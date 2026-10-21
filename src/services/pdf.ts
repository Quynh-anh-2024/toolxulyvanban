/**
 * services/pdf.ts — Trích xuất nội dung PDF trên trình duyệt.
 *
 * Xử lý 3 loại PDF "đa dạng":
 *  1. PDF có lớp chữ (xuất từ Word/web): đọc thẳng text item + tọa độ + font bằng pdf.js.
 *  2. PDF scan (ảnh quét/chụp, không có lớp chữ): render trang thành ảnh rồi OCR
 *     tiếng Việt bằng Tesseract (vie + eng) — dữ liệu chữ + khung tọa độ từ OCR.
 *  3. PDF hỗn hợp: quyết định theo TỪNG TRANG (trang nào < ngưỡng ký tự thì OCR trang đó).
 *
 * Ảnh nhúng trong trang có lớp chữ được trích ra để đưa vào Word; trang scan thì không
 * chèn ảnh toàn trang (tránh trùng với chữ OCR) — chữ ký/dấu mộc trên trang scan sẽ
 * không qua được OCR, xem cảnh báo trả về.
 */
import * as pdfjsLib from "pdfjs-dist";
// Worker của pdf.js — Vite đóng gói sẵn, không phụ thuộc CDN.
// @ts-ignore — khai báo ?url của Vite
import workerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";
// Worker của Tesseract — BẮT BUỘC chỉ định file local: để mặc định, thư viện sẽ
// tải worker từ CDN (jsdelivr) và lỗi 404 khiến OCR treo/Thất bại. Core wasm và
// dữ liệu ngôn ngữ vie/eng cũng đã đóng gói sẵn trong thư mục public/.
// @ts-ignore — khai báo ?url của Vite
import tesseractWorkerUrl from "tesseract.js/dist/worker.min.js?url";
import {
  PdfPageInput,
  PdfTextItem,
  classifyPdf,
  mapPdfJsTextItem,
  pageTextLooksGarbled,
  pdfPagesToBlocks,
  blocksToHtml,
  blocksToRawText,
  type DocBlock,
  type PdfClassification,
} from "../utils/pdfStructure";

pdfjsLib.GlobalWorkerOptions.workerSrc = workerUrl;

/** Số ký tự tối thiểu để coi một trang là "có lớp chữ". */
const MIN_CHARS_PER_PAGE = 24;
/** Độ phóng khi render trang scan cho OCR (2.6 ≈ 190 DPI — đủ nét cho tiếng Việt). */
// Tỉ lệ render trang scan để OCR. Đo thực tế trên văn bản scan thật: 2.6 cho
// độ phủ từ ~83%, nâng lên 3.5 kèm tăng tương phản đạt ~85-86% (Tesseract đọc
// dấu tiếng Việt tốt hơn khi glyph lớn). Cao hơn nữa (4.2) chỉ thêm ~1% nhưng
// canvas A4 phình lên ~9MP, nặng máy người dùng — 3.5 là điểm cân bằng.
const OCR_RENDER_SCALE = 3.5;

export interface PdfConvertResult {
  blocks: DocBlock[];
  html: string;
  rawText: string;
  info: PdfClassification;
  warnings: string[];
  /** Độ tin cậy OCR trung bình (%) — chỉ có khi có trang OCR. */
  ocrConfidence?: number;
}

export type PdfProgress = (stage: string, page: number, totalPages: number) => void;

async function ocrPageImage(
  canvas: HTMLCanvasElement,
  renderScale: number,
  onProgress?: PdfProgress,
  pageNumber = 0,
  totalPages = 0
): Promise<{ items: PdfTextItem[]; confidence: number }> {
  // Nạp tesseract.js theo nhu cầu (chỉ khi gặp trang scan) để không nặng lần tải đầu.
  // Toàn bộ tài nguyên OCR (worker, core wasm, traineddata vie/eng) dùng file LOCAL
  // trong public/ — chạy offline được, không phụ thuộc CDN.
  const { createWorker } = await import("tesseract.js");
  const base = (import.meta as any).env?.BASE_URL || "/";
  const worker = await createWorker(["vie", "eng"], undefined, {
    workerPath: tesseractWorkerUrl,
    corePath: `${base}tesseract-core`,
    langPath: `${base}tessdata`,
    cacheMethod: "write",
    logger: (m: any) => {
      if (m?.status === "recognizing text" && typeof m.progress === "number") {
        onProgress?.(`Đang OCR trang scan (${Math.round(m.progress * 100)}%)`, pageNumber, totalPages);
      }
    },
  } as any);
  try {
    const { data } = await worker.recognize(canvas, {}, { blocks: true } as any);
    const items: PdfTextItem[] = [];
    const blocks: any[] = (data as any).blocks || [];
    for (const block of blocks) {
      for (const para of block.paragraphs || []) {
        for (const line of para.lines || []) {
          for (const word of line.words || []) {
            const text = (word.text || "").trim();
            if (!text) continue;
            const { x0, y0, x1 } = word.bbox;
            const h = (word.bbox.y1 - y0) / renderScale;
            items.push({
              text: text + " ",
              x: x0 / renderScale,
              y: y0 / renderScale,
              width: (x1 - x0) / renderScale,
              height: h,
              fontSize: Math.max(8, h * 0.8),
              bold: false,
              italic: false,
              fontName: "ocr",
            });
          }
        }
      }
    }
    // Tesseract trả text toàn trang khi không có blocks (phiên bản cũ) — dự phòng.
    if (items.length === 0 && data.text) {
      let y = 36;
      for (const lineText of data.text.split(/\r?\n/)) {
        if (!lineText.trim()) continue;
        items.push({
          text: lineText,
          x: 56,
          y,
          width: lineText.length * 6,
          height: 14,
          fontSize: 13,
          bold: false,
          italic: false,
          fontName: "ocr",
        });
        y += 20;
      }
    }
    return { items, confidence: data.confidence || 0 };
  } finally {
    await worker.terminate();
  }
}

async function extractPageImages(page: any): Promise<{ urls: string[]; sizes: Array<{ width: number; height: number }> }> {
  const urls: string[] = [];
  const sizes: Array<{ width: number; height: number }> = [];
  try {
    const ops = await page.getOperatorList();
    const fnArray: number[] = ops.fnArray || [];
    const argsArray: any[] = ops.argsArray || [];
    const paintImage = (pdfjsLib as any).OPS?.paintImageXObject;
    if (paintImage === undefined) return { urls, sizes };
    for (let i = 0; i < fnArray.length && urls.length < 6; i++) {
      if (fnArray[i] !== paintImage) continue;
      const imgName = argsArray[i]?.[0];
      if (!imgName) continue;
      const img = await new Promise<any>((resolve) => {
        try {
          page.objs.get(imgName, (value: any) => resolve(value));
        } catch {
          resolve(null);
        }
      });
      if (!img || !img.data || !img.width || !img.height) continue;
      // Bỏ qua ảnh quá nhỏ (icon/đường kẻ) và ảnh gần như toàn trang (nền scan).
      if (img.width < 40 || img.height < 40) continue;
      const canvas = document.createElement("canvas");
      canvas.width = img.width;
      canvas.height = img.height;
      const ctx = canvas.getContext("2d");
      if (!ctx) continue;
      const imageData = new ImageData(new Uint8ClampedArray(img.data), img.width, img.height);
      ctx.putImageData(imageData, 0, 0);
      urls.push(canvas.toDataURL("image/png"));
      sizes.push({ width: img.width, height: img.height });
    }
  } catch {
    // Trích ảnh là tính năng phụ — lỗi thì bỏ qua, không chặn chuyển đổi.
  }
  return { urls, sizes };
}

export async function convertPdfFile(
  file: File,
  onProgress?: PdfProgress
): Promise<PdfConvertResult> {
  const buffer = await file.arrayBuffer();
  // BẮT BUỘC khai báo cMapUrl/standardFontDataUrl/wasmUrl trỏ tới bản LOCAL
  // trong public/: pdf.js trên trình duyệt KHÔNG tự có các tệp này (khác môi
  // trường Node, nơi chúng nằm sẵn cạnh gói cài đặt). Thiếu CMap, các PDF in
  // từ hệ thống trường học/cổng dịch vụ dùng font CID sẽ bị giải mã thành ký
  // tự sai — lớp chữ hoàn hảo bị bộ lọc tưởng là "lỗi font" và app buộc phải
  // OCR oan, cho ra chữ sai dấu. (Đã gặp thật với CV 5095/BGDĐT và bảng tổng
  // hợp kết quả đánh giá in từ hệ thống quản lý trường học.)
  const base = (import.meta as any).env?.BASE_URL || "/";
  const pdf = await pdfjsLib.getDocument({
    data: buffer,
    cMapUrl: `${base}cmaps/`,
    cMapPacked: true,
    standardFontDataUrl: `${base}standard_fonts/`,
    wasmUrl: `${base}wasm/`,
  }).promise;
  const pages: PdfPageInput[] = [];
  const warnings: string[] = [];
  const confidences: number[] = [];
  let hasSignatureOrStampRisk = false;

  for (let p = 1; p <= pdf.numPages; p++) {
    onProgress?.("Đang đọc trang", p, pdf.numPages);
    const page = await pdf.getPage(p);
    const viewport = page.getViewport({ scale: 1 });
    const textContent = await page.getTextContent();
    // LƯU Ý: pdf.js không phơi tên font gốc (vd "TimesNewRoman-Bold") cho font
    // subset — text item chỉ có mã font nội bộ. Vì vậy đậm/nghiêng trong PDF gốc
    // chỉ nhận diện được khi mã font chứa gợi ý; các vai trò quan trọng (tiêu đề,
    // trích yếu, mục, hàng đầu bảng) được suy ra theo CẤU TRÚC ở pdfStructure.ts
    // và được áp đậm khi xuất Word, nên kết quả vẫn đúng thể thức.
    const items: PdfTextItem[] = [];
    for (const raw of textContent.items as any[]) {
      if (!("str" in raw)) continue;
      const mapped = mapPdfJsTextItem(
        {
          str: raw.str,
          transform: raw.transform,
          width: raw.width,
          height: raw.height,
          fontName: raw.fontName,
        },
        viewport.height
      );
      if (mapped) items.push(mapped);
    }
    const charCount = items.reduce((n, it) => n + it.text.trim().length, 0);
    // Trang có lớp chữ nhưng chữ là rác do lỗi font/mã hoá trong PDF gốc:
    // không được tin lớp chữ đó — đối xử như trang scan (OCR lại từ hình trang).
    const garbledCheck = charCount >= MIN_CHARS_PER_PAGE ? pageTextLooksGarbled(items) : { garbled: false, reason: "" };
    if (garbledCheck.garbled) {
      warnings.push(`Trang ${p}: lớp chữ bị lỗi (${garbledCheck.reason}) — đã OCR lại từ hình trang thay vì dùng lớp chữ.`);
    }

    if (charCount >= MIN_CHARS_PER_PAGE && !garbledCheck.garbled) {
      const { urls, sizes } = await extractPageImages(page);
      pages.push({
        pageNumber: p,
        width: viewport.width,
        height: viewport.height,
        items,
        source: "text",
        imageDataUrls: urls.length ? urls : undefined,
        imageSizes: sizes.length ? sizes : undefined,
      });
    } else {
      // Trang scan / không có lớp chữ → render + OCR.
      onProgress?.("Đang OCR trang scan", p, pdf.numPages);
      const ocrViewport = page.getViewport({ scale: OCR_RENDER_SCALE });
      const rawCanvas = document.createElement("canvas");
      rawCanvas.width = Math.floor(ocrViewport.width);
      rawCanvas.height = Math.floor(ocrViewport.height);
      const rawCtx = rawCanvas.getContext("2d");
      if (!rawCtx) throw new Error("Trình duyệt không tạo được canvas để OCR.");
      await page.render({ canvasContext: rawCtx, viewport: ocrViewport, canvas: rawCanvas } as any).promise;
      // Tiền xử lý nhẹ cho OCR: chuyển xám + tăng tương phản/độ sáng — bản scan
      // giấy thường ngả vàng/mờ, xử lý này giúp Tesseract bắt dấu tiếng Việt
      // tốt hơn (đo thực tế +~1% độ phủ từ, không tốn thêm bộ nhớ đáng kể).
      const canvas = document.createElement("canvas");
      canvas.width = rawCanvas.width;
      canvas.height = rawCanvas.height;
      const ctx = canvas.getContext("2d");
      if (!ctx) throw new Error("Trình duyệt không tạo được canvas để OCR.");
      ctx.filter = "grayscale(100%) contrast(112%) brightness(104%)";
      ctx.drawImage(rawCanvas, 0, 0);
      const { items: ocrItems, confidence } = await ocrPageImage(canvas, OCR_RENDER_SCALE, onProgress, p, pdf.numPages);
      confidences.push(confidence);
      if (ocrItems.length > 0) hasSignatureOrStampRisk = true;
      pages.push({
        pageNumber: p,
        width: viewport.width,
        height: viewport.height,
        items: ocrItems,
        source: "ocr",
      });
    }
  }

  const info = classifyPdf(pages);
  if (info.kind === "empty") {
    warnings.push("Không trích được chữ từ PDF này (có thể file hỏng hoặc được bảo vệ). Hãy thử mở PDF và in lại thành PDF mới.");
  }
  if (info.ocrPages > 0) {
    warnings.push(
      `Có ${info.ocrPages}/${info.totalPages} trang là ảnh scan nên phải OCR: chữ có thể sai chính tả ở chỗ mờ/nghiêng — nên soát lại trước khi ban hành.`
    );
    warnings.push("Chữ ký và con dấu trên trang scan là hình ảnh, OCR không chuyển thành chữ được và ảnh toàn trang không được chèn vào Word.");
    if (hasSignatureOrStampRisk && confidences.length && Math.min(...confidences) < 60) {
      warnings.push("Độ tin cậy OCR của ít nhất một trang dưới 60% — trang đó có thể quét mờ, nên quét lại ở 300 DPI.");
    }
  }

  const blocks = pdfPagesToBlocks(pages);
  return {
    blocks,
    html: blocksToHtml(blocks),
    rawText: blocksToRawText(blocks),
    info,
    warnings,
    ocrConfidence: confidences.length
      ? Math.round(confidences.reduce((a, b) => a + b, 0) / confidences.length)
      : undefined,
  };
}
