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
// Cài vá tương thích trình duyệt cũ TRƯỚC TIÊN (import này phải đứng đầu:
// các import chạy theo thứ tự, trước khi pdf.js được nạp — xem
// utils/browserPolyfills.ts về vụ Safari/iOS cũ lỗi "undefined is not a
// function" khi mở PDF).
import "../utils/installPolyfills";
// Dùng BẢN LEGACY của pdf.js: bản chuẩn v6 đòi trình duyệt rất mới
// (Promise.withResolvers từ Safari 17.4, iterator helpers từ Safari 26...)
// khiến iPhone đời cũ lỗi ngay khi mở PDF. Bản legacy có sẵn polyfill
// core-js cho các API đó. Phần legacy KHÔNG vá được — duyệt ReadableStream
// bằng for await...of (Safari < 18.4) và Promise.withResolvers trong worker —
// do utils/browserPolyfills.ts đảm nhiệm ở cả main thread lẫn worker.
import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";
// Worker của pdf.js — qua điểm vào riêng (src/workers/pdfWorker.ts) để cài
// vá tương thích ngay trong worker trước khi pdf.js chạy.
// @ts-ignore — khai báo ?worker của Vite
import PdfWorker from "../workers/pdfWorker.ts?worker";
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
import { hasGeminiKey, transcribePageImage } from "./ai";
import { aiTextToBlocks } from "../utils/aiText";
import { crossCheckPage, crossCheckWarning } from "../utils/crossCheck";

/** Cách đọc TRANG SCAN (trang không có lớp chữ dùng được) — TỰ ĐỘNG,
 *  không còn nút chọn (theo yêu cầu của thầy Hải 22/10):
 *  - Có khóa Gemini: AI ĐÁNH LẠI từng chữ từ ảnh trang (đo thật trên
 *    CV 5095: phủ 100% từ so với bản chép tay) + OCR chạy nền đối chiếu,
 *    gắn cờ những số liệu hai bên đọc khác nhau (xem utils/crossCheck).
 *  - Chưa có khóa: OCR ngay trên máy (offline; trần ~89% với scan cổng).
 *  Trang có lớp chữ LUÔN đọc thẳng lớp chữ (chuẩn tuyệt đối). */

// Vá tương thích ở main thread đã được cài bởi import đầu tiên ở đầu file
// (utils/installPolyfills) — trước cả khi pdf.js được nạp.
pdfjsLib.GlobalWorkerOptions.workerPort = new PdfWorker();

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
  /** Số trang scan đã được AI đánh lại (chỉ có khi chọn chế độ AI). */
  aiPages?: number;
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
    // GIỮ NGUYÊN cấu trúc dòng/đoạn của Tesseract (lineId/paraId): việc gom
    // dòng lại theo tọa độ y ở tầng cấu trúc đã từng làm vỡ văn bản scan —
    // hộp từ OCR rung vài point là dòng bị xé, đoạn bị máy tách cột tưởng
    // nhầm thành bảng rồi xáo trộn thứ tự đọc, nuốt mất chữ. Tesseract tự
    // phân đoạn trang tốt hơn nhiều; tầng cấu trúc chỉ việc dùng lại.
    let lineSeq = 0;
    let paraSeq = 0;
    for (const block of blocks) {
      for (const para of block.paragraphs || []) {
        paraSeq++;
        for (const line of para.lines || []) {
          lineSeq++;
          const rawWords = (line.words || []).filter((w: any) => (w.text || "").trim());
          const lineItems: PdfTextItem[] = [];
          for (const word of rawWords) {
            const text = (word.text || "").trim();
            const conf = typeof word.confidence === "number" ? word.confidence : 100;
            // Lọc rác theo độ tin cậy: chữ thật trên bản scan sạch hiếm khi
            // dưới 30, còn rác từ con dấu/chữ ký/vết bẩn hầu như luôn dưới 30
            // (đo trên CV 5095: 10/16 từ conf<50 là rác dấu mộc). Ngưỡng 30
            // giữ lại được cả các từ thật bị đọc yếu (vd "trên;" conf 35).
            if (conf < 30) continue;
            // Mảnh vụn 1 ký tự (mảnh dấu, nét chữ ký lạc vào dòng chữ) với độ
            // tin cậy trung bình: bỏ khi dòng còn nhiều từ khác — giữ lại
            // nếu là số (mục lục, số thứ tự) hoặc dòng quá ngắn.
            if (text.length === 1 && conf < 60 && !/[0-9]/.test(text) && rawWords.length >= 4) continue;
            const { x0, y0, x1 } = word.bbox;
            const h = (word.bbox.y1 - y0) / renderScale;
            lineItems.push({
              text: text + " ",
              x: x0 / renderScale,
              y: y0 / renderScale,
              width: (x1 - x0) / renderScale,
              height: h,
              fontSize: Math.max(8, h * 0.8),
              bold: false,
              italic: false,
              fontName: "ocr",
              conf,
              lineId: lineSeq,
              paraId: paraSeq,
            });
          }
          // Dòng không còn từ "giống chữ" nào (≥2 ký tự chữ/số) = thuần ký
          // hiệu rác quanh dấu mộc — bỏ cả dòng, tránh chèn "~=<" vào Word.
          const hasRealWord = lineItems.some((it) => /[0-9A-Za-zÀ-ỹĐđ]{2,}/.test(it.text));
          if (hasRealWord) items.push(...lineItems);
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

/** Render một trang PDF thành canvas ở độ phóng OCR. */
async function renderPageCanvas(page: any): Promise<HTMLCanvasElement> {
  const ocrViewport = page.getViewport({ scale: OCR_RENDER_SCALE });
  const rawCanvas = document.createElement("canvas");
  rawCanvas.width = Math.floor(ocrViewport.width);
  rawCanvas.height = Math.floor(ocrViewport.height);
  const rawCtx = rawCanvas.getContext("2d");
  if (!rawCtx) throw new Error("Trình duyệt không tạo được canvas để đọc trang scan.");
  await page.render({ canvasContext: rawCtx, viewport: ocrViewport, canvas: rawCanvas } as any).promise;
  return rawCanvas;
}

/** OCR một canvas trang đã render (tiền xử lý xám + tương phản như cũ). */
async function ocrCanvasPage(
  rawCanvas: HTMLCanvasElement,
  onProgress: PdfProgress | undefined,
  pageNumber: number,
  totalPages: number
) {
  const canvas = document.createElement("canvas");
  canvas.width = rawCanvas.width;
  canvas.height = rawCanvas.height;
  const ctx = canvas.getContext("2d");
  if (!ctx) throw new Error("Trình duyệt không tạo được canvas để OCR.");
  ctx.filter = "grayscale(100%) contrast(112%) brightness(104%)";
  ctx.drawImage(rawCanvas, 0, 0);
  return ocrPageImage(canvas, OCR_RENDER_SCALE, onProgress, pageNumber, totalPages);
}

/** Chạy một danh sách việc với số luồng song song tối đa cho trước. */
async function runPool<T>(
  items: T[],
  limit: number,
  worker: (item: T) => Promise<void>
): Promise<void> {
  let idx = 0;
  const runners = Array.from(
    { length: Math.max(1, Math.min(limit, items.length)) },
    async () => {
      while (idx < items.length) {
        const current = items[idx++];
        await worker(current);
      }
    }
  );
  await Promise.all(runners);
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
  // Trang scan đọc bằng AI: blocks dựng thẳng từ văn bản AI đánh lại,
  // không qua pages[]/pdfPagesToBlocks (không có tọa độ để mà dựng).
  const aiBlocksByPage = new Map<number, DocBlock[]>();
  let aiPageCount = 0;
  let aiCharCount = 0;
  // B7: việc AI được gom ở vòng đọc trang rồi chạy song song ở pha 2.
  const aiJobs: { pageNumber: number; width: number; height: number; jpeg: string }[] = [];
  let scanPagesWithoutKey = 0;

  /** OCR một trang đã render và ghi vào pages[] (dùng cho cả đường OCR
   *  thẳng lẫn đường dự phòng khi AI lỗi). */
  const ocrRenderedPage = async (
    rawCanvas: HTMLCanvasElement,
    pageNumber: number,
    width: number,
    height: number
  ) => {
    const { items: ocrItems, confidence } = await ocrCanvasPage(
      rawCanvas,
      onProgress,
      pageNumber,
      pdf.numPages
    );
    confidences.push(confidence);
    if (ocrItems.length > 0) hasSignatureOrStampRisk = true;
    pages.push({ pageNumber, width, height, items: ocrItems, source: "ocr" });
  };

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
      warnings.push(`Trang ${p}: lớp chữ bị lỗi (${garbledCheck.reason}) — đã đọc lại từ hình trang thay vì dùng lớp chữ.`);
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
    } else if (hasGeminiKey()) {
      // Trang scan + đã có khóa Gemini: render ảnh trang (bản màu gốc —
      // AI nhìn như người) và GOM VIỆC; việc chép chữ + đối chiếu OCR
      // chạy ở pha 2 sau vòng đọc.
      onProgress?.("Đang render trang scan", p, pdf.numPages);
      const rawCanvas = await renderPageCanvas(page);
      aiJobs.push({
        pageNumber: p,
        width: viewport.width,
        height: viewport.height,
        jpeg: rawCanvas.toDataURL("image/jpeg", 0.85),
      });
    } else {
      // Trang scan nhưng chưa có khóa Gemini: quét OCR ngay.
      scanPagesWithoutKey++;
      onProgress?.("Đang OCR trang scan", p, pdf.numPages);
      const rawCanvas = await renderPageCanvas(page);
      await ocrRenderedPage(rawCanvas, p, viewport.width, viewport.height);
    }
  }

  if (scanPagesWithoutKey > 0) {
    warnings.push(
      `Chưa có khóa Gemini nên ${scanPagesWithoutKey} trang scan đã quét OCR. Dán khóa Gemini vào mục PDF → Word để app dùng AI đánh lại kèm đối chiếu kiểm tra số liệu.`
    );
  }

  // Pha 2: AI đánh lại SONG SONG (tối đa 3 trang cùng lúc) cho các trang
  // scan đã gom. Đồng thời OCR chạy NỀN (một hàng đợi tuần tự dùng chung)
  // cho chính các trang đó để ĐỐI CHIẾU: trang nào AI lỗi thì kết quả
  // OCR nền thành đường dự phòng luôn — một trang chỉ OCR đúng một lần.
  if (aiJobs.length > 0) {
    let aiDoneCount = 0;
    type OcrResult = Awaited<ReturnType<typeof ocrCanvasPage>>;
    const ocrTaskByPage = new Map<number, Promise<OcrResult>>();
    let ocrQueue: Promise<unknown> = Promise.resolve();
    const ensureOcr = (job: { pageNumber: number }): Promise<OcrResult> => {
      const existing = ocrTaskByPage.get(job.pageNumber);
      if (existing) return existing;
      const task = ocrQueue.then(async () => {
        const page = await pdf.getPage(job.pageNumber);
        const rawCanvas = await renderPageCanvas(page);
        return ocrCanvasPage(rawCanvas, undefined, job.pageNumber, pdf.numPages);
      });
      ocrQueue = task.catch(() => undefined);
      ocrTaskByPage.set(job.pageNumber, task);
      return task;
    };
    // Khởi động OCR đối chiếu ngay — chạy trong lúc chờ AI trả lời.
    for (const job of aiJobs) void ensureOcr(job).catch(() => undefined);

    const aiTextByPage = new Map<number, string>();
    onProgress?.("AI đang đánh lại trang scan", 0, aiJobs.length);
    await runPool(aiJobs, 3, async (job) => {
      let ok = false;
      try {
        const aiText = await transcribePageImage(job.jpeg);
        const aiBlocks = aiTextToBlocks(aiText);
        if (aiBlocks.length > 0) {
          aiBlocksByPage.set(job.pageNumber, aiBlocks);
          aiTextByPage.set(job.pageNumber, aiText);
          aiPageCount++;
          aiCharCount += aiText.length;
          ok = true;
        } else {
          warnings.push(`Trang ${job.pageNumber}: AI không đọc được chữ nào — đã quét OCR thay thế.`);
        }
      } catch (e) {
        const msg = e instanceof Error ? e.message : "lỗi không rõ";
        warnings.push(`Trang ${job.pageNumber}: AI đánh lại không thành công (${msg}) — đã quét OCR thay thế cho trang này.`);
      }
      if (!ok) {
        const res = await ensureOcr(job);
        confidences.push(res.confidence);
        if (res.items.length > 0) hasSignatureOrStampRisk = true;
        pages.push({
          pageNumber: job.pageNumber,
          width: job.width,
          height: job.height,
          items: res.items,
          source: "ocr",
        });
      }
      aiDoneCount++;
      onProgress?.(`AI đã đánh lại ${aiDoneCount}/${aiJobs.length} trang scan`, aiDoneCount, aiJobs.length);
    });

    // Đối chiếu AI ↔ OCR: chờ hàng đợi OCR chạy xong rồi so từng trang;
    // chỉ gắn cờ khi số liệu (token chứa chữ số) hai bên đọc khác nhau.
    onProgress?.("Đang đối chiếu AI với OCR", aiJobs.length, aiJobs.length);
    const ocrResults = await Promise.allSettled([...ocrTaskByPage.values()]);
    const ocrByPage = new Map<number, OcrResult>();
    [...ocrTaskByPage.keys()].forEach((pageNumber, i) => {
      const r = ocrResults[i];
      if (r.status === "fulfilled") ocrByPage.set(pageNumber, r.value);
    });
    for (const job of aiJobs) {
      const aiText = aiTextByPage.get(job.pageNumber);
      const ocrRes = ocrByPage.get(job.pageNumber);
      if (!aiText || !ocrRes) continue;
      const ocrText = ocrRes.items.map((it) => it.text).join(" ");
      const w = crossCheckWarning(job.pageNumber, crossCheckPage(aiText, ocrText));
      if (w) warnings.push(w);
    }
  }

  // Trang OCR dự phòng thêm ở pha 2 có thể lệch thứ tự trong pages[] —
  // sắp lại theo số trang để phân loại + dựng blocks luôn đúng thứ tự.
  pages.sort((a, b) => a.pageNumber - b.pageNumber);

  const info = classifyPdf(pages);
  if (aiPageCount > 0) {
    // classifyPdf chỉ nhìn pages[] (lớp chữ + OCR) — cộng phần AI vào cho đúng.
    info.totalPages = pdf.numPages;
    info.totalChars += aiCharCount;
    info.kind = info.textPages > 0 || info.ocrPages > 0 ? "mixed" : "scanned";
    info.label =
      `PDF ${info.totalPages} trang: ${info.textPages} trang lớp chữ` +
      ` · ${aiPageCount} trang AI đánh lại` +
      (info.ocrPages > 0 ? ` · ${info.ocrPages} trang OCR` : "");
    warnings.push(
      `Đã dùng AI đánh lại ${aiPageCount} trang scan — chữ chép rất sát bản gốc (đo thật: đúng ~100% từ trên văn bản hành chính). Chữ ký và con dấu là hình ảnh nên AI không chuyển thành chữ được.`
    );
  }
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

  let blocks: DocBlock[];
  if (aiBlocksByPage.size === 0) {
    blocks = pdfPagesToBlocks(pages);
  } else {
    // Ghép theo đúng thứ tự trang: trang AI lấy blocks AI, trang còn lại
    // dựng từ items như thường (mỗi trang độc lập khi có trang AI chen vào).
    blocks = [];
    const itemByPage = new Map(pages.map((pg) => [pg.pageNumber, pg]));
    for (let p = 1; p <= pdf.numPages; p++) {
      const aiBlocks = aiBlocksByPage.get(p);
      if (aiBlocks) blocks.push(...aiBlocks);
      else {
        const itemPage = itemByPage.get(p);
        if (itemPage) blocks.push(...pdfPagesToBlocks([itemPage]));
      }
    }
  }
  return {
    blocks,
    html: blocksToHtml(blocks),
    rawText: blocksToRawText(blocks),
    info,
    warnings,
    ocrConfidence: confidences.length
      ? Math.round(confidences.reduce((a, b) => a + b, 0) / confidences.length)
      : undefined,
    aiPages: aiPageCount || undefined,
  };
}
