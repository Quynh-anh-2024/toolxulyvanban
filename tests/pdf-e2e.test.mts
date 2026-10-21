/**
 * Test đầu-cuối: PDF thật → pdf.js trích xuất → khối tài liệu → khẳng định nội dung.
 * Phần OCR chạy bằng tesseract.js trên ảnh trang scan (giống đường đi trong trình duyệt,
 * chỉ khác tầng render: ở đây dùng thẳng file PNG của trang scan).
 * Chạy: npx tsx tests/pdf-e2e.test.mts
 */
import assert from "node:assert";
import { readFileSync } from "node:fs";
// @ts-ignore — bản legacy của pdf.js chạy được trong Node
import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";
import {
  PdfPageInput,
  PdfTextItem,
  classifyPdf,
  mapPdfJsTextItem,
  normalizeViText,
  pdfPagesToBlocks,
  blocksToRawText,
} from "../src/utils/pdfStructure";

async function extractTextPage(pdfPath: string): Promise<PdfPageInput> {
  const data = new Uint8Array(readFileSync(pdfPath));
  const pdf = await (pdfjsLib as any).getDocument({ data }).promise;
  const page = await pdf.getPage(1);
  const viewport = page.getViewport({ scale: 1 });
  const tc = await page.getTextContent();
  const resolveFontName = (key: string): string => key;
  const items: PdfTextItem[] = [];
  for (const raw of tc.items as any[]) {
    if (!("str" in raw)) continue;
    const mapped = mapPdfJsTextItem(
      {
        str: raw.str,
        transform: raw.transform,
        width: raw.width,
        height: raw.height,
        fontName: resolveFontName(raw.fontName),
      },
      viewport.height
    );
    if (mapped) items.push(mapped);
  }
  return {
    pageNumber: 1,
    width: viewport.width,
    height: viewport.height,
    items,
    source: items.reduce((n, it) => n + it.text.trim().length, 0) >= 24 ? "text" : "ocr",
  };
}

async function main() {
  let passed = 0;
  const ok = (name: string) => {
    passed++;
    console.log(`  ✓ ${name}`);
  };

  // ---------- 1) PDF có lớp chữ ----------
  const textPage = await extractTextPage("tests/fixtures/cong-van-text.pdf");
  assert.equal(textPage.source, "text");
  const blocks = pdfPagesToBlocks([textPage]);
  const rawNorm = normalizeViText(blocksToRawText(blocks));
  console.log("E2E blocks:", blocks.map((b) => (b.kind === "table" ? `table(${b.role})` : b.kind === "paragraph" ? `p(${b.role})` : "image")).join(", "));

  assert.equal((blocks[0] as any).role, "adminHeader", "khối đầu phải là adminHeader");
  ok("PDF thật: nhận diện phần đầu thể thức");
  assert.ok(rawNorm.includes("SO: 12/KH-THMV"), "thiếu số ký hiệu");
  assert.ok(rawNorm.includes("KE HOACH"), "thiếu tên loại");
  ok("PDF thật: giữ số ký hiệu + tên loại văn bản");
  const subtitle = blocks.find((b) => b.kind === "paragraph" && b.role === "subtitle");
  assert.ok(subtitle, "dòng trích yếu phải thành block subtitle (nhận diện theo vị trí dưới tên loại)");
  ok("PDF thật: trích yếu thành subtitle (nhận diện theo cấu trúc)");
  const dataTable = blocks.find((b) => b.kind === "table" && b.role === "data") as any;
  assert.ok(dataTable && dataTable.rows.length === 3, "bảng dữ liệu phải có 3 hàng");
  ok("PDF thật: bảng dữ liệu 3 hàng theo tọa độ cột");
  const sig = blocks.find((b) => b.kind === "table" && b.role === "signature");
  assert.ok(sig, "thiếu khối chữ ký");
  ok("PDF thật: khối nơi nhận/chữ ký");

  // ---------- 2) PDF scan: pdf.js không có chữ → phải qua OCR ----------
  const scanPage = await extractTextPage("tests/fixtures/thong-bao-scan.pdf");
  const scanChars = scanPage.items.reduce((n, it) => n + it.text.trim().length, 0);
  assert.ok(scanChars < 24, `trang scan phải gần như không có lớp chữ (có ${scanChars})`);
  ok(`PDF scan: lớp chữ gần như trống (${scanChars} ký tự) → app sẽ chuyển sang OCR`);

  console.log("  … đang OCR trang scan bằng Tesseract (vie) — lần đầu cần tải dữ liệu ngôn ngữ");
  const { createWorker } = await import("tesseract.js");
  const worker = await createWorker(["vie", "eng"]);
  const { data } = await worker.recognize("tests/fixtures/scan-page.png", {}, { blocks: true } as any);
  await worker.terminate();

  // Quy đổi tọa độ OCR (ảnh 1240x1754) về point A4 (595x842) — giống services/pdf.ts.
  const scale = 595 / 1240;
  const ocrItems: PdfTextItem[] = [];
  for (const block of (data as any).blocks || []) {
    for (const para of block.paragraphs || []) {
      for (const line of para.lines || []) {
        for (const word of line.words || []) {
          const text = (word.text || "").trim();
          if (!text) continue;
          const h = (word.bbox.y1 - word.bbox.y0) * scale;
          ocrItems.push({
            text: text + " ",
            x: word.bbox.x0 * scale,
            y: word.bbox.y0 * scale,
            width: (word.bbox.x1 - word.bbox.x0) * scale,
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
  assert.ok(ocrItems.length >= 15, `OCR quá ít từ: ${ocrItems.length}`);
  const ocrPage: PdfPageInput = { pageNumber: 1, width: 595, height: 842, items: ocrItems, source: "ocr" };
  assert.equal(classifyPdf([ocrPage]).kind, "scanned");
  const ocrBlocks = pdfPagesToBlocks([ocrPage]);
  const ocrNorm = normalizeViText(blocksToRawText(ocrBlocks));
  console.log("  OCR text:", blocksToRawText(ocrBlocks).replace(/\n+/g, " | ").slice(0, 220));
  assert.ok(ocrNorm.includes("THONG BAO"), "OCR thiếu tiêu đề THÔNG BÁO");
  assert.ok(ocrNorm.includes("QUOC KHANH"), "OCR thiếu 'Quốc khánh'");
  assert.ok(ocrNorm.includes("HOC SINH NGHI LE"), "OCR thiếu nội dung chính");
  ok(`OCR tiếng Việt: nhận đúng nội dung chính (tin cậy TB ${Math.round(data.confidence)}%)`);

  console.log(`\nTest đầu-cuối: ${passed} kiểm tra ĐẠT.`);
}

main().catch((e) => {
  console.error("E2E THẤT BẠI:", e);
  process.exit(1);
});
