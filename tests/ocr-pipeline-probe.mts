/**
 * Probe: đo độ phủ từ qua TỪNG TẦNG của pipeline OCR của app
 * (Tesseract thô -> items theo từ -> pdfPagesToBlocks -> text cuối)
 * trên trang 1 CV 5095 (scan thật), đối chiếu truth-cv-p1.txt chép tay từ ảnh.
 * Chạy: npx tsx tests/ocr-pipeline-probe.mts [anh.png] [scale]
 */
import { promises as fs } from "node:fs";
import { pdfPagesToBlocks, blocksToRawText, type PdfTextItem, type PdfPageInput } from "../src/utils/pdfStructure";

const TOKEN_RE = /[0-9A-Za-zÀ-ỹĐđ]+/g;
function tokens(t: string): Map<string, number> {
  const m = new Map<string, number>();
  for (const tok of t.normalize("NFC").match(TOKEN_RE) || []) m.set(tok, (m.get(tok) || 0) + 1);
  return m;
}
function coverage(src: Map<string, number>, out: Map<string, number>): number {
  let hit = 0, total = 0;
  for (const [k, v] of src) { total += v; hit += Math.min(v, out.get(k) || 0); }
  return total ? hit / total : 1;
}
function missingWords(src: Map<string, number>, out: Map<string, number>, limit = 40): string[] {
  const miss: Array<[string, number]> = [];
  for (const [k, v] of src) { const d = v - (out.get(k) || 0); if (d > 0) miss.push([k, d]); }
  return miss.sort((a, b) => b[1] - a[1]).slice(0, limit).map(([k, d]) => `${k}×${d}`);
}

const img = process.argv[2] || "/tmp/cv-p1-z35-pre.png";
const scale = Number(process.argv[3] || 3.5);
const PAGE_W = 596, PAGE_H = 842;

const truth = await fs.readFile("tests/real/truth-cv-p1.txt", "utf8");
const truthTokens = tokens(truth);
console.log(`Truth: ${[...truthTokens.values()].reduce((a, b) => a + b, 0)} từ`);

const { createWorker } = await import("tesseract.js");
const worker = await createWorker(["vie", "eng"]);
const { data } = await worker.recognize(img, {}, { blocks: true } as any);
await worker.terminate();

// --- Tầng A: text thô của Tesseract
const covA = coverage(truthTokens, tokens(data.text || ""));
console.log(`A. Tesseract data.text : phủ ${(covA * 100).toFixed(1)}% (tin cậy ${Math.round(data.confidence || 0)}%)`);
console.log(`   Từ thiếu nhiều nhất: ${missingWords(truthTokens, tokens(data.text || "")).join(", ")}`);

// --- Tầng B: trích từ theo blocks y hệt ocrPageImage trong services/pdf.ts
// (bản mới: lọc conf<30, gắn lineId/paraId, bỏ dòng thuần ký hiệu rác)
const items: PdfTextItem[] = [];
const blocks: any[] = (data as any).blocks || [];
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
        if (conf < 30) continue;
        if (text.length === 1 && conf < 60 && !/[0-9]/.test(text) && rawWords.length >= 4) continue;
        const { x0, y0, x1 } = word.bbox;
        const h = (word.bbox.y1 - y0) / scale;
        lineItems.push({
          text: text + " ",
          x: x0 / scale,
          y: y0 / scale,
          width: (x1 - x0) / scale,
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
      const hasRealWord = lineItems.some((it) => /[0-9A-Za-zÀ-ỹĐđ]{2,}/.test(it.text));
      if (hasRealWord) items.push(...lineItems);
    }
  }
}
const itemsText = items.map((i) => i.text).join(" ");
const covB = coverage(truthTokens, tokens(itemsText));
console.log(`B. Items trích từ blocks (${items.length} từ, ${blocks.length} blocks): phủ ${(covB * 100).toFixed(1)}%`);
console.log(`   Từ thiếu nhiều nhất: ${missingWords(truthTokens, tokens(itemsText)).join(", ")}`);

// --- Tầng C: qua pdfPagesToBlocks của app
const page: PdfPageInput = { pageNumber: 1, width: PAGE_W, height: PAGE_H, items, source: "ocr" };
const docBlocks = pdfPagesToBlocks([page]);
const finalText = blocksToRawText(docBlocks);
const covC = coverage(truthTokens, tokens(finalText));
console.log(`C. Sau pdfPagesToBlocks (${docBlocks.length} blocks): phủ ${(covC * 100).toFixed(1)}%`);
console.log(`   Từ thiếu nhiều nhất: ${missingWords(truthTokens, tokens(finalText)).join(", ")}`);
await fs.writeFile("/tmp/cv-p1-stageC.txt", finalText, "utf8");
console.log("Đã ghi text tầng C vào /tmp/cv-p1-stageC.txt");
