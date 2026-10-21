/**
 * Kiểm tra thực tế: chạy đúng pipeline của app (pdf.js -> cấu trúc -> .docx,
 * trang scan/lỗi font -> OCR) trên các PDF thật trong tests/real/, xuất .docx +
 * báo cáo JSON để bước kiểm chứng độc lập (PyMuPDF/python-docx) đối chiếu.
 * Chạy: npx tsx tests/real-test.mts [probe|full]
 */
import { createRequire } from "node:module";
import { promises as fs } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import {
  PdfPageInput,
  PdfTextItem,
  classifyPdf,
  mapPdfJsTextItem,
  pageTextLooksGarbled,
  pdfPagesToBlocks,
} from "../src/utils/pdfStructure";
import { buildDocxBuffer } from "../src/utils/docxExport";
import type { FormattingConfig } from "../src/utils/docx";

/** Cấu hình mặc định đúng như App.tsx khi người dùng chưa chỉnh gì. */
const CONFIG: FormattingConfig = {
  font: "Times New Roman",
  size: 14,
  spacing: "1.5",
  textAlign: "justify",
  textIndent: 1.27,
  paraSpacing: 6,
  leftMargin: 3,
  rightMargin: 2,
};

const require = createRequire(import.meta.url);
const REPO = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const REAL = path.join(REPO, "tests", "real");
const OUT = path.join(REAL, "out");
const RENDERED = path.join(REAL, "rendered");
const OCR_RENDER_SCALE = 2.6; // khớp services/pdf.ts của app
const MIN_CHARS = 24;

const pdfjsLib = await import("pdfjs-dist/legacy/build/pdf.mjs");
const { pathToFileURL } = await import("node:url");
pdfjsLib.GlobalWorkerOptions.workerSrc = pathToFileURL(
  require.resolve("pdfjs-dist/legacy/build/pdf.worker.mjs"),
).href;

async function ocrPng(pngPath: string): Promise<{ items: PdfTextItem[]; confidence: number }> {
  const { createWorker } = await import("tesseract.js");
  // Trong Node, để tesseract.js tự dùng gói @tesseract.js-data/* đã cài —
  // ép workerPath/corePath của bản trình duyệt sẽ làm worker web crash trong Node.
  const worker = await createWorker(["vie", "eng"]);
  try {
    const { data } = (await worker.recognize(pngPath, {}, { blocks: true } as any)) as any;
    const inv = 1 / OCR_RENDER_SCALE;
    const items: PdfTextItem[] = [];
    for (const block of data.blocks || []) {
      for (const para of block.paragraphs || []) {
        for (const line of para.lines || []) {
          for (const w of line.words || []) {
            const text = (w.text || "").trim();
            if (!text) continue;
            items.push({
              text,
              x: w.bbox.x0 * inv,
              y: w.bbox.y0 * inv,
              width: (w.bbox.x1 - w.bbox.x0) * inv,
              height: (w.bbox.y1 - w.bbox.y0) * inv,
              fontSize: (w.bbox.y1 - w.bbox.y0) * inv,
              hasEOL: false,
            });
          }
          const last = items[items.length - 1];
          if (last) last.hasEOL = true;
        }
      }
    }
    return { items, confidence: Math.round(data.confidence || 0) };
  } finally {
    await worker.terminate();
  }
}

async function main() {
  const mode = process.argv[2] || "full";
  await fs.mkdir(OUT, { recursive: true });
  const files = (await fs.readdir(REAL)).filter((f) => f.toLowerCase().endsWith(".pdf")).sort();
  const report: any[] = [];

  for (const file of files) {
    const base = file.replace(/\.pdf$/i, "");
    const data = new Uint8Array(await fs.readFile(path.join(REAL, file)));
    const pdf = await pdfjsLib.getDocument({ data }).promise;
    const pages: PdfPageInput[] = [];
    const pageNotes: any[] = [];
    let ocrConfidences: number[] = [];

    for (let p = 1; p <= pdf.numPages; p++) {
      const page = await pdf.getPage(p);
      const viewport = page.getViewport({ scale: 1 });
      const tc = await page.getTextContent();
      const items: PdfTextItem[] = [];
      for (const raw of tc.items as any[]) {
        if (!("str" in raw)) continue;
        const m = mapPdfJsTextItem(raw, viewport.height);
        if (m) items.push(m);
      }
      const charCount = items.reduce((n, it) => n + it.text.trim().length, 0);
      const garbled = charCount >= MIN_CHARS ? pageTextLooksGarbled(items) : { garbled: false, reason: "" };
      const needOcr = charCount < MIN_CHARS || garbled.garbled;

      if (!needOcr) {
        pages.push({ width: viewport.width, height: viewport.height, items, fromOcr: false });
        pageNotes.push({ page: p, mode: "text", chars: charCount });
      } else if (mode === "probe") {
        pageNotes.push({ page: p, mode: "needs-ocr", chars: charCount, garbledReason: garbled.reason || "ít chữ (scan)" });
        pages.push({ width: viewport.width, height: viewport.height, items: [], fromOcr: false });
      } else {
        const png = path.join(RENDERED, base, `page-${p}.png`);
        const { items: ocrItems, confidence } = await ocrPng(png);
        ocrConfidences.push(confidence);
        pages.push({ width: viewport.width, height: viewport.height, items: ocrItems, fromOcr: true, ocrConfidence: confidence });
        pageNotes.push({ page: p, mode: "ocr", chars: charCount, garbledReason: garbled.reason || "ít chữ (scan)", confidence });
      }
    }

    if (mode === "probe") {
      report.push({ file, pages: pdf.numPages, pageNotes });
      continue;
    }

    const blocks = pdfPagesToBlocks(pages);
    const tableBlocks = blocks.filter((b) => b.type === "table");
    const buf = await buildDocxBuffer(blocks, CONFIG);
    await fs.writeFile(path.join(OUT, `${base}.docx`), buf);
    await fs.writeFile(
      path.join(OUT, `${base}.txt`),
      blocks
        .map((b) => {
          if (b.type === "table") return b.rows.map((r) => r.map((c) => c.map((l) => l.map((x) => x.text).join(" ")).join(" ")).join(" | ")).join("\n");
          if (b.type === "image") return "[ẢNH]";
          return (b.runs || []).map((r) => r.text).join("");
        })
        .join("\n"),
      "utf8",
    );
    report.push({
      file,
      pages: pdf.numPages,
      classification: classifyPdf(pages),
      pageNotes,
      blocks: blocks.length,
      tables: tableBlocks.map((t) => ({ rows: t.rows.length, cols: t.rows[0]?.length || 0 })),
      ocrConfidences,
      docxBytes: buf.length,
    });
    console.log(`✔ ${file}: ${pdf.numPages} trang, ${blocks.length} khối, bảng: ${JSON.stringify(report[report.length - 1].tables.slice(0, 6))}`);
  }

  await fs.writeFile(path.join(OUT, "report.json"), JSON.stringify(report, null, 2), "utf8");
  console.log("REPORT:", path.join(OUT, "report.json"));
}

main().catch((e) => {
  console.error("REAL TEST FAILED", e);
  process.exit(1);
});
