/** Thử nghiệm OCR: cùng 1 trang scan, render ở nhiều tỉ lệ -> đo độ phủ từ
 * so với bản gốc có lớp chữ. npx tsx tests/ocr-experiment.mts */
import { createRequire } from "node:module";
import { promises as fs } from "node:fs";
import path from "node:path";

const require = createRequire(import.meta.url);
const TOKEN_RE = /[0-9A-Za-zÀ-ỹĐđ]+/g;
function tokens(t: string): Map<string, number> {
  const m = new Map<string, number>();
  for (const tok of t.normalize("NFC").match(TOKEN_RE) || []) m.set(tok, (m.get(tok) || 0) + 1);
  return m;
}
function coverage(src: Map<string, number>, out: Map<string, number>): number {
  let hit = 0, total = 0;
  for (const [k, v] of src) {
    total += v;
    hit += Math.min(v, out.get(k) || 0);
  }
  return total ? hit / total : 1;
}

const truth = await fs.readFile("tests/real/out/truth-qd-p1.txt", "utf8");
const truthTokens = tokens(truth);
const { createWorker } = await import("tesseract.js");
const worker = await createWorker(["vie", "eng"]);
for (const variant of process.argv.slice(2)) {
  const png = path.join("tests/real/ocr-variants", variant);
  const { data } = await worker.recognize(png, {}, { text: true } as any);
  const cov = coverage(truthTokens, tokens(data.text || ""));
  console.log(`${variant}: phủ ${(cov * 100).toFixed(1)}% (tin cậy ${Math.round(data.confidence || 0)}%)`);
}
await worker.terminate();
