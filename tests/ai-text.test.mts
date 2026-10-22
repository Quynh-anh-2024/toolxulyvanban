/**
 * Test aiTextToBlocks — dựng cấu trúc từ văn bản AI đánh lại trang scan.
 * Mẫu mô phỏng đúng đầu ra thật của Gemini trên CV 5095 (đo 22/10/2026).
 */
import { aiTextToBlocks, aiTextToTables } from "../src/utils/aiText";
import type { DocBlock } from "../src/utils/pdfStructure";

let passed = 0, failed = 0;
function check(name: string, cond: boolean, extra?: string) {
  if (cond) { passed++; console.log(`  ✓ ${name}`); }
  else { failed++; console.log(`  ✗ ${name}${extra ? " — " + extra : ""}`); }
}
const textOf = (b: DocBlock): string =>
  b.kind === "paragraph" ? b.runs.map((r) => r.text).join("") :
  b.kind === "table" ? b.rows.flat(3).map((r: any) => r.text).join(" ") : "";

const SAMPLE = `BỘ GIÁO DỤC VÀ ĐÀO TẠO
CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM
Độc lập – Tự do – Hạnh phúc
Số: 5095/BGDĐT-NGCBQLGD
Hà Nội, ngày 04 tháng 08 năm 2026

V/v xin ý kiến góp ý dự thảo Nghị quyết

Kính gửi: Ủy ban nhân dân các tỉnh, thành phố trực thuộc Trung ương

Thực hiện ý kiến chỉ đạo của Phó Thủ tướng, Bộ phối hợp xây dựng dự thảo Nghị quyết về chế độ trợ cấp.

| STT | Họ và tên | Điểm |
|---|---|---|
| 1 | Trần Văn A | 9 |
| 2 | Nguyễn Thị B | 10 |

Nơi nhận:
- Như trên;
- Lưu: VT, NGCBQLGD.

KT. BỘ TRƯỞNG
THỨ TRƯỞNG

Nguyễn Văn C`;

const blocks = aiTextToBlocks(SAMPLE);
const head = blocks[0];
check("khối đầu là bảng adminHeader", head.kind === "table" && head.role === "adminHeader");
if (head.kind === "table") {
  const left = head.rows[0][0].map((line) => line.map((r) => r.text).join("")).join("\n");
  const right = head.rows[0][1].map((line) => line.map((r) => r.text).join("")).join("\n");
  check("ô trái có cơ quan + số", left.includes("BỘ GIÁO DỤC VÀ ĐÀO TẠO") && left.includes("Số: 5095/BGDĐT-NGCBQLGD"), left);
  check("ô phải có quốc hiệu + ngày", right.includes("CỘNG HÒA XÃ HỘI") && right.includes("ngày 04 tháng 08 năm 2026"), right);
}
const subtitle = blocks.find((b) => b.kind === "paragraph" && b.role === "subtitle");
check("trích yếu V/v thành subtitle", !!subtitle && textOf(subtitle!).startsWith("V/v xin ý kiến"));
const recipient = blocks.find((b) => b.kind === "paragraph" && b.role === "recipient");
check("Kính gửi thành recipient", !!recipient && textOf(recipient!).includes("Ủy ban nhân dân"));
const body = blocks.find((b) => b.kind === "paragraph" && b.role === "body" && textOf(b).includes("dự thảo Nghị quyết về chế độ"));
check("đoạn thân nguyên vẹn", !!body);
const table = blocks.find((b) => b.kind === "table" && b.role === "data");
check("bảng markdown thành bảng data 3 hàng", !!table && table.rows.length === 3);
if (table && table.kind === "table") {
  check("ô bảng đúng chữ", textOf(table).includes("Trần Văn A") && textOf(table).includes("Nguyễn Thị B"));
  check("hàng đầu bảng in đậm", table.rows[0][1][0][0].bold === true);
}
const noiNhan = blocks.findIndex((b) => b.kind === "paragraph" && textOf(b).startsWith("Nơi nhận"));
check("có khối Nơi nhận", noiNhan > 0);
const sigIdx = blocks.findIndex((b) => b.kind === "paragraph" && b.role === "signatureText");
check("khối ký sau Nơi nhận", sigIdx > noiNhan && textOf(blocks[sigIdx]).includes("KT. BỘ TRƯỞNG"));
const nameBlock = blocks[blocks.length - 1];
check("tên người ký là signatureText", nameBlock.kind === "paragraph" && nameBlock.role === "signatureText" && textOf(nameBlock) === "Nguyễn Văn C");

const tables = aiTextToTables(SAMPLE);
check("aiTextToTables trích được 1 bảng 3x3", tables.length === 1 && tables[0].length === 3 && tables[0][1][1] === "Trần Văn A");

/* Văn bản thường (không phải văn bản hành chính): không bịa đầu trang. */
const plain = aiTextToBlocks("Đoạn văn thứ nhất có hai câu. Đây là câu hai.\n\nĐoạn thứ hai.");
check("văn bản thường: 2 đoạn body, không adminHeader", plain.length === 2 && plain.every((b) => b.kind === "paragraph" && b.role === "body"));

console.log(`\nAI-TEXT TEST: ${passed} passed, ${failed} failed`);
process.exit(failed ? 1 : 0);
