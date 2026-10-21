/**
 * Test luồng Excel: file .xlsx → DocBlock[] → Word .docx, và DocBlock[] → .xlsx thật
 * (đọc ngược lại để đối chiếu giá trị). Chạy: npx tsx tests/excel.test.mts
 */
import assert from "node:assert";
import { readFileSync } from "node:fs";
import * as XLSX from "xlsx";
import {
  readWorkbookFromArrayBuffer,
  workbookToBlocks,
  blocksToTables,
  blocksToXlsxBuffer,
} from "../src/utils/spreadsheet";
import { buildDocxBuffer } from "../src/utils/docxExport";
import type { FormattingConfig } from "../src/utils/docx";

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

async function main() {
  let passed = 0;
  const ok = (name: string) => {
    passed++;
    console.log(`  ✓ ${name}`);
  };

  const buf = readFileSync("tests/fixtures/bang-tinh-mau.xlsx");
  const wb = readWorkbookFromArrayBuffer(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer);
  const { blocks, sheets } = workbookToBlocks(wb);

  assert.deepEqual(sheets.map((s) => s.name), ["Danh sach hoc sinh", "Lich truc"]);
  assert.equal(sheets[0].rowCount, 4);
  assert.equal(sheets[0].colCount, 4);
  ok("Đọc .xlsx: 2 sheet, đúng số dòng/cột");

  const tables = blocks.filter((b) => b.kind === "table") as any[];
  assert.equal(tables.length, 2);
  const cell = (t: any, r: number, c: number) => t.rows[r][c].map((l: any) => l.map((x: any) => x.text).join("")).join("");
  assert.equal(cell(tables[0], 0, 1), "Họ và tên");
  assert.equal(cell(tables[0], 1, 1), "Nguyễn Văn An");
  assert.equal(cell(tables[0], 3, 3), "7.25");
  assert.equal(cell(tables[1], 1, 2), "Trần Ngọc Hải");
  ok("Nội dung ô giữ đúng (kể cả số thập phân 7.25, tên tiếng Việt)");

  const headings = blocks.filter((b) => b.kind === "paragraph" && b.role === "heading") as any[];
  assert.equal(headings.length, 2);
  ok("Mỗi sheet có tiêu đề tên sheet trước bảng");

  const docxBuf = await buildDocxBuffer(blocks, CONFIG);
  assert.ok(docxBuf.length > 1024 && docxBuf.subarray(0, 2).toString() === "PK");
  ok(`Xuất Word từ Excel: .docx hợp lệ (${docxBuf.length} bytes)`);

  // Chiều ngược: blocks (như trích từ PDF) → .xlsx thật → đọc lại đối chiếu.
  const extracted = blocksToTables(blocks);
  assert.equal(extracted.length, 2);
  assert.equal(extracted[0].rows[2][1], "Trần Thị Bình");
  const xlsxBuf = blocksToXlsxBuffer(blocks);
  assert.ok(xlsxBuf && xlsxBuf.length > 500);
  const wb2 = XLSX.read(xlsxBuf, { type: "buffer" });
  assert.deepEqual(wb2.SheetNames, ["Bang 1", "Bang 2"]);
  const aoa = XLSX.utils.sheet_to_json<any[]>(wb2.Sheets["Bang 1"], { header: 1, raw: false, defval: "" });
  assert.equal(aoa[0][0], "STT");
  assert.equal(aoa[1][1], "Nguyễn Văn An");
  assert.equal(aoa[3][3], "7.25");
  ok("Xuất .xlsx từ bảng trích xuất: đọc ngược khớp từng ô, 2 sheet Bang 1/Bang 2");

  console.log(`\nTest Excel: ${passed} kiểm tra ĐẠT.`);
}

main().catch((e) => {
  console.error("TEST EXCEL THẤT BẠI:", e);
  process.exit(1);
});
