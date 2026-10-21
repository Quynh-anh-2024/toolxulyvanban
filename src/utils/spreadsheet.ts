/**
 * spreadsheet.ts — Cầu nối Excel ↔ mô hình khối tài liệu (DocBlock[]).
 *
 * Hai chiều:
 *  1. File bảng tính (.xlsx/.xls/.csv) → DocBlock[]: mỗi sheet thành một bảng
 *     Word (kèm tiêu đề tên sheet), để xem trước và xuất .docx chuẩn như PDF.
 *  2. DocBlock[] → file .xlsx THẬT: trích mọi bảng dữ liệu (role "data") từ
 *     PDF/Word/Excel đã nạp, mỗi bảng thành một sheet — mở chuẩn trong Excel.
 *
 * Dùng SheetJS (xlsx) — chạy cả trình duyệt lẫn Node để kiểm thử.
 */
import * as XLSX from "xlsx";
import type { DocBlock, DocRun } from "./pdfStructure";

export interface SheetInfo {
  name: string;
  rowCount: number;
  colCount: number;
}

export interface WorkbookConvertResult {
  blocks: DocBlock[];
  sheets: SheetInfo[];
  warnings: string[];
}

const run = (text: string): DocRun[] => (text ? [{ text }] : []);

/** Giới hạn bảo vệ: bảng quá lớn vào Word sẽ rất nặng — cắt ở 500 dòng/sheet khi dựng khối. */
export const MAX_ROWS_PER_SHEET = 500;

export function workbookToBlocks(workbook: XLSX.WorkBook): WorkbookConvertResult {
  const blocks: DocBlock[] = [];
  const sheets: SheetInfo[] = [];
  const warnings: string[] = [];

  workbook.SheetNames.forEach((sheetName, sheetIdx) => {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return;
    // raw:false → lấy chuỗi đã định dạng (ngày tháng, số thập phân) như mắt thấy trong Excel.
    const aoa = XLSX.utils.sheet_to_json<any[]>(sheet, {
      header: 1,
      raw: false,
      defval: "",
      blankrows: false,
    });
    const rows = aoa
      .map((row) => (row || []).map((cell) => String(cell ?? "").trim()))
      .filter((row) => row.some((cell) => cell.length > 0));
    const colCount = rows.reduce((m, r) => Math.max(m, r.length), 0);
    sheets.push({ name: sheetName, rowCount: rows.length, colCount });

    if (rows.length === 0) {
      warnings.push(`Sheet "${sheetName}" trống, đã bỏ qua.`);
      return;
    }
    if (rows.length > MAX_ROWS_PER_SHEET) {
      warnings.push(
        `Sheet "${sheetName}" có ${rows.length} dòng — bản Word chỉ lấy ${MAX_ROWS_PER_SHEET} dòng đầu để tránh file quá nặng. File .xlsx gốc không bị ảnh hưởng.`
      );
    }
    const used = rows.slice(0, MAX_ROWS_PER_SHEET);

    if (workbook.SheetNames.length > 1 || sheetIdx > 0) {
      blocks.push({
        kind: "paragraph",
        runs: [{ text: sheetName, bold: true }],
        align: "left",
        role: "heading",
      });
    }
    blocks.push({
      kind: "table",
      role: "data",
      rows: used.map((row) =>
        Array.from({ length: colCount }, (_, c) => [run(row[c] || "")])
      ),
    });

    // Ô gộp (merge) trong Excel: SheetJS trả giá trị ở ô góc trên-trái, các ô còn
    // lại trống — bảng Word vẫn giữ đủ cột nhưng nội dung ô gộp chỉ nằm ở ô đầu.
    const merges = (sheet as any)["!merges"] as Array<any> | undefined;
    if (merges && merges.length > 0) {
      warnings.push(
        `Sheet "${sheetName}" có ${merges.length} vùng ô gộp — khi sang Word, chữ chỉ nằm ở ô đầu của vùng gộp.`
      );
    }
  });

  return { blocks, sheets, warnings };
}

export function readWorkbookFromArrayBuffer(buffer: ArrayBuffer): XLSX.WorkBook {
  return XLSX.read(new Uint8Array(buffer), { type: "array", cellDates: true });
}

export interface ExtractedTable {
  name: string;
  rows: string[][];
}

/** Trích các bảng dữ liệu từ DocBlock[] để xuất Excel. */
export function blocksToTables(blocks: DocBlock[]): ExtractedTable[] {
  const tables: ExtractedTable[] = [];
  let dataIdx = 0;
  for (const block of blocks) {
    if (block.kind !== "table") continue;
    // Chỉ xuất bảng dữ liệu + bảng từ Excel; bỏ phần đầu thể thức và khối chữ ký.
    if (block.role !== "data") continue;
    dataIdx++;
    const rows = block.rows.map((row) =>
      row.map((cell) => cell.map((line) => line.map((r) => r.text).join("")).join("\n").trim())
    );
    tables.push({ name: `Bang ${dataIdx}`, rows });
  }
  return tables;
}

export function tablesToWorkbook(tables: ExtractedTable[]): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();
  for (const table of tables) {
    const ws = XLSX.utils.aoa_to_sheet(table.rows);
    // Ước lượng bề rộng cột theo độ dài nội dung (tối thiểu 10, tối đa 60 ký tự).
    const colCount = table.rows.reduce((m, r) => Math.max(m, r.length), 0);
    ws["!cols"] = Array.from({ length: colCount }, (_, c) => {
      const maxLen = table.rows.reduce((m, r) => Math.max(m, (r[c] || "").length), 0);
      return { wch: Math.min(60, Math.max(10, maxLen + 2)) };
    });
    // Tên sheet tối đa 31 ký tự và không chứa ký tự cấm của Excel.
    const safeName = table.name.replace(/[\\/?*[\]:]/g, " ").slice(0, 31) || "Bang";
    XLSX.utils.book_append_sheet(wb, ws, safeName);
  }
  return wb;
}

/** Buffer .xlsx thật (Node kiểm thử + dùng chung logic với bản trình duyệt). */
export function workbookToXlsxBuffer(workbook: XLSX.WorkBook): Buffer {
  return XLSX.write(workbook, { type: "buffer", bookType: "xlsx" }) as Buffer;
}

export function blocksToXlsxBuffer(blocks: DocBlock[]): Buffer | null {
  const tables = blocksToTables(blocks);
  if (tables.length === 0) return null;
  return workbookToXlsxBuffer(tablesToWorkbook(tables));
}
