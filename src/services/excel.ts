/**
 * services/excel.ts — Luồng file bảng tính trên trình duyệt:
 * đọc .xlsx/.xls/.csv thành DocBlock[] và xuất các bảng đã trích ra .xlsx thật.
 */
import {
  readWorkbookFromArrayBuffer,
  workbookToBlocks,
  blocksToTables,
  tablesToWorkbook,
  type SheetInfo,
} from "../utils/spreadsheet";
import { blocksToHtml, blocksToRawText, type DocBlock } from "../utils/pdfStructure";
import * as XLSX from "xlsx";

export interface ExcelConvertResult {
  blocks: DocBlock[];
  html: string;
  rawText: string;
  sheets: SheetInfo[];
  warnings: string[];
}

export async function convertExcelFile(file: File): Promise<ExcelConvertResult> {
  const buffer = await file.arrayBuffer();
  let workbook;
  try {
    workbook = readWorkbookFromArrayBuffer(buffer);
  } catch {
    throw new Error("File bảng tính bị hỏng hoặc định dạng không được hỗ trợ (hãy dùng .xlsx, .xls hoặc .csv).");
  }
  const { blocks, sheets, warnings } = workbookToBlocks(workbook);
  if (blocks.length === 0) {
    throw new Error("Không tìm thấy dữ liệu nào trong file bảng tính (mọi sheet đều trống).");
  }
  return {
    blocks,
    html: blocksToHtml(blocks),
    rawText: blocksToRawText(blocks),
    sheets,
    warnings,
  };
}

/** Xuất mọi bảng dữ liệu trong tài liệu hiện tại thành 1 file .xlsx (mỗi bảng 1 sheet). */
export function exportBlocksTablesToExcel(blocks: DocBlock[], fileName: string): boolean {
  const tables = blocksToTables(blocks);
  if (tables.length === 0) return false;
  const wb = tablesToWorkbook(tables);
  const data = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  const blob = new Blob([data], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `${fileName}_Bang.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(a.href);
  return true;
}
