// Mục 2 — Excel: đọc bảng tính xlsx/xls/csv, tải Word/Excel và trích
// bảng từ PDF ra Excel. Tách khỏi App.tsx (đợt B5); state + hàm xử lý
// vẫn ở App, truyền xuống qua props.
import type { ChangeEvent, ReactNode } from "react";
import { AlertTriangle, Save } from "lucide-react";
import type { SheetInfo } from "../utils/spreadsheet";
import type { DocBlock } from "../utils/pdfStructure";

export interface ExcelSectionProps {
  uploadHint: (label: string) => ReactNode;
  onPick: (handler: (f: File) => void) => (event: ChangeEvent<HTMLInputElement>) => void;
  handleExcelSheetFile: (file: File) => void;
  convertingLine: ReactNode;
  statusLine: ReactNode;
  excelSheets: SheetInfo[] | null;
  excelWarnings: string[];
  excelBlocks: DocBlock[] | null;
  excelTableCount: number;
  handleExportBlocksWord: (blocks: DocBlock[]) => void;
  handleExportTablesExcel: (blocks: DocBlock[]) => void;
  handlePdfTablesFile: (file: File) => void;
  pdfTableBlocks: DocBlock[] | null;
  pdfTableCount: number;
  pdfTableName: string;
}

export function ExcelSection(props: ExcelSectionProps) {
  const {
    uploadHint, onPick, handleExcelSheetFile, convertingLine, statusLine,
    excelSheets, excelWarnings, excelBlocks, excelTableCount,
    handleExportBlocksWord, handleExportTablesExcel, handlePdfTablesFile,
    pdfTableBlocks, pdfTableCount, pdfTableName,
  } = props;
  return (
  <>
    <section className="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
      <h2 className="text-base font-semibold mb-3 text-emerald-700">Bảng tính Excel / CSV</h2>
      <label htmlFor="excelInput" className="cursor-pointer bg-slate-50 border-2 border-dashed border-slate-300 hover:border-emerald-500 hover:bg-emerald-50 transition-colors rounded-lg p-8 flex flex-col items-center">
        {uploadHint("Tải lên file .xlsx, .xls hoặc .csv")}
      </label>
      <input type="file" id="excelInput" accept=".xlsx,.xls,.csv" className="hidden" onChange={onPick((f) => void handleExcelSheetFile(f))} />
      {convertingLine}
      {statusLine}
      {excelSheets && (
        <div className="mt-3 rounded-lg border border-emerald-100 bg-emerald-50 p-3 text-xs text-slate-700 text-left">
          <div className="font-medium text-emerald-800 mb-1">Kết quả đọc bảng tính</div>
          {excelSheets.map((s, i) => (
            <div key={i}>Sheet <b>{s.name}</b>: {s.rowCount} dòng × {s.colCount} cột</div>
          ))}
        </div>
      )}
      {excelWarnings.length > 0 && (
        <div className="mt-2 space-y-1.5 rounded-lg border border-amber-200 bg-amber-50 p-3 text-left">
          {excelWarnings.map((w, i) => (
            <p key={i} className="flex gap-1.5 text-xs text-amber-800"><AlertTriangle className="mt-0.5 h-3.5 w-3.5 shrink-0" /> <span>{w}</span></p>
          ))}
        </div>
      )}
      {excelBlocks && (
        <div className="mt-4 space-y-2">
          <button onClick={() => void handleExportBlocksWord(excelBlocks)} className="w-full bg-blue-800 hover:bg-blue-900 text-white font-medium py-2.5 rounded flex justify-center items-center gap-2 transition-colors">
            <Save className="w-4 h-4" /> Tải File Word (.docx) Chuẩn NĐ 30
          </button>
          <button onClick={() => handleExportTablesExcel(excelBlocks)} className="w-full bg-emerald-700 hover:bg-emerald-800 text-white font-medium py-2.5 rounded flex justify-center items-center gap-2 transition-colors">
            <Save className="w-4 h-4" /> Tải Excel (.xlsx){excelTableCount > 0 ? ` — ${excelTableCount} bảng` : ""}
          </button>
        </div>
      )}
    </section>

    <section className="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
      <h2 className="text-base font-semibold mb-3 text-emerald-700">Trích bảng từ PDF ra Excel</h2>
      <label htmlFor="pdfTableInput" className="cursor-pointer bg-slate-50 border-2 border-dashed border-slate-300 hover:border-emerald-500 hover:bg-emerald-50 transition-colors rounded-lg p-8 flex flex-col items-center">
        {uploadHint("Tải lên file PDF có bảng")}
      </label>
      <input type="file" id="pdfTableInput" accept=".pdf" className="hidden" onChange={onPick((f) => void handlePdfTablesFile(f))} />
      {pdfTableBlocks && (
        <div className="mt-3">
          <p className="text-xs text-slate-600 mb-2">Đã trích <b>{pdfTableCount}</b> bảng từ <b className="break-all">{pdfTableName}</b>.</p>
          <button onClick={() => handleExportTablesExcel(pdfTableBlocks)} disabled={pdfTableCount === 0} className="w-full bg-emerald-700 hover:bg-emerald-800 disabled:opacity-50 text-white font-medium py-2.5 rounded flex justify-center items-center gap-2 transition-colors">
            <Save className="w-4 h-4" /> Tải Excel (.xlsx) — {pdfTableCount} bảng đã trích
          </button>
          {pdfTableCount === 0 && <p className="text-xs text-amber-700 mt-2">PDF này không có bảng nào trích được (bản scan thì bảng chỉ đọc được dạng dòng chữ — xem ở mục PDF → Word).</p>}
        </div>
      )}
    </section>
  </>
  );
}
