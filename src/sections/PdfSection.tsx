// Mục 1 — PDF → Word: ô tải PDF, ô khóa Gemini và panel kết quả. Cách
// đọc trang scan TỰ ĐỘNG (không còn nút chọn): có khóa thì AI đánh lại
// kèm đối chiếu OCR, chưa có khóa thì quét OCR. Tách khỏi App.tsx (đợt
// B5); state + hàm xử lý vẫn ở App và được truyền xuống qua props.
import type { ChangeEvent, ReactNode } from "react";
import { AlertTriangle, Save } from "lucide-react";
import { cn } from "../lib/utils";
import { setGeminiKey } from "../services/ai";
import type { DocBlock, PdfClassification } from "../utils/pdfStructure";

export interface PdfSectionProps {
  uploadHint: (label: string) => ReactNode;
  onPick: (handler: (f: File) => void) => (event: ChangeEvent<HTMLInputElement>) => void;
  handlePdfFile: (file: File) => void;
  geminiKeySaved: boolean;
  setGeminiKeySaved: (saved: boolean) => void;
  geminiKeyInput: string;
  setGeminiKeyInput: (value: string) => void;
  aiChecking: boolean;
  handleCheckAi: () => void;
  showNotice: (kind: "error" | "warn" | "ok", text: string) => void;
  convertingLine: ReactNode;
  statusLine: ReactNode;
  pdfInfo: PdfClassification | null;
  pdfAiPages: number | undefined;
  pdfOcrConfidence: number | undefined;
  pdfWarnings: string[];
  pdfBlocks: DocBlock[] | null;
  handleExportBlocksWord: (blocks: DocBlock[]) => void;
}

export function PdfSection(props: PdfSectionProps) {
  const {
    uploadHint, onPick, handlePdfFile,
    geminiKeySaved, setGeminiKeySaved, geminiKeyInput, setGeminiKeyInput,
    aiChecking, handleCheckAi, showNotice, convertingLine, statusLine,
    pdfInfo, pdfAiPages, pdfOcrConfidence, pdfWarnings, pdfBlocks,
    handleExportBlocksWord,
  } = props;
  return (
  <>
    <section className="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
      <h2 className="text-base font-semibold mb-3 text-blue-700">Chuyển PDF sang Word</h2>
      <label htmlFor="fileInput" className="cursor-pointer bg-slate-50 border-2 border-dashed border-slate-300 hover:border-blue-500 hover:bg-blue-50 transition-colors rounded-lg p-8 flex flex-col items-center">
        {uploadHint("Tải lên file PDF")}
      </label>
      <input type="file" id="fileInput" accept=".pdf" className="hidden" onChange={onPick((f) => void handlePdfFile(f))} />
      {geminiKeySaved ? (
        <div className="mt-3 text-left">
          <span className="text-xs text-slate-500">
            Đã lưu khóa Gemini ·{" "}
            <button type="button" disabled={aiChecking} className="font-medium text-blue-700 underline disabled:opacity-50" onClick={() => void handleCheckAi()}>{aiChecking ? "Đang kiểm tra…" : "Kiểm tra"}</button>
            {" "}·{" "}
            <button type="button" className="font-medium text-blue-700 underline" onClick={() => { setGeminiKey(""); setGeminiKeySaved(false); }}>Xóa</button>
          </span>
        </div>
      ) : (
        <div className="mt-3 rounded-lg border border-blue-100 bg-blue-50 p-3 text-left">
          <label className="block text-xs font-medium text-slate-700" htmlFor="geminiKeyInput">Khóa Gemini — <a href="https://aistudio.google.com/apikey" target="_blank" rel="noreferrer" className="font-semibold text-blue-700 underline">bấm vào đây để lấy khóa miễn phí</a>:</label>
          <div className="mt-1.5 flex gap-2">
            <input id="geminiKeyInput" type="password" value={geminiKeyInput} onChange={(e) => setGeminiKeyInput(e.target.value)} placeholder="Dán khóa Gemini của bạn vào đây" className="min-w-0 flex-1 rounded border border-slate-300 px-2 py-1.5 text-sm" autoComplete="off" />
            <button type="button" className="shrink-0 rounded bg-blue-700 px-3 py-1.5 text-sm font-medium text-white hover:bg-blue-800" onClick={() => { if (geminiKeyInput.trim()) { setGeminiKey(geminiKeyInput); setGeminiKeySaved(true); setGeminiKeyInput(""); showNotice("ok", "Đã lưu khóa Gemini."); } }}>Lưu khóa</button>
          </div>
        </div>
      )}
      {convertingLine}
      {statusLine}
      {pdfInfo && (
        <div className="mt-3 rounded-lg border border-blue-100 bg-blue-50 p-3 text-xs text-slate-700 text-left">
          <div className="font-medium text-blue-800 mb-1">Kết quả chuyển PDF → Word</div>
          <div>Loại PDF: <b>{pdfInfo.label}</b></div>
          <div>Số trang: <b>{pdfInfo.totalPages}</b> · Trang có lớp chữ: <b>{pdfInfo.textPages}</b>{pdfAiPages !== undefined && <> · Trang AI đánh lại: <b>{pdfAiPages}</b></>} · Trang OCR: <b>{pdfInfo.ocrPages}</b></div>
          {pdfOcrConfidence !== undefined && <div>Độ tin cậy OCR trung bình: <b>{pdfOcrConfidence}%</b></div>}
        </div>
      )}
      {pdfWarnings.length > 0 && (
        <div className="mt-2 space-y-1.5 rounded-lg border border-amber-200 bg-amber-50 p-3 text-left">
          {pdfWarnings.map((w, i) => {
            const severe = /không thành công|không đọc được|lỗi/i.test(w);
            return (
              <p key={i} className={cn("flex gap-1.5 text-xs", severe ? "font-medium text-red-700" : "text-amber-800")}>
                <AlertTriangle className="mt-0.5 h-3.5 w-3.5 shrink-0" /> <span>{w}</span>
              </p>
            );
          })}
        </div>
      )}
    </section>

    <section className={cn("bg-white p-5 rounded-xl shadow-sm border border-slate-200", !pdfBlocks && "opacity-50 pointer-events-none")}>
      <h2 className="text-base font-semibold mb-3 text-blue-700">Tải xuống</h2>
      <button onClick={() => pdfBlocks && void handleExportBlocksWord(pdfBlocks)} className="w-full bg-blue-800 hover:bg-blue-900 text-white font-medium py-2.5 rounded flex justify-center items-center gap-2 transition-colors">
        <Save className="w-4 h-4" /> Tải File Word (.docx) Chuẩn NĐ 30
      </button>

    </section>
  </>
  );
}
