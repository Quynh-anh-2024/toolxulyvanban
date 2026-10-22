// Mục 3 — Word & AI: chuẩn hoá file .docx có sẵn, chọn chế độ giữ khung,
// kiểm tra nhanh tài liệu, soát lỗi AI và tải xuống. Tách khỏi App.tsx
// (đợt B5); state + hàm xử lý vẫn ở App, truyền xuống qua props.
import type { ChangeEvent, ReactNode, RefObject } from "react";
import { AlertTriangle, Loader2, Save, SlidersHorizontal, Wand2 } from "lucide-react";
import { cn } from "../lib/utils";
import {
  exportToWord,
  type DocumentProcessingMode,
  type DocumentStats,
  type FormattingConfig,
} from "../utils/docx";

export interface WordSectionProps {
  uploadHint: (label: string) => ReactNode;
  onPick: (handler: (f: File) => void) => (event: ChangeEvent<HTMLInputElement>) => void;
  handleWordFile: (file: File) => void;
  statusLine: ReactNode;
  hasFile: boolean;
  processingModes: Array<{ value: DocumentProcessingMode; title: string }>;
  processingMode: DocumentProcessingMode;
  setProcessingMode: (mode: DocumentProcessingMode) => void;
  showPreviewFrame: boolean;
  setShowPreviewFrame: (show: boolean) => void;
  documentStats: DocumentStats;
  aiStyle: string;
  setAiStyle: (value: string) => void;
  aiMode: string;
  setAiMode: (value: string) => void;
  handleAITask: (taskType: "summarize" | "quiz" | "proofread") => void;
  isAILoading: boolean;
  activeTask: string | null;
  previewRef: RefObject<HTMLDivElement | null>;
  fileName: string;
  config: FormattingConfig;
}

export function WordSection(props: WordSectionProps) {
  const {
    uploadHint, onPick, handleWordFile, statusLine, hasFile,
    processingModes, processingMode, setProcessingMode, showPreviewFrame,
    setShowPreviewFrame, documentStats, aiStyle, setAiStyle, aiMode,
    setAiMode, handleAITask, isAILoading, activeTask, previewRef, fileName,
    config,
  } = props;
  return (
  <>
    <section className="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
      <h2 className="text-base font-semibold mb-3 text-blue-700">Chuẩn hoá file Word có sẵn</h2>
      <label htmlFor="wordInput" className="cursor-pointer bg-slate-50 border-2 border-dashed border-slate-300 hover:border-blue-500 hover:bg-blue-50 transition-colors rounded-lg p-8 flex flex-col items-center">
        {uploadHint("Tải lên file .docx hoặc .doc")}
      </label>
      <input type="file" id="wordInput" accept=".docx,.doc" className="hidden" onChange={onPick(handleWordFile)} />
      {statusLine}
    </section>

    <section className={cn("bg-white p-5 rounded-xl shadow-sm border border-slate-200", !hasFile && "opacity-50 pointer-events-none")}>
      <h2 className="text-base font-semibold mb-3 text-blue-700">Chế Độ Giữ Khung/Bảng</h2>
      <div className="space-y-2">
        {processingModes.map((mode) => (
          <label key={mode.value} className={cn("block rounded-lg border p-3 cursor-pointer transition", processingMode === mode.value ? "border-blue-500 bg-blue-50" : "border-slate-200 bg-white hover:bg-slate-50")}>
            <div className="flex gap-2 items-start">
              <input type="radio" className="mt-1" checked={processingMode === mode.value} onChange={() => setProcessingMode(mode.value)} />
              <div>
                <div className="text-sm font-semibold text-slate-800">{mode.title}</div>
              </div>
            </div>
          </label>
        ))}
      </div>
      <label className="mt-3 flex items-start gap-2 rounded-lg border border-amber-200 bg-amber-50 p-3 cursor-pointer">
        <input type="checkbox" checked={showPreviewFrame} onChange={(event) => setShowPreviewFrame(event.target.checked)} className="mt-1" />
        <span>
          <span className="block text-sm font-semibold text-slate-800">Hiển thị khung hỗ trợ ở bản xem trước</span>
          <span className="block text-xs text-amber-700 leading-5">Khung phần đầu hành chính chỉ để kiểm tra; khi tải xuống Word sẽ tự ẩn nếu là bảng thể thức.</span>
        </span>
      </label>
      <div className="mt-3 rounded-lg border border-slate-200 bg-slate-50 p-3 text-xs text-slate-700">
        <div className="font-medium flex items-center gap-2 mb-2"><SlidersHorizontal className="w-4 h-4" /> Kiểm tra nhanh tài liệu</div>
        <div className="grid grid-cols-2 gap-y-1">
          <span>Số bảng/khung: <b>{documentStats.tableCount}</b></span>
          <span>Số hình ảnh: <b>{documentStats.imageCount}</b></span>
          <span>Thể thức hành chính: <b>{documentStats.administrativeHeaderDetected ? "Có" : "Chưa rõ"}</b></span>
          <span>Khối chữ ký/nơi nhận: <b>{documentStats.signatureTableDetected ? "Có" : "Chưa rõ"}</b></span>
        </div>
        {documentStats.warningMessages.length > 0 && (
          <div className="mt-2 space-y-1 text-amber-700">
            {documentStats.warningMessages.slice(0, 3).map((message, index) => (
              <div key={index} className="flex gap-1"><AlertTriangle className="w-3 h-3 mt-0.5 shrink-0" /> <span>{message}</span></div>
            ))}
          </div>
        )}
      </div>
    </section>

    <section className={cn("bg-white p-5 rounded-xl shadow-sm border border-slate-200", !hasFile && "opacity-50 pointer-events-none")}>
      <h2 className="text-base font-semibold mb-3 text-blue-700">Soát Lỗi AI</h2>
      <div className="space-y-3">
        <select className="w-full border p-2 rounded text-sm text-slate-700" value={aiStyle} onChange={(event) => setAiStyle(event.target.value)}>
          <option value="grammar">Sửa Ngữ pháp & Nối dòng</option>
          <option value="admin">Chuẩn hóa Hành chính (NĐ 30)</option>
          <option value="edu">Chuẩn hóa Giáo dục (Sư phạm)</option>
        </select>
        <div className="grid grid-cols-2 gap-2 text-sm">
          <label className="flex items-center gap-2 border rounded p-2 cursor-pointer"><input type="radio" checked={aiMode === "suggest"} onChange={() => setAiMode("suggest")} className="cursor-pointer" /> Chỉ báo lỗi</label>
          <label className="flex items-center gap-2 border rounded p-2 cursor-pointer"><input type="radio" checked={aiMode === "autofix"} onChange={() => setAiMode("autofix")} className="cursor-pointer" /> AI tự sửa</label>
        </div>
        <button onClick={() => handleAITask("proofread")} disabled={isAILoading} className="w-full bg-indigo-700 hover:bg-indigo-800 text-white font-medium py-2.5 rounded text-sm flex justify-center items-center gap-2 transition-colors">
          {isAILoading && activeTask === "proofread" ? <Loader2 className="w-4 h-4 animate-spin" /> : <Wand2 className="w-4 h-4" />}
          Soát lỗi & Sửa văn bản
        </button>
        <div className="grid grid-cols-2 gap-2">
          <button onClick={() => handleAITask("summarize")} disabled={isAILoading} className="border bg-white hover:bg-slate-50 text-slate-700 py-2 rounded text-xs transition-colors">Tóm tắt văn bản</button>
          <button onClick={() => handleAITask("quiz")} disabled={isAILoading} className="border bg-white hover:bg-slate-50 text-slate-700 py-2 rounded text-xs transition-colors">Tạo Bài tập (10 câu)</button>
        </div>
      </div>
    </section>

    <section className={cn("bg-white p-5 rounded-xl shadow-sm border border-slate-200", !hasFile && "opacity-50 pointer-events-none")}>
      <h2 className="text-base font-semibold mb-3 text-blue-700">Tải xuống</h2>
      <button onClick={() => exportToWord(previewRef.current, fileName, config)} className="w-full bg-blue-800 hover:bg-blue-900 text-white font-medium py-2.5 rounded flex justify-center items-center gap-2 transition-colors">
        <Save className="w-4 h-4" /> Tải Xuống File Word (.doc)
      </button>
    </section>
  </>
  );
}
