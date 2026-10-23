import React, { ChangeEvent, useEffect, useRef, useState } from "react";
import {
  AlertTriangle,
  FileSpreadsheet,
  FileText,
  Loader2,
  Upload,
} from "lucide-react";

import { cn } from "./lib/utils";
import { NoticeBar } from "./components/NoticeBar";
import { setGeminiKey, hasGeminiKey, checkAiConnection } from "./services/ai";
// Tách gói tải theo nhu cầu: các module nặng (pdf.js + Tesseract trong
// services/pdf, xlsx trong services/excel + utils/spreadsheet, docx trong
// utils/docxExport) KHÔNG nạp khi mở app; mỗi thao tác tự nạp module của
// nó bằng import() động ngay trước khi dùng. Ở đây chỉ giữ import KIỂU
// (type) — bị xoá lúc biên dịch, không kéo theo mã thật.
import type { PdfConvertResult } from "./services/pdf";
import type { DocBlock } from "./utils/pdfStructure";
import type { FormattingConfig } from "./utils/docx";

/**
 * Giao diện TỐI GIẢN một màn hình (yêu cầu của Hải 23/10/2026 — "ẩn tất
 * cả, thành cơ chế tự động, nhìn lằng nhằng quá"): tải PDF lên là app tự
 * xử lý (trang có lớp chữ đọc thẳng; trang scan: có khóa Gemini thì AI
 * đánh lại + đối chiếu OCR kiểm tra số liệu, chưa có khóa thì quét OCR),
 * xong chỉ hiện 2 nút tải file Word / Excel. Thể thức lề lối dùng mặc
 * định chuẩn Nghị định 30 bên dưới — không còn mục chỉnh, không xem
 * trước, không mục Word & AI, không đọc file bảng tính.
 */
const DEFAULT_CONFIG: FormattingConfig = {
  font: "Times New Roman",
  size: 14,
  spacing: "1.5",
  textAlign: "justify",
  textIndent: 1.27,
  paraSpacing: 6,
  leftMargin: 3,
  rightMargin: 2,
};

export default function App() {
  const [fileName, setFileName] = useState("Tai_Lieu");
  const [fileStatus, setFileStatus] = useState("Chưa có file nào");
  const [pdfBlocks, setPdfBlocks] = useState<DocBlock[] | null>(null);
  const [pdfWarnings, setPdfWarnings] = useState<string[]>([]);
  const [tableCount, setTableCount] = useState(0);
  const [isConverting, setIsConverting] = useState(false);
  const [convertProgress, setConvertProgress] = useState("");

  const [geminiKeyInput, setGeminiKeyInput] = useState("");
  const [geminiKeySaved, setGeminiKeySaved] = useState<boolean>(() => hasGeminiKey());
  const [aiChecking, setAiChecking] = useState(false);
  const handleCheckAi = async () => {
    setAiChecking(true);
    try {
      const result = await checkAiConnection();
      showNotice(result.ok ? "ok" : "error", result.message);
    } finally {
      setAiChecking(false);
    }
  };

  // Thanh báo trong trang (thay cho alert): đỏ = lỗi, vàng = cảnh báo, xanh = xong.
  const [notice, setNotice] = useState<{ kind: "error" | "warn" | "ok"; text: string } | null>(null);
  const noticeTimer = useRef<number | undefined>(undefined);
  const showNotice = (kind: "error" | "warn" | "ok", text: string) => {
    setNotice({ kind, text });
    window.clearTimeout(noticeTimer.current);
    noticeTimer.current = window.setTimeout(() => setNotice(null), 7000);
  };

  // Nạp TRƯỚC các gói tách ra khi trình duyệt rảnh — trang vẫn mở nhanh vì
  // không phải chờ chúng, nhưng service worker kịp lưu vào bộ nhớ đệm nên
  // app vẫn chạy offline sau lần mở có mạng đầu tiên (không nạp trước thì
  // lần đầu xuất file lúc mất mạng sẽ thiếu gói).
  useEffect(() => {
    const warm = () => {
      void import("./services/pdf").catch(() => undefined);
      void import("./services/excel").catch(() => undefined);
      void import("./utils/docxExport").catch(() => undefined);
      void import("./utils/spreadsheet").catch(() => undefined);
    };
    const ric = (window as any).requestIdleCallback as
      | ((cb: () => void, opts?: { timeout: number }) => number)
      | undefined;
    if (ric) {
      const id = ric(warm, { timeout: 4000 });
      return () => (window as any).cancelIdleCallback?.(id);
    }
    const t = window.setTimeout(warm, 2500);
    return () => window.clearTimeout(t);
  }, []);

  const handlePdfFile = async (file: File) => {
    // Chưa có khóa Gemini cũng không chặn: trang có lớp chữ vẫn đọc thẳng
    // bình thường; riêng trang scan sẽ tự rơi xuống OCR kèm cảnh báo rõ
    // dưới kết quả (xử lý trong services/pdf.ts).
    setIsConverting(true);
    setConvertProgress("Đang mở file PDF...");
    setPdfBlocks(null);
    setPdfWarnings([]);
    try {
      const { convertPdfFile } = await import("./services/pdf");
      const result: PdfConvertResult = await convertPdfFile(
        file,
        (stage, page, total) => {
          setConvertProgress(`${stage} ${page}/${total}...`);
        }
      );
      if (result.info.kind === "empty" || result.blocks.length === 0) {
        showNotice("error", "Không trích được nội dung từ PDF này. File có thể bị hỏng, đặt mật khẩu, hoặc là ảnh scan quá mờ để OCR.");
        setFileStatus("Lỗi: không đọc được PDF");
        return;
      }
      const { blocksToTables } = await import("./utils/spreadsheet");
      setTableCount(blocksToTables(result.blocks).length);
      setPdfBlocks(result.blocks);
      setPdfWarnings(result.warnings);
      setFileName(file.name.replace(/\.pdf$/i, ""));
      setFileStatus(`Đã tải: ${file.name}`);
    } catch (error: any) {
      showNotice("error", `Không thể chuyển đổi PDF này: ${error?.message || "lỗi không xác định"}. Nếu file đặt mật khẩu, hãy gỡ mật khẩu rồi thử lại.`);
      setFileStatus("Lỗi định dạng file PDF");
    } finally {
      setIsConverting(false);
      setConvertProgress("");
    }
  };

  const onPickPdf = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (file) void handlePdfFile(file);
  };

  const handleExportWord = async () => {
    if (!pdfBlocks) return;
    try {
      const { exportBlocksToWord } = await import("./utils/docxExport");
      await exportBlocksToWord(pdfBlocks, fileName, DEFAULT_CONFIG);
    } catch {
      showNotice("error", "Không xuất được file Word. Nếu đang mất mạng và đây là lần dùng đầu tiên, hãy mở app lúc có mạng một lần để app nạp đủ phần cần thiết.");
    }
  };

  const handleExportExcel = () => {
    if (!pdfBlocks) return;
    if (tableCount === 0) {
      showNotice("warn", "PDF này không có bảng nào để xuất Excel.");
      return;
    }
    void import("./services/excel")
      .then((m) => m.exportBlocksTablesToExcel(pdfBlocks, fileName))
      .catch(() =>
        showNotice("error", "Không xuất được file Excel. Nếu đang mất mạng và đây là lần dùng đầu tiên, hãy mở app lúc có mạng một lần để app nạp đủ phần cần thiết.")
      );
  };

  return (
    <div className="min-h-screen bg-slate-100 text-slate-800">
      <header className="bg-gradient-to-r from-blue-950 via-blue-900 to-blue-800 text-white shadow">
        <div className="mx-auto flex w-full max-w-xl items-center gap-3 px-4 py-3">
          <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-lg bg-white/15">
            <FileText className="h-5 w-5" />
          </span>
          <h1 className="text-base font-bold leading-tight md:text-lg">Chuyển PDF sang Word &amp; Excel</h1>
        </div>
      </header>

      <main className="mx-auto w-full max-w-xl px-3 py-4 md:py-8">
        <section className="rounded-xl border border-slate-200 bg-white p-5 shadow-sm">
          <label
            htmlFor="fileInput"
            className="flex cursor-pointer flex-col items-center rounded-lg border-2 border-dashed border-slate-300 bg-slate-50 p-8 transition-colors hover:border-blue-500 hover:bg-blue-50"
            onDragOver={(e) => e.preventDefault()}
            onDrop={(e) => {
              e.preventDefault();
              const file = e.dataTransfer.files?.[0];
              if (file) void handlePdfFile(file);
            }}
          >
            <Upload className="h-10 w-10 text-blue-700" />
            <span className="mt-2 text-sm font-semibold text-slate-700">Tải lên file PDF</span>
          </label>
          <input type="file" id="fileInput" accept=".pdf" className="hidden" onChange={onPickPdf} />

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

          {isConverting && (
            <p className="mt-3 flex items-center gap-2 text-sm font-medium text-blue-700">
              <Loader2 className="h-4 w-4 animate-spin" /> {convertProgress}
            </p>
          )}
          {!isConverting && <p className="mt-3 text-xs text-slate-500">{fileStatus}</p>}

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

          {pdfBlocks && !isConverting && (
            <div className="mt-4 grid grid-cols-1 gap-2 sm:grid-cols-2">
              <button type="button" onClick={() => void handleExportWord()} className="flex items-center justify-center gap-2 rounded-lg bg-blue-700 px-4 py-3 text-sm font-semibold text-white hover:bg-blue-800">
                <FileText className="h-4 w-4" /> Tải file Word (.docx)
              </button>
              <button type="button" onClick={handleExportExcel} className="flex items-center justify-center gap-2 rounded-lg bg-emerald-600 px-4 py-3 text-sm font-semibold text-white hover:bg-emerald-700">
                <FileSpreadsheet className="h-4 w-4" /> Tải file Excel (.xlsx)
              </button>
            </div>
          )}
        </section>
      </main>

      <NoticeBar notice={notice} onClose={() => setNotice(null)} />
    </div>
  );
}
