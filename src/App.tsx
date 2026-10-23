import React, { ChangeEvent, useEffect, useRef, useState } from "react";
import {
  AlertTriangle,
  Check,
  Copy,
  Eye,
  EyeOff,
  FileSpreadsheet,
  FileText,
  Key,
  Loader2,
  RefreshCw,
  Settings,
  ShieldCheck,
  Sparkles,
  Upload,
  X,
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
import type { ExtractedTable } from "./utils/spreadsheet";
import type { FormattingConfig } from "./utils/docx";

/**
 * Giao diện theo KẾ HOẠCH GỘP Hải duyệt 23/10/2026 (file
 * Ke_Hoach_Gop_Nang_Cap_App_PDF.docx): cơ chế GIỮ NGUYÊN — thả PDF vào
 * là app tự xử lý (trang có lớp chữ đọc thẳng; trang scan: có khóa
 * Gemini thì AI đánh lại + đối chiếu OCR kiểm tra số liệu, chưa có khóa
 * thì quét OCR), thể thức dùng mặc định chuẩn Nghị định 30 khóa cứng
 * bên dưới. Phần làm mới chỉ là TẦNG HIỂN THỊ: header kính mờ + pill
 * trạng thái AI mở hộp cài đặt khóa, thẻ file có số trang, tiến trình
 * theo trang, 3 thẻ số liệu sau xử lý, nút Excel ghi số bảng (tự khóa
 * khi 0 bảng), xem trước nội dung dạng xếp gọn và nút sao chép văn bản
 * — mọi thứ thêm vào đều đọc từ kết quả xử lý có sẵn (html, rawText,
 * info), không thêm bước thao tác bắt buộc nào.
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

function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

export default function App() {
  const [fileName, setFileName] = useState("Tai_Lieu");
  const [fileDetails, setFileDetails] = useState<{ name: string; size: string } | null>(null);
  const [pageCount, setPageCount] = useState(0);
  const [pdfBlocks, setPdfBlocks] = useState<DocBlock[] | null>(null);
  const [pdfHtml, setPdfHtml] = useState("");
  const [pdfRawText, setPdfRawText] = useState("");
  const [pdfWarnings, setPdfWarnings] = useState<string[]>([]);
  const [extractedTables, setExtractedTables] = useState<ExtractedTable[]>([]);
  const [tableCount, setTableCount] = useState(0);
  const [isConverting, setIsConverting] = useState(false);
  const [convertProgress, setConvertProgress] = useState("");
  const [dragActive, setDragActive] = useState(false);
  const [showPreview, setShowPreview] = useState(false);
  const [previewTab, setPreviewTab] = useState<"text" | "tables">("text");
  const [copied, setCopied] = useState(false);

  const [showSettings, setShowSettings] = useState(false);
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
    setPdfHtml("");
    setPdfRawText("");
    setPdfWarnings([]);
    setExtractedTables([]);
    setTableCount(0);
    setPageCount(0);
    setShowPreview(false);
    setCopied(false);
    setFileDetails({ name: file.name, size: formatFileSize(file.size) });
    try {
      const { convertPdfFile } = await import("./services/pdf");
      const result: PdfConvertResult = await convertPdfFile(
        file,
        (stage, page, total) => {
          setConvertProgress(`${stage} ${page}/${total}...`);
          if (total > 0) setPageCount(total);
        }
      );
      if (result.info.kind === "empty" || result.blocks.length === 0) {
        showNotice("error", "Không trích được nội dung từ PDF này. File có thể bị hỏng, đặt mật khẩu, hoặc là ảnh scan quá mờ để OCR.");
        return;
      }
      const { blocksToTables } = await import("./utils/spreadsheet");
      const tables = blocksToTables(result.blocks);
      setExtractedTables(tables);
      setTableCount(tables.length);
      setPdfBlocks(result.blocks);
      setPdfHtml(result.html);
      setPdfRawText(result.rawText);
      setPdfWarnings(result.warnings);
      setPageCount(result.info.totalPages);
      setFileName(file.name.replace(/\.pdf$/i, ""));
    } catch (error: any) {
      showNotice("error", `Không thể chuyển đổi PDF này: ${error?.message || "lỗi không xác định"}. Nếu file đặt mật khẩu, hãy gỡ mật khẩu rồi thử lại.`);
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
    if (!pdfBlocks || tableCount === 0) return;
    void import("./services/excel")
      .then((m) => m.exportBlocksTablesToExcel(pdfBlocks, fileName))
      .catch(() =>
        showNotice("error", "Không xuất được file Excel. Nếu đang mất mạng và đây là lần dùng đầu tiên, hãy mở app lúc có mạng một lần để app nạp đủ phần cần thiết.")
      );
  };

  const handleCopyText = () => {
    if (!pdfRawText) return;
    navigator.clipboard
      .writeText(pdfRawText)
      .then(() => {
        setCopied(true);
        showNotice("ok", "Đã sao chép văn bản.");
        window.setTimeout(() => setCopied(false), 2500);
      })
      .catch(() => showNotice("error", "Trình duyệt không cho sao chép tự động. Hãy mở xem trước, bôi đen và sao chép thủ công."));
  };

  const saveGeminiKey = () => {
    if (!geminiKeyInput.trim()) return;
    setGeminiKey(geminiKeyInput.trim());
    setGeminiKeySaved(true);
    setGeminiKeyInput("");
    setShowSettings(false);
    showNotice("ok", "Đã lưu khóa Gemini.");
  };

  return (
    <div className="flex min-h-screen flex-col bg-slate-50/60 font-sans text-slate-800">
      <header className="sticky top-0 z-40 border-b border-slate-200/80 bg-white/90 backdrop-blur-md">
        <div className="mx-auto flex w-full max-w-4xl items-center justify-between px-4 py-3">
          <div className="flex items-center gap-3">
            <span className="flex h-10 w-10 shrink-0 items-center justify-center rounded-xl bg-blue-600 text-white shadow-sm">
              <FileText className="h-5 w-5" />
            </span>
            <div>
              <h1 className="text-base font-bold leading-tight text-slate-900">Chuyển PDF sang Word &amp; Excel</h1>
              <p className="hidden text-xs text-slate-500 sm:block">Chuẩn hóa thể thức Nghị định 30/2020/NĐ-CP · Hỗ trợ AI</p>
            </div>
          </div>
          <button
            type="button"
            onClick={() => setShowSettings(true)}
            className={cn(
              "inline-flex shrink-0 items-center gap-1.5 rounded-full border px-3 py-1.5 text-xs font-medium",
              geminiKeySaved
                ? "border-emerald-200 bg-emerald-50 text-emerald-700 hover:bg-emerald-100"
                : "border-amber-200 bg-amber-50 text-amber-700 hover:bg-amber-100"
            )}
          >
            {geminiKeySaved ? (
              <>
                <Sparkles className="h-3.5 w-3.5 text-emerald-600" />
                <span className="hidden sm:inline">AI Gemini sẵn sàng</span>
                <span className="sm:hidden">AI sẵn sàng</span>
              </>
            ) : (
              <>
                <Key className="h-3.5 w-3.5 text-amber-600" />
                <span>Cài khóa AI</span>
              </>
            )}
            <Settings className="ml-0.5 h-3.5 w-3.5 opacity-60" />
          </button>
        </div>
      </header>

      <main className="mx-auto w-full max-w-3xl flex-1 px-4 py-6 md:py-10">
        <section className="rounded-2xl border border-slate-200/80 bg-white p-5 shadow-sm md:p-6">
          <input type="file" id="fileInput" accept=".pdf" className="hidden" onChange={onPickPdf} />

          {!fileDetails ? (
            <label
              htmlFor="fileInput"
              className={cn(
                "flex cursor-pointer flex-col items-center justify-center rounded-2xl border-2 border-dashed px-4 py-12 text-center transition-all",
                dragActive
                  ? "scale-[1.01] border-blue-500 bg-blue-50/60"
                  : "border-slate-300 bg-slate-50/50 hover:border-blue-500 hover:bg-blue-50/40"
              )}
              onDragOver={(e) => {
                e.preventDefault();
                setDragActive(true);
              }}
              onDragLeave={() => setDragActive(false)}
              onDrop={(e) => {
                e.preventDefault();
                setDragActive(false);
                const file = e.dataTransfer.files?.[0];
                if (file) void handlePdfFile(file);
              }}
            >
              <span className="mb-3 flex h-14 w-14 items-center justify-center rounded-2xl bg-blue-100 text-blue-600">
                <Upload className="h-7 w-7" />
              </span>
              <span className="text-base font-semibold text-slate-800">Kéo thả file PDF vào đây hoặc bấm để chọn</span>
              <span className="mt-1 text-xs text-slate-500">Tự động chuẩn hóa thể thức Nghị định 30 và nhận diện bảng biểu.</span>
            </label>
          ) : (
            <div className="rounded-2xl border border-slate-200 bg-slate-50/80 p-4">
              <div className="flex items-center justify-between gap-3">
                <div className="flex min-w-0 items-center gap-3">
                  <span className="flex h-12 w-12 shrink-0 items-center justify-center rounded-xl bg-red-100 text-xs font-bold text-red-600">PDF</span>
                  <div className="min-w-0">
                    <p className="truncate text-sm font-semibold text-slate-900">{fileDetails.name}</p>
                    <p className="text-xs text-slate-500">
                      {fileDetails.size}
                      {pageCount > 0 ? ` · ${pageCount} trang` : ""}
                    </p>
                  </div>
                </div>
                <label
                  htmlFor="fileInput"
                  className="inline-flex shrink-0 cursor-pointer items-center gap-1.5 rounded-lg border border-slate-300 bg-white px-3 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-50"
                >
                  <RefreshCw className="h-3.5 w-3.5 text-slate-500" /> Đổi file
                </label>
              </div>
              {isConverting && (
                <div className="mt-4 border-t border-slate-200 pt-3">
                  <p className="flex items-center gap-2 text-sm font-medium text-blue-700">
                    <Loader2 className="h-4 w-4 animate-spin" /> {convertProgress}
                  </p>
                </div>
              )}
            </div>
          )}

          {pdfWarnings.length > 0 && (
            <div className="mt-4 space-y-1.5 rounded-xl border border-amber-200 bg-amber-50/80 p-3.5 text-left">
              {pdfWarnings.map((w, i) => {
                const severe = /không thành công|không đọc được|lỗi/i.test(w);
                return (
                  <p key={i} className={cn("flex items-start gap-2 text-xs", severe ? "font-medium text-red-700" : "text-amber-800")}>
                    <AlertTriangle className="mt-0.5 h-3.5 w-3.5 shrink-0" /> <span>{w}</span>
                  </p>
                );
              })}
            </div>
          )}

          {pdfBlocks && !isConverting && (
            <div className="mt-6 border-t border-slate-100 pt-5">
              <div className="mb-5 grid grid-cols-3 gap-3">
                <div className="rounded-xl border border-slate-200/80 bg-slate-50/50 p-3 text-center">
                  <span className="block text-xl font-bold text-slate-800">{pdfBlocks.length}</span>
                  <span className="text-[11px] text-slate-500">Khối nội dung</span>
                </div>
                <div className="rounded-xl border border-slate-200/80 bg-slate-50/50 p-3 text-center">
                  <span className="block text-xl font-bold text-blue-700">{tableCount}</span>
                  <span className="text-[11px] text-slate-500">Bảng biểu tìm thấy</span>
                </div>
                <div className="flex flex-col items-center justify-center rounded-xl border border-slate-200/80 bg-slate-50/50 p-3 text-center">
                  <span className="flex items-center gap-1 text-xs font-semibold text-emerald-600">
                    <ShieldCheck className="h-4 w-4" /> Chuẩn NĐ 30
                  </span>
                  <span className="mt-0.5 text-[11px] text-slate-500">Lề &amp; font tự động</span>
                </div>
              </div>

              <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
                <button
                  type="button"
                  onClick={() => void handleExportWord()}
                  className="flex items-center justify-center gap-2 rounded-xl bg-blue-600 px-5 py-3.5 text-sm font-semibold text-white shadow-sm hover:bg-blue-700"
                >
                  <FileText className="h-4 w-4" /> Tải file Word (.docx)
                </button>
                <button
                  type="button"
                  onClick={handleExportExcel}
                  disabled={tableCount === 0}
                  className={cn(
                    "flex items-center justify-center gap-2 rounded-xl px-5 py-3.5 text-sm font-semibold shadow-sm",
                    tableCount > 0
                      ? "bg-emerald-600 text-white hover:bg-emerald-700"
                      : "cursor-not-allowed bg-slate-100 text-slate-400"
                  )}
                >
                  <FileSpreadsheet className="h-4 w-4" />
                  {tableCount > 0 ? `Tải file Excel (${tableCount} bảng)` : "Không có bảng để xuất Excel"}
                </button>
              </div>

              <div className="mt-4 flex items-center justify-between border-t border-slate-100 pt-3 text-xs">
                <button
                  type="button"
                  onClick={() => setShowPreview(!showPreview)}
                  className="inline-flex items-center gap-1.5 font-medium text-blue-700 hover:text-blue-800"
                >
                  {showPreview ? <EyeOff className="h-3.5 w-3.5" /> : <Eye className="h-3.5 w-3.5" />}
                  <span>{showPreview ? "Thu gọn xem trước" : "Xem trước nội dung"}</span>
                </button>
                <button
                  type="button"
                  onClick={handleCopyText}
                  className="inline-flex items-center gap-1.5 font-medium text-slate-600 hover:text-slate-900"
                >
                  {copied ? <Check className="h-3.5 w-3.5 text-emerald-600" /> : <Copy className="h-3.5 w-3.5" />}
                  <span>{copied ? "Đã chép!" : "Sao chép văn bản"}</span>
                </button>
              </div>

              {showPreview && (
                <div className="mt-4 rounded-xl border border-slate-200 bg-slate-50/50 p-4 text-left">
                  <div className="mb-3 flex gap-2 border-b border-slate-200 pb-2">
                    <button
                      type="button"
                      onClick={() => setPreviewTab("text")}
                      className={cn(
                        "rounded-lg px-3 py-1 text-xs font-medium",
                        previewTab === "text" ? "bg-white text-blue-700 shadow-sm" : "text-slate-600 hover:text-slate-900"
                      )}
                    >
                      Văn bản
                    </button>
                    <button
                      type="button"
                      onClick={() => setPreviewTab("tables")}
                      disabled={tableCount === 0}
                      className={cn(
                        "rounded-lg px-3 py-1 text-xs font-medium disabled:opacity-40",
                        previewTab === "tables" ? "bg-white text-blue-700 shadow-sm" : "text-slate-600 hover:text-slate-900"
                      )}
                    >
                      Bảng biểu ({tableCount})
                    </button>
                  </div>
                  {previewTab === "text" ? (
                    <div
                      className="max-h-80 overflow-y-auto rounded-lg border border-slate-200/60 bg-white px-4 py-3 text-sm leading-relaxed text-slate-800 [&_p]:mb-2 [&_table]:my-2 [&_table]:w-full [&_table]:border-collapse [&_td]:border [&_td]:border-slate-200 [&_td]:p-1.5 [&_td]:align-top"
                      dangerouslySetInnerHTML={{ __html: pdfHtml }}
                    />
                  ) : (
                    <div className="max-h-80 space-y-3 overflow-y-auto">
                      {extractedTables.map((t, idx) => (
                        <div key={idx} className="overflow-x-auto rounded-lg border border-slate-200 bg-white p-2">
                          <p className="mb-1 text-xs font-semibold text-slate-800">{t.name}</p>
                          <table className="w-full border-collapse text-[11px]">
                            <tbody>
                              {t.rows.slice(0, 5).map((row, rIdx) => (
                                <tr key={rIdx} className="border-b border-slate-100 last:border-0">
                                  {row.map((cell, cIdx) => (
                                    <td key={cIdx} className="border-r border-slate-100 p-1 last:border-0">{cell}</td>
                                  ))}
                                </tr>
                              ))}
                            </tbody>
                          </table>
                          {t.rows.length > 5 && <p className="mt-1 text-[11px] text-slate-400">… còn {t.rows.length - 5} dòng nữa</p>}
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              )}
            </div>
          )}
        </section>
      </main>

      {showSettings && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4"
          onClick={() => setShowSettings(false)}
        >
          <div
            className="w-full max-w-md rounded-2xl border border-slate-200 bg-white p-6 text-left shadow-xl"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h3 className="text-sm font-bold text-slate-900">Khóa Gemini</h3>
              <button type="button" onClick={() => setShowSettings(false)} className="text-slate-400 hover:text-slate-600" aria-label="Đóng">
                <X className="h-4 w-4" />
              </button>
            </div>
            <div className="mt-4">
              {geminiKeySaved ? (
                <div className="rounded-xl border border-emerald-200 bg-emerald-50/70 p-3.5">
                  <p className="text-xs font-semibold text-emerald-800">Đã lưu khóa Gemini trên trình duyệt này</p>
                  <div className="mt-3 flex gap-2">
                    <button
                      type="button"
                      disabled={aiChecking}
                      onClick={() => void handleCheckAi()}
                      className="rounded-lg border border-emerald-300 bg-white px-3 py-1.5 text-xs font-medium text-emerald-800 disabled:opacity-50"
                    >
                      {aiChecking ? "Đang kiểm tra…" : "Kiểm tra kết nối"}
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        setGeminiKey("");
                        setGeminiKeySaved(false);
                      }}
                      className="rounded-lg border border-red-200 bg-white px-3 py-1.5 text-xs font-medium text-red-700"
                    >
                      Xóa khóa
                    </button>
                  </div>
                </div>
              ) : (
                <div>
                  <label className="block text-xs font-medium text-slate-700" htmlFor="geminiKeyInput">
                    Khóa Gemini —{" "}
                    <a href="https://aistudio.google.com/apikey" target="_blank" rel="noreferrer" className="font-semibold text-blue-700 underline">
                      bấm vào đây để lấy khóa miễn phí
                    </a>
                    :
                  </label>
                  <input
                    id="geminiKeyInput"
                    type="password"
                    value={geminiKeyInput}
                    onChange={(e) => setGeminiKeyInput(e.target.value)}
                    placeholder="Dán khóa Gemini của bạn vào đây"
                    className="mt-2 w-full rounded-xl border border-slate-300 px-3 py-2 text-sm outline-none focus:border-blue-500"
                    autoComplete="off"
                  />
                  <button
                    type="button"
                    onClick={saveGeminiKey}
                    className="mt-3 w-full rounded-xl bg-blue-600 py-2.5 text-sm font-semibold text-white hover:bg-blue-700"
                  >
                    Lưu khóa
                  </button>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      <NoticeBar notice={notice} onClose={() => setNotice(null)} />
    </div>
  );
}
