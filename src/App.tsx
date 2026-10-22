import React, { ChangeEvent, useEffect, useMemo, useRef, useState } from "react";
import mammoth from "mammoth";
import ReactMarkdown from "react-markdown";
import {
  AlertTriangle,
  CheckCircle2,
  Eye,
  FileQuestion,
  FileSpreadsheet,
  FileText,
  Loader2,
  Save,
  Settings,
  SlidersHorizontal,
  Upload,
  Wand2,
  Wrench,
  X,
} from "lucide-react";

import { cn } from "./lib/utils";
import { runAI, getGeminiKey, setGeminiKey, hasGeminiKey } from "./services/ai";
import { convertPdfFile, type PdfConvertResult, type ScanEngine } from "./services/pdf";
import { convertExcelFile } from "./services/excel";
import { exportBlocksTablesToExcel } from "./services/excel";
import { exportBlocksToWord } from "./utils/docxExport";
import { blocksToTables, type SheetInfo } from "./utils/spreadsheet";
import type { DocBlock, PdfClassification } from "./utils/pdfStructure";
import {
  DocumentProcessingMode,
  DocumentStats,
  FormattingConfig,
  analyzeDocumentHtml,
  enhanceAdministrativeDocumentHtml,
  exportToWord,
  processHtmlToRemoveBullets,
} from "./utils/docx";

const EMPTY_STATS: DocumentStats = {
  tableCount: 0,
  imageCount: 0,
  firstTableHasBorderRisk: false,
  administrativeHeaderDetected: false,
  signatureTableDetected: false,
  warningMessages: [],
};

const processingModes: Array<{
  value: DocumentProcessingMode;
  title: string;
}> = [
  { value: "preserve", title: "Giữ nguyên định dạng gốc tối đa" },
  { value: "nd30", title: "Chuẩn hóa theo NĐ 30" },
  { value: "textOnly", title: "Chỉ lấy nội dung để AI soát lỗi" },
];

/**
 * Giao diện MỘT MÀN HÌNH (như app phiếu học tập): thanh trên + vùng làm
 * việc 2 cột — cột trái là công cụ, cột phải là bản xem trước A4, mỗi cột
 * tự cuộn. Ba MỤC RIÊNG biệt, không chung đụng dữ liệu/nút bấm:
 *  1. "PDF → Word": chỉ nhận PDF, chỉ xuất Word .docx chuẩn NĐ 30.
 *  2. "Excel": bảng tính → Word/Excel và trích bảng từ PDF → Excel.
 *  3. "Word & AI": chuẩn hoá file Word có sẵn + soát lỗi AI.
 * Mỗi mục giữ kết quả riêng (snapshot) — chuyển mục không mất việc đang làm.
 */
type Section = "pdf" | "excel" | "word";

interface DocSnapshot {
  sourceHtml: string;
  rawText: string;
  fileName: string;
  fileStatus: string;
  hasFile: boolean;
  isLegacyDoc: boolean;
  stats: DocumentStats;
}

const EMPTY_SNAPSHOT: DocSnapshot = {
  sourceHtml: "",
  rawText: "",
  fileName: "Tai_Lieu",
  fileStatus: "Chưa có file nào",
  hasFile: false,
  isLegacyDoc: false,
  stats: EMPTY_STATS,
};

const SECTION_TABS: Array<{ value: Section; label: string; icon: React.ReactNode }> = [
  { value: "pdf", label: "PDF → Word", icon: <FileText className="w-4 h-4" /> },
  { value: "excel", label: "Excel", icon: <FileSpreadsheet className="w-4 h-4" /> },
  { value: "word", label: "Word & AI", icon: <Wand2 className="w-4 h-4" /> },
];

export default function App() {
  const [section, setSection] = useState<Section>("pdf");
  const [mobilePane, setMobilePane] = useState<"tools" | "preview">("tools");
  const snapshotsRef = useRef<Record<Section, DocSnapshot>>({
    pdf: EMPTY_SNAPSHOT,
    excel: EMPTY_SNAPSHOT,
    word: EMPTY_SNAPSHOT,
  });

  // --- Trạng thái tài liệu đang hiển thị (của mục đang mở) ---
  const [fileName, setFileName] = useState("Tai_Lieu");
  const [fileStatus, setFileStatus] = useState("Chưa có file nào");
  const [sourceHtml, setSourceHtml] = useState("");
  const [rawText, setRawText] = useState("");
  const [docHtml, setDocHtml] = useState("");
  const [hasFile, setHasFile] = useState(false);
  const [isLegacyDoc, setIsLegacyDoc] = useState(false);
  const [documentStats, setDocumentStats] = useState<DocumentStats>(EMPTY_STATS);

  const [processingMode, setProcessingMode] = useState<DocumentProcessingMode>("preserve");
  const [showPreviewFrame, setShowPreviewFrame] = useState(false);

  // --- Kết quả RIÊNG của từng mục ---
  const [pdfBlocks, setPdfBlocks] = useState<DocBlock[] | null>(null);
  const [pdfInfo, setPdfInfo] = useState<PdfClassification | null>(null);
  const [pdfWarnings, setPdfWarnings] = useState<string[]>([]);
  const [pdfOcrConfidence, setPdfOcrConfidence] = useState<number | undefined>(undefined);
  const [pdfAiPages, setPdfAiPages] = useState<number | undefined>(undefined);
  // Cách đọc trang scan: "ai" = AI đánh lại (chuẩn cao, cần khóa Gemini);
  // "ocr" = quét OCR trên máy (offline). Mặc định AI theo yêu cầu của Hải.
  const [scanEngine, setScanEngine] = useState<ScanEngine>("ai");
  const [geminiKeyInput, setGeminiKeyInput] = useState("");
  const [geminiKeySaved, setGeminiKeySaved] = useState<boolean>(() => hasGeminiKey());

  // Thanh báo trong trang (thay cho alert): đỏ = lỗi, vàng = cảnh báo, xanh = xong.
  const [notice, setNotice] = useState<{ kind: "error" | "warn" | "ok"; text: string } | null>(null);
  const noticeTimer = useRef<number | undefined>(undefined);
  const showNotice = (kind: "error" | "warn" | "ok", text: string) => {
    setNotice({ kind, text });
    window.clearTimeout(noticeTimer.current);
    noticeTimer.current = window.setTimeout(() => setNotice(null), 7000);
  };

  // Xem trước: thu nhỏ trang A4 vừa bề rộng khung (mặc định bật khi màn hình nhỏ).
  const [previewFit, setPreviewFit] = useState<boolean>(
    () => typeof window !== "undefined" && window.innerWidth < 1024
  );
  const [previewScale, setPreviewScale] = useState(1);
  const previewScrollRef = useRef<HTMLDivElement>(null);

  const [excelBlocks, setExcelBlocks] = useState<DocBlock[] | null>(null);
  const [excelSheets, setExcelSheets] = useState<SheetInfo[] | null>(null);
  const [excelWarnings, setExcelWarnings] = useState<string[]>([]);

  const [pdfTableBlocks, setPdfTableBlocks] = useState<DocBlock[] | null>(null);
  const [pdfTableName, setPdfTableName] = useState("");

  const [isConverting, setIsConverting] = useState(false);
  const [convertProgress, setConvertProgress] = useState("");

  const [config, setConfig] = useState<FormattingConfig>({
    font: "Times New Roman",
    size: 14,
    spacing: "1.5",
    textAlign: "justify",
    textIndent: 1.27,
    paraSpacing: 6,
    leftMargin: 3,
    rightMargin: 2,
  });

  const [isAIModalOpen, setIsAIModalOpen] = useState(false);
  const [isAILoading, setIsAILoading] = useState(false);
  const [activeTask, setActiveTask] = useState<string | null>(null);
  const [aiResult, setAiResult] = useState("");
  const [aiError, setAiError] = useState("");
  const [aiMode, setAiMode] = useState("suggest");
  const [aiStyle, setAiStyle] = useState("edu");

  const previewRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const el = previewScrollRef.current;
    if (!el) return;
    const update = () => {
      if (!previewFit) {
        setPreviewScale(1);
        return;
      }
      const w = el.clientWidth;
      if (w > 0) setPreviewScale(Math.max(0.3, Math.min(1, (w - 36) / 794)));
    };
    update();
    const ro = new ResizeObserver(update);
    ro.observe(el);
    return () => ro.disconnect();
  }, [previewFit]);

  useEffect(() => {
    if (!sourceHtml) {
      setDocHtml("");
      setDocumentStats(EMPTY_STATS);
      return;
    }
    const { html, stats } = enhanceAdministrativeDocumentHtml(sourceHtml, {
      mode: processingMode,
      preserveFirstFrame: showPreviewFrame,
    });
    setDocHtml(html);
    setDocumentStats(stats);
  }, [sourceHtml, processingMode, showPreviewFrame]);

  const pdfTableCount = useMemo(
    () => (pdfTableBlocks ? blocksToTables(pdfTableBlocks).length : 0),
    [pdfTableBlocks]
  );
  const excelTableCount = useMemo(
    () => (excelBlocks ? blocksToTables(excelBlocks).length : 0),
    [excelBlocks]
  );

  const modeSummary = useMemo(() => {
    if (section === "pdf") return "PDF (có lớp chữ, scan OCR, hỗn hợp) → Word .docx chuẩn NĐ 30.";
    if (section === "excel") return "Bảng tính → Word/Excel · Trích bảng từ PDF → Excel.";
    if (processingMode === "nd30") return "Ẩn khung kỹ thuật khi xuất Word, căn lại tiêu đề.";
    if (processingMode === "textOnly") return "Chỉ dùng nội dung chữ cho AI.";
    return "Giữ định dạng gốc tối đa, không xóa viền bảng theo vị trí.";
  }, [processingMode, section]);

  /** Chuyển mục: lưu kết quả mục hiện tại, khôi phục kết quả của mục đích. */
  const switchSection = (next: Section) => {
    if (next === section) return;
    snapshotsRef.current[section] = {
      sourceHtml,
      rawText,
      fileName,
      fileStatus,
      hasFile,
      isLegacyDoc,
      stats: documentStats,
    };
    const snap = snapshotsRef.current[next];
    setSection(next);
    setSourceHtml(snap.sourceHtml);
    setRawText(snap.rawText);
    setFileName(snap.fileName);
    setFileStatus(snap.fileStatus);
    setHasFile(snap.hasFile);
    setIsLegacyDoc(snap.isLegacyDoc);
    setDocumentStats(snap.stats);
    setMobilePane("tools");
  };

  /** Nạp một tài liệu đã trích xuất (html + chữ thô) vào vùng xem trước. */
  const loadConverted = (html: string, text: string, name: string) => {
    const stats = analyzeDocumentHtml(html);
    setRawText(text);
    setSourceHtml(html);
    setDocumentStats(stats);
    setFileName(name);
    setHasFile(true);
    setFileStatus(`Đã tải: ${name}`);
    setIsLegacyDoc(false);
    if (stats.administrativeHeaderDetected && section !== "word") {
      setProcessingMode("nd30");
      setShowPreviewFrame(false);
    }
    setMobilePane("preview");
  };

  // ---------------- Mục 1: PDF → Word ----------------
  const handlePdfFile = async (file: File) => {
    // Không chặn cứng khi chọn AI mà chưa có khóa: trang có lớp chữ vẫn đọc
    // thẳng bình thường; riêng trang scan sẽ tự rơi xuống OCR kèm cảnh báo
    // rõ trong panel kết quả (xử lý trong services/pdf.ts).
    setIsConverting(true);
    setConvertProgress("Đang mở file PDF...");
    setPdfAiPages(undefined);
    try {
      const result: PdfConvertResult = await convertPdfFile(
        file,
        (stage, page, total) => {
          setConvertProgress(`${stage} ${page}/${total}...`);
        },
        { scanEngine }
      );
      if (result.info.kind === "empty" || result.blocks.length === 0) {
        showNotice("error", "Không trích được nội dung từ PDF này. File có thể bị hỏng, đặt mật khẩu, hoặc là ảnh scan quá mờ để OCR.");
        setFileStatus("Lỗi: không đọc được PDF");
        setHasFile(false);
        return;
      }
      setPdfBlocks(result.blocks);
      setPdfInfo(result.info);
      setPdfWarnings(result.warnings);
      setPdfOcrConfidence(result.ocrConfidence);
      setPdfAiPages(result.aiPages);
      loadConverted(result.html, result.rawText, file.name.replace(/\.pdf$/i, ""));
    } catch (error: any) {
      showNotice("error", `Không thể chuyển đổi PDF này: ${error?.message || "lỗi không xác định"}. Nếu file đặt mật khẩu, hãy gỡ mật khẩu rồi thử lại.`);
      setFileStatus("Lỗi định dạng file PDF");
      setHasFile(false);
    } finally {
      setIsConverting(false);
      setConvertProgress("");
    }
  };

  // ---------------- Mục 2: Excel ----------------
  const handleExcelSheetFile = async (file: File) => {
    setIsConverting(true);
    setConvertProgress("Đang đọc file bảng tính...");
    try {
      const result = await convertExcelFile(file);
      setExcelBlocks(result.blocks);
      setExcelSheets(result.sheets);
      setExcelWarnings(result.warnings);
      loadConverted(result.html, result.rawText, file.name.replace(/\.(xlsx|xls|csv)$/i, ""));
    } catch (error: any) {
      showNotice("error", error?.message || "Không thể đọc file bảng tính này.");
      setFileStatus("Lỗi định dạng file bảng tính");
      setHasFile(false);
    } finally {
      setIsConverting(false);
      setConvertProgress("");
    }
  };

  const handlePdfTablesFile = async (file: File) => {
    setIsConverting(true);
    setConvertProgress("Đang trích bảng từ PDF...");
    try {
      // Dùng chung cách đọc trang scan với mục PDF → Word; chưa có khóa
      // Gemini thì công cụ này tự quét OCR (không chặn việc trích bảng).
      const engine: ScanEngine = scanEngine === "ai" && !hasGeminiKey() ? "ocr" : scanEngine;
      const result: PdfConvertResult = await convertPdfFile(
        file,
        (stage, page, total) => {
          setConvertProgress(`${stage} ${page}/${total}...`);
        },
        { scanEngine: engine }
      );
      if (result.blocks.length === 0) {
        showNotice("error", "Không trích được nội dung từ PDF này.");
        return;
      }
      setPdfTableBlocks(result.blocks);
      setPdfTableName(file.name);
      loadConverted(result.html, result.rawText, file.name.replace(/\.pdf$/i, ""));
    } catch (error: any) {
      showNotice("error", `Không thể đọc PDF này: ${error?.message || "lỗi không xác định"}.`);
    } finally {
      setIsConverting(false);
      setConvertProgress("");
    }
  };

  // ---------------- Mục 3: Word & AI ----------------
  const handleWordFile = (file: File) => {
    const extension = file.name.split(".").pop()?.toLowerCase();
    setFileName(file.name.replace(/\.(docx|doc)$/i, ""));
    setSourceHtml("");
    setDocHtml("");
    setRawText("");
    setDocumentStats(EMPTY_STATS);
    if (extension === "doc") {
      setIsLegacyDoc(true);
      setFileStatus("Cảnh báo: File .doc cũ có thể lỗi. Nên lưu lại thành .docx.");
    } else {
      setIsLegacyDoc(false);
      setFileStatus(`Đã tải: ${file.name}`);
    }
    const reader = new FileReader();
    reader.onload = async (readerEvent) => {
      try {
        const arrayBuffer = readerEvent.target?.result as ArrayBuffer;
        const htmlResult = await mammoth.convertToHtml(
          { arrayBuffer },
          {
            convertImage: mammoth.images.inline(async (element: any) => {
              const imageBuffer = await element.read("base64");
              return { src: `data:${element.contentType};base64,${imageBuffer}` };
            }),
          }
        );
        const textResult = await mammoth.extractRawText({ arrayBuffer });
        const cleanedHtml = processHtmlToRemoveBullets(htmlResult.value);
        const stats = analyzeDocumentHtml(cleanedHtml);
        setRawText(textResult.value);
        setSourceHtml(cleanedHtml);
        setDocumentStats(stats);
        setHasFile(true);
        setMobilePane("preview");
        if (stats.administrativeHeaderDetected) {
          setProcessingMode("preserve");
          setShowPreviewFrame(false);
        }
      } catch {
        showNotice("error", "Không thể đọc file này. Nếu là file .doc, hãy mở bằng Word rồi Save As sang .docx.");
        setFileStatus("Lỗi định dạng file");
        setHasFile(false);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const onPick = (handler: (f: File) => void) => (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (file) handler(file);
  };

  const handleAITask = async (taskType: "summarize" | "quiz" | "proofread") => {
    if (!rawText.trim()) return;
    setActiveTask(taskType);
    setIsAIModalOpen(true);
    setIsAILoading(true);
    setAiResult("");
    setAiError("");

    let prompt = "";
    if (taskType === "summarize") {
      prompt = "Tóm tắt ngắn gọn, gạch đầu dòng các ý chính yếu của văn bản sau:\n\n" + rawText;
    } else if (taskType === "quiz") {
      prompt = `Dựa tối đa vào văn bản dưới đây, tạo phiếu bài tập trắc nghiệm 10 câu, có đáp án cuối bài.\n\nVăn bản nguồn:\n\"\"\"\n${rawText}\n\"\"\"`;
    } else {
      const modeLine = aiMode === "suggest" ? "CHỈ BÁO LỖI, KHÔNG VIẾT LẠI TOÀN BỘ." : "TỰ ĐỘNG SỬA VĂN BẢN.";
      const styleLine =
        aiStyle === "admin"
          ? "Văn phong hành chính, chuẩn mực theo thể thức văn bản nhà nước."
          : aiStyle === "edu"
          ? "Văn phong giáo dục, chuẩn mực sư phạm, rõ ràng, dễ áp dụng."
          : "Ưu tiên sửa ngữ pháp, chính tả, dấu câu, nối dòng sai.";
      prompt = `Bạn là chuyên gia biên tập văn bản tiếng Việt. ${modeLine}\nPhong cách: ${styleLine}\nYêu cầu: chỉ phân tích/sửa nội dung chữ, không phá bảng, không tự ý thêm thông tin ngoài văn bản.\n\nVăn bản:\n\"\"\"\n${rawText}\n\"\"\"`;
    }

    try {
      const response = await runAI(prompt);
      setAiResult(response);
    } catch (error: any) {
      setAiError(error?.message || "Lỗi AI. Vui lòng thử lại.");
    } finally {
      setIsAILoading(false);
      setActiveTask(null);
    }
  };

  const appendAIToDocument = () => {
    if (!aiResult.trim()) return;
    const formattedHtml = aiResult
      .replace(/### (.*)/g, "<b>$1</b>")
      .replace(/## (.*)/g, "<b>$1</b>")
      .replace(/\*\*(.*?)\*\*/g, "<b>$1</b>")
      .split(/\n\n/)
      .map((paragraph) => `<p class="ai-text-line">${paragraph.trim().replace(/\n/g, "<br>")}</p>`)
      .join("");
    const aiBlock = `<br><div class="ai-wrapper-box"><p class="ai-title-line">--- BẢN GHI TỪ TRỢ LÝ AI ---</p>${formattedHtml}</div>`;
    setDocHtml((previous) => previous + aiBlock);
    setIsAIModalOpen(false);
  };

  const uploadHint = (label: string) => (
    <>
      <Upload className="w-7 h-7 text-slate-500 mb-2" />
      <span className="text-sm text-slate-700 font-medium">{label}</span>
    </>
  );

  const statusLine = (
    <div className="text-xs mt-3 text-center">
      {hasFile ? (
        <span className="text-emerald-600 font-medium italic break-all inline-flex items-center gap-1">
          <CheckCircle2 className="w-3 h-3 shrink-0" /> {fileStatus.replace("Đã tải: ", "")}
        </span>
      ) : (
        <span className="text-slate-500">{fileStatus}</span>
      )}
      {isLegacyDoc && <p className="mt-2 text-amber-600">Nên chuyển file .doc sang .docx trước khi xử lý.</p>}
    </div>
  );

  const convertingLine = isConverting && (
    <p className="text-xs mt-3 text-center text-blue-700 flex items-center justify-center gap-2">
      <Loader2 className="w-3.5 h-3.5 animate-spin" /> {convertProgress || "Đang xử lý..."}
    </p>
  );

  return (
    <div className="h-screen flex flex-col bg-slate-100 font-sans text-slate-800 overflow-hidden">
      {/* ===== Thanh trên ===== */}
      <header className="bg-white border-b border-slate-200 px-4 md:px-6 py-3 flex items-center gap-3 shrink-0">
        <span className="bg-blue-800 text-white w-10 h-10 rounded-lg flex items-center justify-center shrink-0">
          <FileText className="w-5 h-5" />
        </span>
        <div className="min-w-0">
          <h1 className="text-lg md:text-xl font-bold text-blue-800 leading-tight truncate">Chuẩn Hóa Văn Bản & Soát Lỗi AI</h1>
        </div>
        <span className="ml-auto hidden md:inline-flex text-xs font-medium bg-blue-50 text-blue-800 border border-blue-100 rounded-full px-3 py-1">NĐ 30/2020/NĐ-CP</span>
      </header>

      {/* ===== Tab vùng làm việc trên điện thoại ===== */}
      <div className="lg:hidden bg-white border-b border-slate-200 grid grid-cols-2 shrink-0" role="tablist" aria-label="Vùng làm việc">
        <button role="tab" aria-selected={mobilePane === "tools"} onClick={() => setMobilePane("tools")} className={cn("py-2.5 text-sm font-medium flex items-center justify-center gap-2", mobilePane === "tools" ? "text-blue-800 border-b-2 border-blue-800" : "text-slate-500")}>
          <Wrench className="w-4 h-4" /> Công cụ
        </button>
        <button role="tab" aria-selected={mobilePane === "preview"} onClick={() => setMobilePane("preview")} className={cn("py-2.5 text-sm font-medium flex items-center justify-center gap-2", mobilePane === "preview" ? "text-blue-800 border-b-2 border-blue-800" : "text-slate-500")}>
          <Eye className="w-4 h-4" /> Xem trước {hasFile && <span className="w-2 h-2 rounded-full bg-emerald-500" />}
        </button>
      </div>

      {/* ===== Vùng làm việc 1 màn hình ===== */}
      <main className="flex-1 grid grid-cols-1 lg:grid-cols-[400px_minmax(0,1fr)] min-h-0">
        {/* ----- Cột công cụ ----- */}
        <aside className={cn("min-h-0 overflow-y-auto p-4 space-y-4", mobilePane === "tools" ? "block" : "hidden", "lg:block")}>
          {/* Chọn mục */}
          <div className="grid grid-cols-3 gap-1 bg-slate-200/70 rounded-xl p-1" role="tablist" aria-label="Chọn mục công việc">
            {SECTION_TABS.map((tab) => (
              <button
                key={tab.value}
                role="tab"
                aria-selected={section === tab.value}
                onClick={() => switchSection(tab.value)}
                className={cn(
                  "flex items-center justify-center gap-1.5 py-2 rounded-lg text-sm font-medium transition",
                  section === tab.value ? "bg-white text-blue-800 shadow-sm" : "text-slate-600 hover:text-slate-800"
                )}
              >
                {tab.icon} {tab.label}
              </button>
            ))}
          </div>

          {section === "pdf" && (
            <>
              <section className="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
                <h2 className="text-base font-semibold mb-3 text-blue-700">Chuyển PDF sang Word</h2>
                <label htmlFor="fileInput" className="cursor-pointer bg-slate-50 border-2 border-dashed border-slate-300 hover:border-blue-500 hover:bg-blue-50 transition-colors rounded-lg p-8 flex flex-col items-center">
                  {uploadHint("Tải lên file PDF")}
                </label>
                <input type="file" id="fileInput" accept=".pdf" className="hidden" onChange={onPick((f) => void handlePdfFile(f))} />
                <div className="mt-3 flex flex-wrap items-center gap-x-2 gap-y-2 text-left">
                  <span className="text-xs font-medium text-slate-600">Trang scan:</span>
                  <div className="inline-flex overflow-hidden rounded-lg border border-slate-300">
                    <button type="button" onClick={() => setScanEngine("ai")} className={cn("px-3 py-1.5 text-sm font-medium transition-colors", scanEngine === "ai" ? "bg-blue-700 text-white" : "bg-white text-slate-600 hover:bg-slate-50")}>AI đánh lại</button>
                    <button type="button" onClick={() => setScanEngine("ocr")} className={cn("border-l border-slate-300 px-3 py-1.5 text-sm font-medium transition-colors", scanEngine === "ocr" ? "bg-blue-700 text-white" : "bg-white text-slate-600 hover:bg-slate-50")}>Quét OCR</button>
                  </div>
                  {scanEngine === "ai" && geminiKeySaved && (
                    <span className="ml-auto text-xs text-slate-500">
                      Đã lưu khóa Gemini ·{" "}
                      <button type="button" className="font-medium text-blue-700 underline" onClick={() => { setGeminiKey(""); setGeminiKeySaved(false); }}>Xóa</button>
                    </span>
                  )}
                </div>
                {scanEngine === "ai" && !geminiKeySaved && (
                  <div className="mt-2 rounded-lg border border-blue-100 bg-blue-50 p-3 text-left">
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
                <button onClick={() => pdfBlocks && void exportBlocksToWord(pdfBlocks, fileName, config)} className="w-full bg-blue-800 hover:bg-blue-900 text-white font-medium py-2.5 rounded flex justify-center items-center gap-2 transition-colors">
                  <Save className="w-4 h-4" /> Tải File Word (.docx) Chuẩn NĐ 30
                </button>

              </section>
            </>
          )}

          {section === "excel" && (
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
                    <button onClick={() => void exportBlocksToWord(excelBlocks, fileName, config)} className="w-full bg-blue-800 hover:bg-blue-900 text-white font-medium py-2.5 rounded flex justify-center items-center gap-2 transition-colors">
                      <Save className="w-4 h-4" /> Tải File Word (.docx) Chuẩn NĐ 30
                    </button>
                    <button onClick={() => exportBlocksTablesToExcel(excelBlocks, fileName)} className="w-full bg-emerald-700 hover:bg-emerald-800 text-white font-medium py-2.5 rounded flex justify-center items-center gap-2 transition-colors">
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
                    <button onClick={() => exportBlocksTablesToExcel(pdfTableBlocks, fileName)} disabled={pdfTableCount === 0} className="w-full bg-emerald-700 hover:bg-emerald-800 disabled:opacity-50 text-white font-medium py-2.5 rounded flex justify-center items-center gap-2 transition-colors">
                      <Save className="w-4 h-4" /> Tải Excel (.xlsx) — {pdfTableCount} bảng đã trích
                    </button>
                    {pdfTableCount === 0 && <p className="text-xs text-amber-700 mt-2">PDF này không có bảng nào trích được (bản scan thì bảng chỉ đọc được dạng dòng chữ — xem ở mục PDF → Word).</p>}
                  </div>
                )}
              </section>
            </>
          )}

          {section === "word" && (
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
          )}

          {/* Cài đặt xuất Word — dùng chung cho cả 3 mục */}
          <section className="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
            <h2 className="text-base font-semibold mb-3 text-blue-700">Căn Lề NĐ 30 (áp dụng khi tải Word)</h2>
            <div className="space-y-3">
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs text-slate-600 mb-1">Font chữ</label>
                  <select className="w-full border p-2 rounded text-sm" value={config.font} onChange={(event) => setConfig({ ...config, font: event.target.value })}>
                    <option value="Times New Roman">Times New Roman</option>
                    <option value="Arial">Arial</option>
                  </select>
                </div>
                <div>
                  <label className="block text-xs text-slate-600 mb-1">Cỡ chữ</label>
                  <input type="number" className="w-full border p-2 rounded text-sm" value={config.size} onChange={(event) => setConfig({ ...config, size: parseFloat(event.target.value) || 14 })} />
                </div>
              </div>
              <div className="grid grid-cols-2 gap-3 bg-emerald-50 p-2 rounded border border-emerald-100">
                <div>
                  <label className="block text-xs text-slate-600 mb-1">Thụt dòng (cm)</label>
                  <input type="number" step="0.1" className="w-full border p-2 rounded text-sm" value={config.textIndent} onChange={(event) => setConfig({ ...config, textIndent: parseFloat(event.target.value) || 0 })} />
                </div>
                <div>
                  <label className="block text-xs text-slate-600 mb-1">Cách đoạn (pt)</label>
                  <input type="number" className="w-full border p-2 rounded text-sm" value={config.paraSpacing} onChange={(event) => setConfig({ ...config, paraSpacing: parseFloat(event.target.value) || 0 })} />
                </div>
              </div>
              <div className="grid grid-cols-3 gap-2">
                <div>
                  <label className="block text-xs text-slate-600 mb-1">Giãn dòng</label>
                  <select className="w-full border p-2 rounded text-sm" value={config.spacing} onChange={(event) => setConfig({ ...config, spacing: event.target.value })}>
                    <option value="1.0">1.0</option>
                    <option value="1.15">1.15</option>
                    <option value="1.5">1.5</option>
                  </select>
                </div>
                <div>
                  <label className="block text-xs text-slate-600 mb-1">Lề Trái</label>
                  <input type="number" step="0.5" className="w-full border p-2 rounded text-sm" value={config.leftMargin} onChange={(event) => setConfig({ ...config, leftMargin: parseFloat(event.target.value) || 3 })} />
                </div>
                <div>
                  <label className="block text-xs text-slate-600 mb-1">Lề Phải</label>
                  <input type="number" step="0.5" className="w-full border p-2 rounded text-sm" value={config.rightMargin} onChange={(event) => setConfig({ ...config, rightMargin: parseFloat(event.target.value) || 2 })} />
                </div>
              </div>
            </div>
          </section>
        </aside>

        {/* ----- Cột xem trước ----- */}
        <section className={cn("min-h-0 flex flex-col min-w-0", mobilePane === "preview" ? "flex" : "hidden", "lg:flex")}>
          <div className="px-4 md:px-6 pt-4 pb-2 shrink-0 flex flex-wrap items-start gap-2">
            <div className="min-w-0 flex-1 basis-48">
              <h2 className="text-lg font-bold text-slate-800">Bản Xem Trước (A4)</h2>
              <p className="text-xs text-slate-500 mt-1 flex items-center gap-1">
                <CheckCircle2 className="w-3 h-3 shrink-0" /> {modeSummary}
              </p>
            </div>
            <div className="ml-auto inline-flex shrink-0 overflow-hidden rounded-lg border border-slate-300">
              <button type="button" onClick={() => setPreviewFit(true)} className={cn("px-2.5 py-1 text-xs font-medium transition-colors", previewFit ? "bg-blue-700 text-white" : "bg-white text-slate-600")}>Vừa màn hình</button>
              <button type="button" onClick={() => setPreviewFit(false)} className={cn("border-l border-slate-300 px-2.5 py-1 text-xs font-medium transition-colors", !previewFit ? "bg-blue-700 text-white" : "bg-white text-slate-600")}>100%</button>
            </div>
          </div>
          <div ref={previewScrollRef} className="flex-1 min-h-0 bg-slate-300 mx-4 md:mx-6 mb-4 md:mb-6 rounded-xl overflow-auto">
            {/* Bọc trang A4 trong lớp "w-max min-w-full flex justify-center":
                trên điện thoại trang 794px rộng hơn màn hình — để mx-auto trực
                tiếp trong khung cuộn sẽ làm mép TRÁI tràn ra vùng âm, không
                cuộn tới được (chữ bị cắt đầu dòng như ảnh chụp 22/10). Lớp bọc
                rộng bằng nội dung + căn giữa khi vừa khung thì cuộn được cả
                2 phía. */}
            <div className="min-w-full w-max flex justify-center p-4 md:p-8">
            <div className="bg-white shadow-xl relative shrink-0" style={{ width: "794px", minHeight: "1123px", zoom: previewScale }}>
              <style>{`
                .document-content {
                  box-sizing: border-box;
                  width: 794px;
                  min-height: 1123px;
                  padding: 2cm ${config.rightMargin}cm 2cm ${config.leftMargin}cm;
                  background: white;
                  color: black;
                  text-align: justify;
                  overflow: visible;
                }
                .document-content p:not(.ai-wrapper-box p) {
                  font-family: "${config.font}", serif;
                  font-size: ${config.size}pt;
                  line-height: ${config.spacing};
                  text-align: ${config.textAlign};
                  text-indent: ${config.textIndent}cm;
                  margin-top: 0pt;
                  margin-bottom: ${config.paraSpacing}pt;
                }
                .document-content table {
                  width: 100% !important;
                  max-width: 100% !important;
                  table-layout: fixed !important;
                  border-collapse: collapse;
                  margin-top: 6pt;
                  margin-bottom: 12pt;
                  box-sizing: border-box;
                }
                .document-content td,
                .document-content th {
                  border: 1px solid #000;
                  padding: 4pt;
                  vertical-align: top;
                  word-wrap: break-word !important;
                  word-break: break-word !important;
                  overflow-wrap: break-word !important;
                }
                .document-content table p,
                .document-content td p,
                .document-content th p {
                  text-indent: 0cm !important;
                  margin: 0 0 3pt 0 !important;
                  line-height: 1.15 !important;
                }
                .document-content .admin-header-table {
                  width: calc(100% + 16px) !important;
                  margin-left: -8px !important;
                  margin-right: -8px !important;
                  margin-top: 0 !important;
                  margin-bottom: 12pt !important;
                  border: none !important;
                  table-layout: fixed !important;
                  background: transparent !important;
                }
                .document-content .admin-header-table col.admin-left-col { width: 40% !important; }
                .document-content .admin-header-table col.admin-right-col { width: 60% !important; }
                .document-content .admin-header-table td,
                .document-content .admin-header-table th {
                  border: none !important;
                  background: transparent !important;
                  padding: 0pt 3pt !important;
                  text-align: center !important;
                  vertical-align: top !important;
                  overflow: hidden !important;
                }
                .document-content table:first-of-type:not(.admin-header-table) td {
                  width: 50% !important;
                  border: none !important;
                }
                .document-content table:first-of-type:not(.admin-header-table) td p {
                  text-align: center !important;
                  margin-bottom: 2pt !important;
                }
                .document-content table:last-of-type:not(.admin-header-table) td {
                  width: 50% !important;
                  border: none !important;
                }
                .document-content table:last-of-type:not(.admin-header-table) td:last-child p {
                  text-align: center !important;
                }
                .document-content .admin-header-preview-frame td {
                  border: 1px dashed #94a3b8 !important;
                }
                .document-content .admin-header-table p {
                  font-family: "Times New Roman", serif !important;
                  text-align: center !important;
                  text-indent: 0cm !important;
                  margin: 0 0 2pt 0 !important;
                  line-height: 1.12 !important;
                  white-space: normal !important;
                }
                .document-content .admin-agency-line { font-size: 12pt !important; font-weight: normal !important; text-transform: uppercase; }
                .document-content .admin-unit-line { font-size: 12.5pt !important; font-weight: bold !important; text-transform: uppercase; }
                .document-content .admin-number-line { font-size: 13pt !important; font-weight: normal !important; }
                .document-content .admin-national-line { font-size: 12pt !important; font-weight: bold !important; letter-spacing: -0.25pt; text-transform: uppercase; }
                .document-content .admin-motto-line { font-size: 13pt !important; font-weight: bold !important; display: inline-block; border-bottom: 1px solid #000; padding-bottom: 1pt; }
                .document-content .admin-date-line { font-size: 12pt !important; font-style: italic; }
                .document-content .doc-main-title {
                  text-align: center !important;
                  text-indent: 0cm !important;
                  font-weight: bold !important;
                  text-transform: uppercase;
                  margin-top: 8pt !important;
                  margin-bottom: 4pt !important;
                  line-height: 1.2 !important;
                }
                .document-content .doc-sub-title {
                  text-align: center !important;
                  text-indent: 0cm !important;
                  font-weight: bold !important;
                  margin-top: 0 !important;
                  margin-bottom: 4pt !important;
                  line-height: 1.2 !important;
                }
                .document-content img { max-width: 100%; height: auto; display: inline-block; }
                .ai-wrapper-box { border-top: 1px dashed #000; margin-top: 20pt; padding-top: 10pt; }
                .ai-title-line { text-align: center !important; text-indent: 0 !important; font-weight: bold; }
              `}</style>

              {!hasFile ? (
                <div className="absolute inset-0 flex flex-col items-center justify-center text-slate-400 py-32 px-8 text-center">
                  <FileQuestion className="w-16 h-16 mb-4 text-slate-200" />
                  <p className="text-sm">
                    {section === "pdf" && "Tải file PDF ở mục PDF → Word để xem trước văn bản tại đây."}
                    {section === "excel" && "Tải file bảng tính hoặc PDF có bảng ở mục Excel để xem trước tại đây."}
                    {section === "word" && "Tải file .docx ở mục Word & AI để xem trước văn bản tại đây."}
                  </p>
                </div>
              ) : (
                <div ref={previewRef} className="document-content outline-none" dangerouslySetInnerHTML={{ __html: docHtml }} />
              )}
            </div>
            </div>
          </div>
        </section>
      </main>

      {notice && (
        <div
          className={cn(
            "fixed bottom-4 left-1/2 z-50 flex w-[calc(100%-2rem)] max-w-xl -translate-x-1/2 items-start gap-2 rounded-lg px-4 py-3 text-sm text-white shadow-2xl",
            notice.kind === "error" && "bg-red-600",
            notice.kind === "warn" && "bg-amber-500",
            notice.kind === "ok" && "bg-emerald-600"
          )}
          role="alert"
        >
          {notice.kind === "ok" ? (
            <CheckCircle2 className="mt-0.5 h-4 w-4 shrink-0" />
          ) : (
            <AlertTriangle className="mt-0.5 h-4 w-4 shrink-0" />
          )}
          <span className="min-w-0 flex-1">{notice.text}</span>
          <button type="button" aria-label="Đóng thông báo" onClick={() => setNotice(null)} className="shrink-0 opacity-80 hover:opacity-100">
            <X className="h-4 w-4" />
          </button>
        </div>
      )}

      {isAIModalOpen && (
        <div className="fixed inset-0 bg-slate-900/50 flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-4xl flex flex-col max-h-[90vh]">
            <div className="p-4 border-b flex justify-between items-center bg-blue-50">
              <h3 className="font-bold text-blue-900 flex items-center gap-2"><Settings className="w-4 h-4" /> Kết Quả Phân Tích AI</h3>
              <button onClick={() => setIsAIModalOpen(false)}><X className="w-5 h-5 text-slate-400" /></button>
            </div>
            <div className="p-6 overflow-y-auto flex-1 bg-slate-50">
              {isAILoading ? (
                <div className="text-center py-20 text-blue-600"><Loader2 className="animate-spin w-8 h-8 mx-auto mb-2" /> Đang xử lý...</div>
              ) : aiError ? (
                <div className="text-red-600 bg-red-50 p-4 rounded border border-red-200">❌ {aiError}</div>
              ) : (
                <div className="prose max-w-none bg-white p-4 border rounded"><ReactMarkdown>{aiResult}</ReactMarkdown></div>
              )}
            </div>
            <div className="p-4 border-t flex justify-end gap-2 bg-white">
              <button onClick={() => setIsAIModalOpen(false)} className="px-4 py-2 border rounded text-sm">Đóng</button>
              {!isAILoading && !aiError && <button onClick={appendAIToDocument} className="px-4 py-2 bg-blue-600 text-white rounded text-sm">Ghi vào văn bản</button>}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
