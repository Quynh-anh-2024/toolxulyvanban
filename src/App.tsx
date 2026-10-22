import React, { ChangeEvent, useEffect, useMemo, useRef, useState } from "react";
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
import { PdfSection } from "./sections/PdfSection";
import { ExcelSection } from "./sections/ExcelSection";
import { WordSection } from "./sections/WordSection";
import { MarginConfig } from "./components/MarginConfig";
import { PreviewPane } from "./components/PreviewPane";
import { NoticeBar } from "./components/NoticeBar";
import { runAI, getGeminiKey, setGeminiKey, hasGeminiKey, checkAiConnection } from "./services/ai";
// B8 — Tách gói tải theo nhu cầu: các module nặng (pdf.js + Tesseract trong
// services/pdf, xlsx trong services/excel + utils/spreadsheet, docx trong
// utils/docxExport, mammoth) KHÔNG nạp khi mở app; mỗi thao tác tự nạp
// module của nó bằng import() động ngay trước khi dùng. Ở đây chỉ giữ
// import KIỂU (type) — bị xoá lúc biên dịch, không kéo theo mã thật.
import type { PdfConvertResult, ScanEngine } from "./services/pdf";
import type { SheetInfo } from "./utils/spreadsheet";
import type { DocBlock, PdfClassification } from "./utils/pdfStructure";
import {
  DocumentProcessingMode,
  DocumentStats,
  FormattingConfig,
  analyzeDocumentHtml,
  enhanceAdministrativeDocumentHtml,
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

  // B8 (tiếp): nạp TRƯỚC các gói tách ra khi trình duyệt rảnh — trang vẫn
  // mở nhanh vì không phải chờ chúng, nhưng service worker kịp lưu vào bộ
  // nhớ đệm nên mọi mục vẫn chạy offline sau lần mở có mạng đầu tiên
  // (không nạp trước thì lần đầu xuất Word lúc mất mạng sẽ thiếu gói).
  useEffect(() => {
    const warm = () => {
      void import("./services/pdf").catch(() => undefined);
      void import("./services/excel").catch(() => undefined);
      void import("./utils/docxExport").catch(() => undefined);
      void import("./utils/spreadsheet").catch(() => undefined);
      void import("mammoth").catch(() => undefined);
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

  // Số bảng đếm tại thời điểm nạp file bằng module spreadsheet nạp động
  // (trước đây tính bằng useMemo — cách đó kéo cả thư viện xlsx vào gói
  // mở app; xem ghi chú B8 ở đầu file).
  const [pdfTableCount, setPdfTableCount] = useState(0);
  const [excelTableCount, setExcelTableCount] = useState(0);
  const countTables = async (blocks: DocBlock[]): Promise<number> => {
    const { blocksToTables } = await import("./utils/spreadsheet");
    return blocksToTables(blocks).length;
  };
  const handleExportBlocksWord = async (blocks: DocBlock[]) => {
    try {
      const { exportBlocksToWord } = await import("./utils/docxExport");
      await exportBlocksToWord(blocks, fileName, config);
    } catch {
      showNotice("error", "Không xuất được file Word. Nếu đang mất mạng và đây là lần dùng đầu tiên, hãy mở app lúc có mạng một lần để app nạp đủ phần cần thiết.");
    }
  };
  const handleExportTablesExcel = (blocks: DocBlock[]) => {
    void import("./services/excel")
      .then((m) => m.exportBlocksTablesToExcel(blocks, fileName))
      .catch(() =>
        showNotice("error", "Không xuất được file Excel. Nếu đang mất mạng và đây là lần dùng đầu tiên, hãy mở app lúc có mạng một lần để app nạp đủ phần cần thiết.")
      );
  };

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
      const { convertPdfFile } = await import("./services/pdf");
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
      const { convertExcelFile } = await import("./services/excel");
      const result = await convertExcelFile(file);
      setExcelBlocks(result.blocks);
      setExcelTableCount(await countTables(result.blocks));
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
      const { convertPdfFile } = await import("./services/pdf");
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
      setPdfTableCount(await countTables(result.blocks));
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
        const mammoth = (await import("mammoth")).default;
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
            <PdfSection
              uploadHint={uploadHint}
              onPick={onPick}
              handlePdfFile={handlePdfFile}
              scanEngine={scanEngine}
              setScanEngine={setScanEngine}
              geminiKeySaved={geminiKeySaved}
              setGeminiKeySaved={setGeminiKeySaved}
              geminiKeyInput={geminiKeyInput}
              setGeminiKeyInput={setGeminiKeyInput}
              aiChecking={aiChecking}
              handleCheckAi={handleCheckAi}
              showNotice={showNotice}
              convertingLine={convertingLine}
              statusLine={statusLine}
              pdfInfo={pdfInfo}
              pdfAiPages={pdfAiPages}
              pdfOcrConfidence={pdfOcrConfidence}
              pdfWarnings={pdfWarnings}
              pdfBlocks={pdfBlocks}
              handleExportBlocksWord={handleExportBlocksWord}
            />
          )}

          {section === "excel" && (
            <ExcelSection
              uploadHint={uploadHint}
              onPick={onPick}
              handleExcelSheetFile={handleExcelSheetFile}
              convertingLine={convertingLine}
              statusLine={statusLine}
              excelSheets={excelSheets}
              excelWarnings={excelWarnings}
              excelBlocks={excelBlocks}
              excelTableCount={excelTableCount}
              handleExportBlocksWord={handleExportBlocksWord}
              handleExportTablesExcel={handleExportTablesExcel}
              handlePdfTablesFile={handlePdfTablesFile}
              pdfTableBlocks={pdfTableBlocks}
              pdfTableCount={pdfTableCount}
              pdfTableName={pdfTableName}
            />
          )}

          {section === "word" && (
            <WordSection
              uploadHint={uploadHint}
              onPick={onPick}
              handleWordFile={handleWordFile}
              statusLine={statusLine}
              hasFile={hasFile}
              processingModes={processingModes}
              processingMode={processingMode}
              setProcessingMode={setProcessingMode}
              showPreviewFrame={showPreviewFrame}
              setShowPreviewFrame={setShowPreviewFrame}
              documentStats={documentStats}
              aiStyle={aiStyle}
              setAiStyle={setAiStyle}
              aiMode={aiMode}
              setAiMode={setAiMode}
              handleAITask={handleAITask}
              isAILoading={isAILoading}
              activeTask={activeTask}
              previewRef={previewRef}
              fileName={fileName}
              config={config}
            />
          )}

          {/* Cài đặt xuất Word — dùng chung cho cả 3 mục */}
          <MarginConfig config={config} setConfig={setConfig} />
        </aside>

        {/* ----- Cột xem trước ----- */}
        <PreviewPane
          mobilePane={mobilePane}
          modeSummary={modeSummary}
          previewFit={previewFit}
          setPreviewFit={setPreviewFit}
          previewScrollRef={previewScrollRef}
          previewScale={previewScale}
          previewRef={previewRef}
          hasFile={hasFile}
          section={section}
          docHtml={docHtml}
          config={config}
        />
      </main>

      <NoticeBar notice={notice} onClose={() => setNotice(null)} />

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
