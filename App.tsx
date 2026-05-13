import React, { ChangeEvent, useEffect, useMemo, useRef, useState } from "react";
import mammoth from "mammoth";
import ReactMarkdown from "react-markdown";
import {
  AlertTriangle,
  CheckCircle2,
  FileQuestion,
  FileText,
  Loader2,
  Save,
  Settings,
  SlidersHorizontal,
  Upload,
  Wand2,
  X,
} from "lucide-react";

import { cn } from "./lib/utils";
import { runAI } from "./services/ai";
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
  description: string;
}> = [
  {
    value: "preserve",
    title: "Giữ nguyên định dạng gốc tối đa",
    description: "Ưu tiên giữ khung/bảng, phù hợp công văn, đơn, biên bản, biểu mẫu.",
  },
  {
    value: "nd30",
    title: "Chuẩn hóa theo NĐ 30",
    description: "Ẩn khung kỹ thuật phần đầu khi xuất Word, căn lại tiêu đề văn bản.",
  },
  {
    value: "textOnly",
    title: "Chỉ lấy nội dung để AI soát lỗi",
    description: "Không cố chuẩn hóa bố cục, dùng khi tài liệu quá phức tạp.",
  },
];

export default function App() {
  const [fileName, setFileName] = useState("Tai_Lieu");
  const [fileStatus, setFileStatus] = useState("Chưa có file nào");
  const [sourceHtml, setSourceHtml] = useState("");
  const [rawText, setRawText] = useState("");
  const [docHtml, setDocHtml] = useState("");
  const [hasFile, setHasFile] = useState(false);
  const [isLegacyDoc, setIsLegacyDoc] = useState(false);

  const [processingMode, setProcessingMode] = useState<DocumentProcessingMode>("preserve");
  const [showPreviewFrame, setShowPreviewFrame] = useState(false);
  const [documentStats, setDocumentStats] = useState<DocumentStats>(EMPTY_STATS);

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

  const modeSummary = useMemo(() => {
    if (processingMode === "nd30") return "Ẩn khung kỹ thuật khi xuất Word, căn lại tiêu đề.";
    if (processingMode === "textOnly") return "Chỉ dùng nội dung chữ cho AI.";
    return "Giữ định dạng gốc tối đa, không xóa viền bảng theo vị trí.";
  }, [processingMode]);

  const handleFileUpload = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

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

        if (stats.administrativeHeaderDetected) {
          setProcessingMode("preserve");
          setShowPreviewFrame(false);
        }
      } catch {
        alert("Không thể đọc file này. Nếu là file .doc, hãy mở bằng Word rồi Save As sang .docx.");
        setFileStatus("Lỗi định dạng file");
        setHasFile(false);
      }
    };
    reader.readAsArrayBuffer(file);
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

  const disabledSection = !hasFile ? "opacity-50 pointer-events-none" : "";

  return (
    <div className="min-h-screen bg-slate-50 p-4 md:p-8 font-sans text-slate-800">
      <div className="max-w-7xl mx-auto">
        <header className="mb-8 border-b border-slate-300 pb-4">
          <h1 className="text-3xl font-bold text-blue-800 flex items-center gap-2">
            <FileText className="w-8 h-8" /> <span>Chuẩn Hóa Văn Bản & Soát Lỗi AI</span>
          </h1>
          <p className="text-slate-600 mt-2 text-sm">
            Công cụ hỗ trợ tải file Word, soát lỗi văn phong, căn chỉnh theo NĐ 30/2020/NĐ-CP và xuất lại file.
          </p>
        </header>

        <main className="grid grid-cols-1 lg:grid-cols-12 gap-8 items-start">
          <aside className="lg:col-span-4 space-y-5">
            <section className="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
              <h2 className="text-base font-semibold mb-3 flex items-center gap-2 text-blue-700">
                <span className="bg-blue-100 text-blue-800 w-6 h-6 rounded flex items-center justify-center text-sm">1</span>
                Nguồn Tài Liệu
              </h2>
              <label htmlFor="fileInput" className="cursor-pointer bg-slate-50 border-2 border-dashed border-slate-300 hover:border-blue-500 hover:bg-blue-50 transition-colors rounded-lg p-8 flex flex-col items-center">
                <Upload className="w-7 h-7 text-slate-500 mb-2" />
                <span className="text-sm text-slate-700 font-medium">Tải lên file (.docx)</span>
              </label>
              <input type="file" id="fileInput" accept=".docx,.doc" className="hidden" onChange={handleFileUpload} />
              <div className="text-xs mt-3 text-center">
                {hasFile ? (
                  <span className="text-emerald-600 flex flex-col items-center gap-1">
                    <span className="font-medium italic break-all flex items-center gap-1">
                      <CheckCircle2 className="w-3 h-3" /> {fileStatus.replace("Đã tải: ", "")}
                    </span>
                  </span>
                ) : (
                  <span className="text-slate-500">{fileStatus}</span>
                )}
                {isLegacyDoc && <p className="mt-2 text-amber-600">Nên chuyển file .doc sang .docx trước khi xử lý.</p>}
              </div>
            </section>

            <section className={cn("bg-white p-5 rounded-xl shadow-sm border border-slate-200", disabledSection)}>
              <h2 className="text-base font-semibold mb-3 flex items-center gap-2 text-blue-700">
                <span className="bg-blue-100 text-blue-800 w-6 h-6 rounded flex items-center justify-center text-sm">2</span>
                Chế Độ Giữ Khung/Bảng
              </h2>
              <p className="text-xs text-slate-600 mb-3">Kiểu xử lý sau khi upload</p>
              <div className="space-y-2">
                {processingModes.map((mode) => (
                  <label key={mode.value} className={cn("block rounded-lg border p-3 cursor-pointer transition", processingMode === mode.value ? "border-blue-500 bg-blue-50" : "border-slate-200 bg-white hover:bg-slate-50")}>
                    <div className="flex gap-2 items-start">
                      <input type="radio" className="mt-1" checked={processingMode === mode.value} onChange={() => setProcessingMode(mode.value)} />
                      <div>
                        <div className="text-sm font-semibold text-slate-800">{mode.title}</div>
                        <div className="text-xs text-slate-500 leading-5">{mode.description}</div>
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

            <section className={cn("bg-white p-5 rounded-xl shadow-sm border border-slate-200", disabledSection)}>
              <h2 className="text-base font-semibold mb-3 flex items-center gap-2 text-blue-700">
                <span className="bg-blue-100 text-blue-800 w-6 h-6 rounded flex items-center justify-center text-sm">3</span>
                Căn Lề NĐ 30
              </h2>
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

            <section className={cn("bg-white p-5 rounded-xl shadow-sm border border-slate-200", disabledSection)}>
              <h2 className="text-base font-semibold mb-3 flex items-center gap-2 text-blue-700">
                <span className="bg-blue-100 text-blue-800 w-6 h-6 rounded flex items-center justify-center text-sm">4</span>
                Soát Lỗi AI
              </h2>
              <div className="space-y-3">
                <select className="w-full border p-2 rounded text-sm text-slate-700" value={aiStyle} onChange={(event) => setAiStyle(event.target.value)}>
                  <option value="grammar">Sửa Ngữ pháp & Nối dòng</option>
                  <option value="admin">Chuẩn hóa Hành chính (NĐ 30)</option>
                  <option value="edu">Chuẩn hóa Giáo dục (Sư phạm)</option>
                </select>
                <div className="grid grid-cols-2 gap-2 text-sm">
                  <label className="flex items-center gap-2 border rounded p-2"><input type="radio" checked={aiMode === "suggest"} onChange={() => setAiMode("suggest")} /> Chỉ báo lỗi</label>
                  <label className="flex items-center gap-2 border rounded p-2"><input type="radio" checked={aiMode === "autofix"} onChange={() => setAiMode("autofix")} /> AI tự sửa</label>
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

            <section className={cn("bg-white p-5 rounded-xl shadow-sm border border-slate-200", disabledSection)}>
              <h2 className="text-base font-semibold mb-3 flex items-center gap-2 text-blue-700">
                <span className="bg-blue-100 text-blue-800 w-6 h-6 rounded flex items-center justify-center text-sm">5</span>
                Hoàn Tất
              </h2>
              <button onClick={() => exportToWord(previewRef.current, fileName, config)} className="w-full bg-blue-800 hover:bg-blue-900 text-white font-medium py-2.5 rounded flex justify-center items-center gap-2 transition-colors">
                <Save className="w-4 h-4" /> Tải Xuống File Word (.doc)
              </button>
            </section>
          </aside>

          <section className="lg:col-span-8 flex flex-col min-w-0">
            <div className="mb-4">
              <h2 className="text-lg font-bold text-slate-800">Bản Xem Trước (A4)</h2>
              <p className="text-xs text-slate-500 mt-1 flex items-center gap-1">
                <CheckCircle2 className="w-3 h-3" /> {modeSummary}
              </p>
            </div>

            <div className="bg-slate-300 p-4 md:p-8 rounded-xl overflow-auto max-h-[calc(100vh-140px)] min-h-[720px]">
              <div className="mx-auto bg-white shadow-xl relative shrink-0" style={{ width: "794px", minHeight: "1123px" }}>
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
                  
                  /* ĐỊNH DẠNG BẢNG CHUNG */
                  .document-content table {
                    width: 100% !important;
                    max-width: 100% !important;
                    table-layout: fixed !important;
                    border-collapse: collapse !important;
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
                    overflow-wrap: break-word !important;
                  }
                  .document-content table p,
                  .document-content td p,
                  .document-content th p {
                    text-indent: 0cm !important;
                    margin: 0 0 3pt 0 !important;
                    line-height: 1.15 !important;
                  }

                  /* ============================================================== */
                  /* SỬA LỖI NĐ 30: KHÓA TỶ LỆ CỘT VÀ XÓA VIỀN (Dành cho Word Export) */
                  /* ============================================================== */
                  .document-content .admin-header-table,
                  .document-content table:first-of-type,
                  .document-content table:last-of-type {
                    border: none !important;
                    background: transparent !important;
                  }
                  
                  /* Cột Trái: Tên Cơ Quan (Chiếm 40%) */
                  .document-content .admin-header-table td:first-child,
                  .document-content table:first-of-type td:first-child,
                  .document-content table:last-of-type td:first-child {
                    border: none !important;
                    width: 40% !important;
                    text-align: center !important;
                    vertical-align: top !important;
                  }

                  /* Cột Phải: Quốc Hiệu (Chiếm 60%) */
                  .document-content .admin-header-table td:last-child,
                  .document-content table:first-of-type td:last-child,
                  .document-content table:last-of-type td:last-child {
                    border: none !important;
                    width: 60% !important;
                    text-align: center !important;
                    vertical-align: top !important;
                  }

                  /* Khung hỗ trợ (nếu tích chọn ở Bước 2) */
                  .document-content .admin-header-preview-frame td {
                    border: 1px dashed #94a3b8 !important;
                  }

                  .document-content .admin-header-table p,
                  .document-content table:first-of-type p {
                    font-family: "Times New Roman", serif !important;
                    text-align: center !important;
                    text-indent: 0cm !important;
                    margin: 0 0 2pt 0 !important;
                    line-height: 1.15 !important;
                  }
                  
                  /* Format font NĐ 30 */
                  .document-content .admin-agency-line { font-size: 12pt !important; font-weight: normal !important; text-transform: uppercase; }
                  .document-content .admin-unit-line { font-size: 12.5pt !important; font-weight: bold !important; text-transform: uppercase; }
                  .document-content .admin-number-line { font-size: 12pt !important; font-weight: normal !important; }
                  .document-content .admin-national-line { font-size: 12.5pt !important; font-weight: bold !important; text-transform: uppercase; }
                  .document-content .admin-motto-line { font-size: 13pt !important; font-weight: bold !important; display: inline-block; border-bottom: 1px solid #000; padding-bottom: 1pt; }
                  .document-content .admin-date-line { font-size: 12pt !important; font-style: italic; }
                  
                  .document-content .doc-main-title {
                    text-align: center !important;
                    text-indent: 0cm !important;
                    font-weight: bold !important;
                    text-transform: uppercase;
                    margin-top: 12pt !important;
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
                  <div className="absolute inset-0 flex flex-col items-center justify-center text-slate-400 py-32">
                    <FileQuestion className="w-16 h-16 mb-4 text-slate-200" />
                    <p className="text-sm">Tải file .docx để xem trước văn bản tại đây.</p>
                  </div>
                ) : (
                  <div ref={previewRef} className="document-content outline-none" dangerouslySetInnerHTML={{ __html: docHtml }} />
                )}
              </div>
            </div>
          </section>
        </main>
      </div>

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
