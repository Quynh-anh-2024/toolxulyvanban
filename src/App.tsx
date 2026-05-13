import React, { useState, useRef, ChangeEvent } from "react";
import mammoth from "mammoth";
import ReactMarkdown from "react-markdown";
import {
  FileText,
  Save,
  Upload,
  X,
  FileQuestion,
  Loader2,
  CheckCircle2,
  AlertTriangle,
  Wand2,
  SlidersHorizontal,
  Table2,
} from "lucide-react";

import { cn } from "./lib/utils";
import { runAI } from "./services/ai";
import {
  FormattingConfig,
  DocumentProcessingMode,
  DocumentStats,
  exportToWord,
  processHtmlToRemoveBullets,
  enhanceAdministrativeDocumentHtml,
} from "./utils/docx";

const emptyStats: DocumentStats = {
  tableCount: 0,
  imageCount: 0,
  firstTableHasBorderRisk: false,
  administrativeHeaderDetected: false,
  signatureTableDetected: false,
  warningMessages: [],
};

export default function App() {
  const [fileName, setFileName] = useState("Tai_Lieu");
  const [fileStatus, setFileStatus] = useState("Chưa có file nào");
  const [rawText, setRawText] = useState("");
  const [docHtml, setDocHtml] = useState("");
  const [sourceHtml, setSourceHtml] = useState("");
  const [hasFile, setHasFile] = useState(false);
  const [stats, setStats] = useState<DocumentStats>(emptyStats);

  const [processingMode, setProcessingMode] = useState<DocumentProcessingMode>("preserve");
  const [showPreviewFrame, setShowPreviewFrame] = useState(false);

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
  const [isExporting, setIsExporting] = useState(false);

  const previewRef = useRef<HTMLDivElement>(null);

  const rebuildDocumentHtml = (html: string, mode = processingMode, previewFrame = showPreviewFrame) => {
    const enhanced = enhanceAdministrativeDocumentHtml(html, {
      mode,
      preserveFirstFrame: previewFrame,
    });
    setDocHtml(enhanced.html);
    setStats(enhanced.stats);
  };

  const handleProcessingModeChange = (mode: DocumentProcessingMode) => {
    setProcessingMode(mode);
    if (sourceHtml) rebuildDocumentHtml(sourceHtml, mode, showPreviewFrame);
  };

  const handlePreviewFrameChange = (checked: boolean) => {
    setShowPreviewFrame(checked);
    if (sourceHtml) rebuildDocumentHtml(sourceHtml, processingMode, checked);
  };

  const handleFileUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const extension = file.name.split(".").pop()?.toLowerCase();
    setFileName(file.name.replace(/\.(docx|doc)$/i, ""));
    setFileStatus(extension === "doc" ? "Cảnh báo: File .doc cũ có thể bị lỗi." : `Đã tải: ${file.name}`);

    const reader = new FileReader();
    reader.onload = async (event) => {
      try {
        const arrayBuffer = event.target?.result as ArrayBuffer;
        const result = await mammoth.convertToHtml(
          { arrayBuffer },
          {
            convertImage: mammoth.images.inline(async (element: any) => {
              const imageBuffer = await element.read("base64");
              return { src: `data:${element.contentType};base64,${imageBuffer}` };
            }),
          }
        );
        const textResult = await mammoth.extractRawText({ arrayBuffer });
        const cleanedHtml = processHtmlToRemoveBullets(result.value);
        setRawText(textResult.value);
        setSourceHtml(cleanedHtml);
        rebuildDocumentHtml(cleanedHtml, processingMode, showPreviewFrame);
        setHasFile(true);
      } catch (err) {
        alert("Lỗi: Không thể đọc file này. Nếu là file .doc, hãy mở bằng Word rồi Save As sang .docx.");
        setFileStatus("Lỗi định dạng file");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleAITask = async (taskType: "summarize" | "quiz" | "proofread") => {
    if (!rawText) return;
    setActiveTask(taskType);
    setIsAIModalOpen(true);
    setIsAILoading(true);
    setAiResult("");
    setAiError("");

    let prompt = "";
    if (taskType === "summarize") {
      prompt = "Tóm tắt ngắn gọn, gạch đầu dòng các ý chính yếu của văn bản sau:\n\n" + rawText;
    } else if (taskType === "quiz") {
      prompt = `Dựa TỐI ĐA vào văn bản dưới đây, tạo PHIẾU BÀI TẬP TRẮC NGHIỆM 10 câu.\n\nVăn bản nguồn:\n\"\"\"\n${rawText}\n\"\"\"`;
    } else {
      prompt = `Đóng vai chuyên gia văn bản hành chính và giáo dục. Rà soát văn bản dưới đây.\n`;
      prompt += aiMode === "suggest" ? `CHẾ ĐỘ: Chỉ báo lỗi, nêu vị trí và cách sửa.\n` : `CHẾ ĐỘ: Tự động viết lại phần lỗi nhưng không phá bố cục.\n`;
      prompt += `Phong cách ưu tiên: ${aiStyle}.\n\nVĂN BẢN:\n\"\"\"\n${rawText}\n\"\"\"`;
    }

    try {
      const response = await runAI(prompt);
      setAiResult(response);
    } catch (err: any) {
      setAiError(err.message || "Lỗi kết nối AI.");
    } finally {
      setIsAILoading(false);
      setActiveTask(null);
    }
  };

  const appendAIToDocument = () => {
    if (!aiResult) return;
    const formattedHtml = aiResult
      .replace(/### (.*)/g, "<b>$1</b>")
      .replace(/## (.*)/g, "<b>$1</b>")
      .replace(/\*\*(.*?)\*\*/g, "<b>$1</b>")
      .split(/\n\n/)
      .map((pText) => `<p class="ai-text-line">${pText.trim().replace(/\n/g, "<br>")}</p>`)
      .join("");

    const aiBlock = `<br><div class="ai-wrapper-box"><p class="ai-title-line">--- BẢN GHI TỪ TRỢ LÝ AI ---</p>${formattedHtml}</div>`;
    setDocHtml((prev) => prev + aiBlock);
    setIsAIModalOpen(false);
  };

  const handleExport = async () => {
    try {
      setIsExporting(true);
      await exportToWord(previewRef.current, fileName, config);
    } finally {
      setIsExporting(false);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 p-4 md:p-8 font-sans text-slate-800">
      <div className="max-w-7xl mx-auto">
        <header className="mb-8 border-b border-slate-300 pb-4">
          <h1 className="text-3xl font-bold text-blue-800 flex items-center gap-2">
            <FileText className="w-8 h-8" /> <span>Chuẩn Hóa Văn Bản & Soát Lỗi AI</span>
          </h1>
          <p className="text-slate-600 mt-2 text-sm">
            Công cụ hỗ trợ tải file Word, soát lỗi văn phong, căn chỉnh theo NĐ 30/2020/NĐ-CP và xuất lại file Word.
          </p>
        </header>

        <main className="grid grid-cols-1 lg:grid-cols-12 gap-8">
          <aside className="lg:col-span-4 space-y-5">
            <section className="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
              <h2 className="text-base font-semibold mb-3 flex items-center gap-2 text-blue-700">
                <span className="bg-blue-100 text-blue-800 w-6 h-6 rounded flex items-center justify-center text-sm">1</span> Nguồn Tài Liệu
              </h2>
              <label htmlFor="fileInput" className="cursor-pointer bg-slate-50 border-2 border-dashed border-slate-300 hover:border-blue-500 hover:bg-blue-50 transition-colors rounded-lg p-8 flex flex-col items-center">
                <Upload className="w-6 h-6 text-slate-500 mb-2" />
                <span className="text-sm text-slate-700">Tải lên file (.docx)</span>
              </label>
              <input type="file" id="fileInput" accept=".docx,.doc" className="hidden" onChange={handleFileUpload} />
              <div className="text-xs mt-3 text-center">
                {hasFile ? (
                  <span className="text-emerald-600 flex flex-col items-center gap-1">
                    <span>Đã tải:</span>
                    <span className="font-medium italic break-all flex items-center gap-1"><CheckCircle2 className="w-3 h-3" /> {fileStatus.replace("Đã tải: ", "")}</span>
                  </span>
                ) : (
                  <span className="text-slate-500">{fileStatus}</span>
                )}
              </div>
            </section>

            <section className={cn("bg-white p-5 rounded-xl shadow-sm border border-slate-200", !hasFile && "opacity-50 pointer-events-none")}>
              <h2 className="text-base font-semibold mb-3 flex items-center gap-2 text-blue-700">
                <SlidersHorizontal className="w-4 h-4" /> <span className="bg-blue-100 text-blue-800 w-6 h-6 rounded flex items-center justify-center text-sm">2</span> Chế Độ Giữ Khung/Bảng
              </h2>
              <p className="text-xs text-slate-500 mb-3">Kiểu xử lý sau khi upload</p>
              <div className="space-y-2">
                {[
                  { value: "preserve", title: "Giữ nguyên định dạng gốc tối đa", desc: "Ưu tiên giữ khung/bảng, phù hợp công văn, đơn, biên bản, biểu mẫu." },
                  { value: "nd30", title: "Chuẩn hóa theo NĐ 30", desc: "Ẩn khung kỹ thuật ở phần đầu khi xuất Word, căn lại tiêu đề văn bản." },
                  { value: "textOnly", title: "Chỉ lấy nội dung để AI soát lỗi", desc: "Không cố chuẩn hóa bố cục, dùng khi tài liệu quá phức tạp." },
                ].map((item) => (
                  <label key={item.value} className="flex gap-2 rounded-lg border border-slate-200 p-3 cursor-pointer hover:bg-slate-50">
                    <input type="radio" value={item.value} checked={processingMode === item.value} onChange={() => handleProcessingModeChange(item.value as DocumentProcessingMode)} />
                    <span>
                      <span className="block text-sm font-semibold text-slate-800">{item.title}</span>
                      <span className="block text-xs text-slate-500 mt-0.5">{item.desc}</span>
                    </span>
                  </label>
                ))}
                <label className="flex gap-2 rounded-lg border border-amber-200 bg-amber-50 p-3 cursor-pointer">
                  <input type="checkbox" checked={showPreviewFrame} onChange={(e) => handlePreviewFrameChange(e.target.checked)} />
                  <span>
                    <span className="block text-sm font-semibold text-slate-800">Hiển thị khung hỗ trợ ở bản xem trước</span>
                    <span className="block text-xs text-amber-700 mt-0.5">Khung phần đầu chỉ để kiểm tra; khi tải xuống Word sẽ tự ẩn nếu đó là bảng thể thức.</span>
                  </span>
                </label>
              </div>

              {hasFile && (
                <div className="mt-4 rounded-lg border border-slate-200 bg-slate-50 p-3 text-xs text-slate-700">
                  <p className="font-semibold flex items-center gap-1 mb-2"><Table2 className="w-3.5 h-3.5" /> Kiểm tra nhanh tài liệu</p>
                  <div className="grid grid-cols-2 gap-2">
                    <span>Số bảng/khung: <b>{stats.tableCount}</b></span>
                    <span>Số hình ảnh: <b>{stats.imageCount}</b></span>
                    <span>Thể thức hành chính: <b>{stats.administrativeHeaderDetected ? "Có" : "Không"}</b></span>
                    <span>Khối chữ ký/nơi nhận: <b>{stats.signatureTableDetected ? "Có" : "Không"}</b></span>
                  </div>
                  {stats.warningMessages.length > 0 && (
                    <div className="mt-2 space-y-1 text-amber-700">
                      {stats.warningMessages.map((msg) => <p key={msg} className="flex gap-1"><AlertTriangle className="w-3 h-3 mt-0.5 shrink-0" />{msg}</p>)}
                    </div>
                  )}
                </div>
              )}
            </section>

            <section className={cn("bg-white p-5 rounded-xl shadow-sm border border-slate-200", !hasFile && "opacity-50 pointer-events-none")}>
              <h2 className="text-base font-semibold mb-3 flex items-center gap-2 text-blue-700">
                <span className="bg-blue-100 text-blue-800 w-6 h-6 rounded flex items-center justify-center text-sm">3</span> Căn Lề NĐ 30
              </h2>
              <div className="space-y-3">
                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="block text-xs text-slate-600 mb-1">Font chữ</label>
                    <select className="w-full border p-1.5 rounded text-sm" value={config.font} onChange={(e) => setConfig({ ...config, font: e.target.value })}>
                      <option value="Times New Roman">Times New Roman</option>
                      <option value="Arial">Arial</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-xs text-slate-600 mb-1">Cỡ chữ</label>
                    <input type="number" className="w-full border p-1.5 rounded text-sm" value={config.size} onChange={(e) => setConfig({ ...config, size: parseFloat(e.target.value) })} />
                  </div>
                </div>
                <div className="grid grid-cols-2 gap-3 bg-emerald-50 p-3 rounded border border-emerald-100">
                  <div>
                    <label className="block text-xs text-slate-600 mb-1">Thụt dòng (cm)</label>
                    <input type="number" step="0.1" className="w-full border p-1.5 rounded text-sm" value={config.textIndent} onChange={(e) => setConfig({ ...config, textIndent: parseFloat(e.target.value) })} />
                  </div>
                  <div>
                    <label className="block text-xs text-slate-600 mb-1">Cách đoạn (pt)</label>
                    <input type="number" className="w-full border p-1.5 rounded text-sm" value={config.paraSpacing} onChange={(e) => setConfig({ ...config, paraSpacing: parseFloat(e.target.value) })} />
                  </div>
                </div>
                <div className="grid grid-cols-3 gap-3">
                  <div>
                    <label className="block text-xs text-slate-600 mb-1">Giãn dòng</label>
                    <select className="w-full border p-1.5 rounded text-sm" value={config.spacing} onChange={(e) => setConfig({ ...config, spacing: e.target.value })}>
                      <option value="1.0">1.0</option>
                      <option value="1.15">1.15</option>
                      <option value="1.5">1.5</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-xs text-slate-600 mb-1">Lề Trái</label>
                    <input type="number" step="0.1" className="w-full border p-1.5 rounded text-sm" value={config.leftMargin} onChange={(e) => setConfig({ ...config, leftMargin: parseFloat(e.target.value) })} />
                  </div>
                  <div>
                    <label className="block text-xs text-slate-600 mb-1">Lề Phải</label>
                    <input type="number" step="0.1" className="w-full border p-1.5 rounded text-sm" value={config.rightMargin} onChange={(e) => setConfig({ ...config, rightMargin: parseFloat(e.target.value) })} />
                  </div>
                </div>
              </div>
            </section>

            <section className={cn("bg-white p-5 rounded-xl shadow-sm border border-slate-200", !hasFile && "opacity-50 pointer-events-none")}>
              <h2 className="text-base font-semibold mb-3 flex items-center gap-2 text-blue-700">
                <Wand2 className="w-4 h-4" /> <span className="bg-blue-100 text-blue-800 w-6 h-6 rounded flex items-center justify-center text-sm">4</span> Soát Lỗi AI
              </h2>
              <div className="space-y-3">
                <div className="bg-slate-50 p-2 rounded border border-slate-100">
                  <label className="block text-xs text-slate-600 mb-2">Chế độ làm việc</label>
                  <div className="flex gap-4 text-sm">
                    <label className="flex items-center gap-2 cursor-pointer"><input type="radio" checked={aiMode === "suggest"} onChange={() => setAiMode("suggest")} /> Chỉ báo lỗi</label>
                    <label className="flex items-center gap-2 cursor-pointer"><input type="radio" checked={aiMode === "autofix"} onChange={() => setAiMode("autofix")} /> AI tự sửa</label>
                  </div>
                </div>
                <div>
                  <label className="block text-xs text-slate-600 mb-1">Mức độ & Phong cách</label>
                  <select className="w-full border p-2 rounded text-sm text-slate-700" value={aiStyle} onChange={(e) => setAiStyle(e.target.value)}>
                    <option value="grammar">1. Sửa Ngữ pháp & Nối dòng</option>
                    <option value="admin">2. Chuẩn hóa Hành chính (NĐ 30)</option>
                    <option value="edu">3. Chuẩn hóa Giáo dục (Sư phạm)</option>
                  </select>
                </div>
                <button onClick={() => handleAITask("proofread")} disabled={isAILoading} className="w-full bg-blue-600 hover:bg-blue-700 text-white font-medium py-2 rounded text-sm flex justify-center items-center gap-2 transition-colors">
                  {isAILoading && activeTask === "proofread" ? <Loader2 className="w-4 h-4 animate-spin" /> : null} Bắt đầu Soát lỗi & Sửa văn bản
                </button>
                <div className="grid grid-cols-2 gap-2">
                  <button onClick={() => handleAITask("summarize")} disabled={isAILoading} className="border bg-white hover:bg-slate-50 text-slate-700 py-1.5 rounded text-xs transition-colors">Tóm tắt văn bản</button>
                  <button onClick={() => handleAITask("quiz")} disabled={isAILoading} className="border bg-white hover:bg-slate-50 text-slate-700 py-1.5 rounded text-xs transition-colors">Tạo Bài tập</button>
                </div>
              </div>
            </section>

            <section className={cn("bg-white p-5 rounded-xl shadow-sm border border-slate-200", !hasFile && "opacity-50 pointer-events-none")}>
              <h2 className="text-base font-semibold mb-3 flex items-center gap-2 text-blue-700">
                <span className="bg-blue-100 text-blue-800 w-6 h-6 rounded flex items-center justify-center text-sm">5</span> Hoàn Tất
              </h2>
              <button onClick={handleExport} disabled={isExporting} className="w-full bg-blue-800 hover:bg-blue-900 text-white font-medium py-2.5 rounded flex justify-center items-center gap-2 transition-colors">
                {isExporting ? <Loader2 className="w-4 h-4 animate-spin" /> : <Save className="w-4 h-4" />} Tải Xuống File Word (.docx)
              </button>
            </section>
          </aside>

          <section className="lg:col-span-8 flex flex-col min-w-0">
            <div className="mb-4">
              <h2 className="text-lg font-bold text-slate-800">Bản Xem Trước (A4)</h2>
              <p className="text-xs text-slate-500 mt-1 flex items-center gap-1"><CheckCircle2 className="w-3 h-3" /> Bảng đầu văn bản được nhận diện theo nội dung, không xóa viền theo vị trí.</p>
            </div>

            <div className="bg-slate-300 p-4 md:p-8 rounded-xl flex justify-center overflow-auto flex-1 h-[800px]">
              <div className="bg-white shadow-xl relative a4-page" style={{ width: "794px", minHeight: "1123px", flexShrink: 0 }}>
                <style>{`
                  .document-content { padding: 2cm ${config.rightMargin}cm 2cm ${config.leftMargin}cm; background:#fff; color:#000; min-height:1123px; box-sizing:border-box; font-family:"${config.font}",serif; font-size:${config.size}pt; }
                  .document-content p:not(.ai-wrapper-box p) { font-family:"${config.font}",serif; font-size:${config.size}pt; line-height:${config.spacing}; text-align:${config.textAlign}; text-indent:${config.textIndent}cm; margin:0 0 ${config.paraSpacing}pt 0; }
                  .document-content table { width:100%; max-width:100%; table-layout:fixed; border-collapse:collapse; margin:6pt 0 10pt 0; box-sizing:border-box; }
                  .document-content td, .document-content th { border:1px solid #000; padding:4pt; vertical-align:top; overflow-wrap:break-word; }
                  .document-content table p, .document-content td p, .document-content th p { text-indent:0 !important; margin:0 0 2pt 0 !important; line-height:1.15 !important; }
                  .document-content .admin-header-table { width:106% !important; margin-left:-3% !important; margin-right:-3% !important; border:none !important; table-layout:fixed !important; margin-top:0 !important; margin-bottom:12pt !important; }
                  .document-content .admin-header-table col.admin-left-col { width:36% !important; }
                  .document-content .admin-header-table col.admin-right-col { width:64% !important; }
                  .document-content .admin-header-table td { border:none !important; padding:0 3pt !important; text-align:center !important; vertical-align:top !important; background:transparent !important; }
                  .document-content .admin-header-preview-frame, .document-content .admin-header-preview-frame td { border:1px dashed #9ca3af !important; }
                  .document-content .admin-header-table p { font-family:"Times New Roman",serif !important; text-align:center !important; text-indent:0 !important; margin:0 0 1pt 0 !important; line-height:1.12 !important; }
                  .document-content .admin-agency-line { font-size:13pt !important; font-weight:400 !important; text-transform:uppercase !important; }
                  .document-content .admin-unit-line { font-size:13pt !important; font-weight:700 !important; text-transform:uppercase !important; }
                  .document-content .admin-number-line { font-size:13pt !important; font-weight:400 !important; white-space:nowrap !important; }
                  .document-content .admin-national-line { font-size:11.2pt !important; font-weight:700 !important; white-space:nowrap !important; letter-spacing:-0.25pt !important; text-transform:uppercase !important; }
                  .document-content .admin-motto-line { font-size:12.5pt !important; font-weight:700 !important; white-space:nowrap !important; display:inline-block !important; border-bottom:1px solid #000 !important; padding-bottom:1pt !important; }
                  .document-content .admin-date-line { font-size:13pt !important; font-style:italic !important; white-space:nowrap !important; }
                  .document-content .signature-table, .document-content .signature-table td { border:none !important; }
                  .document-content .doc-main-title { text-align:center !important; text-indent:0 !important; font-weight:700 !important; text-transform:uppercase !important; font-size:${Math.max(config.size, 14)}pt !important; margin:8pt 0 4pt 0 !important; line-height:1.2 !important; }
                  .document-content .doc-sub-title { text-align:center !important; text-indent:0 !important; font-weight:700 !important; font-size:${config.size}pt !important; margin:0 0 3pt 0 !important; line-height:1.2 !important; }
                  .document-content img { max-width:100%; display:inline-block; }
                `}</style>
                {!hasFile ? (
                  <div className="h-full flex flex-col items-center justify-center text-slate-400 py-32 absolute inset-0"><FileQuestion className="w-16 h-16 mb-4 text-slate-200" /></div>
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
              <h3 className="font-bold text-blue-900">Kết Quả Phân Tích AI</h3>
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
