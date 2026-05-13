import React, { useState, useRef, ChangeEvent } from "react";
import mammoth from "mammoth";
import ReactMarkdown from "react-markdown";
import {
  FileText, Settings, Wand2, Save, Upload, X, FileQuestion, AlignLeft, Loader2, CheckCircle2, AlertTriangle
} from "lucide-react";

import { cn } from "./lib/utils";
import { runAI } from "./services/ai";
import { FormattingConfig, exportToWord, processHtmlToRemoveBullets } from "./utils/docx";

export default function App() {
  const [fileName, setFileName] = useState("Tai_Lieu");
  const [fileStatus, setFileStatus] = useState("Chưa có file nào");
  const [rawText, setRawText] = useState("");
  const [docHtml, setDocHtml] = useState("");
  const [hasFile, setHasFile] = useState(false);
  const [isLegacyDoc, setIsLegacyDoc] = useState(false);

  const [config, setConfig] = useState<FormattingConfig>({
    font: "Times New Roman", size: 14, spacing: "1.5", textAlign: "justify", textIndent: 1.27, paraSpacing: 6, leftMargin: 3, rightMargin: 2,
  });

  const [isAIModalOpen, setIsAIModalOpen] = useState(false);
  const [isAILoading, setIsAILoading] = useState(false);
  const [activeTask, setActiveTask] = useState<string | null>(null);
  const [aiResult, setAiResult] = useState("");
  const [aiError, setAiError] = useState("");
  const [aiMode, setAiMode] = useState("suggest");
  const [aiStyle, setAiStyle] = useState("edu");

  const previewRef = useRef<HTMLDivElement>(null);

  const handleFileUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    
    const extension = file.name.split('.').pop()?.toLowerCase();
    setFileName(file.name.replace(/\.(docx|doc)$/, ""));
    
    if (extension === 'doc') {
      setIsLegacyDoc(true);
      setFileStatus(`Cảnh báo: File .doc cũ có thể bị lỗi.`);
    } else {
      setIsLegacyDoc(false);
      setFileStatus(`Đang xử lý: ${file.name}...`);
    }

    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const arrayBuffer = e.target?.result as ArrayBuffer;
        const result = await mammoth.convertToHtml({ arrayBuffer }, {
          convertImage: mammoth.images.inline(async (element: any) => {
            const imageBuffer = await element.read("base64");
            return { src: `data:${element.contentType};base64,${imageBuffer}` };
          }),
        });
        const textResult = await mammoth.extractRawText({ arrayBuffer });
        setRawText(textResult.value);
        setDocHtml(processHtmlToRemoveBullets(result.value));
        setHasFile(true);
        if (extension !== 'doc') setFileStatus(`Đã tải: ${file.name}`);
      } catch (err) {
        alert("Lỗi: Không thể đọc file này. Nếu là file .doc, bạn hãy mở bằng Word rồi 'Save As' sang .docx nhé!");
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
      prompt = `Dựa TỐI ĐA vào văn bản dưới đây, tạo PHIẾU BÀI TẬP TRẮC NGHIỆM 10 câu (Ma trận 4-3-3: Nhận biết, Thông hiểu, Vận dụng).
      ⚠️ LƯU Ý: KHÔNG sử dụng ngân hàng câu hỏi cố định. BẮT BUỘC chỉ sử dụng dữ liệu người dùng nhập trong văn bản nguồn để sinh câu hỏi.
      
      Cấu trúc:
      # PHIẾU BÀI TẬP ÔN LUYỆN
      ## PHẦN 1: ĐỀ BÀI
      ### I. Mức 1 (4 câu)
      ### II. Mức 2 (3 câu)
      ### III. Mức 3 (3 câu)
      ## PHẦN 2: ĐÁP ÁN (In đậm đáp án)

      Văn bản nguồn:
      """\n${rawText}\n"""`;
    } else if (taskType === "proofread") {
      prompt = `Đóng vai chuyên gia ngôn ngữ, rà soát văn bản dưới đây.\n`;
      prompt += `⚠️ LƯU Ý TỐI QUAN TRỌNG: Nối lại các câu bị ngắt dòng sai logic (do gõ nhầm Enter giữa câu).\n`;
      prompt += `TUYỆT ĐỐI KHÔNG VIẾT LẠI TOÀN BỘ VĂN BẢN ĐỂ TIẾT KIỆM DUNG LƯỢNG.\n\n`;

      if (aiMode === "suggest") {
        prompt += `⚙️ CHẾ ĐỘ: CHỈ BÁO LỖI. Lập danh sách: "Đoạn lỗi -> Cách sửa -> Lý do".\n\n`;
      } else {
        prompt += `⚙️ CHẾ ĐỘ: TỰ ĐỘNG SỬA. CHỈ in ra những ĐOẠN CẦN SỬA. Bỏ qua các đoạn đúng.\n\n`;
      }

      prompt += `🎨 PHONG CÁCH: `;
      switch (aiStyle) {
        case "grammar": prompt += `Sửa toàn diện ngữ pháp và nối dòng.\n\n`; break;
        case "admin": prompt += `Chuẩn hóa hành chính (NĐ 30).\n\n`; break;
        case "edu": prompt += `Chuẩn hóa văn phong sư phạm.\n\n`; break;
      }
      prompt += `📄 VĂN BẢN:\n"""\n${rawText}\n"""`;
    }

    try {
      const response = await runAI(prompt);
      setAiResult(response);
    } catch (err: any) {
      setAiError(err.message);
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

  return (
    <div className="min-h-screen bg-slate-100 p-4 md:p-8 font-sans text-slate-800">
      <div className="max-w-7xl mx-auto">
        <header className="mb-8 border-b border-slate-300 pb-4">
          <h1 className="text-3xl font-bold text-blue-800 flex items-center gap-2">
            <FileText className="w-8 h-8" /> <span>Chuẩn Hóa Văn Bản & Soát Lỗi AI</span>
          </h1>
          <p className="text-slate-600 mt-2 text-sm">
            Công cụ hỗ trợ tải file Word, tự động soát lỗi văn phong, căn chỉnh chuẩn xác theo NĐ 30/2020/NĐ-CP và xuất lại file.
          </p>
        </header>

        <main className="grid grid-cols-1 lg:grid-cols-12 gap-8">
          <aside className="lg:col-span-4 space-y-6">
            {/* 1. File Source */}
            <section className="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
              <h2 className="text-lg font-semibold mb-4 flex items-center gap-2">
                <span className="text-xl">1️⃣</span> Nguồn Tài Liệu
              </h2>
              <label htmlFor="fileInput" className="cursor-pointer bg-slate-50 border-2 border-dashed border-slate-400 hover:border-blue-700 hover:bg-blue-50 transition-colors rounded-lg p-6 flex flex-col items-center">
                <Upload className="w-8 h-8 text-slate-500 mb-2" />
                <span className="font-medium text-slate-700">Tải lên file (.docx)</span>
              </label>
              <input type="file" id="fileInput" accept=".docx,.doc" className="hidden" onChange={handleFileUpload} />
              <div className="text-sm mt-3 text-center italic">
                {isLegacyDoc ? (
                  <span className="text-amber-600 font-medium flex items-center justify-center gap-1"><AlertTriangle className="w-4 h-4" /> Khuyên dùng file .docx</span>
                ) : hasFile ? (
                  <span className="text-emerald-600 font-medium flex items-center justify-center gap-1"><CheckCircle2 className="w-4 h-4" /> {fileStatus}</span>
                ) : (
                  <span className="text-slate-500">{fileStatus}</span>
                )}
              </div>
            </section>

            {/* 2. Formatting Config */}
            <section className={cn("bg-white p-5 rounded-xl shadow-sm border border-slate-200", !hasFile && "opacity-50 pointer-events-none")}>
              <h2 className="text-lg font-semibold mb-4 flex items-center gap-2">
                <Settings className="w-5 h-5 text-slate-700" /> <span>2️⃣</span> Căn Lề NĐ 30
              </h2>
              <div className="space-y-4">
                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <label className="block text-sm font-medium text-slate-700 mb-1">Font chữ</label>
                    <select className="w-full border p-2 rounded text-sm" value={config.font} onChange={(e) => setConfig({ ...config, font: e.target.value })}>
                      <option value="Times New Roman">Times New Roman</option><option value="Arial">Arial</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-slate-700 mb-1">Cỡ chữ</label>
                    <input type="number" className="w-full border p-2 rounded text-sm" value={config.size} onChange={(e) => setConfig({ ...config, size: parseFloat(e.target.value) })} />
                  </div>
                </div>
                <div className="p-3 bg-emerald-50 border border-emerald-200 rounded-md grid grid-cols-2 gap-3">
                  <div>
                    <label className="block text-sm font-medium text-slate-700 mb-1">Thụt dòng (cm)</label>
                    <input type="number" step="0.1" className="w-full border p-2 rounded text-sm" value={config.textIndent} onChange={(e) => setConfig({ ...config, textIndent: parseFloat(e.target.value) })} />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-slate-700 mb-1">Cách đoạn (pt)</label>
                    <input type="number" className="w-full border p-2 rounded text-sm" value={config.paraSpacing} onChange={(e) => setConfig({ ...config, paraSpacing: parseFloat(e.target.value) })} />
                  </div>
                </div>
                <div className="grid grid-cols-3 gap-3">
                  <div><label className="block text-sm font-medium text-slate-700 mb-1">Giãn dòng</label><select className="w-full border p-2 rounded text-sm" value={config.spacing} onChange={(e) => setConfig({ ...config, spacing: e.target.value })}><option value="1.0">1.0</option><option value="1.5">1.5</option></select></div>
                  <div><label className="block text-sm font-medium text-slate-700 mb-1">Lề Trái</label><input type="number" step="0.5" className="w-full border p-2 rounded text-sm" value={config.leftMargin} onChange={(e) => setConfig({ ...config, leftMargin: parseFloat(e.target.value) })} /></div>
                  <div><label className="block text-sm font-medium text-slate-700 mb-1">Lề Phải</label><input type="number" step="0.5" className="w-full border p-2 rounded text-sm" value={config.rightMargin} onChange={(e) => setConfig({ ...config, rightMargin: parseFloat(e.target.value) })} /></div>
                </div>
              </div>
            </section>

            {/* 3. AI Assistant */}
            <section className={cn("bg-indigo-50 p-5 rounded-xl shadow-sm border border-indigo-200", !hasFile && "opacity-50 pointer-events-none")}>
              <h2 className="text-lg font-semibold mb-3 text-indigo-800 flex items-center gap-2"><Wand2 className="w-5 h-5" /> <span>3️⃣ Soát Lỗi AI</span></h2>
              <div className="space-y-3">
                <div className="bg-white p-3 rounded-md border border-indigo-100">
                  <select className="w-full border border-indigo-200 p-2 rounded text-sm" value={aiStyle} onChange={(e) => setAiStyle(e.target.value)}>
                    <option value="grammar">Sửa Ngữ pháp & Nối dòng sai</option>
                    <option value="admin">Chuẩn hóa Hành chính (NĐ 30)</option>
                    <option value="edu">Chuẩn hóa Giáo dục (Sư phạm)</option>
                  </select>
                </div>
                <div className="grid grid-cols-2 gap-2">
                  <button onClick={() => handleAITask("proofread")} disabled={isAILoading} className="col-span-2 bg-indigo-600 hover:bg-indigo-700 text-white font-medium px-4 py-2 rounded-md text-sm flex justify-center items-center gap-2">
                    {isAILoading && activeTask === "proofread" ? <Loader2 className="w-4 h-4 animate-spin" /> : <Wand2 className="w-4 h-4" />} Soát lỗi & Sửa văn bản
                  </button>
                  <button onClick={() => handleAITask("summarize")} disabled={isAILoading} className="border bg-white text-indigo-700 px-3 py-2 rounded-md text-xs">Tóm tắt văn bản</button>
                  <button onClick={() => handleAITask("quiz")} disabled={isAILoading} className="border bg-white text-indigo-700 px-3 py-2 rounded-md text-xs">Tạo Bài tập (10 câu)</button>
                </div>
              </div>
            </section>

            {/* 4. Export */}
            <section className={cn("bg-white p-5 rounded-xl shadow-sm border border-l-4 border-l-blue-800", !hasFile && "opacity-50 pointer-events-none")}>
              <h2 className="text-lg font-semibold mb-3"><span>💾</span> 4️⃣ Hoàn Tất</h2>
              <button onClick={() => exportToWord(previewRef.current, fileName, config)} className="w-full bg-blue-800 hover:bg-blue-900 text-white font-medium px-4 py-3 rounded-md flex justify-center items-center gap-2">
                <Save className="w-5 h-5" /> Tải Xuống File Word (.doc)
              </button>
            </section>
          </aside>

          {/* RIGHT COLUMN: PREVIEW */}
          <section className="lg:col-span-8 flex flex-col">
            <h2 className="text-xl font-bold mb-4">Bản Xem Trước (A4)</h2>
            <div className="bg-slate-300 p-4 rounded-xl flex justify-center overflow-auto h-[800px]">
              <div className="bg-white shadow-xl w-full max-w-[800px] relative" style={{ minHeight: "842px" }}>
                <style>{`
                  .document-content { padding: 2cm ${config.rightMargin}cm 2cm ${config.leftMargin}cm; background-color: white; color: black; text-align: justify; }
                  
                  /* Đoạn văn bình thường */
                  .document-content p:not(.ai-wrapper-box p) { 
                    font-family: "${config.font}", serif; font-size: ${config.size}pt; 
                    line-height: ${config.spacing}; text-align: ${config.textAlign}; 
                    text-indent: ${config.textIndent}cm; margin-top: 0pt; margin-bottom: ${config.paraSpacing}pt; 
                  }
                  
                  /* Bỏ thụt lề cho tiêu đề in đậm */
                  .document-content p:has(> strong:only-child), .document-content p:has(> b:only-child) { 
                    text-indent: 0cm !important; 
                  }

                  /* ========================================= */
                  /* SỬA LỖI BẢNG BỊ ĐÈ CHỮ (Overlap) */
                  /* ========================================= */
                  .document-content table { 
                    width: 100%; max-width: 100%;
                    table-layout: fixed; /* LỆNH KHÓA KHUNG QUAN TRỌNG NHẤT */
                    border-collapse: collapse; margin-top: 6pt; margin-bottom: 12pt; 
                  }
                  .document-content td, .document-content th { 
                    border: 1px solid #000; padding: 0.5rem; 
                    text-align: left !important; vertical-align: top; word-wrap: break-word;
                  }
                  .document-content table p, .document-content td p, .document-content th p {
                    text-indent: 0cm !important; margin-top: 0pt !important; margin-bottom: 4pt !important;
                    line-height: 1.2 !important;
                  }

                  /* 1. BẢNG ĐẦU TIÊN (Quốc hiệu - Tiêu ngữ) */
                  .document-content table:first-of-type, .document-content table:first-of-type td {
                    border: none !important; padding: 0 !important;
                  }
                  .document-content table:first-of-type td {
                    width: 50%; /* CHIA ĐÔI 50/50 ĐỂ KHÔNG BAO GIỜ BỊ ĐÈ */
                  }
                  .document-content table:first-of-type td p {
                    text-align: center !important; margin-bottom: 2pt !important; 
                  }
                  .document-content table:first-of-type tr:first-child td p { font-weight: bold; }

                  /* 2. BẢNG CUỐI CÙNG (Chữ ký - Nơi nhận) */
                  .document-content table:last-of-type, .document-content table:last-of-type td {
                    border: none !important; padding: 0 !important;
                  }
                  .document-content table:last-of-type td {
                    width: 50%; /* CHIA ĐÔI 50/50 */
                  }
                  .document-content table:last-of-type td:first-child p { text-align: left !important; }
                  .document-content table:last-of-type td:first-child p:first-child { font-weight: bold; font-style: italic; }
                  .document-content table:last-of-type td:last-child p { text-align: center !important; }
                  .document-content table:last-of-type td:last-child p:first-child { font-weight: bold; }

                  .document-content img { max-width: 100%; display: inline-block; }
                  .ai-wrapper-box { border-top: 1px dashed black; margin-top: 2rem; padding-top: 1rem; }
                  .ai-wrapper-box .ai-title-line { text-align: center; font-weight: bold; margin-bottom: 1rem; }
                  .ai-wrapper-box .ai-text-line { text-align: left; margin-bottom: 0.5rem; }
                `}</style>

                {!hasFile ? (
                  <div className="h-full flex flex-col items-center justify-center text-slate-400 py-32">
                    <FileQuestion className="w-16 h-16 mb-4" /> <p>Tài liệu sẽ hiển thị ở đây</p>
                  </div>
                ) : (
                  <div ref={previewRef} className="document-content h-full w-full outline-none" dangerouslySetInnerHTML={{ __html: docHtml }} />
                )}
              </div>
            </div>
          </section>
        </main>

        <footer className="mt-10 pt-6 border-t border-slate-300 text-center pb-8">
          <p className="text-xl font-bold text-blue-800 uppercase tracking-wide">Trường PTDTBT Tiểu học Giàng Chu Phìn</p>
          <p className="text-sm text-slate-500 italic mt-2">(Công cụ hỗ trợ Chuẩn hóa tài liệu Giáo dục)</p>
        </footer>
      </div>

      {/* AI MODAL */}
      {isAIModalOpen && (
        <div className="fixed inset-0 bg-slate-900/50 flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-4xl flex flex-col max-h-[90vh]">
            <div className="p-4 border-b flex justify-between items-center bg-indigo-50">
              <h3 className="text-lg font-bold text-indigo-900 flex items-center gap-2"><AlignLeft className="w-5 h-5" /> Kết Quả Phân Tích AI</h3>
              <button onClick={() => setIsAIModalOpen(false)} className="text-slate-400 hover:text-slate-700"><X /></button>
            </div>
            <div className="p-6 overflow-y-auto flex-1 bg-slate-50">
              {isAILoading ? (
                <div className="flex flex-col items-center justify-center py-20"><Loader2 className="animate-spin text-indigo-600 w-12 h-12 mb-4" /><p>AI đang rà soát và xử lý dữ liệu...</p></div>
              ) : aiError ? (
                <div className="text-red-600 bg-red-50 p-4 rounded border border-red-200">❌ {aiError}</div>
              ) : (
                <div className="prose prose-indigo max-w-none bg-white p-6 rounded-lg border shadow-sm"><ReactMarkdown>{aiResult}</ReactMarkdown></div>
              )}
            </div>
            <div className="p-4 border-t flex justify-end gap-3 bg-white">
              <button onClick={() => setIsAIModalOpen(false)} className="px-5 py-2 border rounded-md">Đóng</button>
              {!isAILoading && !aiError && aiResult && <button onClick={appendAIToDocument} className="px-5 py-2 bg-indigo-600 text-white rounded-md">Ghi kết quả này vào cuối văn bản</button>}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
