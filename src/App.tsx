import React, { useState, useRef, ChangeEvent } from "react";
import mammoth from "mammoth";
import ReactMarkdown from "react-markdown";
import {
  FileText,
  Settings,
  Wand2,
  Save,
  Upload,
  X,
  FileQuestion,
  AlignLeft,
  Loader2,
  CheckCircle2,
} from "lucide-react";

import { cn } from "./lib/utils";
import { runAI } from "./services/ai";
import {
  FormattingConfig,
  exportToWord,
  processHtmlToRemoveBullets,
} from "./utils/docx";

export default function App() {
  // Document State
  const [fileName, setFileName] = useState("Tai_Lieu");
  const [fileStatus, setFileStatus] = useState("Chưa có file nào");
  const [rawText, setRawText] = useState("");
  const [docHtml, setDocHtml] = useState("");
  const [hasFile, setHasFile] = useState(false);

  // Formatting State
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

  // AI Modal State
  const [isAIModalOpen, setIsAIModalOpen] = useState(false);
  const [isAILoading, setIsAILoading] = useState(false);
  const [activeTask, setActiveTask] = useState<string | null>(null);
  const [aiResult, setAiResult] = useState("");
  const [aiError, setAiError] = useState("");

  const [aiMode, setAiMode] = useState("suggest");
  const [aiStyle, setAiStyle] = useState("edu");

  const previewRef = useRef<HTMLDivElement>(null);

  // Handle File Upload
  const handleFileUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setFileName(file.name.replace(".docx", ""));
    setFileStatus(`Đang xử lý: ${file.name}...`);

    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const arrayBuffer = e.target?.result as ArrayBuffer;

        // Convert to HTML
        const result = await mammoth.convertToHtml(
          { arrayBuffer },
          {
            convertImage: mammoth.images.inline(async (element: any) => {
              const imageBuffer = await element.read("base64");
              return { src: `data:${element.contentType};base64,${imageBuffer}` };
            }),
          }
        );

        // Extract Text
        const textResult = await mammoth.extractRawText({ arrayBuffer });
        setRawText(textResult.value);

        // Process and set HTML
        const processedHtml = processHtmlToRemoveBullets(result.value);
        setDocHtml(processedHtml);
        setHasFile(true);
        setFileStatus(`Đã tải: ${file.name}`);
      } catch (err) {
        alert("Lỗi đọc file. Hãy chắc chắn đây là file .docx chuẩn.");
        console.error(err);
        setFileStatus("Lỗi đọc file");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Handle AI Tasks
  const handleAITask = async (taskType: "summarize" | "quiz" | "proofread") => {
    if (!rawText) {
      alert("Vui lòng tải lên tài liệu trước!");
      return;
    }

    setActiveTask(taskType);
    setIsAIModalOpen(true);
    setIsAILoading(true);
    setAiResult("");
    setAiError("");

    let prompt = "";

    if (taskType === "summarize") {
      prompt =
        "Tóm tắt ngắn gọn, gạch đầu dòng các ý chính yếu của văn bản sau:\n\n" +
        rawText;
    } else if (taskType === "quiz") {
      prompt = 
        `Dựa vào văn bản dưới, tạo PHIẾU BÀI TẬP TRẮC NGHIỆM 10 câu (Ma trận 4-3-3: Nhận biết, Thông hiểu, Vận dụng).
        Cấu trúc:
        # PHIẾU BÀI TẬP ÔN LUYỆN
        ## PHẦN 1: ĐỀ BÀI
        ### I. Mức 1 (4 câu)
        ### II. Mức 2 (3 câu)
        ### III. Mức 3 (3 câu)
        ## PHẦN 2: ĐÁP ÁN (In đậm đáp án, kèm giải thích ngắn)

        Văn bản nguồn:
        """\n${rawText}\n"""`;
    } else if (taskType === "proofread") {
      prompt = `Đóng vai chuyên gia ngôn ngữ, rà soát văn bản dưới đây.\n`;
      prompt += `⚠️ LƯU Ý TỐI QUAN TRỌNG ĐỂ TIẾT KIỆM DUNG LƯỢNG: TUYỆT ĐỐI KHÔNG VIẾT LẠI TOÀN BỘ VĂN BẢN.\n`;
      prompt += `CHỈ trích xuất các câu/từ bị lỗi và đưa ra cách sửa ngắn gọn.\n\n`;

      if (aiMode === "suggest") {
        prompt += `⚙️ CHẾ ĐỘ: CHỈ BÁO LỖI. Lập danh sách ngắn gọn: "Đoạn/Câu chứa lỗi -> Cách sửa -> Lý do".\n\n`;
      } else {
        prompt += `⚙️ CHẾ ĐỘ: TỰ ĐỘNG SỬA. CHỈ in ra những ĐOẠN/CÂU CẦN SỬA sau khi đã chỉnh chu. Bỏ qua toàn bộ các đoạn đã đúng để tiết kiệm chữ.\n\n`;
      }

      prompt += `🎨 PHONG CÁCH: `;
      switch (aiStyle) {
        case "spelling":
          prompt += `Chỉ sửa lỗi chính tả, đánh máy.\n\n`;
          break;
        case "grammar":
          prompt += `Sửa chính tả & cấu trúc ngữ pháp lủng củng. Nối các câu bị ngắt dòng sai logic do gõ nhầm phím Enter.\n\n`;
          break;
        case "admin":
          prompt += `Chuẩn hóa hành chính (NĐ 30), trang trọng, súc tích.\n\n`;
          break;
        case "edu":
          prompt += `Chuẩn hóa văn phong sư phạm, chuẩn mực, truyền cảm.\n\n`;
          break;
        case "academic":
          prompt += `Văn phong học thuật, lập luận chặt chẽ.\n\n`;
          break;
        case "original":
          prompt += `Sửa lỗi cơ bản, BẮT BUỘC giữ nguyên văn phong gốc.\n\n`;
          break;
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

    let text = aiResult
      .replace(/### (.*)/g, "<b>$1</b>")
      .replace(/#### (.*)/g, "<b><i>$1</i></b>")
      .replace(/## (.*)/g, "<b>$1</b>")
      .replace(/# (.*)/g, "<b>$1</b>")
      .replace(/\*\*(.*?)\*\*/g, "<b>$1</b>")
      .replace(/\*(.*?)\*/g, "<i>$1</i>");

    let paragraphs = text.split(/\n\n/);

    let formattedHtml = paragraphs
      .map((pText) => {
        let cleanText = pText.trim().replace(/\n/g, "<br>");
        if (!cleanText) return "";
        return `<p class="ai-text-line">${cleanText}</p>`;
      })
      .join("");

    const aiBlock = `
        <br>
        <div class="ai-wrapper-box">
            <p class="ai-title-line">--- BẢN GHI TỪ TRỢ LÝ AI ---</p>
            ${formattedHtml}
        </div>
    `;

    setDocHtml((prev) => prev + aiBlock);
    setIsAIModalOpen(false);
  };

  return (
    <div className="min-h-screen bg-slate-100 text-slate-800 antialiased p-4 md:p-8 font-sans">
      <div className="max-w-7xl mx-auto">
        <header className="mb-8 border-b border-slate-300 pb-4">
          <h1 className="text-3xl font-bold text-blue-800 flex items-center gap-2">
            <FileText className="w-8 h-8" /> <span>Chuẩn Hóa Văn Bản & Soát Lỗi AI</span>
          </h1>
          <p className="text-slate-600 mt-2 text-sm md:text-base">
            Công cụ hỗ trợ tải file Word, tự động soát lỗi văn phong, căn chỉnh
            chuẩn xác theo NĐ 30/2020/NĐ-CP và xuất lại file.
          </p>
        </header>

        <main className="grid grid-cols-1 lg:grid-cols-12 gap-8">
          {/* LEFT COLUMN: CONFIGURATION */}
          <aside className="lg:col-span-4 space-y-6">
            {/* 1. File Source */}
            <section className="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
              <h2 className="text-lg font-semibold mb-4 text-slate-800 flex items-center gap-2">
                <span className="text-xl">1️⃣</span> Nguồn Tài Liệu
              </h2>
              <div className="flex flex-col gap-3">
                <label
                  htmlFor="fileInput"
                  className="cursor-pointer bg-slate-50 border-2 border-dashed border-slate-400 hover:border-blue-700 hover:bg-blue-50 transition-colors rounded-lg p-6 text-center flex flex-col items-center justify-center"
                >
                  <Upload className="w-8 h-8 text-slate-500 mb-2" />
                  <span className="font-medium text-slate-700">
                    Tải lên file (.docx)
                  </span>
                </label>
                <input
                  type="file"
                  id="fileInput"
                  accept=".docx"
                  className="hidden"
                  onChange={handleFileUpload}
                />
                <div className="text-sm text-center italic flex items-center justify-center gap-1">
                  {hasFile ? (
                    <span className="text-emerald-600 font-medium flex items-center gap-1">
                      <CheckCircle2 className="w-4 h-4" /> {fileStatus}
                    </span>
                  ) : (
                    <span className="text-slate-500">{fileStatus}</span>
                  )}
                </div>
              </div>
            </section>

            {/* 2. Formatting Config */}
            <section
              className={cn(
                "bg-white p-5 rounded-xl shadow-sm border border-slate-200 transition-opacity",
                !hasFile && "opacity-50 pointer-events-none"
              )}
            >
              <h2 className="text-lg font-semibold mb-4 text-slate-800 flex items-center gap-2">
                <Settings className="w-5 h-5 text-slate-700" />
                <span>2️⃣</span> Căn Lề NĐ 30
              </h2>

              <div className="space-y-4">
                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <label className="block text-sm font-medium text-slate-700 mb-1">
                      Font chữ
                    </label>
                    <select
                      className="w-full border p-2 rounded focus:ring-blue-800 text-sm"
                      value={config.font}
                      onChange={(e) =>
                        setConfig({ ...config, font: e.target.value })
                      }
                    >
                      <option value="Times New Roman">Times New Roman</option>
                      <option value="Arial">Arial</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-slate-700 mb-1">
                      Cỡ chữ
                    </label>
                    <input
                      type="number"
                      className="w-full border p-2 rounded text-sm"
                      value={config.size}
                      onChange={(e) =>
                        setConfig({
                          ...config,
                          size: parseFloat(e.target.value),
                        })
                      }
                    />
                  </div>
                </div>

                <div className="p-3 bg-emerald-50 border border-emerald-200 rounded-md">
                  <h3 className="text-xs font-bold text-emerald-800 uppercase mb-3">
                    Thông số đoạn văn
                  </h3>
                  <div className="space-y-3">
                    <div>
                      <label className="block text-sm font-medium text-slate-700 mb-1">
                        Căn lề
                      </label>
                      <select
                        className="w-full border p-2 rounded text-sm"
                        value={config.textAlign}
                        onChange={(e) =>
                          setConfig({ ...config, textAlign: e.target.value })
                        }
                      >
                        <option value="justify">Căn đều hai bên</option>
                        <option value="left">Căn trái</option>
                      </select>
                    </div>
                    <div className="grid grid-cols-2 gap-3">
                      <div>
                        <label className="block text-sm font-medium text-slate-700 mb-1">
                          Thụt dòng (cm)
                        </label>
                        <input
                          type="number"
                          step="0.1"
                          className="w-full border p-2 rounded text-sm"
                          value={config.textIndent}
                          onChange={(e) =>
                            setConfig({
                              ...config,
                              textIndent: parseFloat(e.target.value),
                            })
                          }
                        />
                      </div>
                      <div>
                        <label className="block text-sm font-medium text-slate-700 mb-1">
                          Cách đoạn (pt)
                        </label>
                        <input
                          type="number"
                          className="w-full border p-2 rounded text-sm"
                          value={config.paraSpacing}
                          onChange={(e) =>
                            setConfig({
                              ...config,
                              paraSpacing: parseFloat(e.target.value),
                            })
                          }
                        />
                      </div>
                    </div>
                  </div>
                </div>

                <div className="grid grid-cols-3 gap-3">
                  <div>
                    <label className="block text-sm font-medium text-slate-700 mb-1">
                      Giãn dòng
                    </label>
                    <select
                      className="w-full border p-2 rounded text-sm"
                      value={config.spacing}
                      onChange={(e) =>
                        setConfig({ ...config, spacing: e.target.value })
                      }
                    >
                      <option value="1.0">1.0</option>
                      <option value="1.15">1.15</option>
                      <option value="1.5">1.5</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-slate-700 mb-1">
                      Lề Trái
                    </label>
                    <input
                      type="number"
                      step="0.5"
                      className="w-full border p-2 rounded text-sm"
                      value={config.leftMargin}
                      onChange={(e) =>
                        setConfig({
                          ...config,
                          leftMargin: parseFloat(e.target.value),
                        })
                      }
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-slate-700 mb-1">
                      Lề Phải
                    </label>
                    <input
                      type="number"
                      step="0.5"
                      className="w-full border p-2 rounded text-sm"
                      value={config.rightMargin}
                      onChange={(e) =>
                        setConfig({
                          ...config,
                          rightMargin: parseFloat(e.target.value),
                        })
                      }
                    />
                  </div>
                </div>
              </div>
            </section>

            {/* 3. AI Assistant */}
            <section
              className={cn(
                "bg-indigo-50 p-5 rounded-xl shadow-sm border border-indigo-200 transition-opacity",
                !hasFile && "opacity-50 pointer-events-none"
              )}
            >
              <h2 className="text-lg font-semibold mb-3 text-indigo-800 flex items-center gap-2">
                <Wand2 className="w-5 h-5" /> <span>3️⃣ Soát Lỗi & Biên Tập AI</span>
              </h2>

              <div className="space-y-4">
                <div className="bg-white p-3 rounded-md border border-indigo-100">
                  <label className="block text-xs font-bold text-indigo-800 uppercase mb-2">
                    Chế độ làm việc
                  </label>
                  <div className="flex flex-col gap-2 text-sm">
                    <label className="flex items-center gap-2 cursor-pointer">
                      <input
                        type="radio"
                        value="suggest"
                        checked={aiMode === "suggest"}
                        onChange={() => setAiMode("suggest")}
                        className="text-indigo-600 focus:ring-indigo-500"
                      />
                      <span className="text-slate-700">
                        🔍 <b className="text-indigo-700">Chỉ báo lỗi</b> (Chỉ ra
                        lỗi sai để tự sửa)
                      </span>
                    </label>
                    <label className="flex items-center gap-2 cursor-pointer">
                      <input
                        type="radio"
                        value="autofix"
                        checked={aiMode === "autofix"}
                        onChange={() => setAiMode("autofix")}
                        className="text-indigo-600 focus:ring-indigo-500"
                      />
                      <span className="text-slate-700">
                        ✨ <b className="text-emerald-700">AI Tự sửa</b> (Tự động
                        viết lại đoạn lỗi)
                      </span>
                    </label>
                  </div>
                </div>

                <div className="bg-white p-3 rounded-md border border-indigo-100">
                  <label className="block text-xs font-bold text-indigo-800 uppercase mb-2">
                    Mức độ & Phong cách
                  </label>
                  <select
                    className="w-full border border-indigo-200 p-2 rounded text-sm focus:ring-indigo-500 text-slate-700"
                    value={aiStyle}
                    onChange={(e) => setAiStyle(e.target.value)}
                  >
                    <option value="spelling">
                      1. Chỉ sửa chính tả, dấu câu, khoảng trắng
                    </option>
                    <option value="grammar">
                      2. Sửa toàn diện Ngữ pháp & Câu cú
                    </option>
                    <option value="admin">
                      3. Chuẩn hóa Hành chính (Theo NĐ 30)
                    </option>
                    <option value="edu">
                      4. Chuẩn hóa Giáo dục (Văn phong Sư phạm)
                    </option>
                    <option value="academic">
                      5. Viết Học thuật (Khoa học, chuyên sâu)
                    </option>
                    <option value="original">
                      6. Sửa lỗi nhưng GIỮ NGUYÊN văn phong gốc
                    </option>
                  </select>
                </div>

                <div className="grid grid-cols-2 gap-2">
                  <button
                    onClick={() => handleAITask("proofread")}
                    disabled={isAILoading}
                    className={cn(
                      "col-span-2 text-white font-medium px-4 py-2 rounded-md text-sm transition-colors shadow-sm flex justify-center items-center gap-2",
                      isAILoading
                        ? "bg-indigo-400 cursor-not-allowed"
                        : "bg-indigo-600 hover:bg-indigo-700"
                    )}
                  >
                    {isAILoading && activeTask === "proofread" ? (
                      <>
                        <Loader2 className="w-4 h-4 animate-spin" />
                        Đang xử lý...
                      </>
                    ) : (
                      <>
                        <Wand2 className="w-4 h-4" />
                        Bắt đầu Soát lỗi & Sửa văn bản
                      </>
                    )}
                  </button>
                  <button
                    onClick={() => handleAITask("summarize")}
                    disabled={isAILoading}
                    className={cn(
                      "border font-medium px-3 py-2 rounded-md text-xs transition-colors truncate flex justify-center items-center gap-1",
                      isAILoading
                        ? "bg-indigo-50 border-indigo-200 text-indigo-400 cursor-not-allowed"
                        : "bg-white border-indigo-200 hover:bg-indigo-50 text-indigo-700"
                    )}
                  >
                    {isAILoading && activeTask === "summarize" ? (
                      <>
                        <Loader2 className="w-3 h-3 animate-spin" />
                        Đang xử lý...
                      </>
                    ) : (
                      "Tóm tắt văn bản"
                    )}
                  </button>
                  <button
                    onClick={() => handleAITask("quiz")}
                    disabled={isAILoading}
                    className={cn(
                      "border font-medium px-3 py-2 rounded-md text-xs transition-colors truncate flex justify-center items-center gap-1",
                      isAILoading
                        ? "bg-indigo-50 border-indigo-200 text-indigo-400 cursor-not-allowed"
                        : "bg-white border-indigo-200 hover:bg-indigo-50 text-indigo-700"
                    )}
                  >
                    {isAILoading && activeTask === "quiz" ? (
                      <>
                        <Loader2 className="w-3 h-3 animate-spin" />
                        Đang xử lý...
                      </>
                    ) : (
                      "Tạo Phiếu Bài Tập (10 câu)"
                    )}
                  </button>
                </div>
              </div>
            </section>

            {/* 4. Export */}
            <section
              className={cn(
                "bg-white p-5 rounded-xl shadow-sm border border-slate-200 border-l-4 border-l-blue-800 transition-opacity",
                !hasFile && "opacity-50 pointer-events-none"
              )}
            >
              <h2 className="text-lg font-semibold mb-3 flex items-center gap-2">
                <span className="text-xl">💾</span> 4️⃣ Hoàn Tất
              </h2>
              <button
                onClick={() => exportToWord(previewRef.current, fileName, config)}
                className="w-full bg-blue-800 hover:bg-blue-900 text-white font-medium px-4 py-3 rounded-md transition-colors shadow-md flex justify-center items-center gap-2"
              >
                <Save className="w-5 h-5" />
                Tải Xuống File Word (.doc)
              </button>
              <p className="text-xs text-slate-500 mt-2 text-center">
                Tự động chống đè chữ và ép đúng lề.
              </p>
            </section>
          </aside>

          {/* RIGHT COLUMN: PREVIEW */}
          <section className="lg:col-span-8 flex flex-col">
            <div className="flex justify-between items-end mb-4">
              <h2 className="text-xl font-bold text-slate-800">
                Bản Xem Trước (A4)
              </h2>
            </div>

            <div className="bg-slate-300 p-4 md:p-8 rounded-xl flex justify-center overflow-auto flex-1 h-[800px]">
              <div
                className="bg-white shadow-xl w-full max-w-[800px] mx-auto overflow-hidden relative"
                style={{ minHeight: "842px" }}
              >
                {/* Dynamic Configuration Styles */}
                <style>{`
                  .document-content {
                    padding: 2cm ${config.rightMargin}cm 2cm ${config.leftMargin}cm;
                    background-color: white;
                    color: black;
                    text-align: justify;
                  }
                  
                  /* 1. ĐỊNH DẠNG ĐOẠN VĂN CHUẨN (CÓ THỤT LỀ) */
                  .document-content p:not(.ai-wrapper-box p) {
                    font-family: "${config.font}", serif;
                    font-size: ${config.size}pt;
                    line-height: ${config.spacing};
                    text-align: ${config.textAlign};
                    text-indent: ${config.textIndent}cm;
                    margin-top: 0pt;
                    margin-bottom: ${config.paraSpacing}pt;
                  }

                  /* 2. SỬA LỖI TIÊU ĐỀ: Tự động phát hiện dòng chỉ có chữ in đậm -> Hủy thụt lề để thẳng hàng với bảng */
                  .document-content p:has(> strong:only-child),
                  .document-content p:has(> b:only-child) {
                    text-indent: 0cm !important;
                  }
                  
                  /* 3. SỬA LỖI DANH SÁCH (Bullets): Căn chỉnh lại dấu gạch đầu dòng, hủy thụt lề kép */
                  .document-content ul, .document-content ol {
                    margin-top: 0;
                    margin-bottom: ${config.paraSpacing}pt;
                    padding-left: ${parseFloat(config.textIndent.toString()) + 0.5}cm;
                  }
                  .document-content li p {
                    text-indent: 0cm !important;
                    margin-bottom: 0 !important;
                  }
                  
                  /* Formatting AI block */
                  .ai-wrapper-box {
                    border-top: 1px dashed black;
                    margin-top: 2rem;
                    padding-top: 1rem;
                    font-family: "${config.font}", serif;
                    font-size: ${config.size}pt;
                  }
                  .ai-wrapper-box .ai-title-line {
                    text-align: center;
                    font-weight: bold;
                    font-size: 1.1em;
                    margin-bottom: 1rem;
                  }
                  .ai-wrapper-box .ai-text-line {
                    text-align: left;
                    margin-bottom: 0.5rem;
                  }
                  
                  /* ========================================= */
                  /* SỬA LỖI BẢNG (TABLE) TRONG GIÁO ÁN / NĐ 30 */
                  /* ========================================= */

                  /* BẢNG CHUNG: Ép bảng nằm trọn trong trang, không tràn viền */
                  .document-content table { 
                    width: 100%; 
                    max-width: 100%;
                    box-sizing: border-box;
                    border-collapse: collapse; 
                    margin-top: 6pt;
                    margin-bottom: 12pt; 
                  }
                  .document-content td, .document-content th { 
                    border: 1px solid #000; 
                    padding: 0.5rem; 
                    text-align: left !important; 
                  }

                  /* Xóa thụt lề cho chữ nằm BÊN TRONG bảng */
                  .document-content table p, .document-content td p, .document-content th p {
                    text-indent: 0cm !important;
                    margin-bottom: 0pt !important;
                  }
                  
                  /* BẢNG ĐẦU TIÊN (Quốc hiệu - Tiêu ngữ) */
                  .document-content table:first-of-type,
                  .document-content table:first-of-type td,
                  .document-content table:first-of-type th {
                    border: none !important; 
                    padding: 0 !important;
                  }
                  .document-content table:first-of-type td p,
                  .document-content table:first-of-type th p {
                    text-align: center !important; 
                    margin-bottom: 2pt !important; 
                  }
                  .document-content table:first-of-type tr:first-child td p {
                     font-weight: bold; 
                  }

                  /* BẢNG CUỐI CÙNG (Nơi nhận - Chữ ký) */
                  .document-content table:last-of-type,
                  .document-content table:last-of-type td,
                  .document-content table:last-of-type th {
                    border: none !important; 
                    padding: 0 !important;
                  }
                  .document-content table:last-of-type td:first-child p {
                    text-align: left !important;
                  }
                  .document-content table:last-of-type td:first-child p:first-child {
                    font-weight: bold;
                    font-style: italic; 
                  }
                  .document-content table:last-of-type td:last-child p {
                    text-align: center !important; 
                  }
                  .document-content table:last-of-type td:last-child p:first-child {
                    font-weight: bold; 
                  }

                  /* Giữ tỷ lệ ảnh */
                  .document-content img { 
                    max-width: 100%; 
                    display: inline-block;
                  }
                `}</style>

                {!hasFile ? (
                  <div className="h-full w-full flex flex-col items-center justify-center text-slate-400 py-32 absolute inset-0">
                    <FileQuestion className="w-16 h-16 mb-4 text-slate-300" />
                    <p>Tài liệu sẽ hiển thị ở đây</p>
                  </div>
                ) : (
                  <div
                    ref={previewRef}
                    className="document-content h-full w-full outline-none"
                    dangerouslySetInnerHTML={{ __html: docHtml }}
                  />
                )}
              </div>
            </div>
          </section>
        </main>

        <footer className="mt-10 pt-6 border-t border-slate-300 text-center pb-8">
          <p className="text-xl font-bold text-blue-800 uppercase tracking-wide">
            Trường PTDTBT Tiểu học Giàng Chu Phìn
          </p>
          <p className="text-sm text-slate-500 italic mt-2">
            (Công cụ hỗ trợ Chuẩn hóa tài liệu Giáo dục - Nhanh chóng & Hiệu quả)
          </p>
        </footer>
      </div>

      {/* AI MODAL */}
      {isAIModalOpen && (
        <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-4xl flex flex-col max-h-[90vh] overflow-hidden">
            <div className="p-4 border-b border-slate-200 flex justify-between items-center bg-indigo-50">
              <h3 className="text-lg font-bold text-indigo-900 flex items-center gap-2">
                <AlignLeft className="w-5 h-5" />
                Kết Quả Phân Tích & Chỉnh Sửa AI
              </h3>
              <button
                onClick={() => setIsAIModalOpen(false)}
                className="text-slate-400 hover:text-slate-700 font-bold transition-colors"
                aria-label="Close Modal"
              >
                <X className="w-6 h-6" />
              </button>
            </div>

            <div className="bg-blue-50 px-4 py-2 text-xs text-blue-800 border-b border-blue-200 flex items-center gap-2">
              <span className="text-base">⚡</span>
              <span>
                Hệ thống quét <b>TOÀN BỘ</b> tài liệu. Tùy theo độ dài, thời gian AI phản
                hồi có thể dao động từ 10-30 giây.
              </span>
            </div>

            <div className="p-6 overflow-y-auto flex-1 bg-slate-50">
              {isAILoading ? (
                <div className="flex flex-col items-center justify-center py-20">
                  <Loader2 className="animate-spin text-indigo-600 w-12 h-12 mb-4" />
                  <p className="text-indigo-600 font-medium">
                    AI đang phân tích và xử lý, vui lòng đợi giây lát...
                  </p>
                </div>
              ) : aiError ? (
                <div className="text-red-600 font-medium bg-red-50 p-4 rounded border border-red-200">
                  ❌ {aiError}
                </div>
              ) : (
                <div className="prose prose-sm md:prose-base prose-indigo max-w-none bg-white p-6 rounded-lg border border-slate-200 shadow-sm font-sans">
                  <ReactMarkdown>{aiResult}</ReactMarkdown>
                </div>
              )}
            </div>

            <div className="p-4 border-t flex justify-end gap-3 bg-white">
              <button
                onClick={() => setIsAIModalOpen(false)}
                className="px-5 py-2 text-slate-600 font-medium hover:bg-slate-100 rounded-md border border-slate-300 transition-colors"
              >
                Đóng
              </button>
              {!isAILoading && !aiError && aiResult && (
                <button
                  onClick={appendAIToDocument}
                  className="px-5 py-2 bg-indigo-600 hover:bg-indigo-700 text-white font-medium rounded-md transition-colors shadow-sm"
                >
                  Ghi kết quả này vào cuối văn bản
                </button>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
