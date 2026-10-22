// Cột xem trước A4 (bản xem trước + nút Vừa màn hình/100% + CSS của trang
// giấy). Tách khỏi App.tsx (đợt B5); state vẫn ở App, truyền qua props.
import type { RefObject } from "react";
import { CheckCircle2, FileQuestion } from "lucide-react";
import { cn } from "../lib/utils";
import type { FormattingConfig } from "../utils/docx";

export interface PreviewPaneProps {
  mobilePane: "tools" | "preview";
  modeSummary: string;
  previewFit: boolean;
  setPreviewFit: (fit: boolean) => void;
  previewScrollRef: RefObject<HTMLDivElement | null>;
  previewScale: number;
  previewRef: RefObject<HTMLDivElement | null>;
  hasFile: boolean;
  section: "pdf" | "excel" | "word";
  docHtml: string;
  config: FormattingConfig;
}

export function PreviewPane(props: PreviewPaneProps) {
  const {
    mobilePane, modeSummary, previewFit, setPreviewFit, previewScrollRef,
    previewScale, previewRef, hasFile, section, docHtml, config,
  } = props;
  return (
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
  );
}
