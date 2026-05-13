export interface FormattingConfig {
  font: string;
  size: number;
  spacing: string;
  textAlign: string;
  textIndent: number;
  paraSpacing: number;
  leftMargin: number;
  rightMargin: number;
}

export type DocumentProcessingMode = "preserve" | "nd30" | "textOnly";

export interface DocumentStats {
  tableCount: number;
  imageCount: number;
  firstTableHasBorderRisk: boolean;
  administrativeHeaderDetected: boolean;
  signatureTableDetected: boolean;
  warningMessages: string[];
}

export interface EnhanceHtmlOptions {
  mode: DocumentProcessingMode;
  preserveFirstFrame: boolean;
}

function normalizeText(value: string): string {
  return (value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();
}

function cleanText(value: string): string {
  return (value || "")
    .replace(/\u00a0/g, " ")
    .replace(/[\t ]+/g, " ")
    .replace(/\s+([,.;:])/g, "$1")
    .trim();
}

function escapeHtml(value: string): string {
  return cleanText(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function hasAny(text: string, keywords: string[]): boolean {
  const normalized = normalizeText(text);
  return keywords.some((keyword) => normalized.includes(normalizeText(keyword)));
}

function paragraphLinesFromElement(element: Element): string[] {
  const paragraphLines = Array.from(element.querySelectorAll("p"))
    .map((p) => cleanText(p.textContent || ""))
    .filter(Boolean);

  if (paragraphLines.length) return paragraphLines;

  return cleanText(element.textContent || "")
    .split(/\r?\n+/)
    .map(cleanText)
    .filter(Boolean);
}

function uniqueLines(lines: string[]): string[] {
  const seen = new Set<string>();
  const result: string[] = [];

  lines.forEach((line) => {
    const cleaned = cleanText(line);
    const key = normalizeText(cleaned);
    if (!cleaned || seen.has(key)) return;
    seen.add(key);
    result.push(cleaned);
  });

  return result;
}

function isAdministrativeHeaderTable(table: HTMLTableElement, index: number): boolean {
  if (index > 2) return false;
  const text = normalizeText(table.textContent || "");
  const hasNational =
    text.includes("CONG HOA XA HOI CHU NGHIA VIET NAM") ||
    (text.includes("DOC LAP") && text.includes("TU DO") && text.includes("HANH PHUC"));
  const hasAgency = hasAny(text, ["UBND", "UY BAN NHAN DAN", "PHONG GIAO DUC", "SO GIAO DUC", "TRUONG", "SỐ:", "SO:"]);
  return hasNational || hasAgency;
}

function isSignatureTable(table: HTMLTableElement, index: number, total: number): boolean {
  const text = normalizeText(table.textContent || "");
  return index >= Math.max(0, total - 2) && hasAny(text, ["NOI NHAN", "HIEU TRUONG", "NGUOI LAP", "NGUOI KY", "KY TEN"]);
}

function tableHasBorderRisk(table?: HTMLTableElement): boolean {
  if (!table) return false;
  const html = table.outerHTML.toLowerCase();
  return html.includes("border") || html.includes("solid") || html.includes("mso-border") || html.includes("1px");
}

function looksLikeAgency(line: string): boolean {
  return hasAny(line, ["UBND", "ỦY BAN NHÂN DÂN", "PHÒNG GIÁO DỤC", "SỞ GIÁO DỤC"]);
}

function looksLikeUnit(line: string): boolean {
  return hasAny(line, ["TRƯỜNG", "TRUONG", "PTDT", "THCS", "TIỂU HỌC", "TIEU HOC", "TRUNG TÂM"]);
}

function looksLikeNumber(line: string): boolean {
  const text = normalizeText(line);
  return text.startsWith("SO:") || text.startsWith("SO ") || text.startsWith("SỐ:") || text.startsWith("SỐ ");
}

function looksLikeDate(line: string): boolean {
  const text = normalizeText(line);
  return text.includes("NGAY") && text.includes("THANG") && text.includes("NAM");
}

function splitMergedHeaderText(line: string): string[] {
  const cleaned = cleanText(line);
  const normalized = normalizeText(cleaned);

  // Nếu Mammoth gộp nhiều đoạn trong một ô thành một dòng dài, tách theo các mốc quen thuộc.
  if (normalized.includes("UBND") && normalized.includes("TRUONG")) {
    const match = cleaned.match(/^(.*?(?:MÈO VẠC|MEO VAC|XÃ[^A-ZĐ]*|XA[^A-ZD]*))(.*)$/i);
    if (match) return [cleanText(match[1]), cleanText(match[2])].filter(Boolean);
  }

  return [cleaned];
}

function extractHeaderData(table: HTMLTableElement) {
  const rawLines = uniqueLines(
    Array.from(table.querySelectorAll("td, th, p"))
      .flatMap((el) => paragraphLinesFromElement(el))
      .concat(paragraphLinesFromElement(table))
      .flatMap(splitMergedHeaderText)
  );

  const agencyLines = rawLines.filter(looksLikeAgency);
  let unitLines = rawLines.filter((line) => looksLikeUnit(line) && !looksLikeAgency(line));
  if (!unitLines.length) {
    unitLines = rawLines.filter((line) => {
      const n = normalizeText(line);
      return !looksLikeAgency(line) && !looksLikeNumber(line) && !looksLikeDate(line) && !n.includes("CONG HOA") && !n.includes("DOC LAP");
    }).slice(0, 2);
  }

  const numberLine = rawLines.find(looksLikeNumber) || "Số:        /KH-...";
  const dateLine = rawLines.find(looksLikeDate) || "..., ngày ... tháng ... năm ...";

  return {
    agencyLines: agencyLines.length ? agencyLines : [""],
    unitLines,
    numberLine,
    dateLine,
    nationalLine: "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM",
    mottoLine: "Độc lập - Tự do - Hạnh phúc",
  };
}

function createParagraph(doc: Document, text: string, className: string): HTMLParagraphElement {
  const p = doc.createElement("p");
  p.className = className;
  p.textContent = cleanText(text);
  return p;
}

function createCleanAdminHeader(table: HTMLTableElement, showPreviewFrame: boolean): HTMLTableElement {
  const doc = table.ownerDocument;
  const data = extractHeaderData(table);
  const newTable = doc.createElement("table");
  newTable.className = `admin-header-table${showPreviewFrame ? " admin-header-preview-frame" : ""}`;
  newTable.setAttribute("cellspacing", "0");
  newTable.setAttribute("cellpadding", "0");
  newTable.setAttribute("border", "0");

  const colgroup = doc.createElement("colgroup");
  const colLeft = doc.createElement("col");
  const colRight = doc.createElement("col");
  colLeft.className = "admin-left-col";
  colRight.className = "admin-right-col";
  colgroup.append(colLeft, colRight);
  newTable.appendChild(colgroup);

  const row1 = newTable.insertRow();
  const leftTop = row1.insertCell();
  const rightTop = row1.insertCell();
  leftTop.className = "admin-left-cell";
  rightTop.className = "admin-right-cell";

  data.agencyLines.forEach((line) => leftTop.appendChild(createParagraph(doc, line, "admin-agency-line")));
  data.unitLines.forEach((line) => leftTop.appendChild(createParagraph(doc, line, "admin-unit-line")));

  rightTop.appendChild(createParagraph(doc, data.nationalLine, "admin-national-line"));
  rightTop.appendChild(createParagraph(doc, data.mottoLine, "admin-motto-line"));

  const row2 = newTable.insertRow();
  const leftBottom = row2.insertCell();
  const rightBottom = row2.insertCell();
  leftBottom.className = "admin-left-cell";
  rightBottom.className = "admin-right-cell";

  leftBottom.appendChild(createParagraph(doc, data.numberLine, "admin-number-line"));
  rightBottom.appendChild(createParagraph(doc, data.dateLine, "admin-date-line"));

  return newTable;
}

function markTitles(root: HTMLElement) {
  const titleWords = ["KẾ HOẠCH", "QUYẾT ĐỊNH", "THÔNG BÁO", "BÁO CÁO", "BIÊN BẢN", "TỜ TRÌNH", "CÔNG VĂN"];
  const paragraphs = Array.from(root.querySelectorAll("p")).filter((p) => !p.closest("table")) as HTMLParagraphElement[];
  const limit = paragraphs.slice(0, 15);
  const index = limit.findIndex((p) => titleWords.includes(normalizeText(p.textContent || "")));

  if (index < 0) return;

  limit[index].classList.add("doc-main-title");
  for (let i = index + 1; i < Math.min(index + 4, limit.length); i += 1) {
    const text = normalizeText(limit[i].textContent || "");
    if (!text || text.startsWith("CAN CU") || /^(I|II|III|IV|V|A|B|C|\d+)[\.\)]/.test(text)) break;
    limit[i].classList.add("doc-sub-title");
  }
}

function cleanup(root: HTMLElement) {
  root.querySelectorAll("p").forEach((p) => {
    const text = cleanText(p.textContent || "");
    if (!text && !p.querySelector("img")) p.classList.add("doc-empty-paragraph");
  });
}

export function analyzeDocumentHtml(htmlString: string): DocumentStats {
  const tempDiv = document.createElement("div");
  tempDiv.innerHTML = htmlString;
  const tables = Array.from(tempDiv.querySelectorAll("table")) as HTMLTableElement[];
  const images = tempDiv.querySelectorAll("img");

  const administrativeHeaderDetected = tables.some((table, index) => isAdministrativeHeaderTable(table, index));
  const signatureTableDetected = tables.some((table, index) => isSignatureTable(table, index, tables.length));
  const firstTableHasBorderRisk = tableHasBorderRisk(tables[0]);
  const warningMessages: string[] = [];

  if (tables.length) warningMessages.push(`Phát hiện ${tables.length} bảng/khung trong tài liệu.`);
  if (firstTableHasBorderRisk) warningMessages.push("Bảng/khung đầu văn bản có dấu hiệu có viền. Không nên tự động xóa viền theo vị trí.");
  if (administrativeHeaderDetected) warningMessages.push("Có dấu hiệu phần thể thức hành chính: cơ quan ban hành/quốc hiệu/tiêu ngữ/số ký hiệu.");
  if (images.length) warningMessages.push(`Phát hiện ${images.length} hình ảnh. Khi xuất Word cần kiểm tra lại kích thước ảnh.`);

  return {
    tableCount: tables.length,
    imageCount: images.length,
    firstTableHasBorderRisk,
    administrativeHeaderDetected,
    signatureTableDetected,
    warningMessages,
  };
}

export function enhanceAdministrativeDocumentHtml(
  htmlString: string,
  options: EnhanceHtmlOptions
): { html: string; stats: DocumentStats } {
  const tempDiv = document.createElement("div");
  tempDiv.innerHTML = htmlString;
  cleanup(tempDiv);

  const tables = Array.from(tempDiv.querySelectorAll("table")) as HTMLTableElement[];
  tables.forEach((table, index) => {
    const isHeader = isAdministrativeHeaderTable(table, index);
    const isSignature = isSignatureTable(table, index, tables.length);

    if (isHeader && options.mode !== "textOnly") {
      table.replaceWith(createCleanAdminHeader(table, options.preserveFirstFrame));
      return;
    }

    table.classList.add(isSignature && options.mode === "nd30" ? "signature-table" : "form-frame-table");
    if (index === 0 && options.preserveFirstFrame && !isHeader) table.classList.add("preserved-frame-table");
  });

  markTitles(tempDiv);
  const stats = analyzeDocumentHtml(tempDiv.innerHTML);
  return { html: tempDiv.innerHTML, stats };
}

export function processHtmlToRemoveBullets(htmlString: string): string {
  const tempDiv = document.createElement("div");
  tempDiv.innerHTML = htmlString;

  tempDiv.querySelectorAll("li").forEach((li) => {
    const p = document.createElement("p");
    p.innerHTML = li.innerHTML;
    li.replaceWith(p);
  });

  tempDiv.querySelectorAll("ul, ol").forEach((list) => {
    const fragment = document.createDocumentFragment();
    while (list.firstChild) fragment.appendChild(list.firstChild);
    list.replaceWith(fragment);
  });

  return tempDiv.innerHTML;
}

function makeExportCss(config: FormattingConfig): string {
  const spacingMap: Record<string, string> = { "1.0": "100%", "1.15": "115%", "1.5": "150%" };
  const lineHeight = spacingMap[config.spacing] || "150%";
  return `
    body { font-family: "${config.font}", serif; font-size: ${config.size}pt; color:#000; }
    p { margin:0 0 ${config.paraSpacing}pt 0; text-indent:${config.textIndent}cm; text-align:${config.textAlign}; line-height:${lineHeight}; font-family:"${config.font}",serif; font-size:${config.size}pt; }
    .doc-empty-paragraph { display:none; }
    table { border-collapse:collapse; width:100%; table-layout:fixed; margin:6pt 0 10pt 0; }
    td, th { border:1pt solid #000; padding:4pt; vertical-align:top; font-family:"${config.font}",serif; font-size:${config.size}pt; }
    table p, td p, th p { text-indent:0 !important; margin:0 0 2pt 0 !important; line-height:115% !important; }
    .admin-header-table { width:106% !important; margin-left:-3% !important; margin-right:-3% !important; border-collapse:collapse !important; table-layout:fixed !important; border:none !important; margin-top:0 !important; margin-bottom:10pt !important; }
    .admin-header-table col.admin-left-col { width:36% !important; }
    .admin-header-table col.admin-right-col { width:64% !important; }
    .admin-header-table td { border:none !important; padding:0 3pt !important; vertical-align:top !important; text-align:center !important; }
    .admin-header-table p { font-family:"Times New Roman",serif !important; text-align:center !important; text-indent:0 !important; margin:0 0 1pt 0 !important; line-height:115% !important; }
    .admin-agency-line { font-size:13pt !important; font-weight:normal !important; text-transform:uppercase !important; }
    .admin-unit-line { font-size:13pt !important; font-weight:bold !important; text-transform:uppercase !important; }
    .admin-number-line { font-size:13pt !important; font-weight:normal !important; white-space:nowrap !important; }
    .admin-national-line { font-size:11.2pt !important; font-weight:bold !important; text-transform:uppercase !important; white-space:nowrap !important; letter-spacing:-.25pt !important; }
    .admin-motto-line { font-size:12.5pt !important; font-weight:bold !important; white-space:nowrap !important; display:inline-block !important; border-bottom:1pt solid #000 !important; padding-bottom:1pt !important; }
    .admin-date-line { font-size:13pt !important; font-style:italic !important; white-space:nowrap !important; }
    .admin-header-preview-frame, .admin-header-preview-frame td { border:none !important; }
    .signature-table, .signature-table td { border:none !important; }
    .doc-main-title { text-align:center !important; text-indent:0 !important; font-weight:bold !important; text-transform:uppercase !important; font-size:${Math.max(config.size, 14)}pt !important; margin:8pt 0 4pt 0 !important; line-height:120% !important; }
    .doc-sub-title { text-align:center !important; text-indent:0 !important; font-weight:bold !important; font-size:${config.size}pt !important; margin:0 0 3pt 0 !important; line-height:120% !important; }
    img { max-width:100%; }
  `;
}

export async function exportToWord(
  previewElement: HTMLElement | null,
  fileName: string,
  config: FormattingConfig
) {
  if (!previewElement) return;

  const clone = previewElement.cloneNode(true) as HTMLElement;
  clone.querySelectorAll(".admin-header-preview-frame").forEach((el) => el.classList.remove("admin-header-preview-frame"));

  const css = makeExportCss(config);
  const fullHtml = `
  <html xmlns:o="urn:schemas-microsoft-com:office:office"
        xmlns:w="urn:schemas-microsoft-com:office:word"
        xmlns="http://www.w3.org/TR/REC-html40">
    <head>
      <meta charset="utf-8">
      <title>Export</title>
      <style>
        @page WordSection1 {
          size: 21cm 29.7cm;
          margin: 2cm ${config.rightMargin}cm 2cm ${config.leftMargin}cm;
          mso-paper-source: 0;
        }
        div.WordSection1 { page: WordSection1; }
        ${css}
      </style>
    </head>
    <body><div class="WordSection1">${clone.innerHTML}</div></body>
  </html>`;

  const boundary = "----=_NextPart_ChuanND30";
  const htmlBase64 = btoa(unescape(encodeURIComponent(fullHtml)));
  let mhtml = "MIME-Version: 1.0\r\n";
  mhtml += `Content-Type: multipart/related; boundary="${boundary}"\r\n\r\n`;
  mhtml += `--${boundary}\r\n`;
  mhtml += "Content-Type: text/html; charset=\"utf-8\"\r\n";
  mhtml += "Content-Transfer-Encoding: base64\r\n\r\n";
  mhtml += `${htmlBase64.replace(/(.{76})/g, "$1\r\n")}\r\n\r\n`;
  mhtml += `--${boundary}--\r\n`;

  const blob = new Blob([mhtml], { type: "application/msword" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `${fileName}_ChuanND30.doc`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(a.href);
}
