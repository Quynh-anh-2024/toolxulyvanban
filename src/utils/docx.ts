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
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();
}

function hasAnyKeyword(text: string, keywords: string[]): boolean {
  return keywords.some((keyword) => text.includes(normalizeText(keyword)));
}

function getElementText(el: Element): string {
  return normalizeText(el.textContent || "");
}

function tableLooksLikeAdministrativeHeader(table: HTMLTableElement, index: number): boolean {
  const text = getElementText(table);
  const hasNationalHeader =
    text.includes("CONG HOA XA HOI CHU NGHIA VIET NAM") ||
    (text.includes("DOC LAP") && text.includes("TU DO") && text.includes("HANH PHUC"));

  const hasAgencyHeader = hasAnyKeyword(text, [
    "UBND",
    "UY BAN NHAN DAN",
    "PHONG GIAO DUC",
    "PHONG GD",
    "SO GIAO DUC",
    "TRUONG",
    "SO:",
    "Số:",
  ]);

  return index <= 2 && (hasNationalHeader || hasAgencyHeader);
}

function tableLooksLikeSignatureBlock(table: HTMLTableElement, index: number, total: number): boolean {
  const text = getElementText(table);
  const nearEnd = index >= Math.max(0, total - 2);
  const hasSignatureKeyword = hasAnyKeyword(text, [
    "NOI NHAN",
    "Nơi nhận",
    "THU TRUONG",
    "HIỆU TRƯỞNG",
    "HIEU TRUONG",
    "NGUOI LAP",
    "NGƯỜI LẬP",
    "NGUOI KY",
    "Ký tên",
    "KY TEN",
  ]);
  return nearEnd && hasSignatureKeyword;
}

function tableHasLikelyVisibleBorder(table: HTMLTableElement): boolean {
  const html = table.outerHTML.toLowerCase();
  return (
    html.includes("border") ||
    html.includes("mso-border") ||
    html.includes("border-top") ||
    html.includes("border-left") ||
    html.includes("1px") ||
    html.includes("solid")
  );
}

function markTableCells(table: HTMLTableElement) {
  table.querySelectorAll("td, th").forEach((cell) => {
    cell.classList.add("doc-table-cell");
  });
}


function ensureAdminHeaderColumns(table: HTMLTableElement) {
  const existingColGroup = table.querySelector("colgroup");
  if (existingColGroup) existingColGroup.remove();

  const colgroup = document.createElement("colgroup");
  const leftCol = document.createElement("col");
  const rightCol = document.createElement("col");
  leftCol.className = "admin-left-col";
  rightCol.className = "admin-right-col";
  colgroup.appendChild(leftCol);
  colgroup.appendChild(rightCol);
  table.insertBefore(colgroup, table.firstChild);
}

function mergeBrokenNationalHeaderParagraphs(cell: Element) {
  const paragraphs = Array.from(cell.querySelectorAll("p")) as HTMLParagraphElement[];

  for (let index = 0; index < paragraphs.length - 1; index += 1) {
    const current = paragraphs[index];
    const next = paragraphs[index + 1];
    const currentText = normalizeText(current.textContent || "");
    const nextText = normalizeText(next.textContent || "");

    if (
      currentText.includes("CONG HOA XA HOI CHU NGHIA VIET") &&
      (nextText === "NAM" || nextText === "VIET NAM")
    ) {
      current.textContent = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";
      next.remove();
    }
  }
}

function markAdministrativeHeaderCellContent(table: HTMLTableElement) {
  ensureAdminHeaderColumns(table);

  const rows = Array.from(table.rows);
  rows.forEach((row) => {
    Array.from(row.cells).forEach((cell, cellIndex) => {
      cell.classList.add(cellIndex === 0 ? "admin-left-cell" : "admin-right-cell");
      mergeBrokenNationalHeaderParagraphs(cell);

      const paragraphs = Array.from(cell.querySelectorAll("p")) as HTMLParagraphElement[];
      paragraphs.forEach((paragraph) => {
        const text = normalizeText(paragraph.textContent || "");
        paragraph.classList.remove(
          "admin-agency-line",
          "admin-unit-line",
          "admin-number-line",
          "admin-national-line",
          "admin-motto-line",
          "admin-date-line"
        );

        if (!text) return;

        if (text.includes("CONG HOA XA HOI CHU NGHIA VIET NAM")) {
          paragraph.textContent = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";
          paragraph.classList.add("admin-national-line");
          return;
        }

        if (text.includes("DOC LAP") && text.includes("TU DO") && text.includes("HANH PHUC")) {
          paragraph.textContent = "Độc lập - Tự do - Hạnh phúc";
          paragraph.classList.add("admin-motto-line");
          return;
        }

        if (/^(.*NGAY\s+\d{1,2}\s+THANG\s+\d{1,2}\s+NAM\s+\d{4}|.*NGAY\s+\d{1,2}\s+THANG\s+\d{1,2}\s+NAM)$/.test(text)) {
          paragraph.classList.add("admin-date-line");
          return;
        }

        if (/^\d{4}$/.test(text)) {
          const prev = paragraph.previousElementSibling as HTMLElement | null;
          if (prev?.classList.contains("admin-date-line")) {
            prev.textContent = `${prev.textContent || ""} ${paragraph.textContent || ""}`.replace(/\s+/g, " ").trim();
            paragraph.remove();
          }
          return;
        }

        if (text.startsWith("SO:") || text.startsWith("SỐ:") || text.startsWith("SO ")) {
          paragraph.classList.add("admin-number-line");
          return;
        }

        if (cellIndex === 0 && hasAnyKeyword(text, ["UBND", "UY BAN NHAN DAN", "PHONG GIAO DUC", "PHONG GD", "SO GIAO DUC"])) {
          paragraph.classList.add("admin-agency-line");
          return;
        }

        if (cellIndex === 0 && hasAnyKeyword(text, ["TRUONG", "TRƯỜNG", "TRUNG TAM", "PHONG"])) {
          paragraph.classList.add("admin-unit-line");
        }
      });
    });
  });
}


function paragraphLooksLikeMainAdministrativeTitle(text: string): boolean {
  const normalized = normalizeText(text);
  const compact = normalized.replace(/[^A-Z0-9 ]/g, " ").replace(/\s+/g, " ").trim();

  return [
    "KE HOACH",
    "QUYET DINH",
    "CONG VAN",
    "TO TRINH",
    "BAO CAO",
    "BIEN BAN",
    "THONG BAO",
    "GIAY MOI",
    "GIAY MOI HOP",
    "DON",
    "DE NGHI",
    "PHIEU",
    "DANH SACH",
    "CHUONG TRINH",
  ].some((title) => compact === title || compact.startsWith(title + " "));
}

function paragraphLooksLikeAdministrativeSubtitle(text: string): boolean {
  const normalized = normalizeText(text);
  if (!normalized) return false;
  if (normalized.length > 180) return false;
  if (/^(I|II|III|IV|V|A|B|C)[\.\)]\s/.test(normalized)) return false;
  if (/^\d+[\.\)]\s/.test(normalized)) return false;
  if (normalized.startsWith("CAN CU") || normalized.startsWith("THUC HIEN")) return false;
  return true;
}

function markAdministrativeTitleParagraphs(root: HTMLElement) {
  const paragraphs = Array.from(root.querySelectorAll("p")) as HTMLParagraphElement[];
  const contentParagraphs = paragraphs.filter((p) => !p.closest("table"));

  const searchArea = contentParagraphs.slice(0, 12);
  const mainTitleIndex = searchArea.findIndex((p) => paragraphLooksLikeMainAdministrativeTitle(p.textContent || ""));

  if (mainTitleIndex < 0) return;

  const mainTitle = searchArea[mainTitleIndex];
  mainTitle.classList.add("doc-main-title");

  let subtitleCount = 0;
  for (let i = mainTitleIndex + 1; i < searchArea.length && subtitleCount < 4; i += 1) {
    const paragraph = searchArea[i];
    const text = paragraph.textContent || "";
    if (!paragraphLooksLikeAdministrativeSubtitle(text)) break;

    paragraph.classList.add("doc-sub-title");
    subtitleCount += 1;
  }
}

function cleanupEmptyParagraphs(root: HTMLElement) {
  root.querySelectorAll("p").forEach((p) => {
    const text = (p.textContent || "").replace(/\u00a0/g, " ").trim();
    if (!text && !p.querySelector("img")) {
      p.classList.add("doc-empty-paragraph");
    }
  });
}

export function analyzeDocumentHtml(htmlString: string): DocumentStats {
  const tempDiv = document.createElement("div");
  tempDiv.innerHTML = htmlString;

  const tables = Array.from(tempDiv.querySelectorAll("table")) as HTMLTableElement[];
  const images = tempDiv.querySelectorAll("img");
  const firstTable = tables[0];

  const administrativeHeaderDetected = tables.some((table, index) =>
    tableLooksLikeAdministrativeHeader(table, index)
  );
  const signatureTableDetected = tables.some((table, index) =>
    tableLooksLikeSignatureBlock(table, index, tables.length)
  );
  const firstTableHasBorderRisk = !!firstTable && tableHasLikelyVisibleBorder(firstTable);

  const warningMessages: string[] = [];
  if (tables.length > 0) {
    warningMessages.push(`Phát hiện ${tables.length} bảng/khung trong tài liệu.`);
  }
  if (firstTableHasBorderRisk) {
    warningMessages.push("Bảng/khung đầu văn bản có dấu hiệu có viền. Không nên tự động xóa viền theo vị trí.");
  }
  if (administrativeHeaderDetected) {
    warningMessages.push("Có dấu hiệu phần thể thức hành chính: cơ quan ban hành/quốc hiệu/tiêu ngữ/số ký hiệu.");
  }
  if (images.length > 0) {
    warningMessages.push(`Phát hiện ${images.length} hình ảnh. Khi xuất Word cần kiểm tra lại kích thước ảnh.`);
  }

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

  cleanupEmptyParagraphs(tempDiv);

  const tables = Array.from(tempDiv.querySelectorAll("table")) as HTMLTableElement[];

  tables.forEach((table, index) => {
    table.classList.add("doc-table");
    table.removeAttribute("border");
    markTableCells(table);

    const isHeader = tableLooksLikeAdministrativeHeader(table, index);
    const isSignature = tableLooksLikeSignatureBlock(table, index, tables.length);

    table.classList.remove(
      "admin-header-table",
      "admin-header-preview-frame",
      "signature-table",
      "form-frame-table",
      "preserved-frame-table"
    );

    if (options.mode === "textOnly") {
      table.classList.add("form-frame-table");
      return;
    }

    if (isHeader) {
      // Bảng đầu văn bản hành chính thường chỉ là bảng kỹ thuật để căn 2 cột.
      // Mặc định ẩn viền trong preview và khi xuất Word; chỉ hiện nét đứt khi người dùng bật khung hỗ trợ.
      table.classList.add("admin-header-table");
      markAdministrativeHeaderCellContent(table);
      if (options.preserveFirstFrame) {
        table.classList.add("admin-header-preview-frame");
      }
      return;
    }

    if (isSignature && options.mode === "nd30") {
      table.classList.add("signature-table");
      return;
    }

    if (index === 0 && options.preserveFirstFrame) {
      // Chỉ giữ viền bảng đầu khi KHÔNG phải phần thể thức hành chính
      // ví dụ: phiếu biểu mẫu, bảng thống kê, khung thông tin cần in thật.
      table.classList.add("preserved-frame-table");
      return;
    }

    table.classList.add("form-frame-table");
  });

  markAdministrativeTitleParagraphs(tempDiv);

  const stats = analyzeDocumentHtml(tempDiv.innerHTML);
  return { html: tempDiv.innerHTML, stats };
}

export function exportToWord(
  previewElement: HTMLElement | null,
  fileName: string,
  config: FormattingConfig
) {
  if (!previewElement) return;

  const {
    font,
    size,
    spacing,
    textAlign,
    textIndent,
    paraSpacing,
    leftMargin,
    rightMargin,
  } = config;

  let lineSpacingPercent = "150%";
  if (spacing === "1.0") lineSpacingPercent = "100%";
  if (spacing === "1.15") lineSpacingPercent = "115%";
  if (spacing === "1.5") lineSpacingPercent = "150%";

  const cloneDiv = previewElement.cloneNode(true) as HTMLElement;
  cloneDiv.querySelectorAll("*").forEach((el) => {
    if (el.tagName.toLowerCase() !== "img") {
      const existingClass = el.getAttribute("class") || "";
      const isMarkedTable = /doc-table|admin-header-table|admin-header-preview-frame|signature-table|form-frame-table|preserved-frame-table|doc-table-cell|doc-main-title|doc-sub-title/.test(existingClass);
      if (!isMarkedTable) el.removeAttribute("style");
    }
  });

  const imagesToEmbed: { id: string; type: string; data: string }[] = [];
  cloneDiv.querySelectorAll("img").forEach((img, index) => {
    const src = img.getAttribute("src");
    if (src && src.startsWith("data:image/")) {
      const parts = src.match(/^data:(image\/[^;]+);base64,(.*)$/);
      if (parts && parts.length === 3) {
        const id = `image_${index}`;
        imagesToEmbed.push({ id, type: parts[1], data: parts[2] });
        img.setAttribute("src", `cid:${id}`);
      }
    }
  });

  const cleanHTML = cloneDiv.innerHTML;

  const fullHTML = `
  <html xmlns:o="urn:schemas-microsoft-com:office:office"
        xmlns:w="urn:schemas-microsoft-com:office:word"
        xmlns="http://www.w3.org/TR/REC-html40">
  <head>
      <meta charset="utf-8">
      <title>Export</title>
      <style>
          body {
              margin: 0;
              padding: 0;
              font-family: "${font}", serif;
              font-size: ${size}pt;
          }
          @page WordSection1 {
              size: 21cm 29.7cm;
              margin: 2cm ${rightMargin}cm 2cm ${leftMargin}cm;
              mso-header-margin: 35.4pt;
              mso-footer-margin: 35.4pt;
              mso-paper-source: 0;
          }
          div.WordSection1 { page: WordSection1; }
          p {
              margin: 0;
              padding: 0;
              font-family: "${font}", serif;
              font-size: ${size}pt;
              text-align: ${textAlign};
              text-indent: ${textIndent}cm;
              margin-bottom: ${paraSpacing}pt;
              line-height: ${lineSpacingPercent};
              mso-line-height-rule: exactly;
              mso-pagination: none;
              page-break-inside: auto;
          }
          .doc-empty-paragraph { margin-bottom: 0pt !important; line-height: 1pt !important; }
          h1, h2, h3, h4, h5, h6 {
              font-family: "${font}", serif;
              page-break-after: auto;
          }
          .doc-main-title {
              text-align: center !important;
              text-indent: 0cm !important;
              font-weight: bold !important;
              text-transform: uppercase;
              font-size: ${Math.max(size, 14)}pt !important;
              margin-top: 8pt !important;
              margin-bottom: 4pt !important;
              line-height: 120% !important;
          }
          .doc-sub-title {
              text-align: center !important;
              text-indent: 0cm !important;
              font-weight: bold !important;
              margin-top: 0pt !important;
              margin-bottom: 3pt !important;
              line-height: 120% !important;
          }
          table, .doc-table {
              border-collapse: collapse;
              width: 100%;
              max-width: 100%;
              table-layout: fixed;
              margin-top: 6pt;
              margin-bottom: 12pt;
          }
          td, th, .doc-table-cell {
              border: 1pt solid #000;
              padding: 4pt;
              vertical-align: top;
              font-family: "${font}", serif;
              font-size: ${size}pt;
          }
          table p, td p, th p {
              text-align: left !important;
              text-indent: 0cm !important;
              margin-bottom: 0pt !important;
              line-height: 115% !important;
          }
          th, th p {
              text-align: center !important;
              font-weight: bold !important;
          }
          .admin-header-table,
          .admin-header-table td,
          .admin-header-table th,
          .signature-table,
          .signature-table td,
          .signature-table th {
              border: none !important;
          }
          .admin-header-table td,
          .admin-header-table th,
          .signature-table td,
          .signature-table th {
              padding: 0pt 2pt !important;
          }
          .admin-header-table {
              margin-top: 0pt !important;
              margin-bottom: 12pt !important;
              table-layout: fixed !important;
          }
          .admin-header-table col.admin-left-col { width: 45% !important; }
          .admin-header-table col.admin-right-col { width: 55% !important; }
          .admin-header-table p {
              text-align: center !important;
              text-indent: 0cm !important;
              margin-top: 0pt !important;
              margin-bottom: 1pt !important;
              line-height: 115% !important;
              font-family: "Times New Roman", serif !important;
              font-size: 13pt !important;
          }
          .admin-header-table .admin-left-cell,
          .admin-header-table .admin-right-cell { text-align: center !important; }
          .admin-agency-line {
              font-weight: normal !important;
              text-transform: uppercase !important;
          }
          .admin-unit-line {
              font-weight: bold !important;
              text-transform: uppercase !important;
          }
          .admin-number-line {
              text-align: center !important;
              white-space: nowrap !important;
              font-weight: normal !important;
          }
          .admin-national-line {
              font-weight: bold !important;
              text-transform: uppercase !important;
              white-space: nowrap !important;
              letter-spacing: -0.15pt !important;
              font-size: 13pt !important;
          }
          .admin-motto-line {
              font-weight: bold !important;
              white-space: nowrap !important;
              display: inline-block !important;
              padding-bottom: 1pt !important;
              border-bottom: 1pt solid #000 !important;
          }
          .admin-date-line {
              text-align: center !important;
              font-style: italic !important;
              white-space: nowrap !important;
          }
          .signature-table td:first-child p { text-align: left !important; }
          .signature-table td:last-child p { text-align: center !important; }
          .preserved-frame-table,
          .preserved-frame-table td,
          .preserved-frame-table th,
          .form-frame-table,
          .form-frame-table td,
          .form-frame-table th {
              border: 1pt solid #000 !important;
          }
          div.ai-wrapper-box {
              border-top: 1pt dashed #000;
              margin-top: 20pt;
              padding-top: 10pt;
          }
          div.ai-wrapper-box p.ai-title-line {
              text-align: center !important;
              text-indent: 0cm !important;
              font-size: 14pt;
              font-weight: bold;
              margin-bottom: 12pt !important;
          }
          div.ai-wrapper-box p.ai-text-line {
              text-align: left !important;
              text-indent: 0cm !important;
              margin-top: 4pt !important;
              margin-bottom: 4pt !important;
          }
          img { max-width: 100%; display: inline-block; }
      </style>
  </head>
  <body>
      <div class="WordSection1">${cleanHTML}</div>
  </body>
  </html>`;

  const boundary = "----=_NextPart_MHTML_BOUNDARY";
  const htmlBase64 = btoa(unescape(encodeURIComponent(fullHTML)));

  let mhtml = `MIME-Version: 1.0\r\n`;
  mhtml += `Content-Type: multipart/related; boundary="${boundary}"\r\n\r\n`;
  mhtml += `--${boundary}\r\n`;
  mhtml += `Content-Type: text/html; charset="utf-8"\r\n`;
  mhtml += `Content-Transfer-Encoding: base64\r\n\r\n`;
  mhtml += `${htmlBase64.replace(/(.{76})/g, "$1\r\n")}\r\n\r\n`;

  for (const img of imagesToEmbed) {
    mhtml += `--${boundary}\r\n`;
    mhtml += `Content-Type: ${img.type}\r\n`;
    mhtml += `Content-Transfer-Encoding: base64\r\n`;
    mhtml += `Content-ID: <${img.id}>\r\n\r\n`;
    mhtml += `${img.data.replace(/(.{76})/g, "$1\r\n")}\r\n\r\n`;
  }

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

export function processHtmlToRemoveBullets(htmlString: string): string {
  const tempDiv = document.createElement("div");
  tempDiv.innerHTML = htmlString;

  tempDiv.querySelectorAll("li").forEach((li) => {
    const p = document.createElement("p");
    p.innerHTML = li.innerHTML;
    p.className = li.className;
    if (li.parentNode) li.parentNode.replaceChild(p, li);
  });

  tempDiv.querySelectorAll("ul, ol").forEach((list) => {
    const fragment = document.createDocumentFragment();
    while (list.firstChild) fragment.appendChild(list.firstChild);
    if (list.parentNode) list.parentNode.replaceChild(fragment, list);
  });

  return tempDiv.innerHTML;
}
