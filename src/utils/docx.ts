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


function cleanDisplayText(value: string): string {
  return (value || "")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .replace(/\s+([,.;:])/g, "$1")
    .trim();
}

function getPlainLinesFromCell(cell?: HTMLTableCellElement | null): string[] {
  if (!cell) return [];
  const paragraphLines = Array.from(cell.querySelectorAll("p"))
    .map((p) => cleanDisplayText(p.textContent || ""))
    .filter(Boolean);

  if (paragraphLines.length > 0) return paragraphLines;

  return (cell.textContent || "")
    .split(/\r?\n+/)
    .map(cleanDisplayText)
    .filter(Boolean);
}

function uniqueCleanLines(lines: string[]): string[] {
  const seen = new Set<string>();
  const result: string[] = [];
  lines.forEach((line) => {
    const cleaned = cleanDisplayText(line);
    if (!cleaned) return;
    const key = normalizeText(cleaned);
    if (seen.has(key)) return;
    seen.add(key);
    result.push(cleaned);
  });
  return result;
}

function firstLineMatching(lines: string[], matcher: (normalized: string) => boolean): string | undefined {
  return lines.find((line) => matcher(normalizeText(line)));
}

function lineContainsDate(normalized: string): boolean {
  return normalized.includes("NGAY") && normalized.includes("THANG") && normalized.includes("NAM");
}

function normalizeHeaderLine(line: string): string {
  const normalized = normalizeText(line);

  if (normalized.includes("CONG HOA XA HOI CHU NGHIA VIET NAM")) {
    return "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";
  }

  if (normalized.includes("DOC LAP") && normalized.includes("TU DO") && normalized.includes("HANH PHUC")) {
    return "Độc lập - Tự do - Hạnh phúc";
  }

  return cleanDisplayText(line);
}

function extractAdministrativeHeaderData(table: HTMLTableElement) {
  const rows = Array.from(table.rows);
  const row0 = rows[0];
  const row1 = rows[1];

  const leftTop = uniqueCleanLines(getPlainLinesFromCell(row0?.cells?.[0] || null));
  const rightTop = uniqueCleanLines(getPlainLinesFromCell(row0?.cells?.[1] || null));
  const leftBottom = uniqueCleanLines(getPlainLinesFromCell(row1?.cells?.[0] || null));
  const rightBottom = uniqueCleanLines(getPlainLinesFromCell(row1?.cells?.[1] || null));

  const allLines = uniqueCleanLines(
    Array.from(table.querySelectorAll("p"))
      .map((p) => cleanDisplayText(p.textContent || ""))
      .filter(Boolean)
      .concat(
        (table.textContent || "")
          .split(/\r?\n+/)
          .map(cleanDisplayText)
          .filter(Boolean)
      )
  );

  const nationalLine =
    firstLineMatching(rightTop, (text) => text.includes("CONG HOA XA HOI CHU NGHIA VIET NAM")) ||
    firstLineMatching(allLines, (text) => text.includes("CONG HOA XA HOI CHU NGHIA VIET NAM")) ||
    "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";

  const mottoLine =
    firstLineMatching(rightTop, (text) => text.includes("DOC LAP") && text.includes("TU DO") && text.includes("HANH PHUC")) ||
    firstLineMatching(allLines, (text) => text.includes("DOC LAP") && text.includes("TU DO") && text.includes("HANH PHUC")) ||
    "Độc lập - Tự do - Hạnh phúc";

  const numberLine =
    firstLineMatching(leftBottom, (text) => text.startsWith("SO:") || text.startsWith("SO ")) ||
    firstLineMatching(allLines, (text) => text.startsWith("SO:") || text.startsWith("SO ")) ||
    "Số:        /KH-...";

  const dateLine =
    firstLineMatching(rightBottom, lineContainsDate) ||
    firstLineMatching(allLines, lineContainsDate) ||
    rightBottom[0] ||
    "..., ngày ... tháng ... năm ...";

  let agencyLines = leftTop.filter((line) => {
    const text = normalizeText(line);
    return hasAnyKeyword(text, ["UBND", "UY BAN NHAN DAN", "PHONG GIAO DUC", "PHONG GD", "SO GIAO DUC"]);
  });

  let unitLines = leftTop.filter((line) => {
    const text = normalizeText(line);
    return !agencyLines.some((agency) => normalizeText(agency) === text) &&
      !text.startsWith("SO:") &&
      !text.startsWith("SO ") &&
      !lineContainsDate(text);
  });

  if (agencyLines.length === 0 && leftTop.length > 0) {
    agencyLines = [leftTop[0]];
    unitLines = leftTop.slice(1);
  }

  return {
    agencyLines: uniqueCleanLines(agencyLines).map(normalizeHeaderLine),
    unitLines: uniqueCleanLines(unitLines).map(normalizeHeaderLine),
    numberLine: normalizeHeaderLine(numberLine),
    nationalLine: normalizeHeaderLine(nationalLine),
    mottoLine: normalizeHeaderLine(mottoLine),
    dateLine: normalizeHeaderLine(dateLine),
  };
}

function createCleanAdminParagraph(doc: Document, text: string, className: string): HTMLParagraphElement {
  const p = doc.createElement("p");
  p.className = className;
  p.textContent = text;
  return p;
}

function applyCleanAdminCellAttributes(cell: HTMLTableCellElement, side: "left" | "right") {
  cell.className = side === "left" ? "admin-left-cell" : "admin-right-cell";
  cell.setAttribute("valign", "top");
  cell.setAttribute("width", side === "left" ? "42%" : "58%");
  cell.setAttribute(
    "style",
    [
      `width:${side === "left" ? "42%" : "58%"}`,
      "border:none",
      "mso-border-alt:none",
      "background:transparent",
      "padding:0pt 4pt",
      "vertical-align:top",
      "text-align:center",
      "font-family:'Times New Roman',serif",
      "font-size:13pt",
      "line-height:115%",
    ].join(";")
  );
}

function createCleanAdministrativeHeaderTable(sourceTable: HTMLTableElement, showPreviewFrame: boolean): HTMLTableElement {
  const doc = sourceTable.ownerDocument;
  const data = extractAdministrativeHeaderData(sourceTable);

  const table = doc.createElement("table");
  table.className = `doc-table admin-header-table${showPreviewFrame ? " admin-header-preview-frame" : ""}`;
  table.setAttribute("cellspacing", "0");
  table.setAttribute("cellpadding", "0");
  table.setAttribute("border", "0");
  table.setAttribute("style", "width:100%;border-collapse:collapse;table-layout:fixed;border:none;background:transparent;margin:0 0 12pt 0;");

  const colgroup = doc.createElement("colgroup");
  const leftCol = doc.createElement("col");
  const rightCol = doc.createElement("col");
  leftCol.className = "admin-left-col";
  rightCol.className = "admin-right-col";
  leftCol.setAttribute("style", "width:42%;");
  rightCol.setAttribute("style", "width:58%;");
  colgroup.append(leftCol, rightCol);
  table.appendChild(colgroup);

  const firstRow = table.insertRow();
  const leftTopCell = firstRow.insertCell();
  const rightTopCell = firstRow.insertCell();
  applyCleanAdminCellAttributes(leftTopCell, "left");
  applyCleanAdminCellAttributes(rightTopCell, "right");

  const agencyLines = data.agencyLines.length ? data.agencyLines : [""];
  agencyLines.forEach((line) => leftTopCell.appendChild(createCleanAdminParagraph(doc, line, "admin-agency-line")));
  data.unitLines.forEach((line) => leftTopCell.appendChild(createCleanAdminParagraph(doc, line, "admin-unit-line")));

  rightTopCell.appendChild(createCleanAdminParagraph(doc, data.nationalLine, "admin-national-line"));
  rightTopCell.appendChild(createCleanAdminParagraph(doc, data.mottoLine, "admin-motto-line"));

  const secondRow = table.insertRow();
  const leftBottomCell = secondRow.insertCell();
  const rightBottomCell = secondRow.insertCell();
  applyCleanAdminCellAttributes(leftBottomCell, "left");
  applyCleanAdminCellAttributes(rightBottomCell, "right");

  leftBottomCell.appendChild(createCleanAdminParagraph(doc, data.numberLine, "admin-number-line"));
  rightBottomCell.appendChild(createCleanAdminParagraph(doc, data.dateLine, "admin-date-line"));

  return table;
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
      // Dựng lại riêng bảng thể thức hành chính bằng cấu trúc sạch.
      // Không giữ lại style/thuộc tính do Mammoth sinh ra vì Word dễ hiểu sai
      // thành font phóng to, nền xám, viền thật hoặc cột thụt thò.
      const cleanHeaderTable = createCleanAdministrativeHeaderTable(table, options.preserveFirstFrame);
      table.replaceWith(cleanHeaderTable);
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


function setInlineStyles(element: Element, styles: Record<string, string>) {
  const styleText = Object.entries(styles)
    .map(([key, value]) => `${key}: ${value}`)
    .join('; ');
  element.setAttribute('style', styleText);
}

function stripPreviewOnlyClasses(root: HTMLElement) {
  root.querySelectorAll('.admin-header-preview-frame').forEach((el) => {
    el.classList.remove('admin-header-preview-frame');
  });
}

function applyReliableWordHeaderStyles(root: HTMLElement) {
  stripPreviewOnlyClasses(root);

  const headerTables = Array.from(root.querySelectorAll('table.admin-header-table')) as HTMLTableElement[];

  headerTables.forEach((table) => {
    table.removeAttribute('border');
    table.removeAttribute('cellpadding');
    table.removeAttribute('cellspacing');
    table.setAttribute('cellspacing', '0');
    table.setAttribute('cellpadding', '0');

    setInlineStyles(table, {
      width: '100%',
      'border-collapse': 'collapse',
      'table-layout': 'fixed',
      border: 'none',
      'mso-border-alt': 'none',
      'margin-top': '0pt',
      'margin-bottom': '12pt',
      'font-family': 'Times New Roman, serif',
      'font-size': '13pt',
    });

    Array.from(table.rows).forEach((row) => {
      Array.from(row.cells).forEach((cell, cellIndex) => {
        const width = cellIndex === 0 ? '42%' : '58%';
        setInlineStyles(cell, {
          width,
          border: 'none',
          'mso-border-alt': 'none',
          padding: '0pt 3pt',
          'vertical-align': 'top',
          'text-align': 'center',
          'font-family': 'Times New Roman, serif',
          'font-size': '13pt',
          'line-height': '115%',
          'word-break': 'normal',
          'overflow-wrap': 'normal',
        });
      });
    });

    const paragraphs = Array.from(table.querySelectorAll('p')) as HTMLParagraphElement[];
    paragraphs.forEach((paragraph) => {
      const baseStyles: Record<string, string> = {
        margin: '0pt',
        padding: '0pt',
        'text-align': 'center',
        'text-indent': '0cm',
        'font-family': 'Times New Roman, serif',
        'font-size': '13pt',
        'line-height': '115%',
        'mso-line-height-rule': 'exactly',
        border: 'none',
        'word-break': 'normal',
        'overflow-wrap': 'normal',
      };

      if (paragraph.classList.contains('admin-agency-line')) {
        setInlineStyles(paragraph, {
          ...baseStyles,
          'font-weight': 'normal',
          'text-transform': 'uppercase',
        });
        return;
      }

      if (paragraph.classList.contains('admin-unit-line')) {
        setInlineStyles(paragraph, {
          ...baseStyles,
          'font-weight': 'bold',
          'text-transform': 'uppercase',
        });
        return;
      }

      if (paragraph.classList.contains('admin-number-line')) {
        setInlineStyles(paragraph, {
          ...baseStyles,
          'font-weight': 'normal',
          'white-space': 'nowrap',
        });
        return;
      }

      if (paragraph.classList.contains('admin-national-line')) {
        setInlineStyles(paragraph, {
          ...baseStyles,
          'font-weight': 'bold',
          'text-transform': 'uppercase',
          'white-space': 'nowrap',
          'letter-spacing': '-0.1pt',
          'font-size': '12.5pt',
        });
        return;
      }

      if (paragraph.classList.contains('admin-motto-line')) {
        setInlineStyles(paragraph, {
          ...baseStyles,
          'font-weight': 'bold',
          'white-space': 'nowrap',
          'font-size': '13pt',
          'text-decoration': 'underline',
        });
        return;
      }

      if (paragraph.classList.contains('admin-date-line')) {
        setInlineStyles(paragraph, {
          ...baseStyles,
          'font-style': 'italic',
          'white-space': 'nowrap',
        });
        return;
      }

      setInlineStyles(paragraph, baseStyles);
    });
  });
}

function applyReliableWordTitleStyles(root: HTMLElement, font: string, size: number) {
  root.querySelectorAll('.doc-main-title').forEach((el) => {
    setInlineStyles(el, {
      margin: '8pt 0pt 4pt 0pt',
      padding: '0pt',
      'text-align': 'center',
      'text-indent': '0cm',
      'font-family': `${font}, serif`,
      'font-size': `${Math.max(size, 14)}pt`,
      'font-weight': 'bold',
      'text-transform': 'uppercase',
      'line-height': '120%',
      'mso-line-height-rule': 'exactly',
    });
  });

  root.querySelectorAll('.doc-sub-title').forEach((el) => {
    setInlineStyles(el, {
      margin: '0pt 0pt 3pt 0pt',
      padding: '0pt',
      'text-align': 'center',
      'text-indent': '0cm',
      'font-family': `${font}, serif`,
      'font-size': `${size}pt`,
      'font-weight': 'bold',
      'line-height': '120%',
      'mso-line-height-rule': 'exactly',
    });
  });
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

  // Khi xuất Word, không dùng lại style trình duyệt vì Word rất dễ hiểu sai
  // kích thước chữ, viền bảng và khoảng cách. Chỉ giữ style ảnh; còn lại
  // được chuẩn hóa lại bằng CSS/inline style phía dưới.
  cloneDiv.querySelectorAll("*").forEach((el) => {
    const tag = el.tagName.toLowerCase();
    Array.from(el.attributes).forEach((attr) => {
      const name = attr.name.toLowerCase();
      if (name === "class") return;
      if (tag === "img" && ["src", "alt", "title"].includes(name)) return;
      el.removeAttribute(attr.name);
    });
  });

  applyReliableWordHeaderStyles(cloneDiv);
  applyReliableWordTitleStyles(cloneDiv, font, size);

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
              background: transparent !important;
              background-color: transparent !important;
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
              width: 100% !important;
              border-collapse: collapse !important;
              font-family: "Times New Roman", serif !important;
              font-size: 13pt !important;
          }
          .admin-header-table col.admin-left-col { width: 42% !important; }
          .admin-header-table col.admin-right-col { width: 58% !important; }
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
              letter-spacing: -0.25pt !important;
              font-size: 12pt !important;
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
