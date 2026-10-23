/**
 * pdfStructure.ts — Lõi phân tích cấu trúc PDF (thuần logic, không phụ thuộc DOM/pdf.js,
 * chạy được cả trên trình duyệt lẫn Node để kiểm thử).
 *
 * Đầu vào: danh sách "mẩu chữ" (text item) kèm tọa độ của từng trang, lấy từ pdf.js
 * (PDF có lớp chữ) hoặc từ OCR (PDF scan). Đầu ra: mô hình khối tài liệu (DocBlock[])
 * dùng chung cho: (1) dựng HTML xem trước, (2) xuất Word .docx chuẩn NĐ 30.
 *
 * Hệ tọa độ: gốc ở GÓC TRÊN-TRÁI của trang, đơn vị point (1/72 inch), y tăng dần xuống dưới
 * — giống hệ tọa độ màn hình. Tầng trích xuất (services/pdf.ts) chịu trách nhiệm quy đổi.
 */

export interface PdfTextItem {
  text: string;
  /** Tọa độ góc trái của mẩu chữ (point, gốc trên-trái). */
  x: number;
  y: number;
  width: number;
  height: number;
  fontSize: number;
  bold: boolean;
  italic: boolean;
  fontName?: string;
  /** Độ tin cậy OCR 0–100 (chỉ trang OCR mới có). */
  conf?: number;
  /** Số thứ tự DÒNG do Tesseract phân đoạn (chỉ trang OCR) — tầng cấu trúc
   * dùng nguyên, không gom dòng lại theo tọa độ. */
  lineId?: number;
  /** Số thứ tự ĐOẠN do Tesseract phân đoạn (chỉ trang OCR). */
  paraId?: number;
}

export interface PdfPageInput {
  pageNumber: number;
  /** Kích thước trang (point). A4 ≈ 595 x 842. */
  width: number;
  height: number;
  items: PdfTextItem[];
  /** Trang này lấy chữ bằng lớp text của PDF hay bằng OCR. */
  source: "text" | "ocr";
  /** Ảnh nhúng trích từ trang (data URL) — không gồm ảnh toàn trang của trang scan. */
  imageDataUrls?: string[];
  /** Kích thước thật của ảnh tương ứng imageDataUrls (để giữ tỉ lệ khi vào Word). */
  imageSizes?: Array<{ width: number; height: number }>;
}

export interface DocRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
}

export type BlockRole =
  | "body"
  | "title"
  | "subtitle"
  | "heading"
  | "recipient"
  | "signatureText";

export type DocBlock =
  | {
      kind: "paragraph";
      runs: DocRun[];
      align: "left" | "center" | "right" | "justify";
      role: BlockRole;
      /** Cỡ chữ nổi trội của đoạn (point) — giữ lại từ PDF để tham khảo. */
      fontSize?: number;
      firstLineIndent?: boolean;
    }
  | {
      kind: "table";
      /** rows[hàng][ô][dòng trong ô][run] — mỗi ô là danh sách dòng, mỗi dòng là danh sách run. */
      rows: DocRun[][][][];
      role: "data" | "adminHeader" | "signature";
      /** Tỉ lệ bề rộng các cột (%) nếu suy ra được từ PDF. */
      colWidths?: number[];
    }
  | { kind: "image"; dataUrl: string; width: number; height: number };

export interface PdfClassification {
  kind: "text" | "scanned" | "mixed" | "empty";
  totalPages: number;
  textPages: number;
  ocrPages: number;
  totalChars: number;
  label: string;
}

/* ------------------------------------------------------------------ */
/* Tiện ích chữ                                                        */
/* ------------------------------------------------------------------ */

export function normalizeViText(value: string): string {
  return (value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();
}

export function cleanPdfText(value: string): string {
  return (value || "")
    // Chuẩn hoá Unicode sang NFC (dạng dựng sẵn): chữ trích từ PDF thường ở dạng
    // tổ hợp (chữ cái + dấu rời) — để nguyên vào Word dễ bị coi là "lỗi ký tự",
    // dấu hiển thị lệch vị trí. NFC là dạng Word/phần mềm Việt Nam dùng phổ biến.
    .normalize("NFC")
    .replace(/\u00a0/g, " ")
    .replace(/\u200b/g, "")
    .replace(/[ \t]+/g, " ")
    .replace(/\s+([,.;:!?])/g, "$1")
    .replace(/\(\s+/g, "(")
    .replace(/\s+\)/g, ")")
    .trim();
}

/**
 * Đánh giá lớp chữ của một trang có đáng tin không. PDF lỗi font/mã hoá (font
 * không nhúng đúng, bảng mã riêng kiểu TCVN/VNI cũ, thiếu ToUnicode) vẫn có
 * "chữ" nhưng là ký tự rác: vùng Private Use, ký tự thay thế �, ký tự điều
 * khiển — hoặc chữ bị tách rời từng ký tự một. Trang như vậy phải OCR lại
 * thay vì tin lớp chữ, nếu không file Word xuất ra sẽ lỗi ký tự hàng loạt.
 */
export function pageTextLooksGarbled(items: PdfTextItem[]): { garbled: boolean; reason: string } {
  // CHỈ xét các mẩu chữ có nội dung, nối bằng 1 khoảng trắng: pdf.js chèn rất
  // nhiều item chỉ gồm khoảng trắng để tách các mẩu chữ ở xa nhau. Nếu tính
  // cả chúng, token bị vỡ vụn giả tạo — một bảng điểm bình thường (hầu hết ô
  // chỉ có 1 ký tự: Đ, T, số...) sẽ bị tưởng là "chữ vỡ từng ký tự" và bị
  // OCR oan. (Lỗi thật 21/10/2026: bảng tổng hợp kết quả đánh giá in từ hệ
  // thống trường học, lớp chữ hoàn hảo nhưng app OCR ra rác.)
  const meaningful = items.filter((it) => it.text.trim().length > 0);
  const full = meaningful.map((it) => it.text).join(" ");
  const compact = full.replace(/\s/g, "");
  const chars = compact.length;
  if (chars < 80) return { garbled: false, reason: "" };

  // Dạng 1: ký tự vùng riêng/thay thế/điều khiển — font lỗi mã hoá trầm trọng.
  const suspicious = (full.match(/[\uE000-\uF8FF\uFFFD\u0000-\u0008\u000B-\u001F\u007F]/g) || []).length;
  if (suspicious / chars > 0.02) {
    return { garbled: true, reason: `khoảng ${Math.round((suspicious / chars) * 100)}% ký tự lạ (lỗi font/mã hoá)` };
  }

  // Tỉ lệ chữ cái còn dấu tiếng Việt — bằng chứng phụ cho dạng 2 và 3:
  // font lỗi thật làm dấu biến mất gần hết, còn văn bản xịn (kể cả bảng
  // điểm) luôn giữ dấu ở tiêu đề, họ tên, ghi chú.
  const letters = (compact.match(/[A-Za-zÀ-ỹ]/g) || []).length;
  const marked =
    (compact.match(/[ăâđêôơưĂÂĐÊÔƠƯ\u00C0-\u00FF\u1EA0-\u1EF9]/g) || []).length +
    (compact.match(/[\u0300-\u036F]/g) || []).length;
  const markedRatio = letters > 0 ? marked / letters : 0;

  const tokens = full.split(/\s+/).filter(Boolean);
  // Dạng 2: chữ bị tách rời từng ký tự (lỗi giãn chữ trong PDF). Chỉ kết
  // luận khi KÈM bằng chứng mất dấu: bảng biểu nhiều ô 1 ký tự nhưng còn
  // nguyên dấu tiếng Việt thì KHÔNG phải lỗi font.
  if (tokens.length > 120 && letters > 300 && markedRatio < 0.03) {
    const singleChar = tokens.filter((t) => t.replace(/[.,;:]/g, "").length === 1).length;
    if (singleChar / tokens.length > 0.45) {
      return { garbled: true, reason: "chữ bị tách rời từng ký tự và mất gần hết dấu tiếng Việt (lỗi giãn chữ/font trong PDF)" };
    }
  }

  // Dạng 3: font nhúng SAI ÁNH XẠ (hay gặp ở PDF in từ hệ thống quản lý
  // trường học): ký tự vẫn là Unicode "hợp lệ" nên lọt hai lưới trên, nhưng
  // dấu tiếng Việt biến thành ký hiệu lạ (| £ « » ~ { }...) hoặc mất hẳn,
  // các ô bảng dính thành cụm ký tự vô nghĩa. Ba tín hiệu độc lập:
  const pipes = (compact.match(/\|/g) || []).length;
  if (pipes >= 8 && pipes / chars > 0.01) {
    return { garbled: true, reason: "xuất hiện dày đặc ký tự lạ dạng '|' (lỗi font nhúng, sai ánh xạ ký tự)" };
  }
  const hardSymbols = (compact.match(/[£¤¦¨¯´¸«»~^`{[\]\\<>@#§¶×]/g) || []).length;
  if (hardSymbols / chars > 0.05) {
    return { garbled: true, reason: `khoảng ${Math.round((hardSymbols / chars) * 100)}% ký hiệu lạ (lỗi font nhúng, sai ánh xạ ký tự)` };
  }
  // Cụm dài chứa nhiều ký hiệu (ô bảng bị dính, vd "|291120559).|H|"):
  if (tokens.length > 60) {
    const glued = tokens.filter(
      (t) => t.length >= 8 && (t.match(/[|£«»~^`{[\]\\<>#@]/g) || []).length >= 4
    ).length;
    if (glued / tokens.length > 0.03) {
      return { garbled: true, reason: "nhiều cụm ký tự vô nghĩa dính liền nhau (lỗi font nhúng, sai ánh xạ ký tự)" };
    }
  }
  // Trang nhiều chữ cái mà gần như KHÔNG có dấu tiếng Việt nào, đồng thời
  // mật độ ký hiệu cao → dấu đã bị ánh xạ sai gần hết.
  if (letters > 300 && markedRatio < 0.03) {
    const commonPunct = ".,;:!?()'\"/%°-–—";
    const symbolChars = [...compact].filter(
      (c) => !/[A-Za-zÀ-ỹ0-9]/.test(c) && !commonPunct.includes(c)
    ).length;
    if (symbolChars / chars > 0.06) {
      return { garbled: true, reason: "gần như không còn dấu tiếng Việt nhưng dày đặc ký hiệu lạ (lỗi font nhúng, sai ánh xạ ký tự)" };
    }
  }
  return { garbled: false, reason: "" };
}

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** Ánh xạ 1 item thô của pdf.js (getTextContent) sang PdfTextItem. Test được độc lập. */
export function mapPdfJsTextItem(raw: {
  str: string;
  transform: number[];
  width: number;
  height: number;
  fontName?: string;
}, pageHeight: number): PdfTextItem | null {
  const text = raw.str ?? "";
  if (!text) return null;
  const t = raw.transform || [1, 0, 0, 1, 0, 0];
  const fontSize = Math.hypot(t[2] || 0, t[3] || 0) || raw.height || 12;
  const fontName = raw.fontName || "";
  return {
    text,
    x: t[4] || 0,
    // pdf.js: gốc dưới-trái, y là chân chữ → đổi sang gốc trên-trái.
    y: pageHeight - (t[5] || 0) - fontSize,
    width: raw.width || 0,
    height: raw.height || fontSize,
    fontSize,
    bold: /bold|black|heavy|semibold/i.test(fontName),
    italic: /italic|oblique/i.test(fontName),
    fontName,
  };
}

/* ------------------------------------------------------------------ */
/* Gộp mẩu chữ → dòng                                                  */
/* ------------------------------------------------------------------ */

export interface LineSegment {
  runs: DocRun[];
  xStart: number;
  xEnd: number;
  text: string;
}

export interface TextLine {
  segments: LineSegment[];
  runs: DocRun[];
  text: string;
  xStart: number;
  xEnd: number;
  y: number;
  fontSize: number;
  bold: boolean;
  pageNumber: number;
  /** Trang OCR: dòng này mở đầu một đoạn mới theo phân đoạn của Tesseract. */
  paraBreakBefore?: boolean;
}

interface RawLine {
  items: PdfTextItem[];
  y: number;
}

function itemsToRawLines(items: PdfTextItem[]): RawLine[] {
  const sorted = [...items]
    // Bỏ mẩu chữ chỉ chứa khoảng trắng: pdf.js tự chèn chúng để lấp khoảng trống
    // GIỮA HAI CỘT (bề rộng có thể tới cả trăm point) — giữ lại sẽ làm 2 cột dính
    // thành 1 segment. Khoảng cách hình học giữa các mẩu chữ thật đã đủ để tách cột.
    .filter((it) => it.text && it.text.trim().length > 0)
    .sort((a, b) => a.y - b.y || a.x - b.x);
  const lines: RawLine[] = [];
  for (const item of sorted) {
    const last = lines[lines.length - 1];
    const tolerance = Math.max(2.5, item.fontSize * 0.4);
    if (last && Math.abs(item.y - last.y) <= tolerance) {
      last.items.push(item);
    } else {
      lines.push({ items: [item], y: item.y });
    }
  }
  return lines;
}

/** Ngưỡng khoảng trống (pt) để coi là ranh giới cột/ô trên cùng một dòng. */
function columnGapThreshold(fontSize: number): number {
  return Math.max(11, fontSize * 1.05);
}

function buildLine(raw: RawLine, pageNumber: number): TextLine {
  const items = [...raw.items].sort((a, b) => a.x - b.x);
  const segments: LineSegment[] = [];
  let current: { items: PdfTextItem[] } | null = null;
  let prevEnd = 0;
  for (const item of items) {
    if (!current) {
      current = { items: [item] };
      segments.push({ runs: [], xStart: item.x, xEnd: item.x + item.width, text: "" });
    } else {
      const gap = item.x - prevEnd;
      // Khoảng trống lớn bất thường trên cùng dòng = ranh giới cột/ô bảng.
      if (gap > columnGapThreshold(item.fontSize)) {
        current = { items: [item] };
        segments.push({ runs: [], xStart: item.x, xEnd: item.x + item.width, text: "" });
      } else {
        current.items.push(item);
        const seg = segments[segments.length - 1];
        seg.xEnd = item.x + item.width;
        // Chèn khoảng trắng khi 2 mẩu chữ tách rời nhau vừa phải. Ngưỡng phải
        // đủ nhỏ (0,08 lần cỡ chữ): trong văn bản căn đều, pdf.js thường tách
        // mỗi từ thành một mẩu và khoảng trống giữa 2 từ chỉ ~1-2pt — ngưỡng
        // cao hơn sẽ làm chữ dính liền ("quảđánh"). Chỗ tách giữa chừng của
        // một từ (đổi font/kerning) có khoảng trống ≈ 0 nên không bị ảnh hưởng.
        if (gap > item.fontSize * 0.08) {
          const lastItem = current.items[current.items.length - 2];
          if (lastItem && !lastItem.text.endsWith(" ") && !item.text.startsWith(" ")) {
            lastItem.text += " ";
          }
        }
      }
    }
    prevEnd = item.x + item.width;
  }

  // Dựng runs cho từng segment
  const segItems: PdfTextItem[][] = [];
  {
    let idx = 0;
    let prevEnd2 = 0;
    let bucket: PdfTextItem[] = [];
    for (const item of items) {
      if (bucket.length > 0 && item.x - prevEnd2 > columnGapThreshold(item.fontSize)) {
        segItems.push(bucket);
        bucket = [];
        idx++;
      }
      bucket.push(item);
      prevEnd2 = item.x + item.width;
    }
    if (bucket.length) segItems.push(bucket);
    void idx;
  }
  segItems.forEach((bucket, i) => {
    const seg = segments[i];
    if (!seg) return;
    seg.runs = mergeRuns(bucket);
    seg.text = cleanPdfText(bucket.map((b) => b.text).join(""));
    seg.xStart = bucket[0].x;
    seg.xEnd = bucket[bucket.length - 1].x + bucket[bucket.length - 1].width;
  });

  // Runs của dòng = runs của mọi segment nối tiếp nhau. BẮT BUỘC chèn một dấu
  // cách ở ranh giới segment (cột/ô khác nhau), nếu không chữ sẽ dính liền kiểu
  // "STTHọ và tên" khi dòng được dùng như đoạn văn (đây là lỗi "lỗi ký tự" mà
  // người dùng nhìn thấy khi mở file Word). Sao chép run để không sửa vào
  // segment gốc (segment còn dùng làm ô bảng).
  const allRuns: DocRun[] = [];
  for (const seg of segments) {
    for (const r of seg.runs) {
      const copy = { ...r };
      const prev = allRuns[allRuns.length - 1];
      if (prev && !/\s$/.test(prev.text) && !/^\s/.test(copy.text)) {
        prev.text += " ";
      }
      allRuns.push(copy);
    }
  }
  const text = cleanPdfText(segments.map((s) => s.text).join(" "));
  const fontSize = items.reduce((m, it) => Math.max(m, it.fontSize), 0) || 12;
  const boldChars = items.filter((it) => it.bold).reduce((n, it) => n + it.text.length, 0);
  return {
    segments,
    runs: allRuns,
    text,
    xStart: items[0]?.x ?? 0,
    xEnd: Math.max(...items.map((it) => it.x + it.width), 0),
    y: raw.y,
    fontSize,
    bold: text.length > 0 && boldChars >= text.length * 0.6,
    pageNumber,
  };
}

function mergeRuns(items: PdfTextItem[]): DocRun[] {
  const runs: DocRun[] = [];
  for (const item of items) {
    // NFC hoá ngay ở mức run để mọi ngõ ra (Word/HTML/text) đều dùng chữ dựng sẵn.
    const text = (item.text || "").normalize("NFC");
    if (!text) continue;
    const last = runs[runs.length - 1];
    if (last && !!last.bold === !!item.bold && !!last.italic === !!item.italic) {
      last.text += text;
    } else {
      runs.push({ text, bold: item.bold || undefined, italic: item.italic || undefined });
    }
  }
  // Rà soát khoảng trắng ở ranh giới run
  runs.forEach((r) => (r.text = r.text.replace(/[ \t]+/g, " ")));
  return runs.filter((r) => r.text.length > 0);
}

export function pageToLines(page: PdfPageInput): TextLine[] {
  return itemsToRawLines(page.items).map((raw) => buildLine(raw, page.pageNumber));
}

/* ------------------------------------------------------------------ */
/* Trang OCR: dựng dòng theo phân đoạn của Tesseract                   */
/* ------------------------------------------------------------------ */

/**
 * Ngưỡng tách segment cho trang OCR (× cỡ chữ): chỉ khe THẬT RỘNG mới tách
 * (khe giữa 2 cột đầu trang, khe cột của bảng thật). Khe giữa các từ trong
 * đoạn căn đều của bản scan không bao giờ tới mức này. Dùng ngưỡng của
 * trang có lớp chữ (~1,05×) cho trang scan đã từng xé đoạn văn thành nhiều
 * segment → máy dựng bảng tưởng nhầm là bảng dữ liệu → chữ bị xáo trộn
 * thứ tự đọc và bị nuốt mất ở khâu gán cột. Đó là "vấn đề lớn" của cơ chế
 * quét cũ: Tesseract đọc đúng ~89% nhưng bản Word ra lò không dùng được.
 */
const OCR_SEGMENT_GAP_RATIO = 2.4;

function buildOcrLine(items: PdfTextItem[], pageNumber: number, paraBreakBefore: boolean): TextLine {
  const sorted = [...items].sort((a, b) => a.x - b.x);
  const fontSize = sorted.reduce((m, it) => Math.max(m, it.fontSize), 0) || 12;
  const buckets: PdfTextItem[][] = [];
  let bucket: PdfTextItem[] = [];
  let prevEnd = 0;
  for (const item of sorted) {
    if (bucket.length > 0 && item.x - prevEnd > fontSize * OCR_SEGMENT_GAP_RATIO) {
      buckets.push(bucket);
      bucket = [];
    }
    bucket.push(item);
    prevEnd = item.x + item.width;
  }
  if (bucket.length) buckets.push(bucket);
  const segments: LineSegment[] = buckets.map((b) => ({
    runs: mergeRuns(b),
    xStart: b[0].x,
    xEnd: b[b.length - 1].x + b[b.length - 1].width,
    text: cleanPdfText(b.map((x) => x.text).join("")),
  }));
  const allRuns: DocRun[] = [];
  for (const seg of segments) {
    for (const r of seg.runs) {
      const copy = { ...r };
      const prev = allRuns[allRuns.length - 1];
      if (prev && !/\s$/.test(prev.text) && !/^\s/.test(copy.text)) prev.text += " ";
      allRuns.push(copy);
    }
  }
  return {
    segments,
    runs: allRuns,
    text: cleanPdfText(segments.map((s) => s.text).join(" ")),
    xStart: sorted[0]?.x ?? 0,
    xEnd: Math.max(...sorted.map((it) => it.x + it.width), 0),
    y: Math.min(...sorted.map((it) => it.y)),
    fontSize,
    bold: false,
    pageNumber,
    paraBreakBefore,
  };
}

/**
 * Dựng dòng cho trang OCR từ lineId/paraId mà Tesseract đã phân đoạn:
 * giữ nguyên dòng và ranh giới đoạn của Tesseract (nó phân đoạn trang scan
 * tốt hơn nhiều so với gom lại theo tọa độ y — hộp từ OCR rung vài point).
 * Các dòng được sắp lại theo y cho đúng thứ tự đọc của trang 1 cột.
 */
/** Dựng lại runs/text của một dòng OCR khi chỉ giữ một phần segment của nó. */
function rebuildLineSegments(line: TextLine, segs: LineSegment[]): TextLine {
  const allRuns: DocRun[] = [];
  for (const seg of segs) {
    for (const r of seg.runs) {
      const copy = { ...r };
      const prev = allRuns[allRuns.length - 1];
      if (prev && !/\s$/.test(prev.text) && !/^\s/.test(copy.text)) prev.text += " ";
      allRuns.push(copy);
    }
  }
  return {
    ...line,
    segments: segs,
    runs: allRuns,
    text: cleanPdfText(segs.map((s) => s.text).join(" ")),
    xEnd: Math.max(...segs.map((s) => s.xEnd)),
  };
}

/**
 * Vùng cuối trang kiểu hành chính VN: cột trái "Nơi nhận", cột phải khối
 * chữ ký — Tesseract hay gộp 2 cột thành một dòng/đoạn, chữ ký chen vào
 * giữa danh sách nơi nhận. Từ dòng "Nơi nhận" trở xuống: tách segment bên
 * phải (x ≥ 55% bề rộng trang) ra thành một nhóm dòng riêng đặt SAU toàn
 * bộ cột trái — đúng thứ tự đọc của văn bản gốc.
 */
function partitionRecipientZone(lines: TextLine[], pageWidth: number): TextLine[] {
  const markerIdx = lines.findIndex((l) => normalizeViText(l.text).startsWith("NOI NHAN"));
  if (markerIdx < 0) return lines;
  const cutX = pageWidth * 0.55;
  const kept: TextLine[] = [];
  const rightLines: TextLine[] = [];
  lines.forEach((l, idx) => {
    if (idx < markerIdx) {
      kept.push(l);
      return;
    }
    const leftSegs = l.segments.filter((s) => s.xStart < cutX);
    const rightSegs = l.segments.filter((s) => s.xStart >= cutX);
    if (rightSegs.length === 0) {
      kept.push(l);
      return;
    }
    if (leftSegs.length > 0) kept.push(rebuildLineSegments(l, leftSegs));
    for (const seg of rightSegs) {
      rightLines.push({
        segments: [seg],
        runs: seg.runs,
        text: seg.text,
        xStart: seg.xStart,
        xEnd: seg.xEnd,
        y: l.y + rightLines.length * 0.01,
        fontSize: l.fontSize,
        bold: false,
        pageNumber: l.pageNumber,
        paraBreakBefore: rightLines.length === 0,
      });
    }
  });
  return rightLines.length > 0 ? [...kept, ...rightLines] : lines;
}

export function ocrPageToLines(page: PdfPageInput): TextLine[] {
  const withIds = page.items.filter(
    (it) => it.text && it.text.trim().length > 0 && typeof it.lineId === "number"
  );
  if (withIds.length === 0) {
    // Dự phòng (dữ liệu OCR không gắn lineId): gom dòng theo y như cũ nhưng
    // mỗi dòng chỉ 1 segment — vẫn tránh được máy tách cột nuốt chữ.
    return itemsToRawLines(page.items).map((raw) => buildOcrLine(raw.items, page.pageNumber, false));
  }
  const byLine = new Map<number, PdfTextItem[]>();
  for (const it of withIds) {
    const arr = byLine.get(it.lineId!) || [];
    arr.push(it);
    byLine.set(it.lineId!, arr);
  }
  const groups = [...byLine.values()];
  groups.sort((a, b) => Math.min(...a.map((i) => i.y)) - Math.min(...b.map((i) => i.y)));

  // Tách dòng bị Tesseract GỘP 2 cột ở đầu trang: nếu cụm quốc hiệu
  // "CỘNG HÒA…" bắt đầu ở giữa dòng (trước nó là khe rộng), cắt dòng tại
  // từ đó — nửa trái là tên cơ quan ban hành, nửa phải là quốc hiệu. Không
  // tách, extractAdminHeader sẽ nuốt mất tên cơ quan vào ô quốc hiệu (đã
  // gặp thật trên CV 5095: ô trái chỉ còn mảnh rác "Ì").
  const splitGroups: PdfTextItem[][] = [];
  for (const g of groups) {
    const minY = Math.min(...g.map((i) => i.y));
    const sorted = [...g].sort((a, b) => a.x - b.x);
    const fontSize = sorted.reduce((m, it) => Math.max(m, it.fontSize), 0) || 12;
    let cut = -1;
    if (minY < page.height * 0.25) {
      for (let k = 1; k + 1 < sorted.length; k++) {
        if (normalizeViText(sorted[k].text) === "CONG" && normalizeViText(sorted[k + 1].text) === "HOA") {
          const gap = sorted[k].x - (sorted[k - 1].x + sorted[k - 1].width);
          if (gap > fontSize * 1.0) {
            cut = k;
            break;
          }
        }
      }
    }
    if (cut > 0) splitGroups.push(sorted.slice(0, cut), sorted.slice(cut));
    else splitGroups.push(g);
  }

  let prevPara: number | undefined;
  const built = splitGroups.map((g) => {
    const para = g[0].paraId;
    const brk = prevPara !== undefined && para !== undefined && para !== prevPara;
    if (para !== undefined) prevPara = para;
    return buildOcrLine(g, page.pageNumber, brk);
  });
  return partitionRecipientZone(built, page.width);
}

/* ------------------------------------------------------------------ */
/* Nhận diện các thành phần thể thức                                   */
/* ------------------------------------------------------------------ */

const NATIONAL_RE = /CONG HOA XA HOI CHU NGHIA VIET NAM/;
const MOTTO_RE = /DOC LAP.*TU DO.*HANH PHUC/;
const NUMBER_RE = /^SO\s*[:.]/;
const DATE_RE = /NGAY\s+\d{1,2}\s+THANG\s+\d{1,2}\s+NAM\s+\d{4}/;
const AGENCY_KEYWORDS = ["UBND", "UY BAN NHAN DAN", "PHONG", "SO GIAO DUC", "SO GD", "TRUONG", "BAN", "HUY", "XA", "TINH"];
const DOC_TITLES = [
  "KE HOACH", "QUYET DINH", "CONG VAN", "TO TRINH", "BAO CAO", "BIEN BAN",
  "THONG BAO", "GIAY MOI", "DE NGHI", "DON", "PHIEU", "DANH SACH",
  "CHUONG TRINH", "QUY CHE", "HUONG DAN", "NGHI QUYET", "CHI THI",
];
// LƯU Ý: không để từ khoá trần "TRUONG" — sẽ khớp nhầm "TRƯỜNG THPT/Tiểu học"
// ở đầu trang của mọi văn bản trường học và biến 2 dòng tiêu đề thành "bảng
// chữ ký". Chỉ dùng các cụm chức danh ký đầy đủ.
const SIGNATURE_KEYWORDS = ["NOI NHAN", "TM.", "KT.", "HIEU TRUONG", "TRUONG PHONG", "PHO HIEU TRUONG", "NGUOI KY", "CHU TICH", "GIAM DOC", "THU TRUONG"];
const HEADING_RE = /^(PHAN|CHUONG|MUC|BAI)\s+[IVXLCDM0-9]+\b|^[IVXLCDM]{1,5}[.)]\s+\S|^\d{1,2}[.)]\s+\S|^[a-zđ][.)]\s+\S/i;

function isCentered(line: TextLine, pageWidth: number): boolean {
  const mid = (line.xStart + line.xEnd) / 2;
  return Math.abs(mid - pageWidth / 2) < pageWidth * 0.09 && line.xStart > pageWidth * 0.16;
}

function isAllCaps(text: string): boolean {
  // PHẢI kiểm tra trên text gốc: normalizeViText() đã tự viết hoa toàn bộ.
  const letters = text.replace(/[^A-Za-zÀ-ỹĐđ]/g, "");
  if (letters.length < 3) return false;
  return text === text.toUpperCase() && text !== text.toLowerCase();
}

/* ------------------------------------------------------------------ */
/* Phần đầu hành chính (ô số 1 & 2 theo NĐ 30)                         */
/* ------------------------------------------------------------------ */

interface HeaderResult {
  block: DocBlock;
  consumedLines: number;
}

/**
 * Tìm phần đầu thể thức ở đầu trang 1: cột trái = cơ quan chủ quản + cơ quan
 * ban hành + số ký hiệu; cột phải = quốc hiệu + tiêu ngữ + địa danh/ngày tháng.
 * Trả về null nếu không thấy quốc hiệu lẫn cơ quan ban hành.
 */
function extractAdminHeader(lines: TextLine[], pageWidth: number): HeaderResult | null {
  const zone = lines.slice(0, 14);
  const nationalIdx = zone.findIndex((l) => NATIONAL_RE.test(normalizeViText(l.text)));
  const hasAgency = zone.some((l) =>
    l.segments.some((s) => s.xStart < pageWidth * 0.5 && AGENCY_KEYWORDS.some((k) => normalizeViText(s.text).includes(k)))
  );
  if (nationalIdx < 0 && !hasAgency) return null;

  // Vùng header kết thúc ở dòng số/ngày cuối cùng trước tiêu đề văn bản.
  let endIdx = zone.length - 1;
  for (let i = 0; i < zone.length; i++) {
    const norm = normalizeViText(zone[i].text);
    if (
      i > 0 &&
      DOC_TITLES.some((t) => norm === t || norm.startsWith(t + " ")) &&
      isCentered(zone[i], pageWidth)
    ) {
      endIdx = i - 1;
      break;
    }
    if (i > 0 && DATE_RE.test(norm)) endIdx = Math.max(endIdx === zone.length - 1 ? i : endIdx, i);
  }
  // Thu hẹp: lấy tới dòng cuối cùng còn mang tính header trong 8 dòng đầu.
  const headerLines: TextLine[] = [];
  for (let i = 0; i < zone.length && i <= Math.max(endIdx, nationalIdx); i++) {
    const l = zone[i];
    if (i > 0 && headerLines.length > 0) {
      const prev = headerLines[headerLines.length - 1];
      if (l.y - prev.y > l.fontSize * 3.2) break; // khoảng trống lớn = hết header
    }
    headerLines.push(l);
  }
  if (headerLines.length === 0) return null;

  const midX = pageWidth * 0.47;
  const leftLines: string[] = [];
  const rightLines: string[] = [];
  let numberLine = "";
  let dateLine = "";
  let nationalLine = "";
  let mottoLine = "";

  for (const line of headerLines) {
    for (const seg of line.segments) {
      const segText = cleanPdfText(seg.text);
      if (!segText) continue;
      const norm = normalizeViText(segText);
      const isLeft = seg.xStart < midX;
      if (NATIONAL_RE.test(norm)) {
        nationalLine = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";
      } else if (MOTTO_RE.test(norm)) {
        mottoLine = "Độc lập - Tự do - Hạnh phúc";
      } else if (NUMBER_RE.test(norm)) {
        numberLine = segText;
      } else if (DATE_RE.test(norm)) {
        dateLine = segText;
      } else if (isLeft) {
        leftLines.push(segText);
      } else if (!numberLine && seg.xStart >= midX && /NGAY/.test(norm)) {
        dateLine = segText;
      } else {
        rightLines.push(segText);
      }
    }
  }

  if (!nationalLine && !leftLines.length) return null;
  if (!mottoLine && nationalLine) mottoLine = "Độc lập - Tự do - Hạnh phúc";
  if (!dateLine && rightLines.length) {
    const idx = rightLines.findIndex((t) => /ngày/i.test(t));
    if (idx >= 0) dateLine = rightLines.splice(idx, 1)[0];
  }

  const toRuns = (text: string, opts: Partial<DocRun> = {}): DocRun[] => [
    { text, ...opts },
  ];

  const leftCellLines: DocRun[][] = leftLines.map((t, i) =>
    toRuns(t.toUpperCase(), { bold: i === leftLines.length - 1 && leftLines.length > 1 ? true : undefined })
  );
  const rightCellLines: DocRun[][] = [];
  if (nationalLine) rightCellLines.push(toRuns(nationalLine, { bold: true }));
  if (mottoLine) rightCellLines.push(toRuns(mottoLine, { bold: true, underline: true }));
  rightLines.forEach((t) => rightCellLines.push(toRuns(t)));

  const rows: DocRun[][][][] = [
    [leftCellLines, rightCellLines],
    [
      numberLine ? [toRuns(numberLine)] : [],
      dateLine ? [toRuns(dateLine, { italic: true })] : [],
    ],
  ];

  return {
    block: { kind: "table", rows, role: "adminHeader", colWidths: [36, 64] },
    consumedLines: headerLines.length,
  };
}

/* ------------------------------------------------------------------ */
/* Bảng dữ liệu & khối chữ ký                                          */
/* ------------------------------------------------------------------ */

function lineIsTableRow(line: TextLine): boolean {
  return line.segments.length >= 2 && line.segments.every((s) => s.text.length > 0);
}

interface ColumnRange {
  xStart: number;
  xEnd: number;
}

/**
 * Dựng mô hình cột từ các dòng ĐẦY ĐỦ NHẤT (nhiều segment nhất) của vùng bảng:
 * lấy trung vị xStart/xEnd theo từng vị trí. Các dòng đầy đủ nhất luôn có cùng
 * số segment và cùng thứ tự cột nên ghép theo vị trí là an toàn; các dòng
 * thiếu ô (ô trống, ô bị cuốn dòng) sẽ được gán cột theo độ chồng lấn trục x.
 */
function buildColumns(rows: TextLine[]): ColumnRange[] {
  const maxSegs = Math.max(...rows.map((l) => l.segments.length));
  const fullest = rows.filter((l) => l.segments.length === maxSegs);
  const median = (vals: number[]) => vals.sort((a, b) => a - b)[Math.floor(vals.length / 2)];
  const cols: ColumnRange[] = [];
  for (let c = 0; c < maxSegs; c++) {
    cols.push({
      xStart: median(fullest.map((l) => l.segments[c].xStart)),
      xEnd: median(fullest.map((l) => l.segments[c].xEnd)),
    });
  }
  return cols;
}

/** Gán một segment vào cột: ưu tiên cột ĐẦU TIÊN mà nó chồng lấn đáng kể (ô
 * gộp ngang trong hàng tiêu đề), rồi tới cột chồng lấn nhiều nhất, cuối cùng
 * là cột có tâm gần nhất (sai lệch ≤ 22pt). -1 = không thuộc cột nào. */
function assignColumn(seg: LineSegment, cols: ColumnRange[]): number {
  const segWidth = Math.max(1, seg.xEnd - seg.xStart);
  let best = -1;
  let bestOverlap = 0;
  for (let i = 0; i < cols.length; i++) {
    const overlap = Math.min(seg.xEnd, cols[i].xEnd) - Math.max(seg.xStart, cols[i].xStart);
    if (overlap > bestOverlap) {
      bestOverlap = overlap;
      best = i;
    }
    if (overlap >= segWidth * 0.3) return i; // cột đầu tiên chồng lấn đáng kể
  }
  if (best >= 0 && bestOverlap > 0) return best;
  const center = (seg.xStart + seg.xEnd) / 2;
  let nearest = -1;
  let nearestDist = Infinity;
  cols.forEach((c, i) => {
    const d = Math.abs((c.xStart + c.xEnd) / 2 - center);
    if (d < nearestDist) {
      nearestDist = d;
      nearest = i;
    }
  });
  return nearestDist <= 22 ? nearest : -1;
}

/**
 * Kiểm tra một dòng 1-segment có phải dòng CUỐN của một ô trong bảng không.
 * Điều kiện chặt hơn assignColumn thông thường: chữ của một ô không thể rộng
 * hơn ô chứa nó quá nhiều và không thể chồng lấn đáng kể sang 2 cột — nếu
 * không, các dòng tiêu đề văn bản (rộng, căn giữa, nằm ngay trên bảng) sẽ bị
 * hút vào bảng thành một ô rác.
 */
function continuationColumn(seg: LineSegment, cols: ColumnRange[]): number {
  const segWidth = Math.max(1, seg.xEnd - seg.xStart);
  let found = -1;
  for (let i = 0; i < cols.length; i++) {
    const overlap = Math.min(seg.xEnd, cols[i].xEnd + 6) - Math.max(seg.xStart, cols[i].xStart - 6);
    if (overlap >= segWidth * 0.5) {
      if (found >= 0) return -1; // chồng lấn mạnh 2 cột = không phải dòng cuốn của 1 ô
      found = i;
    }
  }
  if (found < 0) return -1;
  const colWidth = Math.max(1, cols[found].xEnd - cols[found].xStart);
  if (segWidth > colWidth * 1.7) return -1;
  return found;
}

function tableRole(lines: TextLine[]): "data" | "signature" {
  const text = normalizeViText(lines.map((l) => l.text).join(" "));
  if (SIGNATURE_KEYWORDS.some((k) => text.includes(k))) return "signature";
  return "data";
}

function linesToTable(
  lines: TextLine[],
  cols: ColumnRange[],
  continuationIdx: Set<number>,
): DocBlock {
  const colCount = cols.length;
  const rows: DocRun[][][][] = [];
  lines.forEach((line, li) => {
    const cells: DocRun[][][] = Array.from({ length: colCount }, () => []);
    for (const seg of line.segments) {
      const c = assignColumn(seg, cols);
      if (c < 0) continue;
      if (cells[c].length > 0) cells[c].push(seg.runs);
      else cells[c] = [seg.runs];
    }
    if (continuationIdx.has(li) && rows.length > 0) {
      // Dòng cuốn của ô ở hàng trên: nối vào ô tương ứng của hàng cuối.
      const prev = rows[rows.length - 1];
      cells.forEach((cellLines, c) => {
        if (cellLines.length > 0) prev[c] = [...prev[c], ...cellLines];
      });
    } else {
      rows.push(cells);
    }
  });
  // Tỉ lệ cột theo khoảng x thực tế của mô hình cột.
  const spanStart = cols[0].xStart;
  const span = Math.max(1, cols[colCount - 1].xEnd - spanStart);
  const colWidths = cols.map((c, i) => {
    const next = i + 1 < colCount ? cols[i + 1].xStart : cols[colCount - 1].xEnd;
    return Math.max(4, Math.round(((next - c.xStart) / span) * 100));
  });
  return { kind: "table", rows, role: tableRole(lines), colWidths };
}

/* ------------------------------------------------------------------ */
/* Gộp dòng → đoạn                                                     */
/* ------------------------------------------------------------------ */

function shouldBreakParagraph(prev: TextLine, next: TextLine, bodyLeft: number): boolean {
  const gap = next.y - prev.y;
  if (gap > prev.fontSize * 2.1) return true; // dòng cách xa = đoạn mới
  if (HEADING_RE.test(next.text.trim())) return true;
  if (next.xStart > bodyLeft + 9 && prev.xEnd > bodyLeft + 60) return true; // thụt đầu dòng ở PDF
  return false;
}

function joinsContinuation(prev: TextLine, next: TextLine, rightEdge: number): boolean {
  // Dòng trước chạm gần mép phải vùng chữ → nhiều khả năng là dòng cuốn của đoạn.
  const reachesRight = prev.xEnd > rightEdge - prev.fontSize * 2.2;
  const nextStartsLower = /^[a-zà-ỹđ]/.test(next.text.trim());
  const prevOpen = !/[.!?:;"”)]$/.test(prev.text.trim());
  return reachesRight || (nextStartsLower && prevOpen);
}

function mergeLineIntoRuns(runs: DocRun[], line: TextLine): DocRun[] {
  const result = runs.map((r) => ({ ...r }));
  const incoming = line.runs.map((r) => ({ ...r }));
  if (result.length === 0) return incoming;
  const lastText = result[result.length - 1].text;
  if (lastText.endsWith("-") && /^[a-zà-ỹđ]/.test(incoming[0]?.text || "")) {
    // Gạch nối cuối dòng trong tiếng Việt hầu như luôn thuộc TỪ GHÉP
    // ("quốc phòng-" + "an ninh") — phải GIỮ gạch nối và nối liền không dấu
    // cách: "quốc phòng-an ninh". Xoá gạch nối sẽ tạo chữ dính sai ("phòngan").
    // (Trường hợp ngắt một từ đơn kiểu phương Tây rất hiếm trong văn bản VN.)
  } else if (!lastText.endsWith(" ")) {
    result[result.length - 1].text = lastText + " ";
  }
  for (const run of incoming) {
    const last = result[result.length - 1];
    if (last && !!last.bold === !!run.bold && !!last.italic === !!run.italic) {
      last.text += run.text;
    } else {
      result.push(run);
    }
  }
  return result;
}

/* ------------------------------------------------------------------ */
/* Hàm chính                                                           */
/* ------------------------------------------------------------------ */

export function classifyPdf(pages: PdfPageInput[]): PdfClassification {
  const textPages = pages.filter((p) => p.source === "text").length;
  const ocrPages = pages.filter((p) => p.source === "ocr").length;
  const totalChars = pages.reduce(
    (n, p) => n + p.items.reduce((m, it) => m + it.text.length, 0),
    0
  );
  let kind: PdfClassification["kind"] = "empty";
  if (pages.length > 0 && textPages === pages.length) kind = "text";
  else if (pages.length > 0 && ocrPages === pages.length) kind = "scanned";
  else if (pages.length > 0) kind = "mixed";
  const label =
    kind === "text"
      ? "PDF có lớp chữ (bản điện tử)"
      : kind === "scanned"
        ? "PDF scan (ảnh chụp/quét — đã OCR)"
        : kind === "mixed"
          ? "PDF hỗn hợp (một số trang có lớp chữ, một số trang scan)"
          : "PDF không đọc được nội dung";
  return { kind, totalPages: pages.length, textPages, ocrPages, totalChars, label };
}

/* ------------------------------------------------------------------ */
/* Trang OCR → khối: đoạn theo phân đoạn Tesseract + bảng thận trọng   */
/* ------------------------------------------------------------------ */

/**
 * Vùng bảng THẬT trên trang scan: ≥4 dòng liên tiếp, mỗi dòng có SỐ segment
 * bằng nhau (≥3 cột) và vị trí bắt đầu từng cột khớp nhau trong 14pt. Điều
 * kiện cố ý chặt: thà để bảng scan ra dòng chữ đúng thứ tự còn hơn nhận
 * nhầm đoạn văn thành bảng rồi xáo trộn/nuốt chữ như cơ chế quét cũ.
 */
function ocrTableRegionAt(lines: TextLine[], start: number): { end: number } | null {
  const segCount = lines[start].segments.length;
  if (segCount < 3) return null;
  let end = start;
  while (end + 1 < lines.length) {
    const next = lines[end + 1];
    if (next.segments.length !== segCount) break;
    if (next.y - lines[end].y > next.fontSize * 2.6) break;
    end++;
  }
  if (end - start + 1 < 4) return null;
  const region = lines.slice(start, end + 1);
  for (let c = 0; c < segCount; c++) {
    const starts = region.map((l) => l.segments[c].xStart);
    if (Math.max(...starts) - Math.min(...starts) > 14) return null;
  }
  return { end };
}

function ocrLinesToTable(lines: TextLine[]): DocBlock {
  const colCount = lines[0].segments.length;
  const rows: DocRun[][][][] = lines.map((line) => line.segments.map((seg) => [seg.runs]));
  const spanStart = Math.min(...lines.map((l) => l.segments[0].xStart));
  const spanEnd = Math.max(...lines.map((l) => l.segments[colCount - 1].xEnd));
  const span = Math.max(1, spanEnd - spanStart);
  const colWidths: number[] = [];
  for (let c = 0; c < colCount; c++) {
    const start = Math.min(...lines.map((l) => l.segments[c].xStart));
    const nextStart =
      c + 1 < colCount ? Math.min(...lines.map((l) => l.segments[c + 1].xStart)) : spanEnd;
    colWidths.push(Math.max(4, Math.round(((nextStart - start) / span) * 100)));
  }
  return { kind: "table", rows, role: tableRole(lines), colWidths };
}

/**
 * Dựng khối cho MỘT trang scan. Khác hẳn đường trang có lớp chữ (giữ
 * nguyên ở pdfPagesToBlocks): đoạn văn đi theo ranh giới đoạn của
 * Tesseract, không qua máy tách cột/gán cột — chữ được giữ đủ, đúng thứ
 * tự đọc; bảng chỉ dựng khi có bằng chứng cột rất mạnh (ocrTableRegionAt).
 */
export function ocrPageBlocks(
  page: PdfPageInput,
  pageIdx: number,
  titleConsumed: boolean
): { blocks: DocBlock[]; titleConsumed: boolean } {
  const blocks: DocBlock[] = [];
  const lines = ocrPageToLines(page).filter((l) => l.text.length > 0);
  let startIdx = 0;

  if (pageIdx === 0 && lines.length > 0) {
    const header = extractAdminHeader(lines, page.width);
    if (header) {
      blocks.push(header.block);
      startIdx = header.consumedLines;
    }
  }
  const content = lines.slice(startIdx);

  let paraLines: TextLine[] = [];
  const flushParagraph = () => {
    if (paraLines.length === 0) return;
    let runs: DocRun[] = [];
    paraLines.forEach((l) => (runs = mergeLineIntoRuns(runs, l)));
    const first = paraLines[0];
    const text = runs.map((r) => r.text).join("");
    const norm = normalizeViText(text);
    let role: BlockRole = "body";
    let align: "left" | "center" | "right" | "justify" = "justify";
    const centered = isCentered(first, page.width) && paraLines.length <= 3;

    if (!titleConsumed && pageIdx === 0 && centered && DOC_TITLES.some((t) => norm === t || norm.startsWith(t + " "))) {
      role = "title";
      align = "center";
      titleConsumed = true;
    } else if (titleConsumed && centered && paraLines.length <= 2 && blocks.length > 0 && (blocks[blocks.length - 1] as any).role === "title") {
      role = "subtitle";
      align = "center";
    } else if (HEADING_RE.test(text.trim()) && (isAllCaps(text) || /^[IVXLCDM]/.test(text.trim()) || paraLines.length === 1)) {
      // OCR không cho biết đậm/nghiêng — nới điều kiện mục so với trang có
      // lớp chữ: đoạn 1 dòng đánh số/viết hoa toàn bộ được coi là mục.
      role = "heading";
      align = "left";
    } else if (/^Nơi nhận/i.test(text.trim())) {
      role = "recipient";
      align = "left";
    } else if (centered && isAllCaps(text) && paraLines.length <= 2) {
      align = "center";
    } else if (first.xStart > page.width * 0.52 && paraLines.length <= 3) {
      align = "right";
    }
    blocks.push({
      kind: "paragraph",
      runs,
      align,
      role,
      fontSize: first.fontSize,
      firstLineIndent: role === "body",
    });
    paraLines = [];
  };

  let i = 0;
  while (i < content.length) {
    const region = ocrTableRegionAt(content, i);
    if (region) {
      flushParagraph();
      blocks.push(ocrLinesToTable(content.slice(i, region.end + 1)));
      i = region.end + 1;
      continue;
    }
    const line = content[i];
    if (paraLines.length > 0) {
      const prev = paraLines[paraLines.length - 1];
      const bigGap = line.y - prev.y > prev.fontSize * 2.1;
      if (line.paraBreakBefore || bigGap || HEADING_RE.test(line.text.trim())) {
        flushParagraph();
      }
    }
    paraLines.push(line);
    i++;
  }
  flushParagraph();

  page.imageDataUrls?.forEach((url, k) =>
    blocks.push({ kind: "image", dataUrl: url, width: page.imageSizes?.[k]?.width || 600, height: page.imageSizes?.[k]?.height || 400 })
  );
  return { blocks, titleConsumed };
}

export function pdfPagesToBlocks(pages: PdfPageInput[]): DocBlock[] {
  const blocks: DocBlock[] = [];
  let titleConsumed = false;

  pages.forEach((page) => {
    // Chỉ số trang TOÀN CỤC suy từ số trang thật (pageNumber), KHÔNG dùng vị
    // trí trong mảng: services/pdf.ts có lúc gọi hàm này cho TỪNG trang một
    // (khi ghép trang AI với trang thường) — nếu dùng vị trí mảng thì mọi
    // trang đều thành "trang 1" và phần đầu thể thức bị dựng nhầm ở mọi
    // trang (nuốt mất nội dung đầu trang 2, 3... — lỗi đã gặp ở CV 5095).
    const pageIdx = page.pageNumber - 1;
    // Trang scan đi đường riêng (giữ cấu trúc của Tesseract) — xem ocrPageBlocks.
    if (page.source === "ocr") {
      const res = ocrPageBlocks(page, pageIdx, titleConsumed);
      blocks.push(...res.blocks);
      titleConsumed = res.titleConsumed;
      return;
    }
    const lines = pageToLines(page).filter((l) => l.text.length > 0);
    let startIdx = 0;

    // Phần đầu thể thức: chỉ xét trang 1.
    if (pageIdx === 0 && lines.length > 0) {
      const header = extractAdminHeader(lines, page.width);
      if (header) {
        blocks.push(header.block);
        startIdx = header.consumedLines;
      }
    }

    const contentLines = lines.slice(startIdx);
    if (contentLines.length === 0 && page.imageDataUrls?.length) {
      page.imageDataUrls.forEach((url, i) =>
        blocks.push({ kind: "image", dataUrl: url, width: page.imageSizes?.[i]?.width || 600, height: page.imageSizes?.[i]?.height || 400 })
      );
      return;
    }

    // Mép trái/phải của vùng chữ để quyết định gộp đoạn.
    const bodyLeft = contentLines.length
      ? Math.min(...contentLines.map((l) => l.xStart))
      : 0;
    const rightEdge = contentLines.length
      ? Math.max(...contentLines.map((l) => l.xEnd))
      : page.width;

    let i = 0;
    let paraLines: TextLine[] = [];
    const flushParagraph = () => {
      if (paraLines.length === 0) return;
      let runs: DocRun[] = [];
      paraLines.forEach((l) => (runs = mergeLineIntoRuns(runs, l)));
      const first = paraLines[0];
      const text = runs.map((r) => r.text).join("");
      const norm = normalizeViText(text);
      let role: BlockRole = "body";
      let align: "left" | "center" | "right" | "justify" = "justify";
      const centered = isCentered(first, page.width) && paraLines.length <= 3;

      if (!titleConsumed && pageIdx === 0 && centered && DOC_TITLES.some((t) => norm === t || norm.startsWith(t + " "))) {
        role = "title";
        align = "center";
        titleConsumed = true;
      } else if (role === "body" && titleConsumed && centered && paraLines.length <= 2 && blocks.length > 0 && (blocks[blocks.length - 1] as any).role === "title") {
        // Trích yếu: dòng căn giữa NGAY DƯỚI tên loại văn bản — nhận diện theo vị trí
        // vì pdf.js thường không cho biết chữ đậm của font subset (xem services/pdf.ts).
        role = "subtitle";
        align = "center";
      } else if (HEADING_RE.test(text.trim()) && (first.bold || isAllCaps(text) || /^[IVXLCDM]/.test(text.trim()))) {
        role = "heading";
        align = "left";
      } else if (/^Nơi nhận/i.test(text.trim())) {
        role = "recipient";
        align = "left";
      } else if (centered && first.bold) {
        align = "center";
      } else if (first.xStart > page.width * 0.52 && paraLines.length <= 3) {
        align = "right";
      }
      blocks.push({
        kind: "paragraph",
        runs,
        align,
        role,
        fontSize: first.fontSize,
        firstLineIndent: role === "body",
      });
      paraLines = [];
    };

    while (i < contentLines.length) {
      const line = contentLines[i];

      // --- Bảng: quét cả VÙNG gồm các dòng nhiều segment liên tiếp ---
      // Cách cũ (so từng cặp dòng, bắt buộc bằng số segment) vỡ trên PDF thật:
      // hàng có ô trống hoặc ô cuốn 2 dòng là mất bảng; và 2 dòng tiêu đề trang
      // (cùng 2 segment) lại bị nhận nhầm thành bảng 2x2. Cách mới: gom vùng,
      // dựng mô hình cột từ các dòng đầy đủ nhất, gán mọi segment theo trục x;
      // dòng chỉ có 1 segment nhưng chồng lên một cột của vùng = dòng cuốn của
      // ô ở hàng trên (ghép vào ô đó thay vì thành hàng riêng).
      if (lineIsTableRow(line)) {
        const region: TextLine[] = [];
        const continuationIdx = new Set<number>();
        let j = i;
        let colsCache: ColumnRange[] | null = null;
        while (j < contentLines.length) {
          const L = contentLines[j];
          const prevLine = region[region.length - 1];
          if (prevLine && L.y - prevLine.y >= L.fontSize * 2.6) break;
          if (lineIsTableRow(L)) {
            region.push(L);
            j++;
            colsCache = null;
            continue;
          }
          if (L.segments.length === 1 && region.filter(lineIsTableRow).length >= 2) {
            if (!colsCache) colsCache = buildColumns(region.filter(lineIsTableRow));
            if (colsCache.length >= 2 && continuationColumn(L.segments[0], colsCache) >= 0) {
              continuationIdx.add(region.length);
              region.push(L);
              j++;
              continue;
            }
          }
          break;
        }
        const candidateRows = region.filter(lineIsTableRow);
        if (candidateRows.length >= 2) {
          const cols = buildColumns(candidateRows);
          const role = tableRole(region);
          // Bảng dữ liệu phải đủ lớn (≥3 hàng hoặc ≥3 cột) mới tin là bảng —
          // tránh nhận nhầm 2 dòng chữ song song (đầu trang, 2 cột chữ ký lửng).
          // Khối chữ ký thì giữ ngưỡng thấp vì bản chất nó chỉ có 2 cột.
          const sizeOk = role === "signature" || region.length >= 3 || cols.length >= 3;
          if (cols.length >= 2 && sizeOk) {
            flushParagraph();
            blocks.push(linesToTable(region, cols, continuationIdx));
            i = j;
            continue;
          }
        }
      }

      // --- Đoạn văn ---
      if (paraLines.length === 0) {
        paraLines = [line];
      } else {
        const prev = paraLines[paraLines.length - 1];
        const centeredNext = isCentered(line, page.width) && line.text.length < 120;
        if (
          centeredNext ||
          shouldBreakParagraph(prev, line, bodyLeft) ||
          !joinsContinuation(prev, line, rightEdge)
        ) {
          // Vẫn gộp khi dòng chỉ đơn giản là dòng cuốn tiếp theo có thụt lề nhẹ do OCR.
          if (!centeredNext && !shouldBreakParagraph(prev, line, bodyLeft)) {
            flushParagraph();
          }
          if (paraLines.length > 0 && (centeredNext || shouldBreakParagraph(prev, line, bodyLeft))) {
            flushParagraph();
          }
        }
        paraLines.push(line);
      }
      i++;
    }
    flushParagraph();

    // Ảnh nhúng của trang (đặt cuối trang, giữ thứ tự trang).
    page.imageDataUrls?.forEach((url, k) =>
      blocks.push({ kind: "image", dataUrl: url, width: page.imageSizes?.[k]?.width || 600, height: page.imageSizes?.[k]?.height || 400 })
    );
  });

  return blocks;
}

export function blocksToRawText(blocks: DocBlock[]): string {
  const parts: string[] = [];
  for (const b of blocks) {
    if (b.kind === "paragraph") parts.push(b.runs.map((r) => r.text).join(""));
    else if (b.kind === "table")
      parts.push(
        b.rows
          .map((row) => row.map((cell) => cell.map((line) => line.map((r) => r.text).join("")).join("\n")).join("\t"))
          .join("\n")
      );
  }
  return parts.join("\n\n");
}

function runsToHtml(runs: DocRun[]): string {
  return runs
    .map((r) => {
      let out = escapeHtml(r.text);
      if (r.bold) out = `<strong>${out}</strong>`;
      if (r.italic) out = `<em>${out}</em>`;
      if (r.underline) out = `<u>${out}</u>`;
      return out;
    })
    .join("");
}

/** Dựng HTML tương thích pipeline xem trước/chuẩn hoá NĐ 30 sẵn có của app. */
export function blocksToHtml(blocks: DocBlock[]): string {
  const out: string[] = [];
  for (const b of blocks) {
    if (b.kind === "paragraph") {
      out.push(`<p>${runsToHtml(b.runs)}</p>`);
    } else if (b.kind === "image") {
      out.push(`<p><img src="${b.dataUrl}" alt="Ảnh trích từ PDF" /></p>`);
    } else {
      const rowsHtml = b.rows
        .map((row) => {
          const cells = row
            .map((cell) => {
              const inner = cell.map((line) => `<p>${runsToHtml(line)}</p>`).join("") || "<p></p>";
              return `<td>${inner}</td>`;
            })
            .join("");
          return `<tr>${cells}</tr>`;
        })
        .join("");
      out.push(`<table>${rowsHtml}</table>`);
    }
  }
  return out.join("\n");
}
