/**
 * aiText.ts — Dựng DocBlock từ văn bản AI "đánh lại" trang scan.
 *
 * Gemini (vision) đọc ảnh trang và trả văn bản thuần có cấu trúc nhẹ:
 * đoạn cách nhau một dòng trống, bảng dạng markdown, **đậm** / *nghiêng*.
 * File này biến văn bản đó thành cùng DocBlock[] mà đường lớp chữ/OCR tạo
 * ra, để xem trước + xuất Word dùng CHUNG một bộ máy định dạng NĐ 30.
 *
 * Nguyên tắc: TRUNG THÀNH với chữ của AI (không sửa nội dung) — chỉ gán
 * vai trò thể thức: đầu văn bản hành chính (bảng 36/64), tiêu đề, trích
 * yếu, "Kính gửi", đoạn thân, bảng, khối ký.
 */
import type { BlockRole, DocBlock, DocRun } from "./pdfStructure";

/** Bóc **đậm** / *nghiêng* thành DocRun, giữ nguyên chữ. */
function parseInline(text: string): DocRun[] {
  const runs: DocRun[] = [];
  const re = /\*\*([^*]+)\*\*|\*([^*]+)\*/g;
  let last = 0;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) runs.push({ text: text.slice(last, m.index) });
    if (m[1] !== undefined) runs.push({ text: m[1], bold: true });
    else if (m[2] !== undefined) runs.push({ text: m[2], italic: true });
    last = m.index + m[0].length;
  }
  if (last < text.length) runs.push({ text: text.slice(last) });
  return runs.length ? runs : [{ text }];
}

function stripHeadingMarks(line: string): string {
  return line.replace(/^#{1,6}\s+/, "").trim();
}

function isAllCaps(text: string): boolean {
  const letters = text.replace(/[^A-Za-zÀ-ỹĐđ]/g, "");
  if (letters.length < 3) return false;
  const upper = letters.replace(/[^A-ZÀ-ỸĐ]/g, "").length;
  return upper / letters.length > 0.75;
}

function para(
  text: string,
  role: BlockRole,
  align: "left" | "center" | "right" | "justify" = "justify",
  opts?: { bold?: boolean; firstLineIndent?: boolean }
): DocBlock {
  const runs = opts?.bold ? [{ text, bold: true }] : parseInline(text);
  return {
    kind: "paragraph",
    runs,
    align,
    role,
    firstLineIndent: opts?.firstLineIndent,
  };
}

/** Một chunk có phải bảng markdown không (mọi dòng đều bắt đầu bằng |). */
function isMarkdownTable(lines: string[]): boolean {
  const nonEmpty = lines.filter((l) => l.trim());
  return nonEmpty.length >= 2 && nonEmpty.every((l) => l.trim().startsWith("|"));
}

function markdownTableToBlock(lines: string[]): DocBlock {
  const rowsRaw = lines
    .map((l) => l.trim())
    .filter((l) => l.startsWith("|"))
    .filter((l) => !/^\|[\s:|-]+\|$/.test(l)); // bỏ dòng phân cách |---|---|
  const rows: DocRun[][][][] = rowsRaw.map((rowLine, ri) => {
    const cells = rowLine.replace(/^\||\|$/g, "").split("|").map((c) => c.trim());
    return cells.map((cell) => {
      if (!cell) return [[{ text: "" }]];
      const cellRuns: DocRun[] = ri === 0 ? [{ text: cell, bold: true }] : parseInline(cell);
      return [cellRuns];
    });
  });
  const colCount = Math.max(...rows.map((r) => r.length), 1);
  const w = Math.floor(100 / colCount);
  return {
    kind: "table",
    rows,
    role: "data",
    colWidths: Array.from({ length: colCount }, (_, i) =>
      i === colCount - 1 ? 100 - w * (colCount - 1) : w
    ),
  };
}

const NATIONAL_RE = /CỘNG H[OÒ]A XÃ HỘI CHỦ NGHĨA|CỘNG HOÀ XÃ HỘI CHỦ NGHĨA/;
const MOTTO_RE = /Độc lập\s*[-–—]\s*Tự do/;
const NUMBER_RE = /^Số\s*[:：]/;
const DATE_RE = /ngày\s+\d+\s+tháng\s+\d+\s+năm\s+\d+|ngày\s+\d+\s*\/\s*\d+\s*\/\s*\d+/;
const HEADING_RE = /^(Phần|Chương|Mục|Tiểu mục|Điều)\s+\S|^[IVXLCDM]+\.\s|^\d+(\.\d+)*\.\s+\S/;

/**
 * Văn bản AI đánh lại → DocBlock[].
 * `text` là toàn bộ các trang nối nhau bằng "\n\n" — hàm xử lý theo từng
 * trang được gọi riêng (mỗi trang một lần gọi) để đầu trang/trang mới
 * không lẫn nhau; trang không phải trang đầu thường không có đầu văn bản.
 */
export function aiTextToBlocks(rawText: string): DocBlock[] {
  const text = (rawText || "").replace(/\r\n/g, "\n").trim();
  if (!text) return [];
  // Tách chunk theo dòng trống; trong chunk giữ từng dòng.
  const chunks: string[][] = text
    .split(/\n[ \t]*\n/)
    .map((c) => c.split("\n").map((l) => stripHeadingMarks(l)).filter((l) => l.trim() !== ""))
    .filter((c) => c.length > 0);

  const blocks: DocBlock[] = [];
  let startChunk = 0;

  /* --- Đầu văn bản hành chính: các dòng đầu của chunk đầu tiên --- */
  if (chunks.length > 0) {
    const lines = chunks[0];
    const nationalIdx = lines.findIndex((l) => NATIONAL_RE.test(l));
    if (nationalIdx > 0) {
      const agencyLines = lines.slice(0, nationalIdx);
      const mottoLine = lines[nationalIdx + 1] && MOTTO_RE.test(lines[nationalIdx + 1]) ? lines[nationalIdx + 1] : undefined;
      const numberLine = lines.find((l) => NUMBER_RE.test(l));
      const dateLine = lines.find((l, i) => i !== nationalIdx && DATE_RE.test(l) && !NUMBER_RE.test(l));
      const consumed = new Set<number>();
      for (let i = 0; i <= nationalIdx; i++) consumed.add(i);
      if (mottoLine) consumed.add(nationalIdx + 1);
      lines.forEach((l, i) => {
        if (l === numberLine || l === dateLine) consumed.add(i);
      });
      const leftCell: DocRun[][] = [
        ...agencyLines.map((l) => [{ text: l, bold: true }]),
        ...(numberLine ? [[{ text: numberLine }]] : []),
      ];
      const rightCell: DocRun[][] = [
        [{ text: lines[nationalIdx], bold: true }],
        ...(mottoLine ? [[{ text: mottoLine, italic: true }]] : []),
        ...(dateLine ? [[{ text: dateLine, italic: true }]] : []),
      ];
      blocks.push({ kind: "table", rows: [[leftCell, rightCell]], role: "adminHeader", colWidths: [36, 64] });
      const rest = lines.filter((_, i) => !consumed.has(i));
      if (rest.length) chunks[0] = rest;
      else startChunk = 1;
    }
  }

  let sawBody = false;
  let sawRecipients = false;
  let subtitleDone = false;

  for (let ci = startChunk; ci < chunks.length; ci++) {
    const lines = chunks[ci];
    const joined = lines.join(" ").replace(/\s+/g, " ").trim();

    if (isMarkdownTable(lines)) {
      blocks.push(markdownTableToBlock(lines));
      sawBody = true;
      continue;
    }

    /* Tiêu đề + trích yếu: chỉ ở phần đầu, trước mọi đoạn thân. */
    if (!sawBody && blocks.length >= 0) {
      const lastBlock = blocks[blocks.length - 1];
      const prevIsHeaderOrTitle =
        !lastBlock || (lastBlock.kind === "table" && lastBlock.role === "adminHeader") ||
        (lastBlock.kind === "paragraph" && (lastBlock.role === "title" || lastBlock.role === "subtitle"));
      if (prevIsHeaderOrTitle && lines.length <= 2 && isAllCaps(joined) && joined.length <= 140) {
        blocks.push(para(joined, "title", "center", { bold: true }));
        continue;
      }
      if (
        prevIsHeaderOrTitle && !subtitleDone &&
        ((lastBlock?.kind === "paragraph" && lastBlock.role === "title") ||
          (lastBlock?.kind === "table" && lastBlock.role === "adminHeader")) &&
        /^V\/v\s/i.test(joined)
      ) {
        blocks.push(para(joined, "subtitle", "center", { bold: true }));
        subtitleDone = true;
        continue;
      }
    }

    if (/^Kính gửi\s*[:：]?/.test(joined)) {
      blocks.push(para(joined, "recipient", "left", { firstLineIndent: true }));
      sawBody = true;
      continue;
    }

    if (/^Nơi nhận\s*[:：]?/.test(lines[0])) {
      sawRecipients = true;
      blocks.push(para(lines[0], "body", "left", { bold: true }));
      for (const l of lines.slice(1)) blocks.push(para(l, "body", "left"));
      sawBody = true;
      continue;
    }

    /* Khối ký: sau "Nơi nhận" hoặc mở đầu bằng KT./TM./TL. */
    if ((sawRecipients || /^(KT\.|TM\.|TL\.|TUQ\.)/.test(joined)) && joined.length <= 160) {
      const align = isAllCaps(joined) || /^(KT\.|TM\.|TL\.|TUQ\.)/.test(joined) ? "center" : "left";
      blocks.push(para(joined, "signatureText", align as "center" | "left", { bold: isAllCaps(joined) }));
      continue;
    }

    /* Dòng liệt kê (- … / • …): mỗi ý một đoạn riêng. */
    if (lines.length > 1 && lines.every((l) => /^[-•*]\s/.test(l))) {
      for (const l of lines) blocks.push(para(l, "body", "justify"));
      sawBody = true;
      continue;
    }

    /* Tiêu đề mục: một dòng, mẫu Phần/Chương/Mục/Điều/số La Mã/số thứ tự. */
    if (lines.length === 1 && HEADING_RE.test(joined) && joined.length <= 110) {
      blocks.push(para(joined, "heading", "left", { bold: true }));
      sawBody = true;
      continue;
    }

    /* Mặc định: đoạn thân — nối các dòng trong chunk thành một đoạn liền. */
    blocks.push(para(joined, "body", "justify", { firstLineIndent: true }));
    sawBody = true;
  }

  return blocks;
}

/** Trích các bảng markdown trong văn bản AI thành lưới ô chữ (cho Excel). */
export function aiTextToTables(rawText: string): string[][][] {
  const tables: string[][][] = [];
  for (const b of aiTextToBlocks(rawText)) {
    if (b.kind !== "table" || b.role !== "data") continue;
    tables.push(
      b.rows.map((row) => row.map((cell) => cell.map((line) => line.map((r) => r.text).join("")).join("\n")))
    );
  }
  return tables;
}
