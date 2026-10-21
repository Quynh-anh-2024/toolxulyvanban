/**
 * docxExport.ts — Xuất mô hình khối (DocBlock[]) thành file Word .docx THẬT
 * (định dạng Office Open XML) trình bày đúng thể thức NĐ 30/2020/NĐ-CP.
 *
 * Khác với đường xuất cũ (file .doc dạng MHTML), file .docx ở đây mở/chỉnh sửa
 * chuẩn trong Word, LibreOffice, Google Docs và giữ đúng:
 *  - Khổ A4; lề trên/dưới 2cm, trái 3cm, phải 2cm (chỉnh được theo cấu hình app).
 *  - Font Times New Roman; nội dung 13–14pt, căn đều 2 bên, giãn dòng theo cấu hình,
 *    cách đoạn 6pt, thụt đầu dòng 1–1,27cm.
 *  - Phần đầu thể thức: bảng 2 cột 36%/64% không viền — cơ quan in hoa (dòng cuối đậm),
 *    quốc hiệu 13pt in hoa đậm, tiêu ngữ 13pt đậm có gạch chân dưới dòng chữ,
 *    số ký hiệu thường, địa danh + ngày tháng nghiêng.
 *  - Tên loại văn bản: in hoa, đậm, căn giữa; trích yếu: đậm, căn giữa.
 *  - Bảng dữ liệu: viền nét liền, hàng đầu đậm căn giữa.
 *  - Khối chữ ký/nơi nhận: bảng không viền, cột phải căn giữa.
 */
import {
  AlignmentType,
  BorderStyle,
  Document,
  ImageRun,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  UnderlineType,
  WidthType,
  convertMillimetersToTwip,
} from "docx";
import type { DocBlock, DocRun } from "./pdfStructure";
import type { FormattingConfig } from "./docx";

const FONT = "Times New Roman";
/** docx: cỡ chữ tính bằng nửa point. */
const hp = (pt: number) => Math.round(pt * 2);

const NO_BORDERS = {
  top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  insideVertical: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
} as const;

const CELL_BORDERS = {
  top: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
  bottom: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
  left: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
  right: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
  insideHorizontal: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
  insideVertical: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
} as const;

function toTextRuns(runs: DocRun[], sizePt: number, force?: Partial<DocRun>): TextRun[] {
  return runs.map(
    (r) =>
      new TextRun({
        // NFC lần cuối trước khi ghi vào Word — bảo hiểm chống "lỗi ký tự".
        text: (r.text || "").normalize("NFC"),
        font: FONT,
        size: hp(sizePt),
        bold: force?.bold ?? r.bold,
        italics: force?.italic ?? r.italic,
        underline:
          (force?.underline ?? r.underline)
            ? { type: UnderlineType.SINGLE }
            : undefined,
      })
  );
}

function lineSpacingTwips(spacing: string): number {
  const f = parseFloat(spacing);
  return Math.round((Number.isFinite(f) && f > 0 ? f : 1.5) * 240);
}

function bodyParagraph(
  block: Extract<DocBlock, { kind: "paragraph" }>,
  config: FormattingConfig
): Paragraph {
  const size = config.size;
  const base = {
    spacing: { after: config.paraSpacing * 20, line: lineSpacingTwips(config.spacing) },
  };
  switch (block.role) {
    case "title":
      return new Paragraph({
        ...base,
        alignment: AlignmentType.CENTER,
        spacing: { before: 240, after: 120, line: lineSpacingTwips(config.spacing) },
        children: toTextRuns(
          block.runs.map((r) => ({ ...r, text: r.text.toUpperCase() })),
          Math.max(size, 14),
          { bold: true }
        ),
      });
    case "subtitle":
      return new Paragraph({
        ...base,
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 120, line: lineSpacingTwips(config.spacing) },
        children: toTextRuns(block.runs, size, { bold: true }),
      });
    case "heading":
      return new Paragraph({
        ...base,
        alignment: AlignmentType.LEFT,
        spacing: { before: 120, after: config.paraSpacing * 20, line: lineSpacingTwips(config.spacing) },
        children: toTextRuns(block.runs, size, { bold: true }),
      });
    case "recipient":
      return new Paragraph({
        ...base,
        alignment: AlignmentType.LEFT,
        children: toTextRuns(block.runs, size - 1),
      });
    default:
      return new Paragraph({
        ...base,
        alignment:
          block.align === "center"
            ? AlignmentType.CENTER
            : block.align === "right"
              ? AlignmentType.RIGHT
              : block.align === "left"
                ? AlignmentType.LEFT
                : AlignmentType.JUSTIFIED,
        indent: block.firstLineIndent
          ? { firstLine: convertMillimetersToTwip(config.textIndent * 10) }
          : undefined,
        children: toTextRuns(block.runs, size),
      });
  }
}

type AlignValue = (typeof AlignmentType)[keyof typeof AlignmentType];

function cellParagraphs(
  cellLines: DocRun[][],
  opts: { align: AlignValue; sizePt: number; bold?: boolean; line?: number }
): Paragraph[] {
  if (cellLines.length === 0) return [new Paragraph({ children: [] })];
  return cellLines.map(
    (line) =>
      new Paragraph({
        alignment: opts.align,
        spacing: { after: 40, line: opts.line ?? 276 },
        children: toTextRuns(line, opts.sizePt, opts.bold !== undefined ? { bold: opts.bold } : undefined),
      })
  );
}

function adminHeaderTable(block: Extract<DocBlock, { kind: "table" }>): Table {
  const [row1, row2] = [block.rows[0] || [[], []], block.rows[1] || [[], []]];
  const leftTop = row1[0] || [];
  const rightTop = row1[1] || [];
  const leftBottom = row2[0] || [];
  const rightBottom = row2[1] || [];

  const centerCell = (lines: Paragraph[], widthPct: number) =>
    new TableCell({
      width: { size: widthPct, type: WidthType.PERCENTAGE },
      borders: NO_BORDERS,
      children: lines.length ? lines : [new Paragraph({ children: [] })],
    });

  // Cột trái: các dòng cơ quan — dòng cơ quan ban hành (cuối) in hoa đậm, dòng chủ quản in hoa thường.
  const leftParas = leftTop.map((line, idx) => {
    const isIssuing = idx === leftTop.length - 1 && leftTop.length > 1;
    return new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 20, line: 240 },
      children: toTextRuns(
        line.map((r) => ({ ...r, text: r.text.toUpperCase() })),
        13,
        { bold: isIssuing }
      ),
    });
  });
  // Cột phải: quốc hiệu 13 in hoa đậm; tiêu ngữ 13–14 đậm, gạch chân đúng độ dài dòng chữ.
  const rightParas = rightTop.map((line, idx) => {
    const isMotto = idx === 1;
    return new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 20, line: 240 },
      children: toTextRuns(line, isMotto ? 14 : 13, {
        bold: true,
        underline: isMotto ? true : undefined,
      }),
    });
  });

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    columnWidths: [convertMillimetersToTwip(58), convertMillimetersToTwip(103)],
    borders: NO_BORDERS,
    rows: [
      new TableRow({
        children: [centerCell(leftParas, 36), centerCell(rightParas, 64)],
      }),
      new TableRow({
        children: [
          centerCell(
            leftBottom.map(
              (line) =>
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 20, line: 240 },
                  children: toTextRuns(line, 13),
                })
            ),
            36
          ),
          centerCell(
            rightBottom.map(
              (line) =>
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 20, line: 240 },
                  children: toTextRuns(line, 13, { italic: true }),
                })
            ),
            64
          ),
        ],
      }),
    ],
  });
}

function dataTable(block: Extract<DocBlock, { kind: "table" }>, config: FormattingConfig): Table {
  const colCount = Math.max(...block.rows.map((r) => r.length), 1);
  const widths = block.colWidths && block.colWidths.length === colCount
    ? block.colWidths
    : Array.from({ length: colCount }, () => Math.floor(100 / colCount));
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: CELL_BORDERS,
    rows: block.rows.map(
      (row, rowIdx) =>
        new TableRow({
          tableHeader: rowIdx === 0,
          children: Array.from({ length: colCount }, (_, c) => {
            const cell = row[c] || [];
            return new TableCell({
              width: { size: widths[c], type: WidthType.PERCENTAGE },
              borders: CELL_BORDERS,
              children: cellParagraphs(cell, {
                align: rowIdx === 0 ? AlignmentType.CENTER : AlignmentType.LEFT,
                sizePt: config.size - 1,
                bold: rowIdx === 0 ? true : undefined,
                line: 276,
              }),
            });
          }),
        })
    ),
  });
}

function signatureTable(block: Extract<DocBlock, { kind: "table" }>, config: FormattingConfig): Table {
  const colCount = Math.max(...block.rows.map((r) => r.length), 1);
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: NO_BORDERS,
    rows: block.rows.map(
      (row) =>
        new TableRow({
          children: Array.from({ length: colCount }, (_, c) => {
            const cell = row[c] || [];
            const isRight = c === colCount - 1 && colCount > 1;
            return new TableCell({
              width: { size: Math.floor(100 / colCount), type: WidthType.PERCENTAGE },
              borders: NO_BORDERS,
              children: cellParagraphs(cell, {
                align: isRight ? AlignmentType.CENTER : AlignmentType.LEFT,
                sizePt: c === 0 ? config.size - 1 : config.size,
                line: 276,
              }),
            });
          }),
        })
    ),
  });
}

function imageBlockToParagraph(block: Extract<DocBlock, { kind: "image" }>): Paragraph {
  const match = block.dataUrl.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.*)$/);
  if (!match) return new Paragraph({ children: [] });
  const maxWidthPx = 460;
  const scale = block.width > maxWidthPx ? maxWidthPx / block.width : 1;
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [
      new ImageRun({
        type: (match[1].split("/")[1] || "png").replace("jpg", "jpeg") as any,
        data: match[2],
        transformation: {
          width: Math.round(block.width * scale),
          height: Math.round(block.height * scale),
        },
      }),
    ],
  });
}

export function blocksToDocxDocument(blocks: DocBlock[], config: FormattingConfig): Document {
  const children: Array<Paragraph | Table> = [];
  for (const block of blocks) {
    if (block.kind === "paragraph") children.push(bodyParagraph(block, config));
    else if (block.kind === "image") children.push(imageBlockToParagraph(block));
    else if (block.role === "adminHeader") children.push(adminHeaderTable(block));
    else if (block.role === "signature") children.push(signatureTable(block, config));
    else children.push(dataTable(block, config));
  }
  if (children.length === 0) children.push(new Paragraph({ children: [] }));

  return new Document({
    styles: {
      default: {
        document: {
          run: { font: FONT, size: hp(config.size) },
          paragraph: { spacing: { after: config.paraSpacing * 20 } },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: convertMillimetersToTwip(210), height: convertMillimetersToTwip(297) },
            margin: {
              top: convertMillimetersToTwip(20),
              bottom: convertMillimetersToTwip(20),
              left: convertMillimetersToTwip(config.leftMargin * 10),
              right: convertMillimetersToTwip(config.rightMargin * 10),
            },
          },
        },
        children,
      },
    ],
  });
}

/** Tạo Blob .docx trong trình duyệt (Packer.toBlob). */
export async function buildDocxBlob(blocks: DocBlock[], config: FormattingConfig): Promise<Blob> {
  return Packer.toBlob(blocksToDocxDocument(blocks, config));
}

/** Tạo Buffer .docx trong Node (dùng cho kiểm thử tự động). */
export async function buildDocxBuffer(blocks: DocBlock[], config: FormattingConfig): Promise<Buffer> {
  return Packer.toBuffer(blocksToDocxDocument(blocks, config));
}

export async function exportBlocksToWord(
  blocks: DocBlock[],
  fileName: string,
  config: FormattingConfig
): Promise<void> {
  const blob = await buildDocxBlob(blocks, config);
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `${fileName}_ChuanND30.docx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(a.href);
}
