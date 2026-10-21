/**
 * Test đơn vị cho lõi chuyển PDF → khối tài liệu (src/utils/pdfStructure.ts).
 * Chạy: npx tsx tests/pdf-structure.test.mts
 */
import assert from "node:assert";
import { writeFileSync } from "node:fs";
import {
  PdfPageInput,
  PdfTextItem,
  classifyPdf,
  pageTextLooksGarbled,
  pdfPagesToBlocks,
  blocksToRawText,
  type DocBlock,
} from "../src/utils/pdfStructure";
import { buildDocxBuffer } from "../src/utils/docxExport";
import type { FormattingConfig } from "../src/utils/docx";

const CONFIG: FormattingConfig = {
  font: "Times New Roman",
  size: 14,
  spacing: "1.5",
  textAlign: "justify",
  textIndent: 1.27,
  paraSpacing: 6,
  leftMargin: 3,
  rightMargin: 2,
};

function item(text: string, x: number, y: number, w: number, opts: Partial<PdfTextItem> = {}): PdfTextItem {
  return {
    text,
    x,
    y,
    width: w,
    height: opts.fontSize || 13,
    fontSize: opts.fontSize || 13,
    bold: opts.bold || false,
    italic: opts.italic || false,
    fontName: opts.fontName || "LiberationSerif",
  };
}

/** Trang A4 (595x842pt) mô phỏng một KẾ HOẠCH hành chính điển hình. */
function sampleAdminPage(): PdfPageInput {
  const items: PdfTextItem[] = [
    // --- Phần đầu thể thức ---
    item("PHÒNG GIÁO DỤC VÀ ĐÀO TẠO HUYỆN MÈO VẠC", 62, 40, 190, { fontSize: 11 }),
    item("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", 268, 40, 262, { fontSize: 12, bold: true }),
    item("TRƯỜNG PTDTBT TH MÈO VẠC", 92, 58, 150, { fontSize: 11, bold: true }),
    item("Độc lập - Tự do - Hạnh phúc", 320, 58, 160, { fontSize: 13, bold: true }),
    item("Số: 12/KH-THMV", 118, 80, 78, { fontSize: 12 }),
    item("Mèo Vạc, ngày 05 tháng 10 năm 2026", 292, 80, 190, { fontSize: 12, italic: true }),
    // --- Tên loại + trích yếu ---
    item("KẾ HOẠCH", 262, 116, 70, { fontSize: 14, bold: true }),
    item("Tổ chức dạy học năm học 2026 - 2027", 180, 138, 235, { fontSize: 14, bold: true }),
    // --- Đoạn 1: 2 dòng cuốn (dòng 2 là tiếp nối, không thụt) ---
    item("Căn cứ Thông tư số 32/2018/TT-BGDĐT ngày 26 tháng 12 năm 2018 của Bộ", 98, 168, 415),
    item("trưởng Bộ Giáo dục và Đào tạo về việc ban hành Chương trình giáo dục phổ thông;", 85, 188, 428),
    // --- Đoạn 2: thụt đầu dòng ---
    item("Trường PTDTBT TH Mèo Vạc xây dựng kế hoạch tổ chức dạy học như sau:", 98, 214, 400),
    // --- Mục lớn ---
    item("I. MỤC ĐÍCH, YÊU CẦU", 85, 244, 170, { bold: true }),
    item("1. Nâng cao chất lượng dạy học các môn học và hoạt động giáo dục.", 98, 270, 395),
    // --- Bảng dữ liệu 3 cột ---
    item("STT", 90, 306, 24, { bold: true }),
    item("Nội dung công việc", 170, 306, 110, { bold: true }),
    item("Thời gian", 400, 306, 60, { bold: true }),
    item("1", 96, 328, 8),
    item("Dạy học theo chương trình", 170, 328, 140),
    item("Tuần 1 - 35", 400, 328, 62),
    item("2", 96, 350, 8),
    item("Kiểm tra định kỳ giữa kỳ", 170, 350, 138),
    item("Tuần 9", 400, 350, 40),
    // --- Khối nơi nhận / chữ ký (2 cột) ---
    item("Nơi nhận:", 85, 400, 55, { italic: true }),
    item("TM. BAN GIÁM HIỆU", 330, 400, 120, { bold: true }),
    item("- Ban Giám hiệu;", 85, 420, 90),
    item("HIỆU TRƯỞNG", 352, 420, 80, { bold: true }),
    item("- Lưu: VT.", 85, 440, 60),
    item("Trần Ngọc Hải", 350, 486, 85, { bold: true }),
  ];
  return { pageNumber: 1, width: 595, height: 842, items, source: "text" };
}

async function main() {
  let passed = 0;
  const check = (name: string, fn: () => void) => {
    fn();
    passed++;
    console.log(`  ✓ ${name}`);
  };

  const blocks = pdfPagesToBlocks([sampleAdminPage()]);
  const raw = blocksToRawText(blocks);
  console.log("Blocks:", blocks.map((b) => (b.kind === "table" ? `table(${b.role})` : b.kind === "paragraph" ? `p(${b.role})` : "image")).join(", "));

  check("Nhận diện bảng thể thức hành chính ở đầu", () => {
    const header = blocks[0];
    assert.equal(header.kind, "table");
    assert.equal((header as any).role, "adminHeader");
  });

  check("Cột trái header: cơ quan + số ký hiệu; cột phải: quốc hiệu + tiêu ngữ + ngày", () => {
    const header = blocks[0] as Extract<DocBlock, { kind: "table" }>;
    const cellText = (r: number, c: number) =>
      header.rows[r][c].map((line) => line.map((run) => run.text).join("")).join("\n");
    assert.match(cellText(0, 0), /PHÒNG GIÁO DỤC/);
    assert.match(cellText(0, 0), /TRƯỜNG PTDTBT TH MÈO VẠC/);
    assert.equal(cellText(0, 1).split("\n")[0], "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM");
    assert.match(cellText(0, 1), /Độc lập - Tự do - Hạnh phúc/);
    assert.match(cellText(1, 0), /Số: 12\/KH-THMV/);
    assert.match(cellText(1, 1), /ngày 05 tháng 10 năm 2026/);
  });

  check("Tên loại văn bản thành block title, trích yếu thành subtitle", () => {
    const title = blocks.find((b) => b.kind === "paragraph" && b.role === "title");
    assert.ok(title, "thiếu block title");
    assert.match((title as any).runs.map((r: any) => r.text).join(""), /KẾ HOẠCH/);
    const sub = blocks.find((b) => b.kind === "paragraph" && b.role === "subtitle");
    assert.ok(sub, "thiếu block subtitle");
  });

  check("2 dòng cuốn của đoạn 'Căn cứ...' được gộp thành 1 đoạn", () => {
    const para = blocks.find(
      (b) => b.kind === "paragraph" && b.runs.map((r) => r.text).join("").includes("Căn cứ Thông tư")
    ) as any;
    assert.ok(para, "không thấy đoạn Căn cứ");
    const text = para.runs.map((r: any) => r.text).join("");
    assert.match(text, /của Bộ trưởng Bộ Giáo dục/);
    assert.equal(para.role, "body");
  });

  check("Mục 'I. ...' thành heading", () => {
    const h = blocks.find((b) => b.kind === "paragraph" && b.role === "heading");
    assert.ok(h);
    assert.match((h as any).runs.map((r: any) => r.text).join(""), /MỤC ĐÍCH/);
  });

  check("Bảng dữ liệu: 3 hàng x 3 cột, giữ đúng nội dung", () => {
    const table = blocks.find((b) => b.kind === "table" && b.role === "data") as any;
    assert.ok(table, "không thấy bảng data");
    assert.equal(table.rows.length, 3);
    assert.equal(table.rows[0].length, 3);
    const cell = (r: number, c: number) => table.rows[r][c].map((l: any) => l.map((x: any) => x.text).join("")).join("");
    assert.equal(cell(0, 1), "Nội dung công việc");
    assert.equal(cell(1, 1), "Dạy học theo chương trình");
    assert.equal(cell(2, 2), "Tuần 9");
  });

  check("Khối nơi nhận/chữ ký thành bảng signature không viền", () => {
    const sig = blocks.find((b) => b.kind === "table" && b.role === "signature") as any;
    assert.ok(sig, "không thấy bảng signature");
    const all = sig.rows.map((r: any) => r.map((c: any) => c.map((l: any) => l.map((x: any) => x.text).join("")).join("")).join(" | ")).join("\n");
    assert.match(all, /Nơi nhận/);
    assert.match(all, /HIỆU TRƯỞNG/);
  });

  check("Phân loại PDF: toàn text / toàn scan / hỗn hợp", () => {
    const page = sampleAdminPage();
    assert.equal(classifyPdf([page]).kind, "text");
    assert.equal(classifyPdf([{ ...page, source: "ocr" }]).kind, "scanned");
    assert.equal(classifyPdf([page, { ...page, pageNumber: 2, source: "ocr" }]).kind, "mixed");
  });

  check("Bộ lọc lỗi font: bắt dạng sai ánh xạ (dấu | dày đặc, cụm dính) — mẫu tổng hợp", () => {
    // Mẫu tổng hợp mô phỏng PDF in từ hệ thống trường học bị lỗi font nhúng:
    // chữ cái thường còn đọc được nhưng dấu biến thành ký hiệu lạ.
    const junk =
      "Z| M 8 on > s 5 = < g 9 s5igsls|Z|5|5|)3|s| £z|ls| 53|lia i|A|5| 3a < |291120559).|H| 6 |D| |D|D|D|D|D|D|T| | J5I.||IlS22)zZ | wu or «@ on . s 2 " +
      "l3 |V MCh ~~ J12/092015| |H|S6|D|D|D| [l4 LIyThiCh ~~ J01042015|X|H [6 |T|D|D| |5 |ThThCh ~~ 02042015 X | H | 6 |D|D|D| | 6 [LyThiChta ~~ [07/092015|X | H [5 |D|D|D| | 7 [VeMiDung ~~ [26/06/2015] |[H [6 |D|D|D| | 8 [Sung ThiGia ~~ [22042015 |X| T [7 |T|T [T |T |T |T|";
    const v = pageTextLooksGarbled([item(junk, 0, 0, 100)]);
    assert.equal(v.garbled, true, "phải bắt được dạng rác này: " + v.reason);
  });

  check("Bộ lọc lỗi font: bảng điểm nhiều ô 1 ký tự + item khoảng trắng của pdf.js KHÔNG bị báo oan", () => {
    // Tái hiện đúng dữ liệu app nhận từ pdf.js với bảng tổng hợp thật:
    // xen kẽ item chỉ chứa khoảng trắng, các ô điểm là ký tự đơn (Đ/T/số)
    // nhưng tiêu đề và họ tên còn nguyên dấu tiếng Việt → lớp chữ TỐT.
    const items: PdfTextItem[] = [];
    const push = (t: string) => items.push(item(t, 0, 0, 10));
    push("BẢNG TỔNG HỢP KẾT QUẢ ĐÁNH GIÁ GIÁO DỤC");
    push(" ");
    push("Phẩm chất chủ yếu");
    push("Năng lực chung");
    push("Năng lực đặc thù");
    const names = [
      "Nguyễn Văn An", "Trần Thị Bích Ngọc", "Lý Văn Cường", "Hoàng Thị Dung",
      "Vũ Minh Đức", "Phạm Thu Hằng", "Đặng Quốc Bảo", "Bùi Thanh Tùng",
      "Đỗ Hải Yến", "Ngô Văn Phúc", "Trịnh Ngọc Anh", "Lê Hồng Sơn",
    ];
    names.forEach((n, i) => {
      push(String(i + 1));
      push(" ");
      push(n);
      push(" ");
      push("12/09/2015");
      push(" ");
      for (let c = 0; c < 14; c++) {
        push(c % 3 === 0 ? "T" : "Đ");
        push(" ");
      }
      push("H");
      push(" ");
    });
    const v = pageTextLooksGarbled(items);
    assert.equal(v.garbled, false, "bảng điểm xịn không được báo lỗi font: " + v.reason);
  });

  check("Bộ lọc lỗi font: chữ vỡ từng ký tự + mất dấu thì vẫn bị bắt", () => {
    // Font lỗi thật: mỗi ký tự một mẩu, dấu biến thành ký hiệu/số lạ.
    const items: PdfTextItem[] = [];
    const junk = "K h o a n 2 d i e u 3 S Y R f a U O O A 1 u | | 5 9 ) . H 6 D D D D x w q z 1 2 3 4 5 6 7 8 9 0 a b c d e f g h i j k l m n o p q r s t u v w x y z A B C D E F G H I J K L M N O P Q R S T U V W X Y Z 1 2 3 4 5 6 7 8 9 0 | ( ) [ ] { } < > = + * # @ ~ ^ ` \\ / ; : , . ? ! % & $ ".repeat(3);
    junk.split(" ").filter(Boolean).forEach((t) => items.push(item(t, 0, 0, 5)));
    const v = pageTextLooksGarbled(items);
    assert.equal(v.garbled, true, "phải bắt được chữ vỡ mất dấu: " + v.reason);
  });

  check("Bộ lọc lỗi font: KHÔNG báo oan văn bản tiếng Việt sạch có bảng điểm số", () => {
    const clean =
      "BẢNG TỔNG HỢP KẾT QUẢ ĐÁNH GIÁ GIÁO DỤC HỌC KỲ II Năm học 2025 - 2026 Môn: Tin học Phẩm chất chủ yếu Năng lực chung Năng lực đặc thù " +
      "STT Họ và tên Ngày sinh Giới tính Lớp Kết quả Ghi chú 1 Nguyễn Văn An 12/09/2015 Nam 5A Đạt Hoàn thành tốt 2 Trần Thị Bích Ngọc 01/04/2015 Nữ 5A Đạt Cần cố gắng thêm " +
      "3 Lý Văn Cường 02/04/2015 Nam 5B Tốt Xuất sắc 4 Hoàng Thị Dung 07/09/2015 Nữ 5B Đạt 5 Vũ Minh Đức 26/06/2015 Nam 5C Tốt " +
      "Ghi chú: Đ = Đạt, T = Tốt. Tổng số học sinh được đánh giá: 632 em. Tỷ lệ hoàn thành tốt đạt 68,5% so với kế hoạch đề ra đầu năm học.";
    const v = pageTextLooksGarbled([item(clean, 0, 0, 100)]);
    assert.equal(v.garbled, false, "không được báo oan: " + v.reason);
  });

  check("Trang scan rỗng (OCR không ra chữ) không sinh block rác", () => {
    const emptyScan: PdfPageInput = { pageNumber: 1, width: 595, height: 842, items: [], source: "ocr" };
    assert.equal(pdfPagesToBlocks([emptyScan]).length, 0);
    assert.equal(classifyPdf([emptyScan]).kind, "scanned");
  });

  const buf = await buildDocxBuffer(blocks, CONFIG);
  assert.equal(buf.subarray(0, 2).toString(), "PK");
  assert.ok(buf.length > 1024);
  writeFileSync("/tmp/test-output.docx", buf);
  passed++;
  console.log("  ✓ File .docx mẫu đã ghi /tmp/test-output.docx (" + buf.length + " bytes)");
  void raw;

  console.log(`\nTất cả ${passed} kiểm tra ĐẠT.`);
}

main().catch((e) => {
  console.error("TEST THẤT BẠI:", e);
  process.exit(1);
});
