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

  check("Phần đầu KHÔNG nuốt thân văn bản khi ngày tháng dự thảo trống số ngày (ca Sở GDĐT Tuyên Quang)", () => {
    // Mô phỏng công văn HƯỚNG DẪN: dòng ngày tháng "ngày   tháng 10 năm
    // 2026" (dự thảo, chưa điền số ngày) → DATE_RE không khớp. Trước khi
    // vá, phần đầu nuốt 14 dòng đầu: V/v, Kính gửi và đoạn Căn cứ bị viết
    // hoa toàn bộ và ép vào ô trái hẹp của bảng đầu.
    const items: PdfTextItem[] = [
      item("ỦY BAN NHÂN DÂN TỈNH TUYÊN QUANG", 40, 40, 205, { bold: true }),
      item("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", 315, 40, 245, { bold: true }),
      item("SỞ GIÁO DỤC VÀ ĐÀO TẠO", 62, 58, 175, { bold: true }),
      item("Độc lập - Tự do - Hạnh phúc", 348, 58, 178, { bold: true }),
      item("Số: 123/SGDĐT-GDPT", 72, 76, 148),
      item("Tuyên Quang, ngày tháng 10 năm 2026", 330, 76, 222, { italic: true }),
      item("V/v hướng dẫn kiểm tra, đánh giá định kỳ từ năm học 2026-2027", 108, 102, 382, { bold: true }),
      item("Kính gửi: Các cơ sở giáo dục phổ thông, giáo dục thường xuyên trên địa bàn tỉnh", 70, 128, 400),
      item("Căn cứ Thông tư số 27/2020/TT-BGDĐT ngày 04/9/2020 của Bộ trưởng Bộ Giáo dục và Đào tạo quy định đánh giá", 70, 152, 455),
      item("học sinh tiểu học; Thông tư số 22/2021/TT-BGDĐT ngày 20/7/2021 của Bộ trưởng Bộ Giáo dục và Đào tạo quy định về đánh giá", 70, 168, 455),
    ];
    const page: PdfPageInput = { pageNumber: 1, width: 595, height: 842, items, source: "text" };
    const bs = pdfPagesToBlocks([page]);
    const raw = blocksToRawText(bs);
    const header = bs[0] as Extract<DocBlock, { kind: "table" }>;
    assert.equal(header.kind, "table");
    assert.equal((header as any).role, "adminHeader");
    const headerText = header.rows
      .flat()
      .flat()
      .map((line) => line.map((r) => r.text).join(""))
      .join("\n");
    assert.ok(!headerText.includes("Kính gửi"), "header nuốt dòng Kính gửi: " + headerText);
    assert.ok(!headerText.includes("Căn cứ"), "header nuốt đoạn Căn cứ: " + headerText);
    assert.ok(!headerText.includes("V/v"), "header nuốt dòng trích yếu: " + headerText);
    assert.ok(raw.includes("Kính gửi: Các cơ sở giáo dục phổ thông"), "mất dòng Kính gửi ở thân văn bản");
    assert.ok(raw.includes("Căn cứ Thông tư số 27/2020/TT-BGDĐT"), "mất đoạn Căn cứ ở thân văn bản");
    assert.ok(raw.includes("V/v hướng dẫn kiểm tra"), "mất dòng trích yếu");
    assert.ok(!raw.includes("CĂN CỨ THÔNG TƯ"), "thân văn bản bị viết hoa oan (dấu hiệu bị nuốt vào ô header)");
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

  // --- Hồi quy vụ "cơ chế quét" (22/10/2026): trang scan phải giữ nguyên
  // dòng/đoạn của Tesseract — đoạn văn căn đều KHÔNG được biến thành bảng,
  // chữ không bị xáo trộn/nuốt; tên cơ quan ở dòng bị gộp 2 cột không bị
  // quốc hiệu nuốt; khối chữ ký chen trong vùng Nơi nhận phải tách ra sau.
  check("Trang scan: đoạn văn giữ nguyên thứ tự, không thành bảng, header đủ 2 cột", () => {
    const oitem = (text: string, x: number, y: number, w: number, lineId: number, paraId: number): PdfTextItem => ({
      text: text + " ",
      x, y, width: w, height: 13, fontSize: 13,
      bold: false, italic: false, fontName: "ocr", conf: 95, lineId, paraId,
    });
    const items: PdfTextItem[] = [
      // Dòng đầu bị Tesseract gộp 2 cột: cơ quan + quốc hiệu (khe 33pt).
      oitem("BỘ", 75, 40, 20, 1, 1), oitem("GIÁO", 98, 40, 25, 1, 1), oitem("DỤC", 126, 40, 22, 1, 1),
      oitem("VÀ", 151, 40, 16, 1, 1), oitem("ĐÀO", 170, 40, 22, 1, 1), oitem("TẠO", 195, 40, 20, 1, 1),
      oitem("CỘNG", 248, 40, 30, 1, 1), oitem("HÒA", 281, 40, 25, 1, 1), oitem("XÃ", 309, 40, 15, 1, 1),
      oitem("HỘI", 327, 40, 22, 1, 1), oitem("CHỦ", 352, 40, 22, 1, 1), oitem("NGHĨA", 377, 40, 32, 1, 1),
      oitem("VIỆT", 412, 40, 25, 1, 1), oitem("NAM", 440, 40, 22, 1, 1),
      oitem("Độc", 330, 58, 20, 2, 1), oitem("lập", 353, 58, 16, 2, 1), oitem("Tự", 380, 58, 14, 2, 1),
      oitem("do", 397, 58, 14, 2, 1), oitem("Hạnh", 425, 58, 25, 2, 1), oitem("phúc", 453, 58, 25, 2, 1),
      oitem("Số:", 75, 84, 18, 3, 2), oitem("5095/BGDĐT-NGCBQLGD", 96, 84, 110, 3, 2),
      oitem("Hà", 270, 84, 16, 3, 2), oitem("Nội,", 289, 84, 22, 3, 2), oitem("ngày", 314, 84, 25, 3, 2),
      oitem("04", 342, 84, 14, 3, 2), oitem("tháng", 359, 84, 28, 3, 2), oitem("08", 390, 84, 14, 3, 2),
      oitem("năm", 407, 84, 20, 3, 2), oitem("2026", 430, 84, 24, 3, 2),
      // Đoạn thân 3 dòng (y rung ±1, có khe căn đều 20pt giữa dòng).
      oitem("Thực", 110, 140, 24, 4, 3), oitem("hiện", 137, 141, 22, 4, 3), oitem("ý", 162, 140, 8, 4, 3),
      oitem("kiến", 173, 140, 24, 4, 3), oitem("chỉ", 200, 141, 16, 4, 3), oitem("đạo", 219, 140, 20, 4, 3),
      oitem("của", 242, 140, 18, 4, 3), oitem("Phó", 263, 140, 18, 4, 3), oitem("Thủ", 284, 141, 18, 4, 3),
      oitem("tướng", 305, 140, 30, 4, 3), oitem("Chính", 355, 140, 26, 4, 3), oitem("phủ", 384, 140, 18, 4, 3),
      oitem("tại", 405, 141, 16, 4, 3),
      oitem("Công", 75, 158, 26, 5, 3), oitem("văn", 104, 158, 20, 5, 3), oitem("số", 127, 159, 14, 5, 3),
      oitem("4283/VPCP-KGVX", 144, 158, 95, 5, 3), oitem("ngày", 242, 158, 25, 5, 3),
      oitem("12/5/2026,", 270, 158, 55, 5, 3), oitem("Bộ", 328, 159, 16, 5, 3), oitem("Giáo", 347, 158, 24, 5, 3),
      oitem("dục", 374, 158, 18, 5, 3), oitem("và", 395, 158, 16, 5, 3), oitem("Đào", 414, 159, 22, 5, 3),
      oitem("tạo", 439, 158, 18, 5, 3), oitem("phối", 460, 158, 22, 5, 3), oitem("hợp", 485, 158, 18, 5, 3),
      oitem("với", 75, 176, 20, 6, 3), oitem("các", 98, 176, 20, 6, 3), oitem("đơn", 121, 177, 20, 6, 3),
      oitem("vị", 144, 176, 12, 6, 3), oitem("có", 159, 176, 16, 6, 3), oitem("liên", 178, 176, 22, 6, 3),
      oitem("quan", 203, 177, 24, 6, 3), oitem("xây", 230, 176, 20, 6, 3), oitem("dựng", 253, 176, 24, 6, 3),
      oitem("dự", 280, 176, 14, 6, 3), oitem("thảo", 297, 177, 24, 6, 3), oitem("Nghị", 324, 176, 22, 6, 3),
      oitem("quyết", 349, 176, 26, 6, 3), oitem("này.", 378, 176, 24, 6, 3),
      // Vùng Nơi nhận bị chữ ký chen segment bên phải vào cùng dòng.
      oitem("Nơi", 75, 640, 20, 7, 4), oitem("nhận:", 98, 640, 28, 7, 4),
      oitem("BỘ", 345, 640, 18, 7, 4), oitem("TRƯỞNG", 366, 640, 40, 7, 4),
      oitem("-", 75, 658, 6, 8, 4), oitem("Như", 84, 658, 20, 8, 4), oitem("trên;", 107, 658, 26, 8, 4),
      oitem("THƯỜNG", 345, 658, 38, 8, 4), oitem("TRỰC", 386, 658, 26, 8, 4),
      oitem("-", 75, 676, 6, 9, 4), oitem("Lưu:", 84, 676, 22, 9, 4), oitem("VT,", 109, 676, 18, 9, 4),
      oitem("NGCBQLGD.", 130, 676, 58, 9, 4),
    ];
    const page: PdfPageInput = { pageNumber: 1, width: 595, height: 842, items, source: "ocr" };
    const blks = pdfPagesToBlocks([page]);
    const rawAll = blocksToRawText(blks);
    // 1) Không có bảng dữ liệu nào sinh ra từ đoạn văn scan.
    assert.ok(!blks.some((b) => b.kind === "table" && b.role === "data"), "đoạn văn scan bị biến thành bảng");
    // 2) Header: bảng adminHeader và tên cơ quan còn nguyên (không bị quốc hiệu nuốt).
    const header = blks[0];
    assert.ok(header.kind === "table" && header.role === "adminHeader", "thiếu bảng đầu trang");
    assert.ok(rawAll.includes("BỘ GIÁO DỤC VÀ ĐÀO TẠO"), "mất tên cơ quan ban hành");
    assert.ok(rawAll.includes("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"), "mất quốc hiệu");
    // 3) Đoạn thân nguyên vẹn, đúng thứ tự từng chữ.
    const body = blks.find((b) => b.kind === "paragraph" && b.runs.map((r) => r.text).join("").includes("Thực hiện"));
    assert.ok(body && body.kind === "paragraph", "không thấy đoạn thân");
    const bodyText = body.kind === "paragraph" ? body.runs.map((r) => r.text).join("") : "";
    assert.ok(
      bodyText.includes("Thực hiện ý kiến chỉ đạo của Phó Thủ tướng Chính phủ tại Công văn số 4283/VPCP-KGVX ngày 12/5/2026, Bộ Giáo dục và Đào tạo phối hợp với các đơn vị có liên quan xây dựng dự thảo Nghị quyết này."),
      "đoạn thân bị xáo trộn/mất chữ: " + bodyText
    );
    // 4) Chữ ký tách khỏi Nơi nhận và đứng sau toàn bộ danh sách nơi nhận.
    const recip = blks.find((b) => b.kind === "paragraph" && b.runs.map((r) => r.text).join("").startsWith("Nơi nhận"));
    assert.ok(recip && recip.kind === "paragraph", "không thấy đoạn Nơi nhận");
    const recipText = recip.kind === "paragraph" ? recip.runs.map((r) => r.text).join("") : "";
    assert.ok(!recipText.includes("TRƯỞNG"), "chữ ký còn chen trong Nơi nhận: " + recipText);
    assert.ok(rawAll.indexOf("Lưu:") < rawAll.indexOf("BỘ TRƯỞNG"), "khối chữ ký phải đứng sau Nơi nhận");
    assert.ok(rawAll.indexOf("Lưu:") < rawAll.indexOf("THƯỜNG TRỰC"), "dòng chức danh ký phải đứng sau Nơi nhận");
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
