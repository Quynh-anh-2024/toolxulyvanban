# Hướng dẫn chạy & test bản mới — thêm chuyển PDF → Word chuẩn NĐ 30

## Chạy app

1. Giải nén zip vào một thư mục.
2. Mở terminal tại thư mục đó, chạy `npm install` (bản này có thêm thư viện
   mới: `pdfjs-dist`, `tesseract.js`, `docx`, `xlsx` — cần cài lại một lần).
3. Chạy `npm run dev` rồi mở http://localhost:3000

## Test nhanh tính năng PDF (5 phút)

Trong thư mục `tests/fixtures/` có sẵn 2 file mẫu:

1. **cong-van-text.pdf** — PDF có lớp chữ (kiểu công văn):
   - Tải lên → panel báo "PDF có lớp chữ (bản điện tử)", 1 trang, 0 trang OCR.
   - Xem trước thấy KẾ HOẠCH + bảng "Nội dung công việc".
   - Bấm **Tải File Word (.docx) Chuẩn NĐ 30** → mở file tải về bằng Word:
     kiểm tra lề trái 3cm, quốc hiệu/tiêu ngữ đúng, bảng có viền,
     khối chữ ký không viền.
2. **thong-bao-scan.pdf** — PDF scan (toàn ảnh, không có lớp chữ):
   - Tải lên → app tự OCR (lần đầu cần mạng để tải dữ liệu tiếng Việt,
     chờ khoảng 30–60 giây) → panel báo "PDF scan", 1 trang OCR,
     kèm cảnh báo soát lại chữ.
   - Nội dung đọc được: THÔNG BÁO nghỉ lễ Quốc khánh.

Sau đó thử với PDF thật của trường: công văn có bảng biểu, quyết định
nhiều trang, bản scan từ máy photocopy.

## Test thêm tính năng Excel (file có bảng tính)

Trong `tests/fixtures/` có sẵn **bang-tinh-mau.xlsx** (2 sheet):

3. **Excel → Word:** tải file .xlsx lên → mỗi sheet thành một bảng trong
   bản xem trước → bấm **Tải File Word (.docx) Chuẩn NĐ 30** để lấy Word.
   App cũng nhận .xls và .csv.
4. **Bảng trong PDF/Excel → Excel:** sau khi tải một PDF có bảng (hoặc
   một file Excel), bấm **Tải Excel (.xlsx) các bảng đã trích** ở mục 5 —
   nhận được file .xlsx thật, mỗi bảng một sheet (Bang 1, Bang 2...).

## Đã kiểm chứng bằng PDF THẬT trên internet (21/10/2026)

Tải 4 văn bản thật công khai và chạy qua app, đối chiếu từng từ với bản gốc
bằng công cụ độc lập (PyMuPDF + python-docx):

| File thật | Kết quả |
| --- | --- |
| Quyết định 2518/QĐ-UBND tỉnh Phú Thọ (5 trang, có bảng vốn đầu tư) | Phủ 99,1% chữ (phần thiếu duy nhất là dòng thông tin CHỮ KÝ SỐ — lớp chữ nằm trong ô chữ ký điện tử, mọi trình chuyển đổi đều không lấy được); bảng vốn 6×5 nhận đúng |
| Danh sách số báo danh/phòng thi THPT Phù Cừ (17 trang, bảng 9 cột lặp mỗi trang) | Phủ **100%** chữ, 0 ký tự lỗi; mỗi trang một bảng 9 cột 37–50 hàng, gộp cả hàng tiêu đề bảng |
| Kế hoạch giáo dục TH Khoái Châu 2 (29 trang, font subset, nhiều bảng lồng cột SL/%) | Phủ 99,8% chữ, 0 ký tự lỗi; nhận đúng cả bảng 14 cột (đánh giá viên chức) |
| Bản scan tạo từ 3 trang văn bản thật ở trên | OCR tiếng Việt phủ ~85% số từ (giới hạn tự nhiên của OCR với bản scan; app có cảnh báo soát lại) |

Các lỗi đã sửa sau vòng test này (đều là lỗi chỉ lộ khi gặp PDF thật):
1. Chữ dính liền không dấu cách giữa các cột khi một dòng có nhiều đoạn
   ("STTHọ và tên") — nguyên nhân hàng đầu gây cảm giác "lỗi ký tự".
2. Bộ nhận diện bảng cũ vỡ trên bảng thật (hàng có ô trống/ô cuốn dòng) —
   đã viết lại theo vùng + mô hình cột, gán ô theo trục x.
3. Hai dòng tiêu đề đầu trang bị nhận nhầm thành bảng 2×2 (vì từ khoá
   "TRƯỜNG" quá rộng và dòng tiêu đề bị hút vào bảng).
4. Gạch nối cuối dòng trong từ ghép ("quốc phòng-an ninh") bị xoá làm dính
   chữ — nay giữ nguyên gạch nối.
5. Chuẩn hoá Unicode NFC cho mọi chữ trích từ PDF (chống lỗi dấu khi mở Word).
6. Trang có lớp chữ nhưng lỗi font/mã hoá (ký tự rác) → tự động OCR lại thay
   vì xuất chữ rác.
7. OCR: render 3.5x + tăng tương phản (đo thực tế cho độ phủ từ cao hơn).

Phần CHUẨN HOÁ Word (docx → .doc) cũng đã rà và sửa 4 lỗi nghiêm trọng:
quốc hiệu bị để 10,8pt (chuẩn 12–13pt), tiêu ngữ 12,5pt (chuẩn 13–14pt);
bảng dữ liệu chứa chữ "Trường" bị nhận nhầm thành bảng đầu văn bản và bị
viết lại; dòng "SỞ GIÁO DỤC" bị nhận nhầm thành số ký hiệu (vì cùng bắt đầu
bằng "SO"); và lỗi bịa số ký hiệu mẫu "Số: /KH-..." khi không tách được số
thật — nay không tìm thấy thì để trống. Bộ test: `tests/ui-normalize-test.py`
(13 kiểm tra, đạt hết) và `tests/real-test.mts` + `tests/verify_docx.py`
(cần tự tải PDF thật vào `tests/real/`).

## PWA — cài ra màn hình chính & dùng OFFLINE (thêm 21/10/2026)

App đã có PWA đầy đủ: manifest + service worker (`public/sw.js`) + bộ icon
riêng. Sau khi deploy lên địa chỉ HTTPS (hoặc chạy localhost):

- **Android/Chrome**: mở app → menu ⋮ → "Cài đặt ứng dụng" / "Thêm vào màn
  hình chính" → icon "Xử Lý Văn Bản" xuất hiện như app thường, mở toàn màn
  hình không thanh địa chỉ.
- **iPhone/Safari**: nút Chia sẻ → "Thêm vào Màn hình chính".
- **Máy tính (Chrome/Edge)**: có biểu tượng cài đặt ở cuối thanh địa chỉ.

**Dùng offline**: sau lần mở đầu tiên CÓ mạng (và nên mở thử 1 file scan để
app tải sẵn dữ liệu OCR tiếng Việt vào bộ nhớ đệm, ~60MB), các lần sau
**không cần mạng**: chuyển PDF/Excel → Word và OCR bản scan đều chạy ngay
trên máy. Riêng tính năng AI (soát lỗi, tóm tắt, tạo bài tập) vẫn cần mạng
vì gọi API. Đã kiểm chứng bằng test tự động `tests/pwa-offline-test.py`:
ngắt mạng hoàn toàn → app vẫn mở, xuất được .docx, OCR vẫn đọc được.

Lưu ý: khi có bản app mới, service worker tự cập nhật ở lần mở sau; nếu thấy
giao diện cũ, đóng app mở lại 2 lần.

## Bản vá 21/10/2026 (chiều) — PDF in từ hệ thống trường học bị OCR oan

Triệu chứng đã gặp thật: nạp bảng tổng hợp kết quả đánh giá (PDF có lớp chữ
hoàn hảo) thì bản xem trước và file xuất ra toàn ký tự rác. Nguyên nhân:
pdf.js chèn nhiều mẩu chỉ chứa khoảng trắng giữa các ô; bộ lọc "trang lỗi
font" đếm cả chúng nên một bảng điểm bình thường (hầu hết ô là 1 ký tự:
Đ, T, số) bị kết luận nhầm là "chữ vỡ từng ký tự" → app OCR oan → rác.

Đã sửa:
- Bộ lọc chỉ xét các mẩu chữ có nội dung; luật "chữ tách rời" phải kèm
  bằng chứng mất gần hết dấu tiếng Việt mới kết luận lỗi font.
- Đóng gói thêm `public/cmaps/`, `public/standard_fonts/`, `public/wasm/`
  và khai báo cho pdf.js — tài nguyên pdf.js cần trên trình duyệt để giải
  mã đúng các font hệ thống (ở môi trường Node chúng có sẵn nên test Node
  không phát hiện được thiếu sót này).
- Thêm test trình duyệt thật `tests/browser-real-pdf-test.py` chạy với
  PDF thật từ hệ thống trường học (để trong `tests/real/`, không đóng gói
  cùng bản phát hành vì có dữ liệu học sinh): xem trước phải đọc thẳng lớp
  chữ, file Excel/Word tải xuống phải sạch. Từ nay, kết luận về PDF hệ
  thống chỉ dựa trên test trình duyệt này, không dựa vào test Node.

Lưu ý: CV 5095/BGDĐT là bản scan thật (mỗi trang là 1 ảnh, lớp chữ chỉ có
thông tin con dấu ký số) nên app đi đường OCR là ĐÚNG — chất lượng chữ
phụ thuộc độ nét của bản scan gốc.

## Kiểm thử tự động (nếu muốn)

- `npm run test:pdf` — chạy test logic + test đầu-cuối (cần `npx tsx`,
  đã có sẵn trong devDependencies).
- Kết quả đối chiếu khi bàn giao: test đơn vị 10/10, test đầu-cuối PDF 7/7,
  test Excel 5/5; test giao diện bằng trình duyệt thật cả 3 luồng đều đạt
  và không lỗi JS: (1) PDF có lớp chữ → panel đúng loại, tải .docx chuẩn;
  (2) PDF scan → OCR ngay trong trình duyệt, tin cậy 95% trên file mẫu;
  (3) Excel → xem trước đủ 2 sheet, tải được cả .docx lẫn .xlsx.
- OCR dùng dữ liệu đóng gói sẵn trong thư mục `public/` (worker, wasm,
  traineddata vie/eng) — chạy offline, không cần mạng.

## Lưu ý

- PDF đặt mật khẩu: cần gỡ mật khẩu trước khi tải lên.
- Chữ ký/con dấu trên trang scan là hình ảnh — OCR không chuyển thành chữ.
- Đậm/nghiêng lẻ trong thân bài PDF có thể mất (pdf.js không đọc được
  thuộc tính này của font subset); tiêu đề, trích yếu, các mục và hàng đầu
  bảng vẫn được áp đậm đúng thể thức khi xuất Word.
