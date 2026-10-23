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

## Bản vá 21/10/2026 (tối) — Lỗi "undefined is not a function" trên iPhone/Safari cũ

Triệu chứng đã gặp thật: mở app trên iPhone, nạp PDF nào cũng báo
"Không thể chuyển đổi PDF này: undefined is not a function (near
'...i of t...')", trong khi trên máy tính/Chrome chạy bình thường.
Nguyên nhân: pdf.js v6 dùng nhiều API JavaScript rất mới mà Safari cũ
chưa có — điển hình là duyệt ReadableStream bằng `for await...of`
(Safari 18.4+), `Promise.withResolvers` (17.4+), `.at()` (15.4+) và
`replaceAll` khi dịch font. Hàm đọc chữ của pdf.js (`getTextContent`)
lặp stream đúng kiểu mới nên PDF nào cũng chết trên Safari cũ.

Đã sửa:
- Chuyển sang pdf.js BẢN LEGACY (có sẵn một phần polyfill core-js).
- Thêm `src/utils/browserPolyfills.ts` vá đúng các API còn thiếu
  (ReadableStream async iterator, Promise.withResolvers/try, .at,
  replaceAll), cài TRƯỚC khi pdf.js nạp, ở cả main thread lẫn worker
  (qua `src/workers/pdfWorker.ts` + `GlobalWorkerOptions.workerPort`),
  gắn kiểu non-enumerable như API bản địa.
- Service worker đổi cache sang `xlvb-v3` để điện thoại tự bỏ bản cũ.
- Thêm `tests/repro-old-safari.py`: giả lập Safari cũ bằng cách gỡ toàn
  bộ API mới cùng lúc mà vẫn chuyển đổi THÀNH CÔNG cả PDF có lớp chữ
  lẫn PDF scan/OCR mới đạt. Lưu ý khi test tay trên iPhone đã cài app ra
  màn hình chính: mở lại app 1–2 lần để bản mới được nạp; nếu vẫn hiện
  bản cũ thì xoá app khỏi màn hình chính rồi thêm lại.

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

## Cập nhật 22/10/2026 — sửa cơ chế quét (OCR) + tách mục giao diện

- **Sự cố**: bản Word từ PDF scan (CV 5095/BGDĐT) ra lò không dùng được:
  đoạn văn bị xáo trộn thứ tự, mất cả dòng chữ, tên cơ quan ban hành bị
  mất, khối chữ ký chen vào giữa "Nơi nhận" — dù Tesseract đọc đúng ~89%.
- **Nguyên nhân**: tầng dựng cấu trúc dùng chung máy tách cột/bảng của PDF
  có lớp chữ cho trang scan: khe căn đều > ~13pt bị coi là ranh cột → đoạn
  văn thành "bảng" giả; gán cột thất bại thì segment bị NUỐT mất; dòng bị
  Tesseract gộp 2 cột đầu trang làm mất tên cơ quan.
- **Cách sửa**: trang scan đi đường riêng — giữ nguyên dòng/đoạn do
  Tesseract phân đoạn (lineId/paraId), chỉ tách segment ở khe rất rộng
  (≥2,4× cỡ chữ), lọc từ rác theo độ tin cậy (<30; mảnh 1 ký tự <60),
  tách khối chữ ký khỏi vùng "Nơi nhận", bảng chỉ dựng khi cột khớp rất
  chặt (≥4 dòng, số cột bằng nhau, lệch ≤14pt). PDF có lớp chữ giữ nguyên
  đường cũ. Đo lại trên CV 5095 trang 1 (bản chép tay đối chiếu): tầng cấu
  trúc mất 5% → 0%; Word tải từ trình duyệt phủ 86,6% → 91,1%, đoạn văn
  nguyên vẹn đúng thứ tự. Phần sai còn lại là Tesseract đọc sai dấu trên
  bản scan cổng (trần của OCR offline), app có cảnh báo soát lại.
- **Giao diện**: một màn hình kiểu app phiếu học tập (thanh trên + cột
  công cụ/cột xem trước tự cuộn, điện thoại có tab Công cụ/Xem trước) với
  3 MỤC RIÊNG: "PDF → Word" (chỉ PDF → .docx), "Excel" (bảng tính →
  Word/Excel + trích bảng từ PDF → Excel), "Word & AI" (chuẩn hoá .docx
  + soát lỗi AI). Mỗi mục giữ kết quả riêng, chuyển mục không mất việc.
- Kiểm chứng: unit 15/15 (thêm test hồi quy trang scan), e2e 7/7, Excel
  5/5, test trình duyệt PDF thật đạt (kể cả luồng Excel mới), UI smoke,
  chuẩn hoá Word, PWA offline đều đạt.

## Cập nhật 22/10/2026 (chiều) — Hai cách đọc trang scan: AI ĐÁNH LẠI / QUÉT OCR

Theo yêu cầu: đẩy 1 file PDF lên thì trang scan (không có lớp chữ) được
đọc bằng MỘT trong hai cách, chọn ngay trong mục "PDF → Word":

1. **AI đánh lại (mặc định, khuyên dùng cho bản scan):** app gửi ẢNH trang
   cho Gemini (model gemini-3.8-flash) chép lại từng chữ như người đánh
   máy, rồi dựng cấu trúc (đầu văn bản 36/64, tiêu đề, bảng markdown →
   bảng thật, Nơi nhận, khối ký) bằng `src/utils/aiText.ts`.
   - Đo thật trên CV 5095/BGDĐT trang 1 (bản scan cổng ký số): AI chép
     đúng **100,0%** số từ so với bản chép tay kiểm chứng; Tesseract chỉ
     đạt ~89%. Đây là mức tương đương ABBYY FineReader (99,8%).
   - Cần: internet + **khóa Gemini của chính người dùng** (miễn phí tại
     aistudio.google.com/apikey), dán MỘT lần vào ô "Khóa Gemini" trong
     app. Khóa chỉ lưu ở localStorage của trình duyệt trên máy đó —
     KHÔNG viết vào mã nguồn (khóa OpenRouter cũ viết cứng trong
     `src/services/ai.ts` đã bị OpenRouter thu hồi — 401 "User not
     found" — vì repo công khai; các tính năng AI chữ của mục Word & AI
     nay tự dùng khóa Gemini này làm đường dự phòng).
   - Trang nào AI lỗi (mất mạng/hết lượt) thì app TỰ rơi xuống quét OCR
     cho trang đó và ghi cảnh báo rõ trong panel kết quả — không treo,
     không mất cả file.
   - Trang có lớp chữ LUÔN đọc thẳng lớp chữ ở cả hai chế độ.
2. **Quét OCR trên máy:** Tesseract như trước — offline, không cần khóa,
   trần ~89–91% với scan cổng (sai dấu ở chỗ mờ là khó tránh).

Công cụ "Trích bảng từ PDF" ở mục Excel dùng chung cách đọc đã chọn ở
mục PDF → Word (chưa có khóa thì tự quét OCR).

Kèm theo: **vá lỗi xem trước trên điện thoại** — trang A4 bị cắt mất mép
trái trong khung cuộn (chữ đầu dòng biến mất như ảnh chụp của thầy Hải);
đã bọc trang trong lớp căn giữa đúng cách, kiểm thử ở khổ iPhone: cuộn
tới được cả hai mép.

Kiểm chứng đợt này: unit pdf-structure 15/15, ai-text 14/14, Excel 5/5,
e2e 7/7, trình duyệt + PDF thật (đường OCR) đạt, UI smoke/chuẩn hoá
Word/PWA offline đạt; chuỗi AI: đo qua API 100% → dựng blocks giữ
100,0% → test trình duyệt `tests/browser-ai-test.py` (nhánh dự phòng
đạt; nhánh AI cần mạng tới Google không bị chặn như trên máy người dùng).

## Cập nhật 22/10/2026 (tối) — Gọn giao diện + các sửa ưu tiên mục A

Theo yêu cầu của thầy Hải: giao diện KHÔNG còn đoạn văn giải thích ở bất
kỳ mục nào (chỉ nhãn, nút, trạng thái); chỗ lấy khóa Gemini là link bấm
được mở thẳng trang aistudio.google.com/apikey. Quy tắc này áp cho mọi
app làm cho thầy từ nay.

Các sửa ưu tiên (mục A trong bản phân tích chuyên gia 22/10):
- A1: chọn cách đọc trang scan thu thành MỘT hàng chọn nhỏ "Trang scan:
  [AI đánh lại | Quét OCR]" ngay dưới ô tải file; ô nhập khóa chỉ hiện
  khi chọn AI mà chưa lưu khóa; đã lưu khóa thì chỉ còn một dòng chữ nhỏ.
- A2: bản xem trước có nút "Vừa màn hình / 100%": chế độ vừa màn hình
  thu nhỏ trang A4 theo bề rộng khung (điện thoại xem trọn trang, không
  cuộn ngang); mặc định bật trên màn hình nhỏ.
- A3: bỏ toàn bộ hộp thoại alert() — lỗi/cảnh báo/thành công hiện bằng
  thanh báo trong trang ở cuối màn hình, phân màu đỏ/vàng/xanh, tự tắt
  sau 7 giây hoặc bấm dấu ×.
- A4: panel kết quả tách hai khối riêng: số liệu (nền xanh) và cảnh báo
  (nền vàng riêng, dòng cảnh báo nặng hiện chữ đỏ).

Kiểm chứng: A-check tự động trên trình duyệt (khổ iPhone) đạt cả A1/A2/
A3 — trang vừa màn hình đo đúng không tràn ngang, chế độ 100% trả đúng
794px, file PDF hỏng hiện thanh báo đỏ và không còn hộp thoại nào; test
giao diện + chuẩn hoá Word đạt; chụp màn hình đối chiếu trực quan đạt.

## Cập nhật 22/10/2026 (tối muộn) — Mục B: cấu trúc cốt lõi

- B5: App.tsx tách thành file riêng cho từng mục: src/sections/PdfSection
  .tsx, ExcelSection.tsx, WordSection.tsx; khung xem trước ở
  src/components/PreviewPane.tsx, thanh báo ở NoticeBar.tsx, khối căn lề
  ở MarginConfig.tsx. State + hàm xử lý vẫn tập trung ở App và truyền
  xuống qua props — sửa một mục không đụng mục khác. App.tsx còn ~700
  dòng (trước ~1.000).
- B6: nút "Kiểm tra" cạnh dòng "Đã lưu khóa Gemini": gọi thử AI một câu
  tối thiểu và báo ngay kết quả bằng thanh báo (khóa sống/model nào đáp,
  hoặc lỗi tiếng Việt dễ hiểu). Lỗi mạng thô của trình duyệt ("Failed to
  fetch") đã được dịch sang tiếng Việt trong services/ai.ts.
- B7: trang scan ở chế độ AI được gom lại và AI đánh lại SONG SONG tối
  đa 3 trang cùng lúc (có báo tiến độ "đã xong x/y trang"); trang nào AI
  lỗi vẫn tự rơi OCR cho trang đó, phần OCR dự phòng chạy tuần tự để
  không nặng máy.
- B8: gói mở app giảm từ ~2.320KB xuống ~430KB: các phần nặng (đọc PDF,
  Excel, xuất Word, đọc .docx) chỉ nạp khi dùng tới; app tự nạp trước
  lúc trình duyệt rảnh để PWA vẫn chạy offline trọn vẹn sau lần mở có
  mạng đầu tiên (đã kiểm lại bằng test offline).

Kiểm chứng: mọi bộ test cũ đều đạt lại (unit 15, AI-text 14, Excel 5,
đầu-cuối 7, PDF thật trên trình duyệt, AI dự phòng, giao diện, chuẩn hoá
Word, kiểm A1–A3, offline PWA) + 2 kiểm mới cho B6/B7 (nút Kiểm tra ra
đúng thanh báo; scan 3 trang rơi OCR đủ 3 trang qua đường song song).

## Cập nhật 22/10/2026 (tối) — Bỏ nút chọn cách đọc, xử lý TỰ ĐỘNG + đối chiếu

Theo yêu cầu của thầy Hải: bỏ hẳn hai nút "AI đánh lại / Quét OCR".
Upload file lên là app tự xử lý theo cách đã bàn:
- Trang có lớp chữ: đọc thẳng lớp chữ (như cũ, chuẩn tuyệt đối).
- Trang scan + ĐÃ có khóa Gemini: AI chép trực tiếp từ ảnh trang (kết
  quả chính), đồng thời OCR chạy nền cho cùng trang; xong thì ĐỐI CHIẾU
  hai bản — chỗ nào số liệu (số hiệu, ngày tháng, số tiền, số lượng) hai
  bên đọc khác nhau sẽ hiện cảnh báo gắn cờ đúng trang và đúng con số
  để kiểm tra trên bản gốc trước khi ban hành. Chữ thường khác dấu
  (điểm yếu quen thuộc của OCR) không gắn cờ để tránh báo nhiễu.
- Trang scan + CHƯA có khóa: quét OCR như cũ, kèm nhắc dán khóa để dùng AI.
- Trang nào AI lỗi vẫn tự rơi OCR cho trang đó (kết quả OCR nền dùng
  luôn, không quét lại lần hai).

Kiểm chứng: test đơn vị bộ đối chiếu 10/10 (sai số quyết định/số tiền
thì gắn cờ, khác dấu chữ thường thì không); kiểm tự động trên trình
duyệt: không còn nút chọn, chưa khóa OCR đủ 3/3 trang, khóa giả (AI lỗi)
rơi OCR đủ 3/3 trang kèm cảnh báo; các bộ test cũ (PDF thật, AI dự
phòng, giao diện, chuẩn hoá, offline PWA) đều đạt lại. Nhánh AI thật +
đối chiếu chỉ chạy được trên máy có mạng tới Google (môi trường kiểm
thử bị chặn) — thầy thử bằng một file scan thật sẽ thấy cảnh báo đối
chiếu nếu có số liệu đọc khác nhau.

## Đợt 23/10/2026 — sửa lỗi file CV 5095 tải xuống + giao diện TỐI GIẢN 1 màn hình

**Lỗi Hải báo (kèm file `CV_5095_BGD_ChuanND30__1.docx` anh tự tải): trang 1 đẹp,
trang 2–3 vỡ nát — bảng 2×2 giả ở đầu mỗi trang, mục 1 trang 2 biến mất, các biểu
mẫu bị ép thành đoạn văn đầy chữ rác OCR.**
- Nguyên nhân 1 (lỗi của app): khi tài liệu có trang AI chen trang OCR, khâu ghép
  gọi dựng cấu trúc TỪNG TRANG MỘT, mà mã dựng "phần đầu thể thức" lại căn theo vị
  trí trang trong mảng truyền vào → trang nào cũng thành "trang 1" → dựng bảng đầu
  giả + nuốt nội dung đầu trang. ĐÃ SỬA: chỉ số trang lấy theo số trang thật
  (`pageNumber`), không theo vị trí mảng. Kiểm hồi quy trong browser-real-pdf-test
  (đúng 1 bảng, mục 1 trang 2 còn nguyên).
- Nguyên nhân 2 (AI lỗi theo trang): trang 1 AI chép hoàn hảo còn trang 2–3 bị lỗi
  gọi AI (quá tải/giới hạn tạm thời) nên rơi xuống OCR — mà OCR với biểu mẫu nhiều
  bảng trống thì rất tệ (đo thực tế trên chính file này). ĐÃ SỬA: trang AI lỗi ở
  vòng song song được THỬ LẠI TUẦN TỰ một vòng (chờ 4 giây/trang) rồi mới xuống
  OCR; cảnh báo chỉ ghi sau khi thử lại vẫn lỗi. Đo bench bằng prompt+model của
  app trên trang 2–3: AI chép đúng gần như tuyệt đối, cả 5 bảng thành bảng thật
  (kể cả Biểu số 2 đủ 13 hàng) — file mẫu đã sửa: `CV_5095_BGD_ban_da_sua.docx`.
  Lưu ý: trang nhiều bảng AI trả lời chậm (~100 giây/trang) và thỉnh thoảng gặp
  503 quá tải — cứ để app tự thử lại; cảnh báo sẽ nêu rõ trang nào phải dùng OCR.

**Giao diện TỐI GIẢN (yêu cầu của Hải: "ẩn tất cả, thành cơ chế tự động, nhìn
lằng nhằng quá"):**
- Chỉ còn MỘT màn hình: ô "Tải lên file PDF" → app tự xử lý → 2 nút "Tải file
  Word (.docx)" / "Tải file Excel (.xlsx)". Ô khóa Gemini giữ nguyên (link bấm
  được; đã lưu thì hiện 1 dòng nhỏ Kiểm tra/Xóa). Cảnh báo theo file vẫn hiện
  ngay dưới ô tải (đây là trạng thái cần biết, không phải chữ giải thích).
- BỎ: bản xem trước A4, mục chỉnh thể thức/lề (mặc định khoá cứng chuẩn NĐ 30:
  Times New Roman 14, căn đều, thụt đầu dòng 1,27cm, lề trái 3cm, phải/trên/dưới
  2cm), mục "Word & AI", mục đọc file bảng tính Excel/CSV, panel số liệu kết quả.
  (Các module xử lý Word/AI/bảng tính vẫn còn trong mã nguồn, chỉ không còn giao
  diện gọi tới — cần khôi phục mục nào thì nói rõ.)
- Tên app đổi thành "Chuyển PDF sang Word & Excel" (tiêu đề trang + manifest);
  bộ nhớ đệm PWA lên `xlvb-v4`.
- Test đã sửa theo giao diện mới: ui_smoke (tải cả Word + Excel), browser-real
  (kiểm qua file tải xuống + 2 kiểm hồi quy lỗi bảng giả), browser-ai (phân nhánh
  bằng cảnh báo + độ phủ file), pwa-offline (chờ nút tải thay vì chữ xem trước).
  `tests/ui-normalize-test.py` đã XOÁ vì mục Word đã gỡ. Toàn bộ test đạt.

## Đợt 23/10/2026 (tiếp) — phần đầu nuốt thân văn bản khi ngày tháng dự thảo trống số ngày

Hải gửi ảnh chụp file Word lỗi: toàn bộ đầu văn bản (V/v, Kính gửi, các đoạn
Căn cứ...) bị VIẾT HOA và ép vào cột trái hẹp của bảng phần đầu. Nguyên nhân:
dòng ngày tháng của bản DỰ THẢO để trống số ngày ("ngày   tháng 10 năm 2026")
nên DATE_RE (đòi số ngày) không khớp → mốc kết thúc vùng header không thu hẹp
được → `extractAdminHeader` nuốt trọn 14 dòng đầu trang vào bảng đầu.
Đã sửa trong `extractAdminHeader` (pdfStructure.ts):
- Chốt chặn cứng khi thu thập dòng header: gặp dòng trích yếu "V/v...", dòng
  "Kính gửi" hoặc dòng thân văn bản dài chạm cả hai mép là DỪNG ngay; vùng
  header tối đa 9 dòng.
- Dòng ngày tháng ở cột phải vẫn được nhận làm mốc kết thúc kể cả khi trống
  số ngày (chỉ cần có đủ "ngày... tháng... năm").
Có test đơn vị mới mô phỏng đúng ca này (Sở GDĐT Tuyên Quang): đã kiểm chứng
test ĐỎ trên code cũ ("mất dòng Kính gửi ở thân văn bản") và XANH sau khi vá;
toàn bộ test còn lại (gồm CV 5095 + bảng tổng hợp trên trình duyệt thật) vẫn đạt.

Bổ sung cùng đợt (rà soát tổng hợp theo yêu cầu của Hải — "suy nghĩ trước
sau, tổng hợp các lỗi có thể phát sinh"):
- Dòng số công văn bị tách segment ("Số:" riêng, "/SGDĐT-GDPT" riêng — gặp
  khi số do dấu ký số đóng riêng): hậu tố nay được nối vào ô số ("Số:
  /SGDĐT-GDPT") thay vì rơi vào ô cơ quan. Test đơn vị đã cập nhật theo.
- Thêm dòng mở đầu trích yếu dạng "Về việc..." vào chốt chặn dừng header.
- Giới hạn đã xác minh trên đúng file của Hải (CV hướng dẫn kiểm tra, đánh
  giá — Sở GDĐT Tuyên Quang, 14 trang, ký số VNPT): số "3616" và giờ ký
  nằm trong LỚP VẼ của chữ ký số (signature appearance), KHÔNG nằm trong
  lớp chữ của PDF — pdf.js không trích được (poppler đọc được vì nó đọc cả
  lớp vẽ chú thích). App không bịa số: ô Số hiện "Số: /SGDĐT-GDPT", người
  dùng điền tay "3616". Quét toàn văn bản ra: 203 đoạn, 0 đoạn viết hoa
  oan/dính chữ/ký tự lạ.
- Đã rà soát thêm các ca biên của phần đầu: ngày tháng viết số ("ngày
  04/8/2026" — không neo được mốc nhưng còn chốt chặn V/v/Kính gửi/dòng
  dài), văn bản không có quốc hiệu lẫn cơ quan (không dựng bảng đầu, nội
  dung giữ nguyên dạng đoạn), trang bìa có tiêu đề lớn ở giữa trang (mốc
  tiêu đề căn giữa vẫn ăn), tài liệu bắt đầu bằng bảng (CV 5095 trang 2-3
  đạt), header lặp ở trang 2+ (chỉ trang 1 được dựng bảng đầu — quy tắc cũ).

## Đợt 23/10/2026 (chiều) — giao diện theo KẾ HOẠCH GỘP (Hải duyệt "Làm luôn")

Theo file Ke_Hoach_Gop_Nang_Cap_App_PDF.docx: GIỮ NGUYÊN pipeline xử lý và
cơ chế tự động; chỉ làm mới tầng hiển thị trong src/App.tsx:
- Header kính mờ cố định: logo + tiêu đề + dòng phụ ngắn; pill trạng thái
  AI (xanh "AI Gemini sẵn sàng" / vàng "Cài khóa AI") bấm mở hộp cài đặt.
- Hộp cài đặt khóa Gemini (modal): lưu/kiểm tra/xóa khóa; link lấy khóa
  bấm được trỏ aistudio.google.com/apikey (giữ quy tắc cũ).
- Ô thả file bo góc lớn, đổi màu + phóng nhẹ khi kéo file vào; thẻ file:
  icon PDF đỏ, tên rút gọn, dung lượng, SỐ TRANG (từ info.totalPages),
  nút Đổi file.
- Tiến trình theo giai đoạn kèm trang x/y như trước.
- Sau xử lý: 3 thẻ số liệu (Khối nội dung / Bảng biểu tìm thấy / Chuẩn
  NĐ 30); nút Word xanh dương; nút Excel ghi "Tải file Excel (N bảng)",
  0 bảng thì khóa hẳn với nhãn "Không có bảng để xuất Excel" (bỏ kiểu
  bấm vào mới báo như bản tối giản).
- Xem trước nội dung dạng xếp gọn (mặc định đóng), 2 thẻ Văn bản / Bảng
  biểu — dùng ĐÚNG html + rawText có sẵn trong PdfConvertResult và
  ExtractedTable {name, rows} của blocksToTables (bảng xem 5 dòng đầu);
  nút Sao chép văn bản dùng rawText. KHÔNG dùng bản App.tsx mẫu trong
  báo cáo của Hải (viết theo cấu trúc dữ liệu tưởng tượng type/text/cells).
- Cache PWA nâng xlvb-v5. ui_smoke.py đã cập nhật theo giao diện mới (thẻ
  file có số trang, thẻ số liệu, xem trước 2 thẻ, nút sao chép) và đạt;
  toàn bộ test còn lại (đơn vị 16/14/5/10/7, trình duyệt PDF thật, nhánh
  AI dự phòng, PWA offline) đều đạt. Kiểm đầu-cuối trên file thật của Hải
  (CV Tuyên Quang 14 trang): thẻ file hiện "14 trang", 197 khối nội dung,
  Excel khóa đúng vì văn bản không có bảng dữ liệu, xem trước hiện đúng
  phần đầu + V/v + Căn cứ.
