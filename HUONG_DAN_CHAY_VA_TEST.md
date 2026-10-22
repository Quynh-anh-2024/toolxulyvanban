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
