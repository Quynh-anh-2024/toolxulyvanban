<div align="center">
<img width="1200" height="475" alt="GHBanner" src="https://github.com/user-attachments/assets/0aa67016-6eaf-458a-adb2-6e31a0763ed6" />
</div>

# Chuẩn Hóa Văn Bản & Soát Lỗi AI

Công cụ tải file Word (.docx) **hoặc PDF**, chuẩn hoá theo thể thức
Nghị định 30/2020/NĐ-CP, soát lỗi bằng AI và xuất lại file Word.

## Tính năng chuyển PDF → Word (mới)

Tải lên một file `.pdf` bất kỳ, app tự xử lý theo từng trang:

| Loại PDF | Cách xử lý |
|---|---|
| Có lớp chữ (xuất từ Word/web) | Đọc trực tiếp chữ + tọa độ + cỡ chữ bằng pdf.js, dựng lại cấu trúc |
| Scan (ảnh quét/chụp, không có lớp chữ) | Render trang rồi OCR tiếng Việt (Tesseract `vie`) |
| Hỗn hợp | Trang nào không có lớp chữ thì OCR riêng trang đó |

Sau khi trích xuất, nội dung được phân tích thành cấu trúc văn bản hành chính:
phần đầu thể thức (cơ quan — quốc hiệu/tiêu ngữ — số ký hiệu — ngày tháng),
tên loại + trích yếu, các mục, bảng dữ liệu (nhận diện theo cột), khối nơi
nhận/chữ ký, ảnh nhúng. Nút xuất khi đó tạo **file .docx thật** chuẩn NĐ 30:

- Khổ A4; lề trên/dưới 2cm, trái 3cm, phải 2cm.
- Times New Roman; nội dung 13–14pt căn đều, giãn dòng/cách đoạn theo cấu hình,
  thụt đầu dòng 1–1,27cm.
- Quốc hiệu 13pt in hoa đậm; tiêu ngữ đậm, gạch chân đúng độ dài dòng chữ;
  cơ quan in hoa (cơ quan ban hành đậm); số ký hiệu thường; ngày tháng nghiêng.
- Bảng dữ liệu có viền, hàng đầu đậm căn giữa; khối chữ ký không viền.

**Giới hạn cần biết:** pdf.js không cho biết chữ đậm/nghiêng của font subset
nhúng trong PDF, nên đậm/nghiêng lẻ trong thân bài có thể bị mất — các thành
phần thể thức chính vẫn được áp đúng nhờ nhận diện cấu trúc. Chữ ký/con dấu
trên trang scan là hình ảnh nên OCR không chuyển thành chữ được.

## Tính năng Excel (file có bảng tính)

- Tải lên `.xlsx` / `.xls` / `.csv`: mỗi sheet thành một bảng trong tài liệu
  (có tiêu đề tên sheet) → xem trước và xuất Word .docx như với PDF.
- Sau khi nạp PDF hoặc Excel, nút **Tải Excel (.xlsx) các bảng đã trích**
  xuất mọi bảng dữ liệu ra một file .xlsx thật, mỗi bảng một sheet.

Kiểm thử tự động: `npm run test:pdf` (test cấu trúc + test đầu-cuối với PDF
mẫu trong `tests/fixtures/`, gồm cả OCR trang scan) và
`npx tsx tests/excel.test.mts` (test hai chiều Excel ↔ Word).

## Run and deploy your AI Studio app

This contains everything you need to run your app locally.

View your app in AI Studio: https://ai.studio/apps/ee85fd56-6461-4926-a428-3d7bfa9f31fa

## Run Locally

**Prerequisites:**  Node.js


1. Install dependencies:
   `npm install`
2. Set the `GEMINI_API_KEY` in [.env.local](.env.local) to your Gemini API key
3. Run the app:
   `npm run dev`
