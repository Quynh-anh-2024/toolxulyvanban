/**
 * workers/pdfWorker.ts — Điểm vào của worker pdf.js.
 *
 * THỨ TỰ IMPORT Ở ĐÂY LÀ BẢN CHẤT CỦA BẢN VÁ: ES module chạy các import
 * theo đúng thứ tự khai báo, trước thân module. Dòng import đầu tiên cài
 * các vá tương thích trình duyệt cũ (xem utils/browserPolyfills.ts) vào
 * phạm vi của worker TRƯỚC KHI mã pdf.js được nạp — vì pdf.js gọi .at()
 * ngay trong lúc khởi tạo worker, cài sau là không kịp. Trong worker cũng
 * có đường giải nén stream dùng for await...of trên ReadableStream và
 * tầng liên lạc dùng Promise.withResolvers; Safari cũ thiếu tất cả.
 */
// 1) Cài vá tương thích trước tiên (side-effect import, chạy ngay khi nạp)
// @ts-ignore — module tự cài đặt, không có giá trị xuất ra cần dùng
import "../utils/installPolyfills";
// 2) Worker gốc của pdf.js — BẢN LEGACY (có sẵn polyfill core-js cho các
// API đời mới khác), tự chứa hoàn toàn, không import thêm gì.
// @ts-ignore — file .mjs của thư viện, không có khai báo kiểu riêng
import "pdfjs-dist/legacy/build/pdf.worker.min.mjs";
