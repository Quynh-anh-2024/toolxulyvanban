/**
 * utils/installPolyfills.ts — Cài vá tương thích NGAY KHI MODULE ĐƯỢC NẠP.
 *
 * Tồn tại riêng vì thứ tự chạy của ES module: các import chạy TRƯỚC thân
 * của module nhập chúng. Nếu chỉ gọi installBrowserPolyfills() ở thân của
 * module chính thì pdf.js (được import cùng chỗ) đã chạy xong phần khởi
 * tạo — mà trong worker, pdf.js gọi .at() ngay trong lúc khởi tạo nên vẫn
 * chết. Module này phải được import ĐẦU TIÊN, trước mọi import pdf.js.
 */
import { installBrowserPolyfills } from "./browserPolyfills";

installBrowserPolyfills();
