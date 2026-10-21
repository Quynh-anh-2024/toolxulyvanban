/**
 * utils/browserPolyfills.ts — Vá tương thích cho trình duyệt cũ.
 *
 * Bối cảnh (21/10/2026): Hải mở app trên iPhone (Safari đời cũ) thì MỌI file
 * PDF đều văng lỗi "undefined is not a function (near '...i of t...')" ngay
 * khi chuyển đổi, trong khi Chrome/Android chạy bình thường. Nguyên nhân:
 * pdf.js v6 (kể cả bản legacy) dùng hàng loạt API JavaScript rất mới:
 *  1. `for await...of` trên ReadableStream — cần
 *     ReadableStream.prototype[Symbol.asyncIterator], Safari mới có từ
 *     18.4 (3/2025). `page.getTextContent()` của pdf.js lặp stream đúng
 *     kiểu này nên PDF nào cũng chết trên Safari cũ — đây là lỗi Hải gặp.
 *  2. `Promise.withResolvers()` — Safari 17.4+; tầng liên lạc với worker
 *     của pdf.js dùng khắp nơi, thiếu là chết ngay khi mở file.
 *  3. `Promise.try()` — Safari 18.2+; tầng stream của pdf.js có dùng.
 *  4. `.at()` trên Array/String/TypedArray — Safari 15.4+; pdf.js dùng
 *     .at(-1) ngay trong lúc khởi tạo worker và trong đường render trang.
 *  5. `String.prototype.replaceAll` — Safari 13.1+; pdf.js dùng khi dịch
 *     font, thiếu là font tải hỏng và trích chữ ra trống.
 * Hàm này chỉ bổ sung đúng phần còn thiếu (trình duyệt nào đã có sẵn thì
 * không đụng vào), dùng được cả ở main thread lẫn trong worker, và PHẢI
 * chạy trước khi mã pdf.js được nạp (xem utils/installPolyfills.ts).
 */

/** Gắn API theo đúng chuẩn bản địa: không liệt kê được (non-enumerable),
 * ghi được, xoá được. Bắt buộc phải vậy: pdf.js có kiểm tra "Array.prototype
 * chứa thuộc tính lạ liệt kê được" và từ chối chạy nếu gắn kiểu gán thường. */
function defineOn(proto: any, key: string | symbol, value: any): void {
  Object.defineProperty(proto, key, {
    value,
    writable: true,
    configurable: true,
    enumerable: false,
  });
}

export function installBrowserPolyfills(): void {
  const g: any = globalThis as any;

  // 1) ReadableStream: duyệt bằng for await...of qua reader.
  try {
    const RS = g.ReadableStream;
    if (typeof RS === "function" && RS.prototype && !RS.prototype[Symbol.asyncIterator]) {
      defineOn(RS.prototype, Symbol.asyncIterator, function (this: any) {
        const reader = this.getReader();
        let finished = false;
        const release = () => {
          try {
            reader.releaseLock();
          } catch {
            /* đã nhả khoá rồi */
          }
        };
        return {
          async next() {
            if (finished) return { done: true, value: undefined };
            try {
              const result = await reader.read();
              if (result.done) {
                finished = true;
                release();
                return { done: true, value: undefined };
              }
              return { done: false, value: result.value };
            } catch (err) {
              finished = true;
              release();
              throw err;
            }
          },
          async return() {
            if (!finished) {
              finished = true;
              try {
                await reader.cancel();
              } catch {
                /* stream đã đóng */
              }
              release();
            }
            return { done: true, value: undefined };
          },
          [Symbol.asyncIterator]() {
            return this;
          },
        };
      });
    }
  } catch {
    /* môi trường không có ReadableStream — pdf.js tự báo lỗi riêng */
  }

  // 2) + 3) Promise.withResolvers / Promise.try
  try {
    const P = g.Promise;
    if (typeof P === "function") {
      if (typeof P.withResolvers !== "function") {
        defineOn(P, "withResolvers", function () {
          let resolve: any;
          let reject: any;
          const promise = new P((res: any, rej: any) => {
            resolve = res;
            reject = rej;
          });
          return { promise, resolve, reject };
        });
      }
      if (typeof P.try !== "function") {
        defineOn(P, "try", function (fn: any, ...args: any[]) {
          return new P((resolve: any) => resolve(fn(...args)));
        });
      }
    }
  } catch {
    /* bỏ qua — môi trường lạ */
  }

  // 4) .at() cho Array / String / TypedArray — Safari 15.4+.
  try {
    const atImpl = function (this: any, index: any) {
      const len = this.length >>> 0;
      let i = Math.trunc(Number(index)) || 0;
      if (i < 0) i += len;
      if (i < 0 || i >= len) return undefined;
      return this[i];
    };
    const installAt = (proto: any) => {
      if (proto && typeof proto.at !== "function") defineOn(proto, "at", atImpl);
    };
    installAt(Array.prototype);
    installAt(String.prototype);
    // %TypedArray%.prototype — nơi .at của mọi TypedArray cư trú
    installAt(Object.getPrototypeOf(Uint8Array.prototype));
  } catch {
    /* bỏ qua — môi trường lạ */
  }

  // 5) String.prototype.replaceAll — Safari 13.1+.
  // Vá dự phòng này lo đúng các dạng pdf.js gọi: chuỗi con thường và
  // RegExp có cờ /g (uỷ thác cho replace). Chuỗi thay thế dạng chuỗi hỗ
  // trợ $$ và $& như bản địa.
  try {
    if (typeof String.prototype.replaceAll !== "function") {
      defineOn(String.prototype, "replaceAll", function (this: any, search: any, replacement: any) {
        const str = String(this);
        if (search instanceof RegExp) {
          if (!search.global) {
            throw new TypeError("String.prototype.replaceAll called with a non-global RegExp argument");
          }
          return str.replace(search, replacement);
        }
        const needle = String(search);
        const isFn = typeof replacement === "function";
        const SENTINEL = String.fromCharCode(1);
        const repStr = isFn
          ? ""
          : String(replacement)
              .split("$$")
              .join(SENTINEL)
              .split("$&")
              .join(needle)
              .split(SENTINEL)
              .join("$");
        if (needle === "") {
          // Chèn vào mọi khe: đầu chuỗi, giữa các ký tự, cuối chuỗi
          let out = "";
          for (let i = 0; i <= str.length; i++) {
            out += isFn ? replacement("", i, str) : repStr;
            if (i < str.length) out += str[i];
          }
          return out;
        }
        let out = "";
        let pos = 0;
        for (;;) {
          const idx = str.indexOf(needle, pos);
          if (idx < 0) break;
          out += str.slice(pos, idx) + (isFn ? replacement(needle, idx, str) : repStr);
          pos = idx + needle.length;
        }
        return out + str.slice(pos);
      });
    }
  } catch {
    /* bỏ qua — môi trường lạ */
  }
}
