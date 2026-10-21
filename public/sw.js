/**
 * Service Worker của app "Chuẩn Hóa Văn Bản & Soát Lỗi AI".
 *
 * Mục tiêu: cài app ra màn hình chính và dùng được KHI KHÔNG CÓ MẠNG —
 * chuyển PDF/Excel → Word và OCR bản scan đều chạy ngay trên máy (dữ liệu
 * OCR tiếng Việt đã đóng gói trong thư mục /tessdata và /tesseract-core),
 * nên chỉ cần mọi tài nguyên cùng nguồn được lưu đệm sau lần mở đầu tiên.
 * Các tính năng AI (soát lỗi, tóm tắt...) gọi API ngoài nên vẫn cần mạng —
 * service worker này KHÔNG đụng vào request khác nguồn.
 *
 * Chiến lược:
 * - Điều hướng (mở app): network-first, mất mạng thì trả bản đã lưu.
 * - Tài nguyên cùng nguồn (JS/CSS/ảnh/dữ liệu OCR...): stale-while-revalidate
 *   — trả ngay từ bộ nhớ đệm nếu có, đồng thời tải bản mới ở nền.
 * - Mỗi lần đổi CACHE_VERSION, bộ nhớ đệm cũ bị xoá sạch khi kích hoạt.
 */
const CACHE_VERSION = "xlvb-v3";
const CORE_ASSETS = [
  "/",
  "/index.html",
  "/manifest.webmanifest",
  "/icons/icon-192.png",
  "/icons/icon-512.png",
  "/icons/icon-maskable-512.png",
  "/icons/apple-touch-icon.png",
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches
      .open(CACHE_VERSION)
      .then((cache) => cache.addAll(CORE_ASSETS))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches
      .keys()
      .then((keys) =>
        Promise.all(keys.filter((k) => k !== CACHE_VERSION).map((k) => caches.delete(k)))
      )
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  const { request } = event;
  if (request.method !== "GET") return;
  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return; // API ngoài (AI) — không can thiệp
  if (request.headers.has("range")) return; // để trình duyệt tự xử lý

  if (request.mode === "navigate") {
    event.respondWith(
      fetch(request)
        .then((response) => {
          const copy = response.clone();
          caches.open(CACHE_VERSION).then((cache) => cache.put("/index.html", copy));
          return response;
        })
        .catch(() => caches.match("/index.html", { ignoreSearch: true }))
    );
    return;
  }

  event.respondWith(
    caches.open(CACHE_VERSION).then(async (cache) => {
      const cached = await cache.match(request, { ignoreSearch: true });
      const network = fetch(request)
        .then((response) => {
          if (response && response.status === 200) {
            cache.put(request, response.clone());
          }
          return response;
        })
        .catch(() => cached);
      return cached || network;
    })
  );
});
