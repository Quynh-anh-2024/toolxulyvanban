"""Tái hiện lỗi Safari/iOS cũ: "undefined is not a function (near '...i of t...')"
khi chuyển đổi PDF (Hải báo 21/10/2026 16:45 trên iPhone).

Cách làm: chạy app bằng WebKit (engine JavaScriptCore giống Safari) nhưng
GỠ BỎ các API JavaScript đời mới mà Safari trên iOS < 18.2 / < 26 chưa có:
- Iterator helpers: iterator.map/filter/toArray/... (Safari 26 mới có)
- Promise.try (Safari 18.2)
- Uint8Array.prototype.toBase64/toHex + Uint8Array.fromBase64/fromHex (Safari 18.2)
- Object.groupBy / Map.groupBy (Safari 18)
Việc gỡ áp dụng cả ở main thread (add_init_script) lẫn trong worker của
pdf.js/Tesseract (chặn request worker, chèn mã gỡ vào đầu file).

Chạy: ~/workspace/.venv-pw/bin/python tests/repro-old-safari.py [chromium|webkit]
(cần server xem trước ở :4173 — vite preview)
"""
import os
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright

REPO = Path(__file__).resolve().parent.parent
BASE = "http://localhost:4173"
_PDFS = {
    "bang": REPO / "tests" / "real" / "bang-tong-hop.pdf",
    "cv": REPO / "tests" / "real" / "cv-5095-bgd.pdf",
}
BANG = _PDFS[os.environ.get("REPRO_PDF", "bang")]
MARKER = os.environ.get("REPRO_MARKER", "Vừ Mí Bình")
ENGINE = sys.argv[1] if len(sys.argv) > 1 else "webkit"

# REPRO_KEEP: danh sách nhóm API KHÔNG gỡ (để phân tầng mức Safari cần giả lập).
# Mặc định gỡ tất cả (giả lập Safari rất cũ).
KEEP = set(filter(None, os.environ.get("REPRO_KEEP", "").split(",")))

_GROUPS = {
    "iter": r"""
    var iterProto = Object.getPrototypeOf(Object.getPrototypeOf(new Map().keys()));
    var names = ["map","filter","flatMap","take","drop","toArray","reduce","some","every","find","forEach","count"];
    for (var i = 0; i < names.length; i++) {
      try { delete iterProto[names[i]]; } catch (e) {}
    }""",
    "promiseTry": "try { delete Promise.try; } catch (e) {}",
    "withResolvers": "try { delete Promise.withResolvers; } catch (e) {}",
    "bytes": """
  try { delete Uint8Array.prototype.toBase64; } catch (e) {}
  try { delete Uint8Array.prototype.toHex; } catch (e) {}
  try { delete Uint8Array.fromBase64; } catch (e) {}
  try { delete Uint8Array.fromHex; } catch (e) {}""",
    "groupBy": """
  try { delete Object.groupBy; } catch (e) {}
  try { delete Map.groupBy; } catch (e) {}
  try { delete Map.prototype.getOrInsert; } catch (e) {}
  try { delete Map.prototype.getOrInsertComputed; } catch (e) {}""",
    "at": "try { delete Array.prototype.at; } catch (e) {}",
    "replaceAll": "try { delete String.prototype.replaceAll; } catch (e) {}",
    # ĐÂY là điều kiện đúng của Hải: Safari < 18.4 không có duyệt
    # ReadableStream bằng for await...of -> pdf.js getTextContent văng
    # "undefined is not a function (near '...i of t...')" trên JavaScriptCore.
    "streamIter": "try { delete ReadableStream.prototype[Symbol.asyncIterator]; } catch (e) {}",
}

STRIP_JS = "(() => {\n" + "\n".join(
    code for name, code in _GROUPS.items() if name not in KEEP
) + "\n})();"


def main() -> int:
    dialogs = []
    page_errors = []
    with sync_playwright() as pw:
        browser = getattr(pw, ENGINE).launch()
        ctx = browser.new_context()
        page = ctx.new_page()
        page.on("dialog", lambda d: (dialogs.append(d.message), d.accept()))
        page.on(
            "pageerror",
            lambda e: page_errors.append(
                str(e) + "\nSTACK: " + (getattr(e, "stack", None) or "(no stack)")
            ),
        )
        page.on(
            "console",
            lambda m: page_errors.append("CONSOLE-" + m.type + ": " + m.text)
            if m.type in ("error", "warning")
            else None,
        )

        def strip_worker(route):
            resp = route.fetch()
            body = resp.text()
            route.fulfill(response=resp, body=STRIP_JS + "\n" + body)

        # Chèn mã gỡ API vào đầu mọi worker (pdf.js + tesseract)
        import re

        page.route(re.compile(r"worker", re.I), strip_worker)
        # Không cho service worker can thiệp vào test
        page.route("**/sw.js", lambda r: r.abort())
        page.add_init_script(STRIP_JS)

        page.goto(BASE + "/", wait_until="networkidle")
        print(f"[{ENGINE}] Đã nạp app (API mới đã bị gỡ để giả lập Safari cũ)")
        page.set_input_files("input[type=file]", str(BANG))
        try:
            page.wait_for_function(
                f"() => document.body.innerText.includes({MARKER!r})",
                timeout=int(os.environ.get("REPRO_TIMEOUT_MS", "300000")),
            )
            print(f"[{ENGINE}] KẾT QUẢ: chuyển đổi THÀNH CÔNG — xem trước có '{MARKER}'")
            outcome = 0
        except Exception:
            print(f"[{ENGINE}] KẾT QUẢ: chuyển đổi THẤT BẠI (không thấy nội dung xem trước)")
            outcome = 1
        for d in dialogs:
            print(f"[{ENGINE}] ALERT:", d)
        for e in page_errors[:8]:
            print(f"[{ENGINE}] PAGEERROR:", e)
        browser.close()
    return outcome


if __name__ == "__main__":
    sys.exit(main())
