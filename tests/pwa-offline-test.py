"""Test PWA + OFFLINE thật cho app xử lý văn bản (bản production):
1) Mở app khi có mạng, kiểm manifest + service worker kích hoạt.
2) Nạp 1 PDF scan để app tải đủ tài nguyên OCR vào bộ nhớ đệm.
3) NGẮT MẠNG hoàn toàn -> tải lại trang: app phải mở được từ bộ nhớ đệm.
4) Khi đang offline: chuyển PDF chữ -> Word và OCR PDF scan phải chạy được.
Chạy: python tests/pwa-offline-test.py   (cần server xem trước ở :4173)
"""
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright

REPO = Path(__file__).resolve().parent.parent
BASE = "http://localhost:4173"
TEXT_PDF = REPO / "tests" / "fixtures" / "cong-van-text.pdf"
SCAN_PDF = REPO / "tests" / "fixtures" / "thong-bao-scan.pdf"
OUT = REPO / "tests" / "real" / "out"
OUT.mkdir(parents=True, exist_ok=True)

failures = []


def check(desc, cond):
    print(("  ✓ " if cond else "  ✘ ") + desc)
    if not cond:
        failures.append(desc)


def main() -> int:
    with sync_playwright() as pw:
        browser = pw.chromium.launch()
        ctx = browser.new_context()
        page = ctx.new_page()
        errors = []
        page.on("pageerror", lambda e: errors.append(str(e)))

        print("Giai đoạn 1 — có mạng: kích hoạt PWA")
        page.goto(BASE + "/", wait_until="networkidle")
        manifest = page.evaluate(
            """async () => {
              const link = document.querySelector('link[rel="manifest"]');
              if (!link) return null;
              const r = await fetch(link.href);
              return { status: r.status, json: await r.json() };
            }"""
        )
        check("Có manifest và tải được", bool(manifest) and manifest["status"] == 200)
        if manifest:
            j = manifest["json"]
            check(
                f"Manifest đúng tên app + có icon 512 ({j.get('short_name')})",
                j.get("name", "").startswith("Chuẩn Hóa") and any("512" in (i.get("sizes") or "") for i in j.get("icons", [])),
            )
        sw_state = page.evaluate(
            """async () => {
              const reg = await navigator.serviceWorker.ready;
              return { active: !!reg.active, scope: reg.scope };
            }"""
        )
        check(f"Service worker đã kích hoạt (scope {sw_state['scope']})", sw_state["active"])
        page.reload(wait_until="networkidle")  # để SW quản lý trang
        controlled = page.evaluate("() => !!navigator.serviceWorker.controller")
        check("Trang đã được service worker phục vụ", controlled)

        print("Giai đoạn 2 — nạp tài nguyên OCR vào bộ nhớ đệm (vẫn có mạng)")
        page.set_input_files("input[type=file]", str(SCAN_PDF))
        page.wait_for_function(
            "() => document.body.innerText.includes('THÔNG BÁO')",
            timeout=240000,
        )
        check("OCR online chạy xong (làm ấm bộ nhớ đệm)", True)
        cache_info = page.evaluate(
            """async () => {
              const names = await caches.keys();
              let entries = 0, tess = 0;
              for (const n of names) {
                const c = await caches.open(n);
                const keys = await c.keys();
                entries += keys.length;
                tess += keys.filter(k => k.url.includes('tessdata') || k.url.includes('tesseract')).length;
              }
              const est = await navigator.storage.estimate();
              return { names, entries, tess, mb: Math.round((est.usage || 0) / 1048576) };
            }"""
        )
        print(f"    bộ nhớ đệm: {cache_info}")
        check("Đã lưu đệm tài nguyên OCR tiếng Việt", cache_info["tess"] >= 3)

        print("Giai đoạn 3 — NGẮT MẠNG: mở lại app từ bộ nhớ đệm")
        ctx.set_offline(True)
        page.reload(wait_until="load")
        page.wait_for_timeout(2500)
        has_ui = page.evaluate(
            "() => !!document.querySelector('input[type=file]') && document.body.innerText.length > 200"
        )
        check("App mở được khi không có mạng", has_ui)

        print("Giai đoạn 4 — OFFLINE: chuyển PDF chữ -> Word")
        page.set_input_files("input[type=file]", str(TEXT_PDF))
        page.wait_for_function(
            "() => document.body.innerText.includes('KẾ HOẠCH')",
            timeout=120000,
        )
        check("Đọc PDF chữ offline: thấy nội dung KẾ HOẠCH", True)
        with page.expect_download(timeout=60000) as dl_info:
            page.click("text=Tải File Word (.docx) Chuẩn NĐ 30")
        dl = dl_info.value
        target = OUT / "offline-pdf-to-word.docx"
        dl.save_as(str(target))
        check(f"Xuất .docx offline thành công ({target.stat().st_size} bytes)", target.stat().st_size > 5000)

        print("Giai đoạn 5 — OFFLINE: OCR bản scan")
        page.set_input_files("input[type=file]", str(SCAN_PDF))
        page.wait_for_function(
            "() => document.body.innerText.includes('nghỉ lễ') || document.body.innerText.includes('Quốc khánh')",
            timeout=240000,
        )
        check("OCR offline đọc được nội dung bản scan", True)

        if errors:
            failures.append(f"Lỗi JS: {errors[:2]}")
        ctx.set_offline(False)
        browser.close()

    print()
    if failures:
        print("CHƯA ĐẠT:", failures)
        return 1
    print("PWA OFFLINE: mọi kiểm tra ĐẠT.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
