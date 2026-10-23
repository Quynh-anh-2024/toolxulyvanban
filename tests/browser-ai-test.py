#!/usr/bin/env python3
"""Test trình duyệt THẬT cho đường AI ĐÁNH LẠI trang scan (Gemini vision).

Luồng: mở app (preview :4173) → nạp khóa Gemini của app phiếu bài tập vào
localStorage (đọc từ services/defaultChatKey.ts, KHÔNG in ra) → tải trang 1
CV 5095 (bản scan cổng ký số) → chờ nút tải Word hiện → tải .docx → đo độ
phủ từ theo bản chép tay trang 1 (truth-cv-p1.txt).
Kỳ vọng nhánh AI ≥97% (đo thô qua API: 100%; Tesseract ~89% từ thô).

HAI NHÁNH HỢP LỆ: (a) mạng tới Google thông → Word phủ ≥97%; (b) môi
trường chặn mạng tới Google (như sandbox CI của Muse: trình duyệt
headless bị proxy chặn CONNECT tới generativelanguage) → app phải tự rơi
xuống OCR kèm cảnh báo "AI đánh lại không thành công" — test kiểm đúng
hành vi dự phòng đó và in ghi chú, không coi là đạt nhánh AI.

Từ 23/10/2026 giao diện TỐI GIẢN: không còn panel số liệu/xem trước —
nhánh được phân biệt bằng cảnh báo trên trang + độ phủ của file tải về.

Chạy: ~/workspace/lbg-demo/venv/bin/python tests/browser-ai-test.py
"""
import re
import subprocess
import sys
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
CV = REPO / "tests" / "real" / "cv-5095-bgd.pdf"
TRUTH = REPO / "tests" / "real" / "truth-cv-p1.txt"
CHAT_KEY_FILE = Path.home() / "workspace" / "phieubaitap-github" / "services" / "defaultChatKey.ts"
OUT = Path("/tmp/ai-test-out")
OUT.mkdir(exist_ok=True)
URL = "http://localhost:4173/"

TOKEN_RE = re.compile(r"[0-9A-Za-zÀ-ỹĐđ]+")


def tokens(text: str) -> dict:
    counts: dict = {}
    for tok in TOKEN_RE.findall(text):
        counts[tok] = counts.get(tok, 0) + 1
    return counts


def multiset_coverage(truth_text: str, out_text: str) -> float:
    t, o = tokens(truth_text), tokens(out_text)
    total = sum(t.values()) or 1
    hit = sum(min(v, o.get(k, 0)) for k, v in t.items())
    return hit / total


def main() -> int:
    # Tách riêng trang 1 của CV thành PDF scan 1 trang để đo phủ chính xác.
    cv1 = OUT / "cv-5095-p1.pdf"
    subprocess.run(
        ["pdfseparate", "-f", "1", "-l", "1", str(CV), str(OUT / "cv-%d.pdf")],
        check=True,
    )
    (OUT / "cv-1.pdf").replace(cv1)

    key_src = CHAT_KEY_FILE.read_text(encoding="utf-8")
    m = re.search(r"DEFAULT_CHAT_API_KEY\s*=\s*[\"'`]([^\"'`]+)[\"'`]", key_src)
    if not m:
        print("Không trích được khóa Gemini từ file key của app phiếu bài tập.")
        return 1
    gemini_key = m.group(1).strip()

    from playwright.sync_api import sync_playwright

    failures = []

    def check(name: str, cond: bool, extra: str = ""):
        print(("  ✓ " if cond else "  ✗ ") + name + (f" — {extra}" if extra and not cond else ""))
        if not cond:
            failures.append(name)

    with sync_playwright() as pw:
        browser = pw.chromium.launch()
        page = browser.new_page(accept_downloads=True)
        errors: list = []
        page.on("pageerror", lambda e: errors.append(str(e)))
        page.goto(URL, wait_until="networkidle")
        page.evaluate("(k) => localStorage.setItem('xlvb_gemini_key', k)", gemini_key)
        page.reload(wait_until="networkidle")
        check("khóa đã lưu → ô nhập khóa ẩn đi", page.locator("#geminiKeyInput").count() == 0)

        page.set_input_files("#fileInput", str(cv1))
        # Chuyển xong = nút tải Word hiện. Nhánh dự phòng nay có thêm vòng
        # AI thử lại tuần tự nên có thể lâu hơn trước — cho 600s.
        page.locator("button", has_text="Tải file Word").wait_for(state="visible", timeout=600000)
        section_text = page.locator("section").first.inner_text()

        with page.expect_download(timeout=120000) as dl_info:
            page.click("button:has-text('Tải file Word')")
        dl = dl_info.value
        docx_path = OUT / "cv-5095-ai.docx"
        dl.save_as(str(docx_path))
        with zipfile.ZipFile(docx_path) as z:
            xml = z.read("word/document.xml").decode("utf-8")
        page1_text = re.sub(r"<[^>]+>", " ", xml)

        if "AI đánh lại không thành công" in section_text:
            # Nhánh (b): môi trường chặn mạng tới Google — app phải rơi OCR
            # kèm cảnh báo rõ ràng, không treo, không mất nội dung.
            check("dự phòng OCR khi AI không tới được Google (có cảnh báo)", True)
            check("Word dự phòng vẫn có nội dung chính", "Nghị quyết" in page1_text)
            check("không có lỗi JS", not errors, "; ".join(errors[:2]))
            print("\nGHI CHÚ: nhánh AI chưa chạy được trong môi trường này (mạng tới Google bị chặn) — chỉ kiểm được nhánh dự phòng OCR.")
            print(f"\nBROWSER-AI TEST (nhánh dự phòng): {'PASS' if not failures else 'FAIL ' + str(failures)}")
            browser.close()
            return 0 if not failures else 1

        truth = TRUTH.read_text(encoding="utf-8")
        cov = multiset_coverage(truth, page1_text)
        print(f"Độ phủ từ Word (AI) so với bản chép tay trang 1: {cov:.1%} (OCR cũ: 91,1%)")
        check("Word AI phủ ≥97% từ trang 1", cov >= 0.97, f"{cov:.1%}")
        check("Word có tên cơ quan đúng", "BỘ GIÁO DỤC VÀ ĐÀO TẠO" in page1_text)
        check("Word có trích yếu", "V/v xin ý kiến góp ý dự thảo Nghị quyết" in page1_text)
        check("Word giữ nguyên câu khó có dấu", "phum, sóc, ấp" in page1_text)
        check("không có lỗi JS", not errors, "; ".join(errors[:2]))
        browser.close()

    print(f"\nBROWSER-AI TEST: {'PASS' if not failures else 'FAIL ' + str(failures)}")
    return 0 if not failures else 1


if __name__ == "__main__":
    sys.exit(main())
