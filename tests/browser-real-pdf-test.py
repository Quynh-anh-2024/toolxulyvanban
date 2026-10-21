"""Test TRÌNH DUYỆT THẬT với PDF thật từ hệ thống trường học (tests/real/):

Bài học 21/10/2026: pdf.js trên trình duyệt thiếu CMap/font chuẩn sẽ giải mã
sai font CID của PDF in từ cổng trường học — lớp chữ hoàn hảo bị tưởng là
"lỗi font" và app OCR oan, cho ra chữ sai dấu ("ý kiến" -> "ý kiện").
Test Node KHÔNG tái hiện được lỗi này (Node có sẵn CMap cạnh gói cài đặt),
nên mọi khẳng định về PDF hệ thống phải qua test trình duyệt này.

Chạy: python tests/browser-real-pdf-test.py   (cần server xem trước ở :4173)
"""
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright

REPO = Path(__file__).resolve().parent.parent
BASE = "http://localhost:4173"
BANG = REPO / "tests" / "real" / "bang-tong-hop.pdf"
CV = REPO / "tests" / "real" / "cv-5095-bgd.pdf"
OUT = REPO / "tests" / "real" / "out"
OUT.mkdir(parents=True, exist_ok=True)

failures = []


def check(desc, cond):
    print(("  ✓ " if cond else "  ✘ ") + desc)
    if not cond:
        failures.append(desc)


def body_text(page):
    return page.evaluate("() => document.body.innerText")


def main() -> int:
    with sync_playwright() as pw:
        browser = pw.chromium.launch()
        page = browser.new_page()
        errors = []
        page.on("pageerror", lambda e: errors.append(str(e)))
        page.on("console", lambda m: errors.append(m.text) if m.type == "error" else None)
        page.goto(BASE + "/", wait_until="networkidle")

        print("File 1 — bảng tổng hợp (PDF in từ hệ thống trường học, font CID)")
        page.set_input_files("input[type=file]", str(BANG))
        page.wait_for_function(
            "() => document.body.innerText.includes('Vừ Mí Bình')",
            timeout=180000,
        )
        text = body_text(page)
        check("Xem trước có tên học sinh đọc đúng từ lớp chữ", "Vừ Mí Bình" in text)
        check("Xem trước có 'Phẩm chất chủ yếu'", "Phẩm chất chủ yếu" in text)
        check("Không còn ký tự rác kiểu OCR ('s5igs', 'wu')", "s5igs" not in text and "wu" not in text)
        check(f"Dấu '|' lạ gần như không còn (đếm: {text.count('|')})", text.count("|") < 5)

        with page.expect_download(timeout=60000) as dl_info:
            page.click("text=Tải Excel (.xlsx)")
        dl = dl_info.value
        xlsx_path = OUT / "bang-tong-hop-app.xlsx"
        dl.save_as(str(xlsx_path))
        import openpyxl

        wb = openpyxl.load_workbook(xlsx_path, data_only=True)
        all_cells = [
            str(c.value)
            for ws in wb.worksheets
            for row in ws.iter_rows()
            for c in row
            if c.value is not None
        ]
        joined = " ".join(all_cells)
        check("Excel tải xuống chứa 'Vừ Mí Bình'", "Vừ Mí Bình" in joined)
        check("Excel tải xuống không chứa rác 's5igs'", "s5igs" not in joined)

        print("File 2 — CV 5095/BGDĐT (bản SCAN thật từ cổng ký số: đúng đường OCR)")
        page.set_input_files("input[type=file]", str(CV))
        page.wait_for_function(
            "() => document.body.innerText.includes('Trang OCR: 3')",
            timeout=300000,
        )
        # Khung thông tin trang hiện trước, nội dung xem trước render ngay
        # sau — chờ chữ OCR xuất hiện thật trong DOM rồi mới kiểm, tránh
        # đọc phải bản xem trước chưa kịp vẽ (đua thời gian render).
        try:
            page.wait_for_function(
                "() => document.body.innerText.includes('Nghị quyết')",
                timeout=60000,
            )
        except Exception:
            pass
        text = body_text(page)
        check("App nhận diện đúng là bản scan (3 trang OCR)", True)
        check("OCR đọc được 'BỘ GIÁO DỤC' ở phần đầu", "BỘ GIÁO DỤC" in text)
        check("OCR đọc được nhắc tới 'Nghị quyết'", "Nghị quyết" in text)

        with page.expect_download(timeout=60000) as dl_info:
            page.click("text=Tải File Word (.docx) Chuẩn NĐ 30")
        dl = dl_info.value
        docx_path = OUT / "cv-5095-app.docx"
        dl.save_as(str(docx_path))
        from docx import Document

        doc = Document(docx_path)
        full = "\n".join(p.text for p in doc.paragraphs)
        for tb in doc.tables:
            for r in tb.rows:
                for c in r.cells:
                    full += "\n" + c.text
        check("Word tải xuống từ bản scan có nội dung đọc được", "Nghị quyết" in full and len(full) > 800)

        if errors:
            print("  (lỗi console/trang:", errors[:3], ")")
        browser.close()

    print()
    if failures:
        print("CHƯA ĐẠT:", failures)
        return 1
    print("TRÌNH DUYỆT + PDF THẬT: mọi kiểm tra ĐẠT.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
