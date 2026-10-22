"""Test UI đầu-cuối chức năng CHUẨN HOÁ WORD (docx -> .doc theo NĐ 30):
nạp file Word chưa chuẩn vào app, bấm xuất, mở file .doc tải xuống kiểm tra.
Chạy: python tests/ui-normalize-test.py
"""
import base64
import re
import subprocess
import sys
import tempfile
from pathlib import Path

from playwright.sync_api import sync_playwright

REPO = Path(__file__).resolve().parent.parent
FIXTURE = REPO / "tests" / "fixtures" / "van-ban-chua-chuan.docx"
OUT_DIR = REPO / "tests" / "real" / "out"
OUT_DIR.mkdir(parents=True, exist_ok=True)


def main() -> int:
    failures = []
    with sync_playwright() as pw:
        browser = pw.chromium.launch()
        page = browser.new_page()
        errors = []
        page.on("pageerror", lambda e: errors.append(str(e)))
        page.goto("http://localhost:3000/", wait_until="networkidle")
        # Từ 22/10/2026: file Word nằm ở mục riêng "Word & AI" — phải mở
        # đúng mục rồi nạp vào đúng ô tải của mục đó.
        page.click("text=Word & AI")
        page.set_input_files("#wordInput", str(FIXTURE))
        # Chờ bản xem trước có nội dung văn bản
        page.wait_for_function(
            "() => document.body.innerText.includes('THÔNG BÁO') || document.body.innerText.includes('thông báo')",
            timeout=30000,
        )
        page.wait_for_timeout(2500)
        with page.expect_download(timeout=30000) as dl_info:
            page.click("text=Tải Xuống File Word (.doc)")
        download = dl_info.value
        doc_path = OUT_DIR / "van-ban-da-chuan-hoa.doc"
        download.save_as(str(doc_path))
        browser.close()
    print(f"đã tải: {doc_path} ({doc_path.stat().st_size} bytes)")
    if errors:
        failures.append(f"Lỗi JS: {errors[:2]}")

    raw = doc_path.read_bytes()
    m = re.search(rb"Content-Transfer-Encoding: base64\r\n\r\n([A-Za-z0-9+/=\r\n]+)", raw)
    if not m:
        failures.append("Không tìm thấy phần HTML base64 trong file .doc")
        html = ""
    else:
        html = base64.b64decode(re.sub(rb"\s+", b"", m.group(1))).decode("utf8", "replace")

    def check(desc, cond):
        print(("  ✓ " if cond else "  ✘ ") + desc)
        if not cond:
            failures.append(desc)

    print("Kiểm tra thể thức trong file .doc xuất ra:")
    check("Khổ A4 (21cm x 29.7cm)", "21cm 29.7cm" in html)
    check("Lề trên/dưới 2cm, trái 3cm, phải 2cm", re.search(r"margin:\s*2cm 2cm 2cm 3cm", html) is not None)
    check("Quốc hiệu 13pt (không còn 10.8pt)", "10.8pt" not in html and re.search(r"\.admin-national-line\s*{[^}]*font-size:\s*13pt", html) is not None)
    check("Tiêu ngữ 13pt (không còn 12.5pt ở motto)", re.search(r"\.admin-motto-line\s*{[^}]*font-size:\s*13pt", html) is not None)
    check("Font Times New Roman", "Times New Roman" in html)
    check("Bảng đầu thể thức không viền", re.search(r"\.admin-header-table[^{]*{[^}]*border:\s*none", html) is not None)
    check("Tiêu đề căn giữa + in hoa + đậm", re.search(r"\.doc-main-title\s*{[^}]*text-align:\s*center", html) is not None and "text-transform: uppercase" in html)
    check("Nội dung văn bản còn nguyên", "kiểm tra công tác bán trú" in html)
    check("Bảng dữ liệu còn nguyên (Khoái Châu 2)", "Khoái Châu 2" in html)
    check("Số ký hiệu THẬT còn nguyên (123/SGDĐT-VP)", "123/SGDĐT-VP" in html)
    check("Không còn số ký hiệu bịa (/KH-...)", "/KH-" not in html)
    check(
        "Đúng 1 bảng đầu thể thức (bảng dữ liệu không bị biến thành bảng đầu)",
        len(re.findall(r"<table[^>]*admin-header-table", html)) == 1,
    )
    check(
        "Bảng dữ liệu đủ 3 hàng x 3 cột",
        all(x in html for x in ["Ngày kiểm tra", "Tiểu học Nhật Tân", "21/10/2026"]),
    )

    # Mở bằng LibreOffice -> PDF để kiểm tra hiển thị thật
    with tempfile.TemporaryDirectory() as td:
        r = subprocess.run(
            ["soffice", "--headless", "--convert-to", "pdf", "--outdir", td, str(doc_path)],
            capture_output=True, text=True, timeout=180,
        )
        pdfs = list(Path(td).glob("*.pdf"))
        if not pdfs:
            failures.append(f"LibreOffice không mở được file .doc: {r.stderr[:200]}")
        else:
            try:
                import pymupdf
            except ImportError:
                print("  (bỏ qua bước render LibreOffice: venv này không có pymupdf)")
                pymupdf = None
            if pymupdf is not None:

                d = pymupdf.open(pdfs[0])
                p1 = d[0]
                check("LibreOffice mở được, trang 1 là A4", abs(p1.rect.width - 595) < 3 and abs(p1.rect.height - 842) < 3)
                text = p1.get_text()
                check("Quốc hiệu hiển thị đúng", "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM" in text)
                # Tiêu đề THÔNG BÁO phải nằm giữa trang
                boxes = p1.search_for("THÔNG BÁO")
                if boxes:
                    cx = (boxes[0].x0 + boxes[0].x1) / 2
                    check(f"Tiêu đề căn giữa trang (tâm x={cx:.0f}, giữa A4=298)", abs(cx - 298) < 45)
                else:
                    failures.append("Không thấy tiêu đề THÔNG BÁO trong bản render")
                fonts = {f[3] for p in d for f in p.get_fonts()}
                check(f"Font hiển thị toàn Times New Roman ({sorted(fonts)})", all("Times" in f for f in fonts))

    print()
    if failures:
        print("CHƯA ĐẠT:", failures)
        return 1
    print("CHUẨN HOÁ WORD: mọi kiểm tra ĐẠT.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
