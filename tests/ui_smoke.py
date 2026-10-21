#!/usr/bin/env python3
"""Smoke test giao diện bằng Playwright (Chromium headless trong máy):
mở app local -> upload PDF mẫu -> chờ panel kết quả -> bấm tải .docx -> kiểm file tải.
Chạy: python tests/ui_smoke.py [duong-dan-pdf]
"""
import sys, zipfile, re
from playwright.sync_api import sync_playwright

PDF = sys.argv[1] if len(sys.argv) > 1 else "tests/fixtures/cong-van-text.pdf"
URL = "http://localhost:3000/"

with sync_playwright() as pw:
    browser = pw.chromium.launch()
    page = browser.new_page(accept_downloads=True)
    errors = []
    page.on("pageerror", lambda e: errors.append(str(e)))
    page.goto(URL, wait_until="networkidle")
    title = page.locator("h1").inner_text()
    print("H1:", title)
    assert "Chuẩn Hóa Văn Bản" in title

    page.set_input_files("#fileInput", PDF)
    # Panel kết quả chuyển PDF xuất hiện sau khi convert xong
    panel = page.locator("text=Kết quả chuyển PDF")
    panel.wait_for(state="visible", timeout=120_000)
    box = panel.locator("..").inner_text()
    print("PANEL:\n", box)
    assert "PDF có lớp chữ" in box
    assert "Số trang" in box

    preview = page.locator(".document-content").inner_text()
    assert "KẾ HOẠCH" in preview, "preview thiếu KẾ HOẠCH"
    assert "Nội dung công việc" in preview, "preview thiếu bảng"
    print("Preview: có KẾ HOẠCH + bảng 'Nội dung công việc'")

    btn = page.locator("button", has_text="Tải File Word (.docx) Chuẩn NĐ 30")
    assert btn.count() == 1, f"không thấy nút tải docx (count={btn.count()})"
    print("Nút tải:", btn.inner_text().strip())
    with page.expect_download(timeout=60_000) as dl_info:
        btn.click()
    dl = dl_info.value
    path = "/tmp/downloaded-from-ui.docx"
    dl.save_as(path)
    print("Đã tải:", dl.suggested_filename, "->", path)
    assert dl.suggested_filename.endswith(".docx")
    z = zipfile.ZipFile(path)
    xml = z.read("word/document.xml").decode("utf-8")
    for probe in ["CỘNG HÒA XÃ HỘI", "KẾ HOẠCH", "Nội dung công việc", "HIỆU TRƯỞNG"]:
        assert probe in xml, f"docx tải về thiếu: {probe}"
    m = re.search(r'<w:pgMar[^/]*/>', xml)
    print("pgMar:", m.group(0) if m else None)
    print("Lỗi JS trên trang:", errors if errors else "không có")
    browser.close()
print("UI SMOKE TEST: ĐẠT")
