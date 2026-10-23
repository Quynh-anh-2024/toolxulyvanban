#!/usr/bin/env python3
"""Smoke test giao diện TỐI GIẢN (từ 23/10/2026: một màn hình duy nhất —
tải PDF lên, app tự xử lý, chỉ hiện 2 nút tải Word/Excel; không xem
trước, không panel số liệu, không mục khác) bằng Playwright:
mở app local -> upload PDF mẫu -> chờ nút tải hiện -> tải .docx + .xlsx
-> kiểm nội dung file tải về.
Chạy: python tests/ui_smoke.py [duong-dan-pdf]
"""
import sys, zipfile
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
    assert "Chuyển PDF sang Word" in title
    # Giao diện tối giản: đúng 1 ô tải file, không tab mục, không ô Word/Excel riêng.
    assert page.locator("input[type=file]").count() == 1
    assert page.get_by_role("tab").count() == 0

    page.set_input_files("#fileInput", PDF)
    btn_word = page.locator("button", has_text="Tải file Word")
    btn_word.wait_for(state="visible", timeout=120_000)
    btn_excel = page.locator("button", has_text="Tải file Excel")
    assert btn_excel.count() == 1
    print("Đã hiện 2 nút tải Word/Excel sau khi chuyển xong")

    with page.expect_download(timeout=60_000) as dl_info:
        btn_word.click()
    dl = dl_info.value
    path = "/tmp/downloaded-from-ui.docx"
    dl.save_as(path)
    print("Đã tải Word:", dl.suggested_filename, "->", path)
    assert dl.suggested_filename.endswith(".docx")
    z = zipfile.ZipFile(path)
    xml = z.read("word/document.xml").decode("utf-8")
    for probe in ["CỘNG HÒA XÃ HỘI", "KẾ HOẠCH", "Nội dung công việc", "HIỆU TRƯỞNG"]:
        assert probe in xml, f"docx tải về thiếu: {probe}"
    import re
    m = re.search(r'<w:pgMar[^/]*/>', xml)
    print("pgMar:", m.group(0) if m else None)

    with page.expect_download(timeout=60_000) as dl_info:
        btn_excel.click()
    dl = dl_info.value
    xpath = "/tmp/downloaded-from-ui.xlsx"
    dl.save_as(xpath)
    print("Đã tải Excel:", dl.suggested_filename, "->", xpath)
    assert dl.suggested_filename.endswith(".xlsx")
    import openpyxl
    wb = openpyxl.load_workbook(xpath, data_only=True)
    cells = " ".join(
        str(c.value) for ws in wb.worksheets for row in ws.iter_rows() for c in row if c.value is not None
    )
    assert "Nội dung công việc" in cells, "xlsx tải về thiếu bảng 'Nội dung công việc'"
    print("Excel: có bảng 'Nội dung công việc'")

    print("Lỗi JS trên trang:", errors if errors else "không có")
    assert not errors
    browser.close()
print("UI SMOKE TEST: ĐẠT")
