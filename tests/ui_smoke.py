#!/usr/bin/env python3
"""Smoke test giao diện theo KẾ HOẠCH GỘP (Hải duyệt 23/10/2026): cơ chế
vẫn tự động (thả PDF vào là xử lý), tầng hiển thị làm mới — header kính
mờ + pill trạng thái AI mở hộp cài đặt khóa, thẻ file có số trang, 3 thẻ
số liệu, nút Excel ghi số bảng, xem trước xếp gọn 2 thẻ, nút sao chép.
Kiểm bằng Playwright: mở app local -> upload PDF mẫu -> chờ nút tải
hiện -> tải .docx + .xlsx -> kiểm nội dung file tải về -> mở xem trước.
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
    # Đúng 1 ô tải file; thẻ xem trước dùng nút thường, không phải role=tab.
    assert page.locator("input[type=file]").count() == 1
    assert page.get_by_role("tab").count() == 0
    # Pill trạng thái AI trên header (chưa lưu khóa -> "Cài khóa AI").
    assert page.locator("button", has_text="Cài khóa AI").count() == 1

    page.set_input_files("#fileInput", PDF)
    btn_word = page.locator("button", has_text="Tải file Word")
    btn_word.wait_for(state="visible", timeout=120_000)
    btn_excel = page.locator("button", has_text="Tải file Excel")
    assert btn_excel.count() == 1
    print("Đã hiện 2 nút tải Word/Excel sau khi chuyển xong")
    # Thẻ file có số trang + 3 thẻ số liệu + nút Excel ghi số bảng.
    section_text = page.locator("section").first.inner_text()
    assert "trang" in section_text, "thẻ file thiếu số trang"
    for probe in ["Khối nội dung", "Bảng biểu tìm thấy", "Chuẩn NĐ 30", "bảng)"]:
        assert probe in section_text, f"thiếu thẻ số liệu/nút: {probe}"
    print("Thẻ file + thẻ số liệu + nút Excel có số bảng: đủ")

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

    # Xem trước xếp gọn: mở ra thấy nội dung văn bản; thẻ Bảng biểu có bảng thật.
    page.click("button:has-text('Xem trước nội dung')")
    page.wait_for_timeout(600)
    preview_text = page.locator("section").first.inner_text()
    assert "KẾ HOẠCH" in preview_text, "xem trước văn bản không có nội dung"
    page.click("button:has-text('Bảng biểu (')")
    page.wait_for_timeout(400)
    assert page.locator("section table").count() >= 1, "thẻ Bảng biểu không dựng được bảng"
    assert page.locator("button", has_text="Sao chép văn bản").count() == 1
    print("Xem trước 2 thẻ + nút sao chép: đạt")

    print("Lỗi JS trên trang:", errors if errors else "không có")
    assert not errors
    browser.close()
print("UI SMOKE TEST: ĐẠT")
