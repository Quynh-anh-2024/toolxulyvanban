#!/usr/bin/env python3
"""Tạo PDF mẫu để test: (1) PDF có lớp chữ kiểu công văn, (2) PDF scan (ảnh toàn trang)."""
import os
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

OUT = os.path.join(os.path.dirname(__file__), "fixtures")
os.makedirs(OUT, exist_ok=True)

SERIF = "/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf"
SERIF_B = "/usr/share/fonts/truetype/liberation/LiberationSerif-Bold.ttf"
SERIF_I = "/usr/share/fonts/truetype/liberation/LiberationSerif-Italic.ttf"
pdfmetrics.registerFont(TTFont("Serif", SERIF))
pdfmetrics.registerFont(TTFont("Serif-Bold", SERIF_B))
pdfmetrics.registerFont(TTFont("Serif-Italic", SERIF_I))

W, H = A4  # 595.27 x 841.89, gốc dưới-trái

def y(top):
    return H - top

# ---------------- (1) PDF có lớp chữ ----------------
c = canvas.Canvas(os.path.join(OUT, "cong-van-text.pdf"), pagesize=A4)
def line(text, x, top, font="Serif", size=13):
    c.setFont(font, size)
    c.drawString(x, y(top), text)

def centered(text, cx, top, font="Serif", size=13):
    c.setFont(font, size)
    c.drawCentredString(cx, y(top), text)

# Header: cột trái tâm ~190, cột phải tâm ~420
centered("PHÒNG GIÁO DỤC VÀ ĐÀO TẠO", 185, 52, "Serif", 11)
centered("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", 415, 52, "Serif-Bold", 12)
centered("TRƯỜNG PTDTBT TH MÈO VẠC", 185, 70, "Serif-Bold", 11)
centered("Độc lập - Tự do - Hạnh phúc", 415, 70, "Serif-Bold", 13)
centered("Số: 12/KH-THMV", 185, 94, "Serif", 12)
centered("Mèo Vạc, ngày 05 tháng 10 năm 2026", 415, 94, "Serif-Italic", 12)

centered("KẾ HOẠCH", W / 2, 128, "Serif-Bold", 14)
centered("Tổ chức dạy học năm học 2026 - 2027", W / 2, 150, "Serif-Bold", 14)

line("Căn cứ Thông tư số 32/2018/TT-BGDĐT ngày 26 tháng 12 năm 2018 của Bộ trưởng", 98, 182)
line("Bộ Giáo dục và Đào tạo ban hành Chương trình giáo dục phổ thông; Trường xây dựng", 85, 202)
line("kế hoạch tổ chức dạy học năm học 2026 - 2027 như sau:", 85, 222)
line("I. MỤC ĐÍCH, YÊU CẦU", 85, 254, "Serif-Bold")
line("Đảm bảo dạy học đúng chương trình, nâng cao chất lượng giáo dục toàn diện.", 98, 282)

# Bảng dữ liệu (chỉ chữ, không kẻ bảng — test nhận diện cột theo tọa độ)
line("STT", 90, 322, "Serif-Bold"); line("Nội dung công việc", 170, 322, "Serif-Bold"); line("Thời gian", 400, 322, "Serif-Bold")
line("1", 96, 344); line("Dạy học theo chương trình", 170, 344); line("Tuần 1 - 35", 400, 344)
line("2", 96, 366); line("Kiểm tra định kỳ giữa kỳ", 170, 366); line("Tuần 9", 400, 366)

# Khối nơi nhận / chữ ký
line("Nơi nhận:", 85, 420, "Serif-Italic")
centered("TM. BAN GIÁM HIỆU", 430, 420, "Serif-Bold")
line("- Ban Giám hiệu;", 85, 442)
centered("HIỆU TRƯỞNG", 430, 442, "Serif-Bold")
line("- Lưu: VT.", 85, 464)
centered("Trần Ngọc Hải", 430, 510, "Serif-Bold")
c.showPage()
c.save()

# ---------------- (2) PDF scan: ảnh toàn trang ----------------
from PIL import Image, ImageDraw
from PIL.ImageFont import truetype as load_font
img = Image.new("RGB", (1240, 1754), "white")
d = ImageDraw.Draw(img)
f_title = load_font(SERIF_B, 44)
f_body = load_font(SERIF, 36)
d.text((620, 120), "THÔNG BÁO", font=f_title, anchor="ma", fill="black")
d.text((90, 260), "Về việc nghỉ lễ Quốc khánh", font=f_body, fill="black")
d.text((90, 340), "Nhà trường thông báo: học sinh nghỉ lễ từ ngày", font=f_body, fill="black")
d.text((90, 400), "01 tháng 9 đến hết ngày 03 tháng 9 năm 2026.", font=f_body, fill="black")
d.text((90, 480), "Đề nghị giáo viên chủ nhiệm thông báo tới phụ huynh.", font=f_body, fill="black")
png_path = os.path.join(OUT, "scan-page.png")
img.save(png_path, dpi=(150, 150))

c2 = canvas.Canvas(os.path.join(OUT, "thong-bao-scan.pdf"), pagesize=A4)
c2.drawImage(png_path, 0, 0, width=W, height=H)
c2.showPage()
c2.save()
print("fixtures done:", os.listdir(OUT))
