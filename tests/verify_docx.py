"""Kiểm chứng độc lập các .docx do app xuất ra từ PDF thật:
- Độ phủ từ vựng so với chữ gốc trong PDF (PyMuPDF đọc, chuẩn NFC).
- Ký tự lỗi: U+FFFD, vùng Private Use, ký tự điều khiển.
- Bảng: số bảng, số hàng/cột; đoạn: số đoạn.
- Thể thức: lề trang (cm), font chữ trong document.xml.
Chạy: python tests/verify_docx.py
"""
import re
import sys
import unicodedata
from collections import Counter
from pathlib import Path

import pymupdf
from docx import Document

REAL = Path(__file__).parent / "real"
OUT = REAL / "out"

TOKEN_RE = re.compile(r"[0-9A-Za-zÀ-ỹĐđ]+")
SUSPICIOUS_RE = re.compile(r"[\uFFFD\uE000-\uF8FF\u0000-\u0008\u000B-\u001F\u007F]")


def tokens(text: str) -> Counter:
    return Counter(TOKEN_RE.findall(unicodedata.normalize("NFC", text)))


def main() -> int:
    failures = 0
    for pdf_path in sorted(REAL.glob("*.pdf")):
        base = pdf_path.stem
        docx_path = OUT / f"{base}.docx"
        if not docx_path.exists():
            print(f"✘ {base}: THIẾU file .docx đầu ra")
            failures += 1
            continue
        src_doc = pymupdf.open(pdf_path)
        src_text = "\n".join(p.get_text() for p in src_doc)
        src_tokens = tokens(src_text)

        doc = Document(docx_path)
        body_texts = [p.text for p in doc.paragraphs]
        table_texts = []
        for t in doc.tables:
            for row in t.rows:
                for cell in row.cells:
                    table_texts.append(cell.text)
        out_text = "\n".join(body_texts + table_texts)
        out_tokens = tokens(out_text)

        covered = sum((src_tokens & out_tokens).values())
        total = sum(src_tokens.values())
        coverage = covered / total if total else 1.0
        suspicious = len(SUSPICIOUS_RE.findall(out_text))

        sec = doc.sections[0]
        margins = (
            round(sec.top_margin.cm, 1),
            round(sec.bottom_margin.cm, 1),
            round(sec.left_margin.cm, 1),
            round(sec.right_margin.cm, 1),
        )
        import zipfile

        xml = zipfile.ZipFile(docx_path).read("word/document.xml").decode("utf8")
        fonts = set(re.findall(r'w:ascii="([^"]+)"', xml))

        tables_info = [(len(t.rows), len(t.columns)) for t in doc.tables]
        print(f"■ {base} ({src_doc.page_count} trang)")
        print(f"  độ phủ chữ so với PDF gốc: {coverage * 100:.1f}%  | ký tự lỗi: {suspicious}")
        print(f"  đoạn: {len(doc.paragraphs)} | bảng: {tables_info[:8]}{'...' if len(tables_info) > 8 else ''}")
        print(f"  lề (trên/dưới/trái/phải cm): {margins} | font: {sorted(fonts)[:4]}")
        if suspicious > 0:
            failures += 1
            print("  ✘ CÓ KÝ TỰ LỖI")
        if src_doc.page_count > 0 and "scan" not in base and coverage < 0.97:
            failures += 1
            print(f"  ✘ ĐỘ PHỦ THẤP (<97%) — mất chữ")
        # scan: OCR so với chữ gốc của văn bản nguồn (đọc từ PDF gốc tương ứng)
        print()
    return failures


if __name__ == "__main__":
    sys.exit(1 if main() else 0)
