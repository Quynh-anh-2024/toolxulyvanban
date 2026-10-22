// Test bộ đối chiếu AI ↔ OCR (utils/crossCheck.ts).
// Chạy: npx tsx tests/cross-check.test.mts
import { crossCheckPage, crossCheckWarning } from "../src/utils/crossCheck.ts";

let passed = 0;
let failed = 0;
function check(name: string, cond: boolean, extra = "") {
  if (cond) {
    passed++;
    console.log(`  ✓ ${name}`);
  } else {
    failed++;
    console.log(`  ✗ ${name}${extra ? " — " + extra : ""}`);
  }
}

const base =
  "Số: 5095/QĐ-BGDĐT Hà Nội, ngày 22 tháng 10 năm 2026. Kinh phí 150.000.000 đồng cho 632 học sinh.";

// 1. Hai bản giống hệt → khớp 100%, không gắn cờ.
{
  const r = crossCheckPage(base, base);
  check("giống hệt: agreement = 1", r.agreement === 1);
  check("giống hệt: không gắn cờ", !r.flagged && crossCheckWarning(1, r) === null);
}

// 2. OCR đọc sai số quyết định → gắn cờ, nêu đúng token khác.
{
  const ocr = base.replace("5095/QĐ-BGDĐT", "5096/QĐ-BGDĐT");
  const r = crossCheckPage(base, ocr);
  check("sai số QĐ: có gắn cờ", r.flagged);
  check(
    "sai số QĐ: phía AI nêu 5095, phía OCR nêu 5096",
    r.aiDigitDiffs.some((t) => t.includes("5095")) &&
      r.ocrDigitDiffs.some((t) => t.includes("5096")),
    JSON.stringify(r)
  );
  const w = crossCheckWarning(3, r);
  check("sai số QĐ: cảnh báo nhắc kiểm tra bản gốc", !!w && w.includes("Trang 3") && w.includes("kiểm tra"));
}

// 3. Chỉ khác dấu tiếng Việt (điểm yếu quen thuộc của OCR) → KHÔNG gắn cờ.
{
  const ocr = "Số: 5095/QĐ-BGDĐT Hà Nội, ngày 22 tháng 10 năm 2026. Kinh phí 150.000.000 đồng cho 632 hoc sinh.";
  const r = crossCheckPage(base, ocr);
  check("khác dấu chữ thường: không gắn cờ", !r.flagged, JSON.stringify(r.aiDigitDiffs));
  check("khác dấu chữ thường: agreement vẫn cao", r.agreement > 0.9);
}

// 4. OCR đọc sai số tiền → gắn cờ.
{
  const ocr = base.replace("150.000.000", "150.000.00");
  const r = crossCheckPage(base, ocr);
  check("sai số tiền: có gắn cờ", r.flagged);
}

// 5. AI sót một con số mà OCR có → phía OCR bị nêu.
{
  const ai = base.replace(" cho 632 học sinh", " cho học sinh");
  const r = crossCheckPage(ai, base);
  check("AI sót số: gắn cờ phía OCR", r.flagged && r.ocrDigitDiffs.some((t) => t.includes("632")));
}

// 6. Văn bản rỗng → không gắn cờ, không lỗi.
{
  const r = crossCheckPage("", "");
  check("rỗng: agreement = 1, không cờ", r.agreement === 1 && !r.flagged);
}

console.log(`\nCROSS-CHECK TEST: ${passed} passed, ${failed} failed`);
if (failed > 0) process.exit(1);
