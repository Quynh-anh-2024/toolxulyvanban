/* ------------------------------------------------------------------ */
/* Đối chiếu AI ↔ OCR cho trang scan (từ 22/10, theo cách thầy Hải     */
/* duyệt): AI chép trực tiếp từ ảnh là kết quả chính; OCR chạy nền     */
/* cho cùng trang và hai bản được so từng từ. Chỗ nào hai bên đọc      */
/* KHÁC nhau ở token chứa chữ số (số hiệu, ngày tháng, số tiền, số     */
/* lượng) thì gắn cờ để người dùng kiểm tra đúng chỗ đó trên bản gốc,  */
/* thay vì đọc lại cả trang. Khác nhau ở chữ thường (dấu tiếng Việt)   */
/* là điểm yếu đã biết của OCR nên không gắn cờ — tránh báo nhiễu.     */
/* ------------------------------------------------------------------ */

export interface CrossCheckResult {
  /** Tỷ lệ token phía AI khớp với OCR (0..1). */
  agreement: number;
  /** Token chứa chữ số chỉ phía AI có (đã khử trùng, tối đa 6). */
  aiDigitDiffs: string[];
  /** Token chứa chữ số chỉ phía OCR có (đã khử trùng, tối đa 6). */
  ocrDigitDiffs: string[];
  /** Có gắn cờ hay không: chỉ khi có token chứa chữ số bị khác. */
  flagged: boolean;
}

interface TokenPair {
  orig: string;
  norm: string;
}

function tokenize(text: string): TokenPair[] {
  return text
    .split(/\s+/)
    .map((t) => t.trim())
    .filter(Boolean)
    .map((orig) => ({
      orig,
      norm: orig
        .toLowerCase()
        .replace(/^[^\p{L}\p{N}]+|[^\p{L}\p{N}]+$/gu, ""),
    }))
    .filter((t) => t.norm.length > 0);
}

function digitDiffsOf(pairs: TokenPair[], otherNorms: Map<string, number>): string[] {
  const remaining = new Map(otherNorms);
  const diffs: string[] = [];
  const seen = new Set<string>();
  for (const t of pairs) {
    const c = remaining.get(t.norm) || 0;
    if (c > 0) {
      remaining.set(t.norm, c - 1);
    } else if (/\d/.test(t.orig) && !seen.has(t.norm)) {
      seen.add(t.norm);
      diffs.push(t.orig);
    }
  }
  return diffs.slice(0, 6);
}

export function crossCheckPage(aiText: string, ocrText: string): CrossCheckResult {
  const aiPairs = tokenize(aiText);
  const ocrPairs = tokenize(ocrText);

  const ocrCounts = new Map<string, number>();
  for (const t of ocrPairs) ocrCounts.set(t.norm, (ocrCounts.get(t.norm) || 0) + 1);
  const aiCounts = new Map<string, number>();
  for (const t of aiPairs) aiCounts.set(t.norm, (aiCounts.get(t.norm) || 0) + 1);

  const remaining = new Map(ocrCounts);
  let matched = 0;
  for (const t of aiPairs) {
    const c = remaining.get(t.norm) || 0;
    if (c > 0) {
      remaining.set(t.norm, c - 1);
      matched++;
    }
  }
  const agreement = aiPairs.length > 0 ? matched / aiPairs.length : 1;

  const aiDigitDiffs = digitDiffsOf(aiPairs, ocrCounts);
  const ocrDigitDiffs = digitDiffsOf(ocrPairs, aiCounts);

  return {
    agreement,
    aiDigitDiffs,
    ocrDigitDiffs,
    flagged: aiDigitDiffs.length > 0 || ocrDigitDiffs.length > 0,
  };
}

/** Dòng cảnh báo gắn cờ cho một trang (null nếu không cần gắn cờ). */
export function crossCheckWarning(pageNumber: number, result: CrossCheckResult): string | null {
  if (!result.flagged) return null;
  const parts: string[] = [];
  if (result.aiDigitDiffs.length > 0) {
    parts.push(`AI đọc: ${result.aiDigitDiffs.map((t) => `"${t}"`).join(", ")}`);
  }
  if (result.ocrDigitDiffs.length > 0) {
    parts.push(`OCR đọc: ${result.ocrDigitDiffs.map((t) => `"${t}"`).join(", ")}`);
  }
  return `Trang ${pageNumber}: đối chiếu AI ↔ OCR thấy khác nhau ở số liệu (${parts.join("; ")}). Hãy kiểm tra các số này trên bản PDF gốc trước khi ban hành.`;
}
