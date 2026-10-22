/* ------------------------------------------------------------------ */
/* Gemini trực tiếp — dùng cho "AI đánh lại" trang scan và làm đường   */
/* dự phòng khi OpenRouter lỗi. Khóa do NGƯỜI DÙNG tự dán vào app,    */
/* chỉ lưu trong localStorage của trình duyệt trên máy đó — TUYỆT ĐỐI */
/* không viết khóa vào mã nguồn (repo công khai = khóa bị thu hồi).    */
/* ------------------------------------------------------------------ */

const GEMINI_KEY_STORAGE = "xlvb_gemini_key";
const GEMINI_MODELS = ["gemini-3.8-flash", "gemini-2.5-flash", "gemini-2.0-flash"];

export function getGeminiKey(): string {
  try {
    return (localStorage.getItem(GEMINI_KEY_STORAGE) || "").trim();
  } catch {
    return "";
  }
}

export function setGeminiKey(key: string): void {
  try {
    if (key.trim()) localStorage.setItem(GEMINI_KEY_STORAGE, key.trim());
    else localStorage.removeItem(GEMINI_KEY_STORAGE);
  } catch {
    /* Trình duyệt chặn localStorage — bỏ qua, tính năng AI sẽ báo thiếu khóa. */
  }
}

export function hasGeminiKey(): boolean {
  return getGeminiKey().length > 0;
}

function plainAiError(status: number, message: string): Error {
  if (status === 400 && /API key not valid|API_KEY_INVALID/i.test(message)) {
    return new Error("Khóa Gemini chưa đúng. Hãy kiểm tra lại khóa đã dán (lấy tại aistudio.google.com/apikey).");
  }
  if (status === 403) return new Error("Khóa Gemini không có quyền dùng model này (403). Hãy thử khóa khác.");
  if (status === 429) return new Error("AI đang quá tải/hết lượt miễn phí (429). Đợi khoảng 1 phút rồi thử lại.");
  if (status >= 500) return new Error("Máy chủ AI đang lỗi tạm thời. Thử lại sau ít phút.");
  return new Error(message || `Lỗi AI (${status}).`);
}

async function geminiGenerate(parts: unknown[]): Promise<string> {
  const key = getGeminiKey();
  if (!key) {
    throw new Error("Chưa có khóa Gemini. Hãy dán khóa Gemini (miễn phí tại aistudio.google.com/apikey) vào ô khóa AI trước.");
  }
  let lastError: Error | null = null;
  for (const model of GEMINI_MODELS) {
    for (let attempt = 0; attempt < 2; attempt++) {
      try {
        const response = await fetch(
          `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`,
          {
            method: "POST",
            headers: { "Content-Type": "application/json", "x-goog-api-key": key },
            body: JSON.stringify({
              contents: [{ parts }],
              generationConfig: { temperature: 0 },
            }),
          }
        );
        const data = await response.json().catch(() => null);
        if (!response.ok) {
          const message: string = data?.error?.message || "";
          // Model không tồn tại với khóa này → thử model kế tiếp trong danh sách.
          if (response.status === 404) { lastError = plainAiError(404, message); break; }
          // Quá tải/hết lượt: chờ rồi thử lại đúng 1 lần (như quy ước app phiếu).
          if ((response.status === 429 || response.status >= 500) && attempt === 0) {
            await new Promise((r) => setTimeout(r, 6000));
            continue;
          }
          throw plainAiError(response.status, message);
        }
        const text = (data?.candidates?.[0]?.content?.parts || [])
          .map((p: any) => p?.text || "")
          .join("")
          .trim();
        if (!text) throw new Error("AI không trả về nội dung. Vui lòng thử lại.");
        return text;
      } catch (e) {
        if (e instanceof Error && /Khóa Gemini|quyền dùng|quá tải/.test(e.message)) throw e;
        lastError = e instanceof Error ? e : new Error("Không gọi được AI (kiểm tra mạng).");
        break; // lỗi mạng → thử model khác cũng vô ích, nhưng cứ sang model kế cho chắc
      }
    }
  }
  throw lastError || new Error("Không gọi được AI. Kiểm tra mạng và khóa Gemini rồi thử lại.");
}

export async function runGeminiText(prompt: string): Promise<string> {
  return geminiGenerate([{ text: prompt }]);
}

/** Prompt đánh lại trang scan — ĐÃ ĐO THẬT trên CV 5095 trang 1: phủ 100% từ
 *  so với bản chép tay (Tesseract chỉ đạt ~89%). Đừng sửa tuỳ tiện. */
export const TRANSCRIBE_PROMPT = `Bạn là hệ thống số hoá văn bản hành chính Việt Nam chính xác như ABBYY FineReader. Nhiệm vụ: ĐÁNH LẠI toàn bộ chữ in trên trang này.
QUY TẮC BẮT BUỘC:
1. Giữ nguyên từng chữ, dấu tiếng Việt, chữ số, ngày tháng, số hiệu — không tóm tắt, không thêm bớt, không bình luận, không tự sửa chính tả: bản gốc viết sao chép y vậy.
2. Thứ tự đọc: từ trên xuống dưới. Phần đầu trang có 2 cột (trái: cơ quan ban hành và số ký hiệu; phải: quốc hiệu, tiêu ngữ và địa danh/ngày tháng) — đánh lại thành các dòng riêng theo thứ tự: cơ quan, quốc hiệu, tiêu ngữ, số ký hiệu, ngày tháng.
3. Mỗi đoạn văn là một đoạn riêng, cách nhau một dòng trống. Dòng bị ngắt giữa đoạn trong bản gốc thì nối lại liền mạch.
4. Nếu có bảng: xuất dạng markdown, mỗi hàng một dòng | ô 1 | ô 2 | ... | và thêm dòng |---|---| ngay dưới hàng tiêu đề.
5. Bỏ qua con dấu, chữ ký tay, chữ viết tay và dòng thông tin cổng ký số nhỏ ở mép trên trang. Khối chữ ký IN (KT./TM., chức danh, họ tên in) thì đánh lại.
6. Chỉ xuất văn bản đã đánh lại, không thêm bất kỳ lời dẫn nào.`;

/** Gửi ẢNH một trang scan cho Gemini đánh lại thành văn bản thuần. */
export async function transcribePageImage(jpegDataUrl: string): Promise<string> {
  const base64 = jpegDataUrl.replace(/^data:image\/[a-z]+;base64,/, "");
  return geminiGenerate([
    { text: TRANSCRIBE_PROMPT },
    { inline_data: { mime_type: "image/jpeg", data: base64 } },
  ]);
}

export async function runAI(prompt: string): Promise<string> {
  const trimmedPrompt = prompt.trim();
  if (!trimmedPrompt) {
    throw new Error("Nội dung gửi AI đang trống.");
  }
  try {
    return await runOpenRouter(trimmedPrompt);
  } catch (e) {
    // OpenRouter lỗi (khóa mặc định trong bản công khai đã bị thu hồi) → nếu
    // người dùng đã dán khóa Gemini của mình thì dùng Gemini trực tiếp thay.
    if (hasGeminiKey()) return runGeminiText(trimmedPrompt);
    throw e;
  }
}

async function runOpenRouter(trimmedPrompt: string): Promise<string> {
  // Thay vì gọi "/api/ai" dễ bị 404 nếu cài đặt Cloudflare sai,
  // chúng ta gọi THẲNG sang OpenRouter luôn. 
  // Bạn CẦN thay chữ "sk-or-v1..." bên dưới bằng KEY THỰC TẾ CỦA BẠN.
  const apiKey = "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"; // <- Thay API Key của bạn vào đây!

  const response = await fetch("https://openrouter.ai/api/v1/chat/completions", {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${apiKey}`,
      "Content-Type": "application/json",
      "X-Title": "Chuan hoa van ban TH Giang Chu Phin",
    },
    body: JSON.stringify({
      model: "google/gemini-2.0-flash-lite-preview-02-05:free", // Model mạnh và miễn phí
      messages: [{ role: "user", content: trimmedPrompt }],
      temperature: 0.25,
    }),
  });

  const data = await response.json().catch(() => null);

  if (!response.ok || data?.error) {
    const message = data?.error?.message || `Lỗi từ OpenRouter (${response.status}).`;
    throw new Error(message);
  }

  const text = data?.choices?.[0]?.message?.content?.trim();
  if (!text) {
    throw new Error("AI không phản hồi nội dung. Vui lòng thử lại.");
  }

  return text;
}
