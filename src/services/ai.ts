export async function runAI(prompt: string): Promise<string> {
  const trimmedPrompt = prompt.trim();
  if (!trimmedPrompt) {
    throw new Error("Nội dung gửi AI đang trống.");
  }

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
