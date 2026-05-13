export async function runAI(prompt: string): Promise<string> {
  const trimmedPrompt = prompt.trim();
  if (!trimmedPrompt) {
    throw new Error("Nội dung gửi AI đang trống.");
  }

  const response = await fetch("/api/ai", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ prompt: trimmedPrompt }),
  });

  const rawText = await response.text();
  let data: any = null;
  try {
    data = rawText ? JSON.parse(rawText) : null;
  } catch {
    data = null;
  }

  if (!response.ok) {
    const message =
      data?.error ||
      data?.message ||
      rawText ||
      `Máy chủ AI trả về lỗi ${response.status}.`;
    throw new Error(message);
  }

  if (typeof data?.text === "string" && data.text.trim()) {
    return data.text.trim();
  }

  const geminiText = data?.candidates?.[0]?.content?.parts
    ?.map((part: { text?: string }) => part.text || "")
    .join("\n")
    .trim();

  if (geminiText) return geminiText;

  if (typeof rawText === "string" && rawText.trim()) {
    return rawText.trim();
  }

  throw new Error("AI không trả về nội dung. Vui lòng thử lại với văn bản ngắn hơn.");
}
