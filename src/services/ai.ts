export type AIProvider = "gemini";

function extractGeminiText(data: any): string {
  return (
    data?.candidates?.[0]?.content?.parts
      ?.map((part: { text?: string }) => part.text || "")
      .join("\n")
      .trim() || ""
  );
}

export async function runAI(prompt: string): Promise<string> {
  const response = await fetch("/api/ai", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ provider: "gemini", prompt }),
  });

  const data = await response.json().catch(() => null);

  if (!response.ok) {
    throw new Error(
      data?.error ||
        `Máy chủ AI lỗi ${response.status}. Vui lòng kiểm tra biến GEMINI_API_KEY/VITE_GEMINI_API_KEY trên Cloudflare và Redeploy.`
    );
  }

  const text = data?.text || extractGeminiText(data);
  if (!text) {
    throw new Error("Gemini không trả về nội dung. Vui lòng thử lại hoặc rút ngắn văn bản.");
  }

  return text;
}
