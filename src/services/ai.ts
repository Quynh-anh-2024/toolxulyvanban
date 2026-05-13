export type AIProvider = "gemini" | "openrouter";

function getEnvValue(name: string): string {
  const env = import.meta.env as Record<string, string | undefined>;
  return (env[name] || "").trim();
}

export async function runAI(prompt: string): Promise<string> {
  try {
    const provider = getEnvValue("VITE_AI_PROVIDER").toLowerCase() === "openrouter" ? "openrouter" : "gemini";

    // Gọi đến trạm trung chuyển nội bộ của bạn thay vì gọi thẳng Google
    const response = await fetch("/api/ai", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ provider, prompt }),
    });

    const data = await response.json();

    if (!response.ok) {
      throw new Error(data?.error || `Lỗi máy chủ trung gian (${response.status})`);
    }

    // Xử lý kết quả trả về tùy theo provider
    if (provider === "gemini") {
      const text = data?.candidates?.[0]?.content?.parts?.[0]?.text;
      if (!text) throw new Error("AI không phản hồi nội dung.");
      return text.trim();
    } else {
      const text = data?.choices?.[0]?.message?.content;
      if (!text) throw new Error("OpenRouter không phản hồi nội dung.");
      return text.trim();
    }
  } catch (error) {
    throw new Error(error instanceof Error ? error.message : "Lỗi kết nối bảo mật.");
  }
}
