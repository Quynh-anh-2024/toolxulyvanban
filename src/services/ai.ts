export type AIProvider = "gemini" | "openrouter";

function getEnvValue(name: string): string {
  const env = import.meta.env as Record<string, string | undefined>;
  return (env[name] || "").trim();
}

function getConfiguredProvider(): AIProvider {
  const provider = getEnvValue("VITE_AI_PROVIDER").toLowerCase();
  return provider === "openrouter" ? "openrouter" : "gemini";
}

async function runGemini(prompt: string, apiKey: string): Promise<string> {
  const response = await fetch(
    `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=${encodeURIComponent(apiKey)}`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        contents: [
          {
            role: "user",
            parts: [{ text: prompt }],
          },
        ],
        generationConfig: {
          temperature: 0.25,
          topP: 0.9,
          maxOutputTokens: 8192,
        },
      }),
    }
  );

  const data = await response.json().catch(() => null);

  if (!response.ok) {
    const message =
      data?.error?.message ||
      `Gemini API lỗi ${response.status}. Kiểm tra lại API key hoặc hạn mức trên Cloudflare.`;
    throw new Error(message);
  }

  const text = data?.candidates?.[0]?.content?.parts
    ?.map((part: { text?: string }) => part.text || "")
    .join("\n")
    .trim();

  if (!text) {
    throw new Error("Gemini không trả về nội dung. Vui lòng thử lại hoặc rút ngắn văn bản.");
  }

  return text;
}

async function runOpenRouter(prompt: string, apiKey: string): Promise<string> {
  const response = await fetch("https://openrouter.ai/api/v1/chat/completions", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json",
      "HTTP-Referer": window.location.origin,
      "X-Title": "Chuan hoa van ban va soat loi AI",
    },
    body: JSON.stringify({
      model: getEnvValue("VITE_OPENROUTER_MODEL") || "google/gemini-flash-1.5",
      messages: [{ role: "user", content: prompt }],
      temperature: 0.25,
      max_tokens: 8192,
    }),
  });

  const data = await response.json().catch(() => null);

  if (!response.ok || data?.error) {
    const message =
      data?.error?.message ||
      `OpenRouter API lỗi ${response.status}. Kiểm tra lại API key hoặc hạn mức trên Cloudflare.`;
    throw new Error(message);
  }

  const text = data?.choices?.[0]?.message?.content?.trim();
  if (!text) {
    throw new Error("OpenRouter không trả về nội dung. Vui lòng thử lại hoặc đổi model/API key.");
  }

  return text;
}

export async function runAI(prompt: string): Promise<string> {
  try {
    const provider = getConfiguredProvider();
    const geminiKey = getEnvValue("VITE_GEMINI_API_KEY");
    const openRouterKey = getEnvValue("VITE_OPENROUTER_API_KEY");

    if (provider === "openrouter") {
      if (!openRouterKey) {
        throw new Error(
          "Chưa cấu hình VITE_OPENROUTER_API_KEY trên Cloudflare Pages. Hãy thêm biến môi trường rồi deploy lại."
        );
      }
      return await runOpenRouter(prompt, openRouterKey);
    }

    if (!geminiKey) {
      throw new Error(
        "Chưa cấu hình VITE_GEMINI_API_KEY trên Cloudflare Pages. Hãy thêm biến môi trường rồi deploy lại."
      );
    }

    return await runGemini(prompt, geminiKey);
  } catch (error) {
    throw new Error(error instanceof Error ? error.message : "Lỗi kết nối AI hoặc máy chủ quá tải.");
  }
}
