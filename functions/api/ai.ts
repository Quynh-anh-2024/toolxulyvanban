type Env = Record<string, string | undefined>;

const DEFAULT_MODEL = "gemini-3-flash-preview";

function json(data: unknown, status = 200): Response {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type",
    },
  });
}

function getEnv(env: Env, names: string[]): string {
  for (const name of names) {
    const value = env?.[name];
    if (typeof value === "string" && value.trim()) return value.trim();
  }
  return "";
}

function extractGeminiText(data: any): string {
  return (
    data?.candidates?.[0]?.content?.parts
      ?.map((part: { text?: string }) => part.text || "")
      .join("\n")
      .trim() || ""
  );
}

export const onRequestOptions = async () => json({ ok: true });

export const onRequestPost = async (context: any) => {
  const { request, env } = context as { request: Request; env: Env };

  try {
    const body = await request.json().catch(() => ({}));
    const prompt = String(body?.prompt || "").trim();

    if (!prompt) {
      return json({ error: "Nội dung gửi AI đang trống." }, 400);
    }

    const apiKey = getEnv(env, [
      "GEMINI_API_KEY",
      "VITE_GEMINI_API_KEY",
      "GOOGLE_API_KEY",
      "AI_API_KEY",
    ]);

    if (!apiKey) {
      return json(
        {
          error:
            "Chưa tìm thấy API Gemini trong Cloudflare. Hãy vào Pages → Settings → Variables and Secrets và thêm một trong các biến: GEMINI_API_KEY hoặc VITE_GEMINI_API_KEY, sau đó Redeploy bản mới nhất.",
        },
        500
      );
    }

    const model =
      getEnv(env, ["GEMINI_MODEL", "VITE_GEMINI_MODEL"]) || DEFAULT_MODEL;

    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(
      model
    )}:generateContent`;

    const aiResponse = await fetch(endpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-goog-api-key": apiKey,
      },
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
          thinkingConfig: { thinkingLevel: "low" },
        },
      }),
    });

    const aiRawText = await aiResponse.text();
    let aiData: any = null;
    try {
      aiData = aiRawText ? JSON.parse(aiRawText) : null;
    } catch {
      aiData = null;
    }

    if (!aiResponse.ok) {
      const message =
        aiData?.error?.message ||
        aiData?.error ||
        aiRawText ||
        `Gemini API trả về lỗi ${aiResponse.status}.`;
      return json({ error: message, model }, aiResponse.status);
    }

    const text = extractGeminiText(aiData);
    if (!text) {
      return json(
        {
          error:
            "Gemini đã phản hồi nhưng không có nội dung văn bản. Hãy thử rút ngắn tài liệu hoặc chạy lại.",
          model,
        },
        502
      );
    }

    return json({ text, model });
  } catch (error: any) {
    return json(
      { error: error?.message || "Lỗi máy chủ AI. Vui lòng thử lại." },
      500
    );
  }
};
