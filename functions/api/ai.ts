export const onRequestPost = async (context: any) => {
  const { request, env } = context;

  const json = (payload: Record<string, unknown>, status = 200) =>
    new Response(JSON.stringify(payload), {
      status,
      headers: {
        "Content-Type": "application/json; charset=utf-8",
        "Cache-Control": "no-store",
      },
    });

  try {
    const body = await request.json().catch(() => ({}));
    const prompt = String(body?.prompt || "").trim();

    if (!prompt) return json({ error: "Thiếu nội dung để AI xử lý." }, 400);

    // Chấp nhận cả 2 tên biến để không bắt bạn phải đổi lại biến đã nhập từ sáng.
    const apiKey = String(env.GEMINI_API_KEY || env.VITE_GEMINI_API_KEY || "").trim();
    const model = String(env.GEMINI_MODEL || env.VITE_GEMINI_MODEL || "gemini-3-flash-preview").trim();

    if (!apiKey) {
      return json(
        {
          error:
            "Thiếu API Gemini trên Cloudflare. Hãy kiểm tra tab Variables: GEMINI_API_KEY hoặc VITE_GEMINI_API_KEY, sau đó Redeploy.",
        },
        500
      );
    }

    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(
      model
    )}:generateContent?key=${encodeURIComponent(apiKey)}`;

    const geminiResponse = await fetch(endpoint, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        contents: [{ role: "user", parts: [{ text: prompt }] }],
        generationConfig: {
          temperature: 0.25,
          topP: 0.9,
          maxOutputTokens: 8192,
        },
      }),
    });

    const data: any = await geminiResponse.json().catch(() => null);

    if (!geminiResponse.ok) {
      const message = data?.error?.message || `Gemini API lỗi ${geminiResponse.status}.`;
      return json({ error: message, model }, geminiResponse.status);
    }

    const text =
      data?.candidates?.[0]?.content?.parts
        ?.map((part: { text?: string }) => part.text || "")
        .join("\n")
        .trim() || "";

    return json({ text, model });
  } catch (error: any) {
    return json({ error: error?.message || "Lỗi máy chủ AI." }, 500);
  }
};
