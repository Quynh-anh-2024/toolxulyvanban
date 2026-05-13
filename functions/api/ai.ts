// Đây là code chạy trên máy chủ Cloudflare, người dùng không thể xem được
export const onRequestPost = async (context: any) => {
  const { request, env } = context;

  try {
    const body = await request.json();
    const { provider, prompt } = body;

    // Lấy Key từ biến môi trường bí mật (không có chữ VITE_)
    const geminiKey = env.GEMINI_API_KEY;
    const openRouterKey = env.OPENROUTER_API_KEY;

    if (provider === "gemini") {
      if (!geminiKey) return new Response("Thiếu Gemini Key trên máy chủ", { status: 500 });

      const response = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=${geminiKey}`,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            contents: [{ role: "user", parts: [{ text: prompt }] }],
            safetySettings: [
              { category: "HARM_CATEGORY_HARASSMENT", threshold: "BLOCK_NONE" },
              { category: "HARM_CATEGORY_HATE_SPEECH", threshold: "BLOCK_NONE" },
              { category: "HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold: "BLOCK_NONE" },
              { category: "HARM_CATEGORY_DANGEROUS_CONTENT", threshold: "BLOCK_NONE" },
            ],
            generationConfig: { temperature: 0.25, maxOutputTokens: 8192 }
          }),
        }
      );
      const data = await response.text();
      return new Response(data, { headers: { "Content-Type": "application/json" } });
    } 
    
    else {
      if (!openRouterKey) return new Response("Thiếu OpenRouter Key trên máy chủ", { status: 500 });

      const response = await fetch("https://openrouter.ai/api/v1/chat/completions", {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${openRouterKey}`,
          "Content-Type": "application/json",
          "X-Title": "Chuan hoa van ban TH Giang Chu Phin",
        },
        body: JSON.stringify({
          model: env.VITE_OPENROUTER_MODEL || "google/gemini-2.0-flash-001",
          messages: [{ role: "user", content: prompt }],
          temperature: 0.25,
        }),
      });
      const data = await response.text();
      return new Response(data, { headers: { "Content-Type": "application/json" } });
    }
  } catch (error: any) {
    return new Response(JSON.stringify({ error: error.message }), { status: 500 });
  }
};
