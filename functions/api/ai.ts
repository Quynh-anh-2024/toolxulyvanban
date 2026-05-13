export const onRequestPost = async (context: any) => {
  const { request, env } = context;

  try {
    const body = await request.json();
    const { provider, prompt } = body;

    const geminiKey = env.GEMINI_API_KEY;
    const openRouterKey = env.OPENROUTER_API_KEY;

    // Hàm bọc lỗi chuẩn JSON để web không bị Crash
    const jsonError = (msg: string, status = 500) => {
      return new Response(JSON.stringify({ error: msg }), { 
        status, 
        headers: { "Content-Type": "application/json" } 
      });
    };

    if (provider === "gemini") {
      if (!geminiKey) return jsonError("Thiếu GEMINI_API_KEY. Vui lòng kiểm tra lại tab Variables trên Cloudflare và nhấn Redeploy.");

      // Cập nhật model Google chuẩn xác nhất hiện hành
      const response = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${geminiKey}`,
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
      return new Response(data, { status: response.status, headers: { "Content-Type": "application/json" } });
    } 
    
    else {
      if (!openRouterKey) return jsonError("Thiếu OPENROUTER_API_KEY. Vui lòng kiểm tra lại tab Variables trên Cloudflare và nhấn Redeploy.");

      // ÉP CỨNG MODEL: Dùng mô hình Gemini 2.0 Flash Lite Miễn phí và ổn định nhất của OpenRouter hiện nay (Tránh lỗi model cũ bị gỡ)
      const stableModel = "google/gemini-2.0-flash-lite-preview-02-05:free";

      const response = await fetch("https://openrouter.ai/api/v1/chat/completions", {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${openRouterKey}`,
          "Content-Type": "application/json",
          "X-Title": "Chuan hoa van ban TH Giang Chu Phin",
        },
        body: JSON.stringify({
          model: stableModel,
          messages: [{ role: "user", content: prompt }],
          temperature: 0.25,
        }),
      });
      const data = await response.text();
      return new Response(data, { status: response.status, headers: { "Content-Type": "application/json" } });
    }
  } catch (error: any) {
    return new Response(JSON.stringify({ error: error.message }), { 
      status: 500, headers: { "Content-Type": "application/json" } 
    });
  }
};
