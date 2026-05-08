export async function runAI(prompt: string): Promise<string> {
  try {
    // Gọi khóa OpenRouter từ Netlify
    const apiKey = import.meta.env.VITE_OPENROUTER_API_KEY;
    
    if (!apiKey) {
        throw new Error("Chưa nạp API Key OpenRouter vào hệ thống Netlify.");
    }
    
    // Gửi tín hiệu sang trạm trung chuyển tại Mỹ
    const response = await fetch("https://openrouter.ai/api/v1/chat/completions", {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${apiKey}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        // Sử dụng bộ não khổng lồ Gemini 2.0 (Bản miễn phí qua OpenRouter)
        model: "microsoft/phi-3-medium-128k-instruct:free", 
        messages: [{ role: "user", content: prompt }]
      })
    });

    const data = await response.json();
    
    if (data.error) {
       throw new Error(data.error.message);
    }

    return data.choices[0].message.content || "Không có kết quả. Vui lòng thử lại.";
    
  } catch (error) {
    throw new Error(
      error instanceof Error ? error.message : "Lỗi kết nối AI hoặc máy chủ quá tải."
    );
  }
}
