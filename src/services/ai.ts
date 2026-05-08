export async function runAI(prompt: string): Promise<string> {
  try {
    // Vite bắt buộc dùng import.meta.env và phải có tiền tố VITE_
    const apiKey = import.meta.env.VITE_GROQ_API_KEY;
    
    if (!apiKey) {
        throw new Error("Chưa nạp API Key vào hệ thống. Vui lòng kiểm tra lại Netlify.");
    }
    
    const response = await fetch("https://api.groq.com/openai/v1/chat/completions", {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${apiKey}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        model: "llama3-8b-8192", // Tốc độ siêu nhanh
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
