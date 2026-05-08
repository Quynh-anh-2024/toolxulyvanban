// Không cần import thư viện Google nữa, chúng ta gọi thẳng đến máy chủ Groq

export async function runAI(prompt: string): Promise<string> {
  try {
    // Lấy API Key từ Netlify
    const apiKey = process.env.GROQ_API_KEY;
    
    const response = await fetch("https://api.groq.com/openai/v1/chat/completions", {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${apiKey}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        model: "llama3-8b-8192", // Mô hình siêu nhanh, thông minh và không chặn IP Việt Nam
        messages: [{ role: "user", content: prompt }]
      })
    });

    const data = await response.json();
    
    // Nếu có lỗi từ máy chủ
    if (data.error) {
       throw new Error(data.error.message);
    }

    // Trả về kết quả văn bản
    return data.choices[0].message.content || "Không có kết quả. Vui lòng thử lại.";
    
  } catch (error) {
    throw new Error(
      error instanceof Error ? error.message : "Lỗi kết nối AI hoặc máy chủ quá tải."
    );
  }
}
