export type AIProvider = "gemini" | "openrouter";

export async function runAI(prompt: string): Promise<string> {
  try {
    // Ép cứng hệ thống LUÔN LUÔN dùng OpenRouter (Vì chúng ta đã có Key này)
    const provider = "openrouter"; 

    const response = await fetch("/api/ai", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ provider, prompt }),
    });

    // Bắt lỗi an toàn, chống văng app
    const textData = await response.text();
    let data;
    try {
      data = JSON.parse(textData);
    } catch {
      throw new Error(textData || `Lỗi máy chủ bất thường (${response.status})`);
    }

    if (!response.ok) {
      throw new Error(data?.error || `Lỗi máy chủ trung gian (${response.status})`);
    }

    // Trả kết quả từ OpenRouter
    const text = data?.choices?.[0]?.message?.content;
    if (!text) throw new Error("OpenRouter không phản hồi nội dung. Hãy thử lại.");
    return text.trim();
    
  } catch (error) {
    throw new Error(error instanceof Error ? error.message : "Lỗi kết nối máy chủ AI.");
  }
}
