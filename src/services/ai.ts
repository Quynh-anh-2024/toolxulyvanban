export async function runAI(prompt: string): Promise<string> {
  try {
    // Ép cứng ứng dụng luôn dùng OpenRouter (vì bạn đã có khóa OpenRouter trên Cloudflare)
    const provider = "openrouter"; 

    const response = await fetch("/api/ai", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ provider, prompt }),
    });

    const textData = await response.text();
    let data;
    try {
      data = JSON.parse(textData);
    } catch {
      throw new Error(textData || `Lỗi kết nối máy chủ (${response.status})`);
    }

    if (!response.ok) {
      throw new Error(data?.error || `Lỗi AI (${response.status})`);
    }

    const text = data?.choices?.[0]?.message?.content;
    if (!text) throw new Error("AI không phản hồi nội dung. Hãy thử lại.");
    return text.trim();
    
  } catch (error) {
    throw new Error(error instanceof Error ? error.message : "Lỗi kết nối máy chủ AI.");
  }
}
