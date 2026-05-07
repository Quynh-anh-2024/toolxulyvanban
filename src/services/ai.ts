import { GoogleGenAI } from "@google/genai";

// Initialize the Gemini client using the environment variable exposed by Vite
const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

export async function runAI(prompt: string): Promise<string> {
  try {
    const response = await ai.models.generateContent({
      model: "gemini-2.0-flash",
      contents: prompt,
    });
    return response.text || "Không có kết quả. Vui lòng thử lại.";
  } catch (error) {
    throw new Error(
      error instanceof Error ? error.message : "Lỗi kết nối AI hoặc máy chủ quá tải."
    );
  }
}
