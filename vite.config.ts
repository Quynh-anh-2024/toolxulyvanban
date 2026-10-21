import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [
    tailwindcss(),
    react()
  ],
  // Worker pdf.js (src/workers/pdfWorker.ts) là ES module (dùng import.meta)
  // nên phải xuất dạng module, không dùng dạng iife mặc định.
  worker: {
    format: 'es'
  },
  base: '/', // Đường dẫn chuẩn cho Vercel
})
