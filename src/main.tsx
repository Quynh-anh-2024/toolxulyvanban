import {StrictMode} from 'react';
import {createRoot} from 'react-dom/client';
import App from './App.tsx';
import './index.css';

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
  </StrictMode>,
);

// PWA: đăng ký service worker (chỉ ở bản production) để cài app ra màn hình
// chính và dùng offline. Lỗi đăng ký không được làm hỏng app — bỏ qua lặng lẽ.
if (import.meta.env.PROD && 'serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker
      .register(`${import.meta.env.BASE_URL}sw.js`)
      .catch(() => {
        /* không có service worker thì app vẫn chạy bình thường khi có mạng */
      });
  });
}
