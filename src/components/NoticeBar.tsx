// Thanh báo trong trang (thay alert từ đợt A3): đỏ = lỗi, vàng = cảnh
// báo, xanh = xong; tự tắt sau 7 giây (App quản lý hẹn giờ) hoặc bấm ×.
import { AlertTriangle, CheckCircle2, X } from "lucide-react";
import { cn } from "../lib/utils";

export interface NoticeData {
  kind: "error" | "warn" | "ok";
  text: string;
}

export function NoticeBar({
  notice,
  onClose,
}: {
  notice: NoticeData | null;
  onClose: () => void;
}) {
  if (!notice) return null;
  return (
  <div
    className={cn(
      "fixed bottom-4 left-1/2 z-50 flex w-[calc(100%-2rem)] max-w-xl -translate-x-1/2 items-start gap-2 rounded-lg px-4 py-3 text-sm text-white shadow-2xl",
      notice.kind === "error" && "bg-red-600",
      notice.kind === "warn" && "bg-amber-500",
      notice.kind === "ok" && "bg-emerald-600"
    )}
    role="alert"
  >
    {notice.kind === "ok" ? (
      <CheckCircle2 className="mt-0.5 h-4 w-4 shrink-0" />
    ) : (
      <AlertTriangle className="mt-0.5 h-4 w-4 shrink-0" />
    )}
    <span className="min-w-0 flex-1">{notice.text}</span>
    <button type="button" aria-label="Đóng thông báo" onClick={onClose} className="shrink-0 opacity-80 hover:opacity-100">
      <X className="h-4 w-4" />
    </button>
  </div>
  );
}
