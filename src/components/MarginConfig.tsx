// Khối "Căn Lề NĐ 30" dùng chung cho cả 3 mục (áp dụng khi tải Word).
// Tách khỏi App.tsx (đợt B5).
import type { FormattingConfig } from "../utils/docx";

export function MarginConfig({
  config,
  setConfig,
}: {
  config: FormattingConfig;
  setConfig: (config: FormattingConfig) => void;
}) {
  return (
  <section className="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
    <h2 className="text-base font-semibold mb-3 text-blue-700">Căn Lề NĐ 30 (áp dụng khi tải Word)</h2>
    <div className="space-y-3">
      <div className="grid grid-cols-2 gap-3">
        <div>
          <label className="block text-xs text-slate-600 mb-1">Font chữ</label>
          <select className="w-full border p-2 rounded text-sm" value={config.font} onChange={(event) => setConfig({ ...config, font: event.target.value })}>
            <option value="Times New Roman">Times New Roman</option>
            <option value="Arial">Arial</option>
          </select>
        </div>
        <div>
          <label className="block text-xs text-slate-600 mb-1">Cỡ chữ</label>
          <input type="number" className="w-full border p-2 rounded text-sm" value={config.size} onChange={(event) => setConfig({ ...config, size: parseFloat(event.target.value) || 14 })} />
        </div>
      </div>
      <div className="grid grid-cols-2 gap-3 bg-emerald-50 p-2 rounded border border-emerald-100">
        <div>
          <label className="block text-xs text-slate-600 mb-1">Thụt dòng (cm)</label>
          <input type="number" step="0.1" className="w-full border p-2 rounded text-sm" value={config.textIndent} onChange={(event) => setConfig({ ...config, textIndent: parseFloat(event.target.value) || 0 })} />
        </div>
        <div>
          <label className="block text-xs text-slate-600 mb-1">Cách đoạn (pt)</label>
          <input type="number" className="w-full border p-2 rounded text-sm" value={config.paraSpacing} onChange={(event) => setConfig({ ...config, paraSpacing: parseFloat(event.target.value) || 0 })} />
        </div>
      </div>
      <div className="grid grid-cols-3 gap-2">
        <div>
          <label className="block text-xs text-slate-600 mb-1">Giãn dòng</label>
          <select className="w-full border p-2 rounded text-sm" value={config.spacing} onChange={(event) => setConfig({ ...config, spacing: event.target.value })}>
            <option value="1.0">1.0</option>
            <option value="1.15">1.15</option>
            <option value="1.5">1.5</option>
          </select>
        </div>
        <div>
          <label className="block text-xs text-slate-600 mb-1">Lề Trái</label>
          <input type="number" step="0.5" className="w-full border p-2 rounded text-sm" value={config.leftMargin} onChange={(event) => setConfig({ ...config, leftMargin: parseFloat(event.target.value) || 3 })} />
        </div>
        <div>
          <label className="block text-xs text-slate-600 mb-1">Lề Phải</label>
          <input type="number" step="0.5" className="w-full border p-2 rounded text-sm" value={config.rightMargin} onChange={(event) => setConfig({ ...config, rightMargin: parseFloat(event.target.value) || 2 })} />
        </div>
      </div>
    </div>
  </section>
  );
}
