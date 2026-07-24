import type { TabKey } from '../types';

const TABS: { key: TabKey; icon: string; label: string }[] = [
  { key: 'home', icon: '🏠', label: '首页' },
  { key: 'new', icon: '➕', label: '录单' },
  { key: 'records', icon: '🔍', label: '查找' },
  { key: 'summary', icon: '📊', label: '汇总' },
  { key: 'manage', icon: '⚙️', label: '管理' },
];

export default function BottomNav({ active, onChange }: { active: TabKey; onChange: (t: TabKey) => void }) {
  return (
    <div className="nav">
      <div className="nav-inner">
        {TABS.map((t) => (
          <div key={t.key} className={`n ${active === t.key ? 'on' : ''}`} onClick={() => onChange(t.key)}>
            <span className="ic">{t.icon}</span>
            {t.label}
          </div>
        ))}
      </div>
    </div>
  );
}
