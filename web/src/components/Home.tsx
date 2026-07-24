import { useEffect, useState } from 'react';
import { api } from '../api';
import type { Dashboard } from '../types';
import { fmt, fmtInt } from '../utils';
import OrderRow from './OrderRow';

export default function Home() {
  const [d, setD] = useState<Dashboard | null>(null);
  const [err, setErr] = useState<string | null>(null);

  useEffect(() => {
    api
      .getDashboard()
      .then((r) => setD(r as Dashboard))
      .catch((e) => setErr(e instanceof Error ? e.message : String(e)));
  }, []);

  if (err) return <div className="error-banner">加载失败: {err}</div>;

  if (!d) {
    return (
      <div className="card">
        <div className="skeleton skeleton-line" style={{ width: '60%' }} />
        <div className="skeleton" style={{ height: 90 }} />
      </div>
    );
  }

  return (
    <>
      <div className="card">
        <h2>
          {d.month} 本月概览
        </h2>
        <div className="stats">
          <div className="stat brand">
            <div className="k">本月总收入</div>
            <div className="v">
              <small>RM</small>
              {fmt(d.monthTotal)}
            </div>
          </div>
          <div className="stat green">
            <div className="k">本月给司机</div>
            <div className="v">
              <small>RM</small>
              {fmt(d.monthCommission)}
            </div>
          </div>
          <div className="stat neutral">
            <div className="k">本月订单</div>
            <div className="v">{fmtInt(d.monthCount)}</div>
          </div>
        </div>
        <div className="hint">
          累计: {fmtInt(d.allCount)} 单 · 总收入 RM{fmt(d.allTotal)} · 给司机 RM{fmt(d.allCommission)}
        </div>
      </div>

      <div className="card">
        <h2>最近订单</h2>
        {d.recent.length === 0 ? (
          <div className="muted-block">还没有订单</div>
        ) : (
          d.recent.map((o) => <OrderRow key={o.id} order={o} />)
        )}
      </div>
    </>
  );
}
