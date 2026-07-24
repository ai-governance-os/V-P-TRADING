import { useEffect, useState } from 'react';
import { api } from '../api';
import type { MonthlySummary } from '../types';
import { fmt } from '../utils';

export default function Summary() {
  const [months, setMonths] = useState<string[]>([]);
  const [month, setMonth] = useState('');
  const [res, setRes] = useState<MonthlySummary | null>(null);

  useEffect(() => {
    api.listAvailableMonths().then((m) => {
      setMonths(m);
      if (m.length) setMonth(m[0]);
    });
  }, []);

  useEffect(() => {
    if (!month) return;
    api.getMonthlySummary(month).then((r) => setRes(r as MonthlySummary));
  }, [month]);

  return (
    <>
      <div className="card">
        <h2>月度汇总</h2>
        <label>月份</label>
        <select value={month} onChange={(e) => setMonth(e.target.value)}>
          {months.map((m) => (
            <option key={m} value={m}>
              {m}
            </option>
          ))}
        </select>
        {res && (
          <>
            <div className="stats" style={{ marginTop: 12 }}>
              <div className="stat brand">
                <div className="k">当月总收入</div>
                <div className="v">
                  <small>RM</small>
                  {fmt(res.totalIncome)}
                </div>
              </div>
              <div className="stat green">
                <div className="k">给司机总额</div>
                <div className="v">
                  <small>RM</small>
                  {fmt(res.totalCommission)}
                </div>
              </div>
            </div>
            <div className="hint">{res.orderCount} 笔订单</div>
          </>
        )}
      </div>
      {res && (
        <div className="card scroll-x">
          <h2>按 Branch 细分</h2>
          <table>
            <thead>
              <tr>
                <th>Branch</th>
                <th>单数</th>
                <th>Total</th>
                <th>抽佣</th>
              </tr>
            </thead>
            <tbody>
              {res.byBranch.map((b) => (
                <tr key={b.branch}>
                  <td>{b.branch}</td>
                  <td>{b.count}</td>
                  <td>{fmt(b.total)}</td>
                  <td>{fmt(b.commission)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </>
  );
}
