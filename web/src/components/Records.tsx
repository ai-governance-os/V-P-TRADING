import { useEffect, useState } from 'react';
import { api } from '../api';
import { useAppData } from '../context/AppDataContext';
import type { Order } from '../types';
import OrderRow from './OrderRow';

export default function Records({ toast }: { toast: (msg: string) => void }) {
  const { data } = useAppData();
  const [months, setMonths] = useState<string[]>([]);
  const [kw, setKw] = useState('');
  const [month, setMonth] = useState('');
  const [region, setRegion] = useState('');
  const [rows, setRows] = useState<Order[] | null>(null);
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    api.listAvailableMonths().then(setMonths).catch(() => {});
  }, []);

  async function search() {
    setLoading(true);
    try {
      const r = (await api.searchOrders({ keyword: kw, month, region })) as Order[];
      setRows(r);
    } catch (e) {
      toast('查找出错: ' + (e instanceof Error ? e.message : String(e)));
    } finally {
      setLoading(false);
    }
  }

  async function del(id: number) {
    if (!confirm(`确定删除订单 #${id}？此操作无法撤销。`)) return;
    try {
      const res = (await api.deleteOrder(id)) as { ok: boolean };
      if (res.ok) {
        setRows((prev) => (prev ? prev.filter((r) => r.id !== id) : prev));
        toast(`已删除 #${id}`);
      } else {
        toast('删除失败，找不到 #' + id);
      }
    } catch (e) {
      toast('出错: ' + (e instanceof Error ? e.message : String(e)));
    }
  }

  return (
    <>
      <div className="card">
        <h2>查找订单</h2>
        <label>关键字 (名字或 branch)</label>
        <input placeholder="留空=不限" value={kw} onChange={(e) => setKw(e.target.value)} />
        <div className="row">
          <div>
            <label>月份</label>
            <select value={month} onChange={(e) => setMonth(e.target.value)}>
              <option value="">全部</option>
              {months.map((m) => (
                <option key={m} value={m}>
                  {m}
                </option>
              ))}
            </select>
          </div>
          <div>
            <label>地区</label>
            <select value={region} onChange={(e) => setRegion(e.target.value)}>
              <option value="">全部</option>
              {data &&
                Object.entries(data.regionLabels).map(([k, v]) => (
                  <option key={k} value={k}>
                    {v}
                  </option>
                ))}
            </select>
          </div>
        </div>
        <button className="btn btn-ghost" onClick={search} disabled={loading}>
          {loading ? '查找中…' : '查找'}
        </button>
      </div>
      <div className="card">
        {rows === null && <div className="muted-block">输入条件后按「查找」</div>}
        {rows !== null && rows.length === 0 && <div className="muted-block">没有找到记录</div>}
        {rows !== null && rows.map((r) => <OrderRow key={r.id} order={r} onDelete={del} />)}
      </div>
    </>
  );
}
