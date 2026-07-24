import { useEffect, useMemo, useRef, useState } from 'react';
import { useAppData } from '../context/AppDataContext';
import { api } from '../api';
import type { Salesman } from '../types';
import { brandOf, fmt, initials, todayISO } from '../utils';

const SET_TYPES = ['8 ITEMS', '9 ITEMS', '10 ITEMS', 'UMBRELLA', '其他'];
const AUTO_ORDER = ['9 ITEMS', '10 ITEMS', '8 ITEMS', 'UMBRELLA'];

export default function NewOrder({ toast }: { toast: (msg: string) => void }) {
  const { data, reload } = useAppData();
  const [query, setQuery] = useState('');
  const [open, setOpen] = useState(false);
  const [sel, setSel] = useState<Salesman | null>(null);
  const [setType, setSetType] = useState<string | null>(null);
  const [price, setPrice] = useState('');
  const [setQty, setSetQty] = useState('1');
  const [detail, setDetail] = useState('');
  const [date, setDate] = useState(todayISO());
  const [status, setStatus] = useState<'OP' | 'PC'>('OP');
  const [payDate, setPayDate] = useState('');
  const [calc, setCalc] = useState({ total: 0, commission: 0 });
  const [saving, setSaving] = useState(false);
  const [waMsg, setWaMsg] = useState<string | null>(null);
  const [addingNew, setAddingNew] = useState(false);
  const [newRegion, setNewRegion] = useState('');
  const [newBranch, setNewBranch] = useState('');
  const [newName, setNewName] = useState('');
  const comboRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    function onDocClick(e: MouseEvent) {
      if (comboRef.current && !comboRef.current.contains(e.target as Node)) setOpen(false);
    }
    document.addEventListener('click', onDocClick);
    return () => document.removeEventListener('click', onDocClick);
  }, []);

  useEffect(() => {
    if (data && !newRegion) setNewRegion(Object.keys(data.regionLabels)[0] || '');
  }, [data, newRegion]);

  const filtered = useMemo(() => {
    if (!data) return [];
    const kw = query.trim().toLowerCase();
    const list = kw
      ? data.salesmen.filter(
          (s) => s.name.toLowerCase().includes(kw) || s.branch.toLowerCase().includes(kw)
        )
      : data.salesmen;
    return list.slice(0, 40);
  }, [data, query]);

  function priceFor(branch: string, type: string): number | null {
    const p = data?.prices[branch];
    if (p && p[type] != null) return p[type];
    return null;
  }

  function pick(s: Salesman) {
    setSel(s);
    setOpen(false);
    setQuery('');
    // auto-pick set type with a known price
    let picked = '其他';
    for (const t of AUTO_ORDER) {
      if (priceFor(s.branch, t) != null) {
        picked = t;
        break;
      }
    }
    applySetType(picked, s);
  }

  function applySetType(t: string, s: Salesman | null = sel) {
    setSetType(t);
    if (!s) return;
    const p = priceFor(s.branch, t);
    if (p != null) setPrice(String(p));
    if (t !== '其他') setDetail(t.toLowerCase());
  }

  function itemKind(t: string | null) {
    return t === 'UMBRELLA' ? 'umbrella' : 'standard';
  }

  useEffect(() => {
    if (!sel || !setType) return;
    const region = sel.region;
    const item = itemKind(setType);
    api
      .previewAmounts(region, item, price, setQty)
      .then((r) => setCalc(r as { total: number; commission: number }))
      .catch(() => {});
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sel, setType, price, setQty]);

  async function handleSave() {
    if (!sel) return toast('请先选 salesman');
    if (!setType) return toast('请选 set 类型');
    setSaving(true);
    try {
      const res = (await api.saveOrder({
        date,
        region: sel.region,
        state: sel.state,
        branch: sel.branch,
        salesman: sel.name,
        item: itemKind(setType),
        setDetail: detail,
        price,
        setQty,
        status,
        paymentDate: payDate,
      })) as { id: number; total: number };
      toast(`✓ 已保存 #${res.id} · RM${res.total}`);
    } catch (e) {
      toast('出错: ' + (e instanceof Error ? e.message : String(e)));
    } finally {
      setSaving(false);
    }
  }

  function handleWa() {
    if (!sel) return toast('请先选 salesman');
    const msg =
      `🆕 新订单 New Order\n` +
      `日期: ${date}\n` +
      `Salesman: ${sel.name}\n` +
      `Branch: ${sel.branch}\n` +
      `State: ${sel.state}\n` +
      `Detail: ${detail || setType}\n` +
      `SET: ${setQty}\n` +
      `Total: RM${fmt(calc.total)}`;
    setWaMsg(msg);
    const phone = data?.driverPhone || '';
    window.open(`https://wa.me/${phone}?text=${encodeURIComponent(msg)}`, '_blank');
  }

  async function submitNewSalesman() {
    if (!newBranch || !newName.trim()) return toast('请选 branch 并输入名字');
    const state = data?.salesmen.find((s) => s.region === newRegion)?.state || '';
    await api.addSalesman(newRegion, state, newBranch, newName.trim());
    await reload();
    toast(`✓ 已新增 ${newName.trim()}`);
    setAddingNew(false);
    const s: Salesman = { name: newName.trim(), branch: newBranch, state, region: newRegion };
    setNewName('');
    pick(s);
  }

  if (!data) return <div className="card">加载中…</div>;

  return (
    <>
      <div className="card">
        <h2>
          ① 选 Salesman <span className="tag">{data.salesmen.length} 人</span>
        </h2>
        <div className="combo" ref={comboRef}>
          {!open && (
            <div className="picked" onClick={() => setOpen(true)}>
              <div className="avatar">{sel ? initials(sel.name) : '?'}</div>
              <div>
                <div className="nm">{sel ? sel.name : '点这里输入名字…'}</div>
                <div className="meta">{sel ? sel.branch : '打字即可搜索，例如 sya'}</div>
              </div>
              <div className="chev">▾</div>
            </div>
          )}
          {open && (
            <div className="dd">
              <div className="search">
                <input
                  autoFocus
                  placeholder="输入名字搜索…"
                  value={query}
                  onChange={(e) => setQuery(e.target.value)}
                />
              </div>
              <div className="list">
                {filtered.length === 0 && <div className="opt empty">没有符合的名字</div>}
                {filtered.map((s, i) => (
                  <div className="opt" key={s.name + s.branch + i} onClick={() => pick(s)}>
                    <div className="avatar" style={{ width: 32, height: 32, fontSize: 13 }}>
                      {initials(s.name)}
                    </div>
                    <div>
                      <div className="nm">{s.name}</div>
                      <div className="meta">
                        {s.branch} · {data.regionLabels[s.region] || s.region}
                      </div>
                    </div>
                  </div>
                ))}
                <div
                  className="opt add"
                  onClick={() => {
                    setAddingNew(true);
                    setOpen(false);
                  }}
                >
                  ＋ 找不到？新增一个 salesman
                </div>
              </div>
            </div>
          )}
        </div>

        {sel && (
          <div className="badges">
            <span className="badge brand">{brandOf(sel.branch)}</span>
            <span className="badge">{sel.branch}</span>
            <span className="badge amber">
              {data.regionLabels[sel.region] || sel.region} · {sel.state}
            </span>
          </div>
        )}

        {addingNew && (
          <div style={{ marginTop: 14, paddingTop: 14, borderTop: '1px dashed var(--line)' }}>
            <label>地区</label>
            <select value={newRegion} onChange={(e) => setNewRegion(e.target.value)}>
              {Object.entries(data.regionLabels).map(([k, v]) => (
                <option key={k} value={k}>
                  {v}
                </option>
              ))}
            </select>
            <label>Branch</label>
            <select value={newBranch} onChange={(e) => setNewBranch(e.target.value)}>
              <option value="">选择 branch…</option>
              {(data.branchesByRegion[newRegion] || []).slice().sort().map((b) => (
                <option key={b} value={b}>
                  {b}
                </option>
              ))}
            </select>
            <label>名字</label>
            <input placeholder="例如 SYAWAN" value={newName} onChange={(e) => setNewName(e.target.value)} />
            <div className="row">
              <button className="btn btn-ghost" onClick={() => setAddingNew(false)}>
                取消
              </button>
              <button className="btn btn-primary" onClick={submitNewSalesman}>
                保存并选用
              </button>
            </div>
          </div>
        )}
      </div>

      {sel && (
        <div className="card">
          <h2>② Set 类型 &amp; 数量</h2>
          <div className="seg">
            {SET_TYPES.map((t) => (
              <div key={t} className={`s ${setType === t ? 'on' : ''}`} onClick={() => applySetType(t)}>
                {t}
              </div>
            ))}
          </div>
          <div className="row" style={{ marginTop: 12 }}>
            <div>
              <label>单价 Price (RM) · 可改</label>
              <input type="number" inputMode="decimal" value={price} onChange={(e) => setPrice(e.target.value)} />
            </div>
            <div>
              <label>数量 SET</label>
              <input type="number" inputMode="numeric" value={setQty} onChange={(e) => setSetQty(e.target.value)} />
            </div>
          </div>
          <label>Set &amp; detail 备注 (可留空)</label>
          <input placeholder="例如 9 items / special item-9" value={detail} onChange={(e) => setDetail(e.target.value)} />

          <div className="calc">
            <div className="c t">
              <div className="k">Total income</div>
              <div className="v">
                <span className="rm">RM</span> {fmt(calc.total)}
              </div>
            </div>
            <div className="c m">
              <div className="k">司机抽佣</div>
              <div className="v">
                <span className="rm">RM</span> {fmt(calc.commission)}
              </div>
            </div>
          </div>
        </div>
      )}

      {sel && (
        <div className="card">
          <h2>③ 保存 &amp; 通知</h2>
          <div className="row">
            <div>
              <label>日期</label>
              <input type="date" value={date} onChange={(e) => setDate(e.target.value)} />
            </div>
            <div>
              <label>状态</label>
              <select value={status} onChange={(e) => setStatus(e.target.value as 'OP' | 'PC')}>
                <option value="OP">OP</option>
                <option value="PC">PC</option>
              </select>
            </div>
          </div>
          <label>Payment date (可留空)</label>
          <input type="date" value={payDate} onChange={(e) => setPayDate(e.target.value)} />

          <button className="btn btn-primary" disabled={saving} onClick={handleSave}>
            {saving ? '保存中…' : '保存订单'}
          </button>
          <button className="btn btn-wa" onClick={handleWa}>
            通知司机 (WhatsApp)
          </button>
          {waMsg && <div className="wa-preview">{waMsg}</div>}
        </div>
      )}
    </>
  );
}
