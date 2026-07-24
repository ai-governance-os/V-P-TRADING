import { useEffect, useState } from 'react';
import { api } from '../api';
import { useAppData } from '../context/AppDataContext';

export default function Manage({ toast }: { toast: (msg: string) => void }) {
  const { data, reload } = useAppData();
  const [region, setRegion] = useState('');
  const [branch, setBranch] = useState('');
  const [name, setName] = useState('');
  const [bRegion, setBRegion] = useState('');
  const [bName, setBName] = useState('');
  const [phone, setPhone] = useState('');
  const [rateVals, setRateVals] = useState<Record<number, string>>({});

  useEffect(() => {
    if (!data) return;
    const first = Object.keys(data.regionLabels)[0] || '';
    if (!region) setRegion(first);
    if (!bRegion) setBRegion(first);
    setPhone(data.driverPhone || '');
  }, [data]); // eslint-disable-line react-hooks/exhaustive-deps

  if (!data) return <div className="card">加载中…</div>;

  const branches = (data.branchesByRegion[region] || []).slice().sort();

  async function addSalesman() {
    if (!branch || !name.trim()) return toast('请选 branch 并输入名字');
    const state = data!.salesmen.find((s) => s.region === region)?.state || '';
    await api.addSalesman(region, state, branch, name.trim());
    await reload();
    setName('');
    toast(`✓ 已新增 ${name.trim()}`);
  }

  async function addBranch() {
    if (!bName.trim()) return toast('请输入 branch 名称');
    const state = data!.salesmen.find((s) => s.region === bRegion)?.state || '';
    await api.addBranch(bRegion, state, bName.trim());
    await reload();
    setBName('');
    toast('✓ 已新增 branch');
  }

  async function saveRate(i: number) {
    const r = data!.rates[i];
    const raw = Number(rateVals[i] ?? (r.Mode === 'percent' ? Number(r.Value) * 100 : r.Value));
    const value = r.Mode === 'percent' ? raw / 100 : raw;
    await api.updateRate(r.Region, r.Item, r.Mode, value);
    await reload();
    toast('规则已更新');
  }

  async function savePhone() {
    await api.setDriverPhone(phone.trim());
    toast('号码已保存');
  }

  return (
    <>
      <div className="card">
        <h2>新增 Salesman</h2>
        <label>地区</label>
        <select
          value={region}
          onChange={(e) => {
            setRegion(e.target.value);
            setBranch('');
          }}
        >
          {Object.entries(data.regionLabels).map(([k, v]) => (
            <option key={k} value={k}>
              {v}
            </option>
          ))}
        </select>
        <label>Branch</label>
        <select value={branch} onChange={(e) => setBranch(e.target.value)}>
          <option value="">选择 branch…</option>
          {branches.map((b) => (
            <option key={b} value={b}>
              {b}
            </option>
          ))}
        </select>
        <label>名字</label>
        <input placeholder="例如 SYAWAN" value={name} onChange={(e) => setName(e.target.value)} />
        <button className="btn btn-primary" onClick={addSalesman}>
          新增 Salesman
        </button>
        <div className="hint">Branch 里找不到？先到下面「新增 Branch」。</div>
      </div>

      <div className="card">
        <h2>新增 Branch</h2>
        <label>地区</label>
        <select value={bRegion} onChange={(e) => setBRegion(e.target.value)}>
          {Object.entries(data.regionLabels).map(([k, v]) => (
            <option key={k} value={k}>
              {v}
            </option>
          ))}
        </select>
        <label>Branch 名称</label>
        <input placeholder="例如 Perodua Bangi" value={bName} onChange={(e) => setBName(e.target.value)} />
        <button className="btn btn-ghost" onClick={addBranch}>
          新增 Branch
        </button>
      </div>

      <div className="card">
        <h2>抽佣规则</h2>
        {data.rates.map((r, i) => {
          const pct = r.Mode === 'percent';
          const label = `${data.regionLabels[r.Region] || r.Region} / ${r.Item}${pct ? ' (百分比 %)' : ' (RM/set)'}`;
          const defaultVal = pct ? Number(r.Value) * 100 : r.Value;
          return (
            <div className="rate-row" key={i}>
              <label>{label}</label>
              <div className="row">
                <input
                  type="number"
                  step="0.01"
                  defaultValue={defaultVal}
                  onChange={(e) => setRateVals((v) => ({ ...v, [i]: e.target.value }))}
                />
                <button className="btn btn-ghost btn-sm" style={{ flex: '0 0 auto' }} onClick={() => saveRate(i)}>
                  保存
                </button>
              </div>
            </div>
          );
        })}
        <div className="hint">NS 与 Johor 部署前请与客户再核对一次。</div>
      </div>

      <div className="card">
        <h2>司机 WhatsApp 号码</h2>
        <label>格式: 国码+号码，不加符号 (例如 60123456789)</label>
        <input inputMode="numeric" value={phone} onChange={(e) => setPhone(e.target.value)} />
        <button className="btn btn-ghost" onClick={savePhone}>
          保存号码
        </button>
      </div>
    </>
  );
}
