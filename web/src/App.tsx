import { useState } from 'react';
import './App.css';
import { AppDataProvider, useAppData } from './context/AppDataContext';
import { useToast } from './hooks/useToast';
import BottomNav from './components/BottomNav';
import Home from './components/Home';
import NewOrder from './components/NewOrder';
import Records from './components/Records';
import Summary from './components/Summary';
import Manage from './components/Manage';
import type { TabKey } from './types';

function Shell() {
  const [tab, setTab] = useState<TabKey>('home');
  const { data, loading, error } = useAppData();
  const { message, show } = useToast();

  return (
    <div className="app">
      <div className="appbar">
        <div className="row-top">
          <div>
            <h1>订单 &amp; 抽佣记录</h1>
            <div className="sub">
              {loading ? '加载中…' : data ? `${data.salesmen.length} 位 salesman` : '—'}
            </div>
          </div>
          <div className="avatar-pill">💼</div>
        </div>
      </div>

      <div className="wrap">
        {error && <div className="error-banner">加载失败: {error}</div>}
        {tab === 'home' && <Home />}
        {tab === 'new' && <NewOrder toast={show} />}
        {tab === 'records' && <Records toast={show} />}
        {tab === 'summary' && <Summary />}
        {tab === 'manage' && <Manage toast={show} />}
      </div>

      <BottomNav active={tab} onChange={setTab} />
      {message && <div className="toast">{message}</div>}
    </div>
  );
}

export default function App() {
  return (
    <AppDataProvider>
      <Shell />
    </AppDataProvider>
  );
}
