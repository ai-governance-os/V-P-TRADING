import type { Order } from '../types';
import { fmt, initials } from '../utils';

export default function OrderRow({ order, onDelete }: { order: Order; onDelete?: (id: number) => void }) {
  return (
    <div className="ord">
      <div className="avatar" style={{ width: 36, height: 36, fontSize: 13 }}>
        {initials(order.salesman)}
      </div>
      <div>
        <div className="nm">
          {order.salesman}
          <span className={`pill ${order.status}`}>{order.status}</span>
        </div>
        <div className="meta">
          #{order.id} · {order.branch} · {order.date} · {order.setQty} set
        </div>
      </div>
      <div className="amt">
        <div className="t">RM{fmt(order.total)}</div>
        <div className="m">佣 RM{fmt(order.commission)}</div>
      </div>
      {onDelete && (
        <div className="del" title="删除" onClick={() => onDelete(order.id)}>
          🗑️
        </div>
      )}
    </div>
  );
}
