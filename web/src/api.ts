// Apps Script 部署网址（.../exec），从环境变量读取。
// 本地开发时在 web/.env.local 里设置 VITE_API_URL=https://script.google.com/macros/s/xxx/exec
const API_URL = import.meta.env.VITE_API_URL as string;

class ApiError extends Error {}

async function apiGet<T>(fn: string, args: unknown[] = []): Promise<T> {
  if (!API_URL) throw new ApiError('未设置 VITE_API_URL，请检查 Vercel 环境变量');
  const url = `${API_URL}?fn=${encodeURIComponent(fn)}&args=${encodeURIComponent(JSON.stringify(args))}`;
  const res = await fetch(url);
  const data = await res.json();
  if (data && data.error) throw new ApiError(data.error);
  return data as T;
}

async function apiPost<T>(fn: string, args: unknown[] = []): Promise<T> {
  if (!API_URL) throw new ApiError('未设置 VITE_API_URL，请检查 Vercel 环境变量');
  // Content-Type: text/plain 是关键 —— 避免浏览器对 Apps Script 发预检请求(它不处理 OPTIONS)。
  const res = await fetch(API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain' },
    body: JSON.stringify({ fn, args }),
  });
  const data = await res.json();
  if (data && data.error) throw new ApiError(data.error);
  return data as T;
}

export const api = {
  getBootstrapData: () => apiGet('getBootstrapData'),
  getDashboard: () => apiGet('getDashboard'),
  searchOrders: (filters: Record<string, string>) => apiGet('searchOrders', [filters]),
  getMonthlySummary: (month: string) => apiGet('getMonthlySummary', [month]),
  listAvailableMonths: () => apiGet<string[]>('listAvailableMonths'),
  previewAmounts: (region: string, item: string, price: number | string, setQty: number | string) =>
    apiGet<{ total: number; commission: number }>('previewAmounts', [region, item, price, setQty]),

  saveOrder: (order: Record<string, unknown>) => apiPost('saveOrder', [order]),
  addSalesman: (region: string, state: string, branch: string, name: string) =>
    apiPost('addSalesman', [region, state, branch, name]),
  addBranch: (region: string, state: string, branch: string) => apiPost('addBranch', [region, state, branch]),
  setPrice: (branch: string, setType: string, price: number) => apiPost('setPrice', [branch, setType, price]),
  updateRate: (region: string, item: string, mode: string, value: number) =>
    apiPost('updateRate', [region, item, mode, value]),
  setDriverPhone: (phone: string) => apiPost('setDriverPhone', [phone]),
  deleteOrder: (id: number) => apiPost<{ ok: boolean }>('deleteOrder', [id]),
};

export { ApiError };
