export interface Salesman {
  name: string;
  branch: string;
  state: string;
  region: string;
}

export interface RateRow {
  Region: string;
  Item: string;
  Mode: 'percent' | 'fixed_per_set';
  Value: number;
  Note?: string;
}

export interface BootstrapData {
  salesmen: Salesman[];
  branchesByRegion: Record<string, string[]>;
  prices: Record<string, Record<string, number>>;
  rates: RateRow[];
  regionLabels: Record<string, string>;
  driverPhone: string;
  userEmail: string;
}

export interface Order {
  id: number;
  date: string;
  region: string;
  state: string;
  branch: string;
  salesman: string;
  item: string;
  setDetail: string;
  price: number;
  setQty: number;
  total: number;
  commission: number;
  status: 'OP' | 'PC';
  paymentDate: string;
  month: string;
}

export interface Dashboard {
  month: string;
  monthTotal: number;
  monthCommission: number;
  monthCount: number;
  allTotal: number;
  allCommission: number;
  allCount: number;
  recent: Order[];
}

export interface MonthlySummary {
  month: string;
  orderCount: number;
  totalIncome: number;
  totalCommission: number;
  byBranch: { branch: string; total: number; commission: number; count: number }[];
}

export type TabKey = 'home' | 'new' | 'records' | 'summary' | 'manage';
