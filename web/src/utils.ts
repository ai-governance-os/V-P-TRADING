export function fmt(n: number | string | undefined | null): string {
  const num = Number(n) || 0;
  return num.toLocaleString('en-MY', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

export function fmtInt(n: number | string | undefined | null): string {
  const num = Number(n) || 0;
  return num.toLocaleString('en-MY');
}

export function initials(name: string): string {
  return (name || '?').trim().charAt(0).toUpperCase();
}

export function brandOf(branch: string): string {
  const w = (branch || '').trim().split(/\s+/)[0];
  return w || '—';
}

export function todayISO(): string {
  const d = new Date();
  const pad = (x: number) => String(x).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}
