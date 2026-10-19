import * as XLSX from 'xlsx';

// Calendar dates, never locale-dependent Date.parse() or timezone conversion.
// The supplied reports use mm/dd/yyyy. ISO text is also accepted explicitly.
export function reportDate(cell: XLSX.CellObject | undefined, date1904 = false): string | null {
  if (!cell || cell.t === 'e' || (cell.f && cell.v === undefined)) return null;
  let y: number, m: number, d: number;
  if (cell.t === 'n' && Number.isFinite(cell.v)) {
    const parts = XLSX.SSF.parse_date_code(cell.v, { date1904 });
    if (!parts) return null;
    ({ y, m, d } = parts);
  } else {
    const text = String(cell.v ?? '').trim();
    const iso = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text);
    const us = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(text);
    if (iso) [, y, m, d] = iso.map(Number);
    else if (us) {
      m = Number(us[1]);
      d = Number(us[2]);
      y = Number(us[3]);
    } else return null;
  }
  const check = new Date(Date.UTC(y, m - 1, d));
  if (
    y < 1900 ||
    y > 9999 ||
    check.getUTCFullYear() !== y ||
    check.getUTCMonth() !== m - 1 ||
    check.getUTCDate() !== d
  )
    return null;
  return `${y.toString().padStart(4, '0')}-${m.toString().padStart(2, '0')}-${d.toString().padStart(2, '0')}`;
}
