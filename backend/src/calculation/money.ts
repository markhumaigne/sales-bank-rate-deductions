import Decimal from 'decimal.js';
Decimal.set({ precision: 40, rounding: Decimal.ROUND_HALF_UP });
export function money(value: string): Decimal | null {
  let text = value.trim().replace(/^(?:₱|PHP)\s*/i, '');
  if (/^\(.*\)$/.test(text)) text = '-' + text.slice(1, -1);
  if (!/^-?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d{1,2})?$/.test(text)) return null;
  try {
    const d = new Decimal(text.replace(/,/g, ''));
    return d.abs().lte('1000000000000') ? d : null;
  } catch {
    return null;
  }
}
export const fixed = (d: Decimal) => d.toFixed(2, Decimal.ROUND_HALF_UP);
export const sum = (values: Decimal[]) => values.reduce((a, b) => a.plus(b), new Decimal(0));
// Largest remainder in whole centavos; ties use ascending original worksheet row order.
export function allocate(total: Decimal, weights: Decimal[]): Decimal[] {
  const cents = total.times(100).toDecimalPlaces(0, Decimal.ROUND_HALF_UP);
  const denominator = sum(weights);
  if (denominator.lte(0)) throw new Error('Allocation requires a positive invoice total.');
  const shares = weights.map((w, index) => {
    const exact = cents.times(w).div(denominator);
    const floor = exact.floor();
    return { index, floor, remainder: exact.minus(floor) };
  });
  const residual = cents.minus(sum(shares.map((s) => s.floor))).toNumber();
  const order = [...shares].sort(
    (a, b) => b.remainder.comparedTo(a.remainder) || a.index - b.index,
  );
  for (let i = 0; i < residual; i++) order[i].floor = order[i].floor.plus(1);
  return shares.map((s) => s.floor.div(100));
}
