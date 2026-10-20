import Decimal from 'decimal.js';
import { BranchDateSummary, ItemResult, isBranchDateSummaryEligible } from '../../../shared/models';
import { normalize } from '../config/config.service';
import { fixed, money } from './money';

export function summarizeByBranchAndDate(items: ItemResult[]): BranchDateSummary[] {
  const groups = new Map<string, BranchDateSummary>();
  for (const item of items) {
    // Reuse completed product calculations: a recognized name alone is not a valid payment.
    // All four amounts must describe the same set of items; unknown costs are never zero.
    if (!isBranchDateSummaryEligible(item)) continue;
    const key = JSON.stringify([item.salesDate, normalize(item.branch)]);
    const group = groups.get(key) ?? {
      date: item.salesDate,
      branch: item.branch.replace(/\s+/g, ' ').trim(),
      sales: '0.00',
      totalCost: '0.00',
      adjustment: '0.00',
      grossMargin: '0.00',
    };
    // Line Total is the existing net sales basis, including discounts and reconciled splits.
    group.sales = fixed(new Decimal(group.sales).plus(money(item.lineTotal)!));
    group.totalCost = fixed(new Decimal(group.totalCost).plus(item.totalCost));
    group.adjustment = fixed(new Decimal(group.adjustment).plus(item.adjustment));
    group.grossMargin = fixed(new Decimal(group.grossMargin).plus(item.grossMargin));
    groups.set(key, group);
  }
  return [...groups.values()].sort(
    (a, b) => a.date.localeCompare(b.date) || a.branch.localeCompare(b.branch),
  );
}
