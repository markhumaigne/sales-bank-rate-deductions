import { PaymentPart } from '../../../shared/models';
import { money } from './money';

// References are retained for invoice deduplication, but never take part in rate matching.
export function paymentDescription(value: string): { method: string; reference: string } {
  const references: string[] = [];
  const method = value
    .replace(
      /(?:\b(?:CC|CARD)\s*#\s*[\d*xX•-]+|[*xX•]{2,}[\d*xX•-]*\d[\d*xX•-]*)/gi,
      (reference) => {
        references.push(reference.replace(/\s+/g, '').toUpperCase());
        return ' ';
      },
    )
    .replace(/\s+/g, ' ')
    .trim();
  return { method, reference: references.join(' / ') };
}

export function parsePayments(
  raw: string,
  startsMethod?: (description: string) => boolean,
): { parts: PaymentPart[]; issues: string[] } {
  const parts: PaymentPart[] = [],
    issues: string[] = [];
  if (!raw.trim()) return { parts, issues: ['Missing Payment method and actual amount.'] };
  for (const [index, entry] of raw.split(/\r\n|[\r\n;|]/).entries()) {
    const line = entry.trim();
    if (!line) continue;
    const addPart = (text: string, amount: string) => {
      const description = paymentDescription(text.replace(/\s*[:=]\s*$/, ''));
      const value = money(amount);
      if (!description.method || !value || value.isNegative())
        issues.push(
          `Payment entry ${index + 1} has an invalid or missing actual payment amount or method: ${JSON.stringify(text + (amount ? ' ' + amount : ''))}. No payment amount was invented.`,
        );
      parts.push({ ...description, amount, description: true });
    };
    // Each explicit monetary token closes its own description, allowing entries
    // on the same line. Consume whole tokens so malformed currency is flagged,
    // never truncated into a different amount. Keep bare card IDs and installment
    // terms inside descriptions unless a complete next payment method follows.
    const tokens = /(?:^|\s)((?:(?:PHP|₱)\s*)?(?:-?\d[\d,.]*|\(\d[\d,.]*\)))(?=\s|$)/gi;
    let start = 0;
    for (const match of line.matchAll(tokens)) {
      const amount = match[1],
        end = match.index! + match[0].length,
        text = line.slice(start, end - amount.length).trim(),
        remaining = line.slice(end).trim();
      if (/(?:CC|CARD)\s*#\s*$/i.test(text)) continue;
      const explicitMoney = /[.,₱]|\bPHP\b/i.test(amount);
      if (!explicitMoney && remaining) {
        if (/^(?:MOS?|MONTHS?|M)\b/i.test(remaining)) continue;
        if (!startsMethod?.(text)) continue;
        if (!startsMethod(remaining)) {
          if (/^[a-z]/i.test(remaining))
            issues.push(
              `Payment entry ${index + 1} has an ambiguous amount or reference before ${JSON.stringify(remaining)}. Use currency-formatted amounts or separate payment lines. No payment amount was invented.`,
            );
          continue;
        }
      }
      addPart(text, amount);
      start = end;
    }
    if (line.slice(start).trim()) addPart(line.slice(start).trim(), '');
  }
  return { parts, issues };
}
