import { Configuration, PaymentMethod } from '../../../shared/models';
import { normalize } from '../config/config.service';
type RateMatch = { method?: PaymentMethod; nonCard?: boolean; reason?: string; note?: string };
const recognized = (value: string) => normalize(value).replace(/(\d)\s*(MOS|MONTHS?)\b/g, '$1 MOS');
function glob(value: string, pattern: string): boolean {
  // Linear wildcard matching; user patterns are never executed as regular expressions.
  const p = normalize(pattern);
  let vi = 0,
    pi = 0,
    star = -1,
    retry = 0;
  while (vi < value.length) {
    if (p[pi] === value[vi]) {
      vi++;
      pi++;
    } else if (p[pi] === '*') {
      star = pi++;
      retry = vi;
    } else if (star >= 0) {
      pi = star + 1;
      vi = ++retry;
    } else return false;
  }
  while (p[pi] === '*') pi++;
  return pi === p.length;
}

export function matchRate(raw: string, c: Configuration, description = false): RateMatch {
  const value = normalize(raw);
  const exact = c.paymentMethods.find((m) => normalize(m.cardName) === value);
  if (exact) return { method: exact };
  const alias = c.aliases.find((a) => normalize(a.alias) === value);
  if (alias)
    return {
      method: c.paymentMethods.find((m) => normalize(m.cardName) === normalize(alias.cardName)),
    };
  if (c.nonCardMethods.some((n) => normalize(n) === value)) return { nonCard: true };
  const name = recognized(value);
  const full = c.paymentMethods.filter((m) => recognized(m.cardName) === name);
  if (full.length === 1)
    return { method: full[0], note: 'Recognized spacing in the configured payment name.' };
  if (full.length > 1) return { reason: 'Ambiguous normalized payment name.' };
  const special = /\b(BNPL|SNPL|PROMO|PROMOTIONAL)\b/.test(value);
  const patterns = c.patterns.filter(
    (p) =>
      (!special || /\b(BNPL|SNPL|PROMO|PROMOTIONAL)\b/.test(normalize(p.pattern))) &&
      glob(value, p.pattern),
  );
  const targets = [...new Set(patterns.map((p) => normalize(p.cardName)))];
  if (targets.length === 1)
    return { method: c.paymentMethods.find((m) => normalize(m.cardName) === targets[0]) };
  if (targets.length > 1) return { reason: 'Ambiguous configured payment patterns.' };
  // The actual Sales Report describes each method followed by its cardholder/reference.
  // Prefix recognition is allowed only for these explicit description fields.
  if (description) {
    const clean = recognized(value.replace(/^CC#\S+\s+/i, ''));
    const candidates = [
      ...c.paymentMethods.map((m) => ({ name: m.cardName, method: m })),
      ...c.aliases.map((a) => ({
        name: a.alias,
        method: c.paymentMethods.find((m) => normalize(m.cardName) === normalize(a.cardName))!,
      })),
    ]
      .filter((x) => {
        const prefix = recognized(x.name);
        if (clean !== prefix && !clean.startsWith(prefix + ' ')) return false;
        const suffix = clean.slice(prefix.length).trim();
        if (
          /\b(BNPL|SNPL|PROMO|PROMOTIONAL|SPLIT)\b/.test(suffix) ||
          /\b\d+\s*MOS\b/.test(suffix) ||
          /[+;|/&]/.test(suffix)
        )
          return false;
        return true;
      })
      .sort((a, b) => recognized(b.name).length - recognized(a.name).length);
    if (candidates.length) {
      const longest = recognized(candidates[0].name).length;
      const best = candidates.filter((x) => recognized(x.name).length === longest);
      if (new Set(best.map((x) => normalize(x.method.cardName))).size === 1)
        return {
          method: best[0].method,
          note: 'Recognized configured payment prefix in the Sales Report description; trailing cardholder/reference preserved in original text.',
        };
      return { reason: 'Ambiguous payment description.' };
    }
    if (c.nonCardMethods.some((n) => clean === recognized(n))) return { nonCard: true };
  }
  return {
    reason: special
      ? 'Unconfigured BNPL/SNPL or promotional method; add an exact method, alias, or explicit promotional pattern.'
      : 'Unknown payment method or missing configured bank rate.',
  };
}
