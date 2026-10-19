import { Injectable } from '@nestjs/common';
import Decimal from 'decimal.js';
import {
  Calculation,
  Configuration,
  ItemResult,
  PaymentPart,
  Sale,
  BankSummary,
} from '../../../shared/models';
import { normalize } from '../config/config.service';
import { allocate, fixed, money, sum } from './money';
import { matchRate } from './rate-matcher';

type Prepared = { parts: PaymentPart[]; scope: 'invoice' | 'line'; error?: string };
type Allocation = {
  bank: string;
  cardName: string;
  amount: Decimal;
  originalAmount: Decimal;
  fee: Decimal;
  rate: string | null;
};
function paymentData(s: Sale, c: Configuration): Prepared {
  if (s.payments.length)
    return {
      parts: s.payments,
      scope: s.paymentScope,
      error: s.paymentIssues.join(' ') || undefined,
    };
  const raw = s.payment;
  if (!raw.trim()) return { parts: [], scope: 'line', error: 'Missing Payment Method.' };
  const direct = matchRate(raw, c);
  if (direct.method || direct.nonCard)
    return {
      parts: [{ method: raw, amount: s.lineTotal, reference: '', description: false }],
      scope: 'line',
    };
  const lines = raw
    .split(/\r?\n/)
    .map((x) => x.trim())
    .filter(Boolean);
  const parts: PaymentPart[] = [];
  for (const line of lines) {
    const match = /^(.*?)\s+((?:PHP\s*|₱\s*)?-?\d[\d,]*\.\d{2}|\(\d[\d,]*\.\d{2}\))$/i.exec(line);
    if (!match) {
      parts.length = 0;
      break;
    }
    parts.push({ method: match[1], amount: match[2], reference: '', description: true });
  }
  if (parts.length) return { parts, scope: s.paymentScope };
  if (/\bSPLIT\b|[+\n;|]/i.test(raw) || /\bCASH\s*[/&]/i.test(raw))
    return {
      parts: [],
      scope: 'line',
      error:
        'Split payment has no reliable per-method amounts. Map payment amounts or correct the Sales Report.',
    };
  return {
    parts: [{ method: raw, amount: s.lineTotal, reference: '', description: true }],
    scope: 'line',
  };
}
// Recognizing a payment name does not require a valid amount or a completed allocation.
// Keep this separate from financial validation so one unresolved method cannot hide another.
function identifiedMethods(s: Sale, prepared: Prepared, c: Configuration) {
  const matches = prepared.parts.flatMap((p) => {
    const method = matchRate(p.method, c, p.description).method;
    return method ? [method] : [];
  });
  const raw = s.original.payment ?? s.payment;
  const whole = matchRate(raw, c).method;
  if (whole) matches.push(whole);
  else {
    // Read complete payment entries, never isolated bank words or substrings in customer names.
    // This also retains names when split amounts are missing or wrap onto another source line.
    for (const entry of raw.split(/\r?\n|\s+[+;|]\s+/)) {
      const method = matchRate(entry.trim(), c, true).method;
      if (method) matches.push(method);
    }
  }
  return [...new Map(matches.map((m) => [normalize(m.cardName), m])).values()];
}
const signedAllocation = (value: Decimal, weights: Decimal[]) =>
  allocate(value.abs(), weights).map((x) => (value.isNegative() ? x.negated() : x));
@Injectable()
export class CalculationService {
  calculate(sales: Sale[], configuration: Configuration): Calculation {
    const allocations = new Map<number, Allocation[]>();
    const prepared = sales.map((s) => paymentData(s, configuration));
    const recognized = sales.map((s, index) =>
      identifiedMethods(s, prepared[index], configuration),
    );
    const items: ItemResult[] = sales.map((s, index) => {
      const issues = [...s.issues],
        warnings = [...s.warnings];
      const net = money(s.lineTotal),
        price = money(s.price),
        discount = money(s.discount),
        cost = money(s.cost);
      let quantity: Decimal | null = null;
      try {
        if (/^\d+(\.\d+)?$/.test(s.quantity.trim())) quantity = new Decimal(s.quantity);
      } catch {}
      let status = 'Calculated';
      if (!net || net.isNegative()) {
        issues.push(
          'Invalid Sales Amount: source net line total is missing, invalid, or a refund.',
        );
        status = 'Invalid Sales Amount';
      }
      if (!quantity || quantity.lte(0)) {
        issues.push('Invalid or nonpositive quantity; refund/adjustment requires review.');
        status = 'Needs Review';
      }
      if (!price || price.isNegative())
        issues.push('Price is missing or invalid; its original value was preserved.');
      if (!discount || discount.isNegative())
        issues.push('Discount is missing or invalid. No zero discount was assumed.');
      if (!s.product.trim()) issues.push('Missing Product Name.');
      if (!s.invoice.trim()) issues.push('Missing invoice reference.');
      if (!s.branch.trim()) issues.push('Missing branch.');
      let totalCost: Decimal | null = null;
      if (!cost || cost.isNegative()) {
        issues.push('Missing Cost or invalid product cost; bank charges can still be calculated.');
        if (status === 'Calculated') status = 'Missing Cost';
      } else if (quantity && quantity.gt(0)) {
        if (s.costBasis === 'line') totalCost = cost;
        else if (s.costBasis === 'unit' || quantity.eq(1))
          totalCost = cost.times(quantity).toDecimalPlaces(2, Decimal.ROUND_HALF_UP);
        else {
          issues.push('Ambiguous Cost basis: choose unit cost or total line cost.');
          status = 'Needs Review';
        }
      }
      return {
        productId: s.productId,
        salesDate: s.date,
        branch: s.branch,
        invoice: s.invoice,
        product: s.product,
        quantity: s.quantity,
        price: s.price,
        cost: s.cost,
        discount: s.discount,
        lineTotal: s.lineTotal,
        bank: [...new Set(recognized[index].map((m) => m.bank))].join(' / '),
        originalPaymentMethod: s.original.payment ?? s.payment,
        appliedRate: null,
        adjustment: null,
        grossMargin: null,
        partialCalculation: false,
        status,
        sourceRow: s.source.row,
        cardName: recognized[index].map((m) => m.cardName).join(' / '),
        explanation: '',
        warnings,
        issues,
        totalCost: totalCost ? fixed(totalCost) : null,
        source: s.source,
      };
    });
    const groups = new Map<string, number[]>();
    sales.forEach((s, i) => {
      const key = JSON.stringify([
        normalize(s.invoice),
        normalize(s.branch),
        s.date,
        s.invoice.trim() ? '' : i,
      ]);
      const g = groups.get(key) || [];
      g.push(i);
      groups.set(key, g);
    });
    const fail = (indices: number[], reason: string, status = 'Needs Review') =>
      indices.forEach((i) => {
        items[i].issues.push(reason);
        if (items[i].status !== 'Invalid Sales Amount') items[i].status = status;
      });
    const process = (indices: number[], parts: PaymentPart[]) => {
      const badRows = indices.filter(
        (i) =>
          !money(sales[i].lineTotal) ||
          money(sales[i].lineTotal)!.isNegative() ||
          !Number.isFinite(Number(sales[i].quantity)) ||
          Number(sales[i].quantity) <= 0 ||
          sales[i].issues.length ||
          !sales[i].invoice.trim(),
      );
      if (badRows.length) {
        fail(
          indices,
          'Payment allocation cannot be verified because source values or dates are invalid (rows ' +
            badRows.map((i) => sales[i].source.row).join(', ') +
            ').',
        );
        return;
      }
      const weights = indices.map((i) => money(sales[i].lineTotal)!);
      const amounts = parts.map((p) => money(p.amount));
      if (!parts.length || amounts.some((a) => !a || a.isNegative())) {
        fail(
          indices,
          'Invalid or missing per-method payment amount; no card portion was invented.',
        );
        return;
      }
      const net = sum(weights),
        paid = sum(amounts as Decimal[]);
      if (!net.eq(paid)) {
        fail(
          indices,
          'Invoice payment discrepancy: explicit payments ' +
            fixed(paid) +
            ' versus net product sales ' +
            fixed(net) +
            ' (difference ' +
            fixed(paid.minus(net)) +
            ').',
        );
        return;
      }
      const matches = parts.map((p) => matchRate(p.method, configuration, p.description));
      const unresolved = matches
        .map((m, p) => (!m.method && !m.nonCard ? p : -1))
        .filter((p) => p >= 0);
      if (unresolved.length) {
        const missing = unresolved.some((p) => !parts[p].method.trim());
        const hasMatchedCard = matches.some((m) => !!m.method);
        fail(
          indices,
          unresolved
            .map(
              (p) =>
                '"' + parts[p].method + '": ' + (matches[p].reason || 'Missing payment method.'),
            )
            .join(' '),
          hasMatchedCard
            ? 'Needs Review'
            : missing
              ? 'Missing Payment Method'
              : 'Missing Bank Rate',
        );
        if (!hasMatchedCard) return;
        indices.forEach((i) => {
          items[i].partialCalculation = true;
          items[i].issues.push(
            'Partial calculation: the applied rate and bank charge cover matched configured card payments only. ' +
              'Charges for unconfigured payment portions are excluded, not assumed to be zero. ' +
              'GrossMargin is provisional after the calculated charges. This row is excluded from confirmed financial totals.',
          );
        });
      }
      // No-charge methods support split allocations but do not make a row eligible
      // for bank calculations on their own. A configured card must be present.
      if (!matches.some((m) => !!m.method)) {
        indices.forEach((i) => {
          const item = items[i];
          item.status = item.issues.length ? 'Needs Review' : 'Not Calculated';
          item.explanation =
            'No configured card matches the payment allocation. Payment method(s) excluded from bank calculations: ' +
            parts.map((p) => p.method).join('; ') +
            '. Applied Bank Rate, Bank Charges / Subsidy and GrossMargin are left blank. ' +
            'The original Payment text is retained for reference.';
        });
        return;
      }
      indices.forEach((i) => allocations.set(i, []));
      parts.forEach((part, p) => {
        const match = matches[p],
          amount = amounts[p]!;
        const amountsPerItem = net.isZero()
          ? weights.map(() => new Decimal(0))
          : allocate(amount, weights);
        const fee = match.method
          ? amount.times(match.method.rate).div(100).toDecimalPlaces(2, Decimal.ROUND_HALF_UP)
          : new Decimal(0);
        const fees = amount.isZero()
          ? weights.map(() => new Decimal(0))
          : signedAllocation(fee, amountsPerItem);
        if (!sum(amountsPerItem).eq(amount) || !sum(fees).eq(fee))
          throw new Error('Payment allocation failed reconciliation.');
        indices.forEach((index, j) => {
          allocations.get(index)!.push({
            bank: match.method?.bank || (match.nonCard ? 'No bank charge' : 'Unconfigured payment'),
            cardName: match.method?.cardName || part.method,
            amount: amountsPerItem[j],
            originalAmount: amount,
            fee: fees[j],
            rate: match.method?.rate || null,
          });
          const item = items[index];
          item.explanation +=
            (item.explanation ? ' ' : '') +
            (match.method
              ? match.method.cardName +
                ' matched to configured rate ' +
                match.method.rate +
                '%. Eligible payment ₱' +
                fixed(amountsPerItem[j]) +
                '; allocated ' +
                (fee.isNegative() ? 'subsidy' : 'charge') +
                ' ₱' +
                fixed(fees[j]) +
                '.'
              : match.nonCard
                ? 'Configured method with no bank charge ' +
                  part.method +
                  ': ₱' +
                  fixed(amountsPerItem[j]) +
                  ' retained in payment allocation and excluded from bank charges.'
                : 'Unconfigured payment ' +
                  part.method +
                  ': ₱' +
                  fixed(amountsPerItem[j]) +
                  ' retained in payment allocation; its unknown charge is excluded from the displayed calculation.');
          if (match.note) item.explanation += ' ' + match.note;
        });
      });
      indices.forEach((index) => {
        const item = items[index],
          a = allocations.get(index)!;
        const adjustment = sum(a.map((p) => p.fee)),
          line = money(sales[index].lineTotal)!;
        item.adjustment = fixed(adjustment);
        item.bank = [...new Set(a.filter((p) => p.rate !== null).map((p) => p.bank))].join(' / ');
        item.cardName = [...new Set(a.filter((p) => p.rate !== null).map((p) => p.cardName))].join(
          ' / ',
        );
        const ratePairs = [
          ...new Map(a.filter((p) => p.rate !== null).map((p) => [p.cardName, p])).values(),
        ];
        item.appliedRate =
          ratePairs.length === 1
            ? ratePairs[0].rate
            : ratePairs.length
              ? ratePairs.map((p) => p.cardName + ': ' + p.rate + '%').join('; ')
              : null;
        if (item.totalCost !== null) {
          item.grossMargin = fixed(line.minus(item.totalCost).minus(adjustment));
          item.explanation +=
            (item.partialCalculation ? ' Provisional Gross Margin = ₱' : ' Gross Margin = ₱') +
            fixed(line) +
            ' − ₱' +
            item.totalCost +
            ' cost − ₱' +
            fixed(adjustment) +
            ' charge/subsidy = ₱' +
            item.grossMargin +
            '.';
        }
        item.explanation +=
          ' Source Cost is ' +
          (sales[index].costBasis === 'line'
            ? 'total line cost'
            : sales[index].costBasis === 'unit'
              ? 'unit cost × quantity'
              : Number(sales[index].quantity) === 1
                ? 'unchanged at quantity 1'
                : 'unresolved') +
          '. Net Line Total already includes discounts; discount was not deducted again.';
        if (!item.issues.length)
          item.status = adjustment.isZero() ? 'No Bank Charge' : 'Calculated';
        else if (item.status === 'Calculated') item.status = 'Needs Review';
      });
    };
    for (const indices of groups.values()) {
      if (indices.some((i) => prepared[i].scope === 'invoice')) {
        if (indices.some((i) => prepared[i].scope !== 'invoice' || prepared[i].error)) {
          fail(indices, 'Incomplete invoice payment evidence across the Sales Report item rows.');
          continue;
        }
        const signatures = new Set(
          indices.map((i) =>
            JSON.stringify(
              prepared[i].parts
                .map((p) => [
                  normalize(p.method),
                  money(p.amount)?.toFixed(2) ?? p.amount,
                  normalize(p.reference),
                ])
                .sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b))),
            ),
          ),
        );
        if (signatures.size !== 1) {
          fail(
            indices,
            'Inconsistent invoice payment amounts repeated across product rows. Correct the Sales Report.',
          );
          continue;
        }
        process(indices, prepared[indices[0]].parts);
      } else
        for (const i of indices) {
          if (prepared[i].error)
            fail(
              [i],
              prepared[i].error!,
              sales[i].payment.trim() ? 'Needs Review' : 'Missing Payment Method',
            );
          else process([i], prepared[i].parts);
        }
    }
    const bankMap = new Map<string, BankSummary & { seen: Set<number>; reviewed: Set<number> }>();
    let totalSales = new Decimal(0),
      charges = new Decimal(0),
      subsidies = new Decimal(0),
      margin = new Decimal(0);
    items.forEach((item, index) => {
      if (item.adjustment === null && recognized[index].length) {
        item.explanation =
          'Recognized Tender: ' +
          recognized[index]
            .map((m) => m.cardName + ' (configured rate ' + m.rate + '%)')
            .join('; ') +
          '. The name is retained from the Sales Report payment details. Bank charges, applied rate and gross margin remain unresolved for the reasons below. ' +
          item.explanation;
      }
      item.explanation = [item.explanation, ...item.issues, ...item.warnings, ...sales[index].notes]
        .filter(Boolean)
        .join(' ');
      const parts = allocations.get(index) || [];
      if (item.adjustment === null || item.partialCalculation) return;
      totalSales = totalSales.plus(money(item.lineTotal)!);
      if (item.grossMargin !== null) margin = margin.plus(item.grossMargin);
      let costWeights = parts.map((p) => p.amount);
      if (sum(costWeights).isZero()) costWeights = parts.map((p) => p.originalAmount);
      if (sum(costWeights).isZero()) costWeights = parts.map(() => new Decimal(1));
      const costs =
        item.totalCost === null ? null : allocate(new Decimal(item.totalCost), costWeights);
      parts.forEach((p, j) => {
        if (p.fee.isNegative()) subsidies = subsidies.plus(p.fee.abs());
        else charges = charges.plus(p.fee);
        const key = JSON.stringify([p.bank, p.cardName, item.branch]);
        const b = bankMap.get(key) || {
          bank: p.bank,
          cardName: p.cardName,
          branch: item.branch,
          sales: '0.00',
          charges: '0.00',
          subsidies: '0.00',
          grossMargin: '0.00',
          processedItems: 0,
          reviewCount: 0,
          seen: new Set<number>(),
          reviewed: new Set<number>(),
        };
        b.sales = fixed(new Decimal(b.sales).plus(p.amount));
        b.charges = fixed(new Decimal(b.charges).plus(p.fee.isPositive() ? p.fee : 0));
        b.subsidies = fixed(new Decimal(b.subsidies).plus(p.fee.isNegative() ? p.fee.abs() : 0));
        if (costs)
          b.grossMargin = fixed(
            new Decimal(b.grossMargin).plus(p.amount.minus(costs[j]).minus(p.fee)),
          );
        b.seen.add(index);
        if (item.issues.length || item.warnings.length) b.reviewed.add(index);
        bankMap.set(key, b);
      });
    });
    // Unresolved and ineligible rows stay in the counts without financial contributions.
    items.forEach((item, index) => {
      if (item.adjustment !== null && !item.partialCalculation) return;
      const needsReview = !!(item.issues.length || item.warnings.length);
      const bank = needsReview ? 'Needs Review' : 'Not calculated';
      const cardName = needsReview ? 'Unresolved' : 'No configured card match';
      const key = JSON.stringify([bank, cardName, item.branch]);
      const b = bankMap.get(key) || {
        bank,
        cardName,
        branch: item.branch,
        sales: '0.00',
        charges: '0.00',
        subsidies: '0.00',
        grossMargin: '0.00',
        processedItems: 0,
        reviewCount: 0,
        seen: new Set<number>(),
        reviewed: new Set<number>(),
      };
      b.seen.add(index);
      if (needsReview) b.reviewed.add(index);
      bankMap.set(key, b);
    });
    const reviews = items
      .map((i, index) => ({
        productId: i.productId,
        salesDate: i.salesDate,
        branch: i.branch,
        invoice: i.invoice,
        product: i.product,
        sourceRow: i.sourceRow,
        status: i.status,
        reason: [...i.issues, ...i.warnings].join(' '),
        source: i.source,
        original: sales[index].original,
      }))
      .filter((_, index) => items[index].issues.length || items[index].warnings.length);
    const bankSummary = [...bankMap.values()].map(({ seen, reviewed, ...b }) => ({
      ...b,
      processedItems: seen.size,
      reviewCount: reviewed.size,
    }));
    if (
      !sum(bankSummary.map((b) => new Decimal(b.sales))).eq(totalSales) ||
      !sum(bankSummary.map((b) => new Decimal(b.charges))).eq(charges) ||
      !sum(bankSummary.map((b) => new Decimal(b.subsidies))).eq(subsidies) ||
      !sum(bankSummary.map((b) => new Decimal(b.grossMargin))).eq(margin)
    )
      throw new Error('Bank summary failed financial reconciliation.');
    return {
      items,
      bankSummary,
      reviews,
      configuration: structuredClone(configuration),
      calculatedAt: new Date().toISOString(),
      reportDates: [...new Set(sales.map((s) => s.date).filter(Boolean))].sort(),
      roundingPolicy:
        'Decimal.js; PHP 2 decimals ROUND_HALF_UP. Explicit invoice payments counted once; each payment fee rounded first and allocated by largest remainder, ties in source row order. Negative rates are subsidies. One output row per original sales item.',
      summary: {
        sales: fixed(totalSales),
        charges: fixed(charges),
        subsidies: fixed(subsidies),
        grossMargin: fixed(margin),
        processedItems: items.length,
        reviewCount: reviews.length,
        feeCalculatedItems: items.filter((i) => i.adjustment !== null).length,
        marginCalculatedItems: items.filter((i) => i.grossMargin !== null && !i.partialCalculation)
          .length,
        partialCalculationItems: items.filter((i) => i.partialCalculation).length,
        missingProductIds: items.filter((i) => !i.productId.trim()).length,
      },
    };
  }
}
