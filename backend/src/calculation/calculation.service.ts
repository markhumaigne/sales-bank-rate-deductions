import { Injectable } from '@nestjs/common';
import Decimal from 'decimal.js';
import {
  Calculation,
  Configuration,
  ItemResult,
  PaymentPart,
  Sale,
  BankSummary,
  isBranchDateSummaryEligible,
} from '../../../shared/models';
import { normalize } from '../config/config.service';
import { allocate, fixed, money, sum } from './money';
import { matchRate, startsPaymentMethod, withDefaultAliases } from './rate-matcher';
import { summarizeByBranchAndDate } from './branch-date-summary';
import { parsePayments, paymentDescription } from './payment-parser';

type Prepared = { parts: PaymentPart[]; scope: 'invoice' | 'line'; error?: string };
type Allocation = {
  bank: string;
  cardName: string;
  amount: Decimal;
  originalAmount: Decimal;
  fee: Decimal;
  rate: string | null;
};
function paymentData(s: Sale, configuration: Configuration): Prepared {
  const parsed = parsePayments(s.payment, (description) =>
    startsPaymentMethod(description, configuration),
  );
  // Payment is authoritative when it supplies amounts. Existing explicit amount
  // columns inside the Sales Report remain usable for method-only Payment cells.
  if (s.payments.length && !parsed.parts.some((p) => p.amount) && !/[₱,.]|\bPHP\b/i.test(s.payment))
    return {
      parts: s.payments,
      scope: s.paymentScope,
      error: s.paymentIssues.join(' ') || undefined,
    };
  return {
    parts: parsed.parts,
    scope: s.paymentScope,
    error: [...parsed.issues, ...s.paymentIssues].join(' ') || undefined,
  };
}
// Recognizing a payment name does not require a valid amount or a completed allocation.
// Keep this separate from financial validation so one unresolved method cannot hide another.
function identifiedMethods(s: Sale, prepared: Prepared, c: Configuration) {
  const matches = prepared.parts.flatMap((p) => {
    const method = matchRate(p.method, c, p.description, s.branch).method;
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
function sourceValue(s: Sale, key: 'lineTotal' | 'quantity', label: string): string {
  const value = s.original[key] ?? s[key];
  const location = s.cells[key] ? 'cell ' + s.cells[key] : 'source row ' + s.source.row;
  return label + ' at ' + location + ' is ' + (value.trim() ? JSON.stringify(value) : 'blank');
}
@Injectable()
export class CalculationService {
  calculate(sales: Sale[], configuration: Configuration): Calculation {
    configuration = withDefaultAliases(configuration);
    const allocations = new Map<number, Allocation[]>();
    const prepared = sales.map((sale) => paymentData(sale, configuration));
    // Only the beginning of Payment controls this routing rule. A later cash
    // entry or a word inside a cardholder name must not hide a configured bank.
    const reviewPayment = sales.map((s) =>
      /^\s*(CASH|CHARGE)\b/i.exec(s.original.payment ?? s.payment)?.[1].toUpperCase(),
    );
    const recognized = sales.map((s, index) =>
      identifiedMethods(s, prepared[index], configuration),
    );
    // Use the same parsed values for item validation and invoice allocation.
    const sourceValues = sales.map((s) => ({
      net: money(s.lineTotal),
      quantity:
        /^-?\d+(\.\d+)?$/.test(s.quantity.trim()) && Number.isFinite(Number(s.quantity))
          ? new Decimal(s.quantity.trim())
          : null,
    }));
    const items: ItemResult[] = sales.map((s, index) => {
      const issues = [...s.issues],
        warnings = [...s.warnings];
      const { net, quantity } = sourceValues[index];
      const price = money(s.price),
        discount = money(s.discount),
        cost = money(s.cost);
      let status = 'Calculated';
      if (!net) {
        issues.push(
          'Invalid Sales Amount: ' +
            sourceValue(s, 'lineTotal', 'Net Total') +
            '. Check the mapped Net Total column and source value.',
        );
        status = 'Invalid Sales Amount';
      } else if (net.isNegative()) {
        issues.push(
          'Negative sales / refund: ' +
            sourceValue(s, 'lineTotal', 'Net Total') +
            '. Original negative value retained for review; bank charges and gross margin are not calculated.',
        );
        status = 'Needs Review';
      }
      if (quantity?.isNegative()) {
        issues.push(
          'Negative quantity / return: ' +
            sourceValue(s, 'quantity', 'Quantity') +
            '. Original negative value retained; refund/adjustment requires review.',
        );
        if (status === 'Calculated') status = 'Needs Review';
      } else if (!quantity || quantity.isZero()) {
        issues.push(
          'Invalid Quantity: ' +
            sourceValue(s, 'quantity', 'Quantity') +
            '. A positive quantity is required; check the mapped Quantity column and source value.',
        );
        if (status === 'Calculated') status = 'Needs Review';
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
        appliedRate1: null,
        appliedRate2: null,
        appliedRate3: null,
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
        s.source.file,
        s.source.sheet,
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
    const routePaymentToReview = (indices: number[], triggers: number[]) => {
      indices.forEach((i) => (items[i].reviewOnly = true));
      fail(
        indices,
        'Payment starts with ' +
          triggers
            .map((i) => `${reviewPayment[i]} at Sales Source Row ${sales[i].source.row}`)
            .join('; ') +
          '. Moved to Needs Review under the Cash/Charge-first payment rule. ' +
          (indices.length > 1
            ? 'All products sharing this invoice payment are held together to prevent partial or duplicate deductions. '
            : '') +
          'Bank charges and gross margin are pending review and excluded from Sheet 1 and Sheet 2. ' +
          triggers
            .map((i) => prepared[i].error ?? '')
            .filter(Boolean)
            .join(' '),
      );
    };
    const process = (indices: number[], parts: PaymentPart[], allowCashChange = true) => {
      const badRows = indices.filter(
        (i) =>
          !sourceValues[i].net ||
          sourceValues[i].net!.isNegative() ||
          !sourceValues[i].quantity ||
          sourceValues[i].quantity!.lte(0) ||
          sales[i].issues.length ||
          !sales[i].invoice.trim() ||
          !sales[i].branch.trim() ||
          !sales[i].date,
      );
      if (badRows.length) {
        // Individual rows already describe their own source problems. Explain the
        // shared hold separately, so valid sibling items are not called invalid.
        if (indices.length > 1)
          fail(
            indices,
            'Shared payment allocation for invoice ' +
              JSON.stringify(sales[indices[0]].invoice) +
              ' is on hold: source rows ' +
              badRows.map((i) => sales[i].source.row).join(', ') +
              ' require review. See their warnings in Needs Review. The invoice payment cannot be allocated safely until those rows are resolved.',
          );
        else
          fail(
            indices,
            "Payment allocation is on hold pending review of this item's source values.",
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
      const matches = parts.map((p) =>
        matchRate(p.method, configuration, p.description, sales[indices[0]].branch),
      );
      const net = sum(weights),
        paid = sum(amounts as Decimal[]);
      // Cash tender can include change. Keep every configured bank amount intact
      // and remove only an excess fully covered by explicitly identified cash.
      // Unknown methods and card-only discrepancies still require review.
      let cashChangeExplanation = '';
      if (
        paid.gt(net) &&
        allowCashChange &&
        matches.some((m) => !!m.method) &&
        matches.every((m) => m.method || m.nonCard)
      ) {
        const cashIndices = matches.flatMap((m, index) => (m.cash ? [index] : []));
        const cashAmounts = cashIndices.map((index) => amounts[index]!);
        const cashTender = sum(cashAmounts),
          change = paid.minus(net);
        if (cashTender.gte(change)) {
          const netCash = cashTender.minus(change);
          const cashPayments = allocate(netCash, cashAmounts);
          cashIndices.forEach((index, j) => (amounts[index] = cashPayments[j]));
          cashChangeExplanation =
            'Cash tender ₱' +
            fixed(cashTender) +
            ' includes ₱' +
            fixed(change) +
            ' change; net cash payment ₱' +
            fixed(netCash) +
            '. Configured bank payment amounts are unchanged; cash receives no bank charge.';
        }
      }
      if (!net.eq(sum(amounts as Decimal[])) && matches.some((m) => !!m.method)) {
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
      const paymentFees = parts.map((_, p) =>
        matches[p].method
          ? amounts[p]!.times(matches[p].method!.rate)
              .div(100)
              .toDecimalPlaces(2, Decimal.ROUND_HALF_UP)
          : new Decimal(0),
      );
      const details = parts
        .map(
          (part, p) =>
            `${p + 1}. ${matches[p].method?.cardName ?? part.method}: actual payment ₱${fixed(money(part.amount)!)}; ` +
            (matches[p].method
              ? `Configured Rate: ${matches[p].method!.rate}%; Bank Charge / Subsidy: ₱${fixed(paymentFees[p])}.`
              : `No matching configured rate${matches[p].nonCard ? ' (method with no bank charge)' : ''}; Bank Charge: ₱0.00. ${matches[p].reason ?? ''}`),
        )
        .join('\n');
      indices.forEach((i, rowIndex) => {
        items[i].explanation =
          `Invoice: ${sales[i].invoice}\nCustomer Payment:\n${details}\nTotal Customer Payment: ₱${fixed(paid)}\nTotal Bank Charges / Subsidy: ₱${fixed(sum(paymentFees))}\nCalculation Basis: Actual customer payment amounts, not Line Total. ` +
          cashChangeExplanation;
        if (indices.length > 1)
          items[i].explanation +=
            ` Invoice payment evidence deduplicated across ${indices.length} product rows. Payments and rounded charges allocated proportionally by positive Line Total (row ₱${fixed(weights[rowIndex])} / invoice ₱${fixed(net)}); zero-Line-Total rows receive no payment or charge. Largest remainder centavo reconciliation; ties follow source row order. `;
        matches.forEach((m) => {
          if (m.reason?.startsWith('Ambiguous')) items[i].issues.push(m.reason);
        });
      });
      // No-charge methods support split allocations but do not make a row eligible
      // for bank calculations on their own. A configured card must be present.
      if (!matches.some((m) => !!m.method)) {
        indices.forEach((i) => {
          const item = items[i];
          item.cardName = '';
          item.bank = '';
          item.adjustment = '0.00';
          item.status = item.issues.length ? 'Needs Review' : 'Not Calculated';
          item.explanation +=
            'No configured card matches the payment allocation. Payment method(s) excluded from bank calculations: ' +
            parts.map((p) => p.method).join('; ') +
            '. Bank Charges / Subsidy is ₱0.00. Applied Bank Rate columns and GrossMargin are left blank under the existing no-card margin rule. ' +
            'The original Payment text is retained for reference.';
          if (!net.eq(paid))
            item.warnings.push(
              `Invoice payment discrepancy: actual payments ₱${fixed(paid)} versus net product sales ₱${fixed(net)}. No bank charges apply; verify the payment evidence.`,
            );
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
        const fee = paymentFees[p];
        const fees = net.isZero()
          ? weights.map(() => new Decimal(0))
          : signedAllocation(fee, weights);
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
                '%. Actual customer payment ₱' +
                fixed(amount) +
                "; this product's allocated payment share ₱" +
                fixed(amountsPerItem[j]) +
                '; allocated ' +
                (fee.isNegative() ? 'subsidy' : 'charge') +
                ' ₱' +
                fixed(fees[j]) +
                '.'
              : match.nonCard
                ? (match.cash
                    ? 'Cash payment with no bank charge '
                    : 'Configured method with no bank charge ') +
                  part.method +
                  ": this product's allocated payment share ₱" +
                  fixed(amountsPerItem[j]) +
                  ' retained in payment allocation and excluded from bank charges.'
                : 'Unconfigured payment ' +
                  part.method +
                  ": this product's allocated payment share ₱" +
                  fixed(amountsPerItem[j]) +
                  ' retained in payment allocation; no matching configured rate, so its bank charge is ₱0.00.');
          if (match.note) item.explanation += ' ' + match.note;
          if (match.method && /\b(?:BNPL|SNPL)\b/i.test(part.method))
            item.explanation += ' Original payment/tender: ' + JSON.stringify(part.method) + '.';
        });
      });
      const rowCharges = new Map(
        indices.map((index) => [index, sum(allocations.get(index)!.map((p) => p.fee))]),
      );
      const allocationTrace = indices
        .filter((index) => !rowCharges.get(index)!.isZero())
        .map(
          (index) =>
            `Sales Source Row ${sales[index].source.row} (${sales[index].product}): ₱${fixed(rowCharges.get(index)!)}`,
        )
        .join('; ');
      indices.forEach((index) => {
        const item = items[index],
          a = allocations.get(index)!;
        const adjustment = rowCharges.get(index)!,
          line = money(sales[index].lineTotal)!;
        if (indices.length > 1) {
          // Put the row-versus-invoice distinction first, especially on free items
          // whose repeated Payment cell still displays the customer's full payment.
          item.explanation =
            `This product (Sales Source Row ${item.sourceRow}): allocated Bank Charges / Subsidy ₱${fixed(adjustment)}. ` +
            (line.isZero()
              ? 'This product has zero Line Total, so no invoice payment or charge is allocated to it. '
              : '') +
            `Invoice Bank Charges / Subsidy ₱${fixed(sum(paymentFees))} counted once. ` +
            (allocationTrace
              ? `Charges / subsidies recorded on: ${allocationTrace}. Find these products using the Sales Source Row column. `
              : '') +
            '\n' +
            item.explanation;
        }
        item.adjustment = fixed(adjustment);
        item.bank = [...new Set(a.filter((p) => p.rate !== null).map((p) => p.bank))].join(' / ');
        item.cardName = [...new Set(a.filter((p) => p.rate !== null).map((p) => p.cardName))].join(
          ' / ',
        );
        const ratePairs = a.filter((p) => p.rate !== null);
        [item.appliedRate1, item.appliedRate2, item.appliedRate3] = [0, 1, 2].map(
          (n) => ratePairs[n]?.rate ?? null,
        );
        const distinctRates = [...new Map(ratePairs.map((p) => [p.cardName, p])).values()];
        item.appliedRate =
          distinctRates.length === 1
            ? distinctRates[0].rate
            : distinctRates.length
              ? distinctRates.map((p) => p.cardName + ': ' + p.rate + '%').join('; ')
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
        item.explanation += ` Row allocated Bank Charges / Subsidy: ₱${item.adjustment}. Net Line Total already includes discounts; discount was not deducted again.`;
        if (!item.issues.length)
          item.status = adjustment.isZero() ? 'No Bank Charge' : 'Calculated';
        else if (item.status === 'Calculated') item.status = 'Needs Review';
      });
    };
    for (const indices of groups.values()) {
      indices.sort((a, b) => sales[a].source.row - sales[b].source.row || a - b);
      if (indices.some((i) => prepared[i].scope === 'invoice')) {
        const reviewTriggers = indices.filter((i) => reviewPayment[i]);
        if (reviewTriggers.length) {
          routePaymentToReview(indices, reviewTriggers);
          continue;
        }
        const evidence = indices.filter((i) => sales[i].payment.trim() || sales[i].payments.length);
        if (
          !evidence.length ||
          evidence.some((i) => prepared[i].scope !== 'invoice' || prepared[i].error)
        ) {
          fail(
            indices,
            'Incomplete invoice payment evidence across the Sales Report item rows. ' +
              (evidence.length ? evidence : indices)
                .map((i) => prepared[i].error)
                .filter(Boolean)
                .join(' '),
          );
          continue;
        }
        // Multiset union: a repeated invoice list counts once, while two identical
        // entries within one list still count twice. Distinct card references stay
        // distinct. Subsets and payment details present on only one row are supported.
        const counts = new Map<string, number>(),
          parts: PaymentPart[] = [];
        // Prefer a complete list's payment order when another product only has a subset.
        const orderedEvidence = [...evidence].sort(
          (a, b) =>
            prepared[b].parts.length - prepared[a].parts.length ||
            sales[a].source.row - sales[b].source.row ||
            a - b,
        );
        for (const i of orderedEvidence) {
          const local = new Map<string, number>();
          for (const p of prepared[i].parts) {
            const description = paymentDescription(p.method);
            const match = matchRate(p.method, configuration, p.description, sales[i].branch);
            const key = JSON.stringify([
              normalize(match.method?.cardName ?? description.method),
              money(p.amount)?.toFixed(2) ?? p.amount,
              normalize(p.reference || description.reference),
            ]);
            const count = (local.get(key) ?? 0) + 1;
            local.set(key, count);
            if (count > (counts.get(key) ?? 0)) {
              counts.set(key, count);
              parts.push(p);
            }
          }
        }
        process(
          indices,
          parts,
          evidence.some((i) => prepared[i].parts.length === parts.length),
        );
      } else
        for (const i of indices) {
          if (reviewPayment[i]) routePaymentToReview([i], [i]);
          else if (prepared[i].error)
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
      if (!item.cardName && !item.reviewOnly) {
        item.reviewOnly = true;
        item.issues.push(
          'No payment method matches a bank rate in Configuration Settings. Moved to Needs Review; excluded from Sheet 1 and Sheet 2.',
        );
        if (item.status === 'Calculated' || item.status === 'Not Calculated')
          item.status = 'Needs Review';
      }
      if (item.adjustment === null && recognized[index].length) {
        item.explanation =
          'Recognized Tender: ' +
          recognized[index]
            .map((m) => m.cardName + ' (configured rate ' + m.rate + '%)')
            .join('; ') +
          '. The payment method matches Configuration Settings. Bank charges, applied rate and gross margin are pending source review. ' +
          item.explanation;
      }
      item.explanation = [
        item.explanation,
        'Source Cost ' +
          JSON.stringify(sales[index].original.cost ?? item.cost) +
          (sales[index].costBasis === 'line'
            ? ' is already the total line cost; quantity is not multiplied again.'
            : sales[index].costBasis === 'unit'
              ? ' × quantity ' + JSON.stringify(item.quantity) + '.'
              : ' has an unspecified basis; only quantity 1 can use it unchanged.') +
          (item.totalCost !== null
            ? ' Cost shown on Sheet 1 is total product cost ₱' + item.totalCost + '.'
            : ' Total product cost is unresolved and left blank on Sheet 1.'),
        item.cardName
          ? isBranchDateSummaryEligible(item)
            ? 'Included in Sheet 2: this row contributes its Line Total, Cost, Bank Charges / Subsidy and GrossMargin once.'
            : 'Excluded from Sheet 2: this row requires payment, cost, quantity, date or branch review; none of its amounts contribute to that summary.'
          : '',
        ...item.issues,
        ...item.warnings,
        ...sales[index].notes,
      ]
        .filter(Boolean)
        .join(' ');
      const parts = allocations.get(index) || [];
      if (!item.cardName || item.adjustment === null || item.partialCalculation) return;
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
      if (item.cardName && item.adjustment !== null && !item.partialCalculation) return;
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
        reason: [...i.issues, ...i.warnings].filter(Boolean).join(' '),
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
      branchDateSummary: summarizeByBranchAndDate(items),
      reviews,
      configuration: structuredClone(configuration),
      calculatedAt: new Date().toISOString(),
      reportDates: [...new Set(sales.map((s) => s.date).filter(Boolean))].sort(),
      roundingPolicy:
        'Decimal.js; PHP 2 decimals ROUND_HALF_UP. Actual invoice payments counted once; each payment fee rounded first and allocated by positive Line Total using largest remainder, ties in source row order. Zero Line Totals receive no allocation. Negative rates are subsidies. No amount is inferred from Line Total.',
      summary: {
        sales: fixed(totalSales),
        charges: fixed(charges),
        subsidies: fixed(subsidies),
        grossMargin: fixed(margin),
        processedItems: items.length,
        reviewCount: reviews.length,
        feeCalculatedItems: items.filter((i) => i.cardName && i.adjustment !== null).length,
        marginCalculatedItems: items.filter((i) => i.grossMargin !== null && !i.partialCalculation)
          .length,
        partialCalculationItems: items.filter((i) => i.partialCalculation).length,
        missingProductIds: items.filter((i) => !i.productId.trim()).length,
      },
    };
  }
}
