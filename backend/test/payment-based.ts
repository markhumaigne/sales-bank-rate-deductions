import assert from 'node:assert/strict';
import Decimal from 'decimal.js';
import ExcelJS from 'exceljs';
import request from 'supertest';
import {
  Configuration,
  Sale,
  ITEM_COLUMNS,
  BRANCH_DATE_SUMMARY_COLUMNS,
  showInItemBankDeductions,
} from '../../shared/models';
import { CalculationService } from '../src/calculation/calculation.service';
import { parsePayments } from '../src/calculation/payment-parser';
import { matchRate } from '../src/calculation/rate-matcher';
import { ConfigService } from '../src/config/config.service';
import { ExcelService } from '../src/excel/excel.service';
import { ExcelExportService } from '../src/excel/excel-export.service';
import { createApplication } from '../src/create-app';
import { sale, part, workbook, upload } from './helpers';

export const paymentTests: { name: string; run: () => unknown | Promise<unknown> }[] = [];
const test = (name: string, run: () => unknown | Promise<unknown>) =>
  paymentTests.push({ name, run });
const configService = new ConfigService(),
  calculator = new CalculationService(),
  excel = new ExcelService();
const exporter = new ExcelExportService(configService, excel);
const settings = (): Configuration => ({
  banks: ['BDO', 'METROBANK', 'BPI'],
  branches: ['ASUS', 'OTHER'],
  paymentMethods: [
    { bank: 'BDO', cardName: 'BDO DEBIT', rate: '2.35' },
    { bank: 'BDO', cardName: 'BDO STRAIGHT', rate: '2.85' },
    { bank: 'METROBANK', cardName: 'METROBANK STRAIGHT', rate: '3.50' },
    { bank: 'BPI', cardName: 'BPI STRAIGHT', rate: '3.00' },
  ],
  aliases: [],
  patterns: [],
  nonCardMethods: ['CASH'],
});
const row = (payment: string, lineTotal: string, values: Partial<Sale> = {}) =>
  sale({
    payment,
    lineTotal,
    price: lineTotal,
    cost: '0',
    payments: [],
    paymentScope: 'invoice',
    ...values,
  });
const calculate = (rows: Sale[], config = settings()) => calculator.calculate(rows, config);
const rates = (item: ReturnType<typeof calculate>['items'][number]) => [
  item.appliedRate1,
  item.appliedRate2,
  item.appliedRate3,
];
const total = (values: (string | null)[]) =>
  values.reduce<Decimal>((sum, value) => sum.plus(value ?? '0'), new Decimal(0)).toFixed(2);

const inlineMetroPayment =
  'CC#2700 METROBANK 12 MOS SAMPLE CUSTOMER 39,999.00 BDO STRAIGHT 10,000.00';

test('Same-line Metrobank and BDO payments use their actual amounts instead of the combined Tender amount', () => {
  const parsed = parsePayments(inlineMetroPayment);
  assert.deepEqual(parsed.issues, []);
  assert.deepEqual(
    parsed.parts.map((p) => [p.method, p.amount, p.reference]),
    [
      ['METROBANK 12 MOS SAMPLE CUSTOMER', '39,999.00', 'CC#2700'],
      ['BDO STRAIGHT', '10,000.00', ''],
    ],
  );
  const sales = [
    row(inlineMetroPayment, '49999', {
      cost: '39291',
      branch: 'ITECH FILINVEST',
      payments: [part('METROBANK 12 MOS SAMPLE CUSTOMER', '49999')],
    }),
  ];
  const original = structuredClone(sales);
  const result = calculate(sales, configService.defaults()),
    item = result.items[0];
  assert.deepEqual(rates(item), ['10.50', '2.85', null]);
  assert.equal(item.adjustment, '4484.90');
  assert.equal(item.grossMargin, '6223.10');
  assert.equal(item.originalPaymentMethod, inlineMetroPayment);
  assert.match(
    item.explanation,
    /METROBANK 12 MOS: actual payment ₱39999.00; Configured Rate: 10.50%; Bank Charge \/ Subsidy: ₱4199.90/,
  );
  assert.match(
    item.explanation,
    /BDO STRAIGHT: actual payment ₱10000.00; Configured Rate: 2.85%; Bank Charge \/ Subsidy: ₱285.00/,
  );
  assert.equal(result.reviews.length, 0);
  assert.equal(result.branchDateSummary[0].adjustment, '4484.90');
  assert.deepEqual(sales, original);
});

test('Two, three and four same-line payments calculate every configured portion in order', () => {
  const methods = ['BDO DEBIT', 'BDO STRAIGHT', 'METROBANK STRAIGHT', 'BPI STRAIGHT'];
  const amounts = [10000, 20000, 15000, 5000];
  for (const count of [2, 3, 4]) {
    for (const decimals of [true, false]) {
      const payment = methods
        .slice(0, count)
        .map((method, i) => `${method} ${decimals ? amounts[i].toFixed(2) : amounts[i]}`)
        .join('  ');
      const result = calculate([
        row(payment, String(amounts.slice(0, count).reduce((a, b) => a + b))),
      ]);
      assert.equal(result.items[0].adjustment, ['805.00', '1330.00', '1480.00'][count - 2]);
      assert.deepEqual(rates(result.items[0]), ['2.35', '2.85', count >= 3 ? '3.50' : null]);
      assert.equal(result.reviews.length, 0);
      if (count === 4)
        assert.match(result.items[0].explanation, /BPI STRAIGHT: actual payment ₱5000.00/);
    }
  }
});

test('Same-line parsing preserves installment terms and card references with integer amounts', () => {
  const payment = 'CC # 2700 METROBANK 12 MOS SAMPLE CUSTOMER 39999 BDO STRAIGHT 7005 10000';
  const result = calculate([row(payment, '49999')], configService.defaults());
  assert.equal(result.items[0].adjustment, '4484.90');
  assert.deepEqual(rates(result.items[0]), ['10.50', '2.85', null]);
  assert.equal(result.reviews.length, 0);
  const ref = calculate([row('BDO STRAIGHT 100 CC # 2700 BDO DEBIT 200', '300')]);
  assert.equal(ref.items[0].adjustment, '7.55');
  assert.equal(ref.reviews.length, 0);
});

test('Same-line mixed methods preserve zero charges and Cash/Charge-first review rules', () => {
  for (const payment of [
    'BDO STRAIGHT ₱44,990.00 SBC QRPH PHP 3,000.00',
    'BDO STRAIGHT 44,990.00 CASH 3,000.00',
    'BDO STRAIGHT 44,990.00 Charge 3,000.00',
  ]) {
    const result = calculate([row(payment, '47990')]);
    assert.equal(result.items[0].adjustment, '1282.22');
    assert.deepEqual(rates(result.items[0]), ['2.85', null, null]);
    assert.equal(result.reviews.length, 0);
  }
  for (const method of ['Cash', 'Charge']) {
    const result = calculate([row(`${method} 3000.00 BDO STRAIGHT 44990.00`, '47990')]);
    assert.equal(result.items[0].adjustment, null);
    assert.equal(result.items.filter(showInItemBankDeductions).length, 0);
    assert.match(result.reviews[0].reason, /Payment starts with/);
  }
});

test('Inline and multiline copies of invoice payments deduplicate and allocate fees only once', () => {
  const rows = ['39999', '10000', '0'].map((amount, i) =>
    row(
      i === 1 ? inlineMetroPayment.replace(' BDO STRAIGHT', '\nBDO STRAIGHT') : inlineMetroPayment,
      amount,
      { source: { file: 'sales.xlsx', sheet: 'Sales Report', row: i + 2 } },
    ),
  );
  const result = calculate(rows, configService.defaults());
  assert.deepEqual(
    result.items.map((item) => item.adjustment),
    ['3587.90', '897.00', '0.00'],
  );
  assert.equal(result.summary.charges, '4484.90');
  assert.equal(result.branchDateSummary[0].adjustment, '4484.90');
  assert.equal(result.reviews.length, 0);
});

test('Same-line payments preserve configured subsidies and branch-specific rates', () => {
  const config = configService.defaults();
  config.paymentMethods.find((method) => method.cardName === 'METROBANK 12 MOS')!.branchRates = {
    'ITECH FILINVEST': '-10.50',
  };
  const result = calculate(
    [row(inlineMetroPayment, '49999', { cost: '39291', branch: 'ITECH FILINVEST' })],
    config,
  );
  assert.deepEqual(rates(result.items[0]), ['-10.50', '2.85', null]);
  assert.equal(result.items[0].adjustment, '-3914.90');
  assert.equal(result.items[0].grossMargin, '14622.90');
  assert.equal(result.summary.charges, '285.00');
  assert.equal(result.summary.subsidies, '4199.90');
  assert.equal(result.branchDateSummary[0].adjustment, '-3914.90');
});

test('Cash/Charge at the start of Payment goes only to Needs Review, including later configured banks', async () => {
  const payments = [
    'Cash 100.00',
    'Charge 100.00',
    '  \r\n cAsH Payment 25.00\r\nBDO STRAIGHT 75.00',
    '\tcharge 25.00;BDO DEBIT 25.00;METROBANK STRAIGHT 50.00',
    'Cash',
    'Charge: missing amount',
  ];
  const rows = payments.map((payment, i) =>
    row(payment, '100', {
      invoice: 'REVIEW-' + i,
      original: { payment },
      source: { file: 'sales.xlsx', sheet: 'Sales Report', row: i + 2 },
    }),
  );
  const snapshot = structuredClone(rows),
    result = calculate(rows);
  assert.equal(result.reviews.length, payments.length);
  assert.equal(result.items.filter(showInItemBankDeductions).length, 0);
  assert.equal(result.summary.charges, '0.00');
  assert.equal(result.summary.sales, '0.00');
  assert.equal(result.summary.feeCalculatedItems, 0);
  assert.deepEqual(result.branchDateSummary, []);
  result.items.forEach((item, i) => {
    assert.equal(item.status, 'Needs Review');
    assert.equal(item.adjustment, null);
    assert.equal(item.grossMargin, null);
    assert.deepEqual(rates(item), [null, null, null]);
    assert.match(result.reviews[i].reason, /Payment starts with (?:CASH|CHARGE)/);
    assert.equal(result.reviews[i].original.payment, payments[i]);
  });
  assert.deepEqual(rows, snapshot);
  const book = new ExcelJS.Workbook();
  await book.xlsx.load((await exporter.report(result)) as any);
  assert.equal(book.worksheets[0].columnCount, 19);
  assert.equal(book.worksheets[0].rowCount, 1);
  assert.equal(book.getWorksheet('Needs Review')!.rowCount, payments.length + 1);
  assert.match(
    String(book.getWorksheet('Needs Review')!.getCell('H4').value),
    /Payment starts with CASH/,
  );
});

test('Bank-first payments keep configured portions when Cash or Charge appears later', () => {
  for (const payment of [
    'BDO STRAIGHT 100.00\nCash 25.00',
    'BDO STRAIGHT 100.00\nCharge 25.00',
    'BDO STRAIGHT CHARGE CUSTOMER 100.00\nSBC QRPH 25.00',
  ]) {
    const result = calculate([row(payment, '125')]);
    assert.equal(result.items.filter(showInItemBankDeductions).length, 1);
    assert.equal(result.items[0].adjustment, '2.85');
    assert.equal(result.reviews.length, 0);
    assert.equal(result.branchDateSummary[0].adjustment, '2.85');
  }
});

test('Unconfigured-only methods move to review and a configured match restores Sheet 1 eligibility', () => {
  const rows = [row('SBC QRPH 100.00', '100')];
  const unmatched = calculate(rows);
  assert.equal(unmatched.items.filter(showInItemBankDeductions).length, 0);
  assert.match(
    unmatched.reviews[0].reason,
    /No payment method matches a bank rate in Configuration Settings/,
  );
  const configured = settings();
  configured.paymentMethods.push({ bank: 'SBC', cardName: 'SBC QRPH', rate: '1.00' });
  const matched = calculate(rows, configured);
  assert.equal(matched.items.filter(showInItemBankDeductions).length, 1);
  assert.equal(matched.items[0].adjustment, '1.00');
  assert.equal(matched.reviews.length, 0);
});

test('Cash-first shared invoices are held together while independent product payments and other invoices calculate', () => {
  const bankFirst = 'BDO STRAIGHT 100.00\nCash 100.00',
    cashFirst = 'Cash 100.00\nBDO STRAIGHT 100.00';
  const shared = calculate([
    row(bankFirst, '100'),
    row(cashFirst, '100', { source: { file: 'sales.xlsx', sheet: 'Sales Report', row: 3 } }),
    row('BDO STRAIGHT 200.00', '200', { invoice: 'OTHER' }),
  ]);
  assert.equal(shared.reviews.length, 2);
  assert.equal(shared.items.filter(showInItemBankDeductions).length, 1);
  assert.deepEqual(
    shared.items.map((item) => item.adjustment),
    [null, null, '5.70'],
  );
  assert.match(
    shared.reviews[0].reason,
    /All products sharing this invoice payment are held together/,
  );
  assert.equal(shared.summary.charges, '5.70');
  assert.equal(shared.branchDateSummary[0].adjustment, '5.70');
  const independent = calculate([
    row(bankFirst, '200', { paymentScope: 'line' }),
    row(cashFirst, '200', { paymentScope: 'line' }),
  ]);
  assert.deepEqual(
    independent.items.map((item) => item.adjustment),
    ['2.85', null],
  );
  assert.equal(independent.reviews.length, 1);
});

for (const [label, payment, paid, expectedRates, fee] of [
  ['one bank', 'BDO STRAIGHT 20,000.00', '20000', ['2.85', null, null], '570.00'],
  [
    'two banks',
    'BDO STRAIGHT 20,000.00\nBDO DEBIT 10,000.00',
    '30000',
    ['2.85', '2.35', null],
    '805.00',
  ],
  [
    'three banks',
    'BDO DEBIT 10,000.00\nBDO STRAIGHT 20,000.00\nMETROBANK STRAIGHT 15,000.00',
    '45000',
    ['2.35', '2.85', '3.50'],
    '1330.00',
  ],
  [
    'configured and unconfigured methods',
    'CC#9259 BDO STRAIGHT NICOLE ALEXANDRA FAJARDO DELA CRUZ 44,990.00\nSBC QRPH 3,000.00',
    '47990',
    ['2.85', null, null],
    '1282.22',
  ],
  [
    'cash and unconfigured QRPH',
    'CASH 10,000.00\nSBC QRPH 3,000.00',
    '13000',
    [null, null, null],
    null,
  ],
] as const)
  test(`Payment amounts: ${label} use independent rates in payment order`, () => {
    const result = calculate([row(payment, paid)]),
      item = result.items[0];
    assert.deepEqual(rates(item), expectedRates);
    assert.equal(item.adjustment, fee);
    assert.equal(item.lineTotal, paid);
    assert.equal(item.originalPaymentMethod, payment);
    assert.equal(item.partialCalculation, false);
    if (expectedRates[0] !== null)
      assert.match(item.explanation, /Actual customer payment amounts, not Line Total/);
    assert.equal(result.reviews.length, expectedRates[0] === null ? 1 : 0);
  });

test('Payment parser handles CRLF, currency, integer amounts, case, spacing and masked card IDs', () => {
  const raw = 'cc # 9259 bdo-straight Nicole  ₱ 44,990.00\r\nsbc qrph PHP 3000';
  const parsed = parsePayments(raw);
  assert.deepEqual(parsed.issues, []);
  assert.equal(parsed.parts[0].reference, 'CC#9259');
  const result = calculate([row(raw, '47990')]);
  assert.equal(result.items[0].adjustment, '1282.22');
  assert.match(
    result.items[0].explanation,
    /actual payment ₱44990.00; Configured Rate: 2.85%; Bank Charge \/ Subsidy: ₱1282.22/,
  );
  assert.match(
    result.items[0].explanation,
    /SBC QRPH: actual payment ₱3000.00; No matching configured rate.*Bank Charge: ₱0.00/i,
  );
  assert.equal(
    calculate([row('****9259 BDO STRAIGHT CUSTOMER 100', '100')]).items[0].adjustment,
    '2.85',
  );
});

test('Invoice payment repeated over three products is charged once with centavo reconciliation', () => {
  const payment = 'BDO STRAIGHT 44,990.00\nSBC QRPH 3,000.00';
  const rows = ['15996.67', '15996.67', '15996.66'].map((amount, i) =>
    row(payment, amount, {
      source: { file: 'sales.xlsx', sheet: 'Sales Report', row: i + 2 },
    }),
  );
  const result = calculate(rows);
  assert.deepEqual(
    result.items.map((item) => item.adjustment),
    ['427.41', '427.41', '427.40'],
  );
  assert.equal(total(result.items.map((item) => item.adjustment)), '1282.22');
  assert.equal(result.summary.charges, '1282.22');
  assert.equal(result.branchDateSummary[0].adjustment, '1282.22');
  result.items.forEach((item) =>
    assert.match(
      item.explanation,
      /deduplicated across 3 product rows.*allocated proportionally by positive Line Total/,
    ),
  );
});

test('Positive products receive proportional fees while zero-Line-Total products receive none', () => {
  const result = calculate(
    ['20000', '10000', '0'].map((amount, i) =>
      row('BDO STRAIGHT 30,000.00', amount, {
        cost: i === 2 ? '238' : '1000',
        source: { file: 'sales.xlsx', sheet: 'Sales Report', row: i + 2 },
      }),
    ),
  );
  assert.deepEqual(
    result.items.map((item) => item.adjustment),
    ['570.00', '285.00', '0.00'],
  );
  assert.equal(result.items[2].grossMargin, '-238.00');
  assert.match(result.items[2].explanation, /Actual customer payment ₱30000.00/);
  assert.match(result.items[2].explanation, /this product's allocated payment share ₱0.00/);
  assert.match(result.items[2].explanation, /Sales Source Row 2 \(Laptop\): ₱570.00/);
  assert.match(result.items[2].explanation, /Sales Source Row 3 \(Laptop\): ₱285.00/);
  assert.equal(result.branchDateSummary[0].adjustment, '855.00');
});

test('Repeated payment on a free accessory points to the product carrying the actual card charge', async () => {
  const payment = 'CC#7500 BDO STRAIGHT SAMPLE CUSTOMER 7,400.00\nCash 70,490.00';
  const rows = [
    { sourceRow: 193, product: 'Free power adapter', lineTotal: '0', cost: '912' },
    { sourceRow: 931, product: 'Phone', lineTotal: '77890', cost: '75585.38' },
    { sourceRow: 1044, product: 'Free bag', lineTotal: '0', cost: '8' },
  ].map(({ sourceRow, product, lineTotal, cost }) =>
    row(payment, lineTotal, {
      product,
      cost,
      source: { file: 'sales.xlsx', sheet: 'Sales Report', row: sourceRow },
    }),
  );
  const result = calculate(rows);
  assert.deepEqual(
    result.items.map((item) => item.adjustment),
    ['0.00', '210.90', '0.00'],
  );
  assert.deepEqual(
    result.items.map((item) => item.grossMargin),
    ['-912.00', '2093.72', '-8.00'],
  );
  assert.equal(result.summary.charges, '210.90');
  assert.equal(result.branchDateSummary[0].adjustment, '210.90');
  for (const item of result.items) {
    assert.deepEqual(rates(item), ['2.85', null, null]);
    assert.match(
      item.explanation,
      /actual payment ₱7400.00; Configured Rate: 2.85%; Bank Charge \/ Subsidy: ₱210.90/,
    );
    assert.match(item.explanation, /Invoice Bank Charges \/ Subsidy ₱210.90 counted once/);
    assert.match(item.explanation, /Sales Source Row 931 \(Phone\): ₱210.90/);
    assert.doesNotMatch(item.explanation, /Eligible payment/);
  }
  assert.match(
    result.items[0].explanation,
    /^This product \(Sales Source Row 193\): allocated Bank Charges \/ Subsidy ₱0.00/,
  );
  assert.match(result.items[0].explanation, /This product has zero Line Total/);
  const book = new ExcelJS.Workbook();
  await book.xlsx.load((await exporter.report(result)) as any);
  const main = book.worksheets[0];
  assert.equal(main.columnCount, 19);
  assert.equal(main.getCell('K2').value, payment);
  assert.equal(main.getCell('L2').value, 0.0285);
  assert.equal(main.getCell('O2').value, 0);
  assert.equal(main.getCell('O3').value, 210.9);
  const note = (cell: ExcelJS.Cell) =>
    typeof cell.note === 'string'
      ? cell.note
      : (cell.note?.texts?.map((text) => text.text).join('') ?? '');
  assert.match(note(main.getCell('O2')), /Sales Source Row 931 \(Phone\): ₱210.90/);
  assert.match(note(main.getCell('L2')), /actual customer payment/);
  assert.equal(main.getCell('S2').value, result.items[0].explanation);
  assert.equal(book.worksheets[1].getCell('E3').value, 210.9);
});

test('Rounding fee allocation uses sales proportions directly and stable source row ties', () => {
  const config = settings();
  config.paymentMethods[0].rate = '50.00';
  const rows = [9, 2, 4].map((sourceRow) =>
    row('BDO DEBIT 0.02\nCASH 0.01', '0.01', {
      source: { file: 'sales.xlsx', sheet: 'Sales Report', row: sourceRow },
    }),
  );
  const result = calculate(rows, config);
  assert.deepEqual(
    result.items.map((item) => item.adjustment),
    ['0.00', '0.01', '0.00'],
  );
  assert.equal(result.summary.charges, '0.01');
});

test('Invoice keys isolate branch, date, file and sheet, while normalized branch names deduplicate', () => {
  const rows = [
    row('BDO STRAIGHT 100', '100'),
    row('BDO STRAIGHT 100', '100', { branch: 'OTHER' }),
    row('BDO STRAIGHT 100', '100', { date: '2026-10-02' }),
    row('BDO STRAIGHT 100', '100', {
      source: { file: 'other.xlsx', sheet: 'Sales Report', row: 2 },
    }),
    row('BDO STRAIGHT 100', '100', { source: { file: 'sales.xlsx', sheet: 'Other', row: 2 } }),
  ];
  assert.equal(calculate(rows).summary.charges, '14.25');
  assert.equal(
    calculate([row('BDO STRAIGHT 100', '60'), row('bdo straight 100', '40', { branch: ' asus ' })])
      .summary.charges,
    '2.85',
  );
});

test('Invoice payment evidence can appear on one product, or as repeated subsets in any order', () => {
  const complete = 'BDO DEBIT 100\nBDO STRAIGHT 200';
  for (const payments of [
    [complete, ''],
    [complete, 'BDO STRAIGHT 200'],
    [complete, 'bdo straight 200\nBDO DEBIT 100'],
    ['BDO DEBIT 100', 'BDO STRAIGHT 200'],
  ]) {
    const result = calculate(
      payments.map((payment, i) =>
        row(payment, '150', { source: { file: 'sales.xlsx', sheet: 'Sales Report', row: i + 2 } }),
      ),
    );
    assert.equal(result.summary.charges, '8.05');
    assert.equal(result.reviews.length, 0);
  }
});

test('Distinct cards and repeated equal payments within one cell remain distinct actual payments', () => {
  for (const payment of [
    'CC#1234 BDO STRAIGHT 100\nCC#5678 BDO STRAIGHT 100',
    'BDO STRAIGHT 100\nBDO STRAIGHT 100',
  ]) {
    const result = calculate([row(payment, '150'), row(payment, '50')]);
    assert.equal(result.summary.charges, '5.70');
    assert.deepEqual(rates(result.items[0]), ['2.85', '2.85', null]);
  }
});

test('A complete invoice payment list determines rate order when the first product only has a subset', () => {
  const result = calculate([
    row('BDO STRAIGHT 200', '150'),
    row('BDO DEBIT 100\nBDO STRAIGHT 200', '150'),
  ]);
  assert.deepEqual(rates(result.items[0]), ['2.35', '2.85', null]);
  assert.equal(result.summary.charges, '8.05');
});

for (const payment of [
  'BDO STRAIGHT',
  'CC#9259 BDO STRAIGHT CUSTOMER',
  'BDO STRAIGHT 1,00.00',
  'BDO STRAIGHT 100.001',
  'BDO STRAIGHT -100.00',
  'BDO STRAIGHT 100.00\nCASH',
  'BDO STRAIGHT 100.00 200.00',
  'BDO STRAIGHT 100.00 BDO DEBIT',
  'BDO STRAIGHT 100.00 BDO DEBIT 1,00.00',
  'BDO STRAIGHT 100.00 BDO DEBIT 100.001',
  'BDO STRAIGHT 100 UNCONFIGURED 200',
  '',
])
  test(`Invalid payment evidence stays unresolved without Line Total fallback: ${JSON.stringify(payment)}`, () => {
    const result = calculate([row(payment, '100')]),
      item = result.items[0];
    assert.equal(item.adjustment, null);
    assert.deepEqual(rates(item), [null, null, null]);
    assert.equal(item.grossMargin, null);
    assert.equal(result.reviews.length, 1);
    assert.match(item.explanation, /missing|invalid|ambiguous/i);
  });

test('Ambiguous bank descriptions never select debit, straight or installment rates by partial names', () => {
  for (const name of [
    'BDO',
    'BDO STRAIGHT DEBIT',
    'BDO STRAIGHT 6 MOS',
    'BDO STRAIGHT 6M',
    'BDO STRAIGHT 6MO',
    'BDO STRAIGHT / BDO DEBIT',
    'SOMETHING BDO STRAIGHT',
  ]) {
    const result = calculate([row(name + ' 100.00', '100')]);
    assert.equal(result.items[0].adjustment, '0.00', name);
    assert.deepEqual(rates(result.items[0]), [null, null, null]);
    assert.match(result.items[0].explanation, /No matching configured rate/);
  }
});

test('Report CREDIT SALES and DEBIT SALES labels do not change a complete payment method', () => {
  for (const payment of [
    'BDO STRAIGHT CREDIT SALES 100',
    'BDO DEBIT DEBIT SALES 100',
    'BDO STRAIGHT SNPL CREDIT SALES 100',
  ]) {
    const result = calculate([row(payment, '100')]);
    assert.equal(result.items[0].adjustment, payment.includes('BDO DEBIT') ? '2.35' : '2.85');
  }
});

test('Broad wildcard settings cannot substitute another bank or payment type', () => {
  const config = settings();
  config.patterns = [{ pattern: 'BDO *', cardName: 'METROBANK STRAIGHT' }];
  const wrongBank = calculate([row('BDO UNKNOWN 100', '100')], config);
  assert.equal(wrongBank.items[0].adjustment, '0.00');
  assert.match(wrongBank.items[0].explanation, /Ambiguous configured pattern/);
  config.paymentMethods = config.paymentMethods.filter((method) => method.cardName !== 'BDO DEBIT');
  config.patterns = [{ pattern: 'BDO *', cardName: 'BDO STRAIGHT' }];
  assert.equal(calculate([row('BDO DEBIT 100', '100')], config).items[0].adjustment, '0.00');
});

test('Branch-specific configured rates override only their matching normalized branch', async () => {
  const config = settings();
  config.paymentMethods[1].branchRates = { ' asus ': '1.50', OTHER: '-2.00' };
  const result = calculate(
    [row('BDO STRAIGHT 100', '100'), row('BDO STRAIGHT 100', '100', { branch: 'OTHER' })],
    configService.validate(config),
  );
  assert.deepEqual(
    result.items.map((item) => item.adjustment),
    ['1.50', '-2.00'],
  );
  assert.deepEqual(
    result.items.map((item) => item.appliedRate1),
    ['1.50', '-2.00'],
  );
  assert.equal(result.items[1].grossMargin, '102.00');
  assert.match(result.items[0].explanation, /Branch-specific rate/);
  assert.equal(matchRate('BDO STRAIGHT', config, false, 'UNLISTED').method!.rate, '2.85');
  assert.deepEqual(exporter.import(upload(await exporter.configuration(config))), config);
  for (const branchRates of [
    { UNKNOWN: '1' },
    { ASUS: '101' },
    { ASUS: '1', ' asus ': '2' },
  ] as Record<string, string>[]) {
    const invalid = settings();
    invalid.paymentMethods[0].branchRates = branchRates;
    assert.throws(() => configService.validate(invalid));
  }
});

test('Four configured payments all contribute fees while only the first three rates are exported', async () => {
  const payment =
    'BDO DEBIT 10000\nBDO STRAIGHT 20000\nMETROBANK STRAIGHT 15000\nBPI STRAIGHT 5000';
  const result = calculate([row(payment, '50000')]);
  assert.deepEqual(rates(result.items[0]), ['2.35', '2.85', '3.50']);
  assert.equal(result.items[0].adjustment, '1480.00');
  assert.match(
    result.items[0].explanation,
    /BPI STRAIGHT: actual payment ₱5000.00; Configured Rate: 3.00%; Bank Charge \/ Subsidy: ₱150.00/,
  );
  const book = new ExcelJS.Workbook();
  await book.xlsx.load((await exporter.report(result)) as any);
  assert.deepEqual(
    (book.worksheets[0].getRow(1).values as unknown[]).slice(1),
    ITEM_COLUMNS.map((column) => column.header),
  );
  assert.deepEqual(
    ['L2', 'M2', 'N2', 'O2'].map((cell) => book.worksheets[0].getCell(cell).value),
    [0.0235, 0.0285, 0.035, 1480],
  );
});

test('Reliable per-product payments preserve distinct product-specific charges', () => {
  const result = calculate([
    row('BDO DEBIT 10000', '10000', { paymentScope: 'line' }),
    row('BDO STRAIGHT 20000', '20000', { paymentScope: 'line' }),
  ]);
  assert.deepEqual(
    result.items.map((item) => item.adjustment),
    ['235.00', '570.00'],
  );
});

test('Conflicting, short or zero-sale invoice payment evidence is flagged with a visible discrepancy', () => {
  for (const rows of [
    [row('BDO STRAIGHT 100', '100'), row('BDO STRAIGHT 120', '100')],
    [row('BDO STRAIGHT 100', '200')],
    [row('BDO STRAIGHT 100', '0')],
    [row('BDO STRAIGHT 100', '100'), row('SBC QRPH 100', '0')],
  ]) {
    const result = calculate(rows);
    result.items.forEach((item) => {
      assert.equal(item.adjustment, null);
      assert.match(item.explanation, /Invoice payment discrepancy/);
    });
    assert.equal(result.summary.charges, '0.00');
  }
});

test('Explicit Payment amounts take priority over stale Tender amounts without mutating source data', () => {
  const rows = [
    row('BDO STRAIGHT 100\nSBC QRPH 100', '200', { payments: [part('BDO STRAIGHT', '200')] }),
  ];
  const original = structuredClone(rows),
    config = settings(),
    snapshot = structuredClone(config);
  assert.equal(calculate(rows, config).items[0].adjustment, '2.85');
  assert.deepEqual(rows, original);
  assert.deepEqual(config, snapshot);
});

test('HTTP upload, invoice allocation and 19-column export reconcile Sheet 1 and Sheet 2', async () => {
  const payment =
    'CC#9259 BDO STRAIGHT NICOLE ALEXANDRA FAJARDO DELA CRUZ 44,990.00\nSBC QRPH 3,000.00';
  const buffer = workbook({
    Sales: [
      [
        'Product ID',
        'Sales Date',
        'Branch',
        'Invoice Number',
        'Product Name',
        'Quantity',
        'Price',
        'Unit Cost',
        'Discount',
        'Line Total',
        'Payment',
      ],
      ['A', '2026-10-01', 'ASUS', '3007-00013640', 'Macbook', 1, 30000, 20000, 0, 30000, payment],
      ['B', '2026-10-01', 'ASUS', '3007-00013640', 'Accessory', 1, 17990, 10000, 0, 17990, payment],
      ['GIFT', '2026-10-01', 'ASUS', '3007-00013640', 'Gift', 1, 0, 238, 0, 0, payment],
    ],
  });
  const app = await createApplication();
  await app.init();
  try {
    const preview = await request(app.getHttpServer())
      .post('/api/excel/preview')
      .attach('file', buffer, 'sales.xlsx')
      .expect(201);
    const sheet = preview.body.sheets[0];
    const response = await request(app.getHttpServer())
      .post('/api/excel/calculate')
      .field(
        'options',
        JSON.stringify({
          sales: { sheet: 'Sales', headerRow: 1, mapping: sheet.suggestedMapping },
          configuration: settings(),
        }),
      )
      .attach('sales', buffer, 'sales.xlsx')
      .expect(201);
    const result = response.body.result;
    assert.equal(result.summary.charges, '1282.22');
    assert.equal(
      total(result.items.map((item: { adjustment: string }) => item.adjustment)),
      '1282.22',
    );
    assert.equal(result.items[2].adjustment, '0.00');
    const download = await request(app.getHttpServer())
      .post('/api/excel/export')
      .send(response.body)
      .buffer(true)
      .parse((res, callback) => {
        const chunks: Buffer[] = [];
        res.on('data', (chunk) => chunks.push(chunk));
        res.on('end', () => callback(null, Buffer.concat(chunks)));
      })
      .expect(201);
    const book = new ExcelJS.Workbook();
    await book.xlsx.load(download.body);
    const main = book.worksheets[0],
      summary = book.worksheets[1];
    assert.equal(main.columnCount, 19);
    assert.equal(main.rowCount, 4);
    assert.deepEqual(
      (summary.getRow(2).values as unknown[]).slice(1),
      BRANCH_DATE_SUMMARY_COLUMNS.map((column) => column.header),
    );
    assert.equal(main.getCell('K2').value, payment);
    for (const [column, summaryColumn] of [
      [10, 3],
      [8, 4],
      [15, 5],
      [16, 6],
    ]) {
      const sum = [2, 3, 4].reduce(
        (value, index) => value.plus(String(main.getCell(index, column).value)),
        new Decimal(0),
      );
      assert.equal(summary.getCell(3, summaryColumn).value, sum.toNumber());
    }
    assert.equal(summary.getCell('E3').value, 1282.22);
  } finally {
    await app.close();
  }
});
