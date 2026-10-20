import 'reflect-metadata';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import Decimal from 'decimal.js';
import request from 'supertest';
import { createApplication } from '../src/create-app';
import { ConfigService } from '../src/config/config.service';
import { ExcelService } from '../src/excel/excel.service';
import { ExcelExportService } from '../src/excel/excel-export.service';
import { CalculationService } from '../src/calculation/calculation.service';
import { summarizeByBranchAndDate } from '../src/calculation/branch-date-summary';
import { matchRate } from '../src/calculation/rate-matcher';
import { reportDate } from '../src/excel/report-date';
import {
  ITEM_COLUMNS,
  BRANCH_DATE_SUMMARY_COLUMNS,
  BRANCH_DATE_SUMMARY_TITLE,
  BRANCH_DATE_SUMMARY_SHEET,
  Sale,
} from '../../shared/models';
import { sale, part, upload, workbook, reportFixture } from './helpers';
const config = new ConfigService(),
  excel = new ExcelService(),
  calc = new CalculationService(),
  exporter = new ExcelExportService(config, excel);
// Fixed, explicitly supplied settings keep arithmetic regressions independent of
// default-rate revisions. The current defaults are checked separately below.
const c = () => {
  const settings = config.defaults();
  const rates: Record<string, string> = {
    'BDO 3 MOS': '3.50',
    'BDO 6 MOS': '6.00',
    'BPI 3 MOS': '3.50',
    'BPI 12 MOS': '12.50',
    'PAYMAYA DEBIT': '2.40',
  };
  for (const method of settings.paymentMethods) method.rate = rates[method.cardName] ?? method.rate;
  settings.paymentMethods.push({ bank: 'BPI', cardName: 'BPI STRAIGHT', rate: '3.00' });
  settings.aliases = [{ alias: 'BDO PAY', cardName: 'BDO STRAIGHT' }];
  return config.validate(settings);
};
const tests: { name: string; run: () => unknown | Promise<unknown> }[] = [];
const test = (name: string, run: () => unknown | Promise<unknown>) => tests.push({ name, run });
const run = (rows: Sale[]) => calc.calculate(rows, c());
const parse = (buf: Buffer, selection: Record<string, string> = {}) => {
  const f = upload(buf),
    p = excel.preview(f).sheets[0];
  return excel.parse(f, {
    sheet: p.name,
    headerRow: p.headerRow,
    mapping: { ...p.suggestedMapping, ...selection },
  });
};
test('Default settings match all 49 approved rates and calculate/export those rates accurately', async () => {
  const expected = [
    ['BDO 24 MOS', '14.00'],
    ['BDO 12 MOS', '10.50'],
    ['BDO 9 MOS', '9.50'],
    ['BDO 6 MOS', '5.50'],
    ['BDO 3 MOS', '3.25'],
    ['BDO PAY/CASA', '3.00'],
    ['BDO QR', '2.50'],
    ['BDO STRAIGHT', '2.85'],
    ['BDO DEBIT', '2.35'],
    ['BDO DINERS', '3.50'],
    ['BDO UNIONPAY', '3.50'],
    ['BDO AMEX', '3.50'],
    ['BPI 3 MOS', '3.00'],
    ['BPI 6 MOS', '5.50'],
    ['BPI 9 MOS', '9.00'],
    ['BPI 12 MOS', '10.00'],
    ['BPI 24 MOS', '14.00'],
    ['METROBANK STRAIGHT', '3.50'],
    ['METROBANK DEBIT', '3.00'],
    ['METROBANK 3 MOS', '3.00'],
    ['METROBANK 6 MOS', '5.50'],
    ['METROBANK 12 MOS', '10.50'],
    ['METROBANK 24 MOS', '12.50'],
    ['METROBANK 36 MOS', '20.50'],
    ['UNIONBANK 3 MOS', '3.50'],
    ['UNIONBANK 6 MOS', '6.50'],
    ['UNIONBANK 12 MOS', '10.50'],
    ['UNIONBANK 24 MOS', '15.50'],
    ['CITI 3 MOS', '3.50'],
    ['CITI 6 MOS', '6.50'],
    ['CITI 12 MOS', '10.50'],
    ['CITI 24 MOS', '15.50'],
    ['RCBC STRAIGHT', '3.93'],
    ['RCBC 3 MOS', '3.93'],
    ['RCBC 6 MOS', '6.92'],
    ['RCBC 12 MOS', '12.89'],
    ['RCBC 24 MOS', '22.50'],
    ['EW 3 MOS', '3.00'],
    ['EW 6 MOS', '6.00'],
    ['EW 12 MOS', '7.50'],
    ['EW 24 MOS', '7.50'],
    ['EW STRAIGHT', '2.50'],
    ['HSBC 3 MOS', '3.50'],
    ['HSBC 12 MOS', '13.00'],
    ['HSBC 6 MOS', '6.50'],
    ['PAYMAYA CREDIT', '2.89'],
    ['PAYMAYA DEBIT', '2.89'],
    ['PAYMAYA QR', '1.99'],
    ['PAYMAYA AMEX', '3.68'],
  ];
  const settings = config.validate(config.defaults());
  assert.equal(expected.length, 49);
  assert.deepEqual(
    settings.paymentMethods.map(({ cardName, rate }) => [cardName, rate]),
    expected,
  );
  for (const method of settings.paymentMethods)
    assert.equal(method.bank, method.cardName.split(' ')[0]);
  assert.deepEqual(settings.aliases, [{ alias: 'BDO PAY', cardName: 'BDO PAY/CASA' }]);
  assert.deepEqual(settings.nonCardMethods, ['CASH', 'SBC QRPH']);
  const result = calc.calculate(
    expected.map(([payment], index) =>
      sale({
        invoice: 'DEFAULT-' + index,
        payment,
        price: '10000',
        lineTotal: '10000',
        cost: '8000',
      }),
    ),
    settings,
  );
  for (const [index, [cardName, rate]] of expected.entries()) {
    const item = result.items[index],
      charge = new Decimal(rate).times(100);
    assert.equal(item.cardName, cardName);
    assert.equal(item.appliedRate, rate);
    assert.equal(item.adjustment, charge.toFixed(2));
    assert.equal(item.grossMargin, new Decimal(2000).minus(charge).toFixed(2));
  }
  const totalCharges = expected.reduce(
    (sum, [, rate]) => sum.plus(new Decimal(rate).times(100)),
    new Decimal(0),
  );
  assert.equal(result.reviews.length, 0);
  assert.deepEqual(result.branchDateSummary, [
    {
      date: '2026-10-01',
      branch: 'ASUS',
      sales: '490000.00',
      totalCost: '392000.00',
      adjustment: totalCharges.toFixed(2),
      grossMargin: new Decimal(98000).minus(totalCharges).toFixed(2),
    },
  ]);
  const alias = calc.calculate(
    [sale({ payment: 'BDO PAY', price: '10000', lineTotal: '10000', cost: '8000' })],
    settings,
  );
  assert.equal(alias.items[0].cardName, 'BDO PAY/CASA');
  assert.equal(alias.items[0].adjustment, '300.00');
  assert.equal(alias.items[0].grossMargin, '1700.00');
  const buffer = await exporter.configuration(settings),
    book = new ExcelJS.Workbook();
  await book.xlsx.load(buffer as any);
  assert.equal(
    book.worksheets[0].getColumn(1).values.filter((value) => value != null && value !== '').length,
    50,
  );
  assert.deepEqual(
    expected.map((_, i) => [
      book.worksheets[0].getCell(i + 2, 1).value,
      new Decimal(String(book.worksheets[0].getCell(i + 2, 2).value)).times(100).toFixed(2),
    ]),
    expected,
  );
  assert.deepEqual(exporter.import(upload(buffer)), settings);
});

test('Single product BDO 3 MOS fee and peso gross margin', () => {
  const r = run([sale()]);
  assert.equal(r.items[0].adjustment, '1050.00');
  assert.equal(r.items[0].grossMargin, '3950.00');
  assert.equal(r.items[0].productId, 'LAPTOP-001');
});
test('BPI 12 MOS uses 12.50 percent', () => {
  const r = run([sale({ payment: 'BPI 12 MOS' })]);
  assert.equal(r.items[0].adjustment, '3750.00');
  assert.equal(r.items[0].grossMargin, '1250.00');
});
test('Multiple products sharing an invoice remain separate, repeated payments counted once', () => {
  const payments = [part('BDO 3 MOS', '25000')];
  const r = run([
    sale({ lineTotal: '20000', cost: '16000', payments }),
    sale({
      source: { file: 'sales.xlsx', sheet: 'Sales Report', row: 3 },
      productId: 'MOUSE',
      product: 'Mouse',
      lineTotal: '5000',
      cost: '2000',
      payments,
    }),
  ]);
  assert.equal(r.items.length, 2);
  assert.deepEqual(
    r.items.map((i) => i.adjustment),
    ['700.00', '175.00'],
  );
  assert.equal(r.summary.charges, '875.00');
});
test('Cash retains its source row with zero bank deductions and no bank rate or margin', async () => {
  const r = run([sale({ payment: 'Cash' })]);
  const item = r.items[0];
  assert.equal(item.status, 'Not Calculated');
  assert.equal(item.cardName, '');
  assert.equal(item.originalPaymentMethod, 'Cash');
  assert.equal(item.lineTotal, '30000');
  assert.equal(item.appliedRate, null);
  assert.equal(item.adjustment, '0.00');
  assert.equal(item.grossMargin, null);
  assert.equal(r.summary.sales, '0.00');
  assert.equal(r.summary.grossMargin, '0.00');
  assert.equal(r.summary.feeCalculatedItems, 0);
  assert.equal(r.summary.processedItems, 1);
  assert.equal(r.summary.reviewCount, 0);
  assert.equal(r.bankSummary[0].processedItems, 1);
  assert.equal(r.bankSummary[0].reviewCount, 0);
  assert.match(item.explanation, /No configured card matches/);
  const book = new ExcelJS.Workbook();
  await book.xlsx.load((await exporter.report(r)) as any);
  const ws = book.worksheets[0];
  assert.equal(ws.rowCount, 1, 'Cash-only rows are omitted from Item Bank Deductions');
  assert.equal(ws.columnCount, 17);
});

test('Cash-only gifts do not create negative margins and adjacent matched BDO rows still calculate', () => {
  const payment = 'Cash 27,100.00';
  const payments = [part('Cash', '27100')];
  const r = run([
    sale({
      payment,
      payments,
      lineTotal: '0',
      price: '0',
      cost: '1600',
      source: { file: 'sales.xlsx', sheet: 'Sales Report', row: 40 },
    }),
    sale({
      payment,
      payments,
      lineTotal: '27100',
      source: { file: 'sales.xlsx', sheet: 'Sales Report', row: 42 },
    }),
    sale({
      invoice: '0002',
      payment: 'BDO STRAIGHT',
      lineTotal: '33999',
      cost: '30264',
      source: { file: 'sales.xlsx', sheet: 'Sales Report', row: 41 },
    }),
  ]);
  for (const item of r.items.slice(0, 2)) {
    assert.equal(item.cardName, '');
    assert.equal(item.appliedRate, null);
    assert.equal(item.adjustment, '0.00');
    assert.equal(item.grossMargin, null);
    assert.doesNotMatch(item.explanation, /Gross Margin =/);
  }
  assert.equal(r.items[0].sourceRow, 40);
  assert.equal(r.items[2].appliedRate, '2.85');
  assert.equal(r.items[2].adjustment, '968.97');
  assert.equal(r.items[2].grossMargin, '2766.03');
  assert.equal(r.summary.sales, '33999.00');
  assert.equal(r.summary.grossMargin, '2766.03');
  assert.equal(r.summary.processedItems, 3);
  assert.equal(r.summary.reviewCount, 0);
});

test('Line-specific Cash is not calculated just because another product on the invoice has a card', () => {
  const r = run([sale(), sale({ payment: 'Cash' })]);
  assert.equal(r.items[0].adjustment, '1050.00');
  assert.equal(r.items[1].cardName, '');
  assert.equal(r.items[1].appliedRate, null);
  assert.equal(r.items[1].adjustment, '0.00');
  assert.equal(r.items[1].grossMargin, null);
  assert.equal(r.summary.grossMargin, '3950.00');
});

test('An explicitly configured zero-percent card remains eligible for calculation', () => {
  const cfg = c();
  cfg.paymentMethods.find((m) => m.cardName === 'BDO 3 MOS')!.rate = '0.00';
  const item = calc.calculate([sale()], cfg).items[0];
  assert.equal(item.cardName, 'BDO 3 MOS');
  assert.equal(item.appliedRate, '0.00');
  assert.equal(item.adjustment, '0.00');
  assert.equal(item.grossMargin, '5000.00');
});
test('Unknown payment method never fabricates a rate or financial totals', () => {
  const r = run([sale({ payment: 'Unknown' })]);
  assert.equal(r.items[0].status, 'Missing Bank Rate');
  assert.equal(r.items[0].adjustment, null);
  assert.equal(r.summary.sales, '0.00');
  assert.equal(r.reviews.length, 1);
  assert.equal(r.items[0].cardName, '');
});

test('Matched Tender calculates its rate, charge and margin despite another unconfigured payment', async () => {
  const cfg = c();
  cfg.nonCardMethods = cfg.nonCardMethods.filter((name) => name !== 'SBC QRPH');
  const payment = 'CC#0001 BDO STRAIGHT EXAMPLE CUSTOMER 44,990.00\nSBC QRPH 3,000.00';
  const payments = [part('BDO STRAIGHT EXAMPLE CUSTOMER', '44990'), part('SBC QRPH', '3000')];
  const r = calc.calculate(
    [
      sale({
        productId: 'GIFT',
        product: 'Bag',
        lineTotal: '0',
        price: '0',
        cost: '238',
        payment,
        payments,
        source: { file: 'sales.xlsx', sheet: 'Sales Report', row: 3 },
      }),
      sale({
        lineTotal: '47990',
        payment,
        payments,
        source: { file: 'sales.xlsx', sheet: 'Sales Report', row: 8 },
      }),
    ],
    cfg,
  );
  for (const item of r.items) {
    assert.equal(item.cardName, 'BDO STRAIGHT');
    assert.equal(item.originalPaymentMethod, payment);
    assert.equal(item.appliedRate, '2.85');
    assert.equal(item.partialCalculation, true);
    assert.match(item.explanation, /Partial calculation/);
    assert.match(item.explanation, /SBC QRPH/);
  }
  assert.deepEqual(
    r.items.map((i) => i.adjustment),
    ['0.00', '1282.22'],
  );
  assert.deepEqual(
    r.items.map((i) => i.grossMargin),
    ['-238.00', '21707.78'],
  );
  assert.equal(r.summary.charges, '0.00');
  assert.equal(r.summary.sales, '0.00');
  assert.equal(r.summary.grossMargin, '0.00');
  assert.equal(r.summary.partialCalculationItems, 2);
  assert.equal(r.summary.reviewCount, 2);
  const b = new ExcelJS.Workbook();
  await b.xlsx.load((await exporter.report(r)) as any);
  assert.equal(b.worksheets[0].getCell('P2').value, 'BDO STRAIGHT');
  assert.equal(b.worksheets[0].getCell('K2').value, payment);
  assert.equal(b.worksheets[0].getCell('L2').value, 0.0285);
  assert.equal(b.worksheets[0].getCell('M2').value, 0);
  assert.equal(b.worksheets[0].getCell('N2').value, -238);
});

test('SBC QRPH is excluded from charges while matched BDO payments calculate rate, fee and margin', async () => {
  assert.ok(c().nonCardMethods.includes('SBC QRPH'));
  const payment = 'CC#0001 BDO STRAIGHT EXAMPLE CUSTOMER 44,990.00\nSBC QRPH 3,000.00';
  const payments = [part('BDO STRAIGHT EXAMPLE CUSTOMER', '44990'), part('SBC QRPH', '3000')];
  const r = run([
    sale({
      productId: 'GIFT',
      product: 'Bag',
      price: '0',
      lineTotal: '0',
      cost: '238',
      payment,
      payments,
    }),
    sale({
      productId: 'LAPTOP',
      price: '47990',
      lineTotal: '47990',
      cost: '40000',
      payment,
      payments,
      source: { file: 'sales.xlsx', sheet: 'Sales Report', row: 3 },
    }),
  ]);
  assert.deepEqual(
    r.items.map((i) => i.appliedRate),
    ['2.85', '2.85'],
  );
  assert.deepEqual(
    r.items.map((i) => i.adjustment),
    ['0.00', '1282.22'],
  );
  assert.deepEqual(
    r.items.map((i) => i.grossMargin),
    ['-238.00', '6707.78'],
  );
  assert.equal(r.summary.sales, '47990.00');
  assert.equal(r.summary.charges, '1282.22');
  assert.equal(r.summary.grossMargin, '6469.78');
  assert.equal(r.summary.reviewCount, 0);
  const qr = r.bankSummary.find((b) => b.cardName === 'SBC QRPH')!;
  assert.equal(qr.sales, '3000.00');
  assert.equal(qr.charges, '0.00');
  for (const item of r.items) {
    assert.equal(item.originalPaymentMethod, payment);
    assert.equal(item.cardName, 'BDO STRAIGHT');
    assert.match(item.explanation, /SBC QRPH:[\s\S]*excluded from bank charges/);
    assert.doesNotMatch(item.explanation, /unresolved|Unknown payment/);
  }
  const b = new ExcelJS.Workbook();
  await b.xlsx.load((await exporter.report(r)) as any);
  assert.equal(b.worksheets[0].getCell('L2').value, 0.0285);
  assert.equal(b.worksheets[0].getCell('M2').value, 0);
  assert.equal(b.worksheets[0].getCell('N2').value, -238);
  assert.equal(b.worksheets[0].getCell('M3').value, 1282.22);
  assert.deepEqual(
    exporter.import(upload(Buffer.from(await b.xlsx.writeBuffer()))),
    r.configuration,
  );
});

test('SBC QRPH-only items retain their source values but are omitted from Sheet 1', async () => {
  for (const payment of ['SBC QRPH', ' sbc   qrph ', 'SBC QRPH 30,000.00']) {
    const r = run([sale({ payment })]);
    const item = r.items[0];
    assert.equal(item.originalPaymentMethod, payment);
    assert.equal(item.cardName, '');
    assert.equal(item.appliedRate, null);
    assert.equal(item.adjustment, null);
    assert.equal(item.grossMargin, null);
    assert.equal(r.summary.feeCalculatedItems, 0);
    assert.equal(r.summary.marginCalculatedItems, 0);
    assert.equal(r.summary.grossMargin, '0.00');
    assert.match(item.explanation, /excluded from bank calculations:.*sbc\s+qrph/i);
    const book = new ExcelJS.Workbook();
    await book.xlsx.load((await exporter.report(r)) as any);
    const ws = book.worksheets[0];
    assert.equal(ws.rowCount, 1);
    assert.equal(ws.columnCount, 17);
  }
});

test('No-charge SBC QRPH stays exact and never hides missing splits, costs or other unknown methods', () => {
  const qr = run([sale({ payment: ' sbc   qrph ' })]);
  assert.equal(qr.items[0].cardName, '');
  assert.equal(qr.items[0].appliedRate, null);
  assert.equal(qr.items[0].adjustment, null);
  assert.equal(qr.items[0].grossMargin, null);
  const text = run([sale({ payment: 'BDO 3 MOS 20,000.00\nSBC QRPH 10,000.00' })]);
  assert.equal(text.items[0].appliedRate, '3.50');
  assert.equal(text.items[0].adjustment, '700.00');
  assert.equal(text.items[0].grossMargin, '4300.00');
  const missingCost = run([
    sale({ cost: '', payments: [part('BDO 3 MOS', '20000'), part('SBC QRPH', '10000')] }),
  ]);
  assert.equal(missingCost.items[0].adjustment, '700.00');
  assert.equal(missingCost.items[0].grossMargin, null);
  assert.equal(run([sale({ payment: 'BDO 3 MOS + SBC QRPH' })]).items[0].adjustment, null);
  const partial = run([
    sale({
      payments: [part('BDO 3 MOS', '20000'), part('SBC QRPH', '5000'), part('Unknown', '5000')],
    }),
  ]);
  assert.equal(partial.items[0].adjustment, '700.00');
  assert.equal(partial.items[0].grossMargin, '4300.00');
  assert.equal(partial.items[0].partialCalculation, true);
  assert.equal(partial.summary.charges, '0.00');
  assert.equal(run([sale({ payment: 'SBC QRPH PROMO' })]).items[0].adjustment, null);
});

test('Single matched Payment description calculates all three values using its net line amount', () => {
  const r = run([sale({ payment: 'BDO STRAIGHT EXAMPLE CUSTOMER' })]);
  assert.equal(r.items[0].cardName, 'BDO STRAIGHT');
  assert.equal(r.items[0].originalPaymentMethod, 'BDO STRAIGHT EXAMPLE CUSTOMER');
  assert.equal(r.items[0].appliedRate, '2.85');
  assert.equal(r.items[0].adjustment, '855.00');
  assert.equal(r.items[0].grossMargin, '4145.00');
  assert.equal(r.items[0].partialCalculation, false);
});

test('Matched splits include BNPL portions once at the underlying configured rate', () => {
  const r = run([
    sale({ payment: 'BDO 3MOS 10,000.00\nBPI 12 MOS 10,000.00\nBDO 6 MOS BNPL 10,000.00' }),
  ]);
  const item = r.items[0];
  assert.equal(item.adjustment, '2200.00');
  assert.equal(item.grossMargin, '2800.00');
  assert.equal(item.cardName, 'BDO 3 MOS / BPI 12 MOS / BDO 6 MOS');
  assert.match(item.appliedRate!, /BDO 3 MOS: 3.50%/);
  assert.match(item.appliedRate!, /BPI 12 MOS: 12.50%/);
  assert.equal(item.partialCalculation, false);
  assert.equal(r.reviews.length, 0);
  assert.equal(r.summary.grossMargin, '2800.00');
  assert.equal(r.branchDateSummary[0].grossMargin, '2800.00');
  const missingCost = run([
    sale({ cost: '', payments: [part('BDO 3 MOS', '20000'), part('Unknown', '10000')] }),
  ]);
  assert.equal(missingCost.items[0].adjustment, '700.00');
  assert.equal(missingCost.items[0].grossMargin, null);
});

test('Payment text retains known names when split amounts or source details need review', () => {
  for (const payment of [
    'CC#0001 BDO STRAIGHT EXAMPLE CUSTOMER\n44,990.00\nSBC QRPH 3,000.00',
    'BDO STRAIGHT + Unknown Method',
    'BDO STRAIGHT 44,990.00\nUnknown Method 3,000.00',
  ]) {
    const item = run([sale({ payment })]).items[0];
    assert.equal(item.cardName, 'BDO STRAIGHT');
    assert.equal(item.adjustment, null);
  }
  const invalid = run([sale({ payment: 'BPI 12 MOS', lineTotal: 'invalid' })]).items[0];
  assert.equal(invalid.cardName, 'BPI 12 MOS');
  assert.equal(invalid.adjustment, null);
});

test('Tender recognition keeps full names, aliases and all known split methods without guessing', () => {
  const cfg = c();
  cfg.aliases = [{ alias: 'Terminal A', cardName: 'BDO STRAIGHT' }];
  const item = calc.calculate([sale({ payment: 'Terminal A + BPI 12 MOS + Unknown' })], cfg)
    .items[0];
  assert.equal(item.cardName, 'BDO STRAIGHT / BPI 12 MOS');
  assert.equal(item.adjustment, null);
  for (const payment of [
    'BDO',
    'NOTBDO STRAIGHT',
    'BDO STRAIGHTER',
    'Unknown Customer BDO STRAIGHT',
    'BDO 6 MOS PROMO',
  ]) {
    assert.equal(run([sale({ payment })]).items[0].cardName, '');
  }
});
test('Missing cost preserves a valid bank charge and leaves margin blank', () => {
  const r = run([sale({ cost: '' })]);
  assert.equal(r.items[0].status, 'Missing Cost');
  assert.equal(r.items[0].adjustment, '1050.00');
  assert.equal(r.items[0].grossMargin, null);
  assert.equal(r.summary.charges, '1050.00');
  assert.equal(r.summary.grossMargin, '0.00');
});
test('Discount already in net line total is never deducted twice', () => {
  const r = run([sale({ price: '1000', discount: '100', lineTotal: '900', cost: '600' })]);
  assert.equal(r.items[0].adjustment, '31.50');
  assert.equal(r.items[0].grossMargin, '268.50');
});
test('Branches and repeated invoice references remain independent', () => {
  const r = run([
    sale(),
    sale({
      branch: 'SMCEBU',
      productId: 'OTHER',
      source: { file: 'sales.xlsx', sheet: 'Sales Report', row: 3 },
    }),
  ]);
  assert.equal(r.items.length, 2);
  assert.equal(r.bankSummary.length, 2);
  assert.equal(r.summary.charges, '2100.00');
});
test('BNPL/SNPL Tender descriptions match underlying methods and retain original Payment', () => {
  for (const [payment, fee] of [
    ['BDO 6 MOS BNPL', '1800.00'],
    ['METROBANK 6 MOS SNPL', '1650.00'],
  ]) {
    const r = run([sale({ payment, payments: [part(payment + ' CUSTOMER', '30000')] })]);
    assert.equal(r.items[0].adjustment, fee);
    assert.equal(r.items[0].originalPaymentMethod, payment);
    assert.match(r.items[0].explanation, /Ignored BNPL\/SNPL/);
    assert.equal(r.reviews.length, 0);
  }
  const cfg = c();
  cfg.aliases = [{ alias: 'BDO 6 MOS BNPL', cardName: 'BDO 6 MOS' }];
  assert.equal(
    calc.calculate([sale({ payment: 'BDO 6 MOS BNPL' })], cfg).items[0].adjustment,
    '1800.00',
  );
});
test('Split payment without explicit amounts stays unresolved', () => {
  const r = run([sale({ payment: 'Cash + BDO 3 MOS' })]);
  assert.equal(r.items[0].status, 'Needs Review');
  assert.equal(r.items[0].adjustment, null);
  assert.match(r.items[0].explanation, /Split payment/);
});
test('BNPL/SNPL normalization preserves existing rates, costs and margins for all label positions', () => {
  for (const [payment, underlying] of [
    ['BDO Straight BNPL', 'BDO STRAIGHT'],
    ['BPI 3 MOS SNPL', 'BPI 3 MOS'],
    ['BNPL BDO 6 MOS', 'BDO 6 MOS'],
    ['SNPL Metrobank Straight', 'METROBANK STRAIGHT'],
    ['BDO BNPL 6 MOS SNPL', 'BDO 6 MOS'],
    ['  bdo  3mos  snpl  ', 'BDO 3 MOS'],
    ['BNPL SNPL BDO DEBIT BNPL', 'BDO DEBIT'],
    ['BDO Debit', 'BDO DEBIT'],
  ]) {
    const source = sale({
      payment,
      quantity: '2',
      price: '16000',
      cost: '12000',
      discount: '2000',
    });
    const snapshot = structuredClone(source);
    const r = run([source]),
      base = run([sale({ ...source, payment: underlying })]);
    assert.deepEqual(source, snapshot, 'Normalization never changes source values');
    assert.equal(r.items[0].originalPaymentMethod, payment);
    assert.equal(r.items[0].cardName, underlying);
    for (const key of ['appliedRate', 'adjustment', 'grossMargin', 'totalCost'] as const)
      assert.equal(r.items[0][key], base.items[0][key], payment + ': ' + key);
    assert.deepEqual(r.branchDateSummary, base.branchDateSummary);
    assert.equal(r.reviews.length, 0);
    assert.equal(r.summary.feeCalculatedItems, 1);
  }
});
test('Explicit BNPL/SNPL methods, aliases and patterns override fallback without duplicate fees', () => {
  const cfg = c();
  cfg.paymentMethods.push({ bank: 'BDO', cardName: 'BDO 6 MOS BNPL', rate: '-2.00' });
  cfg.aliases.push({ alias: 'SNPL BDO 6 MOS', cardName: 'BDO 3 MOS' });
  cfg.patterns.push({ pattern: 'BDO 6 MOS SNPL *', cardName: 'BDO 6 MOS BNPL' });
  const snapshot = structuredClone(cfg);
  for (const [payment, fee, margin] of [
    ['BDO 6 MOS BNPL', '-600.00', '5600.00'],
    ['BDO 6 MOS BNPL CUSTOMER', '-600.00', '5600.00'],
    ['SNPL BDO 6 MOS', '1050.00', '3950.00'],
    ['BDO 6 MOS SNPL CUSTOMER', '-600.00', '5600.00'],
    ['BDO 6 MOS', '1800.00', '3200.00'],
  ]) {
    const r = calc.calculate([sale({ payment })], config.validate(cfg));
    assert.equal(r.items.length, 1);
    assert.equal(r.items[0].adjustment, fee);
    assert.equal(r.items[0].grossMargin, margin);
    assert.equal(r.branchDateSummary[0].adjustment, fee);
    assert.equal(r.branchDateSummary[0].grossMargin, margin);
  }
  assert.deepEqual(cfg, snapshot);
  cfg.patterns.push({ pattern: 'BDO * SNPL CUSTOMER', cardName: 'BPI 3 MOS' });
  const ambiguous = calc.calculate(
    [sale({ payment: 'BDO 6 MOS SNPL CUSTOMER' })],
    config.validate(cfg),
  );
  assert.equal(ambiguous.items[0].adjustment, null);
  assert.match(ambiguous.reviews[0].reason, /Ambiguous/);
  assert.deepEqual(ambiguous.branchDateSummary, []);
});
test('Label removal supports configured aliases and patterns without inventing an underlying method', () => {
  const cfg = c();
  cfg.aliases.push({ alias: 'Terminal A', cardName: 'BDO 3 MOS' });
  cfg.patterns.push({ pattern: 'Terminal B *', cardName: 'BPI 3 MOS' });
  for (const payment of ['BNPL Terminal A', 'Terminal B REF123 SNPL']) {
    const r = calc.calculate([sale({ payment })], config.validate(cfg));
    assert.equal(r.items[0].adjustment, '1050.00');
    assert.equal(r.items[0].partialCalculation, false);
  }
  for (const payment of [
    'BNPL Unknown Bank',
    'Unknown Bank SNPL',
    'BNPL',
    'SNPL',
    'BNPLBDO 6 MOS',
    'BDO SNPL123 6 MOS',
    'BDO 6 MOS BNPL PROMO',
  ]) {
    const r = run([sale({ payment })]);
    assert.equal(r.items[0].adjustment, null, payment);
    assert.equal(r.items[0].grossMargin, null);
    assert.equal(r.reviews.length, 1);
    assert.deepEqual(r.branchDateSummary, []);
  }
  const onlyVariant = config.validate({
    ...c(),
    paymentMethods: [{ bank: 'BDO', cardName: 'BDO 6 MOS BNPL', rate: '7' }],
    aliases: [],
  });
  assert.equal(
    calc.calculate([sale({ payment: 'BDO 6 MOS' })], onlyVariant).items[0].adjustment,
    null,
  );
});
test('Cash and Cash Payment cannot acquire charges from names, aliases or wildcard settings', () => {
  const cfg = c();
  cfg.nonCardMethods = [];
  cfg.paymentMethods.push({ bank: 'BDO', cardName: 'Cash', rate: '99' });
  cfg.aliases.push(
    { alias: 'Cash Payment', cardName: 'BDO 3 MOS' },
    { alias: 'Till', cardName: 'Cash' },
  );
  cfg.patterns.push({ pattern: 'CASH*', cardName: 'BDO 3 MOS' });
  for (const payment of [
    'Cash',
    'Cash Payment',
    ' cash  payment ',
    'SNPL Cash',
    'Cash Payment BNPL',
    'Till',
    'Cash Payment 30,000.00',
  ]) {
    const r = calc.calculate([sale({ payment })], config.validate(cfg));
    assert.equal(r.items[0].adjustment, '0.00', payment);
    assert.equal(r.items[0].appliedRate, null);
    assert.equal(r.items[0].grossMargin, null);
    assert.equal(r.items[0].cardName, '');
    assert.equal(r.items[0].originalPaymentMethod, payment);
    assert.equal(r.summary.charges, '0.00');
    assert.equal(r.summary.subsidies, '0.00');
    assert.equal(r.summary.sales, '0.00');
    assert.equal(r.summary.feeCalculatedItems, 0);
    assert.equal(r.reviews.length, 0);
    assert.deepEqual(r.branchDateSummary, []);
  }
  const split = calc.calculate(
    [sale({ payments: [part('BDO 3 MOS BNPL', '20000'), part('Cash Payment', '10000')] })],
    cfg,
  );
  assert.equal(split.items[0].adjustment, '700.00');
  assert.equal(split.items[0].grossMargin, '4300.00');
  assert.equal(split.branchDateSummary[0].adjustment, '700.00');
  assert.equal(split.items[0].cardName, 'BDO 3 MOS');
  const unresolved = calc.calculate([sale({ payment: 'Cash Payment + BDO 3 MOS BNPL' })], cfg);
  assert.equal(
    unresolved.items[0].adjustment,
    null,
    'Missing split amounts cannot receive a whole-row rate',
  );
});
test('Broad bank patterns cannot charge cash portions or unresolved combined payments', () => {
  const cfg = c();
  cfg.patterns.push({ pattern: 'BDO *', cardName: 'BDO 3 MOS' });
  const r = calc.calculate(
    [sale({ payment: 'BDO 3 MOS BNPL 20,000.00\nCash Payment 10,000.00' })],
    config.validate(cfg),
  );
  assert.equal(r.items[0].adjustment, '700.00');
  assert.equal(r.items[0].grossMargin, '4300.00');
  assert.equal(r.branchDateSummary[0].adjustment, '700.00');
  assert.equal(r.items[0].partialCalculation, false);
  const unresolved = calc.calculate([sale({ payment: 'BDO 3 MOS BNPL + Cash Payment' })], cfg);
  assert.equal(unresolved.items[0].adjustment, null);
  assert.match(unresolved.items[0].explanation, /Split payment has no reliable/);
  assert.deepEqual(unresolved.branchDateSummary, []);
});

test('Sales-only Tender groups preserve labels and reconcile BNPL/SNPL fees across both worksheets', async () => {
  const payment = 'Original Sales Report payment text';
  const tender = [
    'BNPL BDO 6 MOS',
    'ref1',
    20000,
    'BPI 3 MOS SNPL',
    'ref2',
    5000,
    'Cash Payment',
    'ref3',
    5000,
  ];
  const rows = parse(
    workbook({
      Sales: [
        [
          'Product ID',
          'Date',
          'Branch',
          'TM-INV#',
          'Product Name',
          'Qty',
          'Price',
          'Unit Cost',
          'Discount',
          'Net Total',
          'Payment',
          'Tender',
          'Ref',
          'Amount',
          'Tender',
          'Ref',
          'Amount',
          'Tender',
          'Ref',
          'Amount',
        ],
        ['A', '2026-10-19', 'ASUS', '1', 'Laptop', 1, 20000, 16000, 0, 20000, payment, ...tender],
        ['B', '2026-10-19', 'ASUS', '1', 'Accessory', 2, 5000, 4000, 0, 10000, payment, ...tender],
      ],
    }),
  );
  const original = structuredClone(rows);
  const r = run(rows);
  assert.deepEqual(rows, original);
  assert.deepEqual(
    rows[0].payments.map((p) => p.method),
    ['BNPL BDO 6 MOS', 'BPI 3 MOS SNPL', 'Cash Payment'],
  );
  assert.deepEqual(
    r.items.map((i) => i.adjustment),
    ['916.67', '458.33'],
  );
  assert.deepEqual(
    r.items.map((i) => i.originalPaymentMethod),
    [payment, payment],
  );
  assert.equal(r.reviews.length, 0);
  assert.deepEqual(r.branchDateSummary, [
    {
      date: '2026-10-19',
      branch: 'ASUS',
      sales: '30000.00',
      totalCost: '24000.00',
      adjustment: '1375.00',
      grossMargin: '4625.00',
    },
  ]);
  const b = new ExcelJS.Workbook();
  await b.xlsx.load((await exporter.report(r)) as any);
  const main = b.worksheets[0],
    summary = b.worksheets[1];
  assert.equal(main.rowCount, 3, 'Repeated payment evidence never duplicates products');
  assert.equal(main.columnCount, 17);
  for (const row of [2, 3]) {
    assert.equal(main.getCell(row, 11).value, payment);
    assert.equal(main.getCell(row, 16).value, 'BDO 6 MOS / BPI 3 MOS');
    assert.match(String(main.getCell(row, 17).value), /Original payment\/tender: "BNPL BDO 6 MOS"/);
    assert.match(String(main.getCell(row, 17).value), /Original payment\/tender: "BPI 3 MOS SNPL"/);
  }
  assert.equal(summary.columnCount, 6);
  for (const [itemCol, summaryCol] of [
    [10, 3],
    [13, 5],
    [14, 6],
  ])
    assert.equal(
      summary.getCell(3, summaryCol).value,
      new Decimal(String(main.getCell(2, itemCol).value))
        .plus(String(main.getCell(3, itemCol).value))
        .toNumber(),
    );
  assert.equal(summary.getCell('D3').value, 24000);
});

test('Explicit card/cash split applies fees only to card portion', () => {
  const r = run([sale({ payments: [part('BDO 3 MOS', '20000'), part('Cash', '10000')] })]);
  assert.equal(r.items.length, 1);
  assert.equal(r.items[0].adjustment, '700.00');
  assert.equal(r.items[0].grossMargin, '4300.00');
  assert.equal(r.summary.sales, '30000.00');
});

test('PAYMAYA DEBIT with cash uses the configured rate on the card portion exactly once', async () => {
  const payment = 'CC#7297 PAYMAYA DEBIT DEBIT SALES 6,000.00\nCash 3,000.00';
  for (const payments of [[], [part('PAYMAYA DEBIT DEBIT SALES', '6000'), part('Cash', '3000')]]) {
    const rows = [
      sale({ payment, payments, lineTotal: '6000', cost: '4500' }),
      sale({
        payment,
        payments,
        lineTotal: '3000',
        cost: '2500',
        productId: 'SECOND',
        source: { file: 'sales.xlsx', sheet: 'Sales Report', row: 3 },
      }),
    ];
    const r = run(rows);
    assert.deepEqual(
      r.items.map((i) => i.cardName),
      ['PAYMAYA DEBIT', 'PAYMAYA DEBIT'],
    );
    assert.deepEqual(
      r.items.map((i) => i.appliedRate),
      ['2.40', '2.40'],
    );
    assert.deepEqual(
      r.items.map((i) => i.adjustment),
      ['96.00', '48.00'],
    );
    assert.deepEqual(
      r.items.map((i) => i.grossMargin),
      ['1404.00', '452.00'],
    );
    assert.ok(r.items.every((i) => i.originalPaymentMethod === payment && !i.partialCalculation));
    assert.equal(r.summary.charges, '144.00');
    assert.equal(r.summary.grossMargin, '1856.00');
    assert.equal(r.branchDateSummary[0].adjustment, '144.00');
    assert.equal(r.reviews.length, 0);
    const book = new ExcelJS.Workbook();
    await book.xlsx.load((await exporter.report(r)) as any);
    assert.equal(book.worksheets[0].getCell('K2').value, payment);
    assert.equal(book.worksheets[0].getCell('M2').value, 96);
    assert.equal(book.worksheets[0].getCell('M3').value, 48);
    assert.equal(book.worksheets[1].getCell('E3').value, 144);
    const cfg = c();
    cfg.paymentMethods.find((m) => m.cardName === 'PAYMAYA DEBIT')!.rate = '3.00';
    assert.equal(calc.calculate(rows, config.validate(cfg)).summary.charges, '180.00');
  }
});

test('PAYMAYA and cash tender with change calculates from Sales Report inputs and reconciles both sheets', async () => {
  const payment = 'CC#7297 PAYMAYA DEBIT DEBIT SALES 6,000.00\nCash 3,000.00';
  for (const useTenderColumns of [false, true]) {
    const headers = [
      'Product ID',
      'Date',
      'Branch',
      'Invoice Number',
      'Product Name',
      'Qty',
      'Price',
      'Unit Cost',
      'Discount',
      'Net Total',
      'Payment',
    ];
    const row: unknown[] = [
      'PAYMAYA-ITEM',
      '2026-10-01',
      'ASUS',
      'PM-1',
      'Product',
      1,
      8995,
      7000,
      0,
      8995,
      payment,
    ];
    if (useTenderColumns) {
      headers.push('Tender', 'Ref', 'Amount', 'Tender', 'Ref', 'Amount');
      row.push('PAYMAYA DEBIT DEBIT SALES', '7297', 6000, 'Cash', '', 3000);
    }
    const rows = parse(workbook({ 'Sales Report': [headers, row] }));
    const original = structuredClone(rows);
    const r = run(rows),
      item = r.items[0];
    assert.equal(item.cardName, 'PAYMAYA DEBIT');
    assert.equal(item.appliedRate, '2.40');
    assert.equal(item.adjustment, '144.00');
    assert.equal(item.grossMargin, '1851.00');
    assert.equal(item.originalPaymentMethod, payment);
    assert.equal(item.lineTotal, '8995');
    assert.equal(item.partialCalculation, false);
    assert.equal(item.status, 'Calculated');
    assert.equal(r.reviews.length, 0);
    assert.match(
      item.explanation,
      /Cash tender ₱3000.00 includes ₱5.00 change; net cash payment ₱2995.00/,
    );
    assert.match(item.explanation, /Eligible payment ₱6000.00/);
    assert.equal(r.summary.sales, '8995.00');
    assert.equal(r.summary.charges, '144.00');
    assert.deepEqual(r.branchDateSummary, [
      {
        date: '2026-10-01',
        branch: 'ASUS',
        sales: '8995.00',
        totalCost: '7000.00',
        adjustment: '144.00',
        grossMargin: '1851.00',
      },
    ]);
    const book = new ExcelJS.Workbook();
    await book.xlsx.load((await exporter.report(r)) as any);
    const main = book.getWorksheet('Item Bank Deductions')!;
    assert.equal(main.getCell('K2').value, payment);
    assert.equal(main.getCell('L2').value, 0.024);
    assert.equal(main.getCell('M2').value, 144);
    assert.equal(main.getCell('N2').value, 1851);
    assert.deepEqual(
      (book.worksheets[1].getRow(3).values as unknown[]).slice(3),
      [8995, 7000, 144, 1851],
    );
    assert.deepEqual(rows, original, 'Tender amounts and original Payment are preserved');
    for (const [rate, adjustment, grossMargin] of [
      ['3.00', '180.00', '1815.00'],
      ['-2.00', '-120.00', '2115.00'],
      ['0', '0.00', '1995.00'],
    ]) {
      const cfg = c();
      cfg.paymentMethods.find((m) => m.cardName === 'PAYMAYA DEBIT')!.rate = rate;
      const changed = calc.calculate(rows, config.validate(cfg));
      assert.equal(changed.items[0].adjustment, adjustment);
      assert.equal(changed.items[0].grossMargin, grossMargin);
    }
  }
});

test('Cash change and repeated bank payments are allocated once across invoice products', () => {
  const payments = [
    part('PAYMAYA DEBIT', '6000'),
    part('Cash', '1500'),
    part('Cash Payment', '1500'),
  ];
  const payment = 'PAYMAYA DEBIT 6,000.00\nCash 1,500.00\nCash Payment 1,500.00';
  const r = run([
    sale({ payment, payments, lineTotal: '5995', cost: '4500' }),
    sale({
      payment,
      payments,
      lineTotal: '3000',
      cost: '2500',
      source: { file: 'sales.xlsx', sheet: 'Sales Report', row: 3 },
    }),
  ]);
  assert.deepEqual(
    r.items.map((i) => i.adjustment),
    ['95.97', '48.03'],
  );
  assert.deepEqual(
    r.items.map((i) => i.grossMargin),
    ['1399.03', '451.97'],
  );
  assert.equal(r.summary.charges, '144.00');
  assert.equal(r.summary.sales, '8995.00');
  assert.equal(r.summary.grossMargin, '1851.00');
  assert.equal(r.branchDateSummary[0].adjustment, '144.00');
  assert.equal(
    r.bankSummary
      .filter((b) => b.bank === 'No bank charge')
      .reduce((n, b) => n.plus(b.sales), new Decimal(0))
      .toFixed(2),
    '2995.00',
  );
  assert.equal(r.reviews.length, 0);
});

test('Cash change never reduces bank amounts, invents cash or resolves unknown and invalid payments', () => {
  for (const payments of [
    [part('PAYMAYA DEBIT', '9000')],
    [part('PAYMAYA DEBIT', '9000'), part('Cash', '5')],
    [part('PAYMAYA DEBIT', '6000'), part('Voucher', '3000')],
    [part('PAYMAYA DEBIT', '5000'), part('Unknown', '1000'), part('Cash', '3000')],
    [part('PAYMAYA DEBIT', '6000'), part('Cash', '2000')],
    [part('PAYMAYA DEBIT', '6000'), part('Cash', '')],
  ]) {
    const r = run([sale({ payments, lineTotal: '8995' })]);
    assert.equal(r.items[0].adjustment, null);
    assert.deepEqual(r.branchDateSummary, []);
  }
  const cfg = c();
  cfg.nonCardMethods.push('Voucher');
  assert.equal(
    calc.calculate(
      [
        sale({
          lineTotal: '8995',
          payments: [part('PAYMAYA DEBIT', '6000'), part('Voucher', '3000')],
        }),
      ],
      config.validate(cfg),
    ).items[0].adjustment,
    null,
  );
  const refunded = run([
    sale({
      quantity: '-1',
      lineTotal: '-8995',
      payments: [part('PAYMAYA DEBIT', '6000'), part('Cash', '3000')],
    }),
  ]);
  assert.equal(refunded.items[0].adjustment, null);
  assert.match(refunded.items[0].explanation, /Negative sales \/ refund/);
});

test('Combined methods inside descriptions are not mistaken for a single card payment', () => {
  for (const method of ['BDO 3 MOS / Cash', 'BDO STRAIGHT & Cash']) {
    const r = run([sale({ payments: [part(method, '30000')] })]);
    assert.equal(r.items[0].adjustment, null);
    assert.equal(r.summary.charges, '0.00');
  }
});

test('Currency-formatted source totals use the same validated decimal parser in summaries', () => {
  const r = run([sale({ lineTotal: '  ₱30,000.00', price: 'PHP 30,000.00' })]);
  assert.equal(r.summary.sales, '30000.00');
  assert.equal(r.items[0].adjustment, '1050.00');
});
test('Two banks retain one item row with both rates and reconciled summary margins', () => {
  const r = run([sale({ payments: [part('BDO 3 MOS', '20000'), part('BPI 12 MOS', '10000')] })]);
  assert.equal(r.items.length, 1);
  assert.equal(r.items[0].adjustment, '1950.00');
  assert.equal(r.items[0].grossMargin, '3050.00');
  assert.match(r.items[0].appliedRate!, /BPI 12 MOS/);
  assert.equal(Decimal.sum(...r.bankSummary.map((b) => b.grossMargin)).toFixed(2), '3050.00');
});
test('Unit cost and total line cost are interpreted explicitly and displayed consistently on both sheets', async () => {
  const common = {
    quantity: '2',
    lineTotal: '60000',
    price: '30000',
    cost: '25000',
    payment: 'BDO 3 MOS',
  };
  const unit = run([sale({ ...common, costBasis: 'unit' })]),
    line = run([sale({ ...common, costBasis: 'line' })]),
    unknown = run([sale({ ...common, costBasis: 'unknown' })]);
  assert.equal(unit.items[0].grossMargin, '7900.00');
  assert.equal(line.items[0].grossMargin, '32900.00');
  assert.equal(unknown.items[0].grossMargin, null);
  assert.equal(unknown.items[0].adjustment, '2100.00');
  for (const [result, displayedCost] of [
    [unit, 50000],
    [line, 25000],
    [unknown, null],
  ] as const) {
    const book = new ExcelJS.Workbook();
    await book.xlsx.load((await exporter.report(result)) as any);
    const main = book.worksheets[0];
    assert.equal(main.getCell('H2').value, displayedCost);
    assert.match(JSON.stringify(main.getCell('H2').note), /Original source Cost:.*25000/);
    assert.match(result.items[0].explanation, /Source Cost "25000"/);
    assert.equal(
      result.items[0].cost,
      '25000',
      'The original cost is retained in the source audit',
    );
    if (displayedCost !== null) {
      assert.equal(book.worksheets[1].getCell('D3').value, displayedCost);
      assert.match(result.items[0].explanation, /Included in Sheet 2:/);
    } else {
      assert.deepEqual(result.branchDateSummary, []);
      assert.match(result.items[0].explanation, /Excluded from Sheet 2:/);
    }
  }
});
test('Negative rates create subsidies and increase gross margin', () => {
  const cfg = c();
  cfg.paymentMethods.find((m) => m.cardName === 'BDO 3 MOS')!.rate = '-2.00';
  config.validate(cfg);
  const r = calc.calculate([sale()], cfg);
  assert.equal(r.items[0].adjustment, '-600.00');
  assert.equal(r.items[0].grossMargin, '5600.00');
  assert.equal(r.summary.subsidies, '600.00');
  assert.equal(r.summary.charges, '0.00');
});
test('Rounding, zero-price gifts and separate payment fees reconcile in centavos', () => {
  const cfg = c();
  cfg.paymentMethods.find((m) => m.cardName === 'BDO 3 MOS')!.rate = '50.00';
  const payments = [part('BDO 3 MOS', '0.05')],
    rows = ['0.01', '0.02', '0.02', '0'].map((net, i) =>
      sale({
        source: { file: 'sales.xlsx', sheet: 'Sales Report', row: i + 2 },
        lineTotal: net,
        cost: '0',
        price: net,
        payments,
      }),
    );
  const r = calc.calculate(rows, cfg);
  assert.equal(r.summary.charges, '0.03');
  assert.equal(r.items[3].adjustment, '0.00');
});
test('Missing payment, missing discount and invalid sales retain original rows', () => {
  const rows = [sale({ payment: '' }), sale({ discount: '' }), sale({ lineTotal: 'bad' })];
  const r = run(rows);
  assert.equal(r.items.length, 3);
  assert.equal(r.items[0].status, 'Missing Payment Method');
  assert.equal(r.items[1].discount, '');
  assert.ok(r.items[1].issues.some((x) => x.includes('Discount')));
  assert.equal(r.items[2].status, 'Invalid Sales Amount');
  assert.equal(r.items[2].adjustment, null);
});
test('Source row numbers, Product ID text and product association survive parsing', () => {
  const rows = parse(reportFixture());
  assert.equal(rows[0].productId, '000123');
  assert.equal(rows[1].productId, '000124');
  assert.equal(rows[0].source.row, 2);
  assert.equal(rows[1].source.row, 3);
  const r = run(rows);
  assert.equal(r.items[1].product, 'Mouse');
  assert.equal(r.items[1].totalCost, '600.00');
  assert.equal(r.items[1].grossMargin, null);
});
test('All Product ID alias headers auto-detect', () => {
  for (const header of ['Product ID', 'ProductID', 'SKU', 'Item Code', 'Product Code']) {
    const b = XLSX.read(reportFixture(), { type: 'buffer' });
    b.Sheets['Sales Report'].A1.v = header;
    const rows = parse(XLSX.write(b, { type: 'buffer', bookType: 'xlsx' }));
    assert.equal(rows[0].productId, '000123');
  }
});
test('Formatted numeric Product IDs preserve recoverable leading zeros in xlsx and xls', () => {
  for (const bookType of ['xlsx', 'biff8'] as const) {
    const b = XLSX.read(reportFixture(), { type: 'buffer' });
    b.Sheets['Sales Report'].A2 = { t: 'n', v: 123, z: '000000' };
    const f = upload(
      XLSX.write(b, { type: 'buffer', bookType }),
      bookType === 'xlsx' ? 'sales.xlsx' : 'sales.xls',
    );
    const p = excel.preview(f).sheets[0];
    assert.equal(
      excel.parse(f, { sheet: p.name, headerRow: 1, mapping: p.suggestedMapping })[0].productId,
      '000123',
    );
  }
});
test('Missing Product ID is blank with a warning and valid amounts remain calculated', () => {
  const rows = parse(reportFixture(), { productId: '' });
  const r = run(rows);
  assert.equal(r.items[0].productId, '');
  assert.equal(r.items[0].adjustment, '1050.00');
  assert.equal(r.items[0].status, 'Calculated');
  assert.match(r.items[0].warnings.join(' '), /Missing Product ID/);
  assert.equal(r.summary.missingProductIds, 2);
});
test('Duplicate Product IDs are never deduplicated across invoices', () => {
  const r = run([
    sale(),
    sale({ invoice: '0002', source: { file: 'sales.xlsx', sheet: 'Sales Report', row: 8 } }),
  ]);
  assert.equal(r.items.length, 2);
  assert.equal(r.items[1].productId, r.items[0].productId);
  assert.equal(r.items[1].sourceRow, 8);
});
test('Recognized spacing, explicit aliases and source descriptions map to configured names', () => {
  assert.equal(matchRate('BDO 3MOS', c()).method?.cardName, 'BDO 3 MOS');
  assert.equal(matchRate('BDO STRAIGHT CUSTOMER NAME', c(), true).method?.cardName, 'BDO STRAIGHT');
  const cfg = c();
  cfg.aliases = [{ alias: 'Terminal A', cardName: 'BPI 12 MOS' }];
  const r = calc.calculate([sale({ payment: 'Terminal A' })], cfg);
  assert.equal(r.items[0].originalPaymentMethod, 'Terminal A');
  assert.equal(r.items[0].cardName, 'BPI 12 MOS');
});
test('An explicitly configured BDO PAY to BDO STRAIGHT alias calculates each payment once across worksheets', async () => {
  const payment = 'CC#7005 BDO STRAIGHT 7005 31,282.00\nBDO PAY 10,000.00';
  // A synthetic companion product supplies the remaining invoice sales for allocation.
  const rows = parse(
    workbook({
      'Sales Report': [
        [
          'Product ID',
          'Date',
          'Branch',
          'TM-INV#',
          'Product Name',
          'Qty',
          'Price',
          'Unit Cost',
          'Discount',
          'Net Total',
          'Payment',
        ],
        [
          '710425578649',
          '2026-10-01',
          'GAME ONE MEGAMALL',
          '3074-00001539',
          'PS5 G. GTA V (R1)',
          1,
          1250,
          1118,
          0,
          1250,
          payment,
        ],
        [
          'COMPANION',
          '2026-10-01',
          'GAME ONE MEGAMALL',
          '3074-00001539',
          'Remaining invoice sales',
          1,
          40032,
          32000,
          0,
          40032,
          payment,
        ],
      ],
    }),
  );
  const original = structuredClone(rows);
  const r = calc.calculate(rows, config.validate(c()));
  assert.deepEqual(rows, original, 'Source payment descriptions and amounts are preserved');
  assert.deepEqual(
    r.items.map((item) => item.adjustment),
    ['35.63', '1140.91'],
  );
  assert.deepEqual(
    r.items.map((item) => item.grossMargin),
    ['96.37', '6891.09'],
  );
  for (const item of r.items) {
    assert.equal(item.cardName, 'BDO STRAIGHT');
    assert.equal(item.appliedRate, '2.85');
    assert.equal(item.originalPaymentMethod, payment);
    assert.equal(item.partialCalculation, false);
    assert.equal(item.status, 'Calculated');
    assert.doesNotMatch(item.explanation, /Unknown payment|Unconfigured|Provisional/);
  }
  assert.equal(r.reviews.length, 0);
  assert.equal(r.summary.charges, '1176.54', 'Repeated invoice payments are charged only once');
  assert.deepEqual(r.branchDateSummary, [
    {
      date: '2026-10-01',
      branch: 'GAME ONE MEGAMALL',
      sales: '41282.00',
      totalCost: '33118.00',
      adjustment: '1176.54',
      grossMargin: '6987.46',
    },
  ]);
  for (const key of ['adjustment', 'grossMargin'] as const)
    assert.equal(
      Decimal.sum(...r.items.map((item) => item[key]!)).toFixed(2),
      r.branchDateSummary[0][key],
    );

  const buffer = await exporter.report(r);
  const book = new ExcelJS.Workbook();
  await book.xlsx.load(buffer as any);
  const items = book.worksheets[0],
    summary = book.worksheets[1];
  assert.equal(items.rowCount, 3);
  assert.equal(items.getCell('K2').value, payment);
  assert.equal(items.getCell('L2').value, 0.0285);
  assert.equal(items.getCell('M2').value, 35.63);
  assert.equal(items.getCell('N2').value, 96.37);
  assert.equal(items.getCell('P2').value, 'BDO STRAIGHT');
  assert.equal(summary.columnCount, 6);
  assert.deepEqual(
    (summary.getRow(3).values as unknown[]).slice(3),
    [41282, 33118, 1176.54, 6987.46],
  );
  assert.deepEqual(
    exporter.import(upload(buffer)),
    r.configuration,
    'The alias survives configuration export/import',
  );
});

test('BDO PAY uses the editable configured rate and requires a configured target method', () => {
  const cfg = config.validate(c());
  for (const payment of [
    'BDO PAY',
    ' bdo   pay ',
    'BNPL BDO PAY',
    'BDO PAY SNPL',
    'CC#7005 BDO PAY 7005',
  ])
    assert.equal(matchRate(payment, cfg, true).method?.cardName, 'BDO STRAIGHT');
  const payments = [part('BDO STRAIGHT', '1000'), part('BDO PAY', '1000'), part('Cash', '1000')];
  const rows = [sale({ price: '3000', lineTotal: '3000', cost: '2500', payments })];
  for (const [rate, charge, margin] of [
    ['4.00', '80.00', '420.00'],
    ['0.00', '0.00', '500.00'],
    ['-2.00', '-40.00', '540.00'],
  ]) {
    cfg.paymentMethods.find((method) => method.cardName === 'BDO STRAIGHT')!.rate = rate;
    const r = calc.calculate(rows, config.validate(cfg));
    assert.equal(r.items[0].appliedRate, rate);
    assert.equal(
      r.items[0].adjustment,
      charge,
      'Both bank portions use the configured rate; cash remains uncharged',
    );
    assert.equal(r.items[0].grossMargin, margin);
    assert.equal(r.branchDateSummary[0].adjustment, charge);
    assert.equal(r.branchDateSummary[0].grossMargin, margin);
    assert.equal(r.reviews.length, 0);
  }
  cfg.aliases = [];
  cfg.paymentMethods = cfg.paymentMethods.filter(
    (method) => !['BDO STRAIGHT', 'BDO PAY/CASA'].includes(method.cardName),
  );
  const unresolved = calc.calculate([sale({ payment: 'BDO PAY' })], config.validate(cfg));
  assert.equal(unresolved.items[0].status, 'Missing Bank Rate');
  assert.equal(unresolved.items[0].adjustment, null);
  assert.deepEqual(unresolved.branchDateSummary, []);
  assert.equal(unresolved.reviews.length, 1);
});

test('Older settings automatically use the default alias and snapshot the actual configured rate', async () => {
  const oldSettings = c();
  oldSettings.aliases = [];
  oldSettings.paymentMethods.find((method) => method.cardName === 'BDO STRAIGHT')!.rate = '4.00';
  oldSettings.paymentMethods.find((method) => method.cardName === 'BDO PAY/CASA')!.rate = '4.00';
  const imported = exporter.import(upload(await exporter.configuration(oldSettings)));
  assert.deepEqual(imported.aliases, []);
  const snapshot = structuredClone(imported);
  const payment = 'CC#7005 BDO STRAIGHT 7005 2,000.00\nBDO PAY 1,000.00';
  const rows = [sale({ payment, price: '3000', lineTotal: '3000', cost: '2500' })];
  const r = calc.calculate(rows, imported);
  assert.deepEqual(imported, snapshot, 'Calculation does not mutate the supplied settings');
  assert.equal(r.items[0].originalPaymentMethod, payment);
  assert.equal(r.items[0].cardName, 'BDO STRAIGHT / BDO PAY/CASA');
  assert.equal(r.items[0].appliedRate, 'BDO STRAIGHT: 4.00%; BDO PAY/CASA: 4.00%');
  assert.equal(r.items[0].adjustment, '120.00');
  assert.equal(r.items[0].grossMargin, '380.00');
  assert.equal(r.items[0].partialCalculation, false);
  assert.equal(r.reviews.length, 0);
  assert.equal(r.branchDateSummary[0].adjustment, '120.00');
  assert.equal(r.branchDateSummary[0].grossMargin, '380.00');
  assert.deepEqual(r.configuration.aliases, [{ alias: 'BDO PAY', cardName: 'BDO PAY/CASA' }]);
  assert.deepEqual(r.configuration.paymentMethods, imported.paymentMethods);
  const restored = exporter.import(upload(await exporter.report(r)));
  assert.deepEqual(restored, r.configuration);
  const repeated = calc.calculate(rows, restored);
  assert.deepEqual(repeated.configuration, r.configuration);
  assert.deepEqual(repeated.items, r.items, 'Recalculation does not add a second alias or charge');
});

test('Automatic aliases preserve explicit methods, aliases, patterns, no-charge rules and ambiguity', () => {
  const row = sale({ payment: 'BDO PAY', price: '1000', lineTotal: '1000', cost: '800' });
  for (const rule of ['method', 'alias', 'pattern', 'non-card', 'ambiguous']) {
    const cfg = c();
    cfg.aliases = [];
    if (rule === 'method')
      cfg.paymentMethods.push({ bank: 'BDO', cardName: 'BDO PAY', rate: '1.25' });
    if (rule === 'alias') cfg.aliases.push({ alias: ' bdo pay ', cardName: 'BDO 3 MOS' });
    if (rule === 'pattern') cfg.patterns.push({ pattern: 'BDO P*', cardName: 'BDO 3 MOS' });
    if (rule === 'non-card') cfg.nonCardMethods.push('BDO PAY');
    if (rule === 'ambiguous')
      cfg.patterns.push(
        { pattern: 'BDO P*', cardName: 'BDO 3 MOS' },
        { pattern: 'BDO *', cardName: 'BDO 6 MOS' },
      );
    const validated = config.validate(cfg);
    const r = calc.calculate([row], validated);
    assert.deepEqual(r.configuration, validated, rule);
    assert.equal(
      r.items[0].adjustment,
      rule === 'method' ? '12.50' : rule === 'alias' || rule === 'pattern' ? '35.00' : null,
      rule,
    );
    if (rule === 'ambiguous') assert.match(r.reviews[0].reason, /Ambiguous/);
  }
});

test('Repeated Sales Report payment groups, actual headers and cost basis are recognized', () => {
  const f = upload(
    workbook({
      'Sales Report': [
        [
          'Product ID',
          'Date',
          'Branch',
          'TM-INV#',
          'Product Name',
          'Qty',
          'Price',
          'Cost',
          'Discount',
          'Line\n Total',
          'Base\n Qty',
          'Net\n Total',
          'Master Price',
          'Payment',
          'Tender',
          'Ref',
          'Amount',
          'Tender',
          'Ref',
          'Amount',
        ],
        [
          '001',
          46296,
          'ASUS',
          '0001',
          'Laptop',
          2,
          15000,
          10000,
          0,
          30000,
          2,
          30000,
          15000,
          'CC#0001 BDO 3 MOS NAME 20,000.00\nCash 10,000.00',
          'BDO 3 MOS NAME',
          '0001',
          20000,
          'Cash',
          '',
          10000,
        ],
      ],
    }),
  );
  const p = excel.preview(f).sheets[0];
  assert.equal(p.suggestedCostBasis, 'unit');
  assert.equal(p.suggestedMapping.lineTotal, 'Net Total');
  const rows = excel.parse(f, { sheet: p.name, headerRow: 1, mapping: p.suggestedMapping });
  const r = run(rows);
  assert.equal(r.items[0].adjustment, '700.00');
  assert.equal(r.items[0].grossMargin, '9300.00');
  assert.equal(rows[0].payment, rows[0].original.payment);
});
test('Inconsistent or short invoice payment evidence never invents a missing portion', () => {
  const a = sale({ lineTotal: '20000', payments: [part('BDO 3 MOS', '30000')] }),
    b = sale({
      lineTotal: '10000',
      payments: [part('Cash', '10000')],
      source: { file: 'sales.xlsx', sheet: 'Sales Report', row: 3 },
    });
  assert.ok(run([a, b]).items.every((i) => i.adjustment === null));
  assert.equal(run([sale({ payments: [part('BDO 3 MOS', '20000')] })]).items[0].adjustment, null);
});
test('Requested warning text is hidden without removing rows, resolving payments or hiding other warnings', async () => {
  const variants: Partial<Sale>[] = [
    { payment: 'Charge' },
    {
      payment: 'BDO 3 MOS 9,000.00\nCharge 500.00\nOther 500.00',
      lineTotal: '10000',
      cost: '7000',
    },
    {
      payment: 'BDO 6 MOS',
      lineTotal: '9800',
      price: '10000',
      discount: '200',
      payments: [part('BDO 6 MOS', '10000')],
    },
    { payment: 'BDO 6 MOS', lineTotal: '10200', payments: [part('BDO 6 MOS', '10000')] },
    { payment: 'BDO 6 MOS BNPL', quantity: '-1', lineTotal: '-100' },
    { payment: 'BDO 6 MOS', lineTotal: '1000', cost: '800' },
  ];
  const rows = variants.map((variant, i) =>
    sale({
      ...variant,
      invoice: 'INV-' + i,
      productId: 'ITEM-' + i,
      source: { file: 'sales.xlsx', sheet: 'Sales Report', row: i + 2 },
    }),
  );
  const originals = structuredClone(rows);
  const r = run(rows);
  const hidden =
    /"Charge": Unknown payment method or missing configured bank rate\.|Invoice payment discrepancy:/;
  assert.equal(r.items.length, rows.length);
  assert.equal(r.reviews.length, 5);
  for (const item of r.items) assert.doesNotMatch(item.explanation, hidden);
  for (const review of r.reviews) assert.doesNotMatch(review.reason, hidden);
  assert.match(r.items[0].issues.join(' '), hidden, 'Original diagnostics remain available');
  assert.equal(r.items[0].status, 'Missing Bank Rate');
  assert.equal(r.items[0].adjustment, null);
  assert.equal(r.items[1].partialCalculation, true);
  assert.equal(r.items[1].adjustment, '315.00');
  assert.match(r.items[1].explanation, /"Other": Unknown payment method/);
  assert.match(r.reviews[1].reason, /"Other": Unknown payment method/);
  for (const item of r.items.slice(2, 4)) {
    assert.equal(item.status, 'Needs Review');
    assert.equal(item.adjustment, null, 'Hiding a discrepancy never calculates its charge');
    assert.equal(item.grossMargin, null);
    assert.match(item.issues.join(' '), /Invoice payment discrepancy:/);
  }
  assert.match(
    r.items[2].issues.join(' '),
    /explicit payments 10000.00 versus net product sales 9800.00 \(difference 200.00\)/,
  );
  assert.match(r.items[3].issues.join(' '), /difference -200.00/);
  assert.match(r.items[4].explanation, /Negative sales \/ refund/);
  assert.match(r.reviews[4].reason, /Negative quantity \/ return/);
  assert.equal(r.summary.sales, '1000.00');
  assert.equal(r.summary.charges, '60.00');
  assert.equal(r.branchDateSummary[0].grossMargin, '140.00');
  const snapshot = structuredClone(r);
  const book = new ExcelJS.Workbook();
  await book.xlsx.load((await exporter.report(r)) as any);
  const main = book.getWorksheet('Item Bank Deductions')!;
  assert.equal(main.rowCount, 5, 'Matched items remain except negative quantities moved to review');
  assert.equal(main.getCell('K2').value, rows[1].payment);
  assert.equal(main.getCell('M2').value, 315);
  assert.equal(main.getCell('M3').value, null);
  assert.equal(main.getCell('J3').value, 9800);
  assert.equal(main.getCell('I3').value, 200);
  assert.equal(main.getCell('A5').value, 'ITEM-5');
  assert.equal(main.getCell('F5').value, 1);
  assert.equal(main.getCell('J5').value, 1000);
  for (let row = 2; row <= main.rowCount; row++)
    assert.doesNotMatch(String(main.getCell(row, 17).value ?? ''), hidden);
  const review = book.getWorksheet('Needs Review')!;
  assert.equal(review.rowCount, 6);
  for (let row = 2; row <= review.rowCount; row++)
    assert.doesNotMatch(String(review.getCell(row, 8).value ?? ''), hidden);
  assert.match(String(review.getCell('H3').value), /"Other": Unknown payment method/);
  assert.match(String(review.getCell('H6').value), /Negative sales \/ refund/);
  assert.deepEqual(r, snapshot);
  assert.deepEqual(rows, originals);
});

test('Text payment evidence with explicit amounts is parsed without duplicating invoice payments', () => {
  const r = run([sale({ payment: 'CC#0123 BDO 3 MOS CUSTOMER 20,000.00\nCash 10,000.00' })]);
  assert.equal(r.items[0].adjustment, '700.00');
});
test('Source dates remain October 1; invalid dates and refunds go to review', () => {
  assert.equal(reportDate({ t: 'n', v: 46296 }), '2026-10-01');
  assert.equal(reportDate({ t: 'n', v: 44834 }, true), '2026-10-01');
  assert.equal(reportDate({ t: 's', v: '10/01/2026' }), '2026-10-01');
  assert.equal(reportDate({ t: 's', v: '02/30/2026' }), null);
  const r = run([sale({ quantity: '-1', lineTotal: '-30000' })]);
  assert.equal(r.items[0].adjustment, null);
});
test('Negative BDO 6 MOS quantities move to review with original source values and no financial contributions', async () => {
  const sourceRows = [699, 808, 815, 1463, 1587];
  const negativeItems = sourceRows.map((row) =>
    sale({
      source: { file: 'sales.xlsx', sheet: 'Sales Report', row },
      productId: 'RETURN-' + row,
      quantity: '-1',
      lineTotal: '-100.00',
      payment: 'BDO 6 MOS BNPL',
      payments: [part('BDO 6 MOS BNPL', '-500.00')],
      cells: { quantity: 'F' + row, lineTotal: 'J' + row },
      original: { quantity: '-1', lineTotal: '-100.00', payment: 'BDO 6 MOS BNPL' },
    }),
  );
  const rows = [...negativeItems, sale({ invoice: 'VALID', payment: 'BDO 6 MOS' })];
  const before = structuredClone(rows);
  const r = run(rows);
  for (const [i, item] of r.items.slice(0, 5).entries()) {
    assert.equal(item.quantity, '-1');
    assert.equal(item.lineTotal, '-100.00');
    assert.equal(item.originalPaymentMethod, 'BDO 6 MOS BNPL');
    assert.equal(item.cardName, 'BDO 6 MOS');
    assert.equal(item.status, 'Needs Review');
    assert.equal(item.appliedRate, null);
    assert.equal(item.adjustment, null);
    assert.equal(item.grossMargin, null);
    assert.ok(item.explanation.includes('configured rate 6.00%'));
    assert.ok(item.explanation.includes(`Net Total at cell J${sourceRows[i]} is "-100.00"`));
    assert.ok(item.explanation.includes(`Quantity at cell F${sourceRows[i]} is "-1"`));
    assert.match(item.explanation, /Original negative value retained/);
    assert.doesNotMatch(item.explanation, /Invalid Sales Amount|Invalid Quantity/);
    assert.match(item.explanation, /source rows 699, 808, 815, 1463, 1587 require review/);
  }
  assert.equal(r.reviews.length, 5);
  assert.equal(r.items[5].adjustment, '1800.00');
  assert.equal(r.items[5].grossMargin, '3200.00');
  assert.equal(r.summary.charges, '1800.00');
  assert.deepEqual(r.branchDateSummary, [
    {
      date: '2026-10-01',
      branch: 'ASUS',
      sales: '30000.00',
      totalCost: '25000.00',
      adjustment: '1800.00',
      grossMargin: '3200.00',
    },
  ]);
  const book = new ExcelJS.Workbook();
  await book.xlsx.load((await exporter.report(r)) as any);
  const main = book.getWorksheet('Item Bank Deductions')!;
  assert.equal(main.rowCount, 2, 'Only the valid positive-quantity item remains on Sheet 1');
  assert.equal(main.getCell('A2').value, r.items[5].productId);
  assert.equal(main.getCell('M2').value, 1800);
  const review = book.getWorksheet('Needs Review')!;
  for (let i = 0; i < negativeItems.length; i++) {
    const row = review.getRow(i + 2);
    assert.equal(row.getCell(1).value, negativeItems[i].productId);
    assert.equal(row.getCell(6).value, sourceRows[i]);
    assert.match(String(row.getCell(8).value), /Negative quantity \/ return/);
    assert.deepEqual(JSON.parse(String(row.getCell(10).value)), negativeItems[i].original);
  }
  assert.equal(book.getWorksheet('Needs Review')!.rowCount, 6);
  assert.equal(book.getWorksheet(BRANCH_DATE_SUMMARY_SHEET)!.getCell('E3').value, 1800);
  assert.deepEqual(rows, before, 'Source values are never changed to make a refund calculable');
});

test('Negative quantities exclude positive prices and costs while valid losses and subsidies stay in totals', async () => {
  const cfg = c();
  cfg.paymentMethods.find((m) => m.cardName === 'BDO 3 MOS')!.rate = '-2.00';
  const quantities = ['2', '-1', '-2', '-0.5', ' -3.0 '];
  const rows = quantities.map((quantity, index) =>
    sale({
      quantity,
      productId: 'ITEM-' + index,
      invoice: 'INV-' + index,
      lineTotal: '100',
      price: '100',
      cost: '60',
      source: { file: 'sales.xlsx', sheet: 'Sales Report', row: index + 2 },
      original: { quantity, price: '100', lineTotal: '100', cost: '60' },
    }),
  );
  const original = structuredClone(rows);
  const r = calc.calculate(rows, config.validate(cfg));
  assert.equal(r.items.length, quantities.length);
  assert.equal(r.reviews.length, 4);
  assert.equal(r.summary.sales, '100.00');
  assert.equal(r.summary.grossMargin, '-18.00');
  assert.deepEqual(r.branchDateSummary, [
    {
      date: '2026-10-01',
      branch: 'ASUS',
      sales: '100.00',
      totalCost: '120.00',
      adjustment: '-2.00',
      grossMargin: '-18.00',
    },
  ]);
  for (const [index, review] of r.reviews.entries()) {
    assert.equal(review.productId, 'ITEM-' + (index + 1));
    assert.match(review.reason, /Negative quantity \/ return/);
    assert.deepEqual(review.original, rows[index + 1].original);
  }
  // The summary also rejects negative quantities if a record already has financial fields.
  const completedReturns = quantities.slice(1).map((quantity) => ({ ...r.items[0], quantity }));
  assert.deepEqual(
    summarizeByBranchAndDate([r.items[0], ...completedReturns]),
    r.branchDateSummary,
  );
  assert.deepEqual(summarizeByBranchAndDate(completedReturns), []);
  const book = new ExcelJS.Workbook();
  await book.xlsx.load((await exporter.report(r)) as any);
  assert.equal(book.worksheets[0].rowCount, 2);
  assert.equal(book.worksheets[0].getCell('A2').value, 'ITEM-0');
  assert.equal(book.worksheets[0].getCell('N2').value, -18);
  assert.equal(book.getWorksheet('Needs Review')!.rowCount, 5);
  assert.deepEqual(
    (book.worksheets[1].getRow(3).values as unknown[]).slice(3),
    [100, 120, -2, -18],
  );
  assert.deepEqual(
    (book.worksheets[1].getRow(4).values as unknown[]).slice(3),
    [100, 120, -2, -18],
  );
  assert.deepEqual(rows, original);
});

test('Refund warnings distinguish shared invoice holds from independently calculable sales', () => {
  const rows = [
    sale({
      quantity: '-1',
      lineTotal: '-10000',
      payment: 'BDO 6 MOS',
      source: { file: 'sales.xlsx', sheet: 'Sales Report', row: 699 },
    }),
    sale({ payment: 'BDO 6 MOS', source: { file: 'sales.xlsx', sheet: 'Sales Report', row: 700 } }),
  ];
  const separate = run(rows);
  assert.equal(separate.items[0].adjustment, null);
  assert.equal(separate.items[1].adjustment, '1800.00');
  assert.deepEqual(separate.items[1].issues, []);
  const shared = run(rows.map((row) => ({ ...row, payments: [part('BDO 6 MOS', '20000')] })));
  assert.ok(shared.items.every((item) => item.adjustment === null && item.grossMargin === null));
  assert.match(shared.items[0].explanation, /Negative sales \/ refund/);
  assert.doesNotMatch(
    shared.items[1].explanation,
    /Negative sales \/ refund|Invalid Sales Amount|Invalid Quantity/,
  );
  assert.match(
    shared.items[1].explanation,
    /Shared payment allocation for invoice "0001" is on hold: source rows 699 require review/,
  );
  assert.deepEqual(shared.branchDateSummary, []);
});

test('Blank and invalid source values have specific warnings and cannot bypass quantity validation', () => {
  for (const quantity of ['', 'bad', '0', '1e2', '0x10', '9'.repeat(400)]) {
    const r = run([sale({ quantity })]);
    assert.equal(r.items[0].adjustment, null);
    assert.equal(r.items[0].grossMargin, null);
    assert.match(r.items[0].explanation, /Invalid Quantity: Quantity at source row 2/);
    assert.doesNotMatch(r.items[0].explanation, /refund|return/i);
  }
  for (const lineTotal of ['', 'bad', '#FORMULA_WITHOUT_CACHED_VALUE']) {
    const r = run([sale({ lineTotal, quantity: '' })]);
    assert.equal(r.items[0].status, 'Invalid Sales Amount');
    assert.equal(r.items[0].adjustment, null);
    assert.match(r.items[0].explanation, /Invalid Sales Amount: Net Total at source row 2/);
    assert.doesNotMatch(r.items[0].explanation, /refund|return/i);
  }
});

test('Sheet 2 groups and sorts dates/branches and reconciles every amount to eligible Sheet 1 products', async () => {
  const rows = parse(
    workbook({
      Sales: [
        [
          'Product ID',
          'Date',
          'Branch',
          'Invoice Number',
          'Product Name',
          'Qty',
          'Price',
          'Unit Cost',
          'Discount',
          'Net Total',
          'Payment',
        ],
        ['A', '2026-10-20', 'GAME ONE SM NORTH', '1', 'Next day', 1, 100, 80, 0, 100, 'BDO 3 MOS'],
        [
          'B',
          '10/19/2026',
          'GAME ONE TRINOMA',
          '1',
          'Other branch',
          1,
          200,
          150,
          0,
          200,
          'BPI STRAIGHT',
        ],
        [
          'C',
          '2026-10-19',
          'GAME ONE SM NORTH',
          '1',
          'Two products',
          2,
          100,
          60,
          20,
          180,
          'BDO 3 MOS',
        ],
        [
          'C',
          '2026-10-19',
          'GAME ONE SM NORTH',
          '1',
          'Repeated ID',
          1,
          20,
          10,
          0,
          20,
          'BPI STRAIGHT',
        ],
        ['D', '2026-10-19', 'GAME ONE SM NORTH', '2', 'Cash', 1, 500, 400, 0, 500, 'Cash'],
        [
          'E',
          '2026-10-19',
          'GAME ONE SM NORTH',
          '3',
          'Unknown',
          1,
          500,
          400,
          0,
          500,
          'Unknown method',
        ],
      ],
    }),
  );
  const r = run(rows);
  assert.deepEqual(r.branchDateSummary, [
    {
      date: '2026-10-19',
      branch: 'GAME ONE SM NORTH',
      sales: '200.00',
      totalCost: '130.00',
      adjustment: '6.90',
      grossMargin: '63.10',
    },
    {
      date: '2026-10-19',
      branch: 'GAME ONE TRINOMA',
      sales: '200.00',
      totalCost: '150.00',
      adjustment: '6.00',
      grossMargin: '44.00',
    },
    {
      date: '2026-10-20',
      branch: 'GAME ONE SM NORTH',
      sales: '100.00',
      totalCost: '80.00',
      adjustment: '3.50',
      grossMargin: '16.50',
    },
  ]);
  const b = new ExcelJS.Workbook();
  await b.xlsx.load((await exporter.report(r)) as any);
  const main = b.worksheets[0],
    summary = b.worksheets[1];
  assert.equal(main.columnCount, 17);
  assert.equal(main.rowCount, 5, 'Only the four configured bank-payment products are exported');
  assert.equal(main.getCell('H4').value, 120, 'Sheet 1 displays total cost for quantity 2');
  assert.equal(summary.name, 'Gross Margin by Branch & Date');
  assert.equal(summary.columnCount, 6);
  assert.equal(summary.rowCount, 6, 'Title, headers, three groups and totals');
  assert.equal(summary.getCell('A1').value, 'Expected Gross Margin Per Branch by Date');
  assert.deepEqual(summary.model.merges, ['A1:F1']);
  assert.deepEqual((summary.getRow(2).values as string[]).slice(1), [
    'Date',
    'Branch',
    'Total Credit Card Payments',
    'Total Cost / COGS',
    'Total Bank Charges / Subsidy',
    'Gross Margin',
  ]);
  assert.equal(summary.autoFilter, 'A2:F5');
  assert.equal(summary.views[0].state, 'frozen');
  assert.equal((summary.views[0] as ExcelJS.WorksheetViewFrozen).ySplit, 2);
  const grandTotals = [0, 0, 0, 0].map((n) => new Decimal(n));
  r.branchDateSummary.forEach((group, index) => {
    const data = summary.getRow(index + 3);
    const products = r.items.flatMap((item, i) =>
      i < 4 && item.salesDate === group.date && item.branch === group.branch
        ? [main.getRow(i + 2)]
        : [],
    );
    const totals = [0, 0, 0, 0].map((n) => new Decimal(n));
    for (const product of products) {
      totals[0] = totals[0].plus(String(product.getCell(10).value));
      totals[1] = totals[1].plus(String(product.getCell(8).value));
      totals[2] = totals[2].plus(String(product.getCell(13).value));
      totals[3] = totals[3].plus(String(product.getCell(14).value));
    }
    assert.equal((data.getCell(1).value as Date).toISOString(), group.date + 'T00:00:00.000Z');
    assert.equal(data.getCell(1).numFmt, 'dd mmm yyyy');
    assert.equal(data.getCell(2).value, group.branch);
    totals.forEach((total, i) => {
      assert.equal(data.getCell(i + 3).value, total.toNumber());
      assert.match(data.getCell(i + 3).numFmt, /₱.*#,##0\.00/);
      grandTotals[i] = grandTotals[i].plus(total);
    });
  });
  grandTotals.forEach((total, i) =>
    assert.equal(summary.lastRow!.getCell(i + 3).value, total.toNumber()),
  );
  assert.deepEqual(exporter.import(upload(Buffer.from(await b.xlsx.writeBuffer()))), c());
});
test('Sheet 2 excludes cash, unconfigured, partial and incomplete items while retaining the source audit', () => {
  const r = run([
    sale(),
    sale({ payment: 'Cash' }),
    sale({ payment: 'SBC QRPH' }),
    sale({ payment: 'BANK TRANSFER' }),
    sale({ payment: 'BDO 3 MOS PROMO' }),
    sale({ payment: '' }),
    sale({ cost: '' }),
    sale({ cost: 'invalid' }),
    sale({ quantity: '2', costBasis: 'unknown' }),
    sale({ date: '', issues: ['Missing or invalid Sales Date.'] }),
    sale({ branch: '  ' }),
    sale({
      payments: [part('BDO 3 MOS', '20000'), part('UNKNOWN', '10000')],
      paymentScope: 'line',
    }),
    sale({ payments: [part('BDO 3 MOS', '20000')], paymentScope: 'line' }),
    sale({ productId: '', warnings: ['Missing Product ID.'] }),
  ]);
  assert.equal(r.items.length, 14);
  assert.equal(r.items[6].adjustment, '1050.00', 'Missing cost retains its Sheet 1 fee');
  assert.equal(r.items[11].grossMargin, '4300.00', 'Partial margin remains visible on Sheet 1');
  assert.ok(r.items[11].partialCalculation);
  assert.ok(r.reviews.some((review) => /Missing Cost/.test(review.reason)));
  assert.ok(r.reviews.some((review) => /Unknown payment/.test(review.reason)));
  assert.deepEqual(r.branchDateSummary, [
    {
      date: '2026-10-01',
      branch: 'ASUS',
      sales: '60000.00',
      totalCost: '50000.00',
      adjustment: '2100.00',
      grossMargin: '7900.00',
    },
  ]);
});
test('Sheet 2 uses configured aliases, patterns, installment spacing, debit and explicitly configured transfers', () => {
  const cfg = c();
  cfg.paymentMethods.push({ bank: 'BDO', cardName: 'Custom Transfer', rate: '1.25' });
  cfg.aliases.push({ alias: 'Transfer alias', cardName: 'Custom Transfer' });
  cfg.patterns.push({ pattern: 'Approved transfer *', cardName: 'Custom Transfer' });
  const rows = [
    'Transfer alias',
    'Approved transfer REF123',
    'BDO 3MOS',
    'BDO DEBIT',
    'Unconfigured Transfer',
  ].map((payment) => sale({ payment, lineTotal: '100', cost: '80' }));
  const r = calc.calculate(rows, config.validate(cfg));
  assert.deepEqual(r.branchDateSummary, [
    {
      date: '2026-10-01',
      branch: 'ASUS',
      sales: '400.00',
      totalCost: '320.00',
      adjustment: '8.35',
      grossMargin: '71.65',
    },
  ]);
  assert.equal(r.items[4].adjustment, null);
  assert.deepEqual(run(rows.slice(0, 2)).branchDateSummary, [], 'Transfers require configuration');
  assert.deepEqual(
    calc.calculate([sale()], config.validate({ ...c(), paymentMethods: [], aliases: [] }))
      .branchDateSummary,
    [],
  );
  for (const rate of ['', 'invalid', '101']) {
    const invalid = c();
    invalid.paymentMethods[0].rate = rate;
    assert.throws(() => config.validate(invalid), /rate|percentage/i);
  }
});
test('Sheet 2 sums signed subsidies, zero rates and existing margins without a second deduction', () => {
  const cfg = c();
  cfg.paymentMethods.find((m) => m.cardName === 'BDO 3 MOS')!.rate = '-2.00';
  cfg.paymentMethods.find((m) => m.cardName === 'BPI STRAIGHT')!.rate = '0';
  const r = calc.calculate(
    [
      sale({ lineTotal: '100', cost: '80' }),
      sale({ payment: 'BPI STRAIGHT', lineTotal: '100', cost: '80' }),
      sale({ payment: 'BDO DEBIT', lineTotal: '100', cost: '80' }),
    ],
    config.validate(cfg),
  );
  assert.deepEqual(r.branchDateSummary, [
    {
      date: '2026-10-01',
      branch: 'ASUS',
      sales: '300.00',
      totalCost: '240.00',
      adjustment: '0.35',
      grossMargin: '59.65',
    },
  ]);
  const subsidy = calc.calculate([sale({ lineTotal: '100', cost: '80' })], cfg);
  assert.equal(subsidy.branchDateSummary[0].adjustment, '-2.00');
  assert.equal(subsidy.branchDateSummary[0].grossMargin, '22.00');
});
test('Sheet 2 sums each split-payment product once, preserving rounded fees, gift costs and full margins', () => {
  const payments = [part('BDO 3 MOS', '0.51'), part('BPI STRAIGHT', '0.49'), part('Cash', '0.01')];
  const r = run([
    sale({ payments, lineTotal: '0.51', cost: '0.10' }),
    sale({ payments, lineTotal: '0.50', cost: '0.20' }),
    sale({ payments, lineTotal: '0.00', cost: '0.25' }),
  ]);
  assert.equal(r.items.length, 3);
  assert.equal(r.items[2].grossMargin, '-0.25');
  assert.deepEqual(r.branchDateSummary, [
    {
      date: '2026-10-01',
      branch: 'ASUS',
      sales: '1.01',
      totalCost: '0.55',
      adjustment: '0.03',
      grossMargin: '0.43',
    },
  ]);
  assert.equal(
    r.branchDateSummary[0].grossMargin,
    Decimal.sum(...r.items.map((item) => item.grossMargin!)).toFixed(2),
  );
});
test('Sheet 2 uses net sales currency parsing and existing branch normalization', () => {
  const r = run([
    sale({ lineTotal: 'PHP 30,000.00' }),
    sale({ branch: ' asus ', lineTotal: '₱30,000.00' }),
  ]);
  assert.equal(r.branchDateSummary.length, 1);
  assert.equal(r.branchDateSummary[0].sales, '60000.00');
  assert.equal(r.branchDateSummary[0].totalCost, '50000.00');
});
test('Empty Sheet 2 retains its heading, six columns and zero totals', async () => {
  const r = run([sale({ payment: 'Cash' }), sale({ payment: 'Unknown' })]);
  assert.deepEqual(r.branchDateSummary, []);
  const b = new ExcelJS.Workbook();
  await b.xlsx.load((await exporter.report(r)) as any);
  assert.equal(b.worksheets[0].rowCount, 1, 'No matched items leaves only Sheet 1 headers');
  assert.equal(b.worksheets[0].columnCount, 17);
  assert.equal(b.getWorksheet('Needs Review')!.rowCount, 2, 'Unknown payment remains reviewable');
  const ws = b.getWorksheet(BRANCH_DATE_SUMMARY_SHEET)!;
  assert.equal(ws.columnCount, 6);
  assert.equal(ws.rowCount, 3);
  assert.equal(ws.getCell('A1').value, BRANCH_DATE_SUMMARY_TITLE);
  assert.deepEqual(
    (ws.getRow(2).values as string[]).slice(1),
    BRANCH_DATE_SUMMARY_COLUMNS.map((c) => c.header),
  );
  assert.equal(ws.autoFilter, 'A2:F2');
  for (const col of [3, 4, 5, 6]) assert.equal(ws.getCell(3, col).value, 0);
});
test('Only configured Tender items appear in Sheet 1, preserving source order, reviews and totals', async () => {
  const cfg = c();
  cfg.aliases.push({ alias: 'Terminal A', cardName: 'BDO 3 MOS' });
  cfg.patterns.push({ pattern: 'Terminal B *', cardName: 'BPI 3 MOS' });
  cfg.paymentMethods.push(
    { bank: 'BDO', cardName: 'Zero fee card', rate: '0' },
    { bank: 'BPI', cardName: 'Subsidy card', rate: '-2' },
  );
  const variants: Partial<Sale>[] = [
    { payment: 'Cash', productId: '=OMITTED-CASH' },
    { payment: 'BNPL BDO 3 MOS' },
    { payment: '' },
    { payment: '', payments: [part('BPI 3 MOS SNPL', '30000')] },
    { payment: 'Unconfigured Bank' },
    { payment: 'Terminal A' },
    { payment: 'Terminal B REF123' },
    { payment: 'Zero fee card' },
    { payment: 'Subsidy card' },
    { payment: 'BDO 3 MOS', cost: '' },
    { payment: 'SBC QRPH' },
    { payments: [part('BDO 3 MOS', '20000'), part('Unknown', '10000')] },
    { payment: 'BDO 3 MOS', lineTotal: 'invalid' },
  ];
  const rows = variants.map((variant, i) =>
    sale({
      productId: 'ITEM-' + i,
      invoice: 'INV-' + i,
      source: { file: 'sales.xlsx', sheet: 'Sales Report', row: i + 2 },
      ...variant,
    }),
  );
  const r = calc.calculate(rows, config.validate(cfg));
  const expectedIndices = [1, 3, 5, 6, 7, 8, 9, 11, 12];
  const expected = expectedIndices.map((i) => r.items[i]);
  const snapshot = structuredClone(r);
  const b = new ExcelJS.Workbook();
  await b.xlsx.load((await exporter.report(r)) as any);
  const main = b.worksheets[0];
  assert.equal(r.items.length, variants.length, 'Source audit retains all processed items');
  assert.equal(r.summary.processedItems, variants.length);
  assert.equal(main.rowCount, expected.length + 1);
  assert.equal(main.columnCount, 17);
  for (const [i, item] of expected.entries()) {
    const row = main.getRow(i + 2);
    assert.equal(row.getCell(1).value, item.productId);
    assert.equal(row.getCell(11).value, item.originalPaymentMethod);
    assert.equal(row.getCell(15).value, item.sourceRow);
    assert.equal(row.getCell(16).value, item.cardName);
    assert.equal(row.getCell(13).value, item.adjustment === null ? null : Number(item.adjustment));
    assert.equal(row.getCell(8).value, item.totalCost === null ? null : Number(item.totalCost));
    assert.match(
      String(row.getCell(17).value),
      i < 6 ? /Included in Sheet 2:/ : /Excluded from Sheet 2:/,
    );
  }
  assert.deepEqual(
    (b.getWorksheet('Needs Review')!.getColumn(1).values as string[]).slice(2),
    ['ITEM-2', 'ITEM-4', 'ITEM-9', 'ITEM-11', 'ITEM-12'],
    'Missing/unconfigured payment items are still available for review',
  );
  const complete = expected.slice(0, 6);
  for (const key of ['sales', 'totalCost', 'adjustment', 'grossMargin'] as const) {
    assert.equal(
      r.branchDateSummary[0][key],
      complete
        .reduce(
          (total, item) => total.plus(key === 'sales' ? item.lineTotal : item[key]!),
          new Decimal(0),
        )
        .toFixed(2),
    );
  }
  assert.deepEqual(r, snapshot, 'Report filtering does not change calculations or configuration');
});

test('Exact 17-column four-sheet export, Product ID in A as text, dates, values and style', async () => {
  const r = run(parse(reportFixture())),
    b = new ExcelJS.Workbook();
  await b.xlsx.load((await exporter.report(r)) as any);
  assert.deepEqual(
    b.worksheets.map((s) => s.name),
    ['Item Bank Deductions', BRANCH_DATE_SUMMARY_SHEET, 'Configuration Settings', 'Needs Review'],
  );
  const ws = b.worksheets[0];
  assert.equal(ws.columnCount, 17);
  assert.deepEqual(
    (ws.getRow(1).values as string[]).slice(1),
    ITEM_COLUMNS.map((c) => c.header),
  );
  assert.equal(ws.rowCount, 2);
  assert.equal(ws.getCell('A2').value, '000123');
  assert.equal(ws.getCell('A2').numFmt, '@');
  assert.equal(ws.getCell('A2').type, ExcelJS.ValueType.String);
  assert.equal(ws.getCell('M2').value, 1050);
  assert.equal(ws.getCell('N2').value, 3950);
  assert.equal(ws.getCell('O2').value, 2);
  assert.equal((ws.getCell('B2').value as Date).toISOString(), '2026-10-01T00:00:00.000Z');
  assert.equal(ws.views[0].state, 'frozen');
  assert.ok(ws.autoFilter);
});
test('Missing cost stays blank in export; source identifiers and formula-like text are inert', async () => {
  const r = run([sale({ productId: '=SKU', cost: '', product: '=HYPERLINK(\"bad\")' })]),
    b = new ExcelJS.Workbook();
  await b.xlsx.load((await exporter.report(r)) as any);
  assert.equal(b.worksheets[0].getCell('A2').value, '=SKU');
  assert.equal(b.worksheets[0].getCell('A2').type, ExcelJS.ValueType.String);
  assert.equal(typeof b.worksheets[0].getCell('E2').value, 'string');
  assert.equal(b.worksheets[0].getCell('N2').value, null);
  assert.equal(b.worksheets[0].getCell('M2').value, 1050);
});
test('Edited configuration and negative subsidy rates round-trip from Excel', async () => {
  const cfg = c();
  cfg.aliases = [{ alias: 'Terminal A', cardName: 'BDO 3 MOS' }];
  const b = new ExcelJS.Workbook();
  await b.xlsx.load((await exporter.configuration(cfg)) as any);
  const ws = b.worksheets[0],
    row = cfg.paymentMethods.findIndex((m) => m.cardName === 'BDO 3 MOS') + 2;
  ws.getCell(row, 2).value = -0.025;
  const imported = exporter.import(upload(Buffer.from(await b.xlsx.writeBuffer())));
  assert.equal(imported.paymentMethods.find((m) => m.cardName === 'BDO 3 MOS')!.rate, '-2.50');
  assert.deepEqual(imported.aliases, cfg.aliases);
  ws.getCell(row, 2).value = { formula: '1+1', result: 2 };
  const formulaBuffer = Buffer.from(await b.xlsx.writeBuffer());
  assert.throws(() => exporter.import(upload(formulaBuffer)), /formula/i);
});
test('Header preambles and manual Product ID mapping preserve actual Excel rows', () => {
  const f = upload(
    workbook({
      Data: [
        ['Report'],
        [],
        ['Custom SKU', 'Invoice Number', 'Product Name', 'Quantity', 'Net Total'],
        ['00001', 'A1', 'Mouse', 1, 100],
      ],
    }),
  );
  const rows = excel.parse(f, {
    sheet: 'Data',
    headerRow: 3,
    mapping: {
      productId: 'Custom SKU',
      invoice: 'Invoice Number',
      product: 'Product Name',
      quantity: 'Quantity',
      lineTotal: 'Net Total',
    },
  });
  assert.equal(rows[0].productId, '00001');
  assert.equal(rows[0].source.row, 4);
});

test('Sheet 1 copies Payment verbatim, labels configured names Tender and omits Bank and Status', async () => {
  const original = '=SOURCE\n  BDO 3MOS  ';
  const r = run([sale({ payment: 'BDO 3MOS', original: { payment: original } })]);
  const b = new ExcelJS.Workbook();
  await b.xlsx.load((await exporter.report(r)) as any);
  const ws = b.worksheets[0],
    headers = (ws.getRow(1).values as string[]).slice(1);
  assert.deepEqual(headers, [
    'Product ID',
    'Sales Date',
    'Branch',
    'Invoice Number',
    'Product Name',
    'Quantity',
    'Price',
    'Cost',
    'Discount',
    'Line Total',
    'Payment',
    'Applied Bank Rate',
    'Bank Charges / Subsidy',
    'GrossMargin',
    'Sales Source Row',
    'Tender',
    'Calculation Explanation',
  ]);
  assert.equal(ws.getCell('K2').value, original);
  assert.equal(ws.getCell('K2').type, ExcelJS.ValueType.String);
  assert.equal(ws.getCell('P2').value, 'BDO 3 MOS');
  assert.equal(ws.columnCount, 17);
});

test('Four-column settings preserve custom bank links, aliases, patterns and non-card methods', async () => {
  const cfg = c();
  cfg.banks.push('CUSTOM BANK');
  cfg.paymentMethods.push({ bank: 'CUSTOM BANK', cardName: 'My terminal', rate: '1.75' });
  cfg.aliases = [{ alias: 'Terminal alias', cardName: 'My terminal' }];
  cfg.patterns = [{ pattern: 'Terminal *', cardName: 'My terminal' }];
  cfg.nonCardMethods.push('Voucher');
  const b = new ExcelJS.Workbook();
  await b.xlsx.load((await exporter.configuration(cfg)) as any);
  const ws = b.worksheets[0];
  assert.equal(ws.columnCount, 4);
  assert.deepEqual((ws.getRow(1).values as string[]).slice(1), [
    'Card Name',
    'Rate',
    'Bank Reference',
    'Branch',
  ]);
  assert.deepEqual(exporter.import(upload(Buffer.from(await b.xlsx.writeBuffer()))), cfg);
  ws.getCell(cfg.paymentMethods.length + 1, 1).value = 'Renamed terminal';
  ws.getCell(cfg.banks.length + 1, 3).value = 'RENAMED BANK';
  const imported = exporter.import(upload(Buffer.from(await b.xlsx.writeBuffer())));
  assert.equal(imported.paymentMethods.at(-1)!.bank, 'RENAMED BANK');
  assert.equal(imported.paymentMethods.at(-1)!.cardName, 'Renamed terminal');
  assert.equal(imported.aliases[0].cardName, 'Renamed terminal');
  assert.equal(imported.patterns[0].cardName, 'Renamed terminal');
});

test('Four-column settings without notes never infer bank relationships from row positions', () => {
  const buffer = workbook({
    'Configuration Settings': [
      ['Card Name', 'Rate', 'Bank Reference', 'Branch'],
      ['BDO 3 MOS', 0.035, 'BPI', 'ASUS'],
      ['BPI 12 MOS', 0.125, 'BDO', 'SMCEBU'],
    ],
  });
  const cfg = exporter.import(upload(buffer));
  assert.equal(cfg.paymentMethods[0].bank, 'BDO');
  assert.equal(cfg.paymentMethods[1].bank, 'BPI');
  const unknown = workbook({
    'Configuration Settings': [
      ['Card Name', 'Rate', 'Bank Reference', 'Branch'],
      ['Unknown Terminal', 0.035, 'BDO', 'ASUS'],
    ],
  });
  assert.throws(() => exporter.import(upload(unknown)), /bank assignment/);
});

test('Previously exported ten-column configuration remains importable', () => {
  const buffer = workbook({
    'Configuration Settings': [
      [
        'Card Name',
        'Bank',
        'Rate',
        'Bank Reference',
        'Branch',
        'Payment Method Alias',
        'Alias Card Name',
        'Pattern',
        'Pattern Card Name',
        'Non-card Method',
      ],
      [
        'My terminal',
        'CUSTOM',
        0.035,
        'CUSTOM',
        'ASUS',
        'Alias',
        'My terminal',
        'Terminal *',
        'My terminal',
        'Cash',
      ],
    ],
  });
  const imported = exporter.import(upload(buffer));
  assert.equal(imported.paymentMethods[0].bank, 'CUSTOM');
  assert.equal(imported.paymentMethods[0].rate, '3.50');
  assert.equal(imported.aliases[0].alias, 'Alias');
});
test('One-file HTTP workflow previews, calculates, downloads and imports its configuration', async () => {
  const app = await createApplication(true);
  await app.init();
  try {
    const server = app.getHttpServer(),
      buf = reportFixture();
    const defaults = await request(server).get('/api/config/defaults').expect(200);
    assert.deepEqual(defaults.body, config.defaults());
    assert.equal(defaults.body.paymentMethods.length, 49);
    const preview = await request(server)
      .post('/api/excel/preview')
      .attach('file', buf, 'sales.xlsx')
      .expect(201);
    const s = preview.body.sheets[0];
    const response = await request(server)
      .post('/api/excel/calculate')
      .field(
        'options',
        JSON.stringify({
          sales: { sheet: s.name, headerRow: s.headerRow, mapping: s.suggestedMapping },
          configuration: c(),
        }),
      )
      .attach('sales', buf, 'sales.xlsx')
      .expect(201);
    assert.equal(response.body.result.items.length, 2);
    assert.equal(response.body.result.items[0].productId, '000123');
    const download = await request(server)
      .post('/api/excel/export')
      .send(response.body)
      .buffer(true)
      .parse((res, cb) => {
        const chunks: Buffer[] = [];
        res.on('data', (v) => chunks.push(v));
        res.on('end', () => cb(null, Buffer.concat(chunks)));
      })
      .expect(201);
    assert.match(
      download.headers['content-disposition'],
      /Sales_Bank_Rate_Deductions_\d{8}_\d{6}\.xlsx/,
    );
    const imported = await request(server)
      .post('/api/config/import')
      .attach('file', download.body, 'report.xlsx')
      .expect(201);
    assert.deepEqual(imported.body, c());
    response.body.result.items[0].adjustment = '0';
    await request(server).post('/api/excel/export').send(response.body).expect(400);
    await request(server)
      .post('/api/excel/preview')
      .attach('file', Buffer.alloc(10 * 1024 * 1024 + 1), 'oversize.xlsx')
      .expect(413);
    await request(server)
      .post('/api/excel/calculate')
      .field(
        'options',
        JSON.stringify({
          sales: { sheet: s.name, headerRow: 1, mapping: s.suggestedMapping },
          configuration: c(),
        }),
      )
      .expect(400);
  } finally {
    await app.close();
  }
});
async function main() {
  let passed = 0;
  for (const t of tests) {
    try {
      await t.run();
      passed++;
      console.log('PASS ' + t.name);
    } catch (e) {
      console.error('FAIL ' + t.name, e);
      process.exitCode = 1;
    }
  }
  console.log(passed + '/' + tests.length + ' checks passed.');
}
void main();
