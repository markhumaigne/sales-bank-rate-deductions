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
import { matchRate } from '../src/calculation/rate-matcher';
import { reportDate } from '../src/excel/report-date';
import { ITEM_COLUMNS, Sale } from '../../shared/models';
import { sale, part, upload, workbook, reportFixture } from './helpers';
const config = new ConfigService(),
  excel = new ExcelService(),
  calc = new CalculationService(),
  exporter = new ExcelExportService(config, excel);
const c = () => config.defaults();
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
test('Cash without a configured card retains its source row and leaves all three calculations blank', async () => {
  const r = run([sale({ payment: 'Cash' })]);
  const item = r.items[0];
  assert.equal(item.status, 'Not Calculated');
  assert.equal(item.cardName, '');
  assert.equal(item.originalPaymentMethod, 'Cash');
  assert.equal(item.lineTotal, '30000');
  assert.equal(item.appliedRate, null);
  assert.equal(item.adjustment, null);
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
  assert.equal(ws.getCell('K2').value, 'Cash');
  assert.equal(ws.getCell('P2').value, '');
  for (const address of ['L2', 'M2', 'N2']) assert.equal(ws.getCell(address).value, null);
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
    assert.equal(item.adjustment, null);
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
  assert.equal(r.items[1].adjustment, null);
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

test('SBC QRPH-only Payment stays unchanged on Sheet 1 with blank rate, charge and margin', async () => {
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
    assert.equal(ws.getCell('K2').value, payment);
    assert.equal(ws.getCell('P2').value, '');
    for (const address of ['L2', 'M2', 'N2']) assert.equal(ws.getCell(address).value, null);
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

test('Matched split calculations use each known rate while leaving unconfigured promotions under review', () => {
  const r = run([
    sale({ payment: 'BDO 3MOS 10,000.00\nBPI 12 MOS 10,000.00\nBDO 6 MOS BNPL 10,000.00' }),
  ]);
  const item = r.items[0];
  assert.equal(item.adjustment, '1600.00');
  assert.equal(item.grossMargin, '3400.00');
  assert.equal(item.cardName, 'BDO 3 MOS / BPI 12 MOS');
  assert.match(item.appliedRate!, /BDO 3 MOS: 3.50%/);
  assert.match(item.appliedRate!, /BPI 12 MOS: 12.50%/);
  assert.equal(item.partialCalculation, true);
  assert.match(r.reviews[0].reason, /BNPL/);
  assert.equal(r.summary.grossMargin, '0.00');
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
    'BDO 6 MOS BNPL',
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
test('Unconfigured BNPL/SNPL methods require review even with description suffixes', () => {
  for (const payment of ['BDO 6 MOS BNPL', 'METROBANK 6 MOS SNPL']) {
    const r = run([sale({ payment, payments: [part(payment + ' CUSTOMER', '30000')] })]);
    assert.equal(r.items[0].adjustment, null);
    assert.match(r.items[0].explanation, /BNPL/);
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
test('Explicit card/cash split applies fees only to card portion', () => {
  const r = run([sale({ payments: [part('BDO 3 MOS', '20000'), part('Cash', '10000')] })]);
  assert.equal(r.items.length, 1);
  assert.equal(r.items[0].adjustment, '700.00');
  assert.equal(r.items[0].grossMargin, '4300.00');
  assert.equal(r.summary.sales, '30000.00');
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
test('Unit cost and total line cost are interpreted explicitly', () => {
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
test('Exact 17-column four-sheet export, Product ID in A as text, dates, values and style', async () => {
  const r = run(parse(reportFixture())),
    b = new ExcelJS.Workbook();
  await b.xlsx.load((await exporter.report(r)) as any);
  assert.deepEqual(
    b.worksheets.map((s) => s.name),
    ['Item Bank Deductions', 'Bank & Branch Summary', 'Configuration Settings', 'Needs Review'],
  );
  const ws = b.worksheets[0];
  assert.equal(ws.columnCount, 17);
  assert.deepEqual(
    (ws.getRow(1).values as string[]).slice(1),
    ITEM_COLUMNS.map((c) => c.header),
  );
  assert.equal(ws.rowCount, 3);
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
