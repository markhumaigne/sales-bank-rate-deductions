import 'reflect-metadata';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { basename } from 'node:path';
import request from 'supertest';
import Decimal from 'decimal.js';
import ExcelJS from 'exceljs';
import { createApplication } from '../src/create-app';
import { ConfigService } from '../src/config/config.service';
import { ExcelService } from '../src/excel/excel.service';
import { ExcelExportService } from '../src/excel/excel-export.service';
import { CalculationService } from '../src/calculation/calculation.service';
import {
  ITEM_COLUMNS,
  BRANCH_DATE_SUMMARY_COLUMNS,
  BRANCH_DATE_SUMMARY_SHEET,
  BRANCH_DATE_SUMMARY_TITLE,
} from '../../shared/models';
import { upload } from './helpers';

async function main() {
  const args = process.argv.slice(2),
    at = args.indexOf('--sales');
  if (at < 0 || !args[at + 1])
    throw new Error('Usage: npm run test:sample -- --sales "path/to/Sales Report.xlsx"');
  const path = args[at + 1],
    file = upload(readFileSync(path), basename(path));
  const excel = new ExcelService(),
    config = new ConfigService(),
    calc = new CalculationService();
  const exporter = new ExcelExportService(config, excel);
  const preview = excel.preview(file),
    s =
      preview.sheets.find((s) => s.name === 'Sales Report') ||
      preview.sheets.find((s) => s.suggestedMapping.product)!;
  assert.ok(s);
  const selection = { sheet: s.name, headerRow: s.headerRow, mapping: s.suggestedMapping };
  const sales = excel.parse(file, selection),
    result = calc.calculate(sales, config.defaults());
  assert.equal(result.items.length, sales.length);
  assert.equal(result.summary.processedItems, sales.length);
  assert.deepEqual(
    result.items.map((i) => i.sourceRow),
    sales.map((s) => s.source.row),
  );
  assert.ok(result.summary.feeCalculatedItems > 0);
  const screenshotRow = sales.find(
    (row) =>
      row.source.row === 3 &&
      row.payment.includes('BDO STRAIGHT') &&
      row.payment.includes('SBC QRPH'),
  );
  if (screenshotRow) {
    const item = result.items.find((item) => item.sourceRow === screenshotRow.source.row)!;
    assert.equal(
      item.cardName,
      'BDO STRAIGHT',
      'Row 3 calculates BDO STRAIGHT while excluding SBC QRPH from bank charges',
    );
    assert.equal(item.appliedRate, '2.85');
    assert.equal(item.adjustment, '0.00');
    assert.equal(item.grossMargin, '-238.00');
    assert.equal(item.originalPaymentMethod, screenshotRow.payment);
  }
  const cashRow = result.items.find((item) => item.sourceRow === 40);
  if (cashRow?.originalPaymentMethod === 'Cash 27,100.00') {
    assert.equal(cashRow.cardName, '');
    assert.equal(cashRow.appliedRate, null);
    assert.equal(cashRow.adjustment, '0.00');
    assert.equal(cashRow.grossMargin, null);
    assert.doesNotMatch(cashRow.explanation, /Gross Margin =/);
  }
  const sum = (values: string[]) => values.reduce((a, b) => a.plus(b), new Decimal(0)).toFixed(2);
  assert.equal(
    sum(result.items.filter((i) => !i.partialCalculation).map((i) => i.adjustment ?? '0')),
    new Decimal(result.summary.charges).minus(result.summary.subsidies).toFixed(2),
  );
  assert.equal(
    sum(result.items.filter((i) => !i.partialCalculation).map((i) => i.grossMargin ?? '0')),
    result.summary.grossMargin,
  );
  assert.equal(
    sum(
      result.items
        .filter((i) => i.cardName && i.adjustment !== null && !i.partialCalculation)
        .map((i) => i.lineTotal),
    ),
    result.summary.sales,
  );
  for (const [index, item] of result.items.entries()) {
    assert.equal(item.productId, sales[index].productId);
    assert.equal(item.originalPaymentMethod, sales[index].payment);
    if (!item.cardName) {
      assert.equal(item.appliedRate, null, 'No Tender means no calculated rate');
      assert.ok(
        item.adjustment === null || item.adjustment === '0.00',
        'No Tender means no bank charge',
      );
      assert.equal(item.grossMargin, null, 'No Tender means no calculated margin');
    }
    if (item.grossMargin !== null)
      assert.equal(
        item.grossMargin,
        new Decimal(item.lineTotal).minus(item.totalCost!).minus(item.adjustment!).toFixed(2),
      );
    if (item.adjustment === null) {
      assert.equal(item.grossMargin, null);
      if (item.status !== 'Not Calculated' || item.warnings.length)
        assert.ok(result.reviews.some((r) => r.sourceRow === item.sourceRow));
      else {
        assert.equal(item.issues.length, 0);
        assert.equal(item.cardName, '');
        assert.ok(!result.reviews.some((r) => r.sourceRow === item.sourceRow));
      }
    }
    if (item.partialCalculation) {
      assert.notEqual(item.appliedRate, null);
      assert.notEqual(item.adjustment, null);
      assert.ok(result.reviews.some((r) => r.sourceRow === item.sourceRow));
      assert.match(item.explanation, /provisional/i);
    }
  }
  // Independent ExcelJS reader verifies every original source cell, not merely our parser's output.
  if (/\.xlsx$/i.test(path)) {
    const original = new ExcelJS.Workbook();
    await original.xlsx.load(file.buffer as any);
    const ws = original.getWorksheet(s.name)!;
    let rows = 0;
    ws.eachRow((r, n) => {
      if (n > s.headerRow && r.hasValues) rows++;
    });
    assert.equal(rows, sales.length, 'Every original source row retained');
    const value = (address: string) => {
      const c = ws.getCell(address);
      return c.type === ExcelJS.ValueType.Formula ? c.result : c.value;
    };
    for (const [index, row] of sales.entries()) {
      const item = result.items[index];
      for (const key of ['productId', 'invoice', 'product', 'branch', 'payment'] as const) {
        if (row.cells[key]) {
          const raw = value(row.cells[key]);
          if (key === 'productId' && typeof raw === 'number') {
            // The reference uses '# ?/?': integer IDs have presentation-only trailing spaces.
            assert.equal(new Decimal(row.productId.trim()).toString(), String(raw));
          } else
            assert.equal(
              String(raw ?? ''),
              row[key],
              'Original ' + key + ' at source row ' + row.source.row,
            );
        }
      }
      for (const key of ['quantity', 'price', 'cost', 'discount', 'lineTotal'] as const) {
        if (row.cells[key] && value(row.cells[key]) !== null)
          assert.equal(
            new Decimal(String(value(row.cells[key]))).toFixed(2),
            new Decimal(row[key]).toFixed(2),
            'Original ' + key + ' at source row ' + row.source.row,
          );
      }
      if (row.cells.date) {
        const raw = value(row.cells.date);
        assert.ok(raw instanceof Date);
        assert.equal(raw.toISOString().slice(0, 10), item.salesDate);
      }
      if (item.totalCost !== null) {
        const cost = new Decimal(String(value(row.cells.cost))),
          qty = new Decimal(String(value(row.cells.quantity)));
        assert.equal(
          item.totalCost,
          (row.costBasis === 'line' ? cost : cost.times(qty))
            .toDecimalPlaces(2, Decimal.ROUND_HALF_UP)
            .toFixed(2),
        );
      }
    }
    // Independently verify aggregate payment fees for each resolved invoice using literal source groups.
    const groups = new Map<string, number[]>();
    sales.forEach((row, index) => {
      const key = JSON.stringify([row.invoice, row.branch, row.date]);
      const g = groups.get(key) || [];
      g.push(index);
      groups.set(key, g);
    });
    let verifiedInvoices = 0;
    for (const indices of groups.values()) {
      if (indices.some((i) => result.items[i].adjustment === null)) continue;
      const row = sales[indices[0]];
      if (!row.payments.length) continue;
      let independentFee = new Decimal(0);
      for (const part of row.payments) {
        const literal = part.method
          .toUpperCase()
          .replace(/\b(?:BNPL|SNPL)\b/g, ' ')
          .replace(/\s+/g, ' ')
          .trim();
        if (literal === 'CASH' || literal === 'CASH PAYMENT' || literal === 'SBC QRPH') continue;
        const matched = result.configuration.paymentMethods
          .flatMap((method) => [
            { name: method.cardName, method },
            ...result.configuration.aliases
              .filter((alias) => alias.cardName === method.cardName)
              .map((alias) => ({ name: alias.alias, method })),
          ])
          .filter(
            (m) =>
              literal === m.name.toUpperCase() || literal.startsWith(m.name.toUpperCase() + ' '),
          )
          .sort((a, b) => b.name.length - a.name.length)[0]?.method;
        if (!matched || /\b(PROMO|PROMOTIONAL)\b/.test(literal)) {
          assert.ok(
            indices.every((i) => result.items[i].partialCalculation),
            'Unconfigured parts only occur in partial calculations',
          );
          continue;
        }
        independentFee = independentFee.plus(
          new Decimal(part.amount)
            .times(matched.rate)
            .div(100)
            .toDecimalPlaces(2, Decimal.ROUND_HALF_UP),
        );
      }
      assert.equal(
        sum(indices.map((i) => result.items[i].adjustment!)),
        independentFee.toFixed(2),
        'Independent invoice payment fee',
      );
      verifiedInvoices++;
    }
    assert.ok(verifiedInvoices > 0);
    console.log(
      'Independently verified source cells and payment fees for ' +
        verifiedInvoices +
        ' invoices with fully or partially calculated payments.',
    );
  }
  const buffer = await exporter.report(result),
    book = new ExcelJS.Workbook();
  await book.xlsx.load(buffer as any);
  assert.deepEqual(
    book.worksheets.map((w) => w.name),
    ['Item Bank Deductions', BRANCH_DATE_SUMMARY_SHEET, 'Configuration Settings', 'Needs Review'],
  );
  const summary = book.worksheets[1];
  assert.equal(summary.columnCount, 6);
  assert.equal(summary.getCell('A1').value, BRANCH_DATE_SUMMARY_TITLE);
  assert.deepEqual(
    (summary.getRow(2).values as string[]).slice(1),
    BRANCH_DATE_SUMMARY_COLUMNS.map((c) => c.header),
  );
  const eligible = result.items.filter(
    (i) =>
      i.adjustment !== null &&
      i.grossMargin !== null &&
      !i.partialCalculation &&
      i.salesDate &&
      i.branch.trim(),
  );
  for (const key of ['sales', 'totalCost', 'adjustment', 'grossMargin'] as const) {
    assert.equal(
      sum(result.branchDateSummary.map((g) => g[key])),
      sum(eligible.map((i) => (key === 'sales' ? i.lineTotal : i[key]!))),
    );
  }
  for (const [index, group] of result.branchDateSummary.entries()) {
    const items = eligible.filter(
      (i) =>
        i.salesDate === group.date &&
        i.branch.replace(/\s+/g, ' ').trim().toUpperCase() === group.branch.toUpperCase(),
    );
    assert.equal(group.sales, sum(items.map((i) => i.lineTotal)));
    assert.equal(group.totalCost, sum(items.map((i) => i.totalCost!)));
    assert.equal(group.adjustment, sum(items.map((i) => i.adjustment!)));
    assert.equal(group.grossMargin, sum(items.map((i) => i.grossMargin!)));
    assert.deepEqual((summary.getRow(index + 3).values as unknown[]).slice(1), [
      new Date(group.date + 'T00:00:00.000Z'),
      group.branch,
      Number(group.sales),
      Number(group.totalCost),
      Number(group.adjustment),
      Number(group.grossMargin),
    ]);
  }
  const main = book.worksheets[0];
  assert.equal(main.columnCount, 17);
  assert.deepEqual(
    (main.getRow(1).values as string[]).slice(1),
    ITEM_COLUMNS.map((c) => c.header),
  );
  const exportedItems = result.items.filter(
    (item) => item.cardName.trim() && !(Number(item.quantity) < 0),
  );
  assert.equal(main.rowCount, exportedItems.length + 1);
  for (const [index, item] of exportedItems.entries()) {
    assert.equal(main.getCell(index + 2, 1).value, item.productId || '');
    assert.equal(main.getCell(index + 2, 1).type, ExcelJS.ValueType.String);
    assert.equal(main.getCell(index + 2, 5).value, item.product);
    assert.equal(
      main.getCell(index + 2, 8).value,
      item.totalCost === null ? null : Number(item.totalCost),
      'Sheet 1 displays the same line cost used by gross margin and Sheet 2',
    );
    assert.equal(main.getCell(index + 2, 11).value, item.originalPaymentMethod);
    assert.equal(
      main.getCell(index + 2, 13).value,
      item.adjustment === null ? null : Number(item.adjustment),
    );
    assert.equal(
      main.getCell(index + 2, 14).value,
      item.grossMargin === null ? null : Number(item.grossMargin),
    );
    assert.equal(main.getCell(index + 2, 15).value, item.sourceRow);
    assert.equal(main.getCell(index + 2, 16).value, item.cardName);
  }
  assert.deepEqual(exporter.import(upload(buffer)), result.configuration);
  const app = await createApplication(true);
  await app.init();
  try {
    const response = await request(app.getHttpServer())
      .post('/api/excel/calculate')
      .field('options', JSON.stringify({ sales: selection, configuration: config.defaults() }))
      .attach('sales', file.buffer, file.originalname)
      .expect(201);
    assert.deepEqual(response.body.result.summary, result.summary);
    await request(app.getHttpServer())
      .post('/api/excel/export')
      .send(response.body)
      .expect(201)
      .expect('Content-Type', /spreadsheetml/);
  } finally {
    await app.close();
  }
  const statuses: Record<string, number> = {};
  result.items.forEach((i) => (statuses[i.status] = (statuses[i.status] || 0) + 1));
  console.log(
    JSON.stringify(
      {
        sourceRows: sales.length,
        outputRows: result.items.length,
        sourceDates: result.reportDates,
        costBasis: s.suggestedCostBasis,
        columns: main.columnCount,
        summary: result.summary,
        statuses,
        sheets: book.worksheets.map((w) => w.name),
        sourceAndExportChecks: 'PASS',
        httpWorkflow: 'PASS',
        transactionDataSaved: false,
      },
      null,
      2,
    ),
  );
}
main().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
