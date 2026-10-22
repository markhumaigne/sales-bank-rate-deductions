import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import { ExcelService } from '../src/excel/excel.service';
import { ExcelExportService } from '../src/excel/excel-export.service';
import { CalculationService } from '../src/calculation/calculation.service';
import { ConfigService } from '../src/config/config.service';
import { reportFixture, upload, workbook } from './helpers';

export const sourceIntegrityTests: { name: string; run: () => unknown | Promise<unknown> }[] = [];
const test = (name: string, run: () => unknown | Promise<unknown>) =>
  sourceIntegrityTests.push({ name, run });
const excel = new ExcelService(),
  config = new ConfigService(),
  calculator = new CalculationService();
const exporter = new ExcelExportService(config, excel);
const parse = (buffer: Buffer) => {
  const file = upload(buffer),
    preview = excel.preview(file).sheets[0];
  return excel.parse(file, {
    sheet: preview.name,
    headerRow: preview.headerRow,
    mapping: preview.suggestedMapping,
  });
};
const headers = [
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
];

test('Numeric noise is normalized only for numeric cells; fractional cents remain unresolved', async () => {
  for (const [value, expected] of [
    [30000.000000000004, '975.00'],
    [30000.001, null],
    ['30000.000000000004', null],
    [{ formula: '30000' }, null],
  ] as const) {
    const book = new ExcelJS.Workbook();
    await book.xlsx.load(reportFixture() as any);
    book.worksheets[0].getCell('J2').value = value;
    const sales = parse(Buffer.from(await book.xlsx.writeBuffer()));
    const result = calculator.calculate(sales, config.defaults());
    assert.equal(result.items[0].adjustment, expected);
    if (expected) {
      assert.equal(sales[0].lineTotal, '30000.00');
      assert.equal(sales[0].original.lineTotal, String(value));
      assert.match(result.items[0].explanation, /J2: Excel numeric storage noise/);
    } else {
      assert.equal(result.items[0].grossMargin, null);
      assert.deepEqual(result.branchDateSummary, []);
      assert.match(result.reviews[0].reason, /Invalid Sales Amount/);
    }
  }
});

test('1904 dates, zero-padded identifiers and cached formula values retain their source meaning', async () => {
  const book = new ExcelJS.Workbook();
  await book.xlsx.load(reportFixture() as any);
  book.properties.date1904 = true;
  const sheet = book.worksheets[0];
  sheet.getCell('A2').value = 123;
  sheet.getCell('A2').numFmt = '000000';
  sheet.getCell('D2').value = 7;
  sheet.getCell('D2').numFmt = '000000';
  sheet.getCell('B2').value = new Date('2026-10-01T00:00:00.000Z');
  sheet.getCell('B2').numFmt = 'mm/dd/yyyy';
  sheet.getCell('J2').value = { formula: '15000*2', result: 30000 };
  const sales = parse(Buffer.from(await book.xlsx.writeBuffer()));
  const result = calculator.calculate(sales, config.defaults());
  assert.equal(result.items[0].productId, '000123');
  assert.equal(result.items[0].invoice, '000007');
  assert.equal(result.items[0].salesDate, '2026-10-01');
  assert.equal(result.items[0].lineTotal, '30000');
  assert.equal(result.items[0].adjustment, '975.00');
  assert.match(result.items[0].explanation, /J2: cached formula value read/);
  const exported = new ExcelJS.Workbook();
  await exported.xlsx.load((await exporter.report(result)) as any);
  assert.equal(exported.worksheets[0].getCell('A2').value, '000123');
  assert.equal(exported.worksheets[0].getCell('D2').value, '000007');
  assert.equal(
    (exported.worksheets[0].getCell('B2').value as Date).toISOString(),
    '2026-10-01T00:00:00.000Z',
  );
  assert.equal(exported.worksheets[0].getCell('J2').value, 30000);
  assert.equal(exported.worksheets[0].getCell('J2').type, ExcelJS.ValueType.Number);
});

test('Duplicate financial and branch headers require an explicit column choice', () => {
  const file = upload(
    workbook({
      Sales: [
        [...headers, ' Branch ', 'NET TOTAL'],
        [
          '0001',
          '10/01/2026',
          'SM CALAMBA',
          'INV1',
          'Phone',
          2,
          50,
          25,
          0,
          999,
          'UNIONBANK 3 MOS 100.00',
          'SM SEASIDE',
          100,
        ],
      ],
    }),
  );
  const preview = excel.preview(file).sheets[0];
  assert.deepEqual(preview.ambiguousFields, ['branch', 'lineTotal']);
  assert.equal(preview.suggestedMapping.branch, undefined);
  assert.equal(preview.suggestedMapping.lineTotal, undefined);
  const selection = { sheet: preview.name, headerRow: 1, mapping: preview.suggestedMapping };
  assert.throws(() => excel.parse(file, selection), /lineTotal/);
  const mapping = { ...selection.mapping, lineTotal: preview.headers[12] };
  assert.throws(() => excel.parse(file, { ...selection, mapping }), /Duplicate.*branch/);
  const sales = excel.parse(file, {
    ...selection,
    mapping: { ...mapping, branch: preview.headers[11] },
  });
  const result = calculator.calculate(sales, config.defaults());
  assert.equal(sales[0].branch, 'SM SEASIDE');
  assert.equal(sales[0].cells.branch, 'L2');
  assert.equal(sales[0].cells.lineTotal, 'M2');
  assert.equal(result.items[0].lineTotal, '100');
  assert.equal(result.items[0].totalCost, '50.00');
  assert.equal(result.items[0].adjustment, '3.50');
  assert.equal(result.items[0].grossMargin, '46.50');
});

test('A source amount cannot silently serve as both cost and sales or payment', () => {
  const file = upload(reportFixture()),
    preview = excel.preview(file).sheets[0];
  for (const field of ['cost', 'price', 'paymentAmount']) {
    assert.throws(
      () =>
        excel.parse(file, {
          sheet: preview.name,
          headerRow: 1,
          mapping: { ...preview.suggestedMapping, [field]: preview.suggestedMapping.lineTotal },
        }),
      /separate columns/,
    );
  }
});

test('Branch Excel errors and missing formula results cannot receive bank charges or branch totals', async () => {
  for (const value of [
    { error: '#N/A' },
    { formula: 'NA()', result: { error: '#N/A' } },
    { formula: '"SM SEASIDE"' },
  ] as ExcelJS.CellValue[]) {
    const book = new ExcelJS.Workbook();
    await book.xlsx.load(reportFixture() as any);
    book.worksheets[0].getCell('C2').value = value;
    const sales = parse(Buffer.from(await book.xlsx.writeBuffer()));
    const result = calculator.calculate(sales, config.defaults());
    assert.equal(result.items[0].adjustment, null);
    assert.equal(result.items[0].grossMargin, null);
    assert.deepEqual(result.branchDateSummary, []);
    assert.match(result.reviews.find((r) => r.sourceRow === 2)!.reason, /branch/i);
    if ('error' in (value as object) || 'result' in (value as object)) {
      assert.equal(sales[0].original.branch, '#N/A');
      assert.match(sales[0].issues.join(' '), /branch at C2.*Excel error/);
    }
  }
});

test('All report text stays exact and inert in item, summary and review sheets', async () => {
  const book = new ExcelJS.Workbook();
  await book.xlsx.load(reportFixture() as any);
  const sheet = book.worksheets[0];
  const expected = [
    ['=SKU', '  @BRANCH\t', '=00001', '=HYPERLINK("example.invalid","literal")'],
    ["'SKU", '+BRANCH', '-00002', "'Product name"],
  ];
  expected.forEach((values, index) => {
    [1, 3, 4, 5].forEach((column, i) => (sheet.getCell(index + 2, column).value = values[i]));
  });
  const sales = parse(Buffer.from(await book.xlsx.writeBuffer()));
  const result = calculator.calculate(sales, config.defaults());
  const exported = new ExcelJS.Workbook();
  await exported.xlsx.load((await exporter.report(result)) as any);
  for (const [sheetName, index] of [
    ['Item Bank Deductions', 0],
    ['Needs Review', 1],
  ] as const) {
    [1, 3, 4, 5].forEach((column, i) => {
      const cell = exported.getWorksheet(sheetName)!.getCell(2, column);
      assert.equal(cell.value, expected[index][i]);
      assert.equal(cell.type, ExcelJS.ValueType.String);
    });
  }
  assert.equal(exported.worksheets[1].getCell('B3').value, '@BRANCH');
  assert.equal(exported.worksheets[1].getCell('B3').type, ExcelJS.ValueType.String);
  const reviewedOriginal = JSON.parse(
    String(exported.getWorksheet('Needs Review')!.getCell('J2').value),
  );
  assert.deepEqual(reviewedOriginal, sales[1].original);
  for (const worksheet of exported.worksheets)
    worksheet.eachRow((row) =>
      row.eachCell((cell) => assert.notEqual(cell.type, ExcelJS.ValueType.Formula)),
    );
});

test('Blank gaps, hidden rows and duplicate products remain tied to their own source cells', async () => {
  const book = new ExcelJS.Workbook();
  const sheet = book.addWorksheet('Sales');
  sheet.addRow(headers);
  sheet.addRow([
    '0007',
    '10/01/2026',
    'SM SEASIDE',
    'INV1',
    'Phone',
    1,
    100,
    60,
    0,
    100,
    'BDO STRAIGHT 100.00',
  ]);
  sheet.addRow([]);
  sheet.addRow([
    '0007',
    '10/01/2026',
    'SM CALAMBA',
    'INV1',
    'Phone',
    1,
    200,
    150,
    0,
    200,
    'BDO STRAIGHT 200.00',
  ]);
  sheet.getRow(4).hidden = true;
  sheet.addRow([
    '0007',
    '10/02/2026',
    'SM SEASIDE',
    'INV1',
    'Phone',
    1,
    300,
    200,
    0,
    300,
    'Cash 300.00',
  ]);
  book.addWorksheet('Unselected').addRow(['Do not mix other worksheets into this selection']);
  const buffer = Buffer.from(await book.xlsx.writeBuffer());
  const preview = excel.preview(upload(buffer)).sheets[0];
  const sales = parse(buffer);
  assert.equal(preview.recordCount, 3);
  assert.deepEqual(
    sales.map((row) => row.source.row),
    [2, 4, 5],
  );
  const result = calculator.calculate(sales, config.defaults());
  assert.deepEqual(
    result.items.map((item) => [item.sourceRow, item.branch, item.lineTotal, item.adjustment]),
    [
      [2, 'SM SEASIDE', '100', '2.85'],
      [4, 'SM CALAMBA', '200', '5.70'],
      [5, 'SM SEASIDE', '300', null],
    ],
  );
  const exported = new ExcelJS.Workbook();
  await exported.xlsx.load((await exporter.report(result)) as any);
  assert.equal(exported.worksheets[0].getCell('Q2').value, 2);
  assert.equal(exported.worksheets[0].getCell('Q3').value, 4);
  assert.equal(exported.getWorksheet('Needs Review')!.getCell('F2').value, 5);
});
