import { test, expect, Download } from '@playwright/test';
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import {
  ITEM_COLUMNS,
  BRANCH_DATE_SUMMARY_COLUMNS,
  BRANCH_DATE_SUMMARY_TITLE,
  BRANCH_DATE_SUMMARY_SHEET,
} from '../shared/models';
function fixture() {
  const b = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    b,
    XLSX.utils.aoa_to_sheet([['Notes'], ['Select the Sales Report worksheet.']]),
    'Instructions',
  );
  XLSX.utils.book_append_sheet(
    b,
    XLSX.utils.aoa_to_sheet([
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
      ['000123', 46296, 'ASUS', '0001', 'Laptop', 1, 30000, 25000, 0, 30000, 'BDO 3 MOS'],
      ['000124', 46296, 'ASUS', '0002', 'Mouse', 2, 500, 300, 100, 900, 'Cash'],
      ['000125', 46296, 'ASUS', '0003', 'Keyboard', 1, 2000, 1200, 0, 2000, 'BPI 3 MOS'],
    ]),
    'Sales Report',
  );
  return {
    name: 'sales-fixture.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: XLSX.write(b, { type: 'buffer', bookType: 'xlsx' }) as Buffer,
  };
}
async function downloadBuffer(d: Download) {
  const stream = await d.createReadStream();
  if (!stream) throw new Error('Download stream missing.');
  const chunks: Buffer[] = [];
  for await (const chunk of stream) chunks.push(Buffer.from(chunk));
  return Buffer.concat(chunks);
}
test('Sales-only upload, exact 17 columns, search/sort, download and configuration round trip', async ({
  page,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (e) => errors.push(e.message));
  await page.goto('/');
  await expect(
    page.getByRole('heading', { name: 'Bank rate deduction calculator', exact: true }),
  ).toBeVisible();
  const calculate = page.getByRole('button', { name: 'Calculate Bank Rate Deductions' });
  await expect(calculate).toBeDisabled();
  await expect(page.getByLabel('Upload Tender Report', { exact: true })).toHaveCount(0);
  await expect(page.getByRole('tab', { name: /Payment methods \(49\)/ })).toBeVisible();
  await page.screenshot({ path: 'test-results/dashboard-desktop.png', fullPage: true });
  await page.getByLabel('Upload Sales Report', { exact: true }).setInputFiles(fixture());
  await page.getByLabel('Worksheet', { exact: true }).selectOption('Sales Report');
  await expect(calculate).toBeEnabled();
  const responsePromise = page.waitForResponse((r) => r.url().endsWith('/api/excel/calculate'));
  await calculate.click();
  const response = await responsePromise;
  expect(response.ok()).toBeTruthy();
  const payload = await response.json();
  expect(payload.result.summary.charges).toBe('1035.00');
  expect(payload.result.summary.grossMargin).toBe('4765.00');
  expect(payload.result.summary.sales).toBe('32000.00');
  expect(payload.result.items[1].appliedRate).toBeNull();
  expect(payload.result.items[1].adjustment).toBe('0.00');
  expect(payload.result.items[1].grossMargin).toBeNull();
  const table = page.getByRole('table', { name: 'Calculated Sales Report' });
  await expect(
    page.getByRole('tab', { name: 'Item Bank Deductions (2)', exact: true }),
  ).toBeVisible();
  await expect(table.locator('tbody tr')).toHaveCount(2);
  await expect(table.getByRole('cell', { name: '000124', exact: true })).toHaveCount(0);
  await expect(table.locator('thead th')).toHaveCount(17);
  expect(
    (await table.locator('thead th').allTextContents()).map((s) => s.replace(/[↑↓]/g, '').trim()),
  ).toEqual(ITEM_COLUMNS.map((c) => c.header));
  await expect(table.getByRole('cell', { name: '000123', exact: true })).toBeVisible();
  await table.getByRole('button', { name: 'Product ID', exact: true }).click();
  await table.getByRole('button', { name: /^Product ID/ }).click();
  await expect(table.locator('tbody tr').first().locator('td').first()).toHaveText('000125');
  await expect(table.locator('tbody tr').first().locator('td').nth(4)).toHaveText('Keyboard');
  await expect(table.locator('tbody tr').first().locator('td').nth(12)).toHaveText('₱60.00');
  await page.getByLabel('Search results').fill('000123');
  await expect(table.locator('tbody tr')).toHaveCount(1);
  await expect(table.getByRole('cell', { name: 'Laptop', exact: true })).toBeVisible();
  await page.getByLabel('Search results').clear();
  await page.getByRole('tab', { name: BRANCH_DATE_SUMMARY_SHEET, exact: true }).click();
  await expect(
    page.getByRole('heading', { name: BRANCH_DATE_SUMMARY_TITLE, exact: true }),
  ).toBeVisible();
  await expect(table.locator('thead th')).toHaveCount(6);
  expect(
    (await table.locator('thead th').allTextContents()).map((s) => s.replace(/[↑↓]/g, '').trim()),
  ).toEqual(BRANCH_DATE_SUMMARY_COLUMNS.map((c) => c.header));
  await expect(table.locator('tbody tr')).toHaveCount(1);
  await expect(table.locator('tbody tr').first().locator('td')).toHaveText([
    '01 Oct 2026',
    'ASUS',
    '₱32,000.00',
    '₱26,200.00',
    '₱1,035.00',
    '₱4,765.00',
  ]);
  const downloaded = page.waitForEvent('download');
  await page.getByRole('button', { name: /Download Calculated Excel Report/ }).click();
  const download = await downloaded;
  expect(download.suggestedFilename()).toMatch(/^Sales_Bank_Rate_Deductions_\d{8}_\d{6}\.xlsx$/);
  const report = new ExcelJS.Workbook();
  await report.xlsx.load((await downloadBuffer(download)) as any);
  expect(report.worksheets.map((s) => s.name)).toEqual([
    'Item Bank Deductions',
    BRANCH_DATE_SUMMARY_SHEET,
    'Configuration Settings',
    'Needs Review',
  ]);
  const ws = report.worksheets[0];
  expect(ws.columnCount).toBe(17);
  expect(ws.rowCount).toBe(3);
  expect(ws.getCell('A2').value).toBe('000123');
  expect(ws.getCell('M2').value).toBe(975);
  expect(ws.getCell('N2').value).toBe(4025);
  expect(ws.getCell('A3').value).toBe('000125');
  expect(ws.getCell('M3').value).toBe(60);
  expect(ws.getCell('N3').value).toBe(740);
  expect(ws.getCell('O3').value).toBe(4);
  expect(ws.getCell('P3').value).toBe('BPI 3 MOS');
  expect(ws.getCell('A2').numFmt).toBe('@');
  const summary = report.worksheets[1];
  expect(summary.columnCount).toBe(6);
  expect(summary.getCell('A1').value).toBe(BRANCH_DATE_SUMMARY_TITLE);
  expect((summary.getRow(2).values as string[]).slice(1)).toEqual(
    BRANCH_DATE_SUMMARY_COLUMNS.map((c) => c.header),
  );
  expect((summary.getRow(3).values as unknown[]).slice(1)).toEqual([
    new Date('2026-10-01T00:00:00.000Z'),
    'ASUS',
    32000,
    26200,
    1035,
    4765,
  ]);
  await page.getByLabel('Search payment methods').fill('BDO 3 MOS');
  await page.getByLabel('Rate percentage', { exact: true }).fill('4.00');
  await expect(
    page.getByRole('heading', { name: 'Your calculated report will appear here' }),
  ).toBeVisible();
  await calculate.click();
  await expect(page.getByText('₱1,200.00', { exact: true }).first()).toBeVisible();
  await page.getByRole('tab', { name: BRANCH_DATE_SUMMARY_SHEET, exact: true }).click();
  await expect(table.locator('tbody tr').first().locator('td').nth(4)).toHaveText('₱1,260.00');
  await expect(table.locator('tbody tr').first().locator('td').nth(5)).toHaveText('₱4,540.00');
  const exported = page.waitForEvent('download');
  await page.getByRole('button', { name: /Export settings/ }).click();
  const configBook = new ExcelJS.Workbook();
  await configBook.xlsx.load((await downloadBuffer(await exported)) as any);
  let edited = false;
  configBook.worksheets[0].eachRow((row) => {
    if (row.getCell(1).value === 'BDO 3 MOS') {
      row.getCell(2).value = 0.045;
      edited = true;
    }
  });
  expect(edited).toBeTruthy();
  await page.getByLabel('Import configuration workbook').setInputFiles({
    name: 'edited-settings.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: Buffer.from(await configBook.xlsx.writeBuffer()),
  });
  await expect(
    page.getByText('Configuration imported. Review the settings before calculating.'),
  ).toBeVisible();
  await calculate.click();
  await expect(page.getByText('₱1,350.00', { exact: true }).first()).toBeVisible();
  expect(errors).toEqual([]);
});
test('Only configured bank-payment items appear in the preview and exported worksheets', async ({
  page,
}) => {
  const f = fixture(),
    source = XLSX.read(f.buffer, { type: 'buffer' });
  const headers = XLSX.utils.sheet_to_json<unknown[]>(source.Sheets['Sales Report'], {
    header: 1,
  })[0];
  const payments = [
    ['BDO Straight BNPL', 100, 80],
    ['SNPL BPI 3 MOS', 200, 150],
    ['BNPL BDO 6 MOS', 300, 200],
    ['Cash', 100, 80],
    ['Cash Payment', 100, 80],
    ['Unconfigured Bank SNPL', 400, 300],
    ['', 100, 80],
    ['SBC QRPH', 100, 80],
  ] as const;
  source.Sheets['Sales Report'] = XLSX.utils.aoa_to_sheet([
    headers,
    ...payments.map(([payment, sales, cost], i) => [
      'ITEM-' + i,
      '2026-10-19',
      'ASUS',
      'INV-' + i,
      'Product ' + i,
      1,
      sales,
      cost,
      0,
      sales,
      payment,
    ]),
  ]);
  f.buffer = XLSX.write(source, { type: 'buffer', bookType: 'xlsx' });
  await page.goto('/');
  await page.getByLabel('Upload Sales Report', { exact: true }).setInputFiles(f);
  await page.getByLabel('Worksheet', { exact: true }).selectOption('Sales Report');
  const calculated = page.waitForResponse((r) => r.url().endsWith('/api/excel/calculate'));
  await page.getByRole('button', { name: 'Calculate Bank Rate Deductions' }).click();
  const response = await calculated;
  expect(response.ok()).toBeTruthy();
  const { result } = await response.json();
  expect(result.items.map((i: { adjustment: string | null }) => i.adjustment)).toEqual([
    '2.85',
    '6.00',
    '16.50',
    '0.00',
    '0.00',
    null,
    null,
    null,
  ]);
  expect(result.summary.feeCalculatedItems).toBe(3);
  expect(result.reviews).toHaveLength(2);
  const table = page.getByRole('table', { name: 'Calculated Sales Report' });
  await expect(table.locator('tbody tr')).toHaveCount(3);
  await expect(
    page.getByRole('tab', { name: 'Item Bank Deductions (3)', exact: true }),
  ).toBeVisible();
  await expect(table.locator('tbody tr').first().locator('td').nth(10)).toHaveText(
    'BDO Straight BNPL',
  );
  await expect(table.locator('tbody tr').first().locator('td').nth(15)).toHaveText('BDO STRAIGHT');
  for (const omitted of ['ITEM-3', 'ITEM-4', 'ITEM-5', 'ITEM-6', 'ITEM-7'])
    await expect(table.getByRole('cell', { name: omitted, exact: true })).toHaveCount(0);
  await page.getByRole('tab', { name: 'Needs Review (2)', exact: true }).click();
  await expect(table.getByRole('cell', { name: 'ITEM-5', exact: true })).toBeVisible();
  await expect(table.getByRole('cell', { name: 'ITEM-6', exact: true })).toBeVisible();
  await page.getByRole('tab', { name: BRANCH_DATE_SUMMARY_SHEET, exact: true }).click();
  await expect(table.locator('thead th')).toHaveCount(6);
  await expect(table.locator('tbody tr')).toHaveCount(1);
  await expect(table.locator('tbody tr').first().locator('td')).toHaveText([
    '19 Oct 2026',
    'ASUS',
    '₱600.00',
    '₱430.00',
    '₱25.35',
    '₱144.65',
  ]);
  const downloaded = page.waitForEvent('download');
  await page.getByRole('button', { name: /Download Calculated Excel Report/ }).click();
  const report = new ExcelJS.Workbook();
  await report.xlsx.load((await downloadBuffer(await downloaded)) as any);
  const main = report.worksheets[0],
    summary = report.worksheets[1];
  expect(main.rowCount).toBe(4);
  expect(main.columnCount).toBe(17);
  for (const [i, [payment]] of payments.slice(0, 3).entries()) {
    expect(main.getCell(i + 2, 11).value).toBe(payment);
    expect(main.getCell(i + 2, 13).value).toBe(Number(result.items[i].adjustment));
  }
  expect(report.getWorksheet('Needs Review')!.rowCount).toBe(3);
  expect(summary.columnCount).toBe(6);
  expect((summary.getRow(3).values as unknown[]).slice(3)).toEqual([600, 430, 25.35, 144.65]);
});

test('Sheet 1 Cost and Sheet 2 COGS use the same quantity-adjusted cost in preview and export', async ({
  page,
}) => {
  const f = fixture(),
    b = XLSX.read(f.buffer, { type: 'buffer' }),
    source = b.Sheets['Sales Report'];
  source.F2.v = 2;
  source.J2.v = 60000;
  source.F4.v = 3;
  source.J4.v = 6000;
  f.buffer = XLSX.write(b, { type: 'buffer', bookType: 'xlsx' });
  await page.goto('/');
  await page.getByLabel('Upload Sales Report', { exact: true }).setInputFiles(f);
  await page.getByLabel('Worksheet', { exact: true }).selectOption('Sales Report');
  const calculate = page.getByRole('button', { name: 'Calculate Bank Rate Deductions' });
  await calculate.click();
  const table = page.getByRole('table', { name: 'Calculated Sales Report' });
  await expect(table.locator('tbody tr')).toHaveCount(2);
  await expect(table.locator('tbody tr').first().locator('td').nth(7)).toHaveText('₱50,000.00');
  await expect(table.locator('tbody tr').last().locator('td').nth(7)).toHaveText('₱3,600.00');
  await expect(table.locator('tbody tr').first().locator('td').nth(16)).toContainText(
    'Source Cost "25000" × quantity "2"',
  );
  await expect(table.locator('tbody tr').first().locator('td').nth(16)).toContainText(
    'Included in Sheet 2:',
  );
  await page.getByRole('tab', { name: BRANCH_DATE_SUMMARY_SHEET, exact: true }).click();
  await expect(table.locator('tbody tr').first().locator('td')).toHaveText([
    '01 Oct 2026',
    'ASUS',
    '₱66,000.00',
    '₱53,600.00',
    '₱2,130.00',
    '₱10,270.00',
  ]);
  const downloaded = page.waitForEvent('download');
  await page.getByRole('button', { name: /Download Calculated Excel Report/ }).click();
  const report = new ExcelJS.Workbook();
  await report.xlsx.load((await downloadBuffer(await downloaded)) as any);
  const items = report.worksheets[0],
    summary = report.worksheets[1];
  expect(items.getCell('H2').value).toBe(50000);
  expect(items.getCell('H3').value).toBe(3600);
  expect(summary.getCell('D3').value).toBe(
    Number(items.getCell('H2').value) + Number(items.getCell('H3').value),
  );
  expect(summary.lastRow!.getCell('D').value).toBe(53600);
  await page.getByLabel('Cost basis', { exact: true }).selectOption('line');
  await calculate.click();
  await page.getByRole('tab', { name: 'Item Bank Deductions (2)', exact: true }).click();
  await expect(table.locator('tbody tr').first().locator('td').nth(7)).toHaveText('₱25,000.00');
  await expect(table.locator('tbody tr').last().locator('td').nth(7)).toHaveText('₱1,200.00');
  await page.getByRole('tab', { name: BRANCH_DATE_SUMMARY_SHEET, exact: true }).click();
  await expect(table.locator('tbody tr').first().locator('td').nth(3)).toHaveText('₱26,200.00');
  await expect(table.locator('tbody tr').first().locator('td').nth(5)).toHaveText('₱37,670.00');
});

test('Negative quantities move to review and contribute no sales or cost to either worksheet', async ({
  page,
}) => {
  const f = fixture(),
    b = XLSX.read(f.buffer, { type: 'buffer' }),
    source = b.Sheets['Sales Report'];
  source.F2 = { t: 'n', v: -1 };
  source.J2 = { t: 'n', v: 30000 };
  source.K2 = { t: 's', v: 'BDO 6 MOS BNPL' };
  f.buffer = XLSX.write(b, { type: 'buffer', bookType: 'xlsx' });
  await page.goto('/');
  await page.getByLabel('Upload Sales Report', { exact: true }).setInputFiles(f);
  await page.getByLabel('Worksheet', { exact: true }).selectOption('Sales Report');
  await page.getByRole('button', { name: 'Calculate Bank Rate Deductions' }).click();
  const table = page.getByRole('table', { name: 'Calculated Sales Report' });
  await expect(
    page.getByRole('tab', { name: 'Item Bank Deductions (1)', exact: true }),
  ).toBeVisible();
  await expect(table.locator('tbody tr')).toHaveCount(1);
  await expect(table.getByRole('cell', { name: '000123', exact: true })).toHaveCount(0);
  await expect(table.getByRole('cell', { name: '000125', exact: true })).toBeVisible();
  await page.getByRole('tab', { name: 'Needs Review (1)', exact: true }).click();
  await expect(table.getByRole('cell', { name: '000123', exact: true })).toBeVisible();
  await expect(table.locator('tbody tr').first().locator('td').last()).toContainText(
    'Quantity at cell F2 is "-1"',
  );
  await page.getByRole('tab', { name: BRANCH_DATE_SUMMARY_SHEET, exact: true }).click();
  await expect(table.locator('tbody tr')).toHaveCount(1);
  await expect(table.locator('tbody tr').first().locator('td')).toHaveText([
    '01 Oct 2026',
    'ASUS',
    '₱2,000.00',
    '₱1,200.00',
    '₱60.00',
    '₱740.00',
  ]);
  const downloaded = page.waitForEvent('download');
  await page.getByRole('button', { name: /Download Calculated Excel Report/ }).click();
  const report = new ExcelJS.Workbook();
  await report.xlsx.load((await downloadBuffer(await downloaded)) as any);
  expect(report.worksheets[0].rowCount).toBe(2);
  expect(report.worksheets[0].getCell('A2').value).toBe('000125');
  const review = report.getWorksheet('Needs Review')!;
  expect(review.rowCount).toBe(2);
  expect(review.getCell('A2').value).toBe('000123');
  expect(JSON.parse(String(review.getCell('J2').value))).toMatchObject({
    quantity: '-1',
    lineTotal: '30000',
    price: '30000',
    cost: '25000',
  });
  expect((report.worksheets[1].getRow(3).values as unknown[]).slice(3)).toEqual([
    2000, 1200, 60, 740,
  ]);
});

test('Requested warnings stay hidden and negative quantities remain available in review', async ({
  page,
}) => {
  const f = fixture(),
    b = XLSX.read(f.buffer, { type: 'buffer' });
  const headers = XLSX.utils.sheet_to_json<unknown[]>(b.Sheets['Sales Report'], { header: 1 })[0];
  b.Sheets['Sales Report'] = XLSX.utils.aoa_to_sheet([
    headers,
    [
      'PARTIAL',
      '2026-10-01',
      'ASUS',
      '1',
      'Split item',
      1,
      10000,
      8000,
      0,
      10000,
      'BDO 3 MOS 9,500.00\nCharge 500.00',
    ],
    [
      'DISCREPANCY',
      '2026-10-01',
      'ASUS',
      '2',
      'Review item',
      1,
      10000,
      8000,
      200,
      9800,
      'BDO 6 MOS 10,000.00',
    ],
    ['NEGATIVE', '2026-10-01', 'ASUS', '3', 'Return item', -1, 100, 80, 0, -100, 'BDO 6 MOS BNPL'],
    ['VALID', '2026-10-01', 'ASUS', '4', 'Valid item', 1, 1000, 800, 0, 1000, 'BDO 6 MOS'],
  ]);
  f.buffer = XLSX.write(b, { type: 'buffer', bookType: 'xlsx' });
  await page.goto('/');
  await page.getByLabel('Upload Sales Report', { exact: true }).setInputFiles(f);
  await page.getByLabel('Worksheet', { exact: true }).selectOption('Sales Report');
  await page.getByRole('button', { name: 'Calculate Bank Rate Deductions' }).click();
  const table = page.getByRole('table', { name: 'Calculated Sales Report' });
  const hidden =
    /"Charge": Unknown payment method or missing configured bank rate\.|Invoice payment discrepancy:/;
  await expect(
    page.getByRole('tab', { name: 'Item Bank Deductions (3)', exact: true }),
  ).toBeVisible();
  await expect(table.locator('tbody tr')).toHaveCount(3);
  await expect(table).not.toContainText(hidden);
  for (const id of ['PARTIAL', 'DISCREPANCY', 'VALID'])
    await expect(table.getByRole('cell', { name: id, exact: true })).toBeVisible();
  const partial = table.locator('tbody tr').nth(0);
  await expect(partial.locator('td').nth(10)).toHaveText('BDO 3 MOS 9,500.00 Charge 500.00');
  await expect(partial.locator('td').nth(12)).toHaveText('₱308.75');
  await expect(partial.locator('td').nth(16)).toContainText('Partial calculation');
  await expect(table.locator('tbody tr').nth(1).locator('td').nth(12)).toHaveText('—');
  await expect(table.getByRole('cell', { name: 'NEGATIVE', exact: true })).toHaveCount(0);
  await page.getByRole('tab', { name: 'Needs Review (3)', exact: true }).click();
  await expect(table.locator('tbody tr')).toHaveCount(3);
  await expect(table).not.toContainText(hidden);
  await expect(table.getByRole('cell', { name: 'DISCREPANCY', exact: true })).toBeVisible();
  await expect(table.getByRole('cell', { name: 'NEGATIVE', exact: true })).toBeVisible();
  await expect(table).toContainText('Negative sales / refund');
  await page.getByRole('tab', { name: BRANCH_DATE_SUMMARY_SHEET, exact: true }).click();
  await expect(table.locator('tbody tr')).toHaveCount(1);
  await expect(table.locator('tbody tr').first().locator('td')).toHaveText([
    '01 Oct 2026',
    'ASUS',
    '₱1,000.00',
    '₱800.00',
    '₱55.00',
    '₱145.00',
  ]);
});

test('PAYMAYA with cash change calculates the configured card fee in the preview and export', async ({
  page,
}) => {
  const f = fixture(),
    b = XLSX.read(f.buffer, { type: 'buffer' });
  const headers = XLSX.utils.sheet_to_json<unknown[]>(b.Sheets['Sales Report'], { header: 1 })[0];
  const payment = 'CC#7297 PAYMAYA DEBIT DEBIT SALES 6,000.00\nCash 3,000.00';
  b.Sheets['Sales Report'] = XLSX.utils.aoa_to_sheet([
    [...headers, 'Tender', 'Ref', 'Amount', 'Tender', 'Ref', 'Amount'],
    [
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
      'PAYMAYA DEBIT DEBIT SALES',
      '7297',
      6000,
      'Cash',
      '',
      3000,
    ],
  ]);
  f.buffer = XLSX.write(b, { type: 'buffer', bookType: 'xlsx' });
  await page.goto('/');
  await page.getByLabel('Upload Sales Report', { exact: true }).setInputFiles(f);
  await page.getByLabel('Worksheet', { exact: true }).selectOption('Sales Report');
  await page.getByRole('button', { name: 'Calculate Bank Rate Deductions' }).click();
  const table = page.getByRole('table', { name: 'Calculated Sales Report' });
  await expect(
    page.getByRole('tab', { name: 'Item Bank Deductions (1)', exact: true }),
  ).toBeVisible();
  await expect(table.locator('tbody tr')).toHaveCount(1);
  const row = table.locator('tbody tr').first();
  await expect(row.locator('td').nth(9)).toHaveText('₱8,995.00');
  await expect(row.locator('td').nth(10)).toHaveText(payment.replace('\n', ' '));
  await expect(row.locator('td').nth(11)).toHaveText('2.89%');
  await expect(row.locator('td').nth(12)).toHaveText('₱173.40');
  await expect(row.locator('td').nth(13)).toHaveText('₱1,821.60');
  await expect(row.locator('td').nth(15)).toHaveText('PAYMAYA DEBIT');
  await expect(row.locator('td').nth(16)).toContainText(
    'Cash tender ₱3000.00 includes ₱5.00 change; net cash payment ₱2995.00',
  );
  await expect(row.locator('td').nth(16)).not.toContainText('pending source review');
  await expect(page.getByRole('tab', { name: 'Needs Review (0)', exact: true })).toBeVisible();
  await page.getByRole('tab', { name: BRANCH_DATE_SUMMARY_SHEET, exact: true }).click();
  await expect(table.locator('tbody tr').first().locator('td')).toHaveText([
    '01 Oct 2026',
    'ASUS',
    '₱8,995.00',
    '₱7,000.00',
    '₱173.40',
    '₱1,821.60',
  ]);
  const downloaded = page.waitForEvent('download');
  await page.getByRole('button', { name: /Download Calculated Excel Report/ }).click();
  const report = new ExcelJS.Workbook();
  await report.xlsx.load((await downloadBuffer(await downloaded)) as any);
  expect(report.worksheets[0].getCell('K2').value).toBe(payment);
  expect(report.worksheets[0].getCell('L2').value).toBe(0.0289);
  expect(report.worksheets[0].getCell('M2').value).toBe(173.4);
  expect(report.worksheets[0].getCell('N2').value).toBe(1821.6);
  expect((report.worksheets[1].getRow(3).values as unknown[]).slice(3)).toEqual([
    8995, 7000, 173.4, 1821.6,
  ]);
});

test('Older settings automatically calculate BDO PAY with the configured BDO PAY/CASA rate', async ({
  page,
}) => {
  await page.route('**/api/config/defaults', async (route) => {
    const response = await route.fetch();
    const settings = await response.json();
    settings.aliases = [];
    await route.fulfill({ response, json: settings });
  });
  const f = fixture(),
    b = XLSX.read(f.buffer, { type: 'buffer' });
  const payment = 'CC#7005 BDO STRAIGHT 7005 31,282.00\nBDO PAY 10,000.00';
  b.Sheets['Sales Report'] = XLSX.utils.aoa_to_sheet([
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
    ],
    ...[
      ['710425578649', 'PS5 G. GTA V (R1)', 1250, 1118],
      ['COMPANION', 'Remaining invoice sales', 40032, 32000],
    ].map(([id, product, price, cost]) => [
      id,
      '2026-10-01',
      'GAME ONE MEGAMALL',
      '3074-00001539',
      product,
      1,
      price,
      cost,
      0,
      price,
      payment,
      'BDO STRAIGHT 7005',
      '7005',
      31282,
      'BDO PAY',
      '',
      10000,
    ]),
  ]);
  f.buffer = XLSX.write(b, { type: 'buffer', bookType: 'xlsx' });
  await page.goto('/');
  await page.getByLabel('Upload Sales Report', { exact: true }).setInputFiles(f);
  await page.getByLabel('Worksheet', { exact: true }).selectOption('Sales Report');
  const calculated = page.waitForResponse((r) => r.url().endsWith('/api/excel/calculate'));
  await page.getByRole('button', { name: 'Calculate Bank Rate Deductions' }).click();
  const response = await calculated;
  expect(response.ok()).toBeTruthy();
  const payload = await response.json();
  expect(payload.result.configuration.aliases).toContainEqual({
    alias: 'BDO PAY',
    cardName: 'BDO PAY/CASA',
  });
  expect(payload.result.reviews).toHaveLength(0);
  const table = page.getByRole('table', { name: 'Calculated Sales Report' });
  await expect(table.locator('tbody tr')).toHaveCount(2);
  const row = table.locator('tbody tr').first();
  await expect(row.locator('td').nth(10)).toHaveText(payment.replace('\n', ' '));
  await expect(row.locator('td').nth(11)).toHaveText('BDO STRAIGHT: 2.85%; BDO PAY/CASA: 3.00%');
  await expect(row.locator('td').nth(12)).toHaveText('₱36.08');
  await expect(row.locator('td').nth(13)).toHaveText('₱95.92');
  await expect(row.locator('td').nth(15)).toHaveText('BDO STRAIGHT / BDO PAY/CASA');
  await expect(row.locator('td').nth(16)).not.toContainText('Unconfigured');
  await expect(row.locator('td').nth(16)).not.toContainText('Provisional');
  await page.getByRole('tab', { name: BRANCH_DATE_SUMMARY_SHEET, exact: true }).click();
  await expect(table.locator('tbody tr').first().locator('td')).toHaveText([
    '01 Oct 2026',
    'GAME ONE MEGAMALL',
    '₱41,282.00',
    '₱33,118.00',
    '₱1,191.54',
    '₱6,972.46',
  ]);
  const downloaded = page.waitForEvent('download');
  await page.getByRole('button', { name: /Download Calculated Excel Report/ }).click();
  const report = new ExcelJS.Workbook();
  await report.xlsx.load((await downloadBuffer(await downloaded)) as any);
  expect(report.worksheets[0].getCell('M2').value).toBe(36.08);
  expect(report.worksheets[0].getCell('N2').value).toBe(95.92);
  expect((report.worksheets[1].getRow(3).values as unknown[]).slice(3)).toEqual([
    41282, 33118, 1191.54, 6972.46,
  ]);
  await page.getByLabel('Search payment methods').fill('BDO PAY/CASA');
  await page.getByLabel('Rate percentage', { exact: true }).fill('4.00');
  await page.getByRole('button', { name: 'Calculate Bank Rate Deductions' }).click();
  await page.getByRole('tab', { name: 'Item Bank Deductions (2)', exact: true }).click();
  await expect(row.locator('td').nth(12)).toHaveText('₱39.11');
  await expect(row.locator('td').nth(13)).toHaveText('₱92.89');
});

test('Mobile layout, worksheet selection and manual Product ID mapping', async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto('/');
  await expect(page.getByRole('tab', { name: /Payment methods/ })).toBeVisible();
  expect(
    await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth),
  ).toBeTruthy();
  const f = fixture(),
    b = XLSX.read(f.buffer, { type: 'buffer' });
  b.Sheets['Sales Report']['A1'].v = 'Custom SKU';
  b.Sheets['Sales Report']['J1'].v = 'Final Sales Value';
  f.buffer = XLSX.write(b, { type: 'buffer', bookType: 'xlsx' });
  await page.getByLabel('Upload Sales Report', { exact: true }).setInputFiles(f);
  await page.getByLabel('Worksheet', { exact: true }).selectOption('Sales Report');
  await expect(page.getByText('Map required columns to continue')).toBeVisible();
  await page.getByLabel('Product ID / SKU', { exact: true }).selectOption('Custom SKU');
  await page
    .getByLabel('Final net line total *', { exact: true })
    .selectOption('Final Sales Value');
  await page.getByLabel('Cost basis', { exact: true }).selectOption('unit');
  await expect(page.getByRole('button', { name: 'Calculate Bank Rate Deductions' })).toBeEnabled();
  await page.screenshot({ path: 'test-results/dashboard-mobile.png', fullPage: true });
  await page.getByRole('button', { name: 'Calculate Bank Rate Deductions' }).click();
  await expect(page.getByRole('cell', { name: '000123', exact: true })).toBeVisible();
  await page.getByRole('tab', { name: BRANCH_DATE_SUMMARY_SHEET, exact: true }).click();
  await expect(
    page.getByRole('heading', { name: BRANCH_DATE_SUMMARY_TITLE, exact: true }),
  ).toBeVisible();
  await expect(
    page.getByRole('table', { name: 'Calculated Sales Report' }).locator('thead th'),
  ).toHaveCount(6);
  expect(
    await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth),
  ).toBeTruthy();
});
test('Actual Sales Report: matched items retain source references in the four-sheet download', async ({
  page,
}) => {
  test.skip(!process.env.SAMPLE_SALES, 'Set SAMPLE_SALES to run the actual-report browser test.');
  await page.goto('/');
  await page
    .getByLabel('Upload Sales Report', { exact: true })
    .setInputFiles(process.env.SAMPLE_SALES!);
  await expect(page.getByRole('button', { name: 'Calculate Bank Rate Deductions' })).toBeEnabled();
  const calculated = page.waitForResponse((r) => r.url().endsWith('/api/excel/calculate'));
  await page.getByRole('button', { name: 'Calculate Bank Rate Deductions' }).click();
  const response = await calculated;
  expect(response.ok()).toBeTruthy();
  const data = await response.json();
  expect(data.result.reportDates).toEqual(['2026-10-01']);
  expect(data.result.items.length).toBe(data.result.summary.processedItems);
  const matchedItems = data.result.items.filter(
    (item: { cardName: string; quantity: string }) =>
      item.cardName.trim() && !(Number(item.quantity) < 0),
  );
  expect(data.result.summary.feeCalculatedItems).toBeGreaterThan(0);
  expect(data.result.summary.reviewCount).toBeGreaterThan(0);
  for (const partial of data.result.items.filter(
    (item: { partialCalculation: boolean }) => item.partialCalculation,
  )) {
    expect(partial.appliedRate).not.toBeNull();
    expect(partial.adjustment).not.toBeNull();
    expect(partial.explanation).toContain('Partial calculation');
  }
  const rowThree = data.result.items.find((item: { sourceRow: number }) => item.sourceRow === 3);
  expect(rowThree.originalPaymentMethod).toContain('BDO STRAIGHT');
  expect(rowThree.cardName).toBe('BDO STRAIGHT');
  expect(rowThree.appliedRate).toBe('2.85');
  expect(rowThree.adjustment).toBe('0.00');
  expect(rowThree.grossMargin).toBe('-238.00');
  const rowForty = data.result.items.find((item: { sourceRow: number }) => item.sourceRow === 40);
  expect(rowForty.originalPaymentMethod).toBe('Cash 27,100.00');
  expect(rowForty.cardName).toBe('');
  expect(rowForty.appliedRate).toBeNull();
  expect(rowForty.adjustment).toBe('0.00');
  expect(rowForty.grossMargin).toBeNull();
  const rowFortyOne = data.result.items.find(
    (item: { sourceRow: number }) => item.sourceRow === 41,
  );
  expect(rowFortyOne.cardName).toBe('BDO STRAIGHT');
  expect(rowFortyOne.appliedRate).toBe('2.85');
  expect(rowFortyOne.adjustment).toBe('968.97');
  expect(rowFortyOne.grossMargin).toBe('2766.03');
  const sourceRowThree = page
    .getByRole('table', { name: 'Calculated Sales Report' })
    .locator('tbody tr')
    .nth(matchedItems.findIndex((item: { sourceRow: number }) => item.sourceRow === 3));
  await expect(sourceRowThree.locator('td').nth(14)).toHaveText('3');
  await expect(sourceRowThree.locator('td').nth(15)).toHaveText('BDO STRAIGHT');
  await expect(sourceRowThree.locator('td').nth(11)).toHaveText('2.85%');
  await expect(sourceRowThree.locator('td').nth(12)).toHaveText('₱0.00');
  expect(matchedItems.some((item: { sourceRow: number }) => item.sourceRow === 40)).toBe(false);
  await expect(
    page
      .getByRole('table', { name: 'Calculated Sales Report' })
      .getByRole('cell', { name: rowForty.originalPaymentMethod, exact: true }),
  ).toHaveCount(0);
  await page.getByLabel('Search results').fill('3059-00005452');
  const table = page.getByRole('table', { name: 'Calculated Sales Report' });
  await expect(table.locator('tbody tr')).toHaveCount(3);
  await expect(table.getByRole('cell', { name: /Unconfigured BNPL\/SNPL/ })).toHaveCount(0);
  const labeledItems = data.result.items.filter(
    (item: { invoice: string }) => item.invoice === '3059-00005452',
  );
  expect(labeledItems).toHaveLength(3);
  for (const item of labeledItems) {
    expect(item.appliedRate).not.toBeNull();
    expect(item.adjustment).not.toBeNull();
    expect(item.partialCalculation).toBe(false);
  }
  const downloaded = page.waitForEvent('download');
  await page.getByRole('button', { name: /Download Calculated Excel Report/ }).click();
  const book = new ExcelJS.Workbook();
  await book.xlsx.load((await downloadBuffer(await downloaded)) as any);
  expect(book.worksheets.length).toBe(4);
  const ws = book.worksheets[0];
  expect(ws.columnCount).toBe(17);
  expect(ws.rowCount).toBe(matchedItems.length + 1);
  const rowThreeIndex =
    matchedItems.findIndex((item: { sourceRow: number }) => item.sourceRow === 3) + 2;
  expect(ws.getCell(rowThreeIndex, 16).value).toBe('BDO STRAIGHT');
  expect(ws.getCell(rowThreeIndex, 11).value).toBe(rowThree.originalPaymentMethod);
  expect(ws.getCell(rowThreeIndex, 12).value).toBe(0.0285);
  expect(ws.getCell(rowThreeIndex, 13).value).toBe(0);
  expect(ws.getCell(rowThreeIndex, 14).value).toBe(-238);
  for (const [index, item] of matchedItems.entries()) {
    expect(ws.getCell(index + 2, 1).value).toBe(item.productId);
    expect(ws.getCell(index + 2, 15).value).toBe(item.sourceRow);
    expect(ws.getCell(index + 2, 13).value).toBe(
      item.adjustment === null ? null : Number(item.adjustment),
    );
  }
});
