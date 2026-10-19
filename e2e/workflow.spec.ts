import { test, expect, Download } from '@playwright/test';
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { ITEM_COLUMNS } from '../shared/models';
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
  await expect(page.getByRole('tab', { name: /Payment methods \(50\)/ })).toBeVisible();
  await page.screenshot({ path: 'test-results/dashboard-desktop.png', fullPage: true });
  await page.getByLabel('Upload Sales Report', { exact: true }).setInputFiles(fixture());
  await page.getByLabel('Worksheet', { exact: true }).selectOption('Sales Report');
  await expect(calculate).toBeEnabled();
  const responsePromise = page.waitForResponse((r) => r.url().endsWith('/api/excel/calculate'));
  await calculate.click();
  const response = await responsePromise;
  expect(response.ok()).toBeTruthy();
  const payload = await response.json();
  expect(payload.result.summary.charges).toBe('1050.00');
  expect(payload.result.summary.grossMargin).toBe('3950.00');
  expect(payload.result.summary.sales).toBe('30000.00');
  expect(payload.result.items[1].appliedRate).toBeNull();
  expect(payload.result.items[1].adjustment).toBeNull();
  expect(payload.result.items[1].grossMargin).toBeNull();
  const table = page.getByRole('table', { name: 'Calculated Sales Report' });
  await expect(table.locator('thead th')).toHaveCount(17);
  expect(
    (await table.locator('thead th').allTextContents()).map((s) => s.replace(/[↑↓]/g, '').trim()),
  ).toEqual(ITEM_COLUMNS.map((c) => c.header));
  await expect(table.getByRole('cell', { name: '000123', exact: true })).toBeVisible();
  await table.getByRole('button', { name: 'Product ID', exact: true }).click();
  await table.getByRole('button', { name: /^Product ID/ }).click();
  await expect(table.locator('tbody tr').first().locator('td').first()).toHaveText('000124');
  await expect(table.locator('tbody tr').first().locator('td').nth(4)).toHaveText('Mouse');
  for (const column of [11, 12, 13, 15])
    await expect(table.locator('tbody tr').first().locator('td').nth(column)).toHaveText('—');
  await page.getByLabel('Search results').fill('000123');
  await expect(table.locator('tbody tr')).toHaveCount(1);
  await expect(table.getByRole('cell', { name: 'Laptop', exact: true })).toBeVisible();
  await page.getByLabel('Search results').clear();
  const downloaded = page.waitForEvent('download');
  await page.getByRole('button', { name: /Download Calculated Excel Report/ }).click();
  const download = await downloaded;
  expect(download.suggestedFilename()).toMatch(/^Sales_Bank_Rate_Deductions_\d{8}_\d{6}\.xlsx$/);
  const report = new ExcelJS.Workbook();
  await report.xlsx.load((await downloadBuffer(download)) as any);
  expect(report.worksheets.map((s) => s.name)).toEqual([
    'Item Bank Deductions',
    'Bank & Branch Summary',
    'Configuration Settings',
    'Needs Review',
  ]);
  const ws = report.worksheets[0];
  expect(ws.columnCount).toBe(17);
  expect(ws.rowCount).toBe(3);
  expect(ws.getCell('A2').value).toBe('000123');
  expect(ws.getCell('M2').value).toBe(1050);
  expect(ws.getCell('N2').value).toBe(3950);
  for (const address of ['L3', 'M3', 'N3']) expect(ws.getCell(address).value).toBeNull();
  expect(ws.getCell('P3').value).toBe('');
  expect(ws.getCell('A2').numFmt).toBe('@');
  await page.getByLabel('Search payment methods').fill('BDO 3 MOS');
  await page.getByLabel('Rate percentage', { exact: true }).fill('4.00');
  await expect(
    page.getByRole('heading', { name: 'Your calculated report will appear here' }),
  ).toBeVisible();
  await calculate.click();
  await expect(page.getByText('₱1,200.00', { exact: true }).first()).toBeVisible();
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
  expect(
    await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth),
  ).toBeTruthy();
});
test('Actual Sales Report: every source row retained, October 1 dates and four-sheet download', async ({
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
  expect(data.result.summary.feeCalculatedItems).toBeGreaterThan(0);
  expect(data.result.summary.reviewCount).toBeGreaterThan(0);
  const partial = data.result.items.find(
    (item: { partialCalculation: boolean }) => item.partialCalculation,
  );
  expect(partial).toBeTruthy();
  expect(partial.appliedRate).not.toBeNull();
  expect(partial.adjustment).not.toBeNull();
  expect(partial.explanation).toContain('Partial calculation');
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
  expect(rowForty.adjustment).toBeNull();
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
    .nth(1);
  await expect(sourceRowThree.locator('td').nth(14)).toHaveText('3');
  await expect(sourceRowThree.locator('td').nth(15)).toHaveText('BDO STRAIGHT');
  await expect(sourceRowThree.locator('td').nth(11)).toHaveText('2.85%');
  await expect(sourceRowThree.locator('td').nth(12)).toHaveText('₱0.00');
  const sourceRowForty = page
    .getByRole('table', { name: 'Calculated Sales Report' })
    .locator('tbody tr')
    .nth(38);
  await expect(sourceRowForty.locator('td').nth(14)).toHaveText('40');
  for (const column of [11, 12, 13, 15])
    await expect(sourceRowForty.locator('td').nth(column)).toHaveText('—');
  await page.getByLabel('Search results').fill('3059-00005452');
  const table = page.getByRole('table', { name: 'Calculated Sales Report' });
  await expect(table.locator('tbody tr')).toHaveCount(3);
  await expect(table.getByRole('cell', { name: /Unconfigured BNPL\/SNPL/ })).toHaveCount(3);
  const downloaded = page.waitForEvent('download');
  await page.getByRole('button', { name: /Download Calculated Excel Report/ }).click();
  const book = new ExcelJS.Workbook();
  await book.xlsx.load((await downloadBuffer(await downloaded)) as any);
  expect(book.worksheets.length).toBe(4);
  const ws = book.worksheets[0];
  expect(ws.columnCount).toBe(17);
  expect(ws.rowCount).toBe(data.result.items.length + 1);
  expect(ws.getCell('P3').value).toBe('BDO STRAIGHT');
  expect(ws.getCell('K3').value).toBe(rowThree.originalPaymentMethod);
  expect(ws.getCell('L3').value).toBe(0.0285);
  expect(ws.getCell('M3').value).toBe(0);
  expect(ws.getCell('N3').value).toBe(-238);
  expect(ws.getCell('P40').value).toBe('');
  expect(ws.getCell('K40').value).toBe(rowForty.originalPaymentMethod);
  for (const address of ['L40', 'M40', 'N40']) expect(ws.getCell(address).value).toBeNull();
  for (const [index, item] of data.result.items.entries()) {
    expect(ws.getCell(index + 2, 1).value).toBe(item.productId);
    expect(ws.getCell(index + 2, 15).value).toBe(item.sourceRow);
    expect(ws.getCell(index + 2, 13).value).toBe(
      item.adjustment === null ? null : Number(item.adjustment),
    );
  }
});
