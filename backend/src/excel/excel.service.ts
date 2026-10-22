import { BadRequestException, Injectable } from '@nestjs/common';
import * as XLSX from 'xlsx';
import { unzipSync } from 'fflate';
import Decimal from 'decimal.js';
import {
  Sale,
  Selection,
  SheetPreview,
  WorkbookPreview,
  PaymentPart,
} from '../../../shared/models';
import { normalize } from '../config/config.service';
import { reportDate } from './report-date';

export const FIELDS: Record<string, string[]> = {
  productId: ['Product ID', 'ProductID', 'SKU', 'Item Code', 'Product Code'],
  date: ['Sales Date', 'Date', 'Transaction Date'],
  branch: ['Branch', 'Branch Name'],
  invoice: ['TM-INV#', 'Invoice Number', 'Invoice No', 'Reference No'],
  product: ['Product Name', 'Product', 'Item Name', 'Description'],
  quantity: ['Quantity', 'Qty'],
  price: ['Price', 'Unit Price', 'Selling Price'],
  cost: ['Unit Cost', 'Total Line Cost', 'Line Cost', 'Total Cost', 'Cost'],
  discount: ['Discount', 'Discount Amount'],
  lineTotal: ['Net Total', 'Net Sales', 'Net Amount', 'Line Total', 'Item Net Sales Amount'],
  payment: ['Payment', 'Original Tender Payment Method', 'Payment Method', 'Card Name'],
  paymentAmount: ['Line Payment Amount', 'Payment Amount', 'Card Amount', 'Paid Amount'],
};
const REQUIRED = ['invoice', 'product', 'quantity', 'lineTotal'];
const MAX_ROWS = 50000,
  MAX_COLUMNS = 200;
type Loaded = { workbook: XLSX.WorkBook; fileName: string };
@Injectable()
export class ExcelService {
  load(file: Express.Multer.File | undefined): Loaded {
    if (!file?.buffer?.length) throw new BadRequestException('Select an Excel workbook first.');
    if (!/\.xlsx?$/i.test(file.originalname))
      throw new BadRequestException('Only .xlsx and .xls workbooks are accepted.');
    const b = file.buffer;
    const zip = b[0] === 0x50 && b[1] === 0x4b;
    const ole = b
      .subarray(0, 8)
      .equals(Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]));
    if (!zip && !ole)
      throw new BadRequestException(
        'The file is not a valid Excel workbook. CSV/HTML files renamed to .xlsx are not supported.',
      );
    try {
      if (zip) {
        let expanded = 0,
          entries = 0;
        unzipSync(b, {
          filter: (entry) => {
            expanded += entry.originalSize;
            entries++;
            if (expanded > 100 * 1024 * 1024 || entries > 2000)
              throw new Error(
                'Workbook is too large when expanded (100 MB / 2,000 ZIP entries maximum).',
              );
            return false;
          },
        });
      }
      const workbook = XLSX.read(b, {
        type: 'buffer',
        cellText: true,
        cellNF: true,
        cellFormula: true,
        sheetRows: MAX_ROWS + 2,
        bookVBA: false,
        cellHTML: false,
      });
      if (!workbook.SheetNames.length || workbook.SheetNames.length > 30)
        throw new Error('Workbook must contain 1 to 30 worksheets.');
      for (const sheet of Object.values(workbook.Sheets)) {
        const ref = sheet['!fullref'] || sheet['!ref'];
        if (ref) {
          const range = XLSX.utils.decode_range(ref);
          if (range.e.r >= MAX_ROWS + 1 || range.e.c >= MAX_COLUMNS)
            throw new Error('Each sheet must contain at most 50,000 data rows and 200 columns.');
        }
      }
      return { workbook, fileName: file.originalname };
    } catch (e) {
      throw new BadRequestException(
        `Cannot read workbook: ${e instanceof Error ? e.message : 'invalid or password-protected Excel file'}`,
      );
    }
  }
  private text(cell: XLSX.CellObject | undefined, formatted = false): string {
    if (!cell) return '';
    if (cell.t === 'e') return cell.w || '#EXCEL_ERROR';
    if (cell.f && cell.v === undefined) return '#FORMULA_WITHOUT_CACHED_VALUE';
    // Use Excel's displayed text for invoice references, including numeric zero masks.
    if (formatted) return String(cell.w ?? XLSX.utils.format_cell(cell));
    return cell.v == null ? '' : String(cell.v);
  }
  private amount(cell: XLSX.CellObject | undefined, address: string, notes: string[]): string {
    const raw = this.text(cell);
    if (cell?.t !== 'n' || !Number.isFinite(cell.v)) return raw;
    const value = new Decimal(String(cell.v));
    const rounded = value.toDecimalPlaces(2, Decimal.ROUND_HALF_UP);
    // Correct only binary numeric storage noise. Real fractions of a cent stay invalid.
    // Text cells are never silently corrected, even when they look numeric.
    if (!value.eq(rounded) && value.minus(rounded).abs().lte('0.00000001')) {
      notes.push(
        `${address}: Excel numeric storage noise normalized from ${raw} to ${rounded.toFixed(2)}.`,
      );
      return rounded.toFixed(2);
    }
    return raw;
  }

  private costBasis(
    headers: string[],
    mapping: Record<string, string>,
  ): 'unit' | 'line' | 'unknown' {
    const header = normalize(mapping.cost || '');
    if (header === 'UNIT COST') return 'unit';
    if (['TOTAL LINE COST', 'LINE COST', 'TOTAL COST'].includes(header)) return 'line';
    // Verified reference structure: repeated SKU costs remain unit costs at quantities 2, 5 and 8.
    const names = new Set(headers.map(normalize));
    if (
      header === 'COST' &&
      ['PRODUCT ID', 'QTY', 'BASE QTY', 'MASTER PRICE', 'NET TOTAL'].every((n) => names.has(n))
    )
      return 'unit';
    return 'unknown';
  }
  private describe(loaded: Loaded, name: string, headerOverride?: number): SheetPreview {
    const sheet = loaded.workbook.Sheets[name];
    if (!sheet) throw new BadRequestException('Worksheet "' + name + '" does not exist.');
    const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
    let header = 0,
      best = -1;
    for (let r = 0; r <= Math.min(range.e.r, 49); r++) {
      const names = new Set(
        Array.from({ length: range.e.c + 1 }, (_, c) =>
          normalize(this.text(sheet[XLSX.utils.encode_cell({ r, c })])),
        ),
      );
      const score = Object.values(FIELDS).filter((aliases) =>
        aliases.some((a) => names.has(normalize(a))),
      ).length;
      if (score > best) {
        header = r;
        best = score;
      }
    }
    if (headerOverride !== undefined) {
      if (
        !Number.isInteger(headerOverride) ||
        headerOverride < 1 ||
        headerOverride > Math.min(50, range.e.r + 1)
      )
        throw new BadRequestException('Header row must be between 1 and 50.');
      header = headerOverride - 1;
    }
    const seen = new Set<string>();
    const originalHeaders: string[] = [];
    const headers = Array.from({ length: range.e.c + 1 }, (_, c) => {
      const raw =
        this.text(sheet[XLSX.utils.encode_cell({ r: header, c })])
          .replace(/\s+/g, ' ')
          .trim() || 'Column ' + XLSX.utils.encode_col(c);
      originalHeaders.push(normalize(raw));
      let h = raw;
      while (seen.has(normalize(h))) h += ' [' + XLSX.utils.encode_col(c) + ']';
      seen.add(normalize(h));
      return h;
    });
    const mapping: Record<string, string> = {};
    const ambiguousFields: string[] = [];
    for (const [key, aliases] of Object.entries(FIELDS)) {
      for (const alias of aliases) {
        const indices = originalHeaders.flatMap((h, index) =>
          h === normalize(alias) ? [index] : [],
        );
        if (!indices.length) continue;
        if (indices.length === 1) mapping[key] = headers[indices[0]];
        else ambiguousFields.push(key);
        // Do not silently substitute another alias when the preferred heading
        // occurs twice. Only the user can choose the intended source column.
        break;
      }
    }
    let recordCount = 0;
    const preview: SheetPreview['preview'] = [];
    for (let r = header + 1; r <= range.e.r; r++) {
      const cells = headers.map((_, c) => this.text(sheet[XLSX.utils.encode_cell({ r, c })], true));
      if (!cells.some((v) => v.trim())) continue;
      recordCount++;
      if (preview.length < 10) {
        const dc = headers.indexOf(mapping.date);
        if (dc >= 0) {
          const date = reportDate(
            sheet[XLSX.utils.encode_cell({ r, c: dc })],
            !!loaded.workbook.Workbook?.WBProps?.date1904,
          );
          if (date) cells[dc] = date;
        }
        preview.push({ sourceRow: r + 1, cells });
      }
    }
    return {
      name,
      headerRow: header + 1,
      recordCount,
      headers,
      preview,
      suggestedMapping: mapping,
      ambiguousFields,
      suggestedCostBasis: this.costBasis(headers, mapping),
    };
  }
  preview(file: Express.Multer.File | undefined, headerRow?: number): WorkbookPreview {
    const loaded = this.load(file);
    return {
      fileName: loaded.fileName,
      sheets: loaded.workbook.SheetNames.map((name) => this.describe(loaded, name, headerRow)),
    };
  }
  parse(file: Express.Multer.File | undefined, selection: Selection): Sale[] {
    const loaded = this.load(file),
      desc = this.describe(loaded, selection.sheet, selection.headerRow),
      mapping = selection.mapping;
    for (const key of REQUIRED)
      if (!mapping[key] || !desc.headers.includes(mapping[key]))
        throw new BadRequestException(
          'Map required sales field "' + key + '" to an existing column.',
        );
    if (new Set(REQUIRED.map((k) => mapping[k])).size !== REQUIRED.length)
      throw new BadRequestException('Required sales fields must use separate columns.');
    for (const key of desc.ambiguousFields ?? [])
      if (!mapping[key])
        throw new BadRequestException(
          'Duplicate column headings for "' + key + '". Select the intended source column.',
        );
    const mappedColumns = new Map<string, string>();
    for (const [key, value] of Object.entries(mapping)) {
      if (value && (!Object.hasOwn(FIELDS, key) || !desc.headers.includes(value)))
        throw new BadRequestException('Invalid column mapping: ' + key);
      if (!value) continue;
      const previous = mappedColumns.get(value);
      if (previous)
        throw new BadRequestException(
          'Sales fields must use separate columns: "' +
            previous +
            '" and "' +
            key +
            '" both use "' +
            value +
            '".',
        );
      mappedColumns.set(value, key);
    }
    const sheet = loaded.workbook.Sheets[selection.sheet],
      range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
    const result: Sale[] = [];
    const basis =
      selection.costBasis && selection.costBasis !== 'auto'
        ? selection.costBasis
        : this.costBasis(desc.headers, mapping);
    for (let r = desc.headerRow; r <= range.e.r; r++) {
      const row = desc.headers.map((_, c) => this.text(sheet[XLSX.utils.encode_cell({ r, c })]));
      if (!row.some((v) => v.trim())) continue;
      const cells: Record<string, string> = {},
        original: Record<string, string> = {},
        notes: string[] = [],
        issues: string[] = [],
        warnings: string[] = [],
        paymentIssues: string[] = [];
      for (const [key, h] of Object.entries(mapping)) {
        const c = desc.headers.indexOf(h);
        if (c < 0) continue;
        const address = XLSX.utils.encode_cell({ r, c });
        cells[key] = address;
        original[key] = this.text(sheet[address], ['productId', 'invoice', 'date'].includes(key));
        const cell = sheet[address];
        if (cell?.f && cell.v !== undefined)
          notes.push(
            address +
              ': cached formula value read; source workbook must be recalculated before upload.',
          );
        if (
          ['invoice', 'product', 'branch', 'date', 'payment'].includes(key) &&
          (cell?.t === 'e' || (cell?.f && cell.v === undefined))
        )
          issues.push(
            key + ' at ' + address + ' contains an Excel error or missing formula result.',
          );
        if (key === 'invoice' && cell?.t === 'n' && !Number.isSafeInteger(cell.v))
          issues.push('Numeric invoice cannot be recovered reliably; store it as text.');
      }
      const get = (key: string) => original[key] ?? '';
      const amount = (key: string) =>
        cells[key] ? this.amount(sheet[cells[key]], cells[key], notes) : '';
      let productId = get('productId');
      const idCell = sheet[cells.productId];
      if (idCell?.t === 'e' || (idCell?.f && idCell.v === undefined)) {
        productId = '';
        warnings.push('Product ID cell has an Excel error; no identifier was fabricated.');
      }
      if (!productId.trim()) {
        productId = '';
        warnings.push('Missing Product ID; select its column or correct the source identifier.');
      }
      if (idCell?.t === 'n' && !Number.isSafeInteger(idCell.v))
        warnings.push(
          'Product ID is stored as an unsafe numeric value; verify against the original identifier. Displayed text was retained.',
        );
      const date = cells.date
        ? reportDate(sheet[cells.date], !!loaded.workbook.Workbook?.WBProps?.date1904)
        : null;
      if (!date)
        issues.push('Missing or invalid Sales Date. Use an Excel date, YYYY-MM-DD or MM/DD/YYYY.');
      const payments: PaymentPart[] = [];
      if (mapping.paymentAmount) {
        if (get('paymentAmount').trim() || get('payment').trim())
          payments.push({
            method: get('payment'),
            amount: amount('paymentAmount'),
            reference: '',
            description: false,
          });
      } else {
        // These are Sales Report columns, repeated once per product; never additional input files.
        for (let c = 0; c < desc.headers.length; c++) {
          if (
            !/^Tender(?: \[[A-Z]+\])?$/i.test(desc.headers[c]) ||
            !/^Amount(?: \[[A-Z]+\])?$/i.test(desc.headers[c + 2] || '')
          )
            continue;
          if (!row[c]?.trim() && !row[c + 2]?.trim()) continue;
          const address = XLSX.utils.encode_cell({ r, c: c + 2 });
          payments.push({
            method: row[c] || '',
            amount: this.amount(sheet[address], address, notes),
            reference: row[c + 1] || '',
            description: true,
          });
        }
      }
      let payment = get('payment');
      if (!payment && payments.length) {
        payment = payments.map((p) => p.method).join('\n');
        notes.push('Payment text comes from the original Sales Report payment-method columns.');
      }
      result.push({
        source: { file: loaded.fileName, sheet: selection.sheet, row: r + 1 },
        productId,
        date: date || '',
        branch: get('branch'),
        invoice: get('invoice'),
        product: get('product'),
        quantity: get('quantity'),
        price: amount('price'),
        cost: amount('cost'),
        discount: amount('discount'),
        lineTotal: amount('lineTotal'),
        payment,
        costBasis: basis,
        payments,
        paymentScope: selection.paymentScope || 'invoice',
        paymentIssues,
        issues,
        warnings,
        notes,
        cells,
        original,
      });
    }
    if (!result.length)
      throw new BadRequestException('The selected Sales Report worksheet contains no data rows.');
    return result;
  }
}
