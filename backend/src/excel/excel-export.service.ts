import { BadRequestException, Injectable } from '@nestjs/common';
import ExcelJS from 'exceljs';
import Decimal from 'decimal.js';
import * as XLSX from 'xlsx';
import { Calculation, Configuration, ITEM_COLUMNS } from '../../../shared/models';
import { ConfigService } from '../config/config.service';
import { ExcelService } from './excel.service';
import { money } from '../calculation/money';

const currency = '"₱"#,##0.00;[Red]("₱"#,##0.00)';
const safe = (value: unknown): ExcelJS.CellValue =>
  typeof value === 'string' && /^[\s]*[=+\-@\t\r']/.test(value)
    ? "'" + value
    : (value as ExcelJS.CellValue);
const unescape = (v: string) =>
  v.startsWith("'") && /^[\s]*[=+\-@\t\r']/.test(v.slice(1)) ? v.slice(1) : v;
const source = (s: { file: string; sheet: string; row: number }) =>
  `${s.file} / ${s.sheet} / row ${s.row}`;
type Column = { header: string; key: string; width?: number; format?: string };
@Injectable()
export class ExcelExportService {
  constructor(
    private readonly config: ConfigService,
    private readonly excel: ExcelService,
  ) {}
  private table(
    book: ExcelJS.Workbook,
    name: string,
    columns: Column[],
    rows: Record<string, unknown>[],
    totals: string[] = [],
  ) {
    const ws = book.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 1 }] });
    ws.columns = columns.map((c) => ({
      header: c.header,
      key: c.key,
      width: c.width || 24,
      style: c.format ? { numFmt: c.format } : {},
    }));
    rows.forEach((row) =>
      ws.addRow(Object.fromEntries(Object.entries(row).map(([k, v]) => [k, safe(v)]))),
    );
    ws.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: Math.max(1, ws.rowCount), column: columns.length },
    };
    ws.getRow(1).height = 32;
    ws.getRow(1).eachCell((cell) => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF142B49' } };
      cell.alignment = { vertical: 'middle', wrapText: true };
    });
    for (let r = 2; r <= ws.rowCount; r++) {
      const review =
        Boolean(rows[r - 2]?.['needsReview']) ||
        String(rows[r - 2]?.['status'] || '').includes('Review') ||
        name === 'Needs Review';
      ws.getRow(r).eachCell((cell) => {
        cell.font = { name: 'Aptos', size: 10 };
        cell.alignment = { vertical: 'top', wrapText: true };
        if (review || r % 2 === 0)
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: review ? 'FFFFF1D6' : 'FFF3F6FA' },
          };
      });
    }
    if (totals.length) {
      const values: Record<string, unknown> = { [columns[0].key]: 'TOTAL — CONFIRMED' };
      for (const key of totals)
        values[key] = rows
          .filter((r) => !r['status'] || r['status'] === 'Confirmed')
          .reduce((a, r) => a.plus(String(r[key] ?? 0)), new Decimal(0))
          .toNumber();
      const total = ws.addRow(values);
      total.font = { bold: true };
      total.eachCell(
        (c) => (c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDE8F3' } }),
      );
    }
    return ws;
  }
  private settings(book: ExcelJS.Workbook, c: Configuration) {
    const columns: Column[] = [
      { header: 'Card Name', key: 'cardName', width: 28 },
      { header: 'Rate', key: 'rate', width: 14, format: '0.00%' },
      { header: 'Bank Reference', key: 'bankRef', width: 20 },
      { header: 'Branch', key: 'branch', width: 24 },
    ];
    const length = Math.max(
      c.paymentMethods.length,
      c.banks.length,
      c.branches.length,
      c.aliases.length,
      c.patterns.length,
      c.nonCardMethods.length,
      1,
    );
    const rows = Array.from({ length }, (_, i) => ({
      cardName: c.paymentMethods[i]?.cardName ?? '',
      rate: c.paymentMethods[i] ? new Decimal(c.paymentMethods[i].rate).div(100).toNumber() : null,
      bankRef: c.banks[i] ?? '',
      branch: c.branches[i] ?? '',
    }));
    const ws = this.table(book, 'Configuration Settings', columns, rows);
    ws.getCell('B1').note =
      'Excel percentage: enter 3.50% (stored as 0.035). Plain numeric 3.50 means 350% and is rejected. Text "3.50%" is also accepted.';
    ws.getCell('C1').note =
      'Independent bank reference list. Its row position has no relationship to card names or branches.';
    ws.getCell('A1').note =
      'Sales bank configuration v2. Edit these four reference columns. Cell notes preserve explicit card-bank assignments, aliases, patterns and non-card methods for accurate reimport. Manage those additional settings in the application. Keep the notes when editing or copying this sheet. Editing settings requires reimport and recalculation.';
    // Keep four columns without losing the calculation's complete configuration snapshot.
    // Notes travel with their cells and are validated as data, never executed.
    for (let i = 0; i < length; i++) {
      const note = (column: number, value: unknown) => {
        ws.getCell(i + 2, column).note = 'Ledger settings v2\n' + JSON.stringify(value);
      };
      if (c.paymentMethods[i])
        note(1, { cardName: c.paymentMethods[i].cardName, bank: c.paymentMethods[i].bank });
      if (c.aliases[i]) note(2, { alias: c.aliases[i] });
      if (c.banks[i] || c.patterns[i])
        note(3, { bankReference: c.banks[i], pattern: c.patterns[i] });
      if (c.nonCardMethods[i]) note(4, { nonCard: c.nonCardMethods[i] });
    }
    for (let r = 2; r <= length + 101; r++)
      ws.getCell(r, 2).dataValidation = {
        type: 'decimal',
        operator: 'between',
        formulae: [-1, 1],
        allowBlank: true,
        showErrorMessage: true,
        error: 'Enter an Excel percentage from -100% to 100%. Negative rates are subsidies.',
      };
    return ws;
  }
  async configuration(c: Configuration): Promise<Buffer> {
    const book = new ExcelJS.Workbook();
    this.settings(book, c);
    return Buffer.from(await book.xlsx.writeBuffer());
  }
  import(file: Express.Multer.File | undefined): Configuration {
    const { workbook } = this.excel.load(file);
    const ws = workbook.Sheets['Configuration Settings'];
    if (!ws)
      throw new BadRequestException(
        'Workbook must contain a Configuration Settings worksheet. Export a configuration template first.',
      );
    const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, raw: true, defval: '' });
    if (['Card Name', 'Rate', 'Bank Reference', 'Branch'].every((h, i) => rows[0]?.[i] === h))
      return this.importCompactSettings(ws, rows);
    const expected = [
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
    ];
    const headers = (rows[0] || []).map(String);
    if (!expected.every((v, i) => headers[i] === v))
      throw new BadRequestException(
        'Configuration headers have changed. Use the exported template and keep its header row.',
      );
    const c: Configuration = {
      paymentMethods: [],
      banks: [],
      branches: [],
      aliases: [],
      patterns: [],
      nonCardMethods: [],
    };
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const text = (col: number) => unescape(String(row[col] ?? '').trim());
      for (let col = 0; col < 10; col++)
        if (ws[XLSX.utils.encode_cell({ r: i, c: col })]?.f)
          throw new BadRequestException(
            `Configuration row ${i + 1}: formulas are not supported; enter literal settings.`,
          );
      if (text(0) || text(1) || text(2)) {
        let rate: string;
        try {
          const d =
            typeof row[2] === 'string' && text(2).endsWith('%')
              ? new Decimal(text(2).slice(0, -1))
              : new Decimal(String(row[2])).times(100);
          rate = d.toFixed(Math.max(2, d.decimalPlaces()));
        } catch {
          throw new BadRequestException(
            `Configuration row ${i + 1}: invalid rate. Enter 3.50% in Excel.`,
          );
        }
        c.paymentMethods.push({ cardName: text(0), bank: text(1), rate });
      }
      if (text(3)) c.banks.push(text(3));
      if (text(4)) c.branches.push(text(4));
      if (text(5) || text(6)) c.aliases.push({ alias: text(5), cardName: text(6) });
      if (text(7) || text(8)) c.patterns.push({ pattern: text(7), cardName: text(8) });
      if (text(9)) c.nonCardMethods.push(text(9));
    }
    return this.config.validate(c);
  }

  private importCompactSettings(ws: XLSX.WorkSheet, rows: unknown[][]): Configuration {
    const defaults = this.config.defaults();
    const c: Configuration = {
      paymentMethods: [],
      banks: [],
      branches: [],
      aliases: [],
      patterns: [],
      nonCardMethods: [],
    };
    const key = (v: string) => v.trim().replace(/\s+/g, ' ').toUpperCase();
    const renamedCards = new Map<string, string>(),
      renamedBanks = new Map<string, string>();
    const hasSnapshot = (ws.A1?.c || []).some((comment: { t?: string }) =>
      comment.t?.includes('Sales bank configuration v2'),
    );
    const metadata = (r: number, col: number): Record<string, unknown> => {
      const cell = ws[XLSX.utils.encode_cell({ r, c: col })];
      if (cell?.f || cell?.t === 'e')
        throw new BadRequestException(
          'Configuration row ' +
            (r + 1) +
            ': enter literal settings, not formulas or Excel errors.',
        );
      const note = cell?.c?.find((comment: { t?: string }) =>
        comment.t?.startsWith('Ledger settings v2\n'),
      )?.t;
      if (!note) return {};
      try {
        const value = JSON.parse(note.slice('Ledger settings v2\n'.length));
        if (!value || Array.isArray(value) || typeof value !== 'object') throw new Error();
        return value;
      } catch {
        throw new BadRequestException(
          'Configuration row ' + (r + 1) + ': invalid settings note. Export a fresh template.',
        );
      }
    };
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i],
        text = (col: number) => unescape(String(row[col] ?? '').trim());
      const notes = [0, 1, 2, 3].map((col) => metadata(i, col));
      if (text(0) || text(1)) {
        let rate: string;
        try {
          const d =
            typeof row[1] === 'string' && text(1).endsWith('%')
              ? new Decimal(text(1).slice(0, -1))
              : new Decimal(String(row[1])).times(100);
          rate = d.toFixed(Math.max(2, d.decimalPlaces()));
        } catch {
          throw new BadRequestException(
            'Configuration row ' +
              (i + 1) +
              ': invalid rate. Enter an Excel percentage such as 3.50%.',
          );
        }
        const bank =
          typeof notes[0].bank === 'string'
            ? notes[0].bank
            : defaults.paymentMethods.find((m) => key(m.cardName) === key(text(0)))?.bank;
        if (!bank)
          throw new BadRequestException(
            'Configuration row ' +
              (i + 1) +
              ': bank assignment for "' +
              text(0) +
              '" is missing. Add this method in the application and export its settings.',
          );
        c.paymentMethods.push({ cardName: text(0), bank, rate });
        if (typeof notes[0].cardName === 'string')
          renamedCards.set(key(notes[0].cardName), text(0));
      }
      if (text(2)) {
        c.banks.push(text(2));
        if (typeof notes[2].bankReference === 'string')
          renamedBanks.set(key(notes[2].bankReference), text(2));
      }
      if (text(3)) c.branches.push(text(3));
      if (notes[1].alias !== undefined)
        c.aliases.push(notes[1].alias as Configuration['aliases'][number]);
      if (notes[2].pattern !== undefined)
        c.patterns.push(notes[2].pattern as Configuration['patterns'][number]);
      if (notes[3].nonCard !== undefined) c.nonCardMethods.push(notes[3].nonCard as string);
    }
    // Reference lists are independent; never pair cards and banks using row positions.
    c.paymentMethods.forEach((m) => (m.bank = renamedBanks.get(key(m.bank)) ?? m.bank));
    for (const entry of [...c.aliases, ...c.patterns]) {
      if (entry && typeof entry.cardName === 'string')
        entry.cardName = renamedCards.get(key(entry.cardName)) ?? entry.cardName;
    }
    if (!hasSnapshot && !c.nonCardMethods.length) c.nonCardMethods = defaults.nonCardMethods;
    return this.config.validate(c);
  }

  async report(result: Calculation): Promise<Buffer> {
    const book = new ExcelJS.Workbook();
    book.creator = 'Ledger · Sales Bank Rate Calculator';
    book.created = new Date(result.calculatedAt);
    book.description = result.roundingPolicy;
    const date = (value: string) => (value ? new Date(value + 'T00:00:00.000Z') : null);
    const numeric = (value: unknown) =>
      typeof value === 'string' && money(value)
        ? money(value)!.toNumber()
        : value === null
          ? null
          : value;
    const columns: Column[] = ITEM_COLUMNS.map((c) => ({
      header: c.header,
      key: c.key,
      width:
        c.key === 'explanation'
          ? 100
          : c.key === 'product'
            ? 48
            : c.key === 'originalPaymentMethod'
              ? 65
              : c.key === 'productId'
                ? 24
                : 22,
      format:
        c.format === 'money'
          ? currency
          : c.format === 'percent'
            ? '0.00%'
            : c.format === 'date'
              ? 'dd mmm yyyy'
              : c.format === 'text'
                ? '@'
                : undefined,
    }));
    const ws = this.table(
      book,
      'Item Bank Deductions',
      columns,
      result.items.map((i) => ({
        ...i,
        needsReview: !!(i.issues.length || i.warnings.length),
        salesDate: date(i.salesDate),
        quantity:
          /^-?\d+(\.\d+)?$/.test(i.quantity) && Number.isFinite(Number(i.quantity))
            ? Number(i.quantity)
            : i.quantity,
        price: numeric(i.price),
        cost: numeric(i.cost),
        discount: numeric(i.discount),
        lineTotal: numeric(i.lineTotal),
        appliedRate:
          i.appliedRate !== null && /^-?\d+(\.\d+)?$/.test(i.appliedRate)
            ? new Decimal(i.appliedRate).div(100).toNumber()
            : i.appliedRate,
        adjustment: numeric(i.adjustment),
        grossMargin: numeric(i.grossMargin),
      })),
    );
    // An explicit Excel string cell is inert even when an identifier starts with '='.
    // Preserve the actual identifier rather than adding a visible apostrophe to it.
    result.items.forEach((i, n) => {
      ws.getCell(n + 2, 1).value = i.productId;
      ws.getCell(n + 2, 1).numFmt = '@';
      // Copy source Payment verbatim into a literal string cell, including line breaks.
      ws.getCell(n + 2, ws.getColumn('originalPaymentMethod').number).value =
        i.originalPaymentMethod;
    });
    ws.getCell('H1').note =
      'Original Cost value. Unit/line interpretation and normalized total product cost are documented in Calculation Explanation.';
    ws.getCell('J1').note = 'Final net line sales amount. Discount is not deducted a second time.';
    ws.getCell(1, ws.getColumn('cardName').number).note =
      'Recognized configured card names from Sales Report payment details. A name remains visible when another payment or source value needs review; it does not mean the bank charge has been calculated. See Calculation Explanation.';
    ws.getCell(1, ws.getColumn('adjustment').number).note =
      'Positive = bank charge; negative = subsidy. Blank = no configured card match or unresolved calculation. Cash-only and SBC QRPH-only rows are not calculated. In a split, only the matched card portion receives a bank charge. When another payment method is unconfigured, only matched-method charges are shown; see Calculation Explanation. Partial rows are excluded from confirmed totals.';
    ws.getCell(1, ws.getColumn('grossMargin').number).note =
      'Line Total minus total product cost and calculated bank charges/subsidies. Blank when no configured card matches, or cost/charges are unresolved. Provisional when another payment method remains unconfigured; see Calculation Explanation. Partial rows are excluded from confirmed totals.';
    const summary = this.table(
      book,
      'Bank & Branch Summary',
      [
        { header: 'Bank', key: 'bank' },
        { header: 'Configured Payment Method', key: 'cardName', width: 30 },
        { header: 'Branch', key: 'branch' },
        { header: 'Total Sales', key: 'sales', format: currency },
        { header: 'Total Bank Charges', key: 'charges', format: currency },
        { header: 'Total Subsidies', key: 'subsidies', format: currency },
        { header: 'Total Gross Margin', key: 'grossMargin', format: currency },
        { header: 'Total Processed Items', key: 'processedItems' },
        { header: 'Transactions Requiring Review', key: 'reviewCount' },
      ],
      result.bankSummary.map((b) => ({
        ...b,
        sales: numeric(b.sales),
        charges: numeric(b.charges),
        subsidies: numeric(b.subsidies),
        grossMargin: numeric(b.grossMargin),
      })),
      ['sales', 'charges', 'subsidies', 'grossMargin'],
    );
    summary.lastRow!.getCell(8).value = result.summary.processedItems;
    summary.lastRow!.getCell(9).value = result.summary.reviewCount;
    summary.getCell('H1').note =
      'Unique items per bank/method. Split-payment items can occur in more than one group; the grand total is the original source item count.';
    summary.getCell('D1').note =
      'Only validated payment allocations on rows with matched configured cards contribute sales/charges/subsidies. Cash-only and other rows without a configured card contribute counts only. Partial calculations with unconfigured payment methods are excluded from all confirmed financial totals. Gross margin includes only rows with known product cost. Missing identifiers do not discard valid monetary calculations.';
    this.settings(book, result.configuration);
    this.table(
      book,
      'Needs Review',
      [
        { header: 'Product ID', key: 'productId', format: '@' },
        { header: 'Sales Date', key: 'salesDate', format: 'dd mmm yyyy' },
        { header: 'Branch', key: 'branch' },
        { header: 'Invoice Number', key: 'invoice', format: '@' },
        { header: 'Product Name', key: 'product', width: 48 },
        { header: 'Sales Source Row', key: 'sourceRow' },
        { header: 'Status', key: 'status' },
        { header: 'Reason / Warning', key: 'reason', width: 100 },
        { header: 'Source File / Worksheet', key: 'source', width: 65 },
        { header: 'Original Source Values', key: 'original', width: 100 },
      ],
      result.reviews.map((r) => ({
        ...r,
        salesDate: date(r.salesDate),
        source: source(r.source),
        original: JSON.stringify(r.original),
      })),
    );
    const reviewSheet = book.getWorksheet('Needs Review')!;
    result.reviews.forEach((r, n) => {
      reviewSheet.getCell(n + 2, 1).value = r.productId;
      reviewSheet.getCell(n + 2, 1).numFmt = '@';
    });
    return Buffer.from(await book.xlsx.writeBuffer());
  }
}
