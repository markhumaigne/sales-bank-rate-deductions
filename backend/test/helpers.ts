import * as XLSX from 'xlsx';
import { Sale, PaymentPart } from '../../shared/models';
export function sale(values: Partial<Sale> = {}): Sale {
  return {
    source: { file: 'sales.xlsx', sheet: 'Sales Report', row: 2 },
    productId: 'LAPTOP-001',
    date: '2026-10-01',
    branch: 'ASUS',
    invoice: '0001',
    product: 'Laptop',
    quantity: '1',
    price: '30000',
    cost: '25000',
    discount: '0',
    lineTotal: '30000',
    payment: 'BDO 3 MOS',
    costBasis: 'unit',
    payments: [],
    paymentScope: 'invoice',
    paymentIssues: [],
    issues: [],
    warnings: [],
    notes: [],
    cells: {},
    original: {},
    ...values,
  };
}
export const part = (method: string, amount: string): PaymentPart => ({
  method,
  amount,
  reference: '',
  description: true,
});
export const upload = (buffer: Buffer, name = 'sales.xlsx') =>
  ({ buffer, originalname: name, size: buffer.length }) as Express.Multer.File;
export function workbook(sheets: Record<string, unknown[][]>): Buffer {
  const b = XLSX.utils.book_new();
  for (const [name, rows] of Object.entries(sheets))
    XLSX.utils.book_append_sheet(b, XLSX.utils.aoa_to_sheet(rows), name);
  return XLSX.write(b, { type: 'buffer', bookType: 'xlsx' });
}
export function reportFixture(): Buffer {
  return workbook({
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
      ['000123', 46296, 'ASUS', '0001', 'Laptop', 1, 30000, 25000, 0, 30000, 'BDO 3 MOS'],
      ['000124', 46296, 'ASUS', '0002', 'Mouse', 2, 500, 300, 100, 900, 'Cash'],
    ],
  });
}
