export interface PaymentMethod {
  bank: string;
  cardName: string;
  rate: string;
}
export interface Alias {
  alias: string;
  cardName: string;
}
export interface Pattern {
  pattern: string;
  cardName: string;
}
export interface Configuration {
  paymentMethods: PaymentMethod[];
  banks: string[];
  branches: string[];
  aliases: Alias[];
  patterns: Pattern[];
  nonCardMethods: string[];
}
export type CostBasis = 'auto' | 'unit' | 'line';
export interface Selection {
  sheet: string;
  headerRow: number;
  mapping: Record<string, string>;
  costBasis?: CostBasis;
  paymentScope?: 'invoice' | 'line';
}
export interface SheetPreview {
  name: string;
  headerRow: number;
  recordCount: number;
  headers: string[];
  preview: { sourceRow: number; cells: string[] }[];
  suggestedMapping: Record<string, string>;
  suggestedCostBasis: 'unit' | 'line' | 'unknown';
}
export interface WorkbookPreview {
  fileName: string;
  sheets: SheetPreview[];
}
export interface Source {
  file: string;
  sheet: string;
  row: number;
}
export interface PaymentPart {
  method: string;
  amount: string;
  reference: string;
  description: boolean;
}
export interface Sale {
  source: Source;
  productId: string;
  date: string;
  branch: string;
  invoice: string;
  product: string;
  quantity: string;
  price: string;
  cost: string;
  discount: string;
  lineTotal: string;
  payment: string;
  costBasis: 'unit' | 'line' | 'unknown';
  payments: PaymentPart[];
  paymentScope: 'invoice' | 'line';
  paymentIssues: string[];
  issues: string[];
  warnings: string[];
  notes: string[];
  cells: Record<string, string>;
  original: Record<string, string>;
}
export interface ItemResult {
  productId: string;
  salesDate: string;
  branch: string;
  invoice: string;
  product: string;
  quantity: string;
  price: string;
  cost: string;
  discount: string;
  lineTotal: string;
  bank: string;
  originalPaymentMethod: string;
  appliedRate: string | null;
  adjustment: string | null;
  grossMargin: string | null;
  partialCalculation: boolean;
  status: string;
  sourceRow: number;
  cardName: string;
  explanation: string;
  warnings: string[];
  issues: string[];
  totalCost: string | null;
  source: Source;
}
export interface BankSummary {
  bank: string;
  cardName: string;
  branch: string;
  sales: string;
  charges: string;
  subsidies: string;
  grossMargin: string;
  processedItems: number;
  reviewCount: number;
}
export interface Review {
  productId: string;
  salesDate: string;
  branch: string;
  invoice: string;
  product: string;
  sourceRow: number;
  status: string;
  reason: string;
  source: Source;
  original: Record<string, string>;
}
export interface Calculation {
  items: ItemResult[];
  bankSummary: BankSummary[];
  reviews: Review[];
  configuration: Configuration;
  calculatedAt: string;
  reportDates: string[];
  roundingPolicy: string;
  summary: {
    sales: string;
    charges: string;
    subsidies: string;
    grossMargin: string;
    processedItems: number;
    reviewCount: number;
    feeCalculatedItems: number;
    marginCalculatedItems: number;
    partialCalculationItems: number;
    missingProductIds: number;
  };
}
export interface CalculationResponse {
  result: Calculation;
  exportToken: string;
}
export const ITEM_COLUMNS = [
  { header: 'Product ID', key: 'productId', format: 'text' },
  { header: 'Sales Date', key: 'salesDate', format: 'date' },
  { header: 'Branch', key: 'branch', format: 'text' },
  { header: 'Invoice Number', key: 'invoice', format: 'text' },
  { header: 'Product Name', key: 'product', format: 'text' },
  { header: 'Quantity', key: 'quantity', format: 'number' },
  { header: 'Price', key: 'price', format: 'money' },
  { header: 'Cost', key: 'cost', format: 'money' },
  { header: 'Discount', key: 'discount', format: 'money' },
  { header: 'Line Total', key: 'lineTotal', format: 'money' },
  { header: 'Payment', key: 'originalPaymentMethod', format: 'text' },
  { header: 'Applied Bank Rate', key: 'appliedRate', format: 'percent' },
  { header: 'Bank Charges / Subsidy', key: 'adjustment', format: 'money' },
  { header: 'GrossMargin', key: 'grossMargin', format: 'money' },
  { header: 'Sales Source Row', key: 'sourceRow', format: 'number' },
  { header: 'Tender', key: 'cardName', format: 'text' },
  { header: 'Calculation Explanation', key: 'explanation', format: 'text' },
] as const;
