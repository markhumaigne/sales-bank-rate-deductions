export interface PaymentMethod {
  bank: string;
  cardName: string;
  rate: string;
  // Optional rate overrides keyed by branch name; the base rate remains the fallback.
  branchRates?: Record<string, string>;
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
  ambiguousFields?: string[];
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
  appliedRate1: string | null;
  appliedRate2: string | null;
  appliedRate3: string | null;
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
  // Payment selection rules can route an entire shared invoice to Needs Review.
  reviewOnly?: boolean;
  source: Source;
}
// Tender is populated by the existing configuration matcher, including aliases and BNPL/SNPL.
export const hasConfiguredBankPayment = (item: Pick<ItemResult, 'cardName'>): boolean =>
  item.cardName.trim().length > 0;
export const hasNegativeQuantity = (item: Pick<ItemResult, 'quantity'>): boolean =>
  Number(item.quantity) < 0;
// Use one selection rule for the Angular preview, Excel Sheet 1 and Sheet 2.
export const showInItemBankDeductions = (
  item: Pick<ItemResult, 'cardName' | 'quantity' | 'reviewOnly'>,
): boolean => hasConfiguredBankPayment(item) && !hasNegativeQuantity(item) && !item.reviewOnly;
export const isBranchDateSummaryEligible = (
  item: ItemResult,
): item is ItemResult & {
  appliedRate: string;
  adjustment: string;
  totalCost: string;
  grossMargin: string;
} =>
  showInItemBankDeductions(item) &&
  hasConfiguredBankPayment(item) &&
  item.appliedRate !== null &&
  item.adjustment !== null &&
  item.totalCost !== null &&
  item.grossMargin !== null &&
  !item.partialCalculation &&
  !!item.salesDate &&
  !!item.branch.trim();
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
export interface BranchDateSummary {
  date: string;
  branch: string;
  sales: string;
  totalCost: string;
  adjustment: string;
  grossMargin: string;
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
  branchDateSummary: BranchDateSummary[];
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
export const BRANCH_DATE_SUMMARY_TITLE = 'Expected Gross Margin Per Branch by Date';
export const BRANCH_DATE_SUMMARY_SHEET = 'Gross Margin by Branch & Date';
export const BRANCH_DATE_SUMMARY_COLUMNS = [
  { header: 'Date', key: 'date', format: 'date' },
  { header: 'Branch', key: 'branch', format: 'text' },
  { header: 'Total Credit Card Payments', key: 'sales', format: 'money' },
  { header: 'Total Cost / COGS', key: 'totalCost', format: 'money' },
  { header: 'Total Bank Charges / Subsidy', key: 'adjustment', format: 'money' },
  { header: 'Gross Margin', key: 'grossMargin', format: 'money' },
] as const;
export const ITEM_COLUMNS = [
  { header: 'Product ID', key: 'productId', format: 'text' },
  { header: 'Sales Date', key: 'salesDate', format: 'date' },
  { header: 'Branch', key: 'branch', format: 'text' },
  { header: 'Invoice Number', key: 'invoice', format: 'text' },
  { header: 'Product Name', key: 'product', format: 'text' },
  { header: 'Quantity', key: 'quantity', format: 'number' },
  { header: 'Price', key: 'price', format: 'money' },
  { header: 'Cost', key: 'totalCost', format: 'money' },
  { header: 'Discount', key: 'discount', format: 'money' },
  { header: 'Line Total', key: 'lineTotal', format: 'money' },
  { header: 'Payment', key: 'originalPaymentMethod', format: 'text' },
  { header: 'Applied Bank Rate 1', key: 'appliedRate1', format: 'percent' },
  { header: 'Applied Bank Rate 2', key: 'appliedRate2', format: 'percent' },
  { header: 'Applied Bank Rate 3', key: 'appliedRate3', format: 'percent' },
  { header: 'Total Bank Charges / Subsidy', key: 'adjustment', format: 'money' },
  { header: 'GrossMargin', key: 'grossMargin', format: 'money' },
  { header: 'Sales Source Row', key: 'sourceRow', format: 'number' },
  { header: 'Tender', key: 'cardName', format: 'text' },
  { header: 'Calculation Explanation', key: 'explanation', format: 'text' },
] as const;
