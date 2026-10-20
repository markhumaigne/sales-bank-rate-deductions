# Sales Bank Rate Deduction Calculator

A local Angular + NestJS application that reads **one Sales Report**, calculates product bank charges, subsidies and gross margins, and exports a formatted four-sheet workbook. Each product with a configured bank-payment match appears once in Item Bank Deductions. The main worksheet has exactly 17 columns, beginning with Product ID. Configuration stays in local JSON and the current browser session. No database or external storage is required.

## Install and run

Use Node.js 22.12+ and npm. From this repository:

```powershell
npm install
npm start
```

Open **http://127.0.0.1:4200**. The API runs at **http://127.0.0.1:3000**. Stop with Ctrl+C.

To run separately:

```powershell
# Terminal 1
npm run start -w backend

# Terminal 2
npm run start -w frontend
```

Restart the backend after changing backend code or JSON defaults. If an old session is running, stop it with Ctrl+C before starting again; do not start a second instance on the same ports. Reload the browser and upload the Sales Report again after updating.

```powershell
npm run build
```

Build output: backend/dist and frontend/dist/browser. Run a compiled API with npm run start:prod -w backend. A separately hosted frontend must forward /api to the local API; npm start handles this automatically for development.

## Use the calculator

1. Upload a Sales Report (.xlsx or .xls, up to 10 MB).
2. Select its worksheet. Inspect the first ten rows, detected headers and row count. Override column mappings or the header row when necessary.
3. Check **Cost basis** and **Explicit payment amounts apply to**. The supplied report is recognized as unit cost with invoice payment amounts repeated across products.
4. Review rates, banks, branches, aliases, patterns and non-card methods.
5. Click **Calculate Bank Rate Deductions**.
6. Review all 17 columns, search by Product ID/invoice/product, and click any column header to sort. Inspect Needs Review reasons.
7. Click **Download Calculated Excel Report**.

Editing the file selection, mappings, cost basis or configuration clears the previous calculation. The download is available only after successful calculation. Financial results are snapshots: editing settings in Excel requires importing those settings and calculating again.

## Workbook

The filename is Sales_Bank_Rate_Deductions_YYYYMMDD_HHMMSS.xlsx, with generation time in Asia/Manila. Its four sheets are:

1. **Item Bank Deductions** — one row per product with a matched configured payment method in Tender, in source order. Cash-only, other no-charge-only, missing-payment and wholly unmatched products are omitted from this sheet and its preview. Matched products with unresolved amounts remain visible for review. No extra subtotal rows.
2. **Gross Margin by Branch & Date** — headed **Expected Gross Margin Per Branch by Date**, with exactly six columns: Date, Branch, Total Credit Card Payments, Total Cost / COGS, Total Bank Charges / Subsidy, and Gross Margin. Fully calculated matched-payment products are grouped by calendar date and branch, sorted by date then branch, with peso amounts and grand totals.
3. **Configuration Settings** — the exact editable settings snapshot used.
4. **Needs Review** — Product ID, source row, reason, source file/sheet and original parsed values for every flagged item.

The main sheet has these exact columns:

| Column | Name |
| --- | --- |
| A | Product ID |
| B | Sales Date |
| C | Branch |
| D | Invoice Number |
| E | Product Name |
| F | Quantity |
| G | Price |
| H | Cost |
| I | Discount |
| J | Line Total |
| K | Payment |
| L | Applied Bank Rate |
| M | Bank Charges / Subsidy |
| N | GrossMargin |
| O | Sales Source Row |
| P | Tender |
| Q | Calculation Explanation |

Payment copies the original Sales Report Payment cell verbatim, including line breaks. Tender is the matched configured card name. Original BNPL/SNPL Tender descriptions also appear in Calculation Explanation, including when they came from separate Tender columns inside the Sales Report. Bank and Status are omitted from the main worksheet and its preview; review reasons remain in Calculation Explanation and Needs Review. There is no separate payment-report input or cross-report invoice matching.

Product IDs are literal text in Column A. Headers are bold, frozen and filtered. Currency is PHP, rates are percentages and dates display as 01 Oct 2026. Review rows are highlighted. Unresolved amounts remain blank. Item Bank Deductions uses the match already produced by the configured payment matcher, including aliases, patterns and BNPL/SNPL handling. A zero rate or subsidy still qualifies; a positive charge is not required. All source items remain in the calculation audit and processed-item count, while the Item Bank Deductions tab counts only displayed matches. Missing and unmatched payments remain accessible in Needs Review.

## Verified Sales Report structure

The supplied Sales Report.xlsx was inspected and independently checked with ExcelJS against the SheetJS parser:

- Product ID, Product Name, Date, TM-INV#, Qty, Price, Cost, Discount, Net Total, Payment and Branch are detected.
- **Net Total** is the final line amount. When both Line Total and Net Total exist, Net Total takes priority. Discounts are already included and are never subtracted again.
- **Cost is unit cost** in this reference structure: repeated SKUs retain the same cost at quantities 1, 2, 5 and 8. Gross margin uses Cost × Quantity. Sheet 1's Cost column displays that calculated total product cost, matching Sheet 2. Already extended line costs are not multiplied again. Original source costs remain in the calculation audit, Calculation Explanation and exported Cost cell notes.
- Repeated **Tender / Ref / Amount** column groups inside the Sales Report supply explicit invoice payment amounts. Repetition across product rows is checked and counted once. The Payment column remains unchanged in the output.
- Text Payment values with one method/amount per line can also provide explicit invoice payments, for example “BDO 3 MOS 20,000.00” and “Cash 10,000.00”.
- A simple fully identified Payment value without amounts applies to that product's entire net total.
- The reference contains **2,202 rows**, all dated **October 1, 2026**. All are processed; exported matched items retain their original source row numbers even when omitted rows leave gaps.
- BNPL/SNPL transactions qualify when the underlying payment method matches the settings. Cash has zero bank deductions; SBC QRPH remains explicitly excluded from bank charges. Unconfigured names such as Charge, unsupported promotions, refunds and inconsistent payment evidence remain for review.
- Financial totals and review counts must be regenerated with the current matching rules. The private reference workbook is not bundled with the repository; run the real-file test below to verify its current totals.

The real-file test independently compares every source Product ID, product, invoice, branch, payment, quantity, price, cost, discount, net total and date. It checks normalized costs, margins, aggregate fees for resolved invoice payments, all exported identifiers/amounts and the complete HTTP workflow. Private source workbooks are read in memory and are not copied into the repository.

## Calculation and review rules

- Decimal.js uses 40-digit precision and PHP amounts rounded to two decimals, ROUND_HALF_UP. JSON amounts are decimal strings.
- Rates are percentages: a configured 3.50 means 3.50%. Positive rates produce bank charges; negative rates produce subsidies. Defaults use the 49-method rate table supplied on October 20, 2026.
- Bank charge = eligible paid amount × configured rate. GrossMargin = net line total − total product cost − signed charge/subsidy. A negative subsidy increases margin.
- Exact configured names take priority, followed by explicit aliases, then unambiguous recognized spacing or configured wildcard patterns. Recognized installment spacing includes BDO 3MOS → BDO 3 MOS.
- BDO PAY is a default alias of BDO PAY/CASA and uses its current configured rate (3.00% in the supplied defaults). Calculation fills in a missing alias when BDO PAY/CASA exists and saves the effective mapping in the report's Configuration Settings. Explicit methods, aliases, patterns and no-charge settings take priority; missing targets and ambiguous rules remain unresolved. In a BDO STRAIGHT / BDO PAY split, each portion uses its own configured rate (2.85% and 3.00% by default), while the original Payment text is preserved. Reset defaults to replace an older explicit BDO PAY → BDO STRAIGHT mapping.
- The reference's payment descriptions can have a configured method followed by cardholder/reference text. Only recognized description fields allow this prefix interpretation. Original payment text is always retained.
- When Payment identifies the configured method shown in Tender, calculate the matched method's Applied Bank Rate, Bank Charges / Subsidy, and GrossMargin using the actual eligible payment amount and known product cost. A single-method payment description applies to the net line amount. Explicit splits allocate each stated card amount; another unconfigured method no longer prevents calculation of the matched portions.
- A configured card match is required for bank calculations and gross margin. Cash and Cash Payment always have zero bank deductions, even if a name, alias or wildcard setting would otherwise assign a rate. Their rate and gross margin remain blank. SBC QRPH-only rows keep all three fields blank. These rows contribute item counts only, not bank-payment financial totals. A deliberately configured card with a 0% rate still qualifies and calculates a zero charge and a margin.
- If another payment remains unconfigured, the displayed charge covers only matched methods and GrossMargin is provisional after those calculated charges. Calculation Explanation identifies excluded portions, the row remains highlighted/in Needs Review, and its amounts are excluded from confirmed summary totals. Unknown fees are not declared to be zero. SBC QRPH is explicitly configured with no bank charge, so a reconciled BDO STRAIGHT / SBC QRPH split is fully calculated.
- **BNPL/SNPL classification labels are ignored for payment matching.** Explicit matching methods, aliases and label-specific patterns take priority; otherwise standalone BNPL/SNPL words are removed and the existing matching rules resolve the underlying method. Prices, costs, payment amounts and original text are preserved. Unmatched or ambiguous underlying methods remain flagged without an invented rate. Other promotional labels still require explicit configuration.
- CASH and CASH PAYMENT are always non-chargeable. SBC QRPH is a default configured no-charge method. In a split with a matched card, these amounts still participate in the existing payment allocation and gross margin, but never receive the card's rate. Manage additional no-charge methods under Configuration settings → Banks & branches → Methods with no bank charge. Older imported configurations must include SBC QRPH in that list to use its exclusion. Other unknown methods remain unresolved.
- Explicit payment amounts default to the whole invoice. Grouping uses invoice, branch and calendar date from the Sales Report. Payment evidence must agree across those product rows and equal their net sales total. Different branches/dates are separate groups. Select line scope only when the source explicitly supplies per-product payment amounts.
- For a configured bank payment combined with explicitly identified cash, excess tender is treated as cash change only when every payment method is recognized and the cash fully covers the excess. Bank amounts and rates stay unchanged; only net cash participates in allocation. For example, PAYMAYA DEBIT ₱6,000 plus cash tender ₱3,000 against ₱8,995 net sales gives ₱5 change, ₱2,995 net cash, and a ₱173.40 bank charge at the configured 2.89% rate. Original payment text and amounts are retained, and the explanation records the change. Card-only discrepancies, insufficient cash, unknown methods and invalid source values still require review.
- For known split payments, allocate each method's amount proportionally by product net sales using largest remainders in whole centavos. Round that payment's fee first, then allocate the fee using allocated item payments. Ties follow source row order. Zero-paid items receive zero fees.
- Multiple methods are combined into **one result row per product**. The explanation describes each allocated amount and fee. Multiple configured names/rates appear together rather than inventing one blended rate. Separate charge and subsidy totals remain visible in the dashboard; Sheet 2 sums their signed net adjustment.
- Rows without any matched card or configured no-charge method remain unresolved. Missing split amounts, inconsistent payment evidence, negative sales/refunds, invalid quantities or dates still require review before monetary calculation. Known names remain in Tender. A name match never fabricates payment amounts, rates or missing product costs.
- Items with negative quantities (such as -1, -2 or -0.5) are moved from Item Bank Deductions to Needs Review in both the preview and Excel export. They contribute no sales, cost, bank charges or gross margin to Sheet 2. Negative sales amounts with nonnegative quantities still remain visible when the payment is configured. Warnings and source details are preserved, with charges and margins blank pending review. A shared invoice payment stays on hold with the affected source rows identified; independent valid payments continue to calculate.
- The unknown-method warning for `Charge` and the invoice payment discrepancy message are hidden from report explanations and Needs Review text in the preview and Excel export. Item rows, review status, original diagnostics, negative-value warnings and financial validation remain unchanged; hiding these messages does not resolve a payment or calculate a missing fee.
- A missing/invalid cost leaves gross margin blank while preserving a calculable fee. Unit Cost and Total/Line Cost headers are recognized. A generic Cost outside the verified reference structure needs a basis choice for quantities other than one.
- Missing/invalid discount, price, branch or product details are reported. A supplied final net total can still support a valid fee without assuming the missing field is zero.
- Product ID detection supports Product ID, ProductID, SKU, Item Code and Product Code, plus manual mapping. Source text and recoverable displayed numeric formatting are retained, including leading zeros. No zero-padding is invented. Missing IDs stay blank with a separate warning and do not discard valid financial values. Repeated IDs/invoices never deduplicate products.
- Text dates use MM/DD/YYYY or ISO YYYY-MM-DD. Excel serial dates respect the workbook's 1900/1904 date system. Calendar dates are not shifted by browser timezone. Source dates are separate from generation timestamps.
- Monetary inputs accept comma groups, PHP/₱ prefixes and accounting parentheses. Actual fractions of a cent require review. Only numeric Excel storage noise within PHP 0.00000001 of a cent is normalized, with a note.
- Excel formulas are never executed. Cached source values can be read with a note; missing cached results and Excel errors require review.

Sheet 2 sums the existing net Line Total, Cost, signed Bank Charges / Subsidy and GrossMargin once per eligible Sheet 1 product. Cost in both sheets is the calculated total product cost, not a unit cost. Calculation Explanation marks each matched row **Included in Sheet 2** or **Excluded from Sheet 2** so the same rows can be summed directly across both sheets. All four totals use the same products. Negative quantities, cash-only and other no-charge-only rows, unmatched payments, provisional calculations, and products missing cost, date or branch are excluded. Matched items remain on Sheet 1 except negative quantities, which appear only in Needs Review. Missing/unmatched payments and other review reasons remain in Needs Review. The original Sales Report retains all source rows. Missing-ID warnings do not exclude otherwise valid amounts. A configured bank transfer qualifies through the same payment-method matching rules; an unconfigured transfer does not. Reconciled card/cash splits retain the full product sales and cost used by the existing individual gross margin; only the configured card portion incurs a fee. Branch grouping uses the same case/spacing normalization as invoice processing. No fee, subsidy or discount is deducted again.

The dashboard retains its existing totals of known validated amounts: a fully matched fee with missing cost still contributes to dashboard charges, but that incomplete product is excluded from Sheet 2. Review counts include source items and warnings once. Existing settings assign rates by payment method and bank; the branch reference list does not define separate rate overrides.

## Manage configuration

The 49 default payment methods, nine banks and 39 branches are stored in backend/src/config/bank-rates.json and branches.json. The payment-method list matches the supplied rate table; methods omitted from that table are no longer defaults. The Angular editor supports editing names/rates, adding/removing methods, banks and branches, aliases, patterns and non-card methods. Settings are session-only unless exported and reimported. Restart the backend and use **Reset defaults** (or reload the browser) to load updated defaults; imported settings preserve their saved rates and explicit aliases.

Use **Export settings**, or edit Configuration Settings from a calculated report:

- Keep the sheet name and first-row headers.
- The sheet has exactly four columns: **A Card Name, B Rate, C Bank Reference, D Branch**.
- Card Name and Rate form a pair. Bank Reference and Branch are independent lists; their row positions never assign a bank or branch to a card.
- Cell notes preserve explicit card-bank assignments, aliases, patterns and non-card methods for a complete configuration round trip. Keep those notes when editing/copying the worksheet. Manage these additional settings in the application's configuration editor.
- Renaming an existing card or bank in the exported workbook updates its saved relationships on import. Add custom payment methods with their bank assignments in the application, then export settings.
- Previously exported ten-column configurations remain importable. A plain four-column file without notes can recover bank assignments only for exact default card names; unknown assignments produce a readable error instead of guessing from row positions.
- Enter literal percentages, such as 3.50% or -2.00% for a subsidy. Excel stores these as 0.035 and -0.02. Text percentages are also accepted.
- Banks must exist in the bank list; alias/pattern targets must exist in the card list. Duplicates and conflicting names are rejected.
- Configuration formulas are rejected. Import, review and recalculate.

## API

| Endpoint | Request | Response |
| --- | --- | --- |
| GET /api/config/defaults | None | Default JSON |
| POST /api/excel/preview | Multipart file; optional headerRow query | Sheets, mappings, cost basis, counts, first ten rows |
| POST /api/excel/calculate | Multipart sales and JSON-string options | result and exportToken |
| POST /api/excel/export | Unchanged result and exportToken | Four-sheet Excel file |
| POST /api/config/import | Multipart file | Validated settings |
| POST /api/config/export | Configuration JSON | Settings workbook |

Calculation options contain sales: {sheet, headerRow, mapping, costBasis, paymentScope} and configuration. Mapping keys: productId, date, branch, invoice, product, quantity, price, cost, discount, lineTotal, payment, paymentAmount. Required mappings are invoice, product, quantity and lineTotal; missing optional data stays in the output with appropriate review/warnings. Cost basis is auto/unit/line; payment scope is invoice/line.

Export is signed with a random in-memory HMAC key. Modified results or results from a previous API process must be recalculated. No result database or uploaded-file store is used.

## Limits and data handling

Uploads and reports stay in memory. The browser retains the session until reload. Only an explicit download saves a file. Responses use Cache-Control: no-store. No transaction localStorage or cookies are used.

Limits: one file per request, 10 MB/file, 30 worksheets, 50,000 rows, 200 columns, 100 MB expanded ZIP content, 2,000 ZIP entries and 24 MB serialized result. Password-protected files are unsupported. Only genuine Excel ZIP/OLE content is accepted. Split larger reports into batches.

ExcelJS writes Product IDs and original Payment text as literal string cells (never formula objects), keeping their values exact and inert. Other source strings beginning with formula-like characters are escaped. Imported configuration formulas are rejected.

## Tests

```powershell
npm test
npm run build

# Real Sales Report: source comparisons, financial checks, export and HTTP workflow
npm run test:sample -- --sales "C:/Users/Daveee/Downloads/Sales Report.xlsx"

# Desktop/mobile browser tests; installed Google Chrome required
npm run test:e2e

# Include the actual Sales Report in browser upload/calculate/download tests
$env:SAMPLE_SALES="C:/Users/Daveee/Downloads/Sales Report.xlsx"
npm run test:e2e
```

Browser tests use isolated API/frontend ports 3100/4300 and close their servers after completion. They exercise worksheet selection, manual mappings, all 17 columns, search/sort, editable rates, Excel configuration import/export, download and mobile layout. Real-data screenshots, traces and videos are disabled. Screenshots contain only empty or synthetic test data.

Backend tests cover BDO/BPI calculations, individual products, duplicate IDs/invoices, cash, splits, multiple banks, unknown/promotional methods, missing cost/IDs, discounts, unit/line cost, subsidies, rounding, date/source-row preservation, numeric zero masks, .xls, exact export columns, formula safety and signed download validation.

## Project layout

```text
frontend/src/app/
  components/excel-upload/          One file, worksheet, mapping and source preview
  components/configuration/         Session configuration editor
  components/calculation-results/   Seventeen columns, summaries and reviews
  components/dashboard/             Upload → configure → calculate → download
  services/                         HttpClient API and downloads
backend/src/
  excel/                            SheetJS parser, endpoints and ExcelJS exporter
  calculation/                      Decimal.js calculations, allocations and matching
  config/                           Local JSON defaults and validation
shared/models.ts                    API contracts and exact output column definitions
backend/test/                       Unit/integration and actual-report checks
e2e/                                Browser workflow tests
sample-data/                        Instructions; private files are not bundled
```
