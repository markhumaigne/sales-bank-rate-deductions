# Sales Bank Rate Deduction Calculator

A local Angular + NestJS application that reads **one Sales Report**, calculates product bank charges, subsidies and gross margins, and exports a formatted four-sheet workbook. Each source product remains one output row. The main worksheet has exactly 17 columns, beginning with Product ID. Configuration stays in local JSON and the current browser session. No database or external storage is required.

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

1. **Item Bank Deductions** — one row per original product, including unresolved rows, in source order. No extra subtotal rows.
2. **Bank & Branch Summary** — validated sales, charges, subsidies, gross margin, processed item counts and review counts by bank, method and source branch, plus grand totals.
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

Payment copies the original Sales Report Payment cell verbatim, including line breaks. Tender is the matched configured card name. Bank and Status are omitted from the main worksheet and its preview; review reasons remain in Calculation Explanation and Needs Review. There is no separate payment-report input or cross-report invoice matching.

Product IDs are literal text in Column A. Headers are bold, frozen and filtered. Currency is PHP, rates are percentages and dates display as 01 Oct 2026. Review rows are highlighted. Unresolved amounts remain blank. Rows without a configured card match, including cash-only rows, also keep Applied Bank Rate, Bank Charges / Subsidy and GrossMargin blank.

## Verified Sales Report structure

The supplied Sales Report.xlsx was inspected and independently checked with ExcelJS against the SheetJS parser:

- Product ID, Product Name, Date, TM-INV#, Qty, Price, Cost, Discount, Net Total, Payment and Branch are detected.
- **Net Total** is the final line amount. When both Line Total and Net Total exist, Net Total takes priority. Discounts are already included and are never subtracted again.
- **Cost is unit cost** in this reference structure: repeated SKUs retain the same cost at quantities 1, 2, 5 and 8. Gross margin uses Cost × Quantity. The output Cost column preserves the original unit value; the explanation states the normalized total cost.
- Repeated **Tender / Ref / Amount** column groups inside the Sales Report supply explicit invoice payment amounts. Repetition across product rows is checked and counted once. The Payment column remains unchanged in the output.
- Text Payment values with one method/amount per line can also provide explicit invoice payments, for example “BDO 3 MOS 20,000.00” and “Cash 10,000.00”.
- A simple fully identified Payment value without amounts applies to that product's entire net total.
- The reference contains **2,202 rows**, all dated **October 1, 2026**. All 2,202 appear in the output with original source row numbers.
- With the supplied default settings, 305 items have validated fees and margins, and six additional items show matched-method charges and provisional gross margins. These six remain within the 1,746 items requiring review and are excluded from confirmed financial totals. Another 151 rows contain only configured no-charge methods; they remain in the output with all three calculated fields blank. SBC QRPH is explicitly excluded from bank charges as requested. Other unconfigured names such as Charge, unsupported payment methods/promotions, refunds and inconsistent payment evidence are retained for review.
- Verified default totals: validated sales on matched-card rows PHP 2,424,851.00; charges PHP 150,715.17; subsidies PHP 0.00; gross margin PHP 113,246.38. These describe only validated matched-card rows, not all source sales.

The real-file test independently compares every source Product ID, product, invoice, branch, payment, quantity, price, cost, discount, net total and date. It checks normalized costs, margins, aggregate fees for resolved invoice payments, all exported identifiers/amounts and the complete HTTP workflow. Private source workbooks are read in memory and are not copied into the repository.

## Calculation and review rules

- Decimal.js uses 40-digit precision and PHP amounts rounded to two decimals, ROUND_HALF_UP. JSON amounts are decimal strings.
- Rates are percentages: a configured 3.50 means 3.50%. Positive rates produce bank charges; negative rates produce subsidies. Default rates remain unchanged.
- Bank charge = eligible paid amount × configured rate. GrossMargin = net line total − total product cost − signed charge/subsidy. A negative subsidy increases margin.
- Exact configured names take priority, followed by explicit aliases, then unambiguous recognized spacing or configured wildcard patterns. Recognized installment spacing includes BDO 3MOS → BDO 3 MOS.
- The reference's payment descriptions can have a configured method followed by cardholder/reference text. Only recognized description fields allow this prefix interpretation. Original payment text is always retained.
- When Payment identifies the configured method shown in Tender, calculate the matched method's Applied Bank Rate, Bank Charges / Subsidy, and GrossMargin using the actual eligible payment amount and known product cost. A single-method payment description applies to the net line amount. Explicit splits allocate each stated card amount; another unconfigured method no longer prevents calculation of the matched portions.
- A configured card match is required for all three calculations. Cash-only and SBC QRPH-only rows preserve their source values with Applied Bank Rate, Bank Charges / Subsidy and GrossMargin blank, even when their sales and costs are valid. These rows contribute item counts only, not confirmed financial totals. A deliberately configured card with a 0% rate still qualifies and calculates a zero charge and a margin.
- If another payment remains unconfigured, the displayed charge covers only matched methods and GrossMargin is provisional after those calculated charges. Calculation Explanation identifies excluded portions, the row remains highlighted/in Needs Review, and its amounts are excluded from confirmed summary totals. Unknown fees are not declared to be zero. SBC QRPH is explicitly configured with no bank charge, so a reconciled BDO STRAIGHT / SBC QRPH split is fully calculated.
- **BNPL/SNPL and promotional variants require explicit configuration**, as specified in the Sales-only update. Add a specific method, alias or promotional pattern; their labels are not stripped to guess a base rate.
- CASH and SBC QRPH are the default methods with no bank charge. In a split with a matched card, their amounts still participate in payment reconciliation and gross margin. Their portions never receive the card's rate, and these methods alone do not qualify a row for calculation. Manage this list under Configuration settings → Banks & branches → Methods with no bank charge. Older imported configurations must include SBC QRPH in that list to use the exclusion; reset defaults or add the name and re-export the settings. Other unknown methods remain unresolved.
- Explicit payment amounts default to the whole invoice. Grouping uses invoice, branch and calendar date from the Sales Report. Payment evidence must agree across those product rows and equal their net sales total. Different branches/dates are separate groups. Select line scope only when the source explicitly supplies per-product payment amounts.
- For known split payments, allocate each method's amount proportionally by product net sales using largest remainders in whole centavos. Round that payment's fee first, then allocate the fee using allocated item payments. Ties follow source row order. Zero-paid items receive zero fees.
- Multiple methods are combined into **one result row per product**. The explanation describes each allocated amount and fee. Multiple configured names/rates appear together rather than inventing one blended rate. Separate charge and subsidy totals remain visible in the summary.
- Rows without any matched card or configured no-charge method remain unresolved. Missing split amounts, inconsistent payment evidence, negative sales/refunds, invalid quantities or dates still require review before monetary calculation. Known names remain in Tender. A name match never fabricates payment amounts, rates or missing product costs.
- A missing/invalid cost leaves gross margin blank while preserving a calculable fee. Unit Cost and Total/Line Cost headers are recognized. A generic Cost outside the verified reference structure needs a basis choice for quantities other than one.
- Missing/invalid discount, price, branch or product details are reported. A supplied final net total can still support a valid fee without assuming the missing field is zero.
- Product ID detection supports Product ID, ProductID, SKU, Item Code and Product Code, plus manual mapping. Source text and recoverable displayed numeric formatting are retained, including leading zeros. No zero-padding is invented. Missing IDs stay blank with a separate warning and do not discard valid financial values. Repeated IDs/invoices never deduplicate products.
- Text dates use MM/DD/YYYY or ISO YYYY-MM-DD. Excel serial dates respect the workbook's 1900/1904 date system. Calendar dates are not shifted by browser timezone. Source dates are separate from generation timestamps.
- Monetary inputs accept comma groups, PHP/₱ prefixes and accounting parentheses. Actual fractions of a cent require review. Only numeric Excel storage noise within PHP 0.00000001 of a cent is normalized, with a note.
- Excel formulas are never executed. Cached source values can be read with a note; missing cached results and Excel errors require review.

Summary totals include only known validated amounts on matched-card rows. Rows containing only no-charge methods appear under Not calculated with counts and no financial contribution; they do not require review unless source data has issues or warnings. Partial calculations with unconfigured payment portions appear on Sheet 1 but are excluded from confirmed summary totals. A fully matched fee with missing cost contributes to charges but not gross margin. Missing-ID warnings preserve valid amounts. Review counts count source items, including warnings, once. Bank-specific item counts can overlap when a product has multiple banks; the grand count is the original source row count.

## Manage configuration

All 50 default rates, nine banks and 39 branches remain in backend/src/config/bank-rates.json and branches.json. The Angular editor supports editing names/rates, adding/removing methods, banks and branches, aliases, patterns and non-card methods. Settings are session-only unless exported and reimported.

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
