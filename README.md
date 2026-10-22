# Sales Bank Rate Deduction Calculator

A local Angular + NestJS application that reads **one Sales Report**, calculates product bank charges, subsidies and gross margins, and exports a formatted four-sheet workbook. Each product appears once in Item Bank Deductions, with negative quantities retained in Needs Review. Bank rates use actual Payment amounts; invoice payments are counted once. The main worksheet has exactly 19 columns, beginning with Product ID. Configuration stays in local JSON and the current browser session. No database or external storage is required.

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
2. Select its worksheet. Inspect the first ten rows, detected headers and row count. Override column mappings or the header row when necessary. Duplicate headings require an explicit source-column choice before calculation. Each mapped field must use a separate column.
3. Check **Cost basis** and **Explicit payment amounts apply to**. The supplied report is recognized as unit cost with invoice payment amounts repeated across products.
4. Review rates, banks, branches, aliases, patterns and non-card methods.
5. Click **Calculate Bank Rate Deductions**.
6. Review all 19 columns, search by Product ID/invoice/product, and click any column header to sort. Inspect Needs Review reasons.
7. Click **Download Calculated Excel Report**.

Editing the file selection, mappings, cost basis or configuration clears the previous calculation. The download is available only after successful calculation. Financial results are snapshots: editing settings in Excel requires importing those settings and calculating again.

## Workbook

The filename is Sales_Bank_Rate_Deductions_YYYYMMDD_HHMMSS.xlsx, with generation time in Asia/Manila. Its four sheets are:

1. **Item Bank Deductions** — one row per product with a configured bank payment, in source order. Payments starting with Cash or Charge, rows with no configured bank match, and negative quantities appear in Needs Review. Unresolved matched amounts remain blank with visible reasons. No extra subtotal rows.
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
| L | Applied Bank Rate 1 |
| M | Applied Bank Rate 2 |
| N | Applied Bank Rate 3 |
| O | Total Bank Charges / Subsidy |
| P | GrossMargin |
| Q | Sales Source Row |
| R | Tender |
| S | Calculation Explanation |

Payment copies the original Sales Report Payment cell verbatim, including line breaks. Tender is the matched configured card name. Original BNPL/SNPL Tender descriptions also appear in Calculation Explanation, including when they came from separate Tender columns inside the Sales Report. Bank and Status are omitted from the main worksheet and its preview; review reasons remain in Calculation Explanation and Needs Review. There is no separate payment-report input or cross-report invoice matching.

Product IDs are literal text in Column A. Headers are bold, frozen and filtered. Currency is PHP, rates are percentages and dates display as 01 Oct 2026. Review rows are highlighted. Unresolved amounts remain blank. Item Bank Deductions uses the match already produced by the configured payment matcher, including aliases, patterns and BNPL/SNPL handling. A zero rate or subsidy still qualifies; a positive charge is not required. All source items remain in the calculation audit and processed-item count, while the Item Bank Deductions tab counts displayed product rows. Missing, malformed and ambiguous payment details remain accessible in Needs Review.

## Verified Sales Report structure

The supplied Sales Report.xlsx was inspected and independently checked with ExcelJS against the SheetJS parser:

- Product ID, Product Name, Date, TM-INV#, Qty, Price, Cost, Discount, Net Total, Payment and Branch are detected.
- **Net Total** is the final line amount. When both Line Total and Net Total exist, Net Total takes priority. Discounts are already included and are never subtracted again.
- **Cost is unit cost** in this reference structure: repeated SKUs retain the same cost at quantities 1, 2, 5 and 8. Gross margin uses Cost × Quantity. Sheet 1's Cost column displays that calculated total product cost, matching Sheet 2. Already extended line costs are not multiplied again. Original source costs remain in the calculation audit, Calculation Explanation and exported Cost cell notes.
- Repeated **Tender / Ref / Amount** column groups inside the Sales Report supply explicit invoice payment amounts. Repetition across product rows is checked and counted once. The Payment column remains unchanged in the output.
- Text Payment values with one method/amount per line can also provide explicit invoice payments, for example “BDO 3 MOS 20,000.00” and “Cash 10,000.00”.
- A method name without an actual paid amount is unresolved unless explicit payment-amount columns in the same Sales Report supply the amount. Line Total is never used as a payment fallback.
- The reference contains **2,202 rows**, all dated **October 1, 2026**. All are processed; exported matched items retain their original source row numbers even when omitted rows leave gaps.
- BNPL/SNPL transactions qualify when the underlying payment method matches the settings. Cash/Charge-first transactions and rows without a configured bank match go to Needs Review. A bank-first payment may contain later Cash or unconfigured SBC QRPH portions; those portions receive zero deductions. Refunds, malformed amounts, ambiguous names and inconsistent payment evidence remain for review.
- Financial totals and review counts must be regenerated with the current matching rules. The private reference workbook is not bundled with the repository; run the real-file test below to verify its current totals.

The real-file test independently compares every source Product ID, product, invoice, branch, payment, quantity, price, cost, discount, net total and date. It checks normalized costs, margins, aggregate fees for resolved invoice payments, all exported identifiers/amounts and the complete HTTP workflow. Private source workbooks are read in memory and are not copied into the repository.

## Calculation and review rules

- Decimal.js uses 40-digit precision. Currency uses two decimals, ROUND_HALF_UP; JSON amounts are decimal strings. Rates are percentages: 2.85 means 2.85%. Positive rates are charges; negative rates are subsidies.
- **Each actual payment amount × its own matched rate** is rounded first, then allocated to products. Percentages are never combined and applied to Line Total. For BDO STRAIGHT ₱44,990 at 2.85% plus unconfigured SBC QRPH ₱3,000, the invoice fee is **₱1,282.22**.
- Payment accepts multiple methods on the same line or on separate lines (also semicolon or pipe separators), comma-grouped currency, PHP/₱ prefixes and integer or centavo amounts. Each explicit monetary amount ends its payment description; inline plain integers require a complete known next payment method so installment terms and card identifiers are not treated as payments. Card identifiers are retained as deduplication references and excluded from matching; customer text can follow a complete configured method name. Missing, negative, malformed, ambiguous or fractional-cent amounts remain unresolved.
- For example, `CC#2700 METROBANK 12 MOS SAMPLE CUSTOMER 39,999.00 BDO STRAIGHT 10,000.00` uses METROBANK on ₱39,999 and BDO on ₱10,000. With default rates of 10.50% and 2.85%, the charges are ₱4,199.90 and ₱285.00, totaling ₱4,484.90. A combined Tender/Amount value elsewhere in the Sales Report cannot replace these explicit Payment amounts. Three or more inline payments follow the same rule, including negative configured subsidy rates.
- Exact names, explicit aliases, recognized spacing (including installment terms and hyphens), and unambiguous configured patterns are supported. Description prefixes must contain the complete payment method; debit, straight, installment and combined-method suffixes cannot be confused. Wildcard patterns cannot substitute a different recognized bank or payment type. BNPL/SNPL handling is preserved; explicit configured variants take priority.
- Payment text with amounts is authoritative. Existing mapped payment-amount columns and Tender/Ref/Amount groups in the **same Sales Report** remain supported when Payment only supplies names. There is no Tender Report input and no inferred payment amount.
- Unconfigured methods have **zero charges**, with no rate substituted. Cash and Cash Payment always have zero bank deductions. Matched and unconfigured portions can share a fully calculated invoice. The first three matched payment rates fill columns L–N in payment order; all subsequent payments still contribute charges and appear in Calculation Explanation.
- **Payment selection:** if the first word of Payment is Cash or Charge (ignoring leading whitespace and casing), move the row to Needs Review, even when a later payment matches a configured bank. Shared invoice payments are held together if any product has this prefix, with an explicit explanation and no financial contributions until reviewed; independent product-specific payments remain independent. Rows with no configured bank match also go to Needs Review. Bank-first entries with later cash/unconfigured portions still calculate only the configured portions. Angular preview, Sheet 1 and Sheet 2 share this selection rule.
- Invoice keys include normalized Branch and Invoice Number, Sales Date, source file and worksheet. Missing invoice, branch or date prevents financial allocation. Repeated lists count once; repeated equal entries within a single list retain their multiplicity. Distinct card references remain distinct. Details present on only one product, reordered lists and repeated subsets are supported. Conflicting combined evidence that cannot reconcile to net invoice sales is flagged visibly.
- Whole-invoice scope is the default. Product-specific scope is available only for reports with reliable per-product paid amounts. Product rows are never deduplicated.
- Each payment amount and its rounded signed charge are allocated in proportion to **positive product Line Total**, using largest remainders in whole centavos. Ties follow ascending source row. Zero-Line-Total products receive zero payment and zero charge. Allocated fees reconcile exactly to each rounded invoice fee.
- Existing cash-change behavior remains for eligible bank-first transactions: for a complete, recognized card/cash payment list, excess tender fully covered by cash is cash change. Actual card amounts and fees stay unchanged. Explanations show cash tender, change and net cash. Inconsistent partial lists cannot fabricate change.
- Calculation Explanation shows invoice, each actual payment, matched rate (or no configured rate), individual fee, total payments, total signed fees, allocation basis, row allocation and margin. Malformed/unreconciled amounts are visible in both explanations and Needs Review.
- For invoices with multiple products, the explanation starts with this product's allocated charge and identifies the Sales Source Rows receiving the invoice charges. The charge cell also carries this explanation as an Excel note. For example, a repeated BDO STRAIGHT ₱7,400 payment at 2.85% produces one ₱210.90 invoice charge: a free accessory receives ₱0.00 while the positive-value product receives the fee. The customer's actual card payment remains ₱7,400 on both rows; the product allocation is labelled separately.
- Product Price, Discount and Line Total retain their sales meaning. GrossMargin remains **net Line Total − total product cost − signed allocated charges**. Discount is already included in net sales and is never deducted twice; a subsidy increases margin. Unit cost is multiplied by valid quantity; extended line cost is used once. Sheet 1 Cost shows total product cost; original cost and basis remain in notes and explanations.
- No-card rows have blank GrossMargin. Cash/Charge-first rows have pending charges and margin in Needs Review; unconfigured-only rows receive no deduction. Unknown costs leave margin blank while an eligible bank fee can still calculate. Sheet 2 retains its six columns and sums the eligible Sheet 1 sales, total costs, allocated signed fees and margins once, without recalculating rates. Explanations identify Included in Sheet 2 / Excluded from Sheet 2. Missing costs, invalid dates/branches, refunds and negative quantities require review. Missing Product ID warnings alone do not exclude valid financial calculations.
- Product IDs remain literal text with recoverable leading zeros. Dates preserve source calendar dates (MM/DD/YYYY or ISO; Excel serials respect 1900/1904). Formula results are read only from cached source values; formulas are never executed. Only tiny numeric storage noise within PHP 0.00000001 of a cent is normalized; true fractions of a cent remain invalid.

Payment methods may optionally include branch-specific overrides in the JSON configuration:

```json
{"bank":"BDO","cardName":"BDO STRAIGHT","rate":"2.85","branchRates":{"ASUS":"2.50"}}
```

Branch names use the same case/space normalization as invoices and must exist in the branch reference list. An exact branch override takes priority; other branches use the method's base rate. Overrides survive Excel settings export/import and frontend configuration edits; the four-column settings layout remains unchanged. The existing bank/branch reference lists do not assign rates by row position.

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

ExcelJS writes all report text as literal string cells (never formula objects), preserving identifiers, branches, product names and Payment text exactly, including leading formula-like characters. Branch cells containing Excel errors require review and cannot contribute bank charges or branch totals. Imported configuration formulas are rejected.

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

Browser tests use isolated API/frontend ports 3100/4300 and close their servers after completion. They exercise worksheet selection, manual mappings, all 19 columns, search/sort, editable rates, Excel configuration import/export, download and mobile layout. Real-data screenshots, traces and videos are disabled. Screenshots contain only empty or synthetic test data.

Backend tests include the one/two/three/four-payment accounting examples, unconfigured zero charges, masked identifiers, missing amounts, invoice deduplication/subsets, positive/zero sales allocations, branch overrides, currency rounding and HTTP/export reconciliation. They also cover BDO/BPI calculations, individual products, duplicate IDs/invoices, cash, splits, multiple banks, unknown/promotional methods, missing cost/IDs, discounts, unit/line cost, subsidies, rounding, date/source-row preservation, numeric zero masks, .xls, exact export columns, formula safety and signed download validation.

## Project layout

```text
frontend/src/app/
  components/excel-upload/          One file, worksheet, mapping and source preview
  components/configuration/         Session configuration editor
  components/calculation-results/   Nineteen columns, summaries and reviews
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
