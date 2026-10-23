# Phase 2A: verified asset configuration

> Phase 3A migration notice: current deployment scope is the repository-root static allowlist. Paths and release settings in this Phase 2A snapshot are historical; see [Phase 3A](phase-3a-migration.md).

> Historical snapshot. Phase 2C.1 replaced the blocked Business Lead assets with verified synthetic demonstration assets and preserved originals outside the public root. Current status and provenance: [Phase 2C.1 safety record](phase-2c1-safety.md). Earlier blocker/default/path statements below describe the Phase 2A state.

Prepared 2026-10-23. Configuration only; no viewer, page, homepage, stylesheet, media, workbook, or deployment edits.

## Configuration contract

`website/assets/js/projects-config.js` defines `globalThis.PROJECTS_CONFIG`. It performs no fetching or rendering and is not yet loaded by the homepage. Load it once before the future viewer script.

- `projects` contains all seven projects. Each has separate `raw` and `output` groups.
- `sheets` is the sole source of clickable worksheet tabs. Every entry has an actual worksheet name, readable label, verified image path, width, height, byte size, width/height aspect ratio, shape and coverage note.
- Array order follows the workbook worksheet order after intentional exclusions, missing screenshots and safety exclusions. The first included screenshot is `defaultWorksheet`.
- `workbook` contains an approved filename, deployment-root path and download availability, or is null when absent/blocked. Workbooks never become screenshot tabs.
- Missing/blocked worksheet records are diagnostics, never placeholder tabs. Summary and README are absent from the runtime configuration.
- PDF raw input uses `documents` and `defaultDocument`, not image tabs.
- All asset paths start at `/` and assume Vercel's Root Directory is `website/`. Hosting beneath a URL subdirectory would require a later path strategy change.
- Shape thresholds: width > 1600 is wide; height > 1600 is tall. A Standard image can still require scrolling on a phone.
- `viewerDefaults` documents future requirements only: top-left origin, reset on tab change, 100% scale, preserved aspect ratio, zoom/reset/fit and fullscreen-or-expanded actions. No controls are implemented.
- `deploymentReady: false` is advisory metadata, not an access-control or deployment mechanism. Consumers must not turn blocked/missing diagnostics into image or download URLs.

## Public safety and deployment blockers

**Deployment is blocked pending review of both files below.** Configuration omission does not protect files already inside the public deployment directory.

| Exact repository-relative path | Location / risk | Handling |
|---|---|---|
| `website/business-lead-scraper/output/business_leads_scrape_07.01_1.xlsx` | `Business_Leads!N13` contains token-bearing form URLs. Whether these are confidential credentials or public identifiers is unverified. High release risk. | Workbook path and download omitted from runtime configuration. Original unchanged. |
| `website/business-lead-scraper/output/business-lead-output.png` | Screenshot of the affected worksheet, including URL-bearing fields; public safety cannot be established from the preview. Potential disclosure risk. | Screenshot omitted from runtime configuration. Original unchanged. |

Recommended remediation: obtain approval to replace these with reviewed public-safe demonstration copies, or remove them from the deployment directory. Preserve originals outside the public root. Do not silently sanitize originals, hide a button, or assume configuration omission blocks direct HTTP access.

The safe Business Lead output preview currently includes only Invalid_URLs; it has no workbook download. Its default is therefore Invalid_URLs. Business_Leads remains a verified mapping but is not an approved public tab.

All website Excel inputs/outputs were inspected read-only, including cells and XML relationship data; no other credential assignments, embedded files or macros were detected. PNG embedded metadata had no credential-pattern hits. The four PDFs explicitly identify themselves as generated testing samples and have no detected attachment catalog. Prior visual inspection of the unchanged PNGs supplies content/header mapping evidence, not filename guessing.

The user explicitly confirmed that the cleaning and sales customer workbooks are synthetic and public-safe. Their matching screenshot records are treated as the same approved demonstration data. Business input URLs and the invalid-URL preview contain public website references. Catalog/PrestaShop assets use demonstration products and domains. SEO assets contain public site information. Existing professional contact details and branding are retained.

Some approved output workbook summaries expose local machine paths and some workbook/image metadata contains authoring identifiers. No credential values were found there; those details remain unchanged and are a low-risk residual metadata disclosure. This inspection does not establish an unrestricted guarantee about third-party contact data.

## Verified worksheet mappings

PNG filenames below are relative to the listed project's raw/output directory. The configured workbook path, or the blocked source workbook identified above, defines ownership.

| Project | Group | Worksheet | PNG | Tab label | Status |
|---|---|---|---|---|---|
| business-lead-scraper | raw | `Websites_Raw` | `business-lead-scraper-raw.png` | Source Websites | Confirmed; included |
| business-lead-scraper | output | `Business_Leads` | `business-lead-output.png` | Business Leads | Blocked pending safety review |
| business-lead-scraper | output | `Invalid_URLs` | `business-lead-invalid-url.png` | Invalid URLs | Confirmed; included |
| clean-and-validate-excel | raw | `Orders_Raw` | `clean-and-validate-excel-raw.png` | Raw Orders | Confirmed; included |
| clean-and-validate-excel | output | `Customers_Valid` | `clean-customers-valid.png` | Valid Customers | Confirmed; included |
| clean-and-validate-excel | output | `Customers_Invalid` | `clean-customers-invalid.png` | Invalid Customers | Confirmed; included |
| clean-and-validate-excel | output | `Orders_Valid` | `clean-orders-valid.png` | Valid Orders | Confirmed; included |
| clean-and-validate-excel | output | `Orders_Invalid` | `clean-orders-invalid.png` | Invalid Orders | Confirmed; included |
| product-catalog-auditor | raw | `Products_Raw` | `product-catalog-raw.png` | Raw Catalog | Confirmed; included |
| product-catalog-auditor | output | `Products_Valid` | `cleaned-product-catalog-valid.png` | Valid Products | Confirmed; included |
| product-catalog-auditor | output | `Products_Invalid` | `cleaned-product-catalog-invalid.png` | Invalid Products | Confirmed; included |
| prestashop-validator | raw | `PrestaShop_Raw` | `pretashop-raw.png` | Raw Products | Confirmed; included |
| prestashop-validator | output | `PrestaShop_Valid` | `cleaned-pretashop-valid.png` | Valid Products | Confirmed; included |
| prestashop-validator | output | `PrestaShop_Invalid` | `cleaned-pretashop-invalid.png` | Invalid Products | Confirmed; included |
| sales-dashboard | raw | `Orders_Raw` | `sales-orders-raw.png` | Raw Orders | Confirmed; included |
| sales-dashboard | output | `Dashboard` | `cleaned-sales-dashboard.png` | Dashboard | Confirmed; included |
| sales-dashboard | output | `Orders_Clean` | `cleaned-sales-order-valid.png` | Clean Orders | Confirmed; included |
| sales-dashboard | output | `Orders_Invalid` | `cleaned-sales-order-invalid.png` | Invalid Orders | Confirmed; included |
| sales-dashboard | output | `Monthly_Sales` | `cleaned-monthly-sales.png` | Monthly Sales | Confirmed; included |
| sales-dashboard | output | `Category_Sales` | `cleaned-category-sales.png` | Category Sales | Confirmed; included |
| sales-dashboard | output | `Channel_Sales` | `cleaned-channel-sales.png` | Channel Sales | Confirmed; included |
| sales-dashboard | output | `Top_Products` | `cleaned-top-product-sales.png` | Top Products | Confirmed; included |
| seo-auditor | raw | `URLs_Raw` | `urls-raw.png` | Source URLs | Confirmed; included |
| seo-auditor | output | `SEO_Audit` | `cleaned-seo-audit.png` | SEO Audit | Confirmed; included |
| pdf-extractor | output | `Extracted_Data` | `clean-extracted-data.png` | Extracted Data | Confirmed; included |

Defaults are the first included row in each group. PDF raw input defaults to Dental Invoice. Sales output defaults to Dashboard. Cleaning raw defaults to Orders_Raw because Customers_Raw has no supplied PNG.

## Intentionally excluded worksheets

No screenshots or placeholder tabs are needed for these worksheets.

| Workbook | Worksheet |
|---|---|
| `website/business-lead-scraper/output/business_leads_scrape_07.01_1.xlsx` | `Summary` |
| `website/clean-and-validate-excel/output/cleaned_dataset_06.30_1.xlsx` | `Summary` |
| `website/clean-and-validate-excel/raw/large_messy_dataset.xlsx` | `README` |
| `website/pdf-extractor/output/pdf_extraction_report_07.01_1.xlsx` | `Summary` |
| `website/prestashop-validator/output/prestashop_validation_06.30_1.xlsx` | `Summary` |
| `website/product-catalog-auditor/output/product_catalog_audit_06.30_1.xlsx` | `Summary` |
| `website/seo-auditor/output/seo_audit_06.30_1.xlsx` | `Summary` |

## Missing non-excluded worksheets

| Project / group | Worksheet | Decision required |
|---|---|---|
| clean-and-validate-excel / raw | Customers_Raw | Supply a verified PNG later or explicitly approve omission. Current generic raw PNG shows Orders_Raw. |
| pdf-extractor / output | Invalid_PDFs | Header-only sheet; approve omission/empty-state policy or supply a PNG later. |
| seo-auditor / output | URLs_Invalid | No nonempty cell values; approve omission/empty-state policy or supply a PNG later. |

None becomes a tab or a SheetJS fallback. There are no uncertain worksheet identities among the 25 PNGs.

## Screenshot dimensions

Width/height ratio is recorded to six decimal places. Metadata covers all 25 screenshots, including the blocked one; the runtime configuration contains 24 approved screenshot entries.

| Screenshot relative to website/ | Width × height | Bytes | Aspect ratio | Shape |
|---|---|---|---|---|
| `business-lead-scraper/raw/business-lead-scraper-raw.png` | 1757 × 481 | 186905 | 3.652807 | Wide |
| `business-lead-scraper/output/business-lead-output.png` | 6696 × 559 | 319509 | 11.978533 | Wide |
| `business-lead-scraper/output/business-lead-invalid-url.png` | 1895 × 524 | 24304 | 3.616412 | Wide |
| `clean-and-validate-excel/raw/clean-and-validate-excel-raw.png` | 1120 × 3421 | 610562 | 0.327390 | Tall |
| `clean-and-validate-excel/output/clean-customers-valid.png` | 1262 × 3254 | 928144 | 0.387830 | Tall |
| `clean-and-validate-excel/output/clean-customers-invalid.png` | 1517 × 2647 | 904601 | 0.573102 | Tall |
| `clean-and-validate-excel/output/clean-orders-valid.png` | 1257 × 3254 | 526742 | 0.386294 | Tall |
| `clean-and-validate-excel/output/clean-orders-invalid.png` | 1404 × 1469 | 357525 | 0.955752 | Standard |
| `product-catalog-auditor/raw/product-catalog-raw.png` | 1773 × 4249 | 1460241 | 0.417275 | Both |
| `product-catalog-auditor/output/cleaned-product-catalog-valid.png` | 1804 × 5292 | 3520788 | 0.340892 | Both |
| `product-catalog-auditor/output/cleaned-product-catalog-invalid.png` | 1892 × 533 | 245348 | 3.549719 | Wide |
| `prestashop-validator/raw/pretashop-raw.png` | 2529 × 5269 | 7040862 | 0.479977 | Both |
| `prestashop-validator/output/cleaned-pretashop-valid.png` | 2154 × 2744 | 3797517 | 0.784985 | Both |
| `prestashop-validator/output/cleaned-pretashop-invalid.png` | 2430 × 483 | 146989 | 5.031056 | Wide |
| `sales-dashboard/raw/sales-orders-raw.png` | 1609 × 3537 | 934138 | 0.454905 | Both |
| `sales-dashboard/output/cleaned-sales-dashboard.png` | 1323 × 811 | 77233 | 1.631319 | Standard |
| `sales-dashboard/output/cleaned-sales-order-valid.png` | 1581 × 11474 | 10975458 | 0.137790 | Tall |
| `sales-dashboard/output/cleaned-sales-order-invalid.png` | 2113 × 180 | 114204 | 11.738889 | Wide |
| `sales-dashboard/output/cleaned-monthly-sales.png` | 797 × 223 | 37146 | 3.573991 | Standard |
| `sales-dashboard/output/cleaned-category-sales.png` | 501 × 172 | 19974 | 2.912791 | Standard |
| `sales-dashboard/output/cleaned-channel-sales.png` | 501 × 173 | 19451 | 2.895954 | Standard |
| `sales-dashboard/output/cleaned-top-product-sales.png` | 666 × 261 | 43634 | 2.551724 | Standard |
| `seo-auditor/raw/urls-raw.png` | 1297 × 304 | 61154 | 4.266447 | Standard |
| `seo-auditor/output/cleaned-seo-audit.png` | 4225 × 388 | 313363 | 10.889175 | Wide |
| `pdf-extractor/output/clean-extracted-data.png` | 2432 × 252 | 87196 | 9.650794 | Wide |

The largest width is 6696 pixels (blocked Business Leads screenshot). The greatest height, pixel count and file size belong to cleaned-sales-order-valid.png: 1581 × 11474, 18,140,394 pixels, 10,975,458 bytes. All 25 screenshots total 32,752,988 bytes; 24 configured screenshots total 32,433,479 bytes.

## Screenshot coverage

Captures are previews, not complete accessible spreadsheets. Preserve the configuration coverage notes in the later UI.

- Cleaning raw: Orders_Raw excerpt ending around row 130/1201; Customers_Raw is not pictured.
- Valid customers: around row 104/433; valid orders: around row 104/1156.
- Invalid cleaning orders: column A (Order ID) is absent.
- PrestaShop raw: around row 150/217; valid output: around row 100/201.
- Product catalog raw: around row 121/196.
- Sales raw: around row 200/757.
- Clean sales orders: around row 750/751, with visible column lettering skipping J/K; review coverage.
- Blocked Business Leads screenshot: frozen header followed by row 4; rows 2–3 are absent.
- Selection/collaborator markers, blank grid areas and clipped cell text remain part of the supplied images. No PNG was modified.

## PDF inputs

Actual PDF paths are configured under `/pdf-extractor/raw/`:

1. `invoice_001_dental.pdf` — Dental Invoice
2. `invoice_002_construction.pdf` — Construction Invoice
3. `invoice_missing_fields.pdf` — Invoice with Missing Fields
4. `purchase_order_001.pdf` — Purchase Order

Output uses the shared screenshot mapping for Extracted_Data plus the approved extraction workbook download. No PDF was converted or regenerated.

## Existing homepage issues deferred

- `website/index.html:7,137`: images/logo.png does not exist relative to the homepage; actual logo is assets/images/logo.png.
- `website/index.html:9`: style.css does not exist relative to the homepage; actual stylesheet is assets/css/style.css.
- `website/index.html:164–267`: 14 configured preview base paths resolve beneath nonexistent website/datasets or website/outputs directories. Four actual PDF paths share that issue.
- Script labels and displayed workbook metadata contain discrepancies documented in Phase 1.
- No homepage path, SheetJS renderer, CSS or deployment configuration was changed.

## Readiness

Configuration is ready for a Phase 2B viewer to consume the approved subset. Full asset readiness still needs decisions on the three missing non-excluded worksheets. Public deployment remains blocked until the two Business Lead files are safely replaced or removed with approval.

Two website workbooks differed from their source copies before this phase: cleaning raw large_messy_dataset.xlsx and PDF output pdf_extraction_report_07.01_1.xlsx. Phase 1 confirmed equal worksheet order and cell values. Preserve those existing versions; do not overwrite them.

Validation must check JavaScript syntax, all seven projects, exact path casing and website-root containment, image metadata, worksheet membership/order/defaults, exclusion and blocked-path absence, PDF paths, and preservation of all existing file hashes. No dependency or build step is required.

