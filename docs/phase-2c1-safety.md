# Phase 2C.1: safe Business Lead demonstration assets

> Historical asset-sanitization record. Phase 3A subsequently moved the active site from website/ to repository root; public URLs are unchanged. See [current scope and existing GitHub exposure](phase-3a-migration.md). Preserved original hashes below remain valid.

Updated 2026-10-24. This record supersedes the Business Lead safety status in the historical Phase 2A audit. Final Phase 2E QA and deployment are not performed.

## Original preservation

Originals were copied into private-assets/business-lead-scraper-originals/ outside website/, then SHA-256 compared with the public copies before removal. The root .gitignore excludes /private-assets/. No original file contents were modified. Existing source workbooks under outputs/ also remain unchanged. The original screenshot on disk was business-lead-valid.png, not the historical business-lead-output.png name.

- business-lead-valid.png: SHA-256 e5e878b9f144b5853895a9712f54b1493b391bc4de369baf1297309e6ddc79fc
- business_leads_scrape_07.01_1.xlsx: SHA-256 f791b0bf14854150792faaa690ab5374b55a065da3cbb20dcbe6d228950d2508

Original sensitive fields: Business_Leads!N13 (Social Links), URL query parameters accountToken and formGroupToken. Values and complete sensitive URLs are intentionally omitted from this report. Other original lead contact fields were replaced wholesale, rather than assumed public-safe.

## Safe replacements

- website/business-lead-scraper/output/business-leads-demo.xlsx
- website/business-lead-scraper/output/business-leads-demo.png

A fresh workbook was authored from an allowlist of structural labels and approved Invalid_URLs data. It retains Summary, Business_Leads, Invalid_URLs, the 18 original lead columns, and 13 lead rows. All lead data is newly synthetic: Demo Business identifiers, reserved example.invalid URLs/email addresses, no phone numbers, and explicitly synthetic address/owner placeholders. No original lead values, relationships, comments, metadata, hidden content, or arbitrary workbook objects were imported. Summary counts are calculated from the demonstration records. The approved Invalid_URLs cells retain their original values; its existing screenshot is unchanged.

The PNG is rendered directly from the safe Business_Leads worksheet, including all 13 records and 18 columns. It preserves the green header and alternating light rows; widths/heights accommodate the synthetic values. Dimensions: 8112 x 1298 pixels; 549842 bytes. It is a new render, not a redaction overlay on the original screenshot.

## Security validation

The exported ZIP/XML package was inspected: cells and shared strings, formulas and cached results, relationships/hyperlinks, workbook visibility, hidden rows/columns, defined names, comments/person records, embedded objects, custom XML, external links/connections, and metadata. No original token values, credential patterns, private machine paths, external relationships, comments, hidden sheets/rows/columns, embeddings, or author-identifying metadata remain. Eight allowlisted Summary formulas were checked, including zero phone contacts. Lead URLs use only reserved example.invalid hosts without query strings, fragments, or user information; emails use demo identifiers at example.invalid. PNG visually reviewed against these synthetic records and metadata checked.

Only after preservation and replacement verification were the unsafe public workbook and actual original screenshot removed. Both historical original screenshot names and the original workbook URL must return 404. Omitting a link alone is not considered protection.

## Configuration and scope

Output tabs: Valid Leads (Business_Leads, first/default), then Invalid URLs (Invalid_URLs). Only the safe demo workbook is downloadable. The shared Excel viewer is unchanged. Other projects and worksheet mappings are unchanged. The previous Business Lead deployment blocker is cleared; deploymentReady remains false pending final QA, and no deployment is authorized by this change.

## Executed local regression checks

Headless Chrome with a temporary static server rooted at website/ passed all eight routes, 25 approved screenshot tabs, 13 workbook download URLs, four PDF document selections, selected-document-only loading, native PDF Open in New Tab, and simulated PDF failure/retry. Viewer checks covered defaults/order, raw/output independence, zoom, fit/reset, scroll bounds/top-left reset, expansion/focus restoration, and actual keyboard tab activation on Business Lead output. All routes had no horizontal overflow at 320, 375, 430, 768, and 1280 pixels. Footer destinations and safe new-tab attributes were checked on all project routes. Homepage Contact/Get in Touch links reached the final footer, its logo loaded, and there were no project workbook/PDF/screenshot requests on the homepage. No page exceptions or unexpected HTTP errors occurred. Original workbook, both historical screenshot names, and private-directory HTTP probes returned 404 without a server denylist. Desktop/mobile footer and lead previews were visually checked. The full public tree (including decoded workbook packages) was scanned for the original token values and duplicate original hashes, with no matches. Original backup fingerprints remained unchanged.

Physical mobile devices, other browser engines, external destination availability, and final deployment were not tested. Phase 2E remains pending approval.
