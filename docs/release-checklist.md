# Phase 2E / 2F release report and client presentation checklist

> Historical pre-migration release audit. Phase 3A moved the active site to repository root and changed deployment exclusions. The former website/ root and deployment commands below are superseded by [README](../README.md) and [Phase 3A](phase-3a-migration.md). No deployment is authorized by this record.

Checked 2026-10-24 against the existing working tree. No application, dataset, script, workbook, image, viewer, or hosting configuration changes were needed. This report is the only repository file created during the release audit. Existing uncommitted work was preserved.

## Release status

| Gate | Result | Evidence |
|---|---|---|
| Security, local deployment scope | PASS | Inspected all files beneath website/, including decoded workbook packages. No original Business Lead token values, duplicate unsafe originals, credential signatures, or sensitive URL parameter assignments found. Private originals remain outside website/, are ignored by Git, and match preserved hashes. |
| Functional QA | PASS | All eight routes styled; 25 configured screenshot previews, 13 approved workbook download URLs, and four PDF selections passed. |
| Responsive QA | PASS | All routes checked at 375, 768, and 1280 pixels; additional 320 and 430 pixel checks passed. No page-level horizontal overflow. Desktop/mobile footer and lead viewer visually reviewed. |
| Performance behavior | PASS | Homepage has no SheetJS, workbook parsing, workbook/PDF requests, screenshot initialization, or project-image preloading. Screenshot viewers request active images on demand; PDF viewer requests only the selected document. |
| Preview deployment | BLOCKED | Installed Vercel CLI authentication check returned: "The specified token is not valid. Use vercel login to generate a new token." No local .vercel project link or vercel.json exists. Remote project settings could not be inspected. |
| Actual preview URL / live QA | NOT AVAILABLE | No deployment performed, so no preview URL or live verification is claimed. |

Security PASS applies to the inspected website/ document root. Remote Vercel deployment scope must still be confirmed after login. This is not authorization to publish the repository root.

## Executed checks

- Temporary local static server used website/ as its document root; no production server configuration changed.
- Routes: /, /clean-and-validate-excel/, /product-catalog-auditor/, /prestashop-validator/, /sales-dashboard/, /seo-auditor/, /business-lead-scraper/, /pdf-extractor/.
- Excel screenshot checks: configured paths, dimensions, labels, worksheet order, defaults, independent raw/output selection, zoom in/out, reset, fit width, expanded viewing and focus restoration, horizontal/vertical scroll bounds, and top-left reset.
- Business Lead output: Valid Leads first/default, Invalid URLs second, safe demo workbook and screenshot load, actual keyboard tab activation works.
- PDF checks: four configured documents, one selected iframe, selected-document-only requests, new-tab destination, mobile selector, and simulated failure/retry. All four source PDFs identify themselves as testing samples and have no detected catalog actions or attachments.
- Homepage: all seven cards navigate correctly, Contact and Get in Touch reach the final shared footer, redundant contact section absent, no extra project asset loads.
- Footer: shared implementation and correct logo across all eight routes; four exact external destinations and safe new-tab attributes verified on project pages.
- No browser console errors, page exceptions, or unexpected HTTP errors. JavaScript syntax checks and git diff --check passed.
- Original Business Lead workbook URL, business-lead-valid.png, historical business-lead-output.png, and private preservation directory HTTP probes returned actual 404 responses without a test-server denylist.
- Public filesystem inspection found no symlinks into private folders, environment files, keys, source Python scripts, or copied unsafe originals.
- GitHub, Portfolio, Resume: HTTP 200 over the network. LinkedIn: HTTP 999, so automated destination availability is unverified; the configured URL is correct. HTTP 200 alone does not verify every external service interaction or resume download permission.

No staging, commits, pushes, production changes, or deployments were performed. Physical devices and non-Chrome browser engines were not tested. Final deployed-host behavior remains untested.

## Exact preview deployment steps after authentication is restored

1. Open PowerShell in the repository root, not inside website/:

   ```powershell
   Set-Location -LiteralPath 'C:\Users\Axell\Documents\My Projects\Python Automation'
   vercel.cmd login
   vercel.cmd whoami
   vercel.cmd project ls
   ```

2. Select the correct account/team and identify the existing Vercel project associated with `axellgaceta21/python-automation-collection`. Its actual Vercel name is unknown because authentication failed. Inspect that existing project's settings; do not create a duplicate or assume its name matches the repository. If there is no existing project, obtain project/account setup approval before proceeding.

3. In that project's **Settings > Build and Deployment**, confirm:

   | Setting | Value |
   |---|---|
   | Root Directory | `website` |
   | Framework Preset | `Other` |
   | Build Command | Override enabled, command empty (skip build) |
   | Output Directory | `.` relative to website/; Other's default also uses this because no public/ subdirectory exists |
   | Dependencies / environment | No application packages or environment variables required |

   Check the actual existing settings before changing them. Do not publish the repository root or include private-assets/, datasets/, scripts/, or the original outputs/. Nested index.html files provide the project routes; no SPA rewrite or framework build is required.

4. Link the repository root to that existing project, then inspect it:

   ```powershell
   vercel.cmd link
   vercel.cmd project inspect
   ```

   Choose the existing project in the prompts. Do not accept a default that creates a new project. Confirm Root Directory is website before deployment. Do not combine a website/ working directory with a project setting that applies website/ again. Ignore any generated local .vercel link directory in Git before a future approved commit; no link directory was generated by this audit.

5. Deploy a **preview** from the linked repository root:

   ```powershell
   vercel.cmd deploy --target=preview
   ```

   Capture the actual URL printed by the CLI. This uses local working-tree files and does not require a Git push. Do not use --prod, promote, or alter production aliases. If a required account permission or commit blocks the workflow, stop and request authorization.

6. On the actual preview URL, repeat route, screenshot, PDF, workbook-download, and original-asset 404 checks. Verify that client access works under the existing Preview Deployment Protection settings; do not disable protection without authorization. Only then share the verified preview URL for the presentation.

Official references: [Vercel build/root settings](https://vercel.com/docs/builds/configure-a-build), [linking an existing project](https://vercel.com/docs/cli/link), [preview deployment target](https://vercel.com/docs/cli/deploy).

## Short client demonstration sequence

- [ ] Homepage: explain the raw information → Python workflow → usable results progression, then show the seven project cards.
- [ ] Sales Dashboard first: open Dashboard, switch to a summary worksheet, demonstrate Fit Width, zoom and Expand. Describe these as read-only Excel screenshots, not live workbook calculations.
- [ ] Business Lead Scraper: show Valid Leads first, explicitly identify the synthetic example records, then switch to Invalid URLs to explain exception reporting.
- [ ] PDF Extractor: switch between sample invoices and the purchase order, show Open PDF in New Tab, then the Extracted Data output.
- [ ] Show one workbook download to demonstrate that the complete report is available beyond the screenshot preview.
- [ ] Briefly point out Clean and Validate Excel, Product Catalog Auditor, PrestaShop Validator, and SEO Auditor using the homepage cards; open the one closest to the client's workflow.
- [ ] End at the shared footer with Portfolio, GitHub, Resume, and LinkedIn. Manually open LinkedIn before the call because automated availability could not be verified.

## Noncritical follow-up

- Some unchanged approved workbooks retain local file-path/authoring metadata, as documented by the earlier audit; no credential values were detected. Do not alter originals during this fast-track release.
- Some screenshot previews intentionally show excerpts and some worksheets intentionally have no screenshot. Existing coverage notes explain this; no missing screenshots were generated.
- Large screenshots remain unchanged. No loading failures occurred locally, but remote/mobile-network performance cannot be claimed until a real preview is tested. Have the presentation's main views opened ahead of the call.
- deploymentReady remains false in configuration while Vercel authentication, remote scope confirmation, and deployed-host verification are outstanding.
