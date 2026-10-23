# Python Automation Portfolio

The active static website is at the repository root: `index.html`, `assets/`, and seven project directories. Phase 3A moved it out of `website/` without redesigning the site or changing approved worksheet mappings.

## Local preview

Use Node.js 24 or newer:

```powershell
node scripts/serve-portfolio.mjs
```

Open **http://127.0.0.1:5500/**. Stop the server with Ctrl+C. No package installation is required.

Use this preview server instead of a generic repository-root file server. It serves only the static file patterns in `vercel.json`, including the eight routes and approved screenshots, workbooks, and PDFs. It returns 404 for `archive/`, `private-assets/`, `datasets/`, `outputs/`, `scripts/`, `docs/`, configuration files, and Git internals. `.gitignore` and `.vercelignore` do not protect a generic local HTTP server.

## Layout

```text
index.html
assets/                       Shared CSS, JavaScript, and logo
clean-and-validate-excel/
product-catalog-auditor/
prestashop-validator/
sales-dashboard/
seo-auditor/
business-lead-scraper/         Public synthetic lead demonstration only
pdf-extractor/
archive/old-website/           Local-only old homepage, stylesheet, and logo
private-assets/               Local-only preserved unsafe originals
scripts/                      Original Python tools and safe preview server
datasets/                     Original automation inputs; non-public
outputs/                      Original automation reports; non-public
docs/                         Audits and migration records; non-public
.gitignore
.vercelignore
vercel.json
```

Python scripts, input datasets, and original outputs retain their paths and contents because the automation workflows still use them. They were not assumed obsolete. Archive and private originals are ignored by Git and excluded from deployment. Nothing under `.git/` was relocated or rewritten.

## Deployment scope and one-time Vercel settings

Phase 3B is authorized to publish the sanitized migration through the existing GitHub-to-Vercel integration after the pre-push security gate. No manual Vercel deployment or history rewrite is authorized. Phase 3A itself performed no commit, push, deployment, or remote settings change.

Read-only Vercel inspection confirmed the existing **python-automation-collection** project is linked to `axellgaceta21/python-automation-collection` on GitHub. Root Directory is already the repository root (API null/default; CLI `.`), Framework is Other, and Production Branch is **main**. Build, install, and output settings have no custom overrides. No one-time root change is needed. Moving files locally did not change any remote settings; these are the existing verified settings.

- Framework Preset: **Other**.
- No application build/install command or frontend dependency is required.
- Public URL root is `/`; the static files retain their root-relative paths.
- `.vercelignore` independently restricts CLI uploads to the active website and `vercel.json`.
- `vercel.json` explicitly lists `@vercel/static` source patterns. Only those static outputs belong in the final deployment. This also protects Git-triggered deployments, where tracked original reports and scripts can exist in the source checkout despite CLI ignore rules.
- The supported legacy `builds` property is intentional here: it provides an explicit static-output allowlist for this mixed website/Python repository without a custom build script. When present, its static outputs control publishing; do not replace it with an unrestricted “serve every file from `.`” configuration. Any dashboard output-directory field should remain at the root/default, rather than the removed `website/` folder.

Vercel documents that [only builder outputs are deployed when `builds` is specified](https://vercel.com/docs/project-configuration/vercel-json#builds). This is an output restriction, not a frontend framework or Python build. `.vercelignore` is an additional upload exclusion, not a replacement for the output allowlist or review of Git-tracked files. Local file-scope checks passed; an actual Vercel build and deployed-host check remain pending.

## Security follow-up before GitHub release

The website uses only the sanitized `business-leads-demo.xlsx` and its matching screenshot. The original token-bearing workbook remains in the original automation output directory and was already committed before this migration. A read-only public GitHub check on 2026-10-24 returned HTTP 200 for the original report and matched its local SHA-256 hash. The original contains `accountToken` and `formGroupToken` URL parameters; values are not reproduced here.

Excluding it from Vercel does not remove the existing public GitHub file or its history. Review the identifiers' sensitivity and separately authorize appropriate repository cleanup/history handling and any necessary token revocation before a new public release. Phase 3A did not edit the original report, remove tracked source data, or rewrite history.

A read-only Phase 3A check also found the same original workbook at the then-existing production site's `/outputs/business_leads_scrape_07.01_1.xlsx` path (HTTP 200, identical hash). The new static allowlist excludes it, but cannot retroactively secure GitHub history or old deployments. The user has authorized Phase 3B to publish the clean site while retaining separate review/remediation of earlier public copies. See [the security review and provider follow-up](docs/phase-3b-security.md); the affected form service is RecallMax, now MaxAssist. No revocation or history cleanup has been performed.

See [Phase 3A migration report](docs/phase-3a-migration.md) and [file/hash manifest](docs/phase-3a-manifest.json). Earlier Phase 2 documents describe the former `website/` document root and are historical.
