# Phase 3A migration report

Date: 2026-10-24 (Asia/Manila). Status: root migration complete; no commit, push, deployment, or remote settings changes.

## A. Before migration

The Git root was `C:/Users/Axell/Documents/My Projects/Python Automation`. `website/` contained 63 files: eight HTML pages, 26 PNGs (25 worksheet previews plus logo), 13 workbooks, four PDFs, five stylesheets, six scripts, and an empty `.gitkeep`. All active site files were untracked. The old tracked root `index.html`, `style.css`, and `images/logo.png` had already been removed locally. Nothing was staged. `archive/` existed but was empty; no archive paths were tracked. There were no filename conflicts, `.vercelignore`, local `.vercel` link, or Vercel configuration file.

The original Python scripts, datasets, outputs, docs, and private backups were identified separately. Script documentation and relative data paths establish that scripts/datasets/outputs are automation source materials, not confirmed obsolete website files. They remain in place.

## B. Archive

The three absent legacy files were recovered byte-for-byte from Git HEAD into previously unused paths, without changing the Git index or restoring them over the working site:

| Original path | Local archived path |
|---|---|
| index.html | archive/old-website/index.html |
| style.css | archive/old-website/style.css |
| images/logo.png | archive/old-website/images/logo.png |

All three were SHA-256 verified. No existing archive file was overwritten. Archive hashes and source/destination mappings are in `phase-3a-manifest.json`. No tracked archive paths required untracking; no Git index operation was necessary.

## C–E. Root layout and changes

All 63 site files moved from `website/` to the same relative paths at root. The old directory was removed only after destination/hash verification and an empty-directory check. Windows blocked a directory rename, so the remaining files were copied exclusively, verified, and their source copies removed individually. There is no duplicate active website.

All root-relative runtime URLs already matched the new layout. No asset URLs, worksheet mappings, downloads, footer destinations, or PDF paths changed. The only migrated application-file edit was the deployment-root comment in `assets/js/projects-config.js`; all other 62 migrated files are byte-identical. Its runtime data is unchanged.

Created: `.vercelignore`, `vercel.json`, `README.md`, `scripts/serve-portfolio.mjs`, this report, and `phase-3a-manifest.json`. Updated `.gitignore` with `/archive/` and `.vercel/`, preserving existing exclusions. Historical audit headers now point to this current migration record. Original Python scripts, original datasets/reports, and private backup contents are unchanged.

## F. Security and deployment scope

Archive and private backups are ignored by Git and excluded from CLI uploads. No archive or private path is tracked or staged. Explicit `@vercel/static` source rules additionally limit the final static output for Git integrations, where tracked source materials can be present in the checkout. The CLI ignore-pattern selection was checked using Git's native ignore parser for both tracked and untracked files. It selected exactly 63 upload files: 62 public static files and `vercel.json`. The empty `.gitkeep` is preserved locally but not published.

A local inspection artifact was assembled from the static source rules and checked byte-for-byte. Its 62 files comprise the eight routes and all approved assets; it contains no archive, private backups, docs, scripts, root datasets/outputs, environment files, or Git internals. Scanning those files and decoded workbook packages found no original token values, unsafe original duplicates, credential signatures, or token-bearing URL assignments. This is local selection/inspection evidence, not a claim that a remote Vercel build ran.

The local preview server uses the same static source allowlist. Actual HTTP probes returned 404 for archive originals, private backups, root original outputs, Python source, datasets, docs, Git/configuration files, former `/website/index.html`, and both historical unsafe screenshot names. A generic root server would not enforce these rules; use the provided preview command.

**Separate pre-existing GitHub exposure:** `outputs/business_leads_scrape_07.01_1.xlsx` is already tracked in HEAD; local main and origin/main point to the same recorded commit. A public raw GitHub request returned HTTP 200 and a byte-identical workbook. Its URL query token fields were confirmed without printing values. Site exclusion does not remove the public GitHub copy or history. No report deletion, history rewrite, or credential action was performed. Separate remediation review/approval remains required before a new public release.

**Separate pre-existing production exposure:** a read-only request to the existing production site's `/outputs/business_leads_scrape_07.01_1.xlsx` path also returned HTTP 200 and matched the tracked original's hash. The migrated localhost site returns 404 for that path. No production deployment, deletion, access-setting change, or cleanup was performed. New local exclusions do not protect old deployments retroactively. Release remains blocked pending separate review/remediation of the existing GitHub and production copies and any relevant historical copies. Whether the embedded form identifiers are confidential credentials remains a matter for review; no token values are logged or reproduced.

## G. Executed local QA

Tested the repository-root site at `http://127.0.0.1:5500/` using the supplied server handler:

- All eight routes render styled pages, including direct nested navigation.
- All 25 approved screenshot tabs load with correct paths, dimensions, labels, order/defaults, and independent raw/output viewers.
- Zoom in/out, 100% reset, Fit Width, expanded viewing/focus restoration, scroll bounds/top-left reset, and actual keyboard tab activation passed.
- Business Lead output is Valid Leads first/default, then Invalid URLs. Safe demo workbook and screenshot load; unsafe originals remain inaccessible through the site.
- All 13 approved workbook download URLs return 200.
- All four PDFs select and render using the native viewer; only the selected PDF loads. New-tab destination and simulated failure/retry passed.
- Homepage's seven links, Contact/Get in Touch footer anchors, lightweight requests, shared footer, exact external link destinations/new-tab attributes, and logo loading passed.
- All eight routes passed overflow/layout checks at 320, 375, 430, 768, and 1280 pixels. Desktop/mobile footer and lead viewer were visually reviewed.
- No page exceptions, console errors, or unexpected asset HTTP errors. JavaScript syntax checks and git diff --check passed.
- Active HTML/CSS/JavaScript contain no stale `/website/` URLs. Historical audit documents intentionally retain historical paths and carry migration notices.

Unexecuted: physical-device/non-Chrome checks, real remote Vercel build, and deployed-host verification. Earlier external-link HTTP status checks remain historical; they were not repeated for this file migration.

## H. Git verification

The Git index SHA-256 is unchanged from before migration, with no staged paths. New root `index.html` replaces the old tracked homepage in the working tree; assets and seven project directories are ready for a later approved Git add. Old root style/logo deletions were pre-existing and their Git versions are now archived. All source dataset/output/Python and private-file hashes were preserved. No commit, push, index removal, reset, or history rewrite occurred.

## I–J. Vercel and shortest safe next steps

Authentication became available during inspection. Read-only CLI/API calls verified the existing `python-automation-collection` project and its GitHub link to `axellgaceta21/python-automation-collection`. Root Directory is already repository root (API null/default, CLI `.`), Framework is Other, Production Branch is main, and build/install/output overrides are null/default. No root-directory update is needed. No remote project settings were changed, no local project link was created, and no build or deployment was run.

After Phase 3B approval and the separate existing-public-workbook review:

1. Reuse the verified existing Vercel project; do not create a duplicate.
2. Retain the existing root (`./` / default), Other framework, default build/install/output settings, and main production branch. Keep the explicit static output allowlist; the local configuration requires no custom application build command. Recheck these settings if they change before release.
3. Review the exact Git diff and add only intended website/config/docs/preview-tool changes. Keep archive and private originals unstaged. Resolve the already-public tracked report/history concern separately; ignoring a folder does not erase published data.
4. Commit/push only after authorization. A push can trigger automatic deployment; confirm the existing Git integration, scope, and branch first.
5. Verify the actual deployment's eight routes, asset/download URLs, PDF switching, and non-public 404s before sharing it.

References: [Vercel static-output allowlist](https://vercel.com/docs/project-configuration/vercel-json#builds), [Vercel upload exclusions](https://vercel.com/docs/deployments/vercel-ignore), [root-directory settings](https://vercel.com/docs/builds/configure-a-build#root-directory).
