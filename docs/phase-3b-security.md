# Phase 3B pre-push security review

Date: 2026-10-24 (Asia/Manila).

## New commit and deployment scope

The user authorized a normal commit/push through the existing GitHub-to-Vercel integration, subject to checking the files being added. The approved migration files and decoded workbook packages were scanned for original token values, credential signatures, credential assignments, and sensitive URL parameters. No new secret findings were detected. The sanitized business-leads-demo.xlsx and matching PNG are included; the original token-bearing workbook and screenshots are not part of the active site or the explicit static output allowlist.

Archive and preserved private originals remain Git-ignored and unstaged. Original Python scripts, root datasets, and root original outputs are not being changed or added by this migration commit. The Vercel static output rules exclude these non-public directories independently of Git tracking. The original report is already tracked from earlier commits; leaving it unchanged does not remove its GitHub exposure.

The required .gitattributes marks PNG, XLSX, and PDF assets as binary so Windows Git line-ending conversion cannot alter the published source documents. Staged media must match the original local bytes. Existing blank end-of-file lines in text pages are preserved rather than changing the site for cosmetic whitespace warnings.

## Existing exposure and affected service

Earlier read-only checks confirmed the original workbook was retrievable from GitHub and the existing production site's root outputs path. The affected fields are in Business_Leads!N13, Social Links: accountToken and formGroupToken, hosted by form.recallmax.com. Token values and full form URLs are intentionally omitted. No credential-bearing form URL was requested or tested during this review.

The provider is RecallMax, now MaxAssist. These appear in public-form links, but their security semantics, current validity, and revocation requirements have not been established. They must not be assumed harmless or confirmed to be privileged API credentials solely from their names.

## Required owner/provider follow-up

1. The authorized practice/account owner should contact MaxAssist support through its official support channel and identify the affected account/form links privately. Ask whether the two fields grant access or are public form-routing identifiers.
2. If access-bearing or otherwise sensitive, ask the provider to invalidate the exposed identifiers and regenerate the affected form links. Update the legitimate website/forms that use them, and have the provider confirm that the previous links no longer grant unintended access.
3. Ask the provider to review relevant access/submission logs and advise whether additional account credential rotation or other response is necessary. A password reset alone must not be assumed to invalidate these form identifiers.
4. Separately review removal/redaction of the already-public report from the current GitHub tree, older Git history, cached copies, and old Vercel deployments. History cleanup, deployment deletion/protection, or credential revocation requires separate authorization; none is attempted by Phase 3B.

Publishing the new sanitized site can remove the original path from the current production alias, but it does not erase GitHub history or protect older deployment URLs. Production 404 checks are therefore only one part of verification, not evidence that the earlier exposure is fully resolved.

Official provider references: [RecallMax is now MaxAssist](https://maxassist.com/recallmax/), [MaxAssist support](https://maxassist.com/services/unlimited-support/). No public self-service instructions for rotating these specific form fields were verified.
