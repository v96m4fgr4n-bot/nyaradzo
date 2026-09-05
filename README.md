# NFS Policy Mailer

Google Apps Script web app for Nyaradzo Financial Services (NFS). Twice a month, it takes an Excel export of lapsed funeral policies, splits it per sales agent, and emails each agent a branded PDF **and an editable Excel sheet** of just their own lapsed policies — the Excel copy lets agents track and mark off policies as they get reinstated.

Built for Tanyaradzwa Manyeruke. Deployed under `tanashe14@gmail.com`.

## Files

| File | Purpose |
|---|---|
| `Code.gs` | Backend — agent directory, Excel parsing, PDF generation, send/retry logic, auto-send trigger |
| `Index.html` | Frontend — single-page portal (Send, Preview, Agents, Status, History tabs) |
| `appsscript.json` | Manifest — advanced Drive service, OAuth scopes, web app deployment settings |

## How it works

1. Tanya downloads the lapsed-policy export from Easipol and uploads it (Excel) into the **"Tanya Automation"** Drive folder. This step is manual — Easipol has no scheduled/emailed export option, so it can't be automated further without a much riskier browser-automation approach.
2. From the portal's **Send** tab, she picks the file and the agents to send to (or lets the automatic trigger handle it — see below).
3. The backend copies the Excel into a temporary Google Sheet (`Drive.Files.copy`), splits rows by agent name (Column I), builds one PDF per agent (`Utilities.newBlob(...).getAs('application/pdf')`), and emails it via `GmailApp`.
4. Every send attempt is logged to a **"Nyaradzo Mailer Settings"** spreadsheet (`SendLog` tab), which drives the Status and History tabs.
5. `autoSend()` runs on a daily trigger but only actually sends on the 1st and 15th of the month, picking the first Excel file it finds in the folder.

No `UrlFetchApp` and no `Drive.Files.export()` are used anywhere — both have caused permission/format errors in this deployment before, per prior history.

## Deployment

Panashe edits code here; Tanya (`tanashe14@gmail.com`, project owner) must redeploy after any change: **Deploy → Manage Deployments → New Version**.

### Syncing with the live Apps Script project via clasp

This repo is wired up for [`clasp`](https://github.com/google/clasp) (Google's Apps Script CLI), pointed at the live **TANYA AUTOMATION** script project (`.clasp.json`). `.claspignore` restricts what gets pushed to exactly `Code.gs`, `Index.html`, and `appsscript.json` — nothing else in the repo (README, package.json, etc.) touches the live project.

```
npm install        # installs clasp locally (already run once in this repo)
npx clasp login     # one-time: opens a Google OAuth flow — must be run by
                     # someone with edit access to the script (its owner,
                     # tanashe14@gmail.com, or another editor added to it)
npx clasp push      # uploads Code.gs / Index.html / appsscript.json to the
                     # live Apps Script project's editor
npx clasp pull      # pulls the live project's files back down into this repo
npx clasp deploy    # creates a new deployment version — this is what
                     # actually makes changes live on the web app URL
```

Important: `clasp push` updates the code sitting in the Apps Script *editor*, but the deployed web app keeps serving whichever version was last deployed until someone runs `clasp deploy` (or uses Deploy → Manage Deployments → New Version in the web UI). `clasp login` requires an interactive Google OAuth consent — it cannot be completed non-interactively, so a human has to do it once per machine/account.

## Known limitations

- Agent name matching (Excel Column I ↔ agent roster) is an exact, case/whitespace-insensitive string match — no fuzzy matching. A misspelled or new agent name in the source file is caught and flagged (see Fixes below) but not auto-corrected.
- The whole send runs as a single synchronous Apps Script execution; a very large file or agent list could in theory hit the execution time limit mid-run, leaving some agents with no log entry at all for that attempt.
- File upload to the "Tanya Automation" Drive folder is manual — Easipol doesn't support scheduled/emailed exports.

## Fixes applied (this session)

Starting from the live deployment pulled from the "TANYA AUTOMATION" Apps Script project, the following defects were found and fixed:

1. **History double-counted resends** — `getSendHistory()` summed every `"Sent"` log row per cycle instead of the latest one per agent, so using the Override (resend) feature inflated the cycle's agent/policy totals. Fixed by deduping to the most recent log row per agent per cycle before aggregating.
2. **Unmatched agent names silently dropped policies** — the send loop only ever iterated the known agent roster, so a row whose Column I name didn't exactly match any agent (typo, new agent not yet added) was never sent to anyone and never logged. Fixed: `sendEmails` now detects rows with no matching roster agent, logs them under a dedicated `"UNMATCHED"` row, includes them in the run's live log and the summary email to Tanya; `getAgentPolicyCounts` (Preview) now flags unmatched names too, so they're visible *before* sending.
3. **Cycle label/key could desync** — the cycle label shown in logs came from the browser's local clock while the cycle key was always computed server-side; a clock/timezone drift or a send right at a cycle boundary could log a label under the wrong key. Fixed: both are now always derived from the same server-side timestamp inside `sendEmails`.
4. **No real "already sent" protection** — the Send tab's "Override — allow resend to already-sent agents" toggle didn't actually do anything: already-sent agents were never disabled in the checklist in the first place, override or not. Fixed both sides: the Send tab now disables agents already sent this cycle unless Override is on, and `sendEmails` itself now refuses to resend an already-sent agent unless the `override` flag is explicitly passed.
5. **File picker could miss the uploaded file** — `getFilesInFolder()` only matched filenames ending in `.xlsx`/`.xls`/`.csv`; if Google Drive's "convert uploads" setting silently turns the upload into a native Google Sheet, it would disappear from the file list with no explanation. Fixed by also matching on the Google Sheets mimeType.

## Design pass (this session)

The portal's UI was already a mature, well-considered design (tinted shadows, dark mode, hover states, empty/loading states, a consistently applied brand palette) — not the generic "AI slop" pattern most redesign checklists target. Rather than force sweeping stylistic changes, an audit found and fixed a handful of genuine, low-risk gaps in `Index.html`:

- Added `font-variant-numeric: tabular-nums` so the stats/counts don't jitter as values update.
- Added a consistent branded `:focus-visible` ring — buttons, tabs, and checkboxes previously had no custom keyboard-focus indicator.
- Added `:active` press feedback on buttons and tabs (previously only hover states existed).
- Tinted the two remaining plain-black box-shadows (banner logo, modal) to match the navy-tinted shadow language used everywhere else.
- Added semantic landmarks and ARIA: the header banner is now a `<header>`, the tab bar is a `<nav role="tablist">` with `role="tab"`/`aria-selected` per button, and each panel is `role="tabpanel"` with `aria-labelledby` — `switchTab()` now keeps `aria-selected` in sync.

Deliberately left unchanged: the Inter font. Font-swapping is usually the highest-impact fix for marketing/landing pages, but for a dense internal data tool, Inter's legibility and tabular-figure support outweigh "personality" — swapping it here would be change for change's sake.

## Visual revamp (this session, round 2)

Asked for a deeper revamp beyond the polish pass above. Ran the `ui-ux-pro-max` skill's design-system search against "internal financial services admin dashboard" (density 8/10, variance 4/10, motion 4/10) and applied its top matches:

- **Typography:** swapped to a **Fira Code + Fira Sans** pairing (the top database match for "dashboards, analytics, admin panels") — Fira Sans for labels/prose, Fira Code for anything numeric or structured (stat values, badges, pill counts, timestamps, the log panel, email/CSV fields). This also fixed a pre-existing bug where the email and CSV-import fields declared `font-family:'Inter',monospace` — Inter isn't a monospace font, so it never actually rendered as one.
- **Color:** refined the light/dark neutral scale (background, border, muted-text tokens) against a validated navy+gold banking/finance color reference, keeping the exact brand hex (`#1a3080` / `#c8a000`) as the anchor — same CSS variable names throughout, so nothing downstream needed to change.
- **Glassmorphism**, per the style database's "financial dashboards, high-end corporate" recommendation — applied only to the chrome (the header banner + stat cards now form one continuous navy gradient hero band, with the four stat tiles rendered as frosted/translucent glass; the modal overlay now blurs the content behind it), never to the dense data tables/lists, where full opacity and contrast matter more than visual flair.

Verified with a Playwright render (light + dark, Send + Agents tabs, tab-switching) that everything still renders and functions correctly.

## Excel attachment (this session)

Each agent's email now includes an editable `.xlsx` alongside the PDF, so they can mark off policies once they're reinstated — `buildXlsxBlob()` in `Code.gs`.

Since this project cannot use `UrlFetchApp`, `Drive.Files.export()`, or any external library (see Known Issues below), the `.xlsx` is hand-built: an OOXML spreadsheet package (the same zip-of-XML-files format every real `.xlsx` uses internally) assembled directly via `Utilities.zip()`, then relabeled with the correct filename and content type so it opens as a normal Excel file rather than a zip archive. All cells are written as inline strings (no `sharedStrings.xml`) to keep the format minimal.

This was validated outside of Apps Script before shipping: the exact same function, run against synthetic data through a Node.js harness with Drive/Sheets/Gmail mocked out, produces a file that
- round-trips correctly through `openpyxl` (a real, independent OOXML implementation) with zero warnings, and
- is identified by the `file` utility's content-based magic detection as `Microsoft Excel 2007+`, not a zip archive — the same detection mechanism email clients and OS file browsers use.

## Agent column fix (this session)

Live testing against the real Easipol export found the Preview tab not grouping by agent correctly. The real export's actual columns (confirmed via the column diagnostic below) are:

| Col | A | B | C | D | E | F | G | H | I |
|---|---|---|---|---|---|---|---|---|---|
| Header | Policy_Number | Inception_Date | fullname | Cell_Number | EmailAddress | UsualPremium | Currency | **AgentsName** | currstatus |

So the agent name is in **Column H (index 7)**, and Column I (`currstatus`) is the policy status — not what was originally documented (Column I = Agent). Went through two wrong guesses first (Column I itself, then Column J) before adding a proper diagnostic to see the real layout instead of guessing further:

- `getAgentPolicyCounts()` now returns each column's letter, header text, and a sample value from the first data row; the Preview tab renders these as a small strip with the currently-assumed agent column highlighted, so a layout mismatch like this is visible immediately instead of requiring a support round-trip.
- Fixed a real bug this surfaced along the way: the Preview success handler had no error handling, so any exception thrown while rendering left the UI stuck on "Analysing file..." forever with no feedback (`withFailureHandler` only covers server-side errors, not client-side ones in the success callback). It now catches and displays render errors plus the raw server response.
- Fixed the actual cause of one such silent failure: passing a raw `Date` cell value through `google.script.run` can break serialization and deliver `null` to the client instead of a real error. Diagnostic cell values are now converted to plain strings server-side before being returned.

`AGENT_COL` is now `7` in both `sendEmails()` and `getAgentPolicyCounts()`. If Easipol's export layout shifts again in the future, the more durable fix would be to look up the Agent column by matching the header row's text (e.g. a column literally titled "AgentsName") instead of a hardcoded index — the diagnostic strip makes that easy to spot when it happens, but doesn't self-correct it. Flagged as a follow-up, not yet implemented.

## Lapsed/Inactive status split (this session)

Column I (`currstatus`) turned out to hold a mix of `Lapsed` and `Inactive` values, not just lapsed policies as originally assumed — the mailer was combining both into one number everywhere with no way to tell them apart. Added `STATUS_COL` (Column I) and shared helpers `countByStatus_()` / `formatStatusBreakdown_()` in `Code.gs`, used to show the split instead of one combined count in:

- **Preview** — the top summary line and each per-agent row now show e.g. "40 Lapsed, 23 Inactive" alongside the total.
- **Email body** — "Please find attached your policy list for this period (63 policies: 40 Lapsed, 23 Inactive)".
- **PDF** — the "Total Policies" meta box now includes the same breakdown.

The underlying send/skip/duplicate-guard logic is unchanged — this only changes what's *displayed*, not which policies get sent (still: every row matched to that agent, regardless of status, per the "keep sending both" decision). A category other than "Lapsed"/"Inactive" (matched via case-insensitive substring) falls into "Other" so nothing is silently dropped from the counts.

## Per-agent attachment preview (this session)

Added a 👁 button on each Preview-tab row that generates the exact PDF and Excel a given agent would receive — without sending anything — so you can check the output before committing to a real send.

- New backend function `previewAgentAttachments(fileId, agentName)`: re-reads the file, filters to that one agent (same logic as `sendEmails()`), builds the real PDF and Excel via the existing `buildPdfBlob()`/`buildXlsxBlob()`, and returns them as base64 strings.
- The client decodes the base64 into real `Blob` objects in the browser (`base64ToBlob()`) — the PDF opens in a new tab via `URL.createObjectURL()`, the Excel downloads directly. No temporary files are created in Drive, and no email is sent.
- Verified via the dry-run harness that `previewAgentAttachments()` produces byte-correct, independently-openable output (same `openpyxl` validation as the Excel attachment feature) and confirmed zero emails are sent when it runs.

## Grouped Lapsed/Inactive sections in PDF & Excel (this session)

Previously an agent's policy list mixed Lapsed and Inactive rows in whatever order they appeared in the source file. Both attachments now group rows into labeled blocks — Lapsed first, then Inactive, then Other (anything not matching either) — via a new shared helper `groupRowsByStatus_(rows)` in `Code.gs`, used by both `buildPdfBlob()` and `buildXlsxBlob()` (and therefore by `sendEmails()` and the 👁 preview button alike).

- **PDF** — each group gets a full-width colored banner row ("LAPSED (12)", "INACTIVE (5)") above its rows: red for Lapsed, amber for Inactive, gray for Other.
- **Excel** — each group gets a bold, white-on-navy merged banner row spanning all columns, built by hand-adding a second font/fill/cellXf to the workbook's `styles.xml` and a `<mergeCells>` entry per banner row in `sheet1.xml`.
- A status with zero matching rows for that agent produces no banner at all — no empty sections.
- Verified with the dry-run harness + `openpyxl`: banner rows render bold and merged (`A2:I2`, etc.), row order is Lapsed block → Inactive block → Other block, and the PDF HTML contains all three group classes with correct per-group counts.

## Fix unmatched agent names directly from Preview (this session)

The Preview tab already flagged names in the file that don't match anyone in the agent roster (highlighted red, e.g. a typo like "TYPO AGENCT"), but fixing it meant leaving the portal to add the agent manually. Added a **＋ Add** button on every unmatched row that opens an inline "email address" field right there, without leaving the tab.

- New backend function `addAgent(name, email)` in `Code.gs`: validates the email format, rejects a duplicate name (case-insensitive), and appends a single row to the `AgentEmails` settings sheet — it doesn't touch the rest of the roster, unlike `saveAgents()` which replaces the whole sheet.
- Clicking **＋ Add** reveals an inline form (email input + Save/Cancel) under that row; Enter or **Save** calls `addAgent()`, shows a toast, then reloads both the Agents tab and the Preview tab so the name immediately drops out of the "unmatched" list and its policies count toward a real agent.
- Caught and fixed a real bug while testing this with Playwright: the inline form's `hidden` attribute was being overridden by its own inline `display:flex` style (inline styles always win), so it rendered open by default. Fixed by toggling `style.display` directly instead of mixing it with the `hidden` attribute.
- Verified with the dry-run harness (bad email rejected, duplicate name rejected, valid add persists and flips that agent's `known` flag to `true` on the next preview) and with Playwright screenshots of the full click → fill → save → toast flow. Also re-verified the pre-existing "+ Add Agent" modal on the Agents tab still works unchanged (the new client-side function was named `confirmAddAgentFromPreview` specifically to avoid colliding with the existing `confirmAddAgent()`).
