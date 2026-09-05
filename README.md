# NFS Policy Mailer

Google Apps Script web app for Nyaradzo Financial Services (NFS). Twice a month, it takes an Excel export of lapsed funeral policies, splits it per sales agent, and emails each agent a branded PDF of just their own lapsed policies.

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
