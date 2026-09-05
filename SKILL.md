# NFS Policy Mailer — Claude Code Skill

## Project Overview
This is a **Google Apps Script** web application built for **Tanyaradzwa Manyeruke** at **Nyaradzo Financial Services (NFS)**. It automates the bi-weekly distribution of lapsed funeral policy lists to sales agents via email with branded PDF attachments.

## Project Structure
The project lives in Google Apps Script and consists of exactly **two files**:

| File | Purpose |
|------|---------|
| `Code.gs` | Backend — all server-side logic (Google Apps Script / JavaScript) |
| `Index.html` | Frontend — full portal UI (HTML, CSS, vanilla JS) |

There is also an `appsscript.json` manifest file for OAuth scopes and Drive API.

---

## Tech Stack
- **Runtime:** Google Apps Script (V8 engine, ES6 JavaScript)
- **Frontend:** Single-file HTML with inline CSS and vanilla JS
- **Email:** GmailApp (Google Apps Script built-in)
- **Storage:** Google Drive (files), Google Sheets (agent emails + send log)
- **PDF generation:** `Utilities.newBlob().getAs('application/pdf')` — HTML rendered to PDF inline, no external calls
- **File processing:** `Drive.Files.copy()` to convert Excel → Google Sheet, then read with SpreadsheetApp
- **NO external APIs, NO UrlFetchApp, NO npm packages**

---

## Key Constraints
1. **No UrlFetchApp** — previously caused permission errors. All operations use built-in Apps Script services only.
2. **No external libraries** — everything is vanilla JS and built-in GAS services.
3. **Single HTML file** — CSS and JS are inline in `Index.html`. No separate files.
4. **PDF via blob** — PDFs are built from HTML strings using `Utilities.newBlob(html, 'text/html').getAs('application/pdf')`. Do not use Drive export or UrlFetchApp for this.
5. **NFS Logo is base64 encoded** and embedded directly as a constant `LOGO_B64` in `Code.gs`. It is also inlined in `Index.html` for the portal header banner.
6. **No React, no Tailwind, no build tools** — pure HTML/CSS/JS only.

---

## Architecture

### Backend (Code.gs)

#### Constants
```javascript
const FOLDER_NAME = "Tanya Automation";       // Drive folder for Excel uploads
const SETTINGS_SHEET = "AgentEmails";         // Sheet tab for agent email storage
const LOG_SHEET = "SendLog";                  // Sheet tab for send history
const SENDER_NAME = "Tanyaradzwa Manyeruke";
const COMPANY_NAME = "Nyaradzo Financial Services";
const SUMMARY_EMAIL = "panasheb85@gmail.com"; // Post-send summary recipient
const LOGO_B64 = "...";                       // Base64 encoded NFS logo PNG
```

#### Key Functions
| Function | Description |
|----------|-------------|
| `doGet()` | Web app entry point — serves Index.html |
| `getAgents()` | Returns agent list from Settings Sheet (falls back to hardcoded default list of 68 agents) |
| `saveAgents(agentList)` | Saves agent list back to Settings Sheet |
| `getFilesInFolder()` | Lists Excel files in the Tanya Automation Drive folder |
| `sendEmails(selectedAgents, fileId, cycleLabel)` | Main send function — filters per agent, builds PDF, sends email |
| `buildPdfBlob(agentName, cycleLabel, headers, rows)` | Builds branded PDF from HTML string |
| `getStatusSummary()` | Returns sent/pending agents for current bi-weekly cycle |
| `getSendHistory()` | Returns all past cycle summaries |
| `getAgentPolicyCounts(fileId)` | Preview — counts policies per agent before sending |
| `autoSend()` | Runs automatically on 1st and 15th of month at 9am |
| `setupAutoSendTrigger()` | One-time setup to create the daily Apps Script trigger |
| `retryFailedAgents(fileId, cycleLabel)` | Retries only failed agents from current cycle |
| `getBiWeeklyCycleLabel(date)` | Returns "01 Apr – 14 Apr 2026" format label |
| `getBiWeeklyCycleKey(date)` | Returns "2026-04-A" / "2026-04-B" key for cycle tracking |

#### Send Flow
1. `Drive.Files.copy()` converts uploaded Excel → temp Google Sheet
2. `SpreadsheetApp.openById()` reads all rows
3. For each agent: filter rows where Column I (index 8) matches agent name
4. Build HTML string → `getAs('application/pdf')` → attach to email
5. `GmailApp.sendEmail()` with `replyTo: "tanyaradzwa.manyeruke@nyaradzo.co.za"` and `from: "Tanashe14@gmail.com"`
6. Log result to SendLog sheet
7. Clean up temp sheet
8. Send summary email to SUMMARY_EMAIL

#### Data Storage (Google Sheets — "Nyaradzo Mailer Settings")
**AgentEmails sheet:**
| AgentName | Email |
|-----------|-------|

**SendLog sheet:**
| Timestamp | AgentName | Email | PoliciesCount | Status | CycleLabel | CycleKey |

---

### Frontend (Index.html)

#### Brand Colors
```css
--navy: #1a3080;       /* Primary — headers, buttons, tabs */
--navy-dark: #122260;  /* Log panel background */
--navy-light: #2a4090; /* Hover states */
--gold: #c8a000;       /* Action color — Send button */
--gold-light: #e0b800;
```

#### Tabs
| Tab | Description |
|-----|-------------|
| **Send** | Select file + select agents + send button + live log |
| **Preview** | Policy count per agent before sending (bar chart) |
| **Agents** | Agent email directory — add, edit, delete, import CSV |
| **Status** | Current cycle sent/pending with search and filter |
| **History** | All past cycles with agent count, policy count, failed count |

#### Key JS Functions
| Function | Description |
|----------|-------------|
| `loadAll()` | Loads agents, files, status, history on init |
| `renderAgentSendList(agents)` | Renders send tab agent list — disables already-sent agents |
| `toggleOverride(checked)` | Unlocks already-sent agents for intentional resend |
| `triggerSend()` | Collects selected agents + file, calls `sendEmails()` via `google.script.run` |
| `handleSendResults(results)` | Processes send results, updates log, shows retry button if failures |
| `loadPreview()` | Calls `getAgentPolicyCounts()` and renders bar chart |
| `deleteAgent(index)` | Removes agent from list with confirmation |
| `showAddAgentModal()` | Opens Add Agent modal |
| `confirmAddAgent()` | Validates and adds new agent to list |
| `importCsv()` | Bulk imports Name,Email pairs from textarea |
| `buildCycleLabel(date)` | Mirrors backend cycle label logic in JS |
| `toggleDark()` | Switches dark mode via `data-theme="dark"` on `<html>` |

#### Communication Pattern
All backend calls use `google.script.run`:
```javascript
google.script.run
  .withSuccessHandler(function(result) { /* handle */ })
  .withFailureHandler(function(e) { showError(e); })
  .functionName(args);
```

---

## Agent List
68 agents hardcoded as fallback in `getDefaultAgents()`. Agent names in the Excel file (Column I) must match these names exactly (case-insensitive, trimmed) for filtering to work.

Notable agents to be aware of:
- **WALKIN** — not a real agent, represents walk-in/unassigned policies, leave email blank
- **MARY SEVENZAI** — kept (MARY SEVEDZAI was a duplicate typo, removed)
- Several agents have **"-P.E"** suffix (Port Elizabeth branch)

---

## Excel File Format
- Uploaded to **Tanya Automation** Drive folder
- Column I (index 8) = Agent Name
- Column J (index 9) = Status (all lapsed, no filtering needed)
- First row = headers
- ~6,948 rows across 68 agents (April 2026 sample)

---

## PDF Format
Each agent receives a PDF with:
- NFS logo (base64) + "LAPSED POLICY REPORT" header
- Meta bar: Agent Name | Period | Total Policies | Prepared By
- Full data table with NFS navy (`#1a3080`) header rows
- Footer with company name + generation timestamp
- Styled with inline CSS (no external stylesheets)

---

## Deployment
- **Platform:** Google Apps Script Web App
- **Execute as:** Me (Tanashe14@gmail.com — current owner)
- **Access:** Only myself
- **Owner:** Tanashe14@gmail.com (Tanyaradzwa)
- **Editor:** Panashe (panasheb85@gmail.com) — can edit code, cannot deploy
- **Deployment:** Tanya must redeploy after each code change (Deploy → Manage Deployments → New Version)

---

## Auto-Send Trigger
- Runs daily at 9am via Apps Script time-based trigger
- `autoSend()` checks if today is the 1st or 15th — only sends on those days
- Setup: run `setupAutoSendTrigger()` once manually
- Remove: run `removeAutoSendTrigger()`
- Check status: run `checkTriggerStatus()`

---

## Known Issues / History
- `UrlFetchApp.fetch()` was removed after persistent permission errors — do not reintroduce
- `Drive.Files.export()` with `alt=media` was tried and failed — do not use
- CSV attachment was a working intermediate solution before PDF was implemented
- The `script.external_request` OAuth scope was attempted but never resolved — avoid

---

## Future Features (Not Yet Built)
- Archive file after successful send (move to Tanya Automation/Archive subfolder)
- Duplicate file warning if multiple Excel files in folder
- Bounce detection via Gmail inbox scan
- WhatsApp notification alongside email
- Migrate corporate email SMTP (blocked by Nyaradzo IT — SMTP AUTH not enabled)

---

## Files to Load
When working on this project in Claude Code, always load both files together:
- `Code.gs` — backend
- `Index.html` — frontend

Changes to the send logic, PDF styling, agent management, or cycle tracking are in `Code.gs`.
Changes to the UI, tabs, modals, styling, or client-side behaviour are in `Index.html`.
