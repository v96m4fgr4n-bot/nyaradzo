/**
 * NFS Policy Mailer — Backend
 * Nyaradzo Financial Services — bi-weekly lapsed policy distribution.
 *
 * No UrlFetchApp, no Drive.Files.export(), no external libraries.
 * PDFs are built via Utilities.newBlob(html, 'text/html').getAs('application/pdf').
 */

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const FOLDER_NAME = "Tanya Automation";
const SETTINGS_SPREADSHEET_NAME = "Nyaradzo Mailer Settings";
const SETTINGS_SHEET = "AgentEmails";
const LOG_SHEET = "SendLog";
const SENDER_NAME = "Tanyaradzwa Manyeruke";
const COMPANY_NAME = "Nyaradzo Financial Services";
const SUMMARY_EMAIL = "panasheb85@gmail.com";
const REPLY_TO_EMAIL = "tanyaradzwa.manyeruke@nyaradzo.co.za";
const FROM_ALIAS = "Tanashe14@gmail.com";
const AGENT_NAME_COLUMN = 8; // Column I (0-indexed)

// 1x1 transparent PNG placeholder — replace with the real NFS logo base64 string.
const LOGO_B64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";

// ---------------------------------------------------------------------------
// Web app entry point
// ---------------------------------------------------------------------------

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("NFS Policy Mailer")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// ---------------------------------------------------------------------------
// Settings spreadsheet helpers
// ---------------------------------------------------------------------------

function getSettingsSpreadsheet_() {
  var files = DriveApp.getFilesByName(SETTINGS_SPREADSHEET_NAME);
  var ss;
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
  } else {
    ss = SpreadsheetApp.create(SETTINGS_SPREADSHEET_NAME);
  }
  getOrCreateSheet_(ss, SETTINGS_SHEET, ["AgentName", "Email"]);
  getOrCreateSheet_(ss, LOG_SHEET, [
    "Timestamp",
    "AgentName",
    "Email",
    "PoliciesCount",
    "Status",
    "CycleLabel",
    "CycleKey"
  ]);
  var defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet && ss.getSheets().length > 2) {
    ss.deleteSheet(defaultSheet);
  }
  return ss;
}

function getOrCreateSheet_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getTanyaFolder_() {
  var folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(FOLDER_NAME);
}

// ---------------------------------------------------------------------------
// Agent directory
// ---------------------------------------------------------------------------

/**
 * Minimal placeholder roster. The real 68-agent list (including entries like
 * WALKIN and MARY SEVENZAI, some with "-P.E" suffixes) should be entered via
 * the Agents tab (Add Agent / Import CSV) — it is not fabricated here.
 */
function getDefaultAgents() {
  return [{ name: "WALKIN", email: "" }];
}

function getAgents() {
  var ss = getSettingsSpreadsheet_();
  var sheet = ss.getSheetByName(SETTINGS_SHEET);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    var defaults = getDefaultAgents();
    saveAgents(defaults);
    return defaults;
  }
  var agents = [];
  for (var i = 1; i < data.length; i++) {
    var name = data[i][0];
    if (!name) continue;
    agents.push({
      name: String(name).trim(),
      email: data[i][1] ? String(data[i][1]).trim() : ""
    });
  }
  return agents;
}

function saveAgents(agentList) {
  var ss = getSettingsSpreadsheet_();
  var sheet = ss.getSheetByName(SETTINGS_SHEET);
  sheet.clearContents();
  sheet.appendRow(["AgentName", "Email"]);
  agentList.forEach(function (a) {
    sheet.appendRow([a.name, a.email || ""]);
  });
  sheet.setFrozenRows(1);
  return { success: true, count: agentList.length };
}

// ---------------------------------------------------------------------------
// Drive file listing / Excel -> Sheet conversion
// ---------------------------------------------------------------------------

function getFilesInFolder() {
  var folder = getTanyaFolder_();
  var files = [];
  var iterator = folder.getFiles();
  while (iterator.hasNext()) {
    var file = iterator.next();
    var mime = file.getMimeType();
    var isExcel =
      mime === MimeType.MICROSOFT_EXCEL ||
      mime === MimeType.MICROSOFT_EXCEL_LEGACY ||
      mime === MimeType.GOOGLE_SHEETS ||
      /\.xlsx?$/i.test(file.getName());
    if (isExcel) {
      files.push({
        id: file.getId(),
        name: file.getName(),
        lastUpdated: file.getLastUpdated().toISOString()
      });
    }
  }
  files.sort(function (a, b) {
    return new Date(b.lastUpdated) - new Date(a.lastUpdated);
  });
  return files;
}

function convertExcelToSheet_(fileId) {
  var resource = {
    title: "TEMP_" + fileId + "_" + new Date().getTime(),
    mimeType: MimeType.GOOGLE_SHEETS
  };
  var copiedFile = Drive.Files.copy(resource, fileId);
  return copiedFile.id;
}

function readAgentRows_(fileId) {
  var tempId = convertExcelToSheet_(fileId);
  try {
    var ss = SpreadsheetApp.openById(tempId);
    var sheet = ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();
    var headers = data.length ? data[0] : [];
    var rowsByAgent = {};
    for (var i = 1; i < data.length; i++) {
      var agentName = data[i][AGENT_NAME_COLUMN];
      if (!agentName) continue;
      var key = String(agentName).trim().toUpperCase();
      if (!rowsByAgent[key]) rowsByAgent[key] = [];
      rowsByAgent[key].push(data[i]);
    }
    return { headers: headers, rowsByAgent: rowsByAgent };
  } finally {
    DriveApp.getFileById(tempId).setTrashed(true);
  }
}

// ---------------------------------------------------------------------------
// Preview
// ---------------------------------------------------------------------------

function getAgentPolicyCounts(fileId) {
  var parsed = readAgentRows_(fileId);
  var result = Object.keys(parsed.rowsByAgent).map(function (name) {
    return { agentName: name, count: parsed.rowsByAgent[name].length };
  });
  result.sort(function (a, b) {
    return b.count - a.count;
  });
  return result;
}

// ---------------------------------------------------------------------------
// Cycle helpers
// ---------------------------------------------------------------------------

function getBiWeeklyCycleKey(date) {
  var d = date || new Date();
  var half = d.getDate() <= 14 ? "A" : "B";
  var monthStr = ("0" + (d.getMonth() + 1)).slice(-2);
  return d.getFullYear() + "-" + monthStr + "-" + half;
}

function getBiWeeklyCycleLabel(date) {
  var d = date || new Date();
  var monthNames = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
  ];
  var monthName = monthNames[d.getMonth()];
  var startDay, endDay;
  if (d.getDate() <= 14) {
    startDay = 1;
    endDay = 14;
  } else {
    startDay = 15;
    endDay = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
  }
  function pad(n) {
    return ("0" + n).slice(-2);
  }
  return pad(startDay) + " " + monthName + " – " + pad(endDay) + " " + monthName + " " + d.getFullYear();
}

// ---------------------------------------------------------------------------
// Send flow
// ---------------------------------------------------------------------------

function sendEmails(selectedAgents, fileId, cycleLabel) {
  var agents = getAgents();
  var agentMap = {};
  agents.forEach(function (a) {
    agentMap[a.name.toUpperCase()] = a;
  });

  var parsed = readAgentRows_(fileId);
  var headers = parsed.headers;
  var rowsByAgent = parsed.rowsByAgent;
  var cycleKey = getBiWeeklyCycleKey(new Date());
  var results = [];

  selectedAgents.forEach(function (agentName) {
    var key = agentName.trim().toUpperCase();
    var agent = agentMap[key];
    var rows = rowsByAgent[key] || [];

    if (key === "WALKIN") {
      results.push({
        agentName: agentName,
        email: "",
        policiesCount: rows.length,
        status: "SKIPPED",
        message: "Walk-in — no email sent"
      });
      logSendResult_(agentName, "", rows.length, "SKIPPED", cycleLabel, cycleKey);
      return;
    }

    if (!agent || !agent.email) {
      results.push({
        agentName: agentName,
        email: "",
        policiesCount: rows.length,
        status: "FAILED",
        message: "No email on file"
      });
      logSendResult_(agentName, "", rows.length, "FAILED", cycleLabel, cycleKey);
      return;
    }

    if (rows.length === 0) {
      results.push({
        agentName: agentName,
        email: agent.email,
        policiesCount: 0,
        status: "SKIPPED",
        message: "No lapsed policies this cycle"
      });
      logSendResult_(agentName, agent.email, 0, "SKIPPED", cycleLabel, cycleKey);
      return;
    }

    try {
      var pdfBlob = buildPdfBlob(agentName, cycleLabel, headers, rows);
      GmailApp.sendEmail(agent.email, "Lapsed Policy Report — " + cycleLabel, buildEmailBody_(agentName, cycleLabel, rows.length), {
        name: SENDER_NAME,
        from: FROM_ALIAS,
        replyTo: REPLY_TO_EMAIL,
        attachments: [pdfBlob]
      });
      results.push({
        agentName: agentName,
        email: agent.email,
        policiesCount: rows.length,
        status: "SENT",
        message: ""
      });
      logSendResult_(agentName, agent.email, rows.length, "SENT", cycleLabel, cycleKey);
    } catch (err) {
      results.push({
        agentName: agentName,
        email: agent.email,
        policiesCount: rows.length,
        status: "FAILED",
        message: err && err.message ? err.message : String(err)
      });
      logSendResult_(agentName, agent.email, rows.length, "FAILED", cycleLabel, cycleKey);
    }
  });

  sendSummaryEmail_(results, cycleLabel);
  return results;
}

function retryFailedAgents(fileId, cycleLabel) {
  var cycleKey = getBiWeeklyCycleKey(new Date());
  var log = getSendHistoryRows_();
  var failedAgents = [];
  log.forEach(function (row) {
    if (row.cycleKey === cycleKey && row.status === "FAILED") {
      failedAgents.push(row.agentName);
    }
  });
  if (failedAgents.length === 0) return [];
  return sendEmails(failedAgents, fileId, cycleLabel);
}

function buildEmailBody_(agentName, cycleLabel, count) {
  return (
    "Dear " + agentName + ",\n\n" +
    "Please find attached your lapsed funeral policy report for the period " + cycleLabel + " (" + count + " polic" + (count === 1 ? "y" : "ies") + ").\n\n" +
    "Kind regards,\n" + SENDER_NAME + "\n" + COMPANY_NAME
  );
}

// ---------------------------------------------------------------------------
// PDF generation
// ---------------------------------------------------------------------------

function escapeHtml_(value) {
  if (value === null || value === undefined) return "";
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function buildPdfBlob(agentName, cycleLabel, headers, rows) {
  var headerHtml = headers
    .map(function (h) {
      return "<th>" + escapeHtml_(h) + "</th>";
    })
    .join("");

  var rowsHtml = rows
    .map(function (row) {
      var cells = row
        .map(function (cell) {
          return "<td>" + escapeHtml_(cell) + "</td>";
        })
        .join("");
      return "<tr>" + cells + "</tr>";
    })
    .join("");

  var html =
    "<html><head><style>" +
    "body{font-family:Arial,sans-serif;color:#1a1a1a;margin:0;padding:24px;}" +
    ".header{display:flex;align-items:center;border-bottom:3px solid #1a3080;padding-bottom:12px;margin-bottom:16px;}" +
    ".header img{height:44px;margin-right:16px;}" +
    ".header h1{color:#1a3080;font-size:20px;margin:0;letter-spacing:0.5px;}" +
    ".meta{display:flex;justify-content:space-between;background:#f2f4fb;border:1px solid #1a3080;border-radius:6px;padding:10px 16px;margin-bottom:16px;font-size:12px;}" +
    "table{width:100%;border-collapse:collapse;font-size:10px;}" +
    "th{background:#1a3080;color:#fff;padding:6px;text-align:left;}" +
    "td{padding:5px 6px;border-bottom:1px solid #ddd;}" +
    "tr:nth-child(even){background:#f7f8fc;}" +
    ".footer{margin-top:20px;font-size:9px;color:#666;text-align:center;border-top:1px solid #ddd;padding-top:8px;}" +
    "</style></head><body>" +
    '<div class="header">' +
    (LOGO_B64 ? '<img src="data:image/png;base64,' + LOGO_B64 + '">' : "") +
    "<h1>LAPSED POLICY REPORT</h1>" +
    "</div>" +
    '<div class="meta">' +
    "<span><strong>Agent:</strong> " + escapeHtml_(agentName) + "</span>" +
    "<span><strong>Period:</strong> " + escapeHtml_(cycleLabel) + "</span>" +
    "<span><strong>Total Policies:</strong> " + rows.length + "</span>" +
    "<span><strong>Prepared By:</strong> " + SENDER_NAME + "</span>" +
    "</div>" +
    "<table><thead><tr>" + headerHtml + "</tr></thead><tbody>" + rowsHtml + "</tbody></table>" +
    '<div class="footer">' + COMPANY_NAME + " — Generated " + new Date().toLocaleString() + "</div>" +
    "</body></html>";

  var blob = Utilities.newBlob(html, "text/html", "report.html").getAs("application/pdf");
  var safeName = (agentName + "_" + cycleLabel).replace(/[^a-z0-9]/gi, "_");
  blob.setName(safeName + ".pdf");
  return blob;
}

// ---------------------------------------------------------------------------
// Logging / status / history
// ---------------------------------------------------------------------------

function logSendResult_(agentName, email, policiesCount, status, cycleLabel, cycleKey) {
  var ss = getSettingsSpreadsheet_();
  var sheet = ss.getSheetByName(LOG_SHEET);
  sheet.appendRow([new Date(), agentName, email, policiesCount, status, cycleLabel, cycleKey]);
}

function getSendHistoryRows_() {
  var ss = getSettingsSpreadsheet_();
  var sheet = ss.getSheetByName(LOG_SHEET);
  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    rows.push({
      timestamp: data[i][0],
      agentName: data[i][1],
      email: data[i][2],
      policiesCount: data[i][3],
      status: data[i][4],
      cycleLabel: data[i][5],
      cycleKey: data[i][6]
    });
  }
  return rows;
}

function getStatusSummary() {
  var now = new Date();
  var cycleKey = getBiWeeklyCycleKey(now);
  var cycleLabel = getBiWeeklyCycleLabel(now);
  var agents = getAgents();
  var log = getSendHistoryRows_();

  var latestByAgent = {};
  log.forEach(function (row) {
    if (row.cycleKey === cycleKey) {
      latestByAgent[row.agentName.toUpperCase()] = row;
    }
  });

  var sent = [];
  var pending = [];
  agents.forEach(function (a) {
    var row = latestByAgent[a.name.toUpperCase()];
    if (row && row.status === "SENT") {
      sent.push({
        agentName: a.name,
        email: a.email,
        policiesCount: row.policiesCount,
        timestamp: row.timestamp
      });
    } else {
      pending.push({ agentName: a.name, email: a.email });
    }
  });

  return { cycleLabel: cycleLabel, cycleKey: cycleKey, sent: sent, pending: pending };
}

function getSendHistory() {
  var log = getSendHistoryRows_();
  var cycles = {};
  log.forEach(function (row) {
    if (!cycles[row.cycleKey]) {
      cycles[row.cycleKey] = {
        cycleKey: row.cycleKey,
        cycleLabel: row.cycleLabel,
        agentCount: 0,
        policyCount: 0,
        sentCount: 0,
        failedCount: 0
      };
    }
    var c = cycles[row.cycleKey];
    c.agentCount++;
    c.policyCount += Number(row.policiesCount) || 0;
    if (row.status === "SENT") c.sentCount++;
    if (row.status === "FAILED") c.failedCount++;
  });
  var result = Object.keys(cycles).map(function (k) {
    return cycles[k];
  });
  result.sort(function (a, b) {
    return b.cycleKey.localeCompare(a.cycleKey);
  });
  return result;
}

function sendSummaryEmail_(results, cycleLabel) {
  var sent = results.filter(function (r) { return r.status === "SENT"; });
  var failed = results.filter(function (r) { return r.status === "FAILED"; });
  var skipped = results.filter(function (r) { return r.status === "SKIPPED"; });

  var body =
    "Send summary for cycle " + cycleLabel + "\n\n" +
    "Sent: " + sent.length + "\n" +
    "Failed: " + failed.length + "\n" +
    "Skipped: " + skipped.length + "\n";

  if (failed.length > 0) {
    body +=
      "\nFailed agents:\n" +
      failed
        .map(function (r) {
          return "- " + r.agentName + " (" + (r.message || "unknown error") + ")";
        })
        .join("\n") +
      "\n";
  }

  GmailApp.sendEmail(SUMMARY_EMAIL, "NFS Policy Mailer — Send Summary (" + cycleLabel + ")", body, {
    name: SENDER_NAME,
    from: FROM_ALIAS
  });
}

// ---------------------------------------------------------------------------
// Auto-send trigger
// ---------------------------------------------------------------------------

function autoSend() {
  var today = new Date().getDate();
  if (today !== 1 && today !== 15) return;

  var files = getFilesInFolder();
  if (files.length === 0) {
    GmailApp.sendEmail(
      SUMMARY_EMAIL,
      "NFS Policy Mailer — Auto-Send Skipped",
      "No Excel file was found in the \"" + FOLDER_NAME + "\" Drive folder, so auto-send did not run.",
      { name: SENDER_NAME, from: FROM_ALIAS }
    );
    return;
  }

  var latestFile = files[0];
  var agentNames = getAgents().map(function (a) {
    return a.name;
  });
  var cycleLabel = getBiWeeklyCycleLabel(new Date());
  sendEmails(agentNames, latestFile.id, cycleLabel);
}

function setupAutoSendTrigger() {
  removeAutoSendTrigger();
  ScriptApp.newTrigger("autoSend").timeBased().everyDays(1).atHour(9).create();
  return { success: true };
}

function removeAutoSendTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "autoSend") {
      ScriptApp.deleteTrigger(t);
    }
  });
  return { success: true };
}

function checkTriggerStatus() {
  var triggers = ScriptApp.getProjectTriggers().filter(function (t) {
    return t.getHandlerFunction() === "autoSend";
  });
  return { active: triggers.length > 0, count: triggers.length };
}
