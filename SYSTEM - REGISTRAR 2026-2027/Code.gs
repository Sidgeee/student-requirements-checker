/**
 * =====================================================================================
 * CSCQC REGISTRAR — GOOGLE APPS SCRIPT BACKEND
 * =====================================================================================
 * Deploy as: Web App (Execute as: Me, Access: Anyone)
 * Bind to the same Spreadsheet that holds your "Students" sheet.
 *
 * SHEET: "Students"  (row 1 = headers, must match HEADERS array below exactly)
 * SHEET: "Staff"      (row 1 = headers: username, password, dateCreated)
 * =====================================================================================
 */

const SS = SpreadsheetApp.getActiveSpreadsheet();
const STUDENTS_SHEET_NAME = "Students";
const STAFF_SHEET_NAME = "Staff";

// Column order in the "Students" sheet. Keep this in sync with the sheet header row.
const HEADERS = [
  "studentNo", "surname", "firstname", "middlename", "gender", "yearLevel",
  "section", "program", "yearEntry", "status", "previousSchool", "birthdate",
  "dateEnrolled", "operationalStatus",
  "f138", "pic", "f137", "psa", "gmc", "dp", "tor", "hd", "cog", "evalcopy",
  "qrUrl", "lastQRUpdate"
];

const CHECKLIST_FIELDS = ["f138", "pic", "f137", "psa", "gmc", "dp", "tor", "hd", "cog", "evalcopy"];
const CHECKLIST_LABELS = {
  f138: "Form 138",
  pic: "2x2 Picture",
  f137: "Form 137",
  psa: "PSA",
  gmc: "GMC",
  dp: "Diploma",
  tor: "TOR",
  hd: "HD",
  cog: "CERT. OF GRADES",
  evalcopy: "EVAL. COPY"
};

/* ===================================================================================
 * doGet — handles:
 *   ?searchID=2024-0001         -> mobile-friendly HTML profile card (QR scan target)
 *   ?action=read&query=...      -> JSON payload consumed by the dashboard
 * ================================================================================= */
function doGet(e) {
  try {
    const params = (e && e.parameter) || {};

    // --- 1. MOBILE QR SCAN ENDPOINT -------------------------------------------------
    if (params.searchID) {
      const student = findStudentByNo_(params.searchID.toString().trim());
      return HtmlService.createHtmlOutput(renderMobileProfileCard_(student))
        .setTitle("CSCQC Student Profile")
        .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    // --- 2. DASHBOARD DATA READ (existing behavior) ---------------------------------
    if (params.action === "read") {
      const query = (params.query || "").toString().trim().toUpperCase();
      const data = getAllStudents_();
      const filtered = query
        ? data.filter(s =>
            (s.studentNo || "").toUpperCase().includes(query) ||
            (s.surname || "").toUpperCase().includes(query) ||
            (s.firstname || "").toUpperCase().includes(query) ||
            (s.program || "").toUpperCase().includes(query)
          )
        : data;
      return jsonOutput_({ result: "success", data: filtered });
    }

    // --- 3. FALLBACK -----------------------------------------------------------------
    return HtmlService.createHtmlOutput(
      "<p style='font-family:sans-serif;color:#78909c;text-align:center;margin-top:40px;'>" +
      "CSCQC Registrar backend is online. No valid parameters supplied.</p>"
    );

  } catch (err) {
    // CRITICAL: never let doGet fall through without a return value —
    // that silence is exactly what produces "The script completed but did not return anything."
    return HtmlService.createHtmlOutput(
      "<p style='font-family:sans-serif;color:#d32f2f;text-align:center;margin-top:40px;'>" +
      "Server error: " + escapeHtml_(err.message) + "</p>"
    );
  }
}

/* ===================================================================================
 * doPost — handles all write actions dispatched by the dashboard.
 * ================================================================================= */
function doPost(e) {
  try {
    const params = (e && e.parameter) || {};
    const action = params.action;
    let result;

    switch (action) {
      case "enroll":
        result = handleEnroll_(params);
        break;
      case "update":
        result = handleRequirementsUpdate_(params);
        break;
      case "updateStatus":
        result = handleStatusUpdate_(params);
        break;
      case "registerQR":
        result = handleRegisterQR_(params);
        break;
      case "createAccount":
        result = handleCreateAccount_(params);
        break;
      default:
        result = { result: "error", message: "Unrecognized action: " + action };
    }

    return jsonOutput_(result);

  } catch (err) {
    return jsonOutput_({ result: "error", message: err.message });
  }
}

/* ===================================================================================
 * ACTION HANDLERS
 * ================================================================================= */

function handleEnroll_(params) {
  const sheet = getSheet_(STUDENTS_SHEET_NAME);
  const studentNo = (params.studentNo || "").toString().trim();
  if (!studentNo) return { result: "error", message: "Missing studentNo." };

  const existingRow = findRowIndexByStudentNo_(sheet, studentNo);
  const row = HEADERS.map(h => {
    if (h === "operationalStatus") return params.operationalStatus || "Active"; // new records default to Active
    if (h === "qrUrl" || h === "lastQRUpdate") return "";
    return params[h] !== undefined ? params[h] : "";
  });

  if (existingRow > 0) {
    sheet.getRange(existingRow, 1, 1, HEADERS.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }

  // Automated QR lifecycle: brand-new students are Active by default -> generate QR now.
  runQRLifecyclePipeline_(studentNo);

  return { result: "success", message: "Student enrolled.", studentNo: studentNo };
}

function handleRequirementsUpdate_(params) {
  const sheet = getSheet_(STUDENTS_SHEET_NAME);
  const studentNo = (params.studentNo || "").toString().trim();
  const rowIndex = findRowIndexByStudentNo_(sheet, studentNo);
  if (rowIndex < 1) return { result: "error", message: "Student not found: " + studentNo };

  const headerRow = HEADERS;
  CHECKLIST_FIELDS.forEach(field => {
    const col = headerRow.indexOf(field) + 1;
    sheet.getRange(rowIndex, col).setValue(params[field] === "YES" ? "YES" : "");
  });

  return { result: "success", message: "Requirements updated." };
}

function handleStatusUpdate_(params) {
  const sheet = getSheet_(STUDENTS_SHEET_NAME);
  const studentNo = (params.studentNo || "").toString().trim();
  const rowIndex = findRowIndexByStudentNo_(sheet, studentNo);
  if (rowIndex < 1) return { result: "error", message: "Student not found: " + studentNo };

  const newStatus = (params.operationalStatus || "Active").toString();
  const col = HEADERS.indexOf("operationalStatus") + 1;
  sheet.getRange(rowIndex, col).setValue(newStatus);

  // --- AUTOMATED QR STATUS LIFECYCLE ---
  // Any transition INTO "Active" automatically re-runs the QR generation pipeline.
  if (newStatus.toLowerCase() === "active") {
    runQRLifecyclePipeline_(studentNo);
  }

  return { result: "success", message: "Status updated to " + newStatus };
}

function handleRegisterQR_(params) {
  const studentNo = (params.studentNo || "").toString().trim();
  if (!studentNo) return { result: "error", message: "Missing studentNo." };
  const qrUrl = runQRLifecyclePipeline_(studentNo);
  return { result: "success", message: "QR registered.", qrUrl: qrUrl };
}

function handleCreateAccount_(params) {
  const sheet = getSheet_(STAFF_SHEET_NAME);
  const username = (params.username || "").toString().trim();
  const password = (params.password || "").toString().trim();
  if (!username || !password) return { result: "error", message: "Missing credentials." };
  sheet.appendRow([username, password, new Date()]);
  return { result: "success", message: "Account created." };
}

/* ===================================================================================
 * AUTOMATED QR GENERATION LIFECYCLE PIPELINE
 * Generates a QR PNG (server-side, via a QR image API), stores it in Drive so it has a
 * stable public URL, and stamps the sheet with the QR link + timestamp. This runs
 * automatically any time a student's operationalStatus becomes/stays "Active".
 * ================================================================================= */
function runQRLifecyclePipeline_(studentNo) {
  const sheet = getSheet_(STUDENTS_SHEET_NAME);
  const rowIndex = findRowIndexByStudentNo_(sheet, studentNo);
  if (rowIndex < 1) return "";

  const webAppUrl = ScriptApp.getService().getUrl();
  const qrTargetUrl = webAppUrl + "?searchID=" + encodeURIComponent(studentNo);

  // Generate the QR image via a public QR rendering endpoint.
  const qrApiUrl = "https://api.qrserver.com/v1/create-qr-code/?size=300x300&data="
    + encodeURIComponent(qrTargetUrl);

  let fileUrl = "";
  try {
    const response = UrlFetchApp.fetch(qrApiUrl, { muteHttpExceptions: true });
    if (response.getResponseCode() === 200) {
      const blob = response.getBlob().setName("QR_" + studentNo + ".png");
      const folder = getOrCreateQRFolder_();
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = file.getUrl();
    }
  } catch (err) {
    console.error("QR pipeline generation failed for " + studentNo + ": " + err.message);
  }

  const qrUrlCol = HEADERS.indexOf("qrUrl") + 1;
  const lastQRCol = HEADERS.indexOf("lastQRUpdate") + 1;
  sheet.getRange(rowIndex, qrUrlCol).setValue(fileUrl);
  sheet.getRange(rowIndex, lastQRCol).setValue(new Date());

  return fileUrl;
}

function getOrCreateQRFolder_() {
  const FOLDER_NAME = "CSCQC_Student_QR_Codes";
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);
}

/* ===================================================================================
 * MOBILE PROFILE CARD (fixes the "did not return anything" scanner error)
 * ================================================================================= */
function renderMobileProfileCard_(student) {
  if (!student) {
    return "<!DOCTYPE html><html><head><meta name='viewport' content='width=device-width, initial-scale=1.0'></head>" +
      "<body style='font-family:Segoe UI,Roboto,sans-serif;background:#f1f8e9;display:flex;align-items:center;" +
      "justify-content:center;height:100vh;margin:0;'>" +
      "<div style='background:#fff;padding:30px;border-radius:14px;box-shadow:0 10px 30px rgba(0,0,0,0.1);text-align:center;max-width:340px;'>" +
      "<div style='font-size:40px;color:#d32f2f;'>&#10060;</div>" +
      "<h2 style='color:#d32f2f;margin:10px 0 5px;'>Student Not Found</h2>" +
      "<p style='color:#78909c;font-size:14px;'>No record matches this QR code in the CSCQC registry.</p>" +
      "</div></body></html>";
  }

  const mi = student.middlename ? (student.middlename.toString().charAt(0).toUpperCase() + ".") : "";
  const fullName = (student.surname + ", " + student.firstname + (mi ? " " + mi : "")).toUpperCase();
  const opStatus = (student.operationalStatus || "Active").toString();
  const statusColor = opStatus.toLowerCase() === "active" ? "#2e7d32" : "#d32f2f";
  const statusBg = opStatus.toLowerCase() === "active" ? "#e8f5e9" : "#ffebee";
  const category = (student.status || "Regular").toString();

  const checklistRows = CHECKLIST_FIELDS.map(field => {
    const complete = student[field] === "YES";
    const label = CHECKLIST_LABELS[field];
    const icon = complete ? "&#9989;" : "&#10060;";
    const tagColor = complete ? "#2e7d32" : "#d32f2f";
    const tagBg = complete ? "#e8f5e9" : "#ffebee";
    const tagText = complete ? "COMPLETE" : "MISSING";
    return "<div style='display:flex;align-items:center;justify-content:space-between;padding:9px 0;" +
      "border-bottom:1px solid #f0f4f5;'>" +
      "<span style='font-size:13px;color:#37474f;font-weight:600;'>" + icon + " " + label + "</span>" +
      "<span style='font-size:10px;font-weight:800;color:" + tagColor + ";background:" + tagBg +
      ";padding:3px 9px;border-radius:5px;'>" + tagText + "</span>" +
      "</div>";
  }).join("");

  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>CSCQC Student Profile</title>
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <style>
    * { box-sizing: border-box; }
    body {
      margin: 0; padding: 20px;
      font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
      background: linear-gradient(160deg, #e8f5e9, #f1f8e9);
      min-height: 100vh;
    }
    .card {
      max-width: 420px; margin: 0 auto;
      background: #fff; border-radius: 16px;
      box-shadow: 0 15px 35px rgba(0,0,0,0.12);
      overflow: hidden;
    }
    .card-header {
      background: linear-gradient(135deg, #2e7d32, #1b5e20);
      color: #fff; padding: 22px 20px; text-align: center;
    }
    .card-header .name { font-size: 18px; font-weight: 800; letter-spacing: 0.4px; }
    .card-header .id { font-size: 12px; opacity: 0.85; margin-top: 4px; }
    .card-body { padding: 20px; }
    .info-row {
      display: flex; justify-content: space-between; align-items: center;
      padding: 8px 0; border-bottom: 1px solid #f0f4f5; font-size: 13px;
    }
    .info-row span:first-child { color: #78909c; font-weight: 600; }
    .info-row span:last-child { color: #37474f; font-weight: 700; text-align: right; }
    .status-pill {
      display: inline-block; padding: 3px 10px; border-radius: 6px;
      font-size: 11px; font-weight: 800;
    }
    .checklist-title {
      margin: 22px 0 6px; font-size: 13px; font-weight: 800;
      color: #2e7d32; text-transform: uppercase; letter-spacing: 0.5px;
    }
    .footer { text-align: center; padding: 14px; font-size: 11px; color: #90a4ae; }
  </style>
</head>
<body>
  <div class="card">
    <div class="card-header">
      <div class="name">${escapeHtml_(fullName)}</div>
      <div class="id">ID: ${escapeHtml_(student.studentNo)} &nbsp;|&nbsp; ${escapeHtml_(category)}</div>
    </div>
    <div class="card-body">
      <div class="info-row"><span>Course &amp; Section</span><span>${escapeHtml_(student.program || 'N/A')} - ${escapeHtml_(student.section || 'N/A')}</span></div>
      <div class="info-row"><span>Category</span><span>${escapeHtml_(category)}</span></div>
      <div class="info-row"><span>Date Enrolled</span><span>${escapeHtml_(formatDate_(student.dateEnrolled))}</span></div>
      <div class="info-row">
        <span>System Status</span>
        <span class="status-pill" style="color:${statusColor};background:${statusBg};">${escapeHtml_(opStatus.toUpperCase())}</span>
      </div>

      <div class="checklist-title"><i class="fa-solid fa-clipboard-check"></i> Documentary Requirements</div>
      ${checklistRows}
    </div>
    <div class="footer">CSCQC Registrar &middot; College of St. Catherine, Quezon City</div>
  </div>
</body>
</html>`;
}

/* ===================================================================================
 * DATA ACCESS HELPERS
 * ================================================================================= */

function getSheet_(name) {
  const sheet = SS.getSheetByName(name);
  if (!sheet) throw new Error("Sheet not found: " + name);
  return sheet;
}

function getAllStudents_() {
  const sheet = getSheet_(STUDENTS_SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  const headerRow = values[0];
  const rows = values.slice(1);

  return rows
    .filter(row => row[0] !== "") // skip blank rows
    .map(row => {
      const obj = {};
      headerRow.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
}

function findStudentByNo_(studentNo) {
  const all = getAllStudents_();
  return all.find(s => (s.studentNo || "").toString().trim() === studentNo) || null;
}

function findRowIndexByStudentNo_(sheet, studentNo) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if ((values[i][0] || "").toString().trim() === studentNo) return i + 1; // 1-based sheet row
  }
  return -1;
}

function formatDate_(value) {
  if (!value) return "N/A";
  try {
    const d = (value instanceof Date) ? value : new Date(value);
    if (isNaN(d.getTime())) return value.toString();
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "MMM dd, yyyy");
  } catch (e) {
    return value.toString();
  }
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function escapeHtml_(str) {
  if (str === null || str === undefined) return "";
  return str.toString()
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}