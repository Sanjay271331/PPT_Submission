// ============================================================================
// ⚙️ SINGLE Code.gs — Auth + Submission + Resubmission + Deadline Control
// ============================================================================
//
// 📌 SETUP INSTRUCTIONS:
//    1. Open your Google Sheet: https://docs.google.com/spreadsheets/d/19gOqIXlAXiyE6c0GnhbkkB7Tv1jHlFuw5wmxkI4DVrs/edit
//    2. Click: Extensions → Apps Script
//    3. Paste this entire code (replacing everything).
//    4. (Optional) If you have a specific Drive Folder ID, put it in DRIVE_FOLDER_ID below.
//       If you leave it blank (''), the script will AUTOMATICALLY create a folder
//       named "NextGen_Buildathon_Submissions" in your Google Drive!
//    5. Click Deploy → Manage deployments → Edit (pencil icon)
//       → Version: "New version" → Click "Deploy".
//
// 🛑 SUBMISSION DEADLINE CONTROL:
//    - The script automatically creates/checks a "Config" tab with a "Submission Status" column.
//    - You can only enter 1 (Open) or 0 (Closed).
//    - Setting it to 0 immediately displays "Submission deadline has closed" on the portal
//      and blocks both new submissions and resubmissions!
//    - You can also add a column named "Submission Status" in the "Auth" or "Submission" tab.
//
// ============================================================================

// 📁 Optional Google Drive folder ID. Leave as '' to auto-create "NextGen_Buildathon_Submissions"
const DRIVE_FOLDER_ID = '';

// Tab names inside THIS spreadsheet
const AUTH_SHEET_NAME       = 'Auth';        // Tab with registration data (READ ONLY)
const SUBMISSION_SHEET_NAME = 'Submission';  // Tab for storing & verifying submissions
const CONFIG_SHEET_NAME     = 'Config';      // Tab for portal configuration & deadline status

// Column indices (1-based) inside the "Auth" tab
// Mapped to: D = LEADER'S EMAIL (col 4), Y = TEAM SECRET CODE (col 25), Z = TEAM ID (col 26)
const COL_AUTH_TEAM_ID = 26; // Column Z — TEAM ID
const COL_AUTH_EMAIL   = 4;  // Column D — LEADER'S EMAIL
const COL_AUTH_SECRET  = 25; // Column Y — TEAM SECRET CODE

// ============================================================================
// 🌐  doGet — Status check & Deadline check for Frontend
// ============================================================================
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const status = _checkSubmissionStatus(ss);

  // Return JSON status if frontend queries ?action=getStatus or ?check=status
  if (e && e.parameter && (e.parameter.action === 'getStatus' || e.parameter.check === 'status')) {
    return _jsonResponse({
      success: true,
      isOpen: status === 1,
      submissionStatus: status,
      message: status === 1 ? 'Submissions are currently open.' : 'Submission deadline has closed.'
    });
  }

  // Web page status preview
  const statusBadge = status === 1
    ? '<span style="color:#22c55e;font-weight:bold;background:#dcfce7;padding:4px 10px;border-radius:6px;">OPEN (1)</span>'
    : '<span style="color:#ef4444;font-weight:bold;background:#fee2e2;padding:4px 10px;border-radius:6px;">CLOSED (0) - Deadline Passed</span>';

  return HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;padding:2.5rem;max-width:640px;margin:30px auto;line-height:1.6;color:#1e293b;background:#ffffff;border:1px solid #e2e8f0;border-radius:12px;box-shadow:0 4px 16px rgba(0,0,0,0.06);">' +
      '<h2 style="color:#0f172a;margin-top:0;">🚀 NextGen Buildathon API</h2>' +
      '<p style="margin:1rem 0;">Current Portal Status: ' + statusBadge + '</p>' +
      '<p style="color:#64748b;font-size:0.92rem;">To toggle the deadline, open the <strong>' + CONFIG_SHEET_NAME + '</strong> tab in your spreadsheet and change <strong>Submission Status</strong> to <strong>1</strong> (open) or <strong>0</strong> (closed).</p>' +
    '</div>'
  ).setTitle('NextGen Buildathon API');
}

// ============================================================================
// 📨  doPost — Handles Initial Submissions & Resubmissions
// ============================================================================
function doPost(e) {
  // ------------------------------------------------------------------
  // 0. Parse incoming JSON payload
  // ------------------------------------------------------------------
  let payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return _jsonResponse({ success: false, message: 'Error: Malformed request payload.' });
  }

  const {
    teamName,
    teamId,
    email,
    domain,
    secretCode,
    fileName,
    fileBase64,
    action // 'submit' or 'resubmit'
  } = payload;

  const isResubmit = (action === 'resubmit');

  // Validate required inputs
  if (!teamId || !email || !secretCode || !fileName || !fileBase64) {
    return _jsonResponse({ success: false, message: 'Error: Missing required fields.' });
  }

  if (!isResubmit && (!teamName || !domain)) {
    return _jsonResponse({ success: false, message: 'Error: All fields are required for initial submission.' });
  }

  // ------------------------------------------------------------------
  // 1. Acquire lock to prevent race conditions
  // ------------------------------------------------------------------
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    return _jsonResponse({
      success: false,
      message: 'Error: Server is busy. Please try again in a few seconds.'
    });
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // ----------------------------------------------------------------
    // STEP 1 — Check Global Submission Status (Deadline Control)
    // ----------------------------------------------------------------
    const currentStatus = _checkSubmissionStatus(ss);
    if (currentStatus === 0) {
      return _jsonResponse({
        success: false,
        deadlineClosed: true,
        message: 'Submission deadline has closed. Submissions and resubmissions are no longer accepted.'
      });
    }

    // ----------------------------------------------------------------
    // STEP 2 — Initialize / Map "Submission" Tab
    // ----------------------------------------------------------------
    const subContext = _getOrInitSubmissionSheet(ss);
    const subSheet   = subContext.sheet;
    const colMap     = subContext.colMap;

    // Read existing submission records
    const subData = subSheet.getDataRange().getValues();

    // ================================================================
    // BRANCH A: RESUBMISSION FLOW
    // (Check with "Submission" sheet ONLY; replace PPT/PDF Drive URL)
    // ================================================================
    if (isResubmit) {
      let existingRowIndex = -1;
      let existingRowData  = null;

      // Search for existing submission using Team ID
      for (let i = 1; i < subData.length; i++) {
        const rowTeamId = String(subData[i][colMap.teamId - 1]).trim().toLowerCase();
        if (rowTeamId === String(teamId).trim().toLowerCase()) {
          existingRowIndex = i + 1; // 1-based sheet row index
          existingRowData  = subData[i];
          break;
        }
      }

      // 1. Recheck if submitted or not
      if (existingRowIndex === -1) {
        return _jsonResponse({
          success: false,
          message: 'Error: No previous submission found for Team ID "' + teamId + '". Please use "Submit Idea" first.'
        });
      }

      // 2. Verify with "Submission" sheet ONLY using Team ID, Secret Code, and Gmail
      const storedEmail  = String(existingRowData[colMap.email - 1]).trim().toLowerCase();
      const storedSecret = colMap.secret ? String(existingRowData[colMap.secret - 1]).trim() : '';

      const inputEmail   = String(email).trim().toLowerCase();
      const inputSecret  = String(secretCode).trim();

      // Check Email (Gmail)
      if (storedEmail !== inputEmail) {
        return _jsonResponse({
          success: false,
          message: 'Error: Email does not match the registered leader email for Team ID "' + teamId + '".'
        });
      }

      // Check Secret Code against submission sheet
      if (storedSecret !== '') {
        if (storedSecret !== inputSecret) {
          return _jsonResponse({
            success: false,
            message: 'Error: Invalid Secret Code for Team ID "' + teamId + '".'
          });
        }
      } else {
        // Fallback for legacy rows submitted before the Secret Code column was saved:
        // check with Auth sheet once to backfill
        let legacyVerified = false;
        try {
          legacyVerified = _verifyAuthCredentials(ss, teamId, email, secretCode);
        } catch (e) {}

        if (!legacyVerified) {
          return _jsonResponse({
            success: false,
            message: 'Error: Invalid Secret Code or Email for Team ID "' + teamId + '".'
          });
        }
      }

      // 3. Upload new PDF to Google Drive
      let fileUrl;
      try {
        fileUrl = _uploadPdfToDrive(fileName, fileBase64);
      } catch (uploadErr) {
        return _jsonResponse({
          success: false,
          message: 'Error: Failed to upload replacement PDF — ' + uploadErr.message
        });
      }

      // 4. Replace the PPT in that PPT column in the Submission sheet
      subSheet.getRange(existingRowIndex, colMap.pdf).setValue(fileUrl);
      subSheet.getRange(existingRowIndex, colMap.timestamp).setValue(new Date());

      if (teamName && colMap.teamName) {
        subSheet.getRange(existingRowIndex, colMap.teamName).setValue(teamName);
      }
      if (domain && colMap.domain) {
        subSheet.getRange(existingRowIndex, colMap.domain).setValue(domain);
      }
      if (colMap.secret) {
        subSheet.getRange(existingRowIndex, colMap.secret).setValue(inputSecret);
      }
      if (colMap.status) {
        subSheet.getRange(existingRowIndex, colMap.status).setValue(1);
      }

      return _jsonResponse({
        success: true,
        isResubmit: true,
        message: 'Your idea document has been successfully replaced and resubmitted! 🎉'
      });
    }

    // ================================================================
    // BRANCH B: INITIAL SUBMISSION FLOW
    // (Anti-duplication check + Auth sheet verification)
    // ================================================================

    // 1. Anti-Duplication Check: prevent duplicate submissions
    for (let i = 1; i < subData.length; i++) {
      const rowTeamId = String(subData[i][colMap.teamId - 1]).trim().toLowerCase();
      if (rowTeamId === String(teamId).trim().toLowerCase()) {
        return _jsonResponse({
          success: false,
          message: 'Error: Your team has already submitted an idea. Please use the "Resubmit Idea" button to update your submission.'
        });
      }
    }

    // 2. Credential Verification against "Auth" tab
    const authVerified = _verifyAuthCredentials(ss, teamId, email, secretCode);
    if (!authVerified) {
      return _jsonResponse({
        success: false,
        message: 'Error: Invalid Team ID, Email, or Secret Code combination.'
      });
    }

    // 3. PDF Upload to Google Drive
    let fileUrl;
    try {
      fileUrl = _uploadPdfToDrive(fileName, fileBase64);
    } catch (uploadErr) {
      return _jsonResponse({
        success: false,
        message: 'Error: Failed to upload PDF — ' + uploadErr.message
      });
    }

    // 4. Append to "Submission" tab (including Secret Code for future resubmissions)
    const totalCols = Math.max(subSheet.getLastColumn(), 8);
    const newRow = [];

    for (let c = 1; c <= totalCols; c++) {
      if (c === colMap.timestamp) newRow.push(new Date());
      else if (c === colMap.teamName) newRow.push(teamName);
      else if (c === colMap.teamId) newRow.push(teamId);
      else if (c === colMap.email) newRow.push(email);
      else if (c === colMap.domain) newRow.push(domain);
      else if (c === colMap.pdf) newRow.push(fileUrl);
      else if (c === colMap.secret) newRow.push(secretCode);
      else if (c === colMap.status) newRow.push(1);
      else newRow.push('');
    }

    subSheet.appendRow(newRow);

    return _jsonResponse({
      success: true,
      isResubmit: false,
      message: 'Submission successfully completed! 🎉'
    });

  } catch (fatalErr) {
    return _jsonResponse({
      success: false,
      message: 'Error: An unexpected server error occurred — ' + fatalErr.message
    });
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// 🔒 Helper — Check Submission Status (Deadline Control: 1 or 0)
// ============================================================================
function _checkSubmissionStatus(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Check "Config" or "Settings" tab
  let configSheet = ss.getSheetByName(CONFIG_SHEET_NAME) || ss.getSheetByName('Settings');
  if (configSheet) {
    const data = configSheet.getDataRange().getValues();
    if (data.length >= 2) {
      let statusCol = 0;
      for (let c = 0; c < data[0].length; c++) {
        const header = String(data[0][c]).toLowerCase();
        if (header.includes('status')) {
          statusCol = c;
          break;
        }
      }
      const rawVal = String(data[1][statusCol]).trim();
      return (rawVal === '0' || rawVal.toLowerCase() === 'closed' || rawVal.toLowerCase() === 'false') ? 0 : 1;
    }
  }

  // 2. Check "Auth" tab for a column named "Submission Status"
  const authSheet = ss.getSheetByName(AUTH_SHEET_NAME) || ss.getSheetByName('Sheet1');
  if (authSheet) {
    const lastCol = Math.max(authSheet.getLastColumn(), 1);
    const headers = authSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    for (let c = 0; c < headers.length; c++) {
      const h = String(headers[c]).trim().toLowerCase();
      if (h === 'submission status' || h === 'submission_status' || h === 'portal status' || h === 'deadline status') {
        const rawVal = String(authSheet.getRange(2, c + 1).getValue()).trim();
        return (rawVal === '0' || rawVal.toLowerCase() === 'closed' || rawVal.toLowerCase() === 'false') ? 0 : 1;
      }
    }
  }

  // 3. Check "Submission" tab for a column named "Submission Status"
  const subSheet = ss.getSheetByName(SUBMISSION_SHEET_NAME);
  if (subSheet) {
    const lastCol = Math.max(subSheet.getLastColumn(), 1);
    const headers = subSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    for (let c = 0; c < headers.length; c++) {
      const h = String(headers[c]).trim().toLowerCase();
      if (h === 'submission status' || h === 'submission_status' || h === 'portal status' || h === 'deadline status') {
        const rawVal = String(subSheet.getRange(2, c + 1).getValue()).trim();
        return (rawVal === '0' || rawVal.toLowerCase() === 'closed' || rawVal.toLowerCase() === 'false') ? 0 : 1;
      }
    }
  }

  // 4. Auto-create Config tab with Data Validation (1 or 0 only) if not found
  try {
    configSheet = ss.insertSheet(CONFIG_SHEET_NAME);
    configSheet.appendRow(['Submission Status', 'Instructions']);
    configSheet.appendRow([1, 'Enter 1 for Open, 0 for Closed (shows "Submission deadline has closed")']);

    // Style header
    configSheet.getRange('A1:B1').setFontWeight('bold').setBackground('#f1f5f9');
    configSheet.getRange('A2').setHorizontalAlignment('center');

    // Data Validation: strictly enforce 1 or 0 only
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['1', '0'], true)
      .setAllowInvalid(false)
      .setHelpText('Only 1 (Open) or 0 (Closed) is allowed.')
      .build();
    configSheet.getRange('A2:A50').setDataValidation(rule);
  } catch (err) {
    // Handled if concurrent creation
  }

  return 1; // Default to open
}

// ============================================================================
// 📊 Helper — Get or Initialize "Submission" Sheet & Column Map
// ============================================================================
function _getOrInitSubmissionSheet(ss) {
  let subSheet = ss.getSheetByName(SUBMISSION_SHEET_NAME);
  const defaultHeaders = [
    'Timestamp',
    'Team Name',
    'Team ID',
    'Leader Mail',
    'Domain / Track',
    'PDF Drive URL',
    'Team Secret Code',
    'Status'
  ];

  if (!subSheet) {
    subSheet = ss.insertSheet(SUBMISSION_SHEET_NAME);
    subSheet.appendRow(defaultHeaders);
    subSheet.getRange(1, 1, 1, defaultHeaders.length).setFontWeight('bold').setBackground('#f1f5f9');
    subSheet.setFrozenRows(1);

    return {
      sheet: subSheet,
      colMap: {
        timestamp: 1,
        teamName: 2,
        teamId: 3,
        email: 4,
        domain: 5,
        pdf: 6,
        secret: 7,
        status: 8
      }
    };
  }

  // If subSheet exists, read row 1 headers and map column indices dynamically
  const lastCol = Math.max(subSheet.getLastColumn(), 1);
  const headers = subSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());

  let colMap = {
    timestamp: 0,
    teamName: 0,
    teamId: 0,
    email: 0,
    domain: 0,
    pdf: 0,
    secret: 0,
    status: 0
  };

  for (let c = 0; c < headers.length; c++) {
    const h = headers[c].toLowerCase();
    if (h.includes('team id') || h === 'teamid') colMap.teamId = c + 1;
    else if (h.includes('team name') || h === 'teamname') colMap.teamName = c + 1;
    else if (h.includes('mail') || h.includes('email') || h.includes('gmail')) colMap.email = c + 1;
    else if (h.includes('domain') || h.includes('track')) colMap.domain = c + 1;
    else if (h.includes('pdf') || h.includes('ppt') || h.includes('drive url') || h.includes('url')) colMap.pdf = c + 1;
    else if (h.includes('secret')) colMap.secret = c + 1;
    else if (h === 'status') colMap.status = c + 1;
    else if (h.includes('time') || h.includes('date')) colMap.timestamp = c + 1;
  }

  // Defaults if headers weren't named standardly
  if (!colMap.timestamp) colMap.timestamp = 1;
  if (!colMap.teamName)  colMap.teamName  = 2;
  if (!colMap.teamId)    colMap.teamId    = 3;
  if (!colMap.email)     colMap.email     = 4;
  if (!colMap.domain)    colMap.domain    = 5;
  if (!colMap.pdf)       colMap.pdf       = 6;

  // If 'Team Secret Code' column is missing in row 1, add it dynamically
  if (!colMap.secret) {
    const newCol = subSheet.getLastColumn() + 1;
    subSheet.getRange(1, newCol).setValue('Team Secret Code').setFontWeight('bold');
    colMap.secret = newCol;
  }

  // If 'Status' column is missing, add it
  if (!colMap.status) {
    colMap.status = colMap.secret + 1;
    if (subSheet.getLastColumn() < colMap.status) {
      subSheet.getRange(1, colMap.status).setValue('Status').setFontWeight('bold');
    }
  }

  return { sheet: subSheet, colMap: colMap };
}

// ============================================================================
// 🔑 Helper — Verify Credentials in "Auth" Tab (Read Only)
// ============================================================================
function _verifyAuthCredentials(ss, teamId, email, secretCode) {
  let authSheet = ss.getSheetByName(AUTH_SHEET_NAME);
  if (!authSheet) {
    authSheet = ss.getSheetByName('Sheet1') || ss.getSheets()[0];
  }

  if (!authSheet) {
    throw new Error('Registration Auth sheet tab not found.');
  }

  const authData = authSheet.getDataRange().getValues();

  for (let i = 1; i < authData.length; i++) {
    const rowTeamId = String(authData[i][COL_AUTH_TEAM_ID - 1]).trim().toLowerCase();

    if (rowTeamId === String(teamId).trim().toLowerCase()) {
      const rowEmail  = String(authData[i][COL_AUTH_EMAIL - 1]).trim().toLowerCase();
      const rowSecret = String(authData[i][COL_AUTH_SECRET - 1]).trim();

      if (
        rowEmail  === String(email).trim().toLowerCase() &&
        rowSecret === String(secretCode).trim()
      ) {
        return true;
      }
      return false; // Found Team ID but email or secret mismatch
    }
  }

  return false; // Team ID not found in Auth tab
}

// ============================================================================
// ☁️ Helper — PDF Upload to Google Drive
// ============================================================================
function _uploadPdfToDrive(fileName, fileBase64) {
  const decodedBytes = Utilities.base64Decode(fileBase64);
  const blob = Utilities.newBlob(decodedBytes, 'application/pdf', fileName);

  let folder;
  if (DRIVE_FOLDER_ID && DRIVE_FOLDER_ID.trim() !== '' && DRIVE_FOLDER_ID !== 'YOUR_DRIVE_FOLDER_ID_HERE') {
    try {
      folder = DriveApp.getFolderById(DRIVE_FOLDER_ID.trim());
    } catch (fErr) {
      folder = null;
    }
  }

  // If no folder provided, auto-create or use "NextGen_Buildathon_Submissions" folder
  if (!folder) {
    const folderName = 'NextGen_Buildathon_Submissions';
    const existingFolders = DriveApp.getFoldersByName(folderName);
    if (existingFolders.hasNext()) {
      folder = existingFolders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }
  }

  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// ============================================================================
// 🛠  Helper — Build a JSON ContentService response
// ============================================================================
function _jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
