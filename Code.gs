// ============================================================================
// ⚙️ SINGLE Code.gs — Auth + Submission + Resubmission + Domain Number Check + Deadline Control
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
// 🛑 SUBMISSION DEADLINE LOGIC:
//    - The script checks the "Config" tab (or "Submission Status" column).
//    - Admin enters 1 (Open) or 0 (Closed).
//    - When set to 0, both initial submissions and resubmissions are rejected!
//
// 🔑 SEQUENTIAL AUTHENTICATION (Single Row in "Auth" Tab):
//    - Step 1: Check TEAM ID (Column Z, Col 26) sequentially to find the unique team row.
//    - Step 2: Check TEAM SECRET CODE (Column Y, Col 25) in that exact row.
//    - Step 3: Check GMAIL / LEADER EMAIL (Column D, Col 4) in that exact row.
//    - Step 4: Check DOMAIN NUMBER (Column U, Col 21) in that exact row (compares domain number e.g. 1, 2, 3, 4, 5).
//    * Applied to BOTH Initial Submission and Resubmission!
//
// 🔄 RESUBMISSION LOGIC:
//    1. Authenticates against "Auth" tab sequentially (Team ID -> Secret -> Gmail -> Domain Number).
//    2. Rechecks the "Submission" sheet: confirms previous submission exists for this team.
//    3. If exists, THEN ONLY replaces the PPT file link (PDF URL), updates the timestamp,
//       and increments the Resubmission Count (+1).
//
// ============================================================================

// 📁 Optional Google Drive folder ID. Leave as '' to auto-create "NextGen_Buildathon_Submissions"
const DRIVE_FOLDER_ID = '';

// Tab names inside THIS spreadsheet
const AUTH_SHEET_NAME       = 'Auth';        // Tab with registration data (READ ONLY)
const SUBMISSION_SHEET_NAME = 'Submission';  // Tab for storing & verifying submissions
const CONFIG_SHEET_NAME     = 'Config';      // Tab for portal configuration & deadline status

// Column indices (1-based) inside the "Auth" tab
// D = LEADER'S EMAIL (col 4), U = DOMAIN (col 21), Y = TEAM SECRET CODE (col 25), Z = TEAM ID (col 26)
const COL_AUTH_TEAM_ID = 26; // Column Z — TEAM ID
const COL_AUTH_SECRET  = 25; // Column Y — TEAM SECRET CODE
const COL_AUTH_EMAIL   = 4;  // Column D — LEADER'S EMAIL
const COL_AUTH_DOMAIN  = 21; // Column U — DOMAIN

// Standard headers for "Submission" sheet
const SUBMISSION_HEADERS = [
  'Timestamp',
  'Team Name',
  'Team ID',
  'Leader Mail',
  'Domain / Track',
  'PDF Drive URL',
  'Team Secret Code',
  'Resubmission Count',
  'Status'
];

// ============================================================================
// 🌐  doGet — Status check (Deadline) & Diagnostics
// ============================================================================
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Action: getStatus (used by frontend to detect deadline on page load)
  if (e && e.parameter && (e.parameter.action === 'getStatus' || e.parameter.check === 'status')) {
    const status = _checkSubmissionStatus(ss);
    return _jsonResponse({
      success: true,
      isOpen: status === 1,
      submissionStatus: status,
      message: status === 1 ? 'Submissions are currently open.' : 'Submission deadline has closed.'
    });
  }

  // Action: inspectHeaders (developer diagnostic)
  if (e && e.parameter && e.parameter.action === 'inspectHeaders') {
    const authSheet = ss.getSheetByName(AUTH_SHEET_NAME) || ss.getSheetByName('Sheet1') || ss.getSheets()[0];
    const authHeaders = authSheet ? authSheet.getRange(1, 1, 1, Math.max(authSheet.getLastColumn(), 1)).getValues()[0] : [];
    const subSheet = ss.getSheetByName(SUBMISSION_SHEET_NAME);
    const subHeaders = subSheet ? subSheet.getRange(1, 1, 1, Math.max(subSheet.getLastColumn(), 1)).getValues()[0] : [];
    return _jsonResponse({ authHeaders, subHeaders });
  }

  // Web status page
  const status = _checkSubmissionStatus(ss);
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
  if (!teamName || !teamId || !email || !secretCode || !domain || !fileName || !fileBase64) {
    return _jsonResponse({
      success: false,
      message: 'Error: All fields (Team Name, Team ID, Email, Secret Code, Domain, and PDF file) are required.'
    });
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
    // STEP 1 — Check Submission Status (DEADLINE CONTROL)
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
    // STEP 2 — Sequential Authentication in "Auth" Tab (Single Row Check)
    // Mandatory for BOTH initial submission and resubmission:
    // 1) Team ID (first check — finds the unique row)
    // 2) Team Secret Code (verified in that exact row)
    // 3) Gmail / Leader Email (verified in that exact row)
    // 4) Domain Number (compares domain number e.g. 1, 2, 3, 4, 5)
    // ----------------------------------------------------------------
    const authResult = _verifyAuthCredentials(ss, teamId, email, secretCode, domain);
    if (!authResult.verified) {
      return _jsonResponse({
        success: false,
        message: authResult.message || 'Error: Authentication failed against registration records.'
      });
    }

    // ----------------------------------------------------------------
    // STEP 3 — Initialize / Map "Submission" Tab
    // ----------------------------------------------------------------
    const subContext = _getOrInitSubmissionSheet(ss);
    const subSheet   = subContext.sheet;
    const colMap     = subContext.colMap;

    // Read existing submissions
    const subData = subSheet.getDataRange().getValues();

    // Format domain for clean display in Submission sheet
    const domainFormatted = _formatDomainDisplay(domain);

    // ================================================================
    // BRANCH A: RESUBMISSION FLOW
    // 1. Recheck in "Submission" sheet: must already exist!
    // 2. Recheck submission record details (email, domain number).
    // 3. If exists, THEN ONLY replace PPT Drive URL, increment resubmission count.
    // ================================================================
    if (isResubmit) {
      let existingSubRowIndex = -1;
      let existingSubRowData  = null;

      // Recheck if submitted or not in the Submission sheet
      for (let i = 1; i < subData.length; i++) {
        const subTeamId = String(subData[i][colMap.teamId - 1]).trim().toLowerCase();
        if (subTeamId === String(teamId).trim().toLowerCase()) {
          existingSubRowIndex = i + 1; // 1-based row index
          existingSubRowData  = subData[i];
          break;
        }
      }

      // If team has NOT submitted before, reject resubmission:
      if (existingSubRowIndex === -1) {
        return _jsonResponse({
          success: false,
          message: 'Error: No previous submission found for Team ID "' + teamId + '". Please use "Submit Idea" first.'
        });
      }

      // Recheck Email in submission sheet
      const subEmail = String(existingSubRowData[colMap.email - 1]).trim().toLowerCase();
      if (subEmail !== String(email).trim().toLowerCase()) {
        return _jsonResponse({
          success: false,
          message: 'Error: Email does not match the leader email in your previous submission record.'
        });
      }

      // Recheck Domain Number in submission sheet (if previously stored)
      const subDomain = colMap.domain ? String(existingSubRowData[colMap.domain - 1]).trim() : '';
      if (subDomain !== '' && !_isDomainMatch(domain, subDomain)) {
        const enteredNum = _extractDomainNumber(domain);
        const storedNum  = _extractDomainNumber(subDomain);
        return _jsonResponse({
          success: false,
          message: 'Error: Selected Domain (Domain ' + (enteredNum !== null ? enteredNum : domain) +
                   ') does not match your previous submission record (Domain ' + (storedNum !== null ? storedNum : subDomain) + ').'
        });
      }

      // Upload replacement PDF to Google Drive
      let fileUrl;
      try {
        fileUrl = _uploadPdfToDrive(fileName, fileBase64);
      } catch (uploadErr) {
        return _jsonResponse({
          success: false,
          message: 'Error: Failed to upload replacement PDF — ' + uploadErr.message
        });
      }

      // Increment the number of resubmissions in the Resubmission Count column
      const currentResubmitCount = colMap.resubmissionCount
        ? (parseInt(existingSubRowData[colMap.resubmissionCount - 1], 10) || 0)
        : 0;
      const newResubmitCount = currentResubmitCount + 1;

      // REPLACE the PPT PDF Drive URL and update credentials in Submission sheet
      subSheet.getRange(existingSubRowIndex, colMap.pdf).setValue(fileUrl);
      subSheet.getRange(existingSubRowIndex, colMap.timestamp).setValue(new Date());

      if (teamName && colMap.teamName) {
        subSheet.getRange(existingSubRowIndex, colMap.teamName).setValue(teamName);
      }
      if (colMap.domain) {
        subSheet.getRange(existingSubRowIndex, colMap.domain).setValue(domainFormatted);
      }
      if (colMap.secret) {
        subSheet.getRange(existingSubRowIndex, colMap.secret).setValue(secretCode);
      }
      if (colMap.resubmissionCount) {
        subSheet.getRange(existingSubRowIndex, colMap.resubmissionCount).setValue(newResubmitCount);
      }
      if (colMap.status) {
        subSheet.getRange(existingSubRowIndex, colMap.status).setValue(1);
      }

      return _jsonResponse({
        success: true,
        isResubmit: true,
        resubmissionCount: newResubmitCount,
        message: 'Your idea document has been successfully replaced and resubmitted! 🎉 (Resubmission #' + newResubmitCount + ')'
      });
    }

    // ================================================================
    // BRANCH B: INITIAL SUBMISSION FLOW
    // ================================================================

    // Anti-Duplication Check: prevent duplicate initial submissions
    for (let i = 1; i < subData.length; i++) {
      const subTeamId = String(subData[i][colMap.teamId - 1]).trim().toLowerCase();
      if (subTeamId === String(teamId).trim().toLowerCase()) {
        return _jsonResponse({
          success: false,
          message: 'Error: Your team has already submitted an idea. Please use the "Resubmit Idea" button to update your submission.'
        });
      }
    }

    // Upload PDF to Google Drive
    let fileUrl;
    try {
      fileUrl = _uploadPdfToDrive(fileName, fileBase64);
    } catch (uploadErr) {
      return _jsonResponse({
        success: false,
        message: 'Error: Failed to upload PDF — ' + uploadErr.message
      });
    }

    // Append to "Submission" tab — showing ALL entered credentials
    const totalCols = Math.max(subSheet.getLastColumn(), SUBMISSION_HEADERS.length);
    const newRow = [];

    for (let c = 1; c <= totalCols; c++) {
      if (c === colMap.timestamp) newRow.push(new Date());
      else if (c === colMap.teamName) newRow.push(teamName);
      else if (c === colMap.teamId) newRow.push(teamId);
      else if (c === colMap.email) newRow.push(email);
      else if (c === colMap.domain) newRow.push(domainFormatted);
      else if (c === colMap.pdf) newRow.push(fileUrl);
      else if (c === colMap.secret) newRow.push(secretCode);
      else if (c === colMap.resubmissionCount) newRow.push(0); // 0 initial resubmissions
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
// 🔢 Helper — Extract Domain Number (e.g. 1, 2, 3, 4, 5)
// ============================================================================
function _extractDomainNumber(val) {
  if (val === null || val === undefined) return null;
  const s = String(val).trim();

  // Match "Domain 3: IOT and Embedded System", "Domain 3", "Domain: 3", "Domain - 3"
  const domainPrefixMatch = s.match(/domain\s*[:\s#\-]*\s*(\d+)/i) || s.match(/track\s*[:\s#\-]*\s*(\d+)/i);
  if (domainPrefixMatch) {
    return parseInt(domainPrefixMatch[1], 10);
  }

  // Standalone number or leading number like "3", "3.", "#3"
  const digitMatch = s.match(/\b(\d+)\b/) || s.match(/(\d+)/);
  if (digitMatch) {
    return parseInt(digitMatch[1], 10);
  }

  // Fallback by order of domain in dropdown (1 to 5)
  const lower = s.toLowerCase();
  if (lower.includes('nextgen') || lower.includes('emerging') || lower.includes('futuristic')) return 1;
  if (lower.includes('ai') || lower.includes('artificial') || lower.includes('intelligent')) return 2;
  if (lower.includes('iot') || lower.includes('embedded')) return 3;
  if (lower.includes('fintech') || lower.includes('finance')) return 4;
  if (lower.includes('sustainable') || lower.includes('sustainability') || lower.includes('innovation')) return 5;

  return null;
}

// ============================================================================
// 🔤 Helper — Strict Domain Number Comparison
// ============================================================================
function _isDomainMatch(d1, d2) {
  if (d1 === null || d1 === undefined || d2 === null || d2 === undefined) return false;

  const num1 = _extractDomainNumber(d1);
  const num2 = _extractDomainNumber(d2);

  // Compare strictly by domain number!
  if (num1 !== null && num2 !== null) {
    return num1 === num2;
  }

  // Fallback if no digit found: compare normalized strings
  const s1 = String(d1).trim().toLowerCase().replace(/[^a-z0-9]/g, '');
  const s2 = String(d2).trim().toLowerCase().replace(/[^a-z0-9]/g, '');
  return s1 === s2;
}

// ============================================================================
// 🏷️ Helper — Format Domain Display for Sheet
// ============================================================================
function _formatDomainDisplay(val) {
  const num = _extractDomainNumber(val);
  const names = {
    1: 'Domain 1: NEXTGEN — Emerging & Futuristic Technologies',
    2: 'Domain 2: Artificial Intelligence & Intelligent Systems',
    3: 'Domain 3: IOT and Embedded System',
    4: 'Domain 4: FinTech',
    5: 'Domain 5: Sustainable Innovation'
  };
  if (num !== null && names[num]) {
    return names[num];
  }
  return String(val || '');
}

// ============================================================================
// 🔒 Helper — Check Submission Status (DEADLINE CONTROL: 1 or 0)
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
// 📊 Helper — Get or Initialize "Submission" Sheet & Dynamic Column Map
// ============================================================================
function _getOrInitSubmissionSheet(ss) {
  let subSheet = ss.getSheetByName(SUBMISSION_SHEET_NAME);

  if (!subSheet) {
    subSheet = ss.insertSheet(SUBMISSION_SHEET_NAME);
    subSheet.appendRow(SUBMISSION_HEADERS);
    subSheet.getRange(1, 1, 1, SUBMISSION_HEADERS.length).setFontWeight('bold').setBackground('#f1f5f9');
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
        resubmissionCount: 8,
        status: 9
      }
    };
  }

  // Read current headers in row 1
  const lastCol = Math.max(subSheet.getLastColumn(), 1);
  const currentHeaders = subSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());

  // If row 1 headers are corrupted, missing, or don't start with Timestamp / contain Team ID, fix them!
  const hasTeamId = currentHeaders.some(h => h.toLowerCase().includes('team id'));
  const startsWithTimestamp = currentHeaders[0] && currentHeaders[0].toLowerCase().includes('time');

  if (!hasTeamId || !startsWithTimestamp || currentHeaders.length < 5) {
    subSheet.getRange(1, 1, 1, SUBMISSION_HEADERS.length).setValues([SUBMISSION_HEADERS]);
    subSheet.getRange(1, 1, 1, SUBMISSION_HEADERS.length).setFontWeight('bold').setBackground('#f1f5f9');
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
        resubmissionCount: 8,
        status: 9
      }
    };
  }

  // Map columns dynamically
  let colMap = {
    timestamp: 0,
    teamName: 0,
    teamId: 0,
    email: 0,
    domain: 0,
    pdf: 0,
    secret: 0,
    resubmissionCount: 0,
    status: 0
  };

  for (let c = 0; c < currentHeaders.length; c++) {
    const h = currentHeaders[c].toLowerCase();
    if (h.includes('team id') || h === 'teamid') colMap.teamId = c + 1;
    else if (h.includes('team name') || h === 'teamname') colMap.teamName = c + 1;
    else if (h.includes('mail') || h.includes('email') || h.includes('gmail')) colMap.email = c + 1;
    else if (h.includes('domain') || h.includes('track')) colMap.domain = c + 1;
    else if (h.includes('pdf') || h.includes('ppt') || h.includes('drive url') || h.includes('url')) colMap.pdf = c + 1;
    else if (h.includes('secret')) colMap.secret = c + 1;
    else if (h.includes('resubmission') || h.includes('resubmit')) colMap.resubmissionCount = c + 1;
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
  if (!colMap.secret)    colMap.secret    = 7;
  if (!colMap.resubmissionCount) colMap.resubmissionCount = 8;
  if (!colMap.status)    colMap.status    = 9;

  return { sheet: subSheet, colMap: colMap };
}

// ============================================================================
// 🔑 Helper — Sequential Verification in "Auth" Tab (Single Row)
// 1. Team ID (Sequential Search) -> finds the single unique team row
// 2. Secret Code -> verified in that single row
// 3. Gmail / Leader Email -> verified in that single row
// 4. Domain Number -> strictly verified by domain number in that single row
// ============================================================================
function _verifyAuthCredentials(ss, teamId, email, secretCode, domain) {
  let authSheet = ss.getSheetByName(AUTH_SHEET_NAME);
  if (!authSheet) {
    authSheet = ss.getSheetByName('Sheet1') || ss.getSheets()[0];
  }

  if (!authSheet) {
    return { verified: false, message: 'Error: Registration Auth sheet tab not found.' };
  }

  const authData = authSheet.getDataRange().getValues();

  // 1. FIRST: Check Team ID sequentially to find the matching row (Column Z, Col 26)
  let teamRowIndex = -1;
  let teamRowData  = null;

  for (let i = 1; i < authData.length; i++) {
    const rowTeamId = String(authData[i][COL_AUTH_TEAM_ID - 1]).trim().toLowerCase();
    if (rowTeamId === String(teamId).trim().toLowerCase()) {
      teamRowIndex = i + 1;
      teamRowData  = authData[i];
      break;
    }
  }

  // If Team ID not found in Auth sheet:
  if (teamRowIndex === -1) {
    return {
      verified: false,
      message: 'Error: Team ID "' + teamId + '" is not registered in our records.'
    };
  }

  // In that exact single row itself:
  // 2. SECOND: Check Team Secret Code (Column Y, Col 25)
  const rowSecret = String(teamRowData[COL_AUTH_SECRET - 1]).trim();
  if (rowSecret !== String(secretCode).trim()) {
    return {
      verified: false,
      message: 'Error: Invalid Secret Code for Team ID "' + teamId + '".'
    };
  }

  // 3. THIRD: Check Gmail / Leader Email (Column D, Col 4)
  const rowEmail = String(teamRowData[COL_AUTH_EMAIL - 1]).trim().toLowerCase();
  if (rowEmail !== String(email).trim().toLowerCase()) {
    return {
      verified: false,
      message: 'Error: Email does not match the registered leader email for Team ID "' + teamId + '".'
    };
  }

  // 4. FOURTH: Check Domain strictly by Domain Number (Column U, Col 21)
  const rowDomain = String(teamRowData[COL_AUTH_DOMAIN - 1]).trim();
  if (rowDomain !== '' && !_isDomainMatch(domain, rowDomain)) {
    const enteredNum  = _extractDomainNumber(domain);
    const expectedNum = _extractDomainNumber(rowDomain);
    return {
      verified: false,
      message: 'Error: Selected Domain (Domain ' + (enteredNum !== null ? enteredNum : domain) +
               ') does not match your team\'s registered Domain (Domain ' + (expectedNum !== null ? expectedNum : rowDomain) + ').'
    };
  }

  return {
    verified: true,
    teamRowIndex: teamRowIndex,
    teamRowData: teamRowData
  };
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
