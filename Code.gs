// ============================================================================
// ⚙️  SINGLE Code.gs — Auth + Submission in ONE spreadsheet
// ============================================================================
//
// 📌 SETUP:
//    1. Your Google Sheet has TWO tabs:
//       • "Auth"       → Tab with registration data (Team ID, Email, Secret Code)
//       • "Submission"  → Tab for storing submissions (starts with headers only)
//    2. Open this Google Sheet → Extensions → Apps Script
//    3. Paste this code → Fill in DRIVE_FOLDER_ID below
//    4. Deploy → New Deployment → Web App → Execute as "Me" → Anyone
//    5. Copy the Web App URL → paste into index.html as SCRIPT_URL
//
// ============================================================================

// 📁 Google Drive folder for uploaded PDFs
const DRIVE_FOLDER_ID = 'YOUR_DRIVE_FOLDER_ID_HERE';

// Tab names inside THIS spreadsheet
const AUTH_SHEET_NAME       = 'Auth';        // Tab with registration data
const SUBMISSION_SHEET_NAME = 'Submission';  // Tab for storing submissions

// Column indices (1-based) inside the "Auth" tab
// Mapped to your columns: D=LEADER'S EMAIL, Y=TEAM SECRET CODE, Z=TEAM ID
const COL_TEAM_ID  = 26;  // Column Z — TEAM ID
const COL_EMAIL    = 4;   // Column D — LEADER'S EMAIL
const COL_SECRET   = 25;  // Column Y — TEAM SECRET CODE

// ============================================================================
// 🌐  doGet — Status check
// ============================================================================
function doGet() {
  return HtmlService.createHtmlOutput(
    '<h2 style="font-family:sans-serif;color:#333;">✅ NextGen Buildathon API is running.</h2>'
  ).setTitle('NextGen Buildathon API');
}

// ============================================================================
// 📨  doPost — Verifies credentials from "Auth" tab, stores in "Submission" tab
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

  const { teamName, teamId, email, domain, secretCode, fileName, fileBase64 } = payload;

  if (!teamName || !teamId || !email || !domain || !secretCode || !fileName || !fileBase64) {
    return _jsonResponse({ success: false, message: 'Error: All fields are required.' });
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
    // STEP 1 — Anti-Duplication Check ("Submission" tab)
    // ----------------------------------------------------------------
    const subSheet = ss.getSheetByName(SUBMISSION_SHEET_NAME);
    if (!subSheet) {
      return _jsonResponse({ success: false, message: 'Error: Submission tab not found.' });
    }

    const subData = subSheet.getDataRange().getValues();
    for (let i = 1; i < subData.length; i++) {
      if (String(subData[i][2]).trim().toLowerCase() === String(teamId).trim().toLowerCase()) {
        return _jsonResponse({
          success: false,
          message: 'Error: Your team has already submitted an idea.'
        });
      }
    }

    // ----------------------------------------------------------------
    // STEP 2 — Credential Verification ("Auth" tab — READ ONLY)
    // ----------------------------------------------------------------
    const authSheet = ss.getSheetByName(AUTH_SHEET_NAME);
    if (!authSheet) {
      return _jsonResponse({ success: false, message: 'Error: Auth tab not found.' });
    }

    const authData = authSheet.getDataRange().getValues();
    let verified   = false;

    for (let i = 1; i < authData.length; i++) {
      const rowTeamId = String(authData[i][COL_TEAM_ID - 1]).trim().toLowerCase();

      if (rowTeamId === String(teamId).trim().toLowerCase()) {
        const rowEmail  = String(authData[i][COL_EMAIL - 1]).trim().toLowerCase();
        const rowSecret = String(authData[i][COL_SECRET - 1]).trim();

        if (
          rowEmail  === String(email).trim().toLowerCase() &&
          rowSecret === String(secretCode).trim()
        ) {
          verified = true;
        }
        break;
      }
    }

    if (!verified) {
      return _jsonResponse({
        success: false,
        message: 'Error: Invalid Team ID, Email, or Secret Code combination.'
      });
    }

    // ----------------------------------------------------------------
    // STEP 3 — PDF Upload to Google Drive
    // ----------------------------------------------------------------
    let fileUrl;
    try {
      const decodedBytes = Utilities.base64Decode(fileBase64);
      const blob = Utilities.newBlob(decodedBytes, 'application/pdf', fileName);

      const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      const file   = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = file.getUrl();
    } catch (uploadErr) {
      return _jsonResponse({
        success: false,
        message: 'Error: Failed to upload PDF — ' + uploadErr.message
      });
    }

    // ----------------------------------------------------------------
    // STEP 4 — Append to "Submission" tab
    // ----------------------------------------------------------------
    subSheet.appendRow([
      new Date(),           // A — Timestamp
      teamName,             // B — Team Name
      teamId,               // C — Team ID
      email,                // D — Leader Mail
      domain,               // E — Domain / Track
      fileUrl,              // F — PDF Drive URL
      1                     // G — Status (1 = verified & submitted)
    ]);

    // ----------------------------------------------------------------
    // STEP 5 — Return success
    // ----------------------------------------------------------------
    return _jsonResponse({
      success: true,
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
// 🛠  Helper — Build a JSON ContentService response
// ============================================================================
function _jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
