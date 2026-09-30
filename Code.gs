// ============================================================================
// ⚙️ SINGLE Code.gs — Auth + Submission in ONE spreadsheet
// ============================================================================
//
// 📌 SETUP:
//    1. Open your Google Sheet: https://docs.google.com/spreadsheets/d/19gOqIXlAXiyE6c0GnhbkkB7Tv1jHlFuw5wmxkI4DVrs/edit
//    2. Click: Extensions → Apps Script
//    3. Paste this entire code (replacing everything).
//    4. (Optional) If you have a specific Drive Folder ID, put it in DRIVE_FOLDER_ID below.
//       If you leave it blank (''), the script will AUTOMATICALLY create a folder
//       named "NextGen_Buildathon_Submissions" in your Google Drive!
//    5. Click Deploy → Manage deployments → Edit (pencil) → Version: "New version" → Deploy.
//
// ============================================================================

// 📁 Optional Google Drive folder ID. Leave as '' to auto-create "NextGen_Buildathon_Submissions"
const DRIVE_FOLDER_ID = '';

// Tab names inside THIS spreadsheet
const AUTH_SHEET_NAME       = 'Auth';        // Tab with registration data
const SUBMISSION_SHEET_NAME = 'Submission';  // Tab for storing submissions

// Column indices (1-based) inside the "Auth" tab
// Mapped to: D = LEADER'S EMAIL (col 4), Y = TEAM SECRET CODE (col 25), Z = TEAM ID (col 26)
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
    // STEP 1 — Prepare "Submission" tab & Anti-Duplication Check
    // ----------------------------------------------------------------
    let subSheet = ss.getSheetByName(SUBMISSION_SHEET_NAME);
    if (!subSheet) {
      // Auto-create Submission tab with headers if not already created
      subSheet = ss.insertSheet(SUBMISSION_SHEET_NAME);
      subSheet.appendRow([
        'Timestamp',
        'Team Name',
        'Team ID',
        'Leader Mail',
        'Domain / Track',
        'PDF Drive URL',
        'Status'
      ]);
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
    let authSheet = ss.getSheetByName(AUTH_SHEET_NAME);
    if (!authSheet) {
      // Fallback: check if "Sheet1" is present, or use the first sheet
      authSheet = ss.getSheetByName('Sheet1') || ss.getSheets()[0];
    }

    if (!authSheet) {
      return _jsonResponse({ success: false, message: 'Error: Registration Auth sheet tab not found.' });
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
    // STEP 3 — PDF Upload to Google Drive (Auto-Folder Creation)
    // ----------------------------------------------------------------
    let fileUrl;
    try {
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

      // If no valid folder ID provided, auto-create or use "NextGen_Buildathon_Submissions" folder
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
