// ============================================================================
// ⚙️  CONFIGURATION — Paste your own IDs here
// ============================================================================
//
// 📌 HOW TO USE:
//    1. Open your Google Sheet → Extensions → Apps Script
//    2. Delete any default code in Code.gs → Paste this entire file
//    3. Click "+" next to Files → HTML → Name it "Index" → Paste Index.html
//    4. Fill in DRIVE_FOLDER_ID below
//    5. Deploy → New Deployment → Web App → Execute as "Me" → Anyone
//    6. Copy the Web App URL — that's your live form!
//
// If pasting inside the spreadsheet's own script editor,
// leave SPREADSHEET_ID as '' (empty) — it auto-detects.
// If using a standalone script project, paste the full Spreadsheet ID.
// ============================================================================
const SPREADSHEET_ID        = '';                                  // Leave empty if bound to spreadsheet
const REGISTRATION_SHEET    = 'Registrations';                     // Sheet A name (tab name)
const SUBMISSION_SHEET      = 'Submissions';                       // Sheet B name (tab name)
const DRIVE_FOLDER_ID       = 'YOUR_DRIVE_FOLDER_ID_HERE';         // Google Drive folder for PDFs

// Column indices (1-based) inside the REGISTRATION sheet
// Adjust these if your columns are in a different order
const REG_COL_TEAM_ID       = 1;   // Column A — Team ID
const REG_COL_EMAIL         = 2;   // Column B — Registered Team Leader Email
const REG_COL_SECRET        = 3;   // Column C — Team Secret Code

// ============================================================================
// 🌐  doGet — Simple status page (UI is hosted on Vercel, not here)
// ============================================================================
function doGet() {
  return HtmlService.createHtmlOutput(
    '<h2 style="font-family:sans-serif;color:#333;">✅ NextGen Buildathon API is running.</h2>' +
    '<p style="font-family:sans-serif;color:#666;">The submission form UI is hosted separately on Vercel.</p>'
  ).setTitle('NextGen Buildathon API');
}

// ============================================================================
// 📨  doPost — Handles form submissions with full verification pipeline
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

  // Basic server-side presence check
  if (!teamName || !teamId || !email || !domain || !secretCode || !fileName || !fileBase64) {
    return _jsonResponse({ success: false, message: 'Error: All fields are required.' });
  }

  // ------------------------------------------------------------------
  // 1. Acquire a script-level lock to prevent race conditions
  //    (two identical submissions at the same millisecond)
  // ------------------------------------------------------------------
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // wait up to 30 seconds for the lock
  } catch (err) {
    return _jsonResponse({
      success: false,
      message: 'Error: Server is busy processing another submission. Please try again in a few seconds.'
    });
  }

  try {
    // Auto-detect: if bound to a spreadsheet, use getActive; otherwise openById
    const ss = SPREADSHEET_ID
      ? SpreadsheetApp.openById(SPREADSHEET_ID)
      : SpreadsheetApp.getActiveSpreadsheet();

    // ----------------------------------------------------------------
    // STEP 3.1 — Anti-Duplication Check
    // Query the Submissions sheet; reject if Team ID already exists
    // ----------------------------------------------------------------
    const subSheet  = ss.getSheetByName(SUBMISSION_SHEET);
    if (!subSheet) {
      return _jsonResponse({ success: false, message: 'Error: Submissions sheet not found. Contact admin.' });
    }

    const subData   = subSheet.getDataRange().getValues();
    // Column C (index 2) in Submissions = Team ID
    for (let i = 1; i < subData.length; i++) {            // skip header row
      if (String(subData[i][2]).trim().toLowerCase() === String(teamId).trim().toLowerCase()) {
        return _jsonResponse({
          success: false,
          message: 'Error: Your team has already submitted an idea.'
        });
      }
    }

    // ----------------------------------------------------------------
    // STEP 3.2 — Credential Verification against Registrations sheet
    // ----------------------------------------------------------------
    const regSheet = ss.getSheetByName(REGISTRATION_SHEET);
    if (!regSheet) {
      return _jsonResponse({ success: false, message: 'Error: Registrations sheet not found. Contact admin.' });
    }

    const regData  = regSheet.getDataRange().getValues();
    let verified   = false;

    for (let i = 1; i < regData.length; i++) {             // skip header row
      const rowTeamId = String(regData[i][REG_COL_TEAM_ID - 1]).trim().toLowerCase();
      if (rowTeamId === String(teamId).trim().toLowerCase()) {
        // Team ID found — now verify email AND secret code
        const rowEmail  = String(regData[i][REG_COL_EMAIL - 1]).trim().toLowerCase();
        const rowSecret = String(regData[i][REG_COL_SECRET - 1]).trim();

        if (
          rowEmail  === String(email).trim().toLowerCase() &&
          rowSecret === String(secretCode).trim()
        ) {
          verified = true;
        }
        break; // Team ID is unique; stop after first match
      }
    }

    if (!verified) {
      return _jsonResponse({
        success: false,
        message: 'Error: Invalid Team ID, Email, or Secret Code combination.'
      });
    }

    // ----------------------------------------------------------------
    // STEP 3.3 — PDF Upload to Google Drive
    // Decode base64 → Blob → Save to designated Drive folder
    // ----------------------------------------------------------------
    let fileUrl;
    try {
      const decodedBytes = Utilities.base64Decode(fileBase64);
      const blob = Utilities.newBlob(decodedBytes, 'application/pdf', fileName);

      const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      const file   = folder.createFile(blob);

      // Make the file viewable via link so judges / admins can access it
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = file.getUrl();
    } catch (uploadErr) {
      return _jsonResponse({
        success: false,
        message: 'Error: Failed to upload PDF — ' + uploadErr.message
      });
    }

    // ----------------------------------------------------------------
    // STEP 3.4 — Append verified submission to the Submissions sheet
    // Columns: [Timestamp, Team Name, Team ID, Leader Mail, Domain,
    //            PDF Drive URL, Status]
    // ----------------------------------------------------------------
    const timestamp = new Date();
    subSheet.appendRow([
      timestamp,          // A — Timestamp
      teamName,           // B — Team Name
      teamId,             // C — Team ID
      email,              // D — Leader Mail
      domain,             // E — Domain / Track
      fileUrl,            // F — PDF Drive URL
      1                   // G — Status (1 = verified & submitted)
    ]);

    // ----------------------------------------------------------------
    // STEP 3.5 — Return success
    // ----------------------------------------------------------------
    return _jsonResponse({
      success: true,
      message: 'Success! Your idea has been submitted and verified. 🎉'
    });

  } catch (fatalErr) {
    return _jsonResponse({
      success: false,
      message: 'Error: An unexpected server error occurred — ' + fatalErr.message
    });
  } finally {
    // Always release the lock
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
