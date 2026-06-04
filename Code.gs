// ═══════════════════════════════════════════════════════════
// STATE OF DANCE — Photoshoot Registration Backend
// Google Apps Script — paste this into script.google.com
// ═══════════════════════════════════════════════════════════

// ── CONFIGURATION ──────────────────────────────────────────
// After creating your Google Sheet, paste its ID here.
// The ID is in the URL: docs.google.com/spreadsheets/d/SHEET_ID/edit
const SHEET_ID = "YOUR_SHEET_ID_HERE";
const SHEET_NAME = "Registrations";

// Google Drive folder ID where screenshots will be saved.
// Create a folder in Drive, open it, copy the ID from the URL.
const DRIVE_FOLDER_ID = "YOUR_DRIVE_FOLDER_ID_HERE";

// ── MAIN HANDLER ───────────────────────────────────────────
function doPost(e) {
  // Allow requests from any origin (required for your GitHub page)
  const headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST",
    "Access-Control-Allow-Headers": "Content-Type",
    "Content-Type": "application/json"
  };

  try {
    const data = JSON.parse(e.postData.contents);

    // ── Save screenshot to Drive ──────────────────────────
    let screenshotUrl = "No screenshot";
    if (data.screenshotData && data.screenshotData.startsWith("data:image")) {
      try {
        const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
        // Decode base64
        const base64 = data.screenshotData.split(",")[1];
        const mimeMatch = data.screenshotData.match(/data:(image\/\w+);/);
        const mime = mimeMatch ? mimeMatch[1] : "image/jpeg";
        const ext = mime.split("/")[1];
        const blob = Utilities.newBlob(
          Utilities.base64Decode(base64), mime,
          `${data.studentName}_${Date.now()}.${ext}`
        );
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        screenshotUrl = file.getUrl();
      } catch (imgErr) {
        screenshotUrl = "Screenshot upload failed: " + imgErr.message;
      }
    }

    // ── Write to Sheet ────────────────────────────────────
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);

    // Create sheet with headers if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        "Timestamp", "Student Name", "Parent Name", "Branch",
        "Class", "Contact", "Payment Method", "Status",
        "Screenshot", "Amount"
      ]);
      // Style header row
      const headerRange = sheet.getRange(1, 1, 1, 10);
      headerRange.setBackground("#270934");
      headerRange.setFontColor("#deb060");
      headerRange.setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      new Date().toLocaleString("en-PH"),
      data.studentName || "",
      data.parentName  || "",
      data.branch      || "",
      data.className   || "",
      data.contact     || "",
      (data.paymentMethod || "").toUpperCase(),
      "Pending",
      screenshotUrl,
      "₱750"
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── CORS preflight handler ──────────────────────────────────
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "SOD Backend running" }))
    .setMimeType(ContentService.MimeType.JSON);
}
