/**
 * Optional Google Sheets ticket sync endpoint for Kolli Builders.
 * Setup: create a Google Sheet, open Extensions > Apps Script, paste this file,
 * set SPREADSHEET_ID below, then Deploy > New deployment > Web app.
 * Execute as: Me. Access: choose the narrowest access supported by your workflow.
 * Anyone with the deployed URL may be able to submit data; do not use public
 * access for sensitive customer data without adding authentication/validation.
 */
const SPREADSHEET_ID = "PASTE_GOOGLE_SHEET_ID_HERE";
const SHEET_NAME = "Support Tickets";
const HEADERS = ["id", "name", "contact", "category", "priority", "subject", "details", "status", "resolution", "createdAt", "updatedAt"];

function doPost(e) {
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const ticket = payload.ticket;
    if (payload.action !== "upsertTicket" || !ticket || !ticket.id) return json_({ ok: false, error: "Invalid ticket payload" });
    if (!SPREADSHEET_ID || SPREADSHEET_ID === "PASTE_GOOGLE_SHEET_ID_HERE") return json_({ ok: false, error: "Configure SPREADSHEET_ID first" });
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
    const idColumn = HEADERS.indexOf("id") + 1;
    const rows = sheet.getLastRow() > 1 ? sheet.getRange(2, idColumn, sheet.getLastRow() - 1, 1).getDisplayValues() : [];
    const rowIndex = rows.findIndex(row => row[0] === String(ticket.id));
    const values = HEADERS.map(key => String(ticket[key] == null ? "" : ticket[key]).slice(0, 10000));
    if (rowIndex >= 0) sheet.getRange(rowIndex + 2, 1, 1, HEADERS.length).setValues([values]);
    else sheet.appendRow(values);
    return json_({ ok: true, id: ticket.id });
  } catch (error) {
    return json_({ ok: false, error: String(error) });
  }
}
function doGet() { return json_({ ok: true, service: "Kolli Builders support ticket sync" }); }
function json_(data) { return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }
