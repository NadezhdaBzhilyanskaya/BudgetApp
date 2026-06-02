// Budget Tracker — Google Apps Script Backend
// ─────────────────────────────────────────────
// Deploy as a Web App:
//   Execute as: Me
//   Who has access: Anyone
// Then copy the /exec URL into your budget tracker settings.

const SPREAD_SHEET_NAME = 'Budget';
const SHEET_NAME = 'BudgetData';

function getDataSheet() {
  let ss = SpreadsheetApp.openById("1YdfkCdxtp3K_NMfbLsUjDFgL03pv9S9Y8jKXA96yRSA");
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange('A1').setValue(JSON.stringify({ members: [], transactions: [] }));
  }
  return sheet;
}

function readData() {
  const raw = getDataSheet().getRange('A1').getValue();
  try {
    return raw ? JSON.parse(raw) : { members: [], transactions: [] };
  } catch (e) {
    return { members: [], transactions: [] };
  }
}

function writeData(data) {
  getDataSheet().getRange('A1').setValue(JSON.stringify(data));
}

// ── GET: fetch all data ──────────────────────
function doGet(e) {
  const data = readData();
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, data }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── POST: save all data ──────────────────────
// Body: { members: [...], transactions: [...] }
// Either field can be omitted to leave it unchanged.
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const current = readData();
    const updated = {
      members:      payload.members      !== undefined ? payload.members      : current.members,
      transactions: payload.transactions !== undefined ? payload.transactions : current.transactions,
    };
    writeData(updated);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
