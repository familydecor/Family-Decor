// Family Decor - Reviews backend (Google Apps Script)
// Ye script Google Sheet me reviews save karti hai aur website ko list wapas deti hai.
const SHEET_NAME = 'Reviews';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); sh.appendRow(['Date', 'Name', 'Event', 'Rating', 'Review', 'Show (YES/NO)']); }
  return sh;
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function safe_(t) { t = String(t || '').trim(); return /^[=+\-@]/.test(t) ? "'" + t : t; }

// Website review list maangti hai
function doGet(e) {
  const rows = getSheet_().getDataRange().getValues(); rows.shift();
  const list = rows
    .filter(r => r[1] && r[4] && String(r[5]).trim().toUpperCase() !== 'NO')
    .map(r => ({ name: String(r[1]), event: String(r[2] || ''), rating: Number(r[3]) || 5, text: String(r[4]) }))
    .reverse().slice(0, 100);
  return out_(list);
}

// Website naya review bhejti hai
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const name = String(d.name || '').trim().slice(0, 60);
    const text = String(d.text || '').trim().slice(0, 600);
    const event = String(d.event || '').trim().slice(0, 40);
    const rating = Math.max(1, Math.min(5, parseInt(d.rating) || 0));
    if (!name || text.length < 10 || !rating) return out_({ ok: false });
    const lock = LockService.getScriptLock(); lock.waitLock(10000);
    try { getSheet_().appendRow([new Date(), safe_(name), safe_(event), rating, safe_(text), 'YES']); }
    finally { lock.releaseLock(); }
    return out_({ ok: true });
  } catch (err) { return out_({ ok: false }); }
}
