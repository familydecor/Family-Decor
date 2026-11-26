/**
 * FAMILY DECOR - Reviews backend v2 (Google Apps Script)
 * 1) Naya Google Sheet banao -> Extensions -> Apps Script -> yeh poora code paste karo.
 * 2) Neeche ADMIN_KEY me apna secret password likho (yahi site pe "Admin" login me chalega).
 * 3) Deploy -> New deployment -> Web app -> Execute as: Me, Who has access: Anyone -> Deploy -> URL copy.
 * 4) Wo URL index.html me  const REVIEWS_API = '...';  me paste karo.
 * Code badalne ke baad: Deploy -> Manage deployments -> Edit -> New version.
 */
const SHEET_NAME = 'ReviewsV2';
const ADMIN_KEY = 'CHANGE_THIS_PASSWORD';
const HEADERS = ['ID','Date','Name','Email','Event','Rating','Review','Show (YES/NO)','TokenHash'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); sh.appendRow(HEADERS); sh.setFrozenRows(1); }
  return sh;
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function safe_(s) { s = String(s == null ? '' : s); return /^[=+\-@]/.test(s) ? "'" + s : s; }
function hash_(t) {
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(t));
  return Utilities.base64Encode(b);
}
function mask_(e) {
  e = String(e || ''); const i = e.indexOf('@'); if (i < 1) return '';
  return e.slice(0, Math.min(2, i)) + '***' + e.slice(i);
}
function rows_() {
  const sh = sheet_(); const n = sh.getLastRow(); if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, HEADERS.length).getValues().map((r, i) => ({
    row: i + 2, id: String(r[0]), date: r[1], name: String(r[2]), email: String(r[3]), event: String(r[4]),
    rating: Number(r[5]), text: String(r[6]), show: String(r[7]).toUpperCase(), th: String(r[8])
  }));
}
function doGet(e) {
  const p = (e && e.parameter) || {};
  const all = rows_().filter(r => r.show !== 'NO').reverse();
  if (p.action === 'admin') {
    if (p.key !== ADMIN_KEY) return out_({ error: 'unauthorized' });
    return out_(all.map(r => ({ id: r.id, name: r.name, email: r.email, event: r.event, rating: r.rating, text: r.text, date: r.date })));
  }
  return out_(all.map(r => ({ id: r.id, name: r.name, emailMasked: mask_(r.email), event: r.event, rating: r.rating, text: r.text, date: r.date })));
}
function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    let b = {}; try { b = JSON.parse(e.postData.contents); } catch (x) { return out_({ ok: false, error: 'bad' }); }
    const sh = sheet_();
    if (b.action === 'add') {
      if (b.website) return out_({ ok: true });
      const name = String(b.name || '').trim().slice(0, 60), email = String(b.email || '').trim().toLowerCase().slice(0, 100);
      const text = String(b.text || '').trim().slice(0, 600), rating = parseInt(b.rating), token = String(b.token || '');
      if (!name || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email) || text.length < 10 || !(rating >= 1 && rating <= 5) || token.length < 16)
        return out_({ ok: false, error: 'invalid' });
      if (rows_().filter(r => r.email.toLowerCase() === email).length >= 3) return out_({ ok: false, error: 'limit' });
      const id = Utilities.getUuid();
      sh.appendRow([id, new Date(), safe_(name), safe_(email), safe_(String(b.event || '').slice(0, 40)), rating, safe_(text), 'YES', hash_(token)]);
      return out_({ ok: true, id: id });
    }
    if (b.action === 'delete') {
      const r = rows_().find(x => x.id === String(b.id));
      if (r && r.th && r.th === hash_(String(b.token || ''))) { sh.deleteRow(r.row); return out_({ ok: true }); }
      return out_({ ok: false, error: 'forbidden' });
    }
    if (b.action === 'adminDelete') {
      if (b.key !== ADMIN_KEY) return out_({ ok: false, error: 'unauthorized' });
      const r = rows_().find(x => x.id === String(b.id));
      if (r) { sh.deleteRow(r.row); return out_({ ok: true }); }
      return out_({ ok: false, error: 'notfound' });
    }
    return out_({ ok: false, error: 'bad' });
  } finally { lock.releaseLock(); }
}
