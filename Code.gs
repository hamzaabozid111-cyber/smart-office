/**
 * SMART OFFICE: مخزن البيانات داخل Google Sheets
 * الصق هذا الكود في Extensions ← Apps Script داخل الشيت، ثم Deploy ← New deployment ← Web app
 * (Execute as: Me ، Who has access: Anyone). البيانات المحفوظة مشفّرة من التطبيق ولا يقرأها السكربت.
 */
var SHEET = 'data';
var CHUNK = 45000; // حد الخلية الواحدة 50 ألف حرف

function sh_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
}
function find_(sh, id) {
  var n = sh.getLastRow();
  if (n < 1) return 0;
  var ids = sh.getRange(1, 1, n, 1).getValues();
  for (var i = 0; i < n; i++) if (String(ids[i][0]) === id) return i + 1;
  return 0;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function doGet(e) {
  try {
    var id = String((e.parameter || {}).id || '');
    if (!/^[a-f0-9]{32}$/.test(id)) return out_({ error: 'bad id' });
    var sh = sh_(), row = find_(sh, id);
    if (!row) return out_({ ts: null });
    var head = sh.getRange(row, 1, 1, 3).getValues()[0];
    if (e.parameter.a === 'ts') return out_({ ts: Number(head[1]) });
    var cnt = Number(head[2]);
    var parts = cnt ? sh.getRange(row, 4, 1, cnt).getValues()[0] : [];
    return out_({ ts: Number(head[1]), d: parts.join('') });
  } catch (err) { return out_({ error: String(err) }); }
}
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var b = JSON.parse(e.postData.contents);
    if (!/^[a-f0-9]{32}$/.test(String(b.id)) || !b.d) return out_({ error: 'bad request' });
    var sh = sh_(), row = find_(sh, b.id);
    if (!row) row = Math.max(sh.getLastRow(), 0) + 1;
    var parts = [];
    for (var i = 0; i < b.d.length; i += CHUNK) parts.push(b.d.substr(i, CHUNK));
    var wide = Math.max(3 + parts.length, sh.getMaxColumns());
    if (sh.getMaxColumns() < 3 + parts.length) sh.insertColumnsAfter(sh.getMaxColumns(), 3 + parts.length - sh.getMaxColumns());
    sh.getRange(row, 1, 1, sh.getMaxColumns()).clearContent();
    sh.getRange(row, 1, 1, 3 + parts.length).setValues([[b.id, b.ts, parts.length].concat(parts)]);
    return out_({ ok: true });
  } catch (err) { return out_({ error: String(err) }); }
  finally { try { lock.releaseLock(); } catch (x) {} }
}
