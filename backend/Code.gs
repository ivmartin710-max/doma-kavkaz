// Голосование «Дома на Кавказе». Хранит голоса на листе «votes».
var MAX_VOTERS = 8, MAX_PICKS = 2;
var KEYS = ['panda','berloga350','berloga220','berloga400','alas','chalet','narat','cheget','elbrus','djappu','flora'];

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('votes');
  if (!sh) { sh = ss.insertSheet('votes'); sh.appendRow(['name', 'picks', 'token', 'updated']); }
  return sh;
}
function norm_(s) { return String(s || '').trim().replace(/\s+/g, ' ').toLowerCase(); }
function list_(sh) {
  var rows = sh.getDataRange().getValues().slice(1), out = [];
  rows.forEach(function (r) { if (r[0] && r[1]) out.push({ name: String(r[0]), picks: String(r[1]) }); });
  return out;
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === 'vote') {
    return vote_({ name: p.name, token: p.token, picks: String(p.picks || '').split(',').filter(String) });
  }
  return json_({ votes: list_(sheet_()) });
}

function doPost(e) { return vote_(JSON.parse((e && e.postData && e.postData.contents) || '{}')); }

function vote_(body) {
  var lock = null;
  try { lock = LockService.getDocumentLock(); if (!lock.tryLock(3000)) lock = null; } catch (x) { lock = null; }
  try {
    var sh = sheet_();
    var name = String(body.name || '').trim().slice(0, 30), token = String(body.token || '').slice(0, 64);
    if (!name) return json_({ error: 'no_name', votes: list_(sh) });
    var picks = (body.picks || []).map(String).filter(function (k, i, a) { return KEYS.indexOf(k) > -1 && a.indexOf(k) === i; });
    if (picks.length > MAX_PICKS) return json_({ error: 'too_many', votes: list_(sh) });
    var data = sh.getDataRange().getValues(), row = -1, voters = 0;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][1]) voters++;
      if (norm_(data[i][0]) === norm_(name)) row = i + 1;
    }
    if (row > 0) {
      var saved = String(sh.getRange(row, 3).getValue());
      if (saved && saved !== token) return json_({ error: 'taken', votes: list_(sh) });
      if (!picks.length) sh.deleteRow(row);
      else sh.getRange(row, 1, 1, 4).setValues([[name, picks.join(','), token, new Date()]]);
    } else if (picks.length) {
      if (voters >= MAX_VOTERS) return json_({ error: 'full', votes: list_(sh) });
      sh.appendRow([name, picks.join(','), token, new Date()]);
    }
    return json_({ ok: true, votes: list_(sh) });
  } catch (err) {
    return json_({ error: 'server', detail: String(err) });
  } finally { if (lock) { try { lock.releaseLock(); } catch (x) {} } }
}
