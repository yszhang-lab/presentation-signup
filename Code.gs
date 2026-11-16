/**
 * Final presentation sign-up: backend for the public sign-up page.
 * Bound to the "Student Presentation Sign-Up Sheet" spreadsheet.
 *
 * GET  -> current sessions and slots as JSON
 * POST -> claim a slot:  {"session": 1, "slot": 3, "name": "Jane Doe"}
 *         cancel a slot: {"action": "cancel", "session": 1, "slot": 3, "code": "ABC123"}
 *
 * Each sign-up gets a cancel code. Codes are kept in this script's private
 * properties, not in the sheet, and are never sent out when the page loads.
 *
 * Layout it expects on the first sheet: a "SESSION ..." header in column B,
 * then rows with the slot number in column B and the student name in column C.
 */

var SLOT_COL = 2; // B
var NAME_COL = 3; // C
var MAX_NAME = 60;

function doGet() {
  return json_(readBoard_());
}

function doPost(e) {
  var req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad_request' });
  }

  if (req.action === 'cancel') return cancel_(req);

  var name = String(req.name || '').replace(/\s+/g, ' ').trim();
  if (!name || name.length > MAX_NAME) {
    return json_({ ok: false, error: 'bad_name' });
  }
  // Stop a typed name from being run as a spreadsheet formula.
  if (/^[=+\-@]/.test(name)) name = "'" + name;

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return json_({ ok: false, error: 'busy' });
  }

  try {
    var board = readBoard_();
    var session = null, slot = null;
    board.sessions.forEach(function (s) {
      if (s.id === Number(req.session)) session = s;
    });
    if (session) {
      session.slots.forEach(function (x) {
        if (x.n === Number(req.slot)) slot = x;
      });
    }
    if (!slot) return json_({ ok: false, error: 'no_such_slot', board: board });

    var wanted = name.replace(/^'/, '').toLowerCase();
    var already = null;
    board.sessions.forEach(function (s) {
      s.slots.forEach(function (x) {
        if (x.name && x.name.toLowerCase() === wanted) already = { session: s.id, slot: x.n };
      });
    });
    if (already) return json_({ ok: false, error: 'already_signed_up', at: already, board: board });
    if (slot.name) return json_({ ok: false, error: 'taken', board: board });

    var code = newCode_();
    sheet_().getRange(slot.row, NAME_COL).setValue(name);
    PropertiesService.getScriptProperties()
      .setProperty(codeKey_(session.id, slot.n), code + '|' + wanted);
    SpreadsheetApp.flush();
    return json_({ ok: true, code: code, board: readBoard_() });
  } finally {
    lock.releaseLock();
  }
}

function cancel_(req) {
  var code = String(req.code || '').replace(/\s+/g, '').toUpperCase();
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return json_({ ok: false, error: 'busy' });
  }
  try {
    var board = readBoard_();
    var slot = null;
    board.sessions.forEach(function (s) {
      if (s.id !== Number(req.session)) return;
      s.slots.forEach(function (x) {
        if (x.n === Number(req.slot)) slot = x;
      });
    });
    if (!slot || !slot.name) return json_({ ok: false, error: 'not_signed_up', board: board });

    var props = PropertiesService.getScriptProperties();
    var key = codeKey_(req.session, req.slot);
    var saved = String(props.getProperty(key) || '').split('|');
    // The code only counts for the name it was issued to, so a name typed
    // into the sheet by hand can only be removed in the sheet.
    if (saved.length < 2 || saved.slice(1).join('|') !== slot.name.toLowerCase()) {
      return json_({ ok: false, error: 'no_code', board: board });
    }
    if (!code || saved[0] !== code) return json_({ ok: false, error: 'wrong_code', board: board });

    sheet_().getRange(slot.row, NAME_COL).clearContent();
    props.deleteProperty(key);
    SpreadsheetApp.flush();
    return json_({ ok: true, board: readBoard_() });
  } finally {
    lock.releaseLock();
  }
}

function codeKey_(session, slot) {
  return 'code_' + Number(session) + '_' + Number(slot);
}

function newCode_() {
  var chars = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';
  var out = '';
  for (var i = 0; i < 6; i++) out += chars.charAt(Math.floor(Math.random() * chars.length));
  return out;
}

function sheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

function readBoard_() {
  var sh = sheet_();
  var last = sh.getLastRow();
  var values = sh.getRange(1, SLOT_COL, Math.max(last, 1), 2).getDisplayValues();
  var sessions = [];
  var current = null;
  for (var i = 0; i < values.length; i++) {
    var b = String(values[i][0]).trim();
    var c = String(values[i][1]).trim();
    if (/^session\b/i.test(b)) {
      current = { id: sessions.length + 1, title: b, slots: [] };
      sessions.push(current);
    } else if (current && /^\d+$/.test(b)) {
      current.slots.push({ n: Number(b), name: c, row: i + 1 });
    }
  }
  return { sessions: sessions };
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
