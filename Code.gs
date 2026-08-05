/**
 * Attendance backend — Google Apps Script
 *
 * Paste this into Extensions > Apps Script inside your attendance
 * Google Sheet, then deploy it as a Web app (see README).
 *
 * Sheets it expects (run setupSheets() once to create them):
 *   Students    Roll | Name | Section
 *   Courses     Code | Title | Section
 *   Attendance  written by this script
 */

var SHEET_STUDENTS = 'Students';
var SHEET_COURSES = 'Courses';
var SHEET_ATTENDANCE = 'Attendance';

var ATTENDANCE_HEADERS = [
  'Saved at', 'Date', 'Class', 'Course code', 'Course title',
  'Section', 'Roll', 'Name', 'Status'
];

var COL_STATUS = 9;   // column I

/* ================================================================== */
/* Reading                                                            */
/* ================================================================== */

function doGet(e) {
  try {
    var p = (e && e.parameter) || {};

    var bad = pinCheck(p.pin);
    if (bad) return json({ ok: false, code: 'PIN', error: bad });

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (p.action === 'history') {
      return json(readHistory(ss, p.course, p.section));
    }
    return json({
      ok: true,
      courses: readCourses(ss),
      students: readStudents(ss)
    });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function readStudents(ss) {
  var rows = readRows(ss, SHEET_STUDENTS);
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var roll = String(rows[i][0]).trim();
    if (!roll) continue;
    out.push({
      roll: roll,
      name: String(rows[i][1] || '').trim(),
      section: String(rows[i][2] || '').trim()
    });
  }
  return out;
}

function readCourses(ss) {
  var rows = readRows(ss, SHEET_COURSES);
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var code = String(rows[i][0]).trim();
    if (!code) continue;
    out.push({
      code: code,
      title: String(rows[i][1] || '').trim(),
      section: String(rows[i][2] || '').trim()
    });
  }
  return out;
}

function readRows(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Sheet "' + name + '" not found. Run setupSheets() once.');
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
}

/**
 * Everything recorded so far for one course + section.
 *
 * A "session" is one sitting: a date plus a class number, so two lectures of
 * the same course on the same day stay separate.
 *
 *   sessions  [{ key:'2026-01-12#1', date:'2026-01-12', cls:'1' }, ...]
 *   marks     { '210101': { '2026-01-12#1': 'P' }, ... }
 *   names     { '210101': 'Rakib Hasan', ... }
 */
function readHistory(ss, courseCode, section) {
  var sheet = ss.getSheetByName(SHEET_ATTENDANCE);
  if (!sheet || sheet.getLastRow() < 2) {
    return { ok: true, sessions: [], marks: {}, names: {} };
  }

  var values = sheet.getRange(2, 2, sheet.getLastRow() - 1, 8).getValues(); // Date..Status
  var short = { Present: 'P', Late: 'L', Absent: 'A' };
  var seen = {}, sessions = [], marks = {}, names = {};

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (String(row[2]).trim() !== courseCode) continue;
    if (String(row[4]).trim() !== section) continue;

    var date = asDateString(row[0]);
    var cls = normClass(row[1]);
    var roll = String(row[5]).trim();
    if (!date || !roll) continue;

    var key = date + '#' + cls;
    if (!seen[key]) { seen[key] = true; sessions.push({ key: key, date: date, cls: cls }); }
    names[roll] = String(row[6] || '').trim();
    if (!marks[roll]) marks[roll] = {};
    marks[roll][key] = short[String(row[7]).trim()] || 'A';
  }

  sessions.sort(function (a, b) {
    return a.date === b.date ? Number(a.cls) - Number(b.cls) : (a.date < b.date ? -1 : 1);
  });

  return { ok: true, sessions: sessions, marks: marks, names: names };
}

/** Rows written before the Class column existed count as class 1. */
function normClass(value) {
  var s = String(value == null ? '' : value).trim();
  return s === '' ? '1' : s;
}

function asDateString(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value).trim();
}

/* ================================================================== */
/* Writing — every action arrives here                                */
/* ================================================================== */

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var body = JSON.parse(e.postData.contents);

    var bad = pinCheck(body.pin);
    if (bad) return json({ ok: false, code: 'PIN', error: bad });

    lock.waitLock(20000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    switch (body.action || 'save') {
      case 'save':          return json(saveSession(ss, body));
      case 'deleteSession': return json(deleteSession(ss, body));
      case 'setMark':       return json(setMark(ss, body));
      case 'addStudent':    return json(addStudent(ss, body));
      case 'removeStudent': return json(removeStudent(ss, body));
      default: return json({ ok: false, error: 'Unknown action: ' + body.action });
    }
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignored) {}
  }
}

/* ---- save a whole sitting ---------------------------------------- */

function saveSession(ss, body) {
  var sheet = attendanceSheet(ss);
  var cls = normClass(body.cls);
  var existing = findSessionRows(sheet, body.date, cls, body.courseCode, body.section);

  if (existing.length && !body.overwrite) {
    return { ok: false, code: 'ALREADY_SAVED', count: existing.length };
  }
  deleteRows(sheet, existing);

  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var rows = body.records.map(function (r) {
    return [stamp, body.date, cls, body.courseCode, body.courseTitle,
            body.section, String(r.roll), r.name, r.status];
  });

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, ATTENDANCE_HEADERS.length)
         .setValues(rows);
  }
  return { ok: true, saved: rows.length, replaced: existing.length };
}

/* ---- delete a whole sitting -------------------------------------- */

function deleteSession(ss, body) {
  var sheet = attendanceSheet(ss);
  var rows = findSessionRows(sheet, body.date, normClass(body.cls), body.courseCode, body.section);
  if (!rows.length) return { ok: false, error: 'That class is no longer in the sheet.' };
  deleteRows(sheet, rows);
  return { ok: true, deleted: rows.length };
}

/* ---- change one student's mark ----------------------------------- */

function setMark(ss, body) {
  var sheet = attendanceSheet(ss);
  var cls = normClass(body.cls);
  var roll = String(body.roll).trim();

  if (sheet.getLastRow() >= 2) {
    var values = sheet.getRange(2, 2, sheet.getLastRow() - 1, 6).getValues(); // Date..Roll
    for (var i = 0; i < values.length; i++) {
      if (asDateString(values[i][0]) === body.date &&
          normClass(values[i][1]) === cls &&
          String(values[i][2]).trim() === body.courseCode &&
          String(values[i][4]).trim() === body.section &&
          String(values[i][5]).trim() === roll) {
        sheet.getRange(i + 2, COL_STATUS).setValue(body.status);
        return { ok: true, updated: 1 };
      }
    }
  }

  // No row yet — the student joined after this class was taken. Add one.
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  sheet.appendRow([stamp, body.date, cls, body.courseCode, body.courseTitle,
                   body.section, roll, body.name || '', body.status]);
  return { ok: true, added: 1 };
}

/* ---- the roll ----------------------------------------------------- */

function addStudent(ss, body) {
  var sheet = ss.getSheetByName(SHEET_STUDENTS);
  if (!sheet) throw new Error('Sheet "Students" not found. Run setupSheets() once.');

  var roll = String(body.roll || '').trim();
  var section = String(body.section || '').trim();
  if (!roll) return { ok: false, error: 'A roll number is required.' };

  var rows = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues() : [];
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === roll && String(rows[i][2]).trim() === section) {
      return { ok: false, error: 'Roll ' + roll + ' is already on the ' + section + ' roll.' };
    }
  }
  sheet.appendRow([roll, String(body.name || '').trim(), section]);
  return { ok: true, added: 1 };
}

function removeStudent(ss, body) {
  var sheet = ss.getSheetByName(SHEET_STUDENTS);
  if (!sheet) throw new Error('Sheet "Students" not found. Run setupSheets() once.');

  var roll = String(body.roll || '').trim();
  var section = String(body.section || '').trim();
  var removed = 0, purged = 0;

  if (sheet.getLastRow() > 1) {
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    var hits = [];
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === roll && String(rows[i][2]).trim() === section) {
        hits.push(i + 2);
      }
    }
    deleteRows(sheet, hits);
    removed = hits.length;
  }

  if (body.purge) purged = purgeRecords(ss, roll, section);
  return { ok: true, removed: removed, purged: purged };
}

/** Delete every attendance row for one student in one section. */
function purgeRecords(ss, roll, section) {
  var sheet = attendanceSheet(ss);
  if (sheet.getLastRow() < 2) return 0;
  var values = sheet.getRange(2, 6, sheet.getLastRow() - 1, 2).getValues(); // Section, Roll
  var hits = [];
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === section && String(values[i][1]).trim() === roll) {
      hits.push(i + 2);
    }
  }
  deleteRows(sheet, hits);
  return hits.length;
}

/* ================================================================== */
/* Shared helpers                                                     */
/* ================================================================== */

function attendanceSheet(ss) {
  var sheet = ss.getSheetByName(SHEET_ATTENDANCE);
  if (!sheet) throw new Error('Sheet "Attendance" not found. Run setupSheets() once.');
  if (String(sheet.getRange(1, 3).getValue()).trim() !== 'Class') {
    throw new Error('The Attendance sheet is missing the Class column. Run upgradeSheet() once.');
  }
  return sheet;
}

/** Row numbers holding this date + class + course + section. */
function findSessionRows(sheet, date, cls, courseCode, section) {
  if (sheet.getLastRow() < 2) return [];
  var values = sheet.getRange(2, 2, sheet.getLastRow() - 1, 5).getValues(); // Date..Section
  var hits = [];
  for (var i = 0; i < values.length; i++) {
    if (asDateString(values[i][0]) === date &&
        normClass(values[i][1]) === cls &&
        String(values[i][2]).trim() === courseCode &&
        String(values[i][4]).trim() === section) {
      hits.push(i + 2);
    }
  }
  return hits;
}

/** Delete rows by number, bottom-up, in contiguous blocks so it stays fast. */
function deleteRows(sheet, rowNumbers) {
  if (!rowNumbers.length) return;
  var sorted = rowNumbers.slice().sort(function (a, b) { return a - b; });
  var end = sorted.length - 1;
  while (end >= 0) {
    var start = end;
    while (start > 0 && sorted[start - 1] === sorted[start] - 1) start--;
    sheet.deleteRows(sorted[start], end - start + 1);
    end = start - 1;
  }
}

/* ================================================================== */
/* PIN                                                                */
/* ================================================================== */

/**
 * Returns null when the PIN is good, otherwise the reason it is not.
 *
 * This fails CLOSED: with no PIN stored, nothing is served at all. An earlier
 * version treated "no PIN configured" as "open to everyone", which meant a
 * forgotten setup step silently published the whole sheet.
 */
function pinCheck(pin) {
  var stored = PropertiesService.getScriptProperties().getProperty('PIN');
  if (!stored) return 'No PIN is set on this script. Run setPin() in the Apps Script editor.';
  if (String(pin || '') !== stored) return 'That PIN is not correct.';
  return null;
}

/** Edit the value, run this once, then change the value back. */
function setPin() {
  PropertiesService.getScriptProperties().setProperty('PIN', '1234');
}

/** Run this to confirm a PIN really is stored. Check the execution log. */
function checkPinIsSet() {
  var stored = PropertiesService.getScriptProperties().getProperty('PIN');
  Logger.log(stored ? 'PIN is set (' + stored.length + ' characters).' : 'NO PIN IS SET.');
}

/* ================================================================== */
/* Setup and upgrade                                                  */
/* ================================================================== */

function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensure(ss, SHEET_STUDENTS, ['Roll', 'Name', 'Section']);
  ensure(ss, SHEET_COURSES, ['Code', 'Title', 'Section']);
  ensure(ss, SHEET_ATTENDANCE, ATTENDANCE_HEADERS);
  upgradeSheet();
}

function ensure(ss, name, headers) {
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
         .setFontWeight('bold').setBackground('#eef1f6');
    sheet.setFrozenRows(1);
  }
}

/**
 * Safe to run more than once. Adds the Class column to an Attendance sheet
 * created before multiple classes per day were supported, and marks every
 * existing row as class 1.
 */
function upgradeSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ATTENDANCE);
  if (!sheet) throw new Error('Sheet "Attendance" not found. Run setupSheets() first.');
  if (String(sheet.getRange(1, 3).getValue()).trim() === 'Class') return 'Already up to date.';

  sheet.insertColumnBefore(3);
  sheet.getRange(1, 3).setValue('Class').setFontWeight('bold').setBackground('#eef1f6');

  var rows = sheet.getLastRow() - 1;
  if (rows > 0) {
    var fill = [];
    for (var i = 0; i < rows; i++) fill.push(['1']);
    sheet.getRange(2, 3, rows, 1).setValues(fill);
  }
  return 'Added the Class column to ' + rows + ' rows.';
}

/* ================================================================== */

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
