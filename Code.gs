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
var SHEET_COMPONENTS = 'Components';
var SHEET_MARKS = 'Marks';

var ATTENDANCE_HEADERS = [
  'Saved at', 'Date', 'Class', 'Course code', 'Course title',
  'Section', 'Roll', 'Name', 'Status'
];

var COMPONENT_HEADERS = ['Course code', 'Section', 'Component', 'Marks'];
var MARKS_HEADERS = ['Saved at', 'Date', 'Course code', 'Section',
                     'Component', 'Item', 'Max', 'Roll', 'Name', 'Score'];

var COL_STATUS = 9;   // column I on Attendance
var COL_SCORE = 10;   // column J on Marks

/**
 * Suggested starting components, as MARKS ALLOTTED rather than percentages.
 * Editable from the site. The totals here follow the common 30 + 70 split for
 * theory and 60 + 40 for sessional; change them to match your department.
 */
var DEFAULT_COMPONENTS = {
  Sessional: [['Homework', 10], ['Lab task', 20], ['Lab report', 20], ['Class performance', 10]],
  Theory:    [['Homework', 5], ['Assignment', 5], ['Spot test', 5],
              ['Class test', 10], ['Class performance', 5]]
};

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
    if (p.action === 'marks') {
      return json(readMarks(ss, p.course, p.section));
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
  var rows = readRows(ss, SHEET_COURSES, 4);
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var code = String(rows[i][0]).trim();
    if (!code) continue;
    var type = String(rows[i][3] || '').trim();
    out.push({
      code: code,
      title: String(rows[i][1] || '').trim(),
      section: String(rows[i][2] || '').trim(),
      type: type === 'Sessional' ? 'Sessional' : 'Theory'
    });
  }
  return out;
}

function readRows(ss, name, width) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Sheet "' + name + '" not found. Run setupSheets() once.');
  if (sheet.getLastRow() < 2) return [];
  var want = width || 3;
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1,
                            Math.min(want, Math.max(1, sheet.getLastColumn()))).getValues();
  for (var i = 0; i < rows.length; i++) {          // pad short rows
    while (rows[i].length < want) rows[i].push('');
  }
  return rows;
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
  var short = { Present: 'P', Late: 'L', 'Early leave': 'E', Absent: 'A' };
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

/**
 * Components and marks for one course + section.
 *
 * An "item" is one piece of assessed work: a component plus a label, so
 * Homework 1 and Homework 2 stay separate inside the Homework component.
 *
 *   components  [{ name:'Homework', weight:15 }, ...]
 *   items       [{ key:'Homework#HW 1', component, item, date, max }, ...]
 *   scores      { '210101': { 'Homework#HW 1': 8 }, ... }
 */
function readMarks(ss, courseCode, section) {
  var components = [];
  var rows = readRows(ss, SHEET_COMPONENTS, 4);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() !== courseCode) continue;
    if (String(rows[i][1]).trim() !== section) continue;
    var name = String(rows[i][2]).trim();
    if (!name) continue;
    components.push({ name: name, marks: Number(rows[i][3]) || 0 });
  }

  var sheet = ss.getSheetByName(SHEET_MARKS);
  var seen = {}, items = [], scores = {}, names = {};

  if (sheet && sheet.getLastRow() >= 2) {
    var values = sheet.getRange(2, 2, sheet.getLastRow() - 1, 9).getValues(); // Date..Score
    for (var j = 0; j < values.length; j++) {
      var r = values[j];
      if (String(r[1]).trim() !== courseCode) continue;
      if (String(r[2]).trim() !== section) continue;

      var comp = String(r[3]).trim();
      var item = String(r[4]).trim();
      var roll = String(r[6]).trim();
      if (!comp || !item || !roll) continue;

      var key = comp + '#' + item;
      if (!seen[key]) {
        seen[key] = true;
        items.push({ key: key, component: comp, item: item,
                     date: asDateString(r[0]), max: Number(r[5]) || 0 });
      }
      names[roll] = String(r[7] || '').trim();
      if (!scores[roll]) scores[roll] = {};
      if (r[8] !== '' && r[8] !== null) scores[roll][key] = Number(r[8]);
    }
  }

  items.sort(function (a, b) {
    return a.date === b.date ? (a.key < b.key ? -1 : 1) : (a.date < b.date ? -1 : 1);
  });

  return { ok: true, components: components, items: items, scores: scores, names: names };
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
      case 'saveComponents':return json(saveComponents(ss, body));
      case 'saveItem':      return json(saveItem(ss, body));
      case 'setScore':      return json(setScore(ss, body));
      case 'deleteItem':    return json(deleteItem(ss, body));
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

/* ---- components and their weights ---------------------------------- */

/** Replaces the whole component list for one course + section. */
function saveComponents(ss, body) {
  var sheet = ss.getSheetByName(SHEET_COMPONENTS);
  if (!sheet) throw new Error('Sheet "Components" not found. Run setupSheets() once.');

  var hits = [];
  if (sheet.getLastRow() >= 2) {
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === body.courseCode &&
          String(rows[i][1]).trim() === body.section) hits.push(i + 2);
    }
  }
  deleteRows(sheet, hits);

  var out = (body.components || []).filter(function (c) { return String(c.name || '').trim(); })
    .map(function (c) {
      var allotted = c.marks === undefined ? c.weight : c.marks;   // old payloads said weight
      return [body.courseCode, body.section, String(c.name).trim(), Number(allotted) || 0];
    });

  if (out.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, out.length, COMPONENT_HEADERS.length).setValues(out);
  }
  return { ok: true, saved: out.length };
}

/* ---- marks for one item -------------------------------------------- */

/** Replaces every score for one component + item. Blank scores are not stored. */
function saveItem(ss, body) {
  var sheet = ss.getSheetByName(SHEET_MARKS);
  if (!sheet) throw new Error('Sheet "Marks" not found. Run setupSheets() once.');

  var comp = String(body.component || '').trim();
  var item = String(body.item || '').trim();
  if (!comp || !item) return { ok: false, error: 'A component and an item name are both needed.' };

  var hits = findItemRows(sheet, body.courseCode, body.section, comp, item);
  if (hits.length && !body.overwrite && body.isNew) {
    return { ok: false, code: 'ALREADY_SAVED', count: hits.length };
  }
  deleteRows(sheet, hits);

  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var max = Number(body.max) || 0;
  var out = (body.records || [])
    .filter(function (r) { return r.score !== '' && r.score !== null && r.score !== undefined; })
    .map(function (r) {
      return [stamp, body.date, body.courseCode, body.section, comp, item, max,
              String(r.roll), r.name || '', Number(r.score)];
    });

  if (out.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, out.length, MARKS_HEADERS.length).setValues(out);
  }
  return { ok: true, saved: out.length, replaced: hits.length };
}

/**
 * Change one student's score for one item, without touching anybody else.
 * A blank score deletes their row, which excuses them from that item rather
 * than scoring it zero.
 */
function setScore(ss, body) {
  var sheet = ss.getSheetByName(SHEET_MARKS);
  if (!sheet) throw new Error('Sheet "Marks" not found. Run setupSheets() once.');

  var comp = String(body.component || '').trim();
  var item = String(body.item || '').trim();
  var roll = String(body.roll || '').trim();
  if (!comp || !item || !roll) return { ok: false, error: 'Missing component, item or roll.' };

  var blank = body.score === '' || body.score === null || body.score === undefined;

  if (sheet.getLastRow() >= 2) {
    var values = sheet.getRange(2, 3, sheet.getLastRow() - 1, 6).getValues(); // Course..Roll
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][0]).trim() === body.courseCode &&
          String(values[i][1]).trim() === body.section &&
          String(values[i][2]).trim() === comp &&
          String(values[i][3]).trim() === item &&
          String(values[i][5]).trim() === roll) {
        if (blank) { sheet.deleteRow(i + 2); return { ok: true, cleared: 1 }; }
        sheet.getRange(i + 2, COL_SCORE).setValue(Number(body.score));
        return { ok: true, updated: 1 };
      }
    }
  }

  if (blank) return { ok: true, cleared: 0 };   // nothing there to clear

  // No row yet — this student was not marked when the item was saved.
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  sheet.appendRow([stamp, body.date, body.courseCode, body.section, comp, item,
                   Number(body.max) || 0, roll, body.name || '', Number(body.score)]);
  return { ok: true, added: 1 };
}

function deleteItem(ss, body) {
  var sheet = ss.getSheetByName(SHEET_MARKS);
  if (!sheet) throw new Error('Sheet "Marks" not found. Run setupSheets() once.');
  var hits = findItemRows(sheet, body.courseCode, body.section,
                          String(body.component).trim(), String(body.item).trim());
  if (!hits.length) return { ok: false, error: 'That item is no longer in the sheet.' };
  deleteRows(sheet, hits);
  return { ok: true, deleted: hits.length };
}

function findItemRows(sheet, courseCode, section, component, item) {
  if (sheet.getLastRow() < 2) return [];
  var values = sheet.getRange(2, 3, sheet.getLastRow() - 1, 4).getValues(); // Course..Item
  var hits = [];
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === courseCode &&
        String(values[i][1]).trim() === section &&
        String(values[i][2]).trim() === component &&
        String(values[i][3]).trim() === item) hits.push(i + 2);
  }
  return hits;
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

/** Delete every attendance row and every mark for one student in one section. */
function purgeRecords(ss, roll, section) {
  var gone = wipe(attendanceSheet(ss), 6, 7, section, roll);          // Section, Roll
  var marks = ss.getSheetByName(SHEET_MARKS);
  if (marks) gone += wipe(marks, 4, 8, section, roll);                // Section, Roll
  return gone;
}

function wipe(sheet, sectionCol, rollCol, section, roll) {
  if (sheet.getLastRow() < 2) return 0;
  var span = rollCol - sectionCol + 1;
  var values = sheet.getRange(2, sectionCol, sheet.getLastRow() - 1, span).getValues();
  var hits = [];
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === section &&
        String(values[i][span - 1]).trim() === roll) hits.push(i + 2);
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
var MAX_TRIES = 10;        // wrong PINs allowed before a pause
var LOCK_SECONDS = 300;    // how long that pause lasts

/**
 * Returns null when the PIN is good, otherwise the reason it is not.
 *
 * Fails CLOSED: with no PIN stored, nothing is served at all.
 *
 * Because the web app URL is necessarily public, the PIN is the only thing
 * standing between a stranger and the sheet. So repeated wrong answers pause
 * the script for everyone for a few minutes, which turns guessing from
 * "thousands of tries an hour" into something not worth attempting.
 */
function pinCheck(pin) {
  var stored = PropertiesService.getScriptProperties().getProperty('PIN');
  if (!stored) return 'No PIN is set on this script. Run makePin() in the Apps Script editor.';

  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('pinFails') || 0);
  if (fails >= MAX_TRIES) {
    return 'Too many wrong PINs. This will unlock itself in a few minutes.';
  }

  if (String(pin || '') !== stored) {
    cache.put('pinFails', String(fails + 1), LOCK_SECONDS);
    var left = MAX_TRIES - fails - 1;
    return 'That PIN is not correct.' +
           (left <= 3 ? ' ' + left + ' attempt' + (left === 1 ? '' : 's') + ' left before a pause.' : '');
  }

  if (fails) cache.remove('pinFails');
  return null;
}

/**
 * BEST WAY TO SET A PIN. Run this once, then read the execution log: it prints
 * a strong random passphrase and stores it. Because you never type it into the
 * code, it can never end up in your GitHub repo or its history.
 *
 * Run it again any time to rotate the PIN, which is what to do if you ever
 * think it has leaked.
 */
function makePin() {
  var words = 'amber,anchor,basil,cedar,delta,ember,falcon,garnet,harbor,indigo,jasper,kettle,' +
              'lantern,marble,nimbus,onyx,pepper,quartz,rowan,saffron,timber,umber,violet,willow';
  var list = words.split(',');
  var pick = [];
  for (var i = 0; i < 4; i++) {
    pick.push(list[Math.floor(Math.random() * list.length)]);
  }
  var pin = pick.join('-') + '-' + Math.floor(Math.random() * 9000 + 1000);
  PropertiesService.getScriptProperties().setProperty('PIN', pin);
  Logger.log('Your new PIN is:\n\n    ' + pin +
             '\n\nType it into the site once per device. Nothing else to do.');
  return pin;
}

/** Set a PIN of your own instead. Edit the value, run once, then change it back. */
function setPin() {
  PropertiesService.getScriptProperties().setProperty('PIN', '1234');
}

/** Run this to confirm a PIN really is stored. Check the execution log. */
function checkPinIsSet() {
  var stored = PropertiesService.getScriptProperties().getProperty('PIN');
  Logger.log(stored ? 'PIN is set (' + stored.length + ' characters).' : 'NO PIN IS SET.');
}

/** Clears a lockout early, if you locked yourself out by mistyping. */
function unlockNow() {
  CacheService.getScriptCache().remove('pinFails');
  Logger.log('Lockout cleared.');
}

/* ================================================================== */
/* Setup and upgrade                                                  */
/* ================================================================== */

function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensure(ss, SHEET_STUDENTS, ['Roll', 'Name', 'Section']);
  ensure(ss, SHEET_COURSES, ['Code', 'Title', 'Section', 'Type']);
  ensure(ss, SHEET_ATTENDANCE, ATTENDANCE_HEADERS);
  ensure(ss, SHEET_COMPONENTS, COMPONENT_HEADERS);
  ensure(ss, SHEET_MARKS, MARKS_HEADERS);
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var done = [];

  // Components and Marks arrived with internal marks.
  ensure(ss, SHEET_COMPONENTS, COMPONENT_HEADERS);
  ensure(ss, SHEET_MARKS, MARKS_HEADERS);

  // The Components column was called Weight before marks became raw scores.
  var comps = ss.getSheetByName(SHEET_COMPONENTS);
  if (comps && String(comps.getRange(1, 4).getValue()).trim() === 'Weight') {
    comps.getRange(1, 4).setValue('Marks');
    done.push('renamed the Components "Weight" column to "Marks"');
  }

  // Courses gained a Type column (Theory or Sessional).
  var courses = ss.getSheetByName(SHEET_COURSES);
  if (courses && String(courses.getRange(1, 4).getValue()).trim() !== 'Type') {
    courses.getRange(1, 4).setValue('Type').setFontWeight('bold').setBackground('#eef1f6');
    done.push('added the Type column to Courses (blank means Theory)');
  }

  // Attendance gained a Class column when multiple classes per day arrived.
  var sheet = ss.getSheetByName(SHEET_ATTENDANCE);
  if (!sheet) throw new Error('Sheet "Attendance" not found. Run setupSheets() first.');
  if (String(sheet.getRange(1, 3).getValue()).trim() !== 'Class') {
    sheet.insertColumnBefore(3);
    sheet.getRange(1, 3).setValue('Class').setFontWeight('bold').setBackground('#eef1f6');
    var rows = sheet.getLastRow() - 1;
    if (rows > 0) {
      var fill = [];
      for (var i = 0; i < rows; i++) fill.push(['1']);
      sheet.getRange(2, 3, rows, 1).setValues(fill);
    }
    done.push('added the Class column to ' + rows + ' attendance rows');
  }

  return done.length ? 'Upgraded: ' + done.join('; ') + '.' : 'Already up to date.';
}

/**
 * Fills in a starting set of components for one course, based on its Type.
 * The site offers this as a button; you can also run it by hand.
 */
function defaultComponentsFor(type) {
  return DEFAULT_COMPONENTS[type === 'Sessional' ? 'Sessional' : 'Theory'];
}

/* ================================================================== */

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
