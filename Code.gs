// ============================================================
// McIntosh County Curriculum Map Dashboard - Live Backend
// Google Apps Script (doGet -> JSON)
// 2026-2027
// ============================================================
//
// WHAT THIS DOES:
//   Reads the TGE, MCMS, and MCA tabs of the curriculum map
//   workbook and returns one normalized JSON payload. The
//   dashboard HTML fetches this on load and renders it.
//
// WHY IT CANNOT BE BLOCKED BY IT:
//   It uses ONLY SpreadsheetApp.getActiveSpreadsheet(). That
//   call needs no restricted OAuth scope, so the district API
//   Controls wall that killed the old DriveApp/openById version
//   does not apply. It runs as you (Execute as: Me), so no
//   viewer ever authorizes anything.
//
// CHANGES IN THIS VERSION (July 2026), re-validated against the
// live 26-27 workbook header rows, column by column:
//
//   1. SUBJECT HEADERS are detected by BOTH shape and formatting:
//      a row is a subject band only when column A is the ONLY cell
//      with content (a sparse row) AND the row is bold or shaded.
//      - Requiring "sparse" stops shaded DATA rows (e.g. TGE, where
//        the whole grid is colored) from being mistaken for headers.
//      - Requiring "bold/shaded" stops a teacher who simply has not
//        entered unit data yet (just a name in col A) from being
//        mistaken for a header and dropped, which also used to pin
//        that teacher's name onto everyone below as a bogus subject
//        and scramble the row order.
//      Read via getFontWeights()/getBackgrounds(). Run diagnose()
//      to see exactly how each tab's rows are shaped and formatted.
//
//   2. CURRICULUM MAP COLUMN is now read and emitted for all
//      three schools:
//        MCA  -> column B (index 1)
//        MCMS -> column B (index 1)
//        TGE  -> column C (index 2)
//      Previously TGE skipped it, MCMS misread it as the subject,
//      and MCA mislabeled it as the teacher's "course".
//
//   3. MCMS MAP REBUILT to match the live sheet: 13 units, every
//      unit 4 columns wide (Plan, Test Date, Test, Test Talk?).
//      The old map stopped at Unit 10 and treated Units 6-10 as
//      3 wide, so Units 6-13 were read from the wrong columns and
//      Units 11-13 were dropped entirely.
//
//   4. MCMS NO LONGER READS A SUBJECT COLUMN. The MCMS tab has no
//      subject column (col B is Curriculum Map). Subjects, if the
//      tab uses them, are bold/shaded header rows in column A,
//      handled the same way as MCA and TGE.
//
//   (MCA Unit 7's Shared?/Test Talk? swap and the 11-unit MCA
//    layout are real in the sheet and are intentionally kept.)
//
// SETUP:
//   1. Open the bound sheet, Extensions > Apps Script.
//   2. Replace the existing code with this entire file. Save.
//   3. Run testPayload() once from the function dropdown and
//      check the Execution log for row counts, the header guard
//      line "MCA header skipped: true", and the MCMS unit count
//      line "MCMS units read: 13".
//   4. Deploy > Manage deployments > pencil icon > Version:
//      New version > Deploy. (Editing the existing deployment
//      keeps the same /exec URL, so the HTML needs no change.
//      Creating a brand-new deployment would change the URL.)
//        Execute as: Me
//        Who has access: Anyone        <-- must be "Anyone",
//          NOT "Anyone within McIntosh County Schools".
// ============================================================

// Tab names exactly as they appear in the workbook
var TABS = ['MCA', 'TGE', 'MCMS'];

// ---- Curriculum Map column per tab (0-indexed) ----
var MCA_MAP_COL  = 1;   // MCA  column B
var MCMS_MAP_COL = 1;   // MCMS column B
var TGE_MAP_COL  = 2;   // TGE  column C

// ---- Column maps (0-indexed), validated against the live sheet ----

// MCA: [plan, date, shared, testTalk] per unit (11 units).
// Header: A Teacher/Grade | B Curriculum Map | C Unit 1 Plan |
// D Test Date | E Shared? | F Test Talk? | ... Unit 7 has
// Test Talk?/Shared? swapped IN THE SHEET (col 28 Test Talk?,
// col 29 Shared?), so the map compensates.
var MCA_UNITS = [
  [2, 3, 4, 5], [6, 7, 8, 9], [10, 11, 12, 13], [14, 15, 16, 17],
  [18, 19, 20, 21], [22, 23, 24, 25],
  [26, 27, 29, 28],                 // Unit 7: shared=29, testTalk=28 (swapped)
  [30, 31, 32, 33], [34, 35, 36, 37], [38, 39, 40, 41],
  [42, 43, 44, 45]                  // Unit 11: full 4 columns
];

// TGE: [test, date, shared, bwd, testTalk] per unit (7 units).
// Header: A Teacher/Grade | B Deconstructed Standards |
// C Curriculum Map | D Unit 1 Test | E (date) | F (shared) |
// G BWD | H Test Talk | ... first unit therefore starts at col D
// (index 3); columns B and C are not part of any unit.
var TGE_UNITS = [
  [3, 4, 5, 6, 7], [8, 9, 10, 11, 12], [13, 14, 15, 16, 17],
  [18, 19, 20, 21, 22], [23, 24, 25, 26, 27], [28, 29, 30, 31, 32],
  [33, 34, 35, 36, 37]
];

// MCMS: [plan, date, test, testTalk] per unit (13 units, every
// unit 4 columns wide). Header: A Teacher/Grade | B Curriculum
// Map | C Unit 1 Plan | D Test Date | E Test | F Test Talk? | ...
var MCMS_UNITS = [
  [2, 3, 4, 5], [6, 7, 8, 9], [10, 11, 12, 13], [14, 15, 16, 17],
  [18, 19, 20, 21], [22, 23, 24, 25], [26, 27, 28, 29], [30, 31, 32, 33],
  [34, 35, 36, 37], [38, 39, 40, 41], [42, 43, 44, 45], [46, 47, 48, 49],
  [50, 51, 52, 53]
];

// ---- ENTRY POINT ----
function doGet() {
  try {
    var payload = buildPayload();
    return jsonOut({ status: 'success', updatedAt: new Date().toISOString(), schools: payload });
  } catch (err) {
    return jsonOut({ status: 'error', message: String(err) });
  }
}

// ---- BUILD THE FULL PAYLOAD ----
function buildPayload() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  return {
    MCA:  readMCA(ss.getSheetByName('MCA'), tz),
    TGE:  readTGE(ss.getSheetByName('TGE'), tz),
    MCMS: readMCMS(ss.getSheetByName('MCMS'), tz)
  };
}

// ---- READERS ----
// Each reader pulls values plus formatting once, walks the rows
// top-to-bottom (preserving sheet order), skips blank/header
// rows, treats bold-or-shaded rows as subject headers, and emits
// every remaining teacher row.

function readMCA(sheet, tz) {
  var m = sheetMatrix(sheet);
  var out = [], subject = '';
  for (var r = 1; r < m.values.length; r++) {
    var row = m.values[r];
    if (blankRow(row)) continue;
    if (isHeaderRow(row, tz)) continue;          // MCA header sits on sheet row 3
    if (isSubjectHeader(m, r, tz)) { subject = norm(row[0], tz); continue; }
    if (norm(row[0], tz) === '') continue;
    var units = MCA_UNITS.map(function (u) {
      return {
        plan:     u[0] === null ? '' : norm(row[u[0]], tz),
        date:     norm(row[u[1]], tz),
        shared:   norm(row[u[2]], tz),
        testTalk: row[u[3]] === true
      };
    });
    out.push({
      subject: subject,
      name: norm(row[0], tz),
      curriculumMap: mapVal(row[MCA_MAP_COL], tz),
      units: units
    });
  }
  return out;
}

function readTGE(sheet, tz) {
  var m = sheetMatrix(sheet);
  var out = [], subject = '';
  for (var r = 1; r < m.values.length; r++) {
    var row = m.values[r];
    if (blankRow(row)) continue;
    if (isHeaderRow(row, tz)) continue;
    if (isSubjectHeader(m, r, tz)) { subject = norm(row[0], tz); continue; }
    if (norm(row[0], tz) === '') continue;
    var units = TGE_UNITS.map(function (u) {
      return {
        test:     norm(row[u[0]], tz),
        date:     norm(row[u[1]], tz),
        shared:   norm(row[u[2]], tz),
        bwd:      norm(row[u[3]], tz),
        testTalk: row[u[4]] === true
      };
    });
    out.push({
      subject: subject,
      grade: norm(row[0], tz),
      curriculumMap: mapVal(row[TGE_MAP_COL], tz),
      units: units
    });
  }
  return out;
}

function readMCMS(sheet, tz) {
  var m = sheetMatrix(sheet);
  var out = [], subject = '';
  for (var r = 1; r < m.values.length; r++) {
    var row = m.values[r];
    if (blankRow(row)) continue;
    if (isHeaderRow(row, tz)) continue;
    if (isSubjectHeader(m, r, tz)) { subject = norm(row[0], tz); continue; }
    if (norm(row[0], tz) === '') continue;
    var units = MCMS_UNITS.map(function (u) {
      return {
        plan:     norm(row[u[0]], tz),
        date:     norm(row[u[1]], tz),
        test:     norm(row[u[2]], tz),
        testTalk: row[u[3]] === true
      };
    });
    out.push({
      subject: subject,
      teacherGrade: norm(row[0], tz),
      curriculumMap: mapVal(row[MCMS_MAP_COL], tz),
      units: units
    });
  }
  return out;
}

// ---- HELPERS ----

// Read values + formatting for the whole used range in one shot.
function sheetMatrix(sheet) {
  var rng = sheet.getDataRange();
  return {
    values:      rng.getValues(),
    weights:     rng.getFontWeights(),
    backgrounds: rng.getBackgrounds()
  };
}

function norm(v, tz) {
  if (v === null || v === undefined || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, tz, 'M/d/yyyy');
  }
  return String(v).trim();
}

// Curriculum Map cell -> boolean for a checkbox, or the trimmed
// text/link for anything else. Empty stays ''. The dashboard
// treats true or non-empty text as "submitted".
function mapVal(v, tz) {
  if (v === true) return true;
  if (v === false) return false;
  return norm(v, tz);
}

// Header rows start with "Teacher/Grade" in column A. TGE and
// MCMS have their header on sheet row 1 (already skipped by the
// loop starting at r=1); MCA's header sits on sheet row 3 after
// two blank rows, so it must be filtered explicitly.
function isHeaderRow(row, tz) {
  return norm(row[0], tz) === 'Teacher/Grade';
}

// A subject band is a row where column A is the only cell with
// content AND the row is bold or shaded. Both signals are required:
// "sparse" alone would drop teachers who have not entered data yet;
// "formatted" alone would swallow shaded data rows (TGE colors its
// whole grid). Together they isolate the subject bands (ELA, Math).
function isSubjectHeader(m, r, tz) {
  var row = m.values[r];
  if (norm(row[0], tz) === '') return false;   // no label in col A -> not a header
  if (!onlyFirstColFilled(row)) return false;  // real data elsewhere -> a teacher row
  return rowIsBoldOrShaded(m, r);
}

// True when only column A holds real content (every other cell is
// empty or an unchecked checkbox).
function onlyFirstColFilled(row) {
  for (var j = 1; j < row.length; j++) {
    var c = row[j];
    if (c !== '' && c !== null && c !== undefined && c !== false) return false;
  }
  return true;
}

function rowIsBoldOrShaded(m, r) {
  var w = m.weights[r], b = m.backgrounds[r], j;
  for (j = 0; j < w.length; j++) {
    if (w[j] === 'bold') return true;
  }
  for (j = 0; j < b.length; j++) {
    var c = b[j];
    if (!c) continue;
    var s = String(c).toLowerCase();
    if (s !== '' && s !== '#ffffff' && s !== 'white' && s !== '#fff') return true;
  }
  return false;
}

function blankRow(row) {
  for (var j = 0; j < row.length; j++) {
    var c = row[j];
    if (c !== '' && c !== null && c !== undefined && c !== false) return false;
  }
  return true;
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- RUN-ONCE SANITY CHECK ----
// Select this in the dropdown and click Run. Approve permissions.
function testPayload() {
  var p = buildPayload();
  Logger.log('MCA ' + p.MCA.length + ' | TGE ' + p.TGE.length + ' | MCMS ' + p.MCMS.length);

  var bogus = p.MCA.some(function (t) { return t.name === 'Teacher/Grade'; });
  Logger.log('MCA header skipped: ' + !bogus);

  var mcmsUnits = p.MCMS.length ? p.MCMS[0].units.length : 0;
  Logger.log('MCMS units read: ' + mcmsUnits + ' (expected 13)');

  // Spot-check that the Curriculum Map column is coming through.
  Logger.log('MCA[0] curriculumMap: ' + JSON.stringify(p.MCA.length ? p.MCA[0].curriculumMap : null));
  Logger.log('TGE[0] curriculumMap: ' + JSON.stringify(p.TGE.length ? p.TGE[0].curriculumMap : null));
  Logger.log('MCMS[0] curriculumMap: ' + JSON.stringify(p.MCMS.length ? p.MCMS[0].curriculumMap : null));

  var reece = p.MCMS.filter(function (t) { return t.teacherGrade.indexOf('Reece') === 0; })[0];
  if (reece) Logger.log('Reece last unit (U' + reece.units.length + '): ' + JSON.stringify(reece.units[reece.units.length - 1]));

  Logger.log(JSON.stringify(p.MCA[0], null, 2));
}

// ---- FORMATTING / SHAPE DIAGNOSTIC ----
// Select this in the dropdown and Run. For each tab it prints the
// first 20 non-blank rows with: bold?, column A background color,
// sparse? (only col A filled), and the first three column values.
// Use it to confirm how subject bands differ from teacher rows so
// isSubjectHeader can be tuned if a tab hides teachers or leaks
// headers.
function diagnose() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  ['MCA', 'TGE', 'MCMS'].forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) { Logger.log(name + ': NOT FOUND'); return; }
    var m = sheetMatrix(sheet);
    Logger.log('===== ' + name + ' : ' + m.values.length + ' rows x ' +
               (m.values[0] ? m.values[0].length : 0) + ' cols =====');
    var shown = 0;
    for (var r = 0; r < m.values.length && shown < 20; r++) {
      var row = m.values[r];
      if (blankRow(row)) continue;
      shown++;
      var bold = false;
      for (var j = 0; j < m.weights[r].length; j++) {
        if (m.weights[r][j] === 'bold') { bold = true; break; }
      }
      Logger.log('r' + r +
        ' | bold=' + bold +
        ' | bgA=' + m.backgrounds[r][0] +
        ' | sparse=' + onlyFirstColFilled(row) +
        ' | header=' + isSubjectHeader(m, r, tz) +
        ' | A="' + norm(row[0], tz) + '" B="' + norm(row[1], tz) + '" C="' + norm(row[2], tz) + '"');
    }
  });
}

// ---- DATE AUDIT ----
// Flags every Test Date cell that holds TEXT instead of a real
// date value. Text dates (e.g. "Sept. 30") get mis-parsed by the
// dashboard to the 15th of the month and cannot be validated
// against the calendar. Convert the flagged cells to real dates
// (Format > Number > Date) before locking. Re-run until 0 text.
function auditDates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();

  // Test Date column index per unit (0-based), derived from the maps above.
  var specs = [
    { name: 'MCA',  dateCols: MCA_UNITS.map(function (u) { return u[1]; }) },
    { name: 'TGE',  dateCols: TGE_UNITS.map(function (u) { return u[1]; }) },
    { name: 'MCMS', dateCols: MCMS_UNITS.map(function (u) { return u[1]; }) }
  ];

  var totText = 0, totBlank = 0;
  Logger.log('=== DATE AUDIT: text cells must be fixed before locking ===');

  specs.forEach(function (spec) {
    var sheet = ss.getSheetByName(spec.name);
    if (!sheet) { Logger.log(spec.name + ': tab not found'); return; }
    var m = sheetMatrix(sheet);
    var text = 0, blank = 0;

    for (var r = 1; r < m.values.length; r++) {
      var row = m.values[r];
      if (blankRow(row)) continue;
      if (isHeaderRow(row, tz)) continue;
      if (isSubjectHeader(m, r, tz)) continue;
      if (norm(row[0], tz) === '') continue;   // no teacher label = not a teacher row

      var label = norm(row[0], tz);
      spec.dateCols.forEach(function (col, i) {
        var v = row[col];
        if (Object.prototype.toString.call(v) === '[object Date]') return; // real date, ok
        var s = norm(v, tz);
        if (s === '') { blank++; return; }                                 // empty, just unset
        text++;
        Logger.log(spec.name + ' · ' + label + ' · Unit ' + (i + 1) + ' · "' + s + '"');
      });
    }

    Logger.log('--- ' + spec.name + ': ' + text + ' text date(s) to fix, ' + blank + ' empty ---');
    totText += text; totBlank += blank;
  });

  Logger.log('=== TOTAL: ' + totText + ' text dates to convert, ' + totBlank + ' empty ===');
}
