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
//   1. SUBJECT BANDS are detected by the COLORED FILL in column A.
//      Every teacher row has a white column A; only the subject
//      bands (ELA, Math, Science, ...) are filled with a color. This
//      is the workbook's consistent convention on all three tabs.
//      The earlier rules failed on the live data: "only column A
//      has text" dropped teachers who had no data yet and pinned
//      their name onto everyone below as a bogus subject; "bold or
//      any shaded cell" wiped out all of TGE (its whole grid is
//      colored) and dropped MCMS teachers with a stray shaded cell.
//      Read via getBackgrounds(). Run diagnose() to see the fill of
//      every row per tab.
//
//   2. CURRICULUM MAP COLUMN is now read and emitted for all
//      three schools:
//        MCA  -> column B (index 1)
//        MCMS -> column B (index 1)
//        TGE  -> column C (index 2)
//      Previously TGE skipped it, MCMS misread it as the subject,
//      and MCA mislabeled it as the teacher's "course". The cell's
//      hyperlink is also emitted as curriculumMapUrl (via
//      columnLinkUrls) so the dashboard can link to the map; it
//      covers text links, first-run links, and =HYPERLINK() cells.
//
//   3. MCMS MAP REBUILT to match the live sheet: 13 units, every
//      unit 4 columns wide (Plan, Test Date, Test, Test Talk?).
//      The old map stopped at Unit 10 and treated Units 6-10 as
//      3 wide, so Units 6-13 were read from the wrong columns and
//      Units 11-13 were dropped entirely.
//
//   4. MCMS SUBJECT BANDS live in column B with column A blank
//      (e.g. A "" | B "ELA"), so subjectLabel() reads the band's
//      label from column A, or column B when A is empty. Column B on
//      a teacher row is that teacher's Curriculum Map, not a subject.
//
//   5. UNIT RUN DATES (when the unit is TAUGHT, distinct from the
//      Test Date when the assessment is given) are read per teacher
//      per unit from columns titled "Unit n Start" and "Unit n End".
//      Add those columns to the RIGHT of each tab's existing columns,
//      in the same header row as "Teacher/Grade" (one Start and one
//      End per unit: 11 for MCA, 7 for TGE, 13 for MCMS). They are
//      matched by header text via headerMap(), so exact placement
//      does not matter and the payload is unaffected until you add
//      them. Enter the dates as REAL dates, not typed text.
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
// top-to-bottom (preserving sheet order), skips blank/title rows,
// treats color-filled rows as subject bands, and emits every
// remaining teacher row.

function readMCA(sheet, tz) {
  var m = sheetMatrix(sheet);
  var hmap = headerMap(m, tz);
  var mapUrls = columnLinkUrls(sheet, MCA_MAP_COL, m.values.length);
  var out = [], subject = '';
  for (var r = 1; r < m.values.length; r++) {
    var row = m.values[r];
    if (blankRow(row)) continue;
    if (isHeaderRow(row, tz)) continue;          // MCA header sits on sheet row 3
    if (isSubjectHeader(m, r, tz)) { subject = subjectLabel(row, tz); continue; }
    if (norm(row[0], tz) === '') continue;
    var units = MCA_UNITS.map(function (u, i) {
      var run = unitRunDates(hmap, row, i + 1, tz);
      return {
        plan:     u[0] === null ? '' : norm(row[u[0]], tz),
        date:     norm(row[u[1]], tz),
        shared:   norm(row[u[2]], tz),
        testTalk: row[u[3]] === true,
        start:    run.start,
        end:      run.end
      };
    });
    out.push({
      subject: subject,
      name: norm(row[0], tz),
      curriculumMap: mapVal(row[MCA_MAP_COL], tz),
      curriculumMapUrl: mapUrls[r],
      units: units
    });
  }
  return out;
}

function readTGE(sheet, tz) {
  var m = sheetMatrix(sheet);
  var hmap = headerMap(m, tz);
  var mapUrls = columnLinkUrls(sheet, TGE_MAP_COL, m.values.length);
  var out = [], subject = '';
  for (var r = 1; r < m.values.length; r++) {
    var row = m.values[r];
    if (blankRow(row)) continue;
    if (isHeaderRow(row, tz)) continue;
    if (isSubjectHeader(m, r, tz)) { subject = subjectLabel(row, tz); continue; }
    if (norm(row[0], tz) === '') continue;
    var units = TGE_UNITS.map(function (u, i) {
      var run = unitRunDates(hmap, row, i + 1, tz);
      return {
        test:     norm(row[u[0]], tz),
        date:     norm(row[u[1]], tz),
        shared:   norm(row[u[2]], tz),
        bwd:      norm(row[u[3]], tz),
        testTalk: row[u[4]] === true,
        start:    run.start,
        end:      run.end
      };
    });
    out.push({
      subject: subject,
      grade: norm(row[0], tz),
      curriculumMap: mapVal(row[TGE_MAP_COL], tz),
      curriculumMapUrl: mapUrls[r],
      units: units
    });
  }
  return out;
}

function readMCMS(sheet, tz) {
  var m = sheetMatrix(sheet);
  var hmap = headerMap(m, tz);
  var mapUrls = columnLinkUrls(sheet, MCMS_MAP_COL, m.values.length);
  var out = [], subject = '';
  for (var r = 1; r < m.values.length; r++) {
    var row = m.values[r];
    if (blankRow(row)) continue;
    if (isHeaderRow(row, tz)) continue;
    if (isSubjectHeader(m, r, tz)) { subject = subjectLabel(row, tz); continue; }
    if (norm(row[0], tz) === '') continue;
    var units = MCMS_UNITS.map(function (u, i) {
      var run = unitRunDates(hmap, row, i + 1, tz);
      return {
        plan:     norm(row[u[0]], tz),
        date:     norm(row[u[1]], tz),
        test:     norm(row[u[2]], tz),
        testTalk: row[u[3]] === true,
        start:    run.start,
        end:      run.end
      };
    });
    out.push({
      subject: subject,
      teacherGrade: norm(row[0], tz),
      curriculumMap: mapVal(row[MCMS_MAP_COL], tz),
      curriculumMapUrl: mapUrls[r],
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

// Extract the hyperlink URL from every cell of one column, indexed by
// row (0-based). Handles all the ways a Curriculum Map link can be
// stored: a link applied to the cell text, a link on the first run of
// mixed text, or a =HYPERLINK("url","label") formula. Cells with no
// link return ''. Two batched calls, so it stays cheap.
function columnLinkUrls(sheet, col0, numRows) {
  if (numRows < 1) return [];
  var rng = sheet.getRange(1, col0 + 1, numRows, 1);
  var rich = rng.getRichTextValues();
  var forms = rng.getFormulas();
  var out = [];
  for (var i = 0; i < numRows; i++) {
    out.push(extractLinkUrl(rich[i][0], forms[i][0]));
  }
  return out;
}

function extractLinkUrl(rtv, formula) {
  if (rtv) {
    var u = rtv.getLinkUrl();
    if (u) return u;
    var runs = rtv.getRuns();
    for (var j = 0; j < runs.length; j++) {
      var ru = runs[j].getLinkUrl();
      if (ru) return ru;
    }
  }
  if (formula && /^=HYPERLINK\(/i.test(formula)) {
    var m = formula.match(/=HYPERLINK\(\s*"([^"]+)"/i);
    if (m) return m[1];
  }
  return '';
}

// The row that holds the column titles ("Teacher/Grade" in col A).
// TGE/MCMS use row 1; MCA uses row 3, so find it rather than assume.
function headerRowIndex(m, tz) {
  for (var r = 0; r < m.values.length; r++) {
    if (norm(m.values[r][0], tz) === 'Teacher/Grade') return r;
  }
  return 0;
}

// Map of normalized header text -> column index, so new columns can
// be found by their title (e.g. "Unit 3 Start") no matter where they
// sit. First occurrence of a title wins; repeated titles like
// "Test Date" are intentionally left to the fixed column maps above.
function headerMap(m, tz) {
  var map = {}, row = m.values[headerRowIndex(m, tz)];
  for (var c = 0; c < row.length; c++) {
    var h = norm(row[c], tz).toLowerCase().replace(/\s+/g, ' ');
    if (h !== '' && !(h in map)) map[h] = c;
  }
  return map;
}

// Unit run window (when the unit is TAUGHT) for unit number n, read
// from the appended "Unit n Start" / "Unit n End" columns by header.
// Returns '' for either side that has no column or no value yet, so
// the payload is unaffected until those columns are added.
function unitRunDates(hmap, row, n, tz) {
  var sc = hmap['unit ' + n + ' start'];
  var ec = hmap['unit ' + n + ' end'];
  return {
    start: (sc !== undefined) ? norm(row[sc], tz) : '',
    end:   (ec !== undefined) ? norm(row[ec], tz) : ''
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

// Subject bands are the ONLY rows with a colored fill in column A;
// every teacher row has a white column A. This is the workbook's
// consistent visual convention across all three tabs, confirmed by
// diagnose():
//   - MCA/TGE put the subject label in column A (colored fill).
//   - MCMS leaves column A blank and puts the subject in column B,
//     but the band's column A is still colored.
// (The "Teacher/Grade" title row is also colored but is filtered
// earlier by isHeaderRow.) Detecting bands by fill color instead of
// "sparse + bold" avoids two failures seen in the live data: TGE
// teacher rows are bold (so bold cannot mark a band), and some
// teacher rows carry a stray shaded cell (so "any shaded cell in the
// row" wrongly dropped teachers such as MCMS Way-7 and Teschendorf-6
// who simply had no data entered yet).
function isSubjectHeader(m, r, tz) {
  return isColoredFill(m.backgrounds[r][0]);
}

// Subject label for a band: MCA/TGE hold it in column A; MCMS leaves
// column A blank and holds the subject in column B.
function subjectLabel(row, tz) {
  var a = norm(row[0], tz);
  return a !== '' ? a : norm(row[1], tz);
}

function isColoredFill(bg) {
  if (!bg) return false;
  var s = String(bg).toLowerCase();
  return s !== '' && s !== '#ffffff' && s !== 'white' && s !== '#fff';
}

// True when only column A holds real content (used by diagnose()).
function onlyFirstColFilled(row) {
  for (var j = 1; j < row.length; j++) {
    var c = row[j];
    if (c !== '' && c !== null && c !== undefined && c !== false) return false;
  }
  return true;
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
  Logger.log('MCA[0] curriculumMapUrl: ' + JSON.stringify(p.MCA.length ? p.MCA[0].curriculumMapUrl : null));

  // Unit run dates (Start/End) - blank until the columns are added.
  Logger.log('MCA[0] Unit 1 run: start="' + (p.MCA.length ? p.MCA[0].units[0].start : '') +
             '" end="' + (p.MCA.length ? p.MCA[0].units[0].end : '') + '"');

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
