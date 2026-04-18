/**
 * Lesson Logger — Google Apps Script API
 * Version: v2026-04-18 19:00 UTC (standalone-safe + respects Facturas "Active" column)
 *
 * Paste this into Extensions → Apps Script in your lesson-logger Google Sheet
 * (the one with Lessons / Students / Config tabs). Deploy as Web App
 * (Execute as: Me, Access: Anyone).
 *
 * Architecture notes:
 * - Parents and students live in the "Facturas" intake spreadsheet (a SEPARATE
 *   spreadsheet, not this one). The intake form writes to it automatically.
 *   We read from it at request time, with deterministic IDs so parent/student
 *   identity is stable across refreshes.
 * - This spreadsheet keeps:
 *     - Lessons: the canonical lesson log (id, studentId, parentId, date, …)
 *     - Students: now an OVERRIDES table only (per-student pricing overrides)
 *     - Config: standard pricing + API URL
 *     - Parents: left in place as an archive from the old architecture, but
 *       no longer read from. Safe to clear manually.
 * - Add / edit / bulk-import of parents & students is DISABLED here —
 *   everything comes from the Facturas intake form. The frontend hides those
 *   screens; the backend returns a clear error if called anyway.
 *
 * One-time setup after pasting:
 *   1. Open appsscript.json (enable in Project Settings if needed) and make
 *      sure "oauthScopes" includes:
 *        "https://www.googleapis.com/auth/spreadsheets"
 *        "https://www.googleapis.com/auth/script.external_request"
 *      (or delete the oauthScopes array entirely to let Apps Script
 *      auto-detect).
 *   2. Run `authorizeAll` from the editor dropdown → Allow on the prompt.
 *   3. Run `migrateDryRun` from the editor. Inspect the execution log.
 *   4. If the diagnostics look right, run `migrateCommit`. This remaps
 *      existing Lessons rows from old parentId/studentId to Facturas-derived
 *      IDs, and rebuilds the Students sheet keyed on the new IDs while
 *      preserving your pricing overrides.
 *   5. Deploy → Manage deployments → Edit → New version → Deploy. Same URL.
 */

// ─── Configuration ───────────────────────────────────────────────────────────

// The Lesson Logger's OWN spreadsheet (the one with Lessons / Students / Config).
// Required when this Apps Script project is STANDALONE (not container-bound).
// If container-bound (opened via Extensions → Apps Script from the sheet),
// leave this as '' and getActiveSpreadsheet() will be used automatically.
// Find the ID in the sheet URL: docs.google.com/spreadsheets/d/<THIS_PART>/edit
var LESSON_LOGGER_SPREADSHEET_ID = '1DWAPfW721G-uIIapqrDkL24B6pHATS8sCSli4WDolbs';

// The "Facturas" intake spreadsheet (owned by santiagoplopez05@gmail.com).
// Matches invoice-apps-script.gs — same sheet, same tab.
var FACTURAS_SPREADSHEET_ID = '11WEK_G0_OhA18VxuzGziXnkvwi8FIT-JLsaxqlgNmbo';
var FACTURAS_PARENTS_TAB_GID = 1134169331;

var COL_STUDENT_NAME = "Student's Full Name";
var COL_PARENT_NAME  = "Parent/Guardian's Full Name";
var COL_PARENT_EMAIL = "Parent/Guardian's Email";
var COL_ACTIVE       = "Active";  // optional; "NO" (case-insensitive) excludes the row

// ─── One-time authorization helper ───────────────────────────────────────────
//
// Apps Script only grants scopes after the owner approves them in an
// interactive prompt from the editor. Run this once from the editor after
// pasting the file — it touches both scopes the script needs.

function authorizeAll() {
  var ss = SpreadsheetApp.openById(FACTURAS_SPREADSHEET_ID);
  var sheet = findSheetByGid(ss, FACTURAS_PARENTS_TAB_GID);
  var tabName = sheet ? sheet.getName() : '(no tab with that gid)';
  var msg = 'OK — Facturas title: "' + ss.getName() + '", tab: "' + tabName + '"';
  Logger.log(msg);
  return msg;
}

// ─── CORS & Routing ──────────────────────────────────────────────────────────

function doGet(e) {
  var action = e.parameter.action || '';
  var result;

  try {
    switch (action) {
      case 'getAll':
        result = getAllData();
        break;
      case 'getParents':
        result = getParents();
        break;
      case 'getStudents':
        result = getStudents();
        break;
      case 'getLessons':
        result = getLessons(e.parameter.month);
        break;
      case 'getConfig':
        result = getConfig();
        break;
      default:
        result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'Invalid JSON' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var action = data.action || '';
  var result;

  try {
    switch (action) {
      // Read operations (also via POST to avoid CORS)
      case 'getAll':
        result = getAllData();
        break;
      case 'getParents':
        result = getParents();
        break;
      case 'getStudents':
        result = getStudents();
        break;
      case 'getLessons':
        result = getLessons(data.month);
        break;
      case 'getConfig':
        result = getConfig();
        break;
      // Write operations for parents/students — DISABLED (Facturas is SoT)
      case 'addParent':
      case 'updateParent':
      case 'deleteParent':
      case 'addStudent':
      case 'deleteStudent':
      case 'bulkImportParents':
        result = { error: 'Parents and students are managed in the Facturas intake sheet. This action is no longer supported here.' };
        break;
      // updateStudent is still allowed — used for per-student pricing overrides
      case 'updateStudent':
        result = updateStudent(data);
        break;
      case 'addLesson':
        result = addLesson(data);
        break;
      case 'updateLesson':
        result = updateLesson(data);
        break;
      case 'deleteLesson':
        result = deleteLesson(data.id);
        break;
      case 'updateConfig':
        result = updateConfig(data.config);
        break;
      default:
        result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Generic helpers ─────────────────────────────────────────────────────────

function _getLoggerSpreadsheet() {
  if (LESSON_LOGGER_SPREADSHEET_ID) {
    return SpreadsheetApp.openById(LESSON_LOGGER_SPREADSHEET_ID);
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error(
      'This Apps Script project is not bound to a spreadsheet. ' +
      'Set LESSON_LOGGER_SPREADSHEET_ID at the top of apps-script.gs ' +
      'to the Lesson Logger spreadsheet ID (the <ID> in ' +
      'docs.google.com/spreadsheets/d/<ID>/edit).'
    );
  }
  return ss;
}

function getSheet(name) {
  return _getLoggerSpreadsheet().getSheetByName(name);
}

function sheetToObjects(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    result.push(obj);
  }
  return result;
}

function findRowById(sheet, id) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1; // 1-indexed row
  }
  return -1;
}

function generateId() {
  return 'x' + Utilities.getUuid().replace(/-/g, '').substring(0, 7);
}

// ─── ID derivation (deterministic, stable across refreshes) ──────────────────

function sha1hex(s) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_1, String(s));
  var out = '';
  for (var i = 0; i < bytes.length; i++) {
    var v = bytes[i] < 0 ? bytes[i] + 256 : bytes[i];
    var hex = v.toString(16);
    if (hex.length < 2) hex = '0' + hex;
    out += hex;
  }
  return out;
}

function normalizeName(s) {
  return String(s || '')
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // strip accents
    .replace(/[^\w\s]/g, ' ')                          // drop punctuation
    .replace(/\s+/g, ' ')                              // collapse whitespace
    .trim();
}

function normalizeHeader(s) {
  return String(s || '').toLowerCase().replace(/[\s'’]/g, '').trim();
}

function deriveParentId(email, name) {
  var seed = String(email || '').toLowerCase().trim();
  if (!seed) seed = 'name:' + normalizeName(name);
  return 'p_' + sha1hex(seed).substring(0, 10);
}

function deriveStudentId(parentId, studentName) {
  return 's_' + sha1hex(parentId + '|' + normalizeName(studentName)).substring(0, 10);
}

function findSheetByGid(spreadsheet, gid) {
  var sheets = spreadsheet.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === gid) return sheets[i];
  }
  return null;
}

// ─── Facturas loader ─────────────────────────────────────────────────────────
//
// Reads the intake sheet, dedups parents by lowercased email (first row wins —
// same rule as the invoice generator), and emits one student per row.
//
// Returns: { parents: [{id, name, email}], students: [{id, parentId, name}] }

function loadFacturasParentsStudents() {
  var ss = SpreadsheetApp.openById(FACTURAS_SPREADSHEET_ID);
  var sheet = findSheetByGid(ss, FACTURAS_PARENTS_TAB_GID);
  if (!sheet) throw new Error('Facturas tab with gid ' + FACTURAS_PARENTS_TAB_GID + ' not found');

  var rows = sheet.getDataRange().getValues();
  if (rows.length < 2) throw new Error('Facturas sheet is empty');

  // Header row isn't always row 0 (there's a pricing table at the top).
  // Scan until we find one containing the parent name column.
  var headerIdx = -1;
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].some(function(c) { return normalizeHeader(c) === normalizeHeader(COL_PARENT_NAME); })) {
      headerIdx = i;
      break;
    }
  }
  if (headerIdx === -1) throw new Error('Could not find header row in Facturas (no "' + COL_PARENT_NAME + '")');

  var headers = rows[headerIdx].map(normalizeHeader);
  var col = {
    student: headers.indexOf(normalizeHeader(COL_STUDENT_NAME)),
    parent:  headers.indexOf(normalizeHeader(COL_PARENT_NAME)),
    email:   headers.indexOf(normalizeHeader(COL_PARENT_EMAIL))
  };
  var missing = [];
  Object.keys(col).forEach(function(k) { if (col[k] === -1) missing.push(k); });
  if (missing.length) throw new Error('Facturas missing columns: ' + missing.join(', '));

  // Active column is optional. -1 means "not present" → every row is active.
  var activeIdx = headers.indexOf(normalizeHeader(COL_ACTIVE));

  var parentsById = {};
  var students = [];

  for (var r = headerIdx + 1; r < rows.length; r++) {
    var row = rows[r];
    var studentName = String(row[col.student] || '').trim();
    var parentName  = String(row[col.parent]  || '').trim();
    var parentEmail = String(row[col.email]   || '').trim().toLowerCase();

    if (!studentName && !parentName) continue; // skip blank rows

    // Skip rows explicitly marked inactive. Blank / missing column = active.
    if (activeIdx !== -1) {
      var active = String(row[activeIdx] || '').trim().toUpperCase();
      if (active === 'NO') continue;
    }

    var parentId = deriveParentId(parentEmail, parentName);
    if (!parentsById[parentId]) {
      parentsById[parentId] = {
        id: parentId,
        name: parentName,
        email: parentEmail
      };
    }

    if (studentName) {
      var studentId = deriveStudentId(parentId, studentName);
      // Dedup duplicate rows for the same student (e.g., re-submissions)
      if (!students.some(function(s) { return s.id === studentId; })) {
        students.push({
          id: studentId,
          parentId: parentId,
          name: studentName
        });
      }
    }
  }

  var parents = Object.keys(parentsById).map(function(k) { return parentsById[k]; });
  parents.sort(function(a, b) { return a.name.localeCompare(b.name); });
  return { parents: parents, students: students };
}

// ─── GET operations ──────────────────────────────────────────────────────────

function getAllData() {
  var fs = loadFacturasParentsStudents();
  var overrides = loadOverrides();
  var students = fs.students.map(function(s) {
    return {
      id: s.id,
      parentId: s.parentId,
      name: s.name,
      priceOverrides: overrides[s.id] || {}
    };
  });
  return {
    parents: fs.parents.map(function(p) { return { id: p.id, name: p.name }; }),
    students: students,
    lessons: getLessons(),
    config: getConfig()
  };
}

function getParents() {
  var fs = loadFacturasParentsStudents();
  return fs.parents.map(function(p) { return { id: p.id, name: p.name }; });
}

function getStudents() {
  var fs = loadFacturasParentsStudents();
  var overrides = loadOverrides();
  return fs.students.map(function(s) {
    return {
      id: s.id,
      parentId: s.parentId,
      name: s.name,
      priceOverrides: overrides[s.id] || {}
    };
  });
}

function loadOverrides() {
  var sheet = getSheet('Students');
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  var out = {};
  if (data.length < 2) return out;
  // Expected columns: id, parentId, name, override_ind_pres, override_ind_online, override_grp_pres, override_grp_online
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][0] || '');
    if (!id) continue;
    var ov = {};
    if (data[i][3] !== '' && data[i][3] != null) ov.ind_pres   = Number(data[i][3]);
    if (data[i][4] !== '' && data[i][4] != null) ov.ind_online = Number(data[i][4]);
    if (data[i][5] !== '' && data[i][5] != null) ov.grp_pres   = Number(data[i][5]);
    if (data[i][6] !== '' && data[i][6] != null) ov.grp_online = Number(data[i][6]);
    if (Object.keys(ov).length) out[id] = ov;
  }
  return out;
}

function getLessons(month) {
  var sheet = getSheet('Lessons');
  var rawData = sheet.getDataRange().getValues();
  if (rawData.length < 2) return [];
  var headers = rawData[0];
  var all = [];
  for (var i = 1; i < rawData.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[String(headers[j]).trim()] = rawData[i][j];
    }

    // Handle date — may be Date object or string
    var dateStr = '';
    if (obj.date instanceof Date) {
      var d = obj.date;
      dateStr = d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
    } else {
      dateStr = String(obj.date || '').substring(0, 10);
      if (dateStr.match(/^\d{4}-\d{1,2}-\d{1,2}$/)) {
        var p = dateStr.split('-');
        dateStr = p[0] + '-' + p[1].padStart(2,'0') + '-' + p[2].padStart(2,'0');
      }
    }

    // Handle time — may be Date object (time-only) or string
    var timeStr = '';
    if (obj.time instanceof Date) {
      var t = obj.time;
      timeStr = String(t.getHours()).padStart(2,'0') + ':' + String(t.getMinutes()).padStart(2,'0');
    } else if (obj.time) {
      timeStr = String(obj.time);
    }

    all.push({
      id: String(obj.id),
      studentId: String(obj.studentId),
      parentId: String(obj.parentId),
      date: dateStr,
      time: timeStr,
      lessonType: String(obj.lessonType || ''),
      duration: Number(obj.duration) || 0,
      pricePerHour: Number(obj.pricePerHour) || 0,
      total: Number(obj.total) || 0
    });
  }
  if (month) {
    all = all.filter(function(l) {
      return l.date.substring(0, 7) === month;
    });
  }
  return all;
}

function getConfig() {
  var rows = sheetToObjects(getSheet('Config'));
  var config = {};
  rows.forEach(function(r) {
    config[r.key] = r.value;
  });
  return config;
}

// ─── POST operations: Students (pricing overrides only) ──────────────────────
//
// The student NAME is now owned by Facturas — we don't allow renaming here.
// This action only updates pricing overrides for a student whose ID was
// derived from Facturas. If no row exists in the Students sheet yet, one is
// appended.

function updateStudent(data) {
  var id = String(data.id || '');
  if (!id) return { error: 'No student id' };

  // Cross-check: the student must exist in Facturas. Otherwise we'd be
  // saving overrides for a ghost.
  var fs = loadFacturasParentsStudents();
  var facturasMatch = fs.students.filter(function(s) { return s.id === id; })[0];
  if (!facturasMatch) {
    return { error: 'Student id not found in Facturas — overrides must be attached to a Facturas-derived student.' };
  }

  var sheet = getSheet('Students');
  var ov = data.priceOverrides || {};
  var rowIdx = findRowById(sheet, id);

  if (rowIdx === -1) {
    sheet.appendRow([
      id,
      facturasMatch.parentId,
      facturasMatch.name,
      ov.ind_pres   == null || ov.ind_pres   === '' ? '' : Number(ov.ind_pres),
      ov.ind_online == null || ov.ind_online === '' ? '' : Number(ov.ind_online),
      ov.grp_pres   == null || ov.grp_pres   === '' ? '' : Number(ov.grp_pres),
      ov.grp_online == null || ov.grp_online === '' ? '' : Number(ov.grp_online)
    ]);
  } else {
    sheet.getRange(rowIdx, 2).setValue(facturasMatch.parentId);
    sheet.getRange(rowIdx, 3).setValue(facturasMatch.name);
    sheet.getRange(rowIdx, 4).setValue(ov.ind_pres   == null || ov.ind_pres   === '' ? '' : Number(ov.ind_pres));
    sheet.getRange(rowIdx, 5).setValue(ov.ind_online == null || ov.ind_online === '' ? '' : Number(ov.ind_online));
    sheet.getRange(rowIdx, 6).setValue(ov.grp_pres   == null || ov.grp_pres   === '' ? '' : Number(ov.grp_pres));
    sheet.getRange(rowIdx, 7).setValue(ov.grp_online == null || ov.grp_online === '' ? '' : Number(ov.grp_online));
  }
  return { success: true };
}

// ─── POST operations: Lessons ────────────────────────────────────────────────

function addLesson(data) {
  var sheet = getSheet('Lessons');
  var id = generateId();
  var row = [
    id,
    String(data.studentId || ''),
    String(data.parentId || ''),
    String(data.date || ''),
    String(data.time || ''),
    String(data.lessonType || ''),
    Number(data.duration) || 0,
    Number(data.pricePerHour) || 0,
    Number(data.total) || 0
  ];
  sheet.appendRow(row);
  return { success: true, id: id };
}

function updateLesson(data) {
  var sheet = getSheet('Lessons');
  var row = findRowById(sheet, data.id);
  if (row < 0) return { error: 'Lesson not found' };
  // Columns: id, studentId, parentId, date, time, lessonType, duration, pricePerHour, total
  sheet.getRange(row, 4).setValue(data.date);
  sheet.getRange(row, 5).setValue(data.time || '');
  sheet.getRange(row, 6).setValue(data.lessonType);
  sheet.getRange(row, 7).setValue(data.duration);
  sheet.getRange(row, 8).setValue(data.pricePerHour);
  sheet.getRange(row, 9).setValue(data.total);
  return { success: true };
}

function deleteLesson(id) {
  var sheet = getSheet('Lessons');
  var row = findRowById(sheet, id);
  if (row < 0) return { error: 'Lesson not found' };
  sheet.deleteRow(row);
  return { success: true };
}

// ─── POST operations: Config ─────────────────────────────────────────────────

function updateConfig(config) {
  var sheet = getSheet('Config');
  var data = sheet.getDataRange().getValues();

  for (var key in config) {
    var found = false;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(config[key]);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([key, config[key]]);
    }
  }
  return { success: true };
}

// ─── Migration: old local IDs → Facturas-derived IDs ─────────────────────────
//
// Run migrateDryRun first, inspect the execution log, then run migrateCommit
// if the diagnostics look right. Commit: remaps the Lessons sheet's
// parentId/studentId and rebuilds the Students sheet with new IDs + preserved
// overrides. The old Parents sheet is left in place as an archive.

function migrateDryRun() {
  var result = _migrate(true);
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function migrateCommit() {
  var result = _migrate(false);
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function _migrate(dryRun) {
  // Load old local data
  var oldParents = sheetToObjects(getSheet('Parents'));
  var studentsSheet = getSheet('Students');
  var oldStudentsRaw = studentsSheet.getDataRange().getValues();
  var oldStudents = [];
  for (var i = 1; i < oldStudentsRaw.length; i++) {
    var r = oldStudentsRaw[i];
    if (!r[0]) continue;
    oldStudents.push({
      id: String(r[0]),
      parentId: String(r[1] || ''),
      name: String(r[2] || ''),
      ov_ind_pres: r[3],
      ov_ind_online: r[4],
      ov_grp_pres: r[5],
      ov_grp_online: r[6]
    });
  }

  // Load Facturas
  var fs = loadFacturasParentsStudents();

  // Index Facturas parents by normalized name (and by email when available).
  var byNormalizedParentName = {};
  var byParentEmail = {};
  fs.parents.forEach(function(p) {
    byNormalizedParentName[normalizeName(p.name)] = p.id;
    if (p.email) byParentEmail[p.email.toLowerCase()] = p.id;
  });

  // Index Facturas students by (newParentId, normalized name)
  var byParentAndStudent = {};
  fs.students.forEach(function(s) {
    byParentAndStudent[s.parentId + '|' + normalizeName(s.name)] = s;
  });

  // Match old parents → new
  var oldParentIdToNew = {};
  var unmatchedParents = [];
  oldParents.forEach(function(op) {
    var newId = byNormalizedParentName[normalizeName(op.name)];
    if (newId) oldParentIdToNew[String(op.id)] = newId;
    else unmatchedParents.push({ oldId: String(op.id), name: String(op.name) });
  });

  // Match old students → new (idempotent: handles the case where a re-run
  // sees rows that are already in the Facturas ID format)
  var facturasParentIdSet = {};
  fs.parents.forEach(function(p) { facturasParentIdSet[p.id] = true; });

  var oldStudentIdToNew = {};
  var unmatchedStudents = [];
  var rebuiltOverrides = [];
  oldStudents.forEach(function(os) {
    // Resolve parent: either already Facturas-style, or remap from old id.
    var newParentId;
    if (/^p_[0-9a-f]{10}$/.test(os.parentId) && facturasParentIdSet[os.parentId]) {
      newParentId = os.parentId;
    } else {
      newParentId = oldParentIdToNew[os.parentId];
    }
    if (!newParentId) {
      unmatchedStudents.push({ oldId: os.id, name: os.name, reason: 'parent not matched' });
      return;
    }
    var match = byParentAndStudent[newParentId + '|' + normalizeName(os.name)];
    if (!match) {
      unmatchedStudents.push({ oldId: os.id, name: os.name, reason: 'no matching student under parent' });
      return;
    }
    oldStudentIdToNew[os.id] = match.id;
    var hasOv = (os.ov_ind_pres !== '' && os.ov_ind_pres != null) ||
                (os.ov_ind_online !== '' && os.ov_ind_online != null) ||
                (os.ov_grp_pres !== '' && os.ov_grp_pres != null) ||
                (os.ov_grp_online !== '' && os.ov_grp_online != null);
    if (hasOv) {
      rebuiltOverrides.push([
        match.id, match.parentId, match.name,
        os.ov_ind_pres   !== '' && os.ov_ind_pres   != null ? Number(os.ov_ind_pres)   : '',
        os.ov_ind_online !== '' && os.ov_ind_online != null ? Number(os.ov_ind_online) : '',
        os.ov_grp_pres   !== '' && os.ov_grp_pres   != null ? Number(os.ov_grp_pres)   : '',
        os.ov_grp_online !== '' && os.ov_grp_online != null ? Number(os.ov_grp_online) : ''
      ]);
    }
  });

  // Inspect Lessons sheet (read-only for dry run)
  var lessonsSheet = getSheet('Lessons');
  var lessonsData = lessonsSheet.getDataRange().getValues();
  var lheaders = lessonsData[0];
  var sIdx = -1, pIdx = -1;
  for (var h = 0; h < lheaders.length; h++) {
    if (String(lheaders[h]).trim() === 'studentId') sIdx = h;
    if (String(lheaders[h]).trim() === 'parentId')  pIdx = h;
  }
  if (sIdx === -1 || pIdx === -1) {
    return { error: 'Lessons sheet missing studentId or parentId header' };
  }

  var remappedLessons = 0;
  var unremappedLessons = 0;
  var unremappedSample = [];
  for (var rr = 1; rr < lessonsData.length; rr++) {
    var oldSid = String(lessonsData[rr][sIdx] || '');
    var oldPid = String(lessonsData[rr][pIdx] || '');
    if (!oldSid && !oldPid) { unremappedLessons++; continue; }
    // Already a Facturas-style ID? (starts with s_ / p_ and 12 chars long.)
    var alreadyNew = /^s_[0-9a-f]{10}$/.test(oldSid) && /^p_[0-9a-f]{10}$/.test(oldPid);
    if (alreadyNew) { continue; }
    var newSid = oldStudentIdToNew[oldSid];
    var newPid = oldParentIdToNew[oldPid];
    if (newSid && newPid) {
      lessonsData[rr][sIdx] = newSid;
      lessonsData[rr][pIdx] = newPid;
      remappedLessons++;
    } else {
      unremappedLessons++;
      if (unremappedSample.length < 10) {
        unremappedSample.push({
          row: rr + 1,
          oldStudentId: oldSid,
          oldParentId: oldPid,
          date: String(lessonsData[rr][lheaders.indexOf('date')] || '')
        });
      }
    }
  }

  var diagnostics = {
    dryRun: dryRun,
    facturasParents: fs.parents.length,
    facturasStudents: fs.students.length,
    oldParents: oldParents.length,
    matchedParents: Object.keys(oldParentIdToNew).length,
    unmatchedParents: unmatchedParents,
    oldStudents: oldStudents.length,
    matchedStudents: Object.keys(oldStudentIdToNew).length,
    unmatchedStudents: unmatchedStudents,
    rebuiltOverridesRows: rebuiltOverrides.length,
    lessonsTotal: lessonsData.length - 1,
    remappedLessons: remappedLessons,
    unremappedLessons: unremappedLessons,
    unremappedSample: unremappedSample
  };

  if (dryRun) {
    return diagnostics;
  }

  // ─── Commit ────────────────────────────────────────────────────────────────

  // 1. Rewrite Lessons sheet with remapped IDs
  if (lessonsData.length > 1) {
    lessonsSheet.getRange(1, 1, lessonsData.length, lessonsData[0].length).setValues(lessonsData);
  }

  // 2. Rebuild Students sheet — keep header, write only rows that have overrides
  studentsSheet.clear();
  studentsSheet.appendRow([
    'id', 'parentId', 'name',
    'override_ind_pres', 'override_ind_online', 'override_grp_pres', 'override_grp_online'
  ]);
  if (rebuiltOverrides.length > 0) {
    studentsSheet.getRange(2, 1, rebuiltOverrides.length, 7).setValues(rebuiltOverrides);
  }

  diagnostics.success = true;
  return diagnostics;
}
