/**
 * Lesson Logger — Google Apps Script API
 * 
 * Paste this into Extensions → Apps Script in your Google Sheet.
 * Deploy as Web App (Execute as: Me, Access: Anyone).
 */

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
      // Write operations
      case 'addParent':
        result = addParent(data);
        break;
      case 'updateParent':
        result = updateParent(data);
        break;
      case 'deleteParent':
        result = deleteParent(data.id);
        break;
      case 'addStudent':
        result = addStudent(data);
        break;
      case 'updateStudent':
        result = updateStudent(data);
        break;
      case 'deleteStudent':
        result = deleteStudent(data.id);
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
      case 'bulkImportParents':
        result = bulkImportParents(data.parents);
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

// ─── Helpers ─────────────────────────────────────────────────────────────────

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
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
    if (data[i][0] === id) return i + 1; // 1-indexed row
  }
  return -1;
}

function generateId() {
  // Prefix with 'x' to prevent Google Sheets from interpreting as scientific notation
  return 'x' + Utilities.getUuid().replace(/-/g, '').substring(0, 7);
}

// ─── GET operations ──────────────────────────────────────────────────────────

function getAllData() {
  return {
    parents: getParents(),
    students: getStudents(),
    lessons: getLessons(),
    config: getConfig()
  };
}

function getParents() {
  return sheetToObjects(getSheet('Parents'));
}

function getStudents() {
  var raw = sheetToObjects(getSheet('Students'));
  return raw.map(function(s) {
    var overrides = {};
    if (s.override_ind_pres) overrides.ind_pres = Number(s.override_ind_pres);
    if (s.override_ind_online) overrides.ind_online = Number(s.override_ind_online);
    if (s.override_grp_pres) overrides.grp_pres = Number(s.override_grp_pres);
    if (s.override_grp_online) overrides.grp_online = Number(s.override_grp_online);
    return {
      id: s.id,
      parentId: s.parentId,
      name: s.name,
      priceOverrides: overrides
    };
  });
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

// ─── POST operations: Parents ────────────────────────────────────────────────

function addParent(data) {
  var sheet = getSheet('Parents');
  var id = generateId();
  sheet.appendRow([id, data.name || '']);
  return { success: true, id: id };
}

function updateParent(data) {
  var sheet = getSheet('Parents');
  var row = findRowById(sheet, data.id);
  if (row < 0) return { error: 'Parent not found' };
  sheet.getRange(row, 2).setValue(data.name || '');
  return { success: true };
}

function deleteParent(id) {
  var sheet = getSheet('Parents');
  var row = findRowById(sheet, id);
  if (row < 0) return { error: 'Parent not found' };
  sheet.deleteRow(row);
  // Also delete their students
  var studSheet = getSheet('Students');
  var studData = studSheet.getDataRange().getValues();
  for (var i = studData.length - 1; i >= 1; i--) {
    if (studData[i][1] === id) studSheet.deleteRow(i + 1);
  }
  return { success: true };
}

function bulkImportParents(parents) {
  var sheet = getSheet('Parents');
  var studSheet = getSheet('Students');
  
  // Build all rows first, then write in one batch (much faster than appendRow)
  var parentRows = [];
  var studentRows = [];
  var count = 0;
  
  parents.forEach(function(p) {
    var id = generateId();
    parentRows.push([id, p.name || '']);
    count++;
    if (p.students && p.students.length > 0) {
      p.students.forEach(function(s) {
        var sid = generateId();
        studentRows.push([sid, id, s.name || '', '', '', '', '']);
      });
    }
  });
  
  // Batch write parents
  if (parentRows.length > 0) {
    var lastRow = sheet.getLastRow();
    var range = sheet.getRange(lastRow + 1, 1, parentRows.length, 2);
    range.setNumberFormat('@'); // Force plain text
    range.setValues(parentRows);
  }
  
  // Batch write students
  if (studentRows.length > 0) {
    var lastStudRow = studSheet.getLastRow();
    var range2 = studSheet.getRange(lastStudRow + 1, 1, studentRows.length, 7);
    range2.setNumberFormat('@'); // Force plain text
    range2.setValues(studentRows);
  }
  
  return { success: true, count: count };
}

// ─── POST operations: Students ───────────────────────────────────────────────

function addStudent(data) {
  var sheet = getSheet('Students');
  var id = generateId();
  var ov = data.priceOverrides || {};
  sheet.appendRow([
    id,
    data.parentId,
    data.name || '',
    ov.ind_pres || '',
    ov.ind_online || '',
    ov.grp_pres || '',
    ov.grp_online || ''
  ]);
  return { success: true, id: id };
}

function updateStudent(data) {
  var sheet = getSheet('Students');
  var row = findRowById(sheet, data.id);
  if (row < 0) return { error: 'Student not found' };
  var ov = data.priceOverrides || {};
  sheet.getRange(row, 3).setValue(data.name || '');
  sheet.getRange(row, 4).setValue(ov.ind_pres || '');
  sheet.getRange(row, 5).setValue(ov.ind_online || '');
  sheet.getRange(row, 6).setValue(ov.grp_pres || '');
  sheet.getRange(row, 7).setValue(ov.grp_online || '');
  return { success: true };
}

function deleteStudent(id) {
  var sheet = getSheet('Students');
  var row = findRowById(sheet, id);
  if (row < 0) return { error: 'Student not found' };
  sheet.deleteRow(row);
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
