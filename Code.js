// ---------- CONFIG ----------
// These sheet names match the live spreadsheet this script is bound to.
// If you rename a tab in the Sheet, update the matching constant below.
const STUDENT_SHEET_NAME = 'Student List';
const CI_SHEET_NAME = 'CI List';
const INCIDENT_SHEET_NAME = 'Incident';
const EXTENSION_SHEET_NAME = 'Extension';
const FULL_REPORT_SHEET_NAME = 'Full Report';
const CI_LOGIN_LOG_SHEET_NAME = 'CI Login Log';
const ADMIN_SHEET_NAME = 'Admin';
// -----------------------------------

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('COC Phinma College of Nursing')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// JSON API for the standalone (GitHub Pages) frontend — same server functions the
// Apps Script-hosted UI calls via google.script.run, exposed over HTTP instead so a
// static site can reach them with fetch(). Every action name here must match an
// actual function below exactly; args are applied positionally.
function doPost(e) {
  try {
    const body = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
    const action = body.action;
    const fn = API_ACTIONS[action];
    if (!fn) return _jsonResponse({ status: 'error', message: 'Unknown action: ' + action });
    const result = fn.apply(null, Array.isArray(body.args) ? body.args : []);
    return _jsonResponse(result);
  } catch (err) {
    return _jsonResponse({ status: 'error', message: err.message });
  }
}

function _jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function openActiveSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('This script must be bound to a Google Sheet.');
  return ss;
}

function _normalizeHeader(h) { return String(h || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim(); }

function _sheetObjects(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { headers: [], rows: [] };
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(_normalizeHeader);
  const rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues().map(r => {
    const o = {};
    for (let c = 0; c < headers.length; c++) o[headers[c] || ('col' + (c + 1))] = r[c];
    return o;
  });
  return { headers, rows };
}

// Writes a row into `sheetName`, matching values to whatever the sheet's actual
// header order is (rather than assuming a fixed column order). Creates the sheet
// with `defaultHeaders` if it doesn't exist yet.
function _writeRow(ss, sheetName, defaultHeaders, fields) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(defaultHeaders);
  }
  let lastCol = Math.max(sheet.getLastColumn(), defaultHeaders.length);
  let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(_normalizeHeader);
  const row = new Array(lastCol).fill('');
  fields.forEach(f => {
    let idx = headers.findIndex(h => f.patterns.some(p => p.test(h)));
    if (idx === -1) {
      idx = headers.findIndex(h => !h);
      if (idx === -1) {
        idx = headers.length;
        headers.push(_normalizeHeader(f.fallbackHeader || ''));
        row.push('');
        sheet.getRange(1, idx + 1).setValue(f.fallbackHeader || '');
      } else {
        sheet.getRange(1, idx + 1).setValue(f.fallbackHeader || '');
        headers[idx] = _normalizeHeader(f.fallbackHeader || '');
      }
    }
    row[idx] = f.value;
  });
  sheet.appendRow(row);
}

/* ================= LOOKUP LISTS ================= */

// CI List sheet layout is fixed: column A = CI name, column B = password.
function _ciSheetRows() {
  const ss = openActiveSpreadsheet();
  const sheet = ss.getSheetByName(CI_SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + CI_SHEET_NAME + '" not found.');
  const lastRow = sheet.getLastRow();
  const rows = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 2).getValues() : [];
  return { sheet, rows, nameCol: 0, passCol: 1 };
}

function getCIs() {
  try {
    const { rows } = _ciSheetRows();
    const seen = {}; const list = [];
    rows.forEach(r => {
      const name = String(r[0] || '').trim();
      if (!name) return;
      const key = name.toLowerCase();
      if (seen[key]) return;
      seen[key] = true;
      list.push({ name });
    });
    list.sort((a, b) => a.name.localeCompare(b.name, 'en', { sensitivity: 'base' }));
    return { status: 'success', list };
  } catch (e) {
    return { status: 'error', list: [], message: e.message };
  }
}

function getStudents() {
  try {
    const ss = openActiveSpreadsheet();
    const sheet = ss.getSheetByName(STUDENT_SHEET_NAME);
    if (!sheet) return { status: 'error', list: [], message: 'Sheet "' + STUDENT_SHEET_NAME + '" not found.' };
    const { headers, rows } = _sheetObjects(sheet);
    const nameKey = headers.find(h => /(^| )name($| )|student name|full name/.test(h)) || headers[1] || headers[0];
    const idKey = headers.find(h => /(^| )id($| )|student id|id number/.test(h)) || headers[0];
    const locKey = headers.find(h => /location|area|site/.test(h));
    const list = rows.map(r => ({
      name: String(r[nameKey] || '').trim(),
      id: String(r[idKey] || '').trim(),
      location: locKey ? String(r[locKey] || '').trim() : ''
    })).filter(x => x.name);
    list.sort((a, b) => a.name.localeCompare(b.name, 'en', { sensitivity: 'base' }));
    return { status: 'success', list };
  } catch (e) {
    return { status: 'error', list: [], message: e.message };
  }
}

function getIncidents() {
  try {
    const ss = openActiveSpreadsheet();
    const sheet = ss.getSheetByName(INCIDENT_SHEET_NAME);
    if (!sheet) return { status: 'error', list: [], message: 'Sheet "' + INCIDENT_SHEET_NAME + '" not found.' };
    const { headers, rows } = _sheetObjects(sheet);
    const nameKey = headers[0];
    const hoursKey = headers.find(h => /(^| )hours($| )|number of hours|duration|hrs|hour/.test(h)) || headers[1];
    const list = [];
    rows.forEach(r => {
      const name = String(r[nameKey] || '').trim();
      if (!name) return;
      list.push({ incident: name, hours: hoursKey ? String(r[hoursKey] || '').trim() : '' });
    });
    return { status: 'success', list };
  } catch (e) {
    return { status: 'error', list: [], message: e.message };
  }
}

/* ================= AUTH ================= */

function authenticateCI(name, password) {
  try {
    name = (name || '').toString().trim();
    password = (password || '').toString().trim();
    if (!name) return { status: 'error', message: 'Select your name.' };
    const { rows } = _ciSheetRows();
    const match = rows.find(r => String(r[0] || '').trim().toLowerCase() === name.toLowerCase());
    if (!match) return { status: 'error', message: 'Clinical Instructor not found.' };
    const realPass = String(match[1] || '').trim();
    if (realPass && realPass !== password) return { status: 'error', message: 'Incorrect password.' };
    return { status: 'success', name: String(match[0]).trim() };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

// Called by the CI login form (as opposed to the internal _requireCI re-checks used by
// bulk add/remove) so a successful login is traced exactly once per login attempt.
function loginCI(name, password) {
  const result = authenticateCI(name, password);
  if (result.status === 'success') {
    try {
      const ss = openActiveSpreadsheet();
      let sheet = ss.getSheetByName(CI_LOGIN_LOG_SHEET_NAME);
      if (!sheet) {
        sheet = ss.insertSheet(CI_LOGIN_LOG_SHEET_NAME);
        sheet.appendRow(['TIMESTAMP', 'CI NAME']);
      }
      sheet.appendRow([new Date(), result.name]);
    } catch (e) {
      // Logging failure shouldn't block the CI from logging in.
    }
  }
  return result;
}

function changeCIPassword(name, oldPassword, newPassword) {
  try {
    const auth = authenticateCI(name, oldPassword);
    if (auth.status !== 'success') return auth;
    newPassword = (newPassword || '').toString().trim();
    if (!newPassword) return { status: 'error', message: 'Enter a new password.' };

    const { sheet, rows } = _ciSheetRows();
    const rowIdx = rows.findIndex(r => String(r[0] || '').trim().toLowerCase() === auth.name.toLowerCase());
    if (rowIdx === -1) return { status: 'error', message: 'Clinical Instructor not found.' };

    sheet.getRange(rowIdx + 2, 2).setValue(newPassword); // column B
    return { status: 'success' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

// Admin sheet layout is fixed: column A = username, column B = password.
// Auto-creates the sheet with just the header row if missing — an admin account
// must be added manually in the sheet before anyone can log in.
function _adminSheet() {
  const ss = openActiveSpreadsheet();
  let sheet = ss.getSheetByName(ADMIN_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(ADMIN_SHEET_NAME);
    sheet.appendRow(['USERNAME', 'PASSWORD']);
  }
  return sheet;
}

function getAdmins() {
  try {
    const sheet = _adminSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: 'success', list: [] };
    const names = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat()
      .map(v => String(v || '').trim()).filter(Boolean);
    names.sort((a, b) => a.localeCompare(b, 'en', { sensitivity: 'base' }));
    return { status: 'success', list: names };
  } catch (e) {
    return { status: 'error', list: [], message: e.message };
  }
}

function authenticateAdmin(username, password) {
  try {
    username = (username || '').toString().trim();
    password = (password || '').toString().trim();
    if (!username) return { status: 'error', message: 'Enter your admin username.' };
    const sheet = _adminSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: 'error', message: 'Admin account not found.' };
    const rows = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    const match = rows.find(r => String(r[0] || '').trim().toLowerCase() === username.toLowerCase());
    if (!match) return { status: 'error', message: 'Admin account not found.' };
    const realPass = String(match[1] || '').trim();
    if (realPass && realPass !== password) return { status: 'error', message: 'Incorrect password.' };
    return { status: 'success', username: String(match[0]).trim() };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function _requireAdmin(username, password) {
  const auth = authenticateAdmin(username, password);
  if (auth.status !== 'success') throw new Error(auth.message || 'Not authorized.');
  return auth.username;
}

function getStudentProfile(studentId) {
  try {
    studentId = (studentId || '').toString().trim();
    if (!studentId) return { status: 'error', message: 'Enter a Student ID.' };
    const ss = openActiveSpreadsheet();
    const sheet = ss.getSheetByName(STUDENT_SHEET_NAME);
    if (!sheet) return { status: 'error', message: 'Sheet "' + STUDENT_SHEET_NAME + '" not found.' };
    const { headers, rows } = _sheetObjects(sheet);
    const nameKey = headers.find(h => /(^| )name($| )|student name|full name/.test(h)) || headers[1] || headers[0];
    const idKey = headers.find(h => /(^| )id($| )|student id|id number/.test(h)) || headers[0];
    const locKey = headers.find(h => /location|area|site/.test(h));
    const match = rows.find(r => String(r[idKey] || '').trim() === studentId);
    if (match) {
      return {
        status: 'success',
        id: studentId,
        name: String(match[nameKey] || '').trim(),
        location: locKey ? String(match[locKey] || '').trim() : ''
      };
    }

    // Not on the Student List roster — fall back to the Extension log so records
    // for a student aren't hidden just because they haven't been added to the roster yet.
    const extSheet = ss.getSheetByName(EXTENSION_SHEET_NAME);
    if (extSheet) {
      const ext = _sheetObjects(extSheet);
      const extIdKey = ext.headers.find(h => /student id/.test(h));
      const extNameKey = ext.headers.find(h => /student name/.test(h));
      const extLocKey = ext.headers.find(h => /location/.test(h));
      const extMatch = extIdKey ? ext.rows.find(r => String(r[extIdKey] || '').trim() === studentId) : null;
      if (extMatch) {
        return {
          status: 'success',
          id: studentId,
          name: extNameKey ? String(extMatch[extNameKey] || '').trim() : '',
          location: extLocKey ? String(extMatch[extLocKey] || '').trim() : '',
          notOnRoster: true
        };
      }
    }
    return { status: 'error', message: 'Student ID not found.' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

/* ================= EXTENSION / INCIDENT LOG ================= */

function getFullReportColumnCByStudentId(studentId) {
  try {
    studentId = (studentId || '').toString().trim();
    if (!studentId) return { status: 'success', total: 0, rawValues: [] };
    const ss = openActiveSpreadsheet();
    const sheet = ss.getSheetByName(FULL_REPORT_SHEET_NAME);
    if (!sheet) return { status: 'error', total: 0, rawValues: [], message: 'Sheet "' + FULL_REPORT_SHEET_NAME + '" not found.' };
    const lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return { status: 'success', total: 0, rawValues: [] };
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const rawValues = []; let total = 0;
    for (let r = 0; r < data.length; r++) {
      const row = data[r];
      const matched = row.some(cell => cell !== undefined && cell !== null && String(cell).trim() === studentId);
      if (!matched) continue;
      const val = row.length >= 3 ? row[2] : '';
      const strVal = (val === undefined || val === null) ? '' : String(val).trim();
      rawValues.push(strVal);
      const num = parseFloat(strVal.replace(/,/g, ''));
      if (!isNaN(num)) total += num;
    }
    if (Math.abs(Math.round(total) - total) < 1e-9) total = Math.round(total);
    return { status: 'success', total, rawValues };
  } catch (e) {
    return { status: 'error', total: 0, rawValues: [], message: e.message };
  }
}

function submitEntry(data) {
  try {
    if (!data || !data.studentName) return { status: 'error', message: 'Student is required.' };
    const ss = openActiveSpreadsheet();
    const defaultHeaders = ['TIME', 'DATE', 'CI NAME', 'LOCATION', 'STUDENT NAME', 'STUDENT ID', 'REMARKS', 'HOURS', 'INCIDENT'];
    _writeRow(ss, EXTENSION_SHEET_NAME, defaultHeaders, [
      { patterns: [/^time$/, /timestamp/], value: new Date(), fallbackHeader: 'TIME' },
      { patterns: [/^date$/, /selected date/], value: data.selectedDatePacific || '', fallbackHeader: 'DATE' },
      { patterns: [/ci name/, /instructor/], value: data.ciName || '', fallbackHeader: 'CI NAME' },
      { patterns: [/student name/], value: data.studentName || '', fallbackHeader: 'STUDENT NAME' },
      { patterns: [/student id/], value: data.studentId || '', fallbackHeader: 'STUDENT ID' },
      { patterns: [/remarks/], value: data.remarks || '', fallbackHeader: 'REMARKS' },
      { patterns: [/hours/], value: data.numberOfHours || '', fallbackHeader: 'HOURS' },
      { patterns: [/incident/], value: data.incident || '', fallbackHeader: 'INCIDENT' }
    ]);
    return { status: 'success', message: 'Saved.' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

// Parses one pasted row in the order: DATE, CI NAME, LOCATION, STUDENT NAME, STUDENT ID,
// REMARKS, HOURS, INCIDENT. Tabs (native spreadsheet paste) are preferred; falls back to
// commas for typed-in text.
function _parseBulkExtensionRow(line) {
  const parts = (line.indexOf('\t') !== -1 ? line.split('\t') : line.split(',')).map(p => p.trim());
  return {
    date: parts[0] || '',
    ciName: parts[1] || '',
    location: parts[2] || '',
    studentName: parts[3] || '',
    studentId: parts[4] || '',
    remarks: parts[5] || '',
    hours: parts[6] || '',
    incident: parts[7] || ''
  };
}

function bulkAddExtensionEntries(ciName, ciPassword, rowsText) {
  try {
    _requireCI(ciName, ciPassword);
    const lines = String(rowsText || '').split(/\r?\n/).map(l => l.trim()).filter(Boolean);
    if (!lines.length) return { status: 'error', message: 'Paste at least one row.' };

    const ss = openActiveSpreadsheet();
    const defaultHeaders = ['TIME', 'DATE', 'CI NAME', 'LOCATION', 'STUDENT NAME', 'STUDENT ID', 'REMARKS', 'HOURS', 'INCIDENT'];
    let added = 0, skipped = 0;
    lines.forEach(line => {
      const row = _parseBulkExtensionRow(line);
      if (!row.date || !row.studentName || !row.studentId) { skipped++; return; }
      _writeRow(ss, EXTENSION_SHEET_NAME, defaultHeaders, [
        { patterns: [/^time$/, /timestamp/], value: new Date(), fallbackHeader: 'TIME' },
        { patterns: [/^date$/, /selected date/], value: row.date, fallbackHeader: 'DATE' },
        { patterns: [/ci name/, /instructor/], value: row.ciName, fallbackHeader: 'CI NAME' },
        { patterns: [/location/, /area/], value: row.location, fallbackHeader: 'LOCATION' },
        { patterns: [/student name/], value: row.studentName, fallbackHeader: 'STUDENT NAME' },
        { patterns: [/student id/], value: row.studentId, fallbackHeader: 'STUDENT ID' },
        { patterns: [/remarks/], value: row.remarks, fallbackHeader: 'REMARKS' },
        { patterns: [/hours/], value: row.hours, fallbackHeader: 'HOURS' },
        { patterns: [/incident/], value: row.incident, fallbackHeader: 'INCIDENT' }
      ]);
      added++;
    });
    return { status: 'success', added, skipped };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

// Sheets stores date-like values as real Date objects, which google.script.run serializes
// to a full JS toString() (e.g. "Tue Nov 11 2025 00:00:00 GMT+0800 ..."). Format those down
// to a plain "Nov 11 2025" for display; leave plain text values (e.g. a manually-typed date) as-is.
function _formatDate(ss, val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, ss.getSpreadsheetTimeZone(), 'MMM d, yyyy');
  }
  return (val === undefined || val === null) ? '' : String(val).trim();
}

function getStudentHistory(studentId) {
  try {
    studentId = (studentId || '').toString().trim();
    const ss = openActiveSpreadsheet();
    const sheet = ss.getSheetByName(EXTENSION_SHEET_NAME);
    if (!sheet) return { status: 'success', list: [] };
    const { headers, rows } = _sheetObjects(sheet);
    const idKey = headers.find(h => /student id/.test(h));
    const dateKey = headers.find(h => /^date$/.test(h)) || headers.find(h => /date/.test(h));
    const ciKey = headers.find(h => /ci name/.test(h));
    const locKey = headers.find(h => /location/.test(h));
    const incKey = headers.find(h => /incident/.test(h));
    const hoursKey = headers.find(h => /hours/.test(h));
    const remKey = headers.find(h => /remarks/.test(h));
    const list = rows
      .filter(r => idKey && String(r[idKey] || '').trim() === studentId)
      .map(r => ({
        date: dateKey ? _formatDate(ss, r[dateKey]) : '',
        ciName: ciKey ? String(r[ciKey] || '') : '',
        location: locKey ? String(r[locKey] || '') : '',
        incident: incKey ? String(r[incKey] || '') : '',
        hours: hoursKey ? String(r[hoursKey] || '') : '',
        remarks: remKey ? String(r[remKey] || '') : ''
      }));
    return { status: 'success', list };
  } catch (e) {
    return { status: 'error', list: [], message: e.message };
  }
}

function getCIHistory(ciName) {
  try {
    ciName = (ciName || '').toString().trim().toLowerCase();
    const ss = openActiveSpreadsheet();
    const sheet = ss.getSheetByName(EXTENSION_SHEET_NAME);
    if (!sheet) return { status: 'success', list: [], summary: { students: 0, totalHours: 0 } };
    const { headers, rows } = _sheetObjects(sheet);
    const ciKey = headers.find(h => /ci name/.test(h));
    const dateKey = headers.find(h => /date/.test(h));
    const studentKey = headers.find(h => /student name/.test(h));
    const idKey = headers.find(h => /student id/.test(h));
    const incKey = headers.find(h => /incident/.test(h));
    const hoursKey = headers.find(h => /hours/.test(h));
    const remKey = headers.find(h => /remarks/.test(h));
    const filtered = rows.filter(r => ciKey && String(r[ciKey] || '').trim().toLowerCase() === ciName);
    const list = filtered.map(r => ({
      date: dateKey ? _formatDate(ss, r[dateKey]) : '',
      studentName: studentKey ? String(r[studentKey] || '') : '',
      studentId: idKey ? String(r[idKey] || '') : '',
      incident: incKey ? String(r[incKey] || '') : '',
      hours: hoursKey ? String(r[hoursKey] || '') : '',
      remarks: remKey ? String(r[remKey] || '') : ''
    })).reverse();
    const uniqStudents = new Set(list.map(x => x.studentId || x.studentName));
    let totalHours = 0;
    list.forEach(x => { const n = parseFloat(String(x.hours).replace(/,/g, '')); if (!isNaN(n)) totalHours += n; });
    return { status: 'success', list, summary: { students: uniqStudents.size, totalHours } };
  } catch (e) {
    return { status: 'error', list: [], summary: { students: 0, totalHours: 0 }, message: e.message };
  }
}

/* ================= MANAGE STUDENTS (CI-only) ================= */

// Throws if the given CI credentials don't check out. Server-side re-check so this
// can't be bypassed by calling google.script.run functions directly from the console.
function _requireCI(ciName, ciPassword) {
  const auth = authenticateCI(ciName, ciPassword);
  if (auth.status !== 'success') throw new Error(auth.message || 'Not authorized.');
  return auth.name;
}

function _studentSheetInfo() {
  const ss = openActiveSpreadsheet();
  const sheet = ss.getSheetByName(STUDENT_SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + STUDENT_SHEET_NAME + '" not found.');
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(_normalizeHeader);
  // The real "Student List" sheet repeats the Student ID across two columns
  // (student id, Name, student id, LOCATION) — capture every matching column
  // so a pasted ID gets written into all of them, not just the first.
  const idPattern = /(^| )id($| )|student id|id number/;
  let idCols = headers.map((h, i) => idPattern.test(h) ? i : -1).filter(i => i !== -1);
  let nameCol = headers.findIndex(h => /(^| )name($| )|student name|full name/.test(h));
  const locCol = headers.findIndex(h => /location|area|site/.test(h));
  if (idCols.length === 0) idCols = [0];
  if (nameCol === -1) nameCol = 1;
  return { sheet, headers, idCol: idCols[0], idCols, nameCol, locCol };
}

// Handles both paste styles: rows copied straight from a spreadsheet (tab-separated)
// and rows typed as comma-separated text. Names are themselves "Last, First" and
// contain a comma (e.g. "RAMOS, WINZY VAN"), so a naive split(',') mis-parses them —
// tabs are preferred when present, and for commas the first part is the ID, the last
// part is the duplicate ID, and everything in between is rejoined as the name.
function _parseBulkStudentRow(line) {
  if (line.indexOf('\t') !== -1) {
    const parts = line.split('\t').map(p => p.trim());
    return { id: parts[0] || '', name: parts[1] || '' };
  }
  const parts = line.split(',').map(p => p.trim());
  if (parts.length <= 2) return { id: parts[0] || '', name: parts[1] || '' };
  return { id: parts[0], name: parts.slice(1, parts.length - 1).join(', ') };
}

// Normalizes an ID for comparison only (not for storage) so pasted data with different
// dash characters (‑, –, —) or stray whitespace still matches an existing plain-hyphen ID.
function _normalizeId(id) {
  return String(id || '').trim().replace(/[‐-―]/g, '-').replace(/\s+/g, '');
}

function bulkAddStudents(ciName, ciPassword, rowsText) {
  try {
    _requireCI(ciName, ciPassword);
    const info = _studentSheetInfo();
    const lines = String(rowsText || '').split(/\r?\n/).map(l => l.trim()).filter(Boolean);
    if (!lines.length) return { status: 'error', message: 'Paste at least one student row.' };

    const lastRow = info.sheet.getLastRow();
    const existing = lastRow >= 2 ? info.sheet.getRange(2, 1, lastRow - 1, info.headers.length).getValues() : [];
    // Check every ID column, not just the first — legacy rows can have the ID
    // recorded in only one of the two duplicate ID columns.
    const existingIds = new Set();
    existing.forEach(r => {
      info.idCols.forEach(c => {
        const id = String(r[c] || '').trim();
        if (id) existingIds.add(_normalizeId(id));
      });
    });

    // Duplicate Student IDs are rejected outright — both ones already on the roster
    // and repeats within the same pasted batch — rather than silently overwriting.
    let added = 0, duplicates = 0, skipped = 0;
    const seenInBatch = new Set();
    const duplicateIds = [];
    lines.forEach(line => {
      const { id, name } = _parseBulkStudentRow(line);
      if (!id || !name) { skipped++; return; }
      const idKey = _normalizeId(id);
      if (existingIds.has(idKey) || seenInBatch.has(idKey)) {
        duplicates++;
        duplicateIds.push(id);
        return;
      }
      seenInBatch.add(idKey);
      const rowVals = new Array(info.headers.length).fill('');
      info.idCols.forEach(c => { rowVals[c] = id; });
      rowVals[info.nameCol] = name;
      info.sheet.appendRow(rowVals);
      added++;
    });
    return { status: 'success', added, duplicates, skipped, duplicateIds };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function removeStudents(ciName, ciPassword, idsText) {
  try {
    _requireCI(ciName, ciPassword);
    const info = _studentSheetInfo();
    const ids = String(idsText || '').split(/[\r\n,]+/).map(s => s.trim()).filter(Boolean);
    if (!ids.length) return { status: 'error', message: 'Paste at least one Student ID.' };

    const lastRow = info.sheet.getLastRow();
    if (lastRow < 2) return { status: 'success', removed: 0, notFound: ids };
    const idSet = new Set(ids);
    const values = info.sheet.getRange(2, 1, lastRow - 1, info.headers.length).getValues();
    const foundIds = new Set();
    const rowsToDelete = [];
    values.forEach((r, i) => {
      const id = String(r[info.idCol] || '').trim();
      if (idSet.has(id)) { rowsToDelete.push(i + 2); foundIds.add(id); }
    });
    rowsToDelete.sort((a, b) => b - a).forEach(rowNum => info.sheet.deleteRow(rowNum));
    const notFound = ids.filter(id => !foundIds.has(id));
    return { status: 'success', removed: rowsToDelete.length, notFound };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

/* ================= MANAGE CLINICAL INSTRUCTORS (Admin-only) ================= */

// Fixed layout: column A = CI name, column B = password.
function _ciSheetInfo() {
  const ss = openActiveSpreadsheet();
  const sheet = ss.getSheetByName(CI_SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + CI_SHEET_NAME + '" not found.');
  const lastCol = Math.max(sheet.getLastColumn(), 2);
  return { sheet, headers: new Array(lastCol).fill(''), nameCol: 0, passCol: 1 };
}

function bulkAddCIs(adminUser, adminPass, rowsText) {
  try {
    _requireAdmin(adminUser, adminPass);
    const info = _ciSheetInfo();
    const lines = String(rowsText || '').split(/\r?\n/).map(l => l.trim()).filter(Boolean);
    if (!lines.length) return { status: 'error', message: 'Paste at least one CI row.' };

    const lastRow = info.sheet.getLastRow();
    const existing = lastRow >= 2 ? info.sheet.getRange(2, 1, lastRow - 1, info.headers.length).getValues() : [];
    const rowByName = {};
    existing.forEach((r, i) => {
      const name = String(r[info.nameCol] || '').trim().toLowerCase();
      if (name) rowByName[name] = i + 2;
    });

    let added = 0, updated = 0, skipped = 0;
    lines.forEach(line => {
      // Format: Name, Password — split on the LAST comma, since CI names are
      // themselves "Last, First" style and contain a comma (e.g. "ALCANZAR, FELIX").
      const idx = line.lastIndexOf(',');
      if (idx === -1) { skipped++; return; }
      const name = line.slice(0, idx).trim();
      const password = line.slice(idx + 1).trim();
      if (!name) { skipped++; return; }
      const key = name.toLowerCase();
      const rowVals = new Array(info.headers.length).fill('');
      rowVals[info.nameCol] = name;
      rowVals[info.passCol] = password;
      if (rowByName[key]) {
        info.sheet.getRange(rowByName[key], 1, 1, info.headers.length).setValues([rowVals]);
        updated++;
      } else {
        info.sheet.appendRow(rowVals);
        added++;
      }
    });
    return { status: 'success', added, updated, skipped };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function removeCIs(adminUser, adminPass, namesText) {
  try {
    _requireAdmin(adminUser, adminPass);
    const info = _ciSheetInfo();
    // Newline-separated only — CI names contain a comma themselves ("Last, First"),
    // so commas can't be used as a name delimiter here.
    const names = String(namesText || '').split(/\r?\n/).map(s => s.trim()).filter(Boolean);
    if (!names.length) return { status: 'error', message: 'Paste at least one CI name.' };

    const lastRow = info.sheet.getLastRow();
    if (lastRow < 2) return { status: 'success', removed: 0, notFound: names };
    const nameSet = new Set(names.map(n => n.toLowerCase()));
    const values = info.sheet.getRange(2, 1, lastRow - 1, info.headers.length).getValues();
    const foundKeys = new Set();
    const rowsToDelete = [];
    values.forEach((r, i) => {
      const name = String(r[info.nameCol] || '').trim();
      const key = name.toLowerCase();
      if (nameSet.has(key)) { rowsToDelete.push(i + 2); foundKeys.add(key); }
    });
    rowsToDelete.sort((a, b) => b - a).forEach(rowNum => info.sheet.deleteRow(rowNum));
    const notFound = names.filter(n => !foundKeys.has(n.toLowerCase()));
    return { status: 'success', removed: rowsToDelete.length, notFound };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

/* ================= JSON API ACTION MAP (for doPost / static frontend) ================= */
const API_ACTIONS = {
  getCIs, getStudents, getIncidents, getAdmins,
  authenticateCI, loginCI, changeCIPassword, authenticateAdmin,
  getStudentProfile, getFullReportColumnCByStudentId, getStudentHistory, getCIHistory,
  submitEntry, bulkAddExtensionEntries,
  bulkAddStudents, removeStudents,
  bulkAddCIs, removeCIs
};
