function doGet() {
  return HtmlService.createHtmlOutputFromFile('index_optimized')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
    .setTitle('Kineza Clinic Manager');
}

/**
 * Sheet layout guide for patient entry:
 * - Sheet name: Dashboard
 * - Sheet name: Patient Entry
 *   - B3: First Name
 *   - B4: Last Name
 *   - B5: Date of Birth (use a date cell if possible)
 *   - B6: Contact No
 *   - B7: Email
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Clinic Manager')
    .addItem('Go to Dashboard', 'goToDashboard')
    .addItem('Go to Patient Entry', 'goToPatientEntry')
    .addToUi();
}

function activateSheet(sheetName) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Sheet not found: ' + sheetName);
    return null;
  }
  sheet.activate();
  return sheet;
}

function goToDashboard() {
  activateSheet('Dashboard');
}

function goToPatientEntry() {
  activateSheet('Patient Entry');
}

function savePatientFromSheet() {
  const sheet = activateSheet('Patient Entry');
  if (!sheet) return;

  const firstName = String(sheet.getRange('B3').getValue() || '').trim();
  const lastName = String(sheet.getRange('B4').getValue() || '').trim();
  const dobValue = sheet.getRange('B5').getValue();
  const contactNo = String(sheet.getRange('B6').getValue() || '').trim();
  const email = String(sheet.getRange('B7').getValue() || '').trim();
  const dob = dobValue instanceof Date ? dobValue : new Date(dobValue);

  if (!firstName || !lastName || !dobValue || isNaN(dob.getTime())) {
    SpreadsheetApp.getUi().alert(
      'Please complete First Name, Last Name, and a valid Date of Birth.'
    );
    return;
  }

  savePatient({
    firstName,
    lastName,
    dob,
    contactNo,
    email
  });

  sheet.getRange('B3:B7').clearContent();
  SpreadsheetApp.getUi().alert('Patient saved successfully.');
}

const SHEET_ID = '1HiyGRS2-A8yHygWSOirPAWzpf6zLRpEyubtU-MyA1Vo';

/**
 * Branch staff Google accounts (lowercase email -> branch name on the Visits sheet).
 * These users use the full app except the Payments report tab. Staff Commissions are
 * always filtered to their branch on the server (client cannot override).
 * Optional override: Script property BRANCH_VIEWER_MAP as JSON merged on top of this map.
 *
 * Web app deploy (required for email checks):
 * - Execute as: User accessing the web application
 * - Who has access: signed-in users as appropriate for your clinic
 */
const BRANCH_COMMISSION_VIEWERS_DEFAULT = {
  'angeles@kinezabydrakris.com': 'Angeles',
  'olongapo@kinezabydrakris.com': 'Olongapo',
  'quezoncity@kinezabydrakris.com': 'Quezon City'
};

const ATTENDANCE_ADMIN_EMAILS_DEFAULT = ['admin@kinezabydrakris.com', 'your.email@kinezabydrakris.com'];

function getAttendanceAdminEmails() {
  const list = ATTENDANCE_ADMIN_EMAILS_DEFAULT.slice();
  try {
    const raw = PropertiesService.getScriptProperties().getProperty('ATTENDANCE_ADMIN_EMAILS');
    if (raw) {
      const extra = JSON.parse(raw);
      if (Array.isArray(extra)) {
        extra.forEach(function (email) {
          const normalized = String(email || '').trim().toLowerCase();
          if (normalized && list.indexOf(normalized) === -1) {
            list.push(normalized);
          }
        });
      }
    }
  } catch (e) {
    console.warn('ATTENDANCE_ADMIN_EMAILS parse error: ' + e);
  }
  return list;
}

function isAttendanceAdminUser() {
  return true;
}

function getBranchCommissionViewerMap() {
  const map = Object.assign({}, BRANCH_COMMISSION_VIEWERS_DEFAULT);
  try {
    const raw = PropertiesService.getScriptProperties().getProperty('BRANCH_VIEWER_MAP');
    if (raw) {
      const extra = JSON.parse(raw);
      if (extra && typeof extra === 'object') {
        Object.keys(extra).forEach(function (k) {
          map[String(k).trim().toLowerCase()] = String(extra[k]).trim();
        });
      }
    }
  } catch (e) {
    console.warn('BRANCH_VIEWER_MAP parse error: ' + e);
  }
  return map;
}

function getLockedBranchForCurrentUser() {
  return null;
}

function isBranchCommissionViewer() {
  return false;
}

function assertBranchUserCannotViewPaymentsReport() {
  return;
}

/**
 * Called from the HTML client on load. Branch-mapped users use the full app except the Payments report;
 * their Staff Commissions data is locked to their branch on the server (getStaffCommissions).
 */
function getWebAppAccess() {
  return {
    mode: 'full',
    lockBranch: '',
    email: String(Session.getActiveUser().getEmail() || ''),
    canViewPaymentsReport: true,
    canViewSalaryReport: true,
    canViewSalaryTable: true
  };
}

/** Most recent N visits loaded for the Visits tab (newest rows at bottom of sheet). Increase if you need a longer history in the grid. */
const MAX_LIST_VISITS = 800;

/** Most recent N rows for the Visit Procedures report tab. */
const MAX_VISIT_PROCEDURES_LIST = 1000;

/** Script cache max ~100KB per entry; skip caching when JSON is larger so the app stays responsive. */
const CACHE_MAX_CHARS = 95000;

let cachedSS = null;

function safeCachePut(cache, key, obj, ttlSec) {
  try {
    const s = typeof obj === 'string' ? obj : JSON.stringify(obj);
    if (s.length > CACHE_MAX_CHARS) {
      console.warn('Cache skip (payload too large): ' + key + ' len=' + s.length);
      return false;
    }
    cache.put(key, s, ttlSec);
    return true;
  } catch (e) {
    console.error('cache.put failed ' + key + ': ' + e);
    return false;
  }
}

function getSS() {
  if (!cachedSS) {
    cachedSS = SpreadsheetApp.openById(SHEET_ID);
  }
  return cachedSS;
}

function getPatientSheet() {
  const ss = getSS();
  const sheet = ss.getSheetByName('Patients');
  if (!sheet) throw new Error("Tab named 'Patients' not found!");
  return sheet;
}

function calculateAge(dob) {
  const d = dob instanceof Date ? dob : new Date(dob);
  if (isNaN(d.getTime())) return 0;
  const today = new Date();
  let age = today.getFullYear() - d.getFullYear();
  const m = today.getMonth() - d.getMonth();
  if (m < 0 || (m === 0 && today.getDate() < d.getDate())) age--;
  return age;
}

function getPatients() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('patients');
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (parseErr) {
        cache.remove('patients');
      }
    }

    const sheet = getPatientSheet();
    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) return [];

    const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    const tz = Session.getScriptTimeZone();

    let patients = data
      .filter(row => row[0])
      .map((row, i) => ({
        row: i + 2,
        id: row[0],
        sortKey: row[1] ? new Date(row[1]).getTime() : 0,
        dateRegistered: row[1] ? Utilities.formatDate(new Date(row[1]), tz, 'MM/dd/yyyy hh:mm a') : '',
        firstName: row[2],
        lastName: row[3],
        dob: row[4] ? Utilities.formatDate(new Date(row[4]), tz, 'MM/dd/yyyy') : '',
        age: row[5] + ' yrs old',
        gender: row[8] || ''
      }));

    patients.sort((a, b) => b.sortKey - a.sortKey);

    safeCachePut(cache, 'patients', patients, 900);
    return patients;
  } catch (e) {
    console.error('getPatients Error: ' + e);
    throw new Error(
      'Could not load patients. Share the spreadsheet with this Google account if the web app runs as the user, ' +
        'or confirm the Patients tab exists. ' +
        String(e && e.message ? e.message : e)
    );
  }
}

function normalizeDob(value) {
  const dob = value instanceof Date ? value : new Date(value);
  if (isNaN(dob.getTime())) {
    throw new Error('Invalid Date of Birth.');
  }
  return dob;
}

function getDailyExpenseCacheKeys(dateValue, branch) {
  const date = new Date(dateValue);
  if (isNaN(date.getTime())) return [];
  const dayKey = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return [
    `expenses_${dayKey}_${dayKey}_${branch || 'all'}`,
    `expenses_${dayKey}_${dayKey}_all`
  ];
}

function savePatient(data) {
  const sheet = getPatientSheet();
  const lastRow = sheet.getLastRow();
  const patientId = 'KPX-' + String(lastRow).padStart(4, '0');

  if (!data.firstName || !data.lastName) {
    throw new Error('First name and last name are required.');
  }

  if (!data.gender) {
    throw new Error('Gender is required.');
  }

  const dob = normalizeDob(data.dob);
  const age = calculateAge(dob);

  const dateRegistered = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy hh:mm a');

  sheet.appendRow([
    patientId,
    dateRegistered,
    data.firstName,
    data.lastName,
    dob,
    age,
    String(data.contactNo || '').trim(),
    String(data.email || '').trim(),
    data.gender
  ]);

  CacheService.getScriptCache().remove('patients');
  return true;
}

function updatePatient(data) {
  const sheet = getPatientSheet();
  const dob = normalizeDob(data.dob);
  const age = calculateAge(dob);

  if (!data.gender) {
    throw new Error('Gender is required.');
  }

  const range = sheet.getRange(data.row, 3, 1, 4);
  range.setValues([[data.firstName, data.lastName, dob, age]]);

  // Update gender in column 9
  sheet.getRange(data.row, 9).setValue(data.gender);

  CacheService.getScriptCache().remove('patients');
  return true;
}

function deletePatient(row) {
  const sheet = getPatientSheet();
  sheet.deleteRow(row);
  CacheService.getScriptCache().remove('patients');
  return true;
}

function searchPatients(keyword) {
  const patients = getPatients();
  const lowerKeyword = String(keyword || '').toLowerCase();

  return patients
    .filter(p => {
      const terms = [p.firstName, p.lastName].filter(Boolean).join(' ').toLowerCase();
      return terms.includes(lowerKeyword);
    })
    .map(p => ({
      id: p.id,
      firstName: p.firstName,
      lastName: p.lastName,
      dob: p.dob,
      age: p.age
    }));
}

/**
 * Lightweight typeahead for the Add Visit modal: scans the sheet once per call, returns at most 25 rows.
 * Avoids loading every patient into the browser datalist when the roster is large.
 */
function searchPatientsMini(query) {
  try {
    const t = String(query || '').trim().toLowerCase();
    if (t.length < 2) return [];

    const sheet = getPatientSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    const tz = Session.getScriptTimeZone();
    const out = [];

    for (let i = 0; i < data.length && out.length < 25; i++) {
      const row = data[i];
      if (!row[0]) continue;
      const fn = String(row[2] || '').toLowerCase();
      const ln = String(row[3] || '').toLowerCase();
      const full = (fn + ' ' + ln).trim();
      if (fn.indexOf(t) !== -1 || ln.indexOf(t) !== -1 || full.indexOf(t) !== -1) {
        out.push({
          id: row[0],
          firstName: row[2],
          lastName: row[3],
          dob: row[4] ? Utilities.formatDate(new Date(row[4]), tz, 'MM/dd/yyyy') : '',
          age: row[5] != null && row[5] !== '' ? String(row[5]) + ' yrs old' : ''
        });
      }
    }
    return out;
  } catch (e) {
    console.error('searchPatientsMini: ' + e);
    return [];
  }
}

function getProceduresByType(type) {
  const ss = getSS();
  const sheet = ss.getSheetByName('Procedures');
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

  return data
    .filter(row => String(row[0]).trim() === type && row[1])
    .map(row => String(row[1]).trim());
}

function getPatientDetails(lastName) {
  return getPatients().filter(p => p.lastName && p.lastName.toLowerCase() === lastName.toLowerCase());
}

function getVisits() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('visits');
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (parseErr) {
        cache.remove('visits');
      }
    }

    const ss = getSS();
    const visitsSheet = ss.getSheetByName('Visits');
    const patientsSheet = ss.getSheetByName('Patients');
    const paymentsSheet = ss.getSheetByName('Payments');
    const proceduresSheet = ss.getSheetByName('Visit Procedures');

    if (!visitsSheet || !patientsSheet || !paymentsSheet || !proceduresSheet) return [];

    const visitRows = visitsSheet.getLastRow();
    const patientRows = patientsSheet.getLastRow();
    const paymentRows = paymentsSheet.getLastRow();
    const procedureRows = proceduresSheet.getLastRow();

    const dataRowCount = visitRows > 1 ? visitRows - 1 : 0;
    const take = Math.min(MAX_LIST_VISITS, dataRowCount);
    const visitStartRow = take > 0 ? visitRows - take + 1 : 2;
    const visitData = visitRows > 1 ? visitsSheet.getRange(visitStartRow, 1, take, 5).getValues() : [];
    const patientData = patientRows > 1 ? patientsSheet.getRange(2, 1, patientRows - 1, 4).getValues() : [];
    const paymentData = paymentRows > 1 ? paymentsSheet.getRange(2, 1, paymentRows - 1, 2).getValues() : [];
    const procedureData = procedureRows > 1 ? proceduresSheet.getRange(2, 1, procedureRows - 1, 7).getValues() : [];

    const visitIdSet = {};
    visitData.forEach(row => {
      if (row[0]) visitIdSet[String(row[0])] = true;
    });

    const patientMap = {};
    patientData.forEach(row => {
      if (row[0]) patientMap[row[0]] = `${row[2]} ${row[3]}`;
    });

    const totalMap = {};
    const sessionMap = {};
    procedureData.forEach(row => {
      const visitId = String(row[0]);
      if (!visitId || !visitIdSet[visitId]) return;
      const price = Number(row[6]) || 0;
      const currentSession = Number(row[4]) || 1;
      const totalSessions = Number(row[5]) || 1;

      totalMap[visitId] = (totalMap[visitId] || 0) + price;

      if (!sessionMap[visitId]) sessionMap[visitId] = { current: 0, total: 0 };
      sessionMap[visitId].current = Math.max(sessionMap[visitId].current, currentSession);
      sessionMap[visitId].total = Math.max(sessionMap[visitId].total, totalSessions);
    });

    const paidMap = {};
    paymentData.forEach(row => {
      const visitId = String(row[0]);
      if (!visitId || !visitIdSet[visitId]) return;
      const amount = Number(String(row[1]).replace(/[₱,]/g, '')) || 0;
      paidMap[visitId] = (paidMap[visitId] || 0) + amount;
    });

    const tz = Session.getScriptTimeZone();
    const visits = visitData
      .filter(row => row[0])
      .map(row => {
        const visitId = String(row[0]);
        const total = totalMap[visitId] || 0;
        const paid = paidMap[visitId] || 0;
        const balance = Math.max(0, total - paid);
        const sessionData = sessionMap[visitId] || { current: 1, total: 1 };

        let paymentStatus = '';
        if (paid > 0) {
          paymentStatus = paid >= total ? 'PAID' : 'PAID WITH BALANCE';
        }

        let sessionStatus = '';
        if (sessionData.total > 1) {
          sessionStatus = sessionData.current >= sessionData.total ? 'SESSIONS COMPLETE' : 'SESSIONS REMAINING';
        }

        return {
          visitId,
          patientId: String(row[1]),
          patientName: patientMap[row[1]] || 'Unknown Patient',
          date: row[2] ? Utilities.formatDate(new Date(row[2]), tz, 'MM/dd/yyyy hh:mm a') : '',
          branch: String(row[3] || ''),
          remarks: String(row[4] || ''),
          paymentStatus,
          amountPaid: paid,
          sessionStatus,
          balance,
          paymentMethod: '',
          tipAmount: '',
          tipMethod: ''
        };
      });

    const finalData = visits.reverse();
    safeCachePut(cache, 'visits', finalData, 300);
    return finalData;
  } catch (e) {
    console.error('getVisits error: ' + e);
    throw new Error(
      'Could not load visits. Share the spreadsheet with this Google account if the web app runs as the user, ' +
        'or confirm Visits / Patients / Payments / Visit Procedures tabs exist. ' +
        String(e && e.message ? e.message : e)
    );
  }
}

function updateVisit(data) {
  const ss = getSS();
  const sheet = ss.getSheetByName('Visits');
  if (!sheet) return false;

  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.visitId)) {
      const range = sheet.getRange(i + 1, 3, 1, 3);
      range.setValues([[data.date, data.branch, data.remarks]]);

      const cache = CacheService.getScriptCache();
      cache.remove('visits');
      cache.remove('visitProcedures');
      cache.remove('visitDates');
      return true;
    }
  }
  return false;
}

function deleteVisit(visitId) {
  const ss = getSS();
  const sheet = ss.getSheetByName('Visits');
  if (!sheet) return false;

  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(visitId)) {
      sheet.deleteRow(i + 1);

      const cache = CacheService.getScriptCache();
      cache.remove('visits');
      cache.remove('visitProcedures');
      cache.remove('visitDates');
      return true;
    }
  }
  return false;
}

function saveVisit(data) {
  const ss = getSS();
  const sheet = ss.getSheetByName('Visits');
  if (!sheet) throw new Error("Create a sheet named 'Visits'");

  const lastRow = sheet.getLastRow();
  const visitId = 'VIS-' + String(lastRow).padStart(5, '0');

  sheet.appendRow([visitId, data.patientId, data.visitTime, data.branchLocation, data.remarks]);
  CacheService.getScriptCache().remove('visits');
  return true;
}

function saveProcedure(data) {
  const sheet = getSS().getSheetByName('Visit Procedures');
  if (!sheet) throw new Error("Create a sheet named 'Visit Procedures'");

  let bundleId = '';

  if (data.rowIndex) {
    bundleId = sheet.getRange(Number(data.rowIndex), 9).getValue();
    sheet.getRange(Number(data.rowIndex), 1, 1, 9).setValues([[
      data.visitId, data.type, data.name, data.variant, data.session,
      data.totalSessions, data.price, data.staff, bundleId
    ]]);
  } else {
    bundleId = data.bundleId || 'TRT-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMddHHmmss');
    sheet.appendRow([data.visitId, data.type, data.name, data.variant, data.session,
    data.totalSessions, data.price, data.staff, bundleId]);
  }

  const cache = CacheService.getScriptCache();
  cache.remove('visits');
  cache.remove('visitProcedures');
  cache.remove('visitDates');

  // Auto-update related export sheets when procedure data changes.
  try {
    const range = getDefaultSalaryExportRange();
    exportSalaryReportToSheetInternal(range.startDate, range.endDate, '', '');
    exportCommissionsToSheetInternal(range.startDate, range.endDate, '', '');
  } catch (e) {
    console.error('Auto-update exports failed: ' + e);
  }

  return true;
}

function deleteProcedure(rowIndex) {
  const ss = getSS();
  const sheet = ss.getSheetByName('Visit Procedures');
  if (!sheet) throw new Error("Create a sheet named 'Visit Procedures'");

  sheet.deleteRow(rowIndex);

  const cache = CacheService.getScriptCache();
  cache.remove('visits');
  cache.remove('visitProcedures');
  cache.remove('visitDates');

  // Auto-update commission exports when a procedure row is deleted.
  try {
    const range = getDefaultSalaryExportRange();
    exportCommissionsToSheetInternal(range.startDate, range.endDate, '', '');
  } catch (e) {
    console.error('Auto-update commission exports failed: ' + e);
  }

  return true;
}

function getProcedureNames(type) {
  const ss = getSS();
  const sheet = ss.getSheetByName('Procedures Master');
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

  return data
    .filter(row => row[0] == type)
    .map(row => row[1]);
}

function getStaffByBranch(branch) {
  const ss = getSS();
  const sheet = ss.getSheetByName('Staff');
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

  return data
    .filter(row => !branch || row[2] === branch)
    .map(row => ({ name: row[1], branch: row[2] }));
}

function ensureAttendanceSheet() {
  const ss = getSS();
  let sheet = ss.getSheetByName('Attendance');
  if (!sheet) {
    sheet = ss.insertSheet('Attendance');
    sheet.getRange(1, 1, 1, 8).setValues([[
      'Attendance ID',
      'Date',
      'Branch',
      'Staff',
      'Time In',
      'Time Out',
      'Hours Worked',
      'Notes'
    ]]);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(4, 160);
    sheet.setColumnWidth(5, 150);
    sheet.setColumnWidth(6, 150);
    sheet.setColumnWidth(7, 120);
    sheet.setColumnWidth(8, 220);
  }
  return sheet;
}

function getStaffWithSalary(branch) {
  const ss = getSS();
  const sheet = ss.getSheetByName('Staff');
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  return data
    .filter(row => row[1] && (!branch || String(row[2]) === branch))
    .map(row => ({
      name: String(row[1] || ''),
      branch: String(row[2] || ''),
      hourlyRate: Number(row[3]) || 0
    }));
}

function saveStaffSalary(data) {
  const lockedBranch = getLockedBranchForCurrentUser();
  if (lockedBranch && data.branch !== lockedBranch) {
    throw new Error('You can only update salary information for your branch.');
  }
  if (!data.staff || !data.branch) {
    throw new Error('Staff and branch are required.');
  }
  const ss = getSS();
  const sheet = ss.getSheetByName('Staff');
  if (!sheet) throw new Error("Create a sheet named 'Staff'");

  const lastRow = sheet.getLastRow();
  const values = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 4).getValues() : [];
  let found = false;

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][1]) === String(data.staff) && String(values[i][2]) === String(data.branch)) {
      sheet.getRange(i + 2, 4).setValue(Number(data.hourlyRate) || 0);
      found = true;
      break;
    }
  }

  if (!found) {
    sheet.appendRow(['', data.staff, data.branch, Number(data.hourlyRate) || 0]);
  }

  // Auto-update Salary Reports sheet when a daily rate changes
  try {
    const range = getDefaultSalaryExportRange();
    exportSalaryReportToSheetInternal(range.startDate, range.endDate, '', '');
  } catch (e) {
    console.error('Auto-update salary reports failed: ' + e);
  }

  return true;
}

function clockInAttendance(staff, branch, notes) {
  const lockedBranch = getLockedBranchForCurrentUser();
  if (lockedBranch) {
    branch = lockedBranch;
  }
  if (!staff || !branch) {
    throw new Error('Please select a branch and staff member.');
  }

  const sheet = ensureAttendanceSheet();
  const lastRow = sheet.getLastRow();
  const data = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 8).getValues() : [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let openRecordFound = false;
  let todayRecordFound = false;

  data.forEach(row => {
    if (String(row[2]) !== String(branch) || String(row[3]) !== String(staff)) return;
    const rowDate = row[1] instanceof Date ? row[1] : new Date(row[1]);
    if (isNaN(rowDate.getTime())) return;
    const normalizedRowDate = new Date(rowDate);
    normalizedRowDate.setHours(0, 0, 0, 0);

    if (!row[5]) {
      openRecordFound = true;
    }
    if (normalizedRowDate.getTime() === today.getTime()) {
      todayRecordFound = true;
    }
  });

  if (openRecordFound) {
    throw new Error('This staff member is already clocked in and must clock out before clocking in again.');
  }
  if (todayRecordFound) {
    throw new Error('This staff member already has an attendance entry for today and cannot clock in twice.');
  }

  const attendanceId = 'ATT-' + String(lastRow).padStart(5, '0');
  const now = new Date();
  const dateOnly = new Date(now);
  dateOnly.setHours(0, 0, 0, 0);

  sheet.appendRow([attendanceId, dateOnly, branch, staff, now, '', '', String(notes || '')]);
  return {
    attendanceId: attendanceId,
    action: 'clockedIn',
    staff: staff,
    branch: branch,
    timeIn: now
  };
}

function clockOutAttendance(staff, branch) {
  const lockedBranch = getLockedBranchForCurrentUser();
  if (lockedBranch) {
    branch = lockedBranch;
  }
  if (!staff || !branch) {
    throw new Error('Please select a branch and staff member.');
  }

  const sheet = ensureAttendanceSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    throw new Error('No attendance records found to clock out.');
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  let openRowIndex = -1;
  let hasClosedToday = false;
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];
    if (String(row[3]) !== String(staff) || String(row[2]) !== String(branch)) continue;
    const rowDate = row[1] instanceof Date ? row[1] : new Date(row[1]);
    if (!isNaN(rowDate.getTime())) {
      const normalizedRowDate = new Date(rowDate);
      normalizedRowDate.setHours(0, 0, 0, 0);
      if (normalizedRowDate.getTime() === today.getTime() && row[5]) {
        hasClosedToday = true;
      }
    }
    if (!row[5] && row[4]) {
      openRowIndex = i + 2;
      break;
    }
  }

  if (openRowIndex === -1) {
    if (hasClosedToday) {
      throw new Error('This staff member has already clocked out today and cannot clock out again.');
    }
    throw new Error('No open clock-in found for this staff member.');
  }

  const timeInValue = sheet.getRange(openRowIndex, 5).getValue();
  const timeOutValue = new Date();
  const timeInDate = timeInValue instanceof Date ? timeInValue : new Date(timeInValue);
  const hoursWorked = timeInDate && !isNaN(timeInDate.getTime())
    ? Math.round(((timeOutValue.getTime() - timeInDate.getTime()) / 3600000) * 100) / 100
    : 0;

  sheet.getRange(openRowIndex, 6, 1, 2).setValues([[timeOutValue, hoursWorked]]);
  return {
    attendanceId: sheet.getRange(openRowIndex, 1).getValue(),
    action: 'clockedOut',
    staff: staff,
    branch: branch,
    timeOut: timeOutValue,
    hoursWorked: hoursWorked
  };
}

function getAttendanceRecords(startDate, endDate, branch, staff) {
  const lockedBranch = getLockedBranchForCurrentUser();
  if (lockedBranch) {
    branch = lockedBranch;
  }
  const sheet = ensureAttendanceSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const start = startDate ? new Date(startDate) : new Date(0);
  const end = endDate ? new Date(endDate) : new Date(8640000000000000);
  if (endDate) end.setHours(23, 59, 59, 999);
  const branchFilter = branch ? String(branch).trim() : '';
  const staffFilter = staff ? String(staff).trim() : '';
  const tz = Session.getScriptTimeZone();

  return data
    .filter(row => {
      if (!row[0]) return false;
      const branchMatch = branchFilter ? String(row[2] || '') === branchFilter : true;
      const staffMatch = staffFilter ? String(row[3] || '') === staffFilter : true;
      if (!branchMatch || !staffMatch) return false;
      if (startDate || endDate) {
        const dateValue = row[1] ? new Date(row[1]) : null;
        if (!dateValue || isNaN(dateValue.getTime())) return false;
        if (startDate && dateValue < start) return false;
        if (endDate && dateValue > end) return false;
      }
      return true;
    })
    .map(row => ({
      id: String(row[0] || ''),
      date: row[1] ? Utilities.formatDate(new Date(row[1]), tz, 'MM/dd/yyyy') : '',
      branch: String(row[2] || ''),
      staff: String(row[3] || ''),
      timeIn: row[4] ? Utilities.formatDate(new Date(row[4]), tz, 'MM/dd/yyyy hh:mm a') : '',
      timeOut: row[5] ? Utilities.formatDate(new Date(row[5]), tz, 'MM/dd/yyyy hh:mm a') : '',
      hoursWorked: Number(row[6]) || 0,
      notes: String(row[7] || '')
    }));
}

function getOpenAttendanceStatus(branch, staff) {
  const sheet = ensureAttendanceSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { open: false };

  const data = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 8).getValues() : [];
  
  for (let i = data.length - 1; i >= 0; i--) {
    const row = data[i];
    if (String(row[2]) !== String(branch) || String(row[3]) !== String(staff)) continue;
    if (!row[5]) {
      return { open: true };
    }
    break;
  }
  
  return { open: false };
}

function getSalaryReport(startDate, endDate, branch, staff) {
  if (!isAttendanceAdminUser()) {
    throw new Error('You do not have permission to access salary reports.');
  }
  return getSalaryReportData(startDate, endDate, branch, staff);
}

function getSalaryReportData(startDate, endDate, branch, staff) {
  const attendance = getAttendanceRecords(startDate, endDate, branch, staff);
  const staffList = getStaffWithSalary(branch);
  const salaryMap = {};
  staffList.forEach(item => {
    salaryMap[item.name] = {
      hourlyRate: item.hourlyRate || 0,
      branch: item.branch
    };
  });

  const rows = [];
  let totalHours = 0;
  let totalPay = 0;

  attendance.forEach(record => {
    if (!record.hoursWorked || record.hoursWorked <= 0) return;
    const salaryInfo = salaryMap[record.staff] || { hourlyRate: 0, branch: record.branch };
    const pay = Math.round((record.hoursWorked / 8) * salaryInfo.hourlyRate * 100) / 100;
    totalHours += record.hoursWorked;
    totalPay += pay;
    rows.push({
      id: record.id,
      date: record.date,
      branch: record.branch,
      staff: record.staff,
      hoursWorked: record.hoursWorked,
      hourlyRate: salaryInfo.hourlyRate,
      pay: pay,
      notes: record.notes
    });
  });

  const staffSummary = {};
  rows.forEach(row => {
    if (!staffSummary[row.staff]) {
      staffSummary[row.staff] = {
        staff: row.staff,
        branch: row.branch,
        hourlyRate: row.hourlyRate,
        hoursWorked: 0,
        pay: 0
      };
    }
    staffSummary[row.staff].hoursWorked += row.hoursWorked;
    staffSummary[row.staff].pay += row.pay;
  });

  return {
    rows: rows.sort((a, b) => new Date(b.date) - new Date(a.date)),
    totalHours: Math.round(totalHours * 100) / 100,
    totalPay: Math.round(totalPay * 100) / 100,
    staffSummary: Object.values(staffSummary)
  };
}

function getDefaultSalaryExportRange() {
  const nowDate = new Date();
  const endDate = new Date(nowDate);
  const startDate = new Date(nowDate);
  startDate.setDate(nowDate.getDate() - 13);
  return {
    startDate: Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    endDate: Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')
  };
}

function exportSalaryReportToSheetInternal(startDate, endDate, branch, staff) {
  const report = getSalaryReportData(startDate, endDate, branch, staff);
  const ss = getSS();
  let sheet = ss.getSheetByName('Salary Reports');
  if (!sheet) {
    sheet = ss.insertSheet('Salary Reports');
    sheet.getRange(1, 1, 1, 8).setValues([[
      'Date', 'Branch', 'Staff', 'Hours Worked', 'Daily Rate', 'Pay', 'Notes', 'Exported At'
    ]]);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  }

  // Clear existing data except header
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 8).clearContent();
  }

  if (report.rows.length > 0) {
    const data = report.rows.map(row => [
      row.date, row.branch, row.staff, row.hoursWorked, row.hourlyRate, row.pay, row.notes, new Date()
    ]);
    sheet.getRange(2, 1, data.length, 8).setValues(data);
  }

  // Add summary at the bottom
  const summaryRow = report.rows.length + 3;
  sheet.getRange(summaryRow, 1, 1, 8).setValues([[
    'Summary', '', '', report.totalHours, '', report.totalPay, '', new Date()
  ]]);
  sheet.getRange(summaryRow, 1, 1, 8).setFontWeight('bold');

  return true;
}

function exportSalaryReportToSheet(startDate, endDate, branch, staff) {
  if (!isAttendanceAdminUser()) {
    throw new Error('You do not have permission to export salary reports.');
  }
  return exportSalaryReportToSheetInternal(startDate, endDate, branch, staff);
}

function exportCommissionsToSheetInternal(startDate, endDate, branch, staff) {
  const commissions = getStaffCommissions(startDate, endDate, branch, staff);
  const ss = getSS();
  let sheet = ss.getSheetByName('Staff Commissions');
  if (!sheet) {
    sheet = ss.insertSheet('Staff Commissions');
    sheet.getRange(1, 1, 1, 6).setValues([[
      'Staff', 'Date', 'Procedure', 'Commission', 'Tip', 'Total'
    ]]);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }

  // Clear existing data except header
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 6).clearContent();
  }

  if (commissions.length > 0) {
    const data = commissions.map(row => [
      row.staff, row.date, row.procedure, row.amount, row.tip, row.total
    ]);
    sheet.getRange(2, 1, data.length, 6).setValues(data);
  }

  return true;
}

function exportCommissionsToSheet(startDate, endDate, branch, staff) {
  return exportCommissionsToSheetInternal(startDate, endDate, branch, staff);
}

function getVisitProcedures(visitId) {
  try {
    const cache = CacheService.getScriptCache();
    let procedures = cache.get('visitProcedures');
    let visits = cache.get('visitDates');

    if (procedures && visits) {
      procedures = JSON.parse(procedures);
      visits = JSON.parse(visits);
    } else {
      const ss = getSS();
      const procedureSheet = ss.getSheetByName('Visit Procedures');
      const visitSheet = ss.getSheetByName('Visits');

      if (!procedureSheet || !visitSheet) return [];

      const procLastRow = procedureSheet.getLastRow();
      const visitLastRow = visitSheet.getLastRow();

      procedures = procLastRow > 1 ? procedureSheet.getRange(2, 1, procLastRow - 1, 8).getValues() : [];
      visits = visitLastRow > 1 ? visitSheet.getRange(2, 1, visitLastRow - 1, 3).getValues() : [];

      safeCachePut(cache, 'visitProcedures', procedures, 180);
      safeCachePut(cache, 'visitDates', visits, 180);
    }

    let visitDate = '';
    const tz = Session.getScriptTimeZone();

    for (let i = 0; i < visits.length; i++) {
      if (String(visits[i][0]) === String(visitId)) {
        visitDate = Utilities.formatDate(new Date(visits[i][2]), tz, 'MM/dd/yyyy');
        break;
      }
    }

    return procedures
      .filter(row => String(row[0]) === String(visitId))
      .map(row => ({
        procedure: row[2] || '',
        date: visitDate,
        session: row[4] || '',
        totalSessions: row[5] || '',
        staff: row[7] || '',
        price: row[6] || 0
      }));
  } catch (e) {
    console.error(e);
    return [];
  }
}

function savePayment(data) {
  const ss = getSS();
  const sheet = ss.getSheetByName('Payments');
  if (!sheet) throw new Error("Create a sheet named 'Payments'");

  sheet.appendRow([data.visitId, data.amountPaid, data.method, data.tipAmount, data.tipMethod]);

  // Clear caches that depend on payment data
  const cache = CacheService.getScriptCache();
  cache.remove('visits');

  const visitSheet = ss.getSheetByName('Visits');
  if (visitSheet) {
    const rows = visitSheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(data.visitId).trim()) {
        const rowDate = rows[i][2];
        const branch = String(rows[i][3] || '').trim();
        if (rowDate) {
          const formattedDate = Utilities.formatDate(new Date(rowDate), Session.getScriptTimeZone(), 'yyyy-MM-dd');
          cache.remove(`cashflow_${formattedDate}_${branch || 'all'}`);
          cache.remove(`cashflow_${formattedDate}_all`);
        }
        break;
      }
    }
  }

  // Auto-update commission exports when payments change (tips may affect staff commission totals).
  try {
    const range = getDefaultSalaryExportRange();
    exportCommissionsToSheetInternal(range.startDate, range.endDate, '', '');
  } catch (e) {
    console.error('Auto-update commission exports failed: ' + e);
  }

  return true;
}

function getPaymentDetails(visitId) {
  const ss = getSS();
  const sheet = ss.getSheetByName('Payments');
  if (!sheet) {
    return { amountPaid: '', method: '', tipAmount: '', tipMethod: '' };
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return { amountPaid: '', method: '', tipAmount: '', tipMethod: '' };
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const payments = [];
  const methods = [];
  const tipAmounts = [];
  const tipMethods = [];

  data.forEach(row => {
    if (String(row[0]).trim() === String(visitId).trim()) {
      if (row[1]) payments.push(row[1]);
      if (row[2]) methods.push(row[2]);
      if (row[3]) tipAmounts.push(row[3]);
      if (row[4]) tipMethods.push(row[4]);
    }
  });

  return {
    amountPaid: payments.join(' + '),
    method: methods.join(', '),
    tipAmount: tipAmounts.join(' + '),
    tipMethod: tipMethods.join(', ')
  };
}

/**
 * Payment ledger joined to Visits (date, branch) and Patients (name).
 * @param {string} startDate - YYYY-MM-DD or empty for no lower bound
 * @param {string} endDate - YYYY-MM-DD or empty for no upper bound
 * @param {string} branch - exact branch name or empty for all
 * @return {{rows:Object[],totalAmount:number,totalTip:number,grandTotal:number,byMethod:Object,byTipMethod:Object,count:number}}
 */
function getPaymentsReport(startDate, endDate, branch) {
  assertBranchUserCannotViewPaymentsReport();
  const ss = getSS();
  const paySheet = ss.getSheetByName('Payments');
  const visitSheet = ss.getSheetByName('Visits');
  const patientSheet = ss.getSheetByName('Patients');

  const empty = { rows: [], totalAmount: 0, totalTip: 0, grandTotal: 0, byMethod: {}, byTipMethod: {}, count: 0 };
  if (!paySheet || !visitSheet) return empty;

  const parseMoney = function (v) {
    return Number(String(v).replace(/[₱,]/g, '').trim()) || 0;
  };

  const vLast = visitSheet.getLastRow();
  const visitMap = {};
  if (vLast > 1) {
    const vData = visitSheet.getRange(2, 1, vLast - 1, 4).getValues();
    vData.forEach(function (row) {
      if (!row[0]) return;
      visitMap[String(row[0])] = {
        date: new Date(row[2]),
        branch: String(row[3] || ''),
        patientId: row[1]
      };
    });
  }

  const patientMap = {};
  if (patientSheet) {
    const pLast = patientSheet.getLastRow();
    if (pLast > 1) {
      const pData = patientSheet.getRange(2, 1, pLast - 1, 4).getValues();
      pData.forEach(function (row) {
        if (row[0]) patientMap[String(row[0])] = String(row[2] || '') + ' ' + String(row[3] || '');
      });
    }
  }

  const start = startDate ? new Date(startDate) : new Date(0);
  const end = endDate ? new Date(endDate) : new Date(8640000000000000);
  end.setHours(23, 59, 59, 999);

  const branchFilter = branch ? String(branch).trim() : '';

  const payLast = paySheet.getLastRow();
  const rows = [];
  const byMethod = {};
  const byTipMethod = {};
  let totalAmount = 0;
  let totalTip = 0;
  const tz = Session.getScriptTimeZone();

  if (payLast >= 2) {
    const payData = paySheet.getRange(2, 1, payLast - 1, 5).getValues();
    payData.forEach(function (row) {
      const visitId = String(row[0] || '').trim();
      if (!visitId) return;
      const info = visitMap[visitId];
      if (!info) return;

      const vd = info.date;
      if (isNaN(vd.getTime())) return;
      if (vd < start || vd > end) return;
      if (branchFilter && String(info.branch) !== branchFilter) return;

      const amt = parseMoney(row[1]);
      const meth = String(row[2] || '').trim() || '—';
      const tip = parseMoney(row[3]);
      const tipMeth = String(row[4] || '').trim() || '—';

      const patientName = (patientMap[String(info.patientId)] || '').trim() || '—';

      totalAmount += amt;
      totalTip += tip;
      byMethod[meth] = (byMethod[meth] || 0) + amt;
      if (tip > 0) {
        byTipMethod[tipMeth] = (byTipMethod[tipMeth] || 0) + tip;
      }

      rows.push({
        visitId: visitId,
        patientName: patientName,
        visitDate: Utilities.formatDate(vd, tz, 'MM/dd/yyyy hh:mm a'),
        branch: info.branch,
        amountPaid: amt,
        method: meth,
        tipAmount: tip,
        tipMethod: tipMeth
      });
    });
  }

  rows.sort(function (a, b) {
    const ia = visitMap[a.visitId];
    const ib = visitMap[b.visitId];
    const ta = ia && ia.date && !isNaN(ia.date.getTime()) ? ia.date.getTime() : 0;
    const tb = ib && ib.date && !isNaN(ib.date.getTime()) ? ib.date.getTime() : 0;
    return tb - ta;
  });

  return {
    rows: rows,
    totalAmount: totalAmount,
    totalTip: totalTip,
    grandTotal: totalAmount + totalTip,
    byMethod: byMethod,
    byTipMethod: byTipMethod,
    count: rows.length
  };
}

function getAllVisitProcedures(limit) {
  const max = Math.min(5000, Math.max(100, Number(limit) || MAX_VISIT_PROCEDURES_LIST));
  const ss = getSS();
  const sheet = ss.getSheetByName('Visit Procedures');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const body = data.slice(1);
  const slice = body.length > max ? body.slice(body.length - max) : body;

  return slice
    .map(row => ({
      visitId: row[0],
      procedure: row[2],
      staff: row[7],
      price: row[6]
    }))
    .reverse();
}

function getProcedures() {
  const ss = getSS();
  const sheet = ss.getSheetByName('Procedures Master');
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  values.shift();

  return values.map((row, index) => ({
    row: index + 2,
    type: row[0],
    name: row[1],
    price: row[2]
  }));
}

function addProcedure(data) {
  const ss = getSS();
  const sheet = ss.getSheetByName('Procedures Master');
  if (!sheet) throw new Error("Create a sheet named 'Procedures Master'");

  sheet.appendRow([data.type, data.name, data.price]);
  return true;
}

function getStaffCommissions(startDate, endDate, branch, staff) {
  const lockedBranch = getLockedBranchForCurrentUser();
  if (lockedBranch) {
    branch = lockedBranch;
  }
  const ss = getSS();
  const visitSheet = ss.getSheetByName('Visits');
  const procedureSheet = ss.getSheetByName('Visit Procedures');
  const paymentSheet = ss.getSheetByName('Payments');

  if (!visitSheet || !procedureSheet) return [];

  const visitLastRow = visitSheet.getLastRow();
  const procLastRow = procedureSheet.getLastRow();
  const paymentLastRow = paymentSheet ? paymentSheet.getLastRow() : 0;

  const visitData = visitLastRow > 1 ? visitSheet.getRange(2, 1, visitLastRow - 1, 4).getValues() : [];
  const procedureData = procLastRow > 1 ? procedureSheet.getRange(2, 1, procLastRow - 1, 9).getValues() : [];
  const paymentData = paymentLastRow > 1 ? paymentSheet.getRange(2, 1, paymentLastRow - 1, 5).getValues() : [];

  const visitDateMap = {};
  visitData.forEach(row => {
    if (row[0]) visitDateMap[row[0]] = { date: new Date(row[2]), branch: row[3] };
  });

  const bundlePriceMap = {};
  procedureData.forEach(row => {
    const bundleId = row[8];
    const price = Number(row[6]) || 0;
    if (bundleId && price > 0 && !bundlePriceMap[bundleId]) {
      bundlePriceMap[bundleId] = price;
    }
  });

  const visitTipMap = {};
  const parseCurrency = value => Number(String(value).replace(/[₱,]/g, '').trim()) || 0;
  paymentData.forEach(row => {
    const vid = row[0];
    if (!vid) return;
    const tipAmount = parseCurrency(row[3]);
    visitTipMap[vid] = (visitTipMap[vid] || 0) + tipAmount;
  });

  const visitStaffMap = {};
  procedureData.forEach(row => {
    const vid = row[0];
    const staffList = (row[7] || '').split(',').map(s => s.trim()).filter(Boolean);
    if (!visitStaffMap[vid]) {
      visitStaffMap[vid] = new Set();
    }
    staffList.forEach(name => visitStaffMap[vid].add(name));
  });

  const start = new Date(startDate);
  const end = new Date(endDate);
  end.setHours(23, 59, 59, 999);

  const commissionMap = {};
  const tz = Session.getScriptTimeZone();

  procedureData.forEach(row => {
    const visitId = row[0];
    const type = row[1];
    const procedureName = row[2];
    const totalSessions = Number(row[5]) || 1;
    const staffString = row[7] || '';
    const bundleId = row[8];

    const visitInfo = visitDateMap[visitId];
    if (!visitInfo) return;

    const visitDate = visitInfo.date;
    const visitBranch = visitInfo.branch;

    if (visitDate < start || visitDate > end) return;
    if (branch && visitBranch !== branch) return;

    const staffList = staffString.split(',').map(s => s.trim()).filter(Boolean);
    if (staffList.length === 0) return;

    let price = Number(row[6]) || 0;
    if (price === 0 && bundleId) {
      price = bundlePriceMap[bundleId] || 0;
    }

    let totalCommission = 0;
    if (type === 'Doctor') totalCommission = 100;
    else if (type === 'Clinic') totalCommission = (price / totalSessions) * 0.10;
    else if (type === 'Product') totalCommission = price * 0.10;

    const perStaff = totalCommission / staffList.length;

    staffList.forEach(name => {
      if (staff && name !== staff) return;
      const key = `${visitId}::${name}`;
      if (!commissionMap[key]) {
        commissionMap[key] = {
          visitId,
          staff: name,
          date: Utilities.formatDate(visitDate, tz, 'MM/dd/yyyy'),
          procedures: new Set(),
          amount: 0
        };
      }
      commissionMap[key].procedures.add(procedureName);
      commissionMap[key].amount += perStaff;
    });
  });

  const results = Object.values(commissionMap).map(entry => {
    const staffCount = visitStaffMap[entry.visitId] ? visitStaffMap[entry.visitId].size : 0;
    const tip = staffCount ? (visitTipMap[entry.visitId] || 0) / staffCount : 0;
    return {
      staff: entry.staff,
      date: entry.date,
      procedure: Array.from(entry.procedures).join(', '),
      amount: entry.amount,
      tip: tip,
      total: entry.amount + tip
    };
  });

  results.sort((a, b) => new Date(b.date) - new Date(a.date));
  return results;
}

function updateProcedure(data) {
  const ss = getSS();
  const sheet = ss.getSheetByName('Procedures Master');
  if (!sheet) throw new Error("Create a sheet named 'Procedures Master'");

  const range = sheet.getRange(data.row, 1, 1, 3);
  range.setValues([[data.type, data.name, data.price]]);
  return true;
}

function getAmountPaid(visitId) {
  const ss = getSS();
  const paymentSheet = ss.getSheetByName('Payments');
  if (!paymentSheet) return 0;

  const lastRow = paymentSheet.getLastRow();
  if (lastRow < 2) return 0;

  const payments = paymentSheet.getRange(2, 1, lastRow - 1, 2).getValues();
  let totalPaid = 0;

  payments.forEach(row => {
    if (row[0] == visitId) {
      totalPaid += Number(String(row[1]).replace(/[₱,]/g, '')) || 0;
    }
  });

  return totalPaid;
}

function getActiveTreatmentBundles(patientId) {
  const ss = getSS();
  const visitSheet = ss.getSheetByName('Visits');
  const procedureSheet = ss.getSheetByName('Visit Procedures');

  if (!visitSheet || !procedureSheet) return [];

  const visitLast = visitSheet.getLastRow();
  const procLast = procedureSheet.getLastRow();

  if (visitLast < 2 || procLast < 2) return [];

  const visits = visitSheet.getRange(2, 1, visitLast - 1, 2).getValues();
  const procedures = procedureSheet.getRange(2, 1, procLast - 1, 9).getValues();

  const patientVisitIds = visits
    .filter(row => row[1] == patientId)
    .map(row => row[0]);

  const bundles = {};

  procedures.forEach(row => {
    const visitId = row[0];
    if (!patientVisitIds.includes(visitId)) return;

    const bundleId = row[8];
    const session = Number(row[4]);
    const total = Number(row[5]);

    if (!bundleId || session >= total) return;

    if (!bundles[bundleId] || session > bundles[bundleId].lastSession) {
      bundles[bundleId] = {
        bundleId,
        procedure: row[2],
        type: row[1],
        variant: row[3],
        nextSession: session + 1,
        totalSessions: total,
        price: row[6],
        lastSession: session
      };
    }
  });

  return Object.values(bundles);
}

function ensureDailyExpensesSheet() {
  const ss = getSS();
  let sheet = ss.getSheetByName('Daily Expenses');
  if (!sheet) {
    sheet = ss.insertSheet('Daily Expenses');
    // Add headers
    sheet.getRange(1, 1, 1, 5).setValues([['Expense ID', 'Date', 'Branch', 'Description', 'Amount']]);
    // Format the header
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    // Set column widths
    sheet.setColumnWidth(1, 120); // Expense ID
    sheet.setColumnWidth(2, 100); // Date
    sheet.setColumnWidth(3, 120); // Branch
    sheet.setColumnWidth(4, 250); // Description
    sheet.setColumnWidth(5, 100); // Amount
  }
  return sheet;
}

/**
 * Daily Expenses Management Functions
 */

function getDailyExpenses(startDate, endDate, branch) {
  const lockedBranch = getLockedBranchForCurrentUser();
  if (lockedBranch) {
    branch = lockedBranch;
  }

  try {
    // Ensure the Daily Expenses sheet exists
    ensureDailyExpensesSheet();

    const cache = CacheService.getScriptCache();
    const cacheKey = `expenses_${startDate}_${endDate}_${branch || 'all'}`;
    const cached = cache.get(cacheKey);
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (parseErr) {
        cache.remove(cacheKey);
      }
    }

    const ss = getSS();
    const expensesSheet = ss.getSheetByName('Daily Expenses');
    if (!expensesSheet) return [];

    const lastRow = expensesSheet.getLastRow();
    if (lastRow <= 1) return [];

    const data = expensesSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const tz = Session.getScriptTimeZone();

    let expenses = data
      .filter(row => row[0])
      .map((row, i) => ({
        row: i + 2,
        id: row[0],
        date: row[1] ? Utilities.formatDate(new Date(row[1]), tz, 'MM/dd/yyyy') : '',
        branch: String(row[2] || ''),
        description: String(row[3] || ''),
        amount: Number(row[4]) || 0,
        sortKey: row[1] ? new Date(row[1]).getTime() : 0
      }));

    // Apply filters
    if (startDate) {
      const start = new Date(startDate);
      expenses = expenses.filter(exp => new Date(exp.sortKey) >= start);
    }

    if (endDate) {
      const end = new Date(endDate);
      end.setHours(23, 59, 59, 999);
      expenses = expenses.filter(exp => new Date(exp.sortKey) <= end);
    }

    if (branch) {
      expenses = expenses.filter(exp => exp.branch === branch);
    }

    expenses.sort((a, b) => b.sortKey - a.sortKey);

    safeCachePut(cache, cacheKey, expenses, 300);
    return expenses;
  } catch (e) {
    console.error('getDailyExpenses error: ' + e);
    throw new Error('Could not load daily expenses. ' + String(e.message || e));
  }
}

function saveDailyExpense(data) {
  const lockedBranch = getLockedBranchForCurrentUser();
  if (lockedBranch && data.branch !== lockedBranch) {
    throw new Error('You can only add expenses for your branch.');
  }

  const ss = getSS();
  const sheet = ensureDailyExpensesSheet();

  const lastRow = sheet.getLastRow();
  const expenseId = 'EXP-' + String(lastRow).padStart(5, '0');

  const date = new Date(data.date);
  const amount = Number(data.amount) || 0;

  sheet.appendRow([expenseId, date, data.branch, data.description, amount]);

  const cache = CacheService.getScriptCache();
  getDailyExpenseCacheKeys(date, data.branch).forEach(key => cache.remove(key));
  return true;
}

function updateDailyExpense(data) {
  const lockedBranch = getLockedBranchForCurrentUser();
  if (lockedBranch) {
    // Check if the expense belongs to their branch
    const ss = getSS();
    const sheet = ensureDailyExpensesSheet();
    if (sheet) {
      const currentBranch = sheet.getRange(data.row, 3).getValue();
      if (currentBranch !== lockedBranch) {
        throw new Error('You can only edit expenses for your branch.');
      }
    }
  }

  const ss = getSS();
  const sheet = ss.getSheetByName('Daily Expenses');
  if (!sheet) throw new Error("Create a sheet named 'Daily Expenses'");

  const date = new Date(data.date);
  const amount = Number(data.amount) || 0;

  const range = sheet.getRange(data.row, 2, 1, 4);
  range.setValues([[date, data.branch, data.description, amount]]);

  const cache = CacheService.getScriptCache();
  getDailyExpenseCacheKeys(date, data.branch).forEach(key => cache.remove(key));
  return true;
}

function deleteDailyExpense(rowIndex) {
  const lockedBranch = getLockedBranchForCurrentUser();
  if (lockedBranch) {
    // Check if the expense belongs to their branch
    const ss = getSS();
    const sheet = ensureDailyExpensesSheet();
    if (sheet) {
      const currentBranch = sheet.getRange(rowIndex, 3).getValue();
      if (currentBranch !== lockedBranch) {
        throw new Error('You can only delete expenses for your branch.');
      }
    }
  }

  const ss = getSS();
  const sheet = ss.getSheetByName('Daily Expenses');
  if (!sheet) throw new Error("Create a sheet named 'Daily Expenses'");

  const rowValues = sheet.getRange(rowIndex, 2, 1, 2).getValues()[0] || [];
  const date = rowValues[0];
  const branch = String(rowValues[1] || '').trim();

  sheet.deleteRow(rowIndex);

  const cache = CacheService.getScriptCache();
  getDailyExpenseCacheKeys(date, branch).forEach(key => cache.remove(key));
  return true;
}

function getDailyCashFlow(date, branch) {
  const lockedBranch = getLockedBranchForCurrentUser();
  if (lockedBranch) {
    branch = lockedBranch;
  }

  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = `cashflow_${date}_${branch || 'all'}`;
    const cached = cache.get(cacheKey);
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (parseErr) {
        cache.remove(cacheKey);
      }
    }

    const targetDate = new Date(date);
    const startOfDay = new Date(targetDate);
    startOfDay.setHours(0, 0, 0, 0);
    const endOfDay = new Date(targetDate);
    endOfDay.setHours(23, 59, 59, 999);

    // Get cash received from procedures/payments (only cash method)
    const cashReceived = getCashPaymentsForDate(
      Utilities.formatDate(startOfDay, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      Utilities.formatDate(endOfDay, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      branch
    );

    // Get daily expenses
    const expenses = getDailyExpenses(
      Utilities.formatDate(startOfDay, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      Utilities.formatDate(endOfDay, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      branch
    );

    const totalExpenses = expenses.reduce((sum, exp) => sum + exp.amount, 0);

    // Calculate cash on hand (cash received - expenses)
    const cashOnHand = cashReceived - totalExpenses;

    const result = {
      date: Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'MM/dd/yyyy'),
      cashReceived: cashReceived,
      totalExpenses: totalExpenses,
      cashOnHand: cashOnHand,
      expenses: expenses
    };

    safeCachePut(cache, cacheKey, result, 300);
    return result;
  } catch (e) {
    console.error('getDailyCashFlow error: ' + e);
    throw new Error('Could not calculate daily cash flow. ' + String(e.message || e));
  }
}

function getCashPaymentsForDate(startDate, endDate, branch) {
  const lockedBranch = getLockedBranchForCurrentUser();
  if (lockedBranch) {
    branch = lockedBranch;
  }

  const ss = getSS();
  const paySheet = ss.getSheetByName('Payments');
  const visitSheet = ss.getSheetByName('Visits');

  if (!paySheet || !visitSheet) return 0;

  const parseMoney = function (v) {
    return Number(String(v).replace(/[₱,]/g, '').trim()) || 0;
  };

  const vLast = visitSheet.getLastRow();
  const visitMap = {};
  if (vLast > 1) {
    const vData = visitSheet.getRange(2, 1, vLast - 1, 4).getValues();
    vData.forEach(function (row) {
      if (!row[0]) return;
      visitMap[String(row[0])] = {
        date: new Date(row[2]),
        branch: String(row[3] || '')
      };
    });
  }

  const start = new Date(startDate);
  const end = new Date(endDate);
  end.setHours(23, 59, 59, 999);

  const branchFilter = branch ? String(branch).trim() : '';

  const payLast = paySheet.getLastRow();
  let totalCashAmount = 0;
  let totalCashTip = 0;

  if (payLast >= 2) {
    const payData = paySheet.getRange(2, 1, payLast - 1, 5).getValues();
    payData.forEach(function (row) {
      const visitId = String(row[0] || '').trim();
      if (!visitId) return;
      const info = visitMap[visitId];
      if (!info) return;

      const vd = info.date;
      if (isNaN(vd.getTime())) return;
      if (vd < start || vd > end) return;
      if (branchFilter && String(info.branch) !== branchFilter) return;

      // Only count payments made with Cash method
      const method = String(row[2] || '').trim();
      if (method !== 'Cash') return;

      const amt = parseMoney(row[1]);
      const tip = parseMoney(row[3]);

      totalCashAmount += amt;
      totalCashTip += tip;
    });
  }

  return totalCashAmount + totalCashTip;
}
