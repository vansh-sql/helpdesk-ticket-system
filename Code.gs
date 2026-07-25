// ================================================================
//  Vansh Tech Solutions HelpDesk v6 FINAL — Bug Free
//  FEATURES:
//    1. Employee login → Naam + Email + Department (validated)
//    2. Dept passwords → IT, HR, Admin, DME alag alag
//    3. Dept staff → sirf apne dept ke tickets dekhe + close kare
//    4. IT → saari tickets dekhe
//    5. Email notification → Resolved + Closed pe auto email
//    6. Employee stats → Total, Open, InProgress, Resolved
//    7. Search + Filter tickets
//    8. XSS protection (esc() helper)
// ================================================================

// ✅ SIRF YAHI EK LINE CHANGE HUI HAI — NAYA SHEET ID
var SHEET_ID      = '1te1_-y6Q-ZVhTHto0fK3RbemBHgzJHO8ZPX8UylsAv4';
var TICKETS_SHEET = 'Tickets';
var ARCHIVE_SHEET = 'Tickets_Archive';
var COUNTER_SHEET = 'Config';
var TICKET_HEADERS = [
  'Ticket ID','Title','Department','Category','Priority',
  'Status','By','By Dept','Email','Description','Date','Due','Updated','Remarks'
];
var ACTIVE_STATUSES = { 'Open': true, 'Seen': true, 'In Progress': true };
var CLOSED_STATUSES = { 'Resolved': true, 'Closed': true };
var APP_NAME = 'Vansh Tech Solutions';
var MASTER_ADMIN_DEFAULT_PASSWORD = 'master@1234';
var SESSION_TTL_SECONDS = 21600;
var LOCK_RETRY_ATTEMPTS = 5;
var LOCK_WAIT_MS = 8000;

// ── Department Passwords — yahan se change karo ─────────────────
var DEPT_PASSWORDS = {
  'IT'    : 'it@1234',
  'HR'    : 'hr@1234',
  'Admin' : 'admin@1234',
  'DME'   : 'dme@1234'
};
// NOTE: ACCOUNT DEP, CRM DEP, SUPPLY TEAM = sirf ticket destination departments hain, inke login nahi hain

function cleanText_(s) {
  return String(s || '')
    .replace(/[\u200B-\u200D\uFEFF\u00A0]/g, '')
    .trim().replace(/\s+/g, ' ').toLowerCase();
}

function getMasterAdminPassword_() {
  return PropertiesService.getScriptProperties().getProperty('MASTER_ADMIN_PASSWORD') || MASTER_ADMIN_DEFAULT_PASSWORD;
}

function isMasterAdmin_(pass) {
  return String(pass || '') === getMasterAdminPassword_();
}

function createSession_(payload) {
  var token = Utilities.getUuid();
  payload = payload || {};
  payload.createdAt = new Date().getTime();
  CacheService.getScriptCache().put('sess:' + token, JSON.stringify(payload), SESSION_TTL_SECONDS);
  return token;
}

function hashPassword_(str) {
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(str));
  var txtHash = '';
  for (var i = 0; i < rawHash.length; i++) {
    var hashVal = rawHash[i];
    if (hashVal < 0) hashVal += 256;
    if (hashVal.toString(16).length == 1) txtHash += '0';
    txtHash += hashVal.toString(16);
  }
  return txtHash;
}

function logAudit_(action) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('Audit_Logs');
    if (!sheet) {
      sheet = ss.insertSheet('Audit_Logs');
      sheet.appendRow(["Timestamp", "Action"]);
      sheet.getRange("A1:B1").setFontWeight("bold");
    }
    sheet.appendRow([new Date(), action]);
  } catch(e) {}
}

function getSession_(token) {
  if (!token) return null;
  var raw = CacheService.getScriptCache().get('sess:' + token);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch(e) {
    return null;
  }
}

function isMasterAdminAuth_(tokenOrPass) {
  var session = getSession_(tokenOrPass);
  return (session && session.role === 'MASTER') || isMasterAdmin_(tokenOrPass);
}

function sleepBackoff_(attempt) {
  Utilities.sleep(Math.min(2500, 250 * Math.pow(2, attempt)));
}

function withScriptLockRetry_(workFn) {
  var lastErr = null;
  for (var attempt = 0; attempt < LOCK_RETRY_ATTEMPTS; attempt++) {
    var lock = LockService.getScriptLock();
    var locked = false;
    try {
      locked = lock.tryLock(LOCK_WAIT_MS);
      if (!locked) throw new Error('Server busy, retrying...');
      return workFn();
    } catch(e) {
      lastErr = e;
      if (attempt === LOCK_RETRY_ATTEMPTS - 1) break;
      sleepBackoff_(attempt);
    } finally {
      if (locked) lock.releaseLock();
    }
  }
  throw lastErr || new Error('Server busy. Please try again.');
}

function getSettingSheet_(ss) {
  var allSheets = ss.getSheets();
  for (var s = 0; s < allSheets.length; s++) {
    var sn = allSheets[s].getName().trim().toLowerCase();
    if (sn === 'employee_master' || sn === 'setting' || sn === 'settings') return allSheets[s];
  }
  return null;
}

function getOrCreateMappingSheet_(ss) {
  var mapSheet = ss.getSheetByName('Mapping_Rules');
  if (!mapSheet) {
    mapSheet = ss.insertSheet('Mapping_Rules');
    mapSheet.appendRow(['Designation', 'Allowed Departments (comma separated)']);
    mapSheet.getRange(1,1,1,2).setBackground('#3d6fff').setFontColor('#fff').setFontWeight('bold');
    mapSheet.appendRow(['MDO', 'IT, HR, Admin, DME, ACCOUNT DEP, CRM DEP, SUPPLY TEAM']);
    mapSheet.appendRow(['MD', 'IT, HR, Admin, DME, ACCOUNT DEP, CRM DEP, SUPPLY TEAM']);
    mapSheet.appendRow(['GT', 'IT, HR, ACCOUNT DEP']);
    mapSheet.appendRow(['EXPORT', 'IT, HR, SUPPLY TEAM, ACCOUNT DEP']);
    mapSheet.appendRow(['CRM', 'IT, CRM DEP, HR']);
    mapSheet.appendRow(['ACCOUNTS', 'IT, ACCOUNT DEP, HR']);
    mapSheet.appendRow(['HR', 'IT, HR, Admin']);
    mapSheet.appendRow(['MT', 'IT, HR, Admin, DME']);
    mapSheet.appendRow(['SMT', 'IT, HR, Admin, DME, SUPPLY TEAM']);
    SpreadsheetApp.flush();
  }
  return mapSheet;
}

function ensureTicketSheet_(sheet) {
  if (!sheet) return null;
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(TICKET_HEADERS);
    sheet.getRange(1,1,1,TICKET_HEADERS.length).setBackground('#3d6fff').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    return sheet;
  }

  var headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), TICKET_HEADERS.length)).getValues()[0];
  for (var i = 0; i < TICKET_HEADERS.length; i++) {
    if (String(headers[i] || '').trim() !== TICKET_HEADERS[i]) {
      sheet.getRange(1, i + 1).setValue(TICKET_HEADERS[i]);
    }
  }
  sheet.getRange(1,1,1,TICKET_HEADERS.length).setBackground('#3d6fff').setFontColor('#fff').setFontWeight('bold');
  sheet.setFrozenRows(1);
  return sheet;
}

function normalizeTicketRow_(row) {
  var safeRow = row || [];
  return {
    id       : String(safeRow[0]  || ''),
    title    : String(safeRow[1]  || ''),
    dept     : String(safeRow[2]  || ''),
    category : String(safeRow[3]  || ''),
    priority : String(safeRow[4]  || 'Low'),
    status   : String(safeRow[5]  || 'Open'),
    by       : String(safeRow[6]  || ''),
    byDept   : String(safeRow[7]  || ''),
    email    : String(safeRow[8]  || ''),
    desc     : String(safeRow[9]  || ''),
    date     : formatTicketDate_(safeRow[10]),
    due      : String(safeRow[11] || ''),
    updated  : formatTicketDate_(safeRow[12]),
    remarks  : String(safeRow[13] || '')
  };
}

function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function publicError_(e) {
  var msg = String((e && e.message) || e || 'Server error');
  if (/sheet|spreadsheet|range|permission|exception|service invoked too many times/i.test(msg)) {
    return 'Server process complete nahi ho paya. Thodi der baad retry karein.';
  }
  return msg;
}

function formatTicketDate_(value) {
  if (!value) return '';
  var dateObj = null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    dateObj = value;
  } else {
    var text = String(value || '').trim();
    if (!text) return '';
    if (text.indexOf('GMT') === -1) return text;
    var parsed = new Date(text);
    if (!isNaN(parsed.getTime())) dateObj = parsed;
  }
  return dateObj ? Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : String(value || '');
}

function getSheetRows_(sheet) {
  if (!sheet || sheet.getLastRow() <= 1) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, TICKET_HEADERS.length).getValues();
}

function sortTicketsDesc_(tickets) {
  return tickets.sort(function(a, b) {
    var aKey = String(a.updated || a.date || a.id || '');
    var bKey = String(b.updated || b.date || b.id || '');
    return bKey.localeCompare(aKey);
  });
}

// ================================================================
//  SETUP — Pehli baar run karo
// ================================================================
function setupSheets() {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  var tSheet = ss.getSheetByName(TICKETS_SHEET);
  if (!tSheet) tSheet = ss.insertSheet(TICKETS_SHEET);
  ensureTicketSheet_(tSheet);

  var aSheet = ss.getSheetByName(ARCHIVE_SHEET);
  if (!aSheet) aSheet = ss.insertSheet(ARCHIVE_SHEET);
  ensureTicketSheet_(aSheet);

  var empSheet = ss.getSheetByName("Employee_Master");
  if (!empSheet) {
    empSheet = ss.insertSheet("Employee_Master");
    empSheet.appendRow(['EMP Code', 'Name', 'Designation', 'Password', 'Email', 'Status']);
    empSheet.getRange(1, 1, 1, 6).setBackground('#3d6fff').setFontColor('#fff').setFontWeight('bold');
    empSheet.appendRow(['ADMIN01', 'Super Admin', 'MD', hashPassword_('admin123'), 'admin@vanshtech.com', 'Active']);
  }

  var mapSheet = ss.getSheetByName("Email_mapping");
  if (!mapSheet) {
    mapSheet = ss.insertSheet("Email_mapping");
    mapSheet.appendRow(['Department', 'Notification Email']);
    mapSheet.getRange(1, 1, 1, 2).setBackground('#3d6fff').setFontColor('#fff').setFontWeight('bold');
    mapSheet.appendRow(['IT', 'it@vanshtech.com']);
    mapSheet.appendRow(['HR', 'hr@vanshtech.com']);
  }

  var cSheet = ss.getSheetByName(COUNTER_SHEET);
  if (!cSheet) {
    cSheet = ss.insertSheet(COUNTER_SHEET);
    cSheet.appendRow(['Key','Value']);
    cSheet.appendRow(['ticket_counter', 1]);
  } else {
    var cData = cSheet.getDataRange().getValues();
    var found = false;
    for (var i = 0; i < cData.length; i++) {
      if (String(cData[i][0]).trim() === 'ticket_counter') { found = true; break; }
    }
    if (!found) cSheet.appendRow(['ticket_counter', 1]);
  }

  Logger.log('Setup v6 FINAL done!');
  return 'Setup done!';
}

// ================================================================
//  AUTO SETUP — har call pe sheets verify karta hai
// ================================================================
function getOrCreateSheets() {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  var tSheet = ss.getSheetByName(TICKETS_SHEET);
  if (!tSheet) tSheet = ss.insertSheet(TICKETS_SHEET);
  ensureTicketSheet_(tSheet);

  var aSheet = ss.getSheetByName(ARCHIVE_SHEET);
  if (!aSheet) aSheet = ss.insertSheet(ARCHIVE_SHEET);
  ensureTicketSheet_(aSheet);

  var cSheet = ss.getSheetByName(COUNTER_SHEET);
  if (!cSheet) {
    cSheet = ss.insertSheet(COUNTER_SHEET);
    cSheet.appendRow(['Key','Value']);
    cSheet.appendRow(['ticket_counter', 1]);
    SpreadsheetApp.flush();
  } else {
    var cData = cSheet.getDataRange().getValues();
    var found = false;
    for (var i = 0; i < cData.length; i++) {
      if (String(cData[i][0]).trim() === 'ticket_counter') { found = true; break; }
    }
    if (!found) {
      cSheet.appendRow(['ticket_counter', 1]);
      SpreadsheetApp.flush();
    }
  }

  return { ss: ss, tSheet: tSheet, aSheet: aSheet, cSheet: cSheet };
}

// ================================================================
//  doGet
// ================================================================
function doGet() {
  var tmpl = HtmlService.createTemplate(getHTML());
  var html = tmpl.evaluate();
  html.setTitle(APP_NAME);
  html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}

// ================================================================
//  SERVER FUNCTIONS
// ================================================================

function serverCheckDeptPassword(dept, pass) {
  var expected = DEPT_PASSWORDS[dept];
  if (!expected) return false;
  return pass === expected;
}

function serverCheckMasterPassword(pass) {
  return isMasterAdmin_(pass);
}

function serverDeptLogin(dept, pass) {
  try {
    if (!serverCheckDeptPassword(dept, pass)) return { ok: false, msg: 'Galat password!' };
    var token = createSession_({ role: 'staff', name: dept + ' Staff', dept: dept, email: '' });
    return { ok: true, token: token, name: dept + ' Staff', dept: dept, email: '' };
  } catch(e) {
    return { ok: false, msg: publicError_(e) };
  }
}

function serverMasterLogin(pass) {
  try {
    if (!isMasterAdmin_(pass)) return { ok: false, msg: 'Galat Super Admin password!' };
    var token = createSession_({ role: 'MASTER', name: 'Super Admin', dept: 'Master', email: '' });
    return { ok: true, token: token, name: 'Super Admin', dept: 'Master', email: '' };
  } catch(e) {
    return { ok: false, msg: publicError_(e) };
  }
}

// ----------------------------------------------------------------
//  Employee Login: Setting sheet se validate karo
//  Sheet: Col A = EMP Code, Col B = Name, Col C = Designation, Col D = Password
// ----------------------------------------------------------------
function serverEmpLogin(chkCode, chkPass, userEmail) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);

    // Setting sheet dhundo — case-insensitive fallback ke saath
    var setSheet = getSettingSheet_(ss);
    if (!setSheet) {
      return { ok: false, msg: 'Employee master setup complete nahi hai. Admin se contact karein.' };
    }

    // Mapping_Rules auto-create
    var mapSheet = getOrCreateMappingSheet_(ss);

    var inCode = cleanText_(chkCode);
    var inPass = String(chkPass || '').trim();

    if (!inCode) return { ok: false, msg: 'Employee Code khali hai!' };
    if (!inPass) return { ok: false, msg: 'Password khali hai!' };

    var sData = setSheet.getDataRange().getValues();
    var empFound = false;
    var empName  = '';
    var empDesig = '';
    var empCode  = '';
    var empEmail = '';
    var codeFound = false;

    for (var i = 1; i < sData.length; i++) {
      var rowCode = cleanText_(sData[i][0]); // Col A = EMP Code
      var rowPass = String(sData[i][3] || '').trim(); // Col D = Password

      if (!rowCode) continue;

      // Step 1: EMP Code exact match
      if (rowCode !== inCode) continue;
      codeFound = true;

      // Step 2: Name — contains check (dono taraf se)
      if (!rowPass) {
        return { ok: false, msg: 'Is employee ka password setup nahi mila. Admin se update karwayein.' };
      }
      if (rowPass !== inPass && rowPass !== hashPassword_(inPass)) {
        return { ok: false, msg: 'Employee password galat hai!' };
      }

      // Step 3: Designation — contains check
      

      empFound = true;
      empName  = String(sData[i][1] || '').trim();
      empDesig = String(sData[i][2] || '').trim();
      empCode  = String(sData[i][0] || '').trim();
      empEmail = String(sData[i][4] || '').trim(); // Col E = Email Id
      break;
    }

    if (!codeFound) {
      return { ok: false, msg: 'Employee Code "' + chkCode.trim() + '" master record mein nahi mila!' };
    }
    if (!empFound) {
      return { ok: false, msg: 'Login fail hua. Employee Code ya Password check karein.' };
    }
    if (!empEmail || empEmail.indexOf('@') < 1) {
      return { ok: false, msg: 'Is employee ka Email Id valid nahi mila. Admin se update karwayein.' };
    }

    // Mapping_Rules se allowed departments dhundo (contains match)
    var allowed = [];
    var mData = mapSheet.getDataRange().getValues();
    var dLow = empDesig.toLowerCase();
    for (var j = 1; j < mData.length; j++) {
      var mDesig = cleanText_(mData[j][0]);
      if (mDesig && dLow && (dLow.indexOf(mDesig) > -1 || mDesig.indexOf(dLow) > -1)) {
        var parts = String(mData[j][1] || '').split(',');
        for (var k = 0; k < parts.length; k++) {
          var d = parts[k].trim();
          if (d) allowed.push(d);
        }
        break;
      }
    }
    if (allowed.length === 0) {
      allowed = ['IT', 'HR', 'Admin', 'DME', 'ACCOUNT DEP', 'CRM DEP', 'SUPPLY TEAM'];
    }

    var token = createSession_({
      role: 'employee',
      name: empName,
      dept: empDesig,
      code: empCode,
      email: empEmail,
      allowed: allowed
    });

    return { ok: true, token: token, name: empName, dept: empDesig, code: empCode, email: empEmail, allowed: allowed };

  } catch(e) {
    return { ok: false, msg: 'Server error: ' + publicError_(e) };
  }
}

function serverAdminGetData(masterPass) {
  try {
    if (!isMasterAdminAuth_(masterPass)) return { ok: false, msg: 'Super Admin session expire ho gaya. Dobara login karo.' };
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var setSheet = getSettingSheet_(ss);
    if (!setSheet) return { ok: false, msg: 'Employee master setup complete nahi hai.' };
    var mapSheet = getOrCreateMappingSheet_(ss);

    var employees = [];
    var sData = setSheet.getDataRange().getValues();
    for (var i = 1; i < sData.length; i++) {
      if (!String(sData[i][0] || '').trim()) continue;
      employees.push({
        code: String(sData[i][0] || '').trim(),
        name: String(sData[i][1] || '').trim(),
        designation: String(sData[i][2] || '').trim(),
        password: String(sData[i][3] || '').trim(),
        email: String(sData[i][4] || '').trim()
      });
    }

    var mappings = [];
    var mData = mapSheet.getDataRange().getValues();
    for (var j = 1; j < mData.length; j++) {
      if (!String(mData[j][0] || '').trim()) continue;
      mappings.push({
        designation: String(mData[j][0] || '').trim(),
        allowed: String(mData[j][1] || '').trim()
      });
    }

    return { ok: true, employees: employees, mappings: mappings };
  } catch(e) {
    return { ok: false, msg: publicError_(e) };
  }
}

function serverAdminAddEmployee(masterPass, emp) {
  try {
    if (!isMasterAdminAuth_(masterPass)) return { ok: false, msg: 'Super Admin session expire ho gaya. Dobara login karo.' };
    emp = emp || {};
    var code = String(emp.code || '').trim();
    var name = String(emp.name || '').trim();
    var designation = String(emp.designation || '').trim();
    var password = String(emp.password || '').trim();
    var email = String(emp.email || '').trim();

    if (!code || !name || !designation || !password || !email) {
      return { ok: false, msg: 'Employee Code, Name, Designation, Password aur Email Id sab required hain.' };
    }
    if (email.indexOf('@') < 1 || email.indexOf('.') < 1) {
      return { ok: false, msg: 'Valid Email Id bharo.' };
    }

    return withScriptLockRetry_(function() {
      var ss = SpreadsheetApp.openById(SHEET_ID);
      var setSheet = getSettingSheet_(ss);
    if (!setSheet) return { ok: false, msg: 'Employee master setup complete nahi hai.' };

      var data = setSheet.getDataRange().getValues();
      var cleanCode = cleanText_(code);
      for (var i = 1; i < data.length; i++) {
        if (cleanText_(data[i][0]) === cleanCode) {
          return { ok: false, msg: 'Ye Employee Code pehle se master record mein hai.' };
        }
      }

      var hashedPassword = hashPassword_(password);
      setSheet.appendRow([code, name, designation, hashedPassword, email]);
      SpreadsheetApp.flush();
      logAudit_("Super Admin added new employee: " + code + " (" + name + ")");
      return { ok: true, employee: { code: code, name: name, designation: designation, password: password, email: email } };
    });
  } catch(e) {
    return { ok: false, msg: publicError_(e) };
  }
}

function serverAdminSaveMapping(masterPass, designation, allowedCsv) {
  try {
    if (!isMasterAdminAuth_(masterPass)) return { ok: false, msg: 'Super Admin session expire ho gaya. Dobara login karo.' };
    designation = String(designation || '').trim();
    allowedCsv = String(allowedCsv || '').trim();
    if (!designation || !allowedCsv) {
      return { ok: false, msg: 'Designation aur allowed departments required hain.' };
    }

    var parts = allowedCsv.split(',');
    var cleanParts = [];
    for (var p = 0; p < parts.length; p++) {
      var dept = parts[p].trim();
      if (dept) cleanParts.push(dept);
    }
    if (!cleanParts.length) return { ok: false, msg: 'Kam se kam ek department add karo.' };
    allowedCsv = cleanParts.join(', ');

    return withScriptLockRetry_(function() {
      var ss = SpreadsheetApp.openById(SHEET_ID);
      var mapSheet = getOrCreateMappingSheet_(ss);
      var data = mapSheet.getDataRange().getValues();
      var cleanDesignation = cleanText_(designation);
      for (var i = 1; i < data.length; i++) {
        if (cleanText_(data[i][0]) === cleanDesignation) {
          mapSheet.getRange(i + 1, 1, 1, 2).setValues([[designation, allowedCsv]]);
          SpreadsheetApp.flush();
          return { ok: true, mapping: { designation: designation, allowed: allowedCsv } };
        }
      }

      mapSheet.appendRow([designation, allowedCsv]);
      SpreadsheetApp.flush();
      return { ok: true, mapping: { designation: designation, allowed: allowedCsv } };
    });
  } catch(e) {
    return { ok: false, msg: publicError_(e) };
  }
}



function serverGetTickets(query) {
  try {
    query = query || {};
    var session = getSession_(query.token);
    if (!session || !session.role) {
      return { ok: false, msg: 'Session expire ho gaya. Dobara login karo.' };
    }
    var sheets = getOrCreateSheets();
    var liveRows = getSheetRows_(sheets.tSheet);
    var archiveRows = getSheetRows_(sheets.aSheet);
    var role = session.role;
    var dept = String(session.dept || '');
    var email = String(session.email || '').trim().toLowerCase();
    var name = String(session.name || '').trim().toLowerCase();
    var tickets = [];

    function isMine_(ticket) {
      var tEmail = String(ticket.email || '').trim().toLowerCase();
      var tName = String(ticket.by || '').trim().toLowerCase();
      return (email && tEmail === email) || (!email && name && tName === name);
    }

    if (role === 'employee') {
      liveRows.concat(archiveRows).forEach(function(row) {
        var ticket = normalizeTicketRow_(row);
        if (ticket.id && isMine_(ticket)) tickets.push(ticket);
      });
    } else if (role === 'staff') {
      liveRows.forEach(function(row) {
        var ticket = normalizeTicketRow_(row);
        if (!ticket.id) return;
        if (dept === 'IT' || ticket.dept === dept) tickets.push(ticket);
      });
    } else {
      return { ok: false, msg: 'Tickets access allowed nahi hai.' };
    }

    sortTicketsDesc_(tickets);
    return { ok: true, tickets: tickets };
  } catch(e) {
    return { ok: false, msg: publicError_(e) };
  }
}

function serverAddTicket(authToken, title, dept, category, priority, by, byDept, email, desc, due, b64, fname) {
  try {
    var session = getSession_(authToken);
    if (!session || !session.role) return { ok: false, msg: 'Session expire ho gaya. Dobara login karo.' };

    title = String(title || '').trim();
    dept = String(dept || '').trim();
    category = String(category || '').trim();
    priority = String(priority || 'Low').trim();
    desc = String(desc || '').trim();
    due = String(due || '').trim();

    var validPriorities = { Low: true, Medium: true, High: true, Critical: true };
    if (!title || !dept || !category || !desc) return { ok: false, msg: 'Saare required fields bharo.' };
    if (!validPriorities[priority]) return { ok: false, msg: 'Invalid priority' };

    if (session.role === 'employee') {
      var allowed = session.allowed || [];
      var okDept = false;
      for (var a = 0; a < allowed.length; a++) {
        if (String(allowed[a]).trim() === dept) { okDept = true; break; }
      }
      if (!okDept) return { ok: false, msg: 'Aap is department ko ticket nahi bhej sakte.' };
    } else if (session.role !== 'staff') {
      return { ok: false, msg: 'Ticket create access allowed nahi hai.' };
    }

    var submitterName = String(session.name || by || '').trim();
    var submitterDept = String(session.dept || byDept || '').trim();
    var submitterEmail = String(session.email || '').trim();

    return withScriptLockRetry_(function() {
      var sheets  = getOrCreateSheets();
      var tSheet  = sheets.tSheet;
      var cSheet  = sheets.cSheet;

      var cData   = cSheet.getDataRange().getValues();
      var counter = 1;
      var cRowIdx = -1;

      for (var i = 0; i < cData.length; i++) {
        if (String(cData[i][0]).trim() === 'ticket_counter') {
          counter  = Number(cData[i][1]) || 1;
          cRowIdx  = i + 1;
          break;
        }
      }

      if (cRowIdx === -1) {
        cSheet.appendRow(['ticket_counter', 1]);
        cData   = cSheet.getDataRange().getValues();
        counter = 1;
        for (var j = 0; j < cData.length; j++) {
          if (String(cData[j][0]).trim() === 'ticket_counter') {
            cRowIdx = j + 1;
            break;
          }
        }
      }

      var id  = 'TKT-' + ('0000' + counter).slice(-4);
      var dt = new Date();
      var now = Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
      
      if (!due) {
        var hrs = priority === 'Critical' ? 2 : (priority === 'High' ? 4 : (priority === 'Medium' ? 24 : 48));
        dt.setHours(dt.getHours() + hrs);
        due = Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
      }

      var attachmentUrl = '';
      if (b64 && fname) {
        try {
          var folders = DriveApp.getFoldersByName('HelpDesk_Attachments');
          var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('HelpDesk_Attachments');
          folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          var blob = Utilities.newBlob(Utilities.base64Decode(b64), 'application/octet-stream', fname);
          attachmentUrl = folder.createFile(blob).getUrl();
          desc = desc + '\n\nAttachment: ' + attachmentUrl;
        } catch(e) { Logger.log('File upload failed: ' + e); }
      }

      tSheet.appendRow([
        id, title, dept, category, priority,
        'Open', submitterName, submitterDept, submitterEmail, desc, now, due, now, ''
      ]);

      if (cRowIdx > 0) {
        cSheet.getRange(cRowIdx, 2).setValue(counter + 1);
        SpreadsheetApp.flush();
      }

      try {
        var mapSheet = sheets.ss ? sheets.ss.getSheetByName("Email_mapping") : SpreadsheetApp.openById(SHEET_ID).getSheetByName("Email_mapping");
        if (mapSheet) {
          var mData = mapSheet.getDataRange().getValues();
          var notifyEmail = '';
          for (var m = 1; m < mData.length; m++) {
            if (String(mData[m][0]).trim().toLowerCase() === dept.toLowerCase()) {
              notifyEmail = String(mData[m][1]).trim();
              break;
            }
          }
          if (notifyEmail) {
            sendTicketNotification(notifyEmail, dept + " Team", id, title, dept, "Open", "", priority, due);
          }
        }
      } catch(e) {}

      return { ok: true, id: id, date: now, att: attachmentUrl };
    });
  } catch(e) {
    return { ok: false, msg: publicError_(e) };
  }
}

function serverAddRemark(authToken, ticketId, remark) {
  try {
    var session = getSession_(authToken);
    if (!session) return { ok: false, msg: 'Session expire ho gaya.' };
    remark = String(remark || '').trim();
    if (!remark) return { ok: false, msg: 'Remark khali hai.' };
    return withScriptLockRetry_(function() {
      var sheets = getOrCreateSheets();
      var data = sheets.tSheet.getDataRange().getValues();
      var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(ticketId)) {
          var existing = String(data[i][13] || '').trim();
          var finalRem = existing ? existing + '\n\n[' + now + '] ' + session.name + ': ' + remark : '[' + now + '] ' + session.name + ': ' + remark;
          sheets.tSheet.getRange(i + 1, 14).setValue(finalRem);
          sheets.tSheet.getRange(i + 1, 13).setValue(now);
          logAudit_(session.name + " (" + session.role + ") added remark to ticket " + ticketId);
          SpreadsheetApp.flush();
          return { ok: true, remarks: finalRem, updated: now };
        }
      }
      return { ok: false, msg: 'Ticket nahi mila ya close ho chuka hai.' };
    });
  } catch(e) { return { ok: false, msg: publicError_(e) }; }
}

function serverUpdateStatus(authToken, ticketId, newStatus, remarks) {
  try {
    var session = getSession_(authToken);
    if (!session || session.role !== 'staff') return { ok: false, msg: 'Status update access allowed nahi hai.' };

    return withScriptLockRetry_(function() {
      var sheets = getOrCreateSheets();
      var tSheet = sheets.tSheet;
      var aSheet = sheets.aSheet;
      if (!tSheet || !aSheet) return { ok: false, msg: 'Data setup complete nahi hai.' };
      var validStatuses = { 'Open': true, 'Seen': true, 'In Progress': true, 'Resolved': true, 'Closed': true };
      if (!validStatuses[newStatus]) return { ok: false, msg: 'Invalid status' };

      var data = tSheet.getDataRange().getValues();
      var now  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');

      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(ticketId)) {
          var row = data[i].slice(0, TICKET_HEADERS.length);
          while (row.length < TICKET_HEADERS.length) row.push('');

          if (session.dept !== 'IT' && String(row[2] || '') !== String(session.dept || '')) {
            return { ok: false, msg: 'Aap sirf apne department ke tickets update kar sakte ho.' };
          }

          var existingRemarks = String(row[13] || '').trim();
          var newRemark = String(remarks || '').trim();
          var finalRemarks = existingRemarks;
          if (newRemark) {
            finalRemarks = existingRemarks ? existingRemarks + '\n\n[' + now + '] ' + session.name + ': ' + newRemark : '[' + now + '] ' + session.name + ': ' + newRemark;
          }

          row[5] = newStatus;
          row[12] = now;
          row[13] = finalRemarks;

          if (CLOSED_STATUSES[newStatus]) {
            aSheet.appendRow(row);
            tSheet.deleteRow(i + 1);
          } else {
            tSheet.getRange(i + 1, 1, 1, TICKET_HEADERS.length).setValues([row]);
          }
          logAudit_(session.name + " (" + session.role + ") updated ticket " + ticketId + " status to " + newStatus);
          SpreadsheetApp.flush();

          if (CLOSED_STATUSES[newStatus]) {
            var empEmail = String(row[8] || '');
            var empName  = String(row[6] || '');
            var title    = String(row[1] || '');
            var dept     = String(row[2] || '');
            if (empEmail && empEmail.indexOf('@') > -1) {
              try {
                sendTicketNotification(empEmail, empName, ticketId, title, dept, newStatus, finalRemarks);
              } catch(emailErr) {
                Logger.log('Email error: ' + emailErr.toString());
              }
            }
          }
          return { ok: true };
        }
      }
      return { ok: false, msg: 'Ticket nahi mila' };
    });
  } catch(e) {
    return { ok: false, msg: publicError_(e) };
  }
}

// ================================================================
//  EMAIL NOTIFICATION
// ================================================================
function sendTicketNotification(toEmail, empName, ticketId, title, dept, status, remarks, priority, due) {
  if (!toEmail) return;
  var statusEmoji = status === 'Resolved' ? '✅' : (status === 'Open' ? '🔔' : '🔒');
  var subject = statusEmoji + ' Ticket ' + ticketId + ' ' + status + ' — ' + APP_NAME;

  var body = "";
  try {
    var tpl = HtmlService.createTemplateFromFile('EmailTemplate');
    tpl.ticket = {
      recipientName: empName || 'Team',
      id: ticketId,
      title: title,
      priority: priority || 'Normal',
      dept: dept,
      due: due || '',
      status: status,
      remarks: remarks || ''
    };
    body = tpl.evaluate().getContent();
  } catch(e) {
    // Fallback if template missing
    body = "<p>Ticket " + ticketId + " status is now " + status + ".</p>";
  }

  MailApp.sendEmail({ to: toEmail, subject: subject, htmlBody: body });
}

// ================================================================
//  HTML — Cleanly written, no string escape bugs
// ================================================================
function getHTML() {
  return '<!DOCTYPE html>\n<html>\n<head>\n' +
'<meta charset="UTF-8">\n' +
'<meta name="viewport" content="width=device-width,initial-scale=1">\n' +
'<title>Vansh Tech Solutions</title>\n' +
'<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">\n' +
'<style>\n' +
':root{--bg:#0a0f1c;--s1:#111827;--s2:#1f2937;--b1:rgba(255,255,255,0.16);--b2:rgba(255,255,255,0.28);--tx:#ffffff;--mu:#dbe7ff;--ac:#3b82f6;--ac2:#2563eb;--gn:#10b981;--am:#f59e0b;--rd:#ef4444;--pu:#8b5cf6}\n' +
'#pg-app{--bg:#f8fafc;--s1:#ffffff;--s2:#f1f5f9;--b1:rgba(15,23,42,0.08);--b2:rgba(15,23,42,0.15);--tx:#0f172a;--mu:#64748b;background:var(--bg);color:var(--tx)}\n' +
'html[data-theme="dark"] #pg-app{--bg:#0f172a;--s1:#1e293b;--s2:#334155;--b1:rgba(255,255,255,0.08);--b2:rgba(255,255,255,0.15);--tx:#f8fafc;--mu:#94a3b8}\n' +
'*{box-sizing:border-box;margin:0;padding:0}\n' +
'body{font-family:"Plus Jakarta Sans",sans-serif;background:var(--bg);color:var(--tx);min-height:100vh}\n' +
'.pg{display:none;min-height:100vh}.pg.on{display:flex}\n' +

/* LOGIN */
'#pg-login{align-items:center;justify-content:center;flex-direction:column;position:relative;overflow:hidden;padding:18px 0}\n' +
'.night-sky{position:absolute;inset:0;overflow:hidden;z-index:1;pointer-events:none;background:radial-gradient(circle at 20% 20%,rgba(59,130,246,.34),transparent 28%),radial-gradient(ellipse at 30% 60%,#183470,#0a0f1c 72%)}\n' +
'.night-sky::after{content:"";position:absolute;inset:0;background:radial-gradient(circle at 75% 18%,rgba(255,255,255,.18),transparent 18%),linear-gradient(180deg,rgba(255,255,255,.03),transparent 35%);opacity:.9}\n' +
'.star{position:absolute;background:#fff;border-radius:50%;animation:twinkle 4s infinite;box-shadow:0 0 8px rgba(255,255,255,.85)}\n' +
'@keyframes twinkle{0%,100%{opacity:.35;transform:scale(.9)}50%{opacity:1;transform:scale(1.15)}}\n' +
'.shoot{position:absolute;width:2px;height:2px;background:#fff;border-radius:50%;opacity:0;animation:shoot 5s linear infinite;box-shadow:0 0 6px rgba(100,160,255,.8)}\n' +
'.shoot::before{content:"";position:absolute;top:50%;transform:translateY(-50%);width:90px;height:1px;background:linear-gradient(-90deg,rgba(100,160,255,.9),transparent);right:100%}\n' +
'@keyframes shoot{0%{transform:translate(0,0) rotate(-35deg);opacity:1}100%{transform:translate(-900px,900px) rotate(-35deg);opacity:0}}\n' +

'.lbox{position:relative;z-index:10;width:92%;max-width:460px;padding:40px;background:rgba(17,24,39,.84);border:1px solid rgba(255,255,255,.14);border-radius:24px;backdrop-filter:blur(24px);box-shadow:0 30px 80px rgba(0,0,0,.6);transition:transform .3s ease,box-shadow .3s ease}\n' +
'.lbox:hover{transform:translateY(-2px);box-shadow:0 35px 90px rgba(0,0,0,.7)}\n' +
'.lbrand{display:flex;align-items:center;gap:12px;margin-bottom:22px;padding-bottom:18px;border-bottom:1px solid var(--b1)}\n' +
'.llogo{width:42px;height:42px;border-radius:11px;background:linear-gradient(135deg,var(--ac),var(--ac2));display:flex;align-items:center;justify-content:center;font-size:14px;font-weight:700;color:#fff;flex-shrink:0}\n' +
'.lbname{font-size:15px;font-weight:700;color:var(--tx);line-height:1.2}\n' +
'.lbsub{font-size:10px;color:var(--mu);letter-spacing:.04em;margin-top:2px;text-transform:uppercase}\n' +

'.tabs{display:flex;gap:3px;margin-bottom:20px;background:var(--s2);border-radius:10px;padding:4px}\n' +
'.tab{flex:1;padding:8px 4px;border:none;border-radius:7px;cursor:pointer;font-family:"Plus Jakarta Sans",sans-serif;font-size:11px;font-weight:600;transition:transform .22s ease,background .22s ease,color .22s ease,box-shadow .22s ease;white-space:nowrap;background:transparent;color:var(--mu);position:relative;overflow:visible;display:flex;align-items:center;justify-content:center;gap:5px}\n' +
'.tab-ico{display:inline-flex;align-items:center;justify-content:center;width:14px;height:14px;line-height:1;flex-shrink:0;filter:drop-shadow(0 1px 4px rgba(0,0,0,.45));text-shadow:0 1px 5px rgba(255,255,255,.25)}\n' +
'.tab-label{display:inline-block;line-height:1.1;overflow:hidden;text-overflow:ellipsis}\n' +
'.tab:hover{transform:translateY(-1px);background:rgba(255,255,255,.06);color:#fff}.tab:active{transform:scale(.96)}\n' +
'.tab.on{background:linear-gradient(135deg,var(--ac),var(--ac2));color:#fff;box-shadow:0 8px 22px rgba(59,130,246,.34)}\n' +
'#form-emp,#form-IT,#form-HR,#form-Admin,#form-DME,#form-Master{animation:tabSwipe .24s ease both}\n' +
'@keyframes tabSwipe{0%{opacity:0;transform:translateX(16px) scale(.985)}100%{opacity:1;transform:translateX(0) scale(1)}}\n' +

'.lbox h2{font-size:18px;font-weight:700;margin-bottom:4px}\n' +
'.lbox .sub{color:var(--mu);font-size:13px;margin-bottom:16px;line-height:1.5}\n' +
'.fg{margin-bottom:12px}\n' +
'.fg label{display:block;font-size:10px;font-weight:600;color:var(--mu);text-transform:uppercase;letter-spacing:.08em;margin-bottom:5px}\n' +
'.fg input,.fg select,.fg textarea{width:100%;padding:10px 13px;background:var(--s2);border:1px solid var(--b1);border-radius:9px;color:var(--tx);font-family:"Plus Jakarta Sans",sans-serif;font-size:14px;outline:none;transition:border-color .2s,box-shadow .2s}\n' +
'.fg input:focus,.fg select:focus,.fg textarea:focus{border-color:var(--ac);box-shadow:0 0 0 3px rgba(59,130,246,.25)}\n' +
'.fg select option{background:var(--s2)}\n' +
'.fg input.err{border-color:var(--rd)!important;box-shadow:0 0 0 3px rgba(239,68,68,.25)!important}\n' +
'.errmsg{font-size:11px;color:var(--rd);margin-top:4px;display:none}\n' +
'.btn{width:100%;padding:11px;border-radius:9px;border:none;cursor:pointer;font-family:"Plus Jakarta Sans",sans-serif;font-size:14px;font-weight:600;background:linear-gradient(135deg,var(--ac),var(--ac2));color:#fff;transition:transform .22s ease,box-shadow .22s ease,filter .22s ease;margin-top:4px;box-shadow:0 4px 14px rgba(59,130,246,.3)}\n' +
'.btn:hover{transform:translateY(-2px);box-shadow:0 10px 26px rgba(59,130,246,.42);filter:saturate(1.08)}\n' +
'.btn:active{transform:scale(.98)}\n' +
'.btn:disabled{opacity:.4;cursor:not-allowed;transform:none;box-shadow:none}\n' +
'.btn.sec{background:var(--s2);border:1px solid var(--b2);color:var(--tx);margin-top:0}\n' +
'.dbadge{display:inline-flex;align-items:center;gap:6px;font-size:11px;font-weight:600;padding:5px 12px;border-radius:20px;margin-bottom:14px}\n' +
'.dbadge.IT{background:rgba(61,111,255,.15);color:var(--ac)}\n' +
'.dbadge.HR{background:rgba(240,64,120,.15);color:#f0407a}\n' +
'.dbadge.Admin{background:rgba(29,184,116,.12);color:var(--gn)}\n' +
'.dbadge.DME{background:rgba(240,160,32,.15);color:var(--am)}\n' +

/* APP */
'#pg-app{flex-direction:row}\n' +
'.sb{width:220px;background:var(--s1);border-right:1px solid var(--b1);display:flex;flex-direction:column;min-height:100vh;flex-shrink:0}\n' +
'.sb-brand{padding:18px 16px;border-bottom:1px solid var(--b1);display:flex;align-items:center;gap:10px}\n' +
'.sb-logo{width:34px;height:34px;border-radius:8px;background:linear-gradient(135deg,var(--ac),var(--ac2));display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700;color:#fff;flex-shrink:0}\n' +
'.sb-bname{font-size:13px;font-weight:700;color:var(--tx);line-height:1.2}\n' +
'.sb-bsub{font-size:9px;color:var(--mu);letter-spacing:.04em;text-transform:uppercase}\n' +
'.sb-nav{flex:1;padding:12px 10px;display:flex;flex-direction:column;gap:2px}\n' +
'.ni{display:flex;align-items:center;gap:9px;padding:8px 11px;border-radius:8px;cursor:pointer;font-size:13px;font-weight:500;color:var(--mu);transition:all .15s;user-select:none}\n' +
'.ni:hover{background:var(--s2);color:var(--tx)}.ni.on{background:rgba(61,111,255,.14);color:var(--ac)}\n' +
'.ni svg{width:15px;height:15px;flex-shrink:0}\n' +
'.sb-user{padding:12px 16px;border-top:1px solid var(--b1)}\n' +
'.role-badge{display:inline-flex;align-items:center;gap:5px;font-size:11px;font-weight:600;padding:4px 10px;border-radius:7px;margin-bottom:10px}\n' +
'.r-emp{background:rgba(29,184,116,.1);color:var(--gn)}\n' +
'.r-it{background:rgba(61,111,255,.14);color:var(--ac)}\n' +
'.r-hr{background:rgba(240,64,120,.12);color:#f0407a}\n' +
'.r-dme{background:rgba(240,160,32,.12);color:var(--am)}\n' +
'.r-admin{background:rgba(29,184,116,.12);color:var(--gn)}\n' +
'.r-account{background:rgba(155,92,246,.12);color:var(--pu)}\n' +
'.r-crm{background:rgba(240,64,64,.12);color:var(--rd)}\n' +
'.r-supply{background:rgba(61,111,255,.12);color:#92c5fd}\n' +
'.av{width:30px;height:30px;border-radius:7px;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;color:#fff;flex-shrink:0}\n' +
'.uname{font-size:12px;font-weight:600}.udept{font-size:11px;color:var(--mu);margin-top:1px}\n' +

'.mn{flex:1;display:flex;flex-direction:column;overflow-y:auto}\n' +
'.topbar{padding:16px 26px;border-bottom:1px solid var(--b1);display:flex;align-items:center;justify-content:space-between}\n' +
'.topbar h1{font-size:17px;font-weight:700;letter-spacing:-.3px}\n' +
'.topbar-actions{display:flex;gap:8px}\n' +
'.smbtn{padding:7px 14px;border-radius:7px;border:none;cursor:pointer;font-family:"Plus Jakarta Sans",sans-serif;font-size:12px;font-weight:600;transition:transform .2s ease,opacity .2s,box-shadow .2s}\n' +
'.smbtn.primary{background:linear-gradient(135deg,var(--ac),var(--ac2));color:#fff}\n' +
'.smbtn.primary:hover{opacity:.85}\n' +
'.smbtn.outline{background:transparent;border:1px solid var(--b2);color:var(--tx)}\n' +
'.smbtn.green{background:linear-gradient(135deg,var(--gn),#16a85e);color:#fff}\n' +
'.smbtn.green:hover,.smbtn.primary:hover,.smbtn.outline:hover{opacity:.9;transform:translateY(-1px);box-shadow:0 8px 22px rgba(0,0,0,.22)}\n' +

'.ct{padding:22px 26px}\n' +
'.stats-grid{display:grid;gap:11px;margin-bottom:20px}\n' +
'.stat-card{background:var(--s1);border:1px solid var(--b1);border-radius:12px;padding:15px 17px;transition:transform .22s ease,border-color .22s ease,box-shadow .22s ease}\n' +
'.stat-card:hover{transform:translateY(-3px);border-color:rgba(147,197,253,.34);box-shadow:0 14px 30px rgba(0,0,0,.24)}\n' +
'.stat-label{font-size:10px;color:var(--mu);text-transform:uppercase;letter-spacing:.08em;margin-bottom:6px;font-weight:600}\n' +
'.stat-val{font-size:26px;font-weight:700;letter-spacing:-1px}\n' +
'.c-blue{color:var(--ac)}.c-amber{color:var(--am)}.c-green{color:var(--gn)}.c-text{color:var(--tx)}.c-muted{color:var(--mu)}\n' +

'.sect-title{font-size:11px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.08em;margin-bottom:11px}\n' +
'.tbl{width:100%;border-collapse:collapse}\n' +
'.tbl th{text-align:left;font-size:10px;font-weight:600;color:var(--mu);text-transform:uppercase;letter-spacing:.08em;padding:0 13px 8px}\n' +
'.tbl td{padding:10px 13px;border-top:1px solid var(--b1);font-size:13px}\n' +
'.tbl tr.clickable{transition:transform .18s ease}.tbl tr.clickable:hover{transform:translateX(3px)}.tbl tr.clickable:hover td{background:var(--s2);cursor:pointer}\n' +
'.tid{font-family:"JetBrains Mono",monospace;font-size:11px;color:var(--mu);font-weight:500}\n' +

'.badge{display:inline-block;padding:3px 9px;border-radius:20px;font-size:10px;font-weight:600;letter-spacing:.02em}\n' +
'.p-low{background:rgba(107,114,153,.15);color:var(--mu)}\n' +
'.p-medium{background:rgba(240,160,32,.15);color:var(--am)}\n' +
'.p-high{background:rgba(240,64,64,.18);color:var(--rd)}\n' +
'.p-critical{background:rgba(240,64,64,.3);color:#ff7070}\n' +
'.s-open{background:rgba(61,111,255,.15);color:var(--ac)}\n' +
'.s-seen{background:rgba(139,92,246,.18);color:var(--pu)}\n' +
'.s-inprogress{background:rgba(240,160,32,.15);color:var(--am)}\n' +
'.s-resolved{background:rgba(29,184,116,.15);color:var(--gn)}\n' +
'.s-closed{background:rgba(107,114,153,.15);color:var(--mu)}\n' +
'.d-HR{background:rgba(240,64,120,.15);color:#f0407a}\n' +
'.d-IT{background:rgba(61,111,255,.15);color:var(--ac)}\n' +
'.d-DME{background:rgba(240,160,32,.15);color:var(--am)}\n' +
'.d-Admin{background:rgba(29,184,116,.15);color:var(--gn)}\n' +
'.sla-breached{background:rgba(239,68,68,0.06)!important;border-left:3px solid #ef4444;box-shadow:inset 3px 0 0 #ef4444;}\n' +
'.sla-breached td{color:#b91c1c!important;font-weight:600;}\n' +
'.d-Boss{background:rgba(155,92,246,.15);color:var(--pu)}\n' +
'.d-ACCOUNTDEP{background:rgba(155,92,246,.15);color:var(--pu)}\n' +
'.d-CRMDEP{background:rgba(240,64,64,.15);color:var(--rd)}\n' +
'.d-SUPPLYTEAM{background:rgba(61,111,255,.15);color:#92c5fd}\n' +
'.d-Other{background:rgba(107,114,153,.15);color:var(--mu)}\n' +
'.empty{text-align:center;padding:40px 0;color:var(--mu);font-size:13px}\n' +

'.fbar{display:flex;gap:8px;margin-bottom:14px;flex-wrap:wrap;align-items:center}\n' +
'.fsrch{flex:1;min-width:140px;padding:7px 12px;background:var(--s2);border:1px solid var(--b1);border-radius:7px;color:var(--tx);font-family:"Plus Jakarta Sans",sans-serif;font-size:12px;outline:none}\n' +
'.fsrch:focus{border-color:var(--ac)}\n' +
'.fsel{padding:7px 10px;background:var(--s2);border:1px solid var(--b1);border-radius:7px;color:var(--tx);font-family:"Plus Jakarta Sans",sans-serif;font-size:12px;outline:none;cursor:pointer}\n' +
'.fsel:focus{border-color:var(--ac)}\n' +
'.fsel option{background:var(--s2)}\n' +

'.wc{background:linear-gradient(135deg,rgba(61,111,255,.12),rgba(91,63,255,.07));border:1px solid rgba(61,111,255,.22);border-radius:14px;padding:20px 24px;margin-bottom:18px;transition:transform .22s ease,box-shadow .22s ease,border-color .22s ease}\n' +
'.wc:hover{transform:translateY(-2px);border-color:rgba(147,197,253,.36);box-shadow:0 16px 40px rgba(0,0,0,.24)}\n' +
'.wc h2{font-size:17px;font-weight:700;margin-bottom:4px;letter-spacing:-.3px}\n' +
'.wc p{color:var(--mu);font-size:13px;line-height:1.5}\n' +
'.wc-btn{display:inline-flex;align-items:center;gap:7px;margin-top:12px;padding:10px 20px;border-radius:9px;border:none;cursor:pointer;font-family:"Plus Jakarta Sans",sans-serif;font-size:13px;font-weight:600;background:linear-gradient(135deg,var(--ac),var(--ac2));color:#fff;transition:transform .22s ease,box-shadow .22s ease,opacity .2s}\n' +
'.wc-btn:hover{opacity:.92;transform:translateY(-2px);box-shadow:0 10px 24px rgba(59,130,246,.35)}.wc-btn:active,.smbtn:active{transform:scale(.97)}\n' +

'.fc{background:var(--s1);border:1px solid var(--b1);border-radius:14px;padding:22px;max-width:620px}\n' +
'.admin-hero{display:flex;align-items:center;justify-content:space-between;gap:16px;background:linear-gradient(135deg,rgba(59,130,246,.22),rgba(16,185,129,.10));border:1px solid rgba(147,197,253,.26);border-radius:14px;padding:18px 22px;margin-bottom:18px}\n' +
'.admin-hero h2{font-size:18px;font-weight:700;margin-bottom:4px}.admin-hero p{font-size:12px;color:var(--mu);line-height:1.5}.admin-pill{font-size:11px;font-weight:700;color:#bfdbfe;background:rgba(59,130,246,.18);border:1px solid rgba(147,197,253,.25);border-radius:999px;padding:6px 12px;white-space:nowrap}\n' +
'#vp-master .fc{background:rgba(17,24,39,.92);border-color:rgba(219,234,254,.18);box-shadow:0 16px 40px rgba(0,0,0,.22)}\n' +
'#vp-master .sect-title{color:#f8fbff;font-size:12px;margin-bottom:14px}#vp-master .tbl td,#vp-master .tbl th{color:#f8fbff}#vp-master .tbl td[style]{color:#dbe7ff!important}\n' +
'.g2{display:grid;grid-template-columns:1fr 1fr;gap:13px}\n' +
'.fg textarea{width:100%;padding:10px 13px;background:var(--s2);border:1px solid var(--b1);border-radius:9px;color:var(--tx);font-family:"Plus Jakarta Sans",sans-serif;font-size:14px;outline:none;min-height:90px;resize:vertical;line-height:1.6;transition:border-color .2s,box-shadow .2s}\n' +
'.fg textarea:focus{border-color:var(--ac);box-shadow:0 0 0 3px rgba(61,111,255,.12)}\n' +
'.po{display:flex;gap:7px}\n' +
'.po-opt{flex:1;padding:7px;border-radius:7px;border:1px solid var(--b1);background:var(--s2);cursor:pointer;text-align:center;font-size:11px;font-weight:600;color:var(--mu);transition:all .15s;user-select:none}\n' +
'.po-opt.sel-low{border-color:var(--mu);color:var(--tx);background:rgba(107,114,153,.15)}\n' +
'.po-opt.sel-medium{border-color:var(--am);color:var(--am);background:rgba(240,160,32,.1)}\n' +
'.po-opt.sel-high{border-color:var(--rd);color:var(--rd);background:rgba(240,64,64,.1)}\n' +
'.po-opt.sel-critical{border-color:#ff7070;color:#ff7070;background:rgba(240,64,64,.15)}\n' +
'.form-actions{display:flex;gap:9px;margin-top:18px}\n' +
'.form-actions .btn{width:auto;padding:10px 22px;flex:1}\n' +

'.detail-card{background:var(--s1);border:1px solid var(--b1);border-radius:14px;max-width:700px;overflow:hidden}\n' +
'.detail-head{padding:20px 24px;border-bottom:1px solid var(--b1)}\n' +
'.detail-head h2{font-size:17px;font-weight:700;margin:8px 0 12px;letter-spacing:-.3px}\n' +
'.detail-body{padding:20px 24px}\n' +
'.detail-row{display:flex;gap:28px;margin-bottom:16px;flex-wrap:wrap}\n' +
'.detail-item label{font-size:10px;color:var(--mu);text-transform:uppercase;letter-spacing:.08em;font-weight:600;display:block;margin-bottom:3px}\n' +
'.detail-item span{font-size:13px}\n' +
'.desc-box{background:var(--s2);border-radius:9px;padding:13px;font-size:13px;line-height:1.7;margin:12px 0}\n' +
'.update-box{background:rgba(61,111,255,.07);border:1px solid rgba(61,111,255,.15);border-radius:9px;padding:14px;margin-top:6px}\n' +
'.update-box .hint{font-size:11px;color:var(--mu);margin-top:8px}\n' +
'.dsel{padding:7px 12px;background:var(--s2);border:1px solid var(--b1);border-radius:7px;color:var(--tx);font-family:"Plus Jakarta Sans",sans-serif;font-size:12px;font-weight:500;outline:none;cursor:pointer;flex:1;min-width:140px}\n' +
'.backlink{display:flex;align-items:center;gap:5px;color:var(--mu);font-size:12px;cursor:pointer;margin-bottom:16px;font-weight:500;user-select:none}\n' +
'.backlink:hover{color:var(--tx)}\n' +
'.email-info{display:flex;align-items:center;gap:5px;font-size:12px;color:var(--mu);margin-top:5px}\n' +

'.ov{position:fixed;inset:0;background:rgba(10,13,20,.9);z-index:999;display:none;align-items:center;justify-content:center;flex-direction:column;gap:10px}\n' +
'.ov.on{display:flex}\n' +
'.spinner{width:30px;height:30px;border:2px solid var(--b2);border-top-color:var(--ac);border-radius:50%;animation:spin .7s linear infinite}\n' +
'@keyframes spin{to{transform:rotate(360deg)}}\n' +
'.ov-txt{font-size:13px;color:var(--mu);font-weight:500}\n' +
'.vp{display:none}.vp.on{display:block;animation:viewIn .24s ease both}@keyframes viewIn{0%{opacity:0;transform:translateY(8px)}100%{opacity:1;transform:translateY(0)}}\n' +
'.toast{position:fixed;bottom:22px;right:22px;z-index:9999;background:var(--s1);padding:11px 17px;border-radius:10px;font-size:12px;font-weight:600;display:none;max-width:300px;box-shadow:0 6px 20px rgba(0,0,0,.4);transition:opacity .3s}\n' +

'@media(max-width:768px){\n' +
'#pg-login{justify-content:center;min-height:100dvh;overflow-y:auto;padding:16px 0}\n' +
'.night-sky{position:fixed;background:radial-gradient(circle at 18% 12%,rgba(96,165,250,.45),transparent 30%),radial-gradient(circle at 82% 22%,rgba(255,255,255,.16),transparent 20%),linear-gradient(160deg,#07111f,#10245a 48%,#07111f)}\n' +
'.star{width:2px!important;height:2px!important;opacity:.85}.shoot{box-shadow:0 0 10px rgba(147,197,253,.95)}\n' +
'.lbox{width:calc(100% - 24px);max-width:430px;padding:24px 16px 18px;margin:8px auto;background:rgba(15,23,42,.9);border-color:rgba(219,234,254,.18);max-height:none}\n' +
'.lbrand{margin-bottom:18px;padding-bottom:14px}.tabs{display:grid;grid-template-columns:repeat(3,1fr);gap:6px;border-radius:12px;padding:6px}.tab{min-width:0;padding:9px 5px;font-size:10px;text-align:center}.tab-ico{width:15px;height:15px;font-size:12px}.tab-label{max-width:62px}\n' +
'.lbox h2{font-size:24px!important}.lbox .sub{font-size:12px;line-height:1.55}.fg input,.fg select,.fg textarea{min-height:44px}\n' +
'#pg-app{flex-direction:column}\n' +
'.sb{width:100%;min-height:auto;flex-direction:row;flex-wrap:wrap;padding:10px 14px;align-items:center;border-right:none;border-bottom:1px solid var(--b1)}\n' +
'.sb-brand{border-bottom:none;padding:0}\n' +
'.sb-nav{flex-direction:row;overflow-x:auto;padding:0;gap:5px;flex:1;margin:0 8px}\n' +
'.ni{white-space:nowrap;padding:6px 11px}\n' +
'.sb-user{display:none}\n' +
'.ct{padding:13px}\n' +
'.topbar{padding:13px;flex-direction:column;gap:8px;align-items:flex-start}\n' +
'.stats-grid{grid-template-columns:1fr 1fr!important;gap:8px}\n' +
'.detail-row{gap:13px}\n' +
'.tbl{display:block;overflow-x:auto;white-space:nowrap}\n' +
'.fc{max-width:100%}\n' +
'.g2{grid-template-columns:1fr;gap:10px}\n' +
'.po{flex-wrap:wrap}\n' +
'.po-opt{min-width:calc(50% - 4px)}\n' +
'.fbar{gap:6px}\n' +
'}\n' +
'</style>\n</head>\n<body>\n' +

'<div class="ov" id="ov"><div class="spinner"></div><div class="ov-txt" id="ov-txt">Loading...</div></div>\n' +
'<div class="toast" id="toast"></div>\n' +

/* ===============================================================
   LOGIN PAGE
=============================================================== */
'<div class="pg on" id="pg-login">\n' +
'<div class="night-sky" id="night-sky"></div>\n' +
'<div class="lbox">\n' +

'<div class="lbrand">\n' +
'<div class="llogo">VT</div>\n' +
'<div><div class="lbname">Vansh Tech Solutions</div><div class="lbsub">HelpDesk</div></div>\n' +
'</div>\n' +

'<div class="tabs">\n' +
'<button class="tab on" onclick="switchTab(\'emp\')"><span class="tab-ico">\ud83d\udc64</span><span class="tab-label">Employee</span></button>\n' +
'<button class="tab"    onclick="switchTab(\'IT\')"><span class="tab-ico">\ud83d\udd27</span><span class="tab-label">IT</span></button>\n' +
'<button class="tab"    onclick="switchTab(\'HR\')"><span class="tab-ico">\ud83d\udc65</span><span class="tab-label">HR</span></button>\n' +
'<button class="tab"    onclick="switchTab(\'Admin\')"><span class="tab-ico">\ud83c\udfe2</span><span class="tab-label">Admin</span></button>\n' +
'<button class="tab"    onclick="switchTab(\'DME\')"><span class="tab-ico">\u2699\ufe0f</span><span class="tab-label">DME</span></button>\n' +
'<button class="tab"    onclick="switchTab(\'Master\')"><span class="tab-ico">\ud83d\udee1</span><span class="tab-label">Master</span></button>\n' +
'</div>\n' +

/* Employee form */
'<div id="form-emp">\n' +
'<h2>Welcome 👋</h2><p class="sub">Apna Employee Code aur password daalo. Email automatic detect hoga.</p>\n' +
'<div class="fg"><label>Employee Code *</label>' +
'<input type="text" id="e-code" placeholder="e.g. SSIFPL-1" onkeydown="if(event.key===\'Enter\')empLogin()">' +
'<div class="errmsg" id="err-code">Code bharo</div></div>\n' +
'<div class="fg"><label>Password *</label>' +
'<input type="password" id="e-pass" placeholder="••••••••" onkeydown="if(event.key===\'Enter\')empLogin()">' +
'<div class="errmsg" id="err-pass">Password bharo</div></div>\n' +
'<button class="btn" id="e-btn" onclick="empLogin()">Check & Login →</button>\n' +
'</div>\n' +

/* IT form */
'<div id="form-IT" style="display:none">\n' +
'<div class="dbadge IT">🔧 IT Staff Portal</div>\n' +
'<h2>IT Dashboard</h2><p class="sub">IT password daalo</p>\n' +
'<div class="fg"><label>Password *</label><input type="password" id="pass-IT" placeholder="••••••••" onkeydown="if(event.key===\'Enter\')deptLogin(\'IT\')"></div>\n' +
'<button class="btn" onclick="deptLogin(\'IT\')">IT Access →</button>\n' +
'</div>\n' +

/* HR form */
'<div id="form-HR" style="display:none">\n' +
'<div class="dbadge HR">👥 HR Staff Portal</div>\n' +
'<h2>HR Dashboard</h2><p class="sub">HR password daalo</p>\n' +
'<div class="fg"><label>Password *</label><input type="password" id="pass-HR" placeholder="••••••••" onkeydown="if(event.key===\'Enter\')deptLogin(\'HR\')"></div>\n' +
'<button class="btn" onclick="deptLogin(\'HR\')">HR Access →</button>\n' +
'</div>\n' +

/* Admin form */
'<div id="form-Admin" style="display:none">\n' +
'<div class="dbadge Admin">🏢 Admin Staff Portal</div>\n' +
'<h2>Admin Dashboard</h2><p class="sub">Admin password daalo</p>\n' +
'<div class="fg"><label>Password *</label><input type="password" id="pass-Admin" placeholder="••••••••" onkeydown="if(event.key===\'Enter\')deptLogin(\'Admin\')"></div>\n' +
'<button class="btn" onclick="deptLogin(\'Admin\')">Admin Access →</button>\n' +
'</div>\n' +

/* DME form */
'<div id="form-DME" style="display:none">\n' +
'<div class="dbadge DME">\u2699\ufe0f DME Staff Portal</div>\n' +
'<h2>DME Dashboard</h2><p class="sub">DME password daalo</p>\n' +
'<div class="fg"><label>Password *</label><input type="password" id="pass-DME" placeholder="\u2022\u2022\u2022\u2022\u2022\u2022\u2022\u2022" onkeydown="if(event.key===\'Enter\')deptLogin(\'DME\')"></div>\n' +
'<button class="btn" onclick="deptLogin(\'DME\')">DME Access \u2192</button>\n' +
'</div>\n' +

/* Master Admin form */
'<div id="form-Master" style="display:none">\n' +
'<div class="dbadge Admin">\ud83d\udee1 Super Admin Portal</div>\n' +
'<h2>Super Admin</h2><p class="sub">Employee master aur department mapping manage karne ke liye password daalo.</p>\n' +
'<div class="fg"><label>Password *</label><input type="password" id="pass-Master" placeholder="\u2022\u2022\u2022\u2022\u2022\u2022\u2022\u2022" onkeydown="if(event.key===\'Enter\')masterLogin()"></div>\n' +
'<button class="btn" onclick="masterLogin()">Super Admin Access \u2192</button>\n' +
'</div>\n' +

'</div></div>\n' +

/* ===============================================================
   APP PAGE
=============================================================== */
'<div class="pg" id="pg-app">\n' +
'<div class="sb">\n' +
'<div class="sb-brand">\n' +
'<div class="sb-logo">VT</div>\n' +
'<div><div class="sb-bname">Vansh Tech Solutions</div><div class="sb-bsub">HelpDesk</div></div>\n' +
'</div>\n' +
'<nav class="sb-nav" id="sb-nav"></nav>\n' +
'<div class="sb-user">\n' +
'<div id="role-badge" class="role-badge r-emp">👤 Employee</div>\n' +
'<div style="display:flex;align-items:center;gap:8px">\n' +
'<div class="av" id="u-av">U</div>\n' +
'<div><div class="uname" id="u-nm">—</div><div class="udept" id="u-dp">—</div></div>\n' +
'</div></div></div>\n' +

'<div class="mn">\n' +
'<div class="topbar">\n' +
'<h1 id="pg-ttl">Dashboard</h1>\n' +
'<div class="topbar-actions">\n' +
'<button class="smbtn primary" onclick="exportPDF()" id="pdf-btn" style="display:none">📄 PDF</button>\n' +
'<button class="smbtn primary" onclick="exportCSV()" id="exp-btn" style="display:none">📥 CSV</button>\n' +
'<button class="smbtn green" onclick="loadTickets()" id="ref-btn" style="display:none">↻ Refresh</button>\n' +
'<button class="smbtn outline" id="theme-btn" onclick="toggleTheme()" style="padding:4px 8px; font-size:14px; margin-right:4px;" title="Toggle Theme">🌙</button>\n' +
'<button class="smbtn outline" onclick="doLogout()">Logout</button>\n' +
'</div></div>\n' +
'<div class="ct">\n' +

/* ── Staff Dashboard ── */
'<div class="vp" id="vp-it">\n' +
'<div class="stats-grid" style="grid-template-columns:repeat(4,1fr)">\n' +
'<div class="stat-card"><div class="stat-label">Open</div><div class="stat-val c-blue" id="i-op">0</div></div>\n' +
'<div class="stat-card"><div class="stat-label">Seen/In Progress</div><div class="stat-val c-amber" id="i-pr">0</div></div>\n' +
'<div class="stat-card"><div class="stat-label">Resolved/Closed</div><div class="stat-val c-green" id="i-rs">0</div></div>\n' +
'<div class="stat-card"><div class="stat-label">Total</div><div class="stat-val c-text" id="i-tt">0</div></div>\n' +
'</div>\n' +
'<div class="chart-row" style="display:flex;gap:16px;margin:16px 0;height:240px;display:none" id="it-charts">\n' +
'<div class="fc" style="flex:1;display:flex;justify-content:center;position:relative"><canvas id="ch-dept"></canvas></div>\n' +
'<div class="fc" style="flex:1;display:flex;justify-content:center;position:relative"><canvas id="ch-status"></canvas></div>\n' +
'</div>\n' +
'<div class="fbar">\n' +
'<input class="fsrch" type="text" id="srch" placeholder="🔍 Ticket ID, naam, title..." oninput="renderIT()">\n' +
'<select class="fsel" id="f-status" onchange="renderIT()"><option value="">All Status</option><option>Open</option><option>Seen</option><option>In Progress</option><option>Resolved</option><option>Closed</option></select>\n' +
'<select class="fsel" id="f-pri" onchange="renderIT()"><option value="">All Priority</option><option>Low</option><option>Medium</option><option>High</option><option>Critical</option></select>\n' +
'</div>\n' +
'<div class="sect-title" id="it-count">All Tickets</div>\n' +
'<table class="tbl"><thead><tr><th>ID</th><th>Title</th><th>Employee</th><th>Dept</th><th>Priority</th><th>Status</th><th>Date</th></tr></thead>\n' +
'<tbody id="it-tb"></tbody></table>\n' +
'</div>\n' +

/* ── Employee Dashboard ── */
'<div class="vp" id="vp-emp">\n' +
'<div class="wc"><h2 id="wc-nm">Namaste! 👋</h2><p>Koi problem? Naya ticket daalo — IT jaldi solve karega.</p>\n' +
'<button class="wc-btn" onclick="sv(\'new-ticket\')">+ Naya Ticket Daalo</button></div>\n' +
'<div class="stats-grid" style="grid-template-columns:repeat(4,1fr)">\n' +
'<div class="stat-card"><div class="stat-label">Total</div><div class="stat-val c-text" id="e-tt">0</div></div>\n' +
'<div class="stat-card"><div class="stat-label">Open</div><div class="stat-val c-blue" id="e-op">0</div></div>\n' +
'<div class="stat-card"><div class="stat-label">Seen/In Progress</div><div class="stat-val c-amber" id="e-pr">0</div></div>\n' +
'<div class="stat-card"><div class="stat-label">Resolved/Closed</div><div class="stat-val c-green" id="e-rs">0</div></div>\n' +
'</div>\n' +
'<div class="sect-title">Meri Tickets</div>\n' +
'<table class="tbl"><thead><tr><th>ID</th><th>Title</th><th>Priority</th><th>Status</th><th>Date</th></tr></thead>\n' +
'<tbody id="em-tb"></tbody></table>\n' +
'</div>\n' +

/* Super Admin */
'<div class="vp" id="vp-master">\n' +
'<div class="admin-hero"><div><h2>Super Admin Control Center</h2><p>Employee master aur department routing rules yahin se manage honge.</p></div><div class="admin-pill">Master Access</div></div>\n' +
'<div style="display:flex;gap:10px;margin-bottom:16px;border-bottom:1px solid var(--b1);padding-bottom:12px">\n' +
'<button class="btn" style="flex:unset;min-width:120px" id="atab-form" onclick="swATab(\'form\')">Control Center</button>\n' +
'<button class="btn sec" style="flex:unset;min-width:120px" id="atab-list" onclick="swATab(\'list\')">Employee List</button>\n' +
'<button class="btn sec" style="flex:unset;min-width:120px" id="atab-perf" onclick="swATab(\'perf\')">Performance</button>\n' +
'</div>\n' +
'<div id="apanel-form">\n' +
'<div class="g2" style="align-items:start">\n' +
'<div class="fc">\n' +
'<div class="sect-title">Add New Employee</div>\n' +
'<div class="g2">\n' +
'<div class="fg"><label>Employee Code *</label><input type="text" id="a-code" placeholder="e.g. SSIFPL-100"></div>\n' +
'<div class="fg"><label>Name *</label><input type="text" id="a-name" placeholder="Employee name"></div>\n' +
'<div class="fg"><label>Designation *</label><input type="text" id="a-desig" placeholder="e.g. HR"></div>\n' +
'<div class="fg"><label>Password *</label><input type="text" id="a-pass" placeholder="Login password"></div>\n' +
'</div>\n' +
'<div class="fg"><label>Email Id *</label><input type="text" id="a-email" placeholder="employee@company.com"></div>\n' +
'<button class="btn" id="a-add-btn" onclick="adminAddEmployee()">Add Employee</button>\n' +
'</div>\n' +
'<div class="fc">\n' +
'<div class="sect-title">Department Mapping</div>\n' +
'<div class="fg"><label>Existing Designation</label><select id="m-pick" onchange="fillMapping()"><option value="">Select mapping</option></select></div>\n' +
'<div class="fg"><label>Designation *</label><input type="text" id="m-desig" placeholder="e.g. ACCOUNTS"></div>\n' +
'<div class="fg"><label>Allowed Departments *</label><textarea id="m-allowed" placeholder="IT, HR, Admin, DME"></textarea></div>\n' +
'<button class="btn" id="m-save-btn" onclick="adminSaveMapping()">Save Mapping</button>\n' +
'</div>\n' +
'</div>\n' +
'</div>\n' +
'<div id="apanel-list" style="display:none">\n' +
'<div class="fc" style="max-width:none">\n' +
'<div class="sect-title" id="admin-count" style="margin-bottom:0">Employees</div>\n' +
'<input type="text" id="emp-srch" placeholder="🔍 Search Employee..." style="padding:6px 12px; border-radius:6px; border:1px solid var(--b1); background:var(--s2); color:var(--tx); outline:none; font-family:\'Plus Jakarta Sans\'" oninput="filterAdminEmp()">\n' +
'</div>\n' +
'<table class="tbl"><thead><tr><th>Code</th><th>Name</th><th>Designation</th><th>Email Id</th></tr></thead><tbody id="admin-emp-tb"></tbody></table>\n' +
'</div>\n' +
'</div>\n' +
'<div id="apanel-perf" style="display:none">\n' +
'<div class="fc" style="max-width:none">\n' +
'<div class="sect-title">Department SLA & Performance</div>\n' +
'<div style="display:flex;gap:16px;height:280px;margin-bottom:16px">\n' +
'<div style="flex:1;position:relative"><canvas id="perf-ch-1"></canvas></div>\n' +
'<div style="flex:1;position:relative"><canvas id="perf-ch-2"></canvas></div>\n' +
'</div>\n' +
'</div>\n' +
'</div>\n' +
'</div>\n' +

/* ── New Ticket ── */
'<div class="vp" id="vp-new-ticket">\n' +
'<div class="fc">\n' +
'<div class="fg"><label>Title *</label><input type="text" id="t-ti" placeholder="Problem ka short summary..."></div>\n' +
'<div class="g2">\n' +
'<div class="fg"><label>Department *</label>\n' +
'<select id="t-dp" onchange="loadCats()">\n' +
'<option value="">Select Department</option>\n' +
'</select></div>\n' +
'<div class="fg"><label>Category *</label>\n' +
'<select id="t-ct"><option value="">Pehle dept select karo</option></select>\n' +
'</div></div>\n' +
'<div class="fg"><label>Priority *</label>\n' +
'<div class="po">\n' +
'<div class="po-opt sel-low" onclick="selPri(\'Low\')">Low</div>\n' +
'<div class="po-opt" onclick="selPri(\'Medium\')">Medium</div>\n' +
'<div class="po-opt" onclick="selPri(\'High\')">High</div>\n' +
'<div class="po-opt" onclick="selPri(\'Critical\')">Critical</div>\n' +
'</div></div>\n' +
'<div class="fg"><label>Description *</label><textarea id="t-ds" placeholder="Detail mein likho — kya ho raha hai, kab se?"></textarea></div>\n' +
'<div class="g2">\n' +
'<div class="fg"><label>Attachment (Optional)</label><input type="file" id="t-file" accept="image/*,.pdf,.doc,.docx" style="padding:7px"></div>\n' +
'<div class="fg"><label>Due Date</label><input type="date" id="t-du"></div>\n' +
'</div>\n' +
'<div class="form-actions">\n' +
'<button class="btn sec" onclick="goBack()">Cancel</button>\n' +
'<button class="btn" id="sub-btn" onclick="submitTicket()">Submit Ticket ✓</button>\n' +
'</div></div></div>\n' +

/* ── Ticket Detail ── */
'<div class="vp" id="vp-detail">\n' +
'<div class="backlink" onclick="goBack()">\n' +
'<svg fill="none" stroke="currentColor" width="14" height="14" viewBox="0 0 24 24"><path d="M19 12H5M5 12l7 7M5 12l7-7" stroke-width="1.5" stroke-linecap="round"/></svg>Back\n' +
'</div>\n' +
'<div class="detail-card">\n' +
'<div class="detail-head">\n' +
'<span class="tid" id="d-id"></span>\n' +
'<h2 id="d-ti"></h2>\n' +
'<div style="display:flex;gap:7px;flex-wrap:wrap;margin-bottom:7px">\n' +
'<span class="badge" id="d-dp"></span>\n' +
'<span class="badge" id="d-pr"></span>\n' +
'<span class="badge" id="d-st"></span>\n' +
'</div>\n' +
'<div class="email-info">\n' +
'<svg fill="none" stroke="currentColor" width="13" height="13" viewBox="0 0 24 24"><path d="M3 8l9 6 9-6M5 19h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z" stroke-width="1.5" stroke-linecap="round"/></svg>\n' +
'<span id="d-em">—</span>\n' +
'</div>\n' +
'</div>\n' +
'<div class="detail-body">\n' +
'<div class="detail-row">\n' +
'<div class="detail-item"><label>Submitted By</label><span id="d-by"></span></div>\n' +
'<div class="detail-item"><label>Category</label><span id="d-ct"></span></div>\n' +
'<div class="detail-item"><label>Date</label><span id="d-dt"></span></div>\n' +
'<div class="detail-item"><label>Due</label><span id="d-du"></span></div>\n' +
'</div>\n' +
'<div class="desc-box" id="d-ds"></div>\n' +
'<div class="detail-item" style="margin-top:14px"><label>Latest Remarks</label><div class="desc-box" id="d-rm" style="margin-top:8px;min-height:54px">No remarks added</div></div>\n' +
'<div class="update-box">\n' +
'<div style="font-size:10px;color:var(--mu);text-transform:uppercase;letter-spacing:.08em;font-weight:600;margin-bottom:10px">Update Status</div>\n' +
'<div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap">\n' +
'<select class="dsel" id="d-ss"><option>Open</option><option>Seen</option><option>In Progress</option><option>Resolved</option><option>Closed</option></select>\n' +
'<button class="smbtn primary" id="upd-btn" onclick="doUpdate()">✓ Save</button>\n' +
'</div>\n' +
'<div class="hint">💡 Resolved ya Closed karne pe employee ko auto email jayega</div>\n' +
'</div>\n' +
'</div></div></div>\n' +

'</div></div></div>\n' +

/* ===============================================================
   JAVASCRIPT
=============================================================== */
'<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>\n' +
'<script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>\n' +
'<script>\n' +
'"use strict";\n' +
'var cu = null, allowedRoutes = [], deptRole = null, tix = [], selPri_val = "Low", prevView = "", actId = null, authToken = "", masterPass = "", adminEmployees = [], adminMappings = [];\n' +
'var DC = {HR:"#f0407a",IT:"#3d6fff",DME:"#f0a020",Admin:"#1db874",Boss:"#9b5cf6",Other:"#6b7299", "ACCOUNT DEP":"#9b5cf6", "CRM DEP":"#f04040", "SUPPLY TEAM":"#3d6fff"};\n' +
'var CATS = {\n' +
'  IT:    ["Hardware","Software","Network","Access/Account","Email","Other"],\n' +
'  HR:    ["Payroll","Leave","Onboarding","Policy","Other"],\n' +
'  DME:   ["Equipment","Maintenance","Supply","Other"],\n' +
'  Admin: ["Facilities","Travel","Stationery","Other"],\n' +
'  "ACCOUNT DEP": ["Payment","Salary Issue","Invoice","Other"],\n' +
'  "CRM DEP": ["Customer Complaint","Refund","Sale Update","Other"],\n' +
'  "SUPPLY TEAM": ["Stock Status","Delivery Delay","Raw Material","Other"]\n' +
'};\n' +

'function initSky() {\n' +
'  var ns = document.getElementById("night-sky");\n' +
'  if (!ns) return;\n' +
'  var h = "";\n' +
'  for (var i = 0; i < 120; i++) {\n' +
'    h += "<div class=\'star\' style=\'left:" + (Math.random()*100) + "%;top:" + (Math.random()*100) + "%;width:" + (Math.random()*2+0.4) + "px;height:" + (Math.random()*2+0.4) + "px;animation-delay:" + (Math.random()*6) + "s\'></div>";\n' +
'  }\n' +
'  for (var j = 0; j < 8; j++) {\n' +
'    h += "<div class=\'shoot\' style=\'left:" + (Math.random()*150) + "%;top:" + (Math.random()*80-20) + "%;animation-delay:" + (Math.random()*10) + "s;animation-duration:" + (Math.random()*2+4) + "s\'></div>";\n' +
'  }\n' +
'  ns.innerHTML = h;\n' +
'}\n' +
'initSky();\n' +
'var empSub = document.querySelector("#form-emp .sub");\n' +
'if (empSub) empSub.textContent = "Apna Employee Code aur password daalo. Email automatic detect hoga.";\n' +
'var empBtn = document.getElementById("e-btn");\n' +
'if (empBtn) empBtn.textContent = "Login →";\n' +
'var empForm = document.getElementById("form-emp");\n' +
'if (empForm && !document.getElementById("emp-login-note")) {\n' +
'  var loginNote = document.createElement("div");\n' +
'  loginNote.id = "emp-login-note";\n' +
'  loginNote.style.cssText = "margin-top:14px;padding:12px 14px;border:1px solid rgba(59,130,246,.25);border-radius:12px;background:linear-gradient(135deg,rgba(59,130,246,.12),rgba(15,23,42,.35));color:#dbeafe;font-size:12px;line-height:1.6";\n' +
'  loginNote.innerHTML = "<div style=\'font-weight:700;margin-bottom:4px;color:#fff\'>Employee Access</div><div>Employee Code aur password se login karein.</div><div>Ticket updates automatically detected email par milenge.</div>";\n' +
'  empForm.appendChild(loginNote);\n' +
'}\n' +
'if (!document.getElementById("emp-polish-style")) {\n' +
'  var empStyle = document.createElement("style");\n' +
'  empStyle.id = "emp-polish-style";\n' +
'  empStyle.textContent = "#form-emp{padding:8px;border-radius:18px;background:linear-gradient(180deg,rgba(255,255,255,.03),rgba(59,130,246,.04))}#form-emp h2{font-size:28px;letter-spacing:-.02em;margin-bottom:8px}#form-emp .sub{color:#dbeafe;line-height:1.7;font-size:13px;margin-bottom:18px}#form-emp .fg label{font-size:11px;letter-spacing:.06em;text-transform:uppercase;color:#bfdbfe}#form-emp .fg input{border-radius:12px;background:rgba(15,23,42,.82);border:1px solid rgba(148,163,184,.22);padding:12px 14px;min-height:46px}#form-emp .fg input:focus{box-shadow:0 0 0 4px rgba(59,130,246,.18);border-color:#60a5fa}#form-emp .btn{min-height:48px;border-radius:14px;background:linear-gradient(135deg,#2563eb,#1d4ed8);box-shadow:0 14px 30px rgba(37,99,235,.28);transition:transform .18s ease,box-shadow .18s ease,opacity .18s ease}#form-emp .btn:hover{transform:translateY(-1px);box-shadow:0 18px 34px rgba(37,99,235,.34)}.wc-btn,.smbtn,.btn{transition:transform .18s ease,box-shadow .18s ease,opacity .18s ease}.wc-btn:hover,.smbtn:hover,.btn:hover{transform:translateY(-1px)}#vp-emp .wc{background:linear-gradient(135deg,rgba(37,99,235,.22),rgba(15,23,42,.8));border:1px solid rgba(96,165,250,.18);box-shadow:0 20px 50px rgba(2,6,23,.28)}#vp-emp .sect-title{display:flex;align-items:center;justify-content:space-between}";\n' +
'  document.head.appendChild(empStyle);\n' +
'}\n' +
'var statusHint = document.querySelector("#vp-detail .hint");\n' +
'if (statusHint) statusHint.textContent = "Seen status employee dashboard par dikhega. Resolved ya Closed karne par remarks ke saath auto email jayega.";\n' +
'var detailDesc = document.getElementById("d-ds");\n' +
'var updateBox = document.querySelector("#vp-detail .update-box");\n' +
'if (detailDesc && !document.getElementById("d-rm")) {\n' +
'  var remarksView = document.createElement("div");\n' +
'  remarksView.className = "detail-item";\n' +
'  remarksView.style.marginTop = "14px";\n' +
'  remarksView.innerHTML = "<label>Latest Remarks</label><div class=\'desc-box\' id=\'d-rm\' style=\'margin-top:8px;min-height:54px\'>No remarks added</div>";\n' +
'  detailDesc.insertAdjacentElement("afterend", remarksView);\n' +
'}\n' +
'if (updateBox && !document.getElementById("d-remarks")) {\n' +
'  var remarksWrap = document.createElement("div");\n' +
'  remarksWrap.className = "fg";\n' +
'  remarksWrap.style.marginTop = "12px";\n' +
'  remarksWrap.innerHTML = "<label>Remarks (Optional)</label><textarea id=\'d-remarks\' placeholder=\'Yahan optional remarks likh sakte hain...\'></textarea>";\n' +
'  updateBox.insertBefore(remarksWrap, updateBox.querySelector(".hint"));\n' +
'}\n' +

'function switchTab(t) {\n' +
'  var tabIds = ["emp","IT","HR","Admin","DME","Master"];\n' +
'  tabIds.forEach(function(id) {\n' +
'    var f = document.getElementById("form-" + id);\n' +
'    if (f) f.style.display = (id === t) ? "block" : "none";\n' +
'  });\n' +
'  document.querySelectorAll(".tab").forEach(function(btn, idx) {\n' +
'    btn.className = "tab" + (tabIds[idx] === t ? " on" : "");\n' +
'  });\n' +
'}\n' +
'function swATab(t) {\n' +
'  document.getElementById("apanel-form").style.display = t === "form" ? "block" : "none";\n' +
'  document.getElementById("apanel-list").style.display = t === "list" ? "block" : "none";\n' +
'  document.getElementById("apanel-perf").style.display = t === "perf" ? "block" : "none";\n' +
'  document.getElementById("atab-form").className = "btn" + (t === "form" ? "" : " sec");\n' +
'  document.getElementById("atab-list").className = "btn" + (t === "list" ? "" : " sec");\n' +
'  document.getElementById("atab-perf").className = "btn" + (t === "perf" ? "" : " sec");\n' +
'  if (t === "perf") renderPerfCharts();\n' +
'}\n' +

'function validEmail(e) {\n' +
'  e = e.trim();\n' +
'  if (!e) return false;\n' +
'  var parts = e.split("@");\n' +
'  if (parts.length !== 2) return false;\n' +
'  if (!parts[0] || !parts[1]) return false;\n' +
'  if (parts[1].indexOf(".") < 1) return false;\n' +
'  return true;\n' +
'}\n' +

'function empLogin() {\n' +
'  var ec = (document.getElementById("e-code") || {}).value || "";\n' +
'  var ep = (document.getElementById("e-pass") || {}).value || "";\n' +
'  var ok = true;\n' +
'  \n' +
'  ["e-code","e-pass"].forEach(function(i) {\n' +
'     if(document.getElementById(i)) document.getElementById(i).classList.remove("err");\n' +
'  });\n' +
'  document.querySelectorAll(".errmsg").forEach(function(el) { el.style.display = "none"; });\n' +
'  \n' +
'  if (!ec.trim()) { document.getElementById("e-code").classList.add("err"); document.getElementById("err-code").style.display="block"; ok=false; }\n' +
'  if (!ep.trim()) { document.getElementById("e-pass").classList.add("err"); document.getElementById("err-pass").style.display="block"; ok=false; }\n' +
'  if (!ok) return;\n' +
'  \n' +
'  deptRole = null;\n' +
'  showOv("Account Check ho raha hai...");\n' +
'  google.script.run\n' +
'    .withSuccessHandler(function(r) {\n' +
'      hideOv();\n' +
'      if (r.ok) {\n' +
'         authToken = r.token || "";\n' +
'         allowedRoutes = r.allowed;\n' +
'         startApp(r.name, r.dept, r.email);\n' +
'      } else {\n' +
'         showToast(r.msg, true);\n' +
'      }\n' +
'    })\n' +
'    .withFailureHandler(function(e) {\n' +
'      hideOv();\n' +
'      showToast("Error: " + publicClientError(e), true);\n' +
'    })\n' +
'    .serverEmpLogin(ec.trim(), ep);\n' +
'}\n' +

'function deptLogin(dept) {\n' +
'  var pass = document.getElementById("pass-" + dept).value;\n' +
'  if (!pass) { showToast("Password bharo!", true); return; }\n' +
'  showOv("Verify ho raha hai...");\n' +
'  google.script.run\n' +
'    .withSuccessHandler(function(r) {\n' +
'      hideOv();\n' +
'      if (r.ok) { authToken = r.token || ""; allowedRoutes = [dept]; deptRole = dept; startApp(r.name, r.dept, r.email); }\n' +
'      else { showToast("Galat password!", true); document.getElementById("pass-" + dept).value = ""; }\n' +
'    })\n' +
'    .withFailureHandler(function(e) { hideOv(); showToast("Error: " + publicClientError(e), true); })\n' +
'    .serverDeptLogin(dept, pass);\n' +
'}\n' +

'function masterLogin() {\n' +
'  var pass = document.getElementById("pass-Master").value;\n' +
'  if (!pass) { showToast("Password bharo!", true); return; }\n' +
'  showOv("Super Admin verify ho raha hai...");\n' +
'  google.script.run\n' +
'    .withSuccessHandler(function(r) {\n' +
'      hideOv();\n' +
'      if (r.ok) { authToken = r.token || ""; masterPass = r.token || ""; deptRole = "MASTER"; startApp(r.name, r.dept, r.email); }\n' +
'      else { showToast("Galat Super Admin password!", true); document.getElementById("pass-Master").value = ""; }\n' +
'    })\n' +
'    .withFailureHandler(function(e) { hideOv(); showToast("Error: " + publicClientError(e), true); })\n' +
'    .serverMasterLogin(pass);\n' +
'}\n' +

'function startApp(nm, dp, em) {\n' +
'  cu = { name: nm, dept: dp, email: em || "" };\n' +
'  var ini = nm.split(" ").map(function(w) { return w[0] || ""; }).join("").toUpperCase().slice(0,2);\n' +
'  var av = document.getElementById("u-av");\n' +
'  av.textContent = ini;\n' +
'  av.style.background = "linear-gradient(135deg," + (DC[dp] || "#3d6fff") + "," + (DC[dp] || "#3d6fff") + "99)";\n' +
'  document.getElementById("u-nm").textContent = nm;\n' +
'  document.getElementById("u-dp").textContent = dp;\n' +
'  var rb = document.getElementById("role-badge");\n' +
'  var icons = { IT:"🔧", HR:"👥", Admin:"🏢", DME:"⚙️", "ACCOUNT DEP":"💰", "CRM DEP":"📞", "SUPPLY TEAM":"📦" };\n' +
'  var roleClasses = { IT:"r-it", HR:"r-hr", Admin:"r-admin", DME:"r-dme", "ACCOUNT DEP":"r-account", "CRM DEP":"r-crm", "SUPPLY TEAM":"r-supply" };\n' +
'  if (deptRole) {\n' +
'    rb.textContent = deptRole === "MASTER" ? "\ud83d\udee1 Super Admin" : (icons[deptRole] || "\ud83d\udd27") + " " + deptRole + " Staff";\n' +
'    rb.className = "role-badge " + (roleClasses[deptRole] || "r-it");\n' +
'  } else {\n' +
'    rb.textContent = "\ud83d\udc64 " + (dp || "Employee");\n' +
'    rb.className = "role-badge r-emp";\n' +
'  }\n' +
'  document.getElementById("ref-btn").style.display = deptRole !== "MASTER" ? "inline-flex" : "none";\n' +
'  document.getElementById("exp-btn").style.display = deptRole ? "inline-flex" : "none";\n' +
'  document.getElementById("pdf-btn").style.display = deptRole ? "inline-flex" : "none";\n' +
'  buildNav();\n' +
'  document.getElementById("pg-login").classList.remove("on");\n' +
'  document.getElementById("pg-app").classList.add("on");\n' +
'  if (deptRole === "MASTER") { loadAdminData(); return; }\n' +
'  loadTickets();\n' +
'}\n' +

'function doLogout() {\n' +
'  cu = null; tix = []; deptRole = null; allowedRoutes = []; authToken = ""; masterPass = ""; adminEmployees = []; adminMappings = [];\n' +
'  document.getElementById("pg-app").classList.remove("on");\n' +
'  document.getElementById("pg-login").classList.add("on");\n' +
'  if (document.getElementById("e-code"))  document.getElementById("e-code").value = "";\n' +
'  if (document.getElementById("e-pass"))  document.getElementById("e-pass").value = "";\n' +
'  ["IT","HR","Admin","DME"].forEach(function(d) {\n' +
'     var p = document.getElementById("pass-" + d);\n' +
'     if (p) p.value = "";\n' +
'  });\n' +
'  if (document.getElementById("pass-Master")) document.getElementById("pass-Master").value = "";\n' +
'  switchTab("emp");\n' +
'}\n' +

'function buildNav() {\n' +
'  var nav = document.getElementById("sb-nav");\n' +
'  nav.innerHTML = "";\n' +
'  var items = deptRole === "MASTER"\n' +
'    ? [{ label: "Super Admin", view: "master" }]\n' +
'    : (deptRole ? [{ label: "All Tickets", view: "it" }, { label: "New Ticket", view: "new-ticket" }] : [{ label: "Dashboard", view: "emp" }, { label: "New Ticket", view: "new-ticket" }]);\n' +
'  items.forEach(function(item) {\n' +
'    var div = document.createElement("div");\n' +
'    div.className = "ni";\n' +
'    div.textContent = item.label;\n' +
'    div.setAttribute("data-view", item.view);\n' +
'    div.addEventListener("click", function() { sv(item.view); });\n' +
'    nav.appendChild(div);\n' +
'  });\n' +
'}\n' +

'function loadTickets() {\n' +
'  if (deptRole === "MASTER") { loadAdminData(); return; }\n' +
'  var query = deptRole\n' +
'    ? { token: authToken, role: "staff", dept: deptRole }\n' +
'    : { token: authToken, role: "employee" };\n' +
'  showOv("Tickets load ho rahe hain...");\n' +
'  google.script.run\n' +
'    .withSuccessHandler(function(r) {\n' +
'      hideOv();\n' +
'      if (r.ok) { tix = r.tickets || []; sv(deptRole ? "it" : "emp"); }\n' +
'      else { showToast("Error: " + r.msg, true); }\n' +
'    })\n' +
'    .withFailureHandler(function(e) { hideOv(); showToast("Load fail: " + publicClientError(e), true); })\n' +
'    .serverGetTickets(query);\n' +
'}\n' +

'function sv(v) {\n' +
'  if (v !== "detail" && v !== "new-ticket") prevView = v;\n' +
'  document.querySelectorAll(".vp").forEach(function(el) { el.classList.remove("on"); });\n' +
'  document.querySelectorAll(".ni").forEach(function(el) { el.classList.remove("on"); });\n' +
'  var el = document.getElementById("vp-" + v);\n' +
'  if (el) el.classList.add("on");\n' +
'  var titles = { it: "Tickets Dashboard", emp: "Mera Dashboard", master: "Super Admin", "new-ticket": "Naya Ticket", detail: "Ticket Detail" };\n' +
'  document.getElementById("pg-ttl").textContent = titles[v] || "";\n' +
'  document.querySelectorAll(".ni").forEach(function(ni) {\n' +
'    if (ni.getAttribute("data-view") === v) ni.classList.add("on");\n' +
'    if (v === "new-ticket" && ni.getAttribute("data-view") === "new-ticket") ni.classList.add("on");\n' +
'  });\n' +
'  if (v === "it") renderIT();\n' +
'  if (v === "emp") { renderEmp(); var c = document.getElementById("it-charts"); if(c) c.style.display = "none"; }\n' +
'  if (v === "master") renderAdmin();\n' +
'  if (v === "new-ticket") resetForm();\n' +
'}\n' +
'function goBack() { sv(prevView || (deptRole === "MASTER" ? "master" : (deptRole ? "it" : "emp"))); }\n' +

'function loadAdminData() {\n' +
'  showOv("Super Admin data load ho raha hai...");\n' +
'  google.script.run\n' +
'    .withSuccessHandler(function(r) {\n' +
'      hideOv();\n' +
'      if (r.ok) { adminEmployees = r.employees || []; adminMappings = r.mappings || []; sv("master"); }\n' +
'      else { showToast("Error: " + r.msg, true); }\n' +
'    })\n' +
'    .withFailureHandler(function(e) { hideOv(); showToast("Admin load fail: " + publicClientError(e), true); })\n' +
'    .serverAdminGetData(masterPass);\n' +
'}\n' +

'function renderAdmin() {\n' +
'  var pick = document.getElementById("m-pick");\n' +
'  if (pick) {\n' +
'    var chosen = pick.value;\n' +
'    pick.innerHTML = "<option value=\'\'>Select mapping</option>";\n' +
'    adminMappings.forEach(function(m, i) {\n' +
'      var opt = document.createElement("option"); opt.value = String(i); opt.textContent = m.designation; pick.appendChild(opt);\n' +
'    });\n' +
'    pick.value = chosen;\n' +
'  }\n' +
'  var tbody = document.getElementById("admin-emp-tb");\n' +
'  if (!tbody) return;\n' +
'  tbody.innerHTML = "";\n' +
'  document.getElementById("admin-count").textContent = adminEmployees.length + " employees";\n' +
'  if (!adminEmployees.length) {\n' +
'    var tr = document.createElement("tr"); var td = document.createElement("td");\n' +
'    td.colSpan = 4; td.className = "empty"; td.textContent = "Abhi employee list khali hai"; tr.appendChild(td); tbody.appendChild(tr); return;\n' +
'  }\n' +
'  adminEmployees.forEach(function(emp) {\n' +
'    var tr = document.createElement("tr");\n' +
'    tr.innerHTML = "<td><span class=\'tid\'>" + esc(emp.code) + "</span></td><td>" + esc(emp.name) + "</td><td>" + esc(emp.designation) + "</td><td style=\'font-size:11px;color:var(--mu)\'>" + esc(emp.email) + "</td>";\n' +
'    tbody.appendChild(tr);\n' +
'  });\n' +
'}\n' +

'function fillMapping() {\n' +
'  var idx = document.getElementById("m-pick").value;\n' +
'  if (idx === "") return;\n' +
'  var m = adminMappings[Number(idx)];\n' +
'  if (!m) return;\n' +
'  document.getElementById("m-desig").value = m.designation || "";\n' +
'  document.getElementById("m-allowed").value = m.allowed || "";\n' +
'}\n' +

'function adminAddEmployee() {\n' +
'  var emp = {\n' +
'    code: document.getElementById("a-code").value.trim(),\n' +
'    name: document.getElementById("a-name").value.trim(),\n' +
'    designation: document.getElementById("a-desig").value.trim(),\n' +
'    password: document.getElementById("a-pass").value.trim(),\n' +
'    email: document.getElementById("a-email").value.trim()\n' +
'  };\n' +
'  var btn = document.getElementById("a-add-btn"); btn.disabled = true;\n' +
'  showOv("Employee add ho raha hai...");\n' +
'  google.script.run\n' +
'    .withSuccessHandler(function(r) {\n' +
'      hideOv(); btn.disabled = false;\n' +
'      if (r.ok) {\n' +
'        adminEmployees.push(r.employee);\n' +
'        ["a-code","a-name","a-desig","a-pass","a-email"].forEach(function(id){ document.getElementById(id).value = ""; });\n' +
'        renderAdmin(); showToast("Employee add ho gaya");\n' +
'      } else { showToast("Error: " + r.msg, true); }\n' +
'    })\n' +
'    .withFailureHandler(function(e) { hideOv(); btn.disabled = false; showToast("Error: " + publicClientError(e), true); })\n' +
'    .serverAdminAddEmployee(masterPass, emp);\n' +
'}\n' +

'function adminSaveMapping() {\n' +
'  var desig = document.getElementById("m-desig").value.trim();\n' +
'  var allowed = document.getElementById("m-allowed").value.trim();\n' +
'  var btn = document.getElementById("m-save-btn"); btn.disabled = true;\n' +
'  showOv("Mapping save ho rahi hai...");\n' +
'  google.script.run\n' +
'    .withSuccessHandler(function(r) {\n' +
'      hideOv(); btn.disabled = false;\n' +
'      if (r.ok) {\n' +
'        var found = false;\n' +
'        adminMappings = adminMappings.map(function(m){ if (m.designation.toLowerCase() === r.mapping.designation.toLowerCase()) { found = true; return r.mapping; } return m; });\n' +
'        if (!found) adminMappings.push(r.mapping);\n' +
'        renderAdmin(); showToast("Mapping save ho gayi");\n' +
'      } else { showToast("Error: " + r.msg, true); }\n' +
'    })\n' +
'    .withFailureHandler(function(e) { hideOv(); btn.disabled = false; showToast("Error: " + publicClientError(e), true); })\n' +
'    .serverAdminSaveMapping(masterPass, desig, allowed);\n' +
'}\n' +

'function renderIT() {\n' +
'  var base = deptRole && deptRole !== "IT" ? tix.filter(function(t) { return t.dept === deptRole; }) : tix;\n' +
'  var srch = (document.getElementById("srch") || {}).value || "";\n' +
'  var fst  = (document.getElementById("f-status") || {}).value || "";\n' +
'  var fpri = (document.getElementById("f-pri") || {}).value || "";\n' +
'  var filtered = base;\n' +
'  if (srch) {\n' +
'    var s = srch.toLowerCase();\n' +
'    filtered = filtered.filter(function(t) {\n' +
'      return t.id.toLowerCase().indexOf(s) > -1 || t.title.toLowerCase().indexOf(s) > -1 || t.by.toLowerCase().indexOf(s) > -1;\n' +
'    });\n' +
'  }\n' +
'  if (fst)  filtered = filtered.filter(function(t) { return t.status === fst; });\n' +
'  if (fpri) filtered = filtered.filter(function(t) { return t.priority === fpri; });\n' +
'  document.getElementById("i-op").textContent = base.filter(function(t) { return t.status === "Open"; }).length;\n' +
'  document.getElementById("i-pr").textContent = base.filter(function(t) { return t.status === "Seen" || t.status === "In Progress"; }).length;\n' +
'  document.getElementById("i-rs").textContent = base.filter(function(t) { return t.status === "Resolved" || t.status === "Closed"; }).length;\n' +
'  document.getElementById("i-tt").textContent = base.length;\n' +
'  document.getElementById("it-count").textContent = filtered.length + " ticket" + (filtered.length !== 1 ? "s" : "") + " dikh rahe hain";\n' +
'  var tbody = document.getElementById("it-tb");\n' +
'  tbody.innerHTML = "";\n' +
'  if (!filtered.length) {\n' +
'    var tr = document.createElement("tr");\n' +
'    var td = document.createElement("td");\n' +
'    td.colSpan = 7; td.className = "empty"; td.textContent = "Koi ticket nahi mila";\n' +
'    tr.appendChild(td); tbody.appendChild(tr); return;\n' +
'  }\n' +
'  filtered.slice().sort(function(a,b){ return b.id.localeCompare(a.id); }).forEach(function(t) {\n' +
'    var tr = document.createElement("tr");\n' +
'    var isSla = false;\n' +
'    if (t.due && t.status !== "Resolved" && t.status !== "Closed") {\n' +
'       var d = new Date(t.due); if(!isNaN(d.getTime()) && d.getTime() < new Date().getTime()) isSla = true;\n' +
'    }\n' +
'    tr.className = "clickable" + (isSla ? " sla-breached" : "");\n' +
'    tr.innerHTML =\n' +
'      "<td><span class=\'tid\'>" + esc(t.id) + "</span></td>" +\n' +
'      "<td>" + esc(t.title) + "</td>" +\n' +
'      "<td style=\'font-size:11px;color:var(--mu)\'>" + esc(t.by) + "</td>" +\n' +
'      "<td><span class=\'badge d-" + esc(t.dept) + "\'>" + esc(t.dept) + "</span></td>" +\n' +
'      "<td><span class=\'badge p-" + t.priority.toLowerCase() + "\'>" + esc(t.priority) + "</span></td>" +\n' +
'      "<td><span class=\'badge s-" + t.status.replace(/ /g,"").toLowerCase() + "\'>" + esc(t.status) + "</span></td>" +\n' +
'      "<td style=\'font-size:11px;color:var(--mu)\'>" + esc(t.date) + "</td>";\n' +
'    tr.addEventListener("click", function() { openTicket(t.id); });\n' +
'    tbody.appendChild(tr);\n' +
'  });\n' +
'  if (typeof renderCharts === "function" && deptRole) renderCharts(base);\n' +
'}\n' +

'function renderEmp() {\n' +
'  var eEm = (cu.email || "").toLowerCase();\n' +
'  var eNm = (cu.name || "").toLowerCase();\n' +
'  var mine = tix.filter(function(t) {\n' +
'    var tEm = (t.email || "").toLowerCase();\n' +
'    var tNm = (t.by || "").toLowerCase();\n' +
'    return (eEm && tEm === eEm) || (!tEm && tNm === eNm);\n' +
'  });\n' +
'  document.getElementById("wc-nm").textContent = "Namaste, " + cu.name + "! 👋";\n' +
'  document.getElementById("e-tt").textContent = mine.length;\n' +
'  document.getElementById("e-op").textContent = mine.filter(function(t) { return t.status === "Open"; }).length;\n' +
'  document.getElementById("e-pr").textContent = mine.filter(function(t) { return t.status === "Seen" || t.status === "In Progress"; }).length;\n' +
'  document.getElementById("e-rs").textContent = mine.filter(function(t) { return t.status === "Resolved" || t.status === "Closed"; }).length;\n' +
'  var tbody = document.getElementById("em-tb");\n' +
'  tbody.innerHTML = "";\n' +
'  if (!mine.length) {\n' +
'    var tr = document.createElement("tr");\n' +
'    var td = document.createElement("td");\n' +
'    td.colSpan = 5; td.className = "empty"; td.textContent = "Abhi koi ticket nahi — Naya ticket daalo! 🎫";\n' +
'    tr.appendChild(td); tbody.appendChild(tr); return;\n' +
'  }\n' +
'  mine.slice().sort(function(a,b){ return b.id.localeCompare(a.id); }).forEach(function(t) {\n' +
'    var tr = document.createElement("tr");\n' +
'    var isSla = false;\n' +
'    if (t.due && t.status !== "Resolved" && t.status !== "Closed") {\n' +
'       var d = new Date(t.due); if(!isNaN(d.getTime()) && d.getTime() < new Date().getTime()) isSla = true;\n' +
'    }\n' +
'    tr.className = "clickable" + (isSla ? " sla-breached" : "");\n' +
'    tr.innerHTML =\n' +
'      "<td><span class=\'tid\'>" + esc(t.id) + "</span></td>" +\n' +
'      "<td>" + esc(t.title) + "</td>" +\n' +
'      "<td><span class=\'badge p-" + t.priority.toLowerCase() + "\'>" + esc(t.priority) + "</span></td>" +\n' +
'      "<td><span class=\'badge s-" + t.status.replace(/ /g,"").toLowerCase() + "\'>" + esc(t.status) + "</span></td>" +\n' +
'      "<td style=\'font-size:11px;color:var(--mu)\'>" + esc(t.date) + "</td>";\n' +
'    tr.addEventListener("click", function() { openTicket(t.id); });\n' +
'    tbody.appendChild(tr);\n' +
'  });\n' +
'}\n' +

'function openTicket(id) {\n' +
'  var t = tix.find(function(x) { return x.id === id; });\n' +
'  if (!t) return;\n' +
'  actId = id;\n' +
'  document.getElementById("d-id").textContent = t.id;\n' +
'  document.getElementById("d-ti").textContent = t.title;\n' +
'  var dp = document.getElementById("d-dp");\n' +
'  dp.textContent = t.dept; dp.className = "badge d-" + t.dept;\n' +
'  var pr = document.getElementById("d-pr");\n' +
'  pr.textContent = t.priority; pr.className = "badge p-" + t.priority.toLowerCase();\n' +
'  var st = document.getElementById("d-st");\n' +
'  st.textContent = t.status; st.className = "badge s-" + t.status.replace(/ /g,"").toLowerCase();\n' +
'  document.getElementById("d-em").textContent = t.email || "Email save nahi tha";\n' +
'  document.getElementById("d-by").textContent = t.by + " (" + t.byDept + ")";\n' +
'  document.getElementById("d-ct").textContent = t.category;\n' +
'  document.getElementById("d-dt").textContent = t.date;\n' +
'  document.getElementById("d-du").textContent = t.due || "—";\n' +
'  var dsDiv = document.getElementById("d-ds");\n' +
'  var attIdx = t.desc.indexOf("\\n\\nAttachment: ");\n' +
'  if (attIdx !== -1) {\n' +
'    dsDiv.innerHTML = esc(t.desc.substring(0, attIdx)).replace(/\\n/g, "<br>") + "<br><br><a href=\\"" + esc(t.desc.substring(attIdx + 14).trim()) + "\\" target=\\"_blank\\" style=\\"display:inline-block;padding:6px 12px;background:var(--ac);color:#fff;border-radius:4px;text-decoration:none;font-size:12px;font-weight:600\\">📎 View Attachment</a>";\n' +
'  } else {\n' +
'    dsDiv.innerHTML = esc(t.desc).replace(/\\n/g, "<br>");\n' +
'  }\n' +
'  if (document.getElementById("d-rm")) document.getElementById("d-rm").textContent = t.remarks || "No remarks added";\n' +
'  if (document.getElementById("d-remarks")) document.getElementById("d-remarks").value = t.remarks || "";\n' +
'  document.getElementById("d-ss").value = t.status;\n' +
'  var updBox = document.querySelector("#vp-detail .update-box");\n' +
'  if (updBox) updBox.style.display = "block";\n' +
'  var dSs = document.getElementById("d-ss");\n' +
'  if (dSs) dSs.style.display = deptRole ? "inline-block" : "none";\n' +
'  var updHint = document.querySelector("#vp-detail .hint");\n' +
'  if (updHint) updHint.style.display = deptRole ? "block" : "none";\n' +
'  var updTitle = document.querySelector("#vp-detail .update-box > div:first-child");\n' +
'  if (updTitle) updTitle.textContent = deptRole ? "Update Status / Add Reply" : "Add Reply";\n' +
'  sv("detail");\n' +
'}\n' +

'function doUpdate() {\n' +
'  var ns = document.getElementById("d-ss").value;\n' +
'  var remarks = ((document.getElementById("d-remarks") || {}).value || "").trim();\n' +
'  if (!deptRole && !remarks) { showToast("Pehle reply likhein!", true); return; }\n' +
'  var btn = document.getElementById("upd-btn");\n' +
'  btn.disabled = true;\n' +
'  showOv("Update ho raha hai...");\n' +
'  if (!deptRole) {\n' +
'    google.script.run\n' +
'      .withSuccessHandler(function(r) {\n' +
'        hideOv(); btn.disabled = false;\n' +
'        if (r.ok) {\n' +
'          var t = tix.find(function(x) { return x.id === actId; });\n' +
'          if (t) t.remarks = r.remarks;\n' +
'          if (document.getElementById("d-rm")) document.getElementById("d-rm").textContent = r.remarks;\n' +
'          document.getElementById("d-remarks").value = "";\n' +
'          showToast("✓ Reply added");\n' +
'        } else { showToast("Error: " + r.msg, true); }\n' +
'      })\n' +
'      .withFailureHandler(function(e) { hideOv(); btn.disabled = false; showToast("Error: " + publicClientError(e), true); })\n' +
'      .serverAddRemark(authToken, actId, remarks);\n' +
'    return;\n' +
'  }\n' +
'  google.script.run\n' +
'    .withSuccessHandler(function(r) {\n' +
'      hideOv(); btn.disabled = false;\n' +
'      if (r.ok) {\n' +
'        var t = tix.find(function(x) { return x.id === actId; });\n' +
'        if (t) { t.status = ns; t.remarks = remarks; }\n' +
 '        if (deptRole && (ns === "Resolved" || ns === "Closed")) {\n' +
 '          tix = tix.filter(function(x) { return x.id !== actId; });\n' +
 '        }\n' +
'        var st = document.getElementById("d-st");\n' +
'        st.textContent = ns;\n' +
'        st.className = "badge s-" + ns.replace(/ /g,"").toLowerCase();\n' +
'        if (document.getElementById("d-rm")) document.getElementById("d-rm").textContent = remarks || "No remarks added";\n' +
'        var msg = "✓ Status: " + ns;\n' +
'        if (ns === "Resolved" || ns === "Closed") msg += " | 📧 Email bheji gayi!";\n' +
'        showToast(msg);\n' +
'        if (deptRole && (ns === "Resolved" || ns === "Closed")) { sv("it"); return; }\n' +
'      } else { showToast("Error: " + r.msg, true); }\n' +
'    })\n' +
'    .withFailureHandler(function(e) { hideOv(); btn.disabled = false; showToast("Error: " + publicClientError(e), true); })\n' +
'    .serverUpdateStatus(authToken, actId, ns, remarks);\n' +
'}\n' +

'function loadCats() {\n' +
'  var dp = document.getElementById("t-dp").value;\n' +
'  var cats = CATS[dp] || [];\n' +
'  var sel = document.getElementById("t-ct");\n' +
'  sel.innerHTML = "<option value=\'\'>Select category</option>" + cats.map(function(c) { return "<option>" + c + "</option>"; }).join("");\n' +
'}\n' +

'function selPri(p) {\n' +
'  selPri_val = p;\n' +
'  document.querySelectorAll(".po-opt").forEach(function(o) {\n' +
'    o.className = "po-opt";\n' +
'  });\n' +
'  var labels = ["Low","Medium","High","Critical"];\n' +
'  document.querySelectorAll(".po-opt").forEach(function(o, i) {\n' +
'    if (labels[i] === p) o.className = "po-opt sel-" + p.toLowerCase();\n' +
'  });\n' +
'}\n' +

'function resetForm() {\n' +
'  ["t-ti","t-ds","t-du"].forEach(function(id) { document.getElementById(id).value = ""; });\n' +
'  var dpSel = document.getElementById("t-dp");\n' +
'  dpSel.innerHTML = "<option value=\'\'>Select Department</option>";\n' +
'  allowedRoutes.forEach(function(d) {\n' +
'     dpSel.innerHTML += "<option>" + d + "</option>";\n' +
'  });\n' +
'  document.getElementById("t-ct").innerHTML = "<option value=\'\'>Pehle dept select karo</option>";\n' +
'  selPri("Low");\n' +
'}\n' +

'function submitTicket() {\n' +
'  var ti = document.getElementById("t-ti").value.trim();\n' +
'  var dp = document.getElementById("t-dp").value;\n' +
'  var ct = document.getElementById("t-ct").value;\n' +
'  var ds = document.getElementById("t-ds").value.trim();\n' +
'  var du = document.getElementById("t-du").value;\n' +
'  if (!ti || !dp || !ct || !ds) { showToast("Saare * fields bharo!", true); return; }\n' +
'  document.getElementById("sub-btn").disabled = true;\n' +
'  showOv("Ticket submit ho raha hai...");\n' +
'  var fileInput = document.getElementById("t-file");\n' +
'  if (fileInput && fileInput.files.length > 0) {\n' +
'    var file = fileInput.files[0];\n' +
'    if (file.size > 3*1024*1024) { hideOv(); document.getElementById("sub-btn").disabled = false; showToast("File 3MB se badi nahi honi chahiye", true); return; }\n' +
'    var reader = new FileReader();\n' +
'    reader.onload = function(e) { callServer(e.target.result.split(",")[1], file.name); };\n' +
'    reader.readAsDataURL(file);\n' +
'  } else { callServer(null, null); }\n' +
'  function callServer(b64, fname) {\n' +
'    google.script.run\n' +
'      .withSuccessHandler(function(r) {\n' +
'        hideOv(); document.getElementById("sub-btn").disabled = false;\n' +
'        if (r.ok) {\n' +
'          tix.unshift({ id:r.id, title:ti, dept:dp, category:ct, priority:selPri_val,\n' +
'            status:"Open", by:cu.name, byDept:cu.dept, email:cu.email||"",\n' +
'            desc:ds + (r.att ? "\\n\\nAttachment: " + r.att : ""), date:r.date || formatNowDisplay(), due:du||"", remarks:"" });\n' +
'          showToast("✓ Ticket submit! ID: " + r.id);\n' +
'          sv(deptRole ? "it" : "emp");\n' +
'        } else { showToast("Error: " + r.msg, true); }\n' +
'      })\n' +
'      .withFailureHandler(function(e) { hideOv(); document.getElementById("sub-btn").disabled = false; showToast("Error: " + publicClientError(e), true); })\n' +
'      .serverAddTicket(authToken, ti, dp, ct, selPri_val, cu.name, cu.dept, cu.email||"", ds, du, b64, fname);\n' +
'  }\n' +
'}\n' +

'function esc(s) { var d = document.createElement("div"); d.textContent = String(s); return d.innerHTML; }\n' +
'function publicClientError(e) {\n' +
'  var msg = String((e && e.message) || e || "Server error");\n' +
'  return /sheet|spreadsheet|range|permission|exception/i.test(msg) ? "Server process complete nahi ho paya. Thodi der baad retry karein." : msg;\n' +
'}\n' +
'function formatNowDisplay() {\n' +
'  var d = new Date();\n' +
'  var pad = function(n) { return String(n).padStart(2, "0"); };\n' +
'  return d.getFullYear() + "-" + pad(d.getMonth()+1) + "-" + pad(d.getDate()) + " " + pad(d.getHours()) + ":" + pad(d.getMinutes());\n' +
'}\n' +
'function showToast(msg, err) {\n' +
'  var el = document.getElementById("toast");\n' +
'  el.textContent = msg;\n' +
'  el.style.borderLeft = "3px solid " + (err ? "var(--rd)" : "var(--gn)");\n' +
'  el.style.color = err ? "var(--rd)" : "var(--gn)";\n' +
'  el.style.display = "block";\n' +
'  clearTimeout(window._tt);\n' +
'  window._tt = setTimeout(function() { el.style.display = "none"; }, 5000);\n' +
'}\n' +
'function showOv(m) { document.getElementById("ov-txt").textContent = m || "Loading..."; document.getElementById("ov").classList.add("on"); }\n' +
'function hideOv() { document.getElementById("ov").classList.remove("on"); }\n' +
'function exportCSV() {\n' +
'  var data = [];\n' +
'  if (deptRole === "MASTER") {\n' +
'    data.push(["Code", "Name", "Designation", "Email"]);\n' +
'    adminEmployees.forEach(function(e) { data.push([e.code, e.name, e.designation, e.email]); });\n' +
'  } else {\n' +
'    var base = deptRole && deptRole !== "IT" ? tix.filter(function(t) { return t.dept === deptRole; }) : tix;\n' +
'    data.push(["ID", "Title", "Dept", "Priority", "Status", "By", "Date"]);\n' +
'    base.forEach(function(t) { data.push([t.id, t.title, t.dept, t.priority, t.status, t.by, t.date]); });\n' +
'  }\n' +
'  if (data.length < 2) { showToast("Koi data nahi hai export ke liye!", true); return; }\n' +
'  var csv = data.map(function(row) { return row.map(function(v) { return "\\"" + (String(v).replace(/"/g, \'""\')) + "\\""; }).join(","); }).join("\\n");\n' +
'  var blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });\n' +
'  var link = document.createElement("a");\n' +
'  link.href = URL.createObjectURL(blob);\n' +
'  link.download = "HelpDesk_Export_" + formatNowDisplay().replace(/[: ]/g, "_") + ".csv";\n' +
'  link.click();\n' +
'}\n' +
'function exportPDF() {\n' +
'  var element = document.getElementById(deptRole === "MASTER" ? "apanel-list" : "vp-it");\n' +
'  var opt = {\n' +
'    margin: 0.5,\n' +
'    filename: "HelpDesk_Report_" + formatNowDisplay().replace(/[: ]/g, "_") + ".pdf",\n' +
'    image: { type: "jpeg", quality: 0.98 },\n' +
'    html2canvas: { scale: 2, useCORS: true },\n' +
'    jsPDF: { unit: "in", format: "letter", orientation: "landscape" }\n' +
'  };\n' +
'  showOv("PDF generate ho raha hai...");\n' +
'  html2pdf().set(opt).from(element).save().then(function(){ hideOv(); });\n' +
'}\n' +
'function renderCharts(base) {\n' +
'  if (!document.getElementById("it-charts")) return;\n' +
'  document.getElementById("it-charts").style.display = "flex";\n' +
'  var deptCounts = {}, statusCounts = {};\n' +
'  base.forEach(function(t) {\n' +
'    deptCounts[t.dept] = (deptCounts[t.dept] || 0) + 1;\n' +
'    statusCounts[t.status] = (statusCounts[t.status] || 0) + 1;\n' +
'  });\n' +
'  var isDark = document.documentElement.getAttribute("data-theme") === "dark";\n' +
'  Chart.defaults.color = isDark ? "#f8fafc" : "#0f172a";\n' +
'  if (window.chDept) window.chDept.destroy();\n' +
'  if (window.chStatus) window.chStatus.destroy();\n' +
'  window.chDept = new Chart(document.getElementById("ch-dept"), { type: "doughnut", data: { labels: Object.keys(deptCounts), datasets: [{ data: Object.values(deptCounts), backgroundColor: ["#3b82f6","#10b981","#f59e0b","#ef4444","#8b5cf6"], borderColor: isDark ? "#1e293b" : "#ffffff" }] }, options: { maintainAspectRatio: false, plugins: { title: { display: true, text: "Tickets by Department" } }, animation: { duration: 0 } } });\n' +
'  window.chStatus = new Chart(document.getElementById("ch-status"), { type: "bar", data: { labels: Object.keys(statusCounts), datasets: [{ label: "Tickets", data: Object.values(statusCounts), backgroundColor: "#3b82f6" }] }, options: { maintainAspectRatio: false, plugins: { title: { display: true, text: "Tickets by Status" }, legend: { display: false } }, animation: { duration: 0 } } });\n' +
'}\n' +
'function toggleTheme() {\n' +
'  var html = document.documentElement;\n' +
'  var themeBtn = document.getElementById("theme-btn");\n' +
'  if (html.getAttribute("data-theme") === "dark") {\n' +
'    html.removeAttribute("data-theme");\n' +
'    localStorage.setItem("theme", "light");\n' +
'    if(themeBtn) themeBtn.textContent = "🌙";\n' +
'  } else {\n' +
'    html.setAttribute("data-theme", "dark");\n' +
'    localStorage.setItem("theme", "dark");\n' +
'    if(themeBtn) themeBtn.textContent = "☀️";\n' +
'  }\n' +
'  if (deptRole && deptRole !== "MASTER") renderIT();\n' +
'  if (deptRole === "MASTER" && document.getElementById("apanel-perf").style.display === "block") renderPerfCharts();\n' +
'}\n' +
'function filterAdminEmp() {\n' +
'  var s = (document.getElementById("emp-srch") || {}).value || "";\n' +
'  s = s.toLowerCase();\n' +
'  var tbody = document.getElementById("admin-emp-tb");\n' +
'  if (!tbody) return;\n' +
'  var rows = tbody.getElementsByTagName("tr");\n' +
'  for (var i = 0; i < rows.length; i++) {\n' +
'    var text = rows[i].textContent.toLowerCase();\n' +
'    rows[i].style.display = text.indexOf(s) > -1 ? "" : "none";\n' +
'  }\n' +
'}\n' +
'function renderPerfCharts() {\n' +
'  var perfDept = {};\n' +
'  tix.forEach(function(t) {\n' +
'    if(!perfDept[t.dept]) perfDept[t.dept] = { pending: 0, timely: 0, breached: 0 };\n' +
'    var isRes = t.status === "Resolved" || t.status === "Closed";\n' +
'    if(!isRes) perfDept[t.dept].pending++;\n' +
'    if(t.due) {\n' +
'       var dueDt = new Date(t.due).getTime();\n' +
'       if(isRes && dueDt > new Date().getTime()) perfDept[t.dept].timely++;\n' +
'       else if(!isRes && dueDt < new Date().getTime()) perfDept[t.dept].breached++;\n' +
'    }\n' +
'  });\n' +
'  var labels = Object.keys(perfDept);\n' +
'  var pendingData = labels.map(function(d) { return perfDept[d].pending; });\n' +
'  var breachedData = labels.map(function(d) { return perfDept[d].breached; });\n' +
'  var isDark = document.documentElement.getAttribute("data-theme") === "dark";\n' +
'  Chart.defaults.color = isDark ? "#f8fafc" : "#0f172a";\n' +
'  if(window.perfCh1) window.perfCh1.destroy();\n' +
'  if(window.perfCh2) window.perfCh2.destroy();\n' +
'  window.perfCh1 = new Chart(document.getElementById("perf-ch-1"), { type: "bar", data: { labels: labels, datasets: [{ label: "Pending Tickets", data: pendingData, backgroundColor: "#f59e0b" }, { label: "SLA Breached", data: breachedData, backgroundColor: "#ef4444" }] }, options: { maintainAspectRatio: false, plugins: { title: { display: true, text: "Pending & Breached SLA by Dept" } }, animation: { duration: 0 } } });\n' +
'  window.perfCh2 = new Chart(document.getElementById("perf-ch-2"), { type: "doughnut", data: { labels: labels, datasets: [{ data: pendingData, backgroundColor: ["#3b82f6","#10b981","#f59e0b","#ef4444","#8b5cf6"], borderColor: isDark ? "#1e293b" : "#ffffff" }] }, options: { maintainAspectRatio: false, plugins: { title: { display: true, text: "Overall Pending Workload" } }, animation: { duration: 0 } } });\n' +
'}\n' +
'window.addEventListener("DOMContentLoaded", function() {\n' +
'  if (localStorage.getItem("theme") === "dark") {\n' +
'    document.documentElement.setAttribute("data-theme", "dark");\n' +
'    var tb = document.getElementById("theme-btn");\n' +
'    if (tb) tb.textContent = "☀️";\n' +
'  }\n' +
'});\n' +
'</script>\n</body>\n</html>';
}

