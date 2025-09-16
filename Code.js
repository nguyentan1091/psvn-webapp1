// =================================================================
// CẤU HÌNH GOOGLE SHEETS
const SPREADSHEET_ID = ' '; 
const USERS_SHEET = 'Users';
const DATA_SHEET = 'VehicleData';
const TRUCK_LIST_TOTAL_SHEET = 'TruckListTotal';
const HISTORY_LOGIN_SHEET = 'History-login';
const CONTRACT_SHEET = 'ContractData';
const CONTRACT_HEADERS = ['ID', 'Contract No', 'Customer Name', 'Transportion Company', 'Status'];
// === XPPL Weighing Station database ===
const XPPL_DB_ID = '1LJGbMLFU8GnETecJ3i_j_fL5GWz5W1zST5bCQ5A5o3w';
const XPPL_DB_SHEET = 'XPPL-Database';
const XPPL_DB_HEADERS = [
  'ID','No.','W.ID','Weighing Type','TicketID','Truck No','Date In','Time In','Date Out','Time Out',
  'Weight In','Weight Out','Net Weight','Product Name','CoalSource','ProductionCode','Customer Name',
  'DriverName','Id/Passport','CargoLotNo','CargoName','CargoCompany','PackUnit','PackQtt','OrderNo',
  'ContractNo','InvoiceNo','CoNo','OVS_DMT','Plant','Trailer No','Truck Country','Truck Type',
  'WeighStationCode','Note','CreateUser','Transportion Company','Changed Date','Changed Time','Username'
];
// === XPPL TEMPLATE (Google Sheet chứa mẫu in) ===
// ID của file mẫu bạn gửi: https://docs.google.com/spreadsheets/d/18tVwSBr7tLU3uekL8Ay6gyrc4YFIFlS2/...
const XPPL_TEMPLATE_ID = '1p8n8ffm81NaxSWB5F7Wn1GhsaBrQ21XttaWmX5yvBl4';

/** ================== XPPL EXPORT – constants ================== **/
const XPPL_TEMP_PREFIX = 'XPPL_TMP_'; // prefix cho file tạm

// Các alias cho Named Range / Marker
const XPPL_NR_ALIASES = {
  SHEET:        ['Data','DATA','Sheet1'],
  REG_DATE:     ['NR_REG_DATE','REG_DATE'],
  CUSTOMER:     ['NR_CUSTOMER','CUSTOMER','CUSTOMER_CODE','CustomerCode'],
  CONTRACT_NO:  ['NR_CONTRACT_NO','CONTRACT_NO','Contract no'],
  TOTAL_TRUCK:  ['NR_TOTAL_TRUCK','TOTAL_TRUCK'],
  TABLE_START:  ['NR_TABLE_START','TABLE_START']
};


// Thứ tự cột cần đổ vào bảng (sau cột No)
const XPPL_TABLE_COLUMNS = [
  'Truck Plate',           // B
  'Country',               // C
  'Wheel',                 // D
  'Trailer Plate',         // E
  'Driver Name',           // F
  'ID/Passport',           // G
  'Phone number',          // H
  'Transportion Company',  // I
  'Subcontractor'          // J
];


const MAX_LOGIN_ATTEMPTS = 10;
const LOCKOUT_DURATION_1 = 10 * 60 * 1000; // 10 minutes
const LOCKOUT_DURATION_2 = 60 * 60 * 1000; // 1 hour
const SESSION_TIMEOUT_SECONDS = 30 * 60; // 30 minutes


// =============== DATE/TIME NORMALIZATION HELPERS ===============
function stripLeadingApostrophe(v) {
  if (typeof v === 'string' && v.length > 0 && v[0] === "'") return v.slice(1);
  return v;
}

function normalizeDate(v) {
  if (!v) return null;
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  v = stripLeadingApostrophe(v);
  var m = String(v).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;
  var d = parseInt(m[1],10), M = parseInt(m[2],10)-1, y = parseInt(m[3],10);
  var dt = new Date(y,M,d);
  if (isNaN(dt.getTime())) return null;
  return new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
}

function normalizeTime(v) {
  if (!v && v !== 0) return null;
  if (v instanceof Date) {
    var hh=v.getHours(), mm=v.getMinutes(), ss=v.getSeconds();
    return (hh*3600+mm*60+ss)/86400;
  }
  v = stripLeadingApostrophe(v);
  var m = String(v).match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!m) return null;
  var hh=parseInt(m[1],10), mm=parseInt(m[2],10), ss=m[3]?parseInt(m[3],10):0;
  if (hh>23||mm>59||ss>59) return null;
  return (hh*3600+mm*60+ss)/86400;
}

function formatDateForClient(v) {
  if (!v && v!==0) return '';
  if (v instanceof Date) return Utilities.formatDate(v, "Asia/Ho_Chi_Minh", "dd/MM/yyyy");
  return stripLeadingApostrophe(v);
}

function formatTimeForClient(v) {
  if (!v && v!==0) return '';
  if (v instanceof Date) return Utilities.formatDate(v, "Asia/Ho_Chi_Minh", "HH:mm:ss");
  if (typeof v === 'number') {
    var total = Math.round(v*86400);
    var hh = Math.floor(total/3600);
    var mm = Math.floor((total%3600)/60);
    var ss = total%60;
    return String(hh).padStart(2,'0')+':'+String(mm).padStart(2,'0')+':'+String(ss).padStart(2,'0');
  }
  return stripLeadingApostrophe(v);
}

function parseExcelDate_(v) {
  if (v == null || v === '') return '';
  var d;
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    d = v;
  } else if (typeof v === 'number') {
    d = new Date(Math.round((v - 25569) * 86400 * 1000));
  } else {
    var s = String(v).replace(/"/g, '');
    d = new Date(s);
    if (isNaN(d)) {
      var m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      if (m) d = new Date(parseInt(m[3],10), parseInt(m[2],10)-1, parseInt(m[1],10));
    }
  }
  if (!d || isNaN(d)) return '';
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function parseExcelTime_(v) {
  if (v == null || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return Utilities.formatDate(v, "Asia/Ho_Chi_Minh", "HH:mm:ss");
  }
  if (typeof v === 'number') {
    var total = Math.round((v % 1) * 86400);
    var hh = Math.floor(total/3600);
    var mm = Math.floor((total%3600)/60);
    var ss = total%60;
    return String(hh).padStart(2,'0')+':'+String(mm).padStart(2,'0')+':'+String(ss).padStart(2,'0');
  }
  var m = String(v).match(/(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?/);
  if (m) {
    var hh = parseInt(m[1],10);
    var mm = parseInt(m[2],10);
    var ss = m[3]?parseInt(m[3],10):0;
    return String(hh).padStart(2,'0')+':'+String(mm).padStart(2,'0')+':'+String(ss).padStart(2,'0');
  }
  return String(v);
}

function ensureDateTimeFormats(sheet, headers) {
  var dateCol = headers.indexOf('Register Date') + 1;
  var timeCol = headers.indexOf('Time') + 1;
  if (dateCol>0) sheet.getRange(2, dateCol, Math.max(1, sheet.getMaxRows()-1), 1).setNumberFormat("dd/MM/yyyy");
  if (timeCol>0) sheet.getRange(2, timeCol, Math.max(1, sheet.getMaxRows()-1), 1).setNumberFormat("HH:mm:ss");
}

function formatRowForClient_(rowArray, headers) {
  var out = {};
  for (var i=0;i<headers.length;i++){
    var key = headers[i];
    var val = rowArray[i];
    if (key === 'Register Date' || key === 'Date In' || key === 'Date Out' || key === 'Changed Date') {
      out[key] = formatDateForClient(val);
      continue;
    }
    if (key === 'Time' || key === 'Time In' || key === 'Time Out' || key === 'Changed Time') {
      out[key] = formatTimeForClient(val);
      continue;
    }
    out[key] = stripLeadingApostrophe(val);
  }
  return out;
}


const HEADERS_REGISTER = [
  'ID', 'Register Date', 'Contract No', 'Truck Plate', 'Country', 'Wheel', 
  'Trailer Plate', 'Truck weight', 'Pay load', 'Container No1', 'Container No2', 
  'Driver Name', 'ID/Passport', 'Phone number', 'Destination EST', 
  'Transportion Company', 'Subcontractor', 'Vehicle Status', 'Registration Status', 'Time'
];
const HEADERS_TOTAL_LIST = [
  'ID', 'Truck Plate', 'Country', 'Wheel', 'Trailer Plate', 'Truck weight',
  'Pay load', 'Container No1', 'Container No2', 'Driver Name', 'ID/Passport',
  'Phone number', 'Transportion Company', 'Subcontractor', 'Vehicle Status',
  'Activity Status', 'Register Date', 'Time'
];


// =================================================================
// KHỞI TẠO WEB APP
// =================================================================
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Hệ Thống Quản Lý Xe')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// =================================================================
// QUẢN LÝ PHIÊN LÀM VIỆC VÀ XÁC THỰC
// =================================================================

function validateSession(sessionToken) {
  const userCache = CacheService.getUserCache();
  const sessionData = userCache.get('user_session');
  if (sessionData) {
    const session = JSON.parse(sessionData);
    if (session.token === sessionToken) {
      userCache.put('user_session', JSON.stringify(session), SESSION_TIMEOUT_SECONDS);
      return session;
    }
  }
  throw new Error('Bạn chưa đăng nhập hoặc phiên đã hết hạn. Vui lòng đăng nhập lại.');
}

/** ADMIN GUARD (XPPL admin-only) */
function requireAdmin_(sessionToken) {
  const s = validateSession(sessionToken);
  if (!s || s.role !== 'admin') {
    throw new Error('Bạn không có quyền truy cập chức năng này (Admin only).');
  }
  return s;
}

function requireXpplRole_(sessionToken) {
  const s = validateSession(sessionToken);
  if (!s || ['admin','admin-xppl'].indexOf(s.role) === -1) {
    throw new Error('Bạn không có quyền truy cập chức năng này.');
  }
  return s;
}

function logLoginAttempt(username, status) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(HISTORY_LOGIN_SHEET);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp', 'Username', 'Status']);
    }
    sheet.appendRow([new Date(), username, status]);
  } catch (e) {
    Logger.log('Không thể ghi lịch sử đăng nhập: ' + e.message);
  }
}

function checkLogin(credentials) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const userSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USERS_SHEET);
  const username = credentials.username;
  
  try {
    const lockoutUntil = scriptProperties.getProperty(`lockout_until_${username}`);
    if (lockoutUntil && new Date().getTime() < parseFloat(lockoutUntil)) {
      const timeLeft = Math.ceil((parseFloat(lockoutUntil) - new Date().getTime()) / (60 * 1000));
      throw new Error(`Tài khoản của bạn đã bị tạm khóa. Vui lòng thử lại sau ${timeLeft} phút.`);
    }

    if (userSheet.getLastRow() < 2) throw new Error('Không có dữ liệu người dùng.');
    const usersRange = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 8);
    const users = usersRange.getValues();
    let userRowIndex = -1;
    let userRecord = null;

    for (let i = 0; i < users.length; i++) {
        if (users[i][0] === username) {
            userRowIndex = i;
            userRecord = users[i];
            break;
        }
    }

    if (!userRecord || userRecord[1] !== credentials.password) {
      logLoginAttempt(username, 'Failure');
      let failedAttempts = parseInt(scriptProperties.getProperty(`failed_attempts_${username}`) || '0') + 1;
      if (failedAttempts >= MAX_LOGIN_ATTEMPTS) {
        let lockoutLevel = parseInt(scriptProperties.getProperty(`lockout_level_${username}`) || '0') + 1;
        const lockoutDuration = (lockoutLevel === 1) ? LOCKOUT_DURATION_1 : LOCKOUT_DURATION_2;
        const lockoutEndTime = new Date().getTime() + lockoutDuration;
        scriptProperties.setProperty(`lockout_until_${username}`, lockoutEndTime);
        scriptProperties.setProperty(`lockout_level_${username}`, lockoutLevel);
        scriptProperties.deleteProperty(`failed_attempts_${username}`);
        const lockDurationMinutes = Math.ceil(lockoutDuration / (60 * 1000));
        throw new Error(`Bạn đã nhập sai quá nhiều lần. Tài khoản bị khóa trong ${lockDurationMinutes} phút.`);
      } else {
        scriptProperties.setProperty(`failed_attempts_${username}`, failedAttempts);
      }
      throw new Error('Tên đăng nhập hoặc mật khẩu không đúng.');
    }

    const activeToken = userRecord[6]; // Cột G (chỉ số 6)
    const tokenExpiry = userRecord[7]; // Cột H (chỉ số 7)
    
    // **SỬA LỖI:** Kích hoạt lại dòng kiểm tra
    if (activeToken && tokenExpiry && new Date().getTime() < new Date(tokenExpiry).getTime()) {
      throw new Error('Tài khoản này đã được đăng nhập trên một thiết bị khác.');
    }

    logLoginAttempt(username, 'Success');
    scriptProperties.deleteProperty(`failed_attempts_${username}`);
    scriptProperties.deleteProperty(`lockout_until_${username}`);
    scriptProperties.deleteProperty(`lockout_level_${username}`);

    const newSessionToken = Utilities.getUuid();
    const newExpiry = new Date(new Date().getTime() + SESSION_TIMEOUT_SECONDS * 1000);
    
    userSheet.getRange(userRowIndex + 2, 7).setValue(newSessionToken);
    userSheet.getRange(userRowIndex + 2, 8).setValue(newExpiry);

    const userSession = {
      isLoggedIn: true,
      username: userRecord[0],
      role: userRecord[2],
      contractor: userRecord[3],
      token: newSessionToken
    };

    const userCache = CacheService.getUserCache();
    userCache.put('user_session', JSON.stringify(userSession), SESSION_TIMEOUT_SECONDS);

    return userSession;
  } catch (e) {
    Logger.log(e);
    throw new Error(e.message);
  }
}


function logout() {
  const userCache = CacheService.getUserCache();
  const sessionData = userCache.get('user_session');
  if (sessionData) {
    const session = JSON.parse(sessionData);
    const userSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USERS_SHEET);
    const users = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 1).getValues().flat();
    const userRowIndex = users.indexOf(session.username);
    if (userRowIndex !== -1) {
      userSheet.getRange(userRowIndex + 2, 7, 1, 2).clearContent();
    }
  }
  userCache.remove('user_session');
  return { success: true };
}

function changePassword(passwords, sessionToken) {
  const session = validateSession(sessionToken);
  const { currentPassword, newPassword } = passwords;

  try {
    const userSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USERS_SHEET);
    const usersRange = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 2);
    const users = usersRange.getValues();
    const userRowIndex = users.findIndex(u => u[0] === session.username);

    if (userRowIndex === -1) throw new Error('Không tìm thấy người dùng.');
    if (users[userRowIndex][1] !== currentPassword) throw new Error('Mật khẩu hiện tại không đúng.');

    userSheet.getRange(userRowIndex + 2, 2).setValue(newPassword);
    userSheet.getRange(userRowIndex + 2, 5).setValue(new Date());

    return 'Đổi mật khẩu thành công!';
  } catch (e) { Logger.log(e); throw new Error('Lỗi khi đổi mật khẩu: ' + e.message); }
}

function resetPassword(data) {
  const { username, securityCode, newPassword } = data;
  try {
    const userSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USERS_SHEET);
    const usersRange = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 6);
    const users = usersRange.getValues();
    const userRowIndex = users.findIndex(u => u[0] === username);

    if (userRowIndex === -1) throw new Error('Tên đăng nhập không tồn tại.');
    if (users[userRowIndex][5] !== securityCode) throw new Error('Mã bảo mật không chính xác.');

    userSheet.getRange(userRowIndex + 2, 2).setValue(newPassword);
    userSheet.getRange(userRowIndex + 2, 5).setValue(new Date());

    return 'Đặt lại mật khẩu thành công! Vui lòng đăng nhập lại.';
  } catch (e) { Logger.log(e); throw new Error('Lỗi khi đặt lại mật khẩu: ' + e.message); }
}

// =================================================================
// QUẢN LÝ NGƯỜI DÙNG CHO ADMIN
// =================================================================

function getUsers(sessionToken) {
  const session = validateSession(sessionToken);
  if (session.role !== 'admin') throw new Error('Bạn không có quyền truy cập chức năng này.');

  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USERS_SHEET);
    if (sheet.getLastRow() < 2) return [];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
    
    return data.map(row => {
      let formattedDate = '';
      if (row[4] instanceof Date) {
        formattedDate = Utilities.formatDate(row[4], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
      } else if (row[4]) {
        formattedDate = String(row[4]);
      }

      return {
        Username: row[0],
        Password: row[1],
        Role: row[2],
        Contractor: row[3],
        PasswordLastUpdated: formattedDate,
        SecurityCode: row[5]
      }
    });
  } catch (e) { Logger.log(e); throw new Error('Không thể lấy danh sách người dùng.'); }
}

function updateUser(userData, sessionToken) {
  const session = validateSession(sessionToken);
  if (session.role !== 'admin') throw new Error('Bạn không có quyền truy cập chức năng này.');

  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USERS_SHEET);
    const users = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
    const userRowIndex = users.indexOf(userData.Username);

    if (userRowIndex === -1) throw new Error('Không tìm thấy người dùng.');

    sheet.getRange(userRowIndex + 2, 3).setValue(userData.Role);
    sheet.getRange(userRowIndex + 2, 4).setValue(userData.Contractor);
    
    return 'Cập nhật người dùng thành công!';
  } catch (e) { Logger.log(e); throw new Error('Lỗi khi cập nhật người dùng.'); }
}

function adminResetPassword(username, sessionToken) {
  const session = validateSession(sessionToken);
  if (session.role !== 'admin') throw new Error('Bạn không có quyền truy cập chức năng này.');

  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USERS_SHEET);
    const users = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
    const userRowIndex = users.indexOf(username);

    if (userRowIndex === -1) throw new Error('Không tìm thấy người dùng.');

    const newPassword = Math.random().toString(36).slice(-8);
    sheet.getRange(userRowIndex + 2, 2).setValue(newPassword);
    sheet.getRange(userRowIndex + 2, 5).setValue(new Date());

    return `Mật khẩu mới cho ${username} là: ${newPassword}`;
  } catch (e) { Logger.log(e); throw new Error('Lỗi khi đặt lại mật khẩu.'); }
}

function addNewUser(newUserData, sessionToken) {
  const session = validateSession(sessionToken);
  if (session.role !== 'admin') throw new Error('Bạn không có quyền truy cập chức năng này.');

  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USERS_SHEET);
    const users = sheet.getRange(2, 1, sheet.getLastRow(), 1).getValues().flat();
    const userExists = users.some(u => u === newUserData.Username);

    if (userExists) throw new Error('Tên đăng nhập đã tồn tại.');

    const newPassword = Math.random().toString(36).slice(-8);
    const newSecurityCode = Math.random().toString(36).slice(-6).toUpperCase();

    sheet.appendRow([
      newUserData.Username,
      newPassword,
      newUserData.Role,
      newUserData.Contractor,
      new Date(),
      newSecurityCode,
      '',
      ''
    ]);

    return `Đã tạo người dùng ${newUserData.Username} thành công.\nMật khẩu: ${newPassword}\nMã bảo mật: ${newSecurityCode}`;
  } catch (e) { Logger.log(e); throw new Error('Lỗi khi tạo người dùng mới: ' + e.message); }
}

function deleteUser(username, sessionToken) {
  const session = validateSession(sessionToken);
  if (session.role !== 'admin') throw new Error('Bạn không có quyền truy cập chức năng này.');
  if (session.username === username) throw new Error('Bạn không thể tự xóa tài khoản của mình.');

  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USERS_SHEET);
    const users = sheet.getRange(2, 1, sheet.getLastRow(), 1).getValues().flat();
    const userRowIndex = users.indexOf(username);

    if (userRowIndex === -1) throw new Error('Không tìm thấy người dùng.');

    sheet.deleteRow(userRowIndex + 2);
    return `Đã xóa người dùng ${username} thành công!`;
  } catch (e) { Logger.log(e); throw new Error('Lỗi khi xóa người dùng.'); }
}

// =================================================================
// KIỂM TRA THỜI GIAN ĐĂNG KÝ
// =================================================================
function checkRegistrationTime() {
  const now = new Date();
  const nowVn = new Date(now.toLocaleString('en-US', { timeZone: 'Asia/Ho_Chi_Minh' }));
  const hour = nowVn.getHours();
  const minute = nowVn.getMinutes();
  const currentTimeInMinutes = hour * 60 + minute;
  const open1 = 8 * 60, close1 = 16 * 60, open2 = 20 * 60, close2 = 22 * 60;
  let status = { isOpen: false, period: 0, message: '', timeToOpen: 0, timeToClose: 0 };
  
  if ((currentTimeInMinutes >= open1 && currentTimeInMinutes < close1)) {
    status.isOpen = true;
    status.period = 1;
    status.message = 'Hệ thống đăng ký bốc hàng sẽ đóng sau:';
    status.timeToClose = (close1 - currentTimeInMinutes) * 60 * 1000;
  } else if (currentTimeInMinutes >= open2 && currentTimeInMinutes < close2) {
    status.isOpen = true;
    status.period = 2;
    status.message = 'Hệ thống đăng ký bốc hàng sẽ đóng sau:';
    status.timeToClose = (close2 - currentTimeInMinutes) * 60 * 1000;
  } else {
    status.isOpen = false;
    status.message = 'Hệ thống đăng ký bốc hàng đang đóng.';
    if (currentTimeInMinutes < open1) status.timeToOpen = (open1 - currentTimeInMinutes) * 60 * 1000;
    else if (currentTimeInMinutes < open2) status.timeToOpen = (open2 - currentTimeInMinutes) * 60 * 1000;
    else status.timeToOpen = ((24 * 60 - currentTimeInMinutes) + open1) * 60 * 1000;
  }
  return status;
}

// =================================================================
// XỬ LÝ DỮ LIỆU PHÍA MÁY CHỦ (SERVER-SIDE)
// =================================================================

/**
 * Trả về map { ContractNo: [Customer1, Customer2, ...] } từ sheet ContractData.
 * Dùng cho dropdown "Customer Name".
 */
function getCustomersByContracts(contracts, sessionToken) {
  const sess = validateSession(sessionToken);
  if (!sess || sess.role !== 'admin') throw new Error('Chỉ admin.');

  if (!contracts || !contracts.length) return {};

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(typeof CONTRACT_SHEET === 'string' ? CONTRACT_SHEET : 'ContractData');
  if (!sh) return {};

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return {};

  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  // helper: tìm index cột theo nhiều tên khả dĩ
  function findIdx(hs, names) {
    const low = hs.map(h => String(h || '').trim().toLowerCase());
    for (const n of names) {
      const i = low.indexOf(String(n).trim().toLowerCase());
      if (i !== -1) return i;
    }
    return -1;
  }

  const idxNo  = findIdx(headers, ['Contract No','Contract no','Số HĐ','Số hợp đồng','So HD','So hop dong']);
  const idxCus = findIdx(headers, ['Customer Name','Customer','CustomerCode','Customer code','Khách hàng']);
  if (idxNo === -1 || idxCus === -1) return {};

  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const allow = new Set(contracts.map(s => String(s || '').replace(/^'/,'').trim()));
  const out = {};
  for (const r of values) {
    const no  = String(r[idxNo]  || '').replace(/^'/,'').trim();
    const cus = String(r[idxCus] || '').replace(/^'/,'').trim();
    if (!out[no]) out[no] = new Set();
    out[no].add(cus);
  }
  // Set -> Array đã sort
  const outArr = {};
  for (const k in out) outArr[k] = Array.from(out[k]).sort();
  return outArr;
}

function processServerSide(params, sheetName, headers, defaultSortColumnIndex) {
  const userSession = validateSession(params.sessionToken);

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { draw: params.draw, recordsTotal: 0, recordsFiltered: 0, data: [] };
  }

  let allData = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

  if (userSession.role === 'user') {
    const companyColumnIndex = headers.indexOf('Transportion Company');
    if (companyColumnIndex !== -1) {
        allData = allData.filter(row => row[companyColumnIndex] === userSession.contractor);
    }
    const activityColumnIndex = headers.indexOf('Activity Status');
    if (activityColumnIndex !== -1) {
        allData = allData.filter(row => String(row[activityColumnIndex]).toUpperCase() === 'ACTIVE');
    }    
  }

  if (params.dateString) {
    const dateColumnIndex = headers.indexOf('Register Date');
    allData = allData.filter(row => {
      let cellValue = row[dateColumnIndex];
      if (!cellValue) return false;
      let cmp = '';
      if (cellValue instanceof Date) cmp = Utilities.formatDate(cellValue, "Asia/Ho_Chi_Minh", "dd/MM/yyyy");
      else cmp = String(cellValue).trim().replace("'", "");
      return cmp === params.dateString;
    });
  }

  const recordsTotal = allData.length;
  let filteredData = allData;

  if (params.search && params.search.value) {
    const searchValue = params.search.value.toLowerCase();
    filteredData = filteredData.filter(row => {
      return row.some(cell => String(cell).toLowerCase().includes(searchValue));
    });
  }
  const recordsFiltered = filteredData.length;

  
  // === SUMMARY (for 'Xe đã đăng ký') ===
  var summary = null;
  try {
    if (sheetName === DATA_SHEET) {
      var statusIdx = headers.indexOf('Registration Status');
      if (statusIdx !== -1) {
        var total = filteredData.length;
        var pending = 0, approved = 0;
        for (var i = 0; i < filteredData.length; i++) {
          var v = filteredData[i][statusIdx];
          v = (v instanceof Date) ? Utilities.formatDate(v, "Asia/Ho_Chi_Minh", "dd/MM/yyyy")
                                  : String(v || '').replace(/^'/, '').trim();
          if (/^pending approval$/i.test(v)) pending++;
          else if (/^approved$/i.test(v)) approved++;
        }
        summary = { total: total, pending: pending, approved: approved };
      }
    }
  } catch (e) { /* ignore summary errors */ }
if (params.order && params.order.length > 0) {
    const orderInfo = params.order[0];
    const columnIndex = orderInfo.column;
    const direction = orderInfo.dir === 'asc' ? 1 : -1;
    filteredData.sort((a, b) => {
      const valA = a[columnIndex];
      const valB = b[columnIndex];
      if (valA < valB) return -1 * direction;
      if (valA > valB) return 1 * direction;
      return 0;
    });
  } else if (defaultSortColumnIndex !== -1) {
    filteredData.sort((a, b) => (a[defaultSortColumnIndex] < b[defaultSortColumnIndex] ? 1 : -1));
  }


  const paginatedData = filteredData.slice(params.start, params.start + params.length);
  const data = paginatedData.map(row => formatRowForClient_(row, headers));
return {
    draw: parseInt(params.draw),
    recordsTotal: recordsTotal,
    recordsFiltered: recordsFiltered,
    data: data,
    summary: summary
  };
}

function getRegisteredDataServerSide(params) {
  return processServerSide(params, DATA_SHEET, HEADERS_REGISTER, HEADERS_REGISTER.indexOf('Time'));
}

function getTotalListDataServerSide(params) {
  return processServerSide(params, TRUCK_LIST_TOTAL_SHEET, HEADERS_TOTAL_LIST, HEADERS_TOTAL_LIST.indexOf('Register Date'));
}

/** =========================
 *  XPPL — OPTIONS cho dropdown
 *  Input:  getXpplExportOptions({ dateString: 'dd/MM/yyyy' }, sessionToken)
 *  Return: { contracts: string[], customersByContract: { [contractNo]: string[] } }
 * ========================= */
function getXpplExportOptions(filter, sessionToken) {
  // Tùy hệ thống của bạn dùng validateSession/requireAdmin_:
  if (typeof validateSession === 'function') validateSession(sessionToken);

  const s = v => String(v == null ? '' : v).replace(/^'+/, '').trim();
  const normH = x => s(x).toLowerCase().replace(/\s+/g, '');
  const findIdx = (headers, variants) => {
    const H = headers.map(normH);
    for (const v of variants) {
      const i = H.indexOf(v);
      if (i !== -1) return i;
    }
    return -1;
  };

  const dateIn = s(filter && filter.dateString);
  const dateKey = _toDateKey(dateIn);
  if (!dateKey) return { contracts: [], customersByContract: {} };

  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const shV = ss.getSheetByName(DATA_SHEET);       // 'VehicleData'
  const shC = ss.getSheetByName(CONTRACT_SHEET);   // 'ContractData'
  if (!shV || shV.getLastRow() < 2) return { contracts: [], customersByContract: {} };

  // ---- VehicleData: lấy Contract theo ngày + Approved
  const lcV      = shV.getLastColumn();
  const headVRaw = shV.getRange(1, 1, 1, lcV).getValues()[0];
  const idxDate   = findIdx(headVRaw, ['registerdate','ngàydăngký','date','register']);
  const idxNo     = findIdx(headVRaw, ['contractno','contractnumber','sốhđ','sohd','sốhợpđồng']);
  const idxStatus = findIdx(headVRaw, ['registrationstatus','status','trạngtháiđăngký','trangthai']);
  if (idxDate === -1 || idxNo === -1) return { contracts: [], customersByContract: {} };

  const rowsV = shV.getRange(2, 1, shV.getLastRow() - 1, lcV).getValues();
  const setContracts = new Set();
  for (const r of rowsV) {
    if (_toDateKey(r[idxDate]) !== dateKey) continue;
    if (idxStatus !== -1) {
      const st = s(r[idxStatus]).toLowerCase();
      if (st !== 'approved') continue;
    }
    const cno = s(r[idxNo]);
    if (cno) setContracts.add(cno);
  }
  const contracts = Array.from(setContracts).sort();
  if (!contracts.length) return { contracts: [], customersByContract: {} };

  // ---- ContractData: map Contract -> Customers (ưu tiên Status=Active nếu có)
  const customersByContract = {};
  for (const c of contracts) customersByContract[c] = [];

  if (shC && shC.getLastRow() > 1) {
    const lcC      = shC.getLastColumn();
    const headCRaw = shC.getRange(1, 1, 1, lcC).getValues()[0];
    const idxCNo     = findIdx(headCRaw, ['contractno','contractnumber','sốhđ','sohd','sốhợpđồng']);
    const idxCus     = findIdx(headCRaw, ['customername','customer','kháchhàng','khachhang']);
    const idxCStatus = findIdx(headCRaw, ['status','trạngthái','trangthai']);

    if (idxCNo !== -1 && idxCus !== -1) {
      const rowsC = shC.getRange(2, 1, shC.getLastRow() - 1, lcC).getValues();
      for (const r of rowsC) {
        const cno = s(r[idxCNo]);
        if (!(cno in customersByContract)) continue;
        if (idxCStatus !== -1) {
          const st = s(r[idxCStatus]).toLowerCase();
          if (st && st !== 'active') continue;
        }
        const cus = s(r[idxCus]);
        if (cus && customersByContract[cno].indexOf(cus) === -1) {
          customersByContract[cno].push(cus);
        }
      }
      Object.keys(customersByContract).forEach(no => customersByContract[no].sort());
    }
  }

  return { contracts, customersByContract };
}

/** =========================
 *  XPPL — Lấy dữ liệu xuất theo filter
 *  Input:  getXpplExportData({ dateString, contractNo, customerName }, sessionToken)
 *  Return: { ok, errors?, total, headers, rows }
 * ========================= */
function getXpplExportData(filter, sessionToken) {
  if (typeof requireAdmin_ === 'function') requireAdmin_(sessionToken);

  // helpers
  const s = v => String(v == null ? '' : v).replace(/^'+/, '').trim();

  // validate input
  const dateIn       = s(filter && filter.dateString);
  const contractNo   = s(filter && filter.contractNo);
  const customerName = s(filter && filter.customerName);
  const inputErr = [];
  if (!dateIn)       inputErr.push('Thiếu Register Date.');
  if (!contractNo)   inputErr.push('Thiếu Contract No.');
  if (!customerName) inputErr.push('Thiếu Customer Name.');
  if (inputErr.length) return { ok:false, errors: inputErr };

  // open SS + normalize date
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tz = (ss.getSpreadsheetTimeZone && ss.getSpreadsheetTimeZone()) || 'Asia/Ho_Chi_Minh';
  const toDateKey = (v) => {
    if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, tz, 'dd/MM/yyyy');
    let str = String(v||'').trim().replace(/^'+/, '');
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(str)) return str;
    if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
      const d = new Date(str); if (!isNaN(d)) return Utilities.formatDate(d, tz, 'dd/MM/yyyy');
    }
    return str;
  };
  const dateKey = toDateKey(dateIn);

  // 1) Xác thực Contract–Customer (nếu có sheet ContractData)
  const shC = ss.getSheetByName(CONTRACT_SHEET);
  if (shC && shC.getLastRow() > 1) {
    const lc = shC.getLastColumn();
    const H  = shC.getRange(1,1,1,lc).getValues()[0];
    const normalizeHeader = x => String(x||'').trim().toLowerCase().replace(/\s+/g,'');
    const findIndex = (hdr, keys) => {
      const HH = hdr.map(normalizeHeader);
      for (const k of keys) { const i = HH.indexOf(k); if (i !== -1) return i; }
      return -1;
    };
    const iNo  = findIndex(H, ['contractno','contractnumber','sốhđ','sohd','sốhợpđồng']);
    const iCus = findIndex(H, ['customername','customer','kháchhàng','khachhang']);
    const iStt = findIndex(H, ['status','trạngthái','trangthai']);

    if (iNo !== -1 && iCus !== -1) {
      const ok = shC.getRange(2,1,shC.getLastRow()-1,lc).getValues().some(r => {
        if (s(r[iNo]) !== contractNo || s(r[iCus]) !== customerName) return false;
        if (iStt === -1) return true;
        const st = s(r[iStt]).toLowerCase();
        return st === '' || st === 'active';
      });
      if (!ok) {
        return { ok:false, errors:['Customer Name không khớp với Contract No (hoặc hợp đồng không Active).'] };
      }
    }
  }

  // 2) Lọc VehicleData
  const shV = ss.getSheetByName(DATA_SHEET);
  if (!shV || shV.getLastRow() < 2) return { ok:false, errors:['Không có dữ liệu VehicleData.'] };

  const lcV = shV.getLastColumn();
  const HV  = shV.getRange(1,1,1,lcV).getValues()[0];

  // chuẩn hóa header + tìm index (có fuzzy)
  const normalizeHeader = x => String(x||'').trim().toLowerCase().replace(/\s+/g,'');
  const findIdx = (keys) => {
    const H = HV.map(normalizeHeader);
    // exact
    for (const k of keys) {
      const i = H.indexOf(k);
      if (i !== -1) return i;
    }
    // fuzzy: keys[0] chứa các từ cần có
    if (keys.length) {
      const need = keys[0].split(' ').filter(Boolean);
      for (let i = 0; i < H.length; i++) {
        const h = H[i];
        if (need.every(w => h.includes(w))) return i;
      }
    }
    return -1;
  };

  const iDate   = findIdx(['registerdate','ngàydăngký','date','register']);
  const iNo2    = findIdx(['contractno','contractnumber','sốhđ','sohd','sốhợpđồng']);
  const iStReg  = findIdx(['registrationstatus','status','trạngtháiđăngký','trangthai']);

  const iTruck  = findIdx(['truckplate','truck plate','biểnxe','biensoxe']);
  const iCountry= findIdx(['country','quốcgia','quocgia']);
  const iWheel  = findIdx(['wheel','sốtrục','sotruc']);
  const iTrailer= findIdx(['trailerplate','trailer plate','biểnromooc','bienromooc']);
  const iDriver = findIdx(['drivername','driver name','tênlái','tenlai']);
  const iID     = findIdx(['id/passport','idpassport','passport','id']);
  const iPhone  = findIdx(['phone number','phonenumber','điệnthoại','dienthoai']);

  // CHÚ Ý: bắt mọi biến thể "transportion/transportation/transport company"
  const iTrans  = findIdx([
    'transportion company',
    'transport company',
    'transportation company',
    'transportationcompany',
    'transportioncompany',
    'transportcompany'
  ]);

  const iSub    = findIdx(['subcontractor','thầuphụ','thaophu']);

  if (iDate === -1 || iNo2 === -1) {
    return { ok:false, errors:['Thiếu cột bắt buộc trong VehicleData (Register Date / Contract No).'] };
  }

  const all = shV.getRange(2,1,shV.getLastRow()-1,lcV).getValues();
  const rows = [];
  for (const r of all) {
    if (toDateKey(r[iDate]) !== dateKey) continue;
    if (s(r[iNo2]) !== contractNo) continue;
    if (iStReg !== -1 && s(r[iStReg]).toLowerCase() !== 'approved') continue;

    rows.push({
      'Truck Plate':            iTruck  !== -1 ? s(r[iTruck])  : '',
      'Country':                iCountry!== -1 ? s(r[iCountry]): '',
      'Wheel':                  iWheel  !== -1 ? s(r[iWheel])  : '',
      'Trailer Plate':          iTrailer!== -1 ? s(r[iTrailer]): '',
      'Driver Name':            iDriver !== -1 ? s(r[iDriver]) : '',
      'ID/Passport':            iID     !== -1 ? s(r[iID])     : '',
      'Phone number':           iPhone  !== -1 ? s(r[iPhone])  : '',
      'Transportion Company':   iTrans  !== -1 ? s(r[iTrans])  : '',
      'Subcontractor':          iSub    !== -1 ? s(r[iSub])    : ''
    });
  }

  if (!rows.length) return { ok:false, errors:['Không có dòng Approved phù hợp để xuất.'] };

  return { ok:true, filter:{ dateString: dateKey, contractNo, customerName }, total: rows.length, rows };
}


// Tìm range theo danh sách NamedRange / nếu không có thì fallback tìm marker text
function _getRangeByAnyName_(ss, aliases){
  const names = Array.isArray(aliases) ? aliases : [aliases];
  const nr = ss.getNamedRanges();
  const low = {};
  nr.forEach(n => low[String(n.getName()).toLowerCase()] = n.getRange());

  for (const n of names) {
    const k = String(n).toLowerCase().trim();
    if (low[k]) return low[k];
  }
  // fallback: tìm ô chứa đúng chuỗi marker
  return _findMarkerCell_(ss, names);
}

/** Tìm ô có chữ 'NR_TABLE_START' trên sheet (fallback khi thiếu named-range). */
function _findMarkerCell_(ss, names){
  const shNames = XPPL_NR_ALIASES.SHEET;
  for (const sn of shNames){
    const sh = ss.getSheetByName(sn);
    if (!sh) continue;
    const lastR = Math.max(1, sh.getLastRow());
    const lastC = Math.max(1, sh.getLastColumn());
    const values = sh.getRange(1,1,lastR,lastC).getValues();

    for (let r=0;r<values.length;r++){
      for (let c=0;c<values[r].length;c++){
        const v = String(values[r][c]||'').trim();
        if (names.some(n => String(n).trim()===v)){
          return sh.getRange(r+1, c+1);
        }
      }
    }
  }
  return null;
}


// Copy template và ép CONVERT thành Google Sheets trước khi open
function _copyTemplateAsGoogleSheet_(templateFileId, newTitle) {
  var meta = Drive.Files.get(templateFileId); // cần Advanced Drive Service
  if (meta.mimeType === 'application/vnd.google-apps.spreadsheet') {
    // Template là Google Sheet -> copy trực tiếp
    return DriveApp.getFileById(templateFileId).makeCopy(newTitle).getId();
  } else {
    // Template là .xlsx -> convert sang Google Sheet
    var blob = DriveApp.getFileById(templateFileId).getBlob();
    var file = Drive.Files.insert(
      { title: newTitle, mimeType: 'application/vnd.google-apps.spreadsheet' },
      blob,
      { convert: true }
    );
    return file.id;
  }
}



/** Ghi dữ liệu vào bản sao template (Google Sheet). Trả về {ok, fileId, name}. */
function _exportXpplToTemplate_(sheetId, filter, rows) {
  const ss = SpreadsheetApp.openById(sheetId);

  // --- Header ---
  const rDate = _getRangeByAnyName_(ss, XPPL_NR_ALIASES.REG_DATE);
  if (rDate) rDate.setValue(filter.dateString);
  const rCus = _getRangeByAnyName_(ss, XPPL_NR_ALIASES.CUSTOMER);
  if (rCus) rCus.setValue(filter.customerName);
  const rCon = _getRangeByAnyName_(ss, XPPL_NR_ALIASES.CONTRACT_NO);
  if (rCon) rCon.setValue(filter.contractNo);
  const rTot = _getRangeByAnyName_(ss, XPPL_NR_ALIASES.TOTAL_TRUCK);
  if (rTot) rTot.setValue(rows.length);

  // --- Table ---
  const start = _getRangeByAnyName_(ss, XPPL_NR_ALIASES.TABLE_START);
  if (!start) throw new Error('Không tìm thấy named range TABLE_START / NR_TABLE_START');

  const sh = start.getSheet();
  const r0 = start.getRow();     // ví dụ: 9
  const c0 = start.getColumn();  // ví dụ: 1 (cột A)

  // map 10 cột A..J
  const data = rows.map(o => ([
    '', // A - No (đổ sau)
    o['Truck Plate'] || '',
    o['Country'] || '',
    o['Wheel'] || '',
    o['Trailer Plate'] || '',
    o['Driver Name'] || '',
    o['ID/Passport'] || '',
    o['Phone number'] || '',
    o['Transportion Company'] || '',
    o['Subcontractor'] || ''
  ]));

  if (data.length) {
    // Ghi ĐÈ trực tiếp tại TABLE_START để dòng 9 là bản ghi #1
    sh.getRange(r0, c0, data.length, data[0].length).setValues(data);

    // Cột A: No = 1..N
    const nos = Array.from({ length: data.length }, (_, i) => [i + 1]);
    sh.getRange(r0, c0, data.length, 1).setValues(nos);
  }

  // ---------- ĐỊNH DẠNG ----------
  // Row 1 cao ~27.6px => 28px
  sh.setRowHeight(1, 28);
  // Row 3 cao ~31.2px => 31px
  sh.setRowHeight(3, 31);

  if (data.length) {
    // Kẻ ALL BORDERS cho vùng dữ liệu A..J từ dòng r0
    const tableRange = sh.getRange(r0, c0, data.length, 10);
    tableRange
      .setBorder(true, true, true, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID)
      .setWrap(true); // chữ xuống dòng nếu dài

    // (Tuỳ chọn) Font Times New Roman cho bảng
    // tableRange.setFontFamily('Times New Roman');
  }

  SpreadsheetApp.flush();
}



/**
 * Xuất ra XLSX (base64) rồi xóa bản sao Google Sheet để không phình dung lượng.
 * YÊU CẦU: bật Advanced Drive Service (Drive API v2).
 */
function exportXpplAsXlsx(payload, sessionToken) {
  const res = getXpplExportData(payload, sessionToken);
  if (!res || !res.ok) {
    return { ok:false, message:(res && res.errors && res.errors.join('\n')) || 'Không đủ điều kiện để xuất.' };
  }
  const { dateString, contractNo, customerName } = res.filter;
  const rows = res.rows || [];
  if (!rows.length) return { ok:false, message:'Không có dữ liệu để xuất.' };

  // 1) Copy template -> Google Sheet
  const nameSuffix = dateString.replace(/\//g, '-');

  // QUAN TRỌNG: thêm prefix để sweeper tìm và xoá
  const copyName = `${XPPL_TEMP_PREFIX}(${contractNo}_${nameSuffix})-XPPL FORM`;
  const copiedId = _copyTemplateAsGoogleSheet_(XPPL_TEMPLATE_ID, copyName);

  // 2) Ghi dữ liệu vào bản copy
  _exportXpplToTemplate_(copiedId, { dateString, contractNo, customerName }, rows);

  // 3) Flush + đợi 1 nhịp rồi export đúng bản copy
  SpreadsheetApp.flush();
  Utilities.sleep(800);

  const url  = `https://docs.google.com/spreadsheets/d/${copiedId}/export?format=xlsx`;
  const resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) {
    return { ok:false, message:'Export lỗi: ' + resp.getContentText() };
  }

  // 4) Tên file tải về -> làm sạch ký tự cấm
  const safeName = (copyName + '.xlsx').replace(/[\\\/:\*\?"<>\|]/g, '_');

  // 5) (BỎ) trigger one-shot sau 3 phút — không cần nữa
  // try { ScriptApp.newTrigger('cleanupXpplTempFiles').timeBased().after(3*60*1000).create(); } catch(e){}

  // 6) ĐẢM BẢO đã có sweeper chạy định kỳ (nếu chưa có thì tạo 1 lần)
  try { ensureXpplSweeper(); } catch (e) { /* ignore */ }

  return {
    ok: true,
    fileName: safeName,
    base64: Utilities.base64Encode(resp.getBlob().getBytes())
  };
}


// ====== Sweeper dọn file tạm XPPL ======

// Tạo 1 time-based trigger chạy cleanupXpplTempFiles mỗi 5 phút (chỉ tạo 1 lần)
function ensureXpplSweeper() {
  var key = 'XPPL_SWEEPER_CREATED';
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty(key)) return;

  ScriptApp.newTrigger('cleanupXpplTempFiles')
    .timeBased()
    .everyMinutes(5)   // 5 phút/lần
    .create();

  props.setProperty(key, '1');
}

// Hàm dọn rác: xóa các file có tên bắt đầu bằng XPPL_TEMP_PREFIX và cũ > 3 phút
function cleanupXpplTempFiles() {
  var prefix = XPPL_TEMP_PREFIX || 'XPPL_TMP-';
  var cutoff = new Date(Date.now() - 3 * 60 * 1000); // 3 phút trước

  var it = DriveApp.searchFiles('title contains "' + prefix + '" and trashed = false');
  var removed = 0;
  while (it.hasNext()) {
    try {
      var f = it.next();
      if (f.getName().indexOf(prefix) === 0 && f.getDateCreated() < cutoff) {
        f.setTrashed(true);
        removed++;
      }
    } catch (e) {}
  }
  return removed;
}

// (Khuyến nghị) Bảo đảm sweeper tồn tại ngay khi mở project
function onOpen() {
  try { ensureXpplSweeper(); } catch (e) {}
}



// Tạo 1 trigger một-lần chạy vào/ sau thời điểm due sớm nhất
function _ensureCleanupTrigger_(dueTs) {
  var exists = ScriptApp.getProjectTriggers()
    .some(function(t){ return t.getHandlerFunction() === 'xpplCleanupDueFiles'; });
  if (!exists) {
    var now = Date.now();
    var when = Math.max(dueTs, now + 60 * 1000); // luôn >= 1 phút để an toàn
    ScriptApp.newTrigger('xpplCleanupDueFiles').timeBased().at(new Date(when)).create();
  }
}

// Xoá vĩnh viễn tất cả file copy đã đến hạn; nếu còn file chưa đến hạn -> hẹn trigger lần sau
function xpplCleanupDueFiles() {
  var props = PropertiesService.getScriptProperties();
  var all   = props.getProperties();
  var now   = Date.now();
  var prefix = 'xppl_delete_';
  var nextDue = null;

  for (var k in all) {
    if (k.indexOf(prefix) !== 0) continue;
    var fileId = k.substring(prefix.length);
    var due = parseInt(all[k], 10) || 0;

    if (now >= due) {
      // đã đến hạn -> xoá vĩnh viễn
      try { Drive.Files.remove(fileId); } 
      catch (e) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch (_) {} }
      // xoá key
      props.deleteProperty(k);
    } else {
      // chưa đến hạn -> giữ lại và ghi nhận mốc sớm nhất
      if (nextDue === null || due < nextDue) nextDue = due;
    }
  }

  // Nếu vẫn còn file cần xoá trong tương lai -> đặt lại trigger đến mốc sớm nhất
  if (nextDue !== null) _ensureCleanupTrigger_(nextDue);
}



/** Dự phòng _toDateKey nếu dự án chưa có */
function _toDateKey(v) {
  if (v == null || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    const dd = ('0' + v.getDate()).slice(-2);
    const mm = ('0' + (v.getMonth() + 1)).slice(-2);
    const yy = v.getFullYear();
    return `${dd}/${mm}/${yy}`;
  }
  let s = String(v).trim().replace(/^'+/, '');
  // dd/MM/yyyy
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) return s;
  // yyyy-MM-dd or yyyy-MM-ddTHH:mm:ss...
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    const d = new Date(s);
    if (!isNaN(d)) return _toDateKey(d);
  }
  // dd-MM-yyyy
  if (/^\d{2}-\d{2}-\d{4}$/.test(s)) {
    const [d, m, y] = s.split('-').map(Number);
    return _toDateKey(new Date(y, m - 1, d));
  }
  return '';
}



// =================================================================
// LOGIC XỬ LÝ ĐĂNG KÝ XE
// =================================================================

// =================================================================
// LOGIC XỬ LÝ ĐĂNG KÝ XE
// Gom tất cả biển số mới (chưa có trong danh sách tổng) thay vì trả về ngay chiếc đầu tiên
// =================================================================
// =================================================================
// LOGIC XỬ LÝ ĐĂNG KÝ XE
// Gom tất cả biển số mới (chưa có trong danh sách tổng) thay vì trả về ngay chiếc đầu tiên
// =================================================================
// =================================================================
// LOGIC XỬ LÝ ĐĂNG KÝ XE – Gom đủ "Xe mới" và "Xe trùng đơn vị khác"
// =================================================================
function checkVehiclesAgainstTotalList(vehicles) {
  const totalListSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(TRUCK_LIST_TOTAL_SHEET);
  if (totalListSheet.getLastRow() < 2) {
    return { isValid: false, message: 'Danh sách xe tổng chưa có dữ liệu. Không thể đăng ký. Vui lòng liên hệ PSVN.' };
  }

  const totalListData = totalListSheet.getRange(
    2, 1, totalListSheet.getLastRow() - 1, HEADERS_TOTAL_LIST.length
  ).getValues();

  const truckPlateIndex = HEADERS_TOTAL_LIST.indexOf('Truck Plate');
  const companyIndex    = HEADERS_TOTAL_LIST.indexOf('Transportion Company');

  // Map: PLATE -> Company (normalize biển số: uppercase & bỏ khoảng trắng)
  const totalListMap = new Map();
  totalListData.forEach(row => {
    const plate = row[truckPlateIndex];
    if (plate) {
      totalListMap.set(String(plate).toUpperCase().replace(/\s/g, ''), row[companyIndex]);
    }
  });

  const newPlates = [];            // các xe chưa có trong danh sách tổng
  const mismatchPlates = [];       // các xe đã đăng ký với đơn vị khác
  const seenNew = new Set();
  const seenMismatch = new Set();

  for (const vehicle of vehicles) {
    const plateRaw = vehicle['Truck Plate'] || '';
    const plate    = String(plateRaw).toUpperCase().replace(/\s/g, '');
    const company  = vehicle['Transportion Company'];

    if (!plate) continue;

    if (!totalListMap.has(plate)) {
      if (!seenNew.has(plate)) {
        seenNew.add(plate);
        newPlates.push(plate);
      }
      continue;
    }

    const registeredCompany = totalListMap.get(plate);
    if (registeredCompany !== company) {
      if (!seenMismatch.has(plate)) {
        seenMismatch.add(plate);
        mismatchPlates.push(plate);
      }
    }
  }

  // Ưu tiên báo "Xe mới" nếu có
  if (newPlates.length > 0) {
    return {
      isValid: false,
      message: `Xe ${newPlates.join(',')} là xe mới. Yêu cầu gửi đăng ký bổ sung vào danh sách tổng với PSVN.`
    };
  }

  // Nếu có xe trùng đơn vị khác → trả về danh sách biển số
  if (mismatchPlates.length > 0) {
    return {
      isValid: false,
      message: `Xe ${mismatchPlates.join(',')} đã được đăng ký với đơn vị vận chuyển khác. Yêu cầu liên hệ PSVN để được xử lý.`
    };
  }

  return { isValid: true };
}

// =================================================================
// LOGIC XỬ LÝ ĐĂNG KÝ XE – Kiểm tra Activity Status
// =================================================================
function checkVehicleActivityStatus(vehicles) {
  const totalListSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(TRUCK_LIST_TOTAL_SHEET);
  if (totalListSheet.getLastRow() < 2) {
    return { isValid: false, message: 'Danh sách xe tổng chưa có dữ liệu. Không thể đăng ký. Vui lòng liên hệ PSVN.' };
  }

  const totalListData = totalListSheet.getRange(
    2, 1, totalListSheet.getLastRow() - 1, HEADERS_TOTAL_LIST.length
  ).getValues();

  const plateIdx    = HEADERS_TOTAL_LIST.indexOf('Truck Plate');
  const activityIdx = HEADERS_TOTAL_LIST.indexOf('Activity Status');

  const activityMap = new Map();
  totalListData.forEach(row => {
    const plate = row[plateIdx];
    if (plate) {
      activityMap.set(String(plate).toUpperCase().replace(/\s/g, ''), row[activityIdx]);
    }
  });

  const bannedPlates = [];
  vehicles.forEach(v => {
    const plate = String(v['Truck Plate'] || '').toUpperCase().replace(/\s/g, '');
    if (!plate) return;
    const status = activityMap.get(plate);
    if (status && String(status).toLowerCase() === 'banned') {
      bannedPlates.push(plate);
    }
  });

  if (bannedPlates.length > 0) {
    return {
      isValid: false,
      message: `Xe biển số ${bannedPlates.join(', ')} đang trong tình trạng bị cấm, vui lòng liên hệ PSVN để xử lý.`
    };
  }

  return { isValid: true };
}


function getAllDataForExport(dateString, sessionToken, searchQuery) {
  const userSession = validateSession(sessionToken);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(DATA_SHEET);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const headers = HEADERS_REGISTER;
    let rows = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

    // Lọc theo quyền user (nếu là user thường)
    if (userSession.role === 'user') {
      const compIdx = headers.indexOf('Transportion Company');
      rows = rows.filter(r => r[compIdx] === userSession.contractor);
    }

    // Lọc theo ngày (định dạng dd/MM/yyyy, bỏ dấu ')
    if (dateString) {
      const dateIdx = headers.indexOf('Register Date');
      rows = rows.filter(r => {
        const v = r[dateIdx];
        if (!v) return false;
        let s = (v instanceof Date)
          ? Utilities.formatDate(v, "Asia/Ho_Chi_Minh", "dd/MM/yyyy")
          : String(v).replace(/^'/,'').trim();
        return s === dateString;
      });
    }

    // Lọc theo từ khóa search (nếu có)
    if (searchQuery && String(searchQuery).trim()) {
      const q = String(searchQuery).toLowerCase();
      rows = rows.filter(r => r.some(c => String(c).toLowerCase().indexOf(q) !== -1));
    }

    // Trả về mảng object đã format (ngày/thời gian dạng chuỗi chuẩn)
    return rows.map(r => formatRowForClient_(r, headers));
  } catch (e) {
    Logger.log(e);
    throw new Error('Không thể lấy dữ liệu để xuất file: ' + e.message);
  }
}



function saveData(dataToSave, sessionToken) {
  const userSession = validateSession(sessionToken);
  if (!dataToSave || dataToSave.length === 0) throw new Error('Không có dữ liệu để lưu.');

  if (userSession.role !== 'admin') {
    const timeStatus = checkRegistrationTime();
    if (!timeStatus.isOpen) throw new Error('Đã hết thời gian cho phép đăng ký dữ liệu trong ngày.');
  }

  const activityResult = checkVehicleActivityStatus(dataToSave);
  if (!activityResult.isValid) {
    throw new Error(activityResult.message);
  }

  const validationResult = checkVehiclesAgainstTotalList(dataToSave);
  if (!validationResult.isValid) {
    throw new Error(validationResult.message);
  }

  // Kiểm tra Contract No thuộc đúng đơn vị & Active
  (function () {
    const activeMap = buildActiveContractMap_();
    const invalid = [];

    dataToSave.forEach(rec => {
      const cno = String(rec['Contract No'] || '').trim();
      const comp = String(
        (userSession.role === 'user' ? userSession.contractor : rec['Transportion Company']) || ''
      ).trim().toUpperCase();

      if (!cno || !comp || !activeMap.has(comp) || !activeMap.get(comp).has(cno)) {
        invalid.push(`${cno} (${comp})`);
      }
    });

    if (invalid.length > 0) {
      throw new Error('Sai số hợp đồng, vui lòng kiểm tra lại hợp đồng vận chuyển (Contract No phải thuộc đúng đơn vị và đang Active): ' + invalid.join(', '));
    }
  })();

  const dupCheckRecords = dataToSave.map(r => {
    const obj = Object.assign({}, r);
    if (userSession.role === 'user') {
      obj['Transportion Company'] = userSession.contractor;
    }
    return obj;
  });

  const existingDuplicates = checkForExistingRegistrations(dupCheckRecords, sessionToken);
  if (existingDuplicates && existingDuplicates.length > 0) {
      throw new Error(`Các xe sau đã được đăng ký trong ngày: ${existingDuplicates.join(', ')}. Vui lòng kiểm tra lại.`);
  }

  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(DATA_SHEET);
    const dataArray = dataToSave.map(obj => {
      if (userSession.role === 'user') {
        obj['Transportion Company'] = userSession.contractor;
      }
      obj['Register Date'] = normalizeDate(obj['Register Date']);
      obj['Time'] = normalizeTime(Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "HH:mm:ss"));
      obj['Registration Status'] = 'Pending approval';
      return HEADERS_REGISTER.map(header => obj[header] || "");
    });
    sheet.getRange(sheet.getLastRow() + 1, 1, dataArray.length, HEADERS_REGISTER.length).setValues(dataArray);
    return 'Dữ liệu đã được lưu thành công!';
  } catch (error) { Logger.log(error); throw new Error('Lỗi khi lưu dữ liệu: ' + error.message); }
}

function updateData(rowData, sessionToken) {
  const userSession = validateSession(sessionToken);
  if (!rowData || !rowData.ID) throw new Error('Dữ liệu không hợp lệ hoặc thiếu ID.');

  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(DATA_SHEET);
    const dataRange = sheet.getRange(2, 1, sheet.getLastRow(), HEADERS_REGISTER.length);
    const allValues = dataRange.getValues();
    let rowToUpdate = -1;
    let recordTimeStr = '';

    for (let i = 0; i < allValues.length; i++) {
      if (allValues[i][0] == rowData.ID) {
        rowToUpdate = i + 2;
        recordTimeStr = String(allValues[i][19]).replace("'", "");
        break;
      }
    }

    if (rowToUpdate === -1) throw new Error('Không tìm thấy dòng với ID: ' + rowData.ID);
    
    if (userSession.role === 'user') {
      const timeStatus = checkRegistrationTime();
      if (!timeStatus.isOpen) throw new Error('Đã hết thời gian cho phép chỉnh sửa dữ liệu.');

      const recordHour = parseInt(recordTimeStr.split(':')[0]);
      
      if (recordHour >= 8 && recordHour < 16 && timeStatus.period !== 1) {
        throw new Error('Dữ liệu đăng ký từ 8:00-16:00 chỉ có thể sửa trong khung giờ này.');
      }
      if (recordHour >= 20 && recordHour < 22 && timeStatus.period !== 2) {
        throw new Error('Dữ liệu đăng ký từ 20:00-22:00 chỉ có thể sửa trong khung giờ này.');
      }
    }
    
    if (rowData['Register Date']) {
      rowData['Register Date'] = "'" + rowData['Register Date'];
    }
    rowData['Time'] = "'" + Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "HH:mm:ss");
    const dataArray = HEADERS_REGISTER.map(header => rowData[header] || "");
    sheet.getRange(rowToUpdate, 1, 1, HEADERS_REGISTER.length).setValues([dataArray]);
    return 'Dữ liệu đã được cập nhật thành công!';
  } catch (error) { Logger.log(error); throw new Error('Lỗi khi cập nhật dữ liệu: ' + error.message); }
}

function deleteMultipleData(ids,sessionToken) {
  const userSession = validateSession(sessionToken);
  if (!ids || ids.length === 0) throw new Error('Cần cung cấp ID để xóa.');
  if (userSession.role === 'user') {
    const timeStatus = checkRegistrationTime();
    if (!timeStatus.isOpen) throw new Error('Đã hết thời gian cho phép xóa dữ liệu trong ngày.');
  }
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(DATA_SHEET);
    const idColumnValues = sheet.getRange(2, 1, sheet.getLastRow(), 1).getValues().flat();
    const companyColumnValues = sheet.getRange(2, 16, sheet.getLastRow(), 1).getValues().flat();

    let rowsToDelete = [];
    ids.forEach(id => {
      const rowIndex = idColumnValues.indexOf(id);
      if (rowIndex !== -1) {
          if (userSession.role === 'user' && companyColumnValues[rowIndex] !== userSession.contractor) {
            throw new Error(`Bạn không có quyền xóa xe có ID: ${id}.`);
          }
          rowsToDelete.push(rowIndex + 2);
      }
    });
    if (rowsToDelete.length === 0) throw new Error('Không tìm thấy dòng nào với các ID đã cho.');
    rowsToDelete.sort((a, b) => b - a).forEach(rowNum => {
      sheet.deleteRow(rowNum);
    });
    return `Đã xóa thành công ${rowsToDelete.length} mục.`;
  } catch (error) { Logger.log(error); throw new Error('Lỗi khi xóa dữ liệu: ' + error.message); }
}

function checkForExistingRegistrations(recordsToCheck, sessionToken) {
  validateSession(sessionToken);
  if (!recordsToCheck || recordsToCheck.length === 0) return [];

  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(DATA_SHEET);
    const lastRow = sheet.getLastRow();

    // Thu thập các bản ghi đã tồn tại trong ngày
    const existingKeys = new Set();
    if (lastRow >= 2) {
      const allData = sheet.getRange(2, 1, lastRow - 1, HEADERS_REGISTER.length).getValues();
      allData.forEach(row => {
        const dateStr = Utilities.formatDate(new Date(row[1]), "Asia/Ho_Chi_Minh", "yyyy-MM-dd");
        const plate = String(row[3] || '').toUpperCase().replace(/\s/g, '');
        const company = String(row[15] || '').trim().toUpperCase();
        existingKeys.add(`${dateStr}-${plate}-${company}`);
      });
    }

    // Kiểm tra dữ liệu đầu vào (bao gồm cả trùng lặp trong file)
    const seen = new Set();
    
    const duplicates = [];
    recordsToCheck.forEach(rec => {
      const regDate = normalizeDate(rec['Register Date']);
      const dateStr = regDate ? Utilities.formatDate(regDate, "Asia/Ho_Chi_Minh", "yyyy-MM-dd") : '';
      const plate = String(rec['Truck Plate'] || '').toUpperCase().replace(/\s/g, '');
      const company = String(rec['Transportion Company'] || '').trim().toUpperCase();
      const key = `${dateStr}-${plate}-${company}`;

      if (existingKeys.has(key) || seen.has(key)) {
        duplicates.push(plate);
      }
      seen.add(key);
    });
    
    return duplicates;
  } catch (e) {
    Logger.log(e);
    throw new Error('Lỗi khi kiểm tra dữ liệu trùng lặp: ' + e.message);
  }
}

// =================================================================
// XỬ LÝ DỮ LIỆU DANH SÁCH XE TỔNG
// =================================================================

function getTotalListSummary(sessionToken) {
  const userSession = validateSession(sessionToken);
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(TRUCK_LIST_TOTAL_SHEET);
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      return { total: 0, active: 0, banned: 0 };
    }

    const allData = sheet.getRange(2, 1, lastRow - 1, HEADERS_TOTAL_LIST.length).getValues();
    
    let filteredData = allData;
    if (userSession.role === 'user') {
      const companyIndex = HEADERS_TOTAL_LIST.indexOf('Transportion Company');
      const activityIndex = HEADERS_TOTAL_LIST.indexOf('Activity Status');
      filteredData = allData.filter(row => row[companyIndex] === userSession.contractor &&
        String(row[activityIndex]).toUpperCase() === 'ACTIVE');
    }

    const summary = { total: filteredData.length, active: 0, banned: 0 };
    const activityIdx = HEADERS_TOTAL_LIST.indexOf('Activity Status');

    filteredData.forEach(row => {
      const act = String(row[activityIdx]).toUpperCase();
      if (act === 'ACTIVE') summary.active++;
      else if (act === 'BANNED') summary.banned++;
    });
    
    return summary;
  } catch (e) {
    Logger.log(e);
    throw new Error('Không thể lấy thông tin tóm tắt: ' + e.message);
  }
}


function saveTotalTruckData(dataToSave, sessionToken) {
  const userSession = validateSession(sessionToken);
  if (userSession.role !== 'admin') throw new Error('Chỉ có admin mới được thực hiện chức năng này.');
  if (!dataToSave || dataToSave.length === 0) throw new Error('Không có dữ liệu để lưu.');
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(TRUCK_LIST_TOTAL_SHEET);
    if (!sheet) throw new Error('Không tìm thấy sheet Danh sách xe tổng.');

    // Chuẩn hóa biển số để so sánh
    const norm = s => String(s || '').replace(/\s/g, '').toUpperCase();

    // Lấy tất cả biển số hiện có
    const plateIdx = HEADERS_TOTAL_LIST.indexOf('Truck Plate') + 1;
    const companyIdx = HEADERS_TOTAL_LIST.indexOf('Transportion Company') + 1;

    let existingPlates = {};
    const lastRow = sheet.getLastRow();
    if (lastRow > 1 && plateIdx > 0) {
      const plates = sheet.getRange(2, plateIdx, lastRow - 1, 1).getValues().flat();
      plates.forEach(p => { const k = norm(p); if (k) existingPlates[k] = true; });
    }

    // Kiểm tra trùng lặp trong file và với dữ liệu hiện có
    const inFileSeen = {};
    const skippedInFile = [];
    const skippedExisting = [];
    const rowsToAppend = [];

    dataToSave.forEach(obj => {
      const plate = norm(obj['Truck Plate']);
      const company = obj['Transportion Company'] || '';
      if (!plate) return;

      if (inFileSeen[plate]) {
        skippedInFile.push({ plate: plate, company: company });
        return;
      }
      inFileSeen[plate] = true;

      if (existingPlates[plate]) {
        skippedExisting.push({ plate: plate, company: company });
        return;
      }

      // Bổ sung ngày/giờ và map theo header
      obj['Register Date'] = normalizeDate(Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "dd/MM/yyyy"));
      obj['Time'] = normalizeTime(Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "HH:mm:ss"));
      rowsToAppend.push(HEADERS_TOTAL_LIST.map(h => obj[h] || ""));
    });

    // Append thay vì replace
    let inserted = 0;
    if (rowsToAppend.length > 0) {
      const startRow = (lastRow || 1) + 1;
      sheet.getRange(startRow, 1, rowsToAppend.length, HEADERS_TOTAL_LIST.length).setValues(rowsToAppend);
      inserted = rowsToAppend.length;
    }

    // Trả chi tiết để client hiển thị
    return {
      status: 'ok',
      inserted: inserted,
      skippedExisting: skippedExisting,   // trùng với dữ liệu đang có
      skippedInFile: skippedInFile        // trùng trong chính file upload
    };
  } catch (error) {
    Logger.log(error);
    throw new Error('Lỗi khi lưu dữ liệu danh sách xe tổng: ' + error.message);
  }
}

function deleteTotalListVehicles(ids, sessionToken) {
  const session = validateSession(sessionToken);
  if (session.role !== 'admin') throw new Error('Bạn không có quyền truy cập chức năng này.');
  if (!ids || ids.length === 0) throw new Error('Cần cung cấp ID để xóa.');
  
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(TRUCK_LIST_TOTAL_SHEET);
    const idColumnValues = sheet.getRange(2, 1, sheet.getLastRow(), 1).getValues().flat();
    let rowsToDelete = [];

    ids.forEach(id => {
      const rowIndex = idColumnValues.indexOf(id);
      if (rowIndex !== -1) {
        rowsToDelete.push(rowIndex + 2);
      }
    });

    if (rowsToDelete.length === 0) throw new Error('Không tìm thấy xe nào với các ID đã cho.');
    
    rowsToDelete.sort((a, b) => b - a).forEach(rowNum => {
      sheet.deleteRow(rowNum);
    });

    return `Đã xóa thành công ${rowsToDelete.length} xe.`;
  } catch (error) { Logger.log(error); throw new Error('Lỗi khi xóa xe: ' + error.message); }
}

function updateTotalListVehicle(rowData, sessionToken) {
  const session = validateSession(sessionToken);
  if (session.role !== 'admin') throw new Error('Bạn không có quyền truy cập chức năng này.');
  if (!rowData || !rowData.ID) throw new Error('Dữ liệu không hợp lệ hoặc thiếu ID.');

  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(TRUCK_LIST_TOTAL_SHEET);
    const idColumnValues = sheet.getRange(2, 1, sheet.getLastRow(), 1).getValues().flat();
    let rowToUpdate = -1;

    const rowIndex = idColumnValues.indexOf(rowData.ID);
    if (rowIndex !== -1) {
        rowToUpdate = rowIndex + 2;
    }

    if (rowToUpdate === -1) throw new Error('Không tìm thấy xe với ID: ' + rowData.ID);

    const now = new Date();
    rowData['Register Date'] = Utilities.formatDate(now, "Asia/Ho_Chi_Minh", "dd/MM/yyyy");
    rowData['Time'] = "'" + Utilities.formatDate(now, "Asia/Ho_Chi_Minh", "HH:mm:ss");
    const dataArray = HEADERS_TOTAL_LIST.map(header => rowData[header] || "");
    sheet.getRange(rowToUpdate, 1, 1, HEADERS_TOTAL_LIST.length).setValues([dataArray]);
    return 'Cập nhật thông tin xe thành công!';
  } catch (error) { Logger.log(error); throw new Error('Lỗi khi cập nhật thông tin xe: ' + error.message); }
}

// --- Helpers an toàn cho CacheService ---
function safeGetUserCacheJSON(key) {
  try {
    const v = CacheService.getUserCache().get(key);
    return v ? JSON.parse(v) : null;
  } catch (e) {
    Logger.log('CacheService get error: ' + e);
    return null;
  }
}
function safePutUserCacheJSON(key, obj, seconds) {
  try {
    CacheService.getUserCache().put(key, JSON.stringify(obj), seconds || SESSION_TIMEOUT_SECONDS);
  } catch (e) {
    Logger.log('CacheService put error: ' + e);
  }
}

// --- Fallback lookup trong sheet Users bằng sessionToken ---
function lookupSessionFromSheet(sessionToken) {
  if (!sessionToken) return null;
  try {
    const userSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USERS_SHEET);
    const lastRow = userSheet.getLastRow();
    if (lastRow < 2) return null;

    // Cột G (token) & H (expiry)
    const tokens = userSheet.getRange(2, 7, lastRow - 1, 2).getValues(); // [ [token, expiry], ... ]
    const usernames = userSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const roles = userSheet.getRange(2, 3, lastRow - 1, 1).getValues().flat();
    const contractors = userSheet.getRange(2, 4, lastRow - 1, 1).getValues().flat();

    for (let i = 0; i < tokens.length; i++) {
      const tk = tokens[i][0];
      const exp = tokens[i][1];
      if (tk === sessionToken && exp && new Date().getTime() < new Date(exp).getTime()) {
        return {
          isLoggedIn: true,
          username: usernames[i],
          role: roles[i],
          contractor: contractors[i],
          token: tk
        };
      }
    }
  } catch (e) {
    Logger.log('lookupSessionFromSheet error: ' + e);
  }
  return null;
}

// ==========================
// THAY THẾ validateSession()
// ==========================
function validateSession(sessionToken) {
  // 1) Thử đọc từ cache (an toàn)
  let session = safeGetUserCacheJSON('user_session');
  if (session && session.token === sessionToken) {
    // refresh TTL
    safePutUserCacheJSON('user_session', session, SESSION_TIMEOUT_SECONDS);
    return session;
  }

  // 2) Fallback: tra Users sheet theo token nếu cache lỗi / rỗng
  session = lookupSessionFromSheet(sessionToken);
  if (session) {
    safePutUserCacheJSON('user_session', session, SESSION_TIMEOUT_SECONDS);
    return session;
  }

  // 3) Không tìm thấy → báo chưa đăng nhập (để client xử lý)
  throw new Error('Bạn chưa đăng nhập hoặc phiên đã hết hạn. Vui lòng đăng nhập lại.');
}

// ==========================
// THAY THẾ getUserSession()
// ==========================
function getUserSession() {
  try {
    // Ưu tiên cache user nếu có
    var userCache = CacheService.getUserCache();
    var sessionData = userCache.get('user_session');
    if (sessionData) return JSON.parse(sessionData);
  } catch (e) {
    // Bỏ qua lỗi cache, trả về khách ẩn danh
  }
  return { isLoggedIn: false, role: null, contractor: null };
}

// Trả về Map: company (UPPER) -> Set(contractNo) chỉ chứa hợp đồng Active
function buildActiveContractMap_() {
  const sh = ensureContractSheet_();
  const last = sh.getLastRow();
  const rows = last < 2 ? [] : sh.getRange(2, 1, last - 1, CONTRACT_HEADERS.length).getValues();
  const IDX_NO = 1, IDX_COMP = 3, IDX_STATUS = 4;

  const map = new Map();
  for (const r of rows) {
    const no     = String(r[IDX_NO]   || '').trim();
    const comp   = String(r[IDX_COMP] || '').trim().toUpperCase();
    const status = String(r[IDX_STATUS] || '').trim().toLowerCase();
    if (!no || !comp || status !== 'active') continue;
    if (!map.has(comp)) map.set(comp, new Set());
    map.get(comp).add(no);
  }
  return map;
}

// true nếu Contract No thuộc đúng Company và Active
function isContractActiveForCompany_(contractNo, company) {
  const cno = String(contractNo || '').trim();
  const comp = String(company || '').trim().toUpperCase();
  if (!cno || !comp) return false;
  const m = buildActiveContractMap_();
  return m.has(comp) && m.get(comp).has(cno);
}


//Page hop dong van chuyen
function ensureContractSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(CONTRACT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONTRACT_SHEET);
    sh.getRange(1, 1, 1, CONTRACT_HEADERS.length).setValues([CONTRACT_HEADERS]);
  }
  return sh;
}

function genContractId_() {
  const tz = Session.getScriptTimeZone();
  const ddmm = Utilities.formatDate(new Date(), tz, 'dd/MM');
  const rand = Math.random().toString(36).toUpperCase().replace(/[^A-Z0-9]/g,'').slice(-6);
  return `C${ddmm}${rand}`;
}

function getContractDataServerSide(params) {
  const session = validateSession(params.sessionToken);
  const sh = ensureContractSheet_();
  const last = sh.getLastRow();
  const rows = last < 2 ? [] : sh.getRange(2, 1, last - 1, CONTRACT_HEADERS.length).getValues();

  // map -> object
  let data = rows.map(r => ({
    'ID': r[0] || '',
    'Contract No': r[1] || '',
    'Customer Name': r[2] || '',
    'Transportion Company': r[3] || '',
    'Status': r[4] || ''
  }));

  // User chỉ nhìn thấy theo đơn vị mình
  if (session.role !== 'admin') {
    data = data.filter(x => String(x['Transportion Company'] || '') === String(session.contractor || ''));
  }

  // Search toàn cục
  const q = (params.search && params.search.value ? String(params.search.value) : '').toLowerCase();
  let filtered = q
    ? data.filter(o => Object.values(o).some(v => String(v).toLowerCase().includes(q)))
    : data;

  // Order
  const order = Array.isArray(params.order) ? params.order[0] : null;
  if (order && order.column != null) {
    const columns = ['ID','Contract No','Customer Name','Transportion Company','Status']; // đúng thứ tự trả về cho DataTable
    const key = columns[order.column >= columns.length ? columns.length-1 : order.column];
    const dir = (order.dir || 'asc').toLowerCase();
    filtered.sort((a,b) => (String(a[key]).localeCompare(String(b[key]), undefined, {numeric:true}))
      * (dir === 'desc' ? -1 : 1));
  }

  // Paging
  const start = Number(params.start || 0);
  const length = Number(params.length || 50);
  const page = filtered.slice(start, start + length);

  return {
    draw: Number(params.draw || 1),
    recordsTotal: data.length,
    recordsFiltered: filtered.length,
    data: page
  };
}

function upsertContract(contract, sessionToken) {
  const session = validateSession(sessionToken);
  if (session.role !== 'admin') throw new Error('Bạn không có quyền thực hiện.');

  const { ID, 'Contract No': contractNo, 'Customer Name': customerName,
          'Transportion Company': tc, 'Status': status } = contract;
		  
  const sh = ensureContractSheet_();
  const last = sh.getLastRow();
  const ids = last < 2 ? [] : sh.getRange(2, 1, last - 1, 1).getValues().flat();

  if (ID) {
    const idx = ids.indexOf(ID);
    if (idx === -1) throw new Error('Không tìm thấy ID để cập nhật.');
    sh.getRange(idx + 2, 2, 1, 4).setValues([[contractNo, customerName, tc, status]]); // 4 cột
    return 'Đã cập nhật hợp đồng.';
  } else {
    const newId = genContractId_();
    sh.appendRow([newId, contractNo, customerName, tc, status]);
    return 'Đã tạo hợp đồng mới.';
  }
}


function deleteContracts(ids, sessionToken) {
  const session = validateSession(sessionToken);
  if (session.role !== 'admin') throw new Error('Bạn không có quyền thực hiện.');

  if (!Array.isArray(ids) || !ids.length) return 'Không có mục nào để xoá.';
  const sh = ensureContractSheet_();
  const last = sh.getLastRow();
  if (last < 2) return 'Không có dữ liệu.';

  const allIds = sh.getRange(2,1,last-1,1).getValues().flat();
  const rowsToDelete = [];
  ids.forEach(id => {
    const idx = allIds.indexOf(id);
    if (idx !== -1) rowsToDelete.push(idx + 2); // sheet index
  });

  // xoá từ dưới lên
  rowsToDelete.sort((a,b)=>b-a).forEach(r => sh.deleteRow(r));
  return `Đã xoá ${rowsToDelete.length} hợp đồng.`;
}

//Lấy danh sách Contractor từ sheet Users (dropdown “Transportion Company” ở trang Hợp đồng)
function getContractorOptions() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh  = ss.getSheetByName(USERS_SHEET);
  const n   = sh.getLastRow();
  if (n < 2) return [];
  // Cột Contractor là cột D (index 4) theo cấu trúc bạn đang dùng
  const vals = sh.getRange(2, 4, n - 1, 1).getValues().flat();
  const set  = new Set();
  vals.forEach(v => {
    const s = String(v || '').trim();
    if (s) set.add(s);
  });
  return Array.from(set).sort();
}



//Lấy danh sách "Đơn vị vận chuyển" từ sheet TruckListTotal
function getTransportCompanies() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(TRUCK_LIST_TOTAL_SHEET);
  if (!sh) return [];
  const n = sh.getLastRow();
  if (n < 2) return [];
  const idx = HEADERS_TOTAL_LIST.indexOf('Transportion Company') + 1;
  if (idx <= 0) return [];
  const vals = sh.getRange(2, idx, n - 1, 1).getValues().flat();
  const set = new Set();
  vals.forEach(v => {
    const s = String(v || '').trim();
    if (s) set.add(s);
  });
  return Array.from(set).sort();
}

//Lấy Contract No (Status = Active) cho dropdown “Số HĐ” ở trang Đăng ký xe
function getActiveContractNos(sessionToken) {
  const session = validateSession(sessionToken);
  const sh = ensureContractSheet_();
  const n  = sh.getLastRow();
  if (n < 2) return [];
  const data = sh.getRange(2, 1, n - 1, CONTRACT_HEADERS.length).getValues();

  const IDX_NO = 1, IDX_COMP = 3, IDX_STATUS = 4;
  const out = [];
  const seen = new Set();

  for (const r of data) {
    const status = String(r[IDX_STATUS] || '').trim().toLowerCase();
    if (status !== 'active') continue;

    if (session.role !== 'admin') {
      if (String(r[IDX_COMP] || '') !== String(session.contractor || '')) continue;
    }

    const no = String(r[IDX_NO] || '').trim();
    if (no && !seen.has(no)) { seen.add(no); out.push(no); }
  }
  return out.sort();
}



// ====== GS: Trả về danh sách biển số đang có để đánh dấu trùng ======
function getExistingTruckPlates(sessionToken) {
  const session = validateSession(sessionToken);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(TRUCK_LIST_TOTAL_SHEET);
  if (!sheet) throw new Error('Không tìm thấy sheet Danh sách xe tổng.');

  const plateCol = HEADERS_TOTAL_LIST.indexOf('Truck Plate') + 1;
  const lastRow  = sheet.getLastRow();
  if (plateCol <= 0 || lastRow < 2) return [];

  const norm = s => String(s || '').replace(/\s/g, '').toUpperCase();
  const values = sheet.getRange(2, plateCol, lastRow - 1, 1).getValues().flat();
  return values.map(norm).filter(Boolean);
}



// ====== GS: LƯU NỐI TIẾP VÀO "DANH SÁCH XE TỔNG" ======
function saveTotalListAppend(rows, sessionToken) {
  const session = validateSession(sessionToken);

  if (!rows || !rows.length) return 'Không có dữ liệu để lưu.';

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(TRUCK_LIST_TOTAL_SHEET);
  if (!sh) throw new Error('Không tìm thấy sheet Danh sách xe tổng.');

  const header = HEADERS_TOTAL_LIST;
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, header.length).setValues([header]);
  }

  // Format ngày/giờ
  const pad = n => ('0' + n).slice(-2);
  const formatDateText = d =>
    `'${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()}`; // thêm dấu '
  const formatTimeText = d =>
    `'${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`; // thêm dấu '

  const normPlate = s => String(s || '').replace(/\s/g, '').toUpperCase();

  const values = rows.map(obj => {
    const dNow = new Date();
    const regDate = obj['Register Date']
      ? `${obj['Register Date']}`
      : formatDateText(dNow);
    const regTime = obj['Time']
      ? `${obj['Time']}`
      : formatTimeText(dNow);

    return header.map(h => {
      switch (h) {
        case 'Truck Plate':
        case 'Trailer Plate':
          return normPlate(obj[h]);
        case 'Register Date':
          return regDate;
        case 'Time':
          return regTime;
        default:
          return obj[h] == null ? '' : String(obj[h]);
      }
    });
  });

  const startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, values.length, header.length).setValues(values);

  return `Đã thêm ${values.length} dòng mới vào Danh sách xe tổng.`;
}


// Helper tạo ID ngắn, chữ hoa (15 ký tự)
function generateShortId() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 15).toUpperCase();
}

// === THAY TOÀN BỘ addManualVehicle ===
function addManualVehicle(record, sessionToken) {
  const userSession = validateSession(sessionToken);

  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(DATA_SHEET);
    if (!sheet) throw new Error('Không tìm thấy sheet dữ liệu đăng ký.');

    // Chuẩn hóa/điền thêm các field bắt buộc
    const rowObj = Object.assign({}, record);

    // ✅ NEW: khóa Contractor cho user thường
    if (userSession.role === 'user') {
      rowObj['Transportion Company'] = userSession.contractor || rowObj['Transportion Company'];
    }

    const activityCheck = checkVehicleActivityStatus([{ 'Truck Plate': rowObj['Truck Plate'] }]);
    if (!activityCheck.isValid) {
      throw new Error(activityCheck.message);
    }

    // ✅ NEW: 3 kiểm tra đối chiếu "Danh sách tổng" (dùng đúng thông báo như upload)
    const precheck = checkVehiclesAgainstTotalList([{
      'Truck Plate'         : String(rowObj['Truck Plate'] || '').toUpperCase().replace(/\s/g, ''),
      'Transportion Company': rowObj['Transportion Company']
    }]);
    if (!precheck.isValid) {
      throw new Error(precheck.message);
    }

    // ✅ NEW: kiểm tra Contract No thuộc đúng Contractor & Active
    const contractNo = String(rowObj['Contract No'] || '').trim();
    const company    = String(rowObj['Transportion Company'] || '').trim();
    if (!isContractActiveForCompany_(contractNo, company)) {
      throw new Error('Sai số hợp đồng, vui lòng kiểm tra lại hợp đồng vận chuyển (Contract No phải thuộc đúng đơn vị và đang Active).');
    }

    // ✅ NEW: kiểm tra xe đã đăng ký trong ngày (tái dùng logic của saveData)
    const dup = checkForExistingRegistrations([{
      'Register Date'       : rowObj['Register Date'],
      'Truck Plate'         : rowObj['Truck Plate'],
      'Transportion Company': rowObj['Transportion Company']
    }], sessionToken);
    if (dup && dup.length > 0) {
      throw new Error(`Các xe sau đã được đăng ký trong ngày: ${dup.join(', ')}. Vui lòng kiểm tra lại.`);
    }

    // ID do server tự sinh
    rowObj['ID'] = generateShortId();

    // Lưu ngày dạng text (chỉ thêm 1 dấu ')
    if (rowObj['Register Date']) {
      rowObj['Register Date'] = normalizeDate(rowObj['Register Date']);
    }

    // Thời gian tạo (giữ nguyên cách lưu hiện tại)
    rowObj['Time'] = normalizeTime(Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "HH:mm:ss"));

        // ✅ Bổ sung cột Registration Status
    rowObj['Registration Status'] = 'Pending approval';

    // Ghi theo đúng thứ tự header
    const values = [HEADERS_REGISTER.map(h => rowObj[h] ?? "")];
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, HEADERS_REGISTER.length).setValues(values);

    return 'Đăng ký xe thành công!';
  } catch (e) {
    Logger.log(e);
    throw new Error('Lỗi khi thêm mới: ' + e.message);
  }
}



function _cache(){ return CacheService.getScriptCache(); }
function _getCache(key){ try{ const v=_cache().get(key); return v ? JSON.parse(v) : null; } catch(e){ return null; } }
function _putCache(key,obj,sec){ try{ _cache().put(key, JSON.stringify(obj), sec||60); } catch(e){} }
function _bust(keys){ try{ keys.forEach(k=>_cache().put(k,'x',1)); } catch(e){} }



function _toDateKey(v){
  if (v instanceof Date){
    var dd=('0'+v.getDate()).slice(-2);
    var mm=('0'+(v.getMonth()+1)).slice(-2);
    var yy=v.getFullYear();
    return dd+'/'+mm+'/'+yy;
  }
  if (v == null) return '';
  var s = String(v).trim();
  if (s.startsWith("'")) s = s.slice(1);
  var m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) return s;
  var m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m2) return m2[3]+'/'+m2[2]+'/'+m2[1];
  var d = new Date(s);
  if (!isNaN(d.getTime())) return _toDateKey(d);
  return s;
}

function getXpplSnapshot(payload, sessionToken){
  var userSession = requireAdmin_(sessionToken);
  var dateString = payload && payload.dateString;
  if (!dateString) throw new Error('Thiếu ngày đăng ký');
  var scope = 'ALL'; //admin-only
  var ck = 'SNAP_'+dateString+'_'+scope;
  var cached = _getCache(ck);
  if (cached) return cached;

  var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(DATA_SHEET);
  if (!sh) throw new Error('Không tìm thấy sheet dữ liệu');
  var headers = HEADERS_REGISTER;
  var values = sh.getRange(2,1,Math.max(0, sh.getLastRow()-1), headers.length).getValues();

  var idxDate = headers.indexOf('Register Date');
  var idxCno  = headers.indexOf('Contract No');
  var idxComp = headers.indexOf('Transportion Company');
  var idxRS   = headers.indexOf('Registration Status');

  var rows = [];
  var pending=0, approved=0;
  var per={};

  for (var i=0;i<values.length;i++){
    var r=values[i];
    if (_toDateKey(r[idxDate]) !== dateString) continue;
    if (scope !== 'ALL' && String(r[idxComp]||'').trim() !== scope) continue;

    // (SAU) – dùng helper đã có để đảm bảo DD/MM/YYYY + HH:mm:ss
    var obj = formatRowForClient_(r, headers);
    rows.push(obj);

    var st = String(obj['Registration Status']||'');
    if (st === 'Approved') approved++; else pending++;

    var cno = String(obj['Contract No']||'');
    if (!per[cno]) per[cno]={t:0,a:0};
    per[cno].t++; if (st==='Approved') per[cno].a++;
  }

  var fullyApproved = Object.keys(per).filter(function(k){ var v=per[k]; return v.t>0 && v.a===v.t; });

  var contracts = Object.keys(per).sort();

  var ret = { contracts:contracts, rows:rows, pending:pending, approved:approved, sent:0, fullyApproved:fullyApproved };
  _putCache(ck, ret, 300);
  return ret;
}




function getRegistrationsForApproval(payload, sessionToken){
  var snap = getXpplSnapshot({dateString: payload.dateString}, sessionToken);
  var rows = snap.rows;
  if (payload.contracts && payload.contracts.length){
    var set = {};
    payload.contracts.forEach(function(c){ set[String(c).trim()] = true; });
    rows = rows.filter(function(r){ return set[String(r['Contract No']||'').trim()]; });
  }
  // Đếm nhanh theo TẬP ĐANG LỌC (bám sát DataTable XPPL)
  var pending = 0, approved = 0, sent = 0;
  for (var i=0;i<rows.length;i++){
    var st = String(rows[i]['Registration Status']||'').toLowerCase();
    if (st === 'approved') approved++;
    else if (st === 'pending approval') pending++;
    // Cột "đã gửi XPPL" nếu có (tùy chọn)
    var s = rows[i]['Sent XPPL'] || rows[i]['XPPL Sent'] || rows[i]['Sent to XPPL'];
    if (s === true || String(s).toLowerCase()==='yes' || String(s).toLowerCase()==='sent' || s===1) sent++;
  }
  return { rows: rows, pending: pending, approved: approved, sent: sent };
}


function updateRegistrationStatusBulk(filters, newStatus, sessionToken){
  var userSession = requireAdmin_(sessionToken);
  var dateString = filters && filters.dateString;
  var set = {};
  (filters.contracts||[]).forEach(function(c){ set[String(c).trim()] = true; });
  var idsSelected = (filters.idsSelected||[]).map(String);

  if (!dateString) throw new Error('Thiếu ngày đăng ký.');
  if (!newStatus || (['Approved','Pending approval'].indexOf(newStatus)===-1)) throw new Error('Trạng thái không hợp lệ.');

  var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(DATA_SHEET);
  var headers = HEADERS_REGISTER;
  var values = sh.getRange(2,1,Math.max(0, sh.getLastRow()-1), headers.length).getValues();

  var idxDate = headers.indexOf('Register Date');
  var idxCno  = headers.indexOf('Contract No');
  var idxRS   = headers.indexOf('Registration Status');
  var idxID   = headers.indexOf('ID');
  var idxComp = headers.indexOf('Transportion Company');

  var scope = scope = 'ALL';
  var changed = 0;

  for (var r=0; r<values.length; r++){
    var row = values[r];
    if (_toDateKey(row[idxDate]) !== dateString) continue;
    if (scope !== 'ALL' && String(row[idxComp]||'').trim() !== scope) continue;
    if (Object.keys(set).length && !set[String(row[idxCno]||'').trim()]) continue;
    if (idsSelected.length && idsSelected.indexOf(String(row[idxID]||''))===-1) continue;
    if (row[idxRS] === newStatus) continue;
    row[idxRS] = newStatus;
    sh.getRange(r+2, idxRS+1).setValue(newStatus);
    changed++;
  }

  _bust(['SNAP_'+dateString+'_'+scope, 'SNAP_'+dateString+'_ALL']);
  return 'Đã cập nhật ' + changed + ' dòng.';
}


/**
 * Trả về toàn bộ dữ liệu (đÃ format) theo bộ lọc hiện tại để xuất Excel.
 * params: { dateString?: 'dd/MM/yyyy', search?: string }
 */
function exportRegisteredVehicles(params) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('VehicleData'); // đổi tên nếu bạn dùng sheet khác
  const headers = HEADERS_REGISTER; // mảng header của Registered (đang dùng cho DataTable)
  const range = sheet.getRange(2, 1, Math.max(0, sheet.getLastRow() - 1), headers.length);
  let allData = range.getValues();

  // --- lọc theo ngày (chấp nhận Date thật hoặc text có/dính dấu ')
  if (params && params.dateString) {
    const dateIdx = headers.indexOf('Register Date');
    allData = allData.filter(row => {
      const v = row[dateIdx];
      const cmp = (v instanceof Date)
        ? Utilities.formatDate(v, "Asia/Ho_Chi_Minh", "dd/MM/yyyy")
        : String(v).replace(/^'/, '');
      return cmp === params.dateString;
    });
  }

  // --- lọc theo chuỗi tìm kiếm (áp dụng trên toàn hàng đã bóc dấu ')
  if (params && params.search) {
    const q = params.search.toLowerCase();
    allData = allData.filter(row => row.some(cell => String(cell).replace(/^'/, '').toLowerCase().includes(q)));
  }

  // --- map & format để client đổ thẳng vào Excel
  const rows = allData.map(row => formatRowForClient_(row, headers));
  return { headers: headers, rows: rows };
}



function exportXpplToTemplateDownload(filter, sessionToken) {
  const res = getXpplExportData(filter, sessionToken);
  if (!res || !res.ok) return { ok:false, message:(res && res.errors && res.errors.join('\n')) || 'Không đủ điều kiện để xuất.' };

  const { dateString, contractNo, customerName } = res.filter;
  const rows  = res.rows || [];
  if (!rows.length) return { ok:false, message:'Không có dữ liệu để xuất.' };

  // 1) Copy + convert template -> Google Sheet (dễ ghi định dạng)
  const name = `(${contractNo}_${dateString.replace(/\//g,'-')})-XPPL FORM`;
  const copied = Drive.Files.copy({ title:name, mimeType: MimeType.GOOGLE_SHEETS }, XPPL_TEMPLATE_ID);
  const fileId = copied.id;
  const ss = SpreadsheetApp.openById(fileId);

  try {
    // Ghi header
    const rDate = _getRangeByAnyName_(ss, XPPL_NR_ALIASES.REG_DATE);
    if (rDate) rDate.setValue(dateString);

    const rCus  = _getRangeByAnyName_(ss, XPPL_NR_ALIASES.CUSTOMER_NAME);
    if (rCus) rCus.setValue(customerName);

    const rCon  = _getRangeByAnyName_(ss, XPPL_NR_ALIASES.CONTRACT_NO);
    if (rCon) rCon.setValue(contractNo);

    const rTotal= _getRangeByAnyName_(ss, XPPL_NR_ALIASES.TOTAL_TRUCK);
    if (rTotal) rTotal.setValue(rows.length);

    // Ghi bảng Truck list
    const start = _getRangeByAnyName_(ss, XPPL_NR_ALIASES.TABLE_START);
    if (!start) throw new Error('Không tìm thấy TABLE_START / NR_TABLE_START');

    const sh = start.getSheet();
    const r0 = start.getRow();
    const c0 = start.getColumn();

    const aoa = rows.map((r,i)=>[
      i+1,
      r['Truck Plate']||'',
      r['Country']||'',
      r['Wheel']||'',
      r['Trailer Plate']||'',
      r['Driver Name']||'',
      r['ID/Passport']||'',
      r['Phone number']||'',
      r['Transportion Company']||'',
      r['Subcontractor']||''
    ]);
    if (aoa.length) sh.getRange(r0, c0, aoa.length, 10).setValues(aoa);

    // 2) Export về XLSX (blob) rồi xoá file tạm để không tăng dung lượng
    const exportUrl = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx`;
    const resp = UrlFetchApp.fetch(exportUrl, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions:true });
    const blob = resp.getBlob().setName(name + '.xlsx');

    // Xoá vĩnh viễn file Google Sheet tạm
    try { Drive.Files.remove(fileId); } catch (e) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch(_){ } }

    return { ok:true, filename: blob.getName(), base64: Utilities.base64Encode(blob.getBytes()) };

  } catch (e) {
    try { Drive.Files.remove(fileId); } catch(_){}
    return { ok:false, message: 'Xuất thất bại: ' + (e && e.message) };
  }
}

// ===== XPPL Weighing Station functions =====
function saveXpplWeighingData(rows, sessionToken) {
  const user = requireXpplRole_(sessionToken);
  if (!rows || !rows.length) throw new Error('Không có dữ liệu.');

  // build valid Contract-Customer set from ContractData
  const ssMain = SpreadsheetApp.openById(SPREADSHEET_ID);
  const shCon = ssMain.getSheetByName(CONTRACT_SHEET);
  const lc = shCon.getLastColumn();
  const head = shCon.getRange(1,1,1,lc).getValues()[0];
  const idxCNo = head.indexOf('Contract No');
  const idxCus = head.indexOf('Customer Name');
  const validSet = new Set();
  if (idxCNo !== -1 && idxCus !== -1) {
    const data = shCon.getRange(2,1,Math.max(0, shCon.getLastRow()-1), lc).getValues();
    data.forEach(r => {
      const key = String(r[idxCus]).trim() + '|' + String(r[idxCNo]).trim();
      validSet.add(key);
    });
  }

  const ss = SpreadsheetApp.openById(XPPL_DB_ID);
  const sh = ss.getSheetByName(XPPL_DB_SHEET);
  const tz = ss.getSpreadsheetTimeZone() || 'Asia/Ho_Chi_Minh';
  const prefix = Utilities.formatDate(new Date(), tz, 'dd-MM') + '-';
  const dateCols = ['Date In','Date Out'];
  const timeCols = ['Time In','Time Out'];
  let lr = sh.getLastRow();
  if (lr === 0) {
    sh.getRange(1, 1, 1, XPPL_DB_HEADERS.length).setValues([XPPL_DB_HEADERS]);
    lr = 1;
  }

  const toSave = rows.map(r => {
    const key = String(r['Customer Name']||'').trim() + '|' + String(r['ContractNo']||'').trim();
    if (validSet.size && !validSet.has(key)) {
      throw new Error('Sai tên khách hàng hoặc số hợp đồng: ' + key);
    }
    const arr = XPPL_DB_HEADERS.map(h => {
      let v = r[h] || '';
      if (dateCols.indexOf(h) !== -1) v = parseExcelDate_(v);
      else if (timeCols.indexOf(h) !== -1) v = parseExcelTime_(v);
      return v;
    });
    arr[0] = prefix + Math.floor(Math.random()*1e7).toString().padStart(7,'0');
    arr[35] = user.username || user.user || user.email || '';
    return arr;
  });

  if (toSave.length) {
    const startRow = lr + 1;
    sh.getRange(startRow, 1, toSave.length, XPPL_DB_HEADERS.length).setValues(toSave);
    dateCols.forEach(col => {
      const c = XPPL_DB_HEADERS.indexOf(col) + 1;
      if (c > 0) sh.getRange(startRow, c, toSave.length, 1).setNumberFormat('dd/MM/yyyy');
    });
    timeCols.forEach(col => {
      const c = XPPL_DB_HEADERS.indexOf(col) + 1;
      if (c > 0) sh.getRange(startRow, c, toSave.length, 1).setNumberFormat('HH:mm:ss');
    });
  }
  return 'Đã lưu ' + toSave.length + ' dòng.';
}

function getXpplWeighingData(filter, sessionToken) {
  requireXpplRole_(sessionToken);
  const ss = SpreadsheetApp.openById(XPPL_DB_ID);
  const sh = ss.getSheetByName(XPPL_DB_SHEET);
  const lr = sh.getLastRow();
  if (lr < 2) {
    return { data: [], summary: { trucks:0, weight:0 }, contracts: [], customers: [] };
  }
  const s = v => String(v == null ? '' : v).replace(/^'+/, '').trim();
  const filterDate = s(filter && filter.date);
  const dateKey = filterDate ? _toDateKey(filterDate) : null;
  if (!dateKey) {
    return { data: [], summary: { trucks:0, weight:0 }, contracts: [], customers: [] };
  }

  const data = sh.getRange(2,1,lr-1,XPPL_DB_HEADERS.length).getValues();
  const idxDate = XPPL_DB_HEADERS.indexOf('Date Out');
  const idxContract = XPPL_DB_HEADERS.indexOf('ContractNo');
  const idxCustomer = XPPL_DB_HEADERS.indexOf('Customer Name');
  const idxNet = XPPL_DB_HEADERS.indexOf('Net Weight');

  const filterContract = s(filter && filter.contractNo);
  const filterCustomer = s(filter && filter.customerName);

  const contracts = new Set();
  const customers = new Set();
  const rows = [];
  let totalWeight = 0;
  data.forEach(r => {
    const d = _toDateKey(r[idxDate]);
    if (d !== dateKey) return;
    const cno = s(r[idxContract]);
    const cus = s(r[idxCustomer]);
    const net = Number(r[idxNet]) || 0;
    contracts.add(cno);
    customers.add(cus);
    if (filterContract && filterContract !== cno) return;
    if (filterCustomer && filterCustomer !== cus) return;
    rows.push(formatRowForClient_(r, XPPL_DB_HEADERS));
    totalWeight += net;
  });

  return {
    data: rows,
    summary: { trucks: rows.length, weight: totalWeight },
    contracts: Array.from(contracts).sort(),
    customers: Array.from(customers).sort()
  };
}

// ===== WEIGHING RESULT HELPERS =====
function matchTransportionCompanies(filter, sessionToken) {
  const user = requireXpplRole_(sessionToken);
  const main = SpreadsheetApp.openById(SPREADSHEET_ID);
  const totalSh = main.getSheetByName(TRUCK_LIST_TOTAL_SHEET);
  const totalLast = totalSh.getLastRow();
  const totalHead = totalLast > 0 ? totalSh.getRange(1,1,1,totalSh.getLastColumn()).getValues()[0] : [];
  const idxPlateTL = totalHead.indexOf('Truck Plate');
  const idxCompTL = totalHead.indexOf('Transportion Company');
  const plateMap = new Map();
  if (idxPlateTL > -1 && idxCompTL > -1 && totalLast > 1) {
    const vals = totalSh.getRange(2,1,totalLast-1,totalSh.getLastColumn()).getValues();
    vals.forEach(r => {
      const plate = String(r[idxPlateTL]||'').replace(/\s/g,'').toUpperCase();
      if (plate) plateMap.set(plate, String(r[idxCompTL]||'').trim());
    });
  }

  const ss = SpreadsheetApp.openById(XPPL_DB_ID);
  const sh = ss.getSheetByName(XPPL_DB_SHEET);
  const lr = sh.getLastRow();
  if (lr < 2) return 'Không có dữ liệu.';

  const headers = XPPL_DB_HEADERS;
  const idxTruck = headers.indexOf('Truck No');
  const idxComp = headers.indexOf('Transportion Company');
  const idxDate = headers.indexOf('Changed Date');
  const idxTime = headers.indexOf('Changed Time');
  const idxUser = headers.indexOf('Username');
  const idxDateOut = headers.indexOf('Date Out');

  const f = filter || {};
  const from = _toDateKey(f.dateFrom);
  const to = _toDateKey(f.dateTo);  

  const data = sh.getRange(2,1,lr-1,headers.length).getValues();
  const tz = ss.getSpreadsheetTimeZone() || 'Asia/Ho_Chi_Minh';
  const now = new Date();
  const dStr = Utilities.formatDate(now, tz, 'dd/MM/yyyy');
  const tStr = Utilities.formatDate(now, tz, 'HH:mm:ss');
  const uname = user.username || user.user || user.email || '';

  const updates = [];
  data.forEach((r,i) => {
    const dk = _toDateKey(r[idxDateOut]);
    if (from && dk < from) return;
    if (to && dk > to) return;
    const plate = String(r[idxTruck]||'').replace(/\s/g,'').toUpperCase();
    r[idxComp] = plateMap.get(plate) || 'Unknown';
    r[idxDate] = dStr;
    r[idxTime] = tStr;
    r[idxUser] = uname;
    updates.push({row: i+2, values: r});
  });

  if (!updates.length) return 'Không tìm thấy dữ liệu phù hợp.';

  updates.sort((a,b) => a.row - b.row);
  let start = updates[0].row;
  let block = [updates[0].values];
  for (let j = 1; j < updates.length; j++) {
    const cur = updates[j];
    const prev = updates[j-1];
    if (cur.row === prev.row + 1) {
      block.push(cur.values);
    } else {
      sh.getRange(start,1,block.length,headers.length).setValues(block);
      start = cur.row;
      block = [cur.values];
    }
  }
  sh.getRange(start,1,block.length,headers.length).setValues(block);

  return 'Đã đối chiếu ' + updates.length + ' dòng.';
}


function formatWeighResultCell_(header, value) {
  if (header === 'Register Date' || header === 'Date In' || header === 'Date Out' || header === 'Changed Date') {
    return formatDateForClient(value);
  }
  if (header === 'Time' || header === 'Time In' || header === 'Time Out' || header === 'Changed Time') {
    return formatTimeForClient(value);
  }
  const v = stripLeadingApostrophe(value);
  return v == null ? '' : v;
}

function weighResultRowMatchesQuery_(row, headers, queryLower) {
  for (var i = 0; i < headers.length; i++) {
    var text = formatWeighResultCell_(headers[i], row[i]);
    if (text != null && String(text).toLowerCase().indexOf(queryLower) !== -1) {
      return true;
    }
  }
  return false;
}

function getWeighResultData(params) {
  const session = validateSession(params.sessionToken);
  const sh = SpreadsheetApp.openById(XPPL_DB_ID).getSheetByName(XPPL_DB_SHEET);
  const headers = XPPL_DB_HEADERS;
  const lr = sh.getLastRow();  
  const f = params.filter || {};
  const from = _toDateKey(f.dateFrom);
  const to = _toDateKey(f.dateTo);
  const contracts = Array.isArray(f.contracts) && f.contracts.length ? f.contracts : null;
  const companies = Array.isArray(f.companies) && f.companies.length ? f.companies : null;
  const draw = Number(params.draw || 1);
  const empty = { draw: draw, recordsTotal: 0, recordsFiltered: 0, data: [], counts: { unassigned: 0, unknown: 0, assigned: 0 } };

  if (!from && !to && !contracts && !companies && !(params.search && params.search.value)) {
    return empty;
  }

  if (lr < 2) {
    return empty;
  }

  const idxDateOut = headers.indexOf('Date Out');
  const idxContract = headers.indexOf('ContractNo');
  const idxCompany = headers.indexOf('Transportion Company');
  if (idxDateOut === -1 || idxContract === -1 || idxCompany === -1) {
    return empty;    
  }

  let rows = [];

  if (from || to) {
    const dateValues = sh.getRange(2, idxDateOut + 1, lr - 1, 1).getValues();
    let start = 0;
    let end = dateValues.length - 1;
    if (from) {
      while (start <= end && _toDateKey(dateValues[start][0]) < from) start++;
    }
    if (to) {
      while (end >= start && _toDateKey(dateValues[end][0]) > to) end--;
    }
    if (end >= start) {
      rows = sh.getRange(start + 2, 1, end - start + 1, headers.length).getValues();
    }
  } else {
    rows = sh.getRange(2, 1, lr - 1, headers.length).getValues();    
  }

  if (!rows.length) {
    return empty;
  }

  const isUser = session.role === 'user';
  const userCompany = isUser ? String(session.contractor || '') : '';
  const contractSet = contracts ? new Set(contracts.map(function(v){ return String(v || ''); })) : null;
  const companySet = companies ? new Set(companies.map(function(v){ return String(v || ''); })) : null;

  const baseRows = [];
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var dateKey = _toDateKey(row[idxDateOut]);
    if (from && (!dateKey || dateKey < from)) continue;
    if (to && (!dateKey || dateKey > to)) continue;

    var rowContract = String(stripLeadingApostrophe(row[idxContract]) || '');
    if (contractSet && !contractSet.has(rowContract)) continue;

    var rowCompany = String(stripLeadingApostrophe(row[idxCompany]) || '');
    if (isUser && rowCompany !== userCompany) continue;
    if (companySet && !companySet.has(rowCompany)) continue;

    baseRows.push(row);
  }
  
  var totalRecords = baseRows.length;
  if (!totalRecords) {
    return { draw: draw, recordsTotal: 0, recordsFiltered: 0, data: [], counts: { unassigned: 0, unknown: 0, assigned: 0 } };
  }

  const searchValue = (params.search && params.search.value ? String(params.search.value) : '').toLowerCase();
  let filteredForSearch = baseRows;
  if (searchValue) {
    filteredForSearch = [];
    for (var j = 0; j < baseRows.length; j++) {
      if (weighResultRowMatchesQuery_(baseRows[j], headers, searchValue)) {
        filteredForSearch.push(baseRows[j]);
      }
    }
  }

  const counts = { unassigned: 0, unknown: 0, assigned: 0 };
  for (var k = 0; k < filteredForSearch.length; k++) {
    var comp = String(stripLeadingApostrophe(filteredForSearch[k][idxCompany]) || '').trim();
    if (!comp) counts.unassigned++;
    else if (comp.toLowerCase() === 'unknown') counts.unknown++;
    else counts.assigned++;
  }

  let filtered = filteredForSearch;
  if (params.onlyUnknown) {
    filtered = filtered.filter(function(row) {
      var comp = String(stripLeadingApostrophe(row[idxCompany]) || '').trim().toLowerCase();
      return comp === 'unknown';
    });
  } else if (params.excludeUnknown) {
    filtered = filtered.filter(function(row) {
      var comp = String(stripLeadingApostrophe(row[idxCompany]) || '').trim().toLowerCase();
      return comp !== 'unknown';
    });
  }

  const order = Array.isArray(params.order) ? params.order[0] : null;
  if (order && order.column != null) {
    const offset = session.role === 'admin' ? 2 : 0;
    const idx = Number(order.column) - offset;
    if (idx >= 0 && idx < headers.length) {
      const dir = (order.dir || 'asc').toLowerCase() === 'desc' ? -1 : 1;
      filtered.sort(function(a, b) {
        const va = formatWeighResultCell_(headers[idx], a[idx]);
        const vb = formatWeighResultCell_(headers[idx], b[idx]);
        return String(va).localeCompare(String(vb), undefined, { numeric: true }) * dir;
      });
    }
  }

  const start = Math.max(0, Number(params.start || 0));
  const length = Math.max(0, Number(params.length || 50));
  const pageRows = filtered.slice(start, start + length);
  const data = pageRows.map(function(row) {
    return formatRowForClient_(row, headers);
  });

  return {
    draw: draw,
    recordsTotal: totalRecords,
    recordsFiltered: filtered.length,
    data: data,
    counts: counts
  };
}

function updateWeighResultCompany(payload, sessionToken) {
  const user = requireAdmin_(sessionToken);
  const { ID, 'Transportion Company': company } = payload || {};
  if (!ID) throw new Error('Thiếu ID.');

  const ss = SpreadsheetApp.openById(XPPL_DB_ID);
  const sh = ss.getSheetByName(XPPL_DB_SHEET);
  const lr = sh.getLastRow();
  if (lr < 2) throw new Error('Không có dữ liệu.');

  const ids = sh.getRange(2,1,lr-1,1).getValues().flat();
  const rowIdx = ids.indexOf(ID);
  if (rowIdx === -1) throw new Error('Không tìm thấy ID.');

  const idxComp = XPPL_DB_HEADERS.indexOf('Transportion Company') + 1;
  const idxDate = XPPL_DB_HEADERS.indexOf('Changed Date') + 1;
  const idxTime = XPPL_DB_HEADERS.indexOf('Changed Time') + 1;
  const idxUser = XPPL_DB_HEADERS.indexOf('Username') + 1;
  const tz = ss.getSpreadsheetTimeZone() || 'Asia/Ho_Chi_Minh';
  const now = new Date();

  sh.getRange(rowIdx + 2, idxComp).setValue(company);
  sh.getRange(rowIdx + 2, idxDate).setValue(Utilities.formatDate(now, tz, 'dd/MM/yyyy'));
  sh.getRange(rowIdx + 2, idxTime).setValue(Utilities.formatDate(now, tz, 'HH:mm:ss'));
  sh.getRange(rowIdx + 2, idxUser).setValue(user.username || user.user || user.email || '');

  return 'Đã cập nhật.';
}

function deleteWeighResults(ids, sessionToken) {
  const user = requireAdmin_(sessionToken);
  if (!Array.isArray(ids) || !ids.length) return 'Không có ID.';
  const ss = SpreadsheetApp.openById(XPPL_DB_ID);
  const sh = ss.getSheetByName(XPPL_DB_SHEET);
  const lr = sh.getLastRow();
  if (lr < 2) return 'Không có dữ liệu.';
  const idList = sh.getRange(2,1,lr-1,1).getValues().flat();
  const rows = ids.map(id => idList.indexOf(id)).filter(i => i !== -1).map(i => i + 2).sort((a,b) => b - a);
  rows.forEach(r => sh.deleteRow(r));
  return 'Đã xoá ' + rows.length + ' dòng.';
}

/*** END ***/
