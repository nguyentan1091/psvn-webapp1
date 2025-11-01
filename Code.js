// =================================================================
// CẤU HÌNH GOOGLE SHEETS
const SPREADSHEET_ID = '1IHBdQFecC1_JT17dQOTxq-NEz1-HvXHHuSVwg5TGGIM'; 
const VEHICLE_REGISTRATION_CACHE_KEY = 'vehicle_registration_supabase';
// === XPPL Weighing Station database ===
const XPPL_DB_ID = '1LJGbMLFU8GnETecJ3i_j_fL5GWz5W1zST5bCQ5A5o3w';
const XPPL_DB_SHEET = 'XPPL-Database';
const XPPL_DB_HEADERS = [
  'ID','No.','W.ID','Weighing Type','TicketID','Truck No','Date In','Time In','Date Out','Time Out',
  'Weight In','Weight Out','Net Weight','Product Name','CoalSource','ProductionCode','Customer Name',
  'DriverName','Id/Passport','CargoLotNo','CargoName','CargoCompany','PackUnit','PackQtt','OrderNo',
  'ContractNo','InvoiceNo','CoNo','OVS_DMT','Plant','Trailer No','Truck Country','Truck Type','WeighStationCode',
  'Note','CreateUser','Transportation Company','Changed Date','Changed Time','Username'
];
const XPPL_DB_FIELD_MAP = {
  'ID': 'id',
  'No.': 'no',
  'W.ID': 'w_id',
  'Weighing Type': 'weighing_type',
  'TicketID': 'ticket_id',
  'Truck No': 'truck_no',
  'Date In': 'date_in',
  'Time In': 'time_in',
  'Date Out': 'date_out',
  'Time Out': 'time_out',
  'Weight In': 'weight_in',
  'Weight Out': 'weight_out',
  'Net Weight': 'net_weight',
  'Product Name': 'product_name',
  'CoalSource': 'coal_source',
  'ProductionCode': 'production_code',
  'Customer Name': 'customer_name',
  'DriverName': 'driver_name',
  'Id/Passport': 'id_passport',
  'CargoLotNo': 'cargo_lot_no',
  'CargoName': 'cargo_name',
  'CargoCompany': 'cargo_company',
  'PackUnit': 'pack_unit',
  'PackQtt': 'pack_qtt',
  'OrderNo': 'order_no',
  'ContractNo': 'contract_no',
  'InvoiceNo': 'invoice_no',
  'CoNo': 'co_no',
  'OVS_DMT': 'ovs_dmt',
  'Plant': 'plant',
  'Trailer No': 'trailer_no',
  'Truck Country': 'truck_country',
  'Truck Type': 'truck_type',
  'WeighStationCode': 'weigh_station_code',
  'Note': 'note',
  'CreateUser': 'create_user',
  'Transportation Company': 'transportation_company',
  'Changed Date': 'changed_date',
  'Changed Time': 'changed_time',
  'Username': 'username'
};
// === XPPL TEMPLATE (Google Sheet chứa mẫu in) ===
// ID của file mẫu bạn gửi: https://docs.google.com/spreadsheets/d/18tVwSBr7tLU3uekL8Ay6gyrc4YFIFlS2/...
const XPPL_TEMPLATE_ID = '1p8n8ffm81NaxSWB5F7Wn1GhsaBrQ21XttaWmX5yvBl4';
const XPPL_DB_COLUMN_TYPES = {
  'ID': 'text',
  'No.': 'text',
  'W.ID': 'text',
  'Weighing Type': 'text',
  'TicketID': 'text',
  'Truck No': 'text',
  'Date In': 'date',
  'Time In': 'time',
  'Date Out': 'date',
  'Time Out': 'time',
  'Weight In': 'text',
  'Weight Out': 'text',
  'Net Weight': 'text',
  'Product Name': 'text',
  'CoalSource': 'text',
  'ProductionCode': 'text',  
  'Customer Name': 'text',
  'DriverName': 'text',
  'Id/Passport': 'text',
  'CargoLotNo': 'text',
  'CargoName': 'text',
  'CargoCompany': 'text',
  'PackUnit': 'text',
  'PackQtt': 'text',
  'OrderNo': 'text',
  'ContractNo': 'text',
  'InvoiceNo': 'text',
  'CoNo': 'text',
  'OVS_DMT': 'text',
  'Plant': 'text',  
  'Trailer No': 'text',
  'Truck Country': 'text',
  'Truck Type': 'text',
  'WeighStationCode': 'text',
  'Note': 'text',
  'CreateUser': 'text',
  'Transportation Company': 'text',
  'Changed Date': 'date',
  'Changed Time': 'time',
  'Username': 'text'
};

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
  'Transportation Company',  // I
  'Subcontractor'          // J
];


const MAX_LOGIN_ATTEMPTS = 10;
const LOCKOUT_DURATION_1 = 10 * 60 * 1000; // 10 minutes
const LOCKOUT_DURATION_2 = 60 * 60 * 1000; // 1 hour
const SESSION_TIMEOUT_SECONDS = 30 * 60; // 30 minutes

// Default supervision account configuration
const SUPERVISION_DEFAULT_USERNAME = 'LA';
const SUPERVISION_DEFAULT_PASSWORD = 'CRLF@LA111';
const SUPERVISION_DEFAULT_ROLE = 'User-Supervision';

const SERVER_SIDE_CACHE_TTL_SECONDS = 45;
const SHEET_CACHE_VERSION_PREFIX = 'sheet_cache_version::';

// ================= SUPABASE CONFIGURATION =================
const SUPABASE_URL = 'https://mbyrruczihniewdvxokj.supabase.co';
const SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im1ieXJydWN6aWhuaWV3ZHZ4b2tqIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc1ODMzNDE5OSwiZXhwIjoyMDczOTEwMTk5fQ.78Xkt_3GCdvfEL8R0313MVeOEeuXBECDYZ3QHh6XigE';
const SUPABASE_APP_USERS_ENDPOINT = '/rest/v1/app_users';
const SUPABASE_VEHICLE_REG_ENDPOINT = '/rest/v1/vehicle_registration';
const SUPABASE_AUTH_LOGIN_HISTORY_ENDPOINT = '/rest/v1/auth_login_history';
const SUPABASE_HISTORY_VEHICLE_REG_ENDPOINT = '/rest/v1/history_vehicle_registration';
const SUPABASE_CONTRACT_DATA_ENDPOINT = '/rest/v1/contract_data';
const SUPABASE_XPPL_DATABASE_ENDPOINT = '/rest/v1/xppl_database';
const SUPABASE_TRUCK_LIST_TOTAL_ENDPOINT = '/rest/v1/truck_list_total';
const CONTRACT_DATA_SELECT_FIELDS = [
  'id',
  'contract_no',
  'customer_name',
  'transportation_company',
  'status',
  'created_at',
  'created_by',
  'updated_at',
  'updated_by'
];
const SUPABASE_USER_CACHE_PREFIX = 'supabase_user_cache::';
const SUPABASE_USER_CACHE_TTL_SECONDS = 60;
const SUPABASE_USER_MISS_CACHE_PREFIX = 'supabase_user_miss::';
const SUPABASE_USER_MISS_CACHE_TTL_SECONDS = 45;
const SUPABASE_IN_FILTER_BATCH_SIZE = 20;
const SUPABASE_USER_SELECT_FIELDS = [
  'username',
  'password_hash',
  'role',
  'contractor',
  'customer_name',
  'active_session_token',
  'session_token_expiry',
  'security_code',
  'password_last_updated'
];

const VEHICLE_REGISTRATION_SELECT_FIELDS = [
  'id',
  'register_date',
  'contract_no',
  'truck_plate',
  'country',
  'wheel',
  'trailer_plate',
  'truck_weight',
  'pay_load',
  'container_no1',
  'container_no2',
  'driver_name',
  'id_passport',
  'phone_number',
  'destination_est',
  'transportation_company',
  'subcontractor',
  'vehicle_status',
  'registration_status',
  'time',
  'created_at',
  'updated_at',
  'created_by'
];

const VEHICLE_REGISTRATION_COLUMN_MAP = {
  'ID': 'id',
  'Register Date': 'register_date',
  'Contract No': 'contract_no',
  'Truck Plate': 'truck_plate',
  'Country': 'country',
  'Wheel': 'wheel',
  'Trailer Plate': 'trailer_plate',
  'Truck weight': 'truck_weight',
  'Pay load': 'pay_load',
  'Container No1': 'container_no1',
  'Container No2': 'container_no2',
  'Driver Name': 'driver_name',
  'ID/Passport': 'id_passport',
  'Phone number': 'phone_number',
  'Destination EST': 'destination_est',
  'Transportation Company': 'transportation_company',
  'Subcontractor': 'subcontractor',
  'Vehicle Status': 'vehicle_status',
  'Registration Status': 'registration_status',
  'Time': 'time'
};

function buildSupabaseUrl_(path) {
  const base = SUPABASE_URL.replace(/\/$/, '');
  if (!path) return base;
  return base + (path[0] === '/' ? path : '/' + path);
}

function supabaseRequest_(path, options) {
  const opts = options || {};
  const method = (opts.method || 'GET').toUpperCase();
  const headers = Object.assign(
    {
      apikey: SUPABASE_KEY,
      Authorization: 'Bearer ' + SUPABASE_KEY
    },
    opts.headers || {}
  );

  if (method !== 'GET' && !headers['Content-Type'] && !headers['content-type']) {
    headers['Content-Type'] = 'application/json';
  }

  const request = {
    method: method,
    headers: headers,
    muteHttpExceptions: true
  };

  if (opts.payload != null) {
    request.payload = typeof opts.payload === 'string' ? opts.payload : JSON.stringify(opts.payload);
  }

  if (opts.timeout != null) {
    request.timeout = opts.timeout;
  }

  const url = buildSupabaseUrl_(path);
  const returnResponse = !!opts.returnResponse;  
  let response;
  try {
    response = UrlFetchApp.fetch(url, request);
  } catch (fetchError) {
    throw new Error('Supabase request failed: ' + fetchError);
  }

  const status = response.getResponseCode();
  const text = response.getContentText();
  let data = null;
  if (text) {
    try {
      data = JSON.parse(text);
    } catch (parseError) {
      data = text;
    }
  }

  if (status >= 200 && status < 300) {
    if (returnResponse) {
      return {
        data: data,
        status: status,
        headers: response.getAllHeaders(),
        raw: response
      };
    }    
    return data;
  }

  const message = data && data.message ? data.message : text;
  const error = new Error('Supabase request failed (' + status + '): ' + (message || 'Unknown error'));
  if (returnResponse) {
    error.response = response;
  }
  throw error;
}

function fetchContractDataRows_(selectFields, filterParams) {
  const fields = Array.isArray(selectFields) && selectFields.length
    ? selectFields.join(',')
    : '*';
  let query = SUPABASE_CONTRACT_DATA_ENDPOINT + '?select=' + encodeURIComponent(fields);
  if (Array.isArray(filterParams) && filterParams.length) {
    query += '&' + filterParams.join('&');
  }
  try {
    const rows = supabaseRequest_(query);
    return Array.isArray(rows) ? rows : [];
  } catch (e) {
    Logger.log('fetchContractDataRows_ error: ' + e);
    return [];
  }
}

function toSupabaseDateString_(value) {
  const normalized = normalizeDate(value);
  if (!normalized) return null;
  const timezone = Session.getScriptTimeZone() || 'Asia/Ho_Chi_Minh';
  return Utilities.formatDate(normalized, timezone, 'yyyy-MM-dd');
}

function mapVehicleRegistrationRowToArray_(row, headers) {
  if (!row) return headers.map(() => '');
  return headers.map(function (header) {
    const column = VEHICLE_REGISTRATION_COLUMN_MAP[header];
    if (!column) return '';
    const value = row[column];
    if (value == null) return '';
    if (header === 'Register Date') {
      const dt = parseSupabaseTimestamp_(value);
      return dt || value;
    }
    if (header === 'Time') {
      const dt = parseSupabaseTimestamp_(value);
      return dt || value;
    }
    return value;
  });
}

function buildVehicleRegistrationPayload_(record, options) {
  const opts = options || {};
  const payload = {};
  Object.keys(VEHICLE_REGISTRATION_COLUMN_MAP).forEach(function (header) {
    const column = VEHICLE_REGISTRATION_COLUMN_MAP[header];
    if (!column) return;
    const value = record[header];

    if (header === 'ID') {
      if (value !== '' && value !== null && value !== undefined) {
        payload[column] = value;
      }
      return;
    }
    if (header === 'Register Date') {
      const iso = toSupabaseDateString_(value);
      if (iso) payload[column] = iso;
      return;
    }
    if (header === 'Time') {
      if (value) payload[column] = value instanceof Date ? value.toISOString() : String(value);
      return;
    }
    if (value === '' || value === null || value === undefined) {
      if (opts.includeNulls) payload[column] = null;
      return;
    }
    payload[column] = value;
  });
  return payload;
}

function mapTruckListRowToArray_(row, headers) {
  if (!row) return headers.map(() => '');
  return headers.map(function (header) {
    const column = TRUCK_LIST_TOTAL_COLUMN_MAP[header];
    if (!column) return '';
    const value = row[column];
    if (value == null) return '';
    if (header === 'Register Date' || header === 'Time' || header === 'Updated At') {
      const dt = parseSupabaseTimestamp_(value);
      return dt || value;
    }
    return value;
  });
}

function buildTruckListPayload_(record, options) {
  const opts = options || {};
  const payload = {};
  Object.keys(TRUCK_LIST_TOTAL_COLUMN_MAP).forEach(function (header) {
    const column = TRUCK_LIST_TOTAL_COLUMN_MAP[header];
    if (!column) return;
    if (header === 'ID') {
      if (!opts.includeId) return;
      const value = record[header];
      if (value !== '' && value !== null && value !== undefined) {
        payload[column] = value;
      }
      return;
    }
    if (header === 'Register Date' || header === 'Time') {
      // register_date và time được Supabase cập nhật tự động.
      return;
    }
    const value = record[header];
    if (value === '' || value === null || value === undefined) {
      if (opts.includeNulls) payload[column] = null;
      return;
    }
    if (value instanceof Date) {
      payload[column] = value.toISOString();
      return;
    }
    payload[column] = value;
  });
  return payload;
}

function chunkArray_(array, size) {
  if (!Array.isArray(array) || !array.length) return [];
  const chunkSize = size && size > 0 ? size : array.length;
  const chunks = [];
  for (let i = 0; i < array.length; i += chunkSize) {
    chunks.push(array.slice(i, i + chunkSize));
  }
  return chunks;
}

function buildSupabaseInFilter_(column, values) {
  if (!column || !Array.isArray(values) || !values.length) return '';
  const sanitized = values
    .map(function (value) {
      return String(value == null ? '' : value).trim();
    })
    .filter(function (value) { return value.length > 0; })
    .map(function (value) { return value.replace(/"/g, '""'); });
  if (!sanitized.length) return '';
  const quoted = sanitized.map(function (value) { return '"' + value + '"'; });
  return column + '=in.' + encodeURIComponent('(' + quoted.join(',') + ')');
}

function buildSupabaseUserCacheKey_(username) {
  const normalized = String(username == null ? '' : username).trim().toLowerCase();
  return normalized ? SUPABASE_USER_CACHE_PREFIX + normalized : '';
}

function buildSupabaseUserMissCacheKey_(username) {
  const normalized = String(username == null ? '' : username).trim().toLowerCase();
  return normalized ? SUPABASE_USER_MISS_CACHE_PREFIX + normalized : '';
}

function cacheSupabaseUserRecord_(username, record) {
  const key = buildSupabaseUserCacheKey_(username);
  if (!key || !record || typeof record !== 'object') return;
  if (!Object.prototype.hasOwnProperty.call(record, 'password_hash')) return;
  safeScriptCachePutJSON_(key, record, SUPABASE_USER_CACHE_TTL_SECONDS);
  removeSupabaseUserMissCache_(username);
}

function removeSupabaseUserCache_(username) {
  const key = buildSupabaseUserCacheKey_(username);
  if (!key) return;
  safeScriptCacheRemove_(key);
  removeSupabaseUserMissCache_(username);
}

function getSupabaseUserByUsername_(username, options) {
  const normalized = String(username == null ? '' : username).trim();
  if (!normalized) return null;

  const opts = options || {};
  const cacheKey = buildSupabaseUserCacheKey_(normalized);
  const missKey = buildSupabaseUserMissCacheKey_(normalized);
  if (!opts.forceRefresh && missKey) {
    const missHit = safeScriptCacheGetJSON_(missKey);
    if (missHit) return null;
  }
  if (!opts.forceRefresh && cacheKey) {
    const cached = safeScriptCacheGetJSON_(cacheKey);
    if (cached) return cached;
  }

  const selectClause = encodeURIComponent(SUPABASE_USER_SELECT_FIELDS.join(','));
  const query = '?select=' + selectClause + '&username=eq.' + encodeURIComponent(normalized) + '&limit=1';
  const data = supabaseRequest_(SUPABASE_APP_USERS_ENDPOINT + query);
  if (Array.isArray(data) && data.length > 0) {
    const record = data[0];
    cacheSupabaseUserRecord_(normalized, record);
    return record;
  }
  if (cacheKey) {
    safeScriptCacheRemove_(cacheKey);
  }
  if (missKey) {
    safeScriptCachePutJSON_(missKey, 1, SUPABASE_USER_MISS_CACHE_TTL_SECONDS);
  }
  return null;
}

function updateSupabaseUserByUsername_(username, payload, options) {
  const normalized = String(username == null ? '' : username).trim();
  if (!normalized) throw new Error('Username is required');
  const opts = options || {};
  const query = SUPABASE_APP_USERS_ENDPOINT + '?username=eq.' + encodeURIComponent(normalized);
  removeSupabaseUserCache_(normalized);
  const requestOptions = {
    method: 'PATCH',
    payload: payload,
    headers: Object.assign(
      { Prefer: opts.returnMinimal ? 'return=minimal' : 'return=representation' },
      opts.headers || {}
    )
  };
  const result = supabaseRequest_(query, requestOptions);
  if (opts.returnMinimal) {
    if (opts.cacheRecord && typeof opts.cacheRecord === 'object') {
      cacheSupabaseUserRecord_(normalized, opts.cacheRecord);
      return opts.cacheRecord;
    }
    return null;
  }
  if (Array.isArray(result) && result.length > 0) {
    const record = result[0];
    cacheSupabaseUserRecord_(normalized, record);
    return record;
  }
  return null;
}

function insertSupabaseUser_(payload) {
  const result = supabaseRequest_(SUPABASE_APP_USERS_ENDPOINT, {
    method: 'POST',
    payload: payload,
    headers: { Prefer: 'return=representation' }
  });
  if (Array.isArray(result) && result.length > 0 && result[0].username) {
    cacheSupabaseUserRecord_(result[0].username, result[0]);
  }
  return result;  
}

function deleteSupabaseUserByUsername_(username) {
  const normalized = String(username == null ? '' : username).trim();
  if (!normalized) throw new Error('Username is required');
  const query = SUPABASE_APP_USERS_ENDPOINT + '?username=eq.' + encodeURIComponent(normalized);
  removeSupabaseUserCache_(normalized);
  supabaseRequest_(query, { method: 'DELETE' });
}

function clearUserSession_(username, token, options) {
  const opts = options || {};
  if (token) {
    try {
      removeSessionFromCache_(token);
    } catch (cacheError) {
      Logger.log('clearUserSession_ cache error: ' + cacheError);
    }
  }
  if (!username || opts.skipSupabaseUpdate) return;
  try {
    updateSupabaseUserByUsername_(username, {
      active_session_token: null,
      session_token_expiry: null
    }, { returnMinimal: true });
  } catch (e) {
    Logger.log('clearUserSession_ error: ' + e);
  }  
}

function clearUserSessionByToken_(token) {
  if (!token) return;
  try {
    const query = SUPABASE_APP_USERS_ENDPOINT + '?active_session_token=eq.' + encodeURIComponent(token);
    const result = supabaseRequest_(query, {
      method: 'PATCH',
      payload: {
        active_session_token: null,
        session_token_expiry: null
      },
      headers: { Prefer: 'return=representation' }
    });
    if (Array.isArray(result)) {
      result.forEach(function (row) {
        if (row && row.username) {
          removeSupabaseUserCache_(row.username);
          cacheSupabaseUserRecord_(row.username, row);
        }
      });
    }    
  } catch (e) {
    Logger.log('clearUserSessionByToken_ error: ' + e);
  }
  try {
    removeSessionFromCache_(token);
  } catch (cacheError) {
    Logger.log('clearUserSessionByToken_ cache error: ' + cacheError);
  }
}

function parseSupabaseTimestamp_(value) {
  if (!value) return null;
  if (value instanceof Date) return value;
  const dt = new Date(value);
  if (isNaN(dt.getTime())) return null;
  return dt;
}

function formatSupabaseDateTime_(value) {
  const dt = parseSupabaseTimestamp_(value);
  if (!dt) return '';
  try {
    return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  } catch (e) {
    return dt.toISOString();
  }
}

// =============== DATE/TIME NORMALIZATION HELPERS ===============
function stripLeadingApostrophe(v) {
  if (typeof v === 'string' && v.length > 0 && v[0] === "'") return v.slice(1);
  return v;
}

function normalizeDate(v) {
  if (!v) return null;
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  v = stripLeadingApostrophe(v);
  var iso = String(v).match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (iso) {
    var y = parseInt(iso[1], 10);
    var mIso = parseInt(iso[2], 10) - 1;
    var dIso = parseInt(iso[3], 10);
    var dtIso = new Date(y, mIso, dIso);
    if (!isNaN(dtIso.getTime()) && dtIso.getFullYear() === y && dtIso.getMonth() === mIso && dtIso.getDate() === dIso) {
      return new Date(dtIso.getFullYear(), dtIso.getMonth(), dtIso.getDate());
    }
  }
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
  const parsed = parseSupabaseTimestamp_(stripLeadingApostrophe(v));
  if (parsed) return Utilities.formatDate(parsed, "Asia/Ho_Chi_Minh", "dd/MM/yyyy");  
  return stripLeadingApostrophe(v);
}

function formatTimeForClient(v) {
  if (!v && v!==0) return '';
  if (v instanceof Date) return Utilities.formatDate(v, "Asia/Ho_Chi_Minh", "HH:mm:ss");
  const stripped = stripLeadingApostrophe(v);
  const parsed = parseSupabaseTimestamp_(stripped);
  if (parsed) return Utilities.formatDate(parsed, "Asia/Ho_Chi_Minh", "HH:mm:ss");
  if (typeof v === 'number') {
    var total = Math.round(v*86400);
    var hh = Math.floor(total/3600);
    var mm = Math.floor((total%3600)/60);
    var ss = total%60;
    return String(hh).padStart(2,'0')+':'+String(mm).padStart(2,'0')+':'+String(ss).padStart(2,'0');
  }
  const match = String(stripped || '').match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (match) {
    const hh = String(Math.min(23, Math.max(0, parseInt(match[1], 10) || 0))).padStart(2, '0');
    const mm = String(Math.min(59, Math.max(0, parseInt(match[2], 10) || 0))).padStart(2, '0');
    const ss = String(Math.min(59, Math.max(0, match[3] ? parseInt(match[3], 10) : 0))).padStart(2, '0');
    return hh + ':' + mm + ':' + ss;
  }
  return stripped;
}

function toDisplayString_(value) {
  if (value === null || value === undefined) return '';
  return String(value);
}

function parseHistoryDateFilter_(value, endOfDay) {
  if (!value) return null;
  const match = String(value).trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;
  const year = parseInt(match[1], 10);
  const month = parseInt(match[2], 10);
  const day = parseInt(match[3], 10);
  const date = createUtcDate_(year, month, day);
  if (!date) return null;
  const tz = 'Asia/Ho_Chi_Minh';
  const timePart = endOfDay ? "23:59:59" : "00:00:00";
  try {
    return Utilities.formatDate(date, tz, "yyyy-MM-dd'T'" + timePart + "XXX");
  } catch (e) {
    return null;
  }
}

function parseContentRangeTotal_(headers) {
  if (!headers || typeof headers !== 'object') return null;
  const rangeHeader = headers['Content-Range'] || headers['content-range'];
  if (!rangeHeader) return null;
  const parts = String(rangeHeader).split('/');
  if (parts.length !== 2) return null;
  const total = parseInt(parts[1], 10);
  return isNaN(total) ? null : total;
}

function sanitizeHistorySearchTerm_(value) {
  if (!value) return '';
  return String(value).replace(/["'`]/g, '').replace(/%/g, '').replace(/,/g, ' ').trim();
}


function normalizeToUtcDate_(date) {
  if (Object.prototype.toString.call(date) !== '[object Date]' || isNaN(date)) return null;
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
}

function toUtcDateKey_(date) {
  var normalized = normalizeToUtcDate_(date);
  if (!normalized) return '';
  var y = normalized.getUTCFullYear();
  var m = ('0' + (normalized.getUTCMonth() + 1)).slice(-2);
  var d = ('0' + normalized.getUTCDate()).slice(-2);
  return y + '-' + m + '-' + d;
}

function createUtcDate_(year, month, day) {
  if (typeof year !== 'number' || !isFinite(year) || typeof month !== 'number' || !isFinite(month) || typeof day !== 'number' || !isFinite(day)) return null;
  if (year < 1900 || month < 1 || month > 12 || day < 1 || day > 31) return null;
  var date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) {
    return null;
  }
  return date;
}

function collectExcelDateCandidates_(value) {
  var map = {};

  function addCandidate(date, format) {
    var normalized = normalizeToUtcDate_(date);
    if (!normalized) return;
    var key = toUtcDateKey_(normalized);
    if (!key || map[key]) return;
    map[key] = { date: normalized, format: format || 'unknown', key: key };
  }

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    addCandidate(value, 'dateObject');
  } else if (typeof value === 'number' && isFinite(value)) {
    var millis = Math.round((value - 25569) * 86400 * 1000);
    addCandidate(new Date(millis), 'serial');
  }

  if (value != null) {
    var str = String(value).trim();
    if (str) {
      var sanitized = str.replace(/^'+/, '').replace(/"/g, '');
      var digitParts = sanitized.match(/(\d{1,4})/g);

      if (digitParts && digitParts.length === 3) {
        var numbers = digitParts.map(function(part) { return parseInt(part, 10); });
        var lengths = digitParts.map(function(part) { return part.length; });

        function addFromParts(year, month, day, format) {
          var candidate = createUtcDate_(year, month, day);
          if (candidate) addCandidate(candidate, format);
        }

        if (lengths[0] === 4) {
          addFromParts(numbers[0], numbers[1], numbers[2], 'YMD');
        } else if (lengths[2] === 4) {
          addFromParts(numbers[2], numbers[1], numbers[0], 'DMY');
          addFromParts(numbers[2], numbers[0], numbers[1], 'MDY');
        } else if (lengths[1] === 4) {
          addFromParts(numbers[1], numbers[0], numbers[2], 'DMY');
          addFromParts(numbers[1], numbers[2], numbers[0], 'MDY');
        } else {
          addFromParts(numbers[2], numbers[1], numbers[0], 'DMY');
          addFromParts(numbers[2], numbers[0], numbers[1], 'MDY');
        }
      }

      var fallback = new Date(sanitized);
      if (!isNaN(fallback)) {
        addCandidate(fallback, 'native');
      }
    }
  }

  return Object.keys(map).map(function(key) { return map[key]; });
}

function parseExcelDate_(v) {
  if (v == null || v === '') return '';
  var candidates = collectExcelDateCandidates_(v);
  if (!candidates.length) return '';

  var chosen = candidates[0];
  for (var i = 0; i < candidates.length; i++) {
    if (candidates[i].format === 'DMY') {
      chosen = candidates[i];
      break;
    }
  }

  var date = chosen.date;
  return new Date(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
}

function parseExcelTime_(v) {
  if (v == null || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return Utilities.formatDate(v, "Asia/Ho_Chi_Minh", "HH:mm:ss");
    var hh=v.getHours(), mm=v.getMinutes(), ss=v.getSeconds();
    return (hh*3600+mm*60+ss)/86400;    
  }
  if (typeof v === 'number') {
    var frac = v % 1;
    if (frac < 0) frac = (frac + 1) % 1;
    return frac;
  }
  var m = String(v).match(/(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?/);
  if (m) {
    var hh = Math.min(23, Math.max(0, parseInt(m[1],10) || 0));
    var mm = Math.min(59, Math.max(0, parseInt(m[2],10) || 0));
    var ss = Math.min(59, Math.max(0, m[3]?parseInt(m[3],10):0));
    return (hh*3600+mm*60+ss)/86400;
  }
  return '';
}

function sanitizeXpplText_(value) {
  if (value == null || value === '') return '';
  if (typeof value === 'number') return String(value);
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm:ss');
  }
  var str = String(value);
  if (!str) return '';
  str = str.replace(/\r?\n/g, ' ').trim();
  if (str.charAt(0) === '=') {
    str = str.replace(/^=+/, '');
  }
  return str;
}

function normalizeXpplDbValue_(header, value) {
  var type = XPPL_DB_COLUMN_TYPES[header] || 'text';
  if (type === 'date') {
    var d = parseExcelDate_(value);
    return d || '';
  }
  if (type === 'time') {
    var t = parseExcelTime_(value);
    return t === '' ? '' : t;
  }
  return sanitizeXpplText_(value);
}

function applyXpplDbFormats_(sheet, startRow, numRows) {
  if (numRows <= 0) return;
  for (var i = 0; i < XPPL_DB_HEADERS.length; i++) {
    var header = XPPL_DB_HEADERS[i];
    var type = XPPL_DB_COLUMN_TYPES[header] || 'text';
    var format;
    if (type === 'date') format = 'dd/MM/yyyy';
    else if (type === 'time') format = 'HH:mm:ss';
    else format = '@';
    sheet.getRange(startRow, i + 1, numRows, 1).setNumberFormat(format);
  }
}

function toXpplSupabaseDate_(value) {
  if (value === '' || value === null || value === undefined) return '';
  const normalized = normalizeDate(value);
  if (normalized) {
    return Utilities.formatDate(normalized, 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd');
  }
  if (typeof value === 'number' && isFinite(value)) {
    const excelDate = parseExcelDate_(value);
    if (excelDate) {
      return Utilities.formatDate(excelDate, 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd');
    }
  }
  const parsed = parseSupabaseTimestamp_(stripLeadingApostrophe(value));
  if (parsed) {
    return Utilities.formatDate(parsed, 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd');
  }
  const str = sanitizeXpplText_(value);
  if (!str) return '';
  const dmyMatch = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (dmyMatch) {
    const day = parseInt(dmyMatch[1], 10);
    const month = parseInt(dmyMatch[2], 10) - 1;
    const year = parseInt(dmyMatch[3], 10);
    const dateObj = new Date(year, month, day);
    if (!isNaN(dateObj)) {
      return Utilities.formatDate(dateObj, 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd');
    }
  }
  const isoMatch = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) {
    return isoMatch.slice(1, 4).join('-');
  }
  return str;
}

function toXpplSupabaseTime_(value) {
  if (value === '' || value === null || value === undefined) return '';
  if (value instanceof Date && !isNaN(value)) {
    return Utilities.formatDate(value, 'Asia/Ho_Chi_Minh', 'HH:mm:ss');
  }
  let fraction = null;
  if (typeof value === 'number' && isFinite(value)) {
    fraction = value % 1;
    if (fraction < 0) fraction += 1;
  } else {
    fraction = normalizeTime(value);
  }
  if (fraction !== null && fraction !== undefined && !isNaN(fraction)) {
    const totalSeconds = Math.round(fraction * 86400);
    const hh = Math.floor(totalSeconds / 3600);
    const mm = Math.floor((totalSeconds % 3600) / 60);
    const ss = totalSeconds % 60;
    return String(hh).padStart(2, '0') + ':' + String(mm).padStart(2, '0') + ':' + String(ss).padStart(2, '0');
  }
  const parsed = parseSupabaseTimestamp_(stripLeadingApostrophe(value));
  if (parsed) {
    return Utilities.formatDate(parsed, 'Asia/Ho_Chi_Minh', 'HH:mm:ss');
  }
  const str = sanitizeXpplText_(value);
  const match = str.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (match) {
    const hh = String(Math.min(23, Math.max(0, parseInt(match[1], 10) || 0))).padStart(2, '0');
    const mm = String(Math.min(59, Math.max(0, parseInt(match[2], 10) || 0))).padStart(2, '0');
    const ss = String(Math.min(59, Math.max(0, match[3] ? parseInt(match[3], 10) : 0))).padStart(2, '0');
    return hh + ':' + mm + ':' + ss;
  }
  return str;
}

function buildXpplDatabasePayload_(row, options) {
  const opts = options || {};
  const userLabel = sanitizeXpplText_(opts.userLabel || '');
  const payload = {};
  Object.keys(XPPL_DB_FIELD_MAP).forEach(function (header) {
    const column = XPPL_DB_FIELD_MAP[header];
    if (!column) return;
    if (header === 'ID' || header === 'Changed Date' || header === 'Changed Time') return;
    if (header === 'CreateUser') {
      payload[column] = userLabel;
      return;
    }
    const value = row ? row[header] : '';
    if (header === 'Date In' || header === 'Date Out') {
      const iso = toXpplSupabaseDate_(value);
      payload[column] = iso || '';
      return;
    }
    if (header === 'Time In' || header === 'Time Out') {
      payload[column] = toXpplSupabaseTime_(value);
      return;
    }
    if (value === '' || value === null || value === undefined) {
      payload[column] = '';
      return;
    }
    payload[column] = sanitizeXpplText_(value);
  });
  return payload;
}

function mapXpplRecordToRowArray_(record, headers) {
  const cols = headers || XPPL_DB_HEADERS;
  return cols.map(function (header) {
    const column = XPPL_DB_FIELD_MAP[header];
    if (!column) return '';
    const value = record && record.hasOwnProperty(column) ? record[column] : '';
    if (value === null || value === undefined) return '';
    if (header === 'Changed Date' || header === 'Changed Time') {
      return parseSupabaseTimestamp_(value) || value;
    }
    return value;
  });
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
    if (key === 'Updated At') {
      const dt = parseSupabaseTimestamp_(val);
      if (dt) {
        out[key] = Utilities.formatDate(dt, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm:ss');
      } else {
        out[key] = stripLeadingApostrophe(val);
      }
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
  'Transportation Company', 'Subcontractor', 'Vehicle Status', 'Registration Status', 'Time'
];
const NUMERIC_REGISTER_FIELDS = ['Wheel', 'Truck weight', 'Pay load'];
const HEADERS_TOTAL_LIST = [
  'ID', 'Truck Plate', 'Country', 'Wheel', 'Trailer Plate', 'Truck weight',
  'Pay load', 'Container No1', 'Container No2', 'Driver Name', 'ID/Passport',
  'Phone number', 'Transportation Company', 'Subcontractor', 'Vehicle Status',
  'Activity Status', 'Register Date', 'Time', 'Created By', 'Updated At', 'Updated By'
];

const TRUCK_LIST_TOTAL_SELECT_FIELDS = [
  'id',
  'truck_plate',
  'country',
  'wheel',
  'trailer_plate',
  'truck_weight',
  'pay_load',
  'container_no1',
  'container_no2',
  'driver_name',
  'id_passport',
  'phone_number',
  'transportation_company',
  'subcontractor',
  'vehicle_status',
  'activity_status',
  'register_date',
  'time',
  'created_by',
  'updated_at',
  'updated_by'
];

const TRUCK_LIST_TOTAL_COLUMN_MAP = {
  'ID': 'id',
  'Truck Plate': 'truck_plate',
  'Country': 'country',
  'Wheel': 'wheel',
  'Trailer Plate': 'trailer_plate',
  'Truck weight': 'truck_weight',
  'Pay load': 'pay_load',
  'Container No1': 'container_no1',
  'Container No2': 'container_no2',
  'Driver Name': 'driver_name',
  'ID/Passport': 'id_passport',
  'Phone number': 'phone_number',
  'Transportation Company': 'transportation_company',
  'Subcontractor': 'subcontractor',
  'Vehicle Status': 'vehicle_status',
  'Activity Status': 'activity_status',
  'Register Date': 'register_date',
  'Time': 'time',
  'Created By': 'created_by',
  'Updated At': 'updated_at',
  'Updated By': 'updated_by'
};

function coerceNumericRegisterFields_(record) {
  if (!record) return;
  NUMERIC_REGISTER_FIELDS.forEach(field => {
    if (!(field in record)) return;
    const value = record[field];
    if (value === '' || value === null || value === undefined) return;
    const parsed = typeof value === 'number' ? value : parseFloat(String(value).replace(',', '.'));
    if (!isNaN(parsed)) {
      record[field] = parsed;
    }
  });
}


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

function buildLoginHistoryPayload_(username, outcome, context) {
  const payload = {
    occurred_at: new Date().toISOString()
  };

  const normalizedUsername = String(username == null ? '' : username).trim();
  payload.username = normalizedUsername || null;

  const normalizedOutcome = String(outcome == null ? '' : outcome).trim();
  payload.outcome = normalizedOutcome || null;

  const ctx = context && typeof context === 'object' ? context : {};

  const ip = ctx.ip == null ? '' : String(ctx.ip).trim();
  if (ip) {
    payload.ip = ip;
  }

  const latitude = Number(ctx.latitude);
  if (Number.isFinite(latitude)) {
    payload.latitude = latitude;
  }

  const longitude = Number(ctx.longitude);
  if (Number.isFinite(longitude)) {
    payload.longitude = longitude;
  }

  const accuracySource = ctx.accuracy_m != null ? ctx.accuracy_m : ctx.accuracy;
  const accuracy = Number(accuracySource);
  if (Number.isFinite(accuracy)) {
    payload.accuracy_m = Math.round(accuracy);
  }

  const userAgent = ctx.userAgent != null ? ctx.userAgent : ctx.user_agent;
  if (userAgent != null && userAgent !== '') {
    const agentString = String(userAgent);
    payload.user_agent = agentString.length > 1024 ? agentString.slice(0, 1024) : agentString;
  }

  return payload;
}

function logLoginAttempt(username, outcome, context) {
  try {
    const payload = buildLoginHistoryPayload_(username, outcome, context);
    supabaseRequest_(SUPABASE_AUTH_LOGIN_HISTORY_ENDPOINT, {
      method: 'POST',
      payload: payload,
      headers: { Prefer: 'return=minimal' }
    });
  } catch (e) {
    Logger.log('Không thể ghi lịch sử đăng nhập (Supabase): ' + e);
  }
}

function ensureSupervisionAccount_() {
  try {
    const existing = getSupabaseUserByUsername_(SUPERVISION_DEFAULT_USERNAME);
    if (existing) return;

    const payload = {
      username: SUPERVISION_DEFAULT_USERNAME,
      password_hash: SUPERVISION_DEFAULT_PASSWORD,
      role: SUPERVISION_DEFAULT_ROLE,
      contractor: '',
      password_last_updated: new Date().toISOString(),
      security_code: Math.random().toString(36).slice(-6).toUpperCase(),
      customer_name: ''
    };

    insertSupabaseUser_(payload);
  } catch (err) {
    Logger.log('ensureSupervisionAccount_ error: ' + err);
  }
}

function checkLogin(credentials) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const username = String(credentials.username == null ? '' : credentials.username).trim();

  const loginContext = credentials && typeof credentials.context === 'object' ? credentials.context : {};

  try {
    const lockoutUntil = scriptProperties.getProperty(`lockout_until_${username}`);
    if (lockoutUntil && new Date().getTime() < parseFloat(lockoutUntil)) {
      const timeLeft = Math.ceil((parseFloat(lockoutUntil) - new Date().getTime()) / (60 * 1000));
      logLoginAttempt(username, 'Locked', loginContext);      
      throw new Error(`Tài khoản của bạn đã bị tạm khóa. Vui lòng thử lại sau ${timeLeft} phút.`);
    }

    const userRecord = getSupabaseUserByUsername_(username);

    if (!userRecord || String(userRecord.password_hash == null ? '' : userRecord.password_hash) !== String(credentials.password == null ? '' : credentials.password)) {
      logLoginAttempt(username, 'Failure', loginContext);
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

    const nowMs = Date.now();
    const activeToken = String(userRecord.active_session_token == null ? '' : userRecord.active_session_token).trim();
    const tokenExpiry = parseSupabaseTimestamp_(userRecord.session_token_expiry);

    if (activeToken) {
      if (tokenExpiry && nowMs < tokenExpiry.getTime()) {
        logLoginAttempt(username, 'Rejected-ActiveSession', loginContext);        
        throw new Error('Tài khoản này đã được đăng nhập trên một thiết bị khác.');
      }
      clearUserSession_(username, activeToken, { skipSupabaseUpdate: true });
    }

    const customerName = String(userRecord.customer_name == null ? '' : userRecord.customer_name).trim();

    scriptProperties.deleteProperty(`failed_attempts_${username}`);
    scriptProperties.deleteProperty(`lockout_level_${username}`);
    scriptProperties.deleteProperty(`lockout_until_${username}`);    

    const newSessionToken = Utilities.getUuid();
    const tokenExpiryDate = new Date(nowMs + SESSION_TIMEOUT_SECONDS * 1000);

    const updatePayload = {
      active_session_token: newSessionToken,
      session_token_expiry: tokenExpiryDate.toISOString()
    };

    const cachedRecord = Object.assign({}, userRecord, updatePayload);
    try {
      updateSupabaseUserByUsername_(username, updatePayload, {
        returnMinimal: true,
        cacheRecord: cachedRecord
      });
    } catch (updateError) {
      clearUserSession_(username, activeToken);
      logLoginAttempt(username, 'Failure-UpdateSession', loginContext);      
      throw updateError;
    }

    const userRole = userRecord.role;
    const canonicalUsername = String(userRecord.username == null ? '' : userRecord.username).trim();
    const normalizedRole = String(userRole == null ? '' : userRole).trim().toLowerCase();
    const userContractor = String(userRecord.contractor == null ? '' : userRecord.contractor);    

    const userSession = {
      isLoggedIn: true,
      username: canonicalUsername,
      role: normalizedRole,
      roleDisplay: userRole || normalizedRole,
      contractor: userContractor,
      customerName: customerName,
      token: newSessionToken
    };

    logLoginAttempt(username, 'Success', loginContext);
    cacheSession_(userSession);

    return userSession;
  } catch (e) {
    Logger.log(e);
    throw new Error(e.message);
  }
}

function logout(sessionToken) {
  let token = String(sessionToken == null ? '' : sessionToken).trim();
  let session = null;

  if (token) {
    session = getSessionFromCache_(token);
    if (!session) {
      session = lookupSessionFromSheet(token);
    }
  } else {
    const legacy = safeGetUserCacheJSON('user_session');
    if (legacy && legacy.token) {
      session = legacy;
      token = legacy.token;
    }
  }

  try {
    if (session && session.username) {
      clearUserSession_(session.username, token);
    } else if (token) {
      clearUserSessionByToken_(token);
    }
  } catch (e) {
    Logger.log('logout error: ' + e);
  }

  if (token) {
    try {
      removeSessionFromCache_(token);
    } catch (cacheError) {
      Logger.log('logout cache remove error: ' + cacheError);
    }
  }
  safeRemoveUserCacheKey('user_session');
  return { success: true };
}

function changePassword(passwords, sessionToken) {
  const session = validateSession(sessionToken);
  const { currentPassword, newPassword } = passwords;

  try {
    const userRecord = getSupabaseUserByUsername_(session.username);
    if (!userRecord) throw new Error('Không tìm thấy người dùng.');
    if (String(userRecord.password_hash == null ? '' : userRecord.password_hash) !== String(currentPassword == null ? '' : currentPassword)) {
      throw new Error('Mật khẩu hiện tại không đúng.');
    }

    const updatedRecord = Object.assign({}, userRecord, {
      password_hash: newPassword,
      password_last_updated: new Date().toISOString()
    });
    updateSupabaseUserByUsername_(session.username, {
      password_hash: newPassword,
      password_last_updated: updatedRecord.password_last_updated
    }, { returnMinimal: true, cacheRecord: updatedRecord });

    return 'Đổi mật khẩu thành công!';
  } catch (e) { Logger.log(e); throw new Error('Lỗi khi đổi mật khẩu: ' + e.message); }
}

function resetPassword(data) {
  const { username, securityCode, newPassword } = data;
  try {
    const userRecord = getSupabaseUserByUsername_(username);
    if (!userRecord) throw new Error('Tên đăng nhập không tồn tại.');
    const storedCode = String(userRecord.security_code == null ? '' : userRecord.security_code).trim();
    if (storedCode.toUpperCase() !== String(securityCode == null ? '' : securityCode).trim().toUpperCase()) {
      throw new Error('Mã bảo mật không chính xác.');
    }

    const updatedRecord = Object.assign({}, userRecord, {
      password_hash: newPassword,
      password_last_updated: new Date().toISOString()
    });
    updateSupabaseUserByUsername_(username, {
      password_hash: newPassword,
      password_last_updated: updatedRecord.password_last_updated
    }, { returnMinimal: true, cacheRecord: updatedRecord });    
    clearUserSession_(username, String(userRecord.active_session_token || ''));

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
    const data = supabaseRequest_(
      SUPABASE_APP_USERS_ENDPOINT + '?select=username,password_hash,role,contractor,password_last_updated,security_code,customer_name&order=username.asc'
    );
    if (!Array.isArray(data)) return [];

    return data.map(row => ({
      Username: row.username || '',
      Password: row.password_hash || '',
      Role: row.role || '',
      Contractor: row.contractor || '',
      PasswordLastUpdated: formatSupabaseDateTime_(row.password_last_updated),
      SecurityCode: row.security_code || '',
      CustomerName: row.customer_name || ''
    }));
  } catch (e) { Logger.log(e); throw new Error('Không thể lấy danh sách người dùng.'); }
}

function updateUser(userData, sessionToken) {
  const session = validateSession(sessionToken);
  if (session.role !== 'admin') throw new Error('Bạn không có quyền truy cập chức năng này.');

  try {
    const updated = updateSupabaseUserByUsername_(userData.Username, {
      role: userData.Role,
      contractor: userData.Contractor,
      customer_name: userData.CustomerName || ''
    });
    if (!updated) throw new Error('Không tìm thấy người dùng.');

    return 'Cập nhật người dùng thành công!';
  } catch (e) { Logger.log(e); throw new Error('Lỗi khi cập nhật người dùng: ' + e.message); }
}

function adminResetPassword(username, sessionToken) {
  const session = validateSession(sessionToken);
  if (session.role !== 'admin') throw new Error('Bạn không có quyền truy cập chức năng này.');

  try {
    const userRecord = getSupabaseUserByUsername_(username);
    if (!userRecord) throw new Error('Không tìm thấy người dùng.');

    const newPassword = Math.random().toString(36).slice(-8);
    const updatedRecord = Object.assign({}, userRecord, {
      password_hash: newPassword,
      password_last_updated: new Date().toISOString()
    });
    updateSupabaseUserByUsername_(username, {
      password_hash: newPassword,
      password_last_updated: updatedRecord.password_last_updated
    }, { returnMinimal: true, cacheRecord: updatedRecord });    
    clearUserSession_(username, String(userRecord.active_session_token || ''));

    return `Mật khẩu mới cho ${username} là: ${newPassword}`;
  } catch (e) { Logger.log(e); throw new Error('Lỗi khi đặt lại mật khẩu: ' + e.message); }
}

function addNewUser(newUserData, sessionToken) {
  const session = validateSession(sessionToken);
  if (session.role !== 'admin') throw new Error('Bạn không có quyền truy cập chức năng này.');

  try {
    const existing = getSupabaseUserByUsername_(newUserData.Username);
    if (existing) throw new Error('Tên đăng nhập đã tồn tại.');

    const newPassword = Math.random().toString(36).slice(-8);
    const newSecurityCode = Math.random().toString(36).slice(-6).toUpperCase();

    insertSupabaseUser_({
      username: newUserData.Username,
      password_hash: newPassword,
      role: newUserData.Role,
      contractor: newUserData.Contractor,
      customer_name: newUserData.CustomerName || '',
      password_last_updated: new Date().toISOString(),
      security_code: newSecurityCode,
      active_session_token: null,
      session_token_expiry: null
    });

    return `Đã tạo người dùng ${newUserData.Username} thành công.
Mật khẩu: ${newPassword}
Mã bảo mật: ${newSecurityCode}`;
  } catch (e) { Logger.log(e); throw new Error('Lỗi khi tạo người dùng mới: ' + e.message); }
}

function deleteUser(username, sessionToken) {
  const session = validateSession(sessionToken);
  if (session.role !== 'admin') throw new Error('Bạn không có quyền truy cập chức năng này.');
  if (session.username === username) throw new Error('Bạn không thể tự xóa tài khoản của mình.');

  try {
    const userRecord = getSupabaseUserByUsername_(username);
    if (!userRecord) throw new Error('Không tìm thấy người dùng.');

    deleteSupabaseUserByUsername_(username);
    clearUserSession_(username, String(userRecord.active_session_token || ''));
    return `Đã xóa người dùng ${username} thành công!`;
  } catch (e) { Logger.log(e); throw new Error('Lỗi khi xóa người dùng: ' + e.message); }
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

function getVietnamCurrentTime() {
  const timezone = 'Asia/Ho_Chi_Minh';
  const now = new Date();
  return {
    epochMillis: now.getTime(),
    isoDate: Utilities.formatDate(now, timezone, 'yyyy-MM-dd'),
    dateString: Utilities.formatDate(now, timezone, 'dd/MM/yyyy'),
    timeString: Utilities.formatDate(now, timezone, 'HH:mm:ss')
  };
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

  const sanitized = contracts
    .map(function (value) {
      return String(value == null ? '' : value).replace(/^'+/, '').trim();
    })
    .filter(function (value) { return value.length > 0; });

  if (!sanitized.length) return {};

  const filters = [];
  const inFilter = buildSupabaseInFilter_('contract_no', sanitized);
  if (inFilter) filters.push(inFilter);

  const rows = fetchContractDataRows_(['contract_no', 'customer_name', 'status'], filters);
  if (!rows.length) return {};

  const allowed = new Set(sanitized);
  const grouped = {};
  rows.forEach(function (row) {
    const no = String(row.contract_no == null ? '' : row.contract_no).replace(/^'+/, '').trim();
    if (!no || !allowed.has(no)) return;
    const cus = String(row.customer_name == null ? '' : row.customer_name).replace(/^'+/, '').trim();
    if (!cus) return;
    if (!grouped[no]) grouped[no] = new Set();
    grouped[no].add(cus);
  });

  const result = {};
  Object.keys(grouped).forEach(function (contract) {
    result[contract] = Array.from(grouped[contract]).sort();
  });
  return result;
}

function buildEmptyResult_(draw, includeSummary) {
  const result = {
    draw: parseInt(draw, 10),
    recordsTotal: 0,
    recordsFiltered: 0,
    data: []
  };
  if (includeSummary) {
    result.summary = { total: 0, pending: 0, approved: 0 };
  }
  return result;
}

function fetchRowsByIndices_(sheet, rowIndices, columnCount) {
  if (!Array.isArray(rowIndices) || !rowIndices.length) return [];

  const sorted = rowIndices.slice().sort((a, b) => a - b);
  const collected = [];

  const pushBlock = (startIdx, length) => {
    if (length <= 0) return;
    const startRow = startIdx + 2; // +2 vì dữ liệu bắt đầu từ dòng 2
    const values = sheet.getRange(startRow, 1, length, columnCount).getValues();
    for (var i = 0; i < values.length; i++) {
      collected.push(values[i]);
    }
  };

  var blockStart = sorted[0];
  var blockLength = 1;
  for (var i = 1; i < sorted.length; i++) {
    if (sorted[i] === sorted[i - 1] + 1) {
      blockLength++;
    } else {
      pushBlock(blockStart, blockLength);
      blockStart = sorted[i];
      blockLength = 1;
    }
  }
  pushBlock(blockStart, blockLength);

  return collected;
}

function processServerSide(params, sheetName, headers, defaultSortColumnIndex) {
  params = params || {};
  const userSession = validateSession(params.sessionToken);
  const userRole = String(userSession.role || '').toLowerCase();
  const includeSummary = sheetName === VEHICLE_REGISTRATION_CACHE_KEY;

  const cacheKey = buildServerSideCacheKey_(sheetName, params, userRole);
  const cachedResult = cacheKey ? safeScriptCacheGetJSON_(cacheKey) : null;
  if (cachedResult) {
    return cachedResult;
  }

  const respondWithCache = function (result) {
    if (cacheKey && result) {
      safeScriptCachePutJSON_(cacheKey, result, SERVER_SIDE_CACHE_TTL_SECONDS);
    }
    return result;
  };

  if (sheetName === VEHICLE_REGISTRATION_CACHE_KEY) {
    const result = processVehicleRegistrationsServerSide_(params, headers, userSession, defaultSortColumnIndex, includeSummary);
    return respondWithCache(result);
  }  

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return respondWithCache(buildEmptyResult_(params.draw, includeSummary));
  }

  const totalRows = lastRow - 1;
  const columnCount = headers.length;

  const idxRegisterDate = headers.indexOf('Register Date');
  const idxContract = headers.indexOf('Contract No');
  const idxCompany = headers.indexOf('Transportation Company');
  const idxActivity = headers.indexOf('Activity Status');
  const idxStatus = headers.indexOf('Registration Status');

  if (params.dateString && idxRegisterDate === -1) {
    return respondWithCache(buildEmptyResult_(params.draw, includeSummary));
  }

  const columnCache = {};
  function getColumnValues(idx) {
    if (idx === -1) return null;
    if (!(idx in columnCache)) {
      if (totalRows <= 0) {
        columnCache[idx] = [];
      } else {
        const rangeValues = sheet.getRange(2, idx + 1, totalRows, 1).getValues();
        columnCache[idx] = rangeValues.map(function (row) { return row[0]; });
      }
    }
    return columnCache[idx];
  }

  const valuesRegisterDate = (params.dateString && idxRegisterDate !== -1)
    ? getColumnValues(idxRegisterDate)
    : null;
  const valuesContract = (params.contractNo && idxContract !== -1)
    ? getColumnValues(idxContract)
    : null;
  const valuesCompany = (userRole === 'user' && idxCompany !== -1)
    ? getColumnValues(idxCompany)
    : null;
  const valuesActivity = (userRole === 'user' && idxActivity !== -1)
    ? getColumnValues(idxActivity)
    : null;
  const valuesStatus = (userRole === 'user-supervision' && idxStatus !== -1)
    ? getColumnValues(idxStatus)
    : null;

  const dateFilter = params.dateString ? String(params.dateString).trim() : '';
  const contractFilter = params.contractNo
    ? String(params.contractNo).replace(/^'+/, '').trim().toLowerCase()
    : '';
  const contractorValue = String(userSession.contractor == null ? '' : userSession.contractor);

  const timezone = 'Asia/Ho_Chi_Minh';
  const matchedIndices = [];

  for (var r = 0; r < totalRows; r++) {
    if (userRole === 'user') {
      if (valuesCompany) {
        const rawCompany = valuesCompany[r];
        const companyString = String(rawCompany == null ? '' : rawCompany);
        if (companyString !== contractorValue) continue;
      }
      if (valuesActivity) {
        const rawActivity = valuesActivity[r];
        const activityString = String(rawActivity == null ? '' : rawActivity).toUpperCase();
        if (activityString !== 'ACTIVE') continue;
      }
    } else if (userRole === 'user-supervision') {
      if (valuesStatus) {
        const rawStatus = valuesStatus[r];
        const statusString = String(rawStatus == null ? '' : rawStatus).trim().toLowerCase();
        if (statusString !== 'approved') continue;
      }
    }

    if (dateFilter) {
      const cellValue = valuesRegisterDate ? valuesRegisterDate[r] : null;
      if (!cellValue) continue;
      var cmp = '';
      if (cellValue instanceof Date) {
        cmp = Utilities.formatDate(cellValue, timezone, 'dd/MM/yyyy');
      } else {
        cmp = String(cellValue).trim().replace(/^'+/, '');
      }
      if (cmp !== dateFilter) continue;
    }

    if (contractFilter && valuesContract) {
      const rawContract = valuesContract[r];
      const contractString = String(rawContract == null ? '' : rawContract)
        .replace(/^'+/, '')
        .trim()
        .toLowerCase();
      if (contractString !== contractFilter) continue;
    }

    matchedIndices.push(r);
  }

  if (!matchedIndices.length) {
    return respondWithCache(buildEmptyResult_(params.draw, includeSummary));
  }

  let allData = fetchRowsByIndices_(sheet, matchedIndices, columnCount);

  const recordsTotal = allData.length;
  let filteredData = allData;

  if (params.search && params.search.value) {
    const searchValue = params.search.value.toLowerCase();
    filteredData = filteredData.filter(function (row) {
      return row.some(function (cell) {
        return String(cell).toLowerCase().includes(searchValue);
      });
    });
  }

  const recordsFiltered = filteredData.length;

  
  // === SUMMARY (for 'Xe đã đăng ký') ===
  var summary = null;
  try {
    if (sheetName === VEHICLE_REGISTRATION_CACHE_KEY) {
      var statusIdx = headers.indexOf('Registration Status');
      if (statusIdx !== -1) {
        var total = filteredData.length;
        var pending = 0, approved = 0;
        for (var i = 0; i < filteredData.length; i++) {
          var v = filteredData[i][statusIdx];
          v = (v instanceof Date)
            ? Utilities.formatDate(v, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy')
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
    filteredData.sort(function (a, b) {
      const valA = a[columnIndex];
      const valB = b[columnIndex];
      if (valA < valB) return -1 * direction;
      if (valA > valB) return 1 * direction;
      return 0;
    });
  } else if (defaultSortColumnIndex !== -1) {
    filteredData.sort(function (a, b) {
      return (a[defaultSortColumnIndex] < b[defaultSortColumnIndex] ? 1 : -1);
    });
  }  

  const start = Number(params.start || 0);
  const length = Number(params.length || filteredData.length);
  const paginatedData = filteredData.slice(start, start + length);
  const data = paginatedData.map(function (row) { return formatRowForClient_(row, headers); });

  const result = {
    draw: parseInt(params.draw, 10),
    recordsTotal: recordsTotal,
    recordsFiltered: recordsFiltered,
    data: data,
    summary: summary
  };
  return respondWithCache(result);
}

function processVehicleRegistrationsServerSide_(params, headers, userSession, defaultSortColumnIndex, includeSummary) {
  const userRole = String(userSession.role || '').toLowerCase();

  const filters = [];
  const dateFilter = params.dateString ? String(params.dateString).trim() : '';
  if (dateFilter) {
    const iso = toSupabaseDateString_(dateFilter);
    if (iso) filters.push('register_date=eq.' + encodeURIComponent(iso));
    else return buildEmptyResult_(params.draw, includeSummary);
  }

  if (params.contractNo) {
    const contract = String(params.contractNo).replace(/^'+/, '').trim();
    if (contract) filters.push('contract_no=eq.' + encodeURIComponent(contract));
  }

  if (userRole === 'user') {
    const contractor = String(userSession.contractor == null ? '' : userSession.contractor).trim();
    if (!contractor) return buildEmptyResult_(params.draw, includeSummary);
    filters.push('transportation_company=eq.' + encodeURIComponent(contractor));
  } else if (userRole === 'user-supervision') {
    filters.push('registration_status=eq.' + encodeURIComponent('Approved'));
  }

  const queryParams = ['select=' + encodeURIComponent(VEHICLE_REGISTRATION_SELECT_FIELDS.join(',')), 'order=time.desc.nullslast'];
  Array.prototype.push.apply(queryParams, filters);
  const query = SUPABASE_VEHICLE_REG_ENDPOINT + '?' + queryParams.join('&');

  const response = supabaseRequest_(query) || [];
  if (!Array.isArray(response) || !response.length) {
    return buildEmptyResult_(params.draw, includeSummary);
  }

  const allData = response.map(function (row) { return mapVehicleRegistrationRowToArray_(row, headers); });
  let filteredData = allData.slice();

  if (params.search && params.search.value) {
    const searchValue = params.search.value.toLowerCase();
    filteredData = filteredData.filter(function (row) {
      return row.some(function (cell) {
        return String(cell).toLowerCase().includes(searchValue);
      });
    });
  }

  const recordsTotal = allData.length;
  const recordsFiltered = filteredData.length;

  let summary = null;
  if (includeSummary) {
    const statusIdx = headers.indexOf('Registration Status');
    if (statusIdx !== -1) {
      let pending = 0;
      let approved = 0;
      for (var i = 0; i < filteredData.length; i++) {
        const status = String(filteredData[i][statusIdx] || '').trim().toLowerCase();
        if (status === 'pending approval') pending++;
        else if (status === 'approved') approved++;
      }
      summary = { total: filteredData.length, pending: pending, approved: approved };
    }
  }

  if (params.order && params.order.length > 0) {
    const orderInfo = params.order[0];
    const columnIndex = orderInfo.column;
    const direction = orderInfo.dir === 'asc' ? 1 : -1;
    filteredData.sort(function (a, b) {
      const valA = a[columnIndex];
      const valB = b[columnIndex];
      if (valA < valB) return -1 * direction;
      if (valA > valB) return 1 * direction;
      return 0;
    });
  } else if (defaultSortColumnIndex !== -1) {
    filteredData.sort(function (a, b) {
      return (a[defaultSortColumnIndex] < b[defaultSortColumnIndex] ? 1 : -1);
    });
  }

  const paginatedData = filteredData.slice(params.start, params.start + params.length);
  const data = paginatedData.map(function (row) { return formatRowForClient_(row, headers); });

  return {
    draw: parseInt(params.draw, 10),
    recordsTotal: recordsTotal,
    recordsFiltered: recordsFiltered,
    data: data,
    summary: summary
  };
}

function getRegisteredDataServerSide(params) {
  return processServerSide(params, VEHICLE_REGISTRATION_CACHE_KEY, HEADERS_REGISTER, HEADERS_REGISTER.indexOf('Time'));
}

function getRegisteredContractOptions(filter, sessionToken) {
  const session = validateSession(sessionToken);
  const role = String(session.role || '').toLowerCase();
  const dateString = filter && filter.dateString ? String(filter.dateString).trim() : '';
  if (!dateString) return { contracts: [] };
  const iso = toSupabaseDateString_(dateString);
  if (!iso) return { contracts: [] };  

  const queryParts = [
    'select=' + encodeURIComponent(['contract_no', 'transportation_company', 'registration_status'].join(',')),
    'register_date=eq.' + encodeURIComponent(iso)
  ];

  if (role === 'user') {
    const contractor = String(session.contractor == null ? '' : session.contractor).trim();
    if (!contractor) return { contracts: [] };
    queryParts.push('transportation_company=eq.' + encodeURIComponent(contractor));
  } else if (role === 'user-supervision') {
    queryParts.push('registration_status=eq.' + encodeURIComponent('Approved'));
  }

  const rows = supabaseRequest_(SUPABASE_VEHICLE_REG_ENDPOINT + '?' + queryParts.join('&')) || [];
  if (!Array.isArray(rows)) return { contracts: [] };

  const contractsSet = new Set();
  rows.forEach(function (row) {
    const contract = String(row.contract_no == null ? '' : row.contract_no).replace(/^'+/, '').trim();
    if (contract) contractsSet.add(contract);
  });

  return { contracts: Array.from(contractsSet).sort() };
}

function processTruckListTotalServerSide_(params, headers, userSession, defaultSortColumnIndex) {
  params = params || {};
  const userRole = String(userSession.role || '').toLowerCase();
  const start = Number(params.start || 0);
  const length = Number(params.length || 0);

  const queryParts = [
    'select=' + encodeURIComponent(TRUCK_LIST_TOTAL_SELECT_FIELDS.join(',')),
    'order=register_date.desc.nullslast',
    'order=time.desc.nullslast'
  ];

  let rows = supabaseRequest_(SUPABASE_TRUCK_LIST_TOTAL_ENDPOINT + '?' + queryParts.join('&')) || [];
  if (!Array.isArray(rows)) rows = [];

  if (userRole === 'user') {
    const contractor = String(userSession.contractor == null ? '' : userSession.contractor).trim();
    if (!contractor) {
      rows = [];
    } else {
      rows = rows.filter(function (row) {
        const company = String(row.transportation_company == null ? '' : row.transportation_company).trim();
        const activity = String(row.activity_status == null ? '' : row.activity_status).trim().toLowerCase();
        return company === contractor && activity === 'active';
      });
    }
  }

  const allData = rows.map(function (row) { return mapTruckListRowToArray_(row, headers); });
  let filteredData = allData.slice();

  if (params.search && params.search.value) {
    const searchValue = params.search.value.toLowerCase();
    filteredData = filteredData.filter(function (row) {
      return row.some(function (cell) {
        return String(cell).toLowerCase().includes(searchValue);
      });
    });
  }

  const recordsTotal = allData.length;
  const recordsFiltered = filteredData.length;

  if (params.order && params.order.length > 0) {
    const orderInfo = params.order[0];
    const columnIndex = orderInfo.column;
    const direction = orderInfo.dir === 'asc' ? 1 : -1;
    filteredData.sort(function (a, b) {
      const valA = a[columnIndex];
      const valB = b[columnIndex];
      if (valA < valB) return -1 * direction;
      if (valA > valB) return 1 * direction;
      return 0;
    });
  } else if (defaultSortColumnIndex !== -1) {
    filteredData.sort(function (a, b) {
      return (a[defaultSortColumnIndex] < b[defaultSortColumnIndex] ? 1 : -1);
    });
  }

  const paginatedData = filteredData.slice(start, start + (length > 0 ? length : filteredData.length));
  const data = paginatedData.map(function (row) { return formatRowForClient_(row, headers); });

  return {
    draw: parseInt(params.draw, 10),
    recordsTotal: recordsTotal,
    recordsFiltered: recordsFiltered,
    data: data
  };
}

function getTotalListDataServerSide(params) {
  const userSession = validateSession(params && params.sessionToken);
  return processTruckListTotalServerSide_(
    params,
    HEADERS_TOTAL_LIST,
    userSession,
    HEADERS_TOTAL_LIST.indexOf('Register Date')
  );
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

  const dateIn = s(filter && filter.dateString);
  const dateKey = _toDateKey(dateIn);
  if (!dateKey) return { contracts: [], customersByContract: {} };

  const isoDate = toSupabaseDateString_(dateKey);
  if (!isoDate) return { contracts: [], customersByContract: {} };

  const vehicleRows = supabaseRequest_(
    SUPABASE_VEHICLE_REG_ENDPOINT
      + '?select=' + encodeURIComponent(['contract_no', 'register_date', 'registration_status'].join(','))
      + '&register_date=eq.' + encodeURIComponent(isoDate)
      + '&registration_status=eq.' + encodeURIComponent('Approved')
  ) || [];

  if (!Array.isArray(vehicleRows) || !vehicleRows.length) {
    return { contracts: [], customersByContract: {} };
  }

  const setContracts = new Set();
  vehicleRows.forEach(function (row) {
    const cno = s(row.contract_no);
    if (cno) setContracts.add(cno);
 });
  const contracts = Array.from(setContracts).sort();
  if (!contracts.length) return { contracts: [], customersByContract: {} };

  // ---- ContractData: map Contract -> Customers (ưu tiên Status=Active nếu có)
  const customersByContract = {};
  for (const c of contracts) customersByContract[c] = [];

  const filters = [];
  const inFilter = buildSupabaseInFilter_('contract_no', contracts);
  if (inFilter) filters.push(inFilter);

  const contractRows = fetchContractDataRows_(['contract_no', 'customer_name', 'status'], filters);
  contractRows.forEach(function (row) {
    const cno = s(row.contract_no);
    if (!(cno in customersByContract)) return;
    const status = s(row.status).toLowerCase();
    if (status && status !== 'active') return;
    const cus = s(row.customer_name);
    if (!cus) return;
    if (customersByContract[cno].indexOf(cus) === -1) {
      customersByContract[cno].push(cus);
    }
  });
  Object.keys(customersByContract).forEach(function (no) {
    customersByContract[no].sort();
  });

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
  const tz = (Session.getScriptTimeZone && Session.getScriptTimeZone()) || 'Asia/Ho_Chi_Minh';
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

  // 1) Xác thực Contract–Customer trong Supabase contract_data
  const contractFilters = [];
  const inFilter = buildSupabaseInFilter_('contract_no', [contractNo]);
  if (inFilter) contractFilters.push(inFilter);
  const contractRows = fetchContractDataRows_(['contract_no', 'customer_name', 'status'], contractFilters);
  const isValidContract = contractRows.some(function (row) {
    if (s(row.contract_no) !== contractNo) return false;
    if (s(row.customer_name) !== customerName) return false;
    const status = s(row.status).toLowerCase();
    return status === '' || status === 'active';
  });
  if (!isValidContract) {
    return { ok:false, errors:['Customer Name không khớp với Contract No (hoặc hợp đồng không Active).'] };
  }

  // 2) Lọc dữ liệu từ Supabase vehicle_registration
  const isoDate = toSupabaseDateString_(dateKey);
  if (!isoDate) {
  return { ok:false, errors:['Thiếu cột bắt buộc trong dữ liệu đăng ký xe (Register Date / Contract No).'] };
  }

  const vehicleRows = supabaseRequest_(
    SUPABASE_VEHICLE_REG_ENDPOINT
      + '?select=' + encodeURIComponent([
        'truck_plate',
        'country',
        'wheel',
        'trailer_plate',
        'driver_name',
        'id_passport',
        'phone_number',
        'transportation_company',
        'subcontractor',
        'contract_no',
        'register_date',
        'registration_status'
      ].join(','))
      + '&register_date=eq.' + encodeURIComponent(isoDate)
      + '&contract_no=eq.' + encodeURIComponent(contractNo)
      + '&registration_status=eq.' + encodeURIComponent('Approved')
  ) || [];

  if (!Array.isArray(vehicleRows) || !vehicleRows.length) {
    return { ok:false, errors:['Không có dòng Approved phù hợp để xuất.'] };
  }

  const rows = vehicleRows.map(function (row) {
    return {
      'Truck Plate':          s(row.truck_plate),
      'Country':              s(row.country),
      'Wheel':                row.wheel == null ? '' : row.wheel,
      'Trailer Plate':        s(row.trailer_plate),
      'Driver Name':          s(row.driver_name),
      'ID/Passport':          s(row.id_passport),
      'Phone number':         s(row.phone_number),
      'Transportation Company': s(row.transportation_company),
      'Subcontractor':        s(row.subcontractor)
    };
  });

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
    o['Transportation Company'] || '',
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
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    try {
      lock.waitLock(30 * 1000); // đảm bảo tuần tự hoá khi nhiều người cùng export
      locked = true;
    } catch (e) {
      return { ok:false, message:'Hệ thống đang bận. Vui lòng thử lại sau ít phút.' };
    }

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
    let resp;
    try {
      resp = _fetchWithRetry_(url, {
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
      });
    } catch (fetchErr) {
      return { ok:false, message:'Export lỗi: ' + (fetchErr && fetchErr.message ? fetchErr.message : fetchErr) };
    }
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
  } finally {
    if (locked) {
      try { lock.releaseLock(); } catch (e) {}
    }
  }
}

function _fetchWithRetry_(url, options) {
  const maxAttempts = 4;
  const baseDelayMs = 500;
  let lastError = null;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      const resp = UrlFetchApp.fetch(url, options);
      const code = resp.getResponseCode();
      if (code === 429 || (code >= 500 && code < 600)) {
        throw new Error('Drive export quota hit: ' + code);
      }
      return resp;
    } catch (err) {
      lastError = err;
      if (attempt === maxAttempts) {
        throw (err instanceof Error) ? err : new Error(err);
      }
      Utilities.sleep(baseDelayMs * Math.pow(2, attempt - 1));
    }
  }
  throw (lastError instanceof Error) ? lastError : new Error(lastError);
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
const TOTAL_LIST_EMPTY_MESSAGE_VI = 'Danh sách xe tổng chưa có dữ liệu. Không thể đăng ký. Vui lòng liên hệ PSVN.';
const TOTAL_LIST_EMPTY_MESSAGE_EN = 'The total vehicle list has no data. Unable to register. Please contact PSVN.';

function checkVehiclesAgainstTotalList(vehicles) {
  let rows = supabaseRequest_(
    SUPABASE_TRUCK_LIST_TOTAL_ENDPOINT + '?select=' + encodeURIComponent(['truck_plate', 'transportation_company'].join(','))
  ) || [];

  if (!Array.isArray(rows) || !rows.length) {
    return {
      isValid: false,
      message: TOTAL_LIST_EMPTY_MESSAGE_VI,
      messageEn: TOTAL_LIST_EMPTY_MESSAGE_EN
    };
  }

  const normalizePlate = function (value) {
    return String(value == null ? '' : value).toUpperCase().replace(/\s/g, '');
  };

  const totalListMap = new Map();
  rows.forEach(function (row) {
    const plate = normalizePlate(row.truck_plate);
    if (!plate) return;
    totalListMap.set(plate, String(row.transportation_company == null ? '' : row.transportation_company).trim());
  });

  if (!totalListMap.size) {
    return {
      isValid: false,
      message: TOTAL_LIST_EMPTY_MESSAGE_VI,
      messageEn: TOTAL_LIST_EMPTY_MESSAGE_EN
    };
  }

  const newPlates = [];
  const mismatchPlates = [];
  const seenNew = new Set();
  const seenMismatch = new Set();

  vehicles.forEach(function (vehicle) {
    const plate = normalizePlate(vehicle['Truck Plate']);
    if (!plate) return;
    const company = String(vehicle['Transportation Company'] || '').trim().toLowerCase();

    if (!totalListMap.has(plate)) {
      if (!seenNew.has(plate)) {
        seenNew.add(plate);
        newPlates.push(plate);
      }
      return;
    }

    const existingCompanyRaw = String(totalListMap.get(plate) || '').trim();
    const existingCompany = existingCompanyRaw.toLowerCase();
    if (existingCompany === '' && company === '') return;
    if (existingCompany !== company) {
      if (!seenMismatch.has(plate)) {
        seenMismatch.add(plate);
        mismatchPlates.push(plate);
      }
    }
  });

  if (newPlates.length > 0) {
    const message = `Xe biển số ${newPlates.join(', ')} chưa có trong danh sách xe tổng. Vui lòng liên hệ PSVN để đăng ký thêm.`;
    return {
      isValid: false,
      message: message,
      messageEn: `Vehicle plate(s) ${newPlates.join(', ')} are not in the total vehicle list. Please contact PSVN for assistance.`
    };
  }

  if (mismatchPlates.length > 0) {
    const message = `Xe biển số ${mismatchPlates.join(', ')} đã được đăng ký với đơn vị vận chuyển khác. Vui lòng liên hệ PSVN để xử lý.`;
    return {
      isValid: false,
      message: message,
      messageEn: `Vehicle plate(s) ${mismatchPlates.join(', ')} have already been registered with another transport company. Please contact PSVN for assistance.`
    };
  }

  return { isValid: true };
}

// =================================================================
// LOGIC XỬ LÝ ĐĂNG KÝ XE – Kiểm tra Activity Status
// =================================================================
function checkVehicleActivityStatus(vehicles) {
  let rows = supabaseRequest_(
    SUPABASE_TRUCK_LIST_TOTAL_ENDPOINT + '?select=' + encodeURIComponent(['truck_plate', 'activity_status'].join(','))
  ) || [];

  if (!Array.isArray(rows) || !rows.length) {
    return {
      isValid: false,
      message: TOTAL_LIST_EMPTY_MESSAGE_VI,
      messageEn: TOTAL_LIST_EMPTY_MESSAGE_EN
    };
  }

  const normalizePlate = function (value) {
    return String(value == null ? '' : value).toUpperCase().replace(/\s/g, '');
  };

  const activityMap = new Map();
  rows.forEach(function (row) {
    const plate = normalizePlate(row.truck_plate);
    if (!plate) return;
    activityMap.set(plate, String(row.activity_status == null ? '' : row.activity_status).trim().toLowerCase());
  });

  if (!activityMap.size) {
    return {
      isValid: false,
      message: TOTAL_LIST_EMPTY_MESSAGE_VI,
      messageEn: TOTAL_LIST_EMPTY_MESSAGE_EN
    };
  }

  const bannedPlates = [];
  vehicles.forEach(function (vehicle) {
    const plate = normalizePlate(vehicle['Truck Plate']);
    if (!plate) return;
    if (activityMap.get(plate) === 'banned') {
      bannedPlates.push(plate);
    }
  });

  if (bannedPlates.length > 0) {
    const message = `Xe biển số ${bannedPlates.join(', ')} đang trong tình trạng bị cấm, vui lòng liên hệ PSVN để xử lý.`;
    return {
      isValid: false,
      message: message,
      messageEn: `Vehicle plate(s) ${bannedPlates.join(', ')} are currently banned. Please contact PSVN for assistance.`
    };
  }

  return { isValid: true };
}


function getAllDataForExport(dateString, sessionToken, searchQuery, contractNo) {
  const userSession = validateSession(sessionToken);
  const role = String(userSession.role || '').toLowerCase();
  try {
    const params = ['select=' + encodeURIComponent(VEHICLE_REGISTRATION_SELECT_FIELDS.join(','))];
    const headers = HEADERS_REGISTER;

    if (dateString) {
      const iso = toSupabaseDateString_(dateString);
      if (!iso) return [];
      params.push('register_date=eq.' + encodeURIComponent(iso));
    }

    if (contractNo) {
      const contract = String(contractNo).replace(/^'+/, '').trim();
      if (contract) params.push('contract_no=eq.' + encodeURIComponent(contract));
    }

    if (role === 'user') {
      const contractor = String(userSession.contractor == null ? '' : userSession.contractor).trim();
      if (!contractor) return [];
      params.push('transportation_company=eq.' + encodeURIComponent(contractor));
    } else if (role === 'user-supervision') {
      params.push('registration_status=eq.' + encodeURIComponent('Approved'));
    }

    const rows = supabaseRequest_(SUPABASE_VEHICLE_REG_ENDPOINT + '?' + params.join('&')) || [];
    if (!Array.isArray(rows) || !rows.length) return [];

    let mapped = rows.map(function (row) { return mapVehicleRegistrationRowToArray_(row, headers); });

    if (searchQuery && String(searchQuery).trim()) {
      const q = String(searchQuery).toLowerCase();
      mapped = mapped.filter(function (r) {
        return r.some(function (c) { return String(c).toLowerCase().indexOf(q) !== -1; });
      });
    }

    return mapped.map(function (r) { return formatRowForClient_(r, headers); });
  } catch (e) {
    Logger.log(e);
    throw new Error('Cannot retrieve export data: ' + e.message);
  }
}

function resolveLoginHistoryCompanies_(filters) {
  const result = { companies: [], hasFilter: false };
  if (!filters || typeof filters !== 'object') return result;

  const rawCompany = stripLeadingApostrophe(filters.company);
  const company = String(rawCompany == null ? '' : rawCompany).trim();
  const rawContract = stripLeadingApostrophe(filters.contractNo);
  const contract = String(rawContract == null ? '' : rawContract).trim();

  const hasCompanyFilter = company.length > 0;
  const hasContractFilter = contract.length > 0;
  result.hasFilter = hasCompanyFilter || hasContractFilter;

  const selectedCompanies = new Set();
  if (hasCompanyFilter) {
    selectedCompanies.add(company);
  }

  let contractCompanies = [];
  if (hasContractFilter) {
    try {
      const contractFilters = ['contract_no=eq.' + encodeURIComponent(contract)];
      const rows = fetchContractDataRows_(['contract_no', 'transportation_company'], contractFilters);
      if (Array.isArray(rows) && rows.length) {
        const set = new Set();
        rows.forEach(function (row) {
          const comp = String(stripLeadingApostrophe(row && row.transportation_company) || '').trim();
          if (comp) set.add(comp);
        });
        contractCompanies = Array.from(set);
      }
    } catch (e) {
      Logger.log('resolveLoginHistoryCompanies_ error: ' + e);
    }
  }

  if (hasContractFilter) {
    if (selectedCompanies.size) {
      if (contractCompanies.length) {
        const intersection = Array.from(selectedCompanies).filter(function (item) {
          return contractCompanies.indexOf(item) !== -1;
        });
        result.companies = intersection;
      } else {
        result.companies = [];
      }
    } else {
      result.companies = contractCompanies;
    }
  } else {
    result.companies = Array.from(selectedCompanies);
  }

  return result;
}

function fetchLoginHistoryUsernamesForCompanies_(companies) {
  const sanitized = Array.isArray(companies)
    ? companies
        .map(function (value) { return stripLeadingApostrophe(value); })
        .map(function (value) { return String(value == null ? '' : value).trim(); })
        .filter(function (value) { return value.length > 0; })
    : [];
  if (!sanitized.length) return [];

  const unique = Array.from(new Set(sanitized));
  const filter = buildSupabaseInFilter_('contractor', unique);
  if (!filter) return [];

  try {
    const rows = supabaseRequest_(SUPABASE_APP_USERS_ENDPOINT + '?select=username&' + filter);
    if (!Array.isArray(rows) || !rows.length) return [];
    const seen = new Set();
    const usernames = [];
    rows.forEach(function (row) {
      const username = String(row && row.username == null ? '' : row.username).trim();
      if (username && !seen.has(username)) {
        seen.add(username);
        usernames.push(username);
      }
    });
    return usernames;
  } catch (e) {
    Logger.log('fetchLoginHistoryUsernamesForCompanies_ error: ' + e);
    return [];
  }
}

function buildLoginHistoryUserFilterClause_(filters) {
  const info = resolveLoginHistoryCompanies_(filters);
  if (!info.hasFilter) return '';
  if (!info.companies || !info.companies.length) {
    return 'username=eq.__no_matching_user__';
  }
  const usernames = fetchLoginHistoryUsernamesForCompanies_(info.companies);
  if (!usernames.length) {
    return 'username=eq.__no_matching_user__';
  }
  const inFilter = buildSupabaseInFilter_('username', usernames);
  return inFilter || 'username=eq.__no_matching_user__';
}

function buildLoginHistoryQueryParts_(filters, searchValue) {
  const parts = ['select=' + encodeURIComponent('*')];

  const userFilterClause = buildLoginHistoryUserFilterClause_(filters || {});
  if (userFilterClause) parts.push(userFilterClause);
  const fromIso = parseHistoryDateFilter_(filters && filters.dateFrom, false);
  if (fromIso) parts.push('occurred_at=gte.' + encodeURIComponent(fromIso));
  const toIso = parseHistoryDateFilter_(filters && filters.dateTo, true);
  if (toIso) parts.push('occurred_at=lte.' + encodeURIComponent(toIso));
  const sanitized = sanitizeHistorySearchTerm_(searchValue);
  if (sanitized) {
    const likeValue = `%${sanitized}%`;
    const orConditions = [
      `username.ilike.${likeValue}`,
      `outcome.ilike.${likeValue}`,
      `ip.ilike.${likeValue}`,
      `user_agent.ilike.${likeValue}`
    ];
    parts.push('or=' + encodeURIComponent('(' + orConditions.join(',') + ')'));
  }
  return parts;
}

function mapLoginHistoryRow_(row) {
  return {
    occurred_at: formatTimeForClient(row && row.occurred_at),
    username: toDisplayString_(row && row.username),
    outcome: toDisplayString_(row && row.outcome),
    ip: toDisplayString_(row && row.ip),
    latitude: toDisplayString_(row && row.latitude),
    longitude: toDisplayString_(row && row.longitude),
    accuracy_m: toDisplayString_(row && row.accuracy_m),
    user_agent: toDisplayString_(row && row.user_agent),
    created_at: formatTimeForClient(row && row.created_at)
  };
}

function getLoginHistoryServerSide(params) {
  params = params || {};
  requireAdmin_(params.sessionToken);

  const draw = parseInt(params.draw, 10) || 0;
  const start = Math.max(parseInt(params.start, 10) || 0, 0);
  let length = parseInt(params.length, 10);
  if (!isFinite(length) || length <= 0) length = 50;

  const columns = [
    'occurred_at',
    'username',
    'outcome',
    'ip',
    'latitude',
    'longitude',
    'accuracy_m',
    'user_agent',
    'created_at'
  ];

  let sortColumn = columns[0];
  let sortDir = 'desc';
  if (Array.isArray(params.order) && params.order.length) {
    const orderInfo = params.order[0] || {};
    const idx = parseInt(orderInfo.column, 10);
    if (!isNaN(idx) && idx >= 0 && idx < columns.length) {
      sortColumn = columns[idx];
    }
    if (String(orderInfo.dir).toLowerCase() === 'asc') {
      sortDir = 'asc';
    }
  }

  const queryParts = buildLoginHistoryQueryParts_(params.filters || {}, params.search && params.search.value);
  queryParts.push('limit=' + Math.max(length, 10));
  queryParts.push('offset=' + start);
  queryParts.push('order=' + encodeURIComponent(sortColumn + '.' + sortDir));

  const response = supabaseRequest_(SUPABASE_AUTH_LOGIN_HISTORY_ENDPOINT + '?' + queryParts.join('&'), {
    headers: { Prefer: 'count=exact' },
    returnResponse: true
  });
  const rows = Array.isArray(response.data) ? response.data : [];
  const total = parseContentRangeTotal_(response.headers);

  return {
    draw,
    recordsTotal: total != null ? total : rows.length,
    recordsFiltered: total != null ? total : rows.length,
    data: rows.map(mapLoginHistoryRow_)
  };
}

function getAllLoginHistoryForExport(filters, sessionToken, searchQuery) {
  requireAdmin_(sessionToken);
  const queryParts = buildLoginHistoryQueryParts_(filters || {}, searchQuery);
  const rows = supabaseRequest_(SUPABASE_AUTH_LOGIN_HISTORY_ENDPOINT + '?' + queryParts.join('&')) || [];
  const data = Array.isArray(rows) ? rows : [];
  return data.map(mapLoginHistoryRow_);
}

function buildVehicleHistoryQueryParts_(filters, searchValue, options) {
  const opts = options || {};
  const parts = ['select=' + encodeURIComponent('*')];
  const fromIso = parseHistoryDateFilter_(filters && filters.dateFrom, false);
  if (fromIso) parts.push('action_time=gte.' + encodeURIComponent(fromIso));
  const toIso = parseHistoryDateFilter_(filters && filters.dateTo, true);
  if (toIso) parts.push('action_time=lte.' + encodeURIComponent(toIso));

  const contract = String(filters && filters.contractNo || '').trim();
  if (contract) {
    parts.push('contract_no=eq.' + encodeURIComponent(contract));
  }

  const company = String(filters && filters.company || '').trim();
  if (company) {
    parts.push('transportation_comp=eq.' + encodeURIComponent(company));
  }

  const sanitized = sanitizeHistorySearchTerm_(searchValue);
  if (sanitized) {
    const likeValue = `%${sanitized}%`;
    const orConditions = [
      `action_type.ilike.${likeValue}`,
      `truck_plate.ilike.${likeValue}`,
      `driver_name.ilike.${likeValue}`,
      `id_passport.ilike.${likeValue}`,
      `phone_number.ilike.${likeValue}`,
      `subcontractor.ilike.${likeValue}`,
      `created_by.ilike.${likeValue}`
    ];
    if (!opts.skipContractSearch) {
      orConditions.splice(1, 0, `contract_no.ilike.${likeValue}`);
    }
    if (!opts.skipCompanySearch) {
      orConditions.push(`transportation_comp.ilike.${likeValue}`);
    }

    if (orConditions.length) {
      parts.push('or=' + encodeURIComponent('(' + orConditions.join(',') + ')'));
    }
  }

  return parts;
}

function mapVehicleHistoryRow_(row) {
  return {
    action_type: toDisplayString_(row && row.action_type),
    action_time: formatTimeForClient(row && row.action_time),
    register_date: formatDateForClient(row && row.register_date),
    contract_no: toDisplayString_(row && row.contract_no),
    truck_plate: toDisplayString_(row && row.truck_plate),
    country: toDisplayString_(row && row.country),
    wheel: toDisplayString_(row && row.wheel),
    trailer_plate: toDisplayString_(row && row.trailer_plate),
    truck_weight: toDisplayString_(row && row.truck_weight),
    pay_load: toDisplayString_(row && row.pay_load),
    container_no1: toDisplayString_(row && row.container_no1),
    container_no2: toDisplayString_(row && row.container_no2),
    driver_name: toDisplayString_(row && row.driver_name),
    id_passport: toDisplayString_(row && row.id_passport),
    phone_number: toDisplayString_(row && row.phone_number),
    destination_est: toDisplayString_(row && row.destination_est),
    transportation_comp: toDisplayString_(row && row.transportation_comp),
    subcontractor: toDisplayString_(row && row.subcontractor),
    vehicle_status: toDisplayString_(row && row.vehicle_status),
    registration_status: toDisplayString_(row && row.registration_status),
    time: formatTimeForClient(row && row.time),
    created_by: toDisplayString_(row && row.created_by),
    created_at: formatTimeForClient(row && row.created_at)
  };
}

function getVehicleRegistrationHistoryServerSide(params) {
  params = params || {};
  requireAdmin_(params.sessionToken);

  const draw = parseInt(params.draw, 10) || 0;
  const start = Math.max(parseInt(params.start, 10) || 0, 0);
  let length = parseInt(params.length, 10);
  if (!isFinite(length) || length <= 0) length = 50;

  const columns = [
    'action_type',
    'action_time',
    'register_date',
    'contract_no',
    'truck_plate',
    'country',
    'wheel',
    'trailer_plate',
    'truck_weight',
    'pay_load',
    'container_no1',
    'container_no2',
    'driver_name',
    'id_passport',
    'phone_number',
    'destination_est',
    'transportation_comp',
    'subcontractor',
    'vehicle_status',
    'registration_status',
    'time',
    'created_by',
    'created_at'
  ];

  let sortColumn = columns[1];
  let sortDir = 'desc';
  if (Array.isArray(params.order) && params.order.length) {
    const orderInfo = params.order[0] || {};
    const idx = parseInt(orderInfo.column, 10);
    if (!isNaN(idx) && idx >= 0 && idx < columns.length) {
      sortColumn = columns[idx];
    }
    if (String(orderInfo.dir).toLowerCase() === 'asc') {
      sortDir = 'asc';
    }
  }

  const filters = params.filters || {};
  const searchValue = params.search && params.search.value;

  function buildQueryParts(extraOptions) {
    const parts = buildVehicleHistoryQueryParts_(filters, searchValue, extraOptions);
    parts.push('limit=' + Math.max(length, 10));
    parts.push('offset=' + start);
    parts.push('order=' + encodeURIComponent(sortColumn + '.' + sortDir));
    return parts;
  }

  function executeQuery(parts) {
    return supabaseRequest_(SUPABASE_HISTORY_VEHICLE_REG_ENDPOINT + '?' + parts.join('&'), {
      headers: { Prefer: 'count=exact' },
      returnResponse: true
    });
  }

  let response;
  try {
    response = executeQuery(buildQueryParts());
  } catch (error) {
    if (error && /operator does not exist: uuid/i.test(error.message || '')) {
      response = executeQuery(buildQueryParts({ skipContractSearch: true, skipCompanySearch: true }));
    } else {
      throw error;
    }
  }
  const rows = Array.isArray(response.data) ? response.data : [];
  const total = parseContentRangeTotal_(response.headers);

  return {
    draw,
    recordsTotal: total != null ? total : rows.length,
    recordsFiltered: total != null ? total : rows.length,
    data: rows.map(mapVehicleHistoryRow_)
  };
}

function getAllVehicleHistoryForExport(filters, sessionToken, searchQuery) {
  requireAdmin_(sessionToken);
  const paramsFilters = filters || {};

  function fetchRows(extraOptions) {
    const queryParts = buildVehicleHistoryQueryParts_(paramsFilters, searchQuery, extraOptions);
    return supabaseRequest_(SUPABASE_HISTORY_VEHICLE_REG_ENDPOINT + '?' + queryParts.join('&')) || [];
  }

  let rows;
  try {
    rows = fetchRows();
  } catch (error) {
    if (error && /operator does not exist: uuid/i.test(error.message || '')) {
      rows = fetchRows({ skipContractSearch: true, skipCompanySearch: true });
    } else {
      throw error;
    }
  }
  const data = Array.isArray(rows) ? rows : [];
  return data.map(mapVehicleHistoryRow_);
}

function getHistoryFilterOptions(sessionToken) {
  requireAdmin_(sessionToken);
  const rows = fetchContractDataRows_(['contract_no', 'transportation_company', 'status']);
  if (!rows || !rows.length) {
    return { contracts: [], companies: [] };
  }
  const contracts = new Set();
  const companies = new Set();
  rows.forEach(function (row) {
    const status = String(row.status == null ? '' : row.status).trim().toLowerCase();
    if (status !== 'active') return;
    const contract = String(row.contract_no == null ? '' : row.contract_no).replace(/^'+/, '').trim();
    if (contract) contracts.add(contract);
    const company = String(row.transportation_company == null ? '' : row.transportation_company).trim();
    if (company) companies.add(company);
  });
  return {
    contracts: Array.from(contracts).sort(),
    companies: Array.from(companies).sort()
  };
}

function createMessagePicker_(language) {
  const lang = String(language || '').toLowerCase() === 'en' ? 'en' : 'vi';
  return function(vi, en) {
    return lang === 'en' && en ? en : vi;
  };
}

function saveData(dataToSave, sessionToken, language) {
  const userSession = validateSession(sessionToken);

  const pickMessage = createMessagePicker_(language);
  if (!dataToSave || dataToSave.length === 0) {
    throw new Error(pickMessage('Không có dữ liệu để lưu.', 'There is no data to save.'));
  }

  if (userSession.role !== 'admin') {
    const timeStatus = checkRegistrationTime();
    if (!timeStatus.isOpen) {
      throw new Error(pickMessage(
        'Đã hết thời gian cho phép đăng ký dữ liệu trong ngày.',
        'The allowed registration time for today has ended.'
      ));
    }
  }

  const activityResult = checkVehicleActivityStatus(dataToSave);
  if (!activityResult.isValid) {
    throw new Error(pickMessage(activityResult.message, activityResult.messageEn));
  }

  const validationResult = checkVehiclesAgainstTotalList(dataToSave);
  if (!validationResult.isValid) {
    throw new Error(pickMessage(validationResult.message, validationResult.messageEn));
  }

  // Kiểm tra Contract No thuộc đúng đơn vị & Active
  (function () {
    const activeMap = buildActiveContractMap_();
    const invalid = [];

    dataToSave.forEach(rec => {
      const cno = String(rec['Contract No'] || '').trim();
      const comp = String(
        (userSession.role === 'user' ? userSession.contractor : rec['Transportation Company']) || ''
      ).trim().toUpperCase();

      if (!cno || !comp || !activeMap.has(comp) || !activeMap.get(comp).has(cno)) {
        invalid.push(`${cno} (${comp})`);
      }
    });

    if (invalid.length > 0) {
      throw new Error(pickMessage(
        'Sai số hợp đồng, vui lòng kiểm tra lại hợp đồng vận chuyển (Contract No phải thuộc đúng đơn vị và đang Active): ' + invalid.join(', '),
        'Invalid contract numbers. Please verify the transport contract (Contract No must belong to the correct company and be Active): ' + invalid.join(', ')
      ));
    }
  })();

  const dupCheckRecords = dataToSave.map(r => {
    const obj = Object.assign({}, r);
    if (userSession.role === 'user') {
      obj['Transportation Company'] = userSession.contractor;
    }
    return obj;
  });

  const existingDuplicates = checkForExistingRegistrations(dupCheckRecords, sessionToken);
  if (existingDuplicates && existingDuplicates.length > 0) {
      throw new Error(pickMessage(
        `Các xe sau đã được đăng ký trong ngày: ${existingDuplicates.join(', ')}. Vui lòng kiểm tra lại.`,
        `The following vehicles have already been registered today: ${existingDuplicates.join(', ')}. Please verify.`
      ));
  }

  try {
    const payloads = dataToSave.map(function (obj) {
      if (userSession.role === 'user') {
        obj['Transportation Company'] = userSession.contractor;
      }
      coerceNumericRegisterFields_(obj);
      const normalizedDate = normalizeDate(obj['Register Date']);
      obj['Register Date'] = normalizedDate || obj['Register Date'];
      obj['Time'] = new Date();
      obj['Registration Status'] = 'Pending approval';
      const payload = buildVehicleRegistrationPayload_(obj, { includeNulls: true });
      if (payload && typeof payload === 'object' && Object.prototype.hasOwnProperty.call(payload, 'id')) {
        // Let Supabase auto-generate the primary key when inserting multiple records
        delete payload.id;
      }
      return payload;
    });

    supabaseRequest_(SUPABASE_VEHICLE_REG_ENDPOINT, {
      method: 'POST',
      headers: { Prefer: 'return=minimal' },
      payload: payloads
    });

    bumpSheetCacheVersion_(VEHICLE_REGISTRATION_CACHE_KEY);
    return pickMessage('Dữ liệu đã được lưu thành công!', 'Data saved successfully!');
  } catch (error) {
    Logger.log(error);
    throw new Error(pickMessage('Lỗi khi lưu dữ liệu: ' + error.message, 'Error saving data: ' + error.message));
  }
}

function updateData(rowData, sessionToken) {
  const userSession = validateSession(sessionToken);
  if (!rowData || !rowData.ID) throw new Error('Dữ liệu không hợp lệ hoặc thiếu ID.');

  try {
    const selectFields = encodeURIComponent(VEHICLE_REGISTRATION_SELECT_FIELDS.join(','));
    const existingRows = supabaseRequest_(
      SUPABASE_VEHICLE_REG_ENDPOINT + '?id=eq.' + encodeURIComponent(rowData.ID) + '&select=' + selectFields
    );

    if (!Array.isArray(existingRows) || !existingRows.length) {
      throw new Error('Không tìm thấy dòng với ID: ' + rowData.ID);
    }

    const existing = existingRows[0];
    
    if (userSession.role === 'user') {
      const timeStatus = checkRegistrationTime();
      if (!timeStatus.isOpen) throw new Error('Đã hết thời gian cho phép chỉnh sửa dữ liệu.');

      const recordTime = parseSupabaseTimestamp_(existing.time);
      const recordHour = recordTime
        ? parseInt(Utilities.formatDate(recordTime, 'Asia/Ho_Chi_Minh', 'HH'), 10)
        : NaN;

      const recordCompany = String(existing.transportation_company == null ? '' : existing.transportation_company).trim();
      if (recordCompany !== String(userSession.contractor || '').trim()) {
        throw new Error('Bạn không có quyền chỉnh sửa dữ liệu của đơn vị khác.');
      }

      if (!isNaN(recordHour)) {
        if (recordHour >= 8 && recordHour < 16 && timeStatus.period !== 1) {
          throw new Error('Dữ liệu đăng ký từ 8:00-16:00 chỉ có thể sửa trong khung giờ này.');
        }
        if (recordHour >= 20 && recordHour < 22 && timeStatus.period !== 2) {
          throw new Error('Dữ liệu đăng ký từ 20:00-22:00 chỉ có thể sửa trong khung giờ này.');
        }
      }
    }


    if (userSession.role === 'user') {
      rowData['Transportation Company'] = userSession.contractor;
    }
    
    if (rowData['Register Date']) {
      const normalizedDate = normalizeDate(rowData['Register Date']);
      rowData['Register Date'] = normalizedDate || rowData['Register Date'];
    }

    coerceNumericRegisterFields_(rowData);
    rowData['Time'] = new Date();

    const payload = buildVehicleRegistrationPayload_(rowData, { includeNulls: true });
    delete payload.id;

    supabaseRequest_(SUPABASE_VEHICLE_REG_ENDPOINT + '?id=eq.' + encodeURIComponent(rowData.ID), {
      method: 'PATCH',
      headers: { Prefer: 'return=minimal' },
      payload: payload
    });

    bumpSheetCacheVersion_(VEHICLE_REGISTRATION_CACHE_KEY);
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
    const uniqueIds = Array.from(new Set(ids.map(function (id) { return String(id).trim(); }).filter(Boolean)));
    if (!uniqueIds.length) throw new Error('Không tìm thấy dòng nào với các ID đã cho.');

    const select = ['id', 'transportation_company'];
    const selectParam = encodeURIComponent(select.join(','));
    const batches = chunkArray_(uniqueIds, SUPABASE_IN_FILTER_BATCH_SIZE);
    const existingMap = {};

    batches.forEach(function (batch) {
      const batchFilter = buildSupabaseInFilter_('id', batch);
      if (!batchFilter) return;
      const existingBatch = supabaseRequest_(
        SUPABASE_VEHICLE_REG_ENDPOINT + '?select=' + selectParam + '&' + batchFilter
      ) || [];
      if (Array.isArray(existingBatch) && existingBatch.length) {
        existingBatch.forEach(function (row) {
          if (row && row.id != null) {
            existingMap[String(row.id)] = row;
          }
        });
      }
    });

    const existing = Object.values(existingMap);

    if (!existing.length) {
      throw new Error('Không tìm thấy dòng nào với các ID đã cho.');
    }

    if (userSession.role === 'user') {
      const contractor = String(userSession.contractor || '').trim();
      existing.forEach(function (row) {
        const comp = String(row.transportation_company == null ? '' : row.transportation_company).trim();
        if (comp !== contractor) {
          throw new Error(`Bạn không có quyền xóa xe có ID: ${row.id}.`);
        }
      });
    }

    const idsToDelete = Object.keys(existingMap);
    const deleteBatches = chunkArray_(idsToDelete, SUPABASE_IN_FILTER_BATCH_SIZE);

    deleteBatches.forEach(function (batch) {
      const batchFilter = buildSupabaseInFilter_('id', batch);
      if (!batchFilter) return;
      supabaseRequest_(SUPABASE_VEHICLE_REG_ENDPOINT + '?' + batchFilter, {
        method: 'DELETE',
        headers: { Prefer: 'return=minimal' }
      });
    });

    bumpSheetCacheVersion_(VEHICLE_REGISTRATION_CACHE_KEY);
    return `Đã xóa thành công ${existing.length} mục.`;
  } catch (error) { Logger.log(error); throw new Error('Lỗi khi xóa dữ liệu: ' + error.message); }
}

function checkForExistingRegistrations(recordsToCheck, sessionToken) {
  validateSession(sessionToken);
  if (!recordsToCheck || recordsToCheck.length === 0) return [];

  try {
    const normalizedRecords = [];
    const uniqueDates = new Set();

    recordsToCheck.forEach(function (rec) {
      const regDate = normalizeDate(rec['Register Date']);
      const isoDate = toSupabaseDateString_(regDate) || '';
      const plate = String(rec['Truck Plate'] || '').toUpperCase().replace(/\s/g, '');
      const company = String(rec['Transportation Company'] || '').trim().toUpperCase();
      if (!isoDate || !plate || !company) return;
      normalizedRecords.push({ date: isoDate, plate: plate, company: company });
      uniqueDates.add(isoDate);
    });

    if (!normalizedRecords.length) return [];

    const dateFilter = buildSupabaseInFilter_('register_date', Array.from(uniqueDates));
    const existingKeys = new Set();
    if (dateFilter) {
      const existingRows = supabaseRequest_(
        SUPABASE_VEHICLE_REG_ENDPOINT + '?select=' + encodeURIComponent(['register_date', 'truck_plate', 'transportation_company'].join(',')) + '&' + dateFilter
      ) || [];

      if (Array.isArray(existingRows)) {
        existingRows.forEach(function (row) {
          const dateStr = String(row.register_date || '').trim();
          const plate = String(row.truck_plate || '').toUpperCase().replace(/\s/g, '');
          const company = String(row.transportation_company || '').trim().toUpperCase();
          if (dateStr && plate && company) {
            existingKeys.add(`${dateStr}-${plate}-${company}`);
          }
        });
      }
    }

    const seen = new Set();
    const duplicates = [];

    normalizedRecords.forEach(function (rec) {
      const key = `${rec.date}-${rec.plate}-${rec.company}`;

      if (existingKeys.has(key) || seen.has(key)) {
        duplicates.push(rec.plate);
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
    let rows = supabaseRequest_(
      SUPABASE_TRUCK_LIST_TOTAL_ENDPOINT + '?select=' + encodeURIComponent(['transportation_company', 'activity_status'].join(','))
    ) || [];

    if (!Array.isArray(rows) || !rows.length) {
      return { total: 0, active: 0, banned: 0 };
    }

    const role = String(userSession.role || '').toLowerCase();
    if (role === 'user') {
      const contractor = String(userSession.contractor == null ? '' : userSession.contractor).trim();
      rows = rows.filter(function (row) {
        const company = String(row.transportation_company == null ? '' : row.transportation_company).trim();
        const activity = String(row.activity_status == null ? '' : row.activity_status).trim().toLowerCase();
        return company === contractor && activity === 'active';
      });
    }

    const summary = { total: rows.length, active: 0, banned: 0 };
    rows.forEach(function (row) {
      const activity = String(row.activity_status == null ? '' : row.activity_status).trim().toLowerCase();
      if (activity === 'active') summary.active++;
      else if (activity === 'banned') summary.banned++;
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
    const normalizePlate = function (value) {
      return String(value || '').replace(/\s/g, '').toUpperCase();
    };

    let existingRows = supabaseRequest_(
      SUPABASE_TRUCK_LIST_TOTAL_ENDPOINT + '?select=' + encodeURIComponent(['truck_plate'].join(','))
    ) || [];

    const existingPlates = new Set();
    if (Array.isArray(existingRows)) {
      existingRows.forEach(function (row) {
        const plate = normalizePlate(row.truck_plate);
        if (plate) existingPlates.add(plate);
      });
    }

    const inFileSeen = new Set();
    const skippedInFile = [];
    const skippedExisting = [];
    const payloads = [];
    const username = String(userSession.username == null ? '' : userSession.username).trim();

    dataToSave.forEach(function (obj) {
      const plate = normalizePlate(obj['Truck Plate']);
      const company = obj['Transportation Company'] || '';
      if (!plate) return;

      if (inFileSeen.has(plate)) {
        skippedInFile.push({ plate: plate, company: company });
        return;
      }
      inFileSeen.add(plate);

      if (existingPlates.has(plate)) {
        skippedExisting.push({ plate: plate, company: company });
        return;
      }

      const record = Object.assign({}, obj);
      record['Truck Plate'] = plate;
      if (record['Trailer Plate']) {
        record['Trailer Plate'] = normalizePlate(record['Trailer Plate']);
      }
      coerceNumericRegisterFields_(record);
      if (username) {
        record['Created By'] = username;
        record['Updated By'] = username;
      }

      const payload = buildTruckListPayload_(record, { includeNulls: true });
      if (username) {
        payload.created_by = username;
        payload.updated_by = username;
      }
      payload.updated_at = new Date().toISOString();
      payloads.push(payload);
      existingPlates.add(plate);
    });

    if (payloads.length) {
      supabaseRequest_(SUPABASE_TRUCK_LIST_TOTAL_ENDPOINT, {
        method: 'POST',
        headers: { Prefer: 'return=minimal' },
        payload: payloads
      });
    }

    return {
      status: 'ok',
      inserted: payloads.length,
      skippedExisting: skippedExisting,
      skippedInFile: skippedInFile
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
    const sanitized = ids
      .map(function (id) { return String(id == null ? '' : id).trim(); })
      .filter(function (id) { return id.length > 0; });

    if (!sanitized.length) throw new Error('Không có ID hợp lệ để xóa.');

    const unique = Array.from(new Set(sanitized));
    const batches = chunkArray_(unique, SUPABASE_IN_FILTER_BATCH_SIZE);

    if (!batches.length) throw new Error('Không có ID hợp lệ để xóa.');

    let deletedCount = 0;

    batches.forEach(function (batch) {
      const filter = buildSupabaseInFilter_('id', batch);
      if (!filter) return;

      const response = supabaseRequest_(SUPABASE_TRUCK_LIST_TOTAL_ENDPOINT + '?' + filter, {
        method: 'DELETE',
        headers: { Prefer: 'return=representation' }
      });

      if (Array.isArray(response)) {
        deletedCount += response.length;
      } else {
        deletedCount += batch.length;
      }
    });

    if (!deletedCount) throw new Error('Không thể xóa xe.');

    return `Đã xóa thành công ${deletedCount} xe.`;
  } catch (error) { Logger.log(error); throw new Error('Lỗi khi xóa xe: ' + error.message); }
}

function updateTotalListVehicle(rowData, sessionToken) {
  const session = validateSession(sessionToken);
  if (session.role !== 'admin') throw new Error('Bạn không có quyền truy cập chức năng này.');
  if (!rowData || !rowData.ID) throw new Error('Dữ liệu không hợp lệ hoặc thiếu ID.');

  try {
    const id = String(rowData.ID || '').trim();
    if (!id) throw new Error('Thiếu ID hợp lệ.');

    const normalized = Object.assign({}, rowData);
    if (normalized['Truck Plate']) {
      normalized['Truck Plate'] = String(normalized['Truck Plate']).replace(/\s/g, '').toUpperCase();
    }
    if (normalized['Trailer Plate']) {
      normalized['Trailer Plate'] = String(normalized['Trailer Plate']).replace(/\s/g, '').toUpperCase();
    }

    coerceNumericRegisterFields_(normalized);

    const payload = buildTruckListPayload_(normalized, { includeNulls: true });
    payload.updated_at = new Date().toISOString();
    const username = String(session.username == null ? '' : session.username).trim();
    if (username) {
      payload.updated_by = username;
    }

    const response = supabaseRequest_(
      SUPABASE_TRUCK_LIST_TOTAL_ENDPOINT + '?id=eq.' + encodeURIComponent(id),
      {
        method: 'PATCH',
        headers: { Prefer: 'return=representation' },
        payload: payload
      }
    );

    if (!Array.isArray(response) || !response.length) {
      throw new Error('Không tìm thấy xe với ID: ' + rowData.ID);
    }

    return 'Cập nhật thông tin xe thành công!';
  } catch (error) { Logger.log(error); throw new Error('Lỗi khi cập nhật thông tin xe: ' + error.message); }
}

// --- Helpers an toàn cho CacheService ---
function safeScriptCacheGetJSON_(key) {
  if (!key) return null;
  try {
    const cache = CacheService.getScriptCache();
    if (!cache) return null;
    const value = cache.get(key);
    return value ? JSON.parse(value) : null;
  } catch (e) {
    Logger.log('Script cache get error: ' + e);
    return null;
  }
}

function safeScriptCachePutJSON_(key, obj, seconds) {
  if (!key) return;
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(obj), seconds || SERVER_SIDE_CACHE_TTL_SECONDS);
  } catch (e) {
    Logger.log('Script cache put error: ' + e);
  }
}

function safeScriptCacheRemove_(key) {
  if (!key) return;
  try {
    CacheService.getScriptCache().remove(key);
  } catch (e) {
    Logger.log('Script cache remove error: ' + e);
  }
}

function removeSupabaseUserMissCache_(username) {
  const missKey = buildSupabaseUserMissCacheKey_(username);
  if (!missKey) return;
  safeScriptCacheRemove_(missKey);
}

function getSheetCacheVersion_(sheetName) {
  if (!sheetName) return String(Date.now());
  try {
    const props = PropertiesService.getScriptProperties();
    const key = SHEET_CACHE_VERSION_PREFIX + sheetName;
    let version = props.getProperty(key);
    if (!version) {
      version = String(Date.now());
      props.setProperty(key, version);
    }
    return version;
  } catch (e) {
    Logger.log('getSheetCacheVersion_ error: ' + e);
    return String(Date.now());
  }
}

function bumpSheetCacheVersion_(sheetName) {
  if (!sheetName) return;
  try {
    const props = PropertiesService.getScriptProperties();
    const key = SHEET_CACHE_VERSION_PREFIX + sheetName;
    const newVersion = String(Date.now()) + '_' + Math.floor(Math.random() * 1000);
    props.setProperty(key, newVersion);
  } catch (e) {
    Logger.log('bumpSheetCacheVersion_ error: ' + e);
  }
}

function buildServerSideCacheKey_(sheetName, params, userRole) {
  if (!sheetName) return '';
  try {
    params = params || {};
    const version = getSheetCacheVersion_(sheetName);
    const payload = {
      sheet: sheetName,
      version: version,
      role: userRole || '',
      draw: params.draw || '',
      start: params.start || 0,
      length: params.length || 0,
      dateString: params.dateString || '',
      contractNo: params.contractNo || '',
      search: params.search && params.search.value ? String(params.search.value) : '',
      order: Array.isArray(params.order)
        ? params.order.map(function (o) { return o ? [o.column, o.dir] : null; })
        : [],
      columns: Array.isArray(params.columns)
        ? params.columns.map(function (col) {
            return {
              data: col && col.data,
              search: col && col.search ? col.search.value : '',
              searchable: col && col.searchable,
              orderable: col && col.orderable
            };
          })
        : []
    };
    const serialized = JSON.stringify(payload);
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, serialized);
    const hash = Utilities.base64EncodeWebSafe(digest).replace(/=+$/g, '');
    return 'srv_cache:' + sheetName + ':' + version + ':' + hash;
  } catch (e) {
    Logger.log('buildServerSideCacheKey_ error: ' + e);
    return '';
  }
}

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

function safeRemoveUserCacheKey(key) {
  try {
    CacheService.getUserCache().remove(key);
  } catch (e) {
    Logger.log('CacheService remove error: ' + e);
  }
}

function buildSessionCacheKey_(token) {
  return token ? 'user_session_' + token : '';
}

function getSessionFromCache_(token) {
  const key = buildSessionCacheKey_(token);
  if (!key) return null;
  return safeGetUserCacheJSON(key);
}

function cacheSession_(session) {
  if (!session || !session.token) return;
  const key = buildSessionCacheKey_(session.token);
  if (!key) return;
  safePutUserCacheJSON(key, session, SESSION_TIMEOUT_SECONDS);
}

function removeSessionFromCache_(token) {
  const key = buildSessionCacheKey_(token);
  if (!key) return;
  safeRemoveUserCacheKey(key);
}

function refreshSessionExpiry_(username, token) {
  if (!username || !token) return;
  try {
    const query = SUPABASE_APP_USERS_ENDPOINT + '?select=active_session_token,session_token_expiry&username=eq.' + encodeURIComponent(username) + '&limit=1';
    const data = supabaseRequest_(query);
    if (!Array.isArray(data) || data.length === 0) return;
    const user = data[0];
    cacheSupabaseUserRecord_(username, user);    
    if (String(user.active_session_token == null ? '' : user.active_session_token).trim() !== token) return;

    const nowMs = Date.now();
    const currentExpiry = parseSupabaseTimestamp_(user.session_token_expiry);
    const desiredExpiry = new Date(nowMs + SESSION_TIMEOUT_SECONDS * 1000);
    const halfWindowMs = (SESSION_TIMEOUT_SECONDS * 1000) / 2;    

    let shouldUpdate = false;
    if (!currentExpiry) {
      shouldUpdate = true;
    } else if (currentExpiry.getTime() - nowMs < halfWindowMs) {
      shouldUpdate = true;
    }

    if (shouldUpdate) {
      updateSupabaseUserByUsername_(username, {
        session_token_expiry: desiredExpiry.toISOString()
      }, { returnMinimal: true });
    }
  } catch (e) {
    Logger.log('refreshSessionExpiry_ error: ' + e);
  }
}

// --- Fallback lookup trong sheet Users bằng sessionToken ---
function lookupSessionFromSheet(sessionToken) {
  if (!sessionToken) return null;
  try {
    const query = SUPABASE_APP_USERS_ENDPOINT + '?select=username,role,contractor,customer_name,active_session_token,session_token_expiry&active_session_token=eq.' + encodeURIComponent(sessionToken) + '&limit=1';
    const data = supabaseRequest_(query);
    if (!Array.isArray(data) || data.length === 0) return null;
    const user = data[0];
    if (user && user.username) {
      cacheSupabaseUserRecord_(user.username, user);
    }    

    const expiryDate = parseSupabaseTimestamp_(user.session_token_expiry);
    const nowMs = Date.now();
    if (!expiryDate || nowMs >= expiryDate.getTime()) {
      clearUserSession_(user.username, sessionToken);
      return null;
    }

    return {
      isLoggedIn: true,
      username: String(user.username == null ? '' : user.username).trim(),
      role: String(user.role == null ? '' : user.role).trim(),
      contractor: String(user.contractor == null ? '' : user.contractor).trim(),
      customerName: String(user.customer_name == null ? '' : user.customer_name).trim(),
      token: sessionToken
    };
  } catch (e) {
    Logger.log('lookupSessionFromSheet error: ' + e);
  }
  return null;
}

// ==========================
// THAY THẾ validateSession()
// ==========================

function validateSession(sessionToken) {
  const token = String(sessionToken == null ? '' : sessionToken).trim();
  if (!token) {
    throw new Error('Bạn chưa đăng nhập hoặc phiên đã hết hạn. Vui lòng đăng nhập lại.');
  }

  let session = getSessionFromCache_(token);
  if (session && session.token === token) {
    cacheSession_(session);
    refreshSessionExpiry_(session.username, session.token);
    return session;
  }

  session = lookupSessionFromSheet(token);
  if (session) {
    cacheSession_(session);
    refreshSessionExpiry_(session.username, session.token);
    return session;
  }

  throw new Error('Bạn chưa đăng nhập hoặc phiên đã hết hạn. Vui lòng đăng nhập lại.');
}

// ==========================
// THAY THẾ getUserSession()
// ==========================
function getUserSession(sessionToken) {
  try {
    ensureSupervisionAccount_();
  } catch (e) {
    Logger.log('ensureSupervisionAccount_ wrapper error: ' + e);
  }
  try {
    const token = String(sessionToken == null ? '' : sessionToken).trim();
    if (token) {
      return validateSession(token);
    }
  } catch (e) {
    // Bỏ qua lỗi cache, trả về khách ẩn danh
  }
  return { isLoggedIn: false, role: null, contractor: null, customerName: null };
}

// Trả về Map: company (UPPER) -> Set(contractNo) chỉ chứa hợp đồng Active
function buildActiveContractMap_() {
  const rows = fetchContractDataRows_(
    ['contract_no', 'transportation_company', 'status']
  );

  const map = new Map();
  rows.forEach(function (row) {
    const status = String(row.status == null ? '' : row.status).trim().toLowerCase();
    if (status !== 'active') return;
    const no = String(row.contract_no == null ? '' : row.contract_no).replace(/^'+/, '').trim();
    const comp = String(row.transportation_company == null ? '' : row.transportation_company)
      .trim()
      .toUpperCase();
    if (!no || !comp) return;
    if (!map.has(comp)) map.set(comp, new Set());
    map.get(comp).add(no);
  });
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

function getContractDataServerSide(params) {
  const session = validateSession(params.sessionToken);
  const role = String(session.role == null ? '' : session.role).toLowerCase();
  const isAdmin = role === 'admin';

  const rows = fetchContractDataRows_(CONTRACT_DATA_SELECT_FIELDS);

  let data = rows.map(function (row) {
    const createdAt = parseSupabaseTimestamp_(row.created_at);
    const createdAtDisplay = formatSupabaseDateTime_(row.created_at);
    const createdAtSort = createdAt ? createdAt.toISOString() : '';
    return {
      'ID': String(row.id == null ? '' : row.id).trim(),
      'Contract No': String(row.contract_no == null ? '' : row.contract_no).replace(/^'+/, '').trim(),
      'Customer Name': String(row.customer_name == null ? '' : row.customer_name).trim(),
      'Transportation Company': String(row.transportation_company == null ? '' : row.transportation_company).trim(),
      'Status': String(row.status == null ? '' : row.status).trim(),
      'Created At': createdAtDisplay,
      'Created At Sort': createdAtSort,
      'Created By': String(row.created_by == null ? '' : row.created_by).trim()
    };
  });

  if (!isAdmin) {
    const contractor = String(session.contractor == null ? '' : session.contractor).trim();
    if (contractor) {
      data = data.filter(function (item) {
        return String(item['Transportation Company'] || '') === contractor;
      });
    } else {
      data = [];
    }
  }

  const q = (params.search && params.search.value ? String(params.search.value) : '').toLowerCase();
  const searchKeys = ['ID','Contract No','Customer Name','Transportation Company','Status','Created At','Created By'];  
  let filtered = q
    ? data.filter(function (item) {
        return searchKeys.some(function (key) {
          return String(item[key] == null ? '' : item[key]).toLowerCase().includes(q);
        });
      })
    : data;

  const order = Array.isArray(params.order) ? params.order[0] : null;
  if (order && order.column != null) {
    const dataColumns = ['ID','Contract No','Customer Name','Transportation Company','Status','Created At','Created By'];
    let idx = Number(order.column);
    if (isNaN(idx)) idx = 0;
    idx -= 2; // Bỏ qua 2 cột Select & Action
    if (idx < 0) idx = 0;
    if (idx >= dataColumns.length) idx = dataColumns.length - 1;
    const key = dataColumns[idx];
    const sortKey = key === 'Created At' ? 'Created At Sort' : key;
    const dir = (order.dir || 'asc').toLowerCase() === 'desc' ? -1 : 1;
    filtered = filtered.slice().sort(function (a, b) {
      const valA = String(a[sortKey] == null ? '' : a[sortKey]);
      const valB = String(b[sortKey] == null ? '' : b[sortKey]);
      return valA.localeCompare(valB, undefined, { numeric: true }) * dir;
    });
  }

  const start = Number(params.start || 0);
  const length = Number(params.length || 50);
  const page = filtered.slice(start, start + length).map(function (item) {
    const { ['Created At Sort']: _discard, ...rest } = item;
    return rest;
  });

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
          'Transportation Company': tc, 'Status': status } = contract;

  const payload = {
    contract_no: String(contractNo == null ? '' : contractNo).trim(),
    customer_name: String(customerName == null ? '' : customerName).trim(),
    transportation_company: String(tc == null ? '' : tc).trim(),
    status: String(status == null ? '' : status).trim()
  };

  const username = String(session.username == null ? '' : session.username).trim();
  if (username) {
    if (ID) {
      payload.updated_by = username;
    } else {
      payload.created_by = username;
    }
  }

  if (!payload.contract_no) {
    throw new Error('Thiếu Contract No.');
  }

  if (ID) {
    const query = SUPABASE_CONTRACT_DATA_ENDPOINT + '?id=eq.' + encodeURIComponent(String(ID).trim());
    const result = supabaseRequest_(query, {
      method: 'PATCH',
      payload: payload,
      headers: { Prefer: 'return=representation' }
    });
    if (!Array.isArray(result) || !result.length) {
      throw new Error('Không tìm thấy ID để cập nhật.');
    }
    return 'Đã cập nhật hợp đồng.';
  }

  const insertResult = supabaseRequest_(SUPABASE_CONTRACT_DATA_ENDPOINT, {
    method: 'POST',
    payload: payload,
    headers: { Prefer: 'return=representation' }
  });
  if (!Array.isArray(insertResult) || !insertResult.length) {
    throw new Error('Không thể tạo hợp đồng mới.');
  }
  return 'Đã tạo hợp đồng mới.';
}


function deleteContracts(ids, sessionToken) {
  const session = validateSession(sessionToken);
  if (session.role !== 'admin') throw new Error('Bạn không có quyền thực hiện.');

  if (!Array.isArray(ids) || !ids.length) return 'Không có mục nào để xoá.';

  const sanitized = ids
    .map(function (id) { return String(id == null ? '' : id).trim(); })
    .filter(function (id) { return id.length > 0; });

  if (!sanitized.length) return 'Không có mục nào để xoá.';

  const filter = buildSupabaseInFilter_('id', sanitized);
  if (!filter) return 'Không có mục nào để xoá.';

  const result = supabaseRequest_(SUPABASE_CONTRACT_DATA_ENDPOINT + '?' + filter, {
    method: 'DELETE',
    headers: { Prefer: 'return=representation' }
  });

  const count = Array.isArray(result) ? result.length : 0;
  return `Đã xoá ${count} hợp đồng.`;
}

//Lấy danh sách Contractor từ Supabase (dropdown “Transportation Company” ở trang Hợp đồng)
function getContractorOptions() {
  try {
    const data = supabaseRequest_(
      SUPABASE_APP_USERS_ENDPOINT + '?select=contractor&contractor=not.is.null'
    );
    if (!Array.isArray(data)) return [];
    const set = new Set();
    data.forEach(row => {
      const value = String(row.contractor == null ? '' : row.contractor).trim();
      if (value) set.add(value);
    });
    return Array.from(set).sort();
  } catch (e) {
    Logger.log('getContractorOptions error: ' + e);
    return [];
  }
}


//Lấy danh sách "Đơn vị vận chuyển" từ Supabase truck_list_total
function getTransportCompanies() {
  try {
    let rows = supabaseRequest_(
      SUPABASE_TRUCK_LIST_TOTAL_ENDPOINT + '?select=' + encodeURIComponent(['transportation_company'].join(','))
    ) || [];

    if (!Array.isArray(rows) || !rows.length) return [];

    const set = new Set();
    rows.forEach(function (row) {
      const value = String(row.transportation_company == null ? '' : row.transportation_company).trim();
      if (value) set.add(value);
    });
    return Array.from(set).sort();
  } catch (e) {
    Logger.log('getTransportCompanies error: ' + e);
    return [];
  }
}

//Lấy Contract No (Status = Active) cho dropdown “Số HĐ” ở trang Đăng ký xe
function getActiveContractNos(sessionToken) {
  const session = validateSession(sessionToken);
  const role = String(session.role == null ? '' : session.role).toLowerCase();
  const contractor = String(session.contractor == null ? '' : session.contractor).trim();

  if (role !== 'admin' && !contractor) return [];

  const rows = fetchContractDataRows_(['contract_no', 'transportation_company', 'status']);
  if (!rows.length) return [];

  const seen = new Set();
  rows.forEach(function (row) {
    const status = String(row.status == null ? '' : row.status).trim().toLowerCase();
    if (status !== 'active') return;  

    const comp = String(row.transportation_company == null ? '' : row.transportation_company).trim();
    if (role !== 'admin' && contractor && comp !== contractor) return;

    const contractNo = String(row.contract_no == null ? '' : row.contract_no).replace(/^'+/, '').trim();
    if (!contractNo || seen.has(contractNo)) return;
    seen.add(contractNo);
  });

  return Array.from(seen).sort();
}



// ====== GS: Trả về danh sách biển số đang có để đánh dấu trùng ======
function getExistingTruckPlates(sessionToken) {
  validateSession(sessionToken);
  let rows = supabaseRequest_(
    SUPABASE_TRUCK_LIST_TOTAL_ENDPOINT + '?select=' + encodeURIComponent(['truck_plate'].join(','))
  ) || [];

  if (!Array.isArray(rows) || !rows.length) return [];

  const normalizePlate = function (value) {
    return String(value == null ? '' : value).replace(/\s/g, '').toUpperCase();
  };

  return rows
    .map(function (row) { return normalizePlate(row.truck_plate); })
    .filter(function (plate) { return plate.length > 0; });
}



// ====== GS: LƯU NỐI TIẾP VÀO "DANH SÁCH XE TỔNG" ======
function saveTotalListAppend(rows, sessionToken) {
  const session = validateSession(sessionToken);

  if (!rows || !rows.length) return 'Không có dữ liệu để lưu.';

  const normalizePlate = function (value) {
    return String(value || '').replace(/\s/g, '').toUpperCase();
  };

  const username = String(session.username == null ? '' : session.username).trim();
  const payloads = [];

  rows.forEach(function (obj) {
    const record = Object.assign({}, obj);
    const plate = normalizePlate(record['Truck Plate']);
    if (!plate) return;
    record['Truck Plate'] = plate;
    if (record['Trailer Plate']) {
      record['Trailer Plate'] = normalizePlate(record['Trailer Plate']);
    }
    coerceNumericRegisterFields_(record);
    if (username) {
      record['Created By'] = username;
      record['Updated By'] = username;
    }
    const payload = buildTruckListPayload_(record, { includeNulls: true });
    if (username) {
      payload.created_by = username;
      payload.updated_by = username;
    }
    payload.updated_at = new Date().toISOString();
    payloads.push(payload);
  });

  if (payloads.length) {
    supabaseRequest_(SUPABASE_TRUCK_LIST_TOTAL_ENDPOINT, {
      method: 'POST',
      headers: { Prefer: 'return=minimal' },
      payload: payloads
    });
  }

  return `Đã thêm ${payloads.length} dòng mới vào Danh sách xe tổng.`;
}

// === THAY TOÀN BỘ addManualVehicle ===
function addManualVehicle(record, sessionToken, language) {
  const userSession = validateSession(sessionToken);
  const pickMessage = createMessagePicker_(language);

  try {

    // Chuẩn hóa/điền thêm các field bắt buộc
    const rowObj = Object.assign({}, record);

    // ✅ NEW: khóa Contractor cho user thường
    if (userSession.role === 'user') {
      rowObj['Transportation Company'] = userSession.contractor || rowObj['Transportation Company'];
    }

    coerceNumericRegisterFields_(rowObj);

    const activityCheck = checkVehicleActivityStatus([{ 'Truck Plate': rowObj['Truck Plate'] }]);
    if (!activityCheck.isValid) {
     throw new Error(pickMessage(activityCheck.message, activityCheck.messageEn));
    }

    // ✅ NEW: 3 kiểm tra đối chiếu "Danh sách tổng" (dùng đúng thông báo như upload)
    const precheck = checkVehiclesAgainstTotalList([{
      'Truck Plate'         : String(rowObj['Truck Plate'] || '').toUpperCase().replace(/\s/g, ''),
      'Transportation Company': rowObj['Transportation Company']
    }]);
    if (!precheck.isValid) {
      throw new Error(pickMessage(precheck.message, precheck.messageEn));
    }

    // ✅ NEW: kiểm tra Contract No thuộc đúng Contractor & Active
    const contractNo = String(rowObj['Contract No'] || '').trim();
    const company    = String(rowObj['Transportation Company'] || '').trim();
    if (!isContractActiveForCompany_(contractNo, company)) {
      throw new Error(pickMessage(
        'Sai số hợp đồng, vui lòng kiểm tra lại hợp đồng vận chuyển (Contract No phải thuộc đúng đơn vị và đang Active).',
        'Invalid contract numbers. Please verify the transport contract (Contract No must belong to the correct company and be Active).'
      ));
    }

    // ✅ NEW: kiểm tra xe đã đăng ký trong ngày (tái dùng logic của saveData)
    const dup = checkForExistingRegistrations([{
      'Register Date'       : rowObj['Register Date'],
      'Truck Plate'         : rowObj['Truck Plate'],
      'Transportation Company': rowObj['Transportation Company']
    }], sessionToken);
    if (dup && dup.length > 0) {
      throw new Error(pickMessage(
        `Các xe sau đã được đăng ký trong ngày: ${dup.join(', ')}. Vui lòng kiểm tra lại.`,
        `The following vehicles have already been registered today: ${dup.join(', ')}. Please verify.`
      ));
    }

    // Lưu ngày dạng text (chỉ thêm 1 dấu ')
    if (rowObj['Register Date']) {
      rowObj['Register Date'] = normalizeDate(rowObj['Register Date']);
    }

    // Thời gian tạo (giữ nguyên cách lưu hiện tại)
    rowObj['Time'] = new Date();

        // ✅ Bổ sung cột Registration Status
    rowObj['Registration Status'] = 'Pending approval';

    const payload = buildVehicleRegistrationPayload_(rowObj, { includeNulls: true });

    supabaseRequest_(SUPABASE_VEHICLE_REG_ENDPOINT, {
      method: 'POST',
      headers: { Prefer: 'return=minimal' },
      payload: [payload]
    });
    bumpSheetCacheVersion_(VEHICLE_REGISTRATION_CACHE_KEY);

    return pickMessage('Đăng ký xe thành công!', 'Vehicle registered successfully!');
  } catch (e) {
    Logger.log(e);
    throw new Error('Add New Error: ' + e.message);
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

  var isoDate = toSupabaseDateString_(dateString);
  if (!isoDate) throw new Error('Thiếu ngày đăng ký hợp lệ');
  var headers = HEADERS_REGISTER;
  var response = supabaseRequest_(
    SUPABASE_VEHICLE_REG_ENDPOINT
      + '?select=' + encodeURIComponent(VEHICLE_REGISTRATION_SELECT_FIELDS.join(','))
      + '&register_date=eq.' + encodeURIComponent(isoDate)
  ) || [];

  if (!Array.isArray(response)) throw new Error('Không thể đọc dữ liệu đăng ký.');

  var rowsRaw = response.map(function (row) { return mapVehicleRegistrationRowToArray_(row, headers); });

  var rows = [];
  var pending=0, approved=0;
  var per={};

  for (var i=0;i<rowsRaw.length;i++){
    var r=rowsRaw[i];
    var obj = formatRowForClient_(r, headers);
    if (scope !== 'ALL' && String(obj['Transportation Company']||'').trim() !== scope) continue;
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
  requireAdmin_(sessionToken);
  var dateString = filters && filters.dateString;
  var set = {};
  (filters.contracts||[]).forEach(function(c){ set[String(c).trim()] = true; });
  var idsSelected = (filters.idsSelected||[]).map(String);

  if (!dateString) throw new Error('Thiếu ngày đăng ký.');
  if (!newStatus || (['Approved','Pending approval'].indexOf(newStatus)===-1)) throw new Error('Trạng thái không hợp lệ.');

  var isoDate = toSupabaseDateString_(dateString);
  if (!isoDate) throw new Error('Ngày đăng ký không hợp lệ.');

  var baseParams = [
    'select=' + encodeURIComponent(['id','contract_no','registration_status','transportation_company'].join(',')),
    'register_date=eq.' + encodeURIComponent(isoDate)
  ];

  var rows = [];
  if (idsSelected && idsSelected.length) {
    var batches = chunkArray_(idsSelected, SUPABASE_IN_FILTER_BATCH_SIZE);
    batches.forEach(function(batch) {
      var idFilter = buildSupabaseInFilter_('id', batch);
      if (!idFilter) return;
      var batchRows = supabaseRequest_(SUPABASE_VEHICLE_REG_ENDPOINT + '?' + baseParams.concat(idFilter).join('&')) || [];
      if (Array.isArray(batchRows) && batchRows.length) {
        rows = rows.concat(batchRows);
      }
    });
  } else {
    rows = supabaseRequest_(SUPABASE_VEHICLE_REG_ENDPOINT + '?' + baseParams.join('&')) || [];
  }

  if (!Array.isArray(rows) || !rows.length) return 'Không có dòng nào được cập nhật.';

  var changedIdsSet = {};
  rows.forEach(function (row) {
    if (Object.keys(set).length && !set[String(row.contract_no || '').trim()]) return;
    if (String(row.registration_status || '') === newStatus) return;
    if (row.id != null) changedIdsSet[String(row.id)] = true;
  });

  var changedIds = Object.keys(changedIdsSet);
  if (!changedIds.length) return 'Không có dòng nào được cập nhật.';

  var updateBatches = chunkArray_(changedIds, SUPABASE_IN_FILTER_BATCH_SIZE);
  updateBatches.forEach(function (batch) {
    var updateFilter = buildSupabaseInFilter_('id', batch);
    if (!updateFilter) return;
    supabaseRequest_(SUPABASE_VEHICLE_REG_ENDPOINT + '?' + updateFilter, {
      method: 'PATCH',
      headers: { Prefer: 'return=minimal' },
      payload: { registration_status: newStatus }
    });
  });

  bumpSheetCacheVersion_(VEHICLE_REGISTRATION_CACHE_KEY);
  _bust(['SNAP_'+dateString+'_ALL']);
  return 'Đã cập nhật ' + changedIds.length + ' dòng.';
}


/**
 * Trả về toàn bộ dữ liệu (đÃ format) theo bộ lọc hiện tại để xuất Excel.
 * params: { dateString?: 'dd/MM/yyyy', search?: string }
 */
function exportRegisteredVehicles(params) {
  const headers = HEADERS_REGISTER;
  const queryParams = ['select=' + encodeURIComponent(VEHICLE_REGISTRATION_SELECT_FIELDS.join(','))];

  if (params && params.dateString) {
    const iso = toSupabaseDateString_(params.dateString);
    if (iso) {
      queryParams.push('register_date=eq.' + encodeURIComponent(iso));
    }
  }

  let response = supabaseRequest_(SUPABASE_VEHICLE_REG_ENDPOINT + '?' + queryParams.join('&')) || [];
  if (!Array.isArray(response)) response = [];

  let rowsData = response.map(function (row) { return mapVehicleRegistrationRowToArray_(row, headers); });
  if (params && params.search) {
    const q = String(params.search).toLowerCase();
    rowsData = rowsData.filter(function (row) {
      return row.some(function (cell) { return String(cell).toLowerCase().includes(q); });
    });
  }

  const rows = rowsData.map(function (row) { return formatRowForClient_(row, headers); });
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
      r['Transportation Company']||'',
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

  // build valid Contract-Customer set from contract_data in Supabase
  const normalize = value => String(value == null ? '' : value).replace(/^'+/, '').trim();
  const contractNos = Array.from(new Set((rows || []).map(r => normalize(r['ContractNo'])))).filter(Boolean);
  const validSet = new Set();
  if (contractNos.length) {
    const filters = [];
    const inFilter = buildSupabaseInFilter_('contract_no', contractNos);
    if (inFilter) filters.push(inFilter);
    const contractRows = fetchContractDataRows_(['contract_no', 'customer_name'], filters);
    contractRows.forEach(function (row) {
      const key = normalize(row.customer_name) + '|' + normalize(row.contract_no);
      if (key !== '|') validSet.add(key);
    });
  }

  const userLabel = sanitizeXpplText_(user.username || user.user || user.email || '');
  const toInsert = rows.map(function (row) {
    const customer = normalize(row && row['Customer Name']);
    const contract = normalize(row && row['ContractNo']);
    const key = customer + '|' + contract;
    if (validSet.size && !validSet.has(key)) {
      throw new Error('Sai tên khách hàng hoặc số hợp đồng: ' + key);
    }
    const payload = buildXpplDatabasePayload_(row, { userLabel: userLabel });
    const usernameValue = row && Object.prototype.hasOwnProperty.call(row, 'Username')
      ? row['Username']
      : (row && row.Username);
    payload.username = sanitizeXpplText_(usernameValue || userLabel);
    return payload;
  });

  if (!toInsert.length) {
    return 'Không có dữ liệu.';
  }

  const response = supabaseRequest_(SUPABASE_XPPL_DATABASE_ENDPOINT, {
    method: 'POST',
    payload: toInsert,
    headers: { Prefer: 'return=representation' }
  });

  const count = Array.isArray(response) ? response.length : toInsert.length;
  return 'Đã lưu ' + count + ' dòng.';
}

function getXpplWeighingData(filter, sessionToken) {
  requireXpplRole_(sessionToken);
  const s = v => String(v == null ? '' : v).replace(/^'+/, '').trim();
  const filterDate = s(filter && filter.date);
  const dateKey = filterDate ? _toDateKey(filterDate) : null;
  if (!dateKey) {
    return { data: [], summary: { trucks:0, weight:0 }, contracts: [], customers: [] };
  }

  const isoDate = toSupabaseDateString_(dateKey);
  if (!isoDate) {
    return { data: [], summary: { trucks:0, weight:0 }, contracts: [], customers: [] };
  }

  const queryParts = [
    'select=' + encodeURIComponent('*'),
    'date_out=eq.' + encodeURIComponent(isoDate)
  ];

  let records = [];
  try {
    const res = supabaseRequest_(SUPABASE_XPPL_DATABASE_ENDPOINT + '?' + queryParts.join('&'));
    records = Array.isArray(res) ? res : [];
  } catch (e) {
    Logger.log('getXpplWeighingData error: ' + e);
    records = [];
  }

  if (!records.length) {
    return { data: [], summary: { trucks:0, weight:0 }, contracts: [], customers: [] };
  }

  const filterContract = s(filter && filter.contractNo);
  const filterCustomer = s(filter && filter.customerName);

  const contracts = new Set();
  const customers = new Set();
  const rows = [];
  let totalWeight = 0;

  records.forEach(function (record) {
    const arr = mapXpplRecordToRowArray_(record, XPPL_DB_HEADERS);
    const formatted = formatRowForClient_(arr, XPPL_DB_HEADERS);
    const contractValue = s(formatted['ContractNo']);
    const customerValue = s(formatted['Customer Name']);
    contracts.add(contractValue);
    customers.add(customerValue);
    if (filterContract && filterContract !== contractValue) return;
    if (filterCustomer && filterCustomer !== customerValue) return;
    rows.push(formatted);
    const netRaw = record && record.net_weight;
    if (netRaw !== null && netRaw !== undefined && netRaw !== '') {
      let net = netRaw;
      if (typeof net === 'string') {
        net = Number(net.replace(/,/g, ''));
      }
      if (typeof net === 'number' && isFinite(net)) {
        totalWeight += net;
      }
    }
  });

  return {
    data: rows,
    summary: { trucks: rows.length, weight: totalWeight },
    contracts: Array.from(contracts).filter(Boolean).sort(),
    customers: Array.from(customers).filter(Boolean).sort()
  };
}

// ===== WEIGHING RESULT HELPERS =====
function matchTransportationCompanies(filter, sessionToken) {
  const user = requireXpplRole_(sessionToken);
  const normalizePlate = function (value) {
    return String(value == null ? '' : value)
      .replace(/\s/g, '')
      .toUpperCase();
  };
  const f = filter || {};
  const from = _toDateKey(f.dateFrom);
  const to = _toDateKey(f.dateTo);
  const isoFrom = from ? toSupabaseDateString_(from) : null;
  const isoTo = to ? toSupabaseDateString_(to) : null;

  const queryParts = [
    'select=' + encodeURIComponent(['id', 'truck_no', 'transportation_company', 'username', 'date_out'].join(','))
  ];
  if (isoFrom) queryParts.push('date_out=gte.' + encodeURIComponent(isoFrom));
  if (isoTo) queryParts.push('date_out=lte.' + encodeURIComponent(isoTo));

  let records = [];
  try {
    const res = supabaseRequest_(SUPABASE_XPPL_DATABASE_ENDPOINT + '?' + queryParts.join('&'));
    records = Array.isArray(res) ? res : [];
  } catch (e) {
    Logger.log('matchTransportationCompanies error: ' + e);
    records = [];
  }

  if (!records.length) return 'Không tìm thấy dữ liệu phù hợp.';

  const targetPlates = new Set();
  records.forEach(function (record) {
    const plate = normalizePlate(record && record.truck_no);
    if (plate) targetPlates.add(plate);
  });

  const plateMap = new Map();
  if (targetPlates.size) {
    const vehicleQuery = ['select=' + encodeURIComponent(['truck_plate', 'transportation_company'].join(','))];
    if (isoFrom) vehicleQuery.push('register_date=gte.' + encodeURIComponent(isoFrom));
    if (isoTo) vehicleQuery.push('register_date=lte.' + encodeURIComponent(isoTo));
    vehicleQuery.push('transportation_company=not.is.null');

    try {
      const vehicleRows = supabaseRequest_(SUPABASE_VEHICLE_REG_ENDPOINT + '?' + vehicleQuery.join('&'));
      if (Array.isArray(vehicleRows)) {
        vehicleRows.forEach(function (row) {
          const plate = normalizePlate(row && row.truck_plate);
          if (!plate || !targetPlates.has(plate)) return;
          const company = sanitizeXpplText_(row && row.transportation_company);
          if (!company) return;
          if (!plateMap.has(plate)) {
            plateMap.set(plate, company);
          }
        });
      }
    } catch (err) {
      Logger.log('matchTransportationCompanies vehicle registration fetch error: ' + err);
    }
  }

  const uname = sanitizeXpplText_(user.username || user.user || user.email || '');
  const updates = [];

  records.forEach(function (record) {
    const dateKey = _toDateKey(record && record.date_out);
    if (from && dateKey && dateKey < from) return;
    if (to && dateKey && dateKey > to) return;
    const plate = normalizePlate(record && record.truck_no);
    if (!plate || !record || !record.id) return;
    let compName = plateMap.get(plate);
    if (compName) {
      compName = sanitizeXpplText_(compName);
    } else {
      compName = 'Unknown';
    }
    compName = sanitizeXpplText_(compName || 'Unknown');
    updates.push({
      id: record.id,
      transportation_company: compName,
      username: uname
    });
  });

  if (!updates.length) return 'Không tìm thấy dữ liệu phù hợp.';

  supabaseRequest_(SUPABASE_XPPL_DATABASE_ENDPOINT, {
    method: 'POST',
    payload: updates,
    headers: { Prefer: 'resolution=merge-duplicates,return=minimal' }
  });

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
  const headers = XPPL_DB_HEADERS;
  const f = params.filter || {};
  const from = _toDateKey(f.dateFrom);
  const to = _toDateKey(f.dateTo);
  const normalizeListInput = function(value) {
    if (value == null) return [];
    if (Array.isArray(value)) {
      return value.map(function(v){ return String(v == null ? '' : v).trim(); }).filter(function(v){ return v; });
    }
    const str = String(value == null ? '' : value).trim();
    return str ? [str] : [];
  };

  const contractFilter = normalizeListInput(f.contracts);
  const customerFilter = normalizeListInput(f.customers);
  const draw = Number(params.draw || 1);
  const isUser = String(session.role || '').toLowerCase() === 'user';
  const parseCustomerAssignment = function(value) {
    if (value == null) return [];
    if (Array.isArray(value)) {
      return value.map(function(v){ return String(v == null ? '' : v).trim(); }).filter(function(v){ return v; });
    }
    return String(value)
      .split(/[\r\n;|]+/)
      .map(function(v){ return v.trim(); })
      .filter(function(v){ return v; });
  };
  const assignedCustomerNames = isUser
    ? parseCustomerAssignment(session.customerName || session.customerNames)
    : [];
  const uniqueSortedList = function(list) {
    if (!Array.isArray(list)) return [];
    const seen = new Set();
    const out = [];
    for (var i = 0; i < list.length; i++) {
      var val = String(list[i] == null ? '' : list[i]).trim();
      if (!val) continue;
      var key = val.toLowerCase();
      if (seen.has(key)) continue;
      seen.add(key);
      out.push(val);
    }
    return out.sort(function(a, b) {
      return String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: 'base' });
    });
  };
  const buildEmpty = function(customerNames) {
    return {
      draw: draw,
      recordsTotal: 0,
      recordsFiltered: 0,
      data: [],
      counts: { unassigned: 0, unknown: 0, assigned: 0 },
      summary: { trucks: 0, weight: 0 },
      options: {
        contracts: [],
        customers: uniqueSortedList(customerNames)
      }
    };
  };

  if (!from && !to && !contractFilter.length && !customerFilter.length && !(params.search && params.search.value)) {
    return buildEmpty(isUser ? assignedCustomerNames : null);
  }

  const idxDateOut = headers.indexOf('Date Out');
  const idxContract = headers.indexOf('ContractNo');
  const idxCompany = headers.indexOf('Transportation Company');
  const idxCustomer = headers.indexOf('Customer Name');
  const idxNetWeight = headers.indexOf('Net Weight');
  if (idxDateOut === -1 || idxContract === -1 || idxCompany === -1 || idxCustomer === -1) {
    return buildEmpty(isUser ? assignedCustomerNames : null);
  }

  const queryParts = ['select=' + encodeURIComponent('*')];
  if (from) {
    const isoFrom = toSupabaseDateString_(from);
    if (isoFrom) queryParts.push('date_out=gte.' + encodeURIComponent(isoFrom));
  }
  if (to) {
    const isoTo = toSupabaseDateString_(to);
    if (isoTo) queryParts.push('date_out=lte.' + encodeURIComponent(isoTo));
  }
  const contractIn = buildSupabaseInFilter_('contract_no', contractFilter);
  if (contractIn) queryParts.push(contractIn);
  const customerIn = buildSupabaseInFilter_('customer_name', customerFilter);
  if (customerIn) queryParts.push(customerIn);
  queryParts.push('order=date_out.desc');

  let records = [];
  try {
    const res = supabaseRequest_(SUPABASE_XPPL_DATABASE_ENDPOINT + '?' + queryParts.join('&'));
    records = Array.isArray(res) ? res : [];
  } catch (e) {
    Logger.log('getWeighResultData error: ' + e);
    records = [];
  }

  const rows = records.map(function (record) {
    return mapXpplRecordToRowArray_(record, headers);
  });


  if (!rows.length) {
    var fallbackCustomers = null;
    if (isUser) {
      fallbackCustomers = assignedCustomerNames.length ? assignedCustomerNames : customerFilter;
    }
    return buildEmpty(fallbackCustomers);
  }

  const assignedCustomerLowerSet = assignedCustomerNames.length
    ? new Set(assignedCustomerNames.map(function(v){ return String(v == null ? '' : v).trim().toLowerCase(); }))
    : null;
  const requestedCustomerLowerSet = customerFilter.length
    ? new Set(customerFilter.map(function(v){ return String(v == null ? '' : v).trim().toLowerCase(); }))
    : null;

  let accessibleCustomerLowerSet = null;
  let accessibleCustomerNamesForOptions = [];
  if (isUser) {
    if (assignedCustomerLowerSet && assignedCustomerLowerSet.size) {
      accessibleCustomerLowerSet = assignedCustomerLowerSet;
      accessibleCustomerNamesForOptions = assignedCustomerNames.slice();
    } else if (requestedCustomerLowerSet && requestedCustomerLowerSet.size) {
      accessibleCustomerLowerSet = requestedCustomerLowerSet;
      accessibleCustomerNamesForOptions = customerFilter.slice();
    } else {
      return buildEmpty(assignedCustomerNames);
    }
  }

  let filterCustomerLowerSet = null;
  if (isUser) {
    if (assignedCustomerLowerSet && assignedCustomerLowerSet.size) {
      if (requestedCustomerLowerSet && requestedCustomerLowerSet.size) {
        const intersection = [];
        requestedCustomerLowerSet.forEach(function(name){
          if (assignedCustomerLowerSet.has(name)) {
            intersection.push(name);
          }
        });
        filterCustomerLowerSet = new Set(intersection);
        if (filterCustomerLowerSet.size === 0) {
          return buildEmpty(accessibleCustomerNamesForOptions.length ? accessibleCustomerNamesForOptions : assignedCustomerNames);
        }
      } else {
        filterCustomerLowerSet = assignedCustomerLowerSet;
      }
    } else {
      filterCustomerLowerSet = requestedCustomerLowerSet;
    }
  } else {
    filterCustomerLowerSet = requestedCustomerLowerSet;
  }

  const restrictByCustomer = Boolean(accessibleCustomerLowerSet && accessibleCustomerLowerSet.size);

  const contractSet = contractFilter.length
    ? new Set(contractFilter.map(function(v){ return String(v); }))
    : null;
  const customerSet = filterCustomerLowerSet && filterCustomerLowerSet.size
    ? filterCustomerLowerSet
    : null;

  const baseRows = [];
  const optionContracts = new Set();
  const optionCustomers = new Set();
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var dateKey = _toDateKey(row[idxDateOut]);
    if (from && (!dateKey || dateKey < from)) continue;
    if (to && (!dateKey || dateKey > to)) continue;

    var rowContract = String(stripLeadingApostrophe(row[idxContract]) || '').trim();
    var rowCustomer = String(stripLeadingApostrophe(row[idxCustomer]) || '').trim();
    var rowCustomerKey = rowCustomer.toLowerCase();

    if (restrictByCustomer && (!rowCustomerKey || !accessibleCustomerLowerSet.has(rowCustomerKey))) continue;
  
    if (rowContract) optionContracts.add(rowContract);
    if (rowCustomer) optionCustomers.add(rowCustomer);

    if (contractSet && !contractSet.has(rowContract)) continue;
    if (customerSet && (!rowCustomerKey || !customerSet.has(rowCustomerKey))) continue;

    baseRows.push(row);
  }

  var availableContracts = Array.from(optionContracts).sort(function(a, b) {
    return String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: 'base' });
  });
  var availableCustomers = Array.from(optionCustomers).sort(function(a, b) {
    return String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: 'base' });
  });
  if (!availableCustomers.length && accessibleCustomerNamesForOptions.length) {
    availableCustomers = uniqueSortedList(accessibleCustomerNamesForOptions);
  }

  var totalRecords = baseRows.length;
  if (!totalRecords) {
    return {
      draw: draw,
      recordsTotal: 0,
      recordsFiltered: 0,
      data: [],
      counts: { unassigned: 0, unknown: 0, assigned: 0 },
      summary: { trucks: 0, weight: 0 },
      options: { contracts: availableContracts, customers: availableCustomers }
    };
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
  let totalWeight = 0;
  for (var k = 0; k < filteredForSearch.length; k++) {
    var comp = String(stripLeadingApostrophe(filteredForSearch[k][idxCompany]) || '').trim();
    if (!comp) counts.unassigned++;
    else if (comp.toLowerCase() === 'unknown') counts.unknown++;
    else counts.assigned++;

    if (idxNetWeight > -1) {
      var rawWeight = stripLeadingApostrophe(filteredForSearch[k][idxNetWeight]);
      if (typeof rawWeight === 'string') {
        rawWeight = rawWeight.replace(/,/g, '');
      }
      var weightNum = Number(rawWeight);
      if (Number.isFinite(weightNum)) {
        totalWeight += weightNum;
      }
    }
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
    counts: counts,
    summary: { trucks: filteredForSearch.length, weight: totalWeight },
    options: { contracts: availableContracts, customers: availableCustomers }
  };
}

function updateWeighResultCompany(payload, sessionToken) {
  const user = requireAdmin_(sessionToken);
  const { ID, 'Transportation Company': company } = payload || {};
  if (!ID) throw new Error('Thiếu ID.');

  const sanitizedCompany = sanitizeXpplText_(company);
  const userLabel = sanitizeXpplText_(user.username || user.user || user.email || '');
  const response = supabaseRequest_(
    SUPABASE_XPPL_DATABASE_ENDPOINT + '?id=eq.' + encodeURIComponent(ID),
    {
      method: 'PATCH',
      payload: {
        transportation_company: sanitizedCompany,
        username: userLabel
      },
      headers: { Prefer: 'return=representation' }
    }
  );

  if (!Array.isArray(response) || !response.length) {
    throw new Error('Không tìm thấy ID.');
  }

  return 'Đã cập nhật.';
}

function deleteWeighResults(ids, sessionToken) {
  const user = requireAdmin_(sessionToken);
  if (!Array.isArray(ids) || !ids.length) return 'Không có ID.';
  const sanitized = ids
    .map(function(id) { return String(id == null ? '' : id).trim(); })
    .filter(function(id) { return id; });
  if (!sanitized.length) return 'Không có ID.';

  const filter = buildSupabaseInFilter_('id', sanitized);
  if (!filter) return 'Không có ID.';

  supabaseRequest_(SUPABASE_XPPL_DATABASE_ENDPOINT + '?' + filter, {
    method: 'DELETE',
    headers: { Prefer: 'return=minimal' }
  });

  return 'Đã xoá ' + sanitized.length + ' dòng.';
}

/*** END ***/
