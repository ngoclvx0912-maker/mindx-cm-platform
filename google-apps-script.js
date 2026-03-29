// ============================================================
// MindX CM Report Platform — Google Apps Script Web App
// Copy toàn bộ code này vào Google Apps Script, sau đó Deploy
// Phiên bản: 2-tab (ActionPlan + WeeklyReport) + Config
// ============================================================

const SPREADSHEET_ID = '1beULgTt53o_mXun8ImVGoqVATMI2TJkpa2cASsPkoN8';

// Tab names
const TABS = {
  AP: 'ActionPlan',
  WR: 'T3_WeeklyReport',
  FM_NOTES: 'FM_Notes',
  CONFIG: 'Config',
  DAILY: 'DailyReport'
};

// Headers cho từng tab — phải khớp chính xác cột trong Google Sheet
const HEADERS = {
  AP: [
    'bu', 'month', 'week',
    'func', 'chi_so', 'van_de', 'muc_do',
    'root_cause',
    'key_action', 'mo_ta_trien_khai',
    'target_do_luong', 'deadline', 'owner', 'fm_support', 'status',
    'saved_at'
  ],
  WR: [
    'bu', 'month', 'week',
    'kpi', 'target', 'actual', 'notes',
    'saved_at'
  ],
  DAILY: [
    'bu', 'date',
    'lead_mkt', 'calls_n1', 'trial_book', 'trial_done_n1', 'deal_n1', 'revenue_n1',
    'cs_touchpoint', 'phhs_reupsell', 'deal_reupsell', 'deal_referral', 'revenue_n2',
    'direct_sales', 'events', 'lead_n3', 'trial_n3', 'deal_n3', 'revenue_n3',
    'note', 'saved_at'
  ]
};

// Config sheet headers
const CONFIG_HEADERS = ['month', 'section', 'key', 'value', 'saved_at'];

// Danh sách 46 BU — dùng làm fallback nếu chưa có config
const DEFAULT_BU_LIST = [
  "HCM1 - PVT","HCM1 - PXL","HCM1 - TK",
  "HCM2 - LVV","HCM2 - NX","HCM2 - PVD","HCM2 - SH",
  "HCM3 - 3T2","HCM3 - HL","HCM3 - HTLO","HCM3 - PMH","HCM3 - PNL",
  "HCM4 - LBB","HCM4 - TC","HCM4 - TL","HCM4 - TT",
  "HN - HĐT","HN - MK","HN - NCT","HN - NHT","HN - NPS",
  "HN - NVC","HN - OCP","HN - TP","HN - VHHN","HN - VP",
  "K18 - HCM","K18 - HN",
  "MB1 - BN","MB1 - HP","MB1 - QN","MB1 - TS",
  "MB2 - PT","MB2 - TN","MB2 - VP",
  "MN - BD - DA","MN - BD - TA","MN - BD - TDM","MN - BH - PVT",
  "MN - CT - THD","MN - VT - LHP",
  "MT - ĐN","MT - NA","MT - TH",
  "ONL - ART","ONL - COD"
];

// BU_LIST sẽ được đọc từ config, fallback sang DEFAULT_BU_LIST
var BU_LIST = DEFAULT_BU_LIST;

// ============================================================
// CORS HEADERS
// ============================================================
function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type'
  };
}

function makeResponse(data) {
  const output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ============================================================
// doGet — xử lý GET requests
// ============================================================
function doGet(e) {
  try {
    const params = e.parameter || {};
    const action = params.action || '';

    // Load BU_LIST from config if available
    loadBUListFromConfig();

    if (action === 'get') {
      return handleGet(params);
    } else if (action === 'list') {
      return handleList(params);
    } else if (action === 'all_data') {
      return handleAllData(params);
    } else if (action === 'bus') {
      return makeResponse(BU_LIST);
    } else if (action === 'get_fm_note') {
      return handleGetFMNote(params);
    } else if (action === 'save_fm_note') {
      return handleSaveFMNote(params);
    } else if (action === 'mark_fm_note_read') {
      return handleMarkFMNoteRead(params);
    } else if (action === 'delete_fm_note') {
      return handleDeleteFMNote(params);
    } else if (action === 'get_fm_notes_summary') {
      return handleGetFMNotesSummary(params);
    } else if (action === 'get_config') {
      return handleGetConfig(params);
    } else if (action === 'get_config_months') {
      return handleGetConfigMonths();
    } else if (action === 'getDaily') {
      return handleGetDaily(params);
    } else if (action === 'getDailyMTD') {
      return handleGetDailyMTD(params);
    } else if (action === 'getDailyAll') {
      return handleGetDailyAll(params);
    } else if (action === 'qa_list') {
      return handleQAList(params);
    } else if (action === 'qa_upvote') {
      return handleQAUpvote(params);
    } else if (action === 'get_kpi_targets') {
      return handleGetKPITargets(params);
    } else if (action === 'save_kpi_targets_batch') {
      return handleSaveKPIBatch(params);
    } else if (action === 'clear_kpi_targets') {
      return handleClearKPITargets(params);
    } else if (action === 'clear_tab_data') {
      return handleClearTabData(params);
    } else if (action === 'get_staff_list') {
      return handleGetStaffList();
    } else if (action === 'save_row') {
      return handleSaveRow(params);
    } else if (action === 'save_config_via_get') {
      var payloadStr2 = params.payload || '{}';
      try {
        var body2 = JSON.parse(payloadStr2);
        return handleSaveConfig(body2);
      } catch(err) {
        return makeResponse({ error: 'Invalid payload: ' + err.toString() });
      }
    } else {
      return makeResponse({ error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return makeResponse({ error: err.toString() });
  }
}

// ============================================================
// doPost — xử lý POST requests
// ============================================================
function doPost(e) {
  try {
    const params = e.parameter || {};
    const action = params.action || '';

    if (action === 'save') {
      const body = JSON.parse(e.postData.contents);
      return handleSave(params.tab, body);
    } else if (action === 'save_config') {
      const body = JSON.parse(e.postData.contents);
      return handleSaveConfig(body);
    } else if (action === 'saveDaily') {
      const body = JSON.parse(e.postData.contents);
      return handleSaveDaily(body);
    } else if (action === 'qa_post') {
      const body = JSON.parse(e.postData.contents);
      return handleQAPost(body);
    } else if (action === 'mimi_ask') {
      const body = JSON.parse(e.postData.contents);
      return handleMiMiAsk(body);
    } else if (action === 'mimi_direct') {
      const body = JSON.parse(e.postData.contents);
      return handleMiMiDirect(body);
    } else if (action === 'sod_notify') {
      const body = JSON.parse(e.postData.contents);
      return handleSODNotify(body);
    } else if (action === 'save_fm_note') {
      const body = JSON.parse(e.postData.contents);
      return handleSaveFMNote(body);
    } else if (action === 'save_kpi_targets') {
      const body = JSON.parse(e.postData.contents);
      return handleSaveKPITargets(body);
    } else {
      return makeResponse({ error: 'Unknown POST action: ' + action });
    }
  } catch (err) {
    return makeResponse({ error: err.toString() });
  }
}

// ============================================================
// CONFIG: Load BU List from Config sheet
// ============================================================
function loadBUListFromConfig() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = getConfigSheet(ss);
    var data = getSheetData(sheet, CONFIG_HEADERS);

    // Tìm BU list từ config (section=bu_list, lấy config gần nhất)
    var buRows = data.filter(function(row) {
      return row.section === 'bu_list' && row.key === 'list';
    });

    if (buRows.length > 0) {
      // Lấy bản ghi mới nhất
      var latest = buRows[buRows.length - 1];
      var parsed = JSON.parse(latest.value);
      if (Array.isArray(parsed) && parsed.length > 0) {
        BU_LIST = parsed;
      }
    }
  } catch (e) {
    // Fallback sang DEFAULT_BU_LIST
    BU_LIST = DEFAULT_BU_LIST;
  }
}

// ============================================================
// CONFIG: Get or create Config sheet
// ============================================================
function getConfigSheet(ss) {
  var sheet = ss.getSheetByName(TABS.CONFIG);
  if (!sheet) {
    sheet = ss.insertSheet(TABS.CONFIG);
    sheet.getRange(1, 1, 1, CONFIG_HEADERS.length).setValues([CONFIG_HEADERS]);
  }
  return sheet;
}

// ============================================================
// DAILY REPORT: Get or create DailyReport sheet
// ============================================================
function getDailySheet(ss) {
  var sheet = ss.getSheetByName(TABS.DAILY);
  if (!sheet) {
    sheet = ss.insertSheet(TABS.DAILY);
    sheet.getRange(1, 1, 1, HEADERS.DAILY.length).setValues([HEADERS.DAILY]);
  }
  return sheet;
}

// ============================================================
// GET: action=getDaily&bu=X&date=Y
// Returns daily report rows for specific BU + date
// ============================================================
function handleGetDaily(params) {
  var bu = params.bu;
  var date = params.date;
  if (!bu || !date) {
    return makeResponse({ ok: false, error: 'Missing bu or date' });
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getDailySheet(ss);
  var data = getSheetData(sheet, HEADERS.DAILY);

  var filtered = data.filter(function(row) {
    return String(row.bu) === String(bu) && String(row.date) === String(date);
  });

  return makeResponse({ ok: true, data: filtered });
}

// ============================================================
// GET: action=getDailyMTD&bu=X&month=Y (YYYY-MM)
// Returns all daily rows for BU in that month (for MTD summary)
// ============================================================
function handleGetDailyMTD(params) {
  var bu = params.bu;
  var month = params.month; // YYYY-MM
  if (!bu || !month) {
    return makeResponse({ ok: false, error: 'Missing bu or month' });
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getDailySheet(ss);
  var data = getSheetData(sheet, HEADERS.DAILY);

  var filtered = data.filter(function(row) {
    return String(row.bu) === String(bu) && String(row.date).substring(0, 7) === String(month);
  });

  return makeResponse({ ok: true, data: filtered });
}

// ============================================================
// GET: action=getDailyAll&month=YYYY-MM
// Returns ALL daily rows for ALL BUs in that month (for SOD dashboard)
// ============================================================
function handleGetDailyAll(params) {
  var month = params.month; // YYYY-MM
  if (!month) {
    return makeResponse({ ok: false, error: 'Missing month' });
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getDailySheet(ss);
  var data = getSheetData(sheet, HEADERS.DAILY);

  var filtered = data.filter(function(row) {
    return String(row.date).substring(0, 7) === String(month);
  });

  return makeResponse({ ok: true, data: filtered });
}

// ============================================================
// POST: action=saveDaily — UPSERT daily report (delete+insert)
// ============================================================
function handleSaveDaily(body) {
  if (!body || !body.rows || !Array.isArray(body.rows) || body.rows.length === 0) {
    return makeResponse({ ok: false, error: 'Missing rows in body' });
  }

  var row = body.rows[0];
  var bu = row.bu;
  var date = row.date;
  if (!bu || !date) {
    return makeResponse({ ok: false, error: 'Missing bu or date in row' });
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getDailySheet(ss);
  var headers = HEADERS.DAILY;
  var savedAt = new Date().toISOString();

  // Delete existing rows for this BU + date (UPSERT)
  deleteDailyMatchingRows(sheet, headers, bu, date);

  // Insert new row
  var newRow = headers.map(function(h) {
    if (h === 'saved_at') return savedAt;
    var val = row[h];
    return (val === undefined || val === null) ? '' : String(val);
  });

  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, headers.length).setValues([newRow]);

  return makeResponse({ ok: true, saved_at: savedAt });
}

// ============================================================
// HELPER: Delete daily rows matching bu + date
// ============================================================
function deleteDailyMatchingRows(sheet, headers, bu, date) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var buIdx = headers.indexOf('bu');
  var dateIdx = headers.indexOf('date');
  if (buIdx < 0 || dateIdx < 0) return;

  var data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var rowsToDelete = [];

  data.forEach(function(row, i) {
    var rowBu = cellToString(row[buIdx], 'bu');
    var rowDate = cellToString(row[dateIdx], 'date');
    if (rowBu === String(bu) && rowDate === String(date)) {
      rowsToDelete.push(i + 2);
    }
  });

  // Delete from bottom to top to avoid index shifting
  rowsToDelete.reverse().forEach(function(rowNum) {
    sheet.deleteRow(rowNum);
  });
}

// ============================================================
// GET: action=get_config&month=Y
// Trả về config cho tháng cụ thể
// ============================================================
function handleGetConfig(params) {
  var month = params.month;
  if (!month) {
    return makeResponse({ error: 'Missing param: month' });
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getConfigSheet(ss);
  var data = getSheetData(sheet, CONFIG_HEADERS);

  // Filter theo month
  var monthData = data.filter(function(row) {
    return String(row.month) === String(month);
  });

  // Build config object grouped by section
  var config = {};
  monthData.forEach(function(row) {
    if (!config[row.section]) {
      config[row.section] = {};
    }
    // Parse JSON values
    try {
      config[row.section][row.key] = JSON.parse(row.value);
    } catch (e) {
      config[row.section][row.key] = row.value;
    }
  });

  return makeResponse({ month: month, config: config });
}

// ============================================================
// GET: action=get_config_months
// Trả về danh sách các tháng đã có config
// ============================================================
function handleGetConfigMonths() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getConfigSheet(ss);
  var data = getSheetData(sheet, CONFIG_HEADERS);

  var months = {};
  data.forEach(function(row) {
    if (row.month) months[row.month] = true;
  });

  var sortedMonths = Object.keys(months).sort();
  return makeResponse({ months: sortedMonths });
}

// ============================================================
// POST: action=save_config
// Body: { month, config: { section: { key: value, ... }, ... } }
// ============================================================
function handleSaveConfig(body) {
  if (!body || !body.month || !body.config) {
    return makeResponse({ success: false, error: 'Missing month or config in body' });
  }

  var month = body.month;
  var config = body.config;
  var savedAt = new Date().toISOString();

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getConfigSheet(ss);

  // Xóa config cũ của tháng này
  deleteConfigMonth(sheet, month);

  // Ghi config mới
  var newRows = [];
  for (var section in config) {
    if (!config.hasOwnProperty(section)) continue;
    var sectionData = config[section];
    for (var key in sectionData) {
      if (!sectionData.hasOwnProperty(key)) continue;
      var value = sectionData[key];
      // Stringify objects/arrays
      var valueStr = (typeof value === 'object') ? JSON.stringify(value) : String(value);
      newRows.push([month, section, key, valueStr, savedAt]);
    }
  }

  if (newRows.length > 0) {
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, newRows.length, CONFIG_HEADERS.length).setValues(newRows);
  }

  // Cập nhật BU_LIST nếu có thay đổi
  if (config.bu_list && config.bu_list.list) {
    BU_LIST = config.bu_list.list;
  }

  return makeResponse({ success: true, saved_at: savedAt, rows_saved: newRows.length });
}

// ============================================================
// HELPER: Xóa tất cả config rows của 1 tháng
// ============================================================
function deleteConfigMonth(sheet, month) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var monthIdx = CONFIG_HEADERS.indexOf('month');
  var data = sheet.getRange(2, 1, lastRow - 1, CONFIG_HEADERS.length).getValues();
  var toDelete = [];

  data.forEach(function(row, i) {
    if (cellToString(row[monthIdx], 'month') === String(month)) {
      toDelete.push(i + 2);
    }
  });

  toDelete.reverse().forEach(function(rowNum) {
    sheet.deleteRow(rowNum);
  });
}

// ============================================================
// GET handler: action=get&tab=AP|WR&bu=X&month=Y&week=Z
// ============================================================
function handleGet(params) {
  const tab = params.tab;
  const bu = params.bu;
  const month = params.month;
  const week = String(params.week);

  if (!tab || !bu || !month || !week) {
    return makeResponse({ error: 'Missing required params: tab, bu, month, week' });
  }

  if (!TABS[tab]) {
    return makeResponse({ error: 'Invalid tab: ' + tab + '. Must be AP or WR' });
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(TABS[tab]);
  if (!sheet) {
    return makeResponse({ error: 'Sheet not found: ' + TABS[tab] });
  }

  const headers = HEADERS[tab];
  const rows = getMatchingRows(sheet, headers, bu, month, week);

  let savedAt = null;
  const cleanRows = rows.map(row => {
    const r = Object.assign({}, row);
    if (r.saved_at) savedAt = r.saved_at;
    delete r.saved_at;
    delete r.bu;
    delete r.month;
    delete r.week;
    return r;
  });

  return makeResponse({ rows: cleanRows, saved_at: savedAt });
}

// ============================================================
// GET handler: action=list&month=Y&week=Z
// ============================================================
function handleList(params) {
  const month = params.month;
  const week = String(params.week);

  if (!month || !week) {
    return makeResponse({ error: 'Missing required params: month, week' });
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const result = {};

  BU_LIST.forEach(bu => {
    result[bu] = { AP: false, WR: false };
  });

  ['AP', 'WR'].forEach(tab => {
    const sheet = ss.getSheetByName(TABS[tab]);
    if (!sheet) return;
    const headers = HEADERS[tab];
    const allData = getSheetData(sheet, headers);
    allData.forEach(row => {
      if (String(row.month) === String(month) && String(row.week) === String(week)) {
        if (result[row.bu] !== undefined) {
          result[row.bu][tab] = true;
        }
      }
    });
  });

  return makeResponse(result);
}

// ============================================================
// GET handler: action=all_data&month=Y&week=Z
// ============================================================
function handleAllData(params) {
  const month = params.month;
  const week = String(params.week);

  if (!month || !week) {
    return makeResponse({ error: 'Missing required params: month, week' });
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const result = {};

  BU_LIST.forEach(bu => {
    result[bu] = { AP: [], WR: [] };
  });

  ['AP', 'WR'].forEach(tab => {
    const sheet = ss.getSheetByName(TABS[tab]);
    if (!sheet) return;
    const headers = HEADERS[tab];
    const allData = getSheetData(sheet, headers);

    allData.forEach(row => {
      if (String(row.month) === String(month) && String(row.week) === String(week)) {
        const bu = row.bu;
        if (result[bu]) {
          const cleanRow = Object.assign({}, row);
          delete cleanRow.bu;
          delete cleanRow.month;
          delete cleanRow.week;
          delete cleanRow.saved_at;
          result[bu][tab].push(cleanRow);
        }
      }
    });
  });

  return makeResponse(result);
}

// ============================================================
// POST handler: action=save&tab=AP|WR
// ============================================================
function handleSave(tab, body) {
  if (!tab || !body) {
    return makeResponse({ success: false, error: 'Missing tab or body' });
  }

  if (!TABS[tab]) {
    return makeResponse({ success: false, error: 'Invalid tab: ' + tab + '. Must be AP or WR' });
  }

  const { bu, month, week, rows } = body;
  if (!bu || !month || !week) {
    return makeResponse({ success: false, error: 'Missing bu, month, or week in body' });
  }
  if (!rows || !Array.isArray(rows)) {
    return makeResponse({ success: false, error: 'rows must be an array' });
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(TABS[tab]);
  if (!sheet) {
    return makeResponse({ success: false, error: 'Sheet not found: ' + TABS[tab] });
  }

  const headers = HEADERS[tab];
  const savedAt = new Date().toISOString();

  deleteMatchingRows(sheet, headers, bu, String(month), String(week));

  if (rows.length > 0) {
    const newRows = rows.map(row => {
      return headers.map(h => {
        if (h === 'bu') return bu;
        if (h === 'month') return month;
        if (h === 'week') return String(week);
        if (h === 'saved_at') return savedAt;
        const val = row[h];
        return (val === undefined || val === null) ? '' : String(val);
      });
    });

    const lastRow = sheet.getLastRow();
    const startRow = lastRow + 1;
    if (newRows.length > 0) {
      sheet.getRange(startRow, 1, newRows.length, headers.length).setValues(newRows);
    }
  }

  return makeResponse({ success: true, saved_at: savedAt, rows_saved: rows.length });
}

// ============================================================
// HELPER: Chuyển cell value sang string, xử lý Date objects
// ============================================================
function cellToString(value, header) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date) {
    if (header === 'month') {
      const y = value.getFullYear();
      const m = String(value.getMonth() + 1).padStart(2, '0');
      return y + '-' + m;
    }
    if (header === 'date') {
      const y = value.getFullYear();
      const m = String(value.getMonth() + 1).padStart(2, '0');
      const d = String(value.getDate()).padStart(2, '0');
      return y + '-' + m + '-' + d;
    }
    if (header === 'deadline' || header === 'saved_at') {
      return value.toISOString();
    }
    return value.toISOString();
  }
  if (header === 'month' && typeof value === 'number' && value > 40000 && value < 60000) {
    var d1 = new Date(Date.UTC(1899, 11, 30 + value));
    var y1 = d1.getUTCFullYear();
    var m1 = String(d1.getUTCMonth() + 1).padStart(2, '0');
    return y1 + '-' + m1;
  }
  if (header === 'date' && typeof value === 'number' && value > 40000 && value < 60000) {
    var d2 = new Date(Date.UTC(1899, 11, 30 + value));
    var y2 = d2.getUTCFullYear();
    var m2 = String(d2.getUTCMonth() + 1).padStart(2, '0');
    var dd = String(d2.getUTCDate()).padStart(2, '0');
    return y2 + '-' + m2 + '-' + dd;
  }
  return String(value);
}

// ============================================================
// HELPER: Đọc toàn bộ dữ liệu từ sheet thành mảng objects
// ============================================================
function getSheetData(sheet, headers) {
  const lastRow = sheet.getLastRow();
  const lastCol = headers.length;

  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return data
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        obj[h] = row[i] !== undefined ? cellToString(row[i], h) : '';
      });
      return obj;
    });
}

// ============================================================
// HELPER: Lấy dòng khớp bu+month+week
// ============================================================
function getMatchingRows(sheet, headers, bu, month, week) {
  const allData = getSheetData(sheet, headers);
  return allData.filter(row =>
    String(row.bu) === String(bu) &&
    String(row.month) === String(month) &&
    String(row.week) === String(week)
  );
}

// ============================================================
// HELPER: Xóa dòng khớp bu+month+week
// ============================================================
function deleteMatchingRows(sheet, headers, bu, month, week) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const buIdx = headers.indexOf('bu');
  const monthIdx = headers.indexOf('month');
  const weekIdx = headers.indexOf('week');

  if (buIdx < 0 || monthIdx < 0 || weekIdx < 0) return;

  const data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  const rowsToDelete = [];

  data.forEach((row, i) => {
    if (
      cellToString(row[buIdx], 'bu') === String(bu) &&
      cellToString(row[monthIdx], 'month') === String(month) &&
      cellToString(row[weekIdx], 'week') === String(week)
    ) {
      rowsToDelete.push(i + 2);
    }
  });

  rowsToDelete.reverse().forEach(rowNum => {
    sheet.deleteRow(rowNum);
  });
}

// ============================================================
// SETUP: Chạy hàm này 1 lần để tạo headers trong các sheet
// ============================================================
function setupSheetHeaders() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  ['AP', 'WR'].forEach(tab => {
    const sheet = ss.getSheetByName(TABS[tab]);
    if (!sheet) {
      Logger.log('Sheet not found: ' + TABS[tab]);
      return;
    }
    const headers = HEADERS[tab];
    const firstCell = sheet.getRange(1, 1).getValue();
    if (!firstCell) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      Logger.log('Headers đã được set cho: ' + TABS[tab]);
    } else {
      Logger.log('Headers đã tồn tại cho: ' + TABS[tab]);
    }
  });

  // Setup Config sheet
  getConfigSheet(ss);
  Logger.log('Config sheet sẵn sàng');
}

// ============================================================
// FM NOTES
// ============================================================
// ============================================================
// FM NOTES — Thread-based (v2)
// Columns: note_id, bu, month, week, fm_name, note, saved_at, read_by_cm
// ============================================================
const FM_NOTES_HEADERS = ['note_id', 'bu', 'month', 'week', 'fm_name', 'note', 'saved_at', 'read_by_cm'];

function getFMNotesSheet(ss) {
  var sheet = ss.getSheetByName(TABS.FM_NOTES);
  if (!sheet) {
    sheet = ss.insertSheet(TABS.FM_NOTES);
    sheet.getRange(1, 1, 1, FM_NOTES_HEADERS.length).setValues([FM_NOTES_HEADERS]);
  }
  return sheet;
}

function generateNoteId() {
  return 'n_' + new Date().getTime() + '_' + Math.random().toString(36).substr(2, 5);
}

/** get_fm_notes — trả về TOÀN BỘ notes cho 1 BU/month/week (thread) */
function handleGetFMNote(params) {
  var bu = params.bu;
  var month = params.month;
  var week = String(params.week);

  if (!bu || !month || !week) {
    return makeResponse({ error: 'Missing params: bu, month, week' });
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getFMNotesSheet(ss);
  var data = getSheetData(sheet, FM_NOTES_HEADERS);

  var found = data.filter(function(row) {
    return String(row.bu) === String(bu) &&
           String(row.month) === String(month) &&
           String(row.week) === String(week);
  });

  // Sort by saved_at ASC (oldest first)
  found.sort(function(a, b) {
    return String(a.saved_at || '').localeCompare(String(b.saved_at || ''));
  });

  var notes = found.map(function(row) {
    return {
      note_id: row.note_id || '',
      fm_name: row.fm_name || '',
      note: row.note || '',
      saved_at: row.saved_at || '',
      read_by_cm: row.read_by_cm === 'true' || row.read_by_cm === true
    };
  });

  // Backward compat: also return latest as flat fields
  var latest = notes.length ? notes[notes.length - 1] : null;
  return makeResponse({
    notes: notes,
    total: notes.length,
    unread: notes.filter(function(n) { return !n.read_by_cm; }).length,
    note: latest ? latest.note : null,
    saved_at: latest ? latest.saved_at : null,
    read_by_cm: latest ? latest.read_by_cm : false
  });
}

/** save_fm_note — THÊM note mới (không xóa cũ) */
function handleSaveFMNote(params) {
  var bu = params.bu;
  var month = params.month;
  var week = String(params.week);
  var note = params.note || '';
  var fmName = params.fm_name || '';

  if (!bu || !month || !week || !note) {
    return makeResponse({ success: false, error: 'Missing params: bu, month, week, note' });
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getFMNotesSheet(ss);

  var noteId = generateNoteId();
  var savedAt = new Date().toISOString();
  var newRow = [noteId, bu, month, week, fmName, note, savedAt, 'false'];
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, FM_NOTES_HEADERS.length).setValues([newRow]);

  return makeResponse({ success: true, note_id: noteId, saved_at: savedAt });
}

/** mark_fm_note_read — đánh dấu 1 note hoặc tất cả notes đã đọc */
function handleMarkFMNoteRead(params) {
  var bu = params.bu;
  var month = params.month;
  var week = String(params.week);
  var noteId = params.note_id || ''; // nếu trống → mark ALL

  if (!bu || !month || !week) {
    return makeResponse({ success: false, error: 'Missing params: bu, month, week' });
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getFMNotesSheet(ss);

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return makeResponse({ success: false, error: 'No note found' });
  }

  var data = sheet.getRange(2, 1, lastRow - 1, FM_NOTES_HEADERS.length).getValues();
  var noteIdIdx = FM_NOTES_HEADERS.indexOf('note_id');
  var buIdx = FM_NOTES_HEADERS.indexOf('bu');
  var monthIdx = FM_NOTES_HEADERS.indexOf('month');
  var weekIdx = FM_NOTES_HEADERS.indexOf('week');
  var readIdx = FM_NOTES_HEADERS.indexOf('read_by_cm');

  var updated = 0;
  data.forEach(function(row, i) {
    if (
      cellToString(row[buIdx], 'bu') === String(bu) &&
      cellToString(row[monthIdx], 'month') === String(month) &&
      cellToString(row[weekIdx], 'week') === String(week)
    ) {
      // Nếu có note_id cụ thể thì chỉ mark note đó, ngược lại mark all
      if (!noteId || String(row[noteIdIdx]) === String(noteId)) {
        if (String(row[readIdx]) !== 'true') {
          sheet.getRange(i + 2, readIdx + 1).setValue('true');
          updated++;
        }
      }
    }
  });

  if (updated === 0 && noteId) {
    return makeResponse({ success: false, error: 'Note not found or already read' });
  }

  return makeResponse({ success: true, marked: updated });
}

/** delete_fm_note — xóa 1 note cụ thể (FM tự xóa note mình viết) */
function handleDeleteFMNote(params) {
  var noteId = params.note_id;
  if (!noteId) {
    return makeResponse({ success: false, error: 'Missing note_id' });
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getFMNotesSheet(ss);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return makeResponse({ success: false, error: 'No notes' });

  var data = sheet.getRange(2, 1, lastRow - 1, FM_NOTES_HEADERS.length).getValues();
  var noteIdIdx = FM_NOTES_HEADERS.indexOf('note_id');

  for (var i = data.length - 1; i >= 0; i--) {
    if (String(data[i][noteIdIdx]) === String(noteId)) {
      sheet.deleteRow(i + 2);
      return makeResponse({ success: true });
    }
  }
  return makeResponse({ success: false, error: 'Note not found' });
}

/** get_fm_notes_summary — trả về tổng hợp notes cho NHIỀU BUs (batch, dùng cho FM table icon) */
function handleGetFMNotesSummary(params) {
  var month = params.month;
  var week = String(params.week);

  if (!month || !week) {
    return makeResponse({ error: 'Missing params: month, week' });
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getFMNotesSheet(ss);
  var data = getSheetData(sheet, FM_NOTES_HEADERS);

  var summary = {}; // { bu: { total, unread, latest_at } }

  data.forEach(function(row) {
    if (String(row.month) === String(month) && String(row.week) === String(week)) {
      var bu = String(row.bu);
      if (!summary[bu]) {
        summary[bu] = { total: 0, unread: 0, latest_at: '' };
      }
      summary[bu].total++;
      if (row.read_by_cm !== 'true' && row.read_by_cm !== true) {
        summary[bu].unread++;
      }
      if (String(row.saved_at || '') > summary[bu].latest_at) {
        summary[bu].latest_at = row.saved_at || '';
      }
    }
  });

  return makeResponse(summary);
}

function setupFMNotesSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getFMNotesSheet(ss);
  Logger.log('FM_Notes sheet sẵn sàng: ' + sheet.getName());
}

// ============================================================
// Q&A DISCUSSION
// Tab: QA_Discussion
// Headers: id, type, parent_id, bu, content, upvotes, created_at
// ============================================================
const QA_HEADERS = ['id', 'type', 'parent_id', 'bu', 'content', 'upvotes', 'created_at'];
const QA_TAB = 'QA_Discussion';

function getQASheet(ss) {
  var sheet = ss.getSheetByName(QA_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(QA_TAB);
    sheet.getRange(1, 1, 1, QA_HEADERS.length).setValues([QA_HEADERS]);
  }
  return sheet;
}

function generateQAId() {
  return 'qa_' + new Date().getTime() + '_' + Math.random().toString(36).substr(2, 6);
}

// GET: action=qa_list&page=1&limit=20&sort=newest|hot|most_replies&search=keyword
// Trả về danh sách câu hỏi + answers
function handleQAList(params) {
  var page = parseInt(params.page || '1', 10);
  var limit = parseInt(params.limit || '20', 10);
  var isSod = params.is_sod === 'true';
  var sortBy = params.sort || 'newest'; // newest | hot | most_replies
  var searchTerm = (params.search || '').toLowerCase().trim();

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getQASheet(ss);
  var data = getSheetData(sheet, QA_HEADERS);

  // Tách questions và answers
  var questions = data.filter(function(r) { return r.type === 'question'; });
  var answers = data.filter(function(r) { return r.type !== 'question'; });

  // Pre-compute answer_count cho mỗi question
  questions.forEach(function(q) {
    q._answer_count = answers.filter(function(a) { return a.parent_id === q.id; }).length;
    q._upvotes = parseInt(q.upvotes || '0', 10);
  });

  // Tìm kiếm theo từ khóa
  if (searchTerm) {
    questions = questions.filter(function(q) {
      return (q.content || '').toLowerCase().indexOf(searchTerm) >= 0;
    });
  }

  // Sắp xếp theo tiêu chí
  if (sortBy === 'hot') {
    // Hot = nhiều upvote nhất
    questions.sort(function(a, b) {
      if (b._upvotes !== a._upvotes) return b._upvotes - a._upvotes;
      return String(b.created_at || '').localeCompare(String(a.created_at || ''));
    });
  } else if (sortBy === 'most_replies') {
    // Nhiều bình luận nhất
    questions.sort(function(a, b) {
      if (b._answer_count !== a._answer_count) return b._answer_count - a._answer_count;
      return String(b.created_at || '').localeCompare(String(a.created_at || ''));
    });
  } else {
    // Mới nhất (default)
    questions.sort(function(a, b) {
      return String(b.created_at || '').localeCompare(String(a.created_at || ''));
    });
  }

  // Phân trang
  var total = questions.length;
  var start = (page - 1) * limit;
  var paged = questions.slice(start, start + limit);

  // Gắn answers vào từng question
  var result = paged.map(function(q) {
    var qAnswers = answers.filter(function(a) { return a.parent_id === q.id; });
    qAnswers.sort(function(a, b) {
      return String(a.created_at || '').localeCompare(String(b.created_at || ''));
    });

    return {
      id: q.id,
      type: q.type,
      bu: q.bu || '',  // Gửi BU để frontend tạo animal alias
      content: q.content,
      upvotes: q._upvotes,
      created_at: q.created_at,
      answer_count: qAnswers.length,
      answers: qAnswers.map(function(a) {
        return {
          id: a.id,
          type: a.type,
          bu: a.bu || '',
          content: a.content,
          created_at: a.created_at
        };
      })
    };
  });

  return makeResponse({
    questions: result,
    total: total,
    page: page,
    limit: limit,
    has_more: (start + limit) < total
  });
}

// GET: action=qa_upvote&id=X
// Tăng upvote cho câu hỏi
function handleQAUpvote(params) {
  var id = params.id;
  if (!id) {
    return makeResponse({ success: false, error: 'Missing param: id' });
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getQASheet(ss);

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return makeResponse({ success: false, error: 'No data' });
  }

  var data = sheet.getRange(2, 1, lastRow - 1, QA_HEADERS.length).getValues();
  var idIdx = QA_HEADERS.indexOf('id');
  var upvotesIdx = QA_HEADERS.indexOf('upvotes');

  for (var i = 0; i < data.length; i++) {
    if (String(data[i][idIdx]) === String(id)) {
      var current = parseInt(data[i][upvotesIdx] || '0', 10);
      var newVal = current + 1;
      sheet.getRange(i + 2, upvotesIdx + 1).setValue(newVal);
      return makeResponse({ success: true, upvotes: newVal });
    }
  }

  return makeResponse({ success: false, error: 'ID not found' });
}

// POST: action=qa_post
// Body: { type, parent_id, bu, content }
function handleQAPost(body) {
  if (!body || !body.type || !body.content) {
    return makeResponse({ success: false, error: 'Missing type or content' });
  }

  var type = body.type; // question | answer | ai_answer | sod_answer
  var parentId = body.parent_id || '';
  var bu = body.bu || '';
  var content = body.content;

  // Validate type
  var validTypes = ['question', 'answer', 'ai_answer', 'sod_answer'];
  if (validTypes.indexOf(type) < 0) {
    return makeResponse({ success: false, error: 'Invalid type: ' + type });
  }

  // BU là nice to have, không bắt buộc — giữ CM thoải mái ẩn danh

  // Answers phải có parent_id
  if (type !== 'question' && !parentId) {
    return makeResponse({ success: false, error: 'Answer must include parent_id' });
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getQASheet(ss);

  var id = generateQAId();
  // Cho phép truyền created_at tùy chỉnh (dùng cho seeding)
  var createdAt = body.created_at || new Date().toISOString();

  var newRow = [
    id,
    type,
    parentId,
    bu,
    content,
    type === 'question' ? 0 : 0, // upvotes khởi đầu = 0
    createdAt
  ];

  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, QA_HEADERS.length).setValues([newRow]);

  return makeResponse({ success: true, id: id, created_at: createdAt });
}

// Setup helper để tạo QA sheet
function setupQASheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getQASheet(ss);
  Logger.log('QA_Discussion sheet sẵn sàng: ' + sheet.getName());
}

// ============================================================
// MiMi AI — Gọi Perplexity API trả lời câu hỏi CM
// ============================================================
var PERPLEXITY_API_KEY = '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';

var INTERNAL_KEYWORDS = ['lương', 'thưởng', 'kpi cá nhân', 'đánh giá nhân sự', 'sa thải', 'thôi việc', 'ngân sách', 'số liệu tài chính'];

// ============================================================
// MiMi System Prompt — Kiến thức nghiệp vụ MindX chuyên sâu
// ============================================================
function buildMiMiSystemPrompt(internalData) {
  var base = [
    'Bạn là MiMi — trợ lý AI chuyên gia về vận hành và kinh doanh của MindX Technology School.',
    'Nhiệm vụ chính: hỗ trợ các Center Manager (CM) giải quyết vấn đề vận hành trung tâm dựa trên DỮ LIỆU THỰC của họ.',
    '',
    '=== TỔNG QUAN MindX ===',
    'MindX Technology School: hệ thống trường công nghệ cho K12 (6-18 tuổi) và 18+, 44-46 trung tâm (BU) toàn quốc.',
    'Môn học: Coding (Python, Web, App), Robotics, AI, Game Design, Digital Art.',
    '',
    '=== CẤU TRÚC TỔ CHỨC ===',
    '- SOD (Sales Director): Giám đốc kinh doanh, quản lý toàn hệ thống',
    '- FM (Field Manager): Quản lý vùng, phụ trách nhóm BU',
    '- CM (Center Manager): Quản lý trung tâm, báo cáo lên FM/SOD',
    '- Sale: Nhân viên kinh doanh tại BU',
    '- CS (Customer Service): Chăm sóc học viên',
    '',
    '=== OKRs 2026 ===',
    '[N1 - OPTIMIZE] Target 154B — Tối ưu Leads MKT:',
    '  CR12 (KH quan tâm / Lead MKT)',
    '  CR23 (Đặt hẹn / KH quan tâm)',
    '  CR34 (Lên trial / Đặt hẹn)',
    '  CR46 >= 50% (Chốt deal / Trial)',
    '  CR16 >= 15% (Deal / Lead tổng)',
    '  AOV >= 18M (Giá trị đơn hàng trung bình)',
    '',
    '[N2 - OPS] Target 174B — Re-enroll/Upsell/Referral:',
    '  Retention > 60%',
    '  Upsell >= 25%',
    '  Referral >= 8%',
    '',
    '[N3 - GROWTH] Target 100B — Sale tự kiếm:',
    '  Direct Sales, Local Events, Partnership B2B',
    '',
    '=== THUẬT NGỮ QUAN TRỌNG ===',
    'L1 = Lead (MKT), L2 = KH quan tâm, L3 = Đặt hẹn, L4 = Trial, L6 = Deal (chốt)',
    'CR12 = L2/L1, CR23 = L3/L2, CR34 = L4/L3, CR46 = L6/L4, CR16 = L6/L1',
    'AOV = Average Order Value (đơn giá trung bình)',
    'Pace = tiến độ doanh số so với target theo tuần: W1=15%, W2=40%, W3=65%, W4=100%',
    'BU Xanh = BU đạt/vượt pace, BU Đỏ = BU dưới pace',
    '',
    '=== QUY TRÌNH BÁO CÁO ===',
    'CM gửi báo cáo (Weekly Report + Action Plan) vào Chủ nhật và Thứ 2.',
    'Khóa chỉnh sửa sau Thứ 4. Tuần: W1=1-7, W2=8-14, W3=15-21, W4=22+.',
    '',
    '=== ACTION PLAN (AP) ===',
    'Mỗi AP gồm: Func (N1/N2/N3), Chỉ số, Vấn đề, Mức độ, Root Cause, Key Action, Mô tả triển khai, Target đo lường, Deadline, Owner, FM Support, Status.',
    'Chất lượng AP tốt: Root Cause cụ thể (không chung chung), Key Action SMART, có target đo lường rõ ràng.',
    '',
    '=== WEEKLY REPORT (WR) ===',
    'Gồm các KPI: Doanh số N1/N2/N3, số lượng Leads, Trials, Deals, CR các loại, AOV...',
    'Mỗi dòng gồm: KPI, Target, Actual, Notes.',
    '',
    '=== CÁCH PHÂN TÍCH KHI CM HỎI ===',
    '1. Nếu có dữ liệu WR/AP của BU: PHẢI dựa vào dữ liệu thực để trả lời, so benchmark',
    '2. Chỉ ra cụ thể: chỉ số nào đạt, chỉ số nào chưa đạt, gap bao nhiêu',
    '3. Đưa gợi ý hành động cụ thể (không chung chung kiểu "cần cải thiện")',
    '4. Nếu không có dữ liệu: trả lời dựa trên kiến thức chung + khuyên CM nhập dữ liệu để được tư vấn chính xác hơn',
    '',
    '=== BEST PRACTICES (BU Xanh thường làm) ===',
    '- Calls/ngày: 15-20 cuộc gọi outbound',
    '- Follow-up trial trong 24h',
    '- Upsell tại thời điểm re-enroll (cuối khóa)',
    '- Local events: 2-3 events/tháng',
    '- Referral program: hỏi phụ huynh ngay sau feedback tốt',
    '- Trial quality: chuẩn bị giáo trình thử theo độ tuổi, gọi confirm trước 1 ngày',
    '- CS proactive: gọi chăm sóc sau buổi 3, buổi 8',
    '',
    '=== QUY TẮC TRẢ LỜI ===',
    '1. Luôn tiếng Việt, thân thiện, chuyên nghiệp',
    '2. Đi thẳng vào vấn đề, tối đa 3-4 đoạn',
    '3. Nếu có dữ liệu BU: trích dẫn số liệu cụ thể ("CR46 của bạn là 35%, cần đạt 50%")',
    '4. Luôn kết thúc bằng 1-2 hành động cụ thể CM có thể làm ngay',
    '5. Nếu câu hỏi quá nhạy cảm (lương, thưởng, nhân sự): trả lời "Câu hỏi này cần HO trả lời trực tiếp. Nhấn Hỏi HO nhé!"',
    '6. Không bịa số liệu. Nếu không có dữ liệu, nói rõ ràng.',
    '7. Dùng emoji phù hợp để sinh động. Luôn khuyến khích CM.',
    '8. TUYỆT ĐỐI KHÔNG dùng markdown table (|---|). Thay vào đó, trình bày bằng:',
    '   - Dùng số thứ tự: 1. 2. 3.',
    '   - Dùng gạch đầu dòng: -',
    '   - Dùng in đậm **text** cho tiêu đề',
    '   - Ví dụ đúng: "1. Cải thiện Trial Quality - Gọi confirm trước 1 ngày - Kỳ vọng: CR46 tăng 2-3%"',
    '   - KHÔNG dùng cú pháp bảng | cột 1 | cột 2 | vì giao diện không hỗ trợ render bảng'
  ].join('\n');

  // === PHẦN NGHIÊN CỨU SÂU ===
  base += '\n\n=== PHƯƠNG PHÁP TRẢ LỜI ===';
  base += '\nBạn là chuyên gia tư vấn, KHÔNG phải chatbot đơn giản. Mỗi câu trả lời phải:';
  base += '\n1. Dựa trên dữ liệu nội bộ MindX (nếu có) để hiểu bối cảnh cụ thể';
  base += '\n2. Kết hợp nghiên cứu thị trường giáo dục công nghệ, xu hướng EdTech Việt Nam/Đông Nam Á';
  base += '\n3. Áp dụng best practices từ ngành: mô hình bán hàng giáo dục, retention rate chuẩn EdTech, chiến lược enrollment của các hệ thống tương tự';
  base += '\n4. Hiểu đặc thù sản phẩm MindX: STEM/Coding cho trẻ, quyết định mua hàng từ phụ huynh, cần demo/trial, giá trị đơn hàng cao (18M+), học kỳ dài';
  base += '\n5. Xem xét giai đoạn/mùa vụ tuyển sinh: hè là mùa cao điểm, cuối năm là mùa re-enroll, đầu năm là mùa thấp điểm';
  base += '\n6. Đưa ra gợi ý HÀNH ĐỘNG CỤ THỂ (không chung chung), có thể áp dụng ngay tại trung tâm';

  // Inject dữ liệu nội bộ
  if (internalData && internalData.hasData) {
    if (internalData.isSystemWide && internalData.systemSummary) {
      // Chế độ toàn hệ thống
      base += '\n\n=== DỮ LIỆU TOÀN HỆ THỐNG MindX ===';
      base += '\n' + internalData.systemSummary;
      if (internalData.dailySummary) {
        base += '\n\n--- DAILY REPORT TOÀN HỆ THỐNG MTD ---';
        base += '\n' + internalData.dailySummary;
      }
      base += '\n\nLưu ý: CM không chọn BU cụ thể. Trả lời dựa trên dữ liệu toàn hệ thống (bao gồm Daily Report) + nghiên cứu thị trường.';
    } else {
      // Chế độ BU cụ thể
      base += '\n\n=== DỮ LIỆU NỘI BỘ CỦA BU ' + (internalData.bu || '') + ' ===';
      base += '\nTháng: ' + (internalData.month || '?') + ' | Tuần: ' + (internalData.week || '?');

      if (internalData.wrSummary) {
        base += '\n\n--- WEEKLY REPORT (số liệu tuần này) ---';
        base += '\n' + internalData.wrSummary;
      }
      if (internalData.apSummary) {
        base += '\n\n--- ACTION PLAN (kế hoạch hành động hiện tại) ---';
        base += '\n' + internalData.apSummary;
      }
      if (internalData.dailySummary) {
        base += '\n\n--- DAILY REPORT MTD (dữ liệu hàng ngày tích lũy) ---';
        base += '\n' + internalData.dailySummary;
      }
      if (internalData.prevWrSummary) {
        base += '\n\n--- WEEKLY REPORT TUẦN TRƯỚC (so sánh) ---';
        base += '\n' + internalData.prevWrSummary;
      }
      base += '\n\nLưu ý: Dựa vào dữ liệu trên (bao gồm Daily Report MTD) + nghiên cứu thị trường để trả lời chính xác cho CM này.';
    }
  } else {
    base += '\n\n(Không có dữ liệu WR/AP. Hãy dựa vào kiến thức nghiệp vụ + nghiên cứu thị trường EdTech để trả lời chuyên sâu.)';
  }

  return base;
}

// ============================================================
// Trích xuất dữ liệu nội bộ từ Spreadsheet cho MiMi
// Nếu có BU → lấy dữ liệu BU đó
// Nếu không có BU → tổng hợp toàn hệ thống
// ============================================================
function extractBUDataForMiMi(bu, month, week) {
  var result = { bu: bu, month: month, week: week, hasData: false, isSystemWide: false };

  if (!month || !week) return result;

  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var wrSheet = ss.getSheetByName(TABS.WR);
    var apSheet = ss.getSheetByName(TABS.AP);
    var dailySheet = getDailySheet(ss);

    if (bu) {
      // ===== CHẾ ĐỘ 1: Có BU cụ thể → lấy dữ liệu BU đó =====
      if (wrSheet) {
        var wrRows = getMatchingRows(wrSheet, HEADERS.WR, bu, month, week);
        if (wrRows.length > 0) {
          result.hasData = true;
          result.wrSummary = formatWRRows(wrRows);
        }
      }
      if (apSheet) {
        var apRows = getMatchingRows(apSheet, HEADERS.AP, bu, month, week);
        if (apRows.length > 0) {
          result.hasData = true;
          result.apSummary = formatAPRows(apRows);
        }
      }
      // WR tuần trước
      var prevWeek = getPrevWeek(week);
      if (prevWeek && wrSheet) {
        var prevWrRows = getMatchingRows(wrSheet, HEADERS.WR, bu, month, prevWeek);
        if (prevWrRows.length > 0) {
          result.prevWrSummary = formatWRRowsShort(prevWrRows);
        }
      }

      // ===== DAILY REPORT DATA =====
      if (dailySheet) {
        var dailyData = getSheetData(dailySheet, HEADERS.DAILY);
        var buDaily = dailyData.filter(function(r) {
          return String(r.bu) === String(bu) && String(r.date).substring(0,7) === String(month);
        });
        if (buDaily.length > 0) {
          result.hasData = true;
          result.dailySummary = formatDailyForMiMi(buDaily, bu);
        }
      }

    } else {
      // ===== CHẾ ĐỘ 2: Không có BU → tổng hợp toàn hệ thống =====
      result.isSystemWide = true;
      var sysData = extractSystemWideData(ss, month, week);
      if (sysData) {
        result.hasData = true;
        result.systemSummary = sysData;
      }
      // System-wide daily
      if (dailySheet) {
        var allDaily = getSheetData(dailySheet, HEADERS.DAILY);
        var monthDaily = allDaily.filter(function(r) {
          return String(r.date).substring(0,7) === String(month);
        });
        if (monthDaily.length > 0) {
          result.hasData = true;
          result.dailySummary = formatDailySystemForMiMi(monthDaily);
        }
      }
    }

  } catch (e) {
    Logger.log('extractBUDataForMiMi error: ' + e.toString());
  }

  return result;
}

// Format Daily Report data cho MiMi (single BU)
function formatDailyForMiMi(rows, bu) {
  var lines = ['DAILY REPORT MTD cho ' + bu + ' (' + rows.length + ' ng\u00e0y):'];
  var sums = { lead_mkt:0, calls_n1:0, trial_book:0, trial_done_n1:0, deal_n1:0, revenue_n1:0,
    cs_touchpoint:0, phhs_reupsell:0, deal_reupsell:0, deal_referral:0, revenue_n2:0,
    direct_sales:0, events:0, lead_n3:0, trial_n3:0, deal_n3:0, revenue_n3:0 };
  rows.forEach(function(r) {
    for (var k in sums) sums[k] += parseInt(r[k]) || 0;
  });
  var totalDeal = sums.deal_n1 + sums.deal_reupsell + sums.deal_referral + sums.deal_n3;
  var totalRev = sums.revenue_n1 + sums.revenue_n2 + sums.revenue_n3;
  lines.push('T\u1ed5ng Deal MTD: ' + totalDeal + ' | Doanh s\u1ed1 MTD: ' + totalRev + 'M');
  lines.push('N1: Lead=' + sums.lead_mkt + ', Calls=' + sums.calls_n1 + ', Trial Book=' + sums.trial_book + ', Trial Done=' + sums.trial_done_n1 + ', Deal=' + sums.deal_n1 + ', Rev=' + sums.revenue_n1 + 'M');
  var cr16_n1 = sums.lead_mkt > 0 ? (sums.deal_n1/sums.lead_mkt*100).toFixed(1) : '0';
  var cr46_n1 = sums.trial_done_n1 > 0 ? (sums.deal_n1/sums.trial_done_n1*100).toFixed(1) : '0';
  var noshow = sums.trial_book > 0 ? ((1-sums.trial_done_n1/sums.trial_book)*100).toFixed(1) : '0';
  lines.push('CR16(N1)=' + cr16_n1 + '% | CR46(N1)=' + cr46_n1 + '% | No-show=' + noshow + '%');
  lines.push('N2: CS touchpoint=' + sums.cs_touchpoint + ', PHHS t\u01b0 v\u1ea5n=' + sums.phhs_reupsell + ', Deal ReUpsell=' + sums.deal_reupsell + ', Deal Referral=' + sums.deal_referral + ', Rev=' + sums.revenue_n2 + 'M');
  lines.push('N3: Direct=' + sums.direct_sales + ', Events=' + sums.events + ', Lead=' + sums.lead_n3 + ', Trial=' + sums.trial_n3 + ', Deal=' + sums.deal_n3 + ', Rev=' + sums.revenue_n3 + 'M');
  return lines.join('\n');
}

// Format Daily Report data cho MiMi (system-wide)
function formatDailySystemForMiMi(rows) {
  var buSet = {};
  rows.forEach(function(r) { buSet[r.bu] = true; });
  var buCount = Object.keys(buSet).length;
  var lines = ['DAILY REPORT TO\u00c0N H\u1ec6 TH\u1ed0NG (' + buCount + ' BU, ' + rows.length + ' b\u1ea3n ghi):'];
  var sums = { lead_mkt:0, calls_n1:0, trial_book:0, trial_done_n1:0, deal_n1:0, revenue_n1:0,
    cs_touchpoint:0, phhs_reupsell:0, deal_reupsell:0, deal_referral:0, revenue_n2:0,
    direct_sales:0, events:0, lead_n3:0, trial_n3:0, deal_n3:0, revenue_n3:0 };
  rows.forEach(function(r) {
    for (var k in sums) sums[k] += parseInt(r[k]) || 0;
  });
  var totalDeal = sums.deal_n1 + sums.deal_reupsell + sums.deal_referral + sums.deal_n3;
  var totalRev = sums.revenue_n1 + sums.revenue_n2 + sums.revenue_n3;
  lines.push('T\u1ed5ng Deal: ' + totalDeal + ' | T\u1ed5ng Doanh s\u1ed1: ' + totalRev + 'M');
  lines.push('N1: Lead=' + sums.lead_mkt + ', Deal=' + sums.deal_n1 + ', Rev=' + sums.revenue_n1 + 'M, CR16=' + (sums.lead_mkt>0?(sums.deal_n1/sums.lead_mkt*100).toFixed(1):'0') + '%');
  lines.push('N2: CS=' + sums.cs_touchpoint + ', Deal ReUp=' + sums.deal_reupsell + ', Deal Ref=' + sums.deal_referral + ', Rev=' + sums.revenue_n2 + 'M');
  lines.push('N3: Direct=' + sums.direct_sales + ', Events=' + sums.events + ', Deal=' + sums.deal_n3 + ', Rev=' + sums.revenue_n3 + 'M');
  return lines.join('\n');
}

// Format WR rows chi tiết (có % đạt + cảnh báo)
function formatWRRows(rows) {
  return rows.map(function(r) {
    var line = r.kpi || '';
    if (r.target) line += ' | Target: ' + r.target;
    if (r.actual) line += ' | Actual: ' + r.actual;
    var t = parseFloat(r.target);
    var a = parseFloat(r.actual);
    if (!isNaN(t) && t > 0 && !isNaN(a)) {
      var pct = Math.round((a / t) * 100);
      line += ' | Đạt: ' + pct + '%';
      if (pct < 80) line += ' ⚠️ THẤP';
    }
    if (r.notes) line += ' | Ghi chú: ' + r.notes;
    return line;
  }).join('\n');
}

// Format WR rows ngắn gọn (tuần trước)
function formatWRRowsShort(rows) {
  return rows.map(function(r) {
    var line = r.kpi || '';
    if (r.target) line += ' | Target: ' + r.target;
    if (r.actual) line += ' | Actual: ' + r.actual;
    return line;
  }).join('\n');
}

// Format AP rows
function formatAPRows(rows) {
  return rows.map(function(r, idx) {
    var line = 'AP' + (idx + 1) + ': [' + (r.func || '?') + '] ' + (r.chi_so || '');
    if (r.van_de) line += ' | Vấn đề: ' + r.van_de;
    if (r.muc_do) line += ' | Mức độ: ' + r.muc_do;
    if (r.root_cause) line += ' | Root Cause: ' + r.root_cause;
    if (r.key_action) line += ' | Key Action: ' + r.key_action;
    if (r.target_do_luong) line += ' | Target: ' + r.target_do_luong;
    if (r.status) line += ' | Status: ' + r.status;
    return line;
  }).join('\n');
}

// ============================================================
// Tổng hợp dữ liệu toàn hệ thống (khi không chọn BU)
// ============================================================
function extractSystemWideData(ss, month, week) {
  loadBUListFromConfig();
  var wrSheet = ss.getSheetByName(TABS.WR);
  var apSheet = ss.getSheetByName(TABS.AP);

  if (!wrSheet) return null;

  var allWR = getSheetData(wrSheet, HEADERS.WR);
  var weekWR = allWR.filter(function(r) {
    return String(r.month) === String(month) && String(r.week) === String(week);
  });

  if (weekWR.length === 0) return null;

  // Nhóm WR theo BU
  var buMap = {};
  weekWR.forEach(function(r) {
    if (!buMap[r.bu]) buMap[r.bu] = [];
    buMap[r.bu].push(r);
  });

  var pace = { 'W1': 0.15, 'W2': 0.40, 'W3': 0.65, 'W4': 1.0 };
  var weekPace = pace[week] || 0.15;

  var totalBU = BU_LIST.length;
  var reportedBU = Object.keys(buMap).length;
  var buXanh = [];
  var buDo = [];
  var buVang = [];

  // Phân loại BU
  for (var buName in buMap) {
    var rows = buMap[buName];
    var revenueRows = rows.filter(function(r) {
      var k = (r.kpi || '').toLowerCase();
      return k.indexOf('doanh số') >= 0 || k.indexOf('tổng') >= 0 || k.indexOf('revenue') >= 0;
    });

    if (revenueRows.length === 0) {
      buVang.push(buName);
      continue;
    }

    var allMeetPace = revenueRows.every(function(r) {
      var t = parseFloat(r.target);
      var a = parseFloat(r.actual);
      return !isNaN(t) && t > 0 && !isNaN(a) && (a / t) >= weekPace;
    });

    if (allMeetPace) {
      buXanh.push(buName);
    } else {
      buDo.push(buName);
    }
  }

  // Tính trung bình CR/AOV từ tất cả BU
  var crStats = {};
  weekWR.forEach(function(r) {
    var kpi = (r.kpi || '').trim();
    var actual = parseFloat(r.actual);
    if (isNaN(actual)) return;
    if (!crStats[kpi]) crStats[kpi] = { sum: 0, count: 0, values: [] };
    crStats[kpi].sum += actual;
    crStats[kpi].count += 1;
    crStats[kpi].values.push(actual);
  });

  // Tổng hợp AP
  var apSummary = '';
  if (apSheet) {
    var allAP = getSheetData(apSheet, HEADERS.AP);
    var weekAP = allAP.filter(function(r) {
      return String(r.month) === String(month) && String(r.week) === String(week);
    });

    if (weekAP.length > 0) {
      var funcCount = { 'N1 - OPTIMIZE': 0, 'N2 - OPS': 0, 'N3 - GROWTH': 0 };
      var statusCount = { 'Đang thực hiện': 0, 'Hoàn thành': 0, 'Chưa bắt đầu': 0 };
      var topIssues = {};

      weekAP.forEach(function(r) {
        var f = r.func || '';
        if (funcCount[f] !== undefined) funcCount[f]++;
        var s = r.status || 'Chưa bắt đầu';
        if (statusCount[s] !== undefined) statusCount[s]++;
        else statusCount[s] = 1;
        var issue = r.chi_so || r.van_de || '';
        if (issue) {
          if (!topIssues[issue]) topIssues[issue] = 0;
          topIssues[issue]++;
        }
      });

      apSummary = 'Tổng AP: ' + weekAP.length + ' actions từ ' + Object.keys(buMap).length + ' BU';
      apSummary += '\nTheo Func: N1=' + funcCount['N1 - OPTIMIZE'] + ', N2=' + funcCount['N2 - OPS'] + ', N3=' + funcCount['N3 - GROWTH'];
      apSummary += '\nTheo Status: ' + Object.keys(statusCount).map(function(k) { return k + '=' + statusCount[k]; }).join(', ');

      // Top vấn đề phổ biến
      var sortedIssues = Object.keys(topIssues).sort(function(a, b) { return topIssues[b] - topIssues[a]; }).slice(0, 5);
      if (sortedIssues.length > 0) {
        apSummary += '\nVấn đề phổ biến nhất: ' + sortedIssues.map(function(k) { return k + ' (' + topIssues[k] + ' BU)'; }).join(', ');
      }
    }
  }

  // Build summary string
  var summary = 'Tháng: ' + month + ' | Tuần: ' + week + ' | Pace chuẩn: ' + Math.round(weekPace * 100) + '%';
  summary += '\nTổng BU: ' + totalBU + ' | Đã báo cáo: ' + reportedBU + '/' + totalBU;
  summary += '\n\nPHÂN LOẠI BU:';
  summary += '\n🟢 BU Xanh (đạt pace): ' + buXanh.length + ' BU' + (buXanh.length > 0 ? ' — ' + buXanh.slice(0, 8).join(', ') + (buXanh.length > 8 ? '...' : '') : '');
  summary += '\n🔴 BU Đỏ (dưới pace): ' + buDo.length + ' BU' + (buDo.length > 0 ? ' — ' + buDo.slice(0, 8).join(', ') + (buDo.length > 8 ? '...' : '') : '');
  summary += '\n🟡 Chưa rõ: ' + buVang.length + ' BU';

  // Trung bình KPI
  summary += '\n\nTRUNG BÌNH KPI TOÀN HỆ THỐNG:';
  for (var kpi in crStats) {
    var s = crStats[kpi];
    var avg = Math.round((s.sum / s.count) * 100) / 100;
    summary += '\n' + kpi + ': TB=' + avg + ' (từ ' + s.count + ' BU)';
  }

  if (apSummary) {
    summary += '\n\nACTION PLAN TOÀN HỆ THỐNG:';
    summary += '\n' + apSummary;
  }

  return summary;
}

// Tuần trước: W2->W1, W3->W2, W4->W3, W1->null
function getPrevWeek(week) {
  var w = String(week);
  if (w === 'W2') return 'W1';
  if (w === 'W3') return 'W2';
  if (w === 'W4') return 'W3';
  return null;
}

function handleMiMiAsk(body) {
  var questionId = body.question_id || '';
  var questionContent = body.question_content || '';
  var bu = body.bu || '';
  var month = body.month || '';
  var week = body.week || '';

  if (!questionId || !questionContent) {
    return makeResponse({ success: false, error: 'Missing question_id or question_content' });
  }

  // Kiểm tra câu hỏi quá nội bộ
  var qLower = questionContent.toLowerCase();
  for (var i = 0; i < INTERNAL_KEYWORDS.length; i++) {
    if (qLower.indexOf(INTERNAL_KEYWORDS[i]) >= 0) {
      var internalAnswer = 'Câu hỏi này cần HO trả lời trực tiếp. Bạn có thể nhấn "Hỏi HO" để được hỗ trợ nhé!';
      handleQAPost({ type: 'ai_answer', parent_id: questionId, bu: '', content: internalAnswer });
      return makeResponse({ success: true, answer: internalAnswer });
    }
  }

  // Trích xuất dữ liệu nội bộ (BU cụ thể hoặc toàn hệ thống)
  var internalData = extractBUDataForMiMi(bu, month, week);

  // Xây dựng system prompt với dữ liệu nội bộ
  var systemPrompt = buildMiMiSystemPrompt(internalData);

  // Xây dựng user message — thêm context hướng dẫn nghiên cứu
  var userMsg = 'Câu hỏi từ CM';
  if (bu) userMsg += ' (BU: ' + bu + ')';
  userMsg += ':\n\n' + questionContent;
  userMsg += '\n\nHãy kết hợp dữ liệu nội bộ MindX (nếu có) với nghiên cứu từ thị trường giáo dục công nghệ, xu hướng EdTech Việt Nam/Đông Nam Á, và best practices ngành để đưa ra câu trả lời chuyên sâu nhất.';

  try {
    var payload = {
      model: 'sonar-reasoning-pro',
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userMsg }
      ],
      max_tokens: 2000,
      temperature: 0.5,
      search_recency_filter: 'month'
    };

    var options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + PERPLEXITY_API_KEY },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch('https://api.perplexity.ai/chat/completions', options);
    var responseText = response.getContentText();
    var httpCode = response.getResponseCode();
    Logger.log('Perplexity API HTTP ' + httpCode + ': ' + responseText.substring(0, 500));

    var result = JSON.parse(responseText);

    // Kiểm tra API trả về lỗi
    if (httpCode !== 200 || !result.choices || !result.choices.length) {
      var apiError = result.error ? (result.error.message || JSON.stringify(result.error)) : 'Unknown API error (HTTP ' + httpCode + ')';
      Logger.log('Perplexity API error: ' + apiError);
      var errorAnswer = 'MiMi tạm thời không thể trả lời. Bạn có thể nhấn "Hỏi HO" để được hỗ trợ. (API: ' + apiError.substring(0, 80) + ')';
      handleQAPost({ type: 'ai_answer', parent_id: questionId, bu: '', content: errorAnswer });
      return makeResponse({ success: true, answer: errorAnswer });
    }

    // Xử lý <think> tags từ reasoning model
    var aiAnswer = result.choices[0].message.content.trim();
    aiAnswer = aiAnswer.replace(/<think>[\s\S]*?<\/think>/g, '').trim();

    // Lưu câu trả lời vào Sheets
    handleQAPost({ type: 'ai_answer', parent_id: questionId, bu: '', content: aiAnswer });

    return makeResponse({ success: true, answer: aiAnswer });
  } catch (e) {
    Logger.log('handleMiMiAsk exception: ' + e.toString());
    var fallbackAnswer = 'MiMi tạm thời không khả dụng. Bạn có thể thử lại sau hoặc nhấn "Hỏi HO" để được hỗ trợ. (Lỗi: ' + e.toString().substring(0, 100) + ')';
    handleQAPost({ type: 'ai_answer', parent_id: questionId, bu: '', content: fallbackAnswer });
    return makeResponse({ success: true, answer: fallbackAnswer });
  }
}

// ============================================================
// MiMi Direct — Gọi trực tiếp từ Analytics (không lưu Q&A)
// ============================================================
function handleMiMiDirect(body) {
  var question = body.question || '';
  var bu = body.bu || '';
  var month = body.month || '';
  var week = body.week || '';

  if (!question) {
    return makeResponse({ ok: false, error: 'Missing question' });
  }

  var internalData = extractBUDataForMiMi(bu, month, week);
  var systemPrompt = buildMiMiSystemPrompt(internalData);

  var userMsg = 'Phân tích từ SOD/FM Dashboard';
  if (bu) userMsg += ' (BU: ' + bu + ')';
  userMsg += ':\n\n' + question;
  userMsg += '\n\nHãy kết hợp dữ liệu nội bộ MindX với nghiên cứu từ thị trường giáo dục công nghệ, xu hướng EdTech Việt Nam/Đông Nam Á để đưa ra chiến lược cụ thể.';

  try {
    var payload = {
      model: 'sonar-reasoning-pro',
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userMsg }
      ],
      max_tokens: 2000,
      temperature: 0.5,
      search_recency_filter: 'month'
    };

    var options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + PERPLEXITY_API_KEY },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch('https://api.perplexity.ai/chat/completions', options);
    var responseText = response.getContentText();
    var httpCode = response.getResponseCode();

    var result = JSON.parse(responseText);

    if (httpCode !== 200 || !result.choices || !result.choices.length) {
      var apiError = result.error ? (result.error.message || JSON.stringify(result.error)) : 'HTTP ' + httpCode;
      return makeResponse({ ok: false, error: 'API: ' + apiError.substring(0, 100) });
    }

    var aiAnswer = result.choices[0].message.content.trim();
    aiAnswer = aiAnswer.replace(/<think>[\s\S]*?<\/think>/g, '').trim();

    return makeResponse({ ok: true, answer: aiAnswer });
  } catch (e) {
    Logger.log('handleMiMiDirect exception: ' + e.toString());
    return makeResponse({ ok: false, error: e.toString().substring(0, 100) });
  }
}

// ============================================================
// SOD Notify — Gửi câu hỏi CM lên Telegram cho SOD
// ============================================================
var TELEGRAM_BOT_TOKEN = '8661328396:AAECdS1Lxac8uHtbbuez4k2dtWUptkURHKc';
var TELEGRAM_CHAT_ID = '-5190375492';

function handleSODNotify(body) {
  var questionId = body.question_id || '';
  var questionContent = body.question_content || '';

  if (!questionId || !questionContent) {
    return makeResponse({ success: false, error: 'Missing question_id or question_content' });
  }

  var messageText = '📩 <b>Câu hỏi mới từ CM (Ẩn danh)</b>\n'
    + '━━━━━━━━━━━━━━━━━━━━\n\n'
    + questionContent + '\n\n'
    + '━━━━━━━━━━━━━━━━━━━━\n'
    + '💡 <i>Reply tin nhắn này để trả lời CM.</i>\n'
    + '🔑 <code>QID:' + questionId + '</code>';

  try {
    var telegramUrl = 'https://api.telegram.org/bot' + TELEGRAM_BOT_TOKEN + '/sendMessage';
    var payload = {
      chat_id: TELEGRAM_CHAT_ID,
      text: messageText,
      parse_mode: 'HTML'
    };

    var options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(telegramUrl, options);
    var result = JSON.parse(response.getContentText());

    if (result.ok) {
      return makeResponse({ success: true, message_id: result.result.message_id, question_id: questionId });
    } else {
      return makeResponse({ success: false, error: result.description || 'Telegram API error' });
    }
  } catch (e) {
    return makeResponse({ success: false, error: e.toString() });
  }
}

// ============================================================
// Telegram Polling — Kiểm tra reply từ SOD trên Telegram
// Chạy mỗi phút qua time-driven trigger
// ============================================================
function pollTelegramReplies() {
  var props = PropertiesService.getScriptProperties();
  var lastUpdateId = parseInt(props.getProperty('TELEGRAM_LAST_UPDATE_ID') || '0');

  var telegramUrl = 'https://api.telegram.org/bot' + TELEGRAM_BOT_TOKEN + '/getUpdates';
  var params = { offset: lastUpdateId + 1, timeout: 5 };

  try {
    var response = UrlFetchApp.fetch(telegramUrl + '?offset=' + params.offset + '&timeout=' + params.timeout, {
      method: 'get',
      muteHttpExceptions: true
    });
    var data = JSON.parse(response.getContentText());

    if (!data.ok || !data.result || data.result.length === 0) {
      return; // Không có update mới
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = getQASheet(ss);

    for (var i = 0; i < data.result.length; i++) {
      var update = data.result[i];
      lastUpdateId = update.update_id; // Cập nhật ID mới nhất

      var message = update.message || {};
      var text = (message.text || '').trim();
      if (!text) continue;

      // Chỉ xử lý reply tin nhắn (không phải tin nhắn mới)
      var replyTo = message.reply_to_message || null;
      if (!replyTo) continue;

      // Tìm QID trong tin nhắn gốc
      var originalText = replyTo.text || '';
      var qidMatch = originalText.match(/QID:(qa_[^\s\n]+)/);
      if (!qidMatch) continue;

      var questionId = qidMatch[1];

      // Lưu câu trả lời SOD vào Sheets
      var id = generateQAId();
      var createdAt = new Date().toISOString();
      var newRow = [id, 'sod_answer', questionId, 'SOD', text, 0, createdAt];
      var lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, 1, QA_HEADERS.length).setValues([newRow]);

      // Gửi xác nhận lại cho SOD trên Telegram
      try {
        var confirmUrl = 'https://api.telegram.org/bot' + TELEGRAM_BOT_TOKEN + '/sendMessage';
        UrlFetchApp.fetch(confirmUrl, {
          method: 'post',
          contentType: 'application/json',
          payload: JSON.stringify({
            chat_id: message.chat.id,
            text: '✅ Câu trả lời của bạn đã được ghi nhận và hiển thị trên platform cho CM.',
            reply_to_message_id: message.message_id
          }),
          muteHttpExceptions: true
        });
      } catch (e) {
        // Không block
      }

      Logger.log('Saved SOD answer for ' + questionId + ': ' + text.substring(0, 50));
    }

    // Lưu last update ID
    props.setProperty('TELEGRAM_LAST_UPDATE_ID', lastUpdateId.toString());

  } catch (e) {
    Logger.log('Telegram polling error: ' + e.toString());
  }
}

// ============================================================
// Setup Telegram Polling Trigger — Chạy 1 lần để tạo trigger
// ============================================================
function setupTelegramPolling() {
  // Xóa trigger cũ nếu có
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'pollTelegramReplies') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Tạo trigger mới: chạy mỗi 1 phút
  ScriptApp.newTrigger('pollTelegramReplies')
    .timeBased()
    .everyMinutes(1)
    .create();

  Logger.log('Telegram polling trigger đã được tạo — kiểm tra mỗi phút.');
}

// ============================================================
// KPI TARGETS — Lưu và đọc KPI target theo tháng/BU
// ============================================================
var KPI_TARGETS_HEADERS = ['month', 'bu', 'kpi_n1', 'kpi_n2', 'kpi_n3', 'kpi_total', 'saved_at'];

function getKPITargetsSheet(ss) {
  var sheet = ss.getSheetByName('KPI_Targets');
  if (!sheet) {
    sheet = ss.insertSheet('KPI_Targets');
    sheet.getRange(1, 1, 1, KPI_TARGETS_HEADERS.length).setValues([KPI_TARGETS_HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** POST: save_kpi_targets — body = { month, targets: [{ bu, kpi_n1, kpi_n2, kpi_n3, kpi_total }] } */
function handleSaveKPITargets(body) {
  var month = String(body.month);
  var targets = body.targets || [];
  if (!month || targets.length === 0) {
    return makeResponse({ success: false, error: 'Missing month or targets' });
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getKPITargetsSheet(ss);
  var data = sheet.getDataRange().getValues();

  // Keep rows from OTHER months (filter out current month)
  var keepRows = [data[0]]; // header
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) !== month) {
      keepRows.push(data[i]);
    }
  }

  // Add new rows for this month
  var savedAt = new Date().toISOString();
  targets.forEach(function(t) {
    keepRows.push([month, t.bu, Number(t.kpi_n1) || 0, Number(t.kpi_n2) || 0, Number(t.kpi_n3) || 0, Number(t.kpi_total) || 0, savedAt]);
  });

  // Rewrite entire sheet
  sheet.clearContents();
  if (keepRows.length > 0) {
    sheet.getRange(1, 1, keepRows.length, KPI_TARGETS_HEADERS.length).setValues(keepRows);
  }

  return makeResponse({ success: true, count: targets.length, month: month });
}

/** Convert cell value to month string. Handles Date objects from Google Sheets auto-conversion. */
function cellToMonth(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return val.getFullYear() + '-' + String(val.getMonth() + 1).padStart(2, '0');
  }
  return String(val);
}

/** GET: get_kpi_targets — params: month */
function handleGetKPITargets(params) {
  var month = String(params.month || '');
  if (!month) {
    return makeResponse({ success: false, error: 'Missing month' });
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getKPITargetsSheet(ss);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var targets = [];
  for (var i = 1; i < data.length; i++) {
    if (cellToMonth(data[i][0]) === month) {
      var obj = {};
      for (var j = 0; j < headers.length; j++) {
        obj[headers[j]] = j === 0 ? cellToMonth(data[i][j]) : data[i][j];
      }
      targets.push(obj);
    }
  }

  return makeResponse({ success: true, month: month, targets: targets });
}

/** GET: clear_kpi_targets — params: month (delete all KPI for a month) */
function handleClearKPITargets(params) {
  var month = String(params.month || '');
  if (!month) return makeResponse({ success: false, error: 'Missing month' });

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getKPITargetsSheet(ss);
  var data = sheet.getDataRange().getValues();

  var keepRows = [data[0]]; // header
  for (var i = 1; i < data.length; i++) {
    if (cellToMonth(data[i][0]) !== month) {
      keepRows.push(data[i]);
    }
  }

  sheet.clearContents();
  if (keepRows.length > 0) {
    sheet.getRange(1, 1, keepRows.length, KPI_TARGETS_HEADERS.length).setValues(keepRows);
  }
  return makeResponse({ success: true, cleared: month });
}

/** GET: save_kpi_targets_batch — params: month, data (JSON-encoded array of targets) */
function handleSaveKPIBatch(params) {
  var month = String(params.month || '');
  var jsonData = params.data || '[]';
  if (!month) return makeResponse({ success: false, error: 'Missing month' });

  var targets;
  try {
    targets = JSON.parse(jsonData);
  } catch(e) {
    return makeResponse({ success: false, error: 'Invalid JSON: ' + e.toString() });
  }
  if (!targets || targets.length === 0) {
    return makeResponse({ success: false, error: 'Empty targets' });
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getKPITargetsSheet(ss);
  var savedAt = new Date().toISOString();

  var newRows = targets.map(function(t) {
    return [month, t.bu, Number(t.kpi_n1) || 0, Number(t.kpi_n2) || 0, Number(t.kpi_n3) || 0, Number(t.kpi_total) || 0, savedAt];
  });

  var lastRow = sheet.getLastRow();
  var startRow = lastRow + 1;
  sheet.getRange(startRow, 1, newRows.length, KPI_TARGETS_HEADERS.length).setValues(newRows);

  // Force month column to plain text so Google Sheets doesn't auto-convert to Date
  var monthRange = sheet.getRange(startRow, 1, newRows.length, 1);
  monthRange.setNumberFormat('@'); // @ = plain text format

  return makeResponse({ success: true, count: newRows.length, month: month });
}

// ============================================================
// GET-based save for AP/WR (replaces POST which fails due to redirect)
// ============================================================

/** GET: clear_tab_data — params: tab, bu, month, week */
function handleClearTabData(params) {
  var tab = params.tab;
  var bu = params.bu;
  var month = String(params.month);
  var week = String(params.week);
  if (!tab || !bu || !month || !week) {
    return makeResponse({ success: false, error: 'Missing params: tab, bu, month, week' });
  }
  if (!TABS[tab]) {
    return makeResponse({ success: false, error: 'Invalid tab: ' + tab });
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(TABS[tab]);
  if (!sheet) {
    return makeResponse({ success: false, error: 'Sheet not found: ' + TABS[tab] });
  }

  var headers = HEADERS[tab];
  deleteMatchingRows(sheet, headers, bu, month, week);

  return makeResponse({ success: true, cleared: bu + '/' + month + '/W' + week });
}

/** GET: save_row — params: tab, data (JSON with bu, month, week, row, row_index) */
function handleSaveRow(params) {
  var tab = params.tab;
  var jsonData = params.data || '{}';
  if (!tab) return makeResponse({ success: false, error: 'Missing tab' });
  if (!TABS[tab]) return makeResponse({ success: false, error: 'Invalid tab: ' + tab });

  var data;
  try {
    data = JSON.parse(jsonData);
  } catch(e) {
    return makeResponse({ success: false, error: 'Invalid JSON: ' + e.toString() });
  }

  var bu = data.bu;
  var month = String(data.month);
  var week = String(data.week);
  var row = data.row;
  if (!bu || !month || !week || !row) {
    return makeResponse({ success: false, error: 'Missing bu/month/week/row in data' });
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(TABS[tab]);
  if (!sheet) return makeResponse({ success: false, error: 'Sheet not found' });

  var headers = HEADERS[tab];
  var savedAt = new Date().toISOString();

  var newRow = headers.map(function(h) {
    if (h === 'bu') return bu;
    if (h === 'month') return month;
    if (h === 'week') return week;
    if (h === 'saved_at') return savedAt;
    var val = row[h];
    return (val === undefined || val === null) ? '' : String(val);
  });

  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, headers.length).setValues([newRow]);

  return makeResponse({ success: true, row_index: data.row_index });
}

// ============================================================
// get_staff_list — Trả danh sách nhân viên Active từ sheet Staff
// Staff sheet gid=1896503199
// Cột: A=MSNV, B=Họ tên, C=Vị trí, D=Cơ sở (BU), E=Khu vực, F=Trạng thái
// ============================================================

// Mapping từ BU raw + region trong sheet → BU name trong app
const STAFF_BU_MAPPING = {
  // HCM1
  'PVT|HCM1': 'HCM1 - PVT',
  'PXL|HCM1': 'HCM1 - PXL',
  'TK|HCM1': 'HCM1 - TK',
  // HCM2
  'LVV|HCM2': 'HCM2 - LVV',
  'NX|HCM2': 'HCM2 - NX',
  'PVĐ|HCM2': 'HCM2 - PVD',
  'PVD|HCM2': 'HCM2 - PVD',
  'SH|HCM2': 'HCM2 - SH',
  '5H|HCM2': 'HCM2 - SH',
  // HCM3
  '3/2|HCM3': 'HCM3 - 3T2',
  'HL|HCM3': 'HCM3 - HL',
  'HL (Nguyễn Thị Thập)|HCM3': 'HCM3 - HL',
  'HTLO|HCM3': 'HCM3 - HTLO',
  'PMH|HCM3': 'HCM3 - PMH',
  'PNL|HCM3': 'HCM3 - PNL',
  // HCM4
  'LBB|HCM4': 'HCM4 - LBB',
  'Trường Chinh|HCM4': 'HCM4 - TC',
  'TC|HCM4': 'HCM4 - TC',
  'TL|HCM4': 'HCM4 - TL',
  'TT|HCM4': 'HCM4 - TT',
  // HN1
  'MK|HN1': 'HN - MK',
  'NHT|HN1': 'HN - NHT',
  'NVC HN|HN1': 'HN - NVC',
  'NVC|HN1': 'HN - NVC',
  'Ocean Park|HN1': 'HN - OCP',
  'OCP|HN1': 'HN - OCP',
  'TP|HN1': 'HN - TP',
  'Văn Phú|HN1': 'HN - VP',
  'VP|HN1': 'HN - VP',
  // HN2
  'HĐT|HN2': 'HN - HĐT',
  'Hàm Nghi|HN2': 'HN - VHHN',
  'VHHN|HN2': 'HN - VHHN',
  'NCT|HN2': 'HN - NCT',
  'NPS|HN2': 'HN - NPS',
  // MB1 (Tỉnh Bắc 1)
  'BN Lý Thái Tổ|Tỉnh Bắc 1': 'MB1 - BN',
  'BN|Tỉnh Bắc 1': 'MB1 - BN',
  'BN Từ Sơn|Tỉnh Bắc 1': 'MB1 - TS',
  'TS|Tỉnh Bắc 1': 'MB1 - TS',
  'HP|Tỉnh Bắc 1': 'MB1 - HP',
  'QN|Tỉnh Bắc 1': 'MB1 - QN',
  // MB2 (Tỉnh Bắc 2)
  'PT|Tỉnh Bắc 2': 'MB2 - PT',
  'VP|Tỉnh Bắc 2': 'MB2 - VP',
  'TN|Tỉnh Bắc 2': 'MB2 - TN',
  // MN (Tỉnh Nam)
  'BH (Đồng Nai)|Tỉnh Nam': 'MN - BH - PVT',
  'BH|Tỉnh Nam': 'MN - BH - PVT',
  'CT|Tỉnh Nam': 'MN - CT - THD',
  'DA|Tỉnh Nam': 'MN - BD - DA',
  'Lái Thiêu|Tỉnh Nam': 'MN - BD - TA',
  'TA|Tỉnh Nam': 'MN - BD - TA',
  'TDM|Tỉnh Nam': 'MN - BD - TDM',
  'VT|Tỉnh Nam': 'MN - VT - LHP',
  // MT (Tỉnh Trung)
  'DN|Tỉnh Trung': 'MT - ĐN',
  'ĐN|Tỉnh Trung': 'MT - ĐN',
  'TH|Tỉnh Trung': 'MT - TH',
  'NA|Tỉnh Trung': 'MT - NA',
  // K18
  'AM|18+': 'K18 - HCM',
  // ONL
  'Art|ONL': 'ONL - ART',
  'Coding|ONL': 'ONL - COD'
};

function resolveStaffBU(buRaw, region) {
  if (!buRaw) return '';
  var raw = String(buRaw).trim();
  var reg = String(region || '').trim();

  // Try exact match with region
  var key1 = raw + '|' + reg;
  if (STAFF_BU_MAPPING[key1]) return STAFF_BU_MAPPING[key1];

  // Try with cleaned region (remove parentheses etc)
  // Try partial matches
  var keys = Object.keys(STAFF_BU_MAPPING);
  for (var i = 0; i < keys.length; i++) {
    var parts = keys[i].split('|');
    if (parts[0] === raw && reg.indexOf(parts[1]) !== -1) {
      return STAFF_BU_MAPPING[keys[i]];
    }
    if (raw.indexOf(parts[0]) !== -1 && parts[1] === reg) {
      return STAFF_BU_MAPPING[keys[i]];
    }
  }

  return '';
}

function handleGetStaffList() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  // Get sheet by gid
  var sheets = ss.getSheets();
  var staffSheet = null;
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === 1896503199) {
      staffSheet = sheets[i];
      break;
    }
  }
  if (!staffSheet) {
    return makeResponse({ error: 'Staff sheet not found (gid=1896503199)' });
  }

  var data = staffSheet.getDataRange().getValues();
  if (data.length < 2) return makeResponse([]);

  // Find column indices from header row
  var headers = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
  var colMSNV = -1, colName = -1, colPos = -1, colBU = -1, colRegion = -1, colStatus = -1;
  for (var c = 0; c < headers.length; c++) {
    var h = headers[c];
    if (h === 'msnv' || h === 'mã số nhân viên' || h === 'ma so nhan vien') colMSNV = c;
    else if (h === 'họ tên' || h === 'ho ten' || h === 'name' || h === 'họ và tên') colName = c;
    else if (h === 'vị trí' || h === 'vi tri' || h === 'position' || h === 'vi trí') colPos = c;
    else if (h === 'cơ sở' || h === 'co so' || h === 'cơ sở (bu)' || h === 'bu') colBU = c;
    else if (h === 'khu vực' || h === 'khu vuc' || h === 'region') colRegion = c;
    else if (h === 'trạng thái' || h === 'trang thai' || h === 'status') colStatus = c;
  }

  // Fallback: if headers not found, assume A-F
  if (colMSNV === -1) colMSNV = 0;
  if (colName === -1) colName = 1;
  if (colPos === -1) colPos = 2;
  if (colBU === -1) colBU = 3;
  if (colRegion === -1) colRegion = 4;
  if (colStatus === -1) colStatus = 5;

  var result = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var msnv = String(row[colMSNV] || '').trim();
    var status = String(row[colStatus] || '').trim();

    // Only Active staff
    if (!msnv || (status.toLowerCase() !== 'active' && status.toLowerCase() !== 'đang làm việc')) continue;

    var name = String(row[colName] || '').trim();
    var position = String(row[colPos] || '').trim();
    var buRaw = String(row[colBU] || '').trim();
    var region = String(row[colRegion] || '').trim();
    var buApp = resolveStaffBU(buRaw, region);

    result.push({
      msnv: msnv,
      name: name,
      position: position,
      bu_raw: buRaw,
      region: region,
      bu_app: buApp
    });
  }

  return makeResponse(result);
}
