// MindX CM Report Platform - app.js
// Phiên bản: Action Plan only (WR đã loại bỏ)
// Backend: Google Apps Script Web App (Google Sheets)
'use strict';

// ===================== CONFIGURATION =====================
// Thay '__APPS_SCRIPT_URL__' bằng URL Web App sau khi deploy Google Apps Script
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyJpFo_X0M_uvSPCStTOPVIAganyFRaN2AaCGO-Ukv911nlVS3me-C0jfIc8AiTSF1V/exec';
const CGI_BIN = '__CGI_BIN__';

function isConfigured() {
  return APPS_SCRIPT_URL && APPS_SCRIPT_URL !== '__APPS_SCRIPT_URL__';
}

// ===================== STATE =====================
const state = {
  currentView: 'landing',
  cm: {
    bu: '',
    // AP period — tính trực tiếp từ ngày hiện tại
    apMonth: '',
    apWeek: 1,
    activeTab: 'AP',
    savedAt: { AP: null, WR: null },
    apRows: [],
    wrRows: [],
    isLocked: false,
    period: null // detectReportPeriod() result
  },
  dashboard: {
    month: '',
    week: 1,
    allData: null,
    scores: [],
    filteredScores: [],
    selectedBU: null,
    sortCol: 'healthStatus',
    sortDir: 'asc',
    funcFilter: 'ALL'
  },
  config: {
    month: '',
    data: {},        // config data for current month
    months: [],      // list of months that have config
    role: 'sod',     // 'sod' can edit, 'fm' view only
    dirty: false
  },
  discussion: {
    questions: [],       // danh sách câu hỏi đã tải
    page: 1,             // trang hiện tại
    hasMore: false,      // có thêm câu hỏi không
    loading: false,      // đang tải không
    expandedIds: new Set(), // câu hỏi đang mở rộng
    votedIds: new Set(),    // câu hỏi đã upvote (lưu session)
    refreshTimer: null,  // timer auto-refresh
    sortBy: 'newest',    // newest | hot | most_replies
    searchTerm: ''       // từ khóa tìm kiếm
  },
  daily: {
    bu: '',
    date: '',
    data: {},            // field values keyed by field id
    mtd: null,           // MTD summary data
    loaded: false        // whether current selection has been loaded
  }
};

// ===================== CONSTANTS =====================
const BU_LIST = [
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

// Daily Report fields — Hướng A: bao gồm doanh số, thay thế WR hoàn toàn
const DAILY_FIELDS = [
  { id: 'lead_mkt',       label: 'Lead MKT nhận',          group: 'N1', type: 'number' },
  { id: 'calls_n1',       label: 'Cuộc gọi N1',            group: 'N1', type: 'number' },
  { id: 'trial_book',     label: 'Lịch hẹn Trial book',    group: 'N1', type: 'number' },
  { id: 'trial_done_n1',  label: 'Trial thực hiện N1',     group: 'N1', type: 'number' },
  { id: 'deal_n1',        label: 'Deal N1',                group: 'N1', type: 'number' },
  { id: 'revenue_n1',     label: 'Doanh số N1 (triệu)',    group: 'N1', type: 'number' },
  { id: 'cs_touchpoint',  label: 'PHHS chăm sóc',          group: 'N2', type: 'number' },
  { id: 'phhs_reupsell',  label: 'PHHS tư vấn Re/Upsell',  group: 'N2', type: 'number' },
  { id: 'deal_reupsell',  label: 'Deal Re/Upsell',         group: 'N2', type: 'number' },
  { id: 'deal_referral',  label: 'Deal Referral',          group: 'N2', type: 'number' },
  { id: 'revenue_n2',     label: 'Doanh số N2 (triệu)',    group: 'N2', type: 'number' },
  { id: 'direct_sales',   label: 'Lượt đi Direct Sales',   group: 'N3', type: 'number' },
  { id: 'events',         label: 'Event tổ chức',          group: 'N3', type: 'number' },
  { id: 'lead_n3',        label: 'Lead N3 phát sinh',      group: 'N3', type: 'number' },
  { id: 'trial_n3',       label: 'Trial N3',               group: 'N3', type: 'number' },
  { id: 'deal_n3',        label: 'Deal N3',                group: 'N3', type: 'number' },
  { id: 'revenue_n3',     label: 'Doanh số N3 (triệu)',    group: 'N3', type: 'number' },
  { id: 'note',           label: 'Ghi chú',                group: 'TOTAL', type: 'text' }
];

const DAILY_HEADERS = [
  'bu', 'date',
  'lead_mkt', 'calls_n1', 'trial_book', 'trial_done_n1', 'deal_n1', 'revenue_n1',
  'cs_touchpoint', 'phhs_reupsell', 'deal_reupsell', 'deal_referral', 'revenue_n2',
  'direct_sales', 'events', 'lead_n3', 'trial_n3', 'deal_n3', 'revenue_n3',
  'note', 'saved_at'
];

// Action Plan: khởi tạo 3 dòng trống (1 per function) để CM tự điền vấn đề
const AP_DEFAULT_ROWS = [
  { func: 'GROWTH',   chi_so: '', van_de: '', muc_do: '', root_cause: '', key_action: '', mo_ta_trien_khai: '', target_do_luong: '', deadline: '', owner: '', fm_support: '', status: '' },
  { func: 'OPTIMIZE', chi_so: '', van_de: '', muc_do: '', root_cause: '', key_action: '', mo_ta_trien_khai: '', target_do_luong: '', deadline: '', owner: '', fm_support: '', status: '' },
  { func: 'OPS',      chi_so: '', van_de: '', muc_do: '', root_cause: '', key_action: '', mo_ta_trien_khai: '', target_do_luong: '', deadline: '', owner: '', fm_support: '', status: '' }
];

// 17 KPI cố định cho Weekly Report (OKRs 2026)
const WR_DEFAULT_ROWS = [
  { kpi: 'Growth (N3) - L1 Leads tự kiếm' },
  { kpi: 'Growth (N3) - L4 Trials' },
  { kpi: 'Growth (N3) - L6 Deals' },
  { kpi: 'Growth (N3) - Doanh số N3' },
  { kpi: 'Optimize (N1) - L1 Lead MKT' },
  { kpi: 'Optimize (N1) - L4 Trial MKT' },
  { kpi: 'Optimize (N1) - L6 Deal MKT' },
  { kpi: 'Optimize (N1) - CR16%' },
  { kpi: 'Optimize (N1) - CR46%' },
  { kpi: 'Optimize (N1) - AOV' },
  { kpi: 'Optimize (N1) - Doanh số N1' },
  { kpi: 'Ops (N2) - Số PHHS tiềm năng Re/Up' },
  { kpi: 'Ops (N2) - L6 Re/Upsell' },
  { kpi: 'Ops (N2) - L2 Referral' },
  { kpi: 'Ops (N2) - L6 Referral' },
  { kpi: 'Ops (N2) - Doanh số N2' },
  { kpi: 'TỔNG DOANH SỐ BU' }
];

const MONTHS = [
  '2025-01','2025-02','2025-03','2025-04','2025-05','2025-06',
  '2025-07','2025-08','2025-09','2025-10','2025-11','2025-12',
  '2026-01','2026-02','2026-03','2026-04','2026-05','2026-06',
  '2026-07','2026-08','2026-09','2026-10','2026-11','2026-12'
];

// Tỷ lệ doanh số kỳ vọng theo tuần (pace)
const WEEK_PACE = { 1: 0.15, 2: 0.40, 3: 0.70, 4: 1.00 };

// ===================== AUTO PERIOD DETECTION =====================
// Quy tắc tuần: W1=1-7, W2=8-14, W3=15-21, W4=22+
// CM nộp báo cáo CN + T2. Khóa sau T4.
// WR = tuần vừa qua, AP = tuần tới

function getWeekOfMonth(dayOfMonth) {
  if (dayOfMonth <= 7) return 1;
  if (dayOfMonth <= 14) return 2;
  if (dayOfMonth <= 21) return 3;
  return 4;
}

function getPrevMonth(year, month) {
  if (month === 1) return { year: year - 1, month: 12 };
  return { year, month: month - 1 };
}

function getNextMonth(year, month) {
  if (month === 12) return { year: year + 1, month: 1 };
  return { year, month: month + 1 };
}

function fmtMonth(year, month) {
  return `${year}-${String(month).padStart(2, '0')}`;
}

/**
 * Tự động xác định chu kỳ AP dựa trên ngày hiện tại.
 * AP = tuần tiếp theo (kế hoạch cho tuần sắp tới).
 * CM ghi và FM đọc cùng dùng chung chu kỳ này.
 * @returns {{ ap: {month, week, label}, isLocked, lockMessage, deadlineLabel, todayName }}
 */
function detectReportPeriod(now) {
  if (!now) now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1; // 1-12
  const dayOfMonth = now.getDate();
  const dayOfWeek = now.getDay(); // 0=CN, 1=T2, ..., 6=T7

  const currentWeek = getWeekOfMonth(dayOfMonth);

  // AP = tuần TIẾP THEO (kế hoạch cho tuần sắp tới)
  let apMonth, apWeek;
  if (currentWeek < 4) {
    apMonth = fmtMonth(year, month);
    apWeek = currentWeek + 1;
  } else {
    // W4 → W1 tháng sau
    const next = getNextMonth(year, month);
    apMonth = fmtMonth(next.year, next.month);
    apWeek = 1;
  }

  // Deadline: khóa chỉ T5 và T6. Mở lại từ T7.
  const isLocked = dayOfWeek === 4 || dayOfWeek === 5;

  const dayNames = ['Chủ nhật', 'Thứ 2', 'Thứ 3', 'Thứ 4', 'Thứ 5', 'Thứ 6', 'Thứ 7'];
  const todayName = dayNames[dayOfWeek];

  let lockMessage = '';
  let deadlineLabel = '';
  if (isLocked) {
    lockMessage = `Đã khóa chỉnh sửa (hôm nay ${todayName}). Mở lại vào Thứ 7.`;
    deadlineLabel = 'Đã khóa';
  } else {
    let daysUntilDeadline;
    if (dayOfWeek === 6) daysUntilDeadline = 5;
    else if (dayOfWeek === 0) daysUntilDeadline = 4;
    else daysUntilDeadline = 3 - dayOfWeek + 1;
    if (dayOfWeek === 3) {
      deadlineLabel = 'Hạn cuối hôm nay (Thứ 4)';
    } else {
      deadlineLabel = `Còn ${daysUntilDeadline} ngày (hạn Thứ 4)`;
    }
    lockMessage = '';
  }

  function monthLabel(m) {
    const [y, mo] = m.split('-');
    return `Tháng ${parseInt(mo)}/${y}`;
  }

  return {
    ap: {
      month: apMonth,
      week: apWeek,
      label: `${monthLabel(apMonth)} — Tuần ${apWeek}`
    },
    isLocked,
    lockMessage,
    deadlineLabel,
    todayName
  };
}

// ===================== UTILITIES =====================
function fmt(dt) {
  if (!dt) return '';
  const d = new Date(dt);
  if (isNaN(d)) return dt;
  return d.toLocaleString('vi-VN', { day:'2-digit', month:'2-digit', hour:'2-digit', minute:'2-digit' });
}

function fmtNow() {
  return new Date().toLocaleString('vi-VN', {
    day:'2-digit', month:'2-digit', year:'numeric',
    hour:'2-digit', minute:'2-digit', second:'2-digit'
  });
}

function showToast(msg, type = 'success') {
  const container = document.getElementById('toast-container');
  const toast = document.createElement('div');
  toast.className = `toast${type !== 'success' ? ' ' + type : ''}`;
  toast.textContent = msg;
  container.appendChild(toast);
  setTimeout(() => toast.remove(), 3100);
}

function getRegion(buName) {
  if (!buName) return '';
  const n = buName.toUpperCase();
  if (n.startsWith('HCM')) return 'HCM';
  if (n.startsWith('HN')) return 'HN';
  if (n.startsWith('MB')) return 'MB';
  if (n.startsWith('MN')) return 'MN';
  if (n.startsWith('MT')) return 'MT';
  if (n.startsWith('K18')) return 'K18';
  if (n.startsWith('ONL')) return 'ONL';
  return '';
}

function escHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;');
}

// ===================== APPS SCRIPT API CALLS =====================
async function apiFetch(action, extraParams = {}) {
  if (!isConfigured()) {
    throw new Error('APPS_SCRIPT_URL chưa được cấu hình. Xem hướng dẫn thiết lập.');
  }
  const url = new URL(APPS_SCRIPT_URL);
  url.searchParams.set('action', action);
  Object.entries(extraParams).forEach(([k, v]) => url.searchParams.set(k, v));

  // 60s timeout for large responses
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), 60000);
  try {
    const res = await fetch(url.toString(), { signal: controller.signal });
    clearTimeout(timeoutId);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return res.json();
  } catch(e) {
    clearTimeout(timeoutId);
    if (e.name === 'AbortError') throw new Error('Timeout - thử lại sau');
    throw e;
  }
}

async function apiPost(action, tab, body) {
  if (!isConfigured()) {
    throw new Error('APPS_SCRIPT_URL chưa được cấu hình. Xem hướng dẫn thiết lập.');
  }
  const url = new URL(APPS_SCRIPT_URL);
  url.searchParams.set('action', action);
  if (tab) url.searchParams.set('tab', tab);
  const res = await fetch(url.toString(), {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain' },
    body: JSON.stringify(body)
  });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  return res.json();
}

// ===================== SETUP BANNER =====================
function showSetupBanner() {
  const existing = document.getElementById('setup-banner');
  if (existing) return;
  const banner = document.createElement('div');
  banner.id = 'setup-banner';
  banner.innerHTML = `
    <div class="setup-banner-inner">
      <span class="setup-banner-icon">⚙</span>
      <div class="setup-banner-text">
        <strong>Chưa kết nối Google Sheets.</strong>
        Mở file <code>app.js</code>, thay <code>__APPS_SCRIPT_URL__</code> bằng URL Web App của bạn.
        <a href="SETUP.md" target="_blank" rel="noopener noreferrer">Xem hướng dẫn →</a>
      </div>
    </div>
  `;
  document.body.insertBefore(banner, document.getElementById('app-header').nextSibling);
}

// ===================== ACCESS CONTROL =====================
const ACCESS_PROTECTED_VIEWS = ['dashboard', 'config'];
let accessGranted = false;

// Mật khẩu được encode
const ENCODED_PW = btoa('MindX@123');

function checkPassword(input) {
  return btoa(input) === ENCODED_PW;
}

function showPasswordModal(targetView) {
  const modal = document.getElementById('pw-modal');
  if (!modal) return;
  modal.dataset.targetView = targetView;
  modal.classList.add('show');
  const inp = document.getElementById('pw-input');
  if (inp) { inp.value = ''; inp.focus(); }
  const errEl = document.getElementById('pw-error');
  if (errEl) errEl.style.display = 'none';
}

function closePasswordModal() {
  const modal = document.getElementById('pw-modal');
  if (modal) modal.classList.remove('show');
}

function submitPassword() {
  const inp = document.getElementById('pw-input');
  const pw = inp ? inp.value : '';
  if (checkPassword(pw)) {
    accessGranted = true;
    closePasswordModal();
    const targetView = document.getElementById('pw-modal').dataset.targetView;
    navigate(targetView);
  } else {
    const errEl = document.getElementById('pw-error');
    if (errEl) { errEl.textContent = 'Sai mật khẩu. Vui lòng thử lại.'; errEl.style.display = 'block'; }
    if (inp) { inp.value = ''; inp.focus(); }
  }
}

// ===================== ROUTING =====================
function navigate(view) {
  // Kiểm tra mật khẩu cho dashboard và config
  if (ACCESS_PROTECTED_VIEWS.includes(view) && !accessGranted) {
    showPasswordModal(view);
    return;
  }

  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  document.querySelectorAll('.header-nav-btn').forEach(b => b.classList.remove('active'));

  const viewEl = document.getElementById(`view-${view}`);
  if (viewEl) viewEl.classList.add('active');

  const navBtn = document.querySelector(`[data-nav="${view}"]`);
  if (navBtn) navBtn.classList.add('active');

  state.currentView = view;
  window.location.hash = view === 'landing' ? '' : view;

  // Đóng tất cả detail panel khi chuyển view
  const dashPanel = document.getElementById('detail-panel');
  if (dashPanel) dashPanel.classList.remove('open');
  const overlay = document.getElementById('panel-overlay');
  if (overlay) overlay.classList.remove('active');

  // Update diagnostic visibility when switching to CM view
  if (view === 'cm') {
    updateDiagnosticVisibility();
  }

  // Auto-load dashboard khi navigate đến (nếu chưa có data hoặc lần đầu)
  if (view === 'dashboard' && !state.dashboard.allData) {
    runAnalysis();
  }

  // Init config view
  if (view === 'config') {
    initConfigView();
  }

  // Auto-load discussion khi navigate đến
  if (view === 'discussion') {
    initDiscussionView();
  }

  // Auto-init daily view
  if (view === 'daily') {
    initDailyView();
  }
}

function initRouting() {
  const hash = window.location.hash.replace('#', '');
  if (hash === 'cm') navigate('cm');
  else if (hash === 'dashboard') navigate('dashboard');
  else if (hash === 'config') navigate('config');
  else if (hash === 'discussion') navigate('discussion');
  else if (hash === 'daily') navigate('daily');
  else navigate('landing');

  window.addEventListener('hashchange', () => {
    const h = window.location.hash.replace('#', '');
    if (h === 'cm') navigate('cm');
    else if (h === 'dashboard') navigate('dashboard');
    else if (h === 'config') navigate('config');
    else if (h === 'discussion') navigate('discussion');
    else if (h === 'daily') navigate('daily');
    else navigate('landing');
  });
}

// ===================== TOPBAR SELECTS =====================
/**
 * Tạo danh sách chu kỳ AP gần nhất (8 chu kỳ trước + chu kỳ hiện tại)
 * CM và FM dùng chung chu kỳ này để ghi/đọc AP
 */
function generatePeriodHistory(currentPeriod) {
  const periods = [];
  // Bắt đầu từ chu kỳ AP hiện tại, lùi dần về trước
  let m = currentPeriod.ap.month;
  let w = currentPeriod.ap.week;

  for (let i = 0; i < 9; i++) {
    const [y, mo] = m.split('-');
    const label = `Tháng ${parseInt(mo)}/${y} — Tuần ${w}`;
    periods.push({ month: m, week: w, label, isCurrent: i === 0 });

    // Lùi 1 tuần
    if (w > 1) {
      w--;
    } else {
      // W1 → W4 tháng trước
      const prev = getPrevMonth(parseInt(y), parseInt(mo));
      m = fmtMonth(prev.year, prev.month);
      w = 4;
    }
  }

  return periods;
}

function populateSelects() {
  // BU select
  const buSel = document.getElementById('cm-bu');
  BU_LIST.forEach(bu => {
    const o = document.createElement('option');
    o.value = bu;
    o.textContent = bu;
    buSel.appendChild(o);
  });

  // Daily BU select
  const dailyBuSel = document.getElementById('daily-bu');
  if (dailyBuSel) {
    BU_LIST.forEach(bu => {
      const o = document.createElement('option');
      o.value = bu;
      o.textContent = bu;
      dailyBuSel.appendChild(o);
    });
  }

  // Auto-detect period for CM
  const period = detectReportPeriod();
  state.cm.period = period;
  state.cm.apMonth = period.ap.month;
  state.cm.apWeek = period.ap.week;
  state.cm.isLocked = period.isLocked;
  state.cm.bu = buSel.value;

  // Update period display in topbar
  updateCMPeriodDisplay();

  // Dashboard: auto-detect period + populate history
  // FM Dashboard dùng cùng chu kỳ AP với CM Dashboard
  const dashPeriods = generatePeriodHistory(period);
  state.dashboard.periodHistory = dashPeriods;

  // Set dashboard mặc định = chu kỳ AP hiện tại
  state.dashboard.month = period.ap.month;
  state.dashboard.week = period.ap.week;

  // Populate period history dropdown
  const histSel = document.getElementById('dash-period-history');
  dashPeriods.forEach((p, i) => {
    const o = document.createElement('option');
    o.value = i;
    o.textContent = p.isCurrent ? `${p.label} (hiện tại)` : p.label;
    if (p.isCurrent) o.selected = true;
    histSel.appendChild(o);
  });

  // Update dashboard period display
  updateDashPeriodDisplay();

}

/** Cập nhật label kỳ báo cáo trên Dashboard */
function updateDashPeriodDisplay() {
  const el = document.getElementById('dash-current-period');
  if (!el) return;
  const { month, week } = state.dashboard;
  const [y, mo] = month.split('-');
  el.textContent = `Tháng ${parseInt(mo)}/${y} — Tuần ${week}`;
}

function onDashPeriodChange() {
  const histSel = document.getElementById('dash-period-history');
  const idx = parseInt(histSel.value);
  const periods = state.dashboard.periodHistory;
  if (!periods || !periods[idx]) return;

  const selected = periods[idx];
  state.dashboard.month = selected.month;
  state.dashboard.week = selected.week;
  updateDashPeriodDisplay();

  // Auto-load dữ liệu cho kỳ đã chọn
  runAnalysis();
}

function updateCMPeriodDisplay() {
  const period = state.cm.period;
  if (!period) return;

  // Update period labels
  const apPeriodEl = document.getElementById('cm-ap-period');
  if (apPeriodEl) apPeriodEl.textContent = period.ap.label;

  // Update lock status badge
  const lockStatusEl = document.getElementById('cm-lock-status');
  const lockIconEl = lockStatusEl ? lockStatusEl.querySelector('.lock-icon') : null;
  const lockTextEl = document.getElementById('cm-lock-text');

  if (lockStatusEl && lockTextEl && lockIconEl) {
    if (period.isLocked) {
      lockStatusEl.className = 'lock-status locked';
      lockIconEl.textContent = '🔒';
      lockTextEl.textContent = 'Đã khóa';
    } else if (period.deadlineLabel.includes('hôm nay')) {
      lockStatusEl.className = 'lock-status deadline-today';
      lockIconEl.textContent = '⚠️';
      lockTextEl.textContent = period.deadlineLabel;
    } else {
      lockStatusEl.className = 'lock-status';
      lockIconEl.textContent = '🔓';
      lockTextEl.textContent = period.deadlineLabel;
    }
  }

  // Apply lock to form
  applyLockState();
}

function applyLockState() {
  const isLocked = state.cm.isLocked;
  const viewCM = document.getElementById('view-cm');
  if (!viewCM) return;

  // Toggle form-locked class on both tab contents
  const tabAP = document.getElementById('tab-AP');
  const tabWR = document.getElementById('tab-WR');
  if (tabAP) tabAP.classList.toggle('form-locked', isLocked);
  if (tabWR) tabWR.classList.toggle('form-locked', isLocked);

  // Show/remove lock banner
  let banner = document.getElementById('cm-lock-banner');
  if (isLocked) {
    if (!banner) {
      banner = document.createElement('div');
      banner.id = 'cm-lock-banner';
      banner.className = 'lock-banner';
      // Insert after topbar
      const topbar = viewCM.querySelector('.view-topbar');
      if (topbar && topbar.nextSibling) {
        topbar.parentNode.insertBefore(banner, topbar.nextSibling);
      }
    }
    banner.innerHTML = `<span class="lock-banner-icon">🔒</span> ${escHtml(state.cm.period.lockMessage)}`;
  } else {
    if (banner) banner.remove();
  }
}

// ===================== CM TAB SWITCHING =====================
// Navigate from CM view to Daily/Discussion, passing BU context
function navigateFromCM(view) {
  if (view === 'daily' && state.cm.bu) {
    // Pre-set BU for daily view
    state.daily = state.daily || {};
    state.daily.pendingBU = state.cm.bu;
  }
  navigate(view);
}

function initTabs() {
  document.querySelectorAll('#view-cm .tab-btn[data-tab]').forEach(btn => {
    btn.addEventListener('click', () => {
      const tab = btn.dataset.tab;
      const tabContent = document.getElementById(`tab-${tab}`);
      if (!tabContent) return;
      // Only toggle tabs within view-cm
      document.querySelectorAll('#view-cm .tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('#view-cm .tab-content').forEach(c => c.classList.remove('active'));
      btn.classList.add('active');
      tabContent.classList.add('active');
      state.cm.activeTab = tab;
      // Hide AP-specific sections when on Analytics tab
      toggleAPSections(tab !== 'CM-ANA');
      // Init CM analytics when switching to that tab
      if (tab === 'CM-ANA') initCMAnalytics();
    });
  });
}

// Show/hide AP-specific sections (period, lock, FM notes, diagnostic)
function toggleAPSections(show) {
  // Period display + lock status + lock banner: always toggle
  ['cm-period-display', 'cm-lock-status', 'cm-lock-banner'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.style.display = show ? '' : 'none';
  });
  // FM notes + diagnostic: hide when Analytics; restore via their own logic when AP
  ['cm-fm-note-section', 'diagnostic-section'].forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      if (!show) el.style.display = 'none';
      // When returning to AP, let updateDiagnosticVisibility / loadFMNotes handle visibility
    }
  });
  if (show) {
    updateDiagnosticVisibility();
    // FM notes re-render handled by loadCMData already
  }
  // Also hide the topbar dividers that separate period/lock
  const topbar = document.querySelector('#view-cm .view-topbar');
  if (topbar) {
    topbar.querySelectorAll('.topbar-divider').forEach(d => d.style.display = show ? '' : 'none');
  }
}

// Navigate from Daily/Discussion to CM Analytical tab
function switchToCMAnalytical() {
  navigate('cm');
  // Simulate clicking the Analytical tab
  setTimeout(() => {
    const anaBtn = document.querySelector('#view-cm .tab-btn[data-tab="CM-ANA"]');
    if (anaBtn) anaBtn.click();
  }, 50);
}

// ===================== CM TOPBAR CHANGE =====================
function onCMSelectionChange() {
  state.cm.bu = document.getElementById('cm-bu').value;
  // month/week are auto-detected — no manual change needed
  updateDiagnosticVisibility();
  loadCMData();
  // Refresh CM Analytics if currently on that tab
  if (state.cm.activeTab === 'CM-ANA') {
    initCMAnalytics();
  }
}

async function loadCMData() {
  const { bu, apMonth, apWeek } = state.cm;
  if (!bu) return;

  if (!isConfigured()) {
    state.cm.apRows = AP_DEFAULT_ROWS.map(r => ({ ...r }));
    state.cm.wrRows = WR_DEFAULT_ROWS.map(r => ({ ...r }));
    state.cm.savedAt = { AP: null, WR: null };
    renderAPTable();
    renderWRTable();
    updateSavedIndicators();
    updateDiagnosticVisibility();
    return;
  }

  try {
    // Chỉ tải AP — không còn WR
    const rAP = await apiFetch('get', { tab: 'AP', bu, month: apMonth, week: apWeek });

    // Action Plan
    if (rAP.rows && rAP.rows.length > 0) {
      state.cm.apRows = rAP.rows;
      state.cm.savedAt.AP = rAP.saved_at || null;
    } else {
      state.cm.apRows = AP_DEFAULT_ROWS.map(r => ({ ...r }));
      state.cm.savedAt.AP = null;
    }

    // WR: giữ mặc định (không tải từ server)
    state.cm.wrRows = WR_DEFAULT_ROWS.map(r => ({ ...r }));
    state.cm.savedAt.WR = null;

    renderAPTable();
    renderWRTable();
    updateSavedIndicators();
    updateDiagnosticVisibility();
  } catch(e) {
    showToast('Lỗi tải dữ liệu: ' + e.message, 'error');
    state.cm.apRows = AP_DEFAULT_ROWS.map(r => ({ ...r }));
    state.cm.wrRows = WR_DEFAULT_ROWS.map(r => ({ ...r }));
    renderAPTable();
    renderWRTable();
    updateDiagnosticVisibility();
  }

  // Tải FM Note cho CM sau khi data đã sẵn sàng
  if (bu) await loadCMFMNote();
}

function updateSavedIndicators() {
  ['AP','WR'].forEach(tab => {
    const el = document.getElementById(`saved-indicator-${tab}`);
    if (!el) return;
    const ts = state.cm.savedAt[tab];
    if (ts) {
      el.textContent = `Đã lưu lúc ${fmt(ts)}`;
      el.className = 'saved-indicator show';
    } else {
      el.textContent = 'Chưa lưu';
      el.className = 'saved-indicator';
    }
  });
}

// ===================== ACTION PLAN TABLE =====================
function renderAPTable() {
  const tbody = document.getElementById('ap-tbody');
  tbody.innerHTML = '';

  state.cm.apRows.forEach((row, i) => {
    const tr = document.createElement('tr');
    tr.dataset.index = i;

    tr.innerHTML = `
      <td class="row-num">${i+1}</td>
      <td>
        <select class="w-md" onchange="updateAPRow(${i},'func',this.value)">
          <option value="GROWTH"${row.func==='GROWTH'?' selected':''}>GROWTH</option>
          <option value="OPTIMIZE"${row.func==='OPTIMIZE'?' selected':''}>OPTIMIZE</option>
          <option value="OPS"${row.func==='OPS'?' selected':''}>OPS</option>
        </select>
      </td>
      <td><input type="text" class="w-lg" value="${escHtml(row.chi_so||'')}" placeholder="VD: CR16%, L1 Leads, AOV..." onchange="updateAPRow(${i},'chi_so',this.value)"></td>
      <td><textarea class="w-xl" rows="2" placeholder="VD: CR chỉ đạt 8%, thấp hơn chuẩn 50%" onchange="updateAPRow(${i},'van_de',this.value)">${escHtml(row.van_de||'')}</textarea></td>
      <td>
        <select class="w-md" onchange="updateAPRow(${i},'muc_do',this.value)">
          <option value="">—</option>
          <option value="Nghiêm trọng"${row.muc_do==='Nghiêm trọng'?' selected':''}>Nghiêm trọng</option>
          <option value="Cần theo dõi"${row.muc_do==='Cần theo dõi'?' selected':''}>Cần theo dõi</option>
          <option value="Tín hiệu tốt"${row.muc_do==='Tín hiệu tốt'?' selected':''}>Tín hiệu tốt</option>
        </select>
      </td>
      <td><textarea class="w-xl" rows="3" placeholder="Phân tích nguyên nhân gốc rễ..." onchange="updateAPRow(${i},'root_cause',this.value)">${escHtml(row.root_cause||'')}</textarea></td>
      <td><textarea class="w-xl" rows="2" placeholder="Hành động cụ thể cần thực hiện" onchange="updateAPRow(${i},'key_action',this.value)">${escHtml(row.key_action||'')}</textarea></td>
      <td><textarea class="w-xl" rows="3" placeholder="Mô tả cách triển khai, các bước cụ thể..." onchange="updateAPRow(${i},'mo_ta_trien_khai',this.value)">${escHtml(row.mo_ta_trien_khai||'')}</textarea></td>
      <td><input type="text" class="w-md" value="${escHtml(row.target_do_luong||'')}" placeholder="VD: Đạt CR 16% tuần sau" onchange="updateAPRow(${i},'target_do_luong',this.value)"></td>
      <td><input type="date" class="w-date" value="${row.deadline||''}" onchange="updateAPRow(${i},'deadline',this.value)"></td>
      <td><input type="text" class="w-md" value="${escHtml(row.owner||'')}" placeholder="Người phụ trách" onchange="updateAPRow(${i},'owner',this.value)"></td>
      <td><input type="text" class="w-md" value="${escHtml(row.fm_support||'')}" placeholder="Cần hỗ trợ từ FM" onchange="updateAPRow(${i},'fm_support',this.value)"></td>
      <td>
        <select class="w-md" onchange="updateAPRow(${i},'status',this.value)">
          <option value="">—</option>
          <option value="Chưa bắt đầu"${row.status==='Chưa bắt đầu'?' selected':''}>Chưa bắt đầu</option>
          <option value="Đang thực hiện"${row.status==='Đang thực hiện'?' selected':''}>Đang thực hiện</option>
          <option value="Hoàn thành"${row.status==='Hoàn thành'?' selected':''}>Hoàn thành</option>
          <option value="Trì hoãn"${row.status==='Trì hoãn'?' selected':''}>Trì hoãn</option>
        </select>
      </td>
      <td><button class="btn-icon" title="Xóa dòng" onclick="deleteAPRow(${i})">✕</button></td>
    `;
    tbody.appendChild(tr);
  });
}

function calcGapVal(target, actual) {
  const t = parseFloat(target);
  const a = parseFloat(actual);
  if (isNaN(t) || isNaN(a)) return '—';
  return (a - t).toFixed(0);
}

function calcGapPctVal(target, actual) {
  const t = parseFloat(target);
  const a = parseFloat(actual);
  if (isNaN(t) || isNaN(a) || t === 0) return '—';
  return ((a - t) / Math.abs(t) * 100).toFixed(1) + '%';
}

function calcAchievementPct(target, actual) {
  const t = parseFloat(target);
  const a = parseFloat(actual);
  if (isNaN(t) || isNaN(a) || t === 0) return '—';
  return (a / t * 100).toFixed(1) + '%';
}

function updateAPGapDisplay(i) {
  const row = state.cm.apRows[i];
  const gapEl = document.getElementById(`ap-gap-${i}`);
  const pctEl = document.getElementById(`ap-gappct-${i}`);
  if (!gapEl || !pctEl) return;

  const gap = parseFloat(calcGapVal(row.target_prev, row.actual_prev));
  const gapTxt = calcGapVal(row.target_prev, row.actual_prev);
  const pctTxt = calcGapPctVal(row.target_prev, row.actual_prev);

  gapEl.textContent = gapTxt;
  pctEl.textContent = pctTxt;

  if (!isNaN(gap)) {
    gapEl.className = `auto-calc ${gap < 0 ? 'negative' : 'positive'}`;
    pctEl.className = `auto-calc ${gap < 0 ? 'negative' : 'positive'}`;
  } else {
    gapEl.className = 'auto-calc';
    pctEl.className = 'auto-calc';
  }
}

function updateAPRow(i, field, value) {
  if (!state.cm.apRows[i]) return;
  state.cm.apRows[i][field] = value;
}

function calcAPGap(i) {
  updateAPGapDisplay(i);
}

function deleteAPRow(i) {
  state.cm.apRows.splice(i, 1);
  renderAPTable();
}

function addAPRow() {
  state.cm.apRows.push({ func: 'GROWTH', chi_so: '', van_de: '', muc_do: '', root_cause: '', key_action: '', mo_ta_trien_khai: '', target_do_luong: '', deadline: '', owner: '', fm_support: '', status: '' });
  renderAPTable();
  const wrapper = document.querySelector('#tab-AP .form-table-wrapper');
  if (wrapper) wrapper.scrollTop = wrapper.scrollHeight;
}

// ===================== WEEKLY REPORT TABLE =====================
function renderWRTable() {
  const tbody = document.getElementById('wr-tbody');
  tbody.innerHTML = '';

  state.cm.wrRows.forEach((row, i) => {
    const tr = document.createElement('tr');
    const isTotal = row.kpi && row.kpi.toUpperCase().includes('TỔNG');
    if (isTotal) {
      tr.style.fontWeight = '700';
      tr.style.background = '#f0f0f0';
    }

    tr.innerHTML = `
      <td class="row-num">${i+1}</td>
      <td style="font-weight:600;white-space:nowrap;padding:6px 8px">${escHtml(row.kpi||'')}</td>
      <td><input type="number" class="w-num" value="${row.target||''}" placeholder="0" oninput="updateWRRow(${i},'target',this.value); calcWRAchievement(${i})"></td>
      <td><input type="number" class="w-num" value="${row.actual||''}" placeholder="0" oninput="updateWRRow(${i},'actual',this.value); calcWRAchievement(${i})"></td>
      <td class="auto-calc" id="wr-gap-${i}">${calcGapVal(row.target,row.actual)}</td>
      <td class="auto-calc" id="wr-ach-${i}">${calcAchievementPct(row.target,row.actual)}</td>
      <td><textarea rows="2" style="min-width:200px" placeholder="Trạng thái / Ghi chú..." onchange="updateWRRow(${i},'notes',this.value)">${escHtml(row.notes||'')}</textarea></td>
    `;
    tbody.appendChild(tr);
    updateWRDisplay(i);
  });
}

function updateWRDisplay(i) {
  const row = state.cm.wrRows[i];
  const gapEl = document.getElementById(`wr-gap-${i}`);
  const achEl = document.getElementById(`wr-ach-${i}`);
  if (!gapEl || !achEl) return;

  const gap = parseFloat(calcGapVal(row.target, row.actual));
  gapEl.textContent = calcGapVal(row.target, row.actual);
  achEl.textContent = calcAchievementPct(row.target, row.actual);

  if (!isNaN(gap)) {
    gapEl.className = `auto-calc ${gap < 0 ? 'negative' : 'positive'}`;
    achEl.className = `auto-calc ${gap < 0 ? 'negative' : 'positive'}`;
  } else {
    gapEl.className = 'auto-calc';
    achEl.className = 'auto-calc';
  }
}

function updateWRRow(i, field, value) {
  if (!state.cm.wrRows[i]) return;
  state.cm.wrRows[i][field] = value;
  // Re-check diagnostic button khi target/actual thay đổi
  if (field === 'target' || field === 'actual') {
    updateDiagnosticButton();
  }
}

function calcWRAchievement(i) {
  updateWRDisplay(i);
}

// ===================== SAVE =====================
async function saveTab(tab) {
  const { bu, isLocked } = state.cm;
  if (!bu) { showToast('Vui lòng chọn BU trước khi lưu', 'error'); return; }
  if (isLocked) { showToast('Đã khóa chỉnh sửa. Không thể lưu sau Thứ 4.', 'error'); return; }

  // Dùng AP month/week cho tất cả (WR đã loại bỏ)
  const month = state.cm.apMonth;
  const week = state.cm.apWeek;

  if (!isConfigured()) {
    showToast('Chưa cấu hình Google Apps Script URL. Xem SETUP.md để biết cách thiết lập.', 'error');
    return;
  }

  let rows;
  if (tab === 'AP') rows = state.cm.apRows;
  else if (tab === 'WR') rows = state.cm.wrRows;

  const tabLabel = tab === 'AP' ? 'Action Plan' : 'Weekly Report';
  const btnSave = document.querySelector(`#tab-${tab} .btn-primary`);
  if (btnSave) { btnSave.disabled = true; btnSave.textContent = 'Đang lưu...'; }

  try {
    const result = await apiPost('save', tab, { bu, month, week, rows });
    if (result.success) {
      state.cm.savedAt[tab] = result.saved_at;
      updateSavedIndicators();
      showToast(`✓ Đã lưu ${tabLabel} — ${result.rows_saved} dòng lúc ${fmtNow()}`);
    } else {
      showToast('Lỗi lưu dữ liệu: ' + (result.error || 'unknown'), 'error');
    }
  } catch(e) {
    showToast('Lỗi kết nối: ' + e.message, 'error');
  } finally {
    if (btnSave) {
      btnSave.disabled = false;
      btnSave.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg> Lưu ${tabLabel}`;
    }
  }
}

// ===================== SCORING — ACTION PLAN (70 điểm) =====================
function scoreAP(apRows) {
  let score = 0;
  const breakdown = {};

  if (!apRows || apRows.length === 0) return { score: 0, breakdown: { 'Không có dữ liệu': 0 } };

  const n = apRows.length;

  // 1. Nhận diện vấn đề (max 20): % dòng có chi_so + van_de đều có nội dung
  let issueIdentified = 0;
  apRows.forEach(row => {
    const hasChiSo = row.chi_so && String(row.chi_so).trim() !== '';
    const hasVanDe = row.van_de && String(row.van_de).trim() !== '';
    if (hasChiSo && hasVanDe) issueIdentified++;
  });
  const issueScore = Math.round((issueIdentified / n) * 20);
  breakdown['Nhận diện vấn đề (max 20)'] = issueScore;
  score += issueScore;

  // 2. Chất lượng Root Cause (max 20): % dòng có root_cause > 20 ký tự
  let rcGood = 0;
  apRows.forEach(row => {
    if (row.root_cause && String(row.root_cause).trim().length > 20) rcGood++;
  });
  const rcScore = Math.round((rcGood / n) * 20);
  breakdown['Chất lượng Root Cause (max 20)'] = rcScore;
  score += rcScore;

  // 3. Chất lượng Key Action (max 15): % dòng có key_action VÀ target_do_luong
  let actionQuality = 0;
  apRows.forEach(row => {
    const hasAction = row.key_action && String(row.key_action).trim() !== '';
    const hasTarget = row.target_do_luong && String(row.target_do_luong).trim() !== '';
    if (hasAction && hasTarget) actionQuality++;
  });
  const actionScore = Math.round((actionQuality / n) * 15);
  breakdown['Chất lượng Key Action (max 15)'] = actionScore;
  score += actionScore;

  // 4. Tiêu chí SMART (max 10): % dòng có target_do_luong + deadline + owner đều có
  let smartCount = 0;
  apRows.forEach(row => {
    const hasTarget = row.target_do_luong && String(row.target_do_luong).trim() !== '';
    const hasDeadline = row.deadline && String(row.deadline).trim() !== '';
    const hasOwner = row.owner && String(row.owner).trim() !== '';
    if (hasTarget && hasDeadline && hasOwner) smartCount++;
  });
  const smartScore = Math.round((smartCount / n) * 10);
  breakdown['Tiêu chí SMART (max 10)'] = smartScore;
  score += smartScore;

  // 5. Đa dạng Function (max 5)
  const funcs = new Set(apRows.map(r => r.func).filter(Boolean));
  const divScore = funcs.size >= 3 ? 5 : funcs.size >= 2 ? 3 : funcs.size >= 1 ? 1 : 0;
  breakdown['Đa dạng Function (max 5)'] = divScore;
  score += divScore;

  return { score: Math.min(score, 70), breakdown };
}

// ===================== SCORING — WEEKLY REPORT (30 điểm) =====================
function scoreWR(wrRows) {
  let score = 0;
  const breakdown = {};

  if (!wrRows || wrRows.length === 0) return { score: 0, breakdown: { 'Không có dữ liệu': 0 } };

  const n = wrRows.length;

  // 1. Tỷ lệ điền KPI (max 15): % dòng có target + actual đều có
  let filledKPI = 0;
  wrRows.forEach(row => {
    const hasTarget = row.target !== undefined && row.target !== '' && row.target !== null;
    const hasActual = row.actual !== undefined && row.actual !== '' && row.actual !== null;
    if (hasTarget && hasActual) filledKPI++;
  });
  const fillScore = Math.round((filledKPI / n) * 15);
  breakdown['Tỷ lệ điền KPI (max 15)'] = fillScore;
  score += fillScore;

  // 2. Ghi chú khi Gap lớn (max 15): % dòng gap > 5% có notes > 10 ký tự
  let gapCount = 0, noteGood = 0;
  wrRows.forEach(row => {
    const t = parseFloat(row.target);
    const a = parseFloat(row.actual);
    if (!isNaN(t) && !isNaN(a) && t !== 0) {
      const pct = Math.abs((a - t) / t);
      if (pct > 0.05) {
        gapCount++;
        if (row.notes && String(row.notes).trim().length > 10) noteGood++;
      }
    }
  });
  const noteScore = gapCount > 0
    ? Math.round((noteGood / gapCount) * 15)
    : 15; // không có gap lớn = điểm ghi chú tối đa
  breakdown['Ghi chú khi Gap lớn (max 15)'] = noteScore;
  score += noteScore;

  return { score: Math.min(score, 30), breakdown };
}

function getRating(total) {
  if (total >= 80) return { label: 'Tốt', cls: 'tot' };
  if (total >= 60) return { label: 'Đầy đủ', cls: 'day-du' };
  if (total >= 40) return { label: 'Hời hợt', cls: 'hoi-hot' };
  return { label: 'Cần cải thiện', cls: 'can-cb' };
}

// ===================== BU HEALTH STATUS =====================
// Trạng thái sức khỏe BU dựa trên WR doanh số
function computeBUHealth(wrRows, week, paceMap) {
  // Nếu không có WR data thì là xám
  if (!wrRows || wrRows.length === 0) {
    return { healthStatus: 'xam', healthLabel: 'Chưa nộp WR', healthCls: 'health-xam', revenuePct: 0, revenueRows: [] };
  }

  const activePace = paceMap || WEEK_PACE;
  const pace = activePace[week] || activePace[1];

  // Lấy các dòng doanh số: Doanh số N1, Doanh số N2, Doanh số N3, TỔNG DOANH SỐ BU
  const revenueRows = wrRows.filter(row => {
    const k = (row.kpi || '').trim();
    return k.includes('Doanh số') || k.toUpperCase().includes('TỔNG');
  });

  if (revenueRows.length === 0) {
    // Có WR nhưng không có dòng doanh số nào được điền
    return { healthStatus: 'xam', healthLabel: 'Chưa nộp WR', healthCls: 'health-xam', revenuePct: 0, revenueRows: [] };
  }

  // Kiểm tra xem có ít nhất 1 dòng doanh số có target
  const hasAnyTarget = revenueRows.some(row => {
    const t = parseFloat(row.target);
    return !isNaN(t) && t > 0;
  });

  if (!hasAnyTarget) {
    return { healthStatus: 'xam', healthLabel: 'Chưa nộp WR', healthCls: 'health-xam', revenuePct: 0, revenueRows: [] };
  }

  // Tính % đạt cho từng dòng doanh số (actual / (target * pace))
  const rowAnalysis = revenueRows.map(row => {
    const target = parseFloat(row.target);
    const actual = parseFloat(row.actual);
    const kpi = (row.kpi || '').trim();
    const isTotal = kpi.toUpperCase().includes('TỔNG');

    if (isNaN(target) || target <= 0) {
      return { kpi, target: row.target, actual: row.actual, pct: null, paceOk: null, isTotal };
    }
    if (isNaN(actual)) {
      return { kpi, target, actual: 0, pct: 0, paceOk: false, isTotal, expectedPace: Math.round(target * pace) };
    }

    const pct = actual / target; // tỷ lệ actual/target (0..1+)
    const paceOk = pct >= pace;  // đạt pace nếu actual >= target*pace
    return { kpi, target, actual, pct, paceOk, isTotal, expectedPace: Math.round(target * pace) };
  });

  // Lấy dòng tổng doanh số BU
  const totalRow = rowAnalysis.find(r => r.isTotal);
  const totalPct = totalRow ? totalRow.pct : null;

  // Các dòng doanh số cá nhân (không phải tổng)
  const subRows = rowAnalysis.filter(r => !r.isTotal && r.pct !== null);

  let healthStatus, healthLabel, healthCls;

  // Xác định trạng thái
  const allMeetPace = rowAnalysis.filter(r => r.pct !== null).every(r => r.paceOk);
  const anyBelow50 = rowAnalysis.filter(r => r.pct !== null).some(r => r.pct < pace * 0.5);
  const totalBelowPace80 = totalPct !== null && totalPct < pace * 0.8;

  if (allMeetPace) {
    healthStatus = 'xanh';
    healthLabel = 'Xanh';
    healthCls = 'health-xanh';
  } else if (anyBelow50 || totalBelowPace80) {
    healthStatus = 'do';
    healthLabel = 'Đỏ';
    healthCls = 'health-do';
  } else {
    healthStatus = 'vang';
    healthLabel = 'Vàng';
    healthCls = 'health-vang';
  }

  // Tính % đạt tổng (để hiển trong bảng)
  const revenuePct = totalPct !== null ? Math.round(totalPct * 100) : (
    subRows.length > 0 ? Math.round(subRows.reduce((acc, r) => acc + r.pct, 0) / subRows.length * 100) : 0
  );

  return { healthStatus, healthLabel, healthCls, revenuePct, revenueRows: rowAnalysis, pace };
}

// ===================== FUNCTION FILTER =====================
function setFuncFilter(func) {
  state.dashboard.funcFilter = func;

  // Cập nhật nút active
  document.querySelectorAll('.func-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.func === func);
  });

  // Tính lại scores với filter mới
  if (state.dashboard.allData) {
    state.dashboard.scores = computeScores(state.dashboard.allData);

    // Re-compute run rates with new function filter
    const month = state.dashboard.month;
    const kpiTargets = state.dashboard._kpiTargets && state.dashboard._kpiTargets[month];
    if (kpiTargets && kpiTargets.length > 0) {
      const pacePct = calcPacePct(month);
      const kpiMap = {};
      kpiTargets.forEach(t => { kpiMap[t.bu] = t; });
      state.dashboard.scores.forEach(s => {
        const kpi = kpiMap[s.bu];
        if (!kpi) return;
        let rev, tgt;
        if (func === 'GROWTH')   { rev = s.revBU.n3; tgt = Number(kpi.kpi_n3) || 0; }
        else if (func === 'OPTIMIZE') { rev = s.revBU.n1; tgt = Number(kpi.kpi_n1) || 0; }
        else if (func === 'OPS')      { rev = s.revBU.n2; tgt = Number(kpi.kpi_n2) || 0; }
        else                          { rev = s.revBU.total; tgt = Number(kpi.kpi_total) || 0; }
        s.kpiTarget = kpi;
        s._runRev = rev;
        s._runTgt = tgt;
        s.runRatePct = 0; s.runRateDiff = 0; s.runRateStatus = null;
        if (tgt > 0) {
          s.runRatePct = (rev / tgt) * 100;
          s.runRateDiff = s.runRatePct - pacePct;
          s.runRateStatus = getRunRateStatus(s.runRatePct, pacePct);
        }
      });
    }

    state.dashboard.filteredScores = [...state.dashboard.scores];
    renderDashboard();
  }
}

function computeScores(allData) {
  const funcFilter = state.dashboard.funcFilter;
  const week = state.dashboard.week || 1;
  const month = state.dashboard.month;
  const cfgData = getActiveConfig(month);
  const pace = getWeekPace(cfgData);

  // Daily report data (nếu đã load)
  const dailyRows = state.dashboard._dailyData || [];

  // Tính số ngày kỳ vọng (từ đầu tháng đến hôm qua, trừ CN)
  const today = new Date();
  const [yyyy, mm] = month.split('-').map(Number);
  const monthStart = new Date(yyyy, mm - 1, 1);
  const yesterday = new Date(today);
  yesterday.setDate(yesterday.getDate() - 1);
  const expectedDays = [];
  for (let d = new Date(monthStart); d <= yesterday; d.setDate(d.getDate() + 1)) {
    const dow = d.getDay(); // 0=Sun
    if (dow !== 0) { // Trừ Chủ nhật
      expectedDays.push(d.toISOString().slice(0, 10));
    }
  }

  // Group daily data by BU (dates + revenue per source)
  const dailyByBU = {};
  const revByBU = {}; // { buName: { n1, n2, n3, total } }
  dailyRows.forEach(row => {
    const bu = row.bu;
    if (!dailyByBU[bu]) dailyByBU[bu] = new Set();
    if (row.date) dailyByBU[bu].add(row.date);
    // Aggregate revenue per BU per source
    if (!revByBU[bu]) revByBU[bu] = { n1: 0, n2: 0, n3: 0, total: 0 };
    const r1 = parseInt(row.revenue_n1) || 0;
    const r2 = parseInt(row.revenue_n2) || 0;
    const r3 = parseInt(row.revenue_n3) || 0;
    revByBU[bu].n1 += r1;
    revByBU[bu].n2 += r2;
    revByBU[bu].n3 += r3;
    revByBU[bu].total += r1 + r2 + r3;
  });

  return BU_LIST.map(buName => {
    const buData = allData[buName] || { AP: [], WR: [] };

    // Filter AP rows theo function nếu cần
    let apRows = buData.AP || [];
    if (funcFilter !== 'ALL') {
      apRows = apRows.filter(r => r.func === funcFilter);
    }

    const ap = scoreAP(apRows);
    const wr = scoreWR(buData.WR);
    const total = ap.score + wr.score;
    const rating = getRatingFromConfig(total, cfgData);

    const hasAP = apRows.length > 0;
    const hasWR = buData.WR && buData.WR.length > 0;

    // Tính health status từ WR data (with config pace)
    const health = computeBUHealth(buData.WR, week, pace);

    // Daily report: ngày nào thiếu?
    const buDailyDates = dailyByBU[buName] || new Set();
    const dailyMissingDays = expectedDays.filter(d => !buDailyDates.has(d));
    const dailyFilledCount = expectedDays.filter(d => buDailyDates.has(d)).length;
    const dailyExpectedCount = expectedDays.length;

    return {
      bu: buName,
      region: getRegion(buName),
      apScore: ap.score, apMax: 70, apBreakdown: ap.breakdown,
      wrScore: wr.score, wrMax: 30, wrBreakdown: wr.breakdown,
      total, rating: rating.label, ratingCls: rating.cls,
      hasAP, hasWR,
      apRowCount: apRows.length,
      apTotalRowCount: (buData.AP || []).length,
      healthStatus: health.healthStatus,
      healthLabel: health.healthLabel,
      healthCls: health.healthCls,
      revenuePct: health.revenuePct,
      revenueRows: health.revenueRows,
      pace: health.pace,
      dailyMissingDays,
      dailyMissing: dailyMissingDays.length,
      dailyFilled: dailyFilledCount,
      dailyExpected: dailyExpectedCount,
      // Revenue MTD from daily data (VNĐ) — per source + total
      revBU: revByBU[buName] || { n1: 0, n2: 0, n3: 0, total: 0 },
      revenueMTD: (revByBU[buName] || { total: 0 }).total,
      kpiTarget: null, // will be filled after KPI load
      runRatePct: 0,
      runRateDiff: 0,
      runRateStatus: null
    };
  });
}

// ===================== DASHBOARD =====================
async function runAnalysis() {
  // month/week đã được set từ auto-detect hoặc period history selector
  const { month, week } = state.dashboard;

  const btn = document.getElementById('btn-analyze');
  const btnOrigHTML = btn.innerHTML;
  btn.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" class="spin-icon"><path d="M21.5 2v6h-6M2.5 22v-6h6M2 11.5a10 10 0 0 1 18.8-4.3M22 12.5a10 10 0 0 1-18.8 4.2"/></svg> Đang cập nhật...';
  btn.disabled = true;

  if (!isConfigured()) {
    showToast('Chưa cấu hình Google Apps Script URL. Xem SETUP.md để biết cách thiết lập.', 'error');
    btn.innerHTML = btnOrigHTML;
    btn.disabled = false;
    return;
  }

  try {
    // Preload config for this month
    btn.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" class="spin-icon"><path d="M21.5 2v6h-6M2.5 22v-6h6M2 11.5a10 10 0 0 1 18.8-4.3M22 12.5a10 10 0 0 1-18.8 4.2"/></svg> Tải cấu hình...';
    await preloadConfigForMonth(month);

    // Fetch AP data — dùng trực tiếp month/week từ dashboard (= chu kỳ AP)
    btn.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" class="spin-icon"><path d="M21.5 2v6h-6M2.5 22v-6h6M2 11.5a10 10 0 0 1 18.8-4.3M22 12.5a10 10 0 0 1-18.8 4.2"/></svg> Tải AP (${month} W${week})...';
    const apData = await apiFetch('all_data', { month, week });

    // Build allData chỉ từ AP (không cần WR nữa)
    const allData = {};
    Object.keys(apData).forEach(bu => {
      allData[bu] = {
        AP: (apData[bu] && apData[bu].AP) || [],
        WR: []
      };
    });

    state.dashboard.allData = allData;

    // Load Daily Report data cho tháng hiện tại
    btn.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" class="spin-icon"><path d="M21.5 2v6h-6M2.5 22v-6h6M2 11.5a10 10 0 0 1 18.8-4.3M22 12.5a10 10 0 0 1-18.8 4.2"/></svg> Tải Daily...';
    try {
      const dailyResult = await apiFetch('getDailyAll', { month });
      state.dashboard._dailyData = dailyResult.data || [];
    } catch(e) {
      console.warn('Could not load daily data:', e.message);
      state.dashboard._dailyData = [];
    }

    state.dashboard.scores = computeScores(allData);

    // Load KPI targets and compute run rates
    try {
      const kpiTargets = await loadKPITargets(month);
      if (kpiTargets && kpiTargets.length > 0) {
        const pacePct = calcPacePct(month);
        const kpiMap = {};
        kpiTargets.forEach(t => { kpiMap[t.bu] = t; });
        const func = state.dashboard.funcFilter;
        state.dashboard.scores.forEach(s => {
          const kpi = kpiMap[s.bu];
          if (!kpi) return;
          // Pick revenue & target based on function filter
          let rev, tgt;
          if (func === 'GROWTH')   { rev = s.revBU.n3; tgt = Number(kpi.kpi_n3) || 0; }
          else if (func === 'OPTIMIZE') { rev = s.revBU.n1; tgt = Number(kpi.kpi_n1) || 0; }
          else if (func === 'OPS')      { rev = s.revBU.n2; tgt = Number(kpi.kpi_n2) || 0; }
          else                          { rev = s.revBU.total; tgt = Number(kpi.kpi_total) || 0; }
          s.kpiTarget = kpi;
          s._runRev = rev;
          s._runTgt = tgt;
          if (tgt > 0) {
            s.runRatePct = (rev / tgt) * 100;
            s.runRateDiff = s.runRatePct - pacePct;
            s.runRateStatus = getRunRateStatus(s.runRatePct, pacePct);
          }
        });
        state.dashboard._pacePct = pacePct;
      }
    } catch(e) { console.warn('KPI targets load error:', e.message); }

    state.dashboard.filteredScores = [...state.dashboard.scores];

    renderDashboard();
    loadFMNotesSummary().then(() => renderScoreTable());
    showToast(`Đã cập nhật — Tháng ${month} Tuần ${week}`);
  } catch(e) {
    showToast('Lỗi tải dữ liệu dashboard: ' + e.message, 'error');
  } finally {
    btn.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21.5 2v6h-6M2.5 22v-6h6M2 11.5a10 10 0 0 1 18.8-4.3M22 12.5a10 10 0 0 1-18.8 4.2"/></svg> Cập nhật';
    btn.disabled = false;
  }
}

function renderDashboard() {
  renderSummaryCards();
  renderQualityChart();
  renderMTDRunRate();
  applyDashFilters();
}

function renderSummaryCards() {
  const scores = state.dashboard.scores;
  const total = scores.length;

  const apDone = scores.filter(s => s.hasAP).length;
  const apMissing = total - apDone;
  const dailyOk = scores.filter(s => s.dailyMissing === 0 && s.dailyExpected > 0).length;
  const dailyGap = scores.filter(s => s.dailyMissing > 0).length;

  document.getElementById('card-total').textContent = total;
  document.getElementById('card-ap-done').textContent = apDone;
  document.getElementById('card-ap-missing').textContent = apMissing;
  document.getElementById('card-daily-ok').textContent = dailyOk;
  document.getElementById('card-daily-gap').textContent = dailyGap;
}

function renderQualityChart() {
  const scores = state.dashboard.scores;
  const total = scores.length || 1;

  // AP compliance
  const apDone = scores.filter(s => s.hasAP).length;
  const apPct = Math.round(apDone / total * 100);
  const apBar = document.getElementById('cbar-ap');
  if (apBar) apBar.style.width = apPct + '%';
  const apPctEl = document.getElementById('cbar-ap-pct');
  if (apPctEl) apPctEl.textContent = `${apDone}/${total} (${apPct}%)`;

  // Daily compliance
  const dailyOk = scores.filter(s => s.dailyMissing === 0 && s.dailyExpected > 0).length;
  const dailyPct = Math.round(dailyOk / total * 100);
  const dailyBar = document.getElementById('cbar-daily');
  if (dailyBar) dailyBar.style.width = dailyPct + '%';
  const dailyPctEl = document.getElementById('cbar-daily-pct');
  if (dailyPctEl) dailyPctEl.textContent = `${dailyOk}/${total} (${dailyPct}%)`;
}

/** Render MTD Run Rate overview bars (total + N1/N2/N3) */
function renderMTDRunRate() {
  const section = document.getElementById('mtd-runrate-section');
  if (!section) return;

  const scores = state.dashboard.scores;
  const month = state.dashboard.month;
  const pacePct = state.dashboard._pacePct || calcPacePct(month || new Date().toISOString().slice(0, 7));

  // Sum revenue MTD and KPI targets across all BUs
  let revTotal = 0, revN1 = 0, revN2 = 0, revN3 = 0;
  let kpiTotal = 0, kpiN1 = 0, kpiN2 = 0, kpiN3 = 0;
  let hasKPI = false;

  scores.forEach(s => {
    if (s.revBU) {
      revN1 += s.revBU.n1 || 0;
      revN2 += s.revBU.n2 || 0;
      revN3 += s.revBU.n3 || 0;
      revTotal += s.revBU.total || 0;
    }
    if (s.kpiTarget) {
      hasKPI = true;
      kpiN1 += Number(s.kpiTarget.kpi_n1) || 0;
      kpiN2 += Number(s.kpiTarget.kpi_n2) || 0;
      kpiN3 += Number(s.kpiTarget.kpi_n3) || 0;
      kpiTotal += Number(s.kpiTarget.kpi_total) || 0;
    }
  });

  if (!hasKPI) {
    section.style.display = 'none';
    return;
  }
  section.style.display = '';

  // Pace label
  const paceLabel = document.getElementById('mtd-pace-label');
  if (paceLabel) paceLabel.textContent = `(Pace k\u1EF3 v\u1ECDng: ${pacePct.toFixed(0)}%)`;

  // Helper to render a single run rate bar
  function setRRBar(key, rev, kpi) {
    const bar = document.getElementById('rr-bar-' + key);
    const pctEl = document.getElementById('rr-pct-' + key);
    const paceEl = document.getElementById('rr-pace-' + key);
    if (!bar || !pctEl) return;

    const pct = kpi > 0 ? (rev / kpi * 100) : 0;
    const barWidth = Math.min(pct, 100);
    bar.style.width = barWidth + '%';

    // Color based on run rate status
    const diff = pct - pacePct;
    if (key !== 'total') {
      // Keep source colors from HTML
    } else {
      bar.style.background = diff >= 5 ? 'var(--green)' : diff >= -5 ? 'var(--blue)' : diff >= -15 ? 'var(--orange)' : 'var(--red)';
    }

    const revFmt = fmtDops(rev, 'money');
    const kpiFmt = fmtDops(kpi, 'money');
    const status = getRunRateStatus(pct, pacePct);
    pctEl.innerHTML = `<strong>${revFmt}</strong> / ${kpiFmt} <span class="status-badge ${status.cls}" style="font-size:10px;padding:1px 6px;margin-left:4px">${status.icon} ${pct.toFixed(0)}%</span>`;

    // Pace marker
    if (paceEl) {
      paceEl.style.left = Math.min(pacePct, 100) + '%';
      paceEl.title = `Pace: ${pacePct.toFixed(0)}%`;
    }
  }

  setRRBar('total', revTotal, kpiTotal);
  setRRBar('n1', revN1, kpiN1);
  setRRBar('n2', revN2, kpiN2);
  setRRBar('n3', revN3, kpiN3);
}

function applyDashFilters() {
  const searchVal = document.getElementById('filter-search').value.toLowerCase();
  const statusVal = document.getElementById('filter-rating').value;
  const regionVal = document.getElementById('filter-region').value;

  state.dashboard.filteredScores = state.dashboard.scores.filter(s => {
    if (searchVal && !s.bu.toLowerCase().includes(searchVal)) return false;
    if (statusVal === 'ap-done' && !s.hasAP) return false;
    if (statusVal === 'ap-missing' && s.hasAP) return false;
    if (statusVal === 'daily-ok' && (s.dailyMissing > 0 || s.dailyExpected === 0)) return false;
    if (statusVal === 'daily-gap' && s.dailyMissing === 0) return false;
    if (regionVal && s.region !== regionVal) return false;
    return true;
  });

  const countEl = document.getElementById('filter-count');
  if (countEl) countEl.textContent = `${state.dashboard.filteredScores.length} BU`;

  renderScoreTable();
}

function renderScoreTable() {
  const tbody = document.getElementById('score-tbody');
  const { filteredScores, sortCol, sortDir } = state.dashboard;

  // Sắp xếp: mặc định đỏ trước, vàng, xám, xanh
  const healthOrder = { do: 0, vang: 1, xam: 2, xanh: 3 };

  const sorted = [...filteredScores].sort((a, b) => {
    // Nếu đang sắp xếp theo health, dùng healthOrder
    if (sortCol === 'healthStatus') {
      const ha = healthOrder[a.healthStatus] ?? 2;
      const hb = healthOrder[b.healthStatus] ?? 2;
      if (ha !== hb) return sortDir === 'asc' ? ha - hb : hb - ha;
      return b.total - a.total;
    }

    let va = a[sortCol], vb = b[sortCol];
    if (typeof va === 'string') va = va.toLowerCase();
    if (typeof vb === 'string') vb = vb.toLowerCase();
    if (va < vb) return sortDir === 'asc' ? -1 : 1;
    if (va > vb) return sortDir === 'asc' ? 1 : -1;
    return 0;
  });

  if (sorted.length === 0) {
    tbody.innerHTML = `<tr><td colspan="12" style="text-align:center;padding:32px;color:#888">Đang tải...</td></tr>`;
    return;
  }

  const fmNoteCache = state.dashboard._fmNoteCache || {};

  tbody.innerHTML = sorted.map((s, idx) => {
    const noteCache = fmNoteCache[s.bu];
    const hasNotes = noteCache && noteCache.total > 0;
    const unreadCount = noteCache ? noteCache.unread : 0;
    let noteIcon;
    if (hasNotes && unreadCount > 0) {
      noteIcon = `<span title="${noteCache.total} ghi chú, ${unreadCount} chờ CM" style="font-size:14px;position:relative">📝<span class="note-badge-count">${unreadCount}</span></span>`;
    } else if (hasNotes) {
      noteIcon = `<span title="${noteCache.total} ghi chú (đã đọc hết)" style="font-size:14px">📝</span>`;
    } else {
      noteIcon = `<span style="color:#ddd;font-size:14px">📝</span>`;
    }

    // AP status badge
    const apBadge = s.hasAP
      ? `<span class="status-badge status-done">✓ Đã làm</span>`
      : `<span class="status-badge status-missing">✗ Chưa làm</span>`;

    // Daily report status
    let dailyCell;
    if (s.dailyExpected === 0) {
      dailyCell = `<span style="color:#aaa;font-size:11px">—</span>`;
    } else if (s.dailyMissing === 0) {
      dailyCell = `<span class="status-badge status-done">✓ Đầy đủ <small>${s.dailyFilled}/${s.dailyExpected}</small></span>`;
    } else {
      // Hiển thị số ngày thiếu và tooltip với danh sách ngày
      const missingList = s.dailyMissingDays.map(d => {
        const dt = new Date(d + 'T00:00:00');
        return dt.toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit' });
      }).join(', ');
      const severity = s.dailyMissing >= 5 ? 'status-alert' : 'status-warn';
      dailyCell = `<span class="status-badge ${severity}" title="Thiếu: ${missingList}">Thiếu ${s.dailyMissing} ngày <small>${s.dailyFilled}/${s.dailyExpected}</small></span>`;
    }

    return `
    <tr data-bu="${escHtml(s.bu)}" onclick="selectBURow('${escHtml(s.bu)}')" class="${state.dashboard.selectedBU === s.bu ? 'selected' : ''}">
      <td style="color:#888;font-size:11px">${idx+1}</td>
      <td><strong>${escHtml(s.bu)}</strong></td>
      <td><span class="region-badge">${escHtml(s.region)}</span></td>
      <td class="text-center">${apBadge}</td>
      <td>${dailyCell}</td>
      <td>${(() => {
        if (!s.kpiTarget || !s.runRateStatus) return '<span style="color:#ccc;font-size:11px">—</span>';
        const revFmt = fmtDops(s._runRev || s.revenueMTD, 'money');
        const tgtFmt = fmtDops(s._runTgt || s.kpiTarget.kpi_total, 'money');
        const pctFmt = s.runRatePct.toFixed(0) + '%';
        const rs = s.runRateStatus;
        return `<div style="line-height:1.4">
          <div style="font-weight:700;font-size:13px">${revFmt} <small style="font-weight:400;color:var(--gray-400)">/ ${tgtFmt}</small></div>
          <span class="status-badge ${rs.cls}" style="font-size:11px;padding:2px 8px">${rs.icon} ${rs.label} <small>${pctFmt}</small></span>
        </div>`;
      })()}</td>
      <td style="text-align:center">${noteIcon}</td>
      <td style="color:#888;font-size:11px;text-align:center">›</td>
    </tr>
  `;
  }).join('');
}

function selectBURow(bu) {
  state.dashboard.selectedBU = bu;
  document.querySelectorAll('#score-tbody tr').forEach(tr => {
    tr.classList.toggle('selected', tr.dataset.bu === bu);
  });
  openDetailPanel(bu);
}

function sortTable(col) {
  if (state.dashboard.sortCol === col) {
    state.dashboard.sortDir = state.dashboard.sortDir === 'asc' ? 'desc' : 'asc';
  } else {
    state.dashboard.sortCol = col;
    state.dashboard.sortDir = 'desc';
  }
  document.querySelectorAll('.data-table th').forEach(th => {
    th.classList.remove('sorted-asc', 'sorted-desc');
    if (th.dataset.sort === col) {
      th.classList.add(state.dashboard.sortDir === 'asc' ? 'sorted-asc' : 'sorted-desc');
    }
  });
  renderScoreTable();
}

// ===================== DETAIL PANEL =====================
function openDetailPanel(bu) {
  const s = state.dashboard.scores.find(x => x.bu === bu);
  if (!s) return;

  // Header: BU name + meta
  document.getElementById('detail-bu-name').textContent = s.bu;
  const funcInfo = state.dashboard.funcFilter !== 'ALL'
    ? ` · Function: ${state.dashboard.funcFilter}`
    : '';
  document.getElementById('detail-meta').textContent = `Vùng: ${s.region} · ${state.dashboard.month} · Tuần ${state.dashboard.week}${funcInfo}`;

  // Chi tiết Action Plan
  renderAPDetailInPanel(bu);

  // AI phân tích (chỉ hiện khi có dữ liệu AP)
  const aiSection = document.getElementById('detail-ai-section');
  const aiContent = document.getElementById('detail-ai-content');
  if (aiSection) {
    const buData = state.dashboard.allData && state.dashboard.allData[bu];
    const hasAP = buData && buData.AP && buData.AP.length > 0;
    if (hasAP) {
      aiSection.style.display = 'block';
      if (aiContent) aiContent.innerHTML = '';
      const aiBtn = document.getElementById('btn-ai-analyze');
      if (aiBtn) {
        aiBtn.disabled = false;
        aiBtn.innerHTML = '🤖 Phân tích Action Plan bằng AI';
      }
    } else {
      aiSection.style.display = 'none';
    }
  }

  // Mở panel
  document.getElementById('detail-panel').classList.add('open');
  document.getElementById('panel-overlay').classList.add('show');

  // Tải FM Note cho BU này
  loadFMNote(bu);
}

// ===================== AP DETAIL IN FM DASHBOARD =====================
/** Toggle hiển thị chi tiết Action Plan */
function toggleAPDetail() {
  const content = document.getElementById('detail-ap-detail-content');
  const arrow = document.getElementById('ap-detail-arrow');
  if (!content) return;
  const isHidden = content.style.display === 'none';
  content.style.display = isHidden ? 'block' : 'none';
  if (arrow) arrow.textContent = isHidden ? '▼' : '▶';
}

/** Render chi tiết Action Plan trong detail panel của FM Dashboard */
function renderAPDetailInPanel(bu) {
  const container = document.getElementById('detail-ap-detail-content');
  const section = document.getElementById('detail-ap-detail-section');
  if (!container || !section) return;

  // Reset: ẩn nội dung mỗi khi mở panel BU mới
  container.style.display = 'none';
  const arrow = document.getElementById('ap-detail-arrow');
  if (arrow) arrow.textContent = '▶';

  const buData = state.dashboard.allData[bu] || { AP: [], WR: [] };
  let apRows = buData.AP || [];

  // Filter theo function nếu đang lọc
  const funcFilter = state.dashboard.funcFilter;
  if (funcFilter !== 'ALL') {
    apRows = apRows.filter(r => r.func === funcFilter);
  }

  if (apRows.length === 0) {
    container.innerHTML = '<div style="color:#aaa;font-size:13px;padding:12px 0;font-style:italic">Chưa có Action Plan' + (funcFilter !== 'ALL' ? ` cho function ${funcFilter}` : '') + '.</div>';
    // Cập nhật label nút
    const btn = document.getElementById('btn-toggle-ap-detail');
    if (btn) btn.innerHTML = '📋 Xem chi tiết Action Plan (0) <span id="ap-detail-arrow">▶</span>';
    return;
  }

  const funcColors = {
    'GROWTH': { bg: '#e8f5e9', border: '#4caf50', text: '#2e7d32' },
    'OPTIMIZE': { bg: '#e3f2fd', border: '#2196f3', text: '#1565c0' },
    'OPS': { bg: '#fff3e0', border: '#ff9800', text: '#e65100' }
  };

  const severityMap = {
    'Nghiêm trọng': { icon: '🔴', cls: 'ap-severity-critical' },
    'Cần theo dõi': { icon: '🟡', cls: 'ap-severity-watch' },
    '': { icon: '⬜', cls: '' }
  };

  const statusMap = {
    'Hoàn thành': { icon: '✅', cls: 'ap-status-done' },
    'Đang thực hiện': { icon: '🔄', cls: 'ap-status-progress' },
    'Chưa bắt đầu': { icon: '⬜', cls: 'ap-status-pending' },
    '': { icon: '', cls: '' }
  };

  let html = `<div class="ap-detail-list">`;

  apRows.forEach((row, i) => {
    const func = row.func || '';
    const fc = funcColors[func] || { bg: '#f5f5f5', border: '#bbb', text: '#555' };
    const sev = severityMap[row.muc_do] || severityMap[''];
    const st = statusMap[row.status] || statusMap[''];

    const hasContent = row.chi_so || row.van_de || row.key_action;
    if (!hasContent) return;

    html += `
      <div class="ap-detail-card" style="border-left:4px solid ${fc.border}">
        <div class="ap-detail-card-header">
          <span class="ap-detail-func" style="background:${fc.bg};color:${fc.text}">${escHtml(func)}</span>
          ${row.muc_do ? `<span class="ap-detail-severity ${sev.cls}">${sev.icon} ${escHtml(row.muc_do)}</span>` : ''}
          ${row.status ? `<span class="ap-detail-status ${st.cls}">${st.icon} ${escHtml(row.status)}</span>` : ''}
        </div>

        <div class="ap-detail-field">
          <span class="ap-detail-label">Chỉ số:</span>
          <span class="ap-detail-value">${escHtml(row.chi_so || '—')}</span>
        </div>
        <div class="ap-detail-field">
          <span class="ap-detail-label">Vấn đề:</span>
          <span class="ap-detail-value">${escHtml(row.van_de || '—')}</span>
        </div>
        <div class="ap-detail-field">
          <span class="ap-detail-label">Root Cause:</span>
          <span class="ap-detail-value">${escHtml(row.root_cause || '—')}</span>
        </div>
        <div class="ap-detail-field ap-detail-highlight">
          <span class="ap-detail-label">Key Action:</span>
          <span class="ap-detail-value">${escHtml(row.key_action || '—')}</span>
        </div>
        ${row.mo_ta_trien_khai ? `<div class="ap-detail-field">
          <span class="ap-detail-label">Triển khai:</span>
          <span class="ap-detail-value">${escHtml(row.mo_ta_trien_khai)}</span>
        </div>` : ''}

        <div class="ap-detail-meta-row">
          ${row.target_do_luong ? `<span class="ap-detail-meta">🎯 ${escHtml(row.target_do_luong)}</span>` : ''}
          ${row.deadline ? `<span class="ap-detail-meta">📅 ${escHtml(row.deadline)}</span>` : ''}
          ${row.owner ? `<span class="ap-detail-meta">👤 ${escHtml(row.owner)}</span>` : ''}
        </div>
      </div>
    `;
  });

  html += `</div>`;

  // Tổng kết
  const totalActions = apRows.filter(r => r.chi_so || r.van_de || r.key_action).length;
  const completed = apRows.filter(r => r.status === 'Hoàn thành').length;
  const funcs = [...new Set(apRows.map(r => r.func).filter(Boolean))];
  html = `<div class="ap-detail-summary">
    <span>${totalActions} actions</span>
    <span>·</span>
    <span>${funcs.join(', ') || 'Chưa phân loại'}</span>
    ${completed > 0 ? `<span>·</span><span>✅ ${completed} hoàn thành</span>` : ''}
  </div>` + html;

  container.innerHTML = html;
  // Cập nhật label nút toggle với số lượng
  const btn = document.getElementById('btn-toggle-ap-detail');
  const isOpen = container.style.display !== 'none';
  if (btn) btn.innerHTML = `📋 Xem chi tiết Action Plan (${totalActions}) <span id="ap-detail-arrow">${isOpen ? '▼' : '▶'}</span>`;
}

function closeDetailPanel() {
  document.getElementById('detail-panel').classList.remove('open');
  document.getElementById('panel-overlay').classList.remove('show');
  state.dashboard.selectedBU = null;
  document.querySelectorAll('#score-tbody tr').forEach(tr => tr.classList.remove('selected'));
}

// ===================== AI ANALYSIS =====================
async function analyzeWithAI(bu) {
  const s = state.dashboard.scores.find(x => x.bu === bu);
  if (!s) return;

  const buData = state.dashboard.allData[bu] || { AP: [], WR: [] };
  const allData = state.dashboard.allData || {};

  const aiSection = document.getElementById('detail-ai-section');
  const aiContent = document.getElementById('detail-ai-content');
  const aiBtn = document.getElementById('btn-ai-analyze');

  // Show loading
  aiBtn.disabled = true;
  aiBtn.innerHTML = '<span class="ai-spinner"></span> Đang phân tích...';
  aiContent.innerHTML = '<div class="ai-loading">Đang phân tích Action Plan bằng AI...<br><small style="color:#888">Truy xuất dữ liệu lịch sử + benchmark + best-practice BU Xanh...</small></div>';

  // Thu thập AP mẫu từ BU Xanh trong allData (không cần backend fetch lại)
  const bestAps = collectBestAPs(allData, bu, state.dashboard.week);

  try {
    const response = await fetch(`${CGI_BIN}/ai-analyze.py`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        bu: bu,
        week: state.dashboard.week,
        month: state.dashboard.month,
        ap_rows: buData.AP || [],
        wr_rows: buData.WR || [],
        health_status: s.healthLabel || 'Chưa xác định',
        best_aps: bestAps
        // prev_wr_rows: backend sẽ tự fetch từ Google Sheets
      })
    });

    const result = await response.json();

    if (result.success && result.analysis) {
      renderAIAnalysis(result.analysis, buData.AP || []);
    } else {
      aiContent.innerHTML = '<div class="ai-error">Không thể phân tích. Vui lòng thử lại.</div>';
    }
  } catch(e) {
    aiContent.innerHTML = '<div class="ai-error">Lỗi kết nối: ' + escHtml(e.message) + '</div>';
  } finally {
    aiBtn.disabled = false;
    aiBtn.innerHTML = '🤖 Phân tích Action Plan bằng AI';
  }
}

/** Thu thập AP hay từ BU Xanh trong dữ liệu đã tải */
function collectBestAPs(allData, excludeBU, week) {
  const pace = WEEK_PACE[week] || 0.15;
  const bestAps = [];

  for (const [buName, buData] of Object.entries(allData)) {
    if (buName === excludeBU) continue;
    const wr = buData.WR || [];
    const ap = buData.AP || [];
    if (!wr.length || ap.length < 2) continue;

    // Kiểm tra BU Xanh (tất cả doanh số đạt pace)
    const revenueRows = wr.filter(r => {
      const k = (r.kpi || '').trim();
      return k.includes('Doanh số') || k.toUpperCase().includes('TỔNG');
    });
    if (!revenueRows.length) continue;

    const allMeetPace = revenueRows.every(r => {
      const t = parseFloat(r.target);
      const a = parseFloat(r.actual);
      return !isNaN(t) && t > 0 && !isNaN(a) && (a / t) >= pace;
    });

    if (!allMeetPace) continue;

    // Lọc AP có chất lượng
    const qualityAPs = ap.filter(r =>
      r.key_action && String(r.key_action).trim().length > 15 &&
      r.root_cause && String(r.root_cause).trim().length > 15
    ).slice(0, 3);

    if (qualityAPs.length > 0) {
      bestAps.push({ bu: buName, aps: qualityAPs });
    }

    if (bestAps.length >= 3) break;
  }

  return bestAps;
}

function renderAIAnalysis(analysis, apRows) {
  const container = document.getElementById('detail-ai-content');
  let html = '';

  // Trend banner (xu hướng so với tuần trước)
  if (analysis.trend && analysis.trend !== 'NO_DATA') {
    const trendMap = {
      'IMPROVING': { icon: '📈', label: 'Cải thiện', cls: 'ai-trend-up' },
      'DECLINING': { icon: '📉', label: 'Xấu đi', cls: 'ai-trend-down' },
      'STABLE':    { icon: '➡️', label: 'Ổn định', cls: 'ai-trend-stable' }
    };
    const t = trendMap[analysis.trend] || trendMap['STABLE'];
    html += `
      <div class="ai-trend-banner ${t.cls}">
        <span class="ai-trend-icon">${t.icon}</span>
        <span class="ai-trend-label">Xu hướng: ${t.label}</span>
        ${analysis.trend_note ? `<span class="ai-trend-note">${escHtml(analysis.trend_note)}</span>` : ''}
      </div>
    `;
  }

  if (analysis.actions && analysis.actions.length > 0) {
    analysis.actions.forEach((item) => {
      const apRow = apRows[item.index - 1] || {};
      const verdictClass = item.verdict === 'KHA_THI' ? 'ai-verdict-ok' :
                           item.verdict === 'CAN_DIEU_CHINH' ? 'ai-verdict-warn' : 'ai-verdict-bad';
      const verdictIcon = item.verdict === 'KHA_THI' ? '✅' :
                          item.verdict === 'CAN_DIEU_CHINH' ? '⚠️' : '❌';
      const verdictLabel = item.verdict === 'KHA_THI' ? 'Khả thi' :
                           item.verdict === 'CAN_DIEU_CHINH' ? 'Cần điều chỉnh' : 'Không khả thi';

      html += `
        <div class="ai-action-card">
          <div class="ai-action-header">
            <span class="ai-action-num">Action ${item.index}</span>
            <span class="ai-action-func">${escHtml(apRow.func || '')}</span>
            <span class="ai-verdict ${verdictClass}">${verdictIcon} ${verdictLabel}</span>
          </div>
          <div class="ai-action-original">
            <strong>Chỉ số:</strong> ${escHtml(apRow.chi_so || '—')}<br>
            <strong>Key Action:</strong> ${escHtml(apRow.key_action || '—')}
          </div>
          <div class="ai-action-reason">${escHtml(item.reason)}</div>
          ${item.suggestion ? `<div class="ai-action-suggestion"><strong>💡 Gợi ý:</strong> ${escHtml(item.suggestion)}</div>` : ''}
        </div>
      `;
    });
  }

  if (analysis.summary) {
    html += `
      <div class="ai-summary">
        <div class="ai-summary-title">📋 Khuyến nghị cho FM/SOD</div>
        <div class="ai-summary-content">${escHtml(analysis.summary).replace(/\n/g, '<br>')}</div>
      </div>
    `;
  }

  // Data source note
  html += `
    <div style="margin-top:12px;padding:8px 12px;background:#f8f8f8;border:1px solid #e0e0e0;border-radius:4px;font-size:10px;color:#999">
      Model: sonar-pro · Dữ liệu: WR tuần trước + Benchmark OKRs 2026 + Best-practice BU Xanh · Lưu ý: kết quả mang tính tham khảo
    </div>
  `;

  container.innerHTML = html;
}

// ===================== AI DIAGNOSTIC COACH =====================

/** Show/hide diagnostic section based on BU selection + WR data */
function updateDiagnosticVisibility() {
  const section = document.getElementById('diagnostic-section');
  if (!section) return;

  const { bu } = state.cm;
  // Show only when BU is selected AND we're in CM view
  if (bu && state.currentView === 'cm') {
    section.style.display = 'block';
  } else {
    section.style.display = 'none';
  }

  // Reset result when BU changes
  const resultDiv = document.getElementById('diagnostic-result');
  if (resultDiv) {
    resultDiv.style.display = 'none';
    resultDiv.innerHTML = '';
  }

  // Update button state
  updateDiagnosticButton();
}

/** Update diagnostic button enabled/disabled based on current WR data */
function updateDiagnosticButton() {
  const btn = document.getElementById('btn-diagnostic');
  if (!btn) return;
  const { bu, wrRows } = state.cm;
  const hasWRData = wrRows && wrRows.some(r => r.target || r.actual);
  btn.disabled = !hasWRData;
  if (!hasWRData && bu) {
    btn.title = 'Cần điền Weekly Report trước khi chẩn đoán';
  } else {
    btn.title = '';
  }
}

/** Run AI Diagnostic Coach */
async function runDiagnostic() {
  const { bu, apMonth, apWeek, wrRows } = state.cm;
  if (!bu) {
    showToast('Vui lòng chọn BU trước', 'error');
    return;
  }

  const hasWRData = wrRows && wrRows.some(r => r.target || r.actual);
  if (!hasWRData) {
    showToast('Cần có dữ liệu Weekly Report để chẩn đoán', 'error');
    return;
  }

  const btn = document.getElementById('btn-diagnostic');
  const resultDiv = document.getElementById('diagnostic-result');

  // Show loading
  btn.disabled = true;
  btn.innerHTML = '<span class="ai-spinner"></span> Đang chẩn đoán...';
  resultDiv.style.display = 'block';
  resultDiv.innerHTML = `
    <div class="diag-loading">
      <div class="diag-loading-title">Đang chẩn đoán BU ${escHtml(bu)}</div>
      <div class="diag-loading-subtitle">Phân tích 3 giai đoạn: Nhận diện vấn đề → Nguyên nhân gốc rễ → Key Actions...</div>
    </div>
  `;

  try {
    const response = await fetch(`${CGI_BIN}/ai-diagnostic.py`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        bu: bu,
        week: apWeek,
        month: apMonth,
        wr_rows: wrRows.map(r => ({
          kpi: r.kpi || '',
          target: r.target || '',
          actual: r.actual || '',
          notes: r.notes || ''
        }))
        // prev_wr_rows and best_aps will be fetched by backend if needed
      })
    });

    const result = await response.json();

    if (result.success && result.diagnostic) {
      renderDiagnosticResult(result.diagnostic);
      showToast('Chẩn đoán hoàn tất');
    } else {
      resultDiv.innerHTML = '<div class="ai-error" style="padding:16px;text-align:center;color:#cc0000">' +
        escHtml(result.error || 'Không thể chẩn đoán. Vui lòng thử lại.') + '</div>';
    }
  } catch (e) {
    resultDiv.innerHTML = '<div class="ai-error" style="padding:16px;text-align:center;color:#cc0000">Lỗi kết nối: ' +
      escHtml(e.message) + '</div>';
  } finally {
    btn.disabled = false;
    btn.innerHTML = `
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
      Chẩn đoán BU
    `;
  }
}

/** Render 3-phase diagnostic result */
function renderDiagnosticResult(diag) {
  const container = document.getElementById('diagnostic-result');
  container.style.display = 'block';
  let html = '';

  // BU Summary
  if (diag.bu_summary) {
    html += `<div class="diag-summary">${escHtml(diag.bu_summary)}</div>`;
  }

  // Problems
  if (diag.problems && diag.problems.length > 0) {
    diag.problems.forEach((p, pIdx) => {
      // Severity badge
      const isCritical = p.gap_severity === 'NGHIÊM_TRỌNG';
      const sevCls = isCritical ? 'diag-severity-critical' : 'diag-severity-warning';
      const sevLabel = isCritical ? 'Nghiêm trọng' : 'Cần theo dõi';

      // Trend badge
      let trendHtml = '';
      const trendMap = {
        'CẢI_THIỆN': { icon: '↗', cls: 'diag-trend-up', label: 'Cải thiện' },
        'XẤU_ĐI':   { icon: '↘', cls: 'diag-trend-down', label: 'Xấu đi' },
        'ỔN_ĐỊNH':  { icon: '→', cls: 'diag-trend-stable', label: 'Ổn định' },
        'KHÔNG_CÓ_DỮ_LIỆU': { icon: '—', cls: 'diag-trend-nodata', label: '' }
      };
      const tr = trendMap[p.trend] || trendMap['KHÔNG_CÓ_DỮ_LIỆU'];
      if (p.trend && p.trend !== 'KHÔNG_CÓ_DỮ_LIỆU') {
        trendHtml = `<span class="diag-trend ${tr.cls}">${tr.icon} ${tr.label}</span>`;
      }

      // Func badge
      const funcCls = `diag-func-${p.func || 'OPTIMIZE'}`;

      html += `
        <div class="diag-problem-card">
          <div class="diag-problem-header">
            <div class="diag-problem-num">${p.id || pIdx + 1}</div>
            <div class="diag-problem-title">
              <span class="diag-problem-title-indicator">${escHtml(p.indicator || '')}</span>
              <span class="diag-problem-title-func ${funcCls}">${escHtml(p.func || '')}</span>
            </div>
            <span class="diag-severity ${sevCls}">${sevLabel}</span>
            ${trendHtml}
          </div>

          <div class="diag-problem-body">
            <!-- Problem Statement -->
            <div class="diag-problem-statement">${escHtml(p.problem_statement || '')}</div>

            <!-- Phase 2: Root Causes -->
            <div class="diag-phase-label">
              <span class="diag-phase-icon">🔬</span> Nguyên nhân gốc rễ
            </div>
            <div class="diag-rc-list">
      `;

      if (p.root_causes && p.root_causes.length > 0) {
        p.root_causes.forEach(rc => {
          const lkCls = rc.likelihood === 'CAO' ? 'diag-rc-likelihood-high' : 'diag-rc-likelihood-mid';
          const lkLabel = rc.likelihood === 'CAO' ? 'Cao' : 'TB';
          html += `
            <div class="diag-rc-item">
              <div class="diag-rc-header">
                <span class="diag-rc-likelihood ${lkCls}">${lkLabel}</span>
                <span class="diag-rc-cause">${escHtml(rc.cause || '')}</span>
              </div>
              ${rc.evidence ? `<div class="diag-rc-evidence">📊 ${escHtml(rc.evidence)}</div>` : ''}
            </div>
          `;
        });
      }

      html += `
            </div>

            <!-- Phase 3: Suggested Actions -->
            <div class="diag-phase-label">
              <span class="diag-phase-icon">🎯</span> Key Actions gợi ý
            </div>
            <div class="diag-action-list">
      `;

      if (p.suggested_actions && p.suggested_actions.length > 0) {
        p.suggested_actions.forEach((act, aIdx) => {
          const priCls = act.priority === 'P1' ? '' : 'diag-action-priority-p2';
          html += `
            <div class="diag-action-item">
              <div class="diag-action-header">
                <span class="diag-action-priority ${priCls}">${escHtml(act.priority || 'P1')}</span>
                <span class="diag-action-source">${escHtml(act.source || '')}</span>
              </div>
              <div class="diag-action-text">${escHtml(act.action || '')}</div>
              ${act.rationale ? `<div class="diag-action-rationale">${escHtml(act.rationale)}</div>` : ''}
              ${act.target_metric ? `<div class="diag-action-target">${escHtml(act.target_metric)}</div>` : ''}
            </div>
          `;
        });
      }

      // Apply-to-AP button for this problem
      html += `
            </div>
            <button class="diag-apply-btn" onclick="applyDiagnosticToAP(${pIdx})">
              ✚ Áp dụng vào Action Plan
            </button>
          </div>
        </div>
      `;
    });
  }

  // Overall priority
  if (diag.overall_priority) {
    html += `
      <div class="diag-overall">
        <div class="diag-overall-title">⚡ Ưu tiên hành động</div>
        <div class="diag-overall-text">${escHtml(diag.overall_priority)}</div>
      </div>
    `;
  }

  // Source note
  html += `
    <div class="diag-source-note">
      Model: sonar-pro · Phân tích: 3 giai đoạn (Problem ID → Root Cause → Key Actions) · Benchmark: CR16=15%, CR46=50%, AOV=18M · Lưu ý: kết quả mang tính tham khảo, CM cần kiểm chứng trước khi áp dụng
    </div>
  `;

  container.innerHTML = html;

  // Store diagnostic data for apply-to-AP
  state.cm._lastDiagnostic = diag;
}

/** Apply a diagnostic problem to Action Plan */
function applyDiagnosticToAP(problemIdx) {
  const diag = state.cm._lastDiagnostic;
  if (!diag || !diag.problems || !diag.problems[problemIdx]) {
    showToast('Không tìm thấy dữ liệu chẩn đoán', 'error');
    return;
  }

  if (state.cm.isLocked) {
    showToast('Đã khóa chỉnh sửa. Không thể thêm Action Plan.', 'error');
    return;
  }

  const p = diag.problems[problemIdx];

  // Pick the first suggested action (P1 preferred) and first high-likelihood root cause
  const bestRC = (p.root_causes || []).find(rc => rc.likelihood === 'CAO') || (p.root_causes || [])[0];
  const bestAction = (p.suggested_actions || []).find(a => a.priority === 'P1') || (p.suggested_actions || [])[0];

  // Determine mức độ from severity
  const mucDo = p.gap_severity === 'NGHIÊM_TRỌNG' ? 'Nghiêm trọng' : 'Cần theo dõi';

  // Build new AP row
  const newRow = {
    func: p.func || 'OPTIMIZE',
    chi_so: p.indicator || '',
    van_de: p.problem_statement || '',
    muc_do: mucDo,
    root_cause: bestRC ? bestRC.cause : '',
    key_action: bestAction ? bestAction.action : '',
    mo_ta_trien_khai: bestAction && bestAction.rationale ? bestAction.rationale : '',
    target_do_luong: bestAction ? bestAction.target_metric || '' : '',
    deadline: '',
    owner: '',
    fm_support: '',
    status: 'Chưa bắt đầu'
  };

  // Add to AP rows
  state.cm.apRows.push(newRow);
  renderAPTable();

  // Switch to AP tab
  const apTabBtn = document.querySelector('.tab-btn[data-tab="AP"]');
  if (apTabBtn) apTabBtn.click();

  // Scroll to bottom of AP table
  setTimeout(() => {
    const wrapper = document.querySelector('#tab-AP .form-table-wrapper');
    if (wrapper) wrapper.scrollTop = wrapper.scrollHeight;
  }, 100);

  showToast(`Đã thêm vấn đề #${problemIdx + 1} vào Action Plan`);
}

// ===================== CSV EXPORT =====================
function exportCSV() {
  const scores = state.dashboard.filteredScores;
  if (!scores.length) { showToast('Không có dữ liệu để xuất', 'info'); return; }

  const funcFilter = state.dashboard.funcFilter;
  const funcSuffix = funcFilter !== 'ALL' ? `_${funcFilter}` : '';
  const headers = ['BU','Vùng','AP (đã nộp)','WR (đã nộp)','Điểm AP','Điểm WR','Tổng điểm','Xếp loại'];
  const rows = scores.map(s => [
    s.bu, s.region,
    s.hasAP ? 'Có' : 'Không',
    s.hasWR ? 'Có' : 'Không',
    s.apScore, s.wrScore, s.total, s.rating
  ]);

  const csvContent = [headers, ...rows]
    .map(row => row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(','))
    .join('\n');

  const BOM = '\uFEFF'; // UTF-8 BOM cho Excel
  const blob = new Blob([BOM + csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `MindX_CM_Report_${state.dashboard.month}_W${state.dashboard.week}${funcSuffix}.csv`;
  a.click();
  URL.revokeObjectURL(url);
  showToast('Đã xuất CSV thành công');
}

// ===================== INIT =====================
function init() {
  populateSelects();
  initTabs();

  // Khởi tạo CM với dữ liệu mặc định TRƯỚC khi routing
  state.cm.apRows = AP_DEFAULT_ROWS.map(r => ({ ...r }));
  state.cm.wrRows = WR_DEFAULT_ROWS.map(r => ({ ...r }));

  // Render tất cả bảng
  renderAPTable();
  renderWRTable();

  // Routing
  initRouting();

  // CM topbar events (only BU dropdown remains — month/week is auto)
  document.getElementById('cm-bu').addEventListener('change', onCMSelectionChange);

  // Dashboard filter events
  document.getElementById('filter-search').addEventListener('input', applyDashFilters);
  document.getElementById('filter-rating').addEventListener('change', applyDashFilters);
  document.getElementById('filter-region').addEventListener('change', applyDashFilters);

  // Panel overlay click — closes whichever panel is open
  document.getElementById('panel-overlay').addEventListener('click', () => {
    closeDetailPanel();
  });

  // Daily Report events
  const dailyBuEl = document.getElementById('daily-bu');
  const dailyDateEl = document.getElementById('daily-date');
  if (dailyBuEl) dailyBuEl.addEventListener('change', onDailySelectionChange);
  if (dailyDateEl) dailyDateEl.addEventListener('change', onDailySelectionChange);

  // Hiện setup banner nếu chưa cấu hình
  if (!isConfigured()) {
    showSetupBanner();
  }

  // Tải dữ liệu CM nếu đang ở view CM
  if (state.currentView === 'cm') loadCMData();
}

document.addEventListener('DOMContentLoaded', init);

// ============================================================
// DISCUSSION (Thảo luận & Hỏi đáp)
// Không cần mật khẩu — dành cho CM, không theo kỳ báo cáo
// ============================================================

const CGI_MIMI = './cgi-bin/mimi-bot.py';
const CGI_TELEGRAM = './cgi-bin/telegram-bridge.py';

// ---- API wrappers cho Q&A (dùng APPS_SCRIPT_URL giống các module khác) ----

async function qaFetch(action, extraParams = {}) {
  const url = new URL(APPS_SCRIPT_URL);
  url.searchParams.set('action', action);
  Object.entries(extraParams).forEach(([k, v]) => url.searchParams.set(k, v));
  const res = await fetch(url.toString());
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  return res.json();
}

async function qaPost(body) {
  const url = new URL(APPS_SCRIPT_URL);
  // Hỗ trợ action_override cho MiMi/SOD
  const action = body.action_override || 'qa_post';
  url.searchParams.set('action', action);
  // Xóa action_override khỏi body trước khi gửi
  const sendBody = { ...body };
  delete sendBody.action_override;
  const res = await fetch(url.toString(), {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain' },
    body: JSON.stringify(sendBody)
  });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  return res.json();
}

// ---- Animal Avatar System ----
// Mỗi BU được gán 1 biệt danh động vật cố định để phân biệt trong thảo luận mà vẫn ẩn danh
const ANIMAL_ALIASES = [
  { name: 'Hổ', emoji: '🐯' },
  { name: 'Báo', emoji: '🐆' },
  { name: 'Cáo', emoji: '🦊' },
  { name: 'Sóc', emoji: '🐿️' },
  { name: 'Gấu', emoji: '🐻' },
  { name: 'Thỏ', emoji: '🐰' },
  { name: 'Đại Bàng', emoji: '🦅' },
  { name: 'Cú', emoji: '🦉' },
  { name: 'Cá Heo', emoji: '🐬' },
  { name: 'Ngựa', emoji: '🐴' },
  { name: 'Mèo', emoji: '🐱' },
  { name: 'Cánh Cụt', emoji: '🐧' },
  { name: 'Sư Tử', emoji: '🦁' },
  { name: 'Bướm', emoji: '🦋' },
  { name: 'Ong', emoji: '🐝' },
  { name: 'Gà', emoji: '🐔' },
  { name: 'Rồng', emoji: '🐉' },
  { name: 'Voi', emoji: '🐘' },
  { name: 'Hươu', emoji: '🦌' },
  { name: 'Chồn', emoji: '🦝' },
  { name: 'Kỳ Lân', emoji: '🦄' },
  { name: 'Vẹt', emoji: '🦆' },
  { name: 'Cá Ngựa', emoji: '🦛' },
  { name: 'Tuần Lộc', emoji: '🦌' },
  { name: 'Gấu Trúc', emoji: '🐼' },
  { name: 'Sói', emoji: '🐺' },
  { name: 'Cá Voi', emoji: '🐳' },
  { name: 'Khỉ', emoji: '🦜' },
  { name: 'Hải Cẩu', emoji: '🦭' },
  { name: 'Bò Cạp', emoji: '🦂' },
  { name: 'Chuột', emoji: '🐭' },
  { name: 'Đà Điểu', emoji: '🦤' },
  { name: 'Rùa', emoji: '🐢' },
  { name: 'Bạch Tuộc', emoji: '🐙' },
  { name: 'Chim Sáo', emoji: '🐦' },
  { name: 'Khỉ Vàng', emoji: '🐒' },
  { name: 'Côn Trùng', emoji: '🦗' },
  { name: 'Hải Mã', emoji: '🧭' },
  { name: 'Ngỗng', emoji: '🦢' },
  { name: 'Cá Sấu', emoji: '🐊' },
  { name: 'Tê Giác', emoji: '🦏' },
  { name: 'Hả Mã', emoji: '🦛' },
  { name: 'Chim Đại Bàng', emoji: '🦅' },
  { name: 'Cá Vàng', emoji: '🐠' },
  { name: 'Thạch Sùng', emoji: '🦎' },
  { name: 'Cá Mập', emoji: '🦈' },
  { name: 'Hổ Phượng', emoji: '🦩' },
  { name: 'Nhím', emoji: '🦔' },
  { name: 'Cò Sếu', emoji: '🦩' },
  { name: 'Béo', emoji: '🐾' },
];

function hashString(str) {
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    const char = str.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash |= 0;
  }
  return Math.abs(hash);
}

function getAnimalAlias(bu, postId) {
  // Khi có BU: hash từ BU (cùng BU = cùng avatar)
  // Khi ẩn danh (bu trống): hash từ postId (mỗi post = avatar riêng)
  const key = bu || postId || '';
  if (!key) return { name: 'Khách', emoji: '👤', full: '👤 Khách' };
  const idx = hashString(key) % ANIMAL_ALIASES.length;
  const animal = ANIMAL_ALIASES[idx];
  return { name: animal.name, emoji: animal.emoji, full: `${animal.emoji} ${animal.name}` };
}

// ---- Search & Sort ----

function setDiscussionSort(sortBy) {
  state.discussion.sortBy = sortBy;
  // Cập nhật UI buttons
  document.querySelectorAll('.disc-sort-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.sort === sortBy);
  });
  loadDiscussion();
}

let _searchDebounce = null;
function onDiscussionSearch(value) {
  clearTimeout(_searchDebounce);
  _searchDebounce = setTimeout(() => {
    state.discussion.searchTerm = value.trim();
    loadDiscussion();
  }, 400);
}

// ---- Init ----

function initDiscussionView() {
  // Populate BU select (no-op vì đã xóa dropdown)
  populateDiscussionBUSelect();

  // Attach textarea char counter
  const ta = document.getElementById('discussion-question-input');
  if (ta) {
    ta.addEventListener('input', updateDiscussionCharCount);
  }

  // Hiển thị biệt danh của CM trong note
  const anonNote = document.getElementById('discussion-anon-note');
  if (anonNote) {
    const bu = (state.cm && state.cm.bu) ? state.cm.bu : '';
    const alias = getAnimalAlias(bu);
    anonNote.innerHTML = `🔒 Bạn sẽ xuất hiện với biệt danh <strong>${alias.full}</strong> — không ai biết bạn là ai`;
  }

  // Sync sort buttons
  document.querySelectorAll('.disc-sort-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.sort === state.discussion.sortBy);
  });

  // Nếu chưa có dữ liệu thì load
  if (state.discussion.questions.length === 0) {
    loadDiscussion();
  } else {
    renderDiscussionList();
  }

  // Auto-refresh mỗi 30 giây
  if (state.discussion.refreshTimer) {
    clearInterval(state.discussion.refreshTimer);
  }
  state.discussion.refreshTimer = setInterval(() => {
    // Chỉ refresh khi đang ở view discussion
    if (state.currentView === 'discussion') {
      silentRefreshDiscussion();
    }
  }, 30000);
}

function populateDiscussionBUSelect() {
  const sel = document.getElementById('discussion-bu-select');
  if (!sel) return;
  // Giữ option đầu tiên
  while (sel.options.length > 1) sel.remove(1);
  BU_LIST.forEach(bu => {
    const opt = document.createElement('option');
    opt.value = bu;
    opt.textContent = bu;
    sel.appendChild(opt);
  });

  // Tự động chọn BU nếu CM đang chọn ở CM view
  const cmBu = state.cm.bu;
  if (cmBu) {
    sel.value = cmBu;
  }
}

function updateDiscussionCharCount() {
  const ta = document.getElementById('discussion-question-input');
  const counter = document.getElementById('discussion-char-count');
  if (!ta || !counter) return;
  const len = ta.value.length;
  const max = 500;
  counter.textContent = `${len} / ${max}`;
  counter.style.color = len > max ? '#cc0000' : '#999';
}

// ---- Load & Render ----

async function loadDiscussion(reset = true) {
  if (state.discussion.loading) return;
  state.discussion.loading = true;

  if (reset) {
    state.discussion.page = 1;
    state.discussion.questions = [];
  }

  showDiscussionLoading(true);

  try {
    const data = await qaFetch('qa_list', {
      page: state.discussion.page,
      limit: 15,
      sort: state.discussion.sortBy,
      search: state.discussion.searchTerm
    });

    if (reset) {
      state.discussion.questions = data.questions || [];
    } else {
      state.discussion.questions = state.discussion.questions.concat(data.questions || []);
    }

    state.discussion.hasMore = data.has_more || false;

    // Cập nhật label tổng số
    const totalLabel = document.getElementById('discussion-total-label');
    if (totalLabel) {
      totalLabel.textContent = `${data.total || 0} câu hỏi`;
    }

    renderDiscussionList();

  } catch (e) {
    showToast('Lỗi tải danh sách câu hỏi: ' + e.message, 'error');
    showDiscussionEmpty();
  } finally {
    state.discussion.loading = false;
    showDiscussionLoading(false);
  }
}

async function silentRefreshDiscussion() {
  // Refresh không reset danh sách đang expand
  try {
    const data = await qaFetch('qa_list', { page: 1, limit: 15, sort: state.discussion.sortBy, search: state.discussion.searchTerm });
    if (data.questions) {
      state.discussion.questions = data.questions;
      state.discussion.hasMore = data.has_more || false;
      renderDiscussionList();
      // Cập nhật label
      const totalLabel = document.getElementById('discussion-total-label');
      if (totalLabel) totalLabel.textContent = `${data.total || 0} câu hỏi`;
    }
  } catch (e) {
    // Silent fail
  }
}

async function loadMoreQuestions() {
  state.discussion.page++;
  await loadDiscussion(false);
}

function showDiscussionLoading(show) {
  const loading = document.getElementById('discussion-loading');
  const list = document.getElementById('discussion-list');
  if (show) {
    if (loading) loading.style.display = 'flex';
    if (list) list.style.opacity = '0.5';
  } else {
    if (loading) loading.style.display = 'none';
    if (list) list.style.opacity = '1';
  }
}

function showDiscussionEmpty() {
  const list = document.getElementById('discussion-list');
  if (!list) return;
  const loading = document.getElementById('discussion-loading');
  if (loading) loading.style.display = 'none';
  list.innerHTML = `
    <div class="discussion-empty">
      <div class="discussion-empty-icon">❓</div>
      <div class="discussion-empty-title">Chưa có câu hỏi nào</div>
      <div class="discussion-empty-sub">Hãy là người đầu tiên đặt câu hỏi!</div>
    </div>
  `;
}

function renderDiscussionList() {
  const list = document.getElementById('discussion-list');
  if (!list) return;

  const qs = state.discussion.questions;

  if (qs.length === 0) {
    if (state.discussion.searchTerm) {
      list.innerHTML = `
        <div class="discussion-empty">
          <div class="discussion-empty-icon">🔍</div>
          <div class="discussion-empty-title">Không tìm thấy kết quả</div>
          <div class="discussion-empty-sub">Thử tìm với từ khóa khác</div>
        </div>
      `;
    } else {
      showDiscussionEmpty();
    }
    return;
  }

  list.innerHTML = qs.map(q => renderQuestionCard(q)).join('');

  // Load more button
  const loadMoreEl = document.getElementById('discussion-load-more');
  if (loadMoreEl) {
    loadMoreEl.style.display = state.discussion.hasMore ? 'block' : 'none';
  }
}

function renderQuestionCard(q) {
  const isExpanded = state.discussion.expandedIds.has(q.id);
  const hasVoted = state.discussion.votedIds.has(q.id);
  const timeAgo = fmtTimeAgo(q.created_at);
  const answerCount = (q.answers || []).length;
  const alias = getAnimalAlias(q.bu, q.id);

  const answersHtml = isExpanded ? renderAnswersSection(q) : '';

  return `
    <div class="qa-card" id="qa-card-${escHtml(q.id)}">
      <div class="qa-card-header">
        <div class="qa-card-meta">
          <span class="qa-badge qa-badge-anon">${alias.emoji} ${escHtml(alias.name)}</span>
          <span class="qa-time">${escHtml(timeAgo)}</span>
        </div>
        <div class="qa-card-actions">
          <button class="qa-upvote-btn ${hasVoted ? 'voted' : ''}" onclick="upvoteQuestion('${escHtml(q.id)}')" title="Ủng hộ câu hỏi này">
            <svg width="12" height="12" viewBox="0 0 24 24" fill="${hasVoted ? 'currentColor' : 'none'}" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="18 15 12 9 6 15"/></svg>
            <span>${q.upvotes || 0}</span>
          </button>
        </div>
      </div>

      <div class="qa-card-content">${escHtml(q.content)}</div>

      <div class="qa-card-footer">
        <button class="qa-toggle-btn" onclick="toggleAnswers('${escHtml(q.id)}')">
          <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg>
          ${answerCount > 0 ? `${answerCount} trả lời` : 'Trả lời'}
          <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="${isExpanded ? '18 15 12 9 6 15' : '6 9 12 15 18 9'}"/></svg>
        </button>
      </div>

      <div class="qa-answers-section" id="qa-answers-${escHtml(q.id)}" style="display:${isExpanded ? 'block' : 'none'}">
        ${answersHtml}
      </div>
    </div>
  `;
}

function renderAnswersSection(q) {
  const answers = q.answers || [];

  const answersHtml = answers.map(a => renderAnswerItem(a)).join('');

  return `
    <div class="qa-answers-list">
      ${answers.length === 0
        ? '<div class="qa-no-answers">Chưa có câu trả lời nào. Hãy là người đầu tiên!</div>'
        : answersHtml
      }
    </div>
    <div class="qa-reply-area">
      <div class="qa-reply-input-wrap">
        <textarea class="qa-reply-textarea" id="qa-reply-${escHtml(q.id)}" placeholder="Viết câu trả lời của bạn..." rows="2"></textarea>
      </div>
      <div class="qa-reply-actions">
        <button class="qa-action-btn qa-action-reply" onclick="submitAnswer('${escHtml(q.id)}', 'answer')">
          <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/></svg>
          Trả lời ẩn danh
        </button>
      </div>
    </div>
  `;
}

function renderAnswerItem(a) {
  const timeAgo = fmtTimeAgo(a.created_at);
  let badgeHtml = '';

  if (a.type === 'ai_answer') {
    badgeHtml = `<span class="qa-badge qa-badge-ai">🤖 MiMi</span>`;
  } else if (a.type === 'sod_answer') {
    badgeHtml = `<span class="qa-badge qa-badge-sod">⭐ HO</span>`;
  } else {
    const alias = getAnimalAlias(a.bu, a.id);
    badgeHtml = `<span class="qa-badge qa-badge-anon">${alias.emoji} ${escHtml(alias.name)}</span>`;
  }

  // Format nội dung trả lời
  // AI/SOD answers: render markdown đơn giản. CM answers: plain text
  const isRich = (a.type === 'ai_answer' || a.type === 'sod_answer');
  const contentHtml = isRich ? renderSimpleMarkdown(a.content) : escHtml(a.content).replace(/\n/g, '<br>');

  return `
    <div class="qa-answer-item qa-answer-type-${escHtml(a.type)}">
      <div class="qa-answer-meta">
        ${badgeHtml}
        <span class="qa-time">${escHtml(timeAgo)}</span>
      </div>
      <div class="qa-answer-content qa-md">${contentHtml}</div>
    </div>
  `;
}

// ---- Convert markdown table to simple list ----
function convertTableToList(text) {
  if (!text || !text.includes('|')) return text;
  const lines = text.split('\n');
  const result = [];
  let headers = [];
  let inTable = false;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    // Dòng table: bắt đầu và kết thúc bằng | hoặc chứa nhiều |
    const cells = line.split('|').map(c => c.trim()).filter(c => c.length > 0);

    if (cells.length >= 2 && line.includes('|')) {
      // Bỏ qua dòng separator (|---|---|)
      if (/^[\s|:-]+$/.test(line)) {
        continue;
      }
      if (!inTable) {
        // Dòng đầu tiên = headers
        headers = cells;
        inTable = true;
      } else {
        // Dòng dữ liệu: kết hợp header + value
        let itemParts = [];
        for (let j = 0; j < cells.length; j++) {
          if (headers[j] && cells[j]) {
            itemParts.push(headers[j] + ': ' + cells[j]);
          } else if (cells[j]) {
            itemParts.push(cells[j]);
          }
        }
        if (itemParts.length > 0) {
          result.push('- ' + itemParts.join(' | '));
        }
      }
    } else {
      // Hết bảng — reset
      if (inTable) {
        inTable = false;
        headers = [];
      }
      result.push(lines[i]);
    }
  }
  return result.join('\n');
}

// ---- Simple Markdown Renderer cho AI/SOD answers ----
function renderSimpleMarkdown(text) {
  if (!text) return '';
  // Chuyển markdown table thành danh sách đơn giản trước khi render
  text = convertTableToList(text);
  let html = escHtml(text);
  // Headers: ### h3, ## h2, # h1
  html = html.replace(/^### (.+)$/gm, '<h4 class="qa-md-h">$1</h4>');
  html = html.replace(/^## (.+)$/gm, '<h3 class="qa-md-h">$1</h3>');
  html = html.replace(/^# (.+)$/gm, '<h3 class="qa-md-h">$1</h3>');
  // Bold: **text**
  html = html.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
  // Italic: *text* (nhưng không match ** đã xử lý)
  html = html.replace(/(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)/g, '<em>$1</em>');
  // Bullet lists: - item
  html = html.replace(/^- (.+)$/gm, '<li>$1</li>');
  html = html.replace(/(<li>.*<\/li>)/gs, '<ul class="qa-md-list">$1</ul>');
  // Numbered lists: 1. item
  html = html.replace(/^\d+\. (.+)$/gm, '<li>$1</li>');
  // Clean up consecutive <ul> tags
  html = html.replace(/<\/ul>\s*<ul class="qa-md-list">/g, '');
  // Line breaks
  html = html.replace(/\n/g, '<br>');
  // Clean up <br> thừa ngay sau headers, lists, và giữa các dòng trống
  html = html.replace(/<\/h[34]><br>/g, '</h3>');
  html = html.replace(/<\/ul><br>/g, '</ul>');
  html = html.replace(/<ul class="qa-md-list"><br>/g, '<ul class="qa-md-list">');
  html = html.replace(/<br><br><br>/g, '<br>');
  html = html.replace(/<br><br>/g, '<br>');
  return html;
}

// ---- Actions ----

function toggleAnswers(questionId) {
  const section = document.getElementById(`qa-answers-${questionId}`);
  if (!section) return;

  const isExpanded = state.discussion.expandedIds.has(questionId);

  if (isExpanded) {
    state.discussion.expandedIds.delete(questionId);
    section.style.display = 'none';
  } else {
    state.discussion.expandedIds.add(questionId);
    // Tìm question trong state để render content
    const q = state.discussion.questions.find(x => x.id === questionId);
    if (q) {
      section.innerHTML = renderAnswersSection(q);
    }
    section.style.display = 'block';
  }

  // Cập nhật toggle button icon
  const card = document.getElementById(`qa-card-${questionId}`);
  if (card) {
    const toggleBtn = card.querySelector('.qa-toggle-btn');
    if (toggleBtn) {
      const q = state.discussion.questions.find(x => x.id === questionId);
      const newExpanded = state.discussion.expandedIds.has(questionId);
      const answerCount = q ? (q.answers || []).length : 0;
      const arrowPts = newExpanded ? '18 15 12 9 6 15' : '6 9 12 15 18 9';
      toggleBtn.innerHTML = `
        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg>
        ${answerCount > 0 ? `${answerCount} trả lời` : 'Trả lời'}
        <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="${arrowPts}"/></svg>
      `;
    }
  }
}

async function submitQuestion(mode) {
  // mode: 'discuss' | 'mimi' | 'sod'
  if (!mode) mode = 'discuss';

  const ta = document.getElementById('discussion-question-input');
  if (!ta) return;

  const content = ta.value.trim();
  const bu = (state.cm && state.cm.bu) ? state.cm.bu : '';

  if (!content) {
    showToast('Vui lòng nhập câu hỏi', 'error');
    ta.focus();
    return;
  }

  if (content.length > 500) {
    showToast('Câu hỏi tối đa 500 ký tự', 'error');
    return;
  }

  // Disable tất cả 3 buttons
  const allBtns = document.querySelectorAll('.disc-submit-btn');
  allBtns.forEach(b => { b.disabled = true; });
  const activeBtn = document.querySelector(`.disc-submit-${mode}`);
  if (activeBtn) activeBtn.textContent = 'Đang gửi...';

  try {
    const result = await qaPost({
      type: 'question',
      parent_id: '',
      bu: bu,
      content: content
    });

    if (!result.success) {
      showToast('Lỗi: ' + (result.error || 'Không thể đăng câu hỏi'), 'error');
      return;
    }

    ta.value = '';
    updateDiscussionCharCount();

    const questionId = result.id || '';

    if (mode === 'discuss') {
      showToast('✓ Câu hỏi đã được đăng ẩn danh!');
      await loadDiscussion(true);
    } else if (mode === 'mimi') {
      showToast('✓ Câu hỏi đã đăng — MiMi đang trả lời...', 'info');
      await loadDiscussion(true);
      // Gọi MiMi AI trả lời
      if (questionId) {
        await askMiMiFor(questionId, content);
      }
    } else if (mode === 'sod') {
      showToast('✓ Câu hỏi đã đăng — đang gửi đến HO...', 'info');
      await loadDiscussion(true);
      // Gửi qua Telegram cho SOD
      if (questionId) {
        await requestSODAnswer(questionId, content);
      }
    }
  } catch (e) {
    showToast('Lỗi kết nối: ' + e.message, 'error');
  } finally {
    // Re-enable buttons
    allBtns.forEach(b => { b.disabled = false; });
    resetSubmitButtons();
  }
}

function resetSubmitButtons() {
  const d = document.querySelector('.disc-submit-discuss');
  if (d) d.innerHTML = '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg> Thảo luận cùng CM';
  const m = document.querySelector('.disc-submit-mimi');
  if (m) m.innerHTML = '🤖 Hỏi MiMi';
  const s = document.querySelector('.disc-submit-sod');
  if (s) s.innerHTML = '⭐ Hỏi HO';
}

async function submitAnswer(questionId, type) {
  const ta = document.getElementById(`qa-reply-${questionId}`);
  if (!ta) return;

  const content = ta.value.trim();
  if (!content) {
    showToast('Vui lòng nhập nội dung trả lời', 'error');
    ta.focus();
    return;
  }

  // Disable action buttons tạm thời
  const section = document.getElementById(`qa-answers-${questionId}`);
  const actionBtns = section ? section.querySelectorAll('.qa-action-btn') : [];
  actionBtns.forEach(b => { b.disabled = true; });

  try {
    const result = await qaPost({
      type: type,
      parent_id: questionId,
      bu: '',
      content: content
    });

    if (result.success) {
      ta.value = '';
      showToast('✓ Đã gửi câu trả lời');
      // Refresh để lấy câu trả lời mới
      await refreshQuestion(questionId);
    } else {
      showToast('Lỗi: ' + (result.error || 'Không thể gửi'), 'error');
    }
  } catch (e) {
    showToast('Lỗi kết nối: ' + e.message, 'error');
  } finally {
    actionBtns.forEach(b => { b.disabled = false; });
  }
}

async function upvoteQuestion(questionId) {
  if (state.discussion.votedIds.has(questionId)) {
    showToast('Bạn đã ủng hộ câu hỏi này rồi', 'info');
    return;
  }

  try {
    const result = await qaFetch('qa_upvote', { id: questionId });

    if (result.success) {
      // Cập nhật trong state
      const q = state.discussion.questions.find(x => x.id === questionId);
      if (q) q.upvotes = result.upvotes;

      // Đánh dấu đã vote
      state.discussion.votedIds.add(questionId);

      // Cập nhật UI
      const card = document.getElementById(`qa-card-${questionId}`);
      if (card) {
        const upvoteBtn = card.querySelector('.qa-upvote-btn');
        if (upvoteBtn) {
          upvoteBtn.classList.add('voted');
          upvoteBtn.innerHTML = `
            <svg width="12" height="12" viewBox="0 0 24 24" fill="currentColor" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="18 15 12 9 6 15"/></svg>
            <span>${result.upvotes}</span>
          `;
        }
      }
    } else {
      showToast('Không thể ủng hộ: ' + (result.error || ''), 'error');
    }
  } catch (e) {
    showToast('Lỗi kết nối', 'error');
  }
}

async function askMiMiFor(questionId, questionContent) {
  showToast('MiMi đang trả lời...', 'info');

  // Truyền BU + month + week để MiMi trích xuất dữ liệu nội bộ
  const bu = (state.cm && state.cm.bu) ? state.cm.bu : '';
  const month = (state.cm && state.cm.apMonth) ? state.cm.apMonth : (state.dashboard && state.dashboard.month) || '';
  const week = (state.cm && state.cm.apWeek) ? state.cm.apWeek : (state.dashboard && state.dashboard.week) || '';

  try {
    const result = await qaPost({
      action_override: 'mimi_ask',
      question_id: questionId,
      question_content: questionContent,
      bu: bu,
      month: month,
      week: week
    });

    if (result.success) {
      showToast('✓ MiMi đã trả lời!');
      await refreshQuestion(questionId);
    } else {
      showToast('Lỗi MiMi: ' + (result.error || 'Không có phản hồi'), 'error');
    }
  } catch (e) {
    showToast('Lỗi kết nối với MiMi: ' + e.message, 'error');
  }
}

async function requestSODAnswer(questionId, questionContent) {
  try {
    const result = await qaPost({
      action_override: 'sod_notify',
      question_id: questionId,
      question_content: questionContent
    });

    if (result.success) {
      showToast('✓ Đã gửi câu hỏi lên HO qua Telegram!');
    } else {
      showToast('Lỗi gửi Telegram: ' + (result.error || 'Không có phản hồi'), 'error');
    }
  } catch (e) {
    showToast('Lỗi kết nối: ' + e.message, 'error');
  }
}

// Refresh 1 câu hỏi cụ thể sau khi có câu trả lời mới
async function refreshQuestion(questionId) {
  try {
    // Tải lại toàn bộ (page 1) để có dữ liệu mới
    const data = await qaFetch('qa_list', { page: 1, limit: 15 });
    if (data.questions) {
      state.discussion.questions = data.questions;
      state.discussion.hasMore = data.has_more || false;
    }

    // Re-render card cụ thể
    const q = state.discussion.questions.find(x => x.id === questionId);
    if (q) {
      // Đảm bảo vẫn expanded
      state.discussion.expandedIds.add(questionId);

      const card = document.getElementById(`qa-card-${questionId}`);
      if (card) {
        const parent = card.parentNode;
        const newCardHtml = document.createElement('div');
        newCardHtml.innerHTML = renderQuestionCard(q);
        parent.replaceChild(newCardHtml.firstElementChild, card);
      }
    }
  } catch (e) {
    // fallback: reload toàn bộ
    await loadDiscussion(true);
  }
}

// ---- Utilities ----

function fmtTimeAgo(isoString) {
  if (!isoString) return '';
  const now = new Date();
  const then = new Date(isoString);
  const diff = now - then; // ms

  const mins = Math.floor(diff / 60000);
  if (mins < 1) return 'Vừa xong';
  if (mins < 60) return `${mins} phút trước`;

  const hours = Math.floor(mins / 60);
  if (hours < 24) return `${hours} giờ trước`;

  const days = Math.floor(hours / 24);
  if (days < 7) return `${days} ngày trước`;

  // Nhiều hơn 7 ngày: hiển thị ngày
  const d = then.getDate().toString().padStart(2, '0');
  const m = (then.getMonth() + 1).toString().padStart(2, '0');
  const y = then.getFullYear();
  return `${d}/${m}/${y}`;
}

// ===================== FEATURE 1: FM NOTE (Thread v2) =====================

state.cm.fmNotes = []; // array of { note_id, fm_name, note, saved_at, read_by_cm }
state.dashboard._fmNoteCache = {}; // { bu: { total, unread, latest_at } }

// --- FM Note: Dashboard (FM writes) ---

/** Load tất cả FM Notes cho BU đang chọn trong detail panel */
async function loadFMNote(bu) {
  const { month, week } = state.dashboard;
  const existingDiv = document.getElementById('detail-fm-note-existing');
  const textarea = document.getElementById('fm-note-textarea');
  const saveStatus = document.getElementById('fm-note-save-status');
  if (!existingDiv) return;

  existingDiv.innerHTML = '<div style="color:#aaa;font-size:12px">Đang tải ghi chú...</div>';
  if (textarea) textarea.value = '';
  if (saveStatus) saveStatus.textContent = '';

  if (!isConfigured()) {
    existingDiv.innerHTML = '<div style="color:#aaa;font-size:12px;font-style:italic">Chưa kết nối (dev mode)</div>';
    return;
  }

  try {
    const result = await apiFetch('get_fm_note', { bu, month, week });
    renderFMNoteThread(result, existingDiv, 'fm');
  } catch(e) {
    existingDiv.innerHTML = `<div style="color:#cc0000;font-size:12px">Lỗi tải: ${escHtml(e.message)}</div>`;
  }
}

/** Render FM note thread (dùng cho cả FM panel và CM view) */
function renderFMNoteThread(data, container, viewMode) {
  if (!container) return;
  const notes = (data && data.notes) ? data.notes : [];

  if (!notes.length) {
    container.innerHTML = '<div style="color:#aaa;font-size:12px;font-style:italic">Chưa có ghi chú nào.</div>';
    return;
  }

  const total = data.total || notes.length;
  const unread = data.unread || 0;

  let headerHtml = `<div class="note-thread-header">`;
  headerHtml += `<span class="note-thread-count">${total} ghi chú</span>`;
  if (viewMode === 'fm' && unread > 0) {
    headerHtml += `<span class="note-thread-unread">${unread} chờ CM xác nhận</span>`;
  }
  if (viewMode === 'cm' && unread > 0) {
    headerHtml += `<span class="note-thread-unread">${unread} chưa đọc</span>`;
    headerHtml += `<button class="btn-mark-all-read" onclick="markFMNoteRead()">✅ Xác nhận tất cả</button>`;
  }
  headerHtml += `</div>`;

  let threadHtml = notes.map(n => {
    const readCls = n.read_by_cm ? 'read' : 'unread';
    const readBadge = n.read_by_cm
      ? `<span class="note-status-badge read">✓ Đã đọc</span>`
      : `<span class="note-status-badge unread">⏳ Chờ xác nhận</span>`;
    const fmLabel = n.fm_name ? `<span class="note-fm-name">${escHtml(n.fm_name)}</span>` : '';
    const timeStr = n.saved_at ? fmt(n.saved_at) : '';

    let actionBtns = '';
    if (viewMode === 'cm' && !n.read_by_cm) {
      actionBtns = `<button class="btn-note-confirm-single" onclick="markFMNoteReadSingle('${escHtml(n.note_id)}')">✅ Đã đọc</button>`;
    }
    if (viewMode === 'fm') {
      actionBtns = `<button class="btn-note-delete" onclick="deleteFMNote('${escHtml(n.note_id)}')" title="Xóa ghi chú">🗑️</button>`;
    }

    return `
      <div class="note-thread-item ${readCls}">
        <div class="note-thread-item-header">
          <div class="note-thread-item-meta">
            ${fmLabel}
            <span class="note-thread-item-time">${timeStr}</span>
          </div>
          <div class="note-thread-item-actions">
            ${readBadge}
            ${actionBtns}
          </div>
        </div>
        <div class="note-thread-item-text">${escHtml(n.note).replace(/\n/g,'<br>')}</div>
      </div>
    `;
  }).join('');

  container.innerHTML = headerHtml + `<div class="note-thread-list">${threadHtml}</div>`;
}

/** FM saves a note */
async function saveFMNote() {
  const bu = state.dashboard.selectedBU;
  const { month, week } = state.dashboard;
  const textarea = document.getElementById('fm-note-textarea');
  const saveStatus = document.getElementById('fm-note-save-status');

  if (!bu) { showToast('Vui lòng chọn BU trước', 'error'); return; }
  const note = textarea ? textarea.value.trim() : '';
  if (!note) { showToast('Vui lòng nhập nội dung ghi chú', 'error'); return; }

  if (!isConfigured()) {
    showToast('Dev mode: ghi chú đã lưu (giả lập)', 'info');
    return;
  }

  if (saveStatus) saveStatus.textContent = 'Đang lưu...';

  // Lấy FM name từ cache trong bộ nhớ
  let fmName = state.dashboard._fmName || '';
  if (!fmName) {
    fmName = prompt('Nhập tên FM của bạn (lưu trong phiên làm việc):') || '';
    if (fmName) state.dashboard._fmName = fmName;
  }

  try {
    const result = await apiFetch('save_fm_note', {
      bu, month, week, note, fm_name: fmName
    });
    if (result.success) {
      showToast('✓ Đã lưu FM Note');
      if (saveStatus) saveStatus.textContent = `Đã lưu lúc ${fmtNow()}`;
      if (textarea) textarea.value = '';
      await loadFMNote(bu);
      // Cập nhật cache và icon trong bảng
      await loadFMNotesSummary();
      renderScoreTable();
    } else {
      showToast('Lỗi lưu note: ' + (result.error || 'unknown'), 'error');
      if (saveStatus) saveStatus.textContent = '';
    }
  } catch(e) {
    showToast('Lỗi kết nối: ' + e.message, 'error');
    if (saveStatus) saveStatus.textContent = '';
  }
}

/** FM xóa 1 note cụ thể */
async function deleteFMNote(noteId) {
  if (!confirm('Xóa ghi chú này?')) return;
  try {
    const result = await apiFetch('delete_fm_note', { note_id: noteId });
    if (result.success) {
      showToast('✓ Đã xóa ghi chú');
      const bu = state.dashboard.selectedBU;
      if (bu) await loadFMNote(bu);
      await loadFMNotesSummary();
      renderScoreTable();
    } else {
      showToast('Lỗi: ' + (result.error || 'unknown'), 'error');
    }
  } catch(e) {
    showToast('Lỗi kết nối: ' + e.message, 'error');
  }
}

/** Load batch note summary cho tất cả BUs (dùng cho icon trong bảng FM) */
async function loadFMNotesSummary() {
  if (!isConfigured()) return;
  const { month, week } = state.dashboard;
  try {
    const result = await apiFetch('get_fm_notes_summary', { month, week });
    state.dashboard._fmNoteCache = result || {};
  } catch(e) {
    // Silent fail
  }
}

// --- FM Note: CM View (CM reads and acknowledges) ---

/** Tải và hiển thị tất cả FM Notes cho CM */
async function loadCMFMNote() {
  const { bu, apMonth, apWeek } = state.cm;
  if (!bu) return;

  const section = document.getElementById('cm-fm-note-section');
  if (!section) return;

  if (!isConfigured()) {
    section.style.display = 'none';
    return;
  }

  try {
    const result = await apiFetch('get_fm_note', { bu, month: apMonth, week: apWeek });
    state.cm.fmNotes = (result && result.notes) ? result.notes : [];

    if (!result || !result.notes || !result.notes.length) {
      section.style.display = 'none';
      return;
    }

    section.style.display = 'block';
    const container = document.getElementById('cm-fm-note-thread-container');
    if (container) {
      renderFMNoteThread(result, container, 'cm');
    }
  } catch(e) {
    section.style.display = 'none';
  }
}

/** CM đánh dấu tất cả đã đọc */
async function markFMNoteRead() {
  const { bu, apMonth, apWeek } = state.cm;
  if (!bu) return;

  if (!isConfigured()) {
    showToast('Dev mode: đã đánh dấu (giả lập)');
    return;
  }

  try {
    const result = await apiFetch('mark_fm_note_read', { bu, month: apMonth, week: apWeek });
    if (result.success) {
      showToast(`✓ Đã xác nhận ${result.marked || 'tất cả'} ghi chú FM`);
      await loadCMFMNote();
    } else {
      showToast('Lỗi: ' + (result.error || 'unknown'), 'error');
    }
  } catch(e) {
    showToast('Lỗi kết nối: ' + e.message, 'error');
  }
}

/** CM đánh dấu 1 note cụ thể đã đọc */
async function markFMNoteReadSingle(noteId) {
  const { bu, apMonth, apWeek } = state.cm;
  if (!bu) return;

  try {
    const result = await apiFetch('mark_fm_note_read', { bu, month: apMonth, week: apWeek, note_id: noteId });
    if (result.success) {
      showToast('✓ Đã xác nhận ghi chú');
      await loadCMFMNote();
    } else {
      showToast('Lỗi: ' + (result.error || 'unknown'), 'error');
    }
  } catch(e) {
    showToast('Lỗi kết nối: ' + e.message, 'error');
  }
}

// ===================== FEATURE 2: FUNCTION DEEP-DIVE DASHBOARD =====================

const FUNC_COLORS = {
  GROWTH:   { main: '#1a7a3a', bg: '#e8f5ed', border: '#c3e6cb' },
  OPTIMIZE: { main: '#1a5aa0', bg: '#e8eef8', border: '#c3d4e6' },
  OPS:      { main: '#c07000', bg: '#fdf3e0', border: '#f0d9a0' }
};

const FUNC_KPI_CONFIG = {
  GROWTH: {
    label: 'FM Growth (N3)',
    kpis: [
      { key: 'Growth (N3) - L1 Leads tự kiếm', label: 'L1 Leads tự kiếm' },
      { key: 'Growth (N3) - L4 Trials',        label: 'L4 Trials' },
      { key: 'Growth (N3) - L6 Deals',         label: 'L6 Deals' },
      { key: 'Growth (N3) - Doanh số N3',      label: 'Doanh số N3', isRevenue: true }
    ],
    benchmarks: []
  },
  OPTIMIZE: {
    label: 'FM Optimize (N1)',
    kpis: [
      { key: 'Optimize (N1) - L1 Lead MKT',    label: 'L1 Lead MKT' },
      { key: 'Optimize (N1) - L4 Trial MKT',   label: 'L4 Trial MKT' },
      { key: 'Optimize (N1) - L6 Deal MKT',    label: 'L6 Deal MKT' },
      { key: 'Optimize (N1) - CR16%',          label: 'CR16%' },
      { key: 'Optimize (N1) - CR46%',          label: 'CR46%' },
      { key: 'Optimize (N1) - AOV',            label: 'AOV' },
      { key: 'Optimize (N1) - Doanh số N1',    label: 'Doanh số N1', isRevenue: true }
    ],
    benchmarks: [
      { label: 'CR16', value: '15%', desc: 'L1→L6' },
      { label: 'CR46', value: '50%', desc: 'L4→L6' },
      { label: 'AOV',  value: '18M', desc: 'Giá trị đơn hàng' }
    ]
  },
  OPS: {
    label: 'FM Ops (N2)',
    kpis: [
      { key: 'Ops (N2) - Số PHHS tiềm năng Re/Up', label: 'PHHS tiềm năng Re/Up' },
      { key: 'Ops (N2) - L6 Re/Upsell',             label: 'L6 Re/Upsell' },
      { key: 'Ops (N2) - L2 Referral',              label: 'L2 Referral' },
      { key: 'Ops (N2) - L6 Referral',              label: 'L6 Referral' },
      { key: 'Ops (N2) - Doanh số N2',              label: 'Doanh số N2', isRevenue: true }
    ],
    benchmarks: []
  }
};

/** Render benchmark bar khi filter function active */
function renderFuncBenchmarkBar(func) {
  const bar = document.getElementById('func-benchmark-bar');
  if (!bar) return;

  if (func === 'ALL') {
    bar.style.display = 'none';
    return;
  }

  const cfg = FUNC_KPI_CONFIG[func];
  if (!cfg || !cfg.benchmarks || cfg.benchmarks.length === 0) {
    bar.style.display = 'none';
    return;
  }

  const color = FUNC_COLORS[func] || FUNC_COLORS.OPTIMIZE;
  bar.style.display = 'flex';
  bar.innerHTML = `
    <span style="font-size:11px;font-weight:700;color:${color.main};text-transform:uppercase;letter-spacing:0.5px;margin-right:12px">
      Benchmark ${cfg.label}:
    </span>
    ${cfg.benchmarks.map(b => `
      <span class="func-benchmark-item" style="background:${color.bg};border-color:${color.border};color:${color.main}">
        <strong>${escHtml(b.label)}</strong> = ${escHtml(b.value)}
        <span style="opacity:0.7;font-weight:400"> · ${escHtml(b.desc)}</span>
      </span>
    `).join('')}
  `;
}

/** Tổng hợp KPI theo function từ allData */
function aggregateFuncKPIs(func, allData, week) {
  if (!allData || func === 'ALL') return null;
  const cfg = FUNC_KPI_CONFIG[func];
  if (!cfg) return null;

  const pace = WEEK_PACE[week] || WEEK_PACE[1];
  const results = {};

  cfg.kpis.forEach(kpiDef => {
    let totalTarget = 0, totalActual = 0, counted = 0, meetPace = 0;

    BU_LIST.forEach(buName => {
      const buData = allData[buName] || { WR: [] };
      const wr = buData.WR || [];
      const row = wr.find(r => (r.kpi || '').trim() === kpiDef.key);
      if (!row) return;

      const t = parseFloat(row.target);
      const a = parseFloat(row.actual);
      if (isNaN(t) || isNaN(a) || t <= 0) return;

      totalTarget += t;
      totalActual += a;
      counted++;
      if (a / t >= pace) meetPace++;
    });

    const pct = totalTarget > 0 ? Math.round(totalActual / totalTarget * 100) : null;
    results[kpiDef.key] = { label: kpiDef.label, totalTarget, totalActual, pct, counted, meetPace, isRevenue: kpiDef.isRevenue };
  });

  return results;
}

/** Render function KPI cards */
function renderFuncKPICards(func, allData, week) {
  const container = document.getElementById('func-kpi-cards');
  if (!container) return;

  if (func === 'ALL' || !allData) {
    container.style.display = 'none';
    return;
  }

  const cfg = FUNC_KPI_CONFIG[func];
  if (!cfg) {
    container.style.display = 'none';
    return;
  }

  const agg = aggregateFuncKPIs(func, allData, week);
  if (!agg) {
    container.style.display = 'none';
    return;
  }

  const color = FUNC_COLORS[func];
  const pace = WEEK_PACE[week] || WEEK_PACE[1];

  container.style.display = 'grid';
  container.innerHTML = cfg.kpis.map(kpiDef => {
    const d = agg[kpiDef.key];
    if (!d) return '';
    const pctVal = d.pct !== null ? d.pct : null;
    let pctColor = color.main;
    if (pctVal !== null) {
      if (pctVal >= pace * 100) pctColor = '#1a7a3a';
      else if (pctVal < pace * 50) pctColor = '#cc0000';
      else pctColor = '#c07000';
    }
    return `
      <div class="func-kpi-card" style="border-top-color:${color.main}">
        <div class="func-kpi-card-label" style="color:${color.main}">${escHtml(d.label)}</div>
        <div class="func-kpi-card-pct" style="color:${pctColor}">
          ${pctVal !== null ? pctVal + '%' : '—'}
        </div>
        <div class="func-kpi-card-detail">
          <span>T: ${d.totalTarget > 0 ? d.totalTarget.toLocaleString('vi-VN') : '—'}</span>
          <span>A: ${d.totalActual > 0 ? d.totalActual.toLocaleString('vi-VN') : '—'}</span>
        </div>
        <div class="func-kpi-card-pace">
          ${d.counted > 0 ? `${d.meetPace}/${d.counted} BU đạt pace` : 'Chưa có dữ liệu'}
        </div>
        ${d.isRevenue ? `<div class="func-kpi-card-revenue-badge">💰 Doanh số</div>` : ''}
      </div>
    `;
  }).join('');
}

// ===================== CONFIG MANAGEMENT =====================

// Defaults — used when no config exists in sheet
const CONFIG_DEFAULTS = {
  kpi_list: {
    list: WR_DEFAULT_ROWS.map(r => r.kpi)
  },
  benchmarks: {
    cr16: 15, cr46: 50, aov: 18,
    cr16_n1: 15, cr46_n1: 50, aov_n1: 18, noshow: 20,
    cr_reupsell: 15, cr_referral: 5, aov_n2: 15,
    cr16_n3: 10, cr46_n3: 40, aov_n3: 16
  },
  week_pace: {
    w1: 0.15,
    w2: 0.40,
    w3: 0.70,
    w4: 1.00
  },
  rating_thresholds: {
    tot: 80,
    daydu: 60,
    hoihot: 40
  },
  bu_list: {
    list: [...BU_LIST]
  }
};

/** Toggle config section accordion */
function toggleConfigSection(headerEl) {
  const section = headerEl.parentElement;
  section.classList.toggle('open');
}

/** Initialize Config view when navigated to */
function initConfigView() {
  initKPIImportMonthSelector();
  loadConfigMonths().then(() => {
    // Default to current month
    const now = new Date();
    const currentMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
    const sel = document.getElementById('config-month');
    if (sel && sel.value) {
      loadConfigForMonth(sel.value);
    } else {
      state.config.month = currentMonth;
      populateConfigMonthSelector(currentMonth);
      loadConfigForMonth(currentMonth);
    }
  });

  // Open all sections by default
  document.querySelectorAll('.config-section').forEach(s => s.classList.add('open'));

  // Show role badge
  showConfigRoleBadge();

  // Live BU count
  const buTA = document.getElementById('cfg-bu-list');
  if (buTA) buTA.addEventListener('input', updateBUCount);
}

/** Show SOD/FM role badge */
function showConfigRoleBadge() {
  const badge = document.getElementById('config-role-badge');
  if (!badge) return;
  // For now, show role selection — SOD vs FM
  // Ask user which role (simple prompt approach: use a dropdown or just default to SOD)
  badge.innerHTML = `
    <select id="config-role-select" class="topbar-select" style="font-size:11px;min-width:90px" onchange="onConfigRoleChange()">
      <option value="sod">SOD (chỉnh sửa)</option>
      <option value="fm">FM (chỉ xem)</option>
    </select>
  `;
  state.config.role = 'sod';
  applyConfigRole();
}

function onConfigRoleChange() {
  const sel = document.getElementById('config-role-select');
  state.config.role = sel ? sel.value : 'sod';
  applyConfigRole();
}

/** Apply role-based restrictions */
function applyConfigRole() {
  const isReadOnly = state.config.role === 'fm';
  const saveBtn = document.getElementById('btn-config-save');
  const copyBtn = document.getElementById('btn-config-copy');

  if (saveBtn) saveBtn.style.display = isReadOnly ? 'none' : '';
  if (copyBtn) copyBtn.style.display = isReadOnly ? 'none' : '';

  // All textareas
  document.querySelectorAll('#view-config .config-textarea').forEach(ta => {
    ta.readOnly = isReadOnly;
  });
  // All inputs
  document.querySelectorAll('#view-config .config-field input').forEach(inp => {
    inp.readOnly = isReadOnly;
  });
}

/** Load list of months that have config */
async function loadConfigMonths() {
  try {
    const url = `${APPS_SCRIPT_URL}?action=get_config_months`;
    const res = await fetch(url);
    const json = await res.json();
    state.config.months = json.months || [];
  } catch (e) {
    console.error('loadConfigMonths error:', e);
    state.config.months = [];
  }
}

/** Populate month selector with all MONTHS + highlight which have config */
function populateConfigMonthSelector(selectedMonth) {
  const sel = document.getElementById('config-month');
  if (!sel) return;
  sel.innerHTML = '';
  MONTHS.forEach(m => {
    const opt = document.createElement('option');
    opt.value = m;
    const hasConfig = state.config.months.includes(m);
    opt.textContent = m + (hasConfig ? ' ✓' : '');
    if (m === selectedMonth) opt.selected = true;
    sel.appendChild(opt);
  });
}

function onConfigMonthChange() {
  const sel = document.getElementById('config-month');
  if (!sel) return;
  loadConfigForMonth(sel.value);
}

/** Load config for a specific month from sheet */
async function loadConfigForMonth(month) {
  state.config.month = month;
  state.config.dirty = false;
  updateConfigSaveStatus('');

  try {
    const url = `${APPS_SCRIPT_URL}?action=get_config&month=${encodeURIComponent(month)}`;
    const res = await fetch(url);
    const json = await res.json();

    if (json.config && Object.keys(json.config).length > 0) {
      state.config.data = json.config;
    } else {
      // No config for this month — use defaults
      state.config.data = JSON.parse(JSON.stringify(CONFIG_DEFAULTS));
    }
  } catch (e) {
    console.error('loadConfigForMonth error:', e);
    state.config.data = JSON.parse(JSON.stringify(CONFIG_DEFAULTS));
  }

  populateConfigUI();

  // Also load KPI targets for this month and display (force refresh)
  loadKPITargets(month, true).then(targets => {
    if (targets && targets.length > 0) {
      displayImportedKPI(targets, month);
    }
  }).catch(() => {});
}

/** Populate all config form fields from state.config.data */
function populateConfigUI() {
  const d = state.config.data;

  // KPI list
  const kpiTA = document.getElementById('cfg-kpi-list');
  if (kpiTA) {
    const kpis = d.kpi_list && d.kpi_list.list ? d.kpi_list.list : CONFIG_DEFAULTS.kpi_list.list;
    kpiTA.value = (Array.isArray(kpis) ? kpis : []).join('\n');
  }

  // Benchmarks — Blended
  setVal('cfg-bench-cr16', getNestedVal(d, 'benchmarks', 'cr16', CONFIG_DEFAULTS.benchmarks.cr16));
  setVal('cfg-bench-cr46', getNestedVal(d, 'benchmarks', 'cr46', CONFIG_DEFAULTS.benchmarks.cr46));
  setVal('cfg-bench-aov', getNestedVal(d, 'benchmarks', 'aov', CONFIG_DEFAULTS.benchmarks.aov));
  // Benchmarks — N1
  setVal('cfg-bench-cr16-n1', getNestedVal(d, 'benchmarks', 'cr16_n1', CONFIG_DEFAULTS.benchmarks.cr16_n1));
  setVal('cfg-bench-cr46-n1', getNestedVal(d, 'benchmarks', 'cr46_n1', CONFIG_DEFAULTS.benchmarks.cr46_n1));
  setVal('cfg-bench-aov-n1', getNestedVal(d, 'benchmarks', 'aov_n1', CONFIG_DEFAULTS.benchmarks.aov_n1));
  setVal('cfg-bench-noshow', getNestedVal(d, 'benchmarks', 'noshow', CONFIG_DEFAULTS.benchmarks.noshow));
  // Benchmarks — N2
  setVal('cfg-bench-cr-reupsell', getNestedVal(d, 'benchmarks', 'cr_reupsell', CONFIG_DEFAULTS.benchmarks.cr_reupsell));
  setVal('cfg-bench-cr-referral', getNestedVal(d, 'benchmarks', 'cr_referral', CONFIG_DEFAULTS.benchmarks.cr_referral));
  setVal('cfg-bench-aov-n2', getNestedVal(d, 'benchmarks', 'aov_n2', CONFIG_DEFAULTS.benchmarks.aov_n2));
  // Benchmarks — N3
  setVal('cfg-bench-cr16-n3', getNestedVal(d, 'benchmarks', 'cr16_n3', CONFIG_DEFAULTS.benchmarks.cr16_n3));
  setVal('cfg-bench-cr46-n3', getNestedVal(d, 'benchmarks', 'cr46_n3', CONFIG_DEFAULTS.benchmarks.cr46_n3));
  setVal('cfg-bench-aov-n3', getNestedVal(d, 'benchmarks', 'aov_n3', CONFIG_DEFAULTS.benchmarks.aov_n3));

  // Week Pace
  setVal('cfg-pace-w1', getNestedVal(d, 'week_pace', 'w1', CONFIG_DEFAULTS.week_pace.w1));
  setVal('cfg-pace-w2', getNestedVal(d, 'week_pace', 'w2', CONFIG_DEFAULTS.week_pace.w2));
  setVal('cfg-pace-w3', getNestedVal(d, 'week_pace', 'w3', CONFIG_DEFAULTS.week_pace.w3));
  setVal('cfg-pace-w4', getNestedVal(d, 'week_pace', 'w4', CONFIG_DEFAULTS.week_pace.w4));

  // Rating thresholds
  setVal('cfg-rating-tot', getNestedVal(d, 'rating_thresholds', 'tot', CONFIG_DEFAULTS.rating_thresholds.tot));
  setVal('cfg-rating-daydu', getNestedVal(d, 'rating_thresholds', 'daydu', CONFIG_DEFAULTS.rating_thresholds.daydu));
  setVal('cfg-rating-hoihot', getNestedVal(d, 'rating_thresholds', 'hoihot', CONFIG_DEFAULTS.rating_thresholds.hoihot));

  // BU list
  const buTA = document.getElementById('cfg-bu-list');
  if (buTA) {
    const bus = d.bu_list && d.bu_list.list ? d.bu_list.list : CONFIG_DEFAULTS.bu_list.list;
    buTA.value = (Array.isArray(bus) ? bus : []).join('\n');
  }
  updateBUCount();

  applyConfigRole();
}

/** Helper: set input value */
function setVal(id, val) {
  const el = document.getElementById(id);
  if (el) el.value = val;
}

/** Helper: get nested config value with fallback */
function getNestedVal(data, section, key, fallback) {
  if (data && data[section] && data[section][key] !== undefined) return data[section][key];
  return fallback;
}

/** Read all form fields back into a config object */
function readConfigFromUI() {
  const config = {};

  // KPI list
  const kpiTA = document.getElementById('cfg-kpi-list');
  const kpiLines = kpiTA ? kpiTA.value.split('\n').map(l => l.trim()).filter(l => l) : [];
  config.kpi_list = { list: kpiLines };

  // Benchmarks
  config.benchmarks = {
    cr16: parseFloat(document.getElementById('cfg-bench-cr16')?.value) || 15,
    cr46: parseFloat(document.getElementById('cfg-bench-cr46')?.value) || 50,
    aov: parseFloat(document.getElementById('cfg-bench-aov')?.value) || 18,
    cr16_n1: parseFloat(document.getElementById('cfg-bench-cr16-n1')?.value) || 15,
    cr46_n1: parseFloat(document.getElementById('cfg-bench-cr46-n1')?.value) || 50,
    aov_n1: parseFloat(document.getElementById('cfg-bench-aov-n1')?.value) || 18,
    noshow: parseFloat(document.getElementById('cfg-bench-noshow')?.value) || 20,
    cr_reupsell: parseFloat(document.getElementById('cfg-bench-cr-reupsell')?.value) || 15,
    cr_referral: parseFloat(document.getElementById('cfg-bench-cr-referral')?.value) || 5,
    aov_n2: parseFloat(document.getElementById('cfg-bench-aov-n2')?.value) || 15,
    cr16_n3: parseFloat(document.getElementById('cfg-bench-cr16-n3')?.value) || 10,
    cr46_n3: parseFloat(document.getElementById('cfg-bench-cr46-n3')?.value) || 40,
    aov_n3: parseFloat(document.getElementById('cfg-bench-aov-n3')?.value) || 16
  };

  // Week Pace
  config.week_pace = {
    w1: parseFloat(document.getElementById('cfg-pace-w1')?.value) || 0.15,
    w2: parseFloat(document.getElementById('cfg-pace-w2')?.value) || 0.40,
    w3: parseFloat(document.getElementById('cfg-pace-w3')?.value) || 0.70,
    w4: parseFloat(document.getElementById('cfg-pace-w4')?.value) || 1.00
  };

  // Rating thresholds
  config.rating_thresholds = {
    tot: parseInt(document.getElementById('cfg-rating-tot')?.value) || 80,
    daydu: parseInt(document.getElementById('cfg-rating-daydu')?.value) || 60,
    hoihot: parseInt(document.getElementById('cfg-rating-hoihot')?.value) || 40
  };

  // BU list
  const buTA = document.getElementById('cfg-bu-list');
  const buLines = buTA ? buTA.value.split('\n').map(l => l.trim()).filter(l => l) : [];
  config.bu_list = { list: buLines };

  return config;
}

/** Save config to sheet */
async function saveAllConfig() {
  if (state.config.role === 'fm') {
    showToast('FM chỉ có quyền xem, không thể chỉnh sửa cấu hình.', 'error');
    return;
  }

  const month = state.config.month;
  if (!month) {
    showToast('Chưa chọn tháng.', 'error');
    return;
  }

  const config = readConfigFromUI();
  updateConfigSaveStatus('Đang lưu...');

  try {
    const res = await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify({
        action: 'save_config',
        month: month,
        config: config
      })
    });
    const json = await res.json();
    if (json.success) {
      state.config.data = config;
      state.config.dirty = false;
      if (!state.config.months.includes(month)) {
        state.config.months.push(month);
        state.config.months.sort();
      }
      populateConfigMonthSelector(month);
      updateConfigSaveStatus('Đã lưu ✓');
      showToast(`Cấu hình tháng ${month} đã được lưu thành công.`, 'success');
    } else {
      updateConfigSaveStatus('Lỗi!');
      showToast('Lỗi lưu cấu hình: ' + (json.error || 'Unknown'), 'error');
    }
  } catch (e) {
    updateConfigSaveStatus('Lỗi!');
    showToast('Lỗi kết nối khi lưu cấu hình.', 'error');
    console.error('saveAllConfig error:', e);
  }
}

/** Copy config from previous month */
async function copyConfigFromPrev() {
  const currentMonth = state.config.month;
  if (!currentMonth) return;

  // Find previous month
  const idx = MONTHS.indexOf(currentMonth);
  if (idx <= 0) {
    showToast('Không có tháng trước để sao chép.', 'error');
    return;
  }
  const prevMonth = MONTHS[idx - 1];

  // Check if prev month has config
  if (!state.config.months.includes(prevMonth)) {
    showToast(`Tháng ${prevMonth} chưa có cấu hình. Sẽ sao chép giá trị mặc định.`, 'info');
    // Load defaults
    state.config.data = JSON.parse(JSON.stringify(CONFIG_DEFAULTS));
    populateConfigUI();
    return;
  }

  try {
    showToast(`Đang sao chép cấu hình từ ${prevMonth}...`, 'info');
    const url = `${APPS_SCRIPT_URL}?action=get_config&month=${encodeURIComponent(prevMonth)}`;
    const res = await fetch(url);
    const json = await res.json();
    if (json.config && Object.keys(json.config).length > 0) {
      state.config.data = json.config;
      populateConfigUI();
      showToast(`Đã sao chép cấu hình từ ${prevMonth}. Nhấn "Lưu Cấu hình" để áp dụng cho ${currentMonth}.`, 'success');
    } else {
      state.config.data = JSON.parse(JSON.stringify(CONFIG_DEFAULTS));
      populateConfigUI();
    }
  } catch (e) {
    showToast('Lỗi khi sao chép cấu hình.', 'error');
    console.error('copyConfigFromPrev error:', e);
  }
}

function updateConfigSaveStatus(text) {
  const el = document.getElementById('config-save-status');
  if (el) el.textContent = text;
}

function updateBUCount() {
  const ta = document.getElementById('cfg-bu-list');
  const el = document.getElementById('cfg-bu-count');
  if (!ta || !el) return;
  const lines = ta.value.split('\n').map(l => l.trim()).filter(l => l);
  el.textContent = `${lines.length} BU`;
}

// ===================== KPI IMPORT (Excel Upload) =====================

// Temp storage for parsed KPI before confirm
let _kpiPendingImport = null;

/** Normalize BU name: remove extra spaces, standardize separators */
function normalizeBU(name) {
  return String(name).trim()
    .replace(/\s*-\s*/g, ' - ')  // normalize "HCM1-PVT" to "HCM1 - PVT"
    .replace(/\s+/g, ' ');
}

/** Find best matching BU from system list */
function matchBU(excelName, buList) {
  const norm = normalizeBU(excelName);
  // Exact match after normalization
  const exact = buList.find(b => normalizeBU(b) === norm);
  if (exact) return exact;
  // Fuzzy: remove all spaces & dashes, compare
  const key = norm.replace(/[\s\-]/g, '').toUpperCase();
  const fuzzy = buList.find(b => normalizeBU(b).replace(/[\s\-]/g, '').toUpperCase() === key);
  if (fuzzy) return fuzzy;
  return null; // no match
}

/** Handle file selection */
function onKPIFileSelected(input) {
  const file = input.files[0];
  if (!file) return;
  document.getElementById('cfg-kpi-filename').textContent = file.name;

  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const wb = XLSX.read(e.target.result, { type: 'array' });
      const sheetName = wb.SheetNames[0];
      const ws = wb.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      parseExcelKPI(json, sheetName);
    } catch (err) {
      showToast('L\u1ED7i \u0111\u1ECDc file: ' + err.message, 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

/** Parse Excel data and auto-detect columns */
function parseExcelKPI(rows, sheetName) {
  // Find header row (look for "BU" keyword)
  let headerIdx = -1;
  let colBU = -1, colN1 = -1, colN2 = -1, colN3 = -1, colTotal = -1;
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    const row = rows[i].map(c => String(c || '').trim().toUpperCase());
    const buCol = row.findIndex(c => c === 'BU' || c.includes('BU'));
    if (buCol >= 0) {
      headerIdx = i;
      colBU = buCol;
      // Find KPI columns by keywords
      row.forEach((c, j) => {
        if (j === buCol) return;
        const cl = c.toLowerCase();
        if (cl.includes('n1') || cl.includes('kpi n1')) colN1 = j;
        else if (cl.includes('n2') || cl.includes('kpi n2')) colN2 = j;
        else if (cl.includes('n3') || cl.includes('kpi n3')) colN3 = j;
        else if (cl.includes('t\u1ed5ng') || cl.includes('total') || cl.includes('tong')) colTotal = j;
      });
      break;
    }
  }

  if (headerIdx < 0 || colBU < 0) {
    showToast('Kh\u00F4ng t\u00ECm th\u1EA5y c\u1ED9t BU trong file. Ki\u1EC3m tra l\u1EA1i format.', 'error');
    return;
  }

  // If specific columns not found, use sequential columns after BU
  if (colN1 < 0) colN1 = colBU + 1;
  if (colN2 < 0) colN2 = colBU + 2;
  if (colN3 < 0) colN3 = colBU + 3;
  if (colTotal < 0) colTotal = colBU + 4;

  // Parse data rows
  const targets = [];
  let matchCount = 0, noMatchCount = 0;
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    const rawBU = String(row[colBU] || '').trim();
    if (!rawBU) continue;

    const n1 = Number(row[colN1]) || 0;
    const n2 = Number(row[colN2]) || 0;
    const n3 = Number(row[colN3]) || 0;
    const total = Number(row[colTotal]) || (n1 + n2 + n3);
    if (total === 0 && n1 === 0 && n2 === 0 && n3 === 0) continue;

    // Match BU name to system list
    const matchedBU = matchBU(rawBU, BU_LIST);

    targets.push({
      rawBU,
      bu: matchedBU || normalizeBU(rawBU),
      matched: !!matchedBU,
      kpi_n1: n1, kpi_n2: n2, kpi_n3: n3, kpi_total: total
    });
    if (matchedBU) matchCount++; else noMatchCount++;
  }

  if (targets.length === 0) {
    showToast('Kh\u00F4ng t\u00ECm th\u1EA5y d\u1EEF li\u1EC7u KPI trong file.', 'error');
    return;
  }

  _kpiPendingImport = targets;
  showKPIPreview(targets, sheetName, matchCount, noMatchCount);
}

/** Show preview table before confirming import */
function showKPIPreview(targets, sheetName, matchCount, noMatchCount) {
  const el = document.getElementById('cfg-kpi-preview');
  if (!el) return;

  const rows = targets.map(t => {
    const matchIcon = t.matched
      ? '<span style="color:var(--green)">\u2713</span>'
      : '<span style="color:var(--orange)" title="T\u00EAn BU kh\u00F4ng kh\u1EDBp h\u1EC7 th\u1ED1ng, \u0111\u00E3 t\u1EF1 \u0111\u1ED9ng chu\u1EA9n h\u00F3a">\u26A0</span>';
    const buDisplay = t.matched ? escHtml(t.bu)
      : `<span style="color:var(--orange)">${escHtml(t.rawBU)}</span> \u2192 <strong>${escHtml(t.bu)}</strong>`;
    return `<tr>
      <td>${matchIcon}</td>
      <td>${buDisplay}</td>
      <td style="text-align:right">${fmtDops(t.kpi_n1, 'money')}</td>
      <td style="text-align:right">${fmtDops(t.kpi_n2, 'money')}</td>
      <td style="text-align:right">${fmtDops(t.kpi_n3, 'money')}</td>
      <td style="text-align:right;font-weight:700">${fmtDops(t.kpi_total, 'money')}</td>
    </tr>`;
  }).join('');

  el.innerHTML = `
    <div style="font-size:12px;color:var(--gray-600);margin-bottom:8px">
      Sheet: <strong>${escHtml(sheetName)}</strong> \u2014
      ${targets.length} BU \u2014
      <span style="color:var(--green)">${matchCount} match</span>
      ${noMatchCount > 0 ? `<span style="color:var(--orange)">, ${noMatchCount} t\u1EF1 chu\u1EA9n h\u00F3a</span>` : ''}
    </div>
    <div style="max-height:350px;overflow-y:auto;border:1px solid var(--gray-200);border-radius:4px">
      <table class="data-table" style="font-size:12px">
        <thead><tr>
          <th style="width:30px"></th><th>BU</th><th style="text-align:right">KPI N1</th><th style="text-align:right">KPI N2</th><th style="text-align:right">KPI N3</th><th style="text-align:right">KPI T\u1ED5ng</th>
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;

  document.getElementById('cfg-kpi-actions').style.display = '';
}

/** Confirm and save KPI import (using GET-based batch approach) */
async function confirmImportKPI() {
  const month = document.getElementById('cfg-kpi-import-month')?.value;
  const statusEl = document.getElementById('cfg-kpi-import-status');
  if (!month) { showToast('Ch\u01B0a ch\u1ECDn th\u00E1ng', 'error'); return; }
  if (!_kpiPendingImport || _kpiPendingImport.length === 0) { showToast('Ch\u01B0a c\u00F3 d\u1EEF li\u1EC7u \u0111\u1EC3 import', 'error'); return; }

  const targets = _kpiPendingImport.map(t => ({
    bu: t.bu, kpi_n1: t.kpi_n1, kpi_n2: t.kpi_n2, kpi_n3: t.kpi_n3, kpi_total: t.kpi_total
  }));

  if (statusEl) statusEl.textContent = '\u0110ang x\u00F3a d\u1EEF li\u1EC7u c\u0169...';

  try {
    // Step 1: Clear existing KPI for this month (GET)
    const clearResult = await apiFetch('clear_kpi_targets', { month });
    if (!clearResult.success) throw new Error(clearResult.error || 'Clear failed');

    // Step 2: Save in batches of 10 via GET (to avoid URL length limits)
    const BATCH_SIZE = 10;
    let saved = 0;
    for (let i = 0; i < targets.length; i += BATCH_SIZE) {
      const batch = targets.slice(i, i + BATCH_SIZE);
      if (statusEl) statusEl.textContent = `\u0110ang l\u01B0u... ${saved}/${targets.length}`;
      const batchResult = await apiFetch('save_kpi_targets_batch', {
        month,
        data: JSON.stringify(batch)
      });
      if (!batchResult.success) throw new Error(batchResult.error || 'Batch save failed');
      saved += batch.length;
    }

    showToast(`\u2713 \u0110\u00E3 import ${saved} BU KPI cho th\u00E1ng ${month}`);
    if (statusEl) statusEl.textContent = `\u2713 \u0110\u00E3 l\u01B0u ${saved} BU`;
    state.dashboard._kpiTargets = state.dashboard._kpiTargets || {};
    state.dashboard._kpiTargets[month] = targets;
    _kpiPendingImport = null;
    document.getElementById('cfg-kpi-actions').style.display = 'none';
    displayImportedKPI(targets, month);
  } catch (e) {
    showToast('L\u1ED7i: ' + e.message, 'error');
    if (statusEl) statusEl.textContent = 'L\u1ED7i!';
  }
}

/** Display imported KPI in a table */
function displayImportedKPI(targets, month) {
  const el = document.getElementById('cfg-kpi-imported');
  if (!el || !targets || targets.length === 0) {
    if (el) el.innerHTML = '';
    return;
  }
  const rows = targets.map(t => `
    <tr>
      <td style="font-weight:600">${escHtml(t.bu)}</td>
      <td style="text-align:right">${fmtDops(t.kpi_n1 || 0, 'money')}</td>
      <td style="text-align:right">${fmtDops(t.kpi_n2 || 0, 'money')}</td>
      <td style="text-align:right">${fmtDops(t.kpi_n3 || 0, 'money')}</td>
      <td style="text-align:right;font-weight:700">${fmtDops(t.kpi_total || 0, 'money')}</td>
    </tr>
  `).join('');

  el.innerHTML = `
    <div style="font-size:12px;font-weight:700;color:var(--gray-600);margin-bottom:6px">KPI \u0111\u00E3 import cho th\u00E1ng ${month} (${targets.length} BU)</div>
    <div style="max-height:300px;overflow-y:auto;border:1px solid var(--gray-200);border-radius:4px">
      <table class="data-table" style="font-size:12px">
        <thead><tr>
          <th>BU</th><th style="text-align:right">KPI N1</th><th style="text-align:right">KPI N2</th><th style="text-align:right">KPI N3</th><th style="text-align:right">KPI T\u1ED5ng</th>
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;
}

/** Load KPI targets for a month (called from dashboard) */
async function loadKPITargets(month, forceRefresh) {
  if (!isConfigured()) return null;
  state.dashboard._kpiTargets = state.dashboard._kpiTargets || {};
  if (!forceRefresh && state.dashboard._kpiTargets[month]) return state.dashboard._kpiTargets[month];
  try {
    const result = await apiFetch('get_kpi_targets', { month });
    if (result.success && result.targets) {
      state.dashboard._kpiTargets[month] = result.targets;
      return result.targets;
    }
  } catch (e) { console.warn('loadKPITargets error:', e.message); }
  return null;
}

/** Calculate pace (expected % of target achieved by today) */
function calcPacePct(month) {
  const [yyyy, mm] = month.split('-').map(Number);
  const daysInMonth = new Date(yyyy, mm, 0).getDate();
  const today = new Date();
  const currentDay = (today.getFullYear() === yyyy && today.getMonth() + 1 === mm)
    ? today.getDate() : daysInMonth;
  return Math.min(((currentDay - 1) / daysInMonth) * 100, 100);
}

/** Get run rate status for a BU */
function getRunRateStatus(actualPct, pacePct) {
  const diff = actualPct - pacePct;
  if (diff >= 5) return { label: 'Nhanh', cls: 'status-done', icon: '\u2191' };
  if (diff >= -5) return { label: '\u0110\u00FAng ti\u1EBFn \u0111\u1ED9', cls: 'status-badge-blue', icon: '\u2192' };
  if (diff >= -15) return { label: 'Ch\u1EADm', cls: 'status-warn', icon: '\u2193' };
  return { label: 'Nguy hi\u1EC3m', cls: 'status-alert', icon: '\u2193\u2193' };
}

/** Populate KPI import month selector */
function initKPIImportMonthSelector() {
  const sel = document.getElementById('cfg-kpi-import-month');
  if (!sel) return;
  sel.innerHTML = '';
  const now = new Date();
  for (let i = -1; i <= 3; i++) {
    const d = new Date(now.getFullYear(), now.getMonth() + i, 1);
    const val = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
    const opt = document.createElement('option');
    opt.value = val;
    opt.textContent = `Th\u00E1ng ${d.getMonth()+1}/${d.getFullYear()}`;
    if (i === 0) opt.selected = true;
    sel.appendChild(opt);
  }
}

// ===================== CONFIG-AWARE HELPERS =====================
// These functions load config values dynamically for scoring

/** Get active config for a given month — loads from state or returns defaults */
function getActiveConfig(month) {
  // If we have config data loaded for this month, return it
  // For scoring, we need a synchronous approach — cache configs
  return _configCache[month] || null;
}

// Cache for configs loaded for scoring
const _configCache = {};

/** Preload config for a month (called before scoring) */
async function preloadConfigForMonth(month) {
  if (_configCache[month]) return _configCache[month];
  try {
    const url = `${APPS_SCRIPT_URL}?action=get_config&month=${encodeURIComponent(month)}`;
    const res = await fetch(url);
    const json = await res.json();
    if (json.config && Object.keys(json.config).length > 0) {
      _configCache[month] = json.config;
    } else {
      _configCache[month] = null;
    }
  } catch (e) {
    _configCache[month] = null;
  }
  return _configCache[month];
}

/** Get WEEK_PACE from config or use default */
function getWeekPace(configData) {
  if (configData && configData.week_pace) {
    return {
      1: configData.week_pace.w1 || WEEK_PACE[1],
      2: configData.week_pace.w2 || WEEK_PACE[2],
      3: configData.week_pace.w3 || WEEK_PACE[3],
      4: configData.week_pace.w4 || WEEK_PACE[4]
    };
  }
  return WEEK_PACE;
}

/** Get rating thresholds from config or use default */
function getRatingFromConfig(total, configData) {
  const t = configData && configData.rating_thresholds ? configData.rating_thresholds : null;
  const thTot = t ? (t.tot || 80) : 80;
  const thDaydu = t ? (t.daydu || 60) : 60;
  const thHoihot = t ? (t.hoihot || 40) : 40;

  if (total >= thTot) return { label: 'Tốt', cls: 'tot' };
  if (total >= thDaydu) return { label: 'Đầy đủ', cls: 'day-du' };
  if (total >= thHoihot) return { label: 'Hời hợt', cls: 'hoi-hot' };
  return { label: 'Cần cải thiện', cls: 'can-cb' };
}

// ============================================================
// DAILY REPORT (Báo cáo hàng ngày)
// Không cần mật khẩu — dành cho CM
// ============================================================

/** Initialize daily view — set default date, render form fields */
function initDailyView() {
  // Auto-select BU from CM context if navigated via tab
  if (state.daily.pendingBU) {
    const buSelect = document.getElementById('daily-bu');
    if (buSelect) {
      buSelect.value = state.daily.pendingBU;
      state.daily.bu = state.daily.pendingBU;
    }
    delete state.daily.pendingBU;
  }

  const dateEl = document.getElementById('daily-date');
  if (dateEl && !dateEl.value) {
    const today = new Date();
    const yyyy = today.getFullYear();
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const dd = String(today.getDate()).padStart(2, '0');
    dateEl.value = `${yyyy}-${mm}-${dd}`;
    state.daily.date = dateEl.value;
  }
  renderDailyForm();
  // Auto-load if BU is set
  if (state.daily.bu && state.daily.date) {
    loadDailyReport();
  }
}

/** Render input fields into each section grid */
function renderDailyForm() {
  const groups = ['N1', 'N2', 'N3'];
  groups.forEach(group => {
    const container = document.getElementById(`daily-fields-${group}`);
    if (!container) return;
    container.innerHTML = '';
    const fields = DAILY_FIELDS.filter(f => f.group === group);
    fields.forEach(f => {
      const div = document.createElement('div');
      div.className = 'daily-field';
      const label = document.createElement('label');
      label.textContent = f.label;
      div.appendChild(label);
      if (f.type === 'number') {
        const input = document.createElement('input');
        input.type = 'number';
        input.id = `daily-${f.id}`;
        input.min = '0';
        input.value = state.daily.data[f.id] || 0;
        input.addEventListener('input', () => {
          updateDailyData(f.id, parseInt(input.value) || 0);
        });
        div.appendChild(input);
      } else {
        const textarea = document.createElement('textarea');
        textarea.id = `daily-${f.id}`;
        textarea.rows = 2;
        textarea.placeholder = f.label;
        textarea.value = state.daily.data[f.id] || '';
        textarea.addEventListener('input', () => {
          updateDailyData(f.id, textarea.value);
        });
        div.appendChild(textarea);
      }
      container.appendChild(div);
    });
  });
  updateDailyTotals();
}

/** Update a daily data field and recalculate totals */
function updateDailyData(fieldId, value) {
  state.daily.data[fieldId] = value;
  state.daily.loaded = true; // mark dirty
  updateDailyTotals();
  // Update saved indicator
  const indicator = document.getElementById('daily-saved-indicator');
  if (indicator) {
    indicator.textContent = 'Chưa lưu';
    indicator.classList.remove('show');
  }
}

/** Recalculate totals: Deal, Doanh số */
function updateDailyTotals() {
  const d = state.daily.data;
  const dealN1 = parseInt(d.deal_n1) || 0;
  const dealReup = parseInt(d.deal_reupsell) || 0;
  const dealRef = parseInt(d.deal_referral) || 0;
  const dealN3 = parseInt(d.deal_n3) || 0;
  const totalDeal = dealN1 + dealReup + dealRef + dealN3;

  const revN1 = parseInt(d.revenue_n1) || 0;
  const revN2 = parseInt(d.revenue_n2) || 0;
  const revN3 = parseInt(d.revenue_n3) || 0;
  const totalRev = revN1 + revN2 + revN3;

  const el1 = document.getElementById('daily-total-deal');
  if (el1) el1.value = totalDeal;
  const el2 = document.getElementById('daily-total-revenue');
  if (el2) el2.value = totalRev;
}

/** When BU or date changes, load existing data */
function onDailySelectionChange() {
  const bu = document.getElementById('daily-bu').value;
  const date = document.getElementById('daily-date').value;
  state.daily.bu = bu;
  state.daily.date = date;

  if (bu && date) {
    loadDailyReport();
  }
}

/** Load existing daily report for selected BU + date */
async function loadDailyReport() {
  const { bu, date } = state.daily;
  if (!bu || !date) return;

  try {
    const result = await apiFetch('getDaily', { bu, date });
    if (result.ok && result.data && result.data.length > 0) {
      const row = result.data[0];
      // Populate state from loaded data
      state.daily.data = {};
      DAILY_FIELDS.forEach(f => {
        const val = row[f.id];
        state.daily.data[f.id] = f.type === 'number' ? (parseInt(val) || 0) : (val || '');
      });
      // Update indicator
      const indicator = document.getElementById('daily-saved-indicator');
      if (indicator) {
        indicator.textContent = `Đã lưu: ${row.saved_at || ''}`;
        indicator.classList.add('show');
      }
    } else {
      // No data yet — reset to defaults
      state.daily.data = {};
      DAILY_FIELDS.forEach(f => {
        state.daily.data[f.id] = f.type === 'number' ? 0 : '';
      });
      const indicator = document.getElementById('daily-saved-indicator');
      if (indicator) {
        indicator.textContent = 'Chưa có dữ liệu';
        indicator.classList.remove('show');
      }
    }
    // Re-render form with loaded data
    renderDailyForm();
    // Update note field
    const noteEl = document.getElementById('daily-note');
    if (noteEl) noteEl.value = state.daily.data.note || '';
    // Load MTD
    loadDailyMTD();
  } catch (err) {
    console.error('loadDailyReport error:', err);
    showToast('Lỗi tải daily report: ' + err.message);
  }
}

/** Save daily report via POST */
async function saveDailyReport() {
  const bu = document.getElementById('daily-bu').value;
  const date = document.getElementById('daily-date').value;

  if (!bu) { showToast('Vui lòng chọn BU'); return; }
  if (!date) { showToast('Vui lòng chọn ngày'); return; }

  const btn = document.getElementById('btn-save-daily');
  if (btn) { btn.disabled = true; btn.textContent = 'Đang lưu...'; }

  // Build row data
  const rowData = { bu, date };
  DAILY_FIELDS.forEach(f => {
    const val = state.daily.data[f.id];
    rowData[f.id] = f.type === 'number' ? (parseInt(val) || 0) : (val || '');
  });
  rowData.saved_at = new Date().toISOString();

  try {
    const result = await apiPost('saveDaily', 'DAILY', { rows: [rowData] });
    if (result.ok) {
      const indicator = document.getElementById('daily-saved-indicator');
      if (indicator) {
        indicator.textContent = 'Đã lưu';
        indicator.classList.add('show');
      }
      showToast('Đã lưu Daily Report thành công');
      // Reload MTD after save
      loadDailyMTD();
    } else {
      showToast('Lỗi lưu: ' + (result.error || 'Unknown'));
    }
  } catch (err) {
    console.error('saveDailyReport error:', err);
    showToast('Lỗi lưu daily report: ' + err.message);
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg> Lưu Daily Report';
    }
  }
}

/** Load MTD (month-to-date) summary for selected BU */
async function loadDailyMTD() {
  const { bu, date } = state.daily;
  if (!bu || !date) return;

  // Extract month from date (YYYY-MM)
  const month = date.substring(0, 7);

  try {
    const result = await apiFetch('getDailyMTD', { bu, month });
    const panel = document.getElementById('daily-mtd-panel');
    if (!panel) return;

    if (result.ok && result.data && result.data.length > 0) {
      state.daily.mtd = result.data;
      panel.style.display = '';

      // Update meta
      const metaEl = document.getElementById('daily-mtd-meta');
      if (metaEl) metaEl.textContent = `${bu} — Tháng ${month} — ${result.data.length} ngày`;

      // Calculate MTD sums
      const sums = {};
      DAILY_FIELDS.forEach(f => {
        if (f.type === 'number') sums[f.id] = 0;
      });
      result.data.forEach(row => {
        DAILY_FIELDS.forEach(f => {
          if (f.type === 'number') {
            sums[f.id] += parseInt(row[f.id]) || 0;
          }
        });
      });
      // Totals
      const totalDeal = (sums.deal_n1 || 0) + (sums.deal_reupsell || 0) + (sums.deal_referral || 0) + (sums.deal_n3 || 0);
      const totalRev = (sums.revenue_n1 || 0) + (sums.revenue_n2 || 0) + (sums.revenue_n3 || 0);

      // Auto-calculated CRs
      const cr16 = sums.lead_mkt > 0 ? (sums.deal_n1 / sums.lead_mkt * 100) : 0;
      const cr46 = sums.trial_done_n1 > 0 ? (sums.deal_n1 / sums.trial_done_n1 * 100) : 0;
      const aov = sums.deal_n1 > 0 ? (sums.revenue_n1 / sums.deal_n1) : 0;
      const noShowRate = sums.trial_book > 0 ? ((1 - sums.trial_done_n1 / sums.trial_book) * 100) : 0;

      // Render MTD content
      const content = document.getElementById('daily-mtd-content');
      if (!content) return;

      // Section: N1
      let html = '<div class="daily-mtd-group"><div class="daily-mtd-group-title">N1 — OPTIMIZE</div>';
      ['lead_mkt','calls_n1','trial_book','trial_done_n1','deal_n1','revenue_n1'].forEach(id => {
        const f = DAILY_FIELDS.find(x => x.id === id);
        const val = id === 'revenue_n1' ? `${sums[id]}M` : sums[id];
        html += `<div class="daily-mtd-item"><span class="daily-mtd-item-label">${escHtml(f.label)}</span><span class="daily-mtd-item-value">${val}</span></div>`;
      });
      html += '</div>';

      // Section: N2
      html += '<div class="daily-mtd-group"><div class="daily-mtd-group-title">N2 — OPS</div>';
      ['cs_touchpoint','phhs_reupsell','deal_reupsell','deal_referral','revenue_n2'].forEach(id => {
        const f = DAILY_FIELDS.find(x => x.id === id);
        const val = id === 'revenue_n2' ? `${sums[id]}M` : sums[id];
        html += `<div class="daily-mtd-item"><span class="daily-mtd-item-label">${escHtml(f.label)}</span><span class="daily-mtd-item-value">${val}</span></div>`;
      });
      html += '</div>';

      // Section: N3
      html += '<div class="daily-mtd-group"><div class="daily-mtd-group-title">N3 — GROWTH</div>';
      ['direct_sales','events','lead_n3','trial_n3','deal_n3','revenue_n3'].forEach(id => {
        const f = DAILY_FIELDS.find(x => x.id === id);
        const val = id === 'revenue_n3' ? `${sums[id]}M` : sums[id];
        html += `<div class="daily-mtd-item"><span class="daily-mtd-item-label">${escHtml(f.label)}</span><span class="daily-mtd-item-value">${val}</span></div>`;
      });
      html += '</div>';

      // Totals
      html += '<div class="daily-mtd-group daily-mtd-group-total"><div class="daily-mtd-group-title">TỔNG HỢP MTD</div>';
      html += `<div class="daily-mtd-item daily-mtd-item-total"><span class="daily-mtd-item-label">Tổng Deal BU</span><span class="daily-mtd-item-value">${totalDeal}</span></div>`;
      html += `<div class="daily-mtd-item daily-mtd-item-total"><span class="daily-mtd-item-label">Tổng Doanh số</span><span class="daily-mtd-item-value">${totalRev}M</span></div>`;
      html += '</div>';

      // Auto-calculated KPIs
      html += '<div class="daily-mtd-group daily-mtd-group-kpi"><div class="daily-mtd-group-title">CHỈ SỐ TỰ TÍNH</div>';
      html += `<div class="daily-mtd-item"><span class="daily-mtd-item-label">CR16 (Deal N1 / Lead MKT)</span><span class="daily-mtd-item-value">${cr16.toFixed(1)}%</span></div>`;
      html += `<div class="daily-mtd-item"><span class="daily-mtd-item-label">CR46 (Deal N1 / Trial N1)</span><span class="daily-mtd-item-value">${cr46.toFixed(1)}%</span></div>`;
      html += `<div class="daily-mtd-item"><span class="daily-mtd-item-label">AOV (DS N1 / Deal N1)</span><span class="daily-mtd-item-value">${aov.toFixed(1)}M</span></div>`;
      html += `<div class="daily-mtd-item"><span class="daily-mtd-item-label">No-show Rate</span><span class="daily-mtd-item-value">${noShowRate.toFixed(1)}%</span></div>`;
      html += '</div>';

      content.innerHTML = html;
    } else {
      panel.style.display = 'none';
      state.daily.mtd = null;
    }
  } catch (err) {
    console.error('loadDailyMTD error:', err);
  }
}

// ============================================================
// DAILY OPERATIONS DASHBOARD
// ============================================================

state.dailyOps = {
  month: '',
  rawData: null,  // all daily rows for the month
  sums: null      // computed aggregates
};

function switchDashTab(tab) {
  document.querySelectorAll('.dash-tab').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.dash-tab-content').forEach(c => c.classList.remove('active'));
  const btn = document.querySelector(`.dash-tab[data-dtab="${tab}"]`);
  if (btn) btn.classList.add('active');
  const content = document.getElementById(`dtab-${tab}`);
  if (content) content.classList.add('active');

  if (tab === 'daily-ops') {
    initDailyOps();
  }
}

function initDailyOps() {
  // Populate month dropdown
  const monthSel = document.getElementById('dops-month');
  if (monthSel && monthSel.options.length <= 1) {
    monthSel.innerHTML = '';
    MONTHS.forEach(m => {
      const opt = document.createElement('option');
      opt.value = m;
      opt.textContent = m;
      monthSel.appendChild(opt);
    });
    // Default to current month
    const now = new Date();
    const curMonth = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`;
    monthSel.value = curMonth;
  }

  // Populate BU dropdown
  const buSel = document.getElementById('dops-bu');
  if (buSel && buSel.options.length <= 1) {
    buSel.innerHTML = '<option value="ALL">T\u1EA5t c\u1EA3 BU</option>';
    BU_LIST.forEach(bu => {
      const opt = document.createElement('option');
      opt.value = bu;
      opt.textContent = bu;
      buSel.appendChild(opt);
    });
  }

  // Auto-load if no data yet
  if (!state.dailyOps.rawData) {
    loadDailyOps();
  }
}

async function loadDailyOps() {
  const month = document.getElementById('dops-month')?.value;
  if (!month) return;

  state.dailyOps.month = month;

  const loading = document.getElementById('dops-loading');
  const empty = document.getElementById('dops-empty');
  const content = document.getElementById('dops-content');

  if (loading) loading.style.display = '';
  if (empty) empty.style.display = 'none';
  if (content) content.style.display = 'none';

  try {
    // Try getDailyAll first (requires updated Apps Script)
    let allData = null;
    try {
      const result = await apiFetch('getDailyAll', { month });
      if (result.ok && result.data && result.data.length > 0) {
        allData = result.data;
      }
    } catch (e) {
      console.log('getDailyAll not available, falling back to per-BU fetch...');
    }

    // Fallback: fetch each BU's MTD data via getDailyMTD
    if (!allData) {
      allData = [];
      const fetchPromises = BU_LIST.map(async (bu) => {
        try {
          const r = await apiFetch('getDailyMTD', { bu, month });
          if (r.ok && r.data) return r.data;
        } catch (e) { /* skip */ }
        return [];
      });
      const results = await Promise.all(fetchPromises);
      results.forEach(rows => { allData = allData.concat(rows); });
    }

    if (loading) loading.style.display = 'none';

    if (allData.length > 0) {
      state.dailyOps.rawData = allData;
      if (content) content.style.display = '';
      renderDailyOps();
    } else {
      state.dailyOps.rawData = null;
      if (empty) {
        empty.style.display = '';
        empty.textContent = `Ch\u01B0a c\u00F3 d\u1EEF li\u1EC7u Daily Report cho th\u00E1ng ${month}`;
      }
    }
  } catch (err) {
    console.error('loadDailyOps error:', err);
    if (loading) loading.style.display = 'none';
    if (empty) {
      empty.style.display = '';
      empty.textContent = 'L\u1ED7i khi t\u1EA3i d\u1EEF li\u1EC7u. Vui l\u00F2ng th\u1EED l\u1EA1i.';
    }
  }
}

function renderDailyOps() {
  const data = state.dailyOps.rawData;
  if (!data) return;

  const sourceFilter = document.getElementById('dops-source')?.value || 'ALL';
  const buFilter = document.getElementById('dops-bu')?.value || 'ALL';

  // Filter by BU
  const filtered = buFilter === 'ALL' ? data : data.filter(r => r.bu === buFilter);

  // Compute sums
  const s = {};
  DAILY_FIELDS.forEach(f => { if (f.type === 'number') s[f.id] = 0; });
  filtered.forEach(row => {
    DAILY_FIELDS.forEach(f => {
      if (f.type === 'number') s[f.id] += parseInt(row[f.id]) || 0;
    });
  });

  // Totals (raw sums — may contain data entry outliers)
  const totalRev = (s.revenue_n1||0) + (s.revenue_n2||0) + (s.revenue_n3||0);
  const noShow = s.trial_book > 0 ? ((1 - s.trial_done_n1 / s.trial_book) * 100) : 0;

  // --- Outlier filtering: remove rows where deal count > 1000 (likely revenue entered in deal field) ---
  const cleanFiltered = filtered.map(row => {
    const r = Object.assign({}, row);
    if ((parseInt(r.deal_n1)||0) > 1000) { r.deal_n1 = 0; }
    if ((parseInt(r.deal_reupsell)||0) > 1000) { r.deal_reupsell = 0; }
    if ((parseInt(r.deal_referral)||0) > 1000) { r.deal_referral = 0; }
    if ((parseInt(r.deal_n3)||0) > 1000) { r.deal_n3 = 0; }
    return r;
  });
  // Recompute clean deal sums for CR calculations
  let cleanDealN1 = 0, cleanDealReupsell = 0, cleanDealReferral = 0, cleanDealN3 = 0;
  cleanFiltered.forEach(row => {
    cleanDealN1 += parseInt(row.deal_n1) || 0;
    cleanDealReupsell += parseInt(row.deal_reupsell) || 0;
    cleanDealReferral += parseInt(row.deal_referral) || 0;
    cleanDealN3 += parseInt(row.deal_n3) || 0;
  });
  const cleanDealN2 = cleanDealReupsell + cleanDealReferral;
  const cleanDealTotal = cleanDealN1 + cleanDealN2 + cleanDealN3;

  // Per-source CRs (using clean deal counts)
  // N1: CR16 = Deal N1 / Lead MKT, CR46 = Deal N1 / Trial Done N1, AOV N1
  const cr16_n1 = s.lead_mkt > 0 ? (cleanDealN1 / s.lead_mkt * 100) : 0;
  const cr46_n1 = s.trial_done_n1 > 0 ? (cleanDealN1 / s.trial_done_n1 * 100) : 0;
  const aov_n1 = cleanDealN1 > 0 ? (s.revenue_n1 / cleanDealN1) : 0;
  const noShow_n1 = s.trial_book > 0 ? ((1 - s.trial_done_n1 / s.trial_book) * 100) : 0;

  // N2: CR tư vấn = Deal ReUpsell / PHHS tư vấn, CR Referral = Deal Ref / CS touchpoint
  const crReupsell = s.phhs_reupsell > 0 ? (cleanDealReupsell / s.phhs_reupsell * 100) : 0;
  const crReferral = s.cs_touchpoint > 0 ? (cleanDealReferral / s.cs_touchpoint * 100) : 0;
  const aov_n2 = cleanDealN2 > 0 ? (s.revenue_n2 / cleanDealN2) : 0;

  // N3: CR16 = Deal N3 / Lead N3, CR46 = Deal N3 / Trial N3, AOV N3
  const cr16_n3 = s.lead_n3 > 0 ? (cleanDealN3 / s.lead_n3 * 100) : 0;
  const cr46_n3 = s.trial_n3 > 0 ? (cleanDealN3 / s.trial_n3 * 100) : 0;
  const aov_n3 = cleanDealN3 > 0 ? (s.revenue_n3 / cleanDealN3) : 0;

  // Unique BUs and days
  const uniqueBUs = new Set(filtered.map(r => r.bu));
  const uniqueDays = new Set(filtered.map(r => r.date));

  // Update meta
  const metaEl = document.getElementById('dops-meta');
  if (metaEl) {
    const label = buFilter === 'ALL' ? `${uniqueBUs.size} BU` : buFilter;
    metaEl.textContent = `${label} \u2014 ${uniqueDays.size} ng\u00E0y \u2014 Th\u00E1ng ${state.dailyOps.month}`;
  }

  // Blended KPIs (using clean deal counts)
  const totalLeadAll = (s.lead_mkt||0) + (s.lead_n3||0);
  const totalTrialAll = (s.trial_done_n1||0) + (s.trial_n3||0);
  const cleanDealBlend = cleanDealN1 + cleanDealN3;
  const cr16 = totalLeadAll > 0 ? (cleanDealBlend / totalLeadAll * 100) : 0;
  const cr46 = totalTrialAll > 0 ? (cleanDealBlend / totalTrialAll * 100) : 0;
  const aovAll = cleanDealTotal > 0 ? (totalRev / cleanDealTotal) : 0;

  // KPI Cards — formatted
  setText('dops-total-deal', fmtDops(cleanDealTotal, 'count'));
  setText('dops-total-revenue', fmtDops(totalRev, 'money'));

  const aovM = aovAll / 1e6; // VNĐ → triệu for benchmark comparison
  const cr16El = document.getElementById('dops-cr16');
  if (cr16El) {
    cr16El.textContent = fmtDops(cr16, 'pct');
    cr16El.className = 'dops-kpi-value ' + (cr16 >= 15 ? 'dops-kpi-good' : cr16 >= 10 ? 'dops-kpi-warn' : 'dops-kpi-bad');
  }
  const cr46El = document.getElementById('dops-cr46');
  if (cr46El) {
    cr46El.textContent = fmtDops(cr46, 'pct');
    cr46El.className = 'dops-kpi-value ' + (cr46 >= 50 ? 'dops-kpi-good' : cr46 >= 35 ? 'dops-kpi-warn' : 'dops-kpi-bad');
  }
  const aovEl = document.getElementById('dops-aov');
  if (aovEl) {
    aovEl.textContent = fmtDops(aovAll, 'aov');
    aovEl.className = 'dops-kpi-value ' + (aovM >= 18 ? 'dops-kpi-good' : aovM >= 15 ? 'dops-kpi-warn' : 'dops-kpi-bad');
  }
  setText('dops-noshow', fmtDops(noShow, 'pct'));

  // Render N1 metrics
  renderDopsSection('dops-metrics-n1', [
    { label: 'Lead MKT nh\u1EADn', value: fmtDops(s.lead_mkt, 'count') },
    { label: 'Cu\u1ED9c g\u1ECDi N1', value: fmtDops(s.calls_n1, 'count') },
    { label: 'Trial Book', value: fmtDops(s.trial_book, 'count') },
    { label: 'Trial Done', value: fmtDops(s.trial_done_n1, 'count') },
    { label: 'Deal N1', value: fmtDops(cleanDealN1, 'count') },
    { label: 'Doanh s\u1ED1 N1', value: fmtDops(s.revenue_n1, 'money') }
  ]);
  // N1 CR strip
  renderCRStrip('dops-cr-n1', [
    { label: 'CR16 (N1)', value: cr16_n1, bench: 15, benchLabel: 'BM: 15%' },
    { label: 'CR46 (N1)', value: cr46_n1, bench: 50, benchLabel: 'BM: 50%' },
    { label: 'No-show', value: noShow_n1, invert: true, bench: 20, benchLabel: 'BM: <20%' },
    { label: 'AOV (N1)', value: aov_n1, suffix: 'AOV', bench: 18, benchLabel: 'BM: 18M' }
  ]);

  // Render N2 metrics
  renderDopsSection('dops-metrics-n2', [
    { label: 'PHHS ch\u0103m s\u00F3c', value: fmtDops(s.cs_touchpoint, 'count') },
    { label: 'PHHS t\u01B0 v\u1EA5n Re/Upsell', value: fmtDops(s.phhs_reupsell, 'count') },
    { label: 'Deal Re/Upsell', value: fmtDops(cleanDealReupsell, 'count') },
    { label: 'Deal Referral', value: fmtDops(cleanDealReferral, 'count') },
    { label: 'Doanh s\u1ED1 N2', value: fmtDops(s.revenue_n2, 'money') }
  ]);
  // N2 CR strip
  renderCRStrip('dops-cr-n2', [
    { label: 'CR Re/Upsell', value: crReupsell, bench: null },
    { label: 'CR Referral', value: crReferral, bench: null },
    { label: 'AOV (N2)', value: aov_n2, suffix: 'AOV', bench: null }
  ]);

  // Render N3 metrics
  renderDopsSection('dops-metrics-n3', [
    { label: 'Direct Sales', value: fmtDops(s.direct_sales, 'count') },
    { label: 'Event', value: fmtDops(s.events, 'count') },
    { label: 'Lead N3', value: fmtDops(s.lead_n3, 'count') },
    { label: 'Trial N3', value: fmtDops(s.trial_n3, 'count') },
    { label: 'Deal N3', value: fmtDops(cleanDealN3, 'count') },
    { label: 'Doanh s\u1ED1 N3', value: fmtDops(s.revenue_n3, 'money') }
  ]);
  // N3 CR strip
  renderCRStrip('dops-cr-n3', [
    { label: 'CR16 (N3)', value: cr16_n3, bench: 15, benchLabel: 'BM: 15%' },
    { label: 'CR46 (N3)', value: cr46_n3, bench: 50, benchLabel: 'BM: 50%' },
    { label: 'AOV (N3)', value: aov_n3, suffix: 'AOV', bench: 18, benchLabel: 'BM: 18M' }
  ]);

  // Source filter visibility
  document.getElementById('dops-section-n1').style.display = (sourceFilter === 'ALL' || sourceFilter === 'N1') ? '' : 'none';
  document.getElementById('dops-section-n2').style.display = (sourceFilter === 'ALL' || sourceFilter === 'N2') ? '' : 'none';
  document.getElementById('dops-section-n3').style.display = (sourceFilter === 'ALL' || sourceFilter === 'N3') ? '' : 'none';

  // Coverage
  const coverageEl = document.getElementById('dops-coverage-text');
  if (coverageEl) {
    coverageEl.textContent = `${uniqueBUs.size} BU \u0111\u00E3 b\u00E1o c\u00E1o / ${BU_LIST.length} BU t\u1ED5ng \u2014 ${filtered.length} b\u1EA3n ghi t\u1EEB ${uniqueDays.size} ng\u00E0y`;
  }
}

function renderDopsSection(containerId, metrics) {
  const el = document.getElementById(containerId);
  if (!el) return;
  // Dynamically set grid columns based on metric count
  el.className = 'dops-metrics' + (metrics.length === 5 ? ' dops-cols-5' : '');
  el.innerHTML = metrics.map(m =>
    `<div class="dops-metric">
      <div class="dops-metric-label">${m.label}</div>
      <div class="dops-metric-value">${m.value}</div>
    </div>`
  ).join('');
}

function renderCRStrip(containerId, chips) {
  const el = document.getElementById(containerId);
  if (!el) return;
  el.innerHTML = chips.map(c => {
    const val = c.value || 0;
    const isAOV = c.suffix === 'AOV';
    // AOV value is in VNĐ → convert to triệu for display & benchmark
    const aovM = isAOV ? val / 1e6 : 0;
    const display = isAOV ? fmtDops(val, 'aov') : fmtDops(val, 'pct');
    let colorClass = 'dops-cr-neutral';
    if (c.bench !== null && c.bench !== undefined) {
      if (c.invert) {
        colorClass = val <= c.bench ? 'dops-cr-good' : val <= c.bench * 1.5 ? 'dops-cr-warn' : 'dops-cr-bad';
      } else if (isAOV) {
        colorClass = aovM >= c.bench ? 'dops-cr-good' : aovM >= c.bench * 0.8 ? 'dops-cr-warn' : 'dops-cr-bad';
      } else {
        colorClass = val >= c.bench ? 'dops-cr-good' : val >= c.bench * 0.7 ? 'dops-cr-warn' : 'dops-cr-bad';
      }
    }
    const benchHtml = c.benchLabel ? `<div class="dops-cr-chip-bench">${c.benchLabel}</div>` : '';
    return `<div class="dops-cr-chip">
      <div class="dops-cr-chip-label">${c.label}</div>
      <div class="dops-cr-chip-value ${colorClass}">${display}</div>
      ${benchHtml}
    </div>`;
  }).join('');
}

function setText(id, val) {
  const el = document.getElementById(id);
  if (el) el.textContent = val;
}

/** Smart number formatter for Daily Operations
 *  - 'count': integer with thousand separators (1,234)
 *  - 'money': VNĐ → triệu/tỷ with suffix M/B (52.2M, 1.7B)
 *  - 'pct': percentage with 1 decimal (15.3%)
 *  - 'aov': VNĐ per deal → triệu with 1 decimal (18.5M)
 */
function fmtDops(value, type) {
  const v = Number(value) || 0;
  if (type === 'count') {
    // Smart abbreviation for large counts
    if (Math.abs(v) >= 1e9) return (v / 1e9).toFixed(1).replace(/\.0$/, '') + 'B';
    if (Math.abs(v) >= 1e6) return (v / 1e6).toFixed(1).replace(/\.0$/, '') + 'M';
    if (Math.abs(v) >= 10000) return (v / 1e3).toFixed(1).replace(/\.0$/, '') + 'K';
    return v.toLocaleString('en-US');
  }
  if (type === 'money') {
    const m = v / 1e6; // VNĐ → triệu
    if (Math.abs(m) >= 1000) {
      return (m / 1000).toFixed(1).replace(/\.0$/, '') + 'B';
    }
    if (Math.abs(m) >= 100) {
      return Math.round(m).toLocaleString('en-US') + 'M';
    }
    if (Math.abs(m) >= 10) {
      return m.toFixed(1).replace(/\.0$/, '') + 'M';
    }
    return m.toFixed(1) + 'M';
  }
  if (type === 'pct') {
    return v.toFixed(1) + '%';
  }
  if (type === 'aov') {
    const m = v / 1e6;
    return m.toFixed(1) + 'M';
  }
  return String(v);
}

// ============================================================
// CM ANALYTICAL TAB — Single-BU version of FM Analytics
// ============================================================

state.cmAnalytics = {
  month: '',
  rawData: null,  // all BU data (for top-BU comparison)
  buData: null    // filtered to selected BU
};

function initCMAnalytics() {
  const bu = state.cm.bu;
  const noBuEl = document.getElementById('cm-ana-no-bu');
  const wrapperEl = document.getElementById('cm-ana-wrapper');
  if (!bu) {
    if (noBuEl) noBuEl.style.display = '';
    if (wrapperEl) wrapperEl.style.display = 'none';
    return;
  }
  if (noBuEl) noBuEl.style.display = 'none';
  if (wrapperEl) wrapperEl.style.display = '';

  // Populate month dropdown
  const monthSel = document.getElementById('cm-ana-month');
  if (monthSel && monthSel.options.length <= 1) {
    monthSel.innerHTML = '';
    MONTHS.forEach(m => {
      const opt = document.createElement('option');
      opt.value = m; opt.textContent = m;
      monthSel.appendChild(opt);
    });
    const now = new Date();
    monthSel.value = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`;
  }

  loadCMAnalytics();
}

async function loadCMAnalytics() {
  const bu = state.cm.bu;
  const month = document.getElementById('cm-ana-month')?.value;
  if (!bu || !month) return;
  state.cmAnalytics.month = month;

  const loading = document.getElementById('cm-ana-loading');
  const empty = document.getElementById('cm-ana-empty');
  const content = document.getElementById('cm-ana-content');
  if (loading) loading.style.display = '';
  if (empty) empty.style.display = 'none';
  if (content) content.style.display = 'none';

  try {
    // Try to get all data for top-BU comparison
    let allData = null;
    if (state.dailyOps && state.dailyOps.rawData && state.dailyOps.month === month) {
      allData = state.dailyOps.rawData;
    } else if (state.analytics && state.analytics.rawData && state.analytics.month === month) {
      allData = state.analytics.rawData;
    } else {
      try {
        const r = await apiFetch('getDailyAll', { month });
        if (r.ok && r.data) allData = r.data;
      } catch(e) {}
      // Fallback: just get this BU
      if (!allData) {
        try {
          const r = await apiFetch('getDailyMTD', { bu, month });
          if (r.ok && r.data) allData = r.data;
        } catch(e) {}
      }
    }

    if (loading) loading.style.display = 'none';
    if (!allData || allData.length === 0) {
      state.cmAnalytics.rawData = null;
      state.cmAnalytics.buData = null;
      if (empty) { empty.style.display = ''; empty.textContent = `Chưa có dữ liệu Daily Report cho tháng ${month}`; }
      return;
    }

    state.cmAnalytics.rawData = allData;
    const buRows = allData.filter(r => r.bu === bu);
    if (buRows.length === 0) {
      state.cmAnalytics.buData = null;
      if (empty) { empty.style.display = ''; empty.textContent = `Chưa có dữ liệu cho ${bu} trong tháng ${month}`; }
      return;
    }
    state.cmAnalytics.buData = buRows;
    if (content) content.style.display = '';
    renderCMAnalytics();
  } catch(err) {
    console.error('loadCMAnalytics error:', err);
    if (loading) loading.style.display = 'none';
    if (empty) { empty.style.display = ''; empty.textContent = 'Lỗi khi tải dữ liệu.'; }
  }
}

function renderCMAnalytics() {
  const bu = state.cm.bu;
  const filtered = state.cmAnalytics.buData;
  const allData = state.cmAnalytics.rawData || filtered;
  if (!filtered || !bu) return;

  const month = state.cmAnalytics.month;
  const metaEl = document.getElementById('cm-ana-meta');
  if (metaEl) metaEl.textContent = `${bu} — Tháng ${month}`;

  const bm = state.config.data?.benchmarks || CONFIG_DEFAULTS.benchmarks;
  const daysInMonth = new Date(parseInt(month.split('-')[0]), parseInt(month.split('-')[1]), 0).getDate();
  const sortedDays = [...new Set(filtered.map(r => r.date))].sort();
  const daysPassed = sortedDays.length > 0 ? Math.max(...sortedDays.map(d => parseInt(d.split('-')[2]))) : 0;
  const daysRemaining = daysInMonth - daysPassed;
  const avgDays = daysPassed > 0 ? daysPassed : 1;

  // Single BU = buCount is always 1
  const buCount = 1;

  // N1
  const n1_leads = _sum(filtered, 'lead_mkt');
  const n1_calls = _sum(filtered, 'calls_n1');
  const n1_trial_book = _sum(filtered, 'trial_book');
  const n1_trial_done = _sum(filtered, 'trial_done_n1');
  const n1_deals = _sum(filtered, 'deal_n1');
  const n1_rev = _sum(filtered, 'revenue_n1');
  const n1_calls_avg = n1_calls / avgDays;
  const n1_trial_book_avg = n1_trial_book / avgDays;

  // N2
  const n2_cs = _sum(filtered, 'cs_touchpoint');
  const n2_phhs = _sum(filtered, 'phhs_reupsell');
  const n2_deal_re = _sum(filtered, 'deal_reupsell');
  const n2_deal_ref = _sum(filtered, 'deal_referral');
  const n2_rev = _sum(filtered, 'revenue_n2');
  const n2_cs_avg = n2_cs / avgDays;
  const n2_phhs_avg = n2_phhs / avgDays;

  // N3
  const n3_direct = _sum(filtered, 'direct_sales');
  const n3_events = _sum(filtered, 'events');
  const n3_leads = _sum(filtered, 'lead_n3');
  const n3_trial = _sum(filtered, 'trial_n3');
  const n3_deals = _sum(filtered, 'deal_n3');
  const n3_rev = _sum(filtered, 'revenue_n3');
  const n3_direct_avg = n3_direct / avgDays;
  const n3_events_avg = n3_events / avgDays;

  const WM_BENCH = { calls: 15, trial_book: 3, cs_touch: 8, phhs: 3, direct: 2, events: 0.1 };

  function cmWmRow(label, total, avgPerDay, bench, unit) {
    const pct = bench > 0 ? Math.min(avgPerDay / bench * 100, 150) : 0;
    const b = _badge(avgPerDay, bench, bench * 0.6);
    const color = _barColor(avgPerDay, bench, bench * 0.6);
    return `<div class="ana-wm-row">
      <span class="ana-wm-label">${label}</span>
      <span class="ana-wm-value">${total}</span>
      <div class="ana-wm-bar-wrap"><div class="ana-wm-bar" style="width:${Math.min(pct,100)}%;background:${color}"></div></div>
      <span class="ana-wm-badge ${b.cls}">${avgPerDay.toFixed(1)}${unit}/ngày</span>
      <span class="ana-wm-benchmark">(B: ${bench})</span>
    </div>`;
  }

  // SECTION 1: WORK METRICS
  const wmEl = document.getElementById('cm-ana-work-metrics');
  if (wmEl) {
    wmEl.innerHTML = `<div class="ana-wm-grid">
      <div class="ana-wm-source">
        <div class="ana-wm-source-title n1">N1 — Optimize (MKT)</div>
        ${cmWmRow('Calls (gọi ra)', n1_calls, n1_calls_avg, WM_BENCH.calls, '')}
        ${cmWmRow('Trial Book (đặt hẹn)', n1_trial_book, n1_trial_book_avg, WM_BENCH.trial_book, '')}
        <div class="ana-wm-row">
          <span class="ana-wm-label">Lead MKT (nhận)</span>
          <span class="ana-wm-value">${n1_leads}</span>
          <span class="ana-wm-badge" style="background:var(--blue-bg);color:#1a5aa0">${(n1_leads / avgDays).toFixed(1)}/ngày</span>
        </div>
      </div>
      <div class="ana-wm-source">
        <div class="ana-wm-source-title n2">N2 — Ops (Re/Upsell)</div>
        ${cmWmRow('CS Touchpoint', n2_cs, n2_cs_avg, WM_BENCH.cs_touch, '')}
        ${cmWmRow('PHHS tư vấn', n2_phhs, n2_phhs_avg, WM_BENCH.phhs, '')}
        <div class="ana-wm-row">
          <span class="ana-wm-label">Deal ReUpsell</span>
          <span class="ana-wm-value">${n2_deal_re}</span>
          <span class="ana-wm-badge" style="background:var(--blue-bg);color:#1a5aa0">${(n2_deal_re / avgDays).toFixed(2)}/ngày</span>
        </div>
        <div class="ana-wm-row">
          <span class="ana-wm-label">Deal Referral</span>
          <span class="ana-wm-value">${n2_deal_ref}</span>
          <span class="ana-wm-badge" style="background:var(--blue-bg);color:#1a5aa0">${(n2_deal_ref / avgDays).toFixed(2)}/ngày</span>
        </div>
      </div>
      <div class="ana-wm-source">
        <div class="ana-wm-source-title n3">N3 — Growth (Tự kiếm)</div>
        ${cmWmRow('Direct Sales', n3_direct, n3_direct_avg, WM_BENCH.direct, '')}
        ${cmWmRow('Events', n3_events, n3_events_avg, WM_BENCH.events, '')}
        <div class="ana-wm-row">
          <span class="ana-wm-label">Lead N3</span>
          <span class="ana-wm-value">${n3_leads}</span>
          <span class="ana-wm-badge" style="background:var(--blue-bg);color:#1a5aa0">${(n3_leads / avgDays).toFixed(1)}/ngày</span>
        </div>
      </div>
    </div>`;
  }

  // SECTION 2: CONVERSION
  const cr16_n1 = _pct(n1_deals, n1_leads);
  const cr46_n1 = _pct(n1_deals, n1_trial_done);
  const noshow = _pct(n1_trial_book - n1_trial_done, n1_trial_book);
  const cr_re = _pct(n2_deal_re, n2_phhs);
  const cr_ref = _pct(n2_deal_ref, n2_cs);
  const cr16_n3 = _pct(n3_deals, n3_leads);
  const cr46_n3 = _pct(n3_deals, n3_trial);

  function cmCrBadge(val, goodT, warnT, inv) {
    if (inv) return val <= goodT ? 'good' : val <= warnT ? 'warn' : 'bad';
    return val >= goodT ? 'good' : val >= warnT ? 'warn' : 'bad';
  }
  function cmFunnelStep(num, label) {
    return `<div class="ana-funnel-step"><span class="ana-funnel-num">${num}</span><span class="ana-funnel-label">${label}</span></div>`;
  }
  function cmFunnelArrow(crVal, crLabel, goodT, warnT, inv) {
    const cls = cmCrBadge(crVal, goodT, warnT, inv);
    return `<div class="ana-funnel-step"><span class="ana-funnel-arrow">↓</span><span class="ana-funnel-cr ${cls}">${crLabel}: ${crVal.toFixed(1)}% (B: ${goodT}%)</span></div>`;
  }

  const convEl = document.getElementById('cm-ana-conversion');
  if (convEl) {
    convEl.innerHTML = `<div class="ana-conv-grid">
      <div class="ana-conv-source">
        <div class="ana-conv-title n1">N1 — Funnel MKT</div>
        ${cmFunnelStep(n1_leads, 'Lead MKT')}
        ${cmFunnelArrow(cr16_n1, 'CR16', parseFloat(bm.cr16_n1)||15, 10, false)}
        ${cmFunnelStep(n1_trial_book, 'Trial Book')}
        ${cmFunnelArrow(noshow, 'No-show', parseFloat(bm.noshow)||20, 30, true)}
        ${cmFunnelStep(n1_trial_done, 'Trial Done')}
        ${cmFunnelArrow(cr46_n1, 'CR46', parseFloat(bm.cr46_n1)||50, 35, false)}
        ${cmFunnelStep(n1_deals, 'Deal N1')}
        <div class="ana-funnel-step" style="margin-top:6px;border-top:1px solid var(--gray-200);padding-top:8px">
          <span class="ana-funnel-num" style="color:var(--red)">${n1_rev}M</span>
          <span class="ana-funnel-label">Doanh số N1</span>
        </div>
      </div>
      <div class="ana-conv-source">
        <div class="ana-conv-title n2">N2 — Funnel Ops</div>
        ${cmFunnelStep(n2_cs, 'CS Touchpoint')}
        ${cmFunnelArrow(cr_re, 'CR ReUpsell', parseFloat(bm.cr_reupsell)||15, 8, false)}
        ${cmFunnelStep(n2_deal_re, 'Deal ReUpsell')}
        <div style="height:12px"></div>
        ${cmFunnelStep(n2_cs, 'CS Touchpoint')}
        ${cmFunnelArrow(cr_ref, 'CR Referral', parseFloat(bm.cr_referral)||5, 2, false)}
        ${cmFunnelStep(n2_deal_ref, 'Deal Referral')}
        <div class="ana-funnel-step" style="margin-top:6px;border-top:1px solid var(--gray-200);padding-top:8px">
          <span class="ana-funnel-num" style="color:#1a5aa0">${n2_rev}M</span>
          <span class="ana-funnel-label">Doanh số N2</span>
        </div>
      </div>
      <div class="ana-conv-source">
        <div class="ana-conv-title n3">N3 — Funnel Growth</div>
        ${cmFunnelStep(n3_leads, 'Lead N3')}
        ${cmFunnelArrow(cr16_n3, 'CR16', parseFloat(bm.cr16_n3)||10, 5, false)}
        ${cmFunnelStep(n3_trial, 'Trial N3')}
        ${cmFunnelArrow(cr46_n3, 'CR46', parseFloat(bm.cr46_n3)||40, 25, false)}
        ${cmFunnelStep(n3_deals, 'Deal N3')}
        <div class="ana-funnel-step" style="margin-top:6px;border-top:1px solid var(--gray-200);padding-top:8px">
          <span class="ana-funnel-num" style="color:var(--green)">${n3_rev}M</span>
          <span class="ana-funnel-label">Doanh số N3</span>
        </div>
      </div>
    </div>`;
  }

  // SECTION 3: DIAGNOSTIC (single BU)
  const totalDeals = n1_deals + n2_deal_re + n2_deal_ref + n3_deals;
  const totalRev = n1_rev + n2_rev + n3_rev;
  const workScore = (
    n1_calls_avg / WM_BENCH.calls +
    n1_trial_book_avg / WM_BENCH.trial_book +
    n2_cs_avg / WM_BENCH.cs_touch +
    n3_direct_avg / WM_BENCH.direct
  ) / 4 * 100;
  const dealPerDay = totalDeals / avgDays;
  const isHW = workScore >= 70;
  const isGR = dealPerDay >= 0.8 && cr16_n1 >= 12;

  let diagType, diagTitle, diagIcon, diagDesc;
  if (isHW && isGR) {
    diagType = 'star'; diagIcon = '⭐'; diagTitle = 'Chăm chỉ + Kết quả tốt';
    diagDesc = 'BU đang hoạt động tốt. Tiếp tục duy trì và chia sẻ kinh nghiệm cho các BU khác.';
  } else if (isHW && !isGR) {
    diagType = 'smart'; diagIcon = '🧠'; diagTitle = 'Chăm chỉ nhưng KQ thấp';
    diagDesc = 'BU đang làm việc nhiều nhưng chưa hiệu quả. Cần làm việc thông minh hơn — hỏi MiMi để được gợi ý chiến lược.';
  } else {
    diagType = 'effort'; diagIcon = '💪'; diagTitle = 'Chưa đủ nỗ lực';
    diagDesc = 'Các chỉ số công việc còn thấp. Cần tăng cường nỗ lực làm việc trước khi tối ưu chiến lược.';
  }

  const diagEl = document.getElementById('cm-ana-diagnostic');
  if (diagEl) {
    diagEl.innerHTML = `<div class="ana-diag-cards" style="grid-template-columns:1fr">
      <div class="ana-diag-card ${diagType}">
        <div class="ana-diag-card-title">${diagIcon} ${diagTitle}</div>
        <div class="ana-diag-card-desc">${diagDesc}</div>
        <div class="ana-diag-items">
          <strong>${bu}</strong>: Work Score ${workScore.toFixed(0)}% | ${dealPerDay.toFixed(1)} deal/ngày | CR16 N1=${cr16_n1.toFixed(1)}%
        </div>
      </div>
    </div>`;
  }

  // SECTION 4: EOM FORECAST
  const dailyDealRate = daysPassed > 0 ? totalDeals / daysPassed : 0;
  const dailyRevRate = daysPassed > 0 ? totalRev / daysPassed : 0;
  const forecastDeal = Math.round(dailyDealRate * daysInMonth);
  const forecastRev = Math.round(dailyRevRate * daysInMonth);
  const callGap = Math.max(0, WM_BENCH.calls - n1_calls_avg);
  const potentialExtraTrials = callGap * daysRemaining * 0.2;
  const potentialExtraDeals = Math.round(potentialExtraTrials * (cr46_n1 / 100) || 0);
  const pace = state.config.data?.week_pace || CONFIG_DEFAULTS.week_pace;
  const day = new Date().getDate();
  const wk = day <= 7 ? 'w1' : day <= 14 ? 'w2' : day <= 21 ? 'w3' : 'w4';
  const currentPace = pace[wk] || 0.4;
  const progressPct = daysPassed > 0 ? (daysPassed / daysInMonth * 100) : 0;

  const actions = [];
  if (n1_calls_avg < WM_BENCH.calls * 0.6)
    actions.push({ type: 'alert', icon: '📞', text: `<strong>Calls quá thấp (${n1_calls_avg.toFixed(1)}/ngày)</strong> — Benchmark ${WM_BENCH.calls}/ngày. Tăng calls có thể thêm ~${potentialExtraDeals} deal cuối tháng.` });
  else if (n1_calls_avg < WM_BENCH.calls)
    actions.push({ type: 'improve', icon: '📞', text: `<strong>Calls cần cải thiện (${n1_calls_avg.toFixed(1)}/${WM_BENCH.calls}/ngày)</strong> — Tăng thêm ${callGap.toFixed(0)} calls/ngày để đạt benchmark.` });
  else
    actions.push({ type: 'good', icon: '✅', text: `<strong>Calls đạt chuẩn (${n1_calls_avg.toFixed(1)}/ngày)</strong> — Tiếp tục duy trì.` });

  if (cr46_n1 < (parseFloat(bm.cr46_n1) || 50) * 0.7)
    actions.push({ type: 'alert', icon: '🎯', text: `<strong>CR46 N1 thấp (${cr46_n1.toFixed(1)}%)</strong> — Gọn trial nhưng không chốt được. Cần review kỹ năng close và trial quality.` });
  else if (cr46_n1 < (parseFloat(bm.cr46_n1) || 50))
    actions.push({ type: 'improve', icon: '🎯', text: `<strong>CR46 N1 cần cải thiện (${cr46_n1.toFixed(1)}%)</strong> — Benchmark ${bm.cr46_n1 || 50}%. Cải thiện trial quality và follow-up trong 24h.` });

  if (noshow > (parseFloat(bm.noshow) || 20))
    actions.push({ type: 'improve', icon: '📋', text: `<strong>No-show ${noshow.toFixed(1)}%</strong> — Benchmark ${bm.noshow || 20}%. Gọi confirm trước 24h, gửi SMS nhắc lịch.` });

  if (n2_cs_avg < WM_BENCH.cs_touch * 0.6)
    actions.push({ type: 'alert', icon: '🔄', text: `<strong>CS Touchpoint quá thấp (${n2_cs_avg.toFixed(1)}/ngày)</strong> — Chưa đủ chăm sóc học viên. Cần tăng cuộc gọi/nhắn CS.` });

  if (n3_direct_avg < WM_BENCH.direct * 0.5)
    actions.push({ type: 'improve', icon: '🚀', text: `<strong>Direct Sales yếu (${n3_direct_avg.toFixed(1)}/ngày)</strong> — Cần đẩy mạnh outbound và events tại địa phương.` });

  if (actions.length === 0)
    actions.push({ type: 'good', icon: '🌟', text: 'Các chỉ số công việc và kết quả đều tốt. Tiếp tục duy trì!' });

  const fcEl = document.getElementById('cm-ana-forecast');
  if (fcEl) {
    fcEl.innerHTML = `
      <div class="ana-fc-grid">
        <div class="ana-fc-card">
          <div class="ana-fc-label">Deal dự báo EOM</div>
          <div class="ana-fc-value">${forecastDeal}</div>
          <div class="ana-fc-sub">Hiện tại: ${totalDeals} / ${daysPassed} ngày</div>
        </div>
        <div class="ana-fc-card">
          <div class="ana-fc-label">Doanh số dự báo EOM</div>
          <div class="ana-fc-value ana-fc-good">${forecastRev}M</div>
          <div class="ana-fc-sub">Hiện tại: ${totalRev}M</div>
        </div>
        <div class="ana-fc-card">
          <div class="ana-fc-label">Nếu tăng Calls đạt BM</div>
          <div class="ana-fc-value ${potentialExtraDeals > 0 ? 'ana-fc-warn' : 'ana-fc-good'}">${potentialExtraDeals > 0 ? '+' + potentialExtraDeals + ' deal' : 'Đạt rồi'}</div>
          <div class="ana-fc-sub">${potentialExtraDeals > 0 ? `Thêm ${callGap.toFixed(0)} calls/ngày` : 'Calls đã đạt benchmark'}</div>
        </div>
      </div>
      <div class="ana-fc-progress">
        <div class="ana-fc-progress-label">
          <span>Tiến độ tháng: ${daysPassed}/${daysInMonth} ngày (${progressPct.toFixed(0)}%)</span>
          <span>Pace kỳ vọng: ${(currentPace * 100).toFixed(0)}%</span>
        </div>
        <div class="ana-fc-progress-bar">
          <div class="ana-fc-progress-fill" style="width:${progressPct}%;background:${progressPct >= currentPace * 100 ? 'var(--green)' : 'var(--orange)'}"></div>
          <div class="ana-fc-progress-marker" style="left:${currentPace * 100}%" title="Pace kỳ vọng"></div>
        </div>
      </div>
      <div class="ana-fc-actions">
        ${actions.map(a => `<div class="ana-fc-action-item ${a.type}"><span class="ana-fc-action-icon">${a.icon}</span><span class="ana-fc-action-text">${a.text}</span></div>`).join('')}
      </div>
    `;
  }

  // SECTION 5: TOP BU REFERENCE (compare your BU vs top 3)
  const topEl = document.getElementById('cm-ana-top-bu');
  if (topEl && allData) {
    const allBUs = [...new Set(allData.map(r => r.bu))];
    if (allBUs.length > 1) {
      const buScores = allBUs.map(b => {
        const bRows = allData.filter(r => r.bu === b);
        const bDays = [...new Set(bRows.map(r => r.date))].length || 1;
        return {
          bu: b,
          isCurrent: b === bu,
          calls: _sum(bRows, 'calls_n1') / bDays,
          trial: _sum(bRows, 'trial_book') / bDays,
          cs: _sum(bRows, 'cs_touchpoint') / bDays,
          deals: (_sum(bRows,'deal_n1') + _sum(bRows,'deal_reupsell') + _sum(bRows,'deal_referral') + _sum(bRows,'deal_n3')) / bDays,
          rev: (_sum(bRows,'revenue_n1') + _sum(bRows,'revenue_n2') + _sum(bRows,'revenue_n3')) / bDays,
          cr16: _pct(_sum(bRows,'deal_n1'), _sum(bRows,'lead_mkt')),
          cr46: _pct(_sum(bRows,'deal_n1'), _sum(bRows,'trial_done_n1'))
        };
      }).sort((a, b) => b.deals - a.deals);

      // Get top 3 + always include current BU
      const top3 = buScores.slice(0, 3);
      const currentInTop = top3.find(b => b.isCurrent);
      const currentBU = buScores.find(b => b.isCurrent);
      const currentRank = buScores.findIndex(b => b.isCurrent) + 1;

      let rows = top3.map((b, i) => {
        const highlight = b.isCurrent ? 'style="background:#fff5f5;border-left:3px solid var(--red)"' : '';
        return `<div class="ana-top-row" ${highlight}>
          <div class="ana-top-bu-name">${i === 0 ? '🥇' : i === 1 ? '🥈' : '🥉'} ${b.bu} ${b.isCurrent ? '<span style="color:var(--red);font-size:11px">(Bạn)</span>' : ''}</div>
          <div class="ana-top-metrics">
            <span class="ana-top-metric">Calls: <strong>${b.calls.toFixed(1)}/d</strong></span>
            <span class="ana-top-metric">Trial: <strong>${b.trial.toFixed(1)}/d</strong></span>
            <span class="ana-top-metric">CS: <strong>${b.cs.toFixed(1)}/d</strong></span>
            <span class="ana-top-metric">Deal: <strong>${b.deals.toFixed(1)}/d</strong></span>
            <span class="ana-top-metric">Rev: <strong>${b.rev.toFixed(0)}M/d</strong></span>
            <span class="ana-top-metric">CR16: <strong>${b.cr16.toFixed(1)}%</strong></span>
          </div>
        </div>`;
      }).join('');

      // If current BU not in top 3, add it separately
      if (!currentInTop && currentBU) {
        rows += `<div style="border-top:2px dashed var(--gray-200);margin:8px 0"></div>
          <div class="ana-top-row" style="background:#fff5f5;border-left:3px solid var(--red)">
            <div class="ana-top-bu-name">#${currentRank} ${currentBU.bu} <span style="color:var(--red);font-size:11px">(Bạn)</span></div>
            <div class="ana-top-metrics">
              <span class="ana-top-metric">Calls: <strong>${currentBU.calls.toFixed(1)}/d</strong></span>
              <span class="ana-top-metric">Trial: <strong>${currentBU.trial.toFixed(1)}/d</strong></span>
              <span class="ana-top-metric">CS: <strong>${currentBU.cs.toFixed(1)}/d</strong></span>
              <span class="ana-top-metric">Deal: <strong>${currentBU.deals.toFixed(1)}/d</strong></span>
              <span class="ana-top-metric">Rev: <strong>${currentBU.rev.toFixed(0)}M/d</strong></span>
              <span class="ana-top-metric">CR16: <strong>${currentBU.cr16.toFixed(1)}%</strong></span>
            </div>
          </div>`;
      }

      topEl.innerHTML = `<div class="ana-top-grid">${rows}</div>`;
    } else {
      topEl.innerHTML = '<div style="padding:16px;color:var(--gray-400);font-size:12px;font-style:italic">Chưa đủ dữ liệu các BU để so sánh</div>';
    }
  }

  // SECTION 6: MiMi context
  const mimiCtx = document.getElementById('cm-ana-mimi-context');
  if (mimiCtx) {
    mimiCtx.textContent = JSON.stringify({
      bu, month, daysPassed, daysRemaining,
      workMetrics: {
        n1: { calls_avg: n1_calls_avg.toFixed(1), trial_book_avg: n1_trial_book_avg.toFixed(1), leads: n1_leads },
        n2: { cs_avg: n2_cs_avg.toFixed(1), phhs_avg: n2_phhs_avg.toFixed(1) },
        n3: { direct_avg: n3_direct_avg.toFixed(1), events_avg: n3_events_avg.toFixed(1) }
      },
      conversion: { cr16_n1: cr16_n1.toFixed(1), cr46_n1: cr46_n1.toFixed(1), noshow: noshow.toFixed(1), cr_re: cr_re.toFixed(1), cr_ref: cr_ref.toFixed(1), cr16_n3: cr16_n3.toFixed(1) },
      results: { totalDeals, totalRev, forecastDeal, forecastRev },
      diagnostic: { type: diagType, workScore: workScore.toFixed(0) }
    });
  }
}

// Direct Perplexity API call for MiMi Analytics (bypass Apps Script)
const PPLX_KEY = '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';

function mimiMarkdown(text) {
  return text
    .replace(/^### (.+)$/gm, '<h4 style="margin:12px 0 4px;font-size:13px;color:var(--red)">$1</h4>')
    .replace(/^## (.+)$/gm, '<h3 style="margin:14px 0 6px;font-size:14px;color:var(--black)">$1</h3>')
    .replace(/^# (.+)$/gm, '<h2 style="margin:16px 0 8px;font-size:15px;font-weight:800;color:var(--red)">$1</h2>')
    .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
    .replace(/^- (.+)$/gm, '<div style="padding-left:12px">\u2022 $1</div>')
    .replace(/\n/g, '<br>');
}
async function callMiMiDirect(systemPrompt, userMsg) {
  const resp = await fetch('https://api.perplexity.ai/chat/completions', {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' + PPLX_KEY,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      model: 'sonar-reasoning-pro',
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userMsg }
      ],
      max_tokens: 2000,
      temperature: 0.5,
      search_recency_filter: 'month'
    })
  });
  if (!resp.ok) throw new Error('API HTTP ' + resp.status);
  const data = await resp.json();
  if (!data.choices || !data.choices.length) throw new Error('No AI response');
  let answer = data.choices[0].message.content.trim();
  answer = answer.replace(/<think>[\s\S]*?<\/think>/g, '').trim();
  return answer;
}

function buildMiMiSystemPrompt(ctx) {
  return `Bạn là MiMi — trợ lý AI chiến lược của MindX Technology School.
Bạn có dữ liệu nội bộ MindX và cần kết hợp với nghiên cứu thị trường EdTech.

Dữ liệu nội bộ: ${ctx}

Quy tắc:
- Trả lời bằng tiếng Việt, giọng chuyên nghiệp, rõ ràng
- KHÔNG dùng markdown table — chỉ dùng danh sách đánh số, bullet points, chữ in đậm
- Đưa ra gợi ý cụ thể, hành động được
- Kết hợp dữ liệu nội bộ với xu hướng EdTech Việt Nam/Đông Nam Á
- Benchmark: CR16=15%, CR46=50%, AOV=18M, No-show≤20%
- Work metrics/BU/ngày: Calls=15, Trial Book=3, CS=8, PHHS=3, Direct=2`;
}

// MiMi AI for CM Analytics
async function askCMMiMiAnalytics() {
  const btn = document.getElementById('cm-ana-mimi-btn');
  const respEl = document.getElementById('cm-ana-mimi-response');
  const ctxEl = document.getElementById('cm-ana-mimi-context');
  if (!btn || !respEl) return;

  const ctx = ctxEl ? ctxEl.textContent : '{}';
  btn.disabled = true;
  btn.innerHTML = 'Đang phân tích...';
  respEl.style.display = '';
  respEl.innerHTML = '<span style="color:var(--gray-400)">MiMi đang nghiên cứu thị trường và phân tích dữ liệu BU của bạn... (có thể mất 15-30 giây)</span>';

  try {
    const sysPrompt = buildMiMiSystemPrompt(ctx);
    const userMsg = `Phân tích BU ${state.cm.bu || ''}:\n1. Đánh giá mức độ "chăm chỉ" và "hiệu quả" của BU này\n2. Chỉ ra chỉ số công việc nào cần ưu tiên cải thiện nhất\n3. Đưa ra 3 gợi ý cụ thể để cải thiện kết quả trong tháng này\n\nHãy kết hợp dữ liệu nội bộ với nghiên cứu từ thị trường giáo dục công nghệ, xu hướng EdTech Việt Nam/Đông Nam Á để đưa ra chiến lược cụ thể.`;

    const answer = await callMiMiDirect(sysPrompt, userMsg);
    respEl.innerHTML = mimiMarkdown(answer);
  } catch(e) {
    console.error('CM MiMi analytics error:', e);
    respEl.innerHTML = `<span style="color:var(--red)">Lỗi kết nối MiMi: ${e.message}. Vui lòng thử lại.</span>`;
  } finally {
    btn.disabled = false;
    btn.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg> Hỏi MiMi phân tích sâu';
  }
}

// ============================================================
// ANALYTICAL TAB v2 — Work Metrics → Conversion → Diagnostic → Forecast
// Multi-BU: checkbox dropdown with region shortcuts
// ============================================================

state.analytics = {
  month: '',
  rawData: null,
  charts: {},
  selectedBUs: [] // array of selected BU names; empty = ALL
};

// Region grouping for shortcuts
const ANA_REGIONS = [
  { key: 'ALL', label: 'Tất cả' },
  { key: 'HCM1', label: 'HCM1' },
  { key: 'HCM2', label: 'HCM2' },
  { key: 'HCM3', label: 'HCM3' },
  { key: 'HCM4', label: 'HCM4' },
  { key: 'HN', label: 'HN' },
  { key: 'K18', label: 'K18' },
  { key: 'MB1', label: 'MB1' },
  { key: 'MB2', label: 'MB2' },
  { key: 'MN', label: 'MN' },
  { key: 'MT', label: 'MT' },
  { key: 'ONL', label: 'ONL' }
];

function getBUsForRegion(key) {
  if (key === 'ALL') return [...BU_LIST];
  return BU_LIST.filter(bu => bu.startsWith(key + ' '));
}

// Hook into switchDashTab
(function() {
  const origSwitch = switchDashTab;
  switchDashTab = function(tab) {
    origSwitch(tab);
    if (tab === 'analytics') initAnalytics();
  };
})();

function initAnalytics() {
  const monthSel = document.getElementById('ana-month');
  if (monthSel && monthSel.options.length <= 1) {
    monthSel.innerHTML = '';
    MONTHS.forEach(m => {
      const opt = document.createElement('option');
      opt.value = m; opt.textContent = m;
      monthSel.appendChild(opt);
    });
    const now = new Date();
    monthSel.value = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`;
  }
  // Build multi-select checkbox list
  const listEl = document.getElementById('ana-bu-list');
  const shortcutsEl = document.getElementById('ana-bu-shortcuts');
  if (listEl && listEl.children.length === 0) {
    BU_LIST.forEach(bu => {
      const item = document.createElement('label');
      item.className = 'ana-multi-item';
      item.innerHTML = `<input type="checkbox" value="${bu}" checked> ${bu}`;
      listEl.appendChild(item);
    });
  }
  if (shortcutsEl && shortcutsEl.children.length === 0) {
    ANA_REGIONS.forEach(r => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'ana-multi-shortcut' + (r.key === 'ALL' ? ' active' : '');
      btn.textContent = r.label;
      btn.dataset.region = r.key;
      btn.onclick = () => anaRegionToggle(r.key, btn);
      shortcutsEl.appendChild(btn);
    });
  }
  // Default: all selected
  state.analytics.selectedBUs = [...BU_LIST];
  updateAnaBULabel();
  if (!state.analytics.rawData) loadAnalytics();
}

function toggleAnaBUDropdown() {
  const dd = document.getElementById('ana-bu-dropdown');
  if (!dd) return;
  dd.style.display = dd.style.display === 'none' ? 'flex' : 'none';
}

// Close dropdown when clicking outside
document.addEventListener('click', function(e) {
  const wrap = document.getElementById('ana-bu-multi');
  const dd = document.getElementById('ana-bu-dropdown');
  if (wrap && dd && !wrap.contains(e.target)) {
    dd.style.display = 'none';
  }
});

function anaRegionToggle(key, btn) {
  const listEl = document.getElementById('ana-bu-list');
  if (!listEl) return;
  const checks = listEl.querySelectorAll('input[type="checkbox"]');
  if (key === 'ALL') {
    // Toggle all
    const allChecked = Array.from(checks).every(c => c.checked);
    checks.forEach(c => c.checked = !allChecked);
  } else {
    const regionBUs = getBUsForRegion(key);
    const regionChecks = Array.from(checks).filter(c => regionBUs.includes(c.value));
    const allChecked = regionChecks.every(c => c.checked);
    regionChecks.forEach(c => c.checked = !allChecked);
  }
  updateShortcutStates();
}

function updateShortcutStates() {
  const listEl = document.getElementById('ana-bu-list');
  const shortcutsEl = document.getElementById('ana-bu-shortcuts');
  if (!listEl || !shortcutsEl) return;
  const checks = listEl.querySelectorAll('input[type="checkbox"]');
  const checkedVals = new Set(Array.from(checks).filter(c => c.checked).map(c => c.value));
  shortcutsEl.querySelectorAll('.ana-multi-shortcut').forEach(btn => {
    const key = btn.dataset.region;
    const regionBUs = getBUsForRegion(key);
    const allChecked = regionBUs.every(bu => checkedVals.has(bu));
    btn.classList.toggle('active', allChecked);
  });
}

function anaBUSelectAll() {
  const listEl = document.getElementById('ana-bu-list');
  if (!listEl) return;
  listEl.querySelectorAll('input[type="checkbox"]').forEach(c => c.checked = true);
  updateShortcutStates();
}
function anaBUClearAll() {
  const listEl = document.getElementById('ana-bu-list');
  if (!listEl) return;
  listEl.querySelectorAll('input[type="checkbox"]').forEach(c => c.checked = false);
  updateShortcutStates();
}

function anaBUApply() {
  const listEl = document.getElementById('ana-bu-list');
  if (!listEl) return;
  const checks = listEl.querySelectorAll('input[type="checkbox"]');
  const selected = Array.from(checks).filter(c => c.checked).map(c => c.value);
  state.analytics.selectedBUs = selected.length > 0 ? selected : [...BU_LIST];
  updateAnaBULabel();
  document.getElementById('ana-bu-dropdown').style.display = 'none';
  renderAnalytics();
}

function updateAnaBULabel() {
  const labelEl = document.getElementById('ana-bu-label');
  if (!labelEl) return;
  const sel = state.analytics.selectedBUs;
  if (sel.length === BU_LIST.length || sel.length === 0) {
    labelEl.textContent = `Tất cả BU (${BU_LIST.length})`;
  } else if (sel.length === 1) {
    labelEl.textContent = sel[0];
  } else {
    labelEl.textContent = `${sel.length} BU đã chọn`;
  }
}

async function loadAnalytics() {
  const month = document.getElementById('ana-month')?.value;
  if (!month) return;
  state.analytics.month = month;

  const loading = document.getElementById('ana-loading');
  const empty = document.getElementById('ana-empty');
  const content = document.getElementById('ana-content');
  if (loading) loading.style.display = '';
  if (empty) empty.style.display = 'none';
  if (content) content.style.display = 'none';

  try {
    let allData = null;
    if (state.dailyOps.rawData && state.dailyOps.month === month) {
      allData = state.dailyOps.rawData;
    } else {
      try {
        const r = await apiFetch('getDailyAll', { month });
        if (r.ok && r.data) allData = r.data;
      } catch(e) {}
      if (!allData) {
        allData = [];
        const results = await Promise.all(BU_LIST.map(async bu => {
          try { const r = await apiFetch('getDailyMTD', { bu, month }); return r.ok ? r.data : []; } catch(e) { return []; }
        }));
        results.forEach(rows => { allData = allData.concat(rows); });
      }
    }
    if (loading) loading.style.display = 'none';
    if (allData.length > 0) {
      state.analytics.rawData = allData;
      if (content) content.style.display = '';
      renderAnalytics();
    } else {
      state.analytics.rawData = null;
      if (empty) { empty.style.display = ''; empty.textContent = `Chưa có dữ liệu cho tháng ${month}`; }
    }
  } catch(err) {
    console.error('loadAnalytics error:', err);
    if (loading) loading.style.display = 'none';
    if (empty) { empty.style.display = ''; empty.textContent = 'Lỗi khi tải dữ liệu.'; }
  }
}

// === Helpers ===
function _sum(rows, field) { return rows.reduce((s, r) => s + (parseInt(r[field]) || 0), 0); }
function _pct(num, den) { return den > 0 ? (num / den * 100) : 0; }
function _badge(val, good, warn) {
  if (val >= good) return { cls: 'good', text: 'Đạt' };
  if (val >= warn) return { cls: 'warn', text: 'Cận' };
  return { cls: 'bad', text: 'Thấp' };
}
function _barColor(val, good, warn) {
  if (val >= good) return 'var(--green)';
  if (val >= warn) return 'var(--orange)';
  return 'var(--red)';
}

// ============================================================
// Helper: compute per-BU stats for comparison tables
// ============================================================
function _computeBUStats(data, uniqueBUs) {
  const WM_BENCH = { calls: 15, trial_book: 3, cs_touch: 8, phhs: 3, direct: 2, events: 0.1 };
  return uniqueBUs.map(bu => {
    const rows = data.filter(r => r.bu === bu);
    const days = [...new Set(rows.map(r => r.date))].length || 1;
    const n1c = _sum(rows,'calls_n1'), n1tb = _sum(rows,'trial_book'), n1td = _sum(rows,'trial_done_n1');
    const n1l = _sum(rows,'lead_mkt'), n1d = _sum(rows,'deal_n1'), n1r = _sum(rows,'revenue_n1');
    const n2cs = _sum(rows,'cs_touchpoint'), n2ph = _sum(rows,'phhs_reupsell');
    const n2dr = _sum(rows,'deal_reupsell'), n2dref = _sum(rows,'deal_referral'), n2r = _sum(rows,'revenue_n2');
    const n3dir = _sum(rows,'direct_sales'), n3ev = _sum(rows,'events');
    const n3l = _sum(rows,'lead_n3'), n3t = _sum(rows,'trial_n3'), n3d = _sum(rows,'deal_n3'), n3r = _sum(rows,'revenue_n3');
    const totalDeals = n1d + n2dr + n2dref + n3d;
    const totalRev = n1r + n2r + n3r;
    const ws = ((n1c/days)/WM_BENCH.calls + (n1tb/days)/WM_BENCH.trial_book + (n2cs/days)/WM_BENCH.cs_touch + (n3dir/days)/WM_BENCH.direct) / 4 * 100;
    return {
      bu, days,
      calls: n1c/days, trial_book: n1tb/days, cs: n2cs/days, phhs: n2ph/days,
      direct: n3dir/days, events: n3ev/days,
      leads_n1: n1l, deals_n1: n1d, trial_done_n1: n1td, rev_n1: n1r,
      deals_n2_re: n2dr, deals_n2_ref: n2dref, rev_n2: n2r,
      leads_n3: n3l, trial_n3: n3t, deals_n3: n3d, rev_n3: n3r,
      totalDeals, totalRev,
      dealsPerDay: totalDeals / days, revPerDay: totalRev / days,
      cr16_n1: _pct(n1d, n1l), cr46_n1: _pct(n1d, n1td),
      noshow: _pct(n1tb - n1td, n1tb),
      cr_re: _pct(n2dr, n2ph), cr_ref: _pct(n2dref, n2cs),
      cr16_n3: _pct(n3d, n3l), cr46_n3: _pct(n3d, n3t),
      workScore: ws
    };
  });
}

function _cmpCls(val, good, warn, invert) {
  if (invert) return val <= good ? 'cmp-good' : val <= warn ? 'cmp-warn' : 'cmp-bad';
  return val >= good ? 'cmp-good' : val >= warn ? 'cmp-warn' : 'cmp-bad';
}
function _cmpBar(val, maxVal, bench) {
  const pct = maxVal > 0 ? Math.min(val / maxVal * 100, 100) : 0;
  const color = val >= bench ? 'var(--green)' : val >= bench * 0.6 ? 'var(--orange)' : 'var(--red)';
  return `<div class="cmp-bar-wrap"><div class="cmp-bar-fill" style="width:${pct}%;background:${color}"></div><div class="cmp-bar-bench" style="left:${Math.min(bench/maxVal*100,100)}%"></div></div>`;
}
function _cmpBarInv(val, maxVal, bench) {
  const pct = maxVal > 0 ? Math.min(val / maxVal * 100, 100) : 0;
  const color = val <= bench ? 'var(--green)' : val <= bench * 1.5 ? 'var(--orange)' : 'var(--red)';
  return `<div class="cmp-bar-wrap"><div class="cmp-bar-fill" style="width:${pct}%;background:${color}"></div><div class="cmp-bar-bench" style="left:${Math.min(bench/maxVal*100,100)}%"></div></div>`;
}

function renderAnalytics() {
  const data = state.analytics.rawData;
  if (!data) return;

  // Multi-BU filter
  const selectedBUs = state.analytics.selectedBUs;
  const isAllBU = selectedBUs.length === BU_LIST.length;
  const selectedSet = new Set(selectedBUs);
  const filtered = data.filter(r => selectedSet.has(r.bu));
  const month = state.analytics.month;
  const uniqueBUs = [...new Set(filtered.map(r => r.bu))];
  const multiMode = uniqueBUs.length > 1;

  const metaEl = document.getElementById('ana-meta');
  if (metaEl) metaEl.textContent = `${isAllBU ? uniqueBUs.length + ' BU' : uniqueBUs.length + ' BU \u0111\u00E3 ch\u1ECDn'} \u2014 Th\u00E1ng ${month}`;

  // Load benchmarks from config
  const bm = state.config.data?.benchmarks || CONFIG_DEFAULTS.benchmarks;
  const daysInMonth = new Date(parseInt(month.split('-')[0]), parseInt(month.split('-')[1]), 0).getDate();
  const sortedDays = [...new Set(filtered.map(r => r.date))].sort();
  const daysPassed = sortedDays.length > 0 ? Math.max(...sortedDays.map(d => parseInt(d.split('-')[2]))) : 0;
  const daysRemaining = daysInMonth - daysPassed;
  const buCount = uniqueBUs.length;
  const avgDays = daysPassed > 0 ? daysPassed : 1;

  // ============================
  // SECTION 1: WORK METRICS SCORECARD (Summary)
  // ============================
  const n1_leads = _sum(filtered, 'lead_mkt');
  const n1_calls = _sum(filtered, 'calls_n1');
  const n1_trial_book = _sum(filtered, 'trial_book');
  const n1_trial_done = _sum(filtered, 'trial_done_n1');
  const n1_deals = _sum(filtered, 'deal_n1');
  const n1_rev = _sum(filtered, 'revenue_n1');
  const n1_calls_avg = buCount > 0 ? (n1_calls / buCount / avgDays) : 0;
  const n1_trial_book_avg = buCount > 0 ? (n1_trial_book / buCount / avgDays) : 0;

  const n2_cs = _sum(filtered, 'cs_touchpoint');
  const n2_phhs = _sum(filtered, 'phhs_reupsell');
  const n2_deal_re = _sum(filtered, 'deal_reupsell');
  const n2_deal_ref = _sum(filtered, 'deal_referral');
  const n2_rev = _sum(filtered, 'revenue_n2');
  const n2_cs_avg = buCount > 0 ? (n2_cs / buCount / avgDays) : 0;
  const n2_phhs_avg = buCount > 0 ? (n2_phhs / buCount / avgDays) : 0;

  const n3_direct = _sum(filtered, 'direct_sales');
  const n3_events = _sum(filtered, 'events');
  const n3_leads = _sum(filtered, 'lead_n3');
  const n3_trial = _sum(filtered, 'trial_n3');
  const n3_deals = _sum(filtered, 'deal_n3');
  const n3_rev = _sum(filtered, 'revenue_n3');
  const n3_direct_avg = buCount > 0 ? (n3_direct / buCount / avgDays) : 0;
  const n3_events_avg = buCount > 0 ? (n3_events / buCount / avgDays) : 0;

  const WM_BENCH = { calls: 15, trial_book: 3, cs_touch: 8, phhs: 3, direct: 2, events: 0.1 };

  function wmRow(label, total, avgPerDay, bench, unit) {
    const pct = bench > 0 ? Math.min(avgPerDay / bench * 100, 150) : 0;
    const b = _badge(avgPerDay, bench, bench * 0.6);
    const color = _barColor(avgPerDay, bench, bench * 0.6);
    return `<div class="ana-wm-row">
      <span class="ana-wm-label">${label}</span>
      <span class="ana-wm-value">${total}</span>
      <div class="ana-wm-bar-wrap"><div class="ana-wm-bar" style="width:${Math.min(pct,100)}%;background:${color}"></div></div>
      <span class="ana-wm-badge ${b.cls}">${avgPerDay.toFixed(1)}${unit}/BU/ng\u00E0y</span>
      <span class="ana-wm-benchmark">(B: ${bench})</span>
    </div>`;
  }

  const wmEl = document.getElementById('ana-work-metrics');
  if (wmEl) {
    wmEl.innerHTML = `<div class="ana-wm-grid">
      <div class="ana-wm-source">
        <div class="ana-wm-source-title n1">N1 \u2014 Optimize (MKT)</div>
        ${wmRow('Calls (g\u1ECDi ra)', n1_calls, n1_calls_avg, WM_BENCH.calls, '')}
        ${wmRow('Trial Book (\u0111\u1EB7t h\u1EB9n)', n1_trial_book, n1_trial_book_avg, WM_BENCH.trial_book, '')}
        <div class="ana-wm-row">
          <span class="ana-wm-label">Lead MKT (nh\u1EADn)</span>
          <span class="ana-wm-value">${n1_leads}</span>
          <span class="ana-wm-badge" style="background:var(--blue-bg);color:#1a5aa0">${(n1_leads / Math.max(buCount,1) / avgDays).toFixed(1)}/BU/ng\u00E0y</span>
        </div>
      </div>
      <div class="ana-wm-source">
        <div class="ana-wm-source-title n2">N2 \u2014 Ops (Re/Upsell)</div>
        ${wmRow('CS Touchpoint', n2_cs, n2_cs_avg, WM_BENCH.cs_touch, '')}
        ${wmRow('PHHS t\u01B0 v\u1EA5n', n2_phhs, n2_phhs_avg, WM_BENCH.phhs, '')}
        <div class="ana-wm-row">
          <span class="ana-wm-label">Deal ReUpsell</span>
          <span class="ana-wm-value">${n2_deal_re}</span>
          <span class="ana-wm-badge" style="background:var(--blue-bg);color:#1a5aa0">${(n2_deal_re / Math.max(buCount,1) / avgDays).toFixed(2)}/BU/ng\u00E0y</span>
        </div>
        <div class="ana-wm-row">
          <span class="ana-wm-label">Deal Referral</span>
          <span class="ana-wm-value">${n2_deal_ref}</span>
          <span class="ana-wm-badge" style="background:var(--blue-bg);color:#1a5aa0">${(n2_deal_ref / Math.max(buCount,1) / avgDays).toFixed(2)}/BU/ng\u00E0y</span>
        </div>
      </div>
      <div class="ana-wm-source">
        <div class="ana-wm-source-title n3">N3 \u2014 Growth (T\u1EF1 ki\u1EBFm)</div>
        ${wmRow('Direct Sales', n3_direct, n3_direct_avg, WM_BENCH.direct, '')}
        ${wmRow('Events', n3_events, n3_events_avg, WM_BENCH.events, '')}
        <div class="ana-wm-row">
          <span class="ana-wm-label">Lead N3</span>
          <span class="ana-wm-value">${n3_leads}</span>
          <span class="ana-wm-badge" style="background:var(--blue-bg);color:#1a5aa0">${(n3_leads / Math.max(buCount,1) / avgDays).toFixed(1)}/BU/ng\u00E0y</span>
        </div>
      </div>
    </div>`;
  }

  // ============================
  // SECTION 1b: Work Metrics per-BU COMPARISON TABLE
  // ============================
  const buStats = _computeBUStats(data, uniqueBUs);
  const wmCmpEl = document.getElementById('ana-wm-compare');
  const wmCmpBody = document.getElementById('ana-wm-compare-body');
  if (wmCmpEl && wmCmpBody) {
    if (multiMode) {
      wmCmpEl.style.display = '';
      const sorted = [...buStats].sort((a,b) => b.calls - a.calls);
      // Find max values for bar scaling
      const wmMax = {
        calls: Math.max(...sorted.map(s=>s.calls), WM_BENCH.calls) * 1.1,
        trial: Math.max(...sorted.map(s=>s.trial_book), WM_BENCH.trial_book) * 1.1,
        cs: Math.max(...sorted.map(s=>s.cs), WM_BENCH.cs_touch) * 1.1,
        phhs: Math.max(...sorted.map(s=>s.phhs), WM_BENCH.phhs) * 1.1,
        direct: Math.max(...sorted.map(s=>s.direct), WM_BENCH.direct) * 1.1
      };
      let rows = sorted.map(s => {
        return `<tr>
          <td>${s.bu}</td>
          <td class="${_cmpCls(s.calls, WM_BENCH.calls, WM_BENCH.calls*0.6)}">${s.calls.toFixed(1)} ${_cmpBar(s.calls, wmMax.calls, WM_BENCH.calls)}</td>
          <td class="${_cmpCls(s.trial_book, WM_BENCH.trial_book, WM_BENCH.trial_book*0.6)}">${s.trial_book.toFixed(1)} ${_cmpBar(s.trial_book, wmMax.trial, WM_BENCH.trial_book)}</td>
          <td class="${_cmpCls(s.cs, WM_BENCH.cs_touch, WM_BENCH.cs_touch*0.6)}">${s.cs.toFixed(1)} ${_cmpBar(s.cs, wmMax.cs, WM_BENCH.cs_touch)}</td>
          <td class="${_cmpCls(s.phhs, WM_BENCH.phhs, WM_BENCH.phhs*0.6)}">${s.phhs.toFixed(1)} ${_cmpBar(s.phhs, wmMax.phhs, WM_BENCH.phhs)}</td>
          <td class="${_cmpCls(s.direct, WM_BENCH.direct, WM_BENCH.direct*0.6)}">${s.direct.toFixed(1)} ${_cmpBar(s.direct, wmMax.direct, WM_BENCH.direct)}</td>
          <td>${s.events.toFixed(2)}</td>
        </tr>`;
      }).join('');
      rows += `<tr class="ana-cmp-summary-row">
        <td>TB (${buCount} BU)</td>
        <td>${n1_calls_avg.toFixed(1)}</td>
        <td>${n1_trial_book_avg.toFixed(1)}</td>
        <td>${n2_cs_avg.toFixed(1)}</td>
        <td>${n2_phhs_avg.toFixed(1)}</td>
        <td>${n3_direct_avg.toFixed(1)}</td>
        <td>${n3_events_avg.toFixed(2)}</td>
      </tr>`;
      wmCmpBody.innerHTML = `<table class="ana-cmp-table">
        <thead><tr>
          <th>BU</th><th>Calls/d</th><th>Trial/d</th><th>CS/d</th><th>PHHS/d</th><th>Direct/d</th><th>Evts/d</th>
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>
      <div style="font-size:10px;color:var(--gray-400);margin-top:6px">\u0110\u01B0\u1EDDng d\u1ECDc: Benchmark | Calls ${WM_BENCH.calls} | Trial ${WM_BENCH.trial_book} | CS ${WM_BENCH.cs_touch} | PHHS ${WM_BENCH.phhs} | Direct ${WM_BENCH.direct}</div>`;
    } else {
      wmCmpEl.style.display = 'none';
    }
  }

  // ============================
  // SECTION 2: CONVERSION EFFICIENCY (Summary)
  // ============================
  const cr16_n1 = _pct(n1_deals, n1_leads);
  const cr46_n1 = _pct(n1_deals, n1_trial_done);
  const noshow = _pct(n1_trial_book - n1_trial_done, n1_trial_book);
  const cr_re = _pct(n2_deal_re, n2_phhs);
  const cr_ref = _pct(n2_deal_ref, n2_cs);
  const cr16_n3 = _pct(n3_deals, n3_leads);
  const cr46_n3 = _pct(n3_deals, n3_trial);

  function crBadge(val, goodThreshold, warnThreshold, invert) {
    if (invert) {
      if (val <= goodThreshold) return 'good';
      if (val <= warnThreshold) return 'warn';
      return 'bad';
    }
    if (val >= goodThreshold) return 'good';
    if (val >= warnThreshold) return 'warn';
    return 'bad';
  }

  function funnelStep(num, label) {
    return `<div class="ana-funnel-step"><span class="ana-funnel-num">${num}</span><span class="ana-funnel-label">${label}</span></div>`;
  }
  function funnelArrow(crVal, crLabel, goodT, warnT, invert) {
    const cls = crBadge(crVal, goodT, warnT, invert);
    return `<div class="ana-funnel-step"><span class="ana-funnel-arrow">\u2193</span><span class="ana-funnel-cr ${cls}">${crLabel}: ${crVal.toFixed(1)}% (B: ${goodT}%)</span></div>`;
  }

  const convEl = document.getElementById('ana-conversion');
  if (convEl) {
    convEl.innerHTML = `<div class="ana-conv-grid">
      <div class="ana-conv-source">
        <div class="ana-conv-title n1">N1 \u2014 Funnel MKT</div>
        ${funnelStep(n1_leads, 'Lead MKT')}
        ${funnelArrow(cr16_n1, 'CR16', parseFloat(bm.cr16_n1)||15, 10, false)}
        ${funnelStep(n1_trial_book, 'Trial Book')}
        ${funnelArrow(noshow, 'No-show', parseFloat(bm.noshow)||20, 30, true)}
        ${funnelStep(n1_trial_done, 'Trial Done')}
        ${funnelArrow(cr46_n1, 'CR46', parseFloat(bm.cr46_n1)||50, 35, false)}
        ${funnelStep(n1_deals, 'Deal N1')}
        <div class="ana-funnel-step" style="margin-top:6px;border-top:1px solid var(--gray-200);padding-top:8px">
          <span class="ana-funnel-num" style="color:var(--red)">${n1_rev}M</span>
          <span class="ana-funnel-label">Doanh s\u1ED1 N1</span>
        </div>
      </div>
      <div class="ana-conv-source">
        <div class="ana-conv-title n2">N2 \u2014 Funnel Ops</div>
        ${funnelStep(n2_cs, 'CS Touchpoint')}
        ${funnelArrow(cr_re, 'CR ReUpsell', parseFloat(bm.cr_reupsell)||15, 8, false)}
        ${funnelStep(n2_deal_re, 'Deal ReUpsell')}
        <div style="height:12px"></div>
        ${funnelStep(n2_cs, 'CS Touchpoint')}
        ${funnelArrow(cr_ref, 'CR Referral', parseFloat(bm.cr_referral)||5, 2, false)}
        ${funnelStep(n2_deal_ref, 'Deal Referral')}
        <div class="ana-funnel-step" style="margin-top:6px;border-top:1px solid var(--gray-200);padding-top:8px">
          <span class="ana-funnel-num" style="color:#1a5aa0">${n2_rev}M</span>
          <span class="ana-funnel-label">Doanh s\u1ED1 N2</span>
        </div>
      </div>
      <div class="ana-conv-source">
        <div class="ana-conv-title n3">N3 \u2014 Funnel Growth</div>
        ${funnelStep(n3_leads, 'Lead N3')}
        ${funnelArrow(cr16_n3, 'CR16', parseFloat(bm.cr16_n3)||10, 5, false)}
        ${funnelStep(n3_trial, 'Trial N3')}
        ${funnelArrow(cr46_n3, 'CR46', parseFloat(bm.cr46_n3)||40, 25, false)}
        ${funnelStep(n3_deals, 'Deal N3')}
        <div class="ana-funnel-step" style="margin-top:6px;border-top:1px solid var(--gray-200);padding-top:8px">
          <span class="ana-funnel-num" style="color:var(--green)">${n3_rev}M</span>
          <span class="ana-funnel-label">Doanh s\u1ED1 N3</span>
        </div>
      </div>
    </div>`;
  }

  // ============================
  // SECTION 2b: Conversion per-BU COMPARISON TABLE
  // ============================
  const convCmpEl = document.getElementById('ana-conv-compare');
  const convCmpBody = document.getElementById('ana-conv-compare-body');
  if (convCmpEl && convCmpBody) {
    if (multiMode) {
      convCmpEl.style.display = '';
      const sorted = [...buStats].sort((a,b) => b.cr16_n1 - a.cr16_n1);
      const bmCR16 = parseFloat(bm.cr16_n1)||15, bmCR46 = parseFloat(bm.cr46_n1)||50;
      const bmNS = parseFloat(bm.noshow)||20, bmCRre = parseFloat(bm.cr_reupsell)||15;
      const bmCRref = parseFloat(bm.cr_referral)||5, bmCR16n3 = parseFloat(bm.cr16_n3)||10;
      const crMax = {
        cr16: Math.max(...sorted.map(s=>s.cr16_n1), bmCR16) * 1.15,
        cr46: Math.max(...sorted.map(s=>s.cr46_n1), bmCR46) * 1.15,
        ns: Math.max(...sorted.map(s=>s.noshow), bmNS) * 1.15,
        crRe: Math.max(...sorted.map(s=>s.cr_re), bmCRre) * 1.15,
        crRef: Math.max(...sorted.map(s=>s.cr_ref), bmCRref) * 1.15,
        cr16n3: Math.max(...sorted.map(s=>s.cr16_n3), bmCR16n3) * 1.15
      };
      let rows = sorted.map(s => `<tr>
        <td>${s.bu}</td>
        <td class="${_cmpCls(s.cr16_n1, bmCR16, bmCR16*0.7)}">${s.cr16_n1.toFixed(1)}% ${_cmpBar(s.cr16_n1, crMax.cr16, bmCR16)}</td>
        <td class="${_cmpCls(s.cr46_n1, bmCR46, bmCR46*0.7)}">${s.cr46_n1.toFixed(1)}% ${_cmpBar(s.cr46_n1, crMax.cr46, bmCR46)}</td>
        <td class="${_cmpCls(s.noshow, bmNS, bmNS*1.5, true)}">${s.noshow.toFixed(1)}% ${_cmpBarInv(s.noshow, crMax.ns, bmNS)}</td>
        <td class="${_cmpCls(s.cr_re, bmCRre, bmCRre*0.5)}">${s.cr_re.toFixed(1)}% ${_cmpBar(s.cr_re, crMax.crRe, bmCRre)}</td>
        <td class="${_cmpCls(s.cr_ref, bmCRref, bmCRref*0.4)}">${s.cr_ref.toFixed(1)}% ${_cmpBar(s.cr_ref, crMax.crRef, bmCRref)}</td>
        <td class="${_cmpCls(s.cr16_n3, 10, 5)}">${s.cr16_n3.toFixed(1)}% ${_cmpBar(s.cr16_n3, crMax.cr16n3, bmCR16n3)}</td>
      </tr>`).join('');
      // Summary row
      rows += `<tr class="ana-cmp-summary-row">
        <td>T\u1ED5ng</td>
        <td>${cr16_n1.toFixed(1)}%</td>
        <td>${cr46_n1.toFixed(1)}%</td>
        <td>${noshow.toFixed(1)}%</td>
        <td>${cr_re.toFixed(1)}%</td>
        <td>${cr_ref.toFixed(1)}%</td>
        <td>${cr16_n3.toFixed(1)}%</td>
      </tr>`;
      convCmpBody.innerHTML = `<table class="ana-cmp-table">
        <thead><tr>
          <th>BU</th><th>CR16 N1</th><th>CR46 N1</th><th>No-show</th><th>CR Re</th><th>CR Ref</th><th>CR16 N3</th>
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>
      <div style="font-size:10px;color:var(--gray-400);margin-top:6px">Benchmark: CR16 ${bmCR16}% | CR46 ${bmCR46}% | No-show \u2264${bmNS}% | CR Re ${bmCRre}% | CR Ref ${bmCRref}%</div>`;
    } else {
      convCmpEl.style.display = 'none';
    }
  }

  // ============================
  // SECTION 3: DIAGNOSTIC ENGINE
  // ============================
  const totalDeals = n1_deals + n2_deal_re + n2_deal_ref + n3_deals;
  const totalRev = n1_rev + n2_rev + n3_rev;
  const buDiag = { star: [], smart: [], effort: [] };

  if (multiMode) {
    buStats.forEach(s => {
      const isHW = s.workScore >= 70;
      const isGR = s.dealsPerDay >= 0.8 && s.cr16_n1 >= 12;
      const summary = `<strong>${s.bu}</strong>: Work ${s.workScore.toFixed(0)}% | ${s.dealsPerDay.toFixed(1)} deal/ng\u00E0y | CR16=${s.cr16_n1.toFixed(1)}%`;
      if (isHW && isGR) buDiag.star.push(summary);
      else if (isHW && !isGR) buDiag.smart.push(summary);
      else buDiag.effort.push(summary);
    });
  } else {
    const workScore = (n1_calls_avg/WM_BENCH.calls + n1_trial_book_avg/WM_BENCH.trial_book + n2_cs_avg/WM_BENCH.cs_touch + n3_direct_avg/WM_BENCH.direct) / 4 * 100;
    const dealPerDay = totalDeals / avgDays;
    const isHW = workScore >= 70;
    const isGR = dealPerDay >= 0.8 && cr16_n1 >= 12;
    const label = uniqueBUs[0] || 'BU';
    const summary = `<strong>${label}</strong>: Work ${workScore.toFixed(0)}% | ${dealPerDay.toFixed(1)} deal/ng\u00E0y | CR16 N1=${cr16_n1.toFixed(1)}%`;
    if (isHW && isGR) buDiag.star.push(summary);
    else if (isHW && !isGR) buDiag.smart.push(summary);
    else buDiag.effort.push(summary);
  }

  const diagEl = document.getElementById('ana-diagnostic');
  if (diagEl) {
    diagEl.innerHTML = `<div class="ana-diag-cards">
      <div class="ana-diag-card star">
        <div class="ana-diag-card-title">\u2B50 Ch\u0103m ch\u1EC9 + K\u1EBFt qu\u1EA3 t\u1ED1t</div>
        <div class="ana-diag-card-desc">Ti\u1EBFp t\u1EE5c duy tr\u00EC, c\u00F3 th\u1EC3 chia s\u1EBB kinh nghi\u1EC7m</div>
        <div class="ana-diag-items">${buDiag.star.length > 0 ? buDiag.star.join('<br>') : '<span class="ana-diag-empty">Ch\u01B0a c\u00F3 BU</span>'}</div>
      </div>
      <div class="ana-diag-card smart">
        <div class="ana-diag-card-title">\uD83E\uDDE0 Ch\u0103m ch\u1EC9 nh\u01B0ng KQ th\u1EA5p</div>
        <div class="ana-diag-card-desc">C\u1EA7n l\u00E0m vi\u1EC7c th\u00F4ng minh h\u01A1n \u2192 H\u1ECFi MiMi</div>
        <div class="ana-diag-items">${buDiag.smart.length > 0 ? buDiag.smart.join('<br>') : '<span class="ana-diag-empty">Ch\u01B0a c\u00F3 BU</span>'}</div>
      </div>
      <div class="ana-diag-card effort">
        <div class="ana-diag-card-title">\uD83D\uDCAA Ch\u01B0a \u0111\u1EE7 n\u1ED7 l\u1EF1c</div>
        <div class="ana-diag-card-desc">C\u1EA7n t\u0103ng c\u01B0\u1EDDng ch\u1EC9 s\u1ED1 c\u00F4ng vi\u1EC7c</div>
        <div class="ana-diag-items">${buDiag.effort.length > 0 ? buDiag.effort.join('<br>') : '<span class="ana-diag-empty">Ch\u01B0a c\u00F3 BU</span>'}</div>
      </div>
    </div>`;
  }

  // ============================
  // SECTION 4: EOM FORECAST (Summary)
  // ============================
  const dailyDealRate = daysPassed > 0 ? totalDeals / daysPassed : 0;
  const dailyRevRate = daysPassed > 0 ? totalRev / daysPassed : 0;
  const forecastDeal = Math.round(dailyDealRate * daysInMonth);
  const forecastRev = Math.round(dailyRevRate * daysInMonth);
  const callGap = Math.max(0, WM_BENCH.calls - n1_calls_avg);
  const potentialExtraTrials = callGap * buCount * daysRemaining * 0.2;
  const potentialExtraDeals = Math.round(potentialExtraTrials * (cr46_n1 / 100) || 0);

  const pace = state.config.data?.week_pace || CONFIG_DEFAULTS.week_pace;
  function currentWeekKey() {
    const day = new Date().getDate();
    if (day <= 7) return 'w1'; if (day <= 14) return 'w2'; if (day <= 21) return 'w3'; return 'w4';
  }
  const currentPace = pace[currentWeekKey()] || 0.4;
  const progressPct = daysPassed > 0 ? (daysPassed / daysInMonth * 100) : 0;

  const actions = [];
  if (n1_calls_avg < WM_BENCH.calls * 0.6)
    actions.push({ type: 'alert', icon: '\uD83D\uDCDE', text: `<strong>Calls qu\u00E1 th\u1EA5p (${n1_calls_avg.toFixed(1)}/ng\u00E0y)</strong> \u2014 Benchmark ${WM_BENCH.calls}/ng\u00E0y. T\u0103ng calls c\u00F3 th\u1EC3 th\u00EAm ~${potentialExtraDeals} deal cu\u1ED1i th\u00E1ng.` });
  else if (n1_calls_avg < WM_BENCH.calls)
    actions.push({ type: 'improve', icon: '\uD83D\uDCDE', text: `<strong>Calls c\u1EA7n c\u1EA3i thi\u1EC7n (${n1_calls_avg.toFixed(1)}/${WM_BENCH.calls}/ng\u00E0y)</strong> \u2014 T\u0103ng th\u00EAm ${callGap.toFixed(0)} calls/BU/ng\u00E0y \u0111\u1EC3 \u0111\u1EA1t benchmark.` });
  else
    actions.push({ type: 'good', icon: '\u2705', text: `<strong>Calls \u0111\u1EA1t chu\u1EA9n (${n1_calls_avg.toFixed(1)}/ng\u00E0y)</strong> \u2014 Ti\u1EBFp t\u1EE5c duy tr\u00EC.` });
  if (cr46_n1 < (parseFloat(bm.cr46_n1) || 50) * 0.7)
    actions.push({ type: 'alert', icon: '\uD83C\uDFAF', text: `<strong>CR46 N1 th\u1EA5p (${cr46_n1.toFixed(1)}%)</strong> \u2014 C\u1EA7n review k\u1EF9 n\u0103ng close v\u00E0 trial quality.` });
  else if (cr46_n1 < (parseFloat(bm.cr46_n1) || 50))
    actions.push({ type: 'improve', icon: '\uD83C\uDFAF', text: `<strong>CR46 N1 c\u1EA7n c\u1EA3i thi\u1EC7n (${cr46_n1.toFixed(1)}%)</strong> \u2014 Benchmark ${bm.cr46_n1 || 50}%.` });
  if (noshow > (parseFloat(bm.noshow) || 20))
    actions.push({ type: 'improve', icon: '\uD83D\uDCCB', text: `<strong>No-show ${noshow.toFixed(1)}%</strong> \u2014 Benchmark ${bm.noshow || 20}%.` });
  if (n2_cs_avg < WM_BENCH.cs_touch * 0.6)
    actions.push({ type: 'alert', icon: '\uD83D\uDD04', text: `<strong>CS Touchpoint qu\u00E1 th\u1EA5p (${n2_cs_avg.toFixed(1)}/ng\u00E0y)</strong> \u2014 C\u1EA7n t\u0103ng cu\u1ED9c g\u1ECDi CS.` });
  if (n3_direct_avg < WM_BENCH.direct * 0.5)
    actions.push({ type: 'improve', icon: '\uD83D\uDE80', text: `<strong>Direct Sales y\u1EBFu (${n3_direct_avg.toFixed(1)}/ng\u00E0y)</strong> \u2014 C\u1EA7n \u0111\u1EA9y m\u1EA1nh outbound.` });
  if (actions.length === 0)
    actions.push({ type: 'good', icon: '\uD83C\uDF1F', text: 'C\u00E1c ch\u1EC9 s\u1ED1 \u0111\u1EC1u t\u1ED1t. Ti\u1EBFp t\u1EE5c duy tr\u00EC!' });

  const fcEl = document.getElementById('ana-forecast-v2');
  if (fcEl) {
    fcEl.innerHTML = `
      <div class="ana-fc-grid">
        <div class="ana-fc-card">
          <div class="ana-fc-label">Deal d\u1EF1 b\u00E1o EOM</div>
          <div class="ana-fc-value">${forecastDeal}</div>
          <div class="ana-fc-sub">Hi\u1EC7n t\u1EA1i: ${totalDeals} / ${daysPassed} ng\u00E0y</div>
        </div>
        <div class="ana-fc-card">
          <div class="ana-fc-label">Doanh s\u1ED1 d\u1EF1 b\u00E1o EOM</div>
          <div class="ana-fc-value ana-fc-good">${forecastRev}M</div>
          <div class="ana-fc-sub">Hi\u1EC7n t\u1EA1i: ${totalRev}M</div>
        </div>
        <div class="ana-fc-card">
          <div class="ana-fc-label">N\u1EBFu t\u0103ng Calls \u0111\u1EA1t BM</div>
          <div class="ana-fc-value ${potentialExtraDeals > 0 ? 'ana-fc-warn' : 'ana-fc-good'}">${potentialExtraDeals > 0 ? '+' + potentialExtraDeals + ' deal' : '\u0110\u1EA1t r\u1ED3i'}</div>
          <div class="ana-fc-sub">${potentialExtraDeals > 0 ? `Th\u00EAm ${callGap.toFixed(0)} calls/BU/ng\u00E0y` : 'Calls \u0111\u00E3 \u0111\u1EA1t benchmark'}</div>
        </div>
      </div>
      <div class="ana-fc-progress">
        <div class="ana-fc-progress-label">
          <span>Ti\u1EBFn \u0111\u1ED9 th\u00E1ng: ${daysPassed}/${daysInMonth} ng\u00E0y (${progressPct.toFixed(0)}%)</span>
          <span>Pace k\u1EF3 v\u1ECDng: ${(currentPace * 100).toFixed(0)}%</span>
        </div>
        <div class="ana-fc-progress-bar">
          <div class="ana-fc-progress-fill" style="width:${progressPct}%;background:${progressPct >= currentPace * 100 ? 'var(--green)' : 'var(--orange)'}"></div>
          <div class="ana-fc-progress-marker" style="left:${currentPace * 100}%" title="Pace"></div>
        </div>
      </div>
      <div class="ana-fc-actions">
        ${actions.map(a => `<div class="ana-fc-action-item ${a.type}"><span class="ana-fc-action-icon">${a.icon}</span><span class="ana-fc-action-text">${a.text}</span></div>`).join('')}
      </div>`;
  }

  // ============================
  // SECTION 4b: Forecast per-BU COMPARISON TABLE
  // ============================
  const fcCmpEl = document.getElementById('ana-fc-compare');
  const fcCmpBody = document.getElementById('ana-fc-compare-body');
  if (fcCmpEl && fcCmpBody) {
    if (multiMode) {
      fcCmpEl.style.display = '';
      const sorted = [...buStats].sort((a,b) => b.totalRev - a.totalRev);
      const fcMax = {
        deals: Math.max(...sorted.map(s=>s.totalDeals)) * 1.1 || 1,
        rev: Math.max(...sorted.map(s=>s.totalRev)) * 1.1 || 1,
        fcRev: Math.max(...sorted.map(s=>Math.round(s.revPerDay * daysInMonth))) * 1.1 || 1
      };
      let rows = sorted.map(s => {
        const fcDeal = Math.round(s.dealsPerDay * daysInMonth);
        const fcRev = Math.round(s.revPerDay * daysInMonth);
        const revPct = fcMax.rev > 0 ? Math.min(s.totalRev / fcMax.rev * 100, 100) : 0;
        const fcRevPct = fcMax.fcRev > 0 ? Math.min(fcRev / fcMax.fcRev * 100, 100) : 0;
        return `<tr>
          <td>${s.bu}</td>
          <td>${s.totalDeals}</td>
          <td>${s.totalRev}M<div class="cmp-bar-wrap"><div class="cmp-bar-fill" style="width:${revPct}%;background:var(--blue)"></div></div></td>
          <td>${s.dealsPerDay.toFixed(1)}</td>
          <td>${fcDeal}</td>
          <td><strong>${fcRev}M</strong><div class="cmp-bar-wrap"><div class="cmp-bar-fill" style="width:${fcRevPct}%;background:var(--green)"></div></div></td>
        </tr>`;
      }).join('');
      rows += `<tr class="ana-cmp-summary-row">
        <td>T\u1ED5ng</td>
        <td>${totalDeals}</td>
        <td>${totalRev}M</td>
        <td>${dailyDealRate.toFixed(1)}</td>
        <td>${forecastDeal}</td>
        <td><strong>${forecastRev}M</strong></td>
      </tr>`;
      fcCmpBody.innerHTML = `<table class="ana-cmp-table">
        <thead><tr>
          <th>BU</th><th>Deals</th><th>Rev</th><th>Deal/d</th><th>FC Deal</th><th>FC Rev</th>
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>`;
    } else {
      fcCmpEl.style.display = 'none';
    }
  }

  // ============================
  // SECTION 5: TOP BU REFERENCE (within selected BUs)
  // ============================
  const topEl = document.getElementById('ana-top-bu');
  if (topEl && multiMode) {
    const topScores = [...buStats].sort((a, b) => b.dealsPerDay - a.dealsPerDay);
    const top3 = topScores.slice(0, Math.min(3, topScores.length));
    topEl.innerHTML = `<div class="ana-top-grid">
      ${top3.map((b, i) => `<div class="ana-top-row">
        <div class="ana-top-bu-name">${i === 0 ? '\uD83E\uDD47' : i === 1 ? '\uD83E\uDD48' : '\uD83E\uDD49'} ${b.bu}</div>
        <div class="ana-top-metrics">
          <span class="ana-top-metric">Calls: <strong>${b.calls.toFixed(1)}/d</strong></span>
          <span class="ana-top-metric">Trial: <strong>${b.trial_book.toFixed(1)}/d</strong></span>
          <span class="ana-top-metric">CS: <strong>${b.cs.toFixed(1)}/d</strong></span>
          <span class="ana-top-metric">Deal: <strong>${b.dealsPerDay.toFixed(1)}/d</strong></span>
          <span class="ana-top-metric">Rev: <strong>${b.revPerDay.toFixed(0)}M/d</strong></span>
          <span class="ana-top-metric">CR16: <strong>${b.cr16_n1.toFixed(1)}%</strong></span>
          <span class="ana-top-metric">CR46: <strong>${b.cr46_n1.toFixed(1)}%</strong></span>
        </div>
      </div>`).join('')}
    </div>`;
  } else if (topEl) {
    topEl.innerHTML = '<div style="padding:16px;color:var(--gray-400);font-size:12px;font-style:italic">Ch\u1ECDn nhi\u1EC1u BU \u0111\u1EC3 xem top BU tham kh\u1EA3o</div>';
  }

  // ============================
  // SECTION 6: Prepare MiMi context
  // ============================
  const mimiCtx = document.getElementById('ana-mimi-context');
  if (mimiCtx) {
    mimiCtx.textContent = JSON.stringify({
      month, selectedBUs: uniqueBUs, daysPassed, daysRemaining,
      workMetrics: {
        n1: { calls_avg: n1_calls_avg.toFixed(1), trial_book_avg: n1_trial_book_avg.toFixed(1), leads: n1_leads },
        n2: { cs_avg: n2_cs_avg.toFixed(1), phhs_avg: n2_phhs_avg.toFixed(1) },
        n3: { direct_avg: n3_direct_avg.toFixed(1), events_avg: n3_events_avg.toFixed(1) }
      },
      conversion: { cr16_n1: cr16_n1.toFixed(1), cr46_n1: cr46_n1.toFixed(1), noshow: noshow.toFixed(1), cr_re: cr_re.toFixed(1), cr_ref: cr_ref.toFixed(1), cr16_n3: cr16_n3.toFixed(1) },
      results: { totalDeals, totalRev, forecastDeal, forecastRev },
      diagnostic: { star: buDiag.star.length, smart: buDiag.smart.length, effort: buDiag.effort.length }
    });
  }
}

// MiMi AI deep analysis from FM Analytics tab
async function askMiMiAnalytics() {
  const btn = document.getElementById('ana-mimi-btn');
  const respEl = document.getElementById('ana-mimi-response');
  const ctxEl = document.getElementById('ana-mimi-context');
  if (!btn || !respEl) return;

  const ctx = ctxEl ? ctxEl.textContent : '{}';
  btn.disabled = true;
  btn.innerHTML = '\u0110ang ph\u00E2n t\u00EDch...';
  respEl.style.display = '';
  respEl.innerHTML = '<span style="color:var(--gray-400)">MiMi \u0111ang nghi\u00EAn c\u1EE9u th\u1ECB tr\u01B0\u1EDDng v\u00E0 ph\u00E2n t\u00EDch d\u1EEF li\u1EC7u... (c\u00F3 th\u1EC3 m\u1EA5t 15-30 gi\u00E2y)</span>';

  try {
    const sysPrompt = buildMiMiSystemPrompt(ctx);
    const userMsg = `Ph\u00E2n t\u00EDch t\u1ED5ng quan h\u1EC7 th\u1ED1ng:\n1. \u0110\u00E1nh gi\u00E1 m\u1EE9c \u0111\u1ED9 "ch\u0103m ch\u1EC9" v\u00E0 "hi\u1EC7u qu\u1EA3" c\u1EE7a h\u1EC7 th\u1ED1ng\n2. Ch\u1EC9 ra ch\u1EC9 s\u1ED1 c\u00F4ng vi\u1EC7c n\u00E0o c\u1EA7n \u01B0u ti\u00EAn c\u1EA3i thi\u1EC7n nh\u1EA5t\n3. N\u1EBFu \u0111\u00E3 ch\u0103m ch\u1EC9 m\u00E0 k\u1EBFt qu\u1EA3 v\u1EABn th\u1EA5p, \u0111\u01B0a ra 3 chi\u1EBFn l\u01B0\u1EE3c l\u00E0m vi\u1EC7c th\u00F4ng minh h\u01A1n d\u1EF1a tr\u00EAn xu h\u01B0\u1EDBng EdTech hi\u1EC7n t\u1EA1i\n\nH\u00E3y k\u1EBFt h\u1EE3p d\u1EEF li\u1EC7u n\u1ED9i b\u1ED9 v\u1EDBi nghi\u00EAn c\u1EE9u t\u1EEB th\u1ECB tr\u01B0\u1EDDng gi\u00E1o d\u1EE5c c\u00F4ng ngh\u1EC7, xu h\u01B0\u1EDBng EdTech Vi\u1EC7t Nam/\u0110\u00F4ng Nam \u00C1 \u0111\u1EC3 \u0111\u01B0a ra chi\u1EBFn l\u01B0\u1EE3c c\u1EE5 th\u1EC3.`;

    const answer = await callMiMiDirect(sysPrompt, userMsg);
    respEl.innerHTML = mimiMarkdown(answer);
  } catch(e) {
    console.error('MiMi analytics error:', e);
    respEl.innerHTML = `<span style="color:var(--red)">L\u1ED7i k\u1EBFt n\u1ED1i MiMi: ${e.message}. Vui l\u00F2ng th\u1EED l\u1EA1i.</span>`;
  } finally {
    btn.disabled = false;
    btn.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg> H\u1ECFi MiMi ph\u00E2n t\u00EDch s\u00E2u';
  }
}
