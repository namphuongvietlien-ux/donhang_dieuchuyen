// API helpers
const GAS_EXEC_URL = 'https://script.google.com/macros/s/AKfycbwhqeAzzNrPTm1cH7KMmmj44btXb2OL835xxaItHByohT11sLDrdgfw7BrVlI5txqXonw/exec';

async function callJsonApi(urls, options, timeoutMs) {
  let lastError = null;
  const uniqueUrls = Array.isArray(urls) ? urls.filter(function(u, i, arr) { return u && arr.indexOf(u) === i; }) : [urls];
  for (const target of uniqueUrls) {
    const controller = new AbortController();
    const ms = timeoutMs || 120000;
    const timer = setTimeout(function() { controller.abort(); }, ms);
    const reqOptions = Object.assign({}, options || {}, { signal: controller.signal });
    try {
      const res = await fetch(target, reqOptions);
      clearTimeout(timer);
      const txt = await res.text();
      if (!res.ok) {
        lastError = new Error('HTTP ' + res.status + ': ' + txt.slice(0, 200));
        continue;
      }
      try {
        return JSON.parse(txt);
      } catch (e) {
        lastError = new Error('Invalid JSON response from ' + target + ': ' + txt.slice(0, 200));
        continue;
      }
    } catch (err) {
      clearTimeout(timer);
      if (err && err.name === 'AbortError') {
        lastError = new Error('Hết thời gian chờ (' + Math.round(ms / 1000) + 's). Kiểm tra deploy code.gs lên Google Apps Script.');
      } else {
        lastError = err;
      }
    }
  }
  throw lastError || new Error('Không thể kết nối tới máy chủ');
}

function buildGasGetUrl_(action, params, skipCacheBust) {
  const directUrl = new URL(GAS_EXEC_URL);
  directUrl.searchParams.set('action', action);
  if (!skipCacheBust) directUrl.searchParams.set('_ts', String(Date.now()));
  if (params) {
    Object.keys(params).forEach(k => {
      if (params[k] !== undefined && params[k] !== null) directUrl.searchParams.set(k, params[k]);
    });
  }
  return directUrl.toString();
}

function buildProxyGetUrl_(action, params, skipCacheBust) {
  const proxyUrl = new URL('/api/gas-proxy', location.origin);
  proxyUrl.searchParams.set('action', action);
  if (!skipCacheBust) proxyUrl.searchParams.set('_ts', String(Date.now()));
  if (params) {
    Object.keys(params).forEach(k => {
      if (params[k] !== undefined && params[k] !== null) proxyUrl.searchParams.set(k, params[k]);
    });
  }
  return proxyUrl.toString();
}

async function apiGet(action, params, options) {
  options = options || {};
  const timeoutMs = options.timeoutMs || 90000;
  // GET thẳng GAS ổn định hơn proxy (proxy từng 404/chậm làm hỏng mọi tab)
  if (options.directOnly) {
    return callJsonApi([buildGasGetUrl_(action, params, options.skipCacheBust)], { method: 'GET', headers: { 'Accept': 'application/json' } }, timeoutMs);
  }
  const urls = [buildGasGetUrl_(action, params, options.skipCacheBust)];
  if (options.allowProxyFallback !== false) urls.push(buildProxyGetUrl_(action, params, options.skipCacheBust));
  return callJsonApi(urls, { method: 'GET', headers: { 'Accept': 'application/json' } }, timeoutMs);
}

async function apiPost(action, payload, options) {
  options = options || {};
  const body = { action: action, payload: payload };
  const jsonBody = JSON.stringify(body);
  // text/plain: không preflight CORS; proxy dùng làm dự phòng
  const plainOptions = { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: jsonBody };
  const proxyOptions = { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: jsonBody };
  const timeoutMs = options.timeoutMs || 120000;

  if (options.directOnly) {
    return callJsonApi([GAS_EXEC_URL], plainOptions, timeoutMs);
  }

  // Ưu tiên POST thẳng GAS (tránh proxy treo). Proxy chỉ fallback.
  try {
    return await callJsonApi([GAS_EXEC_URL], plainOptions, timeoutMs);
  } catch (directErr) {
    const msg = String(directErr && directErr.message || directErr);
    console.warn('[donhang] direct POST failed, fallback proxy', msg);
    if (options.allowDirectFallback === false && options.allowProxyFallback === false) throw directErr;
    return callJsonApi(['/api/gas-proxy'], proxyOptions, Math.min(timeoutMs, 50000));
  }
}

function showLoginError(message) {
  alert(message || 'Đăng nhập thất bại.');
}

// --- App logic (extracted from original webapp) ---
var APP_BUILD = '2026-08-02-v16';
var shStockWarmState = { ready: false, warming: false, lastMs: 0, promise: null };
console.warn('[donhang] build', APP_BUILD);
(function() {
  var el = document.getElementById('app-build-tag');
  if (el) el.textContent = 'build: ' + APP_BUILD;
})();
var danhMucGoc = {}; var danhMucArr = []; var arrItems = []; var gStores = [];
var storeMap = {};
var catalogLoadState = { loading: false, ready: false, version: '' };
var CATALOG_CACHE_KEY = 'donhang_catalog_v2';
var CATALOG_CACHE_TS_KEY = 'donhang_catalog_ts_v2';
var CATALOG_CACHE_VERSION_KEY = 'donhang_catalog_version_v2';
var CATALOG_CACHE_TTL_MS = 30 * 60 * 1000;
var BOOTSTRAP_CACHE_KEY = 'donhang_bootstrap_v1';
var BOOTSTRAP_CACHE_TS_KEY = 'donhang_bootstrap_ts_v1';
var BOOTSTRAP_CACHE_TTL_MS = 60 * 60 * 1000;
function escapeHtml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function rebuildCatalogArray() {
  var seenCatalogKeys = {};
  danhMucArr = Object.values(danhMucGoc).filter(function(item) {
    if (!item) return false;
    var key = String(item.maHang || '').trim().toUpperCase() + '|' + String(item.maVach || '').trim().toUpperCase() + '|' + String(item.tenHang || '').trim().toUpperCase();
    if (!key || seenCatalogKeys[key]) return false;
    seenCatalogKeys[key] = true;
    return true;
  });
}

function applyCatalogData(res) {
  if (!res || !res.danhMuc) return;
  danhMucGoc = res.danhMuc;
  rebuildCatalogArray();
  catalogLoadState.ready = true;
  catalogLoadState.version = res.version || '';
}

function readCatalogFromLocalStorage(expectedVersion) {
  try {
    var ts = Number(localStorage.getItem(CATALOG_CACHE_TS_KEY) || '0');
    var version = localStorage.getItem(CATALOG_CACHE_VERSION_KEY) || '';
    if (!ts || (Date.now() - ts) > CATALOG_CACHE_TTL_MS) return null;
    if (expectedVersion && version && expectedVersion !== version) return null;
    var raw = localStorage.getItem(CATALOG_CACHE_KEY);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function saveCatalogToLocalStorage(res) {
  try {
    if (!res || !res.danhMuc) return;
    localStorage.setItem(CATALOG_CACHE_KEY, JSON.stringify(res));
    localStorage.setItem(CATALOG_CACHE_TS_KEY, String(Date.now()));
    localStorage.setItem(CATALOG_CACHE_VERSION_KEY, res.version || '');
  } catch (e) {}
}

function clearCatalogLocalStorage() {
  try {
    localStorage.removeItem(CATALOG_CACHE_KEY);
    localStorage.removeItem(CATALOG_CACHE_TS_KEY);
    localStorage.removeItem(CATALOG_CACHE_VERSION_KEY);
  } catch (e) {}
  catalogLoadState.ready = false;
}

function getDefaultExportStore() {
  for (var i = 0; i < gStores.length; i++) {
    var label = storeMap[gStores[i]] || gStores[i];
    if (String(label).indexOf('Q7') !== -1 || String(gStores[i]).indexOf('Q7') !== -1) return gStores[i];
  }
  return gStores[0] || 'Kho Địa điểm kinh doanh Q7';
}

function setCatalogStatus(text) {
  var scanInput = getEl('input-scan');
  if (!scanInput) return;
  scanInput.placeholder = text || 'Nhập để tìm kiếm hoặc quét mã...';
}

function renderStoreDropdowns() {
  var htmlStores = '';
  gStores.forEach(function(s) {
    var disp = storeMap[s] || s;
    htmlStores += '<option value="' + escapeHtml(s) + '">' + escapeHtml(disp) + '</option>';
  });
  var elX = getEl('select-kho-xuat'); if (elX) elX.innerHTML = htmlStores;
  var elN = getEl('select-kho-nhan'); if (elN) elN.innerHTML = htmlStores;
  var elQ = getEl('ql-kho-nhan'); if (elQ) elQ.innerHTML = '<option value="all">-- Tất cả --</option>' + htmlStores;
}

function renderAdminStoreDropdown() {
  var elA = getEl('adm-store');
  var roleEl = getEl('adm-role');
  if (!elA) return;
  var role = roleEl && roleEl.value ? roleEl.value : 'Chi nhánh';
  var html = role === 'Admin'
    ? '<option value="Tất cả">Tất cả (Admin)</option>'
    : '<option value="">-- Chọn kho quản lý --</option>';
  gStores.forEach(function(s) {
    var disp = storeMap[s] || s;
    html += '<option value="' + escapeHtml(s) + '">' + escapeHtml(disp) + '</option>';
  });
  elA.innerHTML = html;
  if (role === 'Admin') elA.value = 'Tất cả';
}

function loadCatalogInBackground(forceReload, expectedVersion) {
  if (catalogLoadState.loading) return Promise.resolve();
  var cached = !forceReload ? readCatalogFromLocalStorage(expectedVersion) : null;
  if (cached && cached.danhMuc) {
    applyCatalogData(cached);
    setCatalogStatus('Quét mã vạch, gõ mã, từ khóa tên hoặc 6 số cuối vạch:');
    return Promise.resolve(cached);
  }

  catalogLoadState.loading = true;
  setCatalogStatus('Đang tải danh mục hàng ở nền...');
  return apiGet('getCatalogData').then(function(res) {
    catalogLoadState.loading = false;
    if (!res || !res.success) {
      setCatalogStatus('Danh mục chưa tải xong - thử lại sau vài giây');
      return res;
    }
    applyCatalogData(res);
    saveCatalogToLocalStorage(res);
    setCatalogStatus('Quét mã vạch, gõ mã, từ khóa tên hoặc 6 số cuối vạch:');
    return res;
  }).catch(function(err) {
    catalogLoadState.loading = false;
    setCatalogStatus('Lỗi tải danh mục: ' + err.message);
    throw err;
  });
}

function applyBootstrapData(res) {
  gStores = res.stores || [];
  storeMap = res.storeMap || {};
  renderStoreDropdowns();
  renderAdminStoreDropdown();
  var nav = getEl('nav-tab-admin');
  if (nav) nav.style.display = sessionUser.role === 'Admin' ? 'block' : 'none';
  updateDashboardHero();
  applyQuyenKho();
}

function readBootstrapFromLocalStorage() {
  try {
    var ts = Number(localStorage.getItem(BOOTSTRAP_CACHE_TS_KEY) || '0');
    if (!ts || (Date.now() - ts) > BOOTSTRAP_CACHE_TTL_MS) return null;
    var raw = localStorage.getItem(BOOTSTRAP_CACHE_KEY);
    if (!raw) return null;
    var parsed = JSON.parse(raw);
    return parsed && parsed.success ? parsed : null;
  } catch (e) {
    return null;
  }
}

function saveBootstrapToLocalStorage(res) {
  try {
    if (!res || !res.success) return;
    localStorage.setItem(BOOTSTRAP_CACHE_KEY, JSON.stringify(res));
    localStorage.setItem(BOOTSTRAP_CACHE_TS_KEY, String(Date.now()));
  } catch (e) {}
}

var phieuData = []; var editRows = []; var currentLoadedRows = []; var currentPhieuObj = null; var currentConfirmPhieuObj = null;
var sessionUser = { user: "", role: "", store: "" };
var deepLinkOrder = new URLSearchParams(location.search).get("soPhieu");
var deepLinkTab = new URLSearchParams(location.search).get("tab");
var INVENTORY_APP_URL = "https://my-inventory-app.vercel.app";

function getEl(id) { return document.getElementById(id); }
function safeText(id, value) { var el = getEl(id); if (el) el.innerText = value; }
function safeDisplay(id, display) { var el = getEl(id); if (el) el.style.display = display; }
function safeValue(id, value) { var el = getEl(id); if (el) el.value = value; }

function getDeepLinkParams() {
  var params = new URLSearchParams(location.search);
  return {
    order: params.get("soPhieu"),
    tab: params.get("tab"),
    public: params.get("public") === "1" || params.get("view") === "public"
  };
}

function openInventoryApp() {
  var targetUrl = new URL(INVENTORY_APP_URL);
  if (sessionUser && sessionUser.user) targetUrl.searchParams.set("user", sessionUser.user);
  var selectedStore = "";
  var storeSelect = document.getElementById("select-kho-nhan");
  if (storeSelect && storeSelect.value) selectedStore = storeSelect.value;
  if (!selectedStore && sessionUser && sessionUser.store) selectedStore = sessionUser.store;
  if (selectedStore) targetUrl.searchParams.set("store", selectedStore);
  targetUrl.searchParams.set("from", "donhang");
  window.open(targetUrl.toString(), "_blank", "noopener,noreferrer");
}

window.onload = function() {
  var loadingOverlay = getEl("loading-overlay"); if (loadingOverlay) loadingOverlay.style.display = "none";
  var pass = getEl("lg-pass"); if(pass) pass.addEventListener("keypress", function(e){ if(e.key==="Enter") doLogin(); });
  document.addEventListener("click", function(){ closeUserMenu(); });
  hidePasswordSection();
  var params = new URLSearchParams(location.search);
  var isPublicView = params.get("public") === "1" || params.get("view") === "public";
  var hasDeepLink = !!(params.get("soPhieu") || params.get("tab") || isPublicView);
  if (hasDeepLink) {
    sessionUser = { user: "", role: "Guest", store: "" };
    safeText("lbl-username", "Guest");
    safeDisplay("login-screen", "none");
    safeDisplay("main-container", "block");
    updateDashboardHero();
    initSystemData();
  }
};

// ================= ĐĂNG NHẬP =================
function doLogin() {
  var uInput = getEl("lg-user"); var pInput = getEl("lg-pass");
  if (!uInput || !pInput) return alert("Không tìm thấy form đăng nhập.");
  var u = uInput.value.trim();
  var p = pInput.value.trim();
  if(!u || !p) return alert("Vui lòng nhập đủ thông tin!");
  showLoad("Đang xác thực...");
  apiPost('loginUser', { username: u, password: p }).then(function(res) {
    hideLoad();
    if (!res || typeof res !== 'object') {
      alert("Không thể xác thực. Vui lòng cập nhật và triển khai lại Google Apps Script.");
      return;
    }
    if(res.success) {
      sessionUser = { user: res.username, role: res.role, store: res.store };
      safeText("lbl-username", sessionUser.user + " (" + sessionUser.role + ")");
      updateDashboardHero();
      safeDisplay("login-screen", "none");
      safeDisplay("main-container", "block");
      initSystemData();
    } else {
      showLoginError("❌ " + (res.msg || res.error || "Không thể đăng nhập."));
    }
  }).catch(function(err){ hideLoad(); showLoginError('Lỗi kết nối: '+err.message); });
}

function initSystemData() {
  var cachedBootstrap = readBootstrapFromLocalStorage();
  if (cachedBootstrap) {
    applyBootstrapData(cachedBootstrap);
    loadCatalogInBackground(false, cachedBootstrap.catalogVersion || '');
  }

  var qlNgay = getEl("ql-ngay"); if (qlNgay && !qlNgay.value) qlNgay.valueAsDate = new Date();
  var qlNgayFast = getEl("ql-ngay-fast"); if (qlNgayFast && !qlNgayFast.value) qlNgayFast.value = "today";
  var confirmNgay = getEl("confirm-ngay"); if (confirmNgay && !confirmNgay.value) confirmNgay.valueAsDate = new Date();
  var confirmNgayFast = getEl("confirm-ngay-fast"); if (confirmNgayFast && !confirmNgayFast.value) confirmNgayFast.value = "today";
  if (!cachedBootstrap) showLoad("Đang tải hệ thống...");

  // Proxy có thể 404 — luôn cho phép fallback GET thẳng GAS
  apiGet('getBootstrapData', null, { allowDirectFallback: true }).then(function(res) {
    hideLoad();
    if (!res || !res.success) {
      if (!cachedBootstrap) alert("Lỗi tải data: " + ((res && (res.error || res.msg)) || res));
      return;
    }
    applyBootstrapData(res);
    saveBootstrapToLocalStorage(res);
    loadCatalogInBackground(false, res.catalogVersion || '');
    openDeepLinkedOrder();
  }).catch(function(err) {
    hideLoad();
    if (!cachedBootstrap) alert('Lỗi: ' + err.message);
  });

  if (cachedBootstrap) openDeepLinkedOrder();
}

function updateDashboardHero() {
  safeText('hero-role', sessionUser.role || 'Guest');
  safeText('hero-store', sessionUser.store && sessionUser.store !== 'Tất cả' ? (storeMap[sessionUser.store] || sessionUser.store) : 'Tất cả');
}

function toggleUserMenu(event) {
  if (event) event.stopPropagation();
  var menu = document.getElementById('user-menu');
  var badge = document.getElementById('user-info-badge');
  if (!menu || !badge) return;
  var isOpen = menu.style.display === 'block';
  menu.style.display = isOpen ? 'none' : 'block';
  badge.classList.toggle('active', !isOpen);
}

function closeUserMenu() {
  var menu = document.getElementById('user-menu');
  var badge = document.getElementById('user-info-badge');
  if (menu) menu.style.display = 'none';
  if (badge) badge.classList.remove('active');
}

function hidePasswordSection() {
  var card = document.getElementById('password-card');
  if (card) card.style.display = 'none';
}

function showPasswordSection() {
  closeUserMenu();
  var card = document.getElementById('password-card');
  if (card) {
    card.style.display = 'block';
    card.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }
  var input = document.getElementById('pw-current');
  if (input) input.focus();
}

function logoutUser() {
  closeUserMenu();
  hidePasswordSection();
  sessionUser = { user: '', role: '', store: '' };
  var nav = document.getElementById('nav-tab-admin');
  if (nav) nav.style.display = 'none';
  safeText('lbl-username', 'Guest');
  safeText('hero-role', 'Guest');
  safeText('hero-store', '-');
  safeDisplay('login-screen', 'flex');
  safeDisplay('main-container', 'none');
  safeValue('lg-user', '');
  safeValue('lg-pass', '');
}

function activateTab(tabId) {
  document.querySelectorAll('.nav-tab').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
  var targetTab = getEl(tabId);
  if (!targetTab) return;
  var nav = Array.from(document.querySelectorAll('.nav-tab')).find(function(el) { var onclick = el.getAttribute('onclick') || ''; return onclick.indexOf("'" + tabId + "'") !== -1; });
  if (nav) nav.classList.add('active');
  targetTab.classList.add('active');
}

function switchTab(tabId) {
  hidePasswordSection();
  activateTab(tabId);
  if(tabId === 'tab-quan-ly') ql_loadPhieu();
  if(tabId === 'tab-xac-nhan') confirm_loadPhieu();
  if(tabId === 'tab-soan-hang') sh_taiDanhSachDon();
  if(tabId === 'tab-dashboard') loadDashboardSummary();
  if(tabId === 'tab-admin') loadDSUser();
}

function formatDateInputValue(date) {
  var yyyy = date.getFullYear();
  var mm = String(date.getMonth() + 1).padStart(2, '0');
  var dd = String(date.getDate()).padStart(2, '0');
  return yyyy + '-' + mm + '-' + dd;
}

function dashboard_onTimelineChange() {
  var timelineEl = document.getElementById('dashboard-timeline');
  var fromWrap = document.getElementById('dashboard-from-wrap');
  var toWrap = document.getElementById('dashboard-to-wrap');
  var fromEl = document.getElementById('dashboard-from');
  var toEl = document.getElementById('dashboard-to');
  var timeline = timelineEl && timelineEl.value ? timelineEl.value : '2days';
  var isCustom = timeline === 'custom';
  if (fromWrap) fromWrap.style.display = isCustom ? 'block' : 'none';
  if (toWrap) toWrap.style.display = isCustom ? 'block' : 'none';
  if (isCustom && fromEl && toEl && !fromEl.value && !toEl.value) {
    var today = new Date();
    var start = new Date(today);
    start.setDate(today.getDate() - 1);
    fromEl.value = formatDateInputValue(start);
    toEl.value = formatDateInputValue(today);
  }
  loadDashboardSummary();
}

function getDashboardTimelineParams() {
  var timelineEl = document.getElementById('dashboard-timeline');
  var fromEl = document.getElementById('dashboard-from');
  var toEl = document.getElementById('dashboard-to');
  return {
    timeline: timelineEl && timelineEl.value ? timelineEl.value : '2days',
    fromDate: fromEl && fromEl.value ? fromEl.value : '',
    toDate: toEl && toEl.value ? toEl.value : ''
  };
}

function getDashboardTimelineLabel(params) {
  if (!params) return '2 ngày gần nhất';
  if (params.timeline === 'today') return 'Hôm nay';
  if (params.timeline === '7days') return '7 ngày gần nhất';
  if (params.timeline === '30days') return '30 ngày gần nhất';
  if (params.timeline === 'all') return 'Tất cả thời gian';
  if (params.timeline === 'custom') {
    if (params.fromDate && params.toDate) return 'Từ ' + params.fromDate + ' đến ' + params.toDate;
    if (params.fromDate) return 'Từ ' + params.fromDate;
    if (params.toDate) return 'Đến ' + params.toDate;
    return 'Khoảng thời gian tùy chỉnh';
  }
  return '2 ngày gần nhất';
}

function loadDashboardSummary() {
  if (!sessionUser || !sessionUser.user) return;
  var grid = document.getElementById('dashboard-summary-grid');
  var recent = document.getElementById('dashboard-recent-orders');
  var timelineLabel = document.getElementById('dashboard-timeline-label');
  var timelineParams = getDashboardTimelineParams();
  var timeline = timelineParams.timeline;
  if (!grid || !recent) return;
  if (timelineLabel) timelineLabel.innerText = 'Đang xem: ' + getDashboardTimelineLabel(timelineParams);
  showLoad('Đang tải tổng quan...');
  apiGet('getDashboardSummary', { userRole: sessionUser.role || '', userStore: sessionUser.store || '', timeline: timeline, fromDate: timelineParams.fromDate, toDate: timelineParams.toDate }).then(function(res) {
    hideLoad();
    if (!res || !res.success || !res.data) return;
    var data = res.data;
    grid.innerHTML = [
      '<div class="card" style="margin:0; padding:14px; background:#eff6ff; border:1px solid #bfdbfe;"><div style="font-size:12px; color:#1d4ed8; font-weight:700; text-transform:uppercase;">Tổng đơn</div><div style="font-size:24px; font-weight:800; color:#1e3a8a;">' + data.totalOrders + '</div></div>',
      '<div class="card" style="margin:0; padding:14px; background:#fefce8; border:1px solid #fde68a;"><div style="font-size:12px; color:#92400e; font-weight:700; text-transform:uppercase;">Đang chờ</div><div style="font-size:24px; font-weight:800; color:#92400e;">' + data.pendingOrders + '</div></div>',
      '<div class="card" style="margin:0; padding:14px; background:#f0fdf4; border:1px solid #bbf7d0;"><div style="font-size:12px; color:#166534; font-weight:700; text-transform:uppercase;">Đã xử lý</div><div style="font-size:24px; font-weight:800; color:#166534;">' + data.processedOrders + '</div></div>',
      '<div class="card" style="margin:0; padding:14px; background:#fef2f2; border:1px solid #fecaca;"><div style="font-size:12px; color:#b91c1c; font-weight:700; text-transform:uppercase;">Đã hủy</div><div style="font-size:24px; font-weight:800; color:#b91c1c;">' + data.canceledOrders + '</div></div>'
    ].join('');

    if (!data.recentOrders || !data.recentOrders.length) {
      recent.innerHTML = '<div style="padding:12px; background:#f8fafc; border:1px dashed #cbd5e1; border-radius:12px; color:#64748b;">Chưa có đơn hàng nào trong khoảng thời gian đã chọn.</div>';
      return;
    }

    var rows = data.recentOrders.map(function(order) {
      return '<tr><td><b>' + order.soPhieu + '</b></td><td>' + (order.khoXuat || '-') + '</td><td>' + (order.khoNhan || '-') + '</td><td><span style="display:inline-block; padding:4px 8px; border-radius:999px; background:#eff6ff; color:#1d4ed8; font-size:12px; font-weight:700;">' + order.status + '</span></td><td>' + (order.thoiGian || '-') + '</td></tr>';
    }).join('');

    recent.innerHTML = '<div style="overflow-x:auto;"><table style="width:100%; border-collapse:collapse;"><thead><tr><th style="text-align:left; padding:8px; border-bottom:1px solid #e2e8f0;">Số phiếu</th><th style="text-align:left; padding:8px; border-bottom:1px solid #e2e8f0;">Kho xuất</th><th style="text-align:left; padding:8px; border-bottom:1px solid #e2e8f0;">Kho nhận</th><th style="text-align:left; padding:8px; border-bottom:1px solid #e2e8f0;">Trạng thái</th><th style="text-align:left; padding:8px; border-bottom:1px solid #e2e8f0;">Cập nhật</th></tr></thead><tbody>' + rows + '</tbody></table></div>';
  }).catch(function(err){ hideLoad(); console.error(err); });
}

// ================= PHÂN QUYỀN KHO =================
function applyQuyenKho() {
  var loaiDon = document.querySelector('input[name="loaiPhieu"]:checked').value;
  var khoXuatEl = document.getElementById("select-kho-xuat");
  var khoNhanEl = document.getElementById("select-kho-nhan");

  if (loaiDon === "DonHang") {
    if(khoXuatEl) { khoXuatEl.value = getDefaultExportStore(); khoXuatEl.setAttribute("disabled", "true"); }
    if (sessionUser.role !== "Admin") {
      if(khoNhanEl) { khoNhanEl.value = sessionUser.store; khoNhanEl.setAttribute("disabled", "true"); }
    } else { if(khoNhanEl) khoNhanEl.removeAttribute("disabled"); }
  } else if (loaiDon === "DieuChuyen") {
    if(khoXuatEl) khoXuatEl.removeAttribute("disabled");
    if(khoNhanEl) khoNhanEl.removeAttribute("disabled");
    if (sessionUser.role !== "Admin" && khoXuatEl) khoXuatEl.value = sessionUser.store;
  }
}

// ================= TÌM KIẾM & TẠO PHIẾU =================
function normalizeSearchText(value) {
  return String(value || "").toLowerCase().normalize('NFKD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]+/g, ' ').trim();
}

function getSearchScore(item, query) {
  var score = 0;
  var fields = [item.maHang, item.maVach, item.tenHang];
  var q = normalizeSearchText(query);
  if (!q) return 0;
  fields.forEach(function(field) {
    var norm = normalizeSearchText(field);
    if (!norm) return;
    if (norm === q) score += 1000;
    else if (norm.indexOf(q) === 0) score += 600;
    else if (norm.indexOf(q) !== -1) score += 250;
    else {
      var parts = q.split(/\s+/).filter(Boolean);
      var matchedParts = parts.filter(function(part) { return norm.indexOf(part) !== -1; }).length;
      if (matchedParts > 0) score += matchedParts * 80;
    }
    if (field === item.maVach && norm.length >= 6 && q.length >= 6 && norm.slice(-6) === q.slice(-6)) score += 180;
  });
  return score;
}

function handleSearchInput(e) {
  if (!catalogLoadState.ready) {
    var box = document.getElementById("suggest-box");
    if (box) {
      box.innerHTML = '<div class="suggest-empty">Danh mục hàng đang tải. Vui lòng đợi vài giây...</div>';
      box.style.display = "block";
    }
    return;
  }
  var inputEl = document.getElementById("input-scan");
  var val = inputEl.value.trim();
  var box = document.getElementById("suggest-box");
  box.style.width = (inputEl.offsetWidth) + "px"; box.style.left = (inputEl.offsetLeft) + "px"; box.style.top = (inputEl.offsetTop + inputEl.offsetHeight) + "px";

  if (val.length < 1) { box.style.display = "none"; return; }
  var kw = val.toUpperCase();

  if (e.key === "Enter") {
    var exactMatch = danhMucGoc[kw];
    if (exactMatch) chonSanPham(exactMatch);
    else {
      var matched = filterProducts(kw);
      if (matched.length > 0) chonSanPham(matched[0]);
      else { arrItems.unshift({ maHang: "LỖI MÃ", maVach: val, tenHang: "❌ Không tồn tại", dvt: "Lỗi", sl: "1" }); renderTable(); }
    }
    inputEl.value = ""; box.style.display = "none"; return;
  }

  var results = filterProducts(kw);
  if (results.length === 0) { box.innerHTML = '<div style="padding:10px; color:#d93025; text-align:center; font-weight:600;">Không tìm thấy sản phẩm phù hợp.</div>'; box.style.display = "block"; return; }

  var html = "";
  results.slice(0, 10).forEach(function(item) {
    var itemStr = encodeURIComponent(JSON.stringify(item));
    html += '<div class="suggest-item" onclick="chonSanPhamFromSuggest(\'' + itemStr + '\')"><div class="sg-title">' + item.tenHang + '</div><div class="sg-desc"><span style="color:#1a73e8; font-weight:700;">Mã hàng: ' + item.maHang + '</span> · Mã vạch: ' + item.maVach + ' · ĐVT: ' + (item.dvt || 'Cái') + '</div></div>';
  });
  box.innerHTML = html; box.style.display = "block";
}

function filterProducts(kw) {
  var query = normalizeSearchText(kw);
  if (!query) return [];
  var scored = danhMucArr.map(function(it) {
    return { item: it, score: getSearchScore(it, query) };
  }).filter(function(entry) {
    return entry.score > 0;
  }).sort(function(a, b) {
    return b.score - a.score || String(a.item.tenHang || "").localeCompare(String(b.item.tenHang || ""));
  });
  var seenResults = {};
  return scored.filter(function(entry) {
    var item = entry.item;
    var key = String(item.maHang || "").trim().toUpperCase() + '|' + String(item.maVach || "").trim().toUpperCase() + '|' + String(item.tenHang || "").trim().toUpperCase();
    if (!key || seenResults[key]) return false;
    seenResults[key] = true;
    return true;
  }).map(function(entry) { return entry.item; });
}

function positionSuggestionBox(box, inputEl) {
  if (!box || !inputEl) return;
  box.style.width = (inputEl.offsetWidth || 280) + "px";
  box.style.left = (inputEl.offsetLeft || 0) + "px";
  box.style.top = (inputEl.offsetTop + inputEl.offsetHeight + 6) + "px";
}

var qlSelectedSuggestedItem = null;

function ql_isItemMatchedWithCode(item, code) {
  if (!item) return false;
  var query = String(code || "").trim().toUpperCase();
  if (!query) return false;
  var maHang = String(item.maHang || "").trim().toUpperCase();
  var maVach = String(item.maVach || "").trim().toUpperCase();
  return query === maHang || query === maVach;
}

function ql_handleAddCodeInput(e) {
  var inputEl = document.getElementById("ql-add-code");
  var box = document.getElementById("ql-suggest-box");
  if (!inputEl || !box) return;
  var val = inputEl.value.trim();
  qlSelectedSuggestedItem = null;
  positionSuggestionBox(box, inputEl);

  if (val.length < 1) { box.style.display = "none"; return; }
  if (e && e.key === "Enter") {
    e.preventDefault();
    box.style.display = "none";
    ql_themMaHang();
    return;
  }

  var results = filterProducts(val);
  if (results.length === 0) {
    box.innerHTML = '<div class="suggest-empty">Không tìm thấy sản phẩm phù hợp.</div>';
    box.style.display = "block";
    return;
  }

  var html = '';
  results.slice(0, 8).forEach(function(item) {
    var itemStr = encodeURIComponent(JSON.stringify(item));
    html += '<div class="suggest-item" onclick="ql_pickSuggestedItem(\'' + itemStr + '\')"><div class="sg-title">' + item.tenHang + '</div><div class="sg-desc"><span style="color:#2563eb; font-weight:700;">Mã hàng: ' + item.maHang + '</span> · Mã vạch: ' + item.maVach + ' · ĐVT: ' + (item.dvt || 'Cái') + '</div></div>';
  });
  box.innerHTML = html;
  box.style.display = "block";
}

function ql_handleAddQtyKeydown(e) {
  if (!e || e.key !== "Enter") return;
  e.preventDefault();
  ql_themMaHang();
}

function ql_pickSuggestedItem(itemStr) {
  var item = JSON.parse(decodeURIComponent(itemStr));
  var input = document.getElementById("ql-add-code");
  var box = document.getElementById("ql-suggest-box");
  var qtyInput = document.getElementById("ql-add-qty");
  qlSelectedSuggestedItem = item;
  if (input) input.value = item.maHang || item.maVach || "";
  if (box) box.style.display = "none";
  if (qtyInput) {
    qtyInput.focus();
    qtyInput.select();
  }
}

function chonSanPhamFromSuggest(itemStr) {
  chonSanPham(JSON.parse(decodeURIComponent(itemStr)));
  document.getElementById("input-scan").value = ""; document.getElementById("suggest-box").style.display = "none"; document.getElementById("input-scan").focus();
}

function chonSanPham(it) {
  var existingIndex = arrItems.findIndex(x => x.maVach === it.maVach && x.maHang !== "LỖI MÃ");
  if(existingIndex !== -1) { arrItems[existingIndex].sl = Number(arrItems[existingIndex].sl) + 1; arrItems[existingIndex].highlight = true; }
  else { arrItems.unshift({ maHang: it.maHang, maVach: it.maVach, tenHang: it.tenHang, dvt: it.dvt, sl: "1", highlight: true }); }
  renderTable();
}

document.addEventListener("click", function(event) {
  var box = document.getElementById("suggest-box");
  var input = document.getElementById("input-scan");
  var qlBox = document.getElementById("ql-suggest-box");
  var qlInput = document.getElementById("ql-add-code");
  if (box && event.target !== box && event.target !== input && !box.contains(event.target)) box.style.display = "none";
  if (qlBox && event.target !== qlBox && event.target !== qlInput && !qlBox.contains(event.target)) qlBox.style.display = "none";
});

function thayDoiSoLuong(index, delta) { var currentSl = Number(arrItems[index].sl) || 0; var newSl = currentSl + delta; if (newSl > 0) { arrItems[index].sl = newSl; renderTable(); } }

function renderTable() {
  var tbody = document.getElementById("tbody-items"); tbody.innerHTML = ""; var tongSl = 0;
  arrItems.forEach((it, i) => {
    var isErr = (it.maHang === "LỖI MÃ" || isNaN(Number(it.sl))); tongSl += (Number(it.sl) || 0);
    var trClass = isErr ? 'row-error' : (it.highlight ? 'scan-highlight' : ''); it.highlight = false;
    tbody.insertAdjacentHTML('beforeend', '<tr class="' + trClass + '"><td>' + (arrItems.length - i) + '</td><td><b>Mã vạch: ' + it.maVach + '</b><br><small style="color:gray;">Mã hàng hóa: ' + it.maHang + '</small></td><td style="font-weight:500;">' + it.tenHang + '</td><td>' + it.dvt + '</td><td><div class="qty-control"><button class="qty-btn" onclick="thayDoiSoLuong(' + i + ', -1)">-</button><input type="number" class="qty-input" value="' + it.sl + '" onchange="arrItems[' + i + '].sl=this.value; renderTable();"><button class="qty-btn" onclick="thayDoiSoLuong(' + i + ', 1)">+</button></div></td><td style="text-align:center;"><button style="color:#d93025; border:none; background:none; font-weight:bold; cursor:pointer; font-size:18px;" onclick="arrItems.splice(' + i + ',1); renderTable();">×</button></td></tr>');
  });
  document.getElementById("lbl-tong-sl").innerText = tongSl;
  if (arrItems.length === 0) {
    tbody.insertAdjacentHTML('beforeend', '<tr><td colspan="6" style="text-align:center; color:#64748b; padding:24px;">Chưa có mặt hàng nào trong phiếu. Hãy tìm kiếm và thêm sản phẩm.</td></tr>');
  }
}

function quickAddSuggestedItem(item) {
  chonSanPham(item);
}

function submitPhieuMoi() {
  if(arrItems.length === 0) return alert("Chưa có hàng!");
  if (!catalogLoadState.ready) return alert("Danh mục hàng đang tải. Vui lòng đợi vài giây rồi thử lại.");
  showLoad("Đang tạo đơn...");
  var lPhieu = document.querySelector('input[name="loaiPhieu"]:checked').value;
  var khoXuat = document.getElementById("select-kho-xuat").value;
  var khoNhan = document.getElementById("select-kho-nhan").value;
  var itemCount = arrItems.length;

  apiPost('luuPhieuTuWebApp', { loaiPhieu: lPhieu, khoXuat: khoXuat, khoNhan: khoNhan, items: arrItems }).then(function(res) {
    hideLoad();
    if(res.coLoi) { alert("⚠️ Có mã lỗi. Sửa trong tab Quản lý!"); arrItems = []; renderTable(); }
    else {
       currentPhieuObj = { soPhieu: res.soPhieu, khoXuat: khoXuat, khoNhan: khoNhan };
       document.getElementById("modal-sophieu").innerText = res.soPhieu; document.getElementById("modal-action").style.display = "flex";
       arrItems = []; renderTable();
       if (document.getElementById("input-scan")) document.getElementById("input-scan").focus();
       if (res.soPhieu && !res.coLoi) {
  apiPost('postProcessNewOrder', {
           soPhieu: res.soPhieu,
           khoXuat: res.khoXuat || khoXuat,
           khoNhan: res.khoNhan || khoNhan,
           itemCount: res.itemCount || itemCount
         }, { allowDirectFallback: false }).catch(function() {});
       }
    }
  }).catch(function(err){ hideLoad(); alert('Lỗi: '+err.message); });
}

function actionPrintNew() { executePrintWeb(currentPhieuObj.soPhieu, currentPhieuObj.khoXuat, currentPhieuObj.khoNhan, arrItems); }
function actionExportNew() { executeExportExcel(currentPhieuObj.soPhieu, currentPhieuObj.khoXuat, currentPhieuObj.khoNhan, arrItems); }
function actionCloseModal() { document.getElementById("modal-action").style.display = "none"; arrItems = []; renderTable(); }

// ================= IN WEB (IFRAME ẨN) & XUẤT EXCEL =================
function executePrintWeb(soPhieu, khoXuat, khoNhan, itemsArray) {
  var styleStr = 'body{font-family: Arial, sans-serif; padding:20px; font-size:12px;} table{width:100%; border-collapse:collapse; margin-top:12px;} th,td{border:1px solid #000; padding:7px; text-align:left; vertical-align:top;} th{background:#f0f0f0;} .title{font-size:16px; font-weight:bold; margin-bottom:8px;} .meta{margin-bottom:8px;} .code-cell{font-size:16px; font-weight:700;} .qty-cell{font-size:16px; font-weight:700; text-align:center;} .note-cell{width:72px;} @media print { @page { margin: 10mm; } }';

  var htmlStr = '<div class="title">Số: ' + soPhieu + '</div><div class="meta"><b>Kho xuất:</b> ' + khoXuat + '<br><b>Kho nhận:</b> ' + khoNhan + '</div>';
  htmlStr += '<table><thead><tr><th>STT</th><th>Mã</th><th>Tên hàng</th><th>ĐVT</th><th>Số lượng</th><th class="note-cell"></th></tr></thead><tbody>';
  var stt = 1; itemsArray.forEach(it => { if(Number(it.sl) > 0) { htmlStr += '<tr><td>'+(stt++)+'</td><td class="code-cell">'+ ((it.maVach || '') + (it.maHang ? ' / ' + it.maHang : '')) +'</td><td>'+it.tenHang+'</td><td>'+it.dvt+'</td><td class="qty-cell">'+it.sl+'</td><td class="note-cell"></td></tr>'; }});
  htmlStr += '</tbody></table><div style="display:flex; justify-content:space-between; margin-top:40px; text-align:center;"><div><b>Người lập phiếu</b><br><br><br>Ký ghi rõ họ tên</div><div><b>Người nhận</b><br><br><br>Ký ghi rõ họ tên</div></div>';

  var iframe = document.createElement('iframe');
  iframe.style.display = 'none';
  document.body.appendChild(iframe);

  var iframeDoc = iframe.contentWindow.document;
  iframeDoc.title = '';
  var styleEl = iframeDoc.createElement('style'); styleEl.innerHTML = styleStr; iframeDoc.head.appendChild(styleEl);
  iframeDoc.body.innerHTML = htmlStr;

  setTimeout(function() {
    iframe.contentWindow.focus(); iframe.contentWindow.print();
    setTimeout(function() { document.body.removeChild(iframe); }, 1500);
  }, 500);
}

function executeExportExcel(soPhieu, khoXuat, khoNhan, itemsArray) {
  showLoad("Đang xử lý File Excel...");
  apiPost('taoFileExcelVaLayLink', { soPhieu: soPhieu, khoXuat: khoXuat, khoNhan: khoNhan, items: itemsArray }).then(function(res) { hideLoad(); if(res.success) window.open(res.url, '_blank'); else alert('Lỗi xuất: '+(res.error||JSON.stringify(res))); }).catch(function(err){ hideLoad(); alert('Lỗi: '+err.message); });
}

// ================= QUẢN LÝ PHIẾU =================
function getNgayFilterParam(inputId, fastSelectId) {
  var fastEl = document.getElementById(fastSelectId);
  var fast = fastEl && fastEl.value ? fastEl.value.trim() : '';
  if (fast === '7days' || fast === 'all') return fast;
  if (fast === 'today' || fast === 'yesterday') return fast;
  var inputEl = document.getElementById(inputId);
  return inputEl && inputEl.value ? inputEl.value : '';
}

function resetQuickDateFilter(fastSelectId) {
  var fastEl = document.getElementById(fastSelectId);
  if (fastEl) fastEl.value = '';
}

function setQuickDateFilter(value, targetId, fastSelectId) {
  var inputEl = document.getElementById(targetId);
  if (!inputEl) return;
  if (!value || value === 'all') {
    inputEl.value = '';
    return;
  }
  if (value === '7days') {
    inputEl.value = '';
    return;
  }
  var today = new Date();
  if (value === 'today') {
    inputEl.value = formatDateInputValue(today);
  } else if (value === 'yesterday') {
    var y = new Date(today);
    y.setDate(today.getDate() - 1);
    inputEl.value = formatDateInputValue(y);
  }
}

function onManualNgayFilterChange(inputId, fastSelectId, reloadFn) {
  resetQuickDateFilter(fastSelectId);
  if (reloadFn === 'confirm') confirm_loadPhieu();
  else ql_loadPhieu();
}

function resetManagementViewAfterSave() {
  var selectEl = document.getElementById("ql-phieu");
  if (selectEl) selectEl.value = "";
  var viewEl = document.getElementById("ql-view-phieu");
  if (viewEl) viewEl.style.display = "none";
  currentPhieuObj = null;
  editRows = [];
  ql_loadPhieu();
}

function resetConfirmViewAfterSave() {
  var selectEl = document.getElementById("confirm-phieu");
  if (selectEl) selectEl.value = "";
  var viewEl = document.getElementById("confirm-view");
  if (viewEl) viewEl.style.display = "none";
  currentConfirmPhieuObj = null;
  confirm_loadPhieu();
}

function unwrapListResponse_(res) {
  // Hỗ trợ cả mảng thuần (GAS cũ/mới) và object { data: [] }
  if (Array.isArray(res)) {
    return {
      rows: res,
      meta: {
        _debugTotalMs: res._debugTotalMs,
        _debugScanned: res._debugScanned,
        _debugRun: res._debugRun,
        _debugStock: res._debugStock
      }
    };
  }
  if (res && Array.isArray(res.data)) return { rows: res.data, meta: res };
  if (res && res.success === false) return { rows: [], meta: res };
  return { rows: [], meta: res || {} };
}

function ql_loadPhieu(selectedSoPhieu) {
  var selectEl = document.getElementById("ql-phieu");
  if (!selectEl) return;
  selectEl.innerHTML = '<option value="">⏳ Đang tải...</option>';
  var ngay = getNgayFilterParam('ql-ngay', 'ql-ngay-fast');
  var t0 = Date.now();
  // #region agent log
  fetch('http://127.0.0.1:7480/ingest/48e8fdfc-ebb8-4d81-9aee-1659862ac812',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'4a6e3c'},body:JSON.stringify({sessionId:'4a6e3c',location:'app.js:ql_loadPhieu',message:'ql list start',data:{ngay:ngay,build:APP_BUILD},timestamp:Date.now(),hypothesisId:'QL-A',runId:'ql-fast-v1'})}).catch(function(){});
  // #endregion
  apiGet('layDanhSachPhieuTheoFilter', { khoNhan: document.getElementById("ql-kho-nhan").value, ngay: ngay, userRole: sessionUser.role, userStore: sessionUser.store }, { directOnly: true, timeoutMs: 45000 }).then(function(res) {
    var parsed = unwrapListResponse_(res);
    var rows = parsed.rows;
    phieuData = rows;
    var countMoi = 0; var countDone = 0; var countCancel = 0;
    var html = '<option value="">-- Chọn Đơn ('+rows.length+') --</option>';
    rows.forEach(function(r) {
      if(r.trangThai === "Mới") countMoi++; else if(r.trangThai === "Đã hủy") countCancel++; else countDone++;
      var shortName = storeMap[r.khoNhan] || storeMap[r.khoXuat] || r.khoNhan || r.khoXuat || '';
      html += '<option value="'+r.soPhieu+'">'+r.soPhieu+' ('+shortName+') ['+r.trangThai+']</option>';
    });
    selectEl.innerHTML = html; document.getElementById("ql-view-phieu").style.display = "none";
    var serverMs = parsed.meta && parsed.meta._debugTotalMs;
    document.getElementById("ql-stats").innerHTML = '<div class="stat-box" style="color:#d93025;">🔔 MỚI: '+countMoi+'</div> | <div class="stat-box" style="color:#137333;">✅ ĐÃ XỬ LÝ: '+countDone+'</div> | <div class="stat-box" style="color:#8b5a2b;">🚫 HỦY: '+countCancel+'</div>' +
      (serverMs ? (' <span style="color:#64748b;font-size:12px;">(' + Math.round(serverMs/1000) + 's)</span>') : '');
    // #region agent log
    fetch('http://127.0.0.1:7480/ingest/48e8fdfc-ebb8-4d81-9aee-1659862ac812',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'4a6e3c'},body:JSON.stringify({sessionId:'4a6e3c',location:'app.js:ql_loadPhieu',message:'ql list done',data:{clientMs:Date.now()-t0,serverMs:serverMs,scanned:parsed.meta&&parsed.meta._debugScanned,count:rows.length,run:parsed.meta&&parsed.meta._debugRun,ngay:ngay},timestamp:Date.now(),hypothesisId:'QL-A',runId:'ql-fast-v1'})}).catch(function(){});
    // #endregion
    if (selectedSoPhieu) {
      var matched = rows.find(function(item) { return item.soPhieu === selectedSoPhieu; });
      if (matched) {
        selectEl.value = selectedSoPhieu;
        ql_hienThiChiTiet(matched);
      }
    }
  }).catch(function(err){
    // #region agent log
    fetch('http://127.0.0.1:7480/ingest/48e8fdfc-ebb8-4d81-9aee-1659862ac812',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'4a6e3c'},body:JSON.stringify({sessionId:'4a6e3c',location:'app.js:ql_loadPhieu',message:'ql list error',data:{error:String(err&&err.message||err),clientMs:Date.now()-t0},timestamp:Date.now(),hypothesisId:'QL-A',runId:'ql-fast-v1'})}).catch(function(){});
    // #endregion
    selectEl.innerHTML = '<option value="">-- Lỗi tải --</option>';
    alert('Lỗi: '+err.message);
  });
}

function ql_onSelectPhieu() {
  var val = document.getElementById("ql-phieu").value;
  if(!val) { document.getElementById("ql-view-phieu").style.display = "none"; return; }
  ql_hienThiChiTiet(phieuData.find(x => x.soPhieu === val));
}

function confirm_loadPhieu(selectedSoPhieu) {
  var selectEl = document.getElementById("confirm-phieu");
  if (!selectEl) return;
  var viewEl = document.getElementById("confirm-view");
  if (!selectedSoPhieu && viewEl) viewEl.style.display = "none";
  if (!selectedSoPhieu) currentConfirmPhieuObj = null;
  selectEl.innerHTML = '<option value="">⏳ Đang tải...</option>';
  var confirmStoreFilter = sessionUser.role === 'Admin' ? 'all' : (sessionUser.store || '');
  var ngay = getNgayFilterParam('confirm-ngay', 'confirm-ngay-fast');
  var t0 = Date.now();
  // #region agent log
  fetch('http://127.0.0.1:7480/ingest/48e8fdfc-ebb8-4d81-9aee-1659862ac812',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'4a6e3c'},body:JSON.stringify({sessionId:'4a6e3c',location:'app.js:confirm_loadPhieu',message:'confirm list start',data:{ngay:ngay,build:APP_BUILD},timestamp:Date.now(),hypothesisId:'QL-B',runId:'ql-fast-v1'})}).catch(function(){});
  // #endregion
  apiGet('layDanhSachPhieuTheoFilter', { khoNhan: confirmStoreFilter, ngay: ngay, userRole: sessionUser.role, userStore: sessionUser.store }, { directOnly: true, timeoutMs: 45000 }).then(function(res) {
    var parsed = unwrapListResponse_(res);
    var confirmableOrders = parsed.rows.filter(function(r) { return r.trangThai === "Đã soạn"; });
    var html = '<option value="">-- Chọn Phiếu (' + confirmableOrders.length + ') --</option>';
    confirmableOrders.forEach(function(r) {
      var shortName = storeMap[r.khoNhan] || storeMap[r.khoXuat] || r.khoNhan || r.khoXuat || '';
      html += '<option value="'+r.soPhieu+'">'+r.soPhieu+' ('+shortName+') ['+r.trangThai+']</option>';
    });
    selectEl.innerHTML = html;
    // #region agent log
    fetch('http://127.0.0.1:7480/ingest/48e8fdfc-ebb8-4d81-9aee-1659862ac812',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'4a6e3c'},body:JSON.stringify({sessionId:'4a6e3c',location:'app.js:confirm_loadPhieu',message:'confirm list done',data:{clientMs:Date.now()-t0,serverMs:parsed.meta&&parsed.meta._debugTotalMs,scanned:parsed.meta&&parsed.meta._debugScanned,count:confirmableOrders.length,run:parsed.meta&&parsed.meta._debugRun,ngay:ngay},timestamp:Date.now(),hypothesisId:'QL-B',runId:'ql-fast-v1'})}).catch(function(){});
    // #endregion
    if (selectedSoPhieu) {
      var found = confirmableOrders.some(function(item) { return item.soPhieu === selectedSoPhieu; });
      if (found) {
        selectEl.value = selectedSoPhieu;
        confirm_onSelectPhieu();
      }
    }
  }).catch(function(err){
    selectEl.innerHTML = '<option value="">-- Lỗi tải --</option>';
    alert('Lỗi: '+err.message);
  });
}

function confirm_onSelectPhieu() {
  var val = document.getElementById("confirm-phieu").value;
  if (!val) { document.getElementById("confirm-view").style.display = "none"; return; }
  showLoad("Đang tải dữ liệu nhận hàng...");
  var t0 = Date.now();
  apiGet('getChiTietPhieu', { soPhieu: val, storeName: sessionUser.store, includeStock: '0' }, { directOnly: true, timeoutMs: 45000 }).then(function(res) {
    hideLoad();
    var parsed = unwrapListResponse_(res);
    var rows = parsed.rows;
    currentConfirmPhieuObj = { soPhieu: val };
    var viewEl = document.getElementById("confirm-view");
    document.getElementById("confirm-lbl-sophieu").innerText = val;
    var serverMs = parsed.meta && parsed.meta._debugTotalMs;
    document.getElementById("confirm-meta").innerText = "Đã tải " + rows.length + " dòng để xác nhận." + (serverMs ? (" (" + Math.round(serverMs/1000) + "s)") : "");
    // #region agent log
    fetch('http://127.0.0.1:7480/ingest/48e8fdfc-ebb8-4d81-9aee-1659862ac812',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'4a6e3c'},body:JSON.stringify({sessionId:'4a6e3c',location:'app.js:confirm_onSelectPhieu',message:'confirm detail done',data:{clientMs:Date.now()-t0,serverMs:serverMs,rows:rows.length,run:parsed.meta&&parsed.meta._debugRun,stock:parsed.meta&&parsed.meta._debugStock},timestamp:Date.now(),hypothesisId:'QL-C',runId:'ql-fast-v1'})}).catch(function(){});
    // #endregion
    var tbody = document.getElementById("confirm-tbody");
    tbody.innerHTML = "";
    rows.filter(function(r){ return r.trangThai !== "Đã hủy dòng" && r.trangThai !== "Đã hủy đơn"; }).forEach(function(r, idx) {
      var packedQty = (r.slThucTe !== undefined && r.slThucTe !== null && r.slThucTe !== "") ? Number(r.slThucTe) : Number(r.slGoc || 0);
      var inputValue = packedQty;
      tbody.insertAdjacentHTML('beforeend', '<tr><td>' + (idx + 1) + '</td><td><b>' + (r.maVach || '') + '</b><br><small style="color:gray;">' + (r.maHang || '') + '</small><br><small>' + (r.tenHang || '') + '</small></td><td>' + (r.dvt || 'Cái') + '</td><td style="font-weight:700;">' + packedQty + '</td><td><input type="number" class="confirm-qty-input same" data-row="' + r.rowIndex + '" data-packed="' + packedQty + '" data-previous="' + packedQty + '" value="' + inputValue + '" min="0" oninput="confirm_updateInput(this)"></td></tr>');
    });
    viewEl.style.display = "block";
  }).catch(function(err){ hideLoad(); alert('Lỗi: '+err.message); });
}

function confirm_updateInput(inputEl) {
  var packed = Number(inputEl.getAttribute("data-packed") || 0);
  var current = Number(inputEl.value || 0);
  inputEl.classList.toggle("changed", current !== packed);
  inputEl.classList.toggle("same", current === packed);
}

function confirm_xacNhanNhanHang() {
  if (!currentConfirmPhieuObj || !sessionUser.user) return alert("Vui lòng chọn phiếu và đăng nhập trước khi xác nhận.");
  var confirmations = [];
  document.querySelectorAll(".confirm-qty-input").forEach(function(inputEl) {
    var row = parseInt(inputEl.getAttribute("data-row"));
    var qty = Number(inputEl.value);
    var previousQty = Number(inputEl.getAttribute("data-previous") || 0);
    if (!isNaN(row) && !isNaN(qty) && qty >= 0) confirmations.push({ row: row, receivedQty: qty, previousQty: previousQty });
  });
  if (!confirmations.length) return alert("Không có dữ liệu xác nhận.");
  showLoad("Đang lưu xác nhận...");
  apiPost('xacNhanNhanHang', { soPhieu: currentConfirmPhieuObj.soPhieu, actor: sessionUser.user, store: sessionUser.store, confirmations: confirmations }).then(function(res) {
    hideLoad();
    if (!res.success) throw new Error(res.error || "Không thể lưu xác nhận.");
    alert("✅ Đã lưu xác nhận nhận hàng.");
    resetConfirmViewAfterSave();
  }).catch(function(err){ hideLoad(); alert('Lỗi: '+err.message); });
}

function openDeepLinkedOrder() {
  var params = getDeepLinkParams();
  var targetOrder = params.order || deepLinkOrder;
  var targetTab = params.tab || deepLinkTab;
  var isPublicView = params.public || params.tab === 'public' || params.order === 'public' || params.view === 'public';
  if (!targetOrder) return;
  if (targetTab === "xac-nhan" || targetTab === "nhan-hang") {
    showLoad("Đang mở đơn hàng...");
    activateTab('tab-xac-nhan');
    confirm_loadPhieu(targetOrder);
    return;
  }
  if (targetTab && targetTab !== "quan-ly" && targetTab !== "public") return;
  showLoad("Đang mở đơn hàng...");
  activateTab('tab-quan-ly');
  apiGet('getThongTinPhieu', { soPhieu: targetOrder }).then(function(phieu) {
    hideLoad();
    if (!phieu || !phieu.soPhieu) {
      alert("Không tìm thấy đơn hàng: " + targetOrder);
      return;
    }
    if (isPublicView) {
      currentPhieuObj = phieu;
      ql_hienThiChiTiet(phieu, { publicView: true });
      setTimeout(function() {
        if (currentPhieuObj && currentPhieuObj.soPhieu) {
          ql_inWeb_FromEdit();
        }
      }, 800);
      return;
    }
    ql_loadPhieu(targetOrder);
  }).catch(function(err) { hideLoad(); alert('Lỗi mở đơn hàng: ' + err.message); });
}

function ql_hienThiChiTiet(phieu, options) {
  if (!phieu) return;
  currentPhieuObj = phieu;
  var isPublicView = !!(options && options.publicView);
  var isAdmin = sessionUser.role === "Admin";
  showLoad("Tải chi tiết...");
  var t0Detail = Date.now();
  apiGet('getChiTietPhieu', { soPhieu: currentPhieuObj.soPhieu, storeName: currentPhieuObj.khoXuat, includeStock: '1' }, { directOnly: true, timeoutMs: 45000 }).then(function(res) {
    hideLoad();
    var parsed = unwrapListResponse_(res);
    var rows = parsed.rows;
    // #region agent log
    fetch('http://127.0.0.1:7480/ingest/48e8fdfc-ebb8-4d81-9aee-1659862ac812',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'4a6e3c'},body:JSON.stringify({sessionId:'4a6e3c',location:'app.js:ql_hienThiChiTiet',message:'ql detail done',data:{clientMs:Date.now()-t0Detail,serverMs:parsed.meta&&parsed.meta._debugTotalMs,rows:rows.length,run:parsed.meta&&parsed.meta._debugRun,stock:parsed.meta&&parsed.meta._debugStock,khoXuat:currentPhieuObj&&currentPhieuObj.khoXuat},timestamp:Date.now(),hypothesisId:'QL-C',runId:'ql-fast-v1'})}).catch(function(){});
    // #endregion
    editRows = rows; currentLoadedRows = rows || []; var tb = document.getElementById("ql-tbody"); tb.innerHTML = "";
    var isConfirmedOrder = rows.some(function(r) { return r.trangThai === "Đã xác nhận nhận hàng"; });
    var isPackedOrder = rows.some(function(r) { return r.trangThai === "Đã soạn hàng"; });
    var canEditRows = !!sessionUser.user && !isPublicView && !isConfirmedOrder && (isAdmin || !isPackedOrder);
    var canAddItems = canEditRows;
    var canCancelOrder = !!sessionUser.user && !isPublicView && !isConfirmedOrder && isAdmin;

    document.getElementById("ql-admin-actions").style.display = canAddItems ? "flex" : "none";
    document.getElementById("ql-admin-column").style.display = canEditRows ? "table-cell" : "none";
    document.getElementById("ql-btn-cancel-order").style.display = canCancelOrder ? "inline-block" : "none";
    document.getElementById("ql-btn-save").style.display = canEditRows ? "inline-block" : "none";
    document.getElementById("ql-btn-print").style.display = "inline-block";
    document.getElementById("ql-btn-excel").style.display = "inline-block";
    if (isPublicView) {
      document.getElementById("ql-order-meta").innerText = "Chế độ xem công khai – chỉ đọc";
    } else {
      document.getElementById("ql-order-meta").innerText = "";
      if (isConfirmedOrder) document.getElementById("ql-order-meta").innerText = "Đơn đã xác nhận nhận hàng - chỉ xem.";
      else if (isPackedOrder && !isAdmin) document.getElementById("ql-order-meta").innerText = "Đơn đã soạn xong - chỉ Admin mới được phép sửa.";
      else if (isPackedOrder && isAdmin) document.getElementById("ql-order-meta").innerText = "Đơn đã soạn xong - Admin sửa sẽ mở lại đơn để soạn lại.";
    }
    document.getElementById("ql-lbl-sophieu").innerText = currentPhieuObj.soPhieu;
    document.getElementById("ql-lbl-khoxuat").innerText = currentPhieuObj.khoXuat + ' (' + (storeMap[currentPhieuObj.khoXuat] || '') + ')';
    document.getElementById("ql-lbl-khonhan").innerText = currentPhieuObj.khoNhan + ' (' + (storeMap[currentPhieuObj.khoNhan] || '') + ')';

    rows.forEach((r, i) => {
      var isCancelled = r.trangThai === "Đã hủy dòng" || r.trangThai === "Đã hủy đơn";
      var isReadOnlyRow = isConfirmedOrder || isCancelled || !canEditRows;
      var rowStyle = isCancelled ? 'background:#fce8e6; color:#777; text-decoration:line-through;' : (r.ghiChu ? 'background:#fff8e1;' : '');
      var rowKey = String(r.tempKey || r.rowIndex);
      var dvtDisplay = r.dvt || "Cái";
      var variantHtml = ql_renderVariantSelector(r, rowKey, isReadOnlyRow);
      var quantityInput = '<input type="number" class="edit-sl-input" data-row="'+r.rowIndex+'" data-new="0" value="'+(r.slGoc || 0)+'" '+(isReadOnlyRow ? 'disabled' : '')+' style="border:2px solid #1a73e8;text-align:center;width:70px;">';
      var cancelButton = canAddItems ? '<td><button type="button" onclick="ql_huyDong('+r.rowIndex+')" '+(isCancelled ? 'disabled' : '')+' style="border:none; background:#d93025; color:white; border-radius:5px; padding:7px 9px; cursor:pointer;">Hủy mã</button></td>' : '<td></td>';
      var requestedDisplay = (!isReadOnlyRow && !isPublicView)
        ? quantityInput
        : ('<b>' + (Number(r.slGoc) || 0) + '</b>');
      var hasReceivedQty = (r.slThucTe !== undefined && r.slThucTe !== null && r.slThucTe !== "");
      var receivedDisplay = hasReceivedQty
        ? ('<b style="color:#166534;">' + Number(r.slThucTe) + '</b>')
        : '<span style="color:#94a3b8;">-</span>';
      tb.insertAdjacentHTML('beforeend', '<tr style="'+rowStyle+'"><td>'+(i+1)+'</td><td>'+variantHtml+'<br><small style="color:gray;">Mã hàng hóa: '+(r.maHang||'')+'</small></td><td>'+r.tenHang+(r.ghiChu?'<br><b style="color:red;font-size:11px;">⚠️ '+r.ghiChu+'</b>':'')+((isCancelled || r.trangThai === "Đã soạn hàng" || r.trangThai === "Đã xác nhận nhận hàng") ? '<br><b style="color:#d93025;font-size:11px;">'+r.trangThai+'</b>' : '')+'</td><td id="ql-dvt-'+rowKey+'">'+dvtDisplay+'</td><td>'+r.stock+'</td><td>'+requestedDisplay+'</td><td>'+receivedDisplay+'</td>'+cancelButton+'</tr>');
    });
    var packerNames = rows.map(function(r){ return r.nguoiSoanHang || ""; }).filter(Boolean);
    if (!isPublicView && packerNames.length) document.getElementById("ql-order-meta").innerText = "Người soạn hàng gần nhất: " + packerNames[packerNames.length - 1];
    if (isConfirmedOrder && !isPublicView) {
      document.getElementById("ql-order-meta").innerText += " | Chế độ chỉ xem sau khi xác nhận nhận hàng";
    } else if (isPackedOrder && !isAdmin && !isPublicView && packerNames.length) {
      document.getElementById("ql-order-meta").innerText += " | Chi nhánh không thể sửa sau khi đã soạn xong";
    }
    document.getElementById("ql-view-phieu").style.display = "block";
  }).catch(function(err){ hideLoad(); alert('Lỗi: '+err.message); });
}

function ql_getVariantOptions(rowMeta) {
  var maHang = String(rowMeta && rowMeta.maHang ? rowMeta.maHang : '').trim().toUpperCase();
  var seen = {};
  var variants = [];
  danhMucArr.forEach(function(item) {
    if (!item) return;
    var itemMaHang = String(item.maHang || '').trim().toUpperCase();
    var sameGroup = maHang && itemMaHang === maHang;
    if (!sameGroup) return;
    var key = String(item.maVach || '').trim().toUpperCase() + '|' + String(item.dvt || 'Cái').trim().toUpperCase();
    if (!key || seen[key]) return;
    seen[key] = true;
    variants.push({
      maVach: item.maVach || '',
      dvt: item.dvt || 'Cái',
      tenHang: item.tenHang || rowMeta.tenHang || '',
      maHang: item.maHang || rowMeta.maHang || ''
    });
  });

  var currentKey = String(rowMeta.maVach || '').trim().toUpperCase() + '|' + String(rowMeta.dvt || 'Cái').trim().toUpperCase();
  if (!seen[currentKey]) {
    variants.unshift({
      maVach: rowMeta.maVach || '',
      dvt: rowMeta.dvt || 'Cái',
      tenHang: rowMeta.tenHang || '',
      maHang: rowMeta.maHang || ''
    });
  }
  return variants;
}

function ql_renderVariantSelector(rowMeta, rowKey, isReadOnlyRow) {
  var variants = ql_getVariantOptions(rowMeta);
  var currentMaVach = String(rowMeta.maVach || '').trim().toUpperCase();
  if (isReadOnlyRow || variants.length <= 1) {
    return '<b>Mã vạch: ' + (rowMeta.maVach || '') + '</b>';
  }
  var optionsHtml = '';
  variants.forEach(function(v) {
    var selected = String(v.maVach || '').trim().toUpperCase() === currentMaVach ? ' selected' : '';
    var payload = encodeURIComponent(JSON.stringify(v));
    optionsHtml += '<option value="' + payload + '"' + selected + '>' + (v.maVach || '-') + ' | ĐVT: ' + (v.dvt || 'Cái') + '</option>';
  });
  return '<select class="ql-variant-select" data-row="' + rowKey + '" onchange="ql_onVariantChange(this)">' + optionsHtml + '</select>';
}

function ql_onVariantChange(selectEl) {
  if (!selectEl) return;
  var rowKey = String(selectEl.getAttribute('data-row') || '');
  var raw = selectEl.value ? decodeURIComponent(selectEl.value) : '';
  if (!rowKey || !raw) return;
  var variant = null;
  try { variant = JSON.parse(raw); } catch(e) { return; }
  var rowMeta = editRows.find(function(item) { return String(item.tempKey || item.rowIndex) === rowKey; });
  if (!rowMeta || !variant) return;
  rowMeta.maVach = variant.maVach || rowMeta.maVach;
  rowMeta.dvt = variant.dvt || rowMeta.dvt || 'Cái';
  rowMeta.tenHang = variant.tenHang || rowMeta.tenHang;
  if (variant.maHang) rowMeta.maHang = variant.maHang;
  var dvtCell = document.getElementById('ql-dvt-' + rowKey);
  if (dvtCell) dvtCell.innerText = rowMeta.dvt || 'Cái';
}

function ql_isDuplicateOrderItem(item) {
  if (!item) return false;
  var itemMaHang = String(item.maHang || "").trim().toUpperCase();
  var itemMaVach = String(item.maVach || "").trim().toUpperCase();
  return editRows.some(function(existing) {
    if (!existing) return false;
    var existingMaHang = String(existing.maHang || "").trim().toUpperCase();
    var existingMaVach = String(existing.maVach || "").trim().toUpperCase();
    if (itemMaHang && existingMaHang && itemMaHang === existingMaHang) return true;
    if (itemMaVach && existingMaVach && itemMaVach === existingMaVach) return true;
    return false;
  });
}

function ql_luuSua() {
  var inputs = document.querySelectorAll(".edit-sl-input");
  var updates = [];
  var newItems = [];
  var variantByRow = {};
  document.querySelectorAll('.ql-variant-select').forEach(function(selectEl) {
    var rowKey = String(selectEl.getAttribute('data-row') || '');
    if (!rowKey || !selectEl.value) return;
    try {
      var variant = JSON.parse(decodeURIComponent(selectEl.value));
      variantByRow[rowKey] = variant;
    } catch(e) {}
  });
  inputs.forEach(function(ip) {
    var isNew = ip.getAttribute("data-new") === "1";
    var rowKey = ip.getAttribute("data-row");
    var qtyValue = ip.value;
    if (!ip.disabled && rowKey) {
      if (isNew) {
        var rowMeta = editRows.find(function(item) { return String(item.tempKey || item.rowIndex) === String(rowKey); });
        if (rowMeta) {
          newItems.push({ maHang: rowMeta.maHang, maVach: rowMeta.maVach, tenHang: rowMeta.tenHang, dvt: rowMeta.dvt, sl: qtyValue });
        }
      } else {
        var variant = variantByRow[String(rowKey)] || null;
        updates.push({
          row: parseInt(rowKey),
          valSl: qtyValue,
          valMaVach: variant ? (variant.maVach || '') : '',
          valDvt: variant ? (variant.dvt || '') : '',
          valTenHang: variant ? (variant.tenHang || '') : ''
        });
      }
    }
  });
  if (!updates.length && !newItems.length) return alert("Không có thay đổi nào cần lưu.");
  showLoad("Đang lưu chỉnh sửa...");
  apiPost('luuChinhSuaPhieu', { soPhieu: currentPhieuObj ? currentPhieuObj.soPhieu : "", updates: updates, newItems: newItems, actor: sessionUser.user }).then(function(res) {
    hideLoad();
    if (!res.success) throw new Error(res.error || "Không thể lưu thay đổi.");
    alert("✅ Đã lưu chỉnh sửa thành công! Thông báo cập nhật đơn sẽ được gửi sau khi lưu.");
    resetManagementViewAfterSave();
  }).catch(function(err){ hideLoad(); alert('Lỗi: '+err.message); });
}

function ql_themMaHang(itemOverride) {
  if (!sessionUser.user) return alert("Vui lòng đăng nhập trước khi thao tác.");
  var inputEl = document.getElementById("ql-add-code");
  var code = (inputEl ? inputEl.value : "").trim().toUpperCase();
  var quantity = Number(document.getElementById("ql-add-qty").value);
  var selectedItem = qlSelectedSuggestedItem && ql_isItemMatchedWithCode(qlSelectedSuggestedItem, code) ? qlSelectedSuggestedItem : null;
  var item = itemOverride || selectedItem || danhMucGoc[code] || filterProducts(code)[0];
  if (!item) return alert("Không tìm thấy mã hàng hóa hoặc mã vạch.");
  if (!quantity || quantity < 1) return alert("Số lượng phải lớn hơn 0.");
  if (ql_isDuplicateOrderItem(item)) {
    alert("⚠️ Mã này đã tồn tại trong đơn hiện tại. Không thể thêm dòng trùng.");
    return;
  }
  qlSelectedSuggestedItem = null;
  if (inputEl) inputEl.value = "";
  document.getElementById("ql-add-qty").value = "1";
  document.getElementById("ql-suggest-box").style.display = "none";

  var tb = document.getElementById("ql-tbody");
  if (tb) {
    var latestQty = Number(quantity) || 1;
    var tempKey = 'new-' + Date.now() + '-' + Math.floor(Math.random() * 1000);
    editRows.push({
      tempKey: tempKey,
      maHang: item.maHang,
      maVach: item.maVach,
      tenHang: item.tenHang,
      dvt: item.dvt || "Cái",
      slGoc: latestQty,
      slThucTe: latestQty,
      stock: 0,
      ghiChu: "",
      trangThai: "Chưa lưu",
      nguoiSoanHang: ""
    });
    tb.insertAdjacentHTML('beforeend', '<tr><td>' + (editRows.length) + '</td><td><b>Mã vạch: ' + (item.maVach || '') + '</b><br><small style="color:gray;">Mã hàng hóa: ' + (item.maHang || '') + '</small></td><td>' + (item.tenHang || '') + '</td><td>' + (item.dvt || 'Cái') + '</td><td>0</td><td><input type="number" class="edit-sl-input" data-row="' + tempKey + '" data-new="1" value="' + latestQty + '" style="border:2px solid #1a73e8;text-align:center;width:70px;"></td><td><span style="color:#94a3b8;">-</span></td><td><button type="button" onclick="ql_huyDong(' + tempKey + ')" style="border:none; background:#d93025; color:white; border-radius:5px; padding:7px 9px; cursor:pointer;">Hủy mã</button></td></tr>');
  }

  alert("✅ Đã thêm mã vào bảng chỉnh sửa. Nhấn Lưu để ghi vào hệ thống.");
}

function ql_huyDong(row) {
  if (!sessionUser.user) return alert("Vui lòng đăng nhập trước khi thao tác.");
  if (!confirm("Hủy mã hàng này khỏi đơn? Dữ liệu sẽ được lưu lịch sử.")) return;
  showLoad("Đang hủy mã hàng...");
  apiPost('huyDongChiTietPhieu', { row: row, actor: sessionUser.user }).then(function(res) {
    hideLoad();
    if (!res.success) throw new Error(res.error || "Không thể hủy mã.");
    alert("✅ Đã hủy mã khỏi đơn thành công! Thông báo cập nhật đơn sẽ được gửi sau khi lưu.");
    ql_hienThiChiTiet(currentPhieuObj);
  }).catch(function(err) { hideLoad(); alert('Lỗi: ' + err.message); });
}

function ql_huyPhieu() {
  if (!sessionUser.user) return alert("Vui lòng đăng nhập trước khi thao tác.");
  if (!confirm("Hủy toàn bộ đơn " + currentPhieuObj.soPhieu + "? Đơn sẽ vẫn được giữ trong Google Sheet với trạng thái Đã hủy.")) return;
  showLoad("Đang hủy đơn...");
  apiPost('huyPhieu', { soPhieu: currentPhieuObj.soPhieu, actor: sessionUser.user }).then(function(res) {
    hideLoad();
    if (!res.success) throw new Error(res.error || "Không thể hủy đơn.");
    alert("✅ Đơn đã được hủy và lưu lịch sử.");
    ql_hienThiChiTiet(currentPhieuObj);
  }).catch(function(err) { hideLoad(); alert('Lỗi: ' + err.message); });
}

function layItemsTuBangSua() {
  var items = [];
  document.querySelectorAll(".edit-sl-input").forEach(function(ip) {
    var rowKey = ip.getAttribute("data-row");
    var rowMeta = editRows.find(function(item) { return String(item.tempKey || item.rowIndex) === String(rowKey); });
    if (rowMeta) {
      items.push({ maHang: rowMeta.maHang, maVach: rowMeta.maVach, tenHang: rowMeta.tenHang, dvt: rowMeta.dvt, sl: ip.value });
    }
  });
  return items;
}
function layItemsForOutput() {
  var isPublicView = new URLSearchParams(location.search).get("public") === "1" || new URLSearchParams(location.search).get("view") === "public";
  if (!isPublicView) return layItemsTuBangSua();
  return (currentLoadedRows || []).map(function(row) {
    return {
      maHang: row.maHang,
      maVach: row.maVach,
      tenHang: row.tenHang,
      dvt: row.dvt,
      sl: (row.slThucTe !== undefined && row.slThucTe !== null && row.slThucTe !== "") ? row.slThucTe : row.slGoc
    };
  });
}
function ql_inWeb_FromEdit() { executePrintWeb(currentPhieuObj.soPhieu, currentPhieuObj.khoXuat, currentPhieuObj.khoNhan, layItemsForOutput()); }
function ql_xuatExcel_FromEdit() { executeExportExcel(currentPhieuObj.soPhieu, currentPhieuObj.khoXuat, currentPhieuObj.khoNhan, layItemsForOutput()); }

// ================= SOẠN HÀNG MOBILE =================
function sh_taiDanhSachDon() {
  var createDateEl = document.getElementById("sh-create-date");
  if (createDateEl && !createDateEl.value) {
    var now = new Date();
    var yyyy = now.getFullYear();
    var mm = String(now.getMonth() + 1).padStart(2, '0');
    var dd = String(now.getDate()).padStart(2, '0');
    createDateEl.value = yyyy + '-' + mm + '-' + dd;
  }
  document.getElementById("sh-list-container").innerHTML = '<div class="card" style="text-align:center; color:gray;">Vui lòng chọn số phiếu ở trên.</div>';
  document.getElementById("sh-footer").style.display = "none";
  document.getElementById("sh-phieu").innerHTML = '<option value="">⏳ Đang tải...</option>';
  apiGet('getDonHangTheoNgay', { ngay: document.getElementById("sh-ngay").value, userRole: sessionUser.role, userStore: sessionUser.store, viewMode: 'packing' }, { allowDirectFallback: true }).then(function(res) {
    var parsed = unwrapListResponse_(res);
    var rows = parsed.rows;
    var countMoi = 0; var countDone = 0;
    var html = '<option value="">-- Chọn đơn ('+rows.length+') --</option>';
    rows.forEach(function(r) {
      if(r.trangThai === "Mới") countMoi++; else countDone++;
      var shortName = storeMap[r.khoNhan] || storeMap[r.khoXuat] || r.khoNhan || r.khoXuat || '';
      html += '<option value="'+r.soPhieu+'">'+r.soPhieu+' ('+shortName+') ['+r.trangThai+']</option>';
    });
    document.getElementById("sh-phieu").innerHTML = html;
    document.getElementById("sh-stats").innerHTML = '<div class="stat-box" style="color:#d93025;">🔔 CẦN SOẠN: '+countMoi+'</div> | <div class="stat-box" style="color:#137333;">✅ ĐÃ SOẠN: '+countDone+'</div>';
  }).catch(function(err){ alert('Lỗi: '+err.message); });
}

var pendingImages = {}; var isCompressing = 0;
function sh_chonDonMobile() {
  var sp = document.getElementById("sh-phieu").value; if(!sp) return;
  document.getElementById("sh-list-container").innerHTML = '<div style="text-align:center;">⏳ Đang tải SP...</div>';
  pendingImages = {}; isCompressing = 0;
  apiGet('getChiTietDonHangMobile', { soPhieu: sp }).then(function(items) {
    var html = "";
    items.forEach((it, j) => {
      var stockDisplay = (it.stock !== undefined && it.stock !== null) ? Number(it.stock) : 0;
      html += '<div class="item-card"><b>'+it.tenHang+'</b><br><small><b>Mã vạch: '+it.maVach+'</b> | Mã hàng hóa: '+(it.maHang||'')+' | ĐVT: '+it.dvt+' | Tồn hiện tại: <b style="color:#1d4ed8;">'+stockDisplay+'</b></small><div class="action-row"><div>SL Yêu Cầu: <b>'+it.slGoc+'</b><br>Thực tế: <input type="number" class="sl-thuc-te" data-row="'+it.rowIndex+'" value="'+it.slThucTe+'"></div><div style="text-align:right;"><label class="btn-camera" for="c-'+j+'">📷 Ảnh</label><input type="file" id="c-'+j+'" accept="image/*" capture="environment" style="display:none;" data-row="'+it.rowIndex+'" data-j="'+j+'" onchange="nenAnh(this)"><img id="p-'+j+'" src="'+(it.anhXacNhan||'')+'" style="display:'+(it.anhXacNhan?'block':'none')+'; width:50px; height:50px; margin-top:5px;"><small id="st-'+j+'"></small></div></div></div>';
    });
    document.getElementById("sh-list-container").innerHTML = html; document.getElementById("sh-footer").style.display = "block";
  }).catch(function(err){ alert('Lỗi: '+err.message); });
}

function nenAnh(inputEl) {
  var r = inputEl.getAttribute("data-row"), j = inputEl.getAttribute("data-j");
  if (inputEl.files[0]) {
    isCompressing++; document.getElementById("st-"+j).innerText = "⏳ Nén...";
    var reader = new FileReader(); reader.onload = function(e) {
      var img = new Image(); img.onload = function() {
        var cvs = document.createElement("canvas"), ctx = cvs.getContext("2d"); var w = img.width, h = img.height; if(w>800){h*=800/w; w=800;}
        cvs.width=w; cvs.height=h; ctx.drawImage(img,0,0,w,h); pendingImages[r] = cvs.toDataURL("image/jpeg", 0.8);
        document.getElementById("p-"+j).src = pendingImages[r]; document.getElementById("p-"+j).style.display="block"; document.getElementById("st-"+j).innerText = "✅ Xong"; isCompressing--;
      }; img.src = e.target.result;
    }; reader.readAsDataURL(inputEl.files[0]);
  }
}

function sh_luuPhieu() {
  if(isCompressing > 0) return alert("Đợi ảnh nén xong!");
  showLoad("Đang lưu kết quả lên hệ thống...");
  var inputs = document.querySelectorAll(".sl-thuc-te"), updates = []; inputs.forEach(ip => updates.push({ row: parseInt(ip.getAttribute("data-row")), val: ip.value }));
  apiPost('luuSoSoanHangVaAnh', { updates: updates, images: pendingImages, actor: sessionUser ? sessionUser.user : '' }).then(function(res) {
    hideLoad();
    var message = (res && res.message) ? res.message : String(res || 'Đã lưu.');
    alert(message);
    pendingImages = {};
    sh_taiDanhSachDon();
    if (res && res.notify && res.notify.soPhieu) {
      apiPost('postProcessPackingOrder', res.notify).catch(function() {});
    }
  }).catch(function(err){ hideLoad(); alert('Lỗi: '+err.message); });
}

var shOrderCandidates = [];

function sh_capNhatTomTatChonDonSoan() {
  var checkboxes = document.querySelectorAll('.sh-order-check');
  var checked = 0;
  checkboxes.forEach(function(cb) { if (cb.checked) checked++; });
  var summaryEl = document.getElementById('sh-order-picker-summary');
  if (summaryEl) {
    summaryEl.innerText = 'Đang chọn ' + checked + '/' + checkboxes.length + ' đơn hợp lệ trong ngày (đơn đã soạn/đã giao đã được ẩn).';
  }
}

function sh_renderDanhSachDonSoan(candidates) {
  var bodyEl = document.getElementById('sh-order-picker-body');
  if (!bodyEl) return;
  var createDateEl = document.getElementById('sh-create-date');
  var ngayLabel = createDateEl && createDateEl.value ? createDateEl.value : 'hôm nay';
  if (!candidates || !candidates.length) {
    bodyEl.innerHTML = '<tr><td colspan="6" style="text-align:center; color:#b91c1c; padding:14px;">Không có đơn mới hợp lệ ngày <b>' + escapeHtml(ngayLabel) + '</b>.<br><small style="color:#64748b;">Thử đổi "Tạo bảng từ đơn ngày" sang ngày có đơn (vd. hôm qua).</small></td></tr>';
    sh_capNhatTomTatChonDonSoan();
    return;
  }
  var html = '';
  candidates.forEach(function(order, idx) {
    html += '<tr>' +
      '<td style="text-align:center; font-weight:700; color:#334155;">' + (idx + 1) + '</td>' +
      '<td style="text-align:center;"><input type="checkbox" class="sh-order-check" data-sophieu="' + order.soPhieu + '" checked onchange="sh_capNhatTomTatChonDonSoan()"></td>' +
      '<td><b>' + order.soPhieu + '</b></td>' +
      '<td>' + (order.khoXuat || '-') + '</td>' +
      '<td>' + (order.khoNhan || '-') + '</td>' +
      '<td>' + (order.thoiGianDat || '-') + '</td>' +
      '</tr>';
  });
  bodyEl.innerHTML = html;
  sh_capNhatTomTatChonDonSoan();
}

function dbgSoanLine_(label, data) {
  try { console.warn('[donhang:' + label + ']', data); } catch (e) {}
}

function sh_taiDanhSachDonSoanChoBang() {
  var createDateEl = document.getElementById('sh-create-date');
  var ngay = createDateEl && createDateEl.value ? createDateEl.value : '';
  var pickerEl = document.getElementById('sh-order-picker');
  if (pickerEl) pickerEl.style.display = 'block';
  var bodyEl = document.getElementById('sh-order-picker-body');
  if (bodyEl) bodyEl.innerHTML = '<tr><td colspan="6" style="text-align:center; color:#64748b; padding:14px;">Đang tải danh sách đơn...</td></tr>';

  showLoad('Đang tải danh sách đơn soạn...');
  // #region agent log
  var _dbgListStart = Date.now();
  fetch('http://127.0.0.1:7480/ingest/48e8fdfc-ebb8-4d81-9aee-1659862ac812',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'4a6e3c'},body:JSON.stringify({sessionId:'4a6e3c',location:'app.js:sh_taiDanhSachDonSoanChoBang',message:'listDonSoan start',data:{ngay:ngay},timestamp:Date.now(),hypothesisId:'F',runId:'post-fix-v2'})}).catch(function(){});
  // #endregion
  dbgSoanLine_('listDonSoan.start', { ngay: ngay, build: APP_BUILD, via: 'proxy' });
  apiGet('getDanhSachDonSoanHang', {
    ngay: ngay,
    userRole: sessionUser.role || '',
    userStore: sessionUser.store || ''
  }, { allowDirectFallback: true, timeoutMs: 120000 }).then(function(res) {
    hideLoad();
    // #region agent log
    var _dbgListMs = Date.now() - _dbgListStart;
    dbgSoanLine_('listDonSoan.done', { clientMs: _dbgListMs, serverMs: res && res._debugTotalMs, total: res && res.total, date: res && res.date, run: res && res._debugRun, build: APP_BUILD });
    fetch('http://127.0.0.1:7480/ingest/48e8fdfc-ebb8-4d81-9aee-1659862ac812',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'4a6e3c'},body:JSON.stringify({sessionId:'4a6e3c',location:'app.js:sh_taiDanhSachDonSoanChoBang',message:'listDonSoan done',data:{clientMs:_dbgListMs,serverMs:res&&res._debugTotalMs,total:res&&res.total,run:res&&res._debugRun},timestamp:Date.now(),hypothesisId:'F',runId:'post-fix-v2'})}).catch(function(){});
    // #endregion
    if (!res || !res.success) {
      throw new Error((res && (res.error || res.msg)) || 'Không thể tải danh sách đơn.');
    }
    shOrderCandidates = Array.isArray(res.orders) ? res.orders : [];
    sh_renderDanhSachDonSoan(shOrderCandidates);
    sh_warmStockInBackground_();
  }).catch(function(err) {
    hideLoad();
    shOrderCandidates = [];
    if (bodyEl) bodyEl.innerHTML = '<tr><td colspan="6" style="text-align:center; color:#b91c1c; padding:14px;">Lỗi tải danh sách: ' + escapeHtml(err.message) + '</td></tr>';
    sh_capNhatTomTatChonDonSoan();
  });
}

function sh_warmStockInBackground_() {
  return sh_ensureStockReady_();
}

function sh_ensureStockReady_() {
  if (shStockWarmState.ready) return Promise.resolve(true);
  if (shStockWarmState.promise) return shStockWarmState.promise;
  shStockWarmState.warming = true;
  var summaryEl = document.getElementById('sh-order-picker-summary');
  if (summaryEl) summaryEl.innerText = 'Đang kiểm tra sheet TON_Q7 (tồn Kho Q7)...';

  shStockWarmState.promise = apiGet('getStockCacheStatus', null, { directOnly: true, timeoutMs: 30000 })
    .then(function(st) {
      // #region agent log
      fetch('http://127.0.0.1:7480/ingest/48e8fdfc-ebb8-4d81-9aee-1659862ac812',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'4a6e3c'},body:JSON.stringify({sessionId:'4a6e3c',location:'app.js:sh_ensureStockReady_',message:'getStockCacheStatus',data:{ready:!!(st&&st.ready),source:st&&st.source,stores:st&&st.stores,run:st&&st._debugRun},timestamp:Date.now(),hypothesisId:'A',runId:'q7-v1'})}).catch(function(){});
      // #endregion
      if (st && st.ready) {
        shStockWarmState.ready = true;
        shStockWarmState.warming = false;
        shStockWarmState.promise = null;
        dbgSoanLine_('warmStock.alreadyReady', { source: st.source, stores: st.stores });
        if (summaryEl) summaryEl.innerText = 'Tồn Q7 sẵn sàng (' + (st.source || 'TON_Q7') + ', ' + (st.stores || 0) + ' mã).';
        return true;
      }
      // Chưa có TON_Q7 → rebuild 1 lần từ sheet tổng (có thể lâu), rồi các lần sau đọc sheet nhẹ
      return apiGet('warmStockIndex', null, { directOnly: true, timeoutMs: 180000 }).then(function(w) {
        shStockWarmState.warming = false;
        shStockWarmState.promise = null;
        shStockWarmState.ready = !!(w && w.success && (w.ready !== false));
        shStockWarmState.lastMs = (w && w._debugTotalMs) || 0;
        dbgSoanLine_('warmStock.done', { serverMs: w && w._debugTotalMs, run: w && w._debugRun, stores: w && w.stores, rows: w && w.rows, ready: shStockWarmState.ready, source: w && w.cacheSource, rebuilt: w && w.rebuilt, via: 'direct-get' });
        // #region agent log
        fetch('http://127.0.0.1:7480/ingest/48e8fdfc-ebb8-4d81-9aee-1659862ac812',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'4a6e3c'},body:JSON.stringify({sessionId:'4a6e3c',location:'app.js:sh_ensureStockReady_',message:'warmStockIndex done',data:{ready:shStockWarmState.ready,source:w&&w.cacheSource,rows:w&&w.rows,rebuilt:w&&w.rebuilt,serverMs:w&&w._debugTotalMs,run:w&&w._debugRun},timestamp:Date.now(),hypothesisId:'C',runId:'q7-v1'})}).catch(function(){});
        // #endregion
        if (summaryEl) {
          summaryEl.innerText = shStockWarmState.ready
            ? ('Tồn Q7 sẵn sàng (' + (w.cacheSource || 'TON_Q7') + ', ' + Math.round(shStockWarmState.lastMs / 1000) + 's).')
            : 'Chưa có TON_Q7 — Admin import lại file tồn kho.';
        }
        return shStockWarmState.ready;
      });
    })
    .catch(function(err) {
      shStockWarmState.warming = false;
      shStockWarmState.promise = null;
      shStockWarmState.ready = false;
      dbgSoanLine_('warmStock.error', { error: String(err && err.message || err), via: 'direct-get' });
      if (summaryEl) summaryEl.innerText = 'Lỗi tải tồn Q7: ' + String(err && err.message || err);
      return false;
    });
  return shStockWarmState.promise;
}

function sh_chonTatCaDonSoan(checked) {
  document.querySelectorAll('.sh-order-check').forEach(function(cb) {
    cb.checked = !!checked;
  });
  sh_capNhatTomTatChonDonSoan();
}

function sh_moBangChonDonSoan() {
  if (!sessionUser || !sessionUser.user) {
    alert('Vui lòng đăng nhập trước khi tạo bảng soạn.');
    return;
  }
  sh_taiDanhSachDonSoanChoBang();
}

function sh_taoBangSoanTuDonDaChon() {
  if (!sessionUser || !sessionUser.user) {
    alert("Vui lòng đăng nhập trước khi tạo bảng soạn.");
    return;
  }
  var createDateEl = document.getElementById("sh-create-date");
  var ngay = createDateEl && createDateEl.value ? createDateEl.value : "";
  var selectedOrders = [];
  document.querySelectorAll('.sh-order-check').forEach(function(cb) {
    if (cb.checked) {
      var soPhieu = cb.getAttribute('data-sophieu') || '';
      if (soPhieu) selectedOrders.push(soPhieu);
    }
  });
  if (!selectedOrders.length) {
    alert('Vui lòng tick ít nhất 1 đơn để tạo bảng soạn.');
    return;
  }

  var _dbgSoanStart = Date.now();
  showLoad("Bước 1/2: Đang kiểm tra tồn Q7 (TON_Q7)...");
  dbgSoanLine_('taoBangSoan.start', { ngay: ngay, selectedCount: selectedOrders.length, build: APP_BUILD });

  sh_ensureStockReady_().then(function(stockOk) {
    showLoad(stockOk ? "Bước 2/2: Đang tạo bảng (đọc TON_Q7)..." : "Bước 2/2: Đang tạo bảng (rebuild TON_Q7 nếu thiếu)...");
    dbgSoanLine_('taoBangSoan.afterWarm', { stockOk: stockOk, warmMs: Date.now() - _dbgSoanStart });
    // forceStock=true nếu chưa có TON_Q7 — POST text/plain thẳng GAS để rebuild
    var payload = {
      ngay: ngay,
      actor: sessionUser.user,
      userRole: sessionUser.role || '',
      userStore: sessionUser.store || '',
      selectedOrders: selectedOrders,
      forceStock: !stockOk
    };
    var postOpts = stockOk
      ? { allowDirectFallback: true, timeoutMs: 55000 }
      : { directOnly: true, timeoutMs: 180000 };
    return apiPost('taoBangSoanHangNgayMai', payload, postOpts);
  }).then(function(res) {
    hideLoad();
    var _dbgClientMs = Date.now() - _dbgSoanStart;
    var stockStep = null;
    if (res && res._debugTimings) {
      for (var si = 0; si < res._debugTimings.length; si++) {
        if (res._debugTimings[si] && res._debugTimings[si].step === 'stockIndex') stockStep = res._debugTimings[si];
      }
    }
    dbgSoanLine_('taoBangSoan.done', { clientMs: _dbgClientMs, serverMs: res && res._debugTotalMs, run: res && res._debugRun, success: !!(res && res.success), stockReady: res && res.stockReady, stockSource: res && res.stockSource, stockStep: stockStep, build: APP_BUILD, steps: res && res._debugTimings });
    // #region agent log
    fetch('http://127.0.0.1:7480/ingest/48e8fdfc-ebb8-4d81-9aee-1659862ac812',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'4a6e3c'},body:JSON.stringify({sessionId:'4a6e3c',location:'app.js:sh_taoBangSoanTuDonDaChon',message:'taoBangSoan done',data:{success:!!(res&&res.success),stockReady:res&&res.stockReady,stockSource:res&&res.stockSource,q7Keys:stockStep&&stockStep.q7Keys,via:stockStep&&stockStep.via,serverMs:res&&res._debugTotalMs,clientMs:_dbgClientMs,run:res&&res._debugRun,missingItems:res&&res.missingItems,totalItems:res&&res.totalItems},timestamp:Date.now(),hypothesisId:'D',runId:'q7-v1'})}).catch(function(){});
    // #endregion
    if (!res || !res.success) {
      alert("❌ Tạo bảng thất bại:\n" + ((res && (res.msg || res.error)) || "Không rõ lỗi") + "\n[Build FE: " + APP_BUILD + "]" + (res && res._debugRun ? (" [GAS: " + res._debugRun + "]") : ""));
      return;
    }
    if (res.stockReady) shStockWarmState.ready = true;
    var msg = "✅ Đã tạo tab: " + (res.sheetName || "SoanNgayMai") + "\n" +
      "- Tổng đơn: " + (res.totalOrders || 0) + "\n" +
      "- Tổng mã: " + (res.totalItems || 0) + "\n" +
      "- Mã thiếu: " + (res.missingItems || 0) + "\n" +
      "- Tồn kho: " + (res.stockReady ? ("CÓ (" + (res.stockSource || "TON_Q7") + ")") : "KHÔNG — Admin import lại file tồn để tạo sheet TON_Q7");
    if (res._debugTotalMs) msg += "\n(Server tạo bảng: " + Math.round(res._debugTotalMs / 1000) + "s)";
    msg += "\n(Tổng chờ: " + Math.round(_dbgClientMs / 1000) + "s)\n[q7-v1 / " + APP_BUILD + "]";
    alert(msg);
    if (res.url) window.open(res.url, '_blank', 'noopener,noreferrer');
  }).catch(function(err) {
    hideLoad();
    dbgSoanLine_('taoBangSoan.error', { clientMs: Date.now() - _dbgSoanStart, error: String(err && err.message || err), build: APP_BUILD });
    alert('Lỗi: ' + (err && err.message || err));
  });
}

// ================= ADMIN: QUẢN LÝ TÀI KHOẢN =================
function checkAdminRole() { renderAdminStoreDropdown(); }
function loadDSUser() {
  showLoad("Đang tải...");
  apiGet('getDanhSachTaiKhoan').then(function(users) {
    hideLoad(); var tb = document.getElementById("adm-table-users"); tb.innerHTML = "";
    (users || []).forEach(function(u) {
      var storeLabel = storeMap[u.store] || u.store;
      tb.insertAdjacentHTML('beforeend', '<tr><td><b>'+escapeHtml(u.user)+'</b></td><td>'+escapeHtml(u.role)+'</td><td>'+escapeHtml(storeLabel)+'</td></tr>');
    });
  }).catch(function(err){ hideLoad(); alert('Lỗi: '+err.message); });
}
function taoTaiKhoan() {
  var role = document.getElementById("adm-role").value;
  var store = document.getElementById("adm-store").value;
  var payload = { user: document.getElementById("adm-user").value.trim(), pass: document.getElementById("adm-pass").value.trim(), role: role, store: store, actor: sessionUser.user };
  if(!payload.user || !payload.pass) return alert("Vui lòng nhập tên đăng nhập và mật khẩu.");
  if(role !== "Admin" && (!store || store === "Tất cả")) return alert("Chi nhánh phải chọn kho quản lý cụ thể.");
  showLoad("Đang tạo...");
  apiPost('taoTaiKhoanMoi', payload).then(function(res) {
    hideLoad();
    if(res.success) {
      alert("Tạo thành công!");
      document.getElementById("adm-user").value="";
      document.getElementById("adm-pass").value="";
      renderAdminStoreDropdown();
      loadDSUser();
    } else alert(res.msg || res.error || "Không thể tạo tài khoản.");
  }).catch(function(err){ hideLoad(); alert('Lỗi: '+err.message); });
}

var importPreviewState = null;

function impNormalizeText(value) {
  return String(value || '').toLowerCase().normalize('NFKD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]+/g, ' ').trim();
}

function readImportFileMatrix(file) {
  return new Promise(function(resolve, reject) {
    if (!file) {
      reject(new Error('Vui lòng chọn file.'));
      return;
    }
    if (typeof XLSX === 'undefined') {
      reject(new Error('Chưa tải được thư viện đọc Excel. Vui lòng tải lại trang.'));
      return;
    }
    var reader = new FileReader();
    reader.onload = function(e) {
      try {
        var workbook = XLSX.read(e.target.result, { type: 'array' });
        var firstSheetName = workbook.SheetNames[0];
        if (!firstSheetName) throw new Error('File không có sheet nào.');
        var worksheet = workbook.Sheets[firstSheetName];
        var rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '', raw: false });
        resolve({ rows: rows, sheetName: firstSheetName });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = function() {
      reject(new Error('Không thể đọc file đã chọn.'));
    };
    reader.readAsArrayBuffer(file);
  });
}

function impGetHeaderRow(rows) {
  if (!rows || !rows.length) return [];
  var limit = Math.min(rows.length, 8);
  for (var i = 0; i < limit; i++) {
    var row = rows[i] || [];
    var score = row.filter(function(cell) { return impNormalizeText(cell).length > 0; }).length;
    if (score >= 3) return row;
  }
  return rows[0] || [];
}

function analyzeImportRows(rows, importType) {
  var header = impGetHeaderRow(rows);
  var normalized = header.map(impNormalizeText).filter(Boolean);
  var joined = normalized.join(' | ');
  var hasStockSignals = /ton kho|so luong ton|stock|onhand|slton|qty/.test(joined) || /kho|cua hang|chi nhanh|store/.test(joined);
  var hasCatalogSignals = /ma hang hoa|ma hang hoa cha|ma vach|ten hang hoa|don vi tinh|gia ban|thue suat/.test(joined);
  var detectedType = hasStockSignals && !hasCatalogSignals ? 'stock' : (hasCatalogSignals ? 'catalog' : 'unknown');
  var warnings = [];
  if (!rows || rows.length < 2) warnings.push('File gần như không có dữ liệu.');
  if (importType === 'stock' && detectedType === 'catalog') warnings.push('File đang giống file nhập khẩu thông tin hơn file tồn kho.');
  if (importType === 'catalog' && detectedType === 'stock') warnings.push('File đang giống file tồn kho hơn file nhập khẩu thông tin.');
  if (importType === 'stock' && !hasStockSignals) warnings.push('Không tìm thấy dấu hiệu cột tồn kho hoặc tên kho trong file.');
  if (importType === 'catalog' && !hasCatalogSignals) warnings.push('Không tìm thấy đủ dấu hiệu cột thông tin hàng hóa trong file.');
  if (normalized.length < 3) warnings.push('Dòng tiêu đề quá ít cột, có thể chọn nhầm sheet hoặc nhầm file.');
  return {
    header: header,
    detectedType: detectedType,
    warnings: warnings,
    rowCount: rows ? rows.length : 0,
    colCount: header ? header.length : 0
  };
}

function renderImportPreview(previewState) {
  var card = document.getElementById('imp-preview-card');
  var meta = document.getElementById('imp-preview-meta');
  var warning = document.getElementById('imp-preview-warning');
  var ok = document.getElementById('imp-preview-ok');
  var head = document.getElementById('imp-preview-head');
  var body = document.getElementById('imp-preview-body');
  var confirmEl = document.getElementById('imp-confirm');
  if (!card || !meta || !warning || !ok || !head || !body || !confirmEl) return;

  if (!previewState) {
    card.style.display = 'none';
    meta.innerHTML = '';
    warning.style.display = 'none';
    ok.style.display = 'none';
    head.innerHTML = '';
    body.innerHTML = '';
    confirmEl.checked = false;
    imp_toggleImportButton();
    return;
  }

  card.style.display = 'block';
  confirmEl.checked = false;
  meta.innerHTML = '<b>File:</b> ' + previewState.fileName + ' | <b>Sheet:</b> ' + previewState.sheetName + ' | <b>Dòng:</b> ' + previewState.analysis.rowCount + ' | <b>Cột:</b> ' + previewState.analysis.colCount + ' | <b>Nhận diện:</b> ' + (previewState.analysis.detectedType === 'stock' ? 'Tồn kho' : previewState.analysis.detectedType === 'catalog' ? 'Nhập khẩu thông tin' : 'Chưa rõ');

  if (previewState.analysis.warnings.length) {
    warning.style.display = 'block';
    warning.innerHTML = '<b>Cảnh báo:</b><br>' + previewState.analysis.warnings.map(function(item) { return '- ' + item; }).join('<br>');
    ok.style.display = 'none';
    ok.innerHTML = '';
  } else {
    ok.style.display = 'block';
    ok.innerHTML = 'Đã kiểm tra sơ bộ, chưa phát hiện dấu hiệu chọn nhầm file.';
    warning.style.display = 'none';
    warning.innerHTML = '';
  }

  var previewRows = previewState.rows.slice(0, 6);
  var header = previewRows[0] || [];
  head.innerHTML = '<tr>' + header.map(function(cell) {
    return '<th style="text-align:left; padding:8px; border-bottom:1px solid #e2e8f0; background:#f8fafc;">' + (cell || '') + '</th>';
  }).join('') + '</tr>';
  body.innerHTML = previewRows.slice(1).map(function(row) {
    return '<tr>' + row.map(function(cell) {
      return '<td style="padding:8px; border-bottom:1px solid #f1f5f9;">' + (cell || '') + '</td>';
    }).join('') + '</tr>';
  }).join('');
  imp_toggleImportButton();
}

function imp_toggleImportButton() {
  var btn = document.getElementById('imp-submit-btn');
  var confirmEl = document.getElementById('imp-confirm');
  var enabled = !!(btn && importPreviewState && confirmEl && confirmEl.checked);
  if (!btn) return;
  btn.disabled = !enabled;
  btn.style.opacity = enabled ? '1' : '0.6';
  btn.style.cursor = enabled ? 'pointer' : 'not-allowed';
}

function imp_handleSelectionChange() {
  var typeEl = document.getElementById('imp-type');
  var fileEl = document.getElementById('imp-file');
  var importType = typeEl && typeEl.value ? typeEl.value : 'stock';
  var file = fileEl && fileEl.files && fileEl.files[0] ? fileEl.files[0] : null;
  importPreviewState = null;
  renderImportPreview(null);
  if (!file) return;
  showLoad('Đang phân tích file để xem trước...');
  readImportFileMatrix(file).then(function(parsed) {
    importPreviewState = {
      importType: importType,
      fileName: file.name,
      sheetName: parsed.sheetName,
      rows: parsed.rows,
      analysis: analyzeImportRows(parsed.rows, importType)
    };
    hideLoad();
    renderImportPreview(importPreviewState);
  }).catch(function(err) {
    hideLoad();
    renderImportPreview(null);
    alert('Lỗi đọc file: ' + err.message);
  });
}

function importDanhMucTonKho() {
  if (!sessionUser || sessionUser.role !== "Admin") return alert("Chỉ Admin mới được phép nhập khẩu dữ liệu.");
  var typeEl = document.getElementById("imp-type");
  var fileEl = document.getElementById("imp-file");
  var importType = typeEl && typeEl.value ? typeEl.value : "stock";
  var file = fileEl && fileEl.files && fileEl.files[0] ? fileEl.files[0] : null;
  if (!file) return alert("Vui lòng chọn file từ máy tính.");
  if (!importPreviewState || importPreviewState.fileName !== file.name || importPreviewState.importType !== importType) {
    return alert('Vui lòng xem trước đúng file trước khi cập nhật.');
  }
  var confirmEl = document.getElementById('imp-confirm');
  if (!confirmEl || !confirmEl.checked) return alert('Vui lòng xác nhận sau khi xem trước dữ liệu.');
  var rowCount = (importPreviewState.rows && importPreviewState.rows.length) || 0;
  showLoad("Đang cập nhật " + rowCount + " dòng lên Google Sheet (có thể 1–3 phút)...");
  var t0 = Date.now();
  // #region agent log
  fetch('http://127.0.0.1:7480/ingest/48e8fdfc-ebb8-4d81-9aee-1659862ac812',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'4a6e3c'},body:JSON.stringify({sessionId:'4a6e3c',location:'app.js:importDanhMucTonKho',message:'import start',data:{importType:importType,rowCount:rowCount,fileName:file.name,build:APP_BUILD},timestamp:Date.now(),hypothesisId:'IMP-A',runId:'import-q7-v2'})}).catch(function(){});
  // #endregion
  Promise.resolve({ rows: importPreviewState.rows, sheetName: importPreviewState.sheetName }).then(function(parsed) {
    // File lớn: POST thẳng GAS, timeout 5 phút (không qua Vercel 60s)
    return apiPost('nhapKhauCapNhatThongTin', {
      importType: importType,
      fileName: file.name,
      sourceSheet: parsed.sheetName,
      fileData: parsed.rows,
      actor: sessionUser.user
    }, { directOnly: true, timeoutMs: 300000 });
  }).then(function(res) {
    hideLoad();
    // #region agent log
    fetch('http://127.0.0.1:7480/ingest/48e8fdfc-ebb8-4d81-9aee-1659862ac812',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'4a6e3c'},body:JSON.stringify({sessionId:'4a6e3c',location:'app.js:importDanhMucTonKho',message:'import done',data:{success:!!(res&&res.success),clientMs:Date.now()-t0,serverMs:res&&res._debugTotalMs,q7Ms:res&&res._debugQ7Ms,q7Rows:res&&res.q7Rows,updatedRows:res&&res.updatedRows,run:res&&res._debugRun,build:APP_BUILD},timestamp:Date.now(),hypothesisId:'IMP-A',runId:'import-q7-v2'})}).catch(function(){});
    // #endregion
    if (!res || !res.success) {
      alert("❌ Nhập khẩu thất bại: " + ((res && (res.error || res.msg)) || "Không rõ lỗi") + "\n[Build: " + APP_BUILD + "]");
      return;
    }
    var targetSheet = res.targetSheet || (importType === 'stock' ? 'TỔNG HỢP TỒN KHO' : 'Data_Excel');
    var msg = "✅ Cập nhật thành công!\n" +
      "- Sheet đích: " + targetSheet + "\n" +
      "- Số dòng: " + (res.updatedRows || 0) + "\n" +
      "- Số cột: " + (res.updatedCols || 0);
    if (importType === 'stock') {
      msg += "\n- Sheet TON_Q7: " + (res.q7Rows || 0) + " mã";
      if (res._debugTotalMs) msg += "\n(Server: " + Math.round(res._debugTotalMs / 1000) + "s)";
    }
    if (res.msg) msg += "\n\n" + res.msg;
    msg += "\n[" + APP_BUILD + "]";
    alert(msg);
    if (fileEl) fileEl.value = '';
    importPreviewState = null;
    renderImportPreview(null);
    if (importType === 'catalog') {
      clearCatalogLocalStorage();
      loadCatalogInBackground(true);
    } else {
      // Nếu Q7 chưa tách được, thử rebuild riêng
      if (!(res.q7Rows > 0)) {
        apiGet('rebuildTonQ7', null, { directOnly: true, timeoutMs: 180000 }).catch(function() {});
      }
      shStockWarmState.ready = !!(res.q7Rows > 0);
      apiGet('getBootstrapData').then(function(bootstrap) {
        if (bootstrap && bootstrap.success) applyBootstrapData(bootstrap);
      }).catch(function() {});
    }
  }).catch(function(err) {
    hideLoad();
    // #region agent log
    fetch('http://127.0.0.1:7480/ingest/48e8fdfc-ebb8-4d81-9aee-1659862ac812',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'4a6e3c'},body:JSON.stringify({sessionId:'4a6e3c',location:'app.js:importDanhMucTonKho',message:'import error',data:{error:String(err&&err.message||err),clientMs:Date.now()-t0,build:APP_BUILD},timestamp:Date.now(),hypothesisId:'IMP-A',runId:'import-q7-v2'})}).catch(function(){});
    // #endregion
    alert('Lỗi: ' + err.message + '\n[Build: ' + APP_BUILD + ']\nGợi ý: deploy lại code.gs (New version) rồi thử lại.');
  });
}

function doiMatKhau() {
  if (!sessionUser.user) return alert("Vui lòng đăng nhập trước khi đổi mật khẩu.");
  var currentPassword = document.getElementById("pw-current").value.trim();
  var newPassword = document.getElementById("pw-new").value.trim();
  var confirmPassword = document.getElementById("pw-confirm").value.trim();
  if (!currentPassword || !newPassword || !confirmPassword) return alert("Vui lòng điền đầy đủ các trường mật khẩu.");
  if (newPassword.length < 4) return alert("Mật khẩu mới phải có ít nhất 4 ký tự.");
  if (newPassword !== confirmPassword) return alert("Xác nhận mật khẩu mới không khớp.");
  showLoad("Đang đổi mật khẩu...");
  apiPost('doiMatKhau', { user: sessionUser.user, oldPassword: currentPassword, newPassword: newPassword }).then(function(res) {
    hideLoad();
    if (!res || !res.success) throw new Error(res && res.msg ? res.msg : "Không thể đổi mật khẩu.");
    alert("✅ Đổi mật khẩu thành công!");
    document.getElementById("pw-current").value = "";
    document.getElementById("pw-new").value = "";
    document.getElementById("pw-confirm").value = "";
  }).catch(function(err){ hideLoad(); alert('Lỗi: '+err.message); });
}

function showLoad(text) { var loadingText = getEl("loading-text"); if (loadingText) loadingText.innerText = text; var overlay = getEl("loading-overlay"); if (overlay) overlay.style.display = "flex"; }
function hideLoad() { var overlay = getEl("loading-overlay"); if (overlay) overlay.style.display = "none"; }
