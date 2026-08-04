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
  const proxyTimeoutMs = Math.min(Math.max(timeoutMs, 20000), 60000);

  if (options.directOnly) {
    return callJsonApi([GAS_EXEC_URL], plainOptions, timeoutMs);
  }

  // Một số action (tạo bảng soạn) hay dính CORS/redirect trên browser → ưu tiên proxy
  if (options.proxyFirst) {
    try {
      return await callJsonApi(['/api/gas-proxy'], proxyOptions, proxyTimeoutMs);
    } catch (proxyErr) {
      console.warn('[donhang] proxy POST failed, fallback direct', String(proxyErr && proxyErr.message || proxyErr));
      if (options.allowDirectFallback === false) throw proxyErr;
      return callJsonApi([GAS_EXEC_URL], plainOptions, timeoutMs);
    }
  }

  // Ưu tiên POST thẳng GAS (tránh proxy treo). Proxy chỉ fallback.
  try {
    return await callJsonApi([GAS_EXEC_URL], plainOptions, timeoutMs);
  } catch (directErr) {
    const msg = String(directErr && directErr.message || directErr);
    console.warn('[donhang] direct POST failed, fallback proxy', msg);
    if (options.allowDirectFallback === false && options.allowProxyFallback === false) throw directErr;
    return callJsonApi(['/api/gas-proxy'], proxyOptions, proxyTimeoutMs);
  }
}

function showLoginError(message) {
  alert(message || 'Đăng nhập thất bại.');
}

// --- App logic (extracted from original webapp) ---
var APP_BUILD = '2026-08-04-v66-misa-upsert';
var shCreateDateUserTouched_ = false;
// Debug: không POST localhost (trình duyệt user không có ingest → ERR_CONNECTION_REFUSED)
var DEBUG_INGEST_ENABLED = false;
var shStockWarmState = { ready: false, warming: false, lastMs: 0, promise: null };
var packingTimelineTimer = null;
console.warn('[donhang] build', APP_BUILD);

function pad2_(n) { return String(n).padStart ? String(n).padStart(2, '0') : ((n < 10 ? '0' : '') + n); }
function formatDateVN_(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return '';
  return pad2_(d.getDate()) + '/' + pad2_(d.getMonth() + 1) + '/' + d.getFullYear();
}
function formatDateTimeVN_(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return '';
  return formatDateVN_(d) + ' ' + pad2_(d.getHours()) + ':' + pad2_(d.getMinutes()) + ':' + pad2_(d.getSeconds());
}
function startOfLocalDay_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 0, 0, 0, 0);
}
function addDaysLocal_(d, days) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate() + days, 0, 0, 0, 0);
}
function atLocalTime_(dayDate, hh, mm, ss) {
  return new Date(dayDate.getFullYear(), dayDate.getMonth(), dayDate.getDate(), hh, mm || 0, ss || 0, 0);
}

/** Chuẩn hóa Date/string về ms local (tránh parse UTC từ chuỗi ISO). */
function toLocalMillis_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value.getTime();
  if (typeof value === 'number' && isFinite(value)) return value;
  if (value === null || value === undefined || value === '') return NaN;
  var s = String(value).trim();
  var mIso = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (mIso) {
    return new Date(+mIso[1], +mIso[2] - 1, +mIso[3], +(mIso[4] || 0), +(mIso[5] || 0), +(mIso[6] || 0), 0).getTime();
  }
  var mVn = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (mVn) {
    return new Date(+mVn[3], +mVn[2] - 1, +mVn[1], +(mVn[4] || 0), +(mVn[5] || 0), +(mVn[6] || 0), 0).getTime();
  }
  var d = new Date(s);
  return isNaN(d.getTime()) ? NaN : d.getTime();
}

/**
 * Cửa sổ ca theo ngày tổng hợp N2:
 * - Chính:     N1 10:00:00 ≤ t < N2 08:00:00
 * - Bổ sung:   N2 08:00:00 ≤ t < N2 10:00:00
 * - Tổng hợp:  N1 10:00:00 ≤ t < N2 10:00:00
 * Đúng mốc 10:00:00 N2 thuộc ngày tổng hợp hôm sau.
 */
function getPackingWindowsForDay_(packingDay) {
  var n2 = startOfLocalDay_(packingDay);
  var n1 = addDaysLocal_(n2, -1);
  var startMs = atLocalTime_(n1, 10, 0, 0).getTime();
  var midMs = atLocalTime_(n2, 8, 0, 0).getTime();
  var endMs = atLocalTime_(n2, 10, 0, 0).getTime();
  return {
    n1: n1,
    n2: n2,
    startMs: startMs,
    midMs: midMs,
    endMs: endMs,
    mainStart: new Date(startMs),
    mainEnd: new Date(midMs),
    suppEnd: new Date(endMs)
  };
}

function getPackingTimelineInfo(now) {
  now = now instanceof Date ? now : new Date();
  var nowMs = now.getTime();
  var today = startOfLocalDay_(now);
  // >= 10:00 → ngày tổng hợp = ngày mai
  var packingDay = nowMs >= atLocalTime_(today, 10, 0, 0).getTime() ? addDaysLocal_(today, 1) : today;
  var win = getPackingWindowsForDay_(packingDay);
  var isMain = nowMs >= win.startMs && nowMs < win.midMs;
  var isSupp = nowMs >= win.midMs && nowMs < win.endMs;
  // Sau 10h đã chuyển packingDay=mai → luôn isMain của ca mai
  var bucket = isMain ? 'main' : 'supp';
  var packingLabel = formatDateVN_(packingDay);
  var title = isMain
    ? ('Đơn lưu lúc này thuộc ĐỢT CHÍNH — ngày giao/tổng hợp ' + packingLabel)
    : ('Đơn lưu lúc này thuộc ĐỢT BỔ SUNG — ngày giao/tổng hợp ' + packingLabel);
  var shipMsg = isMain
    ? ('Sẽ được tổng hợp soạn trong đợt chính (≥10:00 hôm trước & <08:00 hôm nay). Chi nhánh nhận sau khi kho soạn xong ngày ' + packingLabel + '.')
    : ('Sẽ đi cùng bảng bổ sung (≥08:00 & <10:00 ngày ' + packingLabel + '). Chi nhánh nhận sau khi soạn bổ sung.');
  var windowMsg = 'Chính: ' + formatDateTimeVN_(win.mainStart) + ' → ' + formatDateTimeVN_(win.mainEnd) +
    ' · Bổ sung: ' + formatDateTimeVN_(win.mainEnd) + ' → ' + formatDateTimeVN_(win.suppEnd) +
    ' · ≥10:00 thuộc ngày hôm sau.';
  var confirmText = title + '\n\n' + shipMsg + '\n\n' + windowMsg + '\n\nBạn có muốn tiếp tục lưu?';
  return {
    now: now,
    packingDay: packingDay,
    packingLabel: packingLabel,
    bucket: bucket,
    isMain: isMain,
    isSupp: isSupp,
    title: title,
    shipMsg: shipMsg,
    windowMsg: windowMsg,
    confirmText: confirmText,
    shortHtml: '<strong>' + title + '</strong>' + shipMsg + '<small>' + windowMsg + '</small>'
  };
}

function sh_getPackingMode_() {
  var checked = document.querySelector('input[name="sh-packing-mode"]:checked');
  if (checked && checked.value) return String(checked.value);
  var legacy = document.getElementById('sh-only-new-items');
  return legacy && legacy.checked ? 'supp' : 'total';
}

function sh_onPackingModeChange() {
  var mode = sh_getPackingMode_();
  var opts = document.getElementById('sh-only-new-items-opts');
  if (opts) opts.style.display = mode === 'supp' ? 'block' : 'none';
  // Reload danh sách theo mode đang chọn
  var picker = document.getElementById('sh-order-picker');
  if (picker && picker.style.display !== 'none') sh_taiDanhSachDonSoanChoBang();
}

function paintTimelineBanner_(el, info) {
  if (!el || !info) return;
  el.className = 'timeline-banner ' + (info.bucket === 'main' ? 'main' : 'supp');
  el.innerHTML = info.shortHtml;
  el.style.display = 'block';
}

function refreshPackingTimelineBanners() {
  var info = getPackingTimelineInfo(new Date());
  paintTimelineBanner_(document.getElementById('packing-timeline-banner'), info);
  paintTimelineBanner_(document.getElementById('create-timeline-banner'), info);
  var qlBanner = document.getElementById('ql-edit-timeline-banner');
  var qlView = document.getElementById('ql-view-phieu');
  if (qlBanner && qlView && qlView.style.display !== 'none') {
    var editInfo = getPackingTimelineInfo(new Date());
    editInfo.title = (editInfo.isMain ? 'Thêm/sửa mã lúc này vào ĐỢT CHÍNH' : 'Thêm/sửa mã lúc này vào ĐỢT BỔ SUNG') +
      ' — ngày giao/tổng hợp ' + editInfo.packingLabel;
    editInfo.shipMsg = editInfo.isMain
      ? 'Dòng mới/sửa sẽ được gom vào bảng tổng hợp chính (trước 8h).'
      : 'Dòng mới (Thêm mới vào đơn) sẽ vào bảng bổ sung 8h–10h.';
    editInfo.shortHtml = '<strong>' + editInfo.title + '</strong>' + editInfo.shipMsg + '<small>' + editInfo.windowMsg + '</small>';
    paintTimelineBanner_(qlBanner, editInfo);
  }
  return info;
}

function confirmPackingTimelineAction_(actionLabel) {
  var info = getPackingTimelineInfo(new Date());
  var text = (actionLabel ? (actionLabel + '\n\n') : '') + info.confirmText;
  return window.confirm(text);
}

function updateHeroGuideForRole() {
  var list = document.getElementById('hero-guide-list');
  if (!list) return;
  var role = sessionUser && sessionUser.role ? String(sessionUser.role).trim() : '';
  // Hero chỉ map chức năng — bước thao tác nằm trong .tab-guide của từng tab
  if (role === 'Admin') {
    list.innerHTML =
      '<li><b>Tạo Đơn / Quản Lý</b> — tạo &amp; điều hành đơn toàn hệ thống.</li>' +
      '<li><b>Bán kèm DV</b> — xuất bán kèm hóa đơn phần mềm khác.</li>' +
      '<li><b>Soạn Hàng</b> — tổng hợp &amp; soạn tại kho xuất.</li>' +
      '<li><b>Tài Khoản</b> — tạo user, gán kho chi nhánh.</li>';
  } else {
    list.innerHTML =
      '<li><b>Tạo Đơn</b> — đặt hàng / điều chuyển nội bộ.</li>' +
      '<li><b>Bán kèm DV</b> — xuất bán kèm hóa đơn phần mềm khác.</li>' +
      '<li><b>Xác Nhận</b> — xác nhận nhận hàng chi nhánh.</li>' +
      '<li><b>Quản Lý / Tổng Quan</b> — tra cứu &amp; theo dõi đơn.</li>';
  }
}

// #region agent log
function dbgSend_(hypothesisId, location, message, data) {
  try {
    var payload = {
      sessionId: '4a6e3c',
      hypothesisId: hypothesisId || '?',
      location: location || 'app.js',
      message: message || '',
      data: Object.assign({ build: APP_BUILD }, data || {}),
      timestamp: Date.now()
    };
    if (typeof console !== 'undefined' && console.warn) {
      console.warn('[donhang:debug]', hypothesisId || '?', location || '', message || '', payload.data);
    }
    if (!DEBUG_INGEST_ENABLED) return;
    fetch('http://127.0.0.1:7480/ingest/48e8fdfc-ebb8-4d81-9aee-1659862ac812', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'X-Debug-Session-Id': '4a6e3c' },
      body: JSON.stringify(payload)
    }).catch(function() {});
  } catch (e) {}
}
function dbgConfirm_(hypothesisId, location, message, data) {
  dbgSend_(hypothesisId, location, message, data);
}
function dbgStatusBag_(status) {
  var s = String(status || '');
  var codes = [];
  for (var i = 0; i < Math.min(s.length, 24); i++) codes.push(s.charCodeAt(i));
  return { raw: s, len: s.length, codes: codes };
}
// #endregion
(function() {
  var el = document.getElementById('app-build-tag');
  if (el) el.textContent = 'build: ' + APP_BUILD;
})();
var danhMucGoc = {}; var danhMucArr = []; var arrItems = []; var gStores = [];
var storeMap = {};
var newProductsList = [];
var newProductKeySet = {};
var variantIndexByParent = {};
var variantPickerState = { open: false, mode: 'order', parentSku: '', qty: 1, variants: [] };
/** Snapshot dòng đơn vừa tạo — dùng In/Excel sau khi clear arrItems */
var lastCreatedOrderItems_ = [];
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

/** Tra cứu sản phẩm trong Data_Excel (danhMucGoc đã load) */
function lookupCatalogProductClient_(maHang, maVach) {
  var mv = String(maVach || '').trim().toUpperCase();
  var mh = String(maHang || '').trim().toUpperCase();
  if (mv && danhMucGoc[mv]) return danhMucGoc[mv];
  if (mh && danhMucGoc[mh]) return danhMucGoc[mh];
  return null;
}

/** ĐVT bắt buộc theo Data_Excel khi catalog có giá trị — không mặc định giả "Cái" */
function resolveDvtClient_(maHang, maVach, currentDvt) {
  var cat = lookupCatalogProductClient_(maHang, maVach);
  var catalogDvt = cat && String(cat.dvt || '').trim();
  if (catalogDvt) return catalogDvt;
  return String(currentDvt || '').trim();
}

function formatDvtDisplay_(maHang, maVach, currentDvt) {
  return resolveDvtClient_(maHang, maVach, currentDvt) || '—';
}

/** Cảnh báo tồn thấp — chỉ hiển thị, không chặn lên đơn / lưu soạn */
function stockShortageInfo_(stock, needQty) {
  if (stock === '' || stock === null || stock === undefined) return null;
  var s = Number(stock);
  var n = Number(needQty);
  if (isNaN(s) || isNaN(n) || n <= 0) return null;
  if (s >= n) return null;
  return { stock: s, need: n, thieu: Math.round((n - s) * 1000) / 1000 };
}

function formatStockCellHtml_(stock, needQty) {
  var short = stockShortageInfo_(stock, needQty);
  if (stock === '' || stock === null || stock === undefined) {
    return '<span style="color:#94a3b8;">-</span>';
  }
  if (!short) {
    return '<b style="color:#166534;">' + stock + '</b>';
  }
  return '<b style="color:#b91c1c;">' + stock + '</b><br><small style="color:#b91c1c;font-weight:700;">⚠️ THIẾU ' + short.thieu + '</small>';
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
  rebuildVariantIndex_();
}

function rebuildVariantIndex_() {
  variantIndexByParent = {};
  danhMucArr.forEach(function(item) {
    if (!item) return;
    var p = String(item.parentSku || '').trim().toUpperCase();
    if (!p) return;
    if (!variantIndexByParent[p]) variantIndexByParent[p] = [];
    var dedupe = String(item.maHang || '').trim().toUpperCase() + '|' + String(item.maVach || '').trim().toUpperCase();
    var exists = variantIndexByParent[p].some(function(v) {
      return String(v.maHang || '').trim().toUpperCase() + '|' + String(v.maVach || '').trim().toUpperCase() === dedupe;
    });
    if (!exists) variantIndexByParent[p].push(item);
  });
}

function getParentSkuOfItem_(it) {
  if (!it) return '';
  var p = String(it.parentSku || it.Parent_SKU || '').trim().toUpperCase();
  if (p) return p;
  var mh = String(it.maHang || it.variantSku || '').trim().toUpperCase();
  var mv = String(it.maVach || '').trim().toUpperCase();
  if (mv && danhMucGoc[mv] && danhMucGoc[mv].parentSku) {
    return String(danhMucGoc[mv].parentSku).trim().toUpperCase();
  }
  if (mh && danhMucGoc[mh] && danhMucGoc[mh].parentSku) {
    return String(danhMucGoc[mh].parentSku).trim().toUpperCase();
  }
  return '';
}

function getVariantGroupForItem_(it) {
  var p = getParentSkuOfItem_(it);
  if (!p) return [];
  return variantIndexByParent[p] || [];
}

function itemHasVariantGroup_(it) {
  // Có Parent_SKU → mở panel chọn biến thể (kèm tồn), kể cả nhóm 1 mã
  return !!getParentSkuOfItem_(it);
}

/** Mã hiển thị trên kệ/bao bì: ưu tiên Parent_SKU */
function resolveDisplaySku_(it) {
  if (!it) return '';
  if (it.maHangDisplay) return String(it.maHangDisplay).trim();
  var parent = getParentSkuOfItem_(it);
  var child = String(it.maHang || it.variantSku || '').trim();
  if (parent && parent !== child.toUpperCase()) return parent;
  return child;
}

/** Tên + phân loại: [Tên cha] - [Biến thể] (Mã con: SKU) */
function formatPackingProductName_(it) {
  if (!it) return '-';
  if (it.tenHangDisplay) return String(it.tenHangDisplay).trim();
  var parentSku = getParentSkuOfItem_(it);
  var childSku = String(it.variantSku || it.maHang || '').trim();
  var variantName = String(it.variantName || it.tenHang || '').trim();
  var parentName = String(it.parentName || '').trim();
  if (!parentName && parentSku && danhMucGoc[parentSku]) {
    parentName = String(danhMucGoc[parentSku].tenHang || '').trim();
  }
  if (parentSku && childSku && parentSku !== childSku.toUpperCase()) {
    var base = parentName ? (parentName + ' - ' + variantName) : variantName;
    return base + ' (Mã con: ' + childSku + ')';
  }
  return formatVariantDisplayName_(childSku, variantName);
}

/** HTML ô mã SP: Parent nổi bật + mã con phụ */
function formatPackingCodeHtml_(it) {
  var parent = resolveDisplaySku_(it);
  var child = String(it.maHang || it.variantSku || '').trim();
  var mv = String(it.maVach || '').trim();
  var html = '<b style="font-size:15px;color:#0f172a;">' + escapeHtml(parent || '-') + '</b>';
  if (mv) html += '<br><small style="color:#64748b;">MV: ' + escapeHtml(mv) + '</small>';
  if (child && parent && child.toUpperCase() !== String(parent).toUpperCase()) {
    html += '<br><small style="color:#7c3aed;">Mã con: ' + escapeHtml(child) + '</small>';
  } else if (child && !parent) {
    html += '<br><small style="color:gray;">Mã hàng hóa: ' + escapeHtml(child) + '</small>';
  }
  return html;
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
  var elXb = getEl('xb-chi-nhanh'); if (elXb) elXb.innerHTML = htmlStores;
  xb_applyStoreQuyen_();
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
  if (catalogLoadState.loading && !forceReload) return Promise.resolve();
  var cached = !forceReload ? readCatalogFromLocalStorage(expectedVersion) : null;
  if (cached && cached.danhMuc) {
    applyCatalogData(cached);
    setCatalogStatus('Quét mã vạch, gõ mã, từ khóa tên hoặc 6 số cuối vạch:');
    return Promise.resolve(cached);
  }

  catalogLoadState.loading = true;
  setCatalogStatus('Đang tải danh mục hàng ở nền...');
  var params = forceReload ? { nocache: '1', _ts: String(Date.now()) } : null;
  return apiGet('getCatalogData', params, { timeoutMs: 180000, allowDirectFallback: true }).then(function(res) {
    catalogLoadState.loading = false;
    if (!res || !res.success) {
      setCatalogStatus('Danh mục chưa tải xong - thử lại sau vài giây');
      return res;
    }
    applyCatalogData(res);
    saveCatalogToLocalStorage(res);
    var count = danhMucArr ? danhMucArr.length : 0;
    setCatalogStatus('Quét / gõ tên (có hoặc không dấu) · ' + count + ' mã sẵn sàng');
    return res;
  }).catch(function(err) {
    catalogLoadState.loading = false;
    setCatalogStatus('Lỗi tải danh mục: ' + err.message);
    throw err;
  });
}

/** Xóa cache local + gọi getCatalogData?nocache=1 (Data_Excel + mã con từ TON_VARIANT) */
function reloadCatalogNow_() {
  clearCatalogLocalStorage();
  showLoad('Đang tải lại danh mục (Data_Excel + TON_VARIANT)...');
  return loadCatalogInBackground(true).then(function(res) {
    hideLoad();
    var n = danhMucArr ? danhMucArr.length : 0;
    if (!res || !res.success) {
      alert('Không tải được danh mục: ' + ((res && res.error) || 'unknown'));
      return res;
    }
    alert('✅ Đã tải lại danh mục: ' + n + ' sản phẩm (gồm mã cha/con từ TON_VARIANT).\nVersion: ' + (res.version || catalogLoadState.version || '-') + '.');
    var input = document.getElementById('input-scan');
    if (input) {
      input.focus();
      if (input.value && input.value.trim()) handleSearchInput({ key: '' });
    }
    return res;
  }).catch(function(err) {
    hideLoad();
    alert('Lỗi tải lại danh mục: ' + (err && err.message || err));
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
  applyNewProductsData(res.newProducts || []);
}

function applyNewProductsData(list) {
  newProductsList = Array.isArray(list) ? list : [];
  newProductKeySet = {};
  for (var i = 0; i < newProductsList.length; i++) {
    var it = newProductsList[i];
    if (!it) continue;
    var mh = String(it.maHang || '').trim().toUpperCase();
    var mv = String(it.maVach || '').trim().toUpperCase();
    if (mh) newProductKeySet[mh] = true;
    if (mv) newProductKeySet[mv] = true;
  }
  renderNewProductsHighlight();
}

function isNewProductItem_(item) {
  if (!item) return false;
  if (item.isNew === true || item.isNew === 1 || item.isNew === '1' || item.isNew === 'true') return true;
  var mh = String(item.maHang || '').trim().toUpperCase();
  var mv = String(item.maVach || '').trim().toUpperCase();
  if (newProductKeySet[mh] || newProductKeySet[mv]) return true;
  var cat = (mh && danhMucGoc[mh]) || (mv && danhMucGoc[mv]) || null;
  return !!(cat && (cat.isNew === true || cat.isNew === 1 || cat.isNew === 'TRUE' || cat.isNew === 'true'));
}

function renderNewProductsHighlight() {
  // Chỉ 1 block UI duy nhất: #new-products-strip (Trang chủ / Đặt hàng)
  var strip = document.getElementById('new-products-strip');
  var grid = document.getElementById('new-products-grid');
  if (!newProductsList.length) {
    if (strip) strip.style.display = 'none';
    if (grid) grid.innerHTML = '';
    return;
  }

  var cardsHtml = newProductsList.map(function(item, idx) {
    var title = escapeHtml(item.tenHang || '(Không tên)');
    var mh = escapeHtml(item.maHang || '-');
    var mv = escapeHtml(item.maVach || '-');
    var dvt = escapeHtml(item.dvt || 'Cái');
    var ngay = escapeHtml(item.ngayTao || '');
    var reasonBadge = '<span style="display:inline-block;padding:2px 7px;border-radius:999px;font-size:10px;font-weight:800;background:#dbeafe;color:#1d4ed8;">ADMIN CHỌN</span>';
    var payload = encodeURIComponent(JSON.stringify({
      maHang: item.maHang || '',
      maVach: item.maVach || '',
      tenHang: item.tenHang || '',
      dvt: item.dvt || ''
    }));
    return '' +
      '<button type="button" class="new-product-card" onclick="addNewProductToOrder(\'' + payload + '\')">' +
        '<div style="display:flex; justify-content:space-between; gap:8px; align-items:center; flex-wrap:wrap;">' +
          '<span class="badge-new">MỚI #' + (item.rank || (idx + 1)) + '</span>' +
          reasonBadge +
        '</div>' +
        '<div class="np-title">' + title + '</div>' +
        '<div class="np-meta">MH: ' + mh + ' · MV: ' + mv + ' · ĐVT: ' + dvt + (ngay ? (' · ' + ngay) : '') + '</div>' +
        '<div class="np-actions"><span class="np-btn">➕ Thêm vào đơn</span></div>' +
      '</button>';
  }).join('');

  if (grid) grid.innerHTML = cardsHtml;
  if (strip) strip.style.display = 'block';
}

function addNewProductToOrder(encodedItem) {
  try {
    var item = JSON.parse(decodeURIComponent(encodedItem));
    switchTab('tab-tao-phieu');
    chonSanPham(item);
    var input = document.getElementById('input-scan');
    if (input) {
      input.value = '';
      input.focus();
    }
    var box = document.getElementById('suggest-box');
    if (box) box.style.display = 'none';
  } catch (e) {
    alert('Không thể thêm sản phẩm mới: ' + (e.message || e));
  }
}

function loadNewProductsInBackground() {
  apiGet('getNewProductsList', { limit: 10 }, { allowDirectFallback: true, timeoutMs: 45000 })
    .then(function(res) {
      if (res && res.success && Array.isArray(res.data)) applyNewProductsData(res.data);
    })
    .catch(function() {});
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
var INVENTORY_APP_URL = "https://kiem-ke-app-mauve.vercel.app/";

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

  var todayStr = formatDateInputValue(new Date());
  var qlNgay = getEl("ql-ngay"); if (qlNgay && !qlNgay.value) qlNgay.value = todayStr;
  var confirmNgay = getEl("confirm-ngay"); if (confirmNgay && !confirmNgay.value) confirmNgay.value = todayStr;
  var shNgay = getEl("sh-ngay"); if (shNgay && !shNgay.value) shNgay.value = todayStr;
  sh_ensureCreateDateDefaults_();
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
    if (!res.newProducts || !res.newProducts.length) loadNewProductsInBackground();
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
  updateHeroGuideForRole();
  refreshPackingTimelineBanners();
  if (packingTimelineTimer) clearInterval(packingTimelineTimer);
  packingTimelineTimer = setInterval(refreshPackingTimelineBanners, 60000);
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
  refreshPackingTimelineBanners();
  if(tabId === 'tab-quan-ly') ql_loadPhieu();
  if(tabId === 'tab-xac-nhan') confirm_loadPhieu();
  if(tabId === 'tab-soan-hang') sh_taiDanhSachDon();
  if(tabId === 'tab-dashboard') loadDashboardSummary();
  if(tabId === 'tab-admin') {
    loadDSUser();
    if (sessionUser.role === 'Admin') adm_loadNewProductFlags();
  }
  if(tabId === 'tab-ban-kem') xb_onTabOpen();
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
      var kx = formatStoreShortLabel_(order.khoXuat);
      var kn = formatStoreShortLabel_(order.khoNhan);
      var so = String(order.soPhieu || '');
      var soAttr = encodeURIComponent(so);
      return '<tr><td><a href="javascript:void(0)" class="order-pdf-link" onclick="openOrderPdfView(\'' + soAttr + '\')">' + escapeHtml(so) + '</a></td><td>' + kx + '</td><td>' + kn + '</td><td><span style="display:inline-block; padding:4px 8px; border-radius:999px; background:#eff6ff; color:#1d4ed8; font-size:12px; font-weight:700;">' + order.status + '</span></td><td>' + (order.thoiGian || '-') + '</td></tr>';
    }).join('');

    recent.innerHTML = '<div style="overflow-x:auto;"><table style="width:100%; border-collapse:collapse;"><thead><tr><th style="text-align:left; padding:8px; border-bottom:1px solid #e2e8f0;">Số phiếu</th><th style="text-align:left; padding:8px; border-bottom:1px solid #e2e8f0;">Kho xuất</th><th style="text-align:left; padding:8px; border-bottom:1px solid #e2e8f0;">Kho nhận</th><th style="text-align:left; padding:8px; border-bottom:1px solid #e2e8f0;">Trạng thái</th><th style="text-align:left; padding:8px; border-bottom:1px solid #e2e8f0;">Cập nhật</th></tr></thead><tbody>' + rows + '</tbody></table></div>';
  }).catch(function(err){ hideLoad(); console.error(err); });
}

/** Tên ngắn kho từ storeMap (vd. K9 Quận 7); fallback tên gốc */
function formatStoreShortLabel_(storeName) {
  var raw = String(storeName || '').trim();
  if (!raw) return '-';
  if (storeMap && storeMap[raw]) return storeMap[raw];
  // Đôi khi API trả tên ngắn sẵn — giữ nguyên
  for (var full in storeMap) {
    if (!Object.prototype.hasOwnProperty.call(storeMap, full)) continue;
    if (storeMap[full] === raw) return raw;
  }
  return raw;
}

/** Mã ngắn phân biệt 2 kho Q4 — tránh dính chữ "Mới/Cũ" với cột Trạng thái */
var PACKING_STORE_CODES_ = {
  'Kho Địa điểm kinh doanh 01': '178',
  'Kho Địa điểm kinh doanh 06': '275'
};

/** Ô kho trên bảng xem trước soạn hàng — tách sạch khỏi trạng thái đơn */
function formatPackingStoreCellHtml_(storeName) {
  var raw = String(storeName || '').trim();
  if (!raw) return '-';
  var code = PACKING_STORE_CODES_[raw] || '';
  var short = formatStoreShortLabel_(raw);
  var cleaned = String(short).replace(/\s+(Mới|Cũ|Moi|Cu)\s*$/i, '').trim() || short;
  var hint = '';
  if (code) hint = code;
  else if (/mới|moi/i.test(short) && cleaned !== short) hint = 'CN mới';
  else if (/cũ|cu/i.test(short) && cleaned !== short) hint = 'CN cũ';
  if (hint) {
    return '<span class="sh-store-name">' + escapeHtml(cleaned) + '</span>' +
      ' <small class="sh-store-hint" style="color:#64748b;font-weight:600;">· ' + escapeHtml(hint) + '</small>';
  }
  return '<span class="sh-store-name">' + escapeHtml(cleaned) + '</span>';
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
  xb_applyStoreQuyen_();
}

// ================= TÌM KIẾM & TẠO PHIẾU =================
/** Chuẩn hóa mã MISA / mã hàng / số HĐ — GIỮ chữ Đ/đ (HĐC… ≠ HDC…) */
function normalizeMisaCode_(value) {
  return String(value || '')
    .normalize('NFC')
    .replace(/[^a-zA-Z0-9Đđ]/g, '')
    .trim()
    .toUpperCase();
}

function normalizeProductCodeClient_(value) {
  return normalizeMisaCode_(value);
}

/** NFC + lower — nền tảng so khớp */
function toSearchNfcLower_(value) {
  try {
    return String(value == null ? '' : value).normalize('NFC').toLowerCase().trim();
  } catch (e) {
    return String(value == null ? '' : value).toLowerCase().trim();
  }
}

/**
 * Form giữ đ (bỏ dấu tổ hợp á→a, ư→u…) — khớp tên có dấu đã gõ / mã có Đ.
 * VD: "Sữa Tắm" → "sua tam", "HĐC1007" → "hđc1007"
 */
function normalizeSearchText(value) {
  return toSearchNfcLower_(value)
    .normalize('NFKD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9đ\s]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Form không dấu hoàn toàn (đ→d) — gõ "sua tam" / "duong" vẫn ra "sữa tắm" / "đường".
 */
function foldSearchText_(value) {
  return normalizeSearchText(value).replace(/đ/g, 'd');
}

function getSearchFields_(item) {
  if (!item) return [];
  return [
    item.maHang,
    item.maVach,
    item.tenHang,
    item.parentSku,
    item.Parent_SKU,
    item.dvt
  ].filter(function(v) { return v != null && String(v).trim() !== ''; });
}

function tokensAllMatch_(haystack, tokens) {
  if (!tokens || !tokens.length) return false;
  for (var i = 0; i < tokens.length; i++) {
    if (!tokens[i] || haystack.indexOf(tokens[i]) === -1) return false;
  }
  return true;
}

function scoreFieldAgainstQuery_(field, qKeep, qFold, qRawNfc, qCode) {
  if (field == null || field === '') return 0;
  var score = 0;
  var raw = toSearchNfcLower_(field);
  var keep = normalizeSearchText(field);
  var fold = foldSearchText_(field);
  var code = normalizeMisaCode_(field);

  // Khớp mã chính xác / prefix (giữ Đ)
  if (qCode && code) {
    if (code === qCode) score = Math.max(score, 1200);
    else if (code.indexOf(qCode) === 0) score = Math.max(score, 700);
    else if (qCode.length >= 6 && code.slice(-6) === qCode.slice(-6)) score = Math.max(score, 500);
    else if (code.indexOf(qCode) !== -1) score = Math.max(score, 320);
  }

  // Khớp có dấu / giữ đ
  if (qKeep) {
    if (keep === qKeep || raw === qRawNfc) score = Math.max(score, 1000);
    else if (keep.indexOf(qKeep) === 0 || raw.indexOf(qRawNfc) === 0) score = Math.max(score, 650);
    else if (keep.indexOf(qKeep) !== -1 || raw.indexOf(qRawNfc) !== -1) score = Math.max(score, 280);
    else {
      var partsKeep = qKeep.split(/\s+/).filter(Boolean);
      if (partsKeep.length > 1 && tokensAllMatch_(keep, partsKeep)) score = Math.max(score, partsKeep.length * 90);
    }
  }

  // Khớp không dấu (đ→d)
  if (qFold) {
    if (fold === qFold) score = Math.max(score, 950);
    else if (fold.indexOf(qFold) === 0) score = Math.max(score, 600);
    else if (fold.indexOf(qFold) !== -1) score = Math.max(score, 260);
    else {
      var partsFold = qFold.split(/\s+/).filter(Boolean);
      if (partsFold.length > 1 && tokensAllMatch_(fold, partsFold)) score = Math.max(score, partsFold.length * 85);
    }
  }

  return score;
}

function getSearchScore(item, rawQuery) {
  var qRaw = String(rawQuery == null ? '' : rawQuery).trim();
  if (!qRaw) return 0;
  var qRawNfc = toSearchNfcLower_(qRaw);
  var qKeep = normalizeSearchText(qRaw);
  var qFold = foldSearchText_(qRaw);
  var qCode = normalizeMisaCode_(qRaw);
  if (!qKeep && !qFold && !qCode) return 0;

  var fields = getSearchFields_(item);
  var best = 0;
  for (var i = 0; i < fields.length; i++) {
    var s = scoreFieldAgainstQuery_(fields[i], qKeep, qFold, qRawNfc, qCode);
    if (s > best) best = s;
  }
  // Bonus nhẹ khi khớp cả cụm trên tên
  if (item && item.tenHang) {
    var nameFold = foldSearchText_(item.tenHang);
    if (qFold && nameFold.indexOf(qFold) !== -1) best += 40;
  }
  return best;
}

function handleSearchInput(e) {
  if (!catalogLoadState.ready) {
    var box = document.getElementById("suggest-box");
    if (box) {
      box.innerHTML = '<div class="suggest-empty">Danh mục hàng đang tải. Vui lòng đợi vài giây... <a href="javascript:void(0)" onclick="reloadCatalogNow_()">Tải lại</a></div>';
      box.style.display = "block";
    }
    return;
  }
  var inputEl = document.getElementById("input-scan");
  var val = inputEl.value.trim();
  var box = document.getElementById("suggest-box");
  box.style.width = (inputEl.offsetWidth) + "px"; box.style.left = (inputEl.offsetLeft) + "px"; box.style.top = (inputEl.offsetTop + inputEl.offsetHeight) + "px";

  if (val.length < 1) { box.style.display = "none"; return; }

  if (e && e.key === "Enter") {
    var exactKey = normalizeMisaCode_(val) || val.toUpperCase();
    var exactMatch = danhMucGoc[exactKey] || danhMucGoc[val.toUpperCase()] || danhMucGoc[toSearchNfcLower_(val).toUpperCase()];
    if (exactMatch) chonSanPham(exactMatch);
    else {
      var matched = filterProducts(val);
      if (matched.length > 0) chonSanPham(matched[0]);
      else { arrItems.unshift({ maHang: "LỖI MÃ", maVach: val, tenHang: "❌ Không tồn tại", dvt: "Lỗi", sl: "1" }); renderTable(); }
    }
    inputEl.value = ""; box.style.display = "none"; return;
  }

  var results = filterProducts(val);
  if (results.length === 0) {
    box.innerHTML = '<div style="padding:10px; color:#d93025; text-align:center; font-weight:600;">Không tìm thấy sản phẩm phù hợp.<br><small style="color:#64748b;font-weight:500;">Thử <a href="javascript:void(0)" onclick="reloadCatalogNow_()">tải lại danh mục</a> nếu vừa cập nhật Data_Excel.</small></div>';
    box.style.display = "block";
    return;
  }

  var html = "";
  results.slice(0, 12).forEach(function(item) {
    var itemStr = encodeURIComponent(JSON.stringify(item));
    var newBadge = isNewProductItem_(item) ? ' <span class="badge-new">MỚI</span>' : '';
    var groupBadge = itemHasVariantGroup_(item) ? ' <span class="badge-group">NHÓM</span>' : '';
    html += '<div class="suggest-item" onclick="chonSanPhamFromSuggest(\'' + itemStr + '\')"><div class="sg-title">' + escapeHtml(item.tenHang) + newBadge + groupBadge + '</div><div class="sg-desc"><span style="color:#1a73e8; font-weight:700;">Mã hàng: ' + escapeHtml(item.maHang) + '</span> · Mã vạch: ' + escapeHtml(item.maVach) + ' · ĐVT: ' + escapeHtml(formatDvtDisplay_(item.maHang, item.maVach, item.dvt)) + (item.parentSku ? (' · Parent: ' + escapeHtml(item.parentSku)) : '') + '</div></div>';
  });
  box.innerHTML = html; box.style.display = "block";
}

function filterProducts(kw) {
  var query = String(kw == null ? '' : kw).trim();
  if (!query) return [];
  if (!danhMucArr || !danhMucArr.length) return [];
  var scored = danhMucArr.map(function(it) {
    return { item: it, score: getSearchScore(it, query) };
  }).filter(function(entry) {
    return entry.score > 0;
  }).sort(function(a, b) {
    return b.score - a.score || String(a.item.tenHang || "").localeCompare(String(b.item.tenHang || ""), 'vi');
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
    var groupBadge = itemHasVariantGroup_(item) ? ' <span class="badge-group">NHÓM</span>' : '';
    html += '<div class="suggest-item" onclick="ql_pickSuggestedItem(\'' + itemStr + '\')"><div class="sg-title">' + escapeHtml(item.tenHang) + groupBadge + '</div><div class="sg-desc"><span style="color:#2563eb; font-weight:700;">Mã hàng: ' + escapeHtml(item.maHang) + '</span> · Mã vạch: ' + escapeHtml(item.maVach) + ' · ĐVT: ' + escapeHtml(formatDvtDisplay_(item.maHang, item.maVach, item.dvt)) + (item.parentSku ? (' · Parent: ' + escapeHtml(item.parentSku)) : '') + '</div></div>';
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

/** Hiển thị nhãn biến thể: `DC-BANH-DO - Banh gai màu đỏ` */
function formatVariantDisplayName_(maHang, tenHang) {
  var mh = String(maHang || '').trim();
  var th = String(tenHang || '').trim();
  if (!mh) return th || '-';
  if (!th) return mh;
  if (th.toUpperCase().indexOf(mh.toUpperCase()) === 0) return th;
  return mh + ' - ' + th;
}

function closeVariantPicker_() {
  variantPickerState.open = false;
  variantPickerState.parentSku = '';
  variantPickerState.variants = [];
  var modal = document.getElementById('modal-variant-picker');
  if (modal) modal.style.display = 'none';
}

function openVariantPicker_(seedItem, mode, qty) {
  var parentSku = getParentSkuOfItem_(seedItem);
  if (!parentSku) {
    if (mode === 'ql') ql_themMaHangDirect_(seedItem, qty);
    else if (mode === 'xb') xb_chonSanPhamDirect_(seedItem);
    else chonSanPhamDirect_(seedItem, 1);
    return;
  }
  variantPickerState.open = true;
  variantPickerState.mode = mode || 'order';
  variantPickerState.parentSku = parentSku;
  variantPickerState.qty = Number(qty) > 0 ? Number(qty) : 1;
  variantPickerState.variants = [];

  var modal = document.getElementById('modal-variant-picker');
  var sub = document.getElementById('variant-picker-sub');
  var list = document.getElementById('variant-picker-list');
  if (!modal || !list) return;
  if (sub) {
    sub.innerHTML = 'Nhóm cha <b>' + escapeHtml(parentSku) + '</b> — chọn mã con kèm <b>Ton_Hien_Tai</b> (TON_VARIANT), nhập SL rồi bấm <b>Xác nhận chọn</b>.';
  }
  list.innerHTML = '<div style="padding:16px;color:#64748b;text-align:center;">Đang tải danh sách biến thể &amp; tồn hiện tại...</div>';
  modal.style.display = 'flex';

  var localGroup = getVariantGroupForItem_(seedItem).slice();
  function renderRows_(variants) {
    if (!variants || !variants.length) {
      variantPickerState.variants = [];
      list.innerHTML = '<div style="padding:16px;color:#d93025;text-align:center;font-weight:600;">Không có biến thể trong nhóm này.</div>';
      return;
    }
    variantPickerState.variants = variants;
    var html = '<table class="variant-pick-table"><thead><tr>' +
      '<th>Tên chi tiết</th><th>Mã con (SKU)</th><th>Mã vạch</th><th>Ton_Hien_Tai</th><th style="width:80px;">SL</th>' +
      '</tr></thead><tbody>';
    variants.forEach(function(v, idx) {
      var stock = (v.tonHienTai != null && v.tonHienTai !== '') ? Number(v.tonHienTai)
        : ((v.stock === '' || v.stock === null || v.stock === undefined) ? null : Number(v.stock));
      var stockHtml = '';
      if (stock === null || isNaN(stock)) stockHtml = '<span class="badge-stock-zero">?</span>';
      else if (stock <= 0) stockHtml = '<span class="badge-stock-zero">0</span>';
      else stockHtml = '<span class="badge-stock-ok">' + stock + '</span>';
      var zeroClass = (stock !== null && !isNaN(stock) && stock <= 0) ? ' zero' : '';
      html += '<tr class="variant-pick-row' + zeroClass + '" data-vp-idx="' + idx + '">' +
        '<td style="font-weight:600;color:#0f172a;">' + escapeHtml(v.tenHang || '(Không tên)') + '</td>' +
        '<td><b style="color:#1a73e8;">' + escapeHtml(v.maHang || '-') + '</b></td>' +
        '<td>' + escapeHtml(v.maVach || '-') + '</td>' +
        '<td>' + stockHtml + '</td>' +
        '<td><input type="number" class="vp-qty" id="vp-qty-' + idx + '" min="0" step="1" value="0" inputmode="numeric" onclick="event.stopPropagation()"></td>' +
        '</tr>';
    });
    html += '</tbody></table>';
    list.innerHTML = html;
  }

  renderRows_(localGroup.map(function(v) {
    return { maHang: v.maHang, maVach: v.maVach, tenHang: v.tenHang, dvt: v.dvt, parentSku: parentSku, stock: null, tonHienTai: null };
  }));

  apiGet('getVariantStockList', { parentSku: parentSku }, { timeoutMs: 45000 }).then(function(res) {
    if (!variantPickerState.open || variantPickerState.parentSku !== parentSku) return;
    if (res && res.success && Array.isArray(res.variants) && res.variants.length) {
      renderRows_(res.variants.map(function(v) {
        v.tonHienTai = (v.tonHienTai != null ? v.tonHienTai : v.stock);
        return v;
      }));
    }
  }).catch(function() {
    // Giữ list local nếu API lỗi
  });
}

/** Gom các dòng có SL > 0 từ modal → đẩy vào đơn */
function confirmVariantPicker_() {
  var variants = variantPickerState.variants || [];
  if (!variants.length) return alert('Chưa có danh sách biến thể.');
  var picked = [];
  for (var i = 0; i < variants.length; i++) {
    var qtyEl = document.getElementById('vp-qty-' + i);
    var qty = qtyEl ? Number(qtyEl.value) : 0;
    if (!qty || qty <= 0 || isNaN(qty)) continue;
    var v = variants[i];
    picked.push({
      maHang: v.maHang || '',
      maVach: v.maVach || '',
      tenHang: v.tenHang || '',
      dvt: v.dvt || '',
      parentSku: v.parentSku || variantPickerState.parentSku || '',
      sl: qty
    });
  }
  if (!picked.length) return alert('Nhập số lượng (> 0) cho ít nhất một mã con rồi xác nhận.');

  var mode = variantPickerState.mode || 'order';
  closeVariantPicker_();
  picked.forEach(function(item) {
    if (mode === 'ql') ql_themMaHangDirect_(item, item.sl);
    else if (mode === 'xb') xb_chonSanPhamDirect_(item, item.sl);
    else chonSanPhamDirect_(item, item.sl);
  });
}

/** @deprecated giữ tương thích nếu còn gọi click-một-dòng */
function pickVariantFromPicker_(payload) {
  var item = null;
  try { item = JSON.parse(decodeURIComponent(payload)); } catch (e) { return; }
  if (!item) return;
  var mode = variantPickerState.mode || 'order';
  var qty = variantPickerState.qty || 1;
  closeVariantPicker_();
  if (mode === 'ql') ql_themMaHangDirect_(item, qty);
  else if (mode === 'xb') xb_chonSanPhamDirect_(item, qty);
  else chonSanPhamDirect_(item, qty);
}

function chonSanPhamDirect_(it, qty) {
  if (!it) return;
  var addQty = Number(qty);
  if (!addQty || addQty <= 0 || isNaN(addQty)) addQty = Number(it.sl) > 0 ? Number(it.sl) : 1;
  var existingIndex = arrItems.findIndex(function(x) {
    return x.maHang !== "LỖI MÃ" &&
      String(x.maHang || '').trim().toUpperCase() === String(it.maHang || '').trim().toUpperCase() &&
      String(x.maVach || '').trim().toUpperCase() === String(it.maVach || '').trim().toUpperCase();
  });
  if (existingIndex !== -1) {
    arrItems[existingIndex].sl = Number(arrItems[existingIndex].sl) + addQty;
    arrItems[existingIndex].highlight = true;
    if (it.parentSku) arrItems[existingIndex].parentSku = it.parentSku;
  } else {
    arrItems.unshift({
      maHang: it.maHang,
      maVach: it.maVach,
      tenHang: it.tenHang,
      dvt: resolveDvtClient_(it.maHang, it.maVach, it.dvt),
      parentSku: it.parentSku || '',
      sl: String(addQty),
      highlight: true
    });
  }
  renderTable();
}

function chonSanPham(it) {
  if (!it) return;
  if (itemHasVariantGroup_(it)) {
    openVariantPicker_(it, 'order', 1);
    return;
  }
  chonSanPhamDirect_(it);
}

document.addEventListener("click", function(event) {
  var box = document.getElementById("suggest-box");
  var input = document.getElementById("input-scan");
  var qlBox = document.getElementById("ql-suggest-box");
  var qlInput = document.getElementById("ql-add-code");
  var xbBox = document.getElementById("xb-suggest-box");
  var xbInput = document.getElementById("xb-input-scan");
  if (box && event.target !== box && event.target !== input && !box.contains(event.target)) box.style.display = "none";
  if (qlBox && event.target !== qlBox && event.target !== qlInput && !qlBox.contains(event.target)) qlBox.style.display = "none";
  if (xbBox && event.target !== xbBox && event.target !== xbInput && !xbBox.contains(event.target)) xbBox.style.display = "none";
});

function thayDoiSoLuong(index, delta) { var currentSl = Number(arrItems[index].sl) || 0; var newSl = currentSl + delta; if (newSl > 0) { arrItems[index].sl = newSl; renderTable(); } }

function renderTable() {
  var tbody = document.getElementById("tbody-items"); tbody.innerHTML = ""; var tongSl = 0;
  arrItems.forEach((it, i) => {
    var isErr = (it.maHang === "LỖI MÃ" || isNaN(Number(it.sl))); tongSl += (Number(it.sl) || 0);
    var trClass = isErr ? 'row-error' : (it.highlight ? 'scan-highlight' : ''); it.highlight = false;
    var newBadge = (!isErr && isNewProductItem_(it)) ? ' <span class="badge-new">MỚI</span>' : '';
    var groupBadge = (!isErr && getParentSkuOfItem_(it)) ? ' <span class="badge-group">Biến thể</span>' : '';
    var dvtShow = formatDvtDisplay_(it.maHang, it.maVach, it.dvt);
    it.dvt = resolveDvtClient_(it.maHang, it.maVach, it.dvt) || it.dvt;
    var tenShow = formatPackingProductName_(it);
    tbody.insertAdjacentHTML('beforeend', '<tr class="' + trClass + '"><td>' + (arrItems.length - i) + '</td><td>' + formatPackingCodeHtml_(it) + '</td><td style="font-weight:500;">' + escapeHtml(tenShow) + newBadge + groupBadge + '</td><td>' + escapeHtml(dvtShow) + '</td><td><div class="qty-control"><button class="qty-btn" onclick="thayDoiSoLuong(' + i + ', -1)">-</button><input type="number" class="qty-input" value="' + it.sl + '" onchange="arrItems[' + i + '].sl=this.value; renderTable();"><button class="qty-btn" onclick="thayDoiSoLuong(' + i + ', 1)">+</button></div></td><td style="text-align:center;"><button style="color:#d93025; border:none; background:none; font-weight:bold; cursor:pointer; font-size:18px;" onclick="arrItems.splice(' + i + ',1); renderTable();">×</button></td></tr>');
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
  refreshPackingTimelineBanners();
  if (!confirmPackingTimelineAction_('Xác nhận tạo đơn mới')) return;
  showLoad("Đang tạo đơn...");
  var lPhieu = document.querySelector('input[name="loaiPhieu"]:checked').value;
  var khoXuat = document.getElementById("select-kho-xuat").value;
  var khoNhan = document.getElementById("select-kho-nhan").value;
  var itemCount = arrItems.length;
  var itemsToSave = arrItems.map(function(it) {
    return {
      maHang: it.maHang,
      maVach: it.maVach,
      tenHang: it.tenHang,
      dvt: resolveDvtClient_(it.maHang, it.maVach, it.dvt),
      sl: it.sl,
      parentSku: it.parentSku || ''
    };
  });
  lastCreatedOrderItems_ = itemsToSave.slice();
  // #region agent log
  dbgSend_('DVT', 'submitPhieuMoi', 'items dvt before save', {
    sample: itemsToSave.slice(0, 5).map(function(x) { return { ma: x.maHang || x.maVach, dvt: x.dvt }; })
  });
  // #endregion

  apiPost('luuPhieuTuWebApp', { loaiPhieu: lPhieu, khoXuat: khoXuat, khoNhan: khoNhan, items: itemsToSave }).then(function(res) {
    hideLoad();
    if(res.coLoi) { alert("⚠️ Có mã lỗi. Sửa trong tab Quản lý!"); arrItems = []; renderTable(); }
    else {
       currentPhieuObj = { soPhieu: res.soPhieu, khoXuat: khoXuat, khoNhan: khoNhan };
       document.getElementById("modal-sophieu").innerText = res.soPhieu;
       var modalTip = document.getElementById('modal-timeline-tip');
       if (modalTip) {
         var savedTip = getPackingTimelineInfo(new Date());
         modalTip.innerHTML = '<b>' + savedTip.title + '</b><br>' + savedTip.shipMsg +
           '<br><small style="opacity:0.9;">' + savedTip.windowMsg + '</small>';
       }
       document.getElementById("modal-action").style.display = "flex";
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

function actionPrintNew() {
  var items = (lastCreatedOrderItems_ && lastCreatedOrderItems_.length) ? lastCreatedOrderItems_ : arrItems;
  executePrintWeb(currentPhieuObj.soPhieu, currentPhieuObj.khoXuat, currentPhieuObj.khoNhan, items);
}
function actionExportNew() {
  var items = (lastCreatedOrderItems_ && lastCreatedOrderItems_.length) ? lastCreatedOrderItems_ : arrItems;
  executeExportExcel(currentPhieuObj.soPhieu, currentPhieuObj.khoXuat, currentPhieuObj.khoNhan, items);
}
function actionCloseModal() { document.getElementById("modal-action").style.display = "none"; arrItems = []; renderTable(); }

// ================= IN WEB (IFRAME ẨN) & XUẤT EXCEL =================
function executePrintWeb(soPhieu, khoXuat, khoNhan, itemsArray) {
  var styleStr = 'body{font-family: Arial, sans-serif; padding:20px; font-size:12px;} table{width:100%; border-collapse:collapse; margin-top:12px;} th,td{border:1px solid #000; padding:7px; text-align:left; vertical-align:top;} th{background:#f0f0f0;} .title{font-size:16px; font-weight:bold; margin-bottom:8px;} .meta{margin-bottom:8px;} .code-cell{font-size:16px; font-weight:700;} .qty-cell{font-size:16px; font-weight:700; text-align:center;} .note-cell{width:72px;} .variant-line{font-size:11px;color:#475569;margin-top:3px;} @media print { @page { margin: 10mm; } }';

  var htmlStr = '<div class="title">Số: ' + soPhieu + '</div><div class="meta"><b>Kho xuất:</b> ' + khoXuat + '<br><b>Kho nhận:</b> ' + khoNhan + '</div>';
  htmlStr += '<table><thead><tr><th>STT</th><th>Mã Parent (kệ)</th><th>Tên hàng / Phân loại</th><th>ĐVT</th><th>Số lượng (Soạn)</th><th class="note-cell"></th></tr></thead><tbody>';
  var stt = 1;
  (itemsArray || []).forEach(function(it) {
    if (Number(it.sl) > 0) {
      var parentCode = resolveDisplaySku_(it);
      var childCode = String(it.maHang || it.variantSku || '').trim();
      var tenCell = formatPackingProductName_(it);
      var variantLine = '';
      if (parentCode && childCode && parentCode.toUpperCase() !== childCode.toUpperCase()) {
        variantLine = '<div class="variant-line">Phân loại: ' + escapeHtml(it.variantName || it.tenHang || '') +
          ' · Mã con: ' + escapeHtml(childCode) +
          (it.maVach ? (' · MV: ' + escapeHtml(it.maVach)) : '') + '</div>';
      }
      htmlStr += '<tr><td>' + (stt++) + '</td><td class="code-cell">' + escapeHtml(parentCode) +
        '</td><td>' + escapeHtml(tenCell) + variantLine + '</td><td>' + escapeHtml(formatDvtDisplay_(it.maHang, it.maVach, it.dvt)) +
        '</td><td class="qty-cell">' + it.sl + '</td><td class="note-cell"></td></tr>';
    }
  });
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

/** Click số phiếu → tab mới xem PDF web + In */
function openOrderPdfView(soPhieuEncoded) {
  var soPhieu = '';
  try { soPhieu = decodeURIComponent(soPhieuEncoded || ''); } catch (e) { soPhieu = String(soPhieuEncoded || ''); }
  soPhieu = String(soPhieu || '').trim();
  if (!soPhieu) return alert('Thiếu số phiếu.');
  showLoad('Đang tải chi tiết phiếu ' + soPhieu + '...');
  apiGet('getOrderDetail', { soPhieu: soPhieu }, { timeoutMs: 60000 }).then(function(res) {
    hideLoad();
    if (!res || !res.success) {
      // Fallback cũ nếu chưa deploy getOrderDetail
      return apiGet('getChiTietPhieu', { soPhieu: soPhieu, storeName: (sessionUser && sessionUser.store) || '', includeStock: '0' }, { timeoutMs: 45000 }).then(function(legacy) {
        var parsed = unwrapListResponse_(legacy);
        var rows = parsed.rows || [];
        var items = rows.filter(function(r) {
          return r && r.trangThai !== 'Đã hủy dòng' && r.trangThai !== 'Đã hủy đơn';
        }).map(function(r) {
          var sl = Number(r.slSoan) || Number(r.slThucTe) || Number(r.slGoc) || 0;
          return { maHang: r.maHang, maVach: r.maVach, tenHang: r.tenHang, dvt: r.dvt, sl: sl, trangThai: r.trangThai, parentSku: r.parentSku || '', variantSku: r.variantSku || r.maHang || '', variantName: r.variantName || r.tenHang || '', parentName: r.parentName || '', maHangDisplay: r.maHangDisplay || '', tenHangDisplay: r.tenHangDisplay || '' };
        });
        var info = (phieuData || []).find(function(p) { return p && p.soPhieu === soPhieu; }) || {};
        openOrderPdfWindow_({
          success: true,
          soPhieu: soPhieu,
          khoXuat: info.khoXuat || '',
          khoNhan: info.khoNhan || '',
          thoiGian: info.thoiGian || '',
          items: items
        });
      });
    }
    openOrderPdfWindow_(res);
  }).catch(function(err) {
    hideLoad();
    alert('Lỗi tải PDF phiếu: ' + (err && err.message || err));
  });
}

function openOrderPdfWindow_(detail) {
  if (!detail) return;
  var w = window.open('', '_blank');
  if (!w) {
    alert('Trình duyệt chặn tab mới. Cho phép popup rồi thử lại.');
    return;
  }
  w.document.open();
  w.document.write(buildOrderPdfHtml_(detail));
  w.document.close();
}

function buildOrderPdfHtml_(detail) {
  var so = escapeHtml(detail.soPhieu || '');
  var kx = escapeHtml(formatStoreShortLabel_(detail.khoXuat) || detail.khoXuat || '-');
  var kn = escapeHtml(formatStoreShortLabel_(detail.khoNhan) || detail.khoNhan || '-');
  var tg = escapeHtml(detail.thoiGian || '-');
  var rowsHtml = '';
  var stt = 1;
  (detail.items || []).forEach(function(it) {
    var sl = Number(it.sl);
    if (!sl || sl <= 0) return;
    var parentCode = resolveDisplaySku_(it);
    var childCode = String(it.maHang || it.variantSku || '').trim();
    var tenCell = formatPackingProductName_(it);
    var variantLine = '';
    if (parentCode && childCode && parentCode.toUpperCase() !== childCode.toUpperCase()) {
      variantLine = '<div class="variant-line">Phân loại: ' + escapeHtml(it.variantName || it.tenHang || '') +
        ' · Mã con: ' + escapeHtml(childCode) +
        (it.maVach ? (' · MV: ' + escapeHtml(it.maVach)) : '') + '</div>';
    } else if (it.maVach) {
      variantLine = '<div class="variant-line">MV: ' + escapeHtml(it.maVach) + '</div>';
    }
    rowsHtml += '<tr><td>' + (stt++) + '</td><td class="code">' + escapeHtml(parentCode) + '</td><td>' + escapeHtml(tenCell) + variantLine +
      '</td><td>' + escapeHtml(formatDvtDisplay_(it.maHang, it.maVach, it.dvt)) +
      '</td><td class="qty">' + sl + '</td><td class="note"></td></tr>';
  });
  if (!rowsHtml) rowsHtml = '<tr><td colspan="6" style="text-align:center;color:#64748b;">Không có dòng hàng hợp lệ.</td></tr>';
  return '<!DOCTYPE html><html><head><meta charset="utf-8"><title>Phiếu ' + so + '</title>' +
    '<style>' +
    'body{font-family:"Segoe UI",Arial,sans-serif;margin:0;padding:24px;color:#0f172a;background:#f8fafc;}' +
    '.sheet{max-width:920px;margin:0 auto;background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:28px 32px;box-shadow:0 8px 24px rgba(15,23,42,.06);}' +
    '.toolbar{display:flex;gap:10px;justify-content:flex-end;margin-bottom:16px;}' +
    '.toolbar button{border:none;background:#1a73e8;color:#fff;font-weight:700;padding:10px 16px;border-radius:8px;cursor:pointer;}' +
    'h1{margin:0 0 6px;font-size:22px;} .meta{line-height:1.6;margin:12px 0 18px;color:#334155;}' +
    'table{width:100%;border-collapse:collapse;} th,td{border:1px solid #cbd5e1;padding:8px 10px;font-size:13px;vertical-align:top;}' +
    'th{background:#f1f5f9;text-align:left;} .code{font-weight:700;font-size:15px;} .qty{font-weight:800;text-align:center;font-size:15px;}' +
    '.note{width:72px;} .variant-line{margin-top:4px;font-size:11px;color:#475569;} .signs{display:flex;justify-content:space-between;margin-top:48px;text-align:center;}' +
    '@media print{.toolbar{display:none!important;} body{background:#fff;padding:0;} .sheet{box-shadow:none;border:none;border-radius:0;max-width:none;padding:0;} @page{margin:10mm;}}' +
    '</style></head><body><div class="toolbar"><button type="button" onclick="window.print()">🖨️ In PDF</button></div>' +
    '<div class="sheet"><h1>PHIẾU XUẤT / SOẠN HÀNG</h1>' +
    '<div class="meta"><b>Mã phiếu:</b> ' + so + '<br><b>Kho xuất:</b> ' + kx + '<br><b>Kho nhận:</b> ' + kn +
    '<br><b>Thời gian:</b> ' + tg + '</div>' +
    '<table><thead><tr><th>STT</th><th>Mã Parent (kệ)</th><th>Tên hàng / Phân loại</th><th>ĐVT</th><th>SL</th><th class="note"></th></tr></thead><tbody>' +
    rowsHtml + '</tbody></table>' +
    '<div class="signs"><div><b>Người lập phiếu</b><br><br><br>Ký ghi rõ họ tên</div>' +
    '<div><b>Người nhận</b><br><br><br>Ký ghi rõ họ tên</div></div></div></body></html>';
}

// ================= QUẢN LÝ PHIẾU =================
function getNgayFilterParam(inputId) {
  var inputEl = document.getElementById(inputId);
  return inputEl && inputEl.value ? inputEl.value : '';
}

function sh_getSuggestedPackingDayStr_() {
  var info = getPackingTimelineInfo(new Date());
  return formatDateInputValue(info.packingDay);
}

function sh_ensureCreateDateDefaults_() {
  var fromEl = document.getElementById('sh-create-from');
  var toEl = document.getElementById('sh-create-to');
  if (!toEl) return;
  // Mặc định = ngày tổng hợp theo khung giờ (sau 10:00 → ngày mai), KHÔNG phải ngày lịch
  var packingDefault = sh_getSuggestedPackingDayStr_();
  if (!toEl.value || !shCreateDateUserTouched_) {
    toEl.value = packingDefault;
  }
  // from = ngày trước packing day (tự động, dùng cho API cũ)
  if (fromEl) {
    var packing = new Date((toEl.value || packingDefault) + 'T00:00:00');
    var prev = new Date(packing);
    prev.setDate(packing.getDate() - 1);
    fromEl.value = formatDateInputValue(prev);
  }
}

function sh_getCreateDateRange_() {
  sh_ensureCreateDateDefaults_();
  var fromEl = document.getElementById('sh-create-from');
  var toEl = document.getElementById('sh-create-to');
  var packingDay = toEl && toEl.value ? toEl.value : sh_getSuggestedPackingDayStr_();
  if (toEl) toEl.value = packingDay;
  var packing = new Date(packingDay + 'T00:00:00');
  var prev = new Date(packing);
  prev.setDate(packing.getDate() - 1);
  var from = formatDateInputValue(prev);
  if (fromEl) fromEl.value = from;
  return {
    from: from,
    to: packingDay,
    packingDay: packingDay,
    suggestedPackingDay: sh_getSuggestedPackingDayStr_()
  };
}

function sh_onCreateDateRangeChange() {
  shCreateDateUserTouched_ = true;
  sh_getCreateDateRange_();
  var pickerEl = document.getElementById('sh-order-picker');
  if (pickerEl && pickerEl.style.display !== 'none') sh_taiDanhSachDonSoanChoBang();
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
  var ngay = getNgayFilterParam('ql-ngay');
  var khoNhan = document.getElementById("ql-kho-nhan") ? document.getElementById("ql-kho-nhan").value : '';
  var tQl0 = Date.now();
  // #region agent log
  dbgBranch_('B', 'ql_loadPhieu:start', 'branch manage list request', {
    role: sessionUser.role || '',
    store: sessionUser.store || '',
    khoNhan: khoNhan,
    ngay: ngay || ''
  });
  // #endregion
  apiGet('layDanhSachPhieuTheoFilter', { khoNhan: khoNhan, ngay: ngay, userRole: sessionUser.role, userStore: sessionUser.store }, { directOnly: true, timeoutMs: 45000 }).then(function(res) {
    var parsed = unwrapListResponse_(res);
    var rows = parsed.rows;
    // #region agent log
    dbgBranch_('B', 'ql_loadPhieu:ok', 'branch manage list response', {
      ms: Date.now() - tQl0,
      rowCount: rows.length,
      debugMs: parsed.meta && parsed.meta._debugTotalMs,
      debugScanned: parsed.meta && parsed.meta._debugScanned,
      debugStoreCalls: parsed.meta && parsed.meta._debugStoreCalls,
      debugRun: parsed.meta && parsed.meta._debugRun,
      debugUserStore: parsed.meta && parsed.meta._debugUserStore,
      sample: rows[0] ? { soPhieu: rows[0].soPhieu, kx: rows[0].khoXuat, kn: rows[0].khoNhan, st: rows[0].trangThai } : null
    });
    // #endregion
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
    if (selectedSoPhieu) {
      var matched = rows.find(function(item) { return item.soPhieu === selectedSoPhieu; });
      if (matched) {
        selectEl.value = selectedSoPhieu;
        ql_hienThiChiTiet(matched);
      }
    }
  }).catch(function(err){
    // #region agent log
    dbgBranch_('B', 'ql_loadPhieu:err', 'branch manage list failed', {
      ms: Date.now() - tQl0,
      err: String(err && err.message || err)
    });
    // #endregion
    selectEl.innerHTML = '<option value="">-- Lỗi tải --</option>';
    alert('Lỗi: '+err.message);
  });
}

function ql_onSelectPhieu() {
  var val = document.getElementById("ql-phieu").value;
  if(!val) {
    document.getElementById("ql-view-phieu").style.display = "none";
    var ban = document.getElementById('ql-edit-timeline-banner');
    if (ban) ban.style.display = 'none';
    return;
  }
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
  var ngay = getNgayFilterParam('confirm-ngay');
  var t0 = Date.now();
  // #region agent log
  dbgConfirm_('C', 'confirm_loadPhieu:start', 'confirm list request', {
    role: sessionUser.role || '',
    store: sessionUser.store || '',
    khoNhan: confirmStoreFilter,
    ngay: ngay || '',
    build: APP_BUILD
  });
  // #endregion
  apiGet('layDanhSachPhieuTheoFilter', { khoNhan: confirmStoreFilter, ngay: ngay, userRole: sessionUser.role, userStore: sessionUser.store }, { directOnly: true, timeoutMs: 45000 }).then(function(res) {
    try {
      // #region agent log
      dbgConfirm_('A', 'confirm_loadPhieu:response', 'raw response shape', {
        ms: Date.now() - t0,
        isArray: Array.isArray(res),
        type: typeof res,
        keys: res && typeof res === 'object' ? Object.keys(res).slice(0, 20) : [],
        error: res && res.error ? String(res.error).slice(0, 300) : '',
        success: res && res.success,
        debugRun: res && res._debugRun,
        debugMs: res && res._debugTotalMs,
        debugScanned: res && res._debugScanned,
        debugStoreCalls: res && res._debugStoreCalls,
        debugRole: res && res._debugRole,
        debugKhoNhan: res && res._debugKhoNhan,
        debugUserStore: res && res._debugUserStore,
        rowLen: Array.isArray(res) ? res.length : (res && Array.isArray(res.data) ? res.data.length : -1)
      });
      // #endregion
      var parsed = unwrapListResponse_(res);
      var rows = Array.isArray(parsed.rows) ? parsed.rows : [];
      if (res && res.error) {
        // #region agent log
        dbgConfirm_('B', 'confirm_loadPhieu:gas-error', 'GAS returned error field', { error: String(res.error).slice(0, 400), ms: Date.now() - t0 });
        // #endregion
      }
      // Chỉ đơn Đã soạn — loại Đã xác nhận / Đã hủy (không cho mở lại để lưu)
      var statusHist = {};
      rows.forEach(function(r) {
        var st = r && String(r.trangThai || '').trim() || '(empty)';
        statusHist[st] = (statusHist[st] || 0) + 1;
      });
      var confirmableOrders = rows.filter(function(r) {
        return r && String(r.trangThai || '').trim() === 'Đã soạn';
      });
      // #region agent log
      dbgConfirm_('H2', 'confirm_loadPhieu:parsed', 'parsed confirmable + status hist', {
        totalRows: rows.length,
        confirmable: confirmableOrders.length,
        statusHist: statusHist,
        sample: rows.slice(0, 12).map(function(r) {
          return r ? { soPhieu: r.soPhieu, trangThai: r.trangThai, meta: dbgStatusBag_(r.trangThai) } : null;
        }),
        confirmableSample: confirmableOrders.slice(0, 8).map(function(r) { return r && r.soPhieu; }),
        excludedConfirmed: rows.filter(function(r) {
          var st = String((r && r.trangThai) || '');
          return st.indexOf('xác nhận') !== -1 || st.indexOf('Xác nhận') !== -1;
        }).length
      });
      // #endregion
      var html = '<option value="">-- Chọn Phiếu (' + confirmableOrders.length + ') --</option>';
      confirmableOrders.forEach(function(r) {
        var shortName = storeMap[r.khoNhan] || storeMap[r.khoXuat] || r.khoNhan || r.khoXuat || '';
        html += '<option value="'+r.soPhieu+'">'+r.soPhieu+' ('+shortName+') ['+r.trangThai+']</option>';
      });
      selectEl.innerHTML = html;
      if (selectedSoPhieu) {
        var found = confirmableOrders.some(function(item) { return item.soPhieu === selectedSoPhieu; });
        if (found) {
          selectEl.value = selectedSoPhieu;
          confirm_onSelectPhieu();
        } else {
          // Đơn đã xác nhận / không còn trong danh sách chờ
          document.getElementById('confirm-view').style.display = 'none';
          currentConfirmPhieuObj = null;
        }
      }
    } catch (parseErr) {
      // #region agent log
      dbgConfirm_('D', 'confirm_loadPhieu:js-error', 'FE runtime after response', {
        name: parseErr && parseErr.name,
        message: parseErr && parseErr.message ? String(parseErr.message).slice(0, 400) : '',
        ms: Date.now() - t0
      });
      // #endregion
      selectEl.innerHTML = '<option value="">-- Lỗi tải --</option>';
      alert('Lỗi: ' + (parseErr && parseErr.message ? parseErr.message : parseErr));
    }
  }).catch(function(err){
    // #region agent log
    dbgConfirm_('B', 'confirm_loadPhieu:catch', 'API/network/timeout', {
      name: err && err.name,
      message: err && err.message ? String(err.message).slice(0, 500) : '',
      ms: Date.now() - t0,
      role: sessionUser.role || '',
      store: sessionUser.store || ''
    });
    // #endregion
    selectEl.innerHTML = '<option value="">-- Lỗi tải --</option>';
    alert('Lỗi: '+err.message);
  });
}

/** SL hiển thị/điền sẵn ở tab Xác nhận = SL soạn/giao (không dùng SL đặt) */
function confirm_resolvePackedQty_(r) {
  if (!r) return 0;
  var fromSoan = ql_getSlSoan_(r);
  if (fromSoan !== "") return Number(fromSoan) || 0;
  // Fallback cũ: trước khi có cột 16, SL soạn nằm ở slThucTe
  if (r.slThucTe !== undefined && r.slThucTe !== null && r.slThucTe !== "") return Number(r.slThucTe) || 0;
  return Number(r.slGoc) || 0;
}

function confirm_onSelectPhieu() {
  var val = document.getElementById("confirm-phieu").value;
  if (!val) { document.getElementById("confirm-view").style.display = "none"; return; }
  showLoad("Đang tải dữ liệu nhận hàng...");
  apiGet('getChiTietPhieu', { soPhieu: val, storeName: sessionUser.store, includeStock: '0' }, { directOnly: true, timeoutMs: 45000 }).then(function(res) {
    hideLoad();
    var parsed = unwrapListResponse_(res);
    var rows = parsed.rows || [];
    var activeRows = rows.filter(function(r) {
      return r && r.trangThai !== "Đã hủy dòng" && r.trangThai !== "Đã hủy đơn";
    });
    var confirmedCount = activeRows.filter(function(r) {
      return String(r.trangThai || '').trim() === 'Đã xác nhận nhận hàng';
    }).length;
    var pendingCount = activeRows.length - confirmedCount;
    var fullyConfirmed = activeRows.length > 0 && pendingCount === 0;

    // #region agent log
    var lineStatusHist = {};
    activeRows.forEach(function(r) {
      var st = String((r && r.trangThai) || '(empty)');
      lineStatusHist[st] = (lineStatusHist[st] || 0) + 1;
    });
    dbgConfirm_('H3', 'confirm_onSelectPhieu', 'confirm open lock check', {
      soPhieu: val,
      active: activeRows.length,
      confirmedCount: confirmedCount,
      pendingCount: pendingCount,
      fullyConfirmed: fullyConfirmed,
      lineStatusHist: lineStatusHist,
      sampleLines: activeRows.slice(0, 6).map(function(r) {
        return { rowIndex: r.rowIndex, trangThai: r.trangThai, meta: dbgStatusBag_(r.trangThai), slSoan: r.slSoan, slThucTe: r.slThucTe };
      })
    });
    // #endregion

    if (fullyConfirmed) {
      currentConfirmPhieuObj = null;
      document.getElementById("confirm-view").style.display = "none";
      var sel = document.getElementById("confirm-phieu");
      if (sel) sel.value = "";
      alert("Đơn " + val + " đã xác nhận nhận hàng — không thể xác nhận lại.\n(Giống khóa chỉnh sửa ở tab Quản lý.)");
      confirm_loadPhieu();
      return;
    }

    currentConfirmPhieuObj = { soPhieu: val };
    var viewEl = document.getElementById("confirm-view");
    document.getElementById("confirm-lbl-sophieu").innerText = val;
    var serverMs = parsed.meta && parsed.meta._debugTotalMs;
    var metaText = "Đã tải " + activeRows.length + " dòng";
    if (confirmedCount > 0) metaText += " · " + confirmedCount + " dòng đã xác nhận (chỉ lưu các dòng còn lại)";
    if (serverMs) metaText += " (" + Math.round(serverMs / 1000) + "s)";
    document.getElementById("confirm-meta").innerText = metaText;
    var tbody = document.getElementById("confirm-tbody");
    tbody.innerHTML = "";
    activeRows.forEach(function(r, idx) {
      var packedQty = confirm_resolvePackedQty_(r);
      var isConfirmedLine = String(r.trangThai || '').trim() === 'Đã xác nhận nhận hàng';
      var inputValue = isConfirmedLine && r.slThucTe !== "" && r.slThucTe != null
        ? Number(r.slThucTe)
        : packedQty;
      var disabledAttr = isConfirmedLine ? ' disabled' : '';
      var rowNote = isConfirmedLine ? ' <small style="color:#64748b;">(đã xác nhận)</small>' : '';
      tbody.insertAdjacentHTML('beforeend',
        '<tr><td>' + (idx + 1) + '</td><td><b>' + (r.maVach || '') + '</b><br><small style="color:gray;">' + (r.maHang || '') + '</small><br><small>' + (r.tenHang || '') + '</small>' + rowNote + '</td><td>' + escapeHtml(formatDvtDisplay_(r.maHang, r.maVach, r.dvt)) + '</td><td style="font-weight:700;color:#c2410c;">' + packedQty + '</td><td><input type="number" class="confirm-qty-input' + (isConfirmedLine ? '' : ' same') + '" data-row="' + r.rowIndex + '" data-packed="' + packedQty + '" data-previous="' + inputValue + '" data-confirmed="' + (isConfirmedLine ? '1' : '0') + '" value="' + inputValue + '" min="0" oninput="confirm_updateInput(this)"' + disabledAttr + '></td></tr>');
    });
    var saveBtn = document.querySelector('#confirm-view .btn-success');
    if (saveBtn) {
      saveBtn.disabled = pendingCount === 0;
      saveBtn.style.opacity = pendingCount === 0 ? '0.5' : '1';
    }
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
    if (inputEl.getAttribute("data-confirmed") === "1" || inputEl.disabled) return;
    var row = parseInt(inputEl.getAttribute("data-row"));
    var qty = Number(inputEl.value);
    var previousQty = Number(inputEl.getAttribute("data-previous") || 0);
    if (!isNaN(row) && !isNaN(qty) && qty >= 0) confirmations.push({ row: row, receivedQty: qty, previousQty: previousQty });
  });
  if (!confirmations.length) return alert("Không còn dòng nào cần xác nhận (đơn đã khóa).");
  var tSave0 = Date.now();
  showLoad("Đang lưu xác nhận...");
  // #region agent log
  dbgConfirm_('H1', 'confirm_xacNhanNhanHang:start', 'confirm save start', {
    soPhieu: currentConfirmPhieuObj.soPhieu,
    lines: confirmations.length,
    rowIndexes: confirmations.map(function(c) { return c.row; }).slice(0, 40)
  });
  // #endregion
  // POST thẳng GAS, không qua proxy (tránh 502 HTML khi request nặng)
  apiPost('xacNhanNhanHang', {
    soPhieu: currentConfirmPhieuObj.soPhieu,
    actor: sessionUser.user,
    store: sessionUser.store,
    confirmations: confirmations
  }, { directOnly: true, timeoutMs: 180000 }).then(function(res) {
    hideLoad();
    // #region agent log
    dbgConfirm_('H1', 'confirm_xacNhanNhanHang:ok', 'confirm save ok', {
      clientMs: Date.now() - tSave0,
      serverMs: res && res._debugTotalMs,
      debugRun: res && res._debugRun,
      count: res && res.count,
      success: res && res.success,
      error: res && (res.error || res.msg) ? String(res.error || res.msg).slice(0, 300) : '',
      hasNotify: !!(res && res.notify)
    });
    // #endregion
    if (!res || res.success === false) throw new Error((res && (res.error || res.msg)) || "Không thể lưu xác nhận.");
    alert("✅ Đã lưu xác nhận nhận hàng." + (res._debugTotalMs ? ("\n(Lưu: " + Math.round(res._debugTotalMs / 1000) + "s)") : ""));
    resetConfirmViewAfterSave();
    if (res.notify) {
      apiPost('postProcessReceiveOrder', res.notify, { allowDirectFallback: false, timeoutMs: 120000 }).catch(function() {});
    }
  }).catch(function(err){
    hideLoad();
    var errMsg = String(err && err.message || err);
    // #region agent log
    dbgConfirm_('H1', 'confirm_xacNhanNhanHang:err', 'confirm save failed', {
      clientMs: Date.now() - tSave0,
      err: errMsg.slice(0, 500),
      looksAlreadyConfirmed: errMsg.indexOf('đã xác nhận') !== -1 || errMsg.indexOf('Đã xác nhận') !== -1,
      looksTimeout: /timeout|502|Failed to fetch|HTML|NetworkError/i.test(errMsg),
      soPhieu: currentConfirmPhieuObj && currentConfirmPhieuObj.soPhieu
    });
    // #endregion
    // Sau lỗi: reload list để xem đơn còn hiện "Đã soạn" không (khóa một phần)
    try { confirm_loadPhieu(currentConfirmPhieuObj && currentConfirmPhieuObj.soPhieu); } catch (eReload) {}
    alert('Lỗi: ' + err.message + '\nThử lại sau khi deploy code.gs (confirm-fast).');
  });
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
  apiGet('getChiTietPhieu', { soPhieu: currentPhieuObj.soPhieu, storeName: currentPhieuObj.khoXuat, includeStock: '1' }, { directOnly: true, timeoutMs: 45000 }).then(function(res) {
    hideLoad();
    var parsed = unwrapListResponse_(res);
    var rows = parsed.rows;
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

    var lowStockCount = 0;
    var packedLineCount = 0;
    rows.forEach((r, i) => {
      var isCancelled = r.trangThai === "Đã hủy dòng" || r.trangThai === "Đã hủy đơn";
      var isReadOnlyRow = isConfirmedOrder || isCancelled || !canEditRows;
      var slSoanVal = ql_getSlSoan_(r);
      var compareNeed = slSoanVal !== "" ? slSoanVal : r.slGoc;
      var shortInfo = !isCancelled ? stockShortageInfo_(r.stock, compareNeed) : null;
      if (shortInfo) lowStockCount++;
      if (slSoanVal !== "") packedLineCount++;
      var rowStyle = isCancelled
        ? 'background:#fce8e6; color:#777; text-decoration:line-through;'
        : (shortInfo ? 'background:#fef2f2;' : (r.ghiChu ? 'background:#fff8e1;' : ''));
      var rowKey = String(r.tempKey || r.rowIndex);
      var dvtDisplay = formatDvtDisplay_(r.maHang, r.maVach, r.dvt);
      r.dvt = resolveDvtClient_(r.maHang, r.maVach, r.dvt) || r.dvt;
      var variantHtml = ql_renderVariantSelector(r, rowKey, isReadOnlyRow);
      var quantityInput = '<input type="number" class="edit-sl-input" data-row="'+r.rowIndex+'" data-stock="'+(r.stock === '' || r.stock == null ? '' : r.stock)+'" data-new="0" value="'+(r.slGoc || 0)+'" '+(isReadOnlyRow ? 'disabled' : '')+' style="border:2px solid #1a73e8;text-align:center;width:70px;" oninput="ql_onEditQtyInput(this)">';
      var cancelButton = canAddItems ? '<td><button type="button" onclick="ql_huyDong('+r.rowIndex+')" '+(isCancelled ? 'disabled' : '')+' style="border:none; background:#d93025; color:white; border-radius:5px; padding:7px 9px; cursor:pointer;">Hủy mã</button></td>' : '<td></td>';
      var requestedDisplay = (!isReadOnlyRow && !isPublicView)
        ? quantityInput
        : ('<b>' + (Number(r.slGoc) || 0) + '</b>');
      var packedDisplay = slSoanVal !== ""
        ? ('<b style="color:#c2410c;">' + Number(slSoanVal) + '</b>')
        : '<span style="color:#94a3b8;">-</span>';
      var hasReceivedQty = r.trangThai === "Đã xác nhận nhận hàng" && r.slThucTe !== undefined && r.slThucTe !== null && r.slThucTe !== "";
      var receivedDisplay = hasReceivedQty
        ? ('<b style="color:#166534;">' + Number(r.slThucTe) + '</b>')
        : '<span style="color:#94a3b8;">-</span>';
      var stockHtml = formatStockCellHtml_(r.stock, compareNeed);
      var tenShow = formatPackingProductName_(r);
      var codeHtml = formatPackingCodeHtml_(r);
      // Giữ selector biến thể khi sửa; mã Parent vẫn nổi bật phía trên
      var maCell = variantHtml
        ? (codeHtml + '<div style="margin-top:6px;">' + variantHtml + '</div>')
        : codeHtml;
      tb.insertAdjacentHTML('beforeend', '<tr data-stock-row="1" style="'+rowStyle+'"><td>'+(i+1)+'</td><td>'+maCell+'</td><td>'+escapeHtml(tenShow)+(r.ghiChu?'<br><b style="color:red;font-size:11px;">⚠️ '+escapeHtml(r.ghiChu)+'</b>':'')+((isCancelled || r.trangThai === "Đã soạn hàng" || r.trangThai === "Đã xác nhận nhận hàng") ? '<br><b style="color:#d93025;font-size:11px;">'+escapeHtml(r.trangThai)+'</b>' : '')+'</td><td id="ql-dvt-'+rowKey+'">'+dvtDisplay+'</td><td class="ql-stock-cell">'+stockHtml+'</td><td>'+requestedDisplay+'</td><td class="ql-sl-soan">'+packedDisplay+'</td><td>'+receivedDisplay+'</td>'+cancelButton+'</tr>');
    });
    var packerNames = rows.map(function(r){ return r.nguoiSoanHang || ""; }).filter(Boolean);
    var metaBits = [];
    if (!isPublicView && packerNames.length) metaBits.push("Người soạn hàng gần nhất: " + packerNames[packerNames.length - 1]);
    if (isConfirmedOrder && !isPublicView) metaBits.push("Chế độ chỉ xem sau khi xác nhận nhận hàng");
    else if (isPackedOrder && !isAdmin && !isPublicView) metaBits.push("Chi nhánh không thể sửa sau khi đã soạn xong");
    if (packedLineCount > 0) metaBits.push("In/Excel dùng SL Soạn (" + packedLineCount + " mã)");
    if (lowStockCount > 0) metaBits.push("⚠️ " + lowStockCount + " mã tồn thấp hơn SL đặt/soạn (chỉ cảnh báo)");
    document.getElementById("ql-order-meta").innerHTML = metaBits.length
      ? metaBits.map(function(t) {
          return t.indexOf('⚠️') === 0
            ? '<span style="color:#b91c1c;font-weight:700;">' + escapeHtml(t) + '</span>'
            : '<span>' + escapeHtml(t) + '</span>';
        }).join(' | ')
      : '';
    document.getElementById("ql-view-phieu").style.display = "block";
    refreshPackingTimelineBanners();
  }).catch(function(err){ hideLoad(); alert('Lỗi: '+err.message); });
}

/** SL đã soạn: ưu tiên slSoan (cột SL Giao), fallback khi đơn đã soạn còn dùng slThucTe cũ */
function ql_getSlSoan_(row) {
  if (!row) return "";
  if (row.slSoan !== undefined && row.slSoan !== null && row.slSoan !== "") return Number(row.slSoan);
  var packed = row.trangThai === "Đã soạn hàng" || row.trangThai === "Đã xác nhận nhận hàng";
  if (packed && row.slThucTe !== undefined && row.slThucTe !== null && row.slThucTe !== "") return Number(row.slThucTe);
  return "";
}

/** Số lượng dùng khi In / Excel: ưu tiên SL Soạn sau khi đã soạn */
function ql_resolvePrintQty_(row) {
  if (!row) return 0;
  if (row.trangThai === "Đã hủy dòng" || row.trangThai === "Đã hủy đơn") return 0;
  var slSoan = ql_getSlSoan_(row);
  if (slSoan !== "") return Number(slSoan) || 0;
  var ip = document.querySelector('.edit-sl-input[data-row="' + row.rowIndex + '"]');
  if (ip && ip.value !== "" && ip.getAttribute("data-new") !== "1") return Number(ip.value) || 0;
  return Number(row.slGoc) || 0;
}

function ql_onEditQtyInput(input) {
  if (!input) return;
  var tr = input.closest('tr');
  if (!tr) return;
  var stockRaw = input.getAttribute('data-stock');
  var stockCell = tr.querySelector('.ql-stock-cell');
  var need = Number(input.value) || 0;
  if (stockCell) stockCell.innerHTML = formatStockCellHtml_(stockRaw === '' ? '' : stockRaw, need);
  var short = stockShortageInfo_(stockRaw === '' ? '' : stockRaw, need);
  if (tr.getAttribute('data-stock-row') === '1') {
    tr.style.background = short ? '#fef2f2' : '';
  }
}

function ql_getVariantOptions(rowMeta) {
  var maHang = String(rowMeta && rowMeta.maHang ? rowMeta.maHang : '').trim().toUpperCase();
  var parentSku = getParentSkuOfItem_(rowMeta);
  if (!parentSku && maHang && danhMucGoc[maHang]) parentSku = getParentSkuOfItem_(danhMucGoc[maHang]);
  var seen = {};
  var variants = [];
  danhMucArr.forEach(function(item) {
    if (!item) return;
    var itemMaHang = String(item.maHang || '').trim().toUpperCase();
    var itemParent = getParentSkuOfItem_(item);
    var sameGroup = parentSku
      ? (itemParent === parentSku)
      : (maHang && itemMaHang === maHang);
    if (!sameGroup) return;
    var dvtItem = resolveDvtClient_(item.maHang, item.maVach, item.dvt);
    var key = String(item.maHang || '').trim().toUpperCase() + '|' + String(item.maVach || '').trim().toUpperCase() + '|' + String(dvtItem || '').trim().toUpperCase();
    if (!key || seen[key]) return;
    seen[key] = true;
    variants.push({
      maVach: item.maVach || '',
      dvt: dvtItem,
      tenHang: item.tenHang || rowMeta.tenHang || '',
      maHang: item.maHang || rowMeta.maHang || '',
      parentSku: itemParent || parentSku || ''
    });
  });

  var currentDvt = resolveDvtClient_(rowMeta.maHang, rowMeta.maVach, rowMeta.dvt);
  var currentKey = String(rowMeta.maHang || '').trim().toUpperCase() + '|' + String(rowMeta.maVach || '').trim().toUpperCase() + '|' + String(currentDvt || '').trim().toUpperCase();
  if (!seen[currentKey]) {
    variants.unshift({
      maVach: rowMeta.maVach || '',
      dvt: currentDvt,
      tenHang: rowMeta.tenHang || '',
      maHang: rowMeta.maHang || '',
      parentSku: parentSku || ''
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
    var selected = String(v.maVach || '').trim().toUpperCase() === currentMaVach &&
      String(v.maHang || '').trim().toUpperCase() === String(rowMeta.maHang || '').trim().toUpperCase()
      ? ' selected' : '';
    var payload = encodeURIComponent(JSON.stringify(v));
    var label = (v.tenHang || v.maVach || '-') + ' | MH: ' + (v.maHang || '-') + ' | MV: ' + (v.maVach || '-') + ' | ĐVT: ' + (v.dvt || 'Cái');
    optionsHtml += '<option value="' + payload + '"' + selected + '>' + escapeHtml(label.replace(v.dvt || 'Cái', formatDvtDisplay_(v.maHang, v.maVach, v.dvt))) + '</option>';
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
  rowMeta.tenHang = variant.tenHang || rowMeta.tenHang;
  if (variant.maHang) rowMeta.maHang = variant.maHang;
  rowMeta.dvt = resolveDvtClient_(rowMeta.maHang, rowMeta.maVach, variant.dvt || rowMeta.dvt);
  var dvtCell = document.getElementById('ql-dvt-' + rowKey);
  if (dvtCell) dvtCell.innerText = formatDvtDisplay_(rowMeta.maHang, rowMeta.maVach, rowMeta.dvt);
}

function ql_isDuplicateOrderItem(item) {
  if (!item) return false;
  var itemMaHang = String(item.maHang || "").trim().toUpperCase();
  var itemMaVach = String(item.maVach || "").trim().toUpperCase();
  return editRows.some(function(existing) {
    if (!existing) return false;
    var existingMaHang = String(existing.maHang || "").trim().toUpperCase();
    var existingMaVach = String(existing.maVach || "").trim().toUpperCase();
    // Trùng khi cùng MH + MV — cho phép nhiều biến thể khác nhau cùng Parent trên 1 đơn
    if (itemMaHang && existingMaHang && itemMaVach && existingMaVach) {
      return itemMaHang === existingMaHang && itemMaVach === existingMaVach;
    }
    if (itemMaHang && existingMaHang && itemMaHang === existingMaHang && !itemMaVach && !existingMaVach) return true;
    if (itemMaVach && existingMaVach && itemMaVach === existingMaVach && !itemMaHang && !existingMaHang) return true;
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
          newItems.push({
            maHang: rowMeta.maHang,
            maVach: rowMeta.maVach,
            tenHang: rowMeta.tenHang,
            dvt: resolveDvtClient_(rowMeta.maHang, rowMeta.maVach, rowMeta.dvt),
            sl: qtyValue
          });
        }
      } else {
        var variant = variantByRow[String(rowKey)] || null;
        var rowMetaUpd = editRows.find(function(item) { return String(item.rowIndex) === String(rowKey); });
        var dvtUpd = variant
          ? resolveDvtClient_(variant.maHang || (rowMetaUpd && rowMetaUpd.maHang), variant.maVach || (rowMetaUpd && rowMetaUpd.maVach), variant.dvt)
          : resolveDvtClient_(rowMetaUpd && rowMetaUpd.maHang, rowMetaUpd && rowMetaUpd.maVach, rowMetaUpd && rowMetaUpd.dvt);
        updates.push({
          row: parseInt(rowKey),
          valSl: qtyValue,
          valMaHang: variant ? (variant.maHang || '') : '',
          valMaVach: variant ? (variant.maVach || '') : '',
          valDvt: dvtUpd || '',
          valTenHang: variant ? (variant.tenHang || '') : ''
        });
      }
    }
  });
  if (!updates.length && !newItems.length) return alert("Không có thay đổi nào cần lưu.");
  refreshPackingTimelineBanners();
  var tip = getPackingTimelineInfo(new Date());
  var editConfirm = 'Xác nhận lưu sửa đơn' +
    (newItems.length ? (' (có ' + newItems.length + ' mã thêm mới)') : '') +
    '?\n\n' + tip.title + '\n' + tip.shipMsg + '\n\n' + tip.windowMsg;
  if (!window.confirm(editConfirm)) return;
  showLoad("Đang lưu chỉnh sửa...");
  apiPost('luuChinhSuaPhieu', { soPhieu: currentPhieuObj ? currentPhieuObj.soPhieu : "", updates: updates, newItems: newItems, actor: sessionUser.user }).then(function(res) {
    hideLoad();
    if (!res.success) throw new Error(res.error || "Không thể lưu thay đổi.");
    var afterTip = getPackingTimelineInfo(new Date());
    alert("✅ Đã lưu chỉnh sửa thành công!\n\n" + afterTip.title + "\n" + afterTip.shipMsg);
    resetManagementViewAfterSave();
  }).catch(function(err){ hideLoad(); alert('Lỗi: '+err.message); });
}

function ql_themMaHangDirect_(item, quantity) {
  if (!sessionUser.user) return alert("Vui lòng đăng nhập trước khi thao tác.");
  if (!item) return alert("Không tìm thấy mã hàng hóa hoặc mã vạch.");
  quantity = Number(quantity);
  if (!quantity || quantity < 1) return alert("Số lượng phải lớn hơn 0.");
  if (ql_isDuplicateOrderItem(item)) {
    alert("⚠️ Mã này đã tồn tại trong đơn hiện tại. Không thể thêm dòng trùng.");
    return;
  }
  refreshPackingTimelineBanners();
  var addTip = getPackingTimelineInfo(new Date());
  if (!window.confirm(
    'Thêm mã vào đơn?\n\n' + addTip.title + '\n' + addTip.shipMsg +
    '\n\nMã thêm mới sẽ ghi chú "Thêm mới vào đơn" và đi theo khung giờ hiện tại.'
  )) return;

  var inputEl = document.getElementById("ql-add-code");
  qlSelectedSuggestedItem = null;
  if (inputEl) inputEl.value = "";
  var qtyEl = document.getElementById("ql-add-qty");
  if (qtyEl) qtyEl.value = "1";
  var suggestBox = document.getElementById("ql-suggest-box");
  if (suggestBox) suggestBox.style.display = "none";

  var tb = document.getElementById("ql-tbody");
  if (tb) {
    var latestQty = Number(quantity) || 1;
    var tempKey = 'new-' + Date.now() + '-' + Math.floor(Math.random() * 1000);
    var dvtAdd = resolveDvtClient_(item.maHang, item.maVach, item.dvt);
    editRows.push({
      tempKey: tempKey,
      maHang: item.maHang,
      maVach: item.maVach,
      tenHang: item.tenHang,
      dvt: dvtAdd || item.dvt || "Cái",
      parentSku: item.parentSku || "",
      slGoc: latestQty,
      slThucTe: latestQty,
      stock: 0,
      ghiChu: "",
      trangThai: "Chưa lưu",
      nguoiSoanHang: ""
    });
    tb.insertAdjacentHTML('beforeend', '<tr><td>' + (editRows.length) + '</td><td>' + formatPackingCodeHtml_(editRows[editRows.length - 1]) + '</td><td>' + escapeHtml(formatPackingProductName_(editRows[editRows.length - 1])) + '</td><td>' + escapeHtml(formatDvtDisplay_(item.maHang, item.maVach, dvtAdd)) + '</td><td>0</td><td><input type="number" class="edit-sl-input" data-row="' + tempKey + '" data-new="1" value="' + latestQty + '" style="border:2px solid #1a73e8;text-align:center;width:70px;"></td><td><span style="color:#94a3b8;">-</span></td><td><span style="color:#94a3b8;">-</span></td><td><button type="button" onclick="ql_huyDong(\'' + tempKey + '\')" style="border:none; background:#d93025; color:white; border-radius:5px; padding:7px 9px; cursor:pointer;">Hủy mã</button></td></tr>');
  }

  alert("✅ Đã thêm mã vào bảng chỉnh sửa. Nhấn Lưu để ghi vào hệ thống.");
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
  if (itemHasVariantGroup_(item) && !itemOverride) {
    openVariantPicker_(item, 'ql', quantity);
    return;
  }
  ql_themMaHangDirect_(item, quantity);
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
  var sourceRows = (currentLoadedRows && currentLoadedRows.length) ? currentLoadedRows : (editRows || []);
  var items = [];
  function pushOut_(row, sl) {
    items.push({
      maHang: row.maHang,
      maVach: row.maVach,
      tenHang: row.tenHang,
      dvt: row.dvt,
      sl: sl,
      parentSku: row.parentSku || getParentSkuOfItem_(row) || '',
      variantSku: row.variantSku || row.maHang || '',
      variantName: row.variantName || row.tenHang || '',
      parentName: row.parentName || '',
      maHangDisplay: row.maHangDisplay || '',
      tenHangDisplay: row.tenHangDisplay || ''
    });
  }
  sourceRows.forEach(function(row) {
    if (!row) return;
    if (row.trangThai === "Đã hủy dòng" || row.trangThai === "Đã hủy đơn") return;
    // Dòng mới thêm trên UI (chưa lưu): lấy từ ô nhập
    if (row.tempKey || String(row.rowIndex || '').indexOf('temp') === 0) {
      var tip = document.querySelector('.edit-sl-input[data-row="' + (row.tempKey || row.rowIndex) + '"]');
      var tsl = tip ? Number(tip.value) : Number(row.slGoc) || 0;
      if (tsl > 0) pushOut_(row, tsl);
      return;
    }
    var sl = ql_resolvePrintQty_(row);
    if (Number(sl) > 0) pushOut_(row, sl);
  });
  return items;
}
function ql_inWeb_FromEdit() { executePrintWeb(currentPhieuObj.soPhieu, currentPhieuObj.khoXuat, currentPhieuObj.khoNhan, layItemsForOutput()); }
function ql_xuatExcel_FromEdit() { executeExportExcel(currentPhieuObj.soPhieu, currentPhieuObj.khoXuat, currentPhieuObj.khoNhan, layItemsForOutput()); }

// ================= SOẠN HÀNG MOBILE =================
function sh_taiDanhSachDon() {
  sh_ensureCreateDateDefaults_();
  var shNgayEl = document.getElementById("sh-ngay");
  if (shNgayEl && !shNgayEl.value) shNgayEl.value = formatDateInputValue(new Date());
  document.getElementById("sh-list-container").innerHTML = '<div class="card" style="text-align:center; color:gray;">Vui lòng chọn số phiếu ở trên.</div>';
  document.getElementById("sh-footer").style.display = "none";
  document.getElementById("sh-phieu").innerHTML = '<option value="">⏳ Đang tải...</option>';
  var tSh0 = Date.now();
  // #region agent log
  dbgBranch_('C', 'sh_taiDanhSachDon:start', 'packing list request', {
    role: sessionUser.role || '',
    store: sessionUser.store || '',
    ngay: shNgayEl ? shNgayEl.value : '',
    viewMode: 'packing'
  });
  // #endregion
  apiGet('getDonHangTheoNgay', { ngay: shNgayEl ? shNgayEl.value : '', userRole: sessionUser.role, userStore: sessionUser.store, viewMode: 'packing' }, { allowDirectFallback: true }).then(function(res) {
    var parsed = unwrapListResponse_(res);
    var rows = parsed.rows;
    // #region agent log
    dbgBranch_('C', 'sh_taiDanhSachDon:ok', 'packing list response', {
      ms: Date.now() - tSh0,
      rowCount: rows.length,
      isArray: Array.isArray(res),
      sample: rows[0] ? { soPhieu: rows[0].soPhieu, kx: rows[0].khoXuat, kn: rows[0].khoNhan, st: rows[0].trangThai } : null
    });
    // #endregion
    var countMoi = 0; var countDone = 0;
    var html = '<option value="">-- Chọn đơn ('+rows.length+') --</option>';
    rows.forEach(function(r) {
      if(r.trangThai === "Mới") countMoi++; else countDone++;
      var shortName = storeMap[r.khoNhan] || storeMap[r.khoXuat] || r.khoNhan || r.khoXuat || '';
      html += '<option value="'+r.soPhieu+'">'+r.soPhieu+' ('+shortName+') ['+r.trangThai+']</option>';
    });
    document.getElementById("sh-phieu").innerHTML = html;
    document.getElementById("sh-stats").innerHTML = '<div class="stat-box" style="color:#d93025;">🔔 CẦN SOẠN: '+countMoi+'</div> | <div class="stat-box" style="color:#137333;">✅ ĐÃ SOẠN: '+countDone+'</div>';
  }).catch(function(err){
    // #region agent log
    dbgBranch_('C', 'sh_taiDanhSachDon:err', 'packing list failed', {
      ms: Date.now() - tSh0,
      err: String(err && err.message || err)
    });
    // #endregion
    alert('Lỗi: '+err.message);
  });
}

var pendingImages = {}; var isCompressing = 0;

/** SL điền sẵn khi soạn / soạn lại: ưu tiên slSoan đã lưu, rồi slThucTe (cũ), cuối cùng SL đặt */
function sh_resolvePackInputQty_(it) {
  if (!it) return 0;
  if (it.slSoan !== undefined && it.slSoan !== null && it.slSoan !== "") return Number(it.slSoan) || 0;
  if (it.slThucTe !== undefined && it.slThucTe !== null && it.slThucTe !== "") return Number(it.slThucTe) || 0;
  return Number(it.slGoc) || 0;
}

function sh_chonDonMobile() {
  var sp = document.getElementById("sh-phieu").value; if(!sp) return;
  document.getElementById("sh-list-container").innerHTML = '<div style="text-align:center;">⏳ Đang tải SP...</div>';
  pendingImages = {}; isCompressing = 0;
  apiGet('getChiTietDonHangMobile', { soPhieu: sp }, { directOnly: true, timeoutMs: 60000 }).then(function(items) {
    // Fallback nếu API mobile trả rỗng / lỗi hình dạng
    if (!Array.isArray(items) || !items.length) {
      return apiGet('getChiTietPhieu', { soPhieu: sp, includeStock: '0' }, { directOnly: true, timeoutMs: 60000 }).then(function(res2) {
        var parsed = unwrapListResponse_(res2);
        return parsed.rows || [];
      });
    }
    return items;
  }).then(function(items) {
    var html = "";
    var lowStockCount = 0;
    if (!items || !items.length) {
      document.getElementById("sh-list-container").innerHTML = '<div class="card" style="color:#b91c1c;">Không tải được chi tiết đơn <b>' + escapeHtml(sp) + '</b>. Thử lại hoặc kiểm tra deploy code.gs.</div>';
      document.getElementById("sh-footer").style.display = "none";
      return;
    }
    (items || []).forEach(function(it, j) {
      var stockDisplay = (it.stock !== undefined && it.stock !== null && it.stock !== "") ? Number(it.stock) : "";
      // Ưu tiên SL đã soạn (slSoan / cột 16); không dùng lại SL đặt khi đã soạn trước đó
      var packQty = sh_resolvePackInputQty_(it);
      var compareQty = Math.max(Number(it.slGoc) || 0, packQty || 0);
      var short = stockShortageInfo_(stockDisplay, compareQty);
      if (short) lowStockCount++;
      var cardStyle = short ? 'border:1px solid #fecaca; background:#fef2f2;' : '';
      var stockLabel = (stockDisplay === "" || isNaN(stockDisplay))
        ? '<span style="color:#94a3b8;">-</span>'
        : ('<b class="sh-stock-val" style="color:' + (short ? '#b91c1c' : '#1d4ed8') + ';">' + stockDisplay + '</b>');
      var warnHtml = short
        ? ('<div class="sh-stock-warn" style="margin-top:6px; font-size:12px; color:#b91c1c; font-weight:700;">⚠️ Tồn thấp hơn SL đặt/soạn — thiếu ' + short.thieu + ' (vẫn lưu được)</div>')
        : '<div class="sh-stock-warn" style="display:none; margin-top:6px; font-size:12px; color:#b91c1c; font-weight:700;"></div>';
      var prevPackedNote = (it.slSoan !== undefined && it.slSoan !== null && it.slSoan !== "" && Number(it.slSoan) !== Number(it.slGoc))
        ? (' <small style="color:#c2410c;">(đã soạn: ' + Number(it.slSoan) + ')</small>')
        : '';
      html += '<div class="item-card" style="' + cardStyle + '" data-stock="' + (stockDisplay === "" || isNaN(stockDisplay) ? '' : stockDisplay) + '" data-slgoc="' + (Number(it.slGoc) || 0) + '">' +
        '<div style="margin-bottom:4px;">' + formatPackingCodeHtml_(it) + '</div>' +
        '<b>' + escapeHtml(formatPackingProductName_(it)) + '</b><br>' +
        '<small>ĐVT: ' + escapeHtml(formatDvtDisplay_(it.maHang, it.maVach, it.dvt)) + ' | Tồn hiện tại: ' + stockLabel + '</small>' +
        warnHtml +
        '<div class="action-row"><div>SL Yêu Cầu: <b>' + it.slGoc + '</b>' + prevPackedNote + '<br>SL Soạn: ' +
        '<input type="number" class="sl-thuc-te" data-row="' + it.rowIndex + '" value="' + packQty + '" oninput="sh_onPackQtyInput(this)">' +
        '</div><div style="text-align:right;"><label class="btn-camera" for="c-' + j + '">📷 Ảnh</label>' +
        '<input type="file" id="c-' + j + '" accept="image/*" capture="environment" style="display:none;" data-row="' + it.rowIndex + '" data-j="' + j + '" onchange="nenAnh(this)">' +
        '<img id="p-' + j + '" src="' + (it.anhXacNhan || '') + '" style="display:' + (it.anhXacNhan ? 'block' : 'none') + '; width:50px; height:50px; margin-top:5px;">' +
        '<small id="st-' + j + '"></small></div></div></div>';
    });
    if (lowStockCount > 0) {
      html = '<div style="margin:0 0 10px; padding:10px 12px; border-radius:10px; background:#fef2f2; border:1px solid #fecaca; color:#b91c1c; font-size:13px; font-weight:700;">⚠️ ' + lowStockCount + ' mã tồn thấp hơn SL đặt/soạn — chỉ cảnh báo, vẫn lưu được</div>' + html;
    }
    document.getElementById("sh-list-container").innerHTML = html; document.getElementById("sh-footer").style.display = "block";
  }).catch(function(err){ alert('Lỗi: '+err.message); });
}

function sh_onPackQtyInput(input) {
  if (!input) return;
  var card = input.closest('.item-card');
  if (!card) return;
  var stockRaw = card.getAttribute('data-stock');
  var slGoc = Number(card.getAttribute('data-slgoc')) || 0;
  var packQty = Number(input.value) || 0;
  var compareQty = Math.max(slGoc, packQty);
  var short = stockShortageInfo_(stockRaw === '' ? '' : stockRaw, compareQty);
  var stockEl = card.querySelector('.sh-stock-val');
  var warnEl = card.querySelector('.sh-stock-warn');
  if (stockEl) stockEl.style.color = short ? '#b91c1c' : '#1d4ed8';
  if (warnEl) {
    if (short) {
      warnEl.style.display = 'block';
      warnEl.innerText = '⚠️ Tồn thấp hơn SL đặt/soạn — thiếu ' + short.thieu + ' (vẫn lưu được)';
    } else {
      warnEl.style.display = 'none';
      warnEl.innerText = '';
    }
  }
  card.style.border = short ? '1px solid #fecaca' : '';
  card.style.background = short ? '#fef2f2' : '';
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
  // Cảnh báo tồn thấp chỉ mang tính thông báo — không chặn lưu soạn
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
var shBranchMergeHints_ = [];
var shBranchMergeConfirmed_ = false;
var shBranchMergeSkipped_ = false;

function sh_isOrderDefaultChecked_(order) {
  if (!order) return false;
  if (typeof order.defaultChecked === 'boolean') return order.defaultChecked;
  var st = String(order.trangThai || order.status || '').trim().toLowerCase();
  if (!st || st === 'mới' || st === 'moi' || st === 'chờ duyệt' || st === 'cho duyet' || st === 'chờ xử lý' || st === 'cho xu ly' || st === 'đang xử lý') return true;
  return false;
}

function sh_statusBadgeHtml_(order) {
  var label = String((order && (order.trangThai || order.status)) || 'Mới').trim() || 'Mới';
  var tone = String((order && order.statusTone) || '').trim().toLowerCase();
  if (!tone) {
    var st = label.toLowerCase();
    if (st.indexOf('hủy') !== -1 || st.indexOf('huy') !== -1) tone = 'danger';
    else if (st.indexOf('xác nhận') !== -1 || st.indexOf('xac nhan') !== -1 || st.indexOf('đã nhận') !== -1) tone = 'success';
    else if (st.indexOf('soạn') !== -1 || st.indexOf('soan') !== -1) tone = 'warn';
    else tone = 'info';
  }
  // info=blue (mới) | warn=yellow (đã soạn) | success=green (đã nhận) | danger=red (hủy)
  var bg = '#dbeafe';
  var fg = '#1d4ed8';
  if (tone === 'danger') { bg = '#fee2e2'; fg = '#b91c1c'; }
  else if (tone === 'warn') { bg = '#ffedd5'; fg = '#c2410c'; }
  else if (tone === 'success' || tone === 'ok') { bg = '#dcfce7'; fg = '#166534'; }
  return '<span class="sh-status-badge" style="display:inline-block;padding:3px 8px;border-radius:999px;font-size:11px;font-weight:800;background:' + bg + ';color:' + fg + ';">' + escapeHtml(label) + '</span>';
}

function sh_capNhatTomTatChonDonSoan() {
  var checkboxes = document.querySelectorAll('.sh-order-check');
  var checked = 0;
  checkboxes.forEach(function(cb) { if (cb.checked) checked++; });
  var countEl = document.getElementById('sh-order-picker-count');
  if (countEl) {
    countEl.innerText = 'Đã chọn ' + checked + ' / ' + checkboxes.length + ' đơn';
  }
  var summaryEl = document.getElementById('sh-order-picker-summary');
  if (summaryEl && summaryEl.getAttribute('data-base')) {
    summaryEl.innerText = summaryEl.getAttribute('data-base') + ' · Đã chọn ' + checked + '/' + checkboxes.length + ' đơn.';
  } else if (summaryEl) {
    summaryEl.innerText = 'Đã chọn ' + checked + '/' + checkboxes.length + ' đơn trong khung ca (đơn đặc biệt hiện đủ, mặc định bỏ tích).';
  }
}

function sh_renderDanhSachDonSoan(candidates, meta) {
  var bodyEl = document.getElementById('sh-order-picker-body');
  if (!bodyEl) return;
  var range = sh_getCreateDateRange_();
  var packingDay = (meta && meta.packingDay) || range.packingDay || range.to || '';
  var suggested = (meta && meta.suggestedPackingDay) || range.suggestedPackingDay || sh_getSuggestedPackingDayStr_();
  var windowHint = '';
  if (meta && (meta.mainWindow || meta.suppWindow || meta.totalWindow)) {
    windowHint = 'Chính: ' + (meta.mainWindow || '') + ' · Bổ sung: ' + (meta.suppWindow || '') +
      (meta.totalWindow ? (' · Tổng: ' + meta.totalWindow) : '');
  }
  if (!candidates || !candidates.length) {
    var emptyHint = 'Cửa sổ: hôm trước 10:00 → ngày giao 10:00. ' + windowHint;
    if (suggested && packingDay !== suggested) {
      emptyHint += ' · Gợi ý: đơn sau 10h nằm ở ngày tổng hợp ' + suggested + '.';
    }
    bodyEl.innerHTML = '<tr><td colspan="8" style="text-align:center; color:#b91c1c; padding:14px;">Không có đơn trong khung ca cho N2 <b>' + escapeHtml(packingDay) + '</b> / mode <b>' + escapeHtml(sh_getPackingMode_()) + '</b>.<br><small style="color:#64748b;">' + escapeHtml(emptyHint) + '</small></td></tr>';
    sh_capNhatTomTatChonDonSoan();
    return;
  }

  // Mode gộp đã xác nhận: 1 dòng / chi nhánh, SL mã trùng sẽ cộng dồn khi xuất bảng
  if (shBranchMergeConfirmed_) {
    bodyEl.innerHTML = sh_buildMergedBranchRowsHtml_(candidates);
    var summaryMerge = document.getElementById('sh-order-picker-summary');
    if (summaryMerge) {
      var baseM = 'N2 ' + packingDay + ' · Đã XÁC NHẬN GỘP cùng chi nhánh · 1 dòng/kho nhận · SL trùng mã CỘNG DỒN khi xuất';
      if (windowHint) baseM += ' · ' + windowHint;
      summaryMerge.setAttribute('data-base', baseM);
    }
    sh_capNhatTomTatChonDonSoan();
    return;
  }

  var html = '';
  var lastGroup = '';
  var modeNow = sh_getPackingMode_();
  candidates.forEach(function(order, idx) {
    var soPhieuSafe = String(order.soPhieu || '').replace(/\\/g, '\\\\').replace(/'/g, "\\'");
    var soAttr = encodeURIComponent(String(order.soPhieu || ''));
    var bucketLabel = order.packingBucketLabel
      || (order.packingBucket === 'bổ sung' ? 'Đơn Bổ Sung' : (order.packingBucket === 'chính' ? 'Đơn Chính' : ''));
    var bucketBadge = bucketLabel
      ? (' <span style="display:inline-block;margin-left:6px;padding:2px 7px;border-radius:999px;font-size:10px;font-weight:800;background:' +
        (order.packingBucket === 'bổ sung' ? '#ffedd5;color:#c2410c' : '#e0e7ff;color:#3730a3') +
        ';">[' + escapeHtml(bucketLabel) + ']</span>')
      : '';
    var mergeHint = order.mergeWithMain
      ? ' <small style="color:#0f766e;font-weight:600;">· gộp cùng kho</small>'
      : '';
    var checkedAttr = sh_isOrderDefaultChecked_(order) ? ' checked' : '';

    var gk = String(order.groupKey || order.khoNhan || '');
    if (modeNow === 'total' && gk && gk !== lastGroup) {
      lastGroup = gk;
      html += '<tr style="background:#f1f5f9;"><td colspan="8" style="padding:8px 10px;font-size:12px;font-weight:800;color:#334155;">📦 Kho nhận: ' +
        formatPackingStoreCellHtml_(order.khoNhan) + '</td></tr>';
    }

    html += '<tr>' +
      '<td style="text-align:center; font-weight:700; color:#334155;">' + (idx + 1) + '</td>' +
      '<td style="text-align:center;"><input type="checkbox" class="sh-order-check" data-sophieu="' + escapeHtml(order.soPhieu) + '"' + checkedAttr + ' onchange="sh_capNhatTomTatChonDonSoan()"></td>' +
      '<td><a href="javascript:void(0)" class="order-pdf-link" onclick="openOrderPdfView(\'' + soAttr + '\')">' + escapeHtml(order.soPhieu) + '</a>' + bucketBadge + mergeHint + '</td>' +
      '<td>' + formatPackingStoreCellHtml_(order.khoXuat) + '</td>' +
      '<td>' + formatPackingStoreCellHtml_(order.khoNhan) + '</td>' +
      '<td>' + sh_statusBadgeHtml_(order) + '</td>' +
      '<td>' + escapeHtml(order.thoiGianDat || '-') + '</td>' +
      '<td style="text-align:center;">' +
        '<button type="button" class="btn-submit" style="width:auto; padding:6px 10px; border-radius:8px; box-shadow:none; background:#0f766e; font-size:12px;" onclick="sh_inDonTuBangSoan(\'' + soPhieuSafe + '\')">🖨️ In</button>' +
      '</td>' +
      '</tr>';
  });
  bodyEl.innerHTML = html;
  var summaryEl = document.getElementById('sh-order-picker-summary');
  if (summaryEl) {
    var mode = modeNow;
    var modeText = mode === 'main' ? 'Đơn chính' : (mode === 'supp' ? 'Bổ sung' : 'Tổng hợp ca');
    var base = 'N2 ' + packingDay + ' · Mode: ' + modeText + ' · Cột kho tách khỏi trạng thái';
    if (windowHint) base += ' · ' + windowHint;
    base += ' · Click số phiếu để xem PDF';
    if (suggested && packingDay !== suggested) {
      base += ' · Đang lệch ngày gợi ý theo giờ (' + suggested + ')';
    }
    summaryEl.setAttribute('data-base', base);
  }
  sh_capNhatTomTatChonDonSoan();
}

/** Render 1 dòng / chi nhánh sau khi xác nhận gộp */
function sh_buildMergedBranchRowsHtml_(candidates) {
  var byBranch = {};
  var order = [];
  (candidates || []).forEach(function(o) {
    if (!o) return;
    var gk = String(o.groupKey || o.khoNhan || '').trim();
    if (!gk) gk = '_unknown';
    if (!byBranch[gk]) {
      byBranch[gk] = {
        groupKey: gk,
        khoNhan: o.khoNhan,
        khoXuat: o.khoXuat,
        soPhieuList: [],
        defaultChecked: false,
        thoiGianDat: o.thoiGianDat || ''
      };
      order.push(gk);
    }
    byBranch[gk].soPhieuList.push(o.soPhieu);
    if (sh_isOrderDefaultChecked_(o)) byBranch[gk].defaultChecked = true;
  });
  var html = '';
  order.forEach(function(gk, idx) {
    var g = byBranch[gk];
    var codes = g.soPhieuList || [];
    var codesAttr = escapeHtml(codes.join(','));
    var checkedAttr = g.defaultChecked ? ' checked' : '';
    var label = codes.length > 1
      ? (escapeHtml(codes[0]) + ' <small style="color:#0f766e;font-weight:700;">+ ' + (codes.length - 1) + ' đơn gộp</small>')
      : escapeHtml(codes[0] || '-');
    var pdfFirst = encodeURIComponent(String(codes[0] || ''));
    html += '<tr style="background:#ecfdf5;">' +
      '<td style="text-align:center;font-weight:700;">' + (idx + 1) + '</td>' +
      '<td style="text-align:center;"><input type="checkbox" class="sh-order-check sh-order-check-merged" data-sophieu-list="' + codesAttr + '" data-sophieu="' + escapeHtml(codes[0] || '') + '"' + checkedAttr + ' onchange="sh_capNhatTomTatChonDonSoan()"></td>' +
      '<td><a href="javascript:void(0)" class="order-pdf-link" onclick="openOrderPdfView(\'' + pdfFirst + '\')">' + label + '</a>' +
        '<div style="font-size:11px;color:#64748b;margin-top:2px;">' + escapeHtml(codes.join(', ')) + '</div></td>' +
      '<td>' + formatPackingStoreCellHtml_(g.khoXuat) + '</td>' +
      '<td><b>' + formatPackingStoreCellHtml_(g.khoNhan) + '</b> <small style="color:#0f766e;">· 1 lần</small></td>' +
      '<td><span class="sh-status-badge" style="display:inline-block;padding:3px 8px;border-radius:999px;font-size:11px;font-weight:800;background:#dcfce7;color:#166534;">Gộp ' + codes.length + ' đơn</span></td>' +
      '<td>' + escapeHtml(g.thoiGianDat || '-') + '</td>' +
      '<td style="text-align:center;color:#94a3b8;font-size:12px;">—</td>' +
      '</tr>';
  });
  return html;
}

function sh_detectBranchMergeHints_(candidates, apiHints) {
  if (apiHints && apiHints.length) return apiHints;
  var by = {};
  (candidates || []).forEach(function(o) {
    if (!o) return;
    var gk = String(o.groupKey || o.khoNhan || '').trim();
    if (!gk) return;
    if (!by[gk]) by[gk] = { groupKey: gk, khoNhan: o.khoNhan, khoNhanLabel: formatStoreShortLabel_(o.khoNhan) || o.khoNhan, orderCount: 0, soPhieuList: [] };
    by[gk].orderCount++;
    by[gk].soPhieuList.push(o.soPhieu);
  });
  var hints = [];
  Object.keys(by).forEach(function(k) {
    if (by[k].orderCount >= 2) hints.push(by[k]);
  });
  return hints;
}

function sh_showBranchMergeModal_(hints) {
  var modal = document.getElementById('modal-branch-merge');
  var msg = document.getElementById('branch-merge-msg');
  var list = document.getElementById('branch-merge-list');
  if (!modal || !list) return;
  var lines = (hints || []).map(function(h) {
    var name = h.khoNhanLabel || formatStoreShortLabel_(h.khoNhan) || h.khoNhan || h.groupKey;
    return '<div style="padding:8px 0;border-bottom:1px solid #fed7aa;"><b>' + escapeHtml(name) + '</b> — <b>' + h.orderCount + '</b> đơn' +
      '<div style="color:#9a3412;font-size:12px;margin-top:2px;">' + escapeHtml((h.soPhieuList || []).join(', ')) + '</div></div>';
  }).join('');
  if (msg) {
    var first = hints[0];
    var sampleName = (first && (first.khoNhanLabel || formatStoreShortLabel_(first.khoNhan) || first.khoNhan)) || 'chi nhánh';
    msg.innerHTML = 'Phát hiện Chi nhánh <b>' + escapeHtml(sampleName) + '</b>' +
      (hints.length > 1 ? (' và <b>' + (hints.length - 1) + '</b> chi nhánh khác') : '') +
      ' có ≥ 2 đơn trong ca. Bấm <b>Xác Nhận Gộp</b> để gom thành 1 phiếu tổng hợp duy nhất theo chi nhánh (tránh hiển thị trùng; SL mã trùng sẽ <b>CỘNG DỒN</b>).';
  }
  list.innerHTML = lines || '<div>Không có chi tiết.</div>';
  modal.style.display = 'flex';
}

function sh_confirmBranchMerge_() {
  shBranchMergeConfirmed_ = true;
  shBranchMergeSkipped_ = false;
  var modal = document.getElementById('modal-branch-merge');
  if (modal) modal.style.display = 'none';
  sh_renderDanhSachDonSoan(shOrderCandidates, {
    packingDay: (sh_getCreateDateRange_().packingDay || ''),
    mainWindow: '',
    suppWindow: '',
    totalWindow: ''
  });
}

function sh_skipBranchMerge_() {
  shBranchMergeConfirmed_ = false;
  shBranchMergeSkipped_ = true;
  var modal = document.getElementById('modal-branch-merge');
  if (modal) modal.style.display = 'none';
}

/** In từng đơn từ bảng chọn đơn soạn hàng (cùng mẫu in tab Quản lý) */
function sh_inDonTuBangSoan(soPhieu) {
  var code = String(soPhieu || '').trim();
  if (!code) return alert('Thiếu số phiếu.');
  var fromList = (shOrderCandidates || []).find(function(o) {
    return o && String(o.soPhieu || '').trim() === code;
  });
  showLoad('Đang tải đơn ' + code + ' để in...');
  apiGet('getChiTietPhieu', {
    soPhieu: code,
    storeName: (fromList && fromList.khoXuat) || (sessionUser && sessionUser.store) || '',
    includeStock: '0'
  }, { directOnly: true, timeoutMs: 45000 }).then(function(res) {
    hideLoad();
    var parsed = unwrapListResponse_(res);
    var rows = parsed.rows || [];
    var items = [];
    rows.forEach(function(row) {
      if (!row) return;
      if (row.trangThai === 'Đã hủy dòng' || row.trangThai === 'Đã hủy đơn') return;
      var sl = ql_resolvePrintQty_(row);
      if (Number(sl) > 0) {
        items.push({
          maHang: row.maHang,
          maVach: row.maVach,
          tenHang: row.tenHang,
          dvt: resolveDvtClient_(row.maHang, row.maVach, row.dvt),
          sl: sl
        });
      }
    });
    if (!items.length) {
      alert('Đơn ' + code + ' không có dòng hợp lệ để in.');
      return;
    }
    var khoXuat = (fromList && fromList.khoXuat) || '';
    var khoNhan = (fromList && fromList.khoNhan) || '';
    executePrintWeb(code, khoXuat, khoNhan, items);
  }).catch(function(err) {
    hideLoad();
    alert('Lỗi tải đơn để in: ' + (err && err.message || err));
  });
}


function sh_taiDanhSachDonSoanChoBang() {
  var range = sh_getCreateDateRange_();
  var ngay = range.from || '';
  var ngayTo = range.to || ngay;
  var packingMode = sh_getPackingMode_();
  var pickerEl = document.getElementById('sh-order-picker');
  if (pickerEl) pickerEl.style.display = 'block';
  var bodyEl = document.getElementById('sh-order-picker-body');
  if (bodyEl) bodyEl.innerHTML = '<tr><td colspan="7" style="text-align:center; color:#64748b; padding:14px;">Đang tải danh sách đơn...</td></tr>';

  showLoad('Đang tải danh sách đơn soạn...');
  apiGet('getDanhSachDonSoanHang', {
    ngay: ngay,
    ngayTo: ngayTo,
    packingMode: packingMode,
    includePacked: '1', // alias — BE luôn trả đủ trạng thái trong khung ca
    userRole: sessionUser.role || '',
    userStore: sessionUser.store || ''
  }, { allowDirectFallback: true, timeoutMs: 120000 }).then(function(res) {
    hideLoad();
    if (!res || !res.success) {
      throw new Error((res && (res.error || res.msg)) || 'Không thể tải danh sách đơn.');
    }
    shOrderCandidates = Array.isArray(res.orders) ? res.orders : [];
    shBranchMergeConfirmed_ = false;
    shBranchMergeSkipped_ = false;
    shBranchMergeHints_ = sh_detectBranchMergeHints_(shOrderCandidates, res.branchMergeHints || []);
    var suggested = range.suggestedPackingDay || sh_getSuggestedPackingDayStr_();
    // #region agent log
    dbgSend_('PD1', 'sh_taiDanhSachDonSoanChoBang:ok', 'packing picker list', {
      selectedPackingDay: res.packingDay || ngayTo,
      suggestedPackingDay: suggested,
      userTouched: !!shCreateDateUserTouched_,
      total: shOrderCandidates.length,
      mergeHints: shBranchMergeHints_.length,
      hasDH435102: shOrderCandidates.some(function(o) { return o && String(o.soPhieu || '').indexOf('435102') !== -1; }),
      sample: shOrderCandidates.slice(0, 8).map(function(o) { return o && o.soPhieu; })
    });
    // #endregion
    sh_renderDanhSachDonSoan(shOrderCandidates, {
      packingDay: res.packingDay || ngayTo,
      suggestedPackingDay: suggested,
      mainWindow: res.mainWindow || '',
      suppWindow: res.suppWindow || '',
      totalWindow: res.totalWindow || ''
    });
    if (shBranchMergeHints_.length && (packingMode === 'total' || packingMode === 'main' || packingMode === 'supp')) {
      sh_showBranchMergeModal_(shBranchMergeHints_);
    }
    sh_warmStockInBackground_();
  }).catch(function(err) {
    hideLoad();
    shOrderCandidates = [];
    if (bodyEl) bodyEl.innerHTML = '<tr><td colspan="7" style="text-align:center; color:#b91c1c; padding:14px;">Lỗi tải danh sách: ' + escapeHtml(err.message) + '</td></tr>';
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
      if (st && st.ready) {
        shStockWarmState.ready = true;
        shStockWarmState.warming = false;
        shStockWarmState.promise = null;
        if (summaryEl) summaryEl.innerText = 'Tồn Q7 sẵn sàng (' + (st.source || 'TON_Q7') + ', ' + (st.stores || 0) + ' mã).';
        return true;
      }
      // Chưa có TON_Q7 → rebuild 1 lần từ sheet tổng (có thể lâu), rồi các lần sau đọc sheet nhẹ
      return apiGet('warmStockIndex', null, { directOnly: true, timeoutMs: 180000 }).then(function(w) {
        shStockWarmState.warming = false;
        shStockWarmState.promise = null;
        shStockWarmState.ready = !!(w && w.success && (w.ready !== false));
        shStockWarmState.lastMs = (w && w._debugTotalMs) || 0;
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
  // Mỗi lần mở: nếu user chưa tự chọn ngày thì sync theo khung giờ hiện tại
  if (!shCreateDateUserTouched_) {
    var toEl = document.getElementById('sh-create-to');
    if (toEl) toEl.value = sh_getSuggestedPackingDayStr_();
  }
  sh_taiDanhSachDonSoanChoBang();
}

function sh_toggleOnlyNewItemsUi() {
  sh_onPackingModeChange();
}

function sh_taoBangSoanTuDonDaChon() {
  if (!sessionUser || !sessionUser.user) {
    alert("Vui lòng đăng nhập trước khi tạo bảng soạn.");
    return;
  }
  var range = sh_getCreateDateRange_();
  var ngay = range.from || "";
  var selectedOrders = [];
  document.querySelectorAll('.sh-order-check').forEach(function(cb) {
    if (!cb.checked) return;
    var listAttr = cb.getAttribute('data-sophieu-list') || '';
    if (listAttr) {
      listAttr.split(',').forEach(function(sp) {
        sp = String(sp || '').trim();
        if (sp && selectedOrders.indexOf(sp) === -1) selectedOrders.push(sp);
      });
      return;
    }
    var soPhieu = cb.getAttribute('data-sophieu') || '';
    if (soPhieu && selectedOrders.indexOf(soPhieu) === -1) selectedOrders.push(soPhieu);
  });
  if (!selectedOrders.length) {
    alert('Vui lòng tick ít nhất 1 đơn để tạo bảng soạn.');
    return;
  }

  var packingMode = sh_getPackingMode_();
  var onlyNewItems = packingMode === 'supp';
  var afterTimeEl = document.getElementById('sh-new-after-time');
  var beforeTimeEl = document.getElementById('sh-new-before-time');
  var newAfterTime = afterTimeEl && afterTimeEl.value ? afterTimeEl.value : '08:00';
  var newBeforeTime = beforeTimeEl && beforeTimeEl.value ? beforeTimeEl.value : '10:00';

  var modeLabel = packingMode === 'main' ? 'đơn chính' : (packingMode === 'supp' ? 'bổ sung' : 'tổng hợp ca');
  var clientStart = Date.now();
  showLoad("Bước 1/2: Kiểm tra tồn Q7 (" + modeLabel + ")...");

  sh_ensureStockReady_().then(function(stockOk) {
    showLoad(stockOk ? "Bước 2/2: Đang tạo bảng (đọc TON_Q7)..." : "Bước 2/2: Đang tạo bảng (rebuild TON_Q7 nếu thiếu)...");
    var packingDay = range.packingDay || range.to || ngay;
    var payload = {
      ngay: range.from || ngay,
      ngayTo: packingDay,
      actor: sessionUser.user,
      userRole: sessionUser.role || '',
      userStore: sessionUser.store || '',
      selectedOrders: selectedOrders,
      forceStock: !stockOk,
      packingMode: packingMode,
      onlyNewItems: onlyNewItems,
      newAfterTime: newAfterTime,
      newBeforeTime: newBeforeTime || '10:00',
      mergeSameBranch: !!shBranchMergeConfirmed_
    };
    // proxyFirst: tránh CORS redirect POST→GET trên browser (lỗi 502 HTML)
    var postOpts = stockOk
      ? { proxyFirst: true, timeoutMs: 60000 }
      : { proxyFirst: true, timeoutMs: 90000 };
    // #region agent log
    dbgSend_('P502', 'sh_taoBangSoanTuDonDaChon:post', 'create packing table request', {
      selectedCount: selectedOrders.length,
      packingDay: packingDay,
      stockOk: !!stockOk,
      onlyNewItems: !!onlyNewItems,
      sample: selectedOrders.slice(0, 8)
    });
    // #endregion
    return apiPost('taoBangSoanHangNgayMai', payload, postOpts);
  }).then(function(res) {
    hideLoad();
    var clientMs = Date.now() - clientStart;
    if (!res || !res.success) {
      alert("❌ Tạo bảng thất bại:\n" + ((res && (res.msg || res.error)) || "Không rõ lỗi") + "\n[Build FE: " + APP_BUILD + "]" + (res && res._debugRun ? (" [GAS: " + res._debugRun + "]") : ""));
      return;
    }
    if (res.stockReady) shStockWarmState.ready = true;
    var modeRes = res.packingMode || packingMode;
    var modeText = modeRes === 'main'
      ? ("Đơn chính (" + (res.mainWindowLabel || "N1 10:00 → N2 08:00") + ")")
      : (modeRes === 'supp'
        ? ("Bổ sung ≥08:00 & <10:00 (" + (res.newAfterLabel || "") + " → " + (res.newBeforeLabel || "") + ")")
        : ("Tổng hợp ca (" + (res.totalWindowLabel || res.mainWindowLabel || "") + ")"));
    var msg = "✅ Đã tạo tab: " + (res.sheetName || "SoanNgayMai") + "\n" +
      "- Ngày tổng hợp N2: " + (res.packingDay || range.packingDay || range.to || "") + "\n" +
      "- Chế độ: " + modeText + "\n" +
      "- Tổng đơn: " + (res.totalOrders || 0) + "\n" +
      "- Tổng mã: " + (res.totalItems || 0) + "\n" +
      "- Mã thiếu: " + (res.missingItems || 0) + "\n" +
      "- Tồn kho: " + (res.stockReady ? ("CÓ (" + (res.stockSource || "TON_Q7") + ")") : "KHÔNG — Admin import lại file tồn để tạo sheet TON_Q7");
    if (res._debugIncludedMainRows != null) msg += "\n- Dòng chính lấy: " + res._debugIncludedMainRows;
    if (res._debugIncludedNewRows != null) msg += "\n- Dòng bổ sung lấy: " + res._debugIncludedNewRows;
    if (res._debugSkippedByTime) msg += "\n- Bỏ qua (ngoài khung giờ): " + res._debugSkippedByTime;
    if (res._debugTotalMs) msg += "\n(Server tạo bảng: " + Math.round(res._debugTotalMs / 1000) + "s)";
    msg += "\n(Tổng chờ: " + Math.round(clientMs / 1000) + "s)\n[" + APP_BUILD + (res._debugRun ? (" / " + res._debugRun) : "") + "]";
    alert(msg);
    if (res.url) window.open(res.url, '_blank', 'noopener,noreferrer');
  }).catch(function(err) {
    hideLoad();
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

var admNewProductRows_ = [];
var admNewProductFilterQ_ = '';

function admNormalizeSearch_(value) {
  var s = '';
  try {
    s = String(value || '').normalize('NFC');
  } catch (e) {
    s = String(value || '');
  }
  return s
    .toLowerCase()
    .normalize('NFKD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/\s+/g, ' ')
    .trim();
}

function admNormalizeCode_(value) {
  try {
    return String(value || '').normalize('NFC').toUpperCase().trim();
  } catch (e) {
    return String(value || '').toUpperCase().trim();
  }
}

function adm_loadNewProductFlags() {
  if (sessionUser.role !== 'Admin') return;
  var meta = document.getElementById('adm-new-meta');
  if (meta) meta.innerText = 'Đang tải danh mục...';
  showLoad('Đang tải danh mục Hàng Mới...');
  // Quét gần như toàn bộ Data_Excel (không cắt 250 như trước)
  apiGet('getCatalogIsNewAdminList', { q: '', limit: 20000 }, { timeoutMs: 90000 }).then(function(res) {
    hideLoad();
    if (!res || !res.success) throw new Error((res && res.error) || 'Không tải được danh mục.');
    admNewProductRows_ = Array.isArray(res.items) ? res.items : [];
    adm_filterNewProductFlags();
  }).catch(function(err) {
    hideLoad();
    if (meta) meta.innerText = 'Lỗi: ' + (err && err.message || err);
    alert('Lỗi tải Hàng Mới: ' + (err && err.message || err));
  });
}

function adm_itemMatchesNewProductQuery_(it, qRaw) {
  if (!it) return false;
  var qCode = admNormalizeCode_(qRaw);
  var qFold = admNormalizeSearch_(qRaw);
  if (!qCode && !qFold) return true;
  var ma = String(it.maHang || it.maSP || '');
  var mv = String(it.maVach || '');
  var th = String(it.tenHang || '');
  var parent = String(it.parentSku || '');
  var dvt = String(it.dvt || '');
  // SKU / barcode: match NFC uppercase (không phân biệt hoa/thường)
  if (qCode) {
    if (admNormalizeCode_(ma).indexOf(qCode) !== -1) return true;
    if (admNormalizeCode_(mv).indexOf(qCode) !== -1) return true;
    if (admNormalizeCode_(parent).indexOf(qCode) !== -1) return true;
  }
  // Tên / toàn bộ: fold dấu + NFC
  var hay = admNormalizeSearch_([ma, mv, th, parent, dvt].join(' '));
  return !!(qFold && hay.indexOf(qFold) !== -1);
}

function adm_getFilteredNewProductIndexes_() {
  var qEl = document.getElementById('searchNewProductInput') || document.getElementById('adm-new-q');
  var qRaw = qEl ? String(qEl.value || '') : String(admNewProductFilterQ_ || '');
  try {
    qRaw = qRaw.normalize('NFC');
  } catch (e) {}
  admNewProductFilterQ_ = admNormalizeSearch_(qRaw);
  var indexes = [];
  for (var i = 0; i < admNewProductRows_.length; i++) {
    var it = admNewProductRows_[i];
    if (!it) continue;
    if (!String(qRaw || '').trim()) {
      indexes.push(i);
      continue;
    }
    if (adm_itemMatchesNewProductQuery_(it, qRaw)) indexes.push(i);
  }
  return indexes;
}

function adm_filterNewProductFlags() {
  var qEl = document.getElementById('searchNewProductInput') || document.getElementById('adm-new-q');
  var qRaw = qEl ? String(qEl.value || '').trim() : '';
  var indexes = adm_getFilteredNewProductIndexes_();
  // Nếu local không khớp (danh sách cũ / chưa reload) → hỏi server theo q
  if (qRaw && !indexes.length) {
    var meta = document.getElementById('adm-new-meta');
    if (meta) meta.innerText = 'Đang tìm trên server: "' + qRaw + '"...';
    apiGet('getCatalogIsNewAdminList', { q: qRaw, limit: 500 }, { timeoutMs: 60000 }).then(function(res) {
      if (!res || !res.success) {
        adm_renderNewProductFlags_();
        return;
      }
      var remote = Array.isArray(res.items) ? res.items : [];
      if (!remote.length) {
        adm_renderNewProductFlags_();
        return;
      }
      // Merge remote vào cache theo maHang (không mất tick cục bộ nếu trùng)
      var byKey = {};
      admNewProductRows_.forEach(function(row, idx) {
        var k = admNormalizeCode_(row && (row.maHang || row.maSP));
        if (k) byKey[k] = idx;
      });
      remote.forEach(function(row) {
        var k = admNormalizeCode_(row && (row.maHang || row.maSP));
        if (!k) return;
        if (byKey.hasOwnProperty(k)) {
          var prev = admNewProductRows_[byKey[k]];
          if (prev && typeof prev.isNew === 'boolean') row.isNew = prev.isNew;
          admNewProductRows_[byKey[k]] = row;
        } else {
          byKey[k] = admNewProductRows_.length;
          admNewProductRows_.push(row);
        }
      });
      adm_renderNewProductFlags_();
    }).catch(function() {
      adm_renderNewProductFlags_();
    });
    return;
  }
  adm_renderNewProductFlags_();
}

function adm_renderNewProductFlags_() {
  var tb = document.getElementById('adm-new-tbody');
  var meta = document.getElementById('adm-new-meta');
  if (!tb) return;
  if (!admNewProductRows_.length) {
    tb.innerHTML = '<tr><td colspan="6" style="text-align:center;color:#64748b;padding:16px;">Không có sản phẩm. Nhấn "Tải danh sách".</td></tr>';
    if (meta) meta.innerText = 'Chưa có dữ liệu.';
    return;
  }
  var indexes = adm_getFilteredNewProductIndexes_();
  if (!indexes.length) {
    tb.innerHTML = '<tr><td colspan="6" style="text-align:center;color:#64748b;padding:16px;">Không có dòng khớp từ khóa.</td></tr>';
    if (meta) meta.innerText = 'Lọc: 0 / ' + admNewProductRows_.length + ' sản phẩm.';
    return;
  }
  var html = '';
  indexes.forEach(function(idx) {
    var it = admNewProductRows_[idx];
    var checked = it.isNew ? ' checked' : '';
    var parent = String(it.parentSku || '').trim();
    var parentAttr = encodeURIComponent(parent || it.maHang || '');
    var btn = parent
      ? ('<button type="button" class="btn-submit" style="width:auto;padding:6px 10px;box-shadow:none;background:#7c3aed;font-size:12px;" onclick="vm_openForParent(\'' + parentAttr + '\')">Sửa Biến Thể</button>')
      : ('<button type="button" class="btn-submit" style="width:auto;padding:6px 10px;box-shadow:none;background:#64748b;font-size:12px;" onclick="vm_openForParent(\'' + encodeURIComponent(it.maHang || '') + '\')">Sửa Biến Thể</button>');
    html += '<tr class="adm-isnew-row" data-idx="' + idx + '">' +
      '<td style="text-align:center;"><input type="checkbox" class="adm-isnew-check" data-idx="' + idx + '"' + checked +
        ' onchange="adm_onNewProductCheckChange_(this)"></td>' +
      '<td><b>' + escapeHtml(it.maHang || '-') + '</b>' +
        (parent ? ('<br><small style="color:#7c3aed;">Parent: ' + escapeHtml(parent) + '</small>') : '') + '</td>' +
      '<td>' + escapeHtml(it.maVach || '-') + '</td>' +
      '<td>' + escapeHtml(it.tenHang || '-') + '</td>' +
      '<td>' + escapeHtml(it.dvt || '') + '</td>' +
      '<td style="text-align:center;">' + btn + '</td>' +
      '</tr>';
  });
  tb.innerHTML = html;
  if (meta) {
    var qShow = (document.getElementById('searchNewProductInput') || {}).value || '';
    meta.innerText = 'Hiển thị ' + indexes.length + ' / ' + admNewProductRows_.length +
      (qShow ? (' · lọc: "' + qShow + '"') : '') + '. Tick rồi nhấn Lưu.';
  }
}

function adm_onNewProductCheckChange_(cb) {
  if (!cb) return;
  var idx = Number(cb.getAttribute('data-idx'));
  if (!admNewProductRows_[idx]) return;
  admNewProductRows_[idx].isNew = !!cb.checked;
}

function adm_setVisibleNewProductFlags(checked) {
  var want = !!checked;
  var indexes = adm_getFilteredNewProductIndexes_();
  indexes.forEach(function(idx) {
    if (admNewProductRows_[idx]) admNewProductRows_[idx].isNew = want;
  });
  document.querySelectorAll('.adm-isnew-check').forEach(function(cb) {
    cb.checked = want;
    var idx = Number(cb.getAttribute('data-idx'));
    if (admNewProductRows_[idx]) admNewProductRows_[idx].isNew = want;
  });
  var meta = document.getElementById('adm-new-meta');
  if (meta) meta.innerText = (want ? 'Đã CheckAll ' : 'Đã UncheckAll ') + indexes.length + ' dòng đang hiện. Nhấn Lưu để ghi sheet.';
}

function adm_saveNewProductFlags() {
  if (sessionUser.role !== 'Admin') return alert('Chỉ Admin được lưu Hàng Mới.');
  if (!admNewProductRows_.length) return alert('Chưa có danh sách để lưu. Hãy tải danh sách trước.');
  // Đồng bộ checkbox đang hiện → mảng, rồi lưu toàn bộ danh sách đã tải
  document.querySelectorAll('.adm-isnew-check').forEach(function(cb) {
    var idx = Number(cb.getAttribute('data-idx'));
    if (admNewProductRows_[idx]) admNewProductRows_[idx].isNew = !!cb.checked;
  });
  var flags = admNewProductRows_.map(function(row) {
    return {
      sheetRow: row.sheetRow,
      maHang: row.maHang,
      maVach: row.maVach,
      isNew: !!row.isNew
    };
  });
  showLoad('Đang lưu Hàng Mới...');
  apiPost('saveCatalogIsNewFlags', { flags: flags, actor: sessionUser.user }, { timeoutMs: 90000 }).then(function(res) {
    hideLoad();
    if (!res || !res.success) throw new Error((res && (res.error || res.msg)) || 'Lưu thất bại.');
    alert('✅ ' + (res.msg || ('Đã cập nhật ' + (res.changed || 0) + ' sản phẩm.')));
    loadCatalogInBackground(true);
    loadNewProductsInBackground();
    var meta = document.getElementById('adm-new-meta');
    if (meta) meta.innerText = 'Đã lưu · thay đổi: ' + (res.changed || 0) + ' · tổng dòng gửi: ' + flags.length;
  }).catch(function(err) {
    hideLoad();
    alert('Lỗi lưu Hàng Mới: ' + (err && err.message || err));
  });
}

// ================= ADMIN: QUẢN LÝ BIẾN THỂ =================
var admVariantGroups_ = [];
var vmState_ = { parentSku: '', rows: [] };

function adm_loadVariantGroups() {
  if (sessionUser.role !== 'Admin') return;
  var meta = document.getElementById('adm-variant-meta');
  if (meta) meta.innerText = 'Đang tải nhóm Parent_SKU...';
  showLoad('Đang tải nhóm biến thể...');
  apiGet('getParentVariantGroupsAdmin', { q: '', limit: 300 }, { timeoutMs: 60000 }).then(function(res) {
    hideLoad();
    if (!res || !res.success) throw new Error((res && res.error) || 'Không tải được nhóm.');
    admVariantGroups_ = Array.isArray(res.items) ? res.items : [];
    adm_filterVariantGroups_();
  }).catch(function(err) {
    hideLoad();
    if (meta) meta.innerText = 'Lỗi: ' + (err && err.message || err);
    alert('Lỗi tải nhóm biến thể: ' + (err && err.message || err));
  });
}

function adm_filterVariantGroups_() {
  var tb = document.getElementById('adm-variant-tbody');
  var meta = document.getElementById('adm-variant-meta');
  if (!tb) return;
  var q = admNormalizeSearch_((document.getElementById('adm-variant-q') || {}).value || '');
  var rows = admVariantGroups_.filter(function(it) {
    if (!q) return true;
    return admNormalizeSearch_([it.parentSku, it.sampleTen].join(' ')).indexOf(q) !== -1;
  });
  if (!admVariantGroups_.length) {
    tb.innerHTML = '<tr><td colspan="4" style="text-align:center;color:#64748b;padding:16px;">Chưa có Parent_SKU trong Data_Excel.</td></tr>';
    if (meta) meta.innerText = '0 nhóm.';
    return;
  }
  if (!rows.length) {
    tb.innerHTML = '<tr><td colspan="4" style="text-align:center;color:#64748b;padding:16px;">Không khớp từ khóa.</td></tr>';
    if (meta) meta.innerText = 'Lọc 0 / ' + admVariantGroups_.length;
    return;
  }
  var html = '';
  rows.forEach(function(it) {
    var pAttr = encodeURIComponent(it.parentSku || '');
    html += '<tr><td><b style="color:#7c3aed;">' + escapeHtml(it.parentSku) + '</b></td>' +
      '<td>' + escapeHtml(it.sampleTen || '-') + '</td>' +
      '<td style="text-align:center;font-weight:700;">' + (it.childCount || 0) + '</td>' +
      '<td style="text-align:center;"><button type="button" class="btn-submit" style="width:auto;padding:7px 12px;box-shadow:none;background:#7c3aed;font-size:12px;" onclick="vm_openForParent(\'' + pAttr + '\')">Sửa Biến Thể</button></td></tr>';
  });
  tb.innerHTML = html;
  if (meta) meta.innerText = 'Hiển thị ' + rows.length + ' / ' + admVariantGroups_.length + ' nhóm.';
}

function vm_openForParent(parentEncoded) {
  var parentSku = '';
  try { parentSku = decodeURIComponent(parentEncoded || ''); } catch (e) { parentSku = String(parentEncoded || ''); }
  parentSku = String(parentSku || '').trim().toUpperCase();
  if (!parentSku) return alert('Thiếu Parent_SKU.');
  vmState_.parentSku = parentSku;
  vmState_.rows = [];
  var modal = document.getElementById('modal-variant-manager');
  var label = document.getElementById('vm-parent-label');
  var tbody = document.getElementById('vm-tbody');
  if (label) label.textContent = parentSku;
  if (tbody) tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:16px;color:#64748b;">Đang tải mã con...</td></tr>';
  if (modal) modal.style.display = 'flex';
  showLoad('Đang tải biến thể ' + parentSku + '...');
  apiGet('getChildVariantsForAdmin', { parentSku: parentSku }, { timeoutMs: 60000 }).then(function(res) {
    hideLoad();
    var list = (res && res.success && Array.isArray(res.variants)) ? res.variants : [];
    if (!list.length) {
      // fallback local catalog
      list = (variantIndexByParent[parentSku] || []).map(function(v) {
        return { maHang: v.maHang, maVach: v.maVach, tenHang: v.tenHang, dvt: v.dvt, tonBanDau: 0, tonHienTai: 0 };
      });
    }
    vmState_.rows = list.map(function(v) {
      return {
        maHang: v.maHang || '',
        maVach: v.maVach || '',
        tenHang: v.tenHang || '',
        dvt: v.dvt || '',
        tonBanDau: Number(v.tonBanDau) || 0,
        tonHienTai: (v.tonHienTai != null ? Number(v.tonHienTai) : Number(v.stock)) || 0
      };
    });
    if (!vmState_.rows.length) vmState_.rows.push({ maHang: '', maVach: '', tenHang: '', dvt: '', tonBanDau: 0, tonHienTai: 0 });
    vm_renderRows_();
  }).catch(function(err) {
    hideLoad();
    alert('Lỗi tải biến thể: ' + (err && err.message || err));
  });
}

function vm_renderRows_() {
  var tbody = document.getElementById('vm-tbody');
  if (!tbody) return;
  var html = '';
  vmState_.rows.forEach(function(r, idx) {
    html += '<tr>' +
      '<td><input type="text" id="vm-mh-' + idx + '" value="' + escapeHtml(r.maHang) + '" style="width:110px;"></td>' +
      '<td><input type="text" id="vm-th-' + idx + '" value="' + escapeHtml(r.tenHang) + '" style="width:180px;"></td>' +
      '<td><input type="text" id="vm-mv-' + idx + '" value="' + escapeHtml(r.maVach) + '" style="width:120px;"></td>' +
      '<td><input type="text" id="vm-dvt-' + idx + '" value="' + escapeHtml(r.dvt) + '" style="width:70px;"></td>' +
      '<td><input type="number" id="vm-ton-' + idx + '" value="' + (Number(r.tonBanDau) || 0) + '" style="width:90px;"></td>' +
      '<td style="text-align:center;font-weight:700;color:#166534;">' + (Number(r.tonHienTai) || 0) + '</td>' +
      '<td><button type="button" onclick="vm_removeRow_(' + idx + ')" style="border:none;background:#fee2e2;color:#b91c1c;border-radius:6px;padding:6px 8px;cursor:pointer;">Xóa</button></td>' +
      '</tr>';
  });
  tbody.innerHTML = html;
}

function vm_syncFromInputs_() {
  var next = [];
  for (var i = 0; i < vmState_.rows.length; i++) {
    var mhEl = document.getElementById('vm-mh-' + i);
    if (!mhEl) continue;
    next.push({
      maHang: (document.getElementById('vm-mh-' + i) || {}).value || '',
      tenHang: (document.getElementById('vm-th-' + i) || {}).value || '',
      maVach: (document.getElementById('vm-mv-' + i) || {}).value || '',
      dvt: (document.getElementById('vm-dvt-' + i) || {}).value || '',
      tonBanDau: Number((document.getElementById('vm-ton-' + i) || {}).value) || 0,
      tonHienTai: vmState_.rows[i] ? (Number(vmState_.rows[i].tonHienTai) || 0) : 0
    });
  }
  vmState_.rows = next;
}

function vm_addEmptyRow_() {
  vm_syncFromInputs_();
  vmState_.rows.push({ maHang: '', maVach: '', tenHang: '', dvt: '', tonBanDau: 0, tonHienTai: 0 });
  vm_renderRows_();
}

function vm_removeRow_(idx) {
  vm_syncFromInputs_();
  vmState_.rows.splice(idx, 1);
  if (!vmState_.rows.length) vmState_.rows.push({ maHang: '', maVach: '', tenHang: '', dvt: '', tonBanDau: 0, tonHienTai: 0 });
  vm_renderRows_();
}

function vm_close_() {
  var modal = document.getElementById('modal-variant-manager');
  if (modal) modal.style.display = 'none';
}

function vm_save_() {
  if (sessionUser.role !== 'Admin') return alert('Chỉ Admin được sửa biến thể.');
  vm_syncFromInputs_();
  var parentSku = vmState_.parentSku;
  var variants = vmState_.rows.filter(function(r) { return String(r.maHang || '').trim(); }).map(function(r) {
    return {
      maHang: String(r.maHang || '').trim(),
      maVach: String(r.maVach || '').trim(),
      tenHang: String(r.tenHang || '').trim(),
      dvt: String(r.dvt || '').trim(),
      tonBanDau: Number(r.tonBanDau) || 0
    };
  });
  if (!variants.length) return alert('Cần ít nhất 1 mã con (Mã hàng).');
  showLoad('Đang lưu biến thể...');
  apiPost('saveChildVariants', {
    parentSku: parentSku,
    variants: variants,
    actor: sessionUser.user
  }, { timeoutMs: 90000 }).then(function(res) {
    hideLoad();
    if (!res || !res.success) throw new Error((res && (res.error || res.msg)) || 'Lưu thất bại.');
    alert('✅ ' + (res.msg || 'Đã lưu biến thể.'));
    vm_close_();
    loadCatalogInBackground(true);
    adm_loadVariantGroups();
  }).catch(function(err) {
    hideLoad();
    alert('Lỗi lưu biến thể: ' + (err && err.message || err));
  });
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
  // Đ/đ không tách trong NFKD — map sang D/d trước khi lọc a-z, nếu không
  // "Đơn vị tính" → "onvitinh" và "ĐVT" → "vt" (mất cột ĐVT khi nhập khẩu).
  return String(value || '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D')
    .toLowerCase()
    .normalize('NFKD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function dbgBranch_(hypothesisId, location, message, data) {
  // #region agent log
  dbgSend_(hypothesisId, location, message, data);
  // #endregion
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
  var dvtColIdx = impFindColByAliases_(header, ['donvitinh', 'dvtinh', 'dvt', 'donvi', 'unit', 'uom']);
  if (importType === 'stock' && dvtColIdx < 0) warnings.push('Không thấy cột Đơn vị tính (ĐVT) trong tiêu đề — sheet TON_Q7 sẽ thiếu ĐVT.');
  return {
    header: header,
    detectedType: detectedType,
    warnings: warnings,
    rowCount: rows ? rows.length : 0,
    colCount: header ? header.length : 0,
    dvtColIdx: dvtColIdx
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

function impFindColByAliases_(header, aliases) {
  if (!header) return -1;
  for (var c = 0; c < header.length; c++) {
    var n = impNormalizeText(header[c]).replace(/\s+/g, '');
    if (!n) continue;
    for (var a = 0; a < aliases.length; a++) {
      if (n.indexOf(aliases[a]) !== -1) return c;
    }
  }
  return -1;
}

function impIsQ7Header_(value) {
  var n = impNormalizeText(value).replace(/\s+/g, '');
  if (!n) return false;
  if (n.indexOf('q7') !== -1 || n.indexOf('quan7') !== -1) return true;
  if (n.indexOf('khodiadiemkinhdoanhq7') !== -1) return true;
  if (n.indexOf('k9quan7') !== -1) return true;
  return false;
}

function impIsQ7StoreName_(value) {
  return impIsQ7Header_(value);
}

function impParseQty_(value) {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return isNaN(value) ? 0 : value;
  var text = String(value).trim().replace(/\s+/g, '').replace(/\u00A0/g, '');
  if (!text) return 0;
  if (/^\d{1,3}(\.\d{3})+(,\d+)?$/.test(text)) text = text.replace(/\./g, '').replace(',', '.');
  else text = text.replace(/,/g, '');
  var num = Number(text);
  return isNaN(num) ? 0 : num;
}

function impNormCode_(value) {
  return String(value || '').trim().toUpperCase();
}

function impNormDvtKey_(value) {
  return impNormalizeText(value).replace(/\s+/g, '');
}

function impAddStockEntry_(map, prefix, code, qty, dvt) {
  var norm = impNormCode_(code);
  if (!norm) return;
  var q = Number(qty) || 0;
  if (!q) return;
  var dvtRaw = String(dvt || '').trim();
  var dvtKey = impNormDvtKey_(dvtRaw);
  var key = prefix + norm + (dvtKey ? ('|DV:' + dvtKey) : '');
  if (!map[key]) map[key] = { q: 0, d: dvtRaw };
  map[key].q += q;
  if (dvtRaw && !map[key].d) map[key].d = dvtRaw;
}

/** Tách catalog (mã/tên/ĐVT) từ file nhập khẩu thông tin → payload nhỏ */
function extractCatalogEntriesFromRows_(rows) {
  if (!rows || rows.length < 2) return { entries: [], meta: { reason: 'empty' } };
  var headerIndex = 0;
  var bestScore = -1;
  for (var hi = 0; hi < Math.min(rows.length, 10); hi++) {
    var joined = (rows[hi] || []).map(impNormalizeText).join(' ');
    var score = (rows[hi] || []).filter(function(cell) { return impNormalizeText(cell).length > 0; }).length;
    if (/ma hang|mahang|sku|mavach|ma vach/.test(joined)) score += 6;
    if (/ten hang|tenhang|name/.test(joined)) score += 4;
    if (/dvt|don vi|donvi|unit|uom/.test(joined)) score += 5;
    if (score > bestScore) { bestScore = score; headerIndex = hi; }
  }
  var header = rows[headerIndex] || [];
  var maHangIdx = impFindColByAliases_(header, ['mahanghoa', 'mahang', 'sku', 'mahh', 'itemcode']);
  var maVachIdx = impFindColByAliases_(header, ['mavach', 'barcode', 'ean']);
  var tenHangIdx = impFindColByAliases_(header, ['tenhanghoa', 'tenhang', 'name', 'description']);
  var dvtIdx = impFindColByAliases_(header, ['donvitinh', 'dvtinh', 'dvt', 'donvi', 'unit', 'uom', 'basicunit', 'unitname']);
  if (dvtIdx < 0) {
    for (var dc = 0; dc < header.length; dc++) {
      var hn = impNormalizeText(header[dc]);
      if (hn.indexOf('don vi') !== -1 || hn === 'dv' || hn.indexOf('dvt') !== -1) { dvtIdx = dc; break; }
    }
  }
  // Cột Parent_SKU / Mã hàng hóa cha / Ma_Nhom_Ban — đọc riêng, không dùng làm mã hàng chính
  var parentIdx = -1;
  for (var pc = 0; pc < header.length; pc++) {
    var pn = impNormalizeText(header[pc]).replace(/\s+/g, '');
    if (!pn) continue;
    if (pn === 'parentsku' || pn === 'parent' || pn === 'manhomban' || pn === 'nhomban') { parentIdx = pc; break; }
    if ((pn.indexOf('mahang') !== -1 || pn.indexOf('mahanghoa') !== -1) && (pn.indexOf('cha') !== -1 || pn.indexOf('parent') !== -1)) {
      parentIdx = pc; break;
    }
    if (pn.indexOf('parentsku') !== -1 || pn.indexOf('manhomban') !== -1) { parentIdx = pc; break; }
  }
  if (maHangIdx >= 0) {
    var mhHeader = impNormalizeText(header[maHangIdx]).replace(/\s+/g, '');
    if (mhHeader.indexOf('cha') !== -1 || mhHeader.indexOf('parent') !== -1 || maHangIdx === parentIdx) {
      var altMh = -1;
      for (var ac = 0; ac < header.length; ac++) {
        if (ac === maHangIdx || ac === parentIdx) continue;
        var an = impNormalizeText(header[ac]).replace(/\s+/g, '');
        if ((an.indexOf('mahang') !== -1 || an === 'sku') && an.indexOf('cha') === -1 && an.indexOf('parent') === -1) {
          altMh = ac; break;
        }
      }
      if (altMh >= 0) maHangIdx = altMh;
    }
  }

  var entries = [];
  var withDvt = 0;
  var withParent = 0;
  var startRow = headerIndex + 1;
  for (var r = startRow; r < rows.length; r++) {
    var row = rows[r];
    if (!row) continue;
    var mh = maHangIdx >= 0 ? String(row[maHangIdx] == null ? '' : row[maHangIdx]).trim() : '';
    var mv = maVachIdx >= 0 ? String(row[maVachIdx] == null ? '' : row[maVachIdx]).trim() : '';
    var th = tenHangIdx >= 0 ? String(row[tenHangIdx] == null ? '' : row[tenHangIdx]).trim() : '';
    var d = dvtIdx >= 0 ? String(row[dvtIdx] == null ? '' : row[dvtIdx]).trim() : '';
    var p = parentIdx >= 0 ? String(row[parentIdx] == null ? '' : row[parentIdx]).trim() : '';
    if (!mh && !mv) continue;
    if (!th && !d && !mh && !mv) continue;
    if (d) withDvt++;
    if (p) withParent++;
    entries.push({ mh: mh, mv: mv, th: th, d: d, p: p });
  }
  var headerNorm = (header || []).map(function(h) {
    return { raw: String(h == null ? '' : h).slice(0, 40), norm: impNormalizeText(h).replace(/\s+/g, '') };
  });
  // #region agent log
  dbgBranch_('A', 'extractCatalogEntriesFromRows_', 'catalog header dvt detect post-fix', {
    runId: 'post-fix',
    headerIndex: headerIndex,
    dvtIdx: dvtIdx,
    withDvt: withDvt,
    entryCount: entries.length,
    headerNorm: headerNorm.slice(0, 16),
    sampleEntry: entries[0] || null,
    dvtNormOk: headerNorm.some(function(h) { return h.norm.indexOf('donvitinh') !== -1 || h.norm === 'dvt' || h.norm.indexOf('dvt') === 0; })
  });
  // #endregion
  return {
    entries: entries,
    meta: {
      headerIndex: headerIndex,
      headerSample: (header || []).slice(0, 12),
      maHangIdx: maHangIdx,
      maVachIdx: maVachIdx,
      tenHangIdx: tenHangIdx,
      dvtIdx: dvtIdx,
      parentIdx: parentIdx,
      entryCount: entries.length,
      withDvt: withDvt,
      withParent: withParent
    }
  };
}

/** Tách tồn Kho Q7 từ file Excel ngay trên trình duyệt → payload nhỏ */
function extractTonQ7EntriesFromRows_(rows) {
  if (!rows || rows.length < 2) return { entries: [], meta: { reason: 'empty' } };
  // Ưu tiên dòng header có cả mã hàng + ĐVT (tránh nhận nhầm dòng tiêu đề báo cáo)
  var headerIndex = 0;
  var bestScore = -1;
  for (var hi = 0; hi < Math.min(rows.length, 12); hi++) {
    var joined = (rows[hi] || []).map(impNormalizeText).join(' ');
    var score = (rows[hi] || []).filter(function(cell) { return impNormalizeText(cell).length > 0; }).length;
    if (/ma hang|mahang|sku|mavach|ma vach/.test(joined)) score += 5;
    if (/dvt|don vi|donvi|unit|uom|dvtinh/.test(joined)) score += 8;
    if (/ton kho|tonkho|stock|q7|quan 7/.test(joined)) score += 3;
    if (score > bestScore) { bestScore = score; headerIndex = hi; }
  }
  var header = rows[headerIndex] || [];
  var maHangIdx = impFindColByAliases_(header, ['mahanghoa', 'mahang', 'sku', 'mahh', 'itemcode']);
  var maVachIdx = impFindColByAliases_(header, ['mavach', 'barcode', 'ean']);
  var dvtIdx = impFindColByAliases_(header, ['donvitinh', 'dvtinh', 'dvt', 'donvi', 'unit', 'uom', 'basicunit', 'unitname']);
  // Fallback: cột có chữ "đơn vị" trong header
  if (dvtIdx < 0) {
    for (var dc = 0; dc < header.length; dc++) {
      var hn = impNormalizeText(header[dc]);
      if (hn.indexOf('don vi') !== -1 || hn === 'dv' || hn.indexOf('dvt') !== -1) { dvtIdx = dc; break; }
    }
  }
  var tonIdx = impFindColByAliases_(header, ['tonkho', 'soluongton', 'stock', 'onhand', 'slton', 'cuoiky']);
  var tenHangIdx = impFindColByAliases_(header, ['tenhanghoa', 'tenhang', 'name', 'description']);
  var q7Cols = [];
  for (var c = 0; c < header.length; c++) {
    if (c === maHangIdx || c === maVachIdx || c === dvtIdx || c === tonIdx || c === tenHangIdx) continue;
    if (impIsQ7Header_(header[c])) q7Cols.push(c);
  }

  var startRow = headerIndex + 1;
  var marker = rows[startRow] || [];
  var markerHits = 0;
  for (var m = 0; m < marker.length; m++) {
    var mv = marker[m];
    if (typeof mv === 'number' && mv < 0) markerHits++;
    else if (/^\(\d+\)/.test(String(mv || '').trim())) markerHits++;
  }
  if (markerHits >= 2) startRow++;

  var map = {};
  var currentMaHang = '';
  var currentMaVach = '';
  var currentDvt = '';
  var matchedRows = 0;
  var rowsWithDvt = 0;

  for (var r = startRow; r < rows.length; r++) {
    var row = rows[r];
    if (!row) continue;
    var maHang = maHangIdx >= 0 ? row[maHangIdx] : '';
    var maVach = maVachIdx >= 0 ? row[maVachIdx] : '';
    var dvt = dvtIdx >= 0 ? row[dvtIdx] : '';
    var hasOwn = !!(String(maHang || '').trim() || String(maVach || '').trim());
    if (hasOwn) {
      currentMaHang = maHang;
      currentMaVach = maVach;
      if (String(dvt || '').trim()) currentDvt = dvt;
    }
    var useMaHang = hasOwn ? maHang : currentMaHang;
    var useMaVach = hasOwn ? maVach : currentMaVach;
    var useDvt = String(dvt || '').trim() ? dvt : currentDvt;
    if (!String(useMaHang || '').trim() && !String(useMaVach || '').trim()) continue;

    var got = false;
    if (q7Cols.length) {
      for (var qi = 0; qi < q7Cols.length; qi++) {
        var qty = impParseQty_(row[q7Cols[qi]]);
        if (!qty) continue;
        impAddStockEntry_(map, 'MH:', useMaHang, qty, useDvt);
        impAddStockEntry_(map, 'MV:', useMaVach, qty, useDvt);
        got = true;
      }
    } else if (tonIdx >= 0) {
      // Dạng dòng con: tên kho nằm ở cột tên hàng / cột đầu
      var storeCell = tenHangIdx >= 0 ? row[tenHangIdx] : row[0];
      if (impIsQ7StoreName_(storeCell)) {
        var qty2 = impParseQty_(row[tonIdx]);
        if (qty2) {
          impAddStockEntry_(map, 'MH:', useMaHang, qty2, useDvt);
          impAddStockEntry_(map, 'MV:', useMaVach, qty2, useDvt);
          got = true;
        }
      }
    }
    if (got) {
      matchedRows++;
      if (String(useDvt || '').trim()) rowsWithDvt++;
    }
  }

  var entries = [];
  for (var k in map) {
    if (!Object.prototype.hasOwnProperty.call(map, k)) continue;
    entries.push({ k: k, q: map[k].q, d: map[k].d || '' });
  }
  return {
    entries: entries,
    meta: {
      headerIndex: headerIndex,
      headerSample: (header || []).slice(0, 12),
      q7Cols: q7Cols,
      matchedRows: matchedRows,
      rowsWithDvt: rowsWithDvt,
      entryCount: entries.length,
      maHangIdx: maHangIdx,
      maVachIdx: maVachIdx,
      dvtIdx: dvtIdx,
      tonIdx: tonIdx
    }
  };
}

function imp_handleSelectionChange() {
  var typeEl = document.getElementById('imp-type');
  var fileEl = document.getElementById('imp-file');
  var importType = typeEl && typeEl.value ? typeEl.value : 'stock';
  var fullWrap = document.getElementById('imp-full-stock-wrap');
  if (fullWrap) fullWrap.style.display = importType === 'stock' ? 'flex' : 'none';
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

  var allRows = importPreviewState.rows || [];
  var rowCount = allRows.length;
  if (!rowCount) return alert("File không có dữ liệu.");
  var wantFullStock = !!(document.getElementById('imp-full-stock') && document.getElementById('imp-full-stock').checked);
  var t0 = Date.now();

  // ===== STOCK: mặc định import nhanh chỉ TON_Q7 =====
  if (importType === 'stock' && !wantFullStock) {
    showLoad("Đang tách tồn Q7 từ file...");
    var extracted;
    try {
      extracted = extractTonQ7EntriesFromRows_(allRows);
    } catch (ex) {
      hideLoad();
      alert('Lỗi tách tồn Q7: ' + (ex && ex.message || ex));
      return;
    }
    if (!extracted.entries.length) {
      hideLoad();
      alert("Không tách được tồn Kho Q7 từ file.\nKiểm tra file có cột/kho Q7, hoặc tick ghi full sheet.");
      return;
    }
    showLoad("Đang ghi " + extracted.entries.length + " mã Q7 lên TON_Q7...");
    apiPost('nhapKhauCapNhatThongTin', {
      importType: 'stockQ7',
      fileName: file.name,
      q7Entries: extracted.entries,
      actor: sessionUser.user
    }, { directOnly: true, timeoutMs: 120000 }).then(function(res) {
      hideLoad();
      if (!res || !res.success) {
        alert("❌ Nhập khẩu thất bại: " + ((res && (res.error || res.msg)) || "Không rõ lỗi") + "\n[Build: " + APP_BUILD + "]");
        return;
      }
      var dvtColOk = extracted.meta && extracted.meta.dvtIdx >= 0;
      alert("✅ Cập nhật nhanh TON_Q7 thành công!\n" +
        "- Dòng TON_Q7: " + (res.q7Rows || extracted.entries.length) + "\n" +
        "- Biến thể TON_VARIANT: " + (res.variantRows || 0) + " (UPSERT, giữ Da_Xuat / mã con)\n" +
        "- Có ĐVT: " + (res.q7WithDvt != null ? res.q7WithDvt : (extracted.meta.rowsWithDvt || 0)) +
        (dvtColOk ? (" (cột ĐVT #" + (extracted.meta.dvtIdx + 1) + ")") : " (⚠️ không thấy cột ĐVT trong file)") + "\n" +
        "- Thời gian: " + Math.round((Date.now() - t0) / 1000) + "s\n" +
        (res.msg ? ("\n" + res.msg + "\n") : "") +
        "[" + APP_BUILD + " / import-q7-variant-upsert-v3]");
      if (fileEl) fileEl.value = '';
      importPreviewState = null;
      renderImportPreview(null);
      shStockWarmState.ready = true;
    }).catch(function(err) {
      hideLoad();
      alert('Lỗi: ' + err.message + '\n[Build: ' + APP_BUILD + ']\nDeploy lại code.gs rồi thử lại.');
    });
    return;
  }

  // ===== CATALOG nhanh: tách cột cần thiết phía trình duyệt (tránh timeout như tồn Q7) =====
  if (importType === 'catalog') {
    showLoad("Đang tách mã/tên/ĐVT từ file...");
    var catExtracted;
    try {
      catExtracted = extractCatalogEntriesFromRows_(allRows);
    } catch (catEx) {
      hideLoad();
      alert('Lỗi tách catalog: ' + (catEx && catEx.message || catEx));
      return;
    }
    if (!catExtracted.entries.length) {
      hideLoad();
      alert("Không tách được dòng hàng từ file nhập khẩu thông tin.\nKiểm tra có cột mã hàng / mã vạch / tên hàng.");
      return;
    }
    var CAT_CHUNK = 1500;
    var catChunks = [];
    for (var ci = 0; ci < catExtracted.entries.length; ci += CAT_CHUNK) {
      catChunks.push(catExtracted.entries.slice(ci, ci + CAT_CHUNK));
    }
    var catWritten = 0;
    var catLastRes = null;
    function sendCatChunk(idx) {
      var part = catChunks[idx] || [];
      var approxBytes = JSON.stringify(part).length;
      showLoad("Đang ghi catalog " + (idx + 1) + "/" + catChunks.length + " (" + catExtracted.entries.length + " dòng)...");
      return apiPost('nhapKhauCapNhatThongTin', {
        importType: 'catalogFast',
        fileName: file.name,
        catalogEntries: part,
        chunkIndex: idx,
        chunkTotal: catChunks.length,
        actor: sessionUser.user
      }, { directOnly: true, timeoutMs: 180000 }).then(function(res) {
        if (!res || !res.success) {
          throw new Error((res && (res.error || res.msg)) || ("Chunk catalog " + (idx + 1) + " thất bại"));
        }
        catWritten += Number(res.updatedRows) || 0;
        catLastRes = res;
        if (idx + 1 < catChunks.length) return sendCatChunk(idx + 1);
        return res;
      });
    }
    sendCatChunk(0).then(function(res) {
      hideLoad();
      res = res || catLastRes;
      if (!res || !res.success) {
        alert("❌ Nhập khẩu thất bại: " + ((res && (res.error || res.msg)) || "Không rõ lỗi") + "\n[Build: " + APP_BUILD + "]");
        return;
      }
      var dvtOk = catExtracted.meta && catExtracted.meta.dvtIdx >= 0;
      alert("✅ UPSERT danh mục thành công (không xóa mã con)!\n" +
        "- Staging: MISA_IMPORT\n" +
        "- Dòng xử lý: " + catWritten + "\n" +
        "- Cập nhật / Thêm: " + (res.updated || '-') + " / " + (res.appended || '-') + "\n" +
        "- Giữ Parent_SKU: " + (res.preservedParent || 0) + "\n" +
        "- TON_VARIANT meta: " + (res.variantRows || 0) + "\n" +
        "- Có ĐVT: " + (catExtracted.meta.withDvt || 0) +
        (dvtOk ? (" (cột ĐVT #" + (catExtracted.meta.dvtIdx + 1) + ")") : " (⚠️ không thấy cột ĐVT)") + "\n" +
        "- Chunk: " + catChunks.length + "\n" +
        "- Thời gian: " + Math.round((Date.now() - t0) / 1000) + "s\n" +
        (res.msg ? ("\n" + res.msg + "\n") : "") +
        "[" + APP_BUILD + " / import-catalog-upsert-v2]");
      if (fileEl) fileEl.value = '';
      importPreviewState = null;
      renderImportPreview(null);
      clearCatalogLocalStorage();
      loadCatalogInBackground(true);
    }).catch(function(err) {
      hideLoad();
      var tip = /hết thời gian|timeout|abort/i.test(String(err && err.message || ''))
        ? 'File quá lớn hoặc GAS chậm — thử lại; nếu vẫn lỗi thì Deploy code.gs (New version).'
        : 'Nếu báo không nhận action catalogFast: Deploy lại code.gs (New version).';
      alert('Lỗi nhập khẩu thông tin: ' + (err && err.message || err) + '\n[Build: ' + APP_BUILD + ']\n' + tip);
    });
    return;
  }

  // ===== STOCK full (chậm) =====
  var HEADER_PREFIX = Math.min(8, rowCount);
  var BODY_CHUNK = 700;
  var prefix = allRows.slice(0, HEADER_PREFIX);
  var body = allRows.slice(HEADER_PREFIX);
  if (!body.length) body = [[]];
  var chunks = [];
  for (var i = 0; i < body.length; i += BODY_CHUNK) {
    chunks.push(body.slice(i, i + BODY_CHUNK));
  }
  if (!chunks.length) chunks = [[]];

  var totalWritten = 0;
  var lastRes = null;

  function sendChunk(idx) {
    var bodyPart = chunks[idx] || [];
    var writeRows = idx === 0 ? prefix.concat(bodyPart) : bodyPart;
    var parseRows = prefix.concat(bodyPart);
    showLoad("Đang tải lên " + (idx + 1) + "/" + chunks.length + " (" + rowCount + " dòng)...");
    return apiPost('nhapKhauCapNhatThongTin', {
      importType: importType,
      fileName: file.name,
      sourceSheet: importPreviewState.sheetName,
      fileData: writeRows,
      parseMatrix: parseRows,
      chunkIndex: idx,
      chunkTotal: chunks.length,
      actor: sessionUser.user
    }, { directOnly: true, timeoutMs: 180000 }).then(function(res) {
      if (!res || !res.success) {
        throw new Error((res && (res.error || res.msg)) || ("Chunk " + (idx + 1) + " thất bại"));
      }
      totalWritten += Number(res.updatedRows) || 0;
      lastRes = res;
      if (idx + 1 < chunks.length) return sendChunk(idx + 1);
      return res;
    });
  }

  sendChunk(0).then(function(res) {
    hideLoad();
    res = res || lastRes;
    if (!res || !res.success) {
      alert("❌ Nhập khẩu thất bại: " + ((res && (res.error || res.msg)) || "Không rõ lỗi") + "\n[Build: " + APP_BUILD + "]");
      return;
    }
    var targetSheet = res.targetSheet || 'TỔNG HỢP TỒN KHO';
    var msg = "✅ Cập nhật thành công!\n" +
      "- Sheet đích: " + targetSheet + "\n" +
      "- Số dòng ghi: " + totalWritten + "\n" +
      "- Số chunk: " + chunks.length + "\n" +
      "- Thời gian: " + Math.round((Date.now() - t0) / 1000) + "s";
    msg += "\n- Sheet TON_Q7: " + (res.q7Rows || 0) + " mã/ĐVT";
    if (res.msg) msg += "\n\n" + res.msg;
    msg += "\n[" + APP_BUILD + "]";
    alert(msg);
    if (fileEl) fileEl.value = '';
    importPreviewState = null;
    renderImportPreview(null);
    shStockWarmState.ready = !!(res.q7Rows > 0);
  }).catch(function(err) {
    hideLoad();
    alert('Lỗi: ' + (err && err.message || err) + '\n[Build: ' + APP_BUILD + ']');
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

// ================= BÁN KÈM DỊCH VỤ / XUẤT BÁN HÀNG =================
var xbItems = [];
var xbInvoiceLocked = "";

function xb_applyStoreQuyen_() {
  var el = document.getElementById("xb-chi-nhanh");
  if (!el) return;
  if (sessionUser && sessionUser.role !== "Admin" && sessionUser.store) {
    el.value = sessionUser.store;
    el.setAttribute("disabled", "true");
  } else {
    el.removeAttribute("disabled");
  }
}

function xb_setItemsEnabled_(enabled) {
  var card = document.getElementById("xb-items-card");
  var banner = document.getElementById("xb-invoice-banner");
  var lockBtn = document.getElementById("xb-btn-lock");
  var hdInput = document.getElementById("xb-so-hoa-don");
  var cnEl = document.getElementById("xb-chi-nhanh");
  if (card) {
    card.style.opacity = enabled ? "1" : "0.55";
    card.style.pointerEvents = enabled ? "auto" : "none";
  }
  if (banner) banner.style.display = enabled ? "block" : "none";
  if (lockBtn) lockBtn.style.display = enabled ? "none" : "block";
  if (hdInput) hdInput.disabled = !!enabled;
  if (cnEl && enabled) cnEl.setAttribute("disabled", "true");
  else xb_applyStoreQuyen_();
}

function xb_onTabOpen() {
  xb_applyStoreQuyen_();
  xb_renderTable();
  xb_setItemsEnabled_(!!xbInvoiceLocked);
  xb_loadRecent();
  var scan = document.getElementById("xb-input-scan");
  if (xbInvoiceLocked && scan) scan.focus();
  else {
    var hd = document.getElementById("xb-so-hoa-don");
    if (hd) hd.focus();
  }
}

function xb_xacNhanHoaDon() {
  var hdEl = document.getElementById("xb-so-hoa-don");
  var cnEl = document.getElementById("xb-chi-nhanh");
  var soHdRaw = hdEl ? String(hdEl.value || "").trim() : "";
  if (!soHdRaw) return alert("Vui lòng nhập số hóa đơn liên kết (từ phần mềm khác).");
  if (!cnEl || !cnEl.value) return alert("Vui lòng chọn chi nhánh xuất bán.");
  // Giữ nguyên Đ trong số HĐ MISA (HĐCXXX1007 ≠ HDCXXX1007)
  var soHd = normalizeMisaCode_(soHdRaw) || soHdRaw;
  if (hdEl) hdEl.value = soHd;
  xbInvoiceLocked = soHd;
  var lockedLbl = document.getElementById("xb-locked-invoice");
  if (lockedLbl) lockedLbl.textContent = soHd + " · " + (storeMap[cnEl.value] || cnEl.value);
  xb_setItemsEnabled_(true);
  var scan = document.getElementById("xb-input-scan");
  if (scan) {
    scan.placeholder = "Quét / tìm mã bán kèm cho HĐ " + soHd;
    scan.focus();
  }
}

function xb_doiHoaDon() {
  if (xbItems.length && !confirm("Đổi hóa đơn sẽ xóa danh sách mã đang nhập. Tiếp tục?")) return;
  xbItems = [];
  xbInvoiceLocked = "";
  xb_renderTable();
  xb_setItemsEnabled_(false);
  var hdEl = document.getElementById("xb-so-hoa-don");
  if (hdEl) {
    hdEl.disabled = false;
    hdEl.focus();
    hdEl.select();
  }
  var scan = document.getElementById("xb-input-scan");
  if (scan) scan.placeholder = "Nhập số HĐ trước, rồi mới thêm mã...";
}

function xb_handleSearchInput(e) {
  if (!xbInvoiceLocked) {
    alert("Nhập và xác nhận số hóa đơn liên kết trước khi thêm mã.");
    return;
  }
  if (!catalogLoadState.ready) {
    var box0 = document.getElementById("xb-suggest-box");
    if (box0) {
      box0.innerHTML = '<div class="suggest-empty">Danh mục hàng đang tải. Vui lòng đợi vài giây...</div>';
      box0.style.display = "block";
    }
    return;
  }
  var inputEl = document.getElementById("xb-input-scan");
  var box = document.getElementById("xb-suggest-box");
  if (!inputEl || !box) return;
  positionSuggestionBox(box, inputEl);
  var val = inputEl.value.trim();
  if (val.length < 1) { box.style.display = "none"; return; }

  if (e && e.key === "Enter") {
    var exactKey = normalizeMisaCode_(val) || val.toUpperCase();
    var exactMatch = danhMucGoc[exactKey] || danhMucGoc[val.toUpperCase()];
    if (exactMatch) xb_chonSanPham(exactMatch);
    else {
      var matched = filterProducts(val);
      if (matched.length > 0) xb_chonSanPham(matched[0]);
      else {
        xbItems.unshift({ maHang: "LỖI MÃ", maVach: val, tenHang: "❌ Không tồn tại", dvt: "Lỗi", sl: "1" });
        xb_renderTable();
      }
    }
    inputEl.value = "";
    box.style.display = "none";
    return;
  }

  var results = filterProducts(val);
  if (results.length === 0) {
    box.innerHTML = '<div class="suggest-empty">Không tìm thấy sản phẩm phù hợp.</div>';
    box.style.display = "block";
    return;
  }
  var html = "";
  results.slice(0, 10).forEach(function(item) {
    var itemStr = encodeURIComponent(JSON.stringify(item));
    html += '<div class="suggest-item" onclick="xb_chonSanPhamFromSuggest(\'' + itemStr + '\')"><div class="sg-title">' + item.tenHang + '</div><div class="sg-desc"><span style="color:#1a73e8; font-weight:700;">Mã hàng: ' + item.maHang + '</span> · Mã vạch: ' + item.maVach + ' · ĐVT: ' + escapeHtml(formatDvtDisplay_(item.maHang, item.maVach, item.dvt)) + '</div></div>';
  });
  box.innerHTML = html;
  box.style.display = "block";
}

function xb_chonSanPhamFromSuggest(itemStr) {
  xb_chonSanPham(JSON.parse(decodeURIComponent(itemStr)));
  var input = document.getElementById("xb-input-scan");
  var box = document.getElementById("xb-suggest-box");
  if (input) { input.value = ""; input.focus(); }
  if (box) box.style.display = "none";
}

function xb_chonSanPham(it) {
  if (!xbInvoiceLocked) return alert("Xác nhận số hóa đơn liên kết trước.");
  if (itemHasVariantGroup_(it)) {
    openVariantPicker_(it, 'xb', 1);
    return;
  }
  xb_chonSanPhamDirect_(it, 1);
}

function xb_chonSanPhamDirect_(it, qty) {
  if (!xbInvoiceLocked) return alert("Xác nhận số hóa đơn liên kết trước.");
  if (!it) return;
  var addQty = Number(qty);
  if (!addQty || addQty <= 0 || isNaN(addQty)) addQty = Number(it.sl) > 0 ? Number(it.sl) : 1;
  var cat = lookupCatalogProductClient_(it && it.maHang, it && it.maVach) || it || {};
  var maHang = it.maHang || cat.maHang || "";
  var maVach = it.maVach || cat.maVach || "";
  var tenHang = it.tenHang || cat.tenHang || "";
  var dvt = resolveDvtClient_(maHang, maVach, (it && it.dvt) || (cat && cat.dvt));
  var existingIndex = xbItems.findIndex(function(x) {
    return x.maHang !== "LỖI MÃ" &&
      String(x.maHang || '').trim().toUpperCase() === String(maHang || '').trim().toUpperCase() &&
      String(x.maVach || '').trim().toUpperCase() === String(maVach || '').trim().toUpperCase();
  });
  if (existingIndex !== -1) {
    xbItems[existingIndex].sl = Number(xbItems[existingIndex].sl) + addQty;
    xbItems[existingIndex].highlight = true;
    xbItems[existingIndex].dvt = dvt || xbItems[existingIndex].dvt;
  } else {
    xbItems.unshift({
      maHang: maHang,
      maVach: maVach,
      tenHang: tenHang,
      dvt: dvt,
      parentSku: it.parentSku || cat.parentSku || '',
      sl: String(addQty),
      highlight: true
    });
  }
  xb_renderTable();
}

function xb_thayDoiSoLuong(index, delta) {
  var currentSl = Number(xbItems[index].sl) || 0;
  var newSl = currentSl + delta;
  if (newSl > 0) { xbItems[index].sl = newSl; xb_renderTable(); }
}

function xb_renderTable() {
  var tbody = document.getElementById("xb-tbody-items");
  var tongEl = document.getElementById("xb-lbl-tong-sl");
  if (!tbody) return;
  tbody.innerHTML = "";
  var tongSl = 0;
  xbItems.forEach(function(it, i) {
    var isErr = (it.maHang === "LỖI MÃ" || isNaN(Number(it.sl)));
    tongSl += (Number(it.sl) || 0);
    var trClass = isErr ? "row-error" : (it.highlight ? "scan-highlight" : "");
    it.highlight = false;
    tbody.insertAdjacentHTML("beforeend",
      '<tr class="' + trClass + '"><td>' + (xbItems.length - i) + '</td>' +
      '<td>' + formatPackingCodeHtml_(it) + '</td>' +
      '<td style="font-weight:500;">' + escapeHtml(formatPackingProductName_(it)) + '</td><td>' + escapeHtml(formatDvtDisplay_(it.maHang, it.maVach, it.dvt)) + '</td>' +
      '<td><div class="qty-control"><button class="qty-btn" onclick="xb_thayDoiSoLuong(' + i + ', -1)">-</button>' +
      '<input type="number" class="qty-input" value="' + it.sl + '" onchange="xbItems[' + i + '].sl=this.value; xb_renderTable();">' +
      '<button class="qty-btn" onclick="xb_thayDoiSoLuong(' + i + ', 1)">+</button></div></td>' +
      '<td style="text-align:center;"><button style="color:#d93025; border:none; background:none; font-weight:bold; cursor:pointer; font-size:18px;" onclick="xbItems.splice(' + i + ',1); xb_renderTable();">×</button></td></tr>');
  });
  if (tongEl) tongEl.innerText = tongSl;
  if (xbItems.length === 0) {
    tbody.insertAdjacentHTML("beforeend",
      '<tr><td colspan="6" style="text-align:center; color:#64748b; padding:24px;">' +
      (xbInvoiceLocked ? "Đã khóa HĐ — hãy tìm kiếm và thêm sản phẩm bán kèm." : "Nhập số hóa đơn liên kết rồi bấm Xác nhận HĐ để bắt đầu.") +
      "</td></tr>");
  }
}

function xb_submit() {
  if (!xbInvoiceLocked) return alert("Vui lòng xác nhận số hóa đơn liên kết trước.");
  if (!xbItems.length) return alert("Chưa có hàng!");
  if (!catalogLoadState.ready) return alert("Danh mục hàng đang tải. Vui lòng đợi rồi thử lại.");
  if (!sessionUser || !sessionUser.user) return alert("Vui lòng đăng nhập.");
  var cnEl = document.getElementById("xb-chi-nhanh");
  var chiNhanh = cnEl ? cnEl.value : "";
  if (!chiNhanh) return alert("Thiếu chi nhánh xuất bán.");
  if (!confirm("Lưu xuất bán cho HĐ " + xbInvoiceLocked + " (" + xbItems.length + " dòng)?")) return;

  // Đồng bộ lại ĐVT từ catalog Data_Excel trước khi lưu
  var itemsToSave = xbItems.map(function(it) {
    var cat = lookupCatalogProductClient_(it.maHang, it.maVach);
    return {
      maHang: (cat && cat.maHang) || it.maHang,
      maVach: (cat && cat.maVach) || it.maVach,
      tenHang: (cat && cat.tenHang) || it.tenHang,
      dvt: resolveDvtClient_(it.maHang, it.maVach, it.dvt),
      sl: it.sl
    };
  });

  showLoad("Đang lưu xuất bán...");
  apiPost("luuXuatBanHang", {
    soHoaDon: xbInvoiceLocked,
    chiNhanh: chiNhanh,
    items: itemsToSave,
    actor: sessionUser.user
  }).then(function(res) {
    hideLoad();
    if (!res || res.success === false) throw new Error((res && (res.error || res.msg)) || "Không lưu được.");
    alert((res.message || "✅ Đã lưu.") + (res.coLoi ? "\n⚠️ Có dòng lỗi — kiểm tra sheet Xuất Bán Hàng." : ""));
    // Mỗi lần lưu xong phải nhập lại số HĐ liên kết (không giữ khóa HĐ)
    xbItems = [];
    xbInvoiceLocked = "";
    xb_renderTable();
    xb_setItemsEnabled_(false);
    var hdEl = document.getElementById("xb-so-hoa-don");
    if (hdEl) {
      hdEl.disabled = false;
      hdEl.value = "";
      hdEl.focus();
    }
    var scan = document.getElementById("xb-input-scan");
    if (scan) scan.placeholder = "Nhập số HĐ trước, rồi mới thêm mã...";
    xb_loadRecent();
  }).catch(function(err) {
    hideLoad();
    alert("Lỗi: " + err.message + "\nDeploy lại code.gs nếu chưa có API luuXuatBanHang.");
  });
}

function xb_loadRecent() {
  var box = document.getElementById("xb-recent-list");
  if (!box || !sessionUser || !sessionUser.user) return;
  box.textContent = "Đang tải...";
  apiGet("layDanhSachXuatBanHang", {
    ngay: "7days",
    userRole: sessionUser.role || "",
    userStore: sessionUser.store || ""
  }, { timeoutMs: 45000 }).then(function(res) {
    var rows = (res && res.data) ? res.data : [];
    if (!rows.length) {
      box.innerHTML = '<div style="color:#94a3b8;">Chưa có phiếu xuất bán trong 7 ngày gần đây.</div>';
      return;
    }
    var html = '<div style="overflow:auto;"><table style="width:100%; border-collapse:collapse; font-size:13px;">' +
      '<thead><tr style="text-align:left; color:#64748b;"><th style="padding:6px 4px;">Mã XB</th><th>Số HĐ</th><th>Chi nhánh</th><th>SL dòng</th><th>Tổng SL</th><th>Người tạo</th></tr></thead><tbody>';
    rows.slice(0, 30).forEach(function(r) {
      html += '<tr style="border-top:1px solid #e2e8f0;">' +
        '<td style="padding:7px 4px;"><b>' + escapeHtml(r.maPhieu) + '</b></td>' +
        '<td>' + escapeHtml(r.soHoaDon) + '</td>' +
        '<td>' + escapeHtml(formatStoreShortLabel_(r.chiNhanh) || r.chiNhanh) + '</td>' +
        '<td>' + (r.itemCount || 0) + '</td>' +
        '<td>' + (r.tongSl || 0) + '</td>' +
        '<td>' + escapeHtml(r.actor || "") + '</td></tr>';
    });
    html += "</tbody></table></div>";
    box.innerHTML = html;
  }).catch(function(err) {
    box.innerHTML = '<div style="color:#b91c1c;">Không tải được danh sách: ' + escapeHtml(err.message) + "</div>";
  });
}
