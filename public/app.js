// API helpers
const GAS_EXEC_URL = 'https://script.google.com/macros/s/AKfycbwhqeAzzNrPTm1cH7KMmmj44btXb2OL835xxaItHByohT11sLDrdgfw7BrVlI5txqXonw/exec';

async function callJsonApi(urls, options) {
  let lastError = null;
  for (const target of urls) {
    try {
      const res = await fetch(target, options);
      const txt = await res.text();
      if (!res.ok) {
        lastError = new Error('HTTP ' + res.status + ': ' + txt);
        continue;
      }
      try { return JSON.parse(txt); } catch(e) { return txt; }
    } catch (err) {
      lastError = err;
    }
  }
  throw lastError || new Error('Không thể kết nối tới máy chủ');
}

async function apiGet(action, params) {
  const proxyUrl = new URL('/api/gas-proxy', location.origin);
  proxyUrl.searchParams.set('action', action);
  if (params) {
    Object.keys(params).forEach(k => {
      if (params[k] !== undefined && params[k] !== null) proxyUrl.searchParams.set(k, params[k]);
    });
  }

  const directUrl = new URL(GAS_EXEC_URL);
  directUrl.searchParams.set('action', action);
  if (params) {
    Object.keys(params).forEach(k => {
      if (params[k] !== undefined && params[k] !== null) directUrl.searchParams.set(k, params[k]);
    });
  }

  return callJsonApi([proxyUrl.toString(), directUrl.toString()], { method: 'GET', headers: { 'Accept': 'application/json' } });
}

async function apiPost(action, payload) {
  const body = { action: action, payload: payload };
  const options = { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) };
  return callJsonApi(['/api/gas-proxy', GAS_EXEC_URL], options);
}

function showLoginError(message) {
  alert(message || 'Đăng nhập thất bại.');
}

// --- App logic (extracted from original webapp) ---
var danhMucGoc = {}; var danhMucArr = []; var arrItems = []; var gStores = [];
var storeMap = {};
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
  showLoad("Đang tải dữ liệu hệ thống...");
  var qlNgay = getEl("ql-ngay"); if (qlNgay) qlNgay.valueAsDate = new Date();
  apiGet('getInitialData').then(function(res) {
    hideLoad();
    if(!res.success) { alert("Lỗi tải data: " + (res.error||res)); return; }
    gStores = res.stores; danhMucGoc = res.danhMuc;
    var seenCatalogKeys = {};
    danhMucArr = Object.values(danhMucGoc).filter(function(item) {
      if (!item) return false;
      var key = String(item.maHang || "").trim().toUpperCase() + '|' + String(item.maVach || "").trim().toUpperCase() + '|' + String(item.tenHang || "").trim().toUpperCase();
      if (!key || seenCatalogKeys[key]) return false;
      seenCatalogKeys[key] = true;
      return true;
    });
    storeMap = res.storeMap || {};

    var htmlStores = ""; gStores.forEach(function(s) { var disp = storeMap[s] || s; htmlStores += '<option value="'+s+'">'+disp+'</option>'; });
    var elX = document.getElementById("select-kho-xuat"); if(elX) elX.innerHTML = htmlStores;
    var elN = document.getElementById("select-kho-nhan"); if(elN) elN.innerHTML = htmlStores;
    var elQ = document.getElementById("ql-kho-nhan"); if(elQ) elQ.innerHTML = '<option value="all">-- Tất cả --</option>' + htmlStores;
    var elA = document.getElementById("adm-store"); if(elA) elA.innerHTML = '<option value="Tất cả">-- Chọn kho quản lý --</option>' + htmlStores;

    if (sessionUser.role === "Admin") { var nav = document.getElementById("nav-tab-admin"); if(nav) nav.style.display = "block"; }
    updateDashboardHero();
    applyQuyenKho();
    loadDashboardSummary();
    openDeepLinkedOrder();
  }).catch(function(err){ hideLoad(); alert('Lỗi: '+err.message); });
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

function loadDashboardSummary() {
  if (!sessionUser || !sessionUser.user) return;
  var grid = document.getElementById('dashboard-summary-grid');
  var recent = document.getElementById('dashboard-recent-orders');
  if (!grid || !recent) return;
  showLoad('Đang tải tổng quan...');
  apiGet('getDashboardSummary', { userRole: sessionUser.role || '', userStore: sessionUser.store || '' }).then(function(res) {
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
      recent.innerHTML = '<div style="padding:12px; background:#f8fafc; border:1px dashed #cbd5e1; border-radius:12px; color:#64748b;">Chưa có đơn hàng nào trong phạm vi của bạn.</div>';
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
    if(khoXuatEl) { khoXuatEl.value = "Kho Địa điểm kinh doanh Q7"; khoXuatEl.setAttribute("disabled", "true"); }
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
    html += '<div class="suggest-item" onclick="chonSanPhamFromSuggest(\'' + itemStr + '\')"><div class="sg-title">' + item.tenHang + '</div><div class="sg-desc"><span style="color:#1a73e8; font-weight:700;">Mã hàng: ' + item.maHang + '</span> · Mã vạch: ' + item.maVach + '</div></div>';
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

function ql_handleAddCodeInput(e) {
  var inputEl = document.getElementById("ql-add-code");
  var box = document.getElementById("ql-suggest-box");
  if (!inputEl || !box) return;
  var val = inputEl.value.trim();
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
    html += '<div class="suggest-item" onclick="ql_pickSuggestedItem(\'' + itemStr + '\')"><div class="sg-title">' + item.tenHang + '</div><div class="sg-desc"><span style="color:#2563eb; font-weight:700;">Mã hàng: ' + item.maHang + '</span> · Mã vạch: ' + item.maVach + '</div></div>';
  });
  box.innerHTML = html;
  box.style.display = "block";
}

function ql_pickSuggestedItem(itemStr) {
  var item = JSON.parse(decodeURIComponent(itemStr));
  var input = document.getElementById("ql-add-code");
  var box = document.getElementById("ql-suggest-box");
  if (input) input.value = item.maHang;
  if (box) box.style.display = "none";
  ql_themMaHang(item);
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
  showLoad("Đang tạo đơn...");
  var lPhieu = document.querySelector('input[name="loaiPhieu"]:checked').value;
  var khoXuat = document.getElementById("select-kho-xuat").value;
  var khoNhan = document.getElementById("select-kho-nhan").value;

  apiPost('luuPhieuTuWebApp', { loaiPhieu: lPhieu, khoXuat: khoXuat, khoNhan: khoNhan, items: arrItems }).then(function(res) {
    hideLoad();
    if(res.coLoi) { alert("⚠️ Có mã lỗi. Sửa trong tab Quản lý!"); arrItems = []; renderTable(); }
    else {
       currentPhieuObj = { soPhieu: res.soPhieu, khoXuat: khoXuat, khoNhan: khoNhan };
       document.getElementById("modal-sophieu").innerText = res.soPhieu; document.getElementById("modal-action").style.display = "flex";
       arrItems = []; renderTable();
       if (document.getElementById("input-scan")) document.getElementById("input-scan").focus();
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
function setQuickDateFilter(value, targetId) {
  var inputEl = document.getElementById(targetId);
  if (!inputEl) return;
  if (!value) { inputEl.value = ''; return; }
  var today = new Date();
  if (value === 'today') {
    var yyyy = today.getFullYear(); var mm = String(today.getMonth()+1).padStart(2,'0'); var dd = String(today.getDate()).padStart(2,'0');
    inputEl.value = yyyy + '-' + mm + '-' + dd;
  } else if (value === 'yesterday') {
    var y = new Date(today); y.setDate(today.getDate() - 1);
    var yy = y.getFullYear(); var ym = String(y.getMonth()+1).padStart(2,'0'); var yd = String(y.getDate()).padStart(2,'0');
    inputEl.value = yy + '-' + ym + '-' + yd;
  } else if (value === '7days') {
    var d7 = new Date(today); d7.setDate(today.getDate() - 6);
    var y7 = d7.getFullYear(); var m7 = String(d7.getMonth()+1).padStart(2,'0'); var day7 = String(d7.getDate()).padStart(2,'0');
    inputEl.value = y7 + '-' + m7 + '-' + day7;
  } else if (value === 'all') {
    inputEl.value = '';
  }
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

function ql_loadPhieu(selectedSoPhieu) {
  var selectEl = document.getElementById("ql-phieu");
  if (!selectEl) return;
  selectEl.innerHTML = '<option value="">⏳ Đang tải...</option>';
  apiGet('layDanhSachPhieuTheoFilter', { khoNhan: document.getElementById("ql-kho-nhan").value, ngay: document.getElementById("ql-ngay").value, userRole: sessionUser.role, userStore: sessionUser.store }).then(function(res) {
    phieuData = res; var countMoi = 0; var countDone = 0; var countCancel = 0;
    var html = '<option value="">-- Chọn Đơn ('+res.length+') --</option>';
    res.forEach(r => {
      if(r.trangThai === "Mới") countMoi++; else if(r.trangThai === "Đã hủy") countCancel++; else countDone++;
      var shortName = storeMap[r.khoNhan] || storeMap[r.khoXuat] || r.khoNhan || r.khoXuat || '';
      html += '<option value="'+r.soPhieu+'">'+r.soPhieu+' ('+shortName+') ['+r.trangThai+']</option>';
    });
    selectEl.innerHTML = html; document.getElementById("ql-view-phieu").style.display = "none";
    document.getElementById("ql-stats").innerHTML = '<div class="stat-box" style="color:#d93025;">🔔 MỚI: '+countMoi+'</div> | <div class="stat-box" style="color:#137333;">✅ ĐÃ XỬ LÝ: '+countDone+'</div> | <div class="stat-box" style="color:#8b5a2b;">🚫 HỦY: '+countCancel+'</div>';
    if (selectedSoPhieu) {
      var matched = res.find(function(item) { return item.soPhieu === selectedSoPhieu; });
      if (matched) {
        selectEl.value = selectedSoPhieu;
        ql_hienThiChiTiet(matched);
      }
    }
  }).catch(function(err){ alert('Lỗi: '+err.message); });
}

function ql_onSelectPhieu() {
  var val = document.getElementById("ql-phieu").value;
  if(!val) { document.getElementById("ql-view-phieu").style.display = "none"; return; }
  ql_hienThiChiTiet(phieuData.find(x => x.soPhieu === val));
}

function confirm_loadPhieu(selectedSoPhieu) {
  var selectEl = document.getElementById("confirm-phieu");
  if (!selectEl) return;
  selectEl.innerHTML = '<option value="">⏳ Đang tải...</option>';
  apiGet('layDanhSachPhieuTheoFilter', { khoNhan: document.getElementById("ql-kho-nhan").value, ngay: document.getElementById("confirm-ngay").value, userRole: sessionUser.role, userStore: sessionUser.store }).then(function(res) {
    var html = '<option value="">-- Chọn Phiếu --</option>';
    res.forEach(function(r) {
      var shortName = storeMap[r.khoNhan] || storeMap[r.khoXuat] || r.khoNhan || r.khoXuat || '';
      html += '<option value="'+r.soPhieu+'">'+r.soPhieu+' ('+shortName+') ['+r.trangThai+']</option>';
    });
    selectEl.innerHTML = html;
    if (selectedSoPhieu) {
      selectEl.value = selectedSoPhieu;
      confirm_onSelectPhieu();
    }
  }).catch(function(err){ alert('Lỗi: '+err.message); });
}

function confirm_onSelectPhieu() {
  var val = document.getElementById("confirm-phieu").value;
  if (!val) { document.getElementById("confirm-view").style.display = "none"; return; }
  showLoad("Đang tải dữ liệu nhận hàng...");
  apiGet('getChiTietPhieu', { soPhieu: val, storeName: sessionUser.store }).then(function(rows) {
    hideLoad();
    currentConfirmPhieuObj = { soPhieu: val };
    var viewEl = document.getElementById("confirm-view");
    document.getElementById("confirm-lbl-sophieu").innerText = val;
    document.getElementById("confirm-meta").innerText = "Đã tải " + rows.length + " dòng để xác nhận.";
    var tbody = document.getElementById("confirm-tbody");
    tbody.innerHTML = "";
    rows.filter(function(r){ return r.trangThai !== "Đã hủy dòng" && r.trangThai !== "Đã hủy đơn"; }).forEach(function(r, idx) {
      var packedQty = (r.slThucTe !== undefined && r.slThucTe !== null && r.slThucTe !== "") ? Number(r.slThucTe) : Number(r.slGoc || 0);
      var inputValue = packedQty;
      tbody.insertAdjacentHTML('beforeend', '<tr><td>' + (idx + 1) + '</td><td><b>' + (r.maVach || '') + '</b><br><small style="color:gray;">' + (r.maHang || '') + '</small><br><small>' + (r.tenHang || '') + '</small></td><td style="font-weight:700;">' + packedQty + '</td><td><input type="number" class="confirm-qty-input same" data-row="' + r.rowIndex + '" data-packed="' + packedQty + '" data-previous="' + packedQty + '" value="' + inputValue + '" min="0" oninput="confirm_updateInput(this)"></td></tr>');
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
  var canManageOrder = !!sessionUser.user;
  var canReceiveConfirm = !!sessionUser.user && (isAdmin || sessionUser.store === currentPhieuObj.khoNhan || sessionUser.store === "Tất cả" || sessionUser.store === currentPhieuObj.khoXuat);
  showLoad("Tải chi tiết...");
  apiGet('getChiTietPhieu', { soPhieu: currentPhieuObj.soPhieu, storeName: currentPhieuObj.khoXuat }).then(function(rows) {
    hideLoad(); editRows = rows; currentLoadedRows = rows || []; var tb = document.getElementById("ql-tbody"); tb.innerHTML = "";
    var isReadOnlyOrder = rows.some(function(r) { return r.trangThai === "Đã xác nhận nhận hàng"; });
    var canEditRows = !!sessionUser.user && !isReadOnlyOrder && isAdmin && !isPublicView;
    var canAddItems = !!sessionUser.user && !isReadOnlyOrder && !isPublicView;
    var canCancelOrder = !!sessionUser.user && !isReadOnlyOrder && isAdmin && !isPublicView;

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
    }
    document.getElementById("ql-lbl-sophieu").innerText = currentPhieuObj.soPhieu;
    document.getElementById("ql-lbl-khoxuat").innerText = currentPhieuObj.khoXuat + ' (' + (storeMap[currentPhieuObj.khoXuat] || '') + ')';
    document.getElementById("ql-lbl-khonhan").innerText = currentPhieuObj.khoNhan + ' (' + (storeMap[currentPhieuObj.khoNhan] || '') + ')';

    rows.forEach((r, i) => {
      var isCancelled = r.trangThai === "Đã hủy dòng" || r.trangThai === "Đã hủy đơn";
      var isReadOnlyRow = isReadOnlyOrder || isCancelled;
      var rowStyle = isCancelled ? 'background:#fce8e6; color:#777; text-decoration:line-through;' : (r.ghiChu ? 'background:#fff8e1;' : '');
      var latestQty = (r.slThucTe !== undefined && r.slThucTe !== null && r.slThucTe !== "") ? r.slThucTe : r.slGoc;
      var quantityInput = '<input type="number" class="edit-sl-input" data-row="'+r.rowIndex+'" data-new="0" value="'+latestQty+'" '+(isReadOnlyRow || !canEditRows ? 'disabled' : '')+' style="border:2px solid #1a73e8;text-align:center;width:70px;">';
      var cancelButton = canAddItems && !isReadOnlyOrder ? '<td><button type="button" onclick="ql_huyDong('+r.rowIndex+')" '+(isCancelled ? 'disabled' : '')+' style="border:none; background:#d93025; color:white; border-radius:5px; padding:7px 9px; cursor:pointer;">Hủy mã</button></td>' : '<td></td>';
      var displayQty = isPublicView ? (latestQty !== undefined && latestQty !== null && latestQty !== "" ? latestQty : "") : quantityInput;
      tb.insertAdjacentHTML('beforeend', '<tr style="'+rowStyle+'"><td>'+(i+1)+'</td><td><b>Mã vạch: '+r.maVach+'</b><br><small style="color:gray;">Mã hàng hóa: '+(r.maHang||'')+'</small></td><td>'+r.tenHang+(r.ghiChu?'<br><b style="color:red;font-size:11px;">⚠️ '+r.ghiChu+'</b>':'')+(isCancelled || isReadOnlyOrder ? '<br><b style="color:#d93025;font-size:11px;">'+r.trangThai+'</b>' : '')+'</td><td>'+r.stock+'</td><td>'+displayQty+'</td>'+cancelButton+'</tr>');
    });
    var packerNames = rows.map(function(r){ return r.nguoiSoanHang || ""; }).filter(Boolean);
    if (!isPublicView && packerNames.length) document.getElementById("ql-order-meta").innerText = "Người soạn hàng gần nhất: " + packerNames[packerNames.length - 1];
    if (isReadOnlyOrder && !isPublicView) {
      document.getElementById("ql-order-meta").innerText += " | Chế độ chỉ xem sau khi xác nhận nhận hàng";
    }
    document.getElementById("ql-view-phieu").style.display = "block";
  }).catch(function(err){ hideLoad(); alert('Lỗi: '+err.message); });
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
  if (sessionUser.role !== "Admin") return alert("Chỉ quản trị viên được phép sửa đơn.");
  var inputs = document.querySelectorAll(".edit-sl-input");
  var updates = [];
  var newItems = [];
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
        updates.push({ row: parseInt(rowKey), valSl: qtyValue });
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
  var item = itemOverride || danhMucGoc[code] || filterProducts(code)[0];
  if (!item) return alert("Không tìm thấy mã hàng hóa hoặc mã vạch.");
  if (!quantity || quantity < 1) return alert("Số lượng phải lớn hơn 0.");
  if (ql_isDuplicateOrderItem(item)) {
    alert("⚠️ Mã này đã tồn tại trong đơn hiện tại. Không thể thêm dòng trùng.");
    return;
  }
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
      dvt: item.dvt,
      slGoc: latestQty,
      slThucTe: latestQty,
      stock: 0,
      ghiChu: "",
      trangThai: "Chưa lưu",
      nguoiSoanHang: ""
    });
    tb.insertAdjacentHTML('beforeend', '<tr><td>' + (editRows.length) + '</td><td><b>Mã vạch: ' + (item.maVach || '') + '</b><br><small style="color:gray;">Mã hàng hóa: ' + (item.maHang || '') + '</small></td><td>' + (item.tenHang || '') + '</td><td>0</td><td><input type="number" class="edit-sl-input" data-row="' + tempKey + '" data-new="1" value="' + latestQty + '" style="border:2px solid #1a73e8;text-align:center;width:70px;"></td><td><button type="button" onclick="ql_huyDong(' + tempKey + ')" style="border:none; background:#d93025; color:white; border-radius:5px; padding:7px 9px; cursor:pointer;">Hủy mã</button></td></tr>');
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
  document.getElementById("sh-phieu").innerHTML = '<option value="">⏳ Đang tải...</option>';
  apiGet('getDonHangTheoNgay', { ngay: document.getElementById("sh-ngay").value, userRole: sessionUser.role, userStore: sessionUser.store, viewMode: 'packing' }).then(function(res) {
    var countMoi = 0; var countDone = 0;
    var html = '<option value="">-- Chọn đơn ('+res.length+') --</option>';
    res.forEach(r => {
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
      html += '<div class="item-card"><b>'+it.tenHang+'</b><br><small><b>Mã vạch: '+it.maVach+'</b> | Mã hàng hóa: '+(it.maHang||'')+' | ĐVT: '+it.dvt+'</small><div class="action-row"><div>SL Yêu Cầu: <b>'+it.slGoc+'</b><br>Thực tế: <input type="number" class="sl-thuc-te" data-row="'+it.rowIndex+'" value="'+it.slThucTe+'"></div><div style="text-align:right;"><label class="btn-camera" for="c-'+j+'">📷 Ảnh</label><input type="file" id="c-'+j+'" accept="image/*" capture="environment" style="display:none;" data-row="'+it.rowIndex+'" data-j="'+j+'" onchange="nenAnh(this)"><img id="p-'+j+'" src="'+(it.anhXacNhan||'')+'" style="display:'+(it.anhXacNhan?'block':'none')+'; width:50px; height:50px; margin-top:5px;"><small id="st-'+j+'"></small></div></div></div>';
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
  apiPost('luuSoSoanHangVaAnh', { updates: updates, images: pendingImages, actor: sessionUser ? sessionUser.user : '' }).then(function(res) { hideLoad(); alert(res); pendingImages = {}; sh_taiDanhSachDon(); }).catch(function(err){ hideLoad(); alert('Lỗi: '+err.message); });
}

function sh_taoBangSoanNgayMai() {
  if (!sessionUser || !sessionUser.user) {
    alert("Vui lòng đăng nhập trước khi tạo bảng soạn.");
    return;
  }
  var createDateEl = document.getElementById("sh-create-date");
  var ngay = createDateEl && createDateEl.value ? createDateEl.value : "";
  showLoad("Đang tạo bảng tổng hợp soạn hàng ngày mai...");
  apiPost('taoBangSoanHangNgayMai', { ngay: ngay, actor: sessionUser.user }).then(function(res) {
    hideLoad();
    if (!res || !res.success) {
      alert("❌ Tạo bảng thất bại: " + ((res && (res.msg || res.error)) || "Không rõ lỗi"));
      return;
    }
    var msg = "✅ Đã tạo tab: " + (res.sheetName || "SoanNgayMai") + "\n" +
      "- Tổng đơn: " + (res.totalOrders || 0) + "\n" +
      "- Tổng mã: " + (res.totalItems || 0) + "\n" +
      "- Mã thiếu: " + (res.missingItems || 0);
    alert(msg);
    if (res.url) {
      window.open(res.url, '_blank', 'noopener,noreferrer');
    }
  }).catch(function(err) {
    hideLoad();
    alert('Lỗi: ' + err.message);
  });
}

// ================= ADMIN: QUẢN LÝ TÀI KHOẢN =================
function checkAdminRole() { var r = document.getElementById("adm-role").value; if(r === "Admin") document.getElementById("adm-store").value = "Tất cả"; }
function loadDSUser() {
  showLoad("Đang tải...");
  apiGet('getDanhSachTaiKhoan').then(function(users) {
    hideLoad(); var tb = document.getElementById("adm-table-users"); tb.innerHTML = "";
    users.forEach(u => tb.insertAdjacentHTML('beforeend', '<tr><td><b>'+u.user+'</b></td><td>'+u.role+'</td><td>'+u.store+'</td></tr>'));
  }).catch(function(err){ hideLoad(); alert('Lỗi: '+err.message); });
}
function taoTaiKhoan() {
  var payload = { user: document.getElementById("adm-user").value.trim(), pass: document.getElementById("adm-pass").value.trim(), role: document.getElementById("adm-role").value, store: document.getElementById("adm-store").value, actor: sessionUser.user };
  if(!payload.user || !payload.pass || !payload.store) return alert("Vui lòng điền đủ thông tin!");
  showLoad("Đang tạo...");
  apiPost('taoTaiKhoanMoi', payload).then(function(res) {
    hideLoad();
    if(res.success) { alert("Tạo thành công!"); document.getElementById("adm-user").value=""; document.getElementById("adm-pass").value=""; loadDSUser(); }
    else alert(res.msg);
  }).catch(function(err){ hideLoad(); alert('Lỗi: '+err.message); });
}

function importDanhMucTonKho() {
  if (!sessionUser || sessionUser.role !== "Admin") return alert("Chỉ Admin mới được phép nhập khẩu dữ liệu.");
  var sourceName = (document.getElementById("imp-sheet-name").value || "").trim();
  if (!sourceName) return alert("Vui lòng nhập tên sheet nguồn.");
  showLoad("Đang nhập khẩu dữ liệu...");
  apiPost('nhapKhauCapNhatThongTin', { sourceSheet: sourceName, actor: sessionUser.user }).then(function(res) {
    hideLoad();
    if (!res || !res.success) {
      alert("❌ Nhập khẩu thất bại: " + ((res && (res.error || res.msg)) || "Không rõ lỗi"));
      return;
    }
    var msg = "✅ Nhập khẩu thành công!\n" +
      "- Danh mục cập nhật: " + (res.catalogUpdated || 0) + " dòng\n" +
      "- Tồn kho cập nhật: " + (res.stockUpdated || 0) + " dòng";
    if (res.warnings && res.warnings.length) msg += "\n\n⚠️ Ghi chú:\n- " + res.warnings.join("\n- ");
    alert(msg);
    initSystemData();
  }).catch(function(err) {
    hideLoad();
    alert('Lỗi: ' + err.message);
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
