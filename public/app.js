// API helpers
async function apiGet(action, params) {
  const url = new URL('/api/gas-proxy', location.origin);
  url.searchParams.set('action', action);
  if (params) {
    Object.keys(params).forEach(k => {
      if (params[k] !== undefined && params[k] !== null) url.searchParams.set(k, params[k]);
    });
  }
  const res = await fetch(url.toString(), { method: 'GET', headers: { 'Accept': 'application/json' } });
  const txt = await res.text();
  try { return JSON.parse(txt); } catch(e) { return txt; }
}

async function apiPost(action, payload) {
  const body = { action: action, payload: payload };
  const res = await fetch('/api/gas-proxy', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
  const txt = await res.text();
  try { return JSON.parse(txt); } catch(e) { return txt; }
}

// --- App logic (extracted from original webapp) ---
var danhMucGoc = {}; var danhMucArr = []; var arrItems = []; var gStores = [];
var storeMap = {};
var phieuData = []; var editRows = []; var currentPhieuObj = null;
var sessionUser = { user: "", role: "", store: "" };
var deepLinkOrder = new URLSearchParams(location.search).get("soPhieu");
var deepLinkTab = new URLSearchParams(location.search).get("tab");

window.onload = function() {
  document.getElementById("loading-overlay").style.display = "none";
  var pass = document.getElementById("lg-pass"); if(pass) pass.addEventListener("keypress", function(e){ if(e.key==="Enter") doLogin(); });
};

// ================= ĐĂNG NHẬP =================
function doLogin() {
  var u = document.getElementById("lg-user").value.trim();
  var p = document.getElementById("lg-pass").value.trim();
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
      document.getElementById("lbl-username").innerText = sessionUser.user + " (" + sessionUser.role + ")";
      document.getElementById("login-screen").style.display = "none";
      document.getElementById("main-container").style.display = "block";
      initSystemData();
    } else {
      alert("❌ " + (res.msg || res.error || "Không thể đăng nhập."));
    }
  }).catch(function(err){ hideLoad(); alert('Lỗi kết nối: '+err.message); });
}

function initSystemData() {
  showLoad("Đang tải dữ liệu hệ thống...");
  document.getElementById("ql-ngay").valueAsDate = new Date();
  apiGet('getInitialData').then(function(res) {
    hideLoad();
    if(!res.success) { alert("Lỗi tải data: " + (res.error||res)); return; }
    gStores = res.stores; danhMucGoc = res.danhMuc; danhMucArr = Object.values(danhMucGoc);
    storeMap = res.storeMap || {};

    var htmlStores = ""; gStores.forEach(function(s) { var disp = storeMap[s] || s; htmlStores += '<option value="'+s+'">'+disp+'</option>'; });
    var elX = document.getElementById("select-kho-xuat"); if(elX) elX.innerHTML = htmlStores;
    var elN = document.getElementById("select-kho-nhan"); if(elN) elN.innerHTML = htmlStores;
    var elQ = document.getElementById("ql-kho-nhan"); if(elQ) elQ.innerHTML = '<option value="all">-- Tất cả --</option>' + htmlStores;
    var elA = document.getElementById("adm-store"); if(elA) elA.innerHTML = '<option value="Tất cả">-- Chọn kho quản lý --</option>' + htmlStores;

    if (sessionUser.role === "Admin") { var nav = document.getElementById("nav-tab-admin"); if(nav) nav.style.display = "block"; }
    applyQuyenKho();
    openDeepLinkedOrder();
  }).catch(function(err){ hideLoad(); alert('Lỗi: '+err.message); });
}

function activateTab(tabId) {
  document.querySelectorAll('.nav-tab').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
  var nav = Array.from(document.querySelectorAll('.nav-tab')).find(function(el) { return el.getAttribute('onclick').indexOf("'" + tabId + "'") !== -1; });
  if (nav) nav.classList.add('active');
  document.getElementById(tabId).classList.add('active');
}

function switchTab(tabId) {
  activateTab(tabId);
  if(tabId === 'tab-quan-ly') ql_loadPhieu();
  if(tabId === 'tab-soan-hang') sh_taiDanhSachDon();
  if(tabId === 'tab-admin') loadDSUser();
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
function handleSearchInput(e) {
  var inputEl = document.getElementById("input-scan");
  var val = inputEl.value.trim();
  var box = document.getElementById("suggest-box");
  box.style.width = (inputEl.offsetWidth) + "px"; box.style.left = (inputEl.offsetLeft) + "px"; box.style.top = (inputEl.offsetTop + inputEl.offsetHeight) + "px";

  if (val.length < 2) { box.style.display = "none"; return; }
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
  if (results.length === 0) { box.innerHTML = '<div style="padding:10px; color:red; text-align:center;">Không tìm thấy!</div>'; box.style.display = "block"; return; }

  var html = "";
  results.slice(0, 10).forEach(function(item) {
    var itemStr = encodeURIComponent(JSON.stringify(item));
    html += '<div class="suggest-item" onclick="chonSanPhamFromSuggest(\'' + itemStr + '\')"><div class="sg-title">' + item.tenHang + '</div><div class="sg-desc">Mã hàng hóa: <b>' + item.maHang + '</b> | Mã vạch: <b>' + item.maVach + '</b></div></div>';
  });
  box.innerHTML = html; box.style.display = "block";
}

function filterProducts(kw) {
  return danhMucArr.filter(function(it) {
    var maH = it.maHang ? it.maHang.toUpperCase() : ""; var maV = it.maVach ? it.maVach.toUpperCase() : ""; var tenH = it.tenHang ? it.tenHang.toUpperCase() : "";
    var s6Vach = maV.length >= 6 ? maV.substring(maV.length - 6) : maV;
    return maH.indexOf(kw) !== -1 || maV.indexOf(kw) !== -1 || tenH.indexOf(kw) !== -1 || s6Vach.indexOf(kw) !== -1;
  });
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

document.addEventListener("click", function(event) { var box = document.getElementById("suggest-box"); var input = document.getElementById("input-scan"); if (event.target !== box && event.target !== input && !box.contains(event.target)) box.style.display = "none"; });

function thayDoiSoLuong(index, delta) { var currentSl = Number(arrItems[index].sl) || 0; var newSl = currentSl + delta; if (newSl > 0) { arrItems[index].sl = newSl; renderTable(); } }

function renderTable() {
  var tbody = document.getElementById("tbody-items"); tbody.innerHTML = ""; var tongSl = 0;
  arrItems.forEach((it, i) => {
    var isErr = (it.maHang === "LỖI MÃ" || isNaN(Number(it.sl))); tongSl += (Number(it.sl) || 0);
    var trClass = isErr ? 'row-error' : (it.highlight ? 'scan-highlight' : ''); it.highlight = false;
    tbody.insertAdjacentHTML('beforeend', '<tr class="' + trClass + '"><td>' + (arrItems.length - i) + '</td><td><b>Mã vạch: ' + it.maVach + '</b><br><small style="color:gray;">Mã hàng hóa: ' + it.maHang + '</small></td><td style="font-weight:500;">' + it.tenHang + '</td><td>' + it.dvt + '</td><td><div class="qty-control"><button class="qty-btn" onclick="thayDoiSoLuong(' + i + ', -1)">-</button><input type="number" class="qty-input" value="' + it.sl + '" onchange="arrItems[' + i + '].sl=this.value; renderTable();"><button class="qty-btn" onclick="thayDoiSoLuong(' + i + ', 1)">+</button></div></td><td style="text-align:center;"><button style="color:#d93025; border:none; background:none; font-weight:bold; cursor:pointer; font-size:18px;" onclick="arrItems.splice(' + i + ',1); renderTable();">×</button></td></tr>');
  });
  document.getElementById("lbl-tong-sl").innerText = tongSl;
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
    }
  }).catch(function(err){ hideLoad(); alert('Lỗi: '+err.message); });
}

function actionPrintNew() { executePrintWeb(currentPhieuObj.soPhieu, currentPhieuObj.khoXuat, currentPhieuObj.khoNhan, arrItems); }
function actionExportNew() { executeExportExcel(currentPhieuObj.soPhieu, currentPhieuObj.khoXuat, currentPhieuObj.khoNhan, arrItems); }
function actionCloseModal() { document.getElementById("modal-action").style.display = "none"; arrItems = []; renderTable(); }

// ================= IN WEB (IFRAME ẨN) & XUẤT EXCEL =================
function executePrintWeb(soPhieu, khoXuat, khoNhan, itemsArray) {
  var typeTitle = soPhieu.indexOf("DH") !== -1 ? "ĐƠN ĐẶT HÀNG" : "LỆNH ĐIỀU CHUYỂN";
  var styleStr = 'body{font-family: Arial, sans-serif; padding:20px;} h2{text-align:center;} table{width:100%; border-collapse:collapse; margin-top:20px;} th,td{border:1px solid #000; padding:8px; text-align:left;} th{background:#f0f0f0;} @media print { @page { margin: 1cm; } }';

  var htmlStr = '<h2>' + typeTitle + '</h2><p><b>Số:</b> ' + soPhieu + '<br><b>Kho xuất:</b> ' + khoXuat + '<br><b>Kho nhận:</b> ' + khoNhan + '</p>';
  htmlStr += '<table><thead><tr><th>STT</th><th>Mã</th><th>Tên hàng</th><th>ĐVT</th><th>Số lượng</th></tr></thead><tbody>';
  var stt = 1; itemsArray.forEach(it => { if(Number(it.sl) > 0) { htmlStr += '<tr><td>'+(stt++)+'</td><td><b>Mã vạch: '+ (it.maVach || '') +'</b><br><small style="color:gray;">Mã hàng hóa: '+ (it.maHang || '') +'</small></td><td>'+it.tenHang+'</td><td>'+it.dvt+'</td><td style="text-align:center;"><b>'+it.sl+'</b></td></tr>'; }});
  htmlStr += '</tbody></table><div style="display:flex; justify-content:space-between; margin-top:50px; text-align:center;"><div><b>Người lập phiếu</b><br><br><br>Ký ghi rõ họ tên</div><div><b>Người nhận</b><br><br><br>Ký ghi rõ họ tên</div></div>';

  var iframe = document.createElement('iframe');
  iframe.style.display = 'none';
  document.body.appendChild(iframe);

  var iframeDoc = iframe.contentWindow.document;
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
function ql_loadPhieu() {
  document.getElementById("ql-phieu").innerHTML = '<option value="">⏳ Đang tải...</option>';
  apiGet('layDanhSachPhieuTheoFilter', { khoNhan: document.getElementById("ql-kho-nhan").value, ngay: document.getElementById("ql-ngay").value, userRole: sessionUser.role, userStore: sessionUser.store }).then(function(res) {
    phieuData = res; var countMoi = 0; var countDone = 0;
    var html = '<option value="">-- Chọn Đơn ('+res.length+') --</option>';
    res.forEach(r => {
      if(r.trangThai === "Mới") countMoi++; else countDone++;
      var shortName = storeMap[r.khoNhan] || storeMap[r.khoXuat] || r.khoNhan || r.khoXuat || '';
      html += '<option value="'+r.soPhieu+'">'+r.soPhieu+' ('+shortName+') ['+r.trangThai+']</option>';
    });
    document.getElementById("ql-phieu").innerHTML = html; document.getElementById("ql-view-phieu").style.display = "none";
    document.getElementById("ql-stats").innerHTML = '<div class="stat-box" style="color:#d93025;">🔔 MỚI: '+countMoi+'</div> | <div class="stat-box" style="color:#137333;">✅ ĐÃ XỬ LÝ: '+countDone+'</div>';
  }).catch(function(err){ alert('Lỗi: '+err.message); });
}

function ql_onSelectPhieu() {
  var val = document.getElementById("ql-phieu").value;
  if(!val) { document.getElementById("ql-view-phieu").style.display = "none"; return; }
  ql_hienThiChiTiet(phieuData.find(x => x.soPhieu === val));
}

function openDeepLinkedOrder() {
  if (!deepLinkOrder || (deepLinkTab && deepLinkTab !== "quan-ly")) return;
  showLoad("Đang mở đơn hàng...");
  apiGet('getThongTinPhieu', { soPhieu: deepLinkOrder }).then(function(phieu) {
    if (!phieu || !phieu.soPhieu) {
      hideLoad();
      alert("Không tìm thấy đơn hàng: " + deepLinkOrder);
      return;
    }
    activateTab('tab-quan-ly');
    ql_hienThiChiTiet(phieu);
  }).catch(function(err) { hideLoad(); alert('Lỗi mở đơn hàng: ' + err.message); });
}

function ql_hienThiChiTiet(phieu) {
  if (!phieu) return;
  currentPhieuObj = phieu;
  var isAdmin = sessionUser.role === "Admin";
  document.getElementById("ql-admin-actions").style.display = isAdmin ? "flex" : "none";
  document.getElementById("ql-admin-column").style.display = isAdmin ? "table-cell" : "none";
  document.getElementById("ql-btn-cancel-order").style.display = isAdmin ? "block" : "none";
  document.getElementById("ql-btn-save").style.display = isAdmin ? "block" : "none";
  document.getElementById("ql-lbl-sophieu").innerText = currentPhieuObj.soPhieu;
  document.getElementById("ql-lbl-khoxuat").innerText = currentPhieuObj.khoXuat + ' (' + (storeMap[currentPhieuObj.khoXuat] || '') + ')';
  document.getElementById("ql-lbl-khonhan").innerText = currentPhieuObj.khoNhan + ' (' + (storeMap[currentPhieuObj.khoNhan] || '') + ')';
  showLoad("Tải chi tiết...");
  apiGet('getChiTietPhieu', { soPhieu: currentPhieuObj.soPhieu, storeName: currentPhieuObj.khoXuat }).then(function(rows) {
    hideLoad(); editRows = rows; var tb = document.getElementById("ql-tbody"); tb.innerHTML = "";
    rows.forEach((r, i) => {
      var isCancelled = r.trangThai === "Đã hủy dòng" || r.trangThai === "Đã hủy đơn";
      var rowStyle = isCancelled ? 'background:#fce8e6; color:#777; text-decoration:line-through;' : (r.ghiChu ? 'background:#fff8e1;' : '');
      var quantityInput = '<input type="number" class="edit-sl-input" data-row="'+r.rowIndex+'" value="'+r.slGoc+'" '+(isCancelled || sessionUser.role !== "Admin" ? 'disabled' : '')+' style="border:2px solid #1a73e8;text-align:center;width:70px;">';
      var cancelButton = sessionUser.role === "Admin" ? '<td><button type="button" onclick="ql_huyDong('+r.rowIndex+')" '+(isCancelled ? 'disabled' : '')+' style="border:none; background:#d93025; color:white; border-radius:5px; padding:7px 9px; cursor:pointer;">Hủy mã</button></td>' : '';
      tb.insertAdjacentHTML('beforeend', '<tr style="'+rowStyle+'"><td>'+(i+1)+'</td><td><b>Mã vạch: '+r.maVach+'</b><br><small style="color:gray;">Mã hàng hóa: '+(r.maHang||'')+'</small></td><td>'+r.tenHang+(r.ghiChu?'<br><b style="color:red;font-size:11px;">⚠️ '+r.ghiChu+'</b>':'')+(isCancelled?'<br><b style="color:#d93025;font-size:11px;">'+r.trangThai+'</b>':'')+'</td><td>'+r.stock+'</td><td>'+quantityInput+'</td>'+cancelButton+'</tr>');
    });
    document.getElementById("ql-view-phieu").style.display = "block";
  }).catch(function(err){ hideLoad(); alert('Lỗi: '+err.message); });
}

function ql_luuSua() {
  if (sessionUser.role !== "Admin") return alert("Chỉ quản trị viên được phép sửa đơn.");
  showLoad("Đang lưu chỉnh sửa...");
  var inputs = document.querySelectorAll(".edit-sl-input"), updates = [];
  inputs.forEach(ip => { if (!ip.disabled) updates.push({ row: parseInt(ip.getAttribute("data-row")), valSl: ip.value }); });
  apiPost('luuChinhSuaPhieu', { updates: updates, actor: sessionUser.user }).then(function(res) { hideLoad(); if (!res.success) throw new Error(res.error || "Không thể lưu thay đổi."); alert("✅ Đã lưu chỉnh sửa thành công!"); ql_hienThiChiTiet(currentPhieuObj); }).catch(function(err){ hideLoad(); alert('Lỗi: '+err.message); });
}

function ql_themMaHang() {
  if (sessionUser.role !== "Admin") return;
  var code = document.getElementById("ql-add-code").value.trim().toUpperCase();
  var quantity = Number(document.getElementById("ql-add-qty").value);
  var item = danhMucGoc[code] || filterProducts(code)[0];
  if (!item) return alert("Không tìm thấy mã hàng hóa hoặc mã vạch.");
  if (!quantity || quantity < 1) return alert("Số lượng phải lớn hơn 0.");
  showLoad("Đang thêm mã vào đơn...");
  apiPost('themChiTietPhieu', { soPhieu: currentPhieuObj.soPhieu, item: { maHang: item.maHang, maVach: item.maVach, tenHang: item.tenHang, dvt: item.dvt, sl: quantity }, actor: sessionUser.user }).then(function(res) {
    hideLoad();
    if (!res.success) throw new Error(res.error || "Không thể thêm mã.");
    document.getElementById("ql-add-code").value = "";
    document.getElementById("ql-add-qty").value = "1";
    ql_hienThiChiTiet(currentPhieuObj);
  }).catch(function(err) { hideLoad(); alert('Lỗi: ' + err.message); });
}

function ql_huyDong(row) {
  if (!confirm("Hủy mã hàng này khỏi đơn? Dữ liệu sẽ được lưu lịch sử.")) return;
  showLoad("Đang hủy mã hàng...");
  apiPost('huyDongChiTietPhieu', { row: row, actor: sessionUser.user }).then(function(res) {
    hideLoad();
    if (!res.success) throw new Error(res.error || "Không thể hủy mã.");
    ql_hienThiChiTiet(currentPhieuObj);
  }).catch(function(err) { hideLoad(); alert('Lỗi: ' + err.message); });
}

function ql_huyPhieu() {
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
  var items = []; document.querySelectorAll(".edit-sl-input").forEach((ip, idx) => { items.push({ maHang: editRows[idx].maHang, maVach: editRows[idx].maVach, tenHang: editRows[idx].tenHang, dvt: editRows[idx].dvt, sl: ip.value }); }); return items;
}
function ql_inWeb_FromEdit() { executePrintWeb(currentPhieuObj.soPhieu, currentPhieuObj.khoXuat, currentPhieuObj.khoNhan, layItemsTuBangSua()); }
function ql_xuatExcel_FromEdit() { executeExportExcel(currentPhieuObj.soPhieu, currentPhieuObj.khoXuat, currentPhieuObj.khoNhan, layItemsTuBangSua()); }

// ================= SOẠN HÀNG MOBILE =================
function sh_taiDanhSachDon() {
  document.getElementById("sh-phieu").innerHTML = '<option value="">⏳ Đang tải...</option>';
  apiGet('getDonHangTheoNgay', { ngay: document.getElementById("sh-ngay").value, userRole: sessionUser.role, userStore: sessionUser.store }).then(function(res) {
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
  apiPost('luuSoSoanHangVaAnh', { updates: updates, images: pendingImages }).then(function(res) { hideLoad(); alert(res); pendingImages = {}; sh_taiDanhSachDon(); }).catch(function(err){ hideLoad(); alert('Lỗi: '+err.message); });
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
  var payload = { user: document.getElementById("adm-user").value.trim(), pass: document.getElementById("adm-pass").value.trim(), role: document.getElementById("adm-role").value, store: document.getElementById("adm-store").value };
  if(!payload.user || !payload.pass || !payload.store) return alert("Vui lòng điền đủ thông tin!");
  showLoad("Đang tạo...");
  apiPost('taoTaiKhoanMoi', payload).then(function(res) {
    hideLoad();
    if(res.success) { alert("Tạo thành công!"); document.getElementById("adm-user").value=""; document.getElementById("adm-pass").value=""; loadDSUser(); }
    else alert(res.msg);
  }).catch(function(err){ hideLoad(); alert('Lỗi: '+err.message); });
}

function showLoad(text) { document.getElementById("loading-text").innerText = text; document.getElementById("loading-overlay").style.display = "flex"; }
function hideLoad() { document.getElementById("loading-overlay").style.display = "none"; }
