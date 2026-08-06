// ============================================================
// packing_timeline.gs — Khung ca soan hang N1 10:00 -> N2 10:00
// ============================================================


/**
 * Cửa sổ ngày tổng hợp / giao N2 (packing day):
 * - startTime = N1 10:00:00
 * - midTime   = N2 08:00:00
 * - endTime   = N2 10:00:00
 * Chính:     >= start && < mid
 * Bổ sung:   >= mid   && < end
 * Tổng hợp:  >= start && < end
 * Đúng mốc 10:00:00 N2 thuộc ngày tổng hợp N2+1 (không vào ca N2).
 */
function getPackingDayWindows_(packingDayDate, opts) {
  opts = opts || {};
  var packingDay = packingDayDate instanceof Date && !isNaN(packingDayDate.getTime())
    ? new Date(packingDayDate.getFullYear(), packingDayDate.getMonth(), packingDayDate.getDate(), 0, 0, 0, 0)
    : (getScriptTodayStart_() || new Date());
  packingDay = new Date(packingDay.getFullYear(), packingDay.getMonth(), packingDay.getDate(), 0, 0, 0, 0);
  var prevDay = new Date(packingDay.getFullYear(), packingDay.getMonth(), packingDay.getDate() - 1, 0, 0, 0, 0);

  var mainStartTime = opts.mainStartTime || "10:00";
  var mainEndTime = opts.mainEndTime || opts.suppStartTime || "08:00";
  var suppEndTime = opts.suppEndTime || "10:00";

  var mainStart = combineDateAndTime_(prevDay, mainStartTime);
  var mainEnd = combineDateAndTime_(packingDay, mainEndTime);
  var suppEnd = combineDateAndTime_(packingDay, suppEndTime);
  if (!mainStart || !mainEnd || !suppEnd) {
    return {
      packingDay: packingDay,
      prevDay: prevDay,
      mainStart: null,
      mainEnd: null,
      suppStart: null,
      suppEnd: null,
      startMs: NaN,
      midMs: NaN,
      endMs: NaN
    };
  }

  var startMs = mainStart.getTime();
  var midMs = mainEnd.getTime();
  var endMs = suppEnd.getTime();
  var tz = Session.getScriptTimeZone() || "Asia/Ho_Chi_Minh";

  return {
    packingDay: packingDay,
    prevDay: prevDay,
    mainStart: mainStart,
    mainEnd: mainEnd,
    suppStart: mainEnd,
    suppEnd: suppEnd,
    startMs: startMs,
    midMs: midMs,
    endMs: endMs,
    packingDayStr: Utilities.formatDate(packingDay, tz, "yyyy-MM-dd"),
    prevDayStr: Utilities.formatDate(prevDay, tz, "yyyy-MM-dd"),
    mainLabel: Utilities.formatDate(mainStart, tz, "dd/MM HH:mm") + " → " + Utilities.formatDate(mainEnd, tz, "dd/MM HH:mm") + " (không gồm " + Utilities.formatDate(mainEnd, tz, "HH:mm") + ")",
    suppLabel: Utilities.formatDate(mainEnd, tz, "dd/MM HH:mm") + " → " + Utilities.formatDate(suppEnd, tz, "dd/MM HH:mm") + " (không gồm " + Utilities.formatDate(suppEnd, tz, "HH:mm") + ")",
    totalLabel: Utilities.formatDate(mainStart, tz, "dd/MM HH:mm") + " → " + Utilities.formatDate(suppEnd, tz, "dd/MM HH:mm") + " (không gồm " + Utilities.formatDate(suppEnd, tz, "HH:mm") + ")"
  };
}


function isInPackingMainWindow_(createdMs, win) {
  if (!win || isNaN(createdMs) || isNaN(win.startMs) || isNaN(win.midMs)) return false;
  return createdMs >= win.startMs && createdMs < win.midMs;
}


function isInPackingSuppWindow_(createdMs, win) {
  if (!win || isNaN(createdMs) || isNaN(win.midMs) || isNaN(win.endMs)) return false;
  return createdMs >= win.midMs && createdMs < win.endMs;
}


function isInPackingDayWindow_(createdMs, win) {
  if (!win || isNaN(createdMs) || isNaN(win.startMs) || isNaN(win.endMs)) return false;
  return createdMs >= win.startMs && createdMs < win.endMs;
}


/** packingMode: main | supp | total */
function normalizePackingMode_(mode, onlyNewItems) {
  var m = String(mode || "").trim().toLowerCase();
  if (m === "main" || m === "chinh" || m === "chính") return "main";
  if (m === "supp" || m === "supplement" || m === "bosung" || m === "bổ sung" || m === "bo sung") return "supp";
  if (m === "total" || m === "tong" || m === "tổng" || m === "tonghop" || m === "tổng hợp") return "total";
  return onlyNewItems ? "supp" : "total";
}


function isInPackingModeWindow_(createdMs, win, packingMode) {
  var mode = normalizePackingMode_(packingMode, false);
  if (mode === "main") return isInPackingMainWindow_(createdMs, win);
  if (mode === "supp") return isInPackingSuppWindow_(createdMs, win);
  return isInPackingDayWindow_(createdMs, win);
}


function parseIncludePackedFlag_(value) {
  // Giữ alias cũ — bảng tổng hợp luôn trả mọi trạng thái trong khung ca.
  if (value === true || value === 1 || value === "1") return true;
  var s = String(value == null ? "" : value).trim().toLowerCase();
  return s === "true" || s === "yes" || s === "y" || s === "on";
}


/** Gom trạng thái đơn từ các dòng — không dùng để loại đơn khỏi danh sách. */
function resolvePackingOrderStatusMeta_(item) {
  item = item || {};
  // tone: info(blue)=mới | warn(yellow)=đã soạn | success(green)=đã nhận | danger(red)=hủy
  if (item.hasConfirmed) {
    return {
      trangThai: "Đã xác nhận nhận hàng",
      status: "Đã xác nhận nhận hàng",
      statusTone: "success",
      defaultChecked: false,
      alreadyPacked: false
    };
  }
  if (item.hasCancelled && !item.hasOpen && !item.hasPacked) {
    return {
      trangThai: "Đã hủy",
      status: "Đã hủy",
      statusTone: "danger",
      defaultChecked: false,
      alreadyPacked: false
    };
  }
  if (item.hasPacked && !item.hasOpen) {
    return {
      trangThai: "Đã soạn hàng",
      status: "Đã soạn hàng",
      statusTone: "warn",
      defaultChecked: false,
      alreadyPacked: true
    };
  }
  if (item.hasPacked && item.hasOpen) {
    return {
      trangThai: "Chờ xử lý",
      status: "Chờ xử lý",
      statusTone: "info",
      defaultChecked: true,
      alreadyPacked: false
    };
  }
  if (item.hasCancelled && item.hasOpen) {
    return {
      trangThai: "Mới",
      status: "Mới",
      statusTone: "info",
      defaultChecked: true,
      alreadyPacked: false
    };
  }
  return {
    trangThai: "Mới",
    status: "Mới",
    statusTone: "info",
    defaultChecked: true,
    alreadyPacked: false
  };
}


function getDanhSachDonSoanHang(ngayYYYYMMDD, userRole, userStore, ngayToYYYYMMDD, packingMode, includePacked) {
  var _dbgT0 = Date.now();
  // ngayTo = ngày tổng hợp/giao N2. Fallback: ngay hoặc hôm nay.
  var packingDay = parseDateInputYYYYMMDD(ngayToYYYYMMDD) || parseDateInputYYYYMMDD(ngayYYYYMMDD) || getScriptTodayStart_() || new Date();
  packingDay = new Date(packingDay.getFullYear(), packingDay.getMonth(), packingDay.getDate(), 0, 0, 0, 0);
  var win = getPackingDayWindows_(packingDay);
  var mode = normalizePackingMode_(packingMode, false);
  // Luôn lấy FULL đơn trong khung ca — không lọc theo trạng thái
  var orders = getEligibleOrdersForSoanHang(packingDay, userRole, userStore, null, packingDay, win, mode, {
    includeAllStatuses: true
  });
  var _dbgMs = Date.now() - _dbgT0;
  var branchMergeHints = buildBranchMergeHints_(orders);
  return {
    success: true,
    date: win.prevDayStr,
    dateTo: win.packingDayStr,
    packingDay: win.packingDayStr,
    packingMode: mode,
    includePacked: true,
    includeAllStatuses: true,
    mainWindow: win.mainLabel,
    suppWindow: win.suppLabel,
    totalWindow: win.totalLabel,
    total: orders.length,
    orders: orders,
    branchMergeHints: branchMergeHints,
    hasBranchMergeWarning: branchMergeHints.length > 0,
    _debugTotalMs: _dbgMs,
    _debugRun: "packing-window-v6-merge-hint"
  };
}


/**
 * Lấy đơn theo khung ca (createdAt gốc).
 * Không loại theo trạng thái — FE tự default-check theo badge.
 */
function getEligibleOrdersForSoanHang(baseDate, userRole, userStore, historyPack, endDate, packingWin, packingMode, opts) {
  opts = opts || {};
  var ss = getSS();
  var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
  if (!historySheet) return [];
  historyPack = historyPack || readHistoryDataPack_(historySheet, 5000);
  var data = historyPack.data;
  if (!data || data.length < 2) return [];

  var packingDay = endDate || baseDate;
  var win = packingWin || getPackingDayWindows_(packingDay);
  var mode = normalizePackingMode_(packingMode, false);
  var map = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row) continue;
    var soPhieu = row[1] ? String(row[1]).trim() : "";
    if (!soPhieu) continue;

    var ngayTao = row[0];
    var createdMs = toHoChiMinhMillis_(ngayTao);
    // Chỉ lọc theo createdAt gốc trong khung ca: N1 10:00 ≤ t < N2 10:00
    if (!isInPackingDayWindow_(createdMs, win)) continue;

    var khoXuat = row[2] ? String(row[2]).trim() : "";
    var khoNhan = row[3] ? String(row[3]).trim() : "";
    if (userRole !== "Admin" && !isSameStoreName(khoXuat, userStore || "")) continue;

    var rowStatus = row[12] ? String(row[12]).trim() : "Mới";
    var slThucTe = row[8];
    var displayStatus = getDisplayOrderStatus(rowStatus, slThucTe, row[15]);
    var isDupCancel = rowStatus === "Hủy (Trùng đơn)";
    var isLineCancel = rowStatus === "Đã hủy dòng" || rowStatus === "Đã hủy đơn" || isDupCancel;

    if (!map[soPhieu]) {
      map[soPhieu] = {
        soPhieu: soPhieu,
        khoXuat: khoXuat,
        khoNhan: khoNhan,
        createdAtMs: createdMs,
        hasPacked: false,
        hasConfirmed: false,
        hasCancelled: false,
        hasOpen: false,
        isDuplicateCancelled: false,
        inMain: false,
        inSupp: false,
        totalQty: 0,
        skuQty: {}
      };
    }

    var entry = map[soPhieu];
    // Theo dõi trạng thái để trả FE — KHÔNG dùng để loại đơn
    if (displayStatus === "Đã xác nhận") entry.hasConfirmed = true;
    else if (displayStatus === "Đã soạn") entry.hasPacked = true;
    else if (displayStatus === "Đã hủy" || displayStatus === "Đã hủy dòng" || isDupCancel) {
      entry.hasCancelled = true;
      if (isDupCancel) entry.isDuplicateCancelled = true;
    } else entry.hasOpen = true;

    // Giữ mốc createdAt sớm nhất (gốc) — không bị lệch khi sửa/soạn sau 10h
    if (!isNaN(createdMs) && (isNaN(entry.createdAtMs) || createdMs < entry.createdAtMs)) {
      entry.createdAtMs = createdMs;
    }
    // Gom SL + MaSP cho phát hiện trùng (bỏ dòng hủy)
    if (!isLineCancel) {
      var slLine = Number(row[7]) || 0;
      entry.totalQty += slLine;
      var mhKey = normalizeOrderCodeText(row[4] || "");
      if (mhKey) entry.skuQty[mhKey] = (entry.skuQty[mhKey] || 0) + slLine;
    }
  }

  var orders = [];
  var storesWithMain = {};
  for (var key in map) {
    var item = map[key];
    // Đơn đã hủy vì trùng — gạt khỏi bảng soạn hàng
    if (item.isDuplicateCancelled && !item.hasOpen && !item.hasPacked && !item.hasConfirmed) continue;
    // Phân ca theo createdAt gốc của đơn
    item.inMain = isInPackingMainWindow_(item.createdAtMs, win);
    item.inSupp = isInPackingSuppWindow_(item.createdAtMs, win);
    // Chỉ lọc theo mode ca (main/supp/total), không lọc trạng thái
    if (mode === "main" && !item.inMain) continue;
    if (mode === "supp" && !item.inSupp) continue;
    var bucket = item.inSupp ? "bổ sung" : "chính";
    var bucketLabel = item.inSupp ? "Đơn Bổ Sung" : "Đơn Chính";
    var statusMeta = resolvePackingOrderStatusMeta_(item);
    var knKey = normalizeStoreName(item.khoNhan || "") || String(item.khoNhan || "").trim();
    if (item.inMain && knKey) storesWithMain[knKey] = true;
    orders.push({
      soPhieu: item.soPhieu,
      khoXuat: item.khoXuat,
      khoNhan: item.khoNhan,
      createdAt: item.createdAtMs,
      // Cùng hàm format với getThongTinPhieu / PDF — luôn là thời gian tạo đơn (mốc sớm nhất)
      thoiGianDat: formatOrderCreatedAtLabel_(item.createdAtMs),
      thoiGian: formatOrderCreatedAtLabel_(item.createdAtMs),
      thoiGianDatMillis: item.createdAtMs,
      packingBucket: bucket,
      packingBucketLabel: bucketLabel,
      groupKey: knKey,
      totalQty: item.totalQty || 0,
      skuSignature: buildOrderSkuSignature_(item.skuQty || {}),
      trangThai: statusMeta.trangThai,
      status: statusMeta.status,
      statusTone: statusMeta.statusTone,
      defaultChecked: !!statusMeta.defaultChecked,
      alreadyPacked: !!statusMeta.alreadyPacked
    });
  }
  attachDuplicateSuspects_(orders);

  // Mode tổng hợp: gắn cờ gộp khi kho nhận đã có đơn chính + đơn bổ sung trong cùng ca
  if (mode === "total") {
    for (var oi = 0; oi < orders.length; oi++) {
      var o = orders[oi];
      var gk = o.groupKey || "";
      o.mergeWithMain = !!(o.packingBucket === "bổ sung" && gk && storesWithMain[gk]);
      o.groupHint = o.mergeWithMain
        ? ("Gộp với đơn chính kho " + (formatStoreDisplayLabel_(o.khoNhan) || o.khoNhan || ""))
        : "";
    }
    // Sắp xếp: theo kho nhận → chính trước bổ sung → giờ tạo
    orders.sort(function(a, b) {
      var ga = String(a.groupKey || a.khoNhan || "");
      var gb = String(b.groupKey || b.khoNhan || "");
      if (ga !== gb) return ga < gb ? -1 : 1;
      if (a.packingBucket !== b.packingBucket) return a.packingBucket === "chính" ? -1 : 1;
      return a.thoiGianDatMillis - b.thoiGianDatMillis;
    });
  } else {
    orders.sort(function(a, b) { return a.thoiGianDatMillis - b.thoiGianDatMillis; });
  }
  return orders;
}


/** Chi nhánh (kho nhận) có ≥ 2 đơn trong danh sách — dùng cho modal cảnh báo gộp */
function buildBranchMergeHints_(orders) {
  var byBranch = {};
  for (var i = 0; i < (orders || []).length; i++) {
    var o = orders[i];
    if (!o) continue;
    var gk = String(o.groupKey || normalizeStoreName(o.khoNhan || "") || o.khoNhan || "").trim();
    if (!gk) continue;
    if (!byBranch[gk]) {
      byBranch[gk] = {
        groupKey: gk,
        khoNhan: o.khoNhan || gk,
        khoNhanLabel: formatStoreDisplayLabel_(o.khoNhan) || o.khoNhan || gk,
        orderCount: 0,
        soPhieuList: []
      };
    }
    byBranch[gk].orderCount++;
    byBranch[gk].soPhieuList.push(o.soPhieu);
  }
  var hints = [];
  for (var k in byBranch) {
    if (!Object.prototype.hasOwnProperty.call(byBranch, k)) continue;
    if (byBranch[k].orderCount >= 2) hints.push(byBranch[k]);
  }
  hints.sort(function(a, b) { return b.orderCount - a.orderCount; });
  return hints;
}


function taoBangSoanHangNgayMai(payload) {
  // Thuật toán nhanh: chỉ đọc dòng của đơn đã chọn + gom mã + ghi sheet.
  // BỎ đọc "TỔNG HỢP TỒN KHO" (nguyên nhân chậm 1–2 phút).
  var _dbgT0 = Date.now();
  var _dbgSteps = [];
  function _dbgMark(step, extra) {
    var entry = { step: step, ms: Date.now() - _dbgT0 };
    if (extra) {
      for (var ek in extra) {
        if (extra.hasOwnProperty(ek)) entry[ek] = extra[ek];
      }
    }
    _dbgSteps.push(entry);
  }
  var onlyNewItems = !!(payload && payload.onlyNewItems);
  var packingMode = normalizePackingMode_(payload && payload.packingMode, onlyNewItems);
  var newAfterTime = payload && payload.newAfterTime ? String(payload.newAfterTime).trim() : "08:00";
  var newBeforeTime = payload && payload.newBeforeTime ? String(payload.newBeforeTime).trim() : "10:00";
  _dbgMark("start", {
    selectedCount: payload && payload.selectedOrders ? payload.selectedOrders.length : 0,
    algo: "ton-q7",
    packingMode: packingMode,
    onlyNewItems: onlyNewItems,
    newAfterTime: newAfterTime,
    newBeforeTime: newBeforeTime
  });

  var ss = getSS();
  var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
  if (!historySheet) throw new Error("Không tìm thấy sheet Lịch Sử Xuất Kho.");

  // Ngày tổng hợp/giao N2 = ngayTo (ưu tiên) hoặc ngay
  var packingDayInput = parseDateInputYYYYMMDD(payload && payload.ngayTo ? payload.ngayTo : "")
    || parseDateInputYYYYMMDD(payload && payload.ngay ? payload.ngay : "")
    || getScriptTodayStart_()
    || new Date();
  packingDayInput = new Date(packingDayInput.getFullYear(), packingDayInput.getMonth(), packingDayInput.getDate(), 0, 0, 0, 0);
  var win = getPackingDayWindows_(packingDayInput, {
    mainEndTime: newAfterTime || "08:00",
    suppEndTime: newBeforeTime || "10:00"
  });
  if (!win.mainStart || !win.mainEnd || !win.suppEnd || isNaN(win.startMs)) {
    return {
      success: false,
      msg: "Giờ chốt không hợp lệ. Dùng 08:00 / 10:00.",
      _debugTimings: _dbgSteps,
      _debugTotalMs: Date.now() - _dbgT0,
      _debugRun: "packing-window-v2"
    };
  }
  var baseDate = win.prevDay;
  var endDate = win.packingDay;
  var packingDay = win.packingDay;
  var baseDateStr = win.packingDayStr;
  var tzPack = Session.getScriptTimeZone() || "Asia/Ho_Chi_Minh";
  var newAfterLabel = Utilities.formatDate(win.suppStart, tzPack, "dd/MM HH:mm");
  var newBeforeLabel = Utilities.formatDate(win.suppEnd, tzPack, "dd/MM HH:mm");
  var mainWindowLabel = win.mainLabel;
  var totalWindowLabel = win.totalLabel;

  var userRole = payload && payload.userRole ? payload.userRole : "";
  var userStore = payload && payload.userStore ? payload.userStore : "";
  var selectedOrdersRaw = payload && payload.selectedOrders && payload.selectedOrders.length ? payload.selectedOrders : [];
  var selectedSet = buildOrderMatchSet_(selectedOrdersRaw);

  var historyPack;
  if (selectedSet._list && selectedSet._list.length) {
    historyPack = readHistoryForSelectedOrders_(historySheet, selectedSet, "", 2500);
  } else {
    historyPack = readHistoryDataPack_(historySheet, 3000);
    // Full đơn trong khung ca — không lọc trạng thái
    var eligibleOrders = getEligibleOrdersForSoanHang(packingDay, userRole, userStore, historyPack, packingDay, win, packingMode, {
      includeAllStatuses: true
    });
    selectedSet = buildOrderMatchSet_(eligibleOrders.map(function(o) { return o && o.soPhieu; }));
  }
  _dbgMark("readHistory", {
    historyRows: historyPack.data ? historyPack.data.length - 1 : 0,
    scannedRows: historyPack.scannedRows || (historyPack.data ? historyPack.data.length - 1 : 0),
    matchedRows: historyPack.matchedRows || 0,
    targeted: !!selectedOrdersRaw.length
  });

  var data = historyPack.data;
  if (!data || data.length < 2) {
    return {
      success: false,
      msg: "Không tìm thấy dòng hàng của các đơn đã chọn trong lịch sử gần đây.",
      _debugTimings: _dbgSteps,
      _debugTotalMs: Date.now() - _dbgT0,
      _debugRun: "fast-v10"
    };
  }
  if (!Object.keys(selectedSet).length) {
    return {
      success: false,
      msg: "Không có đơn trong khung ca để tạo bảng soạn. Chọn đơn trên danh sách hoặc kiểm tra ngày N2 / mode ca.",
      _debugTimings: _dbgSteps,
      _debugTotalMs: Date.now() - _dbgT0,
      _debugRun: "fast-v10"
    };
  }

  var itemMap = {};
  var storeList = [];
  var orderSeen = {};
  var storeSeen = {};
  var ordersByStore = {}; // khoNhan -> { soPhieu: true }

  // ĐVT: ưu tiên lịch sử → Data_Excel (cache) → TON_Q7. Ghi ngược vào lịch sử nếu trống.
  var catalogLookup = null;
  try { catalogLookup = getCatalogLookup(ss); } catch (catErr) { catalogLookup = null; }
  _dbgMark("catalogLookup", {
    ready: !!(catalogLookup && (Object.keys(catalogLookup.byMaVach || {}).length || Object.keys(catalogLookup.byMaHang || {}).length)),
    byMaVach: catalogLookup ? Object.keys(catalogLookup.byMaVach || {}).length : 0,
    byMaHang: catalogLookup ? Object.keys(catalogLookup.byMaHang || {}).length : 0
  });
  var historyDvtBackfill = []; // { sheetRow, dvt }
  var skippedByTime = 0;
  var skippedNotSupplement = 0;
  var includedNewRows = 0;
  var includedMainRows = 0;

  // Pass 1: mốc createdAt gốc sớm nhất của đơn (không loại theo trạng thái)
  var orderMinCreated = {};
  for (var p1 = 1; p1 < data.length; p1++) {
    var rowP1 = data[p1];
    if (!rowP1) continue;
    var soP1 = rowP1[1] ? String(rowP1[1]).trim() : "";
    if (!soP1 || !orderInMatchSet_(soP1, selectedSet)) continue;
    var msP1 = toMillisSafe_(rowP1[0]);
    if (isNaN(msP1)) continue;
    if (orderMinCreated[soP1] === undefined || msP1 < orderMinCreated[soP1]) {
      orderMinCreated[soP1] = msP1;
    }
  }

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row) continue;
    var soPhieu = row[1] ? String(row[1]).trim() : "";
    if (!soPhieu || !orderInMatchSet_(soPhieu, selectedSet)) continue;

    var khoXuat = row[2] ? String(row[2]).trim() : "";
    var khoNhan = row[3] ? String(row[3]).trim() : "";
    var maHang = row[4] ? String(row[4]).trim() : "";
    var maVach = row[5] ? String(row[5]).trim() : "";
    var tenHang = row[6] ? String(row[6]).trim() : "";
    var soLuong = Number(row[7]) || 0;
    var dvtRaw = String(row[9] || "").trim();
    var dvt = dvtRaw;
    if (catalogLookup) {
      dvt = resolveDvtValue(catalogLookup, maHang, maVach, dvtRaw) || dvtRaw;
    }
    var noteText = row[11] != null ? String(row[11]).trim() : "";
    var status = row[12] ? String(row[12]).trim() : "Đang xử lý";

    if (!khoNhan) continue;
    // Không lọc theo trạng thái — chỉ bỏ dòng SL <= 0
    if (soLuong <= 0) continue;

    // Lọc theo createdAt GỐC của đơn (min), không theo giờ sửa/thêm dòng sau 10h
    var createdMs = orderMinCreated[soPhieu];
    if (createdMs === undefined || isNaN(createdMs)) {
      createdMs = toHoChiMinhMillis_(row[0]);
    }
    if (isNaN(createdMs)) {
      skippedByTime++;
      continue;
    }

    // main / supp / total theo mốc tạo gốc đơn
    if (!isInPackingModeWindow_(createdMs, win, packingMode)) {
      skippedByTime++;
      continue;
    }
    if (isInPackingMainWindow_(createdMs, win)) includedMainRows++;
    if (isInPackingSuppWindow_(createdMs, win)) includedNewRows++;

    // Tách theo mã + ĐVT (cùng mã có thể có Thùng / Túi ...)
    var codeKey = (maHang ? maHang.toUpperCase() : "") || (maVach ? maVach.toUpperCase() : "");
    if (!codeKey) continue;
    var itemKey = codeKey + "|" + normalizeDvtKey_(dvt);

    if (!itemMap[itemKey]) {
      itemMap[itemKey] = {
        maHang: maHang,
        maVach: maVach,
        tenHang: tenHang,
        dvt: dvt || "",
        totalQty: 0,
        byStore: {},
        sourceStores: {}
      };
    }
    var item = itemMap[itemKey];
    if (!item.tenHang && tenHang) item.tenHang = tenHang;
    if (!item.maHang && maHang) item.maHang = maHang;
    if (!item.maVach && maVach) item.maVach = maVach;
    if (!item.dvt && dvt) item.dvt = dvt;
    item.totalQty += soLuong;
    item.byStore[khoNhan] = (item.byStore[khoNhan] || 0) + soLuong;
    if (khoXuat) item.sourceStores[khoXuat] = true;
    if (!storeSeen[khoNhan]) {
      storeSeen[khoNhan] = true;
      storeList.push(khoNhan);
    }
    orderSeen[soPhieu] = true;
    if (!ordersByStore[khoNhan]) ordersByStore[khoNhan] = {};
    ordersByStore[khoNhan][soPhieu] = true;

    // Lịch sử trống ĐVT nhưng đã resolve từ catalog → ghi lại để lần sau không cần lookup
    if (dvt && !dvtRaw && historyPack.orders && historyPack.orders[i - 1]) {
      historyDvtBackfill.push({ sheetRow: historyPack.orders[i - 1], dvt: dvt });
    }
  }

  var keys = Object.keys(itemMap);
  _dbgMark("aggregateItems", {
    itemCount: keys.length,
    storeCount: storeList.length,
    orderCount: Object.keys(orderSeen).length,
    onlyNewItems: onlyNewItems,
    skippedByTime: skippedByTime,
    skippedNotSupplement: skippedNotSupplement,
    includedNewRows: includedNewRows,
    includedMainRows: includedMainRows,
    mainWindowLabel: mainWindowLabel,
    newAfterLabel: newAfterLabel,
    newBeforeLabel: newBeforeLabel
  });
  if (!keys.length) {
    var emptyMsg = packingMode === "supp"
      ? ("Không có dòng trong khung BỔ SUNG " + (newAfterLabel || "N2 08:00") + " → " + (newBeforeLabel || "N2 10:00") + " (không gồm mốc 10:00).")
      : (packingMode === "main"
        ? ("Không có dòng trong khung CHÍNH " + (mainWindowLabel || "N1 10:00 → N2 08:00") + ".")
        : ("Không có dòng trong khung TỔNG HỢP " + (totalWindowLabel || "N1 10:00 → N2 10:00") + "."));
    return {
      success: false,
      msg: emptyMsg + "\nĐã bỏ qua " + skippedByTime + " dòng ngoài khung / biên giờ.",
      onlyNewItems: packingMode === "supp",
      packingMode: packingMode,
      newAfterLabel: newAfterLabel,
      newBeforeLabel: newBeforeLabel,
      mainWindowLabel: mainWindowLabel,
      totalWindowLabel: totalWindowLabel,
      packingDay: baseDateStr,
      _debugTimings: _dbgSteps,
      _debugTotalMs: Date.now() - _dbgT0,
      _debugRun: "packing-window-v2",
      _debugInfo: {
        baseDateStr: baseDateStr,
        selectedList: Object.keys(selectedSet),
        scannedRows: historyPack.scannedRows || 0,
        skippedByTime: skippedByTime,
        skippedNotSupplement: skippedNotSupplement,
        includedNewRows: includedNewRows,
        includedMainRows: includedMainRows,
        packingMode: packingMode,
        startMs: win.startMs,
        midMs: win.midMs,
        endMs: win.endMs
      }
    };
  }

  storeList.sort(function(a, b) {
    return formatShortStoreLabel(a).localeCompare(formatShortStoreLabel(b));
  });
  var activeMap = getActiveStoreMap();

  // Tồn kho soạn hàng: TON_Q7 mặc định; overlay TON_VARIANT nếu mã thuộc nhóm biến thể
  var forceStock = !!(payload && payload.forceStock);
  var q7Bundle = readTonKhoQ7Bundle_(ss);
  var q7Map = q7Bundle.map;
  var q7DvtLabels = q7Bundle.dvtLabels || {};
  if ((!q7Map || !Object.keys(q7Map).length) && forceStock) {
    try { rebuildTonKhoQ7Sheet_(ss); } catch (rebuildErr) { Logger.log(rebuildErr); }
    q7Bundle = readTonKhoQ7Bundle_(ss);
    q7Map = q7Bundle.map;
    q7DvtLabels = q7Bundle.dvtLabels || {};
  }
  var stockReady = !!(q7Map && Object.keys(q7Map).length);
  if (!q7Map) q7Map = {};
  var variantMapPack = {};
  try { variantMapPack = readTonVariantMap_(ss) || {}; } catch (eVar) { variantMapPack = {}; }
  var parentByChildPack = {};
  try { parentByChildPack = buildTonVariantParentLookup_(ss) || {}; } catch (ePar) { parentByChildPack = {}; }
  _dbgMark("stockIndex", {
    via: "TON_Q7+TON_VARIANT",
    stockReady: stockReady,
    q7Keys: Object.keys(q7Map).length,
    variantKeys: Object.keys(variantMapPack).length,
    parentLookupKeys: Object.keys(parentByChildPack).length,
    forceStock: forceStock,
    itemCount: keys.length
  });

  var sheetName = packingMode === "supp"
    ? "__TMP_SOAN_BO_SUNG"
    : (packingMode === "main" ? "__TMP_SOAN_CHINH" : "__TMP_SOAN_NGAY_MAI");
  var reportSheet = recreateTempSheetFast_(ss, sheetName);
  var packingDayTitle = Utilities.formatDate(packingDay, tzPack, "dd/MM/yyyy");
  var title = packingMode === "supp"
    ? ("BẢNG BỔ SUNG — N2 " + packingDayTitle + " | ≥08:00 & <10:00 | " + newAfterLabel + " → " + newBeforeLabel)
    : (packingMode === "main"
      ? ("BẢNG ĐƠN CHÍNH — N2 " + packingDayTitle + " | ≥N1 10:00 & <N2 08:00 | " + mainWindowLabel)
      : ("BẢNG TỔNG HỢP CA — N2 " + packingDayTitle + " | ≥N1 10:00 & <N2 10:00 | " + (totalWindowLabel || "")));
  // Hàng 5 = số đơn theo cột kho; hàng 6 = header tên cột (Q4 - địa chỉ / Q4 Mới…)
  var headerRow = 6;
  var headers = ["STT", "Mã hàng (Parent)", "Mã vạch", "Tên hàng / Phân loại", "ĐVT", "Stock Q7", "Tổng đặt"];
  var storeOrderCountRow = ["", "", "", "", "", "", "Số đơn"];
  for (var h = 0; h < storeList.length; h++) {
    headers.push(formatPackingColumnLabel_(storeList[h]));
    var storeOrders = ordersByStore[storeList[h]] || {};
    storeOrderCountRow.push(Object.keys(storeOrders).length);
  }
  headers.push("Cảnh báo");
  headers.push("Kho xuất");
  storeOrderCountRow.push("");
  storeOrderCountRow.push("");

  var rows = [];
  var missingLines = 0;
  var warningCol = 8 + storeList.length;
  var _dbgDvtFromOrder = 0;
  var _dbgDvtCatalog = 0;
  var _dbgDvtInferred = 0;
  var _dbgDvtEmpty = 0;
  var _dbgDvtSample = [];
  for (var k = 0; k < keys.length; k++) {
    var it = itemMap[keys[k]];
    var sourceStores = Object.keys(it.sourceStores);
    var dvtOut = it.dvt || "";
    var dvtSource = dvtOut ? "resolved" : "";
    if (!dvtOut && catalogLookup) {
      dvtOut = resolveDvtValue(catalogLookup, it.maHang, it.maVach, "") || "";
      if (dvtOut) dvtSource = "catalog";
    }
    if (!dvtOut && stockReady) {
      dvtOut = inferDvtLabelFromStockMap_(q7Map, q7DvtLabels, it.maHang, it.maVach) || "";
      if (dvtOut) dvtSource = "ton_q7";
    }
    if (dvtSource === "resolved") _dbgDvtFromOrder++;
    else if (dvtSource === "catalog") _dbgDvtCatalog++;
    else if (dvtSource === "ton_q7") _dbgDvtInferred++;
    else _dbgDvtEmpty++;
    if (_dbgDvtSample.length < 5) {
      _dbgDvtSample.push({ ma: it.maHang || it.maVach, orderDvt: it.dvt || "", out: dvtOut || "", src: dvtSource || "empty" });
    }
    // Mã cột = Parent_SKU (mã trên bao bì/kệ); tên = Cha - Biến thể (Mã con: …)
    var metaPack = resolveVariantDisplayMeta_(catalogLookup, it.maHang, it.maVach, it.tenHang, parentByChildPack);
    var maPack = metaPack.maHangDisplay || it.maHang || "";
    var tenPack = metaPack.tenHangDisplay || formatVariantDisplayName_(it.maHang, it.tenHang);

    // Cột F — tra cứu đa tầng: bare / MH: / maHang / maVach / parent (maPack)
    var stock = 0;
    var stockSource = "";
    if (stockReady) {
      var code = String(maPack || it.maHang || it.maVach || "").replace(/^MH:/i, "").trim();
      var rawStock = q7Map[code]
        || q7Map["MH:" + code]
        || q7Map[it.maHang]
        || q7Map[it.maVach]
        || q7Map[maPack]
        || 0;
      var hitQ7 = (code && (q7Map[code] != null || q7Map["MH:" + code] != null))
        || (it.maHang && q7Map[it.maHang] != null)
        || (it.maVach && q7Map[it.maVach] != null)
        || (maPack && q7Map[maPack] != null);
      if (!hitQ7) {
        var sQ7Pack = resolveStockValueWithFallback_(
          q7Map, it.maHang, it.maVach, dvtOut || it.dvt, maPack || it.parentSku || it.maHangDisplay
        );
        if (sQ7Pack !== null && sQ7Pack !== undefined) {
          rawStock = sQ7Pack;
          hitQ7 = true;
        }
      }
      var safeStock = Math.max(0, Number(rawStock) || 0);
      stock = safeStock;
      if (hitQ7) stockSource = "TON_Q7";
    }
    var sVarPack = resolveStockValueWithFallback_(variantMapPack, it.maHang, it.maVach, dvtOut || it.dvt, maPack || it.parentSku);
    if (sVarPack !== null && sVarPack !== undefined) {
      stock = Math.max(0, Number(sVarPack) || 0);
      stockSource = "TON_VARIANT";
    }
    var rowHasStock = !!stockSource;
    var canhBao = rowHasStock ? "OK" : "Chưa có TON_Q7";
    if (rowHasStock) {
      var thieu = it.totalQty - stock;
      canhBao = thieu > 0 ? ("THIẾU " + thieu) : (stockSource === "TON_VARIANT" ? "OK (variant)" : "OK");
      if (thieu > 0) missingLines += 1;
    }

    var rowOut = [0, maPack, it.maVach || "", tenPack || "", dvtOut || "", Math.max(0, Number(stock) || 0), it.totalQty];
    for (var c = 0; c < storeList.length; c++) rowOut.push(it.byStore[storeList[c]] || 0);
    rowOut.push(canhBao);
    rowOut.push(sourceStores.map(function(name) { return activeMap[name] || name; }).join(", "));
    rows.push(rowOut);
  }
  rows.sort(function(a, b) {
    var aWarn = String(a[warningCol - 1] || "");
    var bWarn = String(b[warningCol - 1] || "");
    if (stockReady && aWarn !== bWarn) return aWarn.indexOf("THIẾU") === 0 ? -1 : 1;
    return String(a[3] || "").localeCompare(String(b[3] || ""));
  });
  for (var r = 0; r < rows.length; r++) rows[r][0] = r + 1;

  function padRow_(arr, width) {
    var out = [];
    for (var pi = 0; pi < width; pi++) out.push(pi < arr.length ? arr[pi] : "");
    return out;
  }
  var colCount = headers.length;
  var modeLabelSheet = packingMode === "main"
    ? ("Chính: " + includedMainRows + " dòng | " + mainWindowLabel)
    : (packingMode === "supp"
      ? ("Bổ sung: " + includedNewRows + " dòng | " + newAfterLabel + " → " + newBeforeLabel + " (không gồm mốc 10:00)")
      : ("Tổng: Chính " + includedMainRows + " + Bổ sung " + includedNewRows + " dòng | " + (totalWindowLabel || (mainWindowLabel + " → " + newBeforeLabel))));
  var summaryLine = "Tổng đơn: " + Object.keys(orderSeen).length + " | Tổng mã: " + rows.length +
    (stockReady ? (" | Mã thiếu: " + missingLines + " | Tồn: TON_Q7") : " | Tồn: chưa có sheet TON_Q7 — Admin import lại file tồn") +
    " | " + modeLabelSheet;
  var sheetBlock = [
    padRow_([title], colCount),
    padRow_([
      "Ngày tổng hợp N2: " + packingDayTitle +
      " | Chính: N1 10:00 → N2 08:00 (≥10:00 & <08:00)" +
      " | Bổ sung: N2 08:00 → N2 10:00 (≥08:00 & <10:00)" +
      " | Tổng: N1 10:00 → N2 10:00 | Mode: " + packingMode +
      " | Stock từ " + TON_Q7_SHEET_NAME
    ], colCount),
    padRow_([""], colCount),
    padRow_([summaryLine], colCount),
    padRow_(storeOrderCountRow, colCount),
    padRow_(headers, colCount)
  ];
  for (var rb = 0; rb < rows.length; rb++) sheetBlock.push(padRow_(rows[rb], colCount));
  reportSheet.getRange(1, 1, sheetBlock.length, colCount).setValues(sheetBlock);
  reportSheet.setFrozenRows(headerRow);
  _dbgMark("writeSheetData", { outputRows: rows.length, colCount: colCount, missingLines: missingLines, stockReady: stockReady });

  // Ghi ĐVT đã resolve từ Data_Excel vào cột J (ĐVT thực tế) của lịch sử — tối đa 400 ô/lần
  var backfilled = 0;
  if (historyDvtBackfill.length) {
    var maxBackfill = Math.min(historyDvtBackfill.length, 400);
    for (var bf = 0; bf < maxBackfill; bf++) {
      try {
        historySheet.getRange(historyDvtBackfill[bf].sheetRow, 10).setValue(historyDvtBackfill[bf].dvt);
        backfilled++;
      } catch (bfErr) {}
    }
  }
  _dbgMark("backfillHistoryDvt", { candidates: historyDvtBackfill.length, written: backfilled });

  SpreadsheetApp.flush();
  _dbgMark("flush", {});
  var _dbgTotalMs = Date.now() - _dbgT0;

  return {
    success: true,
    sheetName: sheetName,
    totalOrders: Object.keys(orderSeen).length,
    totalItems: rows.length,
    missingItems: missingLines,
    stockReady: stockReady,
    stockSource: TON_Q7_SHEET_NAME,
    onlyNewItems: packingMode === "supp",
    packingMode: packingMode,
    packingDay: baseDateStr,
    mainWindowLabel: mainWindowLabel,
    totalWindowLabel: totalWindowLabel,
    newAfterLabel: newAfterLabel,
    newBeforeLabel: newBeforeLabel,
    url: "https://docs.google.com/spreadsheets/d/" + SHEET_ID + "/edit#gid=" + reportSheet.getSheetId(),
    _debugTimings: _dbgSteps,
    _debugTotalMs: _dbgTotalMs,
    _debugRun: "packing-window-v2",
    _debugIncludedMainRows: includedMainRows,
    _debugIncludedNewRows: includedNewRows,
    _debugSkippedByTime: skippedByTime,
    _debugSkippedNotSupplement: skippedNotSupplement,
    _debugDvtFromOrder: _dbgDvtFromOrder,
    _debugDvtCatalog: _dbgDvtCatalog,
    _debugDvtInferred: _dbgDvtInferred,
    _debugDvtEmpty: _dbgDvtEmpty,
    _debugDvtSample: _dbgDvtSample,
    _debugDvtBackfill: backfilled
  };
}


/**
 * Lịch tuần gom đơn đa kho động (Thứ 2 → Chủ Nhật).
 * Đọc sheet 1 lần (RAM) + CacheService 5 phút (key cal_{weekStart}).
 * @param {string} weekStartYYYYMMDD Monday hoặc bất kỳ ngày trong tuần
 * @param {string} userRole
 * @param {string} userStore
 */
function getPackingWeekCalendar_(weekStartYYYYMMDD, userRole, userStore) {
  var t0 = Date.now();
  var tz = Session.getScriptTimeZone() || "Asia/Ho_Chi_Minh";

  var anchor = parseDateInputYYYYMMDD(weekStartYYYYMMDD) || getScriptTodayStart_() || new Date();
  anchor = new Date(anchor.getFullYear(), anchor.getMonth(), anchor.getDate(), 0, 0, 0, 0);
  var dow = anchor.getDay(); // 0 Sun … 6 Sat
  var offsetToMon = (dow + 6) % 7; // Mon→0 … Sun→6
  var weekStart = new Date(anchor.getFullYear(), anchor.getMonth(), anchor.getDate() - offsetToMon, 0, 0, 0, 0);
  var weekEnd = new Date(weekStart.getFullYear(), weekStart.getMonth(), weekStart.getDate() + 7, 0, 0, 0, 0);
  var weekStartMs = weekStart.getTime();
  var weekEndMs = weekEnd.getTime();
  var weekStartKey = Utilities.formatDate(weekStart, tz, "yyyy-MM-dd");
  var cacheKey = "cal_" + weekStartKey;

  var cache = null;
  try { cache = getScriptCache_(); } catch (eCacheGet) { cache = null; }
  var cached = null;
  try { cached = getCacheJson_(cache, cacheKey); } catch (eCacheRead) { cached = null; }
  if (cached && cached.success && cached.days) {
    var fromCache = filterPackingWeekCalendarForUser_(cached, userRole, userStore);
    fromCache._debugTotalMs = Date.now() - t0;
    fromCache._debugCache = "HIT";
    fromCache._debugRun = "packing-week-calendar-v3-cache";
    return fromCache;
  }

  var ss = getSS();
  var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
  if (!historySheet) {
    return { success: false, error: "Không có sheet Lịch Sử Xuất Kho.", days: [], warehouses: [] };
  }

  // Đọc ĐÚNG 1 LẦN toàn bộ vùng dữ liệu → xử lý trong RAM
  var data = [];
  try {
    data = historySheet.getDataRange().getValues() || [];
  } catch (eRead) {
    Logger.log("getPackingWeekCalendar_ getDataRange error: " + (eRead.message || eRead));
    data = [];
  }

  var map = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row) continue;
    var soPhieu = row[1] ? String(row[1]).trim() : "";
    if (!soPhieu) continue;
    var createdMs = toHoChiMinhMillis_(row[0]);
    if (isNaN(createdMs) || createdMs < weekStartMs || createdMs >= weekEndMs) continue;

    var khoXuat = row[2] ? String(row[2]).trim() : "";
    var khoNhan = row[3] ? String(row[3]).trim() : "";
    var rowStatus = row[12] ? String(row[12]).trim() : "Mới";
    if (rowStatus === "Đã hủy đơn" || rowStatus === "Hủy (Trùng đơn)") continue;
    var isLineCancel = rowStatus === "Đã hủy dòng";

    if (!map[soPhieu]) {
      map[soPhieu] = {
        soPhieu: soPhieu,
        khoXuat: khoXuat,
        khoNhan: khoNhan,
        createdAtMs: createdMs,
        totalQty: 0,
        skuQty: {}
      };
    } else if (createdMs < map[soPhieu].createdAtMs) {
      map[soPhieu].createdAtMs = createdMs;
      if (khoXuat) map[soPhieu].khoXuat = khoXuat;
      if (khoNhan) map[soPhieu].khoNhan = khoNhan;
    }
    if (!isLineCancel) {
      var slCal = Number(row[7]) || 0;
      map[soPhieu].totalQty += slCal;
      var mhCal = normalizeOrderCodeText(row[4] || "");
      if (mhCal) map[soPhieu].skuQty[mhCal] = (map[soPhieu].skuQty[mhCal] || 0) + slCal;
    }
  }

  // Gắn nghi trùng trên toàn bộ đơn tuần (trước khi chia ngày)
  var weekOrderList = [];
  for (var spAll in map) {
    if (!Object.prototype.hasOwnProperty.call(map, spAll)) continue;
    var mo = map[spAll];
    weekOrderList.push({
      soPhieu: mo.soPhieu,
      khoXuat: mo.khoXuat,
      khoNhan: mo.khoNhan,
      createdAt: mo.createdAtMs,
      thoiGianDatMillis: mo.createdAtMs,
      totalQty: mo.totalQty || 0,
      skuSignature: buildOrderSkuSignature_(mo.skuQty || {}),
      groupKey: normalizeStoreName(mo.khoNhan || "") || String(mo.khoNhan || "").trim()
    });
  }
  attachDuplicateSuspects_(weekOrderList);
  var dupBySoPhieu = {};
  for (var di = 0; di < weekOrderList.length; di++) {
    if (weekOrderList[di].duplicateSuspect) {
      dupBySoPhieu[weekOrderList[di].soPhieu] = weekOrderList[di].duplicateSuspect;
    }
  }

  // Precompute 7 ngày + bucket 1 lần (không quét lại map × 7)
  var weekdayNames = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6", "Thứ 7", "Chủ Nhật"];
  var dayMs = 24 * 60 * 60 * 1000;
  var dayMeta = [];
  var dayBuckets = [];
  for (var d = 0; d < 7; d++) {
    var dayDate = new Date(weekStart.getFullYear(), weekStart.getMonth(), weekStart.getDate() + d, 0, 0, 0, 0);
    dayMeta.push({
      dateStr: Utilities.formatDate(dayDate, tz, "yyyy-MM-dd"),
      dateLabel: Utilities.formatDate(dayDate, tz, "dd/MM"),
      weekday: weekdayNames[d],
      startMs: dayDate.getTime()
    });
    dayBuckets.push({});
  }

  var warehouseIndex = {};
  for (var sp in map) {
    if (!Object.prototype.hasOwnProperty.call(map, sp)) continue;
    var ord = map[sp];
    var dayIndex = Math.floor((ord.createdAtMs - weekStartMs) / dayMs);
    if (dayIndex < 0 || dayIndex > 6) continue;

    var rawKho = ord.khoNhan || ord.khoXuat || "Khác";
    var khoKey = normalizeStoreName(rawKho) || String(rawKho).trim() || "other";
    var khoLabel = formatStoreDisplayLabel_(rawKho) || formatShortStoreLabel(rawKho) || rawKho;
    var byWh = dayBuckets[dayIndex];
    if (!byWh[khoKey]) {
      byWh[khoKey] = { khoKey: khoKey, khoLabel: khoLabel, khoRaw: rawKho, orders: [] };
    }
    if (!warehouseIndex[khoKey]) {
      warehouseIndex[khoKey] = { khoKey: khoKey, khoLabel: khoLabel, khoRaw: rawKho, orderCount: 0 };
    }
    warehouseIndex[khoKey].orderCount++;

    var hh = Utilities.formatDate(new Date(ord.createdAtMs), tz, "HH:mm");
    var dupInfo = dupBySoPhieu[ord.soPhieu] || null;
    byWh[khoKey].orders.push({
      soPhieu: ord.soPhieu,
      khoXuat: ord.khoXuat,
      khoNhan: ord.khoNhan,
      createdAt: ord.createdAtMs,
      createdTime: hh,
      thoiGian: formatOrderCreatedAtLabel_(ord.createdAtMs),
      thoiGianPretty: formatOrderCreatedAtPretty_(ord.createdAtMs),
      totalQty: ord.totalQty || 0,
      duplicateSuspect: dupInfo,
      isDuplicateSuspect: !!(dupInfo && !dupInfo.acknowledged)
    });
  }

  var days = [];
  for (var di2 = 0; di2 < 7; di2++) {
    var byWhDay = dayBuckets[di2];
    var warehouses = [];
    for (var wk in byWhDay) {
      if (!Object.prototype.hasOwnProperty.call(byWhDay, wk)) continue;
      byWhDay[wk].orders.sort(function(a, b) { return (a.createdAt || 0) - (b.createdAt || 0); });
      warehouses.push(byWhDay[wk]);
    }
    warehouses.sort(function(a, b) {
      return String(a.khoLabel || "").localeCompare(String(b.khoLabel || ""), "vi");
    });
    var dayOrderCount = 0;
    for (var wi = 0; wi < warehouses.length; wi++) dayOrderCount += warehouses[wi].orders.length;
    days.push({
      date: dayMeta[di2].dateStr,
      dateLabel: dayMeta[di2].dateLabel,
      weekday: dayMeta[di2].weekday,
      isToday: false,
      orderCount: dayOrderCount,
      warehouses: warehouses
    });
  }

  var todayStr = Utilities.formatDate(getScriptTodayStart_() || new Date(), tz, "yyyy-MM-dd");
  for (var ti = 0; ti < days.length; ti++) {
    if (days[ti].date === todayStr) days[ti].isToday = true;
  }

  var warehousesAll = [];
  for (var wkey in warehouseIndex) {
    if (Object.prototype.hasOwnProperty.call(warehouseIndex, wkey)) warehousesAll.push(warehouseIndex[wkey]);
  }
  warehousesAll.sort(function(a, b) {
    return String(a.khoLabel || "").localeCompare(String(b.khoLabel || ""), "vi");
  });

  var result = {
    success: true,
    weekStart: weekStartKey,
    weekEnd: Utilities.formatDate(new Date(weekEnd.getTime() - 1), tz, "yyyy-MM-dd"),
    weekLabel: Utilities.formatDate(weekStart, tz, "dd/MM") +
      " → " + Utilities.formatDate(new Date(weekEnd.getTime() - 1), tz, "dd/MM/yyyy"),
    days: days,
    warehouses: warehousesAll,
    totalOrders: Object.keys(map).length,
    _debugRun: "packing-week-calendar-v3-cache"
  };

  try { putCacheJson_(cache, cacheKey, result, 300); } catch (eCachePut) {}

  var filtered = filterPackingWeekCalendarForUser_(result, userRole, userStore);
  filtered._debugTotalMs = Date.now() - t0;
  filtered._debugCache = "MISS";
  filtered._debugRowsScanned = data.length;
  return filtered;
}


/** Lọc lịch tuần theo kho user (Admin = full). Không mutate object cache. */
function filterPackingWeekCalendarForUser_(payload, userRole, userStore) {
  if (!payload || typeof payload !== "object") return payload;
  if (userRole === "Admin") {
    return payload;
  }
  var filterStore = normalizeStoreName(userStore || "");
  if (!filterStore) return payload;

  var srcDays = payload.days || [];
  var days = [];
  var warehouseIndex = {};
  var totalOrders = 0;

  for (var d = 0; d < srcDays.length; d++) {
    var day = srcDays[d] || {};
    var srcWh = day.warehouses || [];
    var warehouses = [];
    var dayOrderCount = 0;
    for (var w = 0; w < srcWh.length; w++) {
      var wh = srcWh[w] || {};
      var ordersIn = wh.orders || [];
      var ordersOut = [];
      for (var oi = 0; oi < ordersIn.length; oi++) {
        var ord = ordersIn[oi];
        if (!ord) continue;
        if (!isSameStoreName(ord.khoXuat, filterStore) && !isSameStoreName(ord.khoNhan, filterStore)) continue;
        ordersOut.push(ord);
      }
      if (!ordersOut.length) continue;
      warehouses.push({
        khoKey: wh.khoKey,
        khoLabel: wh.khoLabel,
        khoRaw: wh.khoRaw,
        orders: ordersOut
      });
      dayOrderCount += ordersOut.length;
      totalOrders += ordersOut.length;
      if (!warehouseIndex[wh.khoKey]) {
        warehouseIndex[wh.khoKey] = {
          khoKey: wh.khoKey,
          khoLabel: wh.khoLabel,
          khoRaw: wh.khoRaw,
          orderCount: 0
        };
      }
      warehouseIndex[wh.khoKey].orderCount += ordersOut.length;
    }
    days.push({
      date: day.date,
      dateLabel: day.dateLabel,
      weekday: day.weekday,
      isToday: !!day.isToday,
      orderCount: dayOrderCount,
      warehouses: warehouses
    });
  }

  var warehousesAll = [];
  for (var wk in warehouseIndex) {
    if (Object.prototype.hasOwnProperty.call(warehouseIndex, wk)) warehousesAll.push(warehouseIndex[wk]);
  }
  warehousesAll.sort(function(a, b) {
    return String(a.khoLabel || "").localeCompare(String(b.khoLabel || ""), "vi");
  });

  return {
    success: payload.success,
    weekStart: payload.weekStart,
    weekEnd: payload.weekEnd,
    weekLabel: payload.weekLabel,
    days: days,
    warehouses: warehousesAll,
    totalOrders: totalOrders,
    _debugRun: payload._debugRun || "packing-week-calendar-v3-cache"
  };
}


// ========== MODULE: Duplicate detector + Branch SKU history (additive) ==========

function buildOrderSkuSignature_(skuQtyMap) {
  var keys = Object.keys(skuQtyMap || {}).sort();
  var parts = [];
  for (var i = 0; i < keys.length; i++) {
    parts.push(keys[i] + ":" + (Number(skuQtyMap[keys[i]]) || 0));
  }
  return parts.join("|");
}

function sameCalendarDayMs_(msA, msB) {
  if (isNaN(msA) || isNaN(msB)) return false;
  var tz = Session.getScriptTimeZone() || "Asia/Ho_Chi_Minh";
  try {
    return Utilities.formatDate(new Date(msA), tz, "yyyy-MM-dd") ===
      Utilities.formatDate(new Date(msB), tz, "yyyy-MM-dd");
  } catch (e) {
    return false;
  }
}

/**
 * Cùng kho nhận + (cùng ngày hoặc ≤60 phút) + (cùng tổng SL hoặc cùng chữ ký MaSP).
 * Gắn order.duplicateSuspect / isDuplicateSuspect (không đổi struct cũ ngoài field mới).
 */
function attachDuplicateSuspects_(orders) {
  var list = orders || [];
  for (var i = 0; i < list.length; i++) {
    var a = list[i];
    if (!a || a.duplicateSuspect) continue;
    var aMs = Number(a.thoiGianDatMillis || a.createdAt || a.createdAtMs || 0) || 0;
    var aStore = normalizeStoreName(a.khoNhan || "") || String(a.khoNhan || a.groupKey || "").trim();
    if (!aStore || !aMs) continue;
    for (var j = i + 1; j < list.length; j++) {
      var b = list[j];
      if (!b) continue;
      var bMs = Number(b.thoiGianDatMillis || b.createdAt || b.createdAtMs || 0) || 0;
      var bStore = normalizeStoreName(b.khoNhan || "") || String(b.khoNhan || b.groupKey || "").trim();
      if (!bStore || !bMs) continue;
      if (!isSameStoreName(aStore, bStore)) continue;
      var deltaMin = Math.abs(aMs - bMs) / 60000;
      var timeOk = sameCalendarDayMs_(aMs, bMs) || deltaMin <= 60;
      if (!timeOk) continue;
      var qtyA = Number(a.totalQty || 0);
      var qtyB = Number(b.totalQty || 0);
      var sameQty = qtyA > 0 && qtyA === qtyB;
      var sigA = String(a.skuSignature || "");
      var sigB = String(b.skuSignature || "");
      var sameSku = !!(sigA && sigB && sigA === sigB);
      if (!sameQty && !sameSku) continue;
      var reason = sameSku && sameQty ? "cùng danh mục mã & tổng SL" :
        (sameSku ? "cùng danh mục mã hàng" : "cùng tổng số lượng");
      var aUi = formatOrderTimestampUi_(aMs) || formatOrderCreatedAtPretty_(aMs) || "";
      var bUi = formatOrderTimestampUi_(bMs) || formatOrderCreatedAtPretty_(bMs) || "";
      if (!a.duplicateSuspect) {
        a.duplicateSuspect = {
          peerSoPhieu: b.soPhieu,
          peerCreatedAt: bMs,
          peerCreatedUi: bUi,
          reason: reason,
          acknowledged: isDuplicateAcknowledged_(a.soPhieu)
        };
        a.isDuplicateSuspect = !a.duplicateSuspect.acknowledged;
      }
      if (!b.duplicateSuspect) {
        b.duplicateSuspect = {
          peerSoPhieu: a.soPhieu,
          peerCreatedAt: aMs,
          peerCreatedUi: aUi,
          reason: reason,
          acknowledged: isDuplicateAcknowledged_(b.soPhieu)
        };
        b.isDuplicateSuspect = !b.duplicateSuspect.acknowledged;
      }
    }
    if (a.duplicateSuspect && a.duplicateSuspect.acknowledged) a.isDuplicateSuspect = false;
  }
  return list;
}

function isDuplicateAcknowledged_(soPhieu) {
  try {
    var key = "dup_ack_" + normalizeOrderCodeText(soPhieu || "");
    if (!key || key === "dup_ack_") return false;
    var raw = PropertiesService.getDocumentProperties().getProperty(key);
    return !!raw;
  } catch (e) {
    return false;
  }
}

function acknowledgeDuplicateOrder_(soPhieu, actor) {
  try {
    var code = String(soPhieu || "").trim();
    if (!code) return { success: false, error: "Thiếu số phiếu." };
    var key = "dup_ack_" + normalizeOrderCodeText(code);
    PropertiesService.getDocumentProperties().setProperty(key, JSON.stringify({
      at: Date.now(),
      actor: String(actor || ""),
      soPhieu: code
    }));
    try {
      logOrderChange(getSS(), code, "Chấp nhận đơn (không trùng)", actor || "", "", "", "", "", "Ẩn cảnh báo trùng");
    } catch (eLog) {}
    return { success: true, soPhieu: code, acknowledged: true };
  } catch (err) {
    return { success: false, error: err.message || String(err) };
  }
}

/**
 * Chống double-submit: cùng chi nhánh + đơn giống trong ≤5 phút.
 */
function checkDuplicateBeforeSave_(khoNhan, items, userRole, userStore) {
  try {
    var ss = getSS();
    var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
    if (!historySheet) return { success: true, isDuplicate: false };
    var store = normalizeStoreName(khoNhan || userStore || "");
    var skuQty = {};
    var totalQty = 0;
    var arr = items || [];
    for (var i = 0; i < arr.length; i++) {
      var it = arr[i];
      if (!it) continue;
      var sl = Number(it.sl) || 0;
      if (sl <= 0) continue;
      totalQty += sl;
      var mh = normalizeOrderCodeText(it.maHang || "");
      if (mh) skuQty[mh] = (skuQty[mh] || 0) + sl;
    }
    var sig = buildOrderSkuSignature_(skuQty);
    var nowMs = Date.now();
    var windowMs = 5 * 60 * 1000;
    var pack = readHistoryDataPack_(historySheet, 3000);
    var data = pack.data || [[]];
    var map = {};
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      if (!row) continue;
      var sp = row[1] ? String(row[1]).trim() : "";
      if (!sp) continue;
      var st = row[12] ? String(row[12]).trim() : "Mới";
      if (st === "Đã hủy đơn" || st === "Hủy (Trùng đơn)" || st === "Đã hủy dòng") continue;
      var kn = row[3] ? String(row[3]).trim() : "";
      if (store && !isSameStoreName(kn, store)) continue;
      var cms = toHoChiMinhMillis_(row[0]);
      if (isNaN(cms) || (nowMs - cms) > windowMs || cms > nowMs + 60000) continue;
      if (!map[sp]) {
        map[sp] = { soPhieu: sp, khoNhan: kn, createdAtMs: cms, totalQty: 0, skuQty: {} };
      }
      if (cms < map[sp].createdAtMs) map[sp].createdAtMs = cms;
      var slR = Number(row[7]) || 0;
      map[sp].totalQty += slR;
      var mhR = normalizeOrderCodeText(row[4] || "");
      if (mhR) map[sp].skuQty[mhR] = (map[sp].skuQty[mhR] || 0) + slR;
    }
    for (var k in map) {
      if (!Object.prototype.hasOwnProperty.call(map, k)) continue;
      var o = map[k];
      var oSig = buildOrderSkuSignature_(o.skuQty);
      var sameQty = totalQty > 0 && totalQty === o.totalQty;
      var sameSku = !!(sig && oSig && sig === oSig);
      if (!sameQty && !sameSku) continue;
      var mins = Math.round(Math.abs(nowMs - o.createdAtMs) / 60000);
      return {
        success: true,
        isDuplicate: true,
        withinMinutes: mins,
        peerSoPhieu: o.soPhieu,
        peerCreatedAt: o.createdAtMs,
        peerCreatedUi: formatOrderTimestampUi_(o.createdAtMs) || "",
        khoNhan: o.khoNhan,
        reason: sameSku && sameQty ? "cùng danh mục & tổng SL" : (sameSku ? "cùng danh mục mã" : "cùng tổng SL")
      };
    }
    return { success: true, isDuplicate: false };
  } catch (err) {
    return { success: true, isDuplicate: false, warning: err.message || String(err) };
  }
}

/**
 * Lịch sử xuất MaSP của chi nhánh trong N ngày (mặc định 7) — phục vụ SKU History Checker.
 */
function getBranchSkuHistory_(khoNhan, excludeSoPhieu, daysBack) {
  var t0 = Date.now();
  try {
    var ss = getSS();
    var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
    if (!historySheet) return { success: false, error: "Không có lịch sử.", bySku: {} };
    var store = normalizeStoreName(khoNhan || "");
    if (!store) return { success: true, bySku: {}, orders: [] };
    var days = Number(daysBack);
    if (!days || days < 1) days = 7;
    if (days > 31) days = 31;
    var today = getScriptTodayStart_() || new Date();
    var start = new Date(today.getFullYear(), today.getMonth(), today.getDate() - (days - 1), 0, 0, 0, 0);
    var startMs = start.getTime();
    var exclude = normalizeOrderCodeText(excludeSoPhieu || "");
    var pack = readHistoryDataPack_(historySheet, 6000);
    var data = pack.data || [[]];
    var bySku = {};
    var orderSeen = {};
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row) continue;
      var sp = row[1] ? String(row[1]).trim() : "";
      if (!sp) continue;
      if (exclude && normalizeOrderCodeText(sp) === exclude) continue;
      var kn = row[3] ? String(row[3]).trim() : "";
      if (!isSameStoreName(kn, store)) continue;
      var st = row[12] ? String(row[12]).trim() : "Mới";
      if (st === "Đã hủy đơn" || st === "Hủy (Trùng đơn)" || st === "Đã hủy dòng") continue;
      var cms = toHoChiMinhMillis_(row[0]);
      if (isNaN(cms) || cms < startMs) continue;
      var mh = String(row[4] || "").trim();
      var mhKey = normalizeOrderCodeText(mh);
      if (!mhKey) continue;
      var qty = Number(row[7]) || 0;
      if (qty <= 0) continue;
      var dvt = row[9] ? String(row[9]).trim() : "";
      var prev = bySku[mhKey];
      if (!prev || cms >= prev.createdAtMs) {
        var tzHist = Session.getScriptTimeZone() || "Asia/Ho_Chi_Minh";
        bySku[mhKey] = {
          maHang: mh,
          maHangKey: mhKey,
          qty: qty,
          dvt: dvt,
          soPhieu: sp,
          createdAtMs: cms,
          createdUi: formatOrderTimestampUi_(cms) || formatOrderCreatedAtLabel_(cms) || "",
          dateLabel: Utilities.formatDate(new Date(cms), tzHist, "dd/MM/yyyy"),
          storeLabel: formatStoreDisplayLabel_(kn) || kn
        };
      }
      orderSeen[sp] = true;
    }
    return {
      success: true,
      khoNhan: khoNhan,
      storeKey: store,
      daysBack: days,
      excludeSoPhieu: excludeSoPhieu || "",
      bySku: bySku,
      orderCount: Object.keys(orderSeen).length,
      skuCount: Object.keys(bySku).length,
      cacheKey: "bsh_v1_" + store + "_" + days,
      _debugTotalMs: Date.now() - t0,
      _debugRun: "branch-sku-history-v1"
    };
  } catch (err) {
    return { success: false, error: err.message || String(err), bySku: {} };
  }
}

/** Chi tiết nghi trùng cho 1 phiếu (header QL / Soạn) */
function getOrderDuplicateInfo_(soPhieu) {
  try {
    var code = String(soPhieu || "").trim();
    if (!code) return { success: false, isDuplicateSuspect: false };
    if (isDuplicateAcknowledged_(code)) {
      return { success: true, isDuplicateSuspect: false, acknowledged: true, soPhieu: code };
    }
    var info = getThongTinPhieu(code);
    if (!info) return { success: false, isDuplicateSuspect: false, error: "Không tìm thấy phiếu." };
    var detail = getOrderDetail(code);
    var items = (detail && detail.items) || [];
    var skuQty = {};
    var totalQty = 0;
    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      var sl = Number(it.slGoc != null ? it.slGoc : it.sl) || 0;
      totalQty += sl;
      var mh = normalizeOrderCodeText(it.maHang || "");
      if (mh) skuQty[mh] = (skuQty[mh] || 0) + sl;
    }
    var createdMs = info.thoiGianTaoMs || info.thoiGianMs || info.createdAt;
    var ss = getSS();
    var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
    if (!historySheet) return { success: true, isDuplicateSuspect: false };
    var pack = readHistoryDataPack_(historySheet, 5000);
    var data = pack.data || [[]];
    var map = {};
    var windowPad = 24 * 60 * 60 * 1000; // ±1 ngày quanh mốc tạo
    var cMs = Number(createdMs) || 0;
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      if (!row) continue;
      var sp = row[1] ? String(row[1]).trim() : "";
      if (!sp || normalizeOrderCodeText(sp) === normalizeOrderCodeText(code)) continue;
      var st = row[12] ? String(row[12]).trim() : "Mới";
      if (st === "Đã hủy đơn" || st === "Hủy (Trùng đơn)" || st === "Đã hủy dòng") continue;
      var kn = row[3] ? String(row[3]).trim() : "";
      if (!isSameStoreName(kn, info.khoNhan)) continue;
      var cms = toHoChiMinhMillis_(row[0]);
      if (isNaN(cms)) continue;
      if (cMs && Math.abs(cms - cMs) > windowPad) continue;
      if (!map[sp]) map[sp] = { soPhieu: sp, khoNhan: kn, createdAtMs: cms, totalQty: 0, skuQty: {} };
      if (cms < map[sp].createdAtMs) map[sp].createdAtMs = cms;
      var slR = Number(row[7]) || 0;
      map[sp].totalQty += slR;
      var mhR = normalizeOrderCodeText(row[4] || "");
      if (mhR) map[sp].skuQty[mhR] = (map[sp].skuQty[mhR] || 0) + slR;
    }
    var selfSig = buildOrderSkuSignature_(skuQty);
    var selfOrder = {
      soPhieu: code,
      khoNhan: info.khoNhan,
      createdAt: cMs,
      thoiGianDatMillis: cMs,
      totalQty: totalQty,
      skuSignature: selfSig
    };
    var peers = [selfOrder];
    for (var k in map) {
      if (!Object.prototype.hasOwnProperty.call(map, k)) continue;
      peers.push({
        soPhieu: map[k].soPhieu,
        khoNhan: map[k].khoNhan,
        createdAt: map[k].createdAtMs,
        thoiGianDatMillis: map[k].createdAtMs,
        totalQty: map[k].totalQty,
        skuSignature: buildOrderSkuSignature_(map[k].skuQty)
      });
    }
    attachDuplicateSuspects_(peers);
    var hit = peers[0].duplicateSuspect || null;
    return {
      success: true,
      soPhieu: code,
      isDuplicateSuspect: !!(hit && !hit.acknowledged),
      acknowledged: !!(hit && hit.acknowledged) || isDuplicateAcknowledged_(code),
      duplicateSuspect: hit,
      khoNhan: info.khoNhan,
      totalQty: totalQty
    };
  } catch (err) {
    return { success: false, isDuplicateSuspect: false, error: err.message || String(err) };
  }
}
