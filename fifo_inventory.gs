// ============================================================
// fifo_inventory.gs — FIFO / Stock In-Out / In-Memory Cache / Lock
// Thuần Google Sheets + Apps Script (không Supabase)
// ============================================================

var STOCK_LOTS_SHEET_NAME = "STOCK_LOTS";
var STOCK_LOTS_HEADERS = [
  "LotId",       // A
  "MaHang",      // B
  "MaVach",      // C
  "QtyRemain",   // D
  "QtyOriginal", // E
  "NgayNhap",    // F
  "UnitCost",    // G
  "Dvt",         // H
  "Kho",         // I
  "UpdatedAt",   // J
  "Note"         // K
];
var STOCK_LOTS_COL_COUNT = 11;
var CACHE_STOCK_LOTS_KEY = "stock_lots_v1";
var CACHE_FIFO_TTL_SECONDS = 300;

/** Cache trong 1 request GAS (biến global reset mỗi lần invoke) */
var _INV_REQ_CACHE_ = null;

function getInventoryRequestCache_() {
  if (!_INV_REQ_CACHE_) {
    _INV_REQ_CACHE_ = {
      tonQ7Bundle: null,
      lotsRows: null,
      lotsByCode: null,
      seededAt: Date.now()
    };
  }
  return _INV_REQ_CACHE_;
}

function clearInventoryRequestCache_() {
  _INV_REQ_CACHE_ = null;
}

/** Xóa cả CacheService tồn/FIFO sau khi ghi sheet */
function invalidateInventoryCaches_() {
  clearInventoryRequestCache_();
  try {
    var cache = getScriptCache_();
    cache.remove(CACHE_TON_Q7_KEY);
    cache.remove(CACHE_TON_Q7_KEY + "_dvt");
    cache.remove(CACHE_STOCK_LOTS_KEY);
  } catch (eClr) {}
}


/**
 * Khóa Script-level chống race khi ghi tồn / FIFO.
 * Các hàm đã giữ DocumentLock có thể gọi apply* với acquireLock=false.
 */
function withInventoryLock_(fn, waitMs) {
  var lock = LockService.getScriptLock();
  lock.waitLock(waitMs || 30000);
  try {
    return fn();
  } finally {
    try { lock.releaseLock(); } catch (eRel) {}
  }
}


// ---------- Sheet STOCK_LOTS ----------

function ensureStockLotsSheet_(ss) {
  ss = ss || getSS();
  var sh = ss.getSheetByName(STOCK_LOTS_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(STOCK_LOTS_SHEET_NAME);
    sh.getRange(1, 1, 1, STOCK_LOTS_COL_COUNT).setValues([STOCK_LOTS_HEADERS]);
    sh.setFrozenRows(1);
    try { sh.getRange(1, 1, 1, STOCK_LOTS_COL_COUNT).setFontWeight("bold"); } catch (eH) {}
  } else if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, STOCK_LOTS_COL_COUNT).setValues([STOCK_LOTS_HEADERS]);
  }
  return sh;
}


function newStockLotId_() {
  return "LOT-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Ho_Chi_Minh", "yyyyMMdd-HHmmss") +
    "-" + String(Math.floor(Math.random() * 9000) + 1000);
}


/**
 * Đọc TON_Q7 1 lần / request: RAM → CacheService → Sheet.
 * Tái sử dụng readTonKhoQ7Bundle_ (đã có cache sheet).
 */
function getTonQ7BundleCached_(ss) {
  var req = getInventoryRequestCache_();
  if (req.tonQ7Bundle && req.tonQ7Bundle.map) return req.tonQ7Bundle;
  var bundle = readTonKhoQ7Bundle_(ss || getSS());
  req.tonQ7Bundle = bundle || { map: null, dvtLabels: {}, source: "none" };
  return req.tonQ7Bundle;
}

function getTonQ7MapCached_(ss) {
  var b = getTonQ7BundleCached_(ss);
  return (b && b.map) ? b.map : {};
}


/**
 * Seed cache đầu request (TON_Q7 + STOCK_LOTS) — gọi 1 lần.
 */
function seedInventoryRequestCache_(ss) {
  ss = ss || getSS();
  getTonQ7BundleCached_(ss);
  loadStockLotsIndex_(ss);
  return getInventoryRequestCache_();
}


function loadStockLotsIndex_(ss, forceReload) {
  var req = getInventoryRequestCache_();
  if (!forceReload && req.lotsByCode) return req.lotsByCode;

  ss = ss || getSS();
  var cache = getScriptCache_();
  if (!forceReload) {
    var cached = getCacheJson_(cache, CACHE_STOCK_LOTS_KEY);
    if (cached && cached.byCode) {
      req.lotsByCode = cached.byCode;
      req.lotsRows = cached.rows || [];
      return req.lotsByCode;
    }
  }

  var sh = ensureStockLotsSheet_(ss);
  var lastRow = sh.getLastRow();
  var byCode = {};
  var rows = [];
  if (lastRow >= 2) {
    var data = sh.getRange(2, 1, lastRow - 1, STOCK_LOTS_COL_COUNT).getValues();
    for (var i = 0; i < data.length; i++) {
      var r = data[i];
      if (!r) continue;
      var lotId = String(r[0] || "").trim();
      var maHang = String(r[1] || "").replace(/^MH:/i, "").trim();
      if (!lotId || !maHang) continue;
      var qtyRemain = Math.max(0, Number(r[3]) || 0);
      var lot = {
        lotId: lotId,
        maHang: maHang,
        maVach: String(r[2] || "").trim(),
        qtyRemain: qtyRemain,
        qtyOriginal: Math.max(0, Number(r[4]) || 0),
        ngayNhap: r[5],
        ngayNhapMs: toLotTimeMs_(r[5]),
        unitCost: Number(r[6]) || 0,
        dvt: String(r[7] || "").trim(),
        kho: String(r[8] || "").trim(),
        note: String(r[10] || "").trim(),
        rowIndex: i + 2
      };
      rows.push(lot);
      var key = normalizeProductCode(maHang) || maHang.toUpperCase();
      if (!byCode[key]) byCode[key] = [];
      byCode[key].push(lot);
      // alias MH:
      var mhKey = "MH:" + key;
      if (!byCode[mhKey]) byCode[mhKey] = byCode[key];
    }
  }
  req.lotsRows = rows;
  req.lotsByCode = byCode;
  try {
    putCacheJson_(cache, CACHE_STOCK_LOTS_KEY, { byCode: byCode, rows: rows }, CACHE_FIFO_TTL_SECONDS);
  } catch (ePut) {}
  return byCode;
}


function toLotTimeMs_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value.getTime();
  if (typeof value === "number" && isFinite(value)) return value;
  var s = String(value || "").trim();
  if (!s) return 0;
  try {
    if (typeof toHoChiMinhMillis_ === "function") {
      var ms = toHoChiMinhMillis_(value);
      if (!isNaN(ms)) return ms;
    }
  } catch (e1) {}
  var d = new Date(s);
  return isNaN(d.getTime()) ? 0 : d.getTime();
}


function getLotsForItemCode_(ss, itemCode) {
  var code = String(itemCode || "").replace(/^MH:/i, "").trim();
  if (!code) return [];
  var byCode = loadStockLotsIndex_(ss);
  var key = normalizeProductCode(code) || code.toUpperCase();
  var list = byCode[key] || byCode["MH:" + key] || byCode[code] || [];
  // clone shallow — không mutate cache khi sort/trừ ảo
  return list.map(function(l) {
    return {
      lotId: l.lotId,
      maHang: l.maHang,
      maVach: l.maVach,
      qtyRemain: l.qtyRemain,
      qtyOriginal: l.qtyOriginal,
      ngayNhap: l.ngayNhap,
      ngayNhapMs: l.ngayNhapMs,
      unitCost: l.unitCost,
      dvt: l.dvt,
      kho: l.kho,
      note: l.note,
      rowIndex: l.rowIndex
    };
  });
}


/**
 * FIFO thuần JS: trừ lùi theo lô cũ → mới.
 * @returns {{success:boolean, itemCode:string, requestedQty:number, allocatedQty:number,
 *   shortfall:number, totalCost:number, allocations:Array}}
 */
function calculateFifoOut_(itemCode, requestedQty, lotList) {
  var code = String(itemCode || "").replace(/^MH:/i, "").trim();
  var need = Math.max(0, Number(requestedQty) || 0);
  var lots = (lotList || []).slice();
  lots.sort(function(a, b) {
    var ta = Number(a.ngayNhapMs != null ? a.ngayNhapMs : toLotTimeMs_(a.ngayNhap)) || 0;
    var tb = Number(b.ngayNhapMs != null ? b.ngayNhapMs : toLotTimeMs_(b.ngayNhap)) || 0;
    if (ta !== tb) return ta - tb;
    return String(a.lotId || "").localeCompare(String(b.lotId || ""));
  });

  var allocations = [];
  var remain = need;
  var totalCost = 0;
  for (var i = 0; i < lots.length && remain > 0; i++) {
    var lot = lots[i];
    var avail = Math.max(0, Number(lot.qtyRemain) || 0);
    if (avail <= 0) continue;
    var take = Math.min(avail, remain);
    var unitCost = Number(lot.unitCost) || 0;
    allocations.push({
      lotId: lot.lotId,
      maHang: lot.maHang || code,
      qty: take,
      unitCost: unitCost,
      lineCost: take * unitCost,
      ngayNhap: lot.ngayNhap,
      ngayNhapMs: lot.ngayNhapMs || toLotTimeMs_(lot.ngayNhap),
      dvt: lot.dvt || "",
      rowIndex: lot.rowIndex || null,
      qtyRemainAfter: avail - take
    });
    totalCost += take * unitCost;
    remain -= take;
    lot.qtyRemain = avail - take;
  }

  var allocatedQty = need - remain;
  return {
    success: remain <= 0,
    itemCode: code,
    requestedQty: need,
    allocatedQty: allocatedQty,
    shortfall: Math.max(0, remain),
    totalCost: totalCost,
    allocations: allocations
  };
}


/**
 * Nhập kho 1 lô → append STOCK_LOTS + cộng TON_Q7.
 * payload: { maHang, maVach?, qty, unitCost?, dvt?, kho?, ngayNhap?, note?, actor? }
 */
function stockInLot_(payload) {
  return withInventoryLock_(function() {
    var ss = getSS();
    var maHang = String((payload && payload.maHang) || "").replace(/^MH:/i, "").trim();
    var qty = Math.max(0, Number(payload && payload.qty) || 0);
    if (!maHang) return { success: false, error: "Thiếu maHang" };
    if (qty <= 0) return { success: false, error: "Số lượng nhập phải > 0" };

    var sh = ensureStockLotsSheet_(ss);
    var lotId = String((payload && payload.lotId) || "").trim() || newStockLotId_();
    var ngayNhap = (payload && payload.ngayNhap) ? payload.ngayNhap : new Date();
    var row = [
      lotId,
      maHang,
      String((payload && payload.maVach) || "").trim(),
      qty,
      qty,
      ngayNhap,
      Number(payload && payload.unitCost) || 0,
      String((payload && payload.dvt) || "").trim(),
      String((payload && payload.kho) || PACKING_STOCK_STORE || "").trim(),
      new Date(),
      String((payload && (payload.note || payload.actor)) || "").trim()
    ];
    sh.appendRow(row);

    // Cộng TON_Q7 (map bare key)
    try {
      var map = getTonQ7MapCached_(ss) || {};
      var labels = (getTonQ7BundleCached_(ss).dvtLabels) || {};
      var key = normalizeProductCode(maHang) || maHang.toUpperCase();
      map[key] = Math.max(0, (Number(map[key]) || 0) + qty);
      if (payload && payload.dvt && !labels[key]) labels[key] = String(payload.dvt).trim();
      writeTonQ7MapToSheet_(ss, map, labels);
    } catch (eQ7) {
      Logger.log("stockInLot_ TON_Q7 sync error: " + (eQ7.message || eQ7));
    }

    clearInventoryRequestCache_();
    invalidateInventoryCaches_();
    return {
      success: true,
      lotId: lotId,
      maHang: maHang,
      qty: qty,
      msg: "Đã nhập lô " + lotId
    };
  });
}


/**
 * Preview FIFO (không ghi sheet) — FE/Admin kiểm tra trước khi xuất.
 */
function previewFifoOut_(maHang, qty) {
  var ss = getSS();
  seedInventoryRequestCache_(ss);
  var lots = getLotsForItemCode_(ss, maHang);
  return calculateFifoOut_(maHang, qty, lots);
}


/**
 * Áp FIFO xuất kho: trừ QtyRemain trên STOCK_LOTS + trừ TON_Q7.
 * lines: [{maHang, maVach?, qty, dvt?}]
 * opts.acquireLock — mặc định false (caller đã lock).
 */
function applyFifoStockOutBatch_(ss, lines, opts) {
  opts = opts || {};
  var run = function() {
    ss = ss || getSS();
    if (!lines || !lines.length) return { success: true, changed: 0, allocations: [] };

    ensureStockLotsSheet_(ss);
    clearInventoryRequestCache_(); // đọc lots mới nhất trong lock
    seedInventoryRequestCache_(ss);

    var sh = ss.getSheetByName(STOCK_LOTS_SHEET_NAME);
    var allAllocations = [];
    var lotDeltas = {}; // rowIndex -> qty to subtract
    var q7Deltas = {}; // code -> qty to subtract
    var shortfalls = [];

    for (var i = 0; i < lines.length; i++) {
      var line = lines[i];
      if (!line) continue;
      var qty = Number(line.qty);
      if (!qty || qty <= 0 || isNaN(qty)) continue;
      var code = String(line.maHang || line.maVach || "").replace(/^MH:/i, "").trim();
      if (!code) continue;

      var lots = getLotsForItemCode_(ss, code);
      if (!lots.length) {
        // Không có lô FIFO — chỉ trừ TON_Q7 aggregate (tương thích cũ)
        var bare = normalizeProductCode(code) || code.toUpperCase();
        q7Deltas[bare] = (q7Deltas[bare] || 0) + qty;
        shortfalls.push({ maHang: code, requested: qty, shortfall: qty, reason: "no_lots" });
        continue;
      }

      var fifo = calculateFifoOut_(code, qty, lots);
      for (var a = 0; a < fifo.allocations.length; a++) {
        var al = fifo.allocations[a];
        allAllocations.push(al);
        if (al.rowIndex) {
          lotDeltas[al.rowIndex] = (lotDeltas[al.rowIndex] || 0) + al.qty;
        }
      }
      var bare2 = normalizeProductCode(code) || code.toUpperCase();
      q7Deltas[bare2] = (q7Deltas[bare2] || 0) + fifo.allocatedQty;
      if (fifo.shortfall > 0) {
        // Phần thiếu vẫn trừ Q7 nếu còn tồn aggregate
        q7Deltas[bare2] = (q7Deltas[bare2] || 0) + fifo.shortfall;
        shortfalls.push({ maHang: code, requested: qty, shortfall: fifo.shortfall, reason: "partial_fifo" });
      }
    }

    // Ghi STOCK_LOTS — batch theo dải dòng
    var lotRows = Object.keys(lotDeltas).map(Number).filter(function(n) { return n >= 2; }).sort(function(a, b) { return a - b; });
    if (lotRows.length && sh) {
      var minR = lotRows[0];
      var maxR = lotRows[lotRows.length - 1];
      var num = maxR - minR + 1;
      var mat = sh.getRange(minR, 4, num, 1).getValues(); // QtyRemain col D
      var nowMat = sh.getRange(minR, 10, num, 1).getValues(); // UpdatedAt col J
      var now = new Date();
      for (var li = 0; li < lotRows.length; li++) {
        var rr = lotRows[li];
        var off = rr - minR;
        var cur = Math.max(0, Number(mat[off][0]) || 0);
        mat[off][0] = Math.max(0, cur - (lotDeltas[rr] || 0));
        nowMat[off][0] = now;
      }
      sh.getRange(minR, 4, num, 1).setValues(mat);
      sh.getRange(minR, 10, num, 1).setValues(nowMat);
    }

    // Sync TON_Q7
    try {
      var bundle = readTonKhoQ7Bundle_(ss);
      var map = bundle.map || {};
      var labels = bundle.dvtLabels || {};
      var touched = false;
      for (var qk in q7Deltas) {
        if (!Object.prototype.hasOwnProperty.call(q7Deltas, qk)) continue;
        var before = Number(map[qk]);
        if (isNaN(before)) before = Number(map["MH:" + qk]) || 0;
        var after = Math.max(0, before - (Number(q7Deltas[qk]) || 0));
        map[qk] = after;
        if (map["MH:" + qk] != null) map["MH:" + qk] = after;
        touched = true;
      }
      if (touched) writeTonQ7MapToSheet_(ss, map, labels);
    } catch (eSync) {
      Logger.log("applyFifoStockOutBatch_ Q7 sync: " + (eSync.message || eSync));
    }

    clearInventoryRequestCache_();
    invalidateInventoryCaches_();
    return {
      success: true,
      changed: allAllocations.length,
      allocations: allAllocations,
      shortfalls: shortfalls,
      q7Deltas: q7Deltas
    };
  };

  if (opts.acquireLock) return withInventoryLock_(run);
  return run();
}


/**
 * Từ confirmations nhận hàng → lines FIFO.
 */
function applyFifoStockOutFromConfirmations_(ss, historySheet, confirmations, soPhieu, rowCache) {
  var lines = [];
  for (var i = 0; i < (confirmations || []).length; i++) {
    var conf = confirmations[i];
    var row = Number(conf.row);
    var qty = Number(conf.receivedQty);
    if (!row || row < 2 || isNaN(qty) || qty <= 0) continue;
    var orderRow = (rowCache && rowCache[row]) ? rowCache[row] : null;
    if (!orderRow && historySheet) {
      try { orderRow = historySheet.getRange(row, 1, 1, 9).getValues()[0]; } catch (eR) { orderRow = null; }
    }
    if (!orderRow) continue;
    var rowSoPhieu = orderRow[1] ? String(orderRow[1]).trim() : "";
    if (soPhieu && rowSoPhieu && !orderKeysMatch_(rowSoPhieu, soPhieu)) continue;
    lines.push({
      maHang: orderRow[4],
      maVach: orderRow[5],
      dvt: orderRow[9] || "",
      qty: qty
    });
  }
  if (!lines.length) return { success: true, changed: 0 };
  return applyFifoStockOutBatch_(ss, lines, { acquireLock: false });
}


/** API: danh sách lô còn tồn theo mã */
function getStockLotsForItem(maHang) {
  try {
    var ss = getSS();
    seedInventoryRequestCache_(ss);
    var lots = getLotsForItemCode_(ss, maHang).filter(function(l) {
      return (Number(l.qtyRemain) || 0) > 0;
    });
    lots.sort(function(a, b) {
      return (a.ngayNhapMs || 0) - (b.ngayNhapMs || 0);
    });
    return { success: true, maHang: maHang, lots: lots, count: lots.length };
  } catch (e) {
    return { success: false, error: e.message || String(e), lots: [] };
  }
}
