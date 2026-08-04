var SHEET_ID = "1lrMxXon0oEtTUs6nsAydSAthGxryO3GNde0GQusk4j4";

// ================= THÔNG TIN CẤU HÌNH TELEGRAM & WEB APP =================
var TELEGRAM_TOKEN = "8918960838:AAE2w_tEGPD2E25fRz6LK5xUiXGGZGWv8NU"; 
var TELEGRAM_CHAT_ID = "-5408426667"; 
var WEB_APP_URL = "https://donhang-dieuchuyen.vercel.app";
// Mapping original store names -> display (short) names for UI and Telegram
var STORE_MAP = {
  "Kho Địa điểm kinh doanh Q7": "K9 Quận 7",
  "Kho Địa điểm kinh doanh 01": "K9 Quận 4 Mới",
  "Kho Địa điểm kinh doanh 02": "K9 Quận 8",
  "Kho Địa điểm kinh doanh 03": "K9 Phạm Hùng",
  "Kho Địa điểm kinh doanh 04": "K9 Quận 5",
  "Kho Địa điểm kinh doanh 05": "K9 Quận 1",
  "Kho Địa điểm kinh doanh 06": "K9 Quận 4 Cũ"
};
// Mã ngắn để phân biệt cửa hàng trùng nhãn (2× Q4) trên bảng tổng hợp soạn hàng
var STORE_SHORT_CODES = {
  "Kho Địa điểm kinh doanh 01": "178",
  "Kho Địa điểm kinh doanh 06": "275"
};
var GUIDE_SHEET_NAME = "Hướng dẫn";
var CACHE_STORES_KEY = "stores_registry_v1";
var CACHE_CATALOG_PREFIX = "catalog_data_v2_";
var CACHE_STOCK_INDEX_PREFIX = "stock_index_v1_";
var CACHE_NEW_PRODUCTS_PREFIX = "new_products_v1_";
var CACHE_TTL_SECONDS = 1800;
var HISTORY_MAX_ROWS_DEFAULT = 8000;
var NEW_PRODUCTS_DEFAULT_LIMIT = 10;
// Kho soạn hàng chính — sheet nhẹ chỉ chứa tồn Q7 (tạo lúc import file tồn)
var PACKING_STOCK_STORE = "Kho Địa điểm kinh doanh Q7";
var TON_Q7_SHEET_NAME = "TON_Q7";
var CACHE_TON_Q7_KEY = "ton_q7_map_v2";
// Tồn riêng theo biến thể đồ chơi — đối soát: Ton_Ban_Dau - Da_Xuat + Da_Nhan_Nhap = Ton_Hien_Tai
var TON_VARIANT_SHEET_NAME = "TON_VARIANT";
var CACHE_TON_VARIANT_KEY = "ton_variant_map_v2";
var TON_VARIANT_COL_COUNT = 10;
var TON_VARIANT_HEADERS = [
  "Key",            // A: MH_Con|DV
  "Parent_SKU",     // B
  "Ton_Ban_Dau",    // C
  "Da_Xuat",        // D
  "Ton_Hien_Tai",   // E = C - D + F
  "Da_Nhan_Nhap",   // F
  "TenSP_ChiTiet",  // G
  "MaVach",         // H
  "DonViTinh",      // I
  "UpdatedAt"       // J
];
var CATALOG_PARENT_HEADER = "Parent_SKU";
var CATALOG_ISNEW_HEADER = "IsNew";
var CATALOG_COL_COUNT = 11; // A..K: ... Parent_SKU | IsNew

function getScriptCache_() {
  return CacheService.getScriptCache();
}

function isPackingQ7Store_(storeName) {
  if (!storeName) return false;
  if (isStoreNameMatch(storeName, PACKING_STOCK_STORE)) return true;
  return formatShortStoreLabel(storeName) === "Q7";
}

function normalizeDvtKey_(dvt) {
  return normalizeHeaderText(String(dvt || "").trim());
}

function dvtFromStockKey_(key) {
  var text = String(key || "");
  var idx = text.indexOf("|DV:");
  if (idx === -1) return "";
  return text.substring(idx + 4);
}

function writeTonQ7MapToSheet_(ss, map, dvtLabelByKey) {
  ss = ss || getSS();
  var t0 = Date.now();
  map = map || {};
  dvtLabelByKey = dvtLabelByKey || {};
  var sh = ss.getSheetByName(TON_Q7_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(TON_Q7_SHEET_NAME);
  sh.clear();
  sh.getRange(1, 1, 1, 4).setValues([["Key", "Qty", "Dvt", "UpdatedAt"]]);
  var rows = [];
  for (var k in map) {
    if (!Object.prototype.hasOwnProperty.call(map, k)) continue;
    if (k === "__meta") continue;
    var dvtLabel = dvtLabelByKey[k] || dvtFromStockKey_(k) || "";
    rows.push([k, Number(map[k]) || 0, dvtLabel, ""]);
  }
  // Sheet.getRange(row, column, numRows, numColumns)
  if (rows.length) sh.getRange(2, 1, rows.length, 4).setValues(rows);
  sh.getRange(1, 4).setValue(new Date());
  try { SpreadsheetApp.flush(); } catch (e) {}
  try {
    var cache = getScriptCache_();
    putCacheJson_(cache, CACHE_TON_Q7_KEY, map, CACHE_TTL_SECONDS);
    putCacheJson_(cache, CACHE_TON_Q7_KEY + "_dvt", dvtLabelByKey, CACHE_TTL_SECONDS);
  } catch (e2) {}
  return {
    success: true,
    sheetName: TON_Q7_SHEET_NAME,
    rows: rows.length,
    keyCount: rows.length,
    store: PACKING_STOCK_STORE,
    ms: Date.now() - t0,
    _debugRun: "q7-v3"
  };
}

/** Ghi TON_Q7 từ entries {k,q,d} — giữ ĐVT gốc để hiển thị */
function writeTonQ7EntriesToSheet_(ss, entries) {
  var map = {};
  var dvtLabels = {};
  for (var i = 0; i < (entries || []).length; i++) {
    var ent = entries[i];
    if (!ent || !ent.k) continue;
    var key = String(ent.k).trim();
    if (!key) continue;
    map[key] = (Number(map[key]) || 0) + (Number(ent.q) || 0);
    var dLabel = String(ent.d || "").trim();
    if (dLabel) dvtLabels[key] = dLabel;
    else if (!dvtLabels[key]) dvtLabels[key] = dvtFromStockKey_(key);
  }
  return writeTonQ7MapToSheet_(ss, map, dvtLabels);
}

/** Key chuẩn TON_VARIANT: MH_Con|DV (không dấu cách) */
function buildTonVariantKey_(maHang, dvt) {
  var mh = normalizeProductCode(maHang);
  if (!mh) return "";
  var dv = normalizeDvtKey_(dvt);
  return mh + (dv ? ("|" + dv) : "");
}

function calcTonHienTaiVariant_(tonBanDau, daXuat, daNhanNhap) {
  return (Number(tonBanDau) || 0) - (Number(daXuat) || 0) + (Number(daNhanNhap) || 0);
}

function getOrCreateTonVariantSheet_(ss) {
  ss = ss || getSS();
  var sh = ss.getSheetByName(TON_VARIANT_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(TON_VARIANT_SHEET_NAME);
    sh.getRange(1, 1, 1, TON_VARIANT_COL_COUNT).setValues([TON_VARIANT_HEADERS]);
    sh.getRange(1, 1, 1, TON_VARIANT_COL_COUNT).setFontWeight("bold").setBackground("#cfe2f3");
    return sh;
  }
  ensureTonVariantSchema_(sh);
  return sh;
}

/** Migrate sheet cũ (Key|Qty|Parent...) → schema đối soát 10 cột */
function ensureTonVariantSchema_(sh) {
  if (!sh) return;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var header = sh.getRange(1, 1, 1, Math.max(lastCol, TON_VARIANT_COL_COUNT)).getValues()[0] || [];
  var h0 = normalizeHeaderText(header[0]);
  var h1 = normalizeHeaderText(header[1]);
  var h2 = normalizeHeaderText(header[2]);
  var alreadyNew = h1.indexOf("parent") !== -1 && (h2.indexOf("tonbandau") !== -1 || h2.indexOf("bandau") !== -1);
  if (alreadyNew) {
    // Đảm bảo header chuẩn
    sh.getRange(1, 1, 1, TON_VARIANT_COL_COUNT).setValues([TON_VARIANT_HEADERS])
      .setFontWeight("bold").setBackground("#cfe2f3");
    return;
  }
  // Schema cũ: Key | Qty | Parent_SKU | TenSP | MaVach | DonViTinh | UpdatedAt
  var isOld = (h1 === "qty" || h1 === "soluong" || h1.indexOf("ton") !== -1) && h0 === "key";
  var lastRow = sh.getLastRow();
  var migrated = [];
  if (isOld && lastRow >= 2) {
    var old = sh.getRange(2, 1, lastRow - 1, Math.min(lastCol, 7)).getValues();
    for (var i = 0; i < old.length; i++) {
      var k = String(old[i][0] || "").trim();
      if (!k) continue;
      // Chuẩn hoá key cũ MH:xxx|DV:yyy → xxx|yyy
      var keyNorm = k;
      if (keyNorm.indexOf("MH:") === 0) {
        keyNorm = keyNorm.substring(3).replace(/\|DV:/i, "|");
      }
      var qty = Number(old[i][1]) || 0;
      migrated.push([
        keyNorm,
        String(old[i][2] || "").trim(),
        qty,
        0,
        qty,
        0,
        String(old[i][3] || "").trim(),
        String(old[i][4] || "").trim(),
        String(old[i][5] || "").trim(),
        old[i][6] || new Date()
      ]);
    }
  }
  sh.clear();
  sh.getRange(1, 1, 1, TON_VARIANT_COL_COUNT).setValues([TON_VARIANT_HEADERS])
    .setFontWeight("bold").setBackground("#cfe2f3");
  if (migrated.length) {
    sh.getRange(2, 1, migrated.length, TON_VARIANT_COL_COUNT).setValues(migrated);
  }
  try { SpreadsheetApp.flush(); } catch (e) {}
  try { getScriptCache_().remove(CACHE_TON_VARIANT_KEY); } catch (e2) {}
}

/**
 * Import tồn biến thể: ghi Ton_Ban_Dau, reset Da_Xuat=0, Ton_Hien_Tai = Ton_Ban_Dau (+ Da_Nhan_Nhap nếu giữ).
 * entries: {k|maHang, q|qty, p|parentSku, th|tenHang, mv|maVach, d|dvt}
 * Batch: 1 getValues + 1 setValues.
 */
function writeTonVariantEntriesToSheet_(ss, entries) {
  ss = ss || getSS();
  var t0 = Date.now();
  var lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    var sh = getOrCreateTonVariantSheet_(ss);
    var byKey = {};
    var lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      var existing = sh.getRange(2, 1, lastRow - 1, TON_VARIANT_COL_COUNT).getValues();
      for (var r = 0; r < existing.length; r++) {
        var ek = String(existing[r][0] || "").trim();
        if (!ek) continue;
        byKey[ek] = {
          k: ek,
          p: String(existing[r][1] || "").trim(),
          tonBanDau: Number(existing[r][2]) || 0,
          daXuat: Number(existing[r][3]) || 0,
          tonHienTai: Number(existing[r][4]) || 0,
          daNhanNhap: Number(existing[r][5]) || 0,
          th: String(existing[r][6] || "").trim(),
          mv: String(existing[r][7] || "").trim(),
          d: String(existing[r][8] || "").trim()
        };
      }
    }

    for (var i = 0; i < (entries || []).length; i++) {
      var ent = entries[i];
      if (!ent) continue;
      var key = String(ent.k || "").trim();
      if (key.indexOf("MH:") === 0) key = key.substring(3).replace(/\|DV:/i, "|");
      if (!key) {
        key = buildTonVariantKey_(ent.maHang || ent.mh || "", ent.d || ent.dvt || "");
      }
      if (!key) continue;
      var qty = Number(ent.q != null ? ent.q : ent.qty);
      if (isNaN(qty)) qty = 0;
      var prev = byKey[key] || {
        k: key, p: "", tonBanDau: 0, daXuat: 0, tonHienTai: 0, daNhanNhap: 0, th: "", mv: "", d: ""
      };
      // Import mới: Ton_Ban_Dau = qty, reset Da_Xuat, Ton_Hien_Tai = Ton_Ban_Dau (+ nhập bổ sung nếu có)
      prev.tonBanDau = qty;
      prev.daXuat = 0;
      if (ent.keepNhap === true) {
        /* giữ Da_Nhan_Nhap */
      } else {
        prev.daNhanNhap = Number(ent.daNhanNhap != null ? ent.daNhanNhap : 0) || 0;
      }
      prev.tonHienTai = calcTonHienTaiVariant_(prev.tonBanDau, prev.daXuat, prev.daNhanNhap);
      var p = String(ent.p || ent.parentSku || "").trim();
      var th = String(ent.th || ent.tenHang || ent.tenSP || "").trim();
      var mv = String(ent.mv || ent.maVach || "").trim();
      var d = String(ent.d || ent.dvt || "").trim();
      if (p) prev.p = p;
      if (th) prev.th = th;
      if (mv) prev.mv = mv;
      if (d) prev.d = d;
      byKey[key] = prev;
    }

    var rows = [];
    var now = new Date();
    for (var k in byKey) {
      if (!Object.prototype.hasOwnProperty.call(byKey, k)) continue;
      var row = byKey[k];
      row.tonHienTai = calcTonHienTaiVariant_(row.tonBanDau, row.daXuat, row.daNhanNhap);
      rows.push([
        row.k,
        row.p || "",
        Number(row.tonBanDau) || 0,
        Number(row.daXuat) || 0,
        Number(row.tonHienTai) || 0,
        Number(row.daNhanNhap) || 0,
        row.th || "",
        row.mv || "",
        row.d || "",
        now
      ]);
    }
    rows.sort(function(a, b) {
      var ap = String(a[1] || "");
      var bp = String(b[1] || "");
      if (ap !== bp) return ap < bp ? -1 : 1;
      return String(a[0] || "") < String(b[0] || "") ? -1 : 1;
    });

    sh.clear();
    sh.getRange(1, 1, 1, TON_VARIANT_COL_COUNT).setValues([TON_VARIANT_HEADERS])
      .setFontWeight("bold").setBackground("#cfe2f3");
    if (rows.length) sh.getRange(2, 1, rows.length, TON_VARIANT_COL_COUNT).setValues(rows);
    try { SpreadsheetApp.flush(); } catch (e2) {}

    var map = buildTonVariantStockMapFromRows_(rows);
    try { putCacheJson_(getScriptCache_(), CACHE_TON_VARIANT_KEY, map, CACHE_TTL_SECONDS); } catch (e3) {}

    // Đồng bộ tổng Parent → TON_Q7
    try { syncParentVariantTotalsToTonQ7_(ss, rows); } catch (eSync) { Logger.log(eSync); }

    return {
      success: true,
      sheetName: TON_VARIANT_SHEET_NAME,
      rows: rows.length,
      keyCount: rows.length,
      ms: Date.now() - t0,
      _debugRun: "ton-variant-v2-import"
    };
  } finally {
    try { lock.releaseLock(); } catch (eL) {}
  }
}

/** Map lookup (MH:/MV:) từ Ton_Hien_Tai để UI/API dùng getStockValueForItem */
function buildTonVariantStockMapFromRows_(rows) {
  var map = {};
  for (var i = 0; i < (rows || []).length; i++) {
    var r = rows[i];
    if (!r) continue;
    var key = String(r[0] || "").trim();
    var ton = Number(r[4]) || 0; // Ton_Hien_Tai
    var mv = String(r[7] || "").trim();
    var dvt = String(r[8] || "").trim();
    if (!key) continue;
    map[key] = ton;
    var mh = key.split("|")[0];
    var dv = key.indexOf("|") !== -1 ? key.split("|").slice(1).join("|") : normalizeDvtKey_(dvt);
    if (mh) {
      addStockValueByCode(map, "MH:", mh, ton, dv || dvt);
    }
    if (mv) {
      addStockValueByCode(map, "MV:", normalizeProductCode(mv), ton, dv || dvt);
    }
  }
  return map;
}

function readTonVariantMap_(ss) {
  ss = ss || getSS();
  var cache = getScriptCache_();
  var cached = getCacheJson_(cache, CACHE_TON_VARIANT_KEY);
  if (cached && typeof cached === "object" && Object.keys(cached).length) return cached;

  var sh = ss.getSheetByName(TON_VARIANT_SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) return {};
  ensureTonVariantSchema_(sh);
  // getRange(row, column, numRows, numColumns) — đọc Key..DonViTinh
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, TON_VARIANT_COL_COUNT).getValues();
  var map = buildTonVariantStockMapFromRows_(data);
  try { putCacheJson_(cache, CACHE_TON_VARIANT_KEY, map, CACHE_TTL_SECONDS); } catch (e) {}
  return map;
}

/**
 * Khi soạn/xuất: cộng Da_Xuat, tính lại Ton_Hien_Tai, sync tổng Parent → TON_Q7.
 * lines: [{maHang, maVach, dvt, qty, tenHang, parentSku}]
 * Gọi trong LockService đã giữ (batch B).
 */
function applyTonVariantExportBatch_(ss, lines) {
  ss = ss || getSS();
  if (!lines || !lines.length) return { success: true, changed: 0 };
  var sh = getOrCreateTonVariantSheet_(ss);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { success: true, changed: 0, skipped: lines.length };

  var values = sh.getRange(2, 1, lastRow - 1, TON_VARIANT_COL_COUNT).getValues();
  var byKey = {};
  var byMv = {};
  for (var r = 0; r < values.length; r++) {
    var k0 = String(values[r][0] || "").trim();
    if (!k0) continue;
    byKey[k0] = r;
    var mv0 = normalizeProductCode(values[r][7]);
    if (mv0 && byMv[mv0] === undefined) byMv[mv0] = r;
  }

  // Resolve parent từ catalog nếu thiếu
  var catalogLookup = null;
  try { catalogLookup = getCatalogLookup(ss); } catch (eC) { catalogLookup = null; }

  var changed = 0;
  var now = new Date();
  var touchedParents = {};
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    if (!line) continue;
    var qty = Number(line.qty);
    if (!qty || qty <= 0 || isNaN(qty)) continue;
    var key = buildTonVariantKey_(line.maHang, line.dvt);
    var idx = key && byKey[key] !== undefined ? byKey[key] : -1;
    if (idx < 0) {
      var mvKey = normalizeProductCode(line.maVach);
      if (mvKey && byMv[mvKey] !== undefined) idx = byMv[mvKey];
    }
    if (idx < 0) continue; // không phải mã biến thể trong TON_VARIANT

    var daXuat = (Number(values[idx][3]) || 0) + qty;
    var tonBanDau = Number(values[idx][2]) || 0;
    var daNhan = Number(values[idx][5]) || 0;
    values[idx][3] = daXuat;
    values[idx][4] = calcTonHienTaiVariant_(tonBanDau, daXuat, daNhan);
    values[idx][9] = now;
    if (line.tenHang && !values[idx][6]) values[idx][6] = String(line.tenHang).trim();
    if (line.maVach && !values[idx][7]) values[idx][7] = String(line.maVach).trim();
    if (line.dvt && !values[idx][8]) values[idx][8] = String(line.dvt).trim();

    var parent = String(values[idx][1] || line.parentSku || "").trim();
    if (!parent && catalogLookup) {
      var cat = resolveCatalogProduct(catalogLookup, line.maHang, line.maVach);
      if (cat && cat.parentSku) {
        parent = String(cat.parentSku).trim();
        values[idx][1] = parent;
      }
    }
    if (parent) touchedParents[parent.toUpperCase()] = true;
    changed++;
  }

  if (changed) {
    sh.getRange(2, 1, values.length, TON_VARIANT_COL_COUNT).setValues(values);
    try { SpreadsheetApp.flush(); } catch (eF) {}
    var map = buildTonVariantStockMapFromRows_(values);
    try { putCacheJson_(getScriptCache_(), CACHE_TON_VARIANT_KEY, map, CACHE_TTL_SECONDS); } catch (eM) {}
    try { syncParentVariantTotalsToTonQ7_(ss, values, touchedParents); } catch (eS) { Logger.log(eS); }
  }
  return { success: true, changed: changed, parents: Object.keys(touchedParents).length };
}

/**
 * Từ entries tồn Q7/file: lọc mã có Parent_SKU trên Data_Excel → ghi TON_VARIANT (import reset).
 * entries dạng {k, q, d} hoặc {maHang, qty, dvt, ...}
 */
function importTonVariantFromStockEntries_(ss, entries) {
  ss = ss || getSS();
  var catalogLookup = null;
  try { catalogLookup = getCatalogLookup(ss); } catch (e) { catalogLookup = null; }
  if (!catalogLookup) return { rows: 0 };

  var variantEntries = [];
  for (var i = 0; i < (entries || []).length; i++) {
    var ent = entries[i];
    if (!ent) continue;
    var mh = "";
    var mv = "";
    var dvt = String(ent.d || ent.dvt || "").trim();
    var qty = Number(ent.q != null ? ent.q : ent.qty);
    if (isNaN(qty)) qty = 0;

    var key = String(ent.k || "").trim();
    if (key.indexOf("MH:") === 0) {
      var rest = key.substring(3);
      var parts = rest.split("|DV:");
      mh = parts[0] || "";
      if (parts[1] && !dvt) dvt = parts[1];
    } else if (key.indexOf("MV:") === 0) {
      var restMv = key.substring(3);
      var partsMv = restMv.split("|DV:");
      mv = partsMv[0] || "";
      if (partsMv[1] && !dvt) dvt = partsMv[1];
    } else {
      mh = String(ent.maHang || ent.mh || "").trim();
      mv = String(ent.maVach || ent.mv || "").trim();
      if (!mh && key && key.indexOf("|") !== -1) mh = key.split("|")[0];
      else if (!mh && key) mh = key;
    }

    var cat = resolveCatalogProduct(catalogLookup, mh, mv);
    if (!cat || !cat.parentSku) continue;
    var childMh = cat.maHang || mh;
    if (!childMh) continue;
    variantEntries.push({
      maHang: childMh,
      k: buildTonVariantKey_(childMh, dvt || cat.dvt || ""),
      q: qty,
      p: cat.parentSku,
      th: cat.tenHang || ent.th || "",
      mv: cat.maVach || mv || "",
      d: dvt || cat.dvt || ""
    });
  }
  if (!variantEntries.length) return { rows: 0 };
  var written = writeTonVariantEntriesToSheet_(ss, variantEntries);
  return { rows: written.rows || variantEntries.length, ms: written.ms || 0 };
}

/** Tổng Ton_Hien_Tai theo Parent_SKU → ghi vào TON_Q7 (MH:Parent) */
function syncParentVariantTotalsToTonQ7_(ss, rows, onlyParents) {
  ss = ss || getSS();
  var parentSum = {};
  for (var i = 0; i < (rows || []).length; i++) {
    var p = String(rows[i][1] || "").trim().toUpperCase();
    if (!p) continue;
    if (onlyParents && !onlyParents[p]) continue;
    parentSum[p] = (parentSum[p] || 0) + (Number(rows[i][4]) || 0);
  }
  if (!Object.keys(parentSum).length) return { updated: 0 };

  var bundle = readTonKhoQ7Bundle_(ss);
  var map = bundle.map || {};
  var labels = bundle.dvtLabels || {};
  var updated = 0;
  for (var parent in parentSum) {
    if (!Object.prototype.hasOwnProperty.call(parentSum, parent)) continue;
    var mh = normalizeProductCode(parent);
    if (!mh) continue;
    var q7Key = "MH:" + mh;
    map[q7Key] = Number(parentSum[parent]) || 0;
    if (!labels[q7Key]) labels[q7Key] = "";
    updated++;
  }
  writeTonQ7MapToSheet_(ss, map, labels);
  return { updated: updated };
}

/** Trả stock variant (Ton_Hien_Tai) nếu mã có trong TON_VARIANT; null nếu không */
function getVariantStockIfPresent_(map, maHang, maVach, dvt) {
  if (!map || !Object.keys(map).length) return null;
  var mh = normalizeProductCode(maHang);
  var mv = normalizeProductCode(maVach);
  var rawKey = buildTonVariantKey_(maHang, dvt);
  if (rawKey && Object.prototype.hasOwnProperty.call(map, rawKey)) return Number(map[rawKey]) || 0;
  var present = false;
  for (var k in map) {
    if (!Object.prototype.hasOwnProperty.call(map, k) || k === "__meta") continue;
    if (rawKey && k === rawKey) { present = true; break; }
    if (mh && (k === ("MH:" + mh) || k.indexOf("MH:" + mh + "|DV:") === 0 || k === mh || k.indexOf(mh + "|") === 0)) {
      present = true; break;
    }
    if (mv && (k === ("MV:" + mv) || k.indexOf("MV:" + mv + "|DV:") === 0)) { present = true; break; }
  }
  if (!present) return null;
  return getStockValueForItem(map, maHang, maVach, dvt);
}

/** Tên hiển thị biến thể cho soạn/in/excel: `MH - Tên chi tiết` */
function formatVariantDisplayName_(maHang, tenHang) {
  var mh = String(maHang || "").trim();
  var th = String(tenHang || "").trim();
  if (!mh) return th;
  if (!th) return mh;
  if (th.toUpperCase().indexOf(mh.toUpperCase()) === 0) return th;
  return mh + " - " + th;
}

/**
 * API: danh sách biến thể theo Parent_SKU kèm tồn từ TON_VARIANT.
 * @param {string} parentSku
 */
function getVariantStockList(parentSku) {
  try {
    var parent = String(parentSku || "").trim().toUpperCase();
    if (!parent) return { success: false, error: "Thiếu parentSku", parentSku: "", variants: [] };

    var ss = getSS();
    var catalogRes = getCatalogData();
    var danhMuc = (catalogRes && catalogRes.success && catalogRes.danhMuc) ? catalogRes.danhMuc : buildCatalogFromSheet_(ss);
    var seen = {};
    var variants = [];
    for (var key in danhMuc) {
      if (!Object.prototype.hasOwnProperty.call(danhMuc, key)) continue;
      var item = danhMuc[key];
      if (!item) continue;
      var p = String(item.parentSku || "").trim().toUpperCase();
      if (p !== parent) continue;
      var dedupe = String(item.maHang || "").trim().toUpperCase() + "|" + String(item.maVach || "").trim().toUpperCase();
      if (!dedupe || seen[dedupe]) continue;
      seen[dedupe] = true;
      variants.push({
        maHang: item.maHang || "",
        maVach: item.maVach || "",
        tenHang: item.tenHang || "",
        dvt: item.dvt || "",
        parentSku: p,
        stock: 0
      });
    }

    var map = readTonVariantMap_(ss) || {};
    for (var i = 0; i < variants.length; i++) {
      variants[i].stock = getStockValueForItem(map, variants[i].maHang, variants[i].maVach, variants[i].dvt);
    }
    variants.sort(function(a, b) {
      return String(a.tenHang || "").localeCompare(String(b.tenHang || "")) ||
        String(a.maHang || "").localeCompare(String(b.maHang || ""));
    });

    return {
      success: true,
      parentSku: parent,
      variants: variants,
      count: variants.length,
      stockReady: Object.keys(map).length > 0
    };
  } catch (e) {
    return { success: false, error: e.message || String(e), parentSku: String(parentSku || ""), variants: [] };
  }
}

function rebuildTonKhoQ7Sheet_(ss) {
  ss = ss || getSS();
  var map = getStockMapForStoreFromFullSheet_(ss, PACKING_STOCK_STORE);
  return writeTonQ7MapToSheet_(ss, map);
}

function mergeTonQ7FromMatrix_(ss, matrix, reset) {
  var map = reset ? {} : (readTonKhoQ7Map_(ss) || {});
  var chunkMap = buildStockMapForStoreFromData_(matrix, PACKING_STOCK_STORE);
  for (var k in chunkMap) {
    if (!Object.prototype.hasOwnProperty.call(chunkMap, k) || k === "__meta") continue;
    map[k] = (Number(map[k]) || 0) + (Number(chunkMap[k]) || 0);
  }
  return writeTonQ7MapToSheet_(ss, map);
}

function readTonKhoQ7Bundle_(ss) {
  ss = ss || getSS();
  var cache = getScriptCache_();
  var cached = getCacheJson_(cache, CACHE_TON_Q7_KEY);
  var labelCached = getCacheJson_(cache, CACHE_TON_Q7_KEY + "_dvt");
  if (cached && typeof cached === "object" && Object.keys(cached).length) {
    return { map: cached, dvtLabels: labelCached || {} };
  }

  var sh = ss.getSheetByName(TON_Q7_SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) return { map: null, dvtLabels: {} };
  var lastRow = sh.getLastRow();
  var numRows = lastRow - 1;
  var lastCol = Math.max(sh.getLastColumn(), 3);
  var data = sh.getRange(2, 1, numRows, Math.min(lastCol, 3)).getValues();
  var map = {};
  var dvtLabels = {};
  for (var i = 0; i < data.length; i++) {
    var key = String(data[i][0] || "").trim();
    if (!key) continue;
    var qty = Number(data[i][1]) || 0;
    var dvtCol = data[i][2] != null ? String(data[i][2]).trim() : "";
    // Sheet cũ: Key không có |DV: nhưng cột Dvt có giá trị → gắn vào key
    if (dvtCol && key.indexOf("|DV:") === -1) {
      var dvtNorm = normalizeDvtKey_(dvtCol);
      if (dvtNorm) key = key + "|DV:" + dvtNorm;
    }
    map[key] = (Number(map[key]) || 0) + qty;
    if (dvtCol) dvtLabels[key] = dvtCol;
    else if (!dvtLabels[key]) dvtLabels[key] = dvtFromStockKey_(key);
  }
  if (!Object.keys(map).length) return { map: null, dvtLabels: {} };
  try {
    putCacheJson_(cache, CACHE_TON_Q7_KEY, map, CACHE_TTL_SECONDS);
    putCacheJson_(cache, CACHE_TON_Q7_KEY + "_dvt", dvtLabels, CACHE_TTL_SECONDS);
  } catch (e) {}
  return { map: map, dvtLabels: dvtLabels };
}

function readTonKhoQ7Map_(ss) {
  var bundle = readTonKhoQ7Bundle_(ss);
  return bundle.map;
}

/** Parse matrix tồn kho (từ file import hoặc sheet) → map mã+ĐVT cho 1 kho */
function buildStockMapForStoreFromData_(tkData, storeName) {
  var tonKhoMap = {};
  if (!storeName || !tkData || !tkData.length) return tonKhoMap;
  var stockConfig = getStockSheetConfig(tkData);
  var header = tkData[stockConfig.headerIndex] || [];
  var currentMaHang = "";
  var currentMaVach = "";
  var currentDvt = "";
  for (var k = stockConfig.startRow; k < tkData.length; k++) {
    var row = tkData[k];
    if (!row) continue;
    var rowMaHangRaw = getCellValue(row, stockConfig.maHangIdx, "");
    var rowMaVachRaw = getCellValue(row, stockConfig.maVachIdx, "");
    var rowDvtRaw = stockConfig.dvtIdx >= 0 ? getCellValue(row, stockConfig.dvtIdx, "") : "";
    var hasOwnCode = !!(rowMaHangRaw || rowMaVachRaw);
    if (hasOwnCode) {
      currentMaHang = rowMaHangRaw;
      currentMaVach = rowMaVachRaw;
      if (rowDvtRaw) currentDvt = rowDvtRaw;
    }
    var maHangTon = (hasOwnCode ? rowMaHangRaw : currentMaHang) || "";
    var maVachTon = (hasOwnCode ? rowMaVachRaw : currentMaVach) || "";
    var dvtTon = (rowDvtRaw || currentDvt || "");
    if (!maHangTon && !maVachTon) continue;

    var rowStores = getRowStoreNames(row, stockConfig);
    var match = false;
    if (stockConfig.storeHeaderIndexes && stockConfig.storeHeaderIndexes.length) {
      for (var c = 0; c < stockConfig.storeHeaderIndexes.length; c++) {
        var storeHeaderIdx = stockConfig.storeHeaderIndexes[c];
        var storeNameCandidate = getCellValue(header, storeHeaderIdx, "");
        var qty = parseQuantityValue(row[storeHeaderIdx]);
        if (!storeNameCandidate) continue;
        if (isStoreNameMatch(storeNameCandidate, storeName) && qty !== 0) {
          match = true;
          if (maHangTon) addStockValueByCode(tonKhoMap, "MH:", maHangTon, qty, dvtTon);
          if (maVachTon) addStockValueByCode(tonKhoMap, "MV:", maVachTon, qty, dvtTon);
        }
      }
    }
    if (match) continue;

    for (var s = 0; s < rowStores.length; s++) {
      if (isStoreNameMatch(rowStores[s], storeName)) {
        match = true;
        break;
      }
    }
    if (!match) continue;
    var ton = parseQuantityValue(row[stockConfig.tonKhoIdx]);
    if (maHangTon) addStockValueByCode(tonKhoMap, "MH:", maHangTon, ton, dvtTon);
    if (maVachTon) addStockValueByCode(tonKhoMap, "MV:", maVachTon, ton, dvtTon);
  }
  return tonKhoMap;
}

/** Đọc full sheet tổng hợp — chỉ khi rebuild thủ công */
function getStockMapForStoreFromFullSheet_(ss, storeName) {
  if (!storeName) return {};
  var tonKhoSheet = ss.getSheetByName("TỔNG HỢP TỒN KHO");
  if (!tonKhoSheet) return {};
  return buildStockMapForStoreFromData_(tonKhoSheet.getDataRange().getValues(), storeName);
}

function invalidateCatalogCache_() {
  // Catalog cache keys include sheet version suffix; bump is handled on next read.
  try {
    var props = PropertiesService.getScriptProperties();
    props.setProperty("catalog_version_bump", String(Date.now()));
  } catch (e) {}
}

function invalidateStoresCache_() {
  getScriptCache_().remove(CACHE_STORES_KEY);
}

function readHistoryDataPack_(historySheet, maxRows) {
  if (!historySheet) return { data: [[]], startRow: 2 };
  var lastRow = historySheet.getLastRow();
  var lastCol = Math.max(historySheet.getLastColumn(), 16);
  if (lastRow < 2) return { data: [[]], startRow: 2 };

  var limit = maxRows || HISTORY_MAX_ROWS_DEFAULT;
  var startRow = 2;
  if (lastRow - 1 > limit) startRow = lastRow - limit + 1;
  var numRows = lastRow - startRow + 1;
  var body = historySheet.getRange(startRow, 1, numRows, lastCol).getValues();
  var data = [[]];
  for (var i = 0; i < body.length; i++) data.push(body[i]);
  return { data: data, startRow: startRow };
}

function getCacheJson_(cache, key) {
  if (!cache) return null;
  var raw = cache.get(key);
  if (raw) {
    try { return JSON.parse(raw); } catch (e) {}
  }
  var n = cache.get(key + "_n");
  if (!n) return null;
  var parts = [];
  var num = Number(n) || 0;
  for (var i = 0; i < num; i++) {
    var part = cache.get(key + "_" + i);
    if (part === null) return null;
    parts.push(part);
  }
  try { return JSON.parse(parts.join("")); } catch (e) { return null; }
}

function putCacheJson_(cache, key, obj, ttl) {
  if (!cache) return;
  var json = JSON.stringify(obj);
  if (json.length < 95000) {
    cache.put(key, json, ttl);
    return;
  }
  var chunkSize = 90000;
  var numChunks = Math.ceil(json.length / chunkSize);
  cache.put(key + "_n", String(numChunks), ttl);
  for (var ci = 0; ci < numChunks; ci++) {
    cache.put(key + "_" + ci, json.substring(ci * chunkSize, (ci + 1) * chunkSize), ttl);
  }
}

function readHistoryForSelectedOrders_(historySheet, selectedSet, baseDateStr, maxScanRows) {
  // Quét từ dưới lên theo chunk; khi đã thấy đủ số phiếu thì đọc thêm 1 chunk rồi dừng.
  maxScanRows = maxScanRows || 2500;
  var chunkSize = 600;
  var lastRow = historySheet.getLastRow();
  var lastCol = Math.min(Math.max(historySheet.getLastColumn(), 13), 16);
  if (lastRow < 2) return { data: [[]], startRow: 2, scannedRows: 0, matchedRows: 0 };

  var foundOrders = {};
  var needLeft = (selectedSet && selectedSet._list && selectedSet._list.length)
    ? selectedSet._list.length
    : 0;
  if (!needLeft) {
    for (var nk in selectedSet) {
      if (!selectedSet.hasOwnProperty(nk) || nk === "_list") continue;
      needLeft++;
    }
  }
  var matchedRows = [];
  var scanned = 0;
  var endRow = lastRow;
  var extraChunkAfterFound = false;

  while (endRow >= 2 && scanned < maxScanRows) {
    var startRow = Math.max(2, endRow - chunkSize + 1);
    if (scanned + (endRow - startRow + 1) > maxScanRows) {
      startRow = Math.max(2, endRow - (maxScanRows - scanned) + 1);
    }
    var numRows = endRow - startRow + 1;
    var body = historySheet.getRange(startRow, 1, numRows, lastCol).getValues();
    for (var i = 0; i < body.length; i++) {
      var row = body[i];
      if (!row) continue;
      var soPhieu = row[1] ? String(row[1]).trim() : "";
      if (!soPhieu || !orderInMatchSet_(soPhieu, selectedSet)) continue;
      matchedRows.push({ row: row, order: startRow + i });
      if (!foundOrders[soPhieu]) {
        foundOrders[soPhieu] = true;
        needLeft--;
      }
    }
    scanned += numRows;
    endRow = startRow - 1;
    if (needLeft <= 0) {
      if (extraChunkAfterFound) break;
      extraChunkAfterFound = true;
    }
  }

  matchedRows.sort(function(a, b) { return a.order - b.order; });
  var data = [[]];
  var orders = [];
  for (var m = 0; m < matchedRows.length; m++) {
    data.push(matchedRows[m].row);
    orders.push(matchedRows[m].order);
  }
  return {
    data: data,
    orders: orders,
    startRow: matchedRows.length ? matchedRows[0].order : 2,
    scannedRows: scanned,
    matchedRows: matchedRows.length,
    foundOrders: Object.keys(foundOrders).length
  };
}

var STOCK_SHEET_CACHE_NAME = "__CACHE_STOCK_INDEX";

function readStockIndexFromSheet_(ss, version) {
  var sh = ss.getSheetByName(STOCK_SHEET_CACHE_NAME);
  if (!sh || sh.getLastRow() < 2) return null;
  var meta = sh.getRange(1, 1, 1, 2).getValues()[0];
  if (String(meta[0] || "") !== String(version)) return null;
  var lastRow = sh.getLastRow();
  var numRows = lastRow - 1;
  if (numRows < 1) return null;
  var data = sh.getRange(2, 1, numRows, 2).getValues();
  var index = {};
  for (var i = 0; i < data.length; i++) {
    var storeKey = String(data[i][0] || "").trim();
    var raw = data[i][1];
    if (!storeKey || raw === "" || raw === null) continue;
    try {
      index[storeKey] = typeof raw === "string" ? JSON.parse(raw) : raw;
    } catch (e) {}
  }
  return Object.keys(index).length ? index : null;
}

function writeStockIndexToSheet_(ss, version, index) {
  var sh = ss.getSheetByName(STOCK_SHEET_CACHE_NAME);
  if (!sh) sh = ss.insertSheet(STOCK_SHEET_CACHE_NAME);
  sh.clear();
  sh.getRange(1, 1, 1, 2).setValues([[version, new Date()]]);
  var storeKeys = Object.keys(index || {});
  if (!storeKeys.length) return false;
  var rows = [];
  for (var i = 0; i < storeKeys.length; i++) {
    var sk = storeKeys[i];
    var json = JSON.stringify(index[sk] || {});
    // Giới hạn cell ~50k; nếu quá dài thì bỏ qua map kho đó (hiếm)
    if (json.length > 48000) {
      var slim = {};
      var codes = Object.keys(index[sk] || {});
      for (var c = 0; c < codes.length && JSON.stringify(slim).length < 45000; c++) {
        slim[codes[c]] = index[sk][codes[c]];
      }
      json = JSON.stringify(slim);
    }
    rows.push([sk, json]);
  }
  if (rows.length) sh.getRange(2, 1, rows.length, 2).setValues(rows);
  try { SpreadsheetApp.flush(); } catch (e) {}
  return true;
}

function getStockIndexCached_(ss, options) {
  options = options || {};
  ss = ss || getSS();
  var stockSheet = ss.getSheetByName("TỔNG HỢP TỒN KHO");
  if (!stockSheet) return {};
  var version = String(stockSheet.getLastRow()) + "_" + String(stockSheet.getLastColumn());
  var cacheKey = CACHE_STOCK_INDEX_PREFIX + version;
  var cache = getScriptCache_();
  var cached = getCacheJson_(cache, cacheKey);
  if (cached) return cached;
  var sheetCached = readStockIndexFromSheet_(ss, version);
  if (sheetCached) {
    try { putCacheJson_(cache, cacheKey, sheetCached, CACHE_TTL_SECONDS); } catch (e) {}
    return sheetCached;
  }
  // onlyCached: không build trong request tạo bảng (tránh vượt timeout Vercel 60s)
  if (options.onlyCached) return null;
  var stockData = stockSheet.getDataRange().getValues();
  var index = getStockIndexByStore(stockData);
  try { putCacheJson_(cache, cacheKey, index, CACHE_TTL_SECONDS); } catch (e) {}
  try { writeStockIndexToSheet_(ss, version, index); } catch (e2) {}
  return index;
}

function extractStoreFullNameFromCell_(value) {
  var text = String(value || "").trim();
  if (!text) return "";
  var pipeIdx = text.indexOf("|");
  if (pipeIdx !== -1) return String(text.substring(pipeIdx + 1)).trim();
  return text;
}

function extractStoreCodeFromCell_(value) {
  var text = String(value || "").trim();
  if (!text) return "";
  var pipeIdx = text.indexOf("|");
  if (pipeIdx !== -1) return String(text.substring(0, pipeIdx)).trim();
  return text.split(/\s+/)[0] || text;
}

function findGuideSheetColumns_(headerRow) {
  var cols = { stt: -1, khoCode: -1, diaChi: -1, khoFull: -1, tenNgan: -1 };
  if (!headerRow) return cols;
  var khoCols = [];
  for (var c = 0; c < headerRow.length; c++) {
    var token = normalizeHeaderText(headerRow[c]);
    if (!token) continue;
    if (token === "stt") cols.stt = c;
    else if (token.indexOf("diachi") !== -1 || token.indexOf("dia chi") !== -1) cols.diaChi = c;
    else if (token.indexOf("tenngan") !== -1 || token.indexOf("ten ngan") !== -1) cols.tenNgan = c;
    else if (token.indexOf("kho") !== -1) khoCols.push(c);
  }
  if (khoCols.length >= 2) {
    cols.khoCode = khoCols[0];
    cols.khoFull = khoCols[1];
  } else if (khoCols.length === 1) {
    cols.khoFull = khoCols[0];
    cols.khoCode = khoCols[0];
  }
  return cols;
}

function buildFallbackStoreRegistry_() {
  var stores = Object.keys(STORE_MAP);
  return { stores: stores, storeMap: copyObject_(STORE_MAP), storeDetails: [] };
}

function copyObject_(obj) {
  var out = {};
  for (var key in obj) {
    if (Object.prototype.hasOwnProperty.call(obj, key)) out[key] = obj[key];
  }
  return out;
}

function loadStoresFromGuideSheet(ss) {
  ss = ss || getSS();
  var sheet = ss.getSheetByName(GUIDE_SHEET_NAME);
  if (!sheet) return buildFallbackStoreRegistry_();

  var data = sheet.getDataRange().getValues();
  if (!data || data.length < 2) return buildFallbackStoreRegistry_();

  var headerIndex = -1;
  var cols = null;
  for (var r = 0; r < Math.min(data.length, 20); r++) {
    var row = data[r];
    if (!row) continue;
    var probe = findGuideSheetColumns_(row);
    if (probe.khoFull !== -1 && probe.tenNgan !== -1) {
      headerIndex = r;
      cols = probe;
      break;
    }
  }
  if (headerIndex < 0 || !cols) {
    headerIndex = 0;
    cols = { stt: 0, khoCode: 1, diaChi: 2, khoFull: 3, tenNgan: 4 };
  }

  var stores = [];
  var storeMap = {};
  var storeDetails = [];
  for (var i = headerIndex + 1; i < data.length; i++) {
    var row = data[i];
    if (!row) continue;
    var fullName = getCellValue(row, cols.khoFull, "");
    if (!fullName) fullName = extractStoreFullNameFromCell_(getCellValue(row, cols.khoCode, ""));
    fullName = String(fullName).trim();
    if (!fullName) continue;

    var shortName = String(getCellValue(row, cols.tenNgan, "")).trim();
    var codeCell = getCellValue(row, cols.khoCode, "");
    var address = getCellValue(row, cols.diaChi, "");
    var sttVal = getCellValue(row, cols.stt, "");

    if (stores.indexOf(fullName) === -1) stores.push(fullName);
    storeMap[fullName] = shortName || fullName;
    storeDetails.push({
      stt: sttVal,
      code: extractStoreCodeFromCell_(codeCell),
      fullName: fullName,
      shortName: shortName || fullName,
      address: address
    });
  }

  if (!stores.length) return buildFallbackStoreRegistry_();
  return { stores: stores, storeMap: storeMap, storeDetails: storeDetails };
}

function getStoreRegistry(ss) {
  var cache = getScriptCache_();
  var cached = cache.get(CACHE_STORES_KEY);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }
  var registry = loadStoresFromGuideSheet(ss);
  try {
    cache.put(CACHE_STORES_KEY, JSON.stringify(registry), CACHE_TTL_SECONDS);
  } catch (e) {}
  return registry;
}

function getActiveStoreMap() {
  try {
    var registry = getStoreRegistry();
    var merged = copyObject_(STORE_MAP);
    if (registry && registry.storeMap) {
      for (var key in registry.storeMap) {
        if (Object.prototype.hasOwnProperty.call(registry.storeMap, key)) {
          merged[key] = registry.storeMap[key];
        }
      }
    }
    return merged;
  } catch (e) {
    return STORE_MAP;
  }
}

function getRuntimeStores() {
  try {
    var registry = getStoreRegistry();
    if (registry && registry.stores && registry.stores.length) return registry.stores;
  } catch (e) {}
  return Object.keys(STORE_MAP);
}

function getCatalogVersion_() {
  var ss = getSS();
  var sheet = ss.getSheetByName("Data_Excel");
  if (!sheet) return "0";
  var bump = "";
  try {
    bump = PropertiesService.getScriptProperties().getProperty("catalog_version_bump") || "";
  } catch (e) {}
  return String(sheet.getLastRow()) + "_" + String(sheet.getLastColumn()) + "_" + bump;
}

function buildCatalogFromSheet_(ss) {
  var danhMucHangHoa = {};
  ss = ss || getSS();
  var dataSheet = ss.getSheetByName("Data_Excel");
  if (!dataSheet) return danhMucHangHoa;

  var rawData = dataSheet.getDataRange().getValues();
  var tenHangChuanTheoMa = {};
  var headerRowIndex = -1;
  var headerRow = null;
  for (var hdr = 0; hdr < Math.min(rawData.length, 6); hdr++) {
    if (!rawData[hdr]) continue;
    var normalizedHeaderText = "";
    for (var c = 0; c < rawData[hdr].length; c++) {
      var cellText = normalizeHeaderText(rawData[hdr][c]);
      if (cellText.indexOf("mahang") !== -1 || cellText.indexOf("mavach") !== -1 || cellText.indexOf("tenhang") !== -1 || cellText.indexOf("dvt") !== -1 || cellText.indexOf("donvi") !== -1 || cellText.indexOf("unit") !== -1) {
        normalizedHeaderText = cellText;
        break;
      }
    }
    if (normalizedHeaderText) {
      headerRowIndex = hdr;
      headerRow = rawData[hdr];
      break;
    }
  }
  var parentSkuIdx = findCatalogParentColIdx_(headerRow);
  var maHangIdx = findCatalogMaHangColIdx_(headerRow, parentSkuIdx);
  if (maHangIdx === -1) maHangIdx = findColumnIndexByAliases(headerRow, ['mahang', 'sku', 'article', 'code']);
  var maVachIdx = findColumnIndexByAliases(headerRow, ['mavach', 'barcode', 'barcodeid']);
  var tenHangIdx = findColumnIndexByAliases(headerRow, ['tenhang', 'name', 'tênhang', 'description']);
  var dvtIdx = findColumnIndexByAliases(headerRow, ['dvt', 'donvitinh', 'donvi', 'unit', 'uom']);
  var isNewIdx = findColumnIndexByAliases(headerRow, ["isnew", "trangthaimoi", "hangmoi", "newflag"]);
  if (parentSkuIdx === -1 && headerRow && headerRow.length >= 10) {
    // Sheet chuẩn: cột J = Parent_SKU
    var hJ = normalizeHeaderText(headerRow[9]);
    if (!hJ || hJ.indexOf("parent") !== -1 || hJ.indexOf("nhom") !== -1 || hJ.indexOf("cha") !== -1) parentSkuIdx = 9;
  }
  if (isNewIdx === -1 && headerRow && headerRow.length >= 11) {
    var hK = normalizeHeaderText(headerRow[10]);
    if (!hK || hK.indexOf("isnew") !== -1 || hK.indexOf("moi") !== -1) isNewIdx = 10;
  }
  var startRow = headerRowIndex >= 0 ? headerRowIndex + 1 : 2;
  for (var k = startRow; k < rawData.length; k++) {
    if (!rawData[k]) continue;
    var ma = getCellValue(rawData[k], maHangIdx !== -1 ? maHangIdx : 0, "").toUpperCase();
    var ten = getCellValue(rawData[k], tenHangIdx !== -1 ? tenHangIdx : 5, "");
    if (ma !== "" && ten !== "") tenHangChuanTheoMa[ma] = ten;
  }
  for (var i = startRow; i < rawData.length; i++) {
    if (!rawData[i]) continue;
    var maHang = getCellValue(rawData[i], maHangIdx !== -1 ? maHangIdx : 0, "");
    var maVach = getCellValue(rawData[i], maVachIdx !== -1 ? maVachIdx : 2, "");
    var tenHang = getCellValue(rawData[i], tenHangIdx !== -1 ? tenHangIdx : 5, "");
    var dvt = getCellValue(rawData[i], dvtIdx !== -1 ? dvtIdx : 7, "");
    var parentSku = parentSkuIdx !== -1 ? getCellValue(rawData[i], parentSkuIdx, "") : "";
    var isNewFlag = false;
    if (isNewIdx !== -1) {
      var flagRaw = String(rawData[i][isNewIdx] == null ? "" : rawData[i][isNewIdx]).trim().toLowerCase();
      isNewFlag = flagRaw === "1" || flagRaw === "true" || flagRaw === "yes" || flagRaw === "x" || flagRaw === "moi" || flagRaw === "new";
    }
    if (dvt === "" && dvtIdx === -1) {
      var fallbackDvt = getCellValue(rawData[i], 6, "");
      if (fallbackDvt !== "") dvt = fallbackDvt;
    }
    if (tenHang === "" && maHang !== "") tenHang = tenHangChuanTheoMa[maHang.toUpperCase()] || "";
    var obj = { maHang: maHang, maVach: maVach, tenHang: tenHang, dvt: dvt || "", parentSku: parentSku || "", isNew: isNewFlag };
    if (maVach !== "") danhMucHangHoa[maVach.toUpperCase()] = obj;
    if (maHang !== "" && !danhMucHangHoa[maHang.toUpperCase()]) danhMucHangHoa[maHang.toUpperCase()] = obj;
  }
  return danhMucHangHoa;
}
// ========================================================================
function doPost(e) {
  try {
    var contents = e && e.postData && e.postData.contents ? e.postData.contents : "";
    if (!contents) {
      Logger.log("doPost warning: empty payload");
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: "Empty POST payload" })).setMimeType(ContentService.MimeType.JSON);
    }
    var payload = JSON.parse(contents);
    // If request is from webapp proxy (has action), route to server API handlers
    if (payload && payload.action) {
      var action = payload.action;
      try {
        var result = null;
        switch(action) {
          case 'loginUser':
            var credentials = payload.payload || {};
            result = loginUser(credentials.username || '', credentials.password || '');
            break;
          case 'luuPhieuTuWebApp':
            result = luuPhieuTuWebApp(payload.payload || {});
            break;
          case 'luuXuatBanHang':
            requireAuthenticatedAction(payload.payload || {});
            result = luuXuatBanHang(payload.payload || {});
            break;
          case 'luuChinhSuaPhieu':
            requireAdminAction(action, payload.payload || {});
            result = luuChinhSuaPhieu(payload.payload || {});
            break;
          case 'themChiTietPhieu':
            requireAuthenticatedAction(payload.payload || {});
            result = themChiTietPhieu(payload.payload || {});
            break;
          case 'huyDongChiTietPhieu':
            requireAuthenticatedAction(payload.payload || {});
            result = huyDongChiTietPhieu(payload.payload || {});
            break;
          case 'huyPhieu':
            requireAuthenticatedAction(payload.payload || {});
            result = huyPhieu(payload.payload || {});
            break;
          case 'taoTaiKhoanMoi':
            requireAdminAction(action, payload.payload || {});
            result = taoTaiKhoanMoi(payload.payload || {});
            break;
          case 'nhapKhauCapNhatThongTin':
            requireAdminAction(action, payload.payload || {});
            result = nhapKhauCapNhatThongTin(payload.payload || {});
            break;
          case 'doiMatKhau':
            result = doiMatKhau(payload.payload || {});
            break;
          case 'luuSoSoanHangVaAnh':
            result = luuSoSoanHangVaAnh(payload.payload || {});
            break;
          case 'taoFileExcelVaLayLink':
            result = taoFileExcelVaLayLink(payload.payload || {});
            break;
          case 'xacNhanNhanHang':
            requireAuthenticatedAction(payload.payload || {});
            result = xacNhanNhanHang(payload.payload || {});
            break;
          case 'taoBangSoanHangNgayMai':
            requireAuthenticatedAction(payload.payload || {});
            result = taoBangSoanHangNgayMai(payload.payload || {});
            break;
          case 'saveCatalogIsNewFlags':
            requireAdminAction(action, payload.payload || {});
            result = saveCatalogIsNewFlags_(payload.payload || {});
            break;
          case 'postProcessNewOrder':
            result = postProcessNewOrder(payload.payload || {});
            break;
          case 'postProcessPackingOrder':
            result = postProcessPackingOrder(payload.payload || {});
            break;
          case 'postProcessReceiveOrder':
            result = postProcessReceiveOrder(payload.payload || {});
            break;
          default:
            result = { error: 'Unknown action: ' + action };
        }
        return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
      } catch(apiErr) {
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          error: apiErr.message || String(apiErr),
          msg: apiErr.message || String(apiErr),
          action: action,
          _debugRun: "post-fix-v3-catch"
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    var update = payload;
    if (update.message) {
      handleTelegramMessage(update.message);
    } else if (update.edited_message) {
      handleTelegramMessage(update.edited_message);
    } else if (update.channel_post) {
      handleTelegramMessage(update.channel_post);
    } else if (update.callback_query && update.callback_query.message) {
      handleTelegramMessage(update.callback_query.message);
    }
  } catch (err) {
    Logger.log("doPost error: " + err);
  }
  return ContentService.createTextOutput("OK");
}
function handleTelegramMessage(message) {
  var chatId = message.chat.id;
  var username = message.from.username || "";
  var firstName = message.from.first_name || "";
  var text = (message.text || "").trim();
  var normalizedText = text.replace(/^\/(\S+?)(?:@\S+)?(\s|$)/i, function(match, cmd, sep) {
    return "/" + cmd + sep;
  }).trim();
  var commandText = normalizedText.toLowerCase();

  saveTelegramUser(chatId, username, firstName);

  if (commandText === "/start") {
    sendTelegramText(chatId, "Chào bạn! Bot đã được kích hoạt.\n\n" +
        "Sử dụng các lệnh:\n" +
        "/register <tên kho> - đăng ký kho của bạn\n" +
        "/mykho - xem thông tin kho hiện tại\n" +
        "/soanhang - lấy link web app soạn hàng\n" +
        "/help - xem hướng dẫn");
    return;
  }

  if (/^\/register\s+/i.test(normalizedText)) {
    var parts = normalizedText.split(/\s+/);
    var storeName = parts.slice(1).join(" ").trim();
    if (!storeName) {
      sendTelegramText(chatId, "Vui lòng gửi theo định dạng: /register <tên kho>\nVí dụ: /register Kho A");
      return;
    }
    updateTelegramUserStore(chatId, storeName);
    sendTelegramText(chatId, "✅ Bạn đã đăng ký kho: " + storeName + "\n" +
        "Từ nay bạn sẽ nhận thông báo đơn hàng cho kho này.");
    return;
  }

  if (commandText === "/mykho") {
    var ss = getSS();
    var sheet = ss.getSheetByName("TelegramUsers");
    var data = sheet ? sheet.getDataRange().getValues() : [];
    var storeName = "(chưa đăng ký kho)";
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(chatId)) {
        storeName = data[i][4] ? data[i][4].toString() : storeName;
        break;
      }
    }
    sendTelegramText(chatId, "Kho hiện tại của bạn: " + storeName + "\n" +
        "Nếu chưa đúng, hãy gửi lại: /register <tên kho>");
    return;
  }

  if (commandText === "/soanhang") {
    sendTelegramText(chatId, "🔗 Link web app soạn hàng: " + WEB_APP_URL + "\n\n" +
        "Mở link và đăng nhập để thao tác với đơn hàng.");
    return;
  }

  if (commandText === "/help" || commandText === "/?" || commandText === "/commands") {
    sendTelegramText(chatId, "Các lệnh hiện có:\n" +
        "/start - bắt đầu sử dụng bot\n" +
        "/register <tên kho> - đăng ký kho nhận thông báo\n" +
        "/mykho - xem kho hiện tại\n" +
        "/soanhang - nhận link web app\n" +
        "/donmoi - xem các đơn mới của kho\n" +
        "/help - xem trợ giúp");
    return;
  }

  if (commandText === "/donmoi" || commandText === "/myorders") {
    var storeName = getTelegramUserStore(chatId);
    if (!storeName) {
      sendTelegramText(chatId, "Bạn chưa đăng ký kho. Hãy dùng lệnh:\n/register <tên kho>");
      return;
    }
    var orders = getPendingOrdersForStore(storeName, 10);
    if (!orders.length) {
      var short = STORE_MAP[storeName] || storeName;
      sendTelegramText(chatId, "Hiện không có đơn mới cho kho: " + storeName + " (" + short + ").\n\nMở web app để kiểm tra toàn bộ đơn: " + WEB_APP_URL);
      return;
    }
    var lines = orders.map(function(o, idx) {
      var sx = o.khoXuat || ""; var sn = o.khoNhan || "";
      var sxShort = STORE_MAP[sx] || sx;
      var snShort = STORE_MAP[sn] || sn;
      return (idx + 1) + ". " + o.soPhieu + " (" + sxShort + " → " + snShort + ")\n" + getOrderWebUrl(o.soPhieu);
    });
    var short = STORE_MAP[storeName] || storeName;
    sendTelegramText(chatId, "🔔 Đơn mới cho kho " + storeName + " (" + short + "):\n" + lines.join("\n\n"));
    return;
  }

  sendTelegramText(chatId, "Xin lỗi, tôi chưa hiểu lệnh này. Gõ /help để xem danh sách lệnh.");
}
function saveTelegramUser(chatId, username, firstName) {
  var ss = getSS();
  var sheet = ss.getSheetByName("TelegramUsers");
  if (!sheet) {
    sheet = ss.insertSheet("TelegramUsers");
    sheet.appendRow(["username", "chat_id", "first_name", "role", "store", "last_active"]);
    sheet.getRange("A1:F1").setFontWeight("bold");
  }

  var data = sheet.getDataRange().getValues();
  var now = new Date();
  var found = false;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(chatId)) {
      sheet.getRange(i+1, 1, 1, 6).setValues([[username, chatId, firstName, data[i][3], data[i][4], now]]);
      found = true;
      break;
    }
  }
  if (!found) {
    sheet.appendRow([username, chatId, firstName, "", "", now]);
  }
}
function getSS() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.action) {
      var action = e.parameter.action;
      var res = null;
      switch(action) {
        case 'getInitialData':
          res = getInitialData();
          break;
        case 'getBootstrapData':
          res = getBootstrapData();
          break;
        case 'getCatalogData':
          res = getCatalogData();
          break;
        case 'getNewProductsList':
          res = getNewProductsList(e.parameter.limit || NEW_PRODUCTS_DEFAULT_LIMIT);
          break;
        case 'getCatalogIsNewAdminList':
          res = getCatalogIsNewAdminList_(e.parameter.q || '', e.parameter.limit || 200);
          break;
        case 'getVariantStockList':
          res = getVariantStockList(e.parameter.parentSku || e.parameter.parent || '');
          break;
        case 'layDanhSachPhieuTheoFilter':
          res = layDanhSachPhieuTheoFilter(e.parameter.khoNhan || '', e.parameter.ngay || '', e.parameter.userRole || '', e.parameter.userStore || '');
          break;
        case 'getChiTietPhieu':
          res = getChiTietPhieu(e.parameter.soPhieu || '', e.parameter.storeName || '', e.parameter.includeStock);
          break;
        case 'getThongTinPhieu':
          res = getThongTinPhieu(e.parameter.soPhieu || '');
          break;
        case 'getDonHangTheoNgay':
          res = getDonHangTheoNgay(e.parameter.ngay || 'today', e.parameter.userRole || '', e.parameter.userStore || '', e.parameter.viewMode || '');
          break;
        case 'warmStockIndex':
          res = warmStockIndex();
          break;
        case 'getStockCacheStatus':
          res = getStockCacheStatus();
          break;
        case 'rebuildTonQ7':
          res = rebuildTonKhoQ7Sheet_(getSS());
          res._debugRun = "q7-v1";
          break;
        case 'getDashboardSummary':
          res = getDashboardSummary(e.parameter.userRole || '', e.parameter.userStore || '', e.parameter.timeline || '2days', e.parameter.fromDate || '', e.parameter.toDate || '');
          break;
        case 'getDanhSachDonSoanHang':
          res = getDanhSachDonSoanHang(
            e.parameter.ngay || '',
            e.parameter.userRole || '',
            e.parameter.userStore || '',
            e.parameter.ngayTo || '',
            e.parameter.packingMode || '',
            e.parameter.includePacked
          );
          break;
        case 'getChiTietDonHangMobile':
          res = getChiTietDonHangMobile(e.parameter.soPhieu || '');
          break;
        case 'debugOrder':
          res = debugOrderInfo_(e.parameter.key || e.parameter.soPhieu || '');
          break;
        case 'layDanhSachXuatBanHang':
          res = layDanhSachXuatBanHang(e.parameter.ngay || '', e.parameter.userRole || '', e.parameter.userStore || '', e.parameter.soHoaDon || '');
          break;
        case 'getDanhSachTaiKhoan':
          res = getDanhSachTaiKhoan();
          break;
        case 'debugTonKho':
          res = debugTonKhoInfo(e.parameter.key || '', e.parameter.storeName || '', e.parameter.maHang || '');
          break;
        default:
          res = { error: 'Unknown action' };
      }
      return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createTemplateFromFile('WebApp')
      .evaluate()
      .setTitle('⚡ Hệ Thống Quản Lý Kho')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

// Hàm "Mồi" để ép Google cấp quyền
function capQuyenHeThong() {
  try {
    DriveApp.getRootFolder();
    SpreadsheetApp.getActiveSpreadsheet();
    UrlFetchApp.fetch("https://www.google.com"); 
    Logger.log("✅ Đã cấp quyền thành công!");
  } catch(e) { Logger.log("Lỗi: " + e.message); }
}

// --- HÀM GỬI TIN NHẮN TELEGRAM ---
function sendTelegramText(chatId, text) {
  if (!TELEGRAM_TOKEN || !chatId) return;
  var url = "https://api.telegram.org/bot" + TELEGRAM_TOKEN + "/sendMessage";
  var payload = {
    chat_id: chatId,
    text: text
  };
  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  try {
    var response = UrlFetchApp.fetch(url, options);
    Logger.log("Telegram send: " + response.getResponseCode() + " " + response.getContentText());
  } catch (e) {
    Logger.log("Telegram error: " + e.message);
  }
}

function getTelegramUsersByStore(storeName) {
  var ss = getSS();
  var sheet = ss.getSheetByName("TelegramUsers");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var users = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][4]) === storeName) {
      users.push({ chatId: data[i][1], username: data[i][0] });
    }
  }
  return users;
}

function getTelegramUsersByStores(storeNames) {
  var ss = getSS();
  var sheet = ss.getSheetByName("TelegramUsers");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var seen = {};
  var users = [];
  for (var i = 1; i < data.length; i++) {
    var store = String(data[i][4]);
    if (storeNames.indexOf(store) !== -1) {
      var chatId = String(data[i][1]);
      if (chatId && !seen[chatId]) {
        seen[chatId] = true;
        users.push({ chatId: chatId, username: data[i][0] });
      }
    }
  }
  return users;
}

function updateTelegramUserStore(chatId, storeName) {
  var ss = getSS();
  var sheet = ss.getSheetByName("TelegramUsers");
  if (!sheet) return false;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(chatId)) {
      sheet.getRange(i + 1, 5).setValue(storeName);
      sheet.getRange(i + 1, 6).setValue(new Date());
      return true;
    }
  }
  return false;
}

function getTelegramUserStore(chatId) {
  var ss = getSS();
  var sheet = ss.getSheetByName("TelegramUsers");
  if (!sheet) return "";
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(chatId)) {
      return data[i][4] ? data[i][4].toString().trim() : "";
    }
  }
  return "";
}

function getPendingOrdersForStore(storeName, maxCount) {
  var ss = getSS();
  var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
  if (!historySheet) return [];
  var targetStore = normalizeStoreName(storeName || "");
  var data = historySheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var soPhieu = data[i][1] ? data[i][1].toString().trim() : "";
    var khoNhan = data[i][3] ? data[i][3].toString().trim() : "";
    var khoXuat = data[i][2] ? data[i][2].toString().trim() : "";
    var slThucTe = data[i][8];
    var slSoanCol = data[i][15];
    var rowStatus = data[i][12] ? String(data[i][12]).trim() : "Mới";
    if (!soPhieu || !khoNhan) continue;
    if (!isSameStoreName(khoNhan, targetStore)) continue;
    var isDaXuLy = getDisplayOrderStatus(rowStatus, slThucTe, slSoanCol) !== "Mới";
    if (!map[soPhieu]) {
      map[soPhieu] = { soPhieu: soPhieu, khoXuat: khoXuat, khoNhan: khoNhan, daXuLy: isDaXuLy };
    } else if (isDaXuLy) {
      map[soPhieu].daXuLy = true;
    }
  }
  var res = [];
  for (var key in map) {
    if (!map[key].daXuLy) {
      res.push(map[key]);
    }
  }
  res.sort(function(a,b){ return a.soPhieu.localeCompare(b.soPhieu); });
  if (maxCount && res.length > maxCount) res = res.slice(0, maxCount);
  return res;
}

function sendTelegramTextToStoreUsers(storeName, text) {
  var users = getTelegramUsersByStore(storeName);
  users.forEach(function (u) {
    sendTelegramText(u.chatId, text);
  });
}

function sendTelegramTextToStores(storeNames, text) {
  var users = getTelegramUsersByStores(storeNames);
  users.forEach(function (u) {
    sendTelegramText(u.chatId, text);
  });
}

function getOrderWebUrl(soPhieu, tabName, isPublic) {
  var tab = tabName || "quan-ly";
  var url = WEB_APP_URL + "?tab=" + encodeURIComponent(tab) + "&soPhieu=" + encodeURIComponent(soPhieu);
  if (isPublic) {
    url += "&public=1";
  }
  return url;
}

function sanitizeFileNamePart(value) {
  var text = String(value || "").trim();
  if (!text) return "don_hang";
  return text.replace(/[\\/:*?"<>|#%&{}\[\]~]/g, "_").replace(/\s+/g, "_");
}

function normalizeOrderCodeText(value) {
  // Đ/đ không luôn tách thành D trong NFKD → map thủ công để khớp Q7-ĐC… ↔ Q7-DC…
  return String(value || "")
    .trim()
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D")
    .toUpperCase()
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Z0-9]/g, "");
}

/** Khớp số phiếu linh hoạt: Q7-DC318957 ↔ DC-318957 ↔ ĐC-318957 */
function orderKeysMatch_(left, right) {
  var a = normalizeOrderCodeText(left);
  var b = normalizeOrderCodeText(right);
  if (!a || !b) return false;
  if (a === b) return true;
  if (a.length >= 6 && b.indexOf(a) !== -1) return true;
  if (b.length >= 6 && a.indexOf(b) !== -1) return true;
  return false;
}

function buildOrderMatchSet_(soPhieuOrList) {
  var set = { _list: [] };
  var list = Object.prototype.toString.call(soPhieuOrList) === "[object Array]" ? soPhieuOrList : [soPhieuOrList];
  for (var i = 0; i < list.length; i++) {
    var s = String(list[i] || "").trim();
    if (!s) continue;
    set[s] = true;
    set[s.toLowerCase()] = true;
    set[normalizeOrderCodeText(s)] = true;
    set._list.push(s);
  }
  return set;
}

function orderInMatchSet_(soPhieu, matchSet) {
  var s = String(soPhieu || "").trim();
  if (!s || !matchSet) return false;
  if (matchSet[s] || matchSet[s.toLowerCase()] || matchSet[normalizeOrderCodeText(s)]) return true;
  var list = matchSet._list || [];
  for (var i = 0; i < list.length; i++) {
    if (orderKeysMatch_(s, list[i])) return true;
  }
  return false;
}

function taoPdfDonHangVaLayLink(soPhieu) {
  try {
    var target = String(soPhieu || "").trim();
    if (!target) return "";
    var targetNormalized = normalizeOrderCodeText(target);
    var ss = getSS();
    var catalogLookup = getCatalogLookup(ss);
    var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
    if (!historySheet) return "";
    var data = historySheet.getDataRange().getValues();
    if (!data || data.length < 2) return "";

    var rows = [];
    var diffRowIndexes = [];
    var khoXuat = "";
    var khoNhan = "";
    var createdAt = "";
    for (var i = 1; i < data.length; i++) {
      var rowSoPhieu = data[i][1] ? String(data[i][1]).trim() : "";
      if (!rowSoPhieu) continue;
      var rowNormalized = normalizeOrderCodeText(rowSoPhieu);
      if (rowSoPhieu.toLowerCase() !== target.toLowerCase() && rowNormalized !== targetNormalized) continue;
      var rowStatus = data[i][12] ? String(data[i][12]).trim() : "Mới";
      if (rowStatus === "Đã hủy dòng") continue;
      if (!khoXuat && data[i][2]) khoXuat = String(data[i][2]).trim();
      if (!khoNhan && data[i][3]) khoNhan = String(data[i][3]).trim();
      if (!createdAt && data[i][0]) createdAt = data[i][0];
      var slDat = Number(data[i][7]) || 0;
      var hasActual = data[i][8] !== "" && data[i][8] !== null && data[i][8] !== undefined;
      var isReceived = rowStatus === "Đã xác nhận nhận hàng";
      // Cột 16 "SL Giao (Soạn)" lưu riêng số lượng đã soạn/giao, không bị ghi đè khi
      // chi nhánh nhận xác nhận số thực nhận (cột 9 dùng chung cho cả 2 giai đoạn).
      var rawSlGiao = data[i][15];
      var hasSlGiaoColumn = rawSlGiao !== "" && rawSlGiao !== null && rawSlGiao !== undefined;
      var slGiao;
      if (hasSlGiaoColumn) {
        slGiao = Number(rawSlGiao) || 0;
      } else if (hasActual && !isReceived) {
        // Dữ liệu cũ (trước khi có cột 16): cột 9 lúc này vẫn đang là số soạn.
        slGiao = Number(data[i][8]) || 0;
      } else {
        slGiao = slDat;
      }
      var slThucNhan = isReceived && hasActual ? Number(data[i][8]) : "";
      var displayQty = slThucNhan !== "" ? slThucNhan : slGiao;
      if (!displayQty || displayQty <= 0) continue;
      if (slThucNhan !== "" && slThucNhan !== slGiao) diffRowIndexes.push(rows.length);
      rows.push([
        rows.length + 1,
        data[i][4] || "",
        data[i][5] || "",
        formatVariantDisplayName_(data[i][4], data[i][6]),
        resolveDvtValue(catalogLookup, data[i][4], data[i][5], data[i][9]),
        slGiao,
        slThucNhan,
        rowStatus || "Mới"
      ]);
    }
    if (!rows.length) return "";

    var tempSheet = recreateTempSheet(ss, "__TMP_TELE_PDF_DON", ["Pdf_", "__TMP_TELE_PDF_DON"]);
    var title = "PHIẾU CHI TIẾT ĐƠN: " + target;
    tempSheet.getRange("A1:H1").merge().setValue(title).setFontSize(14).setFontWeight("bold").setHorizontalAlignment("center");
    var ngayText = createdAt ? Utilities.formatDate(new Date(createdAt), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : "";
    tempSheet.getRange("A2:H2").merge().setValue("Kho xuất: " + khoXuat + " | Kho nhận: " + khoNhan + (ngayText ? " | Thời gian tạo: " + ngayText : "")).setFontStyle("italic");

    var headers = [["STT", "Mã hàng", "Mã vạch", "Tên hàng", "ĐVT", "SL Giao (Soạn)", "SL Thực Nhận", "Trạng thái dòng"]];
    tempSheet.getRange(4, 1, 1, 8).setValues(headers).setFontWeight("bold").setBackground("#d9ead3").setHorizontalAlignment("center");
    tempSheet.getRange(5, 1, rows.length, 8).setValues(rows);
    tempSheet.getRange(4, 1, rows.length + 1, 8).setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
    tempSheet.getRange(5, 6, rows.length, 2).setHorizontalAlignment("right");
    // Bôi màu nổi bật những dòng có SL Giao khác SL Thực Nhận để dễ phát hiện chênh lệch.
    for (var d = 0; d < diffRowIndexes.length; d++) {
      var sheetRow = 5 + diffRowIndexes[d];
      tempSheet.getRange(sheetRow, 6, 1, 2).setBackground("#f4cccc").setFontColor("#990000").setFontWeight("bold");
    }
    tempSheet.setColumnWidth(1, 45);
    tempSheet.setColumnWidth(2, 110);
    tempSheet.setColumnWidth(3, 120);
    tempSheet.setColumnWidth(4, 280);
    tempSheet.setColumnWidth(5, 60);
    tempSheet.setColumnWidth(6, 70);
    tempSheet.setColumnWidth(7, 90);
    tempSheet.setColumnWidth(8, 140);
    SpreadsheetApp.flush();

    var exportUrl = "https://docs.google.com/spreadsheets/d/" + SHEET_ID + "/export?format=pdf&size=A4&portrait=true&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=true&gid=" + tempSheet.getSheetId();
    var response = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) {
      Logger.log("PDF export error " + response.getResponseCode() + ": " + response.getContentText());
      return "";
    }

    var pdfFolder;
    var folderName = "dieuchuyenhanghoa_pdf";
    var folders = DriveApp.getFoldersByName(folderName);
    if (folders.hasNext()) pdfFolder = folders.next();
    else pdfFolder = DriveApp.createFolder(folderName);

    var fileName = "Phieu_" + sanitizeFileNamePart(target) + "_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss") + ".pdf";
    var file = pdfFolder.createFile(response.getBlob().setName(fileName));
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      Logger.log("PDF share warning: " + shareErr.message);
    }
    // Use a stable direct file URL by ID so Telegram recipients open the exact PDF file.
    return "https://drive.google.com/file/d/" + file.getId() + "/view?usp=sharing";
  } catch (err) {
    Logger.log("taoPdfDonHangVaLayLink error: " + err.message);
    return "";
  }
}

function buildTelegramOrderLinkText(soPhieu, pdfUrl) {
  var publicWebUrl = getOrderWebUrl(soPhieu, "quan-ly", true);
  // Always provide the webapp link as a reliable fallback even when PDF link is invalid or expired.
  if (pdfUrl) {
    return "Mở chi tiết đơn: " + publicWebUrl + "\nXem phiếu PDF: " + pdfUrl;
  }
  return "Mở chi tiết đơn: " + publicWebUrl;
}

function sendTelegramMessage(soPhieu, khoXuat, khoNhan, itemCount, pdfUrl) {
  var typeLabel = soPhieu.indexOf("DH") !== -1 ? "ĐƠN HÀNG MỚI" : "LỆNH ĐIỀU CHUYỂN MỚI";
  var kxShort = STORE_MAP[khoXuat] || khoXuat;
  var knShort = STORE_MAP[khoNhan] || khoNhan;
  var text = "📦 *THÔNG BÁO " + typeLabel + "*\n" +
             "*Trạng thái:* Mới\n" +
             "*Số phiếu:* " + soPhieu + "\n" +
             "*Kho xuất:* " + khoXuat + " (" + kxShort + ")\n" +
             "*Kho nhận:* " + khoNhan + " (" + knShort + ")\n" +
             "*Số mặt hàng:* " + itemCount + "\n\n" +
             buildTelegramOrderLinkText(soPhieu, pdfUrl);
  if (TELEGRAM_CHAT_ID) {
    sendTelegramText(TELEGRAM_CHAT_ID, text);
  }
  sendTelegramTextToStores([khoNhan, khoXuat], "📌 Có đơn mới dành cho kho của bạn:\n" + text);
}

function sendTelegramOrderReady(soPhieu, khoNhan, pdfUrl) {
  var knShort = STORE_MAP[khoNhan] || khoNhan;
  var text = "✅ *ĐÃ HOÀN THÀNH SOẠN HÀNG*\n" +
             "*Trạng thái:* Đã soạn\n" +
             "*Số phiếu:* " + soPhieu + "\n" +
             "*Kho nhận:* " + khoNhan + " (" + knShort + ")\n\n" +
             buildTelegramOrderLinkText(soPhieu, pdfUrl);
  if (TELEGRAM_CHAT_ID) {
    sendTelegramText(TELEGRAM_CHAT_ID, text);
  }
  if (khoNhan) {
    sendTelegramTextToStoreUsers(khoNhan, text);
  }
}

function sendTelegramOrderCancelled(soPhieu, khoXuat, khoNhan, actor, reason) {
  var kxShort = STORE_MAP[khoXuat] || khoXuat;
  var knShort = STORE_MAP[khoNhan] || khoNhan;
  var reasonText = reason ? "\n*Lý do:* " + reason : "";
  var text = "🛑 *ĐÃ HỦY ĐƠN HÀNG*\n" +
             "*Số phiếu:* " + soPhieu + "\n" +
             "*Kho xuất:* " + khoXuat + " (" + kxShort + ")\n" +
             "*Kho nhận:* " + khoNhan + " (" + knShort + ")\n" +
             "*Người hủy:* " + (actor || "Không xác định") + reasonText + "\n\n" +
             "Thông tin này đã được lưu vào lịch sử đơn hàng.";
  if (TELEGRAM_CHAT_ID) {
    sendTelegramText(TELEGRAM_CHAT_ID, text);
  }
  var targetStores = [];
  if (khoNhan) targetStores.push(khoNhan);
  if (khoXuat) targetStores.push(khoXuat);
  if (targetStores.length) {
    sendTelegramTextToStores(targetStores, "⚠️ Đơn hàng đã bị hủy:\n" + text);
  }
}

function sendTelegramOrderChangeSummary(soPhieu, khoXuat, khoNhan, actionLabel, changeCount, actor, extraText, pdfUrl) {
  var kxShort = STORE_MAP[khoXuat] || khoXuat;
  var knShort = STORE_MAP[khoNhan] || khoNhan;
  var text = "🔄 *ĐƠN ĐÃ THAY ĐỔI*\n" +
             "*Số phiếu:* " + soPhieu + "\n" +
             "*Hành động:* " + actionLabel + "\n" +
             "*Số mã thay đổi:* " + changeCount + "\n" +
             "*Người thực hiện:* " + (actor || "Không xác định") + "\n" +
             (extraText ? "*Chi tiết:* " + extraText + "\n" : "") +
             buildTelegramOrderLinkText(soPhieu, pdfUrl);
  if (TELEGRAM_CHAT_ID) {
    sendTelegramText(TELEGRAM_CHAT_ID, text);
  }
  if (khoNhan) {
    sendTelegramTextToStoreUsers(khoNhan, "📢 Đơn hàng vừa được cập nhật:\n" + text);
  }
}

function sendTelegramPackingSummary(soPhieu, khoXuat, khoNhan, changedCount, totalRows, statusLabel, missingCount, extraCount, actor, pdfUrl) {
  var kxShort = STORE_MAP[khoXuat] || khoXuat;
  var knShort = STORE_MAP[khoNhan] || khoNhan;
  var detailText = "*Kết quả:* " + statusLabel + "\n" +
                   "*Thiếu hàng:* " + missingCount + " mã\n" +
                   "*Thừa hàng:* " + extraCount + " mã\n" +
                   "*Tổng dòng:* " + totalRows + " mã\n";
  var text = "📦 *ĐƠN ĐÃ SOẠN XONG*\n" +
             "*Trạng thái:* Đã soạn\n" +
             "*Số phiếu:* " + soPhieu + "\n" +
             "*Kho xuất:* " + khoXuat + " (" + kxShort + ")\n" +
             "*Kho nhận:* " + khoNhan + " (" + knShort + ")\n" +
             "*Mã thay đổi:* " + changedCount + "\n" +
             detailText +
             "*Người thực hiện:* " + (actor || "Không xác định") + "\n\n" +
             buildTelegramOrderLinkText(soPhieu, pdfUrl);
  if (TELEGRAM_CHAT_ID) {
    sendTelegramText(TELEGRAM_CHAT_ID, text);
  }
  if (khoNhan) {
    sendTelegramTextToStoreUsers(khoNhan, "✅ Đơn đã được soạn xong:\n" + text);
  }
}

function sendTelegramReceiveConfirmation(soPhieu, khoNhan, actor, count, confirmedTotal, changedCount, changedQtyTotal, pdfUrl) {
  var knShort = STORE_MAP[khoNhan] || khoNhan;
  var detailText = changedCount > 0 ? "*Dòng có thay đổi số thực nhận:* " + changedCount + "\n" + "*Tổng số lượng thay đổi:* " + changedQtyTotal + "\n" : "*Không có dòng nào thay đổi số thực nhận.*\n";
  var text = "📥 *XÁC NHẬN NHẬN HÀNG*\n" +
             "*Trạng thái:* Đã xác nhận\n" +
             "*Số phiếu:* " + soPhieu + "\n" +
             "*Kho nhận:* " + khoNhan + " (" + knShort + ")\n" +
             "*Số dòng xác nhận:* " + count + "\n" +
             "*Người xác nhận:* " + (actor || "Không xác định") + "\n" +
             "*Tổng số lượng đã xác nhận:* " + confirmedTotal + "\n" +
             detailText +
             buildTelegramOrderLinkText(soPhieu, pdfUrl);
  if (TELEGRAM_CHAT_ID) {
    sendTelegramText(TELEGRAM_CHAT_ID, text);
  }
  if (khoNhan) {
    sendTelegramTextToStoreUsers(khoNhan, "📥 Có xác nhận nhận hàng mới:\n" + text);
  }
}

function getKhoNhanBySoPhieu(soPhieu) {
  var ss = getSS();
  var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
  if (!historySheet) return "";
  var data = historySheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][1].toString().trim().toLowerCase() === soPhieu.toString().trim().toLowerCase()) {
      return data[i][3] ? data[i][3].toString().trim() : "";
    }
  }
  return "";
}

function getThongTinPhieu(soPhieu) {
  try {
    if (!soPhieu) return null;
    var ss = getSS();
    var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
    if (!historySheet) return null;
    // Không full-scan: chunk ngược + match set, dừng khi thấy đủ số phiếu
    var selectedSet = buildOrderMatchSet_(soPhieu);
    var pack = readHistoryForSelectedOrders_(historySheet, selectedSet, "", 3000);
    var data = pack.data || [[]];
    for (var i = 1; i < data.length; i++) {
      if (!data[i] || !data[i][1]) continue;
      var sp = String(data[i][1]).trim();
      if (!orderInMatchSet_(sp, selectedSet)) continue;
      return {
        soPhieu: sp,
        khoXuat: data[i][2] ? String(data[i][2]).trim() : "",
        khoNhan: data[i][3] ? String(data[i][3]).trim() : ""
      };
    }
    return null;
  } catch (e) {
    Logger.log("getThongTinPhieu error: " + (e.message || e));
    return null;
  }
}

// --- API TÀI KHOẢN & ĐĂNG NHẬP ---
function getOrCreateUserSheet(ss) {
  var sheet = ss.getSheetByName("Tài Khoản");
  if (!sheet) {
    sheet = ss.insertSheet("Tài Khoản");
    sheet.appendRow(["Tên đăng nhập", "Mật khẩu", "Phân quyền", "Chi nhánh"]);
    sheet.appendRow(["admin", "123456", "Admin", "Tất cả"]);
    sheet.getRange("A1:D1").setFontWeight("bold").setBackground("#d9ead3");
  }
  return sheet;
}

function loginUser(username, password) {
  try {
    var ss = getSS();
    var sheet = getOrCreateUserSheet(ss);
    var data = sheet.getDataRange().getValues();
    
    var inputUser = String(username).trim();
    var inputPass = String(password).trim();

    for (var i = 1; i < data.length; i++) {
      var sheetUser = String(data[i][0]).trim();
      var sheetPass = String(data[i][1]).trim();
      if (sheetUser === inputUser && sheetPass === inputPass) {
        return { success: true, username: sheetUser, role: data[i][2], store: data[i][3] };
      }
    }
    return { success: false, msg: "Sai tài khoản hoặc mật khẩu!" };
  } catch(e) {
    return { success: false, msg: "Lỗi hệ thống. Vui lòng thử lại sau." };
  }
}

function getDanhSachTaiKhoan() {
  var ss = getSS();
  var sheet = getOrCreateUserSheet(ss);
  var data = sheet.getDataRange().getValues();
  var users = [];
  for (var i = 1; i < data.length; i++) {
    users.push({ user: data[i][0], role: data[i][2], store: data[i][3] });
  }
  return users;
}

function taoTaiKhoanMoi(payload) {
  var user = String(payload && payload.user ? payload.user : "").trim();
  var pass = String(payload && payload.pass ? payload.pass : "").trim();
  var role = String(payload && payload.role ? payload.role : "").trim();
  var storeRaw = String(payload && payload.store ? payload.store : "").trim();

  if (!user || !pass) return { success: false, msg: "Thiếu tên đăng nhập hoặc mật khẩu." };
  if (!role) return { success: false, msg: "Thiếu phân quyền." };
  if (pass.length < 4) return { success: false, msg: "Mật khẩu phải có ít nhất 4 ký tự." };

  if (role === "Admin") {
    storeRaw = "Tất cả";
  } else {
    if (!storeRaw || storeRaw === "Tất cả") {
      return { success: false, msg: "Chi nhánh phải chọn kho quản lý cụ thể." };
    }
    var normalized = normalizeStoreName(storeRaw);
    var validStores = getRuntimeStores();
    var matchedStore = "";
    for (var s = 0; s < validStores.length; s++) {
      if (isSameStoreName(validStores[s], normalized)) {
        matchedStore = validStores[s];
        break;
      }
    }
    if (!matchedStore) {
      return { success: false, msg: "Kho/chi nhánh không hợp lệ: " + storeRaw };
    }
    storeRaw = matchedStore;
  }

  var ss = getSS();
  var sheet = getOrCreateUserSheet(ss);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === user.toLowerCase()) {
      return { success: false, msg: "Tên đăng nhập đã tồn tại!" };
    }
  }
  sheet.appendRow([user, pass, role, storeRaw]);
  return { success: true };
}

function doiMatKhau(payload) {
  var user = String(payload && payload.user ? payload.user : "").trim();
  var oldPassword = String(payload && payload.oldPassword ? payload.oldPassword : "").trim();
  var newPassword = String(payload && payload.newPassword ? payload.newPassword : "").trim();
  if (!user || !oldPassword || !newPassword) throw new Error("Thiếu thông tin mật khẩu.");
  if (newPassword.length < 4) throw new Error("Mật khẩu mới phải có ít nhất 4 ký tự.");

  var ss = getSS();
  var sheet = getOrCreateUserSheet(ss);
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === user) {
      if (String(data[i][1]).trim() !== oldPassword) {
        return { success: false, msg: "Mật khẩu hiện tại không đúng." };
      }
      sheet.getRange(i + 1, 2).setValue(newPassword);
      return { success: true, msg: "Đổi mật khẩu thành công." };
    }
  }

  return { success: false, msg: "Không tìm thấy tài khoản." };
}

function normalizeHeaderText(value) {
  if (value === null || value === undefined) return "";
  // Đ/đ không tách trong NFKD — map D trước khi lọc a-z (tránh "Đơn vị tính" → "onvitinh")
  return String(value)
    .trim()
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D")
    .toLowerCase()
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "");
}

function findColumnIndexByAliases(row, aliases) {
  if (!row) return -1;
  for (var c = 0; c < row.length; c++) {
    var normalized = normalizeHeaderText(row[c]);
    for (var i = 0; i < aliases.length; i++) {
      if (normalized.indexOf(aliases[i]) !== -1) return c;
    }
  }
  return -1;
}

/** Cột Parent_SKU / Mã hàng hóa cha / Ma_Nhom_Ban */
function findCatalogParentColIdx_(headerRow) {
  if (!headerRow) return -1;
  for (var c = 0; c < headerRow.length; c++) {
    var n = normalizeHeaderText(headerRow[c]);
    if (!n) continue;
    if (n === "parentsku" || n === "parent" || n === "manhomban" || n === "nhomban") return c;
    if (n.indexOf("parentsku") !== -1 || n.indexOf("manhomban") !== -1) return c;
    if (n.indexOf("mahang") !== -1 && (n.indexOf("cha") !== -1 || n.indexOf("parent") !== -1)) return c;
    if (n.indexOf("mahanghoa") !== -1 && n.indexOf("cha") !== -1) return c;
  }
  return -1;
}

/** Mã hàng chính — bỏ qua cột cha/parent */
function findCatalogMaHangColIdx_(headerRow, parentIdx) {
  if (!headerRow) return -1;
  for (var c = 0; c < headerRow.length; c++) {
    if (parentIdx >= 0 && c === parentIdx) continue;
    var n = normalizeHeaderText(headerRow[c]);
    if (!n) continue;
    if (n.indexOf("cha") !== -1 || n.indexOf("parent") !== -1) continue;
    if (n.indexOf("mahang") !== -1 || n === "sku" || n.indexOf("article") !== -1 || n === "code" || n.indexOf("itemcode") !== -1 || n === "mahh") {
      return c;
    }
  }
  return -1;
}

function findAllColumnIndicesByAliases(row, aliases) {
  if (!row) return [];
  var indexes = [];
  for (var c = 0; c < row.length; c++) {
    var normalized = normalizeHeaderText(row[c]);
    for (var i = 0; i < aliases.length; i++) {
      if (normalized.indexOf(aliases[i]) !== -1) {
        indexes.push(c);
        break;
      }
    }
  }
  return indexes;
}

function getCellValue(row, index, fallback) {
  if (!row || index === undefined || index === null || index < 0 || index >= row.length) return fallback;
  var value = row[index];
  if (value === null || value === undefined || value === "") return fallback;
  return String(value).trim();
}

function getCatalogLookup(ss) {
  var lookup = { byMaHang: {}, byMaVach: {} };
  var catalogResult = getCatalogData();
  if (!catalogResult || !catalogResult.success || !catalogResult.danhMuc) return lookup;

  var danhMuc = catalogResult.danhMuc;
  for (var key in danhMuc) {
    if (!Object.prototype.hasOwnProperty.call(danhMuc, key)) continue;
    var item = danhMuc[key];
    if (!item) continue;
    if (item.maHang) lookup.byMaHang[String(item.maHang).trim().toUpperCase()] = item;
    if (item.maVach) lookup.byMaVach[String(item.maVach).trim().toUpperCase()] = item;
  }
  return lookup;
}

function resolveCatalogProduct(lookup, maHang, maVach) {
  if (!lookup) return null;
  var mv = String(maVach || "").trim().toUpperCase();
  var mh = String(maHang || "").trim().toUpperCase();
  if (mv && lookup.byMaVach[mv]) return lookup.byMaVach[mv];
  if (mh && lookup.byMaHang[mh]) return lookup.byMaHang[mh];
  return null;
}

function resolveDvtValue(lookup, maHang, maVach, currentDvt) {
  // Ép theo Data_Excel khi có ĐVT catalog — không giữ "Cái"/ĐVT sai đã lưu trên đơn
  var catalogItem = resolveCatalogProduct(lookup, maHang, maVach);
  var catalogDvt = catalogItem && catalogItem.dvt ? String(catalogItem.dvt).trim() : "";
  if (catalogDvt) return catalogDvt;
  return String(currentDvt || "").trim();
}

function findHeaderRowIndex(data, maxScanRows) {
  // Dò dòng tiêu đề bằng cách yêu cầu ÍT NHẤT 2 ô khớp alias trong CÙNG một dòng.
  // Chỉ khớp 1 ô là không đủ, vì dòng tiêu đề báo cáo (vd "TỔNG HỢP TỒN KHO") ở ô A1
  // cũng chứa chuỗi con "tonkho" và dễ bị nhận nhầm là dòng tiêu đề thật.
  var limit = Math.min(data.length, maxScanRows || 8);
  var markerAliases = ['mahang', 'mavach', 'tenhang', 'tonkho', 'soluongton', 'cuahang', 'donvitinh', 'dvt', 'dauky', 'nhapkho', 'xuatkho', 'cuoiky'];
  for (var r = 0; r < limit; r++) {
    var row = data[r];
    if (!row) continue;
    var matchCount = 0;
    for (var c = 0; c < row.length; c++) {
      var token = normalizeHeaderText(row[c]);
      if (!token) continue;
      for (var a = 0; a < markerAliases.length; a++) {
        if (token.indexOf(markerAliases[a]) !== -1 || token === 'kho') {
          matchCount++;
          break;
        }
      }
    }
    if (matchCount >= 2) return r;
  }
  return -1;
}

function getOrCreateCatalogSheet(ss) {
  var sheet = ss.getSheetByName("Data_Excel");
  if (!sheet) {
    sheet = ss.insertSheet("Data_Excel");
    sheet.getRange(1, 1, 1, CATALOG_COL_COUNT).setValues([["Mã hàng", "", "Mã vạch", "", "", "Tên hàng hóa", "", "ĐVT", "Ngày tạo", CATALOG_PARENT_HEADER, CATALOG_ISNEW_HEADER]]);
    sheet.getRange(1, 1, 1, CATALOG_COL_COUNT).setFontWeight("bold").setBackground("#d9ead3");
  } else {
    // Bổ sung header Ngày tạo / Parent_SKU / IsNew nếu sheet cũ chưa có
    try {
      var h9 = String(sheet.getRange(1, 9).getValue() || "").trim();
      if (!h9) {
        sheet.getRange(1, 9).setValue("Ngày tạo").setFontWeight("bold").setBackground("#d9ead3");
      }
      var h10 = String(sheet.getRange(1, 10).getValue() || "").trim();
      if (!h10) {
        sheet.getRange(1, 10).setValue(CATALOG_PARENT_HEADER).setFontWeight("bold").setBackground("#d9ead3");
      }
      var h11 = String(sheet.getRange(1, 11).getValue() || "").trim();
      if (!h11) {
        sheet.getRange(1, 11).setValue(CATALOG_ISNEW_HEADER).setFontWeight("bold").setBackground("#d9ead3");
      }
    } catch (e) {}
  }
  return sheet;
}

/** Timestamp chuẩn ghi cột Ngày tạo: Date object + format hiển thị yyyy-MM-dd HH:mm:ss */
function catalogNowStamp_() {
  var now = new Date();
  var tz = Session.getScriptTimeZone() || "Asia/Ho_Chi_Minh";
  return {
    date: now,
    text: Utilities.formatDate(now, tz, "yyyy-MM-dd HH:mm:ss"),
    tz: tz
  };
}

/** Ghi catalog từ entries gọn {mh,mv,th,d,p,n} — layout Data_Excel + Parent_SKU + IsNew */
function writeCatalogEntriesToSheet_(ss, entries, reset) {
  ss = ss || getSS();
  var sh = getOrCreateCatalogSheet(ss);
  var t0 = Date.now();
  if (reset) {
    var oldLastRow = sh.getLastRow();
    var oldLastCol = Math.max(sh.getLastColumn(), CATALOG_COL_COUNT);
    if (oldLastRow > 0) sh.getRange(1, 1, oldLastRow, oldLastCol).clearContent();
    sh.getRange(1, 1, 1, CATALOG_COL_COUNT).setValues([["Mã hàng", "", "Mã vạch", "", "", "Tên hàng hóa", "", "ĐVT", "Ngày tạo", CATALOG_PARENT_HEADER, CATALOG_ISNEW_HEADER]]);
    sh.getRange(1, 1, 1, CATALOG_COL_COUNT).setFontWeight("bold").setBackground("#d9ead3");
  }
  // Đảm bảo header cột I = Ngày tạo (kể cả sheet cũ)
  try {
    var h9 = String(sh.getRange(1, 9).getValue() || "").trim();
    if (!h9) sh.getRange(1, 9).setValue("Ngày tạo").setFontWeight("bold").setBackground("#d9ead3");
  } catch (eH) {}

  var rows = [];
  var withDvt = 0;
  var withParent = 0;
  var stamp = catalogNowStamp_();
  for (var i = 0; i < (entries || []).length; i++) {
    var e = entries[i];
    if (!e) continue;
    var mh = String(e.mh || "").trim();
    var mv = String(e.mv || "").trim();
    var th = String(e.th || "").trim();
    var d = String(e.d || "").trim();
    var p = String(e.p || e.parentSku || "").trim();
    var nFlag = e.n === true || e.isNew === true || e.n === 1 || e.n === "1" ? "TRUE" : "";
    if (!mh && !mv) continue;
    if (d) withDvt++;
    if (p) withParent++;
    // Cột I (Ngày tạo): luôn stamp thời điểm import/tạo SP
    rows.push([mh, "", mv, "", "", th, "", d, stamp.date, p, nFlag]);
  }
  if (rows.length) {
    var startRow = Math.max(sh.getLastRow() + 1, 2);
    // getRange(row, column, numRows, numColumns)
    sh.getRange(startRow, 1, rows.length, CATALOG_COL_COUNT).setValues(rows);
    try {
      sh.getRange(startRow, 9, rows.length, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
    } catch (eFmt) {}
  }
  try { SpreadsheetApp.flush(); } catch (e) {}
  return {
    rows: rows.length,
    withDvt: withDvt,
    withParent: withParent,
    totalRows: Math.max(sh.getLastRow() - 1, 0),
    ms: Date.now() - t0,
    ngayTaoStamp: stamp.text
  };
}

function getOrCreateStockSheet(ss) {
  var sheet = ss.getSheetByName("TỔNG HỢP TỒN KHO");
  if (!sheet) {
    sheet = ss.insertSheet("TỔNG HỢP TỒN KHO");
    sheet.getRange(4, 1, 1, 8).setValues([["Kho", "Mã hàng", "Mã vạch", "Tên hàng hóa", "ĐVT", "", "Tồn kho", "Cửa hàng"]]);
    sheet.getRange(4, 1, 1, 8).setFontWeight("bold").setBackground("#d9ead3");
    sheet.setFrozenRows(4);
  }
  return sheet;
}

function normalizeImportedMatrix(fileData) {
  if (!fileData || !fileData.length) throw new Error("File tải lên không có dữ liệu.");
  var rows = [];
  var maxCols = 0;
  for (var i = 0; i < fileData.length; i++) {
    var row = Array.isArray(fileData[i]) ? fileData[i].slice() : [fileData[i]];
    var hasData = false;
    for (var c = 0; c < row.length; c++) {
      if (row[c] !== "" && row[c] !== null && row[c] !== undefined) {
        hasData = true;
      }
    }
    if (!hasData && !rows.length) continue;
    rows.push(row);
    if (row.length > maxCols) maxCols = row.length;
  }
  while (rows.length) {
    var lastRow = rows[rows.length - 1];
    var rowHasData = false;
    for (var j = 0; j < lastRow.length; j++) {
      if (lastRow[j] !== "" && lastRow[j] !== null && lastRow[j] !== undefined) {
        rowHasData = true;
        break;
      }
    }
    if (rowHasData) break;
    rows.pop();
  }
  if (!rows.length || maxCols < 1) throw new Error("File tải lên không có dữ liệu hợp lệ.");
  for (var r = 0; r < rows.length; r++) {
    while (rows[r].length < maxCols) rows[r].push("");
  }
  return rows;
}

function removeColumnFromMatrix(fileData, columnIndex) {
  var rows = normalizeImportedMatrix(fileData);
  var result = [];
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i].slice();
    if (columnIndex >= 0 && columnIndex < row.length) {
      row.splice(columnIndex, 1);
    }
    result.push(row);
  }
  return result;
}

function writeImportedDataToSheet(sheet, fileData) {
  var rows = normalizeImportedMatrix(fileData);
  var oldLastRow = sheet.getLastRow();
  var oldLastCol = sheet.getLastColumn();
  var newRowCount = rows.length;
  var newColCount = rows[0].length;
  if (oldLastRow > 0 && oldLastCol > 0) {
    sheet.getRange(1, 1, oldLastRow, oldLastCol).clearContent();
  }
  sheet.getRange(1, 1, newRowCount, newColCount).setValues(rows);
  return { rows: newRowCount, cols: newColCount };
}

function appendImportedDataToSheet_(sheet, fileData) {
  var rows = normalizeImportedMatrix(fileData);
  if (!rows.length) return { rows: 0, cols: 0 };
  var colCount = rows[0].length;
  var startRow = Math.max(sheet.getLastRow() + 1, 1);
  sheet.getRange(startRow, 1, rows.length, colCount).setValues(rows);
  return { rows: rows.length, cols: colCount, startRow: startRow };
}

function isMovementReportColumnHeader(value) {
  // Cột báo cáo biến động kho kiểu "Đầu kỳ / Nhập kho / Xuất kho / Cuối kỳ".
  // Các cột này chứa chữ "kho" (vd "Nhập kho", "Xuất kho") nên dễ bị nhận nhầm
  // thành cột tên cửa hàng/kho nếu chỉ so khớp chuỗi con "kho".
  var text = normalizeHeaderText(String(value || ""));
  if (!text) return false;
  return text.indexOf('dauky') !== -1 || text.indexOf('nhapkho') !== -1 || text.indexOf('nhap') !== -1 ||
    text.indexOf('xuatkho') !== -1 || text.indexOf('xuat') !== -1 || text.indexOf('cuoiky') !== -1 ||
    text.indexOf('tonkho') !== -1 || text.indexOf('soluongton') !== -1 || text.indexOf('stock') !== -1 ||
    text.indexOf('onhand') !== -1 || text.indexOf('slton') !== -1 || text.indexOf('qty') !== -1;
}

function looksLikeStoreHeaderName(value) {
  var text = normalizeHeaderText(String(value || ""));
  if (!text) return false;
  if (text.indexOf('mahang') !== -1 || text.indexOf('mavach') !== -1 || text.indexOf('tenhang') !== -1 || text.indexOf('dvt') !== -1 || text.indexOf('donvi') !== -1 || text.indexOf('unit') !== -1) {
    return false;
  }
  if (isMovementReportColumnHeader(text)) return false;
  if (text.indexOf('kho') !== -1 || text.indexOf('cuahang') !== -1 || text.indexOf('chinhanh') !== -1 || text.indexOf('store') !== -1 || text.indexOf('tenkho') !== -1) return true;
  if (text.indexOf('q7') !== -1 || text.indexOf('q8') !== -1 || text.indexOf('q1') !== -1 || text.indexOf('q4') !== -1 || text.indexOf('q5') !== -1 || text.indexOf('ph') !== -1 || text.indexOf('k9') !== -1 || text.indexOf('quan') !== -1) return true;
  if (text.indexOf('quận') !== -1 || text.indexOf('quan') !== -1) return true;
  return false;
}

// DEBUG TAM THOI - dung de kiem tra dung cau truc sheet ton kho tren Google Sheet that.
// Goi: <exec_url>?action=debugTonKho&key=TK_DEBUG_2026&storeName=K9%20Qu%E1%BA%ADn%207&maHang=TKS2015
// XOA HAM NAY (va case 'debugTonKho' trong doGet) sau khi sua xong loi ton kho.
function debugTonKhoInfo(secret, storeName, maHang) {
  if (secret !== 'TK_DEBUG_2026') {
    return { error: 'Invalid key' };
  }
  var ss = getSS();
  var sheet = ss.getSheetByName("TỔNG HỢP TỒN KHO");
  if (!sheet) return { error: 'Khong tim thay sheet TONG HOP TON KHO' };
  var data = sheet.getDataRange().getValues();
  var stockConfig = getStockSheetConfig(data);
  var header = data[stockConfig.headerIndex] || [];
  var sampleRows = [];
  for (var i = stockConfig.startRow; i < Math.min(data.length, stockConfig.startRow + 15); i++) {
    sampleRows.push(data[i]);
  }
  var result = {
    sheetName: sheet.getName(),
    totalRows: data.length,
    totalCols: header.length,
    headerIndex: stockConfig.headerIndex,
    header: header,
    stockConfig: stockConfig,
    sampleRows: sampleRows
  };
  if (storeName) {
    var tonKhoMap = getStockMapForStore(ss, storeName);
    result.storeNameInput = storeName;
    result.storeNameNormalized = normalizeStoreName(storeName);
    result.tonKhoMapKeyCount = Object.keys(tonKhoMap).length;
    result.tonKhoMapSample = {};
    var keys = Object.keys(tonKhoMap);
    for (var k = 0; k < Math.min(keys.length, 25); k++) {
      result.tonKhoMapSample[keys[k]] = tonKhoMap[keys[k]];
    }
    if (maHang) {
      result.testMaHangValue = getStockValueForItem(tonKhoMap, maHang, maHang);
      var mhKey = 'MH:' + normalizeProductCode(maHang);
      var mvKey = 'MV:' + normalizeProductCode(maHang);
      result.testMaHangKeyLookup = {};
      result.testMaHangKeyLookup[mhKey] = tonKhoMap[mhKey];
      result.testMaHangKeyLookup[mvKey] = tonKhoMap[mvKey];
    }
  }
  return result;
}

function getStockSheetConfig(stockData) {
  var headerIndex = findHeaderRowIndex(stockData, 10);
  if (headerIndex < 0) {
    return { startRow: 4, headerIndex: 0, storeIndexes: [0, 7], storeHeaderIndexes: [], maHangIdx: 1, maVachIdx: 2, tonKhoIdx: 6, dvtIdx: -1, tenHangIdx: 3, requireStoreRowPrefix: false };
  }
  var header = stockData[headerIndex] || [];

  // Xác định trước các cột đã có ý nghĩa rõ ràng (mã hàng, mã vạch, tồn kho, tên hàng, đvt)
  // để loại trừ chúng (và các cột báo cáo biến động như "Nhập kho"/"Xuất kho") khỏi danh sách
  // cột được đoán là "cửa hàng/kho" chỉ vì chứa chữ "kho".
  var maHangIdx = findColumnIndexByAliases(header, ['mahang', 'mahanghoa', 'sku', 'mahh', 'code', 'itemcode']);
  var maVachIdx = findColumnIndexByAliases(header, ['mavach', 'barcode', 'ean', 'barcodeid']);
  var tonKhoIdx = findColumnIndexByAliases(header, ['tonkho', 'soluongton', 'stock', 'onhand', 'slton', 'qty', 'cuoiky']);
  var tenHangIdx = findColumnIndexByAliases(header, ['tenhang', 'tenhanghoa', 'name', 'description']);
  var dvtIdx = findColumnIndexByAliases(header, ['dvt', 'donvitinh', 'donvi', 'unit', 'uom']);
  var claimedIndexes = [maHangIdx, maVachIdx, tonKhoIdx, tenHangIdx, dvtIdx];

  var storeIndexes = findAllColumnIndicesByAliases(header, ['kho', 'cuahang', 'chinhanh', 'store', 'tenkho']).filter(function (idx) {
    if (claimedIndexes.indexOf(idx) !== -1) return false;
    return !isMovementReportColumnHeader(header[idx]);
  });
  var storeHeaderIndexes = [];
  for (var c = 0; c < header.length; c++) {
    if (storeIndexes.indexOf(c) !== -1) continue;
    if (claimedIndexes.indexOf(c) !== -1) continue;
    if (looksLikeStoreHeaderName(header[c])) storeHeaderIndexes.push(c);
  }

  var isSummaryStockLayout = (tenHangIdx !== -1 && maHangIdx !== -1 && tonKhoIdx !== -1 && storeIndexes.length === 0);
  if (isSummaryStockLayout) {
    // File "TỔNG HỢP TỒN KHO (24).xlsx": mỗi sản phẩm có 1 dòng tổng, theo sau là các dòng
    // con (mỗi dòng 1 kho) - tên kho nằm cùng cột với tên hàng hóa (cột A).
    storeIndexes = [tenHangIdx];
    if (maVachIdx === -1) maVachIdx = maHangIdx;
  }

  var startRow = headerIndex + 1;
  var markerRow = stockData[headerIndex + 1] || [];
  if (isMarkerRow(markerRow)) {
    startRow = headerIndex + 2;
  }

  return {
    startRow: startRow,
    headerIndex: headerIndex,
    storeIndexes: storeIndexes.length ? storeIndexes : [0, 7],
    storeHeaderIndexes: storeHeaderIndexes,
    maHangIdx: maHangIdx === -1 ? 1 : maHangIdx,
    maVachIdx: maVachIdx === -1 ? 2 : maVachIdx,
    tonKhoIdx: tonKhoIdx === -1 ? 6 : tonKhoIdx,
    dvtIdx: dvtIdx,
    tenHangIdx: tenHangIdx,
    requireStoreRowPrefix: isSummaryStockLayout
  };
}

function isMarkerRow(row) {
  // Dòng chú thích thứ tự cột kiểu "(1) (2) ... (7) = (4) + (5) - (6)".
  // Google Sheets có thể tự chuyển "(1)" thành số âm -1 (định dạng kế toán),
  // nên phải nhận diện cả trường hợp số âm lẫn chuỗi "(n)".
  if (!row) return false;
  var markerCount = 0;
  for (var c = 0; c < row.length; c++) {
    var v = row[c];
    if (typeof v === 'number' && v < 0 && Math.round(v) === v) {
      markerCount++;
      continue;
    }
    var text = String(v === null || v === undefined ? "" : v).trim();
    if (/^\(\d+\)/.test(text)) markerCount++;
  }
  return markerCount >= 2;
}

function isLikelyStoreRowName(value) {
  var text = String(value || "").trim();
  if (!text) return false;
  var normalized = normalizeHeaderText(text);
  return normalized.indexOf('kho') === 0 || normalized.indexOf('k9') === 0;
}

function getRowStoreNames(row, stockConfig) {
  var stores = [];
  for (var i = 0; i < stockConfig.storeIndexes.length; i++) {
    var name = getCellValue(row, stockConfig.storeIndexes[i], "");
    if (stockConfig.requireStoreRowPrefix && !isLikelyStoreRowName(name)) continue;
    if (name && stores.indexOf(name) === -1) stores.push(name);
  }
  return stores;
}

function parseImportRows(importData) {
  var headerIndex = findHeaderRowIndex(importData, 10);
  if (headerIndex < 0) throw new Error("Không tìm thấy dòng tiêu đề trong sheet nguồn nhập khẩu.");
  var header = importData[headerIndex];
  var idxMaHang = findColumnIndexByAliases(header, ['mahang', 'sku', 'mahh', 'code', 'itemcode']);
  var idxMaVach = findColumnIndexByAliases(header, ['mavach', 'barcode', 'ean', 'barcodeid']);
  var idxTenHang = findColumnIndexByAliases(header, ['tenhang', 'tensanpham', 'name', 'description']);
  var idxDvt = findColumnIndexByAliases(header, ['dvt', 'donvitinh', 'donvi', 'unit', 'uom']);
  var idxTonKho = findColumnIndexByAliases(header, ['tonkho', 'soluongton', 'stock', 'onhand', 'slton', 'qty']);
  var idxKho = findColumnIndexByAliases(header, ['kho', 'cuahang', 'chinhanh', 'store', 'tenkho']);

  if (idxMaHang === -1 && idxMaVach === -1) {
    throw new Error("Sheet nguồn thiếu cột Mã hàng hoặc Mã vạch.");
  }

  var parsed = [];
  for (var r = headerIndex + 1; r < importData.length; r++) {
    var row = importData[r];
    if (!row) continue;
    var maHang = getCellValue(row, idxMaHang, "");
    var maVach = getCellValue(row, idxMaVach, "");
    var tenHang = getCellValue(row, idxTenHang, "");
    var dvt = getCellValue(row, idxDvt, "");
    var kho = getCellValue(row, idxKho, "");
    var tonRaw = idxTonKho !== -1 ? row[idxTonKho] : "";
    var tonKho = (tonRaw === "" || tonRaw === null || tonRaw === undefined) ? "" : Number(tonRaw);

    if (!maHang && !maVach && !tenHang) continue;

    parsed.push({
      maHang: maHang,
      maVach: maVach,
      tenHang: tenHang,
      dvt: dvt || "",
      kho: kho,
      tonKho: isNaN(tonKho) ? "" : tonKho
    });
  }
  return parsed;
}

function nhapKhauCapNhatThongTin(payload) {
  var actor = payload && payload.actor ? String(payload.actor).trim() : "";
  requireAdmin(actor);

  var importType = payload && payload.importType ? String(payload.importType).trim() : "";
  var fileData = payload && payload.fileData ? payload.fileData : null;

  var ss = getSS();

  // Import nhanh: chỉ nhận map Q7 đã tách phía trình duyệt → ghi sheet TON_Q7
  if (importType === 'stockQ7') {
    var tQ7 = Date.now();
    var entries = payload.q7Entries || [];
    if (!entries.length) {
      throw new Error("Không có dữ liệu tồn Q7 để ghi. Kiểm tra file có cột/kho Q7 và cột ĐVT không.");
    }
    var q7Fast = writeTonQ7EntriesToSheet_(ss, entries);
    var withDvt = 0;
    for (var wi = 0; wi < entries.length; wi++) {
      if (entries[wi] && (entries[wi].d || (entries[wi].k && String(entries[wi].k).indexOf("|DV:") !== -1))) withDvt++;
    }
    // Tách mã có Parent_SKU → TON_VARIANT (Ton_Ban_Dau, reset Da_Xuat)
    var variantImport = { rows: 0 };
    try {
      variantImport = importTonVariantFromStockEntries_(ss, entries) || { rows: 0 };
    } catch (eVarImp) {
      Logger.log(eVarImp);
    }
    return {
      success: true,
      importType: importType,
      targetSheet: TON_Q7_SHEET_NAME,
      updatedRows: q7Fast.rows || 0,
      updatedCols: 4,
      q7Sheet: TON_Q7_SHEET_NAME,
      q7Rows: q7Fast.rows || 0,
      q7WithDvt: withDvt,
      variantRows: variantImport.rows || 0,
      done: true,
      _debugTotalMs: Date.now() - tQ7,
      _debugQ7Ms: q7Fast.ms || 0,
      _debugRun: "import-q7-variant-v2",
      msg: "Đã cập nhật " + TON_Q7_SHEET_NAME + " (" + (q7Fast.rows || 0) + " dòng)" +
        ((variantImport.rows || 0) ? (" + " + TON_VARIANT_SHEET_NAME + " (" + variantImport.rows + " biến thể, reset Da_Xuat)") : "") +
        "."
    };
  }

  // Import riêng TON_VARIANT (Admin)
  if (importType === "stockVariant") {
    var tVar = Date.now();
    var vEntries = payload.variantEntries || payload.q7Entries || [];
    if (!vEntries.length) throw new Error("Không có dòng tồn biến thể để ghi.");
    var vInfo = writeTonVariantEntriesToSheet_(ss, vEntries);
    return {
      success: true,
      importType: importType,
      targetSheet: TON_VARIANT_SHEET_NAME,
      updatedRows: vInfo.rows || 0,
      variantRows: vInfo.rows || 0,
      done: true,
      _debugTotalMs: Date.now() - tVar,
      _debugRun: "import-ton-variant-v2",
      msg: "Đã cập nhật " + TON_VARIANT_SHEET_NAME + " (" + (vInfo.rows || 0) + " dòng) — Ton_Ban_Dau mới, Da_Xuat=0."
    };
  }

  // Import nhanh catalog: FE đã tách sẵn mã/tên/ĐVT → ghi Data_Excel (payload nhỏ, tránh timeout)
  if (importType === 'catalogFast') {
    var tCat = Date.now();
    var catEntries = payload.catalogEntries || [];
    if (!catEntries.length) {
      throw new Error("Không có dòng catalog để ghi. Kiểm tra file có cột mã hàng / mã vạch / tên / ĐVT.");
    }
    var chunkIndexC = Number(payload.chunkIndex);
    if (isNaN(chunkIndexC) || chunkIndexC < 0) chunkIndexC = 0;
    var chunkTotalC = Number(payload.chunkTotal);
    if (isNaN(chunkTotalC) || chunkTotalC < 1) chunkTotalC = 1;
    var isFirstC = chunkIndexC === 0;
    var isLastC = chunkIndexC >= chunkTotalC - 1;
    var catInfo = writeCatalogEntriesToSheet_(ss, catEntries, isFirstC);
    if (isLastC) invalidateCatalogCache_();
    return {
      success: true,
      importType: importType,
      targetSheet: 'Data_Excel',
      updatedRows: catInfo.rows || 0,
      updatedCols: 8,
      chunkIndex: chunkIndexC,
      chunkTotal: chunkTotalC,
      done: isLastC,
      withDvt: catInfo.withDvt || 0,
      _debugTotalMs: Date.now() - tCat,
      _debugRun: "import-catalog-fast-v1",
      msg: isLastC
        ? ("Đã cập nhật nhanh Data_Excel (" + (catInfo.totalRows || catInfo.rows || 0) + " dòng, " + (catInfo.withDvt || 0) + " có ĐVT).")
        : ("Đã nhận chunk catalog " + (chunkIndexC + 1) + "/" + chunkTotalC + ".")
    };
  }

  if (fileData && importType) {
    var chunkIndex = Number(payload.chunkIndex);
    if (isNaN(chunkIndex) || chunkIndex < 0) chunkIndex = 0;
    var chunkTotal = Number(payload.chunkTotal);
    if (isNaN(chunkTotal) || chunkTotal < 1) chunkTotal = 1;
    var isFirstChunk = chunkIndex === 0;
    var isLastChunk = chunkIndex >= chunkTotal - 1;
    var parseMatrix = payload.parseMatrix && payload.parseMatrix.length
      ? normalizeImportedMatrix(payload.parseMatrix)
      : null;

    if (importType === 'stock') {
      var tImport0 = Date.now();
      var stockSheetDirect = getOrCreateStockSheet(ss);
      var matrix = normalizeImportedMatrix(fileData);
      var stockWriteInfo = isFirstChunk
        ? writeImportedDataToSheet(stockSheetDirect, matrix)
        : appendImportedDataToSheet_(stockSheetDirect, matrix);
      SpreadsheetApp.flush();
      if (isFirstChunk) invalidateStoresCache_();

      // Full import: chỉ rebuild TON_Q7 ở chunk cuối (tránh ghi lại sheet mỗi lần)
      var q7Info = { rows: 0, ms: 0 };
      if (isLastChunk) {
        try {
          q7Info = rebuildTonKhoQ7Sheet_(ss);
        } catch (q7Err) {
          Logger.log("rebuildTonQ7 after full import error: " + q7Err);
        }
      }

      return {
        success: true,
        importType: importType,
        targetSheet: 'TỔNG HỢP TỒN KHO',
        updatedRows: stockWriteInfo.rows,
        updatedCols: stockWriteInfo.cols,
        chunkIndex: chunkIndex,
        chunkTotal: chunkTotal,
        done: isLastChunk,
        q7Sheet: TON_Q7_SHEET_NAME,
        q7Rows: q7Info.rows || 0,
        _debugTotalMs: Date.now() - tImport0,
        _debugQ7Ms: q7Info.ms || 0,
        _debugRun: "import-chunk-v4",
        msg: isLastChunk
          ? ('Đã cập nhật TỔNG HỢP TỒN KHO và sheet ' + TON_Q7_SHEET_NAME + ' (' + (q7Info.rows || 0) + ' mã/ĐVT Q7).')
          : ('Đã nhận chunk ' + (chunkIndex + 1) + '/' + chunkTotal + '.')
      };
    }
    if (importType === 'catalog') {
      var catalogSheetDirect = getOrCreateCatalogSheet(ss);
      var adjustedCatalogData = removeColumnFromMatrix(fileData, 3);
      var catalogWriteInfo = isFirstChunk
        ? writeImportedDataToSheet(catalogSheetDirect, adjustedCatalogData)
        : appendImportedDataToSheet_(catalogSheetDirect, adjustedCatalogData);
      SpreadsheetApp.flush();
      if (isLastChunk) invalidateCatalogCache_();
      return {
        success: true,
        importType: importType,
        targetSheet: 'Data_Excel',
        updatedRows: catalogWriteInfo.rows,
        updatedCols: catalogWriteInfo.cols,
        chunkIndex: chunkIndex,
        chunkTotal: chunkTotal,
        done: isLastChunk,
        msg: isLastChunk
          ? 'Đã cập nhật file nhập khẩu thông tin lên sheet Data_Excel sau khi bỏ cột D của file tải lên.'
          : ('Đã nhận chunk ' + (chunkIndex + 1) + '/' + chunkTotal + '.')
      };
    }
    throw new Error('Loại cập nhật không hợp lệ.');
  }

  var sourceSheetName = payload && payload.sourceSheet ? String(payload.sourceSheet).trim() : "";
  if (!sourceSheetName) throw new Error("Thiếu tên sheet nguồn nhập khẩu.");
  var sourceSheet = ss.getSheetByName(sourceSheetName);
  if (!sourceSheet) throw new Error("Không tìm thấy sheet nguồn: " + sourceSheetName);

  var sourceData = sourceSheet.getDataRange().getValues();
  if (!sourceData || sourceData.length < 2) throw new Error("Sheet nguồn không có dữ liệu để nhập khẩu.");

  var parsedRows = parseImportRows(sourceData);
  if (!parsedRows.length) throw new Error("Không có dòng hợp lệ để cập nhật.");

  var warnings = [];

  var catalogMap = {};
  for (var i = 0; i < parsedRows.length; i++) {
    var item = parsedRows[i];
    var key = (item.maVach ? item.maVach.toUpperCase() : "") || (item.maHang ? item.maHang.toUpperCase() : "");
    if (!key) continue;
    catalogMap[key] = {
      maHang: item.maHang || "",
      maVach: item.maVach || "",
      tenHang: item.tenHang || "",
      dvt: item.dvt || ""
    };
  }
  var catalogRows = [];
  for (var key in catalogMap) catalogRows.push(catalogMap[key]);

  var catalogSheet = getOrCreateCatalogSheet(ss);
  var catalogData = catalogSheet.getDataRange().getValues();
  var catalogHeaderIndex = findHeaderRowIndex(catalogData, 8);
  if (catalogHeaderIndex < 0) catalogHeaderIndex = 0;
  var catalogHeader = catalogData[catalogHeaderIndex] || [];
  var cMaHang = findColumnIndexByAliases(catalogHeader, ['mahang', 'sku', 'code']);
  var cMaVach = findColumnIndexByAliases(catalogHeader, ['mavach', 'barcode', 'ean']);
  var cTenHang = findColumnIndexByAliases(catalogHeader, ['tenhang', 'name', 'description']);
  var cDvt = findColumnIndexByAliases(catalogHeader, ['dvt', 'donvitinh', 'donvi', 'unit', 'uom']);
  if (cMaHang === -1) cMaHang = 0;
  if (cMaVach === -1) cMaVach = 2;
  if (cTenHang === -1) cTenHang = 5;
  if (cDvt === -1) cDvt = 7;

  var catalogStartRow = catalogHeaderIndex + 2;
  var catalogColumnCount = Math.max(catalogSheet.getLastColumn(), CATALOG_COL_COUNT);
  var cNgayTao = findColumnIndexByAliases(catalogHeader, ["ngaytao", "createdat", "created", "ngaythem", "importedat"]);
  if (cNgayTao === -1) {
    cNgayTao = 8; // cột I mặc định
    try {
      catalogSheet.getRange(1, 9).setValue("Ngày tạo").setFontWeight("bold").setBackground("#d9ead3");
    } catch (eNg) {}
  }
  if (catalogSheet.getLastRow() >= catalogStartRow) {
    catalogSheet.getRange(catalogStartRow, 1, catalogSheet.getLastRow() - catalogStartRow + 1, catalogColumnCount).clearContent();
  }
  if (catalogRows.length) {
    var catalogWrite = [];
    var importStamp = catalogNowStamp_();
    for (var c = 0; c < catalogRows.length; c++) {
      var rowOut = [];
      for (var z = 0; z < catalogColumnCount; z++) rowOut.push("");
      rowOut[cMaHang] = catalogRows[c].maHang;
      rowOut[cMaVach] = catalogRows[c].maVach;
      rowOut[cTenHang] = catalogRows[c].tenHang;
      rowOut[cDvt] = catalogRows[c].dvt;
      rowOut[cNgayTao] = importStamp.date;
      catalogWrite.push(rowOut);
    }
    catalogSheet.getRange(catalogStartRow, 1, catalogWrite.length, catalogColumnCount).setValues(catalogWrite);
    try {
      catalogSheet.getRange(catalogStartRow, cNgayTao + 1, catalogWrite.length, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
    } catch (eFmt2) {}
  }

  var stockRows = parsedRows.filter(function(r) {
    return r.kho && r.tonKho !== "";
  });
  var stockUpdated = 0;
  if (!stockRows.length) {
    warnings.push("Không tìm thấy đủ dữ liệu Kho + Tồn kho trong sheet nguồn, nên chỉ cập nhật danh mục.");
  } else {
    var stockSheet = getOrCreateStockSheet(ss);
    var stockStartRow = 5;
    var stockColumnCount = Math.max(stockSheet.getLastColumn(), 8);
    if (stockSheet.getLastRow() >= stockStartRow) {
      stockSheet.getRange(stockStartRow, 1, stockSheet.getLastRow() - stockStartRow + 1, stockColumnCount).clearContent();
    }

    var stockWrite = [];
    for (var s = 0; s < stockRows.length; s++) {
      var rowStock = [];
      for (var q = 0; q < stockColumnCount; q++) rowStock.push("");
      rowStock[0] = stockRows[s].kho;
      rowStock[1] = stockRows[s].maHang;
      rowStock[2] = stockRows[s].maVach;
      rowStock[3] = stockRows[s].tenHang;
      rowStock[4] = stockRows[s].dvt;
      rowStock[6] = stockRows[s].tonKho;
      rowStock[7] = stockRows[s].kho;
      stockWrite.push(rowStock);
    }
    stockSheet.getRange(stockStartRow, 1, stockWrite.length, stockColumnCount).setValues(stockWrite);
    stockUpdated = stockWrite.length;
  }

  SpreadsheetApp.flush();
  invalidateCatalogCache_();
  invalidateStoresCache_();

  return {
    success: true,
    catalogUpdated: catalogRows.length,
    stockUpdated: stockUpdated,
    sourceSheet: sourceSheetName,
    warnings: warnings
  };
}

// --- API: LẤY DATA BAN ĐẦU ---
function getBootstrapData() {
  try {
    var ss = getSS();
    var registry = getStoreRegistry(ss);
    var newProductsRes = getNewProductsList(NEW_PRODUCTS_DEFAULT_LIMIT);
    return {
      success: true,
      stores: registry.stores,
      storeMap: registry.storeMap,
      storeDetails: registry.storeDetails || [],
      catalogVersion: getCatalogVersion_(),
      newProducts: (newProductsRes && newProductsRes.success && newProductsRes.data) ? newProductsRes.data : []
    };
  } catch (e) {
    return { success: false, error: e.message || String(e) };
  }
}

function getCatalogData() {
  try {
    var version = getCatalogVersion_();
    var cacheKey = CACHE_CATALOG_PREFIX + version;
    var cache = getScriptCache_();
    // Cache chunked — hỗ trợ catalog > 90KB (tránh miss cache khiến đọc sheet mỗi lần)
    var cached = getCacheJson_(cache, cacheKey);
    if (cached && cached.success && cached.danhMuc) return cached;

    var danhMuc = buildCatalogFromSheet_(getSS());
    var result = { success: true, danhMuc: danhMuc, version: version };
    try {
      putCacheJson_(cache, cacheKey, result, CACHE_TTL_SECONDS);
    } catch (e) {}
    return result;
  } catch (e) {
    return { success: false, error: e.message || String(e) };
  }
}

function parseCatalogNgayTaoCell_(rawNgay, tz) {
  tz = tz || Session.getScriptTimeZone() || "Asia/Ho_Chi_Minh";
  var out = { ms: 0, label: "" };
  if (rawNgay == null || rawNgay === "") return out;
  if (rawNgay instanceof Date && !isNaN(rawNgay.getTime())) {
    out.ms = rawNgay.getTime();
    out.label = Utilities.formatDate(rawNgay, tz, "dd/MM/yyyy HH:mm");
    return out;
  }
  var s = String(rawNgay).trim();
  if (!s) return out;
  // Hỗ trợ "yyyy-MM-dd HH:mm:ss" từ stamp import
  var normalized = s.replace(/(\d{4}-\d{2}-\d{2})\s+(\d{2}:\d{2}:\d{2})/, "$1T$2");
  var parsedMs = Date.parse(normalized);
  if (isNaN(parsedMs)) parsedMs = Date.parse(s);
  if (!isNaN(parsedMs)) {
    out.ms = parsedMs;
    out.label = s;
  }
  return out;
}

/**
 * Hàng mới thông minh:
 * 1) Admin tick IsNew
 * 2) Nếu có NgayTao đủ dữ liệu → sort giảm dần
 * 3) Nếu thiếu/trống NgayTao → quét ngược từ dòng CUỐI Data_Excel
 * 4) Lấp đầy Top 8
 */
function getAutoNewProductsList_(ss, limit) {
  var FILL_MIN = 8;
  limit = Math.max(1, Math.min(Number(limit) || NEW_PRODUCTS_DEFAULT_LIMIT, 20));
  var FILL_MAX = Math.max(FILL_MIN, Math.min(limit, 10));
  ss = ss || getSS();
  var dataSheet = ss.getSheetByName("Data_Excel");
  if (!dataSheet) return [];

  var rawData = dataSheet.getDataRange().getValues();
  if (!rawData || rawData.length < 2) return [];

  var headerRowIndex = -1;
  var headerRow = null;
  for (var hdr = 0; hdr < Math.min(rawData.length, 6); hdr++) {
    if (!rawData[hdr]) continue;
    var hit = false;
    for (var c = 0; c < rawData[hdr].length; c++) {
      var cellText = normalizeHeaderText(rawData[hdr][c]);
      if (cellText.indexOf("mahang") !== -1 || cellText.indexOf("mavach") !== -1 || cellText.indexOf("tenhang") !== -1) {
        hit = true;
        break;
      }
    }
    if (hit) {
      headerRowIndex = hdr;
      headerRow = rawData[hdr];
      break;
    }
  }

  var parentSkuIdxNp = findCatalogParentColIdx_(headerRow);
  var maHangIdx = findCatalogMaHangColIdx_(headerRow, parentSkuIdxNp);
  if (maHangIdx === -1) maHangIdx = findColumnIndexByAliases(headerRow, ["mahang", "sku", "article", "code"]);
  var maVachIdx = findColumnIndexByAliases(headerRow, ["mavach", "barcode", "barcodeid"]);
  var tenHangIdx = findColumnIndexByAliases(headerRow, ["tenhang", "name", "tênhang", "description"]);
  var dvtIdx = findColumnIndexByAliases(headerRow, ["dvt", "donvitinh", "donvi", "unit", "uom"]);
  // Không dùng alias "created"/"timestamp" quá rộng — tránh khớp nhầm cột
  var ngayTaoIdx = findColumnIndexByAliases(headerRow, ["ngaytao", "createdat", "ngaythem", "importedat"]);
  if (ngayTaoIdx === -1 && headerRow && headerRow.length >= 9) {
    var hI = normalizeHeaderText(headerRow[8]);
    if (!hI || hI.indexOf("ngay") !== -1 || hI.indexOf("tao") !== -1 || hI.indexOf("created") !== -1) ngayTaoIdx = 8;
  }
  var isNewIdx = findColumnIndexByAliases(headerRow, ["isnew", "trangthaimoi", "hangmoi", "newflag"]);
  if (isNewIdx === -1 && headerRow && headerRow.length >= 11) isNewIdx = 10;
  if (parentSkuIdxNp === -1 && headerRow && headerRow.length >= 10) parentSkuIdxNp = 9;
  var startRow = headerRowIndex >= 0 ? headerRowIndex + 1 : 1;
  var tz = Session.getScriptTimeZone() || "Asia/Ho_Chi_Minh";

  // Quét 1 lần: giữ bản ghi cuối cùng của mỗi key (dòng dưới = mới hơn khi import append)
  var byKey = {};
  var datedCount = 0;
  var orderedBottomUp = []; // unique keys từ dưới lên
  var seenBottom = {};

  for (var i = rawData.length - 1; i >= startRow; i--) {
    if (!rawData[i]) continue;
    var maHang = getCellValue(rawData[i], maHangIdx !== -1 ? maHangIdx : 0, "");
    var maVach = getCellValue(rawData[i], maVachIdx !== -1 ? maVachIdx : 2, "");
    var tenHang = getCellValue(rawData[i], tenHangIdx !== -1 ? tenHangIdx : 5, "");
    var dvt = getCellValue(rawData[i], dvtIdx !== -1 ? dvtIdx : 7, "");
    var parentSkuNp = parentSkuIdxNp !== -1 ? getCellValue(rawData[i], parentSkuIdxNp, "") : "";
    if (!maHang && !maVach) continue;

    var key = String(maHang || maVach).trim().toUpperCase();
    if (!key) continue;

    var ngayInfo = ngayTaoIdx !== -1
      ? parseCatalogNgayTaoCell_(rawData[i][ngayTaoIdx], tz)
      : { ms: 0, label: "" };

    var flaggedNew = false;
    if (isNewIdx !== -1) {
      var flagVal = String(rawData[i][isNewIdx] == null ? "" : rawData[i][isNewIdx]).trim().toLowerCase();
      flaggedNew = flagVal === "1" || flagVal === "true" || flagVal === "yes" || flagVal === "x" || flagVal === "moi" || flagVal === "new";
    }

    var sheetRow = i + 1;
    if (!byKey[key]) {
      byKey[key] = {
        maHang: maHang,
        maVach: maVach,
        tenHang: tenHang,
        dvt: dvt || "",
        parentSku: parentSkuNp || "",
        sheetRow: sheetRow,
        ngayMs: ngayInfo.ms || 0,
        ngayTao: ngayInfo.label || "",
        isNew: flaggedNew
      };
      if (ngayInfo.ms > 0) datedCount++;
      if (!seenBottom[key]) {
        seenBottom[key] = true;
        orderedBottomUp.push(key); // đã đi từ dưới lên → thứ tự "mới nhất trước"
      }
    } else if (flaggedNew) {
      byKey[key].isNew = true;
    }
  }

  // Bước 1: Admin tick
  var flagged = [];
  var flaggedKeys = {};
  for (var k in byKey) {
    if (!Object.prototype.hasOwnProperty.call(byKey, k)) continue;
    if (byKey[k].isNew) {
      flagged.push(byKey[k]);
      flaggedKeys[k] = true;
    }
  }
  flagged.sort(function(a, b) {
    if ((b.ngayMs || 0) !== (a.ngayMs || 0)) return (b.ngayMs || 0) - (a.ngayMs || 0);
    return (b.sheetRow || 0) - (a.sheetRow || 0);
  });

  // Bước 2: NgayTao đủ dữ liệu? (>= 8 mã có mốc) — nếu không → bottom-up
  var hasUsableNgayTao = ngayTaoIdx !== -1 && datedCount >= FILL_MIN;
  var autoCandidates = [];
  if (hasUsableNgayTao) {
    for (var k2 in byKey) {
      if (!Object.prototype.hasOwnProperty.call(byKey, k2)) continue;
      if (flaggedKeys[k2]) continue;
      autoCandidates.push(byKey[k2]);
    }
    autoCandidates.sort(function(a, b) {
      if ((b.ngayMs || 0) !== (a.ngayMs || 0)) return (b.ngayMs || 0) - (a.ngayMs || 0);
      return (b.sheetRow || 0) - (a.sheetRow || 0);
    });
  } else {
    // Bước 2 fallback: quét ngược — orderedBottomUp đã là dòng cuối → đầu
    for (var bi = 0; bi < orderedBottomUp.length; bi++) {
      var bk = orderedBottomUp[bi];
      if (flaggedKeys[bk]) continue;
      autoCandidates.push(byKey[bk]);
    }
  }

  // Bước 3: gom Admin → lấp đầy Top 8
  var top = [];
  for (var fi = 0; fi < flagged.length; fi++) {
    var fItem = flagged[fi];
    fItem.isNew = true;
    fItem.isAdminPick = true;
    fItem.sourceReason = "admin";
    fItem.reasonLabel = "ADMIN CHỌN";
    top.push(fItem);
  }
  if (top.length < FILL_MIN) {
    for (var oi = 0; oi < autoCandidates.length && top.length < FILL_MAX; oi++) {
      var oItem = autoCandidates[oi];
      oItem.isNew = false;
      oItem.isAdminPick = false;
      if (hasUsableNgayTao && oItem.ngayMs > 0) {
        oItem.sourceReason = "auto_date";
        oItem.reasonLabel = oItem.ngayTao ? ("Ngày tạo: " + oItem.ngayTao) : "Mới theo ngày";
      } else {
        oItem.sourceReason = "auto_bottom";
        oItem.reasonLabel = "Cuối bảng · dòng " + (oItem.sheetRow || "?");
      }
      top.push(oItem);
    }
  }
  for (var t = 0; t < top.length; t++) top[t].rank = t + 1;
  return top;
}

/** Wrapper tương thích API cũ */
function buildNewProductsList_(ss, limit) {
  return getAutoNewProductsList_(ss, limit);
}

function getNewProductsList(limit) {
  try {
    var lim = Math.max(1, Math.min(Number(limit) || NEW_PRODUCTS_DEFAULT_LIMIT, 20));
    var version = getCatalogVersion_();
    var cache = getScriptCache_();
    var cacheKey = CACHE_NEW_PRODUCTS_PREFIX + version + "_" + lim;
    var cached = getCacheJson_(cache, cacheKey);
    if (cached && cached.success && cached.data) return cached;

    var data = getAutoNewProductsList_(getSS(), lim);
    var result = {
      success: true,
      data: data,
      limit: lim,
      source: "Data_Excel",
      strategy: "admin_isNew_then_ngayTao_or_bottom_up_fill_8"
    };
    try { putCacheJson_(cache, cacheKey, result, CACHE_TTL_SECONDS); } catch (e) {}
    return result;
  } catch (e) {
    return { success: false, error: e.message || String(e), data: [] };
  }
}

/** Admin: danh sách catalog để tick Hàng Mới */
function getCatalogIsNewAdminList_(query, limit) {
  try {
    limit = Math.max(20, Math.min(Number(limit) || 200, 500));
    var q = String(query || "").trim().toUpperCase();
    var ss = getSS();
    var sh = getOrCreateCatalogSheet(ss);
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { success: true, items: [], total: 0 };

    var width = Math.max(sh.getLastColumn(), CATALOG_COL_COUNT);
    var values = sh.getRange(1, 1, lastRow, width).getValues();
    var header = values[0] || [];
    var parentIdx = findCatalogParentColIdx_(header);
    var mhIdx = findCatalogMaHangColIdx_(header, parentIdx);
    if (mhIdx === -1) mhIdx = 0;
    var mvIdx = findColumnIndexByAliases(header, ["mavach", "barcode", "barcodeid"]);
    if (mvIdx === -1) mvIdx = 2;
    var thIdx = findColumnIndexByAliases(header, ["tenhang", "name", "description"]);
    if (thIdx === -1) thIdx = 5;
    var dvtIdx = findColumnIndexByAliases(header, ["dvt", "donvitinh", "donvi", "unit"]);
    if (dvtIdx === -1) dvtIdx = 7;
    var isNewIdx = findColumnIndexByAliases(header, ["isnew", "trangthaimoi", "hangmoi", "newflag"]);
    if (isNewIdx === -1) isNewIdx = 10;

    var items = [];
    var seen = {};
    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      if (!row) continue;
      var mh = String(row[mhIdx] == null ? "" : row[mhIdx]).trim();
      var mv = String(row[mvIdx] == null ? "" : row[mvIdx]).trim();
      var th = String(row[thIdx] == null ? "" : row[thIdx]).trim();
      var dvt = String(row[dvtIdx] == null ? "" : row[dvtIdx]).trim();
      if (!mh && !mv) continue;
      var key = (mh || mv).toUpperCase();
      if (seen[key]) continue;
      seen[key] = true;
      if (q) {
        var hay = (mh + " " + mv + " " + th).toUpperCase();
        if (hay.indexOf(q) === -1) continue;
      }
      var flagRaw = String(row[isNewIdx] == null ? "" : row[isNewIdx]).trim().toLowerCase();
      var isNew = flagRaw === "1" || flagRaw === "true" || flagRaw === "yes" || flagRaw === "x" || flagRaw === "moi" || flagRaw === "new";
      items.push({
        sheetRow: r + 1,
        maHang: mh,
        maVach: mv,
        tenHang: th,
        dvt: dvt,
        isNew: isNew
      });
      if (items.length >= limit) break;
    }
    return { success: true, items: items, total: items.length, isNewCol: isNewIdx + 1 };
  } catch (e) {
    return { success: false, error: e.message || String(e), items: [] };
  }
}

/**
 * Admin: lưu batch IsNew vào cột K (hoặc cột IsNew hiện có).
 * payload.flags = [{ sheetRow|maHang|maVach, isNew: true/false }]
 */
function saveCatalogIsNewFlags_(payload) {
  var flags = (payload && payload.flags) ? payload.flags : [];
  if (!flags.length) return { success: false, error: "Không có thay đổi IsNew." };

  var ss = getSS();
  var sh = getOrCreateCatalogSheet(ss);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { success: false, error: "Data_Excel trống." };

  var width = Math.max(sh.getLastColumn(), CATALOG_COL_COUNT);
  var values = sh.getRange(1, 1, lastRow, width).getValues();
  var header = values[0] || [];
  var parentIdx = findCatalogParentColIdx_(header);
  var mhIdx = findCatalogMaHangColIdx_(header, parentIdx);
  if (mhIdx === -1) mhIdx = 0;
  var mvIdx = findColumnIndexByAliases(header, ["mavach", "barcode", "barcodeid"]);
  if (mvIdx === -1) mvIdx = 2;
  var isNewIdx = findColumnIndexByAliases(header, ["isnew", "trangthaimoi", "hangmoi", "newflag"]);
  if (isNewIdx === -1) {
    isNewIdx = 10;
    sh.getRange(1, 11).setValue(CATALOG_ISNEW_HEADER).setFontWeight("bold").setBackground("#d9ead3");
  }

  // Map key -> row index (0-based in values)
  var byMh = {};
  var byMv = {};
  for (var r = 1; r < values.length; r++) {
    var mh = String(values[r][mhIdx] == null ? "" : values[r][mhIdx]).trim().toUpperCase();
    var mv = String(values[r][mvIdx] == null ? "" : values[r][mvIdx]).trim().toUpperCase();
    if (mh && byMh[mh] === undefined) byMh[mh] = r;
    if (mv && byMv[mv] === undefined) byMv[mv] = r;
  }

  var changed = 0;
  for (var i = 0; i < flags.length; i++) {
    var f = flags[i];
    if (!f) continue;
    var rowIdx = -1;
    var sr = Number(f.sheetRow);
    if (sr >= 2 && sr <= lastRow) rowIdx = sr - 1;
    if (rowIdx < 0) {
      var kmh = String(f.maHang || "").trim().toUpperCase();
      var kmv = String(f.maVach || "").trim().toUpperCase();
      if (kmh && byMh[kmh] !== undefined) rowIdx = byMh[kmh];
      else if (kmv && byMv[kmv] !== undefined) rowIdx = byMv[kmv];
    }
    if (rowIdx < 1) continue;
    var nextVal = f.isNew === true || f.isNew === 1 || f.isNew === "1" || f.isNew === "true" ? "TRUE" : "";
    var prevVal = String(values[rowIdx][isNewIdx] == null ? "" : values[rowIdx][isNewIdx]).trim();
    var prevBool = /^(1|true|yes|x|moi|new)$/i.test(prevVal);
    var nextBool = nextVal === "TRUE";
    if (prevBool === nextBool) continue;
    values[rowIdx][isNewIdx] = nextVal;
    changed++;
  }

  if (changed) {
    // Cách 1: chỉ ghi đúng 1 cột IsNew.
    // Sheet.getRange(row, column, numRows, numColumns) — KHÔNG phải lastRow/lastCol.
    var colOut = [];
    for (var rr = 0; rr < values.length; rr++) {
      colOut.push([values[rr][isNewIdx] == null ? "" : values[rr][isNewIdx]]);
    }
    var isNewCol = isNewIdx + 1; // thường = 11 (cột K)
    var numRows = colOut.length;
    sh.getRange(1, isNewCol, numRows, 1).setValues(colOut);
    try { SpreadsheetApp.flush(); } catch (e2) {}
    try {
      PropertiesService.getScriptProperties().setProperty("catalog_version_bump", String(Date.now()));
    } catch (e3) {}
  }

  return {
    success: true,
    changed: changed,
    totalFlags: flags.length,
    isNewCol: isNewIdx + 1,
    msg: "Đã cập nhật " + changed + " sản phẩm Hàng Mới."
  };
}

function getInitialData() {
  var bootstrap = getBootstrapData();
  if (!bootstrap.success) return bootstrap;
  var catalog = getCatalogData();
  if (!catalog.success) return catalog;
  return {
    success: true,
    stores: bootstrap.stores,
    storeMap: bootstrap.storeMap,
    danhMuc: catalog.danhMuc
  };
}

// --- API: LƯU ĐƠN / PHIẾU TẠO MỚI ---
function luuPhieuTuWebApp(payload) {
  var ss = getSS();
  var catalogLookup = getCatalogLookup(ss);
  var homNay = new Date();
  var prefix = payload.loaiPhieu === "DonHang" ? "DH" : "DC";
  var soPhieu = prefix + "-" + Math.floor(100000 + Math.random() * 900000);
  var khoXuatNormalized = normalizeStoreName(payload.khoXuat || "");
  var khoNhanNormalized = normalizeStoreName(payload.khoNhan || "");
  
  var dataLichSuArr = [];
  var coLoiCanDieuChinh = false;
  
  for (var i = 0; i < payload.items.length; i++) {
    var item = payload.items[i];
    var dvtResolved = resolveDvtValue(catalogLookup, item ? item.maHang : "", item ? item.maVach : "", item ? item.dvt : "");
    var slNum = Number(item.sl);
    var ghiChuLoi = "";
    var coLoiDongNay = false;
    if (isNaN(slNum) || slNum <= 0) { coLoiDongNay = true; ghiChuLoi = "Lỗi số lượng"; slNum = 0; }
    if (!dvtResolved || dvtResolved === "Không tìm thấy") { coLoiDongNay = true; ghiChuLoi += (ghiChuLoi ? " | " : "") + "Lỗi ĐVT"; }
    if (item.maHang === "LỖI MÃ") { coLoiDongNay = true; ghiChuLoi += (ghiChuLoi ? " | " : "") + "Mã không tồn tại"; }
    dataLichSuArr.push([ homNay, soPhieu, khoXuatNormalized, khoNhanNormalized, item.maHang, item.maVach, item.tenHang, slNum, "", dvtResolved, "", ghiChuLoi, "Mới" ]);
    if (coLoiDongNay) coLoiCanDieuChinh = true;
  }

  var lock = LockService.getDocumentLock();
  try {
    lock.waitLock(10000);
    var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
    if (!historySheet) {
      historySheet = ss.insertSheet("Lịch Sử Xuất Kho");
      historySheet.getRange("A1:M1").setValues([["Thời gian tạo", "Số Phiếu", "Xuất từ Kho", "Kho Nhận", "Mã hàng", "Mã vạch", "Tên hàng hóa", "Số lượng", "Số lượng thực tế", "ĐVT thực tế", "Ảnh xác nhận", "Ghi chú hệ thống", "Trạng thái"]]).setFontWeight("bold").setBackground("#d9ead3");
      historySheet.setFrozenRows(1);
    }
    ensureHistoryStatusColumn(historySheet);
    var lastRow = historySheet.getLastRow();
    historySheet.getRange(lastRow + 1, 1, dataLichSuArr.length, 13).setValues(dataLichSuArr);
    historySheet.getRange(lastRow + 1, 1, dataLichSuArr.length, 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }

  return { success: true, soPhieu: soPhieu, coLoi: coLoiCanDieuChinh, itemCount: payload.items.length, khoXuat: khoXuatNormalized, khoNhan: khoNhanNormalized };
}

function postProcessNewOrder(payload) {
  var soPhieu = String(payload && payload.soPhieu ? payload.soPhieu : "").trim();
  if (!soPhieu) return { success: false, error: "Thiếu số phiếu." };
  var khoXuat = normalizeStoreName(payload && payload.khoXuat ? payload.khoXuat : "");
  var khoNhan = normalizeStoreName(payload && payload.khoNhan ? payload.khoNhan : "");
  var itemCount = Number(payload && payload.itemCount ? payload.itemCount : 0) || 0;
  try {
    var createPdfUrl = taoPdfDonHangVaLayLink(soPhieu);
    sendTelegramMessage(soPhieu, khoXuat, khoNhan, itemCount, createPdfUrl);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function postProcessPackingOrder(payload) {
  var soPhieu = String(payload && payload.soPhieu ? payload.soPhieu : "").trim();
  if (!soPhieu) return { success: false, error: "Thiếu số phiếu." };
  try {
    var khoXuat = normalizeStoreName(payload && payload.khoXuat ? payload.khoXuat : "");
    var khoNhan = normalizeStoreName(payload && payload.khoNhan ? payload.khoNhan : "");
    var updatesCount = Number(payload && payload.updatesCount ? payload.updatesCount : 0) || 0;
    var totalRows = Number(payload && payload.totalRows ? payload.totalRows : 0) || 0;
    var missingCount = Number(payload && payload.missingCount ? payload.missingCount : 0) || 0;
    var extraCount = Number(payload && payload.extraCount ? payload.extraCount : 0) || 0;
    var actor = String(payload && payload.actor ? payload.actor : "Chi nhánh");
    var packedPdfUrl = taoPdfDonHangVaLayLink(soPhieu);
    if (!khoNhan) khoNhan = getKhoNhanBySoPhieu(soPhieu);
    if (khoNhan) sendTelegramOrderReady(soPhieu, khoNhan, packedPdfUrl);
    var statusLabel = (totalRows > 0 && missingCount === 0 && extraCount === 0) ? "Đủ hàng" : (extraCount > 0 && missingCount > 0 ? "Thiếu và thừa hàng" : (extraCount > 0 ? "Thừa hàng" : "Thiếu hàng"));
    sendTelegramPackingSummary(soPhieu, khoXuat, khoNhan, updatesCount, totalRows, statusLabel, missingCount, extraCount, actor, packedPdfUrl);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// --- API: QUẢN LÝ PHIẾU (CÓ PHÂN QUYỀN) ---
function getNgayFilterBounds_(ngayFilter) {
  var filter = String(ngayFilter || "").trim().toLowerCase();
  if (!filter || filter === "all") return { filter: "all", startMs: null, endMs: null, maxScan: 6000 };
  var today = getScriptTodayStart_();
  if (!today) return { filter: filter, startMs: null, endMs: null, maxScan: 4000 };
  var dayMs = 24 * 60 * 60 * 1000;
  if (filter === "today" || /^\d{4}-\d{2}-\d{2}$/.test(filter)) {
    var day = filter === "today" ? today : parseDateInputYYYYMMDD(filter);
    if (!day) return { filter: filter, startMs: null, endMs: null, maxScan: 3000 };
    return { filter: filter, startMs: day.getTime(), endMs: day.getTime() + dayMs - 1, maxScan: 2500, exactDate: filter === "today" ? formatSheetDateYYYYMMDD(today) : filter };
  }
  if (filter === "yesterday") {
    var y = new Date(today);
    y.setDate(today.getDate() - 1);
    return { filter: filter, startMs: y.getTime(), endMs: y.getTime() + dayMs - 1, maxScan: 2500, exactDate: formatSheetDateYYYYMMDD(y) };
  }
  if (filter === "7days") {
    var start = new Date(today);
    start.setDate(today.getDate() - 6);
    return { filter: filter, startMs: start.getTime(), endMs: today.getTime() + dayMs - 1, maxScan: 4500 };
  }
  return { filter: filter, startMs: null, endMs: null, maxScan: 4000 };
}

function storeMatchesFast_(rowStore, targetStore) {
  if (!targetStore) return true;
  var left = String(rowStore || "").trim();
  var right = String(targetStore || "").trim();
  if (!left || !right) return false;
  if (left === right) return true;
  return isSameStoreName(left, right);
}

function layDanhSachPhieuTheoFilter(khoNhan, ngayYYYYMMDD, userRole, userStore) {
  var t0 = Date.now();
  var storeCalls = 0;
  var filterKhoNhan = "";
  var filterUserStore = "";
  var scanned = 0;
  var boundsFilter = "";
  try {
    var ss = getSS();
    var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
    if (!historySheet) {
      return { data: [], _debugTotalMs: Date.now() - t0, _debugRun: "ql-fast-v4", _debugRole: String(userRole || "") };
    }

    filterKhoNhan = normalizeStoreName(khoNhan || "");
    filterUserStore = normalizeStoreName(userStore || "");
    var isAdmin = String(userRole || "") === "Admin";
    var bounds = getNgayFilterBounds_(ngayYYYYMMDD);
    boundsFilter = bounds.filter || "";
    var lastRow = historySheet.getLastRow();
    if (lastRow < 2) {
      return { data: [], _debugTotalMs: Date.now() - t0, _debugRun: "ql-fast-v4", _debugRole: String(userRole || "") };
    }

    // Memo store-match — tránh gọi normalize/so khớp lặp trên cùng cặp tên kho
    var storeMemo = {};
    function matchStoreCounted_(rowStore, targetStore) {
      if (!targetStore) return true;
      storeCalls++;
      var left = String(rowStore || "");
      var key = left + "\0" + targetStore;
      if (Object.prototype.hasOwnProperty.call(storeMemo, key)) return storeMemo[key];
      var ok = storeMatchesFast_(left, targetStore);
      storeMemo[key] = ok;
      return ok;
    }

    var lastCol = Math.min(Math.max(historySheet.getLastColumn(), 16), 16);
    var chunkSize = 600;
    var maxScan = Math.min(bounds.maxScan || 4000, lastRow - 1);
    var map = {};
    var endRow = lastRow;
    var olderChunkStreak = 0;
    var tz = Session.getScriptTimeZone();
    var needKhoNhanFilter = !!(filterKhoNhan && filterKhoNhan !== "all");

    while (endRow >= 2 && scanned < maxScan) {
      var startRow = Math.max(2, endRow - chunkSize + 1);
      if (scanned + (endRow - startRow + 1) > maxScan) {
        startRow = Math.max(2, endRow - (maxScan - scanned) + 1);
      }
      var numRows = endRow - startRow + 1;
      var body = historySheet.getRange(startRow, 1, numRows, lastCol).getValues();
      scanned += numRows;

      var chunkHasInRange = false;
      var chunkAllOlder = bounds.startMs != null;
      for (var i = 0; i < body.length; i++) {
        var row = body[i];
        if (!row) continue;
        var rowSoPhieu = row[1] ? String(row[1]).trim() : "";
        if (!rowSoPhieu) continue;

        var rowNgay = row[0];
        var rowDateStr = "";
        var rowMs = null;
        if (rowNgay instanceof Date && !isNaN(rowNgay.getTime())) {
          rowMs = rowNgay.getTime();
          // Chỉ format chuỗi ngày khi cần exactDate
          if (bounds.exactDate) rowDateStr = Utilities.formatDate(rowNgay, tz, "yyyy-MM-dd");
        } else {
          rowDateStr = formatSheetDateYYYYMMDD(rowNgay);
          if (rowDateStr) {
            var parsed = parseDateInputYYYYMMDD(rowDateStr);
            if (parsed) rowMs = parsed.getTime();
          }
        }

        if (bounds.exactDate) {
          if (rowDateStr !== bounds.exactDate) {
            if (rowMs != null && bounds.startMs != null && rowMs < bounds.startMs) { /* older */ }
            else chunkAllOlder = false;
            continue;
          }
          chunkHasInRange = true;
          chunkAllOlder = false;
        } else if (bounds.startMs != null) {
          if (rowMs == null || rowMs < bounds.startMs || rowMs > bounds.endMs) {
            if (!(rowMs != null && rowMs < bounds.startMs)) chunkAllOlder = false;
            continue;
          }
          chunkHasInRange = true;
          chunkAllOlder = false;
        } else {
          chunkAllOlder = false;
          if (!matchesNgayFilter(rowNgay, ngayYYYYMMDD)) continue;
        }

        var rowKhoXuat = row[2] ? String(row[2]).trim() : "";
        var rowKhoNhan = row[3] ? String(row[3]).trim() : "";
        if (needKhoNhanFilter && !matchStoreCounted_(rowKhoNhan, filterKhoNhan)) continue;
        if (!isAdmin) {
          if (!matchStoreCounted_(rowKhoXuat, filterUserStore) && !matchStoreCounted_(rowKhoNhan, filterUserStore)) continue;
        }

        var displayStatus = getDisplayOrderStatus(row[12] ? String(row[12]).trim() : "Mới", row[8], row[15]);
        var thoiGian = rowMs != null ? rowMs : 0;
        var entry = map[rowSoPhieu];
        if (!entry) {
          map[rowSoPhieu] = {
            soPhieu: rowSoPhieu,
            khoXuat: rowKhoXuat,
            khoNhan: rowKhoNhan,
            thoiGian: thoiGian,
            trangThai: displayStatus === "Đã hủy dòng" ? "Mới" : displayStatus
          };
        } else {
          if (displayStatus === "Đã hủy") entry.trangThai = "Đã hủy";
          else if (displayStatus === "Đã xác nhận" && entry.trangThai !== "Đã hủy") entry.trangThai = "Đã xác nhận";
          else if (displayStatus === "Đã soạn" && entry.trangThai !== "Đã hủy" && entry.trangThai !== "Đã xác nhận") entry.trangThai = "Đã soạn";
          if (thoiGian && thoiGian > (entry.thoiGian || 0)) entry.thoiGian = thoiGian;
        }
      }

      if (bounds.startMs != null) {
        if (chunkHasInRange) olderChunkStreak = 0;
        else if (chunkAllOlder) {
          olderChunkStreak++;
          if (olderChunkStreak >= 2) break;
        } else {
          olderChunkStreak = 0;
        }
      }

      endRow = startRow - 1;
    }

    // Gom kết quả bằng vòng for (nhanh hơn map/filter trên GAS với dataset lớn)
    var res = [];
    for (var key in map) {
      if (Object.prototype.hasOwnProperty.call(map, key)) res.push(map[key]);
    }
    res.sort(function(a, b) { return (b.thoiGian || 0) - (a.thoiGian || 0); });
    return {
      data: res,
      _debugTotalMs: Date.now() - t0,
      _debugScanned: scanned,
      _debugFilter: boundsFilter,
      _debugStoreCalls: storeCalls,
      _debugRole: String(userRole || ""),
      _debugKhoNhan: filterKhoNhan,
      _debugUserStore: filterUserStore,
      _debugRun: "ql-fast-v4"
    };
  } catch (listErr) {
    return {
      data: [],
      error: String(listErr && listErr.message ? listErr.message : listErr),
      _debugTotalMs: Date.now() - t0,
      _debugScanned: scanned,
      _debugStoreCalls: storeCalls,
      _debugRole: String(userRole || ""),
      _debugKhoNhan: filterKhoNhan,
      _debugUserStore: filterUserStore,
      _debugRun: "ql-fast-v4-err"
    };
  }
}

function getDashboardTimelineBounds_(timeline, fromDate, toDate) {
  var today = getScriptTodayStart_();
  if (!today) return { startMs: null, endMs: null, maxScan: 5000 };
  var dayMs = 24 * 60 * 60 * 1000;
  var endMs = today.getTime() + dayMs - 1;
  var selected = String(timeline || "2days").trim();
  if (selected === "all") return { startMs: null, endMs: null, maxScan: 7000 };
  if (selected === "today") {
    return { startMs: today.getTime(), endMs: endMs, maxScan: 2500, exactDate: formatSheetDateYYYYMMDD(today) };
  }
  if (selected === "custom") {
    var start = fromDate ? parseDateInputYYYYMMDD(fromDate) : null;
    var end = toDate ? parseDateInputYYYYMMDD(toDate) : null;
    if (start && end && start.getTime() > end.getTime()) {
      var swap = start; start = end; end = swap;
    }
    return {
      startMs: start ? start.getTime() : null,
      endMs: end ? (end.getTime() + dayMs - 1) : endMs,
      maxScan: 5000
    };
  }
  var days = 2;
  if (selected === "7days") days = 7;
  else if (selected === "30days") days = 30;
  var rangeStart = new Date(today);
  rangeStart.setDate(today.getDate() - (days - 1));
  return {
    startMs: rangeStart.getTime(),
    endMs: endMs,
    maxScan: days <= 2 ? 3000 : (days <= 7 ? 4500 : 6000)
  };
}

function getDashboardSummary(userRole, userStore, timeline, fromDate, toDate) {
  try {
    var role = String(userRole || "");
    var store = String(userStore || "");
    var tl = String(timeline || "2days");
    var from = String(fromDate || "");
    var to = String(toDate || "");
    var cache = getScriptCache_();
    var cacheKey = "dash_sum_v2_" + [role, store, tl, from, to].join("|");
    try {
      var cachedRaw = cache.get(cacheKey);
      if (cachedRaw) {
        var cachedObj = JSON.parse(cachedRaw);
        if (cachedObj && cachedObj.success) return cachedObj;
      }
    } catch (cacheReadErr) {}

    var ss = getSS();
    var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
    var empty = { success: true, data: { totalOrders: 0, pendingOrders: 0, processedOrders: 0, canceledOrders: 0, recentOrders: [] } };
    if (!historySheet) return empty;

    var lastRow = historySheet.getLastRow();
    if (lastRow < 2) return empty;

    var bounds = getDashboardTimelineBounds_(tl, from, to);
    var isAdmin = role === "Admin";
    var filterStore = normalizeStoreName(store);
    var storeMemo = {};
    function matchDashStore_(rowStore) {
      if (isAdmin || !filterStore) return true;
      var left = String(rowStore || "");
      var key = left + "\0" + filterStore;
      if (Object.prototype.hasOwnProperty.call(storeMemo, key)) return storeMemo[key];
      // Giữ tương thích exact-match cũ, bổ sung storeMatchesFast_ khi tên lệch nhẹ
      var ok = left === filterStore || storeMatchesFast_(left, filterStore);
      storeMemo[key] = ok;
      return ok;
    }

    var lastCol = Math.min(Math.max(historySheet.getLastColumn(), 16), 16);
    var chunkSize = 600;
    var maxScan = Math.min(bounds.maxScan || 5000, lastRow - 1);
    var orderMap = {};
    var scanned = 0;
    var endRow = lastRow;
    var olderChunkStreak = 0;
    var tz = Session.getScriptTimeZone();

    while (endRow >= 2 && scanned < maxScan) {
      var startRow = Math.max(2, endRow - chunkSize + 1);
      if (scanned + (endRow - startRow + 1) > maxScan) {
        startRow = Math.max(2, endRow - (maxScan - scanned) + 1);
      }
      var numRows = endRow - startRow + 1;
      var body = historySheet.getRange(startRow, 1, numRows, lastCol).getValues();
      scanned += numRows;

      var chunkHasInRange = false;
      var chunkAllOlder = bounds.startMs != null;
      for (var i = 0; i < body.length; i++) {
        var row = body[i];
        if (!row) continue;
        var rowSoPhieu = row[1] ? String(row[1]).trim() : "";
        if (!rowSoPhieu) continue;

        var rowNgay = row[0];
        var rowMs = null;
        var rowDateStr = "";
        if (rowNgay instanceof Date && !isNaN(rowNgay.getTime())) {
          rowMs = rowNgay.getTime();
          if (bounds.exactDate) rowDateStr = Utilities.formatDate(rowNgay, tz, "yyyy-MM-dd");
        } else {
          rowDateStr = formatSheetDateYYYYMMDD(rowNgay);
          if (rowDateStr) {
            var parsed = parseDateInputYYYYMMDD(rowDateStr);
            if (parsed) rowMs = parsed.getTime();
          }
        }

        if (bounds.exactDate) {
          if (rowDateStr !== bounds.exactDate) {
            if (rowMs != null && bounds.startMs != null && rowMs < bounds.startMs) { /* older */ }
            else chunkAllOlder = false;
            continue;
          }
          chunkHasInRange = true;
          chunkAllOlder = false;
        } else if (bounds.startMs != null) {
          if (rowMs == null || rowMs < bounds.startMs || rowMs > bounds.endMs) {
            if (!(rowMs != null && rowMs < bounds.startMs)) chunkAllOlder = false;
            continue;
          }
          chunkHasInRange = true;
          chunkAllOlder = false;
        } else if (!isDateInTimeline(rowNgay, tl, from, to)) {
          chunkAllOlder = false;
          continue;
        }

        var rowKhoXuat = row[2] ? String(row[2]).trim() : "";
        var rowKhoNhan = row[3] ? String(row[3]).trim() : "";
        if (!matchDashStore_(rowKhoXuat) && !matchDashStore_(rowKhoNhan)) continue;

        var displayStatus = getDisplayOrderStatus(row[12] ? String(row[12]).trim() : "Mới", row[8], row[15]);
        var entry = orderMap[rowSoPhieu];
        if (!entry) {
          entry = {
            soPhieu: rowSoPhieu,
            khoXuat: rowKhoXuat,
            khoNhan: rowKhoNhan,
            thoiGian: rowNgay,
            thoiGianMs: rowMs || 0,
            status: displayStatus === "Đã hủy dòng" ? "Mới" : displayStatus,
            count: 0
          };
          orderMap[rowSoPhieu] = entry;
        }

        if (displayStatus === "Đã hủy") entry.status = "Đã hủy";
        else if (displayStatus === "Đã xác nhận") entry.status = "Đã xác nhận";
        else if (displayStatus === "Đã soạn" && entry.status !== "Đã hủy" && entry.status !== "Đã xác nhận") entry.status = "Đã soạn";
        else if (entry.status !== "Đã hủy" && entry.status !== "Đã soạn" && entry.status !== "Đã xác nhận") entry.status = "Mới";

        entry.count += 1;
        if (rowMs && rowMs > (entry.thoiGianMs || 0)) {
          entry.thoiGian = rowNgay;
          entry.thoiGianMs = rowMs;
        }
      }

      if (bounds.startMs != null) {
        if (chunkHasInRange) olderChunkStreak = 0;
        else if (chunkAllOlder) {
          olderChunkStreak++;
          if (olderChunkStreak >= 2) break;
        } else {
          olderChunkStreak = 0;
        }
      }
      endRow = startRow - 1;
    }

    var orders = [];
    for (var key in orderMap) {
      if (Object.prototype.hasOwnProperty.call(orderMap, key)) orders.push(orderMap[key]);
    }
    orders.sort(function(a, b) { return (b.thoiGianMs || 0) - (a.thoiGianMs || 0); });

    var totalOrders = orders.length;
    var pendingOrders = 0;
    var processedOrders = 0;
    var canceledOrders = 0;
    for (var j = 0; j < orders.length; j++) {
      var st = orders[j].status;
      if (st === "Đã hủy") canceledOrders++;
      else if (st === "Đã soạn" || st === "Đã xác nhận") processedOrders++;
      else pendingOrders++;
    }

    var recentOrders = [];
    var recentLimit = Math.min(8, orders.length);
    for (var r = 0; r < recentLimit; r++) {
      var order = orders[r];
      recentOrders.push({
        soPhieu: order.soPhieu,
        khoXuat: order.khoXuat,
        khoNhan: order.khoNhan,
        status: order.status,
        thoiGian: formatDateTime(order.thoiGian)
      });
    }

    var result = {
      success: true,
      data: {
        totalOrders: totalOrders,
        pendingOrders: pendingOrders,
        processedOrders: processedOrders,
        canceledOrders: canceledOrders,
        recentOrders: recentOrders
      }
    };
    try {
      var json = JSON.stringify(result);
      if (json.length < 90000) cache.put(cacheKey, json, 90);
    } catch (cacheWriteErr) {}
    return result;
  } catch (e) {
    return { success: false, error: e.message || String(e) };
  }
}

function formatDateTime(value) {
  if (!value) return "";
  try {
    var d = value instanceof Date ? value : new Date(value);
    return d.toLocaleString("vi-VN", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" });
  } catch (e) {
    return value;
  }
}

function isDateInTimeline(value, timeline, fromDate, toDate) {
  if (!value) return false;
  var dateStr = formatSheetDateYYYYMMDD(value);
  if (!dateStr) return false;
  var date = parseDateInputYYYYMMDD(dateStr);
  if (!date) return false;

  var today = getScriptTodayStart_();
  if (!today) return false;
  var selected = String(timeline || '2days').trim();
  if (selected === 'all') return true;
  if (selected === 'today') return date.getTime() === today.getTime();
  if (selected === 'custom') {
    var start = fromDate ? parseDateInputYYYYMMDD(fromDate) : null;
    var end = toDate ? parseDateInputYYYYMMDD(toDate) : null;
    if (start && end && start.getTime() > end.getTime()) {
      var swap = start;
      start = end;
      end = swap;
    }
    if (start && date.getTime() < start.getTime()) return false;
    if (end && date.getTime() > end.getTime()) return false;
    return !!(start || end);
  }

  var days = 2;
  if (selected === '7days') days = 7;
  else if (selected === '30days') days = 30;
  var rangeStart = new Date(today);
  rangeStart.setDate(today.getDate() - (days - 1));
  return date.getTime() >= rangeStart.getTime() && date.getTime() <= today.getTime();
}

function getChiTietPhieu(soPhieu, storeName, includeStock) {
  var t0 = Date.now();
  try {
    if (!soPhieu) return [];
    var ss = getSS();
    var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
    if (!historySheet) return [];
    var wantStock = !(includeStock === false || includeStock === 0 || includeStock === "0" || includeStock === "false");
    var selectedSet = buildOrderMatchSet_(soPhieu);
    // Chunk ngược theo match set — không full-scan lịch sử
    var pack = readHistoryForSelectedOrders_(historySheet, selectedSet, "", 4000);
    var data = pack.data || [[]];
    var sheetOrders = pack.orders || [];
    var catalogLookup = null;
    try { catalogLookup = getCatalogLookup(ss); } catch (catErr) { catalogLookup = null; }

    var matchedRows = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i]) continue;
      var rowSoPhieu = data[i][1] ? String(data[i][1]).trim() : "";
      if (!rowSoPhieu || !orderInMatchSet_(rowSoPhieu, selectedSet)) continue;

      var slGoc = Number(data[i][7]) || 0;
      var hasActualQty = (data[i][8] !== "" && data[i][8] !== undefined && data[i][8] !== null);
      var rowStatus = data[i][12] ? String(data[i][12]).trim() : "Mới";
      var isReceived = rowStatus === "Đã xác nhận nhận hàng";
      var rawSlGiao = data[i][15];
      var hasSlGiao = rawSlGiao !== "" && rawSlGiao !== null && rawSlGiao !== undefined;
      var slSoan = "";
      if (hasSlGiao) {
        slSoan = Number(rawSlGiao) || 0;
      } else if (hasActualQty && !isReceived) {
        slSoan = Number(data[i][8]) || 0;
      }
      var slThucTe = "";
      if (isReceived && hasActualQty) slThucTe = Number(data[i][8]);
      else if (!isReceived && hasActualQty && !hasSlGiao) slThucTe = Number(data[i][8]);

      matchedRows.push({
        rowIndex: sheetOrders[i - 1] || (pack.startRow + i - 1),
        maHang: data[i][4],
        maVach: data[i][5],
        tenHang: data[i][6],
        slGoc: slGoc,
        slSoan: slSoan,
        slThucTe: slThucTe,
        dvt: resolveDvtValue(catalogLookup, data[i][4], data[i][5], data[i][9] || "") || data[i][9] || "",
        ghiChu: data[i][11] || "",
        trangThai: rowStatus || "Mới",
        nguoiSoanHang: data[i][13] || "",
        stock: ""
      });
    }

    if (wantStock && matchedRows.length && isPackingQ7Store_(storeName)) {
      var tonKhoMap = getStockMapForStore(ss, storeName) || {};
      for (var j = 0; j < matchedRows.length; j++) {
        matchedRows[j].stock = getStockValueForItem(tonKhoMap, matchedRows[j].maHang, matchedRows[j].maVach, matchedRows[j].dvt);
      }
    }
    // Overlay tồn biến thể (TON_VARIANT) nếu SKU thuộc nhóm Parent — không đụng logic trừ tồn A/B/C
    if (wantStock && matchedRows.length) {
      try {
        var variantMap = readTonVariantMap_(ss);
        if (variantMap && Object.keys(variantMap).length) {
          for (var jv = 0; jv < matchedRows.length; jv++) {
            var vStock = getVariantStockIfPresent_(variantMap, matchedRows[jv].maHang, matchedRows[jv].maVach, matchedRows[jv].dvt);
            if (vStock !== null) matchedRows[jv].stock = vStock;
          }
        }
      } catch (varStockErr) {}
    }

    try {
      matchedRows._debugTotalMs = Date.now() - t0;
      matchedRows._debugScanned = pack.scannedRows || 0;
      matchedRows._debugStock = wantStock;
      matchedRows._debugRun = "ql-fast-v3";
    } catch (metaErr2) {}
    return matchedRows;
  } catch (err) {
    Logger.log("getChiTietPhieu error: " + (err.message || err));
    return [];
  }
}

function getStockMapForStore(ss, storeName) {
  if (!storeName) return {};
  // Ưu tiên sheet nhẹ TON_Q7 khi hỏi tồn kho Q7
  if (isPackingQ7Store_(storeName)) {
    var light = readTonKhoQ7Map_(ss);
    if (light && Object.keys(light).length) return light;
  }
  return getStockMapForStoreFromFullSheet_(ss, storeName);
}

function parseQuantityValue(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return isNaN(value) ? 0 : value;
  var text = String(value).trim();
  if (!text) return 0;
  var normalized = text.replace(/\s+/g, "").replace(/\u00A0/g, "");
  if (/^\d{1,3}(\.\d{3})+(,\d+)?$/.test(normalized)) {
    normalized = normalized.replace(/\./g, "").replace(/,/g, ".");
  } else if (/^\d{1,3}(,\d{3})+(\.\d+)?$/.test(normalized)) {
    normalized = normalized.replace(/,/g, "");
  } else if (/^\d+,\d+$/.test(normalized) && normalized.indexOf(".") === -1) {
    normalized = normalized.replace(/,/g, ".");
  } else {
    normalized = normalized.replace(/,/g, "");
  }
  var n = Number(normalized);
  return isNaN(n) ? 0 : n;
}

function normalizeProductCode(value) {
  if (value === null || value === undefined) return "";
  var text = String(value).trim().toUpperCase();
  if (!text) return "";

  if (text.charAt(0) === "'") text = text.slice(1);
  text = text.replace(/\u00A0/g, "").replace(/\s+/g, "");

  if (/^\d+\.0+$/.test(text)) {
    text = text.replace(/\.0+$/, "");
  }
  if (/^\d{1,3}(\.\d{3})+$/.test(text)) {
    text = text.replace(/\./g, "");
  }

  return text.replace(/[^A-Z0-9]/g, "");
}

function normalizeNumericCode(value) {
  var code = normalizeProductCode(value);
  if (!code) return "";
  if (!/^\d+$/.test(code)) return "";
  var compact = code.replace(/^0+/, "");
  return compact || "0";
}

function addStockValueByCode(tonKhoMap, prefix, code, ton, dvt) {
  var norm = normalizeProductCode(code);
  if (!norm) return;
  var qty = Number(ton) || 0;
  var dvtNorm = normalizeDvtKey_(dvt);
  var suffix = dvtNorm ? ("|DV:" + dvtNorm) : "";
  var key = prefix + norm + suffix;
  tonKhoMap[key] = (tonKhoMap[key] || 0) + qty;
  var compact = normalizeNumericCode(norm);
  if (compact && compact !== norm) {
    var compactKey = prefix + compact + suffix;
    tonKhoMap[compactKey] = (tonKhoMap[compactKey] || 0) + qty;
  }
}

function areCodesEquivalent(leftCode, rightCode) {
  var left = normalizeProductCode(leftCode);
  var right = normalizeProductCode(rightCode);
  if (!left || !right) return false;
  if (left === right) return true;
  var leftCompact = normalizeNumericCode(left);
  var rightCompact = normalizeNumericCode(right);
  return !!(leftCompact && rightCompact && leftCompact === rightCompact);
}

function lookupStockByPrefixCode_(tonKhoMap, prefix, code, dvt) {
  var norm = normalizeProductCode(code);
  if (!norm) return null;
  var dvtNorm = normalizeDvtKey_(dvt);
  var candidates = [];
  if (dvtNorm) {
    candidates.push(prefix + norm + "|DV:" + dvtNorm);
    var compactD = normalizeNumericCode(norm);
    if (compactD) candidates.push(prefix + compactD + "|DV:" + dvtNorm);
  }
  candidates.push(prefix + norm);
  var compact = normalizeNumericCode(norm);
  if (compact) candidates.push(prefix + compact);

  for (var i = 0; i < candidates.length; i++) {
    if (Object.prototype.hasOwnProperty.call(tonKhoMap, candidates[i])) {
      return Number(tonKhoMap[candidates[i]]) || 0;
    }
  }
  // Có ĐVT trên đơn nhưng không khớp ĐVT tồn → không lấy nhầm đơn vị khác
  if (dvtNorm) return 0;

  // Không có ĐVT trên đơn: cộng mọi biến thể ĐVT của mã (fallback)
  var sum = 0;
  var found = false;
  var bases = [prefix + norm];
  if (compact) bases.push(prefix + compact);
  for (var k in tonKhoMap) {
    if (!Object.prototype.hasOwnProperty.call(tonKhoMap, k) || k === "__meta") continue;
    for (var b = 0; b < bases.length; b++) {
      if (k === bases[b] || k.indexOf(bases[b] + "|DV:") === 0) {
        sum += Number(tonKhoMap[k]) || 0;
        found = true;
        break;
      }
    }
  }
  return found ? sum : null;
}

function getStockValueForItem(tonKhoMap, maHang, maVach, dvt) {
  if (!tonKhoMap) return 0;
  var byMaHang = lookupStockByPrefixCode_(tonKhoMap, "MH:", maHang, dvt);
  if (byMaHang !== null) return byMaHang;
  var byMaVach = lookupStockByPrefixCode_(tonKhoMap, "MV:", maVach, dvt);
  return byMaVach !== null ? byMaVach : 0;
}

/** Lấy nhãn ĐVT duy nhất từ TON_Q7 nếu đơn thiếu ĐVT */
function inferDvtLabelFromStockMap_(tonKhoMap, dvtLabels, maHang, maVach) {
  if (!tonKhoMap) return "";
  dvtLabels = dvtLabels || {};
  var codes = [];
  var mh = normalizeProductCode(maHang);
  var mv = normalizeProductCode(maVach);
  if (mh) codes.push("MH:" + mh);
  if (mv) codes.push("MV:" + mv);
  if (!codes.length) return "";
  var found = {};
  for (var k in tonKhoMap) {
    if (!Object.prototype.hasOwnProperty.call(tonKhoMap, k) || k === "__meta") continue;
    for (var c = 0; c < codes.length; c++) {
      if (k === codes[c] || k.indexOf(codes[c] + "|DV:") === 0) {
        var label = dvtLabels[k] || dvtFromStockKey_(k);
        if (label) found[label] = true;
      }
    }
  }
  var list = Object.keys(found);
  return list.length === 1 ? list[0] : "";
}

function isStoreNameMatch(stockStoreName, targetStoreName) {
  var left = normalizeStoreName(stockStoreName || "");
  var right = normalizeStoreName(targetStoreName || "");
  if (!left || !right) return false;

  if (left === right) return true;

  var leftNorm = normalizeHeaderText(left);
  var rightNorm = normalizeHeaderText(right);
  if (leftNorm && rightNorm && (leftNorm === rightNorm || leftNorm.indexOf(rightNorm) !== -1 || rightNorm.indexOf(leftNorm) !== -1)) {
    return true;
  }

  var leftShort = formatShortStoreLabel(left);
  var rightShort = formatShortStoreLabel(right);
  return normalizeHeaderText(leftShort) === normalizeHeaderText(rightShort);
}

function normalizeStoreName(storeName) {
  var raw = String(storeName || "").trim();
  if (!raw) return "";
  if (raw === "all" || raw === "Tất cả") return raw;
  var activeMap = getActiveStoreMap();
  if (activeMap[raw]) return raw;

  var normalizedRaw = normalizeHeaderText(raw);
  for (var fullName in activeMap) {
    if (!Object.prototype.hasOwnProperty.call(activeMap, fullName)) continue;
    var shortName = activeMap[fullName];
    if (normalizeHeaderText(fullName) === normalizedRaw || normalizeHeaderText(shortName) === normalizedRaw) {
      return fullName;
    }
  }
  return raw;
}

function isSameStoreName(leftStore, rightStore) {
  var left = normalizeStoreName(leftStore);
  var right = normalizeStoreName(rightStore);
  if (!left || !right) return false;
  if (left === right) return true;
  return isStoreNameMatch(left, right);
}

function ensureHistoryStatusColumn(historySheet) {
  if (historySheet.getRange(1, 13).getValue() !== "Trạng thái") {
    historySheet.getRange(1, 13).setValue("Trạng thái").setFontWeight("bold").setBackground("#d9ead3");
  }
  if (historySheet.getRange(1, 14).getValue() !== "Người cập nhật cuối") {
    historySheet.getRange(1, 14).setValue("Người cập nhật cuối").setFontWeight("bold").setBackground("#d9ead3");
  }
  if (historySheet.getRange(1, 15).getValue() !== "Nguồn cập nhật") {
    historySheet.getRange(1, 15).setValue("Nguồn cập nhật").setFontWeight("bold").setBackground("#d9ead3");
  }
  if (historySheet.getRange(1, 16).getValue() !== "SL Giao (Soạn)") {
    historySheet.getRange(1, 16).setValue("SL Giao (Soạn)").setFontWeight("bold").setBackground("#d9ead3");
  }
}

function getDisplayOrderStatus(rowStatus, slThucTe, slSoan) {
  var status = String(rowStatus || "").trim();
  if (status === "Đã hủy đơn") return "Đã hủy";
  // Ưu tiên trạng thái xác nhận — không bị cột 9 (SL nhận) kéo về "Đã soạn"
  if (status === "Đã xác nhận nhận hàng") return "Đã xác nhận";
  if (status === "Đã soạn hàng") return "Đã soạn";
  if (status === "Đã hủy dòng") return "Đã hủy dòng";
  // Suy ra "Đã soạn" CHỈ từ cột 16 (SL Giao/Soạn).
  // Không dùng cột 9: sau xác nhận cột 9 = SL nhận, dùng sẽ làm đơn đã nhận vẫn hiện ở tab Xác nhận.
  if (slSoan !== "" && slSoan !== undefined && slSoan !== null) return "Đã soạn";
  // Dữ liệu cũ: chưa có cột 16, SL soạn nằm cột 9 và status còn "Mới"
  if ((!status || status === "Mới") && slThucTe !== "" && slThucTe !== undefined && slThucTe !== null) {
    return "Đã soạn";
  }
  return "Mới";
}

function getOrderState(soPhieu, historySheet, dataRows) {
  var data = dataRows || historySheet.getDataRange().getValues();
  var target = String(soPhieu || "").trim().toLowerCase();
  var state = { isPacked: false, isConfirmed: false, stores: [] };
  if (!target) return state;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1] || "").trim().toLowerCase() !== target) continue;
    var rowStatus = String(data[i][12] || "").trim();
    var slThucTe = data[i][8];
    var khoXuat = String(data[i][2] || "").trim();
    var khoNhan = String(data[i][3] || "").trim();
    if (khoXuat && state.stores.indexOf(khoXuat) === -1) state.stores.push(khoXuat);
    if (khoNhan && state.stores.indexOf(khoNhan) === -1) state.stores.push(khoNhan);
    var slSoanCol = data[i][15];
    if (rowStatus === "Đã xác nhận nhận hàng") state.isConfirmed = true;
    if (rowStatus === "Đã soạn hàng" ||
        ((slSoanCol !== "" && slSoanCol !== undefined && slSoanCol !== null) && rowStatus !== "Đã hủy dòng" && rowStatus !== "Đã hủy đơn") ||
        ((slThucTe !== "" && slThucTe !== undefined && slThucTe !== null) && rowStatus !== "Đã hủy dòng" && rowStatus !== "Đã hủy đơn")) {
      state.isPacked = true;
    }
  }
  return state;
}

function assertActorCanManageOrder(actor, soPhieu, historySheet, dataRows) {
  var account = getAccountByActor(actor);
  if (!account) throw new Error("Tài khoản không tồn tại.");
  var state = getOrderState(soPhieu, historySheet, dataRows);
  if (state.isConfirmed) throw new Error("Đơn đã được xác nhận nhận hàng nên không thể chỉnh sửa nữa.");
  if (String(account.role).trim() === "Admin") {
    return { account: account, state: state };
  }
  var actorStore = String(account.store || "").trim();
  if (actorStore && actorStore !== "Tất cả" && state.stores.length && state.stores.indexOf(actorStore) === -1) {
    throw new Error("Bạn chỉ có thể chỉnh sửa đơn thuộc kho của mình.");
  }
  if (state.isPacked) throw new Error("Đơn đã soạn xong, chỉ quản trị viên mới được phép sửa.");
  return { account: account, state: state };
}

function getAuditSheet(ss) {
  var sheet = ss.getSheetByName("Lịch Sử Thay Đổi Đơn");
  if (!sheet) {
    sheet = ss.insertSheet("Lịch Sử Thay Đổi Đơn");
    sheet.appendRow(["Thời gian", "Số phiếu", "Hành động", "Người thực hiện", "Mã hàng hóa", "Mã vạch", "Giá trị cũ", "Giá trị mới", "Ghi chú"]);
    sheet.getRange("A1:I1").setFontWeight("bold").setBackground("#d9ead3");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getReceiveSheet(ss) {
  var sheet = ss.getSheetByName("Lịch Sử Nhận Hàng");
  if (!sheet) {
    sheet = ss.insertSheet("Lịch Sử Nhận Hàng");
    sheet.appendRow(["Thời gian", "Số phiếu", "Kho nhận", "Người xác nhận", "Mã hàng hóa", "Mã vạch", "Tên hàng", "Số lượng đã nhận", "Số lượng yêu cầu", "Ghi chú"]);
    sheet.getRange("A1:J1").setFontWeight("bold").setBackground("#d9ead3");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function logOrderChangesBatch_(ss, rows) {
  if (!rows || !rows.length) return;
  var sheet = getAuditSheet(ss);
  var start = Math.max(sheet.getLastRow() + 1, 2);
  sheet.getRange(start, 1, rows.length, 9).setValues(rows);
}

function logOrderChange(ss, soPhieu, action, actor, maHang, maVach, oldValue, newValue, note) {
  logOrderChangesBatch_(ss, [[new Date(), soPhieu, action, actor || "", maHang || "", maVach || "", oldValue || "", newValue || "", note || ""]]);
}

function getAccountByActor(actor) {
  var users = getDanhSachTaiKhoan();
  for (var i = 0; i < users.length; i++) {
    if (String(users[i].user).trim() === String(actor || "").trim()) return users[i];
  }
  return null;
}

function isAdminActor(actor) {
  var account = getAccountByActor(actor);
  return !!(account && String(account.role).trim() === "Admin");
}

function requireAuthenticatedAction(payload) {
  var actor = payload && payload.actor ? payload.actor : "";
  if (!actor) throw new Error("Thiếu thông tin người thực hiện.");
  var account = getAccountByActor(actor);
  if (!account) throw new Error("Tài khoản không tồn tại.");
}

function requireAdminAction(action, payload) {
  var adminActions = ['taoTaiKhoanMoi', 'nhapKhauCapNhatThongTin'];
  if (adminActions.indexOf(action) !== -1 && !isAdminActor(payload && payload.actor ? payload.actor : "")) {
    throw new Error("Chỉ quản trị viên được phép thực hiện thao tác này.");
  }
}

function requireAdmin(actor) {
  var account = getAccountByActor(actor);
  if (!account || String(account.role).trim() !== "Admin") throw new Error("Chỉ quản trị viên được phép thay đổi hoặc hủy đơn.");
}

function hasDuplicateItemInOrder(historySheet, soPhieu, item, dataRows) {
  if (!historySheet || !soPhieu || !item) return false;
  var data = dataRows || historySheet.getDataRange().getValues();
  var target = String(soPhieu).trim().toLowerCase();
  var itemMaHang = String(item.maHang || "").trim().toUpperCase();
  var itemMaVach = String(item.maVach || "").trim().toUpperCase();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() !== target) continue;
    var existingMaHang = String(data[i][4] || "").trim().toUpperCase();
    var existingMaVach = String(data[i][5] || "").trim().toUpperCase();
    if (itemMaHang && existingMaHang && itemMaHang === existingMaHang) return true;
    if (itemMaVach && existingMaVach && itemMaVach === existingMaVach) return true;
  }
  return false;
}

function isOrderConfirmedForEditing(soPhieu, historySheet, dataRows) {
  if (!soPhieu || !historySheet) return false;
  var data = dataRows || historySheet.getDataRange().getValues();
  var target = String(soPhieu).trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === target) {
      var status = String(data[i][12] || "").trim();
      if (status === "Đã xác nhận nhận hàng") {
        return true;
      }
    }
  }
  return false;
}

function luuChinhSuaPhieu(payload) {
  try {
    var ss = getSS();
    var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
    if (!historySheet) {
      return { success: false, error: "Không tìm thấy sheet Lịch Sử Xuất Kho", msg: "Không tìm thấy sheet Lịch Sử Xuất Kho" };
    }
    ensureHistoryStatusColumn(historySheet);
    var lock = LockService.getDocumentLock();
    try {
      lock.waitLock(10000);
      var historyData = historySheet.getDataRange().getValues();
      var changeCount = 0;
      var modifiedCount = 0;
      var cancelledCount = 0;
      var shouldNotify = false;
      var orderInfo = null;
      var soPhieu = payload && payload.soPhieu ? String(payload.soPhieu).trim() : "";
      var actor = (payload && payload.actor) ? payload.actor : "";
      var updates = (payload && payload.updates) ? payload.updates : [];
      var newItems = (payload && payload.newItems) ? payload.newItems : [];
      var auditRows = [];
      var nowAudit = new Date();

      var permission = assertActorCanManageOrder(actor, soPhieu, historySheet, historyData);
      var actorRole = String(permission.account.role || "").trim();
      var wasPacked = permission.state.isPacked;
      var orderBaseRow = null;
      var i;
      for (i = 1; i < historyData.length; i++) {
        if (String(historyData[i][1]).trim().toLowerCase() === soPhieu.toLowerCase()) {
          orderBaseRow = historyData[i];
          break;
        }
      }
      // rowPatches[sheetRow] = fields to write on cols 6-15 (batch)
      var rowPatches = {};
      function getPatch_(sheetRow) {
        if (!rowPatches[sheetRow]) rowPatches[sheetRow] = {};
        return rowPatches[sheetRow];
      }

      if (wasPacked && actorRole === "Admin") {
        for (var r = 1; r < historyData.length; r++) {
          if (String(historyData[r][1]).trim().toLowerCase() !== soPhieu.toLowerCase()) continue;
          var rowStatus = String(historyData[r][12] || "").trim();
          if (rowStatus === "Đã hủy đơn" || rowStatus === "Đã hủy dòng") continue;
          var resetPatch = getPatch_(r + 1);
          resetPatch.actualQty = "";
          resetPatch.note = "Cần soạn lại sau khi chỉnh sửa";
          resetPatch.status = "Mới";
          resetPatch.actor = actor;
          resetPatch.source = "Quản lý";
        }
      }

      var newRows = [];
      var catalogLookupForAdd = null;
      try { catalogLookupForAdd = getCatalogLookup(ss); } catch (catAddErr) { catalogLookupForAdd = null; }
      for (var n = 0; n < newItems.length; n++) {
        var newItem = newItems[n];
        if (hasDuplicateItemInOrder(historySheet, soPhieu, newItem, historyData)) {
          throw new Error("Mã này đã tồn tại trong đơn hiện tại. Không thể thêm dòng trùng.");
        }
        var itemQty = Number(newItem.sl);
        if (!itemQty || itemQty < 1) continue;
        newRows.push([
          new Date(), soPhieu,
          orderBaseRow && orderBaseRow[2] ? orderBaseRow[2] : "",
          orderBaseRow && orderBaseRow[3] ? orderBaseRow[3] : "",
          newItem.maHang || "", newItem.maVach || "", newItem.tenHang || "",
          itemQty, "", resolveDvtValue(catalogLookupForAdd, newItem.maHang || "", newItem.maVach || "", newItem.dvt || "") || "", "",
          "Thêm mới vào đơn", "Mới", actor, "Quản lý"
        ]);
        auditRows.push([nowAudit, soPhieu, "Thêm mã vào đơn", actor, newItem.maHang || "", newItem.maVach || "", "", itemQty, newItem.tenHang || ""]);
        changeCount += 1;
        shouldNotify = true;
      }

      for (i = 0; i < updates.length; i++) {
        var u = updates[i];
        var rowIndex = Number(u.row);
        if (!rowIndex || rowIndex < 2) continue;
        var currentRow = historyData[rowIndex - 1];
        if (!currentRow) continue;
        var oldSl = Number(currentRow[7]) || 0;
        var soPhieuValue = currentRow[1] ? String(currentRow[1]).trim() : soPhieu;
        var oldTenHang = currentRow[6] ? String(currentRow[6]).trim() : "";
        var maHang = currentRow[4] ? String(currentRow[4]).trim() : "";
        var maVach = currentRow[5] ? String(currentRow[5]).trim() : "";
        var oldDvt = currentRow[9] ? String(currentRow[9]).trim() : "";
        var newMaHang = u.valMaHang !== undefined && u.valMaHang !== null && String(u.valMaHang).trim() !== "" ? String(u.valMaHang).trim() : maHang;
        var newMaVach = u.valMaVach !== undefined && u.valMaVach !== null && String(u.valMaVach).trim() !== "" ? String(u.valMaVach).trim() : maVach;
        var newDvt = u.valDvt !== undefined && u.valDvt !== null && String(u.valDvt).trim() !== "" ? String(u.valDvt).trim() : oldDvt;
        var newTenHang = u.valTenHang !== undefined && u.valTenHang !== null && String(u.valTenHang).trim() !== "" ? String(u.valTenHang).trim() : oldTenHang;
        if (!orderInfo) orderInfo = getThongTinPhieu(soPhieuValue);

        if (u.valSl !== "" && Number(u.valSl) === 0) {
          var cancelPatch = getPatch_(rowIndex);
          cancelPatch.requestedQty = 0;
          cancelPatch.actualQty = 0;
          cancelPatch.note = "Đã hủy dòng";
          cancelPatch.status = "Đã hủy dòng";
          cancelPatch.actor = actor;
          cancelPatch.source = "Quản lý";
          if (newMaHang) cancelPatch.maHang = newMaHang;
          if (newMaVach) cancelPatch.maVach = newMaVach;
          if (newDvt) cancelPatch.dvt = newDvt;
          if (newTenHang) cancelPatch.tenHang = newTenHang;
          auditRows.push([nowAudit, soPhieuValue, "Hủy mã khỏi đơn", actor, maHang, maVach, oldSl, 0, "Hủy bằng cập nhật số lượng"]);
          changeCount += 1;
          cancelledCount += 1;
          shouldNotify = true;
        } else if (u.valSl !== "") {
          var newVal = Number(u.valSl);
          var editPatch = getPatch_(rowIndex);
          editPatch.requestedQty = newVal;
          editPatch.actualQty = wasPacked ? "" : currentRow[8];
          editPatch.note = wasPacked ? "Cần soạn lại sau khi chỉnh sửa" : "";
          editPatch.status = "Mới";
          editPatch.actor = actor;
          editPatch.source = "Quản lý";
          if (newMaHang) editPatch.maHang = newMaHang;
          if (newMaVach) editPatch.maVach = newMaVach;
          if (newDvt) editPatch.dvt = resolveDvtValue(catalogLookupForAdd, newMaHang || maHang, newMaVach, newDvt) || newDvt;
          if (newTenHang) editPatch.tenHang = newTenHang;

          if (Number(oldSl) !== newVal) {
            auditRows.push([nowAudit, soPhieuValue, "Sửa số lượng", actor, newMaHang || maHang, newMaVach || maVach, oldSl, newVal, ""]);
            changeCount += 1;
            modifiedCount += 1;
            shouldNotify = true;
          }
          if (String(newMaHang || "").trim() !== String(maHang || "").trim()) {
            auditRows.push([nowAudit, soPhieuValue, "Đổi mã hàng (biến thể)", actor, maHang, maVach, maHang, newMaHang, "Đổi mã con trong nhóm Parent"]);
            changeCount += 1;
            shouldNotify = true;
          }
          if (String(newMaVach || "").trim() !== String(maVach || "").trim()) {
            auditRows.push([nowAudit, soPhieuValue, "Đổi mã vạch", actor, newMaHang || maHang, maVach, maVach, newMaVach, "Đổi giữa mã lẻ/mã thùng/biến thể"]);
            changeCount += 1;
            shouldNotify = true;
          }
          if (String(newDvt || "").trim() !== String(oldDvt || "").trim()) {
            auditRows.push([nowAudit, soPhieuValue, "Đổi đơn vị tính", actor, newMaHang || maHang, newMaVach || maVach, oldDvt, newDvt, "Đổi ĐVT theo mã vạch"]);
            changeCount += 1;
            shouldNotify = true;
          }
          if (String(newTenHang || "").trim() !== String(oldTenHang || "").trim()) {
            auditRows.push([nowAudit, soPhieuValue, "Cập nhật tên hàng", actor, newMaHang || maHang, newMaVach || maVach, oldTenHang, newTenHang, "Đồng bộ theo mã biến thể"]);
            changeCount += 1;
            shouldNotify = true;
          }
        }
      }

      if (newRows.length) {
        var startRow = historySheet.getLastRow() + 1;
        historySheet.getRange(startRow, 1, newRows.length, 15).setValues(newRows);
      }

      var patchRowNums = [];
      for (var pk in rowPatches) {
        if (rowPatches.hasOwnProperty(pk)) patchRowNums.push(Number(pk));
      }
      if (patchRowNums.length) {
        patchRowNums.sort(function(a, b) { return a - b; });
        var minR = patchRowNums[0];
        var maxR = patchRowNums[patchRowNums.length - 1];
        var numRows = maxR - minR + 1;
        // cols 6-15: MaVach, Ten, SL, SLTT, DVT, Note, Status, Actor, Source
        // getRange(row, column, numRows, numColumns)
        var mat = historySheet.getRange(minR, 6, numRows, 10).getValues();
        var maHangColOut = [];
        var hasMaHangPatch = false;
        for (var pi = 0; pi < patchRowNums.length; pi++) {
          var sheetRow = patchRowNums[pi];
          var p = rowPatches[sheetRow];
          var off = sheetRow - minR;
          if (p.maVach !== undefined) mat[off][0] = p.maVach;
          if (p.tenHang !== undefined) mat[off][1] = p.tenHang;
          if (p.requestedQty !== undefined) mat[off][2] = p.requestedQty;
          if (p.actualQty !== undefined) mat[off][3] = p.actualQty;
          if (p.dvt !== undefined) mat[off][4] = p.dvt;
          if (p.note !== undefined) mat[off][6] = p.note;
          if (p.status !== undefined) mat[off][7] = p.status;
          if (p.actor !== undefined) mat[off][8] = p.actor;
          if (p.source !== undefined) mat[off][9] = p.source;
        }
        // Cột E (Mã hàng) — ghi riêng để đổi đúng mã con biến thể, không đụng batch A/B/C
        for (var mhPi = 0; mhPi < numRows; mhPi++) {
          var absRow = minR + mhPi;
          var pMh = rowPatches[absRow];
          if (pMh && pMh.maHang !== undefined) {
            maHangColOut.push([pMh.maHang]);
            hasMaHangPatch = true;
          } else {
            maHangColOut.push([historyData[absRow - 1] ? historyData[absRow - 1][4] : ""]);
          }
        }
        historySheet.getRange(minR, 6, numRows, 10).setValues(mat);
        if (hasMaHangPatch) {
          historySheet.getRange(minR, 5, numRows, 1).setValues(maHangColOut);
        }
      }

      logOrderChangesBatch_(ss, auditRows);
      SpreadsheetApp.flush();

      if (orderInfo && shouldNotify && changeCount > 0) {
        var detailText = [];
        if (modifiedCount > 0) detailText.push("Mã sửa số lượng: " + modifiedCount);
        if (cancelledCount > 0) detailText.push("Mã hủy: " + cancelledCount);
        var actionLabel = wasPacked && actorRole === "Admin" ? "Mở lại đơn đã soạn để chỉnh sửa" : "Chỉnh sửa số lượng / hủy mã";
        var extraSummary = detailText.join("; ") || ("Đơn đã thay đổi " + changeCount + " mã.");
        if (wasPacked && actorRole === "Admin") {
          extraSummary = "Trạng thái mới: Mới, cần soạn lại. " + extraSummary;
        }
        var editPdfUrl = taoPdfDonHangVaLayLink(orderInfo.soPhieu);
        sendTelegramOrderChangeSummary(orderInfo.soPhieu, orderInfo.khoXuat, orderInfo.khoNhan, actionLabel, changeCount, actor, extraSummary, editPdfUrl);
      }
    } finally {
      try { lock.releaseLock(); } catch (lockErr) {}
    }
    return { success: true };
  } catch (err) {
    return {
      success: false,
      error: err.message || String(err),
      msg: err.message || String(err)
    };
  }
}

function themChiTietPhieu(payload) {
  requireAuthenticatedAction(payload);
  var ss = getSS();
  var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
  if (!historySheet || !payload.soPhieu || !payload.item) throw new Error("Thiếu dữ liệu đơn hàng hoặc sản phẩm.");
  ensureHistoryStatusColumn(historySheet);
  var data = historySheet.getDataRange().getValues();
  assertActorCanManageOrder(payload.actor, payload.soPhieu, historySheet, data);
  var baseRow = null;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(payload.soPhieu).trim()) { baseRow = data[i]; break; }
  }
  if (!baseRow) throw new Error("Không tìm thấy đơn hàng.");
  var item = payload.item;
  var quantity = Number(item.sl);
  if (!quantity || quantity < 1) throw new Error("Số lượng thêm phải lớn hơn 0.");
  var dvtAdd = resolveDvtValue(getCatalogLookup(ss), item.maHang || "", item.maVach || "", item.dvt || "");
  var row = [new Date(), payload.soPhieu, baseRow[2], baseRow[3], item.maHang || "", item.maVach || "", item.tenHang || "", quantity, "", dvtAdd || "", "", "Thêm mới vào đơn", "Mới", payload.actor || "", "Quản lý"];
  historySheet.appendRow(row);
  logOrderChange(ss, payload.soPhieu, "Thêm mã vào đơn", payload.actor, item.maHang, item.maVach, "", quantity, item.tenHang || "");
  return { success: true };
}

function huyDongChiTietPhieu(payload) {
  try {
    requireAuthenticatedAction(payload);
    var ss = getSS();
    var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
    if (!historySheet) {
      return { success: false, error: "Không tìm thấy sheet Lịch Sử Xuất Kho", msg: "Không tìm thấy sheet Lịch Sử Xuất Kho" };
    }
    ensureHistoryStatusColumn(historySheet);
    var soPhieu = payload && payload.soPhieu ? payload.soPhieu : "";
    var dataRows = historySheet.getDataRange().getValues();
    if (!soPhieu) {
      var lookupRow = Number(payload.row);
      if (lookupRow >= 2 && dataRows[lookupRow - 1]) soPhieu = String(dataRows[lookupRow - 1][1] || "").trim();
    }
    assertActorCanManageOrder(payload.actor, soPhieu, historySheet, dataRows);
    var row = Number(payload.row);
    if (!row || row < 2) throw new Error("Dòng đơn hàng không hợp lệ.");
    var values = dataRows[row - 1];
    if (!values) throw new Error("Dòng đơn hàng không hợp lệ.");

    // Batch 1 lần: cột 8-15 (giữ nguyên ĐVT/ảnh ở cột 10-11)
    var block = historySheet.getRange(row, 8, 1, 8).getValues()[0];
    block[0] = 0;                         // cột 8 — SL đặt
    block[1] = 0;                         // cột 9 — SL thực tế
    block[4] = "Đã hủy dòng";             // cột 12
    block[5] = "Đã hủy dòng";             // cột 13
    block[6] = (payload && payload.actor) ? payload.actor : "";
    block[7] = "Quản lý";                 // cột 15
    historySheet.getRange(row, 8, 1, 8).setValues([block]);

    logOrderChange(ss, values[1], "Hủy mã khỏi đơn", payload.actor, values[4], values[5], values[7], 0, "Hủy từng dòng");
    return { success: true };
  } catch (err) {
    return {
      success: false,
      error: err.message || String(err),
      msg: err.message || String(err)
    };
  }
}

function huyPhieu(payload) {
  try {
    requireAdmin(payload.actor);
    var ss = getSS();
    var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
    if (!historySheet) {
      return { success: false, error: "Không tìm thấy sheet Lịch Sử Xuất Kho", msg: "Không tìm thấy sheet Lịch Sử Xuất Kho" };
    }
    ensureHistoryStatusColumn(historySheet);

    var data = historySheet.getDataRange().getValues();
    if (payload && payload.soPhieu && isOrderConfirmedForEditing(payload.soPhieu, historySheet, data)) {
      throw new Error("Đơn đã được xác nhận nhận hàng nên không thể hủy hoặc chỉnh sửa nữa.");
    }

    var target = String(payload && payload.soPhieu ? payload.soPhieu : "").trim();
    var matchRows = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === target) matchRows.push(i + 1);
    }
    if (!matchRows.length) throw new Error("Không tìm thấy đơn hàng.");

    var minR = matchRows[0];
    var maxR = matchRows[0];
    for (var m = 1; m < matchRows.length; m++) {
      if (matchRows[m] < minR) minR = matchRows[m];
      if (matchRows[m] > maxR) maxR = matchRows[m];
    }
    var numRows = maxR - minR + 1;
    var statusMat = historySheet.getRange(minR, 12, numRows, 2).getValues(); // cột 12-13
    for (var k = 0; k < matchRows.length; k++) {
      var off = matchRows[k] - minR;
      statusMat[off][0] = "Đã hủy đơn";
      statusMat[off][1] = "Đã hủy đơn";
    }
    historySheet.getRange(minR, 12, numRows, 2).setValues(statusMat);

    logOrderChange(ss, payload.soPhieu, "Hủy đơn", payload.actor, "", "", "Đang xử lý", "Đã hủy đơn", (payload && payload.reason) ? payload.reason : "");

    var orderInfo = getThongTinPhieu(payload.soPhieu);
    if (orderInfo) {
      sendTelegramOrderCancelled(payload.soPhieu, orderInfo.khoXuat, orderInfo.khoNhan, payload.actor, (payload && payload.reason) ? payload.reason : "");
    }
    return { success: true };
  } catch (err) {
    return {
      success: false,
      error: err.message || String(err),
      msg: err.message || String(err)
    };
  }
}

function xacNhanNhanHang(payload) {
  var t0 = Date.now();
  try {
    requireAuthenticatedAction(payload);
    var ss = getSS();
    var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
    if (!historySheet) {
      return { success: false, error: "Không tìm thấy dữ liệu đơn hàng.", msg: "Không tìm thấy dữ liệu đơn hàng." };
    }
    var receiveSheet = getReceiveSheet(ss);
    ensureHistoryStatusColumn(historySheet);

    var actorAccount = getAccountByActor(payload.actor);
    var expectedStore = payload.store || (actorAccount ? actorAccount.store : "");
    if (actorAccount && String(actorAccount.role).trim() !== "Admin") {
      if (!expectedStore || (String(actorAccount.store).trim() !== "Tất cả" && String(actorAccount.store).trim() !== String(expectedStore).trim())) {
        throw new Error("Bạn chỉ có thể xác nhận cho chi nhánh của mình.");
      }
    }

    var confirmations = (payload && payload.confirmations) ? payload.confirmations : [];
    if (!confirmations.length) throw new Error("Không có dữ liệu xác nhận.");
    var soPhieuTarget = String(payload.soPhieu || "").trim();
    var actor = (payload && payload.actor) ? payload.actor : "";

    var lock = LockService.getDocumentLock();
    try {
      lock.waitLock(20000);

      // --- Batch đọc các dòng confirm (1 lần getValues) ---
      var minRow = null;
      var maxRow = null;
      var c;
      for (c = 0; c < confirmations.length; c++) {
        var rowScan = Number(confirmations[c].row);
        if (!rowScan || rowScan < 2 || isNaN(rowScan)) continue;
        if (minRow === null || rowScan < minRow) minRow = rowScan;
        if (maxRow === null || rowScan > maxRow) maxRow = rowScan;
      }
      if (minRow === null) throw new Error("Không có dữ liệu xác nhận.");

      var numHistRows = maxRow - minRow + 1;
      var histMatrix = historySheet.getRange(minRow, 1, numHistRows, 16).getValues();
      var rowCache = {};
      for (c = 0; c < confirmations.length; c++) {
        var rowCheck = Number(confirmations[c].row);
        if (!rowCheck || rowCheck < 2 || rowCheck < minRow || rowCheck > maxRow) continue;
        rowCache[rowCheck] = histMatrix[rowCheck - minRow];
      }

      // --- Validate từ cache ---
      var invalidRows = [];
      for (c = 0; c < confirmations.length; c++) {
        var confCheck = confirmations[c];
        var rowV = Number(confCheck.row);
        if (!rowV || rowV < 2) continue;
        var checkValues = rowCache[rowV];
        if (!checkValues) {
          invalidRows.push(rowV + " (không đọc được dòng)");
          continue;
        }
        var rowSoPhieu = checkValues[1] ? String(checkValues[1]).trim() : "";
        var rowStatus = checkValues[12] ? String(checkValues[12]).trim() : "Mới";
        var hasCol9 = checkValues[8] !== "" && checkValues[8] !== null && checkValues[8] !== undefined;
        var hasCol16 = checkValues[15] !== "" && checkValues[15] !== null && checkValues[15] !== undefined;
        var hasPackedQty = hasCol9 || hasCol16;
        if (!rowSoPhieu || !orderKeysMatch_(rowSoPhieu, soPhieuTarget)) {
          invalidRows.push(rowV + " (sai số phiếu)");
          continue;
        }
        if (rowStatus === "Đã hủy dòng" || rowStatus === "Đã hủy đơn") {
          invalidRows.push(rowV + " (đã hủy)");
          continue;
        }
        if (rowStatus === "Đã xác nhận nhận hàng") {
          invalidRows.push(rowV + " (đã xác nhận)");
          continue;
        }
        if (!hasPackedQty && rowStatus !== "Đã soạn hàng") {
          invalidRows.push(rowV + " (chưa soạn)");
        }
      }
      if (invalidRows.length) {
        throw new Error("Không thể xác nhận đơn chưa soạn xong. Dòng lỗi: " + invalidRows.join(", "));
      }

      // --- Ghi lịch sử batch: cột 9 (SL nhận) + 13–15; cột 16 không đụng ---
      var confirmedTotal = 0;
      var changedCount = 0;
      var changedQtyTotal = 0;
      var receiveRows = [];
      var auditRows = [];
      var now = new Date();
      var khoXuatHint = "";
      var khoNhanHint = "";

      for (var i = 0; i < confirmations.length; i++) {
        var conf = confirmations[i];
        var row = Number(conf.row);
        var qty = Number(conf.receivedQty);
        if (!row || row < 2 || isNaN(qty) || qty < 0) continue;
        var values = rowCache[row];
        if (!values) continue;
        var off = row - minRow;
        histMatrix[off][8] = qty;                            // cột 9
        histMatrix[off][12] = "Đã xác nhận nhận hàng";       // cột 13
        histMatrix[off][13] = actor;                         // cột 14
        histMatrix[off][14] = "Xác nhận nhận hàng";          // cột 15
        if (!khoXuatHint && values[2]) khoXuatHint = String(values[2]).trim();
        if (!khoNhanHint && values[3]) khoNhanHint = String(values[3]).trim();
        var previousQty = Number(conf.previousQty) || 0;
        if (qty !== previousQty) {
          changedCount += 1;
          changedQtyTotal += Math.abs(qty - previousQty);
        }
        receiveRows.push([now, soPhieuTarget, expectedStore, actor, values[4] || "", values[5] || "", values[6] || "", qty, values[7] || 0, "Đã xác nhận nhận hàng"]);
        auditRows.push([now, soPhieuTarget, "Xác nhận nhận hàng", actor, values[4] || "", values[5] || "", values[7] || "", qty, "Xác nhận bởi chi nhánh"]);
        confirmedTotal += qty;
      }

      historySheet.getRange(minRow, 1, numHistRows, 16).setValues(histMatrix);

      if (receiveRows.length) {
        var rStart = Math.max(receiveSheet.getLastRow() + 1, 2);
        receiveSheet.getRange(rStart, 1, receiveRows.length, 10).setValues(receiveRows);
      }
      if (auditRows.length) {
        logOrderChangesBatch_(ss, auditRows);
      }

      applyStockDeductionAfterReceive(historySheet, confirmations, soPhieuTarget, rowCache, ss);
      SpreadsheetApp.flush();

      return {
        success: true,
        count: confirmations.length,
        _debugTotalMs: Date.now() - t0,
        _debugRun: "confirm-batch-v2",
        notify: {
          soPhieu: soPhieuTarget,
          khoXuat: khoXuatHint,
          khoNhan: khoNhanHint || expectedStore,
          actor: actor,
          count: confirmations.length,
          confirmedTotal: confirmedTotal,
          changedCount: changedCount,
          changedQtyTotal: changedQtyTotal
        }
      };
    } finally {
      try { lock.releaseLock(); } catch (lockErr) {}
    }
  } catch (err) {
    return {
      success: false,
      error: err.message || String(err),
      msg: err.message || String(err),
      _debugTotalMs: Date.now() - t0,
      _debugRun: "confirm-batch-v2-catch"
    };
  }
}

/** PDF + Telegram sau khi đã lưu xác nhận (không chặn bước lưu) */
function postProcessReceiveOrder(payload) {
  var soPhieu = String(payload && payload.soPhieu ? payload.soPhieu : "").trim();
  if (!soPhieu) return { success: false, error: "Thiếu số phiếu." };
  try {
    var khoNhan = normalizeStoreName(payload.khoNhan || "");
    var actor = String(payload.actor || "");
    var count = Number(payload.count) || 0;
    var confirmedTotal = Number(payload.confirmedTotal) || 0;
    var changedCount = Number(payload.changedCount) || 0;
    var changedQtyTotal = Number(payload.changedQtyTotal) || 0;
    if (!khoNhan) {
      var info = getThongTinPhieu(soPhieu);
      if (info && info.khoNhan) khoNhan = info.khoNhan;
    }
    var receivePdfUrl = taoPdfDonHangVaLayLink(soPhieu);
    sendTelegramReceiveConfirmation(soPhieu, khoNhan, actor, count, confirmedTotal, changedCount, changedQtyTotal, receivePdfUrl);
    return { success: true, _debugRun: "confirm-post-v1" };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function applyStockDeductionAfterReceive(historySheet, confirmations, soPhieu, rowCache, ssOpt) {
  try {
    var ss = ssOpt || getSS();
    var stockSheet = ss.getSheetByName("TỔNG HỢP TỒN KHO");
    if (!stockSheet) return;

    var stockData = stockSheet.getDataRange().getValues();
    var stockConfig = getStockSheetConfig(stockData);
    var tonIdx = stockConfig.tonKhoIdx;

    // Index tồn 1 lần: mã hàng / mã vạch → danh sách dòng (kèm kho), tránh O(confirm × stockRows) quét sheet
    var byMaHang = {};
    var byMaVach = {};
    var currentMaHang = "";
    var currentMaVach = "";
    for (var k = stockConfig.startRow; k < stockData.length; k++) {
      var rowStock = stockData[k];
      if (!rowStock) continue;
      var rowMaHangRaw = getCellValue(rowStock, stockConfig.maHangIdx, "");
      var rowMaVachRaw = getCellValue(rowStock, stockConfig.maVachIdx, "");
      var hasOwnCode = !!(rowMaHangRaw || rowMaVachRaw);
      if (hasOwnCode) {
        currentMaHang = rowMaHangRaw;
        currentMaVach = rowMaVachRaw;
      }
      var rowStores = getRowStoreNames(rowStock, stockConfig);
      if (!rowStores.length) continue;

      var resolvedMaHang = normalizeProductCode((hasOwnCode ? rowMaHangRaw : currentMaHang) || "");
      var resolvedMaVach = normalizeProductCode((hasOwnCode ? rowMaVachRaw : currentMaVach) || "");
      if (!resolvedMaHang && !resolvedMaVach) continue;

      var entry = {
        dataIndex: k,
        sheetRow: k + 1,
        maHang: resolvedMaHang,
        maVach: resolvedMaVach,
        stores: rowStores
      };
      if (resolvedMaHang) {
        if (!byMaHang[resolvedMaHang]) byMaHang[resolvedMaHang] = [];
        byMaHang[resolvedMaHang].push(entry);
      }
      if (resolvedMaVach) {
        if (!byMaVach[resolvedMaVach]) byMaVach[resolvedMaVach] = [];
        byMaVach[resolvedMaVach].push(entry);
      }
    }

    function findStockEntry_(khoXuat, maHang, maVach) {
      var candidates = [];
      var seen = {};
      function addPool_(list) {
        if (!list) return;
        for (var li = 0; li < list.length; li++) {
          var ent = list[li];
          if (seen[ent.sheetRow]) continue;
          seen[ent.sheetRow] = true;
          candidates.push(ent);
        }
      }
      if (maHang) {
        if (byMaHang[maHang]) addPool_(byMaHang[maHang]);
        for (var hk in byMaHang) {
          if (!byMaHang.hasOwnProperty(hk) || hk === maHang) continue;
          if (areCodesEquivalent(maHang, hk)) addPool_(byMaHang[hk]);
        }
      }
      if (maVach) {
        if (byMaVach[maVach]) addPool_(byMaVach[maVach]);
        for (var vk in byMaVach) {
          if (!byMaVach.hasOwnProperty(vk) || vk === maVach) continue;
          if (areCodesEquivalent(maVach, vk)) addPool_(byMaVach[vk]);
        }
      }
      candidates.sort(function(a, b) { return a.sheetRow - b.sheetRow; });
      for (var ci = 0; ci < candidates.length; ci++) {
        var hit = candidates[ci];
        var codeMatch =
          (maHang && hit.maHang && areCodesEquivalent(maHang, hit.maHang)) ||
          (maVach && hit.maVach && areCodesEquivalent(maVach, hit.maVach));
        if (!codeMatch) continue;
        for (var s = 0; s < hit.stores.length; s++) {
          if (isStoreNameMatch(hit.stores[s], khoXuat)) return hit;
        }
      }
      return null;
    }

    var stockUpdates = {}; // sheetRow -> newQty
    for (var i = 0; i < confirmations.length; i++) {
      var conf = confirmations[i];
      var row = Number(conf.row);
      var qty = Number(conf.receivedQty);
      if (!row || row < 2 || isNaN(qty) || qty <= 0) continue;

      var orderRow = (rowCache && rowCache[row]) ? rowCache[row] : null;
      if (!orderRow && historySheet) {
        orderRow = historySheet.getRange(row, 1, 1, 9).getValues()[0];
      }
      if (!orderRow) continue;

      var rowSoPhieu = orderRow[1] ? String(orderRow[1]).trim() : "";
      if (!rowSoPhieu || !orderKeysMatch_(rowSoPhieu, soPhieu || "")) continue;

      var khoXuat = orderRow[2] ? String(orderRow[2]).trim() : "";
      var maHang = orderRow[4] ? normalizeProductCode(orderRow[4]) : "";
      var maVach = orderRow[5] ? normalizeProductCode(orderRow[5]) : "";
      if (!khoXuat || (!maHang && !maVach)) continue;

      var hit = findStockEntry_(khoXuat, maHang, maVach);
      if (!hit) continue;

      var oldStock = parseQuantityValue(stockData[hit.dataIndex][tonIdx]);
      var newStock = oldStock - qty;
      stockData[hit.dataIndex][tonIdx] = newStock;
      stockUpdates[hit.sheetRow] = newStock;
    }

    var updateRows = [];
    for (var ur in stockUpdates) {
      if (stockUpdates.hasOwnProperty(ur)) updateRows.push(Number(ur));
    }
    if (!updateRows.length) return;

    updateRows.sort(function(a, b) { return a - b; });
    var minR = updateRows[0];
    var maxR = updateRows[updateRows.length - 1];
    var numRows = maxR - minR + 1;
    var tonCol = tonIdx + 1;
    var tonMat = stockSheet.getRange(minR, tonCol, numRows, 1).getValues();
    for (var w = 0; w < updateRows.length; w++) {
      tonMat[updateRows[w] - minR][0] = stockUpdates[updateRows[w]];
    }
    stockSheet.getRange(minR, tonCol, numRows, 1).setValues(tonMat);
  } catch (err) {
    Logger.log("applyStockDeductionAfterReceive error: " + (err.message || err));
    throw err;
  }
}

function cleanupLegacyGeneratedSheets(ss, prefixes) {
  if (!ss || typeof ss.getSheets !== "function") return;
  var allSheets = ss.getSheets();
  for (var i = allSheets.length - 1; i >= 0; i--) {
    var sheet = allSheets[i];
    var name = sheet.getName ? sheet.getName() : "";
    for (var p = 0; p < prefixes.length; p++) {
      if (name.indexOf(prefixes[p]) === 0) {
        ss.deleteSheet(sheet);
        break;
      }
    }
  }
}

function recreateTempSheet(ss, sheetName, legacyPrefixes) {
  cleanupLegacyGeneratedSheets(ss, legacyPrefixes || []);
  var existing = ss.getSheetByName(sheetName);
  if (existing) {
    existing.clear();
    return existing;
  }
  return ss.insertSheet(sheetName);
}

function recreateTempSheetFast_(ss, sheetName) {
  var existing = ss.getSheetByName(sheetName);
  if (existing) {
    existing.clear();
    return existing;
  }
  return ss.insertSheet(sheetName);
}

function taoFileExcelVaLayLink(payload) {
  var ss = getSS();
  var catalogLookup = getCatalogLookup(ss);
  var tenTabPhieu = "__TMP_XUAT_EXCEL";
  var targetSheet = recreateTempSheet(ss, tenTabPhieu, ["In_"]);
  
  var khoXuat = payload.khoXuat;
  var khoNhan = payload.khoNhan;
  var tieuDe = payload.soPhieu.indexOf("DH") !== -1 ? "ĐƠN ĐẶT HÀNG" : "LỆNH ĐIỀU CHUYỂN";

  var finalItems = [];
  var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
  if (historySheet && payload.soPhieu) {
    var historyData = historySheet.getDataRange().getValues();
    var targetPhieu = String(payload.soPhieu).trim().toLowerCase();
    for (var i = 1; i < historyData.length; i++) {
      var row = historyData[i];
      if (!row) continue;
      var rowSoPhieu = row[1] ? String(row[1]).trim().toLowerCase() : "";
      if (rowSoPhieu !== targetPhieu) continue;

      var rowStatus = row[12] ? String(row[12]).trim() : "Mới";
      if (rowStatus === "Đã hủy dòng" || rowStatus === "Đã hủy đơn") continue;

      if (!khoXuat && row[2]) khoXuat = String(row[2]).trim();
      if (!khoNhan && row[3]) khoNhan = String(row[3]).trim();

      var slDat = Number(row[7]) || 0;
      var hasActualQty = row[8] !== "" && row[8] !== null && row[8] !== undefined;
      var isReceived = rowStatus === "Đã xác nhận nhận hàng";
      var rawSlGiao = row[15];
      var hasSlGiao = rawSlGiao !== "" && rawSlGiao !== null && rawSlGiao !== undefined;
      // Ưu tiên SL Soạn (cột 16); rồi số thực tế khi đang ở bước soạn; cuối cùng SL đặt
      var slFinal;
      if (hasSlGiao) slFinal = Number(rawSlGiao) || 0;
      else if (hasActualQty && !isReceived) slFinal = Number(row[8]) || 0;
      else if (hasActualQty && isReceived) slFinal = Number(row[8]) || 0; // nhận rồi mà chưa có cột 16 → dùng cột 9
      else slFinal = slDat;
      if (!slFinal || slFinal <= 0) continue;

      finalItems.push({
        maHang: row[4] || "",
        maVach: row[5] || "",
        tenHang: row[6] || "",
        dvt: resolveDvtValue(catalogLookup, row[4], row[5], row[9]),
        sl: slFinal
      });
    }
  }

  // Fallback for legacy calls where order has not been read back from history yet.
  if (!finalItems.length && payload.items && payload.items.length) {
    for (var f = 0; f < payload.items.length; f++) {
      var pItem = payload.items[f];
      var qty = Number(pItem.sl);
      if (!qty || qty <= 0) continue;
      finalItems.push({
        maHang: pItem.maHang || "",
        maVach: pItem.maVach || "",
        tenHang: pItem.tenHang || "",
        dvt: resolveDvtValue(catalogLookup, pItem.maHang, pItem.maVach, pItem.dvt),
        sl: qty
      });
    }
  }
  
  targetSheet.getRange("A4:F4").merge().setValue(tieuDe).setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
  targetSheet.getRange("A6:F6").merge().setValue("Số: " + payload.soPhieu).setFontStyle("italic").setHorizontalAlignment("center");
  targetSheet.getRange("A8").setValue("Kho xuất:").setFontWeight("bold"); targetSheet.getRange("B8").setValue(khoXuat);
  targetSheet.getRange("A9").setValue("Kho nhận:").setFontWeight("bold"); targetSheet.getRange("B9").setValue(khoNhan);
  
  var headers = ["STT", "Mã hàng hóa", "Mã vạch", "Tên hàng hóa", "ĐVT", "Số lượng (Soạn)"];
  targetSheet.getRange("A12:F12").setValues([headers]).setFontWeight("bold").setHorizontalAlignment("center").setBackground("#f8f9fa");
  
  var dataArr = []; var stt = 1;
  for (var j = 0; j < finalItems.length; j++) {
    // Cột MH/MV = mã con chuẩn; cột tên kèm chi tiết biến thể (phẳng, không gộp cha)
    var fi = finalItems[j];
    dataArr.push([
      stt++,
      fi.maHang || "",
      fi.maVach || "",
      formatVariantDisplayName_(fi.maHang, fi.tenHang),
      fi.dvt || "",
      fi.sl
    ]);
  }
  
  if(dataArr.length > 0) {
    targetSheet.getRange(13, 1, dataArr.length, 6).setValues(dataArr);
    targetSheet.getRange(12, 1, dataArr.length + 1, 6).setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
    targetSheet.setColumnWidth(1, 45); targetSheet.setColumnWidth(2, 110); targetSheet.setColumnWidth(3, 120);
    targetSheet.setColumnWidth(4, 380); targetSheet.setColumnWidth(5, 60); targetSheet.setColumnWidth(6, 60);
  }
  SpreadsheetApp.flush();
  return { success: true, url: "https://docs.google.com/spreadsheets/d/" + SHEET_ID + "/export?format=xlsx&gid=" + targetSheet.getSheetId() };
}

function taoPreviewPublicUrl(soPhieu, khoXuat, khoNhan, rows) {
  var ss = getSS();
  var previewSheetName = "__TMP_PREVIEW_CONG_KHAI";
  var sheet = recreateTempSheet(ss, previewSheetName, ["Preview_"]);
  sheet.getRange("A1:G1").setValues([["Số phiếu", "Kho xuất", "Kho nhận", "Mã hàng", "Mã vạch", "Tên hàng", "Số thực nhận"]]).setFontWeight("bold");
  var data = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    data.push([soPhieu, khoXuat, khoNhan, r.maHang || "", r.maVach || "", r.tenHang || "", r.slThucTe !== undefined && r.slThucTe !== null && r.slThucTe !== "" ? r.slThucTe : r.slGoc || 0]);
  }
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, 7).setValues(data);
  }
  sheet.autoResizeColumns(1, 7);
  SpreadsheetApp.flush();
  var url = "https://docs.google.com/spreadsheets/d/" + SHEET_ID + "/edit#gid=" + sheet.getSheetId();
  return { success: true, url: url };
}

// --- API: SOẠN HÀNG MOBILE ---
function getDonHangTheoNgay(ngayChon, userRole, userStore, viewMode) {
  var ss = getSS();
  var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
  if (!historySheet) return [];
  var filterUserStore = normalizeStoreName(userStore || "");
  var data = readHistoryDataPack_(historySheet).data;
  var map = {};
  
  for (var i = 1; i < data.length; i++) {
    var rowNgay = data[i][0];
    var rowSoPhieu = data[i][1] ? data[i][1].toString().trim() : "";
    var rowKhoXuat = data[i][2] ? data[i][2].toString().trim() : "";
    var rowKhoNhan = data[i][3] ? data[i][3].toString().trim() : "";
    var rowStatus = data[i][12] ? String(data[i][12]).trim() : "Mới";
    var isCanceled = rowStatus === "Đã hủy đơn";

    if (isCanceled) continue;

    if (viewMode === "packing") {
      if (userRole !== "Admin" && !isSameStoreName(rowKhoXuat, filterUserStore)) continue;
    } else if (userRole !== "Admin") {
      if (!isSameStoreName(rowKhoXuat, filterUserStore) && !isSameStoreName(rowKhoNhan, filterUserStore)) continue;
    }

    var slThucTe = data[i][8];
    var displayStatus = getDisplayOrderStatus(rowStatus, slThucTe, data[i][15]);

    if (!matchesNgayFilter(rowNgay, ngayChon)) continue;

    if (rowSoPhieu) { 
      if(!map[rowSoPhieu]) map[rowSoPhieu] = { soPhieu: rowSoPhieu, khoXuat: rowKhoXuat, khoNhan: rowKhoNhan, trangThai: displayStatus === "Đã hủy dòng" ? "Mới" : displayStatus }; 
      else if(displayStatus === "Đã xác nhận") map[rowSoPhieu].trangThai = "Đã xác nhận";
      else if(displayStatus === "Đã soạn" && map[rowSoPhieu].trangThai !== "Đã xác nhận") map[rowSoPhieu].trangThai = "Đã soạn";
    }
  }
  var res = []; for(var p in map) res.push(map[p]); return res;
}

function getChiTietDonHangMobile(soPhieu) {
  var ss = getSS();
  var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
  if (!historySheet || !soPhieu) return [];
  ensureHistoryStatusColumn(historySheet);

  var selectedSet = buildOrderMatchSet_(soPhieu);
  var pack = readHistoryForSelectedOrders_(historySheet, selectedSet, "", 6000);
  var data = pack.data || [[]];
  var sheetOrders = pack.orders || [];

  var catalogLookup = null;
  try { catalogLookup = getCatalogLookup(ss); } catch (catErr) { catalogLookup = null; }

  var items = [];
  var khoXuat = "";
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row) continue;
    var rowSoPhieu = row[1] ? String(row[1]).trim() : "";
    if (!rowSoPhieu || !orderInMatchSet_(rowSoPhieu, selectedSet)) continue;
    if (!khoXuat && row[2]) khoXuat = String(row[2]).trim();
    var slGoc = Number(row[7]) || 0;
    var rowStatus = row[12] ? String(row[12]).trim() : "Mới";
    var isReceived = rowStatus === "Đã xác nhận nhận hàng";
    var hasActualQty = row[8] !== "" && row[8] !== undefined && row[8] !== null;
    // Cột 16 = SL đã soạn; cột 9 = thực nhận (sau xác nhận) hoặc SL soạn (dữ liệu cũ)
    var rawSlGiao = row[15];
    var hasSlGiao = rawSlGiao !== "" && rawSlGiao !== null && rawSlGiao !== undefined;
    var slSoan = "";
    if (hasSlGiao) {
      slSoan = Number(rawSlGiao) || 0;
    } else if (hasActualQty && !isReceived) {
      slSoan = Number(row[8]) || 0;
    }
    var packInputQty = slSoan !== "" ? Number(slSoan) : (hasActualQty && !isReceived ? Number(row[8]) : slGoc);
    var slThucTe = isReceived && hasActualQty ? Number(row[8]) : packInputQty;
    var dvtVal = row[9] || "";
    if (catalogLookup) {
      try { dvtVal = resolveDvtValue(catalogLookup, row[4], row[5], row[9]) || dvtVal; } catch (dvtErr) {}
    }
    items.push({
      rowIndex: sheetOrders[i - 1] || (pack.startRow + i - 1),
      maHang: row[4],
      maVach: row[5],
      tenHang: row[6],
      dvt: dvtVal,
      slGoc: slGoc,
      slSoan: slSoan,
      slThucTe: slThucTe,
      trangThai: rowStatus,
      anhXacNhan: (row[10] || ""),
      nguoiSoanHang: row[13] || "",
      _debugPackQty: packInputQty,
      _debugHasCol16: hasSlGiao
    });
  }

  try {
    var tonKhoMap = getStockMapForStore(ss, khoXuat) || {};
    for (var j = 0; j < items.length; j++) {
      items[j].stock = getStockValueForItem(tonKhoMap, items[j].maHang, items[j].maVach, items[j].dvt);
    }
  } catch (stockErr) {
    for (var k = 0; k < items.length; k++) items[k].stock = "";
  }
  return items;
}

/** Debug đơn theo key/suffix (vd. 318957 hoặc Q7-DC318957) */
function debugOrderInfo_(key) {
  key = String(key || "").trim();
  var digits = key.replace(/\D/g, "");
  var ss = getSS();
  var sh = ss.getSheetByName("Lịch Sử Xuất Kho");
  if (!sh) return { success: false, error: "no history sheet" };
  var lastRow = sh.getLastRow();
  var lastCol = Math.max(sh.getLastColumn(), 16);
  var start = Math.max(2, lastRow - 5000);
  var num = lastRow - start + 1;
  var values = sh.getRange(start, 1, num, lastCol).getValues();
  var found = [];
  var minCreatedMs = NaN;
  var tz = Session.getScriptTimeZone() || "Asia/Ho_Chi_Minh";
  for (var i = 0; i < values.length; i++) {
    var sp = values[i][1] != null ? String(values[i][1]).trim() : "";
    if (!sp) continue;
    var hit = false;
    if (key && (sp.indexOf(key) !== -1 || orderKeysMatch_(sp, key))) hit = true;
    if (!hit && digits && sp.indexOf(digits) !== -1) hit = true;
    if (!hit) continue;
    var createdMs = toHoChiMinhMillis_(values[i][0]);
    if (!isNaN(createdMs) && (isNaN(minCreatedMs) || createdMs < minCreatedMs)) minCreatedMs = createdMs;
    found.push({
      sheetRow: start + i,
      soPhieu: sp,
      ngayTao: values[i][0] instanceof Date
        ? Utilities.formatDate(values[i][0], tz, "dd/MM/yyyy HH:mm:ss")
        : String(values[i][0] || ""),
      createdMs: createdMs,
      khoXuat: values[i][2],
      khoNhan: values[i][3],
      slGoc: values[i][7],
      col9: values[i][8],
      status: values[i][12],
      col16_slSoan: values[i][15],
      actor: values[i][13],
      source: values[i][14]
    });
  }

  var packingProbe = null;
  if (!isNaN(minCreatedMs)) {
    // Đoán N2: tạo >= 10:00 → N2 = ngày tạo + 1; tạo < 10:00 → N2 = ngày tạo
    var createdLocal = new Date(minCreatedMs);
    var y = Number(Utilities.formatDate(createdLocal, tz, "yyyy"));
    var m = Number(Utilities.formatDate(createdLocal, tz, "M")) - 1;
    var d = Number(Utilities.formatDate(createdLocal, tz, "d"));
    var hh = Number(Utilities.formatDate(createdLocal, tz, "H"));
    var packingDay = hh >= 10 ? new Date(y, m, d + 1, 0, 0, 0, 0) : new Date(y, m, d, 0, 0, 0, 0);
    var win = getPackingDayWindows_(packingDay);
    packingProbe = {
      packingDay: win.packingDayStr,
      totalWindow: win.totalLabel,
      inTotal: isInPackingDayWindow_(minCreatedMs, win),
      inMain: isInPackingMainWindow_(minCreatedMs, win),
      inSupp: isInPackingSuppWindow_(minCreatedMs, win),
      reasonIfMissing: ""
    };
  }

  var statuses = {};
  for (var fi = 0; fi < found.length; fi++) {
    var st = String(found[fi].status || "Mới").trim() || "Mới";
    statuses[st] = (statuses[st] || 0) + 1;
  }
  if (packingProbe) {
    if (!packingProbe.inTotal) packingProbe.reasonIfMissing = "Ngoài khung giờ tổng hợp N2=" + packingProbe.packingDay;
    else if (statuses["Đã soạn hàng"] || statuses["Đã xác nhận nhận hàng"]) {
      packingProbe.reasonIfMissing = "Trong khung giờ nhưng bị loại vì trạng thái đã soạn/đã xác nhận (picker mặc định)";
    } else {
      packingProbe.reasonIfMissing = "";
    }
  }

  return {
    success: true,
    key: key,
    digits: digits,
    lastRow: lastRow,
    lastCol: lastCol,
    matchCount: found.length,
    minCreatedMs: minCreatedMs,
    minCreatedLabel: isNaN(minCreatedMs) ? "" : Utilities.formatDate(new Date(minCreatedMs), tz, "dd/MM/yyyy HH:mm:ss"),
    statusSummary: statuses,
    packingProbe: packingProbe,
    found: found.slice(0, 20),
    _debugRun: "debug-order-v2"
  };
}

function luuSoSoanHangVaAnh(payload) {
  try {
    var ss = getSS();
    var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
    if (!historySheet) {
      return { success: false, error: "Không tìm thấy sheet Lịch Sử Xuất Kho", msg: "Không tìm thấy sheet Lịch Sử Xuất Kho" };
    }
    ensureHistoryStatusColumn(historySheet);

    var updates = (payload && payload.updates) ? payload.updates : [];
    var images = (payload && payload.images) ? payload.images : {};
    var actor = (payload && payload.actor) ? payload.actor : "";

    var lock = LockService.getDocumentLock();
    try {
      lock.waitLock(20000);

      var minRow = null;
      var maxRow = null;
      var i;

      for (i = 0; i < updates.length; i++) {
        var uRow = Number(updates[i].row);
        if (!uRow || uRow < 2 || isNaN(uRow)) continue;
        if (minRow === null || uRow < minRow) minRow = uRow;
        if (maxRow === null || uRow > maxRow) maxRow = uRow;
      }

      for (var imgKeyScan in images) {
        if (!images.hasOwnProperty(imgKeyScan)) continue;
        if (!images[imgKeyScan] || String(images[imgKeyScan]).indexOf("base64,") === -1) continue;
        var imgRowScan = parseInt(imgKeyScan, 10);
        if (!imgRowScan || imgRowScan < 2 || isNaN(imgRowScan)) continue;
        if (minRow === null || imgRowScan < minRow) minRow = imgRowScan;
        if (maxRow === null || imgRowScan > maxRow) maxRow = imgRowScan;
      }

      var matrix = null;
      var numRows = 0;
      if (minRow !== null) {
        numRows = maxRow - minRow + 1;
        // getRange(row, column, numRows, numColumns) — đọc A..P một lần
        matrix = historySheet.getRange(minRow, 1, numRows, 16).getValues();

        for (i = 0; i < updates.length; i++) {
          var wRow = Number(updates[i].row);
          if (!wRow || wRow < 2 || isNaN(wRow) || wRow < minRow || wRow > maxRow) continue;
          var offset = wRow - minRow;
          var wVal = updates[i].val;
          if (wVal !== "" && wVal !== null && wVal !== undefined) {
            var parsedVal = Number(wVal);
            matrix[offset][8] = parsedVal;       // cột 9 — SL thực tế
            matrix[offset][12] = "Đã soạn hàng"; // cột 13
            matrix[offset][13] = actor;          // cột 14
            matrix[offset][14] = "Soạn hàng";    // cột 15
            matrix[offset][15] = parsedVal;      // cột 16 — SL Giao (Soạn)
          } else {
            matrix[offset][8] = "";
            matrix[offset][15] = "";
            matrix[offset][12] = "Mới";
          }
        }
      }

      var anhDaLuu = 0;
      var hasImages = false;
      for (var imgCheck in images) {
        if (!images.hasOwnProperty(imgCheck)) continue;
        if (images[imgCheck] && String(images[imgCheck]).indexOf("base64,") !== -1) {
          hasImages = true;
          break;
        }
      }

      if (hasImages) {
        var targetFolder;
        try {
          var folders = DriveApp.getFoldersByName("dieuchuyenhanghoa");
          if (folders.hasNext()) targetFolder = folders.next();
          else targetFolder = DriveApp.createFolder("dieuchuyenhanghoa");
        } catch (driveErr) {
          throw new Error("Chưa cấp quyền Drive");
        }

        for (var rIdx in images) {
          if (!images.hasOwnProperty(rIdx)) continue;
          if (!images[rIdx] || String(images[rIdx]).indexOf("base64,") === -1) continue;
          var imgRow = parseInt(rIdx, 10);
          if (!imgRow || imgRow < 2 || isNaN(imgRow)) continue;

          var splitBase = String(images[rIdx]).split("base64,");
          var file = targetFolder.createFile(
            Utilities.newBlob(
              Utilities.base64Decode(splitBase[1]),
              "image/jpeg",
              "XacNhan_" + new Date().getTime() + ".jpg"
            )
          );
          try {
            file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          } catch (shareErr) {}

          var fileUrl = file.getUrl();
          if (matrix && imgRow >= minRow && imgRow <= maxRow) {
            matrix[imgRow - minRow][10] = fileUrl; // cột 11
          } else {
            historySheet.getRange(imgRow, 11).setValue(fileUrl);
          }
          anhDaLuu++;
        }
      }

      if (matrix) {
        historySheet.getRange(minRow, 1, numRows, 16).setValues(matrix);
      }

      // Trừ tồn biến thể: Da_Xuat += SL soạn, Ton_Hien_Tai = Ban_Dau - Da_Xuat + Da_Nhan_Nhap
      var variantExportLines = [];
      if (matrix) {
        for (i = 0; i < updates.length; i++) {
          var expRow = Number(updates[i].row);
          if (!expRow || expRow < minRow || expRow > maxRow) continue;
          var expOff = expRow - minRow;
          var expQty = Number(matrix[expOff][8]);
          if (!expQty || expQty <= 0 || isNaN(expQty)) continue;
          // Chỉ trừ khi trạng thái đã soạn
          if (String(matrix[expOff][12] || "").trim() !== "Đã soạn hàng") continue;
          variantExportLines.push({
            maHang: matrix[expOff][4],
            maVach: matrix[expOff][5],
            tenHang: matrix[expOff][6],
            dvt: matrix[expOff][9],
            qty: expQty
          });
        }
      }
      if (variantExportLines.length) {
        try { applyTonVariantExportBatch_(ss, variantExportLines); } catch (eVarExp) { Logger.log(eVarExp); }
      }

      SpreadsheetApp.flush();

      var soPhieu = "";
      if (matrix) {
        for (i = 0; i < updates.length; i++) {
          var notifyRow = Number(updates[i].row);
          if (!notifyRow || notifyRow < minRow || notifyRow > maxRow) continue;
          var cellSp = matrix[notifyRow - minRow][1];
          if (cellSp) {
            soPhieu = String(cellSp).trim();
            break;
          }
        }
      }

      var khoXuat = "";
      var khoNhan = "";
      var totalRows = 0;
      var missingCount = 0;
      var extraCount = 0;

      if (soPhieu) {
        var matchSet = buildOrderMatchSet_(soPhieu);
        var pack = readHistoryForSelectedOrders_(historySheet, matchSet, "", 4000);
        var historyData = pack.data;
        for (var j = 1; j < historyData.length; j++) {
          var statRow = historyData[j];
          var rowSoPhieu = statRow[1] ? String(statRow[1]).trim() : "";
          if (!rowSoPhieu || !orderKeysMatch_(rowSoPhieu, soPhieu)) continue;
          totalRows += 1;
          var requestedQty = Number(statRow[7]) || 0;
          var actualQty = (statRow[8] !== "" && statRow[8] !== undefined && statRow[8] !== null)
            ? Number(statRow[8])
            : requestedQty;
          if (actualQty < requestedQty) missingCount += 1;
          if (actualQty > requestedQty) extraCount += 1;
          if (!khoXuat && statRow[2]) khoXuat = String(statRow[2]).trim();
          if (!khoNhan && statRow[3]) khoNhan = String(statRow[3]).trim();
        }
      }

      return {
        success: true,
        message: "✅ Đã lưu " + updates.length + " món và " + anhDaLuu + " ảnh!",
        notify: {
          soPhieu: soPhieu,
          khoXuat: khoXuat,
          khoNhan: khoNhan,
          updatesCount: updates.length,
          totalRows: totalRows,
          missingCount: missingCount,
          extraCount: extraCount,
          actor: (payload && payload.actor) ? payload.actor : "Chi nhánh"
        }
      };
    } finally {
      try { lock.releaseLock(); } catch (lockErr) {}
    }
  } catch (err) {
    return {
      success: false,
      error: err.message || String(err),
      msg: err.message || String(err)
    };
  }
}

function parseDateInputYYYYMMDD(value) {
  var raw = String(value || "").trim();
  if (!raw) return null;
  var m = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  var y = Number(m[1]);
  var mm = Number(m[2]) - 1;
  var d = Number(m[3]);
  var date = new Date(y, mm, d);
  if (isNaN(date.getTime())) return null;
  date.setHours(0, 0, 0, 0);
  return date;
}

/** Ghép ngày + "HH:mm" → Date (local components — tránh lệch UTC) */
function combineDateAndTime_(dateObj, timeHHmm) {
  if (!dateObj || !(dateObj instanceof Date) || isNaN(dateObj.getTime())) return null;
  var raw = String(timeHHmm || "").trim();
  var m = raw.match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  var hh = Number(m[1]);
  var mm = Number(m[2]);
  if (isNaN(hh) || isNaN(mm) || hh < 0 || hh > 23 || mm < 0 || mm > 59) return null;
  return new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate(), hh, mm, 0, 0);
}

/**
 * Chuẩn hóa mọi giá trị ngày/giờ về Unix ms theo lịch local (Asia/Ho_Chi_Minh khi script TZ đúng).
 * Tránh so sánh chuỗi và tránh `new Date("yyyy-MM-dd...")` parse UTC.
 */
function toHoChiMinhMillis_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value.getTime();
  if (value === null || value === undefined || value === "") return NaN;
  if (typeof value === "number" && isFinite(value)) return value;

  var s = String(value).trim();
  if (!s) return NaN;

  var mIso = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?(?:\.\d+)?)?/);
  if (mIso) {
    return new Date(
      Number(mIso[1]), Number(mIso[2]) - 1, Number(mIso[3]),
      Number(mIso[4] || 0), Number(mIso[5] || 0), Number(mIso[6] || 0), 0
    ).getTime();
  }

  var mVn = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (mVn) {
    return new Date(
      Number(mVn[3]), Number(mVn[2]) - 1, Number(mVn[1]),
      Number(mVn[4] || 0), Number(mVn[5] || 0), Number(mVn[6] || 0), 0
    ).getTime();
  }

  try {
    var tz = Session.getScriptTimeZone() || "Asia/Ho_Chi_Minh";
    var formats = [
      "dd/MM/yyyy HH:mm:ss",
      "dd/MM/yyyy HH:mm",
      "yyyy-MM-dd HH:mm:ss",
      "yyyy-MM-dd HH:mm",
      "yyyy/MM/dd HH:mm:ss",
      "yyyy/MM/dd HH:mm"
    ];
    for (var fi = 0; fi < formats.length; fi++) {
      try {
        var parsed = Utilities.parseDate(s, tz, formats[fi]);
        if (parsed && !isNaN(parsed.getTime())) return parsed.getTime();
      } catch (parseErr) {}
    }
  } catch (tzErr) {}

  var fallback = new Date(s);
  return isNaN(fallback.getTime()) ? NaN : fallback.getTime();
}

function toMillisSafe_(value) {
  return toHoChiMinhMillis_(value);
}

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

function getScriptTodayStart_() {
  var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  return parseDateInputYYYYMMDD(todayStr);
}

function formatSheetDateYYYYMMDD(value) {
  if (value === null || value === undefined || value === "") return "";
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  var asString = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(asString)) return asString;
  var parsed = new Date(value);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return "";
}

function matchesNgayFilter(rowDate, ngayFilter) {
  var filter = String(ngayFilter || "").trim().toLowerCase();
  if (!filter || filter === "all") return true;

  var rowDateStr = formatSheetDateYYYYMMDD(rowDate);
  if (!rowDateStr) return false;

  if (/^\d{4}-\d{2}-\d{2}$/.test(filter)) {
    return rowDateStr === filter;
  }

  var today = getScriptTodayStart_();
  if (!today) return true;
  var rowDateObj = parseDateInputYYYYMMDD(rowDateStr);
  if (!rowDateObj) return false;

  if (filter === "today") {
    return rowDateObj.getTime() === today.getTime();
  }
  if (filter === "yesterday") {
    var yesterday = new Date(today);
    yesterday.setDate(today.getDate() - 1);
    return rowDateObj.getTime() === yesterday.getTime();
  }
  if (filter === "7days") {
    var start = new Date(today);
    start.setDate(today.getDate() - 6);
    return rowDateObj.getTime() >= start.getTime() && rowDateObj.getTime() <= today.getTime();
  }

  return rowDateStr === filter;
}

function lookupStoreCodeDigits_(storeName) {
  // Không gọi isSameStoreName/formatShortStoreLabel ở đây (tránh đệ quy).
  // Mã link dữ liệu Q4: luôn dùng 178 / 275 (không dùng mã sheet kiểu 004 / 006).
  var raw = String(storeName || "").trim();
  if (!raw) return "";
  var rawNorm = normalizeHeaderText(raw);
  var activeMap = getActiveStoreMap();

  // 1) Cấu hình cố định theo tên đầy đủ / tên ngắn
  if (STORE_SHORT_CODES[raw]) return STORE_SHORT_CODES[raw];
  for (var full in STORE_SHORT_CODES) {
    if (!Object.prototype.hasOwnProperty.call(STORE_SHORT_CODES, full)) continue;
    var fullNorm = normalizeHeaderText(full);
    var shortOfFull = normalizeHeaderText(activeMap[full] || "");
    if (
      rawNorm === fullNorm ||
      (shortOfFull && rawNorm === shortOfFull) ||
      raw === activeMap[full] ||
      rawNorm.indexOf(fullNorm) !== -1
    ) {
      return STORE_SHORT_CODES[full];
    }
  }

  // 2) Suy từ tên hiển thị (Quận 4 Mới / Cũ)
  var display = normalizeHeaderText(activeMap[raw] || raw);
  if ((display.indexOf("q4") !== -1 || display.indexOf("quan4") !== -1) && display.indexOf("moi") !== -1) return "178";
  if ((display.indexOf("q4") !== -1 || display.indexOf("quan4") !== -1) && (display.indexOf("cu") !== -1 || display.indexOf("old") !== -1)) return "275";

  // 3) Sheet Hướng dẫn — chỉ nhận mã link hợp lệ 178/275 (bỏ 004/006…)
  try {
    var registry = getStoreRegistry();
    var details = registry && registry.storeDetails ? registry.storeDetails : [];
    for (var i = 0; i < details.length; i++) {
      var d = details[i];
      if (!d) continue;
      var dFullNorm = normalizeHeaderText(d.fullName || "");
      var dShortNorm = normalizeHeaderText(d.shortName || "");
      if (rawNorm !== dFullNorm && rawNorm !== dShortNorm && raw !== d.fullName && raw !== d.shortName) continue;
      var code = String(d.code || "").trim();
      var m = code.match(/(\d{2,})/);
      if (!m) continue;
      if (m[1] === "178" || m[1] === "275") return m[1];
      // Map mã kho nội bộ thường gặp → mã link
      if (STORE_SHORT_CODES[d.fullName]) return STORE_SHORT_CODES[d.fullName];
    }
  } catch (e) {}

  return "";
}

function formatShortStoreLabel(storeName) {
  var activeMap = getActiveStoreMap();
  var name = String(activeMap[storeName] || storeName || "").trim();
  var normalized = normalizeHeaderText(name);
  var base = "";
  if (normalized.indexOf("q7") !== -1 || normalized.indexOf("quan7") !== -1) base = "Q7";
  else if (normalized.indexOf("q8") !== -1 || normalized.indexOf("quan8") !== -1) base = "Q8";
  else if (normalized.indexOf("phamhung") !== -1) base = "PH";
  else if (normalized.indexOf("q5") !== -1 || normalized.indexOf("quan5") !== -1) base = "Q5";
  else if (normalized.indexOf("q1") !== -1 || normalized.indexOf("quan1") !== -1) base = "Q1";
  else if (normalized.indexOf("q4") !== -1 || normalized.indexOf("quan4") !== -1) base = "Q4";
  else return name || "Khác";

  // Hai cửa hàng Q4 phải tách cột: Q4_178 / Q4_275 (không dùng Q4_004 / Q4_006)
  if (base === "Q4") {
    var code = lookupStoreCodeDigits_(storeName);
    if (!code) code = lookupStoreCodeDigits_(name);
    if (code === "178" || code === "275") return "Q4_" + code;
    // Fallback cuối theo tên đã chuẩn hóa
    if (normalized.indexOf("moi") !== -1) return "Q4_178";
    if (normalized.indexOf("cu") !== -1 || normalized.indexOf("old") !== -1) return "Q4_275";
    return "Q4";
  }
  return base;
}

function getCanonicalStoreKey(storeName) {
  var normalizedStore = normalizeStoreName(storeName || "");
  if (!normalizedStore) return "";
  var shortLabel = formatShortStoreLabel(normalizedStore);
  var key = normalizeHeaderText(shortLabel);
  if (key) return key;
  return normalizeHeaderText(normalizedStore);
}

function getStockMapByStoreName(stockIndex, storeName) {
  if (!stockIndex) return {};
  var key = getCanonicalStoreKey(storeName);
  if (key && stockIndex[key]) return stockIndex[key];
  return {};
}

function getStockIndexByStore(stockData) {
  var index = {};
  var stockConfig = getStockSheetConfig(stockData);
  var header = stockData[stockConfig.headerIndex] || [];
  var currentMaHang = "";
  var currentMaVach = "";
  for (var i = stockConfig.startRow; i < stockData.length; i++) {
    var row = stockData[i];
    if (!row) continue;
    var rowMaHangRaw = getCellValue(row, stockConfig.maHangIdx, "");
    var rowMaVachRaw = getCellValue(row, stockConfig.maVachIdx, "");
    var hasOwnCode = !!(rowMaHangRaw || rowMaVachRaw);
    if (hasOwnCode) {
      currentMaHang = rowMaHangRaw;
      currentMaVach = rowMaVachRaw;
    }
    var maHang = normalizeProductCode((hasOwnCode ? rowMaHangRaw : currentMaHang) || "");
    var maVach = normalizeProductCode((hasOwnCode ? rowMaVachRaw : currentMaVach) || "");

    var storedAny = false;
    if (stockConfig.storeHeaderIndexes && stockConfig.storeHeaderIndexes.length) {
      for (var c = 0; c < stockConfig.storeHeaderIndexes.length; c++) {
        var storeHeaderIdx = stockConfig.storeHeaderIndexes[c];
        var store = getCellValue(header, storeHeaderIdx, "");
        if (!store) continue;
        var qty = parseQuantityValue(row[storeHeaderIdx]);
        if (qty === 0) continue;
        var storeKey = getCanonicalStoreKey(store);
        if (!storeKey) continue;
        if (!index[storeKey]) index[storeKey] = {};
        if (maHang) {
          index[storeKey][maHang] = (index[storeKey][maHang] || 0) + qty;
          var maHangCompact = normalizeNumericCode(maHang);
          if (maHangCompact && maHangCompact !== maHang) index[storeKey][maHangCompact] = (index[storeKey][maHangCompact] || 0) + qty;
        }
        if (maVach) {
          index[storeKey][maVach] = (index[storeKey][maVach] || 0) + qty;
          var maVachCompact = normalizeNumericCode(maVach);
          if (maVachCompact && maVachCompact !== maVach) index[storeKey][maVachCompact] = (index[storeKey][maVachCompact] || 0) + qty;
        }
        storedAny = true;
      }
    }
    if (storedAny) continue;

    var stores = getRowStoreNames(row, stockConfig);
    var qty = parseQuantityValue(row[stockConfig.tonKhoIdx]);
    if (qty === 0) continue;
    if (!stores.length) continue;

    for (var s = 0; s < stores.length; s++) {
      var store = stores[s];
      var storeKey = getCanonicalStoreKey(store);
      if (!storeKey) continue;
      if (!index[storeKey]) index[storeKey] = {};
      if (maHang) {
        index[storeKey][maHang] = (index[storeKey][maHang] || 0) + qty;
        var maHangCompact = normalizeNumericCode(maHang);
        if (maHangCompact && maHangCompact !== maHang) index[storeKey][maHangCompact] = (index[storeKey][maHangCompact] || 0) + qty;
      }
      if (maVach) {
        index[storeKey][maVach] = (index[storeKey][maVach] || 0) + qty;
        var maVachCompact = normalizeNumericCode(maVach);
        if (maVachCompact && maVachCompact !== maVach) index[storeKey][maVachCompact] = (index[storeKey][maVachCompact] || 0) + qty;
      }
    }
  }
  return index;
}

function rowCodesMatchNeeded_(maHang, maVach, neededCodes) {
  if (maHang && neededCodes[maHang]) return true;
  if (maVach && neededCodes[maVach]) return true;
  if (maHang) {
    var c = normalizeNumericCode(maHang);
    if (c && neededCodes[c]) return true;
  }
  if (maVach) {
    var c2 = normalizeNumericCode(maVach);
    if (c2 && neededCodes[c2]) return true;
  }
  return false;
}

function addStockToIndexFiltered_(index, storeKey, maHang, maVach, qty, neededCodes, neededStores) {
  if (!storeKey || !neededStores[storeKey] || !qty) return;
  if (!rowCodesMatchNeeded_(maHang, maVach, neededCodes)) return;
  if (!index[storeKey]) index[storeKey] = {};
  if (maHang) {
    index[storeKey][maHang] = (index[storeKey][maHang] || 0) + qty;
    var maHangCompact = normalizeNumericCode(maHang);
    if (maHangCompact && maHangCompact !== maHang) index[storeKey][maHangCompact] = (index[storeKey][maHangCompact] || 0) + qty;
  }
  if (maVach) {
    index[storeKey][maVach] = (index[storeKey][maVach] || 0) + qty;
    var maVachCompact = normalizeNumericCode(maVach);
    if (maVachCompact && maVachCompact !== maVach) index[storeKey][maVachCompact] = (index[storeKey][maVachCompact] || 0) + qty;
  }
}

function getPartialStockIndexForItems_(ss, itemMap, keys) {
  ss = ss || getSS();
  var stockSheet = ss.getSheetByName("TỔNG HỢP TỒN KHO");
  if (!stockSheet || !keys || !keys.length) return {};

  var neededCodes = {};
  var neededStores = {};
  for (var ki = 0; ki < keys.length; ki++) {
    var it = itemMap[keys[ki]];
    if (!it) continue;
    if (it.maHang) {
      var mh = normalizeProductCode(it.maHang);
      if (mh) neededCodes[mh] = true;
    }
    if (it.maVach) {
      var mv = normalizeProductCode(it.maVach);
      if (mv) neededCodes[mv] = true;
    }
    for (var src in it.sourceStores) {
      var sk = getCanonicalStoreKey(src);
      if (sk) neededStores[sk] = true;
    }
  }
  if (!Object.keys(neededCodes).length || !Object.keys(neededStores).length) return {};

  var lastRow = stockSheet.getLastRow();
  var lastCol = stockSheet.getLastColumn();
  if (lastRow < 2) return {};

  var headerRows = Math.min(20, lastRow);
  var headerData = stockSheet.getRange(1, 1, headerRows, lastCol).getValues();
  var stockConfig = getStockSheetConfig(headerData);
  var header = headerData[stockConfig.headerIndex] || [];

  var colSet = {};
  colSet[stockConfig.maHangIdx] = true;
  colSet[stockConfig.maVachIdx] = true;
  colSet[stockConfig.tonKhoIdx] = true;
  if (stockConfig.storeIndexes) {
    for (var si = 0; si < stockConfig.storeIndexes.length; si++) colSet[stockConfig.storeIndexes[si]] = true;
  }
  if (stockConfig.storeHeaderIndexes) {
    for (var hi = 0; hi < stockConfig.storeHeaderIndexes.length; hi++) {
      var hIdx = stockConfig.storeHeaderIndexes[hi];
      var storeName = getCellValue(header, hIdx, "");
      var storeKey = getCanonicalStoreKey(storeName);
      if (storeKey && neededStores[storeKey]) colSet[hIdx] = true;
    }
  }
  var colList = [];
  for (var ck in colSet) {
    if (colSet.hasOwnProperty(ck) && colSet[ck] !== false) colList.push(Number(ck));
  }
  colList.sort(function(a, b) { return a - b; });
  if (!colList.length) return {};

  var minCol = colList[0] + 1;
  var maxCol = colList[colList.length - 1] + 1;
  var colOffset = minCol - 1;
  var relMaHang = stockConfig.maHangIdx - colOffset;
  var relMaVach = stockConfig.maVachIdx - colOffset;
  var relTonKho = stockConfig.tonKhoIdx - colOffset;
  var relStoreIndexes = (stockConfig.storeIndexes || []).map(function(idx) { return idx - colOffset; });
  var relStoreHeaderIndexes = [];
  if (stockConfig.storeHeaderIndexes) {
    for (var rh = 0; rh < stockConfig.storeHeaderIndexes.length; rh++) {
      var origIdx = stockConfig.storeHeaderIndexes[rh];
      if (!colSet[origIdx]) continue;
      relStoreHeaderIndexes.push({ relIdx: origIdx - colOffset, origIdx: origIdx });
    }
  }

  var slimConfig = {
    startRow: stockConfig.startRow,
    headerIndex: stockConfig.headerIndex,
    maHangIdx: relMaHang,
    maVachIdx: relMaVach,
    tonKhoIdx: relTonKho,
    storeIndexes: relStoreIndexes,
    storeHeaderIndexes: relStoreHeaderIndexes.map(function(x) { return x.relIdx; })
  };

  var index = {};
  var currentMaHang = "";
  var currentMaVach = "";
  var dataStartRow = stockConfig.startRow + 1;
  var chunkSize = 2500;

  for (var startRow = dataStartRow; startRow <= lastRow; startRow += chunkSize) {
    var endRow = Math.min(lastRow, startRow + chunkSize - 1);
    var body = stockSheet.getRange(startRow, minCol, endRow, maxCol).getValues();
    for (var i = 0; i < body.length; i++) {
      var row = body[i];
      if (!row) continue;
      var rowMaHangRaw = getCellValue(row, slimConfig.maHangIdx, "");
      var rowMaVachRaw = getCellValue(row, slimConfig.maVachIdx, "");
      var hasOwnCode = !!(rowMaHangRaw || rowMaVachRaw);
      if (hasOwnCode) {
        currentMaHang = rowMaHangRaw;
        currentMaVach = rowMaVachRaw;
      }
      var maHang = normalizeProductCode((hasOwnCode ? rowMaHangRaw : currentMaHang) || "");
      var maVach = normalizeProductCode((hasOwnCode ? rowMaVachRaw : currentMaVach) || "");
      if (!rowCodesMatchNeeded_(maHang, maVach, neededCodes)) continue;

      var storedAny = false;
      if (slimConfig.storeHeaderIndexes && slimConfig.storeHeaderIndexes.length) {
        for (var c = 0; c < slimConfig.storeHeaderIndexes.length; c++) {
          var storeHeaderIdx = slimConfig.storeHeaderIndexes[c];
          var store = getCellValue(header, colOffset + storeHeaderIdx, "");
          if (!store) continue;
          var qty = parseQuantityValue(row[storeHeaderIdx]);
          if (qty === 0) continue;
          var storeKey = getCanonicalStoreKey(store);
          addStockToIndexFiltered_(index, storeKey, maHang, maVach, qty, neededCodes, neededStores);
          storedAny = storedAny || !!(storeKey && neededStores[storeKey] && qty);
        }
      }
      if (storedAny) continue;

      var stores = getRowStoreNames(row, slimConfig);
      var qty2 = parseQuantityValue(row[slimConfig.tonKhoIdx]);
      if (qty2 === 0 || !stores.length) continue;
      for (var s = 0; s < stores.length; s++) {
        var storeKey2 = getCanonicalStoreKey(stores[s]);
        addStockToIndexFiltered_(index, storeKey2, maHang, maVach, qty2, neededCodes, neededStores);
      }
    }
  }
  return index;
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
  // #region agent log
  var _dbgT0 = Date.now();
  // #endregion
  // ngayTo = ngày tổng hợp/giao N2. Fallback: ngay hoặc hôm nay.
  var packingDay = parseDateInputYYYYMMDD(ngayToYYYYMMDD) || parseDateInputYYYYMMDD(ngayYYYYMMDD) || getScriptTodayStart_() || new Date();
  packingDay = new Date(packingDay.getFullYear(), packingDay.getMonth(), packingDay.getDate(), 0, 0, 0, 0);
  var win = getPackingDayWindows_(packingDay);
  var mode = normalizePackingMode_(packingMode, false);
  // Luôn lấy FULL đơn trong khung ca — không lọc theo trạng thái
  var orders = getEligibleOrdersForSoanHang(packingDay, userRole, userStore, null, packingDay, win, mode, {
    includeAllStatuses: true
  });
  // #region agent log
  var _dbgMs = Date.now() - _dbgT0;
  // #endregion
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
    _debugTotalMs: _dbgMs,
    _debugRun: "packing-window-v5"
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
        inMain: false,
        inSupp: false
      };
    }

    var entry = map[soPhieu];
    // Theo dõi trạng thái để trả FE — KHÔNG dùng để loại đơn
    if (displayStatus === "Đã xác nhận") entry.hasConfirmed = true;
    else if (displayStatus === "Đã soạn") entry.hasPacked = true;
    else if (displayStatus === "Đã hủy" || displayStatus === "Đã hủy dòng") entry.hasCancelled = true;
    else entry.hasOpen = true;

    // Giữ mốc createdAt sớm nhất (gốc) — không bị lệch khi sửa/soạn sau 10h
    if (!isNaN(createdMs) && (isNaN(entry.createdAtMs) || createdMs < entry.createdAtMs)) {
      entry.createdAtMs = createdMs;
    }
  }

  var orders = [];
  var tz = Session.getScriptTimeZone() || "Asia/Ho_Chi_Minh";
  var storesWithMain = {};
  for (var key in map) {
    var item = map[key];
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
      thoiGianDat: Utilities.formatDate(new Date(item.createdAtMs), tz, "dd/MM/yyyy HH:mm"),
      thoiGianDatMillis: item.createdAtMs,
      packingBucket: bucket,
      packingBucketLabel: bucketLabel,
      groupKey: knKey,
      trangThai: statusMeta.trangThai,
      status: statusMeta.status,
      statusTone: statusMeta.statusTone,
      defaultChecked: !!statusMeta.defaultChecked,
      alreadyPacked: !!statusMeta.alreadyPacked
    });
  }

  // Mode tổng hợp: gắn cờ gộp khi kho nhận đã có đơn chính + đơn bổ sung trong cùng ca
  if (mode === "total") {
    for (var oi = 0; oi < orders.length; oi++) {
      var o = orders[oi];
      var gk = o.groupKey || "";
      o.mergeWithMain = !!(o.packingBucket === "bổ sung" && gk && storesWithMain[gk]);
      o.groupHint = o.mergeWithMain
        ? ("Gộp với đơn chính kho " + (formatShortStoreLabel(o.khoNhan) || o.khoNhan || ""))
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

function taoBangSoanHangNgayMai(payload) {
  // Thuật toán nhanh: chỉ đọc dòng của đơn đã chọn + gom mã + ghi sheet.
  // BỎ đọc "TỔNG HỢP TỒN KHO" (nguyên nhân chậm 1–2 phút).
  // #region agent log
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
  // #endregion

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
  // #region agent log
  _dbgMark("readHistory", {
    historyRows: historyPack.data ? historyPack.data.length - 1 : 0,
    scannedRows: historyPack.scannedRows || (historyPack.data ? historyPack.data.length - 1 : 0),
    matchedRows: historyPack.matchedRows || 0,
    targeted: !!selectedOrdersRaw.length
  });
  // #endregion

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
  // #region agent log
  _dbgMark("catalogLookup", {
    ready: !!(catalogLookup && (Object.keys(catalogLookup.byMaVach || {}).length || Object.keys(catalogLookup.byMaHang || {}).length)),
    byMaVach: catalogLookup ? Object.keys(catalogLookup.byMaVach || {}).length : 0,
    byMaHang: catalogLookup ? Object.keys(catalogLookup.byMaHang || {}).length : 0
  });
  // #endregion
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
  // #region agent log
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
  // #endregion
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
  // #region agent log
  _dbgMark("stockIndex", {
    via: "TON_Q7+TON_VARIANT",
    stockReady: stockReady,
    q7Keys: Object.keys(q7Map).length,
    variantKeys: Object.keys(variantMapPack).length,
    forceStock: forceStock,
    itemCount: keys.length
  });
  // #endregion

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
  // Hàng 5 = số đơn theo cột kho; hàng 6 = header tên cột (Q4_178 / Q4_275 / Q8…)
  var headerRow = 6;
  var headers = ["STT", "Mã hàng", "Mã vạch", "Tên hàng", "ĐVT", "Stock Q7", "Tổng đặt"];
  var storeOrderCountRow = ["", "", "", "", "", "", "Số đơn"];
  for (var h = 0; h < storeList.length; h++) {
    headers.push(formatShortStoreLabel(storeList[h]));
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
  // #region agent log
  var _dbgDvtFromOrder = 0;
  var _dbgDvtCatalog = 0;
  var _dbgDvtInferred = 0;
  var _dbgDvtEmpty = 0;
  var _dbgDvtSample = [];
  // #endregion
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
    // #region agent log
    if (dvtSource === "resolved") _dbgDvtFromOrder++;
    else if (dvtSource === "catalog") _dbgDvtCatalog++;
    else if (dvtSource === "ton_q7") _dbgDvtInferred++;
    else _dbgDvtEmpty++;
    if (_dbgDvtSample.length < 5) {
      _dbgDvtSample.push({ ma: it.maHang || it.maVach, orderDvt: it.dvt || "", out: dvtOut || "", src: dvtSource || "empty" });
    }
    // #endregion
    var stock = stockReady ? getStockValueForItem(q7Map, it.maHang, it.maVach, dvtOut || it.dvt) : 0;
    var stockSource = stockReady ? "TON_Q7" : "";
    var vStockPack = getVariantStockIfPresent_(variantMapPack, it.maHang, it.maVach, dvtOut || it.dvt);
    if (vStockPack !== null && vStockPack !== undefined) {
      stock = Number(vStockPack) || 0;
      stockSource = "TON_VARIANT";
    }
    var rowHasStock = !!stockSource;
    var canhBao = rowHasStock ? "OK" : "Chưa có TON_Q7";
    if (rowHasStock) {
      var thieu = it.totalQty - stock;
      canhBao = thieu > 0 ? ("THIẾU " + thieu) : (stockSource === "TON_VARIANT" ? "OK (variant)" : "OK");
      if (thieu > 0) missingLines += 1;
    }
    // MH/MV giữ đúng mã con; tên = "MH - tên chi tiết" để nhặt hàng
    var tenPack = formatVariantDisplayName_(it.maHang, it.tenHang);

    var rowOut = [0, it.maHang || "", it.maVach || "", tenPack || "", dvtOut || "", rowHasStock ? stock : "", it.totalQty];
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
  // #region agent log
  _dbgMark("writeSheetData", { outputRows: rows.length, colCount: colCount, missingLines: missingLines, stockReady: stockReady });
  // #endregion

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
  // #region agent log
  _dbgMark("backfillHistoryDvt", { candidates: historyDvtBackfill.length, written: backfilled });
  // #endregion

  SpreadsheetApp.flush();
  // #region agent log
  _dbgMark("flush", {});
  var _dbgTotalMs = Date.now() - _dbgT0;
  // #endregion

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

function warmStockIndex() {
  var t0 = Date.now();
  var ss = getSS();
  var map = readTonKhoQ7Map_(ss);
  var rebuilt = false;
  if (!map || !Object.keys(map).length) {
    var info = rebuildTonKhoQ7Sheet_(ss);
    rebuilt = true;
    map = readTonKhoQ7Map_(ss);
    return {
      success: true,
      ready: !!(map && Object.keys(map).length),
      stores: 1,
      rows: (map && Object.keys(map).length) || (info && info.rows) || 0,
      cacheSource: "TON_Q7-rebuild",
      rebuilt: rebuilt,
      _debugRun: "q7-v1",
      _debugTotalMs: Date.now() - t0
    };
  }
  return {
    success: true,
    ready: true,
    stores: 1,
    rows: Object.keys(map).length,
    cacheSource: "TON_Q7",
    rebuilt: false,
    _debugRun: "q7-v1",
    _debugTotalMs: Date.now() - t0
  };
}

function getStockCacheStatus() {
  var ss = getSS();
  var map = readTonKhoQ7Map_(ss);
  if (map && Object.keys(map).length) {
    return { success: true, ready: true, source: "TON_Q7", stores: Object.keys(map).length, _debugRun: "q7-v1" };
  }
  var sh = ss.getSheetByName(TON_Q7_SHEET_NAME);
  return { success: true, ready: false, source: sh ? "TON_Q7-empty" : "none", _debugRun: "q7-v1" };
}

// --- XUẤT BÁN HÀNG / BÁN KÈM DỊCH VỤ ---
var XUAT_BAN_SHEET_NAME = "Xuất Bán Hàng";

function ensureXuatBanHangSheet_(ss) {
  ss = ss || getSS();
  var sh = ss.getSheetByName(XUAT_BAN_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(XUAT_BAN_SHEET_NAME);
    sh.getRange("A1:L1").setValues([[
      "Thời gian tạo",
      "Số hóa đơn liên kết",
      "Mã phiếu XB",
      "Chi nhánh xuất bán",
      "Mã hàng",
      "Mã vạch",
      "Tên hàng hóa",
      "ĐVT",
      "Số lượng",
      "Người tạo",
      "Ghi chú",
      "Trạng thái"
    ]]).setFontWeight("bold").setBackground("#d9ead3");
    sh.setFrozenRows(1);
    try { sh.setColumnWidth(2, 160); sh.setColumnWidth(7, 280); } catch (e) {}
  }
  return sh;
}

/** Lưu phiếu xuất bán kèm dịch vụ — bắt buộc số hóa đơn liên kết */
function luuXuatBanHang(payload) {
  payload = payload || {};
  var soHoaDon = String(payload.soHoaDon || "").trim();
  if (!soHoaDon) throw new Error("Vui lòng nhập số hóa đơn liên kết trước khi lưu.");
  var items = payload.items || [];
  if (!items.length) throw new Error("Chưa có mặt hàng nào để lưu.");

  var chiNhanh = normalizeStoreName(payload.chiNhanh || payload.khoXuat || "");
  if (!chiNhanh) throw new Error("Thiếu chi nhánh xuất bán.");
  var actor = String(payload.actor || "").trim();
  var ss = getSS();
  var catalogLookup = getCatalogLookup(ss);
  var now = new Date();
  var maPhieu = "XB-" + Math.floor(100000 + Math.random() * 900000);

  var rows = [];
  var coLoi = false;
  for (var i = 0; i < items.length; i++) {
    var item = items[i] || {};
    // Cùng nguồn Data_Excel như luuPhieuTuWebApp / tab Tạo đơn
    var catalogItem = resolveCatalogProduct(catalogLookup, item.maHang, item.maVach);
    var maHangOut = (catalogItem && catalogItem.maHang) ? catalogItem.maHang : (item.maHang || "");
    var maVachOut = (catalogItem && catalogItem.maVach) ? catalogItem.maVach : (item.maVach || "");
    var tenHangOut = (catalogItem && catalogItem.tenHang) ? catalogItem.tenHang : (item.tenHang || "");
    var dvtIn = item.dvt || (catalogItem && catalogItem.dvt) || "";
    var dvtResolved = resolveDvtValue(catalogLookup, maHangOut, maVachOut, dvtIn);
    var slNum = Number(item.sl);
    var note = "";
    var bad = false;
    if (isNaN(slNum) || slNum <= 0) { bad = true; note = "Lỗi số lượng"; slNum = 0; }
    if (!dvtResolved || dvtResolved === "Không tìm thấy") {
      bad = true;
      note += (note ? " | " : "") + "Lỗi ĐVT";
    }
    if (String(item.maHang || "") === "LỖI MÃ" || String(maHangOut || "") === "LỖI MÃ") {
      bad = true;
      note += (note ? " | " : "") + "Mã không tồn tại";
    }
    if (bad) coLoi = true;
    rows.push([
      now,
      soHoaDon,
      maPhieu,
      chiNhanh,
      maHangOut,
      maVachOut,
      tenHangOut,
      dvtResolved || "",
      slNum,
      actor,
      note,
      "Đã lưu"
    ]);
  }

  var lock = LockService.getDocumentLock();
  try {
    lock.waitLock(15000);
    var sh = ensureXuatBanHangSheet_(ss);
    var lastRow = sh.getLastRow();
    sh.getRange(lastRow + 1, 1, rows.length, 12).setValues(rows);
    sh.getRange(lastRow + 1, 1, rows.length, 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  return {
    success: true,
    maPhieu: maPhieu,
    soHoaDon: soHoaDon,
    chiNhanh: chiNhanh,
    itemCount: items.length,
    coLoi: coLoi,
    message: "✅ Đã lưu xuất bán " + maPhieu + " · HĐ " + soHoaDon + " (" + items.length + " món)"
  };
}

/** Danh sách phiếu xuất bán gần đây (gom theo mã phiếu XB) */
function layDanhSachXuatBanHang(ngayYYYYMMDD, userRole, userStore, soHoaDonFilter) {
  var t0 = Date.now();
  var ss = getSS();
  var sh = ss.getSheetByName(XUAT_BAN_SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) {
    return { success: true, data: [], _debugTotalMs: Date.now() - t0 };
  }
  var lastRow = sh.getLastRow();
  var start = Math.max(2, lastRow - 3000);
  var num = lastRow - start + 1;
  var body = sh.getRange(start, 1, num, 12).getValues();
  var filterHd = String(soHoaDonFilter || "").trim().toLowerCase();
  var filterStore = normalizeStoreName(userStore || "");
  var map = {};

  for (var i = 0; i < body.length; i++) {
    var row = body[i];
    if (!row) continue;
    var maPhieu = row[2] ? String(row[2]).trim() : "";
    var soHd = row[1] ? String(row[1]).trim() : "";
    if (!maPhieu && !soHd) continue;
    if (filterHd && soHd.toLowerCase().indexOf(filterHd) === -1) continue;
    if (!matchesNgayFilter(row[0], ngayYYYYMMDD || "7days")) continue;
    var cn = row[3] ? String(row[3]).trim() : "";
    if (userRole !== "Admin" && filterStore && !isSameStoreName(cn, filterStore)) continue;
    var key = maPhieu || (soHd + "|" + cn);
    if (!map[key]) {
      map[key] = {
        maPhieu: maPhieu,
        soHoaDon: soHd,
        chiNhanh: cn,
        thoiGian: row[0] instanceof Date ? row[0].getTime() : "",
        actor: row[9] ? String(row[9]).trim() : "",
        itemCount: 0,
        tongSl: 0
      };
    }
    map[key].itemCount += 1;
    map[key].tongSl += Number(row[8]) || 0;
  }

  var list = [];
  for (var k in map) {
    if (map.hasOwnProperty(k)) list.push(map[k]);
  }
  list.sort(function(a, b) { return (b.thoiGian || 0) - (a.thoiGian || 0); });
  return {
    success: true,
    data: list.slice(0, 100),
    _debugTotalMs: Date.now() - t0,
    _debugRun: "xuat-ban-v1"
  };
}