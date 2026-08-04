// ============================================================
// catalog_variant.gs — Data_Excel / Parent_SKU / TON_VARIANT / Hang moi
// ============================================================

// Kho soạn hàng chính — sheet nhẹ chỉ chứa tồn Q7 (tạo lúc import file tồn)
var PACKING_STOCK_STORE = "Kho Địa điểm kinh doanh Q7";

var TON_Q7_SHEET_NAME = "TON_Q7";

var CACHE_TON_Q7_KEY = "ton_q7_map_v2";

// Tồn riêng theo biến thể đồ chơi — đối soát: Ton_Ban_Dau - Da_Xuat + Da_Nhan_Nhap = Ton_Hien_Tai
// MASTER Cha–Con + stock biến thể: KHÔNG được sheet.clear() khi import MISA / tồn
var TON_VARIANT_SHEET_NAME = "TON_VARIANT";

var CACHE_TON_VARIANT_KEY = "ton_variant_map_v2";

/** Staging thô từ MISA — được phép ghi đè mỗi lần import */
var MISA_IMPORT_SHEET_NAME = "MISA_IMPORT";

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


function isPackingQ7Store_(storeName) {
  if (!storeName) return false;
  if (isStoreNameMatch(storeName, PACKING_STOCK_STORE)) return true;
  return formatShortStoreLabel(storeName) === "Q7";
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
 * Ghi TON_VARIANT theo map dirtyByKey — KHÔNG clear sheet.
 * Chỉ cập nhật/append các Key trong dirtyByKey; mọi dòng khác giữ nguyên trên sheet.
 */
function persistTonVariantByKeyNoClear_(ss, dirtyByKey) {
  ss = ss || getSS();
  var sh = getOrCreateTonVariantSheet_(ss);
  ensureTonVariantSchema_(sh);
  var lastRow = sh.getLastRow();
  var rowIndexByKey = {};
  if (lastRow >= 2) {
    var existingKeys = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < existingKeys.length; i++) {
      var ek = String(existingKeys[i][0] || "").trim();
      if (ek) rowIndexByKey[ek] = i + 2;
    }
  }

  var now = new Date();
  var appendRows = [];
  var updated = 0;
  for (var k in dirtyByKey) {
    if (!Object.prototype.hasOwnProperty.call(dirtyByKey, k)) continue;
    var row = dirtyByKey[k];
    if (!row || !row.k) continue;
    row.tonHienTai = calcTonHienTaiVariant_(row.tonBanDau, row.daXuat, row.daNhanNhap);
    var vals = [
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
    ];
    var sheetRow = rowIndexByKey[row.k];
    if (sheetRow) {
      sh.getRange(sheetRow, 1, 1, TON_VARIANT_COL_COUNT).setValues([vals]);
      updated++;
    } else {
      appendRows.push(vals);
    }
  }
  if (appendRows.length) {
    var start = Math.max(sh.getLastRow() + 1, 2);
    sh.getRange(start, 1, appendRows.length, TON_VARIANT_COL_COUNT).setValues(appendRows);
  }
  try { SpreadsheetApp.flush(); } catch (eF) {}

  var mapRows = [];
  var lr2 = sh.getLastRow();
  if (lr2 >= 2) {
    mapRows = sh.getRange(2, 1, lr2 - 1, TON_VARIANT_COL_COUNT).getValues();
  }
  var map = buildTonVariantStockMapFromRows_(mapRows);
  try { putCacheJson_(getScriptCache_(), CACHE_TON_VARIANT_KEY, map, CACHE_TTL_SECONDS); } catch (e3) {}
  return {
    success: true,
    rows: mapRows.length,
    updated: updated,
    appended: appendRows.length,
    sheetName: TON_VARIANT_SHEET_NAME
  };
}


/**
 * Đọc toàn bộ TON_VARIANT → map byKey (giữ Da_Xuat / Da_Nhan).
 */
function readTonVariantByKeyMap_(ss) {
  ss = ss || getSS();
  var byKey = {};
  var sh = getOrCreateTonVariantSheet_(ss);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return byKey;
  ensureTonVariantSchema_(sh);
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
  return byKey;
}


/**
 * Import / merge tồn biến thể — UPSERT, KHÔNG clear, MẶC ĐỊNH GIỮ Da_Xuat.
 * entries: {k|maHang, q|qty, p|parentSku, th|tenHang, mv|maVach, d|dvt, resetExport?:bool}
 */
function writeTonVariantEntriesToSheet_(ss, entries) {
  ss = ss || getSS();
  var t0 = Date.now();
  var lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    var byKey = readTonVariantByKeyMap_(ss);
    var dirty = {};

    for (var i = 0; i < (entries || []).length; i++) {
      var ent = entries[i];
      if (!ent) continue;
      var key = String(ent.k || "").trim();
      if (key.indexOf("MH:") === 0) key = key.substring(3).replace(/\|DV:/i, "|");
      if (!key) {
        key = buildTonVariantKey_(ent.maHang || ent.mh || "", ent.d || ent.dvt || "");
      }
      if (!key) continue;
      var qtyRaw = ent.q != null ? ent.q : ent.qty;
      var hasQty = qtyRaw !== "" && qtyRaw !== null && qtyRaw !== undefined;
      var qty = hasQty ? Number(qtyRaw) : NaN;
      if (hasQty && isNaN(qty)) qty = 0;

      var prev = byKey[key] || {
        k: key, p: "", tonBanDau: 0, daXuat: 0, tonHienTai: 0, daNhanNhap: 0, th: "", mv: "", d: ""
      };
      // Chỉ cập nhật Ton_Ban_Dau khi payload có qty — GIỮ Da_Xuat trừ khi resetExport=true
      if (hasQty) prev.tonBanDau = qty;
      if (ent.resetExport === true) {
        prev.daXuat = 0;
        if (ent.keepNhap !== true) {
          prev.daNhanNhap = Number(ent.daNhanNhap != null ? ent.daNhanNhap : 0) || 0;
        }
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
      dirty[key] = prev;
    }

    var written = persistTonVariantByKeyNoClear_(ss, dirty);
    try {
      var shSync = getOrCreateTonVariantSheet_(ss);
      var lrSync = shSync.getLastRow();
      var syncRows = lrSync >= 2 ? shSync.getRange(2, 1, lrSync - 1, TON_VARIANT_COL_COUNT).getValues() : [];
      syncParentVariantTotalsToTonQ7_(ss, syncRows);
    } catch (eSync) {}
    return {
      success: true,
      rows: written.rows || 0,
      updated: written.updated || 0,
      appended: written.appended || 0,
      ms: Date.now() - t0,
      _debugRun: "ton-variant-upsert-noclear-v1"
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
 * Map MH_Con / MV → Parent_SKU từ sheet TON_VARIANT (fallback khi Data_Excel thiếu Parent).
 * @returns {Object.<string,string>}
 */
function buildTonVariantParentLookup_(ss) {
  var out = {};
  try {
    ss = ss || getSS();
    var sh = ss.getSheetByName(TON_VARIANT_SHEET_NAME);
    if (!sh || sh.getLastRow() < 2) return out;
    ensureTonVariantSchema_(sh);
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, TON_VARIANT_COL_COUNT).getValues();
    for (var i = 0; i < data.length; i++) {
      var key = String(data[i][0] || "").trim();
      var parent = String(data[i][1] || "").trim();
      if (!parent) continue;
      var mh = key ? String(key.split("|")[0] || "").trim() : "";
      var mv = String(data[i][7] || "").trim();
      if (mh) out[mh.toUpperCase()] = parent;
      if (mv) out["MV:" + normalizeProductCode(mv)] = parent;
    }
  } catch (e) {}
  return out;
}


/**
 * Map Parent/Variant để kho nhặt theo mã cha trên bao bì.
 * @returns {{parentSku:string,variantSku:string,variantName:string,parentName:string,maHangDisplay:string,tenHangDisplay:string,maVach:string}}
 */
function resolveVariantDisplayMeta_(lookup, maHang, maVach, tenHang, parentByChild) {
  var mh = String(maHang || "").trim();
  var mv = String(maVach || "").trim();
  var th = String(tenHang || "").trim();
  var cat = resolveCatalogProduct(lookup, mh, mv);
  var parentSku = "";
  var variantName = th;
  var parentName = "";
  if (cat) {
    parentSku = String(cat.parentSku || "").trim();
    if (cat.tenHang) variantName = String(cat.tenHang).trim();
  }
  if (!parentSku && parentByChild) {
    parentSku = parentByChild[mh.toUpperCase()] ||
      (mv ? parentByChild["MV:" + normalizeProductCode(mv)] : "") ||
      "";
  }
  if (parentSku && lookup && lookup.byMaHang) {
    var parentCat = lookup.byMaHang[String(parentSku).trim().toUpperCase()];
    if (parentCat && parentCat.tenHang) parentName = String(parentCat.tenHang).trim();
  }

  var isVariant = !!(parentSku && parentSku.toUpperCase() !== mh.toUpperCase());
  var maHangDisplay = isVariant ? parentSku : mh;
  var tenHangDisplay;
  if (isVariant) {
    var baseName = parentName ? (parentName + " - " + variantName) : variantName;
    tenHangDisplay = baseName + " (Mã con: " + mh + ")";
  } else {
    tenHangDisplay = formatVariantDisplayName_(mh, th || variantName);
  }

  return {
    parentSku: parentSku || "",
    variantSku: mh,
    variantName: variantName || th,
    parentName: parentName,
    maHangDisplay: maHangDisplay || mh || mv,
    tenHangDisplay: tenHangDisplay,
    maVach: mv
  };
}


/** Gắn parentSku / variantSku / variantName vào object dòng đơn */
function attachVariantMetaToItem_(item, lookup, parentByChild) {
  if (!item) return item;
  var meta = resolveVariantDisplayMeta_(lookup, item.maHang, item.maVach, item.tenHang, parentByChild);
  item.parentSku = meta.parentSku;
  item.variantSku = meta.variantSku || String(item.maHang || "").trim();
  item.variantName = meta.variantName;
  item.parentName = meta.parentName;
  item.maHangDisplay = meta.maHangDisplay;
  item.tenHangDisplay = meta.tenHangDisplay;
  return item;
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
      variants[i].tonHienTai = variants[i].stock;
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


/**
 * Đọc map IsNew hiện có trên Data_Excel (theo MH / MV) — dùng giữ lại khi import ghi đè catalog.
 */
function readCatalogIsNewFlagMap_(ss) {
  ss = ss || getSS();
  var sh = ss.getSheetByName("Data_Excel");
  var map = {};
  if (!sh || sh.getLastRow() < 2) return map;
  try {
    var width = Math.max(sh.getLastColumn(), CATALOG_COL_COUNT);
    var values = sh.getRange(1, 1, sh.getLastRow(), width).getValues();
    var header = values[0] || [];
    var parentIdx = findCatalogParentColIdx_(header);
    var mhIdx = findCatalogMaHangColIdx_(header, parentIdx);
    if (mhIdx === -1) mhIdx = 0;
    var mvIdx = findColumnIndexByAliases(header, ["mavach", "barcode", "barcodeid"]);
    if (mvIdx === -1) mvIdx = 2;
    var isNewIdx = findColumnIndexByAliases(header, ["isnew", "trangthaimoi", "hangmoi", "newflag"]);
    if (isNewIdx === -1) isNewIdx = 10;
    for (var r = 1; r < values.length; r++) {
      var flagRaw = String(values[r][isNewIdx] == null ? "" : values[r][isNewIdx]).trim().toLowerCase();
      var isNew = flagRaw === "1" || flagRaw === "true" || flagRaw === "yes" || flagRaw === "x" || flagRaw === "moi" || flagRaw === "new";
      if (!isNew) continue;
      var mh = String(values[r][mhIdx] == null ? "" : values[r][mhIdx]).trim().toUpperCase();
      var mv = String(values[r][mvIdx] == null ? "" : values[r][mvIdx]).trim().toUpperCase();
      if (mh) map["MH:" + mh] = true;
      if (mv) map["MV:" + mv] = true;
    }
  } catch (e) {}
  return map;
}


/** Ghi staging MISA_IMPORT (được phép clear/ghi đè). */
function writeMisaImportSheet_(ss, entries, reset) {
  ss = ss || getSS();
  var sh = ss.getSheetByName(MISA_IMPORT_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(MISA_IMPORT_SHEET_NAME);
  var headers = [["Mã hàng", "Mã vạch", "Tên hàng hóa", "ĐVT", "Parent_SKU", "ImportedAt"]];
  if (reset) {
    var oldLastRow = sh.getLastRow();
    var oldLastCol = Math.max(sh.getLastColumn(), 6);
    if (oldLastRow > 0) sh.getRange(1, 1, oldLastRow, oldLastCol).clearContent();
    sh.getRange(1, 1, 1, 6).setValues(headers).setFontWeight("bold").setBackground("#fff2cc");
  } else if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, 6).setValues(headers).setFontWeight("bold").setBackground("#fff2cc");
  }
  var stamp = catalogNowStamp_();
  var rows = [];
  for (var i = 0; i < (entries || []).length; i++) {
    var e = entries[i];
    if (!e) continue;
    var mh = String(e.mh || e.maHang || "").trim();
    var mv = String(e.mv || e.maVach || "").trim();
    var th = String(e.th || e.tenHang || "").trim();
    var d = String(e.d || e.dvt || "").trim();
    var p = String(e.p || e.parentSku || "").trim();
    if (!mh && !mv) continue;
    rows.push([mh, mv, th, d, p, stamp.date]);
  }
  if (rows.length) {
    var start = Math.max(sh.getLastRow() + 1, 2);
    sh.getRange(start, 1, rows.length, 6).setValues(rows);
  }
  try { SpreadsheetApp.flush(); } catch (e) {}
  return { rows: rows.length, sheetName: MISA_IMPORT_SHEET_NAME, totalRows: Math.max(sh.getLastRow() - 1, 0) };
}


/**
 * UPSERT Data_Excel từ MISA: cập nhật tên/ĐVT/MV; GIỮ Parent_SKU + IsNew + mã con không có trong file.
 * Không xóa dòng hiện có — chỉ setValues từng dòng / append.
 */
function mergeCatalogEntriesUpsert_(ss, entries) {
  ss = ss || getSS();
  var sh = getOrCreateCatalogSheet(ss);
  var t0 = Date.now();
  var width = Math.max(sh.getLastColumn(), CATALOG_COL_COUNT);
  var lastRow = sh.getLastRow();
  if (lastRow < 1) {
    sh.getRange(1, 1, 1, CATALOG_COL_COUNT).setValues([["Mã hàng", "", "Mã vạch", "", "", "Tên hàng hóa", "", "ĐVT", "Ngày tạo", CATALOG_PARENT_HEADER, CATALOG_ISNEW_HEADER]]);
    sh.getRange(1, 1, 1, CATALOG_COL_COUNT).setFontWeight("bold").setBackground("#d9ead3");
    lastRow = 1;
    width = CATALOG_COL_COUNT;
  }
  var values = sh.getRange(1, 1, lastRow, width).getValues();
  var header = values[0] || [];
  var parentIdx = findCatalogParentColIdx_(header);
  if (parentIdx === -1) parentIdx = 9;
  var mhIdx = findCatalogMaHangColIdx_(header, parentIdx);
  if (mhIdx === -1) mhIdx = 0;
  var mvIdx = findColumnIndexByAliases(header, ["mavach", "barcode", "barcodeid"]);
  if (mvIdx === -1) mvIdx = 2;
  var thIdx = findColumnIndexByAliases(header, ["tenhang", "name", "description"]);
  if (thIdx === -1) thIdx = 5;
  var dvtIdx = findColumnIndexByAliases(header, ["dvt", "donvitinh", "donvi", "unit"]);
  if (dvtIdx === -1) dvtIdx = 7;
  var ngayIdx = findColumnIndexByAliases(header, ["ngaytao", "createdat", "created"]);
  if (ngayIdx === -1) ngayIdx = 8;
  var isNewIdx = findColumnIndexByAliases(header, ["isnew", "trangthaimoi", "hangmoi", "newflag"]);
  if (isNewIdx === -1) isNewIdx = 10;

  width = Math.max(width, CATALOG_COL_COUNT, parentIdx + 1, isNewIdx + 1);
  try {
    if (!String(sh.getRange(1, parentIdx + 1).getValue() || "").trim()) {
      sh.getRange(1, parentIdx + 1).setValue(CATALOG_PARENT_HEADER).setFontWeight("bold").setBackground("#d9ead3");
    }
    if (!String(sh.getRange(1, isNewIdx + 1).getValue() || "").trim()) {
      sh.getRange(1, isNewIdx + 1).setValue(CATALOG_ISNEW_HEADER).setFontWeight("bold").setBackground("#d9ead3");
    }
    if (!String(sh.getRange(1, ngayIdx + 1).getValue() || "").trim()) {
      sh.getRange(1, ngayIdx + 1).setValue("Ngày tạo").setFontWeight("bold").setBackground("#d9ead3");
    }
  } catch (eH) {}

  var byMh = {};
  var byMv = {};
  for (var r = 1; r < values.length; r++) {
    var mh0 = String(values[r][mhIdx] == null ? "" : values[r][mhIdx]).trim().toUpperCase();
    var mv0 = String(values[r][mvIdx] == null ? "" : values[r][mvIdx]).trim().toUpperCase();
    if (mh0 && byMh[mh0] === undefined) byMh[mh0] = r + 1; // sheet row
    if (mv0 && byMv[mv0] === undefined) byMv[mv0] = r + 1;
  }

  var stamp = catalogNowStamp_();
  var updated = 0;
  var appended = 0;
  var preservedParent = 0;
  var withDvt = 0;
  var withParent = 0;
  var appendRows = [];

  for (var i = 0; i < (entries || []).length; i++) {
    var e = entries[i];
    if (!e) continue;
    var mh = String(e.mh || e.maHang || "").trim();
    var mv = String(e.mv || e.maVach || "").trim();
    var th = String(e.th || e.tenHang || "").trim();
    var d = String(e.d || e.dvt || "").trim();
    var pIn = String(e.p || e.parentSku || "").trim();
    if (!mh && !mv) continue;
    if (d) withDvt++;

    var mhU = mh.toUpperCase();
    var mvU = mv.toUpperCase();
    var sheetRow = (mhU && byMh[mhU]) ? byMh[mhU] : ((mvU && byMv[mvU]) ? byMv[mvU] : 0);

    if (sheetRow > 1) {
      var memIdx = sheetRow - 1;
      var row = values[memIdx] ? values[memIdx].slice() : [];
      while (row.length < width) row.push("");
      var oldParent = String(row[parentIdx] == null ? "" : row[parentIdx]).trim();
      if (mh) row[mhIdx] = mh;
      if (mv) row[mvIdx] = mv;
      if (th) row[thIdx] = th;
      if (d) row[dvtIdx] = d;
      row[ngayIdx] = stamp.date;
      if (pIn) {
        row[parentIdx] = pIn;
        withParent++;
      } else if (oldParent) {
        row[parentIdx] = oldParent;
        preservedParent++;
      }
      // Giữ IsNew nguyên
      sh.getRange(sheetRow, 1, 1, width).setValues([row]);
      if (mhU) byMh[mhU] = sheetRow;
      if (mvU) byMv[mvU] = sheetRow;
      updated++;
    } else {
      var newRow = [];
      for (var c = 0; c < width; c++) newRow.push("");
      newRow[mhIdx] = mh;
      newRow[mvIdx] = mv;
      newRow[thIdx] = th;
      newRow[dvtIdx] = d;
      newRow[ngayIdx] = stamp.date;
      if (pIn) {
        newRow[parentIdx] = pIn;
        withParent++;
      }
      appendRows.push(newRow);
      appended++;
    }
  }

  if (appendRows.length) {
    var start = Math.max(sh.getLastRow() + 1, 2);
    sh.getRange(start, 1, appendRows.length, width).setValues(appendRows);
    try {
      sh.getRange(start, ngayIdx + 1, appendRows.length, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
    } catch (eFmt) {}
  }
  try { SpreadsheetApp.flush(); } catch (e) {}

  return {
    rows: (entries || []).length,
    updated: updated,
    appended: appended,
    preservedParent: preservedParent,
    withDvt: withDvt,
    withParent: withParent,
    totalRows: Math.max(sh.getLastRow() - 1, 0),
    ms: Date.now() - t0,
    ngayTaoStamp: stamp.text,
    mode: "upsert-merge-noclear"
  };
}


/** Ghi catalog từ entries — luôn UPSERT/MERGE (tham số reset chỉ còn ý nghĩa staging MISA). */
function writeCatalogEntriesToSheet_(ss, entries, reset) {
  // reset=true: ghi đè staging MISA_IMPORT; Data_Excel luôn merge
  try { writeMisaImportSheet_(ss, entries, !!reset); } catch (eM) { Logger.log(eM); }
  return mergeCatalogEntriesUpsert_(ss, entries);
}


/**
 * Gộp mã con từ TON_VARIANT (master Cha–Con) vào danhMuc nếu Data_Excel thiếu.
 */
function mergeTonVariantChildrenIntoCatalog_(ss, danhMuc) {
  danhMuc = danhMuc || {};
  try {
    ss = ss || getSS();
    var byKey = readTonVariantByKeyMap_(ss);
    var added = 0;
    for (var k in byKey) {
      if (!Object.prototype.hasOwnProperty.call(byKey, k)) continue;
      var row = byKey[k];
      if (!row) continue;
      var mh = String(row.k || "").split("|")[0].trim();
      var mv = String(row.mv || "").trim();
      var parentSku = String(row.p || "").trim();
      if (!mh && !mv) continue;
      var mhU = mh.toUpperCase();
      var mvU = mv.toUpperCase();
      var existing = (mvU && danhMuc[mvU]) || (mhU && danhMuc[mhU]) || null;
      if (existing) {
        if (!existing.parentSku && parentSku) existing.parentSku = parentSku;
        if (!existing.tenHang && row.th) existing.tenHang = row.th;
        if (!existing.dvt && row.d) existing.dvt = row.d;
        continue;
      }
      var obj = {
        maHang: mh,
        maVach: mv,
        tenHang: row.th || mh,
        dvt: row.d || "",
        parentSku: parentSku,
        isNew: false,
        fromTonVariant: true
      };
      if (mvU) danhMuc[mvU] = obj;
      if (mhU && !danhMuc[mhU]) danhMuc[mhU] = obj;
      added++;
    }
    return { danhMuc: danhMuc, added: added };
  } catch (e) {
    return { danhMuc: danhMuc, added: 0, error: e.message || String(e) };
  }
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
    // Tách mã có Parent_SKU → TON_VARIANT (UPSERT, giữ Da_Xuat — không clear sheet)
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
      _debugRun: "import-q7-variant-upsert-v3",
      msg: "Đã cập nhật " + TON_Q7_SHEET_NAME + " (" + (q7Fast.rows || 0) + " dòng)" +
        ((variantImport.rows || 0) ? (" + UPSERT " + TON_VARIANT_SHEET_NAME + " (" + variantImport.rows + " biến thể, giữ Da_Xuat)") : "") +
        "."
    };
  }

  // Import riêng TON_VARIANT (Admin) — UPSERT, không clear
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
      _debugRun: "import-ton-variant-upsert-v3",
      msg: "Đã UPSERT " + TON_VARIANT_SHEET_NAME + " (" + (vInfo.updated || 0) + " cập nhật, " +
        (vInfo.appended || 0) + " thêm mới) — giữ Da_Xuat / mã con hiện có."
    };
  }

  // Import nhanh catalog: staging MISA_IMPORT + UPSERT Data_Excel (giữ Parent_SKU / mã con)
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
    var variantMeta = { rows: 0 };
    if (isLastC) {
      // Đồng bộ meta Cha–Con sang TON_VARIANT (append mã mới, không đụng Da_Xuat)
      try {
        var tonMetaEntries = [];
        for (var ci = 0; ci < catEntries.length; ci++) {
          var ce = catEntries[ci];
          if (!ce) continue;
          var pSku = String(ce.p || ce.parentSku || "").trim();
          if (!pSku) continue;
          var childMh = String(ce.mh || ce.maHang || "").trim();
          if (!childMh) continue;
          tonMetaEntries.push({
            maHang: childMh,
            mv: ce.mv || ce.maVach || "",
            th: ce.th || ce.tenHang || "",
            d: ce.d || ce.dvt || "",
            p: pSku
            // không gửi q → không đụng Ton_Ban_Dau
          });
        }
        // Quét lại Data_Excel các dòng có Parent để UPSERT meta TON_VARIANT
        try {
          var lookupEnd = getCatalogLookup(ss);
          if (lookupEnd && lookupEnd.byMaHang) {
            for (var mk in lookupEnd.byMaHang) {
              if (!Object.prototype.hasOwnProperty.call(lookupEnd.byMaHang, mk)) continue;
              var itc = lookupEnd.byMaHang[mk];
              if (!itc || !itc.parentSku) continue;
              tonMetaEntries.push({
                maHang: itc.maHang || mk,
                mv: itc.maVach || "",
                th: itc.tenHang || "",
                d: itc.dvt || "",
                p: itc.parentSku
              });
            }
          }
        } catch (eLook) {}
        if (tonMetaEntries.length) {
          variantMeta = writeTonVariantEntriesToSheet_(ss, tonMetaEntries) || { rows: 0 };
        }
      } catch (eTonMeta) { Logger.log(eTonMeta); }
      invalidateCatalogCache_();
    }
    return {
      success: true,
      importType: importType,
      targetSheet: 'Data_Excel',
      stagingSheet: MISA_IMPORT_SHEET_NAME,
      updatedRows: catInfo.rows || 0,
      updatedCols: 8,
      chunkIndex: chunkIndexC,
      chunkTotal: chunkTotalC,
      done: isLastC,
      withDvt: catInfo.withDvt || 0,
      preservedParent: catInfo.preservedParent || 0,
      appended: catInfo.appended || 0,
      updated: catInfo.updated || 0,
      variantRows: variantMeta.rows || 0,
      _debugTotalMs: Date.now() - tCat,
      _debugRun: "import-catalog-upsert-v2",
      msg: isLastC
        ? ("UPSERT danh mục: Data_Excel " + (catInfo.totalRows || 0) + " dòng (cập nhật " +
          (catInfo.updated || 0) + ", thêm " + (catInfo.appended || 0) +
          ", giữ Parent " + (catInfo.preservedParent || 0) + "). Staging: " + MISA_IMPORT_SHEET_NAME +
          ((variantMeta.rows || 0) ? ("; TON_VARIANT meta " + variantMeta.rows + " dòng.") : "."))
        : ("Đã nhận chunk catalog " + (chunkIndexC + 1) + "/" + chunkTotalC + " (UPSERT, không xóa mã con).")
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
      // Legacy matrix → chuyển sang UPSERT entries (không wipe Data_Excel)
      var adjustedCatalogData = removeColumnFromMatrix(fileData, 3);
      var legacyEntries = [];
      for (var li = 0; li < adjustedCatalogData.length; li++) {
        if (li === 0) continue; // skip header-ish
        var lr = adjustedCatalogData[li] || [];
        var lmh = String(lr[0] || "").trim();
        var lmv = String(lr[2] || lr[1] || "").trim();
        var lth = String(lr[5] || lr[4] || lr[3] || "").trim();
        var ld = String(lr[7] || lr[6] || "").trim();
        if (!lmh && !lmv) continue;
        legacyEntries.push({ mh: lmh, mv: lmv, th: lth, d: ld });
      }
      if (!legacyEntries.length) {
        // Fallback: parse bằng extract kiểu thô từ matrix
        for (var lj = 1; lj < adjustedCatalogData.length; lj++) {
          var rowL = adjustedCatalogData[lj] || [];
          var joined = rowL.join(" ");
          if (!String(joined || "").trim()) continue;
          legacyEntries.push({
            mh: String(rowL[0] || "").trim(),
            mv: String(rowL[2] || "").trim(),
            th: String(rowL[5] || "").trim(),
            d: String(rowL[7] || "").trim()
          });
        }
      }
      var legInfo = writeCatalogEntriesToSheet_(ss, legacyEntries, isFirstChunk);
      SpreadsheetApp.flush();
      if (isLastChunk) invalidateCatalogCache_();
      return {
        success: true,
        importType: importType,
        targetSheet: 'Data_Excel',
        stagingSheet: MISA_IMPORT_SHEET_NAME,
        updatedRows: legInfo.rows || 0,
        updatedCols: legInfo.withDvt || 0,
        chunkIndex: chunkIndex,
        chunkTotal: chunkTotal,
        done: isLastChunk,
        msg: isLastChunk
          ? ('UPSERT Data_Excel từ file MISA (giữ Parent_SKU / mã con). Staging: ' + MISA_IMPORT_SHEET_NAME + '.')
          : ('Đã nhận chunk ' + (chunkIndex + 1) + '/' + chunkTotal + ' (UPSERT).')
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

  // UPSERT Data_Excel — không clearContent / không mất Parent_SKU & mã con
  var sourceEntries = catalogRows.map(function(r) {
    return { mh: r.maHang || "", mv: r.maVach || "", th: r.tenHang || "", d: r.dvt || "" };
  });
  try { writeMisaImportSheet_(ss, sourceEntries, true); } catch (eSrcMisa) {}
  var sourceCatInfo = mergeCatalogEntriesUpsert_(ss, sourceEntries);
  var catalogSheet = getOrCreateCatalogSheet(ss);

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
    catalogUpdated: (sourceCatInfo && sourceCatInfo.totalRows) || catalogRows.length,
    catalogUpserted: sourceCatInfo || {},
    stockUpdated: stockUpdated,
    sourceSheet: sourceSheetName,
    stagingSheet: MISA_IMPORT_SHEET_NAME,
    warnings: warnings,
    msg: "UPSERT danh mục từ sheet nguồn (giữ Parent_SKU / mã con). Staging: " + MISA_IMPORT_SHEET_NAME
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


function getCatalogData(forceRefresh) {
  try {
    var version = getCatalogVersion_();
    var cacheKey = CACHE_CATALOG_PREFIX + version;
    var cache = getScriptCache_();
    var force = forceRefresh === true || forceRefresh === 1 || forceRefresh === "1" ||
      forceRefresh === "true" || forceRefresh === "yes";
    // Cache chunked — hỗ trợ catalog > 90KB (tránh miss cache khiến đọc sheet mỗi lần)
    if (!force) {
      var cached = getCacheJson_(cache, cacheKey);
      if (cached && cached.success && cached.danhMuc) return cached;
    }

    var danhMuc = buildCatalogFromSheet_(getSS());
    try {
      var merged = mergeTonVariantChildrenIntoCatalog_(getSS(), danhMuc);
      danhMuc = merged.danhMuc || danhMuc;
    } catch (eMerge) {}
    var keys = 0;
    for (var k in danhMuc) {
      if (Object.prototype.hasOwnProperty.call(danhMuc, k)) keys++;
    }
    var result = {
      success: true,
      danhMuc: danhMuc,
      version: version,
      keyCount: keys,
      forced: !!force,
      mergedTonVariant: true,
      _debugRun: force ? "catalog-nocache-merge-v1" : "catalog-cache-merge-v1"
    };
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
 * Hàng mới: CHỈ Admin tick cột IsNew trên Data_Excel.
 * Không tự gắn theo Ngày tạo / cuối bảng — tránh lệch khi import tồn/danh mục.
 */
function getAutoNewProductsList_(ss, limit) {
  limit = Math.max(1, Math.min(Number(limit) || NEW_PRODUCTS_DEFAULT_LIMIT, 50));
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
  var ngayTaoIdx = findColumnIndexByAliases(headerRow, ["ngaytao", "createdat", "ngaythem", "importedat"]);
  if (ngayTaoIdx === -1 && headerRow && headerRow.length >= 9) ngayTaoIdx = 8;
  var isNewIdx = findColumnIndexByAliases(headerRow, ["isnew", "trangthaimoi", "hangmoi", "newflag"]);
  if (isNewIdx === -1 && headerRow && headerRow.length >= 11) isNewIdx = 10;
  if (parentSkuIdxNp === -1 && headerRow && headerRow.length >= 10) parentSkuIdxNp = 9;
  var startRow = headerRowIndex >= 0 ? headerRowIndex + 1 : 1;
  var tz = Session.getScriptTimeZone() || "Asia/Ho_Chi_Minh";

  var byKey = {};
  for (var i = startRow; i < rawData.length; i++) {
    if (!rawData[i]) continue;
    var maHang = getCellValue(rawData[i], maHangIdx !== -1 ? maHangIdx : 0, "");
    var maVach = getCellValue(rawData[i], maVachIdx !== -1 ? maVachIdx : 2, "");
    var tenHang = getCellValue(rawData[i], tenHangIdx !== -1 ? tenHangIdx : 5, "");
    var dvt = getCellValue(rawData[i], dvtIdx !== -1 ? dvtIdx : 7, "");
    var parentSkuNp = parentSkuIdxNp !== -1 ? getCellValue(rawData[i], parentSkuIdxNp, "") : "";
    if (!maHang && !maVach) continue;

    var flaggedNew = false;
    if (isNewIdx !== -1) {
      var flagVal = String(rawData[i][isNewIdx] == null ? "" : rawData[i][isNewIdx]).trim().toLowerCase();
      flaggedNew = flagVal === "1" || flagVal === "true" || flagVal === "yes" || flagVal === "x" || flagVal === "moi" || flagVal === "new";
    }
    if (!flaggedNew) continue;

    var key = String(maHang || maVach).trim().toUpperCase();
    if (!key) continue;
    var ngayInfo = ngayTaoIdx !== -1
      ? parseCatalogNgayTaoCell_(rawData[i][ngayTaoIdx], tz)
      : { ms: 0, label: "" };
    var sheetRow = i + 1;
    byKey[key] = {
      maHang: maHang,
      maVach: maVach,
      tenHang: tenHang,
      dvt: dvt || "",
      parentSku: parentSkuNp || "",
      sheetRow: sheetRow,
      ngayMs: ngayInfo.ms || 0,
      ngayTao: ngayInfo.label || "",
      isNew: true,
      isAdminPick: true,
      sourceReason: "admin",
      reasonLabel: "ADMIN CHỌN"
    };
  }

  var top = [];
  for (var k in byKey) {
    if (!Object.prototype.hasOwnProperty.call(byKey, k)) continue;
    top.push(byKey[k]);
  }
  top.sort(function(a, b) {
    if ((b.ngayMs || 0) !== (a.ngayMs || 0)) return (b.ngayMs || 0) - (a.ngayMs || 0);
    return (b.sheetRow || 0) - (a.sheetRow || 0);
  });
  if (top.length > limit) top = top.slice(0, limit);
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
      strategy: "admin_isNew_only"
    };
    try { putCacheJson_(cache, cacheKey, result, CACHE_TTL_SECONDS); } catch (e) {}
    return result;
  } catch (e) {
    return { success: false, error: e.message || String(e), data: [] };
  }
}


/** Admin: danh sách catalog để tick Hàng Mới — quét TOÀN BỘ Data_Excel */
function getCatalogIsNewAdminList_(query, limit) {
  try {
    // limit=0 hoặc rất lớn → lấy hết; trần an toàn 30000 tránh payload khổng lồ
    var limRaw = Number(limit);
    if (isNaN(limRaw) || limRaw < 0) limRaw = 20000;
    if (limRaw === 0) limRaw = 20000;
    limit = Math.min(Math.max(limRaw, 1), 30000);

    var qRaw = String(query || "").trim();
    var qUpper = "";
    var qFold = "";
    try {
      qUpper = qRaw.normalize("NFC").toUpperCase();
    } catch (eN) {
      qUpper = qRaw.toUpperCase();
    }
    qFold = String(qUpper)
      .toLowerCase()
      .normalize("NFKD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/đ/g, "d")
      .replace(/\s+/g, " ")
      .trim();

    var ss = getSS();
    var sh = getOrCreateCatalogSheet(ss);
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { success: true, items: [], total: 0, scanned: 0 };

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
    var scanned = 0;
    var matchedBeforeCap = 0;
    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      if (!row) continue;
      var mh = String(row[mhIdx] == null ? "" : row[mhIdx]).trim();
      var mv = String(row[mvIdx] == null ? "" : row[mvIdx]).trim();
      var th = String(row[thIdx] == null ? "" : row[thIdx]).trim();
      var dvt = String(row[dvtIdx] == null ? "" : row[dvtIdx]).trim();
      var parentSku = parentIdx !== -1 ? String(row[parentIdx] == null ? "" : row[parentIdx]).trim() : "";
      if (!mh && !mv) continue;
      scanned++;

      var mhU = "";
      var mvU = "";
      try {
        mhU = mh.normalize("NFC").toUpperCase();
        mvU = mv.normalize("NFC").toUpperCase();
      } catch (eU) {
        mhU = mh.toUpperCase();
        mvU = mv.toUpperCase();
      }
      var key = (mhU || mvU);
      if (!key || seen[key]) continue;
      seen[key] = true;

      if (qUpper || qFold) {
        var thU = "";
        var pU = "";
        try {
          thU = th.normalize("NFC").toUpperCase();
          pU = parentSku.normalize("NFC").toUpperCase();
        } catch (eT) {
          thU = th.toUpperCase();
          pU = parentSku.toUpperCase();
        }
        var hayUpper = (mhU + " " + mvU + " " + thU + " " + pU);
        var hayFold = hayUpper
          .toLowerCase()
          .normalize("NFKD")
          .replace(/[\u0300-\u036f]/g, "")
          .replace(/đ/g, "d")
          .replace(/\s+/g, " ");
        var ok = false;
        if (qUpper && hayUpper.indexOf(qUpper) !== -1) ok = true;
        if (!ok && qFold && hayFold.indexOf(qFold) !== -1) ok = true;
        if (!ok) continue;
      }

      matchedBeforeCap++;
      if (items.length >= limit) continue; // vẫn đếm matched, không early-break scan

      var flagRaw = String(row[isNewIdx] == null ? "" : row[isNewIdx]).trim().toLowerCase();
      var isNew = flagRaw === "1" || flagRaw === "true" || flagRaw === "yes" || flagRaw === "x" || flagRaw === "moi" || flagRaw === "new";
      items.push({
        sheetRow: r + 1,
        maHang: mh,
        maSP: mh,
        maVach: mv,
        tenHang: th,
        dvt: dvt,
        parentSku: parentSku,
        isNew: isNew
      });
    }
    return {
      success: true,
      items: items,
      total: items.length,
      matched: matchedBeforeCap,
      scanned: scanned,
      unique: Object.keys(seen).length,
      truncated: matchedBeforeCap > items.length,
      limit: limit,
      isNewCol: isNewIdx + 1,
      query: qRaw
    };
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


/** Admin: danh sách Parent_SKU (nhóm biến thể) */
function getParentVariantGroupsAdmin_(query, limit) {
  try {
    limit = Math.max(20, Math.min(Number(limit) || 200, 400));
    var q = String(query || "").trim().toUpperCase();
    var ss = getSS();
    var danhMuc = buildCatalogFromSheet_(ss);
    var groups = {};
    for (var key in danhMuc) {
      if (!Object.prototype.hasOwnProperty.call(danhMuc, key)) continue;
      var it = danhMuc[key];
      if (!it) continue;
      var p = String(it.parentSku || "").trim().toUpperCase();
      if (!p) continue;
      if (!groups[p]) {
        groups[p] = {
          parentSku: p,
          childCount: 0,
          sampleTen: "",
          childrenSeen: {}
        };
      }
      var ck = String(it.maHang || "").trim().toUpperCase() + "|" + String(it.maVach || "").trim().toUpperCase();
      if (ck === "|" || groups[p].childrenSeen[ck]) continue;
      groups[p].childrenSeen[ck] = true;
      groups[p].childCount++;
      if (!groups[p].sampleTen && it.tenHang) groups[p].sampleTen = String(it.tenHang);
    }
    var items = [];
    for (var pk in groups) {
      if (!Object.prototype.hasOwnProperty.call(groups, pk)) continue;
      var g = groups[pk];
      if (q && String(g.parentSku + " " + g.sampleTen).toUpperCase().indexOf(q) === -1) continue;
      items.push({
        parentSku: g.parentSku,
        childCount: g.childCount,
        sampleTen: g.sampleTen || ""
      });
      if (items.length >= limit) break;
    }
    items.sort(function(a, b) {
      return String(a.parentSku).localeCompare(String(b.parentSku));
    });
    return { success: true, items: items, total: items.length };
  } catch (e) {
    return { success: false, error: e.message || String(e), items: [] };
  }
}


/** Admin / picker: mã con + Ton_Ban_Dau + Ton_Hien_Tai theo Parent */
function getChildVariantsForAdmin_(parentSku) {
  try {
    var parent = String(parentSku || "").trim().toUpperCase();
    if (!parent) return { success: false, error: "Thiếu parentSku", variants: [] };
    var stockRes = getVariantStockList(parent);
    var variants = (stockRes && stockRes.variants) ? stockRes.variants.slice() : [];
    var ss = getSS();
    var sh = ss.getSheetByName(TON_VARIANT_SHEET_NAME);
    var byMh = {};
    if (sh && sh.getLastRow() >= 2) {
      ensureTonVariantSchema_(sh);
      var data = sh.getRange(2, 1, sh.getLastRow() - 1, TON_VARIANT_COL_COUNT).getValues();
      for (var r = 0; r < data.length; r++) {
        var p = String(data[r][1] || "").trim().toUpperCase();
        if (p !== parent) continue;
        var key = String(data[r][0] || "").trim();
        var mh = key ? String(key.split("|")[0] || "").trim().toUpperCase() : "";
        if (!mh) continue;
        byMh[mh] = {
          tonBanDau: Number(data[r][2]) || 0,
          daXuat: Number(data[r][3]) || 0,
          tonHienTai: Number(data[r][4]) || 0,
          daNhanNhap: Number(data[r][5]) || 0,
          tenHang: String(data[r][6] || "").trim(),
          maVach: String(data[r][7] || "").trim(),
          dvt: String(data[r][8] || "").trim()
        };
      }
    }
    for (var i = 0; i < variants.length; i++) {
      var mhU = String(variants[i].maHang || "").trim().toUpperCase();
      var meta = byMh[mhU];
      variants[i].tonBanDau = meta ? meta.tonBanDau : 0;
      variants[i].tonHienTai = meta ? meta.tonHienTai : (Number(variants[i].stock) || 0);
      variants[i].stock = variants[i].tonHienTai;
      variants[i].daXuat = meta ? meta.daXuat : 0;
      if (meta && meta.tenHang && !variants[i].tenHang) variants[i].tenHang = meta.tenHang;
      if (meta && meta.maVach && !variants[i].maVach) variants[i].maVach = meta.maVach;
      if (meta && meta.dvt && !variants[i].dvt) variants[i].dvt = meta.dvt;
    }
    return {
      success: true,
      parentSku: parent,
      variants: variants,
      count: variants.length
    };
  } catch (e) {
    return { success: false, error: e.message || String(e), parentSku: String(parentSku || ""), variants: [] };
  }
}


/**
 * Upsert TON_VARIANT giữ Da_Xuat / Da_Nhan_Nhap — chỉ cập nhật meta + Ton_Ban_Dau.
 * entries: [{maHang, dvt, parentSku, tenHang, maVach, tonBanDau}]
 * KHÔNG clear sheet.
 */
function upsertTonVariantKeepExport_(ss, entries) {
  ss = ss || getSS();
  var lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    var byKey = readTonVariantByKeyMap_(ss);
    var dirty = {};
    for (var i = 0; i < (entries || []).length; i++) {
      var ent = entries[i];
      if (!ent) continue;
      var key = buildTonVariantKey_(ent.maHang || ent.mh || "", ent.dvt || ent.d || "");
      if (!key) continue;
      var prev = byKey[key] || {
        k: key, p: "", tonBanDau: 0, daXuat: 0, tonHienTai: 0, daNhanNhap: 0, th: "", mv: "", d: ""
      };
      if (ent.tonBanDau !== undefined && ent.tonBanDau !== null && ent.tonBanDau !== "") {
        prev.tonBanDau = Number(ent.tonBanDau) || 0;
      }
      var p = String(ent.parentSku || ent.p || "").trim();
      var th = String(ent.tenHang || ent.th || "").trim();
      var mv = String(ent.maVach || ent.mv || "").trim();
      var d = String(ent.dvt || ent.d || "").trim();
      if (p) prev.p = p;
      if (th) prev.th = th;
      if (mv) prev.mv = mv;
      if (d) prev.d = d;
      prev.tonHienTai = calcTonHienTaiVariant_(prev.tonBanDau, prev.daXuat, prev.daNhanNhap);
      byKey[key] = prev;
      dirty[key] = prev;
    }
    var written = persistTonVariantByKeyNoClear_(ss, dirty);
    try {
      var shSync = getOrCreateTonVariantSheet_(ss);
      var lrSync = shSync.getLastRow();
      var syncRows = lrSync >= 2 ? shSync.getRange(2, 1, lrSync - 1, TON_VARIANT_COL_COUNT).getValues() : [];
      syncParentVariantTotalsToTonQ7_(ss, syncRows);
    } catch (eSync) {}
    return { success: true, rows: written.rows || 0, updated: written.updated || 0, appended: written.appended || 0 };
  } finally {
    try { lock.releaseLock(); } catch (eL) {}
  }
}


/**
 * Admin: lưu danh sách mã con thuộc Parent_SKU.
 * payload: { parentSku, variants:[{maHang, maVach, tenHang, dvt, tonBanDau}], actor }
 */
function saveChildVariants_(payload) {
  try {
    var parentSku = String(payload && payload.parentSku || "").trim().toUpperCase();
    var variants = (payload && payload.variants) ? payload.variants : [];
    if (!parentSku) return { success: false, error: "Thiếu Parent_SKU." };
    if (!variants.length) return { success: false, error: "Danh sách mã con trống." };

    var ss = getSS();
    var sh = getOrCreateCatalogSheet(ss);
    var lastRow = sh.getLastRow();
    var width = Math.max(sh.getLastColumn(), CATALOG_COL_COUNT);
    var values = lastRow >= 1 ? sh.getRange(1, 1, Math.max(lastRow, 1), width).getValues() : [["Mã hàng", "", "Mã vạch", "", "", "Tên hàng hóa", "", "ĐVT", "Ngày tạo", CATALOG_PARENT_HEADER, CATALOG_ISNEW_HEADER]];
    var header = values[0] || [];
    var parentIdx = findCatalogParentColIdx_(header);
    if (parentIdx === -1) parentIdx = 9;
    var mhIdx = findCatalogMaHangColIdx_(header, parentIdx);
    if (mhIdx === -1) mhIdx = 0;
    var mvIdx = findColumnIndexByAliases(header, ["mavach", "barcode", "barcodeid"]);
    if (mvIdx === -1) mvIdx = 2;
    var thIdx = findColumnIndexByAliases(header, ["tenhang", "name", "description"]);
    if (thIdx === -1) thIdx = 5;
    var dvtIdx = findColumnIndexByAliases(header, ["dvt", "donvitinh", "donvi", "unit"]);
    if (dvtIdx === -1) dvtIdx = 7;

    var byMh = {};
    for (var r = 1; r < values.length; r++) {
      var mh0 = String(values[r][mhIdx] == null ? "" : values[r][mhIdx]).trim().toUpperCase();
      if (mh0 && byMh[mh0] === undefined) byMh[mh0] = r;
    }

    var changed = 0;
    var appended = 0;
    var tonEntries = [];
    var stamp = catalogNowStamp_();

    for (var i = 0; i < variants.length; i++) {
      var v = variants[i];
      if (!v) continue;
      var mh = String(v.maHang || "").trim();
      var mv = String(v.maVach || "").trim();
      var th = String(v.tenHang || "").trim();
      var dvt = String(v.dvt || "").trim();
      var tonBanDau = Number(v.tonBanDau);
      if (isNaN(tonBanDau)) tonBanDau = 0;
      if (!mh) continue;
      var mhU = mh.toUpperCase();
      var rowIdx = byMh[mhU];
      if (rowIdx !== undefined) {
        values[rowIdx][mhIdx] = mh;
        values[rowIdx][mvIdx] = mv;
        if (th) values[rowIdx][thIdx] = th;
        values[rowIdx][dvtIdx] = dvt;
        values[rowIdx][parentIdx] = parentSku;
        changed++;
      } else {
        var newRow = [];
        for (var c = 0; c < width; c++) newRow.push("");
        newRow[mhIdx] = mh;
        newRow[mvIdx] = mv;
        newRow[thIdx] = th;
        newRow[dvtIdx] = dvt;
        if (width > 8) newRow[8] = stamp.date;
        newRow[parentIdx] = parentSku;
        values.push(newRow);
        byMh[mhU] = values.length - 1;
        appended++;
      }
      tonEntries.push({
        maHang: mh,
        maVach: mv,
        tenHang: th,
        dvt: dvt,
        parentSku: parentSku,
        tonBanDau: tonBanDau
      });
    }

    // Ghi lại Data_Excel (batch)
    var outW = Math.max(width, CATALOG_COL_COUNT, parentIdx + 1);
    while (header.length < outW) header.push("");
    if (!String(header[parentIdx] || "").trim()) header[parentIdx] = CATALOG_PARENT_HEADER;
    values[0] = header;
    for (var rr = 0; rr < values.length; rr++) {
      while (values[rr].length < outW) values[rr].push("");
    }
    sh.clear();
    sh.getRange(1, 1, values.length, outW).setValues(values);
    try { SpreadsheetApp.flush(); } catch (eF) {}
    invalidateCatalogCache_();

    var tonRes = upsertTonVariantKeepExport_(ss, tonEntries);
    return {
      success: true,
      parentSku: parentSku,
      changed: changed,
      appended: appended,
      variantCount: tonEntries.length,
      tonRows: tonRes && tonRes.rows,
      msg: "Đã lưu " + tonEntries.length + " mã con cho Parent " + parentSku +
        " (cập nhật " + changed + ", thêm mới " + appended + ")."
    };
  } catch (e) {
    return { success: false, error: e.message || String(e) };
  }
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


function getStockMapForStore(ss, storeName) {
  if (!storeName) return {};
  // Ưu tiên sheet nhẹ TON_Q7 khi hỏi tồn kho Q7
  if (isPackingQ7Store_(storeName)) {
    var light = readTonKhoQ7Map_(ss);
    if (light && Object.keys(light).length) return light;
  }
  return getStockMapForStoreFromFullSheet_(ss, storeName);
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
