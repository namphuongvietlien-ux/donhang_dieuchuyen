// ============================================================
// catalog_variant.gs — Data_Excel / Parent_SKU / TON_VARIANT / Hang moi
// ============================================================

// Kho soạn hàng chính — sheet nhẹ chỉ chứa tồn Q7 (tạo lúc import file tồn)
var PACKING_STOCK_STORE = "Kho Địa điểm kinh doanh Q7";

var TON_Q7_SHEET_NAME = "TON_Q7";

var CACHE_TON_Q7_KEY = "ton_q7_map_v7_sheet_bare";

// Tồn riêng theo biến thể đồ chơi — đối soát: Ton_Ban_Dau - Da_Xuat + Da_Nhan_Nhap = Ton_Hien_Tai
// MASTER Cha–Con + stock biến thể: KHÔNG được sheet.clear() khi import MISA / tồn
var TON_VARIANT_SHEET_NAME = "TON_VARIANT";

var CACHE_TON_VARIANT_KEY = "ton_variant_map_v4_stock_fix";

/** Staging thô từ MISA — được phép ghi đè mỗi lần import */
var MISA_IMPORT_SHEET_NAME = "MISA_IMPORT";

var TON_VARIANT_COL_COUNT = 12;

var TON_VARIANT_HEADERS = [
  "Key",            // A: mã SP (không gắn |ĐVT)
  "Parent_SKU",     // B
  "Ton_Ban_Dau",    // C
  "Da_Xuat",        // D
  "Ton_Hien_Tai",   // E = C - D + F
  "Da_Nhan_Nhap",   // F
  "TenSP_ChiTiet",  // G
  "MaVach",         // H
  "DonViTinh",      // I
  "UpdatedAt",      // J
  "IsLocked",       // K: TRUE/FALSE — khóa đặt hàng (giữ khi import MISA)
  "IsOutStock"      // L: TRUE/FALSE — hết hàng (giữ khi import MISA; hết hàng → gỡ IsNew)
];

var CATALOG_PARENT_HEADER = "Parent_SKU";

var CATALOG_ISNEW_HEADER = "IsNew";

var CATALOG_COL_COUNT = 12; // A..L: Mã hàng | … | Mã vạch | … | Tên | … | ĐVT | Ngày tạo | Parent_SKU | IsNew | DonViTinh2
var CATALOG_DVT2_HEADER = "DonViTinh2";


/** Parse cờ TRUE/FALSE / 1 / x trên sheet */
function parseTonVariantLocked_(raw) {
  var s = String(raw == null ? "" : raw).trim().toLowerCase();
  return s === "1" || s === "true" || s === "yes" || s === "x" || s === "locked" || s === "khoa" ||
    s === "out" || s === "hethang" || s === "het hang" || s === "outofstock";
}


function tonVariantLockedCell_(isLocked) {
  return isLocked ? "TRUE" : "FALSE";
}


function parseTonVariantOutStock_(raw) {
  return parseTonVariantLocked_(raw);
}


function tonVariantOutStockCell_(isOutStock) {
  return isOutStock ? "TRUE" : "FALSE";
}


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
  // Gộp Key trùng (MH:X vs MH:X|DV:cai) → 1 dòng / sản phẩm; ĐVT ở cột Dvt
  // Cột A sheet: CHỈ mã SP (không lưu tiền tố MH:/MV:)
  var merged = {};
  var mergedDvt = {};
  var _dbgStripSamples = 0;
  for (var k in map) {
    if (!Object.prototype.hasOwnProperty.call(map, k)) continue;
    if (k === "__meta") continue;
    var ck = canonicalizeStockSheetKey_(k);
    if (!ck) continue;
    var ckBeforeStrip = ck;
    ck = String(ck).replace(/^MH:/i, "").replace(/^MV:/i, "").trim();
    if (!ck) continue;
    // #region agent log
    if (_dbgStripSamples < 3 && ckBeforeStrip !== ck) {
      _dbgStripSamples++;
      try {
        console.log(JSON.stringify({
          sessionId: "f6b0dc",
          hypothesisId: "H-write-mh",
          location: "catalog_variant.gs:writeTonQ7MapToSheet_",
          message: "TON_Q7 write strip MH",
          data: { from: ckBeforeStrip, to: ck },
          timestamp: Date.now()
        }));
      } catch (_dbgW) {}
    }
    // #endregion
    var qty = Number(map[k]) || 0;
    var wasSuffixed = String(k).indexOf("|") !== -1;
    if (merged[ck] === undefined || !wasSuffixed) {
      merged[ck] = qty;
    } else if (!(Number(merged[ck]) > 0) && qty > 0) {
      merged[ck] = qty;
    }
    var dvtLabel = dvtLabelByKey[k] || dvtFromStockKey_(k) || dvtLabelByKey[ck] || dvtLabelByKey[ckBeforeStrip] || "";
    if (dvtLabel) mergedDvt[ck] = dvtLabel;
    else if (!mergedDvt[ck]) mergedDvt[ck] = "";
  }
  map = merged;
  dvtLabelByKey = mergedDvt;

  var sh = ss.getSheetByName(TON_Q7_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(TON_Q7_SHEET_NAME);
  sh.clear();
  sh.getRange(1, 1, 1, 4).setValues([["Key", "Qty", "Dvt", "UpdatedAt"]]);
  var rows = [];
  for (var mk in map) {
    if (!Object.prototype.hasOwnProperty.call(map, mk)) continue;
    rows.push([mk, Number(map[mk]) || 0, dvtLabelByKey[mk] || "", ""]);
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
    _debugRun: "q7-v5-sheet-bare-key"
  };
}


/** Ghi TON_Q7 từ entries {k,q,d} — Key Cột A = mã SP (không MH:/|DV:) */
function writeTonQ7EntriesToSheet_(ss, entries) {
  var map = {};
  var dvtLabels = {};
  for (var i = 0; i < (entries || []).length; i++) {
    var ent = entries[i];
    if (!ent || !ent.k) continue;
    var rawKey = String(ent.k).trim();
    if (!rawKey) continue;
    var key = canonicalizeStockSheetKey_(rawKey);
    if (!key) continue;
    key = String(key).replace(/^MH:/i, "").replace(/^MV:/i, "").trim();
    if (!key) continue;
    map[key] = (Number(map[key]) || 0) + (Number(ent.q) || 0);
    var dLabel = String(ent.d || "").trim() || dvtFromStockKey_(rawKey);
    if (dLabel) dvtLabels[key] = dLabel;
    else if (!dvtLabels[key]) dvtLabels[key] = dvtFromStockKey_(key);
  }
  return writeTonQ7MapToSheet_(ss, map, dvtLabels);
}


/**
 * Key chuẩn TON_VARIANT: CHỈ mã SP (không gắn |ĐVT).
 * ĐVT lưu cột DonViTinh — tránh sinh 2 dòng Parent và Parent|cai.
 */
function buildTonVariantKey_(maHang, dvt) {
  return canonicalizeTonVariantKey_(maHang);
}


function calcTonHienTaiVariant_(tonBanDau, daXuat, daNhanNhap) {
  var raw = (Number(tonBanDau) || 0) - (Number(daXuat) || 0) + (Number(daNhanNhap) || 0);
  if (isNaN(raw)) return 0;
  return Math.max(0, raw);
}


/** Ép số tồn an toàn — giữ giá trị thật (vd 330); rỗng → 0 */
function parseTonStockNumber_(raw) {
  if (raw === "" || raw === null || raw === undefined) return 0;
  if (typeof raw === "number" && isFinite(raw)) return raw;
  var s = String(raw).trim().replace(/\u00A0/g, "").replace(/,/g, "");
  if (!s) return 0;
  var n = Number(s);
  return isFinite(n) ? n : 0;
}


/**
 * Dynamic header map cho TON_VARIANT — tìm cột theo tên (TonKho / Ton_Hien_Tai / Stock…),
 * không phụ thuộc cứng index [4]/[5].
 */
function resolveTonVariantColMap_(shOrHeader) {
  var defaults = {
    key: 0,
    parent: 1,
    tonBanDau: 2,
    daXuat: 3,
    tonHienTai: 4,
    daNhanNhap: 5,
    ten: 6,
    maVach: 7,
    dvt: 8,
    updatedAt: 9,
    isLocked: 10,
    isOutStock: 11
  };
  var header = null;
  if (Object.prototype.toString.call(shOrHeader) === "[object Array]") {
    header = shOrHeader;
  } else if (shOrHeader && typeof shOrHeader.getRange === "function") {
    try {
      var lastCol = Math.max(shOrHeader.getLastColumn(), TON_VARIANT_COL_COUNT);
      header = shOrHeader.getRange(1, 1, 1, lastCol).getValues()[0] || [];
    } catch (eH) {
      header = [];
    }
  }
  var map = {};
  for (var k in defaults) {
    if (Object.prototype.hasOwnProperty.call(defaults, k)) map[k] = defaults[k];
  }
  if (!header || !header.length) return map;

  for (var c = 0; c < header.length; c++) {
    var h = normalizeHeaderText(header[c]);
    if (!h) continue;
    if (h === "key" || h === "mahang" || h === "masp" || h === "sku" || h === "ma") map.key = c;
    else if (h.indexOf("parent") !== -1) map.parent = c;
    else if (h.indexOf("tonbandau") !== -1 || h === "bandau" || h === "opening") map.tonBanDau = c;
    else if (h.indexOf("daxuat") !== -1 || h === "exported" || h === "sold") map.daXuat = c;
    else if (
      h.indexOf("tonhientai") !== -1 || h === "tonkho" || h.indexOf("soluongton") !== -1 ||
      h === "stock" || h === "onhand" || h === "slton" || h === "ton"
    ) map.tonHienTai = c;
    else if (h.indexOf("danhan") !== -1 || h.indexOf("nhapnhap") !== -1 || h === "received") map.daNhanNhap = c;
    else if (h.indexOf("tensp") !== -1 || h.indexOf("tenhang") !== -1 || h.indexOf("chitiet") !== -1) map.ten = c;
    else if (h.indexOf("mavach") !== -1 || h.indexOf("barcode") !== -1) map.maVach = c;
    else if (h === "donvitinh" || h === "dvt" || h === "unit" || h.indexOf("donvi") === 0) map.dvt = c;
    else if (h.indexOf("updated") !== -1) map.updatedAt = c;
    else if (h.indexOf("islocked") !== -1 || h === "locked" || h === "khoa") map.isLocked = c;
    else if (h.indexOf("isoutstock") !== -1 || h.indexOf("hethang") !== -1 || h === "outofstock") map.isOutStock = c;
  }
  return map;
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


/** Migrate sheet cũ → schema đối soát 12 cột (IsLocked + IsOutStock) */
function ensureTonVariantSchema_(sh) {
  if (!sh) return;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var header = sh.getRange(1, 1, 1, Math.max(lastCol, TON_VARIANT_COL_COUNT)).getValues()[0] || [];
  var h0 = normalizeHeaderText(header[0]);
  var h1 = normalizeHeaderText(header[1]);
  var h2 = normalizeHeaderText(header[2]);
  var alreadyNew = h1.indexOf("parent") !== -1 && (h2.indexOf("tonbandau") !== -1 || h2.indexOf("bandau") !== -1);
  if (alreadyNew) {
    // Đảm bảo header chuẩn 12 cột; không đụng dữ liệu — IsLocked / IsOutStock cũ giữ nguyên
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
      // Key chuẩn = mã SP (bỏ |ĐVT / |DV:); ĐVT giữ cột DonViTinh
      var keyNorm = canonicalizeTonVariantKey_(k);
      if (!keyNorm) continue;
      var dvtOld = String(old[i][5] || "").trim() || dvtFromStockKey_(k);
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
        dvtOld,
        old[i][6] || new Date(),
        "FALSE",
        "FALSE"
      ]);
    }
  }
  sh.clear();
  sh.getRange(1, 1, 1, TON_VARIANT_COL_COUNT).setValues([TON_VARIANT_HEADERS])
    .setFontWeight("bold").setBackground("#cfe2f3");
  if (migrated.length) {
    // Dedupe theo Key chuẩn khi migrate
    var migMap = {};
    for (var mi = 0; mi < migrated.length; mi++) {
      var mk = String(migrated[mi][0] || "").trim();
      if (!mk || migMap[mk]) continue;
      migMap[mk] = migrated[mi];
    }
    var migRows = [];
    for (var mk2 in migMap) {
      if (Object.prototype.hasOwnProperty.call(migMap, mk2)) migRows.push(migMap[mk2]);
    }
    if (migRows.length) sh.getRange(2, 1, migRows.length, TON_VARIANT_COL_COUNT).setValues(migRows);
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
  // Index theo Key chuẩn + Key thô (legacy CODE|dvt) → cùng 1 dòng sheet
  var rowIndexByKey = {};
  if (lastRow >= 2) {
    var existingKeys = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < existingKeys.length; i++) {
      var ek = String(existingKeys[i][0] || "").trim();
      if (!ek) continue;
      var sheetRowNum = i + 2;
      rowIndexByKey[ek] = sheetRowNum;
      var ck = canonicalizeTonVariantKey_(ek);
      if (ck && rowIndexByKey[ck] === undefined) rowIndexByKey[ck] = sheetRowNum;
    }
  }

  var now = new Date();
  var appendRows = [];
  var updated = 0;
  var seenCanon = {};
  for (var k in dirtyByKey) {
    if (!Object.prototype.hasOwnProperty.call(dirtyByKey, k)) continue;
    var row = dirtyByKey[k];
    if (!row || !row.k) continue;
    var canonK = canonicalizeTonVariantKey_(row.k);
    if (!canonK) continue;
    if (seenCanon[canonK]) continue;
    seenCanon[canonK] = true;
    row.k = canonK;
    if (!String(row.d || "").trim()) row.d = dvtFromStockKey_(k) || "";
    row.tonHienTai = calcTonHienTaiVariant_(row.tonBanDau, row.daXuat, row.daNhanNhap);
    var locked = row.isLocked === true;
    var outStock = row.isOutStock === true;
    var vals = [
      canonK,
      row.p || "",
      Number(row.tonBanDau) || 0,
      Number(row.daXuat) || 0,
      Number(row.tonHienTai) || 0,
      Number(row.daNhanNhap) || 0,
      row.th || "",
      row.mv || "",
      row.d || "",
      now,
      tonVariantLockedCell_(locked),
      tonVariantOutStockCell_(outStock)
    ];
    var sheetRow = rowIndexByKey[canonK];
    if (sheetRow === undefined) sheetRow = rowIndexByKey[k];
    if (sheetRow) {
      sh.getRange(sheetRow, 1, 1, TON_VARIANT_COL_COUNT).setValues([vals]);
      updated++;
    } else {
      appendRows.push(vals);
      rowIndexByKey[canonK] = -1; // đánh dấu đã append trong batch
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
 * Gộp Key legacy (SKU|dvt) vào Key chuẩn (SKU).
 */
function readTonVariantByKeyMap_(ss) {
  ss = ss || getSS();
  var byKey = {};
  var sh = getOrCreateTonVariantSheet_(ss);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return byKey;
  ensureTonVariantSchema_(sh);
  var cols = resolveTonVariantColMap_(sh);
  var existing = sh.getRange(2, 1, lastRow - 1, TON_VARIANT_COL_COUNT).getValues();
  for (var r = 0; r < existing.length; r++) {
    var row = existing[r];
    if (!row) continue;
    var ek = String(row[cols.key] || "").trim();
    if (!ek) continue;
    var ck = canonicalizeTonVariantKey_(ek);
    if (!ck) continue;
    var dCol = String(row[cols.dvt] || "").trim() || dvtFromStockKey_(ek);
    var lockedFlag = parseTonVariantLocked_(row.length > cols.isLocked ? row[cols.isLocked] : "");
    var outStockFlag = parseTonVariantOutStock_(row.length > cols.isOutStock ? row[cols.isOutStock] : "");
    var tonBanDau = parseTonStockNumber_(row[cols.tonBanDau]);
    var daXuat = parseTonStockNumber_(row[cols.daXuat]);
    var daNhanNhap = parseTonStockNumber_(row[cols.daNhanNhap]);
    var tonCell = row[cols.tonHienTai];
    var hasTonCell = !(tonCell === "" || tonCell === null || tonCell === undefined);
    var tonHienTai = hasTonCell
      ? parseTonStockNumber_(tonCell)
      : calcTonHienTaiVariant_(tonBanDau, daXuat, daNhanNhap);
    var cand = {
      k: ck,
      p: String(row[cols.parent] || "").trim(),
      tonBanDau: tonBanDau,
      daXuat: daXuat,
      tonHienTai: tonHienTai,
      daNhanNhap: daNhanNhap,
      th: String(row[cols.ten] || "").trim(),
      mv: String(row[cols.maVach] || "").trim(),
      d: dCol,
      isLocked: lockedFlag,
      isOutStock: outStockFlag
    };
    var prev = byKey[ck];
    if (!prev) {
      byKey[ck] = cand;
      continue;
    }
    // Ưu tiên dòng Key không hậu tố; gộp Da_Xuat / meta / IsLocked / IsOutStock (OR)
    var prevWasBare = String(prev.k) === ck && ek.indexOf("|") === -1;
    if (!prevWasBare && ek.indexOf("|") === -1) {
      cand.daXuat = Math.max(cand.daXuat, prev.daXuat);
      cand.daNhanNhap = Math.max(cand.daNhanNhap, prev.daNhanNhap);
      if (!cand.p && prev.p) cand.p = prev.p;
      if (!cand.th && prev.th) cand.th = prev.th;
      if (!cand.mv && prev.mv) cand.mv = prev.mv;
      if (!cand.d && prev.d) cand.d = prev.d;
      cand.isLocked = !!(cand.isLocked || prev.isLocked);
      cand.isOutStock = !!(cand.isOutStock || prev.isOutStock);
      if (!cand.tonHienTai && (cand.tonBanDau || cand.daXuat || cand.daNhanNhap)) {
        cand.tonHienTai = calcTonHienTaiVariant_(cand.tonBanDau, cand.daXuat, cand.daNhanNhap);
      }
      byKey[ck] = cand;
    } else {
      prev.daXuat = Math.max(prev.daXuat, cand.daXuat);
      prev.daNhanNhap = Math.max(prev.daNhanNhap, cand.daNhanNhap);
      if (!prev.p && cand.p) prev.p = cand.p;
      if (!prev.th && cand.th) prev.th = cand.th;
      if (!prev.mv && cand.mv) prev.mv = cand.mv;
      if (!prev.d && cand.d) prev.d = cand.d;
      prev.isLocked = !!(prev.isLocked || cand.isLocked);
      prev.isOutStock = !!(prev.isOutStock || cand.isOutStock);
      prev.k = ck;
      // Giữ tonHienTai lớn hơn khi gộp (không đè số thật bằng 0)
      if (cand.tonHienTai > prev.tonHienTai) prev.tonHienTai = cand.tonHienTai;
      else if (!prev.tonHienTai && (prev.tonBanDau || prev.daXuat || prev.daNhanNhap)) {
        prev.tonHienTai = calcTonHienTaiVariant_(prev.tonBanDau, prev.daXuat, prev.daNhanNhap);
      }
    }
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
      var rawKey = String(ent.k || "").trim();
      var dFromKey = dvtFromStockKey_(rawKey);
      var key = canonicalizeTonVariantKey_(rawKey || ent.maHang || ent.mh || "");
      if (!key) {
        key = buildTonVariantKey_(ent.maHang || ent.mh || "", ent.d || ent.dvt || "");
      }
      if (!key) continue;
      var qtyRaw = ent.q != null ? ent.q : ent.qty;
      var hasQty = qtyRaw !== "" && qtyRaw !== null && qtyRaw !== undefined;
      var qty = hasQty ? Number(qtyRaw) : NaN;
      if (hasQty && isNaN(qty)) qty = 0;

      var prev = byKey[key] || {
        k: key, p: "", tonBanDau: 0, daXuat: 0, tonHienTai: 0, daNhanNhap: 0, th: "", mv: "", d: "", isLocked: false, isOutStock: false
      };
      prev.k = key;
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
      var d = String(ent.d || ent.dvt || "").trim() || dFromKey;
      if (p) prev.p = p;
      if (th) prev.th = th;
      if (mv) prev.mv = mv;
      if (d) prev.d = d;
      // BẮT BUỘC GIỮ IsLocked / IsOutStock khi import MISA/tồn — chỉ ghi đè khi payload gửi rõ
      if (ent.isLocked !== undefined && ent.isLocked !== null && ent.isLocked !== "") {
        prev.isLocked = ent.isLocked === true || parseTonVariantLocked_(ent.isLocked);
      } else if (prev.isLocked === undefined) {
        prev.isLocked = false;
      }
      if (ent.isOutStock !== undefined && ent.isOutStock !== null && ent.isOutStock !== "") {
        prev.isOutStock = ent.isOutStock === true || parseTonVariantOutStock_(ent.isOutStock);
      } else if (prev.isOutStock === undefined) {
        prev.isOutStock = false;
      }
      byKey[key] = prev;
      dirty[key] = prev;
    }

    var written = persistTonVariantByKeyNoClear_(ss, dirty);
    try {
      removeDuplicateStockRows_(ss, { skipLock: true });
    } catch (eDed) { Logger.log(eDed); }
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
      _debugRun: "ton-variant-upsert-noclear-v2-key-canon"
    };
  } finally {
    try { lock.releaseLock(); } catch (eL) {}
  }
}


/** Map lookup (MH:/MV:) từ Ton_Hien_Tai để UI/API dùng getStockValueForItem */
function buildTonVariantStockMapFromRows_(rows, colMap) {
  var cols = colMap || {
    key: 0, parent: 1, tonBanDau: 2, daXuat: 3, tonHienTai: 4,
    daNhanNhap: 5, ten: 6, maVach: 7, dvt: 8
  };
  var map = {};
  for (var i = 0; i < (rows || []).length; i++) {
    var r = rows[i];
    if (!r) continue;
    var key = String(r[cols.key] || "").trim();
    if (!key) continue;
    var tonBanDau = parseTonStockNumber_(r[cols.tonBanDau]);
    var daXuat = parseTonStockNumber_(r[cols.daXuat]);
    var daNhanNhap = parseTonStockNumber_(r[cols.daNhanNhap]);
    var tonCell = r[cols.tonHienTai];
    var hasTonCell = !(tonCell === "" || tonCell === null || tonCell === undefined);
    var ton = hasTonCell
      ? parseTonStockNumber_(tonCell)
      : calcTonHienTaiVariant_(tonBanDau, daXuat, daNhanNhap);
    var mv = String(r[cols.maVach] || "").trim();
    var dvt = String(r[cols.dvt] || "").trim();
    map[key] = ton;
    var mh = canonicalizeTonVariantKey_(key) || key.split("|")[0];
    var dv = normalizeDvtKey_(dvt) || (key.indexOf("|") !== -1 ? key.split("|").slice(1).join("|") : "");
    if (mh) {
      // Bare key + MH: — để lookup không phụ thuộc ĐVT lệch chuẩn hoá
      map[mh] = ton;
      addStockValueByCode(map, "MH:", mh, ton, "");
      if (dv || dvt) addStockValueByCode(map, "MH:", mh, ton, dv || dvt);
    }
    if (mv) {
      addStockValueByCode(map, "MV:", normalizeProductCode(mv), ton, "");
      if (dv || dvt) addStockValueByCode(map, "MV:", normalizeProductCode(mv), ton, dv || dvt);
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
  var cols = resolveTonVariantColMap_(sh);
  // getRange(row, column, numRows, numColumns)
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, TON_VARIANT_COL_COUNT).getValues();
  var map = buildTonVariantStockMapFromRows_(data, cols);
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
    var ck0 = canonicalizeTonVariantKey_(k0);
    if (ck0 && byKey[ck0] === undefined) byKey[ck0] = r;
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

    // Chuẩn hóa Key trên sheet nếu còn dạng legacy SKU|dvt
    var canonRowKey = canonicalizeTonVariantKey_(values[idx][0]);
    if (canonRowKey) values[idx][0] = canonRowKey;

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
  var seenChild = {};
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
      if (!mh && key && key.indexOf("|") !== -1) {
        mh = key.split("|")[0];
        if (!dvt) dvt = dvtFromStockKey_(key);
      } else if (!mh && key) mh = key;
    }

    var cat = resolveCatalogProduct(catalogLookup, mh, mv);
    if (!cat || !cat.parentSku) continue;
    var childMh = cat.maHang || mh;
    if (!childMh) continue;
    var vKey = buildTonVariantKey_(childMh, dvt || cat.dvt || "");
    if (!vKey) continue;
    if (Object.prototype.hasOwnProperty.call(seenChild, vKey)) {
      // Cùng mã SP → cộng dồn qty, không tạo dòng 2
      seenChild[vKey].q = (Number(seenChild[vKey].q) || 0) + qty;
      continue;
    }
    var rowEnt = {
      maHang: childMh,
      k: vKey,
      q: qty,
      p: cat.parentSku,
      th: cat.tenHang || ent.th || "",
      mv: cat.maVach || mv || "",
      d: dvt || cat.dvt || ""
    };
    seenChild[vKey] = rowEnt;
    variantEntries.push(rowEnt);
  }
  if (!variantEntries.length) return { rows: 0 };
  var written = writeTonVariantEntriesToSheet_(ss, variantEntries);
  return { rows: written.rows || variantEntries.length, ms: written.ms || 0 };
}


/**
 * Dọn dòng trùng TON_VARIANT: gộp Key `SKU` và `SKU|ĐVT` thành 1 dòng Key = SKU.
 * Giữ Da_Xuat / Da_Nhan lớn nhất; ĐVT đưa vào cột DonViTinh.
 */
function dedupeTonVariantSheet_(ss) {
  ss = ss || getSS();
  var sh = getOrCreateTonVariantSheet_(ss);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { kept: 0, removed: 0 };
  ensureTonVariantSchema_(sh);
  var data = sh.getRange(2, 1, lastRow - 1, TON_VARIANT_COL_COUNT).getValues();
  var byCanon = {};
  var order = [];
  var removed = 0;
  for (var i = 0; i < data.length; i++) {
    var raw = String(data[i][0] || "").trim();
    if (!raw) { removed++; continue; }
    var ck = canonicalizeTonVariantKey_(raw);
    if (!ck) { removed++; continue; }
    var dCol = String(data[i][8] || "").trim() || dvtFromStockKey_(raw);
    var isBare = raw.indexOf("|") === -1;
    var lockedCell = tonVariantLockedCell_(parseTonVariantLocked_(data[i].length > 10 ? data[i][10] : ""));
    var outStockCell = tonVariantOutStockCell_(parseTonVariantOutStock_(data[i].length > 11 ? data[i][11] : ""));
    var cand = [
      ck,
      String(data[i][1] || "").trim(),
      Number(data[i][2]) || 0,
      Number(data[i][3]) || 0,
      Number(data[i][4]) || 0,
      Number(data[i][5]) || 0,
      String(data[i][6] || "").trim(),
      String(data[i][7] || "").trim(),
      dCol,
      data[i][9] || new Date(),
      lockedCell,
      outStockCell
    ];
    if (!Object.prototype.hasOwnProperty.call(byCanon, ck)) {
      byCanon[ck] = cand;
      order.push(ck);
      continue;
    }
    removed++;
    var prev = byCanon[ck];
    // Ưu tiên dòng Key trần (không |ĐVT); gộp xuất/nhập; IsLocked / IsOutStock = OR
    var lockedMerged = tonVariantLockedCell_(
      parseTonVariantLocked_(prev[10]) || parseTonVariantLocked_(cand[10])
    );
    var outStockMerged = tonVariantOutStockCell_(
      parseTonVariantOutStock_(prev[11]) || parseTonVariantOutStock_(cand[11])
    );
    if (isBare) {
      cand[3] = Math.max(Number(cand[3]) || 0, Number(prev[3]) || 0);
      cand[5] = Math.max(Number(cand[5]) || 0, Number(prev[5]) || 0);
      if (!cand[1] && prev[1]) cand[1] = prev[1];
      if (!cand[6] && prev[6]) cand[6] = prev[6];
      if (!cand[7] && prev[7]) cand[7] = prev[7];
      if (!cand[8] && prev[8]) cand[8] = prev[8];
      cand[10] = lockedMerged;
      cand[11] = outStockMerged;
      cand[4] = calcTonHienTaiVariant_(cand[2], cand[3], cand[5]);
      byCanon[ck] = cand;
    } else {
      prev[3] = Math.max(Number(prev[3]) || 0, Number(cand[3]) || 0);
      prev[5] = Math.max(Number(prev[5]) || 0, Number(cand[5]) || 0);
      if (!prev[1] && cand[1]) prev[1] = cand[1];
      if (!prev[6] && cand[6]) prev[6] = cand[6];
      if (!prev[7] && cand[7]) prev[7] = cand[7];
      if (!prev[8] && cand[8]) prev[8] = cand[8];
      prev[0] = ck;
      prev[10] = lockedMerged;
      prev[11] = outStockMerged;
      prev[4] = calcTonHienTaiVariant_(prev[2], prev[3], prev[5]);
    }
  }
  var out = [];
  for (var o = 0; o < order.length; o++) {
    if (byCanon[order[o]]) out.push(byCanon[order[o]]);
  }
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, TON_VARIANT_COL_COUNT).clearContent();
  if (out.length) sh.getRange(2, 1, out.length, TON_VARIANT_COL_COUNT).setValues(out);
  try { SpreadsheetApp.flush(); } catch (e) {}
  try { getScriptCache_().remove(CACHE_TON_VARIANT_KEY); } catch (e2) {}
  return { kept: out.length, removed: removed };
}


/**
 * Dọn dòng trùng TON_Q7: gộp MH:X và MH:X|DV:… thành 1 Key chuẩn.
 */
function dedupeTonQ7Sheet_(ss) {
  ss = ss || getSS();
  var sh = ss.getSheetByName(TON_Q7_SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) return { kept: 0, removed: 0 };
  var lastRow = sh.getLastRow();
  var lastCol = Math.max(sh.getLastColumn(), 4);
  var data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var map = {};
  var labels = {};
  var removed = 0;
  var seen = 0;
  for (var i = 0; i < data.length; i++) {
    var raw = String(data[i][0] || "").trim();
    if (!raw) { removed++; continue; }
    var ck = canonicalizeStockSheetKey_(raw);
    if (!ck) { removed++; continue; }
    var qty = Number(data[i][1]) || 0;
    var dvt = String(data[i][2] || "").trim() || dvtFromStockKey_(raw);
    var wasSuffixed = raw.indexOf("|") !== -1;
    if (!Object.prototype.hasOwnProperty.call(map, ck)) {
      map[ck] = qty;
      if (dvt) labels[ck] = dvt;
      seen++;
    } else {
      removed++;
      if (!wasSuffixed) map[ck] = qty;
      else if (!(Number(map[ck]) > 0) && qty > 0) map[ck] = qty;
      if (dvt && !labels[ck]) labels[ck] = dvt;
    }
  }
  writeTonQ7MapToSheet_(ss, map, labels);
  return { kept: Object.keys(map).length, removed: removed };
}


/**
 * API/Admin: lọc bỏ Key thừa dạng Parent|ĐVT trên TON_VARIANT + TON_Q7.
 * Giữ 1 dòng chuẩn / sản phẩm (Key = mã SP hoặc MH:mã).
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} ss
 * @param {{skipLock?:boolean}=} opt — skipLock=true khi caller đã giữ DocumentLock
 */
function removeDuplicateStockRows_(ss, opt) {
  ss = ss || getSS();
  if (opt && opt.skipLock) return removeDuplicateStockRowsUnlocked_(ss);
  var lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    return removeDuplicateStockRowsUnlocked_(ss);
  } finally {
    try { lock.releaseLock(); } catch (eL) {}
  }
}


/**
 * Hàm public (không dấu _) — hiện trong dropdown Run của Apps Script editor.
 * Chạy tay: chọn removeDuplicateStockRows → Run.
 */
function removeDuplicateStockRows() {
  var res = removeDuplicateStockRows_(getSS());
  try {
    Logger.log(JSON.stringify(res));
  } catch (eLog) {}
  return res;
}


function removeDuplicateStockRowsUnlocked_(ss) {
  ss = ss || getSS();
  var variant = dedupeTonVariantSheet_(ss);
  var q7 = dedupeTonQ7Sheet_(ss);
  try {
    var shV = getOrCreateTonVariantSheet_(ss);
    var lr = shV.getLastRow();
    if (lr >= 2) {
      syncParentVariantTotalsToTonQ7_(ss, shV.getRange(2, 1, lr - 1, TON_VARIANT_COL_COUNT).getValues());
    }
  } catch (eSync) { Logger.log(eSync); }
  return {
    success: true,
    variantKept: variant.kept || 0,
    variantRemoved: variant.removed || 0,
    q7Kept: q7.kept || 0,
    q7Removed: q7.removed || 0,
    msg: "Đã dọn trùng: TON_VARIANT bỏ " + (variant.removed || 0) +
      " dòng, TON_Q7 bỏ " + (q7.removed || 0) + " dòng."
  };
}


/** Tổng Ton_Hien_Tai theo Parent_SKU → ghi vào TON_Q7 (Key = Parent, không MH:) */
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
    var q7Key = String(mh).replace(/^MH:/i, "").replace(/^MV:/i, "").trim() || mh;
    map[q7Key] = Number(parentSum[parent]) || 0;
    if (!labels[q7Key]) labels[q7Key] = "";
    updated++;
  }
  writeTonQ7MapToSheet_(ss, map, labels);
  return { updated: updated };
}


/** Trả stock variant (Ton_Hien_Tai) nếu mã có trong map; null nếu không — tra cứu O(1) */
function getVariantStockIfPresent_(map, maHang, maVach, dvt) {
  if (!map) return null;
  var mh = normalizeProductCode(maHang);
  var mv = normalizeProductCode(maVach);
  var dvtNorm = normalizeDvtKey_(dvt);
  var rawKey = buildTonVariantKey_(maHang, dvt);
  var canon = canonicalizeTonVariantKey_(maHang) || mh;
  var candidates = [];
  if (rawKey) candidates.push(rawKey);
  if (canon) candidates.push(canon);
  if (mh) {
    candidates.push(mh);
    candidates.push("MH:" + mh);
    if (dvtNorm) candidates.push("MH:" + mh + "|DV:" + dvtNorm);
  }
  if (mv) {
    candidates.push("MV:" + mv);
    if (dvtNorm) candidates.push("MV:" + mv + "|DV:" + dvtNorm);
  }
  var rawMh = String(maHang || "").trim();
  if (rawMh && rawMh !== mh) candidates.push(rawMh);

  for (var i = 0; i < candidates.length; i++) {
    var ck = candidates[i];
    if (!ck) continue;
    if (map[ck] !== undefined && map[ck] !== null && map[ck] !== "") {
      return parseTonStockNumber_(map[ck]);
    }
  }
  return null;
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
        isLocked: !!item.isLocked,
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
  // Ưu tiên In-Memory cache theo request (fifo_inventory.gs)
  try {
    var reqCache = getInventoryRequestCache_();
    if (reqCache && reqCache.tonQ7Bundle && reqCache.tonQ7Bundle.map) {
      return reqCache.tonQ7Bundle;
    }
  } catch (eReq) {}

  var cache = getScriptCache_();
  var cached = getCacheJson_(cache, CACHE_TON_Q7_KEY);
  var labelCached = getCacheJson_(cache, CACHE_TON_Q7_KEY + "_dvt");
  if (cached && typeof cached === "object" && Object.keys(cached).length) {
    var fromScriptCache = { map: cached, dvtLabels: labelCached || {}, source: "TON_Q7-cache" };
    try {
      var req2 = getInventoryRequestCache_();
      if (req2) req2.tonQ7Bundle = fromScriptCache;
    } catch (eSeed) {}
    return fromScriptCache;
  }

  var sh = ss.getSheetByName(TON_Q7_SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) return { map: null, dvtLabels: {}, source: "TON_Q7-missing" };
  var lastRow = sh.getLastRow();
  var lastCol = Math.max(sh.getLastColumn(), 3);
  var header = sh.getRange(1, 1, 1, lastCol).getValues()[0] || [];
  var keyIdx = 0;
  var qtyIdx = 1;
  var dvtIdx = 2;
  var foundQtyByName = false;
  for (var c = 0; c < header.length; c++) {
    var h = normalizeHeaderText(header[c]);
    if (!h) continue;
    if (h === "key" || h === "mahang" || h === "masp" || h === "parentsku" || h.indexOf("mahang") === 0) keyIdx = c;
    else if (
      h.indexOf("tonhientai") !== -1 || h.indexOf("tonq7") !== -1 || h === "tonkho" ||
      h.indexOf("soluongton") !== -1 || h === "stock" || h === "qty" || h === "quantity" || h === "soluong"
    ) {
      qtyIdx = c;
      foundQtyByName = true;
    } else if (h === "dvt" || h.indexOf("donvi") === 0 || h === "unit") dvtIdx = c;
  }
  if (!foundQtyByName) {
    // Schema chuẩn Key|Qty|Dvt — giữ mặc định; cảnh báo nếu header không rõ
    Logger.log("TON_Q7: không tìm thấy cột tồn theo tên — dùng cột index " + qtyIdx + " (Qty mặc định). Header=" + JSON.stringify(header));
  }
  var numRows = lastRow - 1;
  var data = sh.getRange(2, 1, numRows, lastCol).getValues();
  var map = {};
  var dvtLabels = {};
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!row) continue;
    // Cột A có thể đã chứa sẵn tiền tố MH: (vd MH:TAM1021) — strip trước khi build key
    var rawCode = String(row[keyIdx] || "").trim();
    if (!rawCode) continue;
    var cleanCode = String(rawCode || "").replace(/^MH:/i, "").trim();
    if (!cleanCode) continue;
    var key = cleanCode;
    // #region agent log
    if (i < 3) {
      try {
        console.log(JSON.stringify({
          sessionId: "f6b0dc",
          hypothesisId: "H-mh-prefix",
          location: "catalog_variant.gs:readTonKhoQ7Bundle_",
          message: "TON_Q7 colA clean",
          data: { rawCode: rawCode, cleanCode: cleanCode, stripped: /^MH:/i.test(rawCode) },
          timestamp: Date.now()
        }));
      } catch (_dbgMhClean) {}
    }
    // #endregion
    var rawQty = row[qtyIdx];
    if (typeof rawQty === "string") rawQty = String(rawQty).replace(/,/g, ".").trim();
    var qty = Number(rawQty);
    if (isNaN(qty)) qty = 0;
    qty = Math.max(0, qty);
    var dvtCol = row[dvtIdx] != null ? String(row[dvtIdx]).trim() : "";
    // Sheet cũ: Key không có |DV: nhưng cột Dvt có giá trị → gắn vào key lookup
    var lookupKey = key;
    if (dvtCol && lookupKey.indexOf("|DV:") === -1) {
      var dvtNorm = normalizeDvtKey_(dvtCol);
      if (dvtNorm) lookupKey = lookupKey + "|DV:" + dvtNorm;
    }
    map[lookupKey] = (Number(map[lookupKey]) || 0) + qty;
    // Alias bare / MH: — chỉ GÁN (không cộng lại qty; tránh 2× so với sheet)
    var bare = canonicalizeStockSheetKey_(key) || key;
    if (bare && bare !== lookupKey) {
      if (map[bare] === undefined) map[bare] = qty;
      else if (!(Number(map[bare]) > 0) && qty > 0) map[bare] = qty;
    }
    var mhOnly = String(bare || "").replace(/^MH:/i, "").replace(/^MV:/i, "");
    mhOnly = normalizeProductCode(mhOnly) || mhOnly;
    if (mhOnly) {
      var dvtNormAlias = normalizeDvtKey_(dvtCol || "");
      var mhKeyFull = "MH:" + mhOnly + (dvtNormAlias ? ("|DV:" + dvtNormAlias) : "");
      var mhKeyBare = "MH:" + mhOnly;
      if (mhKeyFull !== lookupKey) {
        if (map[mhKeyFull] === undefined) map[mhKeyFull] = qty;
        else if (!(Number(map[mhKeyFull]) > 0) && qty > 0) map[mhKeyFull] = qty;
      }
      if (mhKeyBare !== lookupKey && mhKeyBare !== mhKeyFull) {
        if (map[mhKeyBare] === undefined) map[mhKeyBare] = qty;
        else if (!(Number(map[mhKeyBare]) > 0) && qty > 0) map[mhKeyBare] = qty;
      }
    }
    if (dvtCol) dvtLabels[lookupKey] = dvtCol;
    else if (!dvtLabels[lookupKey]) dvtLabels[lookupKey] = dvtFromStockKey_(lookupKey);
  }
  if (!Object.keys(map).length) return { map: null, dvtLabels: {}, source: "TON_Q7-empty" };
  try {
    putCacheJson_(cache, CACHE_TON_Q7_KEY, map, CACHE_TTL_SECONDS);
    putCacheJson_(cache, CACHE_TON_Q7_KEY + "_dvt", dvtLabels, CACHE_TTL_SECONDS);
  } catch (e) {}
  var fromSheet = { map: map, dvtLabels: dvtLabels, source: "TON_Q7" };
  try {
    var req3 = getInventoryRequestCache_();
    if (req3) req3.tonQ7Bundle = fromSheet;
  } catch (eSeed3) {}
  return fromSheet;
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
  // Dynamic header map — tuyệt đối không hardcode cột ĐVT/Tên/MV
  var mappedCols = mapImportHeaderColumns_(headerRow || []);
  var parentSkuIdx = mappedCols.parentSku;
  if (parentSkuIdx === -1) parentSkuIdx = findCatalogParentColIdx_(headerRow);
  var maHangIdx = mappedCols.maHang;
  if (maHangIdx === -1) maHangIdx = findCatalogMaHangColIdx_(headerRow, parentSkuIdx);
  var maVachIdx = mappedCols.maVach;
  var tenHangIdx = mappedCols.tenHang;
  var dvtIdx = mappedCols.dvt;
  var dvt2Idx = mappedCols.dvt2;
  if (dvt2Idx === -1) dvt2Idx = findColumnIndexByAliases(headerRow, ["donvitinh2", "dvt2", "donviquydoi", "dvtphu"]);
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
  var tonIdx = mappedCols.tonKho;
  for (var k = startRow; k < rawData.length; k++) {
    if (!rawData[k]) continue;
    if (maHangIdx < 0 || tenHangIdx < 0) continue;
    var ma = getCellValue(rawData[k], maHangIdx, "").toUpperCase();
    var ten = getCellValue(rawData[k], tenHangIdx, "");
    if (ma !== "" && ten !== "") tenHangChuanTheoMa[ma] = ten;
  }
  var collected = [];
  for (var i = startRow; i < rawData.length; i++) {
    if (!rawData[i]) continue;
    // MaVach/MaHang luôn String — không Number() (mất số 0 đầu barcode)
    var maHang = maHangIdx >= 0 ? String(getCellValue(rawData[i], maHangIdx, "") || "").trim() : "";
    var maVach = maVachIdx >= 0 ? String(getCellValue(rawData[i], maVachIdx, "") || "").trim() : "";
    var tenHang = tenHangIdx >= 0 ? String(getCellValue(rawData[i], tenHangIdx, "") || "").trim() : "";
    var dvtRaw = dvtIdx >= 0 ? getCellValue(rawData[i], dvtIdx, "") : "";
    var dvt = sanitizeImportDvt_(dvtRaw);
    var dvt2 = dvt2Idx >= 0 ? sanitizeImportDvt_(getCellValue(rawData[i], dvt2Idx, "")) : "";
    var parentSku = parentSkuIdx !== -1 ? String(getCellValue(rawData[i], parentSkuIdx, "") || "").trim() : "";
    var tonRaw = tonIdx >= 0 ? getCellValue(rawData[i], tonIdx, "") : "";
    // Phân biệt ô trống (NaN → không gán tonKho) vs tồn = 0 thật
    var tonNum = (tonRaw === "" || tonRaw == null) ? NaN : parseTonStockNumber_(tonRaw);
    var isNewFlag = false;
    if (isNewIdx !== -1) {
      var flagRaw = String(rawData[i][isNewIdx] == null ? "" : rawData[i][isNewIdx]).trim().toLowerCase();
      isNewFlag = flagRaw === "1" || flagRaw === "true" || flagRaw === "yes" || flagRaw === "x" || flagRaw === "moi" || flagRaw === "new";
    }
    if (tenHang === "" && maHang !== "") tenHang = tenHangChuanTheoMa[String(maHang).toUpperCase()] || "";
    var mhNorm = normalizeProductCode(maHang) || String(maHang || "").trim();
    var mvNorm = normalizeProductCode(maVach) || String(maVach || "").trim();
    if (!mhNorm && !mvNorm) continue;
    var parentNorm = normalizeProductCode(parentSku) || String(parentSku || "").trim();

    var units = [];
    if (dvt) units.push(dvt);
    if (dvt2 && normalizeCatalogDvtPart_(dvt2) !== normalizeCatalogDvtPart_(dvt)) units.push(dvt2);
    if (!units.length) units.push("");

    for (var u = 0; u < units.length; u++) {
      var obj = {
        maHang: String(mhNorm || maHang || ""),
        maVach: String(mvNorm || maVach || ""),
        tenHang: tenHang,
        dvt: units[u] || "",
        dvt2: "",
        parentSku: String(parentNorm || parentSku || ""),
        isNew: isNewFlag,
        isLocked: false,
        isOutStock: false
      };
      if (!isNaN(tonNum)) obj.tonKho = tonNum;
      collected.push(obj);
    }
  }

  assignCatalogCompositeKeys_(collected);
  for (var c = 0; c < collected.length; c++) {
    var item = collected[c];
    if (!item || !item.key) continue;
    danhMucHangHoa[item.key] = item;
    var mvAlias = normalizeProductCode(item.maVach);
    if (mvAlias && !danhMucHangHoa[mvAlias]) {
      danhMucHangHoa[mvAlias] = item;
    }
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
  var lookup = { byKey: {}, byMaHang: {}, byMaVach: {}, byMaHangList: {} };
  var catalogResult = getCatalogData();
  if (!catalogResult || !catalogResult.success || !catalogResult.danhMuc) return lookup;

  var danhMuc = catalogResult.danhMuc;
  for (var key in danhMuc) {
    if (!Object.prototype.hasOwnProperty.call(danhMuc, key)) continue;
    var item = danhMuc[key];
    if (!item) continue;
    var itemKey = item.key || key;
    // Bỏ qua alias barcode thuần (trùng object đã index theo composite key)
    if (item.key && key !== item.key && key === (normalizeProductCode(item.maVach) || "")) {
      // alias MV — chỉ ghi byMaVach
    } else {
      lookup.byKey[itemKey] = item;
    }
    var mhKey = normalizeProductCode(item.maHang) || String(item.maHang || "").trim().toUpperCase();
    var mvKey = normalizeProductCode(item.maVach) || String(item.maVach || "").trim().toUpperCase();
    if (mhKey) {
      lookup.byMaHang[mhKey] = item; // last wins (compat)
      if (!lookup.byMaHangList[mhKey]) lookup.byMaHangList[mhKey] = [];
      var already = false;
      for (var li = 0; li < lookup.byMaHangList[mhKey].length; li++) {
        if (lookup.byMaHangList[mhKey][li] === item || lookup.byMaHangList[mhKey][li].key === item.key) {
          already = true; break;
        }
      }
      if (!already) lookup.byMaHangList[mhKey].push(item);
    }
    if (mvKey && !lookup.byMaVach[mvKey]) lookup.byMaVach[mvKey] = item;
  }
  return lookup;
}


function resolveCatalogProduct(lookup, maHang, maVach, dvt) {
  if (!lookup) return null;
  var mh = normalizeProductCode(maHang) || String(maHang || "").trim();
  var mv = normalizeProductCode(maVach) || String(maVach || "").trim();
  var dv = String(dvt || "").trim();

  // 1) Composite Key exact
  if (mh && dv && lookup.byKey) {
    var k1 = buildCatalogCompositeKey_(mh, dv, mv, false);
    var k2 = buildCatalogCompositeKey_(mh, dv, mv, true);
    if (lookup.byKey[k1]) return lookup.byKey[k1];
    if (lookup.byKey[k2]) return lookup.byKey[k2];
  }
  // 2) Cùng MaSP + ĐVT trong list
  if (mh && dv && lookup.byMaHangList && lookup.byMaHangList[mh]) {
    var list = lookup.byMaHangList[mh];
    for (var i = 0; i < list.length; i++) {
      if (normalizeCatalogDvtPart_(list[i].dvt) === normalizeCatalogDvtPart_(dv)) {
        if (!mv || !list[i].maVach || normalizeProductCode(list[i].maVach) === mv) return list[i];
      }
    }
  }
  // 3) MaVach exact (quy cách scan)
  if (mv && lookup.byMaVach[mv]) return lookup.byMaVach[mv];
  // 4) Compat: MaSP only (quy cách đầu / cuối)
  if (mh && lookup.byMaHang[mh]) return lookup.byMaHang[mh];
  return null;
}


function resolveDvtValue(lookup, maHang, maVach, currentDvt) {
  var cur = String(currentDvt || "").trim();
  var catalogItem = resolveCatalogProduct(lookup, maHang, maVach, cur);
  if (catalogItem && catalogItem.dvt) {
    // Nếu current khớp quy cách → giữ; không ép ĐVT1 của mã khác
    if (!cur || normalizeCatalogDvtPart_(cur) === normalizeCatalogDvtPart_(catalogItem.dvt)) {
      return String(catalogItem.dvt).trim();
    }
  }
  if (cur) return cur;
  return catalogItem && catalogItem.dvt ? String(catalogItem.dvt).trim() : "";
}


function getOrCreateCatalogSheet(ss) {
  var sheet = ss.getSheetByName("Data_Excel");
  var standardHeader = [
    "Mã hàng", "", "Mã vạch", "", "", "Tên hàng hóa", "", "ĐVT",
    "Ngày tạo", CATALOG_PARENT_HEADER, CATALOG_ISNEW_HEADER, CATALOG_DVT2_HEADER
  ];
  if (!sheet) {
    sheet = ss.insertSheet("Data_Excel");
    sheet.getRange(1, 1, 1, CATALOG_COL_COUNT).setValues([standardHeader]);
    sheet.getRange(1, 1, 1, CATALOG_COL_COUNT).setFontWeight("bold").setBackground("#d9ead3");
  } else {
    // Bổ sung header Ngày tạo / Parent_SKU / IsNew / DonViTinh2 nếu sheet cũ chưa có
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
      var h12 = String(sheet.getRange(1, 12).getValue() || "").trim();
      if (!h12) {
        sheet.getRange(1, 12).setValue(CATALOG_DVT2_HEADER).setFontWeight("bold").setBackground("#d9ead3");
      }
      // Chuẩn hóa nhãn ĐVT chính nếu ô trống / mơ hồ
      var hDvt = String(sheet.getRange(1, 8).getValue() || "").trim();
      if (!hDvt) {
        sheet.getRange(1, 8).setValue("ĐVT").setFontWeight("bold").setBackground("#d9ead3");
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
 * Xác định cột ghi Data_Excel theo TÊN header.
 * Thiếu cột → ghi vào vị trí chuẩn NẾU ô header trống/đúng vai trò; không thì APPEND cột mới.
 * Tuyệt đối không ghi đè cột đang mang dữ liệu vai trò khác (tránh lệch ĐVT/Tên/MV).
 */
function resolveCatalogWriteColMap_(sh, header, width) {
  header = header || [];
  width = Math.max(Number(width) || 0, header.length, CATALOG_COL_COUNT);
  var mapped = mapImportHeaderColumns_(header);
  var colMap = {
    maHang: mapped.maHang,
    maVach: mapped.maVach,
    tenHang: mapped.tenHang,
    dvt: mapped.dvt,
    dvt2: mapped.dvt2,
    parent: mapped.parentSku >= 0 ? mapped.parentSku : findCatalogParentColIdx_(header),
    ngay: findColumnIndexByAliases(header, ["ngaytao", "createdat", "created"]),
    isNew: findColumnIndexByAliases(header, ["isnew", "trangthaimoi", "hangmoi", "newflag"]),
    width: width
  };

  var CANON = { maHang: 0, maVach: 2, tenHang: 5, dvt: 7, ngay: 8, parent: 9, isNew: 10, dvt2: 11 };
  var LABELS = {
    maHang: "Mã hàng",
    maVach: "Mã vạch",
    tenHang: "Tên hàng hóa",
    dvt: "ĐVT",
    dvt2: CATALOG_DVT2_HEADER,
    ngay: "Ngày tạo",
    parent: CATALOG_PARENT_HEADER,
    isNew: CATALOG_ISNEW_HEADER
  };

  function roleScoreAt_(idx, roleKey) {
    if (idx < 0 || idx >= header.length) return 0;
    var n = normalizeHeaderText(header[idx]);
    if (!n) return 0;
    if (roleKey === "parent") return scoreImportHeaderRole_(n, "parentSku");
    if (roleKey === "ngay") {
      if (n.indexOf("ngaytao") !== -1 || n.indexOf("created") !== -1) return 90;
      return 0;
    }
    if (roleKey === "isNew") {
      if (n.indexOf("isnew") !== -1 || n === "moi" || n.indexOf("hangmoi") !== -1) return 90;
      return 0;
    }
    return scoreImportHeaderRole_(n, roleKey);
  }

  function ensureCol_(roleKey) {
    if (colMap[roleKey] >= 0 && roleScoreAt_(colMap[roleKey], roleKey) >= 70) return;
    if (colMap[roleKey] >= 0 && String(header[colMap[roleKey]] || "").trim() === "") {
      // giữ index đã map nhưng header trống — gắn nhãn
      try {
        sh.getRange(1, colMap[roleKey] + 1).setValue(LABELS[roleKey]).setFontWeight("bold").setBackground("#d9ead3");
      } catch (eL) {}
      header[colMap[roleKey]] = LABELS[roleKey];
      return;
    }
    var canon = CANON[roleKey];
    var hAtCanon = String(header[canon] == null ? "" : header[canon]).trim();
    if (!hAtCanon || roleScoreAt_(canon, roleKey) >= 70) {
      colMap[roleKey] = canon;
      if (!hAtCanon) {
        try {
          sh.getRange(1, canon + 1).setValue(LABELS[roleKey]).setFontWeight("bold").setBackground("#d9ead3");
        } catch (eC) {}
        while (header.length <= canon) header.push("");
        header[canon] = LABELS[roleKey];
      }
      colMap.width = Math.max(colMap.width, canon + 1);
      return;
    }
    // Cột chuẩn đang bị chiếm bởi dữ liệu khác → append cột mới
    var newIdx = colMap.width;
    colMap[roleKey] = newIdx;
    colMap.width = newIdx + 1;
    while (header.length <= newIdx) header.push("");
    header[newIdx] = LABELS[roleKey];
    try {
      sh.getRange(1, newIdx + 1).setValue(LABELS[roleKey]).setFontWeight("bold").setBackground("#d9ead3");
    } catch (eA) {}
  }

  ensureCol_("maHang");
  ensureCol_("maVach");
  ensureCol_("tenHang");
  ensureCol_("dvt");
  ensureCol_("dvt2");
  ensureCol_("ngay");
  ensureCol_("parent");
  ensureCol_("isNew");
  return colMap;
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
    sh.getRange(1, 1, 1, CATALOG_COL_COUNT).setValues([[
      "Mã hàng", "", "Mã vạch", "", "", "Tên hàng hóa", "", "ĐVT",
      "Ngày tạo", CATALOG_PARENT_HEADER, CATALOG_ISNEW_HEADER, CATALOG_DVT2_HEADER
    ]]);
    sh.getRange(1, 1, 1, CATALOG_COL_COUNT).setFontWeight("bold").setBackground("#d9ead3");
    lastRow = 1;
    width = CATALOG_COL_COUNT;
  }
  var values = sh.getRange(1, 1, lastRow, width).getValues();
  var header = values[0] || [];
  // Map cột theo TÊN header — không fallback hardcode vào cột sai
  var writeCols = resolveCatalogWriteColMap_(sh, header, width);
  var mhIdx = writeCols.maHang;
  var mvIdx = writeCols.maVach;
  var thIdx = writeCols.tenHang;
  var dvtIdx = writeCols.dvt;
  var dvt2Idx = writeCols.dvt2;
  var ngayIdx = writeCols.ngay;
  var parentIdx = writeCols.parent;
  var isNewIdx = writeCols.isNew;
  width = writeCols.width;
  values[0] = header;

  // Index sheet rows by Composite Key (MaSP_ĐVT[_MaVach])
  var existingItems = [];
  for (var r = 1; r < values.length; r++) {
    var mh0 = String(values[r][mhIdx] == null ? "" : values[r][mhIdx]).trim();
    var mv0 = String(values[r][mvIdx] == null ? "" : values[r][mvIdx]).trim();
    var dvt0 = sanitizeImportDvt_(values[r][dvtIdx]);
    if (!mh0 && !mv0) continue;
    existingItems.push({
      maHang: mh0,
      maVach: mv0,
      dvt: dvt0,
      _sheetRow: r + 1,
      _memIdx: r
    });
  }
  assignCatalogCompositeKeys_(existingItems);
  var byKey = {};
  for (var ei = 0; ei < existingItems.length; ei++) {
    var ex = existingItems[ei];
    if (ex && ex.key && byKey[ex.key] === undefined) byKey[ex.key] = ex;
  }

  var stamp = catalogNowStamp_();
  var updated = 0;
  var appended = 0;
  var preservedParent = 0;
  var withDvt = 0;
  var withDvt2 = 0;
  var skippedBadDvt = 0;
  var withParent = 0;
  var appendRows = [];

  var upsertList = expandCatalogEntriesByDvt_(entries || []);
  assignCatalogCompositeKeys_(upsertList);
  var tonSyncEntries = [];

  for (var i = 0; i < upsertList.length; i++) {
    var e = upsertList[i];
    if (!e) continue;
    var mh = String(e.mh || e.maHang || "").trim();
    var mv = String(e.mv || e.maVach || "").trim();
    var th = String(e.th || e.tenHang || "").trim();
    var dRaw = String(e.d || e.dvt || "").trim();
    var d = sanitizeImportDvt_(dRaw);
    var pIn = String(e.p || e.parentSku || "").trim();
    if (!mh && !mv) continue;
    if (dRaw && !d) skippedBadDvt++;
    if (d) withDvt++;

    var rowKey = e.key || buildCatalogCompositeKey_(mh, d, mv, false);
    var hit = byKey[rowKey];
    if (!hit) {
      var alt1 = buildCatalogCompositeKey_(mh, d, mv, true);
      var alt2 = buildCatalogCompositeKey_(mh, d, mv, false);
      hit = byKey[alt1] || byKey[alt2] || null;
    }
    var hasTon = e.ton !== undefined && e.ton !== null && e.ton !== "";
    var tonNum = hasTon ? Number(e.ton) : NaN;

    if (hit && hit._sheetRow > 1) {
      var sheetRow = hit._sheetRow;
      var memIdx = hit._memIdx;
      var row = values[memIdx] ? values[memIdx].slice() : [];
      while (row.length < width) row.push("");
      var oldParent = String(row[parentIdx] == null ? "" : row[parentIdx]).trim();
      if (mh) row[mhIdx] = mh;
      if (mv) row[mvIdx] = String(mv);
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
      sh.getRange(sheetRow, 1, 1, width).setValues([row]);
      byKey[rowKey] = { key: rowKey, maHang: mh, maVach: mv, dvt: d, _sheetRow: sheetRow, _memIdx: memIdx };
      updated++;
      if (hasTon && !isNaN(tonNum) && tonNum >= 0 && mh) {
        tonSyncEntries.push({ maHang: mh, mv: mv, th: th, d: d, p: pIn || oldParent || "", q: tonNum });
      }
    } else {
      var newRow = [];
      for (var c = 0; c < width; c++) newRow.push("");
      newRow[mhIdx] = mh;
      newRow[mvIdx] = String(mv);
      newRow[thIdx] = th;
      newRow[dvtIdx] = d;
      newRow[ngayIdx] = stamp.date;
      if (pIn) {
        newRow[parentIdx] = pIn;
        withParent++;
      }
      appendRows.push(newRow);
      appended++;
      byKey[rowKey] = { key: rowKey, maHang: mh, maVach: mv, dvt: d, _sheetRow: -1, _pending: true };
      if (hasTon && !isNaN(tonNum) && tonNum >= 0 && mh) {
        tonSyncEntries.push({ maHang: mh, mv: mv, th: th, d: d, p: pIn || "", q: tonNum });
      }
    }
  }

  if (appendRows.length) {
    var start = Math.max(sh.getLastRow() + 1, 2);
    sh.getRange(start, 1, appendRows.length, width).setValues(appendRows);
    try {
      sh.getRange(start, ngayIdx + 1, appendRows.length, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
    } catch (eFmt) {}
  }
  // Chỉ ghi tồn khi có số thực — KHÔNG gán đè = 0
  if (tonSyncEntries.length) {
    try { writeTonVariantEntriesToSheet_(ss, tonSyncEntries); } catch (eTonBatch) {}
  }
  try { SpreadsheetApp.flush(); } catch (e) {}

  return {
    rows: upsertList.length,
    updated: updated,
    appended: appended,
    preservedParent: preservedParent,
    withDvt: withDvt,
    withDvt2: withDvt2,
    skippedBadDvt: skippedBadDvt,
    withParent: withParent,
    tonSynced: tonSyncEntries.length,
    totalRows: Math.max(sh.getLastRow() - 1, 0),
    ms: Date.now() - t0,
    ngayTaoStamp: stamp.text,
    mode: "upsert-composite-key"
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
    var collected = [];
    var seenObj = {};
    for (var mapKey in danhMuc) {
      if (!Object.prototype.hasOwnProperty.call(danhMuc, mapKey)) continue;
      var existingItem = danhMuc[mapKey];
      if (!existingItem || !existingItem.key) continue;
      if (seenObj[existingItem.key]) continue;
      seenObj[existingItem.key] = true;
      collected.push(existingItem);
    }
    for (var k in byKey) {
      if (!Object.prototype.hasOwnProperty.call(byKey, k)) continue;
      var row = byKey[k];
      if (!row) continue;
      var mh = String(row.k || "").split("|")[0].trim();
      var mv = String(row.mv || "").trim();
      var parentSku = String(row.p || "").trim();
      var dvt = String(row.d || "").trim();
      if (!mh && !mv) continue;
      var ton = row.tonHienTai != null ? Number(row.tonHienTai) : NaN;

      // Tìm quy cách khớp MaSP+ĐVT(+MV) đã có
      var matched = null;
      for (var dk in danhMuc) {
        if (!Object.prototype.hasOwnProperty.call(danhMuc, dk)) continue;
        var cur = danhMuc[dk];
        if (!cur || !cur.key) continue;
        if (normalizeProductCode(cur.maHang) !== normalizeProductCode(mh) &&
            String(cur.maHang || "").trim().toUpperCase() !== String(mh || "").trim().toUpperCase()) continue;
        if (dvt && normalizeCatalogDvtPart_(cur.dvt) !== normalizeCatalogDvtPart_(dvt)) continue;
        if (mv && cur.maVach && normalizeProductCode(cur.maVach) !== normalizeProductCode(mv)) continue;
        matched = cur;
        break;
      }
      if (matched) {
        if (!matched.parentSku && parentSku) matched.parentSku = parentSku;
        if (!matched.tenHang && row.th) matched.tenHang = row.th;
        if (!matched.dvt && dvt) matched.dvt = dvt;
        if (!matched.maVach && mv) matched.maVach = mv;
        if (!isNaN(ton)) matched.tonKho = ton;
        if (row.isLocked) matched.isLocked = true;
        if (row.isOutStock) {
          matched.isOutStock = true;
          matched.isNew = false;
        }
        continue;
      }
      var obj = {
        maHang: mh,
        maVach: mv,
        tenHang: row.th || mh,
        dvt: dvt || "",
        dvt2: "",
        parentSku: parentSku,
        isNew: false,
        isLocked: !!row.isLocked,
        isOutStock: !!row.isOutStock,
        fromTonVariant: true
      };
      if (!isNaN(ton)) obj.tonKho = ton;
      collected.push(obj);
      added++;
    }
    assignCatalogCompositeKeys_(collected);
    // Rebuild map theo composite key (giữ alias barcode nếu trống)
    var rebuilt = {};
    for (var ci = 0; ci < collected.length; ci++) {
      var it = collected[ci];
      if (!it || !it.key) continue;
      rebuilt[it.key] = it;
      var mvA = normalizeProductCode(it.maVach);
      if (mvA && !rebuilt[mvA]) rebuilt[mvA] = it;
    }
    // Giữ các item cũ chưa nằm trong collected (alias-only keys đã có .key)
    for (var oldK in danhMuc) {
      if (!Object.prototype.hasOwnProperty.call(danhMuc, oldK)) continue;
      var oldIt = danhMuc[oldK];
      if (!oldIt || !oldIt.key) continue;
      if (!rebuilt[oldIt.key]) rebuilt[oldIt.key] = oldIt;
    }
    return { danhMuc: rebuilt, added: added };
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
  // Score-based map — không dùng alias "code" (khớp nhầm barcode)
  var stockMap = mapImportHeaderColumns_(header);
  var maHangIdx = stockMap.maHang;
  var maVachIdx = stockMap.maVach;
  var tonKhoIdx = stockMap.tonKho;
  var tenHangIdx = stockMap.tenHang;
  var dvtIdx = stockMap.dvt;
  var claimedIndexes = [maHangIdx, maVachIdx, tonKhoIdx, tenHangIdx, dvtIdx, stockMap.dvt2];

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
    // Không fallback hardcode 1/2/6 khi không nhận diện được — tránh lệch cột
    maHangIdx: maHangIdx,
    maVachIdx: maVachIdx,
    tonKhoIdx: tonKhoIdx,
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
  var headerIndex = findImportHeaderRowIndex_(importData, 15);
  if (headerIndex < 0) headerIndex = findHeaderRowIndex(importData, 10);
  if (headerIndex < 0) throw new Error("Không tìm thấy dòng tiêu đề trong sheet nguồn nhập khẩu.");
  var header = importData[headerIndex];
  var cols = mapImportHeaderColumns_(header);
  var idxMaHang = cols.maHang;
  var idxMaVach = cols.maVach;
  var idxTenHang = cols.tenHang;
  var idxDvt = cols.dvt;
  var idxTonKho = cols.tonKho;
  var idxKho = findColumnIndexByAliases(header, ['kho', 'cuahang', 'chinhanh', 'store', 'tenkho']);

  if (idxMaHang === -1 && idxMaVach === -1) {
    throw new Error("Sheet nguồn thiếu cột Mã hàng hoặc Mã vạch (không nhận diện được theo tiêu đề).");
  }

  var parsed = [];
  for (var r = headerIndex + 1; r < importData.length; r++) {
    var row = importData[r];
    if (!row) continue;
    if (isImportJunkDataRow_(row, cols)) continue;
    var maHang = idxMaHang >= 0 ? getCellValue(row, idxMaHang, "") : "";
    var maVach = idxMaVach >= 0 ? getCellValue(row, idxMaVach, "") : "";
    var tenHang = idxTenHang >= 0 ? getCellValue(row, idxTenHang, "") : "";
    var dvt = sanitizeImportDvt_(idxDvt >= 0 ? getCellValue(row, idxDvt, "") : "");
    var kho = idxKho >= 0 ? getCellValue(row, idxKho, "") : "";
    var tonRaw = idxTonKho !== -1 ? row[idxTonKho] : "";
    var tonKho = (tonRaw === "" || tonRaw === null || tonRaw === undefined) ? "" : Number(tonRaw);

    if (!maHang && !maVach) continue;
    if (!normalizeProductCode(maHang) && !normalizeProductCode(maVach)) continue;

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
    var dedupeInfo = { variantRemoved: 0, q7Removed: 0 };
    try {
      dedupeInfo = removeDuplicateStockRows_(ss) || dedupeInfo;
    } catch (eDed) {
      Logger.log(eDed);
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
      dedupeVariantRemoved: dedupeInfo.variantRemoved || 0,
      dedupeQ7Removed: dedupeInfo.q7Removed || 0,
      done: true,
      _debugTotalMs: Date.now() - tQ7,
      _debugQ7Ms: q7Fast.ms || 0,
      _debugRun: "import-q7-variant-upsert-v4-key-canon",
      msg: "Đã cập nhật " + TON_Q7_SHEET_NAME + " (" + (q7Fast.rows || 0) + " dòng)" +
        ((variantImport.rows || 0) ? (" + UPSERT " + TON_VARIANT_SHEET_NAME + " (" + variantImport.rows + " biến thể, giữ Da_Xuat)") : "") +
        ((dedupeInfo.variantRemoved || dedupeInfo.q7Removed)
          ? (" — đã dọn " + ((dedupeInfo.variantRemoved || 0) + (dedupeInfo.q7Removed || 0)) + " dòng Key trùng")
          : "") +
        "."
    };
  }

  // Import riêng TON_VARIANT (Admin) — UPSERT, không clear
  if (importType === "stockVariant") {
    var tVar = Date.now();
    var vEntries = payload.variantEntries || payload.q7Entries || [];
    if (!vEntries.length) throw new Error("Không có dòng tồn biến thể để ghi.");
    var vInfo = writeTonVariantEntriesToSheet_(ss, vEntries);
    var dedupeVar = { variantRemoved: 0, q7Removed: 0 };
    try { dedupeVar = removeDuplicateStockRows_(ss) || dedupeVar; } catch (eDv) { Logger.log(eDv); }
    return {
      success: true,
      importType: importType,
      targetSheet: TON_VARIANT_SHEET_NAME,
      updatedRows: vInfo.rows || 0,
      variantRows: vInfo.rows || 0,
      dedupeVariantRemoved: dedupeVar.variantRemoved || 0,
      done: true,
      _debugTotalMs: Date.now() - tVar,
      _debugRun: "import-ton-variant-upsert-v4-key-canon",
      msg: "Đã UPSERT " + TON_VARIANT_SHEET_NAME + " (" + (vInfo.updated || 0) + " cập nhật, " +
        (vInfo.appended || 0) + " thêm mới) — giữ Da_Xuat / mã con hiện có" +
        ((dedupeVar.variantRemoved || 0) ? ("; đã dọn " + dedupeVar.variantRemoved + " dòng Key trùng") : "") +
        "."
    };
  }

  // Import nhanh catalog: staging MISA_IMPORT + UPSERT Data_Excel (giữ Parent_SKU / mã con)
  if (importType === 'catalogFast') {
    var tCat = Date.now();
    var catEntries = payload.catalogEntries || [];
    if (!catEntries.length) {
      throw new Error("Không có dòng catalog để ghi. Kiểm tra file có cột mã hàng / mã vạch / tên / ĐVT.");
    }
    // Sanitize ĐVT lần nữa phía server (chặn số lượng/ngày/mã SP nhảy cột)
    for (var si = 0; si < catEntries.length; si++) {
      if (!catEntries[si]) continue;
      catEntries[si].d = sanitizeImportDvt_(catEntries[si].d || catEntries[si].dvt || "");
      catEntries[si].d2 = sanitizeImportDvt_(catEntries[si].d2 || catEntries[si].dvt2 || "");
      catEntries[si].dvt = catEntries[si].d;
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
      // Legacy matrix → dynamic header map (KHÔNG hardcode cột 0/2/5/7)
      var catalogMatrix = normalizeImportedMatrix(fileData);
      var parsedLegacy = extractCatalogEntriesFromMatrix_(catalogMatrix);
      var legacyEntries = parsedLegacy.entries || [];
      if (!legacyEntries.length) {
        throw new Error("Không đọc được dòng catalog. Kiểm tra file có dòng tiêu đề: Mã hàng / Tên / ĐVT / Mã vạch.");
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
        withDvt: legInfo.withDvt || 0,
        skippedBadDvt: legInfo.skippedBadDvt || parsedLegacy.meta.skippedBadDvt || 0,
        headerMap: parsedLegacy.meta.labels || {},
        chunkIndex: chunkIndex,
        chunkTotal: chunkTotal,
        done: isLastChunk,
        _debugRun: "import-catalog-dynamic-header-v1",
        msg: isLastChunk
          ? ('UPSERT Data_Excel (map cột động theo header). ĐVT hợp lệ: ' + (legInfo.withDvt || 0) +
            '. Staging: ' + MISA_IMPORT_SHEET_NAME + '.')
          : ('Đã nhận chunk ' + (chunkIndex + 1) + '/' + chunkTotal + ' (UPSERT, dynamic header).')
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
    if (force) {
      try { cache.remove(CACHE_TON_VARIANT_KEY); } catch (eClr1) {}
      try { cache.remove(CACHE_TON_Q7_KEY); } catch (eClr2) {}
      try { cache.remove(cacheKey); } catch (eClr3) {}
    }
    // Cache chunked — hỗ trợ catalog > 90KB (tránh miss cache khiến đọc sheet mỗi lần)
    if (!force) {
      var cached = getCacheJson_(cache, cacheKey);
      if (cached && cached.success && cached.danhMuc) return cached;
    }

    var ssCat = getSS();
    var danhMuc = buildCatalogFromSheet_(ssCat);
    try {
      var merged = mergeTonVariantChildrenIntoCatalog_(ssCat, danhMuc);
      danhMuc = merged.danhMuc || danhMuc;
    } catch (eMerge) {}
    try {
      applyProductLocksToCatalog_(ssCat, danhMuc);
    } catch (eLock) {}
    try {
      applyProductOutOfStockToCatalog_(ssCat, danhMuc);
    } catch (eOut) {}
    var stockMerge = { source: "none", mapKeys: 0, applied: 0 };
    try {
      stockMerge = applyMasterStockToCatalog_(ssCat, danhMuc) || stockMerge;
      danhMuc = stockMerge.danhMuc || danhMuc;
    } catch (eTon) {}
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
      stockFromTonVariant: stockMerge.source === "TON_VARIANT",
      stockSource: stockMerge.source || "none"
    };
    try {
      putCacheJson_(cache, cacheKey, result, CACHE_TTL_SECONDS);
    } catch (e) {}
    return result;
  } catch (e) {
    return { success: false, error: e.message || String(e) };
  }
}


/**
 * Tra cứu tồn linh hoạt — cùng logic applyMasterStockToCatalog_:
 * getStockValueForItem + fallback bare / MH: / getVariantStockIfPresent_ / parentSku.
 * @returns {number|null} tồn >= 0 nếu tìm thấy key; null nếu không có trong map
 */
function resolveStockValueWithFallback_(map, maHang, maVach, dvt, parentSku) {
  if (!map || typeof map !== "object") return null;
  var mh = normalizeProductCode(maHang);
  var hasKey = false;
  if (mh) {
    hasKey = Object.prototype.hasOwnProperty.call(map, mh) ||
      Object.prototype.hasOwnProperty.call(map, "MH:" + mh) ||
      Object.prototype.hasOwnProperty.call(map, String(maHang || "").trim());
  }
  if (!hasKey) {
    var mv = normalizeProductCode(maVach);
    if (mv) {
      hasKey = Object.prototype.hasOwnProperty.call(map, "MV:" + mv);
    }
  }
  var looked = null;
  if (!hasKey) {
    var vTry = getVariantStockIfPresent_(map, maHang, maVach, dvt);
    if (vTry !== null && vTry !== undefined) looked = vTry;
  } else {
    looked = getStockValueForItem(map, maHang, maVach, dvt);
    if ((looked === 0 || looked == null) && mh && map[mh] != null) looked = Number(map[mh]) || 0;
    if ((looked === 0 || looked == null) && mh && map["MH:" + mh] != null) looked = Number(map["MH:" + mh]) || 0;
  }
  if (looked !== null && looked !== undefined && looked !== "") {
    var n = Number(looked);
    if (isNaN(n)) n = 0;
    return Math.max(0, n);
  }

  // Fallback Parent_SKU / maHangDisplay — O(1) key trực tiếp
  var parentRaw = String(parentSku || "").trim();
  if (!parentRaw) return null;
  var pNorm = normalizeProductCode(parentRaw) || parentRaw.toUpperCase();
  if (!pNorm) return null;
  if (mh && pNorm === mh) return null;

  var dvtNorm = normalizeDvtKey_(dvt);
  var pKeys = [pNorm, "MH:" + pNorm];
  if (dvtNorm) pKeys.push("MH:" + pNorm + "|DV:" + dvtNorm);
  if (parentRaw !== pNorm) {
    pKeys.push(parentRaw);
    pKeys.push("MH:" + parentRaw);
    if (dvtNorm) pKeys.push("MH:" + parentRaw + "|DV:" + dvtNorm);
  }

  for (var pi = 0; pi < pKeys.length; pi++) {
    var pk = pKeys[pi];
    if (!pk) continue;
    if (map[pk] !== undefined && map[pk] !== null && map[pk] !== "") {
      var pn = Number(parseTonStockNumber_(map[pk]));
      if (isNaN(pn)) pn = 0;
      // #region agent log
      try {
        console.log(JSON.stringify({
          sessionId: "f6b0dc",
          hypothesisId: "H-parent",
          location: "catalog_variant.gs:resolveStockValueWithFallback_",
          message: "parentSku fallback hit",
          data: { maHang: mh || "", parentRaw: parentRaw, hitKey: pk, stock: Math.max(0, pn) },
          timestamp: Date.now()
        }));
      } catch (_dbgParentHit) {}
      // #endregion
      return Math.max(0, pn);
    }
  }
  // #region agent log
  try {
    console.log(JSON.stringify({
      sessionId: "f6b0dc",
      hypothesisId: "H-parent",
      location: "catalog_variant.gs:resolveStockValueWithFallback_",
      message: "parentSku fallback miss",
      data: { maHang: mh || "", parentRaw: parentRaw, pNorm: pNorm, tried: pKeys.length },
      timestamp: Date.now()
    }));
  } catch (_dbgParentMiss) {}
  // #endregion
  return null;
}


/**
 * Gắn tồn kho vào catalog — ƯU TIÊN sheet TON_Q7, sau đó TON_VARIANT, giữ Data_Excel nếu đã có.
 * Gán đủ alias: TonKho / tonKho / stock / tonHienTai.
 */
function applyMasterStockToCatalog_(ss, danhMuc) {
  danhMuc = danhMuc || {};
  ss = ss || getSS();
  var source = "none";
  var map = null;
  var q7Bundle = null;
  try {
    q7Bundle = readTonKhoQ7Bundle_(ss);
    if (q7Bundle && q7Bundle.map && Object.keys(q7Bundle.map).length) {
      map = q7Bundle.map;
      source = q7Bundle.source || "TON_Q7";
    }
  } catch (eQ7) {}
  if (!map || !Object.keys(map).length) {
    try {
      map = readTonVariantMap_(ss) || {};
      if (Object.keys(map).length) source = "TON_VARIANT";
    } catch (eVar) {}
  }
  if (!map || !Object.keys(map).length) {
    Logger.log("CẢNH BÁO TỒN KHO: không đọc được TON_Q7 / TON_VARIANT — giữ tonKho từ Data_Excel (không gán 0 hàng loạt).");
    return { danhMuc: danhMuc, source: source, mapKeys: 0 };
  }

  var applied = 0;
  for (var key in danhMuc) {
    if (!Object.prototype.hasOwnProperty.call(danhMuc, key)) continue;
    var item = danhMuc[key];
    if (!item) continue;
    var stockVal = resolveStockValueWithFallback_(map, item.maHang, item.maVach, item.dvt);
    if (stockVal === null || stockVal === undefined) continue;
    item.TonKho = stockVal;
    item.tonKho = stockVal;
    item.stock = stockVal;
    item.tonHienTai = stockVal;
    applied++;
  }
  return { danhMuc: danhMuc, source: source, mapKeys: Object.keys(map).length, applied: applied };
}


/** @deprecated dùng applyMasterStockToCatalog_ (ưu tiên TON_Q7) */
function applyTonVariantStockToCatalog_(ss, danhMuc) {
  var res = applyMasterStockToCatalog_(ss, danhMuc);
  return res.danhMuc || danhMuc;
}


/**
 * Gắn isLocked từ TON_VARIANT vào từng sản phẩm catalog (theo maHang / maVach).
 */
function applyProductLocksToCatalog_(ss, danhMuc) {
  danhMuc = danhMuc || {};
  ss = ss || getSS();
  var byKey = readTonVariantByKeyMap_(ss);
  var lockByCode = {};
  for (var k in byKey) {
    if (!Object.prototype.hasOwnProperty.call(byKey, k)) continue;
    var row = byKey[k];
    if (!row || !row.isLocked) continue;
    var mh = canonicalizeTonVariantKey_(row.k) || String(row.k || "").trim().toUpperCase();
    var mv = normalizeProductCode(row.mv);
    if (mh) lockByCode[mh] = true;
    if (mv) {
      lockByCode[mv] = true;
      lockByCode["MV:" + mv] = true;
    }
  }
  for (var key in danhMuc) {
    if (!Object.prototype.hasOwnProperty.call(danhMuc, key)) continue;
    var item = danhMuc[key];
    if (!item) continue;
    var mhU = String(item.maHang || "").trim().toUpperCase();
    var mvU = normalizeProductCode(item.maVach);
    var locked = !!(lockByCode[mhU] || (mvU && (lockByCode[mvU] || lockByCode["MV:" + mvU])));
    item.isLocked = locked;
  }
  return danhMuc;
}


/**
 * Admin API: khóa / mở khóa đặt hàng trên TON_VARIANT (cột IsLocked).
 * payload: { actor, maHang, maVach?, tenHang?, dvt?, parentSku?, isLocked: true|false }
 * hoặc { actor, items: [{ maHang, isLocked, ... }] }
 */
function updateProductLockStatus_(payload) {
  try {
    var actor = payload && payload.actor ? String(payload.actor).trim() : "";
    requireAdmin(actor);
    var items = [];
    if (payload && payload.items && payload.items.length) {
      items = payload.items;
    } else if (payload) {
      items = [payload];
    }
    if (!items.length) return { success: false, error: "Thiếu danh sách sản phẩm." };

    var ss = getSS();
    var lock = LockService.getDocumentLock();
    lock.waitLock(20000);
    try {
      var byKey = readTonVariantByKeyMap_(ss);
      var dirty = {};
      var updated = 0;
      for (var i = 0; i < items.length; i++) {
        var ent = items[i];
        if (!ent) continue;
        var mh = String(ent.maHang || ent.mh || "").trim();
        var key = buildTonVariantKey_(mh, ent.dvt || ent.d || "");
        if (!key && ent.k) key = canonicalizeTonVariantKey_(ent.k);
        if (!key) continue;
        var prev = byKey[key] || {
          k: key, p: "", tonBanDau: 0, daXuat: 0, tonHienTai: 0, daNhanNhap: 0, th: "", mv: "", d: "", isLocked: false, isOutStock: false
        };
        prev.k = key;
        var p = String(ent.parentSku || ent.p || "").trim();
        var th = String(ent.tenHang || ent.th || "").trim();
        var mv = String(ent.maVach || ent.mv || "").trim();
        var d = String(ent.dvt || ent.d || "").trim();
        if (p) prev.p = p;
        if (th) prev.th = th;
        if (mv) prev.mv = mv;
        if (d) prev.d = d;
        if (!prev.th || !prev.mv || !prev.d || !prev.p) {
          try {
            var catLookup = getCatalogLookup(ss);
            var cat = resolveCatalogProduct(catLookup, mh || key, mv);
            if (cat) {
              if (!prev.th && cat.tenHang) prev.th = cat.tenHang;
              if (!prev.mv && cat.maVach) prev.mv = cat.maVach;
              if (!prev.d && cat.dvt) prev.d = cat.dvt;
              if (!prev.p && cat.parentSku) prev.p = cat.parentSku;
            }
          } catch (eCat) {}
        }
        prev.isLocked = ent.isLocked === true || parseTonVariantLocked_(ent.isLocked);
        if (prev.isOutStock === undefined) prev.isOutStock = false;
        prev.tonHienTai = calcTonHienTaiVariant_(prev.tonBanDau, prev.daXuat, prev.daNhanNhap);
        byKey[key] = prev;
        dirty[key] = prev;
        updated++;
      }
      if (!updated) return { success: false, error: "Không có mã hợp lệ để cập nhật." };
      persistTonVariantByKeyNoClear_(ss, dirty);
      invalidateCatalogCache_();
      return {
        success: true,
        updated: updated,
        msg: "Đã cập nhật khóa đặt hàng cho " + updated + " sản phẩm."
      };
    } finally {
      try { lock.releaseLock(); } catch (eL) {}
    }
  } catch (e) {
    return { success: false, error: e.message || String(e) };
  }
}


/**
 * Gắn isOutStock từ TON_VARIANT vào catalog; hết hàng → isNew = false trên object trả về.
 */
function applyProductOutOfStockToCatalog_(ss, danhMuc) {
  danhMuc = danhMuc || {};
  ss = ss || getSS();
  var byKey = readTonVariantByKeyMap_(ss);
  var outByCode = {};
  for (var k in byKey) {
    if (!Object.prototype.hasOwnProperty.call(byKey, k)) continue;
    var row = byKey[k];
    if (!row || !row.isOutStock) continue;
    var mh = canonicalizeTonVariantKey_(row.k) || String(row.k || "").trim().toUpperCase();
    var mv = normalizeProductCode(row.mv);
    if (mh) outByCode[mh] = true;
    if (mv) {
      outByCode[mv] = true;
      outByCode["MV:" + mv] = true;
    }
  }
  for (var key in danhMuc) {
    if (!Object.prototype.hasOwnProperty.call(danhMuc, key)) continue;
    var item = danhMuc[key];
    if (!item) continue;
    var mhU = String(item.maHang || "").trim().toUpperCase();
    var mvU = normalizeProductCode(item.maVach);
    var out = !!(outByCode[mhU] || (mvU && (outByCode[mvU] || outByCode["MV:" + mvU])));
    item.isOutStock = out;
    if (out) item.isNew = false;
  }
  return danhMuc;
}


/** Gỡ IsNew trên Data_Excel theo danh sách mã SP / mã vạch */
function clearCatalogIsNewForCodes_(ss, codes) {
  ss = ss || getSS();
  codes = codes || [];
  var want = {};
  for (var i = 0; i < codes.length; i++) {
    var c = String(codes[i] || "").trim().toUpperCase();
    if (c) want[c] = true;
  }
  if (!Object.keys(want).length) return { changed: 0 };
  var sh = ss.getSheetByName("Data_Excel");
  if (!sh || sh.getLastRow() < 2) return { changed: 0 };
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
  var changed = 0;
  for (var r = 1; r < values.length; r++) {
    var mh = String(values[r][mhIdx] == null ? "" : values[r][mhIdx]).trim().toUpperCase();
    var mv = String(values[r][mvIdx] == null ? "" : values[r][mvIdx]).trim().toUpperCase();
    if (!(want[mh] || want[mv])) continue;
    var prevVal = String(values[r][isNewIdx] == null ? "" : values[r][isNewIdx]).trim();
    if (!prevVal) continue;
    if (!/^(1|true|yes|x|moi|new)$/i.test(prevVal)) continue;
    values[r][isNewIdx] = "";
    changed++;
  }
  if (changed) {
    var colOut = [];
    for (var rr = 0; rr < values.length; rr++) {
      colOut.push([values[rr][isNewIdx] == null ? "" : values[rr][isNewIdx]]);
    }
    sh.getRange(1, isNewIdx + 1, colOut.length, 1).setValues(colOut);
    try { SpreadsheetApp.flush(); } catch (eF) {}
  }
  return { changed: changed };
}


/**
 * Upsert cờ IsOutStock trên TON_VARIANT.
 * outOfStock=true → đồng thời gỡ IsNew trên Data_Excel.
 */
function updateProductOutOfStockStatus_(payload, outOfStock) {
  try {
    var actor = payload && payload.actor ? String(payload.actor).trim() : "";
    requireAdmin(actor);
    var items = [];
    if (payload && payload.items && payload.items.length) items = payload.items;
    else if (payload) items = [payload];
    if (!items.length) return { success: false, error: "Thiếu danh sách sản phẩm." };

    var ss = getSS();
    var lock = LockService.getDocumentLock();
    lock.waitLock(20000);
    try {
      var byKey = readTonVariantByKeyMap_(ss);
      var dirty = {};
      var updated = 0;
      var clearCodes = [];
      for (var i = 0; i < items.length; i++) {
        var ent = items[i];
        if (!ent) continue;
        var mh = String(ent.maHang || ent.mh || ent.maSP || "").trim();
        var key = buildTonVariantKey_(mh, ent.dvt || ent.d || "");
        if (!key && ent.k) key = canonicalizeTonVariantKey_(ent.k);
        if (!key) continue;
        var prev = byKey[key] || {
          k: key, p: "", tonBanDau: 0, daXuat: 0, tonHienTai: 0, daNhanNhap: 0, th: "", mv: "", d: "", isLocked: false, isOutStock: false
        };
        prev.k = key;
        var p = String(ent.parentSku || ent.p || "").trim();
        var th = String(ent.tenHang || ent.th || "").trim();
        var mv = String(ent.maVach || ent.mv || "").trim();
        var d = String(ent.dvt || ent.d || "").trim();
        if (p) prev.p = p;
        if (th) prev.th = th;
        if (mv) prev.mv = mv;
        if (d) prev.d = d;
        if (!prev.th || !prev.mv || !prev.d || !prev.p) {
          try {
            var catLookup = getCatalogLookup(ss);
            var cat = resolveCatalogProduct(catLookup, mh || key, mv);
            if (cat) {
              if (!prev.th && cat.tenHang) prev.th = cat.tenHang;
              if (!prev.mv && cat.maVach) prev.mv = cat.maVach;
              if (!prev.d && cat.dvt) prev.d = cat.dvt;
              if (!prev.p && cat.parentSku) prev.p = cat.parentSku;
            }
          } catch (eCat) {}
        }
        if (prev.isLocked === undefined) prev.isLocked = false;
        prev.isOutStock = outOfStock === true;
        prev.tonHienTai = calcTonHienTaiVariant_(prev.tonBanDau, prev.daXuat, prev.daNhanNhap);
        byKey[key] = prev;
        dirty[key] = prev;
        updated++;
        if (outOfStock) {
          if (mh) clearCodes.push(mh);
          if (prev.mv) clearCodes.push(prev.mv);
          if (mv) clearCodes.push(mv);
        }
      }
      if (!updated) return { success: false, error: "Không có mã hợp lệ để cập nhật." };
      persistTonVariantByKeyNoClear_(ss, dirty);
      var clearedNew = { changed: 0 };
      if (outOfStock) {
        try { clearedNew = clearCatalogIsNewForCodes_(ss, clearCodes); } catch (eClear) {}
      }
      invalidateCatalogCache_();
      return {
        success: true,
        updated: updated,
        clearedIsNew: clearedNew.changed || 0,
        isOutStock: !!outOfStock,
        msg: outOfStock
          ? ("Đã báo hết hàng " + updated + " SP" + (clearedNew.changed ? ("; gỡ IsNew " + clearedNew.changed + " dòng") : "") + ".")
          : ("Đã gỡ báo hết hàng / bán lại " + updated + " sản phẩm.")
      };
    } finally {
      try { lock.releaseLock(); } catch (eL2) {}
    }
  } catch (e) {
    return { success: false, error: e.message || String(e) };
  }
}


function markOutOfStockBatch_(payload) {
  return updateProductOutOfStockStatus_(payload || {}, true);
}


function markInStockBatch_(payload) {
  return updateProductOutOfStockStatus_(payload || {}, false);
}


/**
 * One-time cleanup: quét Data_Excel + TON_VARIANT, xóa ĐVT bất thường
 * (số lượng / ngày / mã SP nhảy cột). Không tự suy ĐVT mới nếu không chắc.
 */
function fixWrongDVTOnSheet_(payload) {
  try {
    var actor = payload && payload.actor ? String(payload.actor).trim() : "";
    requireAdmin(actor);
    var ss = getSS();
    var clearedCatalog = 0;
    var clearedVariant = 0;
    var samples = [];

    // --- Data_Excel ---
    var sh = ss.getSheetByName("Data_Excel");
    if (sh && sh.getLastRow() >= 2) {
      var width = Math.max(sh.getLastColumn(), CATALOG_COL_COUNT);
      var values = sh.getRange(1, 1, sh.getLastRow(), width).getValues();
      var header = values[0] || [];
      var mapped = mapImportHeaderColumns_(header);
      var dvtIdx = mapped.dvt;
      if (dvtIdx < 0) dvtIdx = findColumnIndexByAliases(header, ["donvitinh", "dvtinh", "tendvt", "dvt"]);
      if (dvtIdx < 0) dvtIdx = 7;
      var dvt2Idx = mapped.dvt2;
      if (dvt2Idx < 0) dvt2Idx = findColumnIndexByAliases(header, ["donvitinh2", "dvt2", "donviquydoi"]);
      var mhIdx = mapped.maHang;
      if (mhIdx < 0) mhIdx = 0;
      var dirty = false;
      for (var r = 1; r < values.length; r++) {
        var dvtRaw = String(values[r][dvtIdx] == null ? "" : values[r][dvtIdx]).trim();
        if (dvtRaw && !isPlausibleDvtValue_(dvtRaw)) {
          if (samples.length < 12) {
            samples.push({
              sheet: "Data_Excel",
              row: r + 1,
              maHang: String(values[r][mhIdx] || ""),
              badDvt: dvtRaw.slice(0, 40)
            });
          }
          values[r][dvtIdx] = "";
          clearedCatalog++;
          dirty = true;
        }
        if (dvt2Idx >= 0) {
          var d2Raw = String(values[r][dvt2Idx] == null ? "" : values[r][dvt2Idx]).trim();
          if (d2Raw && !isPlausibleDvtValue_(d2Raw)) {
            values[r][dvt2Idx] = "";
            clearedCatalog++;
            dirty = true;
          }
        }
      }
      if (dirty) {
        sh.getRange(1, 1, values.length, width).setValues(values);
      }
    }

    // --- TON_VARIANT cột DonViTinh (I = index 8) ---
    var shV = ss.getSheetByName(TON_VARIANT_SHEET_NAME);
    if (shV && shV.getLastRow() >= 2) {
      ensureTonVariantSchema_(shV);
      var lastV = shV.getLastRow();
      var valsV = shV.getRange(2, 1, lastV - 1, TON_VARIANT_COL_COUNT).getValues();
      var dirtyV = false;
      for (var vr = 0; vr < valsV.length; vr++) {
        var dvtV = String(valsV[vr][8] == null ? "" : valsV[vr][8]).trim();
        if (dvtV && !isPlausibleDvtValue_(dvtV)) {
          if (samples.length < 12) {
            samples.push({
              sheet: TON_VARIANT_SHEET_NAME,
              row: vr + 2,
              maHang: String(valsV[vr][0] || ""),
              badDvt: dvtV.slice(0, 40)
            });
          }
          valsV[vr][8] = "";
          clearedVariant++;
          dirtyV = true;
        }
      }
      if (dirtyV) {
        shV.getRange(2, 1, valsV.length, TON_VARIANT_COL_COUNT).setValues(valsV);
        try { getScriptCache_().remove(CACHE_TON_VARIANT_KEY); } catch (eC) {}
      }
    }

    invalidateCatalogCache_();
    try { SpreadsheetApp.flush(); } catch (eF) {}
    var total = clearedCatalog + clearedVariant;
    return {
      success: true,
      clearedCatalog: clearedCatalog,
      clearedVariant: clearedVariant,
      cleared: total,
      samples: samples,
      msg: total
        ? ("Đã xóa " + total + " ô ĐVT bất thường (Data_Excel: " + clearedCatalog +
          ", TON_VARIANT: " + clearedVariant + "). Hãy Import lại file MISA để ghi ĐVT chuẩn.")
        : "Không phát hiện ĐVT bất thường trên sheet."
    };
  } catch (e) {
    return { success: false, error: e.message || String(e) };
  }
}


/** Danh sách mã đang hết hàng (Admin) */
function getOutOfStockList_(query, limit) {
  try {
    var lim = Number(limit);
    if (isNaN(lim) || lim <= 0) lim = 500;
    lim = Math.min(lim, 2000);
    var qRaw = String(query || "").trim().toLowerCase();
    var ss = getSS();
    var byKey = readTonVariantByKeyMap_(ss);
    var items = [];
    for (var k in byKey) {
      if (!Object.prototype.hasOwnProperty.call(byKey, k)) continue;
      var row = byKey[k];
      if (!row || !row.isOutStock) continue;
      var mh = canonicalizeTonVariantKey_(row.k) || String(row.k || "").trim();
      var mv = String(row.mv || "").trim();
      var th = String(row.th || "").trim();
      if (qRaw) {
        var hay = (mh + " " + mv + " " + th).toLowerCase();
        if (hay.indexOf(qRaw) === -1) continue;
      }
      items.push({
        maHang: mh,
        maVach: mv,
        tenHang: th,
        dvt: String(row.d || "").trim(),
        parentSku: String(row.p || "").trim(),
        isOutStock: true
      });
      if (items.length >= lim) break;
    }
    items.sort(function(a, b) {
      return String(a.maHang || "").localeCompare(String(b.maHang || ""), "vi");
    });
    return { success: true, items: items, count: items.length };
  } catch (e) {
    return { success: false, error: e.message || String(e), items: [] };
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

    var ssNp = getSS();
    var data = getAutoNewProductsList_(ssNp, lim);
    // Loại SP đang hết hàng khỏi danh sách Hàng Mới
    try {
      var outMap = readTonVariantByKeyMap_(ssNp);
      data = (data || []).filter(function(it) {
        if (!it) return false;
        var mh = String(it.maHang || "").trim().toUpperCase();
        var mv = normalizeProductCode(it.maVach);
        var row = (mh && outMap[mh]) || null;
        if (!row && mv) {
          for (var ok in outMap) {
            if (!Object.prototype.hasOwnProperty.call(outMap, ok)) continue;
            if (normalizeProductCode(outMap[ok].mv) === mv) { row = outMap[ok]; break; }
          }
        }
        return !(row && row.isOutStock);
      });
    } catch (eFilt) {}
    var result = {
      success: true,
      data: data,
      limit: lim,
      source: "Data_Excel",
      strategy: "admin_isNew_only_exclude_outstock"
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
        k: key, p: "", tonBanDau: 0, daXuat: 0, tonHienTai: 0, daNhanNhap: 0, th: "", mv: "", d: "", isLocked: false
      };
      prev.k = key;
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
      // Giữ IsLocked trừ khi Admin gửi rõ
      if (ent.isLocked !== undefined && ent.isLocked !== null && ent.isLocked !== "") {
        prev.isLocked = ent.isLocked === true || parseTonVariantLocked_(ent.isLocked);
      } else if (prev.isLocked === undefined) {
        prev.isLocked = false;
      }
      prev.tonHienTai = calcTonHienTaiVariant_(prev.tonBanDau, prev.daXuat, prev.daNhanNhap);
      byKey[key] = prev;
      dirty[key] = prev;
    }
    var written = persistTonVariantByKeyNoClear_(ss, dirty);
    try {
      removeDuplicateStockRows_(ss, { skipLock: true });
    } catch (eDed) { Logger.log(eDed); }
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
