// ============================================================
// utils_helpers.gs — Helpers: cache, ngay gio, store, matrix, auth
// ============================================================


function getScriptCache_() {
  return CacheService.getScriptCache();
}


function normalizeDvtKey_(dvt) {
  return normalizeHeaderText(String(dvt || "").trim());
}


/**
 * Chuẩn hóa Key sheet tồn (TON_Q7 / lookup): giữ prefix MH:/MV:, BỎ hậu tố |DV: / |ĐVT.
 * ĐVT lưu cột Dvt riêng — tránh sinh 2 dòng MH:X và MH:X|DV:cai.
 */
function canonicalizeStockSheetKey_(key) {
  var k = String(key || "").trim();
  if (!k) return "";
  var prefix = "";
  if (k.indexOf("MH:") === 0) {
    prefix = "MH:";
    k = k.substring(3);
  } else if (k.indexOf("MV:") === 0) {
    prefix = "MV:";
    k = k.substring(3);
  }
  var dvIdx = k.toUpperCase().indexOf("|DV:");
  if (dvIdx !== -1) k = k.substring(0, dvIdx);
  else {
    var pipe = k.indexOf("|");
    if (pipe !== -1) k = k.substring(0, pipe);
  }
  k = normalizeProductCode(k) || String(k || "").trim().toUpperCase();
  return k ? (prefix + k) : "";
}


/**
 * Key TON_VARIANT duy nhất = mã SP (Parent hoặc mã con). KHÔNG gắn |ĐVT.
 */
function canonicalizeTonVariantKey_(keyOrMaHang) {
  var k = String(keyOrMaHang || "").trim();
  if (!k) return "";
  if (k.indexOf("MH:") === 0) k = k.substring(3);
  if (k.indexOf("MV:") === 0) k = k.substring(3);
  var dvIdx = k.toUpperCase().indexOf("|DV:");
  if (dvIdx !== -1) k = k.substring(0, dvIdx);
  var pipe = k.indexOf("|");
  if (pipe !== -1) k = k.substring(0, pipe);
  return normalizeProductCode(k) || String(k || "").trim().toUpperCase();
}


function dvtFromStockKey_(key) {
  var text = String(key || "");
  var idx = text.toUpperCase().indexOf("|DV:");
  if (idx !== -1) return text.substring(idx + 4);
  // Legacy TON_VARIANT: CODE|cai
  var pipe = text.indexOf("|");
  if (pipe !== -1 && text.indexOf("MH:") !== 0 && text.indexOf("MV:") !== 0) {
    return text.substring(pipe + 1);
  }
  return "";
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
    var merged = {};
    // Ưu tiên STORE_MAP cố định (2× Q4 phân biệt Mới/Cũ); Guide chỉ bổ sung kho thiếu
    if (registry && registry.storeMap) {
      for (var key in registry.storeMap) {
        if (Object.prototype.hasOwnProperty.call(registry.storeMap, key)) {
          merged[key] = registry.storeMap[key];
        }
      }
    }
    for (var fixed in STORE_MAP) {
      if (Object.prototype.hasOwnProperty.call(STORE_MAP, fixed)) {
        merged[fixed] = STORE_MAP[fixed];
      }
    }
    return merged;
  } catch (e) {
    return STORE_MAP;
  }
}


/**
 * Tên ngắn trên UI: tách 2 kho Q4 thành "Q4 Mới" / "Q4 Cũ".
 * Cột nội bộ soạn hàng vẫn dùng formatShortStoreLabel → Q4_178 / Q4_275.
 */
function formatStoreDisplayLabel_(storeName) {
  var raw = String(storeName || "").trim();
  if (!raw) return "";
  var code = lookupStoreCodeDigits_(raw);
  if (!code) {
    var normProbe = normalizeHeaderText(raw);
    if (normProbe.indexOf("178") !== -1) code = "178";
    else if (normProbe.indexOf("275") !== -1) code = "275";
    else if ((normProbe.indexOf("q4") !== -1 || normProbe.indexOf("quan4") !== -1) && normProbe.indexOf("moi") !== -1) code = "178";
    else if ((normProbe.indexOf("q4") !== -1 || normProbe.indexOf("quan4") !== -1) && (normProbe.indexOf("cu") !== -1 || normProbe.indexOf("old") !== -1)) code = "275";
  }
  if (code === "178") return "Q4 Mới";
  if (code === "275") return "Q4 Cũ";

  var activeMap = getActiveStoreMap();
  if (activeMap[raw]) {
    var mapped = String(activeMap[raw]).trim();
    if (STORE_SHORT_CODES[raw] === "178") return "Q4 Mới";
    if (STORE_SHORT_CODES[raw] === "275") return "Q4 Cũ";
    return mapped;
  }
  if (STORE_SHORT_CODES[raw] === "178") return "Q4 Mới";
  if (STORE_SHORT_CODES[raw] === "275") return "Q4 Cũ";

  try {
    var registry = getStoreRegistry();
    var details = registry && registry.storeDetails ? registry.storeDetails : [];
    var rawNorm = normalizeHeaderText(raw);
    for (var i = 0; i < details.length; i++) {
      var d = details[i];
      if (!d) continue;
      if (raw === d.fullName || raw === d.shortName ||
          rawNorm === normalizeHeaderText(d.fullName) ||
          rawNorm === normalizeHeaderText(d.shortName)) {
        if (STORE_SHORT_CODES[d.fullName] === "178") return "Q4 Mới";
        if (STORE_SHORT_CODES[d.fullName] === "275") return "Q4 Cũ";
        return String(d.shortName || d.fullName || raw).trim();
      }
    }
  } catch (e2) {}
  return raw;
}


/** Địa chỉ kho từ sheet Hướng dẫn — phân biệt 2 điểm Q4 theo mã 178/275 */
function lookupStoreAddress_(storeName) {
  var raw = String(storeName || "").trim();
  if (!raw) return "";
  try {
    var registry = getStoreRegistry();
    var details = registry && registry.storeDetails ? registry.storeDetails : [];
    var rawNorm = normalizeHeaderText(raw);
    var codeHint = lookupStoreCodeDigits_(raw) || "";
    if (!codeHint) {
      if (rawNorm.indexOf("178") !== -1) codeHint = "178";
      else if (rawNorm.indexOf("275") !== -1) codeHint = "275";
      else if (rawNorm.indexOf("moi") !== -1 && (rawNorm.indexOf("q4") !== -1 || rawNorm.indexOf("quan4") !== -1)) codeHint = "178";
      else if ((rawNorm.indexOf("cu") !== -1 || rawNorm.indexOf("old") !== -1) && (rawNorm.indexOf("q4") !== -1 || rawNorm.indexOf("quan4") !== -1)) codeHint = "275";
    }

    var byCode = "";
    for (var i = 0; i < details.length; i++) {
      var d = details[i];
      if (!d) continue;
      var dFull = normalizeHeaderText(d.fullName || "");
      var dShort = normalizeHeaderText(d.shortName || "");
      var dCode = String(d.code || "").trim();
      var addr = String(d.address || "").trim();
      if (!addr) continue;
      if (
        raw === d.fullName || raw === d.shortName ||
        rawNorm === dFull || rawNorm === dShort
      ) {
        return addr;
      }
      if (codeHint && (dCode.indexOf(codeHint) !== -1 || STORE_SHORT_CODES[d.fullName] === codeHint)) {
        byCode = addr;
      }
    }
    if (byCode) return byCode;
  } catch (e) {}
  return "";
}


/**
 * Nhãn in / header bảng tổng hợp: ưu tiên "Q4 - [địa chỉ]".
 * Fallback: "Q4 Mới" / "Q4 Cũ".
 */
function formatStorePrintLabel_(storeName) {
  var address = lookupStoreAddress_(storeName);
  var shortName = formatStoreDisplayLabel_(storeName) || String(storeName || "").trim();
  if (!shortName) return "";
  if (address) {
    var base = shortName;
    if (shortName === "Q4 Mới" || shortName === "Q4 Cũ" || shortName === "Q4") base = "Q4";
    return base + " - " + address;
  }
  return shortName;
}


/**
 * Header cột kho trên bảng tổng hợp soạn hàng (Excel).
 * Không dùng Q4_178 — dùng địa chỉ hoặc Q4 Mới/Cũ.
 */
function formatPackingColumnLabel_(storeName) {
  return formatStorePrintLabel_(storeName) || formatStoreDisplayLabel_(storeName) || formatShortStoreLabel(storeName) || String(storeName || "");
}


function getRuntimeStores() {
  try {
    var registry = getStoreRegistry();
    if (registry && registry.stores && registry.stores.length) return registry.stores;
  } catch (e) {}
  return Object.keys(STORE_MAP);
}

function getSS() {
  return SpreadsheetApp.openById(SHEET_ID);
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
  // Giữ Đ/đ — MISA phân biệt HĐCXXX1007 ≠ HDCXXX1007 (không map Đ→D).
  // NFC: tránh lệch so sánh Tiếng Việt (composed vs decomposed).
  var text = String(value == null ? "" : value).trim();
  if (!text) return "";
  try { text = text.normalize("NFC"); } catch (eNfc) {}
  return text
    .replace(/[^a-zA-Z0-9Đđ]/g, "")
    .trim()
    .toUpperCase();
}


/** Alias chuẩn hóa mã chứng từ / hóa đơn MISA (giữ Đ/đ) */
function normalizeMisaDocumentCode_(value) {
  return normalizeOrderCodeText(value);
}


/** Khớp số HĐ / mã chứng từ — CHỈ exact sau chuẩn hóa (tránh Q4 khớp nhầm Q4-275) */
function invoiceKeysMatch_(left, right) {
  var a = normalizeMisaDocumentCode_(left);
  var b = normalizeMisaDocumentCode_(right);
  if (!a || !b) return false;
  return a === b;
}


/**
 * Phân loại dòng xuất bán: HANG (vật lý) vs DV (dịch vụ đi kèm).
 * Ưu tiên cột loaiDong; fallback heuristic theo mã/ĐVT/tên.
 */
function isXuatBanServiceLine_(maHang, tenHang, dvt, loaiDong) {
  var loai = String(loaiDong || "").trim().toUpperCase();
  if (loai === "DV" || loai === "DICHVU" || loai === "SERVICE" || loai === "DỊCH VỤ" || loai === "DICH VU") return true;
  if (loai === "HANG" || loai === "SP" || loai === "PRODUCT") return false;
  var mh = String(maHang || "").trim().toUpperCase();
  var th = String(tenHang || "").trim().toUpperCase();
  var dv = String(dvt || "").trim().toUpperCase();
  if (mh.indexOf("DV-") === 0 || mh.indexOf("DV_") === 0 || mh === "DV") return true;
  if (dv === "DV" || dv === "DICH VU" || dv === "DỊCH VỤ" || dv === "LAN" || dv === "LẦN") return true;
  if (th.indexOf("DICH VU") === 0 || th.indexOf("DỊCH VỤ") === 0 || th.indexOf("PHÍ ") === 0) return true;
  return false;
}


/** Khớp số phiếu — CHỈ exact sau chuẩn hóa (Q4-275 ≠ Q4) */
function orderKeysMatch_(left, right) {
  var a = normalizeOrderCodeText(left);
  var b = normalizeOrderCodeText(right);
  if (!a || !b) return false;
  return a === b;
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
  var norm = normalizeOrderCodeText(s);
  if (matchSet[s] || matchSet[s.toLowerCase()] || (norm && matchSet[norm])) return true;
  var list = matchSet._list || [];
  for (var i = 0; i < list.length; i++) {
    // Exact only — không substring (tránh Q4 khớp Q4-275)
    if (orderKeysMatch_(s, list[i])) return true;
  }
  return false;
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


function findColumnIndexByAliases(row, aliases, excludeIndexes) {
  if (!row) return -1;
  var i;
  var c;
  var normalized;
  var skip = excludeIndexes || {};
  // Pass 1: khớp exact header — tránh "barcode".indexOf("code") / "unitname".indexOf("name")
  for (c = 0; c < row.length; c++) {
    if (skip[c]) continue;
    normalized = normalizeHeaderText(row[c]);
    if (!normalized) continue;
    for (i = 0; i < aliases.length; i++) {
      if (normalized === aliases[i]) return c;
    }
  }
  // Pass 2: contains, bỏ alias ngắn dễ đụng (code/name/sku đơn lẻ đã thử exact ở trên)
  for (c = 0; c < row.length; c++) {
    if (skip[c]) continue;
    normalized = normalizeHeaderText(row[c]);
    if (!normalized) continue;
    for (i = 0; i < aliases.length; i++) {
      var alias = aliases[i];
      if (!alias || alias.length < 4) continue;
      if (normalized.indexOf(alias) !== -1) return c;
    }
  }
  return -1;
}


/** Header ĐVT 2 / quy đổi — không được dùng làm ĐVT chính khi import */
function isImportDvt2HeaderNorm_(norm) {
  if (!norm) return false;
  if (norm.indexOf("quydoi") !== -1) return true;
  if (norm.indexOf("dvt2") !== -1 || norm.indexOf("donvi2") !== -1) return true;
  if (norm.indexOf("dvtphu") !== -1 || norm.indexOf("donviphu") !== -1) return true;
  if (norm.indexOf("unit2") !== -1 || norm.indexOf("altunit") !== -1) return true;
  if (norm.indexOf("heso") !== -1 && (norm.indexOf("dvt") !== -1 || norm.indexOf("donvi") !== -1)) return true;
  return false;
}


/** Header ĐVT chính (ĐVT 1) */
function isImportDvt1HeaderNorm_(norm) {
  if (!norm || isImportDvt2HeaderNorm_(norm)) return false;
  if (norm === "dvt" || norm === "donvi" || norm === "donvitinh" || norm === "dvtinh") return true;
  if (norm === "tendvt" || norm === "dvtchinh" || norm === "donvichinh" || norm === "basicunit") return true;
  if (norm === "unit" || norm === "uom" || norm === "unitname") return true;
  if (norm.indexOf("donvitinh") !== -1) return true;
  if (norm.indexOf("dvtchinh") !== -1 || norm.indexOf("tendvt") !== -1) return true;
  if (norm.indexOf("basicunit") !== -1) return true;
  if (norm.indexOf("dvt") === 0 && norm.length <= 12) return true;
  return false;
}


/**
 * Điểm khớp header → vai trò cột (cao hơn = chắc hơn).
 * Tránh first-match: "barcode".indexOf("code"), "unitname".indexOf("name").
 */
function scoreImportHeaderRole_(norm, role) {
  if (!norm) return 0;
  if (role === "parentSku") {
    if (norm === "parentsku" || norm === "parent" || norm === "manhomban") return 100;
    if ((norm.indexOf("mahang") !== -1 || norm.indexOf("mahanghoa") !== -1) &&
        (norm.indexOf("cha") !== -1 || norm.indexOf("parent") !== -1)) return 95;
    if (norm.indexOf("parentsku") !== -1 || norm.indexOf("manhomban") !== -1) return 90;
    return 0;
  }
  if (role === "dvt2") {
    if (isImportDvt2HeaderNorm_(norm)) {
      if (norm.indexOf("quydoi") !== -1) return 100;
      if (norm.indexOf("dvt2") !== -1 || norm.indexOf("donvi2") !== -1) return 95;
      return 85;
    }
    return 0;
  }
  if (role === "dvt") {
    if (isImportDvt2HeaderNorm_(norm)) return 0;
    if (norm.indexOf("price") !== -1 || norm.indexOf("gia") !== -1) return 0;
    if (norm.indexOf("madvt") !== -1 || norm === "madvt") return 0; // mã ĐVT ≠ tên ĐVT
    if (norm === "donvitinh" || norm === "dvtinh") return 100;
    if (norm === "tendvt" || norm === "dvtchinh" || norm === "donvichinh") return 98;
    if (norm === "dvt" || norm === "donvi") return 96;
    if (norm === "basicunit" || norm === "unitname" || norm === "uom" || norm === "unit") return 88;
    if (norm.indexOf("donvitinh") !== -1) return 94;
    if (norm.indexOf("tendvt") !== -1 || norm.indexOf("dvtchinh") !== -1) return 92;
    if (norm.indexOf("dvt") === 0 && norm.length <= 12) return 80;
    return 0;
  }
  if (role === "maHang") {
    if (norm.indexOf("cha") !== -1 || norm.indexOf("parent") !== -1) return 0;
    if (norm.indexOf("mavach") !== -1 || norm.indexOf("barcode") !== -1) return 0;
    if (norm === "mahanghoa" || norm === "masanpham") return 100;
    if (norm === "masp" || norm === "mahang" || norm === "mahh") return 96;
    if (norm === "itemcode" || norm === "article" || norm === "sku") return 90;
    if (norm.indexOf("mahanghoa") !== -1 || norm.indexOf("masanpham") !== -1) return 94;
    if (norm.indexOf("mahang") !== -1 && norm.indexOf("vach") === -1) return 88;
    // KHÔNG dùng alias "code" — khớp nhầm barcode
    return 0;
  }
  if (role === "maVach") {
    if (norm === "mavach" || norm === "barcode" || norm === "barcodeid" || norm === "ean") return 100;
    if (norm.indexOf("mavach") !== -1) return 95;
    if (norm.indexOf("barcode") !== -1) return 90;
    if (norm.indexOf("ean") !== -1 && norm.length <= 8) return 80;
    return 0;
  }
  if (role === "tenHang") {
    if (norm.indexOf("dvt") !== -1 || norm.indexOf("donvi") !== -1) return 0;
    if (norm.indexOf("unit") !== -1) return 0;
    if (norm === "tenhanghoa" || norm === "tensanpham") return 100;
    if (norm === "tensp" || norm === "tenhang") return 96;
    if (norm === "description") return 85;
    if (norm === "name") return 70; // exact only — không contains
    if (norm.indexOf("tenhanghoa") !== -1 || norm.indexOf("tensanpham") !== -1) return 94;
    if (norm.indexOf("tenhang") !== -1 || norm.indexOf("tensp") !== -1) return 90;
    return 0;
  }
  if (role === "tonKho") {
    if (norm === "tonkho" || norm === "soluongton" || norm === "slton" || norm === "cuoiky") return 100;
    if (norm === "soluong" || norm === "soton" || norm === "onhand" || norm === "stock") return 92;
    if (norm.indexOf("tonkho") !== -1 || norm.indexOf("soluongton") !== -1) return 95;
    if (norm.indexOf("soton") !== -1) return 90;
    if (norm === "qty" || norm === "quantity") return 75;
    if (norm.indexOf("soluong") !== -1 && norm.indexOf("quydoi") === -1) return 80;
    // Tránh Nhập kho / Xuất kho
    if (norm.indexOf("nhapkho") !== -1 || norm.indexOf("xuatkho") !== -1) return 0;
    return 0;
  }
  return 0;
}


/**
 * Map cột upload theo ĐIỂM SỐ tiêu đề (KHÔNG hardcode index, không first-match yếu).
 * @returns {{maHang:number,tenHang:number,dvt:number,dvt2:number,maVach:number,parentSku:number,tonKho:number,labels:Object,scores:Object}}
 */
function mapImportHeaderColumns_(headerRow) {
  var out = {
    maHang: -1,
    tenHang: -1,
    dvt: -1,
    dvt2: -1,
    maVach: -1,
    parentSku: -1,
    tonKho: -1,
    labels: {},
    scores: {}
  };
  if (!headerRow || !headerRow.length) return out;

  var roles = ["parentSku", "dvt2", "dvt", "maHang", "maVach", "tenHang", "tonKho"];
  var best = {};
  var r;
  for (r = 0; r < roles.length; r++) best[roles[r]] = { idx: -1, score: 0 };

  var c;
  for (c = 0; c < headerRow.length; c++) {
    var norm = normalizeHeaderText(headerRow[c]);
    if (!norm) continue;
    for (r = 0; r < roles.length; r++) {
      var role = roles[r];
      var sc = scoreImportHeaderRole_(norm, role);
      if (sc > best[role].score) {
        best[role] = { idx: c, score: sc };
      }
    }
  }

  // Gán không trùng cột: ưu tiên điểm cao hơn khi xung đột
  var taken = {};
  var orderByPriority = ["parentSku", "dvt2", "maVach", "maHang", "dvt", "tenHang", "tonKho"];
  // Sort roles by their best score desc for conflict resolution
  var assignOrder = orderByPriority.slice().sort(function(a, b) {
    return (best[b].score || 0) - (best[a].score || 0);
  });
  for (r = 0; r < assignOrder.length; r++) {
    var roleA = assignOrder[r];
    var cand = best[roleA];
    if (!cand || cand.idx < 0 || cand.score < 70) continue; // ngưỡng tối thiểu
    if (taken[cand.idx]) continue;
    out[roleA] = cand.idx;
    out.scores[roleA] = cand.score;
    taken[cand.idx] = roleA;
  }

  out.labels = {
    maHang: out.maHang >= 0 ? String(headerRow[out.maHang] || "") : "",
    tenHang: out.tenHang >= 0 ? String(headerRow[out.tenHang] || "") : "",
    dvt: out.dvt >= 0 ? String(headerRow[out.dvt] || "") : "",
    dvt2: out.dvt2 >= 0 ? String(headerRow[out.dvt2] || "") : "",
    maVach: out.maVach >= 0 ? String(headerRow[out.maVach] || "") : "",
    tonKho: out.tonKho >= 0 ? String(headerRow[out.tonKho] || "") : "",
    parentSku: out.parentSku >= 0 ? String(headerRow[out.parentSku] || "") : ""
  };
  return out;
}


/** Dòng rác / tiêu đề phụ / tổng cộng — bỏ khi import */
function isImportJunkDataRow_(row, cols) {
  if (!row) return true;
  var cells = [];
  var i;
  for (i = 0; i < row.length; i++) {
    var t = String(row[i] == null ? "" : row[i]).trim();
    if (t) cells.push(t);
  }
  if (!cells.length) return true;
  var joined = cells.join(" ").toLowerCase();
  var joinedNorm = normalizeHeaderText(joined);
  if (joinedNorm.indexOf("tongcong") !== -1 || joinedNorm === "tong" || joinedNorm.indexOf("total") === 0) return true;
  if (joinedNorm.indexOf("congty") !== -1 && cells.length <= 3) return true;
  if (joinedNorm.indexOf("baocao") !== -1 || joinedNorm.indexOf("phieukiem") !== -1) return true;
  if (joinedNorm.indexOf("ngaylap") !== -1 || joinedNorm.indexOf("trang") === 0) return true;
  // Dòng lặp lại header
  if (cols) {
    var lookLikeHeader = 0;
    if (cols.maHang >= 0 && scoreImportHeaderRole_(normalizeHeaderText(row[cols.maHang]), "maHang") >= 70) lookLikeHeader++;
    if (cols.dvt >= 0 && scoreImportHeaderRole_(normalizeHeaderText(row[cols.dvt]), "dvt") >= 70) lookLikeHeader++;
    if (cols.tenHang >= 0 && scoreImportHeaderRole_(normalizeHeaderText(row[cols.tenHang]), "tenHang") >= 70) lookLikeHeader++;
    if (lookLikeHeader >= 2) return true;
  }
  return false;
}


/**
 * ĐVT hợp lệ: cái/hộp/thùng… — loại số lượng, ngày, mã SP nhảy cột.
 * Giữ nguyên chữ Đ/đ trong giá trị (không fold).
 */
function isPlausibleDvtValue_(value) {
  var raw = String(value == null ? "" : value).trim();
  if (!raw) return false;
  if (raw.length > 40) return false;
  // Thuần số / số lượng
  if (/^[\d.,\s]+$/.test(raw)) return false;
  // Ngày tháng
  if (/^\d{1,4}[\/\-.]\d{1,2}[\/\-.]\d{1,4}/.test(raw)) return false;
  if (/^\d{4}-\d{2}-\d{2}/.test(raw)) return false;
  // Mã vạch dài / SKU số dài
  if (/^\d{8,}$/.test(raw.replace(/\s+/g, ""))) return false;
  // Scientific
  if (/^[+-]?\d+(\.\d+)?e[+-]?\d+$/i.test(raw)) return false;
  // Quá nhiều chữ số so với chữ → nghi là mã/số lượng
  var digits = (raw.match(/\d/g) || []).length;
  var letters = (raw.match(/[a-zA-ZÀ-ỹĐđ]/g) || []).length;
  if (digits >= 6 && digits > letters) return false;
  return true;
}


function sanitizeImportDvt_(value) {
  var raw = String(value == null ? "" : value).trim();
  if (!raw) return "";
  // Bỏ nháy Excel / NBSP — giữ UTF-8 / Đ/đ
  if (raw.charAt(0) === "'") raw = raw.slice(1).trim();
  raw = raw.replace(/\u00A0/g, " ").replace(/\s+/g, " ").trim();
  if (!isPlausibleDvtValue_(raw)) return "";
  return raw;
}


/** Tìm dòng header trong matrix upload (điểm theo score MaSP/MV/Tên/ĐVT) */
function findImportHeaderRowIndex_(rows, maxScan) {
  if (!rows || !rows.length) return 0;
  var limit = Math.min(rows.length, maxScan || 15);
  var bestScore = -1;
  var best = 0;
  for (var hi = 0; hi < limit; hi++) {
    var row = rows[hi] || [];
    var mapped = mapImportHeaderColumns_(row);
    var score = 0;
    if (mapped.maHang >= 0) score += (mapped.scores.maHang || 80);
    if (mapped.maVach >= 0) score += (mapped.scores.maVach || 70);
    if (mapped.tenHang >= 0) score += (mapped.scores.tenHang || 70);
    if (mapped.dvt >= 0) score += (mapped.scores.dvt || 80);
    if (mapped.dvt2 >= 0) score += 20;
    if (mapped.tonKho >= 0) score += 30;
    // Cộng nhẹ số ô có chữ (tiêu đề thật thường dày)
    var nonEmpty = 0;
    for (var c = 0; c < row.length; c++) {
      if (String(row[c] == null ? "" : row[c]).trim()) nonEmpty++;
    }
    score += Math.min(nonEmpty, 8);
    if (score > bestScore) { bestScore = score; best = hi; }
  }
  return best;
}


/**
 * Tách catalog entries từ matrix Excel/MISA bằng dynamic header map.
 * Output: [{mh,mv,th,d,d2,p}]
 */
function extractCatalogEntriesFromMatrix_(rows) {
  if (!rows || rows.length < 2) return { entries: [], meta: { reason: "empty" } };
  var headerIndex = findImportHeaderRowIndex_(rows, 15);
  var header = rows[headerIndex] || [];
  var cols = mapImportHeaderColumns_(header);
  if (cols.maHang < 0 && cols.maVach < 0) {
    return {
      entries: [],
      meta: {
        reason: "missing_ma_sp_or_barcode",
        headerIndex: headerIndex,
        headerSample: (header || []).slice(0, 16),
        labels: cols.labels,
        entryCount: 0
      }
    };
  }
  var entries = [];
  var withDvt = 0;
  var skippedBadDvt = 0;
  var withDvt2 = 0;
  var skippedJunk = 0;
  for (var r = headerIndex + 1; r < rows.length; r++) {
    var row = rows[r];
    if (!row) continue;
    if (isImportJunkDataRow_(row, cols)) { skippedJunk++; continue; }
    var mh = cols.maHang >= 0 ? getCellValue(row, cols.maHang, "") : "";
    var mv = cols.maVach >= 0 ? getCellValue(row, cols.maVach, "") : "";
    var th = cols.tenHang >= 0 ? getCellValue(row, cols.tenHang, "") : "";
    var dRaw = cols.dvt >= 0 ? getCellValue(row, cols.dvt, "") : "";
    var d2Raw = cols.dvt2 >= 0 ? getCellValue(row, cols.dvt2, "") : "";
    var p = cols.parentSku >= 0 ? getCellValue(row, cols.parentSku, "") : "";
    mh = String(mh || "").trim();
    mv = String(mv || "").trim();
    th = String(th || "").trim();
    p = String(p || "").trim();
    // Giữ nguyên Đ/đ / UTF-8 — chỉ trim, không fold
    if (!mh && !mv) continue;
    // Bỏ dòng chỉ có số tồn / không có mã thật
    if (!normalizeProductCode(mh) && !normalizeProductCode(mv)) continue;
    var d = sanitizeImportDvt_(dRaw);
    var d2 = sanitizeImportDvt_(d2Raw);
    if (dRaw && !d) skippedBadDvt++;
    if (d) withDvt++;
    if (d2) withDvt2++;
    entries.push({ mh: mh, mv: mv, th: th, d: d, d2: d2, p: p });
  }
  return {
    entries: entries,
    meta: {
      headerIndex: headerIndex,
      headerSample: (header || []).slice(0, 16),
      maHangIdx: cols.maHang,
      maVachIdx: cols.maVach,
      tenHangIdx: cols.tenHang,
      dvtIdx: cols.dvt,
      dvt2Idx: cols.dvt2,
      tonKhoIdx: cols.tonKho,
      parentIdx: cols.parentSku,
      labels: cols.labels,
      scores: cols.scores,
      entryCount: entries.length,
      withDvt: withDvt,
      withDvt2: withDvt2,
      skippedBadDvt: skippedBadDvt,
      skippedJunk: skippedJunk
    }
  };
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
  // Barcode/SKU dạng số từ Excel — không để scientific notation phá tìm kiếm mã vạch
  if (typeof value === "number" && isFinite(value)) {
    if (Math.floor(value) === value && Math.abs(value) < 1e16) {
      return String(Math.round(value));
    }
  }
  return String(value).trim();
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


function formatDateTime(value) {
  if (!value) return "";
  // Thống nhất với bảng soạn / PDF: dd/MM/yyyy HH:mm (không dùng toLocaleString lệch TZ)
  var label = formatOrderCreatedAtLabel_(value);
  if (label) return label;
  try {
    return String(value);
  } catch (e) {
    return "";
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
  // Số Excel thuần (barcode/SKU) — giữ dạng String, tránh scientific notation
  // Lưu ý: nếu Excel đã lưu barcode dạng Number thì số 0 đầu đã mất từ nguồn;
  // ô text / apostrophe vẫn giữ nguyên qua String().trim().
  if (typeof value === "number" && isFinite(value)) {
    if (Math.floor(value) === value && Math.abs(value) < 1e16) {
      value = String(Math.round(value));
    } else {
      value = String(value);
    }
  }
  var text = String(value).trim();
  if (!text) return "";
  try { text = text.normalize("NFC"); } catch (eNfc) {}

  if (text.charAt(0) === "'") text = text.slice(1);
  text = text.replace(/\u00A0/g, "").replace(/\s+/g, "");

  // Excel scientific notation: 8.93E+12 / 8.93e+12 — vẫn ra String digit, không Number() barcode thường
  if (/^[+-]?\d+(\.\d+)?e[+-]?\d+$/i.test(text)) {
    var sciNum = Number(text);
    if (isFinite(sciNum) && Math.abs(sciNum) < 1e16) {
      text = String(Math.round(sciNum));
    }
  }

  // Số Excel kiểu 12345.0 / 1.234.567 — xử lý trước khi lọc ký tự
  var upperProbe = text.toUpperCase();
  if (/^\d+\.0+$/.test(upperProbe)) {
    text = text.replace(/\.0+$/, "");
  }
  if (/^\d{1,3}(\.\d{3})+$/.test(text)) {
    text = text.replace(/\./g, "");
  }

  // BẮT BUỘC giữ Đ/đ — MISA: HĐCXXX1007 ≠ HDCXXX1007
  return text.replace(/[^a-zA-Z0-9Đđ]/g, "").trim().toUpperCase();
}


function normalizeNumericCode(value) {
  var code = normalizeProductCode(value);
  if (!code) return "";
  if (!/^\d+$/.test(code)) return "";
  // KHÔNG cắt số 0 đầu — barcode "0123" ≠ "123"
  return code;
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
  var adminActions = ['taoTaiKhoanMoi', 'nhapKhauCapNhatThongTin', 'removeDuplicateStockRows', 'fixWrongDVTOnSheet', 'saveCatalogIsNewFlags', 'saveChildVariants', 'updateProductLockStatus', 'markOutOfStockBatch', 'markInStockBatch'];
  if (adminActions.indexOf(action) !== -1 && !isAdminActor(payload && payload.actor ? payload.actor : "")) {
    throw new Error("Chỉ quản trị viên được phép thực hiện thao tác này.");
  }
}


function requireAdmin(actor) {
  var account = getAccountByActor(actor);
  if (!account || String(account.role).trim() !== "Admin") throw new Error("Chỉ quản trị viên được phép thay đổi hoặc hủy đơn.");
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
 * Nhãn thời gian tạo đơn thống nhất trên UI / bảng soạn.
 * Format: dd/MM/yyyy HH:mm (Asia/Ho_Chi_Minh).
 */
function formatOrderCreatedAtLabel_(valueOrMs) {
  var ms = toHoChiMinhMillis_(valueOrMs);
  if (isNaN(ms)) return "";
  var tz = Session.getScriptTimeZone() || "Asia/Ho_Chi_Minh";
  try {
    return Utilities.formatDate(new Date(ms), tz, "dd/MM/yyyy HH:mm");
  } catch (e) {
    return "";
  }
}


/**
 * Nhãn giờ tạo đơn cho header PDF / mẫu in.
 * Format: HH:mm - dd/MM/yyyy (vd: 10:15 - 05/08/2026).
 */
function formatOrderCreatedAtPretty_(valueOrMs) {
  var ms = toHoChiMinhMillis_(valueOrMs);
  if (isNaN(ms)) return "";
  var tz = Session.getScriptTimeZone() || "Asia/Ho_Chi_Minh";
  try {
    return Utilities.formatDate(new Date(ms), tz, "HH:mm - dd/MM/yyyy");
  } catch (e) {
    return "";
  }
}


/**
 * Từ các dòng lịch sử cùng số phiếu → mốc tạo đơn = thời gian sớm nhất (cột A).
 * Dùng chung cho bảng soạn hàng + click số phiếu / PDF — tránh lệch giờ dòng sửa sau.
 * @param {Array} rows mảng dòng sheet (mỗi row là array) hoặc {row: array}
 * @param {string=} soPhieuHint
 * @returns {{ms:number, label:string, soPhieu:string, khoXuat:string, khoNhan:string}|null}
 */
function extractOrderCreatedAtFromHistoryRows_(rows, soPhieuHint) {
  if (!rows || !rows.length) return null;
  var hintNorm = normalizeOrderCodeText(String(soPhieuHint || "").trim());
  var bestMs = NaN;
  var soPhieu = String(soPhieuHint || "").trim();
  var khoXuat = "";
  var khoNhan = "";
  for (var i = 0; i < rows.length; i++) {
    var raw = rows[i];
    var row = raw && raw.row ? raw.row : raw;
    if (!row) continue;
    var sp = row[1] != null ? String(row[1]).trim() : "";
    if (!sp) continue;
    if (hintNorm && normalizeOrderCodeText(sp) !== hintNorm) continue;
    var ms = toHoChiMinhMillis_(row[0]);
    if (isNaN(ms)) continue;
    if (isNaN(bestMs) || ms < bestMs) {
      bestMs = ms;
      soPhieu = sp;
      if (row[2]) khoXuat = String(row[2]).trim();
      if (row[3]) khoNhan = String(row[3]).trim();
    } else if (ms === bestMs) {
      if (!khoXuat && row[2]) khoXuat = String(row[2]).trim();
      if (!khoNhan && row[3]) khoNhan = String(row[3]).trim();
    }
  }
  if (isNaN(bestMs)) return null;
  return {
    ms: bestMs,
    label: formatOrderCreatedAtLabel_(bestMs),
    soPhieu: soPhieu,
    khoXuat: khoXuat,
    khoNhan: khoNhan
  };
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
