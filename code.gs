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
var CACHE_TTL_SECONDS = 1800;
var HISTORY_MAX_ROWS_DEFAULT = 8000;
// Kho soạn hàng chính — sheet nhẹ chỉ chứa tồn Q7 (tạo lúc import file tồn)
var PACKING_STOCK_STORE = "Kho Địa điểm kinh doanh Q7";
var TON_Q7_SHEET_NAME = "TON_Q7";
var CACHE_TON_Q7_KEY = "ton_q7_map_v2";

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
  var maHangIdx = findColumnIndexByAliases(headerRow, ['mahang', 'sku', 'article', 'code']);
  var maVachIdx = findColumnIndexByAliases(headerRow, ['mavach', 'barcode', 'barcodeid']);
  var tenHangIdx = findColumnIndexByAliases(headerRow, ['tenhang', 'name', 'tênhang', 'description']);
  var dvtIdx = findColumnIndexByAliases(headerRow, ['dvt', 'donvitinh', 'donvi', 'unit', 'uom']);
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
    if (dvt === "" && dvtIdx === -1) {
      var fallbackDvt = getCellValue(rawData[i], 6, "");
      if (fallbackDvt !== "") dvt = fallbackDvt;
    }
    if (tenHang === "" && maHang !== "") tenHang = tenHangChuanTheoMa[maHang.toUpperCase()] || "";
    var obj = { maHang: maHang, maVach: maVach, tenHang: tenHang, dvt: dvt || "" };
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
          case 'postProcessNewOrder':
            result = postProcessNewOrder(payload.payload || {});
            break;
          case 'postProcessPackingOrder':
            result = postProcessPackingOrder(payload.payload || {});
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
          res = getDanhSachDonSoanHang(e.parameter.ngay || '', e.parameter.userRole || '', e.parameter.userStore || '', e.parameter.ngayTo || '');
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
        data[i][6] || "",
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
  var ss = getSS();
  var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
  if (!historySheet || !soPhieu) return null;
  var selectedSet = buildOrderMatchSet_(soPhieu);
  var pack = readHistoryForSelectedOrders_(historySheet, selectedSet, "", 8000);
  var data = pack.data || [[]];
  for (var i = 1; i < data.length; i++) {
    if (!data[i] || !data[i][1]) continue;
    return {
      soPhieu: String(data[i][1]).trim(),
      khoXuat: data[i][2] ? String(data[i][2]).trim() : "",
      khoNhan: data[i][3] ? String(data[i][3]).trim() : ""
    };
  }
  return null;
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
  return String(value).trim().toLowerCase().normalize("NFKD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]+/g, "");
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
  var current = String(currentDvt || "").trim();
  var catalogItem = resolveCatalogProduct(lookup, maHang, maVach);
  var catalogDvt = catalogItem && catalogItem.dvt ? String(catalogItem.dvt).trim() : "";
  if (catalogDvt && (!current || normalizeHeaderText(current) === "cai")) return catalogDvt;
  return current || catalogDvt || "";
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
    sheet.getRange(1, 1, 1, 8).setValues([["Mã hàng", "", "Mã vạch", "", "", "Tên hàng hóa", "", "ĐVT"]]);
    sheet.getRange(1, 1, 1, 8).setFontWeight("bold").setBackground("#d9ead3");
  }
  return sheet;
}

/** Ghi catalog từ entries gọn {mh,mv,th,d} — layout Data_Excel cố định */
function writeCatalogEntriesToSheet_(ss, entries, reset) {
  ss = ss || getSS();
  var sh = getOrCreateCatalogSheet(ss);
  var t0 = Date.now();
  if (reset) {
    var oldLastRow = sh.getLastRow();
    var oldLastCol = Math.max(sh.getLastColumn(), 8);
    if (oldLastRow > 0) sh.getRange(1, 1, oldLastRow, oldLastCol).clearContent();
    sh.getRange(1, 1, 1, 8).setValues([["Mã hàng", "", "Mã vạch", "", "", "Tên hàng hóa", "", "ĐVT"]]);
    sh.getRange(1, 1, 1, 8).setFontWeight("bold").setBackground("#d9ead3");
  }
  var rows = [];
  var withDvt = 0;
  for (var i = 0; i < (entries || []).length; i++) {
    var e = entries[i];
    if (!e) continue;
    var mh = String(e.mh || "").trim();
    var mv = String(e.mv || "").trim();
    var th = String(e.th || "").trim();
    var d = String(e.d || "").trim();
    if (!mh && !mv) continue;
    if (d) withDvt++;
    rows.push([mh, "", mv, "", "", th, "", d]);
  }
  if (rows.length) {
    var startRow = Math.max(sh.getLastRow() + 1, 2);
    sh.getRange(startRow, 1, rows.length, 8).setValues(rows);
  }
  try { SpreadsheetApp.flush(); } catch (e) {}
  return {
    rows: rows.length,
    withDvt: withDvt,
    totalRows: Math.max(sh.getLastRow() - 1, 0),
    ms: Date.now() - t0
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
    return {
      success: true,
      importType: importType,
      targetSheet: TON_Q7_SHEET_NAME,
      updatedRows: q7Fast.rows || 0,
      updatedCols: 4,
      q7Sheet: TON_Q7_SHEET_NAME,
      q7Rows: q7Fast.rows || 0,
      q7WithDvt: withDvt,
      done: true,
      _debugTotalMs: Date.now() - tQ7,
      _debugQ7Ms: q7Fast.ms || 0,
      _debugRun: "import-q7-fast-v2",
      msg: "Đã cập nhật nhanh sheet " + TON_Q7_SHEET_NAME + " (" + (q7Fast.rows || 0) + " dòng, " + withDvt + " có ĐVT) — không ghi full TỔNG HỢP TỒN KHO."
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
  var catalogColumnCount = Math.max(catalogSheet.getLastColumn(), 8);
  if (catalogSheet.getLastRow() >= catalogStartRow) {
    catalogSheet.getRange(catalogStartRow, 1, catalogSheet.getLastRow() - catalogStartRow + 1, catalogColumnCount).clearContent();
  }
  if (catalogRows.length) {
    var catalogWrite = [];
    for (var c = 0; c < catalogRows.length; c++) {
      var rowOut = [];
      for (var z = 0; z < catalogColumnCount; z++) rowOut.push("");
      rowOut[cMaHang] = catalogRows[c].maHang;
      rowOut[cMaVach] = catalogRows[c].maVach;
      rowOut[cTenHang] = catalogRows[c].tenHang;
      rowOut[cDvt] = catalogRows[c].dvt;
      catalogWrite.push(rowOut);
    }
    catalogSheet.getRange(catalogStartRow, 1, catalogWrite.length, catalogColumnCount).setValues(catalogWrite);
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
    var registry = getStoreRegistry();
    return {
      success: true,
      stores: registry.stores,
      storeMap: registry.storeMap,
      storeDetails: registry.storeDetails || [],
      catalogVersion: getCatalogVersion_()
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getCatalogData() {
  try {
    var version = getCatalogVersion_();
    var cacheKey = CACHE_CATALOG_PREFIX + version;
    var cache = getScriptCache_();
    var cached = cache.get(cacheKey);
    if (cached) {
      try { return JSON.parse(cached); } catch (e) {}
    }

    var danhMuc = buildCatalogFromSheet_(getSS());
    var result = { success: true, danhMuc: danhMuc, version: version };
    try {
      var json = JSON.stringify(result);
      if (json.length < 90000) cache.put(cacheKey, json, CACHE_TTL_SECONDS);
    } catch (e) {}
    return result;
  } catch (e) {
    return { success: false, error: e.message };
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
      return { data: [], _debugTotalMs: Date.now() - t0, _debugRun: "ql-fast-v3", _debugRole: String(userRole || "") };
    }

    filterKhoNhan = normalizeStoreName(khoNhan || "");
    filterUserStore = normalizeStoreName(userStore || "");
    var bounds = getNgayFilterBounds_(ngayYYYYMMDD);
    boundsFilter = bounds.filter || "";
    var lastRow = historySheet.getLastRow();
    if (lastRow < 2) {
      return { data: [], _debugTotalMs: Date.now() - t0, _debugRun: "ql-fast-v3", _debugRole: String(userRole || "") };
    }

    function matchStoreCounted_(rowStore, targetStore) {
      storeCalls++;
      return storeMatchesFast_(rowStore, targetStore);
    }

    var lastCol = Math.min(Math.max(historySheet.getLastColumn(), 16), 16);
    var chunkSize = 500;
    var maxScan = Math.min(bounds.maxScan || 4000, lastRow - 1);
    var map = {};
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
        var rowDateStr = "";
        var rowMs = null;
        if (rowNgay instanceof Date && !isNaN(rowNgay.getTime())) {
          rowMs = rowNgay.getTime();
          rowDateStr = Utilities.formatDate(rowNgay, tz, "yyyy-MM-dd");
        } else {
          rowDateStr = formatSheetDateYYYYMMDD(rowNgay);
          if (rowDateStr) {
            var parsed = parseDateInputYYYYMMDD(rowDateStr);
            if (parsed) rowMs = parsed.getTime();
          }
        }

        if (bounds.exactDate) {
          if (rowDateStr !== bounds.exactDate) {
            if (rowMs != null && rowMs < bounds.startMs) { /* older */ }
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
        if (filterKhoNhan && filterKhoNhan !== "all" && !matchStoreCounted_(rowKhoNhan, filterKhoNhan)) continue;
        if (userRole !== "Admin") {
          if (!matchStoreCounted_(rowKhoXuat, filterUserStore) && !matchStoreCounted_(rowKhoNhan, filterUserStore)) continue;
        }

        var slThucTe = row[8];
        var slSoanCol = row[15];
        var rowStatus = row[12] ? String(row[12]).trim() : "Mới";
        var displayStatus = getDisplayOrderStatus(rowStatus, slThucTe, slSoanCol);
        var thoiGian = rowMs != null ? rowMs : "";

        if (!map[rowSoPhieu]) {
          map[rowSoPhieu] = { soPhieu: rowSoPhieu, khoXuat: rowKhoXuat, khoNhan: rowKhoNhan, thoiGian: thoiGian, trangThai: displayStatus === "Đã hủy dòng" ? "Mới" : displayStatus };
        } else {
          if (displayStatus === "Đã hủy") map[rowSoPhieu].trangThai = "Đã hủy";
          else if (displayStatus === "Đã xác nhận" && map[rowSoPhieu].trangThai !== "Đã hủy") map[rowSoPhieu].trangThai = "Đã xác nhận";
          else if (displayStatus === "Đã soạn" && map[rowSoPhieu].trangThai !== "Đã hủy" && map[rowSoPhieu].trangThai !== "Đã xác nhận") map[rowSoPhieu].trangThai = "Đã soạn";
          if (thoiGian && (!map[rowSoPhieu].thoiGian || thoiGian > map[rowSoPhieu].thoiGian)) map[rowSoPhieu].thoiGian = thoiGian;
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

    var res = [];
    for (var key in map) res.push(map[key]);
    res.sort(function(a, b) { return (b.thoiGian || 0) - (a.thoiGian || 0); });
    // Object wrapper giữ meta debug qua JSON (array custom props bị mất khi stringify).
    return {
      data: res,
      _debugTotalMs: Date.now() - t0,
      _debugScanned: scanned,
      _debugFilter: boundsFilter,
      _debugStoreCalls: storeCalls,
      _debugRole: String(userRole || ""),
      _debugKhoNhan: filterKhoNhan,
      _debugUserStore: filterUserStore,
      _debugRun: "ql-fast-v3"
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
      _debugRun: "ql-fast-v3-err"
    };
  }
}

function getDashboardSummary(userRole, userStore, timeline, fromDate, toDate) {
  try {
    var ss = getSS();
    var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
    if (!historySheet) {
      return { success: true, data: { totalOrders: 0, pendingOrders: 0, processedOrders: 0, canceledOrders: 0, recentOrders: [] } };
    }

    var data = readHistoryDataPack_(historySheet).data;
    var orderMap = {};
    for (var i = 1; i < data.length; i++) {
      var rowSoPhieu = data[i][1] ? data[i][1].toString().trim() : "";
      if (!rowSoPhieu) continue;
      if (!isDateInTimeline(data[i][0], timeline || '2days', fromDate || '', toDate || '')) continue;
      var rowKhoXuat = data[i][2] ? data[i][2].toString().trim() : "";
      var rowKhoNhan = data[i][3] ? data[i][3].toString().trim() : "";
      if (userRole !== "Admin") {
        if (rowKhoXuat !== userStore && rowKhoNhan !== userStore) continue;
      }

      var status = data[i][12] ? String(data[i][12]).trim() : "Mới";
      var slThucTe = data[i][8];
      var displayStatus = getDisplayOrderStatus(status, slThucTe, data[i][15]);
      var entry = orderMap[rowSoPhieu];
      if (!entry) {
        entry = { soPhieu: rowSoPhieu, khoXuat: rowKhoXuat, khoNhan: rowKhoNhan, thoiGian: data[i][0], status: displayStatus === "Đã hủy dòng" ? "Mới" : displayStatus, count: 0 };
        orderMap[rowSoPhieu] = entry;
      }

      if (displayStatus === "Đã hủy") entry.status = "Đã hủy";
      else if (displayStatus === "Đã xác nhận") entry.status = "Đã xác nhận";
      else if (displayStatus === "Đã soạn" && entry.status !== "Đã hủy" && entry.status !== "Đã xác nhận") entry.status = "Đã soạn";
      else if (entry.status !== "Đã hủy" && entry.status !== "Đã soạn" && entry.status !== "Đã xác nhận") entry.status = "Mới";

      entry.count += 1;
      if (!entry.thoiGian || (data[i][0] instanceof Date && entry.thoiGian instanceof Date && data[i][0].getTime() > entry.thoiGian.getTime())) {
        entry.thoiGian = data[i][0];
      }
    }

    var orders = [];
    for (var key in orderMap) orders.push(orderMap[key]);
    orders.sort(function(a, b) {
      var aTime = a.thoiGian instanceof Date ? a.thoiGian.getTime() : 0;
      var bTime = b.thoiGian instanceof Date ? b.thoiGian.getTime() : 0;
      return bTime - aTime;
    });

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

    var recentOrders = orders.slice(0, 8).map(function(order) {
      return {
        soPhieu: order.soPhieu,
        khoXuat: order.khoXuat,
        khoNhan: order.khoNhan,
        status: order.status,
        thoiGian: formatDateTime(order.thoiGian)
      };
    });

    return { success: true, data: { totalOrders: totalOrders, pendingOrders: pendingOrders, processedOrders: processedOrders, canceledOrders: canceledOrders, recentOrders: recentOrders } };
  } catch (e) {
    return { success: false, error: e.message };
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
  var ss = getSS();
  var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
  if (!historySheet || !soPhieu) return [];
  var wantStock = !(includeStock === false || includeStock === 0 || includeStock === "0" || includeStock === "false");
  var selectedSet = buildOrderMatchSet_(soPhieu);
  var pack = readHistoryForSelectedOrders_(historySheet, selectedSet, "", 6000);
  var data = pack.data || [[]];
  var sheetOrders = pack.orders || [];

  // Không load full catalog ở đây (rất chậm) — dùng ĐVT đã lưu trên lịch sử
  var matchedRows = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i]) continue;
    var slGoc = Number(data[i][7]) || 0;
    var hasActualQty = (data[i][8] !== "" && data[i][8] !== undefined && data[i][8] !== null);
    var rowStatus = data[i][12] ? String(data[i][12]).trim() : "Mới";
    var isReceived = rowStatus === "Đã xác nhận nhận hàng";
    // Cột 16 (index 15): SL Giao (Soạn) — giữ số đã soạn sau khi chi nhánh xác nhận nhận
    var rawSlGiao = data[i][15];
    var hasSlGiao = rawSlGiao !== "" && rawSlGiao !== null && rawSlGiao !== undefined;
    var slSoan = "";
    if (hasSlGiao) {
      slSoan = Number(rawSlGiao) || 0;
    } else if (hasActualQty && !isReceived) {
      // Dữ liệu cũ: cột 9 đang là số soạn
      slSoan = Number(data[i][8]) || 0;
    }
    // slThucTe: số thực nhận (sau xác nhận); trước đó để trống để UI không nhầm với SL soạn
    var slThucTe = "";
    if (isReceived && hasActualQty) slThucTe = Number(data[i][8]);
    else if (!isReceived && hasActualQty && !hasSlGiao) slThucTe = Number(data[i][8]); // tương thích cũ
    matchedRows.push({
      rowIndex: sheetOrders[i - 1] || (pack.startRow + i - 1),
      maHang: data[i][4],
      maVach: data[i][5],
      tenHang: data[i][6],
      slGoc: slGoc,
      slSoan: slSoan,
      slThucTe: slThucTe,
      dvt: data[i][9] || "",
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

  // Trả mảng thuần — tương thích FE cũ (rows.forEach / rows.filter)
  try {
    matchedRows._debugTotalMs = Date.now() - t0;
    matchedRows._debugScanned = pack.scannedRows || 0;
    matchedRows._debugStock = wantStock;
    matchedRows._debugRun = "ql-fast-v2";
  } catch (metaErr2) {}
  return matchedRows;
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
  if (status === "Đã xác nhận nhận hàng") return "Đã xác nhận";
  if (status === "Đã soạn hàng") return "Đã soạn";
  if (status === "Đã hủy dòng") return "Đã hủy dòng";
  // Có SL soạn (cột 16) hoặc SL cột 9 → coi như đã soạn (kể cả khi status sheet còn "Mới")
  if (slSoan !== "" && slSoan !== undefined && slSoan !== null) return "Đã soạn";
  if (slThucTe !== "" && slThucTe !== undefined && slThucTe !== null) return "Đã soạn";
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

function logOrderChange(ss, soPhieu, action, actor, maHang, maVach, oldValue, newValue, note) {
  getAuditSheet(ss).appendRow([new Date(), soPhieu, action, actor, maHang || "", maVach || "", oldValue || "", newValue || "", note || ""]);
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
  var ss = getSS();
  var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
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
    var permission = assertActorCanManageOrder(payload.actor, soPhieu, historySheet, historyData);
    var actorRole = String(permission.account.role || "").trim();
    var wasPacked = permission.state.isPacked;
    var orderBaseRow = null;
    for (var i = 1; i < historyData.length; i++) {
      if (String(historyData[i][1]).trim().toLowerCase() === soPhieu.toLowerCase()) {
        orderBaseRow = historyData[i];
        break;
      }
    }
    if (wasPacked && actorRole === "Admin") {
      for (var r = 1; r < historyData.length; r++) {
        if (String(historyData[r][1]).trim().toLowerCase() !== soPhieu.toLowerCase()) continue;
        var rowStatus = String(historyData[r][12] || "").trim();
        if (rowStatus === "Đã hủy đơn" || rowStatus === "Đã hủy dòng") continue;
        historySheet.getRange(r + 1, 9).clearContent();
        historySheet.getRange(r + 1, 12, 1, 4).setValues([["Cần soạn lại sau khi chỉnh sửa", "Mới", payload.actor || "", "Quản lý"]]);
      }
    }
    var newRows = [];
    if (payload.newItems && payload.newItems.length) {
      for (var n = 0; n < payload.newItems.length; n++) {
        var newItem = payload.newItems[n];
        if (hasDuplicateItemInOrder(historySheet, soPhieu, newItem, historyData)) {
          throw new Error("Mã này đã tồn tại trong đơn hiện tại. Không thể thêm dòng trùng.");
        }
        var itemQty = Number(newItem.sl);
        if (!itemQty || itemQty < 1) continue;
        newRows.push([new Date(), soPhieu, orderBaseRow && orderBaseRow[2] ? orderBaseRow[2] : "", orderBaseRow && orderBaseRow[3] ? orderBaseRow[3] : "", newItem.maHang || "", newItem.maVach || "", newItem.tenHang || "", itemQty, "", newItem.dvt || "", "", "Thêm mới vào đơn", "Mới", payload.actor || "", "Quản lý"]);
        logOrderChange(ss, soPhieu, "Thêm mã vào đơn", payload.actor, newItem.maHang, newItem.maVach, "", itemQty, newItem.tenHang || "");
        changeCount += 1;
        shouldNotify = true;
      }
    }
    var pendingUpdates = [];
    for (var i = 0; i < payload.updates.length; i++) {
      var u = payload.updates[i];
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
      var newMaVach = u.valMaVach !== undefined && u.valMaVach !== null && String(u.valMaVach).trim() !== "" ? String(u.valMaVach).trim() : maVach;
      var newDvt = u.valDvt !== undefined && u.valDvt !== null && String(u.valDvt).trim() !== "" ? String(u.valDvt).trim() : oldDvt;
      var newTenHang = u.valTenHang !== undefined && u.valTenHang !== null && String(u.valTenHang).trim() !== "" ? String(u.valTenHang).trim() : oldTenHang;
      if (!orderInfo) orderInfo = getThongTinPhieu(soPhieuValue);
      if (u.valSl !== "" && Number(u.valSl) === 0) {
        pendingUpdates.push({row: rowIndex, requestedQty: 0, actualQty: 0, note: "Đã hủy dòng", status: "Đã hủy dòng", actor: payload.actor || "", source: "Quản lý", maVach: newMaVach, dvt: newDvt, tenHang: newTenHang});
        logOrderChange(ss, soPhieuValue, "Hủy mã khỏi đơn", payload.actor, maHang, maVach, oldSl, 0, "Hủy bằng cập nhật số lượng");
        changeCount += 1;
        cancelledCount += 1;
        shouldNotify = true;
      } else if (u.valSl !== "") {
        var newVal = Number(u.valSl);
        pendingUpdates.push({row: rowIndex, requestedQty: newVal, actualQty: wasPacked ? "" : currentRow[8], note: wasPacked ? "Cần soạn lại sau khi chỉnh sửa" : "", status: "Mới", actor: payload.actor || "", source: "Quản lý", maVach: newMaVach, dvt: newDvt, tenHang: newTenHang});
        if (Number(oldSl) !== newVal) {
          logOrderChange(ss, soPhieuValue, "Sửa số lượng", payload.actor, maHang, maVach, oldSl, newVal, "");
          changeCount += 1;
          modifiedCount += 1;
          shouldNotify = true;
        }
        if (String(newMaVach || "").trim() !== String(maVach || "").trim()) {
          logOrderChange(ss, soPhieuValue, "Đổi mã vạch", payload.actor, maHang, maVach, maVach, newMaVach, "Đổi giữa mã lẻ/mã thùng");
          changeCount += 1;
          shouldNotify = true;
        }
        if (String(newDvt || "").trim() !== String(oldDvt || "").trim()) {
          logOrderChange(ss, soPhieuValue, "Đổi đơn vị tính", payload.actor, maHang, newMaVach || maVach, oldDvt, newDvt, "Đổi ĐVT theo mã vạch");
          changeCount += 1;
          shouldNotify = true;
        }
        if (String(newTenHang || "").trim() !== String(oldTenHang || "").trim()) {
          logOrderChange(ss, soPhieuValue, "Cập nhật tên hàng", payload.actor, maHang, newMaVach || maVach, oldTenHang, newTenHang, "Đồng bộ theo mã vạch mới");
          changeCount += 1;
          shouldNotify = true;
        }
      }
    }
    if (newRows.length) {
      var startRow = historySheet.getLastRow() + 1;
      historySheet.getRange(startRow, 1, newRows.length, 15).setValues(newRows);
    }
    if (pendingUpdates.length) {
      pendingUpdates.sort(function(a, b) { return a.row - b.row; });
      for (var g = 0; g < pendingUpdates.length; g++) {
        var itemUpdate = pendingUpdates[g];
        historySheet.getRange(itemUpdate.row, 8).setValue(itemUpdate.requestedQty);
        if (itemUpdate.actualQty === "") historySheet.getRange(itemUpdate.row, 9).clearContent();
        else historySheet.getRange(itemUpdate.row, 9).setValue(itemUpdate.actualQty);
        if (itemUpdate.maVach !== undefined && itemUpdate.maVach !== null && String(itemUpdate.maVach).trim() !== "") {
          historySheet.getRange(itemUpdate.row, 6).setValue(itemUpdate.maVach);
        }
        if (itemUpdate.tenHang !== undefined && itemUpdate.tenHang !== null && String(itemUpdate.tenHang).trim() !== "") {
          historySheet.getRange(itemUpdate.row, 7).setValue(itemUpdate.tenHang);
        }
        if (itemUpdate.dvt !== undefined && itemUpdate.dvt !== null && String(itemUpdate.dvt).trim() !== "") {
          historySheet.getRange(itemUpdate.row, 10).setValue(itemUpdate.dvt);
        }
        historySheet.getRange(itemUpdate.row, 12, 1, 4).setValues([[itemUpdate.note, itemUpdate.status, itemUpdate.actor, itemUpdate.source]]);
      }
    }
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
      sendTelegramOrderChangeSummary(orderInfo.soPhieu, orderInfo.khoXuat, orderInfo.khoNhan, actionLabel, changeCount, payload.actor, extraSummary, editPdfUrl);
    }
  } finally { lock.releaseLock(); }
  return { success: true };
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
  var row = [new Date(), payload.soPhieu, baseRow[2], baseRow[3], item.maHang || "", item.maVach || "", item.tenHang || "", quantity, "", item.dvt || "", "", "Thêm mới vào đơn", "Mới", payload.actor || "", "Quản lý"];
  historySheet.appendRow(row);
  logOrderChange(ss, payload.soPhieu, "Thêm mã vào đơn", payload.actor, item.maHang, item.maVach, "", quantity, item.tenHang || "");
  return { success: true };
}

function huyDongChiTietPhieu(payload) {
  requireAuthenticatedAction(payload);
  var ss = getSS();
  var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
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
  var values = historySheet.getRange(row, 1, 1, 13).getValues()[0];
  historySheet.getRange(row, 8).setValue(0);
  historySheet.getRange(row, 9).setValue(0);
  historySheet.getRange(row, 12).setValue("Đã hủy dòng");
  historySheet.getRange(row, 13).setValue("Đã hủy dòng");
  historySheet.getRange(row, 14).setValue(payload.actor || "");
  historySheet.getRange(row, 15).setValue("Quản lý");
  logOrderChange(ss, values[1], "Hủy mã khỏi đơn", payload.actor, values[4], values[5], values[7], 0, "Hủy từng dòng");
  return { success: true };
}

function huyPhieu(payload) {
  requireAdmin(payload.actor);
  var ss = getSS();
  var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
  ensureHistoryStatusColumn(historySheet);
  if (payload && payload.soPhieu && isOrderConfirmedForEditing(payload.soPhieu, historySheet)) {
    throw new Error("Đơn đã được xác nhận nhận hàng nên không thể hủy hoặc chỉnh sửa nữa.");
  }
  var data = historySheet.getDataRange().getValues();
  var found = false;
  var orderInfo = getThongTinPhieu(payload.soPhieu);
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(payload.soPhieu).trim()) {
      historySheet.getRange(i + 1, 13).setValue("Đã hủy đơn");
      historySheet.getRange(i + 1, 12).setValue("Đã hủy đơn");
      found = true;
    }
  }
  if (!found) throw new Error("Không tìm thấy đơn hàng.");
  logOrderChange(ss, payload.soPhieu, "Hủy đơn", payload.actor, "", "", "Đang xử lý", "Đã hủy đơn", payload.reason || "");
  if (orderInfo) {
    sendTelegramOrderCancelled(payload.soPhieu, orderInfo.khoXuat, orderInfo.khoNhan, payload.actor, payload.reason || "");
  }
  return { success: true };
}

function xacNhanNhanHang(payload) {
  requireAuthenticatedAction(payload);
  var ss = getSS();
  var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
  var receiveSheet = getReceiveSheet(ss);
  if (!historySheet) throw new Error("Không tìm thấy dữ liệu đơn hàng.");
  ensureHistoryStatusColumn(historySheet);
  var actorAccount = getAccountByActor(payload.actor);
  var expectedStore = payload.store || (actorAccount ? actorAccount.store : "");
  if (actorAccount && String(actorAccount.role).trim() !== "Admin") {
    if (!expectedStore || (String(actorAccount.store).trim() !== "Tất cả" && String(actorAccount.store).trim() !== String(expectedStore).trim())) {
      throw new Error("Bạn chỉ có thể xác nhận cho chi nhánh của mình.");
    }
  }
  var confirmations = payload.confirmations || [];
  if (!confirmations.length) throw new Error("Không có dữ liệu xác nhận.");

  var invalidRows = [];
  for (var c = 0; c < confirmations.length; c++) {
    var confCheck = confirmations[c];
    var rowCheck = Number(confCheck.row);
    if (!rowCheck || rowCheck < 2) continue;
    var checkValues = historySheet.getRange(rowCheck, 1, 1, 13).getValues()[0];
    var rowSoPhieu = checkValues[1] ? String(checkValues[1]).trim() : "";
    var rowStatus = checkValues[12] ? String(checkValues[12]).trim() : "Mới";
    var hasPackedQty = checkValues[8] !== "" && checkValues[8] !== null && checkValues[8] !== undefined;
    if (!rowSoPhieu || rowSoPhieu.toLowerCase() !== String(payload.soPhieu || "").trim().toLowerCase()) {
      invalidRows.push(rowCheck + " (sai số phiếu)");
      continue;
    }
    if (rowStatus === "Đã hủy dòng" || rowStatus === "Đã hủy đơn") {
      invalidRows.push(rowCheck + " (đã hủy)");
      continue;
    }
    if (rowStatus === "Đã xác nhận nhận hàng") {
      invalidRows.push(rowCheck + " (đã xác nhận)");
      continue;
    }
    if (!hasPackedQty && rowStatus !== "Đã soạn hàng") {
      invalidRows.push(rowCheck + " (chưa soạn)");
    }
  }
  if (invalidRows.length) {
    throw new Error("Không thể xác nhận đơn chưa soạn xong. Dòng lỗi: " + invalidRows.join(", "));
  }

  var confirmedTotal = 0;
  var changedCount = 0;
  var changedQtyTotal = 0;
  for (var i = 0; i < confirmations.length; i++) {
    var conf = confirmations[i];
    var row = Number(conf.row);
    var qty = Number(conf.receivedQty);
    if (!row || row < 2 || isNaN(qty) || qty < 0) continue;
    var values = historySheet.getRange(row, 1, 1, 13).getValues()[0];
    historySheet.getRange(row, 9).setValue(qty);
    historySheet.getRange(row, 13).setValue("Đã xác nhận nhận hàng");
    historySheet.getRange(row, 14).setValue(payload.actor || "");
    historySheet.getRange(row, 15).setValue("Xác nhận nhận hàng");
    var requestedQty = Number(values[7]) || 0;
    var previousQty = Number(conf.previousQty) || 0;
    var actualChanged = qty !== previousQty;
    if (actualChanged) {
      changedCount += 1;
      changedQtyTotal += Math.abs(qty - previousQty);
    }
    receiveSheet.appendRow([new Date(), payload.soPhieu, expectedStore, payload.actor, values[4] || "", values[5] || "", values[6] || "", qty, values[7] || 0, "Đã xác nhận nhận hàng"]);
    logOrderChange(ss, payload.soPhieu, "Xác nhận nhận hàng", payload.actor, values[4], values[5], values[7], qty, "Xác nhận bởi chi nhánh");
    confirmedTotal += qty;
  }

  // Trừ tồn kho ngay khi xác nhận nhận hàng để đơn kế tiếp nhìn thấy tồn thực tế.
  applyStockDeductionAfterReceive(historySheet, confirmations, payload.soPhieu);

  var orderInfo = getThongTinPhieu(payload.soPhieu);
  if (orderInfo) {
    SpreadsheetApp.flush();
    var receivePdfUrl = taoPdfDonHangVaLayLink(payload.soPhieu);
    sendTelegramReceiveConfirmation(payload.soPhieu, orderInfo.khoNhan || expectedStore, payload.actor, confirmations.length, confirmedTotal, changedCount, changedQtyTotal, receivePdfUrl);
  }
  return { success: true, count: confirmations.length };
}

function applyStockDeductionAfterReceive(historySheet, confirmations, soPhieu) {
  var ss = getSS();
  var stockSheet = ss.getSheetByName("TỔNG HỢP TỒN KHO");
  if (!stockSheet) return;

  var stockData = stockSheet.getDataRange().getValues();
  var stockConfig = getStockSheetConfig(stockData);

  for (var i = 0; i < confirmations.length; i++) {
    var conf = confirmations[i];
    var row = Number(conf.row);
    var qty = Number(conf.receivedQty);
    if (!row || row < 2 || isNaN(qty) || qty <= 0) continue;

    var orderRow = historySheet.getRange(row, 1, 1, 9).getValues()[0];
    var rowSoPhieu = orderRow[1] ? String(orderRow[1]).trim() : "";
    if (!rowSoPhieu || rowSoPhieu.toLowerCase() !== String(soPhieu || "").trim().toLowerCase()) continue;

    var khoXuat = orderRow[2] ? String(orderRow[2]).trim() : "";
    var maHang = orderRow[4] ? normalizeProductCode(orderRow[4]) : "";
    var maVach = orderRow[5] ? normalizeProductCode(orderRow[5]) : "";
    if (!khoXuat || (!maHang && !maVach)) continue;

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
      var isSameStore = false;
      for (var s = 0; s < rowStores.length; s++) {
        if (isStoreNameMatch(rowStores[s], khoXuat)) {
          isSameStore = true;
          break;
        }
      }
      if (!isSameStore) continue;

      var rowMaHang = normalizeProductCode((hasOwnCode ? rowMaHangRaw : currentMaHang) || "");
      var rowMaVach = normalizeProductCode((hasOwnCode ? rowMaVachRaw : currentMaVach) || "");
      var codeMatch = (maHang && rowMaHang && areCodesEquivalent(maHang, rowMaHang)) || (maVach && rowMaVach && areCodesEquivalent(maVach, rowMaVach));
      if (!codeMatch) continue;

      var oldStock = parseQuantityValue(rowStock[stockConfig.tonKhoIdx]);
      var newStock = oldStock - qty;
      stockData[k][stockConfig.tonKhoIdx] = newStock;
      stockSheet.getRange(k + 1, stockConfig.tonKhoIdx + 1).setValue(newStock);
      break;
    }
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
    dataArr.push([stt++, finalItems[j].maHang, finalItems[j].maVach, finalItems[j].tenHang, finalItems[j].dvt, finalItems[j].sl]);
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
  for (var i = 0; i < values.length; i++) {
    var sp = values[i][1] != null ? String(values[i][1]).trim() : "";
    if (!sp) continue;
    var hit = false;
    if (key && (sp.indexOf(key) !== -1 || orderKeysMatch_(sp, key))) hit = true;
    if (!hit && digits && sp.indexOf(digits) !== -1) hit = true;
    if (!hit) continue;
    found.push({
      sheetRow: start + i,
      soPhieu: sp,
      slGoc: values[i][7],
      col9: values[i][8],
      status: values[i][12],
      col16_slSoan: values[i][15],
      actor: values[i][13],
      source: values[i][14]
    });
  }
  return {
    success: true,
    key: key,
    digits: digits,
    lastRow: lastRow,
    lastCol: lastCol,
    matchCount: found.length,
    found: found,
    _debugRun: "debug-order-v1"
  };
}

function luuSoSoanHangVaAnh(payload) {
  var ss = getSS();
  var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
  ensureHistoryStatusColumn(historySheet);
  var updates = payload.updates || [];
  var images = payload.images || {}; 
  var lock = LockService.getDocumentLock();
  try {
    lock.waitLock(20000);
    var actor = payload.actor || "";
    if (updates.length) {
      var minRow = updates[0].row;
      var maxRow = updates[0].row;
      var rowMap = {};
      for (var ui = 0; ui < updates.length; ui++) {
        minRow = Math.min(minRow, updates[ui].row);
        maxRow = Math.max(maxRow, updates[ui].row);
        rowMap[updates[ui].row] = updates[ui];
      }
      // Ghi từng dòng bằng A1 (tránh nhầm getRange numRows/endRow) — cột 9 + 13–16
      for (var wr = 0; wr < updates.length; wr++) {
        var wRow = Number(updates[wr].row);
        if (!wRow || wRow < 2 || isNaN(wRow)) continue;
        var wVal = updates[wr].val;
        if (wVal !== "" && wVal !== null && wVal !== undefined) {
          var parsedVal = Number(wVal);
          historySheet.getRange(wRow, 9).setValue(parsedVal);
          historySheet.getRange(wRow, 13).setValue("Đã soạn hàng");
          historySheet.getRange(wRow, 14).setValue(actor);
          historySheet.getRange(wRow, 15).setValue("Soạn hàng");
          historySheet.getRange(wRow, 16).setValue(parsedVal); // SL Giao (Soạn)
        } else {
          historySheet.getRange(wRow, 9).clearContent();
          historySheet.getRange(wRow, 16).clearContent();
          historySheet.getRange(wRow, 13).setValue("Mới");
        }
      }
    }
    var targetFolder;
    try {
      var folders = DriveApp.getFoldersByName("dieuchuyenhanghoa");
      if (folders.hasNext()) targetFolder = folders.next();
      else targetFolder = DriveApp.createFolder("dieuchuyenhanghoa");
    } catch (e) { throw new Error("Chưa cấp quyền Drive"); }
    
    var anhDaLuu = 0;
    for (var rIdx in images) {
      if (images[rIdx] && images[rIdx].indexOf("base64,") !== -1) {
        var splitBase = images[rIdx].split('base64,');
        var file = targetFolder.createFile(Utilities.newBlob(Utilities.base64Decode(splitBase[1]), "image/jpeg", 'XacNhan_' + new Date().getTime() + '.jpg'));
        try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
        historySheet.getRange(parseInt(rIdx), 11).setValue(file.getUrl());
        anhDaLuu++;
      }
    }
    SpreadsheetApp.flush();

    var soPhieu = "";
    var khoXuat = "";
    var khoNhan = "";
    var totalRows = 0;
    var missingCount = 0;
    var extraCount = 0;
    for (var u = 0; u < updates.length; u++) {
      var notifyRow = updates[u].row;
      var currentSoPhieu = historySheet.getRange(notifyRow, 2).getValue();
      if (currentSoPhieu) { soPhieu = String(currentSoPhieu).trim(); break; }
    }
    var historyData = readHistoryDataPack_(historySheet, 12000).data;
    for (var j = 1; j < historyData.length; j++) {
      var statRow = historyData[j];
      var rowSoPhieu = statRow[1] ? String(statRow[1]).trim() : "";
      if (!rowSoPhieu || rowSoPhieu !== soPhieu) continue;
      totalRows += 1;
      var requestedQty = Number(statRow[7]) || 0;
      var actualQty = (statRow[8] !== "" && statRow[8] !== undefined) ? Number(statRow[8]) : requestedQty;
      if (actualQty < requestedQty) missingCount += 1;
      if (actualQty > requestedQty) extraCount += 1;
      if (!khoXuat && statRow[2]) khoXuat = String(statRow[2]).trim();
      if (!khoNhan && statRow[3]) khoNhan = String(statRow[3]).trim();
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
        actor: payload.actor || "Chi nhánh"
      }
    };
  } finally { lock.releaseLock(); }
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

/** Ghép ngày + "HH:mm" → Date (timezone script) */
function combineDateAndTime_(dateObj, timeHHmm) {
  if (!dateObj || !(dateObj instanceof Date) || isNaN(dateObj.getTime())) return null;
  var raw = String(timeHHmm || "").trim();
  var m = raw.match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  var hh = Number(m[1]);
  var mm = Number(m[2]);
  if (isNaN(hh) || isNaN(mm) || hh < 0 || hh > 23 || mm < 0 || mm > 59) return null;
  var out = new Date(dateObj.getTime());
  out.setHours(hh, mm, 0, 0);
  return out;
}

/**
 * Cửa sổ ngày tổng hợp / giao (packing day D):
 * - Đơn chính (tổng hợp): (D-1) 10:00 → D 08:00 (không gồm 08:00)
 * - Bổ sung / đơn mới: D 08:00 → D 10:00 (gồm 10:00)
 * - Sau D 10:00 → thuộc ngày tổng hợp D+1
 */
function getPackingDayWindows_(packingDayDate, opts) {
  opts = opts || {};
  var packingDay = packingDayDate instanceof Date && !isNaN(packingDayDate.getTime())
    ? new Date(packingDayDate.getTime())
    : (getScriptTodayStart_() || new Date());
  packingDay.setHours(0, 0, 0, 0);
  var prevDay = new Date(packingDay.getTime());
  prevDay.setDate(prevDay.getDate() - 1);

  var mainStartTime = opts.mainStartTime || "10:00";
  var mainEndTime = opts.mainEndTime || opts.suppStartTime || "08:00";
  var suppEndTime = opts.suppEndTime || "10:00";

  var mainStart = combineDateAndTime_(prevDay, mainStartTime);
  var mainEnd = combineDateAndTime_(packingDay, mainEndTime);
  var suppEnd = combineDateAndTime_(packingDay, suppEndTime);
  var tz = Session.getScriptTimeZone();
  return {
    packingDay: packingDay,
    prevDay: prevDay,
    mainStart: mainStart,
    mainEnd: mainEnd,
    suppStart: mainEnd,
    suppEnd: suppEnd,
    packingDayStr: Utilities.formatDate(packingDay, tz, "yyyy-MM-dd"),
    prevDayStr: Utilities.formatDate(prevDay, tz, "yyyy-MM-dd"),
    mainLabel: mainStart && mainEnd
      ? (Utilities.formatDate(mainStart, tz, "dd/MM HH:mm") + " → " + Utilities.formatDate(mainEnd, tz, "dd/MM HH:mm"))
      : "",
    suppLabel: mainEnd && suppEnd
      ? (Utilities.formatDate(mainEnd, tz, "dd/MM HH:mm") + " → " + Utilities.formatDate(suppEnd, tz, "dd/MM HH:mm"))
      : ""
  };
}

function isInPackingMainWindow_(createdMs, win) {
  if (!win || !win.mainStart || !win.mainEnd || isNaN(createdMs)) return false;
  return createdMs >= win.mainStart.getTime() && createdMs < win.mainEnd.getTime();
}

function isInPackingSuppWindow_(createdMs, win) {
  if (!win || !win.suppStart || !win.suppEnd || isNaN(createdMs)) return false;
  return createdMs >= win.suppStart.getTime() && createdMs <= win.suppEnd.getTime();
}

function isInPackingDayWindow_(createdMs, win) {
  if (!win || !win.mainStart || !win.suppEnd || isNaN(createdMs)) return false;
  return createdMs >= win.mainStart.getTime() && createdMs <= win.suppEnd.getTime();
}

function toMillisSafe_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value.getTime();
  if (value === null || value === undefined || value === "") return NaN;
  var d = new Date(value);
  return isNaN(d.getTime()) ? NaN : d.getTime();
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

function getDanhSachDonSoanHang(ngayYYYYMMDD, userRole, userStore, ngayToYYYYMMDD) {
  // #region agent log
  var _dbgT0 = Date.now();
  // #endregion
  // ngayTo = ngày tổng hợp/giao (packing day). Fallback: ngay hoặc hôm nay.
  var packingDay = parseDateInputYYYYMMDD(ngayToYYYYMMDD) || parseDateInputYYYYMMDD(ngayYYYYMMDD) || getScriptTodayStart_() || new Date();
  packingDay.setHours(0, 0, 0, 0);
  var win = getPackingDayWindows_(packingDay);
  var orders = getEligibleOrdersForSoanHang(packingDay, userRole, userStore, null, packingDay, win);
  // #region agent log
  var _dbgMs = Date.now() - _dbgT0;
  // #endregion
  return {
    success: true,
    date: win.prevDayStr,
    dateTo: win.packingDayStr,
    packingDay: win.packingDayStr,
    mainWindow: win.mainLabel,
    suppWindow: win.suppLabel,
    total: orders.length,
    orders: orders,
    _debugTotalMs: _dbgMs,
    _debugRun: "packing-window-v1"
  };
}

function getEligibleOrdersForSoanHang(baseDate, userRole, userStore, historyPack, endDate, packingWin) {
  var ss = getSS();
  var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
  if (!historySheet) return [];
  historyPack = historyPack || readHistoryDataPack_(historySheet, 5000);
  var data = historyPack.data;
  if (!data || data.length < 2) return [];

  var packingDay = endDate || baseDate;
  var win = packingWin || getPackingDayWindows_(packingDay);
  var map = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row) continue;
    var soPhieu = row[1] ? String(row[1]).trim() : "";
    if (!soPhieu) continue;

    var ngayTao = row[0];
    var createdMs = toMillisSafe_(ngayTao);
    // Chỉ lấy đơn có dòng trong cửa sổ ngày tổng hợp (prev 10h → D 10h)
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
        createdAt: ngayTao instanceof Date ? ngayTao : new Date(createdMs),
        hasPacked: false,
        hasConfirmed: false,
        hasCancelled: false,
        inMain: false,
        inSupp: false
      };
    }

    var entry = map[soPhieu];
    if (displayStatus === "Đã xác nhận") entry.hasConfirmed = true;
    else if (displayStatus === "Đã soạn") entry.hasPacked = true;
    else if (displayStatus === "Đã hủy" || displayStatus === "Đã hủy dòng") entry.hasCancelled = true;

    if (!isNaN(createdMs)) {
      if (isInPackingMainWindow_(createdMs, win)) entry.inMain = true;
      if (isInPackingSuppWindow_(createdMs, win)) entry.inSupp = true;
      var createdDate = ngayTao instanceof Date ? ngayTao : new Date(createdMs);
      if (entry.createdAt instanceof Date && createdDate.getTime() < entry.createdAt.getTime()) {
        entry.createdAt = createdDate;
      }
    }
  }

  var orders = [];
  var tz = Session.getScriptTimeZone();
  for (var key in map) {
    var item = map[key];
    if (item.hasPacked || item.hasConfirmed || item.hasCancelled) continue;
    var bucket = item.inMain && item.inSupp ? "cả hai" : (item.inSupp ? "bổ sung" : "chính");
    orders.push({
      soPhieu: item.soPhieu,
      khoXuat: item.khoXuat,
      khoNhan: item.khoNhan,
      thoiGianDat: Utilities.formatDate(item.createdAt, tz, "dd/MM/yyyy HH:mm"),
      thoiGianDatMillis: item.createdAt.getTime(),
      packingBucket: bucket
    });
  }

  orders.sort(function(a, b) { return a.thoiGianDatMillis - b.thoiGianDatMillis; });
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
  var newAfterTime = payload && payload.newAfterTime ? String(payload.newAfterTime).trim() : "08:00";
  var newBeforeTime = payload && payload.newBeforeTime ? String(payload.newBeforeTime).trim() : "10:00";
  _dbgMark("start", {
    selectedCount: payload && payload.selectedOrders ? payload.selectedOrders.length : 0,
    algo: "ton-q7",
    onlyNewItems: onlyNewItems,
    newAfterTime: newAfterTime,
    newBeforeTime: newBeforeTime
  });
  // #endregion

  var ss = getSS();
  var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
  if (!historySheet) throw new Error("Không tìm thấy sheet Lịch Sử Xuất Kho.");

  // Ngày tổng hợp/giao = ngayTo (ưu tiên) hoặc ngay
  var packingDayInput = parseDateInputYYYYMMDD(payload && payload.ngayTo ? payload.ngayTo : "")
    || parseDateInputYYYYMMDD(payload && payload.ngay ? payload.ngay : "")
    || getScriptTodayStart_()
    || new Date();
  packingDayInput.setHours(0, 0, 0, 0);
  var win = getPackingDayWindows_(packingDayInput, {
    mainEndTime: newAfterTime || "08:00",
    suppEndTime: newBeforeTime || "10:00"
  });
  if (!win.mainStart || !win.mainEnd || !win.suppEnd) {
    return {
      success: false,
      msg: "Giờ chốt không hợp lệ. Dùng 08:00 / 10:00.",
      _debugTimings: _dbgSteps,
      _debugTotalMs: Date.now() - _dbgT0,
      _debugRun: "packing-window-v1"
    };
  }
  var baseDate = win.prevDay;
  var endDate = win.packingDay;
  var packingDay = win.packingDay;
  var baseDateStr = win.packingDayStr;
  var newAfterLabel = Utilities.formatDate(win.suppStart, Session.getScriptTimeZone(), "dd/MM HH:mm");
  var newBeforeLabel = Utilities.formatDate(win.suppEnd, Session.getScriptTimeZone(), "dd/MM HH:mm");
  var mainWindowLabel = win.mainLabel;

  var userRole = payload && payload.userRole ? payload.userRole : "";
  var userStore = payload && payload.userStore ? payload.userStore : "";
  var selectedOrdersRaw = payload && payload.selectedOrders && payload.selectedOrders.length ? payload.selectedOrders : [];
  var selectedSet = {};
  for (var so0 = 0; so0 < selectedOrdersRaw.length; so0++) {
    var pick0 = String(selectedOrdersRaw[so0] || "").trim();
    if (pick0) selectedSet[pick0] = true;
  }

  var historyPack;
  if (Object.keys(selectedSet).length) {
    historyPack = readHistoryForSelectedOrders_(historySheet, selectedSet, "", 2500);
  } else {
    historyPack = readHistoryDataPack_(historySheet, 3000);
    var eligibleOrders = getEligibleOrdersForSoanHang(packingDay, userRole, userStore, historyPack, packingDay, win);
    for (var eo = 0; eo < eligibleOrders.length; eo++) {
      selectedSet[String(eligibleOrders[eo].soPhieu).trim()] = true;
    }
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
      msg: "Không có đơn hợp lệ để tạo bảng soạn. Đơn đã soạn hoặc đã giao không được tính.",
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

  // Pass 1: mốc tạo đơn sớm nhất (để nhận đơn mới trong cửa sổ 8h–10h)
  var orderMinCreated = {};
  for (var p1 = 1; p1 < data.length; p1++) {
    var rowP1 = data[p1];
    if (!rowP1) continue;
    var soP1 = rowP1[1] ? String(rowP1[1]).trim() : "";
    if (!soP1 || !selectedSet[soP1]) continue;
    var stP1 = rowP1[12] ? String(rowP1[12]).trim() : "";
    if (stP1 === "Đã hủy đơn" || stP1 === "Đã hủy dòng") continue;
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
    if (!soPhieu || !selectedSet[soPhieu]) continue;

    var khoXuat = row[2] ? String(row[2]).trim() : "";
    var khoNhan = row[3] ? String(row[3]).trim() : "";
    var maHang = row[4] ? String(row[4]).trim() : "";
    var maVach = row[5] ? String(row[5]).trim() : "";
    var tenHang = row[6] ? String(row[6]).trim() : "";
    var soLuong = Number(row[7]) || 0;
    var dvtRaw = String(row[9] || "").trim();
    var dvt = dvtRaw;
    if ((!dvt || normalizeHeaderText(dvt) === "cai") && catalogLookup) {
      dvt = resolveDvtValue(catalogLookup, maHang, maVach, dvtRaw) || dvtRaw;
    }
    var noteText = row[11] != null ? String(row[11]).trim() : "";
    var status = row[12] ? String(row[12]).trim() : "Đang xử lý";

    if (!khoNhan) continue;
    if (status === "Đã hủy đơn" || status === "Đã hủy dòng") continue;
    if (soLuong <= 0) continue;

    var createdMs = toMillisSafe_(row[0]);
    if (isNaN(createdMs)) {
      skippedByTime++;
      continue;
    }

    if (onlyNewItems) {
      // Tick "Bảng bổ sung": lấy TẤT CẢ dòng trong khung 8h–10h (đơn mới + mã thêm + đơn chưa soạn trong khung).
      // Không bỏ qua chỉ vì thiếu ghi chú "Thêm mới vào đơn".
      if (!isInPackingSuppWindow_(createdMs, win)) {
        skippedByTime++;
        continue;
      }
      includedNewRows++;
    } else {
      // Tổng hợp đầy đủ: cả đợt chính + bổ sung trong ngày D
      // (D-1) 10:00 → D 10:00 — đơn chưa soạn đều vào bảng, không cắt ở 8h.
      if (!isInPackingDayWindow_(createdMs, win)) {
        skippedByTime++;
        continue;
      }
      if (isInPackingMainWindow_(createdMs, win)) includedMainRows++;
      else includedNewRows++;
    }

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
    return {
      success: false,
      msg: onlyNewItems
        ? ("Không có dòng hàng trong khung bổ sung " +
          (newAfterLabel || "?") + " → " + (newBeforeLabel || "?") +
          ".\nĐã bỏ qua " + skippedByTime + " dòng ngoài khung 8h–10h.")
        : ("Không gom được mã hàng trong cửa sổ tổng hợp " +
          (mainWindowLabel || "(D-1) 10:00") + " → " + (newBeforeLabel || "D 10:00") +
          " (gồm cả đợt chính + bổ sung).\nĐã bỏ qua " + skippedByTime + " dòng ngoài khung / sau 10h."),
      onlyNewItems: onlyNewItems,
      newAfterLabel: newAfterLabel,
      newBeforeLabel: newBeforeLabel,
      mainWindowLabel: mainWindowLabel,
      packingDay: baseDateStr,
      _debugTimings: _dbgSteps,
      _debugTotalMs: Date.now() - _dbgT0,
      _debugRun: "packing-window-v1",
      _debugInfo: {
        baseDateStr: baseDateStr,
        selectedList: Object.keys(selectedSet),
        scannedRows: historyPack.scannedRows || 0,
        skippedByTime: skippedByTime,
        skippedNotSupplement: skippedNotSupplement,
        includedNewRows: includedNewRows,
        includedMainRows: includedMainRows
      }
    };
  }

  storeList.sort(function(a, b) {
    return formatShortStoreLabel(a).localeCompare(formatShortStoreLabel(b));
  });
  var activeMap = getActiveStoreMap();

  // Tồn kho soạn hàng: chỉ dùng sheet nhẹ TON_Q7 (tách lúc import file tồn)
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
  // #region agent log
  _dbgMark("stockIndex", { via: "TON_Q7", stockReady: stockReady, q7Keys: Object.keys(q7Map).length, forceStock: forceStock, itemCount: keys.length });
  // #endregion

  var sheetName = onlyNewItems ? "__TMP_SOAN_BO_SUNG" : "__TMP_SOAN_NGAY_MAI";
  var reportSheet = recreateTempSheetFast_(ss, sheetName);
  var packingDayTitle = Utilities.formatDate(packingDay, Session.getScriptTimeZone(), "dd/MM/yyyy");
  var title = onlyNewItems
    ? ("BẢNG BỔ SUNG (toàn bộ đơn/dòng 8h–10h chưa soạn) — ngày " + packingDayTitle +
      " | " + newAfterLabel + " → " + newBeforeLabel)
    : ("BẢNG TỔNG HỢP SOẠN HÀNG NGÀY " + packingDayTitle +
      " | Chính+Bổ sung: " + mainWindowLabel + " → " + newBeforeLabel);
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
    var canhBao = "Chưa có TON_Q7";
    if (stockReady) {
      var thieu = it.totalQty - stock;
      canhBao = thieu > 0 ? ("THIẾU " + thieu) : "OK";
      if (thieu > 0) missingLines += 1;
    }

    var rowOut = [0, it.maHang || "", it.maVach || "", it.tenHang || "", dvtOut || "", stockReady ? stock : "", it.totalQty];
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
  var summaryLine = "Tổng đơn: " + Object.keys(orderSeen).length + " | Tổng mã: " + rows.length +
    (stockReady ? (" | Mã thiếu: " + missingLines + " | Tồn: TON_Q7") : " | Tồn: chưa có sheet TON_Q7 — Admin import lại file tồn") +
    (onlyNewItems
      ? (" | Bổ sung: " + includedNewRows + " dòng | " + newAfterLabel + " → " + newBeforeLabel)
      : (" | Chính: " + includedMainRows + " + Bổ sung: " + includedNewRows + " dòng | " + mainWindowLabel + " → " + newBeforeLabel));
  var sheetBlock = [
    padRow_([title], colCount),
    padRow_([
      "Ngày tổng hợp: " + packingDayTitle +
      " | Chính: " + mainWindowLabel +
      " | Bổ sung: " + newAfterLabel + " → " + newBeforeLabel +
      " | Sau 10h → ngày hôm sau" +
      " | Gom theo mã" + (onlyNewItems ? " (chỉ 8h–10h)" : " (chính + bổ sung)") + " | Stock từ " + TON_Q7_SHEET_NAME
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
    onlyNewItems: onlyNewItems,
    packingDay: baseDateStr,
    mainWindowLabel: mainWindowLabel,
    newAfterLabel: newAfterLabel,
    newBeforeLabel: newBeforeLabel,
    url: "https://docs.google.com/spreadsheets/d/" + SHEET_ID + "/edit#gid=" + reportSheet.getSheetId(),
    _debugTimings: _dbgSteps,
    _debugTotalMs: _dbgTotalMs,
    _debugRun: "packing-window-v1",
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