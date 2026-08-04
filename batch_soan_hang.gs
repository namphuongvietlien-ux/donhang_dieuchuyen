// ============================================================
// batch_soan_hang.gs — Batch A/B1/C: soan / nhan / sua / huy don
// ============================================================


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
