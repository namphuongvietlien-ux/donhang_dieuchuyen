// ============================================================
// api_routes.gs — doGet / doPost / Telegram / Auth / Xuat ban
// ============================================================

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
          case 'saveChildVariants':
            requireAdminAction(action, payload.payload || {});
            result = saveChildVariants_(payload.payload || {});
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
        case 'getParentVariantGroupsAdmin':
          res = getParentVariantGroupsAdmin_(e.parameter.q || '', e.parameter.limit || 200);
          break;
        case 'getChildVariantsForAdmin':
          res = getChildVariantsForAdmin_(e.parameter.parentSku || e.parameter.parent || '');
          break;
        case 'getOrderDetail':
          res = getOrderDetail(e.parameter.soPhieu || '');
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
  var soHoaDon = normalizeOrderCodeText(payload.soHoaDon || "") || String(payload.soHoaDon || "").trim();
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
  var filterHd = normalizeOrderCodeText(soHoaDonFilter || "");
  var filterStore = normalizeStoreName(userStore || "");
  var map = {};

  for (var i = 0; i < body.length; i++) {
    var row = body[i];
    if (!row) continue;
    var maPhieu = row[2] ? String(row[2]).trim() : "";
    var soHd = row[1] ? String(row[1]).trim() : "";
    if (!maPhieu && !soHd) continue;
    if (filterHd) {
      var soHdNorm = normalizeOrderCodeText(soHd);
      if (soHdNorm.indexOf(filterHd) === -1) continue;
    }
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
