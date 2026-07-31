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
// ========================================================================
function doPost(e) {
  try {
    var contents = e && e.postData && e.postData.contents ? e.postData.contents : "";
    if (!contents) {
      Logger.log("doPost warning: empty payload");
      return ContentService.createTextOutput("OK");
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
          default:
            result = { error: 'Unknown action: ' + action };
        }
        return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
      } catch(apiErr) {
        return ContentService.createTextOutput(JSON.stringify({ error: apiErr.message })).setMimeType(ContentService.MimeType.JSON);
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
        case 'layDanhSachPhieuTheoFilter':
          res = layDanhSachPhieuTheoFilter(e.parameter.khoNhan || '', e.parameter.ngay || '', e.parameter.userRole || '', e.parameter.userStore || '');
          break;
        case 'getChiTietPhieu':
          res = getChiTietPhieu(e.parameter.soPhieu || '', e.parameter.storeName || '');
          break;
        case 'getThongTinPhieu':
          res = getThongTinPhieu(e.parameter.soPhieu || '');
          break;
        case 'getDonHangTheoNgay':
          res = getDonHangTheoNgay(e.parameter.ngay || 'today', e.parameter.userRole || '', e.parameter.userStore || '');
          break;
        case 'getDashboardSummary':
          res = getDashboardSummary(e.parameter.userRole || '', e.parameter.userStore || '', e.parameter.timeline || '2days', e.parameter.fromDate || '', e.parameter.toDate || '');
          break;
        case 'getChiTietDonHangMobile':
          res = getChiTietDonHangMobile(e.parameter.soPhieu || '');
          break;
        case 'getDanhSachTaiKhoan':
          res = getDanhSachTaiKhoan();
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
    text: text,
    parse_mode: "Markdown"
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
  var data = historySheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var soPhieu = data[i][1] ? data[i][1].toString().trim() : "";
    var khoNhan = data[i][3] ? data[i][3].toString().trim() : "";
    var khoXuat = data[i][2] ? data[i][2].toString().trim() : "";
    var slThucTe = data[i][8];
    var rowStatus = data[i][12] ? String(data[i][12]).trim() : "Mới";
    if (!soPhieu || !khoNhan) continue;
    if (khoNhan !== storeName) continue;
    var isDaXuLy = getDisplayOrderStatus(rowStatus, slThucTe) !== "Mới";
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

function sendTelegramMessage(soPhieu, khoXuat, khoNhan, itemCount) {
  var typeLabel = soPhieu.indexOf("DH") !== -1 ? "ĐƠN HÀNG MỚI" : "LỆNH ĐIỀU CHUYỂN MỚI";
  var kxShort = STORE_MAP[khoXuat] || khoXuat;
  var knShort = STORE_MAP[khoNhan] || khoNhan;
  var text = "📦 *THÔNG BÁO " + typeLabel + "*\n" +
             "*Trạng thái:* Mới\n" +
             "*Số phiếu:* " + soPhieu + "\n" +
             "*Kho xuất:* " + khoXuat + " (" + kxShort + ")\n" +
             "*Kho nhận:* " + khoNhan + " (" + knShort + ")\n" +
             "*Số mặt hàng:* " + itemCount + "\n\n" +
             "Mở chi tiết đơn: " + getOrderWebUrl(soPhieu);
  if (TELEGRAM_CHAT_ID) {
    sendTelegramText(TELEGRAM_CHAT_ID, text);
  }
  sendTelegramTextToStores([khoNhan, khoXuat], "📌 Có đơn mới dành cho kho của bạn:\n" + text);
}

function sendTelegramOrderReady(soPhieu, khoNhan) {
  var knShort = STORE_MAP[khoNhan] || khoNhan;
  var text = "✅ *ĐÃ HOÀN THÀNH SOẠN HÀNG*\n" +
             "*Trạng thái:* Đã soạn\n" +
             "*Số phiếu:* " + soPhieu + "\n" +
             "*Kho nhận:* " + khoNhan + " (" + knShort + ")\n\n" +
             "Mở chi tiết đơn: " + getOrderWebUrl(soPhieu);
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

function sendTelegramOrderChangeSummary(soPhieu, khoXuat, khoNhan, actionLabel, changeCount, actor, extraText) {
  var kxShort = STORE_MAP[khoXuat] || khoXuat;
  var knShort = STORE_MAP[khoNhan] || khoNhan;
  var text = "🔄 *ĐƠN ĐÃ THAY ĐỔI*\n" +
             "*Số phiếu:* " + soPhieu + "\n" +
             "*Hành động:* " + actionLabel + "\n" +
             "*Số mã thay đổi:* " + changeCount + "\n" +
             "*Người thực hiện:* " + (actor || "Không xác định") + "\n" +
             (extraText ? "*Chi tiết:* " + extraText + "\n" : "") +
             "Mở chi tiết đơn: " + getOrderWebUrl(soPhieu);
  if (TELEGRAM_CHAT_ID) {
    sendTelegramText(TELEGRAM_CHAT_ID, text);
  }
  if (khoNhan) {
    sendTelegramTextToStoreUsers(khoNhan, "📢 Đơn hàng vừa được cập nhật:\n" + text);
  }
}

function sendTelegramPackingSummary(soPhieu, khoXuat, khoNhan, changedCount, totalRows, statusLabel, missingCount, extraCount, actor) {
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
             "Mở chi tiết đơn: " + getOrderWebUrl(soPhieu);
  if (TELEGRAM_CHAT_ID) {
    sendTelegramText(TELEGRAM_CHAT_ID, text);
  }
  if (khoNhan) {
    sendTelegramTextToStoreUsers(khoNhan, "✅ Đơn đã được soạn xong:\n" + text);
  }
}

function sendTelegramReceiveConfirmation(soPhieu, khoNhan, actor, count, confirmedTotal, changedCount, changedQtyTotal) {
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
             "Mở chi tiết đơn: " + getOrderWebUrl(soPhieu, "quan-ly", true);
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
  var data = historySheet.getDataRange().getValues();
  var target = String(soPhieu).trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] && String(data[i][1]).trim().toLowerCase() === target) {
      return {
        soPhieu: String(data[i][1]).trim(),
        khoXuat: data[i][2] ? String(data[i][2]).trim() : "",
        khoNhan: data[i][3] ? String(data[i][3]).trim() : ""
      };
    }
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
  for (var i = 1; i < data.length; i++) { users.push({ user: data[i][0], pass: data[i][1], role: data[i][2], store: data[i][3] }); }
  return users;
}

function taoTaiKhoanMoi(payload) {
  var ss = getSS();
  var sheet = getOrCreateUserSheet(ss);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) { if (String(data[i][0]).trim() === payload.user) return { success: false, msg: "Tên đăng nhập đã tồn tại!" }; }
  sheet.appendRow([payload.user, payload.pass, payload.role, payload.store]);
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

function findHeaderRowIndex(data, maxScanRows) {
  var limit = Math.min(data.length, maxScanRows || 8);
  for (var r = 0; r < limit; r++) {
    var row = data[r];
    if (!row) continue;
    var hasMarker = false;
    for (var c = 0; c < row.length; c++) {
      var token = normalizeHeaderText(row[c]);
      if (!token) continue;
      if (token.indexOf('mahang') !== -1 || token.indexOf('mavach') !== -1 || token.indexOf('tenhang') !== -1 || token.indexOf('tonkho') !== -1 || token.indexOf('soluongton') !== -1 || token === 'kho' || token.indexOf('cuahang') !== -1) {
        hasMarker = true;
        break;
      }
    }
    if (hasMarker) return r;
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

function getStockSheetConfig(stockData) {
  var headerIndex = findHeaderRowIndex(stockData, 10);
  if (headerIndex < 0) {
    return { startRow: 4, storeIndexes: [0, 7], maHangIdx: 1, maVachIdx: 2, tonKhoIdx: 6 };
  }
  var header = stockData[headerIndex] || [];
  var storeIndexes = findAllColumnIndicesByAliases(header, ['kho', 'cuahang', 'chinhanh', 'store', 'tenkho']);
  var maHangIdx = findColumnIndexByAliases(header, ['mahang', 'sku', 'mahh', 'code', 'itemcode']);
  var maVachIdx = findColumnIndexByAliases(header, ['mavach', 'barcode', 'ean', 'barcodeid']);
  var tonKhoIdx = findColumnIndexByAliases(header, ['tonkho', 'soluongton', 'stock', 'onhand', 'slton', 'qty']);
  return {
    startRow: headerIndex + 1,
    storeIndexes: storeIndexes.length ? storeIndexes : [0, 7],
    maHangIdx: maHangIdx === -1 ? 1 : maHangIdx,
    maVachIdx: maVachIdx === -1 ? 2 : maVachIdx,
    tonKhoIdx: tonKhoIdx === -1 ? 6 : tonKhoIdx
  };
}

function getRowStoreNames(row, stockConfig) {
  var stores = [];
  for (var i = 0; i < stockConfig.storeIndexes.length; i++) {
    var name = getCellValue(row, stockConfig.storeIndexes[i], "");
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
    var dvt = getCellValue(row, idxDvt, "Cái");
    var kho = getCellValue(row, idxKho, "");
    var tonRaw = idxTonKho !== -1 ? row[idxTonKho] : "";
    var tonKho = (tonRaw === "" || tonRaw === null || tonRaw === undefined) ? "" : Number(tonRaw);

    if (!maHang && !maVach && !tenHang) continue;

    parsed.push({
      maHang: maHang,
      maVach: maVach,
      tenHang: tenHang,
      dvt: dvt || "Cái",
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
  if (fileData && importType) {
    if (importType === 'stock') {
      var stockSheetDirect = getOrCreateStockSheet(ss);
      var stockWriteInfo = writeImportedDataToSheet(stockSheetDirect, fileData);
      SpreadsheetApp.flush();
      return {
        success: true,
        importType: importType,
        targetSheet: 'TỔNG HỢP TỒN KHO',
        updatedRows: stockWriteInfo.rows,
        updatedCols: stockWriteInfo.cols,
        msg: 'Đã cập nhật file tồn kho lên sheet TỔNG HỢP TỒN KHO.'
      };
    }
    if (importType === 'catalog') {
      var catalogSheetDirect = getOrCreateCatalogSheet(ss);
      var catalogWriteInfo = writeImportedDataToSheet(catalogSheetDirect, fileData);
      SpreadsheetApp.flush();
      return {
        success: true,
        importType: importType,
        targetSheet: 'Data_Excel',
        updatedRows: catalogWriteInfo.rows,
        updatedCols: catalogWriteInfo.cols,
        msg: 'Đã cập nhật file nhập khẩu thông tin lên sheet Data_Excel.'
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
      dvt: item.dvt || "Cái"
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

  return {
    success: true,
    catalogUpdated: catalogRows.length,
    stockUpdated: stockUpdated,
    sourceSheet: sourceSheetName,
    warnings: warnings
  };
}

// --- API: LẤY DATA BAN ĐẦU ---
function getInitialData() {
  try {
    var ss = getSS();
    var danhMucHangHoa = {};
    var stores = ["Kho Địa điểm kinh doanh Q7"];
    
    var tonKhoSheet = ss.getSheetByName("TỔNG HỢP TỒN KHO");
    if (tonKhoSheet) {
      var tonKhoData = tonKhoSheet.getDataRange().getValues();
      var stockConfig = getStockSheetConfig(tonKhoData);
      for (var i = stockConfig.startRow; i < tonKhoData.length; i++) {
        if(!tonKhoData[i]) continue;
        var rowStores = getRowStoreNames(tonKhoData[i], stockConfig);
        for (var s = 0; s < rowStores.length; s++) {
          if (stores.indexOf(rowStores[s]) === -1) stores.push(rowStores[s]);
        }
      }
    }

    var dataSheet = ss.getSheetByName("Data_Excel");
    if (dataSheet) {
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
        if(!rawData[k]) continue;
        var ma = getCellValue(rawData[k], maHangIdx !== -1 ? maHangIdx : 0, "").toUpperCase();
        var ten = getCellValue(rawData[k], tenHangIdx !== -1 ? tenHangIdx : 5, "");
        if (ma !== "" && ten !== "") tenHangChuanTheoMa[ma] = ten;
      }
      for (var i = startRow; i < rawData.length; i++) {
        if(!rawData[i]) continue;
        var maHang = getCellValue(rawData[i], maHangIdx !== -1 ? maHangIdx : 0, "");
        var maVach = getCellValue(rawData[i], maVachIdx !== -1 ? maVachIdx : 2, "");
        var tenHang = getCellValue(rawData[i], tenHangIdx !== -1 ? tenHangIdx : 5, "");
        var dvt = getCellValue(rawData[i], dvtIdx !== -1 ? dvtIdx : 7, "");
        if (dvt === "" && dvtIdx === -1) {
          var fallbackDvt = getCellValue(rawData[i], 6, "");
          if (fallbackDvt !== "") dvt = fallbackDvt;
        }
        if (tenHang === "" && maHang !== "") tenHang = tenHangChuanTheoMa[maHang.toUpperCase()] || "";
        var obj = { maHang: maHang, maVach: maVach, tenHang: tenHang, dvt: dvt || "Cái" };
        if (maVach !== "") danhMucHangHoa[maVach.toUpperCase()] = obj;
        if (maHang !== "" && !danhMucHangHoa[maHang.toUpperCase()]) danhMucHangHoa[maHang.toUpperCase()] = obj;
      }
    }
    // Mapping original store names -> display (short) names
    var storeMap = {
      "Kho Địa điểm kinh doanh Q7": "K9 Quận 7",
      "Kho Địa điểm kinh doanh 01": "K9 Quận 4 Mới",
      "Kho Địa điểm kinh doanh 02": "K9 Quận 8",
      "Kho Địa điểm kinh doanh 03": "K9 Phạm Hùng",
      "Kho Địa điểm kinh doanh 04": "K9 Quận 5",
      "Kho Địa điểm kinh doanh 05": "K9 Quận 1",
      "Kho Địa điểm kinh doanh 06": "K9 Quận 4 Cũ"
    };
    // Ensure every store has a mapping (fallback to itself)
    for (var i = 0; i < stores.length; i++) {
      var s = stores[i];
      if (!storeMap[s]) storeMap[s] = s;
    }
    return { success: true, stores: stores, danhMuc: danhMucHangHoa, storeMap: storeMap };
  } catch(e) { return { success: false, error: e.message }; }
}

// --- API: LƯU ĐƠN / PHIẾU TẠO MỚI ---
function luuPhieuTuWebApp(payload) {
  var ss = getSS();
  var homNay = new Date();
  var prefix = payload.loaiPhieu === "DonHang" ? "DH" : "DC";
  var soPhieu = prefix + "-" + Math.floor(100000 + Math.random() * 900000);
  
  var dataLichSuArr = [];
  var coLoiCanDieuChinh = false;
  
  for (var i = 0; i < payload.items.length; i++) {
    var item = payload.items[i];
    var slNum = Number(item.sl);
    var ghiChuLoi = "";
    var coLoiDongNay = false;
    if (isNaN(slNum) || slNum <= 0) { coLoiDongNay = true; ghiChuLoi = "Lỗi số lượng"; slNum = 0; }
    if (!item.dvt || item.dvt === "Không tìm thấy" || item.dvt === "") { coLoiDongNay = true; ghiChuLoi += (ghiChuLoi ? " | " : "") + "Lỗi ĐVT"; }
    if (item.maHang === "LỖI MÃ") { coLoiDongNay = true; ghiChuLoi += (ghiChuLoi ? " | " : "") + "Mã không tồn tại"; }
    dataLichSuArr.push([ homNay, soPhieu, payload.khoXuat, payload.khoNhan, item.maHang, item.maVach, item.tenHang, slNum, "", item.dvt, "", ghiChuLoi, "Mới" ]);
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
  
  if (!coLoiCanDieuChinh) {
    sendTelegramMessage(soPhieu, payload.khoXuat, payload.khoNhan, payload.items.length);
  }
  return { success: true, soPhieu: soPhieu, coLoi: coLoiCanDieuChinh };
}

// --- API: QUẢN LÝ PHIẾU (CÓ PHÂN QUYỀN) ---
function layDanhSachPhieuTheoFilter(khoNhan, ngayYYYYMMDD, userRole, userStore) {
  var ss = getSS();
  var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
  if (!historySheet) return [];
  var data = historySheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var rowNgay = data[i][0]; 
    var rowSoPhieu = data[i][1] ? data[i][1].toString().trim() : "";
    var rowKhoXuat = data[i][2] ? data[i][2].toString().trim() : "";
    var rowKhoNhan = data[i][3] ? data[i][3].toString().trim() : "";
    var slThucTe = data[i][8];
    var rowStatus = data[i][12] ? String(data[i][12]).trim() : "Mới";
    var displayStatus = getDisplayOrderStatus(rowStatus, slThucTe);
    var thoiGian = rowNgay instanceof Date ? rowNgay.getTime() : "";

    if (!rowSoPhieu) continue;
    if (khoNhan && khoNhan !== "all" && rowKhoNhan !== khoNhan) continue;
    
    // BỘ LỌC BẢO MẬT PHÂN QUYỀN: Chi nhánh chỉ thấy đơn của mình
    if (userRole !== "Admin") {
      if (rowKhoXuat !== userStore && rowKhoNhan !== userStore) continue;
    }

    if (ngayYYYYMMDD && rowNgay) {
      var dObj = new Date(rowNgay);
      if (!isNaN(dObj)) {
        var m = dObj.getMonth() + 1; var d = dObj.getDate(); var y = dObj.getFullYear();
        var rowDateStr = y + "-" + (m < 10 ? '0' : '') + m + "-" + (d < 10 ? '0' : '') + d;
        if (rowDateStr !== ngayYYYYMMDD) continue;
      }
    }
    
    if (!map[rowSoPhieu]) { 
      map[rowSoPhieu] = { soPhieu: rowSoPhieu, khoXuat: rowKhoXuat, khoNhan: rowKhoNhan, thoiGian: thoiGian, trangThai: displayStatus === "Đã hủy dòng" ? "Mới" : displayStatus };
    } else {
      if (displayStatus === "Đã hủy") map[rowSoPhieu].trangThai = "Đã hủy";
      else if (displayStatus === "Đã xác nhận" && map[rowSoPhieu].trangThai !== "Đã hủy") map[rowSoPhieu].trangThai = "Đã xác nhận";
      else if (displayStatus === "Đã soạn" && map[rowSoPhieu].trangThai !== "Đã hủy" && map[rowSoPhieu].trangThai !== "Đã xác nhận") map[rowSoPhieu].trangThai = "Đã soạn";
    }
  }
  var res = []; for(var key in map) res.push(map[key]); 
  res.sort(function(a,b){ return b.thoiGian - a.thoiGian; }); 
  return res; 
}

function getDashboardSummary(userRole, userStore, timeline, fromDate, toDate) {
  try {
    var ss = getSS();
    var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
    if (!historySheet) {
      return { success: true, data: { totalOrders: 0, pendingOrders: 0, processedOrders: 0, canceledOrders: 0, recentOrders: [] } };
    }

    var data = historySheet.getDataRange().getValues();
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
      var displayStatus = getDisplayOrderStatus(status, slThucTe);
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
  var date = value instanceof Date ? new Date(value) : new Date(value);
  if (isNaN(date.getTime())) return false;
  date.setHours(0, 0, 0, 0);

  var today = new Date();
  today.setHours(0, 0, 0, 0);
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
  var start = new Date(today);
  start.setDate(today.getDate() - (days - 1));
  return date.getTime() >= start.getTime() && date.getTime() <= today.getTime();
}

function getChiTietPhieu(soPhieu, storeName) {
  var ss = getSS();
  var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
  var data = historySheet.getDataRange().getValues();
  var matchedRows = [];
  for (var i = 1; i < data.length; i++) { 
    if (data[i][1] && data[i][1].toString().toLowerCase() === soPhieu.toLowerCase()) {
      var slGoc = Number(data[i][7]) || 0;
      var slThucTe = (data[i][8] !== "" && data[i][8] !== undefined) ? Number(data[i][8]) : slGoc;
      var effectiveQty = (slThucTe !== undefined && slThucTe !== null && slThucTe !== "") ? slThucTe : slGoc;
      matchedRows.push({ rowIndex: i + 1, maHang: data[i][4], maVach: data[i][5], tenHang: data[i][6], slGoc: slGoc, slThucTe: effectiveQty, dvt: data[i][9], ghiChu: data[i][11]||"", trangThai: data[i][12] || "Mới", nguoiSoanHang: data[i][13] || "" });
    }
  }
  var tonKhoSheet = ss.getSheetByName("TỔNG HỢP TỒN KHO");
  var tonKhoMap = {};
  if(tonKhoSheet) {
    var tkData = tonKhoSheet.getDataRange().getValues();
    var stockConfig = getStockSheetConfig(tkData);
    for (var k = stockConfig.startRow; k < tkData.length; k++) {
      var rowStores = getRowStoreNames(tkData[k], stockConfig);
      var match = false;
      for (var s = 0; s < rowStores.length; s++) {
        if (rowStores[s].indexOf(storeName) !== -1 || storeName.indexOf(rowStores[s]) !== -1) {
          match = true;
          break;
        }
      }
      var maHangTon = getCellValue(tkData[k], stockConfig.maHangIdx, "");
      if(match && maHangTon) tonKhoMap[maHangTon.toUpperCase()] = Number(tkData[k][stockConfig.tonKhoIdx])||0;
    }
  }
  for (var j = 0; j < matchedRows.length; j++) { matchedRows[j].stock = tonKhoMap[matchedRows[j].maHang.toUpperCase()] || 0; }
  return matchedRows;
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
}

function getDisplayOrderStatus(rowStatus, slThucTe) {
  var status = String(rowStatus || "").trim();
  if (status === "Đã hủy đơn") return "Đã hủy";
  if (status === "Đã xác nhận nhận hàng") return "Đã xác nhận";
  if (status === "Đã soạn hàng") return "Đã soạn";
  if (status === "Đã hủy dòng") return "Đã hủy dòng";
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
    if (rowStatus === "Đã xác nhận nhận hàng") state.isConfirmed = true;
    if (rowStatus === "Đã soạn hàng" || ((slThucTe !== "" && slThucTe !== undefined && slThucTe !== null) && rowStatus !== "Đã hủy dòng" && rowStatus !== "Đã hủy đơn")) {
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
      var maHang = currentRow[4] ? String(currentRow[4]).trim() : "";
      var maVach = currentRow[5] ? String(currentRow[5]).trim() : "";
      if (!orderInfo) orderInfo = getThongTinPhieu(soPhieuValue);
      if (u.valSl !== "" && Number(u.valSl) === 0) {
        pendingUpdates.push({row: rowIndex, requestedQty: 0, actualQty: 0, note: "Đã hủy dòng", status: "Đã hủy dòng", actor: payload.actor || "", source: "Quản lý"});
        logOrderChange(ss, soPhieuValue, "Hủy mã khỏi đơn", payload.actor, maHang, maVach, oldSl, 0, "Hủy bằng cập nhật số lượng");
        changeCount += 1;
        cancelledCount += 1;
        shouldNotify = true;
      } else if (u.valSl !== "") {
        var newVal = Number(u.valSl);
        pendingUpdates.push({row: rowIndex, requestedQty: newVal, actualQty: wasPacked ? "" : currentRow[8], note: wasPacked ? "Cần soạn lại sau khi chỉnh sửa" : "", status: "Mới", actor: payload.actor || "", source: "Quản lý"});
        if (Number(oldSl) !== newVal) {
          logOrderChange(ss, soPhieuValue, "Sửa số lượng", payload.actor, maHang, maVach, oldSl, newVal, "");
          changeCount += 1;
          modifiedCount += 1;
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
      sendTelegramOrderChangeSummary(orderInfo.soPhieu, orderInfo.khoXuat, orderInfo.khoNhan, actionLabel, changeCount, payload.actor, extraSummary);
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
  var orderInfo = getThongTinPhieu(payload.soPhieu);
  if (orderInfo) {
    sendTelegramReceiveConfirmation(payload.soPhieu, orderInfo.khoNhan || expectedStore, payload.actor, confirmations.length, confirmedTotal, changedCount, changedQtyTotal);
  }
  return { success: true, count: confirmations.length };
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
    ss.deleteSheet(existing);
  }
  return ss.insertSheet(sheetName);
}

function taoFileExcelVaLayLink(payload) {
  var ss = getSS();
  var tenTabPhieu = "__TMP_XUAT_EXCEL";
  var targetSheet = recreateTempSheet(ss, tenTabPhieu, ["In_"]);
  
  var khoXuat = payload.khoXuat; var khoNhan = payload.khoNhan;
  var tieuDe = payload.soPhieu.indexOf("DH") !== -1 ? "ĐƠN ĐẶT HÀNG" : "LỆNH ĐIỀU CHUYỂN";
  
  targetSheet.getRange("A4:F4").merge().setValue(tieuDe).setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
  targetSheet.getRange("A6:F6").merge().setValue("Số: " + payload.soPhieu).setFontStyle("italic").setHorizontalAlignment("center");
  targetSheet.getRange("A8").setValue("Kho xuất:").setFontWeight("bold"); targetSheet.getRange("B8").setValue(khoXuat);
  targetSheet.getRange("A9").setValue("Kho nhận:").setFontWeight("bold"); targetSheet.getRange("B9").setValue(khoNhan);
  
  var headers = ["STT", "Mã hàng hóa", "Mã vạch", "Tên hàng hóa", "ĐVT", "Số lượng"];
  targetSheet.getRange("A12:F12").setValues([headers]).setFontWeight("bold").setHorizontalAlignment("center").setBackground("#f8f9fa");
  
  var dataArr = []; var stt = 1;
  for(var i=0; i<payload.items.length; i++) {
    if(Number(payload.items[i].sl) > 0) dataArr.push([stt++, payload.items[i].maHang, payload.items[i].maVach, payload.items[i].tenHang, payload.items[i].dvt, payload.items[i].sl]);
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
  var data = historySheet.getDataRange().getValues();
  var map = {};
  var today = new Date(); today.setHours(0,0,0,0);
  var yesterday = new Date(today); yesterday.setDate(today.getDate() - 1);
  
  for (var i = 1; i < data.length; i++) {
    var rowNgay = new Date(data[i][0]); rowNgay.setHours(0,0,0,0);
    var rowSoPhieu = data[i][1] ? data[i][1].toString().trim() : "";
    var rowKhoXuat = data[i][2] ? data[i][2].toString().trim() : "";
    var rowKhoNhan = data[i][3] ? data[i][3].toString().trim() : "";
    var rowStatus = data[i][12] ? String(data[i][12]).trim() : "Mới";
    var isCanceled = rowStatus === "Đã hủy đơn";

    if (isCanceled) continue;

    if (viewMode === "packing") {
      if (userRole !== "Admin" && rowKhoXuat !== userStore) continue;
    } else if (userRole !== "Admin") {
      if (rowKhoXuat !== userStore && rowKhoNhan !== userStore) continue;
    }

    var slThucTe = data[i][8];
    var displayStatus = getDisplayOrderStatus(rowStatus, slThucTe);
    
    var match = (ngayChon === 'today' && rowNgay.getTime() === today.getTime()) ||
                (ngayChon === 'yesterday' && rowNgay.getTime() === yesterday.getTime()) ||
                (ngayChon === 'all');
                
    if (match && rowSoPhieu) { 
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
  var data = historySheet.getDataRange().getValues();
  var items = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][1].toString().trim().toLowerCase() === soPhieu.toLowerCase()) {
      var slGoc = Number(data[i][7]) || 0;
      var slThucTe = (data[i][8] !== "" && data[i][8] !== undefined) ? Number(data[i][8]) : slGoc;
      items.push({ rowIndex: i + 1, maHang: data[i][4], maVach: data[i][5], tenHang: data[i][6], dvt: data[i][9] || "Cái", slGoc: slGoc, slThucTe: slThucTe, anhXacNhan: (data[i][10]||""), nguoiSoanHang: data[i][13] || "" });
    }
  }
  return items;
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
    for (var i = 0; i < updates.length; i++) {
      var val = updates[i].val;
      if (val !== "") {
        var parsedVal = Number(val);
        historySheet.getRange(updates[i].row, 9).setValue(parsedVal);
        historySheet.getRange(updates[i].row, 14).setValue(payload.actor || "");
        historySheet.getRange(updates[i].row, 15).setValue("Soạn hàng");
        historySheet.getRange(updates[i].row, 13).setValue("Đã soạn hàng");
      } else {
        historySheet.getRange(updates[i].row, 9).clearContent();
        historySheet.getRange(updates[i].row, 13).setValue("Mới");
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
    for (var i = 0; i < updates.length; i++) {
      var row = updates[i].row;
      var currentSoPhieu = historySheet.getRange(row, 2).getValue();
      if (currentSoPhieu) { soPhieu = String(currentSoPhieu).trim(); break; }
    }
    var allRows = historySheet.getDataRange().getValues();
    for (var i = 1; i < allRows.length; i++) {
      var rowSoPhieu = allRows[i][1] ? String(allRows[i][1]).trim() : "";
      if (!rowSoPhieu || rowSoPhieu !== soPhieu) continue;
      totalRows += 1;
      var requestedQty = Number(allRows[i][7]) || 0;
      var actualQty = (allRows[i][8] !== "" && allRows[i][8] !== undefined) ? Number(allRows[i][8]) : requestedQty;
      if (actualQty < requestedQty) missingCount += 1;
      if (actualQty > requestedQty) extraCount += 1;
      if (!khoXuat && allRows[i][2]) khoXuat = String(allRows[i][2]).trim();
      if (!khoNhan && allRows[i][3]) khoNhan = String(allRows[i][3]).trim();
    }
    if (soPhieu) {
      if (!khoNhan) khoNhan = getKhoNhanBySoPhieu(soPhieu);
      if (khoNhan) {
        sendTelegramOrderReady(soPhieu, khoNhan);
      }
      var statusLabel = (totalRows > 0 && missingCount === 0 && extraCount === 0) ? "Đủ hàng" : (extraCount > 0 && missingCount > 0 ? "Thiếu và thừa hàng" : (extraCount > 0 ? "Thừa hàng" : "Thiếu hàng"));
      sendTelegramPackingSummary(soPhieu, khoXuat, khoNhan, updates.length, totalRows, statusLabel, missingCount, extraCount, payload.actor || "Chi nhánh");
    }

    return "✅ Đã lưu " + updates.length + " món và " + anhDaLuu + " ảnh!";
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

function formatShortStoreLabel(storeName) {
  var name = String(STORE_MAP[storeName] || storeName || "").trim();
  var normalized = normalizeHeaderText(name);
  if (normalized.indexOf("q7") !== -1 || normalized.indexOf("quan7") !== -1) return "Q7";
  if (normalized.indexOf("q8") !== -1 || normalized.indexOf("quan8") !== -1) return "Q8";
  if (normalized.indexOf("phamhung") !== -1) return "PH";
  if (normalized.indexOf("q5") !== -1 || normalized.indexOf("quan5") !== -1) return "Q5";
  if (normalized.indexOf("q1") !== -1 || normalized.indexOf("quan1") !== -1) return "Q1";
  if (normalized.indexOf("q4") !== -1 || normalized.indexOf("quan4") !== -1) return "Q4";
  return name || "Khác";
}

function getStockIndexByStore(stockData) {
  var index = {};
  var stockConfig = getStockSheetConfig(stockData);
  for (var i = stockConfig.startRow; i < stockData.length; i++) {
    var row = stockData[i];
    if (!row) continue;
    var stores = getRowStoreNames(row, stockConfig);
    var maHang = getCellValue(row, stockConfig.maHangIdx, "").toUpperCase();
    var maVach = getCellValue(row, stockConfig.maVachIdx, "").toUpperCase();
    var qty = Number(row[stockConfig.tonKhoIdx]) || 0;
    if (qty === 0) continue;
    if (!stores.length) continue;

    for (var s = 0; s < stores.length; s++) {
      var store = stores[s];
      if (!index[store]) index[store] = {};
      if (maHang) index[store][maHang] = (index[store][maHang] || 0) + qty;
      if (maVach) index[store][maVach] = (index[store][maVach] || 0) + qty;
    }
  }
  return index;
}

function taoBangSoanHangNgayMai(payload) {
  var ss = getSS();
  var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
  if (!historySheet) throw new Error("Không tìm thấy sheet Lịch Sử Xuất Kho.");

  var inputDate = parseDateInputYYYYMMDD(payload && payload.ngay ? payload.ngay : "");
  var baseDate = inputDate || new Date();
  baseDate.setHours(0, 0, 0, 0);
  var tomorrow = new Date(baseDate);
  tomorrow.setDate(tomorrow.getDate() + 1);

  var data = historySheet.getDataRange().getValues();
  if (!data || data.length < 2) throw new Error("Chưa có dữ liệu đơn hàng để tổng hợp.");

  var itemMap = {};
  var storeList = [];
  var orderSeen = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row) continue;
    var ngayTao = row[0] instanceof Date ? new Date(row[0]) : new Date(row[0]);
    if (isNaN(ngayTao.getTime())) continue;
    ngayTao.setHours(0, 0, 0, 0);
    if (ngayTao.getTime() !== baseDate.getTime()) continue;

    var soPhieu = row[1] ? String(row[1]).trim() : "";
    var khoXuat = row[2] ? String(row[2]).trim() : "";
    var khoNhan = row[3] ? String(row[3]).trim() : "";
    var maHang = row[4] ? String(row[4]).trim() : "";
    var maVach = row[5] ? String(row[5]).trim() : "";
    var tenHang = row[6] ? String(row[6]).trim() : "";
    var soLuong = Number(row[7]) || 0;
    var status = row[12] ? String(row[12]).trim() : "Đang xử lý";

    if (!soPhieu || !khoNhan) continue;
    if (status === "Đã hủy đơn" || status === "Đã hủy dòng") continue;
    if (soLuong <= 0) continue;

    var itemKey = (maHang ? maHang.toUpperCase() : "") || (maVach ? maVach.toUpperCase() : "");
    if (!itemKey) continue;

    if (!itemMap[itemKey]) {
      itemMap[itemKey] = {
        maHang: maHang,
        maVach: maVach,
        tenHang: tenHang,
        totalQty: 0,
        byStore: {},
        sourceStores: {},
        orders: {}
      };
    }

    var item = itemMap[itemKey];
    if (!item.tenHang && tenHang) item.tenHang = tenHang;
    if (!item.maHang && maHang) item.maHang = maHang;
    if (!item.maVach && maVach) item.maVach = maVach;
    item.totalQty += soLuong;
    item.byStore[khoNhan] = (item.byStore[khoNhan] || 0) + soLuong;
    if (khoXuat) item.sourceStores[khoXuat] = true;
    item.orders[soPhieu] = true;

    if (storeList.indexOf(khoNhan) === -1) storeList.push(khoNhan);
    orderSeen[soPhieu] = true;
  }

  var keys = [];
  for (var key in itemMap) keys.push(key);
  if (!keys.length) {
    return { success: false, msg: "Không có đơn hợp lệ trong ngày đã chọn để tạo bảng soạn." };
  }

  storeList.sort(function(a, b) {
    return formatShortStoreLabel(a).localeCompare(formatShortStoreLabel(b));
  });

  var stockSheet = ss.getSheetByName("TỔNG HỢP TỒN KHO");
  var stockIndex = stockSheet ? getStockIndexByStore(stockSheet.getDataRange().getValues()) : {};

  var sheetName = "__TMP_SOAN_NGAY_MAI";
  var reportSheet = recreateTempSheet(ss, sheetName, ["SoanNgayMai_"]);

  var title = "BẢNG TỔNG HỢP SOẠN HÀNG NGÀY " + Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), "dd/MM/yyyy");
  reportSheet.getRange("A1").setValue(title).setFontSize(14).setFontWeight("bold");
  reportSheet.getRange("A2").setValue("Nguồn dữ liệu: Đơn tạo ngày " + Utilities.formatDate(baseDate, Session.getScriptTimeZone(), "dd/MM/yyyy") + " | Gom theo mã để 1 lượt lấy đủ cho tất cả đơn.").setFontStyle("italic");
  reportSheet.getRange("A3").setValue("Quy trình gợi ý: 1) In tab này 2) Đi theo từng mã từ trên xuống 3) Lấy đủ theo cột Tổng đặt 4) Chia theo cột Q7/Q8/PH... 5) Mã thiếu xử lý theo cột Đề xuất.");

  var headers = ["STT", "Mã hàng", "Mã vạch", "Tên hàng", "Stock khả dụng", "Tổng đặt"];
  for (var h = 0; h < storeList.length; h++) {
    headers.push(formatShortStoreLabel(storeList[h]));
  }
  headers.push("Cảnh báo");
  headers.push("Đề xuất xử lý thiếu");
  headers.push("Kho xuất");

  var headerRow = 5;
  reportSheet.getRange(headerRow, 1, 1, headers.length).setValues([headers]);
  reportSheet.getRange(headerRow, 1, 1, headers.length).setFontWeight("bold").setBackground("#d9ead3").setHorizontalAlignment("center");
  reportSheet.setFrozenRows(headerRow);

  var rows = [];
  var missingLines = 0;
  for (var k = 0; k < keys.length; k++) {
    var item = itemMap[keys[k]];
    var sourceStores = [];
    for (var src in item.sourceStores) sourceStores.push(src);

    var stock = 0;
    for (var s = 0; s < sourceStores.length; s++) {
      var store = sourceStores[s];
      var storeStockMap = stockIndex[store] || {};
      var codeA = item.maHang ? item.maHang.toUpperCase() : "";
      var codeB = item.maVach ? item.maVach.toUpperCase() : "";
      var qtyByMaHang = codeA ? (storeStockMap[codeA] || 0) : 0;
      var qtyByMaVach = codeB ? (storeStockMap[codeB] || 0) : 0;
      stock += qtyByMaHang > 0 ? qtyByMaHang : qtyByMaVach;
    }

    var thieu = item.totalQty - stock;
    var canhBao = thieu > 0 ? ("THIẾU " + thieu) : "ĐỦ";
    var deXuat = thieu > 0
      ? ("Thiếu " + thieu + ": ưu tiên đơn gấp, điều chuyển nội bộ hoặc nhập bổ sung trước khi in phiếu giao.")
      : "OK - có thể soạn gộp theo mã.";
    if (thieu > 0) missingLines += 1;

    var rowOut = [
      0,
      item.maHang || "",
      item.maVach || "",
      item.tenHang || "",
      stock,
      item.totalQty
    ];

    for (var c = 0; c < storeList.length; c++) {
      rowOut.push(item.byStore[storeList[c]] || 0);
    }

    rowOut.push(canhBao);
    rowOut.push(deXuat);
    rowOut.push(sourceStores.map(function(name) { return STORE_MAP[name] || name; }).join(", "));
    rows.push(rowOut);
  }

  rows.sort(function(a, b) {
    var aWarn = String(a[6 + storeList.length]);
    var bWarn = String(b[6 + storeList.length]);
    if (aWarn !== bWarn) return aWarn.indexOf("THIẾU") === 0 ? -1 : 1;
    var aName = String(a[3] || "");
    var bName = String(b[3] || "");
    return aName.localeCompare(bName);
  });
  for (var r = 0; r < rows.length; r++) rows[r][0] = r + 1;

  reportSheet.getRange(4, 1).setValue("Tổng đơn: " + Object.keys(orderSeen).length + " | Tổng mã: " + rows.length + " | Mã thiếu: " + missingLines).setFontWeight("bold");

  var startRow = headerRow + 1;
  reportSheet.getRange(startRow, 1, rows.length, headers.length).setValues(rows);
  reportSheet.getRange(startRow, 5, rows.length, 2 + storeList.length).setHorizontalAlignment("right");
  reportSheet.getRange(startRow, 1, rows.length, headers.length).setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);

  var warningCol = 7 + storeList.length;
  for (var x = 0; x < rows.length; x++) {
    if (String(rows[x][warningCol - 1]).indexOf("THIẾU") === 0) {
      reportSheet.getRange(startRow + x, warningCol).setBackground("#f4cccc").setFontWeight("bold");
      reportSheet.getRange(startRow + x, 5).setBackground("#fff2cc");
      reportSheet.getRange(startRow + x, 6).setBackground("#fce5cd");
    } else {
      reportSheet.getRange(startRow + x, warningCol).setBackground("#d9ead3");
    }
  }

  reportSheet.autoResizeColumns(1, headers.length);
  reportSheet.setColumnWidth(4, 260);
  reportSheet.setColumnWidth(8 + storeList.length, 360);
  SpreadsheetApp.flush();

  return {
    success: true,
    sheetName: sheetName,
    totalOrders: Object.keys(orderSeen).length,
    totalItems: rows.length,
    missingItems: missingLines,
    url: "https://docs.google.com/spreadsheets/d/" + SHEET_ID + "/edit#gid=" + reportSheet.getSheetId()
  };
}