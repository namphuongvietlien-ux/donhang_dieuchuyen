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
    if (!soPhieu || !khoNhan) continue;
    if (khoNhan !== storeName) continue;
    var isDaXuLy = (slThucTe !== "" && slThucTe !== undefined && Number(slThucTe) > 0);
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

function getOrderWebUrl(soPhieu, tabName) {
  var tab = tabName || "quan-ly";
  return WEB_APP_URL + "?tab=" + encodeURIComponent(tab) + "&soPhieu=" + encodeURIComponent(soPhieu);
}

function sendTelegramMessage(soPhieu, khoXuat, khoNhan, itemCount) {
  var typeLabel = soPhieu.indexOf("DH") !== -1 ? "ĐƠN HÀNG MỚI" : "LỆNH ĐIỀU CHUYỂN MỚI";
  var kxShort = STORE_MAP[khoXuat] || khoXuat;
  var knShort = STORE_MAP[khoNhan] || khoNhan;
  var text = "📦 *THÔNG BÁO " + typeLabel + "*\n" +
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

function sendTelegramPackingSummary(soPhieu, khoXuat, khoNhan, changedCount, totalRows, statusLabel, missingCount, actor) {
  var kxShort = STORE_MAP[khoXuat] || khoXuat;
  var knShort = STORE_MAP[khoNhan] || khoNhan;
  var text = "📦 *ĐƠN ĐÃ SOẠN XONG*\n" +
             "*Số phiếu:* " + soPhieu + "\n" +
             "*Kho xuất:* " + khoXuat + " (" + kxShort + ")\n" +
             "*Kho nhận:* " + khoNhan + " (" + knShort + ")\n" +
             "*Mã thay đổi:* " + changedCount + "\n" +
             "*Kết quả:* " + statusLabel + "\n" +
             "*Thiếu hàng:* " + missingCount + " mã / tổng " + totalRows + " mã\n" +
             "*Người thực hiện:* " + (actor || "Không xác định") + "\n\n" +
             "Mở chi tiết đơn: " + getOrderWebUrl(soPhieu);
  if (TELEGRAM_CHAT_ID) {
    sendTelegramText(TELEGRAM_CHAT_ID, text);
  }
  if (khoNhan) {
    sendTelegramTextToStoreUsers(khoNhan, "✅ Đơn đã được soạn xong:\n" + text);
  }
}

function sendTelegramReceiveConfirmation(soPhieu, khoNhan, actor, count, confirmedTotal) {
  var knShort = STORE_MAP[khoNhan] || khoNhan;
  var text = "📥 *XÁC NHẬN NHẬN HÀNG*\n" +
             "*Số phiếu:* " + soPhieu + "\n" +
             "*Kho nhận:* " + khoNhan + " (" + knShort + ")\n" +
             "*Số dòng xác nhận:* " + count + "\n" +
             "*Người xác nhận:* " + (actor || "Không xác định") + "\n" +
             "*Tổng số lượng đã xác nhận:* " + confirmedTotal + "\n\n" +
             "Mở chi tiết đơn: " + getOrderWebUrl(soPhieu, "xac-nhan");
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

// --- API: LẤY DATA BAN ĐẦU ---
function getInitialData() {
  try {
    var ss = getSS();
    var danhMucHangHoa = {};
    var stores = ["Kho Địa điểm kinh doanh Q7"];
    
    var tonKhoSheet = ss.getSheetByName("TỔNG HỢP TỒN KHO");
    if (tonKhoSheet) {
      var tonKhoData = tonKhoSheet.getDataRange().getValues();
      for (var i = 4; i < tonKhoData.length; i++) {
        if(!tonKhoData[i]) continue;
        var tenDong = tonKhoData[i][0] ? tonKhoData[i][0].toString().trim() : "";
        var cuaHang = tonKhoData[i][7] ? tonKhoData[i][7].toString().trim() : "";
        if (tenDong.indexOf("Kho ") !== -1 && stores.indexOf(tenDong) === -1) stores.push(tenDong);
        else if (cuaHang !== "" && stores.indexOf(cuaHang) === -1) stores.push(cuaHang);
      }
    }

    var dataSheet = ss.getSheetByName("Data_Excel");
    if (dataSheet) {
      var rawData = dataSheet.getDataRange().getValues();
      var tenHangChuanTheoMa = {};
      for (var k = 2; k < rawData.length; k++) {
        if(!rawData[k]) continue;
        var ma = rawData[k][0] ? rawData[k][0].toString().trim().toUpperCase() : "";
        var ten = rawData[k][5] ? rawData[k][5].toString().trim() : "";
        if (ma !== "" && ten !== "") tenHangChuanTheoMa[ma] = ten;
      }
      for (var i = 2; i < rawData.length; i++) {
        if(!rawData[i]) continue;
        var maHang = rawData[i][0] ? rawData[i][0].toString().trim() : "";
        var maVach = rawData[i][2] ? rawData[i][2].toString().trim() : "";
        var tenHang = rawData[i][5] ? rawData[i][5].toString().trim() : "";
        var dvt = rawData[i][7] ? rawData[i][7].toString().trim() : "Cái";
        if (tenHang === "" && maHang !== "") tenHang = tenHangChuanTheoMa[maHang.toUpperCase()] || "";
        var obj = { maHang: maHang, maVach: maVach, tenHang: tenHang, dvt: dvt };
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
    dataLichSuArr.push([ homNay, soPhieu, payload.khoXuat, payload.khoNhan, item.maHang, item.maVach, item.tenHang, slNum, "", item.dvt, "", ghiChuLoi, "Đang xử lý" ]);
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
    var rowStatus = data[i][12] ? String(data[i][12]).trim() : "Đang xử lý";
    var isDaXuLy = rowStatus === "Đã hủy đơn" || (slThucTe !== "" && Number(slThucTe) > 0);
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
      map[rowSoPhieu] = { soPhieu: rowSoPhieu, khoXuat: rowKhoXuat, khoNhan: rowKhoNhan, thoiGian: thoiGian, trangThai: rowStatus === "Đã hủy đơn" ? "Đã hủy" : (isDaXuLy ? "Đã xử lý" : "Mới") };
    } else {
      if (rowStatus === "Đã hủy đơn") map[rowSoPhieu].trangThai = "Đã hủy";
      else if (isDaXuLy && map[rowSoPhieu].trangThai !== "Đã hủy") map[rowSoPhieu].trangThai = "Đã xử lý";
    }
  }
  var res = []; for(var key in map) res.push(map[key]); 
  res.sort(function(a,b){ return b.thoiGian - a.thoiGian; }); 
  return res; 
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
      matchedRows.push({ rowIndex: i + 1, maHang: data[i][4], maVach: data[i][5], tenHang: data[i][6], slGoc: slGoc, slThucTe: effectiveQty, dvt: data[i][9], ghiChu: data[i][11]||"", trangThai: data[i][12] || "Đang xử lý", nguoiSoanHang: data[i][13] || "" });
    }
  }
  var tonKhoSheet = ss.getSheetByName("TỔNG HỢP TỒN KHO");
  var tonKhoMap = {};
  if(tonKhoSheet) {
    var tkData = tonKhoSheet.getDataRange().getValues();
    for (var k = 4; k < tkData.length; k++) {
      var td = tkData[k][0]?tkData[k][0].toString():""; var ch = tkData[k][7]?tkData[k][7].toString():"";
      var match = (td.indexOf(storeName)!==-1) || (ch.indexOf(storeName)!==-1);
      if(match && tkData[k][1]) tonKhoMap[tkData[k][1].toString().toUpperCase()] = Number(tkData[k][6])||0;
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
  var adminActions = ['luuChinhSuaPhieu', 'taoTaiKhoanMoi'];
  if (adminActions.indexOf(action) !== -1 && !isAdminActor(payload && payload.actor ? payload.actor : "")) {
    throw new Error("Chỉ quản trị viên được phép thực hiện thao tác này.");
  }
}

function requireAdmin(actor) {
  var account = getAccountByActor(actor);
  if (!account || String(account.role).trim() !== "Admin") throw new Error("Chỉ quản trị viên được phép thay đổi hoặc hủy đơn.");
}

function luuChinhSuaPhieu(payload) {
  var ss = getSS();
  var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
  requireAdmin(payload.actor);
  ensureHistoryStatusColumn(historySheet);
  var lock = LockService.getDocumentLock();
  try {
    lock.waitLock(10000);
    var changeCount = 0;
    var shouldNotify = false;
    var orderInfo = null;
    for (var i = 0; i < payload.updates.length; i++) {
      var u = payload.updates[i];
      var oldSl = historySheet.getRange(u.row, 8).getValue();
      var soPhieu = historySheet.getRange(u.row, 2).getValue();
      var maHang = historySheet.getRange(u.row, 5).getValue();
      var maVach = historySheet.getRange(u.row, 6).getValue();
      if (!orderInfo) orderInfo = getThongTinPhieu(soPhieu);
      if (Number(u.valSl) === 0) {
         historySheet.getRange(u.row, 8).setValue(0);
         historySheet.getRange(u.row, 9).setValue(0);
         historySheet.getRange(u.row, 12).setValue("Đã hủy dòng");
         historySheet.getRange(u.row, 13).setValue("Đã hủy dòng");
         historySheet.getRange(u.row, 14).setValue(payload.actor || "");
         historySheet.getRange(u.row, 15).setValue("Quản lý");
         logOrderChange(ss, soPhieu, "Hủy mã khỏi đơn", payload.actor, maHang, maVach, oldSl, 0, "Hủy bằng cập nhật số lượng");
         changeCount += 1;
         shouldNotify = true;
      } else {
         if (u.valSl !== "") {
           var newVal = Number(u.valSl);
           historySheet.getRange(u.row, 8).setValue(newVal);
           historySheet.getRange(u.row, 9).setValue(newVal);
           historySheet.getRange(u.row, 12).clearContent();
           historySheet.getRange(u.row, 13).setValue("Đang xử lý");
           historySheet.getRange(u.row, 14).setValue(payload.actor || "");
           historySheet.getRange(u.row, 15).setValue("Quản lý");
           if (Number(oldSl) !== newVal) {
             logOrderChange(ss, soPhieu, "Sửa số lượng", payload.actor, maHang, maVach, oldSl, newVal, "");
             changeCount += 1;
             shouldNotify = true;
           }
         }
      }
    }
    SpreadsheetApp.flush();
    if (orderInfo && shouldNotify && changeCount > 0) {
      sendTelegramOrderChangeSummary(orderInfo.soPhieu, orderInfo.khoXuat, orderInfo.khoNhan, "Chỉnh sửa số lượng / hủy mã", changeCount, payload.actor, "Đơn đã thay đổi " + changeCount + " mã.");
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
  var baseRow = null;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(payload.soPhieu).trim()) { baseRow = data[i]; break; }
  }
  if (!baseRow) throw new Error("Không tìm thấy đơn hàng.");
  var item = payload.item;
  var quantity = Number(item.sl);
  if (!quantity || quantity < 1) throw new Error("Số lượng thêm phải lớn hơn 0.");
  var row = [new Date(), payload.soPhieu, baseRow[2], baseRow[3], item.maHang || "", item.maVach || "", item.tenHang || "", quantity, "", item.dvt || "", "", "Thêm bởi quản trị viên", "Đang xử lý"];
  historySheet.appendRow(row);
  logOrderChange(ss, payload.soPhieu, "Thêm mã vào đơn", payload.actor, item.maHang, item.maVach, "", quantity, item.tenHang || "");
  var orderInfo = getThongTinPhieu(payload.soPhieu);
  if (orderInfo) {
    sendTelegramOrderChangeSummary(payload.soPhieu, orderInfo.khoXuat, orderInfo.khoNhan, "Thêm mã vào đơn", 1, payload.actor, "Đơn đã thêm 1 mã mới.");
  }
  return { success: true };
}

function huyDongChiTietPhieu(payload) {
  requireAuthenticatedAction(payload);
  var ss = getSS();
  var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
  ensureHistoryStatusColumn(historySheet);
  var row = Number(payload.row);
  if (!row || row < 2) throw new Error("Dòng đơn hàng không hợp lệ.");
  var values = historySheet.getRange(row, 1, 1, 13).getValues()[0];
  historySheet.getRange(row, 8).setValue(0);
  historySheet.getRange(row, 12).setValue("Đã hủy dòng");
  historySheet.getRange(row, 13).setValue("Đã hủy dòng");
  logOrderChange(ss, values[1], "Hủy mã khỏi đơn", payload.actor, values[4], values[5], values[7], 0, "Hủy từng dòng");
  var orderInfo = getThongTinPhieu(values[1]);
  if (orderInfo) {
    sendTelegramOrderChangeSummary(values[1], orderInfo.khoXuat, orderInfo.khoNhan, "Hủy mã khỏi đơn", 1, payload.actor, "Đơn đã hủy 1 mã khỏi danh sách.");
  }
  return { success: true };
}

function huyPhieu(payload) {
  requireAdmin(payload.actor);
  var ss = getSS();
  var historySheet = ss.getSheetByName("Lịch Sử Xuất Kho");
  ensureHistoryStatusColumn(historySheet);
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
  for (var i = 0; i < confirmations.length; i++) {
    var conf = confirmations[i];
    var row = Number(conf.row);
    var qty = Number(conf.receivedQty);
    if (!row || row < 2 || isNaN(qty) || qty < 0) continue;
    var values = historySheet.getRange(row, 1, 1, 13).getValues()[0];
    receiveSheet.appendRow([new Date(), payload.soPhieu, expectedStore, payload.actor, values[4] || "", values[5] || "", values[6] || "", qty, values[7] || 0, "Đã xác nhận nhận hàng"]);
    logOrderChange(ss, payload.soPhieu, "Xác nhận nhận hàng", payload.actor, values[4], values[5], values[7], qty, "Xác nhận bởi chi nhánh");
    confirmedTotal += qty;
  }
  var orderInfo = getThongTinPhieu(payload.soPhieu);
  if (orderInfo) {
    sendTelegramReceiveConfirmation(payload.soPhieu, orderInfo.khoNhan || expectedStore, payload.actor, confirmations.length, confirmedTotal);
  }
  return { success: true, count: confirmations.length };
}

function taoFileExcelVaLayLink(payload) {
  var ss = getSS();
  var tenTabPhieu = "In_" + payload.soPhieu;
  var targetSheet = ss.getSheetByName(tenTabPhieu);
  if (targetSheet) { ss.deleteSheet(targetSheet); }
  targetSheet = ss.insertSheet(tenTabPhieu);
  
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

// --- API: SOẠN HÀNG MOBILE ---
function getDonHangTheoNgay(ngayChon, userRole, userStore) {
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
    var rowKhoNhan = data[i][3] || "N/A";
    
    if (userRole !== "Admin") {
      if (rowKhoXuat !== userStore && rowKhoNhan !== userStore) continue;
    }

    var slThucTe = data[i][8];
    var isDaXuLy = (slThucTe !== "" && Number(slThucTe) > 0);
    
    var match = (ngayChon === 'today' && rowNgay.getTime() === today.getTime()) ||
                (ngayChon === 'yesterday' && rowNgay.getTime() === yesterday.getTime()) ||
                (ngayChon === 'all');
                
    if (match && rowSoPhieu) { 
      if(!map[rowSoPhieu]) map[rowSoPhieu] = { soPhieu: rowSoPhieu, khoNhan: rowKhoNhan, trangThai: isDaXuLy ? "Đã xử lý" : "Mới" }; 
      else if(isDaXuLy) map[rowSoPhieu].trangThai = "Đã xử lý";
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
  var updates = payload.updates || [];
  var images = payload.images || {}; 
  var lock = LockService.getDocumentLock();
  try {
    lock.waitLock(20000); 
    for (var i = 0; i < updates.length; i++) {
      var val = updates[i].val;
      if (val !== "") {
        var parsedVal = Number(val);
        historySheet.getRange(updates[i].row, 8).setValue(parsedVal);
        historySheet.getRange(updates[i].row, 9).setValue(parsedVal);
        historySheet.getRange(updates[i].row, 14).setValue(payload.actor || "");
        historySheet.getRange(updates[i].row, 15).setValue("Soạn hàng");
        historySheet.getRange(updates[i].row, 13).setValue("Đang xử lý");
      } else {
        historySheet.getRange(updates[i].row, 8).clearContent();
        historySheet.getRange(updates[i].row, 9).clearContent();
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
      if (!khoXuat && allRows[i][2]) khoXuat = String(allRows[i][2]).trim();
      if (!khoNhan && allRows[i][3]) khoNhan = String(allRows[i][3]).trim();
    }
    if (soPhieu) {
      if (!khoNhan) khoNhan = getKhoNhanBySoPhieu(soPhieu);
      if (khoNhan) {
        sendTelegramOrderReady(soPhieu, khoNhan);
      }
      var statusLabel = (totalRows > 0 && missingCount === 0) ? "Đủ hàng" : "Thiếu hàng";
      sendTelegramPackingSummary(soPhieu, khoXuat, khoNhan, updates.length, totalRows, statusLabel, missingCount, payload.actor || "Chi nhánh");
    }

    return "✅ Đã lưu " + updates.length + " món và " + anhDaLuu + " ảnh!";
  } finally { lock.releaseLock(); }
}