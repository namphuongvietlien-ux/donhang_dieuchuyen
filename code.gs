// ============================================================
// code.gs — Cấu hình toàn cục (GAS shared global scope)
// Modules: api_routes.gs | batch_soan_hang.gs | packing_timeline.gs
//           catalog_variant.gs | utils_helpers.gs
// ============================================================

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
