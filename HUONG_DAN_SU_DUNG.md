# Hướng Dẫn Sử Dụng Hệ Thống Điều Chuyển / Đơn Hàng

## 1. Mục đích hệ thống

Hệ thống dùng để:
- Tạo đơn hàng hoặc lệnh điều chuyển
- Quản lý và chỉnh sửa đơn
- Soạn hàng theo từng phiếu hoặc theo bảng gom mã
- Xác nhận nhận hàng thực tế
- Theo dõi tồn kho đối chiếu với nhu cầu

## 2. Các thành phần chính

- `code.gs`: xử lý logic backend trên Google Apps Script
- `public/`: giao diện web đang dùng qua Vercel
- `api/gas-proxy.js`: trung gian gọi từ frontend sang Apps Script
- `TỔNG HỢP TỒN KHO`: dữ liệu tồn kho hiện tại
- `Lịch Sử Xuất Kho`: dữ liệu gốc của từng đơn

## 3. Luồng sử dụng chuẩn

### Bước 1: Nhập danh mục và tồn kho

Vào tab `Admin`:
- chọn `Loại cập nhật`
- chọn file local từ máy tính
- bấm `CẬP NHẬT TỰ ĐỘNG LÊN GOOGLE SHEET`

Kết quả:
- nếu chọn `Cập nhật tồn kho`:
	dữ liệu file được dán trực tiếp lên `TỔNG HỢP TỒN KHO`
- nếu chọn `Cập nhật file nhập khẩu thông tin`:
	dữ liệu file được dán trực tiếp lên `Data_Excel`

Lưu ý:
- chỉ tài khoản `Admin` mới nhìn thấy và thao tác được chức năng này
- file nhập khẩu thông tin có thể thay đổi số cột, hệ thống sẽ lưu nguyên dữ liệu file vào `Data_Excel`
- cần làm bước này trước khi tạo đơn nếu dữ liệu mã hàng hoặc tồn kho mới thay đổi

### Bước 2: Tạo đơn

Ở tab tạo đơn:
- chọn loại phiếu: `Đơn Hàng` hoặc `Điều Chuyển`
- quét mã hoặc tìm mã hàng
- nhập số lượng
- gửi tạo phiếu

Lưu ý:
- nếu có mã lỗi, hệ thống vẫn tạo đơn nhưng sẽ cảnh báo để sửa trong tab quản lý
- `Số lượng yêu cầu` là số lượng gốc của đơn, không nên dùng để ghi số thực soạn

### Bước 3: Quản lý / chỉnh sửa đơn

Ở tab `Quản lý`:
- mở phiếu cần sửa
- chi nhánh có thể sửa số lượng yêu cầu, thêm mã, hủy dòng trước khi đơn được soạn xong
- admin có thể sửa số lượng, thêm mã, hủy dòng, hủy đơn

Lưu ý:
- đơn đã `Xác nhận nhận hàng` sẽ không được sửa nữa
- khi đơn đã `Đã soạn hàng`, tài khoản chi nhánh chỉ được xem, không được sửa nữa
- nếu admin sửa một đơn đã `Đã soạn hàng`, hệ thống sẽ mở lại đơn để soạn lại
- hủy dòng và hủy đơn vẫn giữ lịch sử để tra cứu

### Bước 4: Soạn hàng theo từng phiếu

Ở tab `Soạn Hàng`:
- chọn ngày xem đơn
- chọn phiếu cần soạn
- nhập `Thực tế`
- chụp ảnh nếu cần
- bấm lưu

Ý nghĩa dữ liệu:
- `SL Yêu Cầu`: số lượng gốc của đơn
- `Thực tế`: số lượng đã soạn thực tế

Lưu ý quan trọng:
- từ bản sửa hiện tại, khi lưu soạn hàng hệ thống chỉ cập nhật `Số lượng thực tế`
- hệ thống không còn ghi đè `Số lượng yêu cầu`
- điều này giúp báo cáo thiếu/thừa và xác nhận nhận hàng chính xác hơn

### Bước 5: Tạo bảng gom mã để đi nhặt hàng nhanh

Ở tab `Soạn Hàng`:
- chọn ngày tạo bảng từ đơn ngày
- bấm `TẠO BẢNG SOẠN NGÀY MAI`

Hệ thống sẽ:
- gom toàn bộ đơn của ngày đã chọn
- tạo 1 tab mới trong Google Sheet
- cộng tổng nhu cầu theo từng mã
- tách cột theo từng cửa hàng nhận như Q7, Q8, PH...
- đối chiếu tồn kho
- cảnh báo mã thiếu
- gợi ý cách xử lý thiếu

Cách dùng bảng này:
- in bảng ra giấy
- nhân viên đi theo từng mã
- mỗi mã lấy 1 lần đủ cho tất cả đơn
- sau đó chia theo từng cột cửa hàng

## 4. Xác nhận nhận hàng

Ở tab `Xác nhận nhận hàng`:
- chọn phiếu
- kiểm tra số lượng đã soạn
- nhập số lượng thực nhận
- bấm lưu xác nhận

Từ bản sửa hiện tại:
- số thực nhận sẽ được ghi lại vào dữ liệu đơn
- trạng thái được cập nhật thành `Đã xác nhận nhận hàng`
- lưu được người xác nhận và nguồn cập nhật

Lưu ý:
- chi nhánh chỉ xác nhận được phiếu thuộc kho của mình
- admin có thể xem toàn bộ

## 5. Ý nghĩa trạng thái

- `Mới`: đơn vừa tạo, chưa có dữ liệu soạn thực tế
- `Đã soạn hàng`: kho xuất đã nhập số lượng soạn thực tế xong
- `Đã xác nhận nhận hàng`: kho nhận đã xác nhận xong
- `Đã hủy dòng`: mã hàng đó bị hủy khỏi đơn
- `Đã hủy đơn`: toàn bộ đơn bị hủy

## 6. Các lỗi logic đã được sửa

### Đã sửa
- lưu soạn hàng không còn ghi đè `Số lượng yêu cầu`
- xác nhận nhận hàng đã ghi lại số thực nhận vào dữ liệu đơn
- tab xác nhận không còn phụ thuộc sai vào bộ lọc của tab quản lý
- action tạo bảng soạn ngày mai không còn mở qua `GET`
- tab `Tài Khoản` được ẩn hoàn toàn với tài khoản chi nhánh
- tài khoản chi nhánh được sửa đơn trước khi đơn chuyển sang `Đã soạn hàng`
- đơn đã `Đã soạn hàng` chỉ còn admin được phép sửa
- admin sửa đơn đã soạn sẽ mở lại đơn để cần soạn lại
- ô số lượng bị xóa trống không còn bị hiểu nhầm là hủy dòng

### Tác động sau khi sửa
- báo cáo thiếu/thừa đúng hơn
- số lượng gốc của đơn được giữ nguyên
- màn hình xác nhận nhận hàng phản ánh đúng dữ liệu gần nhất
- giảm nguy cơ tạo tab ngoài ý muốn do gọi sai URL

## 7. Khuyến nghị vận hành

- mỗi ngày nên cập nhật tồn kho trước khi tạo bảng gom mã
- chỉ dùng tab `Quản lý` để sửa số lượng gốc của đơn
- chỉ dùng tab `Soạn Hàng` để nhập số thực soạn
- chỉ dùng tab `Xác nhận nhận hàng` để nhập số thực nhận cuối cùng
- hạn chế sửa đơn sau khi kho đã bắt đầu soạn

## 8. Khi cần kiểm tra lỗi

Nên kiểm tra theo thứ tự:
- mã hàng có tồn tại trong `Data_Excel` không
- tồn kho đã cập nhật mới nhất chưa
- dòng đơn đang ở trạng thái gì
- số lượng gốc và số lượng thực tế có bị nhập nhầm tab không
- tài khoản đang dùng có đúng quyền và đúng kho không
