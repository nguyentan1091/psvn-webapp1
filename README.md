# psvn-webapp1
Google Apps Script WebApp cho dự án XPPL của PSVN

## Vai trò User-Supervision

Ứng dụng tự động bảo đảm có tài khoản giám sát với thông tin sau:

- **Username:** `LA`
- **Password:** `CRLF@LA111`
- **Role:** `User-Supervision`

Tài khoản này được thiết kế cho bên thứ ba chỉ giám sát và trích xuất dữ liệu, nên không thể sửa, xoá hay thao tác gây ảnh hưởng đến dữ liệu nguồn. Các giới hạn chính gồm:

- Chỉ truy cập các trang **Xe đã đăng ký**, **Kết quả bốc hàng** và **Tài khoản** (để tự đổi mật khẩu).
- Ở trang **Xe đã đăng ký**, bảng dữ liệu chỉ hiển thị các xe đã được *Approved* và ẩn toàn bộ cột hành động/chọn dòng.
- Ở trang **Kết quả bốc hàng**, chỉ xem bộ lọc và dữ liệu kết quả, các nút thao tác và cột chọn dòng/hành động bị ẩn hoàn toàn.
- Không có quyền quản lý người dùng và không thể truy cập các trang chức năng khác trong menu.

Vai trò này phục vụ nhu cầu giám sát, bảo đảm bên thứ ba chỉ thu thập số liệu cần thiết mà không làm thay đổi dữ liệu vận hành của hệ thống.
