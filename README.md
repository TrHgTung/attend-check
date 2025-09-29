QR Attendance System (Google Sheet + Apps Script)
Giới thiệu

Hệ thống điểm danh bằng mã QR sử dụng Google Sheets và Google Apps Script.
Mỗi người tham gia có một Token duy nhất để sinh ra mã QR, khi quét mã sẽ ghi lại thông tin điểm danh vào Google Sheet.

Cấu trúc Sheet

Họ tên: tên người tham dự

Số điện thoại: số liên hệ

Token: mã duy nhất dùng để tạo QR và check-in

Email: tùy chọn

Đã tham dự: trạng thái điểm danh (Có / trống)

CheckinAt: thời gian check-in (full datetime)

Ngày: ngày check-in

Giờ: giờ check-in

Nhóm: nhóm tham dự (tùy chỉnh)

Checkin_URL: đường dẫn check-in (Web App)

Chức năng chính

Sinh Token

Sử dụng hàm generateTokensSafe() để tạo Token.

Token giữ cố định, không đổi khi chỉnh sửa các cột khác.

Xuất mã QR

Sử dụng hàm exportQRCodesWithLabels() để sinh mã QR kèm tên, nhóm, số điện thoại.

QR chứa đường dẫn check-in với Token.

Điểm danh (Check-in)

Web App (doGet) xác thực theo Token.

Khi quét mã:

API tự ghi nhận "Có" vào cột Đã tham dự

Lưu thời gian, ngày, giờ check-in

Nếu đã check-in trước đó sẽ hiện cảnh báo.

Triển khai

Mở Google Apps Script từ Google Sheet.

Copy code vào file Code.gs.

Điều chỉnh biến cấu hình:
