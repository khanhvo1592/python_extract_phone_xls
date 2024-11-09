# Cách sử dụng
## Chuẩn bị file Excel
File Excel phải có ít nhất 2 cột:
Cột chứa tên khách hàng
Cột chứa số điện thoại
Hỗ trợ cả định dạng .xls và .xlsx
Số điện thoại có thể ở các định dạng:
10 số (VD: 0912345678)
9 số (VD: 912345678)
Có mã quốc gia (VD: 84912345678)
Có dấu phân cách (VD: 0912-345-678)
## Các bước thực hiện
Chọn file Excel
Nhấn nút "Chọn file Excel"
Tìm và chọn file Excel cần xử lý
Chọn Sheet
Chọn sheet chứa dữ liệu từ danh sách dropdown
Nếu file chỉ có 1 sheet, hệ thống sẽ tự động chọn
## Chọn cột dữ liệu
Chọn cột chứa tên khách hàng từ dropdown "Cột tên"
Chọn cột chứa số điện thoại từ dropdown "Cột SĐT"
Xử lý dữ liệu
Nhấn nút "Xử lý" để bắt đầu trích xuất
Chờ quá trình xử lý hoàn tất
## Kết quả
File kết quả "results.xlsx" sẽ được tạo trong cùng thư mục với file gốc
File kết quả bao gồm 3 cột:
No id: Số thứ tự
Khách hàng: Tên khách hàng
sdt: Số điện thoại đã chuẩn hóa
## Quy tắc chuẩn hóa số điện thoại
Số điện thoại hợp lệ phải:
Có 10 số
Bắt đầu bằng số 0
Không bắt đầu bằng: 00, 01, 02, 030, 031
Các trường hợp xử lý:
Số 9 chữ số: thêm số 0 vào đầu
Số có mã 84: chuyển thành số 0
Loại bỏ các ký tự đặc biệt (-, , ., space)
## Xử lý lỗi thường gặp
Không đọc được file Excel
Kiểm tra định dạng file (.xls hoặc .xlsx)
Đảm bảo file không bị hỏng
Đóng file Excel nếu đang mở
Không tìm thấy số điện thoại hợp lệ
Kiểm tra định dạng số trong file gốc
Đảm bảo chọn đúng cột số điện thoại
Kiểm tra số điện thoại có đúng định dạng Việt Nam
## Lỗi khi ghi file kết quả
Đảm bảo đã đóng file results.xlsx
Kiểm tra quyền ghi trong thư mục
Đảm bảo ổ đĩa còn đủ dung lượng
## Lưu ý quan trọng
Sao lưu dữ liệu gốc trước khi xử lý
Không đổi tên file results.xlsx trong quá trình sử dụng
Mỗi lần xử lý sẽ thêm dữ liệu mới vào cuối file results.xlsx
Kiểm tra kết quả sau mỗi lần xử lý
## Hỗ trợ
Nếu cần hỗ trợ thêm, vui lòng liên hệ:
Email: khanhvo1592@gmail.com
Phone: 0914257604