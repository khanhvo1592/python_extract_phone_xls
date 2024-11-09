# python_extract_phone_xls

### Bước 1: Cài đặt Python
1. Tải Python từ trang chủ: https://www.python.org/downloads/
2. Chọn phiên bản Python 3.9 (khuyến nghị)
3. Trong quá trình cài đặt, nhớ tích vào "Add Python to PATH"

### Bước 2: Tạo thư mục dự án và môi trường ảo

```bash
# Tạo thư mục dự án
mkdir phone_extractor
cd phone_extractor

# Tạo môi trường ảo
python -m venv venv

# Kích hoạt môi trường ảo
# Trên Windows:
venv\Scripts\activate
# Trên Linux/Mac:
source venv/bin/activate
```

### Bước 3: Cài đặt các thư viện cần thiết

```bash
# Nâng cấp pip
pip install --upgrade pip

# Cài đặt các thư viện
pip install pandas
pip install openpyxl==3.1.2
```

### Bước 4: Tạo file code
Tạo file `app.py` và copy code sau vào:
### Bước 5: Chạy chương trình

```bash
# Đảm bảo đang ở trong môi trường ảo
python app.py
```

### Cách sử dụng:
1. Chương trình sẽ hiển thị một cửa sổ với nút "Chọn file Excel"
2. Nhấn nút để chọn file Excel cần xử lý
3. Chương trình sẽ:
   - Tìm tất cả số điện thoại trong file Excel
   - Chuẩn hóa số điện thoại (thêm số 0 cho số 9 chữ số)
   - Loại bỏ các số không hợp lệ (00x, 01x, 02x, 030x, 031x)
   - Tạo file txt chứa danh sách số điện thoại đã xử lý

### Lưu ý:
- File kết quả sẽ được lưu trong cùng thư mục với file Excel gốc
- Tên file kết quả sẽ có dạng: `phone_numbers_YYYYMMDD_HHMMSS.txt`
- Mỗi số điện thoại sẽ được ghi trên một dòng
- Các số điện thoại sẽ được sắp xếp theo thứ tự tăng dần