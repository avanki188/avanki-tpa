# Bộ kiểm tra tự động – Avanki TPA

Chạy toàn bộ: `python3 tests/run_tests.py`

- Cần: `pip install playwright openpyxl` và Chrome/Chromium (đặt biến `CHROME=` nếu dùng Chrome riêng).
- Mở `index.html` trong trình duyệt không giao diện, kho dữ liệu giả lập trong bộ nhớ – **không đụng OneDrive**.
- Toàn bộ file Excel được tạo lúc chạy, **tên và số liệu đều bịa**.
- Kiểm tra: đọc Excel (kể cả file hệ thống có tiền tố `x:`), Chuyển đổi DPBT, Lập BK, Giấy ĐNTT, Fix lỗi, Đối chiếu số tiền + dấu ✓, Theo dõi HS TPA (Kho, cảnh báo Công văn sai mẫu TPA, giữ dấu ✓ khi sửa).
- Chạy trước mỗi lần đưa bản mới lên; có mục ✗ thì sửa xong mới đưa lên.
