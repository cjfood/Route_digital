# HƯỚNG DẪN SỬ DỤNG BẢN ĐỒ TUYẾN BÁN HÀNG DMS - CJ FOODS (OFFLINE STANDALONE)

Hệ thống hiện tại đã được nâng cấp thành **Web App Trực Tiếp (Standalone)**:
👉 **KHÔNG CẦN CHẠY BẤT KỲ LỆNH NÀO**, không cần cài đặt Web Server!

---

## 1. Cách Mở Web App Trực Tiếp
* Bạn chỉ cần vào thư mục dự án và **nhấp đúp chuột trực tiếp vào file [`index.html`](file:///d:/CJ/5.%20Route%20Setting/Project%20Map_CJ_Golden/index.html)**.
* Trình duyệt web (Chrome, Edge, Cốc Cốc, Safari...) sẽ tự động mở trang web lên ngay lập tức.
* **Xem trên điện thoại:** Bạn có thể copy cả thư mục này sang điện thoại, gửi file qua Zalo/Drive hoặc mở qua trình duyệt điện thoại để xem offline mà không phụ thuộc vào server.

---

## 2. Các Tính Năng Đã Tích Hợp Sẵn Trong Trang `index.html`

1. **Bản đồ Google Maps Chuẩn (100% Rõ Nét, Đầy Đủ Toàn Bộ Điểm Bán Toàn Quốc):**
   * Hiển thị toàn bộ **57.630 điểm Shop CJ** và **61.746 điểm Golden Store** trên khắp cả nước (không bị giới hạn/cắt bớt điểm bán như trước).
   * Có nút chuyển đổi linh hoạt: **Bản đồ Google**, **Google Vệ tinh (Satellite)** và **OpenStreetMap**.
2. **Hệ Thống Điểm & Vòng Gom Cụm (Cluster) Đồng Bộ Màu Sắc Tuyệt Đối:**
   * **Shop CJ (MCP):** Màu **Xanh đậm** (`#0A2E50`) hình tròn viền trắng, vòng gom cụm cũng có màu Xanh đậm có chữ **CJ**.
   * **Golden Store (Loại A):** Màu **Cam hơi đậm** (`#D35400`) kích thước lớn 32px viền vàng kim có chữ **GS**, vòng gom cụm có màu Cam đậm có chữ **GS**.
   * Hoàn toàn không còn các vòng tròn màu xanh lá/vàng mặc định gây nhầm lẫn.
3. **Tùy Chọn Bật/Tắt Lớp Bản Đồ Thông Minh:**
   * `[x] Hiển thị Shop CJ (MCP)`
   * `[x] Hiển thị Golden Store (Loại A)`
   * `[x] Gom cụm bản đồ (Cluster)`: Tắt để thấy toàn bộ các điểm bán riêng lẻ và đường lộ trình khi soi tuyến một NVBH.
   * `[x] Vẽ lộ trình nối các điểm viếng thăm`
5. **Nút Định Vị Thực Tế Của Bạn (GPS My Location):** Bấm nút `📍` ở góc dưới bên trái để định vị chấm xanh vị trí thực tế của bạn trên bản đồ khi đi thị trường.
6. **Xuất Tuyến Ra Excel Trực Tiếp:** Bấm nút **"Xuất Tuyến Excel"** góc trên bên phải để tải ngay file Excel về máy tính mà không cần qua server trung gian.

---

## 3. Cách Cập Nhật Dữ Liệu Khi Có File Excel Mới
Mỗi khi bạn có file Excel/XLSB mới trong tháng:
1. Copy file mới vào thư mục tương ứng:
   * File MCP mới $\rightarrow$ Thư mục `Shop CJ/`
   * File Tiềm năng mới $\rightarrow$ Thư mục `Shop Potential/`
2. **Nhấp đúp chuột vào file [`cap_nhat_data.bat`](file:///d:/CJ/5.%20Route%20Setting/Project%20Map_CJ_Golden/cap_nhat_data.bat)**.
   * Script sẽ tự động cập nhật Master Database và xuất lại file `dms_data.js` mới nhất trong vòng ~30 giây.
3. Sau đó, bạn chỉ cần mở lại file `index.html` (hoặc bấm F5) để xem dữ liệu mới!
