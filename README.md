# 🌸 LUV-DOC-TRAN — Trình Dịch Thuật Song Ngữ Kỹ Thuật Thương Mại

Ứng dụng Web độc lập, giao diện **Pastel Kawaii siêu cute, thân thiện và tối giản**, tích hợp toàn bộ sức mạnh dịch thuật song ngữ kỹ thuật đỉnh cao từ công nghệ **DOC-TRAN** của hệ sinh thái TitanPDF.

---

[![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com/deploy?repo=https://github.com/NguyenChiTinh27/luv-doc-tran)

---

## ✨ Điểm Nổi Bật & Tính Năng Thương Mại

1. **Giao Diện Pastel Kawaii Thân Thiện**:
   - Thiết kế chuẩn thương mại với bảng màu pastel ngọt ngào (hồng đào, tím lavender, xanh mint).
   - Linh vật chibi **Docy** đồng hành, tạo cảm giác dễ chịu và gần gũi cho kỹ sư & người dịch.

2. **Cơ Chế Dùng Thử 48 Giờ (Anti-Tamper 2-Day Trial)**:
   - Tự động nhận diện thiết bị (Client Device Fingerprint) và kích hoạt 48 giờ dùng thử đầy đủ tính năng.
   - Đồng hồ đếm ngược thời gian thực trên thanh tiêu đề. Khi hết hạn 48h, hệ thống tự động khóa và yêu cầu Mã Bản Quyền.

3. **Cơ Chế Khóa Bản Quyền Lai (Hybrid Cryptographic License)**:
   - **Tương thích 100% với ứng dụng native `LUV_License_Manager.exe`** (sử dụng cùng thuật toán băm SHA-256 với `SECRET_SALT = "LUV_PDF_OFFICIAL_SECRET_KEY_2026"`).
   - Hỗ trợ cả kích hoạt Offline tức thì và đồng bộ kiểm tra trực tuyến với máy chủ Quản trị TitanPDF (`tinh-of-2m9j`).

4. **Ưu Tiên Tốc Độ Word & Tùy Chọn Xuất PDF Với Hoạt Họa Đợi**:
   - **Ưu tiên Word (.docx)**: Nút tải nổi bật phát sáng, hoàn thành trong 3-5 giây, dễ dàng chỉnh sửa.
   - **Tùy chọn xuất PDF vector**: Hiển thị modal hoạt hình dễ thương kèm thông báo rõ ràng *"Đang kết xuất tệp PDF vector độ nét cao, vui lòng đợi khoảng 2 - 3 phút nhen (Khuyên dùng: Tải file Word sẽ nhanh hơn rất nhiều)"*.
   - Bộ 3 nút tải hoàn chỉnh: Word (.docx), PDF (.pdf), ZIP trọn bộ.

5. **Đa Định Dạng Mạnh Mẽ**:
   - Word: `.docx` và `.doc` (Word 97-2003 nhị phân).
   - PDF: Tự động chuyển đổi sang Word, dịch song ngữ và hỗ trợ xuất ngược PDF vector.
   - Excel: `.xlsx` và `.xls` bảo toàn 100% công thức (`=SUM`, `=IF`,...), merged cells và định dạng viền bảng.

6. **Review Split-View Studio**:
   - Màn hình so sánh 2 cột trực quan (Văn bản gốc vs. Bản dịch song ngữ).
   - Thanh công cụ HUD: Zoom (+/-), Cuộn đồng bộ (Synchronized Scroll), In ấn, Chỉnh sửa nhanh (Inline Edit).

---

## 🚀 Hướng Dẫn Triển Khai Lên Render (1-Click Deploy)

### Cách 1: Sử dụng Docker Environment (Khuyên Dùng)
1. Tạo một repository mới trên GitHub (ví dụ: `luv-doc-tran`) và push toàn bộ thư mục này lên.
2. Đăng nhập [Render.com](https://render.com) -> Chọn **New Web Service**.
3. Kết nối với repository vừa tạo.
4. Render sẽ tự động nhận diện `Dockerfile` hoặc `render.yaml`:
   - **Environment**: `Docker`
   - **Region**: `Singapore`
   - **Health Check Path**: `/api/status`
5. Bấm **Deploy Web Service** và hoàn tất!

### Cách 2: Chạy Local Trên Máy Tính
```bash
# Cài đặt thư viện
pip install -r requirements.txt

# Khởi chạy server
python server.py
# hoặc
uvicorn server:app --host 0.0.0.0 --port 8080 --reload
```
Truy cập trình duyệt: `http://localhost:8080`.

---

## 🔑 Cấu Trúc Khóa Bản Quyền

- Định dạng: `LUV-PRO-XXXX-XXXX-XXXX` hoặc `LUV-ENT-XXXX-XXXX-XXXX`
- Cấp khóa trực tiếp từ:
  - Ứng dụng Desktop: `LUV_License_Manager.exe` trong ổ `KKK`
  - Web Quản trị TitanPDF: Tab `🔑 Quản Lý Bản Quyền LUV`
- Master Demo Key để kiểm thử nhanh:
  - `LUV-PDF-2026-PRO-VN99`
  - `LUV-2026-ENTERPRISE`
