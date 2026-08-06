# 🚀 RIVA Checklist - Hệ Thống Quản Lý Phân Công Công Việc 6 Dự Án Chiến Lược

Hệ thống Dashboard theo dõi phân công công việc trực quan, tự động làm sạch danh sách nhân sự và giám sát tiến độ dự án real-time cho 6 danh mục chiến lược của RIVA.

🌐 **Trang Web Dashboard Trực Tuyến**: [https://akashi0310.github.io/RIVA_checklist/](https://akashi0310.github.io/RIVA_checklist/)

---

## 🔒 Bảo Mật Truy Cập
Trang Dashboard được trang bị màn hình khóa xác thực bảo mật nội bộ để tránh truy cập trái phép khi publish công khai trên GitHub Pages.
- **Mật khẩu truy cập nội bộ**: `riva2026` hoặc `RIVA2026`

---

## 🌟 Tính Năng Nổi Bật
1. **Tra Cứu Nhân Viên Toàn Hệ Thống**: Tìm kiếm & liệt kê 100% công việc của từng nhân viên trên cả 6 dự án cùng lúc.
2. **Mã Màu Phân Loại Dự Án**: 
   - 🏆 **AMSIO**: Tím Xanh (`#6366f1`)
   - 🎓 **AP**: Xanh Lá (`#10b981`)
   - ✈️ **IENA (Đức)**: Vàng Cam (`#f59e0b`)
   - 🇹🇭 **IPITEX (Thái Lan)**: Hồng (`#ec4899`)
   - 🔬 **NCKH**: Xanh Ngọc (`#06b6d4`)
   - 📣 **Sản Phẩm Truyền Thông**: Tím (`#a855f7`)
3. **Cảnh Báo Deadline Nóng (≤ 4 ngày)**: Công việc chưa hoàn thành có deadline sắp tới hoặc quá hạn sẽ được hiển thị **Màu Đỏ In Đậm kèm Biểu Tượng Cảnh Báo ⚠️**.
4. **Lưu Trạng Thái Đã Xong/Chưa Xong**: Tự động lưu vào bộ nhớ trình duyệt `localStorage`.
5. **Xuất Excel Báo Cáo Cá Nhân & Dự Án**: Hỗ trợ xuất file `.csv` chuẩn UTF-8 BOM.

---

## 📁 Cấu Trúc Thư Mục
- `index.html`: Giao diện Web Dashboard chính tích hợp màn khóa bảo mật.
- `dashboard_data.js`: Cơ sở dữ liệu 240 công việc đã xử lý.
- `update_dashboard.py`: Script Python tự động đọc 6 file Excel & làm sạch dữ liệu.
- `cap_nhat_du_lieu.bat`: File nhấp đôi chuột 1-click để cập nhật lại dữ liệu khi sửa Excel.
- `*.xlsx`: 6 file dữ liệu Excel gốc (AMSIO, AP, IENA, IPITEX, NCKH, SẢN PHẨM TRUYỀN THÔNG).

---

## 🛠️ Hướng Dẫn Cập Nhật Dữ Liệu
Khi bạn chỉnh sửa hoặc bổ sung bất kỳ file Excel nào trong thư mục này:
1. Nhấp đôi chuột vào file `cap_nhat_du_lieu.bat`.
2. Hệ thống sẽ tự động cập nhật file `dashboard_data.js`.
3. Mở lại hoặc nhấn **F5** trên `index.html` để thấy dữ liệu mới!

---

## 🌐 Hướng Dẫn Đẩy Code Lên GitHub `akashi0310/RIVA_checklist`
Nếu bạn sử dụng Git:
```bash
git add .
git commit -m "Update RIVA Checklist Dashboard with security password"
git push origin main
```
Hoặc kéo thả trực tiếp các file trong thư mục `RIVA_checklist` vào giao diện Repository [https://github.com/akashi0310/RIVA_checklist](https://github.com/akashi0310/RIVA_checklist).
