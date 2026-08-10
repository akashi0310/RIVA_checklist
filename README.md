# 📋 RIVA Checklist & Dashboard Quản Lý Phân Công Công Việc 5 Dự Án Chiến Lược

Hệ thống Dashboard theo dõi tiến độ, phân công công việc & tra cứu nhân sự trực quan dành cho 5 dự án trọng điểm RIVA.

---

## 🌟 Tính Năng Nổi Bật

- **Báo Cáo Tổng Quan 5 Dự Án**: AMSIO Việt Nam, Tuyển Sinh AP, Đội Tuyển IENA (Đức), Đội Tuyển IPITEX (Thái Lan), NCKH (2026-2027).
- **Tra Cứu Nhân Sự Hàng Đầu (Global Search Taskbar)**: Tìm kiếm & chọn nhanh nhân sự ngay ở hàng trên cùng của trang web.
- **Tự Động Cảnh Báo Deadline (≤ 4 ngày & Quá hạn)**: Công việc gần deadline hiển thị màu đỏ in đậm kèm biểu tượng cảnh báo ⚠️ và được tự động đẩy lên đầu danh sách.
- **Chế Độ Xem Hộp Ngang Khi Chọn 1 Nhân Sự (Horizontal Single Person Box)**: Hiển thị giao diện dàn ngang rộng rãi, dễ quan sát chi tiết từng nhiệm vụ.
- **Tương Tác Tự Động Định Vị Công Việc (Jump & Flash Highlight)**: Nhấp vào bất kỳ công việc nào trong mục Phân Công Nhân Sự sẽ tự động chuyển sang Bảng Công Việc, cuộn màn hình và phát sáng dòng công việc đó.
- **Bảo Mật Đăng Nhập An Toàn**: Bảo mật phân quyền đăng nhập mở khóa hệ thống.

---

## 📁 Cấu Trúc File Dự Án

```text
RIVA_checklist/
├── index.html           # File giao diện web chính (HTML/CSS/JavaScript)
├── update_dashboard.py  # Script Python cập nhật dữ liệu từ thư mục Excel_file
├── dashboard_data.js    # Cơ sở dữ liệu 5 dự án & phân công nhân sự
├── cap_nhat_du_lieu.bat # File chạy nhanh 1-click để cập nhật dữ liệu
├── Excel_file/          # Thư mục chứa các file Excel dữ liệu nguồn
└── README.md            # Tài liệu hướng dẫn sử dụng
```

---

## 🌐 Trải Nghiệm Trực Tuyến

Truy cập hệ thống Dashboard tại: [https://akashi0310.github.io/RIVA_checklist/](https://akashi0310.github.io/RIVA_checklist/)
