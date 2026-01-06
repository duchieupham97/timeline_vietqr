# 📊 HƯỚNG DẪN TẠO OVERVIEW SHEET

## 🚀 CÁC BƯỚC THỰC HIỆN

### Bước 1: Mở Apps Script
1. Mở Google Sheet của bạn
2. Vào **Extensions** → **Apps Script**

### Bước 2: Paste Code
1. **Xóa toàn bộ** code mặc định trong editor
2. Mở file `create_overview_sheet.gs` trong workspace
3. **Copy toàn bộ** nội dung
4. **Paste** vào Apps Script editor

### Bước 3: Chạy Script
1. Nhấn nút **Run** (▶️) trên thanh công cụ
2. Đảm bảo function được chọn là: `createOverviewSheet`
3. **Lần đầu chạy**, Google sẽ yêu cầu cấp quyền:
   - Click **"Review permissions"**
   - Chọn tài khoản Google của bạn
   - Click **"Advanced"** → **"Go to [project name] (unsafe)"**
   - Click **"Allow"**

### Bước 4: Xong! 🎉
Sheet **"Overview"** sẽ được tạo tự động ở vị trí đầu tiên.

---

## ✨ TÍNH NĂNG ĐÃ CÓ

| # | Tính năng | Mô tả |
|---|-----------|-------|
| 1 | **KPI Dashboard** | Tổng task, đã hoàn thành, đang làm, testing, chờ xử lý, urgent, % hoàn thành |
| 2 | **Biểu đồ Status** | Biểu đồ tròn thống kê theo trạng thái với số lượng và % |
| 3 | **Biểu đồ Priority** | Biểu đồ cột theo độ ưu tiên |
| 4 | **Thống kê Assignee** | Số task của từng người (hỗ trợ multiple select) |
| 5 | **Task sắp hết hạn** | Danh sách task trong 7 ngày tới chưa hoàn thành |
| 6 | **Bảng chi tiết Workload** | Mỗi người: Done, In Progress, Testing, Pending, Urgent, High, Normal, Low |
| 7 | **Task đang làm** | Hiển thị FNo. và tên task mỗi người đang làm (In Progress) |
| 8 | **Cảnh báo Priority** | Highlight các task Urgent/High cần xử lý |

---

## 🔄 REALTIME UPDATE

Tất cả dữ liệu trong Overview sẽ **tự động cập nhật** khi bạn thay đổi Task List.
Không cần chạy lại script!

---

## ➕ THÊM ASSIGNEE MỚI

Nếu team có thêm thành viên mới:

1. Mở **Extensions** → **Apps Script**
2. Tìm dòng `assignees:` trong phần CONFIG
3. Thêm tên mới vào mảng:

```javascript
assignees: ["Duy Anh", "Trường", "Đức", "Triều", "Nghĩa", "Hiếu Phạm", "Quyết", "Hiếu Hà", "Tôn", "Tên mới"],
```

4. Chạy lại function `createOverviewSheet`

---

## 📝 LƯU Ý

- **Multiple Select Assignee**: Script đã xử lý trường hợp 1 task có nhiều người được giao
- **Group Headers**: Script bỏ qua các hàng header nhóm (Customer Support, BackEnd)
- **Remaining Time**: Định dạng hh:mm được hỗ trợ
- **Biểu đồ**: Tự động tạo biểu đồ tròn và biểu đồ cột

---

## 📌 MENU TẮT

Sau khi chạy script lần đầu, bạn sẽ thấy menu mới:

**📊 Task Overview** →
- 🔄 Tạo/Cập nhật Overview
- 📈 Chỉ cập nhật biểu đồ
- ℹ️ Hướng dẫn
