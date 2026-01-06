# 📊 HƯỚNG DẪN TẠO SHEET "OVERVIEW" CHO TASK LIST

## 🚀 CÁCH NHANH NHẤT: Sử dụng Google Apps Script

### Bước 1: Mở Apps Script
1. Mở Google Sheet của bạn: https://docs.google.com/spreadsheets/d/1N_f8TaqdUu1RKuKSFk0essrEQ95fdUbR5t4mvnsZj8c/edit
2. Vào menu **Extensions** → **Apps Script**

### Bước 2: Copy code
1. Xóa toàn bộ code mặc định (function myFunction() {...})
2. Copy toàn bộ nội dung file `create_overview_sheet.gs` và paste vào

### Bước 3: Điều chỉnh CONFIG (QUAN TRỌNG!)
Tìm phần **CONFIG** ở đầu file và điều chỉnh theo cấu trúc sheet "Task List" của bạn:

```javascript
const CONFIG = {
  taskListSheetName: "Task List",  // Tên sheet chứa task
  
  // Vị trí cột (A=1, B=2, C=3, ...)
  columns: {
    taskId: 1,        // Cột A
    taskName: 2,      // Cột B
    description: 3,   // Cột C
    assignee: 4,      // Cột D - QUAN TRỌNG: cột Người được giao
    status: 5,        // Cột E - QUAN TRỌNG: cột Trạng thái
    priority: 6,      // Cột F - QUAN TRỌNG: cột Độ ưu tiên
    dueDate: 7,       // Cột G
    remainingTime: 8, // Cột H - Thời gian còn lại
    startDate: 9      // Cột I
  },
  
  // Điều chỉnh giá trị Status theo sheet của bạn
  status: {
    done: ["Finished", "Closed"],     // Các status = "Done"
    inProgress: ["In Progress"],       // Status đang làm
    pending: ["To Do", "Open"]         // Status chưa làm
  },
  
  // Điều chỉnh giá trị Priority theo sheet của bạn
  priority: {
    urgent: "Urgent",   // hoặc "Khẩn cấp"
    high: "High",       // hoặc "Cao"
    medium: "Medium",   // hoặc "Trung bình"
    low: "Low"          // hoặc "Thấp"
  }
};
```

### Bước 4: Chạy Script
1. Nhấn nút **Run** (▶️) ở thanh công cụ
2. Chọn function: **createOverviewSheet**
3. Lần đầu chạy, Google sẽ yêu cầu cấp quyền:
   - Click "Review permissions"
   - Chọn tài khoản Google của bạn
   - Click "Advanced" → "Go to [project name] (unsafe)"
   - Click "Allow"

### Bước 5: Hoàn tất! 🎉
Sheet "Overview" sẽ được tạo tự động với:
- ✅ Bảng KPI tổng quan (tổng task, đã xong, đang làm, %)
- ✅ Thống kê theo Status (số lượng + phần trăm)
- ✅ Thống kê theo Priority (với cảnh báo task urgent)
- ✅ Thống kê theo Assignee
- ✅ Danh sách task sắp hết hạn
- ✅ Bảng chi tiết từng người (done/in progress/pending + priority + task đang làm)

---

## 📝 CÁCH THỦ CÔNG: Dùng công thức trực tiếp

Nếu bạn không muốn dùng Apps Script, có thể tự tạo sheet Overview và nhập các công thức sau:

### Giả sử cấu trúc Task List:
- Cột D: Assignee
- Cột E: Status  
- Cột F: Priority
- Cột H: Remaining Time (số ngày)

### 1️⃣ Thống kê Status

| Ô | Nội dung |
|---|----------|
| A1 | `Trạng thái` |
| B1 | `Số lượng` |
| C1 | `Phần trăm` |
| A2 | `To Do` |
| A3 | `In Progress` |
| A4 | `Finished` |
| A5 | `Closed` |
| B2 | `=COUNTIF('Task List'!E:E,"To Do")` |
| B3 | `=COUNTIF('Task List'!E:E,"In Progress")` |
| B4 | `=COUNTIF('Task List'!E:E,"Finished")` |
| B5 | `=COUNTIF('Task List'!E:E,"Closed")` |
| C2 | `=IF(SUM($B$2:$B$5)>0,B2/SUM($B$2:$B$5),0)` |

### 2️⃣ Thống kê Assignee (dùng QUERY)

```
=QUERY('Task List'!D2:E,"SELECT D, COUNT(D) WHERE D IS NOT NULL GROUP BY D ORDER BY COUNT(D) DESC LABEL COUNT(D) 'Số Task'")
```

### 3️⃣ Bảng chi tiết Assignee

| Cột | Header | Công thức (cho hàng 2) |
|-----|--------|------------------------|
| A | Assignee | `=UNIQUE('Task List'!D2:D)` |
| B | Tổng | `=COUNTIF('Task List'!D:D,A2)` |
| C | Done | `=COUNTIFS('Task List'!D:D,A2,'Task List'!E:E,"Finished")+COUNTIFS('Task List'!D:D,A2,'Task List'!E:E,"Closed")` |
| D | In Progress | `=COUNTIFS('Task List'!D:D,A2,'Task List'!E:E,"In Progress")` |
| E | Pending | `=B2-C2-D2` |
| F | Urgent | `=COUNTIFS('Task List'!D:D,A2,'Task List'!F:F,"Urgent",'Task List'!E:E,"<>Finished",'Task List'!E:E,"<>Closed")` |
| G | High | `=COUNTIFS('Task List'!D:D,A2,'Task List'!F:F,"High",'Task List'!E:E,"<>Finished",'Task List'!E:E,"<>Closed")` |
| H | Medium | `=COUNTIFS('Task List'!D:D,A2,'Task List'!F:F,"Medium",'Task List'!E:E,"<>Finished",'Task List'!E:E,"<>Closed")` |
| I | Low | `=COUNTIFS('Task List'!D:D,A2,'Task List'!F:F,"Low",'Task List'!E:E,"<>Finished",'Task List'!E:E,"<>Closed")` |
| J | Task đang làm | `=TEXTJOIN(", ",TRUE,FILTER('Task List'!B:B,('Task List'!D:D=A2)*('Task List'!E:E="In Progress")))` |

### 4️⃣ Task sắp hết hạn (trong 3 ngày)

```
=FILTER({'Task List'!B2:B,'Task List'!D2:D,'Task List'!F2:F,'Task List'!G2:G,'Task List'!H2:H,'Task List'!E2:E},('Task List'!H2:H<=3)*('Task List'!H2:H>=0)*('Task List'!E2:E<>"Finished")*('Task List'!E2:E<>"Closed"))
```

### 5️⃣ Thống kê Priority

| Ô | Nội dung |
|---|----------|
| A1 | `Priority` |
| B1 | `Tổng` |
| C1 | `Chưa xong` |
| D1 | `%` |
| A2 | `Urgent` |
| B2 | `=COUNTIF('Task List'!F:F,"Urgent")` |
| C2 | `=COUNTIFS('Task List'!F:F,"Urgent",'Task List'!E:E,"<>Finished",'Task List'!E:E,"<>Closed")` |
| D2 | `=IF(SUM($B$2:$B$5)>0,B2/SUM($B$2:$B$5),0)` |

---

## ❓ CẦN HỖ TRỢ?

Nếu cấu trúc Task List của bạn khác với giả định trên, hãy cho tôi biết:

1. **Screenshot** header row của sheet "Task List"
2. **Các giá trị Status** có thể có
3. **Các giá trị Priority** có thể có
4. **Cột Remaining Time** là số hay text (ví dụ: "2 days")?

Tôi sẽ điều chỉnh script/công thức cho phù hợp!

---

## 📁 FILES ĐÃ TẠO

1. `create_overview_sheet.gs` - Google Apps Script hoàn chỉnh
2. `google_sheets_overview_guide.md` - Hướng dẫn chi tiết với công thức
3. `HUONG_DAN_TAO_OVERVIEW_SHEET.md` - File này
