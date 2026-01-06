# Hướng dẫn tạo Sheet "Overview" cho Task List

## Giả định cấu trúc Sheet "Task List"

Giả sử sheet "Task List" của bạn có các cột như sau (điều chỉnh theo thực tế):

| Cột | Tên cột | Mô tả |
|-----|---------|-------|
| A | Task ID | Mã task |
| B | Task Name | Tên task |
| C | Description | Mô tả |
| D | Assignee | Người được giao |
| E | Status | Trạng thái (To Do, In Progress, Finished, Closed) |
| F | Priority | Độ ưu tiên (Urgent, High, Medium, Low) |
| G | Due Date | Ngày hết hạn |
| H | Remaining Time | Thời gian còn lại |
| I | Start Date | Ngày bắt đầu |

**Lưu ý:** Hãy điều chỉnh các công thức bên dưới theo đúng vị trí cột trong sheet của bạn.

---

## BƯỚC 1: Tạo Sheet "Overview"

1. Mở Google Sheet của bạn
2. Click vào dấu "+" ở góc dưới bên trái để tạo sheet mới
3. Đặt tên là "Overview"

---

## BƯỚC 2: Thống kê Task theo Status (Biểu đồ tròn)

### 2.1. Tạo bảng dữ liệu cho biểu đồ

Tại vị trí **A1**, nhập các công thức sau:

```
A1: Status
A2: To Do
A3: In Progress
A4: Finished
A5: Closed
A6: TỔNG

B1: Số lượng
B2: =COUNTIF('Task List'!E:E,"To Do")
B3: =COUNTIF('Task List'!E:E,"In Progress")
B4: =COUNTIF('Task List'!E:E,"Finished")
B5: =COUNTIF('Task List'!E:E,"Closed")
B6: =SUM(B2:B5)

C1: Phần trăm
C2: =IF($B$6>0,B2/$B$6,0)
C3: =IF($B$6>0,B3/$B$6,0)
C4: =IF($B$6>0,B4/$B$6,0)
C5: =IF($B$6>0,B5/$B$6,0)
C6: =SUM(C2:C5)
```

**Format cột C:** Chọn C2:C6 → Format → Number → Percent

### 2.2. Tạo biểu đồ tròn

1. Chọn vùng A1:C5
2. Insert → Chart
3. Chọn Chart type: Pie chart
4. Customize theo ý muốn

---

## BƯỚC 3: Thống kê Task theo Assignee (Biểu đồ cột)

### 3.1. Tạo bảng dữ liệu

Tại vị trí **E1**, nhập:

```
E1: Assignee
F1: Số Task

E2: =UNIQUE('Task List'!D2:D)
```

Sau đó tại **F2**, nhập công thức và kéo xuống:
```
F2: =COUNTIF('Task List'!D:D,E2)
```

### 3.2. Cách khác - Dùng QUERY function

```
E1: ={"Assignee","Số Task";QUERY('Task List'!D2:D,"SELECT D, COUNT(D) WHERE D IS NOT NULL GROUP BY D LABEL COUNT(D) ''")}
```

### 3.3. Tạo biểu đồ cột

1. Chọn vùng dữ liệu
2. Insert → Chart
3. Chọn Chart type: Bar chart hoặc Column chart

---

## BƯỚC 4: Bảng chi tiết theo Assignee (Thống kê đầy đủ)

### 4.1. Tạo bảng thống kê chi tiết

Tại vị trí **A10**, tạo bảng:

```
A10: Assignee
B10: Tổng Task
C10: Done (Finished/Closed)
D10: In Progress
E10: Pending
F10: Urgent
G10: High
H10: Medium
I10: Low
J10: Task đang làm

A11: =UNIQUE('Task List'!D2:D)
```

Cho mỗi hàng Assignee (bắt đầu từ hàng 11), nhập các công thức:

```
B11: =COUNTIF('Task List'!D:D,A11)
C11: =COUNTIFS('Task List'!D:D,A11,'Task List'!E:E,"Finished")+COUNTIFS('Task List'!D:D,A11,'Task List'!E:E,"Closed")
D11: =COUNTIFS('Task List'!D:D,A11,'Task List'!E:E,"In Progress")
E11: =B11-C11-D11
F11: =COUNTIFS('Task List'!D:D,A11,'Task List'!F:F,"Urgent")
G11: =COUNTIFS('Task List'!D:D,A11,'Task List'!F:F,"High")
H11: =COUNTIFS('Task List'!D:D,A11,'Task List'!F:F,"Medium")
I11: =COUNTIFS('Task List'!D:D,A11,'Task List'!F:F,"Low")
J11: =TEXTJOIN(", ",TRUE,FILTER('Task List'!B:B,('Task List'!D:D=A11)*('Task List'!E:E="In Progress"),"Không có"))
```

### 4.2. Công thức ALL-IN-ONE với QUERY (Nâng cao)

Bạn cũng có thể dùng công thức QUERY phức tạp hơn:

```
=QUERY('Task List'!A:H,"SELECT D, COUNT(D), SUM(CASE WHEN E='Finished' OR E='Closed' THEN 1 ELSE 0 END) WHERE D IS NOT NULL GROUP BY D")
```

---

## BƯỚC 5: Bảng Task sắp hết hạn

### 5.1. Lọc task theo Remaining Time

Tại vị trí **A25**, tạo bảng:

```
A25: TASK SẮP HẾT HẠN (trong 3 ngày tới)
A26: Task Name
B26: Assignee
C26: Due Date
D26: Remaining Time
E26: Status
F26: Priority

A27: =FILTER('Task List'!B:H, ('Task List'!H:H<=3)*('Task List'!H:H>0)*('Task List'!E:E<>"Finished")*('Task List'!E:E<>"Closed"), "Không có task sắp hết hạn")
```

### 5.2. Nếu Remaining Time là text (ví dụ: "2 days")

```
A27: =FILTER('Task List'!B:H, (VALUE(REGEXEXTRACT('Task List'!H:H,"\d+"))<=3)*('Task List'!E:E<>"Finished")*('Task List'!E:E<>"Closed"), "Không có")
```

### 5.3. Nếu dùng Due Date để tính

```
A27: =FILTER('Task List'!B:H, ('Task List'!G:G-TODAY()<=3)*('Task List'!G:G-TODAY()>=0)*('Task List'!E:E<>"Finished")*('Task List'!E:E<>"Closed"), "Không có task sắp hết hạn")
```

---

## BƯỚC 6: Thống kê theo Priority

### 6.1. Tạo bảng Priority

Tại vị trí **A40**, nhập:

```
A40: THỐNG KÊ THEO ĐỘ ƯU TIÊN
A41: Priority
B41: Số lượng
C41: Phần trăm
D41: Chưa hoàn thành

A42: Urgent
A43: High
A44: Medium
A45: Low
A46: TỔNG

B42: =COUNTIF('Task List'!F:F,"Urgent")
B43: =COUNTIF('Task List'!F:F,"High")
B44: =COUNTIF('Task List'!F:F,"Medium")
B45: =COUNTIF('Task List'!F:F,"Low")
B46: =SUM(B42:B45)

C42: =IF($B$46>0,B42/$B$46,0)
C43: =IF($B$46>0,B43/$B$46,0)
C44: =IF($B$46>0,B44/$B$46,0)
C45: =IF($B$46>0,B45/$B$46,0)
C46: =SUM(C42:C45)

D42: =COUNTIFS('Task List'!F:F,"Urgent",'Task List'!E:E,"<>Finished",'Task List'!E:E,"<>Closed")
D43: =COUNTIFS('Task List'!F:F,"High",'Task List'!E:E,"<>Finished",'Task List'!E:E,"<>Closed")
D44: =COUNTIFS('Task List'!F:F,"Medium",'Task List'!E:E,"<>Finished",'Task List'!E:E,"<>Closed")
D45: =COUNTIFS('Task List'!F:F,"Low",'Task List'!E:E,"<>Finished",'Task List'!E:E,"<>Closed")
D46: =SUM(D42:D45)
```

**Format cột C:** Chọn C42:C46 → Format → Number → Percent

---

## BƯỚC 7: Thêm Conditional Formatting (Định dạng có điều kiện)

### 7.1. Highlight task Urgent

1. Chọn cột Priority trong bảng chi tiết
2. Format → Conditional formatting
3. Format cells if: Text contains → "Urgent"
4. Formatting style: Background màu đỏ

### 7.2. Highlight task sắp hết hạn

1. Chọn cột Remaining Time
2. Format → Conditional formatting
3. Format cells if: Less than or equal to → 3
4. Formatting style: Background màu vàng/cam

---

## BƯỚC 8: Thêm Dashboard Cards (KPI)

Tại vị trí **H1**, tạo các KPI cards:

```
H1: 📊 TỔNG QUAN
H2: Tổng Task
I2: =COUNTA('Task List'!A2:A)

H3: ✅ Đã hoàn thành
I3: =COUNTIF('Task List'!E:E,"Finished")+COUNTIF('Task List'!E:E,"Closed")

H4: 🔄 Đang thực hiện
I4: =COUNTIF('Task List'!E:E,"In Progress")

H5: ⏳ Chưa bắt đầu
I5: =COUNTIF('Task List'!E:E,"To Do")

H6: 🚨 Task Urgent
I6: =COUNTIFS('Task List'!F:F,"Urgent",'Task List'!E:E,"<>Finished",'Task List'!E:E,"<>Closed")

H7: ⚠️ Sắp hết hạn
I7: =COUNTIFS('Task List'!H:H,"<=3",'Task List'!E:E,"<>Finished",'Task List'!E:E,"<>Closed")

H8: 📈 % Hoàn thành
I8: =I3/I2
```

---

## BƯỚC 9: Tạo biểu đồ cho Priority

1. Chọn vùng A41:B45
2. Insert → Chart
3. Chọn Chart type: Doughnut chart hoặc Pie chart
4. Thêm data labels để hiển thị phần trăm

---

## MẸO BỔ SUNG

### Tự động cập nhật danh sách Assignee

Dùng UNIQUE để lấy danh sách unique và ArrayFormula để áp dụng công thức cho tất cả:

```
=ARRAYFORMULA(IF(A11:A<>"",COUNTIF('Task List'!D:D,A11:A),""))
```

### Sắp xếp task theo độ ưu tiên và deadline

```
=SORT(FILTER('Task List'!A:H,'Task List'!E:E="In Progress"),6,FALSE,8,TRUE)
```

### Tạo Alert cho task quá hạn

```
=IF(COUNTIFS('Task List'!H:H,"<0",'Task List'!E:E,"<>Finished",'Task List'!E:E,"<>Closed")>0,"⚠️ CÓ "&COUNTIFS('Task List'!H:H,"<0",'Task List'!E:E,"<>Finished",'Task List'!E:E,"<>Closed")&" TASK QUÁ HẠN!","✅ Không có task quá hạn")
```

---

## LƯU Ý QUAN TRỌNG

1. **Điều chỉnh tên cột:** Thay đổi các tham chiếu cột (A, B, C, D, E, F, G, H) theo đúng vị trí trong sheet "Task List" của bạn.

2. **Điều chỉnh giá trị Status:** Nếu Status của bạn khác (ví dụ: "Done" thay vì "Finished"), hãy thay đổi trong các công thức.

3. **Điều chỉnh giá trị Priority:** Tương tự, thay đổi theo giá trị thực tế (ví dụ: "Critical" thay vì "Urgent").

4. **Realtime update:** Tất cả các công thức sẽ tự động cập nhật khi bạn thay đổi dữ liệu trong sheet "Task List".

5. **Tên sheet:** Nếu tên sheet của bạn có khoảng trắng hoặc ký tự đặc biệt, hãy dùng dấu nháy đơn: `'Task List'!A:A`

---

## CẦN HỖ TRỢ THÊM?

Nếu bạn cung cấp cho tôi:
- Screenshot cấu trúc sheet "Task List"
- Các giá trị Status và Priority thực tế

Tôi sẽ tạo công thức chính xác hơn cho bạn!
