/**
 * HƯỚNG DẪN SỬ DỤNG:
 * 1. Mở Google Sheet của bạn.
 * 2. Vào menu "Extensions" (Tiện ích mở rộng) > "Apps Script".
 * 3. Xóa mọi mã trong đó và dán toàn bộ mã này vào.
 * 4. Chỉnh sửa phần "CẤU HÌNH" bên dưới nếu tên cột của bạn khác.
 * 5. Lưu lại và chọn hàm "createOverviewDashboard" để chạy.
 * 6. Cấp quyền truy cập khi được hỏi.
 */

// --- CẤU HÌNH (Vui lòng kiểm tra và sửa tên cột cho khớp với file của bạn) ---
const CONFIG = {
  SOURCE_SHEET_NAME: "Task List", // Tên sheet chứa dữ liệu
  DEST_SHEET_NAME: "Overview",    // Tên sheet thống kê sẽ được tạo
  COLUMNS: {
    TASK_NAME: "Task Name",       // Tên cột Tên Task
    ASSIGNEE: "Assignee",         // Tên cột Người được giao
    STATUS: "Status",             // Tên cột Trạng thái
    PRIORITY: "Priority",         // Tên cột Độ ưu tiên
    REMAINING: "Remaining Time"   // Tên cột Thời gian còn lại
  },
  STATUS_KEYWORDS: {
    DONE: ["Finished", "Closed", "Done", "Hoàn thành"], // Các trạng thái coi là Đã xong
    IN_PROGRESS: ["In Progress", "Doing", "Đang làm"]   // Các trạng thái coi là Đang làm
  },
  PRIORITY_KEYWORDS: {
    URGENT: ["Urgent", "Khẩn cấp"],
    HIGH: ["High", "Cao"],
    MEDIUM: ["Medium", "Vừa", "Normal", "Trung bình"],
    LOW: ["Low", "Thấp"]
  }
};

function createOverviewDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(CONFIG.SOURCE_SHEET_NAME);
  
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert(`Không tìm thấy sheet "${CONFIG.SOURCE_SHEET_NAME}". Vui lòng kiểm tra lại tên sheet.`);
    return;
  }

  // 1. Lấy dữ liệu
  const dataRange = sourceSheet.getDataRange();
  const values = dataRange.getValues();
  const headers = values[0];
  const data = values.slice(1);

  // Map column headers to indices
  const colMap = {};
  for (const [key, name] of Object.entries(CONFIG.COLUMNS)) {
    const index = headers.indexOf(name);
    if (index === -1) {
      // Thử tìm kiếm không phân biệt hoa thường nếu không tìm thấy chính xác
      const indexCaseInsensitive = headers.findIndex(h => h.toString().toLowerCase() === name.toLowerCase());
      if (indexCaseInsensitive === -1) {
        SpreadsheetApp.getUi().alert(`Không tìm thấy cột "${name}". Vui lòng kiểm tra lại cấu hình.`);
        return;
      }
      colMap[key] = indexCaseInsensitive;
    } else {
      colMap[key] = index;
    }
  }

  // 2. Xử lý dữ liệu
  const stats = processData(data, colMap);

  // 3. Chuẩn bị Sheet Overview
  let destSheet = ss.getSheetByName(CONFIG.DEST_SHEET_NAME);
  if (destSheet) {
    destSheet.clear();
    const charts = destSheet.getCharts();
    charts.forEach(c => destSheet.removeChart(c));
  } else {
    destSheet = ss.insertSheet(CONFIG.DEST_SHEET_NAME);
  }

  // 4. Vẽ giao diện và biểu đồ
  drawDashboard(destSheet, stats);
}

function processData(data, colMap) {
  const stats = {
    totalTasks: 0,
    statusCounts: {},
    priorityCounts: {},
    assigneeStats: {},
    expiringTasks: [] // Logic này sẽ cần điều chỉnh tùy format thời gian
  };

  data.forEach(row => {
    const taskName = row[colMap.TASK_NAME];
    const assignee = row[colMap.ASSIGNEE] || "Unassigned";
    const status = row[colMap.STATUS] || "Unknown";
    const priority = row[colMap.PRIORITY] || "None";
    const remaining = row[colMap.REMAINING];

    if (!taskName) return; // Bỏ qua dòng trống

    stats.totalTasks++;

    // Thống kê Status
    stats.statusCounts[status] = (stats.statusCounts[status] || 0) + 1;

    // Thống kê Priority
    stats.priorityCounts[priority] = (stats.priorityCounts[priority] || 0) + 1;

    // Thống kê Assignee
    if (!stats.assigneeStats[assignee]) {
      stats.assigneeStats[assignee] = {
        total: 0,
        done: 0,
        inProgress: 0,
        priorities: { Urgent: 0, High: 0, Medium: 0, Low: 0, Other: 0 },
        currentTasks: []
      };
    }
    const userStat = stats.assigneeStats[assignee];
    userStat.total++;

    // Check Done
    if (CONFIG.STATUS_KEYWORDS.DONE.includes(status)) {
      userStat.done++;
    }
    // Check In Progress
    else if (CONFIG.STATUS_KEYWORDS.IN_PROGRESS.includes(status)) {
      userStat.inProgress++;
      userStat.currentTasks.push(`${taskName} (${priority})`);
    }

    // Count Priority per Assignee
    if (CONFIG.PRIORITY_KEYWORDS.URGENT.includes(priority)) userStat.priorities.Urgent++;
    else if (CONFIG.PRIORITY_KEYWORDS.HIGH.includes(priority)) userStat.priorities.High++;
    else if (CONFIG.PRIORITY_KEYWORDS.MEDIUM.includes(priority)) userStat.priorities.Medium++;
    else if (CONFIG.PRIORITY_KEYWORDS.LOW.includes(priority)) userStat.priorities.Low++;
    else userStat.priorities.Other++;

    // Check Expiring (Logic đơn giản: Nếu có chữ "days" hoặc số nhỏ)
    // Bạn có thể tùy chỉnh logic này
    if (remaining) {
       stats.expiringTasks.push({ taskName, assignee, remaining, priority, status });
    }
  });

  return stats;
}

function drawDashboard(sheet, stats) {
  let currentRow = 1;
  
  // Style cơ bản
  const headerStyle = SpreadsheetApp.newTextStyle().setBold(true).setFontSize(12).build();
  
  // --- PHẦN 1: THỐNG KÊ TỔNG QUAN (STATUS & PRIORITY) ---
  sheet.getRange(currentRow, 1).setValue("THỐNG KÊ TRẠNG THÁI (STATUS)").setTextStyle(headerStyle);
  sheet.getRange(currentRow, 4).setValue("THỐNG KÊ ĐỘ ƯU TIÊN (PRIORITY)").setTextStyle(headerStyle);
  currentRow++;

  // Bảng Status
  const statusData = [["Status", "Số lượng", "Tỷ lệ (%)"]];
  for (const [status, count] of Object.entries(stats.statusCounts)) {
    statusData.push([status, count, (count / stats.totalTasks * 100).toFixed(1) + "%"]);
  }
  sheet.getRange(currentRow, 1, statusData.length, 3).setValues(statusData).setBorder(true, true, true, true, true, true);
  
  // Biểu đồ Status (Pie Chart)
  const statusChart = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange(currentRow + 1, 1, statusData.length - 1, 2))
    .setPosition(currentRow, 1, 0, 0) // Vẽ đè lên bảng một chút hoặc bên cạnh? Để bên dưới.
    .setOption('title', 'Tỷ lệ trạng thái Task')
    .build();
  // Chúng ta sẽ đặt biểu đồ sau
  
  // Bảng Priority (bên cạnh)
  const priorityData = [["Priority", "Số lượng", "Tỷ lệ (%)"]];
  for (const [prio, count] of Object.entries(stats.priorityCounts)) {
    priorityData.push([prio, count, (count / stats.totalTasks * 100).toFixed(1) + "%"]);
  }
  sheet.getRange(currentRow, 4, priorityData.length, 3).setValues(priorityData).setBorder(true, true, true, true, true, true);
  
  // Vẽ biểu đồ Status xuống dưới
  sheet.insertChart(statusChart);
  const chartHeight = 15; // rows
  statusChart.setPosition(currentRow + Math.max(statusData.length, priorityData.length) + 1, 1, 0, 0);
  sheet.updateChart(statusChart);

  currentRow += Math.max(statusData.length, priorityData.length) + 18; // Dịch chuyển xuống dưới biểu đồ

  // --- PHẦN 2: THỐNG KÊ THEO ASSIGNEE ---
  sheet.getRange(currentRow, 1).setValue("CHI TIẾT THEO NHÂN SỰ (ASSIGNEE)").setTextStyle(headerStyle);
  currentRow++;

  const assigneeHeaders = ["Assignee", "Tổng Task", "Đã xong", "Đang làm", "Đang làm gì?", "Urgent", "High", "Medium", "Low"];
  const assigneeData = [assigneeHeaders];
  
  for (const [name, s] of Object.entries(stats.assigneeStats)) {
    assigneeData.push([
      name,
      s.total,
      s.done,
      s.inProgress,
      s.currentTasks.join(",\n"), // Xuống dòng nếu nhiều task
      s.priorities.Urgent,
      s.priorities.High,
      s.priorities.Medium,
      s.priorities.Low
    ]);
  }

  const assigneeRange = sheet.getRange(currentRow, 1, assigneeData.length, assigneeHeaders.length);
  assigneeRange.setValues(assigneeData);
  assigneeRange.setBorder(true, true, true, true, true, true);
  assigneeRange.setVerticalAlignment("top"); // Căn trên để dễ đọc task đang làm
  
  // Format cột "Đang làm gì?" cho wrap text
  sheet.getRange(currentRow + 1, 5, assigneeData.length - 1, 1).setWrap(true);

  // Biểu đồ Assignee (Bar Chart)
  const assigneeChart = sheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(sheet.getRange(currentRow, 1, assigneeData.length, 2)) // Cột Tên và Tổng
    .setPosition(currentRow, 10, 0, 0) // Đặt bên phải bảng
    .setOption('title', 'Số lượng Task theo Nhân sự')
    .build();
  sheet.insertChart(assigneeChart);

  currentRow += assigneeData.length + 2;

  // --- PHẦN 3: TASK SẮP HẾT HẠN (EXPIRING) ---
  sheet.getRange(currentRow, 1).setValue("DANH SÁCH TASK CẦN LƯU Ý (REMAINING TIME)").setTextStyle(headerStyle);
  currentRow++;

  // Sắp xếp theo Remaining Time (cái này hơi khó nếu format không chuẩn, tạm thời list ra hết hoặc sort string)
  // Giả sử user muốn xem list này
  const expiringHeaders = ["Task Name", "Assignee", "Status", "Priority", "Remaining Time"];
  const expiringData = [expiringHeaders];
  
  // Lọc sơ bộ hoặc lấy hết
  stats.expiringTasks.forEach(t => {
      expiringData.push([t.taskName, t.assignee, t.status, t.priority, t.remaining]);
  });

  if (expiringData.length > 1) {
    sheet.getRange(currentRow, 1, expiringData.length, expiringHeaders.length).setValues(expiringData).setBorder(true, true, true, true, true, true);
  } else {
    sheet.getRange(currentRow, 1).setValue("Không có dữ liệu thời gian.");
  }

  // Auto resize columns
  sheet.autoResizeColumns(1, 10);
}
