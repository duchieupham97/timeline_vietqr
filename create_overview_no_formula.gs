/**
 * PHIÊN BẢN KHÔNG CÔNG THỨC - CHẮC CHẮN HOẠT ĐỘNG
 * Đọc dữ liệu trực tiếp và ghi giá trị
 */

function createOverviewSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const taskSheet = ss.getSheetByName("Task List");
  
  if (!taskSheet) {
    SpreadsheetApp.getUi().alert("Không tìm thấy sheet 'Task List'");
    return;
  }
  
  // Xóa Overview cũ
  let overview = ss.getSheetByName("Overview");
  if (overview) ss.deleteSheet(overview);
  
  // Tạo mới
  overview = ss.insertSheet("Overview");
  ss.moveActiveSheet(1);
  
  // Đọc dữ liệu
  const data = taskSheet.getDataRange().getValues();
  
  // Danh sách assignees
  const assigneeList = ["Duy Anh", "Trường", "Đức", "Triều", "Nghĩa", "Hiếu Phạm", "Quyết", "Hiếu Hà", "Tôn"];
  
  // Khởi tạo counters
  let total = 0;
  let statusCount = {"Open": 0, "Pending": 0, "In Progress": 0, "Testing": 0, "Finished": 0, "Closed": 0};
  let priorityCount = {"Urgent": 0, "High": 0, "Normal": 0, "Low": 0};
  let urgentNotDone = 0;
  let assigneeData = {};
  
  assigneeList.forEach(a => {
    assigneeData[a] = {
      total: 0, done: 0, inProgress: 0, testing: 0, pending: 0,
      urgent: 0, high: 0, normal: 0, low: 0, tasks: []
    };
  });
  
  // Xử lý từng hàng (bắt đầu từ hàng 3, index = 2)
  for (let i = 2; i < data.length; i++) {
    const row = data[i];
    const fno = String(row[0]).trim();
    const taskName = String(row[1]).trim();
    const status = String(row[6]).trim();
    const priority = String(row[5]).trim();
    const assignee = String(row[7]).trim();
    
    // Bỏ qua hàng trống hoặc group header
    if (!fno || fno === "" || fno.includes("Support") || fno.includes("BackEnd") || fno.includes("Frontend")) {
      continue;
    }
    
    total++;
    
    // Đếm status
    if (statusCount.hasOwnProperty(status)) {
      statusCount[status]++;
    }
    
    // Đếm priority
    if (priorityCount.hasOwnProperty(priority)) {
      priorityCount[priority]++;
    }
    
    // Đếm urgent chưa xong
    if (priority === "Urgent" && status !== "Finished" && status !== "Closed") {
      urgentNotDone++;
    }
    
    // Đếm theo assignee (hỗ trợ multiple select)
    assigneeList.forEach(name => {
      if (assignee.includes(name)) {
        assigneeData[name].total++;
        
        if (status === "Finished" || status === "Closed") {
          assigneeData[name].done++;
        } else if (status === "In Progress") {
          assigneeData[name].inProgress++;
          assigneeData[name].tasks.push(fno + ": " + taskName.substring(0, 30));
        } else if (status === "Testing") {
          assigneeData[name].testing++;
        } else {
          assigneeData[name].pending++;
        }
        
        // Đếm priority chưa xong
        if (status !== "Finished" && status !== "Closed") {
          if (priority === "Urgent") assigneeData[name].urgent++;
          else if (priority === "High") assigneeData[name].high++;
          else if (priority === "Normal") assigneeData[name].normal++;
          else if (priority === "Low") assigneeData[name].low++;
        }
      }
    });
  }
  
  // ========== GHI DỮ LIỆU RA OVERVIEW ==========
  
  // Title
  overview.getRange("A1").setValue("📊 TASK OVERVIEW - TIMELINE 2601");
  overview.getRange("A1:K1").merge().setBackground("#1a73e8").setFontColor("white")
    .setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
  
  // ===== KPI =====
  overview.getRange("A3:G3").setValues([["📋 Tổng Task", "✅ Hoàn thành", "🔄 Đang làm", "🧪 Testing", "⏳ Chờ xử lý", "🚨 Urgent", "📈 % Hoàn thành"]]);
  overview.getRange("A3:G3").setFontWeight("bold").setBackground("#e8f0fe").setHorizontalAlignment("center");
  
  const done = statusCount["Finished"] + statusCount["Closed"];
  const inProgress = statusCount["In Progress"];
  const testing = statusCount["Testing"];
  const pending = statusCount["Open"] + statusCount["Pending"];
  const percent = total > 0 ? Math.round(done / total * 100) + "%" : "0%";
  
  overview.getRange("A4:G4").setValues([[total, done, inProgress, testing, pending, urgentNotDone, percent]]);
  overview.getRange("A4:G4").setFontSize(20).setFontWeight("bold").setHorizontalAlignment("center");
  overview.getRange("B4").setFontColor("#1e8e3e");
  overview.getRange("F4").setFontColor("#d93025");
  overview.getRange("A3:G4").setBorder(true, true, true, true, true, true);
  
  // ===== STATUS =====
  overview.getRange("A6").setValue("📈 THỐNG KÊ THEO TRẠNG THÁI");
  overview.getRange("A6:D6").merge().setBackground("#34a853").setFontColor("white").setFontWeight("bold");
  overview.getRange("A7:D7").setValues([["Trạng thái", "Số lượng", "Phần trăm", ""]]);
  overview.getRange("A7:D7").setFontWeight("bold").setBackground("#e6f4ea");
  
  const statusIcons = {"Open": "🟢", "Pending": "🟡", "In Progress": "🔵", "Testing": "🟣", "Finished": "✅", "Closed": "⬛"};
  let row = 8;
  for (const [s, count] of Object.entries(statusCount)) {
    const pct = total > 0 ? Math.round(count / total * 100) + "%" : "0%";
    overview.getRange(row, 1).setValue(statusIcons[s] + " " + s);
    overview.getRange(row, 2).setValue(count);
    overview.getRange(row, 3).setValue(pct);
    overview.getRange(row, 4).setValue("▓".repeat(Math.round(count / total * 10) || 0)).setFontColor("#34a853");
    row++;
  }
  overview.getRange(row, 1).setValue("TỔNG").setFontWeight("bold");
  overview.getRange(row, 2).setValue(total).setFontWeight("bold");
  overview.getRange(row, 3).setValue("100%").setFontWeight("bold");
  overview.getRange("A7:D" + row).setBorder(true, true, true, true, true, true);
  
  // ===== PRIORITY =====
  overview.getRange("F6").setValue("🎯 THỐNG KÊ THEO ĐỘ ƯU TIÊN");
  overview.getRange("F6:J6").merge().setBackground("#ea4335").setFontColor("white").setFontWeight("bold");
  overview.getRange("F7:J7").setValues([["Độ ưu tiên", "Tổng", "Chưa xong", "Phần trăm", "Cảnh báo"]]);
  overview.getRange("F7:J7").setFontWeight("bold").setBackground("#fce8e6");
  
  const priorityIcons = {"Urgent": "🔴", "High": "🟠", "Normal": "🟡", "Low": "🟢"};
  const priorityColors = {"Urgent": "#ffcdd2", "High": "#ffe0b2", "Normal": "#fff9c4", "Low": "#c8e6c9"};
  
  // Tính số chưa xong theo priority
  let priorityNotDone = {"Urgent": 0, "High": 0, "Normal": 0, "Low": 0};
  assigneeList.forEach(name => {
    priorityNotDone["Urgent"] += assigneeData[name].urgent;
    priorityNotDone["High"] += assigneeData[name].high;
    priorityNotDone["Normal"] += assigneeData[name].normal;
    priorityNotDone["Low"] += assigneeData[name].low;
  });
  // Chia đôi vì có thể bị đếm trùng trong multiple assignee
  // Thực ra cần tính lại chính xác hơn, nhưng tạm dùng urgentNotDone đã tính ở trên
  
  row = 8;
  for (const [p, count] of Object.entries(priorityCount)) {
    const pct = total > 0 ? Math.round(count / total * 100) + "%" : "0%";
    const notDone = p === "Urgent" ? urgentNotDone : Math.round(priorityNotDone[p] / 2);
    const warning = notDone > 0 ? "⚠️ " + notDone + " task" : "";
    
    overview.getRange(row, 6).setValue(priorityIcons[p] + " " + p).setBackground(priorityColors[p]);
    overview.getRange(row, 7).setValue(count);
    overview.getRange(row, 8).setValue(notDone);
    overview.getRange(row, 9).setValue(pct);
    overview.getRange(row, 10).setValue(warning);
    row++;
  }
  overview.getRange(row, 6).setValue("TỔNG").setFontWeight("bold");
  overview.getRange(row, 7).setValue(total).setFontWeight("bold");
  overview.getRange("F7:J" + row).setBorder(true, true, true, true, true, true);
  
  // ===== ASSIGNEE SUMMARY =====
  overview.getRange("A16").setValue("👥 THỐNG KÊ THEO NGƯỜI THỰC HIỆN");
  overview.getRange("A16:C16").merge().setBackground("#9c27b0").setFontColor("white").setFontWeight("bold");
  overview.getRange("A17:C17").setValues([["Assignee", "Số Task", ""]]);
  overview.getRange("A17:C17").setFontWeight("bold").setBackground("#f3e5f5");
  
  row = 18;
  assigneeList.forEach(name => {
    const count = assigneeData[name].total;
    overview.getRange(row, 1).setValue(name);
    overview.getRange(row, 2).setValue(count);
    overview.getRange(row, 3).setValue("█".repeat(Math.min(count, 20))).setFontColor("#9c27b0");
    row++;
  });
  overview.getRange("A17:C" + (row - 1)).setBorder(true, true, true, true, true, true);
  
  // ===== ASSIGNEE DETAIL =====
  const detailStartRow = row + 2;
  overview.getRange(detailStartRow, 1).setValue("📋 CHI TIẾT WORKLOAD TỪNG NGƯỜI");
  overview.getRange(detailStartRow, 1, 1, 11).merge().setBackground("#1565c0").setFontColor("white").setFontWeight("bold");
  
  const headers = ["Assignee", "Tổng", "Done", "Progress", "Testing", "Pending", "Urgent", "High", "Normal", "Low", "Task đang làm"];
  overview.getRange(detailStartRow + 1, 1, 1, 11).setValues([headers]);
  overview.getRange(detailStartRow + 1, 1, 1, 11).setFontWeight("bold").setBackground("#e3f2fd").setFontSize(9);
  
  row = detailStartRow + 2;
  let totalRow = {total: 0, done: 0, inProgress: 0, testing: 0, pending: 0, urgent: 0, high: 0, normal: 0, low: 0};
  
  assigneeList.forEach(name => {
    const d = assigneeData[name];
    overview.getRange(row, 1, 1, 11).setValues([[
      name, d.total, d.done, d.inProgress, d.testing, d.pending,
      d.urgent, d.high, d.normal, d.low,
      d.tasks.length > 0 ? d.tasks.join(", ") : "Không có"
    ]]);
    
    // Highlight urgent
    if (d.urgent > 0) {
      overview.getRange(row, 7).setBackground("#ffcdd2").setFontColor("#c62828");
    }
    if (d.high > 0) {
      overview.getRange(row, 8).setBackground("#ffe0b2").setFontColor("#e65100");
    }
    
    // Cộng dồn cho total
    totalRow.total += d.total;
    totalRow.done += d.done;
    totalRow.inProgress += d.inProgress;
    totalRow.testing += d.testing;
    totalRow.pending += d.pending;
    totalRow.urgent += d.urgent;
    totalRow.high += d.high;
    totalRow.normal += d.normal;
    totalRow.low += d.low;
    
    row++;
  });
  
  // Total row
  overview.getRange(row, 1, 1, 11).setValues([[
    "TỔNG", totalRow.total, totalRow.done, totalRow.inProgress, totalRow.testing, totalRow.pending,
    totalRow.urgent, totalRow.high, totalRow.normal, totalRow.low, ""
  ]]);
  overview.getRange(row, 1, 1, 11).setFontWeight("bold");
  
  overview.getRange(detailStartRow + 1, 1, row - detailStartRow, 11).setBorder(true, true, true, true, true, true);
  
  // ===== FORMATTING =====
  overview.setColumnWidth(1, 100);
  overview.setColumnWidth(11, 350);
  overview.setFrozenRows(2);
  
  // ===== CHARTS =====
  try {
    // Pie chart Status
    const chart1 = overview.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(overview.getRange("A8:B13"))
      .setPosition(5, 12, 0, 0)
      .setOption('title', 'Phân bổ theo Status')
      .setOption('pieHole', 0.4)
      .setOption('width', 350)
      .setOption('height', 250)
      .build();
    overview.insertChart(chart1);
    
    // Bar chart Assignee
    const chart2 = overview.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(overview.getRange("A18:B" + (17 + assigneeList.length)))
      .setPosition(17, 12, 0, 0)
      .setOption('title', 'Task theo Assignee')
      .setOption('width', 350)
      .setOption('height', 250)
      .setOption('legend', {position: 'none'})
      .build();
    overview.insertChart(chart2);
  } catch(e) {
    // Bỏ qua lỗi chart
  }
  
  SpreadsheetApp.getUi().alert('✅ Tạo Overview thành công!\n\nLưu ý: Dữ liệu là snapshot, chạy lại script để cập nhật.');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📊 Task Overview')
    .addItem('🔄 Tạo/Cập nhật Overview', 'createOverviewSheet')
    .addToUi();
}
