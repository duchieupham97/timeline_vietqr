/**
 * SCRIPT TEST ĐƠN GIẢN NHẤT
 * Chạy function testFormulas để kiểm tra
 */

function testFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Test 1: Đọc dữ liệu từ Task List
  const taskSheet = ss.getSheetByName("Task List");
  if (!taskSheet) {
    ui.alert("Không tìm thấy sheet 'Task List'");
    return;
  }
  
  // Lấy tất cả dữ liệu
  const data = taskSheet.getDataRange().getValues();
  
  let info = "📊 PHÂN TÍCH DỮ LIỆU:\n\n";
  info += "Tổng số hàng: " + data.length + "\n\n";
  
  // Đếm Status
  let statusCount = {};
  let priorityCount = {};
  let assigneeCount = {};
  
  for (let i = 1; i < data.length; i++) { // Bỏ qua header (hàng 0)
    const row = data[i];
    const colG = row[6]; // Status (cột G = index 6)
    const colF = row[5]; // Priority (cột F = index 5)
    const colH = row[7]; // Assignee (cột H = index 7)
    
    // Đếm Status
    if (colG && colG !== "") {
      statusCount[colG] = (statusCount[colG] || 0) + 1;
    }
    
    // Đếm Priority
    if (colF && colF !== "") {
      priorityCount[colF] = (priorityCount[colF] || 0) + 1;
    }
    
    // Đếm Assignee (tách multiple)
    if (colH && colH !== "") {
      const assignees = String(colH).split(",");
      assignees.forEach(a => {
        const name = a.trim();
        if (name) assigneeCount[name] = (assigneeCount[name] || 0) + 1;
      });
    }
  }
  
  info += "📌 STATUS:\n";
  for (const [k, v] of Object.entries(statusCount)) {
    info += "- " + k + ": " + v + "\n";
  }
  
  info += "\n📌 PRIORITY:\n";
  for (const [k, v] of Object.entries(priorityCount)) {
    info += "- " + k + ": " + v + "\n";
  }
  
  info += "\n📌 ASSIGNEE:\n";
  for (const [k, v] of Object.entries(assigneeCount)) {
    info += "- " + k + ": " + v + "\n";
  }
  
  ui.alert("KẾT QUẢ TEST", info, ui.ButtonSet.OK);
  Logger.log(info);
}

/**
 * Tạo Overview với công thức CỰC KỲ ĐƠN GIẢN
 */
function createSimpleOverview() {
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
  
  // ===== ĐỌC DỮ LIỆU TRỰC TIẾP =====
  const data = taskSheet.getDataRange().getValues();
  
  // Khởi tạo counters
  let total = 0;
  let statusCount = {"Open": 0, "Pending": 0, "In Progress": 0, "Testing": 0, "Finished": 0, "Closed": 0};
  let priorityCount = {"Urgent": 0, "High": 0, "Normal": 0, "Low": 0};
  let assigneeData = {};
  
  // Danh sách assignees
  const assigneeList = ["Duy Anh", "Trường", "Đức", "Triều", "Nghĩa", "Hiếu Phạm", "Quyết", "Hiếu Hà", "Tôn"];
  assigneeList.forEach(a => {
    assigneeData[a] = {total: 0, done: 0, inProgress: 0, testing: 0, pending: 0, urgent: 0, high: 0, normal: 0, low: 0, tasks: []};
  });
  
  // Xử lý từng hàng
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const fno = row[0];      // A - FNo
    const taskName = row[1]; // B - Functional
    const status = row[6];   // G - Status
    const priority = row[5]; // F - Priority
    const assignee = row[7]; // H - Assignee
    
    // Bỏ qua hàng trống hoặc group header
    if (!fno || fno === "" || String(fno).includes("Support") || String(fno).includes("BackEnd") || String(fno).includes("Frontend")) {
      continue;
    }
    
    total++;
    
    // Đếm status
    if (status && statusCount.hasOwnProperty(status)) {
      statusCount[status]++;
    }
    
    // Đếm priority
    if (priority && priorityCount.hasOwnProperty(priority)) {
      priorityCount[priority]++;
    }
    
    // Đếm theo assignee
    if (assignee) {
      assigneeList.forEach(name => {
        if (String(assignee).includes(name)) {
          assigneeData[name].total++;
          
          if (status === "Finished" || status === "Closed") {
            assigneeData[name].done++;
          } else if (status === "In Progress") {
            assigneeData[name].inProgress++;
            assigneeData[name].tasks.push(fno + ": " + taskName);
          } else if (status === "Testing") {
            assigneeData[name].testing++;
          } else {
            assigneeData[name].pending++;
          }
          
          if (status !== "Finished" && status !== "Closed") {
            if (priority === "Urgent") assigneeData[name].urgent++;
            if (priority === "High") assigneeData[name].high++;
            if (priority === "Normal") assigneeData[name].normal++;
            if (priority === "Low") assigneeData[name].low++;
          }
        }
      });
    }
  }
  
  // ===== GHI DỮ LIỆU RA OVERVIEW =====
  
  // Title
  overview.getRange("A1").setValue("📊 TASK OVERVIEW - TIMELINE 2601");
  overview.getRange("A1:K1").merge().setBackground("#1a73e8").setFontColor("white").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
  
  // KPI
  overview.getRange("A3:G3").setValues([["📋 Tổng Task", "✅ Hoàn thành", "🔄 Đang làm", "🧪 Testing", "⏳ Chờ xử lý", "🚨 Urgent", "📈 % Hoàn thành"]]);
  overview.getRange("A3:G3").setFontWeight("bold").setBackground("#e8f0fe");
  
  const done = statusCount["Finished"] + statusCount["Closed"];
  const inProgress = statusCount["In Progress"];
  const testing = statusCount["Testing"];
  const pending = statusCount["Open"] + statusCount["Pending"];
  const urgentNotDone = priorityCount["Urgent"]; // Cần tính lại chính xác hơn
  const percent = total > 0 ? (done / total * 100).toFixed(0) + "%" : "0%";
  
  overview.getRange("A4:G4").setValues([[total, done, inProgress, testing, pending, urgentNotDone, percent]]);
  overview.getRange("A4:G4").setFontSize(20).setFontWeight("bold").setHorizontalAlignment("center");
  overview.getRange("A3:G4").setBorder(true, true, true, true, true, true);
  
  // Status
  overview.getRange("A6").setValue("📈 THỐNG KÊ THEO TRẠNG THÁI");
  overview.getRange("A6:C6").merge().setBackground("#34a853").setFontColor("white").setFontWeight("bold");
  overview.getRange("A7:C7").setValues([["Trạng thái", "Số lượng", "Phần trăm"]]).setFontWeight("bold").setBackground("#e6f4ea");
  
  let row = 8;
  const statusIcons = {"Open": "🟢", "Pending": "🟡", "In Progress": "🔵", "Testing": "🟣", "Finished": "✅", "Closed": "⬛"};
  for (const [s, count] of Object.entries(statusCount)) {
    const pct = total > 0 ? (count / total * 100).toFixed(0) + "%" : "0%";
    overview.getRange(row, 1).setValue((statusIcons[s] || "") + " " + s);
    overview.getRange(row, 2).setValue(count);
    overview.getRange(row, 3).setValue(pct);
    row++;
  }
  overview.getRange("A7:C" + (row - 1)).setBorder(true, true, true, true, true, true);
  
  // Priority
  overview.getRange("E6").setValue("🎯 THỐNG KÊ THEO ĐỘ ƯU TIÊN");
  overview.getRange("E6:G6").merge().setBackground("#ea4335").setFontColor("white").setFontWeight("bold");
  overview.getRange("E7:G7").setValues([["Độ ưu tiên", "Số lượng", "Phần trăm"]]).setFontWeight("bold").setBackground("#fce8e6");
  
  row = 8;
  const priorityIcons = {"Urgent": "🔴", "High": "🟠", "Normal": "🟡", "Low": "🟢"};
  const priorityColors = {"Urgent": "#ffcdd2", "High": "#ffe0b2", "Normal": "#fff9c4", "Low": "#c8e6c9"};
  for (const [p, count] of Object.entries(priorityCount)) {
    const pct = total > 0 ? (count / total * 100).toFixed(0) + "%" : "0%";
    overview.getRange(row, 5).setValue((priorityIcons[p] || "") + " " + p).setBackground(priorityColors[p]);
    overview.getRange(row, 6).setValue(count);
    overview.getRange(row, 7).setValue(pct);
    row++;
  }
  overview.getRange("E7:G" + (row - 1)).setBorder(true, true, true, true, true, true);
  
  // Assignee Summary
  overview.getRange("A15").setValue("👥 THỐNG KÊ THEO NGƯỜI THỰC HIỆN");
  overview.getRange("A15:C15").merge().setBackground("#9c27b0").setFontColor("white").setFontWeight("bold");
  overview.getRange("A16:C16").setValues([["Assignee", "Số Task", ""]]).setFontWeight("bold").setBackground("#f3e5f5");
  
  row = 17;
  assigneeList.forEach(name => {
    const count = assigneeData[name].total;
    overview.getRange(row, 1).setValue(name);
    overview.getRange(row, 2).setValue(count);
    overview.getRange(row, 3).setValue("█".repeat(count)).setFontColor("#9c27b0");
    row++;
  });
  overview.getRange("A16:C" + (row - 1)).setBorder(true, true, true, true, true, true);
  
  // Assignee Detail
  const detailStartRow = row + 2;
  overview.getRange(detailStartRow, 1).setValue("📋 CHI TIẾT WORKLOAD TỪNG NGƯỜI");
  overview.getRange(detailStartRow, 1, 1, 11).merge().setBackground("#1565c0").setFontColor("white").setFontWeight("bold");
  
  const headers = ["Assignee", "Tổng", "Done", "Progress", "Testing", "Pending", "Urgent", "High", "Normal", "Low", "Task đang làm"];
  overview.getRange(detailStartRow + 1, 1, 1, 11).setValues([headers]).setFontWeight("bold").setBackground("#e3f2fd");
  
  row = detailStartRow + 2;
  assigneeList.forEach(name => {
    const d = assigneeData[name];
    overview.getRange(row, 1).setValue(name);
    overview.getRange(row, 2).setValue(d.total);
    overview.getRange(row, 3).setValue(d.done);
    overview.getRange(row, 4).setValue(d.inProgress);
    overview.getRange(row, 5).setValue(d.testing);
    overview.getRange(row, 6).setValue(d.pending);
    overview.getRange(row, 7).setValue(d.urgent);
    overview.getRange(row, 8).setValue(d.high);
    overview.getRange(row, 9).setValue(d.normal);
    overview.getRange(row, 10).setValue(d.low);
    overview.getRange(row, 11).setValue(d.tasks.join(", ") || "Không có");
    
    // Highlight urgent
    if (d.urgent > 0) {
      overview.getRange(row, 7).setBackground("#ffcdd2").setFontColor("#c62828");
    }
    if (d.high > 0) {
      overview.getRange(row, 8).setBackground("#ffe0b2").setFontColor("#e65100");
    }
    row++;
  });
  
  overview.getRange(detailStartRow + 1, 1, row - detailStartRow, 11).setBorder(true, true, true, true, true, true);
  
  // Column widths
  overview.setColumnWidth(1, 100);
  overview.setColumnWidth(11, 300);
  
  SpreadsheetApp.getUi().alert("✅ Tạo Overview thành công!\n\nDữ liệu được tính toán trực tiếp (không dùng công thức).\nĐể cập nhật, chạy lại script.");
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu("📊 Overview")
    .addItem("🔍 Test đọc dữ liệu", "testFormulas")
    .addItem("🔄 Tạo Overview (không công thức)", "createSimpleOverview")
    .addToUi();
}
