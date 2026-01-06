/**
 * Google Apps Script để tạo Sheet "Overview" 
 * cho Task List của Team Timeline 2601
 * 
 * HƯỚNG DẪN SỬ DỤNG:
 * 1. Mở Google Sheet: https://docs.google.com/spreadsheets/d/1N_f8TaqdUu1RKuKSFk0essrEQ95fdUbR5t4mvnsZj8c/edit
 * 2. Vào Extensions → Apps Script
 * 3. Xóa code mặc định và paste toàn bộ code này
 * 4. Nhấn Run (▶️) và chọn function "createOverviewSheet"
 * 5. Cấp quyền khi được yêu cầu
 */

// ==================== CẤU HÌNH THEO SHEET CỦA BẠN ====================
const CONFIG = {
  taskListSheetName: "Task List",
  overviewSheetName: "Overview",
  
  // Vị trí cột trong Task List (A=1, B=2, ...)
  columns: {
    taskId: 1,         // A - FNo.
    taskName: 2,       // B - Functional
    startDate: 3,      // C - Start Date
    endDate: 4,        // D - End Date
    remainingTime: 5,  // E - Remaining Time (hh:mm)
    priority: 6,       // F - Priority
    status: 7,         // G - Status
    assignee: 8,       // H - Assignee (MULTIPLE SELECT)
    tester: 9,         // I - Tester
    progress: 10,      // J - Progress (%)
    note: 11           // K - Reference/Note
  },
  
  // Giá trị Status
  status: {
    done: ["Finished", "Closed"],
    inProgress: ["In Progress"],
    testing: ["Testing"],
    pending: ["Open", "Pending"]
  },
  
  // Giá trị Priority
  priority: {
    urgent: "Urgent",
    high: "High",
    normal: "Normal",
    low: "Low"
  },
  
  // Danh sách Assignee (để thống kê chính xác với multiple select)
  assignees: ["Duy Anh", "Trường", "Đức", "Triều", "Nghĩa", "Hiếu Phạm", "Quyết", "Hiếu Hà", "Tôn"],
  
  // Hàng bắt đầu dữ liệu (bỏ qua header)
  dataStartRow: 3,
  
  // Các hàng là header group (Customer Support, BackEnd) - sẽ bỏ qua
  groupHeaderRows: [3, 23] // Điều chỉnh nếu cần
};

// ==================== MAIN FUNCTION ====================
function createOverviewSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Xóa sheet Overview cũ nếu có
  let overviewSheet = ss.getSheetByName(CONFIG.overviewSheetName);
  if (overviewSheet) {
    ss.deleteSheet(overviewSheet);
  }
  
  // Tạo sheet Overview mới
  overviewSheet = ss.insertSheet(CONFIG.overviewSheetName);
  ss.setActiveSheet(overviewSheet);
  ss.moveActiveSheet(1);
  
  // Thiết lập các phần
  setupDashboardKPIs(overviewSheet);
  setupStatusStats(overviewSheet);
  setupPriorityStats(overviewSheet);
  setupAssigneeOverview(overviewSheet);
  setupUpcomingDeadlines(overviewSheet);
  setupAssigneeDetailTable(overviewSheet);
  
  // Format
  formatOverviewSheet(overviewSheet);
  
  // Tạo biểu đồ
  createCharts(overviewSheet);
  
  SpreadsheetApp.getUi().alert('✅ Sheet "Overview" đã được tạo thành công!\n\nDữ liệu sẽ tự động cập nhật realtime khi bạn thay đổi Task List.');
}

// ==================== HELPER: Get column letter ====================
function getColLetter(colNum) {
  let letter = '';
  while (colNum > 0) {
    let mod = (colNum - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    colNum = Math.floor((colNum - mod) / 26);
  }
  return letter;
}

// ==================== DASHBOARD KPIs ====================
function setupDashboardKPIs(sheet) {
  const tl = CONFIG.taskListSheetName;
  const statusCol = getColLetter(CONFIG.columns.status);
  const priorityCol = getColLetter(CONFIG.columns.priority);
  const taskIdCol = getColLetter(CONFIG.columns.taskId);
  
  // Title
  sheet.getRange("A1").setValue("📊 TASK OVERVIEW - TIMELINE 2601").setFontSize(18).setFontWeight("bold");
  sheet.getRange("A1:J1").merge().setBackground("#1a73e8").setFontColor("white").setHorizontalAlignment("center");
  
  // KPI Row
  sheet.getRange("A3").setValue("📋 Tổng Task");
  sheet.getRange("B3").setValue("✅ Hoàn thành");
  sheet.getRange("C3").setValue("🔄 Đang làm");
  sheet.getRange("D3").setValue("🧪 Testing");
  sheet.getRange("E3").setValue("⏳ Chờ xử lý");
  sheet.getRange("F3").setValue("🚨 Urgent");
  sheet.getRange("G3").setValue("📈 % Hoàn thành");
  sheet.getRange("A3:G3").setFontWeight("bold").setBackground("#e8f0fe").setHorizontalAlignment("center");
  
  // KPI Values - Đếm task có FNo. không rỗng và không phải header group
  sheet.getRange("A4").setFormula(`=COUNTIF('${tl}'!${taskIdCol}:${taskIdCol},"F*")`);
  sheet.getRange("B4").setFormula(`=COUNTIFS('${tl}'!${statusCol}:${statusCol},"Finished")+COUNTIFS('${tl}'!${statusCol}:${statusCol},"Closed")`);
  sheet.getRange("C4").setFormula(`=COUNTIF('${tl}'!${statusCol}:${statusCol},"In Progress")`);
  sheet.getRange("D4").setFormula(`=COUNTIF('${tl}'!${statusCol}:${statusCol},"Testing")`);
  sheet.getRange("E4").setFormula(`=COUNTIF('${tl}'!${statusCol}:${statusCol},"Open")+COUNTIF('${tl}'!${statusCol}:${statusCol},"Pending")`);
  sheet.getRange("F4").setFormula(`=COUNTIFS('${tl}'!${priorityCol}:${priorityCol},"Urgent",'${tl}'!${statusCol}:${statusCol},"<>Finished",'${tl}'!${statusCol}:${statusCol},"<>Closed")`);
  sheet.getRange("G4").setFormula(`=IF(A4>0,B4/A4,0)`);
  
  // Format
  sheet.getRange("A4:F4").setFontSize(24).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("G4").setFontSize(24).setFontWeight("bold").setHorizontalAlignment("center").setNumberFormat("0.0%");
  sheet.getRange("F4").setFontColor("#d93025"); // Red for urgent
  sheet.getRange("B4").setFontColor("#1e8e3e"); // Green for done
  sheet.getRange("A3:G4").setBorder(true, true, true, true, true, true);
}

// ==================== STATUS STATISTICS ====================
function setupStatusStats(sheet) {
  const tl = CONFIG.taskListSheetName;
  const statusCol = getColLetter(CONFIG.columns.status);
  
  // Header
  sheet.getRange("A6").setValue("📈 THỐNG KÊ THEO TRẠNG THÁI").setFontSize(12).setFontWeight("bold");
  sheet.getRange("A6:D6").merge().setBackground("#34a853").setFontColor("white");
  
  // Table headers
  const headers = ["Trạng thái", "Số lượng", "Phần trăm", "Biểu đồ"];
  headers.forEach((h, i) => sheet.getRange(7, i + 1).setValue(h));
  sheet.getRange("A7:D7").setFontWeight("bold").setBackground("#e6f4ea");
  
  // Status data
  const statuses = [
    ["🟢 Open", "Open"],
    ["🟡 Pending", "Pending"],
    ["🔵 In Progress", "In Progress"],
    ["🟣 Testing", "Testing"],
    ["✅ Finished", "Finished"],
    ["⬛ Closed", "Closed"]
  ];
  
  statuses.forEach((status, i) => {
    const row = 8 + i;
    sheet.getRange(row, 1).setValue(status[0]);
    sheet.getRange(row, 2).setFormula(`=COUNTIF('${tl}'!${statusCol}:${statusCol},"${status[1]}")`);
    sheet.getRange(row, 3).setFormula(`=IF($A$4>0,B${row}/$A$4,0)`).setNumberFormat("0.0%");
    sheet.getRange(row, 4).setFormula(`=REPT("█",ROUND(C${row}*20))&REPT("░",20-ROUND(C${row}*20))`).setFontSize(8);
  });
  
  // Total
  const totalRow = 8 + statuses.length;
  sheet.getRange(totalRow, 1).setValue("TỔNG").setFontWeight("bold");
  sheet.getRange(totalRow, 2).setFormula(`=SUM(B8:B${totalRow-1})`).setFontWeight("bold");
  sheet.getRange(totalRow, 3).setValue("100%").setFontWeight("bold");
  
  sheet.getRange(`A7:D${totalRow}`).setBorder(true, true, true, true, true, true);
}

// ==================== PRIORITY STATISTICS ====================
function setupPriorityStats(sheet) {
  const tl = CONFIG.taskListSheetName;
  const priorityCol = getColLetter(CONFIG.columns.priority);
  const statusCol = getColLetter(CONFIG.columns.status);
  
  // Header
  sheet.getRange("F6").setValue("🎯 THỐNG KÊ THEO ĐỘ ƯU TIÊN").setFontSize(12).setFontWeight("bold");
  sheet.getRange("F6:J6").merge().setBackground("#ea4335").setFontColor("white");
  
  // Table headers
  sheet.getRange("F7").setValue("Độ ưu tiên");
  sheet.getRange("G7").setValue("Tổng");
  sheet.getRange("H7").setValue("Chưa xong");
  sheet.getRange("I7").setValue("Phần trăm");
  sheet.getRange("J7").setValue("⚠️ Cảnh báo");
  sheet.getRange("F7:J7").setFontWeight("bold").setBackground("#fce8e6");
  
  // Priority data
  const priorities = [
    ["🔴 Urgent", "Urgent", "#ffcdd2"],
    ["🟠 High", "High", "#ffe0b2"],
    ["🟡 Normal", "Normal", "#fff9c4"],
    ["🟢 Low", "Low", "#c8e6c9"]
  ];
  
  priorities.forEach((p, i) => {
    const row = 8 + i;
    sheet.getRange(row, 6).setValue(p[0]).setBackground(p[2]);
    sheet.getRange(row, 7).setFormula(`=COUNTIF('${tl}'!${priorityCol}:${priorityCol},"${p[1]}")`);
    sheet.getRange(row, 8).setFormula(`=COUNTIFS('${tl}'!${priorityCol}:${priorityCol},"${p[1]}",'${tl}'!${statusCol}:${statusCol},"<>Finished",'${tl}'!${statusCol}:${statusCol},"<>Closed")`);
    sheet.getRange(row, 9).setFormula(`=IF($A$4>0,G${row}/$A$4,0)`).setNumberFormat("0.0%");
    sheet.getRange(row, 10).setFormula(`=IF(H${row}>0,"⚠️ Cần xử lý "&H${row}&" task","")`);
  });
  
  // Total
  sheet.getRange(12, 6).setValue("TỔNG").setFontWeight("bold");
  sheet.getRange(12, 7).setFormula("=SUM(G8:G11)").setFontWeight("bold");
  sheet.getRange(12, 8).setFormula("=SUM(H8:H11)").setFontWeight("bold");
  
  sheet.getRange("F7:J12").setBorder(true, true, true, true, true, true);
}

// ==================== ASSIGNEE OVERVIEW (với Multiple Select) ====================
function setupAssigneeOverview(sheet) {
  const tl = CONFIG.taskListSheetName;
  const assigneeCol = getColLetter(CONFIG.columns.assignee);
  
  // Header
  sheet.getRange("A16").setValue("👥 THỐNG KÊ THEO NGƯỜI THỰC HIỆN").setFontSize(12).setFontWeight("bold");
  sheet.getRange("A16:C16").merge().setBackground("#9c27b0").setFontColor("white");
  
  // Table headers
  sheet.getRange("A17").setValue("Assignee");
  sheet.getRange("B17").setValue("Số Task");
  sheet.getRange("C17").setValue("Biểu đồ");
  sheet.getRange("A17:C17").setFontWeight("bold").setBackground("#f3e5f5");
  
  // Assignee data - dùng REGEXMATCH để đếm vì là multiple select
  CONFIG.assignees.forEach((assignee, i) => {
    const row = 18 + i;
    sheet.getRange(row, 1).setValue(assignee);
    // Dùng SUMPRODUCT với REGEXMATCH để đếm task chứa tên assignee
    sheet.getRange(row, 2).setFormula(`=SUMPRODUCT(REGEXMATCH('${tl}'!${assigneeCol}:${assigneeCol},"(?i).*${assignee}.*")*1)`);
    sheet.getRange(row, 3).setFormula(`=REPT("█",B${row})&" ("&B${row}&")")`).setFontSize(9);
  });
  
  const endRow = 17 + CONFIG.assignees.length;
  sheet.getRange(`A17:C${endRow}`).setBorder(true, true, true, true, true, true);
}

// ==================== UPCOMING DEADLINES ====================
function setupUpcomingDeadlines(sheet) {
  const tl = CONFIG.taskListSheetName;
  const taskNameCol = getColLetter(CONFIG.columns.taskName);
  const assigneeCol = getColLetter(CONFIG.columns.assignee);
  const statusCol = getColLetter(CONFIG.columns.status);
  const priorityCol = getColLetter(CONFIG.columns.priority);
  const endDateCol = getColLetter(CONFIG.columns.endDate);
  const remainingCol = getColLetter(CONFIG.columns.remainingTime);
  const taskIdCol = getColLetter(CONFIG.columns.taskId);
  
  // Header
  sheet.getRange("E16").setValue("⏰ TASK SẮP/QUÁ HẾT HẠN (trong 3 ngày)").setFontSize(12).setFontWeight("bold");
  sheet.getRange("E16:K16").merge().setBackground("#f57c00").setFontColor("white");
  
  // Table headers
  sheet.getRange("E17").setValue("FNo.");
  sheet.getRange("F17").setValue("Task Name");
  sheet.getRange("G17").setValue("Assignee");
  sheet.getRange("H17").setValue("Priority");
  sheet.getRange("I17").setValue("End Date");
  sheet.getRange("J17").setValue("Còn lại");
  sheet.getRange("K17").setValue("Status");
  sheet.getRange("E17:K17").setFontWeight("bold").setBackground("#fff3e0");
  
  // Filter - sử dụng End Date để lọc task sắp hết hạn
  // Remaining Time = End Date - NOW(), hiển thị hh:mm
  // Lọc: End Date trong 3 ngày tới hoặc đã quá hạn, và chưa hoàn thành
  sheet.getRange("E18").setFormula(`=IFERROR(
    SORT(
      FILTER(
        {'${tl}'!${taskIdCol}2:${taskIdCol},'${tl}'!${taskNameCol}2:${taskNameCol},'${tl}'!${assigneeCol}2:${assigneeCol},'${tl}'!${priorityCol}2:${priorityCol},'${tl}'!${endDateCol}2:${endDateCol},'${tl}'!${remainingCol}2:${remainingCol},'${tl}'!${statusCol}2:${statusCol}},
        ('${tl}'!${endDateCol}2:${endDateCol}<>"")*
        ('${tl}'!${endDateCol}2:${endDateCol}<=TODAY()+3)*
        ('${tl}'!${statusCol}2:${statusCol}<>"Finished")*
        ('${tl}'!${statusCol}2:${statusCol}<>"Closed")*
        ('${tl}'!${taskIdCol}2:${taskIdCol}<>"")
      ),
      5, TRUE
    ),
    "✅ Không có task sắp hết hạn"
  )`);
  
  sheet.getRange("E17:K27").setBorder(true, true, true, true, true, true);
  
  // Thêm alert cho task quá hạn
  sheet.getRange("E28").setFormula(`=IF(COUNTIFS('${tl}'!${endDateCol}:${endDateCol},"<"&TODAY(),'${tl}'!${statusCol}:${statusCol},"<>Finished",'${tl}'!${statusCol}:${statusCol},"<>Closed",'${tl}'!${taskIdCol}:${taskIdCol},"F*")>0,"🚨 CÓ "&COUNTIFS('${tl}'!${endDateCol}:${endDateCol},"<"&TODAY(),'${tl}'!${statusCol}:${statusCol},"<>Finished",'${tl}'!${statusCol}:${statusCol},"<>Closed",'${tl}'!${taskIdCol}:${taskIdCol},"F*")&" TASK ĐÃ QUÁ HẠN!","")`);
  sheet.getRange("E28").setFontSize(12).setFontWeight("bold").setFontColor("#d32f2f");
}

// ==================== ASSIGNEE DETAIL TABLE ====================
function setupAssigneeDetailTable(sheet) {
  const tl = CONFIG.taskListSheetName;
  const taskNameCol = getColLetter(CONFIG.columns.taskName);
  const assigneeCol = getColLetter(CONFIG.columns.assignee);
  const statusCol = getColLetter(CONFIG.columns.status);
  const priorityCol = getColLetter(CONFIG.columns.priority);
  const taskIdCol = getColLetter(CONFIG.columns.taskId);
  
  // Header
  sheet.getRange("A30").setValue("📋 CHI TIẾT THEO TỪNG NGƯỜI - PHÂN TÍCH WORKLOAD").setFontSize(12).setFontWeight("bold");
  sheet.getRange("A30:L30").merge().setBackground("#1565c0").setFontColor("white");
  
  // Table headers
  const headers = [
    "Assignee", "Tổng Task", "✅ Done", "🔄 In Progress", "🧪 Testing", 
    "⏳ Pending", "🔴 Urgent", "🟠 High", "🟡 Normal", "🟢 Low", "📝 Task đang làm"
  ];
  headers.forEach((h, i) => sheet.getRange(31, i + 1).setValue(h));
  sheet.getRange("A31:K31").setFontWeight("bold").setBackground("#e3f2fd");
  
  // Data rows for each assignee
  CONFIG.assignees.forEach((assignee, i) => {
    const row = 32 + i;
    
    // Assignee name
    sheet.getRange(row, 1).setValue(assignee);
    
    // Tổng Task (multiple select - dùng REGEXMATCH)
    sheet.getRange(row, 2).setFormula(`=SUMPRODUCT(REGEXMATCH('${tl}'!${assigneeCol}:${assigneeCol},"(?i).*${assignee}.*")*1)`);
    
    // Done (Finished + Closed)
    sheet.getRange(row, 3).setFormula(`=SUMPRODUCT(REGEXMATCH('${tl}'!${assigneeCol}:${assigneeCol},"(?i).*${assignee}.*")*('${tl}'!${statusCol}:${statusCol}="Finished")*1)+SUMPRODUCT(REGEXMATCH('${tl}'!${assigneeCol}:${assigneeCol},"(?i).*${assignee}.*")*('${tl}'!${statusCol}:${statusCol}="Closed")*1)`);
    
    // In Progress
    sheet.getRange(row, 4).setFormula(`=SUMPRODUCT(REGEXMATCH('${tl}'!${assigneeCol}:${assigneeCol},"(?i).*${assignee}.*")*('${tl}'!${statusCol}:${statusCol}="In Progress")*1)`);
    
    // Testing
    sheet.getRange(row, 5).setFormula(`=SUMPRODUCT(REGEXMATCH('${tl}'!${assigneeCol}:${assigneeCol},"(?i).*${assignee}.*")*('${tl}'!${statusCol}:${statusCol}="Testing")*1)`);
    
    // Pending (Open + Pending)
    sheet.getRange(row, 6).setFormula(`=SUMPRODUCT(REGEXMATCH('${tl}'!${assigneeCol}:${assigneeCol},"(?i).*${assignee}.*")*(('${tl}'!${statusCol}:${statusCol}="Open")+('${tl}'!${statusCol}:${statusCol}="Pending"))*1)`);
    
    // Urgent (chưa xong)
    sheet.getRange(row, 7).setFormula(`=SUMPRODUCT(REGEXMATCH('${tl}'!${assigneeCol}:${assigneeCol},"(?i).*${assignee}.*")*('${tl}'!${priorityCol}:${priorityCol}="Urgent")*('${tl}'!${statusCol}:${statusCol}<>"Finished")*('${tl}'!${statusCol}:${statusCol}<>"Closed")*1)`);
    
    // High (chưa xong)
    sheet.getRange(row, 8).setFormula(`=SUMPRODUCT(REGEXMATCH('${tl}'!${assigneeCol}:${assigneeCol},"(?i).*${assignee}.*")*('${tl}'!${priorityCol}:${priorityCol}="High")*('${tl}'!${statusCol}:${statusCol}<>"Finished")*('${tl}'!${statusCol}:${statusCol}<>"Closed")*1)`);
    
    // Normal (chưa xong)
    sheet.getRange(row, 9).setFormula(`=SUMPRODUCT(REGEXMATCH('${tl}'!${assigneeCol}:${assigneeCol},"(?i).*${assignee}.*")*('${tl}'!${priorityCol}:${priorityCol}="Normal")*('${tl}'!${statusCol}:${statusCol}<>"Finished")*('${tl}'!${statusCol}:${statusCol}<>"Closed")*1)`);
    
    // Low (chưa xong)
    sheet.getRange(row, 10).setFormula(`=SUMPRODUCT(REGEXMATCH('${tl}'!${assigneeCol}:${assigneeCol},"(?i).*${assignee}.*")*('${tl}'!${priorityCol}:${priorityCol}="Low")*('${tl}'!${statusCol}:${statusCol}<>"Finished")*('${tl}'!${statusCol}:${statusCol}<>"Closed")*1)`);
    
    // Task đang làm (In Progress)
    sheet.getRange(row, 11).setFormula(`=IFERROR(TEXTJOIN(", ",TRUE,FILTER('${tl}'!${taskIdCol}:${taskIdCol}&": "&'${tl}'!${taskNameCol}:${taskNameCol},REGEXMATCH('${tl}'!${assigneeCol}:${assigneeCol},"(?i).*${assignee}.*")*('${tl}'!${statusCol}:${statusCol}="In Progress"))),"Không có")`);
  });
  
  const endRow = 31 + CONFIG.assignees.length;
  
  // Conditional formatting cho Urgent
  CONFIG.assignees.forEach((_, i) => {
    const row = 32 + i;
    sheet.getRange(row, 7).setFormula(sheet.getRange(row, 7).getFormula()); // Keep formula
  });
  
  // Total row
  sheet.getRange(endRow + 1, 1).setValue("TỔNG").setFontWeight("bold");
  for (let col = 2; col <= 10; col++) {
    sheet.getRange(endRow + 1, col).setFormula(`=SUM(${getColLetter(col)}32:${getColLetter(col)}${endRow})`).setFontWeight("bold");
  }
  
  sheet.getRange(`A31:K${endRow + 1}`).setBorder(true, true, true, true, true, true);
  
  // Thêm conditional formatting cho cột Urgent
  const urgentRange = sheet.getRange(`G32:G${endRow}`);
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground("#ffcdd2")
    .setFontColor("#c62828")
    .setRanges([urgentRange])
    .build();
  
  const highRange = sheet.getRange(`H32:H${endRow}`);
  const rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground("#ffe0b2")
    .setFontColor("#e65100")
    .setRanges([highRange])
    .build();
  
  sheet.setConditionalFormatRules([rule, rule2]);
}

// ==================== FORMATTING ====================
function formatOverviewSheet(sheet) {
  // Set column widths
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 90);
  sheet.setColumnWidth(6, 150);
  sheet.setColumnWidth(7, 90);
  sheet.setColumnWidth(8, 90);
  sheet.setColumnWidth(9, 90);
  sheet.setColumnWidth(10, 90);
  sheet.setColumnWidth(11, 350);
  
  // Freeze rows
  sheet.setFrozenRows(2);
  
  // Set default font
  sheet.getRange("A1:K100").setFontFamily("Arial");
}

// ==================== TẠO BIỂU ĐỒ ====================
function createCharts(sheet) {
  // Biểu đồ tròn Status
  const statusChart = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange("A8:B13"))
    .setPosition(6, 12, 0, 0)
    .setOption('title', '📈 Phân bổ theo Trạng thái')
    .setOption('pieHole', 0.4)
    .setOption('width', 380)
    .setOption('height', 280)
    .setOption('legend', {position: 'right'})
    .setOption('colors', ['#4caf50', '#ffeb3b', '#2196f3', '#9c27b0', '#8bc34a', '#607d8b'])
    .build();
  sheet.insertChart(statusChart);
  
  // Biểu đồ cột Priority
  const priorityChart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange("F8:H11"))
    .setPosition(16, 12, 0, 0)
    .setOption('title', '🎯 Task theo Độ ưu tiên')
    .setOption('width', 380)
    .setOption('height', 280)
    .setOption('legend', {position: 'top'})
    .setOption('colors', ['#9e9e9e', '#f44336'])
    .setOption('hAxis', {title: 'Priority'})
    .setOption('vAxis', {title: 'Số lượng'})
    .build();
  sheet.insertChart(priorityChart);
}

// ==================== MENU ====================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('📊 Task Overview')
    .addItem('🔄 Tạo/Cập nhật Overview', 'createOverviewSheet')
    .addItem('📈 Chỉ cập nhật biểu đồ', 'updateChartsOnly')
    .addSeparator()
    .addItem('ℹ️ Hướng dẫn', 'showHelp')
    .addToUi();
}

function updateChartsOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.overviewSheetName);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Vui lòng tạo Overview Sheet trước!');
    return;
  }
  
  // Xóa charts cũ
  sheet.getCharts().forEach(c => sheet.removeChart(c));
  
  // Tạo lại
  createCharts(sheet);
  SpreadsheetApp.getUi().alert('Đã cập nhật biểu đồ!');
}

function showHelp() {
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial; padding: 15px;">
      <h2>📊 Task Overview - Hướng dẫn</h2>
      
      <h3>🔹 Tính năng</h3>
      <ul>
        <li><b>KPI Dashboard:</b> Tổng quan số task, % hoàn thành</li>
        <li><b>Thống kê Status:</b> Biểu đồ tròn theo trạng thái</li>
        <li><b>Thống kê Priority:</b> Biểu đồ cột theo độ ưu tiên</li>
        <li><b>Workload Assignee:</b> Phân tích task từng người</li>
        <li><b>Task đang làm:</b> Hiển thị task In Progress của mỗi người</li>
        <li><b>Deadline Alert:</b> Danh sách task sắp hết hạn</li>
      </ul>
      
      <h3>🔹 Lưu ý Multiple Select</h3>
      <p>Script đã được tối ưu để đếm chính xác khi 1 task có nhiều Assignee.</p>
      
      <h3>🔹 Cập nhật dữ liệu</h3>
      <p>Dữ liệu tự động cập nhật realtime khi thay đổi Task List.</p>
      
      <h3>🔹 Thêm Assignee mới</h3>
      <p>Vào Apps Script, thêm tên vào mảng <code>assignees</code> trong CONFIG.</p>
    </div>
  `).setWidth(450).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Hướng dẫn sử dụng');
}
