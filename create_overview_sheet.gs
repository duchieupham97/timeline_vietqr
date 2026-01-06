/**
 * Google Apps Script để tạo Sheet "Overview" 
 * với thống kê realtime từ Sheet "Task List"
 * 
 * HƯỚNG DẪN SỬ DỤNG:
 * 1. Mở Google Sheet của bạn
 * 2. Vào Extensions → Apps Script
 * 3. Xóa code mặc định và paste toàn bộ code này vào
 * 4. Điều chỉnh CONFIG bên dưới theo cấu trúc sheet của bạn
 * 5. Nhấn nút Run (▶️) và chọn function "createOverviewSheet"
 * 6. Cấp quyền khi được yêu cầu
 * 7. Sheet "Overview" sẽ được tạo tự động!
 */

// ==================== CẤU HÌNH - ĐIỀU CHỈNH THEO SHEET CỦA BẠN ====================
const CONFIG = {
  // Tên sheet chứa danh sách task
  taskListSheetName: "Task List",
  
  // Tên sheet overview sẽ được tạo
  overviewSheetName: "Overview",
  
  // Vị trí các cột trong Task List (điều chỉnh theo thứ tự cột của bạn)
  // Số thứ tự bắt đầu từ 1 (A=1, B=2, C=3, ...)
  columns: {
    taskId: 1,        // Cột A - Task ID
    taskName: 2,      // Cột B - Tên task
    description: 3,   // Cột C - Mô tả
    assignee: 4,      // Cột D - Người được giao
    status: 5,        // Cột E - Trạng thái
    priority: 6,      // Cột F - Độ ưu tiên
    dueDate: 7,       // Cột G - Ngày hết hạn
    remainingTime: 8, // Cột H - Thời gian còn lại (số ngày hoặc text)
    startDate: 9      // Cột I - Ngày bắt đầu
  },
  
  // Các giá trị Status
  status: {
    done: ["Finished", "Closed"],           // Các status được coi là "Done"
    inProgress: ["In Progress"],            // Status "Đang thực hiện"
    pending: ["To Do", "Open", "Pending"]   // Status "Chờ xử lý"
  },
  
  // Các giá trị Priority (theo thứ tự từ cao đến thấp)
  priority: {
    urgent: "Urgent",
    high: "High",
    medium: "Medium",
    low: "Low"
  },
  
  // Số ngày để cảnh báo task sắp hết hạn
  deadlineWarningDays: 3
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
  
  // Di chuyển sheet Overview lên đầu
  ss.setActiveSheet(overviewSheet);
  ss.moveActiveSheet(1);
  
  // Thiết lập các phần thống kê
  setupDashboardKPIs(overviewSheet);
  setupStatusStats(overviewSheet);
  setupPriorityStats(overviewSheet);
  setupAssigneeStats(overviewSheet);
  setupUpcomingDeadlines(overviewSheet);
  setupAssigneeDetailTable(overviewSheet);
  
  // Format sheet
  formatOverviewSheet(overviewSheet);
  
  SpreadsheetApp.getUi().alert('✅ Sheet "Overview" đã được tạo thành công!\n\nTất cả dữ liệu sẽ tự động cập nhật khi bạn thay đổi Task List.');
}

// ==================== DASHBOARD KPIs ====================
function setupDashboardKPIs(sheet) {
  const taskListName = CONFIG.taskListSheetName;
  const statusCol = getColLetter(CONFIG.columns.status);
  const priorityCol = getColLetter(CONFIG.columns.priority);
  const remainingCol = getColLetter(CONFIG.columns.remainingTime);
  
  const doneStatuses = CONFIG.status.done.map(s => `"${s}"`).join(",");
  const inProgressStatuses = CONFIG.status.inProgress.map(s => `"${s}"`).join(",");
  
  // Header
  sheet.getRange("A1").setValue("📊 TỔNG QUAN TASK").setFontSize(16).setFontWeight("bold");
  sheet.getRange("A1:E1").merge().setBackground("#4285f4").setFontColor("white");
  
  // KPI Cards
  const kpis = [
    ["📋 Tổng Task", `=COUNTA('${taskListName}'!A2:A)`],
    ["✅ Đã hoàn thành", `=SUMPRODUCT((ISNUMBER(MATCH('${taskListName}'!${statusCol}2:${statusCol},{${doneStatuses}},0)))*1)`],
    ["🔄 Đang thực hiện", `=COUNTIF('${taskListName}'!${statusCol}:${statusCol},"${CONFIG.status.inProgress[0]}")`],
    ["⏳ Chờ xử lý", `=A3-B3-C3`],
    ["🚨 Task Urgent chưa xong", `=COUNTIFS('${taskListName}'!${priorityCol}:${priorityCol},"${CONFIG.priority.urgent}",'${taskListName}'!${statusCol}:${statusCol},"<>${CONFIG.status.done[0]}",'${taskListName}'!${statusCol}:${statusCol},"<>${CONFIG.status.done[1]}")`],
    ["📈 % Hoàn thành", `=IF(A3>0,B3/A3,0)`]
  ];
  
  sheet.getRange("A2").setValue(kpis[0][0]);
  sheet.getRange("B2").setValue(kpis[1][0]);
  sheet.getRange("C2").setValue(kpis[2][0]);
  sheet.getRange("D2").setValue(kpis[3][0]);
  sheet.getRange("E2").setValue(kpis[4][0]);
  sheet.getRange("F2").setValue(kpis[5][0]);
  
  sheet.getRange("A3").setFormula(kpis[0][1]);
  sheet.getRange("B3").setFormula(kpis[1][1]);
  sheet.getRange("C3").setFormula(kpis[2][1]);
  sheet.getRange("D3").setFormula(kpis[3][1]);
  sheet.getRange("E3").setFormula(kpis[4][1]);
  sheet.getRange("F3").setFormula(kpis[5][1]).setNumberFormat("0.0%");
  
  // Style KPI cells
  sheet.getRange("A2:F2").setBackground("#e8f0fe").setFontWeight("bold");
  sheet.getRange("A3:F3").setFontSize(18).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("E3").setFontColor("#ea4335"); // Red for urgent
}

// ==================== STATUS STATISTICS ====================
function setupStatusStats(sheet) {
  const taskListName = CONFIG.taskListSheetName;
  const statusCol = getColLetter(CONFIG.columns.status);
  
  // Header
  sheet.getRange("A5").setValue("📈 THỐNG KÊ THEO TRẠNG THÁI").setFontSize(14).setFontWeight("bold");
  sheet.getRange("A5:D5").merge().setBackground("#34a853").setFontColor("white");
  
  // Table headers
  sheet.getRange("A6").setValue("Trạng thái");
  sheet.getRange("B6").setValue("Số lượng");
  sheet.getRange("C6").setValue("Phần trăm");
  sheet.getRange("D6").setValue("Thanh tiến độ");
  sheet.getRange("A6:D6").setFontWeight("bold").setBackground("#e6f4ea");
  
  // Data rows
  const allStatuses = [...CONFIG.status.done, ...CONFIG.status.inProgress, ...CONFIG.status.pending];
  let row = 7;
  
  allStatuses.forEach(status => {
    sheet.getRange(row, 1).setValue(status);
    sheet.getRange(row, 2).setFormula(`=COUNTIF('${taskListName}'!${statusCol}:${statusCol},"${status}")`);
    sheet.getRange(row, 3).setFormula(`=IF($A$3>0,B${row}/$A$3,0)`).setNumberFormat("0.0%");
    sheet.getRange(row, 4).setFormula(`=REPT("█",ROUND(C${row}*20))&REPT("░",20-ROUND(C${row}*20))`);
    row++;
  });
  
  // Total row
  sheet.getRange(row, 1).setValue("TỔNG").setFontWeight("bold");
  sheet.getRange(row, 2).setFormula(`=SUM(B7:B${row-1})`).setFontWeight("bold");
  sheet.getRange(row, 3).setFormula(`=SUM(C7:C${row-1})`).setNumberFormat("0.0%").setFontWeight("bold");
}

// ==================== PRIORITY STATISTICS ====================
function setupPriorityStats(sheet) {
  const taskListName = CONFIG.taskListSheetName;
  const priorityCol = getColLetter(CONFIG.columns.priority);
  const statusCol = getColLetter(CONFIG.columns.status);
  
  // Header
  sheet.getRange("F5").setValue("🎯 THỐNG KÊ THEO ĐỘ ƯU TIÊN").setFontSize(14).setFontWeight("bold");
  sheet.getRange("F5:J5").merge().setBackground("#ea4335").setFontColor("white");
  
  // Table headers
  sheet.getRange("F6").setValue("Độ ưu tiên");
  sheet.getRange("G6").setValue("Tổng");
  sheet.getRange("H6").setValue("Chưa xong");
  sheet.getRange("I6").setValue("Phần trăm");
  sheet.getRange("J6").setValue("Cảnh báo");
  sheet.getRange("F6:J6").setFontWeight("bold").setBackground("#fce8e6");
  
  const priorities = [
    [CONFIG.priority.urgent, "🔴"],
    [CONFIG.priority.high, "🟠"],
    [CONFIG.priority.medium, "🟡"],
    [CONFIG.priority.low, "🟢"]
  ];
  
  let row = 7;
  priorities.forEach(([priority, icon]) => {
    const doneConditions = CONFIG.status.done.map(s => `'${taskListName}'!${statusCol}:${statusCol},"<>${s}"`).join(",");
    
    sheet.getRange(row, 6).setValue(`${icon} ${priority}`);
    sheet.getRange(row, 7).setFormula(`=COUNTIF('${taskListName}'!${priorityCol}:${priorityCol},"${priority}")`);
    sheet.getRange(row, 8).setFormula(`=COUNTIFS('${taskListName}'!${priorityCol}:${priorityCol},"${priority}",'${taskListName}'!${statusCol}:${statusCol},"<>${CONFIG.status.done[0]}",'${taskListName}'!${statusCol}:${statusCol},"<>${CONFIG.status.done[1]}")`);
    sheet.getRange(row, 9).setFormula(`=IF($A$3>0,G${row}/$A$3,0)`).setNumberFormat("0.0%");
    sheet.getRange(row, 10).setFormula(`=IF(H${row}>0,"⚠️ "&H${row}&" task cần xử lý","")`);
    row++;
  });
  
  // Total row
  sheet.getRange(row, 6).setValue("TỔNG").setFontWeight("bold");
  sheet.getRange(row, 7).setFormula(`=SUM(G7:G${row-1})`).setFontWeight("bold");
  sheet.getRange(row, 8).setFormula(`=SUM(H7:H${row-1})`).setFontWeight("bold");
  sheet.getRange(row, 9).setFormula(`=SUM(I7:I${row-1})`).setNumberFormat("0.0%").setFontWeight("bold");
}

// ==================== ASSIGNEE STATISTICS ====================
function setupAssigneeStats(sheet) {
  const taskListName = CONFIG.taskListSheetName;
  const assigneeCol = getColLetter(CONFIG.columns.assignee);
  
  // Header
  sheet.getRange("A14").setValue("👥 THỐNG KÊ THEO NGƯỜI THỰC HIỆN").setFontSize(14).setFontWeight("bold");
  sheet.getRange("A14:B14").merge().setBackground("#fbbc04").setFontColor("white");
  
  // Dùng QUERY để lấy danh sách unique assignees và đếm
  sheet.getRange("A15").setValue("Assignee");
  sheet.getRange("B15").setValue("Số Task");
  sheet.getRange("A15:B15").setFontWeight("bold").setBackground("#fef7e0");
  
  // Query để lấy thống kê
  sheet.getRange("A16").setFormula(`=IFERROR(QUERY('${taskListName}'!${assigneeCol}2:${assigneeCol},"SELECT ${assigneeCol}, COUNT(${assigneeCol}) WHERE ${assigneeCol} IS NOT NULL GROUP BY ${assigneeCol} ORDER BY COUNT(${assigneeCol}) DESC LABEL COUNT(${assigneeCol}) ''"),"")`);
}

// ==================== UPCOMING DEADLINES ====================
function setupUpcomingDeadlines(sheet) {
  const taskListName = CONFIG.taskListSheetName;
  const taskNameCol = getColLetter(CONFIG.columns.taskName);
  const assigneeCol = getColLetter(CONFIG.columns.assignee);
  const statusCol = getColLetter(CONFIG.columns.status);
  const priorityCol = getColLetter(CONFIG.columns.priority);
  const dueDateCol = getColLetter(CONFIG.columns.dueDate);
  const remainingCol = getColLetter(CONFIG.columns.remainingTime);
  
  // Header
  sheet.getRange("D14").setValue(`⏰ TASK SẮP HẾT HẠN (trong ${CONFIG.deadlineWarningDays} ngày)`).setFontSize(14).setFontWeight("bold");
  sheet.getRange("D14:I14").merge().setBackground("#ea4335").setFontColor("white");
  
  // Table headers
  sheet.getRange("D15").setValue("Task Name");
  sheet.getRange("E15").setValue("Assignee");
  sheet.getRange("F15").setValue("Priority");
  sheet.getRange("G15").setValue("Due Date");
  sheet.getRange("H15").setValue("Còn lại");
  sheet.getRange("I15").setValue("Status");
  sheet.getRange("D15:I15").setFontWeight("bold").setBackground("#fce8e6");
  
  // Filter formula - lọc task sắp hết hạn
  // Giả sử Remaining Time là số ngày
  sheet.getRange("D16").setFormula(`=IFERROR(FILTER({'${taskListName}'!${taskNameCol}2:${taskNameCol},'${taskListName}'!${assigneeCol}2:${assigneeCol},'${taskListName}'!${priorityCol}2:${priorityCol},'${taskListName}'!${dueDateCol}2:${dueDateCol},'${taskListName}'!${remainingCol}2:${remainingCol},'${taskListName}'!${statusCol}2:${statusCol}},('${taskListName}'!${remainingCol}2:${remainingCol}<=${CONFIG.deadlineWarningDays})*('${taskListName}'!${remainingCol}2:${remainingCol}>=-1)*('${taskListName}'!${statusCol}2:${statusCol}<>"${CONFIG.status.done[0]}")*('${taskListName}'!${statusCol}2:${statusCol}<>"${CONFIG.status.done[1]}")),"✅ Không có task sắp hết hạn")`);
}

// ==================== ASSIGNEE DETAIL TABLE ====================
function setupAssigneeDetailTable(sheet) {
  const taskListName = CONFIG.taskListSheetName;
  const taskNameCol = getColLetter(CONFIG.columns.taskName);
  const assigneeCol = getColLetter(CONFIG.columns.assignee);
  const statusCol = getColLetter(CONFIG.columns.status);
  const priorityCol = getColLetter(CONFIG.columns.priority);
  
  // Header
  sheet.getRange("A28").setValue("📋 BẢNG CHI TIẾT THEO NGƯỜI THỰC HIỆN").setFontSize(14).setFontWeight("bold");
  sheet.getRange("A28:K28").merge().setBackground("#9c27b0").setFontColor("white");
  
  // Table headers
  const headers = [
    "Assignee", "Tổng Task", "✅ Done", "🔄 In Progress", "⏳ Pending",
    "🔴 Urgent", "🟠 High", "🟡 Medium", "🟢 Low", "📝 Task đang làm"
  ];
  
  headers.forEach((header, i) => {
    sheet.getRange(29, i + 1).setValue(header);
  });
  sheet.getRange("A29:J29").setFontWeight("bold").setBackground("#f3e5f5");
  
  // Get unique assignees formula
  sheet.getRange("A30").setFormula(`=IFERROR(UNIQUE(FILTER('${taskListName}'!${assigneeCol}2:${assigneeCol},'${taskListName}'!${assigneeCol}2:${assigneeCol}<>"")),"Không có dữ liệu")`);
  
  // Công thức cho các cột khác (sẽ được áp dụng cho từng dòng)
  // Giả sử có tối đa 20 assignees
  for (let row = 30; row <= 49; row++) {
    // Tổng Task
    sheet.getRange(row, 2).setFormula(`=IF(A${row}<>"",COUNTIF('${taskListName}'!${assigneeCol}:${assigneeCol},A${row}),"")`);
    
    // Done (Finished + Closed)
    sheet.getRange(row, 3).setFormula(`=IF(A${row}<>"",COUNTIFS('${taskListName}'!${assigneeCol}:${assigneeCol},A${row},'${taskListName}'!${statusCol}:${statusCol},"${CONFIG.status.done[0]}")+COUNTIFS('${taskListName}'!${assigneeCol}:${assigneeCol},A${row},'${taskListName}'!${statusCol}:${statusCol},"${CONFIG.status.done[1]}"),"")`);
    
    // In Progress
    sheet.getRange(row, 4).setFormula(`=IF(A${row}<>"",COUNTIFS('${taskListName}'!${assigneeCol}:${assigneeCol},A${row},'${taskListName}'!${statusCol}:${statusCol},"${CONFIG.status.inProgress[0]}"),"")`);
    
    // Pending
    sheet.getRange(row, 5).setFormula(`=IF(A${row}<>"",B${row}-C${row}-D${row},"")`);
    
    // Urgent
    sheet.getRange(row, 6).setFormula(`=IF(A${row}<>"",COUNTIFS('${taskListName}'!${assigneeCol}:${assigneeCol},A${row},'${taskListName}'!${priorityCol}:${priorityCol},"${CONFIG.priority.urgent}",'${taskListName}'!${statusCol}:${statusCol},"<>${CONFIG.status.done[0]}",'${taskListName}'!${statusCol}:${statusCol},"<>${CONFIG.status.done[1]}"),"")`);
    
    // High
    sheet.getRange(row, 7).setFormula(`=IF(A${row}<>"",COUNTIFS('${taskListName}'!${assigneeCol}:${assigneeCol},A${row},'${taskListName}'!${priorityCol}:${priorityCol},"${CONFIG.priority.high}",'${taskListName}'!${statusCol}:${statusCol},"<>${CONFIG.status.done[0]}",'${taskListName}'!${statusCol}:${statusCol},"<>${CONFIG.status.done[1]}"),"")`);
    
    // Medium
    sheet.getRange(row, 8).setFormula(`=IF(A${row}<>"",COUNTIFS('${taskListName}'!${assigneeCol}:${assigneeCol},A${row},'${taskListName}'!${priorityCol}:${priorityCol},"${CONFIG.priority.medium}",'${taskListName}'!${statusCol}:${statusCol},"<>${CONFIG.status.done[0]}",'${taskListName}'!${statusCol}:${statusCol},"<>${CONFIG.status.done[1]}"),"")`);
    
    // Low
    sheet.getRange(row, 9).setFormula(`=IF(A${row}<>"",COUNTIFS('${taskListName}'!${assigneeCol}:${assigneeCol},A${row},'${taskListName}'!${priorityCol}:${priorityCol},"${CONFIG.priority.low}",'${taskListName}'!${statusCol}:${statusCol},"<>${CONFIG.status.done[0]}",'${taskListName}'!${statusCol}:${statusCol},"<>${CONFIG.status.done[1]}"),"")`);
    
    // Task đang làm
    sheet.getRange(row, 10).setFormula(`=IF(A${row}<>"",IFERROR(TEXTJOIN(", ",TRUE,FILTER('${taskListName}'!${taskNameCol}:${taskNameCol},('${taskListName}'!${assigneeCol}:${assigneeCol}=A${row})*('${taskListName}'!${statusCol}:${statusCol}="${CONFIG.status.inProgress[0]}"))),"Không có"),"")`);
  }
}

// ==================== FORMATTING ====================
function formatOverviewSheet(sheet) {
  // Set column widths
  sheet.setColumnWidth(1, 150);  // A
  sheet.setColumnWidth(2, 100);  // B
  sheet.setColumnWidth(3, 100);  // C
  sheet.setColumnWidth(4, 150);  // D
  sheet.setColumnWidth(5, 120);  // E
  sheet.setColumnWidth(6, 120);  // F
  sheet.setColumnWidth(7, 80);   // G
  sheet.setColumnWidth(8, 100);  // H
  sheet.setColumnWidth(9, 80);   // I
  sheet.setColumnWidth(10, 300); // J - Task đang làm
  
  // Freeze first row
  sheet.setFrozenRows(1);
  
  // Add conditional formatting for urgent tasks
  const urgentRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("Urgent")
    .setBackground("#ffcdd2")
    .setRanges([sheet.getRange("F7:F10"), sheet.getRange("F30:F49")])
    .build();
  
  // Add conditional formatting for high priority
  const highRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("High")
    .setBackground("#ffe0b2")
    .setRanges([sheet.getRange("F7:F10"), sheet.getRange("F30:F49")])
    .build();
  
  // Apply rules
  const rules = sheet.getConditionalFormatRules();
  rules.push(urgentRule);
  rules.push(highRule);
  sheet.setConditionalFormatRules(rules);
  
  // Add borders
  sheet.getRange("A6:D12").setBorder(true, true, true, true, true, true);
  sheet.getRange("F6:J11").setBorder(true, true, true, true, true, true);
  sheet.getRange("A15:B25").setBorder(true, true, true, true, true, true);
  sheet.getRange("D15:I25").setBorder(true, true, true, true, true, true);
  sheet.getRange("A29:J49").setBorder(true, true, true, true, true, true);
}

// ==================== HELPER FUNCTIONS ====================
function getColLetter(colNum) {
  let letter = '';
  while (colNum > 0) {
    let mod = (colNum - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    colNum = Math.floor((colNum - mod) / 26);
  }
  return letter;
}

// ==================== MENU ====================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Task Overview')
    .addItem('🔄 Tạo/Cập nhật Overview Sheet', 'createOverviewSheet')
    .addItem('ℹ️ Hướng dẫn', 'showHelp')
    .addToUi();
}

function showHelp() {
  const htmlOutput = HtmlService.createHtmlOutput(`
    <h2>📊 Hướng dẫn sử dụng Task Overview</h2>
    <h3>Bước 1: Cấu hình</h3>
    <p>Mở Apps Script (Extensions → Apps Script) và điều chỉnh phần CONFIG theo cấu trúc sheet của bạn:</p>
    <ul>
      <li><b>taskListSheetName:</b> Tên sheet chứa danh sách task</li>
      <li><b>columns:</b> Vị trí các cột (A=1, B=2, ...)</li>
      <li><b>status:</b> Các giá trị trạng thái</li>
      <li><b>priority:</b> Các giá trị độ ưu tiên</li>
    </ul>
    <h3>Bước 2: Chạy Script</h3>
    <p>Click menu "📊 Task Overview" → "🔄 Tạo/Cập nhật Overview Sheet"</p>
    <h3>Lưu ý</h3>
    <p>- Dữ liệu sẽ tự động cập nhật realtime<br>
    - Chạy lại script nếu muốn reset layout</p>
  `)
    .setWidth(500)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Hướng dẫn');
}

// ==================== TẠO BIỂU ĐỒ ====================
function createCharts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.overviewSheetName);
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Vui lòng chạy "Tạo Overview Sheet" trước!');
    return;
  }
  
  // Xóa charts cũ
  const charts = sheet.getCharts();
  charts.forEach(chart => sheet.removeChart(chart));
  
  // Biểu đồ tròn cho Status (A6:C11)
  const statusChart = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange("A7:B11"))
    .setPosition(5, 12, 0, 0)
    .setOption('title', 'Phân bổ theo Trạng thái')
    .setOption('pieHole', 0.4)
    .setOption('width', 400)
    .setOption('height', 300)
    .build();
  sheet.insertChart(statusChart);
  
  // Biểu đồ cột cho Priority
  const priorityChart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange("F7:H10"))
    .setPosition(14, 12, 0, 0)
    .setOption('title', 'Task theo Độ ưu tiên')
    .setOption('width', 400)
    .setOption('height', 300)
    .setOption('colors', ['#ea4335', '#fbbc04'])
    .build();
  sheet.insertChart(priorityChart);
}
