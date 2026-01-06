/**
 * Google Apps Script để tạo Sheet "Overview" 
 * cho Task List của Team Timeline 2601
 * VERSION 2 - ĐÃ SỬA LỖI
 */

// ==================== CẤU HÌNH ====================
const CONFIG = {
  taskListSheetName: "Task List",
  overviewSheetName: "Overview",
  
  // Vị trí cột trong Task List
  columns: {
    taskId: 1,         // A - FNo.
    taskName: 2,       // B - Functional
    startDate: 3,      // C - Start Date
    endDate: 4,        // D - End Date
    remainingTime: 5,  // E - Remaining Time
    priority: 6,       // F - Priority
    status: 7,         // G - Status
    assignee: 8,       // H - Assignee (MULTIPLE SELECT)
    tester: 9,         // I - Tester
    progress: 10,      // J - Progress (%)
    note: 11           // K - Reference/Note
  },
  
  // Danh sách Assignee
  assignees: ["Duy Anh", "Trường", "Đức", "Triều", "Nghĩa", "Hiếu Phạm", "Quyết", "Hiếu Hà", "Tôn"],
  
  // Dòng cuối cùng có data (điều chỉnh nếu cần)
  lastDataRow: 200
};

// ==================== MAIN ====================
function createOverviewSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Xóa sheet cũ nếu có
  let sheet = ss.getSheetByName(CONFIG.overviewSheetName);
  if (sheet) ss.deleteSheet(sheet);
  
  // Tạo mới
  sheet = ss.insertSheet(CONFIG.overviewSheetName);
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(1);
  
  const tl = CONFIG.taskListSheetName;
  const lastRow = CONFIG.lastDataRow;
  
  // ===== TITLE =====
  sheet.getRange("A1").setValue("📊 TASK OVERVIEW - TIMELINE 2601")
    .setFontSize(18).setFontWeight("bold");
  sheet.getRange("A1:K1").merge().setBackground("#1a73e8").setFontColor("white").setHorizontalAlignment("center");
  
  // ===== KPI ROW =====
  sheet.getRange("A3:G3").setValues([["📋 Tổng Task", "✅ Hoàn thành", "🔄 Đang làm", "🧪 Testing", "⏳ Chờ xử lý", "🚨 Urgent", "📈 % Hoàn thành"]]);
  sheet.getRange("A3:G3").setFontWeight("bold").setBackground("#e8f0fe").setHorizontalAlignment("center");
  
  // KPI formulas - dùng range cụ thể
  sheet.getRange("A4").setFormula(`=COUNTIF('${tl}'!A3:A${lastRow},"F*")`);
  sheet.getRange("B4").setFormula(`=COUNTIF('${tl}'!G3:G${lastRow},"Finished")+COUNTIF('${tl}'!G3:G${lastRow},"Closed")`);
  sheet.getRange("C4").setFormula(`=COUNTIF('${tl}'!G3:G${lastRow},"In Progress")`);
  sheet.getRange("D4").setFormula(`=COUNTIF('${tl}'!G3:G${lastRow},"Testing")`);
  sheet.getRange("E4").setFormula(`=COUNTIF('${tl}'!G3:G${lastRow},"Open")+COUNTIF('${tl}'!G3:G${lastRow},"Pending")`);
  sheet.getRange("F4").setFormula(`=COUNTIFS('${tl}'!F3:F${lastRow},"Urgent",'${tl}'!G3:G${lastRow},"<>Finished",'${tl}'!G3:G${lastRow},"<>Closed")`);
  sheet.getRange("G4").setFormula(`=IFERROR(B4/A4,0)`);
  
  sheet.getRange("A4:F4").setFontSize(24).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("G4").setFontSize(24).setFontWeight("bold").setHorizontalAlignment("center").setNumberFormat("0.0%");
  sheet.getRange("F4").setFontColor("#d93025");
  sheet.getRange("B4").setFontColor("#1e8e3e");
  sheet.getRange("A3:G4").setBorder(true, true, true, true, true, true);
  
  // ===== STATUS STATS =====
  sheet.getRange("A6").setValue("📈 THỐNG KÊ THEO TRẠNG THÁI").setFontSize(12).setFontWeight("bold");
  sheet.getRange("A6:D6").merge().setBackground("#34a853").setFontColor("white");
  
  sheet.getRange("A7:D7").setValues([["Trạng thái", "Số lượng", "Phần trăm", "Biểu đồ"]]);
  sheet.getRange("A7:D7").setFontWeight("bold").setBackground("#e6f4ea");
  
  const statuses = ["Open", "Pending", "In Progress", "Testing", "Finished", "Closed"];
  const statusIcons = ["🟢", "🟡", "🔵", "🟣", "✅", "⬛"];
  
  statuses.forEach((s, i) => {
    const row = 8 + i;
    sheet.getRange(row, 1).setValue(`${statusIcons[i]} ${s}`);
    sheet.getRange(row, 2).setFormula(`=COUNTIF('${tl}'!G3:G${lastRow},"${s}")`);
    sheet.getRange(row, 3).setFormula(`=IFERROR(B${row}/$A$4,0)`).setNumberFormat("0.0%");
    sheet.getRange(row, 4).setFormula(`=REPT("█",ROUND(C${row}*20))&REPT("░",20-ROUND(C${row}*20))`).setFontSize(8);
  });
  
  sheet.getRange(14, 1).setValue("TỔNG").setFontWeight("bold");
  sheet.getRange(14, 2).setFormula("=SUM(B8:B13)").setFontWeight("bold");
  sheet.getRange(14, 3).setValue("100%").setFontWeight("bold");
  sheet.getRange("A7:D14").setBorder(true, true, true, true, true, true);
  
  // ===== PRIORITY STATS =====
  sheet.getRange("F6").setValue("🎯 THỐNG KÊ THEO ĐỘ ƯU TIÊN").setFontSize(12).setFontWeight("bold");
  sheet.getRange("F6:J6").merge().setBackground("#ea4335").setFontColor("white");
  
  sheet.getRange("F7:J7").setValues([["Độ ưu tiên", "Tổng", "Chưa xong", "Phần trăm", "⚠️ Cảnh báo"]]);
  sheet.getRange("F7:J7").setFontWeight("bold").setBackground("#fce8e6");
  
  const priorities = [
    ["🔴 Urgent", "Urgent", "#ffcdd2"],
    ["🟠 High", "High", "#ffe0b2"],
    ["🟡 Normal", "Normal", "#fff9c4"],
    ["🟢 Low", "Low", "#c8e6c9"]
  ];
  
  priorities.forEach((p, i) => {
    const row = 8 + i;
    sheet.getRange(row, 6).setValue(p[0]).setBackground(p[2]);
    sheet.getRange(row, 7).setFormula(`=COUNTIF('${tl}'!F3:F${lastRow},"${p[1]}")`);
    sheet.getRange(row, 8).setFormula(`=COUNTIFS('${tl}'!F3:F${lastRow},"${p[1]}",'${tl}'!G3:G${lastRow},"<>Finished",'${tl}'!G3:G${lastRow},"<>Closed")`);
    sheet.getRange(row, 9).setFormula(`=IFERROR(G${row}/$A$4,0)`).setNumberFormat("0.0%");
    sheet.getRange(row, 10).setFormula(`=IF(H${row}>0,"⚠️ "&H${row}&" task cần xử lý","")`);
  });
  
  sheet.getRange(12, 6).setValue("TỔNG").setFontWeight("bold");
  sheet.getRange(12, 7).setFormula("=SUM(G8:G11)").setFontWeight("bold");
  sheet.getRange(12, 8).setFormula("=SUM(H8:H11)").setFontWeight("bold");
  sheet.getRange("F7:J12").setBorder(true, true, true, true, true, true);
  
  // ===== ASSIGNEE STATS (dùng COUNTIF với wildcard) =====
  sheet.getRange("A16").setValue("👥 THỐNG KÊ THEO NGƯỜI THỰC HIỆN").setFontSize(12).setFontWeight("bold");
  sheet.getRange("A16:C16").merge().setBackground("#9c27b0").setFontColor("white");
  
  sheet.getRange("A17:C17").setValues([["Assignee", "Số Task", "Biểu đồ"]]);
  sheet.getRange("A17:C17").setFontWeight("bold").setBackground("#f3e5f5");
  
  CONFIG.assignees.forEach((name, i) => {
    const row = 18 + i;
    sheet.getRange(row, 1).setValue(name);
    // Dùng COUNTIF với wildcard * để tìm trong multiple select
    sheet.getRange(row, 2).setFormula(`=COUNTIF('${tl}'!H3:H${lastRow},"*${name}*")`);
    sheet.getRange(row, 3).setFormula(`=REPT("█",B${row})&" ("&B${row}&")")`).setFontSize(9);
  });
  
  const assigneeEndRow = 17 + CONFIG.assignees.length;
  sheet.getRange(`A17:C${assigneeEndRow}`).setBorder(true, true, true, true, true, true);
  
  // ===== UPCOMING DEADLINES =====
  sheet.getRange("E16").setValue("⏰ TASK SẮP/QUÁ HẾT HẠN").setFontSize(12).setFontWeight("bold");
  sheet.getRange("E16:K16").merge().setBackground("#f57c00").setFontColor("white");
  
  sheet.getRange("E17:K17").setValues([["FNo.", "Task Name", "Assignee", "Priority", "End Date", "Còn lại", "Status"]]);
  sheet.getRange("E17:K17").setFontWeight("bold").setBackground("#fff3e0");
  
  // Filter task sắp hết hạn (End Date <= TODAY + 3)
  sheet.getRange("E18").setFormula(`=IFERROR(SORT(FILTER({'${tl}'!A3:A${lastRow},'${tl}'!B3:B${lastRow},'${tl}'!H3:H${lastRow},'${tl}'!F3:F${lastRow},'${tl}'!D3:D${lastRow},'${tl}'!E3:E${lastRow},'${tl}'!G3:G${lastRow}},('${tl}'!D3:D${lastRow}<>"")*(('${tl}'!D3:D${lastRow}<=TODAY()+3)+('${tl}'!D3:D${lastRow}<TODAY()))*('${tl}'!G3:G${lastRow}<>"Finished")*('${tl}'!G3:G${lastRow}<>"Closed")*('${tl}'!A3:A${lastRow}<>"")),5,TRUE),"Không có task sắp hết hạn")`);
  
  sheet.getRange("E17:K27").setBorder(true, true, true, true, true, true);
  
  // Alert quá hạn
  sheet.getRange("E28").setFormula(`=IF(COUNTIFS('${tl}'!D3:D${lastRow},"<"&TODAY(),'${tl}'!G3:G${lastRow},"<>Finished",'${tl}'!G3:G${lastRow},"<>Closed",'${tl}'!A3:A${lastRow},"F*")>0,"🚨 CÓ "&COUNTIFS('${tl}'!D3:D${lastRow},"<"&TODAY(),'${tl}'!G3:G${lastRow},"<>Finished",'${tl}'!G3:G${lastRow},"<>Closed",'${tl}'!A3:A${lastRow},"F*")&" TASK ĐÃ QUÁ HẠN!","")`);
  sheet.getRange("E28").setFontSize(12).setFontWeight("bold").setFontColor("#d32f2f");
  
  // ===== DETAILED ASSIGNEE TABLE =====
  sheet.getRange("A30").setValue("📋 CHI TIẾT THEO TỪNG NGƯỜI - PHÂN TÍCH WORKLOAD").setFontSize(12).setFontWeight("bold");
  sheet.getRange("A30:K30").merge().setBackground("#1565c0").setFontColor("white");
  
  const detailHeaders = ["Assignee", "Tổng", "✅Done", "🔄Progress", "🧪Testing", "⏳Pending", "🔴Urgent", "🟠High", "🟡Normal", "🟢Low", "📝 Task đang làm"];
  sheet.getRange("A31:K31").setValues([detailHeaders]);
  sheet.getRange("A31:K31").setFontWeight("bold").setBackground("#e3f2fd").setFontSize(9);
  
  CONFIG.assignees.forEach((name, i) => {
    const row = 32 + i;
    
    sheet.getRange(row, 1).setValue(name);
    
    // Tổng (dùng COUNTIF với wildcard)
    sheet.getRange(row, 2).setFormula(`=COUNTIF('${tl}'!H3:H${lastRow},"*${name}*")`);
    
    // Done
    sheet.getRange(row, 3).setFormula(`=COUNTIFS('${tl}'!H3:H${lastRow},"*${name}*",'${tl}'!G3:G${lastRow},"Finished")+COUNTIFS('${tl}'!H3:H${lastRow},"*${name}*",'${tl}'!G3:G${lastRow},"Closed")`);
    
    // In Progress
    sheet.getRange(row, 4).setFormula(`=COUNTIFS('${tl}'!H3:H${lastRow},"*${name}*",'${tl}'!G3:G${lastRow},"In Progress")`);
    
    // Testing
    sheet.getRange(row, 5).setFormula(`=COUNTIFS('${tl}'!H3:H${lastRow},"*${name}*",'${tl}'!G3:G${lastRow},"Testing")`);
    
    // Pending
    sheet.getRange(row, 6).setFormula(`=COUNTIFS('${tl}'!H3:H${lastRow},"*${name}*",'${tl}'!G3:G${lastRow},"Open")+COUNTIFS('${tl}'!H3:H${lastRow},"*${name}*",'${tl}'!G3:G${lastRow},"Pending")`);
    
    // Urgent (chưa xong)
    sheet.getRange(row, 7).setFormula(`=COUNTIFS('${tl}'!H3:H${lastRow},"*${name}*",'${tl}'!F3:F${lastRow},"Urgent",'${tl}'!G3:G${lastRow},"<>Finished",'${tl}'!G3:G${lastRow},"<>Closed")`);
    
    // High
    sheet.getRange(row, 8).setFormula(`=COUNTIFS('${tl}'!H3:H${lastRow},"*${name}*",'${tl}'!F3:F${lastRow},"High",'${tl}'!G3:G${lastRow},"<>Finished",'${tl}'!G3:G${lastRow},"<>Closed")`);
    
    // Normal
    sheet.getRange(row, 9).setFormula(`=COUNTIFS('${tl}'!H3:H${lastRow},"*${name}*",'${tl}'!F3:F${lastRow},"Normal",'${tl}'!G3:G${lastRow},"<>Finished",'${tl}'!G3:G${lastRow},"<>Closed")`);
    
    // Low
    sheet.getRange(row, 10).setFormula(`=COUNTIFS('${tl}'!H3:H${lastRow},"*${name}*",'${tl}'!F3:F${lastRow},"Low",'${tl}'!G3:G${lastRow},"<>Finished",'${tl}'!G3:G${lastRow},"<>Closed")`);
    
    // Task đang làm
    sheet.getRange(row, 11).setFormula(`=IFERROR(TEXTJOIN(", ",TRUE,FILTER('${tl}'!A3:A${lastRow}&": "&'${tl}'!B3:B${lastRow},('${tl}'!H3:H${lastRow}<>"")*(ISNUMBER(SEARCH("${name}",'${tl}'!H3:H${lastRow})))*('${tl}'!G3:G${lastRow}="In Progress"))),"Không có")`);
  });
  
  const detailEndRow = 31 + CONFIG.assignees.length;
  
  // Total row
  sheet.getRange(detailEndRow + 1, 1).setValue("TỔNG").setFontWeight("bold");
  for (let col = 2; col <= 10; col++) {
    const colLetter = String.fromCharCode(64 + col);
    sheet.getRange(detailEndRow + 1, col).setFormula(`=SUM(${colLetter}32:${colLetter}${detailEndRow})`).setFontWeight("bold");
  }
  
  sheet.getRange(`A31:K${detailEndRow + 1}`).setBorder(true, true, true, true, true, true);
  
  // Conditional formatting
  const urgentRange = sheet.getRange(`G32:G${detailEndRow}`);
  const highRange = sheet.getRange(`H32:H${detailEndRow}`);
  
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground("#ffcdd2")
      .setFontColor("#c62828")
      .setRanges([urgentRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground("#ffe0b2")
      .setFontColor("#e65100")
      .setRanges([highRange])
      .build()
  ];
  sheet.setConditionalFormatRules(rules);
  
  // ===== FORMATTING =====
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 70);
  sheet.setColumnWidth(3, 70);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 70);
  sheet.setColumnWidth(6, 150);
  sheet.setColumnWidth(7, 70);
  sheet.setColumnWidth(8, 70);
  sheet.setColumnWidth(9, 70);
  sheet.setColumnWidth(10, 70);
  sheet.setColumnWidth(11, 300);
  
  sheet.setFrozenRows(2);
  sheet.getRange("A1:K100").setFontFamily("Arial");
  
  // ===== CHARTS =====
  createCharts(sheet);
  
  SpreadsheetApp.getUi().alert('✅ Sheet "Overview" đã được tạo thành công!');
}

function createCharts(sheet) {
  // Pie chart cho Status
  const statusChart = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange("A8:B13"))
    .setPosition(6, 13, 0, 0)
    .setOption('title', 'Phân bổ theo Trạng thái')
    .setOption('pieHole', 0.4)
    .setOption('width', 350)
    .setOption('height', 250)
    .build();
  sheet.insertChart(statusChart);
  
  // Column chart cho Priority  
  const priorityChart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange("F8:H11"))
    .setPosition(16, 13, 0, 0)
    .setOption('title', 'Task theo Độ ưu tiên')
    .setOption('width', 350)
    .setOption('height', 250)
    .build();
  sheet.insertChart(priorityChart);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📊 Task Overview')
    .addItem('🔄 Tạo/Cập nhật Overview', 'createOverviewSheet')
    .addToUi();
}
