/**
 * PHIÊN BẢN REALTIME - DÙNG CÔNG THỨC
 * Bắt đầu từ hàng 3, dùng dấu , (Apps Script tự convert)
 */

function createOverviewSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tl = "Task List";
  
  // Xóa sheet cũ
  let sheet = ss.getSheetByName("Overview");
  if (sheet) ss.deleteSheet(sheet);
  
  // Tạo mới
  sheet = ss.insertSheet("Overview");
  ss.moveActiveSheet(1);
  
  // Danh sách assignees
  const assignees = ["Duy Anh", "Trường", "Đức", "Triều", "Nghĩa", "Hiếu Phạm", "Quyết", "Hiếu Hà", "Tôn"];
  
  // ===== TITLE =====
  sheet.getRange("A1").setValue("📊 TASK OVERVIEW - TIMELINE 2601");
  sheet.getRange("A1:K1").merge().setBackground("#1a73e8").setFontColor("white")
    .setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
  
  // ===== KPI =====
  sheet.getRange("A3:G3").setValues([["📋 Tổng Task", "✅ Hoàn thành", "🔄 Đang làm", "🧪 Testing", "⏳ Chờ xử lý", "🚨 Urgent", "📈 %"]]);
  sheet.getRange("A3:G3").setFontWeight("bold").setBackground("#e8f0fe").setHorizontalAlignment("center");
  
  // Công thức KPI - bắt đầu từ hàng 3
  sheet.getRange("A4").setFormula("=COUNTA('Task List'!A3:A500)");
  sheet.getRange("B4").setFormula("=COUNTIF('Task List'!G3:G500,\"Finished\")+COUNTIF('Task List'!G3:G500,\"Closed\")");
  sheet.getRange("C4").setFormula("=COUNTIF('Task List'!G3:G500,\"In Progress\")");
  sheet.getRange("D4").setFormula("=COUNTIF('Task List'!G3:G500,\"Testing\")");
  sheet.getRange("E4").setFormula("=COUNTIF('Task List'!G3:G500,\"Open\")+COUNTIF('Task List'!G3:G500,\"Pending\")");
  sheet.getRange("F4").setFormula("=COUNTIFS('Task List'!F3:F500,\"Urgent\",'Task List'!G3:G500,\"<>Finished\",'Task List'!G3:G500,\"<>Closed\")");
  sheet.getRange("G4").setFormula("=IFERROR(B4/A4,0)");
  
  sheet.getRange("A4:F4").setFontSize(20).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("G4").setFontSize(20).setFontWeight("bold").setNumberFormat("0%");
  sheet.getRange("B4").setFontColor("#1e8e3e");
  sheet.getRange("F4").setFontColor("#d93025");
  sheet.getRange("A3:G4").setBorder(true, true, true, true, true, true);
  
  // ===== STATUS =====
  sheet.getRange("A6").setValue("📈 THỐNG KÊ THEO TRẠNG THÁI");
  sheet.getRange("A6:D6").merge().setBackground("#34a853").setFontColor("white").setFontWeight("bold");
  sheet.getRange("A7:D7").setValues([["Trạng thái", "Số lượng", "%", ""]]);
  sheet.getRange("A7:D7").setFontWeight("bold").setBackground("#e6f4ea");
  
  const statuses = [
    ["🟢 Open", "Open"],
    ["🟡 Pending", "Pending"],
    ["🔵 In Progress", "In Progress"],
    ["🟣 Testing", "Testing"],
    ["✅ Finished", "Finished"],
    ["⬛ Closed", "Closed"]
  ];
  
  statuses.forEach((s, i) => {
    const r = 8 + i;
    sheet.getRange(r, 1).setValue(s[0]);
    sheet.getRange(r, 2).setFormula("=COUNTIF('Task List'!G3:G500,\"" + s[1] + "\")");
    sheet.getRange(r, 3).setFormula("=IFERROR(B" + r + "/$B$14,0)").setNumberFormat("0%");
    sheet.getRange(r, 4).setFormula("=REPT(\"▓\",ROUND(C" + r + "*10))").setFontColor("#34a853");
  });
  
  sheet.getRange(14, 1).setValue("TỔNG").setFontWeight("bold");
  sheet.getRange(14, 2).setFormula("=SUM(B8:B13)").setFontWeight("bold");
  sheet.getRange(14, 3).setValue("100%").setFontWeight("bold");
  sheet.getRange("A7:D14").setBorder(true, true, true, true, true, true);
  
  // ===== PRIORITY =====
  sheet.getRange("F6").setValue("🎯 THỐNG KÊ ĐỘ ƯU TIÊN");
  sheet.getRange("F6:I6").merge().setBackground("#ea4335").setFontColor("white").setFontWeight("bold");
  sheet.getRange("F7:I7").setValues([["Độ ưu tiên", "Tổng", "Chưa xong", "%"]]);
  sheet.getRange("F7:I7").setFontWeight("bold").setBackground("#fce8e6");
  
  const priorities = [
    ["🔴 Urgent", "Urgent", "#ffcdd2"],
    ["🟠 High", "High", "#ffe0b2"],
    ["🟡 Normal", "Normal", "#fff9c4"],
    ["🟢 Low", "Low", "#c8e6c9"]
  ];
  
  priorities.forEach((p, i) => {
    const r = 8 + i;
    sheet.getRange(r, 6).setValue(p[0]).setBackground(p[2]);
    sheet.getRange(r, 7).setFormula("=COUNTIF('Task List'!F3:F500,\"" + p[1] + "\")");
    sheet.getRange(r, 8).setFormula("=COUNTIFS('Task List'!F3:F500,\"" + p[1] + "\",'Task List'!G3:G500,\"<>Finished\",'Task List'!G3:G500,\"<>Closed\")");
    sheet.getRange(r, 9).setFormula("=IFERROR(G" + r + "/$G$12,0)").setNumberFormat("0%");
  });
  
  sheet.getRange(12, 6).setValue("TỔNG").setFontWeight("bold");
  sheet.getRange(12, 7).setFormula("=SUM(G8:G11)").setFontWeight("bold");
  sheet.getRange(12, 8).setFormula("=SUM(H8:H11)").setFontWeight("bold");
  sheet.getRange("F7:I12").setBorder(true, true, true, true, true, true);
  
  // ===== ASSIGNEE SUMMARY =====
  sheet.getRange("A16").setValue("👥 THỐNG KÊ THEO NGƯỜI");
  sheet.getRange("A16:C16").merge().setBackground("#9c27b0").setFontColor("white").setFontWeight("bold");
  sheet.getRange("A17:C17").setValues([["Assignee", "Task", ""]]);
  sheet.getRange("A17:C17").setFontWeight("bold").setBackground("#f3e5f5");
  
  assignees.forEach((name, i) => {
    const r = 18 + i;
    sheet.getRange(r, 1).setValue(name);
    sheet.getRange(r, 2).setFormula("=COUNTIF('Task List'!H3:H500,\"*" + name + "*\")");
    sheet.getRange(r, 3).setFormula("=REPT(\"█\",B" + r + ")").setFontColor("#9c27b0").setFontSize(9);
  });
  
  const assEndRow = 17 + assignees.length;
  sheet.getRange("A17:C" + assEndRow).setBorder(true, true, true, true, true, true);
  
  // ===== ASSIGNEE DETAIL =====
  sheet.getRange("A29").setValue("📋 CHI TIẾT WORKLOAD");
  sheet.getRange("A29:K29").merge().setBackground("#1565c0").setFontColor("white").setFontWeight("bold");
  
  const headers = ["Assignee", "Tổng", "Done", "Progress", "Testing", "Pending", "Urgent", "High", "Normal", "Low", "Task đang làm"];
  sheet.getRange("A30:K30").setValues([headers]);
  sheet.getRange("A30:K30").setFontWeight("bold").setBackground("#e3f2fd").setFontSize(9);
  
  assignees.forEach((name, i) => {
    const r = 31 + i;
    sheet.getRange(r, 1).setValue(name);
    
    // Tổng
    sheet.getRange(r, 2).setFormula("=COUNTIF('Task List'!H3:H500,\"*" + name + "*\")");
    
    // Done
    sheet.getRange(r, 3).setFormula("=COUNTIFS('Task List'!H3:H500,\"*" + name + "*\",'Task List'!G3:G500,\"Finished\")+COUNTIFS('Task List'!H3:H500,\"*" + name + "*\",'Task List'!G3:G500,\"Closed\")");
    
    // In Progress
    sheet.getRange(r, 4).setFormula("=COUNTIFS('Task List'!H3:H500,\"*" + name + "*\",'Task List'!G3:G500,\"In Progress\")");
    
    // Testing
    sheet.getRange(r, 5).setFormula("=COUNTIFS('Task List'!H3:H500,\"*" + name + "*\",'Task List'!G3:G500,\"Testing\")");
    
    // Pending
    sheet.getRange(r, 6).setFormula("=COUNTIFS('Task List'!H3:H500,\"*" + name + "*\",'Task List'!G3:G500,\"Open\")+COUNTIFS('Task List'!H3:H500,\"*" + name + "*\",'Task List'!G3:G500,\"Pending\")");
    
    // Urgent (chưa xong)
    sheet.getRange(r, 7).setFormula("=COUNTIFS('Task List'!H3:H500,\"*" + name + "*\",'Task List'!F3:F500,\"Urgent\",'Task List'!G3:G500,\"<>Finished\",'Task List'!G3:G500,\"<>Closed\")");
    
    // High
    sheet.getRange(r, 8).setFormula("=COUNTIFS('Task List'!H3:H500,\"*" + name + "*\",'Task List'!F3:F500,\"High\",'Task List'!G3:G500,\"<>Finished\",'Task List'!G3:G500,\"<>Closed\")");
    
    // Normal
    sheet.getRange(r, 9).setFormula("=COUNTIFS('Task List'!H3:H500,\"*" + name + "*\",'Task List'!F3:F500,\"Normal\",'Task List'!G3:G500,\"<>Finished\",'Task List'!G3:G500,\"<>Closed\")");
    
    // Low
    sheet.getRange(r, 10).setFormula("=COUNTIFS('Task List'!H3:H500,\"*" + name + "*\",'Task List'!F3:F500,\"Low\",'Task List'!G3:G500,\"<>Finished\",'Task List'!G3:G500,\"<>Closed\")");
    
    // Task đang làm
    sheet.getRange(r, 11).setFormula("=IFERROR(TEXTJOIN(\", \",TRUE,FILTER('Task List'!A3:A500&\": \"&'Task List'!B3:B500,(ISNUMBER(SEARCH(\"" + name + "\",'Task List'!H3:H500)))*('Task List'!G3:G500=\"In Progress\"))),\"Không có\")");
  });
  
  const detailEndRow = 30 + assignees.length;
  
  // Total row
  sheet.getRange(detailEndRow + 1, 1).setValue("TỔNG").setFontWeight("bold");
  for (let col = 2; col <= 10; col++) {
    const letter = String.fromCharCode(64 + col);
    sheet.getRange(detailEndRow + 1, col).setFormula("=SUM(" + letter + "31:" + letter + detailEndRow + ")").setFontWeight("bold");
  }
  
  sheet.getRange("A30:K" + (detailEndRow + 1)).setBorder(true, true, true, true, true, true);
  
  // Conditional formatting
  const urgentRange = sheet.getRange("G31:G" + detailEndRow);
  const highRange = sheet.getRange("H31:H" + detailEndRow);
  
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground("#ffcdd2").setFontColor("#c62828")
      .setRanges([urgentRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground("#ffe0b2").setFontColor("#e65100")
      .setRanges([highRange]).build()
  ]);
  
  // Formatting
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(11, 350);
  sheet.setFrozenRows(2);
  
  // Charts
  try {
    sheet.insertChart(sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sheet.getRange("A8:B13"))
      .setPosition(5, 11, 0, 0)
      .setOption('title', 'Status')
      .setOption('pieHole', 0.4)
      .setOption('width', 300)
      .setOption('height', 200)
      .build());
      
    sheet.insertChart(sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(sheet.getRange("A18:B" + assEndRow))
      .setPosition(16, 11, 0, 0)
      .setOption('title', 'Assignee')
      .setOption('width', 300)
      .setOption('height', 200)
      .setOption('legend', {position: 'none'})
      .build());
  } catch(e) {}
  
  SpreadsheetApp.getUi().alert('✅ Tạo Overview thành công!\n\nDữ liệu sẽ tự động cập nhật realtime.');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📊 Overview')
    .addItem('🔄 Tạo Overview', 'createOverviewSheet')
    .addToUi();
}
