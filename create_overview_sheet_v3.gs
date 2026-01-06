/**
 * Google Apps Script - PHIÊN BẢN 3 (ĐÃ SỬA HOÀN TOÀN)
 * Cấu trúc: Header hàng 1, Dữ liệu từ hàng 2
 */

const CONFIG = {
  taskListSheetName: "Task List",
  overviewSheetName: "Overview",
  assignees: ["Duy Anh", "Trường", "Đức", "Triều", "Nghĩa", "Hiếu Phạm", "Quyết", "Hiếu Hà", "Tôn"],
  lastRow: 500
};

function createOverviewSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tl = CONFIG.taskListSheetName;
  const lr = CONFIG.lastRow;
  
  // Xóa sheet cũ
  let sheet = ss.getSheetByName(CONFIG.overviewSheetName);
  if (sheet) ss.deleteSheet(sheet);
  
  // Tạo mới
  sheet = ss.insertSheet(CONFIG.overviewSheetName);
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(1);
  
  // ========== TITLE ==========
  sheet.getRange("A1").setValue("📊 TASK OVERVIEW - TIMELINE 2601");
  sheet.getRange("A1:K1").merge().setBackground("#1a73e8").setFontColor("white")
    .setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
  
  // ========== KPI DASHBOARD ==========
  const kpiLabels = [["📋 Tổng Task", "✅ Hoàn thành", "🔄 Đang làm", "🧪 Testing", "⏳ Chờ xử lý", "🚨 Urgent", "📈 % Hoàn thành"]];
  sheet.getRange("A3:G3").setValues(kpiLabels);
  sheet.getRange("A3:G3").setFontWeight("bold").setBackground("#e8f0fe").setHorizontalAlignment("center");
  
  // KPI Values - QUAN TRỌNG: dùng COUNTA thay vì COUNTIF để đếm ô không rỗng
  // Và dùng range từ hàng 2 để bao gồm tất cả dữ liệu
  sheet.getRange("A4").setFormula(`=SUMPRODUCT(('${tl}'!A2:A${lr}<>"")*1)-COUNTIF('${tl}'!A2:A${lr},"*Support*")-COUNTIF('${tl}'!A2:A${lr},"*BackEnd*")-COUNTIF('${tl}'!A2:A${lr},"*Frontend*")`);
  sheet.getRange("B4").setFormula(`=COUNTIF('${tl}'!G2:G${lr},"Finished")+COUNTIF('${tl}'!G2:G${lr},"Closed")`);
  sheet.getRange("C4").setFormula(`=COUNTIF('${tl}'!G2:G${lr},"In Progress")`);
  sheet.getRange("D4").setFormula(`=COUNTIF('${tl}'!G2:G${lr},"Testing")`);
  sheet.getRange("E4").setFormula(`=COUNTIF('${tl}'!G2:G${lr},"Open")+COUNTIF('${tl}'!G2:G${lr},"Pending")`);
  sheet.getRange("F4").setFormula(`=COUNTIFS('${tl}'!F2:F${lr},"Urgent",'${tl}'!G2:G${lr},"<>Finished",'${tl}'!G2:G${lr},"<>Closed")`);
  sheet.getRange("G4").setFormula(`=IFERROR(B4/A4,0)`);
  
  sheet.getRange("A4:F4").setFontSize(20).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("G4").setFontSize(20).setFontWeight("bold").setHorizontalAlignment("center").setNumberFormat("0%");
  sheet.getRange("F4").setFontColor("#d93025");
  sheet.getRange("B4").setFontColor("#1e8e3e");
  sheet.getRange("A3:G4").setBorder(true, true, true, true, true, true);
  
  // ========== THỐNG KÊ STATUS ==========
  sheet.getRange("A6").setValue("📈 THỐNG KÊ THEO TRẠNG THÁI");
  sheet.getRange("A6:D6").merge().setBackground("#34a853").setFontColor("white").setFontWeight("bold");
  
  sheet.getRange("A7:D7").setValues([["Trạng thái", "Số lượng", "Phần trăm", ""]]);
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
    const row = 8 + i;
    sheet.getRange(row, 1).setValue(s[0]);
    sheet.getRange(row, 2).setFormula(`=COUNTIF('${tl}'!G2:G${lr},"${s[1]}")`);
    sheet.getRange(row, 3).setFormula(`=IFERROR(B${row}/B14,0)`).setNumberFormat("0%");
    sheet.getRange(row, 4).setFormula(`=REPT("▓",ROUND(C${row}*10))`).setFontColor("#34a853");
  });
  
  sheet.getRange(14, 1).setValue("TỔNG").setFontWeight("bold");
  sheet.getRange(14, 2).setFormula("=SUM(B8:B13)").setFontWeight("bold");
  sheet.getRange(14, 3).setValue("100%").setFontWeight("bold");
  sheet.getRange("A7:D14").setBorder(true, true, true, true, true, true);
  
  // ========== THỐNG KÊ PRIORITY ==========
  sheet.getRange("F6").setValue("🎯 THỐNG KÊ THEO ĐỘ ƯU TIÊN");
  sheet.getRange("F6:J6").merge().setBackground("#ea4335").setFontColor("white").setFontWeight("bold");
  
  sheet.getRange("F7:J7").setValues([["Độ ưu tiên", "Tổng", "Chưa xong", "Phần trăm", "Cảnh báo"]]);
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
    sheet.getRange(row, 7).setFormula(`=COUNTIF('${tl}'!F2:F${lr},"${p[1]}")`);
    sheet.getRange(row, 8).setFormula(`=COUNTIFS('${tl}'!F2:F${lr},"${p[1]}",'${tl}'!G2:G${lr},"<>Finished",'${tl}'!G2:G${lr},"<>Closed")`);
    sheet.getRange(row, 9).setFormula(`=IFERROR(G${row}/G12,0)`).setNumberFormat("0%");
    sheet.getRange(row, 10).setFormula(`=IF(H${row}>0,"⚠️ "&H${row}&" task","")`);
  });
  
  sheet.getRange(12, 6).setValue("TỔNG").setFontWeight("bold");
  sheet.getRange(12, 7).setFormula("=SUM(G8:G11)").setFontWeight("bold");
  sheet.getRange(12, 8).setFormula("=SUM(H8:H11)").setFontWeight("bold");
  sheet.getRange("F7:J12").setBorder(true, true, true, true, true, true);
  
  // ========== THỐNG KÊ ASSIGNEE ==========
  sheet.getRange("A16").setValue("👥 THỐNG KÊ THEO NGƯỜI THỰC HIỆN");
  sheet.getRange("A16:C16").merge().setBackground("#9c27b0").setFontColor("white").setFontWeight("bold");
  
  sheet.getRange("A17:C17").setValues([["Assignee", "Số Task", ""]]);
  sheet.getRange("A17:C17").setFontWeight("bold").setBackground("#f3e5f5");
  
  CONFIG.assignees.forEach((name, i) => {
    const row = 18 + i;
    sheet.getRange(row, 1).setValue(name);
    sheet.getRange(row, 2).setFormula(`=COUNTIF('${tl}'!H2:H${lr},"*${name}*")`);
    sheet.getRange(row, 3).setFormula(`=REPT("█",B${row})`).setFontColor("#9c27b0").setFontSize(9);
  });
  
  const assEndRow = 17 + CONFIG.assignees.length;
  sheet.getRange(`A17:C${assEndRow}`).setBorder(true, true, true, true, true, true);
  
  // ========== TASK SẮP HẾT HẠN ==========
  sheet.getRange("E16").setValue("⏰ TASK SẮP HẾT HẠN (3 ngày tới)");
  sheet.getRange("E16:K16").merge().setBackground("#f57c00").setFontColor("white").setFontWeight("bold");
  
  sheet.getRange("E17:K17").setValues([["FNo.", "Task", "Assignee", "Priority", "End Date", "Còn lại", "Status"]]);
  sheet.getRange("E17:K17").setFontWeight("bold").setBackground("#fff3e0");
  
  // Filter với điều kiện đơn giản hơn
  sheet.getRange("E18").setFormula(
    `=IFERROR(FILTER({'${tl}'!A2:A${lr},'${tl}'!B2:B${lr},'${tl}'!H2:H${lr},'${tl}'!F2:F${lr},'${tl}'!D2:D${lr},'${tl}'!E2:E${lr},'${tl}'!G2:G${lr}},` +
    `('${tl}'!D2:D${lr}<>"")*(('${tl}'!D2:D${lr}-TODAY())<=3)*(('${tl}'!D2:D${lr}-TODAY())>=-7)*` +
    `('${tl}'!G2:G${lr}<>"Finished")*('${tl}'!G2:G${lr}<>"Closed")*('${tl}'!A2:A${lr}<>"")),` +
    `"✅ Không có task sắp hết hạn")`
  );
  
  sheet.getRange("E17:K27").setBorder(true, true, true, true, true, true);
  
  // Alert quá hạn
  sheet.getRange("E28").setFormula(
    `=IF(COUNTIFS('${tl}'!D2:D${lr},"<"&TODAY(),'${tl}'!G2:G${lr},"<>Finished",'${tl}'!G2:G${lr},"<>Closed",'${tl}'!D2:D${lr},"<>")>0,` +
    `"🚨 CÓ "&COUNTIFS('${tl}'!D2:D${lr},"<"&TODAY(),'${tl}'!G2:G${lr},"<>Finished",'${tl}'!G2:G${lr},"<>Closed",'${tl}'!D2:D${lr},"<>")&" TASK QUÁ HẠN!","")`
  );
  sheet.getRange("E28").setFontWeight("bold").setFontColor("#d32f2f");
  
  // ========== BẢNG CHI TIẾT ASSIGNEE ==========
  sheet.getRange("A30").setValue("📋 CHI TIẾT WORKLOAD TỪNG NGƯỜI");
  sheet.getRange("A30:K30").merge().setBackground("#1565c0").setFontColor("white").setFontWeight("bold");
  
  const headers = ["Assignee", "Tổng", "Done", "Progress", "Testing", "Pending", "Urgent", "High", "Normal", "Low", "Task đang làm"];
  sheet.getRange("A31:K31").setValues([headers]);
  sheet.getRange("A31:K31").setFontWeight("bold").setBackground("#e3f2fd").setFontSize(9);
  
  CONFIG.assignees.forEach((name, i) => {
    const row = 32 + i;
    sheet.getRange(row, 1).setValue(name);
    
    // Tổng
    sheet.getRange(row, 2).setFormula(`=COUNTIF('${tl}'!H2:H${lr},"*${name}*")`);
    
    // Done
    sheet.getRange(row, 3).setFormula(`=COUNTIFS('${tl}'!H2:H${lr},"*${name}*",'${tl}'!G2:G${lr},"Finished")+COUNTIFS('${tl}'!H2:H${lr},"*${name}*",'${tl}'!G2:G${lr},"Closed")`);
    
    // In Progress
    sheet.getRange(row, 4).setFormula(`=COUNTIFS('${tl}'!H2:H${lr},"*${name}*",'${tl}'!G2:G${lr},"In Progress")`);
    
    // Testing
    sheet.getRange(row, 5).setFormula(`=COUNTIFS('${tl}'!H2:H${lr},"*${name}*",'${tl}'!G2:G${lr},"Testing")`);
    
    // Pending
    sheet.getRange(row, 6).setFormula(`=COUNTIFS('${tl}'!H2:H${lr},"*${name}*",'${tl}'!G2:G${lr},"Open")+COUNTIFS('${tl}'!H2:H${lr},"*${name}*",'${tl}'!G2:G${lr},"Pending")`);
    
    // Urgent (chưa xong)
    sheet.getRange(row, 7).setFormula(`=COUNTIFS('${tl}'!H2:H${lr},"*${name}*",'${tl}'!F2:F${lr},"Urgent",'${tl}'!G2:G${lr},"<>Finished",'${tl}'!G2:G${lr},"<>Closed")`);
    
    // High
    sheet.getRange(row, 8).setFormula(`=COUNTIFS('${tl}'!H2:H${lr},"*${name}*",'${tl}'!F2:F${lr},"High",'${tl}'!G2:G${lr},"<>Finished",'${tl}'!G2:G${lr},"<>Closed")`);
    
    // Normal
    sheet.getRange(row, 9).setFormula(`=COUNTIFS('${tl}'!H2:H${lr},"*${name}*",'${tl}'!F2:F${lr},"Normal",'${tl}'!G2:G${lr},"<>Finished",'${tl}'!G2:G${lr},"<>Closed")`);
    
    // Low
    sheet.getRange(row, 10).setFormula(`=COUNTIFS('${tl}'!H2:H${lr},"*${name}*",'${tl}'!F2:F${lr},"Low",'${tl}'!G2:G${lr},"<>Finished",'${tl}'!G2:G${lr},"<>Closed")`);
    
    // Task đang làm - dùng TEXTJOIN với FILTER và SEARCH
    sheet.getRange(row, 11).setFormula(
      `=IFERROR(TEXTJOIN(", ",TRUE,FILTER('${tl}'!A2:A${lr}&"-"&'${tl}'!B2:B${lr},` +
      `(ISNUMBER(SEARCH("${name}",'${tl}'!H2:H${lr})))*('${tl}'!G2:G${lr}="In Progress"))),"Không có")`
    );
  });
  
  const detailEndRow = 31 + CONFIG.assignees.length;
  
  // TỔNG
  sheet.getRange(detailEndRow + 1, 1).setValue("TỔNG").setFontWeight("bold");
  for (let col = 2; col <= 10; col++) {
    const colLetter = String.fromCharCode(64 + col);
    sheet.getRange(detailEndRow + 1, col).setFormula(`=SUM(${colLetter}32:${colLetter}${detailEndRow})`).setFontWeight("bold");
  }
  
  sheet.getRange(`A31:K${detailEndRow + 1}`).setBorder(true, true, true, true, true, true);
  
  // Conditional formatting cho Urgent và High
  const urgentRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground("#ffcdd2")
    .setFontColor("#c62828")
    .setRanges([sheet.getRange(`G32:G${detailEndRow}`)])
    .build();
    
  const highRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground("#ffe0b2")
    .setFontColor("#e65100")
    .setRanges([sheet.getRange(`H32:H${detailEndRow}`)])
    .build();
    
  sheet.setConditionalFormatRules([urgentRule, highRule]);
  
  // ========== FORMATTING ==========
  sheet.setColumnWidths(1, 1, 90);
  sheet.setColumnWidths(2, 9, 65);
  sheet.setColumnWidth(11, 280);
  sheet.setFrozenRows(2);
  
  // ========== CHARTS ==========
  try {
    // Pie chart Status
    const chart1 = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sheet.getRange("A8:B13"))
      .setPosition(5, 12, 0, 0)
      .setOption('title', 'Phân bổ Status')
      .setOption('pieHole', 0.4)
      .setOption('width', 320)
      .setOption('height', 220)
      .build();
    sheet.insertChart(chart1);
    
    // Bar chart Assignee
    const chart2 = sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(sheet.getRange(`A18:B${assEndRow}`))
      .setPosition(17, 12, 0, 0)
      .setOption('title', 'Task theo Assignee')
      .setOption('width', 320)
      .setOption('height', 220)
      .setOption('legend', {position: 'none'})
      .build();
    sheet.insertChart(chart2);
  } catch(e) {
    // Bỏ qua nếu lỗi chart
  }
  
  SpreadsheetApp.getUi().alert('✅ Tạo Overview thành công!');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📊 Task Overview')
    .addItem('🔄 Tạo/Cập nhật Overview', 'createOverviewSheet')
    .addToUi();
}
