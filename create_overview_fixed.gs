/**
 * 📊 TASK OVERVIEW - FIXED VERSION
 * Đã sửa tất cả bugs
 */

function createOverviewSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Xóa sheet cũ
  let sheet = ss.getSheetByName("Overview");
  if (sheet) ss.deleteSheet(sheet);
  
  // Tạo mới
  sheet = ss.insertSheet("Overview");
  ss.moveActiveSheet(1);
  
  const assignees = ["Duy Anh", "Trường", "Đức", "Triều", "Nghĩa", "Hiếu Phạm", "Quyết", "Hiếu Hà", "Tôn"];
  
  // ============================================================
  // SECTION 1: HEADER
  // ============================================================
  
  sheet.getRange("A1:L1").merge().setValue("📊 TASK OVERVIEW DASHBOARD")
    .setBackground("#0d47a1").setFontColor("white")
    .setFontSize(18).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.setRowHeight(1, 40);
  
  sheet.getRange("A2:L2").merge().setValue("Timeline 2601 - Realtime Statistics")
    .setBackground("#1565c0").setFontColor("#bbdefb")
    .setFontSize(10).setHorizontalAlignment("center");
  
  // ============================================================
  // SECTION 2: KPI CARDS (Sửa lại - không merge phức tạp)
  // ============================================================
  
  // Header row
  sheet.getRange("A4").setValue("📋 TỔNG TASK");
  sheet.getRange("B4").setValue("✅ HOÀN THÀNH");
  sheet.getRange("C4").setValue("🔄 ĐANG LÀM");
  sheet.getRange("D4").setValue("🧪 TESTING");
  sheet.getRange("E4").setValue("⏳ CHỜ XỬ LÝ");
  sheet.getRange("F4").setValue("🚨 URGENT");
  sheet.getRange("G4").setValue("📈 TIẾN ĐỘ");
  sheet.getRange("A4:G4").setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center").setBackground("#e3f2fd");
  
  // Value row
  setF(sheet, "A5", '=COUNTIF(\'Task List\'!G3:G500;"<>")');
  setF(sheet, "B5", '=COUNTIF(\'Task List\'!G3:G500;"Finished")+COUNTIF(\'Task List\'!G3:G500;"Closed")');
  setF(sheet, "C5", '=COUNTIF(\'Task List\'!G3:G500;"In Progress")');
  setF(sheet, "D5", '=COUNTIF(\'Task List\'!G3:G500;"Testing")');
  setF(sheet, "E5", '=COUNTIF(\'Task List\'!G3:G500;"Open")+COUNTIF(\'Task List\'!G3:G500;"Pending")');
  setF(sheet, "F5", '=COUNTIFS(\'Task List\'!F3:F500;"Urgent";\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed")');
  setF(sheet, "G5", '=IFERROR(B5/A5;0)');
  
  sheet.getRange("A5:F5").setFontSize(22).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("G5").setFontSize(22).setFontWeight("bold").setHorizontalAlignment("center").setNumberFormat("0%");
  
  // Màu sắc cho từng KPI
  sheet.getRange("A4:A5").setBackground("#e3f2fd"); // Blue
  sheet.getRange("B4:B5").setBackground("#e8f5e9"); sheet.getRange("B5").setFontColor("#2e7d32"); // Green
  sheet.getRange("C4:C5").setBackground("#fff3e0"); sheet.getRange("C5").setFontColor("#ef6c00"); // Orange
  sheet.getRange("D4:D5").setBackground("#f3e5f5"); sheet.getRange("D5").setFontColor("#7b1fa2"); // Purple
  sheet.getRange("E4:E5").setBackground("#fce4ec"); sheet.getRange("E5").setFontColor("#c2185b"); // Pink
  sheet.getRange("F4:F5").setBackground("#ffebee"); sheet.getRange("F5").setFontColor("#c62828"); // Red
  sheet.getRange("G4:G5").setBackground("#e0f7fa"); sheet.getRange("G5").setFontColor("#00838f"); // Cyan
  
  sheet.getRange("A4:G5").setBorder(true, true, true, true, true, true);
  sheet.setRowHeight(5, 40);
  
  // ============================================================
  // SECTION 3: CẢNH BÁO
  // ============================================================
  
  sheet.getRange("A7:L7").merge().setValue("⚠️ CẢNH BÁO")
    .setBackground("#ff7043").setFontColor("white")
    .setFontWeight("bold").setHorizontalAlignment("center");
  
  sheet.getRange("A8:D8").merge();
  setF(sheet, "A8", '=IF(F5>0;"🔴 "&F5&" task URGENT cần xử lý!";"")');
  sheet.getRange("A8").setFontColor("#c62828").setFontWeight("bold");
  
  sheet.getRange("E8:H8").merge();
  setF(sheet, "E8", '=IFERROR(IF(COUNTIFS(\'Task List\'!D3:D500;"<"&TODAY();\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed";\'Task List\'!D3:D500;"<>")>0;"⏰ "&COUNTIFS(\'Task List\'!D3:D500;"<"&TODAY();\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed";\'Task List\'!D3:D500;"<>")&" task QUÁ HẠN!";"");"")');
  sheet.getRange("E8").setFontColor("#d84315").setFontWeight("bold");
  
  sheet.getRange("I8:L8").merge();
  setF(sheet, "I8", '=IFERROR(IF(COUNTIFS(\'Task List\'!D3:D500;">="&TODAY();\'Task List\'!D3:D500;"<="&TODAY()+3;\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed";\'Task List\'!D3:D500;"<>")>0;"📅 "&COUNTIFS(\'Task List\'!D3:D500;">="&TODAY();\'Task List\'!D3:D500;"<="&TODAY()+3;\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed";\'Task List\'!D3:D500;"<>")&" task sắp hết hạn";"✅ Không có cảnh báo");"✅ OK")');
  sheet.getRange("I8").setFontColor("#2e7d32").setFontWeight("bold");
  
  sheet.getRange("A7:L8").setBorder(true, true, true, true, true, true);
  
  // ============================================================
  // SECTION 4: STATUS STATISTICS
  // ============================================================
  
  sheet.getRange("A10:E10").merge().setValue("📈 THỐNG KÊ THEO TRẠNG THÁI")
    .setBackground("#43a047").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  
  sheet.getRange("A11:E11").setValues([["Trạng thái", "SL", "%", "Tiến độ", ""]]);
  sheet.getRange("A11:E11").setFontWeight("bold").setBackground("#c8e6c9");
  
  const statuses = [
    {name: "Open", icon: "🟢", color: "#e8f5e9"},
    {name: "Pending", icon: "🟡", color: "#fff8e1"},
    {name: "In Progress", icon: "🔵", color: "#e3f2fd"},
    {name: "Testing", icon: "🟣", color: "#f3e5f5"},
    {name: "Finished", icon: "✅", color: "#e8f5e9"},
    {name: "Closed", icon: "⬛", color: "#eceff1"}
  ];
  
  statuses.forEach((s, i) => {
    const r = 12 + i;
    sheet.getRange(r, 1).setValue(s.icon + " " + s.name).setBackground(s.color);
    setF(sheet, "B" + r, '=COUNTIF(\'Task List\'!G3:G500;"' + s.name + '")');
    setF(sheet, "C" + r, '=IFERROR(B' + r + '/$B$18;0)');
    sheet.getRange(r, 3).setNumberFormat("0%");
    setF(sheet, "D" + r, '=REPT("▓";ROUND(C' + r + '*10))&REPT("░";10-ROUND(C' + r + '*10))');
    sheet.getRange(r, 4).setFontSize(9).setFontColor("#43a047");
    setF(sheet, "E" + r, '=IF(B' + r + '>0;B' + r + ';"")');
  });
  
  sheet.getRange(18, 1).setValue("TỔNG").setFontWeight("bold");
  setF(sheet, "B18", '=SUM(B12:B17)');
  sheet.getRange(18, 2).setFontWeight("bold");
  sheet.getRange(18, 3).setValue("100%").setFontWeight("bold");
  sheet.getRange("A11:E18").setBorder(true, true, true, true, true, true);
  
  // ============================================================
  // SECTION 5: PRIORITY STATISTICS (Sửa cột cảnh báo)
  // ============================================================
  
  sheet.getRange("G10:K10").merge().setValue("🎯 THỐNG KÊ THEO ĐỘ ƯU TIÊN")
    .setBackground("#e53935").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  
  sheet.getRange("G11:K11").setValues([["Độ ưu tiên", "Tổng", "Chưa xong", "%", "Cảnh báo"]]);
  sheet.getRange("G11:K11").setFontWeight("bold").setBackground("#ffcdd2");
  
  const priorities = [
    {name: "Urgent", icon: "🔴", color: "#ffcdd2"},
    {name: "High", icon: "🟠", color: "#ffe0b2"},
    {name: "Normal", icon: "🟡", color: "#fff9c4"},
    {name: "Low", icon: "🟢", color: "#c8e6c9"}
  ];
  
  priorities.forEach((p, i) => {
    const r = 12 + i;
    sheet.getRange(r, 7).setValue(p.icon + " " + p.name).setBackground(p.color);
    setF(sheet, "H" + r, '=COUNTIF(\'Task List\'!F3:F500;"' + p.name + '")');
    setF(sheet, "I" + r, '=COUNTIFS(\'Task List\'!F3:F500;"' + p.name + '";\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed")');
    setF(sheet, "J" + r, '=IFERROR(H' + r + '/$H$16;0)');
    sheet.getRange(r, 10).setNumberFormat("0%");
    // SỬA: Công thức cột cảnh báo đơn giản hơn
    setF(sheet, "K" + r, '=IF(I' + r + '>0;"⚠️ "&I' + r + ';"✅")');
  });
  
  sheet.getRange(16, 7).setValue("TỔNG").setFontWeight("bold");
  setF(sheet, "H16", '=SUM(H12:H15)');
  setF(sheet, "I16", '=SUM(I12:I15)');
  sheet.getRange(16, 8).setFontWeight("bold");
  sheet.getRange(16, 9).setFontWeight("bold");
  sheet.getRange("G11:K16").setBorder(true, true, true, true, true, true);
  
  // ============================================================
  // SECTION 6: ASSIGNEE SUMMARY
  // ============================================================
  
  sheet.getRange("A20:E20").merge().setValue("👥 THỐNG KÊ THEO NGƯỜI THỰC HIỆN")
    .setBackground("#7b1fa2").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  
  sheet.getRange("A21:E21").setValues([["Assignee", "Task", "Done", "%Done", "Workload"]]);
  sheet.getRange("A21:E21").setFontWeight("bold").setBackground("#e1bee7");
  
  assignees.forEach((name, i) => {
    const r = 22 + i;
    sheet.getRange(r, 1).setValue(name);
    setF(sheet, "B" + r, '=COUNTIF(\'Task List\'!H3:H500;"*' + name + '*")');
    setF(sheet, "C" + r, '=COUNTIFS(\'Task List\'!H3:H500;"*' + name + '*";\'Task List\'!G3:G500;"Finished")+COUNTIFS(\'Task List\'!H3:H500;"*' + name + '*";\'Task List\'!G3:G500;"Closed")');
    setF(sheet, "D" + r, '=IFERROR(C' + r + '/B' + r + ';0)');
    sheet.getRange(r, 4).setNumberFormat("0%");
    setF(sheet, "E" + r, '=IF(B' + r + '>0;REPT("█";B' + r + ');"")');
    sheet.getRange(r, 5).setFontColor("#7b1fa2").setFontSize(9);
  });
  
  const assEndRow = 21 + assignees.length;
  sheet.getRange("A21:E" + assEndRow).setBorder(true, true, true, true, true, true);
  
  // ============================================================
  // SECTION 7: TOP PERFORMERS (Sửa - để rỗng khi không có data)
  // ============================================================
  
  sheet.getRange("G20:K20").merge().setValue("🏆 TOP PERFORMERS")
    .setBackground("#ff6f00").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  
  sheet.getRange("G21:K21").setValues([["#", "Assignee", "Done", "%Done", "🏅"]]);
  sheet.getRange("G21:K21").setFontWeight("bold").setBackground("#ffe0b2");
  
  // Top 5 - SỬA: dùng IFERROR để trả về rỗng khi không có data
  for (let i = 1; i <= 5; i++) {
    const r = 21 + i;
    sheet.getRange(r, 7).setValue(i);
    // Chỉ hiển thị nếu có ít nhất i người có done > 0
    setF(sheet, "H" + r, '=IFERROR(IF(LARGE($C$22:$C$' + assEndRow + ';' + i + ')>0;INDEX($A$22:$A$' + assEndRow + ';MATCH(LARGE($C$22:$C$' + assEndRow + ';' + i + ');$C$22:$C$' + assEndRow + ';0));"");"")');
    setF(sheet, "I" + r, '=IFERROR(IF(LARGE($C$22:$C$' + assEndRow + ';' + i + ')>0;LARGE($C$22:$C$' + assEndRow + ';' + i + ');"");"")');
    setF(sheet, "J" + r, '=IFERROR(IF(H' + r + '<>"";INDEX($D$22:$D$' + assEndRow + ';MATCH(H' + r + ';$A$22:$A$' + assEndRow + ';0));"");"")');
    sheet.getRange(r, 10).setNumberFormat("0%");
    // Medal chỉ hiển thị khi có data
    setF(sheet, "K" + r, '=IF(H' + r + '<>"";IF(' + i + '=1;"🥇";IF(' + i + '=2;"🥈";IF(' + i + '=3;"🥉";"")));"")');
  }
  sheet.getRange("G21:K27").setBorder(true, true, true, true, true, true);
  
  // ============================================================
  // SECTION 8: CHI TIẾT WORKLOAD (Sửa layout cột)
  // ============================================================
  
  const detailStartRow = 34;
  
  sheet.getRange("A" + detailStartRow + ":L" + detailStartRow).merge()
    .setValue("📋 CHI TIẾT WORKLOAD THEO TỪNG NGƯỜI")
    .setBackground("#1565c0").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  
  const headerRow = detailStartRow + 1;
  const headers = ["Assignee", "Tổng", "Done", "Progress", "Testing", "Pending", "Urgent", "High", "Normal", "Low", "Done%", "Task đang làm"];
  sheet.getRange("A" + headerRow + ":L" + headerRow).setValues([headers]);
  sheet.getRange("A" + headerRow + ":L" + headerRow).setFontWeight("bold").setBackground("#bbdefb").setFontSize(9);
  
  assignees.forEach((name, i) => {
    const r = headerRow + 1 + i;
    sheet.getRange(r, 1).setValue(name);
    
    setF(sheet, "B" + r, '=COUNTIF(\'Task List\'!H3:H500;"*' + name + '*")');
    setF(sheet, "C" + r, '=COUNTIFS(\'Task List\'!H3:H500;"*' + name + '*";\'Task List\'!G3:G500;"Finished")+COUNTIFS(\'Task List\'!H3:H500;"*' + name + '*";\'Task List\'!G3:G500;"Closed")');
    setF(sheet, "D" + r, '=COUNTIFS(\'Task List\'!H3:H500;"*' + name + '*";\'Task List\'!G3:G500;"In Progress")');
    setF(sheet, "E" + r, '=COUNTIFS(\'Task List\'!H3:H500;"*' + name + '*";\'Task List\'!G3:G500;"Testing")');
    setF(sheet, "F" + r, '=COUNTIFS(\'Task List\'!H3:H500;"*' + name + '*";\'Task List\'!G3:G500;"Open")+COUNTIFS(\'Task List\'!H3:H500;"*' + name + '*";\'Task List\'!G3:G500;"Pending")');
    setF(sheet, "G" + r, '=COUNTIFS(\'Task List\'!H3:H500;"*' + name + '*";\'Task List\'!F3:F500;"Urgent";\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed")');
    setF(sheet, "H" + r, '=COUNTIFS(\'Task List\'!H3:H500;"*' + name + '*";\'Task List\'!F3:F500;"High";\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed")');
    setF(sheet, "I" + r, '=COUNTIFS(\'Task List\'!H3:H500;"*' + name + '*";\'Task List\'!F3:F500;"Normal";\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed")');
    setF(sheet, "J" + r, '=COUNTIFS(\'Task List\'!H3:H500;"*' + name + '*";\'Task List\'!F3:F500;"Low";\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed")');
    setF(sheet, "K" + r, '=IFERROR(C' + r + '/B' + r + ';0)');
    sheet.getRange(r, 11).setNumberFormat("0%");
    setF(sheet, "L" + r, '=IFERROR(IF(D' + r + '>0;TEXTJOIN("; ";TRUE;FILTER(\'Task List\'!A3:A500&"-"&\'Task List\'!B3:B500;(ISNUMBER(SEARCH("' + name + '";\'Task List\'!H3:H500)))*(\'Task List\'!G3:G500="In Progress")));"Không có");"Không có")');
  });
  
  const detailEndRow = headerRow + assignees.length;
  
  // Total row
  sheet.getRange(detailEndRow + 1, 1).setValue("TỔNG").setFontWeight("bold");
  for (let col = 2; col <= 10; col++) {
    const letter = String.fromCharCode(64 + col);
    setF(sheet, letter + (detailEndRow + 1), '=SUM(' + letter + (headerRow + 1) + ':' + letter + detailEndRow + ')');
    sheet.getRange(detailEndRow + 1, col).setFontWeight("bold");
  }
  
  sheet.getRange("A" + headerRow + ":L" + (detailEndRow + 1)).setBorder(true, true, true, true, true, true);
  
  // Conditional formatting
  const urgentRange = sheet.getRange("G" + (headerRow + 1) + ":G" + detailEndRow);
  const highRange = sheet.getRange("H" + (headerRow + 1) + ":H" + detailEndRow);
  const doneRange = sheet.getRange("K" + (headerRow + 1) + ":K" + detailEndRow);
  
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground("#ffcdd2").setFontColor("#c62828")
      .setRanges([urgentRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground("#ffe0b2").setFontColor("#e65100")
      .setRanges([highRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(0.8)
      .setBackground("#c8e6c9").setFontColor("#2e7d32")
      .setRanges([doneRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(0.01, 0.29)
      .setBackground("#ffcdd2").setFontColor("#c62828")
      .setRanges([doneRange]).build()
  ]);
  
  // ============================================================
  // COLUMN WIDTHS (Sửa - điều chỉnh phù hợp)
  // ============================================================
  
  sheet.setColumnWidth(1, 100);  // A - Assignee
  sheet.setColumnWidth(2, 70);   // B
  sheet.setColumnWidth(3, 70);   // C
  sheet.setColumnWidth(4, 80);   // D
  sheet.setColumnWidth(5, 100);  // E - Workload bar
  sheet.setColumnWidth(6, 70);   // F - Pending (SỬA: tăng từ 20 lên 70)
  sheet.setColumnWidth(7, 100);  // G
  sheet.setColumnWidth(8, 60);   // H
  sheet.setColumnWidth(9, 70);   // I
  sheet.setColumnWidth(10, 60);  // J
  sheet.setColumnWidth(11, 60);  // K
  sheet.setColumnWidth(12, 350); // L - Task đang làm
  
  sheet.setFrozenRows(2);
  
  // ============================================================
  // CHARTS
  // ============================================================
  
  try {
    // Pie chart Status
    sheet.insertChart(sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sheet.getRange("A12:B17"))
      .setPosition(10, 13, 0, 0)
      .setOption('title', 'Phân bổ theo Status')
      .setOption('pieHole', 0.4)
      .setOption('width', 320)
      .setOption('height', 220)
      .setOption('colors', ['#4caf50', '#ffeb3b', '#2196f3', '#9c27b0', '#8bc34a', '#607d8b'])
      .build());
    
    // Bar chart Assignee
    sheet.insertChart(sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(sheet.getRange("A22:B" + assEndRow))
      .setPosition(20, 13, 0, 0)
      .setOption('title', 'Workload theo Assignee')
      .setOption('width', 320)
      .setOption('height', 250)
      .setOption('legend', {position: 'none'})
      .setOption('colors', ['#7b1fa2'])
      .build());
      
  } catch(e) {}
  
  SpreadsheetApp.getUi().alert('✅ Tạo Overview thành công!\n\n📊 Đã sửa tất cả bugs\n🔄 Dữ liệu realtime');
}

/**
 * Set formula
 */
function setF(sheet, cell, formula) {
  sheet.getRange(cell).setFormula(formula);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📊 Task Overview')
    .addItem('🔄 Tạo/Cập nhật Dashboard', 'createOverviewSheet')
    .addToUi();
}
