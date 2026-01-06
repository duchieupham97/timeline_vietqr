/**
 * 📊 TASK OVERVIEW - PREMIUM VERSION
 * Dashboard quản lý task chuyên nghiệp
 * - Realtime với công thức
 * - Biểu đồ đẹp
 * - Giao diện hiện đại
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
  // SECTION 1: HEADER & KPI CARDS
  // ============================================================
  
  // Title với gradient effect
  sheet.getRange("A1:N1").merge().setValue("📊 TASK OVERVIEW DASHBOARD")
    .setBackground("#0d47a1").setFontColor("white")
    .setFontSize(20).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.setRowHeight(1, 45);
  
  // Subtitle
  sheet.getRange("A2:N2").merge().setValue("Timeline 2601 - Realtime Statistics")
    .setBackground("#1565c0").setFontColor("#bbdefb")
    .setFontSize(11).setHorizontalAlignment("center");
  
  // KPI Cards Row
  sheet.setRowHeight(4, 30);
  sheet.setRowHeight(5, 50);
  
  // Card 1: Tổng Task
  createKPICard(sheet, "A4:B5", "📋 TỔNG TASK", "A5", 
    '=COUNTIF(\'Task List\'!G3:G500;"<>")', "#e3f2fd", "#1565c0");
  
  // Card 2: Hoàn thành
  createKPICard(sheet, "C4:D5", "✅ HOÀN THÀNH", "C5",
    '=COUNTIF(\'Task List\'!G3:G500;"Finished")+COUNTIF(\'Task List\'!G3:G500;"Closed")', "#e8f5e9", "#2e7d32");
  
  // Card 3: Đang làm
  createKPICard(sheet, "E4:F5", "🔄 ĐANG LÀM", "E5",
    '=COUNTIF(\'Task List\'!G3:G500;"In Progress")', "#fff3e0", "#ef6c00");
  
  // Card 4: Testing
  createKPICard(sheet, "G4:H5", "🧪 TESTING", "G5",
    '=COUNTIF(\'Task List\'!G3:G500;"Testing")', "#f3e5f5", "#7b1fa2");
  
  // Card 5: Chờ xử lý
  createKPICard(sheet, "I4:J5", "⏳ CHỜ XỬ LÝ", "I5",
    '=COUNTIF(\'Task List\'!G3:G500;"Open")+COUNTIF(\'Task List\'!G3:G500;"Pending")', "#fce4ec", "#c2185b");
  
  // Card 6: Urgent
  createKPICard(sheet, "K4:L5", "🚨 URGENT", "K5",
    '=COUNTIFS(\'Task List\'!F3:F500;"Urgent";\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed")', "#ffebee", "#c62828");
  
  // Card 7: % Hoàn thành
  createKPICard(sheet, "M4:N5", "📈 TIẾN ĐỘ", "M5",
    '=IFERROR(C5/A5;0)', "#e0f7fa", "#00838f", true);
  
  // ============================================================
  // SECTION 2: ALERT BOX
  // ============================================================
  
  sheet.getRange("A7:N7").merge().setValue("⚠️ CẢNH BÁO")
    .setBackground("#ff8a65").setFontColor("white")
    .setFontWeight("bold").setHorizontalAlignment("center");
  
  sheet.getRange("A8:D8").merge();
  setF(sheet, "A8", '=IF(K5>0;"🔴 "&K5&" task URGENT cần xử lý!";"")');
  sheet.getRange("A8").setFontColor("#c62828").setFontWeight("bold");
  
  sheet.getRange("E8:H8").merge();
  setF(sheet, "E8", '=IF(COUNTIFS(\'Task List\'!D3:D500;"<"&TODAY();\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed";\'Task List\'!D3:D500;"<>")>0;"⏰ "&COUNTIFS(\'Task List\'!D3:D500;"<"&TODAY();\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed";\'Task List\'!D3:D500;"<>")&" task QUÁ HẠN!";"")');
  sheet.getRange("E8").setFontColor("#d84315").setFontWeight("bold");
  
  sheet.getRange("I8:N8").merge();
  setF(sheet, "I8", '=IF(COUNTIFS(\'Task List\'!D3:D500;">="&TODAY();\'Task List\'!D3:D500;"<="&TODAY()+3;\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed";\'Task List\'!D3:D500;"<>")>0;"📅 "&COUNTIFS(\'Task List\'!D3:D500;">="&TODAY();\'Task List\'!D3:D500;"<="&TODAY()+3;\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed";\'Task List\'!D3:D500;"<>")&" task sắp hết hạn (3 ngày)";"✅ Không có cảnh báo")');
  sheet.getRange("I8").setFontColor("#f57c00").setFontWeight("bold");
  
  sheet.getRange("A7:N8").setBorder(true, true, true, true, true, true);
  
  // ============================================================
  // SECTION 3: STATUS & PRIORITY STATISTICS
  // ============================================================
  
  // --- STATUS ---
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
    sheet.getRange(r, 4).setFontSize(9);
    setF(sheet, "E" + r, '=IF(B' + r + '>0;B' + r + ';"")');
  });
  
  sheet.getRange(18, 1).setValue("TỔNG").setFontWeight("bold");
  setF(sheet, "B18", '=SUM(B12:B17)');
  sheet.getRange(18, 2).setFontWeight("bold");
  sheet.getRange("A11:E18").setBorder(true, true, true, true, true, true);
  
  // --- PRIORITY ---
  sheet.getRange("G10:K10").merge().setValue("🎯 THỐNG KÊ THEO ĐỘ ƯU TIÊN")
    .setBackground("#e53935").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  
  sheet.getRange("G11:K11").setValues([["Độ ưu tiên", "Tổng", "Chưa xong", "%", "⚠️"]]);
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
    setF(sheet, "K" + r, '=IF(I' + r + '>0;"⚠️ "&I' + r + '";"✅")');
  });
  
  sheet.getRange(16, 7).setValue("TỔNG").setFontWeight("bold");
  setF(sheet, "H16", '=SUM(H12:H15)');
  setF(sheet, "I16", '=SUM(I12:I15)');
  sheet.getRange(16, 8).setFontWeight("bold");
  sheet.getRange(16, 9).setFontWeight("bold");
  sheet.getRange("G11:K16").setBorder(true, true, true, true, true, true);
  
  // ============================================================
  // SECTION 4: ASSIGNEE SUMMARY & WORKLOAD
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
    setF(sheet, "E" + r, '=REPT("█";B' + r + ')');
    sheet.getRange(r, 5).setFontColor("#7b1fa2").setFontSize(9);
  });
  
  const assEndRow = 21 + assignees.length;
  sheet.getRange("A21:E" + assEndRow).setBorder(true, true, true, true, true, true);
  
  // ============================================================
  // SECTION 5: TOP PERFORMERS & DEADLINE
  // ============================================================
  
  // --- TOP PERFORMERS ---
  sheet.getRange("G20:K20").merge().setValue("🏆 TOP PERFORMERS")
    .setBackground("#ff6f00").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  
  sheet.getRange("G21:K21").setValues([["#", "Assignee", "Done", "%", "🏅"]]);
  sheet.getRange("G21:K21").setFontWeight("bold").setBackground("#ffe0b2");
  
  // Top 5 performers - sử dụng LARGE và INDEX/MATCH
  for (let i = 1; i <= 5; i++) {
    const r = 21 + i;
    sheet.getRange(r, 7).setValue(i);
    setF(sheet, "H" + r, '=IFERROR(INDEX($A$22:$A$' + assEndRow + ';MATCH(LARGE($C$22:$C$' + assEndRow + ';' + i + ');$C$22:$C$' + assEndRow + ';0));"")');
    setF(sheet, "I" + r, '=IFERROR(LARGE($C$22:$C$' + assEndRow + ';' + i + ');"")');
    setF(sheet, "J" + r, '=IFERROR(INDEX($D$22:$D$' + assEndRow + ';MATCH(H' + r + ';$A$22:$A$' + assEndRow + ';0));"")');
    sheet.getRange(r, 10).setNumberFormat("0%");
    const medal = i === 1 ? "🥇" : (i === 2 ? "🥈" : (i === 3 ? "🥉" : ""));
    sheet.getRange(r, 11).setValue(medal);
  }
  sheet.getRange("G21:K27").setBorder(true, true, true, true, true, true);
  
  // ============================================================
  // SECTION 6: TASK SẮP HẾT HẠN
  // ============================================================
  
  sheet.getRange("M10:N10").merge().setValue("⏰ SẮP HẾT HẠN")
    .setBackground("#d84315").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  
  sheet.getRange("M11:N11").setValues([["Task", "End Date"]]);
  sheet.getRange("M11:N11").setFontWeight("bold").setBackground("#ffccbc");
  
  setF(sheet, "M12", '=IFERROR(FILTER(\'Task List\'!A3:A500&" - "&\'Task List\'!B3:B500;(\'Task List\'!D3:D500>=TODAY())*(\'Task List\'!D3:D500<=TODAY()+3)*(\'Task List\'!G3:G500<>"Finished")*(\'Task List\'!G3:G500<>"Closed")*(\'Task List\'!D3:D500<>""));"Không có")');
  setF(sheet, "N12", '=IFERROR(FILTER(\'Task List\'!D3:D500;(\'Task List\'!D3:D500>=TODAY())*(\'Task List\'!D3:D500<=TODAY()+3)*(\'Task List\'!G3:G500<>"Finished")*(\'Task List\'!G3:G500<>"Closed")*(\'Task List\'!D3:D500<>""));"")');
  
  sheet.getRange("M11:N20").setBorder(true, true, true, true, true, true);
  
  // ============================================================
  // SECTION 7: CHI TIẾT WORKLOAD
  // ============================================================
  
  const detailStartRow = assEndRow + 3;
  
  sheet.getRange("A" + detailStartRow + ":N" + detailStartRow).merge()
    .setValue("📋 CHI TIẾT WORKLOAD THEO TỪNG NGƯỜI")
    .setBackground("#1565c0").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  
  const headerRow = detailStartRow + 1;
  const headers = ["Assignee", "Tổng", "✅Done", "🔄Progress", "🧪Testing", "⏳Pending", "🔴Urgent", "🟠High", "🟡Normal", "🟢Low", "Done%", "📝 Task đang làm"];
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
    setF(sheet, "L" + r, '=IFERROR(TEXTJOIN("; ";TRUE;FILTER(\'Task List\'!A3:A500&"-"&\'Task List\'!B3:B500;(ISNUMBER(SEARCH("' + name + '";\'Task List\'!H3:H500)))*(\'Task List\'!G3:G500="In Progress")));"Không có")');
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
      .whenNumberLessThan(0.3)
      .setBackground("#ffcdd2").setFontColor("#c62828")
      .setRanges([doneRange]).build()
  ]);
  
  // ============================================================
  // FORMATTING
  // ============================================================
  
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 60);
  sheet.setColumnWidth(3, 60);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 20);
  sheet.setColumnWidth(7, 100);
  sheet.setColumnWidth(8, 60);
  sheet.setColumnWidth(9, 70);
  sheet.setColumnWidth(10, 60);
  sheet.setColumnWidth(11, 60);
  sheet.setColumnWidth(12, 350);
  sheet.setColumnWidth(13, 150);
  sheet.setColumnWidth(14, 100);
  
  sheet.setFrozenRows(2);
  
  // ============================================================
  // CHARTS
  // ============================================================
  
  try {
    // Pie chart Status
    sheet.insertChart(sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sheet.getRange("A12:B17"))
      .setPosition(10, 15, 0, 0)
      .setOption('title', 'Phân bổ theo Status')
      .setOption('pieHole', 0.4)
      .setOption('width', 350)
      .setOption('height', 250)
      .setOption('colors', ['#4caf50', '#ffeb3b', '#2196f3', '#9c27b0', '#8bc34a', '#607d8b'])
      .build());
    
    // Bar chart Assignee Workload
    sheet.insertChart(sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(sheet.getRange("A22:B" + assEndRow))
      .setPosition(20, 15, 0, 0)
      .setOption('title', 'Workload theo Assignee')
      .setOption('width', 350)
      .setOption('height', 280)
      .setOption('legend', {position: 'none'})
      .setOption('colors', ['#7b1fa2'])
      .build());
      
    // Column chart Priority
    sheet.insertChart(sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sheet.getRange("G12:I15"))
      .setPosition(detailStartRow, 15, 0, 0)
      .setOption('title', 'Priority: Tổng vs Chưa xong')
      .setOption('width', 350)
      .setOption('height', 250)
      .setOption('colors', ['#9e9e9e', '#f44336'])
      .build());
      
  } catch(e) {}
  
  SpreadsheetApp.getUi().alert('✅ Tạo Overview Premium thành công!\n\n📊 Dashboard đẹp với biểu đồ\n🔄 Dữ liệu realtime\n⚠️ Cảnh báo tự động');
}

/**
 * Tạo KPI Card
 */
function createKPICard(sheet, range, title, valueCell, formula, bgColor, textColor, isPercent) {
  const cells = sheet.getRange(range);
  cells.merge().setBorder(true, true, true, true, null, null, "#bdbdbd", SpreadsheetApp.BorderStyle.SOLID);
  cells.setBackground(bgColor);
  
  const titleCell = range.split(":")[0];
  sheet.getRange(titleCell).setValue(title).setFontSize(10).setFontColor("#616161").setHorizontalAlignment("center");
  
  setF(sheet, valueCell, formula);
  sheet.getRange(valueCell).setFontSize(24).setFontWeight("bold").setFontColor(textColor).setHorizontalAlignment("center");
  
  if (isPercent) {
    sheet.getRange(valueCell).setNumberFormat("0%");
  }
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
