/**
 * 📊 TASK OVERVIEW - CLEAN VERSION
 * - Độ rộng cột đều nhau
 * - Merge cells cho từng bảng
 * - Layout chuẩn
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
  // ĐẶT ĐỘ RỘNG CỘT ĐỀU NHAU
  // ============================================================
  const COL_WIDTH = 75;
  for (let i = 1; i <= 20; i++) {
    sheet.setColumnWidth(i, COL_WIDTH);
  }
  
  // ============================================================
  // SECTION 1: HEADER (Merge A1:T1)
  // ============================================================
  
  sheet.getRange("A1:T1").merge().setValue("📊 TASK OVERVIEW DASHBOARD - Timeline 2601")
    .setBackground("#0d47a1").setFontColor("white")
    .setFontSize(18).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.setRowHeight(1, 45);
  
  // ============================================================
  // SECTION 2: KPI CARDS (Row 3-4)
  // ============================================================
  
  // Card 1: Tổng Task (A3:C4)
  sheet.getRange("A3:C3").merge().setValue("📋 TỔNG TASK").setBackground("#e3f2fd")
    .setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("A4:C4").merge().setBackground("#e3f2fd");
  setF(sheet, "A4", '=COUNTIF(\'Task List\'!G3:G500;"<>")');
  sheet.getRange("A4").setFontSize(28).setFontWeight("bold").setFontColor("#1565c0").setHorizontalAlignment("center");
  
  // Card 2: Hoàn thành (D3:F4)
  sheet.getRange("D3:F3").merge().setValue("✅ HOÀN THÀNH").setBackground("#e8f5e9")
    .setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("D4:F4").merge().setBackground("#e8f5e9");
  setF(sheet, "D4", '=COUNTIF(\'Task List\'!G3:G500;"Finished")+COUNTIF(\'Task List\'!G3:G500;"Closed")');
  sheet.getRange("D4").setFontSize(28).setFontWeight("bold").setFontColor("#2e7d32").setHorizontalAlignment("center");
  
  // Card 3: Đang làm (G3:I4)
  sheet.getRange("G3:I3").merge().setValue("🔄 ĐANG LÀM").setBackground("#fff3e0")
    .setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("G4:I4").merge().setBackground("#fff3e0");
  setF(sheet, "G4", '=COUNTIF(\'Task List\'!G3:G500;"In Progress")');
  sheet.getRange("G4").setFontSize(28).setFontWeight("bold").setFontColor("#ef6c00").setHorizontalAlignment("center");
  
  // Card 4: Testing (J3:L4)
  sheet.getRange("J3:L3").merge().setValue("🧪 TESTING").setBackground("#f3e5f5")
    .setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("J4:L4").merge().setBackground("#f3e5f5");
  setF(sheet, "J4", '=COUNTIF(\'Task List\'!G3:G500;"Testing")');
  sheet.getRange("J4").setFontSize(28).setFontWeight("bold").setFontColor("#7b1fa2").setHorizontalAlignment("center");
  
  // Card 5: Chờ xử lý (M3:O4)
  sheet.getRange("M3:O3").merge().setValue("⏳ CHỜ XỬ LÝ").setBackground("#fce4ec")
    .setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("M4:O4").merge().setBackground("#fce4ec");
  setF(sheet, "M4", '=COUNTIF(\'Task List\'!G3:G500;"Open")+COUNTIF(\'Task List\'!G3:G500;"Pending")');
  sheet.getRange("M4").setFontSize(28).setFontWeight("bold").setFontColor("#c2185b").setHorizontalAlignment("center");
  
  // Card 6: Urgent (P3:R4)
  sheet.getRange("P3:R3").merge().setValue("🚨 URGENT").setBackground("#ffebee")
    .setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("P4:R4").merge().setBackground("#ffebee");
  setF(sheet, "P4", '=COUNTIFS(\'Task List\'!F3:F500;"Urgent";\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed")');
  sheet.getRange("P4").setFontSize(28).setFontWeight("bold").setFontColor("#c62828").setHorizontalAlignment("center");
  
  // Card 7: % Tiến độ (S3:T4)
  sheet.getRange("S3:T3").merge().setValue("📈 TIẾN ĐỘ").setBackground("#e0f7fa")
    .setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("S4:T4").merge().setBackground("#e0f7fa");
  setF(sheet, "S4", '=IFERROR(D4/A4;0)');
  sheet.getRange("S4").setFontSize(28).setFontWeight("bold").setFontColor("#00838f")
    .setHorizontalAlignment("center").setNumberFormat("0%");
  
  sheet.getRange("A3:T4").setBorder(true, true, true, true, true, true);
  sheet.setRowHeight(4, 50);
  
  // ============================================================
  // SECTION 3: CẢNH BÁO (Row 6)
  // ============================================================
  
  sheet.getRange("A6:T6").merge().setValue("⚠️ CẢNH BÁO").setBackground("#ff7043")
    .setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  
  sheet.getRange("A7:F7").merge().setBackground("#fff3e0");
  setF(sheet, "A7", '=IF(P4>0;"🔴 "&P4&" task URGENT cần xử lý ngay!";"✅ Không có task Urgent")');
  sheet.getRange("A7").setFontWeight("bold").setHorizontalAlignment("center");
  
  sheet.getRange("G7:M7").merge().setBackground("#fff3e0");
  setF(sheet, "G7", '=IFERROR(IF(COUNTIFS(\'Task List\'!D3:D500;"<"&TODAY();\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed";\'Task List\'!D3:D500;"<>")>0;"⏰ "&COUNTIFS(\'Task List\'!D3:D500;"<"&TODAY();\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed";\'Task List\'!D3:D500;"<>")&" task ĐÃ QUÁ HẠN!";"✅ Không có task quá hạn");"✅ OK")');
  sheet.getRange("G7").setFontWeight("bold").setHorizontalAlignment("center");
  
  sheet.getRange("N7:T7").merge().setBackground("#fff3e0");
  setF(sheet, "N7", '=IFERROR(IF(COUNTIFS(\'Task List\'!D3:D500;">="&TODAY();\'Task List\'!D3:D500;"<="&TODAY()+3;\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed";\'Task List\'!D3:D500;"<>")>0;"📅 "&COUNTIFS(\'Task List\'!D3:D500;">="&TODAY();\'Task List\'!D3:D500;"<="&TODAY()+3;\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed";\'Task List\'!D3:D500;"<>")&" task sắp hết hạn (3 ngày)";"✅ Không có task sắp hết hạn");"✅ OK")');
  sheet.getRange("N7").setFontWeight("bold").setHorizontalAlignment("center");
  
  sheet.getRange("A6:T7").setBorder(true, true, true, true, true, true);
  
  // ============================================================
  // SECTION 4: BẢNG STATUS (A9:G18) - Merge để độc lập
  // ============================================================
  
  sheet.getRange("A9:G9").merge().setValue("📈 THỐNG KÊ THEO TRẠNG THÁI")
    .setBackground("#43a047").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  
  // Header
  sheet.getRange("A10:C10").merge().setValue("Trạng thái").setFontWeight("bold").setBackground("#c8e6c9").setHorizontalAlignment("center");
  sheet.getRange("D10:E10").merge().setValue("Số lượng").setFontWeight("bold").setBackground("#c8e6c9").setHorizontalAlignment("center");
  sheet.getRange("F10:G10").merge().setValue("Phần trăm").setFontWeight("bold").setBackground("#c8e6c9").setHorizontalAlignment("center");
  
  const statuses = [
    {name: "Open", icon: "🟢", color: "#e8f5e9"},
    {name: "Pending", icon: "🟡", color: "#fff8e1"},
    {name: "In Progress", icon: "🔵", color: "#e3f2fd"},
    {name: "Testing", icon: "🟣", color: "#f3e5f5"},
    {name: "Finished", icon: "✅", color: "#e8f5e9"},
    {name: "Closed", icon: "⬛", color: "#eceff1"}
  ];
  
  statuses.forEach((s, i) => {
    const r = 11 + i;
    sheet.getRange("A" + r + ":C" + r).merge().setValue(s.icon + " " + s.name).setBackground(s.color);
    sheet.getRange("D" + r + ":E" + r).merge().setBackground(s.color).setHorizontalAlignment("center");
    setF(sheet, "D" + r, '=COUNTIF(\'Task List\'!G3:G500;"' + s.name + '")');
    sheet.getRange("F" + r + ":G" + r).merge().setBackground(s.color).setHorizontalAlignment("center");
    setF(sheet, "F" + r, '=IFERROR(D' + r + '/$D$17;0)');
    sheet.getRange("F" + r).setNumberFormat("0%");
  });
  
  // Total
  sheet.getRange("A17:C17").merge().setValue("TỔNG").setFontWeight("bold");
  sheet.getRange("D17:E17").merge().setFontWeight("bold").setHorizontalAlignment("center");
  setF(sheet, "D17", '=SUM(D11:D16)');
  sheet.getRange("F17:G17").merge().setValue("100%").setFontWeight("bold").setHorizontalAlignment("center");
  
  sheet.getRange("A9:G17").setBorder(true, true, true, true, true, true);
  
  // ============================================================
  // SECTION 5: BẢNG PRIORITY (I9:O16) - Merge để độc lập
  // ============================================================
  
  sheet.getRange("I9:O9").merge().setValue("🎯 THỐNG KÊ THEO ĐỘ ƯU TIÊN")
    .setBackground("#e53935").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  
  // Header
  sheet.getRange("I10:K10").merge().setValue("Độ ưu tiên").setFontWeight("bold").setBackground("#ffcdd2").setHorizontalAlignment("center");
  sheet.getRange("L10").setValue("Tổng").setFontWeight("bold").setBackground("#ffcdd2").setHorizontalAlignment("center");
  sheet.getRange("M10").setValue("Chưa xong").setFontWeight("bold").setBackground("#ffcdd2").setHorizontalAlignment("center");
  sheet.getRange("N10:O10").merge().setValue("Cảnh báo").setFontWeight("bold").setBackground("#ffcdd2").setHorizontalAlignment("center");
  
  const priorities = [
    {name: "Urgent", icon: "🔴", color: "#ffcdd2"},
    {name: "High", icon: "🟠", color: "#ffe0b2"},
    {name: "Normal", icon: "🟡", color: "#fff9c4"},
    {name: "Low", icon: "🟢", color: "#c8e6c9"}
  ];
  
  priorities.forEach((p, i) => {
    const r = 11 + i;
    sheet.getRange("I" + r + ":K" + r).merge().setValue(p.icon + " " + p.name).setBackground(p.color);
    sheet.getRange("L" + r).setBackground(p.color).setHorizontalAlignment("center");
    setF(sheet, "L" + r, '=COUNTIF(\'Task List\'!F3:F500;"' + p.name + '")');
    sheet.getRange("M" + r).setBackground(p.color).setHorizontalAlignment("center");
    setF(sheet, "M" + r, '=COUNTIFS(\'Task List\'!F3:F500;"' + p.name + '";\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed")');
    sheet.getRange("N" + r + ":O" + r).merge().setBackground(p.color).setHorizontalAlignment("center");
    setF(sheet, "N" + r, '=IF(M' + r + '>0;"⚠️ "&M' + r + '&" task";"✅")');
  });
  
  // Total
  sheet.getRange("I15:K15").merge().setValue("TỔNG").setFontWeight("bold");
  sheet.getRange("L15").setFontWeight("bold").setHorizontalAlignment("center");
  setF(sheet, "L15", '=SUM(L11:L14)');
  sheet.getRange("M15").setFontWeight("bold").setHorizontalAlignment("center");
  setF(sheet, "M15", '=SUM(M11:M14)');
  
  sheet.getRange("I9:O15").setBorder(true, true, true, true, true, true);
  
  // ============================================================
  // SECTION 6: BẢNG ASSIGNEE (Q9:T+) - Merge để độc lập
  // ============================================================
  
  sheet.getRange("Q9:T9").merge().setValue("👥 THỐNG KÊ THEO NGƯỜI")
    .setBackground("#7b1fa2").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  
  // Header
  sheet.getRange("Q10:R10").merge().setValue("Assignee").setFontWeight("bold").setBackground("#e1bee7").setHorizontalAlignment("center");
  sheet.getRange("S10").setValue("Task").setFontWeight("bold").setBackground("#e1bee7").setHorizontalAlignment("center");
  sheet.getRange("T10").setValue("Done").setFontWeight("bold").setBackground("#e1bee7").setHorizontalAlignment("center");
  
  assignees.forEach((name, i) => {
    const r = 11 + i;
    sheet.getRange("Q" + r + ":R" + r).merge().setValue(name);
    sheet.getRange("S" + r).setHorizontalAlignment("center");
    setF(sheet, "S" + r, '=COUNTIF(\'Task List\'!H3:H500;"*' + name + '*")');
    sheet.getRange("T" + r).setHorizontalAlignment("center");
    setF(sheet, "T" + r, '=COUNTIFS(\'Task List\'!H3:H500;"*' + name + '*";\'Task List\'!G3:G500;"Finished")+COUNTIFS(\'Task List\'!H3:H500;"*' + name + '*";\'Task List\'!G3:G500;"Closed")');
  });
  
  const assEndRow = 10 + assignees.length;
  sheet.getRange("Q9:T" + assEndRow).setBorder(true, true, true, true, true, true);
  
  // ============================================================
  // SECTION 7: TOP PERFORMERS (A19:G25)
  // ============================================================
  
  sheet.getRange("A19:G19").merge().setValue("🏆 TOP PERFORMERS")
    .setBackground("#ff6f00").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  
  sheet.getRange("A20:B20").merge().setValue("#").setFontWeight("bold").setBackground("#ffe0b2").setHorizontalAlignment("center");
  sheet.getRange("C20:E20").merge().setValue("Assignee").setFontWeight("bold").setBackground("#ffe0b2").setHorizontalAlignment("center");
  sheet.getRange("F20").setValue("Done").setFontWeight("bold").setBackground("#ffe0b2").setHorizontalAlignment("center");
  sheet.getRange("G20").setValue("🏅").setFontWeight("bold").setBackground("#ffe0b2").setHorizontalAlignment("center");
  
  for (let i = 1; i <= 5; i++) {
    const r = 20 + i;
    sheet.getRange("A" + r + ":B" + r).merge().setValue(i).setHorizontalAlignment("center");
    sheet.getRange("C" + r + ":E" + r).merge();
    setF(sheet, "C" + r, '=IFERROR(IF(LARGE($T$11:$T$' + assEndRow + ';' + i + ')>0;INDEX($Q$11:$Q$' + assEndRow + ';MATCH(LARGE($T$11:$T$' + assEndRow + ';' + i + ');$T$11:$T$' + assEndRow + ';0));"");"")');
    sheet.getRange("F" + r).setHorizontalAlignment("center");
    setF(sheet, "F" + r, '=IFERROR(IF(LARGE($T$11:$T$' + assEndRow + ';' + i + ')>0;LARGE($T$11:$T$' + assEndRow + ';' + i + ');"");"")');
    sheet.getRange("G" + r).setHorizontalAlignment("center");
    setF(sheet, "G" + r, '=IF(F' + r + '<>"";IF(' + i + '=1;"🥇";IF(' + i + '=2;"🥈";IF(' + i + '=3;"🥉";"")));"")');
  }
  
  sheet.getRange("A19:G25").setBorder(true, true, true, true, true, true);
  
  // ============================================================
  // SECTION 8: CHI TIẾT WORKLOAD (A27:T+)
  // ============================================================
  
  const detailStartRow = 27;
  
  sheet.getRange("A" + detailStartRow + ":T" + detailStartRow).merge()
    .setValue("📋 CHI TIẾT WORKLOAD THEO TỪNG NGƯỜI")
    .setBackground("#1565c0").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  
  const headerRow = detailStartRow + 1;
  
  // Headers với merge
  sheet.getRange("A" + headerRow + ":B" + headerRow).merge().setValue("Assignee").setFontWeight("bold").setBackground("#bbdefb").setHorizontalAlignment("center");
  sheet.getRange("C" + headerRow).setValue("Tổng").setFontWeight("bold").setBackground("#bbdefb").setHorizontalAlignment("center");
  sheet.getRange("D" + headerRow).setValue("Done").setFontWeight("bold").setBackground("#bbdefb").setHorizontalAlignment("center");
  sheet.getRange("E" + headerRow).setValue("Progress").setFontWeight("bold").setBackground("#bbdefb").setHorizontalAlignment("center");
  sheet.getRange("F" + headerRow).setValue("Testing").setFontWeight("bold").setBackground("#bbdefb").setHorizontalAlignment("center");
  sheet.getRange("G" + headerRow).setValue("Pending").setFontWeight("bold").setBackground("#bbdefb").setHorizontalAlignment("center");
  sheet.getRange("H" + headerRow).setValue("Urgent").setFontWeight("bold").setBackground("#ffcdd2").setHorizontalAlignment("center");
  sheet.getRange("I" + headerRow).setValue("High").setFontWeight("bold").setBackground("#ffe0b2").setHorizontalAlignment("center");
  sheet.getRange("J" + headerRow).setValue("Normal").setFontWeight("bold").setBackground("#fff9c4").setHorizontalAlignment("center");
  sheet.getRange("K" + headerRow).setValue("Low").setFontWeight("bold").setBackground("#c8e6c9").setHorizontalAlignment("center");
  sheet.getRange("L" + headerRow).setValue("Done%").setFontWeight("bold").setBackground("#bbdefb").setHorizontalAlignment("center");
  sheet.getRange("M" + headerRow + ":T" + headerRow).merge().setValue("📝 Task đang làm").setFontWeight("bold").setBackground("#bbdefb").setHorizontalAlignment("center");
  
  assignees.forEach((name, i) => {
    const r = headerRow + 1 + i;
    sheet.getRange("A" + r + ":B" + r).merge().setValue(name);
    
    sheet.getRange("C" + r).setHorizontalAlignment("center");
    setF(sheet, "C" + r, '=COUNTIF(\'Task List\'!H3:H500;"*' + name + '*")');
    
    sheet.getRange("D" + r).setHorizontalAlignment("center");
    setF(sheet, "D" + r, '=COUNTIFS(\'Task List\'!H3:H500;"*' + name + '*";\'Task List\'!G3:G500;"Finished")+COUNTIFS(\'Task List\'!H3:H500;"*' + name + '*";\'Task List\'!G3:G500;"Closed")');
    
    sheet.getRange("E" + r).setHorizontalAlignment("center");
    setF(sheet, "E" + r, '=COUNTIFS(\'Task List\'!H3:H500;"*' + name + '*";\'Task List\'!G3:G500;"In Progress")');
    
    sheet.getRange("F" + r).setHorizontalAlignment("center");
    setF(sheet, "F" + r, '=COUNTIFS(\'Task List\'!H3:H500;"*' + name + '*";\'Task List\'!G3:G500;"Testing")');
    
    sheet.getRange("G" + r).setHorizontalAlignment("center");
    setF(sheet, "G" + r, '=COUNTIFS(\'Task List\'!H3:H500;"*' + name + '*";\'Task List\'!G3:G500;"Open")+COUNTIFS(\'Task List\'!H3:H500;"*' + name + '*";\'Task List\'!G3:G500;"Pending")');
    
    sheet.getRange("H" + r).setHorizontalAlignment("center");
    setF(sheet, "H" + r, '=COUNTIFS(\'Task List\'!H3:H500;"*' + name + '*";\'Task List\'!F3:F500;"Urgent";\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed")');
    
    sheet.getRange("I" + r).setHorizontalAlignment("center");
    setF(sheet, "I" + r, '=COUNTIFS(\'Task List\'!H3:H500;"*' + name + '*";\'Task List\'!F3:F500;"High";\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed")');
    
    sheet.getRange("J" + r).setHorizontalAlignment("center");
    setF(sheet, "J" + r, '=COUNTIFS(\'Task List\'!H3:H500;"*' + name + '*";\'Task List\'!F3:F500;"Normal";\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed")');
    
    sheet.getRange("K" + r).setHorizontalAlignment("center");
    setF(sheet, "K" + r, '=COUNTIFS(\'Task List\'!H3:H500;"*' + name + '*";\'Task List\'!F3:F500;"Low";\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed")');
    
    sheet.getRange("L" + r).setHorizontalAlignment("center").setNumberFormat("0%");
    setF(sheet, "L" + r, '=IFERROR(D' + r + '/C' + r + ';0)');
    
    sheet.getRange("M" + r + ":T" + r).merge();
    setF(sheet, "M" + r, '=IFERROR(IF(E' + r + '>0;TEXTJOIN("; ";TRUE;FILTER(\'Task List\'!A3:A500&"-"&\'Task List\'!B3:B500;(ISNUMBER(SEARCH("' + name + '";\'Task List\'!H3:H500)))*(\'Task List\'!G3:G500="In Progress")));"Không có");"Không có")');
  });
  
  const detailEndRow = headerRow + assignees.length;
  
  // Total row
  const totalRow = detailEndRow + 1;
  sheet.getRange("A" + totalRow + ":B" + totalRow).merge().setValue("TỔNG").setFontWeight("bold");
  for (let col = 3; col <= 11; col++) {
    const letter = String.fromCharCode(64 + col);
    sheet.getRange(totalRow, col).setHorizontalAlignment("center").setFontWeight("bold");
    setF(sheet, letter + totalRow, '=SUM(' + letter + (headerRow + 1) + ':' + letter + detailEndRow + ')');
  }
  
  sheet.getRange("A" + headerRow + ":T" + totalRow).setBorder(true, true, true, true, true, true);
  
  // Conditional formatting
  const urgentRange = sheet.getRange("H" + (headerRow + 1) + ":H" + detailEndRow);
  const highRange = sheet.getRange("I" + (headerRow + 1) + ":I" + detailEndRow);
  const doneRange = sheet.getRange("L" + (headerRow + 1) + ":L" + detailEndRow);
  
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
      .setRanges([doneRange]).build()
  ]);
  
  // ============================================================
  // FREEZE & CHARTS
  // ============================================================
  
  sheet.setFrozenRows(2);
  
  try {
    // Pie chart Status
    sheet.insertChart(sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sheet.getRange("A11:D16"))
      .setPosition(19, 9, 0, 0)
      .setOption('title', 'Phân bổ theo Status')
      .setOption('pieHole', 0.4)
      .setOption('width', 350)
      .setOption('height', 230)
      .build());
      
  } catch(e) {}
  
  SpreadsheetApp.getUi().alert('✅ Tạo Overview thành công!\n\n📊 Giao diện đẹp với merge cells\n🔄 Dữ liệu realtime\n⚠️ Validation chuẩn');
}

function setF(sheet, cell, formula) {
  sheet.getRange(cell).setFormula(formula);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📊 Task Overview')
    .addItem('🔄 Tạo/Cập nhật Dashboard', 'createOverviewSheet')
    .addToUi();
}
