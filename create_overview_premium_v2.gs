/**
 * 📊 TASK OVERVIEW - PREMIUM V2
 * - Giữ nguyên giao diện đẹp
 * - Độ rộng cột đều nhau
 * - Merge cells cho các bảng độc lập
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
  const COL_WIDTH = 70;
  for (let i = 1; i <= 26; i++) {
    sheet.setColumnWidth(i, COL_WIDTH);
  }
  
  // ============================================================
  // SECTION 1: HEADER (Merge A1:Z1)
  // ============================================================
  
  sheet.getRange("A1:Z1").merge().setValue("📊 TASK OVERVIEW DASHBOARD")
    .setBackground("#0d47a1").setFontColor("white")
    .setFontSize(20).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.setRowHeight(1, 45);
  
  sheet.getRange("A2:Z2").merge().setValue("Timeline 2601 - Realtime Statistics")
    .setBackground("#1565c0").setFontColor("#bbdefb")
    .setFontSize(11).setHorizontalAlignment("center");
  
  // ============================================================
  // SECTION 2: KPI CARDS (Row 4-5) - Mỗi card 3 cột
  // Merge TỪNG HÀNG RIÊNG: hàng 4 = title, hàng 5 = value
  // ============================================================
  
  sheet.setRowHeight(4, 28);
  sheet.setRowHeight(5, 45);
  
  // Card 1: Tổng Task (A4:C4 title, A5:C5 value)
  sheet.getRange("A4:C4").merge().setValue("📋 TỔNG TASK")
    .setBackground("#e3f2fd").setFontSize(10).setFontColor("#616161").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange("A5:C5").merge().setBackground("#e3f2fd");
  setF(sheet, "A5", '=COUNTIF(\'Task List\'!G3:G500;"<>")');
  sheet.getRange("A5").setFontSize(26).setFontWeight("bold").setFontColor("#1565c0").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange("A4:C5").setBorder(true, true, true, true, null, null, "#bdbdbd", SpreadsheetApp.BorderStyle.SOLID);
  
  // Card 2: Hoàn thành (D4:F4 title, D5:F5 value)
  sheet.getRange("D4:F4").merge().setValue("✅ HOÀN THÀNH")
    .setBackground("#e8f5e9").setFontSize(10).setFontColor("#616161").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange("D5:F5").merge().setBackground("#e8f5e9");
  setF(sheet, "D5", '=COUNTIF(\'Task List\'!G3:G500;"Finished")+COUNTIF(\'Task List\'!G3:G500;"Closed")');
  sheet.getRange("D5").setFontSize(26).setFontWeight("bold").setFontColor("#2e7d32").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange("D4:F5").setBorder(true, true, true, true, null, null, "#bdbdbd", SpreadsheetApp.BorderStyle.SOLID);
  
  // Card 3: Đang làm (G4:I4 title, G5:I5 value)
  sheet.getRange("G4:I4").merge().setValue("🔄 ĐANG LÀM")
    .setBackground("#fff3e0").setFontSize(10).setFontColor("#616161").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange("G5:I5").merge().setBackground("#fff3e0");
  setF(sheet, "G5", '=COUNTIF(\'Task List\'!G3:G500;"In Progress")');
  sheet.getRange("G5").setFontSize(26).setFontWeight("bold").setFontColor("#ef6c00").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange("G4:I5").setBorder(true, true, true, true, null, null, "#bdbdbd", SpreadsheetApp.BorderStyle.SOLID);
  
  // Card 4: Testing (J4:L4 title, J5:L5 value)
  sheet.getRange("J4:L4").merge().setValue("🧪 TESTING")
    .setBackground("#f3e5f5").setFontSize(10).setFontColor("#616161").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange("J5:L5").merge().setBackground("#f3e5f5");
  setF(sheet, "J5", '=COUNTIF(\'Task List\'!G3:G500;"Testing")');
  sheet.getRange("J5").setFontSize(26).setFontWeight("bold").setFontColor("#7b1fa2").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange("J4:L5").setBorder(true, true, true, true, null, null, "#bdbdbd", SpreadsheetApp.BorderStyle.SOLID);
  
  // Card 5: Chờ xử lý (M4:O4 title, M5:O5 value)
  sheet.getRange("M4:O4").merge().setValue("⏳ CHỜ XỬ LÝ")
    .setBackground("#fce4ec").setFontSize(10).setFontColor("#616161").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange("M5:O5").merge().setBackground("#fce4ec");
  setF(sheet, "M5", '=COUNTIF(\'Task List\'!G3:G500;"Open")+COUNTIF(\'Task List\'!G3:G500;"Pending")');
  sheet.getRange("M5").setFontSize(26).setFontWeight("bold").setFontColor("#c2185b").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange("M4:O5").setBorder(true, true, true, true, null, null, "#bdbdbd", SpreadsheetApp.BorderStyle.SOLID);
  
  // Card 6: Urgent (P4:R4 title, P5:R5 value)
  sheet.getRange("P4:R4").merge().setValue("🚨 URGENT")
    .setBackground("#ffebee").setFontSize(10).setFontColor("#616161").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange("P5:R5").merge().setBackground("#ffebee");
  setF(sheet, "P5", '=COUNTIFS(\'Task List\'!F3:F500;"Urgent";\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed")');
  sheet.getRange("P5").setFontSize(26).setFontWeight("bold").setFontColor("#c62828").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange("P4:R5").setBorder(true, true, true, true, null, null, "#bdbdbd", SpreadsheetApp.BorderStyle.SOLID);
  
  // Card 7: % Hoàn thành (S4:U4 title, S5:U5 value)
  sheet.getRange("S4:U4").merge().setValue("📈 TIẾN ĐỘ")
    .setBackground("#e0f7fa").setFontSize(10).setFontColor("#616161").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange("S5:U5").merge().setBackground("#e0f7fa");
  setF(sheet, "S5", '=IFERROR(D5/A5;0)');
  sheet.getRange("S5").setFontSize(26).setFontWeight("bold").setFontColor("#00838f").setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0%");
  sheet.getRange("S4:U5").setBorder(true, true, true, true, null, null, "#bdbdbd", SpreadsheetApp.BorderStyle.SOLID);
  
  // ============================================================
  // SECTION 3: ALERT BOX (Row 7-8)
  // ============================================================
  
  sheet.getRange("A7:U7").merge().setValue("⚠️ CẢNH BÁO")
    .setBackground("#ff8a65").setFontColor("white")
    .setFontWeight("bold").setHorizontalAlignment("center");
  
  sheet.getRange("A8:G8").merge();
  setF(sheet, "A8", '=IF(P5>0;"🔴 "&P5&" task URGENT cần xử lý!";"✅ Không có Urgent")');
  sheet.getRange("A8").setFontColor("#c62828").setFontWeight("bold").setHorizontalAlignment("center");
  
  sheet.getRange("H8:N8").merge();
  setF(sheet, "H8", '=IFERROR(IF(COUNTIFS(\'Task List\'!D3:D500;"<"&TODAY();\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed";\'Task List\'!D3:D500;"<>")>0;"⏰ "&COUNTIFS(\'Task List\'!D3:D500;"<"&TODAY();\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed";\'Task List\'!D3:D500;"<>")&" task QUÁ HẠN!";"✅ Không có quá hạn");"✅ OK")');
  sheet.getRange("H8").setFontColor("#d84315").setFontWeight("bold").setHorizontalAlignment("center");
  
  sheet.getRange("O8:U8").merge();
  setF(sheet, "O8", '=IFERROR(IF(COUNTIFS(\'Task List\'!D3:D500;">="&TODAY();\'Task List\'!D3:D500;"<="&TODAY()+3;\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed";\'Task List\'!D3:D500;"<>")>0;"📅 "&COUNTIFS(\'Task List\'!D3:D500;">="&TODAY();\'Task List\'!D3:D500;"<="&TODAY()+3;\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed";\'Task List\'!D3:D500;"<>")&" task sắp hết hạn";"✅ OK");"✅ OK")');
  sheet.getRange("O8").setFontColor("#f57c00").setFontWeight("bold").setHorizontalAlignment("center");
  
  sheet.getRange("A7:U8").setBorder(true, true, true, true, true, true);
  
  // ============================================================
  // SECTION 4: BẢNG STATUS (A10:G18) - 7 cột merged
  // ============================================================
  
  sheet.getRange("A10:G10").merge().setValue("📈 THỐNG KÊ THEO TRẠNG THÁI")
    .setBackground("#43a047").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  
  // Header
  sheet.getRange("A11:B11").merge().setValue("Trạng thái").setFontWeight("bold").setBackground("#c8e6c9").setHorizontalAlignment("center");
  sheet.getRange("C11").setValue("SL").setFontWeight("bold").setBackground("#c8e6c9").setHorizontalAlignment("center");
  sheet.getRange("D11").setValue("%").setFontWeight("bold").setBackground("#c8e6c9").setHorizontalAlignment("center");
  sheet.getRange("E11:G11").merge().setValue("Tiến độ").setFontWeight("bold").setBackground("#c8e6c9").setHorizontalAlignment("center");
  
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
    sheet.getRange("A" + r + ":B" + r).merge().setValue(s.icon + " " + s.name).setBackground(s.color);
    sheet.getRange("C" + r).setBackground(s.color).setHorizontalAlignment("center");
    setF(sheet, "C" + r, '=COUNTIF(\'Task List\'!G3:G500;"' + s.name + '")');
    sheet.getRange("D" + r).setBackground(s.color).setHorizontalAlignment("center").setNumberFormat("0%");
    setF(sheet, "D" + r, '=IFERROR(C' + r + '/$C$18;0)');
    sheet.getRange("E" + r + ":G" + r).merge().setBackground(s.color).setFontSize(9);
    setF(sheet, "E" + r, '=REPT("▓";ROUND(D' + r + '*10))&REPT("░";10-ROUND(D' + r + '*10))');
  });
  
  // Total
  sheet.getRange("A18:B18").merge().setValue("TỔNG").setFontWeight("bold");
  sheet.getRange("C18").setFontWeight("bold").setHorizontalAlignment("center");
  setF(sheet, "C18", '=SUM(C12:C17)');
  sheet.getRange("D18").setValue("100%").setFontWeight("bold").setHorizontalAlignment("center");
  
  sheet.getRange("A10:G18").setBorder(true, true, true, true, true, true);
  
  // ============================================================
  // SECTION 5: BẢNG PRIORITY (I10:O16) - 7 cột merged
  // ============================================================
  
  sheet.getRange("I10:O10").merge().setValue("🎯 THỐNG KÊ THEO ĐỘ ƯU TIÊN")
    .setBackground("#e53935").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  
  // Header
  sheet.getRange("I11:K11").merge().setValue("Độ ưu tiên").setFontWeight("bold").setBackground("#ffcdd2").setHorizontalAlignment("center");
  sheet.getRange("L11").setValue("Tổng").setFontWeight("bold").setBackground("#ffcdd2").setHorizontalAlignment("center");
  sheet.getRange("M11").setValue("Chưa").setFontWeight("bold").setBackground("#ffcdd2").setHorizontalAlignment("center");
  sheet.getRange("N11:O11").merge().setValue("⚠️ Cảnh báo").setFontWeight("bold").setBackground("#ffcdd2").setHorizontalAlignment("center");
  
  const priorities = [
    {name: "Urgent", icon: "🔴", color: "#ffcdd2"},
    {name: "High", icon: "🟠", color: "#ffe0b2"},
    {name: "Normal", icon: "🟡", color: "#fff9c4"},
    {name: "Low", icon: "🟢", color: "#c8e6c9"}
  ];
  
  priorities.forEach((p, i) => {
    const r = 12 + i;
    sheet.getRange("I" + r + ":K" + r).merge().setValue(p.icon + " " + p.name).setBackground(p.color);
    sheet.getRange("L" + r).setBackground(p.color).setHorizontalAlignment("center");
    setF(sheet, "L" + r, '=COUNTIF(\'Task List\'!F3:F500;"' + p.name + '")');
    sheet.getRange("M" + r).setBackground(p.color).setHorizontalAlignment("center");
    setF(sheet, "M" + r, '=COUNTIFS(\'Task List\'!F3:F500;"' + p.name + '";\'Task List\'!G3:G500;"<>Finished";\'Task List\'!G3:G500;"<>Closed")');
    sheet.getRange("N" + r + ":O" + r).merge().setBackground(p.color).setHorizontalAlignment("center");
    setF(sheet, "N" + r, '=IF(M' + r + '>0;"⚠️ "&M' + r + '&" task";"✅")');
  });
  
  // Total
  sheet.getRange("I16:K16").merge().setValue("TỔNG").setFontWeight("bold");
  sheet.getRange("L16").setFontWeight("bold").setHorizontalAlignment("center");
  setF(sheet, "L16", '=SUM(L12:L15)');
  sheet.getRange("M16").setFontWeight("bold").setHorizontalAlignment("center");
  setF(sheet, "M16", '=SUM(M12:M15)');
  
  sheet.getRange("I10:O16").setBorder(true, true, true, true, true, true);
  
  // ============================================================
  // SECTION 6: BẢNG ASSIGNEE (Q10:U+) - 5 cột merged
  // ============================================================
  
  sheet.getRange("Q10:U10").merge().setValue("👥 THỐNG KÊ THEO NGƯỜI")
    .setBackground("#7b1fa2").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  
  // Header
  sheet.getRange("Q11:R11").merge().setValue("Assignee").setFontWeight("bold").setBackground("#e1bee7").setHorizontalAlignment("center");
  sheet.getRange("S11").setValue("Task").setFontWeight("bold").setBackground("#e1bee7").setHorizontalAlignment("center");
  sheet.getRange("T11").setValue("Done").setFontWeight("bold").setBackground("#e1bee7").setHorizontalAlignment("center");
  sheet.getRange("U11").setValue("Workload").setFontWeight("bold").setBackground("#e1bee7").setHorizontalAlignment("center");
  
  assignees.forEach((name, i) => {
    const r = 12 + i;
    sheet.getRange("Q" + r + ":R" + r).merge().setValue(name);
    sheet.getRange("S" + r).setHorizontalAlignment("center");
    setF(sheet, "S" + r, '=COUNTIF(\'Task List\'!H3:H500;"*' + name + '*")');
    sheet.getRange("T" + r).setHorizontalAlignment("center");
    setF(sheet, "T" + r, '=COUNTIFS(\'Task List\'!H3:H500;"*' + name + '*";\'Task List\'!G3:G500;"Finished")+COUNTIFS(\'Task List\'!H3:H500;"*' + name + '*";\'Task List\'!G3:G500;"Closed")');
    setF(sheet, "U" + r, '=REPT("█";S' + r + ')');
    sheet.getRange("U" + r).setFontColor("#7b1fa2").setFontSize(8);
  });
  
  const assEndRow = 11 + assignees.length;
  sheet.getRange("Q10:U" + assEndRow).setBorder(true, true, true, true, true, true);
  
  // ============================================================
  // SECTION 7: TOP PERFORMERS (A20:G26)
  // ============================================================
  
  sheet.getRange("A20:G20").merge().setValue("🏆 TOP PERFORMERS")
    .setBackground("#ff6f00").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  
  // Header
  sheet.getRange("A21").setValue("#").setFontWeight("bold").setBackground("#ffe0b2").setHorizontalAlignment("center");
  sheet.getRange("B21:D21").merge().setValue("Assignee").setFontWeight("bold").setBackground("#ffe0b2").setHorizontalAlignment("center");
  sheet.getRange("E21").setValue("Done").setFontWeight("bold").setBackground("#ffe0b2").setHorizontalAlignment("center");
  sheet.getRange("F21").setValue("%").setFontWeight("bold").setBackground("#ffe0b2").setHorizontalAlignment("center");
  sheet.getRange("G21").setValue("🏅").setFontWeight("bold").setBackground("#ffe0b2").setHorizontalAlignment("center");
  
  for (let i = 1; i <= 5; i++) {
    const r = 21 + i;
    sheet.getRange("A" + r).setValue(i).setHorizontalAlignment("center");
    sheet.getRange("B" + r + ":D" + r).merge();
    setF(sheet, "B" + r, '=IFERROR(IF(LARGE($T$12:$T$' + assEndRow + ';' + i + ')>0;INDEX($Q$12:$Q$' + assEndRow + ';MATCH(LARGE($T$12:$T$' + assEndRow + ';' + i + ');$T$12:$T$' + assEndRow + ';0));"");"")');
    sheet.getRange("E" + r).setHorizontalAlignment("center");
    setF(sheet, "E" + r, '=IFERROR(IF(LARGE($T$12:$T$' + assEndRow + ';' + i + ')>0;LARGE($T$12:$T$' + assEndRow + ';' + i + ');"");"")');
    sheet.getRange("F" + r).setHorizontalAlignment("center").setNumberFormat("0%");
    setF(sheet, "F" + r, '=IFERROR(IF(B' + r + '<>"";INDEX($T$12:$T$' + assEndRow + ';MATCH(B' + r + ';$Q$12:$Q$' + assEndRow + ';0))/INDEX($S$12:$S$' + assEndRow + ';MATCH(B' + r + ';$Q$12:$Q$' + assEndRow + ';0));"");"")');
    sheet.getRange("G" + r).setHorizontalAlignment("center");
    setF(sheet, "G" + r, '=IF(E' + r + '<>"";IF(' + i + '=1;"🥇";IF(' + i + '=2;"🥈";IF(' + i + '=3;"🥉";"")));"")');
  }
  
  sheet.getRange("A20:G26").setBorder(true, true, true, true, true, true);
  
  // ============================================================
  // SECTION 8: TASK SẮP HẾT HẠN (I20:O28)
  // ============================================================
  
  sheet.getRange("I20:O20").merge().setValue("⏰ TASK SẮP HẾT HẠN (3 ngày)")
    .setBackground("#d84315").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  
  sheet.getRange("I21:L21").merge().setValue("Task").setFontWeight("bold").setBackground("#ffccbc").setHorizontalAlignment("center");
  sheet.getRange("M21:O21").merge().setValue("End Date").setFontWeight("bold").setBackground("#ffccbc").setHorizontalAlignment("center");
  
  sheet.getRange("I22:L28").merge();
  setF(sheet, "I22", '=IFERROR(TEXTJOIN(CHAR(10);TRUE;FILTER(\'Task List\'!A3:A500&" - "&\'Task List\'!B3:B500;(\'Task List\'!D3:D500>=TODAY())*(\'Task List\'!D3:D500<=TODAY()+3)*(\'Task List\'!G3:G500<>"Finished")*(\'Task List\'!G3:G500<>"Closed")*(\'Task List\'!D3:D500<>"")));"Không có task sắp hết hạn ✅")');
  sheet.getRange("I22").setWrap(true).setVerticalAlignment("top").setFontSize(9);
  
  sheet.getRange("M22:O28").merge();
  setF(sheet, "M22", '=IFERROR(TEXTJOIN(CHAR(10);TRUE;FILTER(TEXT(\'Task List\'!D3:D500;"dd/mm");(\'Task List\'!D3:D500>=TODAY())*(\'Task List\'!D3:D500<=TODAY()+3)*(\'Task List\'!G3:G500<>"Finished")*(\'Task List\'!G3:G500<>"Closed")*(\'Task List\'!D3:D500<>"")));"")');
  sheet.getRange("M22").setWrap(true).setVerticalAlignment("top").setFontSize(9).setHorizontalAlignment("center");
  
  sheet.getRange("I20:O28").setBorder(true, true, true, true, true, true);
  
  // ============================================================
  // SECTION 9: TASK QUÁ HẠN (Q20:U28)
  // ============================================================
  
  sheet.getRange("Q20:U20").merge().setValue("🔥 TASK QUÁ HẠN")
    .setBackground("#b71c1c").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  
  sheet.getRange("Q21:S21").merge().setValue("Task").setFontWeight("bold").setBackground("#ffcdd2").setHorizontalAlignment("center");
  sheet.getRange("T21:U21").merge().setValue("Quá").setFontWeight("bold").setBackground("#ffcdd2").setHorizontalAlignment("center");
  
  sheet.getRange("Q22:S28").merge();
  setF(sheet, "Q22", '=IFERROR(TEXTJOIN(CHAR(10);TRUE;FILTER(\'Task List\'!A3:A500&" - "&\'Task List\'!B3:B500;(\'Task List\'!D3:D500<TODAY())*(\'Task List\'!D3:D500<>"")*(\'Task List\'!G3:G500<>"Finished")*(\'Task List\'!G3:G500<>"Closed")));"Không có task quá hạn ✅")');
  sheet.getRange("Q22").setWrap(true).setVerticalAlignment("top").setFontSize(9);
  
  sheet.getRange("T22:U28").merge();
  setF(sheet, "T22", '=IFERROR(TEXTJOIN(CHAR(10);TRUE;FILTER(TODAY()-\'Task List\'!D3:D500&" ngày";(\'Task List\'!D3:D500<TODAY())*(\'Task List\'!D3:D500<>"")*(\'Task List\'!G3:G500<>"Finished")*(\'Task List\'!G3:G500<>"Closed")));"")');
  sheet.getRange("T22").setWrap(true).setVerticalAlignment("top").setFontSize(9).setHorizontalAlignment("center").setFontColor("#c62828");
  
  sheet.getRange("Q20:U28").setBorder(true, true, true, true, true, true);
  
  // ============================================================
  // SECTION 10: CHI TIẾT WORKLOAD (A30:U+)
  // ============================================================
  
  const detailStartRow = 30;
  
  sheet.getRange("A" + detailStartRow + ":U" + detailStartRow).merge()
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
  sheet.getRange("H" + headerRow).setValue("🔴Urgent").setFontWeight("bold").setBackground("#ffcdd2").setHorizontalAlignment("center");
  sheet.getRange("I" + headerRow).setValue("🟠High").setFontWeight("bold").setBackground("#ffe0b2").setHorizontalAlignment("center");
  sheet.getRange("J" + headerRow).setValue("🟡Normal").setFontWeight("bold").setBackground("#fff9c4").setHorizontalAlignment("center");
  sheet.getRange("K" + headerRow).setValue("🟢Low").setFontWeight("bold").setBackground("#c8e6c9").setHorizontalAlignment("center");
  sheet.getRange("L" + headerRow).setValue("Done%").setFontWeight("bold").setBackground("#bbdefb").setHorizontalAlignment("center");
  sheet.getRange("M" + headerRow + ":U" + headerRow).merge().setValue("📝 Task đang làm (In Progress)").setFontWeight("bold").setBackground("#bbdefb").setHorizontalAlignment("center");
  
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
    
    sheet.getRange("M" + r + ":U" + r).merge().setFontSize(9);
    setF(sheet, "M" + r, '=IFERROR(IF(E' + r + '>0;TEXTJOIN("; ";TRUE;FILTER(\'Task List\'!A3:A500&"-"&\'Task List\'!B3:B500;(ISNUMBER(SEARCH("' + name + '";\'Task List\'!H3:H500)))*(\'Task List\'!G3:G500="In Progress")));"—");"—")');
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
  
  sheet.getRange("A" + headerRow + ":U" + totalRow).setBorder(true, true, true, true, true, true);
  
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
      .setRanges([doneRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0.3)
      .setBackground("#ffcdd2").setFontColor("#c62828")
      .setRanges([doneRange]).build()
  ]);
  
  // ============================================================
  // FREEZE ROWS
  // ============================================================
  
  sheet.setFrozenRows(2);
  
  // ============================================================
  // CHARTS (sau các bảng)
  // ============================================================
  
  try {
    // Pie chart Status - đặt ở cột W (sau bảng cuối)
    sheet.insertChart(sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sheet.getRange("A12:C17"))
      .setPosition(10, 23, 0, 0)  // Row 10, Column W
      .setOption('title', 'Phân bổ theo Status')
      .setOption('pieHole', 0.4)
      .setOption('width', 320)
      .setOption('height', 220)
      .setOption('colors', ['#4caf50', '#ffeb3b', '#2196f3', '#9c27b0', '#8bc34a', '#607d8b'])
      .build());
    
    // Bar chart Assignee Workload
    sheet.insertChart(sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(sheet.getRange("Q12:S" + assEndRow))
      .setPosition(20, 23, 0, 0)  // Row 20, Column W
      .setOption('title', 'Workload theo Assignee')
      .setOption('width', 320)
      .setOption('height', 250)
      .setOption('legend', {position: 'none'})
      .setOption('colors', ['#7b1fa2'])
      .build());
      
    // Column chart Priority
    sheet.insertChart(sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sheet.getRange("I12:M15"))
      .setPosition(detailStartRow, 23, 0, 0)
      .setOption('title', 'Priority: Tổng vs Chưa xong')
      .setOption('width', 320)
      .setOption('height', 220)
      .setOption('colors', ['#9e9e9e', '#f44336'])
      .build());
      
  } catch(e) {
    Logger.log("Chart error: " + e);
  }
  
  SpreadsheetApp.getUi().alert('✅ Tạo Overview Premium V2 thành công!\n\n📊 Giao diện đẹp + Merge cells\n🔄 Dữ liệu realtime\n📈 3 biểu đồ\n⚠️ Cảnh báo tự động\n🏆 Top Performers');
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
