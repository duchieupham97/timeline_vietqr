/**
 * DEBUG SCRIPT - Kiểm tra cấu trúc sheet
 * Chạy function này để xem thông tin sheet của bạn
 */

function debugSheetInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Liệt kê tất cả các sheet
  const sheets = ss.getSheets();
  let info = "📋 DANH SÁCH SHEETS:\n";
  sheets.forEach((s, i) => {
    info += `${i + 1}. "${s.getName()}"\n`;
  });
  
  // Tìm sheet Task List
  const taskSheet = ss.getSheetByName("Task List");
  if (taskSheet) {
    info += "\n✅ Tìm thấy sheet 'Task List'\n";
    
    // Đọc header row
    const headerRow1 = taskSheet.getRange("A1:K1").getValues()[0];
    const headerRow2 = taskSheet.getRange("A2:K2").getValues()[0];
    
    info += "\n📌 NỘI DUNG HÀNG 1:\n";
    headerRow1.forEach((cell, i) => {
      const col = String.fromCharCode(65 + i);
      info += `${col}: "${cell}"\n`;
    });
    
    info += "\n📌 NỘI DUNG HÀNG 2:\n";
    headerRow2.forEach((cell, i) => {
      const col = String.fromCharCode(65 + i);
      info += `${col}: "${cell}"\n`;
    });
    
    // Đọc vài dòng dữ liệu
    info += "\n📌 MẪU DỮ LIỆU (hàng 3-5):\n";
    const sampleData = taskSheet.getRange("A3:K5").getValues();
    sampleData.forEach((row, rowIdx) => {
      info += `Hàng ${rowIdx + 3}: `;
      row.forEach((cell, colIdx) => {
        if (cell !== "") {
          const col = String.fromCharCode(65 + colIdx);
          info += `${col}="${cell}" | `;
        }
      });
      info += "\n";
    });
    
    // Đếm số hàng có dữ liệu
    const lastRow = taskSheet.getLastRow();
    info += `\n📊 Số hàng có dữ liệu: ${lastRow}`;
    
  } else {
    info += "\n❌ KHÔNG tìm thấy sheet 'Task List'";
    info += "\nHãy kiểm tra lại tên sheet chính xác!";
  }
  
  // Hiển thị kết quả
  const ui = SpreadsheetApp.getUi();
  ui.alert("DEBUG INFO", info, ui.ButtonSet.OK);
  
  // Cũng log ra console
  Logger.log(info);
}

/**
 * Test một công thức đơn giản
 */
function testSimpleFormula() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const taskSheet = ss.getSheetByName("Task List");
  
  if (!taskSheet) {
    SpreadsheetApp.getUi().alert("Không tìm thấy sheet 'Task List'");
    return;
  }
  
  // Test đếm số dòng có dữ liệu ở cột A
  const colA = taskSheet.getRange("A:A").getValues();
  let count = 0;
  colA.forEach(row => {
    if (row[0] && String(row[0]).startsWith("F")) count++;
  });
  
  // Test đếm Status
  const colG = taskSheet.getRange("G:G").getValues();
  let statusCount = {};
  colG.forEach(row => {
    const val = row[0];
    if (val && val !== "" && val !== "Status") {
      statusCount[val] = (statusCount[val] || 0) + 1;
    }
  });
  
  let info = `📊 KẾT QUẢ TEST:\n\n`;
  info += `Số task có FNo. bắt đầu bằng "F": ${count}\n\n`;
  info += `Thống kê Status:\n`;
  for (const [status, cnt] of Object.entries(statusCount)) {
    info += `- ${status}: ${cnt}\n`;
  }
  
  SpreadsheetApp.getUi().alert("TEST RESULT", info, SpreadsheetApp.getUi().ButtonSet.OK);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🔧 Debug')
    .addItem('📋 Xem thông tin Sheet', 'debugSheetInfo')
    .addItem('🧪 Test đếm dữ liệu', 'testSimpleFormula')
    .addToUi();
}
