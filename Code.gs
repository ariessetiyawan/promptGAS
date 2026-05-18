const SAact=SpreadsheetApp.getActiveSpreadsheet()
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('GAS Multi-Phase Prompt Generator')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Mengambil data fitur dari Kolom A di Spreadsheet
function getFeatures() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  if (data.length === 0 || (data.length === 1 && data[0][0] === "")) {
    const defaultFeatures = [
      "Operasi CRUD (Create, Read, Update, Delete) lengkap",
      "Koneksi Database ke Google Spreadsheet",
      "UI Dashboard Command Center bergaya ekstrim fungsional",
      "Sistem Autentikasi Login (Session/LocalStorage)",
      "Optimasi Kecepatan dengan CacheService",
      "Login/ Auth System"
    ];
    
    const sheetData = defaultFeatures.map(f => [f]);
    sheet.getRange(1, 2, sheetData.length, 1).setValues(sheetData);
    return defaultFeatures;
  }
  
  return data.map(row => row[0]).filter(val => val !== "");
}

// Menyimpan fitur baru
function saveFeature(newFeature) {
  const sheet = SAact.getActiveSheet();
  sheet.appendRow([newFeature]);
  return getFeatures(); 
}

// Membuat Sheet Database berdasarkan skema yang dirancang di UI
function initAppDatabase(sheetName, schemaData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear(); 
  }

  if (schemaData && schemaData.length > 0) {
    const headers = schemaData.map(col => col.name);
    const types = schemaData.map(col => col.type);

    // Header Baris 1
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight("bold").setBackground("#0ea5e9").setFontColor("white");

    // Tipe Data Baris 2
    const typeRange = sheet.getRange(2, 1, 1, types.length);
    typeRange.setValues([types]);
    typeRange.setFontStyle("italic").setFontColor("#64748b"); 

    sheet.autoResizeColumns(1, headers.length);
  }

  return `Database '${sheetName}' berhasil dibuat dan siap digunakan!`;
}
