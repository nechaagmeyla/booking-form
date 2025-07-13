// === APPS SCRIPT: Code.gs ===
// Hilangkan pembatasan email agar bisa menerima dari siapa saja
// function isAuthorized() {
//   const email = Session.getActiveUser().getEmail();
//   return email === "agmeylanecha@gmail.com";
// }

function doGet(e) {
  const action = e.parameter.action;
  if (action === "get_mrs_view") return doGetMRSView(e);
  
  const tanggal = e.parameter.tanggal_penerimaan;
  
  if (action === "get_mrs") {
    return doGetMRS(e);
  }
  
  if (!tanggal) return ContentService.createTextOutput("❌ Tanggal kosong");

  const date = new Date(tanggal);
  const folderName = `Booking_${date.getFullYear()}_${String(date.getMonth() + 1).padStart(2, '0')}`;
  const parentId = "1ypv_G7bkmc2BOZAHn6sbD50t9k4aJou1";
  const parent = DriveApp.getFolderById(parentId);
  const folders = parent.getFoldersByName(folderName);
  if (!folders.hasNext()) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);

  const folder = folders.next();
  const files = folder.getFilesByName(folderName);
  if (!files.hasNext()) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);

  const ss = SpreadsheetApp.open(files.next());
  const sheet = ss.getSheetByName(tanggal);
  if (!sheet) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const result = [];

  for (let i = 1; i < data.length; i++) {
    const row = {};
    for (let j = 0; j < headers.length; j++) {
      let val = data[i][j];
      if (val instanceof Date) val = val.toISOString();
      row[headers[j].toLowerCase().replace(/ /g, "_")] = val;
    }
    row.sheetName = sheet.getName();
    row.id_row = data[i][0]; // Add id_row for edit/delete operations
    result.push(row);
  }

  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doGetMRSView(e) {
  var bulan = e.parameter.bulan; // Contoh: 'Mei'
  var tahun = e.parameter.tahun; // Contoh: '2025'
  if (!bulan || !tahun) {
    return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
  }

  // Nama file dan sheet
  var monthNum = getMonthNumber(bulan); // 'Mei' -> '05'
  var fileName = 'Booking_' + tahun + '_' + monthNum;
  var sheetName = 'MRS_' + bulan;

  // Cari file di Google Drive
  var files = DriveApp.getFilesByName(fileName);
  if (!files.hasNext()) {
    return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
  }
  var file = files.next();
  var ss = SpreadsheetApp.open(file);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) { // Mulai dari 1, skip header
    result.push({
      id_row: i + 1, // baris di sheet (1-based, +1 karena header)
      nomor_polisi: data[i][0],
      bulan: data[i][1],
      ch: data[i][2],
      nama_customer: data[i][3],
      no_hp: data[i][4]
    });
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// Helper: Ubah nama bulan ke angka 2 digit
function getMonthNumber(bulan) {
  var bulanMap = {
    'Januari': '01', 'Februari': '02', 'Maret': '03', 'April': '04',
    'Mei': '05', 'Juni': '06', 'Juli': '07', 'Agustus': '08',
    'September': '09', 'Oktober': '10', 'November': '11', 'Desember': '12'
  };
  return bulanMap[bulan] || '01';
}

function doGetMRS(e) {
  const month = e.parameter.month;
  
  try {
    const mrsSS = getMRSSpreadsheet();
    const sheet = mrsSS.getSheetByName("MRS_Data");
    if (!sheet) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const result = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = {};
      for (let j = 0; j < headers.length; j++) {
        let val = data[i][j];
        if (val instanceof Date) val = val.toISOString();
        row[headers[j].toLowerCase().replace(/ /g, "_")] = val;
      }
      row.id_row = data[i][0];
      result.push(row);
    }
    
    // Filter by month if specified
    if (month && month !== "") {
      const filteredResult = result.filter(row => row.reminder_bulan === month);
      return ContentService.createTextOutput(JSON.stringify(filteredResult)).setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  const p = e.parameter;
  
  // Handle different actions
  if (p.action === "delete") return doDelete(e);
  if (p.action === "edit") return doEdit(e);
  if (p.action === "export_rekap") return doExportRekap(e);
  if (p.action === "add_mrs") return doAddMRSV3(e);
  if (p.action === "delete_mrs") return doDeleteMRS(e);
  
  // Default action: add new booking
  const tglBooking = formatTanggalLengkap(new Date());
  const sheet = getOrCreateSheet(p.tanggal_penerimaan);
  const idRow = `${Date.now()}_${p.plat}`;
  const newRow = [
    idRow, "", p.status_kehadiran || "", p.plat, p.tipe, p.tahun, p.tipe_telepon || "", p.nohp, p.datang, p.nama,
    tglBooking, p.tanggal_penerimaan, p.jam_penerimaan, p.via, p.admin, p.keluhan, p.sales || ""
  ];
  insertDataSortedByJam(sheet, newRow);
  return ContentService.createTextOutput("✅ Data berhasil ditambahkan.");
}

function doEdit(e) {
  const p = e.parameter;
  const sheet = getSheetBySheetName(p.tanggal_penerimaan, p.sheetName);
  if (!sheet) return ContentService.createTextOutput("❌ Sheet tidak ditemukan");
  
  const data = sheet.getDataRange().getValues();
  const rowIdx = data.findIndex(row => row[0] === p.id_row);
  if (rowIdx === -1) return ContentService.createTextOutput("❌ Data tidak ditemukan");

  const tglBooking = formatTanggalLengkap(new Date());
  const updateRow = [
    p.id_row, "", p.status_kehadiran || "", p.plat, p.tipe, p.tahun, p.tipe_telepon || "", p.nohp, p.datang, p.nama,
    tglBooking, p.tanggal_penerimaan, p.jam_penerimaan, p.via, p.admin, p.keluhan, p.sales || ""
  ];

  sheet.deleteRow(rowIdx + 1);
  insertDataSortedByJam(sheet, updateRow);

  return ContentService.createTextOutput("✅ Data berhasil diupdate.");
}

function doDelete(e) {
  const { id_row, tanggal_penerimaan, sheetName } = e.parameter;
  const sheet = getSheetBySheetName(tanggal_penerimaan, sheetName);
  if (!sheet) return ContentService.createTextOutput("❌ Sheet tidak ditemukan");
  
  const data = sheet.getDataRange().getValues();
  const rowIdx = data.findIndex(row => row[0] === id_row);
  if (rowIdx === -1 || rowIdx < 1) return ContentService.createTextOutput("❌ Data tidak ditemukan atau tidak valid");
  
  sheet.deleteRow(rowIdx + 1);
  updateRowNumbers(sheet);
  return ContentService.createTextOutput("✅ Data berhasil dihapus");
}

function doExportRekap(e) {
  const { tanggal_penerimaan } = e.parameter;
  if (!tanggal_penerimaan) return ContentService.createTextOutput("❌ Tanggal kosong");

  try {
    const date = new Date(tanggal_penerimaan);
    const folderName = `Booking_${date.getFullYear()}_${String(date.getMonth() + 1).padStart(2, '0')}`;
    const parentId = "1ypv_G7bkmc2BOZAHn6sbD50t9k4aJou1";
    const parent = DriveApp.getFolderById(parentId);
    const folders = parent.getFoldersByName(folderName);
    
    if (!folders.hasNext()) return ContentService.createTextOutput("❌ Folder tidak ditemukan");

    const folder = folders.next();
    const files = folder.getFilesByName(folderName);
    if (!files.hasNext()) return ContentService.createTextOutput("❌ Spreadsheet tidak ditemukan");

    const ss = SpreadsheetApp.open(files.next());
    const sheet = ss.getSheetByName(tanggal_penerimaan);
    if (!sheet) return ContentService.createTextOutput("❌ Sheet tidak ditemukan");

    // Create rekap sheet
    const rekapSheetName = `Rekap_${tanggal_penerimaan}`;
    let rekapSheet = ss.getSheetByName(rekapSheetName);
    if (rekapSheet) {
      ss.deleteSheet(rekapSheet);
    }
    rekapSheet = ss.insertSheet(rekapSheetName);

    // Get data and create summary
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find status column index
    const statusColIdx = headers.findIndex(h => h.toLowerCase().includes('status'));
    
    // Count statistics
    let showCount = 0, noShowCount = 0, rescheduleCount = 0, unknownCount = 0;
    const filteredData = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = statusColIdx >= 0 ? (row[statusColIdx] || "") : "";
      
      if (status === "show") showCount++;
      else if (status === "no show") noShowCount++;
      else if (status === "reschedule") rescheduleCount++;
      else unknownCount++; // Untuk status kosong atau tidak diketahui
      
      // Add to filtered data for detailed table
      filteredData.push(row);
    }

    // Create summary section
    rekapSheet.getRange(1, 1).setValue("REKAP SHOW/NO SHOW/RESCHEDULE");
    rekapSheet.getRange(1, 1, 1, 1).setFontWeight("bold").setFontSize(14);
    
    rekapSheet.getRange(3, 1).setValue("Tanggal:");
    rekapSheet.getRange(3, 2).setValue(tanggal_penerimaan);
    
    rekapSheet.getRange(5, 1).setValue("Ringkasan:");
    rekapSheet.getRange(5, 1).setFontWeight("bold");
    
    rekapSheet.getRange(6, 1).setValue("Show:");
    rekapSheet.getRange(6, 2).setValue(showCount);
    
    rekapSheet.getRange(7, 1).setValue("No Show:");
    rekapSheet.getRange(7, 2).setValue(noShowCount);
    
    rekapSheet.getRange(8, 1).setValue("Reschedule:");
    rekapSheet.getRange(8, 2).setValue(rescheduleCount);
    
    rekapSheet.getRange(9, 1).setValue("Tidak Diketahui:");
    rekapSheet.getRange(9, 2).setValue(unknownCount);
    
    rekapSheet.getRange(10, 1).setValue("Total:");
    rekapSheet.getRange(10, 1).setFontWeight("bold");
    rekapSheet.getRange(10, 2).setValue(filteredData.length);
    rekapSheet.getRange(10, 2).setFontWeight("bold");

    // Create detailed table
    if (filteredData.length > 0) {
      rekapSheet.getRange(12, 1).setValue("Detail Data:");
      rekapSheet.getRange(12, 1).setFontWeight("bold");
      
      // Add headers for rekap table
      const rekapHeaders = ["No", "Status", "Nomor Polisi", "Nama Customer", "Tanggal Penerimaan", "Jam Penerimaan", "Via", "Admin", "Sales"];
      rekapSheet.getRange(13, 1, 1, rekapHeaders.length).setValues([rekapHeaders]);
      rekapSheet.getRange(13, 1, 1, rekapHeaders.length).setFontWeight("bold");
      
      // Add data rows
      for (let i = 0; i < filteredData.length; i++) {
        const row = filteredData[i];
        const status = statusColIdx >= 0 ? (row[statusColIdx] || "") : "";
        const displayStatus = status || "-"; // Tampilkan "-" jika kosong
        
        const rekapRow = [
          i + 1,
          displayStatus,
          row[3] || "", // Nomor Polisi
          row[8] || "", // Nama Customer
          row[10] || "", // Tanggal Penerimaan
          row[11] || "", // Jam Penerimaan
          row[12] || "", // Via
          row[13] || "", // Admin
          row[15] || ""  // Sales
        ];
        
        rekapSheet.getRange(14 + i, 1, 1, rekapRow.length).setValues([rekapRow]);
      }
    }

    // Auto-resize columns
    rekapSheet.autoResizeColumns(1, rekapHeaders.length);
    
    return ContentService.createTextOutput(`✅ Rekap berhasil dibuat di sheet "${rekapSheetName}"`);
    
  } catch (error) {
    return ContentService.createTextOutput(`❌ Error: ${error.toString()}`);
  }
}

function insertDataSortedByJam(sheet, newRow) {
  const allData = sheet.getDataRange().getValues();
  const dataRows = allData.slice(1);
  const jamBaru = timeToMinutes(newRow[11]);

  let insertIndex = dataRows.length + 1;
  for (let i = 0; i < dataRows.length; i++) {
    const jamLama = timeToMinutes(dataRows[i][11]);
    if (jamBaru < jamLama) {
      insertIndex = i + 2;
      break;
    }
  }

  if (insertIndex === 1) insertIndex = 2;
  sheet.insertRowBefore(insertIndex);
  sheet.getRange(insertIndex, 1, 1, newRow.length).setValues([newRow]);
  updateRowNumbers(sheet);
}

function timeToMinutes(jamStr) {
  if (!jamStr) return 0;
  if (jamStr instanceof Date) return jamStr.getHours() * 60 + jamStr.getMinutes();

  const parts = jamStr.toString().split(":");
  if (parts.length >= 2) {
    const jam = parseInt(parts[0], 10);
    const menit = parseInt(parts[1], 10);
    if (!isNaN(jam) && !isNaN(menit)) return jam * 60 + menit;
  }
  return 0;
}

function updateRowNumbers(sheet) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    sheet.getRange(i + 1, 2).setValue(i);
  }
}

function getOrCreateSheet(tanggal) {
  const date = new Date(tanggal);
  const folderName = `Booking_${date.getFullYear()}_${String(date.getMonth() + 1).padStart(2, '0')}`;
  const parentId = "1ypv_G7bkmc2BOZAHn6sbD50t9k4aJou1";
  const parent = DriveApp.getFolderById(parentId);
  const folders = parent.getFoldersByName(folderName);
  const folder = folders.hasNext() ? folders.next() : parent.createFolder(folderName);

  const spreadsheetName = folderName;
  let spreadsheet;
  const files = folder.getFilesByName(spreadsheetName);
  if (files.hasNext()) {
    spreadsheet = SpreadsheetApp.open(files.next());
  } else {
    const ss = SpreadsheetApp.create(spreadsheetName);
    DriveApp.getFileById(ss.getId()).moveTo(folder);
    spreadsheet = ss;
  }

  let sheet = spreadsheet.getSheetByName(tanggal);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(tanggal);
    sheet.appendRow([
      "ID_ROW", "No", "Status Kehadiran", "Nomor Polisi", "Tipe Kendaraan", "Tahun", "Tipe Telepon", "No HP",
      "Customer Akan Datang", "Nama Customer", "Tanggal Booking",
      "Tanggal Penerimaan", "Jam Penerimaan", "VIA", "Admin", "Jenis Pekerjaan", "Nama Sales"
    ]);
  }
  return sheet;
}

function getSheetBySheetName(tanggal, sheetName) {
  const date = new Date(tanggal);
  const folderName = `Booking_${date.getFullYear()}_${String(date.getMonth() + 1).padStart(2, '0')}`;
  const parentId = "1ypv_G7bkmc2BOZAHn6sbD50t9k4aJou1";
  const parent = DriveApp.getFolderById(parentId);
  const folders = parent.getFoldersByName(folderName);
  if (!folders.hasNext()) return null;
  const folder = folders.next();
  const files = folder.getFilesByName(folderName);
  if (!files.hasNext()) return null;
  const ss = SpreadsheetApp.open(files.next());
  return ss.getSheetByName(sheetName);
}

function formatTanggalLengkap(date) {
  const hari = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
  const bulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  return `${hari[date.getDay()]}, ${date.getDate()} ${bulan[date.getMonth()]} ${date.getFullYear()}`;
}

function doAddMRSV3(e) {
  const p = e.parameter;
  const tahun = p.tahun || new Date().getFullYear();
  const bulanInput = p.reminder_month;
  const bulanMap = {
    "Januari": "01", "Februari": "02", "Maret": "03", "April": "04", "Mei": "05", "Juni": "06",
    "Juli": "07", "Agustus": "08", "September": "09", "Oktober": "10", "November": "11", "Desember": "12"
  };
  const bulanAngka = bulanMap[bulanInput];
  if (!bulanAngka) return ContentService.createTextOutput("❌ Bulan tidak valid");

  // 1. Cari folder Booking_YYYY_MM
  const folderName = `Booking_${tahun}_${bulanAngka}`;
  const parentId = "1ypv_G7bkmc2BOZAHn6sbD50t9k4aJou1";
  const parent = DriveApp.getFolderById(parentId);
  let folder;
  const folders = parent.getFoldersByName(folderName);
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = parent.createFolder(folderName);
  }

  // 2. Cari spreadsheet Booking_YYYY_MM
  let spreadsheet;
  const files = folder.getFilesByName(folderName);
  if (files.hasNext()) {
    spreadsheet = SpreadsheetApp.open(files.next());
  } else {
    const ss = SpreadsheetApp.create(folderName);
    DriveApp.getFileById(ss.getId()).moveTo(folder);
    spreadsheet = ss;
  }

  // 3. Sheet MRS di paling kiri, nama: MRS_[NamaBulan]
  const sheetName = "MRS_" + bulanInput.toUpperCase();
  let mrsSheet = spreadsheet.getSheetByName(sheetName);
  if (!mrsSheet) {
    let firstSheet = spreadsheet.getSheets()[0];
    if (firstSheet.getName() === "Sheet1" && firstSheet.getLastRow() === 0) {
      firstSheet.setName(sheetName);
      mrsSheet = firstSheet;
    } else {
      mrsSheet = spreadsheet.insertSheet(sheetName, 0);
    }
    mrsSheet.appendRow(["Nomor Polisi", "Bulan", "CH", "Nama Customer", "No HP"]);
  } else {
    spreadsheet.setActiveSheet(mrsSheet);
    spreadsheet.moveActiveSheet(0);
  }

  // 4. Tambahkan data
  mrsSheet.appendRow([
    p.plat_mrs,
    bulanInput,
    p.ch_mrs,
    p.nama_customer_mrs,
    p.nohp_mrs
  ]);

  return ContentService.createTextOutput("✅ Data MRS berhasil ditambahkan.");
}

function doDeleteMRS(e) {
  const { id_row } = e.parameter;
  
  try {
    const mrsSS = getMRSSpreadsheet();
    const sheet = mrsSS.getSheetByName("MRS_Data");
    const data = sheet.getDataRange().getValues();
    const rowIdx = data.findIndex(row => row[0] === id_row);
    
    if (rowIdx === -1 || rowIdx < 1) {
      return ContentService.createTextOutput("❌ Data MRS tidak ditemukan");
    }
    
    sheet.deleteRow(rowIdx + 1);
    return ContentService.createTextOutput("✅ Data MRS berhasil dihapus");
    
  } catch (error) {
    return ContentService.createTextOutput(`❌ Error: ${error.toString()}`);
  }
}

function getMRSSpreadsheet() {
  const parentId = "1ypv_G7bkmc2BOZAHn6sbD50t9k4aJou1";
  const parent = DriveApp.getFolderById(parentId);
  const spreadsheetName = "MRS_System";
  
  // Check if MRS spreadsheet already exists
  const files = parent.getFilesByName(spreadsheetName);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }
  
  // Create new MRS spreadsheet
  const ss = SpreadsheetApp.create(spreadsheetName);
  const file = DriveApp.getFileById(ss.getId());
  file.moveTo(parent);
  
  // Create MRS_Data sheet with headers
  const sheet = ss.getSheetByName("Sheet1");
  sheet.setName("MRS_Data");
  sheet.appendRow([
    "ID_ROW", "Nomor Polisi", "Reminder Bulan", "CH", "Nama Customer", "Tanggal Input"
  ]);
  
  // Format headers
  sheet.getRange(1, 1, 1, 6).setFontWeight("bold");
  sheet.autoResizeColumns(1, 6);
  
  return ss;
} 