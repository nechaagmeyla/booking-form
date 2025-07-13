// === APPS SCRIPT: Code.gs ===
// Hilangkan pembatasan email agar bisa menerima dari siapa saja
// function isAuthorized() {
//   const email = Session.getActiveUser().getEmail();
//   return email === "agmeylanecha@gmail.com";
// }

// SALES_LIST: Daftar sales yang valid untuk normalisasi dan rekap
const SALES_LIST = [
  'zen', 'januar', 'nia', 'avro', 'mila', 'indra', 'hendra', 'irvan', 'rony', 'stevy', 'osla', 'imron', 'asep aj', 'nadia', 'daulat', 'other'
];

function doGet(e) {
  const action = e.parameter.action;
  if (action === "get_mrs_view") return doGetMRSView(e);
  if (action === "get_mrs") return doGetMRS(e);
  if (action === "get_all_booking") return doGetAllBooking(e);
  
  const tanggal = e.parameter.tanggal_penerimaan;
  
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
  if (p.action === "update_sales_rekap") return ContentService.createTextOutput(updateRekapSalesSheet(p));
  if (p.action === "export_rekap_sales") return ContentService.createTextOutput(exportRekapSalesSheet(p));

  // === TAMBAH BOOKING ===
  const tglBooking = formatTanggalLengkap(new Date());
  const sheet = getOrCreateSheet(p.tanggal_penerimaan);
  const idRow = `${Date.now()}_${p.plat}`;

  // Normalisasi sales
  let salesVal = (p.sales || "").toString().trim().toLowerCase();
  if (!SALES_LIST.includes(salesVal) && salesVal !== "") {
    salesVal = p.sales; // biarkan sesuai input user (other)
  }

  // === LOGIKA NO HP & CUSTOMER AKAN DATANG ===
  let nohp = "";
  let datang = "";
  if ((p.tipe_telepon || "").toUpperCase() === "SA") {
    // Jika tipe telepon SA, No HP dan Customer Akan Datang = kode SA (RZ, PS, DB, DM, DR)
    nohp = p.nohp_sa || "";
    datang = p.datang_dropdown || "";
  } else {
    // Jika tipe telepon TELP, No HP = nomor HP (dengan prefix 62, tanpa label), Customer Akan Datang = input user
    let hp = (p.nohp || "").replace(/\D/g, "");
    if (hp.startsWith("62")) hp = hp.substring(2);
    if (hp.length > 0) {
      nohp = "62" + hp;
    } else {
      nohp = "";
    }
    datang = p.datang || "";
  }

  // Status Kehadiran harus selalu kosong saat tambah booking
  const statusKehadiran = "";
  const newRow = [
    idRow, "", statusKehadiran, p.plat, p.tipe, p.tahun, p.tipe_telepon || "", nohp, datang, p.nama,
    tglBooking, p.tanggal_penerimaan, p.jam_penerimaan, p.via, p.admin, p.keluhan, salesVal
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
  // Normalisasi sales
  let salesVal = (p.sales || "").toString().trim().toLowerCase();
  if (!SALES_LIST.includes(salesVal) && salesVal !== "") {
    salesVal = p.sales;
  }
  // === LOGIKA NO HP & CUSTOMER AKAN DATANG (sama seperti tambah) ===
  let nohp = "";
  let datang = "";
  if ((p.tipe_telepon || "").toUpperCase() === "SA") {
    nohp = p.nohp_sa || "";
    datang = p.datang_dropdown || "";
  } else {
    let hp = (p.nohp || "").replace(/\D/g, "");
    if (hp.startsWith("62")) hp = hp.substring(2);
    if (hp.length > 0) {
      nohp = "62" + hp;
    } else {
      nohp = "";
    }
    datang = p.datang || "";
  }
  // Status Kehadiran boleh diisi saat edit
  const statusKehadiran = p.status_kehadiran || "";
  const updateRow = [
    p.id_row, "", statusKehadiran, p.plat, p.tipe, p.tahun, p.tipe_telepon || "", nohp, datang, p.nama,
    tglBooking, p.tanggal_penerimaan, p.jam_penerimaan, p.via, p.admin, p.keluhan, salesVal
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
  const dataRows = allData.slice(1); // skip header
  const jamBaru = timeToMinutes(newRow[12]); // kolom ke-13: Jam Penerimaan

  let insertIndex = dataRows.length + 1; // default: paling bawah
  for (let i = 0; i < dataRows.length; i++) {
    const jamLama = timeToMinutes(dataRows[i][12]); // kolom ke-13: Jam Penerimaan
    // Jika jamBaru valid dan lebih kecil dari jamLama, sisipkan di sini
    if (jamBaru > 0 && (jamLama === 0 || jamBaru < jamLama)) {
      insertIndex = i + 2; // +2 karena header
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
  // Ambil hanya bagian jam dan menit (misal: '7:00', '09:00', '10:00')
  const match = jamStr.match(/^(\d{1,2}):(\d{2})/);
  if (match) {
    const jam = parseInt(match[1], 10);
    const menit = parseInt(match[2], 10);
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
  const correctHeader = [
    "ID_ROW", "No", "Status Kehadiran", "Nomor Polisi", "Tipe Kendaraan", "Tahun", "Tipe Telepon", "No HP",
    "Customer Akan Datang", "Nama Customer", "Tanggal Booking",
    "Tanggal Penerimaan", "Jam Penerimaan", "VIA", "Admin", "Jenis Pekerjaan", "Nama Sales"
  ];
  if (!sheet) {
    sheet = spreadsheet.insertSheet(tanggal);
    sheet.appendRow(correctHeader);
  } else {
    // Jika header sheet lama belum ada kolom 'Tipe Telepon', update header dan sisipkan kolom
    const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (header.length < correctHeader.length || header[6] !== "Tipe Telepon") {
      // Sisipkan kolom ke-7 (setelah Tahun)
      sheet.insertColumnAfter(6);
      sheet.getRange(1, 7).setValue("Tipe Telepon");
      // Geser data lama ke kanan jika perlu
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, 7, lastRow - 1, 1).setValue("");
      }
      // Update header lain jika perlu
      for (var i = 0; i < correctHeader.length; i++) {
        sheet.getRange(1, i + 1).setValue(correctHeader[i]);
      }
    }
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

function doGetAllBooking(e) {
  try {
    const allBookings = getAllBookingData();
    return ContentService.createTextOutput(JSON.stringify(allBookings)).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
  }
}

function getAllBookingData() {
  const parentId = "1ypv_G7bkmc2BOZAHn6sbD50t9k4aJou1";
  const parent = DriveApp.getFolderById(parentId);
  const allBookings = [];
  // Ambil semua folder Booking_YYYY_MM
  const folders = parent.getFolders();
  while (folders.hasNext()) {
    const folder = folders.next();
    const folderName = folder.getName();
    if (folderName.startsWith('Booking_')) {
      const files = folder.getFilesByName(folderName);
      if (files.hasNext()) {
        const ss = SpreadsheetApp.open(files.next());
        const sheets = ss.getSheets();
        for (let i = 0; i < sheets.length; i++) {
          const sheet = sheets[i];
          const sheetName = sheet.getName();
          // Skip MRS sheets dan sheet kosong
          if (sheetName.startsWith('MRS_') || sheet.getLastRow() <= 1) continue;
          const data = sheet.getDataRange().getValues();
          const headers = data[0];
          for (let j = 1; j < data.length; j++) {
            const row = data[j];
            const booking = {};
            for (let k = 0; k < headers.length; k++) {
              let val = row[k];
              if (val instanceof Date) val = val.toISOString();
              booking[headers[k].toLowerCase().replace(/ /g, "_")] = val;
            }
            // Normalisasi sales
            booking.sales = (booking.sales || "").toString().trim().toLowerCase();
            booking.nama_sales = (booking.nama_sales || "").toString().trim().toLowerCase();
            booking.sheetName = sheetName;
            booking.folderName = folderName;
            booking.id_row = row[0];
            allBookings.push(booking);
          }
        }
      }
    }
  }
  return allBookings;
}

// Tambahkan fungsi untuk mengunci sheet dari edit manual
function protectSheet(sheet) {
  var protection = sheet.protect();
  protection.setDescription('Sheet dikunci otomatis');
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
}

function updateSalesRekap() {
  try {
    const salesRekapSS = createSalesRekapSpreadsheet();
    const allBookings = getAllBookingData();
    // Filter only TAB bookings
    const tabBookings = allBookings.filter(booking => (booking.via || '').toLowerCase() === 'tab');
    const updateDate = new Date().toISOString();
    // Process Sales_Detail sheet
    const detailSheet = salesRekapSS.getSheetByName("Sales_Detail");
    if (detailSheet.getLastRow() > 1) {
      detailSheet.clear();
      detailSheet.appendRow([
        "ID", "Bulan", "Tahun", "Nama Sales", "Nomor Polisi", "Nama Customer", 
        "Tanggal Penerimaan", "Jam Penerimaan", "Via", "Admin", "Tanggal Update"
      ]);
    }
    tabBookings.forEach((booking, index) => {
      const tanggalPenerimaan = booking.tanggal_penerimaan || booking.tanggal_booking || '';
      const date = new Date(tanggalPenerimaan);
      const bulan = date.getMonth() + 1;
      const tahun = date.getFullYear();
      let sales = (booking.nama_sales || booking.sales || '').toString().trim().toLowerCase();
      if (!SALES_LIST.includes(sales) && sales !== "") {
        sales = booking.nama_sales || booking.sales || '';
      }
      detailSheet.appendRow([
        index + 1,
        bulan,
        tahun,
        sales,
        booking.nomor_polisi || booking.plat || '',
        booking.nama_customer || booking.nama || '',
        tanggalPenerimaan,
        booking.jam_penerimaan || '',
        booking.via || '',
        booking.admin || '',
        updateDate
      ]);
    });
    // Process Sales_Summary sheet
    const summarySheet = salesRekapSS.getSheetByName("Sales_Summary");
    if (summarySheet.getLastRow() > 1) {
      summarySheet.clear();
      summarySheet.appendRow([
        "Bulan", "Tahun", "Nama Sales", "Total Booking", "Via TAB", "Via Lainnya", "Tanggal Update"
      ]);
    }
    // Kumpulkan semua kombinasi bulan-tahun dari tabBookings
    const bulanTahunSet = new Set();
    tabBookings.forEach(booking => {
      const tanggalPenerimaan = booking.tanggal_penerimaan || booking.tanggal_booking || '';
      const date = new Date(tanggalPenerimaan);
      const bulan = date.getMonth() + 1;
      const tahun = date.getFullYear();
      bulanTahunSet.add(`${bulan}_${tahun}`);
    });
    // Hitung summary per bulan-tahun-sales
    const summaryData = {};
    tabBookings.forEach(booking => {
      const tanggalPenerimaan = booking.tanggal_penerimaan || booking.tanggal_booking || '';
      const date = new Date(tanggalPenerimaan);
      const bulan = date.getMonth() + 1;
      const tahun = date.getFullYear();
      let sales = (booking.nama_sales || booking.sales || '').toString().trim().toLowerCase();
      if (!SALES_LIST.includes(sales) && sales !== "") {
        sales = booking.nama_sales || booking.sales || '';
      }
      const via = booking.via || '';
      const key = `${bulan}_${tahun}_${sales}`;
      if (!summaryData[key]) {
        summaryData[key] = {
          bulan: bulan,
          tahun: tahun,
          sales: sales,
          total: 0,
          viaTab: 0,
          viaLainnya: 0
        };
      }
      summaryData[key].total++;
      if (via.toLowerCase() === 'tab') {
        summaryData[key].viaTab++;
      } else {
        summaryData[key].viaLainnya++;
      }
    });
    // Untuk setiap kombinasi bulan-tahun, pastikan semua sales ada (meskipun 0)
    bulanTahunSet.forEach(bulanTahun => {
      const [bulan, tahun] = bulanTahun.split('_');
      SALES_LIST.forEach(sales => {
        const key = `${bulan}_${tahun}_${sales}`;
        const summary = summaryData[key] || {
          bulan: Number(bulan),
          tahun: Number(tahun),
          sales: sales,
          total: 0,
          viaTab: 0,
          viaLainnya: 0
        };
        summarySheet.appendRow([
          summary.bulan,
          summary.tahun,
          summary.sales,
          summary.total,
          summary.viaTab,
          summary.viaLainnya,
          updateDate
        ]);
      });
    });
    // Kunci sheet summary & detail
    protectSheet(summarySheet);
    protectSheet(detailSheet);
    // Urutkan sheet: MRS (jika ada), Sales_Summary, Sales_Detail
    var mrsSheet = salesRekapSS.getSheetByName('MRS');
    if (mrsSheet) {
      salesRekapSS.setActiveSheet(mrsSheet);
      salesRekapSS.moveActiveSheet(0);
    }
    salesRekapSS.setActiveSheet(summarySheet);
    salesRekapSS.moveActiveSheet(1);
    salesRekapSS.setActiveSheet(detailSheet);
    salesRekapSS.moveActiveSheet(2);
    // Format header
    detailSheet.getRange(1, 1, 1, 11).setFontWeight("bold");
    summarySheet.getRange(1, 1, 1, 7).setFontWeight("bold");
    return "✅ Sales Rekap berhasil diupdate";
  } catch (error) {
    return `❌ Error: ${error.toString()}`;
  }
} 

// Fungsi untuk update sheet REKAP_SALES_BULAN di spreadsheet Booking_YYYY_MM
function updateRekapSalesSheet(p) {
  try {
    const bulanNum = p.bulan || '';
    const tahun = p.tahun || '';
    if (!bulanNum || !tahun) return '❌ Bulan/tahun tidak valid';
    // Nama bulan
    const bulanMap = {
      '01': 'JANUARI', '02': 'FEBRUARI', '03': 'MARET', '04': 'APRIL', '05': 'MEI', '06': 'JUNI',
      '07': 'JULI', '08': 'AGUSTUS', '09': 'SEPTEMBER', '10': 'OKTOBER', '11': 'NOVEMBER', '12': 'DESEMBER'
    };
    const bulanStr = bulanMap[bulanNum] || bulanNum;
    const folderName = `Booking_${tahun}_${bulanNum}`;
    const parentId = "1ypv_G7bkmc2BOZAHn6sbD50t9k4aJou1";
    const parent = DriveApp.getFolderById(parentId);
    const folders = parent.getFoldersByName(folderName);
    if (!folders.hasNext()) return '❌ Folder tidak ditemukan';
    const folder = folders.next();
    const files = folder.getFilesByName(folderName);
    if (!files.hasNext()) return '❌ Spreadsheet tidak ditemukan';
    const ss = SpreadsheetApp.open(files.next());
    // Nama sheet rekap
    const rekapSheetName = `REKAP_SALES_${bulanStr}`;
    // Hapus sheet jika sudah ada
    const oldSheet = ss.getSheetByName(rekapSheetName);
    if (oldSheet) ss.deleteSheet(oldSheet);
    // Buat sheet baru
    const rekapSheet = ss.insertSheet(rekapSheetName);
    // Data booking bulan ini
    const allBookings = getAllBookingData();
    const filtered = allBookings.filter(row => {
      if ((row.via || '').toLowerCase() !== 'tab') return false;
      const tgl = row.tanggal_penerimaan || row.tanggal_booking || '';
      if (!tgl) return false;
      const d = new Date(tgl);
      const m = (d.getMonth() + 1).toString().padStart(2, '0');
      const y = d.getFullYear().toString();
      return m === bulanNum && y === tahun;
    });
    // Hitung jumlah booking per sales
    const salesCount = {};
    SALES_LIST.forEach(s => salesCount[s] = 0);
    filtered.forEach(row => {
      let salesVal = (row.nama_sales || row.sales || '').toString().trim().toLowerCase();
      if (SALES_LIST.includes(salesVal)) salesCount[salesVal]++;
    });
    // Header
    rekapSheet.appendRow(['Nama Sales', 'Jumlah Booking']);
    SALES_LIST.forEach(sales => {
      if (sales === 'other') return;
      rekapSheet.appendRow([sales.toUpperCase(), salesCount[sales]]);
    });
    // Tempatkan sheet setelah MRS_BULAN
    const mrsSheetName = `MRS_${bulanStr}`;
    const mrsSheet = ss.getSheetByName(mrsSheetName);
    if (mrsSheet) {
      ss.setActiveSheet(rekapSheet);
      ss.moveActiveSheet(mrsSheet.getIndex() + 1);
    }
    return `✅ Sheet ${rekapSheetName} berhasil diupdate`;
  } catch (err) {
    return `❌ Error: ${err}`;
  }
} 

// Fungsi untuk export hasil view rekap & grafik ke sheet EXPORT_REKAP_SALES_BULAN
function exportRekapSalesSheet(p) {
  try {
    const bulanNum = p.bulan || '';
    const tahun = p.tahun || '';
    if (!bulanNum || !tahun) return '❌ Bulan/tahun tidak valid';
    const bulanMap = {
      '01': 'JANUARI', '02': 'FEBRUARI', '03': 'MARET', '04': 'APRIL', '05': 'MEI', '06': 'JUNI',
      '07': 'JULI', '08': 'AGUSTUS', '09': 'SEPTEMBER', '10': 'OKTOBER', '11': 'NOVEMBER', '12': 'DESEMBER'
    };
    const bulanStr = bulanMap[bulanNum] || bulanNum;
    const folderName = `Booking_${tahun}_${bulanNum}`;
    const parentId = "1ypv_G7bkmc2BOZAHn6sbD50t9k4aJou1";
    const parent = DriveApp.getFolderById(parentId);
    const folders = parent.getFoldersByName(folderName);
    if (!folders.hasNext()) return '❌ Folder tidak ditemukan';
    const folder = folders.next();
    const files = folder.getFilesByName(folderName);
    if (!files.hasNext()) return '❌ Spreadsheet tidak ditemukan';
    const ss = SpreadsheetApp.open(files.next());
    // Nama sheet export
    const exportSheetName = `EXPORT_REKAP_SALES_${bulanStr}`;
    // Hapus sheet jika sudah ada
    const oldSheet = ss.getSheetByName(exportSheetName);
    if (oldSheet) ss.deleteSheet(oldSheet);
    // Buat sheet baru
    const exportSheet = ss.insertSheet(exportSheetName);
    // Data booking bulan ini
    const allBookings = getAllBookingData();
    const filtered = allBookings.filter(row => {
      if ((row.via || '').toLowerCase() !== 'tab') return false;
      const tgl = row.tanggal_penerimaan || row.tanggal_booking || '';
      if (!tgl) return false;
      const d = new Date(tgl);
      const m = (d.getMonth() + 1).toString().padStart(2, '0');
      const y = d.getFullYear().toString();
      return m === bulanNum && y === tahun;
    });
    // Hitung jumlah booking per sales
    const salesCount = {};
    SALES_LIST.forEach(s => salesCount[s] = 0);
    filtered.forEach(row => {
      let salesVal = (row.nama_sales || row.sales || '').toString().trim().toLowerCase();
      if (SALES_LIST.includes(salesVal)) salesCount[salesVal]++;
    });
    // Header tabel utama
    exportSheet.appendRow(['Nama Sales', 'Jumlah Booking']);
    SALES_LIST.forEach(sales => {
      if (sales === 'other') return;
      exportSheet.appendRow([sales.toUpperCase(), salesCount[sales]]);
    });
    // Data untuk grafik: sales dengan booking > 0
    exportSheet.appendRow(['']);
    exportSheet.appendRow(['Data Grafik']);
    exportSheet.appendRow(['Nama Sales', 'Jumlah Booking']);
    SALES_LIST.forEach(sales => {
      if (sales === 'other') return;
      if (salesCount[sales] > 0) {
        exportSheet.appendRow([sales.toUpperCase(), salesCount[sales]]);
      }
    });
    // Tempatkan sheet setelah REKAP_SALES_BULAN (atau setelah MRS jika belum ada)
    const rekapSheetName = `REKAP_SALES_${bulanStr}`;
    const mrsSheetName = `MRS_${bulanStr}`;
    const rekapSheet = ss.getSheetByName(rekapSheetName);
    const mrsSheet = ss.getSheetByName(mrsSheetName);
    if (rekapSheet) {
      ss.setActiveSheet(exportSheet);
      ss.moveActiveSheet(rekapSheet.getIndex() + 1);
    } else if (mrsSheet) {
      ss.setActiveSheet(exportSheet);
      ss.moveActiveSheet(mrsSheet.getIndex() + 1);
    }
    return `✅ Sheet ${exportSheetName} berhasil diexport`;
  } catch (err) {
    return `❌ Error: ${err}`;
  }
} 