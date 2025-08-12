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
  const action = e.parameter.action || "";
  const bulan = e.parameter.bulan || "";
  const tahun = e.parameter.tahun || "";
  const id_row = e.parameter.id_row || "";
  const tanggal = e.parameter.tanggal_penerimaan || "";

  // Ambil data MRS
  if (action === "get_mrs") return doGetMRS(e);

  // Ambil data MRS view
  if (action === "get_mrs_view") return getMRSView(e);

  // Ambil semua booking
  if (action === "get_all_booking") return doGetAllBooking(e);

  // THS
  if (action === "get_ths") {
    return ContentService
      .createTextOutput(JSON.stringify(getTHSData(bulan, tahun)))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Gabungan TAB + THS
  if (action === "get_ths_sales_merge") {
    return ContentService
      .createTextOutput(JSON.stringify(getTHSSalesMerge(bulan, tahun)))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Booking + Sales merge
  if (action === "get_booking_sales_merge") {
    return ContentService
      .createTextOutput(JSON.stringify(getBookingSalesMerge(bulan, tahun)))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Default: ambil data booking per tanggal
  if (!tanggal) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: "Tanggal kosong" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

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
    row.id_row = data[i][0]; // Add id_row for edit/delete
    result.push(row);
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getTHSData(bulan, tahun) {
  const parentId = "1ypv_G7bkmc2BOZAHn6sbD50t9k4aJou1";
  const parent = DriveApp.getFolderById(parentId);

  const folderName = `THS_${tahun}_${bulan}`;
  const folders = parent.getFoldersByName(folderName);
  if (!folders.hasNext()) return [];

  const folder = folders.next();
  const files = folder.getFilesByName(folderName);
  if (!files.hasNext()) return [];

  const ss = SpreadsheetApp.open(files.next());

  const bulanNamaMap = {
    "01": "JANUARI", "02": "FEBRUARI", "03": "MARET", "04": "APRIL",
    "05": "MEI", "06": "JUNI", "07": "JULI", "08": "AGUSTUS",
    "09": "SEPTEMBER", "10": "OKTOBER", "11": "NOVEMBER", "12": "DESEMBER"
  };
  const bulanNama = bulanNamaMap[bulan] || bulan;

  const sheetName = `THS_${bulanNama}`;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().trim());
  const result = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i].join("").trim() === "") continue;
    const row = {};
    for (let j = 0; j < headers.length; j++) {
      let val = data[i][j];
      if (val instanceof Date) val = val.toISOString();
      row[headers[j].toLowerCase().replace(/ /g, "_")] = val;
    }
    row.source_type = "ths";
    if (!row.id_unik) row.id_unik = `ths_${tahun}${bulan}_${i}`;
    result.push(row);
  }
  return result;
}

function getBookingSalesMerge(bulan, tahun) {
  // Ambil data dari booking bengkel via TAB
  let tabData = getAllBookingData()
    .filter(row => (row.via || "").toLowerCase() === "tab");

  if (bulan && tahun) {
    tabData = tabData.filter(row => {
      const d = new Date(row.tanggal_penerimaan || row.tanggal_booking || "");
      return (d.getMonth() + 1).toString().padStart(2, "0") === bulan &&
             d.getFullYear().toString() === tahun;
    });
  }

  // Tambahkan id_unik untuk TAB
  const tabWithType = tabData.map((row, idx) => {
    row.source_type = "bengkel";
    row.id_unik = `tab_${tahun}${bulan}_${idx + 1}`;
    return row;
  });

  // Ambil data THS sesuai bulan & tahun
  const thsWithType = getTHSData(bulan, tahun).map((row, idx) => {
    row.source_type = "ths";
    if (!row.id_unik) row.id_unik = `ths_${tahun}${bulan}_${idx + 1}`;
    return row;
  });

  // Gabungkan tanpa duplikat berdasarkan id_unik
  const mergedMap = {};
  [...tabWithType, ...thsWithType].forEach(item => {
    if (!mergedMap[item.id_unik]) {
      mergedMap[item.id_unik] = item;
    }
  });

  // Normalisasi nama sales
  Object.values(mergedMap).forEach(item => {
    let rawName = (item.nama_sales || item.sales || "").trim();
    if (!rawName) rawName = "Tanpa Nama";
    item.nama_sales_key = rawName.toLowerCase(); // untuk penggabungan
    item.nama_sales = rawName.toLowerCase().replace(/\b\w/g, c => c.toUpperCase()); // untuk tampilan
  });

  return Object.values(mergedMap);
}

//dogetmrs

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

function getTHSSalesMerge(bulan, tahun) {
  // Ambil data THS saja dari folder THS_TAHUN_BULAN
  const thsData = getTHSData(bulan, tahun); // sudah ada id_unik & source_type
  
  return thsData.map((row, idx) => ({
    ...row,
    id_unik: row.id_unik || `ths_${tahun}${bulan}_${idx + 1}`,
    source_type: "ths"
  }));
}


function getBookingSalesMerge(bulan, tahun) {
  let tabData = getAllBookingData()
    .filter(row => (row.via || "").toLowerCase() === "tab");

  if (bulan && tahun) {
    tabData = tabData.filter(row => {
      const d = new Date(row.tanggal_penerimaan || row.tanggal_booking || "");
      return (d.getMonth() + 1).toString().padStart(2, "0") === bulan && d.getFullYear().toString() === tahun;
    });
  }

  // Tambahkan id_unik untuk TAB
  const tabWithType = tabData.map((row, idx) => {
    row.source_type = "bengkel";
    row.id_unik = `tab_${tahun}${bulan}_${idx + 1}`;
    return row;
  });

  const thsWithType = getTHSData(bulan, tahun); // sudah ada id_unik

  return [...tabWithType, ...thsWithType];
}


function getMRSView(e) {
  const bulan = parseInt(e.parameter.bulan, 10);
  const tahun = parseInt(e.parameter.tahun, 10);

  if (!bulan || !tahun) {
    return ContentService.createTextOutput(JSON.stringify({ current: [], previous: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Hitung bulan sebelumnya
  let prevMonth = bulan - 1;
  let prevYear = tahun;
  if (prevMonth < 1) {
    prevMonth = 12;
    prevYear -= 1;
  }

  // Ambil data bulan sekarang dan sebelumnya
  const currentData = readMRSFile(tahun, bulan);
  const prevData = readMRSFile(prevYear, prevMonth);

  const result = {
    current: currentData,
    previous: prevData
  };

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Ambil data dari file Booking_{tahun}_{bulan}
 * bulan dalam angka (1-12)
 */
function readMRSFile(tahun, bulan) {
  const monthStr = bulan.toString().padStart(2, "0");
  const fileName = `Booking_${tahun}_${monthStr}`;
  const sheetName = `MRS_${getMonthName(bulan)}`;

  const files = DriveApp.getFilesByName(fileName);
  if (!files.hasNext()) return [];

  const file = files.next();
  const ss = SpreadsheetApp.open(file);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const result = [];

  for (let i = 1; i < data.length; i++) {
    result.push({
      nomor_polisi: data[i][0],
      bulan: data[i][1],
      tahun: tahun, // biar bisa hapus tanpa id_row
      ch: data[i][2],
      nama_customer: data[i][3],
      no_hp: data[i][4],
      status: data[i][5],
      catatan: data[i][6],
      admin_update: data[i][7],
      tanggal_update: data[i][8]
    });
  }
  return result;
}


// Helper: nomor bulan → nama bulan
function getMonthName(num) {
  const bulanMap = [
    '', 'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
    'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
  ];
  return bulanMap[num] || '';
}


function doPost(e) {
  const p = e.parameter;
  const action = e.parameter.action;

  // === Handle Booking THS ===
  if (action === "add_ths_booking") return doAddTHS(e);
  if (action === "edit_ths_booking") return editTHSBooking(e);
  if (action === "delete_ths_booking") return deleteTHSBooking(e);

  // === Handle MRS ===
  if (action === "update_mrs_status") return doUpdateMRSStatus(e);
  if (action === "add_mrs") return doAddMRSV3(e);
 if (action === "delete_mrs") {
  return doDeleteMRS(e); // langsung jalankan, tanpa cek id_row
}

  // === Handle Booking Umum ===
  if (action === "delete") return doDelete(e);
  if (action === "edit") return doEdit(e);
  if (action === "export_rekap") return doExportRekap(e);
  if (action === "update_sales_rekap") return ContentService.createTextOutput(updateRekapSalesSheet(p));
  if (action === "export_rekap_sales") return ContentService.createTextOutput(exportRekapSalesSheet(p));
  if (action === "update_kedatangan") return doUpdateKedatangan(e);

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

function updateMRSStatusHandler(e) {
  const p = e.parameter;
  const idRow = Number(p.id_row);
  const bulan = p.bulan;
  const tahun = p.tahun;

  if (!idRow || !bulan || !tahun) {
    return ContentService.createTextOutput("❌ Parameter 'tahun', 'bulan', dan 'id_row' harus disediakan.");
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MRS");
  const status = p.status || "";
  const catatan = p.catatan || "";
  const admin = p.admin || "";
  const tanggalUpdate = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");

  // Asumsi kolom: A=Nomor Polisi, B=Bulan, C=CH, D=Nama Customer, E=No HP, F=Status, G=Catatan, H=Admin, I=Tanggal Update
  sheet.getRange(idRow, 6).setValue(status); // Status
  sheet.getRange(idRow, 7).setValue(catatan); // Catatan
  sheet.getRange(idRow, 8).setValue(admin);   // Admin
  sheet.getRange(idRow, 9).setValue(tanggalUpdate); // Tanggal Update

  return ContentService.createTextOutput("✅ Status berhasil diupdate");
}

function doAddTHS(e) {
  const p = e.parameter;
  const tgl = new Date(p.tanggal_penerimaan);
  const tahun = tgl.getFullYear();
  const bulan = String(tgl.getMonth() + 1).padStart(2, '0');

  const bulanNamaMap = {
    "01": "JANUARI", "02": "FEBRUARI", "03": "MARET", "04": "APRIL",
    "05": "MEI", "06": "JUNI", "07": "JULI", "08": "AGUSTUS",
    "09": "SEPTEMBER", "10": "OKTOBER", "11": "NOVEMBER", "12": "DESEMBER"
  };
  const bulanNama = bulanNamaMap[bulan];

  // === UBAH: simpan di folder THS_TAHUN_BULAN ===
  const parentId = "1ypv_G7bkmc2BOZAHn6sbD50t9k4aJou1";
  const parent = DriveApp.getFolderById(parentId);
  const folderName = `THS_${tahun}_${bulan}`;
  
  let folder;
  let folders = parent.getFoldersByName(folderName);
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = parent.createFolder(folderName);
  }

  // === UBAH: spreadsheet juga pakai nama THS_TAHUN_BULAN ===
  let spreadsheet;
  const files = folder.getFilesByName(folderName);
  if (files.hasNext()) {
    spreadsheet = SpreadsheetApp.open(files.next());
  } else {
    const ss = SpreadsheetApp.create(folderName);
    folder.addFile(DriveApp.getFileById(ss.getId()));
    DriveApp.getRootFolder().removeFile(DriveApp.getFileById(ss.getId())); // pindahkan
    spreadsheet = ss;
  }

  const sheetName = `THS_${bulanNama}`;
  let thsSheet = spreadsheet.getSheetByName(sheetName);
  if (!thsSheet) {
    thsSheet = spreadsheet.insertSheet(sheetName);
    thsSheet.appendRow(["ID Unik", "Nama Sales", "Nomor Polisi", "Keterangan", "Tanggal Penerimaan"]);
  }

  const idUnik = `ths_${tahun}${bulan}_${Date.now()}`;
  thsSheet.appendRow([
    idUnik,
    p.nama_sales,
    p.nopol,
    p.keterangan,
    p.tanggal_penerimaan
  ]);

  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    message: "✅ Data THS berhasil ditambahkan.",
    id_unik: idUnik
  })).setMimeType(ContentService.MimeType.JSON);
}



function getFolderByName(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : null;
}

function doDeleteMRS(e) {
  const nomorPolisi = (e.parameter.nomor_polisi || "").trim().toUpperCase();
  const bulan = e.parameter.bulan;
  const tahun = e.parameter.tahun;

  if (!nomorPolisi || !bulan || !tahun) {
    return ContentService.createTextOutput("❌ Parameter 'nomor_polisi', 'bulan', dan 'tahun' wajib diisi.");
  }

  const monthNum = getMonthNumber(bulan);
  const fileName = `Booking_${tahun}_${monthNum}`;
  const sheetName = `MRS_${bulan.toUpperCase()}`;

  const files = DriveApp.getFilesByName(fileName);
  if (!files.hasNext()) {
    return ContentService.createTextOutput("❌ File tidak ditemukan.");
  }

  const ss = SpreadsheetApp.open(files.next());
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return ContentService.createTextOutput("❌ Sheet tidak ditemukan.");
  }

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (
      (data[i][0] || "").toString().trim().toUpperCase() === nomorPolisi &&
      (data[i][1] || "").toString().trim() === bulan
    ) {
      sheet.deleteRow(i + 1);
      return ContentService.createTextOutput("✅ Data berhasil dihapus.");
    }
  }

  return ContentService.createTextOutput("❌ Data tidak ditemukan.");
}



function doEdit(e) {
  const p = e.parameter;
  const sheet = getSheetBySheetName(p.tanggal_penerimaan, p.sheetName);
  if (!sheet) return ContentService.createTextOutput("❌ Sheet tidak ditemukan");
  const data = sheet.getDataRange().getValues();
  const rowIdx = data.findIndex(row => String(row[0]).trim() === String(p.id_row).trim());
  if (rowIdx === -1) return ContentService.createTextOutput("❌ Data tidak ditemukan");
  
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
  
  // Status Kehadiran boleh diisi saat edit - pastikan tidak kosong jika sudah diisi
  const statusKehadiran = p.status_kehadiran || data[rowIdx][2] || "";
  
  // Update existing row directly - NEVER create new data
  const updateRow = [
    p.id_row, 
    data[rowIdx][1], // Keep existing No
    statusKehadiran, 
    p.plat, 
    p.tipe, 
    p.tahun, 
    p.tipe_telepon || "", 
    nohp, 
    datang, 
    p.nama,
    data[rowIdx][10], // Keep existing Tanggal Booking
    p.tanggal_penerimaan, 
    p.jam_penerimaan, 
    p.via, 
    p.admin, 
    p.keluhan, 
    salesVal
  ];
  
  // Update the existing row directly - NO insertion or deletion
  sheet.getRange(rowIdx + 1, 1, 1, updateRow.length).setValues([updateRow]);
  
  // DO NOT call insertDataSortedByJam or any function that adds new rows
  // Only update row numbers to ensure consistency
  updateRowNumbers(sheet);
  
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
  // Setelah insert, urutkan ulang seluruh data berdasarkan jam penerimaan ascending
  const allRows = sheet.getDataRange().getValues();
  if (allRows.length > 2) {
    const header = allRows[0];
    const rows = allRows.slice(1);
    rows.sort((a, b) => timeToMinutes(a[12]) - timeToMinutes(b[12]));
    // Tulis ulang data (tanpa header)
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
    updateRowNumbers(sheet);
  }
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

function updateExistingRow(sheet, rowIdx, updateData) {
  // Update existing row without deleting and reinserting
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const updateRow = [];
  
  // Build update row based on headers
  for (let i = 0; i < headers.length; i++) {
    if (updateData[headers[i].toLowerCase().replace(/ /g, "_")]) {
      updateRow.push(updateData[headers[i].toLowerCase().replace(/ /g, "_")]);
    } else {
      // Keep existing value
      updateRow.push(sheet.getRange(rowIdx + 1, i + 1).getValue());
    }
  }
  
  // Update the row directly
  sheet.getRange(rowIdx + 1, 1, 1, updateRow.length).setValues([updateRow]);
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
    "Januari": "01", "Februari": "02", "Maret": "03", "April": "04",
    "Mei": "05", "Juni": "06", "Juli": "07", "Agustus": "08",
    "September": "09", "Oktober": "10", "November": "11", "Desember": "12"
  };
  const bulanAngka = bulanMap[bulanInput];
  if (!bulanAngka) return ContentService.createTextOutput("❌ Bulan tidak valid");

  // 1. Cari folder
  const folderName = `Booking_${tahun}_${bulanAngka}`;
  const parentId = "1ypv_G7bkmc2BOZAHn6sbD50t9k4aJou1";
  const parent = DriveApp.getFolderById(parentId);
  let folder = parent.getFoldersByName(folderName).hasNext()
    ? parent.getFoldersByName(folderName).next()
    : parent.createFolder(folderName);

  // 2. Cari spreadsheet
  let spreadsheet = folder.getFilesByName(folderName).hasNext()
    ? SpreadsheetApp.open(folder.getFilesByName(folderName).next())
    : (() => {
        const ss = SpreadsheetApp.create(folderName);
        DriveApp.getFileById(ss.getId()).moveTo(folder);
        return ss;
      })();

  // 3. Sheet MRS
  const sheetName = "MRS_" + bulanInput.toUpperCase();
  let mrsSheet = spreadsheet.getSheetByName(sheetName);
  if (!mrsSheet) {
    mrsSheet = spreadsheet.insertSheet(sheetName, 0);
    mrsSheet.appendRow(["Nomor Polisi", "Bulan", "CH", "Nama Customer", "No HP", "Status", "Catatan", "Admin Update", "Tanggal Update"]);
  }

  // 4. Tambahkan data
  const adminUpdate = p.admin_update || "";
  const tanggalUpdate = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");

  mrsSheet.appendRow([
    p.plat_mrs,
    bulanInput,
    p.ch_mrs,
    p.nama_customer_mrs,
    p.nohp_mrs,
    "Belum Dihubungi",
    "",
    adminUpdate,
    tanggalUpdate
  ]);

  return ContentService.createTextOutput("✅ Data MRS berhasil ditambahkan.");
}



function doUpdateMRSStatus(e) {
  const p = e.parameter;
  const tahun = p.tahun || new Date().getFullYear();
  const bulanInput = p.bulan;
  const nopol = (p.nomor_polisi || "").replace(/\s+/g, "").toUpperCase(); // hilangkan spasi

  if (!nopol) {
    return ContentService.createTextOutput("❌ Nomor polisi harus diisi");
  }

  const bulanMap = {
    "Januari": "01",
    "Februari": "02",
    "Maret": "03",
    "April": "04",
    "Mei": "05",
    "Juni": "06",
    "Juli": "07",
    "Agustus": "08",
    "September": "09",
    "Oktober": "10",
    "November": "11",
    "Desember": "12"
  };
  const bulanAngka = bulanMap[bulanInput];
  if (!bulanAngka) return ContentService.createTextOutput("❌ Bulan tidak valid");

  // Akses folder & file
  const folderName = `Booking_${tahun}_${bulanAngka}`;
  const parentId = "1ypv_G7bkmc2BOZAHn6sbD50t9k4aJou1";
  const parent = DriveApp.getFolderById(parentId);
  const folders = parent.getFoldersByName(folderName);
  if (!folders.hasNext()) return ContentService.createTextOutput("❌ Folder tidak ditemukan");
  const folder = folders.next();

  const files = folder.getFilesByName(folderName);
  if (!files.hasNext()) return ContentService.createTextOutput("❌ Spreadsheet tidak ditemukan");
  const ss = SpreadsheetApp.open(files.next());

  const sheetName = "MRS_" + bulanInput.toUpperCase();
  const mrsSheet = ss.getSheetByName(sheetName);
  if (!mrsSheet) return ContentService.createTextOutput("❌ Sheet MRS tidak ditemukan");

  const data = mrsSheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().trim().toLowerCase());

  // Cari index kolom berdasarkan header
  const colNopol = headers.indexOf("nomor polisi");
  const colStatus = headers.indexOf("status");
  const colCatatan = headers.indexOf("catatan");
  const colAdmin = headers.indexOf("admin update");
  const colTanggal = headers.indexOf("tanggal update");

  if (colNopol === -1) {
    return ContentService.createTextOutput("❌ Kolom 'Nomor Polisi' tidak ditemukan di sheet.");
  }

  const status = p.status || "";
  const catatan = p.catatan || "";
  const adminUpdate = p.admin || "";
  let tanggalUpdate = p.tanggal_update || "";

  if (status.toLowerCase().includes("sudah") && !tanggalUpdate) {
    tanggalUpdate = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");
  }

  let updated = false;
  for (let i = 1; i < data.length; i++) {
    const plat = (data[i][colNopol] || "").toString().replace(/\s+/g, "").toUpperCase();
    if (plat === nopol) {
      if (colStatus !== -1) mrsSheet.getRange(i + 1, colStatus + 1).setValue(status);
      if (colCatatan !== -1) mrsSheet.getRange(i + 1, colCatatan + 1).setValue(catatan);
      if (colAdmin !== -1) mrsSheet.getRange(i + 1, colAdmin + 1).setValue(adminUpdate);
      if (colTanggal !== -1) mrsSheet.getRange(i + 1, colTanggal + 1).setValue(tanggalUpdate);
      updated = true;
      break;
    }
  }

  if (!updated) {
    return ContentService.createTextOutput("❌ Data dengan nomor polisi tersebut tidak ditemukan");
  }

  return ContentService.createTextOutput("✅ Status MRS berhasil diperbarui");
}



function doDeleteMRS(e) {
  const p = e.parameter;
  const tahun = p.tahun;
  const bulanInput = p.bulan;
  const nopol = (p.nomor_polisi || "").toUpperCase().trim();

  if (!tahun || !bulanInput || !nopol) {
    return ContentService.createTextOutput("❌ Parameter 'tahun', 'bulan', dan 'nomor_polisi' harus disediakan.");
  }

  const bulanMap = {
    "Januari": "01", "Februari": "02", "Maret": "03", "April": "04",
    "Mei": "05", "Juni": "06", "Juli": "07", "Agustus": "08",
    "September": "09", "Oktober": "10", "November": "11", "Desember": "12"
  };
  const bulanAngka = bulanMap[bulanInput];
  if (!bulanAngka) return ContentService.createTextOutput("❌ Bulan tidak valid");

  try {
    // Akses file
    const folderName = `Booking_${tahun}_${bulanAngka}`;
    const parentId = "1ypv_G7bkmc2BOZAHn6sbD50t9k4aJou1";
    const parent = DriveApp.getFolderById(parentId);
    const folder = parent.getFoldersByName(folderName).next();
    const ss = SpreadsheetApp.open(folder.getFilesByName(folderName).next());

    const sheetName = "MRS_" + bulanInput.toUpperCase();
    const mrsSheet = ss.getSheetByName(sheetName);
    if (!mrsSheet) return ContentService.createTextOutput("❌ Sheet tidak ditemukan");

    // Cari baris berdasarkan nomor polisi
    const data = mrsSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if ((data[i][0] || "").toUpperCase().trim() === nopol) {
        mrsSheet.deleteRow(i + 1);
        return ContentService.createTextOutput("✅ Data MRS berhasil dihapus.");
      }
    }

    return ContentService.createTextOutput("❌ Data tidak ditemukan.");
  } catch (err) {
    return ContentService.createTextOutput("❌ Gagal menghapus data MRS: " + err.message);
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

function doUpdateKedatangan(e) {
  const p = e.parameter;
  const { id_row, sheetName, status_kehadiran, catatan } = p;
  
  console.log('doUpdateKedatangan received:', { id_row, sheetName, status_kehadiran, catatan });
  
  if (!id_row || !sheetName || !status_kehadiran) {
    return ContentService.createTextOutput("❌ Data tidak lengkap");
  }
  
  try {
    // Get all bookings to find the target
    const allBookings = getAllBookingData();
    console.log('Total bookings found:', allBookings.length);
    
    // Find the target booking
    const targetBooking = allBookings.find(booking => 
      String(booking.id_row) === String(id_row)
    );
    
    if (!targetBooking) {
      console.log('Target booking not found for id_row:', id_row);
      return ContentService.createTextOutput("❌ Data booking tidak ditemukan");
    }
    
    console.log('Found target booking:', targetBooking);
    
    // Get the sheet directly using the booking data
    const sheet = getSheetBySheetName(targetBooking.tanggal_penerimaan, targetBooking.sheetName);
    if (!sheet) {
      return ContentService.createTextOutput("❌ Sheet tidak ditemukan");
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    console.log('Sheet headers:', headers);
    
    // Find the row by id_row
    const rowIdx = data.findIndex(row => String(row[0]) === String(id_row));
    if (rowIdx === -1) {
      console.log('Row not found for id_row:', id_row);
      return ContentService.createTextOutput("❌ Data tidak ditemukan di sheet");
    }
    
    console.log('Found row at index:', rowIdx);
    console.log('Row data before update:', data[rowIdx]);
    
    // Find Status Kehadiran column - use exact column index 2 (Status Kehadiran)
    const statusColIdx = 2; // Column C is index 2 (0-based)
    console.log('Using status column index:', statusColIdx);
    console.log('Status column header:', headers[statusColIdx]);
    
    // Get current value before update
    const currentValue = data[rowIdx][statusColIdx];
    console.log('Current value before update:', currentValue);
    
    // Update the status directly in the data array and then write it back
    data[rowIdx][statusColIdx] = status_kehadiran;
    
    // Add catatan if provided
    if (catatan && catatan.trim() !== '') {
      const keluhanColIdx = 15; // Jenis Pekerjaan column
      const currentKeluhan = data[rowIdx][keluhanColIdx] || '';
      const updatedKeluhan = currentKeluhan + (currentKeluhan ? ' | ' : '') + `Catatan: ${catatan}`;
      data[rowIdx][keluhanColIdx] = updatedKeluhan;
    }
    
    // Write the entire updated row back to the sheet
    console.log('Writing updated row:', data[rowIdx]);
    sheet.getRange(rowIdx + 1, 1, 1, data[rowIdx].length).setValues([data[rowIdx]]);
    
    // Verify the update by reading it back
    const verificationValue = sheet.getRange(rowIdx + 1, statusColIdx + 1).getValue();
    console.log('Verification - Updated value:', verificationValue);
    
    if (String(verificationValue).toLowerCase() !== String(status_kehadiran).toLowerCase()) {
      console.log('Warning: Update verification failed');
      console.log('Expected:', status_kehadiran, 'Got:', verificationValue);
      return ContentService.createTextOutput("❌ Gagal mengupdate status kehadiran - verifikasi gagal");
    }
    
    // Update row numbers to ensure consistency
    updateRowNumbers(sheet);
    
    console.log('Status kedatangan updated successfully:', status_kehadiran);
    
    return ContentService.createTextOutput("✅ Status kedatangan berhasil diupdate");
    
  } catch (error) {
    console.error('Error in doUpdateKedatangan:', error);
    return ContentService.createTextOutput(`❌ Error: ${error.toString()}`);
  }
} 