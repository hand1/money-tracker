/**
 * Money Tracker App v1.0.0
 * Backend Script untuk Google Apps Script (GAS) dengan Spreadsheet Storage
 */

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Money Tracker App')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Inisialisasi atau dapatkan sheet secara aman dengan validasi header otomatis
function getOrCreateSheet(sheetName, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    // Format header agar rapi
    sheet.getRange(1, 1, 1, headers.length)
         .setFontWeight("bold")
         .setBackground("#16a34a")
         .setFontColor("white");
  }
  return sheet;
}

// Setup awal untuk seluruh tabel database di Spreadsheet
function initDatabase() {
  getOrCreateSheet('transaksi', ['ID', 'Tanggal', 'Kategori', 'Jumlah', 'Tipe', 'Keterangan']);
  getOrCreateSheet('kategori', ['Tipe', 'Nama']);
  getOrCreateSheet('kantong', ['ID', 'Nama', 'Target', 'Terkumpul']);
  getOrCreateSheet('anggaran', ['Kategori', 'Limit']);
  getOrCreateSheet('pengaturan', ['Kunci', 'Nilai']);
  getOrCreateSheet('pengingat', ['ID', 'Nama', 'Tanggal', 'Status']);
  
  // Set Kategori Default jika tabel kategori masih kosong
  var katSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('kategori');
  if (katSheet.getLastRow() === 1) {
    var defaultKategori = [
      ['Pemasukan', 'Gaji'], ['Pemasukan', 'Bonus'], ['Pemasukan', 'Investasi'], ['Pemasukan', 'Deposito'], ['Pemasukan', 'Dividen'], ['Pemasukan', 'Lain-lain'],
      ['Pengeluaran', 'Makan'], ['Pengeluaran', 'Transport'], ['Pengeluaran', 'Sewa Rumah'], ['Pengeluaran', 'Bensin'], ['Pengeluaran', 'Belanja'], ['Pengeluaran', 'Asuransi'], ['Pengeluaran', 'Kesehatan'], ['Pengeluaran', 'Lain-lain']
    ];
    defaultKategori.forEach(function(row) {
      katSheet.appendRow(row);
    });
  }

  // Set Pengaturan Default jika kosong
  var setSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('pengaturan');
  if (setSheet.getLastRow() === 1) {
    setSheet.appendRow(['username', 'Handy']);
    setSheet.appendRow(['isPinEnabled', 'false']);
    setSheet.appendRow(['savedPin', '1234']);
  }
}

// Membaca seluruh data dari Spreadsheet untuk dikirim ke Frontend
function loadAllData() {
  initDatabase();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Ambil Transaksi (Konversi tanggal ke format string YYYY-MM-DD agar aman diserialisasi)
  var transSheet = ss.getSheetByName('transaksi');
  var transValues = transSheet.getDataRange().getValues();
  var transaksi = [];
  for (var i = 1; i < transValues.length; i++) {
    var rawDate = transValues[i][1];
    var formattedDate = "";
    if (rawDate instanceof Date) {
      formattedDate = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else {
      formattedDate = String(rawDate);
    }
    transaksi.push({
      id: transValues[i][0],
      tanggal: formattedDate,
      kategori: transValues[i][2],
      jumlah: Number(transValues[i][3]),
      type: transValues[i][4],
      nama: transValues[i][5] // keterangan transaksi
    });
  }
  // Urutkan transaksi terbaru di atas
  transaksi.reverse();

  // 2. Ambil Kategori
  var katSheet = ss.getSheetByName('kategori');
  var katValues = katSheet.getDataRange().getValues();
  var kategoriObj = { "Pemasukan": [], "Pengeluaran": [] };
  for (var j = 1; j < katValues.length; j++) {
    var tipe = katValues[j][0];
    var namaKat = katValues[j][1];
    if (kategoriObj[tipe]) {
      kategoriObj[tipe].push(namaKat);
    }
  }

  // 3. Ambil Kantong
  var kantongSheet = ss.getSheetByName('kantong');
  var kantongValues = kantongSheet.getDataRange().getValues();
  var pockets = [];
  for (var k = 1; k < kantongValues.length; k++) {
    pockets.push({
      id: kantongValues[k][0],
      nama: kantongValues[k][1],
      target: Number(kantongValues[k][2]),
      terkumpul: Number(kantongValues[k][3])
    });
  }

  // 4. Ambil Anggaran (Budgeting)
  var angSheet = ss.getSheetByName('anggaran');
  var angValues = angSheet.getDataRange().getValues();
  var budgets = [];
  for (var l = 1; l < angValues.length; l++) {
    budgets.push({
      kategori: angValues[l][0],
      limit: Number(angValues[l][1])
    });
  }

  // 5. Ambil Pengingat (Reminder)
  var remSheet = ss.getSheetByName('pengingat');
  var remValues = remSheet.getDataRange().getValues();
  var reminders = [];
  for (var n = 1; n < remValues.length; n++) {
    var rawRemDate = remValues[n][2];
    var formattedRemDate = "";
    if (rawRemDate instanceof Date) {
      formattedRemDate = Utilities.formatDate(rawRemDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else {
      formattedRemDate = String(rawRemDate);
    }
    reminders.push({
      id: remValues[n][0],
      nama: remValues[n][1],
      tanggal: formattedRemDate,
      status: remValues[n][3]
    });
  }

  // 6. Ambil Pengaturan
  var setSheet = ss.getSheetByName('pengaturan');
  var setValues = setSheet.getDataRange().getValues();
  var settings = {};
  for (var m = 1; m < setValues.length; m++) {
    settings[setValues[m][0]] = setValues[m][1];
  }

  return {
    transaksi: transaksi,
    kategori: kategoriObj,
    pockets: pockets,
    budgets: budgets,
    reminders: reminders,
    settings: settings
  };
}

// ----------------- CRUD OPERATIONS ON SPREADSHEET -----------------

// Simpan atau Tambah Transaksi Baru
function addTransaction(date, category, amount, type, description) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('transaksi');
  var id = "TX" + new Date().getTime();
  sheet.appendRow([id, date, category, amount, type, description]);
  return { success: true };
}

// Tambah Kategori Baru
function addCategory(type, name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('kategori');
  sheet.appendRow([type, name]);
  return { success: true };
}

// Edit Kategori
function updateCategory(type, oldName, newName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('kategori');
  var range = sheet.getDataRange();
  var values = range.getValues();
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === type && values[i][1] === oldName) {
      sheet.getCell(i + 1, 2).setValue(newName);
      break;
    }
  }
  return { success: true };
}

// Hapus Kategori
function deleteCategory(type, name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('kategori');
  var range = sheet.getDataRange();
  var values = range.getValues();
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === type && values[i][1] === name) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return { success: true };
}

// Simpan Kantong (Tambah / Edit)
function savePocket(id, name, target, saved) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('kantong');
  var range = sheet.getDataRange();
  var values = range.getValues();
  var found = false;
  
  if (id) {
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] === id) {
        sheet.getRange(i + 1, 1, 1, 4).setValues([[id, name, target, saved]]);
        found = true;
        break;
      }
    }
  }
  
  if (!found) {
    var newId = "PK" + new Date().getTime();
    sheet.appendRow([newId, name, target, saved]);
  }
  return { success: true };
}

// Hapus Kantong
function deletePocket(id) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('kantong');
  var range = sheet.getDataRange();
  var values = range.getValues();
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === id) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return { success: true };
}

// Simpan Anggaran (Tambah / Edit Limit)
function saveBudget(category, limit) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('anggaran');
  var range = sheet.getDataRange();
  var values = range.getValues();
  var found = false;
  
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === category) {
      sheet.getRange(i + 1, 2).setValue(limit);
      found = true;
      break;
    }
  }
  
  if (!found) {
    sheet.appendRow([category, limit]);
  }
  return { success: true };
}

// Hapus Anggaran
function deleteBudget(category) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('anggaran');
  var range = sheet.getDataRange();
  var values = range.getValues();
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === category) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return { success: true };
}

// Simpan Pengingat (Tambah / Edit / Update Status Verifikasi)
function saveReminder(id, name, date, status) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('pengingat');
  var range = sheet.getDataRange();
  var values = range.getValues();
  var found = false;
  
  if (id) {
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] === id) {
        sheet.getRange(i + 1, 1, 1, 4).setValues([[id, name, date, status]]);
        found = true;
        break;
      }
    }
  }
  
  if (!found) {
    var newId = "RM" + new Date().getTime();
    sheet.appendRow([newId, name, date, status]);
  }
  return { success: true };
}

// Hapus Pengingat
function deleteReminder(id) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('pengingat');
  var range = sheet.getDataRange();
  var values = range.getValues();
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === id) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return { success: true };
}

// Update Pengaturan Secara Fleksibel (Username, Status PIN, Nilai PIN)
function updateSetting(key, value) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('pengaturan');
  var range = sheet.getDataRange();
  var values = range.getValues();
  var found = false;
  
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(String(value));
      found = true;
      break;
    }
  }
  
  if (!found) {
    sheet.appendRow([key, String(value)]);
  }
  return { success: true };
}
