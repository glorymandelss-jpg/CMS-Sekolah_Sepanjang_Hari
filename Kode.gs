function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('CMS Sekolah Sepanjang Hari')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 2. HELPER UTILS
function saveFileToDrive(base64Data, filename, mimeType) {
  var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, filename);
  var file = DriveApp.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// 3. MODUL 1: AUTENTIKASI
function validasiLogin(email, password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users");
  if (!sheet) return { status: "error", message: "Sheet Users tidak ditemukan!" };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var userEmail = String(data[i][2]).trim(); // Kolom Email
    var userPass = String(data[i][3]).trim();  // Kolom Password

    if (userEmail.toLowerCase() === String(email).toLowerCase() && userPass === String(password)) {
      return {
        status: "success",
        user: {
          id: String(data[i][0]),
          nama: String(data[i][1]),
          email: String(data[i][2]),
          role: String(data[i][4])
        }
      };
    }
  }

  return { status: "error", message: "Email atau Password salah!" };
}

// Alias untuk menjaga kompatibilitas dengan frontend
function prosesLogin(email, password) {
  return validasiLogin(email, password);
}

function catatPresensi(userId, fotoBase64, lat, lng, jenis) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // Tunggu kunci antrean maks 10 detik

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Presensi");
    if (!sheet) {
      sheet = ss.insertSheet("Presensi");
      sheet.appendRow(["id", "user_id", "tanggal", "jam_masuk", "jam_keluar", "foto_masuk_url", "foto_keluar_url", "lat_masuk", "lng_masuk", "lat_keluar", "lng_keluar", "status", "catatan"]);
    }

    var filename = "PRESENSI_" + userId + "_" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd_HHmmss") + ".jpg";
    var urlFoto = saveFileToDrive(fotoBase64, filename, "image/jpeg");

    var now = new Date();
    var tglFormatted = Utilities.formatDate(now, "GMT+7", "yyyy-MM-dd");
    var jamFormatted = Utilities.formatDate(now, "GMT+7", "HH:mm:ss");
    var presensiId = "PR" + now.getTime();

    sheet.appendRow([
      presensiId, userId, tglFormatted, jamFormatted, "", 
      urlFoto, "", lat, lng, "", "", "hadir", "Presensi via Web App"
    ]);

    return { status: "success", message: "Presensi " + jenis + " berhasil dicatat!" };
  } catch (err) {
    return { status: "error", message: "Gagal menyimpan presensi: " + err.toString() };
  } finally {
    lock.releaseLock();
  }
}

function getRekapPresensi() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Presensi");
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  var headers = data[0];
  var result = [];
  
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var val = data[i][j];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, "GMT+7", "yyyy-MM-dd HH:mm:ss");
      }
      obj[headers[j]] = String(val);
    }
    result.push(obj);
  }
  return result;
}

// 5. MODUL 3: MATERI PEMBELAJARAN
function uploadMateri(userId, judul, deskripsi, kategori, fileBase64, fileName, mimeType) {
  try {
    var fileUrl = saveFileToDrive(fileBase64, fileName, mimeType);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Materi");
    if (!sheet) {
      sheet = ss.insertSheet("Materi");
      sheet.appendRow(["id", "judul", "deskripsi", "file_url", "tipe_file", "ukuran_file", "uploaded_by", "kategori", "is_aktif", "created_at"]);
    }

    sheet.appendRow([
      "MAT" + new Date().getTime(), judul, deskripsi, fileUrl, mimeType, 0, userId, kategori, true, new Date()
    ]);
    return { status: "success", message: "Materi berhasil diunggah!" };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function getDaftarMateri() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Materi");
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  var headers = data[0];
  var result = [];
  
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = String(data[i][j]);
    }
    result.push(obj);
  }
  return result;
}

// 6. MODUL 4: LOG HARIAN
function uploadLogHarian(userId, bulan, tahun, catatanGuru, fileBase64, fileName, mimeType) {
  try {
    var fileUrl = saveFileToDrive(fileBase64, fileName, mimeType);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Log_Harian");
    if (!sheet) {
      sheet = ss.insertSheet("Log_Harian");
      sheet.appendRow(["id", "user_id", "bulan", "tahun", "file_laporan_url", "tipe_file", "catatan_guru", "status_review", "catatan_koordinator", "reviewed_by", "submitted_at", "reviewed_at"]);
    }

    sheet.appendRow([
      "LOG" + new Date().getTime(), userId, bulan, tahun, fileUrl, mimeType, catatanGuru, "pending", "", "", new Date(), ""
    ]);
    return { status: "success", message: "Laporan Log Harian berhasil dikirim!" };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function getDaftarLogHarian() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Log_Harian");
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  var headers = data[0];
  var result = [];
  
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var val = data[i][j];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, "GMT+7", "yyyy-MM-dd HH:mm:ss");
      }
      obj[headers[j]] = String(val);
    }
    result.push(obj);
  }
  return result;
}

function reviewLogHarian(logId, statusReview, catatanKoordinator, reviewerUserId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Log_Harian");
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === logId) {
        sheet.getRange(i + 1, 8).setValue(statusReview); // status_review
        sheet.getRange(i + 1, 9).setValue(catatanKoordinator); // catatan_koordinator
        sheet.getRange(i + 1, 10).setValue(reviewerUserId); // reviewed_by
        sheet.getRange(i + 1, 12).setValue(new Date()); // reviewed_at
        return { status: "success", message: "Status laporan berhasil diperbarui!" };
      }
    }
    return { status: "error", message: "ID Laporan tidak ditemukan." };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

// 7. MODUL 5 & 6: KELOLA USERS & SETTINGS
function getAllUsers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  var headers = data[0];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = String(data[i][j]);
    }
    result.push(obj);
  }
  return result;
}

function tambahUser(nama, email, password, role, noHp) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Users");
    var newId = "USR" + Math.floor(100 + Math.random() * 900);
    sheet.appendRow([newId, nama, email, password, role, noHp, "aktif", "", new Date()]);
    return { status: "success", message: "Pengguna baru berhasil ditambahkan!" };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function getWebsiteSettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Websie_Settings") || ss.getSheetByName("Website_Settings");
  if (!sheet) return {};
  
  var data = sheet.getDataRange().getValues();
  var settings = {};
  
  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][0]).trim();
    var val = String(data[i][1]).trim();
    if (key) {
      settings[key] = val;
    }
  }
  return settings;
}

function getTemplateLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Template_Log");
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var isAktif = String(data[i][4]).toUpperCase();
    if (isAktif === 'TRUE' || isAktif === '1') {
      result.push({
        id: String(data[i][0]),
        nama_template: String(data[i][1]),
        file_url: String(data[i][2]),
        deskripsi: String(data[i][3])
      });
    }
  }
  return result;
}