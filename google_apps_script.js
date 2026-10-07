/**
 * ============================================================================
 * GOOGLE APPS SCRIPT - RSVP & UCAPAN WEDDING KIKI & AMALDA
 * ============================================================================
 * 
 * PETUNJUK PEMASANGAN:
 * 1. Buka Google Spreadsheet: https://docs.google.com/spreadsheets/d/12AD2AMjzLGUaBmZAEVbohW02Q-W2YBPlB9ezHtZrF-c/edit
 * 2. Di menu atas, klik: Ekstensi (Extensions) > Apps Script
 * 3. Hapus semua kode yang ada di editor Apps Script, lalu Salin & Tempel seluruh kode di file ini.
 * 4. Klik tombol Simpan (ikon disket / Ctrl+S).
 * 5. (Opsional) Pilih fungsi "testKoneksi" di menu atas lalu klik "Jalankan" (Run) untuk menguji dan memberi izin akses Google.
 * 6. Klik tombol "Terapkan" (Deploy) di kanan atas > pilih "Penerapan baru" (New deployment).
 * 7. Klik ikon Roda Gigi (Select type) > pilih jenis: "Aplikasi Web" (Web app).
 * 8. Isi konfigurasi:
 *    - Deskripsi: RSVP Undangan Kiki & Amalda
 *    - Jalankan sebagai: "Saya" (Me)
 *    - Yang memiliki akses: "Siapa saja" (Anyone)  <-- PENTING!
 * 9. Klik "Terapkan" (Deploy). Berikan izin akses Google jika diminta.
 * 10. Salin "URL Aplikasi Web" (Web App URL) yang berakhiran `/exec`.
 * 11. Tempelkan URL tersebut ke variabel `GAS_WEBAPP_URL` di file index.html.
 * ============================================================================
 */

var SPREADSHEET_ID = "12AD2AMjzLGUaBmZAEVbohW02Q-W2YBPlB9ezHtZrF-c";
var SHEET_NAME = "Sheet1";

/**
 * Mendapatkan atau membuat sheet secara aman
 */
function getSheet() {
  var ss;
  try {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  
  if (!ss) {
    throw new Error("Spreadsheet tidak ditemukan! Pastikan ID Spreadsheet benar.");
  }
  
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  
  // Inisialisasi header jika sheet masih kosong
  if (sheet.getLastRow() === 0) {
    var headers = ["Timestamp", "Nama Tamu", "Ucapan", "Konfirmasi Kehadiran", "Jumlah Tamu"];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#685952")
      .setFontColor("#FFFFFF")
      .setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

/**
 * Fungsi untuk tes manual di Apps Script Editor
 * (Pilih 'testKoneksi' lalu klik tombol 'Jalankan' / 'Run')
 */
function testKoneksi() {
  var sheet = getSheet();
  Logger.log("✅ Berhasil terhubung ke spreadsheet: " + sheet.getName());
  Logger.log("Jumlah baris saat ini: " + sheet.getLastRow());
}

/**
 * GET Handler: Mengambil seluruh daftar ucapan dari Google Spreadsheet
 */
function doGet(e) {
  try {
    var sheet = getSheet();
    var data = [];
    var lastRow = sheet.getLastRow();
    
    if (lastRow > 1) {
      var rows = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
      for (var i = rows.length - 1; i >= 0; i--) { // Urutkan dari yang terbaru (DESC)
        var row = rows[i];
        if (row[1] || row[2]) { // Pastikan ada nama atau ucapan
          var timestampStr = "";
          if (row[0] instanceof Date) {
            timestampStr = Utilities.formatDate(row[0], "Asia/Jakarta", "dd MMM yyyy, HH:mm");
          } else {
            timestampStr = String(row[0] || "");
          }
          
          data.push({
            timestamp: timestampStr,
            nama: String(row[1] || "").trim(),
            ucapan: String(row[2] || "").trim(),
            kehadiran: String(row[3] || "").trim(),
            jumlah: String(row[4] || "1").trim()
          });
        }
      }
    }
    
    var response = {
      status: "success",
      total: data.length,
      data: data
    };
    
    return ContentService.createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    var errorResponse = {
      status: "error",
      message: err.toString()
    };
    return ContentService.createTextOutput(JSON.stringify(errorResponse))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * POST Handler: Menyimpan ucapan & RSVP baru ke Google Spreadsheet
 */
function doPost(e) {
  try {
    var sheet = getSheet();
    
    var params = {};
    if (e && e.postData && e.postData.contents) {
      try {
        params = JSON.parse(e.postData.contents);
      } catch (jsonErr) {
        params = e.parameter || {};
      }
    } else if (e && e.parameter) {
      params = e.parameter;
    }
    
    var nama = (params.nama || params.name || params.author || "").trim();
    var ucapan = (params.ucapan || params.message || params.comment || "").trim();
    var kehadiran = (params.kehadiran || params.attendance || "notsure").trim();
    var jumlah = (params.jumlah || params.guest || params.guests || "1").trim();
    
    // Normalisasi label kehadiran
    var labelKehadiran = "Hadir";
    if (kehadiran === "notpresent" || kehadiran.toLowerCase().indexOf("tidak") !== -1) {
      labelKehadiran = "Tidak Hadir";
      jumlah = "-";
    } else if (kehadiran === "present" || kehadiran.toLowerCase().indexOf("hadir") !== -1) {
      labelKehadiran = "Hadir";
    } else {
      labelKehadiran = "Ragu-ragu";
    }
    
    if (!nama) nama = "Tamu Undangan";
    if (!ucapan) ucapan = "-";
    
    var now = new Date();
    var formattedTime = Utilities.formatDate(now, "Asia/Jakarta", "dd MMM yyyy, HH:mm");
    
    // Simpan baris ke Google Sheets:
    // Kolom: 1. Timestamp | 2. Nama Tamu | 3. Ucapan | 4. Konfirmasi Kehadiran | 5. Jumlah Tamu
    sheet.appendRow([now, nama, ucapan, labelKehadiran, jumlah]);
    
    var result = {
      status: "success",
      message: "Ucapan berhasil disimpan!",
      entry: {
        timestamp: formattedTime,
        nama: nama,
        ucapan: ucapan,
        kehadiran: labelKehadiran,
        jumlah: jumlah
      }
    };
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    var errResult = {
      status: "error",
      message: err.toString()
    };
    return ContentService.createTextOutput(JSON.stringify(errResult))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
