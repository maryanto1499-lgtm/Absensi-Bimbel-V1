/**
 * KONFIGURASI SISTEM ABSENSI BIMBEL
 */
const CONFIG = {
  SHEET_SISWA: 'Siswa',
  SHEET_ABSENSI: 'Absensi',
  PIN_TENTOR: 'BIMBEL2026', // Ubah PIN keamanan ini sesuai keinginan Anda
  TIMEZONE: 'Asia/Jakarta', // WIB
  CEGAH_DOUBLE_SCAN_MENIT: 15 // Mencegah siswa ter-scan 2x dalam rentang menit ini
};

/**
 * Endpoint GET: Untuk cek status API atau tes langsung via browser
 * Contoh tes: URL_WEB_APP?pin=BIMBEL2026&id_siswa=SB001&tentor=Kak%20Budi
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.id_siswa) {
    return prosesAbsensi(e.parameter);
  }
  return buatResponJSON({
    status: 'online',
    pesan: 'API Absensi Bimbel aktif dan siap menerima data scan.'
  });
}

/**
 * Endpoint POST: Menerima data JSON/Form dari aplikasi scanner HP tentor
 */
function doPost(e) {
  try {
    let data = {};
    if (e.postData && e.postData.contents) {
      // Mendukung payload JSON maupun URL-encoded form
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter || {};
      }
    } else {
      data = e.parameter || {};
    }
    return prosesAbsensi(data);
  } catch (error) {
    return buatResponJSON({
      status: 'error',
      pesan: 'Terjadi kesalahan server: ' + error.message
    });
  }
}

/**
 * Fungsi Utama Validasi & Pencatatan Kehadiran
 */
function prosesAbsensi(params) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000); // Mencegah tabrakan data jika 2 tentor scan bersamaan

  try {
    const pin = String(params.pin || '').trim();
    const idSiswa = String(params.id_siswa || '').trim().toUpperCase();
    const namaTentor = String(params.tentor || 'Tentor').trim();
    const keterangan = String(params.keterangan || 'Hadir').trim();

    // 1. Validasi PIN Keamanan
    if (CONFIG.PIN_TENTOR && pin !== CONFIG.PIN_TENTOR) {
      return buatResponJSON({
        status: 'error',
        pesan: 'PIN Tentor salah! Akses ditolak.'
      });
    }

    // 2. Validasi ID Siswa tidak kosong
    if (!idSiswa) {
      return buatResponJSON({
        status: 'error',
        pesan: 'Kode Barcode/QR kosong atau tidak terbaca.'
      });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetSiswa = ss.getSheetByName(CONFIG.SHEET_SISWA);
    const sheetAbsensi = ss.getSheetByName(CONFIG.SHEET_ABSENSI);

    if (!sheetSiswa || !sheetAbsensi) {
      return buatResponJSON({
        status: 'error',
        pesan: 'Tab Siswa atau Absensi tidak ditemukan di Spreadsheet.'
      });
    }

    // 3. Cek Data Siswa di Tab 'Siswa' (Mulai baris ke-6 sesuai template)
    const lastRowSiswa = sheetSiswa.getLastRow();
    let dataSiswaDitemukan = null;

    if (lastRowSiswa >= 6) {
      const dataSiswa = sheetSiswa.getRange(6, 1, lastRowSiswa - 5, 4).getValues();
      for (let i = 0; i < dataSiswa.length; i++) {
        const rowId = String(dataSiswa[i][0]).trim().toUpperCase();
        if (rowId === idSiswa) {
          dataSiswaDitemukan = {
            id: rowId,
            nama: String(dataSiswa[i][1]).trim(),
            kelas: String(dataSiswa[i][2]).trim(),
            status: String(dataSiswa[i][3]).trim()
          };
          break;
        }
      }
    }

    if (!dataSiswaDitemukan) {
      return buatResponJSON({
        status: 'error',
        pesan: 'ID Siswa "' + idSiswa + '" tidak terdaftar di database!'
      });
    }

    if (dataSiswaDitemukan.status.toLowerCase() !== 'aktif') {
      return buatResponJSON({
        status: 'error',
        pesan: 'Siswa ' + dataSiswaDitemukan.nama + ' berstatus NONAKTIF.'
      });
    }

    // 4. Siapkan Waktu Sekarang (WIB)
    const now = new Date();
    const timestampStr = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    const tanggalStr = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    const jamStr = Utilities.formatDate(now, CONFIG.TIMEZONE, 'HH:mm');

    // 5. Cari baris kosong berikutnya di tab 'Absensi' (berdasarkan kolom A / Timestamp)
    // Kita tidak memakai appendRow() agar rumus ARRAYFORMULA di kolom E (Nama Siswa) tidak terganggu
    const colAValues = sheetAbsensi.getRange('A6:A').getValues();
    let nextRow = 6;
    for (let i = 0; i < colAValues.length; i++) {
      if (colAValues[i][0] !== '' && colAValues[i][0] !== null) {
        nextRow = 6 + i + 1;
      } else {
        break;
      }
    }

    // 6. Cek Double Scan (apakah siswa ini baru saja di-scan dalam X menit terakhir)
    if (nextRow > 6 && CONFIG.CEGAH_DOUBLE_SCAN_MENIT > 0) {
      const startCheck = Math.max(6, nextRow - 30); // Cek 30 entri terakhir
      const recentLogs = sheetAbsensi.getRange(startCheck, 1, nextRow - startCheck, 4).getValues();
      for (let i = recentLogs.length - 1; i >= 0; i--) {
        const logId = String(recentLogs[i][3]).trim().toUpperCase();
        if (logId === idSiswa) {
          const logTime = new Date(recentLogs[i][0]);
          const selisihMenit = (now.getTime() - logTime.getTime()) / (1000 * 60);
          if (!isNaN(selisihMenit) && selisihMenit >= 0 && selisihMenit < CONFIG.CEGAH_DOUBLE_SCAN_MENIT) {
            return buatResponJSON({
              status: 'warning',
              pesan: dataSiswaDitemukan.nama + ' sudah di-scan ' + Math.floor(selisihMenit) + ' menit lalu.',
              siswa: dataSiswaDitemukan,
              waktu: jamStr
            });
          }
          break;
        }
      }
    }

    // 7. Tulis Data ke Kolom A:D dan F:G (Kolom E dibiarkan terisi otomatis oleh rumus sheet)
    sheetAbsensi.getRange(nextRow, 1, 1, 4).setValues([[
      timestampStr,
      tanggalStr,
      jamStr,
      dataSiswaDitemukan.id
    ]]);
    sheetAbsensi.getRange(nextRow, 6, 1, 2).setValues([[
      namaTentor,
      keterangan
    ]]);

    return buatResponJSON({
      status: 'success',
      pesan: 'Berhasil absen: ' + dataSiswaDitemukan.nama + ' (' + jamStr + ' WIB)',
      siswa: dataSiswaDitemukan,
      waktu: jamStr,
      tanggal: tanggalStr
    });

  } finally {
    lock.releaseLock();
  }
}

/**
 * Helper untuk mengembalikan response JSON ke browser/HP
 */
function buatResponJSON(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}