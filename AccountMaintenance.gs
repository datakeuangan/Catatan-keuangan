// ============================================================
// CATATAN KEUANGAN - BACKEND MODULE
// Generated from the verified 7F-16 baseline.
// Functions are preserved in full; only file ownership was reorganized.
// ============================================================

function migrasiEmailKeUserKey() {
  var emailLama = 'restorepoin12@gmail.com'; 
  var namaBaru = 'putu';                     
  var kodeBaru = '1212';                     

  var userKeyBaru = buatUserKey_(namaBaru, kodeBaru);
  var emailLamaLower = String(emailLama).trim().toLowerCase();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  gantiOwnerLama_(ss.getSheetByName('Transaksi'), 9, emailLamaLower, userKeyBaru);
  gantiOwnerLama_(ss.getSheetByName('Rekening'), 4, emailLamaLower, userKeyBaru);
  gantiOwnerLama_(ss.getSheetByName('Kategori'), 6, emailLamaLower, userKeyBaru);
  gantiOwnerLama_(ss.getSheetByName('Transfer'), 6, emailLamaLower, userKeyBaru);

  Logger.log('Migrasi selesai. Data email "' + emailLama + '" sekarang jadi userKey: ' + userKeyBaru);
}

function gantiOwnerLama_(sheet, kolomOwner, nilaiLama, nilaiBaru) {
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  for (var r = 2; r <= lastRow; r++) {
    var cell = sheet.getRange(r, kolomOwner);
    var isiSekarang = String(cell.getValue() || '').trim().toLowerCase();
    if (isiSekarang === nilaiLama) {
      cell.setValue(nilaiBaru);
    }
  }
}

function pastikanDataDefault_(userKey) {
  if (!userKey) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetInit = ss.getSheetByName('Init Akun');
  if (!sheetInit) {
    sheetInit = ss.insertSheet('Init Akun');
    sheetInit.appendRow(['UserKey']);
  }

  var dataInit = sheetInit.getDataRange().getValues();
  for (var i = 1; i < dataInit.length; i++) {
    if (String(dataInit[i][0] || '').trim() === userKey) return;
  }

  var sudahPunyaData = cekSudahPunyaData_(ss, userKey);
  if (!sudahPunyaData) {
    seedKategoriDefault_(ss, userKey);
    seedRekeningDefault_(ss, userKey);
  }

  sheetInit.appendRow([userKey]);
}

function cekSudahPunyaData_(ss, userKey) {
  var sheetKat = ss.getSheetByName('Kategori');
  if (sheetKat) {
    var dataKat = sheetKat.getDataRange().getValues();
    for (var i = 1; i < dataKat.length; i++) {
      if (String(dataKat[i][5] || '').trim() === userKey) return true;
    }
  }
  var sheetRek = ss.getSheetByName('Rekening');
  if (sheetRek) {
    var dataRek = sheetRek.getDataRange().getValues();
    for (var i = 1; i < dataRek.length; i++) {
      if (String(dataRek[i][3] || '').trim() === userKey) return true;
    }
  }
  return false;
}

function seedKategoriDefault_(ss, userKey) {
  var sheet = ss.getSheetByName('Kategori');
  if (!sheet) {
    sheet = ss.insertSheet('Kategori');
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Tipe", "Kategori", "Sub Kategori", "Icon Kategori", "Icon Sub Kategori", "UserKey"]);
  }

  var defaultKategori = [
    // Pemasukan (+)
    ["'+", "Pemasukan", "Gaji Utama", "💰", "💵", userKey],
    ["'+", "Pemasukan", "Bonus & Tunjangan", "💰", "🎁", userKey],
    ["'+", "Pemasukan", "Usaha & Bisnis", "💰", "💼", userKey],
    ["'+", "Pemasukan", "Investasi", "💰", "📈", userKey],
    ["'+", "Pemasukan", "Lain-lain", "💰", "🪙", userKey],

    // Pengeluaran (-)
    ["'-", "Makanan & Minuman", "Makan Harian", "🍔", "🍚", userKey],
    ["'-", "Makanan & Minuman", "Camilan & Jajanan", "🍔", "🍰", userKey],
    ["'-", "Makanan & Minuman", "Kopi & Minuman", "🍔", "☕", userKey],

    ["'-", "Belanja & Kebutuhan", "Pakaian & Fashion", "🛒", "👕", userKey],
    ["'-", "Belanja & Kebutuhan", "Kebutuhan Bulanan", "🛒", "📦", userKey],

    ["'-", "Transportasi", "Bensin", "⛽", "⛽", userKey],
    ["'-", "Transportasi", "Parkir & Tol", "⛽", "🅿️", userKey],
    ["'-", "Transportasi", "Transport Umum", "⛽", "🚌", userKey],

    ["'-", "Tagihan & Utilitas", "Listrik & Air", "🧾", "💡", userKey],
    ["'-", "Tagihan & Utilitas", "Pulsa & Paket Data", "🧾", "📱", userKey],
    ["'-", "Tagihan & Utilitas", "Internet WiFi", "🧾", "📡", userKey],

    ["'-", "Hiburan & Rekreasi", "Jalan-jalan", "🎉", "🏖️", userKey],
    ["'-", "Hiburan & Rekreasi", "Hobi & Game", "🎉", "🎮", userKey],

    ["'-", "Kesehatan", "Obat & Dokter", "🏥", "💊", userKey],
    ["'-", "Kesehatan", "Perawatan Diri", "🏥", "🧼", userKey]
  ];
  defaultKategori.forEach(function(row) { sheet.appendRow(row); });
}

function seedRekeningDefault_(ss, userKey) {
  var sheet = ss.getSheetByName('Rekening');
  if (!sheet) {
    sheet = ss.insertSheet('Rekening');
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Nama", "Saldo", "Icon", "UserKey"]);
  }

  var defaultRekening = [
    ["CASH / TUNAI", 0, "💵", userKey],
    ["MANDIRI", 0, "💳", userKey],
    ["BCA", 0, "💳", userKey],
    ["E-WALLET (DANA/OVO/GOPAY)", 0, "👛", userKey]
  ];
  defaultRekening.forEach(function(row) { sheet.appendRow(row); });
}

function resetSemuaDataAkun(userKey) {
  userKey = String(userKey || '').trim();

  if (!userKey) {
    return { sukses: false, pesan: 'UserKey tidak valid.' };
  }

  if (!validasiUserKey(userKey)) {
    return { sukses: false, pesan: 'Sesi akun tidak valid atau sudah tidak aktif.' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lock = LockService.getScriptLock();

  lock.waitLock(15000);

  try {
    // Hapus seluruh data finansial milik akun ini.
    hapusDataByUserKey_(ss.getSheetByName('Transaksi'), 9, userKey);
    hapusDataByUserKey_(ss.getSheetByName('Rekening'), 4, userKey);
    hapusDataByUserKey_(ss.getSheetByName('Kategori'), 6, userKey);
    hapusDataByUserKey_(ss.getSheetByName('Transfer'), 6, userKey);
    hapusDataByUserKey_(ss.getSheetByName('Goals'), 9, userKey);
    hapusDataByUserKey_(ss.getSheetByName('HutangPiutang'), 9, userKey);

    // Hapus penanda inisialisasi agar akun dapat dibuat ulang
    // ke kondisi awal.
    hapusDataByUserKey_(ss.getSheetByName('Init Akun'), 1, userKey);

    // Hapus pengaturan tampilan tersimpan untuk akun ini,
    // sehingga perangkat kembali menggunakan konfigurasi default.
    var sheetSettings = ss.getSheetByName('Pengaturan Pengguna');
    if (sheetSettings) {
      hapusDataByUserKey_(sheetSettings, 1, userKey);
    }

    // Kembalikan akun ke kondisi awal.
    // Akun "Pengguna" sendiri tetap dipertahankan.
    pastikanDataDefault_(userKey);

    SpreadsheetApp.flush();

    Logger.log('Reset semua data akun berhasil. UserKey: ' + userKey);

    return {
      sukses: true,
      pesan: 'Semua data akun berhasil direset. Akun tetap dipertahankan dan data default telah dibuat kembali.'
    };

  } catch (e) {
    Logger.log('Gagal reset data akun ' + userKey + ': ' + e);
    return {
      sukses: false,
      pesan: 'Gagal mereset data: ' + (e && e.message ? e.message : e)
    };
  } finally {
    lock.releaseLock();
  }
}
