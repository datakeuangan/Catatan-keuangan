// ============================================================
// CATATAN KEUANGAN - BACKEND MODULE
// Generated from the verified 7F-16 baseline.
// Functions are preserved in full; only file ownership was reorganized.
// ============================================================

function getOrCreatePenggunaSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Pengguna');
  if (!sheet) {
    sheet = ss.insertSheet('Pengguna');
    sheet.appendRow(['Nama', 'Kode', 'Admin', 'Status', 'Email', 'Tanggal']);
  } else {
    if (String(sheet.getRange(1, 3).getValue() || '').trim() === '') {
      sheet.getRange(1, 3).setValue('Admin');
    }
    if (String(sheet.getRange(1, 4).getValue() || '').trim() === '') {
      sheet.getRange(1, 4).setValue('Status');
    }
    if (String(sheet.getRange(1, 5).getValue() || '').trim() === '') {
      sheet.getRange(1, 5).setValue('Email');
    }
    if (String(sheet.getRange(1, 6).getValue() || '').trim() === '') {
      sheet.getRange(1, 6).setValue('Tanggal');
    }
  }
  return sheet;
}

function buatUserKey_(nama, kode) {
  return String(nama).trim().toLowerCase() + '::' + String(kode).trim();
}

function cekLogin(nama, kode) {
  var hasil = { sukses: false };
  if (!nama || !kode) return hasil;

  var sheet = getOrCreatePenggunaSheet_();
  var data = sheet.getDataRange().getValues();
  var namaInput = String(nama).trim();
  var kodeInput = String(kode).trim();
  if (!namaInput || !kodeInput) return hasil;

  for (var i = 1; i < data.length; i++) {
    var namaBaris = String(data[i][0] || '').trim();
    var kodeBaris = String(data[i][1] || '').trim();
    if (namaBaris && namaBaris.toLowerCase() === namaInput.toLowerCase() && kodeBaris === kodeInput) {
      var status = String(data[i][3] || 'APPROVED').trim().toUpperCase();
      
      if (status === 'PENDING') {
        hasil.sukses = false;
        hasil.status = 'PENDING';
        hasil.pesan = 'Akun Anda sedang menunggu persetujuan Admin.';
        return hasil;
      }
      
      if (status === 'REJECTED') {
        hasil.sukses = false;
        hasil.status = 'REJECTED';
        hasil.pesan = 'Pendaftaran akun Anda ditolak oleh Admin.';
        return hasil;
      }

      var userKey = buatUserKey_(namaBaris, kodeBaris);
      pastikanDataDefault_(userKey);
      hasil.sukses = true;
      hasil.userKey = userKey;
      hasil.nama = namaBaris;
      return hasil;
    }
  }
  
  hasil.pesan = 'Nama atau Kode Akses salah.';
  return hasil;
}

function getAdminEmailFromSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheetPengguna = ss.getSheetByName('Pengguna');
  if (sheetPengguna) {
    var data = sheetPengguna.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var isAdminFlag = String(data[i][2] || '').trim().toUpperCase();
      var emailUser = String(data[i][4] || '').trim();
      if ((isAdminFlag === 'TRUE' || isAdminFlag === 'ADMIN' || isAdminFlag === 'YES' || isAdminFlag === '1') && emailUser && emailUser.indexOf('@') !== -1) {
        return emailUser;
      }
    }
  }

  if (sheetPengguna) {
    var data = sheetPengguna.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var emailUser = String(data[i][4] || '').trim();
      if (emailUser && emailUser.indexOf('@') !== -1) {
        return emailUser;
      }
    }
  }

  try {
    var ownerEmail = Session.getEffectiveUser().getEmail();
    if (ownerEmail) return ownerEmail;
  } catch (e) {}

  return '';
}

function daftarPenggunaBaru(nama, kode, email) {
  if (!nama || !kode) {
    return { sukses: false, pesan: 'Nama dan Kode Akses wajib diisi!' };
  }

  var namaClean = String(nama).trim();
  var kodeClean = String(kode).trim();
  var emailClean = String(email || '').trim();

  var sheet = getOrCreatePenggunaSheet_();
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var namaBaris = String(data[i][0] || '').trim();
    if (namaBaris.toLowerCase() === namaClean.toLowerCase()) {
      return { sukses: false, pesan: 'Nama "' + namaClean + '" sudah terdaftar. Silakan gunakan nama lain.' };
    }
  }

  var tz = 'Asia/Jakarta';
  try { tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(); } catch (e) {}
  var nowStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  sheet.appendRow([namaClean, kodeClean, '', 'PENDING', emailClean, nowStr]);

  kirimNotifikasiEmailAdmin_(namaClean, emailClean, nowStr);

  return {
    sukses: true,
    pesan: 'Pendaftaran berhasil! Akun Anda sedang menunggu persetujuan Admin.'
  };
}

function kirimEmailAmanHtml_(targetEmail, subject, body, htmlBody) {
  try {
    MailApp.sendEmail({
      to: targetEmail,
      subject: subject,
      body: body,
      htmlBody: htmlBody
    });
    return true;
  } catch (e1) {
    try {
      GmailApp.sendEmail(targetEmail, subject, body, {
        htmlBody: htmlBody
      });
      return true;
    } catch (e2) {
      Logger.log('Gagal kirim email: ' + e2.message);
      return false;
    }
  }
}

function kirimNotifikasiEmailAdmin_(namaPendaftar, emailPendaftar, tanggal) {
  var adminEmail = getAdminEmailFromSheet_();
  if (!adminEmail) {
    Logger.log('Gagal: Alamat email admin tidak ditemukan.');
    return 'Alamat email admin belum diisi!';
  }

  var subject = '🔔 [Catatan Keuangan] Permintaan Pendaftaran Akun Baru: ' + namaPendaftar;
  
  var body = 'Halo Admin,\n\n' +
    'Ada permintaan pendaftaran akun baru di aplikasi Catatan Keuangan:\n\n' +
    '• Nama Pendaftar: ' + namaPendaftar + '\n' +
    '• Email Pendaftar: ' + (emailPendaftar ? emailPendaftar : '(tidak dicantumkan)') + '\n' +
    '• Tanggal: ' + tanggal + '\n\n' +
    'Silakan buka aplikasi Catatan Keuangan lalu masuk ke menu "🛠️ Kelola Pengguna" untuk menyetujui atau menolak pendaftaran ini.\n\n' +
    'Terima kasih.';

  var htmlBody = '<div style="font-family: Arial, sans-serif; padding: 20px; background-color: #121212; color: #e0e0e0;">' +
    '<div style="max-width: 500px; margin: 0 auto; background: #1e1e1e; padding: 20px; border-radius: 12px; border: 1px solid #333;">' +
    '<h2 style="color: #9c27b0; margin-top: 0;">🔔 Permintaan Pendaftaran Baru</h2>' +
    '<p style="color:#bbb;">Halo Admin, ada pengguna baru yang mengajukan pendaftaran akun di <b>Catatan Keuangan</b>:</p>' +
    '<table style="width: 100%; border-collapse: collapse; margin: 15px 0; color: #fff;">' +
    '<tr><td style="padding: 10px; border-bottom:1px solid #333; font-weight: bold; width: 40%;">Nama:</td><td style="padding: 10px; border-bottom:1px solid #333;">' + namaPendaftar + '</td></tr>' +
    '<tr><td style="padding: 10px; border-bottom:1px solid #333; font-weight: bold;">Email:</td><td style="padding: 10px; border-bottom:1px solid #333;">' + (emailPendaftar ? emailPendaftar : '<i style="color:#888;">Tidak dicantumkan</i>') + '</td></tr>' +
    '<tr><td style="padding: 10px; border-bottom:1px solid #333; font-weight: bold;">Waktu:</td><td style="padding: 10px; border-bottom:1px solid #333;">' + tanggal + '</td></tr>' +
    '</table>' +
    '<p style="color:#aaa; font-size: 13px;">Silakan buka Web App Catatan Keuangan Anda, masuk sebagai Admin, lalu buka menu <b>🛠️ Kelola Pengguna</b> untuk menyetujui (Approve) atau menolak pendaftaran ini.</p>' +
    '</div></div>';

  var terkirim = kirimEmailAmanHtml_(adminEmail, subject, body, htmlBody);
  if (terkirim) {
    Logger.log('Email notifikasi berhasil dikirim ke: ' + adminEmail);
    return 'Email terkirim ke ' + adminEmail;
  } else {
    Logger.log('Gagal menginstal/mengirim email.');
    return 'Gagal kirim email notifikasi';
  }
}

function tesKirimEmailAdmin() {
  var email = getAdminEmailFromSheet_();
  if (!email) {
    Logger.log('ERROR: Alamat email admin masih kosong. Silakan isi alamat email Anda di sheet Pengguna!');
    return 'ERROR: Isi email Anda di sheet Pengguna!';
  }
  
  var terkirim = kirimEmailAmanHtml_(email, '[Tes Email Catatan Keuangan]', 'Ini adalah email uji coba dari sistem Catatan Keuangan.', '<h3>Tes Email Catatan Keuangan</h3><p>Email notifikasi berfungsi dengan baik!</p>');
  if (terkirim) {
    Logger.log('SUKSES: Email uji coba berhasil dikirim ke ' + email);
    return 'SUKSES kirim ke ' + email;
  } else {
    Logger.log('ERROR: Belum ada izin email. Jalankan fungsi mintaIzinEmail terlebih dahulu!');
    return 'ERROR: Silakan pilih dan jalankan fungsi mintaIzinEmail untuk memberikan izin!';
  }
}

var REMEMBER_SESSION_PREFIX_ = 'AUTH_REMEMBER_SESSION_';
var REMEMBER_SESSION_DURATION_MS_ = 30 * 24 * 60 * 60 * 1000;

function buatSesiIngatSaya(userKey) {
  if (!userKey || !validasiUserKey(userKey)) {
    return { sukses: false, pesan: 'Sesi login tidak valid.' };
  }

  var token = Utilities.getUuid() + '-' + Utilities.getUuid();
  var now = Date.now();
  var payload = {
    userKey: String(userKey),
    createdAt: now,
    expiresAt: now + REMEMBER_SESSION_DURATION_MS_
  };

  PropertiesService.getScriptProperties().setProperty(
    REMEMBER_SESSION_PREFIX_ + token,
    JSON.stringify(payload)
  );

  return {
    sukses: true,
    token: token,
    expiresAt: payload.expiresAt
  };
}

function validasiSesiIngatSaya(token) {
  if (!token) return { sukses: false };

  var props = PropertiesService.getScriptProperties();
  var key = REMEMBER_SESSION_PREFIX_ + String(token);
  var raw = props.getProperty(key);
  if (!raw) return { sukses: false };

  var payload;
  try {
    payload = JSON.parse(raw);
  } catch (e) {
    props.deleteProperty(key);
    return { sukses: false };
  }

  if (!payload || !payload.userKey || !payload.expiresAt || Date.now() >= Number(payload.expiresAt)) {
    props.deleteProperty(key);
    return { sukses: false };
  }

  if (!validasiUserKey(payload.userKey)) {
    props.deleteProperty(key);
    return { sukses: false };
  }

  var sheet = getOrCreatePenggunaSheet_();
  var data = sheet.getDataRange().getValues();
  var nama = '';
  for (var i = 1; i < data.length; i++) {
    var uk = buatUserKey_(data[i][0], data[i][1]);
    if (uk === payload.userKey) {
      nama = String(data[i][0] || '').trim();
      break;
    }
  }

  if (!nama) {
    props.deleteProperty(key);
    return { sukses: false };
  }

  return {
    sukses: true,
    userKey: payload.userKey,
    nama: nama,
    expiresAt: payload.expiresAt
  };
}

function cabutSesiIngatSaya(token) {
  if (!token) return false;
  PropertiesService.getScriptProperties().deleteProperty(
    REMEMBER_SESSION_PREFIX_ + String(token)
  );
  return true;
}

function validasiUserKey(userKey) {
  if (!userKey) return false;
  var sheet = getOrCreatePenggunaSheet_();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var uk = buatUserKey_(data[i][0], data[i][1]);
    var status = String(data[i][3] || 'APPROVED').trim().toUpperCase();
    if (uk === userKey && status !== 'PENDING' && status !== 'REJECTED') return true;
  }
  return false;
}

function cekApakahAdmin(userKey) {
  if (!userKey) return false;
  var sheet = getOrCreatePenggunaSheet_();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var uk = buatUserKey_(data[i][0], data[i][1]);
    if (uk === userKey) {
      var flag = String(data[i][2] || '').trim().toUpperCase();
      return flag === 'TRUE' || flag === 'YES' || flag === 'ADMIN' || flag === '1';
    }
  }
  return false;
}

function getDaftarPengguna(userKey) {
  if (!cekApakahAdmin(userKey)) return [];
  var sheet = getOrCreatePenggunaSheet_();
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    var nama = String(data[i][0] || '').trim();
    var status = String(data[i][3] || 'APPROVED').trim().toUpperCase();
    if (nama && status !== 'PENDING') {
      list.push({
        row: i + 1,
        nama: nama,
        kode: String(data[i][1] || '').trim(),
        admin: String(data[i][2] || '').trim().toUpperCase() === 'TRUE',
        status: status,
        email: String(data[i][4] || '').trim()
      });
    }
  }
  return list;
}

function getDaftarPenggunaPending(userKey) {
  if (!cekApakahAdmin(userKey)) return [];
  var sheet = getOrCreatePenggunaSheet_();
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    var nama = String(data[i][0] || '').trim();
    var status = String(data[i][3] || 'APPROVED').trim().toUpperCase();
    if (nama && status === 'PENDING') {
      list.push({
        row: i + 1,
        nama: nama,
        kode: String(data[i][1] || '').trim(),
        email: String(data[i][4] || '').trim(),
        tanggal: String(data[i][5] || '').trim()
      });
    }
  }
  return list;
}

function setujuiPengguna(rowNum, userKey) {
  if (!cekApakahAdmin(userKey)) return 'Ditolak';
  var sheet = getOrCreatePenggunaSheet_();
  if (rowNum < 2 || rowNum > sheet.getLastRow()) return 'Gagal';

  sheet.getRange(rowNum, 4).setValue('APPROVED');
  
  var namaPendaftar = String(sheet.getRange(rowNum, 1).getValue() || '').trim();
  var emailPendaftar = String(sheet.getRange(rowNum, 5).getValue() || '').trim();

  var kodePendaftar = String(sheet.getRange(rowNum, 2).getValue() || '').trim();
  var uk = buatUserKey_(namaPendaftar, kodePendaftar);
  pastikanDataDefault_(uk);

  if (emailPendaftar) {
    var subject = '[Catatan Keuangan] Akun Anda Telah Disetujui!';
    var body = 'Halo ' + namaPendaftar + ',\n\n' +
      'Pendaftaran akun Anda di Catatan Keuangan telah DISETUJUI oleh Admin.\n' +
      'Anda sekarang bisa masuk (login) menggunakan Nama dan Kode Akses yang telah Anda daftarkan.\n\n' +
      'Selamat menggunakan!';
    var htmlBody = '<div style="font-family: Arial, sans-serif; padding: 20px; background-color: #121212; color: #e0e0e0;">' +
      '<div style="max-width: 500px; margin: 0 auto; background: #1e1e1e; padding: 20px; border-radius: 12px; border: 1px solid #333;">' +
      '<h2 style="color: #4caf50; margin-top: 0;">🎉 Akun Disetujui!</h2>' +
      '<p>Halo <b>' + namaPendaftar + '</b>,</p>' +
      '<p>Pendaftaran akun Anda di <b>Catatan Keuangan</b> telah <b>DISETUJUI</b> oleh Admin.</p>' +
      '<p>Anda sekarang dapat masuk ke aplikasi menggunakan Nama dan Kode Akses Anda.</p>' +
      '</div></div>';
    kirimEmailAmanHtml_(emailPendaftar, subject, body, htmlBody);
  }

  return getDaftarPenggunaPending(userKey);
}

function tolakPengguna(rowNum, userKey) {
  if (!cekApakahAdmin(userKey)) return 'Ditolak';
  var sheet = getOrCreatePenggunaSheet_();
  if (rowNum < 2 || rowNum > sheet.getLastRow()) return 'Gagal';

  var emailPendaftar = String(sheet.getRange(rowNum, 5).getValue() || '').trim();
  var namaPendaftar = String(sheet.getRange(rowNum, 1).getValue() || '').trim();

  sheet.deleteRow(rowNum);

  if (emailPendaftar) {
    var subject = '[Catatan Keuangan] Status Pendaftaran Akun';
    var body = 'Halo ' + namaPendaftar + ',\n\n' +
      'Mohon maaf, pendaftaran akun Anda di Catatan Keuangan belum dapat disetujui oleh Admin.\n\n' +
      'Terima kasih.';
    var htmlBody = '<div style="font-family: Arial, sans-serif; padding: 20px; background-color: #121212; color: #e0e0e0;">' +
      '<div style="max-width: 500px; margin: 0 auto; background: #1e1e1e; padding: 20px; border-radius: 12px; border: 1px solid #333;">' +
      '<h2 style="color: #f44336; margin-top: 0;">Status Pendaftaran Akun</h2>' +
      '<p>Halo <b>' + namaPendaftar + '</b>,</p>' +
      '<p>Mohon maaf, pendaftaran akun Anda di Catatan Keuangan belum dapat disetujui oleh Admin saat ini.</p>' +
      '</div></div>';
    kirimEmailAmanHtml_(emailPendaftar, subject, body, htmlBody);
  }

  return getDaftarPenggunaPending(userKey);
}

function tambahPengguna(nama, kode, userKey) {
  if (!cekApakahAdmin(userKey)) return 'Ditolak';
  if (!nama || !kode) return 'Gagal';
  var sheet = getOrCreatePenggunaSheet_();
  var tz = 'Asia/Jakarta';
  try { tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(); } catch (e) {}
  var nowStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  sheet.appendRow([String(nama).trim(), String(kode).trim(), '', 'APPROVED', '', nowStr]);
  return 'Sukses';
}

function hapusPengguna(rowNum, userKey) {
  if (!cekApakahAdmin(userKey)) return 'Ditolak';
  return hapusPenggunaLengkap(rowNum);
}

function hapusPenggunaLengkap(rowNum) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPengguna = getOrCreatePenggunaSheet_();

  var rowData = sheetPengguna.getRange(rowNum, 1, 1, 2).getValues()[0];
  var nama = String(rowData[0] || '').trim();
  var kode = String(rowData[1] || '').trim();

  if (!nama || !kode) return getDaftarPenggunaInternal_();

  var userKeyTarget = buatUserKey_(nama, kode);

  hapusDataByUserKey_(ss.getSheetByName('Transaksi'), 9, userKeyTarget);
  hapusDataByUserKey_(ss.getSheetByName('Rekening'),  4, userKeyTarget);
  hapusDataByUserKey_(ss.getSheetByName('Kategori'),  6, userKeyTarget);
  hapusDataByUserKey_(ss.getSheetByName('Transfer'),  6, userKeyTarget);
  hapusDataByUserKey_(ss.getSheetByName('Goals'),     9, userKeyTarget);
  hapusDataByUserKey_(ss.getSheetByName('HutangPiutang'), 9, userKeyTarget);
  hapusDataByUserKey_(ss.getSheetByName('Init Akun'), 1, userKeyTarget);

  sheetPengguna.deleteRow(rowNum);

  Logger.log('Pengguna "' + nama + '" dan semua datanya telah dihapus. UserKey: ' + userKeyTarget);
  return getDaftarPenggunaInternal_();
}

function getDaftarPenggunaInternal_() {
  var sheet = getOrCreatePenggunaSheet_();
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    var nama = String(data[i][0] || '').trim();
    var status = String(data[i][3] || 'APPROVED').trim().toUpperCase();
    if (nama && status !== 'PENDING') {
      list.push({
        row: i + 1,
        nama: nama,
        kode: String(data[i][1] || '').trim(),
        admin: String(data[i][2] || '').trim().toUpperCase() === 'TRUE',
        status: status,
        email: String(data[i][4] || '').trim()
      });
    }
  }
  return list;
}

function hapusDataByUserKey_(sheet, kolomOwner, userKey) {
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  for (var r = lastRow; r >= 2; r--) {
    var val = String(sheet.getRange(r, kolomOwner).getValue() || '').trim();
    if (val === userKey) {
      sheet.deleteRow(r);
    }
  }
}
