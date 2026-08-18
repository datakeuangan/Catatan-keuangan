/**
 * ============================================================
 * NOTIFIKASI CICILAN - V1
 * ============================================================
 * Gmail otomatis + data pengaturan notifikasi.
 *
 * Catatan:
 * - Modul ini tidak mengirim WhatsApp.
 * - Pengiriman dilakukan menggunakan MailApp Apps Script.
 * - Trigger harian dibuat melalui fungsi setupTriggerNotifikasiCicilan().
 * - Pengaturan disimpan di sheet NotifikasiCicilan.
 * ============================================================
 */

const NOTIF_CICILAN_CONFIG = {
  SHEET: 'NotifikasiCicilan',
  HEADER: [
    'ID',
    'UserKey',
    'Email',
    'Aktif',
    'H7',
    'H3',
    'H1',
    'HariH',
    'SetelahJatuhTempo',
    'MaksHariTerlambat',
    'JamPengecekan',
    'TerakhirDijalankan',
    'UpdatedAt'
  ],
  DEFAULT_MAX_LATE_DAYS: 7
};

/**
 * Log pengiriman dipisahkan dari pengaturan agar satu pengingat
 * tidak dikirim berulang-ulang untuk cicilan dan jenis yang sama.
 */

/**
 * Konfigurasi notifikasi PER ID CICILAN.
 * Sistem lama (NotifikasiCicilan) tetap dipertahankan sebagai fallback
 * untuk data lama yang belum memiliki konfigurasi per-ID.
 */
const NOTIF_CICILAN_PER_ID_CONFIG = {
  SHEET: 'NotifikasiCicilanPerID',
  HEADER: [
    'ID',
    'UserKey',
    'CicilanID',
    'Email',
    'Aktif',
    'H7',
    'H3',
    'H1',
    'HariH',
    'SetelahJatuhTempo',
    'MaksHariTerlambat',
    'JamPengecekan',
    'TerakhirDijalankan',
    'UpdatedAt'
  ],
  DEFAULT_MAX_LATE_DAYS: 7
};

const NOTIF_CICILAN_LOG_CONFIG = {
  SHEET: 'NotifikasiCicilanLog',
  HEADER: [
    'ID',
    'UserKey',
    'CicilanID',
    'Jenis',
    'TanggalJatuhTempo',
    'TanggalKunci',
    'DikirimPada',
    'Email'
  ]
};

function setupNotifikasiCicilan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(NOTIF_CICILAN_CONFIG.SHEET);

  if (!sh) sh = ss.insertSheet(NOTIF_CICILAN_CONFIG.SHEET);

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, NOTIF_CICILAN_CONFIG.HEADER.length)
      .setValues([NOTIF_CICILAN_CONFIG.HEADER]);
    sh.setFrozenRows(1);
  } else {
    ensureNotificationHeaders_(sh);
  }

  return {
    success: true,
    message: 'Sheet NotifikasiCicilan siap digunakan.'
  };
}


function setupNotifikasiCicilanPerID_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(NOTIF_CICILAN_PER_ID_CONFIG.SHEET);

  if (!sh) sh = ss.insertSheet(NOTIF_CICILAN_PER_ID_CONFIG.SHEET);

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, NOTIF_CICILAN_PER_ID_CONFIG.HEADER.length)
      .setValues([NOTIF_CICILAN_PER_ID_CONFIG.HEADER]);
    sh.setFrozenRows(1);
  } else {
    const headers = NOTIF_CICILAN_PER_ID_CONFIG.HEADER;
    const current = sh.getRange(
      1, 1, 1, Math.max(sh.getLastColumn(), headers.length)
    ).getValues()[0];

    headers.forEach(function(h, i) {
      if (current[i] !== h) sh.getRange(1, i + 1).setValue(h);
    });
    sh.setFrozenRows(1);
  }

  return sh;
}

function findNotificationRowPerID_(sh, userKey, cicilanID) {
  if (!sh || sh.getLastRow() < 2) return -1;

  const requestedUser = normalisasiUserKeyNotifikasi_(userKey);
  const requestedCicilan = String(cicilanID || '').trim();
  if (!requestedUser || !requestedCicilan) return -1;

  const rows = sh.getRange(
    2, 1, sh.getLastRow() - 1,
    Math.max(sh.getLastColumn(), NOTIF_CICILAN_PER_ID_CONFIG.HEADER.length)
  ).getValues();

  for (let i = 0; i < rows.length; i++) {
    const storedUser = normalisasiUserKeyNotifikasi_(rows[i][1]);
    const storedCicilan = String(rows[i][2] == null ? '' : rows[i][2]).trim();

    if (storedUser === requestedUser && storedCicilan === requestedCicilan) {
      return i + 2;
    }
  }

  return -1;
}

function simpanPengaturanNotifikasiCicilanPerID(data) {
  const sh = setupNotifikasiCicilanPerID_();

  data = data || {};
  const userKey = String(data.userKey || '').trim();
  const cicilanID = String(data.cicilanID || '').trim();
  const email = String(data.email || '').trim();

  if (!userKey) throw new Error('UserKey wajib diisi.');
  if (!cicilanID) throw new Error('CicilanID wajib diisi.');
  if (!email || !isValidEmail_(email)) {
    throw new Error('Email pengingat tidak valid.');
  }

  // Pastikan ID benar-benar milik user yang sedang menyimpan.
  const cicilan = typeof ambilCicilanById === 'function'
    ? ambilCicilanById(cicilanID, userKey)
    : null;

  if (!cicilan) {
    throw new Error('CicilanID tidak ditemukan atau bukan milik akun ini.');
  }

  const jamPengecekan = normalisasiJamPengecekan_(data.jamPengecekan || '07:00');
  const rowIndex = findNotificationRowPerID_(sh, userKey, cicilanID);
  const now = new Date();

  const row = [
    rowIndex > 1 ? sh.getRange(rowIndex, 1).getValue() :
      'NOTIF-ID-' + Utilities.getUuid().replace(/-/g, '').slice(0, 16),
    userKey,
    cicilanID,
    email,
    data.aktif !== false,
    data.h7 !== false,
    data.h3 !== false,
    data.h1 !== false,
    data.hariH !== false,
    data.setelahJatuhTempo !== false,
    Math.max(
      0,
      parseInt(data.maksHariTerlambat, 10) ||
        NOTIF_CICILAN_PER_ID_CONFIG.DEFAULT_MAX_LATE_DAYS
    ),
    jamPengecekan,
    rowIndex > 1 ? sh.getRange(rowIndex, 13).getValue() : '',
    now
  ];

  if (rowIndex > 1) {
    sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    sh.getRange(rowIndex, 12).setNumberFormat('@').setValue(jamPengecekan);
  } else {
    sh.appendRow(row);
    const newRow = sh.getLastRow();
    sh.getRange(newRow, 12).setNumberFormat('@').setValue(jamPengecekan);
  }

  SpreadsheetApp.flush();

  return {
    success: true,
    id: row[0],
    userKey: userKey,
    cicilanID: cicilanID,
    email: email,
    aktif: data.aktif !== false,
    jamPengecekan: jamPengecekan,
    message: 'Pengaturan notifikasi Cicilan berhasil terhubung ke ID transaksi.'
  };
}

function ambilPengaturanNotifikasiCicilanPerID(cicilanID, userKey) {
  const sh = setupNotifikasiCicilanPerID_();
  const rowIndex = findNotificationRowPerID_(sh, userKey, cicilanID);

  if (rowIndex < 2) {
    return {
      found: false,
      cicilanID: String(cicilanID || '').trim(),
      userKey: String(userKey || '').trim(),
      email: '',
      aktif: true,
      h7: true,
      h3: true,
      h1: true,
      hariH: true,
      setelahJatuhTempo: true,
      maksHariTerlambat: NOTIF_CICILAN_PER_ID_CONFIG.DEFAULT_MAX_LATE_DAYS,
      jamPengecekan: '07:00'
    };
  }

  const r = sh.getRange(
    rowIndex, 1, 1, NOTIF_CICILAN_PER_ID_CONFIG.HEADER.length
  ).getValues()[0];

  return {
    found: true,
    id: r[0],
    userKey: String(r[1] || '').trim(),
    cicilanID: String(r[2] || '').trim(),
    email: String(r[3] || '').trim(),
    aktif: r[4] !== false,
    h7: r[5] !== false,
    h3: r[6] !== false,
    h1: r[7] !== false,
    hariH: r[8] !== false,
    setelahJatuhTempo: r[9] !== false,
    maksHariTerlambat: Number(r[10]) || 7,
    jamPengecekan: normalisasiJamPengecekan_(r[11]),
    terakhirDijalankan: r[12]
  };
}

function simpanPengaturanNotifikasiCicilan(data) {
  setupNotifikasiCicilan();

  data = data || {};
  const userKey = String(data.userKey || '').trim();
  if (!userKey) throw new Error('UserKey wajib diisi.');

  const email = String(data.email || '').trim();
  if (!email || !isValidEmail_(email)) {
    throw new Error('Email pengingat tidak valid.');
  }

  const sh = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(NOTIF_CICILAN_CONFIG.SHEET);

  const rowIndex = findNotificationRow_(sh, userKey);
  const now = new Date();

  // JamPengecekan sengaja disimpan sebagai TEKS "HH:mm".
  // Ini mencegah Google Sheets mengubahnya menjadi Date object
  // dengan basis tanggal 1899 dan interpretasi timezone.
  const jamPengecekan = normalisasiJamPengecekan_(data.jamPengecekan || '07:00');

  const row = [
    rowIndex > 1 ? sh.getRange(rowIndex, 1).getValue() :
      'NOTIF-' + Utilities.getUuid().replace(/-/g, '').slice(0, 16),
    userKey,
    email,
    data.aktif !== false,
    data.h7 !== false,
    data.h3 !== false,
    data.h1 !== false,
    data.hariH !== false,
    data.setelahJatuhTempo !== false,
    Math.max(
      0,
      parseInt(data.maksHariTerlambat, 10) ||
        NOTIF_CICILAN_CONFIG.DEFAULT_MAX_LATE_DAYS
    ),
    jamPengecekan,
    rowIndex > 1 ? sh.getRange(rowIndex, 12).getValue() : '',
    now
  ];

  if (rowIndex > 1) {
    sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);

    // Paksa kolom JamPengecekan menjadi plain text dan tulis ulang
    // nilai HH:mm agar nilai lama Date tidak terbawa.
    sh.getRange(rowIndex, 11)
      .setNumberFormat('@')
      .setValue(jamPengecekan);
  } else {
    sh.appendRow(row);

    // appendRow dapat mengikuti format kolom lama; paksa cell baru
    // menjadi teks juga.
    const newRow = sh.getLastRow();
    sh.getRange(newRow, 11)
      .setNumberFormat('@')
      .setValue(jamPengecekan);
  }

  SpreadsheetApp.flush();

  return {
    success: true,
    message: 'Pengaturan notifikasi berhasil disimpan.',
    userKey: userKey,
    jamPengecekan: jamPengecekan
  };
}


function ambilPengaturanNotifikasiCicilan(userKey) {
  setupNotifikasiCicilan();

  const sh = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(NOTIF_CICILAN_CONFIG.SHEET);

  const rowIndex = findNotificationRow_(sh, userKey);

  if (rowIndex < 2) {
    return {
      found: false,
      userKey: String(userKey || '').trim(),
      email: '',
      aktif: true,
      h7: true,
      h3: true,
      h1: true,
      hariH: true,
      setelahJatuhTempo: true,
      maksHariTerlambat: NOTIF_CICILAN_CONFIG.DEFAULT_MAX_LATE_DAYS,
      jamPengecekan: '07:00'
    };
  }

  const r = sh.getRange(rowIndex, 1, 1, NOTIF_CICILAN_CONFIG.HEADER.length).getValues()[0];

  return {
    found: true,
    id: r[0],
    userKey: String(r[1] == null ? '' : r[1]).trim(),
    email: String(r[2] == null ? '' : r[2]).trim(),
    aktif: r[3] !== false,
    h7: r[4] !== false,
    h3: r[5] !== false,
    h1: r[6] !== false,
    hariH: r[7] !== false,
    setelahJatuhTempo: r[8] !== false,
    maksHariTerlambat: Number(r[9]) || 7,
    jamPengecekan: normalisasiJamPengecekan_(r[10]),
    terakhirDijalankan: r[11]
  };
}

/**
 * Fungsi utama untuk trigger harian.
 *
 * Karena fungsi ini berjalan tanpa membuka Form.html, fungsi akan
 * memeriksa seluruh pengaturan notifikasi dan seluruh cicilan.
 */

function setupLogNotifikasiCicilan_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(NOTIF_CICILAN_LOG_CONFIG.SHEET);

  if (!sh) sh = ss.insertSheet(NOTIF_CICILAN_LOG_CONFIG.SHEET);

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, NOTIF_CICILAN_LOG_CONFIG.HEADER.length)
      .setValues([NOTIF_CICILAN_LOG_CONFIG.HEADER]);
    sh.setFrozenRows(1);
  } else {
    const headers = NOTIF_CICILAN_LOG_CONFIG.HEADER;
    const current = sh.getRange(
      1, 1, 1, Math.max(sh.getLastColumn(), headers.length)
    ).getValues()[0];

    headers.forEach(function(h, i) {
      if (current[i] !== h) sh.getRange(1, i + 1).setValue(h);
    });
    sh.setFrozenRows(1);
  }

  return sh;
}

function buildNotificationKey_(item, jenis) {
  return normalizeNotificationKeyParts_(
    item.id,
    jenis,
    item.tanggalJatuhTempo
  );
}

function normalizeNotificationDueKey_(value) {
  if (value === null || value === undefined || value === '') return '-';

  const tz = Session.getScriptTimeZone() || 'Asia/Jakarta';

  if (Object.prototype.toString.call(value) === '[object Date]' &&
      !isNaN(value.getTime())) {
    return Utilities.formatDate(value, tz, 'yyyy-MM-dd');
  }

  const text = String(value).trim();

  // Jika sudah yyyy-MM-dd, pertahankan apa adanya.
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;

  // Jika berupa tanggal/waktu yang bisa diparse, normalkan ke tanggal lokal.
  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, tz, 'yyyy-MM-dd');
  }

  return text;
}

function normalizeNotificationKeyParts_(cicilanId, jenis, tanggalKunci) {
  return [
    String(cicilanId || '').trim(),
    String(jenis || '').trim(),
    normalizeNotificationDueKey_(tanggalKunci)
  ].join('|');
}


const NOTIF_CICILAN_PROPERTY_PREFIX = 'NOTIF_SENT_V1|';

function notificationPropertyKey_(key) {
  return NOTIF_CICILAN_PROPERTY_PREFIX + String(key || '').trim();
}

function getSentNotificationProperties_() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const keys = {};

  Object.keys(props).forEach(function(propKey) {
    if (propKey.indexOf(NOTIF_CICILAN_PROPERTY_PREFIX) !== 0) return;

    const key = propKey.slice(NOTIF_CICILAN_PROPERTY_PREFIX.length);
    if (key) keys[key] = true;
  });

  return keys;
}

function isNotificationAlreadySent_(key) {
  if (!key) return false;

  const propKey = notificationPropertyKey_(key);
  return PropertiesService.getScriptProperties().getProperty(propKey) === '1';
}

function markNotificationPropertySent_(key) {
  if (!key) return;

  PropertiesService.getScriptProperties().setProperty(
    notificationPropertyKey_(key),
    '1'
  );
}

function removeNotificationProperty_(key) {
  if (!key) return;

  PropertiesService.getScriptProperties().deleteProperty(
    notificationPropertyKey_(key)
  );
}

function getSentNotificationKeys_() {
  const sh = setupLogNotifikasiCicilan_();
  const keys = getSentNotificationProperties_();

  if (sh.getLastRow() < 2) return keys;

  const rows = sh.getRange(
    2, 1, sh.getLastRow() - 1, NOTIF_CICILAN_LOG_CONFIG.HEADER.length
  ).getValues();

  rows.forEach(function(r) {
    const cicilanId = String(r[2] || '').trim();
    const jenis = String(r[3] || '').trim();
    const tanggalKunci = normalizeNotificationDueKey_(r[5]);

    if (cicilanId && jenis && tanggalKunci !== '-') {
      const key = normalizeNotificationKeyParts_(
        cicilanId,
        jenis,
        tanggalKunci
      );
      keys[key] = true;
    }
  });

  return keys;
}

function markNotificationSent_(item, jenis, email) {
  const sh = setupLogNotifikasiCicilan_();
  const tz = Session.getScriptTimeZone() || 'Asia/Jakarta';

  const dueKey = item.tanggalJatuhTempo
    ? Utilities.formatDate(
        new Date(item.tanggalJatuhTempo),
        tz,
        'yyyy-MM-dd'
      )
    : '-';

  sh.appendRow([
    'NLOG-' + Utilities.getUuid().replace(/-/g, '').slice(0, 16),
    String(item.userKey || ''),
    String(item.id || ''),
    String(jenis || ''),
    item.tanggalJatuhTempo || '',
    dueKey,
    new Date(),
    String(email || '')
  ]);
  SpreadsheetApp.flush();
}


function prosesNotifikasiCicilanPerIDHarian_() {
  const sh = setupNotifikasiCicilanPerID_();
  const sentKeys = getSentNotificationKeys_();
  const rows = sh.getLastRow() >= 2
    ? sh.getRange(2, 1, sh.getLastRow() - 1, NOTIF_CICILAN_PER_ID_CONFIG.HEADER.length).getValues()
    : [];

  const result = {
    checked: 0,
    sent: 0,
    skipped: 0,
    duplicateSkipped: 0,
    errors: []
  };

  rows.forEach(function(r, index) {
    const cfg = {
      rowIndex: index + 2,
      id: String(r[0] || '').trim(),
      userKey: String(r[1] || '').trim(),
      cicilanID: String(r[2] || '').trim(),
      email: String(r[3] || '').trim(),
      aktif: r[4] !== false,
      h7: r[5] !== false,
      h3: r[6] !== false,
      h1: r[7] !== false,
      hariH: r[8] !== false,
      setelahJatuhTempo: r[9] !== false,
      maksHariTerlambat: Number(r[10]) || 7,
      jamPengecekan: normalisasiJamPengecekan_(r[11])
    };

    try {
      if (!cfg.userKey || !cfg.cicilanID || !cfg.aktif || !cfg.email || !isValidEmail_(cfg.email)) {
        result.skipped++;
        return;
      }

      result.checked++;

      const item = typeof ambilCicilanById === 'function'
        ? ambilCicilanById(cfg.cicilanID, cfg.userKey)
        : null;

      if (!item || item.status === 'Selesai' || Number(item.sisa) <= 0) {
        result.skipped++;
        return;
      }

      const statusList = cekStatusCicilan(cfg.userKey) || [];
      const statusItem = statusList.find(function(x) {
        return String(x.id || '').trim() === cfg.cicilanID;
      });

      if (!statusItem || statusItem.statusWaktu === 'Selesai') {
        result.skipped++;
        return;
      }

      const jenis = tentukanJenisPengingat_(statusItem, cfg);
      if (!jenis) {
        result.skipped++;
        return;
      }

      const key = buildNotificationKey_(statusItem, jenis);
      if (isNotificationAlreadySent_(key) || sentKeys[key]) {
        result.duplicateSkipped++;
        return;
      }

      const subject = buatSubjectPengingat_(statusItem, jenis);
      const bodyHtml = buatEmailPengingat_(statusItem, jenis);

      MailApp.sendEmail({
        to: cfg.email,
        subject: subject,
        htmlBody: bodyHtml,
        body: stripHtml_(bodyHtml),
        name: 'Pengingat Keuangan'
      });

      markNotificationPropertySent_(key);
      markNotificationSent_(statusItem, jenis, cfg.email);
      sentKeys[key] = true;
      result.sent++;

      sh.getRange(cfg.rowIndex, 13).setValue(new Date());
    } catch (err) {
      result.errors.push({
        cicilanID: cfg.cicilanID,
        userKey: cfg.userKey,
        error: String(err && err.message || err)
      });
    }
  });

  SpreadsheetApp.flush();
  return result;
}

function prosesNotifikasiCicilanHarian() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return {
      checkedUsers: 0,
      sent: 0,
      skipped: 0,
      duplicateSkipped: 0,
      errors: ['Notifikasi sedang diproses oleh eksekusi lain.']
    };
  }

  try {
    const perIdResult = prosesNotifikasiCicilanPerIDHarian_();

    setupModulCicilan();
  setupNotifikasiCicilan();
  setupLogNotifikasiCicilan_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(NOTIF_CICILAN_CONFIG.SHEET);
  const configs = getNotificationRows_(sh);
  const perIdSh = setupNotifikasiCicilanPerID_();
  const perIdRows = perIdSh.getLastRow() >= 2
    ? perIdSh.getRange(2, 1, perIdSh.getLastRow() - 1, NOTIF_CICILAN_PER_ID_CONFIG.HEADER.length).getValues()
    : [];
  const perIdKeys = {};
  perIdRows.forEach(function(r) {
    const u = String(r[1] || '').trim();
    const c = String(r[2] || '').trim();
    if (u && c) perIdKeys[normalisasiUserKeyNotifikasi_(u) + '|' + c] = true;
  });
  const sentKeys = getSentNotificationKeys_();

  const result = {
    checkedUsers: 0,
    sent: 0,
    skipped: 0,
    duplicateSkipped: 0,
    errors: []
  };

  configs.forEach(function(cfg) {
    try {
      if (!cfg.aktif || !cfg.email || !isValidEmail_(cfg.email)) {
        result.skipped++;
        return;
      }

      result.checkedUsers++;

      const statuses = cekStatusCicilan(cfg.userKey);

      statuses.forEach(function(item) {
        if (item.statusWaktu === 'Selesai') return;

        const perIdKey = normalisasiUserKeyNotifikasi_(cfg.userKey) + '|' + String(item.id || '').trim();
        // Jika transaksi sudah memiliki konfigurasi per-ID, konfigurasi
        // tersebut menjadi sumber kebenaran dan fallback user-level dilewati.
        if (perIdKeys[perIdKey]) return;

        const jenis = tentukanJenisPengingat_(item, cfg);
        if (!jenis) return;

        const key = buildNotificationKey_(item, jenis);

        if (isNotificationAlreadySent_(key) || sentKeys[key]) {
          result.duplicateSkipped++;
          return;
        }

        const subject = buatSubjectPengingat_(item, jenis);
        const bodyHtml = buatEmailPengingat_(item, jenis);

        MailApp.sendEmail({
          to: cfg.email,
          subject: subject,
          htmlBody: bodyHtml,
          body: stripHtml_(bodyHtml),
          name: 'Pengingat Keuangan'
        });

        // Tandai deduplikasi terlebih dahulu pada store persisten,
        // lalu simpan audit trail di Spreadsheet.
        markNotificationPropertySent_(key);
        markNotificationSent_(item, jenis, cfg.email);
        sentKeys[key] = true;
        result.sent++;
      });

      updateLastRun_(sh, cfg.rowIndex);
    } catch (err) {
      result.errors.push({
        userKey: cfg.userKey,
        error: String(err && err.message || err)
      });
    }
  });

    result.checkedUsers += Number(perIdResult.checked || 0);
    result.sent += Number(perIdResult.sent || 0);
    result.skipped += Number(perIdResult.skipped || 0);
    result.duplicateSkipped += Number(perIdResult.duplicateSkipped || 0);
    if (Array.isArray(perIdResult.errors)) {
      result.errors = result.errors.concat(perIdResult.errors);
    }

    return result;
  } finally {
    lock.releaseLock();
  }
}

function tentukanJenisPengingat_(item, cfg) {
  if (item.statusWaktu === 'Terlambat') {
    if (!cfg.setelahJatuhTempo) return '';
    if (item.selisihHari > cfg.maksHariTerlambat) return '';
    return 'TERLAMBAT';
  }

  if (item.statusWaktu === 'JatuhTempoHariIni') {
    return cfg.hariH ? 'HARI_H' : '';
  }

  if (item.statusWaktu !== 'AkanJatuhTempo') return '';

  if (item.selisihHari === 7 && cfg.h7) return 'H7';
  if (item.selisihHari === 3 && cfg.h3) return 'H3';
  if (item.selisihHari === 1 && cfg.h1) return 'H1';

  return '';
}

function buatSubjectPengingat_(item, jenis) {
  if (jenis === 'TERLAMBAT') {
    return '🚨 Cicilan terlambat: ' + (item.nama || 'Cicilan');
  }

  if (jenis === 'HARI_H') {
    return '🔴 Jatuh tempo hari ini: ' + (item.nama || 'Cicilan');
  }

  return '🔔 Pengingat cicilan: ' + (item.nama || 'Cicilan');
}

function buatEmailPengingat_(item, jenis) {
  const tz = Session.getScriptTimeZone() || 'Asia/Jakarta';
  const tanggal = item.tanggalJatuhTempo
    ? Utilities.formatDate(new Date(item.tanggalJatuhTempo), tz, 'dd MMMM yyyy')
    : '-';

  let title = 'Pengingat Cicilan';
  let intro = 'Cicilan berikut perlu diperhatikan.';

  if (jenis === 'H7' || jenis === 'H3' || jenis === 'H1') {
    title = '🔔 Pengingat Cicilan';
    intro = 'Cicilan <b>' + escapeHtml_(item.nama || 'Cicilan') +
      '</b> akan jatuh tempo dalam <b>' + item.selisihHari + ' hari</b>.';
  } else if (jenis === 'HARI_H') {
    title = '🔴 Jatuh Tempo Hari Ini';
    intro = 'Cicilan <b>' + escapeHtml_(item.nama || 'Cicilan') +
      '</b> jatuh tempo <b>hari ini</b>.';
  } else if (jenis === 'TERLAMBAT') {
    title = '🚨 Pembayaran Belum Tercatat';
    intro = 'Cicilan <b>' + escapeHtml_(item.nama || 'Cicilan') +
      '</b> sudah terlambat <b>' + item.selisihHari + ' hari</b>.';
  }

  return `
    <div style="font-family:Arial,sans-serif;max-width:560px;margin:auto;color:#222;">
      <div style="background:#9c27b0;color:#fff;padding:20px;border-radius:14px 14px 0 0;">
        <div style="font-size:20px;font-weight:700;">${title}</div>
      </div>
      <div style="padding:22px;border:1px solid #eee;border-top:0;border-radius:0 0 14px 14px;">
        <p style="font-size:14px;line-height:1.6;">${intro}</p>
        <table style="width:100%;border-collapse:collapse;font-size:13px;">
          <tr><td style="padding:7px 0;color:#777;">Nama</td><td style="padding:7px 0;font-weight:700;">${escapeHtml_(item.nama || '-')}</td></tr>
          <tr><td style="padding:7px 0;color:#777;">Nominal cicilan</td><td style="padding:7px 0;font-weight:700;">${formatRupiah_(item.nominalCicilan)}</td></tr>
          <tr><td style="padding:7px 0;color:#777;">Jatuh tempo</td><td style="padding:7px 0;font-weight:700;">${tanggal}</td></tr>
          <tr><td style="padding:7px 0;color:#777;">Sisa</td><td style="padding:7px 0;font-weight:700;">${formatRupiah_(item.sisa)}</td></tr>
        </table>
        <p style="margin-top:20px;color:#777;font-size:12px;">
          Pesan ini dikirim otomatis oleh aplikasi keuangan kamu.
        </p>
      </div>
    </div>
  `;
}

/**
 * Kirim email percobaan secara manual dari aplikasi.
 */
function setupTriggerNotifikasiCicilan() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'prosesNotifikasiCicilanHarian') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('prosesNotifikasiCicilanHarian')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();

  return {
    success: true,
    message: 'Trigger harian notifikasi cicilan berhasil dibuat.'
  };
}

function getNotificationRows_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];

  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn())
    .getValues()
    .map(function(r, i) {
      return {
        id: r[0],
        userKey: String(r[1] || ''),
        email: String(r[2] || '').trim(),
        aktif: r[3] !== false,
        h7: r[4] !== false,
        h3: r[5] !== false,
        h1: r[6] !== false,
        hariH: r[7] !== false,
        setelahJatuhTempo: r[8] !== false,
        maksHariTerlambat: Number(r[9]) || 7,
        jamPengecekan: normalisasiJamPengecekan_(r[10]),
        terakhirDijalankan: r[11],
        rowIndex: i + 2
      };
    });
}

function normalisasiUserKeyNotifikasi_(value) {
  return String(value == null ? '' : value)
    .normalize('NFKC')
    .replace(/[\u00A0\u200B-\u200D\uFEFF]/g, '')
    .trim()
    .replace(/\s+/g, ' ');
}

function findNotificationRow_(sh, userKey) {
  if (!sh || sh.getLastRow() < 2) return -1;

  const requested = normalisasiUserKeyNotifikasi_(userKey);
  if (!requested) return -1;

  const requestedLower = requested.toLowerCase();
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();

  // 1) Pencocokan persis setelah normalisasi.
  for (let i = 0; i < rows.length; i++) {
    const stored = normalisasiUserKeyNotifikasi_(rows[i][1]);
    if (stored && stored === requested) return i + 2;
  }

  // 2) Fallback aman untuk perbedaan kapitalisasi.
  for (let i = 0; i < rows.length; i++) {
    const stored = normalisasiUserKeyNotifikasi_(rows[i][1]);
    if (stored && stored.toLowerCase() === requestedLower) return i + 2;
  }

  return -1;
}

function updateLastRun_(sh, rowIndex) {
  sh.getRange(rowIndex, 12).setValue(new Date());
}

function ensureNotificationHeaders_(sh) {
  const headers = NOTIF_CICILAN_CONFIG.HEADER;
  const current = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), headers.length)).getValues()[0];

  headers.forEach(function(h, i) {
    if (current[i] !== h) sh.getRange(1, i + 1).setValue(h);
  });

  sh.setFrozenRows(1);
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function formatRupiah_(n) {
  n = Number(n) || 0;
  return 'Rp' + Math.round(n).toLocaleString('id-ID');
}

function escapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function stripHtml_(html) {
  return String(html || '')
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}


/**
 * ============================================================
 * TES NOTIFIKASI CICILAN - READ ONLY / TANPA KIRIM EMAIL
 * ============================================================
 * Menguji logika jenis pengingat:
 * H7, H3, H1, HARI_H, TERLAMBAT.
 *
 * Tes ini TIDAK:
 * - mengirim email
 * - membuat trigger
 * - mengubah konfigurasi
 * - mengubah data Cicilan
 *
 * Tes hanya memanggil fungsi logika dan membaca data.
 * ============================================================
 */
function normalisasiJamPengecekan_(value) {
  if (value === null || value === undefined || value === '') return '07:00';

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    const tz = Session.getScriptTimeZone() || 'Asia/Jakarta';
    return Utilities.formatDate(value, tz, 'HH:mm');
  }

  if (typeof value === 'number' && isFinite(value)) {
    const fraction = ((value % 1) + 1) % 1;
    const totalMinutes = Math.round(fraction * 24 * 60) % (24 * 60);
    return String(Math.floor(totalMinutes / 60)).padStart(2, '0') + ':' +
           String(totalMinutes % 60).padStart(2, '0');
  }

  const text = String(value).trim();
  const match = text.match(/^(\d{1,2}):(\d{1,2})$/);
  if (match) {
    const hh = Number(match[1]), mm = Number(match[2]);
    if (hh >= 0 && hh <= 23 && mm >= 0 && mm <= 59) {
      return String(hh).padStart(2, '0') + ':' + String(mm).padStart(2, '0');
    }
  }

  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    const tz = Session.getScriptTimeZone() || 'Asia/Jakarta';
    return Utilities.formatDate(parsed, tz, 'HH:mm');
  }

  return '07:00';
}


/**
 * ============================================================
 * SET JAM NOTIFIKASI PRODUKSI KE 07:00
 * ============================================================
 * Mengubah HANYA konfigurasi jam pengecekan untuk user produksi.
 *
 * TIDAK mengubah cicilan.
 * TIDAK mengubah pembayaran.
 * TIDAK mengirim email.
 * TIDAK membuat / menghapus trigger.
 *
 * User:
 *   putu::1212
 *
 * Jam target:
 *   07:00
 *
 * Pengaturan notifikasi lain dipertahankan sesuai konfigurasi
 * produksi yang sudah digunakan sebelumnya.
 * ============================================================
 */

/**
 * ============================================================
 * DIAGNOSTIK RAW KONFIGURASI JAM PRODUKSI
 * ============================================================
 * READ ONLY.
 * Tidak mengirim email.
 * Tidak membuat / menghapus trigger.
 * Tidak mengubah cicilan, pembayaran, atau konfigurasi.
 *
 * Tujuan:
 * mencari tahu mengapa nilai yang disimpan sebagai 07:00
 * terbaca kembali sebagai 07:11.
 * ============================================================
 */

/**
 * ============================================================
 * VERIFIKASI FINAL PENYIMPANAN JAM PRODUKSI
 * ============================================================
 * Menulis 07:00 melalui fungsi produksi lalu membaca kembali
 * nilai mentah dan display cell.
 *
 * Tidak mengirim email.
 * Tidak membuat / menghapus trigger.
 * Tidak mengubah cicilan atau pembayaran.
 * ============================================================
 */
function setJamNotifikasiProduksi0700() {
  const userKey = 'putu::1212';

  const activeEmail = String(Session.getActiveUser().getEmail() || '').trim();
  const effectiveEmail = String(Session.getEffectiveUser().getEmail() || '').trim();
  const emailAkun = activeEmail || effectiveEmail;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NOTIF_CICILAN_CONFIG.SHEET);

  if (!sheet) {
    throw new Error(
      'Sheet konfigurasi notifikasi "' +
      NOTIF_CICILAN_CONFIG.SHEET +
      '" tidak ditemukan.'
    );
  }

  const rows = getNotificationRows_(sheet);
  const existing = rows.find(function(row) {
    return row.userKey === userKey;
  });

  if (!existing) {
    throw new Error(
      'Konfigurasi produksi untuk userKey "' +
      userKey +
      '" tidak ditemukan. ' +
      'Jalankan setupKonfigurasiProduksiNotifikasiCicilan terlebih dahulu.'
    );
  }

  const email = existing.email || emailAkun;

  if (!email || !isValidEmail_(email)) {
    throw new Error(
      'Email konfigurasi produksi tidak valid: ' + String(email || '(kosong)')
    );
  }

  const hasil = simpanPengaturanNotifikasiCicilan({
    userKey: userKey,
    email: email,
    aktif: existing.aktif !== false,
    h7: existing.h7 !== false,
    h3: existing.h3 !== false,
    h1: existing.h1 !== false,
    hariH: existing.hariH !== false,
    setelahJatuhTempo: existing.setelahJatuhTempo !== false,
    maksHariTerlambat: Number(existing.maksHariTerlambat || 7),
    jamPengecekan: '07:00'
  });

  Logger.log('================ SET JAM PRODUKSI 07:00 ================');
  Logger.log(JSON.stringify({
    status: hasil && hasil.success ? 'PASS' : 'CEK MASALAH',
    userKey: userKey,
    email: email,
    jamPengecekanBaru: '07:00',
    message: hasil ? hasil.message : 'Tidak ada hasil dari fungsi penyimpanan.'
  }, null, 2));
  Logger.log('==========================================================');

  return hasil;
}

/**
 * Verifikasi khusus setelah jam produksi diset ke 07:00.
 * READ ONLY: tidak mengubah data, trigger, atau mengirim email.
 */
function setupKonfigurasiProduksiNotifikasiCicilan() {
  const userKey = 'putu::1212';

  const activeEmail = String(Session.getActiveUser().getEmail() || '').trim();
  const effectiveEmail = String(Session.getEffectiveUser().getEmail() || '').trim();
  const email = activeEmail || effectiveEmail;

  if (!email || !isValidEmail_(email)) {
    throw new Error(
      'Email akun Google tidak dapat dibaca atau tidak valid. ' +
      'Active: ' + (activeEmail || '(kosong)') +
      ', Effective: ' + (effectiveEmail || '(kosong)')
    );
  }

  const hasil = simpanPengaturanNotifikasiCicilan({
    userKey: userKey,
    email: email,
    aktif: true,
    h7: true,
    h3: true,
    h1: true,
    hariH: true,
    setelahJatuhTempo: true,
    maksHariTerlambat: 7,
    jamPengecekan: '07:00'
  });

  Logger.log('================ SETUP KONFIGURASI PRODUKSI ================');
  Logger.log(JSON.stringify({
    status: hasil.success ? 'PASS' : 'CEK MASALAH',
    userKey: userKey,
    email: email,
    aktif: true,
    h7: true,
    h3: true,
    h1: true,
    hariH: true,
    setelahJatuhTempo: true,
    maksHariTerlambat: 7,
    jamPengecekan: '07:00',
    message: hasil.message
  }, null, 2));
  Logger.log('==============================================================');

  return hasil;
}

/**
 * ============================================================
 * OTORISASI AKSES TRIGGER - TIDAK MEMBUAT TRIGGER
 * ============================================================
 * Fungsi ini hanya meminta izin ScriptApp yang diperlukan
 * untuk membaca trigger project.
 *
 * Tidak membuat trigger.
 * Tidak menghapus trigger.
 * Tidak mengirim email.
 * Tidak mengubah data cicilan.
 *
 * Jalankan fungsi ini SATU KALI jika verifikasi preflight
 * menampilkan "Specified permissions are not sufficient..."
 */
function otorisasiAksesTriggerNotifikasiCicilan() {
  const triggers = ScriptApp.getProjectTriggers();

  const hasil = {
    test: 'OTORISASI AKSES TRIGGER',
    mode: 'READ ONLY - TIDAK MEMBUAT / MENGHAPUS TRIGGER',
    aksesBerhasil: true,
    jumlahTriggerSaatIni: triggers.length,
    status: 'PASS'
  };

  Logger.log('================ OTORISASI AKSES TRIGGER ================');
  Logger.log(JSON.stringify(hasil, null, 2));
  Logger.log('===========================================================');

  return hasil;
}

/**
 * ============================================================
 * PREFLIGHT SEBELUM TRIGGER PRODUKSI
 * ============================================================
 * READ ONLY.
 * Tidak membuat trigger dan tidak mengirim email.
 *
 * Memastikan:
 * - konfigurasi notifikasi tersedia
 * - email valid
 * - user aktif terbaca
 * - cicilan aktif terbaca
 * - fungsi handler tersedia
 * - tidak ada duplikat trigger handler
 * ============================================================
 */
