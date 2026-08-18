// ============================================================
// CATATAN KEUANGAN - BACKEND MODULE
// Generated from the verified 7F-16 baseline.
// Functions are preserved in full; only file ownership was reorganized.
// ============================================================

function getOrCreatePengaturanPenggunaSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Pengaturan Pengguna');

  if (!sheet) {
    sheet = ss.insertSheet('Pengaturan Pengguna');
    sheet.getRange(1, 1, 1, 7).setValues([[
      'UserKey',
      'WidgetConfig',
      'RekeningVisibility',
      'CategoryOrderPlus',
      'CategoryOrderMinus',
      'LimitRiwayat',
      'UpdatedAt'
    ]]);
    return sheet;
  }

  // Pastikan header tersedia jika sheet pernah dibuat manual.
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 7).setValues([[
      'UserKey',
      'WidgetConfig',
      'RekeningVisibility',
      'CategoryOrderPlus',
      'CategoryOrderMinus',
      'LimitRiwayat',
      'UpdatedAt'
    ]]);
  }

  return sheet;
}

function parsePengaturanJson_(value, fallback) {
  if (value === null || value === undefined || value === '') return fallback;
  try {
    var parsed = JSON.parse(String(value));
    return parsed;
  } catch (e) {
    return fallback;
  }
}

function getPengaturanPengguna(userKey) {
  userKey = String(userKey || '').trim();
  if (!userKey) {
    return { exists: false };
  }

  var sheet = getOrCreatePengaturanPenggunaSheet_();
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return { exists: false };
  }

  var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();

  for (var i = 0; i < data.length; i++) {
    var rowUserKey = String(data[i][0] || '').trim();

    if (rowUserKey === userKey) {
      var widgetConfig = parsePengaturanJson_(data[i][1], null);
      var rekeningVisibility = parsePengaturanJson_(data[i][2], {});
      var categoryOrderPlus = parsePengaturanJson_(data[i][3], []);
      var categoryOrderMinus = parsePengaturanJson_(data[i][4], []);

      var limit = parseInt(data[i][5], 10);
      if (isNaN(limit)) limit = 5;
      limit = Math.max(1, Math.min(25, limit));

      return {
        exists: true,
        userKey: userKey,
        widgetConfig: Array.isArray(widgetConfig) ? widgetConfig : [],
        rekeningVisibility: (rekeningVisibility && typeof rekeningVisibility === 'object' && !Array.isArray(rekeningVisibility))
          ? rekeningVisibility
          : {},
        categoryOrderPlus: Array.isArray(categoryOrderPlus) ? categoryOrderPlus : [],
        categoryOrderMinus: Array.isArray(categoryOrderMinus) ? categoryOrderMinus : [],
        limitRiwayat: limit
      };
    }
  }

  return { exists: false };
}

function simpanPengaturanPengguna(settings) {
  settings = settings || {};

  var userKey = String(settings.userKey || '').trim();
  if (!userKey) {
    return { sukses: false, pesan: 'UserKey tidak valid.' };
  }

  var widgetConfig = Array.isArray(settings.widgetConfig) ? settings.widgetConfig : [];
  var rekeningVisibility =
    (settings.rekeningVisibility &&
     typeof settings.rekeningVisibility === 'object' &&
     !Array.isArray(settings.rekeningVisibility))
      ? settings.rekeningVisibility
      : {};

  var categoryOrderPlus = Array.isArray(settings.categoryOrderPlus)
    ? settings.categoryOrderPlus
    : [];

  var categoryOrderMinus = Array.isArray(settings.categoryOrderMinus)
    ? settings.categoryOrderMinus
    : [];

  var limit = parseInt(settings.limitRiwayat, 10);
  if (isNaN(limit)) limit = 5;
  limit = Math.max(1, Math.min(25, limit));

  var sheet = getOrCreatePengaturanPenggunaSheet_();
  var lock = LockService.getScriptLock();

  lock.waitLock(10000);

  try {
    var now = new Date();
    var lastRow = sheet.getLastRow();

    if (lastRow >= 2) {
      var keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

      for (var i = 0; i < keys.length; i++) {
        if (String(keys[i][0] || '').trim() === userKey) {
          sheet.getRange(i + 2, 1, 1, 7).setValues([[
            userKey,
            JSON.stringify(widgetConfig),
            JSON.stringify(rekeningVisibility),
            JSON.stringify(categoryOrderPlus),
            JSON.stringify(categoryOrderMinus),
            limit,
            now
          ]]);

          return { sukses: true, updated: true };
        }
      }
    }

    sheet.appendRow([
      userKey,
      JSON.stringify(widgetConfig),
      JSON.stringify(rekeningVisibility),
      JSON.stringify(categoryOrderPlus),
      JSON.stringify(categoryOrderMinus),
      limit,
      now
    ]);

    return { sukses: true, updated: false };
  } finally {
    lock.releaseLock();
  }
}
