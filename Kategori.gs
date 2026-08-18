// ============================================================
// CATATAN KEUANGAN - BACKEND MODULE
// Generated from the verified 7F-16 baseline.
// Functions are preserved in full; only file ownership was reorganized.
// ============================================================

function getKategoriData(userKey) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Kategori');
    if (!sheet || !userKey) return [];

    var data = sheet.getDataRange().getValues();
    var list = [];
    for (var i = 1; i < data.length; i++) {
      var ownerBaris = String(data[i][5] || '').trim();
      if (ownerBaris !== userKey) continue;
      list.push({
        row: i + 1,
        tipe: String(data[i][0] || '').replace("'", "").trim(),
        kategori: String(data[i][1] || '').trim(),
        subKategori: String(data[i][2] || '').trim(),
        iconKat: String(data[i][3] || '📁').trim(),
        iconSub: String(data[i][4] || '').trim()
      });
    }
    return list;
  } catch (e) {
    return [];
  }
}

function simpanKategoriBaru(param) {
  if (!param || !param.userKey) return [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Kategori');
  if (!sheet) {
    sheet = ss.insertSheet('Kategori');
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Tipe", "Kategori", "Sub Kategori", "Icon Kategori", "Icon Sub Kategori", "UserKey"]);
  }

  var tipeFormatted = "'" + String(param.tipe).replace("'", "");
  var subKat = param.subKategori || '';
  var iconSubFinal = subKat ? (param.iconSub || '📌') : '📌';

  sheet.appendRow([
    tipeFormatted,
    param.kategori,
    subKat,
    param.iconKat || '📁',
    iconSubFinal,
    param.userKey
  ]);
  return getKategoriData(param.userKey);
}

function updateKategoriDetail(param) {
  if (!param || !param.userKey) return [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Kategori');
  if (sheet && param.row) {
    var ownerBaris = String(sheet.getRange(param.row, 6).getValue() || '').trim();
    if (ownerBaris === param.userKey) {
      if (param.target === 'kat') {
        sheet.getRange(param.row, 2).setValue(param.nama);
        sheet.getRange(param.row, 4).setValue(param.icon);
      } else {
        sheet.getRange(param.row, 3).setValue(param.nama);
        sheet.getRange(param.row, 5).setValue(param.icon);

        if (param.newParent) {
          sheet.getRange(param.row, 2).setValue(param.newParent);
          var iconIndukBaru = cariIconKategoriUtama_(sheet, param.newParent, param.userKey);
          if (iconIndukBaru) sheet.getRange(param.row, 4).setValue(iconIndukBaru);
        }
      }
    }
  }
  return getKategoriData(param.userKey);
}

function cariIconKategoriUtama_(sheet, namaKategori, userKey) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var kat = String(data[i][1] || '').trim();
    var owner = String(data[i][5] || '').trim();
    if (owner === userKey && kat.toLowerCase() === String(namaKategori).trim().toLowerCase()) {
      var icon = String(data[i][3] || '').trim();
      if (icon) return icon;
    }
  }
  return '';
}

function hapusKategori(rowNum, userKey) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Kategori');
  if (sheet && rowNum && userKey) {
    var ownerBaris = String(sheet.getRange(rowNum, 6).getValue() || '').trim();
    if (ownerBaris === userKey) {
      sheet.deleteRow(rowNum);
    }
  }
  return getKategoriData(userKey);
}
