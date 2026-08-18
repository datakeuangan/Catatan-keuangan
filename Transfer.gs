// ============================================================
// CATATAN KEUANGAN - BACKEND MODULE
// Generated from the verified 7F-16 baseline.
// Functions are preserved in full; only file ownership was reorganized.
// ============================================================

function getOrCreateTransferSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Transfer');
  if (!sheet) {
    sheet = ss.insertSheet('Transfer');
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Tanggal", "Dari", "Kepada", "Jumlah", "Catatan", "UserKey"]);
  }
  return sheet;
}

function simpanTransfer(param) {
  if (!param || !param.dari || !param.kepada || !param.userKey) return 'Gagal';
  if (param.dari === param.kepada) return 'Gagal';

  var jumlah = Number(param.jumlah) || 0;
  if (jumlah <= 0) return 'Gagal';

  var sheet = getOrCreateTransferSheet_();
  sheet.appendRow([
    param.tanggal,
    param.dari,
    param.kepada,
    jumlah,
    param.catatan || '',
    param.userKey
  ]);

  updateSaldoRekening(param.userKey, param.dari, '-', jumlah);
  updateSaldoRekening(param.userKey, param.kepada, '+', jumlah);

  return 'Sukses';
}

function getRiwayatTransfer(userKey) {
  var sheet = getOrCreateTransferSheet_();
  var data = sheet.getDataRange().getValues();
  var riwayat = [];
  if (!userKey) return riwayat;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var ownerBaris = String(row[5] || '').trim();
    if (row[0] && ownerBaris === userKey) {
      riwayat.push({
        row: i + 1,
        tanggal: formatDate(row[0]),
        dari: String(row[1] || '').trim(),
        kepada: String(row[2] || '').trim(),
        jumlah: Number(row[3]) || 0,
        catatan: String(row[4] || '').trim()
      });
    }
  }
  return riwayat.reverse();
}

function hapusTransfer(row, userKey) {
  var sheet = getOrCreateTransferSheet_();
  if (!row || !userKey) return 'Gagal';

  var rowData = sheet.getRange(row, 1, 1, 6).getValues()[0];
  var ownerBaris = String(rowData[5] || '').trim();
  if (ownerBaris !== userKey) return 'Gagal';

  var dari = String(rowData[1] || '').trim();
  var kepada = String(rowData[2] || '').trim();
  var jumlah = Number(rowData[3]) || 0;

  sheet.deleteRow(row);

  updateSaldoRekening(userKey, dari, '+', jumlah);
  updateSaldoRekening(userKey, kepada, '-', jumlah);

  return 'Sukses';
}
