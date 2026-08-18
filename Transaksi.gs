// ============================================================
// CATATAN KEUANGAN - BACKEND MODULE
// Generated from the verified 7F-16 baseline.
// Functions are preserved in full; only file ownership was reorganized.
// ============================================================

function simpanData(data) {
  if (!data || !data.userKey) return 'Gagal';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Transaksi');
  if (!sheet) {
    sheet = ss.insertSheet('Transaksi');
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Tanggal", "Tipe", "Kategori", "Sub Kategori", "jumlah", "Rekening", "Dari/kepada", "Catatan", "UserKey"]);
  }

  sheet.appendRow([
    data.tanggal,
    "'" + String(data.tipe).replace("'", ""),
    data.kategori,
    data.subKategori || '',
    Number(data.jumlah),
    data.rekening,
    data.dariKepada || '',
    data.catatan || '',
    data.userKey
  ]);

  updateSaldoRekening(data.userKey, data.rekening, data.tipe, Number(data.jumlah));

  return "Sukses";
}

function hapusTransaksi(row, userKey) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Transaksi');
  if (!sheet || !row || !userKey) return 'Gagal';

  var rowData = sheet.getRange(row, 1, 1, 9).getValues()[0];
  var ownerBaris = String(rowData[8] || '').trim();
  if (ownerBaris !== userKey) return 'Gagal';

  var tipe = String(rowData[1] || '').replace("'", "").trim();
  var jumlah = Number(rowData[4]) || 0;
  var rekening = String(rowData[5] || '').trim();

  sheet.deleteRow(row);

  var tipeBalik = tipe === '+' ? '-' : '+';
  updateSaldoRekening(userKey, rekening, tipeBalik, jumlah);

  return 'Sukses';
}

function updateTransaksi(param) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Transaksi');
  if (!sheet || !param || !param.row || !param.userKey) return 'Gagal';

  var row = Number(param.row);
  var oldData = sheet.getRange(row, 1, 1, 9).getValues()[0];
  var ownerBaris = String(oldData[8] || '').trim();
  if (ownerBaris !== param.userKey) return 'Gagal';

  var oldTipe = String(oldData[1] || '').replace("'", "").trim();
  var oldJumlah = Number(oldData[4]) || 0;
  var oldRekening = String(oldData[5] || '').trim();

  var tipeBalik = oldTipe === '+' ? '-' : '+';
  updateSaldoRekening(param.userKey, oldRekening, tipeBalik, oldJumlah);

  sheet.getRange(row, 1, 1, 9).setValues([[
    param.tanggal,
    "'" + String(param.tipe).replace("'", ""),
    param.kategori,
    param.subKategori || '',
    Number(param.jumlah),
    param.rekening,
    param.dariKepada || '',
    param.catatan || '',
    ownerBaris
  ]]);

  updateSaldoRekening(param.userKey, param.rekening, param.tipe, Number(param.jumlah));

  return 'Sukses';
}
