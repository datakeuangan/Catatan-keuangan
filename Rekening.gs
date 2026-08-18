// ============================================================
// CATATAN KEUANGAN - BACKEND MODULE
// Generated from the verified 7F-16 baseline.
// Functions are preserved in full; only file ownership was reorganized.
// ============================================================

function getRekeningData(userKey) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetRek = ss.getSheetByName('Rekening');

  var rekeningList = [];
  if (!sheetRek || !userKey) return rekeningList;

  var dataRek = sheetRek.getDataRange().getValues();
  for (var i = 1; i < dataRek.length; i++) {
    var nama = String(dataRek[i][0] || '').trim();
    var ownerBaris = String(dataRek[i][3] || '').trim();
    if (nama && ownerBaris === userKey) {
      var saldo = Number(dataRek[i][1]) || 0;
      rekeningList.push({
        row: i + 1,
        nama: nama,
        saldoAwal: saldo,
        icon: String(dataRek[i][2] || '💳').trim(),
        saldoAkhir: saldo
      });
    }
  }
  return rekeningList;
}

function simpanRekeningBaru(param) {
  if (!param || !param.userKey) return [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Rekening');
  if (!sheet) {
    sheet = ss.insertSheet('Rekening');
    sheet.appendRow(["Nama", "Saldo", "Icon", "UserKey"]);
  }
  
  var saldoAwal = Number(param.saldoAwal) || 0;
  sheet.appendRow([param.nama, saldoAwal, param.icon || '💳', param.userKey]);

  if (saldoAwal > 0) {
    var sheetTx = ss.getSheetByName('Transaksi');
    if (!sheetTx) {
      sheetTx = ss.insertSheet('Transaksi');
      sheetTx.appendRow(["Tanggal", "Tipe", "Kategori", "Sub Kategori", "jumlah", "Rekening", "Dari/kepada", "Catatan", "UserKey"]);
    }
    var todayStr = formatDate(new Date());
    sheetTx.appendRow([
      todayStr,
      "'+",
      "Pemasukan",
      "Saldo Awal",
      saldoAwal,
      param.nama,
      "",
      "Saldo Awal Rekening " + param.nama,
      param.userKey
    ]);
  }

  return getRekeningData(param.userKey);
}

function updateRekeningDetail(param) {
  if (!param || !param.userKey) return [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Rekening');
  if (!sheet) return getRekeningData(param.userKey);

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var nama = String(data[i][0] || '').trim();
    var ownerBaris = String(data[i][3] || '').trim();
    if (ownerBaris === param.userKey && nama.toLowerCase() === String(param.namaLama).trim().toLowerCase()) {
      var saldoLama = Number(data[i][1]) || 0;
      var saldoBaru = Number(param.saldoBaru) || 0;
      var selisih = saldoBaru - saldoLama;

      sheet.getRange(i + 1, 1).setValue(param.namaBaru);
      sheet.getRange(i + 1, 2).setValue(saldoBaru);
      sheet.getRange(i + 1, 3).setValue(param.icon || '💳');

      if (selisih !== 0) {
        var sheetTx = ss.getSheetByName('Transaksi');
        if (!sheetTx) {
          sheetTx = ss.insertSheet('Transaksi');
          sheetTx.appendRow(["Tanggal", "Tipe", "Kategori", "Sub Kategori", "jumlah", "Rekening", "Dari/kepada", "Catatan", "UserKey"]);
        }
        var todayStr = formatDate(new Date());
        var tipeTx = selisih > 0 ? "'+ " : "'-";
        var katTx = selisih > 0 ? "Pemasukan" : "Belanja";
        var subKatTx = "Penyesuaian Saldo";
        var ketTx = (selisih > 0 ? "Penambahan" : "Pengurangan") + " Saldo Rekening " + param.namaBaru;

        sheetTx.appendRow([
          todayStr,
          tipeTx,
          katTx,
          subKatTx,
          Math.abs(selisih),
          param.namaBaru,
          "",
          ketTx,
          param.userKey
        ]);
      }
      break;
    }
  }
  return getRekeningData(param.userKey);
}

function updateSaldoRekening(userKey, namaRekening, tipe, jumlah) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Rekening');
  if (!sheet || !userKey) return;

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var nama = String(data[i][0] || '').trim();
    var ownerBaris = String(data[i][3] || '').trim();
    if (ownerBaris === userKey && nama.toLowerCase() === String(namaRekening).trim().toLowerCase()) {
      var saldoLama = Number(data[i][1]) || 0;
      var saldoBaru = tipe === '+' ? saldoLama + jumlah : saldoLama - jumlah;
      sheet.getRange(i + 1, 2).setValue(saldoBaru);
      break;
    }
  }
}

function syncSaldoAwal() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetRek = ss.getSheetByName('Rekening');
  var sheetTx = ss.getSheetByName('Transaksi');
  if (!sheetRek || !sheetTx) return;

  var dataRek = sheetRek.getDataRange().getValues();
  var dataTx = sheetTx.getDataRange().getValues();

  for (var i = 1; i < dataRek.length; i++) {
    var nama = String(dataRek[i][0] || '').trim();
    var ownerRek = String(dataRek[i][3] || '').trim();
    if (!nama) continue;

    var saldo = 0;
    for (var j = 1; j < dataTx.length; j++) {
      var tipe = String(dataTx[j][1] || '').replace("'", "").trim();
      var jumlah = Number(dataTx[j][4]) || 0;
      var rekNama = String(dataTx[j][5] || '').trim();
      var ownerTx = String(dataTx[j][8] || '').trim();
      if (rekNama.toLowerCase() === nama.toLowerCase() && ownerTx === ownerRek) {
        if (tipe === '+') saldo += jumlah;
        if (tipe === '-') saldo -= jumlah;
      }
    }
    sheetRek.getRange(i + 1, 2).setValue(saldo);
  }
}

function hapusRekening(rowNum, userKey) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Rekening');
  if (sheet && rowNum && userKey) {
    var ownerBaris = String(sheet.getRange(rowNum, 4).getValue() || '').trim();
    if (ownerBaris === userKey) {
      sheet.deleteRow(rowNum);
    }
  }
  return getRekeningData(userKey);
}
