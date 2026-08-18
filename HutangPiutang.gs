// ============================================================
// CATATAN KEUANGAN - BACKEND MODULE
// Generated from the verified 7F-16 baseline.
// Functions are preserved in full; only file ownership was reorganized.
// ============================================================

function getOrCreateHutangPiutangSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('HutangPiutang');
  if (!sheet) {
    sheet = ss.insertSheet('HutangPiutang');
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["ID", "Tipe", "Nama", "Jumlah", "Terbayar", "TglJatuhTempo", "Status", "Catatan", "UserKey"]);
  }
  return sheet;
}

function getHutangPiutangData(userKey) {
  var sheet = getOrCreateHutangPiutangSheet_();
  var data = sheet.getDataRange().getValues();
  var list = [];
  var summary = {
    totalHutang: 0,
    totalPiutang: 0,
    sisaHutang: 0,
    sisaPiutang: 0
  };

  if (!userKey) return { list: list, summary: summary };

  for (var i = 1; i < data.length; i++) {
    var ownerBaris = String(data[i][8] || '').trim();
    if (ownerBaris !== userKey) continue;

    var id = String(data[i][0] || ('HP_' + i)).trim();
    var tipe = String(data[i][1] || 'HUTANG').trim().toUpperCase(); // 'HUTANG' atau 'PIUTANG'
    var nama = String(data[i][2] || '').trim();
    var jumlah = Number(data[i][3]) || 0;
    var terbayar = Number(data[i][4]) || 0;
    var tglJatuhTempo = data[i][5] ? formatDate(data[i][5]) : '';
    var status = String(data[i][6] || 'BELUM LUNAS').trim().toUpperCase();
    var catatan = String(data[i][7] || '').trim();

    var sisa = Math.max(0, jumlah - terbayar);
    var persentase = jumlah > 0 ? Math.min(100, Math.round((terbayar / jumlah) * 100)) : 0;

    if (tipe === 'HUTANG') {
      summary.totalHutang += jumlah;
      summary.sisaHutang += sisa;
    } else {
      summary.totalPiutang += jumlah;
      summary.sisaPiutang += sisa;
    }

    list.push({
      row: i + 1,
      id: id,
      tipe: tipe,
      nama: nama,
      jumlah: jumlah,
      terbayar: terbayar,
      sisa: sisa,
      persentase: persentase,
      tglJatuhTempo: tglJatuhTempo,
      status: status,
      catatan: catatan
    });
  }

  list.reverse();
  return { list: list, summary: summary };
}

function simpanHutangPiutangBaru(param) {
  if (!param || !param.nama || !param.userKey) return getHutangPiutangData(param ? param.userKey : '');
  var sheet = getOrCreateHutangPiutangSheet_();
  var id = 'HP_' + Date.now();

  var jumlah = Number(param.jumlah) || 0;
  var terbayar = Number(param.terbayar) || 0;
  var status = (jumlah > 0 && terbayar >= jumlah) ? 'LUNAS' : 'BELUM LUNAS';
  var tipe = String(param.tipe || 'HUTANG').trim().toUpperCase();

  sheet.appendRow([
    id,
    tipe,
    String(param.nama).trim(),
    jumlah,
    terbayar,
    param.tglJatuhTempo || '',
    status,
    param.catatan || '',
    param.userKey
  ]);

  if (param.rekening) {
    if (tipe === 'HUTANG') {
      updateSaldoRekening(param.userKey, param.rekening, '+', jumlah);
    } else {
      updateSaldoRekening(param.userKey, param.rekening, '-', jumlah);
    }
  }

  return getHutangPiutangData(param.userKey);
}

function bayarHutangPiutang(param) {
  if (!param || !param.row || !param.userKey) return getHutangPiutangData(param ? param.userKey : '');
  var sheet = getOrCreateHutangPiutangSheet_();
  var row = Number(param.row);
  var ownerBaris = String(sheet.getRange(row, 9).getValue() || '').trim();

  if (ownerBaris === param.userKey) {
    var tipe = String(sheet.getRange(row, 2).getValue() || 'HUTANG').trim().toUpperCase();
    var terbayarLama = Number(sheet.getRange(row, 5).getValue()) || 0;
    var jumlahBayar = Number(param.jumlah) || 0;
    var terbayarBaru = terbayarLama + jumlahBayar;
    var jumlahTarget = Number(sheet.getRange(row, 4).getValue()) || 0;

    sheet.getRange(row, 5).setValue(terbayarBaru);
    if (jumlahTarget > 0 && terbayarBaru >= jumlahTarget) {
      sheet.getRange(row, 7).setValue('LUNAS');
    }

    if (param.rekening) {
      if (tipe === 'HUTANG') {
        updateSaldoRekening(param.userKey, param.rekening, '-', jumlahBayar);
      } else {
        updateSaldoRekening(param.userKey, param.rekening, '+', jumlahBayar);
      }
    }
  }
  return getHutangPiutangData(param.userKey);
}

function updateHutangPiutangDetail(param) {
  if (!param || !param.row || !param.userKey) return getHutangPiutangData(param ? param.userKey : '');
  var sheet = getOrCreateHutangPiutangSheet_();
  var row = Number(param.row);
  var ownerBaris = String(sheet.getRange(row, 9).getValue() || '').trim();

  if (ownerBaris === param.userKey) {
    var jumlah = Number(param.jumlah) || 0;
    var terbayar = Number(param.terbayar) || 0;
    var status = (jumlah > 0 && terbayar >= jumlah) ? 'LUNAS' : (param.status || 'BELUM LUNAS');

    sheet.getRange(row, 2).setValue(String(param.tipe || 'HUTANG').trim().toUpperCase());
    sheet.getRange(row, 3).setValue(String(param.nama).trim());
    sheet.getRange(row, 4).setValue(jumlah);
    sheet.getRange(row, 5).setValue(terbayar);
    sheet.getRange(row, 6).setValue(param.tglJatuhTempo || '');
    sheet.getRange(row, 7).setValue(status);
    sheet.getRange(row, 8).setValue(param.catatan || '');
  }
  return getHutangPiutangData(param.userKey);
}

function hapusHutangPiutang(rowNum, userKey) {
  var sheet = getOrCreateHutangPiutangSheet_();
  if (sheet && rowNum && userKey) {
    var ownerBaris = String(sheet.getRange(rowNum, 9).getValue() || '').trim();
    if (ownerBaris === userKey) {
      sheet.deleteRow(rowNum);
    }
  }
  return getHutangPiutangData(userKey);
}
