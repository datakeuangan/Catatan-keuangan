// ============================================================
// CATATAN KEUANGAN - BACKEND MODULE
// Generated from the verified 7F-16 baseline.
// Functions are preserved in full; only file ownership was reorganized.
// ============================================================

function getOrCreateGoalsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Goals');
  if (!sheet) {
    sheet = ss.insertSheet('Goals');
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["ID", "Nama", "Target", "Terkumpul", "TanggalTarget", "Icon", "Status", "Catatan", "UserKey"]);
  }
  return sheet;
}

function getGoalsData(userKey) {
  var sheet = getOrCreateGoalsSheet_();
  var data = sheet.getDataRange().getValues();
  var list = [];
  if (!userKey) return list;

  for (var i = 1; i < data.length; i++) {
    var ownerBaris = String(data[i][8] || '').trim();
    if (ownerBaris !== userKey) continue;

    var id = String(data[i][0] || ('GOAL_' + i)).trim();
    var nama = String(data[i][1] || '').trim();
    var target = Number(data[i][2]) || 0;
    var terkumpul = Number(data[i][3]) || 0;
    var tglTarget = data[i][4] ? formatDate(data[i][4]) : '';
    var icon = String(data[i][5] || '🎯').trim();
    var status = String(data[i][6] || 'AKTIF').trim().toUpperCase();
    var catatan = String(data[i][7] || '').trim();

    var persentase = target > 0 ? Math.min(100, Math.round((terkumpul / target) * 100)) : 0;
    var sisa = Math.max(0, target - terkumpul);

    list.push({
      row: i + 1,
      id: id,
      nama: nama,
      target: target,
      terkumpul: terkumpul,
      sisa: sisa,
      persentase: persentase,
      tglTarget: tglTarget,
      icon: icon,
      status: status,
      catatan: catatan
    });
  }
  return list;
}

function simpanGoalBaru(param) {
  if (!param || !param.nama || !param.userKey) return getGoalsData(param ? param.userKey : '');
  var sheet = getOrCreateGoalsSheet_();
  var goalId = 'GOAL_' + Date.now();

  sheet.appendRow([
    goalId,
    String(param.nama).trim(),
    Number(param.target) || 0,
    Number(param.terkumpul) || 0,
    param.tglTarget || '',
    param.icon || '🎯',
    'AKTIF',
    param.catatan || '',
    param.userKey
  ]);
  return getGoalsData(param.userKey);
}

function setorKeGoal(param) {
  if (!param || !param.row || !param.userKey) return getGoalsData(param ? param.userKey : '');
  var sheet = getOrCreateGoalsSheet_();
  var row = Number(param.row);
  var ownerBaris = String(sheet.getRange(row, 9).getValue() || '').trim();

  if (ownerBaris === param.userKey) {
    var terkumpulLama = Number(sheet.getRange(row, 4).getValue()) || 0;
    var jumlahSetor = Number(param.jumlah) || 0;
    var terkumpulBaru = terkumpulLama + jumlahSetor;
    var target = Number(sheet.getRange(row, 3).getValue()) || 0;

    sheet.getRange(row, 4).setValue(terkumpulBaru);
    if (target > 0 && terkumpulBaru >= target) {
      sheet.getRange(row, 7).setValue('SELESAI');
    }

    if (param.rekening) {
      updateSaldoRekening(param.userKey, param.rekening, '-', jumlahSetor);
    }
  }
  return getGoalsData(param.userKey);
}

function tarikDariGoal(param) {
  if (!param || !param.row || !param.userKey) return getGoalsData(param ? param.userKey : '');
  var sheet = getOrCreateGoalsSheet_();
  var row = Number(param.row);
  var ownerBaris = String(sheet.getRange(row, 9).getValue() || '').trim();

  if (ownerBaris === param.userKey) {
    var terkumpulLama = Number(sheet.getRange(row, 4).getValue()) || 0;
    var jumlahTarik = Number(param.jumlah) || 0;
    var terkumpulBaru = Math.max(0, terkumpulLama - jumlahTarik);
    var target = Number(sheet.getRange(row, 3).getValue()) || 0;

    sheet.getRange(row, 4).setValue(terkumpulBaru);
    if (target > 0 && terkumpulBaru < target) {
      sheet.getRange(row, 7).setValue('AKTIF');
    }

    if (param.rekening) {
      updateSaldoRekening(param.userKey, param.rekening, '+', jumlahTarik);
    }
  }
  return getGoalsData(param.userKey);
}

function updateGoalDetail(param) {
  if (!param || !param.row || !param.userKey) return getGoalsData(param ? param.userKey : '');
  var sheet = getOrCreateGoalsSheet_();
  var row = Number(param.row);
  var ownerBaris = String(sheet.getRange(row, 9).getValue() || '').trim();

  if (ownerBaris === param.userKey) {
    sheet.getRange(row, 2).setValue(param.nama);
    sheet.getRange(row, 3).setValue(Number(param.target) || 0);
    sheet.getRange(row, 4).setValue(Number(param.terkumpul) || 0);
    sheet.getRange(row, 5).setValue(param.tglTarget || '');
    sheet.getRange(row, 6).setValue(param.icon || '🎯');
    sheet.getRange(row, 7).setValue(param.status || 'AKTIF');
    sheet.getRange(row, 8).setValue(param.catatan || '');
  }
  return getGoalsData(param.userKey);
}

function hapusGoal(rowNum, userKey) {
  var sheet = getOrCreateGoalsSheet_();
  if (sheet && rowNum && userKey) {
    var ownerBaris = String(sheet.getRange(rowNum, 9).getValue() || '').trim();
    if (ownerBaris === userKey) {
      sheet.deleteRow(rowNum);
    }
  }
  return getGoalsData(userKey);
}
