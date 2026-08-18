// ============================================================
// CATATAN KEUANGAN - BACKEND MODULE
// Generated from the verified 7F-16 baseline.
// Functions are preserved in full; only file ownership was reorganized.
// ============================================================

function getIkhtisarData(userKey) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Transaksi');
  var res = {
    hariIni: {in:0, out:0},
    kemarin: {in:0, out:0},
    tujuhHari: {in:0, out:0},
    bulanIni: {in:0, out:0},
    bulanLalu: {in:0, out:0},
    tahunIni: {in:0, out:0},
    tahunLalu: {in:0, out:0},
    total: {in:0, out:0}
  };
  if (!sheet || !userKey) return res;

  var data = sheet.getDataRange().getValues();
  var now = new Date();
  var today = formatDate(now);
  
  var nowTime = now.getTime();
  var yesterdayStr = formatDate(new Date(nowTime - 86400000));
  var sevenDaysAgoStr = formatDate(new Date(nowTime - 7 * 86400000));

  var curMonthStr = today.substring(0, 7);
  var curYearStr = today.substring(0, 4);
  var lastMonthDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  var lastMonthStr = formatDate(lastMonthDate).substring(0, 7);
  var lastYearStr = String(Number(curYearStr) - 1);

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var ownerBaris = String(data[i][8] || '').trim();
    if (ownerBaris !== userKey) continue;

    var tglStr = formatDate(data[i][0]);
    var tipe = String(data[i][1] || '').replace("'", "").trim();
    var val = Number(data[i][4]) || 0;
    var bucket = tipe === '+' ? 'in' : (tipe === '-' ? 'out' : null);
    if (!bucket) continue;

    res.total[bucket] += val;

    if (tglStr === today) res.hariIni[bucket] += val;
    if (tglStr === yesterdayStr) res.kemarin[bucket] += val;
    if (tglStr >= sevenDaysAgoStr && tglStr <= today) res.tujuhHari[bucket] += val;
    if (tglStr.substring(0, 7) === curMonthStr) res.bulanIni[bucket] += val;
    if (tglStr.substring(0, 7) === lastMonthStr) res.bulanLalu[bucket] += val;
    if (tglStr.substring(0, 4) === curYearStr) res.tahunIni[bucket] += val;
    if (tglStr.substring(0, 4) === lastYearStr) res.tahunLalu[bucket] += val;
  }
  return res;
}

function getRiwayatTransaksi(userKey) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Transaksi");
  if (!sheet || !userKey) return [];

  var data = sheet.getDataRange().getValues();
  var kategoriList = getKategoriData(userKey);
  var riwayat = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var ownerBaris = String(row[8] || '').trim();
    if (row[0] && ownerBaris === userKey) {
      var tglStr = formatDate(row[0]);
      var tipe = String(row[1] || '').replace("'", "").trim();
      var katName = String(row[2] || '').trim();
      var subKatName = String(row[3] || '').trim();

      var icon = tipe === '+' ? '📥' : '📤';
      for (var k = 0; k < kategoriList.length; k++) {
        if (kategoriList[k].kategori.toLowerCase() === katName.toLowerCase()) {
          if (subKatName && kategoriList[k].subKategori.toLowerCase() === subKatName.toLowerCase()) {
            icon = kategoriList[k].iconSub || kategoriList[k].iconKat;
            break;
          } else if (!subKatName) {
            icon = kategoriList[k].iconKat;
          }
        }
      }

      riwayat.push({
        row: i + 1,
        tanggal: tglStr,
        tipe: tipe,
        kategori: katName,
        subKategori: subKatName,
        jumlah: Number(row[4]) || 0,
        rekening: String(row[5] || '').trim(),
        dariKepada: String(row[6] || '').trim(),
        catatan: String(row[7] || '').trim(),
        icon: icon
      });
    }
  }
  return riwayat.reverse();
}

function getRiwayatPeriode(userKey, mode, tanggalISO) {
  var res = { list: [], totalIn: 0, totalOut: 0 };
  if (!userKey || !tanggalISO) return res;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Transaksi');
  if (!sheet) return res;

  var data = sheet.getDataRange().getValues();
  var kategoriList = getKategoriData(userKey);

  var keyMatch;
  if (mode === 'bulan') {
    keyMatch = String(tanggalISO).substring(0, 7);
  } else if (mode === 'tahun') {
    keyMatch = String(tanggalISO).substring(0, 4);
  } else {
    keyMatch = String(tanggalISO).substring(0, 10);
  }

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    var ownerBaris = String(row[8] || '').trim();
    if (ownerBaris !== userKey) continue;

    var tglStr = formatDate(row[0]);
    var cocok = false;
    if (mode === 'bulan') cocok = (tglStr.substring(0, 7) === keyMatch);
    else if (mode === 'tahun') cocok = (tglStr.substring(0, 4) === keyMatch);
    else cocok = (tglStr === keyMatch);

    if (!cocok) continue;

    var tipe = String(row[1] || '').replace("'", "").trim();
    var katName = String(row[2] || '').trim();
    var subKatName = String(row[3] || '').trim();
    var jumlah = Number(row[4]) || 0;

    var icon = tipe === '+' ? '📥' : '📤';
    for (var k = 0; k < kategoriList.length; k++) {
      if (kategoriList[k].kategori.toLowerCase() === katName.toLowerCase()) {
        if (subKatName && kategoriList[k].subKategori.toLowerCase() === subKatName.toLowerCase()) {
          icon = kategoriList[k].iconSub || kategoriList[k].iconKat;
          break;
        } else if (!subKatName) {
          icon = kategoriList[k].iconKat;
        }
      }
    }

    res.list.push({
      row: i + 1,
      tanggal: tglStr,
      tipe: tipe,
      kategori: katName,
      subKategori: subKatName,
      jumlah: jumlah,
      rekening: String(row[5] || '').trim(),
      dariKepada: String(row[6] || '').trim(),
      catatan: String(row[7] || '').trim(),
      icon: icon
    });

    if (tipe === '+') res.totalIn += jumlah;
    else if (tipe === '-') res.totalOut += jumlah;
  }

  res.list.reverse();
  return res;
}

function getBreakdownKategoriBulanan(userKey, tipe, viewMode, bulanISO) {
  var res = { total: 0, items: [], dominant: null };
  if (!userKey || !bulanISO) return res;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Transaksi');
  if (!sheet) return res;

  var data = sheet.getDataRange().getValues();
  var kategoriList = getKategoriData(userKey);

  var targetBulan = String(bulanISO).substring(0, 7);
  var tipeTarget = String(tipe || '-').replace("'", "").trim();

  var mapGroup = {};
  var grandTotal = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    var ownerBaris = String(row[8] || '').trim();
    if (ownerBaris !== userKey) continue;

    var tglStr = formatDate(row[0]);
    if (tglStr.substring(0, 7) !== targetBulan) continue;

    var t = String(row[1] || '').replace("'", "").trim();
    if (t !== tipeTarget) continue;

    var katName = String(row[2] || '').trim();
    var subKatName = String(row[3] || '').trim();
    var jumlah = Number(row[4]) || 0;

    var keyName = (viewMode === 'sub' && subKatName) ? (katName + ' › ' + subKatName) : katName;

    var icon = t === '+' ? '📥' : '📤';
    for (var k = 0; k < kategoriList.length; k++) {
      if (kategoriList[k].kategori.toLowerCase() === katName.toLowerCase()) {
        if (subKatName && kategoriList[k].subKategori.toLowerCase() === subKatName.toLowerCase()) {
          icon = kategoriList[k].iconSub || kategoriList[k].iconKat;
          break;
        } else if (!subKatName) {
          icon = kategoriList[k].iconKat;
        }
      }
    }

    if (!mapGroup[keyName]) {
      mapGroup[keyName] = {
        name: keyName,
        kat: katName,
        sub: subKatName,
        jumlah: 0,
        icon: icon
      };
    }
    mapGroup[keyName].jumlah += jumlah;
    grandTotal += jumlah;
  }

  var itemList = [];
  for (var kName in mapGroup) {
    var item = mapGroup[kName];
    var persentase = grandTotal > 0 ? (item.jumlah / grandTotal) * 100 : 0;
    itemList.push({
      name: item.name,
      kat: item.kat,
      sub: item.sub,
      jumlah: item.jumlah,
      icon: item.icon,
      persentase: Math.round(persentase * 10) / 10
    });
  }

  itemList.sort(function(a, b) { return b.jumlah - a.jumlah; });

  res.total = grandTotal;
  res.items = itemList;
  res.dominant = itemList.length > 0 ? itemList[0] : null;

  return res;
}
