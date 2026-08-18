/**
 * ============================================================
 * MODUL CICILAN - V1
 * ============================================================
 * File terpisah dari Code.gs utama.
 *
 * Sheet yang dibuat otomatis:
 *   1. Cicilan
 *   2. CicilanPembayaran
 *
 * Prinsip:
 * - Tidak mengubah sheet HutangPiutang yang sudah ada.
 * - Setiap cicilan mengacu ke ID HutangPiutang.
 * - UserKey memisahkan data antar pengguna.
 * - Sheet dibuat otomatis jika belum ada.
 * ============================================================
 */

const CICILAN_CONFIG = {
  SHEET_CICILAN: 'Cicilan',
  SHEET_PEMBAYARAN: 'CicilanPembayaran',

  HEADER_CICILAN: [
    'ID',
    'HutangPiutangID',
    'Nama',
    'Tipe',
    'Total',
    'NominalCicilan',
    'Tenor',
    'TanggalMulai',
    'TanggalJatuhTempo',
    'Frekuensi',
    'CicilanKe',
    'SudahDibayar',
    'Sisa',
    'Status',
    'Catatan',
    'UserKey',
    'CreatedAt',
    'UpdatedAt'
  ],

  HEADER_PEMBAYARAN: [
    'ID',
    'CicilanID',
    'HutangPiutangID',
    'TanggalBayar',
    'Nominal',
    'CicilanKe',
    'Catatan',
    'UserKey',
    'CreatedAt',
    'Rekening',
    'TransaksiID'
  ]
};

/**
 * Jalankan sekali secara manual setelah file ini ditambahkan.
 * Fungsi ini aman dijalankan berulang kali.
 */
function setupModulCicilan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = {
    cicilan: ensureCicilanSheet_(ss),
    pembayaran: ensureCicilanPembayaranSheet_(ss)
  };

  return {
    success: true,
    message: 'Modul Cicilan siap digunakan.',
    sheets: result
  };
}

function ensureCicilanSheet_(ss) {
  let sh = ss.getSheetByName(CICILAN_CONFIG.SHEET_CICILAN);

  if (!sh) {
    sh = ss.insertSheet(CICILAN_CONFIG.SHEET_CICILAN);
  }

  ensureHeaders_(sh, CICILAN_CONFIG.HEADER_CICILAN);
  return sh.getName();
}

function ensureCicilanPembayaranSheet_(ss) {
  let sh = ss.getSheetByName(CICILAN_CONFIG.SHEET_PEMBAYARAN);

  if (!sh) {
    sh = ss.insertSheet(CICILAN_CONFIG.SHEET_PEMBAYARAN);
  }

  ensureHeaders_(sh, CICILAN_CONFIG.HEADER_PEMBAYARAN);
  return sh.getName();
}

function ensureHeaders_(sh, headers) {
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    return;
  }

  const existing = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), headers.length))
    .getValues()[0];

  let changed = false;

  headers.forEach((header, i) => {
    if (existing[i] !== header) {
      sh.getRange(1, i + 1).setValue(header);
      changed = true;
    }
  });

  if (changed) sh.setFrozenRows(1);
}


/**
 * Membuat master Cicilan sekaligus membuat pengaturan Notifikasi
 * yang terikat ke ID Cicilan yang baru dibuat.
 *
 * Fungsi ini menjadi satu pintu untuk tombol "Simpan Cicilan".
 * Fungsi buatCicilan(data) lama tetap dipertahankan agar alur lama
 * dan fungsi lain tidak ikut berubah.
 */
function buatCicilanDenganNotifikasi(data) {
  if (!data) throw new Error('Data cicilan tidak ditemukan.');

  const payload = Object.assign({}, data);
  const notif = payload.notifikasi || {};

  // Validasi notifikasi dilakukan sebelum membuat master agar kesalahan
  // input tidak meninggalkan Cicilan tanpa konfigurasi yang dimaksud.
  const email = String(notif.email || '').trim();
  const jam = String(notif.jamPengecekan || '').trim();

  if (!email) throw new Error('Alamat email notifikasi wajib diisi.');
  if (typeof isValidEmail_ === 'function' && !isValidEmail_(email)) {
    throw new Error('Alamat email notifikasi tidak valid.');
  }
  if (!/^\d{2}:\d{2}$/.test(jam)) {
    throw new Error('Jam notifikasi tidak valid.');
  }

  // Hanya master Cicilan yang dibuat melalui fungsi lama.
  // ID hasilnya kemudian menjadi kunci notifikasi per transaksi.
  const hasilCicilan = buatCicilan(payload);

  if (!hasilCicilan || hasilCicilan.success !== true || !hasilCicilan.id) {
    throw new Error('Master Cicilan gagal dibuat.');
  }

  let hasilNotifikasi;
  try {
    hasilNotifikasi = simpanPengaturanNotifikasiCicilanPerID({
      userKey: String(payload.userKey || '').trim(),
      cicilanID: String(hasilCicilan.id).trim(),
      email: email,
      aktif: notif.aktif !== false,
      h7: notif.h7 !== false,
      h3: notif.h3 !== false,
      h1: notif.h1 !== false,
      hariH: notif.hariH !== false,
      setelahJatuhTempo: notif.setelahJatuhTempo !== false,
      maksHariTerlambat: Number(notif.maksHariTerlambat || 7),
      jamPengecekan: jam
    });
  } catch (errNotif) {
    hasilNotifikasi = {
      success: false,
      message: String(errNotif && errNotif.message || errNotif || 'Gagal menyimpan notifikasi.')
    };
  }

  return {
    success: true,
    id: hasilCicilan.id,
    message: hasilCicilan.message || 'Cicilan berhasil dibuat.',
    notification: hasilNotifikasi
  };
}

function buatCicilan(data) {
  setupModulCicilan();

  if (!data) throw new Error('Data cicilan tidak ditemukan.');

  const userKey = String(data.userKey || '').trim();
  if (!userKey) throw new Error('UserKey wajib diisi.');

  const hutangPiutangID = String(data.hutangPiutangID || '').trim();
  let tipeCicilan = String(data.tipe || 'Hutang').trim();
  let namaCicilan = String(data.nama || '').trim();
  let totalCicilan = toNumber_(data.total);

  // Jika cicilan berasal dari Hutang/Piutang, data master menjadi sumber kebenaran.
  if (hutangPiutangID && typeof getHutangPiutangData === 'function') {
    const hpData = getHutangPiutangData(userKey);
    const hp = (hpData.list || []).find(function(item) {
      return String(item.id || '').trim() === hutangPiutangID;
    });
    if (!hp) throw new Error('Hutang/Piutang yang dipilih tidak ditemukan atau bukan milik akun ini.');
    if (Number(hp.sisa || 0) <= 0) throw new Error('Hutang/Piutang yang dipilih sudah lunas.');
    tipeCicilan = hp.tipe === 'PIUTANG' ? 'Piutang' : 'Hutang';
    namaCicilan = hp.nama;
    totalCicilan = Number(hp.sisa || 0);
  }

  const total = totalCicilan;
  const nominal = toNumber_(data.nominalCicilan);
  const tenor = Math.max(1, parseInt(data.tenor, 10) || 1);

  if (total <= 0) throw new Error('Total cicilan harus lebih dari 0.');
  if (nominal <= 0) throw new Error('Nominal cicilan harus lebih dari 0.');

  const id = data.id || ('CIC-' + Utilities.getUuid().replace(/-/g, '').slice(0, 16));
  const now = new Date();

  const sh = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CICILAN_CONFIG.SHEET_CICILAN);

  const sudahDibayar = Math.max(0, toNumber_(data.sudahDibayar));
  const sisa = Math.max(0, total - sudahDibayar);
  const status = sisa <= 0 ? 'Selesai' : 'Aktif';

  const row = [
    id,
    hutangPiutangID,
    namaCicilan,
    tipeCicilan,
    total,
    nominal,
    tenor,
    parseDate_(data.tanggalMulai),
    parseDate_(data.tanggalJatuhTempo),
    data.frekuensi || 'Bulanan',
    parseInt(data.cicilanKe, 10) || 1,
    sudahDibayar,
    sisa,
    status,
    data.catatan || '',
    userKey,
    now,
    now
  ];

  sh.appendRow(row);

  return {
    success: true,
    id: id,
    message: 'Cicilan berhasil dibuat.'
  };
}

function ambilCicilan(userKey) {
  setupModulCicilan();

  userKey = String(userKey || '').trim();
  if (!userKey) return [];

  const sh = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CICILAN_CONFIG.SHEET_CICILAN);

  const values = getDataRows_(sh);
  return values
    .filter(r => String(r[15] || '') === userKey)
    .map(cicilanRowToObject_);
}

function ambilCicilanById(id, userKey) {
  return ambilCicilan(userKey).find(x => String(x.id) === String(id)) || null;
}

function updateCicilan(data) {
  setupModulCicilan();

  const id = String(data && data.id || '').trim();
  const userKey = String(data && data.userKey || '').trim();

  if (!id || !userKey) throw new Error('ID dan UserKey wajib diisi.');

  const sh = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CICILAN_CONFIG.SHEET_CICILAN);

  const rowIndex = findRowByIdAndUser_(sh, id, userKey, 1, 16);
  if (rowIndex < 2) throw new Error('Data cicilan tidak ditemukan.');

  const current = sh.getRange(rowIndex, 1, 1, CICILAN_CONFIG.HEADER_CICILAN.length).getValues()[0];

  const total = data.total !== undefined ? toNumber_(data.total) : toNumber_(current[4]);
  const nominal = data.nominalCicilan !== undefined ? toNumber_(data.nominalCicilan) : toNumber_(current[5]);
  const tenor = data.tenor !== undefined ? Math.max(1, parseInt(data.tenor, 10) || 1) : current[6];
  const sudahDibayar = data.sudahDibayar !== undefined ? toNumber_(data.sudahDibayar) : toNumber_(current[11]);
  const sisa = Math.max(0, total - sudahDibayar);

  current[1] = data.hutangPiutangID !== undefined ? data.hutangPiutangID : current[1];
  current[2] = data.nama !== undefined ? data.nama : current[2];
  current[3] = data.tipe !== undefined ? data.tipe : current[3];
  current[4] = total;
  current[5] = nominal;
  current[6] = tenor;
  current[7] = data.tanggalMulai !== undefined ? parseDate_(data.tanggalMulai) : current[7];
  current[8] = data.tanggalJatuhTempo !== undefined ? parseDate_(data.tanggalJatuhTempo) : current[8];
  current[9] = data.frekuensi !== undefined ? data.frekuensi : current[9];
  current[10] = data.cicilanKe !== undefined ? parseInt(data.cicilanKe, 10) || 1 : current[10];
  current[11] = sudahDibayar;
  current[12] = sisa;
  current[13] = sisa <= 0 ? 'Selesai' : 'Aktif';
  current[14] = data.catatan !== undefined ? data.catatan : current[14];
  current[17] = new Date();

  sh.getRange(rowIndex, 1, 1, current.length).setValues([current]);

  return {
    success: true,
    message: 'Cicilan berhasil diperbarui.',
    data: cicilanRowToObject_(current)
  };
}

function ensureCicilanTransaksiColumn_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Transaksi');

  if (!sh) {
    sh = ss.insertSheet('Transaksi');
    sh.appendRow([
      'Tanggal',
      'Tipe',
      'Kategori',
      'Sub Kategori',
      'jumlah',
      'Rekening',
      'Dari/kepada',
      'Catatan',
      'UserKey',
      'ID'
    ]);
  } else if (sh.getLastColumn() < 10) {
    sh.getRange(1, 10).setValue('ID');
  } else if (!String(sh.getRange(1, 10).getValue() || '').trim()) {
    sh.getRange(1, 10).setValue('ID');
  }

  return sh;
}

function buatTransaksiPembayaranCicilan_(data) {
  const sh = ensureCicilanTransaksiColumn_();
  const transaksiId = 'TX-CIC-' + Utilities.getUuid().replace(/-/g, '').slice(0, 16);

  sh.appendRow([
    data.tanggal,
    "'" + data.tipe,
    'Hutang/Piutang',
    data.tipe === '+' ? 'Terima Piutang - Cicilan' : 'Bayar Hutang - Cicilan',
    Number(data.nominal),
    data.rekening,
    data.nama || '',
    (data.hutangPiutangID ? '[HP:' + data.hutangPiutangID + '] ' : '') + (data.catatan || ''),
    data.userKey,
    transaksiId
  ]);

  updateSaldoRekening(
    data.userKey,
    data.rekening,
    data.tipe,
    Number(data.nominal)
  );

  return transaksiId;
}

function hapusTransaksiPembayaranCicilan_(transaksiId, userKey) {
  if (!transaksiId) return false;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Transaksi');
  if (!sh || sh.getLastRow() < 2) return false;

  const lastCol = Math.max(10, sh.getLastColumn());
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const rowId = String(row[9] || '').trim();
    const rowUser = String(row[8] || '').trim();

    if (rowId !== String(transaksiId).trim() || rowUser !== String(userKey).trim()) continue;

    const tipe = String(row[1] || '').replace("'", '').trim();
    const jumlah = Number(row[4]) || 0;
    const rekening = String(row[5] || '').trim();

    sh.deleteRow(i + 2);

    if (jumlah > 0 && rekening) {
      const tipeBalik = tipe === '-' ? '+' : '-';
      updateSaldoRekening(userKey, rekening, tipeBalik, jumlah);
    }

    return true;
  }

  return false;
}

function hapusCicilan(id, userKey) {
  setupModulCicilan();

  const requestedId = normalizeCicilanKey_(id);
  const requestedUser = String(userKey || '').trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CICILAN_CONFIG.SHEET_CICILAN);
  const psh = ss.getSheetByName(CICILAN_CONFIG.SHEET_PEMBAYARAN);

  const rowIndex = findRowByNormalizedIdAndUser_(sh, requestedId, requestedUser, 1, 16);
  if (rowIndex < 2) throw new Error('Data cicilan tidak ditemukan.');

  const paymentRows = getDataRows_(psh).filter(r =>
    normalizeCicilanKey_(r[1]) === requestedId
  );

  const removedTransactions = [];

  try {
    paymentRows.forEach(payment => {
      const transaksiId = String(payment[10] || '').trim();
      if (transaksiId) {
        const removed = hapusTransaksiPembayaranCicilan_(transaksiId, requestedUser);
        if (removed) removedTransactions.push(transaksiId);
      }
    });

    const masterId = String(sh.getRange(rowIndex, 1).getValue() || '').trim();
    const hpId = String(sh.getRange(rowIndex, 2).getValue() || '').trim();
    sh.deleteRow(rowIndex);

    const rows = getDataRows_(psh);
    for (let i = rows.length - 1; i >= 0; i--) {
      if (normalizeCicilanKey_(rows[i][1]) === requestedId) {
        psh.deleteRow(i + 2);
      }
    }

    // Karena Cicilan adalah detail dari Hutang/Piutang, hapus master
    // yang menjadi induknya juga. Rekening tidak disentuh di sini;
    // semua pembayaran sudah dibalik melalui Transaksi sebelumnya.
    const hsh = ss.getSheetByName('HutangPiutang');
    if (hsh && hpId && hsh.getLastRow() >= 2) {
      const hRows = hsh.getRange(
        2, 1, hsh.getLastRow() - 1, Math.max(12, hsh.getLastColumn())
      ).getValues();

      for (let i = hRows.length - 1; i >= 0; i--) {
        const hOwner = String(hRows[i][8] || '').trim();
        const hId = String(hRows[i][0] || '').trim();
        const hCicilanId = String(hRows[i][10] || '').trim();

        if (hOwner === requestedUser &&
            (hId === hpId || normalizeCicilanKey_(hCicilanId) === requestedId)) {
          hsh.deleteRow(i + 2);
          break;
        }
      }
    }

    return {
      success: true,
      message: 'Cicilan ' + masterId + ', Hutang/Piutang terkait, seluruh riwayat pembayaran, dan transaksi terkait berhasil dihapus. Saldo rekening dikembalikan.'
    };
  } catch (err) {
    throw new Error('Gagal menghapus cicilan: ' + (err && err.message ? err.message : String(err)));
  }
}

function sinkronkanHutangPiutangDariCicilan_(cicilanId, userKey) {
  const id = String(cicilanId || '').trim();
  const owner = String(userKey || '').trim();
  if (!id || !owner) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const csh = ss.getSheetByName(CICILAN_CONFIG.SHEET_CICILAN);
  const hsh = ss.getSheetByName('HutangPiutang');
  if (!csh || !hsh || csh.getLastRow() < 2 || hsh.getLastRow() < 2) return;

  const crows = csh.getRange(2, 1, csh.getLastRow() - 1, Math.max(18, csh.getLastColumn())).getValues();
  let cicilanRow = null;
  for (let i = 0; i < crows.length; i++) {
    if (String(crows[i][0] || '').trim() === id && String(crows[i][15] || '').trim() === owner) {
      cicilanRow = crows[i];
      break;
    }
  }
  if (!cicilanRow) return;

  const hpId = String(cicilanRow[1] || '').trim();
  if (!hpId) return;

  const hrows = hsh.getRange(2, 1, hsh.getLastRow() - 1, Math.max(11, hsh.getLastColumn())).getValues();
  for (let i = 0; i < hrows.length; i++) {
    if (String(hrows[i][0] || '').trim() !== hpId || String(hrows[i][8] || '').trim() !== owner) continue;

    const totalHP = Number(hrows[i][3]) || 0;
    const totalCicilan = Number(cicilanRow[4]) || 0;
    const sudahDibayarCicilan = Number(cicilanRow[11]) || 0;
    const pembayaranSebelumCicilan = Math.max(0, totalHP - totalCicilan);
    const terbayarGabungan = Math.min(totalHP, pembayaranSebelumCicilan + sudahDibayarCicilan);
    const status = terbayarGabungan >= totalHP ? 'LUNAS' : 'BELUM LUNAS';

    hsh.getRange(i + 2, 5).setValue(terbayarGabungan);
    hsh.getRange(i + 2, 7).setValue(status);
    hsh.getRange(i + 2, 10).setValue('CICILAN');
    hsh.getRange(i + 2, 11).setValue(id);
    return;
  }
}

function catatPembayaranCicilan(data) {
  setupModulCicilan();

  const cicilanIdRaw = String(data && data.cicilanId || '').trim();
  const userKey = String(data && data.userKey || '').trim();
  const nominal = toNumber_(data && data.nominal);
  const rekening = String(data && data.rekening || '').trim();

  if (!cicilanIdRaw || !userKey) throw new Error('CicilanID dan UserKey wajib diisi.');
  if (nominal <= 0) throw new Error('Nominal pembayaran harus lebih dari 0.');
  if (!rekening) throw new Error('Rekening transaksi wajib dipilih.');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  let transaksiId = '';
  let transaksiCreated = false;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(CICILAN_CONFIG.SHEET_CICILAN);
    const psh = ss.getSheetByName(CICILAN_CONFIG.SHEET_PEMBAYARAN);

    const requestedId = normalizeCicilanKey_(cicilanIdRaw);
    const rowIndex = findRowByNormalizedIdAndUser_(sh, requestedId, userKey, 1, 16);
    if (rowIndex < 2) throw new Error('Cicilan tidak ditemukan atau bukan milik akun ini.');

    const row = sh.getRange(rowIndex, 1, 1, CICILAN_CONFIG.HEADER_CICILAN.length).getValues()[0];
    const tipe = String(row[3] || '').trim().toUpperCase();
    const arahTransaksi = tipe === 'PIUTANG' ? '+' : '-';

    if (tipe !== 'HUTANG' && tipe !== 'PIUTANG') {
      throw new Error('Jenis cicilan tidak valid. Pilih HUTANG atau PIUTANG.');
    }

    const rekeningList = typeof getRekeningData === 'function'
      ? getRekeningData(userKey)
      : [];
    const rekeningObj = rekeningList.find(r =>
      String(r.nama || '').trim().toLowerCase() === rekening.toLowerCase()
    );

    if (!rekeningObj) {
      throw new Error('Rekening tidak ditemukan atau bukan milik akun ini.');
    }

    if (arahTransaksi === '-' && Number(rekeningObj.saldoAkhir || rekeningObj.saldoAwal || 0) < nominal) {
      throw new Error('Saldo rekening ' + rekening + ' tidak mencukupi untuk pembayaran ini.');
    }

    const total = toNumber_(row[4]);
    const sudahDibayarLama = toNumber_(row[11]);
    const sisaLama = Math.max(0, total - sudahDibayarLama);

    if (sisaLama <= 0) throw new Error('Cicilan ini sudah lunas.');
    if (nominal > sisaLama) throw new Error('Nominal pembayaran melebihi sisa cicilan.');

    const paymentRows = getDataRows_(psh).filter(r =>
      normalizeCicilanKey_(r[1]) === requestedId
    );
    const nomorPembayaran = paymentRows.length + 1;
    const tenor = Math.max(1, parseInt(row[6], 10) || 1);
    const cicilanKeBaru = Math.min(tenor, Math.max(1, nomorPembayaran));
    const sudahDibayarBaru = sudahDibayarLama + nominal;
    const sisaBaru = Math.max(0, total - sudahDibayarBaru);

    transaksiId = buatTransaksiPembayaranCicilan_({
      tanggal: parseDate_(data.tanggalBayar) || new Date(),
      tipe: arahTransaksi,
      nominal: nominal,
      rekening: rekening,
      nama: String(row[2] || 'Cicilan').trim(),
      catatan: data.catatan || '',
      userKey: userKey,
      hutangPiutangID: String(row[1] || '').trim()
    });
    transaksiCreated = true;

    row[10] = cicilanKeBaru;
    row[11] = sudahDibayarBaru;
    row[12] = sisaBaru;
    row[13] = sisaBaru <= 0 ? 'Selesai' : 'Aktif';
    row[17] = new Date();
    sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);

    psh.appendRow([
      'CP-' + Utilities.getUuid().replace(/-/g, '').slice(0, 16),
      row[0],
      row[1],
      parseDate_(data.tanggalBayar) || new Date(),
      nominal,
      cicilanKeBaru,
      data.catatan || '',
      userKey,
      new Date(),
      rekening,
      transaksiId
    ]);

    // Sinkronkan angka Terbayar/Sisa pada Hutang & Piutang agar widget
    // dashboard menunjukkan angka yang sama dengan menu Cicilan.
    sinkronkanHutangPiutangDariCicilan_(row[0], userKey);

    return {
      success: true,
      message: sisaBaru <= 0 ? 'Cicilan selesai.' : 'Pembayaran cicilan berhasil dicatat.',
      data: cicilanRowToObject_(row),
      rekening: rekening,
      transaksiId: transaksiId,
      tipe: tipe,
      arahTransaksi: arahTransaksi
    };
  } catch (err) {
    if (transaksiCreated && transaksiId) {
      try {
        hapusTransaksiPembayaranCicilan_(transaksiId, userKey);
      } catch (rollbackErr) {
        console.error('Rollback transaksi pembayaran gagal:', rollbackErr);
      }
    }
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function normalizeCicilanKey_(value) {
  return String(value == null ? '' : value)
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .trim()
    .replace(/\s+/g, '')
    .toUpperCase();
}

/**
 * Hapus satu transaksi pembayaran cicilan.
 *
 * Aman untuk multi-user:
 * - Payment ID harus ditemukan.
 * - Kepemilikan diverifikasi dari master Cicilan + UserKey sesi.
 * - Setelah payment dihapus, SudahDibayar, Sisa, Status,
 *   dan nomor CicilanKe dihitung ulang dari seluruh riwayat yang tersisa.
 */
function hapusPembayaranCicilan(paymentId, userKey) {
  setupModulCicilan();

  const requestedPaymentId = normalizeCicilanKey_(paymentId);
  const requestedUser = String(userKey || '').trim();
  if (!requestedPaymentId || !requestedUser) {
    throw new Error('ID transaksi dan UserKey wajib diisi.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const csh = ss.getSheetByName(CICILAN_CONFIG.SHEET_CICILAN);
    const psh = ss.getSheetByName(CICILAN_CONFIG.SHEET_PEMBAYARAN);

    const paymentRows = getDataRows_(psh);
    let paymentRowIndex = -1;
    let paymentRow = null;

    for (let i = 0; i < paymentRows.length; i++) {
      if (normalizeCicilanKey_(paymentRows[i][0]) === requestedPaymentId) {
        paymentRowIndex = i + 2;
        paymentRow = paymentRows[i];
        break;
      }
    }

    if (paymentRowIndex < 2 || !paymentRow) {
      throw new Error('Transaksi pembayaran tidak ditemukan.');
    }

    const cicilanId = String(paymentRow[1] || '').trim();
    if (!cicilanId) throw new Error('CicilanID pada transaksi tidak ditemukan.');

    const requestedCicilanId = normalizeCicilanKey_(cicilanId);
    const cicilanRows = getDataRows_(csh);
    let masterRowIndex = -1;
    let masterRow = null;

    for (let i = 0; i < cicilanRows.length; i++) {
      if (normalizeCicilanKey_(cicilanRows[i][0]) === requestedCicilanId) {
        masterRowIndex = i + 2;
        masterRow = cicilanRows[i];
        break;
      }
    }

    if (masterRowIndex < 2 || !masterRow) {
      throw new Error('Cicilan utama untuk transaksi ini tidak ditemukan.');
    }

    const ownerKey = String(masterRow[15] || '').trim();
    if (!ownerKey || ownerKey !== requestedUser) {
      throw new Error('Akses menghapus transaksi tidak diizinkan.');
    }

    const transaksiId = String(paymentRow[10] || '').trim();
    if (transaksiId) {
      const removed = hapusTransaksiPembayaranCicilan_(transaksiId, requestedUser);
      if (!removed) throw new Error('Transaksi keuangan terkait tidak ditemukan. Pembayaran tidak dihapus agar saldo tetap aman.');
    }

    psh.deleteRow(paymentRowIndex);

    const remainingRows = getDataRows_(psh)
      .map((r, i) => ({ rowIndex: i + 2, data: r }))
      .filter(x => normalizeCicilanKey_(x.data[1]) === requestedCicilanId);

    const total = toNumber_(masterRow[4]);
    const tenor = Math.max(1, parseInt(masterRow[6], 10) || 1);
    const sudahDibayar = remainingRows.reduce((sum, x) => sum + toNumber_(x.data[4]), 0);
    const sisa = Math.max(0, total - sudahDibayar);
    const cicilanKe = remainingRows.length > 0 ? Math.min(tenor, remainingRows.length) : 1;

    remainingRows.forEach((x, index) => {
      const nomor = Math.min(tenor, index + 1);
      psh.getRange(x.rowIndex, 6).setValue(nomor);
    });

    masterRow[10] = cicilanKe;
    masterRow[11] = sudahDibayar;
    masterRow[12] = sisa;
    masterRow[13] = sisa <= 0 ? 'Selesai' : 'Aktif';
    masterRow[17] = new Date();
    csh.getRange(masterRowIndex, 1, 1, masterRow.length).setValues([masterRow]);

    // Setelah pembayaran dihapus, kembalikan angka Terbayar/Sisa di
    // Hutang & Piutang ke kondisi yang sesuai dengan riwayat Cicilan.
    sinkronkanHutangPiutangDariCicilan_(masterRow[0], requestedUser);

    return {
      success: true,
      message: 'Transaksi pembayaran berhasil dihapus dan saldo rekening dikembalikan.',
      deletedPaymentId: paymentRow[0],
      cicilanId: masterRow[0],
      transaksiId: transaksiId,
      data: cicilanRowToObject_(masterRow)
    };
  } finally {
    lock.releaseLock();
  }
}

function ambilRiwayatPembayaranCicilan(cicilanId, userKey) {
  setupModulCicilan();

  const requestedId = normalizeCicilanKey_(cicilanId);
  const requestedUser = String(userKey || '').trim();
  if (!requestedId) throw new Error('CicilanID wajib diisi.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const csh = ss.getSheetByName(CICILAN_CONFIG.SHEET_CICILAN);
  const psh = ss.getSheetByName(CICILAN_CONFIG.SHEET_PEMBAYARAN);

  const cicilanRows = getDataRows_(csh);
  const cicilan = cicilanRows.find(r => normalizeCicilanKey_(r[0]) === requestedId);
  if (!cicilan) throw new Error('Cicilan tidak ditemukan. ID: ' + String(cicilanId));

  const ownerKey = String(cicilan[15] || '').trim();
  if (!ownerKey) throw new Error('Pemilik cicilan tidak ditemukan.');
  if (requestedUser && requestedUser !== ownerKey) {
    throw new Error('Akses riwayat cicilan tidak diizinkan.');
  }

  const rows = getDataRows_(psh);
  return rows
    .filter(r => normalizeCicilanKey_(r[1]) === requestedId)
    .map(r => ({
      id: r[0],
      cicilanId: r[1],
      hutangPiutangId: r[2],
      // Google Apps Script tidak aman mengirim objek Date mentah
      // melalui google.script.run. Ubah menjadi string sebelum dikirim ke UI.
      tanggalBayar: serializeDateForClient_(r[3]),
      nominal: r[4],
      cicilanKe: r[5],
      catatan: r[6],
      userKey: r[7],
      createdAt: serializeDateForClient_(r[8]),
      rekening: r[9] || '',
      transaksiId: r[10] || ''
    }));
}

/**
 * Diagnostik aman untuk memastikan ID cicilan yang dikirim UI cocok dengan
 * data di Sheet CicilanPembayaran. Tidak mengubah data apa pun.
 */
function diagnosaRiwayatPembayaranCicilan(cicilanId, userKey) {
  setupModulCicilan();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const csh = ss.getSheetByName(CICILAN_CONFIG.SHEET_CICILAN);
  const psh = ss.getSheetByName(CICILAN_CONFIG.SHEET_PEMBAYARAN);

  const requestedId = normalizeCicilanKey_(cicilanId);
  const requestedUser = String(userKey || '').trim();

  const cRows = getDataRows_(csh);
  const pRows = getDataRows_(psh);

  const cicilan = cRows.find(r => normalizeCicilanKey_(r[0]) === requestedId);
  const directMatches = pRows.filter(r => normalizeCicilanKey_(r[1]) === requestedId);

  const ownerKey = cicilan ? String(cicilan[15] || '').trim() : '';
  const paymentUserKeys = directMatches.map(r => String(r[7] || '').trim());

  return {
    success: true,
    requestedId: requestedId,
    requestedUser: requestedUser,
    cicilanFound: !!cicilan,
    cicilanOwner: ownerKey,
    userAuthorized: !!cicilan && (!requestedUser || requestedUser === ownerKey),
    pembayaranRowsTotal: pRows.length,
    pembayaranMatches: directMatches.length,
    paymentIds: directMatches.map(r => r[0]),
    paymentCicilanIds: directMatches.map(r => r[1]),
    paymentUserKeys: paymentUserKeys,
    paymentUserKeyMismatchCount: ownerKey
      ? paymentUserKeys.filter(k => k && k !== ownerKey).length
      : 0
  };
}

/**
 * Perbaikan satu kali untuk data pembayaran yang sudah ada sebelum logika
 * nomor cicilan diperbaiki. Menomori riwayat berdasarkan urutan tanggal/row.
 * Jalankan manual sekali jika data lama masih menunjukkan CicilanKe = 1.
 */
function perbaikiDataCicilanPembayaran() {
  setupModulCicilan();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const csh = ss.getSheetByName(CICILAN_CONFIG.SHEET_CICILAN);
  const psh = ss.getSheetByName(CICILAN_CONFIG.SHEET_PEMBAYARAN);
  const cRows = getDataRows_(csh);
  const pRows = getDataRows_(psh);

  const ownerById = {};
  const tenorById = {};
  const masterRowById = {};

  cRows.forEach((r, i) => {
    const id = normalizeCicilanKey_(r[0]);
    if (!id) return;

    ownerById[id] = String(r[15] || '').trim();
    tenorById[id] = Math.max(1, parseInt(r[6], 10) || 1);
    masterRowById[id] = i + 2;
  });

  const groups = {};
  pRows.forEach((r, i) => {
    const id = normalizeCicilanKey_(r[1]);
    if (!id || !ownerById[id]) return;

    if (!groups[id]) groups[id] = [];
    groups[id].push({
      rowIndex: i + 2,
      date: r[3] ? new Date(r[3]).getTime() : 0,
      nominal: toNumber_(r[4])
    });
  });

  Object.keys(groups).forEach(id => {
    groups[id].sort((a, b) => a.date - b.date || a.rowIndex - b.rowIndex);

    const tenor = tenorById[id] || 1;
    const owner = ownerById[id];

    groups[id].forEach((item, index) => {
      const ke = Math.min(tenor, index + 1);

      // Normalisasi relasi pembayaran ke master Cicilan.
      psh.getRange(item.rowIndex, 2).setValue(
        csh.getRange(masterRowById[id], 1).getValue()
      );
      psh.getRange(item.rowIndex, 6).setValue(ke);
      psh.getRange(item.rowIndex, 8).setValue(owner);
    });

    const cicilanRowIndex = masterRowById[id];
    if (cicilanRowIndex >= 2) {
      const current = csh.getRange(
        cicilanRowIndex,
        1,
        1,
        CICILAN_CONFIG.HEADER_CICILAN.length
      ).getValues()[0];

      const total = toNumber_(current[4]);
      const sudah = groups[id].reduce((sum, item) => sum + item.nominal, 0);
      const sisa = Math.max(0, total - sudah);

      current[10] = Math.min(tenor, Math.max(1, groups[id].length));
      current[11] = sudah;
      current[12] = sisa;
      current[13] = sisa <= 0 ? 'Selesai' : 'Aktif';
      current[17] = new Date();

      csh.getRange(cicilanRowIndex, 1, 1, current.length).setValues([current]);
    }
  });

  return {
    success: true,
    message: 'Data pembayaran, ID, UserKey, nomor cicilan, dan saldo cicilan berhasil disinkronkan.'
  };
}

function cekStatusCicilan(userKey, referenceDate) {
  const list = ambilCicilan(userKey);
  const today = startOfDay_(referenceDate ? parseDate_(referenceDate) : new Date());

  return list.map(item => {
    const due = startOfDay_(item.tanggalJatuhTempo);
    let statusWaktu = 'BelumJatuhTempo';
    let hari = Math.round((due - today) / 86400000);

    if (item.status === 'Selesai' || item.sisa <= 0) {
      statusWaktu = 'Selesai';
    } else if (hari < 0) {
      statusWaktu = 'Terlambat';
      hari = Math.abs(hari);
    } else if (hari === 0) {
      statusWaktu = 'JatuhTempoHariIni';
    } else {
      statusWaktu = 'AkanJatuhTempo';
    }

    return Object.assign({}, item, {
      statusWaktu: statusWaktu,
      selisihHari: hari
    });
  });
}

function cicilanRowToObject_(r) {
  return {
    id: r[0],
    hutangPiutangID: r[1],
    nama: r[2],
    tipe: r[3],
    total: r[4],
    nominalCicilan: r[5],
    tenor: r[6],
    tanggalMulai: serializeDateForClient_(r[7]),
    tanggalJatuhTempo: serializeDateForClient_(r[8]),
    frekuensi: r[9],
    cicilanKe: r[10],
    sudahDibayar: r[11],
    sisa: r[12],
    status: r[13],
    catatan: r[14],
    userKey: r[15],
    createdAt: serializeDateForClient_(r[16]),
    updatedAt: serializeDateForClient_(r[17])
  };
}


function serializeDateForClient_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    if (isNaN(value.getTime())) return '';
    return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
  }
  return String(value);
}

function getDataRows_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
}

function findRowByNormalizedIdAndUser_(sh, normalizedId, userKey, idCol, userCol) {
  if (!sh || sh.getLastRow() < 2) return -1;

  const rows = getDataRows_(sh);
  for (let i = 0; i < rows.length; i++) {
    if (normalizeCicilanKey_(rows[i][idCol - 1]) === String(normalizedId) &&
        String(rows[i][userCol - 1]).trim() === String(userKey).trim()) {
      return i + 2;
    }
  }
  return -1;
}

function findRowByIdAndUser_(sh, id, userKey, idCol, userCol) {
  if (!sh || sh.getLastRow() < 2) return -1;

  const rows = getDataRows_(sh);
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][idCol - 1]) === String(id) &&
        String(rows[i][userCol - 1]) === String(userKey)) {
      return i + 2;
    }
  }
  return -1;
}

function toNumber_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : 0;
  if (value === null || value === undefined || value === '') return 0;

  let s = String(value).trim().replace(/[^\d,.-]/g, '');
  if (s.indexOf(',') >= 0 && s.indexOf('.') >= 0) {
    if (s.lastIndexOf(',') > s.lastIndexOf('.')) {
      s = s.replace(/\./g, '').replace(',', '.');
    } else {
      s = s.replace(/,/g, '');
    }
  } else if (s.indexOf(',') >= 0) {
    s = s.replace(',', '.');
  }

  const n = Number(s);
  return isFinite(n) ? n : 0;
}

function parseDate_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]') return value;

  const d = new Date(value);
  if (isNaN(d.getTime())) throw new Error('Tanggal tidak valid: ' + value);
  return d;
}

function startOfDay_(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  return d;
}


/**
 * ============================================================
 * TES FINAL INTEGRITAS MODUL CICILAN
 * ============================================================
 * READ-ONLY. Tidak mengubah data.
 *
 * Memeriksa:
 * 1. Master Cicilan ditemukan.
 * 2. Semua payment terhubung berdasarkan CicilanID.
 * 3. UserKey master dan payment.
 * 4. Total pembayaran vs SudahDibayar.
 * 5. Sisa vs Total - SudahDibayar.
 * 6. Nomor Cicilan berurutan.
 * ============================================================
 */
