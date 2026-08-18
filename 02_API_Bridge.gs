// ============================================================
// CATATAN KEUANGAN - BACKEND MODULE
// Generated from the verified 7F-16 baseline.
// Functions are preserved in full; only file ownership was reorganized.
// ============================================================

var API_FUNCTIONS_ = {
  cekLogin: cekLogin, daftarPenggunaBaru: daftarPenggunaBaru, validasiUserKey: validasiUserKey,
  simpanPengaturanPengguna: simpanPengaturanPengguna, getPengaturanPengguna: getPengaturanPengguna,
  getDaftarPenggunaPending: getDaftarPenggunaPending, setujuiPengguna: setujuiPengguna, tolakPengguna: tolakPengguna,
  getDaftarPengguna: getDaftarPengguna, tambahPengguna: tambahPengguna, hapusPengguna: hapusPengguna, cekApakahAdmin: cekApakahAdmin,
  getRekeningData: getRekeningData, simpanTransfer: simpanTransfer, simpanRekeningBaru: simpanRekeningBaru,
  hapusRekening: hapusRekening, updateRekeningDetail: updateRekeningDetail, getKategoriData: getKategoriData,
  getIkhtisarData: getIkhtisarData, getRiwayatTransaksi: getRiwayatTransaksi, getRiwayatTransfer: getRiwayatTransfer,
  getBreakdownKategoriBulanan: getBreakdownKategoriBulanan, getRiwayatPeriode: getRiwayatPeriode,
  getHutangPiutangData: getHutangPiutangData, updateHutangPiutangDetail: updateHutangPiutangDetail,
  simpanHutangPiutangBaru: simpanHutangPiutangBaru, bayarHutangPiutang: bayarHutangPiutang, hapusHutangPiutang: hapusHutangPiutang,
  getGoalsData: getGoalsData, updateGoalDetail: updateGoalDetail, simpanGoalBaru: simpanGoalBaru, setorKeGoal: setorKeGoal,
  tarikDariGoal: tarikDariGoal, hapusGoal: hapusGoal, hapusTransfer: hapusTransfer, hapusTransaksi: hapusTransaksi,
  updateKategoriDetail: updateKategoriDetail, hapusKategori: hapusKategori, simpanKategoriBaru: simpanKategoriBaru,
  updateTransaksi: updateTransaksi, simpanData: simpanData, buatCicilan: buatCicilan,
  catatPembayaranCicilan: catatPembayaranCicilan, diagnosaRiwayatPembayaranCicilan: diagnosaRiwayatPembayaranCicilan,
  resetSemuaDataAkun: resetSemuaDataAkun
};

function apiHealth_() {
  return ContentService.createTextOutput(JSON.stringify({ok:true, service:'catatan-keuangan-api', version:1}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) return apiJson_({ok:false, error:'Body request kosong.'});
    var payload = JSON.parse(e.postData.contents);
    var action = String(payload.action || '').trim();
    var args = Array.isArray(payload.args) ? payload.args : [];
    if (!API_FUNCTIONS_[action]) return apiJson_({ok:false, error:'API action tidak diizinkan: '+action});
    var result = API_FUNCTIONS_[action].apply(null, args);
    return apiJson_({ok:true, result:result});
  } catch (err) {
    console.error('API bridge error:', err);
    return apiJson_({ok:false, error:err && err.message ? err.message : String(err)});
  }
}

function apiJson_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
