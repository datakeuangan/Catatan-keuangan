// ============================================================
// CATATAN KEUANGAN - BACKEND MODULE
// Generated from the verified 7F-16 baseline.
// Functions are preserved in full; only file ownership was reorganized.
// ============================================================

function doGet(e) {
  // API mode untuk frontend GitHub. Mode HTML lama tetap dipakai jika
  // parameter api tidak dikirim, sehingga aplikasi Apps Script lama tidak rusak.
  if (e && e.parameter && e.parameter.api === '1') {
    return apiHealth_();
  }

  return HtmlService.createTemplateFromFile('Form')
    .evaluate()
    .setTitle('Catatan Keuangan')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
