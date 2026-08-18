// ============================================================
// CATATAN KEUANGAN - BACKEND MODULE
// Generated from the verified 7F-16 baseline.
// Functions are preserved in full; only file ownership was reorganized.
// ============================================================

function mintaIzinEmail() {
  var email = Session.getEffectiveUser().getEmail();
  MailApp.sendEmail(email, '[Otorisasi Catatan Keuangan]', 'Selamat! Otorisasi izin pengiriman email telah berhasil.');
  Logger.log('Otorisasi Berhasil! Email uji coba dikirim ke: ' + email);
}

function formatDate(d) {
  if (!d) return '';

  if (typeof d === 'string') {
    var str = d.trim();
    var match = str.match(/^(\d{4}-\d{2}-\d{2})/);
    if (match) return match[1];
  }

  if (Object.prototype.toString.call(d) === '[object Date]' || d instanceof Date) {
    if (isNaN(d.getTime())) return '';
    var tz = 'Asia/Jakarta';
    try {
      tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    } catch (e) {
      try {
        tz = Session.getScriptTimeZone();
      } catch (err) {}
    }
    return Utilities.formatDate(d, tz, "yyyy-MM-dd");
  }

  var strFallback = String(d).trim();
  var matchFallback = strFallback.match(/^(\d{4}-\d{2}-\d{2})/);
  if (matchFallback) return matchFallback[1];

  return strFallback;
}
