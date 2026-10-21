/**
 * OVVO — Coming Soon sign-ups
 * Apps Script project owned by info@ovvoshop.com, writing to the sheet below.
 * Each sign-up from ovvoshop.com is added as a row in the "Signups" tab
 * and an email is sent to info@ovvoshop.com.
 */

const NOTIFY_TO = 'info@ovvoshop.com';
const SPREADSHEET_ID = '1-nd0VSgik5JMslcZ8aFEgZjYK2PKsGK__hlb7_lEzBw'; // "OVVO – Coming Soon Sign-ups"
const SHEET_NAME = 'Signups';
const HEADERS = ['التاريخ / Date', 'الاسم / Name', 'الموبايل / Mobile', 'الإيميل / Email', 'اللغة / Lang', 'الصفحة / Page', 'الجهاز / Device'];

function doPost(e) {
  const p = (e && e.parameter) || {};

  // Spam trap: real visitors never fill this hidden field.
  if (p.company) return json_({ ok: true });

  const name  = clean_(p.name, 80);
  const phone = clean_(p.phone, 20);
  const email = clean_(p.email, 120);
  if (name.length < 2 || phone.replace(/\D/g, '').length < 8) return json_({ ok: false });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet_();
    // Leading apostrophe keeps the 0/+ at the start of phone numbers and blocks formulas.
    sh.appendRow([new Date(), safe_(name), "'" + phone, safe_(email), clean_(p.lang, 5), clean_(p.page, 200), clean_(p.ua, 150)]);
  } finally {
    lock.releaseLock();
  }

  const total = Math.max(getSheet_().getLastRow() - 1, 0);
  const wa = phone.replace(/\D/g, '').replace(/^0/, '962');
  MailApp.sendEmail({
    to: NOTIFY_TO,
    replyTo: email || NOTIFY_TO,
    subject: 'OVVO – تسجيل جديد: ' + name,
    htmlBody:
      '<div dir="rtl" style="font-family:Arial,sans-serif;font-size:15px;color:#0D1226">' +
      '<h2 style="color:#791126;margin:0 0 12px">تسجيل جديد من صفحة Coming Soon</h2>' +
      '<p><b>الاسم:</b> ' + esc_(name) + '</p>' +
      '<p><b>الموبايل:</b> <span dir="ltr">' + esc_(phone) + '</span> — ' +
      '<a href="https://wa.me/' + wa + '">واتساب</a></p>' +
      '<p><b>الإيميل:</b> ' + (email ? esc_(email) : '—') + '</p>' +
      '<p style="color:#666;font-size:13px">مجموع المسجلين لحد الآن: ' + total + '</p>' +
      '<p><a href="' + SpreadsheetApp.openById(SPREADSHEET_ID).getUrl() + '">افتح جدول التسجيلات</a></p>' +
      '</div>'
  });

  return json_({ ok: true });
}

function doGet() {
  return json_({ ok: true, service: 'OVVO sign-ups' });
}

/* ---------- helpers ---------- */
function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#0D1226').setFontColor('#F4F4F4');
    sh.setColumnWidths(1, HEADERS.length, 170);
  }
  return sh;
}
function clean_(v, max) { return String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').trim().slice(0, max); }
function safe_(v) { return /^[=+\-@]/.test(v) ? "'" + v : v; }
function esc_(v) { return String(v).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/** Run once from the editor (▶ Run → setup) to create the tab and approve permissions. */
function setup() { getSheet_(); }
