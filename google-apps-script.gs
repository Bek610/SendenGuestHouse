/**
 * Senden Guest House — захиалгыг Google Sheets руу бичиж, имэйлээр мэдэгдэх,
 * захиалгатай өдрүүдийг сайтад харуулж, давхар захиалгаас хамгаалах скрипт.
 *
 * Шинэчлэх заавар (хуучин скриптийг солих):
 * 1. Google Sheet → Extensions → Apps Script. Байгаа кодыг бүгдийг устгаад энэ файлыг хуулж тавина. Save.
 * 2. Deploy → Manage deployments → харандаа (Edit) → Version: "New version" → Deploy.
 *    (URL өөрчлөгдөхгүй. Шинэ зөвшөөрөл (Google Calendar) асуувал Allow дарна.)
 *
 * Хэрэглэх:
 * - "Төлөв" баганыг "Цуцалсан" болговол тэр өдрүүд сайт дээр дахин сул болно.
 * - Утсаар авсан захиалгыг Sheet-д шинэ мөр болгон гараар нэмнэ:
 *   Ирэх, Гарах баганад огноог 2026-12-20 хэлбэрээр бичнэ. Тэгвэл сайт дээр хаагдана.
 */

const SHEET_NAME = 'Захиалга';
const NOTIFY_EMAIL = 'sendenguesthouse@gmail.com';
const ADD_TO_CALENDAR = true;   // Захиалга бүрийг Google Calendar-т нэмэх эсэх
const HEADERS = ['Огноо', 'Төлөв', 'Нэр', 'Утас', 'Өрөө', 'Ирэх', 'Гарах', 'Шөнө',
                 'Хүн', 'Нийт', 'Аялал', 'Аяллын өдөр', 'Тэмдэглэл', 'Хэл'];
const COL_STATUS = 1, COL_IN = 5, COL_OUT = 6;   // 0-ээс эхэлсэн баганын дугаар

// Google Calendar-ын зөвшөөрөл өгөх: дээрээс 'authorizeCalendar'-ийг сонгоод Run дарна (нэг удаа).
function authorizeCalendar() {
  Logger.log('Хуанли: ' + CalendarApp.getDefaultCalendar().getName() + ' — зөвшөөрөл OK');
}

// Сайт захиалгатай өдрүүдийг эндээс уншина: GET ...exec?action=booked
function doGet(e) {
  try {
    return reply_({ ok: true, booked: bookedRanges_() });
  } catch (err) {
    return reply_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const d = JSON.parse(e.postData.contents);
    if (!d.name || !d.phone) return reply_({ ok: false, error: 'missing fields' });
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d.checkin) || !/^\d{4}-\d{2}-\d{2}$/.test(d.checkout) || d.checkin >= d.checkout)
      return reply_({ ok: false, error: 'bad dates' });

    lock.waitLock(20000);   // Зэрэг ирсэн хоёр захиалгыг дарааллаар нь шалгана
    const clash = bookedRanges_().find(r => d.checkin < r.to && r.from < d.checkout);
    if (clash) return reply_({ ok: false, error: 'dates_taken', clash: clash });

    const sheet = getSheet_();
    const row = [
      new Date(), 'Шинэ', d.name, "'" + d.phone, d.room, d.checkin, d.checkout, d.nights,
      d.guests, d.total, d.tour, d.tourDate, d.note, d.lang
    ].map(clean_);
    sheet.appendRow(row);
    SpreadsheetApp.flush();
    lock.releaseLock();

    let calNote = '';
    if (ADD_TO_CALENDAR) {
      try {
        CalendarApp.getDefaultCalendar().createAllDayEvent(
          'Senden: ' + clean_(d.name) + ' (' + clean_(d.guests) + ' хүн)',
          parseDate_(d.checkin), parseDate_(d.checkout),
          { description: 'Утас: ' + d.phone + '\nНийт: ' + d.total + (d.note ? '\n' + d.note : '') });
      } catch (calErr) { calNote = String(calErr); /* Хуанли алдаа гаргавал захиалга хэвээрээ хадгалагдана */ }
    }

    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: 'Шинэ захиалга: ' + clean_(d.name) + ' (' + clean_(d.checkin) + ')',
      body:
        'Нэр: ' + d.name + '\n' +
        'Утас: ' + d.phone + '\n' +
        'Өрөө: ' + d.room + '\n' +
        'Ирэх: ' + d.checkin + '  Гарах: ' + d.checkout + ' (' + d.nights + ' шөнө)\n' +
        'Хүний тоо: ' + d.guests + '\n' +
        'Нийт: ' + d.total + '\n' +
        (d.tour ? 'Аялал: ' + d.tour + ' — ' + d.tourDate + '\n' : '') +
        (d.note ? 'Тэмдэглэл: ' + d.note + '\n' : '') +
        '\nБүх захиалга: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
    });

    return reply_({ ok: true, calendar: calNote || 'ok' });
  } catch (err) {
    return reply_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

// Цуцлагдаагүй, дуусаагүй захиалгуудын өдрүүд: [{from:'2026-12-20', to:'2026-12-22'}] (to = гарах өдөр)
function bookedRanges_() {
  const sh = getSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  return sh.getRange(2, 1, n, HEADERS.length).getValues()
    .filter(r => String(r[COL_STATUS]).trim() !== 'Цуцалсан')
    .map(r => ({ from: ymd_(r[COL_IN], tz), to: ymd_(r[COL_OUT], tz) }))
    .filter(r => r.from && r.to && r.from < r.to && r.to > today)
    .sort((a, b) => a.from < b.from ? -1 : 1);
}

function ymd_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  const m = String(v).trim().match(/^(\d{4})[-.\/](\d{1,2})[-.\/](\d{1,2})/);
  return m ? m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2) : '';
}

function parseDate_(s) {
  const p = s.split('-').map(Number);
  return new Date(p[0], p[1] - 1, p[2]);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#1C3D6E').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Шинэ', 'Баталгаажсан', 'Цуцалсан'], true).build();
    sh.getRange('B2:B').setDataValidation(rule);
  }
  return sh;
}

// Текстийг 500 тэмдэгтээр хязгаарлаж, томьёо болж ажиллахаас хамгаална.
function clean_(v) {
  if (v === undefined || v === null) return '';
  if (typeof v !== 'string') return v;
  v = v.slice(0, 500);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function reply_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
