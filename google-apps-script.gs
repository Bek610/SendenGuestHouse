/**
 * Senden Guest House — захиалгыг Google Sheets руу бичиж, имэйлээр мэдэгдэх скрипт.
 *
 * Тохируулах заавар:
 * 1. sendenguesthouse@gmail.com-оор нэвтэрч шинэ Google Sheet үүсгэнэ ("Senden захиалга").
 * 2. Extensions → Apps Script. Тэнд байгаа кодыг устгаад энэ файлыг бүхэлд нь хуулж тавина. Save.
 * 3. Deploy → New deployment → төрөл: Web app.
 *      Execute as: Me (sendenguesthouse@gmail.com)
 *      Who has access: Anyone
 *    Deploy → Authorize access → Google бүртгэлээ сонгож зөвшөөрнө.
 * 4. Гарсан "Web app URL" (https://script.google.com/macros/s/.../exec)-ийг хуулж Claude-д өгнө.
 */

const SHEET_NAME = 'Захиалга';
const NOTIFY_EMAIL = 'sendenguesthouse@gmail.com';
const HEADERS = ['Огноо', 'Төлөв', 'Нэр', 'Утас', 'Өрөө', 'Ирэх', 'Гарах', 'Шөнө',
                 'Хүн', 'Нийт (₮)', 'Аялал', 'Аяллын өдөр', 'Тэмдэглэл', 'Хэл'];

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (!d.name || !d.phone) return reply_({ ok: false, error: 'missing fields' });

    const sheet = getSheet_();
    const row = [
      new Date(), 'Шинэ', d.name, d.phone, d.room, d.checkin, d.checkout, d.nights,
      d.guests, d.total, d.tour, d.tourDate, d.note, d.lang
    ].map(clean_);
    sheet.appendRow(row);

    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: 'Шинэ захиалга: ' + clean_(d.name) + ' (' + clean_(d.checkin) + ')',
      body:
        'Нэр: ' + d.name + '\n' +
        'Утас: ' + d.phone + '\n' +
        'Өрөө: ' + d.room + '\n' +
        'Ирэх: ' + d.checkin + '  Гарах: ' + d.checkout + ' (' + d.nights + ' шөнө)\n' +
        'Хүний тоо: ' + d.guests + '\n' +
        'Нийт: ' + d.total + '₮\n' +
        (d.tour ? 'Аялал: ' + d.tour + ' — ' + d.tourDate + '\n' : '') +
        (d.note ? 'Тэмдэглэл: ' + d.note + '\n' : '') +
        '\nБүх захиалга: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
    });

    return reply_({ ok: true });
  } catch (err) {
    return reply_({ ok: false, error: String(err) });
  }
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#1d5a42').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    // "Төлөв" баганад сонголт: Шинэ / Баталгаажсан / Цуцалсан
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
