/**
 * Senden Guest House — захиалгын систем (Google Sheets + Apps Script)
 *
 * Юу хийдэг вэ:
 *  1. Сайтаас ирсэн захиалгыг "Захиалга" хуудсанд мөр болгон бичиж, танд имэйл илгээнэ.
 *     Тэр өдрүүд сайт дээр шууд хаагдана.
 *  2. Зочин данс руу урьдчилгаа шилжүүлээд "Шилжүүлсэн" дарахад танд "[Төлбөр шалгах]" имэйл ирнэ.
 *  3. Та дансаа шалгаад sendenguesthouse.com/admin хуудсанд "Төлбөр орсон" дарахад
 *     Sheet дээр "Баталгаажсан" болж, зочинд имэйл очно (имэйл бичсэн бол).
 *  4. Booking.com / Trip.com-ын iCal холбоосыг ICAL_FEEDS-д тавьбал тэндээс ирсэн захиалгууд
 *     15 минут тутамд Sheet-д орж, сайт дээр тэр өдрүүд хаагдана.
 *  5. ...exec?action=ical холбоосыг Booking.com-д "Import calendar"-аар өгвөл
 *     манай сайтын захиалгууд тэнд хаагдана.
 *
 * Суулгах / шинэчлэх (нэг удаа):
 *  1. Google Sheet → Extensions → Apps Script. Байгаа кодыг бүгдийг устгаад энэ файлыг хуулж тавина. Save.
 *  2. Project Settings (араа) → Script Properties → Add script property:
 *       ADMIN_PASSWORD = (8-аас дээш тэмдэгттэй нууц үг)  ← админ хуудсанд нэвтрэх нууц үг
 *  3. Доорх BANK хэсэгт дансаа бичнэ.
 *  4. Дээрээс 'setup' функцийг сонгоод Run → Allow (зөвшөөрөл). Энэ нь баганыг шинэчилж,
 *     Booking.com/Trip.com-той 15 минут тутам синк хийх таймерыг тохируулна.
 *  5. Deploy → Manage deployments → харандаа (Edit) → Version: "New version" → Deploy.
 *     (URL өөрчлөгдөхгүй.)
 *
 * Sheet-ийн төлөв:  Шинэ → Төлбөр шалгах → Баталгаажсан   (эсвэл Цуцалсан)
 *  - "Цуцалсан" болговол тэр өдрүүд сайт дээр дахин сул болно.
 *  - Утсаар авсан захиалгыг гараар нэмж болно: Ирэх, Гарах баганад 2026-12-20 хэлбэрээр бичнэ.
 */

const SHEET_NAME = 'Захиалга';
const NOTIFY_EMAIL = 'sendenguesthouse@gmail.com';
const SITE_URL = 'https://sendenguesthouse.com';
const ADD_TO_CALENDAR = true;          // Захиалга бүрийг Google Calendar-т нэмэх

const PRICE_USD = 50;                  // Нэг шөнө, бүтэн байр
const PRICE_MNT = 180000;
const TRANSFER_USD = 30;               // Тосох, үдэх эсвэл хоёулаа — нийт $30 (хоёр талдаа)
const DEPOSIT_MNT = 180000;            // Монгол зочдын урьдчилгаа (1 шөнө)

// Монгол зочид урьдчилгаагаа шилжүүлэх данс. Хоосон бол сайт дансны мэдээлэл харуулахгүй.
const BANK = { name: '', account: '', holder: '' };

// Booking.com / Trip.com-ын "Export calendar" (iCal, .ics) холбоос. Хоосон бол алгасна.
const ICAL_FEEDS = [
  { name: 'Booking.com', url: '' },
  { name: 'Trip.com', url: '' }
];

// Багана (0-ээс эхэлсэн). Хуучин 14 багана хэвээр, шинэ нь араас нь нэмэгдэнэ.
const HEADERS = ['Огноо', 'Төлөв', 'Нэр', 'Утас', 'Өрөө', 'Ирэх', 'Гарах', 'Шөнө',
                 'Хүн', 'Нийт', 'Аялал', 'Аяллын өдөр', 'Тэмдэглэл', 'Хэл',
                 'Код', 'Имэйл', 'Тосох/үдэх', 'Төлбөр орсон', 'Эх сурвалж', 'Гадны ID'];
const C = { date: 0, status: 1, name: 2, phone: 3, room: 4, in: 5, out: 6, nights: 7, guests: 8,
            total: 9, tour: 10, tourDate: 11, note: 12, lang: 13, code: 14, email: 15,
            transfer: 16, paid: 17, source: 18, uid: 19 };
const ST = { NEW: 'Шинэ', CHECK: 'Төлбөр шалгах', OK: 'Баталгаажсан', CANCEL: 'Цуцалсан' };
const SITE_SOURCE = 'Сайт';
const TRANSFERS = { pick: ['Тосох', 1], drop: ['Үдэх', 1], both: ['Тосох ба үдэх', 1] };

/* ---------------- нэг удаагийн тохиргоо ---------------- */

function setup() {
  getSheet_();
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'syncOta')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('syncOta').timeBased().everyMinutes(15).create();
  if (ADD_TO_CALENDAR) CalendarApp.getDefaultCalendar();
  if (!adminPassword_()) Logger.log('АНХААР: Script Properties-д ADMIN_PASSWORD нэмээгүй байна.');
  syncOta();
  Logger.log('Бэлэн. iCal экспорт: ' + ScriptApp.getService().getUrl() + '?action=ical');
}

/* ---------------- GET ---------------- */

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'booked';
  try {
    if (action === 'ical') {
      return ContentService.createTextOutput(icalExport_()).setMimeType(ContentService.MimeType.ICAL);
    }
    return reply_({ ok: true, booked: bookedRanges_(), transferUsd: TRANSFER_USD });
  } catch (err) {
    return reply_({ ok: false, error: String(err) });
  }
}

/* ---------------- POST ---------------- */

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return reply_({ ok: false, error: 'bad json' }); }
  try {
    switch (d.action || 'book') {
      case 'book': return reply_(book_(d));
      case 'claim': return reply_(claim_(d));
      case 'admin_list': return reply_(adminList_(d));
      case 'admin_set': return reply_(adminSet_(d));
      default: return reply_({ ok: false, error: 'bad action' });
    }
  } catch (err) {
    return reply_({ ok: false, error: String(err) });
  }
}

/* ---------------- захиалга ---------------- */

function book_(d) {
  if (!d.name || !d.phone) return { ok: false, error: 'missing fields' };
  if (!isYmd_(d.checkin) || !isYmd_(d.checkout) || d.checkin >= d.checkout) return { ok: false, error: 'bad dates' };
  const guests = Math.min(Math.max(parseInt(d.guests, 10) || 1, 1), 5);
  const nights = Math.round((parseDate_(d.checkout) - parseDate_(d.checkin)) / 864e5);
  const tr = TRANSFERS[d.transfer] || null;
  const totalUsd = nights * PRICE_USD + (tr ? tr[1] * TRANSFER_USD : 0);
  const mn = !d.lang || d.lang === 'mn';   // ru/zh/ko/en зочид — гадаад, ирэхдээ төлнө

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);   // Зэрэг ирсэн хоёр захиалгыг дарааллаар нь шалгана
  let code;
  try {
    const clash = bookedRanges_().find(r => d.checkin < r.to && r.from < d.checkout);
    if (clash) return { ok: false, error: 'dates_taken', clash: clash };
    const sh = getSheet_();
    code = newCode_(sh);
    const row = new Array(HEADERS.length).fill('');
    row[C.date] = new Date(); row[C.status] = ST.NEW; row[C.name] = d.name; row[C.phone] = "'" + d.phone;
    row[C.room] = 'Бүтэн байр'; row[C.in] = d.checkin; row[C.out] = d.checkout; row[C.nights] = nights;
    row[C.guests] = guests; row[C.total] = '$' + totalUsd; row[C.tour] = d.tour || ''; row[C.tourDate] = d.tourDate || '';
    row[C.note] = d.note || ''; row[C.lang] = d.lang || 'mn'; row[C.code] = code; row[C.email] = d.email || '';
    row[C.transfer] = tr ? tr[0] + (d.flight ? ' — ' + d.flight : '') : ''; row[C.source] = SITE_SOURCE;
    sh.appendRow(row.map(clean_));
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  if (ADD_TO_CALENDAR) {
    try {
      CalendarApp.getDefaultCalendar().createAllDayEvent(
        'Senden: ' + clean_(d.name) + ' (' + guests + ' хүн) ' + code,
        parseDate_(d.checkin), parseDate_(d.checkout),
        { description: 'Утас: ' + d.phone + '\nНийт: $' + totalUsd + (tr ? '\n' + tr[0] : '') });
    } catch (calErr) { /* Хуанли алдаа гаргавал захиалга хэвээрээ хадгалагдана */ }
  }

  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    subject: 'Шинэ захиалга ' + code + ': ' + clean_(d.name) + ' (' + d.checkin + ')',
    body:
      'Код: ' + code + '\n' +
      'Нэр: ' + d.name + '\n' +
      'Утас: ' + d.phone + (d.email ? '\nИмэйл: ' + d.email : '') + '\n' +
      'Ирэх: ' + d.checkin + '  Гарах: ' + d.checkout + ' (' + nights + ' шөнө)\n' +
      'Хүний тоо: ' + guests + '\n' +
      'Нийт: $' + totalUsd + (tr ? ' (' + tr[0] + ' $' + tr[1] * TRANSFER_USD + ' орсон)' : '') + '\n' +
      (tr ? 'Тосох/үдэх: ' + tr[0] + (d.flight ? ' — ' + d.flight : '') + '\n' : '') +
      (d.tour ? 'Аялал: ' + d.tour + ' — ' + (d.tourDate || '') + '\n' : '') +
      (d.note ? 'Тэмдэглэл: ' + d.note + '\n' : '') +
      '\n' + (mn ? 'Монгол зочин — урьдчилгаа ' + DEPOSIT_MNT.toLocaleString('en-US') + '₮ дансаар.'
                 : 'Гадаад зочин — ирэхдээ бэлнээр төлнө. Баталгаажуулахаа мартуузай.') +
      '\n\nАдмин: ' + SITE_URL + '/admin\nSheet: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
  });

  const res = { ok: true, code: code, totalUsd: totalUsd, nights: nights };
  if (mn && BANK.account) res.bank = { name: BANK.name, account: BANK.account, holder: BANK.holder, amount: DEPOSIT_MNT };
  return res;
}

// Зочин "Шилжүүлсэн" дарав → төлөв "Төлбөр шалгах", танд имэйл
function claim_(d) {
  const code = String(d.code || '').toUpperCase();
  if (!/^SG\d{4}$/.test(code)) return { ok: false, error: 'bad code' };
  const cache = CacheService.getScriptCache(), key = 'claim:' + code;
  const tries = Number(cache.get(key) || 0);
  if (tries >= 5) return { ok: false, error: 'too many' };
  cache.put(key, String(tries + 1), 3600);

  const sh = getSheet_();
  const f = findByCode_(sh, code);
  if (!f || digits_(f.row[C.phone]).slice(-8) !== digits_(d.phone).slice(-8)) return { ok: false, error: 'not found' };
  if (f.row[C.status] === ST.NEW) {
    sh.getRange(f.index, C.status + 1).setValue(ST.CHECK);
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: '[Төлбөр шалгах] ' + code + ' — ' + f.row[C.name] + ' — ' + DEPOSIT_MNT.toLocaleString('en-US') + '₮',
      body: f.row[C.name] + ' (' + f.row[C.phone] + ') "' + code + '" гүйлгээний утгатай ' +
        DEPOSIT_MNT.toLocaleString('en-US') + '₮ шилжүүлсэн гэж мэдэгдлээ.\n' +
        'Ирэх: ' + ymd_(f.row[C.in]) + '  Гарах: ' + ymd_(f.row[C.out]) + '\n\n' +
        'Дансаа шалгаад админ хуудсанд "Төлбөр орсон" дарна уу:\n' + SITE_URL + '/admin'
    });
  }
  return { ok: true };
}

/* ---------------- админ ---------------- */

function adminList_(d) {
  const denied = adminGuard_(d.password);
  if (denied) return denied;
  const sh = getSheet_(), n = sh.getLastRow() - 1;
  const today = ymd_(new Date());
  if (n < 1) return { ok: true, items: [], ical: ScriptApp.getService().getUrl() + '?action=ical' };
  const rows = sh.getRange(2, 1, n, HEADERS.length).getValues();
  const items = rows
    .filter(r => r[C.status] !== ST.CANCEL && ymd_(r[C.out]) >= today)
    .map(r => ({
      code: String(r[C.code] || ''), status: String(r[C.status] || ST.NEW), name: String(r[C.name]),
      phone: String(r[C.phone]), email: String(r[C.email] || ''), checkin: ymd_(r[C.in]), checkout: ymd_(r[C.out]),
      nights: r[C.nights], guests: r[C.guests], total: String(r[C.total]), transfer: String(r[C.transfer] || ''),
      tour: String(r[C.tour] || ''), note: String(r[C.note] || ''), lang: String(r[C.lang] || ''),
      source: String(r[C.source] || SITE_SOURCE), paid: r[C.paid] ? ymd_(r[C.paid]) : ''
    }))
    .sort((a, b) => a.checkin < b.checkin ? -1 : 1);
  return { ok: true, items: items, depositMnt: DEPOSIT_MNT, ical: ScriptApp.getService().getUrl() + '?action=ical' };
}

function adminSet_(d) {
  const denied = adminGuard_(d.password);
  if (denied) return denied;
  const sh = getSheet_();
  const f = findByCode_(sh, String(d.code || '').toUpperCase());
  if (!f) return { ok: false, error: 'Захиалга олдсонгүй' };
  if (d.op === 'confirm') {
    sh.getRange(f.index, C.status + 1).setValue(ST.OK);
    sh.getRange(f.index, C.paid + 1).setValue(new Date());
    const email = String(f.row[C.email] || '');
    if (/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) {
      const en = !!f.row[C.lang] && f.row[C.lang] !== 'mn';
      MailApp.sendEmail({
        to: email, replyTo: NOTIFY_EMAIL, name: 'Senden Guest House',
        subject: en ? 'Your stay at Senden Guest House is confirmed (' + f.row[C.code] + ')'
                    : 'Senden Guest House — захиалга баталгаажлаа (' + f.row[C.code] + ')',
        body: en
          ? 'Hello ' + f.row[C.name] + ',\n\nYour booking is confirmed.\nCheck-in: ' + ymd_(f.row[C.in]) +
            '\nCheck-out: ' + ymd_(f.row[C.out]) + '\nTotal: ' + f.row[C.total] + ' (pay on arrival)\n\n' +
            'Address: 3rd floor above Ugooj store, Khugjil complex, Zuunmod, Töv Province\nPhone / WhatsApp: +976 9901 2835\n\nSee you soon!\nSenden Guest House'
          : 'Сайн байна уу, ' + f.row[C.name] + '.\n\nТаны төлбөр орж, захиалга баталгаажлаа.\nИрэх: ' + ymd_(f.row[C.in]) +
            '\nГарах: ' + ymd_(f.row[C.out]) + '\nНийт: ' + f.row[C.total] + '\n\n' +
            'Хаяг: Төв аймаг, Зуунмод сум, Хөгжил хотхон, Өгөөж дэлгүүрийн 3 давхарт\nУтас: 9901 2835\n\nТаныг хүлээж байна.\nSenden Guest House'
      });
    }
  } else if (d.op === 'cancel') {
    sh.getRange(f.index, C.status + 1).setValue(ST.CANCEL);
  } else {
    return { ok: false, error: 'bad op' };
  }
  return { ok: true };
}

function adminPassword_() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD') || '';
}

// Нууц үг таахаас хамгаалж, 15 минутад 20-оос олон буруу оролдлогыг хаана.
function adminGuard_(pw) {
  const want = adminPassword_();
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('admin-fail') || 0);
  if (fails >= 20) return { ok: false, error: 'Хэт олон буруу оролдлого. 15 минутын дараа дахин оролдоно уу.' };
  if (want.length >= 8 && String(pw || '') === want) return null;
  cache.put('admin-fail', String(fails + 1), 900);
  return { ok: false, error: want.length < 8 ? 'ADMIN_PASSWORD тохируулаагүй байна.' : 'Нууц үг буруу.' };
}

/* ---------------- Booking.com / Trip.com синк ---------------- */

// Таймер 15 минут тутамд ажиллуулна. Гараар Run дарж ч болно.
function syncOta() {
  const feeds = ICAL_FEEDS.filter(f => f.url);
  if (!feeds.length) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_(), today = ymd_(new Date());
    const n = sh.getLastRow() - 1;
    const rows = n > 0 ? sh.getRange(2, 1, n, HEADERS.length).getValues() : [];
    feeds.forEach(feed => {
      let events;
      try { events = parseIcal_(UrlFetchApp.fetch(feed.url, { muteHttpExceptions: true }).getContentText()); }
      catch (err) { Logger.log(feed.name + ': ' + err); return; }
      if (!events) return;   // Татаж чадаагүй бол юуг ч цуцлахгүй
      const seen = {};
      events.filter(ev => ev.to > today).forEach(ev => {
        const uid = feed.name + ':' + ev.uid;
        seen[uid] = true;
        const i = rows.findIndex(r => r[C.uid] === uid);
        if (i >= 0) {
          if (ymd_(rows[i][C.in]) !== ev.from || ymd_(rows[i][C.out]) !== ev.to) {
            sh.getRange(i + 2, C.in + 1, 1, 2).setValues([[ev.from, ev.to]]);
            mailOta_(feed.name + ' захиалга өөрчлөгдлөө', ev, feed.name);
          }
          return;
        }
        // Манай сайтын захиалгыг OTA буцааж "хаалттай" гэж илгээсэн бол давхар бичихгүй
        const own = rows.some(r => r[C.source] !== feed.name && r[C.status] !== ST.CANCEL &&
                                   ymd_(r[C.in]) === ev.from && ymd_(r[C.out]) === ev.to);
        if (own) return;
        const row = new Array(HEADERS.length).fill('');
        row[C.date] = new Date(); row[C.status] = ST.OK; row[C.name] = ev.summary || feed.name;
        row[C.room] = 'Бүтэн байр'; row[C.in] = ev.from; row[C.out] = ev.to;
        row[C.nights] = Math.round((parseDate_(ev.to) - parseDate_(ev.from)) / 864e5);
        row[C.source] = feed.name; row[C.uid] = uid; row[C.note] = feed.name + '-оор хийгдсэн захиалга';
        sh.appendRow(row.map(clean_));
        rows.push(row);
        mailOta_(feed.name + '-оос шинэ захиалга', ev, feed.name);
      });
      // OTA дээр цуцлагдсан захиалга → Sheet дээр Цуцалсан
      rows.forEach((r, i) => {
        if (r[C.source] === feed.name && r[C.status] !== ST.CANCEL && !seen[r[C.uid]] && ymd_(r[C.out]) > today) {
          sh.getRange(i + 2, C.status + 1).setValue(ST.CANCEL);
          r[C.status] = ST.CANCEL;
          mailOta_(feed.name + ' захиалга цуцлагдлаа', { from: ymd_(r[C.in]), to: ymd_(r[C.out]), summary: r[C.name] }, feed.name);
        }
      });
    });
  } finally {
    lock.releaseLock();
  }
}

function mailOta_(subject, ev, source) {
  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    subject: subject + ': ' + ev.from + ' – ' + ev.to,
    body: subject + '\nИрэх: ' + ev.from + '\nГарах: ' + ev.to + '\n' + (ev.summary || '') +
      '\n\nСайт дээр эдгээр өдрүүд хаагдсан. Дэлгэрэнгүйг ' + source + '-ийн удирдлагаас харна уу.\n' +
      SpreadsheetApp.getActiveSpreadsheet().getUrl()
  });
}

// .ics → [{uid, from, to, summary}]  (to = гарах өдөр). Буруу хариу бол null.
function parseIcal_(text) {
  if (!/BEGIN:VCALENDAR/.test(text)) return null;
  const lines = text.replace(/\r\n[ \t]/g, '').split(/\r?\n/);
  const out = []; let ev = null;
  lines.forEach(l => {
    if (l === 'BEGIN:VEVENT') ev = {};
    else if (l === 'END:VEVENT') { if (ev && ev.from && ev.to && ev.from < ev.to) out.push(ev); ev = null; }
    else if (ev) {
      const m = l.match(/^([A-Z-]+)(;[^:]*)?:(.*)$/);
      if (!m) return;
      if (m[1] === 'UID') ev.uid = m[3];
      else if (m[1] === 'SUMMARY') ev.summary = m[3].replace(/\\,/g, ',').slice(0, 120);
      else if (m[1] === 'DTSTART' || m[1] === 'DTEND') {
        const dm = m[3].match(/^(\d{4})(\d{2})(\d{2})/);
        if (dm) ev[m[1] === 'DTSTART' ? 'from' : 'to'] = dm[1] + '-' + dm[2] + '-' + dm[3];
      }
    }
  });
  out.forEach(e => { if (!e.uid) e.uid = e.from + '_' + e.to; });
  return out;
}

// Манай сайтын (OTA-аас ирээгүй) захиалгууд → Booking.com-д импорт хийх iCal
function icalExport_() {
  const sh = getSheet_(), n = sh.getLastRow() - 1;
  const stamp = Utilities.formatDate(new Date(), 'UTC', "yyyyMMdd'T'HHmmss'Z'");
  const ev = [];
  if (n > 0) {
    sh.getRange(2, 1, n, HEADERS.length).getValues().forEach((r, i) => {
      if (r[C.status] === ST.CANCEL) return;
      if (ICAL_FEEDS.some(f => f.name === r[C.source])) return;
      const a = ymd_(r[C.in]), b = ymd_(r[C.out]);
      if (!a || !b || a >= b) return;
      ev.push('BEGIN:VEVENT', 'UID:senden-' + (r[C.code] || 'row' + (i + 2)) + '@sendenguesthouse.com',
        'DTSTAMP:' + stamp, 'DTSTART;VALUE=DATE:' + a.replace(/-/g, ''), 'DTEND;VALUE=DATE:' + b.replace(/-/g, ''),
        'SUMMARY:Senden - Not available', 'END:VEVENT');
    });
  }
  return ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Senden Guest House//Booking//MN', 'CALSCALE:GREGORIAN']
    .concat(ev, ['END:VCALENDAR']).join('\r\n');
}

/* ---------------- туслах ---------------- */

// Цуцлагдаагүй, дуусаагүй бүх захиалгын өдрүүд: [{from, to}] (to = гарах өдөр)
function bookedRanges_() {
  const sh = getSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const today = ymd_(new Date());
  return sh.getRange(2, 1, n, HEADERS.length).getValues()
    .filter(r => String(r[C.status]).trim() !== ST.CANCEL)
    .map(r => ({ from: ymd_(r[C.in]), to: ymd_(r[C.out]) }))
    .filter(r => r.from && r.to && r.from < r.to && r.to > today)
    .sort((a, b) => a.from < b.from ? -1 : 1);
}

function findByCode_(sh, code) {
  const n = sh.getLastRow() - 1;
  if (n < 1 || !code) return null;
  const rows = sh.getRange(2, 1, n, HEADERS.length).getValues();
  const i = rows.findIndex(r => String(r[C.code]).toUpperCase() === code);
  return i < 0 ? null : { index: i + 2, row: rows[i] };
}

function newCode_(sh) {
  const n = sh.getLastRow() - 1;
  const used = n > 0 ? sh.getRange(2, C.code + 1, n, 1).getValues().map(r => String(r[0])) : [];
  let code;
  do { code = 'SG' + (1000 + Math.floor(Math.random() * 9000)); } while (used.indexOf(code) >= 0);
  return code;
}

function isYmd_(s) { return /^\d{4}-\d{2}-\d{2}$/.test(String(s)); }
function digits_(s) { return String(s || '').replace(/\D/g, ''); }

function ymd_(v) {
  const tz = Session.getScriptTimeZone();
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  const m = String(v).trim().match(/^(\d{4})[-.\/](\d{1,2})[-.\/](\d{1,2})/);
  return m ? m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2) : '';
}

function parseDate_(s) {
  const p = String(s).split('-').map(Number);
  return new Date(p[0], p[1] - 1, p[2]);
}

// Хуудсыг үүсгэх / шинэ баганыг нэмэх (хуучин мөрүүд хэвээр үлдэнэ)
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  const head = sh.getLastRow() ? sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0] : [];
  if (head.length < HEADERS.length || head[HEADERS.length - 1] !== HEADERS[HEADERS.length - 1]) {
    if (sh.getMaxColumns() < HEADERS.length) sh.insertColumnsAfter(sh.getMaxColumns(), HEADERS.length - sh.getMaxColumns());
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
      .setFontWeight('bold').setBackground('#1C3D6E').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList([ST.NEW, ST.CHECK, ST.OK, ST.CANCEL], true).build();
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
