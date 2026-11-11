// Mari Takayama — lesson booking: requests, one-tap confirm, availability and statistics in one Google Sheet.
//
// How to use (once):
//   1. Make a new Google Sheet. Menu: Extensions > Apps Script. Delete what is there and paste this whole file.
//   2. Choose "setup" at the top and press Run. Allow access when Google asks.
//      This makes three tabs: リクエスト (requests), 空き状況 (availability), 統計 (statistics).
//   3. Deploy > New deployment > type "Web app". Execute as: Me. Who has access: Anyone. Deploy.
//   4. Copy the Web app URL (https://script.google.com/macros/s/.../exec) and put it in lessons.html (SHEET_APP).
//
// How it works:
//   - The website asks for free times (?slots=1): the lesson hours below, minus busy times in Mari's Google Calendar.
//     Only free/busy leaves the sheet — never the titles or details of her plans.
//   - A request is added to リクエスト and emailed to Mari with one button per requested time.
//     One tap confirms: the customer gets a confirmation email and the lesson goes into Mari's Google Calendar.

// The sheet's id (from its address). Leave empty when the script is opened from the sheet itself (Extensions > Apps Script).
const SHEET_ID = '';
const book_ = () => SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActive();
const TZ = 'America/New_York';
const USE_GOOGLE_CALENDAR = true;   // busy times in every calendar Mari owns block lesson times
const CAL_ICS = '';                 // or another calendar as an iCal (.ics) link
const REPLY_TO = 'maritakayama.official@gmail.com';  // customers' replies go here
// Confirmed lessons are also sent as a calendar invitation here (e.g. Mari's Yahoo address), so they appear in that calendar.
const EXTRA_CALENDAR_EMAIL = '';
const SENDER_NAME = 'Mari Takayama';
// Paying ahead: Venmo name without @ (empty = the email says Mari will send the details). No cash.
const VENMO = 'maritakayama_ballet';
// Zelle QR code as base64 PNG. Kept only in Mari's own copy of the script (it carries her phone number), never in the public files.
const ZELLE_QR_B64 = '';
// Mari's Zoom meeting link for online lessons (put it only in her own copy of the script). Empty = the email says she will send it.
const ZOOM_LINK = '';
const RATE_PER_30 = 65;   // dollars per 30 minutes, used for FuncPhysio lessons
const BOOKING_PAGE = 'https://bostonterrier808.github.io/mari-takayama-site/lessons.html';
// Lesson hours by weekday (0 = Sunday), 24-hour clock, New York time. A lesson must fit inside these hours.
const HOURS = { 0: [['09:00', '21:00']], 1: [['19:00', '22:00']], 2: [['19:00', '22:00']], 3: [['19:00', '22:00']], 4: [['19:00', '22:00']], 5: [['19:00', '22:00']], 6: [] };
// The 空き状況 tab can open (◯) or close (×) these parts of a day.
const PARTS = [['09:00', '12:00'], ['12:00', '17:00'], ['19:00', '22:00']];
const MIN_NOTICE_HOURS = 12;        // earliest bookable time = now + this
const DAYS_AHEAD = 28;
const PLACES = {
  FuncPhysio: { ja: 'FuncPhysio Physical Therapy（2 W 45th Street, Suite 1600, New York）', en: 'FuncPhysio Physical Therapy, 2 W 45th Street, Suite 1600, New York' },
  Studio: { ja: 'ダンススタジオ（場所は別途ご連絡します）', en: 'A dance studio (Mari will send the address)' },
  Travel: { ja: 'ご指定の場所（出張）', en: 'Your place (Mari comes to you)' },
  Zoom: { ja: 'オンライン（Zoom）', en: 'Online (Zoom)' },
  Any: { ja: '相談して決めます', en: 'To be decided together' }
};

const REQ = 'リクエスト', AV = '空き状況', ST = '統計';
const HEAD = ['受付日時', '初めて／2回目以降', '名前', '連絡方法', 'メール', '場所', 'エリア', '対象', 'バレエ経験',
  '第1希望', '第2希望', '第3希望', '第1希望の曜日', '第1希望の時間帯', '第2希望の曜日', '第2希望の時間帯', '第3希望の曜日', '第3希望の時間帯',
  'メッセージ', '言語', '状況', 'その他の連絡先', 'レッスンの長さ（分）', '確定した日時', 'id', 'key', 'data'];
const C = Object.fromEntries(HEAD.map((h, i) => [h, i + 1]));
const JA = {
  visit: { First: '初めて', Returning: '2回目以降' },
  place: { FuncPhysio: 'FuncPhysio', Studio: 'ダンススタジオ', Travel: '出張', Zoom: 'Zoom', Any: 'どこでも' },
  age: { Child: '子ども', Teen: '10代', Adult: '大人' },
  level: { 'New to ballet': '初めて', 'Some classes': '少し経験あり', 'Trains regularly': '定期的に練習' },
  day: { Sun: '日', Mon: '月', Tue: '火', Wed: '水', Thu: '木', Fri: '金', Sat: '土' }
};
const WD_EN = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'], WD_JA = ['日', '月', '火', '水', '木', '金', '土'];
// time zones a customer can see times in (the website's list); the last number: 1 = in the U.S. (Venmo / Zelle work)
const ZONES = {
  'America/New_York': ['ニューヨーク時間', 'New York time', 1], 'America/Chicago': ['シカゴ時間（中部）', 'Central time', 1],
  'America/Denver': ['デンバー時間（山岳部）', 'Mountain time', 1], 'America/Phoenix': ['アリゾナ時間', 'Arizona time', 1],
  'America/Los_Angeles': ['ロサンゼルス時間（太平洋）', 'Pacific time', 1], 'America/Anchorage': ['アラスカ時間', 'Alaska time', 1],
  'Pacific/Honolulu': ['ハワイ時間', 'Hawaii time', 1], 'Asia/Tokyo': ['日本時間', 'Japan time', 0], 'Europe/London': ['ロンドン時間', 'London time', 0]
};

function setup() {
  const ss = book_();
  ss.setSpreadsheetTimeZone(TZ);
  const req = ss.getSheetByName(REQ) || ss.insertSheet(REQ, 0);
  req.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold').setBackground('#ecdfc6');
  req.setFrozenRows(1);
  req.getRange('A2:A').setNumberFormat('yyyy/mm/dd hh:mm');
  req.getRange(2, C['状況'], req.getMaxRows() - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['新規', '時間を提案中', '返信済み', '予約確定', 'レッスン済み', 'お断り'], true).build());
  req.setColumnWidths(1, HEAD.length, 120);
  try { req.hideColumns(C['id'], 3); } catch (e) {}
  const av = ss.getSheetByName(AV) || ss.insertSheet(AV, 1);
  if (av.getLastRow() < 2) {
    av.getRange(1, 1, 1, 5).setValues([['日付', '曜日', '午前 9〜12時', '午後 12〜17時', '夜 19〜22時']]).setFontWeight('bold').setBackground('#ecdfc6');
    av.setFrozenRows(1);
    const rows = []; const d = new Date(); d.setHours(0, 0, 0, 0);
    for (let i = 0; i < 84; i++) { const x = new Date(d); x.setDate(d.getDate() + i); rows.push([x, '=TEXT(A' + (i + 2) + ',"ddd")', '', '', '']); }
    av.getRange(2, 1, rows.length, 5).setValues(rows);
    av.getRange('A2:A').setNumberFormat('yyyy/mm/dd');
  }
  av.getRange('C2:E').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['◯', '×'], true).setAllowInvalid(true).build()).setHorizontalAlignment('center');
  av.getRange('G1').setValue('ふだんは使わなくて大丈夫です。◯ = いつもは閉じている時間を開ける　× = その時間を閉じる（Google カレンダーの予定は自動で×になります）');
  const st = ss.getSheetByName(ST) || ss.insertSheet(ST, 2);
  st.clear();
  const R = "'" + REQ + "'!";
  const block = (row, col, title, formula) => { st.getRange(row, col).setValue(title).setFontWeight('bold'); st.getRange(row + 1, col).setFormula(formula); };
  st.getRange('A1').setValue('レッスンのリクエスト　統計').setFontSize(14).setFontWeight('bold');
  st.getRange('A3:B7').setValues([['リクエストの合計', ''], ['今月', ''], ['初めて', ''], ['2回目以降', ''], ['予約確定', '']]);
  st.getRange('B3').setFormula('=COUNTA(' + R + 'C2:C)');
  st.getRange('B4').setFormula('=COUNTIFS(' + R + 'A2:A,">="&(EOMONTH(TODAY(),-1)+1))');
  st.getRange('B5').setFormula('=COUNTIF(' + R + 'B2:B,"初めて")');
  st.getRange('B6').setFormula('=COUNTIF(' + R + 'B2:B,"2回目以降")');
  st.getRange('B7').setFormula('=COUNTIF(' + R + 'U2:U,"予約確定")+COUNTIF(' + R + 'U2:U,"レッスン済み")');
  const q = (rng, label, unit) => '=IFERROR(QUERY(' + rng + ',"select Col1, count(Col1) where Col1 is not null and Col1 <> \'\' group by Col1 order by count(Col1) desc label Col1 \'' + label + '\', count(Col1) \'' + unit + '\'",0),"まだありません")';
  block(9, 1, '月ごと', '=IFERROR(QUERY({ARRAYFORMULA(IF(' + R + 'A2:A="",,TEXT(' + R + 'A2:A,"yyyy-mm")))},"select Col1, count(Col1) where Col1 is not null group by Col1 order by Col1 desc label Col1 \'月\', count(Col1) \'件数\'",0),"まだありません")');
  block(9, 4, '場所', q('{' + R + 'F2:F}', '場所', '件数'));
  block(9, 7, '連絡方法', q('{' + R + 'D2:D}', '連絡方法', '件数'));
  block(23, 1, '人気の曜日（第1〜第3希望の合計）', q('{' + R + 'M2:M;' + R + 'O2:O;' + R + 'Q2:Q}', '曜日', '回数'));
  block(23, 4, '人気の時間帯', q('{' + R + 'N2:N;' + R + 'P2:P;' + R + 'R2:R}', '時間帯', '回数'));
  block(23, 7, '状況', q('{' + R + 'U2:U}', '状況', '件数'));
  block(37, 1, 'レッスンの長さ', q('{' + R + 'W2:W}', '分', '件数'));
  block(37, 4, '対象', q('{' + R + 'H2:H}', '対象', '件数'));
  block(37, 7, 'バレエ経験', q('{' + R + 'I2:I}', '経験', '件数'));
  st.setColumnWidths(1, 8, 130);
}

// ---------- the website sends a request ----------
function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (!looksReal_(d)) return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
    const len = Math.max(30, Math.min(240, +d.length || 60));
    const c = (d.choices || []).slice(0, 3).filter(x => /^\d{4}-\d\d-\d\d$/.test(x.date));
    const id = Utilities.getUuid().slice(0, 8), key = Utilities.getUuid().replace(/-/g, '');
    const tz = zoneOk_(d.tz); d.tz = tz;
    const area = d.place === 'Zoom' ? zoneName_(tz, true) : clip(d.area);
    const row = [new Date(), JA.visit[d.visit] || d.visit || '', clip(d.name), d.via || '', clip(d.email), JA.place[d.place] || d.place || '', area,
      JA.age[d.age] || d.age || '', JA.level[d.level] || d.level || ''];
    for (let i = 0; i < 3; i++) row.push(c[i] ? choiceJa_(c[i], len) : '');
    for (let i = 0; i < 3; i++) row.push(c[i] ? WD_JA[dateOf_(c[i].date).getDay()] : '', c[i] ? partOf_(c[i]) : '');
    row.push(clip(d.note, 700), d.language || '', '新規', clip(d.contact), len, '', id, key, JSON.stringify({ id, choices: c, len, place: d.place || '', area: String(d.area || '').slice(0, 120), name: String(d.name || '').slice(0, 100), email: String(d.email || '').slice(0, 120), language: d.language === 'ja' ? 'ja' : 'en', tz }));
    book_().getSheetByName(REQ).appendRow(row);
    notifyMari_(d, c, len, id, key);
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } finally { lock.releaseLock(); }
}

// quietly drop bot or flood submissions (the visitor still sees "sent", so bots learn nothing)
function looksReal_(d) {
  if (!d || d.hp) return false;                                   // hidden field filled = a bot
  if (typeof d.ms === 'number' && d.ms < 4000) return false;      // filled in under 4 seconds = a bot
  const email = String(d.email || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || email.length > 120) return false;
  if (!String(d.name || '').trim() || String(d.name).length > 100) return false;
  if (!Array.isArray(d.choices) || !d.choices.length || d.choices.length > 3) return false;
  const cache = CacheService.getScriptCache();
  const bump = (k, max, secs) => { const n = +(cache.get(k) || 0) + 1; cache.put(k, String(n), secs); return n <= max; };
  if (!bump('m:' + email, 3, 3600)) return false;                  // same email: at most 3 an hour
  if (!bump('all', 30, 3600)) return false;                        // everyone together: at most 30 an hour
  return true;
}
function notifyMari_(d, c, len, id, key) {
  const me = Session.getEffectiveUser().getEmail(); if (!me) return;
  const url = ScriptApp.getService().getUrl();
  const btn = (href, label, main) => '<a href="' + href + '" style="display:inline-block;margin:4px 0;padding:12px 18px;background:' + (main ? '#ecdfc6' : '#ffffff') + ';border:1px solid #c9b48c;color:#2c2a30;text-decoration:none;font-size:16px">' + label + '</a>';
  const rows = c.map((x, i) => '<p style="margin:10px 0">' + btn(url + '?act=confirm&id=' + id + '&k=' + key + '&c=' + i, '第' + (i + 1) + '希望で確定：' + choiceJa_(x, len), true) + '</p>').join('');
  const info = [
    ['名前', esc_(d.name) + '（' + (JA.visit[d.visit] || '') + '）'], ['メール', esc_(d.email)],
    d.contact ? ['連絡方法', esc_(d.via) + '　' + esc_(d.contact)] : null,
    ['場所', esc_(JA.place[d.place] || d.place || '') + (d.area && d.place !== 'Zoom' ? '（' + esc_(d.area) + '）' : '')], ['長さ', len + '分'],
    d.tz !== TZ ? ['お客さんの時間', esc_(zoneName_(d.tz, true)) + '<br>' + c.map((x, i) => '第' + (i + 1) + '希望：' + whenIn_(x, len, d.tz, true)).join('<br>')] : null,
    d.age ? ['対象', esc_(JA.age[d.age] || d.age)] : null, d.level ? ['バレエ経験', esc_(JA.level[d.level] || d.level)] : null,
    d.note ? ['メッセージ', esc_(d.note)] : null
  ].filter(Boolean).map(([a, b]) => '<tr><td style="padding:3px 12px 3px 0;color:#6a6670">' + a + '</td><td>' + b + '</td></tr>').join('');
  const html = '<div style="font-family:sans-serif;color:#2c2a30;font-size:15px;line-height:1.6">'
    + '<p><b>レッスンのリクエストが届きました。</b>確定する時間のボタンを押すと、お客さんに確定メールが届き、Google カレンダーにも入ります。' + (d.tz !== TZ ? '<br>ボタンの時間はニューヨーク時間です（お客さんの時間は下にあります）。' : '') + '</p>'
    + rows
    + '<p style="margin:16px 0 6px">' + btn(url + '?act=propose&id=' + id + '&k=' + key, '時間を変えて提案する（例：30分ずらす）', false) + '</p>'
    + '<p style="margin:6px 0 16px">' + btn(url + '?act=decline&id=' + id + '&k=' + key, 'どの時間も無理なので、お断りする（ほかの日を選び直してもらう）', false) + '</p>'
    + '<table style="margin-top:14px;font-size:14px">' + info + '</table>'
    + '<p style="color:#6a6670;font-size:13px;margin-top:18px">ボタンを押すと確認の画面が出ます。そこでもう一度押すと送られます。<br>シート：' + book_().getUrl() + '</p></div>';
  MailApp.sendEmail({ to: me, subject: 'レッスンのリクエスト：' + String(d.name || '').slice(0, 60) + '（' + (JA.visit[d.visit] || '') + '）', htmlBody: html, body: String(d.message || '').slice(0, 3000), name: 'Lesson booking' });
}

// ---------- the website asks for free times, or Mari taps a button ----------
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.slots) return json_({ days: freeTimes_() });
  if (p.availability) return json_({ days: {} });
  if (p.act) return actionPage_(p);
  return HtmlService.createHtmlOutput('<p>OK</p>');
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// free time ranges per day, in minutes from midnight (New York): {'2026-11-20': [[1140, 1320]], ...}
function freeTimes_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('free'); if (hit) return JSON.parse(hit);
  const start = new Date(); start.setHours(0, 0, 0, 0);
  const end = new Date(start); end.setDate(start.getDate() + DAYS_AHEAD + 1);
  const busy = busyEvents_(start, end);
  const sheet = sheetMarks_();
  const earliest = new Date(Date.now() + MIN_NOTICE_HOURS * 3600000);
  const out = {};
  for (let i = 0; i <= DAYS_AHEAD; i++) {
    const day = new Date(start); day.setDate(start.getDate() + i);
    const k = Utilities.formatDate(day, TZ, 'yyyy-MM-dd');
    const wd = +Utilities.formatDate(day, TZ, 'u') % 7;
    let ranges = (HOURS[wd] || []).map(([a, b]) => [mins_(a), mins_(b)]);
    const marks = sheet[k];
    if (marks) PARTS.forEach(([a, b], j) => {
      if (marks[j] === 'o') ranges = add_(ranges, [mins_(a), mins_(b)]);
      if (marks[j] === 'x') ranges = cut_(ranges, [mins_(a), mins_(b)]);
    });
    const dayStart = atNY_(k, 0, 0);
    busy.forEach(ev => {
      if (ev.allDay || ev.end <= dayStart || ev.start >= new Date(dayStart.getTime() + 86400000 + 3600000)) return;
      ranges = cut_(ranges, [Math.floor((ev.start - dayStart) / 60000), Math.ceil((ev.end - dayStart) / 60000)]);
    });
    if (earliest > dayStart) ranges = cut_(ranges, [0, Math.ceil((earliest - dayStart) / 60000)]);
    ranges = ranges.filter(([a, b]) => b - a >= 60);
    if (ranges.length) out[k] = ranges;
  }
  cache.put('free', JSON.stringify(out), 300);
  return out;
}
function add_(ranges, [a, b]) {
  const all = ranges.concat([[a, b]]).sort((x, y) => x[0] - y[0]); const out = [];
  all.forEach(r => { const last = out[out.length - 1]; if (last && r[0] <= last[1]) last[1] = Math.max(last[1], r[1]); else out.push(r.slice()); });
  return out;
}
function cut_(ranges, [a, b]) {
  const out = []; ranges.forEach(([x, y]) => { if (b <= x || a >= y) out.push([x, y]); else { if (a > x) out.push([x, a]); if (b < y) out.push([b, y]); } }); return out;
}
function mins_(hhmm) { const [h, m] = hhmm.split(':').map(Number); return h * 60 + m; }
function sheetMarks_() {
  const out = {}; const av = book_().getSheetByName(AV); if (!av || av.getLastRow() < 2) return out;
  av.getRange(2, 1, av.getLastRow() - 1, 5).getValues().forEach(r => {
    if (!(r[0] instanceof Date)) return;
    const m = [r[2], r[3], r[4]].map(v => /×|✕|^x$/i.test(String(v)) ? 'x' : /◯|○|〇|^o$/i.test(String(v)) ? 'o' : '');
    if (m.some(Boolean)) out[Utilities.formatDate(r[0], TZ, 'yyyy-MM-dd')] = m;
  });
  return out;
}
function busyEvents_(start, end) {
  let events = [];
  if (USE_GOOGLE_CALENDAR) {
    try {
      CalendarApp.getAllOwnedCalendars().forEach(cal => cal.getEvents(start, end).forEach(ev => {
        if (ev.getMyStatus && ev.getMyStatus() === CalendarApp.GuestStatus.NO) return;
        events.push({ start: ev.getStartTime(), end: ev.getEndTime(), allDay: ev.isAllDayEvent() });
      }));
    } catch (err) { console.log(err); }
  }
  if (CAL_ICS) { try { events = events.concat(parseIcs_(UrlFetchApp.fetch(CAL_ICS, { muteHttpExceptions: true }).getContentText())); } catch (err) { console.log(err); } }
  return events;
}

// ---------- Mari's buttons: confirm one time, or decline ----------
function actionPage_(p) {
  const sh = book_().getSheetByName(REQ);
  const ids = sh.getLastRow() > 1 ? sh.getRange(2, C['id'], sh.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
  const i = ids.indexOf(String(p.id));
  if (i < 0) return page_('見つかりません', 'このリクエストは見つかりませんでした。');
  const r = i + 2;
  const data = JSON.parse(sh.getRange(r, C['data']).getValue() || '{}');
  const status = String(sh.getRange(r, C['状況']).getValue());
  if (p.act === 'accept') return acceptPage_(p, sh, r, data, status);
  if (String(sh.getRange(r, C['key']).getValue()) !== String(p.k)) return page_('開けません', 'このボタンは使えません。');
  if (status === '予約確定' || status === 'お断り') return page_('対応済みです', 'このリクエストはもう「' + status + '」になっています。' + (status === '予約確定' ? '<br>確定した日時：' + esc_(sh.getRange(r, C['確定した日時']).getDisplayValue()) : ''));
  const self = ScriptApp.getService().getUrl() + '?act=' + encodeURIComponent(p.act) + '&id=' + encodeURIComponent(p.id) + '&k=' + encodeURIComponent(p.k) + (p.c != null ? '&c=' + encodeURIComponent(p.c) : '');
  if (p.act === 'confirm') {
    const ch = (data.choices || [])[+p.c]; if (!ch) return page_('見つかりません', 'その希望日時が見つかりませんでした。');
    const s = atNY_(ch.date, ...ch.start.split(':').map(Number)), en = new Date(s.getTime() + data.len * 60000);
    const clash = busyEvents_(s, en).filter(ev => !ev.allDay && ev.start < en && ev.end > s).length;
    if (!p.go) return page_('この時間で確定しますか？', '<b>' + esc_(data.name) + '</b> さん<br>' + choiceJa_(ch, data.len) + '（ニューヨーク時間）' + (data.tz && data.tz !== TZ ? '<br>お客さんの時間：' + whenIn_(ch, data.len, data.tz, true) + '（' + zoneName_(data.tz, true) + '）' : '') + '<br>' + esc_(PLACES[data.place] ? PLACES[data.place].ja : '') + (clash ? '<p style="color:#8b3a3a">この時間には、Google カレンダーにほかの予定があります。</p>' : '') + '<p>確定すると、お客さんに確定メールが届き、Google カレンダーに入ります。</p>', self + '&go=1', '確定して、お客さんにメールを送る');
    bookLesson_(sh, r, data, ch, data.len, '');
    return page_('確定しました', esc_(data.name) + ' さんと Mari さんに確定メールを送り、Google カレンダーに入れました。<br>' + choiceJa_(ch, data.len));
  }
  if (p.act === 'propose') {
    const ch0 = (data.choices || [])[0] || { date: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), start: '19:00' };
    if (!p.go) {
      const opt = (v, cur, label) => '<option value="' + v + '"' + (String(v) === String(cur) ? ' selected' : '') + '>' + (label || v) + '</option>';
      let times = ''; for (let t = 7 * 60; t <= 22 * 60; t += 15) times += opt(Math.floor(t / 60) + ':' + String(t % 60).padStart(2, '0'), ch0.start.replace(/^0/, ''));
      const lens = [30, 45, 60, 75, 90, 120, 150, 180].map(n => opt(n, data.len, n + '分')).join('');
      const base = ScriptApp.getService().getUrl();
      const asked = (data.choices || []).map((x, j) => '第' + (j + 1) + '希望：' + choiceJa_(x, data.len)).join('<br>');
      const form = '<form method="get" action="' + base + '" target="_top" style="margin-top:14px">'
        + '<input type="hidden" name="act" value="propose"><input type="hidden" name="id" value="' + esc_(p.id) + '"><input type="hidden" name="k" value="' + esc_(p.k) + '"><input type="hidden" name="go" value="1">'
        + '<p>日にち<br><input type="date" name="date" value="' + ch0.date + '" style="font-size:17px;padding:8px;width:100%;box-sizing:border-box"></p>'
        + '<p>始まる時間<br><select name="start" style="font-size:17px;padding:8px;width:100%">' + times + '</select></p>'
        + '<p>長さ<br><select name="len" style="font-size:17px;padding:8px;width:100%">' + lens + '</select></p>'
        + '<p>ひとこと（なくても大丈夫）<br><textarea name="msg" rows="3" style="font-size:16px;padding:8px;width:100%;box-sizing:border-box" placeholder="例：19:30 だと少し間に合わないので、20:00 からでいかがでしょうか。"></textarea></p>'
        + '<p><button type="submit" style="font-size:17px;padding:14px 22px;background:#ecdfc6;border:1px solid #c9b48c;color:#2c2a30">この時間をお客さんに提案する</button></p></form>';
      return page_('時間を変えて提案する', '<b>' + esc_(data.name) + '</b> さんの希望：<br>' + asked + form);
    }
    if (!/^\d{4}-\d\d-\d\d$/.test(p.date || '') || !/^\d{1,2}:\d\d$/.test(p.start || '')) return page_('入力を確かめてください', '日にちと時間を選んでください。');
    const prop = { date: p.date, start: p.start, len: Math.max(15, Math.min(240, +p.len || data.len)) };
    data.id = data.id || String(p.id); data.proposal = prop; data.ckey = data.ckey || Utilities.getUuid().replace(/-/g, '');
    sh.getRange(r, C['data']).setValue(JSON.stringify(data));
    sh.getRange(r, C['状況']).setValue('時間を提案中');
    sendCustomer_(data, 'propose', prop, String(p.msg || '').slice(0, 500));
    return page_('提案を送りました', esc_(data.name) + ' さんに、' + choiceJa_(prop, prop.len) + '（' + prop.len + '分）を提案しました。<br>お客さんが「この時間で予約する」を押すと、自動で確定して、Google カレンダーに入り、Mari さんにもメールが届きます。');
  }
  if (p.act === 'decline') {
    if (!p.go) return page_('お断りしますか？', '<b>' + esc_(data.name) + '</b> さんに、「希望の時間はうまってしまいました。ほかの日時を選び直してください」というメールを送ります。', self + '&go=1', 'お断りのメールを送る');
    sendCustomer_(data, 'decline');
    sh.getRange(r, C['状況']).setValue('お断り');
    return page_('送りました', esc_(data.name) + ' さんに、ほかの日時を選び直してもらうメールを送りました。');
  }
  return page_('?', '');
}
// one confirmed lesson: Google Calendar (+ invitation to the extra calendar), emails to the customer and to Mari, the sheet
function bookLesson_(sh, r, data, ch, len, how) {
  const s = atNY_(ch.date, ...ch.start.split(':').map(Number)), en = new Date(s.getTime() + len * 60000);
  const place = placeText_(data, false);
  const theirs = data.tz && data.tz !== TZ ? '\nお客さんの時間：' + whenIn_(ch, len, data.tz, true) + '（' + zoneName_(data.tz, true) + '）' : '';
  const opts = { location: place, description: 'メール：' + data.email + (data.area ? '\nエリア：' + data.area : '') + theirs + '\n（予約サイトから' + (how ? '・' + how : '') + '）' };
  if (EXTRA_CALENDAR_EMAIL) { opts.guests = EXTRA_CALENDAR_EMAIL; opts.sendInvites = true; }
  CalendarApp.getDefaultCalendar().createEvent('バレエレッスン：' + data.name, s, en, opts);
  data.len = len;
  sendCustomer_(data, 'confirm', ch);
  sh.getRange(r, C['状況']).setValue('予約確定'); sh.getRange(r, C['確定した日時']).setValue(choiceJa_(ch, len));
  CacheService.getScriptCache().remove('free');
  const me = Session.getEffectiveUser().getEmail();
  if (me) MailApp.sendEmail({ to: me, subject: '予約が確定しました：' + data.name + '（' + choiceJa_(ch, len) + '）',
    body: data.name + ' さんのレッスンが確定しました' + (how ? '（' + how + '）' : '') + '。\n\n日時：' + choiceJa_(ch, len) + '（' + len + '分・ニューヨーク時間）' + theirs + '\n場所：' + placeText_(data, true) + (data.area && data.place !== 'Zoom' ? '（' + data.area + '）' : '') + '\nメール：' + data.email + '\n\nお客さんにも確定メールを送りました。Google カレンダーに入れました' + (EXTRA_CALENDAR_EMAIL ? '（Yahoo カレンダーにも招待として送りました）' : '') + '。\n添付のファイルを開くと、ほかのカレンダーにも追加できます。',
    attachments: [ics_(data, ch, len)], name: 'Lesson booking' });
}
function ics_(data, ch, len) {
  const s = atNY_(ch.date, ...ch.start.split(':').map(Number)), en = new Date(s.getTime() + len * 60000);
  const z = d => Utilities.formatDate(d, 'UTC', "yyyyMMdd'T'HHmmss'Z'");
  const txt = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Mari Takayama//Lesson//EN', 'METHOD:PUBLISH', 'BEGIN:VEVENT',
    'UID:' + (data.id || Utilities.getUuid()) + '@maritakayama', 'DTSTAMP:' + z(new Date()), 'DTSTART:' + z(s), 'DTEND:' + z(en),
    'SUMMARY:Ballet lesson — Mari Takayama / ' + String(data.name || '').replace(/[,;\n]/g, ' '),
    'LOCATION:' + placeText_(data, false).replace(/,/g, '\\,'), 'END:VEVENT', 'END:VCALENDAR'].join('\r\n');
  return Utilities.newBlob(txt, 'text/calendar', 'lesson.ics');
}
function acceptPage_(p, sh, r, data, status) {
  const ja = data.language === 'ja';
  if (!data.ckey || String(p.k) !== data.ckey || !data.proposal) return page_(ja ? '開けません' : 'Link not valid', ja ? 'このリンクは使えません。' : 'This link can no longer be used.');
  const prop = data.proposal; const when = whenIn_(prop, prop.len, data.tz || TZ, ja);
  if (status === '予約確定') return page_(ja ? '予約は確定しています' : 'Your lesson is booked', (ja ? '日時：' : 'When: ') + esc_(sh.getRange(r, C['確定した日時']).getDisplayValue() || when));
  if (status === 'お断り') return page_(ja ? 'このご予約は終了しています' : 'This request is closed', ja ? 'お手数ですが、予約ページから選び直してください。<br>' + BOOKING_PAGE : 'Please choose a new time here:<br>' + BOOKING_PAGE);
  if (!p.go) {
    const self = ScriptApp.getService().getUrl() + '?act=accept&id=' + encodeURIComponent(p.id) + '&k=' + encodeURIComponent(p.k) + '&go=1';
    return page_(ja ? 'この時間で予約しますか？' : 'Book this time?', '<b>' + custWhen_(data, prop, prop.len, ja).replace(/\n/g, '<br>') + '</b><br>' + esc_(PLACES[data.place] ? PLACES[data.place][ja ? 'ja' : 'en'] : ''), self, ja ? 'この時間で予約する' : 'Book this time');
  }
  bookLesson_(sh, r, data, prop, prop.len, '提案した時間で、お客さんが予約');
  return page_(ja ? '予約が確定しました' : 'Your lesson is booked', (ja ? '確定のメールをお送りしました。<br>日時：' : 'A confirmation email is on its way.<br>When: ') + when);
}
function page_(title, body, goHref, goLabel) {
  const html = '<!doctype html><meta name="viewport" content="width=device-width,initial-scale=1"><div style="font-family:Georgia,serif;color:#2c2a30;max-width:520px;margin:40px auto;padding:0 18px;line-height:1.7;font-size:17px">'
    + '<h1 style="font-weight:400;font-size:28px;margin:0 0 14px">' + title + '</h1><div>' + body + '</div>'
    + (goHref ? '<p style="margin-top:22px"><a href="' + goHref + '" target="_top" style="display:inline-block;padding:14px 22px;background:#ecdfc6;border:1px solid #c9b48c;color:#2c2a30;text-decoration:none">' + goLabel + '</a></p>' : '') + '</div>';
  return HtmlService.createHtmlOutput(html).setTitle(title).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
// the friendly "please pay ahead" part of the confirmation email
function payNote_(data, ch, ja) {
  const z = ZONES[data.tz || TZ];
  if (!z || !z[2]) return (ja ? ['レッスン代は前もってお願いしています。お支払いの方法は、別のメールでお知らせします。', '', 'ご予定が変わったときは：', '・2日前まで：全額お返しします', '・前日：半分お返しします', '・当日：お返しできません']
    : ['Lessons are paid ahead of time. Mari will email you how to pay from where you are.', '', 'If your plans change:', '• Up to 2 days before: fully refunded', '• The day before: half refunded', '• On the day: no refund']).join('\n');
  const amount = data.place === 'FuncPhysio' ? Math.round(RATE_PER_30 * data.len / 30) : 0;
  const d = dateOf_(ch.date); const memo = 'Ballet lesson ' + (d.getMonth() + 1) + '/' + d.getDate();
  const link = VENMO ? 'https://venmo.com/' + VENMO + '?txn=pay' + (amount ? '&amount=' + amount : '') + '&note=' + encodeURIComponent(memo) : '';
  const lines = ja ? [
    'レッスン代は、前もって Venmo' + (ZELLE_QR_B64 ? ' か Zelle' : '') + ' でお願いします（現金は受け取っていません）。',
    amount ? '金額：$' + amount : '金額は、別のメールでお知らせします。',
    link ? 'Venmo：こちらを開くと、金額が入った状態で開きます\n' + link : (VENMO ? '' : '送り先は、別のメールでお知らせします。'),
    ZELLE_QR_B64 ? 'Zelle：添付の QR コードを、銀行のアプリで読み取ってください' : '',
    '', 'ご予定が変わったときは：',
    '・2日前まで：全額お返しします', '・前日：半分お返しします', '・当日：お返しできません'
  ] : [
    'Please pay for the lesson ahead of time by Venmo' + (ZELLE_QR_B64 ? ' or Zelle' : '') + ' (sorry, no cash).',
    amount ? 'Amount: $' + amount : 'Mari will send you the amount in a separate email.',
    link ? 'Venmo — this link opens with everything filled in:\n' + link : (VENMO ? '' : 'Mari will send you where to send it.'),
    ZELLE_QR_B64 ? 'Zelle — scan the attached QR code in your banking app' : '',
    '', 'If your plans change:',
    '• Up to 2 days before: fully refunded', '• The day before: half refunded', '• On the day: no refund'
  ];
  return lines.filter((x, i) => x !== '' || lines[i - 1] !== '').join('\n');
}
function sendCustomer_(data, kind, ch, note) {
  if (!data.email) return;
  const ja = data.language === 'ja';
  const place = placeText_(data, ja);
  const zoom = data.place === 'Zoom';
  let subject, body;
  if (kind === 'confirm') {
    const when = whenIn_(ch, data.len, data.tz || TZ, ja), full = custWhen_(data, ch, data.len, ja);
    const prep = zoom
      ? (ja ? (ZOOM_LINK ? '' : 'Zoom のリンクは、レッスンの前にお送りします。\n') + '動きやすい服装で、少し動ける広さの場所をご用意ください。全身が映るようにカメラを置き、椅子の背やテーブルなど、つかまれるものがあるとバーの代わりになります。'
          : (ZOOM_LINK ? '' : 'Mari will send you the Zoom link before the lesson.\n') + 'Please wear comfortable clothes and clear a little space to move. Set the camera so your whole body is in view; a chair back or a counter works as a barre.')
      : (ja ? '動きやすい服装でお越しください。バレエシューズがあればお持ちください（靴下でも大丈夫です）。' : 'Please wear comfortable clothes you can move in, and bring ballet slippers if you have them (socks are fine too).');
    subject = ja ? 'レッスンのご予約が確定しました（' + when + '）' : 'Your ballet lesson is confirmed — ' + when;
    body = ja
      ? data.name + ' 様\n\nご予約ありがとうございます。レッスンの日時が確定しました。\n\n日時：' + full + '\n場所：' + place + (data.area && !zoom ? '（' + data.area + '）' : '') + '\nレッスン：' + data.len + '分\n\n' + payNote_(data, ch, true) + '\n\n' + prep + '\n変更やキャンセルは、このメールにご返信ください。\n\nお会いできるのを楽しみにしています。\n髙山まり'
      : 'Hi ' + data.name + ',\n\nThank you for booking — your lesson is confirmed.\n\nWhen: ' + full + '\nWhere: ' + place + (data.area && !zoom ? ' (' + data.area + ')' : '') + '\nLesson: ' + data.len + ' minutes\n\n' + payNote_(data, ch, false) + '\n\n' + prep + '\nTo change or cancel, just reply to this email.\n\nSee you soon!\nMari Takayama';
  } else if (kind === 'propose') {
    const when = whenIn_(ch, ch.len, data.tz || TZ, ja), full = custWhen_(data, ch, ch.len, ja);
    const link = ScriptApp.getService().getUrl() + '?act=accept&id=' + encodeURIComponent(data.id || '') + '&k=' + data.ckey;
    subject = ja ? 'レッスンの時間のご提案（' + when + '）' : 'A time for your lesson — ' + when;
    body = ja
      ? data.name + ' 様\n\nお問い合わせありがとうございます。ご希望の時間に近い、こちらの時間はいかがでしょうか。\n\n日時：' + full + '\n場所：' + place + '\nレッスン：' + ch.len + '分\n' + (note ? '\n' + note + '\n' : '') + '\nこの時間でよろしければ、下のリンクを開いて「この時間で予約する」を押してください。予約が確定します。\n' + link + '\n\nほかの日時がよろしければ、こちらから選び直していただけます。\n' + BOOKING_PAGE + '\n\n髙山まり'
      : 'Hi ' + data.name + ',\n\nThank you for your request. Would this time work for you?\n\nWhen: ' + full + '\nWhere: ' + place + '\nLesson: ' + ch.len + ' minutes\n' + (note ? '\n' + note + '\n' : '') + '\nIf it does, open this link and tap “Book this time” — your lesson will be confirmed:\n' + link + '\n\nIf another time is better, you can choose here:\n' + BOOKING_PAGE + '\n\nMari Takayama';
  } else {
    subject = ja ? 'レッスンのご希望日時について' : 'About your lesson request';
    body = ja
      ? data.name + ' 様\n\nお問い合わせありがとうございます。申し訳ありません、ご希望の日時はほかの予定が入ってしまいました。\n\nお手数ですが、こちらからほかの日時を選び直していただけますか。\n' + BOOKING_PAGE + '\n\nよろしくお願いいたします。\n髙山まり'
      : 'Hi ' + data.name + ',\n\nThank you for your request. I’m sorry — the times you chose are no longer available.\n\nCould you pick another time here?\n' + BOOKING_PAGE + '\n\nThank you,\nMari Takayama';
  }
  const mail = { to: data.email, subject: subject, body: body, name: SENDER_NAME, replyTo: REPLY_TO };
  if (kind === 'confirm') { mail.attachments = [ics_(data, ch, data.len)].concat(ZELLE_QR_B64 ? [Utilities.newBlob(Utilities.base64Decode(ZELLE_QR_B64), 'image/png', 'zelle-qr.png')] : []); mail.body += ja ? '\n\n添付のファイルを開くと、お使いのカレンダーに追加できます。' : '\n\nOpen the attached file to add the lesson to your calendar.'; }
  MailApp.sendEmail(mail);
}

// ---------- helpers ----------
function dateOf_(ymd) { return Utilities.parseDate(ymd + ' 12:00', TZ, 'yyyy-MM-dd HH:mm'); }
function atNY_(ymd, h, m) { return Utilities.parseDate(ymd + ' ' + h + ':' + m, TZ, 'yyyy-MM-dd H:m'); }
function addMin_(hhmm, n) { const t = mins_(hhmm) + n; return Math.floor(t / 60) + ':' + String(t % 60).padStart(2, '0'); }
function partOf_(c) { const t = mins_(c.start || '19:00'); return t < 720 ? '午前' : t < 1020 ? '午後' : '夜'; }
function choiceJa_(c, len) { const d = dateOf_(c.date); return (d.getMonth() + 1) + '/' + d.getDate() + '（' + WD_JA[d.getDay()] + '）' + c.start + '〜' + addMin_(c.start, len); }
function choiceEn_(c, len) {
  const d = dateOf_(c.date); const f = hhmm => { let [h, m] = hhmm.split(':').map(Number); const ap = h >= 12 ? 'pm' : 'am'; h = h % 12 || 12; return h + ':' + String(m).padStart(2, '0') + ' ' + ap; };
  return WD_EN[d.getDay()] + ', ' + (d.getMonth() + 1) + '/' + d.getDate() + ', ' + f(c.start) + '–' + f(addMin_(c.start, len));
}
function zoneOk_(z) { z = String(z || ''); return /^[A-Za-z]+(\/[A-Za-z0-9_+\-]+){1,2}$/.test(z) ? z : TZ; }
function zoneName_(z, ja) { const n = ZONES[z]; return n ? n[ja ? 0 : 1] : z.split('/').pop().replace(/_/g, ' ') + (ja ? ' の時間' : ' time'); }
// a lesson time as read in time zone z
function whenIn_(c, len, z, ja) {
  const s = atNY_(c.date, ...c.start.split(':').map(Number)), e = new Date(s.getTime() + len * 60000);
  const wd = +Utilities.formatDate(s, z, 'u') % 7;
  if (ja) return Utilities.formatDate(s, z, 'M/d') + '（' + WD_JA[wd] + '）' + Utilities.formatDate(s, z, 'H:mm') + '〜' + Utilities.formatDate(e, z, 'H:mm');
  const f = d => Utilities.formatDate(d, z, 'h:mm a').toLowerCase();
  return WD_EN[wd] + ', ' + Utilities.formatDate(s, z, 'M/d') + ', ' + f(s) + '–' + f(e);
}
// the time a customer reads: in their own time zone, plus New York time when that is different
function custWhen_(data, c, len, ja) {
  const z = data.tz || TZ; const own = whenIn_(c, len, z, ja) + (ja ? '（' + zoneName_(z, true) + '）' : ' (' + zoneName_(z, false) + ')');
  return z === TZ ? own : own + (ja ? '\n　　ニューヨーク時間では ' : '\n      In New York: ') + whenIn_(c, len, TZ, ja);
}
function placeText_(data, ja) {
  if (data.place === 'Zoom') return (ja ? 'オンライン（Zoom）' : 'Online (Zoom)') + (ZOOM_LINK ? ' ' + ZOOM_LINK : '');
  return PLACES[data.place] ? PLACES[data.place][ja ? 'ja' : 'en'] : '';
}
function esc_(v) { return String(v == null ? '' : v).replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch])); }
function clip(v, n) { v = String(v == null ? '' : v); v = v.replace(/^[=+\-@]/, "'$&"); return v.slice(0, n || 200); }

// ---------- iCal reading (only used when CAL_ICS is set) ----------
function icsDate_(v, params) {
  if (/VALUE=DATE(?!-)/.test(params) || /^\d{8}$/.test(v)) return { date: Utilities.parseDate(v.slice(0, 8), TZ, 'yyyyMMdd'), allDay: true };
  if (/Z$/.test(v)) return { date: Utilities.parseDate(v.slice(0, 15), 'UTC', "yyyyMMdd'T'HHmmss"), allDay: false };
  const tz = (params.match(/TZID=([^;:]+)/) || [])[1];
  return { date: Utilities.parseDate(v.slice(0, 15), tz ? tz.replace(/"/g, '') : TZ, "yyyyMMdd'T'HHmmss"), allDay: false };
}
function parseIcs_(text) {
  const events = []; let ev = null;
  text.replace(/\r?\n[ \t]/g, '').split(/\r?\n/).forEach(line => {
    if (line === 'BEGIN:VEVENT') { ev = {}; return; }
    if (line === 'END:VEVENT') { if (ev && ev.start && !/CANCELLED/i.test(ev.status || '')) events.push({ start: ev.start, end: ev.end || new Date(ev.start.getTime() + 3600000), allDay: ev.allDay }); ev = null; return; }
    if (!ev) return; const i = line.indexOf(':'); if (i < 0) return;
    const head = line.slice(0, i), val = line.slice(i + 1), name = head.split(';')[0].toUpperCase(), params = head.slice(name.length);
    if (name === 'DTSTART') { const d = icsDate_(val, params); ev.start = d.date; ev.allDay = d.allDay; }
    else if (name === 'DTEND') ev.end = icsDate_(val, params).date;
    else if (name === 'STATUS') ev.status = val;
  });
  return events;
}
