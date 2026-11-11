// Mari Takayama — lesson requests, availability and statistics in one Google Sheet.
//
// How to use (once):
//   1. Make a new Google Sheet. Menu: Extensions > Apps Script. Delete what is there and paste this whole file.
//   2. Choose "setup" at the top and press Run. Allow access when Google asks.
//      This makes three tabs: リクエスト (requests), 空き状況 (availability), 統計 (statistics).
//   3. Deploy > New deployment > type "Web app". Execute as: Me. Who has access: Anyone. Deploy.
//   4. Copy the Web app URL (https://script.google.com/macros/s/.../exec) and put it in lessons.html (SHEET_APP).
//
// Every new request is added to リクエスト and sent to your Gmail.
// In 空き状況, type ◯ / △ / × for each date and time. Blank = the usual week shown on the site.

// The sheet's id (from its address). Leave empty when the script is opened from the sheet itself (Extensions > Apps Script).
const SHEET_ID = '';
const book_ = () => SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActive();
const TZ = 'America/New_York';
// Busy times in Mari's Google Calendar (every calendar she owns) become × on the site.
// Only free/busy goes to the website — never the titles or details of her plans.
const USE_GOOGLE_CALENDAR = true;
// Or another calendar as an iCal (.ics) link (e.g. Yahoo Calendar > Share > ICS link).
const CAL_ICS = '';
// the three lesson times on the site (24-hour clock, New York time)
const SLOTS = [[9, 0, 12, 0], [12, 0, 17, 0], [19, 0, 21, 30]];
const REQ = 'リクエスト', AV = '空き状況', ST = '統計';
const HEAD = ['受付日時', '初めて／2回目以降', '名前', '連絡方法', 'メール', '場所', 'エリア', '対象', 'バレエ経験',
  '第1希望', '第2希望', '第3希望', '第1希望の曜日', '第1希望の時間帯', '第2希望の曜日', '第2希望の時間帯', '第3希望の曜日', '第3希望の時間帯',
  'メッセージ', '言語', '状況', 'その他の連絡先'];
const JA = {
  visit: { First: '初めて', Returning: '2回目以降' },
  place: { FuncPhysio: 'FuncPhysio', Studio: 'ダンススタジオ', Travel: '出張', Any: 'どこでも' },
  age: { Child: '子ども', Teen: '10代', Adult: '大人' },
  level: { 'New to ballet': '初めて', 'Some classes': '少し経験あり', 'Trains regularly': '定期的に練習' },
  day: { Sun: '日', Mon: '月', Tue: '火', Wed: '水', Thu: '木', Fri: '金', Sat: '土' },
  time: { Morning: '午前', Afternoon: '午後', Evening: '夜' }
};

function setup() {
  const ss = book_();
  ss.setSpreadsheetTimeZone(TZ);
  // requests
  const req = ss.getSheetByName(REQ) || ss.insertSheet(REQ, 0);
  req.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold').setBackground('#ecdfc6');
  req.setFrozenRows(1);
  req.getRange('A2:A').setNumberFormat('yyyy/mm/dd hh:mm');
  req.getRange('U2:U').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['新規', '返信済み', '予約確定', 'レッスン済み', 'お断り'], true).build());
  req.setColumnWidths(1, HEAD.length, 120);
  // availability: the next 12 weeks, one row per day
  const av = ss.getSheetByName(AV) || ss.insertSheet(AV, 1);
  if (av.getLastRow() < 2) {
    av.getRange(1, 1, 1, 5).setValues([['日付', '曜日', '午前 9〜12時', '午後 12〜17時', '夜 19時〜']]).setFontWeight('bold').setBackground('#ecdfc6');
    av.setFrozenRows(1);
    const rows = []; const d = new Date(); d.setHours(0, 0, 0, 0);
    for (let i = 0; i < 84; i++) { const x = new Date(d); x.setDate(d.getDate() + i); rows.push([x, '=TEXT(A' + (i + 2) + ',"ddd")', '', '', '']); }
    av.getRange(2, 1, rows.length, 5).setValues(rows);
    av.getRange('A2:A').setNumberFormat('yyyy/mm/dd');
  }
  av.getRange('C2:E').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['◯', '△', '×'], true).setAllowInvalid(true).build()).setHorizontalAlignment('center');
  av.getRange('G1').setValue('◯ 空いていそう　△ 要相談　× 不可　（空白はいつもの週：平日は夜だけ◯、日曜は◯）');
  // statistics
  const st = ss.getSheetByName(ST) || ss.insertSheet(ST, 2);
  st.clear();
  const R = "'" + REQ + "'!";
  const block = (row, col, title, formula) => { st.getRange(row, col).setValue(title).setFontWeight('bold'); st.getRange(row + 1, col).setFormula(formula); };
  st.getRange('A1').setValue('レッスンのリクエスト　統計').setFontSize(14).setFontWeight('bold');
  st.getRange('A3:B6').setValues([['リクエストの合計', ''], ['今月', ''], ['初めて', ''], ['2回目以降', '']]);
  st.getRange('B3').setFormula('=COUNTA(' + R + 'C2:C)');
  st.getRange('B4').setFormula('=COUNTIFS(' + R + 'A2:A,">="&(EOMONTH(TODAY(),-1)+1))');
  st.getRange('B5').setFormula('=COUNTIF(' + R + 'B2:B,"初めて")');
  st.getRange('B6').setFormula('=COUNTIF(' + R + 'B2:B,"2回目以降")');
  block(8, 1, '月ごと', '=IFERROR(QUERY({ARRAYFORMULA(IF(' + R + 'A2:A="",,TEXT(' + R + 'A2:A,"yyyy-mm")))},"select Col1, count(Col1) where Col1 is not null group by Col1 order by Col1 desc label Col1 \'月\', count(Col1) \'件数\'",0),"まだありません")');
  block(8, 4, '場所', '=IFERROR(QUERY(' + R + 'F2:F,"select F, count(F) where F <> \'\' group by F order by count(F) desc label F \'場所\', count(F) \'件数\'",0),"まだありません")');
  block(8, 7, '連絡方法', '=IFERROR(QUERY(' + R + 'D2:D,"select D, count(D) where D <> \'\' group by D order by count(D) desc label D \'連絡方法\', count(D) \'件数\'",0),"まだありません")');
  block(22, 1, '人気の曜日（第1〜第3希望の合計）', '=IFERROR(QUERY({' + R + 'M2:M;' + R + 'O2:O;' + R + 'Q2:Q},"select Col1, count(Col1) where Col1 is not null and Col1 <> \'\' group by Col1 order by count(Col1) desc label Col1 \'曜日\', count(Col1) \'回数\'",0),"まだありません")');
  block(22, 4, '人気の時間帯', '=IFERROR(QUERY({' + R + 'N2:N;' + R + 'P2:P;' + R + 'R2:R},"select Col1, count(Col1) where Col1 is not null and Col1 <> \'\' group by Col1 order by count(Col1) desc label Col1 \'時間帯\', count(Col1) \'回数\'",0),"まだありません")');
  block(22, 7, '状況', '=IFERROR(QUERY(' + R + 'U2:U,"select U, count(U) where U <> \'\' group by U label U \'状況\', count(U) \'件数\'",0),"まだありません")');
  block(36, 1, '対象', '=IFERROR(QUERY(' + R + 'H2:H,"select H, count(H) where H <> \'\' group by H label H \'対象\', count(H) \'件数\'",0),"まだありません")');
  block(36, 4, 'バレエ経験', '=IFERROR(QUERY(' + R + 'I2:I,"select I, count(I) where I <> \'\' group by I label I \'経験\', count(I) \'件数\'",0),"まだありません")');
  st.setColumnWidths(1, 8, 130);
}

// a request from the website
function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const c = (d.choices || []).slice(0, 3);
    const choiceText = i => c[i] ? Utilities.formatDate(new Date(c[i].date + 'T12:00:00'), TZ, 'M/d') + '（' + (JA.day[c[i].weekday] || c[i].weekday) + '）' + (JA.time[c[i].time] || c[i].time) : '';
    const row = [new Date(), JA.visit[d.visit] || d.visit || '', clip(d.name), d.via || '', clip(d.email), JA.place[d.place] || d.place || '', clip(d.area),
      JA.age[d.age] || d.age || '', JA.level[d.level] || d.level || '', choiceText(0), choiceText(1), choiceText(2)];
    for (let i = 0; i < 3; i++) row.push(c[i] ? (JA.day[c[i].weekday] || c[i].weekday) : '', c[i] ? (JA.time[c[i].time] || c[i].time) : '');
    row.push(clip(d.note, 700), d.language || '', '新規', clip(d.contact));
    book_().getSheetByName(REQ).appendRow(row);
    const me = Session.getEffectiveUser().getEmail();
    if (me) MailApp.sendEmail(me, 'レッスンのリクエスト：' + String(d.name || '').slice(0, 100) + '（' + (JA.visit[d.visit] || '') + '）',
      String(d.message || '').slice(0, 3000) + '\n\n――\nメール：' + String(d.email || '').slice(0, 200) + (d.contact ? '\n連絡方法：' + (d.via || '') + '　' + String(d.contact).slice(0, 200) : '') + '\nシート：' + book_().getUrl());
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } finally { lock.releaseLock(); }
}

// the website asks for availability: ?availability=1
// the sheet wins when a cell has ◯ / △ / ×; otherwise a busy calendar time becomes ×
function doGet(e) {
  const days = {};
  if (e && e.parameter && e.parameter.availability) {
    busyDays_().forEach((v, k) => { days[k] = v; });
    const av = book_().getSheetByName(AV);
    const today = new Date(); today.setHours(0, 0, 0, 0);
    if (av && av.getLastRow() > 1) av.getRange(2, 1, av.getLastRow() - 1, 5).getValues().forEach(r => {
      if (!(r[0] instanceof Date) || r[0] < today) return;
      if (!r[2] && !r[3] && !r[4]) return;
      const k = Utilities.formatDate(r[0], TZ, 'yyyy-MM-dd'); const cal = days[k] || ['', '', ''];
      days[k] = [0, 1, 2].map(i => String(r[2 + i] || cal[i] || ''));
    });
  }
  return ContentService.createTextOutput(JSON.stringify({ days })).setMimeType(ContentService.MimeType.JSON);
}

// ---------- calendar (free / busy only) ----------
function busyDays_() {
  const out = new Map();
  if (!USE_GOOGLE_CALENDAR && !CAL_ICS) return out;
  const cache = CacheService.getScriptCache();
  const hit = cache.get('busy');
  if (hit) { JSON.parse(hit).forEach(([k, v]) => out.set(k, v)); return out; }
  const start = new Date(); start.setHours(0, 0, 0, 0);
  const end = new Date(start); end.setDate(start.getDate() + 36);
  let events = [];
  if (USE_GOOGLE_CALENDAR) {
    try {
      CalendarApp.getAllOwnedCalendars().forEach(cal => cal.getEvents(start, end).forEach(ev => {
        const mine = ev.getMyStatus && ev.getMyStatus();
        if (mine === CalendarApp.GuestStatus.NO) return;
        events.push({ start: ev.getStartTime(), end: ev.getEndTime(), allDay: ev.isAllDayEvent() });
      }));
    } catch (err) { console.log(err); }
  }
  if (CAL_ICS) {
    try { events = events.concat(parseIcs_(UrlFetchApp.fetch(CAL_ICS, { muteHttpExceptions: true, followRedirects: true }).getContentText())); } catch (err) { console.log(err); }
  }
  for (let i = 0; i < 35; i++) {
    const day = new Date(start); day.setDate(start.getDate() + i);
    const k = Utilities.formatDate(day, TZ, 'yyyy-MM-dd');
    const marks = SLOTS.map(([h1, m1, h2, m2]) => {
      const a = atNY_(k, h1, m1), b = atNY_(k, h2, m2);
      return events.some(ev => ev.allDay ? false : ev.start < b && ev.end > a) ? '×' : '';
    });
    if (marks.some(Boolean)) out.set(k, marks);
  }
  cache.put('busy', JSON.stringify([...out]), 600); // keep for 10 minutes
  return out;
}
function atNY_(ymd, h, m) { return Utilities.parseDate(ymd + ' ' + h + ':' + m, TZ, 'yyyy-MM-dd H:m'); }
function icsDate_(v, params) {
  if (/VALUE=DATE(?!-)/.test(params) || /^\d{8}$/.test(v)) return { date: Utilities.parseDate(v.slice(0, 8), TZ, 'yyyyMMdd'), allDay: true };
  const tz = (params.match(/TZID=([^;:]+)/) || [])[1];
  if (/Z$/.test(v)) return { date: Utilities.parseDate(v.slice(0, 15), 'UTC', "yyyyMMdd'T'HHmmss"), allDay: false };
  let zone = TZ; if (tz) { try { zone = tz.replace(/"/g, ''); Utilities.formatDate(new Date(), zone, 'H'); } catch (err) { zone = TZ; } }
  return { date: Utilities.parseDate(v.slice(0, 15), zone, "yyyyMMdd'T'HHmmss"), allDay: false };
}
function parseIcs_(text) {
  const lines = text.replace(/\r?\n[ \t]/g, '').split(/\r?\n/);
  const events = []; let ev = null;
  const horizon = new Date(); horizon.setDate(horizon.getDate() + 40);
  const from = new Date(); from.setDate(from.getDate() - 1);
  lines.forEach(line => {
    if (line === 'BEGIN:VEVENT') { ev = { ex: [] }; return; }
    if (line === 'END:VEVENT') { if (ev && ev.start && !/CANCELLED/i.test(ev.status || '') && !/TRANSPARENT/i.test(ev.transp || '')) expand_(ev, from, horizon).forEach(x => events.push(x)); ev = null; return; }
    if (!ev) return;
    const i = line.indexOf(':'); if (i < 0) return;
    const head = line.slice(0, i), val = line.slice(i + 1); const name = head.split(';')[0].toUpperCase(); const params = head.slice(name.length);
    if (name === 'DTSTART') { const d = icsDate_(val, params); ev.start = d.date; ev.allDay = d.allDay; }
    else if (name === 'DTEND') ev.end = icsDate_(val, params).date;
    else if (name === 'DURATION') ev.dur = val;
    else if (name === 'RRULE') ev.rrule = val;
    else if (name === 'EXDATE') val.split(',').forEach(v => ev.ex.push(icsDate_(v, params).date.getTime()));
    else if (name === 'STATUS') ev.status = val;
    else if (name === 'TRANSP') ev.transp = val;
  });
  return events;
}
function expand_(ev, from, horizon) {
  let len = ev.end ? ev.end - ev.start : 0;
  if (!ev.end && ev.dur) { const m = ev.dur.match(/P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?)?/); if (m) len = (((+m[1] || 0) * 24 + (+m[2] || 0)) * 60 + (+m[3] || 0)) * 60000; }
  if (!len) len = ev.allDay ? 86400000 : 3600000;
  const one = s => ({ start: s, end: new Date(s.getTime() + len), allDay: ev.allDay });
  if (!ev.rrule) return ev.start < horizon && new Date(ev.start.getTime() + len) > from ? [one(ev.start)] : [];
  const r = {}; ev.rrule.split(';').forEach(p => { const [a, b] = p.split('='); r[a] = b; });
  const until = r.UNTIL ? icsDate_(r.UNTIL, '').date : null; const count = r.COUNT ? +r.COUNT : Infinity; const step = +(r.INTERVAL || 1);
  const wd = ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA']; const byday = r.BYDAY ? r.BYDAY.split(',').map(x => wd.indexOf(x.slice(-2))) : null;
  const out = []; let n = 0;
  for (let d = new Date(ev.start), guard = 0; d < horizon && guard < 3000; guard++) {
    const hits = [];
    if (r.FREQ === 'WEEKLY' && byday) { const ws = new Date(d); ws.setDate(d.getDate() - d.getDay()); byday.forEach(k => { const x = new Date(ws); x.setDate(ws.getDate() + k); if (x >= ev.start) hits.push(x); }); hits.sort((a, b) => a - b); }
    else hits.push(new Date(d));
    for (const x of hits) {
      if (until && x > until) return out; if (n >= count) return out; n++;
      if (x >= from && x < horizon && ev.ex.indexOf(x.getTime()) < 0) out.push(one(x));
    }
    if (r.FREQ === 'DAILY') d.setDate(d.getDate() + step);
    else if (r.FREQ === 'WEEKLY') d.setDate(d.getDate() - (byday ? d.getDay() : 0) + 7 * step);
    else if (r.FREQ === 'MONTHLY') d.setMonth(d.getMonth() + step);
    else if (r.FREQ === 'YEARLY') d.setFullYear(d.getFullYear() + step);
    else break;
  }
  return out;
}

function clip(v, n) { v = String(v == null ? '' : v); v = v.replace(/^[=+\-@]/, "'$&"); return v.slice(0, n || 200); }
