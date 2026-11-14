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
const RATE_PER_30 = 65;
// when Mari books another studio for the customer: studio cost + this handling fee (only shown in emails)
const STUDIO_FEE = 10;   // dollars per 30 minutes, used for FuncPhysio lessons
const SITE = 'https://bostonterrier808.github.io/mari-takayama-site/';
const BOOKING_PAGE = 'https://bostonterrier808.github.io/mari-takayama-site/lessons.html';
// Lesson hours by weekday (0 = Sunday), 24-hour clock, New York time. A lesson must fit inside these hours.
const HOURS = { 0: [['10:30', '21:00']], 1: [['19:00', '22:00']], 2: [['19:00', '22:00']], 3: [['19:00', '22:00']], 4: [['19:00', '22:00']], 5: [['19:00', '22:00']], 6: [] };
// The 空き状況 tab can open (◯) or close (×) these parts of a day.
const PARTS = [['09:00', '12:00'], ['12:00', '17:00'], ['19:00', '22:00']];
const MIN_NOTICE_HOURS = 12;        // earliest bookable time = now + this
const DAYS_AHEAD = 92;
const PLACES = {
  FuncPhysio: { ja: 'FuncPhysio Physical Therapy（2 W 45th Street, Suite 1600, New York）', en: 'FuncPhysio Physical Therapy, 2 W 45th Street, Suite 1600, New York' },
  Studio: { ja: 'ダンススタジオ（場所は別途ご連絡します）', en: 'A dance studio (Mari will send the address)' },
  Travel: { ja: 'ご指定の場所（出張）', en: 'Your place (Mari comes to you)' },
  Zoom: { ja: 'オンライン（Zoom）', en: 'Online (Zoom)' },
  Any: { ja: '相談して決めます', en: 'To be decided together' }
};

const REQ = 'リクエスト', AV = '空き状況', ST = '統計', PK = '回数券';
// 5-lesson pack: paid ahead, half a lesson off, used within 3 months (for one lesson length)
const PACK_SIZE = 5, PACK_MONTHS = 3;
const PACK_HEAD = ['購入日', '名前', 'メール', '長さ（分）', '回数', '使った回数', '期限', 'メモ'];
// Whole days with no lessons at all (tours, holidays): [first day, last day], New York dates. Keep in step with AWAY in lessons.html.
const AWAY = [['2026-11-22', '2027-01-09']];
const isAway_ = k => AWAY.some(([a, b]) => k >= a && k <= b);
const HEAD = ['受付日時', '初めて／2回目以降', '名前', '連絡方法', 'メール', '場所', 'エリア', '対象', 'バレエ経験',
  '第1希望', '第2希望', '第3希望', '第1希望の曜日', '第1希望の時間帯', '第2希望の曜日', '第2希望の時間帯', '第3希望の曜日', '第3希望の時間帯',
  'メッセージ', '言語', '状況', 'その他の連絡先', 'レッスンの長さ（分）', '確定した日時', 'id', 'key', 'data', 'きっかけ', '紹介者', 'ニックネーム', '性別', 'SNS', 'レッスン日', '支払い', '催促', 'レッスンの言語', '受けたい理由'];
const C = Object.fromEntries(HEAD.map((h, i) => [h, i + 1]));
const JA = {
  visit: { First: '初めて', Returning: '2回目以降' },
  place: { FuncPhysio: 'FuncPhysio', Studio: 'ダンススタジオ', Travel: '出張', Zoom: 'Zoom', Any: 'どこでも' },
  age: { Child: '子ども', Teen: '10代', Adult: '大人' },
  goals: { 'Start ballet': 'バレエを始めたい', 'Strong basics': '基礎', 'Better technique': 'テクニックを上げたい', 'Expressiveness': '表現力を上げたい', 'Pointe work': 'ポアント', 'Learn a variation': 'バリエーション', 'Competition or audition': 'コンクール・オーディション', 'Fitness and posture': '体づくり・姿勢', 'Just for fun': '楽しみとして', 'Something else': 'その他' },
  tongue: { English: '英語', Japanese: '日本語', Either: 'どちらでも' },
  gender: { Female: '女性', Male: '男性', Nonbinary: 'ノンバイナリー', Other: 'その他', None: '答えない' },
  heard: { Instagram: 'Instagram', Website: 'ウェブサイト・検索', Mari: 'Mari本人から', Friend: '友達の紹介', Studio: 'スタジオ・学校', Performance: '公演を見て', Other: 'その他' },
  level: { 'New to ballet': '初めて', 'Some classes': '少し経験あり', 'Trains regularly': '定期的に練習', 'Professional dancer': 'プロのダンサー', 'Ballet student': 'バレエ学生', 'Dance student': 'ダンス学生', 'Musical theater (singer first)': 'ミュージカル（歌がメイン）', 'Musical theater (dancer first)': 'ミュージカル（ダンスがメイン）' },
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
  packSheet_();
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
    const c = (d.choices || []).slice(0, 5).filter(x => /^\d{4}-\d\d-\d\d$/.test(x.date));
    const id = 'r' + Utilities.getUuid().slice(0, 7), key = Utilities.getUuid().replace(/-/g, '');
    const tz = zoneOk_(d.tz); d.tz = tz;
    const area = d.place === 'Zoom' ? zoneName_(tz, true) : clip(d.area);
    const row = [new Date(), JA.visit[d.visit] || d.visit || '', clip(d.name), d.via || '', clip(d.email), JA.place[d.place] || d.place || '', area,
      JA.age[d.age] || d.age || '', JA.level[d.level] || d.level || ''];
    for (let i = 0; i < 3; i++) row.push(c[i] ? choiceJa_(c[i], len) : '');
    for (let i = 0; i < 3; i++) row.push(c[i] ? WD_JA[dateOf_(c[i].date).getDay()] : '', c[i] ? partOf_(c[i]) : '');
    const more = c.slice(3).map((x, i) => '第' + (i + 4) + '希望：' + choiceJa_(x, len)).join('\n');
    row.push(clip((more ? more + (d.note ? '\n' : '') : '') + (d.note || ''), 900), d.language || '', '新規', clip(d.contact), len, '', id, key, JSON.stringify({ id, choices: c, len, place: d.place || '', area: String(d.area || '').slice(0, 120), name: String(d.name || '').slice(0, 100), email: String(d.email || '').slice(0, 120), language: d.language === 'ja' ? 'ja' : 'en', tz, pack: d.pack === '5' ? '5' : '', first: d.visit === 'First', free: d.free === true, studio: d.place === 'Studio' ? (d.studio === 'mari' ? 'mari' : 'self') : '', friend: d.visit === 'First' ? String(d.friend || '').slice(0, 120) : '', multi: d.multi === true && c.length > 1 }), d.visit === 'First' ? (JA.heard[d.heard] || '') + (d.heard === 'Other' && d.heardNote ? '（' + clip(d.heardNote, 200) + '）' : '') : '', d.visit === 'First' && d.heard === 'Friend' ? clip(d.friend, 120) : '', clip(d.nickname, 60), d.visit === 'First' ? (JA.gender[d.gender] || '') : '', clip(d.social, 120), '', '', '', d.visit === 'First' ? (JA.tongue[d.tongue] || '') : '', goalsJa_(d));
    heardStats_();
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
  if (!Array.isArray(d.choices) || !d.choices.length || d.choices.length > 5) return false;
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
  const multi = d.multi === true && c.length > 1;
  const rows = multi
    ? '<p style="margin:10px 0"><b>まとめて予約（' + c.length + '回）：</b><br>' + c.map((x, i) => (i + 1) + '回目：' + choiceJa_(x, len)).join('<br>') + '</p><p style="margin:10px 0">' + btn(url + '?act=confirmall&id=' + id + '&k=' + key, 'この' + c.length + '回をまとめて確定', true) + '</p>'
    : c.map((x, i) => '<p style="margin:10px 0">' + btn(url + '?act=confirm&id=' + id + '&k=' + key + '&c=' + i, '第' + (i + 1) + '希望で確定：' + choiceJa_(x, len), true) + '</p>').join('');
  const info = [
    ['名前', esc_(d.name) + '（' + (JA.visit[d.visit] || '') + '）'], d.nickname ? ['ニックネーム', esc_(String(d.nickname).slice(0, 60))] : null,
    d.visit === 'First' && JA.gender[d.gender] ? ['性別', JA.gender[d.gender]] : null, ['メール', esc_(d.email)],
    d.social ? ['SNS', esc_(String(d.social).slice(0, 120))] : null,
    d.visit === 'First' && JA.tongue[d.tongue] ? ['レッスンの言語', JA.tongue[d.tongue]] : null,
    goalsJa_(d) ? ['受けたい理由', esc_(goalsJa_(d))] : null,
    d.contact ? ['連絡方法', esc_(d.via) + '　' + esc_(d.contact)] : null,
    ['場所', esc_(JA.place[d.place] || d.place || '') + (d.area && d.place !== 'Zoom' ? '（' + esc_(d.area) + '）' : '')], ['長さ', len + '分' + (d.pack === '5' ? '（5回券を希望）' : '')],
    d.place === 'Studio' ? ['スタジオ', d.studio === 'mari' ? '<b>Mari が代わりに予約</b>（スタジオ代＋手数料 $' + STUDIO_FEE + '）' : 'お客さんが自分で予約'] : null,
    d.free === true ? ['種類', '<b>無料の Zoom 相談（60分）</b>'] : null,
    d.visit === 'First' && d.heard ? ['知ったきっかけ', esc_(JA.heard[d.heard] || d.heard) + (d.heard === 'Friend' && d.friend ? '（' + esc_(clip(d.friend)) + ' さん）' : '') + (d.heard === 'Other' && d.heardNote ? '（' + esc_(clip(d.heardNote, 200)) + '）' : '')] : null,
    d.tz !== TZ ? ['お客さんの時間', esc_(zoneName_(d.tz, true)) + '<br>' + c.map((x, i) => '第' + (i + 1) + '希望：' + whenIn_(x, len, d.tz, true)).join('<br>')] : null,
    d.age ? ['対象', esc_(JA.age[d.age] || d.age)] : null, d.level ? ['バレエ経験', esc_(JA.level[d.level] || d.level)] : null,
    d.note ? ['メッセージ', esc_(d.note)] : null
  ].filter(Boolean).map(([a, b]) => '<tr><td style="padding:3px 12px 3px 0;color:#6a6670">' + a + '</td><td>' + b + '</td></tr>').join('');
  const html = '<div style="color:#2c2a30;font-size:15px;line-height:1.6">'
    + '<p><b>レッスンのリクエストが届きました。</b>確定する時間のボタンを押すと、お客さんに確定メールが届き、Google カレンダーにも入ります。' + (d.tz !== TZ ? '<br>ボタンの時間はニューヨーク時間です（お客さんの時間は下にあります）。' : '') + '</p>'
    + rows
    + '<p style="margin:16px 0 6px">' + btn(url + '?act=propose&id=' + id + '&k=' + key, '時間を変えて提案する（例：30分ずらす）', false) + '</p>'
    + '<p style="margin:6px 0 16px">' + btn(url + '?act=decline&id=' + id + '&k=' + key, 'どの時間も無理なので、お断りする（ほかの日を選び直してもらう）', false) + '</p>'
    + '<table style="margin-top:14px;font-size:14px">' + info + '</table>'
    + '<p style="color:#6a6670;font-size:13px;margin-top:18px">ボタンを押すと確認の画面が出ます。そこでもう一度押すと送られます。<br>シート：' + book_().getUrl() + '</p></div>';
  MailApp.sendEmail({ to: me, subject: 'レッスンのリクエスト：' + String(d.name || '').slice(0, 60) + '（' + (JA.visit[d.visit] || '') + '）', htmlBody: pretty_(html, true), body: String(d.message || '').slice(0, 3000), name: 'Lesson booking' });
}

// ---------- the website asks for free times, or Mari taps a button ----------
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.slots) return json_({ days: freeTimes_() });
  if (p.availability) return json_({ days: {} });
  if (p.pack) return json_({ packs: packsFor_(p.pack) });
  if (p.cleanup) { cleanupTests_(); return json_({ ok: true }); }
  if (p.reminders) { const c = CacheService.getScriptCache(); if (!c.get('remind')) { c.put('remind', '1', 6 * 3600); paymentReminders(); } return json_({ ok: true }); }
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
    if (isAway_(k)) continue;
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
  if (p.act === 'mine' || p.act === 'cancel' || p.act === 'change') return customerPage_(p, sh, r, data);
  if (p.act === 'take' || p.act === 'decl') return offerAnswer_(p, sh, r, data);
  if (String(sh.getRange(r, C['key']).getValue()) !== String(p.k)) return page_('開けません', 'このボタンは使えません。');
  if (p.act === 'move' || p.act === 'keep') return moveAnswer_(p, sh, r, data);
  if (p.act === 'mmove' || p.act === 'mcancel') return mariChange_(p, sh, r, data);
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
  if (p.act === 'confirmall') {
    const list = (data.choices || []);
    if (!list.length) return page_('見つかりません', '希望日時が見つかりませんでした。');
    if (!p.go) return page_('まとめて確定しますか？', '<b>' + esc_(data.name) + '</b> さん（' + list.length + '回）<br>' + list.map((x, i) => (i + 1) + '回目：' + choiceJa_(x, data.len)).join('<br>') + '<br>（ニューヨーク時間）', self + '&go=1', 'この' + list.length + '回をまとめて確定する');
    list.forEach(x => bookLesson_(sh, r, data, x, data.len, 'まとめて予約'));
    sh.getRange(r, C['確定した日時']).setValue(list.map(x => choiceJa_(x, data.len)).join(' ／ '));
    return page_('確定しました', esc_(data.name) + ' さんの ' + list.length + '回分を確定して、Google カレンダーに入れました。<br>' + list.map(x => choiceJa_(x, data.len)).join('<br>'));
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
  const ev = CalendarApp.getDefaultCalendar().createEvent((data.free ? '無料 Zoom 相談：' : 'バレエレッスン：') + data.name, s, en, opts);
  data.len = len;
  data.packInfo = data.free ? null : usePack_(data, ch, len);
  data.id = data.id || String(sh.getRange(r, C['id']).getValue());
  data.ckey = data.ckey || Utilities.getUuid().replace(/-/g, '');
  data.booked = data.booked || [];
  data.booked.push({ date: ch.date, start: ch.start, len: len, ev: ev.getId(), pack: data.packInfo ? data.packInfo.row : 0 });
  data._b = data.booked.length - 1;
  sh.getRange(r, C['data']).setValue(JSON.stringify(data));
  const friendLine = data.first && data.friend ? '\n紹介：' + data.friend + ' さんの紹介です。' : '';  // record only, no automatic credit
  sendCustomer_(data, 'confirm', ch);
  sh.getRange(r, C['状況']).setValue('予約確定'); sh.getRange(r, C['確定した日時']).setValue(choiceJa_(ch, len));
  heardStats_();
  const firstDay = String(sh.getRange(r, C['レッスン日']).getValue() || '');
  if (!firstDay || firstDay > ch.date + ' ' + ch.start) sh.getRange(r, C['レッスン日']).setNumberFormat('@').setValue(ch.date + ' ' + ch.start);
  const noPay = data.free || (data.packInfo && !data.packInfo.bought);
  const pc = sh.getRange(r, C['支払い']);
  if (noPay) { pc.clearDataValidations(); pc.setValue('不要'); } else { pc.insertCheckboxes(); pc.setValue(false); }
  CacheService.getScriptCache().remove('free');
  const me = Session.getEffectiveUser().getEmail();
  if (me) sendPretty_({ to: me, subject: '予約が確定しました：' + data.name + '（' + choiceJa_(ch, len) + '）',
    body: data.name + ' さんのレッスンが確定しました' + (how ? '（' + how + '）' : '') + '。\n\n日時：' + choiceJa_(ch, len) + '（' + len + '分・ニューヨーク時間）' + theirs + '\n場所：' + placeText_(data, true) + (data.area && data.place !== 'Zoom' ? '（' + data.area + '）' : '') + '\nメール：' + data.email + packLine_(data.packInfo) + friendLine + '\n\nお客さんにも確定メールを送りました。Google カレンダーに入れました' + (EXTRA_CALENDAR_EMAIL ? '（Yahoo カレンダーにも招待として送りました）' : '') + '。\n添付のファイルを開くと、ほかのカレンダーにも追加できます。'
      + '\n\nMari さんの都合で変えるときは：\n' + mariUrl_(sh, r, data, 'mmove') + '\n' + mariUrl_(sh, r, data, 'mcancel'),
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
  const html = '<!doctype html><meta name="viewport" content="width=device-width,initial-scale=1"><div style="font-family:-apple-system,BlinkMacSystemFont,\'Hiragino Sans\',\'Hiragino Kaku Gothic ProN\',\'Yu Gothic\',Meiryo,\'Helvetica Neue\',Arial,sans-serif;color:#222222;max-width:520px;margin:40px auto;padding:0 18px;line-height:1.8;font-size:17px">'
    + '<div style="font-family:Georgia,serif;font-size:15px;letter-spacing:.12em;color:#8a7140;margin-bottom:18px">Mari Takayama</div>'
    + '<h1 style="font-family:Georgia,serif;font-weight:400;font-size:26px;margin:0 0 14px">' + title + '</h1><div>' + body + '</div>'
    + (goHref ? '<p style="margin-top:22px"><a href="' + goHref + '" target="_top" style="display:inline-block;padding:14px 22px;background:#ecdfc6;border:1px solid #c9b48c;color:#2c2a30;text-decoration:none">' + goLabel + '</a></p>' : '') + '</div>';
  return HtmlService.createHtmlOutput(html).setTitle(title).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
// the friendly "please pay ahead" part of the confirmation email
function payNote_(data, ch, ja) {
  if (data.free) return ja ? 'この相談は無料です（お支払いはいりません）。\nご都合が変わったときは、このメールに返信してお知らせください。' : 'This call is free, so there is nothing to pay.\nIf your plans change, just reply to this email.';
  const pk = data.packInfo;
  const rules = pk
    ? (ja ? ['', 'ご予定が変わったときは：', '・2日前まで：回数券に1回分お戻しします', '・前日・当日：1回使ったことになります', '・Mari の急な予定変更でキャンセルになった場合：回数券に1回分お戻しします']
          : ['', 'If your plans change:', '• Up to 2 days before: the lesson goes back on your pack', '• The day before or on the day: it counts as used', '• If Mari has to cancel because of a sudden schedule change: the lesson goes back on your pack'])
    : (ja ? ['', 'ご予定が変わったときは：', '・2日前まで：全額お返しします', '・前日：半分お返しします', '・当日：お返しできません', '・Mari の急な予定変更でキャンセルになった場合：全額お返しします']
          : ['', 'If your plans change:', '• Up to 2 days before: fully refunded', '• The day before: half refunded', '• On the day: no refund', '• If Mari has to cancel because of a sudden schedule change: fully refunded']);
  const studio = data.place !== 'Studio' ? [] : data.studio === 'mari'
    ? [ja ? 'スタジオは Mari が代わりに予約します。スタジオ代と手数料 $' + STUDIO_FEE + ' を、レッスン代と一緒にお願いします（スタジオ代は Mari からお知らせします）。' : 'Mari will book the studio for you. Please add the studio cost plus a $' + STUDIO_FEE + ' booking fee to the lesson fee (Mari will tell you the studio cost).']
    : [ja ? 'スタジオの予約が取れたら、スタジオの名前・住所・部屋を、このメールに返信してお知らせください。' : 'Once you have booked the studio, please reply to this email with its name, address and room.'];
  const travel = (data.place === 'Travel' ? [ja ? 'マンハッタンの外への出張は、交通費を別にお願いしています（金額は Mari からお知らせします）。' : 'For lessons outside Manhattan, travel costs are added; Mari will let you know the amount.'] : []).concat(studio);
  // a lesson from a pack already paid for
  if (pk && !pk.bought) return [ja ? '今回のレッスンは回数券を使います（お支払いはいりません）。' : 'This lesson uses your lesson pack, so there is nothing to pay.',
    ja ? '残り：' + pk.left + '回（' + pk.until + 'まで）' : 'Left on your pack: ' + pk.left + ' (use by ' + pk.until + ')'].concat(travel, rules).join('\n');
  const per = data.place === 'FuncPhysio' ? Math.round(RATE_PER_30 * data.len / 30) : 0;
  const amount = pk ? (per ? Math.round(per * PACK_SIZE - per / 2) : 0) : per;
  const what = pk ? (ja ? ['5回券（今回が1回目、残り' + pk.left + '回は' + pk.until + 'まで使えます）'] : ['5-lesson pack (this is lesson 1; the other ' + pk.left + ' can be used until ' + pk.until + ')']) : [];
  const z = ZONES[data.tz || TZ];
  if (!z || !z[2]) return what.concat([ja ? 'レッスン代は前もって（遅くとも ' + payBy_(ch).ja + 'まで）お願いしています。お支払いの方法は、別のメールでお知らせします。' : 'Lessons are paid ahead of time (at the latest by ' + payBy_(ch).en + '). Mari will email you how to pay from where you are.'], travel, rules).join('\n');
  const d = dateOf_(ch.date); const memo = (pk ? 'Ballet lessons (5-pack) from ' : 'Ballet lesson ') + (d.getMonth() + 1) + '/' + d.getDate();
  const link = VENMO ? 'https://venmo.com/' + VENMO + '?txn=pay' + (amount ? '&amount=' + amount : '') + '&note=' + encodeURIComponent(memo) : '';
  const dl = payBy_(ch);
  const lines = what.concat(ja ? [
    'お支払いについて：',
    'ご予約を確実にするため、今すぐお支払いをお願いします。',
    '',
    '金額：' + (amount ? '$' + amount + (pk ? '（' + PACK_SIZE + '回分から1回分の半額をお引きしています）' : '') : '別のメールでお知らせします'),
    '期限：' + dl.ja + 'まで',
    '方法：Venmo' + (ZELLE_QR_B64 ? ' か Zelle' : '') + '（現金不可）',
    '',
    link ? link : (VENMO ? '' : '送り先は、別のメールでお知らせします。'),
    ZELLE_QR_B64 ? 'Zelle の場合は、添付の QR コードを銀行のアプリで読み取ってください。' : ''
  ] : [
    'Payment:',
    'To secure your booking, please pay now.',
    '',
    'Amount: ' + (amount ? '$' + amount + (pk ? ' (' + PACK_SIZE + ' lessons, less half a lesson)' : '') : 'Mari will send it in a separate email'),
    'Pay by: ' + dl.en,
    'How: Venmo' + (ZELLE_QR_B64 ? ' or Zelle' : '') + ' (sorry, no cash)',
    '',
    link ? link : (VENMO ? '' : 'Mari will send you where to send it.'),
    ZELLE_QR_B64 ? 'For Zelle, scan the attached QR code in your banking app.' : ''
  ], travel, rules);
  return lines.filter((x, i) => x !== '' || lines[i - 1] !== '').join('\n');
}
function packSheet_() {
  const ss = book_(); let sh = ss.getSheetByName(PK);
  if (!sh) {
    sh = ss.insertSheet(PK);
    sh.getRange(1, 1, 1, PACK_HEAD.length).setValues([PACK_HEAD]).setFontWeight('bold').setBackground('#ecdfc6');
    sh.setFrozenRows(1); sh.getRange('A2:A').setNumberFormat('yyyy/mm/dd'); sh.getRange('G2:G').setNumberFormat('yyyy/mm/dd');
    sh.getRange('J1').setValue('予約が確定すると自動で記録されます。キャンセルで1回分戻すときは「使った回数」を1つ減らしてください。');
  }
  return sh;
}
// on a confirmed lesson: use one from this person's pack, or start a new pack if they asked for one
function usePack_(data, ch, len) {
  const sh = packSheet_(), v = sh.getDataRange().getValues(), day = dateOf_(ch.date), email = String(data.email || '').trim().toLowerCase();
  const md = x => Utilities.formatDate(new Date(x), TZ, 'M/d');
  for (let i = 1; i < v.length; i++) {
    const em = String(v[i][2]).trim().toLowerCase(), l = +v[i][3], n = +v[i][4], used = +v[i][5], until = v[i][6];
    if (em && em === email && l === +len && used < n && (!until || new Date(until).getTime() + 86400000 > day.getTime())) {
      sh.getRange(i + 1, 6).setValue(used + 1);
      return { left: n - used - 1, until: until ? md(until) : '', row: i + 1 };
    }
  }
  if (data.pack !== '5') return null;
  const until = new Date(day); until.setMonth(until.getMonth() + PACK_MONTHS);
  sh.appendRow([new Date(), data.name, data.email, len, PACK_SIZE, 1, until, '']);
  return { bought: true, left: PACK_SIZE - 1, until: md(until), row: sh.getLastRow() };
}
// "Bring a friend": the first lesson of someone a friend sent adds one lesson to that friend's pack
function creditFriend_(data) {
  const who = String(data.friend).trim().toLowerCase(), me = String(data.email || '').trim().toLowerCase();
  if (!who || who === me) return '';
  const sh = packSheet_(), v = sh.getDataRange().getValues(), now = Date.now();
  for (let i = v.length - 1; i >= 1; i--) {
    const em = String(v[i][2]).trim().toLowerCase(), nm = String(v[i][1]).trim().toLowerCase(), until = v[i][6];
    if ((em && em === who) || (nm && nm === who)) {
      if (until && new Date(until).getTime() + 86400000 < now) continue;
      sh.getRange(i + 1, 5).setValue(+v[i][4] + 1);
      sh.getRange(i + 1, 8).setValue(String(v[i][7] || '') + (v[i][7] ? '／' : '') + '紹介：' + data.name + '（+1回）');
      return '\n紹介：' + v[i][1] + ' さんの紹介です。' + v[i][1] + ' さんの5回券に1回分を追加しました。';
    }
  }
  return '\n紹介：' + data.friend + ' さんの紹介です（5回券が見つからなかったので、まだ何も追加していません。お礼をするなら「回数券」タブで足してください）。';
}
// remaining lessons on someone's packs (only numbers and dates go back to the page)
function packsFor_(email) {
  email = String(email || '').trim().toLowerCase(); if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return [];
  const cache = CacheService.getScriptCache(), k = 'pk:' + email, n = +(cache.get(k) || 0) + 1; cache.put(k, String(n), 600); if (n > 10) return [];
  const sh = book_().getSheetByName(PK); if (!sh || sh.getLastRow() < 2) return [];
  const now = Date.now();
  return sh.getDataRange().getValues().slice(1).filter(v => String(v[2]).trim().toLowerCase() === email && +v[5] < +v[4] && (!v[6] || new Date(v[6]).getTime() + 86400000 > now))
    .map(v => ({ left: +v[4] - +v[5], len: +v[3], until: v[6] ? Utilities.formatDate(new Date(v[6]), TZ, 'yyyy-MM-dd') : '' }));
}
// make sure the request sheet has the two new columns and the stats tab counts them
function heardStats_() {
  const ss = book_(), req = ss.getSheetByName(REQ);
  if (req && req.getRange(1, C['きっかけ']).getValue() !== 'きっかけ') req.getRange(1, C['きっかけ'], 1, 2).setValues([['きっかけ', '紹介者']]).setFontWeight('bold').setBackground('#ecdfc6');
  if (req && req.getRange(1, C['ニックネーム']).getValue() !== 'ニックネーム') req.getRange(1, C['ニックネーム'], 1, 3).setValues([['ニックネーム', '性別', 'SNS']]).setFontWeight('bold').setBackground('#ecdfc6');
  if (req && req.getRange(1, C['受けたい理由']).getValue() !== '受けたい理由') req.getRange(1, C['レッスンの言語'], 1, 2).setValues([['レッスンの言語', '受けたい理由']]).setFontWeight('bold').setBackground('#ecdfc6');
  if (req && req.getRange(1, C['支払い']).getValue() !== '支払い') { req.getRange(1, C['レッスン日'], 1, 3).setValues([['レッスン日', '支払い', '催促']]).setFontWeight('bold').setBackground('#ecdfc6'); try { req.hideColumns(C['レッスン日']); } catch (e) {} }
  const st = ss.getSheetByName(ST); if (!st || st.getRange('A51').getValue()) return;
  const R = "'" + REQ + "'!", q = (rng, label) => '=IFERROR(QUERY(' + rng + ',"select Col1, count(Col1) where Col1 is not null and Col1 <> \'\' group by Col1 order by count(Col1) desc label Col1 \'' + label + '\', count(Col1) \'人数\'",0),"まだありません")';
  st.getRange('A51').setValue('どこで知ったか').setFontWeight('bold'); st.getRange('A52').setFormula(q('{' + R + 'AB2:AB}', 'きっかけ'));
  st.getRange('D51').setValue('紹介してくれた人').setFontWeight('bold'); st.getRange('D52').setFormula(q('{' + R + 'AC2:AC}', '名前'));
}
function packLine_(pk) {
  if (!pk) return '';
  return pk.bought ? '\n回数券：新しく5回券を購入（今回が1回目・残り' + pk.left + '回・' + pk.until + 'まで）。「回数券」タブに記録しました。'
    : '\n回数券：1回使いました（残り' + pk.left + '回・' + pk.until + 'まで）';
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
    subject = data.free ? (ja ? '無料の Zoom 相談が確定しました（' + when + '）' : 'Your free Zoom call is confirmed — ' + when) : ja ? 'レッスンのご予約が確定しました（' + when + '）' : 'Your ballet lesson is confirmed — ' + when;
    body = ja
      ? '# ' + (data.free ? '無料相談' : 'ご予約') + '確定のお知らせ\n\n' + data.name + ' 様\n\nご予約ありがとうございます。' + (data.free ? '無料相談' : 'レッスン') + 'の日時が確定しました。\n\n日時：' + full + '\n場所：' + place + (data.area && !zoom ? '（' + data.area + '）' : '') + '\n' + (data.free ? '相談' : 'レッスン') + '：' + data.len + '分\n\n' + payNote_(data, ch, true) + '\n\n' + (data.free ? (ZOOM_LINK ? '' : 'Zoom のリンクは、当日までにお送りします。\n') + '今どこで踊っているか、目指していることなどを、気軽にお話しください。' : prep) + manageLine_(data, true) + '\n\nお会いできるのを楽しみにしています。\n髙山まり'
      : '# ' + (data.free ? 'Your free call is confirmed' : 'Your lesson is confirmed') + '\n\nHi ' + data.name + ',\n\nThank you for booking — your ' + (data.free ? 'free call' : 'lesson') + ' is confirmed.\n\nWhen: ' + full + '\nWhere: ' + place + (data.area && !zoom ? ' (' + data.area + ')' : '') + '\n' + (data.free ? 'Call' : 'Lesson') + ': ' + data.len + ' minutes\n\n' + payNote_(data, ch, false) + '\n\n' + (data.free ? (ZOOM_LINK ? '' : 'Mari will send you the Zoom link before the call.\n') + 'Come ready to talk about where you dance now and what you are working toward.' : prep) + manageLine_(data, false) + '\n\nSee you soon!\nMari Takayama';
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
  mail.htmlBody = pretty_(mail.body); mail.body = plain_(mail.body);
  MailApp.sendEmail(mail);
}

// ---------- customers change or cancel a booked lesson themselves ----------
const CHANGE_DAYS = 2;  // the date can be changed online up to this many days before; cancelling is always possible
function manageUrl_(data, b) { return ScriptApp.getService().getUrl() + '?act=mine&id=' + encodeURIComponent(data.id || '') + '&k=' + data.ckey + '&b=' + b; }
function manageLine_(data, ja) {
  if (data._b == null || !data.ckey) return ja ? '\n変更やキャンセルは、このメールにご返信ください。' : '\nTo change or cancel, just reply to this email.';
  return '\n\n' + (ja ? '日時の変更やキャンセルは、こちらからできます。' : 'You can change or cancel your lesson here:') + '\n' + manageUrl_(data, data._b);
}
function daysTo_(ymd) { return Math.round((dateOf_(ymd).getTime() - dateOf_(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd')).getTime()) / 86400000); }
// what happens to the money (or the pack) if this lesson is cancelled now
function refundOf_(data, bk) {
  const d = daysTo_(bk.date), per = data.place === 'FuncPhysio' ? Math.round(RATE_PER_30 * bk.len / 30) : 0;
  if (data.free) return { kind: 'free', ja: '無料相談なので、お支払いはありません。', en: 'This call is free, so there is nothing to refund.', mari: '無料相談' };
  if (bk.pack) return d >= 2
    ? { kind: 'packback', ja: '回数券に1回分をお戻しします。', en: 'The lesson goes back on your pack.', mari: '回数券に1回分戻しました' }
    : { kind: 'packused', ja: '前日・当日のキャンセルなので、回数券は1回使ったことになります。', en: 'As it is the day before or the day of the lesson, it counts as used on your pack.', mari: '前日・当日なので回数券は使ったまま' };
  if (d >= 2) return { kind: 'full', amt: per, ja: '全額お返しします。', en: 'You will get a full refund.', mari: '全額返金' + (per ? '（$' + per + '）' : '') };
  if (d === 1) return { kind: 'half', amt: Math.round(per / 2), ja: '前日のキャンセルなので、半分お返しします。', en: 'As it is the day before, you will get half back.', mari: '半額返金' + (per ? '（$' + Math.round(per / 2) + '）' : '') };
  return { kind: 'none', amt: 0, ja: '当日のキャンセルなので、お返しできません。', en: 'As it is the day of the lesson, there is no refund.', mari: '当日なので返金なし' };
}
function custBtn_(href, label, dark) {
  return '<a href="' + href + '" target="_top" style="display:block;margin:10px 0;padding:15px 12px;text-align:center;text-decoration:none;font-size:17px;letter-spacing:.04em;' + (dark ? 'background:#2c2a30;color:#ffffff;border:1px solid #2c2a30' : 'background:#ffffff;color:#2c2a30;border:1px solid #2c2a30') + '">' + label + '</a>';
}
function infoBox_(rows) {
  return '<table style="width:100%;border-collapse:collapse;margin:6px 0 18px;background:#faf7f1;border:1px solid #eee3cf">' + rows.map(([a, b]) => '<tr><td style="padding:10px 0 10px 14px;color:#8a7140;font-size:14px;white-space:nowrap;vertical-align:top">' + a + '</td><td style="padding:10px 14px;font-weight:600">' + b + '</td></tr>').join('') + '</table>';
}
function customerPage_(p, sh, r, data) {
  const ja = data.language === 'ja';
  const bk = (data.booked || [])[+p.b];
  if (!data.ckey || String(p.k) !== data.ckey || !bk) return page_(ja ? '開けません' : 'Link not valid', ja ? 'このリンクは使えません。このメールにご返信ください。' : 'This link can no longer be used. Please reply to the email instead.');
  const when = custWhen_(data, bk, bk.len, ja).replace(/\n/g, '<br>').replace(/(（[^）]*時間）| \([^)]*time\))/g, '<br><span style="font-weight:400;font-size:15px">$1</span>');
  const box = infoBox_([[ja ? '日時' : 'When', when], [ja ? '場所' : 'Where', esc_(placeText_(data, ja))], [ja ? 'レッスン' : 'Lesson', bk.len + (ja ? '分' : ' minutes')]]);
  if (bk.cancelled) return page_(ja ? 'キャンセル済みです' : 'This lesson is cancelled', box + (ja ? 'またのご予約をお待ちしています。' : 'Hope to see you another time.') + custBtn_(BOOKING_PAGE, ja ? '新しく予約する' : 'Book a new lesson', false));
  const s = atNY_(bk.date, ...bk.start.split(':').map(Number));
  if (s.getTime() < Date.now()) return page_(ja ? 'このレッスンは終わっています' : 'This lesson has passed', box + custBtn_(BOOKING_PAGE, ja ? '次のレッスンを予約する' : 'Book your next lesson', true));
  const base = manageUrl_(data, +p.b), rf = refundOf_(data, bk), d = daysTo_(bk.date);
  const pending = data.change && +data.change.b === +p.b;
  if (p.act === 'mine') {
    const policy = ja
      ? '<p style="font-size:15px;color:#555">キャンセルについて：<br><span style="white-space:nowrap">2日前まで 全額お返し</span>／<span style="white-space:nowrap">前日 半分</span>／<span style="white-space:nowrap">当日 お返しできません</span></p>'
      : '<p style="font-size:15px;color:#555">Cancelling: up to 2 days before, full refund; the day before, half; on the day, no refund.</p>';
    const changeBtn = pending
      ? '<p>' + (ja ? '<b>日時の変更をご希望中です</b>（' + whenIn_(data.change, bk.len, data.tz || TZ, true) + '）。Mari からのメールをお待ちください。' : '<b>You asked to move this lesson</b> to ' + whenIn_(data.change, bk.len, data.tz || TZ, false) + '. Mari will email you soon.') + '</p>'
      : d >= CHANGE_DAYS ? custBtn_(base.replace('act=mine', 'act=change'), ja ? '日時を変更する' : 'Change the date', true)
      : '<p>' + (ja ? 'レッスンの前日・当日の日時の変更は、確定メールにご返信ください。' : 'To move a lesson the day before or on the day, please reply to the confirmation email.') + '</p>';
    return page_(ja ? 'ご予約の変更・キャンセル' : 'Change or cancel', box + changeBtn + custBtn_(base.replace('act=mine', 'act=cancel'), ja ? 'キャンセルする' : 'Cancel the lesson', false) + policy);
  }
  if (p.act === 'cancel') {
    if (!p.go) return page_(ja ? 'キャンセルしますか？' : 'Cancel this lesson?', box + '<p>' + (ja ? '今キャンセルすると：' : 'If you cancel now: ') + '<b>' + (ja ? rf.ja : rf.en) + '</b></p>' + custBtn_(base.replace('act=mine', 'act=cancel') + '&go=1', ja ? 'キャンセルする' : 'Yes, cancel', true) + custBtn_(base, ja ? 'やめる（予約はそのまま）' : 'Keep my lesson', false));
    cancelLesson_(sh, r, data, +p.b, rf);
    return page_(ja ? 'キャンセルしました' : 'Your lesson is cancelled', box + '<p><b>' + (ja ? rf.ja : rf.en) + '</b>' + (rf.kind === 'full' || rf.kind === 'half' ? (ja ? '（お支払い済みの場合、Mari から返金します）' : ' (if you have paid, Mari will send it back)') : '') + '</p><p>' + (ja ? '確認のメールをお送りしました。' : 'A confirmation email is on its way.') + '</p>' + custBtn_(BOOKING_PAGE, ja ? '別の日に予約する' : 'Book another time', false));
  }
  // change: pick a new free time; Mari confirms it with one tap
  if (pending) return page_(ja ? '変更のご希望を受け付けています' : 'Your request is with Mari', box + (ja ? 'ご希望の日時：' : 'New time you asked for: ') + '<b>' + whenIn_(data.change, bk.len, data.tz || TZ, ja) + '</b><br>' + (ja ? 'Mari からのメールをお待ちください。' : 'Mari will email you soon.'));
  if (d < CHANGE_DAYS) return page_(ja ? 'ご返信ください' : 'Please reply to the email', ja ? 'レッスンの前日・当日の日時の変更は、確定メールにご返信ください。' : 'To move a lesson the day before or on the day, please reply to the confirmation email.');
  const free = freeTimes_(), slots = [];
  Object.keys(free).sort().forEach(k => {
    if (daysTo_(k) < 1 || daysTo_(k) > 60) return;
    const opts = [];
    free[k].forEach(([a, b2]) => { for (let m = Math.ceil(a / 30) * 30; m + bk.len <= b2; m += 30) opts.push({ date: k, start: Math.floor(m / 60) + ':' + String(m % 60).padStart(2, '0') }); });
    if (opts.length) slots.push(opts);
  });
  if (!p.go) {
    if (!slots.length) return page_(ja ? '空きがありません' : 'No free times', ja ? '今は空いている時間がありません。確定メールにご返信ください。' : 'There are no free times right now. Please reply to the confirmation email.');
    const sel = slots.map(day => '<optgroup label="' + (ja ? whenIn_(day[0], bk.len, data.tz || TZ, true).split('）')[0] + '）' : whenIn_(day[0], bk.len, data.tz || TZ, false).split(', ').slice(0, 2).join(', ')) + '">' + day.map(c => '<option value="' + c.date + ' ' + c.start + '">' + whenIn_(c, bk.len, data.tz || TZ, ja) + '</option>').join('') + '</optgroup>').join('');
    const u = ScriptApp.getService().getUrl();
    const form = '<form method="get" action="' + u + '" target="_top">'
      + '<input type="hidden" name="act" value="change"><input type="hidden" name="id" value="' + esc_(p.id) + '"><input type="hidden" name="k" value="' + esc_(p.k) + '"><input type="hidden" name="b" value="' + esc_(p.b) + '"><input type="hidden" name="go" value="1">'
      + '<p style="margin:0 0 6px;font-weight:600">' + (ja ? '新しい日時' : 'New time') + (data.tz && data.tz !== TZ ? '（' + esc_(zoneName_(data.tz, ja)) + '）' : '') + '</p>'
      + '<select name="slot" style="font-size:17px;padding:12px;width:100%;box-sizing:border-box;border:1px solid #c9b48c;background:#fff">' + sel + '</select>'
      + '<button type="submit" style="display:block;width:100%;margin:16px 0 8px;padding:15px 12px;background:#2c2a30;color:#fff;border:0;font-size:17px;letter-spacing:.04em">' + (ja ? 'この日時に変更をお願いする' : 'Ask to move to this time') + '</button></form>'
      + '<p style="font-size:15px;color:#555">' + (ja ? 'Mari が確認して、変更の確定メールをお送りします。それまでは今の日時のままです。お支払いはそのまま引き継がれます。' : 'Mari will check and email you to confirm. Until then your lesson stays at the current time. Your payment carries over.') + '</p>';
    return page_(ja ? '日時を変更する' : 'Change the date', '<p style="margin:0 0 4px">' + (ja ? '今のご予約：' : 'Now booked: ') + '</p>' + box + form);
  }
  const m = /^(\d{4}-\d\d-\d\d) (\d{1,2}:\d\d)$/.exec(String(p.slot || ''));
  const ok = m && slots.some(day => day.some(c => c.date === m[1] && c.start === m[2]));
  if (!ok) return page_(ja ? 'その時間は空いていません' : 'That time is not free', (ja ? 'ほかの時間を選んでください。' : 'Please choose another time.') + custBtn_(base.replace('act=mine', 'act=change'), ja ? '選び直す' : 'Choose again', true));
  data.change = { b: +p.b, date: m[1], start: m[2], at: Date.now() };
  sh.getRange(r, C['data']).setValue(JSON.stringify(data));
  sh.getRange(r, C['状況']).setValue('変更の希望');
  const me = Session.getEffectiveUser().getEmail(), u = ScriptApp.getService().getUrl(), key = String(sh.getRange(r, C['key']).getValue());
  const mb = (href, label, main) => '<a href="' + href + '" style="display:block;margin:8px 0;padding:14px 12px;text-align:center;background:' + (main ? '#2c2a30;color:#ffffff' : '#ffffff;color:#2c2a30') + ';border:1px solid #2c2a30;text-decoration:none;font-size:16px">' + label + '</a>';
  if (me) MailApp.sendEmail({ to: me, name: 'Lesson booking', subject: '【日時変更の希望】' + data.name + '（' + choiceJa_(bk, bk.len) + ' → ' + choiceJa_(data.change, bk.len) + '）',
    body: data.name + ' さんから日時変更の希望です。\n今：' + choiceJa_(bk, bk.len) + '\n希望：' + choiceJa_(data.change, bk.len),
    htmlBody: pretty_('<p><b>' + esc_(data.name) + '</b> さんから、日時変更の希望が届きました。</p>' + infoBox_([['今', choiceJa_(bk, bk.len)], ['希望', choiceJa_(data.change, bk.len)]])
      + mb(u + '?act=move&id=' + encodeURIComponent(data.id) + '&k=' + key, 'この日時に変更する', true) + mb(u + '?act=keep&id=' + encodeURIComponent(data.id) + '&k=' + key, '変更できない（今の日時のまま）', false)
      + '<p style="color:#6a6670;font-size:13px">変更すると、Google カレンダーも自動で移動し、お客さんに確定メールが届きます。</p>', true) });
  return page_(ja ? '変更のご希望を送りました' : 'Request sent', (ja ? 'ご希望の日時：' : 'New time you asked for: ') + '<b>' + whenIn_(data.change, bk.len, data.tz || TZ, ja) + '</b><p>' + (ja ? 'Mari が確認して、変更の確定メールをお送りします。それまでは今の日時のままです。' : 'Mari will check and email you to confirm. Until then your lesson stays at the current time.') + '</p>');
}
function eventOf_(data, bk) {
  const cal = CalendarApp.getDefaultCalendar();
  try { if (bk.ev) { const e = cal.getEventById(bk.ev); if (e) return e; } } catch (e) {}
  const s = atNY_(bk.date, ...bk.start.split(':').map(Number));
  return cal.getEvents(s, new Date(s.getTime() + bk.len * 60000)).filter(e => e.getStartTime().getTime() === s.getTime() && e.getTitle().indexOf(String(data.name)) > -1)[0] || null;
}
function refreshRow_(sh, r, data) {
  const live = (data.booked || []).filter(x => !x.cancelled).sort((a, b) => (a.date + a.start).localeCompare(b.date + b.start));
  sh.getRange(r, C['確定した日時']).setValue(live.length ? live.map(x => choiceJa_(x, x.len)).join(' ／ ') : 'キャンセル');
  sh.getRange(r, C['レッスン日']).setNumberFormat('@').setValue(live.length ? live[0].date + ' ' + live[0].start : '');
  sh.getRange(r, C['状況']).setValue(live.length ? (data.change ? '変更の希望' : '予約確定') : 'キャンセル');
  sh.getRange(r, C['data']).setValue(JSON.stringify(data));
  CacheService.getScriptCache().remove('free');
}
function cancelLesson_(sh, r, data, b, rf, byMari, msg) {
  const bk = data.booked[b], ja = data.language === 'ja';
  const ev = eventOf_(data, bk); if (ev) ev.deleteEvent();
  if (rf.kind === 'packback') { const ps = packSheet_(); const c = ps.getRange(bk.pack, 6); c.setValue(Math.max(0, +c.getValue() - 1)); }
  bk.cancelled = true; if (data.change && +data.change.b === b) data.change = null;
  refreshRow_(sh, r, data);
  const paid = sh.getRange(r, C['支払い']).getValue();
  const when = custWhen_(data, bk, bk.len, ja);
  const body = byMari ? (ja
    ? '# レッスンキャンセルのお知らせ\n\n' + data.name + ' 様\n\nご迷惑をおかけして、申し訳ございません。こちらの都合で、次のレッスンをキャンセルさせてください。\n\n日時：' + when + '\nレッスン：' + bk.len + '分\n' + (msg ? '\n' + msg + '\n' : '') + '\n' + rf.ja + (paid === true && rf.kind === 'full' ? '\nVenmo でお返しします。' : '') + '\n\nご都合のよい日時で、また予約していただけたらうれしいです。\n' + BOOKING_PAGE + '\n\n本当に申し訳ございません。\n髙山まり'
    : '# Your lesson is cancelled\n\nHi ' + data.name + ',\n\nI’m very sorry for the inconvenience — I have to cancel this lesson.\n\nWhen: ' + when + '\nLesson: ' + bk.len + ' minutes\n' + (msg ? '\n' + msg + '\n' : '') + '\n' + rf.en + (paid === true && rf.kind === 'full' ? '\nI will send it back by Venmo.' : '') + '\n\nI would love to see you at another time:\n' + BOOKING_PAGE + '\n\nMari Takayama')
    : ja
    ? '# キャンセル受付のお知らせ\n\n' + data.name + ' 様\n\nレッスンのキャンセルを承りました。\n\n日時：' + when + '\nレッスン：' + bk.len + '分\n\n' + rf.ja + (rf.kind === 'full' || rf.kind === 'half' ? (paid === true ? '\nVenmo でお返しします。' : '') : '') + '\n\nまたのご予約をお待ちしています。\n' + BOOKING_PAGE + '\n\n髙山まり'
    : '# Cancellation confirmed\n\nHi ' + data.name + ',\n\nYour lesson is cancelled.\n\nWhen: ' + when + '\nLesson: ' + bk.len + ' minutes\n\n' + rf.en + (rf.kind === 'full' || rf.kind === 'half' ? (paid === true ? '\nMari will send it back by Venmo.' : '') : '') + '\n\nHope to see you another time:\n' + BOOKING_PAGE + '\n\nMari Takayama';
  if (data.email) MailApp.sendEmail({ to: data.email, subject: byMari ? (ja ? 'レッスンのキャンセルのお知らせ（' + whenIn_(bk, bk.len, data.tz || TZ, true) + '）' : 'Your lesson is cancelled — ' + whenIn_(bk, bk.len, data.tz || TZ, false)) : ja ? 'キャンセルを承りました（' + whenIn_(bk, bk.len, data.tz || TZ, true) + '）' : 'Your lesson is cancelled — ' + whenIn_(bk, bk.len, data.tz || TZ, false), body: plain_(body), htmlBody: pretty_(body), name: SENDER_NAME, replyTo: REPLY_TO });
  const me = Session.getEffectiveUser().getEmail();
  const todo = rf.kind === 'full' || rf.kind === 'half' ? (paid === true ? '支払い済みなので、Venmo で ' + rf.mari + ' をお願いします。' : 'まだ支払いのチェックがないので、返金はいりません（受け取っていたら ' + rf.mari + '）。') : rf.mari + '。';
  if (byMari) return todo;
  if (me) MailApp.sendEmail({ to: me, name: 'Lesson booking', subject: '【キャンセル】' + data.name + '（' + choiceJa_(bk, bk.len) + '）',
    body: data.name + ' さんがキャンセルしました。' + choiceJa_(bk, bk.len) + '\n' + todo,
    htmlBody: pretty_('<p><b>' + esc_(data.name) + '</b> さんが、レッスンをキャンセルしました。Google カレンダーからも消しました。</p>' + infoBox_([['日時', choiceJa_(bk, bk.len)], ['返金', esc_(rf.mari)], ['メール', esc_(data.email)]]) + '<p>' + esc_(todo) + '</p>', true) });
}
function moveTo_(sh, r, data, b, ch) {
  const bk = data.booked[b], old = { date: bk.date, start: bk.start };
  const s = atNY_(ch.date, ...ch.start.split(':').map(Number)), en = new Date(s.getTime() + bk.len * 60000);
  const ev = eventOf_(data, bk);
  if (ev) ev.setTime(s, en); else bk.ev = CalendarApp.getDefaultCalendar().createEvent('バレエレッスン：' + data.name, s, en, { location: placeText_(data, false), description: 'メール：' + data.email }).getId();
  bk.date = ch.date; bk.start = ch.start; data._b = b;
  refreshRow_(sh, r, data);
  return old;
}
// ---------- Mari changes or cancels a booked lesson ----------
function mariUrl_(sh, r, data, act, b) {
  return ScriptApp.getService().getUrl() + '?act=' + act + '&id=' + encodeURIComponent(data.id || '') + '&k=' + String(sh.getRange(r, C['key']).getValue()) + '&b=' + (b == null ? data._b : b);
}
function fullRefund_(data, bk) {
  const per = data.place === 'FuncPhysio' ? Math.round(RATE_PER_30 * bk.len / 30) : 0;
  if (data.free) return { kind: 'free', ja: '無料相談なので、お支払いはありません。', en: 'This call was free, so there is nothing to refund.', mari: '無料相談' };
  if (bk.pack) return { kind: 'packback', ja: '回数券に1回分をお戻しします。', en: 'The lesson goes back on your pack.', mari: '回数券に1回分戻しました' };
  return { kind: 'full', amt: per, ja: '全額お返しします。', en: 'You will get a full refund.', mari: '全額返金' + (per ? '（$' + per + '）' : '') };
}
// ready-made notes Mari can tap when she moves or cancels a lesson (customers who chose English get the English ones)
const MEMOS_JA = {
  move: ['急にリハーサルが入ってしまいました。', '公演の予定が入ってしまいました。', 'スタジオが使えなくなってしまいました。', '体調を崩してしまい、念のため日にちを変えさせてください。', 'ご都合が合わなければ、遠慮なくキャンセルしてください。'],
  cancel: ['急にリハーサルが入ってしまいました。', '公演の予定が入ってしまいました。', 'スタジオが使えなくなってしまいました。', '体調を崩してしまいました。', 'またご都合のよい日に、ぜひお会いできたらうれしいです。']
};
const MEMOS_EN = {
  move: ['A rehearsal came up suddenly.', 'I have a performance scheduled at that time.', 'The studio is no longer available at that time.', 'I’m not feeling well, so I’d like to move the lesson to be safe.', 'If the new time doesn’t work for you, please feel free to cancel.'],
  cancel: ['A rehearsal came up suddenly.', 'I have a performance scheduled at that time.', 'The studio is no longer available at that time.', 'I’m not feeling well.', 'I hope to see you at another time soon.']
};
function mariChange_(p, sh, r, data) {
  const bk = (data.booked || [])[+p.b];
  if (!bk) return page_('見つかりません', 'このレッスンが見つかりませんでした（ボタンができる前の予約は、お客さんにメールで連絡してください）。');
  if (bk.cancelled) return page_('キャンセル済みです', 'このレッスンは、もうキャンセルされています。');
  const head = '<b>' + esc_(data.name) + '</b> さん<br>' + choiceJa_(bk, bk.len) + '（ニューヨーク時間）';
  const u = ScriptApp.getService().getUrl();
  const hidden = '<input type="hidden" name="act" value="' + p.act + '"><input type="hidden" name="id" value="' + esc_(p.id) + '"><input type="hidden" name="k" value="' + esc_(p.k) + '"><input type="hidden" name="b" value="' + esc_(p.b) + '"><input type="hidden" name="go" value="1">';
  const en = data.language !== 'ja';
  const memos = (en ? MEMOS_EN : MEMOS_JA)[p.act === 'mcancel' ? 'cancel' : 'move'];
  const chips = memos.map(m => '<button type="button" onclick="var t=document.getElementById(\'msg\');t.value=(t.value?t.value+\'\\n\':\'\')+this.textContent;t.focus()" style="display:block;width:100%;text-align:left;margin:6px 0;padding:10px 12px;background:#faf7f1;border:1px solid #e6dac3;color:#222;font-size:15px;line-height:1.5">' + esc_(m) + '</button>').join('');
  const msgBox = '<p style="margin:16px 0 4px;font-weight:600">お客さんへのメモ（' + (en ? '英語で送られます' : '日本語') + '）</p><p style="margin:0 0 6px;font-size:14px;color:#555">押すと下に入ります（いくつ選んでも大丈夫・書き直しもできます）。</p>' + chips
    + '<textarea id="msg" name="msg" rows="4" style="font-size:16px;padding:10px;width:100%;box-sizing:border-box;margin-top:8px;border:1px solid #c9b48c" placeholder="自由に書いても大丈夫です（なくても送れます）"></textarea>';
  const submit = label => '<button type="submit" style="display:block;width:100%;margin:14px 0;padding:15px 12px;background:#2c2a30;color:#fff;border:0;font-size:17px">' + label + '</button>';
  if (p.act === 'mcancel') {
    const rf = fullRefund_(data, bk);
    if (!p.go) return page_('キャンセルしますか？', head + '<p>Mari さんの都合のキャンセルなので、お客さんには「' + rf.ja + '」とお知らせします。Google カレンダーからも消します。</p><form method="get" action="' + u + '" target="_top">' + hidden + msgBox + submit('キャンセルして、お客さんにメールを送る') + '</form>');
    const todo = cancelLesson_(sh, r, data, +p.b, rf, true, String(p.msg || '').slice(0, 500));
    return page_('キャンセルしました', head + '<p>お客さんにお知らせのメールを送り、Google カレンダーから消しました。</p><p><b>' + esc_(todo) + '</b></p>');
  }
  // mmove: Mari picks a new time; the customer says OK or cancels with a full refund
  if (!p.go) {
    const opt = (v, cur) => '<option value="' + v + '"' + (v === cur ? ' selected' : '') + '>' + v + '</option>';
    let times = ''; for (let m = 7 * 60; m <= 22 * 60; m += 15) times += opt(Math.floor(m / 60) + ':' + String(m % 60).padStart(2, '0'), bk.start.replace(/^0/, ''));
    return page_('日時を変更する', head + '<form method="get" action="' + u + '" target="_top" style="margin-top:14px">' + hidden
      + '<p>新しい日にち<br><input type="date" name="date" value="' + bk.date + '" style="font-size:17px;padding:8px;width:100%;box-sizing:border-box"></p>'
      + '<p>始まる時間<br><select name="start" style="font-size:17px;padding:8px;width:100%">' + times + '</select></p>' + msgBox
      + submit('この日時をお客さんに提案する') + '</form><p style="font-size:15px;color:#555">お客さんが「OK」を押すと、自動で変更されてカレンダーも移ります。合わないときは、お客さんがキャンセルでき、全額返金（回数券なら1回分戻す）になります。それまでは今の日時のままです。</p>');
  }
  if (!/^\d{4}-\d\d-\d\d$/.test(p.date || '') || !/^\d{1,2}:\d\d$/.test(p.start || '')) return page_('入力を確かめてください', '日にちと時間を選んでください。');
  data.offer = { b: +p.b, date: p.date, start: p.start };
  sh.getRange(r, C['data']).setValue(JSON.stringify(data));
  const ja = data.language === 'ja', ck = ScriptApp.getService().getUrl() + '?id=' + encodeURIComponent(data.id) + '&k=' + data.ckey + '&b=' + p.b;
  const msg = String(p.msg || '').slice(0, 500);
  const body = ja
    ? '# レッスン日時変更のお願い\n\n' + data.name + ' 様\n\nご迷惑をおかけして、申し訳ございません。こちらの都合で、レッスンの日時を変更させていただけないでしょうか。\n\n今のご予約：' + whenIn_(bk, bk.len, data.tz || TZ, true) + '\n新しい日時：' + custWhen_(data, data.offer, bk.len, true) + '\n' + (msg ? '\n' + msg + '\n' : '') + '\n## 新しい日時でよろしい場合\n「この日時で OK」を押してください。予約が変更されます。\n' + ck.replace('?', '?act=take&') + '\n\n## ご都合が合わない場合（キャンセル）\nキャンセルにして、全額お返しします' + (bk.pack ? '（回数券なら1回分お戻しします）' : '') + '。\n' + ck.replace('?', '?act=decl&') + '\n\nご迷惑をおかけして、本当に申し訳ございません。\n髙山まり'
    : '# Request to move your lesson\n\nHi ' + data.name + ',\n\nI’m very sorry for the inconvenience — could we move your lesson?\n\nNow booked: ' + whenIn_(bk, bk.len, data.tz || TZ, false) + '\nNew time: ' + custWhen_(data, data.offer, bk.len, false) + '\n' + (msg ? '\n' + msg + '\n' : '') + '\n## If the new time works\nTap “This time works” and your lesson will be moved.\n' + ck.replace('?', '?act=take&') + '\n\n## If it doesn’t work (cancel)\nYou can cancel and get a full refund' + (bk.pack ? ' (or the lesson back on your pack)' : '') + '.\n' + ck.replace('?', '?act=decl&') + '\n\nI’m truly sorry for the inconvenience.\nMari Takayama';
  MailApp.sendEmail({ to: data.email, subject: ja ? 'レッスンの日時変更のお願い' : 'Could we move your lesson?', body: plain_(body), htmlBody: pretty_(body), name: SENDER_NAME, replyTo: REPLY_TO });
  return page_('提案を送りました', head + '<br>→ <b>' + choiceJa_(data.offer, bk.len) + '</b><p>お客さんが OK すると、自動で変更されて Mari さんにもメールが届きます。</p>');
}
// the customer answers Mari's new time
function offerAnswer_(p, sh, r, data) {
  const ja = data.language === 'ja', of = data.offer, bk = of && (data.booked || [])[of.b];
  if (!data.ckey || String(p.k) !== data.ckey || !of || +p.b !== of.b || !bk || bk.cancelled) return page_(ja ? '対応済みです' : 'Already done', ja ? 'このご提案は、もう終わっています。' : 'This request has already been answered.');
  const self = ScriptApp.getService().getUrl() + '?act=' + p.act + '&id=' + encodeURIComponent(p.id) + '&k=' + encodeURIComponent(p.k) + '&b=' + p.b + '&go=1';
  const me = Session.getEffectiveUser().getEmail();
  if (p.act === 'take') {
    if (!p.go) return page_(ja ? 'この日時に変更しますか？' : 'Move to this time?', '<b>' + custWhen_(data, of, bk.len, ja).replace(/\n/g, '<br>') + '</b>', self, ja ? 'この日時で OK' : 'Yes, this time works');
    data.offer = null; const old = moveTo_(sh, r, data, of.b, of);
    const body = ja ? '# レッスン日時変更のお知らせ\n\n' + data.name + ' 様\n\nご迷惑をおかけして申し訳ございません。ご協力ありがとうございます。レッスンの日時を変更しました。\n\n日時：' + custWhen_(data, bk, bk.len, true) + '\n場所：' + placeText_(data, true) + '\nレッスン：' + bk.len + '分\n\n（変更前：' + whenIn_(old, bk.len, data.tz || TZ, true) + '）\n\nお支払いは、そのまま新しい日時に引き継がれます。' + manageLine_(data, true) + '\n\n髙山まり'
      : '# Your lesson has been moved\n\nHi ' + data.name + ',\n\nThank you! Your lesson has been moved.\n\nWhen: ' + custWhen_(data, bk, bk.len, false) + '\nWhere: ' + placeText_(data, false) + '\nLesson: ' + bk.len + ' minutes\n\n(Before: ' + whenIn_(old, bk.len, data.tz || TZ, false) + ')\n\nYour payment carries over to the new time.' + manageLine_(data, false) + '\n\nMari Takayama';
    MailApp.sendEmail({ to: data.email, subject: ja ? 'レッスンの日時を変更しました（' + whenIn_(bk, bk.len, data.tz || TZ, true) + '）' : 'Your lesson has been moved — ' + whenIn_(bk, bk.len, data.tz || TZ, false), body: plain_(body), htmlBody: pretty_(body), name: SENDER_NAME, replyTo: REPLY_TO, attachments: [ics_(data, bk, bk.len)] });
    if (me) MailApp.sendEmail({ to: me, name: 'Lesson booking', subject: '【変更OK】' + data.name + '（' + choiceJa_(old, bk.len) + ' → ' + choiceJa_(bk, bk.len) + '）', body: data.name + ' さんが新しい日時を OK しました。', htmlBody: pretty_('<p><b>' + esc_(data.name) + '</b> さんが、新しい日時を OK しました。Google カレンダーも移しました。</p>' + infoBox_([['前', choiceJa_(old, bk.len)], ['新しい日時', choiceJa_(bk, bk.len)]]), true) });
    return page_(ja ? '変更しました' : 'Your lesson is moved', (ja ? '確認のメールをお送りしました。<br>日時：' : 'A confirmation email is on its way.<br>When: ') + custWhen_(data, bk, bk.len, ja).replace(/\n/g, '<br>'));
  }
  const rf = fullRefund_(data, bk);
  if (!p.go) return page_(ja ? 'キャンセルしますか？' : 'Cancel this lesson?', whenIn_(bk, bk.len, data.tz || TZ, ja) + '<p><b>' + (ja ? rf.ja : rf.en) + '</b></p>', self, ja ? 'キャンセルする' : 'Yes, cancel');
  data.offer = null; cancelLesson_(sh, r, data, of.b, rf, false);
  return page_(ja ? 'キャンセルしました' : 'Your lesson is cancelled', '<b>' + (ja ? rf.ja : rf.en) + '</b><p>' + (ja ? 'ご迷惑をおかけして、申し訳ございませんでした。' : 'Sorry again for the trouble.') + '</p>' + custBtn_(BOOKING_PAGE, ja ? '別の日に予約する' : 'Book another time', false));
}

// Mari answers a change request
function moveAnswer_(p, sh, r, data) {
  const ch = data.change; if (!ch) return page_('対応済みです', '変更の希望はもうありません。');
  const bk = data.booked[ch.b], self = ScriptApp.getService().getUrl() + '?act=' + p.act + '&id=' + encodeURIComponent(p.id) + '&k=' + encodeURIComponent(p.k) + '&go=1';
  const ja = data.language === 'ja';
  if (p.act === 'keep') {
    if (!p.go) return page_('変更をお断りしますか？', '<b>' + esc_(data.name) + '</b> さん<br>今：' + choiceJa_(bk, bk.len) + '<br>希望：' + choiceJa_(ch, bk.len) + '<p>お客さんに「その日時は都合がつかないので、今の日時のままです」とメールします。</p>', self, 'お断りのメールを送る');
    data.change = null; refreshRow_(sh, r, data);
    const body = ja
      ? data.name + ' 様\n\n申し訳ありません、ご希望の日時（' + whenIn_(ch, bk.len, data.tz || TZ, true) + '）は都合がつきませんでした。レッスンは今の日時のままです。\n\n日時：' + custWhen_(data, bk, bk.len, true) + '\n\nほかの日時をご希望でしたら、こちらからもう一度お選びください。\n' + manageUrl_(data, ch.b) + '\n\n髙山まり'
      : 'Hi ' + data.name + ',\n\nI’m sorry — the time you asked for (' + whenIn_(ch, bk.len, data.tz || TZ, false) + ') doesn’t work for me. Your lesson stays at:\n\nWhen: ' + custWhen_(data, bk, bk.len, false) + '\n\nTo choose another time:\n' + manageUrl_(data, ch.b) + '\n\nMari Takayama';
    MailApp.sendEmail({ to: data.email, subject: ja ? 'レッスンの日時について' : 'About moving your lesson', body: plain_(body), htmlBody: pretty_(body), name: SENDER_NAME, replyTo: REPLY_TO });
    return page_('送りました', esc_(data.name) + ' さんに、今の日時のままだとお知らせしました。');
  }
  if (!p.go) return page_('この日時に変更しますか？', '<b>' + esc_(data.name) + '</b> さん<br>今：' + choiceJa_(bk, bk.len) + '<br>新しい日時：<b>' + choiceJa_(ch, bk.len) + '</b><p>Google カレンダーの予定を移して、お客さんに変更の確定メールを送ります。</p>', self, '変更して、お客さんにメールを送る');
  data.change = null;
  const old = moveTo_(sh, r, data, ch.b, ch);
  const paid = sh.getRange(r, C['支払い']).getValue();
  const payPart = paid === false ? payNote_(data, bk, ja) : (ja ? 'お支払いは、そのまま新しい日時に引き継がれます。' : 'Your payment carries over to the new time.');
  const body = ja
    ? '# レッスン日時変更のお知らせ\n\n' + data.name + ' 様\n\nレッスンの日時の変更が確定しました。\n\n日時：' + custWhen_(data, bk, bk.len, true) + '\n場所：' + placeText_(data, true) + '\nレッスン：' + bk.len + '分\n\n（変更前：' + whenIn_(old, bk.len, data.tz || TZ, true) + '）\n\n' + payPart + manageLine_(data, true) + '\n\nお会いできるのを楽しみにしています。\n髙山まり'
    : '# Your lesson has been moved\n\nHi ' + data.name + ',\n\nYour lesson has been moved.\n\nWhen: ' + custWhen_(data, bk, bk.len, false) + '\nWhere: ' + placeText_(data, false) + '\nLesson: ' + bk.len + ' minutes\n\n(Before: ' + whenIn_(old, bk.len, data.tz || TZ, false) + ')\n\n' + payPart + manageLine_(data, false) + '\n\nSee you soon!\nMari Takayama';
  MailApp.sendEmail({ to: data.email, subject: ja ? 'レッスンの日時を変更しました（' + whenIn_(bk, bk.len, data.tz || TZ, true) + '）' : 'Your lesson has been moved — ' + whenIn_(bk, bk.len, data.tz || TZ, false), body: plain_(body), htmlBody: pretty_(body), name: SENDER_NAME, replyTo: REPLY_TO, attachments: [ics_(data, bk, bk.len)] });
  return page_('変更しました', esc_(data.name) + ' さんのレッスンを ' + choiceJa_(bk, bk.len) + ' に変更しました。Google カレンダーも移しました。');
}

// ---------- payment deadline and reminders ----------
const PAY_DAYS = 3;  // lessons must be paid at least this many days before
function payBy_(ch) {
  const d = dateOf_(ch.date); d.setDate(d.getDate() - PAY_DAYS);
  const today = dateOf_(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'));
  if (d.getTime() <= today.getTime()) return { ja: '今日中', en: 'today', late: true };
  return { ja: (d.getMonth() + 1) + '/' + d.getDate() + '（' + WD_JA[d.getDay()] + '）', en: Utilities.formatDate(d, TZ, 'EEE, MMM d'), late: false };
}
// run once from the editor: sends payment reminders every morning
function setupPaymentReminders() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'paymentReminders').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('paymentReminders').timeBased().everyDays(1).atHour(9).inTimezone(TZ).create();
  heardStats_();
}
// test bookings (name starting with TEST, or with テスト) leave Mari's calendar 30 minutes after they are confirmed
function isTest_(name) { name = String(name || ''); return /^\s*test\b/i.test(name) || /テスト/.test(name); }
function cleanupTests_() {
  const now = Date.now(), from = new Date(now - 2 * 86400000), to = new Date(now + (DAYS_AHEAD + 7) * 86400000);
  let n = 0;
  CalendarApp.getDefaultCalendar().getEvents(from, to).forEach(ev => {
    const title = ev.getTitle();
    if (!/^(バレエレッスン|無料 Zoom 相談)：/.test(title) || !isTest_(title.split('：').slice(1).join('：'))) return;
    if (ev.getDateCreated().getTime() > now - 30 * 60000) return;
    ev.deleteEvent(); n++;
  });
  if (n) CacheService.getScriptCache().remove('free');
  return n;
}
function paymentReminders() {
  const sh = book_().getSheetByName(REQ); if (!sh || sh.getLastRow() < 2) return;
  heardStats_();
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, HEAD.length).getValues();
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), t0 = dateOf_(today).getTime(), me = Session.getEffectiveUser().getEmail();
  v.forEach((row, i) => {
    const paid = row[C['支払い'] - 1], when = String(row[C['レッスン日'] - 1] || '');
    if (!/^(予約確定|変更の希望)$/.test(row[C['状況'] - 1]) || paid !== false || !when) return;
    const [date, start] = when.split(' '); if (date < today) return;
    let data; try { data = JSON.parse(row[C['data'] - 1]); } catch (e) { return; }
    if (isTest_(data.name)) return;
    const ch = { date, start: start || '19:00' }, days = Math.round((dateOf_(date).getTime() - t0) / 86400000);
    const log = String(row[C['催促'] - 1] || ''), stamp = Utilities.formatDate(new Date(), TZ, 'M/d');
    let tag = '';
    if (days <= PAY_DAYS && !/3日前/.test(log)) tag = '3日前';
    else if (days === PAY_DAYS + 1 && !/4日前/.test(log)) tag = '4日前';
    else if (!log && new Date(row[0]).getTime() < Date.now() - 86400000) tag = '翌日';
    if (!tag) return;
    data.len = +row[C['レッスンの長さ（分）'] - 1] || data.len || 60;
    sendPayReminder_(data, ch, tag);
    if (tag === '3日前' && me) MailApp.sendEmail({ to: me, subject: '【未払い】' + data.name + '（' + choiceJa_(ch, data.len) + '）', name: 'Lesson booking',
      body: data.name + ' さんのレッスン代が、まだ「支払い」にチェックされていません。\n\n日時：' + choiceJa_(ch, data.len) + '\nメール：' + data.email + '\n\nお客さんにはもう一度お支払いのお願いを送りました。受け取ったらシートの「支払い」にチェックを入れてください。キャンセルにするかどうかは Mari さんが決めてください。',
      htmlBody: pretty_('<p><b>' + esc_(data.name) + '</b> さんのレッスン代が、まだ「支払い」にチェックされていません。</p><p>日時：' + choiceJa_(ch, data.len) + '<br>メール：' + esc_(data.email) + '</p><p>お客さんにはもう一度お支払いのお願いを送りました。受け取ったらシートの「支払い」にチェックを入れてください。キャンセルにするかどうかは Mari さんが決めてください。</p>', true) });
    sh.getRange(i + 2, C['催促']).setValue((log ? log + '、' : '') + stamp + ' ' + tag);
  });
}
function sendPayReminder_(data, ch, tag) {
  if (!data.email) return;
  const ja = data.language === 'ja', when = custWhen_(data, ch, data.len, ja), dl = payBy_(ch);
  const per = data.place === 'FuncPhysio' ? Math.round(RATE_PER_30 * data.len / 30) : 0;
  const d = dateOf_(ch.date), memo = 'Ballet lesson ' + (d.getMonth() + 1) + '/' + d.getDate();
  const link = VENMO ? 'https://venmo.com/' + VENMO + '?txn=pay' + (per ? '&amount=' + per : '') + '&note=' + encodeURIComponent(memo) : '';
  const last = tag === '3日前';
  const subject = ja ? (last ? '【お支払いのお願い】' : 'お支払いのお願い：') + 'バレエレッスン（' + whenIn_(ch, data.len, data.tz || TZ, true) + '）' : (last ? 'Payment needed: ' : 'Payment reminder: ') + 'your ballet lesson on ' + whenIn_(ch, data.len, data.tz || TZ, false);
  const body = ja
    ? data.name + ' 様\n\n' + (last ? 'レッスン代のお支払いがまだ確認できていません。ご予約を続けるため、今日中にお支払いをお願いします。' : 'レッスン代のお支払いがまだ確認できていません。' + (dl.late ? '今日中' : dl.ja + 'まで') + 'にお支払いをお願いします（レッスンの' + PAY_DAYS + '日前まで）。') + '\n\n日時：' + when + (per ? '\n金額：$' + per : '') + (link ? '\n\nVenmo：こちらを開くと、金額が入った状態で開きます\n' + link : '') + '\n\nもうお支払い済みでしたら、行き違いですみません。このメールに返信してお知らせください。\n\nMari Takayama'
    : 'Hi ' + data.name + ',\n\n' + (last ? 'I haven’t received the payment for your lesson yet. To keep your booking, please pay today.' : 'I haven’t received the payment for your lesson yet. Please pay by ' + (dl.late ? 'today' : dl.en) + ' (' + PAY_DAYS + ' days before the lesson).') + '\n\nWhen: ' + when + (per ? '\nAmount: $' + per : '') + (link ? '\n\nVenmo — this link opens with everything filled in:\n' + link : '') + '\n\nIf you have already paid, sorry for the mix-up — just reply to let me know.\n\nMari Takayama';
  MailApp.sendEmail({ to: data.email, subject, body: plain_(body), htmlBody: pretty_(body), name: SENDER_NAME, replyTo: REPLY_TO });
}
// elegant HTML version of an email (plain text in, or html when isHtml)
function goalsJa_(d) { return d.visit === 'First' && Array.isArray(d.goals) ? d.goals.slice(0, 8).map(g => JA.goals[g] || '').filter(Boolean).join('、') : ''; }
function plain_(s) { return String(s).replace(/^#{1,2} /gm, ''); }
function sendPretty_(m) { m.htmlBody = pretty_(m.body); m.body = plain_(m.body); m.body += '\n\n' + SITE; MailApp.sendEmail(m); }
function pretty_(text, isHtml) {
  const jaMail = /[\u3040-\u30ff]/.test(text);
  const link = u => /[?&]act=(mmove|mcancel|take|decl)&/.test(u)
    ? '<a href="' + u + '" style="display:block;margin:6px 0;padding:14px 12px;text-align:center;text-decoration:none;font-size:16px;letter-spacing:.04em;' + (/act=(mmove|take)&/.test(u) ? 'background:#2c2a30;color:#ffffff;border:1px solid #2c2a30' : 'background:#ffffff;color:#2c2a30;border:1px solid #2c2a30') + '">' + ({ mmove: '日時を変更する（お客さんに提案）', mcancel: 'キャンセルする', take: jaMail ? 'この日時で OK' : 'This time works', decl: jaMail ? 'キャンセルする（全額返金）' : 'Doesn’t work — cancel with a full refund' })[u.match(/act=(\w+)/)[1]] + '</a>'
    : u === BOOKING_PAGE
    ? '<a href="' + u + '" style="display:block;margin:6px 0;padding:14px 12px;text-align:center;text-decoration:none;font-size:16px;letter-spacing:.04em;background:#ffffff;color:#2c2a30;border:1px solid #2c2a30">' + (jaMail ? '予約ページを開く' : 'Open the booking page') + '</a>'
    : /[?&]act=mine/.test(u)
    ? '<a href="' + u + '" style="display:block;margin:4px 0;padding:14px 12px;background:#ffffff;border:1px solid #2c2a30;color:#2c2a30;text-align:center;text-decoration:none;font-size:17px;letter-spacing:.04em">' + (jaMail ? '日時の変更・キャンセル' : 'Change or cancel') + '</a>'
    : /venmo\.com/.test(u)
    ? '<a href="' + u + '" style="display:block;margin:4px 0;padding:16px 12px;background:#2c2a30;color:#ffffff;text-align:center;text-decoration:none;font-size:18px;font-weight:600;letter-spacing:.06em">' + (jaMail ? 'Venmo で今すぐ支払う &rarr;' : 'Pay now with Venmo &rarr;') + '</a>'
    : '<a href="' + u + '" style="color:#7a6233;word-break:break-all">' + u + '</a>';
  const fmt = s => esc_(s).replace(/(https?:\/\/[^\s<]+)/g, link);
  const isInfo = l => /^[^：:\s][^：:]{0,9}[：:]\s?\S/.test(l) && !/^https?:/.test(l);
  const block = p => {
    if (/^# /.test(p)) return '<div style="font-size:21px;font-weight:600;letter-spacing:.04em;color:#2c2a30;margin:0 0 22px;padding:0 0 10px;border-bottom:2px solid #c9b48c">' + esc_(p.slice(2)) + '</div>';
    if (/^## /.test(p)) { const ls = p.split('\n'); return '<div style="font-size:17px;font-weight:600;color:#2c2a30;margin:28px 0 8px;padding-left:10px;border-left:3px solid #c9b48c">' + esc_(ls.shift().slice(3)) + '</div>' + (ls.length ? block(ls.join('\n')) : ''); }
    const lines = p.split('\n');
    // "When: …  Where: …" style lines -> a soft box with labels and values
    if (lines.length > 1 && lines.every(isInfo)) {
      return '<table role="presentation" style="width:100%;border-collapse:collapse;margin:0 0 20px;background:#faf7f1;border:1px solid #eee3cf">' + lines.map(l => {
        const m = l.match(/^([^：:]+)[：:]\s?(.*)$/);
        return '<tr><td style="padding:10px 0 10px 14px;color:#8a7140;font-size:14px;white-space:nowrap;vertical-align:top">' + esc_(m[1]) + '</td><td style="padding:10px 14px;font-size:17px;color:#222222;font-weight:600">' + fmt(m[2]) + '</td></tr>';
      }).join('') + '</table>';
    }
    // a heading line ending with ：/: followed by bullet lines
    const head = /[：:]$/.test(lines[0]) && lines.length > 1 ? '<div style="font-weight:600;margin:0 0 6px">' + fmt(lines.shift()) + '</div>' : '';
    const body = lines.map(l => /^[・•]/.test(l) ? '<div style="padding-left:1em;text-indent:-1em;margin:2px 0">' + fmt(l) + '</div>' : fmt(l)).join('<br>').replace(/<\/div><br>/g, '</div>');
    return '<div style="margin:0 0 20px">' + head + body + '</div>';
  };
  const inner = isHtml ? text : text.split(/\n\n+/).map(block).join('');
  return '<div style="background:#faf7f1;padding:24px 10px"><div style="max-width:560px;margin:0 auto;background:#ffffff;border:1px solid #e6dac3;padding:30px 24px;font-family:-apple-system,BlinkMacSystemFont,\'Hiragino Sans\',\'Hiragino Kaku Gothic ProN\',\'Yu Gothic\',YuGothic,Meiryo,\'Helvetica Neue\',Arial,sans-serif;font-size:17px;line-height:1.9;letter-spacing:.02em;color:#222222">'
    + '<div style="text-align:center;margin:0 0 24px;padding-bottom:16px;border-bottom:1px solid #e6dac3"><div style="font-family:Georgia,\'Times New Roman\',serif;font-size:24px;letter-spacing:.08em;color:#2c2a30">Mari Takayama</div><div style="font-size:11px;letter-spacing:.32em;color:#8a7140;text-transform:uppercase;margin-top:4px">Ballet · New York</div></div>'
    + inner + '<div style="text-align:center;margin-top:28px;padding-top:18px;border-top:1px solid #e6dac3"><a href="' + SITE + '" style="font-family:Georgia,\'Times New Roman\',serif;font-size:15px;letter-spacing:.12em;color:#7a6233;text-decoration:none">Mari Takayama &middot; Website &rarr;</a></div></div></div>';
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
  if (data.place === 'Studio' && data.studio === 'self') return ja ? 'ダンススタジオ（ご自身で予約されたスタジオ）' : 'A dance studio (the one you book)';
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
