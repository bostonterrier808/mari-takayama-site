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
const REQ = 'リクエスト', AV = '空き状況', ST = '統計';
const HEAD = ['受付日時', '初めて／2回目以降', '名前', '連絡方法', '連絡先', '場所', 'エリア', '対象', 'バレエ経験',
  '第1希望', '第2希望', '第3希望', '第1希望の曜日', '第1希望の時間帯', '第2希望の曜日', '第2希望の時間帯', '第3希望の曜日', '第3希望の時間帯',
  'メッセージ', '言語', '状況'];
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
    const row = [new Date(), JA.visit[d.visit] || d.visit || '', clip(d.name), d.via || '', clip(d.contact), JA.place[d.place] || d.place || '', clip(d.area),
      JA.age[d.age] || d.age || '', JA.level[d.level] || d.level || '', choiceText(0), choiceText(1), choiceText(2)];
    for (let i = 0; i < 3; i++) row.push(c[i] ? (JA.day[c[i].weekday] || c[i].weekday) : '', c[i] ? (JA.time[c[i].time] || c[i].time) : '');
    row.push(clip(d.note, 700), d.language || '', '新規');
    book_().getSheetByName(REQ).appendRow(row);
    const me = Session.getEffectiveUser().getEmail();
    if (me) MailApp.sendEmail(me, 'レッスンのリクエスト：' + String(d.name || '').slice(0, 100) + '（' + (JA.visit[d.visit] || '') + '）',
      String(d.message || '').slice(0, 3000) + '\n\n――\n連絡方法：' + (d.via || '') + '　' + String(d.contact || '').slice(0, 200) + '\nシート：' + book_().getUrl());
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } finally { lock.releaseLock(); }
}

// the website asks for availability: ?availability=1
function doGet(e) {
  const days = {};
  if (e && e.parameter && e.parameter.availability) {
    const av = book_().getSheetByName(AV);
    const today = new Date(); today.setHours(0, 0, 0, 0);
    if (av && av.getLastRow() > 1) av.getRange(2, 1, av.getLastRow() - 1, 5).getValues().forEach(r => {
      if (!(r[0] instanceof Date) || r[0] < today) return;
      if (!r[2] && !r[3] && !r[4]) return;
      days[Utilities.formatDate(r[0], TZ, 'yyyy-MM-dd')] = [String(r[2]), String(r[3]), String(r[4])];
    });
  }
  return ContentService.createTextOutput(JSON.stringify({ days })).setMimeType(ContentService.MimeType.JSON);
}

function clip(v, n) { v = String(v == null ? '' : v); v = v.replace(/^[=+\-@]/, "'$&"); return v.slice(0, n || 200); }
