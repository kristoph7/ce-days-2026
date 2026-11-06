/**
 * CE Days 2026 Theme Votes — ACES, Aklan State University
 *
 * HOW TO INSTALL (one time, about 3 minutes):
 *  1. In the "CE Days 2026 Theme Votes" sheet, open Extensions > Apps Script.
 *  2. Delete the starter code, paste this whole file, and click Save.
 *  3. In the function dropdown at the top, choose  setup  and click Run.
 *     Approve the Google permissions when asked. This builds the live Results dashboard,
 *     sets the timezone to Manila, and schedules the final ranking for the deadline.
 *  4. Click Deploy > New deployment > type "Web app".
 *     Execute as: Me    |    Who has access: Anyone    →  Deploy.
 *  5. Copy the Web app URL and send it to Claude.
 */

const DEADLINE = new Date('2026-11-08T23:59:59+08:00'); // Sun, Nov 8, 11:59 PM (PH time)
const THEMES = [
  'Minecraft', 'Street Wear', 'Halloween', 'Wrong Party Attended', 'After Dark', 'Historic',
  'Celestial Midnight Soiree', 'God & Goddess', 'Celestial Seas', 'Barrio Festival', 'Glitz & Glam'
];
const VOTES = 'Votes', RESULTS = 'Results', FINAL = 'Final Ranking';

/* ---------------- receiving votes from the website ---------------- */
function doPost(e) {
  // the website asks "is this student still on the Votes list?" so a deleted row unlocks them
  try {
    const q = JSON.parse(e.postData.contents);
    if (q.action === 'check') return json_({ ok: true, voted: hasVoted_(q.name) });
  } catch (_) {}
  if (new Date() > DEADLINE) return json_({ ok: false, error: 'Voting is closed.' });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const theme = String(d.theme || '');
    if (THEMES.indexOf(theme) === -1) return json_({ ok: false, error: 'Unknown theme.' });
    const sheet = votesSheet_();
    // one vote per student: the same full name (ignoring case, spaces, dots and dashes) is rejected,
    // even with a different year or section
    if (hasVoted_(d.name)) {
      return json_({ ok: false, duplicate: true, error: 'Already voted.' });
    }
    sheet.appendRow([
      new Date(),
      clean_(d.name, 80),
      clean_(d.year, 20),
      clean_(d.section, 30),
      theme
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  const open = new Date() <= DEADLINE;
  return ContentService.createTextOutput('CE Days 2026 vote endpoint · ' + (open ? 'voting open' : 'voting closed'));
}

/* ---------------- one-time setup: live Results dashboard + deadline trigger ---------------- */
// Results layout
//   rows 1-2   title + status (WINNER / TIE once voting closes)
//   rows 4-5   live counters: total votes · leading theme · last vote received
//   rows 7-18  ranking table, re-sorted automatically by votes (most first)
//   rows 20-24 votes by year level
//   H7:I18     helper tally (one COUNTIF per theme so 0-vote themes still show)
const R = { total: '$A$5', first: 8, last: 18 };

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Manila');
  ss.setRecalculationInterval(SpreadsheetApp.RecalculationInterval.MINUTE); // keeps NOW()-based status fresh

  // Votes tab
  const v = votesSheet_();
  v.setName(VOTES);
  v.getRange('A1:E1').setValues([['Timestamp', 'Name', 'Year', 'Section', 'Theme']])
    .setFontWeight('bold').setBackground('#0d2547').setFontColor('#ffffff');
  v.setFrozenRows(1);
  v.getRange('A:A').setNumberFormat('mmm d, yyyy h:mm:ss am/pm');
  v.setColumnWidth(1, 190); v.setColumnWidth(2, 220); v.setColumnWidth(3, 90); v.setColumnWidth(4, 110); v.setColumnWidth(5, 210);

  // Results tab
  let r = ss.getSheetByName(RESULTS) || ss.insertSheet(RESULTS);
  r.getRange('A1:L40').breakApart();
  r.clear(); r.clearConditionalFormatRules();
  const f = R.first, l = R.last, n = THEMES.length;
  const C = 'C' + f + ':C' + l, B = 'B' + f + ':B' + l;

  // title + status
  r.getRange('A1:F1').merge(); r.getRange('A1').setValue('CE Days 2026 · Live Theme Ranking');
  r.getRange('A2:F2').merge(); r.getRange('A2').setFormula(
    '=IF(NOW()<DATE(2026,11,8)+TIME(23,59,59),' +
    '"🟢 Voting open · updates live with every vote · closes Sun, Nov 8, 11:59 PM",' +
    'IF(' + R.total + '=0,"Voting closed · no votes received",' +
    'IF(COUNTIF(' + C + ',$C$' + f + ')>1,' +
    '"🏁 Voting closed · TIE for 1st: "&TEXTJOIN(" / ",TRUE,FILTER(' + B + ',' + C + '=$C$' + f + '))&" ("&$C$' + f + '&" votes each)",' +
    '"🏆 WINNER: "&$B$' + f + '&" · "&$C$' + f + '&" votes")))'
  );

  // live counters
  r.getRange('A4:B4').merge().setValue('TOTAL VOTES');
  r.getRange('C4:D4').merge().setValue('LEADING THEME');
  r.getRange('E4:F4').merge().setValue('LAST VOTE RECEIVED');
  r.getRange('A5:B5').merge(); r.getRange('A5').setFormula('=SUM(I' + f + ':I' + l + ')');
  r.getRange('C5:D5').merge(); r.getRange('C5').setFormula(
    '=IF(' + R.total + '=0,"No votes yet",IF(COUNTIF(' + C + ',$C$' + f + ')>1,' +
    '"Tie: "&TEXTJOIN(" / ",TRUE,FILTER(' + B + ',' + C + '=$C$' + f + ')),$B$' + f + '&" ("&$C$' + f + '&")"))');
  r.getRange('E5:F5').merge(); r.getRange('E5').setFormula(
    '=IF(COUNT(' + VOTES + '!A2:A)=0,"—",TEXT(MAX(' + VOTES + '!A2:A),"mmm d · h:mm:ss AM/PM"))');

  // ranking table (auto-sorted)
  r.getRange('A7:E7').setValues([['Rank', 'Theme', 'Votes', 'Share', 'Bar']]);
  r.getRange('H7:I7').setValues([['Theme', 'Votes']]);
  r.getRange(f, 8, n, 1).setValues(THEMES.map(t => [t]));
  r.getRange(f, 9, n, 1).setFormulas(THEMES.map((_, i) => ['=COUNTIF(' + VOTES + '!$E:$E,H' + (f + i) + ')']));
  r.getRange('B' + f).setFormula('=SORT(H' + f + ':I' + l + ',2,FALSE,1,TRUE)');
  r.getRange('A' + f).setFormula('=ARRAYFORMULA(IF(' + C + '="",,RANK(' + C + ',' + C + ',FALSE)))');
  r.getRange('D' + f).setFormula('=ARRAYFORMULA(IF(' + C + '="",,IF(' + R.total + '=0,0,' + C + '/' + R.total + ')))');
  r.getRange('E' + f).setFormula('=ARRAYFORMULA(IF(' + C + '="",,IF(' + R.total + '=0,"",REPT("█",ROUND(' + C + '/MAX(' + C + ')*20)))))');

  // votes by year
  r.getRange('A20:C20').setValues([['Year level', '', 'Votes']]);
  const years = ['1st Year', '2nd Year', '3rd Year', '4th Year'];
  r.getRange('A21:A24').setValues(years.map(y => [y]));
  r.getRange('C21:C24').setFormulas(years.map((_, i) => ['=COUNTIF(' + VOTES + '!$C:$C,A' + (21 + i) + ')']));

  // formatting
  const navy = '#0d2547', gold = '#b7791f', soft = '#6b7a90';
  r.getRange('A1').setFontSize(18).setFontWeight('bold').setFontColor(navy);
  r.getRange('A2').setFontSize(12).setFontWeight('bold').setFontColor(gold);
  r.getRange('A4:F4').setFontSize(9).setFontWeight('bold').setFontColor(soft).setHorizontalAlignment('center');
  r.getRange('A5:F5').setBackground('#eef3fb').setHorizontalAlignment('center').setVerticalAlignment('middle').setFontWeight('bold').setFontColor(navy);
  r.getRange('A5').setFontSize(30);
  r.getRange('C5').setFontSize(14).setWrap(true);
  r.getRange('E5').setFontSize(12);
  r.setRowHeight(5, 58);
  r.getRange('A7:E7').setFontWeight('bold').setBackground(navy).setFontColor('#ffffff');
  r.getRange('H7:I7').setFontWeight('bold').setFontColor(soft);
  r.getRange('H' + f + ':I' + l).setFontColor(soft);
  r.getRange('A' + f + ':A' + l).setHorizontalAlignment('center');
  r.getRange('C' + f + ':C' + l).setHorizontalAlignment('center').setFontWeight('bold');
  r.getRange('D' + f + ':D' + l).setNumberFormat('0.0%').setHorizontalAlignment('center');
  r.getRange('E' + f + ':E' + l).setFontColor('#e0a526');
  r.getRange('A20:C20').setFontWeight('bold').setBackground(navy).setFontColor('#ffffff');
  r.getRange('C21:C24').setHorizontalAlignment('center').setFontWeight('bold');
  r.setColumnWidth(1, 70); r.setColumnWidth(2, 230); r.setColumnWidth(3, 80); r.setColumnWidth(4, 80);
  r.setColumnWidth(5, 190); r.setColumnWidth(6, 120); r.setColumnWidth(7, 30); r.setColumnWidth(8, 200);
  r.setFrozenRows(2);
  const leader = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A' + f + '=1,$C' + f + '>0)')
    .setBackground('#fde9b0').setBold(true)
    .setRanges([r.getRange('A' + f + ':E' + l)]).build();
  r.setConditionalFormatRules([leader]);

  // Results first, so it opens on the dashboard
  ss.setActiveSheet(r); ss.moveActiveSheet(1);

  // schedule the final ranking snapshot at the deadline
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'finalizeRanking')
    .forEach(t => ScriptApp.deleteTrigger(t));
  if (new Date() < DEADLINE) {
    ScriptApp.newTrigger('finalizeRanking').timeBased().at(DEADLINE).create();
  }
}

/* ---------------- runs automatically at the deadline ---------------- */
// Freezes the ranking into a "Final Ranking" tab (plain values, so it never changes afterward).
function finalizeRanking() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.flush();
  const r = ss.getSheetByName(RESULTS);
  if (!r) return;
  const rows = r.getRange('A' + R.first + ':D' + R.last).getValues();
  const total = r.getRange('A5').getValue();
  const status = r.getRange('A2').getDisplayValue();

  let f = ss.getSheetByName(FINAL) || ss.insertSheet(FINAL);
  f.clear(); f.clearConditionalFormatRules();
  f.getRange('A1').setValue('CE Days 2026 · Final Theme Ranking').setFontSize(16).setFontWeight('bold').setFontColor('#0d2547');
  f.getRange('A2').setValue(status).setFontSize(13).setFontWeight('bold').setFontColor('#b7791f');
  f.getRange('A3').setValue('Locked ' + Utilities.formatDate(new Date(), 'Asia/Manila', 'MMM d, yyyy h:mm a') + ' (Philippine time)').setFontColor('#6b7a90');
  f.getRange('A5:D5').setValues([['Rank', 'Theme', 'Votes', 'Share']]).setFontWeight('bold').setBackground('#0d2547').setFontColor('#ffffff');
  f.getRange(6, 1, rows.length, 4).setValues(rows);
  f.getRange(6, 4, rows.length, 1).setNumberFormat('0.0%');
  f.getRange(6, 1, rows.length, 1).setHorizontalAlignment('center');
  f.getRange(7 + rows.length, 1, 1, 3).setValues([['Total votes', '', total]]).setFontWeight('bold');
  rows.forEach((row, i) => {
    if (row[0] === 1 && row[2] > 0) f.getRange(6 + i, 1, 1, 4).setBackground('#fde9b0').setFontWeight('bold');
  });
  f.setColumnWidth(1, 70); f.setColumnWidth(2, 240); f.setColumnWidth(3, 80); f.setColumnWidth(4, 80);
  ss.setActiveSheet(f);
  ss.moveActiveSheet(1);
}

/* ---------------- helpers ---------------- */
function votesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(VOTES) || ss.getSheets()[0];
}
function clean_(v, max) {
  return String(v || '').replace(/^[=+\-@]/, "'$&").slice(0, max); // stops formula injection
}
function hasVoted_(name) {
  const sheet = votesSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return false;
  const key = norm_(name);
  if (!key) return false;
  const names = sheet.getRange(2, 2, last - 1, 1).getValues(); // Name column
  return names.some(r => norm_(r[0]) === key);
}
function norm_(v) {
  return String(v || '').replace(/^'/, '').toLowerCase().replace(/[\s\-_.,]+/g, '');
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
