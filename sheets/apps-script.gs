/**
 * CE Days 2026 Theme Votes — ACES, Aklan State University
 *
 * HOW TO INSTALL (one time, about 3 minutes):
 *  1. In the "CE Days 2026 Theme Votes" sheet, open Extensions > Apps Script.
 *  2. Delete the starter code, paste this whole file, and click Save.
 *  3. In the function dropdown at the top, choose  setup  and click Run.
 *     Approve the Google permissions when asked. This builds the Results tab,
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

/* ---------------- one-time setup: Results tab + deadline trigger ---------------- */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Manila');
  ss.setRecalculationInterval(SpreadsheetApp.RecalculationInterval.MINUTE); // keeps the open/closed status fresh

  // Votes tab
  const v = votesSheet_();
  v.setName(VOTES);
  v.getRange('A1:E1').setValues([['Timestamp', 'Name', 'Year', 'Section', 'Theme']])
    .setFontWeight('bold').setBackground('#0d2547').setFontColor('#ffffff');
  v.setFrozenRows(1);
  v.getRange('A:A').setNumberFormat('mmm d, yyyy h:mm:ss am/pm');
  v.setColumnWidth(1, 190); v.setColumnWidth(2, 220); v.setColumnWidth(5, 210);

  // Results tab (live tally + ranking, all formulas)
  let r = ss.getSheetByName(RESULTS) || ss.insertSheet(RESULTS);
  r.clear(); r.clearConditionalFormatRules();
  const n = THEMES.length, last = 4 + n; // rows 5..15

  r.getRange('A1').setValue('CE Days 2026 · Theme Ranking');
  r.getRange('A2').setFormula(
    '=IF(NOW()<DATE(2026,11,8)+TIME(23,59,59),' +
    '"Voting open · live count · closes Sun, Nov 8, 11:59 PM",' +
    'IF($C$17=0,"Voting closed · no votes received",' +
    'IF(COUNTIF($C$5:$C$15,$C$5)>1,' +
    '"Voting closed · TIE for 1st: "&TEXTJOIN(" / ",TRUE,FILTER($B$5:$B$15,$C$5:$C$15=$C$5))&" ("&$C$5&" votes each)",' +
    '"WINNER: "&$B$5&" · "&$C$5&" votes")))'
  );
  r.getRange('A4:D4').setValues([['Rank', 'Theme', 'Votes', 'Share']]);

  // helper tally (columns H:I) — one COUNTIF per theme so themes with 0 votes still appear
  r.getRange('H4:I4').setValues([['Theme', 'Votes']]);
  r.getRange(5, 8, n, 1).setValues(THEMES.map(t => [t]));
  r.getRange(5, 9, n, 1).setFormulas(THEMES.map((_, i) => ['=COUNTIF(' + VOTES + '!$E:$E,H' + (5 + i) + ')']));

  // ranking (sorted by votes, then A–Z)
  r.getRange('B5').setFormula('=SORT(H5:I' + last + ',2,FALSE,1,TRUE)');
  r.getRange('A5').setFormula('=ARRAYFORMULA(IF(C5:C' + last + '="",,RANK(C5:C' + last + ',C5:C' + last + ',FALSE)))');
  r.getRange('D5').setFormula('=ARRAYFORMULA(IF($C$17=0,0,C5:C' + last + '/$C$17))');
  r.getRange('A17').setValue('Total votes');
  r.getRange('C17').setFormula('=SUM(I5:I' + last + ')');

  // formatting
  r.getRange('A1').setFontSize(16).setFontWeight('bold').setFontColor('#0d2547');
  r.getRange('A2').setFontSize(12).setFontWeight('bold').setFontColor('#b7791f');
  r.getRange('A4:D4').setFontWeight('bold').setBackground('#0d2547').setFontColor('#ffffff');
  r.getRange('H4:I4').setFontWeight('bold').setFontColor('#6b7a90');
  r.getRange('H5:I' + last).setFontColor('#6b7a90');
  r.getRange('D5:D' + last).setNumberFormat('0.0%');
  r.getRange('A5:A' + last).setHorizontalAlignment('center');
  r.getRange('A17:C17').setFontWeight('bold');
  r.setColumnWidth(1, 70); r.setColumnWidth(2, 240); r.setColumnWidth(3, 80); r.setColumnWidth(4, 80);
  r.setColumnWidth(8, 200);
  r.setFrozenRows(4);
  const gold = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A5=1,$C5>0)')
    .setBackground('#fde9b0').setBold(true)
    .setRanges([r.getRange('A5:D' + last)]).build();
  r.setConditionalFormatRules([gold]);

  // schedule the final ranking snapshot at the deadline
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'finalizeRanking')
    .forEach(t => ScriptApp.deleteTrigger(t));
  if (new Date() < DEADLINE) {
    ScriptApp.newTrigger('finalizeRanking').timeBased().at(DEADLINE).create();
  }
  ss.setActiveSheet(r);
}

/* ---------------- runs automatically at the deadline ---------------- */
// Freezes the ranking into a "Final Ranking" tab (plain values, so it never changes afterward).
function finalizeRanking() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.flush();
  const r = ss.getSheetByName(RESULTS);
  if (!r) return;
  const n = THEMES.length, last = 4 + n;
  const rows = r.getRange('A5:D' + last).getValues();
  const total = r.getRange('C17').getValue();
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
