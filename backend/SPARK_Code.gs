/**
 * SPARK — TVS Lanka Horana Skill Evaluation Portal
 * Backend: receives exam submissions from the candidate web page,
 * appends a row to the Results sheet, and saves the face photo to Drive.
 *
 * SETUP (one time):
 * 1. Open the "SPARK Basic Exam Results" Google Sheet.
 * 2. Extensions > Apps Script.
 * 3. Delete any placeholder code, paste this whole file in.
 * 4. Click "Deploy" > "New deployment".
 *    - Select type: "Web app"
 *    - Execute as: Me (denuwantvsl@gmail.com)
 *    - Who has access: Anyone
 * 5. Click Deploy, authorize the permissions it asks for.
 * 6. Copy the resulting Web App URL — it ends in /exec.
 * 7. Paste that URL into spark_exam.html, replacing
 *    "PASTE_YOUR_DEPLOYED_WEB_APP_URL_HERE" in the APPS_SCRIPT_URL constant.
 *
 * If you ever edit this code, you must create a NEW deployment version
 * (Deploy > Manage deployments > pencil icon > New version) for changes
 * to take effect on the live URL.
 */

const SHEET_ID = '1JYnywXmYnN-n89A9nuEU2i_Ub6-sh-7o-bnt5NXRAVA';
const PHOTOS_FOLDER_ID = '1abYuGGJnBI0Y5AFJSaEwPZsB1SZfTfrl';
const SHEET_NAME = 'Results';

// Fixed per-tool question labels and their correct answers, in the same
// order as QUESTIONS in spark_exam.html. Because the exam shuffles both
// question order and option order per candidate, these columns are what
// make every candidate's "Q1" mean the same physical tool — without this,
// answers land in different columns for every candidate and can't be
// audited or color-coded.
const QUESTION_LABELS = [
  'Q1: Open end spanner', 'Q2: Adjustable wrench', 'Q3: Combination wrench',
  'Q4: Hammer', 'Q5: Allen key set', 'Q6: T-handle', 'Q7: Long nose pliers',
  'Q8: Impact driver', 'Q9: Multimeter', 'Q10: Screwdriver (long)',
  'Q11: Flat screwdriver', 'Q12: Heat gun', 'Q13: Electric drill',
  'Q14: Spark plug', 'Q15: Bench vice', 'Q16: Vernier caliper',
  'Q17: Micrometer', 'Q18: Torque wrench', 'Q19: Brand logo', 'Q20: Tool box'
];
const CORRECT_ANSWERS = [
  'Open end spanner', 'Adjuster spanner', 'Wrench spanner', 'Hammer', 'Alan key',
  'T Handle', 'Long nose pliers', 'DC Tool', 'Multimeter', 'Screw driver',
  'Flat Screw Driver', 'Heat blower gun', 'Electric drill', 'Spark plug', 'Bench wise',
  'Vernier caliper', 'Micrometer', 'Torque wrench', 'TVS trade mark', 'Tool Box'
];
const FIRST_QUESTION_COL = 12; // column L — where Q1's answer lands

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

    if (sheet.getLastRow() === 0) {
      const headers = [
        'Timestamp', 'Name', 'Phone', 'Employee Type', 'ID Number',
        'Score', 'Total Questions', 'Percentage', 'Category',
        'Time Taken (sec)', 'Photo Link'
      ];
      headers.push(...QUESTION_LABELS);
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
    }

    // Save photo to Drive — entirely best-effort. Whatever goes wrong
    // here (permissions, quota, anything), it must never stop the exam
    // result itself from being recorded. If it fails, photoUrl just
    // stays empty and everything else still saves normally.
    let photoUrl = '';
    if (data.photoBase64) {
      try {
        const folder = DriveApp.getFolderById(PHOTOS_FOLDER_ID);
        const safeName = String(data.name || 'candidate').replace(/[^a-zA-Z0-9]/g, '_');
        const fileName = safeName + '_' + (data.idNumber || 'noid') + '_' + new Date().getTime() + '.jpg';
        const base64 = data.photoBase64.split(',').pop();
        const bytes = Utilities.base64Decode(base64);
        const blob = Utilities.newBlob(bytes, 'image/jpeg', fileName);
        const file = folder.createFile(blob);
        photoUrl = file.getUrl();
        try {
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        } catch (shareErr) {
          console.error('Photo sharing step failed (non-fatal): ' + shareErr.toString());
        }
      } catch (photoErr) {
        console.error('Photo save failed (non-fatal): ' + photoErr.toString());
        photoUrl = 'PHOTO_SAVE_FAILED: ' + photoErr.toString();
      }
    }

    const total = data.total || 20;
    const score = data.score || 0;
    const pct = Math.round((score / total) * 100);
    let category;
    if (pct >= 90) category = 'Advanced';
    else if (pct >= 80) category = 'Intermediate';
    else if (pct >= 70) category = 'Basic';
    else category = 'Not Qualified';

    const row = [
      new Date(),
      data.name || '',
      data.phone || '',
      data.employeeType || '',
      data.idNumber || '',
      score,
      total,
      pct + '%',
      category,
      data.timeTaken || '',
      photoUrl
    ];
    const answers = data.answers || [];
    for (let i = 0; i < 20; i++) row.push(answers[i] || '');

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', percentage: pct, category: category }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    console.error('SPARK doPost failed: ' + err.toString() + ' | stack: ' + (err.stack || 'n/a'));
    // Best-effort: write the error directly into the sheet so it's visible
    // even if the Executions/Cloud Logging panels don't show it.
    try {
      const ss2 = SpreadsheetApp.openById(SHEET_ID);
      const errSheet = ss2.getSheetByName('Errors') || ss2.insertSheet('Errors');
      if (errSheet.getLastRow() === 0) errSheet.appendRow(['Timestamp', 'Error', 'Stack']);
      errSheet.appendRow([new Date(), err.toString(), err.stack || '']);
    } catch (err2) {
      // if even this fails, there's nothing more we can do here
    }
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * ONE-TIME SETUP — run this once manually (select it in the function
 * dropdown next to the Run button, then click Run) to add the Phone
 * column to a sheet that already has the old header row without it.
 * Safe to run more than once; it checks before inserting.
 */
function setupPhoneColumn() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { console.log('Results sheet not found yet — nothing to fix.'); return; }
  const c1 = sheet.getRange('C1').getValue();
  if (c1 !== 'Phone') {
    sheet.insertColumnBefore(3);
    sheet.getRange('C1').setValue('Phone');
    console.log('Inserted Phone column at C.');
  } else {
    console.log('Phone column already present — nothing to do.');
  }
}

/**
 * ONE-TIME AUTH TRIGGER — run this once manually (select it in the
 * function dropdown, click Run) purely to force the Drive permission
 * consent screen. doPost runs as a Web App and can't prompt for
 * authorization itself, so this function exists just to surface it.
 * Approve ALL permissions it asks for, especially Drive.
 */
function testDriveAccess() {
  const folder = DriveApp.getFolderById(PHOTOS_FOLDER_ID);
  console.log('Drive access OK — folder name: ' + folder.getName());
}

/**
 * ONE-TIME MIGRATION — run this once if your Results sheet already has
 * generic "Q1 Answer".."Q20 Answer" headers from before this update.
 * Renames L1:AE1 to the fixed per-tool labels. Safe to run more than
 * once — it just overwrites those 20 header cells either way.
 * NOTE: any rows saved before this update have answers in shuffled
 * (not per-tool) order and won't line up with these labels — delete
 * those old test rows manually first.
 */
function relabelQuestionColumns() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { console.log('Results sheet not found yet.'); return; }
  sheet.getRange(1, FIRST_QUESTION_COL, 1, QUESTION_LABELS.length).setValues([QUESTION_LABELS]);
  console.log('Relabeled ' + QUESTION_LABELS.length + ' question columns.');
}

/**
 * ONE-TIME SETUP — run this once (select it in the function dropdown,
 * click Run) to color-code every candidate's answers: green background
 * if they matched the correct answer for that tool, red if not. Applies
 * to every row from row 2 down to row 1000, so it keeps working as new
 * submissions come in — no need to re-run this for each new candidate.
 * Safe to run again later (e.g. after extending the row range).
 */
function setupAnswerColorCoding() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { console.log('Results sheet not found yet.'); return; }

  const lastRow = 1000; // covers well beyond 150 candidates; re-run if you ever exceed this
  const existingRules = sheet.getConditionalFormatRules();
  const keepRules = existingRules.filter(r => {
    const ranges = r.getRanges();
    return !ranges.some(rg => rg.getColumn() >= FIRST_QUESTION_COL &&
                               rg.getColumn() < FIRST_QUESTION_COL + QUESTION_LABELS.length);
  });

  const newRules = [];
  for (let i = 0; i < QUESTION_LABELS.length; i++) {
    const col = FIRST_QUESTION_COL + i;
    const colLetter = columnToLetter(col);
    const range = sheet.getRange(2, col, lastRow - 1, 1);
    const correctText = CORRECT_ANSWERS[i].replace(/"/g, '""');

    newRules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=' + colLetter + '2="' + correctText + '"')
        .setBackground('#d9ead3') // light green
        .setRanges([range])
        .build()
    );
    newRules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND(' + colLetter + '2<>"",' + colLetter + '2<>"' + correctText + '")')
        .setBackground('#f4cccc') // light red
        .setRanges([range])
        .build()
    );
  }

  sheet.setConditionalFormatRules(keepRules.concat(newRules));
  console.log('Color coding applied to ' + QUESTION_LABELS.length + ' question columns.');
}

function columnToLetter(col) {
  let letter = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'SPARK backend is live' }))
    .setMimeType(ContentService.MimeType.JSON);
}
