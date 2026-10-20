/**
 * SPARK — TVS Lanka Horana Skill Evaluation Portal
 * Backend for three pages:
 *  - index.html  (Basic Entry Exam): appends a row to the "Results" tab
 *    and saves the face photo to Drive.
 *  - assess.html (SKD Assembly Skill Assessment, assessor use): appends a
 *    row to "SKD Assessments" and updates the "SKD Skill Matrix" tab.
 *  - matrix.html (Employee Skill Matrix, PIN-protected): reads everyone's
 *    basic exam + SKD results.
 *
 * The web app is already deployed (Execute as: Me, Who has access: Anyone).
 * After ANY edit: Deploy > Manage deployments > pencil icon >
 * Version: New version > Deploy. Never use "New deployment" — that
 * creates a new /exec URL and the live pages would stop saving.
 *
 * The SKD page needs a Script Property ASSESSOR_PIN
 * (Project Settings > Script Properties). Without it, every SKD request
 * is refused.
 */

const SHEET_ID = '1JYnywXmYnN-n89A9nuEU2i_Ub6-sh-7o-bnt5NXRAVA';
const PHOTOS_FOLDER_ID = '1abYuGGJnBI0Y5AFJSaEwPZsB1SZfTfrl';
const SHEET_NAME = 'Results';

// Fixed per-tool question labels and their correct answers, in the same
// order as QUESTIONS in index.html. Because the exam shuffles both
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
    if (data.formType === 'skd') return handleSkd(data);

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

/* =====================================================================
 * SKD ASSEMBLY SKILL ASSESSMENT (assess.html) + EMPLOYEE SKILL MATRIX (matrix.html)
 * Based on "SKD Assembly.xlsx" (Galhena): 19 key areas x 8 models.
 * Each area has 5 criteria marked Pass/Fail by the assessor. An area is
 * SKILLED (1) only if all 5 pass, otherwise 0, or N/A if it doesn't
 * apply to the model. Model % = skilled / applicable areas.
 * All-models % = average of the 8 model % (not-assessed models count 0),
 * the same as "Employee working knowledge % for all models" in the Excel.
 * Levels use the basic exam bands: 90+ Advanced, 80-89 Intermediate,
 * 70-79 Basic, below 70 Not Qualified.
 * Model and area order MUST match MODELS / AREAS in assess.html and matrix.html.
 * Every SKD action needs the Script Property ASSESSOR_PIN.
 * Actions: checkPin, lookup (one employee), submit, matrix (everyone).
 * ===================================================================== */
const SKD_ASSESS_SHEET = 'SKD Assessments';
const SKD_MATRIX_SHEET = 'SKD Skill Matrix';
const SKD_MODELS = ['Sport 110', 'Ntorq 125', 'Ronin 225', 'RTR', 'iQube', 'Orbiter', 'Jupiter', 'Raider 125'];
const SKD_AREAS = [
  ['Depacking Process', ['Safety compliance', 'Correct part identification', 'Proper handling / damage prevention', 'Work method & sequence', 'Housekeeping']],
  ['Trolly Loading Process', ['Correct loading sequence', 'Load stability', 'Safe handling', 'Damage prevention', 'Workplace housekeeping']],
  ['Hoister Usage', ['Pre-use safety check', 'Correct lifting point', 'Safe operation', 'Load balance / stability', 'Post-use handling']],
  ['Handle Bar Assembly', ['Correct part & orientation', 'Proper fitment', 'Torque / tightening', 'Alignment', 'Final quality check']],
  ['Speedometer Assembly', ['Correct part identification', 'Correct fitment', 'Wiring / connector condition', 'Function check', 'Workmanship']],
  ['HSG Head Lamp Assembly', ['Correct fitment', 'Wiring / connector condition', 'Alignment', 'Fastener tightening', 'Function check']],
  ['TSL FR L & R Assembly', ['Correct left/right parts', 'Correct fitment', 'Wiring / connector condition', 'Alignment', 'Final inspection']],
  ['Front Fender Assembly', ['Correct fitment', 'Alignment', 'Fastener tightening', 'Damage prevention', 'Final inspection']],
  ['Wheel Assembly – Front', ['Correct wheel / component identification', 'Axle & fitment', 'Torque / tightening', 'Alignment', 'Brake / rotation check']],
  ['Muffler Assembly', ['Correct fitment', 'Fastener torque', 'Clearance', 'Leakage / abnormal noise check', 'Final inspection']],
  ['Crash Guard Assembly', ['Correct position', 'Fitment', 'Torque / tightening', 'Clearance', 'Final inspection']],
  ['Main Footrest Assembly', ['Correct fitment', 'Torque / tightening', 'Movement / function', 'Alignment', 'Final inspection']],
  ['Footrest L & R Assembly', ['Correct left/right parts', 'Correct fitment', 'Torque / tightening', 'Movement / function', 'Final inspection']],
  ['Unloading', ['Safe handling', 'Correct unloading sequence', 'Damage prevention', 'Material identification', 'Housekeeping']],
  ['Testing', ['Correct test procedure', 'Safety compliance', 'Defect identification', 'Accurate result recording', 'Final confirmation']],
  ['Minor Repair', ['Correct diagnosis', 'Correct repair method', 'Proper tool usage', 'Quality after repair', 'Documentation / confirmation']],
  ['Major Repair', ['Correct diagnosis', 'Correct repair method', 'Safety compliance', 'Quality after repair', 'Final verification']],
  ['Identified Damage Parts & Shortages', ['Correct identification', 'Accurate quantity', 'Proper segregation', 'Timely reporting', 'Documentation']],
  ['Quality of Work', ['First-time-right', 'Workmanship', 'Defect prevention', 'Standard compliance', 'Final quality awareness']]
];
// SKD Assessments: A Timestamp .. J Remarks, K..AC areas 1-19, AD Failed Criteria, AE Phone, AF Photo Link
const SKD_FIRST_AREA_COL = 11;                                   // K
const SKD_ASSESS_PHONE_COL = SKD_FIRST_AREA_COL + SKD_AREAS.length + 1; // AE (31)
const SKD_ASSESS_PHOTO_COL = SKD_ASSESS_PHONE_COL + 1;            // AF (32)
// SKD Skill Matrix: A Name, B Type, C ID, D..K models, L All %, M Models Assessed, N Last Assessed,
// O Last Assessor, P Phone, Q Photo Link
const SKD_MATRIX_FIRST_MODEL_COL = 4;                             // D
const SKD_MATRIX_WIDTH = 3 + SKD_MODELS.length + 6;               // A..Q (17)
const TZ = 'Asia/Colombo';

function skdReply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function normId(v) { return String(v || '').trim().toUpperCase(); }

function fmtDate(d) {
  return (d instanceof Date) ? Utilities.formatDate(d, TZ, 'dd MMM yyyy') : String(d || '');
}

// Drive file URL -> image URL a page can show (file must be shared "anyone with link").
function photoThumb(url) {
  const m = String(url || '').match(/\/d\/([a-zA-Z0-9_-]{20,})/) || String(url || '').match(/[?&]id=([a-zA-Z0-9_-]{20,})/);
  return m ? 'https://drive.google.com/thumbnail?id=' + m[1] + '&sz=w400' : '';
}

// Results tab stores "95%" (Sheets may turn it into 0.95).
function pctNumber(v) {
  if (typeof v === 'number') return v <= 1 ? Math.round(v * 100) : Math.round(v);
  const n = parseFloat(String(v).replace('%', ''));
  return isNaN(n) ? null : Math.round(n);
}

function handleSkd(data) {
  const expectedPin = PropertiesService.getScriptProperties().getProperty('ASSESSOR_PIN');
  if (!expectedPin) {
    return skdReply({ status: 'error', message: 'Assessor PIN is not set up yet (Apps Script > Project Settings > Script Properties > ASSESSOR_PIN).' });
  }
  if (String(data.pin || '').trim() !== String(expectedPin).trim()) {
    return skdReply({ status: 'error', message: 'Wrong assessor PIN.' });
  }
  if (data.action === 'checkPin') return skdReply({ status: 'success' });
  if (data.action === 'lookup') return skdLookup(normId(data.idNumber));
  if (data.action === 'matrix') return skdReply({ status: 'success', employees: skdAllEmployees() });
  if (data.action !== 'submit') return skdReply({ status: 'error', message: 'Unknown action.' });

  // ---- validate ----
  const model = String(data.model || '');
  if (SKD_MODELS.indexOf(model) === -1) return skdReply({ status: 'error', message: 'Unknown model: ' + model });
  const idNumber = normId(data.idNumber);
  const name = String(data.name || '').trim();
  const phone = String(data.phone || '').trim();
  if (!idNumber || !name) return skdReply({ status: 'error', message: 'Employee name and ID are required.' });
  const areas = data.areas;
  if (!Array.isArray(areas) || areas.length !== SKD_AREAS.length) {
    return skdReply({ status: 'error', message: 'Expected ' + SKD_AREAS.length + ' areas, got ' + (areas ? areas.length : 0) + '.' });
  }

  // ---- score on the server (don't trust the page's own %) ----
  const areaResults = [];
  const failedText = [];
  let applicable = 0, skilled = 0;
  for (let a = 0; a < SKD_AREAS.length; a++) {
    const m = areas[a] || {};
    if (m.na === true) { areaResults.push('N/A'); continue; }
    const crit = Array.isArray(m.crit) ? m.crit : [];
    const critNames = SKD_AREAS[a][1];
    if (crit.length !== critNames.length || crit.some(v => v !== 0 && v !== 1)) {
      return skdReply({ status: 'error', message: 'Area ' + (a + 1) + ' (' + SKD_AREAS[a][0] + ') is not fully marked.' });
    }
    applicable++;
    const failed = critNames.filter((_, i) => crit[i] !== 1);
    if (failed.length === 0) { skilled++; areaResults.push(1); }
    else { areaResults.push(0); failedText.push((a + 1) + '. ' + SKD_AREAS[a][0] + ': ' + failed.join(', ')); }
  }
  if (applicable === 0) return skdReply({ status: 'error', message: 'All areas are N/A. At least one area must be assessed.' });
  const modelPct = Math.round(skilled / applicable * 100);
  const empType = data.employeeType === 'company' ? 'Company' : 'Manpower';
  const now = new Date();

  // ---- photo: best-effort, never blocks the result ----
  let newPhotoUrl = '';
  if (data.photoBase64) newPhotoUrl = skdSavePhoto(data.photoBase64, name, idNumber);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);

    // ---- 2. find this employee's matrix row (needed for the photo fallback) ----
    const mSheet = skdMatrixSheet(ss);
    const nModels = SKD_MODELS.length;
    const last = mSheet.getLastRow();
    let r = -1;
    if (last >= 2) {
      const ids = mSheet.getRange(2, 3, last - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (normId(ids[i][0]) === idNumber) { r = i + 2; break; }
      }
    }
    let mRow;
    if (r === -1) { r = last + 1; mRow = new Array(SKD_MATRIX_WIDTH).fill(''); }
    else { mRow = mSheet.getRange(r, 1, 1, SKD_MATRIX_WIDTH).getValues()[0]; }

    // Photo for this employee: new one, else the one already on file, else the basic exam selfie.
    let photoUrl = newPhotoUrl;
    if (!photoUrl || photoUrl.indexOf('PHOTO_SAVE_FAILED') === 0) {
      const kept = String(mRow[SKD_MATRIX_WIDTH - 1] || '');
      photoUrl = kept || (skdBasicResult(ss, idNumber) || {}).photoUrl || newPhotoUrl || '';
    }

    // ---- 1. full record of this assessment ----
    const aSheet = skdAssessSheet(ss);
    const row = [now, String(data.assessor || ''), name, empType, idNumber, model,
                 applicable, skilled, modelPct / 100, String(data.remarks || '')];
    row.push(...areaResults);
    row.push(failedText.join('\n'));
    row.push(phone ? "'" + phone : '', newPhotoUrl);
    aSheet.appendRow(row);

    // ---- 2. latest result per employee x model ----
    mRow[0] = name;
    mRow[1] = empType;
    mRow[2] = idNumber;
    mRow[SKD_MATRIX_FIRST_MODEL_COL - 1 + SKD_MODELS.indexOf(model)] = modelPct / 100;
    let sum = 0, assessed = 0;
    const matrix = {};
    for (let k = 0; k < nModels; k++) {
      const v = mRow[SKD_MATRIX_FIRST_MODEL_COL - 1 + k];
      if (v === '' || v === null) { matrix[SKD_MODELS[k]] = null; continue; }
      sum += Number(v); assessed++;
      matrix[SKD_MODELS[k]] = Math.round(Number(v) * 100);
    }
    const allPct = sum / nModels;
    mRow[3 + nModels] = allPct;
    mRow[3 + nModels + 1] = assessed;
    mRow[3 + nModels + 2] = now;
    mRow[3 + nModels + 3] = String(data.assessor || '');
    // Leading ' keeps the phone as text, so 07... doesn't lose its 0 when written back.
    const keepPhone = (phone || String(mRow[3 + nModels + 4] || '')).replace(/^'+/, '');
    mRow[3 + nModels + 4] = keepPhone ? "'" + keepPhone : '';
    mRow[3 + nModels + 5] = photoUrl;
    mSheet.getRange(r, 1, 1, SKD_MATRIX_WIDTH).setValues([mRow]);

    return skdReply({
      status: 'success', model: model, applicable: applicable, skilled: skilled,
      modelPct: modelPct, allModelsPct: Math.round(allPct * 100), matrix: matrix,
      photo: photoThumb(photoUrl)
    });
  } finally {
    lock.releaseLock();
  }
}

function skdSavePhoto(photoBase64, name, idNumber) {
  try {
    const folder = DriveApp.getFolderById(PHOTOS_FOLDER_ID);
    const safeName = String(name || 'employee').replace(/[^a-zA-Z0-9]/g, '_');
    const fileName = 'SKD_' + safeName + '_' + idNumber + '_' + new Date().getTime() + '.jpg';
    const bytes = Utilities.base64Decode(String(photoBase64).split(',').pop());
    const file = folder.createFile(Utilities.newBlob(bytes, 'image/jpeg', fileName));
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      console.error('SKD photo sharing failed (non-fatal): ' + shareErr.toString());
    }
    return file.getUrl();
  } catch (photoErr) {
    console.error('SKD photo save failed (non-fatal): ' + photoErr.toString());
    return 'PHOTO_SAVE_FAILED: ' + photoErr.toString();
  }
}

// Latest basic exam (Results tab) row for this ID, or null.
function skdBasicResult(ss, idNumber) {
  const all = skdBasicResultsById(ss);
  return all[idNumber] || null;
}

function skdBasicResultsById(ss) {
  const out = {};
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) return out;
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 11).getValues();
  vals.forEach(v => {
    const id = normId(v[4]);
    if (!id || !String(v[1] || '').trim()) return;   // skip blank / junk rows
    const photo = String(v[10] || '');
    out[id] = {                                       // later rows overwrite = latest attempt wins
      name: String(v[1]), phone: String(v[2] || ''),
      empType: String(v[3] || '').toLowerCase() === 'company' ? 'Company' : 'Manpower',
      score: v[5], total: v[6], pct: pctNumber(v[7]), category: String(v[8] || ''),
      date: fmtDate(v[0]), photoUrl: photo.indexOf('PHOTO_SAVE_FAILED') === 0 ? '' : photo
    };
  });
  return out;
}

function skdLookup(idNumber) {
  if (!idNumber) return skdReply({ status: 'error', message: 'Enter the ID number first.' });
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const basic = skdBasicResult(ss, idNumber);

  // Latest result + date per model, from the full assessment log.
  const models = {};
  const aSheet = ss.getSheetByName(SKD_ASSESS_SHEET);
  if (aSheet && aSheet.getLastRow() >= 2) {
    const vals = aSheet.getRange(2, 1, aSheet.getLastRow() - 1, 9).getValues();
    vals.forEach(v => {
      if (normId(v[4]) !== idNumber) return;
      models[String(v[5])] = { pct: Math.round(Number(v[8]) * 100), date: fmtDate(v[0]) };
    });
  }

  let emp = null, allPct = null;
  const mSheet = ss.getSheetByName(SKD_MATRIX_SHEET);
  if (mSheet && mSheet.getLastRow() >= 2) {
    const rows = mSheet.getRange(2, 1, mSheet.getLastRow() - 1, SKD_MATRIX_WIDTH).getValues();
    const m = rows.find(row => normId(row[2]) === idNumber);
    if (m) {
      emp = { name: String(m[0]), empType: String(m[1]), phone: String(m[SKD_MATRIX_WIDTH - 2] || ''), photoUrl: String(m[SKD_MATRIX_WIDTH - 1] || '') };
      allPct = m[3 + SKD_MODELS.length] === '' ? null : Math.round(Number(m[3 + SKD_MODELS.length]) * 100);
    }
  }
  if (!emp && basic) emp = { name: basic.name, empType: basic.empType, phone: basic.phone, photoUrl: basic.photoUrl };
  if (emp && !emp.photoUrl && basic) emp.photoUrl = basic.photoUrl;
  if (emp && !emp.phone && basic) emp.phone = basic.phone;

  return skdReply({
    status: 'success', found: !!emp,
    employee: emp ? { name: emp.name, empType: emp.empType, phone: emp.phone, photo: photoThumb(emp.photoUrl) } : null,
    models: models, allModelsPct: allPct,
    basic: basic ? { pct: basic.pct, score: basic.score, total: basic.total, category: basic.category, date: basic.date } : null
  });
}

// Everyone who has a basic exam result and/or an SKD assessment.
function skdAllEmployees() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const basicById = skdBasicResultsById(ss);
  const byId = {};
  const mSheet = ss.getSheetByName(SKD_MATRIX_SHEET);
  if (mSheet && mSheet.getLastRow() >= 2) {
    mSheet.getRange(2, 1, mSheet.getLastRow() - 1, SKD_MATRIX_WIDTH).getValues().forEach(m => {
      const id = normId(m[2]);
      if (!id) return;
      const models = {};
      SKD_MODELS.forEach((name, k) => {
        const v = m[SKD_MATRIX_FIRST_MODEL_COL - 1 + k];
        models[name] = (v === '' || v === null) ? null : Math.round(Number(v) * 100);
      });
      const allV = m[3 + SKD_MODELS.length];
      byId[id] = {
        id: id, name: String(m[0]), empType: String(m[1]),
        phone: String(m[SKD_MATRIX_WIDTH - 2] || ''), photoUrl: String(m[SKD_MATRIX_WIDTH - 1] || ''),
        models: models, allPct: allV === '' ? null : Math.round(Number(allV) * 100),
        modelsAssessed: Number(m[3 + SKD_MODELS.length + 1]) || 0,
        lastAssessed: fmtDate(m[3 + SKD_MODELS.length + 2])
      };
    });
  }
  Object.keys(basicById).forEach(id => {
    const b = basicById[id];
    if (!byId[id]) {
      byId[id] = { id: id, name: b.name, empType: b.empType, phone: b.phone, photoUrl: b.photoUrl,
                   models: {}, allPct: null, modelsAssessed: 0, lastAssessed: '' };
    }
    const e = byId[id];
    if (!e.photoUrl) e.photoUrl = b.photoUrl;
    if (!e.phone) e.phone = b.phone;
    e.basic = { pct: b.pct, score: b.score, total: b.total, category: b.category, date: b.date };
  });
  return Object.keys(byId).map(id => {
    const e = byId[id];
    return { id: e.id, name: e.name, empType: e.empType, phone: e.phone, photo: photoThumb(e.photoUrl),
             models: e.models, allPct: e.allPct, modelsAssessed: e.modelsAssessed,
             lastAssessed: e.lastAssessed, basic: e.basic || null };
  });
}

// Writes any missing header cells (used when new columns are added to an existing tab).
function skdEnsureHeaders(sh, headers) {
  const cur = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  headers.forEach((h, i) => { if (cur[i] === '' || cur[i] === null) sh.getRange(1, i + 1).setValue(h).setFontWeight('bold').setWrap(true); });
}

function skdAssessSheet(ss) {
  const headers = ['Timestamp', 'Assessor', 'Employee Name', 'Employee Type', 'ID Number', 'Model',
                   'Applicable Areas', 'Skilled Areas', 'Model Skill %', 'Remarks'];
  SKD_AREAS.forEach((a, i) => headers.push((i + 1) + '. ' + a[0]));
  headers.push('Failed Criteria', 'Phone', 'Photo Link');
  let sh = ss.getSheetByName(SKD_ASSESS_SHEET);
  if (sh && sh.getLastRow() > 0) { skdEnsureHeaders(sh, headers); return sh; }
  if (!sh) sh = ss.insertSheet(SKD_ASSESS_SHEET);
  sh.appendRow(headers);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setWrap(true);
  sh.getRange('I2:I').setNumberFormat('0%');
  // Area columns: 1 = green, 0 = red, N/A = grey
  const areaRange = sh.getRange(2, SKD_FIRST_AREA_COL, 999, SKD_AREAS.length);
  sh.setConditionalFormatRules(sh.getConditionalFormatRules().concat([
    SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(1).setBackground('#d9ead3').setRanges([areaRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(0).setBackground('#f4cccc').setRanges([areaRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('N/A').setBackground('#e0e0e0').setRanges([areaRange]).build()
  ]));
  return sh;
}

function skdMatrixSheet(ss) {
  const headers = ['Employee Name', 'Employee Type', 'ID Number'].concat(SKD_MODELS)
    .concat(['All Models %', 'Models Assessed', 'Last Assessed', 'Last Assessor', 'Phone', 'Photo Link']);
  let sh = ss.getSheetByName(SKD_MATRIX_SHEET);
  if (sh && sh.getLastRow() > 0) { skdEnsureHeaders(sh, headers); return sh; }
  if (!sh) sh = ss.insertSheet(SKD_MATRIX_SHEET);
  sh.appendRow(headers);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setWrap(true);
  const pctCols = sh.getRange(2, SKD_MATRIX_FIRST_MODEL_COL, 999, SKD_MODELS.length + 1);
  pctCols.setNumberFormat('0%');
  // ISNUMBER keeps not-assessed (blank) models uncoloured.
  const c = 'D2';
  sh.setConditionalFormatRules(sh.getConditionalFormatRules().concat([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER(' + c + '),' + c + '>=0.9)').setBackground('#d9ead3').setRanges([pctCols]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER(' + c + '),' + c + '>=0.7,' + c + '<0.9)').setBackground('#fff2cc').setRanges([pctCols]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER(' + c + '),' + c + '<0.7)').setBackground('#f4cccc').setRanges([pctCols]).build()
  ]));
  return sh;
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'SPARK backend is live' }))
    .setMimeType(ContentService.MimeType.JSON);
}
