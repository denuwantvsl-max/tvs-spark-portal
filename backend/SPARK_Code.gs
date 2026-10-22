/**
 * SPARK — TVS Lanka Horana Skill Evaluation Portal
 * Backend for three pages:
 *  - index.html  (Basic Entry Exam): appends a row to the "Results" tab
 *    and saves the face photo to Drive.
 *  - assess.html (skill assessment, all lines, assessor use): appends a
 *    row to "SKD Assessments" or "Workstation Assessments" and rebuilds "SKD Skill Matrix".
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
 * SKILL ASSESSMENT (assess.html) + EMPLOYEE SKILL MATRIX (matrix.html)
 * Lines and models: SKD (12 models), CKD (XL 100 i-Touch), 3W (3W, EV Max 3W).
 * Two evaluation methods (templates, defined in templates.js — copied below):
 *  - 'areas'    : Gala's original SKD Assembly.xlsx, 19 key areas x 5 own criteria.
 *                 Logged in "SKD Assessments".
 *  - 'stations' : Gala's "Test Revised with criteria.xlsx" workstation sheets
 *                 (RTR 160 4V / RTR 200, CKD XL 100, 3W / EV Max 3W). The assessor
 *                 picks the workstation(s); every activity is marked on the same
 *                 5 criteria. Logged in "Workstation Assessments".
 * An item (area/activity) is SKILLED only if all 5 criteria pass; N/A items are
 * left out. Model % = skilled / applicable items.
 * Overall % = average over all models of every line the employee has been
 * assessed in (not-assessed models in those lines count 0 — the SKD Excel rule).
 * Levels = basic exam bands: 90+ Advanced, 80-89 Intermediate, 70-79 Basic, <70 Not Qualified.
 * "SKD Skill Matrix" is REBUILT from the two logs on every save — don't edit it by hand.
 * Every action needs Script Property ASSESSOR_PIN: checkPin, lookup, submit, matrix.
 * ===================================================================== */
// ---- SPARK_TEMPLATES (generated from templates.js — do not edit here) ----
const SPARK = {
 "criteria5": [
  "Assembly Quality & Accuracy",
  "Fastener Tightening & Torque Compliance",
  "Workmanship & Component Handling",
  "Safety & Standard Work Compliance",
  "Inspection & Non-Conformity Management"
 ],
 "templates": {
  "skd19": {
   "kind": "areas",
   "title": "SKD assembly – 19 key areas",
   "source": "SKD Assembly.xlsx",
   "areas": [
    {
     "name": "Depacking Process",
     "crit": [
      "Safety compliance",
      "Correct part identification",
      "Proper handling / damage prevention",
      "Work method & sequence",
      "Housekeeping"
     ]
    },
    {
     "name": "Trolly Loading Process",
     "crit": [
      "Correct loading sequence",
      "Load stability",
      "Safe handling",
      "Damage prevention",
      "Workplace housekeeping"
     ]
    },
    {
     "name": "Hoister Usage",
     "crit": [
      "Pre-use safety check",
      "Correct lifting point",
      "Safe operation",
      "Load balance / stability",
      "Post-use handling"
     ]
    },
    {
     "name": "Handle Bar Assembly",
     "crit": [
      "Correct part & orientation",
      "Proper fitment",
      "Torque / tightening",
      "Alignment",
      "Final quality check"
     ]
    },
    {
     "name": "Speedometer Assembly",
     "crit": [
      "Correct part identification",
      "Correct fitment",
      "Wiring / connector condition",
      "Function check",
      "Workmanship"
     ]
    },
    {
     "name": "HSG Head Lamp Assembly",
     "crit": [
      "Correct fitment",
      "Wiring / connector condition",
      "Alignment",
      "Fastener tightening",
      "Function check"
     ]
    },
    {
     "name": "TSL FR L & R Assembly",
     "crit": [
      "Correct left/right parts",
      "Correct fitment",
      "Wiring / connector condition",
      "Alignment",
      "Final inspection"
     ]
    },
    {
     "name": "Front Fender Assembly",
     "crit": [
      "Correct fitment",
      "Alignment",
      "Fastener tightening",
      "Damage prevention",
      "Final inspection"
     ]
    },
    {
     "name": "Wheel Assembly – Front",
     "crit": [
      "Correct wheel / component identification",
      "Axle & fitment",
      "Torque / tightening",
      "Alignment",
      "Brake / rotation check"
     ]
    },
    {
     "name": "Muffler Assembly",
     "crit": [
      "Correct fitment",
      "Fastener torque",
      "Clearance",
      "Leakage / abnormal noise check",
      "Final inspection"
     ]
    },
    {
     "name": "Crash Guard Assembly",
     "crit": [
      "Correct position",
      "Fitment",
      "Torque / tightening",
      "Clearance",
      "Final inspection"
     ]
    },
    {
     "name": "Main Footrest Assembly",
     "crit": [
      "Correct fitment",
      "Torque / tightening",
      "Movement / function",
      "Alignment",
      "Final inspection"
     ]
    },
    {
     "name": "Footrest L & R Assembly",
     "crit": [
      "Correct left/right parts",
      "Correct fitment",
      "Torque / tightening",
      "Movement / function",
      "Final inspection"
     ]
    },
    {
     "name": "Unloading",
     "crit": [
      "Safe handling",
      "Correct unloading sequence",
      "Damage prevention",
      "Material identification",
      "Housekeeping"
     ]
    },
    {
     "name": "Testing",
     "crit": [
      "Correct test procedure",
      "Safety compliance",
      "Defect identification",
      "Accurate result recording",
      "Final confirmation"
     ]
    },
    {
     "name": "Minor Repair",
     "crit": [
      "Correct diagnosis",
      "Correct repair method",
      "Proper tool usage",
      "Quality after repair",
      "Documentation / confirmation"
     ]
    },
    {
     "name": "Major Repair",
     "crit": [
      "Correct diagnosis",
      "Correct repair method",
      "Safety compliance",
      "Quality after repair",
      "Final verification"
     ]
    },
    {
     "name": "Identified Damage Parts & Shortages",
     "crit": [
      "Correct identification",
      "Accurate quantity",
      "Proper segregation",
      "Timely reporting",
      "Documentation"
     ]
    },
    {
     "name": "Quality of Work",
     "crit": [
      "First-time-right",
      "Workmanship",
      "Defect prevention",
      "Standard compliance",
      "Final quality awareness"
     ]
    }
   ]
  },
  "rtr_ws": {
   "kind": "stations",
   "title": "SKD workstation evaluation – RTR 160 4V / RTR 200",
   "source": "Test Revised with criteria.xlsx (RTR 160 , 200)",
   "stations": [
    {
     "code": "L-1",
     "items": [
      "Unpack the motorcycle parts from the box.",
      "Remove the polythene covers from the parts.",
      "Fit the fuel tank cover [safety cover] over the fuel tank before starting assembly.",
      "Inspect the parts for defects and report any findings to the Transit Team."
     ]
    },
    {
     "code": "L-2",
     "items": [
      "Place the front wheel, saree guard, crash guard, and seat inside the trolley.",
      "Remove the hardware box, separate the bolts, and arrange them in the tray.",
      "Send the left and right tail covers and front fender to the sub-assembly area."
     ]
    },
    {
     "code": "L-3",
     "items": [
      "Remove the wooden piece fitted at the handle mounting area.",
      "Handle assembly.",
      "Install the left pillion footrest.",
      "Assemble the front brake master pump.",
      "Install the rear number plate.",
      "Remove the crash guard and center-stand mounting brackets from the main wooden plate.",
      "Use the hoist to lift the motorcycle and place it on the two-wheeler assembly trolley.",
      "Insert the rear shock absorber bolt (14 mm).",
      "Remove the rear-wheel mounting bracket and the front-wheel axle.",
      "Fix the center stand.",
      "Secure the rear wheel to the two-wheeler assembly trolley using the ratchet."
     ]
    },
    {
     "code": "SUB Process - 1",
     "items": [
      "Fit the left and right tail covers to the center tail cover.",
      "Fit the tail lamp to the complete tail-cover assembly."
     ]
    },
    {
     "code": "SUB Process - 2",
     "items": [
      "Carry out the engine guard sub-assembly.",
      "Carry out the front fender sub-assembly."
     ]
    },
    {
     "code": "L-4",
     "items": [
      "Fit and adjust the throttle cable.",
      "Pre-fit the crash guard to the motorcycle.",
      "Fit the tail-cover assembly to the motorcycle.",
      "Open the right frame cover and connect the battery terminals.",
      "Tighten the center-stand fasteners.",
      "Tighten the rear shock absorber fasteners correctly.",
      "Install the prop stand and ensure the stand switch is fitted correctly.",
      "Fit and tighten the saree guard correctly."
     ]
    },
    {
     "code": "L-5",
     "items": [
      "Tighten the crash guard fasteners correctly.",
      "Tighten the front brake master pump fasteners correctly.",
      "Tighten the frame-cover fasteners correctly.",
      "Install the speedometer correctly.",
      "Fit the headlamp housing (HSG) correctly.",
      "Install the grab handle correctly."
     ]
    },
    {
     "code": "L-6",
     "items": [
      "Install the muffler assembly correctly.",
      "Install the front wheel correctly and mark the fastener after torque tightening.",
      "Install the front fender correctly.",
      "Fit the engine guard securely to the motorcycle.",
      "Install the front brake caliper correctly onto the front wheel disc assembly.",
      "Route and connect the ABS system wiring correctly.",
      "Fit the speedometer cover bracket correctly to the motorcycle.",
      "Install the front number plate and tighten the rear number plate securely.",
      "Remove the fuel tank cover and release the ratchet.",
      "Unload the motorcycle for testing."
     ]
    },
    {
     "code": "Testing",
     "items": [
      "Carry out the vehicle tests."
     ]
    },
    {
     "code": "Repair",
     "items": [
      "Carry out major and minor repairs as required."
     ]
    }
   ]
  },
  "ckd_xl100": {
   "kind": "stations",
   "title": "CKD workstation evaluation – XL 100",
   "source": "Test Revised with criteria.xlsx (CKD)",
   "stations": [
    {
     "code": "L-1",
     "items": [
      "Attach stay head lamp bracket to the fork legs",
      "Fix fork assy to the frame",
      "Fix handle lock to the frame",
      "Attach sub assembled handle bar and Tighten the steering lock nut",
      "Attach and tighten cap nut to handle column",
      "Arrange wiring harness through the H bracket and attach cable guide LH",
      "Load vehicle to the conveyor",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "L-2",
     "items": [
      "Place Chassis on the work table and Apply grease on the racers and place cup set balls",
      "Attach Front seat on the frame",
      "Route wiring harness on the chassis",
      "Fix horn with front number plate bracket to the vehicle",
      "Connect meter wire sockets with the main wiring harness",
      "Attach rear fender to the vehicle and attach seat mounting nuts",
      "Attach rear fender front side bottom bolt",
      "Attach air cleaner assembly to the vehicle and attach cable guide RH to the frame",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "L-3",
     "items": [
      "Attach shock absorber to frame",
      "Fix tail lamp to vehicle and Reflector and mud flap assembly",
      "Attach swing arm to the frame",
      "Attach Rear drum sprocket to the swing arm",
      "Tighten both side attached front seat mounting nuts",
      "Cover tool box center rear assembly",
      "Attach luggage carrier mounting bolts to the frame",
      "Attach Pillion Handle to the luggage",
      "Cushion rubber attach to the bracket of the frame",
      "Attach rear wheel to the swing arm assist with other side operator",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "L-4",
     "items": [
      "Attach rear fender to the frame bracket",
      "Attach luggage carrier RH mounting screws",
      "Fix Rear wheel to the vehicle",
      "Attach torque link with the rear brake panel",
      "Shock absorber RH and Lady handle assembly; attach Pillion Handle to the luggage",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "L-5",
     "items": [
      "Attach engine mount bolts",
      "Tighten cover tool box center and luggage mounting screws",
      "Battery assembly to the vehicle",
      "Side stand place to the frame and attach bolt",
      "Tighten side stand fixing bolt",
      "Fix SAI unit to the engine",
      "Front chain sprocket assembly",
      "Route rear brake cable through head lamp bracket",
      "Fix tail lamp to signal lamp bracket and Attach tail lamp to the vehicle",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "L-6",
     "items": [
      "Fix engine to the vehicle and fix Ignition coil mounting bracket to the vehicle",
      "Route engine wire through the tool box bottom and attach cover tool box center rear RH screws",
      "Tighten engine mount bolts and tighten kick lever mounting bolt",
      "Control unit attach to the vehicle and connect wire clips",
      "Mount ignition coil to the chassis, attach engine ground wire",
      "Route choke cable and Acc. Cable through the chassis bracket and air cleaner breather hose attach to the engine",
      "Attach RH side tail lamp bracket bolt",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "L-7",
     "items": [
      "Take the kitting bin and muffler from the trolley and it place on the conveyor",
      "Tighten the drum sprocket nut",
      "Fix pillion footrest LH and Tighten shock absorber top bolt and bottom nut",
      "Attach screw clips and beading to cover tool box and place it on tool box rear",
      "Fix tool box RH screws and Attach and tighten air cleaner mount screw",
      "Attach front fender mounting RH side screws",
      "Attach front brake cable with brake panel and attach choke cable to lever",
      "Fix sari guard to the vehicle",
      "Chain case assembly",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "L-8",
     "items": [
      "Attach fender front RH mounting screws",
      "Attach split pin and Tighten rear wheel axle nut",
      "Fix muffler to the vehicle",
      "Removed head lamp place on the conveyor",
      "Tighten tool box mounting RH screws and Tighten air cleaner bellow screw",
      "Attach rear brake cable and adjust rear brake as required",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "L-9",
     "items": [
      "Attach fuel tank on the frame",
      "Tighten tail lamp mounting bracket screws and fender front mounting screws",
      "Fix engine sprocket cover",
      "Tighten fuel tank mounting screws",
      "Attach cover frame to frame",
      "Attach head lamp housing to the stay head lamp bracket",
      "Connect all wire sockets and connectors in head lamp housing",
      "Fix seat lock to the carrier",
      "Fix head lamp to the head lamp rear",
      "Cable clip attach with the rear brake cable and choke cable",
      "Adjust the front brake",
      "Take the empty hardware bin on the conveyor and it place on the side trolley",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "L-10",
     "items": [
      "Attach acc. Cable top end to the acc. Grip and tighten control switch bolts",
      "Fix pilion foot rest RH and Tighten shock absorber bolts and Tighten swing arm nut",
      "Attach and tighten both side floor board mounting screws and Tighten fender front stabiliser tube RH screws",
      "Fix leg guard to the vehicle",
      "Connect LH/RH TSL couplers and attach cable strip for handle switch wire",
      "Fix tool box cover lock to the utility box",
      "Attach and tighten head lamp mounting screw and fix rear seat",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "Testing",
     "items": [
      "Place vehicle on the testing bay",
      "Follow the tester's check list for test the vehicle",
      "Rectify defects which can be rectify within the stage time",
      "Remove fuel hose from tank and connect testing fuel can tube and fill fuel to the carb sump",
      "Start the vehicle using self starter and do the necessary checks",
      "Start the vehicle using kick starter",
      "Do the test run",
      "If found any defect which cant be rectified within stage time, update the defect card",
      "Park vehicle on the relevant work station & return back to the testing area"
     ]
    }
   ]
  },
  "w3_ws": {
   "kind": "stations",
   "title": "3W workstation evaluation – 3W / EV Max 3W",
   "source": "Test Revised with criteria.xlsx (3Wheel)",
   "stations": [
    {
     "code": "W-1",
     "items": [
      "3W structure tilting and Depacking",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "W-2",
     "items": [
      "Prepare for assembling",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "W-3",
     "items": [
      "Remove battery from the vehicle & send to charging operation",
      "Fit charged battery to the vehicle",
      "Fix side boards, passenger seat & passenger backrest to the vehicle",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "W-4",
     "items": [
      "Remove driver seat mounting bracket",
      "Fix driver seat mounting bracket",
      "Check tightness of the passenger compartment mounting bolts",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "W-5",
     "items": [
      "Pre fit tail lamps LHS/RHS & number plate lamp",
      "Fix rear bumper to the vehicle",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "W-6",
     "items": [
      "Fix rear shock absorber and fix Inner Mud Flap LHS",
      "Pre-fit rear inner mudguard assy",
      "Fix rear LH side mud flap to the vehicle",
      "Fix rear hood bar cage & driver backrest to the vehicle",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "W-7",
     "items": [
      "Fix rear shock absorber and fix Inner Mud Flap RHS",
      "Pre-fit rear inner mudguard assy",
      "Fix rear RH side mud flap to the vehicle and rear number plate to the vehicle",
      "Fix rear hood bar cage & driver backrest to the vehicle",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "W-8",
     "items": [
      "Fit steering column & front wheel to the vehicle and Tighten steering column & fit handle bar",
      "Tighten the front wheel mounting nut and the handle mounting bolt",
      "Fix footrest to flow board",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "W-9",
     "items": [
      "Pre fit wind shield frame",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "W-10",
     "items": [
      "Fix wiper motor, wiper blade & front number plate to the vehicle",
      "Fix windshield frame brackets to the vehicle",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "W-11",
     "items": [
      "Fix windshield apron plastic cover",
      "Attach plastic cover to utility box",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "W-12",
     "items": [
      "Attach the center hood bar to the vehicle",
      "Fix side hood bar LHS to the vehicle",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "Testing",
     "items": [
      "Remove air bubbles in brake system and Unload the vehicle from 3W dolley",
      "Fill oil to the engine & differential",
      "Fuel filling & wheel nut tightening",
      "Vehicle testing"
     ]
    },
    {
     "code": "Hood",
     "items": [
      "Fix hood strips to the vehicle",
      "Fix hood timber to the vehicle",
      "Hood making-A",
      "Hood making-B",
      "Hood making-C",
      "Remove the LH side door lock and side door lock bracket (DVA)",
      "Record nonconformity and move to next vehicle"
     ]
    },
    {
     "code": "Repair",
     "items": [
      "Minor and Major rectification work"
     ]
    }
   ]
  }
 },
 "lines": [
  {
   "id": "SKD",
   "name": "SKD",
   "models": [
    "Sport 110",
    "Ntorq 125",
    "NTORQ 150",
    "Ronin 225",
    "RTR 160 4V",
    "RTR 200",
    "RTR 310",
    "RTR 310 RR",
    "iQube",
    "Orbiter",
    "Jupiter",
    "Raider 125"
   ]
  },
  {
   "id": "CKD",
   "name": "CKD",
   "models": [
    "XL 100 i-Touch"
   ]
  },
  {
   "id": "3W",
   "name": "3W",
   "models": [
    "3W",
    "EV Max 3W"
   ]
  }
 ],
 "modelTemplate": {
  "Sport 110": "skd19",
  "Ntorq 125": "skd19",
  "NTORQ 150": "skd19",
  "Ronin 225": "skd19",
  "RTR 160 4V": "rtr_ws",
  "RTR 200": "rtr_ws",
  "RTR 310": "skd19",
  "RTR 310 RR": "skd19",
  "iQube": "skd19",
  "Orbiter": "skd19",
  "Jupiter": "skd19",
  "Raider 125": "skd19",
  "XL 100 i-Touch": "ckd_xl100",
  "3W": "w3_ws",
  "EV Max 3W": "w3_ws"
 }
};
// ---- end SPARK_TEMPLATES ----

const SKD_ASSESS_SHEET = 'SKD Assessments';
const WS_ASSESS_SHEET = 'Workstation Assessments';
const SKD_MATRIX_SHEET = 'SKD Skill Matrix';
const ALL_MODELS = SPARK.lines.reduce((a, l) => a.concat(l.models), []);
const MODEL_LINE = {};
SPARK.lines.forEach(l => l.models.forEach(m => { MODEL_LINE[m] = l.id; }));
const SKD_AREAS = SPARK.templates.skd19.areas.map(a => [a.name, a.crit]);
const SKD_FIRST_AREA_COL = 11; // K, in "SKD Assessments"
const MATRIX_HEADERS = ['Employee Name', 'Employee Type', 'ID Number'].concat(ALL_MODELS)
  .concat(['Overall %', 'Lines', 'Models Assessed', 'Last Assessed', 'Last Assessor', 'Phone', 'Photo Link']);
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
function goodPhoto(u) { u = String(u || ''); return u && u.indexOf('PHOTO_SAVE_FAILED') !== 0 ? u : ''; }

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
  return skdSubmit(data);
}

/* ---------- scoring ---------- */
// Returns {error} or {applicable, skilled, results:[1|0|'N/A'], failed:[text], stationText:[text], stations:[codes]}
function scoreItems(templateId, data) {
  const tpl = SPARK.templates[templateId];
  const items = data.items || data.areas;   // 'areas' = name used by the first assess.html version
  let list;                                   // [{label, crit:[names]}] in template order
  let stations = [];
  if (tpl.kind === 'areas') {
    list = tpl.areas.map((a, i) => ({ label: (i + 1) + '. ' + a.name, crit: a.crit, station: '' }));
  } else {
    const wanted = Array.isArray(data.stations) ? data.stations.map(String) : [];
    stations = tpl.stations.map(s => s.code).filter(c => wanted.indexOf(c) !== -1);
    if (!stations.length) return { error: 'Choose at least one workstation.' };
    list = [];
    tpl.stations.forEach(s => {
      if (stations.indexOf(s.code) === -1) return;
      s.items.forEach((t, i) => list.push({ label: s.code + ' #' + (i + 1) + ' ' + t, crit: SPARK.criteria5, station: s.code }));
    });
  }
  if (!Array.isArray(items) || items.length !== list.length) {
    return { error: 'Expected ' + list.length + ' items, got ' + (items ? items.length : 0) + '. Reload the page and try again.' };
  }
  const results = [], failed = [], per = {};
  let applicable = 0, skilled = 0;
  for (let k = 0; k < list.length; k++) {
    const m = items[k] || {}, it = list[k];
    if (it.station && !per[it.station]) per[it.station] = { a: 0, s: 0 };
    if (m.na === true) { results.push('N/A'); continue; }
    const crit = Array.isArray(m.crit) ? m.crit : [];
    if (crit.length !== it.crit.length || crit.some(v => v !== 0 && v !== 1)) {
      return { error: it.label + ' is not fully marked.' };
    }
    applicable++;
    if (it.station) per[it.station].a++;
    const bad = it.crit.filter((_, i) => crit[i] !== 1);
    if (!bad.length) { skilled++; results.push(1); if (it.station) per[it.station].s++; }
    else { results.push(0); failed.push(it.label + ': ' + bad.join(', ')); }
  }
  if (!applicable) return { error: 'Everything is N/A. At least one item must be assessed.' };
  const stationText = stations.map(c => {
    const p = per[c];
    return c + ': ' + (p.a ? p.s + '/' + p.a + ' (' + Math.round(p.s / p.a * 100) + '%)' : 'N/A');
  });
  return { applicable: applicable, skilled: skilled, results: results, failed: failed, stationText: stationText, stations: stations };
}

function skdSubmit(data) {
  const model = String(data.model || '');
  const templateId = SPARK.modelTemplate[model];
  if (!templateId) return skdReply({ status: 'error', message: 'Unknown model: ' + model });
  const idNumber = normId(data.idNumber);
  const name = String(data.name || '').trim();
  const phone = String(data.phone || '').trim().replace(/^'+/, '');
  if (!idNumber || !name) return skdReply({ status: 'error', message: 'Employee name and ID are required.' });

  const sc = scoreItems(templateId, data);
  if (sc.error) return skdReply({ status: 'error', message: sc.error });
  const modelPct = Math.round(sc.skilled / sc.applicable * 100);
  const empType = data.employeeType === 'company' ? 'Company' : 'Manpower';
  const now = new Date();
  const photoUrl = data.photoBase64 ? skdSavePhoto(data.photoBase64, name, idNumber) : '';

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    if (SPARK.templates[templateId].kind === 'areas') {
      const row = [now, String(data.assessor || ''), name, empType, idNumber, model,
                   sc.applicable, sc.skilled, modelPct / 100, String(data.remarks || '')];
      row.push(...sc.results);
      row.push(sc.failed.join('\n'), phone ? "'" + phone : '', photoUrl);
      skdAssessSheet(ss).appendRow(row);
    } else {
      wsAssessSheet(ss).appendRow([now, String(data.assessor || ''), name, empType, idNumber, model, MODEL_LINE[model],
        sc.stations.join(', '), sc.applicable, sc.skilled, modelPct / 100, String(data.remarks || ''),
        sc.stationText.join('\n'), sc.failed.join('\n'), phone ? "'" + phone : '', photoUrl]);
    }
    const emps = skdRebuildMatrix(ss);
    const e = emps[idNumber];
    const matrix = {};
    ALL_MODELS.forEach(m => { matrix[m] = e.models[m] ? e.models[m].pct : null; });
    return skdReply({
      status: 'success', model: model, line: MODEL_LINE[model], applicable: sc.applicable, skilled: sc.skilled,
      modelPct: modelPct, overallPct: e.overallPct, matrix: matrix, photo: photoThumb(e.photoUrl)
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

/* ---------- reading the logs ---------- */
// Latest basic exam (Results tab) row per ID.
function skdBasicResultsById(ss) {
  const out = {};
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 11).getValues().forEach(v => {
    const id = normId(v[4]);
    if (!id || !String(v[1] || '').trim()) return;   // skip blank / junk rows
    out[id] = {                                       // later rows overwrite = latest attempt wins
      name: String(v[1]), phone: String(v[2] || ''),
      empType: String(v[3] || '').toLowerCase() === 'company' ? 'Company' : 'Manpower',
      score: v[5], total: v[6], pct: pctNumber(v[7]), category: String(v[8] || ''),
      date: fmtDate(v[0]), photoUrl: goodPhoto(v[10])
    };
  });
  return out;
}

// Every assessment from both logs as {ts, assessor, name, empType, id, model, pct, phone, photo}.
function skdLogRows(ss) {
  const rows = [];
  const read = (sheetName, width, map) => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return;
    sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues().forEach(v => {
      const id = normId(v[4]);
      if (!id || !(v[0] instanceof Date)) return;
      rows.push(map(v, id));
    });
  };
  const nA = SKD_FIRST_AREA_COL - 1 + SKD_AREAS.length + 3;  // .. AD Failed, AE Phone, AF Photo
  read(SKD_ASSESS_SHEET, nA, (v, id) => ({ ts: v[0], assessor: String(v[1]), name: String(v[2]), empType: String(v[3]), id: id,
    model: String(v[5]), pct: Math.round(Number(v[8]) * 100), phone: String(v[nA - 2] || '').replace(/^'+/, ''), photo: goodPhoto(v[nA - 1]) }));
  read(WS_ASSESS_SHEET, 16, (v, id) => ({ ts: v[0], assessor: String(v[1]), name: String(v[2]), empType: String(v[3]), id: id,
    model: String(v[5]), pct: Math.round(Number(v[10]) * 100), phone: String(v[14] || '').replace(/^'+/, ''), photo: goodPhoto(v[15]) }));
  rows.sort((a, b) => a.ts - b.ts);
  return rows;
}

// {id: {id, name, empType, phone, photoUrl, models:{model:{pct,date}}, overallPct, lines, modelsAssessed, last, lastAssessor, basic}}
// Includes basic-exam-only people (no models).
function skdEmployees(ss) {
  const basic = skdBasicResultsById(ss);
  const emps = {};
  skdLogRows(ss).forEach(r => {
    if (!MODEL_LINE[r.model]) return;            // models no longer in the list (e.g. old single "RTR") are ignored
    const e = emps[r.id] || (emps[r.id] = { id: r.id, models: {}, phone: '', photoUrl: '' });
    e.name = r.name; e.empType = r.empType; e.last = r.ts; e.lastAssessor = r.assessor;
    if (r.phone) e.phone = r.phone;
    if (r.photo) e.photoUrl = r.photo;
    e.models[r.model] = { pct: r.pct, date: fmtDate(r.ts) };
  });
  Object.keys(emps).forEach(id => {
    const e = emps[id];
    const lines = SPARK.lines.filter(l => l.models.some(m => e.models[m]));
    let sum = 0, n = 0;
    lines.forEach(l => l.models.forEach(m => { n++; if (e.models[m]) sum += e.models[m].pct; }));
    e.overallPct = n ? Math.round(sum / n) : null;
    e.lines = lines.map(l => l.id);
    e.modelsAssessed = Object.keys(e.models).length;
  });
  Object.keys(basic).forEach(id => {
    const b = basic[id];
    const e = emps[id] || (emps[id] = { id: id, name: b.name, empType: b.empType, models: {}, overallPct: null, lines: [], modelsAssessed: 0, phone: '', photoUrl: '' });
    if (!e.phone) e.phone = b.phone;
    if (!e.photoUrl) e.photoUrl = b.photoUrl;
    e.basic = { pct: b.pct, score: b.score, total: b.total, category: b.category, date: b.date };
  });
  return emps;
}

/* ---------- actions ---------- */
function skdLookup(idNumber) {
  if (!idNumber) return skdReply({ status: 'error', message: 'Enter the ID number first.' });
  const e = skdEmployees(SpreadsheetApp.openById(SHEET_ID))[idNumber];
  if (!e) return skdReply({ status: 'success', found: false, employee: null, models: {}, overallPct: null, basic: null });
  return skdReply({
    status: 'success', found: true,
    employee: { name: e.name, empType: e.empType, phone: e.phone, photo: photoThumb(e.photoUrl) },
    models: e.models, overallPct: e.overallPct, basic: e.basic || null
  });
}

function skdAllEmployees() {
  const emps = skdEmployees(SpreadsheetApp.openById(SHEET_ID));
  return Object.keys(emps).map(id => {
    const e = emps[id];
    const models = {};
    Object.keys(e.models).forEach(m => { models[m] = e.models[m].pct; });
    return { id: e.id, name: e.name, empType: e.empType, phone: e.phone, photo: photoThumb(e.photoUrl),
             models: models, overallPct: e.overallPct, lines: e.lines, modelsAssessed: e.modelsAssessed,
             lastAssessed: e.last ? fmtDate(e.last) : '', basic: e.basic || null };
  });
}

/* ---------- sheets ---------- */
// Rewrites "SKD Skill Matrix" from the logs (one row per assessed employee). Returns the employees map.
function skdRebuildMatrix(ss) {
  const emps = skdEmployees(ss);
  let sh = ss.getSheetByName(SKD_MATRIX_SHEET) || ss.insertSheet(SKD_MATRIX_SHEET);
  const W = MATRIX_HEADERS.length;
  const cur = sh.getLastRow() > 0 ? sh.getRange(1, 1, 1, Math.max(W, sh.getLastColumn())).getValues()[0] : [];
  const headerOk = MATRIX_HEADERS.every((h, i) => cur[i] === h) && cur.length === W;
  if (!headerOk) {
    sh.clear();
    sh.getRange(1, 1, 1, W).setValues([MATRIX_HEADERS]).setFontWeight('bold').setWrap(true);
    sh.setFrozenRows(1);
    sh.setFrozenColumns(3);
    const pctCols = sh.getRange(2, 4, 999, ALL_MODELS.length + 1);
    const c = 'D2';   // ISNUMBER keeps not-assessed (blank) models uncoloured
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER(' + c + '),' + c + '>=0.9)').setBackground('#b7e1b9').setRanges([pctCols]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER(' + c + '),' + c + '>=0.8,' + c + '<0.9)').setBackground('#d9ead3').setRanges([pctCols]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER(' + c + '),' + c + '>=0.7,' + c + '<0.8)').setBackground('#fff2cc').setRanges([pctCols]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER(' + c + '),' + c + '<0.7)').setBackground('#f4cccc').setRanges([pctCols]).build()
    ]);
  } else if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, W).clearContent();
  }
  const rows = Object.keys(emps).map(id => emps[id]).filter(e => e.modelsAssessed > 0)
    .sort((a, b) => (b.overallPct || 0) - (a.overallPct || 0))
    .map(e => [e.name, e.empType, e.id]
      .concat(ALL_MODELS.map(m => e.models[m] ? e.models[m].pct / 100 : ''))
      .concat([e.overallPct / 100, e.lines.join(', '), e.modelsAssessed, e.last, e.lastAssessor,
               e.phone ? "'" + e.phone : '', e.photoUrl]));
  if (rows.length) {
    sh.getRange(2, 1, rows.length, W).setValues(rows);
    sh.getRange(2, 4, rows.length, ALL_MODELS.length + 1).setNumberFormat('0%');
  }
  return emps;
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
  const areaRange = sh.getRange(2, SKD_FIRST_AREA_COL, 999, SKD_AREAS.length);
  sh.setConditionalFormatRules(sh.getConditionalFormatRules().concat([
    SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(1).setBackground('#d9ead3').setRanges([areaRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(0).setBackground('#f4cccc').setRanges([areaRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('N/A').setBackground('#e0e0e0').setRanges([areaRange]).build()
  ]));
  return sh;
}

function wsAssessSheet(ss) {
  const headers = ['Timestamp', 'Assessor', 'Employee Name', 'Employee Type', 'ID Number', 'Model', 'Line',
                   'Workstations', 'Applicable Activities', 'Skilled Activities', 'Model Skill %', 'Remarks',
                   'Workstation Results', 'Failed Activities', 'Phone', 'Photo Link'];
  let sh = ss.getSheetByName(WS_ASSESS_SHEET);
  if (sh && sh.getLastRow() > 0) { skdEnsureHeaders(sh, headers); return sh; }
  if (!sh) sh = ss.insertSheet(WS_ASSESS_SHEET);
  sh.appendRow(headers);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setWrap(true);
  sh.getRange('K2:K').setNumberFormat('0%');
  return sh;
}

/* =====================================================================
 * 6-MONTHLY SKILL EVALUATION REMINDER (email)
 * A monthly trigger runs reminderTick at about 8 AM on the 1st; it only
 * sends in January and July. Install it ONCE: select setupReminderTrigger
 * in the function dropdown, click Run, approve the permissions.
 * sendReminderTest sends the same email to the script owner only.
 * ===================================================================== */
const REMINDER_TO = ['dinidu@tvslanka.lk', 'malindu.u@tvslanka.lk', 'hrdept@tvslanka.lk'];
const REMINDER_MONTHS = [0, 6];   // January, July (0 = January)
const PORTAL_URL = 'https://denuwantvsl-max.github.io/tvs-spark-portal/';
const OVERDUE_DAYS = 182;

function setupReminderTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'reminderTick')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('reminderTick').timeBased().onMonthDay(1).atHour(8).create();
  console.log('Reminder trigger installed: 1st of every month ~8 AM, sends in January and July.');
}

function reminderTick() {
  const month = Number(Utilities.formatDate(new Date(), TZ, 'M')) - 1;
  if (REMINDER_MONTHS.indexOf(month) === -1) return;
  sendSkillEvalReminder(REMINDER_TO);
}

function sendReminderTest() {
  sendSkillEvalReminder([Session.getEffectiveUser().getEmail()]);
}

function sendSkillEvalReminder(recipients) {
  const emps = skdEmployees(SpreadsheetApp.openById(SHEET_ID));
  const now = new Date();
  const overdue = [], neverAssessed = [];
  let assessedCount = 0;
  Object.keys(emps).forEach(id => {
    const e = emps[id];
    if (e.modelsAssessed > 0) {
      assessedCount++;
      const days = (e.last instanceof Date) ? Math.floor((now - e.last) / 86400000) : 9999;
      if (days >= OVERDUE_DAYS) overdue.push([e.name, e.id, e.lines.join(', '), fmtDate(e.last), e.overallPct + '%']);
    } else if (e.basic) {
      neverAssessed.push([e.name, e.id, e.basic.pct + '% ' + e.basic.category, e.basic.date]);
    }
  });

  const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const table = (head, rows) => rows.length
    ? '<table cellpadding="6" style="border-collapse:collapse;font-size:13px">' +
      '<tr style="background:#1c2126;color:#fff">' + head.map(h => '<th align="left">' + esc(h) + '</th>').join('') + '</tr>' +
      rows.map(r => '<tr style="border-bottom:1px solid #ddd">' + r.map(c => '<td>' + esc(c) + '</td>').join('') + '</tr>').join('') +
      '</table>'
    : '<p style="color:#5b6670">None.</p>';

  const period = Utilities.formatDate(now, TZ, 'MMMM yyyy');
  const html =
    '<div style="font-family:Arial,sans-serif;color:#1c2126;max-width:680px">' +
    '<h2 style="margin:0 0 4px">SPARK — 6-monthly skill evaluation due</h2>' +
    '<p style="margin:0 0 16px;color:#5b6670">TVS Lanka Horana Assembly Plant · ' + esc(period) + '</p>' +
    '<p>This is the scheduled reminder to carry out the 6-monthly skill evaluation of assembly staff ' +
    '(SKD, CKD and 3W lines, one model at a time).</p>' +
    '<p><b>' + overdue.length + '</b> of ' + assessedCount + ' assessed employees have not been assessed in the last 6 months.<br>' +
    '<b>' + neverAssessed.length + '</b> people have a basic exam result but no skill assessment yet.</p>' +
    '<p>' +
    '<a href="' + PORTAL_URL + 'assess.html" style="background:#f2b705;color:#1c2126;padding:10px 16px;text-decoration:none;font-weight:bold;border-radius:3px">Start an assessment</a> ' +
    '&nbsp;<a href="' + PORTAL_URL + 'matrix.html" style="color:#1565c0">Open the employee skill matrix</a> · ' +
    '<a href="' + PORTAL_URL + 'home.html" style="color:#1565c0">SPARK home</a></p>' +
    '<h3>Due for re-assessment (last assessed 6+ months ago)</h3>' +
    table(['Employee', 'ID', 'Lines', 'Last assessed', 'Overall'], overdue) +
    '<h3>Basic exam done, not yet skill-assessed</h3>' +
    table(['Name', 'ID', 'Basic exam', 'Exam date'], neverAssessed) +
    '<p style="color:#5b6670;font-size:12px;margin-top:20px">Sent automatically by SPARK on 1 January and 1 July. ' +
    'The assessor PIN is needed to open the assessment and matrix pages.</p></div>';

  MailApp.sendEmail({
    to: recipients.join(','),
    subject: 'SPARK: 6-monthly skill evaluation due (' + period + ')',
    htmlBody: html,
    name: 'SPARK – TVS Lanka Horana'
  });
  console.log('Reminder sent to ' + recipients.join(', ') + ' (overdue ' + overdue.length + ', never assessed ' + neverAssessed.length + ')');
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'SPARK backend is live' }))
    .setMimeType(ContentService.MimeType.JSON);
}
