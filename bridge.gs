/**
 * STUDENT INTAKE → GOOGLE SHEET BRIDGE
 *
 * HOW TO DEPLOY AS A WEB APP
 * --------------------------
 * 1. Open script.google.com, create a project, paste this file (or bind it to your Sheet:
 *    Extensions → Apps Script).
 * 2. Set SPREADSHEET_ID below to your Google Sheet’s ID (from the URL between /d/ and /edit).
 * 3. Row 1 headers should match your sheet (see HEADERS below), e.g.:
 *    Timestamp | Student Name | Student Email | Task Description | Word Count | Teacher Concerns | Essay Text
 * 4. Click Deploy → New deployment → Type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone (or “Anyone with Google account” if you prefer)
 * 5. Authorize when prompted. Copy the Web App URL and use it as the POST endpoint from your form.
 * 6. Your frontend should POST JSON with Content-Type: application/json.
 *
 * CORS: Browsers send OPTIONS before JSON POST. doOptions and doPost attach Access-Control-*
 * headers on the TextOutput. Use the /exec Web App URL from Deploy → Test deployments / Manage
 * deployments. If fetch still fails, confirm “Who has access” matches your form’s audience.
 */

/** Replace with your Google Sheet ID from the spreadsheet URL. */
var SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';

/** Sheet name (tab) to write to. */
var SHEET_NAME = 'Intake';

var HEADERS = [
  'Timestamp',
  'Student Name',
  'Student Email',
  'Task Description',
  'Word Count',
  'Teacher Concerns',
  'Essay Text'
];

/**
 * JSON TextOutput for Web App responses. Sets CORS headers when the runtime exposes
 * TextOutput.setHeader (some environments/tutorials use this; the public reference
 * only documents setMimeType—if setHeader is missing, doOptions + JSON MIME still help clients).
 */
function corsJsonTextOutput_(payload) {
  var json = typeof payload === 'string' ? payload : JSON.stringify(payload);
  var out = ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
  if (typeof out.setHeader === 'function') {
    out.setHeader('Access-Control-Allow-Origin', '*');
    out.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    out.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  }
  return out;
}

/**
 * Preflight for cross-origin JSON POST (browsers send OPTIONS first).
 */
function doOptions() {
  return corsJsonTextOutput_({ ok: true });
}

/**
 * Accepts JSON keys (exact names): name, email, task, wordCount, teacherConcerns, essayText
 * (Optional snake_case: word_count, teacher_concerns, essay_text — prefer camelCase from index.html.)
 *
 * Always returns ContentService TextOutput with MIME JSON and CORS headers so browsers can read
 * the response from another origin (e.g. file:// or your static site).
 */
function doPost(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return corsJsonTextOutput_({ ok: false, error: 'Missing POST body' });
  }

  var raw;
  try {
    raw = JSON.parse(e.postData.contents);
  } catch (err) {
    return corsJsonTextOutput_({ ok: false, error: 'Invalid JSON' });
  }

  var name = pick_(raw, ['name']);
  var email = pick_(raw, ['email']);
  var task = pick_(raw, ['task']);
  var wordCount = pick_(raw, ['wordCount', 'word_count']);
  var teacherConcerns = pick_(raw, ['teacherConcerns', 'teacher_concerns']);
  var essayText = pick_(raw, ['essayText', 'essay_text']);

  if (!name || !email) {
    return corsJsonTextOutput_({ ok: false, error: 'name and email are required' });
  }

  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = getOrCreateSheet_(ss, SHEET_NAME);
    ensureHeaderRow_(sheet);

    // Column order matches HEADERS: Timestamp → … → Essay Text (JSON keys are not the header text).
    var row = [
      new Date(),
      String(name),
      String(email),
      task != null ? String(task) : '',
      wordCount != null ? String(wordCount) : '',
      teacherConcerns != null ? String(teacherConcerns) : '',
      essayText != null ? String(essayText) : ''
    ];

    sheet.appendRow(row);

    return corsJsonTextOutput_({
      ok: true,
      message: 'Submission received and saved successfully.'
    });
  } catch (err) {
    return corsJsonTextOutput_({
      ok: false,
      error: 'Could not write to spreadsheet. Check SPREADSHEET_ID and permissions.'
    });
  }
}

function pick_(obj, keys) {
  for (var i = 0; i < keys.length; i++) {
    if (obj[keys[i]] !== undefined && obj[keys[i]] !== null && obj[keys[i]] !== '') {
      return obj[keys[i]];
    }
  }
  return '';
}

function getOrCreateSheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  return sh;
}

function ensureHeaderRow_(sheet) {
  var first = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  var empty = true;
  for (var c = 0; c < first.length; c++) {
    if (first[c]) {
      empty = false;
      break;
    }
  }
  if (empty) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
}
