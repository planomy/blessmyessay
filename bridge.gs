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

  /** QUIK DRAFT FETCH — returns next row not marked processed (see isRowProcessed_). */
  if (raw.action === 'fetchNext') {
    return fetchNextHandler_();
  }

  /** QUIK DRAFT SEND — email feedback, write J, mark I/K/L processed. */
  if (raw.action === 'sendFeedback') {
    return sendFeedbackHandler_(raw);
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

/** Column I (9): teacher marked “Sent”. Column L (12): “YES” = grading complete. Either marks row processed for FETCH. */
function isRowProcessed_(sheet, row) {
  var colI = String(sheet.getRange(row, 9).getValue() || '')
    .trim()
    .toLowerCase();
  var colL = sheet.getRange(row, 12).getValue();
  var lStr = String(colL === true ? 'yes' : colL || '')
    .trim()
    .toUpperCase();
  if (colI === 'sent') return true;
  if (lStr === 'YES' || lStr === 'TRUE') return true;
  return false;
}

/**
 * First data row (≥2) that is not processed and has a name or essay in the sheet.
 * Matches quikdraft/index.html: row, name, task, essay, focus / teacherConcerns.
 */
function fetchNextHandler_() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      return corsJsonTextOutput_({ ok: false, error: 'Sheet not found: ' + SHEET_NAME });
    }
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return corsJsonTextOutput_({ row: null });
    }
    for (var r = 2; r <= lastRow; r++) {
      if (isRowProcessed_(sheet, r)) continue;
      var name = sheet.getRange(r, 2).getValue();
      var essay = sheet.getRange(r, 7).getValue();
      if (!name && !essay) continue;
      var concerns = String(sheet.getRange(r, 6).getValue() || '');
      return corsJsonTextOutput_({
        row: r,
        name: name != null ? String(name) : '',
        task: String(sheet.getRange(r, 4).getValue() || ''),
        essay: essay != null ? String(essay) : '',
        teacherConcerns: concerns,
        focus: concerns
      });
    }
    return corsJsonTextOutput_({ row: null });
  } catch (err) {
    return corsJsonTextOutput_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

/** QUIK DRAFT — column J feedback text, C email, I/K/L status (same indices as fetchNext skip rules). */
var FEEDBACK_TEXT_COL = 10;
var STUDENT_EMAIL_COL = 3;
var STUDENT_NAME_COL = 2;

function escapeHtml_(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function sendFeedbackHandler_(data) {
  var row = data.row;
  var feedbackText = data.feedbackText != null ? String(data.feedbackText) : '';
  var feedbackHtml = data.feedbackHtml != null ? String(data.feedbackHtml) : '';

  if (!row || Number(row) < 2) {
    return corsJsonTextOutput_({ ok: false, error: 'Invalid row' });
  }

  if (!feedbackText.trim() && !feedbackHtml.trim()) {
    return corsJsonTextOutput_({ ok: false, error: 'Missing feedback text' });
  }

  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      return corsJsonTextOutput_({ ok: false, error: 'Sheet not found: ' + SHEET_NAME });
    }

    var email = String(sheet.getRange(Number(row), STUDENT_EMAIL_COL).getValue() || '').trim();
    if (!email) {
      return corsJsonTextOutput_({ ok: false, error: 'No student email in row' });
    }

    var studentName = String(sheet.getRange(Number(row), STUDENT_NAME_COL).getValue() || '').trim();
    var subject = studentName ? ('Writing feedback: ' + studentName) : 'Your writing feedback';

    sheet.getRange(Number(row), FEEDBACK_TEXT_COL).setValue(feedbackText);

    var plainTextBody = feedbackText;
    var htmlBody = feedbackHtml;
    if (!htmlBody.trim()) {
      htmlBody =
        '<div style="font-family:Segoe UI,system-ui,sans-serif;font-size:11pt;line-height:1.6;color:#334155;">' +
        escapeHtml_(plainTextBody).replace(/\n/g, '<br>') +
        '</div>';
    }

    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: plainTextBody,
      htmlBody: htmlBody
    });

    sheet.getRange(Number(row), 9).setValue('Sent');
    sheet.getRange(Number(row), 11).setValue(new Date());
    sheet.getRange(Number(row), 12).setValue('YES');

    return corsJsonTextOutput_({ ok: true, message: 'Sent' });
  } catch (err) {
    return corsJsonTextOutput_({
      ok: false,
      error: String(err && err.message ? err.message : err)
    });
  }
}
