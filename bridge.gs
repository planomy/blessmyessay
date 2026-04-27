/**
 * STUDENT INTAKE → GOOGLE SHEET BRIDGE (Quikdraft: fetch / send)
 *
 * Deploy: Web app, Execute as: Me, Who has access: Anyone.
 * CORS: use the /exec URL. doOptions + JSON responses for cross-origin fetches.
 */

const SPREADSHEET_ID = '1tk4_YTBPWFYYQvCE2jgE8AyCNB4OPP65wdahb0qmL7s';
const SHEET_NAME = 'Submissions';

/**
 * 13 columns A–M. G = essay, J = feedback “Marked Essay”, I/K/L = status (sendFeedback writes I, K, L).
 */
var HEADER_ROW_ = [
  'Timestamp',
  'Name',
  'Email',
  'Task',
  'Word Count',
  'Teacher Concerns',
  'Essay Draft',
  'Submitted At',
  'Status',
  'Marked Essay',
  'Date Returned',
  'Sent',
  'Source'
];

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

function doOptions() {
  return corsJsonTextOutput_({ ok: true });
}

function doGet() {
  return ContentService
    .createTextOutput('Apps Script running')
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * name, email, task, wordCount, teacherConcerns, essay from essayText / essayDraft, submittedAt, source
 * fetchNext, sendFeedback
 */
function doPost(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return corsJsonTextOutput_({ ok: false, error: 'Missing POST body' });
  }

  var data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return corsJsonTextOutput_({ ok: false, error: 'Invalid JSON' });
  }

  if (data.action === 'sendFeedback') {
    return sendFeedbackHandler_(data);
  }

  if (data.action === 'fetchNext') {
    return fetchNextHandler_();
  }

  try {
    var sheet = getSheet_();
    var essay = pick_(
      data,
      ['essayText', 'essay_text', 'essayDraft']
    );

    sheet.appendRow([
      new Date(),
      data.name != null ? String(data.name) : '',
      data.email != null ? String(data.email) : '',
      data.task != null ? String(data.task) : '',
      data.wordCount != null ? String(data.wordCount) : '',
      data.teacherConcerns != null ? String(data.teacherConcerns) : '',
      String(essay),
      data.submittedAt != null ? String(data.submittedAt) : '',
      '',
      '',
      '',
      '',
      data.source != null ? String(data.source) : ''
    ]);

    return corsJsonTextOutput_({ ok: true });
  } catch (err) {
    return corsJsonTextOutput_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function pick_(obj, keys) {
  for (var i = 0; i < keys.length; i++) {
    if (
      obj[keys[i]] !== undefined &&
      obj[keys[i]] !== null &&
      obj[keys[i]] !== ''
    ) {
      return obj[keys[i]];
    }
  }
  return '';
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADER_ROW_.length).setValues([HEADER_ROW_]);
  }
  return sheet;
}

/**
 * First row (after header) with essay in G and no “finished” value in L (12).
 * L truthy = skip (matches your original: !sent on column L).
 * focus = first up to 3 non-empty lines from F, for Quikdraft.
 */
function fetchNextHandler_() {
  try {
    const sheet = getSheet_();
    const values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      var rowData = values[i];
      var essayDraft = rowData[6];
      var sent = rowData[11];
      if (essayDraft && !sent) {
        var focusRaw = String(rowData[5] || '');
        var focus = focusRaw
          .split(/\n|,/g)
          .map(function (f) {
            return f.trim();
          })
          .filter(function (f) {
            return f !== '';
          })
          .slice(0, 3)
          .join('\n');
        return corsJsonTextOutput_({
          row: i + 1,
          name: rowData[1] != null ? String(rowData[1]) : '',
          task: rowData[3] != null ? String(rowData[3]) : '',
          focus: focus,
          teacherConcerns: focusRaw,
          essay: essayDraft != null ? String(essayDraft) : ''
        });
      }
    }
    return corsJsonTextOutput_({ row: null, name: '', task: '', essay: '' });
  } catch (err) {
    return corsJsonTextOutput_({
      ok: false,
      error: String(err && err.message ? err.message : err)
    });
  }
}

var FEEDBACK_TEXT_COL = 10;
var STUDENT_EMAIL_COL = 3;
var STUDENT_NAME_COL = 2;
var TASK_COL = 4;
var STATUS_COL = 9;
var DATE_RETURNED_COL = 11;
var SENT_COL = 12;

function escapeHtml_(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function sendFeedbackHandler_(data) {
  var row = Number(data.row);
  var feedbackText = data.feedbackText != null ? String(data.feedbackText) : '';
  var feedbackHtml = data.feedbackHtml != null ? String(data.feedbackHtml) : '';
  if (!row || row < 2) {
    return corsJsonTextOutput_({ ok: false, error: 'Invalid row' });
  }
  if (!feedbackText.trim() && !feedbackHtml.trim()) {
    return corsJsonTextOutput_({ ok: false, error: 'Missing feedback text' });
  }
  var sheet;
  var email;
  try {
    sheet = getSheet_();
    email = String(sheet.getRange(row, STUDENT_EMAIL_COL).getValue() || '')
      .trim();
    if (!email) {
      return corsJsonTextOutput_({ ok: false, error: 'No student email in row' });
    }
    var name = sheet.getRange(row, STUDENT_NAME_COL).getValue();
    var task = sheet.getRange(row, TASK_COL).getValue();
    sheet.getRange(row, FEEDBACK_TEXT_COL).setValue(feedbackText);
    var taskStr = task != null ? String(task).trim() : '';
    var nameStr = name != null ? String(name).trim() : '';
    var subject;
    if (taskStr) {
      subject = 'Feedback for your ' + taskStr;
    } else if (nameStr) {
      subject = 'Writing feedback: ' + nameStr;
    } else {
      subject = 'Your writing feedback';
    }
    var plainTextBody = feedbackText;
    var htmlBody = feedbackHtml;
    if (htmlBody.trim()) {
      // Root wrapper so clients get a full block; Quikdraft already inlines blue/highlight styles
      htmlBody =
        '<div style="font-family:Segoe UI,system-ui,sans-serif;font-size:11pt;line-height:1.6;color:#1f2937;max-width:100%;">' +
        feedbackHtml +
        '</div>';
    } else {
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
    sheet.getRange(row, STATUS_COL).setValue('Sent');
    sheet.getRange(row, DATE_RETURNED_COL).setValue(new Date());
    sheet.getRange(row, SENT_COL).setValue('YES');
    return corsJsonTextOutput_({ ok: true, message: 'Sent' });
  } catch (err) {
    return corsJsonTextOutput_({
      ok: false,
      error: String(err && err.message ? err.message : err)
    });
  }
}
