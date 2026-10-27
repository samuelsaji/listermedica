/**
 * Listermedica website enquiries -> Google Sheet + email alert
 *
 * Setup (one time):
 * 1. Create a Google Sheet, e.g. "Listermedica Website Enquiries".
 * 2. In the sheet: Extensions > Apps Script. Delete the sample code and paste this file.
 * 3. Click Deploy > New deployment > type "Web app".
 *      Execute as: Me
 *      Who has access: Anyone
 *    Click Deploy and approve the permissions.
 * 4. Copy the Web app URL (ends in /exec) and put it in SCRIPT_URL in index.html.
 */

const NOTIFY_EMAIL = 'listermedica.cok@gmail.com'; // who gets the alert (comma-separate for several)
const SHEET_NAME = 'Enquiries';
const HEADERS = ['Received', 'Name', 'Organisation', 'Phone', 'Email', 'Interested in', 'Requirement'];

function doPost(e) {
  const p = (e && e.parameter) || {};

  // Spam trap: real visitors never fill this hidden field.
  if (p.website) return json_({ ok: true });

  const clean = (v, max) => {
    let s = String(v || '').trim().slice(0, max || 500);
    if (/^[=+\-@]/.test(s)) s = "'" + s; // stop formulas being injected into the sheet
    return s;
  };
  const row = {
    name: clean(p.name, 120),
    org: clean(p.org, 160),
    phone: clean(p.phone, 40),
    email: clean(p.email, 160),
    need: clean(p.need, 80),
    message: clean(p.message, 3000),
  };
  if (!row.name || !row.phone) return json_({ ok: false, error: 'missing fields' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    sh.appendRow([new Date(), row.name, row.org, row.phone, row.email, row.need, row.message]);
  } finally {
    lock.releaseLock();
  }

  const mail = {
    to: NOTIFY_EMAIL,
    subject: 'New website enquiry: ' + (row.need || 'General') + ' - ' + (row.org || row.name),
    body:
      'New enquiry from listermedica.com\n\n' +
      'Name: ' + row.name + '\n' +
      'Organisation: ' + row.org + '\n' +
      'Phone: ' + row.phone + '\n' +
      'Email: ' + (row.email || '-') + '\n' +
      'Interested in: ' + row.need + '\n\n' +
      'Requirement:\n' + (row.message || '-') + '\n\n' +
      'All enquiries: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl(),
  };
  if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(row.email)) mail.replyTo = row.email;
  MailApp.sendEmail(mail);

  return json_({ ok: true });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
