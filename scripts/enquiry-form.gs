/**
 * Grid Indigo – website enquiry form receiver (Google Apps Script, free).
 *
 * What it does when someone submits "Discuss a Requirement" on the website:
 *   1. Saves the answers as a new row in the "Enquiries" tab of this Google Sheet.
 *   2. Emails GI (GI_EMAIL) with the full enquiry. Hitting Reply answers the client directly.
 *   3. Sends the client a short automatic acknowledgement.
 *
 * Setup: see the steps in the GI chat, or the README note at the bottom of this file.
 * Nothing here is secret; do not add passwords or keys to this file.
 */

const GI_EMAIL = 'gridindigosystems@gmail.com';
const SHEET_NAME = 'Enquiries';

// Column order in the sheet. Keys match the website form's field names.
const FIELDS = [
  ['received', 'Received'],
  ['first_name', 'First name'],
  ['last_name', 'Last name'],
  ['email', 'Email'],
  ['company', 'Company'],
  ['role', 'Role'],
  ['website', 'Website'],
  ['org_type', 'Organisation type'],
  ['country', 'Country'],
  ['areas', 'Areas'],
  ['description', 'Requirement'],
  ['supports', 'Supports'],
  ['goal', 'Goal'],
  ['stage', 'Stage'],
  ['model', 'Engagement model'],
  ['start', 'Start'],
  ['budget', 'Budget'],
  ['details', 'Details'],
  ['source', 'Heard about us'],
  ['status', 'Status'],
];

const MAX_PER_EMAIL_PER_HOUR = 3; // stops someone using the form to flood one inbox

function doPost(e) {
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    // Spam trap: real visitors never fill this hidden field.
    if (d._gotcha) return reply_({ ok: true });

    const email = clean_(d.email, 200);
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email) || !clean_(d.first_name, 100) || !clean_(d.description, 5000)) {
      return reply_({ ok: false, error: 'missing' });
    }

    const cache = CacheService.getScriptCache();
    const key = 'e:' + email.toLowerCase();
    const count = Number(cache.get(key) || 0);
    if (count >= MAX_PER_EMAIL_PER_HOUR) return reply_({ ok: true });
    cache.put(key, String(count + 1), 3600);

    d.received = new Date();
    d.status = 'New';

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      sheet_().appendRow(FIELDS.map(([k]) => (k === 'received' ? d.received : cell_(d[k]))));
    } finally {
      lock.releaseLock();
    }

    notifyGi_(d, email);
    if (MailApp.getRemainingDailyQuota() > 5) acknowledge_(d, email);

    return reply_({ ok: true });
  } catch (err) {
    console.error(err);
    return reply_({ ok: false, error: 'server' });
  }
}

/** Run once from the editor to create the tab and header row. */
function setup() {
  sheet_();
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(FIELDS.map(([, label]) => label));
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, FIELDS.length).setFontWeight('bold').setBackground('#e8e3f3');
  }
  return sh;
}

function notifyGi_(d, email) {
  const who = [clean_(d.first_name, 100), clean_(d.last_name, 100)].join(' ').trim();
  const company = clean_(d.company, 200) || 'no company given';
  const lines = FIELDS
    .filter(([k]) => k !== 'received' && k !== 'status' && clean_(d[k], 5000))
    .map(([k, label]) => label + ': ' + clean_(d[k], 5000));
  MailApp.sendEmail({
    to: GI_EMAIL,
    replyTo: email,
    name: 'GI Website',
    subject: 'New enquiry: ' + company + ' (' + who + ')',
    body: 'A new enquiry came in from the website.\n\n' + lines.join('\n') +
      '\n\nReply to this email to answer the client directly. All enquiries are in the "Enquiries" tab of the GI Website Enquiries sheet.',
  });
}

function acknowledge_(d, email) {
  // Fixed wording on purpose: nothing the visitor typed is repeated except their first name,
  // so the form cannot be used to send someone else a message.
  const first = clean_(d.first_name, 40).replace(/[<>:/\\]/g, '') || 'there';
  MailApp.sendEmail({
    to: email,
    replyTo: GI_EMAIL,
    name: 'Grid Indigo',
    subject: 'We received your enquiry – Grid Indigo',
    body:
      'Hi ' + first + ',\n\n' +
      'Thank you for reaching out to Grid Indigo. We have received your requirement and will review it shortly.\n\n' +
      'If there is a potential fit, we will contact you to talk through the scope, priorities and next steps.\n\n' +
      'Kind regards,\nGrid Indigo\nYour digital delivery partner for what’s next.\n\n' +
      'This is an automatic acknowledgement. If you did not submit an enquiry, you can ignore this email.',
  });
}

function clean_(v, max) {
  return String(v == null ? '' : v).trim().slice(0, max);
}

// Stops a value like "=IMPORTDATA(...)" from running as a formula in the sheet.
function cell_(v) {
  const s = clean_(v, 5000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function reply_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/*
README – setup (about 15 minutes, signed in as gridindigosystems@gmail.com)
1. Google Drive → New → Google Sheets. Name it "GI Website Enquiries".
2. In the sheet: Extensions → Apps Script. Delete what is there, paste this whole file, Save.
3. Choose the function "setup" at the top and click Run. Approve the permissions (Advanced → Go to project).
4. Deploy → New deployment → type "Web app".
   Execute as: Me (gridindigosystems@gmail.com). Who has access: Anyone. → Deploy.
5. Copy the Web app URL (ends in /exec) and send it to Claude. It goes into ENDPOINT on get-started.html.
If you change this script later: Deploy → Manage deployments → Edit → Version: New version (the URL stays the same).
*/
