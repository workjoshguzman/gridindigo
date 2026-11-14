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
// Sender shown to clients. Needs the Resend API key saved in Project Settings > Script Properties as
// RESEND_API_KEY (a key with sending access to gridindigo.com). Without it, emails go out from Gmail instead.
const FROM_ADDRESS = 'enquiry@gridindigo.com';
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

// Internal alert to the GI team: always sent by Gmail (gridindigosystems@gmail.com), never from the client-facing address.
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
  const first = clean_(d.first_name, 40).replace(/[<>:/\\&"']/g, '') || 'there';
  const paragraphs = [
    'Hi ' + first + ',',
    'Thank you for reaching out to Grid Indigo. We have received your requirement and will review it shortly.',
    'If there is a potential fit, we will contact you to talk through the scope, priorities and next steps.',
    'Kind regards,\nGrid Indigo',
  ];
  sendMail_({
    to: email,
    replyTo: GI_EMAIL,
    name: 'Grid Indigo',
    subject: 'We received your enquiry \u2013 Grid Indigo',
    body: paragraphs.join('\n\n') + '\n\n--\n' + FOOTER_LINE + '\n\n' + AUTO_NOTE + '\n\n' + CONFIDENTIALITY_NOTICE + '\n\n' + PRIVACY_NOTICE,
    htmlBody: emailHtml_(paragraphs),
    withLogo: true,
  });
}

// Signature: same layout and notices as Delivery Hub client emails.
const FOOTER_LINE = 'Grid Indigo \u00b7 Digital Delivery & Technology Partner \u00b7 gridindigo.com';
const AUTO_NOTE = 'This is an automatic acknowledgement. If you did not submit an enquiry, you can ignore this email.';
const CONFIDENTIALITY_NOTICE =
  'CONFIDENTIALITY NOTICE: This email and any attachments are confidential and intended only for the named ' +
  'recipient(s). If you received it in error, please notify the sender and delete it. Any unauthorized use, ' +
  'disclosure or copying is prohibited.';
const PRIVACY_NOTICE =
  'DATA PRIVACY: Grid Indigo processes the personal data you submitted only to respond to your enquiry, in line ' +
  'with applicable data protection laws. See our Privacy Policy at gridindigo.com/privacy.html or reply to this ' +
  'email to ask about or update your data.';

function emailHtml_(paragraphs) {
  const esc = (t) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const small = 'font-size:11px;line-height:1.5;color:#8a8aa6;margin:8px 0 0';
  return [
    '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.6;color:#1c1c30">',
    paragraphs.map((p) => '<p style="margin:0 0 14px">' + esc(p).replace(/\n/g, '<br>') + '</p>').join(''),
    '<div style="margin-top:24px;padding-top:12px;border-top:1px solid #e2e2ee">',
    '<img src="cid:giLogo" alt="Grid Indigo" width="120" style="display:block;width:120px;height:auto;border:0;margin:0 0 8px">',
    '<p style="' + small + '">' + esc(FOOTER_LINE) + '</p>',
    '<p style="' + small + '">' + esc(AUTO_NOTE) + '</p>',
    '<p style="' + small + '">' + esc(CONFIDENTIALITY_NOTICE) + '</p>',
    '<p style="' + small + '">' + esc(PRIVACY_NOTICE) + '</p>',
    '</div></div>',
  ].join('');
}

// Grid Indigo v3 wordmark (purple, 240 px wide), embedded so it shows without loading anything external.
const LOGO_PNG_BASE64 = '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';

/**
 * Client-facing emails only. Sends through Resend (from enquiry@gridindigo.com) when the API key is set,
 * otherwise through Gmail.
 * If Resend refuses or cannot be reached, the email still goes out through Gmail.
 */
function sendMail_(m) {
  const key = PropertiesService.getScriptProperties().getProperty('RESEND_API_KEY');
  if (key) {
    try {
      const res = UrlFetchApp.fetch('https://api.resend.com/emails', {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + key },
        muteHttpExceptions: true,
        payload: JSON.stringify({
          from: m.name + ' <' + FROM_ADDRESS + '>',
          to: [m.to],
          reply_to: m.replyTo,
          subject: m.subject,
          text: m.body,
          html: m.htmlBody,
          attachments: m.withLogo
            ? [{ filename: 'grid-indigo.png', content: LOGO_PNG_BASE64, content_type: 'image/png', content_id: 'giLogo' }]
            : undefined,
        }),
      });
      if (res.getResponseCode() < 300) return;
      console.error('Resend refused the email: ' + res.getResponseCode() + ' ' + res.getContentText());
    } catch (err) {
      console.error('Could not reach Resend: ' + err);
    }
  }
  const opts = { to: m.to, replyTo: m.replyTo, name: m.name, subject: m.subject, body: m.body };
  if (m.htmlBody) opts.htmlBody = m.htmlBody;
  if (m.withLogo) opts.inlineImages = { giLogo: Utilities.newBlob(Utilities.base64Decode(LOGO_PNG_BASE64), 'image/png', 'grid-indigo.png') };
  MailApp.sendEmail(opts);
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
6. Optional, to send from enquiry@gridindigo.com: Resend > API Keys > Create (Sending access, domain gridindigo.com).
   Apps Script > Project Settings > Script Properties > Add: RESEND_API_KEY = the key. Never paste the key in this file.
If you change this script later: Deploy → Manage deployments → Edit → Version: New version (the URL stays the same).
*/
