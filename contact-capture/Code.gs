/**
 * Contact Card → Google Sheet logger
 * Paste into: Sheet "Contact Card — Incoming Contacts" › Extensions › Apps Script
 * Deploy › New deployment › Web app › Execute as: Me › Who has access: Anyone
 * Copy the /exec URL into CONTACT_ENDPOINT in contact-form-snippet.html
 */
const SHEET_ID = '1y1Hw2NagoSsP79ACvfVLlFwlp68V3wU3JmsSM_s1q6Y';
const NOTIFY_EMAIL = 'sjlove@shannonjeffreylove.com'; // set to '' to turn off email alerts
const MAX_LEN = 500;

function doPost(e) {
  try {
    const p = (e && e.parameter) || {};

    // Spam trap: real people never fill the hidden "website" field
    if (p.website) return json_({ ok: true });

    // Trim first, then neutralize anything Sheets would read as a formula
    const clean = v => {
      const s = String(v || '').trim().slice(0, MAX_LEN);
      return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
    };
    const name = clean(p.name);
    const email = clean(p.email);
    const phone = clean(p.phone);
    if (!name || (!email && !phone)) return json_({ ok: false, error: 'Name and an email or phone are required.' });

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      SpreadsheetApp.openById(SHEET_ID).getSheets()[0].appendRow([
        new Date(), name, email, phone, clean(p.company), clean(p.met), clean(p.note), clean(p.source) || 'contact.shannonjlove.cloud'
      ]);
    } finally {
      lock.releaseLock();
    }

    if (NOTIFY_EMAIL) {
      MailApp.sendEmail(NOTIFY_EMAIL, 'New contact from your card: ' + name,
        ['Name: ' + name, 'Email: ' + email, 'Phone: ' + phone, 'Company / Role: ' + clean(p.company),
         'How we met: ' + clean(p.met), 'Note: ' + clean(p.note)].join('\n'));
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: 'Server error' });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
