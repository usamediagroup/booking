/**
 * USA Media Group — Shoot Booking backend (Google Apps Script)
 *
 * Rules:
 *  - Mon–Fri only, three 1-hour slots: 12:00, 13:00, 14:00 (Denver time)
 *  - ONE booking per day: once any slot is booked, the whole day closes
 *  - To cancel a booking: delete its row, or write "cancelled" in the Status column
 *  - To block a day yourself: add a row with just the date (yyyy-mm-dd) in the Date column
 */

const CONFIG = {
  SHEET_NAME: 'Bookings',
  TIMEZONE: 'America/Denver',
  SLOTS: ['12:00', '13:00', '14:00'],
  WEEKS_AHEAD: 4,          // must match the page
  MIN_DAYS_AHEAD: 1,       // 1 = earliest bookable day is tomorrow
  NOTIFY_EMAIL: '',        // empty = the Google account that owns this script
  SEND_CLIENT_CONFIRMATION: true,
  BUSINESS_NAME: 'USA Media Group',
};

const HEADERS = ['Created', 'Date', 'Time', 'Name', 'Business', 'Phone', 'Email', 'Details', 'Status'];

/** Run once from the editor: creates the sheet and asks for permissions. */
function setup() {
  const sheet = getSheet_();
  Logger.log('Sheet ready: ' + sheet.getName() + '. Booked dates: ' + JSON.stringify(getBookedDates_()));
}

function doGet() {
  return json_({ ok: true, booked: getBookedDates_() });
}

function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad_request' });
  }

  // Honeypot: bots fill the hidden "website" field
  if (data.website) return json_({ ok: true });

  const problem = validate_(data);
  if (problem) return json_({ ok: false, error: problem });

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (err) {
    return json_({ ok: false, error: 'busy_try_again' });
  }

  try {
    const booked = getBookedDates_();
    if (booked.indexOf(data.date) !== -1) {
      return json_({ ok: false, error: 'taken', booked: booked });
    }
    getSheet_().appendRow([
      new Date(),
      data.date,
      data.time,
      clean_(data.name),
      clean_(data.business),
      clean_(data.phone),
      clean_(data.email),
      clean_(data.details),
      'booked',
    ]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  sendEmails_(data);
  return json_({ ok: true });
}

/* ---------- helpers ---------- */

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    // Keep Date and Time as plain text so Sheets doesn't convert them
    sheet.getRange('B:C').setNumberFormat('@');
  }
  return sheet;
}

function getBookedDates_() {
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const rows = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
  const today = todayStr_();
  const out = {};
  rows.forEach(function (r) {
    const date = normalizeDate_(r[1], tz);
    const status = String(r[8] || '').trim().toLowerCase();
    if (date && date >= today && status !== 'cancelled' && status !== 'canceled') out[date] = true;
  });
  return Object.keys(out).sort();
}

function normalizeDate_(v, tz) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  const s = String(v).trim().replace(/^'/, '');
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const d = new Date(s);
  return isNaN(d) ? '' : Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}

function validate_(d) {
  if (!d || typeof d !== 'object') return 'bad_request';
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(d.date || ''))) return 'bad_date';
  if (CONFIG.SLOTS.indexOf(d.time) === -1) return 'bad_time';
  if (!String(d.name || '').trim()) return 'missing_name';
  if (!String(d.phone || '').trim()) return 'missing_phone';
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(d.email || '').trim())) return 'bad_email';

  const dow = new Date(d.date + 'T12:00:00Z').getUTCDay();
  if (dow === 0 || dow === 6) return 'weekend';

  const today = todayStr_();
  if (d.date < addDays_(today, CONFIG.MIN_DAYS_AHEAD)) return 'too_soon';
  if (d.date > addDays_(today, CONFIG.WEEKS_AHEAD * 7)) return 'too_far';
  return '';
}

function sendEmails_(d) {
  const owner = CONFIG.NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
  const when = prettyDate_(d.date) + ' at ' + prettyTime_(d.time) + ' (Mountain Time)';
  const who = String(d.business || '').trim() || String(d.name).trim();

  try {
    MailApp.sendEmail({
      to: owner,
      replyTo: String(d.email).trim(),
      subject: 'New shoot booked: ' + d.date + ' ' + prettyTime_(d.time) + ' — ' + who,
      body:
        'New booking\n\n' +
        'When: ' + when + '\n' +
        'Name: ' + d.name + '\n' +
        'Business: ' + (d.business || '-') + '\n' +
        'Phone: ' + d.phone + '\n' +
        'Email: ' + d.email + '\n' +
        'Details: ' + (d.details || '-') + '\n\n' +
        'To cancel: delete the row in the sheet or set Status to "cancelled".',
    });
  } catch (err) {
    Logger.log('Owner email failed: ' + err);
  }

  if (!CONFIG.SEND_CLIENT_CONFIRMATION) return;
  try {
    MailApp.sendEmail({
      to: String(d.email).trim(),
      replyTo: owner,
      name: CONFIG.BUSINESS_NAME,
      subject: 'Your shoot is booked — ' + prettyDate_(d.date),
      body:
        'Hi ' + String(d.name).trim() + ',\n\n' +
        'Your 1-hour shoot with ' + CONFIG.BUSINESS_NAME + ' is booked for ' + when + '.\n\n' +
        'If you need to reschedule, just reply to this email.\n\n' +
        '— ' + CONFIG.BUSINESS_NAME,
    });
  } catch (err) {
    Logger.log('Client email failed: ' + err);
  }
}

function clean_(v) {
  let s = String(v || '').trim().slice(0, 1000);
  if (/^[=+\-@]/.test(s)) s = "'" + s; // stop formula injection in Sheets
  return s;
}

function todayStr_() {
  return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

function addDays_(dateStr, n) {
  const d = new Date(dateStr + 'T12:00:00Z');
  d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
}

function prettyDate_(dateStr) {
  return Utilities.formatDate(new Date(dateStr + 'T12:00:00Z'), 'UTC', 'EEEE, MMMM d, yyyy');
}

function prettyTime_(t) {
  const h = parseInt(t.split(':')[0], 10);
  return (h % 12 || 12) + ':00 ' + (h < 12 ? 'AM' : 'PM');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
