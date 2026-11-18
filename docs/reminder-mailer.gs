/**
 * Standalone wedding reminder mailer.
 *
 * IMPORTANT
 * - Create this Apps Script from Extensions > Apps Script inside a separate
 *   "Wedding Reminder Send List" spreadsheet copy.
 * - Do not add this file to the production RSVP endpoint project.
 * - This script only reads the active spreadsheet and only writes by appending
 *   a result row to the "Send Log" tab.
 * - It installs no triggers and sends nothing automatically.
 *
 * Required "Send List" headers:
 * Email | Name | Table | Guest Count | Attendance | Send
 *
 * Required "Send Log" headers:
 * Timestamp | Email | Name | Table | Status | Error | Version
 */

const REMINDER_SEND_TAB = 'Send List';
const REMINDER_LOG_TAB = 'Send Log';
const REMINDER_MAILER_VERSION = '2026-11-18-v1';

// Safety lock. Keep false until the test email has been reviewed and approved.
const REMINDER_BATCH_ENABLED = false;

// Required before running sendReminderTest().
const REMINDER_TEST_EMAIL = '';

// Limits one manual run even when more rows are marked YES.
const REMINDER_MAX_PER_RUN = 50;

const REMINDER_WEBSITE_URL = 'https://minchianddavid.github.io/celebratewithus/';
const REMINDER_MAPS_URL = 'https://maps.google.com/?q=Taipei+Marriott+Hotel';

// These become available after the reviewed JPGs are committed and published.
const REMINDER_CARD_URLS = {
  home: REMINDER_WEBSITE_URL + 'output/email-reminder/01-home-page-wedding-reminder.jpg',
  weddingDay: REMINDER_WEBSITE_URL + 'output/email-reminder/02-wedding-day-photo-card.jpg',
  schedule: REMINDER_WEBSITE_URL + 'output/email-reminder/03-wedding-schedule-and-venue.jpg',
  arrival: REMINDER_WEBSITE_URL + 'output/email-reminder/04-arrival-guide-36f.jpg',
};

const REQUIRED_SEND_HEADERS = [
  'Email',
  'Name',
  'Table',
  'Guest Count',
  'Attendance',
  'Send',
];

const REQUIRED_LOG_HEADERS = [
  'Timestamp',
  'Email',
  'Name',
  'Table',
  'Status',
  'Error',
  'Version',
];

/**
 * Sends one preview to REMINDER_TEST_EMAIL using the first eligible row.
 * Test sends are logged as TEST and do not mark the guest as sent.
 */
function sendReminderTest() {
  const testEmail = normalizeEmail_(REMINDER_TEST_EMAIL);
  if (!isValidEmail_(testEmail)) {
    throw new Error('Set REMINDER_TEST_EMAIL to a valid address before testing.');
  }

  const context = getReminderContext_();
  const candidates = getEligibleReminderRows_(context.sendSheet, new Set());
  if (!candidates.length) {
    throw new Error('No eligible row found. Mark one Send List row as Attendance=yes and Send=YES.');
  }

  const guest = candidates[0];
  try {
    sendReminderEmail_(testEmail, guest, true);
    appendSendLog_(context.logSheet, guest, 'TEST', 'Sent to ' + testEmail);
  } catch (error) {
    appendSendLog_(context.logSheet, guest, 'TEST_ERROR', error.message || String(error));
    throw error;
  }
}

/**
 * Sends reminders manually to eligible guests.
 * Requires REMINDER_BATCH_ENABLED=true and a typed confirmation.
 */
function sendReminderBatch() {
  if (!REMINDER_BATCH_ENABLED) {
    throw new Error('Batch sending is locked. Review a test email, then set REMINDER_BATCH_ENABLED=true.');
  }

  const context = getReminderContext_();
  const alreadySent = getSuccessfullySentEmails_(context.logSheet);
  const candidates = getEligibleReminderRows_(context.sendSheet, alreadySent)
    .slice(0, REMINDER_MAX_PER_RUN);

  if (!candidates.length) {
    throw new Error('No eligible unsent reminder rows were found.');
  }

  const quota = MailApp.getRemainingDailyQuota();
  if (quota < candidates.length) {
    throw new Error('Not enough remaining MailApp quota. Needed: ' + candidates.length + '; available: ' + quota + '.');
  }

  const ui = SpreadsheetApp.getUi();
  const confirmation = ui.prompt(
    'Send wedding reminders',
    'This will send ' + candidates.length + ' individual email(s). Type SEND to continue.',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirmation.getSelectedButton() !== ui.Button.OK || confirmation.getResponseText().trim() !== 'SEND') {
    throw new Error('Send cancelled. No reminder emails were sent.');
  }

  candidates.forEach(function (guest) {
    try {
      sendReminderEmail_(guest.email, guest, false);
      appendSendLog_(context.logSheet, guest, 'SENT', '');
    } catch (error) {
      appendSendLog_(context.logSheet, guest, 'ERROR', error.message || String(error));
    }
  });
}

function getReminderContext_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('Open the separate Wedding Reminder Send List spreadsheet and run this bound script there.');
  }

  const sendSheet = spreadsheet.getSheetByName(REMINDER_SEND_TAB);
  const logSheet = spreadsheet.getSheetByName(REMINDER_LOG_TAB);
  if (!sendSheet) {
    throw new Error('Missing sheet tab: ' + REMINDER_SEND_TAB + '.');
  }
  if (!logSheet) {
    throw new Error('Missing sheet tab: ' + REMINDER_LOG_TAB + '.');
  }

  assertHeaders_(sendSheet, REQUIRED_SEND_HEADERS);
  assertHeaders_(logSheet, REQUIRED_LOG_HEADERS);
  return { sendSheet: sendSheet, logSheet: logSheet };
}

function assertHeaders_(sheet, requiredHeaders) {
  const lastColumn = sheet.getLastColumn();
  if (!lastColumn) {
    throw new Error(sheet.getName() + ' has no header row.');
  }

  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  const missing = requiredHeaders.filter(function (header) {
    return headers.indexOf(header) === -1;
  });
  if (missing.length) {
    throw new Error(sheet.getName() + ' is missing header(s): ' + missing.join(', '));
  }
}

function getEligibleReminderRows_(sheet, alreadySent) {
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) {
    return [];
  }

  const headers = values.shift();
  const newestByEmail = new Map();

  values.forEach(function (row, index) {
    const data = headers.reduce(function (record, header, columnIndex) {
      record[header] = String(row[columnIndex] || '').trim();
      return record;
    }, {});

    const email = normalizeEmail_(data.Email);
    if (!isValidEmail_(email)) {
      return;
    }

    // If a duplicate address exists, the lower (later) row wins.
    newestByEmail.set(email, {
      rowNumber: index + 2,
      email: email,
      name: data.Name,
      table: data.Table,
      guestCount: data['Guest Count'],
      attendance: data.Attendance.toLowerCase(),
      send: data.Send.toUpperCase(),
    });
  });

  return Array.from(newestByEmail.values()).filter(function (guest) {
    return guest.attendance === 'yes'
      && guest.send === 'YES'
      && guest.name
      && isUsableTable_(guest.table)
      && !alreadySent.has(guest.email);
  });
}

function isUsableTable_(value) {
  const table = String(value || '').trim();
  return Boolean(table) && table.toUpperCase() !== 'NOT FOUND';
}

function getSuccessfullySentEmails_(logSheet) {
  const values = logSheet.getDataRange().getDisplayValues();
  if (values.length < 2) {
    return new Set();
  }

  const headers = values.shift();
  const emailIndex = headers.indexOf('Email');
  const statusIndex = headers.indexOf('Status');
  const sent = new Set();

  values.forEach(function (row) {
    if (String(row[statusIndex]).trim().toUpperCase() === 'SENT') {
      const email = normalizeEmail_(row[emailIndex]);
      if (email) {
        sent.add(email);
      }
    }
  });
  return sent;
}

function appendSendLog_(logSheet, guest, status, errorMessage) {
  logSheet.appendRow([
    new Date(),
    guest.email,
    guest.name,
    guest.table,
    status,
    errorMessage || '',
    REMINDER_MAILER_VERSION,
  ]);
}

function sendReminderEmail_(recipient, guest, isTest) {
  const subjectPrefix = isTest ? '[TEST] ' : '';
  const subject = subjectPrefix + '婚禮行前提醒｜Minchi & David｜2026.11.29';

  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    body: buildReminderPlainText_(guest),
    htmlBody: buildReminderHtml_(guest),
    name: 'Minchi & David',
  });
}

function buildReminderPlainText_(guest) {
  const tableLabel = formatTableLabel_(guest.table);
  return [
    'Hi ' + guest.name + ',',
    '',
    '婚禮即將到來，期待在 11 月 29 日與你見面。',
    '',
    '你的桌次：' + tableLabel,
    '賓客報到：11:30 AM',
    '宴客開始：12:00 PM',
    '',
    '台北萬豪酒店 36F 寰宇廳',
    '台北市中山區樂群二路 199 號',
    '',
    '由飯店正門進入後，從 Concierge 櫃檯左側通道直走到底，',
    '即可搭乘前往 36 樓的電梯。',
    '',
    'Wedding website: ' + REMINDER_WEBSITE_URL,
    'Google Maps: ' + REMINDER_MAPS_URL,
    '',
    'Minchi & David',
    '旻淇・大為',
  ].join('\n');
}

function buildReminderHtml_(guest) {
  const safeName = escapeHtml_(guest.name);
  const safeTable = escapeHtml_(formatTableLabel_(guest.table));
  const cards = [
    { url: REMINDER_CARD_URLS.home, alt: 'Minchi and David wedding reminder' },
    { url: REMINDER_CARD_URLS.weddingDay, alt: 'Wedding Day' },
    { url: REMINDER_CARD_URLS.schedule, alt: 'Wedding schedule and venue' },
    { url: REMINDER_CARD_URLS.arrival, alt: 'Directions to the 36th floor' },
  ].map(function (card) {
    return '<tr><td style="padding:0 0 14px;">'
      + '<img src="' + card.url + '" width="520" alt="' + card.alt + '" '
      + 'style="display:block;width:100%;max-width:520px;height:auto;border:0;">'
      + '</td></tr>';
  }).join('');

  return [
    '<!doctype html>',
    '<html lang="zh-Hant">',
    '<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"></head>',
    '<body style="margin:0;padding:0;background:#f4f1eb;">',
    '<div style="margin:0;padding:34px 12px;background:#f4f1eb;color:#37352f;font-family:Arial,\'PingFang TC\',\'Noto Sans TC\',sans-serif;">',
    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="width:100%;border-collapse:collapse;">',
    '<tr><td align="center">',
    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="width:100%;max-width:560px;border-collapse:separate;background:#ffffff;">',
    '<tr><td style="padding:46px 32px 34px;">',
    '<p style="margin:0 0 18px;font-size:16px;line-height:1.8;color:#56514b;">Hi ' + safeName + ',</p>',
    '<p style="margin:0 0 30px;font-size:15px;line-height:1.9;color:#37352f;">婚禮即將到來，<br>期待在 11 月 29 日與你見面。</p>',
    '<div style="margin:0 0 34px;padding:24px 20px;border-top:1px solid #d9c7aa;border-bottom:1px solid #d9c7aa;text-align:center;">',
    '<div style="margin:0 0 7px;font-size:12px;letter-spacing:.14em;color:#8f7049;">YOUR TABLE · 您的桌次</div>',
    '<div style="font-family:Georgia,\'Times New Roman\',serif;font-size:32px;line-height:1.25;color:#37352f;">' + safeTable + '</div>',
    '</div>',
    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="width:100%;border-collapse:collapse;">',
    cards,
    '</table>',
    '<p style="margin:18px 0 8px;text-align:center;"><a href="' + REMINDER_WEBSITE_URL + '" style="color:#8f7049;font-size:14px;text-decoration:none;">Open the Invitation &#8599;</a></p>',
    '<p style="margin:0;text-align:center;"><a href="' + REMINDER_MAPS_URL + '" style="color:#716a62;font-size:13px;text-decoration:none;">Google Maps &#8599;</a></p>',
    '<div style="width:72px;margin:34px auto 22px;border-top:1px solid #d8d2c8;"></div>',
    '<p style="margin:0;text-align:center;font-family:Georgia,\'Times New Roman\',serif;font-size:20px;font-style:italic;color:#37352f;">Minchi &amp; David</p>',
    '<p style="margin:7px 0 0;text-align:center;font-size:12px;letter-spacing:.14em;color:#716a62;">旻淇・大為</p>',
    '</td></tr></table>',
    '</td></tr></table>',
    '</div>',
    '</body></html>',
  ].join('');
}

function formatTableLabel_(value) {
  const table = String(value || '').trim();
  if (!table) {
    return '';
  }
  if (table.toUpperCase() === 'MAIN TABLE') {
    return '主桌';
  }
  if (/桌$/.test(table)) {
    return table;
  }
  return '第 ' + table + ' 桌';
}

function normalizeEmail_(value) {
  return String(value || '').trim().toLowerCase();
}

function isValidEmail_(value) {
  return value.length <= 254 && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
}

function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
