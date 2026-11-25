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
 * Email | Name | Table | Guest Count | Send
 *
 * Required "Send Log" headers:
 * Timestamp | Email | Name | Table | Status | Error | Version
 */

const REMINDER_SEND_TAB = 'Send List';
const REMINDER_LOG_TAB = 'Send Log';
const REMINDER_MAILER_VERSION = '2026-11-18-v1';
const REMINDER_TIME_ZONE = 'Asia/Taipei';
const REMINDER_WEDDING_DATE = '2026-11-29';

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
  monogram: REMINDER_WEBSITE_URL + 'output/email-reminder/email-emblem.png?v=20261124-1',
  signature: REMINDER_WEBSITE_URL + 'output/email-reminder/email-signature.png',
  signatureZh: REMINDER_WEBSITE_URL + 'output/email-reminder/email-signature-zh.png',
};

const REQUIRED_SEND_HEADERS = [
  'Email',
  'Name',
  'Table',
  'Guest Count',
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
    throw new Error('No eligible row found. Mark one valid Send List row as Send=YES.');
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
      send: data.Send.toUpperCase(),
    });
  });

  return Array.from(newestByEmail.values()).filter(function (guest) {
    return guest.send === 'YES'
      && guest.name
      && hasTableValue_(guest.table)
      && !alreadySent.has(guest.email);
  });
}

function hasTableValue_(value) {
  const table = String(value || '').trim();
  return Boolean(table);
}

function hasTableAssignment_(value) {
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
  const tableSubject = hasTableAssignment_(guest.table)
    ? '｜' + formatTableLabel_(guest.table)
    : '';
  const subject = subjectPrefix + '婚禮行前提醒' + tableSubject + '｜Minchi & David';
  const sendTime = new Date();

  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    body: buildReminderPlainText_(guest, sendTime),
    htmlBody: buildReminderHtml_(guest, sendTime),
    name: 'Minchi & David',
  });
}

function buildReminderPlainText_(guest, sendTime) {
  const tableLabel = formatTableLabel_(guest.table);
  const guestCountLabel = formatGuestCountLabel_(guest.guestCount);
  const countdownCopy = formatWeddingCountdown_(sendTime);
  const tableLines = hasTableAssignment_(guest.table)
    ? [
      '你的桌次：' + tableLabel,
      guestCountLabel ? '出席人數：' + guestCountLabel : '',
      '',
    ]
    : [];

  return [
    'Hi ' + guest.name + ',',
    '',
    '歡迎來到郭大為 & 蔡旻淇的婚禮！' + countdownCopy + '期待見到你，和你一起分享這份喜悅 🤍',
    '— 郭大為 & 蔡旻淇',
    '',
  ].concat(tableLines, [
    '2026.11.29 · Sunday',
    '賓客報到：11:30 AM',
    '午宴開始：12:00 PM',
    '',
    '台北萬豪酒店 36F 寰宇廳',
    '台北市中山區樂群二路 199 號',
    '',
    '前往 36 樓的電梯位於飯店正門 Concierge 左側。',
    '請放心直走到底，電梯真的在那裡。',
    '',
    '停車資訊：',
    '當日有多場婚宴，車位有限。',
    'B2–B4 設有地下停車場。',
    '接待桌備有 QR Code，離場前掃描即可免費折抵。',
    '',
    'Wedding website: ' + REMINDER_WEBSITE_URL,
    'Google Maps: ' + REMINDER_MAPS_URL,
    '',
    'Minchi & David',
    '旻淇・大為',
  ]).join('\n');
}

function buildReminderHtml_(guest, sendTime) {
  const safeName = escapeHtml_(guest.name);
  const safeTable = escapeHtml_(formatTableLabel_(guest.table));
  const safeGuestCount = escapeHtml_(formatGuestCountLabel_(guest.guestCount));
  const safeCountdownCopy = escapeHtml_(formatWeddingCountdown_(sendTime));
  const guestCountHtml = safeGuestCount
    ? '<div style="margin:7px 0 0;font-size:13px;line-height:1.6;color:#716a62;">' + safeGuestCount + '</div>'
    : '';
  const tableSectionHtml = hasTableAssignment_(guest.table)
    ? '<div style="margin:0 0 40px;padding:23px 16px 22px;border-top:1px solid #d8d2c8;border-bottom:1px solid #d8d2c8;text-align:center;">'
      + '<div style="margin:0 0 8px;font-family:\'Avenir Next\',Avenir,\'Helvetica Neue\',Arial,sans-serif;font-size:11px;font-weight:600;line-height:1.5;letter-spacing:.18em;color:#8f7049;">YOUR TABLE</div>'
      + '<div style="font-family:\'PingFang TC\',\'Noto Sans TC\',sans-serif;font-size:29px;font-weight:500;line-height:1.35;color:#37352f;">' + safeTable + '</div>'
      + guestCountHtml
      + '</div>'
    : '';

  return [
    '<!doctype html>',
    '<html lang="zh-Hant">',
    '<head><meta charset="UTF-8"><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"></head>',
    '<body style="margin:0;padding:0;background:#f4f4f2;">',
    '<div style="display:none;max-height:0;overflow:hidden;opacity:0;color:transparent;">11 月 29 日婚禮桌次、時間與抵達提醒</div>',
    '<div lang="zh-Hant" style="margin:0;padding:40px 16px;background:#f4f4f2;color:#37352f;font-family:Arial,\'PingFang TC\',\'Noto Sans TC\',sans-serif;">',
    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="width:100%;border-collapse:collapse;">',
    '<tr><td align="center">',
    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="width:100%;max-width:520px;border:1px solid #e6e3dd;border-collapse:separate;background:#ffffff;">',
    '<tr><td style="padding:54px 48px 50px;">',
    '<div style="margin:0 0 12px;text-align:center;"><img src="' + REMINDER_CARD_URLS.monogram + '" width="150" alt="Minchi and David wedding emblem" style="display:inline-block;width:150px;max-width:100%;height:auto;border:0;"></div>',
    '<div style="margin:0 0 6px;text-align:center;font-family:\'Avenir Next\',Avenir,\'Helvetica Neue\',Arial,sans-serif;font-size:13px;font-weight:600;line-height:1.5;letter-spacing:.14em;text-transform:uppercase;color:#8f7049;">Wedding Day Reminder</div>',
    '<div style="margin:0 0 46px;text-align:center;font-family:\'PingFang TC\',\'Noto Sans TC\',sans-serif;font-size:12px;font-weight:400;line-height:1.6;letter-spacing:.08em;color:#716a62;">婚禮行前提醒</div>',
    '<p style="margin:0 0 24px;font-family:\'Avenir Next\',Avenir,\'Helvetica Neue\',Arial,sans-serif;font-size:16px;font-weight:400;line-height:1.8;color:#56514b;">Hi ' + safeName + ',</p>',
    '<p style="margin:0 0 42px;font-family:\'PingFang TC\',\'Noto Sans TC\',sans-serif;font-size:15px;font-weight:400;line-height:1.9;color:#37352f;overflow-wrap:anywhere;word-break:break-all;">歡迎來到郭大為 &amp; 蔡旻淇的婚禮！' + safeCountdownCopy + '期待見到你，和你一起分享這份喜悅 &#129293;<span style="display:block;margin-top:8px;color:#716a62;font-family:\'PingFang TC\',\'Noto Sans TC\',sans-serif;font-size:13px;line-height:1.6;">&mdash; 郭大為 &amp; 蔡旻淇</span></p>',
    tableSectionHtml,
    '<div style="margin:0 0 38px;"><img src="' + REMINDER_CARD_URLS.weddingDay + '" width="422" alt="Minchi and David on their wedding day" style="display:block;width:100%;max-width:422px;height:auto;border:0;"></div>',
    '<div style="margin:0 0 42px;">',
    '<div style="margin:0 0 18px;font-family:\'Avenir Next\',Avenir,\'Helvetica Neue\',Arial,sans-serif;font-size:12px;font-weight:600;line-height:1.5;letter-spacing:.14em;color:#8f7049;">2026.11.29 · SUNDAY</div>',
    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="width:100%;border-collapse:collapse;font-family:\'PingFang TC\',\'Noto Sans TC\',sans-serif;color:#37352f;">',
    '<tr><td style="width:72px;padding:0 0 9px;font-size:14px;line-height:1.7;color:#716a62;">11:30</td><td style="padding:0 0 9px;font-size:14px;line-height:1.7;">賓客報到</td></tr>',
    '<tr><td style="width:72px;padding:0 0 21px;font-size:14px;line-height:1.7;color:#716a62;">12:00</td><td style="padding:0 0 21px;font-size:14px;line-height:1.7;">午宴開始</td></tr>',
    '</table>',
    '<p style="margin:0 0 5px;font-family:\'PingFang TC\',\'Noto Sans TC\',sans-serif;font-size:15px;font-weight:500;line-height:1.7;color:#37352f;">台北萬豪酒店 · 36F 寰宇廳</p>',
    '<p style="margin:0 0 12px;font-family:\'PingFang TC\',\'Noto Sans TC\',sans-serif;font-size:13px;font-weight:400;line-height:1.7;color:#716a62;">台北市中山區樂群二路 199 號</p>',
    '<p style="margin:0;"><a href="' + REMINDER_MAPS_URL + '" target="_blank" style="color:#8f7049;font-family:\'Avenir Next\',Avenir,\'Helvetica Neue\',Arial,sans-serif;font-size:14px;font-weight:400;line-height:1.6;text-decoration:none;">Open in Google Maps &#8599;</a></p>',
    '</div>',
    '<div style="margin:0 0 24px;padding-top:31px;border-top:1px solid #d8d2c8;">',
    '<div style="margin:0 0 13px;font-family:\'Avenir Next\',Avenir,\'Helvetica Neue\',Arial,sans-serif;font-size:11px;font-weight:600;line-height:1.5;letter-spacing:.16em;color:#8f7049;">ARRIVAL NOTE</div>',
    '<p style="margin:0;font-family:\'PingFang TC\',\'Noto Sans TC\',sans-serif;font-size:14px;font-weight:400;line-height:1.9;color:#56514b;">前往 36 樓的電梯位於<br>飯店正門 Concierge 左側，<br>請放心直走到底，<br>電梯真的在那裡。</p>',
    '</div>',
    '<div style="margin:0 0 31px;padding-top:29px;border-top:1px solid #d8d2c8;">',
    '<div style="margin:0 0 13px;font-family:\'Avenir Next\',Avenir,\'Helvetica Neue\',Arial,sans-serif;font-size:11px;font-weight:600;line-height:1.5;letter-spacing:.16em;color:#8f7049;">PARKING</div>',
    '<p style="margin:0;font-family:\'PingFang TC\',\'Noto Sans TC\',sans-serif;font-size:14px;font-weight:400;line-height:1.9;color:#56514b;">當日有多場婚宴，車位有限。<br>B2–B4 設有地下停車場。<br>接待桌備有 QR Code，<br>離場前掃描即可免費折抵。</p>',
    '</div>',
    '<p style="margin:0 0 24px;"><a href="' + REMINDER_WEBSITE_URL + '" target="_blank" style="color:#8f7049;font-family:\'Avenir Next\',Avenir,\'Helvetica Neue\',Arial,sans-serif;font-size:14px;font-weight:400;line-height:1.6;letter-spacing:.01em;text-decoration:none;">Open the Invitation &#8599;</a></p>',
    '<div style="width:84px;margin:0 0 16px;border-top:1px solid #d8d2c8;"></div>',
    '<p style="margin:0 0 5px;text-align:left;"><img src="' + REMINDER_CARD_URLS.signature + '" width="205" alt="Minchi &amp; David" style="display:block;width:205px;max-width:100%;height:auto;border:0;"></p>',
    '<p style="margin:0 0 0 8px;text-align:left;"><img src="' + REMINDER_CARD_URLS.signatureZh + '" width="89" alt="旻淇・大為" style="display:block;width:89px;max-width:100%;height:auto;border:0;"></p>',
    '</td></tr></table>',
    '</td></tr></table>',
    '</div>',
    '</body></html>',
  ].join('');
}

function formatWeddingCountdown_(sendTime) {
  const daysUntilWedding = getDaysUntilWedding_(sendTime || new Date());
  if (daysUntilWedding === 0) {
    return '婚禮就是今天！';
  }
  return '距離婚禮只剩 ' + daysUntilWedding + ' 天了，';
}

function getDaysUntilWedding_(sendTime) {
  const localDate = Utilities.formatDate(sendTime, REMINDER_TIME_ZONE, 'yyyy-MM-dd');
  const localParts = localDate.split('-').map(Number);
  const weddingParts = REMINDER_WEDDING_DATE.split('-').map(Number);
  const localUtc = Date.UTC(localParts[0], localParts[1] - 1, localParts[2]);
  const weddingUtc = Date.UTC(weddingParts[0], weddingParts[1] - 1, weddingParts[2]);
  return Math.max(0, Math.round((weddingUtc - localUtc) / 86400000));
}

function formatGuestCountLabel_(value) {
  const guestCount = String(value || '').trim();
  if (!guestCount) {
    return '';
  }
  if (/^\d+$/.test(guestCount)) {
    return guestCount + ' 位賓客';
  }
  return guestCount;
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
