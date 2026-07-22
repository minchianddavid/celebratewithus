/**
 * Append-only Google Apps Script endpoint for wedding RSVPs.
 *
 * This must be created from Extensions > Apps Script inside the RSVP
 * spreadsheet. It only accesses that attached spreadsheet, only appends
 * submissions to the "RSVP" tab, and may send one confirmation email to the
 * validated address included in the current submission.
 *
 * RSVP columns:
 * Timestamp | Name | Attendance | Guest Count | Invitation Type | Email |
 * Address | Message
 */

const RSVP_TAB = 'RSVP';
const RSVP_ENDPOINT_VERSION = '2026-07-22-email-v53';
const EMAIL_MONOGRAM_PNG_BASE64 = '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';
const EMAIL_SIGNATURE_PNG_BASE64 = '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';

const EMAIL_SIGNATURE_ZH_PNG_BASE64 = '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';

function normalizeEmail(value) {
  return String(value || '').trim().toLowerCase();
}

function isValidEmail(value) {
  return value.length <= 254 && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function sendRsvpConfirmation(email, name) {
  const safeName = escapeHtml(name);
  const subject = 'RSVP Confirmed｜婚禮邀請函｜Minchi & David 蔡旻淇・郭大為｜2026.11.29';
  const websiteUrl = 'https://minchianddavid.github.io/celebratewithus/';
  const mapsUrl = 'https://maps.google.com/?q=Taipei+Marriott+Hotel';
  const calendarUrl = websiteUrl + 'assets/calendar/minchi-david-wedding.ics';
  const inlineImages = {
    rsvpMonogram: Utilities.newBlob(Utilities.base64Decode(EMAIL_MONOGRAM_PNG_BASE64), 'image/png', 'md-monogram.png'),
    rsvpSignature: Utilities.newBlob(Utilities.base64Decode(EMAIL_SIGNATURE_PNG_BASE64), 'image/png', 'minchi-and-david.png'),
    rsvpSignatureZh: Utilities.newBlob(Utilities.base64Decode(EMAIL_SIGNATURE_ZH_PNG_BASE64), 'image/png', 'minchi-david-zh.png'),
  };
  const body = [
    'M·D',
    '',
    'RSVP CONFIRMED',
    '',
    'Hi ' + name + ',',
    '',
    '太棒了！',
    '期待與你一起分享這個特別的日子 🤍',
    '',
    '婚禮前三天，我們會再寄送一封 Email，',
    '通知你的桌次資訊與行前提醒。',
    '',
    '2026.11.29 • Taipei Marriott Hotel',
    'Add to calendar: ' + calendarUrl,
    mapsUrl,
    '',
    'Open the Invitation ↗ ' + websiteUrl,
    '',
    '────────────',
    '',
    'Minchi & David',
    '旻淇・大為',
  ].join('\n');
  const htmlBody = [
    '<div lang="zh-Hant" style="margin:0;padding:40px 16px;background:#f4f4f2;color:#37352f;font-family:Arial,\'PingFang TC\',\'Noto Sans TC\',sans-serif;">',
    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="width:100%;border-collapse:collapse;">',
    '<tr><td align="center">',
    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="width:100%;max-width:520px;border:1px solid #e6e3dd;border-collapse:separate;background:#ffffff;">',
    '<tr><td style="padding:54px 48px 50px;">',
    '<div style="margin:0 0 8px;text-align:center;"><img src="cid:rsvpMonogram" width="82" alt="M·D" style="display:inline-block;width:82px;max-width:100%;height:auto;border:0;"></div>',
    '<div style="margin:0 0 46px;text-align:center;font-family:\'Avenir Next\',Avenir,\'Helvetica Neue\',Arial,sans-serif;font-size:13px;font-weight:600;line-height:1.5;letter-spacing:.14em;text-transform:uppercase;color:#716b63;">RSVP Confirmed</div>',
    '<p style="margin:0 0 24px;font-family:\'Avenir Next\',Avenir,\'Helvetica Neue\',Arial,sans-serif;font-size:16px;font-weight:400;line-height:1.8;color:#56514b;">Hi ' + safeName + ',</p>',
    '<p style="margin:0 0 26px;font-family:\'PingFang TC\',\'Noto Sans TC\',sans-serif;font-size:15px;font-weight:400;line-height:1.8;color:#37352f;">太棒了！<br>期待與你一起分享這個特別的日子 🤍</p>',
    '<p style="margin:0 0 52px;font-family:\'PingFang TC\',\'Noto Sans TC\',sans-serif;font-size:15px;font-weight:400;line-height:1.8;color:#56514b;">婚禮前三天，我們會再寄送一封 Email，<br>通知你的桌次資訊與行前提醒。</p>',
    '<p style="margin:0 0 8px;font-family:\'Avenir Next\',Avenir,\'Helvetica Neue\',Arial,sans-serif;font-size:14px;font-weight:400;line-height:1.6;color:#56514b;"><a href="' + calendarUrl + '" target="_blank" style="color:#56514b;font-weight:400;text-decoration:none;">2026.11.29</a> <span style="font-size:9px;vertical-align:1px;">•</span> <a href="' + mapsUrl + '" target="_blank" style="color:#56514b;font-weight:400;text-decoration:none;">Taipei Marriott Hotel</a></p>',
    '<p style="margin:0 0 24px;"><a href="' + websiteUrl + '" target="_blank" style="color:#8f7049;font-family:\'Avenir Next\',Avenir,\'Helvetica Neue\',Arial,sans-serif;font-size:14px;font-weight:400;line-height:1.6;letter-spacing:.01em;text-decoration:none;">Open the Invitation ↗</a></p>',
    '<div style="width:84px;margin:0 0 16px;border-top:1px solid #d8d2c8;"></div>',
    '<p style="margin:0 0 8px;text-align:left;"><img src="cid:rsvpSignature" width="180" alt="Minchi &amp; David" style="display:block;width:180px;max-width:100%;height:auto;border:0;"></p>',
    '<p style="margin:0 0 0 8px;text-align:left;"><img src="cid:rsvpSignatureZh" width="89" alt="旻淇・大為" style="display:block;width:89px;max-width:100%;height:auto;border:0;"></p>',
    '</td></tr></table>',
    '</td></tr></table>',
    '</div>',
  ].join('');

  MailApp.sendEmail({
    to: email,
    subject: subject,
    body: body,
    htmlBody: htmlBody,
    inlineImages: inlineImages,
    name: 'Minchi & David',
  });
}

function getRsvpSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    return {
      error: 'RSVP endpoint is not attached to a spreadsheet. Open the RSVP spreadsheet, choose Extensions > Apps Script, and redeploy the web app from there.',
      sheet: null,
    };
  }

  const sheet = spreadsheet.getSheetByName(RSVP_TAB);
  if (!sheet) {
    return {
      error: 'RSVP sheet not found. Create a sheet tab named "RSVP" and redeploy the web app.',
      sheet: null,
    };
  }

  return { error: '', sheet: sheet };
}

function doPost(e) {
  try {
    const params = e.parameter;
    const name = (params.name || '').trim();
    const email = normalizeEmail(params.email);
    const destination = getRsvpSheet();

    if (!destination.sheet) {
      return jsonResponse({
        result: 'error',
        message: destination.error,
      });
    }

    const sheet = destination.sheet;
    const attendance = params.attendance || 'yes';

    if (!name) {
      return jsonResponse({
        result: 'error',
        message: 'Name is required.',
      });
    }

    if (email && !isValidEmail(email)) {
      return jsonResponse({
        result: 'error',
        message: 'Please enter a valid email address.',
      });
    }

    sheet.appendRow([
      new Date(),
      name,
      attendance,
      parseInt(params.guest_count, 10) || 1,
      (params.invitation_type || '').trim(),
      email,
      (params.address || '').trim(),
      (params.message || '').trim(),
    ]);

    let emailSent = false;
    let emailErrorMessage = '';
    if (email && attendance === 'yes') {
      try {
        sendRsvpConfirmation(email, name);
        emailSent = true;
      } catch (emailError) {
        console.error('RSVP saved but confirmation email failed: ' + emailError.message);
        emailErrorMessage = emailError.message || 'Confirmation email could not be sent. Reauthorize MailApp and redeploy the web app.';
      }
    }

    return jsonResponse({
      result: 'success',
      message: 'RSVP received!',
      email_sent: emailSent,
      email_error: emailErrorMessage,
      endpoint_version: RSVP_ENDPOINT_VERSION,
    });
  } catch (error) {
    return jsonResponse({
      result: 'error',
      message: 'Server error: ' + error.message,
    });
  }
}

function doGet() {
  const destination = getRsvpSheet();

  if (!destination.sheet) {
    return jsonResponse({
      result: 'error',
      message: destination.error,
    });
  }

  return jsonResponse({
    result: 'ok',
    message: 'RSVP endpoint is active.',
    endpoint_version: RSVP_ENDPOINT_VERSION,
  });
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
