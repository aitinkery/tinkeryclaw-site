/**
 * AI Tinkery -- Network Interest Form: auto-welcome (v4, newsletter-matched banner)
 *
 * Matches tinkery-newsletter-preview.pages.dev design system:
 *   - dark #0e0e10 header, cream mascot + "AI TINKERY" wordmark
 *   - 2px #ff6b35 orange rule
 *   - cream #f4f1ea page bg, Georgia serif body
 *   - orange #d94e1f CTA
 *
 * The newsletter CTA always points at NEWSLETTER_URL (stable root).
 * To update the welcome copy later, edit TEMPLATE_HTML below -- no reinstall.
 *
 * INSTALL (one-time)
 *   1. Open https://docs.google.com/spreadsheets/d/16P7AXGmR53n5ti_xBe9jhpBhcBbKpuLAFyOOj9k6BhI/edit
 *   2. Extensions -> Apps Script.
 *   3. Paste this whole file into Code.gs (replace anything there).
 *   4. Save. Run installTrigger() once -- grant permissions when prompted
 *      (MailApp + SpreadsheetApp scopes).
 *   5. Submit the form yourself as a test. Check:
 *        - Sent At + Send Log columns populate on the new row
 *        - DRY_RUN inbox (wilsong@) receives the "[DRY-RUN -> you@...]" email
 *   6. Flip DRY_RUN = false. Done.
 */

// ============================================================================
// CONFIG
// ============================================================================

const DRY_RUN = true;
const DRY_RUN_TO = 'wilsong@stanford.edu';

const FROM_NAME = 'AI Tinkery';
const REPLY_TO  = 'wilsong@stanford.edu';
const BCC       = '';

const NEWSLETTER_SUBJECT = 'Welcome to the Tinkery Network — here’s the latest newsletter';

const NEWSLETTER_URL          = 'https://tinkery-newsletter-preview.pages.dev/';
const PREVIOUS_NEWSLETTER_URL = 'https://tinkery-newsletter-preview.pages.dev/previous/';

const COL_SENT_AT   = 'Sent At';
const COL_SEND_LOG  = 'Send Log';
const COL_TEMPLATE  = 'Welcome Template';
const TEMPLATE_NAME = 'newsletter-v2';


// ============================================================================
// PUBLIC ENTRY POINTS
// ============================================================================

function onFormSubmit(e) {
  const sheet = e && e.range ? e.range.getSheet() : SpreadsheetApp.getActiveSheet();
  const rowIdx = e && e.range ? e.range.getRow() : sheet.getLastRow();
  processRow_(sheet, rowIdx, false);
}

function installTrigger() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onFormSubmit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
  ensureAuxColumns_(ss.getSheets()[0]);
  Logger.log('Trigger installed. Sheet aux columns ensured.');
}

function resendRow(rowIdx) {
  const sheet = SpreadsheetApp.getActive().getSheets()[0];
  processRow_(sheet, rowIdx, true);
}

function testWelcomeEmail() {
  const html = TEMPLATE_HTML.replace(/\{\{first_name\}\}/g, 'Test');
  const plain = buildPlainEmail_('Test');
  MailApp.sendEmail({
    to: Session.getActiveUser().getEmail(),
    subject: '[TEST] ' + NEWSLETTER_SUBJECT,
    name: FROM_NAME,
    replyTo: REPLY_TO,
    body: plain,
    htmlBody: html,
    inlineImages: {
      'tinkery-mascot': getLogoBlob_()
    }
  });
}


// ============================================================================
// CORE
// ============================================================================

function processRow_(sheet, rowIdx, forceSend) {
  if (rowIdx < 2) return;
  ensureAuxColumns_(sheet);
  const cols = headerMap_(sheet);
  const width = sheet.getLastColumn();
  const row = sheet.getRange(rowIdx, 1, 1, width).getValues()[0];

  const alreadySent = row[cols[COL_SENT_AT] - 1];
  if (alreadySent && !forceSend) {
    Logger.log('Row ' + rowIdx + ' already sent at ' + alreadySent + '. Skip.');
    return;
  }

  const email = firstNonEmpty_([
    row[(cols['Email Address'] || 0) - 1],
    row[(cols['Email Address (2)'] || 0) - 1],
    row[(cols['Email'] || 0) - 1]
  ]);
  const fullName = String(row[(cols['Full Name'] || cols['Name'] || 0) - 1] || '').trim();
  const firstName = fullName.split(/[\s,]+/)[0] || 'Tinkerer';

  if (!email || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) {
    logSend_(sheet, rowIdx, cols, 'SKIP: no valid email');
    return;
  }

  const html = TEMPLATE_HTML.replace(/\{\{first_name\}\}/g, escapeHtml_(firstName));
  const plain = buildPlainEmail_(firstName);
  const toAddr = DRY_RUN ? DRY_RUN_TO : email;
  const subject = DRY_RUN ? '[DRY-RUN → ' + email + '] ' + NEWSLETTER_SUBJECT : NEWSLETTER_SUBJECT;

  try {
    const opts = { name: FROM_NAME, replyTo: REPLY_TO, htmlBody: html, inlineImages: { 'tinkery-mascot': getLogoBlob_() } };
    if (BCC) opts.bcc = BCC;
    MailApp.sendEmail(toAddr, subject, plain, opts);
    logSend_(sheet, rowIdx, cols, 'OK' + (DRY_RUN ? ' (dry-run)' : ''));
  } catch (err) {
    logSend_(sheet, rowIdx, cols, 'ERR: ' + err.message);
    throw err;
  }
}


// ============================================================================
// EMAIL BODY
// ============================================================================

function buildPlainEmail_(firstName) {
  return '' +
    'THE TINKERY NETWORK\n' +
    'AI Tinkery · Stanford\n\n' +
    '---\n\n' +
    'Welcome, ' + firstName + '!\n\n' +
    'Thanks for joining the Tinkery Network. Below is the latest newsletter so ' +
    'you’re caught up on what’s happening at the Tinkery and across the network. ' +
    'Reply to this email if anything in it sparks a question.\n\n' +
    'READ THE CURRENT NEWSLETTER:\n  ' + NEWSLETTER_URL + '\n\n' +
    'Previous issue:\n  ' + PREVIOUS_NEWSLETTER_URL + '\n\n' +
    '---\n\n' +
    'See you around,\n' +
    'Gregory & Karin\n' +
    'AI Tinkery · Stanford Accelerator for Learning\n\n' +
    'You’re receiving this because you signed up for the Tinkery Network at ' +
    'https://aitinkery.com/network. Reply to this email to reach us directly.';
}

const TEMPLATE_HTML = `<!doctype html><html><body style="margin:0;padding:0;background:#f4f1ea;font-family:Georgia,'Iowan Old Style','Palatino Linotype',Palatino,serif;color:#1a1a1a;">
<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="background:#f4f1ea;padding:24px 12px 40px;">
<tr><td align="center">
<table role="presentation" width="640" cellspacing="0" cellpadding="0" border="0" style="max-width:640px;width:100%;background:#ffffff;border-radius:12px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.04);">

<!-- Brand header: dark bg, mascot + wordmark -->
<tr><td style="background:#0e0e10;padding:28px 32px;">
  <table role="presentation" cellspacing="0" cellpadding="0" border="0"><tr>
    <td style="padding-right:14px;vertical-align:middle;">
      <img src="cid:tinkery-mascot" alt="AI Tinkery" width="56" height="56" style="display:block;width:56px;height:56px;">
    </td>
    <td style="vertical-align:middle;">
      <div style="font-family:Impact,'Arial Black','Helvetica Neue',Arial,sans-serif;font-size:22px;line-height:1;letter-spacing:0.02em;color:#f4f1ea;font-weight:900;">AI TINKERY</div>
    </td>
  </tr></table>
</td></tr>

<!-- Orange rule -->
<tr><td style="background:#ff6b35;height:2px;line-height:2px;font-size:0;">&nbsp;</td></tr>

<!-- Body -->
<tr><td style="padding:40px 44px;">
  <div style="font-family:-apple-system,BlinkMacSystemFont,'Helvetica Neue',Arial,sans-serif;font-size:11px;letter-spacing:0.14em;text-transform:uppercase;color:#6a6a6a;margin:0 0 18px;">The Tinkery Network</div>

  <h1 style="font-family:Georgia,serif;font-weight:700;font-size:30px;line-height:1.2;margin:0 0 10px;color:#1a1a1a;">Welcome, {{first_name}}!</h1>

  <p style="font-style:italic;color:#6a6a6a;margin:0 0 28px;font-size:15px;">You&rsquo;re in. Here&rsquo;s the latest from the Tinkery.</p>

  <p style="font-size:17px;line-height:1.55;margin:0 0 16px;">Thanks for joining the Tinkery Network. Below is the latest newsletter so you&rsquo;re caught up on what&rsquo;s happening at the Tinkery and across the network.</p>

  <p style="font-size:17px;line-height:1.55;margin:0 0 28px;">Click through to read it &mdash; and reply to this email if anything in it sparks a question.</p>

  <!-- CTA -->
  <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="margin:8px 0 24px 0;"><tr><td style="background:#d94e1f;border-radius:6px;">
    <a href="https://tinkery-newsletter-preview.pages.dev/" style="display:inline-block;padding:14px 28px;color:#ffffff;text-decoration:none;font-weight:700;font-size:17px;font-family:-apple-system,BlinkMacSystemFont,'Helvetica Neue',Arial,sans-serif;">Read the current newsletter &rarr;</a>
  </td></tr></table>

  <p style="font-size:15px;line-height:1.55;margin:0 0 28px;color:#6a6a6a;">Missed the one before that? <a href="https://tinkery-newsletter-preview.pages.dev/previous/" style="color:#d94e1f;text-decoration:underline;">Previous issue &rarr;</a></p>

  <hr style="border:none;border-top:1px solid #eaeaea;margin:28px 0 20px 0;">

  <p style="font-family:Georgia,serif;font-size:16px;line-height:1.5;margin:0 0 4px 0;color:#1a1a1a;">See you around,</p>
  <p style="font-family:Georgia,serif;font-size:16px;line-height:1.5;margin:0;color:#1a1a1a;"><strong>Gregory &amp; Karin</strong><br>
  <span style="color:#6a6a6a;">AI Tinkery &middot; Stanford Accelerator for Learning</span></p>
</td></tr>

<!-- Footer -->
<tr><td style="padding:20px 44px 28px 44px;background:#faf3e2;border-top:1px solid #eaeaea;font-family:-apple-system,BlinkMacSystemFont,'Helvetica Neue',Arial,sans-serif;font-size:13px;line-height:1.5;color:#6a6a6a;">
You&rsquo;re receiving this because you signed up for the Tinkery Network at <a href="https://aitinkery.com/network" style="color:#d94e1f;text-decoration:none;">aitinkery.com/network</a>. Reply to this email to reach us directly.
</td></tr>

</table>
</td></tr>
</table>
</body></html>`;


// ============================================================================
// UTIL
// ============================================================================

function headerMap_(sheet) {
  const hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  hdr.forEach((h, i) => {
    const name = String(h || '').trim();
    if (!name) return;
    if (map[name]) {
      let n = 2;
      while (map[name + ' (' + n + ')']) n++;
      map[name + ' (' + n + ')'] = i + 1;
    } else {
      map[name] = i + 1;
    }
  });
  return map;
}

function ensureAuxColumns_(sheet) {
  const cols = headerMap_(sheet);
  const need = [COL_SENT_AT, COL_SEND_LOG, COL_TEMPLATE];
  let nextCol = sheet.getLastColumn() + 1;
  need.forEach(name => {
    if (!cols[name]) {
      sheet.getRange(1, nextCol).setValue(name);
      nextCol++;
    }
  });
}

function firstNonEmpty_(arr) {
  for (const v of arr) {
    if (v && String(v).trim()) return String(v).trim();
  }
  return '';
}

function logSend_(sheet, rowIdx, cols, msg) {
  const now = new Date();
  if (msg.indexOf('OK') === 0) {
    sheet.getRange(rowIdx, cols[COL_SENT_AT]).setValue(now);
    sheet.getRange(rowIdx, cols[COL_TEMPLATE]).setValue(TEMPLATE_NAME);
  }
  sheet.getRange(rowIdx, cols[COL_SEND_LOG]).setValue(msg);
  Logger.log('Row ' + rowIdx + ': ' + msg);
}

function escapeHtml_(s) {
  return String(s).replace(/[&<>"']/g, c => ({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
  }[c]));
}

const LOGO_BASE64 = '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';

function getLogoBlob_() {
  return Utilities.newBlob(
    Utilities.base64Decode(LOGO_BASE64),
    'image/png',
    'tinkery-mascot.png'
  );
}
