// ============================================================
// GOOGLE APPS SCRIPT - metrocity-hoalac.com Lead Collection
// Chi gui 1 email thong bao cho moi lead.
// ============================================================
// CAI DAT:
// 1. Mo Google Sheet > Extensions > Apps Script.
// 2. Xoa code cu, dan toan bo code nay vao.
// 3. Kiem tra SHEET_ID, SHEET_NAME va NOTIFY_EMAILS.
// 4. Save > Deploy > New deployment.
//    Type: Web app
//    Execute as: Me
//    Who has access: Anyone
// 5. Copy URL Web app va dan vao bien SCRIPT_URL trong index.html.
//
// LUU Y:
// - Khong chay setupInstallableTriggers() cho ban code nay.
// - Neu code cu da tung tao trigger onSheetEdit/onSheetChange,
//   chay removeOldSheetTriggers() mot lan de xoa trigger cu.
// ============================================================

// ==================== CAU HINH ====================
const SHEET_ID = '10c6qLnV1q46wwtGxCEGGKRRjQSH0o6sdMHGMiHwCyzM';
const SHEET_NAME = 'Sheet1';
const SHEET_URL = 'https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit';

// Chi gui email toi cac dia chi trong danh sach nay.
// Khong tu dong lay email Owner/Editor/Viewer cua Google Sheet.
const NOTIFY_EMAILS = [
  'phamvuduchuynd@gmail.com'
];

const BRAND = {
  projectName: 'Lead Dat nen Hoa Lac',
  tagline: 'metrocity-hoalac.com',
  primary: '#09353F',
  primaryDark: '#061F25',
  bg: '#F4F8F9',
  card: '#FFFFFF',
  text: '#0F172A',
  subtext: '#475569',
  hotline: '0868868686',
  website: 'https://metrocity-hoalac.com'
};

const MIN_FILL_MS = 3000;
const COOLDOWN_MS = 30000;
const DUP_WINDOW_SECONDS = 600;

// ==================== WEB APP ====================
function doGet(e) {
  const params = getParams(e);
  const isSubmit = String(params.action || '').toLowerCase() === 'submit' ||
    !!(params.name || params.phone || params.finance);

  if (!isSubmit) return jsonResponse({ ok: true, status: 'alive' });
  return handleSubmission(params);
}

function doPost(e) {
  return handleSubmission(getParams(e));
}

function handleSubmission(params) {
  try {
    const now = Date.now();

    if (String(params.website || '').trim() !== '') {
      return jsonResponse({ ok: true, status: 'blocked_honeypot' });
    }

    const startedAt = Number(params.form_started_at || 0);
    if (!Number.isFinite(startedAt) || startedAt <= 0 || now - startedAt < MIN_FILL_MS) {
      return jsonResponse({ ok: true, status: 'blocked_too_fast' });
    }

    const name = String(params.name || '').trim();
    const finance = String(params.finance || '').trim();
    const phone = normalizePhone(firstNonEmpty([
      params.phone,
      params.phone_text,
      params.sdt
    ]));

    if (!isValidVNPhone(phone)) {
      return jsonResponse({ ok: true, status: 'blocked_bad_phone' });
    }

    const cache = CacheService.getScriptCache();
    const cooldownKey = 'lead-cooldown:' + hashKey(phone);
    const lastSubmitted = Number(cache.get(cooldownKey) || 0);

    if (lastSubmitted && now - lastSubmitted < COOLDOWN_MS) {
      return jsonResponse({ ok: true, status: 'blocked_cooldown' });
    }
    cache.put(cooldownKey, String(now), Math.ceil(COOLDOWN_MS / 1000));

    const duplicateValue = [
      name.toLowerCase().replace(/\s+/g, ' '),
      phone,
      finance.toLowerCase()
    ].join('|');
    const duplicateKey = 'lead-duplicate:' + hashKey(duplicateValue);

    if (cache.get(duplicateKey)) {
      return jsonResponse({ ok: true, status: 'blocked_duplicate' });
    }
    cache.put(duplicateKey, '1', DUP_WINDOW_SECONDS);

    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.getSheets()[0];
    if (!sheet) {
      return jsonResponse({ ok: false, status: 'sheet_not_found' });
    }

    ensureSheetHeader(sheet);

    const timeText = Utilities.formatDate(
      new Date(),
      'Asia/Ho_Chi_Minh',
      'dd/MM/yyyy HH:mm:ss'
    );
    const phoneForSheet = phone ? "'" + phone : '';

    sheet.appendRow([timeText, name, phoneForSheet, finance]);
    formatLastRow(sheet);

    // Chi gui email lead chinh. Loi email khong lam mat du lieu da luu.
    try {
      sendLeadEmail({
        name: name,
        phone: phone,
        finance: finance,
        submittedAt: new Date()
      });
    } catch (mailError) {
      Logger.log('Lead email error: ' + getErrorMessage(mailError));
    }

    return jsonResponse({ ok: true, status: 'saved' });
  } catch (error) {
    Logger.log('Submission error: ' + getErrorMessage(error));
    return jsonResponse({
      ok: false,
      status: 'server_error',
      error: getErrorMessage(error)
    });
  }
}

// ==================== GOOGLE SHEET ====================
function ensureSheetHeader(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Thoi Gian',
      'Ho Va Ten',
      'So Dien Thoai',
      'Nhu Cau Cu The'
    ]);
    styleHeader(sheet);
    sheet.getRange('C:C').setNumberFormat('@STRING@');
    return;
  }

  // Dong bo ten cot neu Sheet cu van dang dung ten Tai Chinh.
  const header = String(sheet.getRange(1, 4).getDisplayValue() || '').trim().toLowerCase();
  if (header === 'tài chính' || header === 'tai chinh') {
    sheet.getRange(1, 4).setValue('Nhu Cau Cu The');
  }
}

function styleHeader(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, 4);
  headerRange.setFontWeight('bold');
  headerRange.setBackground(BRAND.primary);
  headerRange.setFontColor('#F5C842');
}

function formatLastRow(sheet) {
  const row = sheet.getLastRow();
  if (row % 2 === 0) {
    sheet.getRange(row, 1, 1, 4).setBackground('#f0f4f2');
  }
  sheet.autoResizeColumns(1, 4);
}

// ==================== EMAIL ====================
function sendLeadEmail(lead) {
  const recipients = getNotificationEmails();
  if (!recipients.length) {
    Logger.log('Chua co email nhan thong bao.');
    return;
  }

  const submittedAt = lead.submittedAt || new Date();
  const when = Utilities.formatDate(
    submittedAt,
    'Asia/Ho_Chi_Minh',
    'dd/MM/yyyy HH:mm:ss'
  );
  const name = String(lead.name || '').trim() || 'Khach hang';
  const phone = String(lead.phone || '').trim();
  const finance = String(lead.finance || '').trim() || '(chua dien)';
  const subject = '[LEAD MOI] ' + BRAND.projectName + ' | ' + name +
    (phone ? ' | ' + phone : '');

  const textBody = [
    'LEAD MOI - ' + BRAND.projectName,
    '----------------------------------------',
    'Ho va ten        : ' + name,
    'So dien thoai    : ' + (phone || '(trong)'),
    'Nhu cau cu the   : ' + finance,
    'Thoi gian        : ' + when,
    '',
    'Goi nhanh: ' + (phone ? 'tel:' + phone : 'tel:' + BRAND.hotline),
    'Mo bang data: ' + SHEET_URL,
    'Website: ' + BRAND.website
  ].join('\n');

  MailApp.sendEmail({
    to: recipients.join(','),
    subject: subject,
    body: textBody,
    htmlBody: buildLeadEmailHtml({
      name: name,
      phone: phone,
      finance: finance,
      when: when,
      sheetUrl: SHEET_URL
    }),
    name: BRAND.projectName + ' Lead Bot',
    noReply: true
  });
}

function getNotificationEmails() {
  const seen = {};
  const result = [];

  NOTIFY_EMAILS.forEach(function(email) {
    const normalized = String(email || '').trim().toLowerCase();
    if (!normalized || !isSimpleEmail(normalized) || seen[normalized]) return;
    seen[normalized] = true;
    result.push(normalized);
  });

  return result;
}

function buildLeadEmailHtml(lead) {
  const name = escapeHtml(lead.name || 'Khach hang');
  const phone = escapeHtml(lead.phone || '(chua dien)');
  const finance = escapeHtml(lead.finance || '(chua dien)');
  const when = escapeHtml(lead.when || '');
  const phoneHref = escapeHtml(lead.phone ? 'tel:' + lead.phone : 'tel:' + BRAND.hotline);
  const sheetUrl = escapeHtml(lead.sheetUrl || SHEET_URL);

  return '<!doctype html>' +
    '<html><head><meta charset="UTF-8"></head>' +
    '<body style="margin:0;padding:24px 12px;background:' + BRAND.bg + ';font-family:Arial,sans-serif;color:' + BRAND.text + ';">' +
    '<table width="100%" cellspacing="0" cellpadding="0"><tr><td align="center">' +
    '<table width="640" cellspacing="0" cellpadding="0" style="width:100%;max-width:640px;background:#fff;border:1px solid #e2e8f0;border-radius:14px;overflow:hidden;">' +
    '<tr><td style="padding:24px;background:' + BRAND.primary + ';color:#fff;">' +
    '<div style="font-size:12px;color:#cbd5e1;">THONG BAO LEAD MOI</div>' +
    '<div style="font-size:24px;font-weight:700;margin-top:6px;">' + escapeHtml(BRAND.projectName) + '</div>' +
    '</td></tr>' +
    '<tr><td style="padding:24px;">' +
    '<p style="margin:0 0 16px;font-size:15px;">Co khach hang vua dien form tren <strong>' + escapeHtml(BRAND.tagline) + '</strong>.</p>' +
    '<table width="100%" cellspacing="0" cellpadding="8" style="border-collapse:collapse;">' +
    '<tr><td width="150" style="color:' + BRAND.subtext + ';">Ho va ten</td><td><strong>' + name + '</strong></td></tr>' +
    '<tr><td style="color:' + BRAND.subtext + ';">So dien thoai</td><td><a href="' + phoneHref + '">' + phone + '</a></td></tr>' +
    '<tr><td style="color:' + BRAND.subtext + ';">Nhu cau cu the</td><td>' + finance + '</td></tr>' +
    '<tr><td style="color:' + BRAND.subtext + ';">Thoi gian</td><td>' + when + '</td></tr>' +
    '</table>' +
    '<div style="margin-top:22px;">' +
    '<a href="' + phoneHref + '" style="display:inline-block;margin-right:8px;background:' + BRAND.primary + ';color:#fff;text-decoration:none;font-size:14px;font-weight:700;padding:12px 18px;border-radius:8px;">Goi ngay</a>' +
    '<a href="' + sheetUrl + '" style="display:inline-block;background:#e2e8f0;color:' + BRAND.text + ';text-decoration:none;font-size:14px;font-weight:700;padding:12px 18px;border-radius:8px;">Mo bang data Sheet</a>' +
    '</div>' +
    '<p style="margin:20px 0 0;color:' + BRAND.subtext + ';font-size:12px;">Email tu dong tu he thong form ' + escapeHtml(BRAND.tagline) + '.</p>' +
    '</td></tr></table></td></tr></table>' +
    '</body></html>';
}

// ==================== TIEN ICH ====================
function getParams(e) {
  const params = e && e.parameter ? e.parameter : {};
  if (Object.keys(params).length) return params;

  if (e && e.postData && e.postData.contents) {
    try {
      const parsed = JSON.parse(String(e.postData.contents));
      if (parsed && typeof parsed === 'object') return parsed;
    } catch (error) {
      Logger.log('Cannot parse POST body: ' + getErrorMessage(error));
    }
  }
  return {};
}

function firstNonEmpty(values) {
  for (let i = 0; i < values.length; i++) {
    const value = String(values[i] || '').trim();
    if (value) return value;
  }
  return '';
}

function normalizePhone(phone) {
  let value = String(phone || '').trim().replace(/\s+/g, '');
  if (value.indexOf('+84') === 0) value = '0' + value.slice(3);
  if (value.indexOf('84') === 0) value = '0' + value.slice(2);
  return value.replace(/[^0-9]/g, '');
}

function isValidVNPhone(phone) {
  return /^0\d{9,10}$/.test(String(phone || '').trim());
}

function isSimpleEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(String(email || '').trim());
}

function hashKey(value) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(value),
    Utilities.Charset.UTF_8
  );
  return Utilities.base64EncodeWebSafe(bytes).slice(0, 40);
}

function escapeHtml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function getErrorMessage(error) {
  return String(error && error.message ? error.message : error);
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==================== DON DEP TRIGGER CU ====================
// Chay ham nay mot lan neu truoc day ban da chay
// setupInstallableTriggers() cua code cu.
function removeOldSheetTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    const handler = trigger.getHandlerFunction();
    if (handler === 'onSheetEdit' || handler === 'onSheetChange') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  Logger.log('Da xoa cac trigger theo doi Sheet cu.');
}

// ==================== TEST THU CONG ====================
function testEmail() {
  sendLeadEmail({
    name: 'Nguyen Van Test',
    phone: '0912345678',
    finance: 'Can tu van lo goc, dien tich 80m2',
    submittedAt: new Date()
  });
  Logger.log('Da gui email test toi: ' + getNotificationEmails().join(', '));
}
