// ============================================================
// GOOGLE APPS SCRIPT – metrocity-hoalac.com Lead Collection
// Lưu dữ liệu form (Tên, SĐT, Nhu cầu cụ thể) → Google Sheet
// Gửi email thông báo lead đến danh sách email được cấu hình
// ============================================================
//
// CÁCH CÀI ĐẶT:
// 1. Mở Google Sheet của bạn → Extensions (Tiện ích mở rộng) → Apps Script
// 2. Xóa code cũ, dán toàn bộ code này vào
// 3. Cập nhật SHEET_ID bên dưới (lấy từ URL Google Sheet)
// 4. Bấm "Save" (Ctrl+S)
// 5. Bấm "Deploy" → "New deployment"
//    - Type: Web app
//    - Execute as: Me (tài khoản của bạn)
//    - Who has access: Anyone
// 6. Bấm "Deploy" → Copy URL
//    (dạng: https://script.google.com/macros/s/ABC.../exec)
// 7. Dán URL vào file index.html, thay giá trị biến SCRIPT_URL
//
// SAU KHI DEPLOY LẦN ĐẦU:
// - Chạy hàm setupInstallableTriggers() một lần để theo dõi
//   thay đổi trực tiếp trên Sheet (nếu cần)
// ============================================================

// ─── CẤU HÌNH CHÍNH ─────────────────────────────────────────
// ID của Google Sheet (lấy từ phần đường dẫn URL của sheet)
// Ví dụ: https://docs.google.com/spreadsheets/d/SHEET_ID/edit
const SHEET_ID   = '10c6qLnV1q46wwtGxCEGGKRRjQSH0o6sdMHGMiHwCyzM';
const SHEET_NAME = 'Sheet1'; // Tên tab trong Google Sheet

// Chỉ gửi tới các địa chỉ trong danh sách này, không tự lấy người có quyền trên Sheet.
// Có thể thêm nhiều địa chỉ nếu cần nhận thông báo ở nhiều hộp thư.
const DEFAULT_NOTIFY_EMAIL = 'phamvuduchuynd@gmail.com';
const NOTIFY_EMAILS = [DEFAULT_NOTIFY_EMAIL];

// ─── BẢO VỆ SPAM ─────────────────────────────────────────────
const MIN_FILL_MS       = 3000;   // Tối thiểu 3 giây mới được submit
const COOLDOWN_MS       = 30000;  // Cùng SĐT phải chờ 30 giây
const DUP_WINDOW_SECONDS = 600;   // Chặn trùng lặp trong 10 phút

// ─── THƯƠNG HIỆU ─────────────────────────────────────────────
const BRAND = {
  projectName : 'Lead Đất nền Hòa Lạc',
  tagline     : 'metrocity-hoalac.com',
  primary     : '#09353F',
  primaryDark : '#061F25',
  accent      : '#F5C842',
  accentLight : '#FFD95A',
  bg          : '#F4F8F9',
  card        : '#FFFFFF',
  text        : '#0F172A',
  subtext     : '#475569',
  hotline     : '0868868686',           // ← cập nhật hotline thực tế
  website     : 'https://metrocity-hoalac.com'
};

// ─── ENTRY POINTS ────────────────────────────────────────────
function doGet(e) {
  const p = getParams(e);
  const isSubmit = (String(p.action || '').toLowerCase() === 'submit') ||
                   !!(p.name || p.phone || p.finance);
  if (!isSubmit) return ok('alive');
  return handleSubmission(p);
}

function doPost(e) {
  const p = getParams(e);
  return handleSubmission(p);
}

// ─── XỬ LÝ CHÍNH ─────────────────────────────────────────────
function handleSubmission(p) {
  try {
    const now = Date.now();

    // 1) Honeypot — bot thường điền trường ẩn này
    if (String(p.website || '').trim() !== '') {
      return ok('blocked_honeypot');
    }

    // 2) Time trap — phải điền ≥ MIN_FILL_MS ms
    const startedAt = Number(p.form_started_at || 0);
    if (!Number.isFinite(startedAt) || startedAt <= 0 || (now - startedAt) < MIN_FILL_MS) {
      return ok('blocked_too_fast');
    }

    // 3) Validate & normalize SĐT (bắt buộc)
    const phoneRaw = firstNonEmpty([p.phone, p.phone_text, p.sdt]);
    const phone    = normalizePhone(phoneRaw);
    if (!isValidVNPhone(phone)) {
      return ok('blocked_bad_phone');
    }

    // 4) Cooldown theo SĐT
    const cache     = CacheService.getScriptCache();
    const cdKey     = 'cd:' + hashKey(phone);
    const lastTime  = Number(cache.get(cdKey) || 0);
    if (lastTime && (now - lastTime) < COOLDOWN_MS) {
      return ok('blocked_cooldown');
    }
    cache.put(cdKey, String(now), Math.ceil(COOLDOWN_MS / 1000));

    // 5) Duplicate guard — tránh ghi 2 lần giống nhau
    const name    = String(p.name || '').trim();
    const finance = String(p.finance || '').trim();
    const fpRaw   = [name.toLowerCase().replace(/\s+/g, ' '), phone, finance.toLowerCase()].join('|');
    const fpKey   = 'fp:' + hashKey(fpRaw);
    if (cache.get(fpKey)) {
      return ok('blocked_duplicate');
    }
    cache.put(fpKey, '1', DUP_WINDOW_SECONDS);

    // 6) Lấy Sheet
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    if (!sh) return fail('sheet_not_found', 'Không tìm thấy worksheet');

    // 7) Tạo header nếu sheet trống
    if (sh.getLastRow() === 0) {
      sh.appendRow(['Thời Gian', 'Họ Và Tên', 'Số Điện Thoại', 'Nhu Cầu Cụ Thể']);
      const hr = sh.getRange(1, 1, 1, 4);
      hr.setFontWeight('bold');
      hr.setBackground(BRAND.primary);
      hr.setFontColor(BRAND.accent);
      sh.getRange('C:C').setNumberFormat('@STRING@'); // Giữ số 0 đầu SĐT
    }

    // 8) Ghi dữ liệu
    const timeStr     = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm:ss');
    const phoneSheet  = phone ? "'" + phone : '';  // Dấu ' giữ số 0 đầu trong Sheet
    sh.appendRow([timeStr, name, phoneSheet, finance]);

    // 9) Format dòng mới (highlight xen kẽ + auto-resize)
    const lastRow = sh.getLastRow();
    if (lastRow % 2 === 0) {
      sh.getRange(lastRow, 1, 1, 4).setBackground('#f0f4f2');
    }
    sh.autoResizeColumns(1, 4);

    // 10) Gửi email thông báo (không để lỗi mail làm hỏng request)
    try {
      const recipients = resolveNotificationEmails(SHEET_ID);
      sendLeadNotification(recipients, {
        name   : name,
        phone  : phone,
        finance: finance,
        submittedAt: new Date()
      });
    } catch (mailErr) {
      Logger.log('Mail error: ' + String(mailErr && mailErr.message ? mailErr.message : mailErr));
    }

    return ok('saved');
  } catch (err) {
    return fail('server_error', err);
  }
}

// ─── TIỆN ÍCH CHUNG ──────────────────────────────────────────
function getParams(e) {
  const params = e && e.parameter ? e.parameter : {};
  if (Object.keys(params).length > 0) return params;
  if (e && e.postData && e.postData.contents) {
    const raw = String(e.postData.contents || '').trim();
    if (raw) {
      try {
        const parsed = JSON.parse(raw);
        if (parsed && typeof parsed === 'object') return parsed;
      } catch (_) {}
    }
  }
  return {};
}

function firstNonEmpty(list) {
  for (let i = 0; i < list.length; i++) {
    const v = String(list[i] || '').trim();
    if (v) return v;
  }
  return '';
}

function normalizePhone(phone) {
  let s = String(phone || '').trim().replace(/\s+/g, '');
  if (s.startsWith('+84')) s = '0' + s.slice(3);
  return s.replace(/[^0-9]/g, '');
}

// SĐT Việt Nam hợp lệ: bắt đầu 0, tổng 10-11 chữ số
function isValidVNPhone(phone) {
  return /^0\d{9,10}$/.test(String(phone || '').trim());
}

function isSimpleEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(String(email || '').trim().toLowerCase());
}

function hashKey(str) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8
  );
  return Utilities.base64EncodeWebSafe(bytes).slice(0, 40);
}

// ─── RESOLVE EMAIL NHẬN THÔNG BÁO ────────────────────────────
// Không gửi theo quyền chia sẻ Sheet để tránh gửi tới hộp thư lỗi/đã đầy.
function resolveNotificationEmails(sheetId) {
  const candidates = NOTIFY_EMAILS;

  const seen = {}, out = [];
  for (let i = 0; i < candidates.length; i++) {
    const em = String(candidates[i] || '').trim().toLowerCase();
    if (!em || !isSimpleEmail(em) || seen[em]) continue;
    seen[em] = true;
    out.push(em);
  }
  return out;
}

function getSheetShareRecipients(sheetId) {
  try {
    const file    = DriveApp.getFileById(sheetId);
    const seen    = {}, list = [];
    const addEmail = (em) => {
      em = String(em || '').trim().toLowerCase();
      if (em && !seen[em]) { seen[em] = true; list.push(em); }
    };

    const owner = file.getOwner();
    if (owner) addEmail(owner.getEmail());

    file.getEditors().forEach(u => addEmail(u.getEmail()));
    file.getViewers().forEach(u => addEmail(u.getEmail()));

    return list;
  } catch (err) {
    Logger.log('getSheetShareRecipients error: ' + String(err && err.message ? err.message : err));
    return [];
  }
}

// ─── GỬI EMAIL THÔNG BÁO LEAD ────────────────────────────────
function sendLeadNotification(recipients, lead) {
  if (!recipients || recipients.length === 0) return;

  const submittedAt = (lead && lead.submittedAt) ? lead.submittedAt : new Date();
  const when    = Utilities.formatDate(submittedAt, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm:ss');
  const name    = String((lead && lead.name)    || '').trim() || 'Khách hàng';
  const phone   = String((lead && lead.phone)   || '').trim();
  const finance = String((lead && lead.finance) || '').trim() || '(chưa điền)';

  const subject = '[LEAD MỚI] ' + BRAND.projectName + ' | ' + name + (phone ? ' | ' + phone : '');

  const callHref = phone ? 'tel:' + phone : 'tel:' + BRAND.hotline;
  const mailtoSubject = encodeURIComponent('Tư vấn đất nền - ' + name);
  const mailtoBody = encodeURIComponent(
    'Xin chào,\n\nThông tin lead mới:\n' +
    '- Họ tên: ' + name + '\n' +
    '- Số điện thoại: ' + phone + '\n' +
    '- Nhu cầu cụ thể: ' + finance + '\n' +
    '- Thời gian: ' + when + '\n\n' +
    'Nguồn: Form website ' + BRAND.website
  );

  const textBody = [
    'LEAD MỚI - ' + BRAND.projectName,
    '-'.repeat(40),
    'Họ tên         : ' + name,
    'Số điện thoại  : ' + (phone || '(trống)'),
    'Nhu cầu cụ thể : ' + finance,
    'Thời gian      : ' + when,
    '',
    'Gọi nhanh: ' + callHref,
    'Website: ' + BRAND.website
  ].join('\n');

  const htmlBody = buildLeadEmailHtml({
    name, phone, finance, when,
    callHref,
    gmailHref   : 'mailto:' + encodeURIComponent(recipients[0] || DEFAULT_NOTIFY_EMAIL) +
                  '?subject=' + mailtoSubject + '&body=' + mailtoBody,
    websiteHref : 'https://docs.google.com/spreadsheets/d/10c6qLnV1q46wwtGxCEGGKRRjQSH0o6sdMHGMiHwCyzM/edit'
  });

  MailApp.sendEmail({
    to      : recipients.join(','),
    subject : subject,
    body    : textBody,
    htmlBody: htmlBody,
    name    : BRAND.projectName + ' Lead Bot',
    noReply : true
  });
}

// ─── HTML EMAIL ───────────────────────────────────────────────
function buildLeadEmailHtml(v) {
  const name    = escapeHtml(v.name    || 'Khách hàng');
  const phone   = escapeHtml(v.phone   || '(chưa điền)');
  const finance = escapeHtml(v.finance || '(chưa điền)');
  const when    = escapeHtml(v.when    || '');
  const phoneDisplay = escapeHtml(v.phone || BRAND.hotline);

  return '<!doctype html>' +
  '<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
  '<body style="margin:0;padding:0;background:' + BRAND.bg + ';font-family:Arial,sans-serif;color:' + BRAND.text + ';">' +

  // Wrapper
  '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:' + BRAND.bg + ';padding:24px 12px;">' +
  '<tr><td align="center">' +
  '<table role="presentation" width="640" cellspacing="0" cellpadding="0" style="width:100%;max-width:640px;background:' + BRAND.card + ';border-radius:16px;overflow:hidden;border:1px solid #e2e8f0;">' +

  // Header
  '<tr><td style="background:linear-gradient(135deg,' + BRAND.primary + ' 0%,' + BRAND.primaryDark + ' 100%);padding:24px;">' +
  '<div style="font-size:11px;letter-spacing:1.4px;text-transform:uppercase;color:#94a3b8;">Thông báo Lead mới</div>' +
  '<div style="font-size:26px;font-weight:800;color:#ffffff;margin-top:6px;">' + escapeHtml(BRAND.projectName) + '</div>' +
  '<div style="font-size:13px;color:#e2e8f0;margin-top:8px;">Có khách hàng vừa điền form trên <strong>' + escapeHtml(BRAND.tagline) + '</strong>. Hãy phản hồi ngay để tối ưu tỷ lệ chốt!</div>' +
  '</td></tr>' +

  // Thông tin lead
  '<tr><td style="padding:24px 28px 12px;">' +
  '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:separate;border-spacing:0 12px;">' +

  // Row: Họ tên
  '<tr>' +
  '<td width="160" style="font-size:13px;color:' + BRAND.subtext + ';vertical-align:top;padding-right:12px;">👤 Họ và tên</td>' +
  '<td style="font-size:16px;font-weight:700;color:' + BRAND.text + ';">' + name + '</td>' +
  '</tr>' +

  // Row: SĐT
  '<tr>' +
  '<td style="font-size:13px;color:' + BRAND.subtext + ';vertical-align:top;padding-right:12px;">📞 Số điện thoại</td>' +
  '<td style="font-size:16px;font-weight:700;color:' + BRAND.text + ';">' +
  '<a href="tel:' + escapeHtml(v.phone || BRAND.hotline) + '" style="color:' + BRAND.primary + ';text-decoration:none;">' + phone + '</a>' +
  '</td>' +
  '</tr>' +

  // Row: Nhu cầu cụ thể
  '<tr>' +
  '<td style="font-size:13px;color:' + BRAND.subtext + ';vertical-align:top;padding-right:12px;">💬 Nhu cầu cụ thể</td>' +
  '<td style="font-size:15px;color:' + BRAND.text + ';">' + finance + '</td>' +
  '</tr>' +

  // Row: Thời gian
  '<tr>' +
  '<td style="font-size:13px;color:' + BRAND.subtext + ';vertical-align:top;padding-right:12px;">🕐 Thời gian</td>' +
  '<td style="font-size:14px;color:' + BRAND.text + ';">' + when + '</td>' +
  '</tr>' +

  '</table>' +
  '</td></tr>' +

  // Divider
  '<tr><td style="padding:0 28px;"><div style="height:1px;background:#e2e8f0;"></div></td></tr>' +

  // Action buttons
  '<tr><td style="padding:20px 28px 24px;">' +
  '<table role="presentation" cellspacing="0" cellpadding="0"><tr>' +

  '<td style="padding-right:10px;">' +
  '<a href="' + escapeHtml(v.callHref) + '" style="display:inline-block;background:' + BRAND.primary + ';color:#ffffff;text-decoration:none;font-size:14px;font-weight:700;padding:13px 20px;border-radius:10px;">' +
  '📞 Gọi ngay: ' + phoneDisplay +
  '</a>' +
  '</td>' +

  '<td style="padding-right:10px;">' +
  '<a href="' + escapeHtml(v.websiteHref) + '" style="display:inline-block;background:#e2e8f0;color:#0f172a;text-decoration:none;font-size:14px;font-weight:700;padding:13px 20px;border-radius:10px;">' +
  '📊 Xem thông tin khách' +
  '</a>' +
  '</td>' +

  '</tr></table>' +
  '</td></tr>' +

  // Footer
  '<tr><td style="padding:14px 28px;background:#f8fafc;border-top:1px solid #e2e8f0;font-size:12px;line-height:1.6;color:#64748b;">' +
  'Email tự động từ hệ thống form <strong>' + escapeHtml(BRAND.tagline) + '</strong>. ' +
  'Vui lòng phản hồi lead trong <strong>5 phút</strong> đầu để tối ưu tỷ lệ chốt.' +
  '</td></tr>' +

  '</table>' +
  '</td></tr>' +
  '</table>' +
  '</body></html>';
}

// ─── HELPER HTML ESCAPE ───────────────────────────────────────
function escapeHtml(text) {
  return String(text || '')
    .replace(/&/g,  '&amp;')
    .replace(/</g,  '&lt;')
    .replace(/>/g,  '&gt;')
    .replace(/"/g,  '&quot;')
    .replace(/'/g,  '&#39;');
}

// ─── RESPONSE HELPERS ─────────────────────────────────────────
function ok(status) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, status: status }))
    .setMimeType(ContentService.MimeType.JSON);
}

function fail(status, err) {
  return ContentService
    .createTextOutput(JSON.stringify({
      ok    : false,
      status: status,
      error : String(err && err.message ? err.message : err)
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── TRIGGER: THEO DÕI THAY ĐỔI TRỰC TIẾP TRÊN SHEET ─────────
// (Chạy hàm setupInstallableTriggers() một lần sau khi deploy)

function setupInstallableTriggers() {
  const ss           = SpreadsheetApp.openById(SHEET_ID);
  const spreadsheetId = ss.getId();
  const allTriggers  = ScriptApp.getProjectTriggers();

  // Xóa trigger cũ tránh trùng
  allTriggers.forEach(function(t) {
    const fn = t.getHandlerFunction();
    if ((fn === 'onSheetEdit' || fn === 'onSheetChange') &&
        t.getTriggerSourceId() === spreadsheetId) {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('onSheetEdit')
    .forSpreadsheet(spreadsheetId)
    .onEdit()
    .create();

  ScriptApp.newTrigger('onSheetChange')
    .forSpreadsheet(spreadsheetId)
    .onChange()
    .create();

  Logger.log('✅ Installable triggers đã được tạo cho: ' + spreadsheetId);
}

function removeInstallableTriggers() {
  const ss           = SpreadsheetApp.openById(SHEET_ID);
  const spreadsheetId = ss.getId();

  ScriptApp.getProjectTriggers().forEach(function(t) {
    const fn = t.getHandlerFunction();
    if ((fn === 'onSheetEdit' || fn === 'onSheetChange') &&
        t.getTriggerSourceId() === spreadsheetId) {
      ScriptApp.deleteTrigger(t);
    }
  });

  Logger.log('🗑️ Installable triggers đã được xóa.');
}

// Được gọi khi ai đó sửa cell trực tiếp trên Sheet
function onSheetEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (!sheet || sheet.getName() !== SHEET_NAME) return;
    const row = e.range.getRow();
    if (row <= 1) return; // Bỏ qua hàng header

    // Gửi thông báo về thay đổi
    _sendSheetChangeNotif('EDIT', sheet, row, e);
  } catch (err) {
    Logger.log('onSheetEdit error: ' + String(err && err.message ? err.message : err));
  }
}

// Được gọi khi có thay đổi cấu trúc (thêm dòng, v.v.)
function onSheetChange(e) {
  try {
    if (!e || !e.source) return;
    const changeType = String(e.changeType || '').toUpperCase();
    if (changeType !== 'INSERT_ROW' && changeType !== 'EDIT') return;

    const sh = e.source.getSheetByName(SHEET_NAME) || e.source.getSheets()[0];
    if (!sh) return;

    const lastRow = sh.getLastRow();
    if (lastRow <= 1) return;

    _sendSheetChangeNotif(changeType, sh, lastRow, e);
  } catch (err) {
    Logger.log('onSheetChange error: ' + String(err && err.message ? err.message : err));
  }
}

// Nội bộ: gửi email thông báo thay đổi sheet
function _sendSheetChangeNotif(changeType, sheet, row, eventObj) {
  const lastCol  = Math.max(1, sheet.getLastColumn());
  const rowVals  = sheet.getRange(row, 1, 1, lastCol).getDisplayValues()[0] || [];

  const name     = String(rowVals[1] || '').trim();
  const phone    = String(rowVals[2] || '').trim().replace(/^'/, '');
  const finance  = String(rowVals[3] || '').trim();

  const recipients = resolveNotificationEmails(SHEET_ID);
  if (!recipients.length) {
    Logger.log('Không tìm thấy người nhận thông báo thay đổi sheet.');
    return;
  }

  const now       = new Date();
  const when      = Utilities.formatDate(now, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm:ss');
  const editedCell = eventObj && eventObj.range ? eventObj.range.getA1Notation() : 'N/A';

  const subject = '[SHEET ' + changeType + '] ' + BRAND.projectName + ' | Dòng ' + row;
  const body = [
    'Phát hiện thay đổi dữ liệu trên Google Sheet ' + BRAND.projectName + '.',
    '',
    'Loại thay đổi  : ' + changeType,
    'Sheet          : ' + sheet.getName(),
    'Dòng           : ' + row,
    'Ô vừa chỉnh   : ' + editedCell,
    '',
    'Họ tên         : ' + (name    || '(trống)'),
    'Số điện thoại  : ' + (phone   || '(trống)'),
    'Nhu cầu cụ thể : ' + (finance || '(trống)'),
    'Thời gian      : ' + when
  ].join('\n');

  MailApp.sendEmail({
    to     : recipients.join(','),
    subject: subject,
    body   : body,
    name   : BRAND.projectName + ' Sheet Watcher',
    noReply: true
  });
}

// ─── HÀM TEST (không deploy – chạy thủ công trong IDE) ───────
function testSubmit() {
  const fakeParams = {
    name           : 'Nguyễn Văn Test',
    phone          : '0912345678',
    finance        : '1–2 tỷ',
    form_started_at: String(Date.now() - 5000) // giả lập điền 5 giây
  };
  const result = handleSubmission(fakeParams);
  Logger.log('testSubmit result: ' + result.getContent());
}

function testEmail() {
  const recipients = resolveNotificationEmails(SHEET_ID);
  Logger.log('Recipients: ' + recipients.join(', '));
  sendLeadNotification(recipients, {
    name       : 'Nguyễn Văn Test',
    phone      : '0912345678',
    finance    : '1–2 tỷ',
    submittedAt: new Date()
  });
  Logger.log('✅ Test email đã gửi.');
}
