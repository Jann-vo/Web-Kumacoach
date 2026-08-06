// ══════════════════════════════════════════════════════════
//  KumaCoach — Google Apps Script v3
//  Xử lý 2 loại request:
//  1. Đơn hàng từ website (có field "name")
//  2. Webhook từ Sepay khi có tiền vào (có field "transferAmount")
// ══════════════════════════════════════════════════════════

var SHEET_NAME  = 'Đơn Hàng';
var ADMIN_EMAIL = 'voquocan.13122000@gmail.com';

// Mã bí mật cho link "Xác nhận 1 chạm" trong email cảnh báo (có thể đổi thành chuỗi của bạn).
var CONFIRM_TOKEN = 'kmc_x7Qp2Lr9Vt4Wn8Zc6Bd1Hf3Ms5Jy0Gk';
// URL web app /exec (dự phòng khi ScriptApp.getService().getUrl() rỗng). Cập nhật nếu deploy đổi URL.
var GAS_EXEC_URL  = 'https://script.google.com/macros/s/AKfycbymZw0gaSWUexI3TBeojLDW2o1NCyE5t3SQSQYV6FREM2bfmnl2JxCmi-cGHA7sQFY3/exec';
// Bảng admin: mở bằng  <GAS_EXEC_URL>?action=admin&token=<CONFIRM_TOKEN>
var LINK_SHEET   = 'Link Ebook';   // tab lưu link ebook — sửa từ panel, khỏi sửa code
var CMS_SHEET_ID = '1-QXao7RuHbztzQYGvNw8fjuKReJ_daVN9A0sdnjKTPI';  // spreadsheet số liệu hiển thị trên web (loadCMS)

// Map tên sản phẩm → link Google Drive (anyone with link can view)
var EBOOK_LINKS = {
  'The Model Elevate Program':                'https://drive.google.com/file/d/1heSdztyVOtUGpZx84JCMHctQ6yzUr3cp/view?usp=sharing',
  "The Men's Physique Training Method":       'https://drive.google.com/file/d/1X5FAJYlcCQpYv_BvuUkOLtVtSkiIYjPu/view?usp=sharing',
  'The Bikini Heaven':                        'https://drive.google.com/file/d/1Jg-aqL2Fstxgbr2SkBA88Uwo73xJmYTx/view?usp=sharing',
  'The Wellness Blueprint':                   'https://drive.google.com/file/d/1HbJAq3Vs7d7W5cTB3LNXZC9BG9N8Ih-k/view?usp=sharing',
  'The Classic Aesthetics Program':           'https://drive.google.com/file/d/10hyZIg5GGAr2F_UfPh9CWZI1NcoIpwvF/view?usp=sharing',
  'A Science-Based Guide to Building Muscle': 'https://drive.google.com/file/d/1VFXpCLRANL4y-4MFy3gMwojVRIB_fEGl/view?usp=sharing',
  'An Toàn Trên Đường Đua':                   'https://drive.google.com/file/d/1D9vxTdSf9tJNsJmNEHKsT2C9DtGuEu9e/view?usp=drive_link'
};

// ── Entry point ───────────────────────────────────────────
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // Phân biệt: Sepay webhook hay đơn hàng từ website
    if (data.transferAmount !== undefined) {
      return handleSepayWebhook(data);
    } else {
      return handleNewOrder(data);
    }

  } catch (err) {
    return ok({ status: 'error', message: err.message });
  }
}

// ── 1. Xử lý đơn hàng mới từ website ────────────────────
function handleNewOrder(data) {
  var sheet = getOrCreateSheet();
  var now   = timestamp();

  // Chống ghi trùng: web ghi đơn 2 lần (sớm ở bước 2 khi hiện mã + lúc bấm nút xác nhận).
  // Nếu mã đơn đã có trong sheet thì bỏ qua, KHÔNG ghi/không gửi email lần 2.
  if (data.orderCode) {
    var existing = sheet.getDataRange().getValues();
    for (var j = 1; j < existing.length; j++) {
      if (normCode(existing[j][1]) === normCode(data.orderCode)) {
        return ok({ status: 'duplicate' });
      }
    }
  }

  sheet.appendRow([
    now,
    data.orderCode || '',
    data.name      || '',
    data.phone     || '',
    data.email     || '',
    data.note      || '',
    data.items     || '',
    Number(data.total) || 0,
    'Chờ xác nhận'
  ]);

  sendAdminEmail(data, now);
  if (data.email) sendCustomerPendingEmail(data, now);

  return ok({ status: 'ok' });
}

// ── 2. Xử lý webhook Sepay ───────────────────────────────
function handleSepayWebhook(data) {
  // Lấy nội dung chuyển khoản
  var content = (data.content || data.code || '').toUpperCase();

  // Tìm mã đơn KUMA trong nội dung CK
  var orderCode = extractOrderCode(content);
  if (!orderCode) {
    // Không có mã KUMA — ghi nhận + CẢNH BÁO admin (khách CK quên ghi mã đơn)
    logUnmatchedTransaction(data);
    notifyUnmatched(data);
    return ok({ status: 'no_match' });
  }

  // Tìm đơn hàng trong Sheet
  var sheet = getOrCreateSheet();
  var rows  = sheet.getDataRange().getValues();
  var found = -1;

  for (var i = 1; i < rows.length; i++) {
    if (normCode(rows[i][1]) === orderCode &&
        rows[i][8] === 'Chờ xác nhận') {
      found = i + 1; // row index (1-based)
      break;
    }
  }

  if (found === -1) {
    logUnmatchedTransaction(data);
    notifyUnmatched(data);
    return ok({ status: 'order_not_found', code: orderCode });
  }

  // Cập nhật trạng thái → Đã xác nhận
  sheet.getRange(found, 9).setValue('✅ Đã xác nhận');
  sheet.getRange(found, 9).setBackground('#d4edda').setFontColor('#155724');

  // Ghi thêm: số tiền thực nhận & thời gian xác nhận
  var row         = rows[found - 1];
  var codeDisplay = row[1] || orderCode;   // mã gốc có dấu cách để hiển thị cho khách
  var custEmail   = row[4];
  var custName    = row[2];
  var items       = row[6];
  var orderAmt    = row[7];

  // Gửi email xác nhận + link ebook cho khách
  if (custEmail) {
    sendCustomerConfirmedEmail({
      email     : custEmail,
      name      : custName,
      orderCode : codeDisplay,
      items     : items,
      total     : orderAmt
    });
  }

  // Thông báo admin
  sendAdminConfirmedEmail({
    orderCode     : codeDisplay,
    name          : custName,
    email         : custEmail,
    items         : items,
    total         : orderAmt,
    transferAmount: data.transferAmount
  });

  return ok({ status: 'confirmed', code: codeDisplay });
}

// ── Chuẩn hoá mã: bỏ MỌI ký tự không phải chữ/số, viết HOA ──
// Ngân hàng/khách rất hay xoá dấu cách (hoặc thay bằng '-', '.') trong
// nội dung CK, nên phải so khớp theo dạng đã chuẩn hoá — KHÔNG phụ thuộc dấu cách.
function normCode(s) {
  return String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
}

// ── Tìm mã KUMA trong nội dung CK (không phụ thuộc dấu cách) ──
function extractOrderCode(content) {
  var flat  = normCode(content);            // "KUMA 260803 743" | "KUMA-260803-743" → "KUMA260803743"
  var match = flat.match(/KUMA\d{6}\d{3}/); // YYMMDD (6 số) + 3 số ngẫu nhiên
  return match ? match[0] : null;           // trả về dạng chuẩn hoá "KUMA260803743"
}

// ── Log giao dịch không khớp ──────────────────────────────
function logUnmatchedTransaction(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Giao Dịch Khác') || ss.insertSheet('Giao Dịch Khác');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Thời gian', 'Số tiền', 'Nội dung', 'Ngân hàng', 'Mã GD']);
  }
  sheet.appendRow([
    timestamp(),
    data.transferAmount || 0,
    data.content || data.code || '',
    data.gateway || '',
    data.referenceCode || ''
  ]);
}

// ── Email: Khách đặt hàng thành công (chờ xác nhận) ──────
function sendCustomerPendingEmail(data, time) {
  var total   = Number(data.total).toLocaleString() + ' ₫';
  var subject = '✅ Đặt hàng thành công [' + data.orderCode + '] — KumaCoach';
  var html = emailWrap(
    '📦 ĐẶT HÀNG THÀNH CÔNG',
    '<p style="color:#ccc;font-size:14px;line-height:1.8">Xin chào <strong style="color:#fff">' + (data.name||'bạn') + '</strong>,<br>Chúng tôi đã nhận đơn hàng. Vui lòng chờ xác nhận thanh toán.</p>'
    + infoBox([
        ['Mã đơn hàng', '<strong style="color:#D4A017;letter-spacing:2px">' + data.orderCode + '</strong>'],
        ['Sản phẩm',    data.items || ''],
        ['Tổng tiền',   '<strong style="color:#D4A017">' + total + '</strong>'],
        ['Email nhận',  data.email || '']
      ])
    + '<div style="background:rgba(212,160,23,.08);border:1px solid rgba(212,160,23,.2);border-radius:8px;padding:12px 16px;font-size:13px;color:#aaa;line-height:1.8;margin-top:16px">'
    + '⚠️ Ebook sẽ được gửi trong vòng <strong style="color:#fff">24h</strong> sau khi xác nhận thanh toán.<br>📩 Email chứa link tải có thể nằm ở mục <strong style="color:#fff">Thư rác / Quảng cáo</strong> — nhớ kiểm tra ở đó nếu chưa thấy.</div>'
    + footer()
  );
  MailApp.sendEmail({ to: data.email, subject: subject, htmlBody: html });
}

// ── Email: Khách đã thanh toán — gửi link ebook ──────────
function sendCustomerConfirmedEmail(d) {
  var total    = Number(d.total).toLocaleString() + ' ₫';
  var subject  = '🎉 Thanh toán xác nhận — Link tải Ebook! [' + d.orderCode + ']';
  var ebookBtn = buildEbookButtons(d.items || '');
  var html = emailWrap(
    '🎉 THANH TOÁN XÁC NHẬN',
    '<p style="color:#ccc;font-size:14px;line-height:1.8">Xin chào <strong style="color:#fff">' + (d.name||'bạn') + '</strong>,<br>'
    + 'Thanh toán của bạn đã được xác nhận. Bấm nút bên dưới để tải ebook ngay!</p>'
    + infoBox([
        ['Mã đơn hàng', '<strong style="color:#D4A017">' + d.orderCode + '</strong>'],
        ['Sản phẩm',    (d.items||'').replace(/\|/g,'<br>')],
        ['Tổng tiền',   '<strong style="color:#D4A017">' + total + '</strong>']
      ])
    + ebookBtn
    + '<p style="font-size:12px;color:#555;margin-top:20px">Thắc mắc? <a href="https://www.facebook.com/ThanhDatPham281" style="color:#D4A017">Facebook</a> · <a href="https://www.instagram.com/teamkumacoaching/" style="color:#D4A017">Instagram</a></p>'
    + footer()
  );
  MailApp.sendEmail({ to: d.email, subject: subject, htmlBody: html });
}

// Tạo nút tải ebook dựa vào tên sản phẩm
function buildEbookButtons(items) {
  var names = items.split('|').map(function(s){ return s.trim(); });
  var html  = '<div style="margin-top:20px">';
  names.forEach(function(name) {
    var link = getEbookLink(name);
    if (link) {
      html += '<a href="' + link + '" target="_blank" style="display:block;text-align:center;background:#D4A017;color:#000;font-weight:700;font-size:14px;text-decoration:none;padding:14px 24px;border-radius:8px;margin-bottom:10px">📥 Tải ebook: ' + name + '</a>';
    }
  });
  html += '</div>';
  return html;
}

// ── Email: Admin — đơn hàng mới ──────────────────────────
function sendAdminEmail(data, time) {
  var total   = Number(data.total).toLocaleString() + ' ₫';
  var subject = '🛒 ĐƠN MỚI [' + (data.orderCode||'') + '] — ' + (data.name||'Khách hàng');
  var html = emailWrap(
    '🛒 ĐƠN HÀNG MỚI',
    infoBox([
        ['Mã đơn',     '<strong style="color:#D4A017">' + (data.orderCode||'') + '</strong>'],
        ['Thời gian',  time],
        ['Họ tên',     data.name||''],
        ['SĐT',        data.phone||''],
        ['Email',      data.email||''],
        ['Nội dung CK','<strong style="color:#D4A017">' + (data.note||'') + '</strong>'],
        ['Sản phẩm',   (data.items||'').replace(/\|/g,'<br>')],
        ['Tổng tiền',  '<strong style="color:#D4A017;font-size:18px">' + total + '</strong>']
      ])
    + '<p style="font-size:12px;color:#555;margin-top:16px">→ Kiểm tra ngân hàng → Sepay sẽ tự xác nhận khi nhận được tiền.</p>'
  );
  MailApp.sendEmail({ to: ADMIN_EMAIL, subject: subject, htmlBody: html });
}

// ── Email: Admin — đã nhận tiền tự động ──────────────────
function sendAdminConfirmedEmail(d) {
  var subject = '💰 NHẬN TIỀN [' + d.orderCode + '] — ' + (d.name||'');
  var html = emailWrap(
    '💰 ĐÃ NHẬN TIỀN — TỰ ĐỘNG XÁC NHẬN',
    '<p style="color:#4cd98a;font-size:14px">Sepay đã phát hiện giao dịch và tự động xác nhận đơn hàng.</p>'
    + infoBox([
        ['Mã đơn',        '<strong style="color:#D4A017">' + d.orderCode + '</strong>'],
        ['Khách hàng',    (d.name||'') + ' · ' + (d.email||'')],
        ['Sản phẩm',      (d.items||'').replace(/\|/g,'<br>')],
        ['Tiền đặt hàng', Number(d.total).toLocaleString() + ' ₫'],
        ['Tiền nhận được','<strong style="color:#4cd98a">' + Number(d.transferAmount).toLocaleString() + ' ₫</strong>']
      ])
    + '<p style="font-size:13px;color:#D4A017;margin-top:16px;font-weight:700">→ Gửi ebook về email: ' + (d.email||'') + '</p>'
  );
  MailApp.sendEmail({ to: ADMIN_EMAIL, subject: subject, htmlBody: html });
}

// ── Helpers ───────────────────────────────────────────────
function getOrCreateSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Thời gian','Mã đơn','Họ tên','SĐT','Email','Nội dung CK','Sản phẩm','Tổng tiền (₫)','Trạng thái']);
    var r = sheet.getRange(1,1,1,9);
    r.setFontWeight('bold').setBackground('#111111').setFontColor('#D4A017');
    sheet.setColumnWidths(1,9,140);
  }
  return sheet;
}

function timestamp() {
  return Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm:ss');
}

function ok(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function infoBox(rows) {
  var html = '<table style="width:100%;border-collapse:collapse;font-size:14px;margin-top:12px">';
  rows.forEach(function(r) {
    html += '<tr><td style="padding:7px 0;color:#666;font-size:12px;letter-spacing:1px;text-transform:uppercase;border-bottom:1px solid #1a1a1a;width:35%">'
          + r[0] + '</td><td style="padding:7px 0;color:#fff;border-bottom:1px solid #1a1a1a">' + r[1] + '</td></tr>';
  });
  return html + '</table>';
}

function emailWrap(title, body) {
  return '<div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;background:#0d0d0d;color:#fff;border-radius:12px;overflow:hidden">'
    + '<div style="background:#D4A017;padding:16px 24px"><h2 style="margin:0;color:#000;font-size:16px">' + title + '</h2></div>'
    + '<div style="padding:24px">' + body + '</div></div>';
}

function footer() {
  return '<p style="font-size:11px;color:#333;margin-top:20px">© 2026 KumaCoach · Team Đạt Kuma</p>';
}

// ── Test thủ công ─────────────────────────────────────────
function testSetup() {
  var sheet = getOrCreateSheet();
  sheet.appendRow([timestamp(),'KUMA 260629 001','Test User','0909123456',ADMIN_EMAIL,'KUMA 260629 001','The Model Elevate Program',1000000,'Chờ xác nhận']);
  Logger.log('✅ Đã thêm đơn test');
}

function testSepayWebhook() {
  // Giả lập Sepay gửi webhook
  var fakeData = {
    transferAmount: 1000000,
    content: 'KUMA 260629 001 chuyen tien',
    code: 'KUMA 260629 001',
    gateway: 'BIDV',
    referenceCode: 'TEST123'
  };
  var result = handleSepayWebhook(fakeData);
  Logger.log(result.getContent());
}

function testExtract() {
  // Kiểm tra regex khớp mã bất kể ngân hàng có giữ dấu cách hay không
  ['KUMA 260803 743','KUMA260803743','KUMA-260803-743','CT DEN KUMA260803743 CAM ON']
    .forEach(function(c){ Logger.log(c + '  →  ' + extractOrderCode(c)); });
}

// ══════════════════════════════════════════════════════════
//  DỰ PHÒNG: xác nhận & gửi ebook THỦ CÔNG cho 1 đơn
//  Dùng khi Sepay không tự khớp: mở sheet "Đơn Hàng", bấm vào
//  hàng đơn cần giao → Menu 🛒 KumaCoach → Xác nhận & gửi ebook.
// ══════════════════════════════════════════════════════════
function confirmSelectedOrderManual() {
  var ui    = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { ui.alert('Không tìm thấy sheet "' + SHEET_NAME + '".'); return; }

  var row = sheet.getActiveRange().getRow();
  if (row <= 1) { ui.alert('Hãy bấm chọn đúng HÀNG đơn hàng (không phải hàng tiêu đề).'); return; }

  var v         = sheet.getRange(row, 1, 1, 9).getValues()[0];
  var orderCode = v[1], name = v[2], email = v[4], items = v[6], total = v[7], status = v[8];
  if (!email) { ui.alert('Hàng này không có email.'); return; }

  if (String(status).indexOf('xác nhận') !== -1 && String(status).indexOf('Chờ') === -1) {
    var r = ui.alert('Đơn này đã xác nhận rồi. Gửi lại link ebook?', ui.ButtonSet.YES_NO);
    if (r !== ui.Button.YES) return;
  }

  sendCustomerConfirmedEmail({ email: email, name: name, orderCode: orderCode, items: items, total: total });

  sheet.getRange(row, 9).setValue('✅ Đã xác nhận (thủ công)')
       .setBackground('#d4edda').setFontColor('#155724');
  ui.alert('✅ Đã gửi link ebook cho: ' + email);
}

// Gửi ebook + tạo đơn THỦ CÔNG cho 1 khách chưa có trong sheet (hiếm).
// Sửa 5 biến dưới rồi Run hàm này trong editor.
function rescueOrderManual() {
  var CODE  = 'KUMA 000000 000';            // ← mã đơn (bịa cũng được)
  var NAME  = 'Tên khách';                  // ←
  var EMAIL = 'email_khach@gmail.com';      // ←
  var ITEMS = 'The Model Elevate Program';  // ← đúng tên trong EBOOK_LINKS
  var TOTAL = 1000000;                      // ←
  sendCustomerConfirmedEmail({ email: EMAIL, name: NAME, orderCode: CODE, items: ITEMS, total: TOTAL });
  getOrCreateSheet().appendRow([timestamp(), CODE, NAME, '', EMAIL, CODE, ITEMS, TOTAL, '✅ Đã xác nhận (thủ công)']);
  Logger.log('Đã gửi ebook cho ' + EMAIL);
}

// Tạo menu khi mở sheet. Nếu project này DÙNG CHUNG với script ebook (đã có
// onOpen khác) thì bỏ hàm onOpen này đi và gộp addItem vào menu sẵn có.
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛒 KumaCoach')
    .addItem('✅ Xác nhận & gửi ebook (đơn đang chọn)', 'confirmSelectedOrderManual')
    .addToUi();
}

// ══════════════════════════════════════════════════════════
//  CẢNH BÁO + XÁC NHẬN 1 CHẠM khi tiền vào KHÔNG khớp đơn
//  (khách chuyển khoản quên ghi mã đơn → đơn kẹt "Chờ xác nhận")
// ══════════════════════════════════════════════════════════

// Sepay báo tiền vào mà không khớp mã → tìm đơn "Chờ xác nhận" có số tiền
// gần khớp → email admin kèm nút "Xác nhận & gửi ebook". KHÔNG tự xác nhận.
// Không có đơn nghi khớp thì KHÔNG gửi (tránh spam với tiền vào không liên quan).
function notifyUnmatched(data) {
  var amount = Number(data.transferAmount) || 0;
  if (amount <= 0) return;
  var sheet = getOrCreateSheet();
  var rows  = sheet.getDataRange().getValues();
  var cands = [];
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][8] || '').indexOf('Chờ') < 0) continue;    // chỉ đơn đang chờ
    var total = Number(rows[i][7]) || 0;
    if (total > 0 && Math.abs(total - amount) <= 20000) {         // số tiền gần khớp (±20k)
      cands.push({ code: rows[i][1], name: rows[i][2], email: rows[i][4],
                   items: rows[i][6], total: total, time: rows[i][0] });
    }
  }
  if (!cands.length) return;        // không có đơn nghi khớp → coi như tiền khác, không báo
  cands.reverse();                  // đơn mới nhất lên đầu

  var base = ScriptApp.getService().getUrl() || GAS_EXEC_URL;
  var blocks = cands.map(function(c) {
    var link = base + '?action=confirm&token=' + encodeURIComponent(CONFIRM_TOKEN)
             + '&code=' + encodeURIComponent(normCode(c.code));
    return '<div style="border:1px solid #333;border-radius:8px;padding:14px;margin:12px 0">'
      + '<div style="color:#D4A017;font-weight:700;letter-spacing:1px">' + c.code + '</div>'
      + '<div style="color:#ccc;font-size:13px;margin-top:4px">' + c.name + ' · ' + c.email + '</div>'
      + '<div style="color:#ccc;font-size:13px">' + (c.items || '') + ' · '
      + Number(c.total).toLocaleString() + ' ₫ · ' + c.time + '</div>'
      + '<div style="margin-top:12px"><a href="' + link + '" '
      + 'style="display:inline-block;background:#D4A017;color:#000;font-weight:700;text-decoration:none;padding:12px 22px;border-radius:8px">'
      + '✅ Xác nhận &amp; gửi ebook cho khách này</a></div></div>';
  }).join('');

  var html = emailWrap('⚠️ TIỀN VÀO CHƯA KHỚP ĐƠN — CẦN BẠN DUYỆT',
    '<p style="color:#ccc;font-size:14px;line-height:1.8">Có tiền vào nhưng nội dung chuyển khoản '
    + 'không chứa mã đơn nên hệ thống KHÔNG tự xác nhận.</p>'
    + infoBox([
        ['Số tiền nhận', '<strong style="color:#4cd98a">' + amount.toLocaleString() + ' ₫</strong>'],
        ['Nội dung CK',  (data.content || data.code || '(trống)')],
        ['Thời gian',    timestamp()]
      ])
    + '<p style="color:#ccc;font-size:14px;margin-top:16px">Đơn nghi khớp (bấm nút để xác nhận + gửi ebook):</p>'
    + blocks
    + '<p style="font-size:12px;color:#777;margin-top:8px">Chỉ bấm khi chắc đúng đơn. Bấm là khách của đơn đó nhận ebook ngay.</p>'
    + footer());
  MailApp.sendEmail({ to: ADMIN_EMAIL,
    subject: '⚠️ Tiền vào chưa khớp đơn — cần duyệt [+' + amount.toLocaleString() + 'đ]',
    htmlBody: html });
}

// Link "Xác nhận 1 chạm" từ email gọi vào đây. Mã bí mật chặn người ngoài.
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.token !== CONFIRM_TOKEN) return htmlPage_('❌ Không có quyền truy cập.');
  if (p.action === 'admin') return adminPage_();
  if (p.action === 'confirm' && p.code) {
    var r = confirmOrderByCode(p.code, Number(p.amount) || 0);
    return htmlPage_((r.ok ? '✅ ' : '⚠️ ') + r.msg);
  }
  return htmlPage_('KumaCoach webhook đang hoạt động.');
}

// Xác nhận 1 đơn theo mã + gửi ebook (idempotent: đã xác nhận thì không gửi lại).
function confirmOrderByCode(rawCode, transferAmount) {
  var target = normCode(rawCode);
  var sheet  = getOrCreateSheet();
  var rows   = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (normCode(rows[i][1]) !== target) continue;
    if (String(rows[i][8]).indexOf('Đã xác nhận') >= 0)
      return { ok: false, msg: 'Đơn ' + (rows[i][1] || rawCode) + ' đã được xác nhận trước đó.' };
    sheet.getRange(i + 1, 9).setValue('✅ Đã xác nhận').setBackground('#d4edda').setFontColor('#155724');
    var codeDisplay = rows[i][1] || rawCode;
    var email = rows[i][4], name = rows[i][2], items = rows[i][6], total = rows[i][7];
    if (email) sendCustomerConfirmedEmail({ email: email, name: name, orderCode: codeDisplay, items: items, total: total });
    sendAdminConfirmedEmail({ orderCode: codeDisplay, name: name, email: email, items: items, total: total, transferAmount: transferAmount || total });
    return { ok: true, msg: 'Đã xác nhận đơn ' + codeDisplay + ' và gửi ebook cho ' + (email || 'khách') + '.' };
  }
  return { ok: false, msg: 'Không tìm thấy đơn có mã ' + rawCode + '.' };
}

function htmlPage_(msg) {
  return HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;background:#0d0d0d;color:#fff;min-height:100vh;padding:48px 20px;text-align:center">'
    + '<div style="max-width:480px;margin:0 auto">'
    + '<h2 style="color:#D4A017;margin-bottom:16px">KumaCoach</h2>'
    + '<p style="font-size:17px;line-height:1.6">' + msg + '</p></div></div>')
    .setTitle('KumaCoach');
}

// ══════════════════════════════════════════════════════════
//  BẢNG ADMIN — mở: <GAS_EXEC_URL>?action=admin&token=<CONFIRM_TOKEN>
//  Thống kê · Quản đơn · Link ebook · Số liệu web
// ══════════════════════════════════════════════════════════

function chkToken_(t) { if (t !== CONFIRM_TOKEN) throw new Error('Token sai — không có quyền.'); }
function admDay_(dt) { return Utilities.formatDate(dt, 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd'); }
function admDayOf_(v) {
  if (v instanceof Date) return admDay_(v);
  var m = String(v == null ? '' : v).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  return m ? (m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2)) : null;
}
function admNum_(v) { if (typeof v === 'number') return v; var t = String(v == null ? '' : v).replace(/[^\d]/g, ''); return t ? parseInt(t, 10) : 0; }
function admFmt_(n) { n = Math.round(Number(n) || 0); var neg = n < 0; n = Math.abs(n); var s = String(n), o = '', c = 0; for (var i = s.length - 1; i >= 0; i--) { o = s.charAt(i) + o; if (++c % 3 === 0 && i > 0) o = '.' + o; } return (neg ? '-' : '') + o; }
function admEsc_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }

// Link ebook lưu ở sheet "Link Ebook" (sửa từ panel), fallback map cứng EBOOK_LINKS.
function ebookLinkSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(LINK_SHEET);
  if (!sh) {
    sh = ss.insertSheet(LINK_SHEET);
    sh.appendRow(['Sản phẩm', 'Link Google Drive']);
    for (var k in EBOOK_LINKS) if (EBOOK_LINKS.hasOwnProperty(k)) sh.appendRow([k, EBOOK_LINKS[k]]);
  }
  return sh;
}
function getEbookLink(name) {
  try {
    var rows = ebookLinkSheet_().getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(name).trim() && rows[i][1]) return String(rows[i][1]).trim();
    }
  } catch (e) {}
  return EBOOK_LINKS[name] || '';
}
function admCmsSheet_(ss) {
  var shs = ss.getSheets();
  for (var i = 0; i < shs.length; i++) {
    var h = shs[i].getRange(1, 1, 1, 2).getValues()[0];
    if (String(h[0]).toLowerCase() === 'key') return shs[i];
  }
  return shs[0];
}
function adminCmsRead_() {
  var out = { overalls: '', first_place: '', students: '', years: '', _err: '' };
  try {
    var sh = admCmsSheet_(SpreadsheetApp.openById(CMS_SHEET_ID));
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) { var key = String(rows[i][0]).trim(); if (out.hasOwnProperty(key)) out[key] = rows[i][1]; }
  } catch (e) { out._err = e.message; }
  return out;
}

// Gom số liệu + danh sách đơn cho panel
function adminGather_() {
  var rows = getOrCreateSheet().getDataRange().getValues();
  var today = admDay_(new Date());
  var month = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyy-MM');
  var last7 = {}, chartDays = [], byDay = {};
  for (var a = 6; a >= 0; a--) last7[admDay_(new Date(Date.now() - a * 86400000))] = 1;
  for (var b = 13; b >= 0; b--) { var dd = admDay_(new Date(Date.now() - b * 86400000)); chartDays.push(dd); byDay[dd] = 0; }

  var st = { revC: 0, revP: 0, nC: 0, nP: 0, revToday: 0, rev7: 0, revMonth: 0 };
  var prod = {}, orders = [];
  for (var r = 1; r < rows.length; r++) {
    var row = rows[r];
    var status = String(row[8] || '');
    var confirmed = status.indexOf('Đã xác nhận') >= 0 || status.indexOf('✅') >= 0;
    var pending = !confirmed && status.indexOf('Chờ') >= 0;
    if (!confirmed && !pending) continue;   // bỏ dòng không phải đơn (log ngân hàng...)
    var name = String(row[2] || '');
    var total = admNum_(row[7]);
    orders.push({ time: String(row[0] || ''), code: String(row[1] || ''), name: name, email: String(row[4] || ''),
                  items: String(row[6] || ''), total: total, confirmed: confirmed, pending: pending });
    if (name.toUpperCase().indexOf('TEST') === 0) continue;  // thống kê tiền: loại đơn test
    var dk = admDayOf_(row[0]);
    if (confirmed) {
      st.nC++; st.revC += total;
      if (dk) { if (dk === today) st.revToday += total; if (last7[dk]) st.rev7 += total; if (dk.indexOf(month) === 0) st.revMonth += total; if (byDay[dk] !== undefined) byDay[dk] += total; }
      var clean = []; String(row[6] || '').split('|').forEach(function (x) { x = x.trim(); if (x) clean.push(x); });
      var share = clean.length ? total / clean.length : 0;
      clean.forEach(function (pn) { if (!prod[pn]) prod[pn] = { n: 0, rev: 0 }; prod[pn].n++; prod[pn].rev += share; });
    } else { st.nP++; st.revP += total; }
  }
  orders.reverse();
  var top = []; for (var k in prod) if (prod.hasOwnProperty(k)) top.push({ name: k, n: prod[k].n, rev: prod[k].rev });
  top.sort(function (x, y) { return y.rev - x.rev; });
  return { st: st, orders: orders, top: top, chart: chartDays.map(function (d) { return { day: d, rev: byDay[d] }; }) };
}

// ── Mutations (đều kiểm token) ────────────────────────────
function adminConfirm(token, code) { chkToken_(token); return confirmOrderByCode(code, 0); }
function adminResend(token, code) {
  chkToken_(token);
  var rows = getOrCreateSheet().getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (normCode(rows[i][1]) === normCode(code)) {
      var email = rows[i][4]; if (!email) return { ok: false, msg: 'Đơn không có email.' };
      sendCustomerConfirmedEmail({ email: email, name: rows[i][2], orderCode: rows[i][1] || code, items: rows[i][6], total: rows[i][7] });
      return { ok: true, msg: 'Đã gửi LẠI ebook cho ' + email + '.' };
    }
  }
  return { ok: false, msg: 'Không tìm thấy đơn ' + code };
}
function adminSetLink(token, product, url) {
  chkToken_(token);
  var sh = ebookLinkSheet_(), rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(product).trim()) { sh.getRange(i + 1, 2).setValue(url); return { ok: true, msg: 'Đã đổi link: ' + product }; }
  }
  sh.appendRow([product, url]); return { ok: true, msg: 'Đã thêm link: ' + product };
}
function adminSetCmsAll(token, o) {
  chkToken_(token);
  try {
    var sh = admCmsSheet_(SpreadsheetApp.openById(CMS_SHEET_ID));
    var rows = sh.getDataRange().getValues(), keys = ['overalls', 'first_place', 'students', 'years'];
    keys.forEach(function (key) {
      if (o[key] === undefined || o[key] === '') return;
      var done = false;
      for (var i = 1; i < rows.length; i++) { if (String(rows[i][0]).trim() === key) { sh.getRange(i + 1, 2).setValue(o[key]); done = true; break; } }
      if (!done) sh.appendRow([key, o[key]]);
    });
    return { ok: true, msg: 'Đã cập nhật số liệu web. Web tự cập nhật sau ít phút.' };
  } catch (e) { return { ok: false, msg: 'Không ghi được sheet số liệu (cần cấp quyền EDIT cho ' + ADMIN_EMAIL + '): ' + e.message }; }
}

// ── Trang admin ───────────────────────────────────────────
function adminPage_() {
  var g = adminGather_(), cms = adminCmsRead_(), st = g.st;
  var maxRev = 1; g.chart.forEach(function (d) { if (d.rev > maxRev) maxRev = d.rev; });

  function card(l, v, s) {
    return '<div style="background:#111;border:1px solid #1e1e1e;border-radius:14px;padding:16px 18px">'
      + '<div style="font-size:11px;letter-spacing:1px;text-transform:uppercase;color:#7a7a7a">' + l + '</div>'
      + '<div style="font-size:23px;font-weight:800;color:#fff;margin-top:6px">' + v + '</div>'
      + (s ? '<div style="font-size:12px;color:#8a7a52;margin-top:3px">' + s + '</div>' : '') + '</div>';
  }
  var cards = '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px">'
    + card('Doanh thu đã xác nhận', '<span style="color:#4cd98a">' + admFmt_(st.revC) + 'đ</span>', st.nC + ' đơn')
    + card('Đang chờ', admFmt_(st.revP) + 'đ', st.nP + ' đơn')
    + card('Hôm nay', '<span style="color:#D4A017">' + admFmt_(st.revToday) + 'đ</span>', '')
    + card('7 ngày', admFmt_(st.rev7) + 'đ', 'Tháng: ' + admFmt_(st.revMonth) + 'đ') + '</div>';

  var bars = '';
  g.chart.forEach(function (d) {
    var h = Math.round((d.rev / maxRev) * 90); if (d.rev > 0 && h < 3) h = 3;
    bars += '<div style="flex:1;display:flex;flex-direction:column;align-items:center;gap:3px">'
      + '<div title="' + d.day + ': ' + admFmt_(d.rev) + 'đ" style="width:70%;height:' + h + 'px;background:' + (d.rev > 0 ? '#D4A017' : '#222') + ';border-radius:3px 3px 0 0"></div>'
      + '<div style="font-size:8px;color:#666">' + d.day.slice(8) + '/' + d.day.slice(5, 7) + '</div></div>';
  });
  var chart = '<div style="background:#111;border:1px solid #1e1e1e;border-radius:14px;padding:16px 18px;margin-top:14px"><div style="font-size:11px;letter-spacing:1px;text-transform:uppercase;color:#7a7a7a;margin-bottom:10px">Doanh thu 14 ngày</div><div style="display:flex;align-items:flex-end;gap:4px;height:110px">' + bars + '</div></div>';

  var topH = g.top.length ? '' : '<div style="color:#666;font-size:13px">Chưa có đơn.</div>';
  g.top.slice(0, 8).forEach(function (p) {
    topH += '<div style="display:flex;justify-content:space-between;gap:10px;padding:8px 0;border-bottom:1px solid #1a1a1a"><span style="color:#ddd;font-size:13px">' + admEsc_(p.name) + '</span><span style="color:#fff;font-size:13px;font-weight:700;white-space:nowrap">' + admFmt_(p.rev) + 'đ <span style="color:#666;font-weight:400">· ' + p.n + '</span></span></div>';
  });
  var topBox = '<div style="background:#111;border:1px solid #1e1e1e;border-radius:14px;padding:16px 18px;margin-top:14px"><div style="font-size:11px;letter-spacing:1px;text-transform:uppercase;color:#7a7a7a;margin-bottom:6px">Ebook bán chạy</div>' + topH + '</div>';

  var orderRows = '';
  g.orders.slice(0, 100).forEach(function (o) {
    var badge = o.confirmed ? '<span style="color:#4cd98a">✓ xác nhận</span>' : '<span style="color:#D4A017">chờ</span>';
    var act = o.pending
      ? '<button onclick="confirmOrder(\'' + admEsc_(o.code) + '\')" style="background:#D4A017;color:#000;border:none;border-radius:6px;padding:6px 10px;font-weight:700;font-size:11px;cursor:pointer">✅ Xác nhận &amp; gửi</button>'
      : '<button onclick="resendOrder(\'' + admEsc_(o.code) + '\')" style="background:#1c1c1c;color:#D4A017;border:1px solid #333;border-radius:6px;padding:6px 10px;font-weight:700;font-size:11px;cursor:pointer">📩 Gửi lại</button>';
    orderRows += '<tr data-st="' + (o.pending ? 'pending' : 'confirmed') + '">'
      + '<td style="padding:8px 6px;border-bottom:1px solid #161616;color:#999;font-size:11px;white-space:nowrap">' + admEsc_(o.time) + '</td>'
      + '<td style="padding:8px 6px;border-bottom:1px solid #161616;color:#fff;font-size:12px">' + admEsc_(o.name) + '<br><span style="color:#777;font-size:11px">' + admEsc_(o.email) + '</span></td>'
      + '<td style="padding:8px 6px;border-bottom:1px solid #161616;color:#bbb;font-size:11px">' + admEsc_(o.items) + '</td>'
      + '<td style="padding:8px 6px;border-bottom:1px solid #161616;color:#fff;font-size:12px;text-align:right;white-space:nowrap">' + admFmt_(o.total) + 'đ<br>' + badge + '</td>'
      + '<td style="padding:8px 6px;border-bottom:1px solid #161616;text-align:right">' + act + '</td></tr>';
  });
  var ordersBox = '<div style="background:#111;border:1px solid #1e1e1e;border-radius:14px;padding:14px;margin-top:14px;overflow-x:auto">'
    + '<div style="display:flex;gap:8px;margin-bottom:10px;flex-wrap:wrap"><b style="color:#fff;font-size:14px;margin-right:auto">Đơn hàng</b>'
    + '<button onclick="filt(\'all\')" style="background:#1c1c1c;color:#ccc;border:1px solid #333;border-radius:6px;padding:5px 12px;font-size:11px;cursor:pointer">Tất cả</button>'
    + '<button onclick="filt(\'pending\')" style="background:#1c1c1c;color:#D4A017;border:1px solid #333;border-radius:6px;padding:5px 12px;font-size:11px;cursor:pointer">Chờ</button>'
    + '<button onclick="filt(\'confirmed\')" style="background:#1c1c1c;color:#4cd98a;border:1px solid #333;border-radius:6px;padding:5px 12px;font-size:11px;cursor:pointer">Đã xác nhận</button></div>'
    + '<table style="width:100%;border-collapse:collapse;min-width:560px">' + orderRows + '</table></div>';

  var linkRows = '', links = ebookLinkSheet_().getDataRange().getValues(), linkNames = {};
  for (var i = 1; i < links.length; i++) {
    linkNames[i] = links[i][0];
    linkRows += '<div style="padding:10px 0;border-bottom:1px solid #1a1a1a">'
      + '<div style="color:#ddd;font-size:13px;margin-bottom:6px">' + admEsc_(links[i][0]) + '</div>'
      + '<div style="display:flex;gap:8px"><input id="lk_' + i + '" value="' + admEsc_(links[i][1]) + '" style="flex:1;background:#0a0a0a;border:1px solid #333;border-radius:6px;color:#fff;padding:8px;font-size:11px">'
      + '<button onclick="saveLink(' + i + ')" style="background:#D4A017;color:#000;border:none;border-radius:6px;padding:0 14px;font-weight:700;font-size:12px;cursor:pointer">Lưu</button></div></div>';
  }
  var linksBox = '<div style="background:#111;border:1px solid #1e1e1e;border-radius:14px;padding:16px 18px;margin-top:14px"><div style="font-size:11px;letter-spacing:1px;text-transform:uppercase;color:#7a7a7a;margin-bottom:6px">Link ebook (đổi khỏi cần sửa code)</div>' + linkRows + '</div>';

  function cmsInput(id, label, v) {
    return '<div style="margin-bottom:10px"><div style="font-size:12px;color:#8a7a52;margin-bottom:4px">' + label + '</div><input id="' + id + '" value="' + admEsc_(v) + '" style="width:100%;box-sizing:border-box;background:#0a0a0a;border:1px solid #333;border-radius:6px;color:#fff;padding:9px;font-size:14px"></div>';
  }
  var cmsErr = cms._err ? '<div style="color:#ff9b9b;font-size:11px;margin-bottom:8px">⚠️ ' + admEsc_(cms._err) + '</div>' : '';
  var cmsBox = '<div style="background:#111;border:1px solid #1e1e1e;border-radius:14px;padding:16px 18px;margin-top:14px"><div style="font-size:11px;letter-spacing:1px;text-transform:uppercase;color:#7a7a7a;margin-bottom:10px">Số liệu hiển thị trên web</div>' + cmsErr
    + cmsInput('cms_overalls', 'Overall Champion (overalls)', cms.overalls)
    + cmsInput('cms_first_place', '1st Place (first_place)', cms.first_place)
    + cmsInput('cms_students', 'Học viên (students)', cms.students)
    + cmsInput('cms_years', 'Số năm (years)', cms.years)
    + '<button onclick="saveCms()" style="background:#D4A017;color:#000;border:none;border-radius:8px;padding:11px 20px;font-weight:800;font-size:13px;cursor:pointer">Lưu số liệu</button></div>';

  var tabBtn = function (id, txt) { return '<button onclick="tab(\'' + id + '\')" id="t_' + id + '" style="background:none;border:none;color:#888;font-size:13px;font-weight:700;padding:10px 6px;cursor:pointer;border-bottom:2px solid transparent">' + txt + '</button>'; };
  var body = '<div style="font-family:-apple-system,Arial,sans-serif;background:#0a0a0a;min-height:100vh;color:#fff;padding:16px">'
    + '<div style="max-width:920px;margin:0 auto">'
    + '<div style="font-size:19px;font-weight:900;letter-spacing:1px;margin-bottom:4px">KUMACOACH <span style="color:#D4A017">· ADMIN</span></div>'
    + '<div style="display:flex;gap:14px;border-bottom:1px solid #1e1e1e;margin-bottom:14px;flex-wrap:wrap">'
    + tabBtn('stats', '📊 Thống kê') + tabBtn('orders', '🛒 Đơn hàng') + tabBtn('links', '🔗 Link ebook') + tabBtn('cms', '⚙️ Số liệu web') + '</div>'
    + '<div id="stats">' + cards + chart + topBox + '</div>'
    + '<div id="orders" style="display:none">' + ordersBox + '</div>'
    + '<div id="links" style="display:none">' + linksBox + '</div>'
    + '<div id="cms" style="display:none">' + cmsBox + '</div>'
    + '</div></div>';

  var js = '<script>var T=' + JSON.stringify(CONFIRM_TOKEN) + ';var LN=' + JSON.stringify(linkNames) + ';'
    + 'function tab(id){["stats","orders","links","cms"].forEach(function(x){document.getElementById(x).style.display=(x===id)?"":"none";var b=document.getElementById("t_"+x);b.style.color=(x===id)?"#D4A017":"#888";b.style.borderBottomColor=(x===id)?"#D4A017":"transparent";});}'
    + 'function call(fn,args){google.script.run.withSuccessHandler(function(r){if(r&&r.msg)alert(r.msg);location.reload();}).withFailureHandler(function(e){alert("Lỗi: "+e.message);})[fn].apply(google.script.run,args);}'
    + 'function confirmOrder(c){if(confirm("Xác nhận & gửi ebook cho đơn "+c+"?"))call("adminConfirm",[T,c]);}'
    + 'function resendOrder(c){if(confirm("Gửi LẠI ebook cho đơn "+c+"?"))call("adminResend",[T,c]);}'
    + 'function saveLink(i){var p=LN[i];var u=document.getElementById("lk_"+i).value.trim();if(u&&p)call("adminSetLink",[T,p,u]);}'
    + 'function saveCms(){call("adminSetCmsAll",[T,{overalls:v("cms_overalls"),first_place:v("cms_first_place"),students:v("cms_students"),years:v("cms_years")}]);}'
    + 'function v(id){return document.getElementById(id).value.trim();}'
    + 'function filt(s){var r=document.querySelectorAll("[data-st]");for(var i=0;i<r.length;i++){r[i].style.display=(s==="all"||r[i].getAttribute("data-st")===s)?"":"none";}}'
    + 'tab("stats");</script>';

  return HtmlService.createHtmlOutput(body + js).setTitle('KumaCoach Admin').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
