// ═══════════════════════════════════════════════════════════════
//  KUMACOACH EBOOK STORE — GOOGLE APPS SCRIPT
//  Hướng dẫn:
//  1. Mở Google Sheet (sheet kết nối với Form)
//  2. Extensions → Apps Script → paste toàn bộ code này → Save
//  3. Chạy setupSheet()  (1 lần duy nhất)
//  4. Chạy setupTrigger() (1 lần duy nhất)
//  5. Xong! Từ giờ mọi thứ tự động.
// ═══════════════════════════════════════════════════════════════

// ── CẤU HÌNH ────────────────────────────────────────────────────
const CFG = {
  EBOOK_FOLDER_ID : '1S-NEmc3UiL3urh1ZacHcF22ay3hDgnyK', // folder Drive chứa PDF
  SENDER_EMAIL    : 'solynd.contact@gmail.com',
  SENDER_NAME     : 'KumaCoach Store',
  BANK_NAME       : 'VIETCOMBANK',
  BANK_ACCOUNT    : '1055780662',
  BANK_OWNER      : 'PHAM THANH DAT',
  MESSENGER       : 'https://m.me/ThanhDatPham281',
};

// ── DANH SÁCH EBOOK ─────────────────────────────────────────────
// Key = tên CHÍNH XÁC trong Form (copy y chang từ Form)
const EBOOKS = {
  'A science-based guide to building muscle . KUMATRAINING': {
    price    : 1190000,
    filename : 'A science-based guide to building muscle . KUMATRAINING.pdf',
    label    : 'A Science-Based Guide to Building Muscle',
  },
  'THE BIKINI HEAVEN. KUMATRAINING': {
    price    : 1050000,
    filename : 'THE BIKINI HEAVEN. KUMATRAINING.pdf',
    label    : 'The Bikini Heaven',
  },
  'THE CLASSIC AESTHETICS PROGRAM. KUMATRAINING': {
    price    : 1200000,
    filename : 'THE CLASSIC AESTHETICS PROGRAM. KUMATRAINING.pdf',
    label    : 'The Classic Aesthetics Program',
  },
  "The Men's Physique Training Method": {
    price    : 1150000,
    filename : "The Men’s Physique Training Method.pdf",
    label    : "The Men's Physique Training Method",
  },
  'The Model Elevate program. KUMATRAINING': {
    price    : 1130000,
    filename : 'The Model Elevate program. KUMATRAINING.pdf',
    label    : 'The Model Elevate Program',
  },
  'THE WELLNESS BLUEPRINT. KUMATRAINING': {
    price    : 1000000,
    filename : 'THE WELLNESS BLUEPRINT. KUMATRAINING.pdf',
    label    : 'The Wellness Blueprint',
  },
};

// ── VỊ TRÍ CỘT TRONG SHEET (A=1, B=2...) ───────────────────────
const COL = {
  TIMESTAMP : 1,  // A - Dấu thời gian (Form tự điền)
  NAME      : 2,  // B - Họ và Tên
  EMAIL     : 3,  // C - Email của bạn
  PHONE     : 4,  // D - SĐT của bạn
  EBOOK     : 5,  // E - Ebook muốn mua
  ORDER_ID  : 6,  // F - Mã đơn hàng     (script tự điền)
  TOTAL     : 7,  // G - Tổng tiền        (script tự điền)
  STATUS    : 8,  // H - Trạng thái       (script tự điền)
  SENT_AT   : 9,  // I - Ngày giao ebook  (script tự điền)
};

// ── HELPER ──────────────────────────────────────────────────────
function fmtMoney(n) {
  return n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.') + 'đ';
}

function makeOrderId() {
  return 'KUMA' + new Date().getTime().toString().slice(-7);
}

// Tìm ebook trong EBOOKS map, normalize apostrophe cong/thẳng
function findEbook(key) {
  if (EBOOKS[key]) return EBOOKS[key];
  var norm = key.replace(/[‘’ʼ]/g, "'");
  if (EBOOKS[norm]) return EBOOKS[norm];
  var lk = norm.toLowerCase();
  for (var k in EBOOKS) {
    if (k.replace(/[‘’ʼ]/g, "'").toLowerCase() === lk) return EBOOKS[k];
  }
  return null;
}

// Tìm file PDF trong folder (và subfolder 1 cấp)
function findFile(filename) {
  var folder = DriveApp.getFolderById(CFG.EBOOK_FOLDER_ID);
  var files  = folder.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    if (f.getName() === filename) return f;
  }
  var subs = folder.getFolders();
  while (subs.hasNext()) {
    var sub   = subs.next();
    var sfiles = sub.getFiles();
    while (sfiles.hasNext()) {
      var sf = sfiles.next();
      if (sf.getName() === filename) return sf;
    }
  }
  return null;
}

// ── 1. TỰ ĐỘNG KHI CÓ ĐƠN MỚI ─────────────────────────────────
function onFormSubmit(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var row   = e.range.getRow();
  var vals  = sheet.getRange(row, 1, 1, COL.EBOOK).getValues()[0];

  var name     = vals[COL.NAME  - 1] || '';
  var email    = vals[COL.EMAIL - 1] || '';
  var phone    = vals[COL.PHONE - 1] || '';
  var ebookKey = vals[COL.EBOOK - 1] || '';

  var eb      = findEbook(ebookKey);
  var price   = eb ? eb.price : 0;
  var orderId = makeOrderId();

  // Ghi thêm vào sheet
  sheet.getRange(row, COL.ORDER_ID).setValue(orderId);
  sheet.getRange(row, COL.TOTAL).setValue(price);
  sheet.getRange(row, COL.STATUS).setValue('Cho thanh toan');
  sheet.getRange(row, 1, 1, COL.SENT_AT).setBackground('#fff8e1');

  if (!email) return;

  // ── Email gửi khách ──
  var subject = '[KumaCoach] Xac nhan don hang - Ma ' + orderId;
  var html = '<div style="font-family:Arial,sans-serif;max-width:580px;margin:0 auto;border:1px solid #eee;border-radius:10px;overflow:hidden">'
    + '<div style="background:#0a0a0a;padding:28px;text-align:center">'
    + '<h2 style="color:#d4a017;margin:0;font-size:22px;letter-spacing:3px">KUMACOACH STORE</h2>'
    + '<p style="color:#888;margin:6px 0 0;font-size:13px">Vietnam-Based Coaching Team</p>'
    + '</div>'
    + '<div style="padding:32px;background:#fafafa">'
    + '<p style="font-size:16px">Chao <strong>' + name + '</strong>,</p>'
    + '<p>Cam on ban da dat mua ebook tai KumaCoach! Vui long chuyen khoan de nhan ebook.</p>'
    + '<div style="background:#fff;border:1px solid #f0e0b0;border-radius:8px;padding:20px;margin:20px 0">'
    + '<p style="margin:0 0 10px"><strong>Ebook:</strong> ' + (eb ? eb.label : ebookKey) + '</p>'
    + '<p style="margin:0 0 10px"><strong>Ma don:</strong> <code style="background:#f5f5f5;padding:3px 10px;border-radius:4px;font-size:14px">' + orderId + '</code></p>'
    + '<p style="margin:0;font-size:20px;color:#d4a017;font-weight:bold">' + fmtMoney(price) + '</p>'
    + '</div>'
    + '<div style="background:#e8f5e9;border-left:4px solid #4caf50;border-radius:4px;padding:18px;margin:20px 0">'
    + '<p style="margin:0 0 12px;font-weight:bold;font-size:15px">THONG TIN CHUYEN KHOAN</p>'
    + '<table style="border-collapse:collapse;width:100%">'
    + '<tr><td style="padding:5px 0;color:#555;width:140px">Ngan hang</td><td><strong>' + CFG.BANK_NAME + '</strong></td></tr>'
    + '<tr><td style="padding:5px 0;color:#555">So tai khoan</td><td><strong style="font-size:18px;letter-spacing:1px">' + CFG.BANK_ACCOUNT + '</strong></td></tr>'
    + '<tr><td style="padding:5px 0;color:#555">Chu tai khoan</td><td><strong>' + CFG.BANK_OWNER + '</strong></td></tr>'
    + '<tr><td style="padding:5px 0;color:#555">So tien</td><td><strong style="color:#d4a017">' + fmtMoney(price) + '</strong></td></tr>'
    + '<tr><td style="padding:5px 0;color:#555">Noi dung CK</td><td><strong style="color:#e53935;font-size:16px">' + orderId + '</strong> (bat buoc)</td></tr>'
    + '</table>'
    + '</div>'
    + '<p style="color:#555">Link download se duoc gui ve email nay trong vong <strong>15 phut</strong> sau khi xac nhan thanh toan.</p>'
    + '</div>'
    + '<div style="background:#0a0a0a;padding:16px;text-align:center">'
    + '<p style="color:#555;margin:0;font-size:12px">KumaCoach Store &middot; solynd.contact@gmail.com</p>'
    + '</div></div>';

  GmailApp.sendEmail(email, subject, '', { name: CFG.SENDER_NAME, htmlBody: html });

  // ── Thong bao cho admin ──
  GmailApp.sendEmail(
    CFG.SENDER_EMAIL,
    '[Don moi] ' + orderId + ' - ' + name + ' - ' + fmtMoney(price),
    'Ebook: ' + ebookKey + '\nEmail: ' + email + '\nSDT: ' + phone + '\nMa don: ' + orderId
  );
}

// ── 2. XÁC NHẬN THANH TOÁN & GỬI LINK (gọi từ menu) ───────────
function confirmAndSendEbook() {
  var ui    = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var row   = sheet.getActiveRange().getRow();

  if (row <= 1) { ui.alert('❌ Chọn đúng hàng đơn hàng, không chọn hàng tiêu đề.'); return; }

  var vals     = sheet.getRange(row, 1, 1, COL.SENT_AT).getValues()[0];
  var name     = vals[COL.NAME     - 1];
  var email    = vals[COL.EMAIL    - 1];
  var ebookKey = vals[COL.EBOOK    - 1];
  var orderId  = vals[COL.ORDER_ID - 1];
  var status   = vals[COL.STATUS   - 1];

  if (!email) { ui.alert('❌ Không có email trong hàng này.'); return; }

  if (status === '✅ Đã giao') {
    var res = ui.alert('⚠️ Đơn này đã giao rồi. Gửi lại link?', ui.ButtonSet.YES_NO);
    if (res !== ui.Button.YES) return;
  }

  var eb = findEbook(ebookKey);
  if (!eb) { ui.alert('Khong nhan dien ebook: "' + ebookKey + '"'); return; }

  var file = findFile(eb.filename);
  if (!file) {
    ui.alert('❌ Không tìm thấy file:\n"' + eb.filename + '"\n\nHãy upload PDF vào Drive folder và thử lại.');
    return;
  }

  // Tạo link chia sẻ
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var link = 'https://drive.google.com/file/d/' + file.getId() + '/view?usp=sharing';

  // ── Email gửi link download ──
  var subject = '[KumaCoach] Link Download Ebook - ' + orderId;
  var html = '<div style="font-family:Arial,sans-serif;max-width:580px;margin:0 auto;border:1px solid #eee;border-radius:10px;overflow:hidden">'
    + '<div style="background:#0a0a0a;padding:28px;text-align:center">'
    + '<h2 style="color:#d4a017;margin:0;font-size:22px;letter-spacing:3px">KUMACOACH STORE</h2>'
    + '<p style="color:#888;margin:6px 0 0;font-size:13px">Vietnam-Based Coaching Team</p>'
    + '</div>'
    + '<div style="padding:32px;background:#fafafa">'
    + '<p style="font-size:16px">Chao <strong>' + name + '</strong>,</p>'
    + '<p>Thanh toan da duoc xac nhan! Day la link download ebook cua ban:</p>'
    + '<div style="background:#fff;border:1px solid #c8e6c9;border-radius:10px;padding:28px;margin:24px 0;text-align:center">'
    + '<p style="margin:0 0 6px;font-size:17px;font-weight:bold;color:#333">' + eb.label + '</p>'
    + '<p style="margin:0 0 20px;color:#999;font-size:13px">Ma don: ' + orderId + '</p>'
    + '<a href="' + link + '" style="background:#d4a017;color:#000;padding:15px 36px;border-radius:8px;text-decoration:none;font-weight:bold;font-size:14px;letter-spacing:1px;display:inline-block">TAI EBOOK NGAY</a>'
    + '<p style="margin:16px 0 0;font-size:11px;color:#aaa">Hoac copy link: ' + link + '</p>'
    + '</div>'
    + '<p style="color:#c0392b;font-size:13px">Link chi danh rieng cho ban. Khong chia se de tranh bi khoa.</p>'
    + '<p style="margin-top:24px">Chuc ban luyen tap hieu qua!</p>'
    + '</div>'
    + '<div style="background:#0a0a0a;padding:16px;text-align:center">'
    + '<p style="color:#555;margin:0;font-size:12px">KumaCoach Store &middot; solynd.contact@gmail.com</p>'
    + '</div></div>';

  GmailApp.sendEmail(email, subject, '', { name: CFG.SENDER_NAME, htmlBody: html });

  // Cập nhật sheet
  var now = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm');
  sheet.getRange(row, COL.STATUS).setValue('✅ Đã giao');
  sheet.getRange(row, COL.SENT_AT).setValue(now);
  sheet.getRange(row, 1, 1, COL.SENT_AT).setBackground('#e8f5e9');

  ui.alert('✅ Đã gửi link download cho\n' + email);
}

// ── 3. SETUP SHEET (chạy 1 lần) ────────────────────────────────
function setupSheet() {
  var sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var headers = [
    'Dấu thời gian','Họ và Tên','Email','SĐT',
    'Ebook muốn mua','Mã đơn','Tổng tiền','Trạng thái','Ngày giao'
  ];
  var hRow = sheet.getRange(1, 1, 1, headers.length);
  hRow.setValues([headers]);
  hRow.setBackground('#0a0a0a').setFontColor('#d4a017').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(COL.EBOOK, 260);
  sheet.setColumnWidth(COL.ORDER_ID, 120);
  sheet.setColumnWidth(COL.STATUS, 150);
  SpreadsheetApp.getUi().alert('✅ Sheet đã setup xong!');
}

// ── 4. SETUP TRIGGER (chạy 1 lần) ──────────────────────────────
function setupTrigger() {
  // Xóa trigger cũ (tránh trùng)
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'onFormSubmit') ScriptApp.deleteTrigger(t);
  });
  // Cài trigger mới
  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();
  SpreadsheetApp.getUi().alert('✅ Trigger đã cài xong!\nMỗi đơn mới sẽ tự động gửi email cho khách.');
}

// ── 5. TỰ ĐỘNG KIỂM TRA THANH TOÁN QUA GMAIL ───────────────────
function checkPayments() {
  var sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Lấy tất cả đơn đang chờ thanh toán
  var data = sheet.getRange(2, 1, lastRow - 1, COL.SENT_AT).getValues();
  var pending = {};
  data.forEach(function(row, i) {
    var orderId = String(row[COL.ORDER_ID - 1] || '').trim();
    var status  = String(row[COL.STATUS  - 1] || '');
    if (orderId && status.indexOf('Cho') !== -1) {
      pending[orderId] = i + 2;
    }
  });
  if (!Object.keys(pending).length) return;

  // Tìm email từ VCBDigibank trong 30 phút gần nhất
  var threads = GmailApp.search('from:VCBDigibank newer_than:30m', 0, 20);
  threads.forEach(function(thread) {
    thread.getMessages().forEach(function(msg) {
      var body = msg.getPlainBody() + ' ' + msg.getSubject();
      Object.keys(pending).forEach(function(orderId) {
        if (body.indexOf(orderId) !== -1) {
          autoSendEbook(sheet, pending[orderId]);
          delete pending[orderId];
          msg.markRead();
        }
      });
    });
  });
}

function autoSendEbook(sheet, row) {
  var vals     = sheet.getRange(row, 1, 1, COL.SENT_AT).getValues()[0];
  var name     = vals[COL.NAME     - 1];
  var email    = vals[COL.EMAIL    - 1];
  var ebookKey = vals[COL.EBOOK    - 1];
  var orderId  = vals[COL.ORDER_ID - 1];
  if (!email) return;

  var eb   = findEbook(ebookKey);
  if (!eb) return;
  var file = findFile(eb.filename);
  if (!file) return;

  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var link = 'https://drive.google.com/file/d/' + file.getId() + '/view?usp=sharing';

  var subject = '[KumaCoach] Link Download Ebook - ' + orderId;
  var html = '<div style="font-family:Arial,sans-serif;max-width:580px;margin:0 auto;border:1px solid #eee;border-radius:10px;overflow:hidden">'
    + '<div style="background:#0a0a0a;padding:28px;text-align:center">'
    + '<h2 style="color:#d4a017;margin:0;font-size:22px;letter-spacing:3px">KUMACOACH STORE</h2>'
    + '<p style="color:#888;margin:6px 0 0;font-size:13px">Vietnam-Based Coaching Team</p>'
    + '</div>'
    + '<div style="padding:32px;background:#fafafa">'
    + '<p style="font-size:16px">Chao <strong>' + name + '</strong>,</p>'
    + '<p>Thanh toan da duoc xac nhan! Day la link download ebook cua ban:</p>'
    + '<div style="background:#fff;border:1px solid #c8e6c9;border-radius:10px;padding:28px;margin:24px 0;text-align:center">'
    + '<p style="margin:0 0 6px;font-size:17px;font-weight:bold;color:#333">' + eb.label + '</p>'
    + '<p style="margin:0 0 20px;color:#999;font-size:13px">Ma don: ' + orderId + '</p>'
    + '<a href="' + link + '" style="background:#d4a017;color:#000;padding:15px 36px;border-radius:8px;text-decoration:none;font-weight:bold;font-size:14px;letter-spacing:1px;display:inline-block">TAI EBOOK NGAY</a>'
    + '<p style="margin:16px 0 0;font-size:11px;color:#aaa">Hoac copy link: ' + link + '</p>'
    + '</div>'
    + '<p style="color:#c0392b;font-size:13px">Link chi danh rieng cho ban. Khong chia se de tranh bi khoa.</p>'
    + '<p>Chuc ban luyen tap hieu qua!</p>'
    + '</div>'
    + '<div style="background:#0a0a0a;padding:16px;text-align:center">'
    + '<p style="color:#555;margin:0;font-size:12px">KumaCoach Store &middot; solynd.contact@gmail.com</p>'
    + '</div></div>';

  GmailApp.sendEmail(email, subject, '', { name: CFG.SENDER_NAME, htmlBody: html });

  var now = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm');
  sheet.getRange(row, COL.STATUS).setValue('Da giao (tu dong)');
  sheet.getRange(row, COL.SENT_AT).setValue(now);
  sheet.getRange(row, 1, 1, COL.SENT_AT).setBackground('#e8f5e9');
}

function setupPaymentTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'checkPayments') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkPayments')
    .timeBased().everyMinutes(2).create();
  SpreadsheetApp.getUi().alert('Da cai trigger kiem tra thanh toan moi 2 phut!');
}

// ── 6. KIỂM TRA FILES DRIVE ────────────────────────────────────
function listDriveFiles() {
  var folder = DriveApp.getFolderById(CFG.EBOOK_FOLDER_ID);
  var out    = 'Files trong folder "' + folder.getName() + '":\n\n';
  var files  = folder.getFiles();
  var count  = 0;
  while (files.hasNext()) { out += '• ' + files.next().getName() + '\n'; count++; }
  if (count === 0) out += '(Trống — hãy upload PDF ebook vào đây)\n';
  var subs = folder.getFolders();
  while (subs.hasNext()) {
    var s  = subs.next();
    out   += '\n📁 ' + s.getName() + ':\n';
    var sf = s.getFiles();
    while (sf.hasNext()) out += '  • ' + sf.next().getName() + '\n';
  }
  SpreadsheetApp.getUi().alert(out);
}

// ── MENU ────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛒 KumaCoach')
    .addItem('✅ Xác nhận thanh toán & Gửi ebook', 'confirmAndSendEbook')
    .addSeparator()
    .addItem('🔧 Setup Sheet (chay 1 lan)', 'setupSheet')
    .addItem('🔧 Setup Trigger Form (chay 1 lan)', 'setupTrigger')
    .addItem('🔧 Setup Tu Dong Kiem Tra Thanh Toan (chay 1 lan)', 'setupPaymentTrigger')
    .addItem('📁 Kiem tra files trong Drive', 'listDriveFiles')
    .addToUi();
}
