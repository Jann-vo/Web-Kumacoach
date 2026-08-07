// ══════════════════════════════════════════════════════════
//  KumaCoach — BẢNG ADMIN (script RIÊNG, KHÔNG đụng webhook đơn hàng)
//
//  CÀI ĐẶT:
//   1) Tạo Apps Script project MỚI bằng tài khoản có quyền vào sheet đơn hàng.
//   2) Dán toàn bộ file này.
//   3) Deploy → New deployment → Web app:
//        Execute as: Me   |   Who has access: ONLY MYSELF
//      → khoá bằng đăng nhập Google của bạn, KHỎI CẦN token. Khách không thấy.
//   4) Mở URL web app → bookmark.
//  (Webhook đơn hàng giữ nguyên, không phải sửa gì.)
// ══════════════════════════════════════════════════════════

var ORDERS_SHEET_ID = '1R6VZvZrcC0UpZ8mrlTUIuW1rGu13TK26T9In6jIiSzc';  // spreadsheet chứa tab "Đơn Hàng"
var CMS_SHEET_ID    = '1-QXao7RuHbztzQYGvNw8fjuKReJ_daVN9A0sdnjKTPI';  // spreadsheet số liệu hiển thị trên web
var ADMIN_EMAIL     = 'voquocan.13122000@gmail.com';
var SHEET_NAME      = 'Đơn Hàng';
var LINK_SHEET      = 'Link Ebook';

// Link ebook mặc định — dùng khi tab "Link Ebook" chưa có / thiếu.
var EBOOK_LINKS = {
  'The Model Elevate Program':                'https://drive.google.com/file/d/1heSdztyVOtUGpZx84JCMHctQ6yzUr3cp/view?usp=sharing',
  "The Men's Physique Training Method":       'https://drive.google.com/file/d/1X5FAJYlcCQpYv_BvuUkOLtVtSkiIYjPu/view?usp=sharing',
  'The Bikini Heaven':                        'https://drive.google.com/file/d/1Jg-aqL2Fstxgbr2SkBA88Uwo73xJmYTx/view?usp=sharing',
  'The Wellness Blueprint':                   'https://drive.google.com/file/d/1HbJAq3Vs7d7W5cTB3LNXZC9BG9N8Ih-k/view?usp=sharing',
  'The Classic Aesthetics Program':           'https://drive.google.com/file/d/10hyZIg5GGAr2F_UfPh9CWZI1NcoIpwvF/view?usp=sharing',
  'A Science-Based Guide to Building Muscle': 'https://drive.google.com/file/d/1VFXpCLRANL4y-4MFy3gMwojVRIB_fEGl/view?usp=sharing',
  'An Toàn Trên Đường Đua':                   'https://drive.google.com/file/d/1D9vxTdSf9tJNsJmNEHKsT2C9DtGuEu9e/view?usp=drive_link'
};

function doGet(e) { return adminPage_(); }

// ── Truy cập dữ liệu ──────────────────────────────────────
function ordersSheet_() { return SpreadsheetApp.openById(ORDERS_SHEET_ID).getSheetByName(SHEET_NAME); }
function linkSheet_() {
  var ss = SpreadsheetApp.openById(ORDERS_SHEET_ID), sh = ss.getSheetByName(LINK_SHEET);
  if (!sh) {
    sh = ss.insertSheet(LINK_SHEET);
    sh.appendRow(['Sản phẩm', 'Link Google Drive']);
    for (var k in EBOOK_LINKS) if (EBOOK_LINKS.hasOwnProperty(k)) sh.appendRow([k, EBOOK_LINKS[k]]);
  }
  return sh;
}
function getEbookLink_(name) {
  try {
    var rows = linkSheet_().getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) if (String(rows[i][0]).trim() === String(name).trim() && rows[i][1]) return String(rows[i][1]).trim();
  } catch (e) {}
  return EBOOK_LINKS[name] || '';
}
function cmsSheet_(ss) {
  var shs = ss.getSheets();
  for (var i = 0; i < shs.length; i++) { var h = shs[i].getRange(1, 1, 1, 2).getValues()[0]; if (String(h[0]).toLowerCase() === 'key') return shs[i]; }
  return shs[0];
}

// ── Helpers ───────────────────────────────────────────────
function normCode_(s) { return String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); }
function num_(v) { if (typeof v === 'number') return v; var t = String(v == null ? '' : v).replace(/[^\d]/g, ''); return t ? parseInt(t, 10) : 0; }
function fmt_(n) { n = Math.round(Number(n) || 0); var g = n < 0; n = Math.abs(n); var s = String(n), o = '', c = 0; for (var i = s.length - 1; i >= 0; i--) { o = s.charAt(i) + o; if (++c % 3 === 0 && i > 0) o = '.' + o; } return (g ? '-' : '') + o; }
function esc_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
function day_(dt) { return Utilities.formatDate(dt, 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd'); }
function dayOf_(v) { if (v instanceof Date) return day_(v); var m = String(v == null ? '' : v).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/); return m ? (m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2)) : null; }

// ── Gom số liệu + đơn ─────────────────────────────────────
function gather_() {
  var rows = ordersSheet_().getDataRange().getValues();
  var today = day_(new Date()), month = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyy-MM');
  var last7 = {}, chartDays = [], byDay = {};
  for (var a = 6; a >= 0; a--) last7[day_(new Date(Date.now() - a * 86400000))] = 1;
  for (var b = 13; b >= 0; b--) { var dd = day_(new Date(Date.now() - b * 86400000)); chartDays.push(dd); byDay[dd] = 0; }
  var st = { revC: 0, revP: 0, nC: 0, nP: 0, revToday: 0, rev7: 0, revMonth: 0 }, prod = {}, orders = [];
  for (var r = 1; r < rows.length; r++) {
    var row = rows[r], status = String(row[8] || '');
    var confirmed = status.indexOf('Đã xác nhận') >= 0 || status.indexOf('✅') >= 0;
    var pending = !confirmed && status.indexOf('Chờ') >= 0;
    if (!confirmed && !pending) continue;
    var name = String(row[2] || ''), total = num_(row[7]);
    orders.push({ time: String(row[0] || ''), code: String(row[1] || ''), name: name, email: String(row[4] || ''), items: String(row[6] || ''), total: total, confirmed: confirmed, pending: pending });
    if (name.toUpperCase().indexOf('TEST') === 0) continue;
    var dk = dayOf_(row[0]);
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
function cmsRead_() {
  var out = { overalls: '', first_place: '', students: '', years: '', _err: '' };
  try {
    var sh = cmsSheet_(SpreadsheetApp.openById(CMS_SHEET_ID)), rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) { var key = String(rows[i][0]).trim(); if (out.hasOwnProperty(key)) out[key] = rows[i][1]; }
  } catch (e) { out._err = e.message; }
  return out;
}

// ── Gửi ebook (email giao) ────────────────────────────────
function sendDelivery_(o) {
  var link = getEbookLink_(String(o.items || '').split('|')[0].trim());
  var btn = link
    ? '<div style="margin-top:20px"><a href="' + link + '" target="_blank" style="display:block;text-align:center;background:#D4A017;color:#000;font-weight:700;font-size:14px;text-decoration:none;padding:14px 24px;border-radius:8px">📥 Tải ebook: ' + esc_(o.items) + '</a></div>'
    : '<p style="color:#ff9b9b">⚠️ Chưa có link cho sản phẩm này — kiểm tab "Link ebook".</p>';
  var html = '<div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;background:#0d0d0d;color:#fff;border-radius:12px;overflow:hidden">'
    + '<div style="background:#D4A017;padding:16px 24px"><h2 style="margin:0;color:#000;font-size:16px">🎉 THANH TOÁN XÁC NHẬN</h2></div>'
    + '<div style="padding:24px"><p style="color:#ccc;font-size:14px;line-height:1.8">Xin chào <strong style="color:#fff">' + esc_(o.name || 'bạn') + '</strong>,<br>Thanh toán đã được xác nhận. Bấm nút bên dưới để tải ebook ngay!</p>'
    + '<table style="width:100%;border-collapse:collapse;font-size:14px;margin-top:12px">'
    + '<tr><td style="padding:7px 0;color:#666;font-size:12px;text-transform:uppercase;border-bottom:1px solid #1a1a1a;width:35%">Mã đơn</td><td style="padding:7px 0;color:#D4A017;border-bottom:1px solid #1a1a1a"><strong>' + esc_(o.code) + '</strong></td></tr>'
    + '<tr><td style="padding:7px 0;color:#666;font-size:12px;text-transform:uppercase;border-bottom:1px solid #1a1a1a">Sản phẩm</td><td style="padding:7px 0;color:#fff;border-bottom:1px solid #1a1a1a">' + esc_(o.items) + '</td></tr></table>'
    + btn
    + '<p style="font-size:12px;color:#555;margin-top:20px">Thắc mắc? <a href="https://www.facebook.com/ThanhDatPham281" style="color:#D4A017">Facebook</a> · <a href="https://www.instagram.com/teamkumacoaching/" style="color:#D4A017">Instagram</a></p>'
    + '<p style="font-size:11px;color:#333;margin-top:16px">© KumaCoach · Team Đạt Kuma</p></div></div>';
  MailApp.sendEmail({ to: o.email, subject: '🎉 Link tải Ebook — KumaCoach [' + o.code + ']', htmlBody: html });
}
function findOrder_(code) {
  var sh = ordersSheet_(), rows = sh.getDataRange().getValues(), t = normCode_(code);
  for (var i = 1; i < rows.length; i++) if (normCode_(rows[i][1]) === t) return { sh: sh, row: i + 1, data: rows[i] };
  return null;
}

// ── Hành động (deploy Only myself nên không cần token) ────
function adminConfirm(code) {
  var f = findOrder_(code); if (!f) return { ok: false, msg: 'Không tìm thấy đơn ' + code };
  if (String(f.data[8]).indexOf('Đã xác nhận') >= 0) return { ok: false, msg: 'Đơn đã xác nhận trước đó.' };
  if (!f.data[4]) return { ok: false, msg: 'Đơn không có email.' };
  f.sh.getRange(f.row, 9).setValue('✅ Đã xác nhận').setBackground('#d4edda').setFontColor('#155724');
  sendDelivery_({ code: f.data[1], name: f.data[2], email: f.data[4], items: f.data[6] });
  return { ok: true, msg: 'Đã xác nhận + gửi ebook cho ' + f.data[4] };
}
function adminResend(code) {
  var f = findOrder_(code); if (!f) return { ok: false, msg: 'Không tìm thấy đơn ' + code };
  if (!f.data[4]) return { ok: false, msg: 'Đơn không có email.' };
  sendDelivery_({ code: f.data[1], name: f.data[2], email: f.data[4], items: f.data[6] });
  return { ok: true, msg: 'Đã gửi LẠI ebook cho ' + f.data[4] };
}
function adminSetLink(product, url) {
  var sh = linkSheet_(), rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) if (String(rows[i][0]).trim() === String(product).trim()) { sh.getRange(i + 1, 2).setValue(url); return { ok: true, msg: 'Đã đổi link: ' + product }; }
  sh.appendRow([product, url]); return { ok: true, msg: 'Đã thêm link: ' + product };
}
function adminSetCmsAll(o) {
  try {
    var sh = cmsSheet_(SpreadsheetApp.openById(CMS_SHEET_ID)), rows = sh.getDataRange().getValues(), keys = ['overalls', 'first_place', 'students', 'years'];
    keys.forEach(function (key) {
      if (o[key] === undefined || o[key] === '') return;
      var done = false;
      for (var i = 1; i < rows.length; i++) if (String(rows[i][0]).trim() === key) { sh.getRange(i + 1, 2).setValue(o[key]); done = true; break; }
      if (!done) sh.appendRow([key, o[key]]);
    });
    return { ok: true, msg: 'Đã cập nhật số liệu web (web tự đổi sau ít phút).' };
  } catch (e) { return { ok: false, msg: 'Không ghi được sheet số liệu (cần cấp quyền EDIT cho ' + ADMIN_EMAIL + '): ' + e.message }; }
}

// ── Trang admin (HTML) ────────────────────────────────────
function adminPage_() {
  var g = gather_(), cms = cmsRead_(), st = g.st;
  var maxRev = 1; g.chart.forEach(function (d) { if (d.rev > maxRev) maxRev = d.rev; });
  function card(l, v, s) {
    return '<div style="background:#111;border:1px solid #1e1e1e;border-radius:14px;padding:16px 18px">'
      + '<div style="font-size:11px;letter-spacing:1px;text-transform:uppercase;color:#7a7a7a">' + l + '</div>'
      + '<div style="font-size:23px;font-weight:800;color:#fff;margin-top:6px">' + v + '</div>'
      + (s ? '<div style="font-size:12px;color:#8a7a52;margin-top:3px">' + s + '</div>' : '') + '</div>';
  }
  var cards = '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px">'
    + card('Doanh thu đã xác nhận', '<span style="color:#4cd98a">' + fmt_(st.revC) + 'đ</span>', st.nC + ' đơn')
    + card('Đang chờ', fmt_(st.revP) + 'đ', st.nP + ' đơn')
    + card('Hôm nay', '<span style="color:#D4A017">' + fmt_(st.revToday) + 'đ</span>', '')
    + card('7 ngày', fmt_(st.rev7) + 'đ', 'Tháng: ' + fmt_(st.revMonth) + 'đ') + '</div>';
  var bars = '';
  g.chart.forEach(function (d) {
    var h = Math.round((d.rev / maxRev) * 90); if (d.rev > 0 && h < 3) h = 3;
    bars += '<div style="flex:1;display:flex;flex-direction:column;align-items:center;gap:3px"><div title="' + d.day + ': ' + fmt_(d.rev) + 'đ" style="width:70%;height:' + h + 'px;background:' + (d.rev > 0 ? '#D4A017' : '#222') + ';border-radius:3px 3px 0 0"></div><div style="font-size:8px;color:#666">' + d.day.slice(8) + '/' + d.day.slice(5, 7) + '</div></div>';
  });
  var chart = '<div style="background:#111;border:1px solid #1e1e1e;border-radius:14px;padding:16px 18px;margin-top:14px"><div style="font-size:11px;letter-spacing:1px;text-transform:uppercase;color:#7a7a7a;margin-bottom:10px">Doanh thu 14 ngày</div><div style="display:flex;align-items:flex-end;gap:4px;height:110px">' + bars + '</div></div>';
  var topH = g.top.length ? '' : '<div style="color:#666;font-size:13px">Chưa có đơn.</div>';
  g.top.slice(0, 8).forEach(function (p) { topH += '<div style="display:flex;justify-content:space-between;gap:10px;padding:8px 0;border-bottom:1px solid #1a1a1a"><span style="color:#ddd;font-size:13px">' + esc_(p.name) + '</span><span style="color:#fff;font-size:13px;font-weight:700;white-space:nowrap">' + fmt_(p.rev) + 'đ <span style="color:#666;font-weight:400">· ' + p.n + '</span></span></div>'; });
  var topBox = '<div style="background:#111;border:1px solid #1e1e1e;border-radius:14px;padding:16px 18px;margin-top:14px"><div style="font-size:11px;letter-spacing:1px;text-transform:uppercase;color:#7a7a7a;margin-bottom:6px">Ebook bán chạy</div>' + topH + '</div>';

  var orderRows = '';
  g.orders.slice(0, 100).forEach(function (o) {
    var badge = o.confirmed ? '<span style="color:#4cd98a">✓ xác nhận</span>' : '<span style="color:#D4A017">chờ</span>';
    var act = o.pending
      ? '<button onclick="confirmOrder(\'' + esc_(o.code) + '\')" style="background:#D4A017;color:#000;border:none;border-radius:6px;padding:6px 10px;font-weight:700;font-size:11px;cursor:pointer">✅ Xác nhận &amp; gửi</button>'
      : '<button onclick="resendOrder(\'' + esc_(o.code) + '\')" style="background:#1c1c1c;color:#D4A017;border:1px solid #333;border-radius:6px;padding:6px 10px;font-weight:700;font-size:11px;cursor:pointer">📩 Gửi lại</button>';
    orderRows += '<tr data-st="' + (o.pending ? 'pending' : 'confirmed') + '">'
      + '<td style="padding:8px 6px;border-bottom:1px solid #161616;color:#999;font-size:11px;white-space:nowrap">' + esc_(o.time) + '</td>'
      + '<td style="padding:8px 6px;border-bottom:1px solid #161616;color:#fff;font-size:12px">' + esc_(o.name) + '<br><span style="color:#777;font-size:11px">' + esc_(o.email) + '</span></td>'
      + '<td style="padding:8px 6px;border-bottom:1px solid #161616;color:#bbb;font-size:11px">' + esc_(o.items) + '</td>'
      + '<td style="padding:8px 6px;border-bottom:1px solid #161616;color:#fff;font-size:12px;text-align:right;white-space:nowrap">' + fmt_(o.total) + 'đ<br>' + badge + '</td>'
      + '<td style="padding:8px 6px;border-bottom:1px solid #161616;text-align:right">' + act + '</td></tr>';
  });
  var ordersBox = '<div style="background:#111;border:1px solid #1e1e1e;border-radius:14px;padding:14px;margin-top:14px;overflow-x:auto">'
    + '<div style="display:flex;gap:8px;margin-bottom:10px;flex-wrap:wrap"><b style="color:#fff;font-size:14px;margin-right:auto">Đơn hàng</b>'
    + '<button onclick="filt(\'all\')" style="background:#1c1c1c;color:#ccc;border:1px solid #333;border-radius:6px;padding:5px 12px;font-size:11px;cursor:pointer">Tất cả</button>'
    + '<button onclick="filt(\'pending\')" style="background:#1c1c1c;color:#D4A017;border:1px solid #333;border-radius:6px;padding:5px 12px;font-size:11px;cursor:pointer">Chờ</button>'
    + '<button onclick="filt(\'confirmed\')" style="background:#1c1c1c;color:#4cd98a;border:1px solid #333;border-radius:6px;padding:5px 12px;font-size:11px;cursor:pointer">Đã xác nhận</button></div>'
    + '<table style="width:100%;border-collapse:collapse;min-width:560px">' + orderRows + '</table></div>';

  var linkRows = '', links = linkSheet_().getDataRange().getValues(), linkNames = {};
  for (var i = 1; i < links.length; i++) {
    linkNames[i] = links[i][0];
    linkRows += '<div style="padding:10px 0;border-bottom:1px solid #1a1a1a"><div style="color:#ddd;font-size:13px;margin-bottom:6px">' + esc_(links[i][0]) + '</div><div style="display:flex;gap:8px"><input id="lk_' + i + '" value="' + esc_(links[i][1]) + '" style="flex:1;background:#0a0a0a;border:1px solid #333;border-radius:6px;color:#fff;padding:8px;font-size:11px"><button onclick="saveLink(' + i + ')" style="background:#D4A017;color:#000;border:none;border-radius:6px;padding:0 14px;font-weight:700;font-size:12px;cursor:pointer">Lưu</button></div></div>';
  }
  var linksBox = '<div style="background:#111;border:1px solid #1e1e1e;border-radius:14px;padding:16px 18px;margin-top:14px"><div style="font-size:11px;letter-spacing:1px;text-transform:uppercase;color:#7a7a7a;margin-bottom:6px">Link ebook (đổi khỏi cần sửa code)</div>' + linkRows + '</div>';

  function cmsInput(id, label, v) { return '<div style="margin-bottom:10px"><div style="font-size:12px;color:#8a7a52;margin-bottom:4px">' + label + '</div><input id="' + id + '" value="' + esc_(v) + '" style="width:100%;box-sizing:border-box;background:#0a0a0a;border:1px solid #333;border-radius:6px;color:#fff;padding:9px;font-size:14px"></div>'; }
  var cmsErr = cms._err ? '<div style="color:#ff9b9b;font-size:11px;margin-bottom:8px">⚠️ ' + esc_(cms._err) + ' (cấp quyền Editor sheet số liệu cho ' + ADMIN_EMAIL + ')</div>' : '';
  var cmsBox = '<div style="background:#111;border:1px solid #1e1e1e;border-radius:14px;padding:16px 18px;margin-top:14px"><div style="font-size:11px;letter-spacing:1px;text-transform:uppercase;color:#7a7a7a;margin-bottom:10px">Số liệu hiển thị trên web</div>' + cmsErr
    + cmsInput('cms_overalls', 'Overall Champion (overalls)', cms.overalls) + cmsInput('cms_first_place', '1st Place (first_place)', cms.first_place)
    + cmsInput('cms_students', 'Học viên (students)', cms.students) + cmsInput('cms_years', 'Số năm (years)', cms.years)
    + '<button onclick="saveCms()" style="background:#D4A017;color:#000;border:none;border-radius:8px;padding:11px 20px;font-weight:800;font-size:13px;cursor:pointer">Lưu số liệu</button></div>';

  function tb(id, txt) { return '<button onclick="tab(\'' + id + '\')" id="t_' + id + '" style="background:none;border:none;color:#888;font-size:13px;font-weight:700;padding:10px 6px;cursor:pointer;border-bottom:2px solid transparent">' + txt + '</button>'; }
  var body = '<div style="font-family:-apple-system,Arial,sans-serif;background:#0a0a0a;min-height:100vh;color:#fff;padding:16px"><div style="max-width:920px;margin:0 auto">'
    + '<div style="font-size:19px;font-weight:900;letter-spacing:1px;margin-bottom:4px">KUMACOACH <span style="color:#D4A017">· ADMIN</span></div>'
    + '<div style="display:flex;gap:14px;border-bottom:1px solid #1e1e1e;margin-bottom:14px;flex-wrap:wrap">' + tb('stats', '📊 Thống kê') + tb('orders', '🛒 Đơn hàng') + tb('links', '🔗 Link ebook') + tb('cms', '⚙️ Số liệu web') + '</div>'
    + '<div id="stats">' + cards + chart + topBox + '</div>'
    + '<div id="orders" style="display:none">' + ordersBox + '</div>'
    + '<div id="links" style="display:none">' + linksBox + '</div>'
    + '<div id="cms" style="display:none">' + cmsBox + '</div></div></div>';
  var js = '<script>var LN=' + JSON.stringify(linkNames) + ';'
    + 'function tab(id){["stats","orders","links","cms"].forEach(function(x){document.getElementById(x).style.display=(x===id)?"":"none";var b=document.getElementById("t_"+x);b.style.color=(x===id)?"#D4A017":"#888";b.style.borderBottomColor=(x===id)?"#D4A017":"transparent";});}'
    + 'function call(fn,args){google.script.run.withSuccessHandler(function(r){if(r&&r.msg)alert(r.msg);location.reload();}).withFailureHandler(function(e){alert("Lỗi: "+e.message);})[fn].apply(google.script.run,args);}'
    + 'function confirmOrder(c){if(confirm("Xác nhận & gửi ebook cho đơn "+c+"?"))call("adminConfirm",[c]);}'
    + 'function resendOrder(c){if(confirm("Gửi LẠI ebook cho đơn "+c+"?"))call("adminResend",[c]);}'
    + 'function saveLink(i){var p=LN[i];var u=document.getElementById("lk_"+i).value.trim();if(u&&p)call("adminSetLink",[p,u]);}'
    + 'function saveCms(){call("adminSetCmsAll",[{overalls:v("cms_overalls"),first_place:v("cms_first_place"),students:v("cms_students"),years:v("cms_years")}]);}'
    + 'function v(id){return document.getElementById(id).value.trim();}'
    + 'function filt(s){var r=document.querySelectorAll("[data-st]");for(var i=0;i<r.length;i++){r[i].style.display=(s==="all"||r[i].getAttribute("data-st")===s)?"":"none";}}'
    + 'tab("stats");</script>';
  return HtmlService.createHtmlOutput(body + js).setTitle('KumaCoach Admin').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
