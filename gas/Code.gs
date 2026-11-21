/**
 * BAHULU KUNDUR · KOMASKO — Backend App Konsainan (v2: Staff + Admin)
 * Google Apps Script + Google Sheets
 *
 * CARA PASANG / KEMAS KINI:
 * 1. Extensions > Apps Script, padam kod lama, tampal SEMUA kod ini, Save
 * 2. Pilih fungsi "setup" di bar atas, tekan Run (selamat dijalankan semula, data tak terpadam)
 * 3. Deploy > Manage deployments > pensel (Edit) > Version: New version > Deploy
 *    (URL kekal sama, app tak perlu diubah)
 *
 * TELEGRAM (notifikasi admin): lihat panduan di bahagian TELEGRAM di bawah.
 *  1. Buat bot di @BotFather, salin token ke sheet Tetapan baris tg_token
 *  2. Admin tekan Start pada bot
 *  3. Run fungsi sambungTelegram  (pilihan: Run pasangRingkasanHarian)
 *
 * PIN:
 *  - pin_admin : semua akses (edit, padam, tetapan, laporan). Min 6 digit.
 *  - pin_staff : hantar stok, kutip, cetak sahaja. Min 4 digit.
 *  Kalau sheet Tetapan ada PIN lama ("pin"), ia jadi PIN admin.
 *  PIN asal baru: admin 123456, staff 1234 — TUKAR dalam app (Tetapan > Keselamatan).
 */

const SS_ID = ''; // kosongkan jika skrip ini dibuka dari dalam Google Sheet itu sendiri

const SHEETS = {
  Tetapan: ['key', 'value'],
  Produk:  ['id', 'name', 'price', 'order'],
  Kedai:   ['id', 'name', 'phone', 'updatedAt'],
  Resit:   ['id', 'no', 'shopId', 'date', 'status', 'items', 'settle', 'updatedAt', 'deleted'],
  Log:     ['masa', 'peranan', 'tindakan', 'rujukan', 'butiran']
};

function ss_() { return SS_ID ? SpreadsheetApp.openById(SS_ID) : SpreadsheetApp.getActiveSpreadsheet(); }

/* ---------- SETUP (selamat dijalankan semula) ---------- */
function setup() {
  const ss = ss_();
  Object.keys(SHEETS).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const head = SHEETS[name];
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#7A3F14').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), head.length).setNumberFormat('@');
  });
  const s = readSettings_();
  if (!s.name) setSetting_('name', 'Bahulu Kundur');
  if (s.sub === undefined) setSetting_('sub', 'Keluaran KOMASKO');
  if (s.phone === undefined) setSetting_('phone', '');
  if (!s.note) setSetting_('note', 'Terima kasih. Bayaran ikut jualan sebenar.');
  if (!s.pin_admin) setSetting_('pin_admin', s.pin || '123456');
  if (!s.pin_staff) setSetting_('pin_staff', '1234');
  if (s.tg_token === undefined) setSetting_('tg_token', '');
  if (s.tg_admin === undefined) setSetting_('tg_admin', 'abufayyadh');
  if (s.tg_chat === undefined) setSetting_('tg_chat', '');
  const p = ss.getSheetByName('Produk');
  if (p.getLastRow() < 2) {
    p.getRange(2, 1, 2, 4).setValues([
      ['p1', 'Bahulu Ori (balang)', '12', '1'],
      ['p2', 'Bahulu Perisa (balang)', '15', '2']
    ]);
  }
  const s1 = ss.getSheetByName('Sheet1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 5) ss.deleteSheet(s1);
  return 'Setup siap';
}

/* ---------- API ---------- */
function doGet(e) { return handle_(e.parameter || {}); }
function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents || '{}'); } catch (err) { return out_({ ok: false, error: 'Data tidak sah' }); }
  return handle_(p);
}

function handle_(p) {
  try {
    const a = auth_(p);
    if (a.err) return out_({ ok: false, error: a.err, code: a.code });
    if (p.action === 'login') {
      log_(a.role, 'Log masuk', '', 'Device ' + String(p.dev || '').slice(-6));
      return out_({ ok: true, role: a.role, data: readAll_(), time: Date.now() });
    }
    if (p.action === 'all') return out_({ ok: true, role: a.role, data: readAll_(), time: Date.now() });
    if (p.action === 'sync') {
      const r = applyOps_(p.ops || [], a.role);
      r.notes.forEach(function (m) { tgSend_(m); });
      return out_({ ok: true, role: a.role, applied: r.applied, rejected: r.rejected, data: readAll_(), time: Date.now() });
    }
    return out_({ ok: false, error: 'Tindakan tidak dikenali' });
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- PIN + had cubaan ---------- */
function auth_(p) {
  const s = readSettings_();
  const cache = CacheService.getScriptCache();
  const dev = String(p.dev || 'tiada').slice(0, 40);
  const kd = 'fail_' + dev, kg = 'fail_semua';
  const fd = +cache.get(kd) || 0, fg = +cache.get(kg) || 0;
  if (fd >= 5 || fg >= 30) return { err: 'Terlalu banyak cubaan PIN salah. Cuba lagi selepas 15 minit.', code: 'LOCK' };
  const pin = String(p.pin || '');
  const pa = String(s.pin_admin || s.pin || ''), ps = String(s.pin_staff || '');
  let role = null;
  if (pin && pin === pa) role = 'admin';
  else if (pin && ps && pin === ps) role = 'staff';
  if (!role) {
    cache.put(kd, String(fd + 1), 900);
    cache.put(kg, String(fg + 1), 900);
    if (p.action === 'login') log_('-', 'PIN salah', '', 'Device ' + dev.slice(-6) + ' (cubaan ' + (fd + 1) + ')');
    return { err: 'PIN salah', code: 'PIN' };
  }
  return { role: role };
}

/* ---------- BACA ---------- */
function rows_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const head = SHEETS[name];
  return sh.getRange(2, 1, sh.getLastRow() - 1, head.length).getValues()
    .filter(function (r) { return String(r[0]) !== ''; })
    .map(function (r) { const o = {}; head.forEach(function (h, i) { o[h] = r[i]; }); return o; });
}

function readSettings_() {
  const o = {};
  rows_('Tetapan').forEach(function (r) { o[String(r.key)] = String(r.value); });
  return o;
}

function readProducts_() {
  return rows_('Produk')
    .sort(function (a, b) { return (+a.order || 0) - (+b.order || 0); })
    .map(function (r) { return { id: String(r.id), name: String(r.name), price: +r.price || 0 }; });
}

function toCons_(r) {
  return {
    id: String(r.id), no: String(r.no), shopId: String(r.shopId), date: String(r.date), status: String(r.status),
    items: parse_(r.items, []), settle: parse_(r.settle, null), updatedAt: +r.updatedAt || 0
  };
}

function readAll_() {
  const s = readSettings_();
  const shops = rows_('Kedai').map(function (r) {
    return { id: String(r.id), name: String(r.name), phone: String(r.phone || ''), updatedAt: +r.updatedAt || 0 };
  });
  const cons = rows_('Resit')
    .filter(function (r) { return String(r.deleted) !== 'true'; })
    .map(toCons_)
    .sort(function (a, b) { return (b.updatedAt || 0) - (a.updatedAt || 0); });
  return {
    supplier: { name: s.name || '', sub: s.sub || '', phone: s.phone || '', note: s.note || '' },
    products: readProducts_(), shops: shops, cons: cons
  };
}

function parse_(v, d) { try { return v ? JSON.parse(v) : d; } catch (e) { return d; } }

/* ---------- TULIS (dengan semakan peranan) ---------- */
function applyOps_(ops, role) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    let applied = 0;
    const rejected = [];
    const notes = [];
    const who = admin_(role);
    const admin = role === 'admin';
    const products = readProducts_();
    const shopName = function (id) {
      const r = rows_('Kedai').filter(function (x) { return String(x.id) === String(id); })[0];
      return r ? String(r.name) : '?';
    };

    const reject = function (what, ref) { rejected.push(what); log_(role, 'Ditolak (perlu Admin)', ref || '', what); };
    ops.forEach(function (op) {
      const d = op.data;
      if (op.t === 'supplier' || op.t === 'products' || op.t === 'pin') {
        if (!admin) { reject(op.t); return; }
        if (op.t === 'supplier') {
          ['name', 'sub', 'phone', 'note'].forEach(function (k) { setSetting_(k, d[k] || ''); });
          log_(role, 'Tetapan diubah', '', 'Nama/telefon/nota resit');
        } else if (op.t === 'products') {
          const before = products.map(function (x) { return x.name + ' RM' + x.price; }).join(', ');
          writeProducts_(d || []);
          const after = (d || []).map(function (x) { return x.name + ' RM' + (+x.price || 0); }).join(', ');
          if (before !== after) log_(role, 'Produk/harga diubah', '', 'Sebelum: ' + before + ' | Selepas: ' + after);
        } else {
          if (d.admin) { if (!/^\d{6,8}$/.test(String(d.admin))) { rejected.push('pin'); return; } setSetting_('pin_admin', String(d.admin)); log_(role, 'PIN admin ditukar', '', ''); }
          if (d.staff) { if (!/^\d{4,8}$/.test(String(d.staff))) { rejected.push('pin'); return; } setSetting_('pin_staff', String(d.staff)); log_(role, 'PIN staff ditukar', '', ''); }
        }
        applied++;
        return;
      }

      if (op.t === 'shop') {
        const ex = find_('Kedai', d.id);
        if (ex && !admin) { reject('ubah kedai', d.name); return; }
        if (upsert_('Kedai', d.id, [d.id, d.name || '', d.phone || '', String(d.updatedAt || Date.now())], d.updatedAt, 3)) {
          applied++;
          log_(role, ex ? 'Kedai diubah' : 'Kedai baru', '', (ex ? String(ex.obj.name) + ' → ' : '') + (d.name || ''));
        }
        return;
      }

      if (op.t === 'cons') {
        const ex = find_('Resit', d.id);
        const old = ex && String(ex.obj.deleted) !== 'true' ? toCons_(ex.obj) : null;

        if (d.deleted) {
          if (!admin) { reject('padam resit', d.no); return; }
          if (!old) return;
          writeCons_(d);
          notes.push('🗑 <b>RESIT DIPADAM</b>\nNo: ' + esc_(old.no) + '\nKedai: ' + esc_(shopName(old.shopId)) + '\nOleh: ' + who + '\nMasa: ' + time_(Date.now()));
          log_(role, 'Resit dipadam', old.no, shopName(old.shopId) + ' · ' + (old.status === 'selesai' ? 'dah bayar ' + rm_(old.settle && old.settle.total) : 'belum kutip ' + itemsTxt_(old.items)));
          applied++;
          return;
        }

        // kira semula jumlah bayaran di server (jangan percaya nombor dari app)
        if (d.status === 'selesai' && d.settle) {
          if (!admin) {
            d.settle.prices = {};
            (d.items || []).forEach(function (i) {
              const p = products.filter(function (x) { return x.id === i.pid; })[0];
              d.settle.prices[i.pid] = p ? p.price : +i.price || 0;
            });
          }
          d.settle.total = totalOf_(d);
        }

        if (!old) {
          if (!admin && d.status !== 'aktif' && d.status !== 'selesai') { reject('resit', d.no); return; }
          writeCons_(d);
          log_(role, 'Hantar stok', d.no, shopName(d.shopId) + ': ' + itemsTxt_(d.items) + ' · nilai ' + rm_(valueOf_(d.items)));
          if (d.status === 'selesai') log_(role, 'Kutip', d.no, soldTxt_(d) + ' · bayar ' + rm_(d.settle.total) + ' · tarikh ' + d.settle.date);
          notes.push(msgHantar_(d, shopName(d.shopId), who));
          if (d.status === 'selesai') notes.push({ kutip: d, shop: shopName(d.shopId), who: who, edit: false });
          applied++;
          return;
        }

        if ((+old.updatedAt || 0) > (+d.updatedAt || 0)) return; // data dalam sheet lebih baru

        const sameItems = JSON.stringify(old.items) === JSON.stringify(d.items);
        if (!admin) {
          // staff hanya boleh kutip resit yang belum kutip, tanpa ubah item hantaran
          if (!(old.status === 'aktif' && d.status === 'selesai' && sameItems)) { reject('edit resit', d.no); return; }
        }
        writeCons_(d);
        applied++;

        if (!sameItems) log_(role, 'Edit hantaran', d.no, 'Sebelum: ' + itemsTxt_(old.items) + ' | Selepas: ' + itemsTxt_(d.items));
        if (old.status === 'aktif' && d.status === 'selesai') {
          const was = old.settle ? ' (sebelum ini ' + rm_(old.settle.total) + ')' : '';
          notes.push({ kutip: d, shop: shopName(d.shopId), who: who, edit: !!old.settle, before: old.settle ? old.settle.total : null });
          log_(role, old.settle ? 'Edit kutipan' : 'Kutip', d.no, soldTxt_(d) + ' · bayar ' + rm_(d.settle.total) + was + ' · tarikh ' + d.settle.date);
        } else if (old.status === 'selesai' && d.status === 'aktif') {
          log_(role, 'Buka semula kutipan', d.no, 'Untuk edit · sebelum ini ' + rm_(old.settle && old.settle.total));
        } else if (old.status === 'selesai' && d.status === 'selesai' && JSON.stringify(old.settle) !== JSON.stringify(d.settle)) {
          log_(role, 'Edit kutipan', d.no, 'Sebelum: ' + soldTxt_(old) + ' ' + rm_(old.settle.total) + ' | Selepas: ' + soldTxt_(d) + ' ' + rm_(d.settle.total));
        }
      }
    });
    // mesej kutipan dibina selepas semua data ditulis supaya "jumlah hari ini" tepat
    return {
      applied: applied, rejected: rejected,
      notes: notes.map(function (n) { return typeof n === 'string' ? n : msgKutip_(n); })
    };
  } finally {
    lock.releaseLock();
  }
}

function writeCons_(d) {
  const row = [d.id, d.no || '', d.shopId || '', d.date || '', d.status || 'aktif',
    JSON.stringify(d.items || []), d.settle ? JSON.stringify(d.settle) : '',
    String(d.updatedAt || Date.now()), d.deleted ? 'true' : ''];
  upsert_('Resit', d.id, row, d.updatedAt, 7);
}

/* ---------- Bantuan teks log ---------- */
function rm_(n) { return 'RM' + (Number(n) || 0).toFixed(2); }
function valueOf_(items) { return (items || []).reduce(function (s, i) { return s + (+i.qty || 0) * (+i.price || 0); }, 0); }
function itemsTxt_(items) { return (items || []).map(function (i) { return i.name + ' ' + i.qty; }).join(', '); }
function priceOf_(c, i) { const s = c.settle; return s && s.prices && s.prices[i.pid] != null ? +s.prices[i.pid] : +i.price || 0; }
function totalOf_(c) {
  let t = 0;
  (c.items || []).forEach(function (i) {
    const b = Math.min(+i.qty || 0, Math.max(0, +(c.settle.baki || {})[i.pid] || 0));
    t += ((+i.qty || 0) - b) * priceOf_(c, i);
  });
  return Math.round(t * 100) / 100;
}
function soldTxt_(c) {
  if (!c.settle) return '';
  return 'Terjual ' + (c.items || []).map(function (i) {
    const b = +(c.settle.baki || {})[i.pid] || 0;
    return i.name + ' ' + ((+i.qty || 0) - b) + '/' + i.qty + ' @RM' + priceOf_(c, i);
  }).join(', ');
}

function log_(role, action, ref, detail) {
  const sh = ss_().getSheetByName('Log');
  if (!sh) return;
  const t = Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'yyyy-MM-dd HH:mm:ss');
  sh.getRange(sh.getLastRow() + 1, 1, 1, 5).setNumberFormat('@').setValues([[t, role === 'admin' ? 'Admin' : role === 'staff' ? 'Staff' : '-', action, ref || '', detail || '']]);
}

/* ---------- Tulis asas ---------- */
function setSetting_(key, val) {
  const sh = ss_().getSheetByName('Tetapan');
  const last = sh.getLastRow();
  const keys = last >= 2 ? sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]); }) : [];
  const i = keys.indexOf(key);
  if (i >= 0) sh.getRange(i + 2, 2).setValue(String(val));
  else sh.getRange(last + 1, 1, 1, 2).setNumberFormat('@').setValues([[key, String(val)]]);
}

function writeProducts_(list) {
  const sh = ss_().getSheetByName('Produk');
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 4).clearContent();
  if (!list.length) return;
  sh.getRange(2, 1, list.length, 4).setValues(list.map(function (p, i) {
    return [String(p.id), String(p.name || ''), String(+p.price || 0), String(i + 1)];
  }));
}

function find_(name, id) {
  const sh = ss_().getSheetByName(name);
  const last = sh.getLastRow();
  if (last < 2) return null;
  const head = SHEETS[name];
  const data = sh.getRange(2, 1, last - 1, head.length).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      const o = {}; head.forEach(function (h, k) { o[h] = data[i][k]; });
      return { row: i + 2, obj: o };
    }
  }
  return null;
}

/* Kemaskini baris ikut id; hanya jika data baru lebih terkini (updatedAt) */
function upsert_(name, id, row, updatedAt, updCol) {
  const sh = ss_().getSheetByName(name);
  const width = SHEETS[name].length;
  const ex = find_(name, id);
  if (ex) {
    if ((+ex.obj[SHEETS[name][updCol]] || 0) > (+updatedAt || 0)) return false;
    sh.getRange(ex.row, 1, 1, width).setValues([row.map(String)]);
    return true;
  }
  sh.getRange(sh.getLastRow() + 1, 1, 1, width).setNumberFormat('@').setValues([row.map(String)]);
  return true;
}

/* ================== TELEGRAM ==================
 * PANDUAN:
 * 1. Di Telegram, buka @BotFather > /newbot > beri nama (contoh: Bahulu Kundur Notifikasi)
 *    dan username bot (mesti berakhir "bot", contoh: bahulukundur_notif_bot). Salin TOKEN.
 * 2. Tampal token dalam sheet Tetapan, baris "tg_token" (lajur value).
 * 3. Admin (username dalam baris "tg_admin") buka bot tadi dan tekan START.
 * 4. Di Apps Script, pilih fungsi "sambungTelegram" > Run. Admin akan dapat mesej ujian.
 * 5. (Pilihan) Run "pasangRingkasanHarian" untuk ringkasan setiap hari jam 9 malam.
 * 6. Deploy > Manage deployments > Edit > New version > Deploy.
 */
function sambungTelegram() {
  const s = readSettings_();
  const token = String(s.tg_token || '').trim();
  if (!token) throw new Error('Isi tg_token dalam sheet Tetapan dulu (token dari @BotFather).');
  const want = String(s.tg_admin || '').replace('@', '').toLowerCase();
  const res = JSON.parse(UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/getUpdates', { muteHttpExceptions: true }).getContentText());
  if (!res.ok) throw new Error('Token tidak sah: ' + (res.description || ''));
  let chat = null;
  (res.result || []).forEach(function (u) {
    const m = u.message || u.my_chat_member || u.edited_message;
    if (!m || !m.chat) return;
    const un = String((m.from && m.from.username) || m.chat.username || '').toLowerCase();
    if (!want || un === want || m.chat.type !== 'private') chat = m.chat;
  });
  if (!chat) throw new Error('Tak jumpa chat @' + want + '. Pastikan admin dah buka bot dan tekan START (atau hantar apa-apa mesej), kemudian Run semula.');
  setSetting_('tg_chat', String(chat.id));
  tgSend_('✅ <b>Notifikasi Bahulu Kundur aktif</b>\nAnda akan terima makluman setiap kali staff hantar stok atau kutip bayaran di kedai.');
  return 'Berjaya. Chat ID: ' + chat.id + ' (' + (chat.username || chat.title || '') + ')';
}

function tgSend_(text) {
  try {
    const s = readSettings_();
    if (!s.tg_token || !s.tg_chat) return;
    UrlFetchApp.fetch('https://api.telegram.org/bot' + s.tg_token + '/sendMessage', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ chat_id: s.tg_chat, text: text, parse_mode: 'HTML', disable_web_page_preview: true })
    });
  } catch (e) { /* notifikasi gagal tak boleh ganggu simpanan data */ }
}

function esc_(t) { return String(t == null ? '' : t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function admin_(role) { return role === 'admin' ? 'Admin' : role === 'staff' ? 'Staff' : '-'; }
function time_(ms) { return Utilities.formatDate(new Date(+ms || Date.now()), 'Asia/Kuala_Lumpur', 'dd/MM/yyyy, h:mm a'); }
function lateNote_(ms) { const m = Math.round((Date.now() - (+ms || Date.now())) / 60000); return m > 10 ? '\n<i>(direkod offline, sampai ke sistem ' + m + ' minit kemudian)</i>' : ''; }
function dmy_(ymd) { return String(ymd || '').split('-').reverse().join('/'); }

function msgHantar_(d, shop, who) {
  const u = (d.items || []).reduce(function (a, i) { return a + (+i.qty || 0); }, 0);
  return '🚚 <b>HANTAR STOK</b>\n' +
    '🏪 <b>' + esc_(shop) + '</b>\n' +
    '🕒 ' + time_(d.updatedAt) + '\n' +
    '👤 ' + who + ' · No ' + esc_(d.no) + '\n\n' +
    (d.items || []).map(function (i) { return '• ' + esc_(i.name) + ' × <b>' + i.qty + '</b> (' + rm_(i.qty * i.price) + ')'; }).join('\n') +
    '\n\nJumlah: <b>' + u + ' balang</b> · Nilai ' + rm_(valueOf_(d.items)) + lateNote_(d.updatedAt);
}

function msgKutip_(n) {
  const d = n.kutip;
  let jual = 0, baki = 0;
  const lines = (d.items || []).map(function (i) {
    const b = +(d.settle.baki || {})[i.pid] || 0, j = (+i.qty || 0) - b;
    jual += j; baki += b;
    return '• ' + esc_(i.name) + ': terjual <b>' + j + '</b>/' + i.qty + ', baki ' + b + ' → ' + rm_(j * priceOf_(d, i));
  });
  const td = todaySummary_(d.settle.date);
  return (n.edit ? '✏️ <b>KUTIPAN DIUBAH</b>' : '💰 <b>KUTIPAN</b>') + '\n' +
    '🏪 <b>' + esc_(n.shop) + '</b>\n' +
    '🕒 ' + time_(d.updatedAt) + (d.settle.date !== Utilities.formatDate(new Date(+d.updatedAt || Date.now()), 'Asia/Kuala_Lumpur', 'yyyy-MM-dd') ? ' (tarikh kutip ' + dmy_(d.settle.date) + ')' : '') + '\n' +
    '👤 ' + n.who + ' · No ' + esc_(d.no) + ' · hantar ' + dmy_(d.date) + '\n\n' +
    lines.join('\n') + '\n\n' +
    'Terjual <b>' + jual + '</b> balang · Baki di kedai <b>' + baki + '</b> balang\n' +
    '💵 Kutipan: <b>' + rm_(d.settle.total) + '</b>' + (n.edit && n.before != null ? ' (sebelum ini ' + rm_(n.before) + ')' : '') + '\n' +
    '📊 Jumlah kutipan ' + dmy_(d.settle.date) + ': <b>' + rm_(td.total) + '</b> (' + td.n + ' kedai)' + lateNote_(d.updatedAt);
}

function todaySummary_(ymd) {
  const list = rows_('Resit').filter(function (r) { return String(r.deleted) !== 'true' && String(r.status) === 'selesai'; })
    .map(toCons_).filter(function (c) { return c.settle && c.settle.date === ymd; });
  return { n: list.length, total: list.reduce(function (a, c) { return a + (+c.settle.total || 0); }, 0), list: list };
}

/* Ringkasan harian (pilihan) */
function ringkasanHarian() {
  const ymd = Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'yyyy-MM-dd');
  const all = rows_('Resit').filter(function (r) { return String(r.deleted) !== 'true'; }).map(toCons_);
  const shops = {}; rows_('Kedai').forEach(function (k) { shops[String(k.id)] = String(k.name); });
  const hantar = all.filter(function (c) { return c.date === ymd; });
  const kutip = all.filter(function (c) { return c.status === 'selesai' && c.settle && c.settle.date === ymd; });
  const aktif = all.filter(function (c) { return c.status === 'aktif'; });
  const tot = kutip.reduce(function (a, c) { return a + (+c.settle.total || 0); }, 0);
  const balHantar = hantar.reduce(function (a, c) { return a + c.items.reduce(function (s, i) { return s + (+i.qty || 0); }, 0); }, 0);
  const stok = aktif.reduce(function (a, c) { return a + valueOf_(c.items); }, 0);
  tgSend_('📋 <b>RINGKASAN HARI INI</b> (' + dmy_(ymd) + ')\n\n' +
    '🚚 Hantar: <b>' + hantar.length + '</b> kedai, ' + balHantar + ' balang\n' +
    '💰 Kutip: <b>' + kutip.length + '</b> kedai, <b>' + rm_(tot) + '</b>\n' +
    (kutip.length ? kutip.map(function (c) { return '   • ' + esc_(shops[c.shopId] || '?') + ': ' + rm_(c.settle.total); }).join('\n') + '\n' : '') +
    '\n📦 Stok masih di kedai: ' + rm_(stok) + ' (' + aktif.length + ' resit belum kutip)');
}

function pasangRingkasanHarian() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'ringkasanHarian') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('ringkasanHarian').timeBased().everyDays(1).atHour(21).inTimezone('Asia/Kuala_Lumpur').create();
  return 'Ringkasan harian dipasang: setiap hari jam 9 malam';
}
