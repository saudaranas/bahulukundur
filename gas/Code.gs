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
          log_(role, old.settle ? 'Edit kutipan' : 'Kutip', d.no, soldTxt_(d) + ' · bayar ' + rm_(d.settle.total) + was + ' · tarikh ' + d.settle.date);
        } else if (old.status === 'selesai' && d.status === 'aktif') {
          log_(role, 'Buka semula kutipan', d.no, 'Untuk edit · sebelum ini ' + rm_(old.settle && old.settle.total));
        } else if (old.status === 'selesai' && d.status === 'selesai' && JSON.stringify(old.settle) !== JSON.stringify(d.settle)) {
          log_(role, 'Edit kutipan', d.no, 'Sebelum: ' + soldTxt_(old) + ' ' + rm_(old.settle.total) + ' | Selepas: ' + soldTxt_(d) + ' ' + rm_(d.settle.total));
        }
      }
    });
    return { applied: applied, rejected: rejected };
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
