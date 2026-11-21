/**
 * BAHULU KUNDUR · KOMASKO — Backend App Konsainan
 * Google Apps Script + Google Sheets
 *
 * CARA PASANG:
 * 1. Buat Google Sheet baru (contoh nama: "Bahulu Kundur - Data App")
 * 2. Extensions > Apps Script, padam kod asal, tampal SEMUA kod ini
 * 3. Pilih fungsi "setup" di bar atas, tekan Run (benarkan akses bila diminta)
 * 4. Deploy > New deployment > Web app
 *      Execute as: Me
 *      Who has access: Anyone
 * 5. Salin URL Web App (berakhir dengan /exec) dan beri kepada developer
 *
 * PIN asal: 1234 (tukar dalam app: Tetapan > Keselamatan, atau terus di sheet Tetapan)
 */

const SS_ID = ''; // kosongkan jika skrip ini dibuka dari dalam Google Sheet itu sendiri

const SHEETS = {
  Tetapan: ['key', 'value'],
  Produk:  ['id', 'name', 'price', 'order'],
  Kedai:   ['id', 'name', 'phone', 'updatedAt'],
  Resit:   ['id', 'no', 'shopId', 'date', 'status', 'items', 'settle', 'updatedAt', 'deleted']
};

function ss_() { return SS_ID ? SpreadsheetApp.openById(SS_ID) : SpreadsheetApp.getActiveSpreadsheet(); }

/* ---------- SETUP (jalankan sekali) ---------- */
function setup() {
  const ss = ss_();
  Object.keys(SHEETS).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const head = SHEETS[name];
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#7A3F14').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    // semua lajur sebagai teks supaya tarikh, no. telefon & no. resit tak ditukar oleh Sheets
    sh.getRange(1, 1, sh.getMaxRows(), head.length).setNumberFormat('@');
  });
  const t = ss.getSheetByName('Tetapan');
  if (t.getLastRow() < 2) {
    t.getRange(2, 1, 5, 2).setValues([
      ['pin', '1234'],
      ['name', 'Bahulu Kundur'],
      ['sub', 'Keluaran KOMASKO'],
      ['phone', ''],
      ['note', 'Terima kasih. Bayaran ikut jualan sebenar.']
    ]);
  }
  const p = ss.getSheetByName('Produk');
  if (p.getLastRow() < 2) {
    p.getRange(2, 1, 2, 4).setValues([
      ['p1', 'Bahulu Ori (balang)', '12', '1'],
      ['p2', 'Bahulu Perisa (balang)', '15', '2']
    ]);
  }
  const s1 = ss.getSheetByName('Sheet1');
  if (s1 && ss.getSheets().length > 4 && s1.getLastRow() === 0) ss.deleteSheet(s1);
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
    const set = readSettings_();
    if (String(p.pin || '') !== String(set.pin || '')) return out_({ ok: false, error: 'PIN salah', code: 'PIN' });
    if (p.action === 'all') return out_({ ok: true, data: readAll_(), time: Date.now() });
    if (p.action === 'sync') {
      const n = applyOps_(p.ops || []);
      return out_({ ok: true, applied: n, data: readAll_(), time: Date.now() });
    }
    return out_({ ok: false, error: 'Tindakan tidak dikenali' });
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

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
  rows_('Tetapan').forEach(function (r) { o[r.key] = String(r.value); });
  return o;
}

function readAll_() {
  const s = readSettings_();
  const products = rows_('Produk')
    .sort(function (a, b) { return (+a.order || 0) - (+b.order || 0); })
    .map(function (r) { return { id: String(r.id), name: String(r.name), price: +r.price || 0 }; });
  const shops = rows_('Kedai').map(function (r) {
    return { id: String(r.id), name: String(r.name), phone: String(r.phone || ''), updatedAt: +r.updatedAt || 0 };
  });
  const cons = rows_('Resit')
    .filter(function (r) { return String(r.deleted) !== 'true'; })
    .map(function (r) {
      return {
        id: String(r.id), no: String(r.no), shopId: String(r.shopId), date: String(r.date), status: String(r.status),
        items: parse_(r.items, []), settle: parse_(r.settle, null), updatedAt: +r.updatedAt || 0
      };
    })
    .sort(function (a, b) { return (b.updatedAt || 0) - (a.updatedAt || 0); });
  return {
    supplier: { name: s.name || '', sub: s.sub || '', phone: s.phone || '', note: s.note || '' },
    products: products, shops: shops, cons: cons
  };
}

function parse_(v, d) { try { return v ? JSON.parse(v) : d; } catch (e) { return d; } }

/* ---------- TULIS ---------- */
function applyOps_(ops) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    let n = 0;
    ops.forEach(function (op) {
      const d = op.data;
      if (op.t === 'supplier') {
        ['name', 'sub', 'phone', 'note'].forEach(function (k) { setSetting_(k, d[k] || ''); }); n++;
      } else if (op.t === 'pin') {
        if (String(d).length >= 4) { setSetting_('pin', String(d)); n++; }
      } else if (op.t === 'products') {
        writeProducts_(d || []); n++;
      } else if (op.t === 'shop') {
        if (upsert_('Kedai', d.id, [d.id, d.name || '', d.phone || '', String(d.updatedAt || Date.now())], d.updatedAt, 3)) n++;
      } else if (op.t === 'cons') {
        const row = [d.id, d.no || '', d.shopId || '', d.date || '', d.status || 'aktif',
          JSON.stringify(d.items || []), d.settle ? JSON.stringify(d.settle) : '',
          String(d.updatedAt || Date.now()), d.deleted ? 'true' : ''];
        if (upsert_('Resit', d.id, row, d.updatedAt, 7)) n++;
      }
    });
    return n;
  } finally {
    lock.releaseLock();
  }
}

function setSetting_(key, val) {
  const sh = ss_().getSheetByName('Tetapan');
  const last = sh.getLastRow();
  const keys = last >= 2 ? sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]); }) : [];
  const i = keys.indexOf(key);
  if (i >= 0) sh.getRange(i + 2, 2).setValue(String(val));
  else sh.appendRow([key, String(val)]);
}

function writeProducts_(list) {
  const sh = ss_().getSheetByName('Produk');
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 4).clearContent();
  if (!list.length) return;
  sh.getRange(2, 1, list.length, 4).setValues(list.map(function (p, i) {
    return [String(p.id), String(p.name || ''), String(+p.price || 0), String(i + 1)];
  }));
}

/* Kemaskini baris ikut id; hanya jika data baru lebih terkini (updatedAt) */
function upsert_(name, id, row, updatedAt, updCol) {
  const sh = ss_().getSheetByName(name);
  const last = sh.getLastRow();
  const width = SHEETS[name].length;
  if (last >= 2) {
    const data = sh.getRange(2, 1, last - 1, width).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        if ((+data[i][updCol] || 0) > (+updatedAt || 0)) return false; // data dalam sheet lebih baru
        sh.getRange(i + 2, 1, 1, width).setValues([row.map(String)]);
        return true;
      }
    }
  }
  sh.getRange(last + 1, 1, 1, width).setNumberFormat('@').setValues([row.map(String)]);
  return true;
}
