// Rekap Invoice Berdu — logika aplikasi (parsing, filter, pengelompokan, render, cetak).
// Konfigurasi pengirim ada di konstanta SENDER.

const SENDER = { nama: 'Daffi Hijab', telp: '0813-3244-2527', pickup: 'DIAMBIL DI GUDANG DROPSHIP' };
const MONTHS = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
// Pola nomor resi: Shopee (SPXID..), J&T (JY/JZ..), dan SP/angka panjang
const RESI_RE = /\b(SPXID[0-9A-Z]+|JY\d{8,}|JZ\d{8,}|SP\d{10,}|\d{16,})\b/g;

let FORMAT = 'ekspor';  // 'ekspor' | 'gudang'
let ALL = [];           // semua order hasil parse
let WARN = [];          // peringatan pemeriksaan data: { level: 'error'|'warn'|'info', msg }
const warn = (level, msg) => WARN.push({ level, msg });

const $ = id => document.getElementById(id);
const sum = a => a.reduce((x, y) => x + y, 0);
const rp = n => 'Rp ' + Number(n || 0).toLocaleString('id-ID');
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const toInt = x => parseInt(String(x || '').replace(/[^\d-]/g, ''), 10) || 0;
const normName = s => String(s || '').trim().replace(/\s+/g, ' ').toLowerCase();

// ---------- parsing file ----------
function parseCSV(text) {
  const rows = []; let row = [], field = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) {
      if (c === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else q = false; }
      else field += c;
    } else if (c === '"') q = true;
    else if (c === ',') { row.push(field); field = ''; }
    else if (c === '\n' || c === '\r') {
      if (c === '\r' && text[i + 1] === '\n') i++;
      row.push(field); rows.push(row); row = []; field = '';
    } else field += c;
  }
  if (field || row.length) { row.push(field); rows.push(row); }
  return rows;
}

function fileToRows(file) {
  return new Promise((resolve, reject) => {
    const isX = /\.xlsx?$/i.test(file.name);
    const r = new FileReader();
    r.onload = () => {
      try {
        if (isX) {
          const wb = XLSX.read(r.result, { type: 'array' });
          const ws = wb.Sheets[wb.SheetNames[0]];
          Object.keys(ws).forEach(a => {
            const c = ws[a];
            if (a[0] !== '!' && c && c.t === 'n' && Math.abs(c.v) >= 1e15 && !/E\+/i.test(c.w || '')) {
              warn('error', `Sel ${a} berisi angka ${c.w || c.v} (lebih dari 15 digit). Excel hanya menyimpan 15 digit, nomor resi/ID di sel ini bisa keliru. Ubah kolom ke Teks lalu simpan ulang.`);
            }
          });
          resolve(XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' }));
        } else resolve(parseCSV(r.result.replace(/^﻿/, '')));
      } catch (e) { reject(e); }
    };
    r.onerror = reject;
    isX ? r.readAsArrayBuffer(file) : r.readAsText(file);
  });
}

function parseDate(s) {
  const m = /(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(s || '');
  return m ? new Date(+m[3], +m[2] - 1, +m[1]) : null;
}
const fmtDate = d => d ? `${d.getDate()} ${MONTHS[d.getMonth()]} ${d.getFullYear()}` : '-';
const ymdCompact = d => d ? `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}` : 'NA';
// Ambil nilai warna saja (WARNA:BLACK -> BLACK). Atribut lain tetap ditampilkan.
const fmtWarna = v => {
  if (!v) return '—';
  const parts = String(v).split(',').map(p => p.trim()).filter(Boolean).map(p => {
    const i = p.indexOf(':');
    if (i < 0) return p.toUpperCase();
    const k = p.slice(0, i).trim().toUpperCase(), val = p.slice(i + 1).trim().toUpperCase();
    return (k === 'WARNA' || k === 'COLOR') ? val : `${k}: ${val}`;
  });
  return parts.join(' · ');
};
const fmtVar = v => {
  if (!v) return '';
  const i = v.indexOf(':');
  return i < 0 ? v.toUpperCase() : v.slice(0, i).toUpperCase() + ': ' + v.slice(i + 1).trim().toUpperCase();
};

// Catatan dari template gudang: pisahkan resi dari teks, buang link drive
function cleanNote(raw) {
  const resi = [...String(raw || '').matchAll(RESI_RE)].map(m => m[1]);
  const text = String(raw || '')
    .replace(/https?:\/\/\S+/g, '')
    .replace(RESI_RE, '')
    .split('\n').map(s => s.trim()).filter(Boolean).join('\n');
  return { note: text, resi: [...new Set(resi)] };
}

// Dispatcher: deteksi format dari header
function parseOrders(rows) {
  // Abaikan baris penanda pemisah Excel (mis. "sep=,")
  if (rows.length && /^sep=/i.test(String(rows[0][0] || ''))) rows = rows.slice(1);
  checkNotation(rows);
  const head = rows[0].map(h => String(h).trim());
  // Template gudang dicek lebih dulu: versi terbaru juga memuat kolom "Order Id"
  if (head.includes('Nomor') && head.includes('Nama Produk') && head.includes('Metode Pengiriman')) {
    FORMAT = 'gudang'; return parseGudang(rows, head);
  }
  if (head.includes('Order Id')) { FORMAT = 'ekspor'; return parseEkspor(rows, head); }
  throw new Error('Format tidak dikenali. Header harus berisi "Order Id" (ekspor berdu.id) atau "Nomor" + "Metode Pengiriman" (template gudang).');
}

// Format ekspor berdu.id: baris dengan Order Id = header order (total & diskon), tanpa Order Id = item lanjutan
function parseEkspor(rows, head) {
  const idx = n => head.indexOf(n);
  const I = {
    id: idx('Order Id'), buyer: idx('Nama Depan'), reseller: idx('Nama Reseller'),
    sent: idx('Waktu Dikirim'), status: idx('Status Pengiriman'), prod: idx('Nama Produk'),
    variasi: idx('Variasi Produk'), sku: idx('SKU Produk'), qty: idx('Jumlah Produk'),
    price: idx('Harga Produk'), disc: idx('Nilai Diskon'), total: idx('Total'), created: idx('Waktu Dibuat')
  };
  const findCol = re => head.findIndex(h => re.test(h));
  const OPT = {
    alamat: findCol(/alamat|address/i),
    telp: findCol(/telp|telepon|phone|no\.?\s*hp|whatsapp/i),
    catatan: findCol(/catatan|note/i)
  };
  const get = (r, k) => I[k] >= 0 ? String(r[I[k]] ?? '').trim() : '';
  const opt = (r, k) => OPT[k] >= 0 ? String(r[OPT[k]] ?? '').trim() : '';

  const orders = []; const seen = new Set(); let cur = null; let skipping = false;
  rows.slice(1).forEach((r, k) => {
    const line = k + 1;
    if (!r.some(c => String(c).trim())) return;
    if (get(r, 'id')) {
      const id = get(r, 'id');
      if (seen.has(id)) {
        warn('error', `Order ${id} muncul lebih dari sekali (baris ${line}). Duplikat DIABAIKAN, hanya kemunculan pertama yang dihitung.`);
        skipping = true; cur = null; return;
      }
      seen.add(id); skipping = false;
      if (get(r, 'sent') && !parseDate(get(r, 'sent'))) warn('warn', `Tanggal kirim tidak terbaca pada order ${id}: "${get(r, 'sent')}"`);
      cur = {
        id, buyer: get(r, 'buyer'), reseller: get(r, 'reseller'),
        sent: parseDate(get(r, 'sent')), status: get(r, 'status'), metode: '',
        created: get(r, 'created'), discount: toInt(get(r, 'disc')), total: toInt(get(r, 'total')), items: [],
        addr: opt(r, 'alamat'), tel: opt(r, 'telp'), note: opt(r, 'catatan'), resi: []
      };
      orders.push(cur);
    }
    if (get(r, 'prod')) {
      if (!cur) { if (!skipping) warn('error', `Baris produk tanpa order di baris ${line} (diabaikan): "${get(r, 'prod')}"`); return; }
      const rawQty = get(r, 'qty');
      if (!/^\d+$/.test(rawQty)) warn('error', `Jumlah "${rawQty || '(kosong)'}" tidak valid pada order ${cur.id}, produk ${get(r, 'prod')} (dihitung 1, periksa manual)`);
      const qty = toInt(rawQty) || 1, price = toInt(get(r, 'price'));
      cur.items.push({ prod: get(r, 'prod'), variasi: get(r, 'variasi'), sku: get(r, 'sku'),
                       qty, price, sub: qty * price, buyer: cur.buyer, orderId: cur.id, metode: '' });
    }
  });
  orders.forEach(o => {
    o.subtotal = sum(o.items.map(i => i.sub));
    o.mismatch = o.subtotal - o.discount !== o.total;
    if (o.mismatch) warn('warn', `Order ${o.id}: total ${rp(o.total)} tidak sama dengan subtotal − diskon (${rp(o.subtotal - o.discount)})`);
  });
  return orders;
}

// ID seperti RIT-11012024-27 yang menempel di akhir nama (bukan bagian nama customer)
const ID_IN_NAME = /\s+[A-Za-z]{2,}\d*-\d{6,}(?:-\d+)*\s*$/;
const cleanName = raw => String(raw || '').trim().replace(/\s+/g, ' ').replace(ID_IN_NAME, '');

// Template gudang: baris dengan Order Id / Nomor = order baru, baris kosong = item lanjutan. Tidak ada harga.
function parseGudang(rows, head) {
  const idx = n => head.indexOf(n);
  const C = {
    no: idx('Nomor'), oid: idx('Order Id'), created: idx('Waktu Dibuat'), buyer: idx('Nama Depan'), prod: idx('Nama Produk'),
    variasi: idx('Variasi Produk'), qty: idx('Jumlah Produk'), status: idx('Status Pengiriman'),
    metode: idx('Metode Pengiriman'), catatan: idx('Catatan dari pembeli')
  };
  const get = (r, k) => C[k] >= 0 ? String(r[C[k]] ?? '').trim() : '';

  const orders = []; const seen = new Set(); let cur = null; let skipping = false;
  rows.slice(1).forEach((r, k) => {
    const line = k + 1;
    if (!r.some(c => String(c).trim())) return;
    const oid = get(r, 'oid'), no = get(r, 'no');
    if (oid || no) {
      const id = oid || no;
      if (seen.has(id)) {
        warn('error', `Order ${id} muncul lebih dari sekali (baris ${line}). Duplikat DIABAIKAN, hanya kemunculan pertama yang dihitung.`);
        skipping = true; cur = null; return;
      }
      seen.add(id); skipping = false;
      const { note, resi } = cleanNote(get(r, 'catatan'));
      const metode = get(r, 'metode');
      if (!metode) warn('warn', `Metode pengiriman kosong pada order ${id}`);
      else if (!/PUSAT|DROPSHIP/i.test(metode)) warn('warn', `Metode pengiriman tidak dikenali "${metode}" pada order ${id}`);
      const rawBuyer = get(r, 'buyer');
      const buyer = cleanName(rawBuyer);
      if (!buyer) warn('warn', `Nama pembeli kosong pada order ${id}`);
      else if (buyer !== rawBuyer.replace(/\s+/g, ' ')) warn('info', `ID dilepas dari nama "${rawBuyer}" (order ${id}), customer dikenali sebagai "${buyer}".`);
      else if (/\d/.test(buyer)) warn('info', `Nama pembeli memuat angka: "${buyer}" (order ${id}). Pastikan penulisan konsisten.`);
      const created = get(r, 'created');
      const sent = parseDate(created);
      if (!sent) warn('warn', `Tanggal tidak terbaca pada order ${id}: "${created || '(kosong)'}"`);
      cur = {
        id, buyer, reseller: buyer,
        sent, status: get(r, 'status'), metode,
        created, discount: 0, total: 0, items: [],
        addr: '', tel: '', note, resi, mismatch: false
      };
      orders.push(cur);
    }
    if (get(r, 'prod')) {
      if (!cur) { if (!skipping) warn('error', `Baris produk tanpa order di baris ${line} (diabaikan): "${get(r, 'prod')}"`); return; }
      const rawQty = get(r, 'qty');
      if (!/^\d+$/.test(rawQty)) warn('error', `Jumlah "${rawQty || '(kosong)'}" tidak valid pada order ${cur.id}, produk ${get(r, 'prod')} (dihitung 1, periksa manual)`);
      cur.items.push({ prod: get(r, 'prod'), variasi: get(r, 'variasi'), sku: '',
                       qty: toInt(rawQty) || 1, price: 0, sub: 0,
                       buyer: cur.buyer, orderId: cur.id, metode: cur.metode });
    }
  });
  orders.forEach(o => o.subtotal = 0);
  return orders;
}

// Sel angka berformat ilmiah (1.23457E+18) berarti nomor sudah rusak di file sumber
function checkNotation(rows) {
  const head = rows[0] || [];
  rows.forEach((r, k) => {
    if (!k) return;
    r.forEach((c, j) => {
      const v = String(c).trim();
      if (/^-?\d+(\.\d+)?E[+-]\d+$/i.test(v)) {
        const col = head[j] ? String(head[j]).trim() : `kolom ${j + 1}`;
        warn('error', `Angka "${v}" di ${col} (baris ${k}) berformat ilmiah, nomor di sel ini rusak. Ubah kolom ke Teks lalu simpan ulang.`);
      }
    });
  });
}

// Jarak edit (Levenshtein) untuk mendeteksi typo
function lev(a, b) {
  if (a === b) return 0;
  if (!a.length) return b.length;
  if (!b.length) return a.length;
  let prev = Array.from({ length: b.length + 1 }, (_, j) => j);
  for (let i = 1; i <= a.length; i++) {
    const cur = [i];
    for (let j = 1; j <= b.length; j++) {
      cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
    prev = cur;
  }
  return prev[b.length];
}

// Pasangan nama yang mirip dan jarang muncul: kemungkinan typo dari nama yang lebih sering
function similarPairs(counts, minLen) {
  const names = [...counts.keys()], out = [];
  for (let i = 0; i < names.length; i++) for (let j = i + 1; j < names.length; j++) {
    const a = names[i], b = names[j];
    if (Math.min(a.length, b.length) < minLen || lev(a, b) > 1) continue;
    if (a.replace(/\d/g, '') === b.replace(/\d/g, '')) continue;   // hanya angka yang beda (HAPPY 1 / HAPPY 2)
    const [rare, common] = counts.get(a) <= counts.get(b) ? [a, b] : [b, a];
    if (counts.get(rare) <= 2 && counts.get(common) >= counts.get(rare)) out.push([rare, common]);
  }
  return out;
}

function checkTypos() {
  const prod = new Map(), warna = new Map(), cust = new Map();
  const bump = (m, k) => m.set(k, (m.get(k) || 0) + 1);
  for (const o of ALL) {
    const key = normName(o.reseller || o.buyer);
    if (key) bump(cust, key);
    for (const i of o.items) {
      bump(prod, i.prod.trim().replace(/\s+/g, ' ').toUpperCase());
      const w = fmtWarna(i.variasi);
      if (w !== '—') bump(warna, w);
    }
  }
  similarPairs(prod, 4).forEach(([r, c]) => warn('warn', `Kemungkinan typo produk: "${r}" (${prod.get(r)} baris) mirip "${c}" (${prod.get(c)} baris). Cek penulisan nama produk.`));
  similarPairs(warna, 4).forEach(([r, c]) => warn('warn', `Kemungkinan typo warna: "${r}" (${warna.get(r)} baris) mirip "${c}" (${warna.get(c)} baris).`));
  similarPairs(cust, 5).forEach(([r, c]) => warn('warn', `Kemungkinan dua customer sama: "${r}" mirip "${c}". Cek penulisan nama.`));
}

const LVL = {
  error: { cls: 'border-rose-200 bg-rose-50 text-rose-800', dot: 'bg-rose-500', label: 'error' },
  warn:  { cls: 'border-amber-200 bg-amber-50 text-amber-800', dot: 'bg-amber-500', label: 'peringatan' },
  info:  { cls: 'border-sky-200 bg-sky-50 text-sky-800', dot: 'bg-sky-500', label: 'info' },
};

// Panel pemeriksaan data di sidebar
function renderWarnings() {
  const box = $('sum-warn');
  if (!ALL.length) { box.hidden = true; box.innerHTML = ''; return; }
  box.hidden = false;
  if (!WARN.length) {
    box.innerHTML = `<div class="flex items-center gap-2 rounded-lg border border-emerald-200 bg-emerald-50 px-3 py-2 text-xs text-emerald-800">${icon('check')}<span>Semua data terbaca, tidak ada typo yang terdeteksi.</span></div>`;
    return;
  }
  const n = { error: 0, warn: 0, info: 0 };
  WARN.forEach(w => n[w.level]++);
  const counts = Object.keys(LVL).filter(l => n[l]).map(l => `${n[l]} ${LVL[l].label}`).join(' · ');
  box.innerHTML = `
    <div class="mb-2 flex items-center justify-between gap-2 text-xs">
      <span class="flex items-center gap-1.5 font-semibold text-gray-800">${icon('alert')} Pemeriksaan data</span>
      <span class="text-gray-500">${counts}</span>
    </div>
    <ul class="max-h-64 space-y-1.5 overflow-auto pr-1">
      ${WARN.map(w => `<li class="flex gap-2 rounded-md border px-2.5 py-1.5 text-[11px] leading-snug ${LVL[w.level].cls}"><span class="mt-1 h-1.5 w-1.5 shrink-0 rounded-full ${LVL[w.level].dot}"></span><span>${esc(w.msg)}</span></li>`).join('')}
    </ul>`;
}

// ---------- konfigurasi cetak & filter ----------
const PAPER_KEY = 'berdu-inv:paper';
const PAPERS = {
  thermal80: { label: 'Thermal 80 mm', w: '80mm', cls: 'paper-thermal80', thermal: true },
  thermal58: { label: 'Thermal 58 mm', w: '58mm', cls: 'paper-thermal58', thermal: true },
  a5:        { label: 'A5',            w: '148mm', cls: 'paper-a5', thermal: false },
};

// ---------- state ----------
let CONVERTED = false;
let FILE_NAME = '';
let METODE_SEL = '';
let PAPER = 'thermal80';
let INVOICES = [];   // satu invoice = satu produk milik satu customer (per metode)
let PERSONS = [];    // customer, berisi daftar invoice produk
let REKAP = [];      // rekap per produk, berisi warna
let SELECTED = null; // key invoice yang sedang dipreview
let TAB = 'invoice';
const EXPANDED = new Set(); // id baris yang sedang dibuka: "c:<person>", "p:<key>", "r:<produk>"
const wait = ms => new Promise(r => setTimeout(r, ms));

try { const v = localStorage.getItem(PAPER_KEY); if (PAPERS[v]) PAPER = v; } catch (e) { /* storage tidak tersedia */ }

// ---------- helper ----------
const metodeShort = m => /PUSAT/i.test(m) ? 'Gudang Pusat' : /DROPSHIP/i.test(m) ? 'Dropship' : '';
const pillCls = m => /PUSAT/i.test(m)
  ? 'bg-indigo-50 text-indigo-700'
  : 'bg-amber-50 text-amber-700';
const emptyHTML = t => `<div class="m-4 rounded-lg border border-dashed border-gray-200 px-4 py-10 text-center text-sm text-gray-500">${t}</div>`;

// Item digabung per produk; warna dijumlahkan
function productGroups(items) {
  const map = new Map();
  for (const i of items) {
    const k = i.prod.trim().toUpperCase();
    if (!map.has(k)) map.set(k, { prod: k, colors: new Map(), qty: 0, sub: 0, orderIds: [] });
    const g = map.get(k);
    const col = fmtWarna(i.variasi);
    if (!g.colors.has(col)) g.colors.set(col, { warna: col, qty: 0, sub: 0, prices: new Set() });
    const cc = g.colors.get(col);
    cc.qty += i.qty; cc.sub += i.sub; cc.prices.add(i.price);
    g.qty += i.qty; g.sub += i.sub; g.orderIds.push(i.orderId);
  }
  return [...map.values()].sort((a, b) => a.prod.localeCompare(b.prod)).map(g => ({
    prod: g.prod,
    qty: g.qty,
    sub: g.sub,
    orderIds: [...new Set(g.orderIds)],
    colors: [...g.colors.values()].sort((a, b) => a.warna.localeCompare(b.warna)).map(cc => ({
      warna: cc.warna, qty: cc.qty, sub: cc.sub,
      price: cc.prices.size === 1 ? [...cc.prices][0] : null
    }))
  }));
}

// ---------- filter order ----------
function filteredOrders() {
  const q = $('q').value.trim().toLowerCase();
  return ALL.filter(o => {
    if (METODE_SEL && o.metode !== METODE_SEL) return false;
    if (q) {
      const hay = [o.id, o.buyer, o.reseller, o.note, ...o.resi,
                   ...o.items.flatMap(i => [i.prod, i.variasi, i.sku])].join(' ').toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });
}

// ---------- pengelompokan ----------
// Indeks tetap dari SEMUA order (tidak ikut filter), supaya nomor invoice tidak berubah saat filter dipakai
function fullIndex() {
  const units = new Map();
  for (const o of ALL) {
    const raw = o.reseller || o.buyer || '(tanpa nama)';
    const person = normName(raw) || '(tanpa nama)';
    const metode = o.metode || '';
    const key = person + '|' + metode;
    if (!units.has(key)) units.set(key, { person, metode, name: raw.trim().replace(/\s+/g, ' '), orders: [], items: [] });
    const u = units.get(key);
    u.orders.push(o);
    u.items.push(...o.items);
  }
  const sorted = [...units.values()].sort((a, b) => a.name.localeCompare(b.name) || a.metode.localeCompare(b.metode));
  const idx = new Map();
  sorted.forEach((u, ui) => {
    const groups = productGroups(u.items);
    const parts = new Map(groups.map((g, gi) => [g.prod, { part: gi + 1, total: groups.length }]));
    const first = u.orders.map(o => o.sent).filter(Boolean).sort((a, b) => a - b)[0] || null;
    idx.set(u.person + '|' + u.metode, { no: ui + 1, first, parts });
  });
  return idx;
}

// Satu invoice per produk, per customer dan metode pengiriman
function buildInvoices(list) {
  const full = fullIndex();
  const units = new Map();
  for (const o of list) {
    const raw = o.reseller || o.buyer || '(tanpa nama)';
    const person = normName(raw) || '(tanpa nama)';
    const metode = o.metode || '';
    const key = person + '|' + metode;
    if (!units.has(key)) units.set(key, { person, metode, name: raw.trim().replace(/\s+/g, ' '), orders: [], items: [] });
    const u = units.get(key);
    u.orders.push(o);
    u.items.push(...o.items);
  }
  const sorted = [...units.values()].sort((a, b) => a.name.localeCompare(b.name) || a.metode.localeCompare(b.metode));
  const out = [];
  for (const u of sorted) {
    const f = full.get(u.person + '|' + u.metode);
    for (const g of productGroups(u.items)) {
      const pi = f.parts.get(g.prod);
      const orders = u.orders.filter(o => o.items.some(i => i.prod.trim().toUpperCase() === g.prod));
      out.push({
        key: u.person + '|' + u.metode + '|' + g.prod,
        person: u.person, name: u.name, metode: u.metode, metodeLabel: metodeShort(u.metode),
        pickup: u.metode || 'Metode belum diisi',
        first: f.first,
        invNo: `INV/${ymdCompact(f.first)}/${String(f.no).padStart(3, '0')}-P${pi.part}`,
        prod: g.prod, part: pi.part, partTotal: pi.total,
        colors: g.colors, pcs: g.qty, sub: g.sub,
        orderIds: g.orderIds,
        resi: [...new Set(orders.flatMap(o => o.resi))],
        notes: [...new Set(orders.map(o => o.note).filter(Boolean))],
        orders,
        mismatch: orders.some(o => o.mismatch),
        discount: sum(orders.map(o => o.discount)) > 0,
      });
    }
  }
  return out;
}

// Customer = gabungan invoice produk dengan nama yang sama
function groupPersons(invs) {
  const map = new Map();
  for (const inv of invs) {
    if (!map.has(inv.person)) map.set(inv.person, { person: inv.person, name: inv.name, invs: [], orders: new Map() });
    const p = map.get(inv.person);
    p.invs.push(inv);
    inv.orders.forEach(o => p.orders.set(o.id, o));
  }
  return [...map.values()].map((p, i) => {
    const orders = [...p.orders.values()];
    return {
      no: i + 1, person: p.person, name: p.name, invs: p.invs,
      metodes: [...new Set(p.invs.map(x => x.metode).filter(Boolean))],
      notes: [...new Set(p.invs.flatMap(x => x.notes))].map(n => n.replace(/\s*\n\s*/g, ' ')),
      orderCount: orders.length,
      pcs: sum(p.invs.map(x => x.pcs)),
      total: sum(orders.map(o => o.total)),
    };
  });
}

// Rekap per produk, dengan rincian warna
function buildRekap(list) {
  const map = new Map();
  for (const o of list) for (const i of o.items) {
    const pk = i.prod.trim().toUpperCase();
    if (!map.has(pk)) map.set(pk, { prod: pk, colors: new Map(), pusat: 0, ds: 0, total: 0, orders: new Set() });
    const p = map.get(pk);
    const warna = fmtWarna(i.variasi);
    if (!p.colors.has(warna)) p.colors.set(warna, { warna, pusat: 0, ds: 0, total: 0, orders: new Set() });
    const c = p.colors.get(warna);
    const isP = /PUSAT/i.test(i.metode), isD = /DROPSHIP/i.test(i.metode);
    for (const t of [p, c]) {
      t.total += i.qty;
      if (isP) t.pusat += i.qty; else if (isD) t.ds += i.qty;
      t.orders.add(o.id);
    }
    p.orders.add(o.id);
  }
  return [...map.values()].sort((a, b) => a.prod.localeCompare(b.prod)).map(p => ({
    prod: p.prod, pusat: p.pusat, ds: p.ds, total: p.total, orders: p.orders.size,
    colors: [...p.colors.values()].sort((a, b) => a.warna.localeCompare(b.warna))
      .map(c => ({ warna: c.warna, pusat: c.pusat, ds: c.ds, total: c.total, orders: c.orders.size })),
  }));
}

// ---------- ikon & tombol ----------
const ICON = {
  eye: '<path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/>',
  printer: '<path d="M6 9V3h12v6"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="7" rx="1"/>',
  chevR: '<path d="M9 6l6 6-6 6"/>',
  chevL: '<path d="M15 6l-6 6 6 6"/>',
  chevsL: '<path d="M11 6l-6 6 6 6"/><path d="M18 6l-6 6 6 6"/>',
  chevsR: '<path d="M13 6l6 6-6 6"/><path d="M6 6l6 6-6 6"/>',
  check: '<path d="M20 6 9 17l-5-5"/>',
  alert: '<path d="M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/><path d="M12 9v4M12 17h.01"/>',
  note: '<path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>',
};
const icon = (k, cls = 'h-3.5 w-3.5') => `<svg class="${cls}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICON[k]}</svg>`;
const btn = (kind, ic, label, attrs = '') => `<button type="button" class="btn btn-${kind}" ${attrs}>${icon(ic)}<span>${label}</span></button>`;
const toggleBtn = (id, open) => `<button type="button" class="toggle" data-toggle="${esc(id)}" aria-expanded="${open}" aria-label="${open ? 'Tutup detail' : 'Lihat detail'}">${icon('chevR', 'chev')}</button>`;

// ---------- tabel: state paginasi & animasi ----------
const PAGE = { invoice: 1, rekap: 1 };
const SIZE = { invoice: 10, rekap: 10 };
const JUST = new Set();   // id yang baru dibuka; anaknya dianimasikan sekali
let LAST_SIG = '';
let BUSY = false;

const animCls = parent => JUST.has(parent) ? ' anim' : '';
const fx = (parent, i) => JUST.has(parent) ? ` style="--d:${Math.min(i, 6) * 30}ms"` : '';

// ---------- tabel: render ----------
function invoiceTableHTML(rows, price) {
  const body = rows.map(p => {
    const id = 'c:' + p.person;
    const open = EXPANDED.has(id);
    let h = `<tr class="grp${open ? ' open' : ''}" data-row="${esc(id)}">
      <td class="w-12">${toggleBtn(id, open)}</td>
      <td><div class="font-semibold">${esc(p.name)}</div><div class="text-xs text-gray-500">${p.orderCount} order</div>${p.notes.length ? `<div class="mt-1 flex max-w-xs items-start gap-1 text-[11px] text-amber-800" title="${esc(p.notes.join(' | '))}">${icon('note', 'mt-px h-3 w-3 shrink-0')}<span class="line-clamp-2">${esc(p.notes.join(' · '))}</span></div>` : ''}</td>
      <td><div class="flex flex-wrap gap-1">${p.metodes.map(m => `<span class="pill ${pillCls(m)}">${esc(metodeShort(m))}</span>`).join('')}</div></td>
      <td class="num">${p.invs.length}</td>
      <td class="num">${p.pcs}</td>
      ${price ? `<td class="num font-semibold">${rp(p.total)}</td>` : ''}
      <td class="act">${btn('ghost', 'printer', 'Cetak semua', `data-print-person="${esc(p.person)}"`)}</td>
    </tr>`;
    if (open) p.invs.forEach((inv, i) => {
      const pid = 'p:' + inv.key;
      const po = EXPANDED.has(pid);
      h += `<tr class="lvl1${animCls(id)}" data-parent="${esc(id)}" data-row="${esc(pid)}"${fx(id, i)}>
        <td>${toggleBtn(pid, po)}</td>
        <td><div class="font-medium text-gray-900">${esc(inv.prod)}</div><div class="text-[11px] text-gray-500">Part ${inv.part} dari ${inv.partTotal} · ${esc(inv.invNo)}</div></td>
        <td><span class="pill ${pillCls(inv.metode)}">${esc(inv.metodeLabel || '—')}</span></td>
        <td></td>
        <td class="num">${inv.pcs}</td>
        ${price ? `<td class="num">${rp(inv.sub)}</td>` : ''}
        <td class="act">${btn('ghost', 'eye', 'Preview', `data-preview="${esc(inv.key)}"`)}${btn('dark', 'printer', 'Cetak', `data-print-one="${esc(inv.key)}"`)}</td>
      </tr>`;
      if (po) inv.colors.forEach((c, j) => {
        h += `<tr class="lvl2${animCls(pid)}" data-parent="${esc(pid)}"${fx(pid, j)}>
          <td></td>
          <td>${esc(c.warna)}${price && c.price !== null ? ` <span class="text-gray-400">· ${rp(c.price)}</span>` : ''}</td>
          <td></td><td></td>
          <td class="num">${c.qty}x</td>
          ${price ? `<td class="num">${rp(c.sub)}</td>` : ''}
          <td></td>
        </tr>`;
      });
    });
    return h;
  }).join('');
  const head = `<tr>
    <th class="w-12"></th><th>Customer</th><th>Metode</th>
    <th class="num">Invoice</th><th class="num">Pcs</th>
    ${price ? '<th class="num">Total</th>' : ''}
    <th class="act-h">Aksi</th></tr>`;
  return `<table class="dt"><thead>${head}</thead><tbody>${body}</tbody></table>`;
}

function rekapTableHTML(rows, split) {
  const head = `<tr><th class="w-12"></th><th>Produk / Warna</th>
    ${split ? '<th class="num">Gudang Pusat</th><th class="num">Dropship</th>' : ''}
    <th class="num">Total qty</th><th class="num">Order</th></tr>`;
  const body = rows.map(p => {
    const id = 'r:' + p.prod;
    const open = EXPANDED.has(id);
    let h = `<tr class="grp${open ? ' open' : ''}" data-row="${esc(id)}">
      <td class="w-12">${toggleBtn(id, open)}</td>
      <td><div class="font-semibold">${esc(p.prod)}</div><div class="text-xs text-gray-500">${p.colors.length} warna</div></td>
      ${split ? `<td class="num">${p.pusat}</td><td class="num">${p.ds}</td>` : ''}
      <td class="num font-semibold">${p.total}</td>
      <td class="num">${p.orders}</td>
    </tr>`;
    if (open) p.colors.forEach((c, j) => {
      h += `<tr class="lvl2${animCls(id)}" data-parent="${esc(id)}"${fx(id, j)}>
        <td></td>
        <td>${esc(c.warna)}</td>
        ${split ? `<td class="num">${c.pusat}</td><td class="num">${c.ds}</td>` : ''}
        <td class="num">${c.total}</td>
        <td class="num">${c.orders}</td>
      </tr>`;
    });
    return h;
  }).join('');
  return `<table class="dt"><thead>${head}</thead><tbody>${body}</tbody></table>`;
}

// nomor halaman yang ditampilkan, '…' untuk celah
function pageNums(cur, pages) {
  if (pages <= 7) return Array.from({ length: pages }, (_, i) => i + 1);
  const arr = [...new Set([1, pages, cur - 1, cur, cur + 1].filter(n => n >= 1 && n <= pages))].sort((a, b) => a - b);
  const out = [];
  arr.forEach((n, i) => { if (i && n - arr[i - 1] > 1) out.push('…'); out.push(n); });
  return out;
}

function pageSlice(tab, items) {
  const size = SIZE[tab];
  const pages = Math.max(1, Math.ceil(items.length / size));
  PAGE[tab] = Math.min(Math.max(1, PAGE[tab]), pages);
  const start = (PAGE[tab] - 1) * size;
  return { rows: items.slice(start, start + size), cur: PAGE[tab], pages, start, total: items.length };
}

function pagerHTML(tab, info) {
  if (!info.total) return '';
  const { cur, pages, start, total } = info;
  const end = Math.min(start + SIZE[tab], total);
  const ctrl = (n, disabled, ic, label) => `<button type="button" class="pager-ctrl w-8 justify-center px-0" data-page="${tab}:${n}" ${disabled ? 'disabled' : ''} aria-label="${label}" title="${label}">${icon(ic)}</button>`;
  const nums = pageNums(cur, pages).map(n => n === '…'
    ? '<span class="pager-gap" aria-hidden="true">…</span>'
    : `<button type="button" class="pg-num${n === cur ? ' pg-on' : ''}" data-page="${tab}:${n}"${n === cur ? ' aria-current="page"' : ''}>${n}</button>`).join('');
  return `<nav class="pager" aria-label="Navigasi halaman">
    <div class="pager-info">Menampilkan <b>${start + 1}–${end}</b> dari <b>${total}</b> ${tab === 'invoice' ? 'customer' : 'produk'}</div>
    <div class="pager-mid">
      ${ctrl(cur - 1, cur <= 1, 'chevL', 'Halaman sebelumnya')}
      <div class="pager-nums">${nums}</div>
      ${ctrl(cur + 1, cur >= pages, 'chevR', 'Halaman berikutnya')}
    </div>
  </nav>`;
}

function sizeHTML(tab) {
  return `<label class="pager-size">Baris per halaman
    <select data-pagesize="${tab}">${[10, 25, 50].map(n => `<option value="${n}"${n === SIZE[tab] ? ' selected' : ''}>${n}</option>`).join('')}</select>
  </label>`;
}

function renderTables() {
  const price = FORMAT === 'ekspor';
  const split = FORMAT === 'gudang';

  const inv = pageSlice('invoice', PERSONS);
  $('list').innerHTML = inv.total ? invoiceTableHTML(inv.rows, price) : emptyHTML('Tidak ada data sesuai filter.');
  $('sz-invoice').innerHTML = sizeHTML('invoice');
  $('pg-invoice').innerHTML = pagerHTML('invoice', inv);

  const rk = pageSlice('rekap', REKAP);
  $('rekap').innerHTML = rk.total ? rekapTableHTML(rk.rows, split) : emptyHTML('Tidak ada data sesuai filter.');
  $('sz-rekap').innerHTML = sizeHTML('rekap');
  $('pg-rekap').innerHTML = pagerHTML('rekap', rk);

  JUST.clear();
}

// ---------- tabel: toggle dengan animasi tutup ----------
function markClosing(id) {
  const opener = [...document.querySelectorAll('tr[data-row]')].find(tr => tr.dataset.row === id);
  if (!opener) return;
  const gone = new Set([id]);
  for (const tr of opener.parentElement.rows) {
    if (gone.has(tr.dataset.parent)) {
      tr.classList.add('closing');
      if (tr.dataset.row) gone.add(tr.dataset.row);
    }
  }
}

function onToggle(id) {
  if (BUSY) return;
  if (EXPANDED.has(id)) {
    BUSY = true;
    markClosing(id);
    setTimeout(() => { EXPANDED.delete(id); BUSY = false; renderTables(); }, 200);
  } else {
    EXPANDED.add(id);
    JUST.add(id);
    renderTables();
  }
}

function onTableClick(e) {
  const t = e.target.closest('[data-toggle]');
  if (t) return onToggle(t.dataset.toggle);
  const pv = e.target.closest('[data-preview]');
  if (pv) return showPreview(pv.dataset.preview);
  const one = e.target.closest('[data-print-one]');
  if (one) return doPrint(INVOICES.filter(i => i.key === one.dataset.printOne));
  const per = e.target.closest('[data-print-person]');
  if (per) return doPrint(INVOICES.filter(i => i.person === per.dataset.printPerson));
}

function onPager(e) {
  const b = e.target.closest('[data-page]');
  if (!b || b.disabled) return;
  const [tab, n] = b.dataset.page.split(':');
  PAGE[tab] = Number(n);
  renderTables();
  $(tab === 'invoice' ? 'list' : 'rekap').scrollTop = 0;
}

function onPageSize(e) {
  const s = e.target.closest('[data-pagesize]');
  if (!s) return;
  const tab = s.dataset.pagesize;
  SIZE[tab] = Number(s.value);
  PAGE[tab] = 1;
  renderTables();
}

// ---------- invoice ----------
function invoiceHTML(inv) {
  const cls = PAPERS[PAPER].cls;
  const narrow = PAPER === 'thermal58';
  const ekspor = FORMAT === 'ekspor';
  const showHarga = ekspor && !narrow;
  const head = `<tr><th>Warna</th><th class="num">Qty</th>${showHarga ? '<th class="num">Harga</th>' : ''}${ekspor ? '<th class="num">Subtotal</th>' : ''}</tr>`;
  const rows = inv.colors.map(c => `<tr>
      <td>${esc(c.warna)}</td>
      <td class="num">${c.qty}x</td>
      ${showHarga ? `<td class="num">${c.price === null ? '—' : rp(c.price)}</td>` : ''}
      ${ekspor ? `<td class="num">${rp(c.sub)}</td>` : ''}
    </tr>`).join('');
  const created = inv.orders[0] ? inv.orders[0].created : '';
  return `
  <article class="inv-sheet ${cls}">
    <header class="inv-center">
      <div class="inv-store">${esc(SENDER.nama)}</div>
      <div>${esc(SENDER.telp)}</div>
    </header>
    <hr class="sep">
    <div class="inv-title-row"><span class="inv-title">INVOICE</span><span class="inv-no">${esc(inv.invNo)}</span></div>
    <dl class="inv-kv">
      <dt>Tanggal</dt><dd>${fmtDate(inv.first)}</dd>
      ${created ? `<dt>Waktu order</dt><dd>${esc(created)}</dd>` : ''}
      <dt>Kepada</dt><dd><b>${esc(inv.name)}</b></dd>
      <dt>Pengiriman</dt><dd>${esc(inv.pickup)}</dd>
      <dt>Order ID</dt><dd>${inv.orderIds.map(esc).join(', ')}</dd>
      ${inv.resi.length ? `<dt>Resi</dt><dd>${inv.resi.map(esc).join('<br>')}</dd>` : ''}
    </dl>
    ${inv.notes.length ? `<div class="inv-note"><b>Catatan:</b> ${inv.notes.map(esc).join(' | ')}</div>` : ''}
    <hr class="sep">
    <div class="inv-part">Produk ${inv.part} dari ${inv.partTotal}</div>
    <div class="inv-prod">${esc(inv.prod)}</div>
    <table class="inv-tbl">
      <thead>${head}</thead>
      <tbody>${rows}</tbody>
      <tfoot><tr>
        <td>Total</td>
        <td class="num">${inv.pcs}x</td>
        ${showHarga ? '<td></td>' : ''}
        ${ekspor ? `<td class="num">${rp(inv.sub)}</td>` : ''}
      </tr></tfoot>
    </table>
    ${ekspor ? `<div class="inv-sum"><span>Subtotal produk</span><b>${rp(inv.sub)}</b></div>` : ''}
    ${ekspor && inv.discount ? '<div class="inv-small">Order ini memakai diskon/voucher. Potongannya tidak dialokasikan per produk.</div>' : ''}
    ${inv.mismatch ? '<div class="inv-small">Total di file berbeda dari subtotal − diskon.</div>' : ''}
    <hr class="sep">
    <div class="inv-foot">Bagian ${inv.part} dari ${inv.partTotal} untuk ${esc(inv.name)}</div>
  </article>`;
}

function rekapSheetHTML() {
  const cls = PAPERS[PAPER].cls;
  const split = FORMAT === 'gudang';
  const body = REKAP.map(p => `
    <tr class="grp-print">
      <td><b>${esc(p.prod)}</b></td>
      ${split ? `<td class="num">${p.pusat}</td><td class="num">${p.ds}</td>` : ''}
      <td class="num"><b>${p.total}</b></td>
    </tr>
    ${p.colors.map(c => `<tr><td>&nbsp;&nbsp;${esc(c.warna)}</td>${split ? `<td class="num">${c.pusat}</td><td class="num">${c.ds}</td>` : ''}<td class="num">${c.total}</td></tr>`).join('')}`).join('');
  const pcs = sum(REKAP.map(p => p.total));
  return `
  <article class="inv-sheet ${cls}">
    <header class="inv-center">
      <div class="inv-store">REKAP PRODUK</div>
      <div>${fmtDate(new Date())}</div>
    </header>
    <hr class="sep">
    <table class="inv-tbl">
      <thead><tr><th>Produk / Warna</th>${split ? '<th class="num">Pusat</th><th class="num">DS</th>' : ''}<th class="num">Qty</th></tr></thead>
      <tbody>${body}</tbody>
      <tfoot><tr><td>Total pcs</td>${split ? '<td></td><td></td>' : ''}<td class="num">${pcs}</td></tr></tfoot>
    </table>
  </article>`;
}

// ---------- render ----------
function renderAll() {
  const list = filteredOrders();
  INVOICES = buildInvoices(list);
  PERSONS = groupPersons(INVOICES);
  REKAP = buildRekap(list);

  const pcs = sum(list.flatMap(o => o.items).map(i => i.qty));
  $('st-orders').textContent = list.length;
  $('st-cust').textContent = PERSONS.length;
  $('st-invoices').textContent = INVOICES.length;
  if (FORMAT === 'gudang') {
    $('st-total-label').textContent = 'Total pcs';
    $('st-total').textContent = pcs;
  } else {
    $('st-total-label').textContent = 'Total';
    $('st-total').textContent = rp(sum(list.map(o => o.total)));
  }
  $('cust-count').textContent = PERSONS.length ? `${PERSONS.length} customer · ${INVOICES.length} invoice` : '';
  const sig = [METODE_SEL, $('q').value, FORMAT, FILE_NAME].join('|');
  if (sig !== LAST_SIG) { PAGE.invoice = 1; PAGE.rekap = 1; LAST_SIG = sig; }
  renderTables();
  $('btn-print-all').disabled = !INVOICES.length;
  $('btn-print-rekap').disabled = !REKAP.length;

  if (SELECTED && !INVOICES.some(i => i.key === SELECTED)) closeModal();
  else if (SELECTED) refreshPreview();
}

// ---------- cetak ----------
// Thermal: gulungan kertas tidak punya tinggi tetap, jadi setiap struk diukur
// dan diberi halaman bernama dengan lebar kertas + tinggi struk itu sendiri.
// A5: halaman tetap, satu customer per halaman.
function preparePrint(sheetsHTML) {
  const area = $('print-area');
  const p = PAPERS[PAPER];
  area.classList.toggle('thermal', p.thermal);
  area.classList.toggle('a5', !p.thermal);
  area.innerHTML = sheetsHTML;

  let css;
  if (p.thermal) {
    area.style.cssText = 'display:block;position:absolute;left:-10000px;top:0;';
    const pxToMm = 25.4 / 96;
    css = '';
    area.querySelectorAll('.inv-sheet').forEach((el, i) => {
      const h = Math.ceil(el.getBoundingClientRect().height * pxToMm) + 3;
      css += `@page pg${i} { size: ${p.w} ${h}mm; margin: 0; }\n`;
      el.style.setProperty('page', `pg${i}`);
    });
    area.style.cssText = '';
  } else {
    css = '@page { size: A5 portrait; margin: 6mm; }';
  }
  let st = document.getElementById('page-css');
  if (!st) { st = document.createElement('style'); st.id = 'page-css'; document.head.appendChild(st); }
  st.textContent = css;
}

function doPrint(list) {
  if (!list.length) return;
  preparePrint(list.map(invoiceHTML).join(''));
  window.print();
}

function printRekap() {
  if (!REKAP.length) return;
  preparePrint(rekapSheetHTML());
  window.print();
}

// ---------- modal preview ----------
function siblingsOf(inv) {
  return INVOICES.filter(x => x.person === inv.person && x.metode === inv.metode);
}

function refreshPreview() {
  const inv = INVOICES.find(i => i.key === SELECTED);
  if (!inv) return;
  const sib = siblingsOf(inv);
  const idx = sib.findIndex(x => x.key === inv.key);
  $('pv-title').textContent = `${inv.name} · ${inv.prod}`;
  $('pv-pos').textContent = `${inv.metodeLabel ? inv.metodeLabel + ' · ' : ''}Produk ${inv.part} dari ${inv.partTotal}`;
  $('btn-prev').disabled = idx <= 0;
  $('btn-next').disabled = idx >= sib.length - 1;
  $('pv').innerHTML = invoiceHTML(inv);
  $('warn').innerHTML = '';
}

function showPreview(key) {
  SELECTED = key;
  refreshPreview();
  $('modal').hidden = false;
  document.body.classList.add('modal-open');
}

function closeModal() {
  $('modal').hidden = true;
  document.body.classList.remove('modal-open');
  SELECTED = null;
}

function stepPreview(d) {
  const inv = INVOICES.find(i => i.key === SELECTED);
  if (!inv) return;
  const sib = siblingsOf(inv);
  const target = sib[sib.findIndex(x => x.key === inv.key) + d];
  if (target) { SELECTED = target.key; refreshPreview(); }
}

// ---------- tab ----------
function setTab(t) {
  TAB = t;
  $('pane-invoice').hidden = t !== 'invoice';
  $('pane-rekap').hidden = t !== 'rekap';
  document.querySelectorAll('[data-tab]').forEach(b => {
    const on = b.dataset.tab === t;
    b.classList.toggle('seg-on', on);
    b.setAttribute('aria-selected', String(on));
  });
}

// ---------- tahap aplikasi: upload -> konversi -> hasil ----------
const CONVERT_STEPS = [
  'Membaca struktur file',
  'Mengelompokkan order per customer',
  'Menyusun rekap produk',
  'Menyiapkan invoice per produk',
];

function setViews() {
  const has = ALL.length > 0;
  $('view-upload').hidden = CONVERTED;
  $('view-result').hidden = !CONVERTED;
  $('picked').hidden = !has || CONVERTED;
  $('picked-name').textContent = FILE_NAME;
  $('filters').disabled = !CONVERTED;
  $('filter-hint').hidden = CONVERTED;
  $('btn-convert').disabled = !has || CONVERTED;
  $('btn-convert').querySelector('span').textContent = CONVERTED ? 'Sudah dikonversi' : 'Convert to Invoice';
  $('btn-reset').hidden = !has;
}

function updateSummary() {
  const has = ALL.length > 0;
  $('fname').textContent = has ? FILE_NAME : 'Belum ada file.';
  $('fmt').textContent = !has ? 'Belum ada data'
    : FORMAT === 'gudang' ? 'Template gudang (tanpa harga)' : 'Ekspor berdu.id (dengan harga)';
  if (!CONVERTED) {
    ['st-orders', 'st-cust', 'st-invoices', 'st-total'].forEach(id => $(id).textContent = '–');
    $('st-total-label').textContent = 'Total';
    renderWarnings();
    return;
  }
  const pcs = sum(ALL.flatMap(o => o.items).map(i => i.qty));
  const persons = new Set(ALL.map(o => normName(o.reseller || o.buyer))).size;
  $('st-orders').textContent = ALL.length;
  $('st-cust').textContent = persons;
  $('st-invoices').textContent = has ? buildInvoices(ALL).length : 0;
  if (FORMAT === 'gudang') {
    $('st-total-label').textContent = 'Total pcs';
    $('st-total').textContent = pcs;
  } else {
    $('st-total-label').textContent = 'Total';
    $('st-total').textContent = rp(sum(ALL.map(o => o.total)));
  }
  renderWarnings();
}

function paintMetode() {
  const hasMetode = ALL.some(o => o.metode);
  document.querySelectorAll('[data-metode]').forEach(b => {
    const on = b.dataset.metode === METODE_SEL;
    b.classList.toggle('seg-on', on);
    b.disabled = !hasMetode && b.dataset.metode !== '';
  });
}

function fillFilterOptions() {
  METODE_SEL = '';
  PAGE.invoice = PAGE.rekap = 1;
  LAST_SIG = '';
  $('q').value = '';
  paintMetode();
}

async function loadFile(f) {
  if (!f) return;
  WARN = [];
  try {
    const rows = await fileToRows(f);
    ALL = parseOrders(rows);
    checkTypos();
    FILE_NAME = f.name;
    CONVERTED = false;
    SELECTED = null;
    EXPANDED.clear();
    fillFilterOptions();
    updateSummary();
    setViews();
  } catch (err) {
    alert('Gagal membaca file: ' + err.message);
  }
}

function setBar(p) {
  $('loader-bar').style.width = p + '%';
  $('loader-pct').textContent = p + '%';
}

function setStep(i, state) {
  const li = $('loader-steps').children[i];
  if (!li) return;
  const cls = { pending: 'text-gray-400', active: 'font-medium text-accent-600', done: 'text-emerald-600' };
  const icon = { pending: '○', active: '●', done: '✓' };
  li.className = 'flex items-center gap-2 ' + cls[state];
  li.firstElementChild.textContent = icon[state];
}

// Setiap langkah menjalankan pekerjaan nyata, lalu memperbarui progress
async function runConvert() {
  if (!ALL.length || CONVERTED) return;
  $('loader-steps').innerHTML = CONVERT_STEPS.map(s =>
    `<li class="flex items-center gap-2 text-gray-400"><span class="w-4 text-center">○</span><span>${esc(s)}</span></li>`).join('');
  $('loader-title').textContent = 'Mengonversi data';
  $('loader').hidden = false;
  setBar(0);

  const n = CONVERT_STEPS.length;
  let invoices = 0;
  for (let i = 0; i < n; i++) {
    setStep(i, 'active');
    $('loader-step').textContent = CONVERT_STEPS[i] + '…';
    await wait(400);
    if (i === 1) invoices = buildInvoices(ALL).length;
    if (i === 2) buildRekap(ALL);
    setStep(i, 'done');
    setBar(Math.round(((i + 1) / n) * 100));
    await wait(180);
  }

  $('loader-title').textContent = 'Selesai';
  $('loader-step').textContent = `${invoices} invoice siap`;
  await wait(500);
  $('loader').hidden = true;
  CONVERTED = true;
  setTab('invoice');
  setViews();
  renderAll();
}

function resetAll() {
  ALL = [];
  WARN = [];
  FILE_NAME = '';
  CONVERTED = false;
  SELECTED = null;
  EXPANDED.clear();
  closeModal();
  $('file').value = '';
  fillFilterOptions();
  updateSummary();
  setViews();
}

// ---------- events ----------
$('btn-choose').addEventListener('click', () => $('file').click());
$('file').addEventListener('change', e => loadFile(e.target.files[0]));

const dz = $('dropzone');
const dzOn = () => dz.classList.add('border-accent-600', 'bg-accent-50');
const dzOff = () => dz.classList.remove('border-accent-600', 'bg-accent-50');
['dragenter', 'dragover'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dzOn(); }));
['dragleave', 'drop'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dzOff(); }));
dz.addEventListener('drop', e => loadFile(e.dataTransfer.files[0]));

$('btn-convert').addEventListener('click', runConvert);
$('btn-reset').addEventListener('click', resetAll);

$('q').addEventListener('input', () => { if (CONVERTED) renderAll(); });
document.querySelectorAll('[data-metode]').forEach(b => b.addEventListener('click', () => {
  METODE_SEL = b.dataset.metode;
  paintMetode();
  if (CONVERTED) renderAll();
}));

document.querySelectorAll('[data-tab]').forEach(b => b.addEventListener('click', () => setTab(b.dataset.tab)));

// Pengaturan kertas
$('paper').value = PAPER;
$('paper').addEventListener('change', e => {
  PAPER = e.target.value;
  try { localStorage.setItem(PAPER_KEY, PAPER); } catch (err) { /* abaikan */ }
  if (SELECTED) refreshPreview();
});

// Tabel: toggle, aksi, dan paginasi
$('list').addEventListener('click', onTableClick);
$('rekap').addEventListener('click', onTableClick);
$('pg-invoice').addEventListener('click', onPager);
$('pg-rekap').addEventListener('click', onPager);
$('sz-invoice').addEventListener('change', onPageSize);
$('sz-rekap').addEventListener('change', onPageSize);

// Modal
document.querySelectorAll('[data-close]').forEach(el => el.addEventListener('click', closeModal));
$('btn-prev').addEventListener('click', () => stepPreview(-1));
$('btn-next').addEventListener('click', () => stepPreview(1));
$('btn-print-one').addEventListener('click', () => doPrint(INVOICES.filter(i => i.key === SELECTED)));
document.addEventListener('keydown', e => {
  if ($('modal').hidden) return;
  if (e.key === 'Escape') closeModal();
  if (e.key === 'ArrowLeft') stepPreview(-1);
  if (e.key === 'ArrowRight') stepPreview(1);
});

$('btn-print-all').addEventListener('click', () => doPrint(INVOICES));
$('btn-print-rekap').addEventListener('click', printRekap);

setTab('invoice');
setViews();
updateSummary();
