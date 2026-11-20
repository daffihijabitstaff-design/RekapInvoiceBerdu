// Rekap Invoice Berdu — logika aplikasi (parsing, filter, pengelompokan, render, cetak).
// Konfigurasi pengirim ada di konstanta SENDER.

const SENDER = { nama: 'Daffi Hijab', telp: '0813-3244-2527', pickup: 'DIAMBIL DI GUDANG DROPSHIP' };
const MONTHS = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
// Pola nomor resi: Shopee (SPXID..), J&T (JY/JZ..), dan SP/angka panjang
const RESI_RE = /\b(SPXID[0-9A-Z]+|JY\d{8,}|JZ\d{8,}|SP\d{10,}|\d{16,})\b/g;

let FORMAT = 'ekspor';  // 'ekspor' | 'gudang'
let ALL = [];           // semua order hasil parse

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
  const head = rows[0].map(h => String(h).trim());
  if (head.includes('Order Id')) { FORMAT = 'ekspor'; return parseEkspor(rows, head); }
  if (head.includes('Nomor') && head.includes('Nama Produk') && head.includes('Metode Pengiriman')) {
    FORMAT = 'gudang'; return parseGudang(rows, head);
  }
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

  const orders = []; let cur = null;
  for (const r of rows.slice(1)) {
    if (!r.some(c => String(c).trim())) continue;
    if (get(r, 'id')) {
      cur = {
        id: get(r, 'id'), buyer: get(r, 'buyer'), reseller: get(r, 'reseller'),
        sent: parseDate(get(r, 'sent')), status: get(r, 'status'), metode: '',
        created: get(r, 'created'), discount: toInt(get(r, 'disc')), total: toInt(get(r, 'total')), items: [],
        addr: opt(r, 'alamat'), tel: opt(r, 'telp'), note: opt(r, 'catatan'), resi: []
      };
      orders.push(cur);
    }
    if (cur && get(r, 'prod')) {
      const qty = toInt(get(r, 'qty')) || 1, price = toInt(get(r, 'price'));
      cur.items.push({ prod: get(r, 'prod'), variasi: get(r, 'variasi'), sku: get(r, 'sku'),
                       qty, price, sub: qty * price, buyer: cur.buyer, orderId: cur.id, metode: '' });
    }
  }
  orders.forEach(o => {
    o.subtotal = sum(o.items.map(i => i.sub));
    o.mismatch = o.subtotal - o.discount !== o.total;
  });
  return orders;
}

// Template gudang: baris dengan Nomor = order baru, baris kosong Nomor = item lanjutan. Tidak ada harga.
function parseGudang(rows, head) {
  const idx = n => head.indexOf(n);
  const C = {
    no: idx('Nomor'), created: idx('Waktu Dibuat'), buyer: idx('Nama Depan'), prod: idx('Nama Produk'),
    variasi: idx('Variasi Produk'), qty: idx('Jumlah Produk'), status: idx('Status Pengiriman'),
    metode: idx('Metode Pengiriman'), catatan: idx('Catatan dari pembeli')
  };
  const get = (r, k) => C[k] >= 0 ? String(r[C[k]] ?? '').trim() : '';

  const orders = []; let cur = null;
  for (const r of rows.slice(1)) {
    if (!r.some(c => String(c).trim())) continue;
    if (get(r, 'no')) {
      const { note, resi } = cleanNote(get(r, 'catatan'));
      cur = {
        id: get(r, 'no'), buyer: get(r, 'buyer'), reseller: get(r, 'buyer'),
        sent: parseDate(get(r, 'created')), status: get(r, 'status'), metode: get(r, 'metode'),
        created: get(r, 'created'), discount: 0, total: 0, items: [],
        addr: '', tel: '', note, resi, mismatch: false
      };
      orders.push(cur);
    }
    if (cur && get(r, 'prod')) {
      cur.items.push({ prod: get(r, 'prod'), variasi: get(r, 'variasi'), sku: '',
                       qty: toInt(get(r, 'qty')) || 1, price: 0, sub: 0,
                       buyer: cur.buyer, orderId: cur.id, metode: cur.metode });
    }
  }
  orders.forEach(o => o.subtotal = 0);
  return orders;
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
const toggleBtn = (id, open) => `<button type="button" class="toggle" data-toggle="${esc(id)}" aria-expanded="${open}" aria-label="Lihat detail"><svg class="chev" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M9 6l6 6-6 6"/></svg></button>`;

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
// Satu invoice per produk, per customer dan metode pengiriman
function buildInvoices(list) {
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
  sorted.forEach((u, ui) => {
    const groups = productGroups(u.items);
    const first = u.orders.map(o => o.sent).filter(Boolean).sort((a, b) => a - b)[0] || null;
    groups.forEach((g, gi) => {
      const orders = u.orders.filter(o => o.items.some(i => i.prod.trim().toUpperCase() === g.prod));
      out.push({
        key: u.person + '|' + u.metode + '|' + g.prod,
        person: u.person, name: u.name, metode: u.metode, metodeLabel: metodeShort(u.metode),
        pickup: u.metode || SENDER.pickup,
        first,
        invNo: `INV/${ymdCompact(first)}/${String(ui + 1).padStart(3, '0')}-P${gi + 1}`,
        prod: g.prod, part: gi + 1, partTotal: groups.length,
        colors: g.colors, pcs: g.qty, sub: g.sub,
        orderIds: g.orderIds,
        resi: [...new Set(orders.flatMap(o => o.resi))],
        notes: [...new Set(orders.map(o => o.note).filter(Boolean))],
        orders,
        mismatch: orders.some(o => o.mismatch),
        discount: sum(orders.map(o => o.discount)) > 0,
      });
    });
  });
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

// ---------- tabel ----------
function customersTableHTML() {
  if (!PERSONS.length) return emptyHTML('Tidak ada data sesuai filter.');
  const price = FORMAT === 'ekspor';
  const body = PERSONS.map(p => {
    const id = 'c:' + p.person;
    const open = EXPANDED.has(id);
    let h = `<tr class="grp">
      <td class="pl-3 w-10">${toggleBtn(id, open)}</td>
      <td><div class="font-semibold">${esc(p.name)}</div><div class="text-xs text-gray-500">${p.orderCount} order</div></td>
      <td><div class="flex flex-wrap gap-1">${p.metodes.map(m => `<span class="rounded-full px-2 py-0.5 text-[11px] font-medium ${pillCls(m)}">${esc(metodeShort(m))}</span>`).join('')}</div></td>
      <td class="num">${p.invs.length}</td>
      <td class="num">${p.pcs}</td>
      ${price ? `<td class="num font-semibold">${rp(p.total)}</td>` : ''}
      <td class="pr-4 text-right whitespace-nowrap"><button class="btn-sm" data-print-person="${esc(p.person)}">Cetak semua</button></td>
    </tr>`;
    if (open) p.invs.forEach(inv => {
      const pid = 'p:' + inv.key;
      const po = EXPANDED.has(pid);
      h += `<tr class="sub">
        <td class="pl-6">${toggleBtn(pid, po)}</td>
        <td class="py-2"><div class="font-medium text-gray-900">${esc(inv.prod)}</div><div class="text-[11px] text-gray-500">Part ${inv.part} dari ${inv.partTotal} · ${esc(inv.invNo)}</div></td>
        <td><span class="rounded-full px-2 py-0.5 text-[11px] font-medium ${pillCls(inv.metode)}">${esc(inv.metodeLabel || '—')}</span></td>
        <td></td>
        <td class="num">${inv.pcs}</td>
        ${price ? `<td class="num">${rp(inv.sub)}</td>` : ''}
        <td class="pr-4 text-right whitespace-nowrap">
          <button class="btn-sm" data-preview="${esc(inv.key)}">Preview</button>
          <button class="btn-sm" data-print-one="${esc(inv.key)}">Cetak</button>
        </td>
      </tr>`;
      if (po) inv.colors.forEach(c => {
        h += `<tr class="sub">
          <td></td>
          <td class="py-1.5 pl-9 text-gray-600">${esc(c.warna)}${price && c.price !== null ? ` <span class="text-gray-400">· ${rp(c.price)}</span>` : ''}</td>
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
    <th class="w-10"></th><th>Customer</th><th>Metode</th>
    <th class="num">Invoice</th><th class="num">Pcs</th>
    ${price ? '<th class="num">Total</th>' : ''}
    <th class="pr-4 text-right">Aksi</th></tr>`;
  return `<table class="dt"><thead>${head}</thead><tbody>${body}</tbody></table>`;
}

function rekapTableHTML() {
  if (!REKAP.length) return emptyHTML('Tidak ada data sesuai filter.');
  const split = FORMAT === 'gudang';
  const head = `<tr><th class="w-10"></th><th>Produk / Warna</th>
    ${split ? '<th class="num">Gudang Pusat</th><th class="num">Dropship</th>' : ''}
    <th class="num">Total qty</th><th class="pr-4 text-right">Order</th></tr>`;
  const body = REKAP.map(p => {
    const id = 'r:' + p.prod;
    const open = EXPANDED.has(id);
    let h = `<tr class="grp">
      <td class="pl-3">${toggleBtn(id, open)}</td>
      <td><div class="font-semibold">${esc(p.prod)}</div><div class="text-xs text-gray-500">${p.colors.length} warna</div></td>
      ${split ? `<td class="num">${p.pusat}</td><td class="num">${p.ds}</td>` : ''}
      <td class="num font-semibold">${p.total}</td>
      <td class="pr-4 text-right tabular-nums">${p.orders}</td>
    </tr>`;
    if (open) p.colors.forEach(c => {
      h += `<tr class="sub">
        <td></td>
        <td class="py-1.5 pl-9 text-gray-600">${esc(c.warna)}</td>
        ${split ? `<td class="num">${c.pusat}</td><td class="num">${c.ds}</td>` : ''}
        <td class="num">${c.total}</td>
        <td class="pr-4 text-right tabular-nums">${c.orders}</td>
      </tr>`;
    });
    return h;
  }).join('');
  return `<table class="dt"><thead>${head}</thead><tbody>${body}</tbody></table>`;
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
      <dt>Ambil</dt><dd>${esc(inv.pickup)}</dd>
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
    <div class="inv-foot strong">Terima kasih atas pesanan Anda.</div>
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
  $('st-items').textContent = sum(list.map(o => o.items.length));
  $('st-invoices').textContent = INVOICES.length;
  if (FORMAT === 'gudang') {
    $('st-total-label').textContent = 'Total pcs';
    $('st-total').textContent = pcs;
  } else {
    $('st-total-label').textContent = 'Total';
    $('st-total').textContent = rp(sum(list.map(o => o.total)));
  }
  $('cust-count').textContent = PERSONS.length ? `${PERSONS.length} customer · ${INVOICES.length} invoice` : '';
  $('list').innerHTML = customersTableHTML();
  $('rekap').innerHTML = rekapTableHTML();
  $('btn-print-all').disabled = !INVOICES.length;
  $('btn-print-rekap').disabled = !REKAP.length;

  if (SELECTED && !INVOICES.some(i => i.key === SELECTED)) closeModal();
  else if (SELECTED) refreshPreview();
}

function refreshInvoiceTable() { $('list').innerHTML = customersTableHTML(); }
function refreshRekapTable() { $('rekap').innerHTML = rekapTableHTML(); }

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
  $('btn-convert').textContent = CONVERTED ? 'Sudah dikonversi' : 'Convert to Invoice';
  $('btn-reset').hidden = !has;
}

function updateSummary() {
  const has = ALL.length > 0;
  $('fname').textContent = has ? FILE_NAME : 'Belum ada file.';
  $('fmt').textContent = !has ? 'Belum ada data'
    : FORMAT === 'gudang' ? 'Template gudang (tanpa harga)' : 'Ekspor berdu.id (dengan harga)';
  const pcs = sum(ALL.flatMap(o => o.items).map(i => i.qty));
  const persons = new Set(ALL.map(o => normName(o.reseller || o.buyer))).size;
  $('st-orders').textContent = ALL.length;
  $('st-cust').textContent = persons;
  $('st-items').textContent = sum(ALL.map(o => o.items.length));
  $('st-invoices').textContent = has ? buildInvoices(ALL).length : 0;
  if (FORMAT === 'gudang') {
    $('st-total-label').textContent = 'Total pcs';
    $('st-total').textContent = pcs;
  } else {
    $('st-total-label').textContent = 'Total';
    $('st-total').textContent = rp(sum(ALL.map(o => o.total)));
  }
  const mm = ALL.filter(o => o.mismatch).length;
  $('sum-warn').hidden = !mm;
  $('sum-warn').textContent = mm ? `⚠ ${mm} order total beda dari subtotal − diskon` : '';
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
  $('q').value = '';
  paintMetode();
}

async function loadFile(f) {
  if (!f) return;
  try {
    const rows = await fileToRows(f);
    ALL = parseOrders(rows);
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

// Delegasi: toggle baris & aksi tabel
function onTableClick(e, refresh) {
  const t = e.target.closest('[data-toggle]');
  if (t) {
    const id = t.dataset.toggle;
    if (EXPANDED.has(id)) EXPANDED.delete(id); else EXPANDED.add(id);
    refresh();
    return;
  }
  const pv = e.target.closest('[data-preview]');
  if (pv) return showPreview(pv.dataset.preview);
  const one = e.target.closest('[data-print-one]');
  if (one) return doPrint(INVOICES.filter(i => i.key === one.dataset.printOne));
  const per = e.target.closest('[data-print-person]');
  if (per) return doPrint(INVOICES.filter(i => i.person === per.dataset.printPerson));
}
$('list').addEventListener('click', e => onTableClick(e, refreshInvoiceTable));
$('rekap').addEventListener('click', e => onTableClick(e, refreshRekapTable));

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
