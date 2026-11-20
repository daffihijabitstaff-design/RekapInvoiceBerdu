// Rekap Invoice Berdu — logika aplikasi (parsing, filter, pengelompokan, render, cetak).
// Konfigurasi pengirim ada di konstanta SENDER.

const SENDER = { nama: 'Daffi Hijab', telp: '0813-3244-2527', pickup: 'DIAMBIL DI GUDANG DROPSHIP' };
const MONTHS = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
// Pola nomor resi: Shopee (SPXID..), J&T (JY/JZ..), dan SP/angka panjang
const RESI_RE = /\b(SPXID[0-9A-Z]+|JY\d{8,}|JZ\d{8,}|SP\d{10,}|\d{16,})\b/g;

let FORMAT = 'ekspor';  // 'ekspor' | 'gudang'
let ALL = [];           // semua order hasil parse
let CUST = [];          // customer hasil filter (terurut)
let SELECTED = null;    // key customer yang sedang dipreview

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

// ---------- filter & group ----------
const dOf = s => { if (!s) return null; const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d); };

function filtered() {
  const q = $('q').value.trim().toLowerCase();
  const st = $('status').value, mt = $('metode').value;
  const from = dOf($('from').value), to = dOf($('to').value);
  if (to) to.setHours(23, 59, 59, 999);
  return ALL.filter(o => {
    if (st && o.status !== st) return false;
    if (mt && o.metode !== mt) return false;
    if (from && (!o.sent || o.sent < from)) return false;
    if (to && (!o.sent || o.sent > to)) return false;
    if (q) {
      const hay = [o.id, o.buyer, o.reseller, o.note, ...o.resi,
                   ...o.items.flatMap(i => [i.prod, i.variasi, i.sku])].join(' ').toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });
}

// Label metode pendek untuk badge
const metodeShort = m => /PUSAT/i.test(m) ? 'Gudang Pusat' : /DROPSHIP/i.test(m) ? 'Dropship' : '';

// Kelompokkan item per produk, lalu per warna di dalamnya
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

// Satu invoice = satu produk milik satu customer (per metode pengiriman).
// Customer dengan 5 produk menghasilkan 5 invoice (P1..P5).
function customers(list) {
  const units = new Map();
  for (const o of list) {
    const raw = o.reseller || o.buyer || '(tanpa nama)';
    const person = normName(raw) || '(tanpa nama)';
    const metode = o.metode || '';
    const key = person + '|' + metode;
    if (!units.has(key)) units.set(key, { key, person, metode, name: raw.trim().replace(/\s+/g, ' '), orders: [], items: [] });
    const u = units.get(key);
    u.orders.push(o);
    u.items.push(...o.items);
  }
  const sorted = [...units.values()].sort((a, b) => a.name.localeCompare(b.name) || a.metode.localeCompare(b.metode));
  const out = [];
  sorted.forEach((u, ui) => {
    const groups = productGroups(u.items);
    const dates = u.orders.map(o => o.sent).filter(Boolean).sort((a, b) => a - b);
    const first = dates[0] || null;
    groups.forEach((g, gi) => {
      const orders = u.orders.filter(o => o.items.some(i => i.prod.trim().toUpperCase() === g.prod));
      out.push({
        key: u.key + '|' + g.prod,
        person: u.person, name: u.name, metode: u.metode,
        metodeLabel: metodeShort(u.metode),
        pickup: u.metode || SENDER.pickup,
        first,
        invNo: `INV/${ymdCompact(first)}/${String(ui + 1).padStart(3, '0')}-P${gi + 1}`,
        prod: g.prod, part: gi + 1, partTotal: groups.length,
        colors: g.colors, pcs: g.qty, sub: g.sub, total: g.sub,
        orderIds: g.orderIds,
        orders, items: u.items.filter(i => i.prod.trim().toUpperCase() === g.prod),
        mismatch: orders.filter(o => o.mismatch)
      });
    });
  });
  out.forEach((inv, i) => inv.no = i + 1);
  return out;
}

// Gudang tidak punya harga, jadi invoice selalu tanpa harga
const hidePrice = () => FORMAT === 'gudang';

// ---------- invoice markup (satu produk per invoice) ----------
function invoiceHTML(c) {
  const hp = hidePrice();
  const orderLabel = FORMAT === 'gudang' ? 'Order (Nomor)' : 'Order ID';
  const body = c.colors.map(r => hp
    ? `<tr><td>${esc(r.warna)}</td><td class="c">${r.qty}x</td></tr>`
    : `<tr><td>${esc(r.warna)}</td><td class="c">${r.qty}x</td><td class="r">${r.price === null ? '—' : rp(r.price)}</td><td class="r">${rp(r.sub)}</td></tr>`
  ).join('');
  const thead = hp
    ? `<tr><th>Warna</th><th class="c">Qty</th></tr>`
    : `<tr><th>Warna</th><th class="c">Qty</th><th class="r">Harga</th><th class="r">Subtotal</th></tr>`;
  const tfoot = hp
    ? `<tr><td>Total</td><td class="c">${c.pcs}x</td></tr>`
    : `<tr><td>Total</td><td class="c">${c.pcs}x</td><td></td><td class="r">${rp(c.sub)}</td></tr>`;
  return `
  <div class="sheet">
    <div class="sheet-head">
      <span>${fmtDate(c.first)}</span>
      <span class="inv">${esc(c.invNo)}</span>
    </div>
    <div class="blk">
      <div class="lbl">Pengirim</div>
      <b>${esc(SENDER.nama)} (${esc(SENDER.telp)})</b>
      <div class="pickup">${esc(c.pickup)}</div>
    </div>
    <div class="blk">
      <div class="lbl">Kepada</div>
      <b>${esc(c.name)}</b>
      <div class="part">Part ${c.part} dari ${c.partTotal}</div>
    </div>
    <div class="blk">
      <div class="lbl">Produk</div>
      <div class="prod">${esc(c.prod)}</div>
      <div class="ref">${orderLabel}: ${esc(c.orderIds.join(', '))}</div>
    </div>
    <table class="items">
      <thead>${thead}</thead>
      <tbody>${body}</tbody>
      <tfoot>${tfoot}</tfoot>
    </table>
  </div>`;
}

// ---------- rekap produk (untuk inventory / picking) ----------
function rekapHTML(list) {
  const map = new Map();
  for (const o of list) for (const i of o.items) {
    const key = i.prod.toUpperCase() + '|' + fmtWarna(i.variasi);
    if (!map.has(key)) map.set(key, { prod: i.prod.toUpperCase(), warna: fmtWarna(i.variasi), pusat: 0, ds: 0, total: 0, orders: new Set() });
    const r = map.get(key);
    r.total += i.qty;
    if (i.metode.includes('PUSAT')) r.pusat += i.qty;
    else if (i.metode.includes('DROPSHIP')) r.ds += i.qty;
    r.orders.add(o.id);
  }
  const rows = [...map.values()].sort((a, b) => a.prod.localeCompare(b.prod) || a.warna.localeCompare(b.warna));
  if (!rows.length) return '<p class="muted">Tidak ada data sesuai filter.</p>';
  const split = FORMAT === 'gudang';
  return `
  <table class="rekap">
    <thead><tr><th>Produk</th><th>Warna</th>${split ? '<th class="n">Gudang Pusat</th><th class="n">Dropship</th>' : ''}<th class="n">Total qty</th><th class="n">Order</th></tr></thead>
    <tbody>${rows.map(r => `<tr><td>${esc(r.prod)}</td><td>${esc(r.warna)}</td>${split ? `<td class="n">${r.pusat}</td><td class="n">${r.ds}</td>` : ''}<td class="n"><b>${r.total}</b></td><td class="n">${r.orders.size}</td></tr>`).join('')}</tbody>
  </table>`;
}

// ---------- render ----------
function renderAll() {
  const list = filtered();
  CUST = customers(list);

  const pcs = sum(list.flatMap(o => o.items).map(i => i.qty));
  $('st-orders').textContent = list.length;
  $('st-cust').textContent = new Set(CUST.map(c => c.person)).size;
  $('st-items').textContent = sum(list.map(o => o.items.length));
  if (FORMAT === 'gudang') {
    $('st-total-label').textContent = 'Total pcs';
    $('st-total').textContent = pcs;
  } else {
    $('st-total-label').textContent = 'Total';
    $('st-total').textContent = rp(sum(list.map(o => o.total)));
  }
  $('cust-count').textContent = CUST.length ? `(${CUST.length})` : '';

  $('list').innerHTML = CUST.map(c => `
    <div class="card ${SELECTED === c.key ? 'sel' : ''}">
      <div>
        <b>${c.no}. ${esc(c.name)}</b>
        <div class="muted">${esc(c.prod)} · part ${c.part}/${c.partTotal} · ${c.pcs} pcs</div>
        ${c.metodeLabel ? `<span class="pill">${esc(c.metodeLabel)}</span>` : ''}
        ${c.mismatch.length ? `<div class="warn">⚠ ${c.mismatch.length} order total beda dari CSV</div>` : ''}
      </div>
      <div class="r">
        ${FORMAT === 'ekspor' ? `<b>${rp(c.total)}</b>` : ''}
        <button data-key="${esc(c.key)}">Preview</button>
      </div>
    </div>`).join('') || '<p class="muted">Tidak ada data sesuai filter.</p>';

  $('rekap').innerHTML = rekapHTML(list);
  $('btn-print-all').disabled = !CUST.length;
  if (SELECTED && !CUST.some(c => c.key === SELECTED)) { closeModal(); return; }
  if (SELECTED) refreshPreview();
}

function refreshPreview() {
  const c = CUST.find(x => x.key === SELECTED);
  if (!c) return;
  $('pv-title').textContent = `${c.name} · ${c.prod}` + (c.metodeLabel ? ` · ${c.metodeLabel}` : '');
  $('pv').innerHTML = invoiceHTML(c);
  $('warn').innerHTML = c.mismatch.length
    ? `<div class="warnbox">⚠ Order berikut total di file tidak sama dengan subtotal − diskon (kemungkinan ada voucher/biaya lain). Invoice memakai nilai Total dari file:<br>${c.mismatch.map(o => esc(o.id)).join('<br>')}</div>`
    : '';
}

function openModal() {
  $('modal').hidden = false;
  document.body.classList.add('modal-open');
}
function closeModal() {
  $('modal').hidden = true;
  document.body.classList.remove('modal-open');
  SELECTED = null;
  renderAll();
}

function showPreview(key) {
  SELECTED = key;
  openModal();
  renderAll();
}

function doPrint(list) {
  $('print-area').innerHTML = list.map(invoiceHTML).join('');
  window.print();
}

function doPrintRekap() {
  const list = filtered();
  $('print-area').innerHTML = `<div class="sheet"><div class="sheet-head"><span>Rekap produk</span><span class="inv">${fmtDate(new Date())}</span></div>${rekapHTML(list)}</div>`;
  window.print();
}

// ---------- tahap aplikasi: upload -> konversi -> hasil ----------
let CONVERTED = false;
let FILE_NAME = '';
const wait = ms => new Promise(r => setTimeout(r, ms));
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

// Ringkasan di sidebar kanan, dihitung dari file sebelum konversi
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
  $('st-invoices').textContent = has ? customers(ALL).length : 0;
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

function fillFilterOptions() {
  const uniq = k => [...new Set(ALL.map(o => o[k]).filter(Boolean))];
  $('status').innerHTML = '<option value="">Semua</option>' + uniq('status').map(s => `<option>${esc(s)}</option>`).join('');
  $('metode').innerHTML = '<option value="">Semua</option>' + uniq('metode').map(s => `<option>${esc(s)}</option>`).join('');
  $('q').value = ''; $('from').value = ''; $('to').value = '';
}

async function loadFile(f) {
  if (!f) return;
  try {
    const rows = await fileToRows(f);
    ALL = parseOrders(rows);
    FILE_NAME = f.name;
    CONVERTED = false;
    SELECTED = null;
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

// Konversi: setiap langkah menjalankan pekerjaan nyata, lalu memperbarui progress
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
    if (i === 1) invoices = customers(ALL).length;   // pengelompokan
    if (i === 2) rekapHTML(ALL);                      // rekap produk
    setStep(i, 'done');
    setBar(Math.round(((i + 1) / n) * 100));
    await wait(180);
  }

  $('loader-title').textContent = 'Selesai';
  $('loader-step').textContent = `${invoices} invoice siap`;
  await wait(500);
  $('loader').hidden = true;
  CONVERTED = true;
  setViews();
  renderAll();
}

function resetAll() {
  ALL = [];
  FILE_NAME = '';
  CONVERTED = false;
  SELECTED = null;
  $('modal').hidden = true;
  document.body.classList.remove('modal-open');
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

['q', 'status', 'metode', 'from', 'to'].forEach(id =>
  $(id).addEventListener(id === 'q' ? 'input' : 'change', () => { if (CONVERTED) renderAll(); }));

$('list').addEventListener('click', e => {
  const b = e.target.closest('button[data-key]');
  if (b) showPreview(b.dataset.key);
});

$('btn-print-one').addEventListener('click', () => {
  const c = CUST.find(x => x.key === SELECTED);
  if (c) doPrint([c]);
});
$('btn-print-all').addEventListener('click', () => doPrint(CUST));
$('btn-print-rekap').addEventListener('click', doPrintRekap);

document.querySelectorAll('[data-close]').forEach(el => el.addEventListener('click', closeModal));
document.addEventListener('keydown', e => { if (e.key === 'Escape' && !$('modal').hidden) closeModal(); });

setViews();
updateSummary();
