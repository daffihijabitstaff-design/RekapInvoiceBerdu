# Rekap Invoice Berdu

Aplikasi statis untuk membaca ekspor order berdu.id (dengan harga) dan template gudang (tanpa harga), lalu:

- menampilkan ringkasan dan filter (metode pengiriman dan pencarian nama),
- membuat **rekap produk** per produk dan warna, untuk daftar ambil barang,
- membuat **satu invoice per produk** untuk setiap customer, dengan rincian warna dan qty,
- mencetak invoice dalam ukuran A5.

Semua proses terjadi di browser. File tidak diunggah ke server.

## Teknologi

- **Node.js** (>= 18) untuk tooling dan server lokal (tanpa framework)
- **Tailwind CSS 3** untuk styling, dikompilasi dari `src/input.css`
- **SheetJS** (`xlsx`) untuk membaca file XLSX; CSV dibaca langsung

## Struktur

```
.
├── public/
│   ├── index.html        # markup halaman, memakai utilitas Tailwind
│   └── js/app.js         # logika: parsing, filter, pengelompokan, render, cetak
├── src/
│   └── input.css         # sumber Tailwind (base, components, print)
├── scripts/
│   └── copy-vendor.js    # menyalin xlsx dari node_modules ke public/vendor
├── server.js             # server statis untuk pengembangan lokal
├── tailwind.config.js
├── vercel.json
└── docs/FLOW.md          # alur admin dan arsitektur developer
```

`public/css/styles.css` dan `public/vendor/` dibuat ulang oleh `npm run build` dan tidak disimpan di git.

## Menjalankan lokal

```bash
npm install
npm run dev        # build lalu jalankan server di http://localhost:3000
```

Untuk mengedit styling sambil berjalan, buka terminal kedua:

```bash
npm run watch:css
```

## Script

| Perintah | Fungsi |
|---|---|
| `npm run build` | Kompilasi Tailwind (minified) dan salin xlsx ke `public/vendor` |
| `npm run dev` | `build` lalu jalankan `server.js` |
| `npm run watch:css` | Kompilasi Tailwind otomatis saat file berubah |
| `npm start` | Jalankan `server.js` saja (perlu `build` sebelumnya) |

## Deploy ke Vercel

1. Import repo ini di Vercel (Framework: Other).
2. Vercel memakai `vercel.json`: `npm install`, lalu `npm run build`, dan menyajikan folder `public`.
3. Tidak perlu environment variable.

## Konfigurasi

Data pengirim ada di konstanta `SENDER` pada bagian atas `public/js/app.js`:

```js
const SENDER = { nama: 'Daffi Hijab', telp: '0813-3244-2527', pickup: 'DIAMBIL DI GUDANG DROPSHIP' };
```

## Format file yang didukung

- **Ekspor berdu.id**: header berisi `Order Id`. Harga dan diskon ikut dibaca.
- **Template gudang**: header berisi `Nomor`, `Nama Produk`, dan `Metode Pengiriman`. Tidak ada harga.

Format dikenali otomatis dari nama kolom.

## Privasi

- File ekspor (`*.csv`, `*.xlsx`, `*.xls`) berisi data pembeli. File-file ini sudah di-ignore git dan tidak boleh di-commit.
- Aplikasi ini tidak punya autentikasi. Siapa pun yang membuka URL deploy bisa memakai aplikasi, tetapi data pembeli tetap hanya ada di browser mereka selama file diunggah secara lokal.

## Pengembangan lanjut

Lihat [docs/FLOW.md](docs/FLOW.md) bagian B.9 untuk roadmap. Langkah yang paling disarankan berikutnya: memisahkan fungsi parsing dari DOM dan menambahkan test dengan fixture yang sudah dianonimkan.
