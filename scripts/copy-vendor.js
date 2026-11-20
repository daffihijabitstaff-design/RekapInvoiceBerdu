// Salin library pihak ketiga dari node_modules ke public/vendor
// supaya bisa dimuat dari origin sendiri (tanpa CDN).
const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..');
const copies = [
  {
    from: path.join(root, 'node_modules', 'xlsx', 'dist', 'xlsx.full.min.js'),
    to: path.join(root, 'public', 'vendor', 'xlsx.full.min.js'),
  },
];

for (const { from, to } of copies) {
  if (!fs.existsSync(from)) {
    console.error(`Tidak ditemukan: ${from}. Jalankan "npm install" dulu.`);
    process.exit(1);
  }
  fs.mkdirSync(path.dirname(to), { recursive: true });
  fs.copyFileSync(from, to);
  console.log(`copied ${path.relative(root, to)}`);
}
