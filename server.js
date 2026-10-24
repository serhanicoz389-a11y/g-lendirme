// Dilek Ofis - Kırtasiye Fiyat Kataloğu
// Harici paket gerektirmez: sadece Node.js (v18+) yeterlidir.
//   Başlatmak için:  node server.js
//   Müşteri sayfası: http://<bilgisayar-ip>:3000/
//   Yönetim paneli:  http://localhost:3000/admin

const http = require('http');
const fs = require('fs');
const path = require('path');
const os = require('os');
const crypto = require('crypto');

const PORT = Number(process.env.PORT) || 3000;
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || 'dilek123';
const DATA_DIR = path.join(__dirname, 'data');
const PUBLIC_DIR = path.join(__dirname, 'public');
const PRODUCTS_FILE = path.join(DATA_DIR, 'products.json');
const SETTINGS_FILE = path.join(DATA_DIR, 'settings.json');

const DEFAULT_SETTINGS = {
  shopName: 'Dilek Ofis',
  slogan: 'Kırtasiye & Ofis Malzemeleri',
  phone: '',
  address: '',
  publicUrl: '',
  orderNo: 1,
};

// ---------- Veri dosyaları ----------

function readJson(file, fallback) {
  try {
    return JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch {
    return fallback;
  }
}

function writeJson(file, data) {
  // Önce geçici dosyaya yazıp sonra taşıyoruz; elektrik kesilirse veri bozulmaz.
  const tmp = file + '.tmp';
  fs.writeFileSync(tmp, JSON.stringify(data, null, 2), 'utf8');
  fs.renameSync(tmp, file);
}

const loadProducts = () => readJson(PRODUCTS_FILE, []);
const saveProducts = (list) => writeJson(PRODUCTS_FILE, list);
const loadSettings = () => ({ ...DEFAULT_SETTINGS, ...readJson(SETTINGS_FILE, {}) });
const saveSettings = (s) => writeJson(SETTINGS_FILE, s);

// Kategoriler: ayarlardaki sıralı liste + ürünlerde geçip listede olmayanlar.
function categoryList(products = loadProducts(), settings = loadSettings()) {
  const list = [...(settings.categories || [])];
  for (const p of products) if (!list.includes(p.category)) list.push(p.category);
  return list;
}

function barcodeOwner(products, barcode, exceptId) {
  return barcode ? products.find((p) => p.barcode === barcode && p.id !== exceptId) : undefined;
}

function cleanProduct(input, existing = {}) {
  const p = { ...existing };
  if (input.name !== undefined) p.name = String(input.name).trim();
  if (input.code !== undefined) p.code = String(input.code).trim();
  if (input.barcode !== undefined) p.barcode = String(input.barcode).replace(/\s+/g, '');
  if (input.category !== undefined) p.category = String(input.category).trim() || 'Diğer';
  if (input.unit !== undefined) p.unit = String(input.unit).trim() || 'Adet';
  if (input.price !== undefined) {
    const price = Number(String(input.price).replace(',', '.'));
    p.price = Number.isFinite(price) && price >= 0 ? Math.round(price * 100) / 100 : 0;
  }
  if (input.active !== undefined) p.active = Boolean(input.active);
  p.category ??= 'Diğer';
  p.unit ??= 'Adet';
  p.code ??= '';
  p.barcode ??= '';
  p.price ??= 0;
  p.active ??= true;
  p.updatedAt = new Date().toISOString();
  return p;
}

// ---------- CSV (Excel ile uyumlu, noktalı virgül ayraçlı) ----------

const CSV_HEADERS = ['Kod', 'Barkod', 'Ürün Adı', 'Kategori', 'Birim', 'Fiyat', 'Aktif'];

function toCsv(products) {
  const esc = (v) => {
    const s = String(v ?? '');
    return /[;"\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  const rows = products.map((p) =>
    [p.code, p.barcode, p.name, p.category, p.unit, String(p.price).replace('.', ','), p.active ? 'Evet' : 'Hayır']
      .map(esc)
      .join(';')
  );
  return '﻿' + [CSV_HEADERS.join(';'), ...rows].join('\r\n');
}

function parseCsv(text) {
  text = text.replace(/^﻿/, '');
  const firstLine = text.split(/\r?\n/, 1)[0];
  const sep = firstLine.split(';').length >= firstLine.split(',').length ? ';' : ',';
  const rows = [];
  let row = [], field = '', quoted = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (quoted) {
      if (c === '"' && text[i + 1] === '"') { field += '"'; i++; }
      else if (c === '"') quoted = false;
      else field += c;
    } else if (c === '"') quoted = true;
    else if (c === sep) { row.push(field); field = ''; }
    else if (c === '\n' || c === '\r') {
      if (c === '\r' && text[i + 1] === '\n') i++;
      row.push(field); rows.push(row); row = []; field = '';
    } else field += c;
  }
  if (field || row.length) { row.push(field); rows.push(row); }
  return rows.filter((r) => r.some((f) => f.trim() !== ''));
}

function importCsv(text) {
  const rows = parseCsv(text);
  if (rows.length < 2) throw new Error('Dosyada ürün satırı bulunamadı.');
  const header = rows[0].map((h) => h.trim().toLowerCase());
  const col = (...names) => header.findIndex((h) => names.some((n) => h.startsWith(n)));
  const iCode = col('kod'), iBarcode = col('barkod'), iName = col('ürün', 'urun', 'ad'), iCat = col('kategori');
  const iUnit = col('birim'), iPrice = col('fiyat'), iActive = col('aktif');
  if (iName < 0 || iPrice < 0) throw new Error('"Ürün Adı" ve "Fiyat" sütunları bulunmalı.');

  const products = loadProducts();
  let added = 0, updated = 0;
  for (const r of rows.slice(1)) {
    const name = (r[iName] || '').trim();
    if (!name) continue;
    const code = iCode >= 0 ? (r[iCode] || '').trim() : '';
    const barcode = iBarcode >= 0 ? (r[iBarcode] || '').replace(/\s+/g, '') : '';
    // Fiyat "1.250,50" veya "1250.50" biçiminde olabilir.
    let priceText = (r[iPrice] || '0').replace(/[^\d.,]/g, '');
    if (priceText.includes(',')) priceText = priceText.replace(/\./g, '').replace(',', '.');
    const input = { name, code, price: priceText };
    if (iBarcode >= 0) input.barcode = barcode;
    if (iCat >= 0) input.category = r[iCat] || '';
    if (iUnit >= 0) input.unit = r[iUnit] || '';
    if (iActive >= 0) input.active = !/^(hayır|hayir|0|false|pasif)$/i.test((r[iActive] || '').trim());

    // Eşleştirme sırası: barkod, ürün kodu, ürün adı
    const match =
      (barcode && products.find((p) => p.barcode === barcode)) ||
      (code && products.find((p) => p.code === code)) ||
      (!code && products.find((p) => p.name.toLocaleLowerCase('tr') === name.toLocaleLowerCase('tr')));
    if (match) { Object.assign(match, cleanProduct(input, match)); updated++; }
    else { products.push(cleanProduct(input, { id: crypto.randomUUID() })); added++; }
  }
  saveProducts(products);
  return { added, updated };
}

// ---------- Yardımcılar ----------

// Telefonların ulaşabileceği yerel adresler; en olası olan (Wi-Fi/Ethernet, 192.168.x) başta.
// Sanal ağ kartları (VirtualBox, Hyper-V, VPN…) telefondan erişilemez, sona atılır.
function lanAddresses() {
  const virtual = /virtual|vmware|vbox|hyper-v|vethernet|wsl|docker|loopback|bluetooth|tailscale|zerotier|vpn|tap|tun/i;
  const out = [];
  for (const [name, list] of Object.entries(os.networkInterfaces())) {
    for (const a of list || []) {
      if (a.family !== 'IPv4' || a.internal || a.address.startsWith('169.254.')) continue;
      let score = 0;
      if (virtual.test(name)) score -= 10;
      if (/wi-?fi|wlan|kablosuz|wireless/i.test(name)) score += 3;
      if (/ethernet|eth|en\d/i.test(name) && !virtual.test(name)) score += 2;
      if (a.address.startsWith('192.168.')) score += 2;
      else if (a.address.startsWith('10.')) score += 1;
      out.push({ url: `http://${a.address}:${PORT}/`, score });
    }
  }
  return out.sort((a, b) => b.score - a.score).map((x) => x.url);
}

function send(res, status, body, type = 'application/json; charset=utf-8', extra = {}) {
  res.writeHead(status, { 'Content-Type': type, 'Cache-Control': 'no-store', ...extra });
  res.end(typeof body === 'string' || Buffer.isBuffer(body) ? body : JSON.stringify(body));
}

function readBody(req, limit = 5 * 1024 * 1024) {
  return new Promise((resolve, reject) => {
    let size = 0;
    const chunks = [];
    req.on('data', (c) => {
      size += c.length;
      if (size > limit) { reject(new Error('Gönderilen veri çok büyük.')); req.destroy(); }
      else chunks.push(c);
    });
    req.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')));
    req.on('error', reject);
  });
}

async function readJsonBody(req) {
  const text = await readBody(req);
  return text ? JSON.parse(text) : {};
}

function isAdmin(req) {
  const given = Buffer.from(String(req.headers['x-admin-password'] || ''));
  const expected = Buffer.from(ADMIN_PASSWORD);
  return given.length === expected.length && crypto.timingSafeEqual(given, expected);
}

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.ico': 'image/x-icon',
};

function serveStatic(res, urlPath) {
  const routes = { '/': 'index.html', '/admin': 'admin.html' };
  const rel = routes[urlPath] || urlPath.replace(/^\/+/, '');
  const file = path.normalize(path.join(PUBLIC_DIR, rel));
  if (!file.startsWith(PUBLIC_DIR) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) {
    return send(res, 404, 'Sayfa bulunamadı', 'text/plain; charset=utf-8');
  }
  send(res, 200, fs.readFileSync(file), MIME[path.extname(file)] || 'application/octet-stream');
}

// ---------- API ----------

async function handleApi(req, res, url) {
  const { pathname } = url;
  const method = req.method;

  // Müşterinin gördüğü katalog (sadece aktif ürünler)
  if (pathname === '/api/catalog' && method === 'GET') {
    const settings = loadSettings();
    const { shopName, slogan, phone, address } = settings;
    const all = loadProducts();
    const products = all
      .filter((p) => p.active)
      .map(({ id, name, barcode, category, unit, price, updatedAt }) => ({ id, name, barcode, category, unit, price, updatedAt }));
    const categories = categoryList(all, settings);
    return send(res, 200, { settings: { shopName, slogan, phone, address }, categories, products });
  }

  if (pathname === '/api/login' && method === 'POST') {
    const { password } = await readJsonBody(req);
    req.headers['x-admin-password'] = password;
    return isAdmin(req) ? send(res, 200, { ok: true }) : send(res, 401, { error: 'Şifre hatalı.' });
  }

  if (!pathname.startsWith('/api/admin/')) return send(res, 404, { error: 'Bulunamadı' });
  if (!isAdmin(req)) return send(res, 401, { error: 'Yetkisiz. Lütfen tekrar giriş yapın.' });

  const route = pathname.slice('/api/admin/'.length);
  const idMatch = route.match(/^products\/([\w-]+)$/);

  if (route === 'products' && method === 'GET') return send(res, 200, loadProducts());

  if (route === 'products' && method === 'POST') {
    const input = await readJsonBody(req);
    if (!String(input.name || '').trim()) return send(res, 400, { error: 'Ürün adı boş olamaz.' });
    const products = loadProducts();
    const product = cleanProduct(input, { id: crypto.randomUUID() });
    const owner = barcodeOwner(products, product.barcode);
    if (owner) return send(res, 409, { error: `Bu barkod zaten "${owner.name}" ürününde kayıtlı.` });
    products.push(product);
    saveProducts(products);
    return send(res, 201, product);
  }

  if (idMatch && method === 'PUT') {
    const products = loadProducts();
    const i = products.findIndex((p) => p.id === idMatch[1]);
    if (i < 0) return send(res, 404, { error: 'Ürün bulunamadı.' });
    const updated = cleanProduct(await readJsonBody(req), products[i]);
    const owner = barcodeOwner(products, updated.barcode, updated.id);
    if (owner) return send(res, 409, { error: `Bu barkod zaten "${owner.name}" ürününde kayıtlı.` });
    products[i] = updated;
    saveProducts(products);
    return send(res, 200, products[i]);
  }

  if (idMatch && method === 'DELETE') {
    const products = loadProducts();
    const next = products.filter((p) => p.id !== idMatch[1]);
    if (next.length === products.length) return send(res, 404, { error: 'Ürün bulunamadı.' });
    saveProducts(next);
    return send(res, 200, { ok: true });
  }

  // Toplu fiyat güncelleme: örn. tüm ürünlere %10 zam
  if (route === 'bulk-price' && method === 'POST') {
    const { percent, category, roundTo } = await readJsonBody(req);
    const pct = Number(percent);
    if (!Number.isFinite(pct) || pct <= -100) return send(res, 400, { error: 'Geçerli bir yüzde girin.' });
    const step = Number(roundTo) > 0 ? Number(roundTo) : 0.01;
    const products = loadProducts();
    let count = 0;
    for (const p of products) {
      if (category && p.category !== category) continue;
      const raised = p.price * (1 + pct / 100);
      p.price = Math.round(Math.round(raised / step) * step * 100) / 100;
      p.updatedAt = new Date().toISOString();
      count++;
    }
    saveProducts(products);
    return send(res, 200, { count });
  }

  // ---- Kategoriler ----
  if (route === 'categories' && method === 'GET') return send(res, 200, categoryList());

  if (route === 'categories' && method === 'POST') {
    // action: add | rename | delete | order
    const { action, name, to, list } = await readJsonBody(req);
    const products = loadProducts();
    const s = loadSettings();
    let cats = categoryList(products, s);
    const clean = (v) => String(v || '').trim();

    if (action === 'add') {
      const n = clean(name);
      if (!n) return send(res, 400, { error: 'Kategori adı boş olamaz.' });
      if (cats.includes(n)) return send(res, 409, { error: 'Bu kategori zaten var.' });
      cats.push(n);
    } else if (action === 'rename') {
      const from = clean(name), n = clean(to);
      if (!n) return send(res, 400, { error: 'Yeni ad boş olamaz.' });
      for (const p of products) if (p.category === from) p.category = n;
      // Aynı ada sahip kategori varsa ikisi birleşir.
      cats = cats.includes(n) ? cats.filter((c) => c !== from) : cats.map((c) => (c === from ? n : c));
      saveProducts(products);
    } else if (action === 'delete') {
      const n = clean(name), target = clean(to);
      const used = products.filter((p) => p.category === n);
      if (used.length && !target) return send(res, 400, { error: 'Bu kategoride ürün var; taşınacak kategoriyi seçin.' });
      for (const p of used) p.category = target;
      cats = cats.filter((c) => c !== n);
      if (target && !cats.includes(target)) cats.push(target);
      if (used.length) saveProducts(products);
    } else if (action === 'order') {
      if (!Array.isArray(list)) return send(res, 400, { error: 'Liste gerekli.' });
      cats = [...new Set([...list.map(clean).filter((c) => cats.includes(c)), ...cats])];
    } else return send(res, 400, { error: 'Bilinmeyen işlem.' });

    s.categories = cats;
    saveSettings(s);
    return send(res, 200, categoryList(loadProducts(), s));
  }

  if (route === 'settings' && method === 'GET') return send(res, 200, loadSettings());

  if (route === 'settings' && method === 'PUT') {
    const input = await readJsonBody(req);
    const s = loadSettings();
    for (const k of ['shopName', 'slogan', 'phone', 'address', 'publicUrl']) {
      if (input[k] !== undefined) s[k] = String(input[k]).trim();
    }
    saveSettings(s);
    return send(res, 200, s);
  }

  // Sipariş formu numarası: her yazdırmada bir artar
  if (route === 'next-order-no' && method === 'POST') {
    const s = loadSettings();
    const no = s.orderNo;
    s.orderNo = no + 1;
    saveSettings(s);
    return send(res, 200, { orderNo: no });
  }

  if (route === 'info' && method === 'GET') return send(res, 200, { lanUrls: lanAddresses(), port: PORT });

  if (route === 'export.csv' && method === 'GET') {
    const date = new Date().toISOString().slice(0, 10);
    return send(res, 200, toCsv(loadProducts()), 'text/csv; charset=utf-8', {
      'Content-Disposition': `attachment; filename="dilek-ofis-urunler-${date}.csv"`,
    });
  }

  if (route === 'import' && method === 'POST') {
    try {
      return send(res, 200, importCsv(await readBody(req)));
    } catch (e) {
      return send(res, 400, { error: e.message });
    }
  }

  return send(res, 404, { error: 'Bulunamadı' });
}

// ---------- Sunucu ----------

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
if (!fs.existsSync(PRODUCTS_FILE)) saveProducts([]);

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, 'http://localhost');
  try {
    if (url.pathname.startsWith('/api/')) await handleApi(req, res, url);
    else if (req.method === 'GET') serveStatic(res, url.pathname);
    else send(res, 405, 'İzin verilmeyen istek', 'text/plain; charset=utf-8');
  } catch (e) {
    console.error(e);
    send(res, 500, { error: 'Sunucu hatası: ' + e.message });
  }
});

server.listen(PORT, () => {
  console.log('\n  Dilek Ofis Fiyat Kataloğu çalışıyor\n');
  console.log(`  Yönetim paneli : http://localhost:${PORT}/admin`);
  const [best, ...others] = lanAddresses();
  if (best) console.log(`  Müşteri sayfası: ${best}   <- QR kod bu adrese yönlenmeli`);
  for (const u of others) console.log(`  Diğer adres    : ${u}`);
  if (!best) console.log('  Uyarı: Bilgisayar bir ağa (Wi-Fi/kablo) bağlı görünmüyor.');
  const { publicUrl } = loadSettings();
  if (publicUrl && /^http:\/\/(\d+\.){3}\d+/.test(publicUrl) && !lanAddresses().includes(publicUrl)) {
    console.log(`\n  !!! DİKKAT: Basılı QR kod ${publicUrl} adresine yönleniyor,`);
    console.log('  !!! ama bu bilgisayarın adresi değişmiş. Müşteriler sayfayı açamaz.');
    console.log('  !!! Çözüm: Yönetim paneli > QR Kod sekmesindeki talimatlara bakın.');
  }
  if (ADMIN_PASSWORD === 'dilek123') console.log('\n  Uyarı: Varsayılan şifre (dilek123) kullanılıyor. ADMIN_PASSWORD ile değiştirin.');
  console.log('\n  Kapatmak için bu pencerede Ctrl+C\n');
});
