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
const { exec } = require('child_process');

let PORT = Number(process.env.PORT) || 3000;
const START_PORT = PORT;
const SITE_URL = 'https://dilekofis.github.io/';
const SITE_REPO = 'dilekofis/dilekofis.github.io';
const LOG_FILE = path.join(__dirname, 'hata-kaydi.txt');

// Beklenmeyen hatalar hem ekrana hem hata-kaydi.txt dosyasına yazılır.
function logError(e) {
  const line = `[${new Date().toLocaleString('tr-TR')}] ${e && e.stack ? e.stack : e}\n`;
  console.error('\n  HATA: ' + (e && e.message ? e.message : e));
  try { fs.appendFileSync(LOG_FILE, line); } catch {}
}
process.on('uncaughtException', (e) => { logError(e); process.exitCode = 1; });
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || 'dilek123';
const DATA_DIR = path.join(__dirname, 'data');
const PUBLIC_DIR = path.join(__dirname, 'public');
const PRODUCTS_FILE = path.join(DATA_DIR, 'products.json');
const SETTINGS_FILE = path.join(DATA_DIR, 'settings.json');
const ORDERS_FILE = path.join(DATA_DIR, 'orders.json');
const ORDER_REPO = 'dilekofis/siparisler';

const DEFAULT_SETTINGS = {
  shopName: 'Dilek Ofis',
  slogan: 'Kırtasiye & Ofis Malzemeleri',
  phone: '',
  address: '',
  publicUrl: SITE_URL,
  orderNo: 1,
  // İnternette yayınlama (GitHub + Cloudflare Pages)
  githubRepo: SITE_REPO,
  githubBranch: 'main',
  githubToken: '',
  siteUrl: SITE_URL,
  autoPublish: true,
  // İnternetten sipariş alma: müşteri siparişi bu (gizli) depoya şifreli "issue" olarak bırakır, program alır.
  orderRepo: ORDER_REPO,
  orderToken: '',
  ordersEnabled: true,
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
const saveProducts = (list) => { writeJson(PRODUCTS_FILE, list); schedulePublish(); };
// Eski adresler (pages.dev bazı internet şirketlerinde açılmıyor; ardından github.io/g-lendirme) kayıtlı ayarlarda
// varsa yenisine çevrilir. Site artık "dilekofis" kuruluşunun dilekofis.github.io deposundan yayınlanır.
const OLD_SITE_URL = /^https?:\/\/(dilekofis\.pages\.dev|serhanicoz389-a11y\.github\.io\/g-lendirme)\/?$/;
const OLD_SITE_REPO = 'serhanicoz389-a11y/g-lendirme';
function loadSettings() {
  const s = { ...DEFAULT_SETTINGS, ...readJson(SETTINGS_FILE, {}) };
  if (OLD_SITE_URL.test(s.siteUrl || '')) s.siteUrl = SITE_URL;
  if (OLD_SITE_URL.test(s.publicUrl || '')) s.publicUrl = SITE_URL;
  if ((s.githubRepo || '').trim() === OLD_SITE_REPO) s.githubRepo = SITE_REPO;
  return s;
}
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

// ---------- İnternette yayınlama ----------
// Müşteri sayfası ve güncel katalog GitHub'daki "main" dalına gönderilir;
// GitHub Pages bu dalı https://<kullanıcı>.github.io/<depo>/ adresinde yayınlar.

const PUBLISH_FILES = ['index.html', 'style.css', 'vendor/zxing-reader.js', 'vendor/zxing-LICENSE.txt'];
// İkili (metin olmayan) dosyalar: GitHub'a ayrı "blob" olarak yüklenir, zaten oradaysa tekrar gönderilmez.
const PUBLISH_BINARY = ['vendor/zxing_reader.wasm'];
const gitBlobSha = (buf) => crypto.createHash('sha1').update(`blob ${buf.length}\0`).update(buf).digest('hex');
const AUTO_PUBLISH_DELAY = 15000;
let publishTimer = null;
let publishing = null;

// mode 'local': dükkân bilgisayarından açılan sayfa (sipariş doğrudan bu programa gelir);
// mode 'site': internetteki sayfa (sipariş, sipariş deposuna şifreli bırakılır; ayarlı değilse sipariş kapalı).
function catalogData(all = loadProducts(), settings = loadSettings(), mode = 'local') {
  const { shopName, slogan, phone, address } = settings;
  const products = all
    .filter((p) => p.active)
    .map(({ id, name, barcode, category, unit, price, updatedAt }) => ({ id, name, barcode, category, unit, price, updatedAt }));
  let order = null;
  if (settings.ordersEnabled !== false) {
    if (mode === 'local') order = { local: true };
    else if (ordersConfigured(settings)) order = { repo: cleanRepo(settings.orderRepo), k: hideToken(settings.orderToken), pub: orderKeys().pub };
  }
  return { settings: { shopName, slogan, phone, address }, categories: categoryList(all, settings), products, order };
}

function siteFiles() {
  const files = PUBLISH_FILES.map((f) => ({ path: f, content: fs.readFileSync(path.join(PUBLIC_DIR, f), 'utf8') }));
  for (const f of PUBLISH_BINARY) {
    const data = fs.readFileSync(path.join(PUBLIC_DIR, f));
    files.push({ path: f, data, sha: gitBlobSha(data) });
  }
  files.push({ path: 'catalog.json', content: JSON.stringify(catalogData(loadProducts(), loadSettings(), 'site')) });
  files.push({ path: 'siparis-durum.json', content: JSON.stringify(orderStatusMap()) });
  files.push({ path: '.nojekyll', content: '' });
  files.push({ path: '_headers', content: '/\n  Cache-Control: no-cache\n/index.html\n  Cache-Control: no-cache\n/catalog.json\n  Cache-Control: no-cache\n/siparis-durum.json\n  Cache-Control: no-cache\n' });
  return files;
}

const hashFiles = (files) => crypto.createHash('sha256').update(JSON.stringify(files.map((f) => [f.path, f.sha || f.content]))).digest('hex');

function publishConfigured(s = loadSettings()) {
  return Boolean(s.githubRepo && s.githubToken);
}

function publishStatus() {
  const s = loadSettings();
  const configured = publishConfigured(s);
  return {
    configured,
    autoPublish: s.autoPublish,
    siteUrl: s.siteUrl,
    dirty: configured && hashFiles(siteFiles()) !== s.publishedHash,
    publishing: Boolean(publishing),
    lastPublishedAt: s.lastPublishedAt || null,
    lastPublishError: s.lastPublishError || null,
  };
}

async function github(method, apiPath, token, body) {
  const res = await fetch((process.env.GITHUB_API_URL || 'https://api.github.com') + apiPath, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
      'User-Agent': 'dilek-ofis-katalog',
      ...(body ? { 'Content-Type': 'application/json' } : {}),
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  return { status: res.status, data: await res.json().catch(() => ({})) };
}

async function doPublish() {
  const s = loadSettings();
  if (!publishConfigured(s)) throw new Error('Önce Ayarlar > İnternette Yayınlama bölümünü doldurun.');
  const repo = cleanRepo(s.githubRepo);
  const branch = (s.githubBranch || 'main').trim();
  const token = s.githubToken;

  const repoInfo = await github('GET', `/repos/${repo}`, token);
  if (repoInfo.status === 401) throw new Error('GitHub anahtarı geçersiz veya süresi dolmuş. Yeni anahtar oluşturup Ayarlar\'a yapıştırın.');
  if (repoInfo.status === 404) throw new Error(`"${repo}" deposu bulunamadı ya da anahtarın bu depoya erişimi yok.`);
  if (repoInfo.status >= 400) throw new Error(`GitHub hatası (${repoInfo.status}): ${repoInfo.data.message || ''}`);

  const files = siteFiles();
  const ref = await github('GET', `/repos/${repo}/git/ref/heads/${encodeURIComponent(branch)}`, token);
  const parent = ref.status === 200 ? ref.data.object.sha : null;

  for (const f of files.filter((x) => x.data)) {
    const have = await github('GET', `/repos/${repo}/git/blobs/${f.sha}`, token);
    if (have.status === 200) continue;
    const blob = await github('POST', `/repos/${repo}/git/blobs`, token, { content: f.data.toString('base64'), encoding: 'base64' });
    if (blob.status === 403 || blob.status === 404) throw new Error('GitHub anahtarının "Contents: Read and write" izni yok.');
    if (blob.status >= 400) throw new Error(`GitHub hatası (${blob.status}): ${blob.data.message || ''}`);
  }
  const tree = await github('POST', `/repos/${repo}/git/trees`, token, {
    tree: files.map((f) => (f.data
      ? { path: f.path, mode: '100644', type: 'blob', sha: f.sha }
      : { path: f.path, mode: '100644', type: 'blob', content: f.content })),
  });
  if (tree.status === 403 || tree.status === 404) throw new Error('GitHub anahtarının "Contents: Read and write" izni yok.');
  if (tree.status >= 400) throw new Error(`GitHub hatası (${tree.status}): ${tree.data.message || ''}`);

  const stamp = new Date().toLocaleString('tr-TR', { timeZone: 'Europe/Istanbul' });
  const commit = await github('POST', `/repos/${repo}/git/commits`, token, {
    message: `Fiyat listesi güncellendi (${stamp})`,
    tree: tree.data.sha,
    parents: parent ? [parent] : [],
  });
  if (commit.status >= 400) throw new Error(`GitHub hatası (${commit.status}): ${commit.data.message || ''}`);

  const update = parent
    ? await github('PATCH', `/repos/${repo}/git/refs/heads/${encodeURIComponent(branch)}`, token, { sha: commit.data.sha, force: true })
    : await github('POST', `/repos/${repo}/git/refs`, token, { ref: `refs/heads/${branch}`, sha: commit.data.sha });
  if (update.status >= 400) throw new Error(`GitHub hatası (${update.status}): ${update.data.message || ''}`);

  return hashFiles(files);
}

// Aynı anda tek yayın; hatayı ayarlara yazar ki panel gösterebilsin.
function publishSite() {
  clearTimeout(publishTimer);
  publishing ||= doPublish()
    .then((hash) => {
      const s = loadSettings();
      Object.assign(s, { publishedHash: hash, lastPublishedAt: new Date().toISOString(), lastPublishError: null });
      saveSettings(s);
      console.log('  İnternete yayınlandı: ' + new Date().toLocaleTimeString('tr-TR'));
    })
    .catch((e) => {
      const s = loadSettings();
      s.lastPublishError = e.message;
      saveSettings(s);
      console.log('  Yayınlama hatası: ' + e.message);
      throw e;
    })
    .finally(() => { publishing = null; });
  return publishing;
}

// Değişiklikten 15 sn sonra otomatik yayınla (art arda düzenlemeler tek seferde gider).
function schedulePublish() {
  const s = loadSettings();
  if (!s.autoPublish || !publishConfigured(s)) return;
  clearTimeout(publishTimer);
  publishTimer = setTimeout(() => {
    if (publishStatus().dirty) publishSite().catch(() => {});
  }, AUTO_PUBLISH_DELAY);
}

// Tarayıcıya giden ayarlarda GitHub anahtarları ve sipariş şifre anahtarı yer almaz.
function publicSettings(s = loadSettings()) {
  const { githubToken, publishedHash, orderToken, orderKeyPrivate, orderKeyPublic, ...rest } = s;
  return { ...rest, githubTokenSet: Boolean(githubToken), orderTokenSet: Boolean(orderToken) };
}

const cleanRepo = (r) => String(r || '').trim().replace(/^https?:\/\/github\.com\//, '').replace(/\.git$/, '').replace(/\/+$/, '');

// ---------- Siparişler ----------
// Müşteri sepetini onaylayınca sipariş iki yoldan gelir:
//  • Dükkân Wi-Fi'ından açılan sayfa: doğrudan bu programa (POST /api/orders).
//  • İnternetteki sayfa: sipariş, dükkânın şifre anahtarıyla (RSA + AES) şifrelenip gizli sipariş deposuna "issue"
//    olarak bırakılır; bu program 20 saniyede bir depoya bakar, siparişi çözüp kaydeder ve issue'yu kapatır.
//    Bilgisayar kapalıyken gelen siparişler depoda bekler, kaybolmaz. Sayfadaki anahtar yalnız bu depoya
//    sipariş bırakabilir; siparişin içeriği şifreli olduğu için başkası okuyamaz.
// Siparişin durumu (alındı, hazırlanıyor, hazır…) siparis-durum.json ile müşteri sayfasına yayınlanır
// (yalnız sipariş numarası ve durum; ad, telefon gibi bilgiler yayınlanmaz).

const ORDER_STATUSES = ['yeni', 'hazirlaniyor', 'hazir', 'teslim', 'iptal'];
const ORDER_MARK = 'DILEKOFIS-SIPARIS v1';
const ORDER_ID = /^[A-HJ-NP-Z2-9]{8}$/;

const loadOrders = () => readJson(ORDERS_FILE, []);
const saveOrders = (list) => writeJson(ORDERS_FILE, list);

function ordersConfigured(s = loadSettings()) {
  return Boolean(s.orderToken && cleanRepo(s.orderRepo));
}

// Anahtar sayfada açık metin olarak durmaz (GitHub'ın taramasına takılıp iptal edilmesin diye ters çevrilip kodlanır).
const hideToken = (t) => Buffer.from(String(t).split('').reverse().join(''), 'utf8').toString('base64');

// Sipariş şifreleme anahtar çifti ilk gerektiğinde oluşturulur; gizli anahtar yalnız bu bilgisayarda kalır.
function orderKeys() {
  const s = loadSettings();
  if (!s.orderKeyPrivate || !s.orderKeyPublic) {
    const { publicKey, privateKey } = crypto.generateKeyPairSync('rsa', { modulusLength: 2048 });
    s.orderKeyPublic = publicKey.export({ type: 'spki', format: 'der' }).toString('base64');
    s.orderKeyPrivate = privateKey.export({ type: 'pkcs8', format: 'pem' });
    saveSettings(s);
  }
  return { pub: s.orderKeyPublic, priv: s.orderKeyPrivate };
}

// Tarayıcının WebCrypto ile ürettiği zarfı açar: k = RSA-OAEP(SHA-256) ile şifreli AES anahtarı,
// iv = 12 bayt, d = AES-256-GCM şifreli metin + 16 bayt doğrulama etiketi.
function decryptOrder(env) {
  const key = crypto.privateDecrypt(
    { key: orderKeys().priv, padding: crypto.constants.RSA_PKCS1_OAEP_PADDING, oaepHash: 'sha256' },
    Buffer.from(env.k, 'base64')
  );
  const data = Buffer.from(env.d, 'base64');
  const decipher = crypto.createDecipheriv('aes-256-gcm', key, Buffer.from(env.iv, 'base64'));
  decipher.setAuthTag(data.subarray(data.length - 16));
  return JSON.parse(Buffer.concat([decipher.update(data.subarray(0, data.length - 16)), decipher.final()]).toString('utf8'));
}

const clip = (v, n) => String(v ?? '').replace(/[\u0000-\u001f\u007f]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, n);

// Gelen siparişi doğrular ve temizler; ürünler güncel listeyle eşleştirilir (ad ve fiyat buradan alınır,
// müşterinin sepette gördüğü fiyat da saklanır ki arada fiyat değiştiyse görülsün).
function cleanOrder(input, source) {
  if (!input || typeof input !== 'object') throw new Error('Geçersiz sipariş.');
  const id = String(input.id || '');
  if (!ORDER_ID.test(id)) throw new Error('Geçersiz sipariş numarası.');
  const customer = clip(input.customer, 60);
  if (customer.length < 2) throw new Error('Ad soyad yazılmalı.');
  const phone = clip(input.phone, 20).replace(/[^\d+ ()-]/g, '');
  if (!Array.isArray(input.items) || !input.items.length) throw new Error('Sepet boş.');
  if (input.items.length > 200) throw new Error('Sepette en fazla 200 çeşit ürün olabilir.');
  const products = loadProducts();
  const merged = new Map();
  for (const it of input.items) {
    const qty = Math.floor(Number(it && it.qty));
    if (!Number.isFinite(qty) || qty < 1 || qty > 9999) throw new Error('Geçersiz ürün adedi.');
    const pid = String((it && it.id) || '').slice(0, 64);
    const p = products.find((x) => x.id === pid);
    const seen = Number(it.price);
    const row = merged.get(pid) || {
      id: pid,
      name: p ? p.name : clip(it.name, 120) || 'Bilinmeyen ürün',
      unit: p ? p.unit : clip(it.unit, 20) || 'Adet',
      price: p ? p.price : (Number.isFinite(seen) && seen >= 0 ? Math.round(seen * 100) / 100 : 0),
      seenPrice: Number.isFinite(seen) && seen >= 0 ? Math.round(seen * 100) / 100 : null,
      qty: 0,
      missing: !p,
    };
    row.qty = Math.min(9999, row.qty + qty);
    merged.set(pid, row);
  }
  const items = [...merged.values()];
  const now = new Date().toISOString();
  return {
    id,
    name: clip(input.name, 60),
    customer,
    phone,
    note: clip(input.note, 500),
    items,
    total: Math.round(items.reduce((t, x) => t + x.price * x.qty, 0) * 100) / 100,
    seenTotal: Math.round(items.reduce((t, x) => t + (x.seenPrice ?? x.price) * x.qty, 0) * 100) / 100,
    createdAt: !Number.isNaN(Date.parse(input.createdAt)) && Math.abs(Date.parse(input.createdAt) - Date.now()) < 7 * 864e5
      ? new Date(input.createdAt).toISOString() : now,
    receivedAt: now,
    status: 'yeni',
    statusAt: now,
    source,
  };
}

// Aynı sipariş (aynı numara) iki kez gelirse ikincisi yok sayılır (müşteri "tekrar dene"ye basmış olabilir).
function addOrder(order) {
  const orders = loadOrders();
  if (orders.some((o) => o.id === order.id)) return false;
  orders.push(order);
  saveOrders(orders);
  console.log(`  Yeni sipariş: ${order.id} · ${order.customer}${order.name ? ' · ' + order.name : ''} · ${order.items.length} çeşit`);
  schedulePublish(); // müşteri "Siparişiniz alındı" görsün
  return true;
}

// Müşteri sayfasına yayınlanan durum listesi: son 60 günün siparişleri, numara → [durum, zaman].
function orderStatusMap(orders = loadOrders()) {
  const since = Date.now() - 60 * 864e5;
  const out = {};
  for (const o of orders) if (Date.parse(o.createdAt) > since) out[o.id] = [o.status, o.statusAt];
  return out;
}

// Sipariş deposundaki açık issue'lar okunur, çözülür, kaydedilir ve issue kapatılır.
const orderPoll = { lastAt: null, lastError: null, running: false };
async function fetchOrders() {
  const s = loadSettings();
  if (!ordersConfigured(s) || s.ordersEnabled === false || orderPoll.running) return { added: 0 };
  orderPoll.running = true;
  let added = 0;
  try {
    const repo = cleanRepo(s.orderRepo);
    const list = await github('GET', `/repos/${repo}/issues?state=open&per_page=100&sort=created&direction=asc`, s.orderToken);
    if (list.status === 401) throw new Error('Sipariş anahtarı geçersiz veya süresi dolmuş.');
    if (list.status === 404) throw new Error(`"${repo}" sipariş deposu bulunamadı ya da anahtarın bu depoya erişimi yok.`);
    if (list.status === 403) throw new Error('Sipariş anahtarının "Issues: Read and write" izni yok.');
    if (list.status >= 400) throw new Error(`GitHub hatası (${list.status}): ${list.data.message || ''}`);
    for (const issue of Array.isArray(list.data) ? list.data : []) {
      if (issue.pull_request) continue;
      let reason = 'completed';
      try {
        const body = String(issue.body || '');
        if (!body.startsWith(ORDER_MARK)) throw new Error('sipariş değil');
        const order = cleanOrder(decryptOrder(JSON.parse(Buffer.from(body.slice(ORDER_MARK.length).trim(), 'base64').toString('utf8'))), 'internet');
        if (addOrder(order)) added++;
      } catch (e) {
        reason = 'not_planned';
        logError(new Error(`Sipariş #${issue.number} okunamadı: ${e.message}`));
      }
      const closed = await github('PATCH', `/repos/${repo}/issues/${issue.number}`, s.orderToken, { state: 'closed', state_reason: reason });
      if (closed.status >= 400) throw new Error(`Sipariş alındı ama depoda kapatılamadı (${closed.status}).`);
    }
    orderPoll.lastError = null;
  } catch (e) {
    orderPoll.lastError = e.message === 'fetch failed' ? 'İnternete bağlanılamadı.' : e.message;
  } finally {
    orderPoll.lastAt = new Date().toISOString();
    orderPoll.running = false;
  }
  return { added };
}
setInterval(() => fetchOrders().catch(() => {}), 20000).unref();

// Yerel ağdan gelen siparişler için basit sınır: aynı adresten dakikada en fazla 10 sipariş.
const orderHits = new Map();
function orderRateOk(ip) {
  const now = Date.now();
  const hits = (orderHits.get(ip) || []).filter((t) => now - t < 60000);
  hits.push(now);
  orderHits.set(ip, hits);
  return hits.length <= 10;
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
  '.wasm': 'application/wasm',
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
    return send(res, 200, catalogData(loadProducts(), loadSettings(), 'local'));
  }

  // Dükkân Wi-Fi'ından açılan sayfadan gelen sipariş
  if (pathname === '/api/orders' && method === 'POST') {
    if (loadSettings().ordersEnabled === false) return send(res, 403, { error: 'Sipariş alma şu an kapalı.' });
    if (!orderRateOk(req.socket.remoteAddress)) return send(res, 429, { error: 'Çok fazla deneme. Biraz bekleyin.' });
    let order;
    try {
      order = cleanOrder(JSON.parse(await readBody(req, 200 * 1024)), 'yerel');
    } catch (e) {
      return send(res, 400, { error: e.message });
    }
    addOrder(order);
    return send(res, 200, { ok: true, id: order.id });
  }

  if (pathname === '/api/order-status' && method === 'GET') return send(res, 200, orderStatusMap());

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
    schedulePublish();
    return send(res, 200, categoryList(loadProducts(), s));
  }

  // ---- Siparişler ----
  if (route === 'orders' && method === 'GET') {
    return send(res, 200, { orders: loadOrders(), poll: { ...orderPoll, configured: ordersConfigured() } });
  }

  if (route === 'orders/check' && method === 'POST') {
    const r = await fetchOrders();
    return send(res, 200, { ...r, poll: { ...orderPoll, configured: ordersConfigured() } });
  }

  const orderMatch = route.match(/^orders\/([A-Z0-9]+)$/);
  if (orderMatch && method === 'PUT') {
    const { status } = await readJsonBody(req);
    if (!ORDER_STATUSES.includes(status)) return send(res, 400, { error: 'Geçersiz durum.' });
    const orders = loadOrders();
    const o = orders.find((x) => x.id === orderMatch[1]);
    if (!o) return send(res, 404, { error: 'Sipariş bulunamadı.' });
    o.status = status;
    o.statusAt = new Date().toISOString();
    saveOrders(orders);
    schedulePublish(); // müşteri durumu görsün
    return send(res, 200, o);
  }

  if (orderMatch && method === 'DELETE') {
    const orders = loadOrders();
    const next = orders.filter((x) => x.id !== orderMatch[1]);
    if (next.length === orders.length) return send(res, 404, { error: 'Sipariş bulunamadı.' });
    saveOrders(next);
    schedulePublish();
    return send(res, 200, { ok: true });
  }

  if (route === 'settings' && method === 'GET') return send(res, 200, publicSettings());

  if (route === 'settings' && method === 'PUT') {
    const input = await readJsonBody(req);
    const s = loadSettings();
    for (const k of ['shopName', 'slogan', 'phone', 'address', 'publicUrl', 'githubRepo', 'githubBranch', 'siteUrl', 'orderRepo']) {
      if (input[k] !== undefined) s[k] = String(input[k]).trim();
    }
    if (input.githubToken) s.githubToken = String(input.githubToken).trim();
    if (input.orderToken) s.orderToken = String(input.orderToken).trim();
    if (input.autoPublish !== undefined) s.autoPublish = Boolean(input.autoPublish);
    if (input.ordersEnabled !== undefined) s.ordersEnabled = Boolean(input.ordersEnabled);
    saveSettings(s);
    schedulePublish();
    if (input.orderToken || input.orderRepo !== undefined) fetchOrders().catch(() => {});
    return send(res, 200, publicSettings(s));
  }

  if (route === 'publish' && method === 'GET') return send(res, 200, publishStatus());

  if (route === 'publish' && method === 'POST') {
    try {
      await publishSite();
      return send(res, 200, publishStatus());
    } catch (e) {
      return send(res, 502, { error: e.message, ...publishStatus() });
    }
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

// Tarayıcıyı ancak sunucu gerçekten hazır olunca aç (Windows).
function openBrowser(url) {
  if (process.platform !== 'win32' || process.env.NO_BROWSER) return;
  exec(`start "" "${url}"`, { windowsHide: true });
}

function onListening() {
  console.log('\n  Dilek Ofis Fiyat Kataloğu çalışıyor\n');
  console.log(`  Yönetim paneli : http://localhost:${PORT}/admin`);
  const [best, ...others] = lanAddresses();
  const { publicUrl } = loadSettings();
  console.log(`  QR kod adresi  : ${publicUrl || '(QR Kod sekmesinde sabitlenmemiş)'}`);
  if (best) console.log(`  Dükkân Wi-Fi'ı : ${best}`);
  for (const u of others) console.log(`  Diğer adres    : ${u}`);
  const pub = publishStatus();
  if (pub.configured) {
    console.log(`\n  İnternet adresi: ${pub.siteUrl || '(Ayarlar\'da site adresi girilmemiş)'}`);
    if (pub.dirty && pub.autoPublish) schedulePublish();
  }
  if (publicUrl && /^http:\/\/(\d+\.){3}\d+/.test(publicUrl) && !lanAddresses().includes(publicUrl)) {
    console.log(`\n  !!! DİKKAT: Basılı QR kod ${publicUrl} adresine yönleniyor,`);
    console.log('  !!! ama bu bilgisayarın adresi değişmiş. Müşteriler sayfayı açamaz.');
    console.log('  !!! Çözüm: Yönetim paneli > QR Kod sekmesindeki talimatlara bakın.');
  }
  if (ordersConfigured()) {
    console.log('  İnternetten sipariş alma açık (20 sn\'de bir kontrol edilir).');
    fetchOrders().catch(() => {});
  }
  if (ADMIN_PASSWORD === 'dilek123') console.log('\n  Uyarı: Varsayılan şifre (dilek123) kullanılıyor. ADMIN_PASSWORD ile değiştirin.');
  console.log('\n  Kapatmak için bu pencerede Ctrl+C\n');
  openBrowser(`http://localhost:${PORT}/admin`);
}

// Port doluysa: zaten çalışan bir Dilek Ofis varsa onu aç, yoksa sıradaki boş portu dene.
server.on('error', async (e) => {
  if (e.code !== 'EADDRINUSE') { logError(e); process.exit(1); }
  try {
    const r = await fetch(`http://localhost:${PORT}/api/catalog`, { signal: AbortSignal.timeout(1500) });
    if (r.ok && (await r.json()).products) {
      // Arka planda açık kalmış eski kopya (belki eski sürüm) kapatılır, yerine bu sürüm çalışır.
      if (!retriedAfterStop && stopOldCopy(PORT)) {
        retriedAfterStop = true;
        console.log('  Açık kalan eski program kapatıldı, yenisi başlatılıyor...');
        return setTimeout(() => server.listen(PORT), 1500);
      }
      console.log(`\n  Program zaten açık (http://localhost:${PORT}/admin). Tarayıcıda açılıyor...`);
      console.log('  Bu pencereyi kapatabilirsiniz.\n');
      openBrowser(`http://localhost:${PORT}/admin`);
      return setTimeout(() => process.exit(0), 1500);
    }
  } catch {}
  if (PORT >= START_PORT + 10) {
    logError(new Error(`${START_PORT}-${PORT} arasındaki portların hepsi başka programlarca kullanılıyor.`));
    process.exit(1);
  }
  console.log(`  ${PORT} numaralı port başka bir program tarafından kullanılıyor, ${PORT + 1} deneniyor...`);
  PORT += 1;
  server.listen(PORT);
});

// Windows'ta portu tutan eski Node sürecini bulup kapatır (yalnız bizim programımız olduğu anlaşıldıktan sonra çağrılır).
let retriedAfterStop = false;
function stopOldCopy(port) {
  if (process.platform !== 'win32') return false;
  try {
    const out = require('child_process').execSync('netstat -ano -p tcp', { encoding: 'utf8' });
    const pids = new Set();
    for (const line of out.split(/\r?\n/)) {
      const m = line.trim().split(/\s+/);
      if (m.length >= 5 && /^tcp/i.test(m[0]) && /:0$/.test(m[2]) && m[1].endsWith(':' + port) && +m[4] !== process.pid) pids.add(m[4]);
    }
    for (const pid of pids) require('child_process').execSync(`taskkill /PID ${pid} /F`, { stdio: 'ignore' });
    return pids.size > 0;
  } catch { return false; }
}

server.on('listening', onListening);
server.listen(PORT);
