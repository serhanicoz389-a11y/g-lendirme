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

// ---------- Excel ile alışveriş (ek paket gerektirmez) ----------
// Dışa aktarma gerçek Excel dosyasıdır (.xlsx): Kod ve Barkod sütunları "Metin" biçimlidir, böylece Excel uzun
// barkodları 8,69724E+12 gibi bozmaz; yeni satıra yazılan barkod da metin kalır. Son sütundaki "No" ürünün
// kimliğidir: ürünün adı ya da barkodu değişse bile aynı ürün olarak tanınır.
// İçe aktarma .xlsx ve .csv (UTF-8 ya da Windows Türkçe) okur. Önce bir önizleme gösterilir: kaç ürün değişecek,
// eklenecek, dosyada olmadığı için silinecek ve sorunlu satırlar; onaylanınca uygulanır.

const zlib = require('zlib');

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();
function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function zipFiles(files) {
  const parts = [], central = [];
  let offset = 0;
  for (const f of files) {
    const name = Buffer.from(f.name, 'utf8'), data = Buffer.from(f.data, 'utf8'), packed = zlib.deflateRawSync(data);
    const crc = crc32(data);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0); local.writeUInt16LE(20, 4); local.writeUInt16LE(0x0800, 6); local.writeUInt16LE(8, 8);
    local.writeUInt32LE(crc, 14); local.writeUInt32LE(packed.length, 18); local.writeUInt32LE(data.length, 22); local.writeUInt16LE(name.length, 26);
    const cen = Buffer.alloc(46);
    cen.writeUInt32LE(0x02014b50, 0); cen.writeUInt16LE(20, 4); cen.writeUInt16LE(20, 6); cen.writeUInt16LE(0x0800, 8); cen.writeUInt16LE(8, 10);
    cen.writeUInt32LE(crc, 16); cen.writeUInt32LE(packed.length, 20); cen.writeUInt32LE(data.length, 24); cen.writeUInt16LE(name.length, 28);
    cen.writeUInt32LE(offset, 42);
    parts.push(local, name, packed);
    central.push(cen, name);
    offset += 30 + name.length + packed.length;
  }
  const size = central.reduce((s, b) => s + b.length, 0), end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(files.length, 8); end.writeUInt16LE(files.length, 10);
  end.writeUInt32LE(size, 12); end.writeUInt32LE(offset, 16);
  return Buffer.concat([...parts, ...central, end]);
}

function unzip(buf) {
  let e = buf.length - 22;
  while (e >= 0 && buf.readUInt32LE(e) !== 0x06054b50) e--;
  if (e < 0) throw new Error('Dosya açılamadı (bozuk ya da Excel dosyası değil).');
  const count = buf.readUInt16LE(e + 10);
  let p = buf.readUInt32LE(e + 16);
  const out = new Map();
  for (let i = 0; i < count; i++) {
    if (buf.readUInt32LE(p) !== 0x02014b50) break;
    const method = buf.readUInt16LE(p + 10), csize = buf.readUInt32LE(p + 20), nlen = buf.readUInt16LE(p + 28);
    const xlen = buf.readUInt16LE(p + 30), clen = buf.readUInt16LE(p + 32), lo = buf.readUInt32LE(p + 42);
    const name = buf.subarray(p + 46, p + 46 + nlen).toString('utf8');
    const start = lo + 30 + buf.readUInt16LE(lo + 26) + buf.readUInt16LE(lo + 28), raw = buf.subarray(start, start + csize);
    out.set(name.replace(/^\/+/, ''), method === 8 ? zlib.inflateRawSync(raw) : raw);
    p += 46 + nlen + xlen + clen;
  }
  return out;
}

const xmlEsc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
  .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '');
const xmlText = (s) => s.replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (m, c) =>
  c[0] === '#' ? String.fromCodePoint(c[1] === 'x' || c[1] === 'X' ? parseInt(c.slice(2), 16) : Number(c.slice(1)))
    : { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }[c.toLowerCase()]);
const colName = (i) => (i >= 26 ? colName(Math.floor(i / 26) - 1) : '') + String.fromCharCode(65 + (i % 26));
const colIndex = (ref) => [...ref.replace(/\d+/g, '')].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;

// cols: [{ title, width, kind: 'text' | 'money' | 'general' }]
function makeXlsx(cols, rows) {
  const style = { text: 2, money: 3, general: 0 };
  const cell = (v, ci, ri) => {
    const r = colName(ci) + ri, s = style[cols[ci].kind];
    if (cols[ci].kind === 'money' && typeof v === 'number') return `<c r="${r}" s="${s}"><v>${v}</v></c>`;
    return v === '' || v == null ? `<c r="${r}" s="${s}"/>` : `<c r="${r}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${xmlEsc(v)}</t></is></c>`;
  };
  const last = colName(cols.length - 1);
  const sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
    '<cols>' + cols.map((c, i) => `<col min="${i + 1}" max="${i + 1}" width="${c.width}" style="${style[c.kind]}" customWidth="1"/>`).join('') + '</cols>' +
    '<sheetData><row r="1">' + cols.map((c, i) => `<c r="${colName(i)}1" s="1" t="inlineStr"><is><t>${xmlEsc(c.title)}</t></is></c>`).join('') + '</row>' +
    rows.map((row, ri) => `<row r="${ri + 2}">` + row.map((v, ci) => cell(v, ci, ri + 2)).join('') + '</row>').join('') +
    `</sheetData><autoFilter ref="A1:${last}${rows.length + 1}"/></worksheet>`;
  const styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00"/></numFmts>' +
    '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
    '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFCE4EC"/></patternFill></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="49" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyNumberFormat="1"/>' +
    '<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>' +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
  return zipFiles([
    { name: '[Content_Types].xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>' },
    { name: '_rels/.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>' },
    { name: 'xl/workbook.xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Ürünler" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">\'Ürünler\'!$A$1:$' + last + '$' + (rows.length + 1) + '</definedName></definedNames></workbook>' },
    { name: 'xl/_rels/workbook.xml.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>' },
    { name: 'xl/worksheets/sheet1.xml', data: sheet },
    { name: 'xl/styles.xml', data: styles },
  ]);
}

// İlk sayfanın hücrelerini metin olarak okur (sayılar tam haneleriyle: 8697236130339).
function readXlsx(buf) {
  const files = unzip(buf), get = (n) => files.get(n) && files.get(n).toString('utf8');
  const wb = get('xl/workbook.xml');
  if (!wb) throw new Error('Bu bir Excel (.xlsx) dosyası değil.');
  let sheetPath = 'xl/worksheets/sheet1.xml';
  const first = /<sheet\b[^>]*\br:id="([^"]+)"/.exec(wb), rels = get('xl/_rels/workbook.xml.rels') || '';
  if (first) {
    const rel = new RegExp(`<Relationship\\b[^>]*\\bId="${first[1]}"[^>]*>`).exec(rels), target = rel && /Target="([^"]+)"/.exec(rel[0]);
    if (target) sheetPath = target[1].startsWith('/') ? target[1].slice(1) : 'xl/' + target[1].replace(/^\.\//, '');
  }
  const shared = [];
  for (const si of (get('xl/sharedStrings.xml') || '').match(/<si>[\s\S]*?<\/si>/g) || []) {
    shared.push(xmlText((si.match(/<t\b[^>]*>[\s\S]*?<\/t>/g) || []).map((t) => t.replace(/<[^>]+>/g, '')).join('')));
  }
  const xml = get(sheetPath);
  if (!xml) throw new Error('Excel dosyasında sayfa bulunamadı.');
  const rows = [];
  for (const rm of xml.matchAll(/<row\b([^>]*)>([\s\S]*?)<\/row>/g)) {
    const rn = /\br="(\d+)"/.exec(rm[1]), row = [];
    for (const cm of rm[2].matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const ref = /\br="([A-Z]+\d+)"/.exec(cm[1]), t = (/\bt="(\w+)"/.exec(cm[1]) || [])[1], body = cm[2] || '';
      const v = /<v>([\s\S]*?)<\/v>/.exec(body);
      let val = '';
      if (t === 's') val = v ? shared[Number(v[1])] ?? '' : '';
      else if (t === 'inlineStr') val = xmlText((body.match(/<t\b[^>]*>[\s\S]*?<\/t>/g) || []).map((x) => x.replace(/<[^>]+>/g, '')).join(''));
      else if (t === 'str') val = v ? xmlText(v[1]) : '';
      else if (t === 'b') val = v && v[1] === '1' ? 'DOĞRU' : 'YANLIŞ';
      else if (t === 'e') val = '';
      else if (v) {
        const n = Number(v[1]);
        val = Number.isFinite(n) && Number.isInteger(n) && Math.abs(n) < 1e21 ? n.toFixed(0) : Number.isFinite(n) ? String(Math.round(n * 1e6) / 1e6) : v[1];
      }
      row[ref ? colIndex(ref[1]) : row.length] = val;
    }
    rows[rn ? Number(rn[1]) - 1 : rows.length] = Array.from(row, (x) => x ?? '');
  }
  return Array.from(rows, (r) => r || []);
}

// CSV: UTF-8 değilse Windows Türkçe (1254) kabul edilir (Excel'in "CSV (virgülle ayrılmış)" kaydı böyledir).
const CP1254 = { 0xd0: 'Ğ', 0xdd: 'İ', 0xde: 'Ş', 0xf0: 'ğ', 0xfd: 'ı', 0xfe: 'ş', 0x80: '€', 0x8a: 'Š', 0x9a: 'š', 0x8c: 'Œ', 0x9c: 'œ', 0x9f: 'Ÿ' };
function decodeText(buf) {
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(buf);
  } catch {
    let s = '';
    for (const b of buf) s += CP1254[b] || String.fromCharCode(b);
    return s;
  }
}

function parseCsv(text) {
  text = text.replace(/^﻿/, '');
  const firstLine = text.split(/\r?\n/, 1)[0];
  const sep = firstLine.split(';').length >= firstLine.split(',').length ? ';' : (firstLine.includes('\t') ? '\t' : ',');
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
  return rows;
}

function readSheet(buf) {
  if (buf.length > 4 && buf.readUInt32LE(0) === 0x04034b50) return readXlsx(buf);
  if (buf.length > 8 && buf.readUInt32LE(0) === 0xe011cfd0) {
    throw new Error('Bu eski Excel biçimi (.xls). Excel\'de "Farklı Kaydet" → "Excel Çalışma Kitabı (.xlsx)" seçip o dosyayı yükleyin.');
  }
  return parseCsv(decodeText(buf));
}

const EXCEL_COLS = [
  { title: 'Kod', width: 10, kind: 'text' },
  { title: 'Barkod', width: 17, kind: 'text' },
  { title: 'Ürün Adı', width: 40, kind: 'general' },
  { title: 'Kategori', width: 20, kind: 'general' },
  { title: 'Birim', width: 9, kind: 'general' },
  { title: 'Fiyat', width: 11, kind: 'money' },
  { title: 'Satışta', width: 9, kind: 'general' },
  { title: 'No (değiştirmeyin)', width: 38, kind: 'text' },
];

function exportXlsx(products) {
  const cats = categoryList(products);
  const list = [...products].sort((a, b) => cats.indexOf(a.category) - cats.indexOf(b.category) || a.name.localeCompare(b.name, 'tr'));
  return makeXlsx(EXCEL_COLS, list.map((p) => [p.code, p.barcode, p.name, p.category, p.unit, p.price, p.active ? 'Evet' : 'Hayır', p.id]));
}

const trLower = (s) => String(s || '').toLocaleLowerCase('tr').trim();

// Fiyat: "1.250,50", "1250,5", "1250.50", "₺ 45", "45 TL" … → sayı; anlaşılamazsa null.
function parsePrice(raw) {
  let t = String(raw ?? '').replace(/[^\d.,-]/g, '');
  if (!t) return null;
  if (t.includes(',') && t.includes('.')) t = t.lastIndexOf(',') > t.lastIndexOf('.') ? t.replace(/\./g, '').replace(',', '.') : t.replace(/,/g, '');
  else if (t.includes(',')) t = t.replace(',', '.');
  else if (/^\d{1,3}(\.\d{3})+$/.test(t)) t = t.replace(/\./g, ''); // 1.250 → 1250
  const n = Number(t);
  return Number.isFinite(n) && n >= 0 ? Math.round(n * 100) / 100 : null;
}

// Dosyayı mevcut ürünlerle karşılaştırıp ne olacağını hesaplar; hiçbir şeyi kaydetmez.
function planImport(rows, products) {
  const problems = [];
  let h = rows.findIndex((r, i) => i < 10 && r.some((c) => /^(ürün|urun)/i.test(String(c).trim())) && r.some((c) => /^fiyat/i.test(String(c).trim())));
  if (h < 0) throw new Error('Dosyada "Ürün Adı" ve "Fiyat" başlıklı sütunlar bulunamadı. Programdan "Excel\'e Aktar" ile aldığınız dosyayı kullanın.');
  const header = rows[h].map((x) => trLower(x));
  const col = (re) => header.findIndex((x) => re.test(x));
  const C = {
    id: col(/^(no|id)\b/), code: col(/^kod/), barcode: col(/^barkod/), name: col(/^(ürün|urun)/), cat: col(/^kategori/),
    unit: col(/^birim/), price: col(/^fiyat/), active: col(/^(satışta|satista|aktif|durum)/),
  };
  const byId = new Map(products.map((p) => [p.id, p]));
  const used = new Set(), seenBarcode = new Map(), out = [];
  for (let i = h + 1; i < rows.length; i++) {
    const r = rows[i] || [], line = i + 1;
    const g = (k) => (C[k] >= 0 ? String(r[C[k]] ?? '').trim() : undefined);
    if (!r.some((c) => String(c).trim())) continue;
    const name = g('name');
    if (!name) { problems.push({ line, msg: 'Ürün adı boş; satır atlandı.' }); continue; }
    const input = { name };
    let barcode = g('barcode');
    if (barcode !== undefined) {
      barcode = barcode.replace(/^'/, '').replace(/\s+/g, '');
      if (/^\d+\.0+$/.test(barcode)) barcode = barcode.replace(/\.0+$/, '');
      if (/^[\d.,]+e\+?\d+$/i.test(barcode)) {
        problems.push({ line, name, msg: `Barkod Excel'de bozulmuş (${barcode}); bu ürünün barkodu değiştirilmedi. Barkod sütununu "Metin" yapıp barkodu yeniden yazın.` });
        barcode = undefined;
      } else if (barcode && seenBarcode.has(barcode)) {
        problems.push({ line, name, msg: `Barkod ${barcode} dosyada ${seenBarcode.get(barcode)}. satırda da var; bu satırın barkodu yok sayıldı.` });
        barcode = undefined;
      }
      if (barcode) seenBarcode.set(barcode, line);
      if (barcode !== undefined) input.barcode = barcode;
    }
    if (g('code') !== undefined) input.code = g('code');
    if (g('cat') !== undefined) input.category = g('cat') || 'Diğer';
    if (g('unit') !== undefined) input.unit = g('unit') || 'Adet';
    const rawPrice = g('price');
    const price = parsePrice(rawPrice);
    if (price === null) problems.push({ line, name, msg: rawPrice ? `Fiyat anlaşılamadı (${rawPrice}); fiyat değiştirilmedi.` : 'Fiyat boş; fiyat değiştirilmedi.' });
    else input.price = price;
    if (g('active') !== undefined && g('active') !== '') input.active = !/^(hayır|hayir|h|0|false|yanlış|yanlis|pasif|yok)$/i.test(g('active'));

    // Eşleştirme: No, barkod, kod, ad (aynı ürün iki satırla eşleşmez)
    const free = (p) => p && !used.has(p.id);
    const id = g('id');
    let match = (id && free(byId.get(id)) && byId.get(id)) ||
      (barcode && products.find((p) => free(p) && p.barcode === barcode)) ||
      (input.code && products.find((p) => free(p) && p.code && p.code === input.code)) ||
      products.find((p) => free(p) && trLower(p.name) === trLower(name));
    if (id && byId.get(id) && used.has(id)) problems.push({ line, name, msg: 'Bu satırdaki No başka bir satırda da var; yeni ürün olarak eklenecek.' });
    if (match) used.add(match.id);
    out.push({ line, match, input });
  }
  // Barkod başka (dosyada olmayan ama silinmeyecek) bir üründe kayıtlıysa çakışma
  return { rows: out, problems, missing: products.filter((p) => !used.has(p.id)) };
}

function applyPlan(plan, products, deleteMissing) {
  const sameish = (a, b) => ['name', 'code', 'barcode', 'category', 'unit', 'price', 'active'].every((k) => a[k] === b[k]);
  let next = deleteMissing ? products.filter((p) => !plan.missing.includes(p)) : [...products];
  const added = [], updated = [], unchanged = [];
  for (const r of plan.rows) {
    if (r.match) {
      const cand = cleanProduct(r.input, r.match);
      if (sameish(cand, r.match)) { unchanged.push(r.match); continue; }
      const changes = ['name', 'code', 'barcode', 'category', 'unit', 'price', 'active'].filter((k) => cand[k] !== r.match[k]);
      next = next.map((p) => (p.id === r.match.id ? cand : p));
      updated.push({ before: r.match, after: cand, changes });
    } else {
      const p = cleanProduct(r.input, { id: crypto.randomUUID() });
      next.push(p);
      added.push(p);
    }
  }
  // Aynı barkod iki üründe kalmasın: dosyadan geleni koru, diğerinin barkodunu sil.
  const problems = [];
  const fromFile = new Set([...added, ...updated.map((u) => u.after)].map((p) => p.id));
  const owner = new Map();
  for (const p of next) {
    if (!p.barcode) continue;
    const o = owner.get(p.barcode);
    if (!o) { owner.set(p.barcode, p); continue; }
    const loser = fromFile.has(p.id) && !fromFile.has(o.id) ? o : p;
    problems.push({ name: loser.name, msg: `Barkod ${p.barcode} "${(loser === p ? o : p).name}" ürününde de var; "${loser.name}" ürününden kaldırıldı.` });
    if (loser === o) owner.set(p.barcode, p);
    next = next.map((x) => (x.id === loser.id ? { ...x, barcode: '' } : x));
  }
  return { next, added, updated, unchanged, deleted: deleteMissing ? plan.missing : [], problems };
}

// Yüklenen dosya önizlemeden sonra "Uygula"ya basılana kadar bellekte bekler (10 dk).
const pendingImports = new Map();

function importPreview(buf) {
  const rows = readSheet(buf);
  const products = loadProducts();
  const plan = planImport(rows, products);
  if (!plan.rows.length) throw new Error('Dosyada ürün satırı bulunamadı.');
  const r = applyPlan(plan, products, true);
  const token = crypto.randomBytes(12).toString('hex');
  for (const [k, v] of pendingImports) if (Date.now() - v.at > 600000) pendingImports.delete(k);
  pendingImports.set(token, { rows, at: Date.now() });
  const FIELD = { name: 'ad', code: 'kod', barcode: 'barkod', category: 'kategori', unit: 'birim', price: 'fiyat', active: 'satışta' };
  const show = (k, v) => (k === 'price' ? String(v).replace('.', ',') + ' ₺' : k === 'active' ? (v ? 'evet' : 'hayır') : v === '' ? '(boş)' : v);
  return {
    token,
    total: plan.rows.length,
    added: r.added.map((p) => p.name),
    updated: r.updated.map((u) => ({ name: u.after.name, changes: u.changes.map((k) => `${FIELD[k]}: ${show(k, u.before[k])} → ${show(k, u.after[k])}`) })),
    unchanged: r.unchanged.length,
    missing: plan.missing.map((p) => p.name),
    problems: [...plan.problems, ...r.problems],
  };
}

function importApply(token, deleteMissing) {
  const pending = pendingImports.get(token);
  if (!pending) throw new Error('Önizlemenin süresi doldu. Dosyayı yeniden seçin.');
  const products = loadProducts();
  const r = applyPlan(planImport(pending.rows, products), products, deleteMissing);
  saveProducts(r.next);
  pendingImports.delete(token);
  return { added: r.added.length, updated: r.updated.length, deleted: r.deleted.length };
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
  return readRaw(req, limit).then((b) => b.toString('utf8'));
}

function readRaw(req, limit = 5 * 1024 * 1024) {
  return new Promise((resolve, reject) => {
    let size = 0;
    const chunks = [];
    req.on('data', (c) => {
      size += c.length;
      if (size > limit) { reject(new Error('Gönderilen veri çok büyük.')); req.destroy(); }
      else chunks.push(c);
    });
    req.on('end', () => resolve(Buffer.concat(chunks)));
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

  if ((route === 'export.xlsx' || route === 'export.csv') && method === 'GET') {
    const date = new Date().toISOString().slice(0, 10);
    return send(res, 200, exportXlsx(loadProducts()), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', {
      'Content-Disposition': `attachment; filename="dilek-ofis-urunler-${date}.xlsx"`,
    });
  }

  // Excel/CSV yükleme: önce önizleme, onaydan sonra uygulama
  if (route === 'import' && method === 'POST') {
    try {
      return send(res, 200, importPreview(await readRaw(req, 20 * 1024 * 1024)));
    } catch (e) {
      return send(res, 400, { error: e.message });
    }
  }

  if (route === 'import/apply' && method === 'POST') {
    const { token, deleteMissing } = await readJsonBody(req);
    try {
      return send(res, 200, importApply(String(token || ''), Boolean(deleteMissing)));
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
