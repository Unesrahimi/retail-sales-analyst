#!/usr/bin/env node
// Retail sales analyzer: reads Excel/CSV sales exports (Persian or English), computes
// sales, customer (RFM), product (ABC), discount, basket and inventory metrics, and writes
// analysis.json + customers.csv (+ optional report.xlsx).
//
// Usage:
//   node analyze.mjs <file-or-folder> [more files...] [options]
// Options:
//   --out <dir>            output folder (default: ./sales-analysis)
//   --stock <file>         optional inventory file (sku/product + stock [+ cost])
//   --customers <file>     optional customer database (name + phone [+ code]) to fill missing phones by name
//   --map <json>           column overrides, e.g. '{"net":"Total Paid","customer":"Buyer"}'
//   --divisor <n>          divide all money by n (10 = Rial -> Toman). Default 1.
//   --currency <label>     label for money in outputs (e.g. Toman, Rial, EUR). Default "unit".
//   --preset shopify|square  known export layouts (column map + revenue basis)
//   --basis net|final|gross|unit  force which amount is revenue (default: auto, by column fill rate)
//   --date-format auto|dmy|mdy    how to read 03/04/2026 (default auto: detect, else dmy)
//   --decimal auto|dot|comma      decimal separator in text numbers (default auto)
//   --excel                also write report.xlsx
//   --lang auto|fa|en      output labels (default auto: fa if headers are Persian, else en)
//   --lead-days <n>        restock lead time for low-stock flag (default 21)
//   --sheet <name>         only read this sheet name from each workbook
import { createRequire } from 'module';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const require = createRequire(import.meta.url);
let XLSX;
try { XLSX = require('xlsx'); }
catch {
  const here = path.dirname(fileURLToPath(import.meta.url));
  console.error(`Missing dependency "xlsx". Run:  cd "${here}" && npm install`);
  process.exit(2);
}

// ---------- args ----------
const argv = process.argv.slice(2);
const opts = { out: 'sales-analysis', divisor: 1, currency: '', excel: false, lang: 'auto', leadDays: 21, map: {}, inputs: [], basis: 'auto', dateFormat: 'auto', decimal: 'auto' };
const PRESETS = {
  shopify: { basis: 'unit', map: { invoice: 'Name', date: 'Created at', customer: 'Billing Name', email: 'Email', phone: 'Billing Phone', product: 'Lineitem name', sku: 'Lineitem sku', qty: 'Lineitem quantity', unitPrice: 'Lineitem price', discount: 'Lineitem discount', returnAmount: 'Refunded Amount', tax: 'Taxes', status: 'Financial Status' } },
  square: { basis: 'net', map: { date: 'Date', product: 'Item', qty: 'Qty', sku: 'SKU', category: 'Category', gross: 'Gross Sales', discount: 'Discounts', net: 'Net Sales', tax: 'Tax', invoice: 'Transaction ID', customerId: 'Customer ID', customer: 'Customer Name' } },
};
for (let i = 0; i < argv.length; i++) {
  const a = argv[i];
  const next = () => argv[++i];
  if (a === '--out') opts.out = next();
  else if (a === '--stock') opts.stock = next();
  else if (a === '--customers') opts.customersFile = next();
  else if (a === '--map') opts.map = JSON.parse(next());
  else if (a === '--divisor') opts.divisor = Number(next()) || 1;
  else if (a === '--currency') opts.currency = next();
  else if (a === '--excel') opts.excel = true;
  else if (a === '--lang') opts.lang = next();
  else if (a === '--lead-days') opts.leadDays = Number(next()) || 21;
  else if (a === '--sheet') opts.sheet = next();
  else if (a === '--preset') opts.preset = next().toLowerCase();
  else if (a === '--basis') opts.basis = next();
  else if (a === '--date-format') opts.dateFormat = next();
  else if (a === '--decimal') opts.decimal = next();
  else opts.inputs.push(a);
}
if (opts.preset) {
  const p = PRESETS[opts.preset];
  if (!p) { console.error(`Unknown preset ${opts.preset}. Known: ${Object.keys(PRESETS).join(', ')}`); process.exit(1); }
  opts.map = { ...p.map, ...opts.map };
  if (opts.basis === 'auto') opts.basis = p.basis;
}
if (!opts.inputs.length) { console.error('Usage: node analyze.mjs <file-or-folder> [...] [--out dir] [--stock file] [--divisor 10] [--currency Toman] [--excel]'); process.exit(1); }

// ---------- text / number normalization ----------
const FA_DIGITS = '۰۱۲۳۴۵۶۷۸۹', AR_DIGITS = '٠١٢٣٤٥٦٧٨٩';
function normText(s) {
  if (s === null || s === undefined) return '';
  let t = String(s);
  t = t.replace(/[۰-۹]/g, d => FA_DIGITS.indexOf(d)).replace(/[٠-٩]/g, d => AR_DIGITS.indexOf(d));
  t = t.replace(/ي/g, 'ی').replace(/ى/g, 'ی').replace(/ك/g, 'ک').replace(/ة/g, 'ه').replace(/[أإ]/g, 'ا');
  t = t.replace(/[‌‏‎ ]/g, ' ').replace(/\s+/g, ' ').trim();
  return t;
}
const key = s => normText(s).toLowerCase();
const currencySeen = {};
const CURRENCY_RX = /(€|\$|£|¥|₺|₽|kč|czk|eur|usd|gbp|chf|zł|pln|huf|ft|sek|nok|dkk|aed|ریال|تومان|irr|irt)/i;
function num(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;
  let t = normText(v);
  const cur = t.match(CURRENCY_RX);
  if (cur) currencySeen[cur[1].toUpperCase()] = (currencySeen[cur[1].toUpperCase()] || 0) + 1;
  t = t.replace(/[\s'’]/g, '').replace(/٫/g, '.').replace(/[٬،]/g, ',');
  const firstDigit = t.search(/\d/);
  if (firstDigit < 0) return null;
  const neg = /^\(.*\)$/.test(t) || /[-−]/.test(t.slice(0, firstDigit)) || /\d[-−]$/.test(t);
  t = t.replace(/[^\d.,]/g, '');
  const hasDot = t.includes('.'), hasComma = t.includes(',');
  if (hasDot && hasComma) {
    // the separator that appears last is the decimal one: 1.234,56 (EU) or 1,234.56 (US)
    t = t.lastIndexOf(',') > t.lastIndexOf('.') ? t.replace(/\./g, '').replace(',', '.') : t.replace(/,/g, '');
  } else if (hasComma) {
    const thousandsLike = /^\d{1,3}(,\d{3})+$/.test(t);
    if (opts.decimal === 'comma' || (opts.decimal === 'auto' && !thousandsLike && (t.match(/,/g) || []).length === 1)) t = t.replace(',', '.');
    else t = t.replace(/,/g, '');
  } else if (hasDot) {
    const multi = (t.match(/\./g) || []).length > 1;
    if (multi || (opts.decimal === 'comma' && /^\d{1,3}(\.\d{3})+$/.test(t))) t = t.replace(/\./g, '');
  }
  const n = Number(t);
  return isFinite(n) ? (neg ? -n : n) : null;
}

// ---------- dates ----------
function jalaliToGregorian(jy, jm, jd) {
  jy += 1595;
  let days = -355668 + 365 * jy + Math.floor(jy / 33) * 8 + Math.floor(((jy % 33) + 3) / 4) + jd + (jm < 7 ? (jm - 1) * 31 : (jm - 7) * 30 + 186);
  let gy = 400 * Math.floor(days / 146097);
  days %= 146097;
  if (days > 36524) { gy += 100 * Math.floor(--days / 36524); days %= 36524; if (days >= 365) days++; }
  gy += 4 * Math.floor(days / 1461);
  days %= 1461;
  if (days > 365) { gy += Math.floor((days - 1) / 365); days = (days - 1) % 365; }
  let gd = days + 1;
  const ml = [0, 31, (gy % 4 === 0 && gy % 100 !== 0) || gy % 400 === 0 ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  let gm;
  for (gm = 0; gm < 13 && gd > ml[gm]; gm++) gd -= ml[gm];
  return [gy, gm, gd];
}
const MONTH_NAMES = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12, mär: 3, mrz: 3, mai: 5, okt: 10, dez: 12, jän: 1 };
const monthFromName = s => MONTH_NAMES[s.toLowerCase().slice(0, 3)] || null;
const DMY_RX = /(\d{1,2})[\/\-.]\s?(\d{1,2})[\/\-.]\s?(\d{2,4})(?!\d)/;
// Decide whether ambiguous dates like 03/04/2026 are day-first or month-first, from the whole column.
function detectDateOrder(values) {
  if (opts.dateFormat === 'dmy' || opts.dateFormat === 'mdy') return { order: opts.dateFormat, ambiguous: false };
  let dmy = 0, mdy = 0, seen = 0;
  for (const v of values) {
    if (typeof v !== 'string') continue;
    const s = normText(v);
    if (/(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/.test(s)) continue;
    const m = s.match(DMY_RX);
    if (!m) continue;
    seen++;
    if (+m[1] > 12) dmy++;
    else if (+m[2] > 12) mdy++;
  }
  if (dmy && !mdy) return { order: 'dmy', ambiguous: false };
  if (mdy && !dmy) return { order: 'mdy', ambiguous: false };
  return { order: 'dmy', ambiguous: seen > 0 };
}
// returns { t: utc ms (day), y, m, d, cal: 'jalali'|'gregorian' } with y/m/d in the source calendar
function parseDate(v, order = 'dmy') {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date && !isNaN(v)) {
    const t = Date.UTC(v.getFullYear(), v.getMonth(), v.getDate());
    return { t, y: v.getFullYear(), m: v.getMonth() + 1, d: v.getDate(), cal: 'gregorian' };
  }
  if (typeof v === 'number') {
    if (v > 13000000 && v < 15000000) return parseDate(String(v)); // 14050701
    if (v > 20000 && v < 80000) { // Excel serial
      const ms = Math.round((v - 25569) * 86400000);
      const dt = new Date(ms);
      return { t: Date.UTC(dt.getUTCFullYear(), dt.getUTCMonth(), dt.getUTCDate()), y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), cal: 'gregorian' };
    }
    return null;
  }
  const s = normText(v);
  let m = s.match(/(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/);
  let y, mo, d;
  if (m) { y = +m[1]; mo = +m[2]; d = +m[3]; }
  else if ((m = s.match(/^(\d{4})(\d{2})(\d{2})$/))) { y = +m[1]; mo = +m[2]; d = +m[3]; }
  else if ((m = s.match(DMY_RX))) {
    y = +m[3]; if (y < 100) y += 2000;
    if (order === 'mdy') { mo = +m[1]; d = +m[2]; } else { d = +m[1]; mo = +m[2]; }
    if (mo > 12 && d <= 12) [mo, d] = [d, mo];
  }
  else if ((m = s.match(/(\d{1,2})\.?\s+([a-zäéû]{3,})\.?,?\s+(\d{4})/i)) && monthFromName(m[2])) { d = +m[1]; mo = monthFromName(m[2]); y = +m[3]; }
  else if ((m = s.match(/([a-zäéû]{3,})\.?\s+(\d{1,2}),?\s+(\d{4})/i)) && monthFromName(m[1])) { mo = monthFromName(m[1]); d = +m[2]; y = +m[3]; }
  else return null;
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  if (y < 1700) {
    const [gy, gm, gd] = jalaliToGregorian(y, mo, d);
    return { t: Date.UTC(gy, gm - 1, gd), y, m: mo, d, cal: 'jalali' };
  }
  return { t: Date.UTC(y, mo - 1, d), y, m: mo, d, cal: 'gregorian' };
}
const DAY = 86400000;
const WEEKDAYS_FA = ['یکشنبه', 'دوشنبه', 'سه‌شنبه', 'چهارشنبه', 'پنجشنبه', 'جمعه', 'شنبه'];
const WEEKDAYS_EN = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const JALALI_MONTHS = ['فروردین', 'اردیبهشت', 'خرداد', 'تیر', 'مرداد', 'شهریور', 'مهر', 'آبان', 'آذر', 'دی', 'بهمن', 'اسفند'];
const pad = n => String(n).padStart(2, '0');

// ---------- column detection ----------
// Header synonyms: Persian (Iranian accounting: Sepidar, Hamkaran, Mahak, Holoo), English (Shopify, Square,
// WooCommerce, Lightspeed, generic), German and Czech accounting/POS exports. Matching is case-insensitive.
const SYN = {
  invoice: ['شماره فاکتور', 'فاکتور', 'شماره سند', 'شماره رسید', 'invoice', 'invoice no', 'invoice number', 'invoice #', 'order id', 'order', 'order #', 'order no', 'order number', 'order name', 'receipt', 'receipt no', 'receipt number', 'transaction id', 'transaction', 'ticket', 'sale id', 'číslo dokladu', 'doklad', 'číslo faktury', 'faktura', 'číslo objednávky', 'objednávka', 'číslo účtenky', 'rechnungsnummer', 'rechnung', 'belegnummer', 'beleg', 'bestellnummer'],
  date: ['تاریخ', 'تاریخ فاکتور', 'تاریخ فروش', 'date', 'order date', 'invoice date', 'sale date', 'transaction date', 'created at', 'created', 'paid at', 'datetime', 'date time', 'datum', 'datum prodeje', 'datum vystavení', 'datum dokladu', 'datum objednávky', 'rechnungsdatum', 'belegdatum', 'verkaufsdatum'],
  customer: ['نام خریدار', 'خریدار', 'مشتری', 'نام مشتری', 'طرف حساب', 'نام و نام خانوادگی', 'نام کامل', 'customer', 'customer name', 'customer full name', 'client', 'client name', 'buyer', 'billing name', 'full name', 'zákazník', 'jméno zákazníka', 'odběratel', 'kunde', 'kundenname'],
  customerId: ['کد مشتری', 'کد خریدار', 'شماره مشتری', 'customer id', 'customer code', 'customer number', 'client id', 'customer reference id', 'id zákazníka', 'číslo zákazníka', 'kundennummer'],
  email: ['ایمیل', 'پست الکترونیک', 'email', 'e-mail', 'customer email', 'billing email', 'email address'],
  phone: ['موبایل', 'تلفن همراه', 'شماره موبایل', 'تلفن', 'شماره تماس', 'phone', 'mobile', 'phone number', 'billing phone', 'customer phone', 'telefon', 'telefonní číslo', 'mobil'],
  sku: ['کد کالا', 'کد کالا/درآمد', 'کد محصول', 'بارکد', 'sku', 'lineitem sku', 'variant sku', 'product code', 'item code', 'item id', 'barcode', 'product id', 'ean', 'plu', 'kód', 'kód zboží', 'katalogové číslo', 'artikelnummer', 'artikel-nr'],
  product: ['کالا', 'نام کالا', 'کالا/درآمد', 'شرح کالا', 'محصول', 'نام محصول', 'product', 'product name', 'product title', 'item', 'item name', 'lineitem name', 'title', 'description', 'název', 'název zboží', 'zboží', 'položka', 'produkt', 'artikel', 'bezeichnung', 'artikelbezeichnung'],
  category: ['گروه کالا', 'گروه', 'دسته', 'دسته بندی', 'دسته‌بندی', 'category', 'product category', 'product type', 'department', 'group', 'kategorie', 'skupina', 'warengruppe'],
  qty: ['مقدار', 'تعداد', 'qty', 'quantity', 'lineitem quantity', 'units', 'items sold', 'množství', 'počet', 'ks', 'menge', 'anzahl'],
  unitPrice: ['فی', 'قیمت واحد', 'مبلغ واحد', 'unit price', 'price', 'lineitem price', 'item price', 'cena', 'jednotková cena', 'cena za kus', 'cena/ks', 'preis', 'einzelpreis', 'stückpreis'],
  gross: ['مبلغ کل فروش', 'مبلغ کل', 'جمع کل', 'gross', 'gross sales', 'subtotal', 'sub total', 'sales amount', 'celkem bez slevy', 'brutto', 'bruttoumsatz', 'umsatz'],
  discount: ['تخفیفات', 'تخفیف', 'مبلغ تخفیف', 'discount', 'discounts', 'discount amount', 'lineitem discount', 'sleva', 'slevy', 'rabatt'],
  tax: ['عوارض مالیات', 'مالیات', 'مالیات و عوارض', 'ارزش افزوده', 'tax', 'taxes', 'tax amount', 'vat', 'vat amount', 'dph', 'mwst', 'ust', 'steuer'],
  final: ['مبلغ نهایی', 'مبلغ قابل پرداخت', 'total', 'line total', 'order total', 'grand total', 'total price', 'final amount', 'amount', 'total amount', 'celkem', 'celkem s dph', 'k úhradě', 'gesamt', 'gesamtbetrag', 'endbetrag'],
  returnQty: ['مقدار برگشتی', 'تعداد برگشتی', 'return qty', 'returned qty', 'returned quantity', 'refunded qty'],
  returnAmount: ['مبلغ نهایی برگشتی', 'مبلغ کل برگشتی', 'مبلغ برگشتی', 'return amount', 'refund', 'refund amount', 'refunded amount', 'refunds', 'returns', 'vráceno', 'erstattung'],
  net: ['مبلغ کل خالص', 'خالص', 'فروش خالص', 'net', 'net sales', 'net amount', 'net total', 'total sales', 'celkem bez dph', 'netto', 'nettoumsatz'],
  cost: ['بهای تمام شده', 'قیمت خرید', 'قیمت تمام شده', 'cost', 'unit cost', 'cogs', 'cost price', 'nákupní cena', 'einkaufspreis'],
  status: ['وضعیت', 'وضعیت سفارش', 'status', 'order status', 'financial status', 'payment status', 'stav', 'stav objednávky', 'bestellstatus'],
  stock: ['موجودی', 'مانده', 'موجودی انبار', 'stock', 'inventory', 'on hand', 'qty on hand', 'quantity on hand', 'available', 'stock level', 'skladem', 'zásoba', 'stav skladu', 'bestand', 'lagerbestand'],
};
function detectColumns(headers, wanted, overrides = {}) {
  const hk = headers.map(key);
  const map = {};
  const used = new Set();
  for (const [f, h] of Object.entries(overrides)) {
    const i = hk.indexOf(key(h));
    if (i >= 0) { map[f] = i; used.add(i); }
  }
  // pass 1: exact matches, in synonym order
  for (const f of wanted) {
    if (map[f] !== undefined) continue;
    for (const s of SYN[f]) { const i = hk.findIndex((h, j) => h === key(s) && !used.has(j)); if (i >= 0) { map[f] = i; used.add(i); break; } }
  }
  // pass 2: header contains synonym (longest synonyms first to avoid "مبلغ نهایی" grabbing "مبلغ نهایی برگشتی")
  const cands = [];
  for (const f of wanted) if (map[f] === undefined) for (const s of SYN[f]) cands.push([f, key(s)]);
  cands.sort((a, b) => b[1].length - a[1].length);
  for (const [f, s] of cands) {
    if (map[f] !== undefined || s.length < 4) continue;
    const i = hk.findIndex((h, j) => !used.has(j) && h.includes(s));
    if (i >= 0) { map[f] = i; used.add(i); }
  }
  return map;
}
function findHeaderRow(rows) {
  let best = 0, bestScore = -1;
  for (let r = 0; r < Math.min(rows.length, 15); r++) {
    const cells = (rows[r] || []).map(key);
    let score = 0;
    for (const syns of Object.values(SYN)) if (cells.some(c => c && syns.some(s => c === key(s) || c.includes(key(s))))) score++;
    if (score > bestScore) { bestScore = score; best = r; }
  }
  return best;
}
// The xlsx library can hang forever on a damaged zip, so validate the zip structure first.
function assertValidZip(file) {
  const buf = fs.readFileSync(file);
  if (buf.length < 22 || buf.readUInt32LE(0) !== 0x04034b50) throw new Error('not a valid xlsx (zip) file');
  const from = Math.max(0, buf.length - 65557);
  let eocd = -1;
  for (let i = buf.length - 22; i >= from; i--) if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  if (eocd < 0) throw new Error('damaged xlsx: zip directory missing');
  const cdSize = buf.readUInt32LE(eocd + 12), cdOff = buf.readUInt32LE(eocd + 16);
  if (cdOff !== 0xffffffff && (cdOff + cdSize > eocd || buf.readUInt32LE(cdOff) !== 0x02014b50)) throw new Error('damaged xlsx: zip directory offset is invalid (re-save the file in Excel)');
}
function readTable(file) {
  const ext = path.extname(file).toLowerCase();
  if (['.xlsx', '.xlsm'].includes(ext)) assertValidZip(file);
  let wb;
  if (ext === '.csv' || ext === '.tsv' || ext === '.txt') {
    let txt = fs.readFileSync(file, 'utf8').replace(/^﻿/, '');
    wb = XLSX.read(txt, { type: 'string', raw: true, FS: ext === '.tsv' ? '\t' : undefined });
  } else wb = XLSX.readFile(file, { cellDates: true });
  const out = [];
  for (const name of wb.SheetNames) {
    if (opts.sheet && name !== opts.sheet) continue;
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: null, raw: true });
    if (rows.length < 2) continue;
    const h = findHeaderRow(rows);
    out.push({ sheet: name, headers: rows[h].map(x => (x === null ? '' : String(x))), rows: rows.slice(h + 1) });
  }
  return out;
}
function expandInputs(list) {
  const files = [];
  for (const p of list) {
    if (fs.existsSync(p) && fs.statSync(p).isDirectory()) {
      for (const f of fs.readdirSync(p)) if (/\.(xlsx|xls|xlsm|csv|tsv)$/i.test(f) && !f.startsWith('~$')) files.push(path.join(p, f));
    } else files.push(p);
  }
  return files;
}

// ---------- customers ----------
const HONORIFICS = ['سرکار خانم', 'جناب آقای', 'آقای', 'آقا', 'خانم', 'دکتر', 'مهندس', 'حاج', 'حاجی', 'mr.', 'mr', 'mrs.', 'mrs', 'ms.', 'ms', 'miss', 'dr.', 'dr'];
const GENERIC = ['مشتری عمومی', 'عمومی', 'نقدی', 'مشتری نقدی', 'متفرقه', 'فروش نقدی', 'walk-in', 'walk in', 'walk-in customer', 'guest', 'cash', 'cash customer', 'anonymous', 'unknown', 'n/a', '-', 'neznámý', 'hotovost', 'barverkauf', 'laufkunde'];
function cleanName(s) {
  let t = normText(s);
  let changed = true;
  while (changed) {
    changed = false;
    for (const h of HONORIFICS) {
      const lt = t.toLowerCase();
      if (lt.startsWith(h + ' ')) { t = t.slice(h.length).trim(); changed = true; }
    }
  }
  return t;
}
function genderHint(s) {
  const t = normText(s).toLowerCase();
  if (/^(سرکار )?خانم/.test(t) || /^(mrs|ms|miss)\b/.test(t)) return 'F';
  if (/^(جناب )?آقا/.test(t) || /^mr\b/.test(t)) return 'M';
  return '';
}
// Iranian numbers -> 09xxxxxxxxx; other international numbers -> +<country><number>; local numbers kept as digits.
function normPhone(s) {
  const raw = normText(s);
  let t = raw.replace(/\D/g, '');
  if (t.length < 6) return '';
  const intl = /^\s*(\+|00)/.test(raw);
  if (t.startsWith('00')) t = t.slice(2);
  if ((intl || t.length === 12) && t.startsWith('98') && t.length === 12) return '0' + t.slice(2);
  if (intl) return '+' + t;
  if (t.length === 10 && /^9[0-9]{9}$/.test(t) && !/^0/.test(raw)) return '0' + t; // Iranian mobile without leading 0
  return t;
}

// ---------- load sales ----------
const W = ['invoice', 'date', 'customer', 'customerId', 'email', 'phone', 'sku', 'product', 'category', 'qty', 'unitPrice', 'gross', 'discount', 'tax', 'final', 'returnQty', 'returnAmount', 'net', 'cost', 'status'];
const files = expandInputs(opts.inputs);
let lines = [];
const headerLang = { fa: 0, other: 0 };
const quality = { files: [], rowsRead: 0, rowsUsed: 0, skippedNoDate: 0, skippedEmpty: 0, skippedTotals: 0, duplicateRowsAcrossFiles: 0, negativeNetLines: 0, skippedSheets: [], warnings: [] };
const seenRowSig = new Map(); // signature -> file
const revenueBasisCount = {};
for (const file of files) {
  let tables;
  try { tables = readTable(file); } catch (e) { quality.warnings.push(`Could not read ${path.basename(file)}: ${e.message}`); continue; }
  for (const tb of tables) {
    const cm = detectColumns(tb.headers, W, opts.map);
    const colInfo = Object.fromEntries(Object.entries(cm).map(([f, i]) => [f, tb.headers[i]]));
    quality.files.push({ file: path.basename(file), sheet: tb.sheet, rows: tb.rows.length, columns: colInfo });
    const hasItem = cm.product !== undefined || cm.sku !== undefined;
    const hasMoney = ['net', 'final', 'gross', 'unitPrice'].some(k => cm[k] !== undefined);
    if (cm.date === undefined || !hasItem || !hasMoney) {
      quality.files.pop();
      quality.skippedSheets.push({ file: path.basename(file), sheet: tb.sheet, reason: cm.date === undefined ? 'no date column' : 'no product/amount column', headers: tb.headers.filter(Boolean).slice(0, 12) });
      continue;
    }
    const g = (row, f) => (cm[f] === undefined ? null : row[cm[f]]);
    if (headerLang.fa === 0 && headerLang.other === 0) for (const h of tb.headers) { if (/[؀-ۿ]/.test(h)) headerLang.fa++; else if (/[a-z]/i.test(h)) headerLang.other++; }
    // date order (DD/MM vs MM/DD) is decided for the whole column, not per cell
    const dOrder = detectDateOrder(tb.rows.slice(0, 2000).map(r => r && r[cm.date]));
    if (dOrder.ambiguous) quality.warnings.push(`${path.basename(file)}: dates like 03/04 are ambiguous; read as DAY/MONTH. Pass --date-format mdy if they are US-style MONTH/DAY.`);
    // revenue basis: pick the most specific amount column that is filled on (almost) every line.
    // Order-level totals that appear only on an order's first line (Shopify "Total") are thereby avoided.
    const dated = tb.rows.filter(r => r && r[cm.date] !== null && r[cm.date] !== undefined && String(r[cm.date]).trim() !== '');
    const fill = f => (cm[f] === undefined || !dated.length ? 0 : dated.filter(r => num(r[cm[f]]) !== null).length / dated.length);
    let basis = opts.basis;
    if (basis === 'auto') basis = ['net', 'final', 'gross'].find(f => fill(f) >= 0.9) || (cm.unitPrice !== undefined ? 'unit' : ['net', 'final', 'gross'].sort((a, b) => fill(b) - fill(a))[0]);
    const BASIS_LABEL = { net: 'net column', final: 'final - returns', gross: 'gross - discount + tax - returns', unit: 'unit price x qty - discount - returns' };
    for (const row of tb.rows) {
      quality.rowsRead++;
      if (!row || row.every(c => c === null || String(c).trim() === '')) { quality.skippedEmpty++; continue; }
      const firstText = key(row.find(c => c !== null && String(c).trim() !== '') ?? '');
      if (/^(جمع|جمع کل|مجموع|total|grand total|sum|celkem|gesamt|summe)\b/.test(firstText)) { quality.skippedTotals++; continue; }
      const dt = parseDate(g(row, 'date'), dOrder.order);
      if (!dt) { quality.skippedNoDate++; continue; }
      const qty = num(g(row, 'qty')) ?? 1;
      const rq = Math.abs(num(g(row, 'returnQty')) ?? 0);
      const gross = num(g(row, 'gross'));
      const disc = Math.abs(num(g(row, 'discount')) ?? 0); // some POS export discounts as negative numbers
      const tax = num(g(row, 'tax')) ?? 0;
      const fin = num(g(row, 'final'));
      const ret = Math.abs(num(g(row, 'returnAmount')) ?? 0);
      const netCol = num(g(row, 'net'));
      const unit = num(g(row, 'unitPrice'));
      let net = null;
      if (basis === 'net') net = netCol;
      else if (basis === 'final') net = fin === null ? null : fin - ret;
      else if (basis === 'gross') net = gross === null ? null : gross - disc + tax - ret;
      else if (basis === 'unit') net = unit === null ? null : unit * qty - disc - (ret || unit * rq);
      let used = BASIS_LABEL[basis] || basis;
      if (net === null) { // this line lacks the chosen amount: fall back to whatever it has
        if (netCol !== null) { net = netCol; used = 'net column (fallback)'; }
        else if (unit !== null) { net = unit * (qty - rq) - disc; used = 'unit price (fallback)'; }
        else if (fin !== null) { net = fin - ret; used = 'final (fallback)'; }
        else { net = 0; used = 'none'; }
      }
      revenueBasisCount[used] = (revenueBasisCount[used] || 0) + 1;
      const rawCust = g(row, 'customer');
      const phone = normPhone(g(row, 'phone'));
      const email = normText(g(row, 'email')).toLowerCase();
      const cid = normText(g(row, 'customerId'));
      const cname = cleanName(rawCust);
      const generic = !cname || GENERIC.includes(cname.toLowerCase());
      const ckey = phone ? 'p:' + phone : email.includes('@') ? 'e:' + email : cid ? 'id:' + cid : !generic ? 'n:' + key(cname).replace(/\s/g, '') : '';
      const pname = normText(g(row, 'product'));
      const sku = normText(g(row, 'sku'));
      const inv = normText(g(row, 'invoice'));
      const D = opts.divisor;
      const line = {
        inv: inv || `${dt.t}|${ckey || 'walkin'}`, t: dt.t, y: dt.y, m: dt.m, d: dt.d, cal: dt.cal,
        ckey, cname: generic ? '' : cname, gender: genderHint(rawCust), phone, email: email.includes('@') ? email : '',
        status: key(g(row, 'status')), hasInvoiceCol: cm.invoice !== undefined,
        sku, pkey: sku ? 's:' + key(sku) : 'n:' + key(pname), pname, category: normText(g(row, 'category')),
        qty, rq, netQty: qty - rq,
        gross: (gross ?? (unit !== null ? unit * qty : net + disc - tax + ret)) / D,
        disc: disc / D, tax: tax / D, ret: ret / D, net: net / D,
        cost: num(g(row, 'cost')),
      };
      const sig = [line.inv, line.t, line.pkey, line.qty, line.net, line.ckey].join('|');
      const prevFile = seenRowSig.get(sig);
      if (prevFile && prevFile !== file) { quality.duplicateRowsAcrossFiles++; continue; }
      seenRowSig.set(sig, file);
      if (line.net < 0) quality.negativeNetLines++;
      lines.push(line);
      quality.rowsUsed++;
    }
  }
}
quality.revenueBasis = revenueBasisCount;
// drop cancelled / voided / fully refunded / failed orders. Status may sit only on an order's first line
// (Shopify), so it is applied to the whole invoice.
const EXCLUDED_STATUS = /^(cancel+ed|void|voided|failed|refunded|fully refunded|declined|abandoned|draft|trash|deleted|expired)$/;
const EXCLUDED_STATUS_LOCAL = /(لغو|باطل|ناموفق|zrušen|stornov|storniert|abgebrochen|annul)/;
const isExcluded = s => !!s && (EXCLUDED_STATUS.test(s.replace(/[_\-]+/g, ' ').trim()) || EXCLUDED_STATUS_LOCAL.test(s));
const excludedInv = new Set(lines.filter(l => isExcluded(l.status)).map(l => (l.hasInvoiceCol ? l.inv : l)));
if (excludedInv.size) {
  const before = lines.length;
  lines = lines.filter(l => !excludedInv.has(l.hasInvoiceCol ? l.inv : l));
  quality.excludedByStatus = { orders: excludedInv.size, lines: before - lines.length };
  quality.rowsUsed = lines.length;
  quality.warnings.push(`${excludedInv.size} cancelled/voided/refunded order(s) (${before - lines.length} lines) were excluded based on the status column.`);
}
if (opts.lang === 'auto') opts.lang = headerLang.fa > headerLang.other ? 'fa' : 'en';
if (!opts.currency) {
  const ISO = { '€': 'EUR', '$': 'USD', '£': 'GBP', '¥': 'JPY', '₺': 'TRY', '₽': 'RUB', 'KČ': 'CZK', 'ZŁ': 'PLN', 'FT': 'HUF', 'ریال': 'IRR (Rial)', 'تومان': 'Toman' };
  const top = Object.entries(currencySeen).sort((a, b) => b[1] - a[1])[0];
  opts.currency = top ? ISO[top[0]] || top[0] : 'unit';
}
// One customer identity per invoice (Shopify puts name/phone only on the first line), then merge
// identifiers that co-occur (same person seen once with phone, once with email) via union-find.
{
  const byInv = groupBy0(lines.filter(l => l.hasInvoiceCol), l => l.inv);
  for (const ls of byInv.values()) {
    const src = ls.find(l => l.phone) || ls.find(l => l.email) || ls.find(l => l.ckey);
    if (!src) continue;
    for (const l of ls) { l.phone ||= src.phone; l.email ||= src.email; l.cname ||= src.cname; l.gender ||= src.gender; }
  }
  const parent = new Map();
  const find = x => { while (parent.get(x) !== x) { parent.set(x, parent.get(parent.get(x))); x = parent.get(x); } return x; };
  const add = x => { if (!parent.has(x)) parent.set(x, x); };
  const ids = l => [l.phone && 'p:' + l.phone, l.email && 'e:' + l.email, l.ckey].filter(Boolean);
  for (const l of lines) { const xs = ids(l); xs.forEach(add); for (let i = 1; i < xs.length; i++) parent.set(find(xs[i]), find(xs[0])); }
  for (const l of lines) { const xs = ids(l); if (xs.length) l.ckey = find(xs[0]); }
}
function groupBy0(arr, f) { const m = new Map(); for (const x of arr) { const k = f(x); if (!m.has(k)) m.set(k, []); m.get(k).push(x); } return m; }
// money precision: whole units for Rial/Toman-sized prices, cents for EUR/USD-sized prices
const avgAbsLine = lines.reduce((s, l) => s + Math.abs(l.net), 0) / Math.max(1, lines.length);
const MONEY_DEC = avgAbsLine >= 1000 ? 0 : 2;
if (quality.skippedSheets.length) quality.warnings.push(`${quality.skippedSheets.length} sheet(s) did not look like sales data (price lists, customer lists...) and were skipped: ${[...new Set(quality.skippedSheets.map(s => s.file))].join(', ')}. If one IS sales data, pass --map with its column headers.`);
if (quality.duplicateRowsAcrossFiles) quality.warnings.push(`${quality.duplicateRowsAcrossFiles} rows were identical to rows in another file (e.g. "- Copy" files) and were ignored.`);
if (!lines.length) { console.error('No usable sales rows found.\n' + JSON.stringify(quality, null, 2)); process.exit(3); }

// ---------- helpers ----------
const sum = (arr, f) => arr.reduce((s, x) => s + (f(x) || 0), 0);
const round = (n, d = MONEY_DEC) => (n === null || n === undefined || !isFinite(n) ? null : Math.round(n * 10 ** d) / 10 ** d);
const pct = (a, b) => (b ? round((a / b) * 100, 1) : null);
const change = (cur, prev) => (prev ? round(((cur - prev) / Math.abs(prev)) * 100, 1) : null);
function groupBy(arr, f) { const m = new Map(); for (const x of arr) { const k = f(x); if (!m.has(k)) m.set(k, []); m.get(k).push(x); } return m; }
const cal = lines.filter(l => l.cal === 'jalali').length >= lines.length / 2 ? 'jalali' : 'gregorian';
const dateLabel = new Map();
for (const l of lines) if (!dateLabel.has(l.t)) dateLabel.set(l.t, l.cal === 'jalali' ? `${l.y}/${pad(l.m)}/${pad(l.d)}` : null);
const monthOfT = new Map(lines.map(l => [l.t, monthKey(l)]));
const fmtDate = t => {
  const dt = new Date(t);
  const g = `${dt.getUTCFullYear()}-${pad(dt.getUTCMonth() + 1)}-${pad(dt.getUTCDate())}`;
  return cal === 'gregorian' ? g : dateLabel.get(t) || g;
};
function monthKey(l) { return `${l.y}-${pad(l.m)}`; }
const JALALI_MONTHS_EN = ['Farvardin', 'Ordibehesht', 'Khordad', 'Tir', 'Mordad', 'Shahrivar', 'Mehr', 'Aban', 'Azar', 'Dey', 'Bahman', 'Esfand'];
const GREG_MONTHS_EN = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const GREG_MONTHS_FA = ['ژانویه', 'فوریه', 'مارس', 'آوریل', 'مه', 'ژوئن', 'ژوئیه', 'اوت', 'سپتامبر', 'اکتبر', 'نوامبر', 'دسامبر'];
const monthName = m => (cal === 'jalali' ? (opts.lang === 'en' ? JALALI_MONTHS_EN : JALALI_MONTHS) : opts.lang === 'en' ? GREG_MONTHS_EN : GREG_MONTHS_FA)[m - 1];
const monthLabel = k => { const [y, m] = k.split('-').map(Number); return `${monthName(m)} ${y}`; };

// ---------- invoices ----------
const invoices = [...groupBy(lines, l => l.inv).entries()].map(([inv, ls]) => ({
  inv, t: Math.min(...ls.map(l => l.t)), ckey: ls.find(l => l.ckey)?.ckey || '', month: monthKey(ls[0]),
  net: sum(ls, l => l.net), gross: sum(ls, l => l.gross), disc: sum(ls, l => l.disc), qty: sum(ls, l => l.netQty), lines: ls.length,
  hasReturn: ls.some(l => l.rq > 0 || l.ret > 0), products: [...new Set(ls.map(l => l.pkey))],
}));
const saleInvoices = invoices.filter(i => i.net > 0);
const minT = Math.min(...lines.map(l => l.t)), maxT = Math.max(...lines.map(l => l.t));
const spanDays = Math.round((maxT - minT) / DAY) + 1;

// ---------- overview ----------
function overview(ls, invs) {
  const sInv = invs.filter(i => i.net > 0);
  const net = sum(ls, l => l.net), gross = sum(ls, l => l.gross), disc = sum(ls, l => l.disc), ret = sum(ls, l => l.ret);
  const customers = new Set(ls.filter(l => l.ckey).map(l => l.ckey));
  return {
    netRevenue: round(net), grossSales: round(gross), discounts: round(disc), tax: round(sum(ls, l => l.tax)), returns: round(ret),
    invoices: sInv.length, itemsSold: round(sum(ls, l => l.netQty), 2), customers: customers.size,
    avgOrderValue: round(sInv.length ? sum(sInv, i => i.net) / sInv.length : 0),
    itemsPerInvoice: round(sInv.length ? sum(sInv, i => i.qty) / sInv.length : 0, 2),
    avgItemPrice: round(sum(ls, l => l.netQty) ? net / sum(ls, l => l.netQty) : 0),
    discountRatePct: pct(disc, gross), discountedInvoicesPct: pct(sInv.filter(i => i.disc > 0).length, sInv.length),
    returnRatePct: pct(ret, gross), invoicesWithReturns: invs.filter(i => i.hasReturn).length,
    walkInLinesPct: pct(ls.filter(l => !l.ckey).length, ls.length),
  };
}
const totals = overview(lines, invoices);

// ---------- monthly ----------
const byMonth = [...groupBy(lines, monthKey).entries()].sort((a, b) => a[0].localeCompare(b[0]));
const invByMonth = groupBy(invoices, i => i.month);
const monthly = [];
byMonth.forEach(([k, ls], idx) => {
  const o = overview(ls, invByMonth.get(k) || []);
  const prev = monthly[idx - 1];
  monthly.push({ month: k, label: monthLabel(k), ...o, revenueMoMPct: prev ? change(o.netRevenue, prev.netRevenue) : null });
});
// same month across years (YoY)
const yoy = [];
const monthsByNum = groupBy(monthly, r => r.month.split('-')[1]);
for (const [mn, rows] of monthsByNum) {
  if (rows.length < 2) continue;
  rows.sort((a, b) => a.month.localeCompare(b.month));
  for (let i = 1; i < rows.length; i++) {
    const a = rows[i - 1], b = rows[i];
    yoy.push({ month: monthName(+mn), from: a.month, to: b.month,
      revenuePct: change(b.netRevenue, a.netRevenue), invoicesPct: change(b.invoices, a.invoices), itemsPct: change(b.itemsSold, a.itemsSold),
      customersPct: change(b.customers, a.customers), aovPct: change(b.avgOrderValue, a.avgOrderValue),
      prev: { revenue: a.netRevenue, invoices: a.invoices, items: a.itemsSold, customers: a.customers }, cur: { revenue: b.netRevenue, invoices: b.invoices, items: b.itemsSold, customers: b.customers } });
  }
}
yoy.sort((a, b) => a.to.localeCompare(b.to));
// period-over-period when multiple years: compare same set of months
const years = [...new Set(monthly.map(r => r.month.split('-')[0]))].sort();
let yearComparison = null;
if (years.length >= 2) {
  const ya = years[years.length - 2], yb = years[years.length - 1];
  const mb = new Set(monthly.filter(r => r.month.startsWith(yb)).map(r => r.month.split('-')[1]));
  const la = lines.filter(l => String(l.y) === ya && mb.has(pad(l.m))), lb = lines.filter(l => String(l.y) === yb);
  const ia = invoices.filter(i => i.month.startsWith(ya) && mb.has(i.month.split('-')[1])), ib = invoices.filter(i => i.month.startsWith(yb));
  const A = overview(la, ia), B = overview(lb, ib);
  const deltas = {};
  for (const k of Object.keys(B)) if (typeof B[k] === 'number' && typeof A[k] === 'number') deltas[k] = change(B[k], A[k]);
  yearComparison = { note: `Same months only (${[...mb].sort().join(',')}) of ${ya} vs ${yb}`, [ya]: A, [yb]: B, changePct: deltas };
}

// ---------- weekday ----------
const weekday = [...groupBy(saleInvoices, i => new Date(i.t).getUTCDay()).entries()]
  .map(([d, inv]) => ({ day: opts.lang === 'en' ? WEEKDAYS_EN[d] : WEEKDAYS_FA[d], dow: d, invoices: inv.length, revenue: round(sum(inv, i => i.net)), sharePct: pct(sum(inv, i => i.net), sum(saleInvoices, i => i.net)) }))
  // Iranian week starts Saturday; ISO/European week starts Monday
  .sort((a, b) => (cal === 'jalali' ? ((a.dow + 1) % 7) - ((b.dow + 1) % 7) : ((a.dow + 6) % 7) - ((b.dow + 6) % 7)));
// day-of-month thirds (salary-cycle effect)
const domBuckets = { '1-10': 0, '11-20': 0, '21-31': 0 };
for (const l of lines) domBuckets[l.d <= 10 ? '1-10' : l.d <= 20 ? '11-20' : '21-31'] += l.net;
const dayOfMonth = Object.entries(domBuckets).map(([k, v]) => ({ days: k, revenue: round(v), sharePct: pct(v, totals.netRevenue) }));

// ---------- products ----------
const mid = minT + (maxT - minT) / 2;
const prodGroups = groupBy(lines, l => l.pkey);
let products = [...prodGroups.entries()].map(([pk, ls]) => {
  const names = groupBy(ls.filter(l => l.pname), l => l.pname);
  const name = [...names.entries()].sort((a, b) => b[1].length - a[1].length)[0]?.[0] || ls[0].sku;
  const qty = sum(ls, l => l.netQty), rev = sum(ls, l => l.net);
  const h1 = sum(ls.filter(l => l.t <= mid), l => l.net), h2 = sum(ls.filter(l => l.t > mid), l => l.net);
  const last = Math.max(...ls.map(l => l.t));
  return { key: pk, sku: ls[0].sku, name, model: name.split(' ')[0] || name, category: ls[0].category, qty: round(qty, 2), revenue: round(rev),
    invoices: new Set(ls.map(l => l.inv)).size, customers: new Set(ls.filter(l => l.ckey).map(l => l.ckey)).size,
    avgPrice: round(qty ? rev / qty : 0), discountPct: pct(sum(ls, l => l.disc), sum(ls, l => l.gross)), returnedQty: round(sum(ls, l => l.rq), 2),
    firstHalfRevenue: round(h1), secondHalfRevenue: round(h2), trendPct: change(h2, h1),
    lastSale: fmtDate(last), daysSinceLastSale: Math.round((maxT - last) / DAY) };
}).sort((a, b) => b.revenue - a.revenue);
let cum = 0;
const totRev = sum(products, p => Math.max(p.revenue, 0));
// a product is A if the cumulative share BEFORE it is under 80% (so the item crossing 80% is still A)
for (const p of products) { const before = totRev ? (cum / totRev) * 100 : 0; cum += Math.max(p.revenue, 0); p.sharePct = pct(p.revenue, totRev); p.cumSharePct = pct(cum, totRev); p.abc = before < 80 ? 'A' : before < 95 ? 'B' : 'C'; }
const abcSummary = ['A', 'B', 'C'].map(c => { const ps = products.filter(p => p.abc === c); return { class: c, products: ps.length, revenue: round(sum(ps, p => p.revenue)), revenueSharePct: pct(sum(ps, p => p.revenue), totRev), productSharePct: pct(ps.length, products.length) }; });
function rollup(field) {
  const g = groupBy(products.filter(p => p[field]), p => p[field]);
  return [...g.entries()].map(([k, ps]) => ({ [field]: k, products: ps.length, qty: round(sum(ps, p => p.qty), 2), revenue: round(sum(ps, p => p.revenue)), sharePct: pct(sum(ps, p => p.revenue), totRev) })).sort((a, b) => b.revenue - a.revenue);
}
const categories = rollup('category');
const models = rollup('model').filter(m => m.products > 1 || m.revenue > 0).slice(0, 40);

// ---------- basket ----------
const multi = saleInvoices.filter(i => i.products.length > 1);
const pairCount = new Map();
for (const i of multi) {
  const ps = [...i.products].sort();
  for (let a = 0; a < ps.length; a++) for (let b = a + 1; b < ps.length; b++) { const k = ps[a] + '||' + ps[b]; pairCount.set(k, (pairCount.get(k) || 0) + 1); }
}
const pname = new Map(products.map(p => [p.key, p.name]));
const basket = { multiItemInvoicesPct: pct(multi.length, saleInvoices.length),
  topPairs: [...pairCount.entries()].sort((a, b) => b[1] - a[1]).slice(0, 15).map(([k, n]) => { const [a, b] = k.split('||'); return { a: pname.get(a), b: pname.get(b), invoices: n }; }) };
const modelPair = new Map();
for (const i of multi) {
  const ms = [...new Set(i.products.map(p => (pname.get(p) || '').split(' ')[0]))].sort();
  for (let a = 0; a < ms.length; a++) for (let b = a + 1; b < ms.length; b++) { const k = ms[a] + ' + ' + ms[b]; modelPair.set(k, (modelPair.get(k) || 0) + 1); }
}
basket.topModelPairs = [...modelPair.entries()].sort((a, b) => b[1] - a[1]).slice(0, 10).map(([pair, invoices]) => ({ pair, invoices }));

// ---------- discounts ----------
const dInv = saleInvoices.filter(i => i.disc > 0), fInv = saleInvoices.filter(i => i.disc <= 0);
const discounts = { discountedInvoices: dInv.length, fullPriceInvoices: fInv.length,
  avgDiscountPctOnDiscounted: pct(sum(dInv, i => i.disc), sum(dInv, i => i.gross)),
  aovDiscounted: round(dInv.length ? sum(dInv, i => i.net) / dInv.length : 0), aovFullPrice: round(fInv.length ? sum(fInv, i => i.net) / fInv.length : 0),
  itemsPerInvoiceDiscounted: round(dInv.length ? sum(dInv, i => i.qty) / dInv.length : 0, 2), itemsPerInvoiceFullPrice: round(fInv.length ? sum(fInv, i => i.qty) / fInv.length : 0, 2),
  byMonth: monthly.map(m => ({ month: m.month, label: m.label, discountRatePct: m.discountRatePct, revenue: m.netRevenue })) };

// ---------- customers / RFM ----------
const custLines = groupBy(lines.filter(l => l.ckey), l => l.ckey);
const custInv = groupBy(saleInvoices.filter(i => i.ckey), i => i.ckey);
let customers = [...custLines.entries()].map(([ck, ls]) => {
  const invs = custInv.get(ck) || [];
  const dates = [...new Set(invs.map(i => i.t))].sort((a, b) => a - b);
  const first = dates[0] ?? Math.min(...ls.map(l => l.t)), last = dates[dates.length - 1] ?? Math.max(...ls.map(l => l.t));
  const named = ls.find(l => l.cname);
  const cats = groupBy(ls, l => l.category || l.pname.split(' ')[0]);
  return { key: ck, firstT: first, name: named?.cname || '', phone: ls.find(l => l.phone)?.phone || '', email: ls.find(l => l.email)?.email || '', gender: ls.find(l => l.gender)?.gender || '',
    frequency: dates.length, visits: dates.length, invoices: invs.length, monetary: round(sum(ls, l => l.net)), items: round(sum(ls, l => l.netQty), 2),
    firstPurchase: fmtDate(first), lastPurchase: fmtDate(last), recencyDays: Math.round((maxT - last) / DAY), tenureDays: Math.round((last - first) / DAY),
    avgGapDays: dates.length > 1 ? Math.round((last - first) / DAY / (dates.length - 1)) : null,
    favourite: [...cats.entries()].sort((a, b) => sum(b[1], l => l.net) - sum(a[1], l => l.net))[0]?.[0] || '',
    usedDiscount: ls.some(l => l.disc > 0) };
});
// fill phones from a customer database, matching on cleaned name (ambiguous names are flagged, not guessed)
if (opts.customersFile) {
  const book = new Map(), emails = new Map();
  try {
    for (const tb of readTable(opts.customersFile)) {
      const cm = detectColumns(tb.headers, ['customer', 'customerId', 'phone'], opts.map);
      if (cm.customer === undefined) { quality.warnings.push(`Customer file sheet ${tb.sheet}: needs a name column (plus phone and/or email).`); continue; }
      // use every phone-like column; prefer an Iranian mobile (09...) over a landline
      const phoneCols = tb.headers.map((h, i) => [key(h), i]).filter(([h]) => /تلفن|موبایل|همراه|phone|mobile|tel|mobil/.test(h)).map(([, i]) => i);
      const emailCol = detectColumns(tb.headers, ['email'], opts.map).email;
      for (const row of tb.rows) {
        const phones = phoneCols.map(i => normPhone(row[i])).filter(Boolean);
        const n = key(cleanName(row[cm.customer])).replace(/\s/g, ''), ph = phones.find(p => /^(09\d{9}|\+\d{9,})$/.test(p)) || phones[0];
        const em = emailCol !== undefined ? normText(row[emailCol]).toLowerCase() : '';
        if (!n) continue;
        if (ph) { if (!book.has(n)) book.set(n, new Set()); book.get(n).add(ph); }
        if (em.includes('@') && !emails.has(n)) emails.set(n, em);
      }
    }
  } catch (e) { quality.warnings.push(`Could not read customer file: ${e.message}`); }
  let filled = 0, ambiguous = 0;
  for (const c of customers) {
    const nk = key(c.name).replace(/\s/g, '');
    if (!c.email && emails.has(nk)) c.email = emails.get(nk);
    if (c.phone) continue;
    const ph = book.get(nk);
    if (!ph) continue;
    if (ph.size === 1) { c.phone = [...ph][0]; filled++; } else { c.phoneCandidates = [...ph].join(' / '); ambiguous++; }
  }
  quality.customerPhoneFill = { databaseNames: book.size, filled, ambiguousNames: ambiguous, stillMissing: customers.filter(c => !c.phone).length };
}
// monetary quintile
const sortedM = customers.map(c => c.monetary).sort((a, b) => a - b);
const quint = v => { if (!sortedM.length) return 0; let lo = 0, hi = sortedM.length; while (lo < hi) { const m = (lo + hi) >> 1; if (sortedM[m] < v) lo = m + 1; else hi = m; } return Math.min(5, Math.floor((lo / sortedM.length) * 5) + 1); };
// thresholds scale with data span so short exports still segment sensibly
const RECENT = Math.max(30, Math.round(spanDays * 0.15)), ACTIVE = Math.max(90, Math.round(spanDays * 0.35)), LAPSED = Math.max(180, Math.round(spanDays * 0.6));
const SEG = {
  vip: { fa: 'VIP / قهرمان', en: 'VIP / Champions' }, loyal: { fa: 'وفادار', en: 'Loyal' }, atRisk: { fa: 'در خطر ریزش', en: 'At risk' },
  lostRepeat: { fa: 'تکراری از‌دست‌رفته', en: 'Lost repeat' }, highValueOnce: { fa: 'یک‌بار خرید ارزشمند', en: 'High-value one-timer' },
  newRecent: { fa: 'جدید', en: 'New' }, oneTimeCooling: { fa: 'یک‌بار خرید - رو به سرد شدن', en: 'One-time cooling' }, oneTimeLost: { fa: 'یک‌بار خرید - غیرفعال', en: 'One-time dormant' },
};
for (const c of customers) {
  c.mScore = quint(c.monetary);
  const f = c.frequency, r = c.recencyDays;
  let s;
  if (f >= 2) s = r <= ACTIVE && (f >= 3 || c.mScore >= 4) ? 'vip' : r <= LAPSED ? (r <= ACTIVE ? 'loyal' : 'atRisk') : 'lostRepeat';
  else if (r <= RECENT) s = c.mScore >= 5 ? 'highValueOnce' : 'newRecent';
  else if (r <= ACTIVE) s = c.mScore >= 5 ? 'highValueOnce' : 'oneTimeCooling';
  else s = c.mScore >= 5 && r <= LAPSED ? 'highValueOnce' : 'oneTimeLost';
  c.segment = s; c.segmentLabel = SEG[s][opts.lang === 'en' ? 'en' : 'fa'];
}
customers.sort((a, b) => b.monetary - a.monetary);
const custRevTotal = sum(customers, c => c.monetary);
const segments = Object.keys(SEG).map(s => { const cs = customers.filter(c => c.segment === s); return { segment: s, label: SEG[s][opts.lang === 'en' ? 'en' : 'fa'], customers: cs.length, revenue: round(sum(cs, c => c.monetary)), revenueSharePct: pct(sum(cs, c => c.monetary), custRevTotal), withPhone: cs.filter(c => c.phone).length, withEmail: cs.filter(c => c.email).length, avgRecencyDays: cs.length ? Math.round(sum(cs, c => c.recencyDays) / cs.length) : null }; }).filter(s => s.customers);
const repeaters = customers.filter(c => c.frequency >= 2);
const top10n = Math.max(1, Math.round(customers.length * 0.1));
const customerStats = {
  identifiedCustomers: customers.length, repeatCustomers: repeaters.length, repeatRatePct: pct(repeaters.length, customers.length),
  repeatRevenueSharePct: pct(sum(repeaters, c => c.monetary), custRevTotal),
  top10PctCustomersRevenueSharePct: pct(sum(customers.slice(0, top10n), c => c.monetary), custRevTotal),
  avgCustomerValue: round(customers.length ? custRevTotal / customers.length : 0),
  medianDaysBetweenVisits: (() => { const g = repeaters.map(c => c.avgGapDays).filter(x => x !== null).sort((a, b) => a - b); return g.length ? g[Math.floor(g.length / 2)] : null; })(),
  withPhonePct: pct(customers.filter(c => c.phone).length, customers.length),
  withEmailPct: pct(customers.filter(c => c.email).length, customers.length),
  genderSplit: { F: customers.filter(c => c.gender === 'F').length, M: customers.filter(c => c.gender === 'M').length, unknown: customers.filter(c => !c.gender).length },
  newCustomersByMonth: [...groupBy(customers, c => monthOfT.get(c.firstT) || '?').entries()].sort().map(([month, cs]) => ({ month, label: monthLabel(month), newCustomers: cs.length })),
  segmentThresholdsDays: { recent: RECENT, active: ACTIVE, lapsed: LAPSED },
};

// ---------- inventory ----------
let inventory = null;
const productSlowMovers = products.filter(p => p.daysSinceLastSale >= 60 && p.revenue > 0).sort((a, b) => b.daysSinceLastSale - a.daysSinceLastSale).slice(0, 30)
  .map(p => ({ name: p.name, sku: p.sku, lastSale: p.lastSale, daysSinceLastSale: p.daysSinceLastSale, qtySold: p.qty }));
if (opts.stock) {
  const tbs = readTable(opts.stock);
  const items = [];
  for (const tb of tbs) {
    const cm = detectColumns(tb.headers, ['sku', 'product', 'category', 'stock', 'cost', 'unitPrice'], opts.map);
    if (cm.stock === undefined) { quality.warnings.push(`Stock file sheet ${tb.sheet}: no stock column found.`); continue; }
    for (const row of tb.rows) {
      const st = num(row[cm.stock]);
      if (st === null) continue;
      const sku = cm.sku !== undefined ? normText(row[cm.sku]) : '', name = cm.product !== undefined ? normText(row[cm.product]) : '';
      if (!sku && !name) continue;
      items.push({ sku, name, stock: st, cost: cm.cost !== undefined ? num(row[cm.cost]) : null, price: cm.unitPrice !== undefined ? num(row[cm.unitPrice]) : null, category: cm.category !== undefined ? normText(row[cm.category]) : '' });
    }
  }
  const bySku = new Map(products.filter(p => p.sku).map(p => [key(p.sku), p])), byName = new Map(products.map(p => [key(p.name), p]));
  const win = Math.min(90, spanDays);
  const recentLines = groupBy(lines.filter(l => l.t > maxT - win * DAY), l => l.pkey);
  const rows = items.map(it => {
    const p = (it.sku && bySku.get(key(it.sku))) || byName.get(key(it.name));
    const sold = p ? sum(recentLines.get(p.key) || [], l => l.netQty) : 0;
    const vel = sold / win;
    const cover = vel > 0 ? Math.round(it.stock / vel) : null;
    let status;
    if (it.stock <= 0) status = p && sold > 0 ? 'out_of_stock_selling' : 'out_of_stock';
    else if (!p || sold === 0) status = !p ? 'dead_never_sold' : 'dead_no_recent_sales';
    else if (cover < opts.leadDays) status = 'low_stock';
    else if (cover > 180) status = 'overstock';
    else status = 'healthy';
    const value = it.stock > 0 ? it.stock * ((it.cost ?? it.price ?? (p ? p.avgPrice * opts.divisor : 0)) / opts.divisor) : 0;
    return { sku: it.sku, name: it.name || p?.name, category: it.category || p?.category || '', stock: it.stock, soldLastWindow: round(sold, 2), windowDays: win, dailyVelocity: round(vel, 3), daysOfCover: cover, abc: p?.abc || '-', status, stockValue: round(value) };
  });
  const st = groupBy(rows, r => r.status);
  inventory = { items: rows.length, matchedToSales: rows.filter(r => r.status !== 'dead_never_sold' && r.status !== 'out_of_stock').length,
    statusSummary: [...st.entries()].map(([s, rs]) => ({ status: s, items: rs.length, units: round(sum(rs, r => r.stock), 2), stockValue: round(sum(rs, r => r.stockValue)) })),
    lowStock: rows.filter(r => r.status === 'low_stock' || r.status === 'out_of_stock_selling').sort((a, b) => (a.daysOfCover ?? -1) - (b.daysOfCover ?? -1)).slice(0, 40),
    deadStock: rows.filter(r => r.status.startsWith('dead')).sort((a, b) => b.stockValue - a.stockValue).slice(0, 40),
    overstock: rows.filter(r => r.status === 'overstock').sort((a, b) => b.stockValue - a.stockValue).slice(0, 20),
    all: rows };
}

// ---------- assemble ----------
const result = {
  meta: { generatedAt: new Date().toISOString(), files: files.map(f => path.basename(f)), calendar: cal, from: fmtDate(minT), to: fmtDate(maxT), spanDays, currency: opts.currency, divisor: opts.divisor },
  dataQuality: quality, totals, yearComparison, monthly, yoy, weekday, dayOfMonth,
  products: { count: products.length, top: products.slice(0, 25), bottom: products.filter(p => p.revenue > 0).slice(-15).reverse(), abcSummary, risingStars: products.filter(p => p.firstHalfRevenue > 0 && p.trendPct >= 50 && p.revenue > 0).sort((a, b) => b.secondHalfRevenue - a.secondHalfRevenue).slice(0, 10), declining: products.filter(p => p.abc !== 'C' && p.trendPct !== null && p.trendPct <= -40).slice(0, 10), mostReturned: products.filter(p => p.returnedQty > 0).sort((a, b) => b.returnedQty - a.returnedQty).slice(0, 10), slowMovers: productSlowMovers },
  categories, models, basket, discounts,
  customers: { stats: customerStats, segments, top: customers.slice(0, 25).map(({ key, firstT, ...c }) => c) },
  inventory: inventory ? { ...inventory, all: undefined } : null,
};
fs.mkdirSync(opts.out, { recursive: true });
fs.writeFileSync(path.join(opts.out, 'analysis.json'), JSON.stringify(result, null, 1));
// full lists for campaigns / drill-down
const csvEsc = v => { const s = v === null || v === undefined ? '' : String(v); return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s; };
const toCsv = (rows, cols) => '﻿' + [cols.join(','), ...rows.map(r => cols.map(c => csvEsc(r[c])).join(','))].join('\n');
const custCols = ['name', 'phone', 'phoneCandidates', 'email', 'gender', 'segment', 'segmentLabel', 'frequency', 'monetary', 'items', 'firstPurchase', 'lastPurchase', 'recencyDays', 'avgGapDays', 'favourite', 'usedDiscount'];
fs.writeFileSync(path.join(opts.out, 'customers.csv'), toCsv(customers, custCols));
const prodCols = ['sku', 'name', 'model', 'category', 'qty', 'revenue', 'sharePct', 'cumSharePct', 'abc', 'invoices', 'customers', 'avgPrice', 'discountPct', 'returnedQty', 'trendPct', 'lastSale', 'daysSinceLastSale'];
fs.writeFileSync(path.join(opts.out, 'products.csv'), toCsv(products, prodCols));
if (inventory) fs.writeFileSync(path.join(opts.out, 'inventory.csv'), toCsv(inventory.all, Object.keys(inventory.all[0] || { sku: 1 })));

if (opts.excel) {
  const fa = opts.lang !== 'en';
  const wb = XLSX.utils.book_new();
  wb.Workbook = { Views: [{ RTL: fa }] };
  const add = (name, rows) => { if (rows && rows.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), name.slice(0, 31)); };
  const L = fa ? { netRevenue: 'فروش خالص', grossSales: 'فروش ناخالص', discounts: 'تخفیف', tax: 'مالیات', returns: 'برگشتی', invoices: 'تعداد فاکتور', itemsSold: 'تعداد کالا', customers: 'تعداد مشتری', avgOrderValue: 'میانگین فاکتور', itemsPerInvoice: 'کالا در هر فاکتور', avgItemPrice: 'میانگین قیمت کالا', discountRatePct: 'درصد تخفیف', discountedInvoicesPct: 'درصد فاکتور با تخفیف', returnRatePct: 'درصد برگشتی', invoicesWithReturns: 'فاکتور با برگشتی', walkInLinesPct: 'درصد فروش بدون نام مشتری' } : {};
  add(fa ? 'خلاصه' : 'Summary', [
    { [fa ? 'شاخص' : 'Metric']: fa ? 'بازه' : 'Period', [fa ? 'مقدار' : 'Value']: `${result.meta.from} → ${result.meta.to}` },
    { [fa ? 'شاخص' : 'Metric']: fa ? 'واحد پول' : 'Currency', [fa ? 'مقدار' : 'Value']: opts.currency },
    ...Object.entries(totals).map(([k, v]) => ({ [fa ? 'شاخص' : 'Metric']: L[k] || k, [fa ? 'مقدار' : 'Value']: v })),
  ]);
  add(fa ? 'ماهانه' : 'Monthly', monthly);
  if (yoy.length) add(fa ? 'مقایسه سالانه' : 'YoY', yoy.map(({ prev, cur, ...r }) => ({ ...r, prevRevenue: prev.revenue, curRevenue: cur.revenue, prevInvoices: prev.invoices, curInvoices: cur.invoices })));
  add(fa ? 'روز هفته' : 'Weekday', weekday.map(({ dow, ...r }) => r));
  add(fa ? 'محصولات' : 'Products', products.map(p => Object.fromEntries(prodCols.map(c => [c, p[c]]))));
  add(fa ? 'سگمنت‌ها' : 'Segments', segments);
  add(fa ? 'مشتریان' : 'Customers', customers.map(c => Object.fromEntries(custCols.map(k => [k, c[k]]))));
  if (inventory) add(fa ? 'موجودی' : 'Inventory', inventory.all);
  add(fa ? 'کیفیت داده' : 'Data quality', [...quality.files.map(f => ({ file: f.file, sheet: f.sheet, rows: f.rows, columns: JSON.stringify(f.columns) })), ...quality.warnings.map(w => ({ file: 'warning', sheet: w }))]);
  XLSX.writeFile(wb, path.join(opts.out, 'report.xlsx'));
}

// ---------- console summary ----------
const f = n => (n === null || n === undefined ? '-' : Number(n).toLocaleString('en-US'));
console.log(`Period ${result.meta.from} -> ${result.meta.to} (${spanDays} days, ${cal}), ${files.length} file(s), ${quality.rowsUsed}/${quality.rowsRead} rows used`);
console.log(`Revenue basis: ${JSON.stringify(revenueBasisCount)}`);
console.log(`Net revenue ${f(totals.netRevenue)} ${opts.currency} | invoices ${f(totals.invoices)} | items ${f(totals.itemsSold)} | customers ${f(totals.customers)} | AOV ${f(totals.avgOrderValue)}`);
console.log(`Repeat rate ${customerStats.repeatRatePct}% | discount rate ${totals.discountRatePct}% | returns ${totals.returnRatePct}% | A-class products ${abcSummary[0].products}/${products.length}`);
if (quality.warnings.length) console.log('Warnings:\n - ' + quality.warnings.join('\n - '));
console.log(`Wrote ${path.resolve(opts.out)}`);
