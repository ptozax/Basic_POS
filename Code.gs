/************************************************************
 *  LocalPOS — Google Apps Script backend (Cloud + Login)
 *  ----------------------------------------------------------
 *  เก็บข้อมูลบน Google Sheets แบบ key-value (chunked) และมีระบบ
 *  login ด้วย username/password (hash SHA-256) + token (CacheService)
 *
 *  วิธีติดตั้งย่อ ๆ (ดูละเอียดใน README-DEPLOY.md):
 *   1) สร้าง Google Sheet ใหม่ 1 ไฟล์ (เปล่า ๆ ก็ได้)
 *   2) เมนู Extensions > Apps Script
 *   3) วางไฟล์นี้ลงใน Code.gs และสร้างไฟล์ HTML ชื่อ "Index"
 *      แล้ววางเนื้อหา index.html ทั้งหมดลงไป
 *   4) รันฟังก์ชัน setup() หนึ่งครั้ง (จะสร้างชีต + ผู้ใช้ admin)
 *   5) Deploy > New deployment > Web app
 *        - Execute as: Me
 *        - Who has access: Anyone (หรือ Anyone with Google account)
 *   6) เปิด URL ที่ได้ แล้ว login ด้วย admin / admin123 (รีบเปลี่ยนรหัส)
 ************************************************************/

/* ==================== CONFIG ==================== */
var STORE_SHEET = '_store';       // เก็บข้อมูล key-value
var USER_SHEET  = '_users';       // เก็บบัญชีผู้ใช้
var CHUNK_SIZE  = 45000;          // ตัดข้อความต่อ 1 แถว (ลิมิตช่อง Sheet = 50,000)
var SESSION_TTL = 6 * 60 * 60;    // token อายุ 6 ชั่วโมง (วินาที)
var CACHE_PREFIX = 'possess_';

/* ==================== WEB APP ENTRY ==================== */
function doGet() {
  // ใช้ createHtmlOutputFromFile (เสิร์ฟ HTML ตรง ๆ) แทน createTemplateFromFile
  // เพราะไฟล์มีข้อความ <?xml ... ?> ซึ่ง template engine จะตีความเป็น scriptlet แล้ว error
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบ POS — LocalPOS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ==================== SETUP (รันครั้งเดียว) ==================== */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('เปิดสคริปต์นี้จากภายใน Google Sheet (Extensions > Apps Script)');

  // สร้างชีตเก็บข้อมูล
  var store = ss.getSheetByName(STORE_SHEET);
  if (!store) {
    store = ss.insertSheet(STORE_SHEET);
    store.getRange(1, 1, 1, 3).setValues([['key', 'idx', 'value']]);
    store.setFrozenRows(1);
  }

  // สร้างชีตผู้ใช้ + admin เริ่มต้น
  var users = ss.getSheetByName(USER_SHEET);
  if (!users) {
    users = ss.insertSheet(USER_SHEET);
    users.getRange(1, 1, 1, 6).setValues([['username', 'salt', 'hash', 'displayName', 'role', 'active']]);
    users.setFrozenRows(1);
  }
  if (users.getLastRow() < 2) {
    _createUser_(users, 'admin', 'admin123', 'ผู้ดูแลระบบ', 'admin');
  }
  return 'ตั้งค่าเรียบร้อย — login: admin / admin123 (โปรดเปลี่ยนรหัสผ่าน)';
}

/* ==================== AUTH ==================== */
function login(username, password) {
  username = String(username || '').trim().toLowerCase();
  var users = _sheet_(USER_SHEET);
  var data = users.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[0]).trim().toLowerCase() !== username) continue;
    if (String(row[5]).toLowerCase() === 'false' || row[5] === false) {
      return { ok: false, error: 'บัญชีนี้ถูกปิดใช้งาน' };
    }
    var salt = String(row[1]);
    var hash = String(row[2]);
    if (_sha256_(salt + password) !== hash) {
      return { ok: false, error: 'รหัสผ่านไม่ถูกต้อง' };
    }
    var user = { username: String(row[0]), displayName: String(row[3] || row[0]), role: String(row[4] || 'staff') };
    var token = Utilities.getUuid();
    CacheService.getScriptCache().put(CACHE_PREFIX + token, JSON.stringify(user), SESSION_TTL);
    return { ok: true, token: token, user: user };
  }
  return { ok: false, error: 'ไม่พบชื่อผู้ใช้นี้' };
}

// ตรวจ token ที่ client ถืออยู่ (ใช้ตอนเปิดหน้าใหม่/รีเฟรช)
function resume(token) {
  var u = _session_(token);
  return u ? { ok: true, user: u } : { ok: false };
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove(CACHE_PREFIX + token);
  return { ok: true };
}

// เปลี่ยนรหัสผ่านของตัวเอง
function changePassword(token, oldPass, newPass) {
  var u = _requireSession_(token);
  if (!newPass || String(newPass).length < 4) return { ok: false, error: 'รหัสผ่านใหม่ต้องยาวอย่างน้อย 4 ตัว' };
  var users = _sheet_(USER_SHEET);
  var data = users.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === u.username.toLowerCase()) {
      if (_sha256_(String(data[i][1]) + oldPass) !== String(data[i][2])) {
        return { ok: false, error: 'รหัสผ่านเดิมไม่ถูกต้อง' };
      }
      var salt = Utilities.getUuid();
      users.getRange(i + 1, 2).setValue(salt);
      users.getRange(i + 1, 3).setValue(_sha256_(salt + newPass));
      return { ok: true };
    }
  }
  return { ok: false, error: 'ไม่พบผู้ใช้' };
}

// จัดการผู้ใช้ (เฉพาะ admin): เพิ่ม/รีเซ็ตรหัส
function addUser(token, username, password, displayName, role) {
  var u = _requireSession_(token);
  if (u.role !== 'admin') return { ok: false, error: 'เฉพาะผู้ดูแลระบบเท่านั้น' };
  username = String(username || '').trim();
  if (!username || !password) return { ok: false, error: 'กรอกชื่อผู้ใช้และรหัสผ่าน' };
  var users = _sheet_(USER_SHEET);
  var data = users.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === username.toLowerCase()) {
      return { ok: false, error: 'มีชื่อผู้ใช้นี้อยู่แล้ว' };
    }
  }
  _createUser_(users, username, password, displayName || username, role || 'staff');
  return { ok: true };
}

function _createUser_(sheet, username, password, displayName, role) {
  var salt = Utilities.getUuid();
  sheet.appendRow([username, salt, _sha256_(salt + password), displayName, role, true]);
}

/* ==================== DATA STORE (key-value, chunked) ==================== */
// โหลดข้อมูลทั้งหมดคืนเป็น object { key: parsedValue }
function loadAll(token) {
  _requireSession_(token);
  var sheet = _sheet_(STORE_SHEET);
  var last = sheet.getLastRow();
  var out = {};
  if (last < 2) return { ok: true, data: out };
  var rows = sheet.getRange(2, 1, last - 1, 3).getValues();
  var buckets = {}; // key -> [{idx, value}]
  for (var i = 0; i < rows.length; i++) {
    var k = rows[i][0];
    if (k === '' || k === null) continue;
    (buckets[k] = buckets[k] || []).push({ idx: Number(rows[i][1]) || 0, value: String(rows[i][2] == null ? '' : rows[i][2]) });
  }
  for (var key in buckets) {
    buckets[key].sort(function (a, b) { return a.idx - b.idx; });
    var joined = buckets[key].map(function (c) { return c.value; }).join('');
    try { out[key] = joined === '' ? null : JSON.parse(joined); }
    catch (e) { out[key] = null; }
  }
  return { ok: true, data: out };
}

// บันทึกหลาย key พร้อมกัน  batch = { key: <ค่าที่ JSON.stringify มาแล้ว หรือ ค่าดิบ> }
function saveMulti(token, batch) {
  _requireSession_(token);
  if (!batch) return { ok: true };
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = _sheet_(STORE_SHEET);
    for (var key in batch) {
      var raw = batch[key];
      var text = (typeof raw === 'string') ? raw : JSON.stringify(raw);
      _writeKey_(sheet, key, text);
    }
    SpreadsheetApp.flush();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// ลบแถวเดิมของ key แล้วเขียนใหม่เป็น chunk
function _writeKey_(sheet, key, text) {
  var last = sheet.getLastRow();
  if (last >= 2) {
    var keys = sheet.getRange(2, 1, last - 1, 1).getValues();
    for (var r = keys.length - 1; r >= 0; r--) {
      if (keys[r][0] === key) sheet.deleteRow(r + 2);
    }
  }
  var chunks = _chunk_(text, CHUNK_SIZE);
  if (chunks.length === 0) chunks = [''];
  var rows = [];
  for (var i = 0; i < chunks.length; i++) rows.push([key, i, chunks[i]]);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
}

function _chunk_(s, size) {
  var out = [];
  for (var i = 0; i < s.length; i += size) out.push(s.substr(i, size));
  return out;
}

/* ==================== CONCURRENCY-SAFE OPS ====================
 * ปัญหาเดิม: client โหลดข้อมูลทั้งก้อนตอน login แล้วเขียนทับทั้ง key เวลาบันทึก
 * → เมื่อหลายเครื่องขายพร้อมกัน เครื่องที่เซฟทีหลังจะ "ทับ" บิลของเครื่องอื่น (บิลหาย),
 *   เลขบิลซ้ำ, และสต็อกเพี้ยน
 * วิธีแก้: ทุกการเปลี่ยนแปลงที่สำคัญให้ทำที่ server ภายใน LockService เดียว โดย
 *   อ่านค่า "ล่าสุด" จากชีตก่อนเสมอ แล้วค่อย merge/append/หัก-delta (ไม่ใช้ snapshot เก่าจาก client)
 */

// อ่านค่าปัจจุบันของ key เดียว (ประกอบ chunk แล้ว JSON.parse) — คืน null ถ้าไม่มี/พังพาร์ส
function _readKey_(sheet, key) {
  var last = sheet.getLastRow();
  if (last < 2) return null;
  var rows = sheet.getRange(2, 1, last - 1, 3).getValues();
  var chunks = [];
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][0] === key) chunks.push({ idx: Number(rows[i][1]) || 0, value: String(rows[i][2] == null ? '' : rows[i][2]) });
  }
  if (!chunks.length) return null;
  chunks.sort(function (a, b) { return a.idx - b.idx; });
  var joined = chunks.map(function (c) { return c.value; }).join('');
  if (joined === '') return null;
  try { return JSON.parse(joined); } catch (e) { return null; }
}

// โหลดเฉพาะบาง key (ใช้รีเฟรชระหว่างใช้งานให้เห็นข้อมูลของเครื่องอื่น โดยไม่ต้องโหลดใหม่ทั้งหมด)
function loadKeys(token, keys) {
  _requireSession_(token);
  var sheet = _sheet_(STORE_SHEET);
  var out = {};
  if (!keys || !keys.length) return { ok: true, data: out };
  for (var i = 0; i < keys.length; i++) out[keys[i]] = _readKey_(sheet, keys[i]);
  return { ok: true, data: out };
}

var PROD_KEY  = 'pos_products';   // ต้องตรงกับ K.PROD ฝั่ง client
var SALES_KEY = 'pos_sales';      // ต้องตรงกับ K.SALES ฝั่ง client

function _pad5_(n) { n = String(n); while (n.length < 5) n = '0' + n; return n; }

// upsert เรคคอร์ดตาม id: มีอยู่แล้ว → แทนที่, ยังไม่มี → ใส่หน้าสุด (ล่าสุดอยู่บน)
function _upsert_(sheet, key, records) {
  var arr = _readKey_(sheet, key);
  if (!(arr instanceof Array)) arr = [];
  var pos = {};
  for (var i = 0; i < arr.length; i++) if (arr[i] && arr[i].id != null) pos[arr[i].id] = i;
  for (var j = 0; j < records.length; j++) {
    var rec = records[j];
    if (rec && rec.id != null && pos[rec.id] !== undefined) {
      arr[pos[rec.id]] = rec;
    } else {
      arr.unshift(rec);
      pos = {};
      for (var k = 0; k < arr.length; k++) if (arr[k] && arr[k].id != null) pos[arr[k].id] = k;
    }
  }
  _writeKey_(sheet, key, JSON.stringify(arr));
  return arr.length;
}

// ปรับสต็อกแบบ delta (หัก/คืน) บนค่าล่าสุด — ข้ามสินค้าสต็อกไม่จำกัด
function _applyStock_(sheet, deltas) {
  if (!deltas) return;
  var products = _readKey_(sheet, PROD_KEY);
  if (!(products instanceof Array)) return;
  var changed = false;
  for (var i = 0; i < products.length; i++) {
    var p = products[i];
    if (p && p.id != null && deltas[p.id] !== undefined && !p.unlimited) {
      p.stock = (Number(p.stock) || 0) + Number(deltas[p.id]);
      changed = true;
    }
  }
  if (changed) _writeKey_(sheet, PROD_KEY, JSON.stringify(products));
}

// บันทึกการขาย 1 บิลแบบ atomic: จองเลขบิล + เพิ่มบิล + ตัดสต็อก ในล็อกเดียว
// server เป็นผู้ออกเลขบิล (กันเลขซ้ำ) แล้วส่งกลับให้ client
function commitSale(token, sale, stockDeltas, seqKey) {
  _requireSession_(token);
  if (!sale) return { ok: false, error: 'ไม่มีข้อมูลบิล' };
  seqKey = seqKey || 'pos_seq';
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = _sheet_(STORE_SHEET);
    var seq = (Number(_readKey_(sheet, seqKey)) || 0) + 1;   // เลขบิลจากค่าล่าสุด
    sale.no = 'INV-' + _pad5_(seq);
    _upsert_(sheet, SALES_KEY, [sale]);                      // เพิ่มบิล (ไม่ทับของเครื่องอื่น)
    _applyStock_(sheet, stockDeltas);                        // ตัดสต็อกแบบ delta
    _writeKey_(sheet, seqKey, JSON.stringify(seq));          // เดินเลขรันนิ่ง
    SpreadsheetApp.flush();
    return { ok: true, no: sale.no, seq: seq };
  } finally {
    lock.releaseLock();
  }
}

// ยกเลิก/คืนเงินบิลแบบ atomic + idempotent: ถ้าถูกยกเลิกไปแล้วจะไม่คืนสต็อกซ้ำ
function commitVoid(token, saleId, stockDeltas, refundedAt) {
  _requireSession_(token);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = _sheet_(STORE_SHEET);
    var arr = _readKey_(sheet, SALES_KEY);
    if (!(arr instanceof Array)) return { ok: false, error: 'ไม่พบข้อมูลการขาย' };
    var idx = -1;
    for (var i = 0; i < arr.length; i++) { if (arr[i] && arr[i].id === saleId) { idx = i; break; } }
    if (idx < 0) return { ok: false, error: 'ไม่พบบิลนี้ (อาจถูกลบไปแล้ว)' };
    if (arr[idx].status === 'refunded') return { ok: true, already: true };  // ยกเลิกแล้ว → ไม่คืนสต็อกซ้ำ
    arr[idx].status = 'refunded';
    arr[idx].refundedAt = refundedAt || '';
    _writeKey_(sheet, SALES_KEY, JSON.stringify(arr));
    _applyStock_(sheet, stockDeltas);                        // คืนสต็อกแบบ delta
    SpreadsheetApp.flush();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ==================== HELPERS ==================== */
function _sheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('ยังไม่ได้ตั้งค่า — โปรดรันฟังก์ชัน setup() ก่อน (ไม่พบชีต ' + name + ')');
  return sh;
}

function _session_(token) {
  if (!token) return null;
  var raw = CacheService.getScriptCache().get(CACHE_PREFIX + token);
  if (!raw) return null;
  // ต่ออายุ session อัตโนมัติเมื่อมีการใช้งาน
  CacheService.getScriptCache().put(CACHE_PREFIX + token, raw, SESSION_TTL);
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function _requireSession_(token) {
  var u = _session_(token);
  if (!u) throw new Error('SESSION_EXPIRED');
  return u;
}

function _sha256_(str) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8);
  var hex = '';
  for (var i = 0; i < bytes.length; i++) {
    var b = (bytes[i] < 0 ? bytes[i] + 256 : bytes[i]).toString(16);
    hex += (b.length === 1 ? '0' : '') + b;
  }
  return hex;
}
