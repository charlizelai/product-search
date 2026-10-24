// ============================================================
//  Product API v5.4 — 完整版
//  v4.x：CRUD + 圖片 + pHash + 快取 + CDN + OOS + 掃碼新增
//  v5.0：OOS 回報新增「訂單資訊」（單日期/單號尾4/派送時段/單類型）
//  v5.1：OOS 原因移至 _OOSReasons 獨立工作表
//  v5.2：新增 oosHistory API
//  v5.3：barcode 只抓真正 GTIN-13，不用 itemNum 冒充
//  v5.4：新增公眾假期（1823 API）+ 自訂繁忙日（_BusyDays）
//        + 冬至/母親節/父親節自動算 + prefillBusyDays()
// ============================================================

const SHEET_ID   = '1U3zVg2ljge-xxDHvAIRXX8KoDoPZBZq4ZRvH9GBWCww';
const SHEET_NAME = 'Sheet1';

// ── API Token ────────────────────────────────────────────────
// ⚠️ 安全性：舊版 token 曾經寫死在前端 index.html，等同已對外公開，
//    已於本次輪替作廢（請勿再沿用舊值）。
//    現在優先讀 ScriptProperties，讀不到才用下面的預設值。
//    請在 Apps Script 編輯器執行一次 setupTokens() 把 token 存進 ScriptProperties。
const DEFAULT_SECRET_TOKEN = 'user';
const DEFAULT_ADMIN_TOKEN  = 'admin';
const ADMIN_EMAIL  = 'charlizelai@gmail.com';

function getSecretToken_() {
  return PropertiesService.getScriptProperties().getProperty('SECRET_TOKEN') || DEFAULT_SECRET_TOKEN;
}
function getAdminToken_() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_TOKEN') || DEFAULT_ADMIN_TOKEN;
}

/**
 * 一次性設定 / 輪替 token。
 * 執行後請：
 *   1. 把印出的「管理員密碼」填到 index.html 的 ADMIN_TOKEN
 *   2. 通知同仁新的登入密碼
 */
function setupTokens(secret, admin) {
  var props = PropertiesService.getScriptProperties();
  var s = secret || DEFAULT_SECRET_TOKEN;
  var a = admin  || DEFAULT_ADMIN_TOKEN;
  props.setProperty('SECRET_TOKEN', s);
  props.setProperty('ADMIN_TOKEN', a);
  Logger.log('一般登入密碼：' + s);
  Logger.log('管理員密碼：' + a);
  return {ok: true, secret: s, admin: a};
}

const IMG_CELL_W = 100;
const IMG_CELL_H = 100;

const COL_MAP = {
  id: 1, barcode: 2, name: 3, name_zh: 4, price: 5, original_price: 6,
  image: 7, category_en: 8, category_zh: 9, description: 10,
  temp_location: 12, combine: 13, warehouse: 14, island: 15,
  aisle: 16, crosslane: 17, shelfnumber: 18, floor: 19,
  brand: 24,                                 // #21：惠康 JSON-LD 的 brand.name
  sku: 25                                    // #31：Y 欄 SKU code（人手維護，唯讀）
};

const IMAGE_COL_INDEX = 7;
const COL_TIMESTAMP   = 20;
const COL_SOURCE      = 21;
// ⚠️ 第 22 欄原本同時被 api.gs（phash）與 weeklySheetUpdate.gs（_removed_at）
//    使用，兩邊互相覆蓋。現已把 phash 移到第 23 欄，_removed_at 留在第 22 欄。
// ⚠️ 第 25 欄（Y 欄）= SKU code，由使用者手動維護，程式唔應該讀寫。
//    新增欄位請由第 26 欄開始，唔好擴大 MAX_COL 去蓋過第 25 欄。
const COL_PHASH       = 23;
// #21：第 24 欄 = brand（只在最後追加，不動既有欄位配置）
// #31/#36：提到 25 係為了讀寫第 25 欄（SKU）。
//      SKU 由前端編輯表單維護（FIELDS 已包含 sku），所以 updateProductInternal 會寫入。
//      ⚠️ 前端有防護：若 list 回傳嘅產品本身冇 sku 欄位（例如 API 未更新、
//         或 CDN 未重新匯出），編輯時會停用該欄並且唔送 sku，
//         避免用空白蓋走試算表原有資料。
//      ⚠️ 第 25 欄必須係純文字格式（見 ensureSkuTextFormat_），否則 0 開頭會消失。
//      ⚠️ weeklySheetUpdate.gs 只寫到第 24 欄，所以爬蟲唔會覆蓋 SKU。
const MAX_COL         = 25;

const ALL_FIELDS = Object.keys(COL_MAP);

const CACHE_KEY = 'product_list_v1';
const CACHE_TTL = 300;
const CACHE_CHUNK_SIZE = 90000;

const GITHUB_DATA_PATH = 'data/products.json';
const GITHUB_BRANCH = 'main';
const CONFIG_SHEET = '_Config';
const OOS_SHEET    = '_OOS';
const OOS_REASONS_SHEET = '_OOSReasons';
const BUSY_DAYS_SHEET   = '_BusyDays';
const USERS_SHEET       = '_Users';        // #38：個人帳戶
const SESSIONS_SHEET    = '_Sessions';     // #38：登入 session
const PW_ITERATIONS     = 1000;            // #38：密碼雜湊疊代次數
const SESSION_DAYS_REMEMBER = 30;          // 記住我
const SESSION_HOURS_DEFAULT = 12;          // 否則 12 小時

const SOURCE_SCAN_AUTO = 'scan_auto';

const API_VERSION = 'v5.4-2026-10-06';


// ============================================================
//  工具
// ============================================================
function getSheet() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 所有「會寫入」的操作都必須包在這裡。
 * 原因：findRow() 回傳的是列號；若期間另一個請求 deleteRow，
 *      後面的列會整批上移，接著的 setValue 就會寫到錯的產品上。
 * 注意：LockService 不支援重入，不要再於已持鎖的函式內呼叫 withLock。
 */
function withLock(fn) {
  var lock = LockService.getScriptLock();
  var ok = false;
  try { ok = lock.tryLock(25000); } catch (e) { ok = false; }
  if (!ok) return {error: '系統忙碌中，請稍後再試（Busy）'};
  try {
    return fn();
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * #38：把 token 解析成身分。
 *   1. 舊共用密碼（個人帳戶未完全接手之前仍然有效，唔會鎖死任何人）
 *   2. Session token（個人帳戶登入後發出）
 * 失敗回 null。
 */
function resolveAuth_(token) {
  var t = safeStr(token).trim();
  if (!t) return null;

  if (t === getAdminToken_()) {
    return { userId: 'legacy-admin', name: 'admin', display: '管理員', role: 'admin', legacy: true };
  }
  if (t === getSecretToken_()) {
    return { userId: 'legacy-staff', name: 'staff', display: '同事', role: 'staff', legacy: true };
  }
  return lookupSession_(t);
}

function isAuthorized(token) {
  return !!resolveAuth_(token);
}

function isAdminToken(token) {
  var a = resolveAuth_(token);
  return !!a && a.role === 'admin';
}

/** 取得呼叫者身分（記錄「邊個做」時用） */
function currentUser_(token) {
  return resolveAuth_(token) || { userId: '', name: '', display: '', role: 'staff' };
}

// ============================================================
//  #38  個人帳戶 + Session
// ============================================================

function getUsersSheet_(createIfMissing) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(USERS_SHEET);
  if (!sh && createIfMissing) {
    sh = ss.insertSheet(USERS_SHEET);
    sh.getRange(1, 1, 1, 9).setValues([[
      'user_id','name','display_name','role','pw_hash','pw_salt','active','created_at','last_login'
    ]]);
    sh.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#e8f0fe');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 80);
    sh.setColumnWidth(2, 120);
    sh.setColumnWidth(3, 140);
    sh.setColumnWidth(4, 70);
    sh.setColumnWidth(5, 250);
    sh.setColumnWidth(6, 150);
    sh.setColumnWidth(7, 70);
    sh.setColumnWidth(8, 170);
    sh.setColumnWidth(9, 170);
  }
  return sh;
}

function getSessionsSheet_(createIfMissing) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SESSIONS_SHEET);
  if (!sh && createIfMissing) {
    sh = ss.insertSheet(SESSIONS_SHEET);
    sh.getRange(1, 1, 1, 6).setValues([[
      'token','user_id','name','role','issued_at','expires_at'
    ]]);
    sh.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#e8f0fe');
    sh.setFrozenRows(1);
    try { sh.hideSheet(); } catch (e) {}
  }
  return sh;
}

function bytesToHex_(bytes) {
  var hex = '';
  for (var i = 0; i < bytes.length; i++) {
    var b = (bytes[i] + 256) % 256;
    hex += (b < 16 ? '0' : '') + b.toString(16);
  }
  return hex;
}

function makeSalt_() {
  return Utilities.getUuid().replace(/-/g, '').substring(0, 16);
}

/**
 * 密碼雜湊：每個用戶獨立 salt + SHA-256 疊代。
 * Apps Script 冇 bcrypt／PBKDF2，呢個係可行範圍內最好嘅做法。
 */
function hashPassword_(password, salt) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(salt) + '|' + String(password),
    Utilities.Charset.UTF_8
  );
  for (var i = 0; i < PW_ITERATIONS; i++) {
    bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes);
  }
  return bytesToHex_(bytes);
}

function readUsers_() {
  var sh = getUsersSheet_(false);
  if (!sh) return [];
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var vals = sh.getRange(2, 1, lastRow - 1, 9).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    var name = safeStr(r[1]).trim();
    if (!name) continue;
    var actRaw = r[6];
    out.push({
      row: i + 2,
      user_id: safeStr(r[0]).trim(),
      name: name,
      display: safeStr(r[2]).trim(),
      role: safeStr(r[3]).trim() || 'staff',
      pw_hash: safeStr(r[4]).trim(),
      pw_salt: safeStr(r[5]).trim(),
      active: !(actRaw === false || safeStr(actRaw).trim().toUpperCase() === 'FALSE'),
      created_at: safeStr(r[7]),
      last_login: safeStr(r[8])
    });
  }
  return out;
}

function findUser_(name) {
  var want = safeStr(name).trim().toLowerCase();
  if (!want) return null;
  var users = readUsers_();
  for (var i = 0; i < users.length; i++) {
    if (users[i].name.toLowerCase() === want) return users[i];
  }
  return null;
}

function lookupSession_(token) {
  var t = safeStr(token).trim();
  if (!t) return null;

  var cache = CacheService.getScriptCache();
  var hit = cache.get('sess_' + t);
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }

  var sh = getSessionsSheet_(false);
  if (!sh) return null;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  var vals = sh.getRange(2, 1, lastRow - 1, 6).getValues();
  var now = Date.now();
  for (var i = 0; i < vals.length; i++) {
    if (safeStr(vals[i][0]).trim() !== t) continue;
    var exp = Date.parse(safeStr(vals[i][5])) || 0;
    if (exp && exp < now) return null;
    var rec = {
      userId: safeStr(vals[i][1]).trim(),
      name:   safeStr(vals[i][2]).trim(),
      display: safeStr(vals[i][2]).trim(),
      role:   safeStr(vals[i][3]).trim() || 'staff',
      expires: exp
    };
    try { cache.put('sess_' + t, JSON.stringify(rec), 300); } catch (e) {}
    return rec;
  }
  return null;
}

function createSession_(user, remember) {
  var sh = getSessionsSheet_(true);
  var token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  var now = new Date();
  var ms = remember
    ? SESSION_DAYS_REMEMBER * 24 * 3600 * 1000
    : SESSION_HOURS_DEFAULT * 3600 * 1000;
  var exp = new Date(now.getTime() + ms);
  sh.appendRow([token, user.user_id, user.name, user.role, now.toISOString(), exp.toISOString()]);
  return { token: token, expires: exp.toISOString() };
}

function dropSession_(token) {
  var t = safeStr(token).trim();
  if (!t) return false;
  try { CacheService.getScriptCache().remove('sess_' + t); } catch (e) {}
  var sh = getSessionsSheet_(false);
  if (!sh) return false;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return false;
  var vals = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = vals.length - 1; i >= 0; i--) {
    if (safeStr(vals[i][0]).trim() === t) { sh.deleteRow(i + 2); return true; }
  }
  return false;
}

function revokeUserSessions_(name) {
  var sh = getSessionsSheet_(false);
  if (!sh) return 0;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;
  var vals = sh.getRange(2, 1, lastRow - 1, 6).getValues();
  var rows = [];
  for (var i = 0; i < vals.length; i++) {
    if (safeStr(vals[i][2]).trim() === name) {
      try { CacheService.getScriptCache().remove('sess_' + safeStr(vals[i][0]).trim()); } catch (e) {}
      rows.push(i + 2);
    }
  }
  for (var k = rows.length - 1; k >= 0; k--) { sh.deleteRow(rows[k]); }
  return rows.length;
}

function pruneSessions_() {
  var sh = getSessionsSheet_(false);
  if (!sh) return;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  var vals = sh.getRange(2, 1, lastRow - 1, 6).getValues();
  var now = Date.now();
  var dead = [];
  for (var i = 0; i < vals.length; i++) {
    var exp = Date.parse(safeStr(vals[i][5])) || 0;
    if (!exp || exp < now) dead.push(i + 2);
  }
  for (var k = dead.length - 1; k >= 0; k--) { sh.deleteRow(dead[k]); }
}

function getInviteCode_() {
  var props = PropertiesService.getScriptProperties();
  var c = props.getProperty('INVITE_CODE');
  if (!c) {
    c = 'WELLCOME' + Math.floor(Math.random() * 9000 + 1000);
    props.setProperty('INVITE_CODE', c);
  }
  return c;
}

/** 設定 / 查看邀請碼（可喺編輯器手動執行） */
function setInviteCode(code) {
  var c = safeStr(code).trim() || ('WELLCOME' + Math.floor(Math.random() * 9000 + 1000));
  PropertiesService.getScriptProperties().setProperty('INVITE_CODE', c);
  Logger.log('邀請碼：' + c);
  return c;
}

/* ---------- 公開動作：註冊 / 登入 / 登出 ---------- */

function registerUser(body) {
  if (!isObj(body)) return { error: 'Invalid body' };
  var name     = safeStr(body.name).trim();
  var display  = safeStr(body.displayName || body.display).trim();
  var password = safeStr(body.password);
  var invite   = safeStr(body.inviteCode).trim();

  if (!name) return { error: '請輸入名稱' };
  if (name.length < 2) return { error: '名稱至少 2 個字' };
  if (password.length < 6) return { error: '密碼至少 6 位' };
  if (invite !== getInviteCode_()) return { error: '邀請碼不正確' };
  if (findUser_(name)) return { error: '名稱「' + name + '」已經有人用' };

  var sh = getUsersSheet_(true);
  var users = readUsers_();
  var maxN = 0;
  for (var i = 0; i < users.length; i++) {
    var m = users[i].user_id.match(/^U([0-9]+)$/);
    if (m) maxN = Math.max(maxN, parseInt(m[1], 10));
  }
  var userId = 'U' + ('000' + (maxN + 1)).slice(-4);
  var salt = makeSalt_();
  var hash = hashPassword_(password, salt);
  var now  = new Date().toISOString();

  sh.appendRow([userId, name, display || name, 'staff', hash, salt, true, now, '']);

  var session = createSession_({ user_id: userId, name: name, role: 'staff' }, true);
  return { ok: true, token: session.token, expires: session.expires,
           userId: userId, name: name, display: display || name, role: 'staff' };
}

function loginUser(body) {
  if (!isObj(body)) return { error: 'Invalid body' };
  var name     = safeStr(body.name).trim();
  var password = safeStr(body.password);
  if (!name || !password) return { error: '請輸入名稱同密碼' };

  var u = findUser_(name);
  if (!u) return { error: '名稱或密碼不正確' };
  if (!u.active) return { error: '此帳戶已停用，請聯絡管理員' };

  var hash = hashPassword_(password, u.pw_salt);
  if (hash !== u.pw_hash) return { error: '名稱或密碼不正確' };

  var sh = getUsersSheet_(false);
  if (sh) sh.getRange(u.row, 9).setValue(new Date().toISOString());

  pruneSessions_();
  var session = createSession_({ user_id: u.user_id, name: u.name, role: u.role }, !!body.remember);
  return { ok: true, token: session.token, expires: session.expires,
           userId: u.user_id, name: u.name, display: u.display || u.name, role: u.role };
}

function logoutUser(body) {
  dropSession_(safeStr(body && body.token).trim());
  return { ok: true };
}

/* ---------- 管理員：員工管理 ---------- */

function listUsers() {
  var users = readUsers_();
  return {
    ok: true,
    inviteCode: getInviteCode_(),
    users: users.map(function(u) {
      return { userId: u.user_id, name: u.name, display: u.display, role: u.role,
               active: u.active, createdAt: u.created_at, lastLogin: u.last_login };
    })
  };
}

function setUserActive(body) {
  var name = safeStr(body && body.name).trim();
  var active = !!(body && body.active);
  var u = findUser_(name);
  if (!u) return { error: '找不到使用者：' + name };
  if (u.role === 'admin' && !active) return { error: '唔可以停用管理員帳戶' };
  getUsersSheet_(false).getRange(u.row, 7).setValue(active);
  if (!active) revokeUserSessions_(u.name);
  return { ok: true, name: name, active: active };
}

function setUserRole(body) {
  var name = safeStr(body && body.name).trim();
  var role = safeStr(body && body.role).trim() === 'admin' ? 'admin' : 'staff';
  var u = findUser_(name);
  if (!u) return { error: '找不到使用者：' + name };
  if (u.role === 'admin' && role !== 'admin' && u.name === currentUser_(body && body.token).name) {
    return { error: '唔可以取消自己嘅管理員權限' };
  }
  getUsersSheet_(false).getRange(u.row, 4).setValue(role);
  revokeUserSessions_(u.name);
  return { ok: true, name: name, role: role };
}

function resetUserPassword(body) {
  var name = safeStr(body && body.name).trim();
  var pw   = safeStr(body && body.password);
  if (pw.length < 6) return { error: '密碼至少 6 位' };
  var u = findUser_(name);
  if (!u) return { error: '找不到使用者：' + name };
  var salt = makeSalt_();
  var hash = hashPassword_(pw, salt);
  var sh = getUsersSheet_(false);
  sh.getRange(u.row, 5).setValue(hash);
  sh.getRange(u.row, 6).setValue(salt);
  revokeUserSessions_(u.name);
  return { ok: true, name: name };
}

/** 手動將某人設為管理員（第一次 bootstrap 用；亦可由 API 呼叫） */
function promoteToAdmin(name) {
  var u = findUser_(name);
  if (!u) { Logger.log('找不到使用者：' + name); return null; }
  getUsersSheet_(false).getRange(u.row, 4).setValue('admin');
  Logger.log(u.name + ' 已設為管理員');
  return u.name;
}


function isValidRow(r) {
  return r && Array.isArray(r) && r.length > 0;
}

function isObj(o) {
  return o && typeof o === 'object' && !Array.isArray(o);
}

function safeStr(v) {
  return v != null ? String(v) : '';
}

function isTempId(id) {
  return /^\d{11}$/.test(safeStr(id).trim());
}

function pad2(n) { return n < 10 ? '0' + n : '' + n; }

function ensureColumns() {
  const sheet = getSheet();
  const headers = {
    12: 'temp_location', 13: 'combine', 14: 'warehouse', 15: 'island',
    16: 'aisle', 17: 'crosslane', 18: 'shelfnumber', 19: 'floor',
    20: 'last_modified', 21: 'last_source', 23: 'phash', 24: 'brand', 25: 'sku'
  };
  Object.keys(headers).forEach(function(col) {
    var c = parseInt(col);
    if (!sheet.getRange(1, c).getValue()) sheet.getRange(1, c).setValue(headers[col]);
  });

  ensureSkuTextFormat_(sheet);
  ensureBarcodeTextFormat_(sheet);
}

/**
 * #42：第 2 欄（Barcode）亦必須係純文字。
 *
 * 條碼可能 0 開頭（美加 UPC-A 常見），試算表預設會當成數字，
 * 前導 0 會直接消失 —— 同 SKU、單號尾4 一模一樣嘅問題。
 *
 * ⚠️ 只改格式，唔會改動任何已有嘅值；已經冇咗前導 0 嘅舊值救唔返。
 */
function ensureBarcodeTextFormat_(sheet) {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('BARCODE_TEXT_FMT') === '1') return;
  var rows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, COL_MAP.barcode, rows, 1).setNumberFormat('@');
  props.setProperty('BARCODE_TEXT_FMT', '1');
}

/** 手動重新套用 B 欄（Barcode）純文字格式 */
function forceBarcodeTextFormat() {
  PropertiesService.getScriptProperties().deleteProperty('BARCODE_TEXT_FMT');
  ensureBarcodeTextFormat_(getSheet());
  Logger.log('已將第 ' + COL_MAP.barcode + ' 欄（B 欄 Barcode）設為純文字格式');
  return true;
}

/**
 * #34：第 25 欄（Y 欄）SKU code 必須以「純文字」儲存。
 *
 * SKU 係純數字，但可能 0 開頭（例如 "0123"）。試算表預設會當成數字 123，
 * 前導 0 會直接消失 —— 同條碼、單號尾4 一樣嘅問題。
 *
 * 只設定一次（記在 ScriptProperties），避免每次請求都改格式。
 */
function ensureSkuTextFormat_(sheet) {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('SKU_TEXT_FMT') === '1') return;
  var rows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, COL_MAP.sku, rows, 1).setNumberFormat('@');
  props.setProperty('SKU_TEXT_FMT', '1');
}

/**
 * 手動重新套用 Y 欄純文字格式。
 * 適用情況：你插過新行、或者改動過格式、或者想確認一次。
 * 只改格式，唔會改任何儲存格嘅內容。
 */
function forceSkuTextFormat() {
  PropertiesService.getScriptProperties().deleteProperty('SKU_TEXT_FMT');
  var sheet = getSheet();
  ensureSkuTextFormat_(sheet);
  Logger.log('已將第 ' + COL_MAP.sku + ' 欄（Y 欄）設為純文字格式');
  return true;
}

/**
 * 一次性遷移（只需執行一次，可在 Apps Script 編輯器手動跑）：
 *   舊版 api.gs 把 phash 寫在第 22 欄，weeklySheetUpdate.gs 也把 _removed_at
 *   寫在第 22 欄，兩邊互相覆蓋，該欄資料已混雜。
 *   本函式把第 22 欄中「看起來像 phash」的值（16 位十六進位）搬到第 23 欄，
 *   並清空第 22 欄，之後由 weeklySheetUpdate.gs 重新建立 _removed_at。
 */
function migratePhashColumn() {
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {ok: true, moved: 0};
  var col = sheet.getRange(2, 22, lastRow - 1, 1).getValues();
  var out = [];
  var moved = 0;
  for (var i = 0; i < col.length; i++) {
    var v = safeStr(col[i][0]).trim();
    if (/^[0-9a-fA-F]{16}$/.test(v)) { out.push([v]); moved++; }
    else { out.push(['']); }
  }
  sheet.getRange(2, 23, out.length, 1).setValues(out);
  sheet.getRange(2, 22, col.length, 1).clearContent();
  sheet.getRange(1, 23).setValue('phash');
  Logger.log('migratePhashColumn: moved=' + moved);
  return {ok: true, moved: moved};
}


/**
 * #45 一次性修復：統一分類名（第 8 欄 category_en）嘅大細楷／空格分裂。
 *
 * 成因：同一個惠康分類，喺唔同時期、唔同頁面被抓到唔同寫法，例如
 *       Case Deal / case deal / CASE DEAL / case Deal / CASe deal
 *       Beverages / Beverage / beverages / BEVERAGES
 *       結果：前端「分類」下拉出現 36 個選項（實際只有 21 個真分類），
 *             而且揀「Beverages」只會出 1,466 件，另外 11 件完全消失。
 *
 * 做法：只做正規化（統一大小楷 + 壓縮空格），唔會改變任何一件貨嘅意思。
 *       認唔到嘅寫法原封不動，並喺 Logger 列出，等你決定要唔要合併。
 *
 * 執行：Apps Script 編輯器 → 函式揀 fixCategoryNames → 執行一次。
 *       （手動函式，唔經 /exec，所以唔需要重新部署）
 */
var FIX_CAT_CANONICAL = [
  'Beauty & Personal Care', 'Beer, Wines & Spirits', 'Beverages', 'Breakfast & Bakery',
  'Canned, Preserved & Soup', 'Case Deal', 'Cat', 'Condiment & Sauce',
  'Dairy, Chilled & Eggs', 'Dog', 'Feminine Care', 'Festive items', 'Fitness & Wellness',
  'Frozen Food', 'Fruits & Vegetables', 'Health Care', 'Household Supplies',
  'Kitchen & Dining', 'Meat & Seafood', 'Milk Powder', 'Mother & Baby',
  'Outdoor & Sundries', 'Rice, Oil & Noodles', 'SIM Card', 'Snacks & Confectionery',
  'UHT Milk'
];

// 明顯同義嘅別名（唔係大細楷問題，係寫法唔同）。冇十足信心嘅唔好放呢度。
var FIX_CAT_ALIAS = {
  'Beverage': 'Beverages'          // 單數寫法，8 件
};

function fixCatKey_(s) {
  return safeStr(s)
    .replace(/\s+/g, ' ')
    .replace(/\s*&\s*/g, ' & ')
    .trim()
    .toLowerCase();
}

/**
 * #51 用條碼移除重複記錄。
 *
 * ⚠️ 呢個係【刪除資料】嘅操作，所以設計上極度保守：
 *   1. 只會處理「同一條碼 + 同一個 ID」嘅行 —— 唔同 ID 撞條碼一律唔碰，另外列出。
 *   2. 先跑 previewDuplicateBarcodes()（只讀，唔會改任何嘢）。
 *   3. 執行前會把每一行被刪嘅資料完整備份到 _DeletedDuplicates 工作表（可以還原）。
 *   4. 公式欄（例如 K 欄嘅圖片縮圖公式）會保留：重寫完之後用第 2 行嘅公式
 *      copyTo 落新範圍，Sheets 會自動調整相對參照。
 *   4. 合併規則：先用保留行嘅值；保留行為空就用另一行嘅值（唔會漏資料）。
 *      品牌衝突 → 用品名對得返嘅嗰個（實測 24/24 組都對得返）。
 *      價格衝突 → 用較新嘅記錄（惠康每星期更新價格）。
 *      分類衝突 → 判斷唔到，整組跳過並列出，等你決定。
 */
var DUP_BACKUP_SHEET = '_DeletedDuplicates';
var DUP_REPORT_SHEET = '_DuplicateReport';

function previewDuplicateBarcodes() { return dedupeByBarcode_(true); }
function removeDuplicateBarcodes()  { return dedupeByBarcode_(false); }

/**
 * #51 診斷：純檢查，唔會改任何嘢。用嚟查「點解去重跑唔到」。
 * 執行後睇「執行記錄」，會直接講結論。
 */
function checkDedupeBlocker() {
  var out = [];
  try {
    var sheet = getSheet();
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var n = lastRow - 1;
    var cols = Math.max(MAX_COL, lastCol || MAX_COL);
    Logger.log('① 主表：lastRow=' + lastRow + '　資料列=' + n +
               '　getLastColumn=' + lastCol + '　重寫用 cols=' + cols);

    var lock = LockService.getScriptLock();
    var gotLock = lock.tryLock(5000);
    if (gotLock) lock.releaseLock();
    Logger.log('② 可以即刻取得腳本鎖：' + gotLock + (gotLock ? '' : '（有程式跑緊 → 去重會拒絕執行）'));

    var fmls = sheet.getRange(2, 1, n, cols).getFormulas();
    var cells = [];
    for (var i = 0; i < fmls.length && cells.length < 20; i++) {
      for (var j = 0; j < fmls[i].length && cells.length < 20; j++) {
        if (fmls[i][j]) {
          cells.push(sheet.getRange(i + 2, j + 1).getA1Notation() + '  =  ' + fmls[i][j]);
        }
      }
    }
    Logger.log('③ 有公式嘅格：' + (cells.length ? ('\n　　' + cells.join('\n　　')) : '冇'));

    Logger.log('④ 結論：' + (cells.length
      ? '表內有公式（正常，例如圖片縮圖欄）→ 新版會【保留公式】自動處理，唔會再擋住去重。'
      : '冇公式 → 去重直接執行。'));

    out.push('lastRow=' + lastRow, 'lastCol=' + lastCol, 'lock=' + gotLock,
             'formulaCells=' + cells.length);
    Logger.log('⑤ 摘要：' + out.join('　'));
    return { ok: true, lastRow: lastRow, lastCol: lastCol, lock: gotLock,
             formulaCells: cells };
  } catch (e) {
    Logger.log('✗ 診斷都出錯：' + e.message);
    return { error: e.message };
  }
}

function dedupeByBarcode_(dryRun) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(60000)) {
    Logger.log('#51 ✗ 取不到鎖（有其他程式正在執行），乜都冇做');
    return { error: '系統忙碌中，請稍後再試' };
  }
  try {
    var sheet = getSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 3) {
      Logger.log('#51 資料太少，唔需要處理');
      return { ok: true, message: '資料太少，唔需要處理' };
    }

    var n = lastRow - 1;
    var cols = Math.max(MAX_COL, sheet.getLastColumn() || MAX_COL);
    var vals = sheet.getRange(2, 1, n, cols).getValues();

    // ---- 1) 按條碼分組 ----
    var byBarcode = {};   // 條碼 -> [陣列索引]
    for (var i = 0; i < n; i++) {
      var b = safeStr(vals[i][COL_MAP.barcode - 1]).trim();
      if (!b) continue;
      if (!byBarcode[b]) byBarcode[b] = [];
      byBarcode[b].push(i);
    }

    var groups = [];          // { barcode, rows:[i], ids:[], needsReview, reason }
    for (var b2 in byBarcode) {
      var idxs = byBarcode[b2];
      if (idxs.length < 2) continue;

      var ids = {};
      idxs.forEach(function (k) { ids[safeStr(vals[k][COL_MAP.id - 1]).trim()] = true; });
      var idList = Object.keys(ids);

      if (idList.length > 1) {
        groups.push({ barcode: b2, rows: idxs, ids: idList, needsReview: true,
                      reason: '同一個條碼有 ' + idList.length + ' 個唔同 ID，唔敢自動處理' });
      } else {
        groups.push({ barcode: b2, rows: idxs, ids: idList, needsReview: false, reason: '' });
      }
    }

    // ---- 2) 為每組決定保留行 + 合併值 ----
    var updates = [];     // { rowIndex, row }
    var deletes = [];     // { rowIndex, backup }
    var review = [];      // 需要人手睇
    var mergedCount = 0;

    groups.forEach(function (g) {
      if (g.needsReview) { review.push(g); return; }

      var rows = g.rows.map(function (k) { return { i: k, data: vals[k] }; });

      // 保留行：非空欄位最多；同分就用較新
      rows.forEach(function (r) {
        r.filled = 0;
        for (var f = 0; f < cols; f++) if (safeStr(r.data[f]).trim() !== '') r.filled++;
        r.t = safeStr(r.data[COL_TIMESTAMP - 1]);
      });
      rows.sort(function (a, b) {
        if (b.filled !== a.filled) return b.filled - a.filled;
        return a.t < b.t ? 1 : (a.t > b.t ? -1 : 0);
      });

      var keep = rows[0];
      var others = rows.slice(1);
      var merged = keep.data.slice();          // 用保留行做底
      var catConflict = false;

      for (var ci = 1; ci <= cols; ci++) {
        var fk = null;
        for (var kk in COL_MAP) { if (COL_MAP[kk] === ci) { fk = kk; break; } }
        if (!fk) continue;

        var curV = safeStr(merged[ci - 1]);
        var vals2 = [curV].concat(others.map(function (o) { return safeStr(o.data[ci - 1]); }));
        var uniq = {};
        vals2.forEach(function (x) { var t2 = String(x).trim(); if (t2) uniq[t2] = true; });
        var keys = Object.keys(uniq);
        if (keys.length === 0) continue;

        if (!curV.trim()) {                     // 保留行空白 → 借另一行嘅值
          for (var oi = 0; oi < others.length; oi++) {
            var ov = safeStr(others[oi].data[ci - 1]);
            if (ov.trim()) { merged[ci - 1] = ov; mergedCount++; break; }
          }
          continue;
        }
        if (keys.length === 1) continue;        // 冇衝突

        // ---- 有衝突 ----
        if (fk === 'brand') {
          var nmAll = safeStr(merged[COL_MAP.name_zh - 1]) + ' ' + safeStr(merged[COL_MAP.name - 1]);
          var lowAll = nmAll.toLowerCase();
          var matched = vals2.filter(function (x) {
            return x.trim() && lowAll.indexOf(String(x).trim().toLowerCase()) >= 0;
          });
          var uniqMatched = {};
          matched.forEach(function (x) { uniqMatched[String(x).trim()] = x; });
          var mk = Object.keys(uniqMatched);
          if (mk.length === 1) { merged[ci - 1] = uniqMatched[mk[0]]; mergedCount++; }
          // 對唔到就保留保留行原本嘅值
        } else if (fk === 'price' || fk === 'original_price') {
          var newest = keep;
          others.forEach(function (o) { if (safeStr(o.data[COL_TIMESTAMP - 1]) > safeStr(newest.data[COL_TIMESTAMP - 1])) newest = o; });
          var nv = safeStr(newest.data[ci - 1]);
          if (nv.trim() && nv !== merged[ci - 1]) { merged[ci - 1] = nv; mergedCount++; }
        }
        // 分類（category_en / category_zh）嘅衝突唔喺呢度處理，
        // 因為 H 同 I 係一對，要一齊揀，見下面。
      }

      // ---- 分類係一對（H = category_en、I = category_zh），要一齊決定 ----
      var catSeen = {};
      rows.forEach(function (r) {
        var e2 = safeStr(r.data[COL_MAP.category_en - 1]).trim();
        if (e2) catSeen[e2] = true;
      });
      if (Object.keys(catSeen).length > 1) {
        // 用戶規則：一律保留 Case Deal（原箱）—— 連埋佢自己嗰個中文細分，
        // 唔可以 H 用 Case Deal 但 I 用另一個分類嘅細分。
        var cdRow = null;
        for (var ri = 0; ri < rows.length; ri++) {
          if (/^case deal$/i.test(safeStr(rows[ri].data[COL_MAP.category_en - 1]).trim())) { cdRow = rows[ri]; break; }
        }
        if (cdRow) {
          var cdEn = safeStr(cdRow.data[COL_MAP.category_en - 1]);
          var cdZh = safeStr(cdRow.data[COL_MAP.category_zh - 1]);
          if (safeStr(merged[COL_MAP.category_en - 1]) !== cdEn) { merged[COL_MAP.category_en - 1] = cdEn; mergedCount++; }
          if (safeStr(merged[COL_MAP.category_zh - 1]) !== cdZh) { merged[COL_MAP.category_zh - 1] = cdZh; mergedCount++; }
        } else {
          g.needsReview = true;
          g.reason = '兩行嘅分類唔同（' + keys0(others, keep, COL_MAP.category_en) + '），而且冇 Case Deal，等你自己揀';
          review.push(g);
          return;
        }
      }

      var changed = false;
      for (var q = 0; q < cols; q++) {
        if (safeStr(merged[q]) !== safeStr(keep.data[q])) { changed = true; break; }
      }
      if (changed) {
        merged[COL_TIMESTAMP - 1] = new Date().toISOString();
        merged[COL_SOURCE - 1] = 'user';
        updates.push({ rowIndex: keep.i, row: merged });
      }
      others.forEach(function (o) {
        deletes.push({ rowIndex: o.i, backup: o.data, keptId: safeStr(keep.data[COL_MAP.id - 1]) });
      });
    });

    // 小工具：列出分類值，畀 review 訊息用
    function keys0(os, kp, col) {
      return [safeStr(kp.data[col - 1])].concat(os.map(function (o) { return safeStr(o.data[col - 1]); }))
        .filter(function (x) { return String(x).trim(); }).join(' vs ');
    }

    var summary = {
      ok: true,
      dryRun: !!dryRun,
      duplicateGroups: groups.length,
      readyGroups: groups.length - review.length,
      reviewGroups: review.length,
      willUpdate: updates.length,
      willDelete: deletes.length,
      mergedFields: mergedCount
    };

    // ---- 3) 寫報告表 ----
    writeDupReport_(groups, review, summary);

    if (dryRun) {
      Logger.log('#51 預覽：' + JSON.stringify(summary));
      summary.message = '預覽完成：' + groups.length + ' 組重複，可以自動處理 ' +
        (groups.length - review.length) + ' 組（會刪 ' + deletes.length + ' 行），' +
        '需要人手睇 ' + review.length + ' 組。詳情見 _DuplicateReport 工作表。';
      return summary;
    }

    // ---- 4) 真正執行：先備份，再一次過重寫整個資料區 ----
    //
    // ⚠️ 唔可以逐行 deleteRow()。13,682 行嘅表上面每刪一行要幾百毫秒至一秒，
    //    180 行就要 3-6 分鐘，會直接撞爆 Apps Script 6 分鐘上限（上次就係咁死）。
    //    改為：計出保留邊啲行 → 一次 setValues 寫返上去 → 清走尾段多出嚟嘅行。
    if (deletes.length) {
      // ---- 4a) 處理公式欄（例如 K 欄嘅 =IF(G2="","",IMAGE(G2,4,100,100))）----
      //     唔可以當佢係錯：呢啲係圖片縮圖公式，用戶要保留。
      //     做法：公式欄唔寫值 → 重寫完之後，用第 2 行嘅公式 copyTo 落新範圍，
      //          Sheets 會自動調整相對參照（G2→G3→G4…），所以每行都對得返自己。
      var fmls = sheet.getRange(2, 1, n, cols).getFormulas();
      var formulaCols = {};        // 欄號 -> 第 2 行嘅公式
      var badCols = [];            // 第 2 行冇公式但下面有 → 冇辦法安全重排
      for (var fj2 = 0; fj2 < cols; fj2++) {
        if (!fmls[0][fj2]) {
          for (var fi2 = 1; fi2 < fmls.length; fi2++) {
            if (fmls[fi2][fj2]) { badCols.push(fj2 + 1); break; }
          }
          continue;
        }
        formulaCols[fj2 + 1] = fmls[0][fj2];
      }
      if (badCols.length) {
        var bmsg = '第 ' + badCols.join('、') + ' 欄有公式但唔係由第 2 行開始，' +
                   '冇辦法安全重排。請話我知，我另外處理。';
        Logger.log('#51 ✗ ' + bmsg);
        return { ok: false, error: bmsg };
      }
      Logger.log('#51 公式欄（會保留）：' + (Object.keys(formulaCols).join('、') || '冇'));

      // 有篩選嘅話，Google 會封鎖部分範圍操作（會出「含有已篩除列的範圍」）
      try {
        if (sheet.getFilter()) {
          Logger.log('#51 ⚠️ 主表有篩選（Filter）—— 建議先跑 removeSheetFilter()。今次照做。');
        }
      } catch (eFilt) {}

      try {
        backupDeletedRows_(deletes, cols);
      } catch (eBk) {
        var emsg = '備份失敗，為安全起見【冇刪任何嘢】。再跑一次通常就得：' + eBk.message;
        Logger.log('#51 ✗ ' + emsg);
        return { ok: false, error: emsg };
      }

      var deleteSet = {};
      deletes.forEach(function (dd) { deleteSet[dd.rowIndex] = true; });
      var updateMap = {};
      updates.forEach(function (u) { updateMap[u.rowIndex] = u.row; });

      var out = [];
      for (var ri2 = 0; ri2 < n; ri2++) {
        if (deleteSet[ri2]) continue;
        out.push(updateMap[ri2] ? updateMap[ri2] : vals[ri2]);
      }

      // 公式欄先寫空白，之後再補返公式（唔可以寫值，否則會壓死公式）
      var fCols = Object.keys(formulaCols).map(Number);
      if (fCols.length) {
        for (var oi2 = 0; oi2 < out.length; oi2++) {
          for (var fi3 = 0; fi3 < fCols.length; fi3++) out[oi2][fCols[fi3] - 1] = '';
        }
      }

      // 分塊寫入（每次 2000 行），避免單次 setValues 太大而失敗
      var CHUNK = 2000;
      for (var c0 = 0; c0 < out.length; c0 += CHUNK) {
        var part = out.slice(c0, c0 + CHUNK);
        sheet.getRange(2 + c0, 1, part.length, cols).setValues(part);
        Logger.log('#51 寫入進度 ' + Math.min(c0 + CHUNK, out.length) + '/' + out.length);
      }

      // 尾段多出嚟嘅行一次過刪走（deleteRows 係單一個 API 呼叫，比逐行 deleteRow 快好多）
      if (out.length < n) {
        var removed = n - out.length;
        try {
          sheet.deleteRows(2 + out.length, removed);
          Logger.log('#51 已刪尾段 ' + removed + ' 行');
        } catch (eDel) {
          Logger.log('#51 ⚠️ deleteRows 失敗（' + eDel.message + '），改用清空內容');
          sheet.getRange(2 + out.length, 1, removed, cols).clearContent();
        }
      }

      // ---- 4b) 補返公式欄 ----
      //   用第 2 行嘅公式，copyTo 落新範圍；Sheets 會自動調整相對參照。
      fCols.forEach(function (cnum) {
        var tpl = formulaCols[cnum];
        if (!tpl) return;
        try {
          var cell = sheet.getRange(2, cnum);
          cell.setFormula(tpl);
          if (out.length > 1) {
            cell.copyTo(sheet.getRange(2, cnum, out.length, 1));
          }
          Logger.log('#51 已補返第 ' + cnum + ' 欄公式（' + out.length + ' 行）');
        } catch (eF) {
          Logger.log('#51 ⚠️ 第 ' + cnum + ' 欄公式補唔返：' + eF.message);
        }
      });
    } else if (updates.length) {
      updates.forEach(function (u) {
        sheet.getRange(u.rowIndex + 2, 1, 1, cols).setValues([u.row]);
      });
    }

    invalidateCache();
    scheduleExport();

    Logger.log('#51 已執行：' + JSON.stringify(summary));
    summary.message = '完成：刪除 ' + deletes.length + ' 行重複，更新 ' + updates.length +
      ' 行。需要人手睇 ' + review.length + ' 組（見 _DuplicateReport）。備份喺 _DeletedDuplicates。';
    return summary;
  } finally {
    lock.releaseLock();
  }
}

/**
/**
 * #51 修復：重填第 11 欄（K 欄）嘅圖片縮圖公式。
 * 用喺：去重複／整表重寫之後，如果喺 Google Sheet 見到縮圖有窿就用呢個補返。
 *   ① 如果表有篩選，先跑 removeSheetFilter()
 *   ② 再跑 refillImageFormula()
 */
function refillImageFormula() {
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('冇資料'); return { ok: true, rows: 0 }; }

  var COL = 11;
  var tpl = '';
  try { tpl = sheet.getRange(2, COL).getFormula(); } catch (e) {}
  if (!tpl) tpl = '=IF(G2="","",IMAGE(G2,4,100,100))';
  Logger.log('模板公式（第 2 行）：' + tpl);

  // ⚠️ 唔用 copyTo —— 如果工作表有篩選，copyTo 會出
  //    「這項操作不支援含有已篩除列的範圍」。改為自己砌好每一行嘅公式，
  //    再用 setFormulas 一次過寫入（單一 API 呼叫，冇篩選限制）。
  var out = [];
  for (var r = 2; r <= lastRow; r++) {
    out.push([dupRebaseFormula_(tpl, 2, r)]);
  }

  try {
    sheet.getRange(2, COL, out.length, 1).setFormulas(out);
  } catch (e) {
    var msg = '填公式失敗：' + e.message +
              '（如果係「含有已篩除列的範圍」，請先跑 removeSheetFilter() 再試）';
    Logger.log('✗ ' + msg);
    return { ok: false, error: msg };
  }

  // 覆核：睇下有幾多格真係有公式
  var check = sheet.getRange(2, COL, out.length, 1).getFormulas();
  var filled = 0;
  for (var i = 0; i < check.length; i++) if (check[i][0]) filled++;

  Logger.log('✅ 已重填第 ' + COL + ' 欄公式：' + filled + ' / ' + out.length + ' 行有公式');
  if (filled < out.length) {
    Logger.log('⚠️ 有 ' + (out.length - filled) + ' 行填唔到。如果工作表有篩選，' +
               '請先跑 removeSheetFilter() 再跑一次。');
  }
  return { ok: true, rows: out.length, filled: filled, formula: tpl };
}

/**
 * 把「第 fromRow 行嘅公式」改寫成「第 toRow 行嘅版本」。
 * 例如模板 =IF(G2="","",IMAGE(G2,4,100,100))、toRow=5 → =IF(G5="","",IMAGE(G5,4,100,100))
 */
function dupRebaseFormula_(tpl, fromRow, toRow) {
  var s = String(tpl || '');
  if (!s || fromRow === toRow) return s;
  // 只換「欄字母 + 行號」而且行號後面唔係數字（避免 G2 撞 G20）
  var re = new RegExp('([A-Za-z]{1,3})' + fromRow + '(?![0-9])', 'g');
  return s.replace(re, function (m, col) { return col + toRow; });
}

/** 移除主表嘅篩選（Google Sheets 對有篩選嘅範圍會封鎖部分操作） */
function removeSheetFilter() {
  var sheet = getSheet();
  var f = null;
  try { f = sheet.getFilter(); } catch (e) {}
  if (!f) { Logger.log('主表冇篩選'); return { ok: true, had: false }; }
  f.remove();
  Logger.log('✅ 已移除主表篩選');
  return { ok: true, had: true };
}

/** * 把即將被刪嘅行完整備份（連原本行號），出錯可以還原。
 * ⚠️ 新建工作表之後即刻大批寫入，Google 有時會出「服務錯誤：試算表」，
 *    所以要：新建後等一等 + 分塊寫（每 40 行）+ 每塊最多重試 3 次。
 *    如果最終都寫唔入，會 throw —— 上層會中止，唔會刪任何嘢。
 */
function backupDeletedRows_(deletes, cols) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(DUP_BACKUP_SHEET);
  var created = false;

  if (!sh) {
    sh = ss.insertSheet(DUP_BACKUP_SHEET);
    dupWriteWithRetry_(function () {
      sh.getRange(1, 1, 1, 4).setValues([['備份時間', '原本行號', '保留嘅 ID', '完整資料 (JSON)']])
        .setFontWeight('bold').setBackground('#fdecea');
    }, '寫備份表頭');
    sh.setFrozenRows(1);
    created = true;
  }
  // 等 Google 後台追得上，再 flush 一次
  try { if (typeof SpreadsheetApp.flush === 'function') SpreadsheetApp.flush(); } catch (eF0) {}
  Utilities.sleep(created ? 1500 : 800);

  var now = new Date().toISOString();
  // ⚠️ 用【單欄 JSON】而唔係 28 欄：179 行由 5,012 格減到 716 格，
  //    大幅減少「服務錯誤：試算表」嘅機會。還原時 JSON.parse 就得。
  var out = deletes.map(function (dd) {
    return [now, dd.rowIndex + 2, dd.keptId, JSON.stringify(dd.backup)];
  });
  if (!out.length) return;

  var startRow = sh.getLastRow() + 1;
  var CHUNK = 20;

  for (var i = 0; i < out.length; i += CHUNK) {
    var part = out.slice(i, i + CHUNK);
    (function (off, chunk) {
      try {
        dupWriteWithRetry_(function () {
          sh.getRange(startRow + off, 1, chunk.length, 4).setValues(chunk);
        }, '寫備份第 ' + (off + 1) + ' 筆起');
      } catch (e) {
        throw new Error('備份寫入失敗（第 ' + (off + 1) + '～' + (off + chunk.length) + ' 筆）：' + e.message);
      }
    })(i, part);
    Utilities.sleep(250);
  }
  Logger.log('#51 已備份 ' + out.length + ' 行去 ' + DUP_BACKUP_SHEET + '（單欄 JSON 格式）');
}

/** 寫試算表嘅通用重試（「服務錯誤：試算表」通常係暫時性，等多一陣再試就得） */
function dupWriteWithRetry_(fn, label) {
  var last = null;
  for (var a = 1; a <= 3; a++) {
    try { return fn(); }
    catch (e) {
      last = e;
      Logger.log('#51 ' + (label || '寫入') + ' 重試 ' + a + '：' + e.message);
      if (a < 3) Utilities.sleep(1500 * a);
    }
  }
  throw last || new Error('未知錯誤');
}

/** 每次重寫一份報告，方便你喺 Google Sheet 篩 */
function writeDupReport_(groups, review, summary) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(DUP_REPORT_SHEET);
    // 唔用 deleteSheet + insertSheet（會增加 Google 服務錯誤風險），改為清空重用
    if (sh) {
      sh.clear();
    } else {
      sh = ss.insertSheet(DUP_REPORT_SHEET);
      sh.setFrozenRows(1);
      Utilities.sleep(800);
    }
    sh.getRange(1, 1, 1, 4).setValues([['條碼', 'ID', '狀態', '說明']])
      .setFontWeight('bold').setBackground('#e8f0fe');

    var rows = [];
    var reviewKeys = {};
    review.forEach(function (g) { reviewKeys[g.barcode] = g.reason; });

    groups.forEach(function (g) {
      var isRev = !!reviewKeys[g.barcode];
      rows.push([
        g.barcode,
        g.ids.join(' / '),
        isRev ? '⚠️ 要人手決定' : '✅ 會自動合併',
        isRev ? reviewKeys[g.barcode] : ('保留資料最完整嗰行，刪 ' + (g.rows.length - 1) + ' 行')
      ]);
    });
    if (rows.length) sh.getRange(2, 1, rows.length, 4).setValues(rows);
  } catch (e) {
    Logger.log('writeDupReport_ 失敗：' + e.message);
  }
}

/**
 * #52 用惠康官方資料覆核／修正「品牌」欄（第 24 欄）。
 *
 * 背景：表內有幾百行品牌寫錯（例如「每日鮮語」牛奶寫成「吉列」）。成因係惠康
 *       個網站本身嘅品牌資料亂，本機亦冇辦法判斷邊個才對 —— 唯一可靠來源係
 *       惠康產品頁嘅 JSON-LD："brand": {"@type":"Brand","name":"..."}。
 *
 * 做法：
 *   1. 逐件產品抓一次產品頁，讀 JSON-LD 品牌。
 *   2. 同表內比較，【只有唔同而且唔係空白】才寫入 —— 所以亂掃都唔會破壞資料。
 *   3. 進度記錄喺 _BrandCheck 工作表，下次跑會續落去，唔會重複抓。
 *   4. 優先處理「品牌對唔上品名」嘅可疑個案，之後才掃其餘。
 *
 * 執行：函式揀 checkBrandsFromWellcome → 執行（每跑一次約 120 件、4 分鐘）。
 */
var BRAND_CHECK_SHEET = '_BrandCheck';
var BRAND_CHECK_PER_RUN = 120;      // 每次跑幾多件（~1.5 秒/件，留返時間餘裕）
var BRAND_CHECK_TIME_MS = 5 * 60 * 1000;   // 5 分鐘就收手，避免爆 6 分鐘上限

function checkBrandsFromWellcome() { return brandCheckRun_(false); }
function previewBrandFixes()       { return brandCheckRun_(true); }

function brandCheckSheet_(createIfMissing) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(BRAND_CHECK_SHEET);
  if (!sh && createIfMissing) {
    sh = ss.insertSheet(BRAND_CHECK_SHEET);
    sh.getRange(1, 1, 1, 5).setValues([['id', '狀態', '表內品牌', '惠康品牌', '時間']])
      .setFontWeight('bold').setBackground('#e8f0fe');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 120); sh.setColumnWidth(3, 160); sh.setColumnWidth(4, 160);
  }
  return sh;
}

/** 品牌同品名對唔對得上（用嚟排優先前後） */
function brandMatchesName_(brand, name) {
  var b = safeStr(brand).trim();
  if (!b) return true;
  var low = safeStr(name).toLowerCase();
  if (!low) return true;
  if (low.indexOf(b.toLowerCase()) >= 0) return true;
  var toks = b.split(/[\s·・、,，\/&\-\.]+/);
  for (var i = 0; i < toks.length; i++) {
    var t = toks[i].trim();
    if (t.length >= 2 && low.indexOf(t.toLowerCase()) >= 0) return true;
  }
  return false;
}

/** 由產品頁 HTML 抽 JSON-LD 嘅品牌 */
function brandFromHtml_(html) {
  var m = safeStr(html).match(/<script[^>]*application\/ld\+json[^>]*>([\s\S]*?)<\/script>/i);
  if (!m) return '';
  try {
    var j = JSON.parse(m[1]);
    if (j && j.brand) return safeStr(j.brand.name || j.brand).trim();
  } catch (e) {}
  return '';
}

function brandCheckRun_(dryRun) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { error: '系統忙碌中，請稍後再試' };
  var started = Date.now();
  try {
    var sheet = getSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { ok: true, message: '冇資料' };

    var n = lastRow - 1;
    var ids    = sheet.getRange(2, COL_MAP.id, n, 1).getValues();
    var brands = sheet.getRange(2, COL_MAP.brand, n, 1).getValues();
    var names  = sheet.getRange(2, COL_MAP.name_zh, n, 1).getValues();
    var names2 = sheet.getRange(2, COL_MAP.name, n, 1).getValues();

    // 已經查過嘅
    var sh2 = brandCheckSheet_(false);
    var done = {}, failCount = {};
    if (sh2 && sh2.getLastRow() > 1) {
      var prev = sh2.getRange(2, 1, sh2.getLastRow() - 1, 2).getValues();
      for (var pi = 0; pi < prev.length; pi++) {
        var pid = safeStr(prev[pi][0]).trim();
        var st  = safeStr(prev[pi][1]).trim();
        if (!pid) continue;
        if (st === '抓取失敗') failCount[pid] = (failCount[pid] || 0) + 1;
        else done[pid] = true;
      }
    }
    // 連續失敗 3 次就當處理唔到（例如貨品已下架），唔再浪費額度
    Object.keys(failCount).forEach(function (k) { if (failCount[k] >= 3) done[k] = true; });

    // 未查過嘅，可疑嘅排先
    var suspects = [], others = [];
    for (var i = 0; i < n; i++) {
      var id = safeStr(ids[i][0]).trim();
      if (!id || done[id]) continue;
      var nm = safeStr(names[i][0]) + ' ' + safeStr(names2[i][0]);
      var item = { row: i + 2, id: id, brand: safeStr(brands[i][0]).trim(), name: nm };
      if (!brandMatchesName_(item.brand, nm)) suspects.push(item);
      else others.push(item);
    }
    var queue = suspects.concat(others).slice(0, BRAND_CHECK_PER_RUN);

    if (!queue.length) {
      return { ok: true, checked: 0, fixed: 0,
               message: '✅ 全部 ' + n + ' 件貨嘅品牌都已經覆核過。' };
    }

    var results = [], fixed = 0, same = 0, failed = 0;
    for (var q = 0; q < queue.length; q++) {
      if (Date.now() - started > BRAND_CHECK_TIME_MS) break;
      var it = queue[q];
      var official = '';
      try {
        var url = 'https://www.wellcome.com.hk/zh-hant/wellcome/p/p/i/' + it.id + '.html';
        var resp = UrlFetchApp.fetch(url, {
          muteHttpExceptions: true,
          followRedirects: true,
          headers: { 'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15' }
        });
        if (resp.getResponseCode() === 200) official = brandFromHtml_(resp.getContentText());
      } catch (e) { official = ''; }

      if (!official) { failed++; results.push([it.id, '抓取失敗', it.brand, '', new Date().toISOString()]); continue; }

      var sameBrand = official.replace(/\s+/g, '').toLowerCase() === it.brand.replace(/\s+/g, '').toLowerCase();
      if (sameBrand || !it.brand) {
        // 一樣；或者表內本來空白 → 補上（都算有用）
        if (!sameBrand && !it.brand && !dryRun) {
          sheet.getRange(it.row, COL_MAP.brand).setValue(official);
          sheet.getRange(it.row, COL_TIMESTAMP).setValue(new Date().toISOString());
          sheet.getRange(it.row, COL_SOURCE).setValue('wellcome');
          fixed++;
          results.push([it.id, '已補上', it.brand, official, new Date().toISOString()]);
        } else {
          same++;
          results.push([it.id, '一致', it.brand, official, new Date().toISOString()]);
        }
      } else {
        if (!dryRun) {
          sheet.getRange(it.row, COL_MAP.brand).setValue(official);
          sheet.getRange(it.row, COL_TIMESTAMP).setValue(new Date().toISOString());
          sheet.getRange(it.row, COL_SOURCE).setValue('wellcome');
        }
        fixed++;
        results.push([it.id, dryRun ? '會改' : '已修正', it.brand, official, new Date().toISOString()]);
      }
    }

    // 寫進度
    if (!dryRun && results.length) {
      var out = brandCheckSheet_(true);
      out.getRange(out.getLastRow() + 1, 1, results.length, 5).setValues(results);
      if (fixed > 0) { invalidateCache(); scheduleExport(); }
    }

    var summary = {
      ok: true, dryRun: !!dryRun,
      checked: results.length,
      fixed: fixed,
      same: same,
      failed: failed,
      remainingSuspects: Math.max(0, suspects.length - queue.filter(function (x) { return suspects.indexOf(x) >= 0; }).length),
      message: (dryRun ? '[預覽] ' : '') + '覆核 ' + results.length + ' 件：' +
               (dryRun ? '會改 ' : '修正 ') + fixed + ' 件，一致 ' + same + ' 件' +
               (failed ? '，抓取失敗 ' + failed + ' 件' : '') +
               (dryRun ? '（預覽唔會寫入，亦唔會記錄進度）' : '')
    };
    Logger.log('#52 ' + JSON.stringify(summary));
    return summary;
  } finally {
    lock.releaseLock();
  }
}

/** 建立每 6 小時自動覆核嘅觸發器（跑一次就得） */
function createBrandCheckTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'checkBrandsFromWellcome') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkBrandsFromWellcome').timeBased().everyHours(6).create();
  Logger.log('已建立每 6 小時品牌覆核觸發器');
}

function fixCategoryNames() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { error: '系統忙碌中，請稍後再試' };
  try {
    var sheet = getSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { ok: true, changed: 0, message: '冇資料' };

    // 1) 讀
    var rng  = sheet.getRange(2, COL_MAP.category_en, lastRow - 1, 1);
    var vals = rng.getValues();

    // 2) 先統計每種寫法出現幾多次 —— 用嚟決定「認唔到嘅值」要統一成邊個寫法。
    //    例如 room temperature milk ×5 + ROOM TEMPERATURE MILK ×2
    //    會統一成 room temperature milk（用得最多嗰個），唔會再分裂。
    var groups = {};            // key -> { 原寫法: 次數 }
    for (var i = 0; i < vals.length; i++) {
      var s0 = safeStr(vals[i][0]).trim();
      if (!s0) continue;
      var k0 = fixCatKey_(s0);
      if (!groups[k0]) groups[k0] = {};
      groups[k0][s0] = (groups[k0][s0] || 0) + 1;
    }

    // 3) 每個 key 決定一個正式寫法
    var byKey = {};
    FIX_CAT_CANONICAL.forEach(function (n) { byKey[fixCatKey_(n)] = n; });
    Object.keys(FIX_CAT_ALIAS).forEach(function (k) { byKey[fixCatKey_(k)] = FIX_CAT_ALIAS[k]; });

    var target = {};            // key -> 正式寫法
    var unknown = {};           // 認唔到嘅組（原封不動，只統一大小楷）
    Object.keys(groups).forEach(function (k) {
      if (byKey[k]) { target[k] = byKey[k]; return; }
      var v = groups[k], best = null, bestN = -1;
      Object.keys(v).forEach(function (x) {
        if (v[x] > bestN) { bestN = v[x]; best = x; }
      });
      target[k] = best;
      unknown[best] = true;
    });

    // 4) 套用
    var changed = 0;
    for (var j = 0; j < vals.length; j++) {
      var raw = safeStr(vals[j][0]);
      var t = raw.trim();
      if (!t) continue;
      var want = target[fixCatKey_(t)];
      if (want && want !== raw) { vals[j][0] = want; changed++; }
    }

    if (changed > 0) {
      rng.setValues(vals);
      invalidateCache();
      scheduleExport();
    }

    var unk = Object.keys(unknown).sort();
    Logger.log('fixCategoryNames: 統一咗 ' + changed + ' 件。認唔到（只統一大小楷，冇合併）：' +
               (unk.length ? unk.join('、') : '冇'));

    return {
      ok: true,
      changed: changed,
      unknown: unknown,
      message: '已統一 ' + changed + ' 件貨嘅分類名' +
               (unk.length ? '；認唔到：' + unk.join('、') : '')
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * #45 批量套用「建議分類」。
 *   items = [{ id, category_en, category_zh }]
 *
 * 安全設計（呢個係一次改幾百行，所以特別保守）：
 *   1. 每一行都要先對得上客戶端提供嘅 expect 值（佢以為而家係咩分類）。
 *      對唔上就跳過 —— 即係話任何已經有分類嘅行（包括 Case Deal 原箱）
 *      都唔會被覆蓋，除非客戶端明確講明佢就係呢個值。
 *   2. 一次過讀 5 欄、一次過寫 4 欄 —— 唔會逐行 update。
 *      （457 行 × 每次 2 個 sheet call = 914 次，會爆 6 分鐘同長時間霸住鎖）
 *   3. H 欄（英文分類）同 I 欄（中文細分）一定一齊寫，唔會得一半。
 *   4. 寫入 T（時間）／U（來源）方便日後追查。
 */
function bulkFixCategory(items) {
  if (!Array.isArray(items) || !items.length) return { error: '冇資料' };

  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, applied: 0, skipped: 0, appliedIds: [] };

  var n = lastRow - 1;
  var rngId = sheet.getRange(2, COL_MAP.id, n, 1);
  var rngH  = sheet.getRange(2, COL_MAP.category_en, n, 1);
  var rngI  = sheet.getRange(2, COL_MAP.category_zh, n, 1);
  var rngT  = sheet.getRange(2, COL_TIMESTAMP, n, 1);
  var rngU  = sheet.getRange(2, COL_SOURCE, n, 1);

  var ids = rngId.getValues();
  var hs  = rngH.getValues();
  var is  = rngI.getValues();
  var ts  = rngT.getValues();
  var us  = rngU.getValues();

  var rowOf = {};
  for (var i = 0; i < n; i++) {
    var id = safeStr(ids[i][0]).trim();
    if (id) rowOf[id] = i;
  }

  var nowIso = new Date().toISOString();
  var applied = 0, skipped = 0;
  var appliedIds = [];

  for (var j = 0; j < items.length; j++) {
    var it = items[j] || {};
    var pid = safeStr(it.id).trim();
    var k = rowOf[pid];
    if (k === undefined) { skipped++; continue; }

    // 安全鎖：客戶端要講明「佢以為 H 欄而家係咩值」，同實際唔一樣就唔改。
    //   冇提供 expect 時當作空字串 → 只會寫入「完全冇分類」嘅行。
    //   ⚠️ Case Deal（原箱）係真分類，所以呢支更加唔可以靠「Case Deal 就代表未分類」。
    var curH = safeStr(hs[k][0]).trim();
    var expect = safeStr(it.expect).trim();
    if (curH !== expect) { skipped++; continue; }

    var want = safeStr(it.category_en).trim();
    if (!want) { skipped++; continue; }             // 「Case Deal」係真分類，可以寫
    if (curH === want) { skipped++; continue; }     // 值一樣，唔使寫（免得白改時間戳）

    hs[k][0] = want;
    is[k][0] = safeStr(it.category_zh).trim();
    ts[k][0] = nowIso;
    us[k][0] = 'user';
    applied++;
    appliedIds.push(pid);
  }

  if (applied > 0) {
    rngH.setValues(hs);
    rngI.setValues(is);
    rngT.setValues(ts);
    rngU.setValues(us);
    invalidateCache();
    scheduleExport();
  }

  Logger.log('bulkFixCategory: applied=' + applied + ' skipped=' + skipped);
  return { ok: true, applied: applied, skipped: skipped, appliedIds: appliedIds };
}

/**
 * #48 一次性修復：還原「原箱（Case Deal）」被批量改走嘅分類。
 *
 * 背景：Case Deal（原箱）其實係【真分類】—— 實務上原箱係掛喺貨品類別之下，
 *       用嚟辨識大客買大量貨。但係之前嘅程式錯誤地當咗佢係「促銷分類、唔係商品分類」，
 *       結果 2026-10-19 用「一鍵套用建議分類」時，把 440 件原箱貨改走咗。
 *
 * 做法：只還原【嗰一次批量操作】改過嘅行 —— 靠 T 欄（最後修改時間）嘅
 *       時間戳辨識，所以唔會影響任何本來就有商品分類嘅原箱貨。
 *
 * 用法：
 *   ① 先跑 previewCaseDealRestore()  → 只係睇，唔會改任何嘢
 *   ② 確認數字對就跑 restoreCaseDealBulk()
 */
var CASE_DEAL_BULK_TS = '2026-10-19T13:24:49';

function previewCaseDealRestore() { return caseDealRestore_(true); }
function restoreCaseDealBulk()    { return caseDealRestore_(false); }

function caseDealRestore_(dryRun) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { error: '系統忙碌中，請稍後再試' };
  try {
    var sheet = getSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { ok: true, matched: 0, restored: 0 };

    var n = lastRow - 1;
    var rngH  = sheet.getRange(2, COL_MAP.category_en, n, 1);
    var rngT  = sheet.getRange(2, COL_TIMESTAMP, n, 1);
    var rngU  = sheet.getRange(2, COL_SOURCE, n, 1);
    var rngZ  = sheet.getRange(2, COL_MAP.name_zh, n, 1);
    var rngE  = sheet.getRange(2, COL_MAP.name, n, 1);

    var hs = rngH.getValues();
    var ts = rngT.getValues();
    var us = rngU.getValues();
    var zs = rngZ.getValues();
    var es = rngE.getValues();

    var nowIso = new Date().toISOString();
    var matched = 0, restored = 0;
    var samples = [];

    for (var i = 0; i < n; i++) {
      var t = safeStr(ts[i][0]);
      if (t.indexOf(CASE_DEAL_BULK_TS) < 0) continue;
      matched++;

      var cur = safeStr(hs[i][0]).trim();
      if (/^case deal$/i.test(cur)) continue;      // 已經係 Case Deal，唔使改

      restored++;
      if (samples.length < 10) {
        samples.push(safeStr(zs[i][0]) || safeStr(es[i][0]) || ('row ' + (i + 2)));
      }
      if (!dryRun) {
        hs[i][0] = 'Case Deal';
        ts[i][0] = nowIso;
        us[i][0] = 'user';
      }
    }

    if (!dryRun && restored > 0) {
      rngH.setValues(hs);
      rngT.setValues(ts);
      rngU.setValues(us);
      invalidateCache();
      scheduleExport();
    }

    var msg = (dryRun ? '[預覽] ' : '[已執行] ') +
              '時間戳符合 ' + matched + ' 件，需要還原 ' + restored + ' 件';
    Logger.log('caseDealRestore: ' + msg + '\n樣本：' + samples.join(' / '));

    return { ok: true, dryRun: dryRun, matched: matched, restored: restored,
             samples: samples, message: msg };
  } finally {
    lock.releaseLock();
  }
}

function emptyProduct() {
  var o = {};
  ALL_FIELDS.forEach(function(k) { o[k] = ''; });
  o._last_modified = '';
  o._last_source = '';
  return o;
}

function rowToProduct(row) {
  if (!isValidRow(row)) return emptyProduct();
  var obj = {};
  ALL_FIELDS.forEach(function(k) {
    var idx = COL_MAP[k] - 1;
    obj[k] = safeStr(row[idx]);
  });
  obj._last_modified = safeStr(row[COL_TIMESTAMP - 1]);
  obj._last_source   = safeStr(row[COL_SOURCE - 1]);
  return obj;
}

function productToRow(p, timestamp, source) {
  if (!isObj(p)) {
    Logger.log('⚠️ productToRow received non-object: ' + JSON.stringify(p));
    p = {};
  }
  var row = new Array(MAX_COL).fill('');
  ALL_FIELDS.forEach(function(k) {
    row[COL_MAP[k] - 1] = p[k] != null ? p[k] : '';
  });
  row[COL_TIMESTAMP - 1] = timestamp || new Date().toISOString();
  row[COL_SOURCE - 1]    = source || '';
  return row;
}

function hammingDistance(h1, h2) {
  if (!h1 || !h2 || h1.length !== h2.length) return 999;
  var dist = 0;
  for (var i = 0; i < h1.length; i++) {
    var xor = parseInt(h1[i], 16) ^ parseInt(h2[i], 16);
    while (xor) { dist += xor & 1; xor >>= 1; }
  }
  return dist;
}


// ============================================================
//  OOS 工具
// ============================================================
function getOOSSheet(createIfMissing) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(OOS_SHEET);
  if (!sheet && createIfMissing) {
    sheet = ss.insertSheet(OOS_SHEET);
    sheet.getRange(1, 1, 1, 13).setValues([
      ['時間', '產品ID', '產品名稱', '回報人', '原因', '狀態', '補貨時間',
       '單日期', '單號尾4', '派送時段', '單類型', '數量', '跟進動作']
    ]);
    sheet.getRange(1, 1, 1, 13).setFontWeight('bold').setBackground('#fce8e6');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 170);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 220);
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 220);
    sheet.setColumnWidth(6, 100);
    sheet.setColumnWidth(7, 170);
    sheet.setColumnWidth(8, 100);
    sheet.setColumnWidth(9, 90);
    sheet.setColumnWidth(10, 110);
    sheet.setColumnWidth(11, 90);
    sheet.setColumnWidth(12, 80);
    sheet.setColumnWidth(13, 130);
  }
  if (sheet) {
    var h8 = sheet.getRange(1, 8).getValue();
    if (!h8) {
      sheet.getRange(1, 8, 1, 4).setValues([
        ['單日期', '單號尾4', '派送時段', '單類型']
      ]);
      sheet.getRange(1, 8, 1, 4).setFontWeight('bold').setBackground('#fce8e6');
      sheet.setColumnWidth(8, 100);
      sheet.setColumnWidth(9, 90);
      sheet.setColumnWidth(10, 110);
      sheet.setColumnWidth(11, 90);
    }
    // #18：第 12 欄「數量」＝該次回報的缺貨件數（同產品多張單會累加）。
    //      舊試算表沒有這欄，這裡自動補上。
    var h12 = sheet.getRange(1, 12).getValue();
    if (!h12) {
      sheet.getRange(1, 12).setValue('數量');
      sheet.getRange(1, 12).setFontWeight('bold').setBackground('#fce8e6');
      sheet.setColumnWidth(12, 80);
    }

    // #53：第 13 欄「跟進動作」——每次情況唔同，所以記喺【每筆記錄】身上，
    //      唔可以綁死喺原因度（同一個原因今日要搵、聽日等補貨）。
    var h13 = sheet.getRange(1, 13).getValue();
    if (!h13) {
      sheet.getRange(1, 13).setValue('跟進動作');
      sheet.getRange(1, 13).setFontWeight('bold').setBackground('#fce8e6');
      sheet.setColumnWidth(13, 130);
    }

    // #19：第 9 欄「單號尾4」強制純文字格式。
    //      否則 0 開頭的單號（例："0123"）會被試算表當成數字、存成 123。
    //      只需設定一次，用 ScriptProperties 記住，避免每次回報都重設格式。
    var props = PropertiesService.getScriptProperties();
    if (props.getProperty('OOS_TAIL_TEXT_FMT') !== '1') {
      sheet.getRange(2, 9, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat('@');
      props.setProperty('OOS_TAIL_TEXT_FMT', '1');
    }
  }
  return sheet;
}

function getActiveOOSMap() {
  var result = {};
  var sheet = getOOSSheet(false);
  if (!sheet) return result;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return result;

  var rows = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
  rows.forEach(function(r) {
    if (!isValidRow(r)) return;
    var id = safeStr(r[1]).trim();
    var status = safeStr(r[5]).trim();
    if (!id || status !== 'active') return;

    // #18：同一件貨可能有多張單缺貨（多筆 active），數量要累加。
    //      列是由舊到新，所以後面覆蓋前面 → 訂單資訊取「最新一張單」。
    // 舊記錄沒有「數量」欄 → 當作 1 件（一筆回報至少缺 1 件），
    // 否則歷史資料的累計會全部變成 0。
    var qRaw = safeStr(r[11]).trim();
    var q = qRaw === '' ? 1 : parseInt(qRaw.replace(/[^0-9]/g, ''), 10);
    if (isNaN(q) || q < 0) q = 1;            // #60：0 件合法（行架巡查）

    if (!result[id]) {
      result[id] = { at: '', reason: '', reporter: '', orderDate: '', orderTail: '',
                     timeSlot: '', orderType: '', qty: 0, reports: 0 };
    }
    result[id].qty     += q;
    result[id].reports += 1;
    result[id].at         = safeStr(r[0]);
    result[id].reason     = safeStr(r[4]);
    result[id].reporter   = safeStr(r[3]);
    result[id].orderDate  = safeStr(r[7]);
    result[id].orderTail  = safeStr(r[8]);
    result[id].timeSlot   = safeStr(r[9]);
    result[id].orderType  = safeStr(r[10]);
  });
  return result;
}

/**
 * #20：回傳 _OOS 的「逐筆記錄」（不是按產品彙總的 getActiveOOSMap）。
 *      前端要用「單日期」篩出某日有哪幾張單缺貨，
 *      但 getActiveOOSMap 每件貨只保留最新一張單，無法做到，所以另開這支。
 *      date（YYYYMMDD，可選）留空＝回傳全部。
 */
function getOOSReports(date) {
  var sheet = getOOSSheet(false);
  if (!sheet) return {ok: true, date: safeStr(date).trim(), reports: []};

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {ok: true, date: safeStr(date).trim(), reports: []};

  var want = safeStr(date).trim();
  var rows = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  var out = [];

  // 由新到舊（試算表是舊在上）
  for (var i = rows.length - 1; i >= 0; i--) {
    var r = rows[i];
    if (!isValidRow(r)) continue;
    var orderDate = safeStr(r[7]).trim();
    if (want && orderDate !== want) continue;

    var qRaw = safeStr(r[11]).trim();
    var q = qRaw === '' ? 1 : parseInt(qRaw.replace(/[^0-9]/g, ''), 10);
    if (isNaN(q) || q < 0) q = 1;            // #60：0 件合法（行架巡查）

    out.push({
      row:         i + 2,
      at:          safeStr(r[0]),
      productId:   safeStr(r[1]).trim(),
      productName: safeStr(r[2]),
      reporter:    safeStr(r[3]),
      reason:      safeStr(r[4]),
      status:      safeStr(r[5]).trim(),
      restockedAt: safeStr(r[6]),
      orderDate:   orderDate,
      orderTail:   safeStr(r[8]).trim(),
      timeSlot:    safeStr(r[9]).trim(),
      orderType:   safeStr(r[10]).trim(),
      qty:         q,
      action:      safeStr(r[12]).trim()      // #53：跟進動作
    });
  }
  return {ok: true, date: want, reports: out};
}

// #54：_OOS 各欄位置
var OOS_COL_DATE = 8, OOS_COL_TAIL = 9, OOS_COL_SLOT = 10,
    OOS_COL_TYPE = 11, OOS_COL_QTY = 12, OOS_COL_ACTION = 13;

/**
 * #54 事後修改某一筆缺貨記錄 —— 日期／單號／時段／類型／件數／跟進動作。
 * （回報時打錯日子、揀錯時段，喺 App 嘅缺貨紀錄度就改得返，唔使重新回報。）
 *
 * 用 row ＋ productId ＋ at 三重核對，避免行號移咗而改錯行。
 * 只會改有提供嘅欄位，冇提供嘅一律唔碰。
 */
function updateOOSRecord(body) {
  if (!isObj(body)) return { error: 'Invalid body' };
  var sheet = getOOSSheet(false);
  if (!sheet) return { error: '未有缺貨記錄' };
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { error: '未有缺貨記錄' };

  var pid = safeStr(body.productId).trim();
  var at  = safeStr(body.at).trim();
  if (!pid) return { error: 'Missing productId' };

  var f = isObj(body.fields) ? body.fields : {};
  var out = {};

  if (f.orderDate !== undefined) {
    var d = safeStr(f.orderDate).trim();
    if (d && !/^\d{8}$/.test(d)) return { error: '日期格式需為 YYYYMMDD（8 位數字）' };
    out[OOS_COL_DATE] = d;
  }
  if (f.orderTail !== undefined) {
    var t = safeStr(f.orderTail).trim();
    if (t && !/^\d{1,4}$/.test(t)) return { error: '單號尾需為 1~4 位數字' };
    out[OOS_COL_TAIL] = t;
  }
  if (f.timeSlot !== undefined) {
    var s = safeStr(f.timeSlot).trim();
    if (s && ['早','中','晚','放假'].indexOf(s) < 0) return { error: '派送時段需為 早 / 中 / 晚 / 放假' };
    out[OOS_COL_SLOT] = s;
  }
  if (f.orderType !== undefined) {
    var ty = safeStr(f.orderType).trim();
    if (ty && ['O2O','CNC','B2B'].indexOf(ty) < 0) return { error: '單類型需為 O2O / CNC / B2B' };
    out[OOS_COL_TYPE] = ty;
  }
  if (f.qty !== undefined) {
    var q = parseInt(String(f.qty).replace(/[^0-9]/g, ''), 10);
    if (isNaN(q) || q < 0) return { error: '缺貨件數需為 0 或以上' };
    if (q > 9999) q = 9999;
    out[OOS_COL_QTY] = q;                    // #60：0 合法（巡查缺貨）
  }
  if (f.action !== undefined) {
    var a = safeStr(f.action).trim();
    if (a.length > 40) a = a.substring(0, 40);
    out[OOS_COL_ACTION] = a;
  }

  var cols = Object.keys(out).map(Number);
  if (!cols.length) return { error: '冇嘢要改' };

  // ---- 定位（row + productId + at 核對）----
  var rows = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  var targetRow = 0;
  var r0 = parseInt(body.row, 10);
  if (!isNaN(r0) && r0 >= 2 && r0 <= lastRow) {
    var rr = rows[r0 - 2];
    if (safeStr(rr[1]).trim() === pid && (!at || safeStr(rr[0]).trim() === at)) targetRow = r0;
  }
  if (!targetRow) {
    for (var i = rows.length - 1; i >= 0; i--) {
      if (safeStr(rows[i][1]).trim() !== pid) continue;
      if (at && safeStr(rows[i][0]).trim() !== at) continue;
      targetRow = i + 2;
      break;
    }
  }
  if (!targetRow) return { error: '搵唔到呢筆缺貨記錄' };

  // 單號尾要鎖純文字，否則 0 開頭會變數字
  if (out[OOS_COL_TAIL] !== undefined) sheet.getRange(targetRow, OOS_COL_TAIL).setNumberFormat('@');
  cols.forEach(function (c) { sheet.getRange(targetRow, c).setValue(out[c]); });

  invalidateCache();
  scheduleExport();
  Logger.log('#54 updateOOSRecord row=' + targetRow + ' cols=' + cols.join(','));
  return { ok: true, row: targetRow, updated: cols };
}

/**
 * #53 舊入口：只改跟進動作。
 * ⚠️ body.action 已經用嚟做派發（'setOOSAction'），唔可以再當動作值讀，
 *    所以動作值要用 body.actionValue。
 */
function setOOSAction(body) {
  if (!isObj(body)) return { error: 'Invalid body' };
  var val = (body.actionValue !== undefined) ? body.actionValue
          : (isObj(body.fields) ? body.fields.action : '');
  return updateOOSRecord({
    row: body.row, productId: body.productId, at: body.at,
    fields: { action: val }
  });
}

function getProductOOSHistory(productId, limit) {
  if (!productId) return {error: 'Missing productId'};
  var sheet = getOOSSheet(false);
  if (!sheet) return {ok: true, productId: productId, history: []};

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {ok: true, productId: productId, history: []};

  var rows = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  var target = String(productId).trim();
  var history = [];
  var max = limit || 5;

  for (var i = rows.length - 1; i >= 0; i--) {
    if (!isValidRow(rows[i])) continue;
    var rid = safeStr(rows[i][1]).trim();
    if (rid !== target) continue;
    history.push({
      at:          safeStr(rows[i][0]),
      reporter:    safeStr(rows[i][3]),
      reason:      safeStr(rows[i][4]),
      status:      safeStr(rows[i][5]),
      restockedAt: safeStr(rows[i][6]),
      orderDate:   safeStr(rows[i][7]),
      orderTail:   safeStr(rows[i][8]),
      timeSlot:    safeStr(rows[i][9]),
      orderType:   safeStr(rows[i][10]),
      qty:         safeStr(rows[i][11]),
      action:      safeStr(rows[i][12])       // #53
    });
    if (history.length >= max) break;
  }
  return {ok: true, productId: productId, history: history};
}


// ============================================================
//  臨時 ID 清單
// ============================================================
function getTempIdProducts() {
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {count: 0, items: []};
  var rows = sheet.getRange(2, 1, lastRow - 1, MAX_COL).getValues();
  var items = [];
  for (var i = 0; i < rows.length; i++) {
    if (!isValidRow(rows[i])) continue;
    var id = safeStr(rows[i][0]).trim();
    var source = safeStr(rows[i][COL_SOURCE - 1]).trim();
    if (source === SOURCE_SCAN_AUTO && isTempId(id)) {
      items.push({
        id: id,
        row: i + 2,
        barcode: safeStr(rows[i][1]),
        name: safeStr(rows[i][2]),
        name_zh: safeStr(rows[i][3]),
        created_at: safeStr(rows[i][COL_TIMESTAMP - 1])
      });
    }
  }
  return {count: items.length, items: items};
}


// ============================================================
//  OOS 原因（_OOSReasons）
// ============================================================
function getOOSReasonsSheet(createIfMissing) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(OOS_REASONS_SHEET);
  if (!sheet && createIfMissing) {
    sheet = ss.insertSheet(OOS_REASONS_SHEET);
    sheet.getRange(1, 1, 1, 4).setValues([
      ['儲存值 (value)', '顯示名稱 (label)', '排序', '跟進動作 (D 欄)']
    ]);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#fff3cd');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 220);
    sheet.setColumnWidth(2, 260);
    sheet.setColumnWidth(3, 70);
    sheet.setColumnWidth(4, 160);

    sheet.getRange(2, 1, 6, 4).setValues([
      ['貨架空了', '🛒 貨架空了', 1, '要搵'],
      ['客人買光', '🔥 客人買光', 2, '等補貨'],
      ['補貨中',   '⏳ 補貨中',   3, '等補貨'],
      ['找不到貨', '🔍 找不到貨', 4, '要搵'],
      ['貨物損壞', '💔 貨物損壞', 5, '等補貨'],
      ['過期',     '📅 過期',     6, '等補貨']
    ]);
  }
  return sheet;
}

function getOOSReasons() {
  var sheet = getOOSReasonsSheet(false);
  if (!sheet) {
    sheet = getOOSReasonsSheet(true);
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // #50：自動補 D 欄（跟進動作）嘅標題，舊表都唔會出事
  var head = safeStr(sheet.getRange(1, 4).getValue()).trim();
  if (!head) {
    sheet.getRange(1, 4).setValue('跟進動作 (D 欄)').setFontWeight('bold').setBackground('#fff3cd');
    sheet.setColumnWidth(4, 160);
  }

  var rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var v = safeStr(rows[i][0]).trim();
    var l = safeStr(rows[i][1]).trim();
    var o = rows[i][2];
    if (!v) continue;
    out.push({
      value: v,
      label: l || v,
      order: (o === '' || o == null) ? 999 : Number(o),
      action: safeStr(rows[i][3]).trim()     // #50：D 欄＝跟進動作（自由文字）
    });
  }
  out.sort(function(a, b) { return a.order - b.order; });
  return out.map(function(x) {
    return { value: x.value, label: x.label, action: x.action };
  });
}

// ============================================================
//  #53 跟進動作（_OOSActions）——每次情況唔同，所以記喺每筆缺貨記錄度
// ============================================================
var OOS_ACTIONS_SHEET = '_OOSActions';

function getOOSActionsSheet(createIfMissing) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(OOS_ACTIONS_SHEET);
  if (!sh && createIfMissing) {
    sh = ss.insertSheet(OOS_ACTIONS_SHEET);
    sh.getRange(1, 1, 1, 3).setValues([['儲存值 (value)', '顯示名稱 (label)', '排序']]);
    sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#e6f4ea');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 160);
    sh.setColumnWidth(2, 260);
    sh.setColumnWidth(3, 70);
    sh.getRange(2, 1, 3, 3).setValues([
      ['要搵',       '🔍 要搵（搬咗去地堆／唔知擺邊）', 1],
      ['等補貨',     '⏳ 等補貨（真係賣晒）',           2],
      ['要問runner', '❓ 要問 runner（DDC 網購預留）',   3]
    ]);
  }
  return sh;
}

function getOOSActions() {
  var sh = getOOSActionsSheet(true);
  if (!sh) return [];
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var rows = sh.getRange(2, 1, lastRow - 1, 3).getValues();
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var v = safeStr(rows[i][0]).trim();
    if (!v) continue;
    var l = safeStr(rows[i][1]).trim();
    var o = rows[i][2];
    out.push({ value: v, label: l || v, order: (o === '' || o == null) ? 999 : Number(o) });
  }
  out.sort(function (a, b) { return a.order - b.order; });
  return out.map(function (x) { return { value: x.value, label: x.label }; });
}

function initOOSActionsSheet() {
  getOOSActionsSheet(true);
  Logger.log('✅ _OOSActions 已建立');
  return { ok: true };
}

// ============================================================
//  #57 訂單件數（_OrderTotals）—— 一張單有幾多件，按貨類分
//    A 時間 | B 單日期 | C 單號尾4 | D 派送時段 | E 單類型
//    F 干貨 | G 旦貨 | H 風貨(0-4度) | I 凍貨(-18度) | J 大件(sim卡/禮券)
//    K 總件數 | L 填寫人
//    同一張單（日期+單號尾+時段+類型）只有一組數字，再存會覆蓋更新。
// ============================================================
var ORDER_TOTALS_SHEET = '_OrderTotals';
var OT_KINDS = [
  { key: 'dry',      label: '干貨',               col: 6 },
  { key: 'egg',      label: '旦貨',               col: 7 },
  { key: 'chill',    label: '風貨（0-4度）',       col: 8 },
  { key: 'frozen',   label: '凍貨（-18度）',       col: 9 },
  { key: 'bulky',    label: '大件（sim卡/禮券）',  col: 10 },
  { key: 'original', label: '原箱',               col: 11 }   // #60
];

function getOrderTotalsSheet(createIfMissing) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(ORDER_TOTALS_SHEET);
  if (!sh && createIfMissing) {
    sh = ss.insertSheet(ORDER_TOTALS_SHEET);
    otWriteHeader_(sh);
    sh.setFrozenRows(1);
    sh.getRange(2, 3, Math.max(1, sh.getMaxRows() - 1), 1).setNumberFormat('@');
  } else {
    // #60：舊表冇「原箱」欄 → 喺第 11 欄插入，總件數／填寫人自動右移
    var h11 = safeStr(sh.getRange(1, 11).getValue()).trim();
    if (h11 !== '原箱') {
      sh.insertColumnBefore(11);
      otWriteHeader_(sh);
      sh.getRange(2, 3, Math.max(1, sh.getMaxRows() - 1), 1).setNumberFormat('@');
      Logger.log('#60 _OrderTotals 已插入「原箱」欄');
    }
  }
  return sh;
}

/** #60 _OrderTotals 表頭（13 欄） */
function otWriteHeader_(sh) {
  sh.getRange(1, 1, 1, 13).setValues([[
    '時間', '單日期', '單號尾4', '派送時段', '單類型',
    '干貨', '旦貨', '風貨(0-4度)', '凍貨(-18度)', '大件(sim卡/禮券)', '原箱',
    '總件數', '填寫人'
  ]]);
  sh.getRange(1, 1, 1, 13).setFontWeight('bold').setBackground('#e8f0fe');
  var widths = [160, 100, 90, 110, 90, 70, 70, 110, 110, 150, 70, 80, 100];
  for (var wi = 0; wi < widths.length; wi++) sh.setColumnWidth(wi + 1, widths[wi]);
}

/**
 * #58 訂單件數嘅身分 = 單日期 ＋ 單號尾4。
 *   時段／類型只係附加資料，唔計入 key —— 因為同一張單嘅缺貨可能跨時段，
 *   如果用四樣做 key，缺貨紀錄度就對唔返。
 */
function orderTotalsKey_(date, tail) {
  return safeStr(date).trim() + '|' + safeStr(tail).trim();
}

function getOrderTotals() {
  var sh = getOrderTotalsSheet(false);
  if (!sh) return { ok: true, totals: [] };
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: true, totals: [] };

  var rows = sh.getRange(2, 1, lastRow - 1, 13).getValues();
  var out = [];
  for (var i = rows.length - 1; i >= 0; i--) {          // 由新到舊
    var r = rows[i];
    var date = safeStr(r[1]).trim();
    if (!date) continue;
    out.push({
      row:       i + 2,
      at:        safeStr(r[0]),
      orderDate: date,
      orderTail: safeStr(r[2]).trim(),
      timeSlot:  safeStr(r[3]).trim(),
      orderType: safeStr(r[4]).trim(),
      dry:      Number(safeStr(r[5]))  || 0,
      egg:      Number(safeStr(r[6]))  || 0,
      chill:    Number(safeStr(r[7]))  || 0,
      frozen:   Number(safeStr(r[8]))  || 0,
      bulky:    Number(safeStr(r[9]))  || 0,
      original: Number(safeStr(r[10])) || 0,
      total:    Number(safeStr(r[11])) || 0,
      by:       safeStr(r[12])
    });
  }
  return { ok: true, totals: out };
}

/**
 * 儲存／更新一張單嘅件數。
 * body = { orderDate, orderTail, timeSlot, orderType,
 *          counts: {dry,egg,chill,frozen,bulky}, by }
 * 同一 key（日期+單號尾+時段+類型）已存在 → 覆蓋該行；否則新增。
 */
function saveOrderTotals(body) {
  if (!isObj(body)) return { error: 'Invalid body' };

  var date = safeStr(body.orderDate).trim();
  var tail = safeStr(body.orderTail).trim();
  var slot = safeStr(body.timeSlot).trim();
  var type = safeStr(body.orderType).trim();

  if (!date) return { error: '請填單日期' };
  if (!/^\d{8}$/.test(date)) return { error: '日期格式需為 YYYYMMDD（8 位數字）' };
  if (!tail) return { error: '請填單號尾（件數係跟單號走）' };
  if (!/^\d{1,4}$/.test(tail)) return { error: '單號尾需為 1~4 位數字' };
  if (slot && ['早','中','晚','放假'].indexOf(slot) < 0) return { error: '派送時段需為 早 / 中 / 晚 / 放假' };
  if (type && ['O2O','CNC','B2B'].indexOf(type) < 0) return { error: '單類型需為 O2O / CNC / B2B' };

  var c = isObj(body.counts) ? body.counts : {};
  var nums = {};
  for (var k = 0; k < OT_KINDS.length; k++) {
    var kd = OT_KINDS[k];
    var raw = safeStr(c[kd.key]).replace(/[^0-9]/g, '');
    var n = raw === '' ? 0 : parseInt(raw, 10);
    if (isNaN(n) || n < 0) return { error: kd.label + ' 件數不可為負數' };
    if (n > 9999) n = 9999;
    nums[kd.key] = n;                       // #60：0 係合法（巡查用）
  }
  var total = nums.dry + nums.egg + nums.chill + nums.frozen + nums.bulky + nums.original;

  var sh = getOrderTotalsSheet(true);
  var lastRow = sh.getLastRow();

  var key = orderTotalsKey_(date, tail);
  var targetRow = 0;
  if (lastRow >= 2) {
    var kc = sh.getRange(2, 2, lastRow - 1, 2).getValues();     // 只需 B 日期 + C 單號尾
    for (var i = 0; i < kc.length; i++) {
      if (orderTotalsKey_(kc[i][0], kc[i][1]) === key) {
        targetRow = i + 2;
        break;
      }
    }
  }

  var rowData = [new Date().toISOString(), date, tail, slot, type,
                 nums.dry, nums.egg, nums.chill, nums.frozen, nums.bulky, nums.original,
                 total, safeStr(body.by) || ''];
  var isUpdate = targetRow > 0;
  if (!isUpdate) targetRow = sh.getLastRow() + 1;
  sh.getRange(targetRow, 3).setNumberFormat('@');
  sh.getRange(targetRow, 1, 1, 13).setValues([rowData]);

  scheduleExport();
  Logger.log('#57 saveOrderTotals ' + key + ' total=' + total + ' row=' + targetRow);
  return { ok: true, row: targetRow, updated: isUpdate, total: total,
           counts: nums, orderDate: date, orderTail: tail, timeSlot: slot, orderType: type };
}

// ============================================================
//  #61 替代貨品（_Substitutes）
//    A 更新時間 | B 原貨品ID | C 原貨品名 | D 替代貨品ID | E 替代條碼 | F 替代品名 | G 使用次數
//    揀 🎩替代貨品 缺貨時記錄；下次同一件貨再揀，前端會問「係咪用返上次嗰件」。
// ============================================================
var SUBSTITUTES_SHEET = '_Substitutes';

function getSubstitutesSheet(createIfMissing) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SUBSTITUTES_SHEET);
  if (!sh && createIfMissing) {
    sh = ss.insertSheet(SUBSTITUTES_SHEET);
    sh.getRange(1, 1, 1, 7).setValues([[
      '更新時間', '原貨品ID', '原貨品名', '替代貨品ID', '替代條碼', '替代品名', '使用次數'
    ]]);
    sh.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#fff3e0');
    sh.setFrozenRows(1);
    var w = [170, 110, 220, 110, 150, 220, 80];
    for (var i = 0; i < w.length; i++) sh.setColumnWidth(i + 1, w[i]);
    sh.getRange(2, 2, Math.max(1, sh.getMaxRows() - 1), 1).setNumberFormat('@');
  }
  return sh;
}

function getSubstitutes() {
  var sh = getSubstitutesSheet(false);
  if (!sh) return { ok: true, subs: [] };
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: true, subs: [] };

  var rows = sh.getRange(2, 1, lastRow - 1, 7).getValues();
  var out = [];
  for (var i = rows.length - 1; i >= 0; i--) {
    var r = rows[i];
    var pid = safeStr(r[1]).trim();
    if (!pid) continue;
    out.push({
      at:          safeStr(r[0]),
      productId:   pid,
      productName: safeStr(r[2]),
      subId:       safeStr(r[3]).trim(),
      subBarcode:  safeStr(r[4]).trim(),
      subName:     safeStr(r[5]),
      uses:        Number(safeStr(r[6])) || 0
    });
  }
  return { ok: true, subs: out };
}

/** 儲存／更新某件貨嘅替代貨品（同一件原貨品只保留一個替代貨品） */
function saveSubstitute(body) {
  if (!isObj(body)) return { error: 'Invalid body' };
  var pid = safeStr(body.productId).trim();
  if (!pid) return { error: 'Missing productId' };

  var subId  = safeStr(body.subId).trim();
  var subBc  = safeStr(body.subBarcode).trim();
  var subNm  = safeStr(body.subName);
  if (!subId && !subBc) return { error: '請提供替代貨品' };

  var sh = getSubstitutesSheet(true);
  var lastRow = sh.getLastRow();
  var targetRow = 0, uses = 0;
  if (lastRow >= 2) {
    var ids = sh.getRange(2, 2, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (safeStr(ids[i][0]).trim() === pid) {
        targetRow = i + 2;
        uses = Number(safeStr(sh.getRange(targetRow, 7).getValue())) || 0;
        break;
      }
    }
  }

  var row = [new Date().toISOString(), pid, safeStr(body.productName),
             subId, subBc, subNm, uses + 1];
  if (!targetRow) targetRow = sh.getLastRow() + 1;
  sh.getRange(targetRow, 2).setNumberFormat('@');
  sh.getRange(targetRow, 1, 1, 7).setValues([row]);

  scheduleExport();
  Logger.log('#61 saveSubstitute ' + pid + ' → ' + (subId || subBc) + '（第 ' + uses + 1 + ' 次）');
  return { ok: true, row: targetRow, updated: targetRow <= lastRow, uses: uses + 1 };
}

function initOOSReasonsSheet() {
  var sheet = getOOSReasonsSheet(true);
  Logger.log('✅ _OOSReasons 已建立');
  return {ok: true};
}


// ============================================================
//  ★ v5.4：自訂繁忙日（_BusyDays）
// ============================================================
function getBusyDaysSheet(createIfMissing) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(BUSY_DAYS_SHEET);
  if (!sheet && createIfMissing) {
    sheet = ss.insertSheet(BUSY_DAYS_SHEET);
    sheet.getRange(1, 1, 1, 4).setValues([
      ['日期 (YYYY-MM-DD)', '名稱', '類型', '備註']
    ]);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#fce8e6');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 140);
    sheet.setColumnWidth(2, 180);
    sheet.setColumnWidth(3, 100);
    sheet.setColumnWidth(4, 260);

    sheet.getRange(2, 1, 4, 4).setValues([
      ['2026-02-14', '情人節', 'festival', '鮮花、朱古力、餐廳高峯'],
      ['2026-10-31', '萬聖節', 'festival', '糖果、裝飾品高峯'],
      ['2026-11-11', '雙11',   'promo',    '促銷高峯'],
      ['2026-12-24', '平安夜', 'festival', '節日食品、禮盒高峯']
    ]);
  }
  return sheet;
}

function initBusyDaysSheet() {
  getBusyDaysSheet(true);
  Logger.log('✅ _BusyDays 已建立（含 4 個示例）');
  Logger.log('如需預填未來 5 年節日，執行 prefillBusyDays()');
  return {ok: true};
}

function getCustomBusyDays() {
  var sheet = getBusyDaysSheet(false);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  var result = [];
  rows.forEach(function(r) {
    var raw = safeStr(r[0]).trim();
    var name = safeStr(r[1]).trim();
    var type = safeStr(r[2]).trim() || 'busy';
    var note = safeStr(r[3]).trim();
    if (!raw || !name) return;

    var date = raw;
    if (/^\d{8}$/.test(raw)) {
      date = raw.substring(0,4) + '-' + raw.substring(4,6) + '-' + raw.substring(6,8);
    }
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return;

    result.push({ date: date, name: name, type: type, note: note });
  });
  result.sort(function(a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
  return result;
}

/**
 * 求某年某月第 n 個星期 X 的日期
 * @param {number} year    年
 * @param {number} month   月（1-12）
 * @param {number} weekday 星期（0=日, 1=一, ..., 6=六）
 * @param {number} n       第 n 個
 * @returns {string} YYYY-MM-DD
 */
function nthWeekdayOfMonth(year, month, weekday, n) {
  var first = new Date(year, month - 1, 1);
  var firstWd = first.getDay();
  var diff = (weekday - firstWd + 7) % 7;
  var day = 1 + diff + (n - 1) * 7;
  return year + '-' + pad2(month) + '-' + pad2(day);
}

/**
 * 冬至（天文現象，2000–2099 用近似公式）
 */
function getWinterSolstice(year) {
  if (year < 2000 || year > 2099) return '';
  var Y = year - 2000;
  var D = Math.floor(Y * 0.2422 + 21.94) - Math.floor(Y / 4);
  return year + '-12-' + pad2(D);
}

/**
 * ★ 預填未來 5 年繁忙日（固定 + 浮動 + 冬至）
 * 已存在的日期不會重複加
 */
function prefillBusyDays() {
  var sheet = getBusyDaysSheet(true);
  var existing = getCustomBusyDays();
  var existingDates = {};
  existing.forEach(function(b) { existingDates[b.date] = true; });

  var now = new Date();
  var startYear = now.getFullYear();
  var endYear = startYear + 4;
  var toAdd = [];

  for (var y = startYear; y <= endYear; y++) {
    // 固定日期
    toAdd.push([y + '-01-01', '元旦',   'festival', '']);
    toAdd.push([y + '-02-14', '情人節', 'festival', '鮮花、朱古力、餐廳高峯']);
    toAdd.push([y + '-10-31', '萬聖節', 'festival', '糖果、裝飾品高峯']);
    toAdd.push([y + '-11-11', '雙11',   'promo',    '促銷高峯']);
    toAdd.push([y + '-12-24', '平安夜', 'festival', '節日食品、禮盒高峯']);
    toAdd.push([y + '-12-25', '聖誕節', 'festival', '節日食品、禮盒高峯']);
    toAdd.push([y + '-12-31', '除夕',   'festival', '派對食品、飲品高峯']);

    // 浮動日期
    var mother = nthWeekdayOfMonth(y, 5, 0, 2);
    var father = nthWeekdayOfMonth(y, 6, 0, 3);
    toAdd.push([mother, '母親節', 'festival', '禮物、餐廳高峯']);
    toAdd.push([father, '父親節', 'festival', '禮物、餐廳高峯']);

    // 冬至
    var solstice = getWinterSolstice(y);
    if (solstice) {
      toAdd.push([solstice, '冬至', 'festival', '做冬、團年飯、湯圓高峯']);
    }
  }

  var newRows = [];
  toAdd.forEach(function(r) {
    if (!existingDates[r[0]]) newRows.push(r);
  });

  if (newRows.length === 0) {
    Logger.log('所有繁忙日已存在，不需新增');
    return { ok: true, added: 0 };
  }

  var startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, newRows.length, 4).setValues(newRows);
  Logger.log('✅ 已新增 ' + newRows.length + ' 個繁忙日（' + startYear + '-' + endYear + '）');
  return { ok: true, added: newRows.length };
}


// ============================================================
//  ★ v5.4：公眾假期（1823 API）+ 繁忙日合併端點
// ============================================================
const HOLIDAY_CACHE_KEY = 'hk_holidays_v1';
const HOLIDAY_CACHE_TTL = 24 * 60 * 60;

function getHKHolidaysRaw() {
  var cached = cacheGet(HOLIDAY_CACHE_KEY);
  if (cached && cached.holidays) return cached.holidays;

  var holidays = [];
  try {
    var res = UrlFetchApp.fetch(
      'https://www.1823.gov.hk/common/ical/tc.json',
      { muteHttpExceptions: true, followRedirects: true }
    );
    if (res.getResponseCode() !== 200) {
      throw new Error('HTTP ' + res.getResponseCode());
    }
    var json = JSON.parse(res.getContentText('UTF-8'));
    var cal = json.vcalendar || [];
    cal.forEach(function(item) {
      var dates = item.dtstart || [];
      if (!Array.isArray(dates)) dates = [dates];
      var name = item.summary || '';
      dates.forEach(function(d) {
        var ds = String(d).substring(0, 10);
        if (/^\d{4}-\d{2}-\d{2}$/.test(ds)) {
          holidays.push({ date: ds, name: name, type: 'public' });
        }
      });
    });
    cachePut(HOLIDAY_CACHE_KEY, { holidays: holidays }, HOLIDAY_CACHE_TTL);
  } catch (e) {
    Logger.log('1823 API 失敗：' + e.message);
  }
  return holidays;
}

function getHolidayData() {
  return {
    ok: true,
    updated: new Date().toISOString(),
    holidays: getHKHolidaysRaw(),
    busyDays: getCustomBusyDays()
  };
}


// ============================================================
//  選項管理（_Config）
// ============================================================
function initConfigSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(CONFIG_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG_SHEET);
  sheet.clear();

  var config = [
    ['選項類別', '選項值', '顯示名稱'],
    ['warehouse', 'A倉', 'A倉'],
    ['warehouse', 'B倉', 'B倉'],
    ['warehouse', 'C倉', 'C倉'],
    ['warehouse', 'D倉', 'D倉'],
    ['warehouse', '臨時倉', '臨時倉']
  ];
  for (var i = 1; i <= 20; i++) {
    var n = ('0' + i).slice(-2);
    config.push(['aisle', n, n + ' 道']);
  }
  'ABCDEFGH'.split('').forEach(function(c) { config.push(['crosslane', c, c + ' 巷']); });
  ['1','2','3','4','5'].forEach(function(n) { config.push(['floor', n, '第 ' + n + ' 層']); });
  ['涼爽','冷凍','常溫','特別處理'].forEach(function(v) { config.push(['combine', v, v]); });

  var cats = [
    ['Beer, Wines & Spirits', '啤酒、烈酒'],
    ['Beverages', '飲品'],
    ['Frozen Food', '急凍食品'],
    ['Breakfast & Bakery', '早餐及烘焙'],
    ['Dairy, Chilled & Eggs', '奶類、冷藏及蛋'],
    ['Canned, Preserved & Soup', '罐頭、醃製及湯'],
    ['Condiment & Sauce', '調味醬料'],
    ['Fruits & Vegetables', '蔬果'],
    ['Health Care', '健康護理'],
    ['Kitchen & Dining', '廚房及餐具'],
    ['Meat & Seafood', '肉類及海鮮'],
    ['Mother & Baby', '母嬰'],
    ['Rice, Oil & Noodles', '米、油、麵'],
    ['Snacks & Confectionery', '零食及糖果'],
    ['Feminine Care', '女性護理'],
    ['Festive items', '節日食品']
  ];
  cats.forEach(function(pair) { config.push(['category_en', pair[0], pair[1]]); });

  sheet.getRange(1, 1, config.length, 3).setValues(config);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#e8f0fe');
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 200);
  sheet.setFrozenRows(1);

  return {ok: true, count: config.length - 1};
}

function getOptions() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(CONFIG_SHEET);
  if (!sheet || sheet.getLastRow() < 2) {
    initConfigSheet();
    sheet = ss.getSheetByName(CONFIG_SHEET);
  }
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  var options = {
    warehouse: [], aisle: [], crosslane: [], floor: [],
    combine: [], category_en: [],
    oos_reason: getOOSReasons(),
    oos_action: getOOSActions()          // #53：跟進動作選項
  };
  rows.forEach(function(row) {
    if (!isValidRow(row)) return;
    var type  = safeStr(row[0]).trim();
    var value = safeStr(row[1]).trim();
    var label = safeStr(row[2] != null ? row[2] : value).trim();
    if (!type || !value) return;
    if (type === 'oos_reason') return;
    if (options[type]) options[type].push({value: value, label: label});
  });
  return options;
}

function saveOptions(body) {
  var options = body && body.options;
  if (!isObj(options)) return {error: 'Missing options'};

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(CONFIG_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG_SHEET);
  sheet.clear();

  var TYPES = ['warehouse', 'aisle', 'crosslane', 'floor', 'combine', 'category_en'];
  var config = [['選項類別', '選項值', '顯示名稱']];

  TYPES.forEach(function(type) {
    var list = options[type];
    if (Array.isArray(list)) {
      list.forEach(function(opt) {
        if (!isObj(opt)) return;
        var v = safeStr(opt.value).trim();
        var l = safeStr(opt.label || v).trim();
        if (v) config.push([type, v, l]);
      });
    }
  });

  sheet.getRange(1, 1, config.length, 3).setValues(config);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#e8f0fe');
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 200);
  sheet.setFrozenRows(1);

  return {ok: true, count: config.length - 1};
}


// ============================================================
//  OOS API
// ============================================================
function reportOOS(body) {
  if (!isObj(body)) return {error: 'Invalid body'};
  var productId   = safeStr(body.productId).trim();
  var reason      = safeStr(body.reason).trim();
  var productName = safeStr(body.productName);
  var reporter    = safeStr(body.reporter) || 'user';

  var orderDate = safeStr(body.orderDate).trim();
  var orderTail = safeStr(body.orderTail).trim();
  var timeSlot  = safeStr(body.timeSlot).trim();
  var orderType = safeStr(body.orderType).trim();

  // #53：跟進動作 —— 每次情況唔同，所以逐筆記錄，可以留空。
  // ⚠️ 一定要用 body.oosAction，唔可以讀 body.action —— 後者係派發用嘅動作名。
  var action = safeStr(body.oosAction).trim();
  if (action.length > 40) action = action.substring(0, 40);

  // #18：缺貨件數（可選，預設 1）。同一件貨多張單會累加。
  var qtyRaw = safeStr(body.qty).trim();
  var qty = qtyRaw === '' ? 1 : parseInt(qtyRaw.replace(/[^0-9]/g, ''), 10);
  // #60：0 件係合法 —— 行架巡查見到缺貨但未知會缺幾多時用 0
  if (isNaN(qty) || qty < 0) qty = 1;
  if (qty > 9999) qty = 9999;

  if (!productId) return {error: 'Missing productId'};
  if (!reason) return {error: '必須填寫缺貨原因'};
  if (reason.length < 2) return {error: '缺貨原因太短（至少 2 個字）'};

  if (orderDate && !/^\d{8}$/.test(orderDate)) {
    return {error: '日期格式錯誤，需 YYYYMMDD（8 位數字）'};
  }
  if (orderTail && !/^\d{1,4}$/.test(orderTail)) {
    return {error: '單號尾需為 1~4 位數字'};
  }
  if (timeSlot && ['早','中','晚','放假'].indexOf(timeSlot) < 0) {
    return {error: '派送時段需為 早 / 中 / 晚 / 放假'};
  }
  if (orderType && ['O2O','CNC','B2B'].indexOf(orderType) < 0) {
    return {error: '單類型需為 O2O / CNC / B2B'};
  }

  // #18：每次回報都「新增一列」，不再覆寫舊記錄。
  //      同一件貨若有多張單都缺貨，各自記錄件數，
  //      再由 getActiveOOSMap() 把同一產品的 active 列累加出總數。
  //      （補貨 restockOOS 仍會一次把該產品所有 active 列結案。）
  var sheet = getOOSSheet(true);

  // #19：單號尾4 必須存成文字，否則 0 開頭的單號（"0123"）會變成 123。
  //      寫入前先把該列的格式鎖成純文字，不依賴試算表現有格式，
  //      所以就算之後新增列數也一定正確。
  var rowNum = sheet.getLastRow() + 1;
  sheet.getRange(rowNum, 9).setNumberFormat('@');
  sheet.getRange(rowNum, 1, 1, 13).setValues([[
    new Date().toISOString(),
    productId,
    productName,
    reporter,
    reason,
    'active',
    '',
    orderDate,
    orderTail,
    timeSlot,
    orderType,
    qty,
    action          // #53
  ]]);

  scheduleExport();

  // 回傳累計總數，前端可即時顯示（不必等下一次同步）
  var agg = getActiveOOSMap()[productId] || null;
  return {
    ok: true,
    created: true,
    productId: productId,
    qty: qty,
    totalQty: agg ? agg.qty : qty,
    reports: agg ? agg.reports : 1
  };
}

function restockOOS(body) {
  if (!isObj(body)) return {error: 'Invalid body'};
  var productId = safeStr(body.productId).trim();
  if (!productId) return {error: 'Missing productId'};

  var sheet = getOOSSheet(false);
  if (!sheet) return {error: '沒有 OOS 記錄'};

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {error: '沒有 OOS 記錄'};

  var rows = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  var found = 0;
  for (var i = 0; i < rows.length; i++) {
    if (!isValidRow(rows[i])) continue;
    var rid = safeStr(rows[i][1]).trim();
    var rstatus = safeStr(rows[i][5]).trim();
    if (rid === productId && rstatus === 'active') {
      sheet.getRange(i + 2, 6).setValue('restocked');
      sheet.getRange(i + 2, 7).setValue(new Date().toISOString());
      found++;
    }
  }

  if (found === 0) return {error: '此產品沒有 active 缺貨記錄'};
  scheduleExport();
  return {ok: true, restocked: found, productId: productId};
}


// ============================================================
//  快取
// ============================================================
function cacheGet(key) {
  var cache = CacheService.getScriptCache();
  var info = cache.get(key + '_info');
  if (!info) return null;
  var chunks = parseInt(info);
  var keys = [];
  for (var i = 0; i < chunks; i++) keys.push(key + '_' + i);
  var parts = cache.getAll(keys);
  var full = '';
  for (var j = 0; j < chunks; j++) {
    var part = parts[key + '_' + j];
    if (!part) return null;
    full += part;
  }
  try { return JSON.parse(full); } catch (e) { return null; }
}

function cachePut(key, data, seconds) {
  var cache = CacheService.getScriptCache();
  var full = JSON.stringify(data);
  var chunks = Math.ceil(full.length / CACHE_CHUNK_SIZE);
  if (chunks > 50) return false;
  var entries = {};
  for (var i = 0; i < chunks; i++) {
    entries[key + '_' + i] = full.substr(i * CACHE_CHUNK_SIZE, CACHE_CHUNK_SIZE);
  }
  entries[key + '_info'] = String(chunks);
  try { cache.putAll(entries, seconds); return true; }
  catch (e) { Logger.log('cachePut failed: ' + e.message); return false; }
}

function cacheRemove(key) {
  var cache = CacheService.getScriptCache();
  var info = cache.get(key + '_info');
  if (info) {
    var chunks = parseInt(info);
    var keys = [key + '_info'];
    for (var i = 0; i < chunks; i++) keys.push(key + '_' + i);
    cache.removeAll(keys);
  }
}

function invalidateCache() { cacheRemove(CACHE_KEY); }


// ============================================================
//  pHash 鎖
// ============================================================
function acquirePhashLock() {
  var props = PropertiesService.getScriptProperties();
  var now = Date.now();
  var lockUntil = parseInt(props.getProperty('PHASH_LOCK_UNTIL') || '0');
  if (now < lockUntil) return {ok: false, remainSec: Math.ceil((lockUntil - now) / 1000)};
  props.setProperty('PHASH_LOCK_UNTIL', String(now + 30 * 60 * 1000));
  return {ok: true};
}

function releasePhashLock() {
  PropertiesService.getScriptProperties().deleteProperty('PHASH_LOCK_UNTIL');
  return {ok: true};
}


// ============================================================
//  CDN 匯出
// ============================================================
function exportToGitHub() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('GITHUB_TOKEN');
  var repo  = props.getProperty('GITHUB_REPO') || 'charlizelai/product-search';
  if (!token) throw new Error('未設定 GITHUB_TOKEN');

  var products = readProductsFromSheet();
  var oosMap = getActiveOOSMap();

  var items = products
    .filter(function(p) { return isObj(p) && p.id; })
    .map(function(p) {
      var item = {
        id: safeStr(p.id), barcode: safeStr(p.barcode),
        name: safeStr(p.name), name_zh: safeStr(p.name_zh),
        price: safeStr(p.price), original_price: safeStr(p.original_price),
        image: safeStr(p.image),
        category_en: safeStr(p.category_en), category_zh: safeStr(p.category_zh),
        description: safeStr(p.description), brand: safeStr(p.brand),
        sku: safeStr(p.sku),
        temp_location: safeStr(p.temp_location), combine: safeStr(p.combine),
        warehouse: safeStr(p.warehouse), island: safeStr(p.island),
        aisle: safeStr(p.aisle), crosslane: safeStr(p.crosslane),
        shelfnumber: safeStr(p.shelfnumber), floor: safeStr(p.floor),
        _last_modified: safeStr(p._last_modified),
        _last_source: safeStr(p._last_source)
      };
      if (isTempId(p.id)) item._temp_id = true;
      var oos = oosMap[String(p.id)];
      if (oos) {
        item._oos = true;
        item._oos_reason = oos.reason;
        item._oos_at = oos.at;
        item._oos_reporter = oos.reporter;
        item._oos_orderDate = oos.orderDate;
        item._oos_orderTail = oos.orderTail;
        item._oos_timeSlot = oos.timeSlot;
        item._oos_orderType = oos.orderType;
        item._oos_qty = oos.qty;
        item._oos_reports = oos.reports;
      }
      return item;
    });

  var payload = {
    version: Date.now(),
    updated: new Date().toISOString(),
    count: items.length,
    items: items
  };

  var jsonStr = JSON.stringify(payload);
  var encoded = Utilities.base64Encode(jsonStr, Utilities.Charset.UTF_8);
  var apiBase = 'https://api.github.com/repos/' + repo + '/contents/' + GITHUB_DATA_PATH;
  var headers = {
    'Authorization': 'Bearer ' + token,
    'Accept': 'application/vnd.github+json',
    'X-GitHub-Api-Version': '2022-11-28'
  };

  var sha = null;
  try {
    var r = UrlFetchApp.fetch(apiBase + '?ref=' + GITHUB_BRANCH, { headers: headers, muteHttpExceptions: true });
    if (r.getResponseCode() === 200) sha = JSON.parse(r.getContentText()).sha;
  } catch (e) {}

  var body = {
    message: 'Auto-export (' + items.length + ' items) - ' + new Date().toISOString(),
    content: encoded,
    branch: GITHUB_BRANCH
  };
  if (sha) body.sha = sha;

  var putRes = UrlFetchApp.fetch(apiBase, {
    method: 'put', headers: headers,
    contentType: 'application/json',
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  var code = putRes.getResponseCode();
  if (code === 200 || code === 201) {
    return {ok: true, count: items.length, size: Math.round(jsonStr.length / 1024) + ' KB'};
  }
  throw new Error('GitHub 匯出失敗 HTTP ' + code + ': ' + putRes.getContentText().substring(0, 300));
}

function scheduleExport() {
  var props = PropertiesService.getScriptProperties();
  var now = Date.now();
  var last = parseInt(props.getProperty('LAST_EXPORT') || '0');
  if (now - last < 5000) return;
  props.setProperty('LAST_EXPORT', String(now));
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'exportToGitHub' && t.getEventType() === ScriptApp.EventType.CLOCK) {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('exportToGitHub').timeBased().after(3000).create();
}

function manualExport() {
  var result = exportToGitHub();
  Logger.log('✅ 匯出成功：' + JSON.stringify(result));
  return result;
}

function createDailyExportTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'exportToGitHub' && t.getEventType() === ScriptApp.EventType.CLOCK) {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('exportToGitHub').timeBased().everyHours(1).create();
  Logger.log('✅ 已建立每小時匯出觸發');
}




// ============================================================
//  #43  自動跟隨惠康推廣品牌
// ============================================================
/*
 * 做法：
 *   1. 抓惠康首頁 → 抽出當期主推嘅商品 ID（實測約 177 個）
 *   2. 用試算表已有嘅 brand 欄對返品牌
 *   3. 按出現次數排序 → 存起頭 N 個
 *   4. 前端品牌下拉：手動釘住 → 惠康推廣 → 其餘按產品數量
 *
 * 成本：每次更新只需 1 個網頁請求 + 2 個單欄讀取，非常輕。
 * 失敗保護：抽唔到商品就保留上一次嘅清單，唔會清空。
 */

const PROMO_BRAND_LIMIT = 12;   // 自動推廣品牌最多保留幾個

/** 抓惠康首頁，抽出商品 ID（失敗回空陣列） */
function fetchWellcomeHomepageIds_() {
  try {
    var res = UrlFetchApp.fetch('https://www.wellcome.com.hk/zh-hant/', {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept-Language': 'zh-HK,zh;q=0.9'
      }
    });
    if (res.getResponseCode() !== 200) return [];
    var html = res.getContentText('UTF-8');
    var re = /\/i\/(\d+)\.html/g, m, set = {};
    while ((m = re.exec(html)) !== null) { set[m[1]] = true; }
    return Object.keys(set);
  } catch (e) {
    Logger.log('fetchWellcomeHomepageIds_ 失敗：' + e.message);
    return [];
  }
}

/**
 * 更新推廣品牌清單（可手動執行，或由每日觸發器呼叫）。
 * 抽唔到首頁商品 → 保留舊清單，唔會清空。
 */
function updatePromoBrands() {
  var ids = fetchWellcomeHomepageIds_();
  if (!ids.length) {
    wsuLog_('PROMO_SKIP', '首頁抽唔到商品，保留上一次清單');
    return { ok: false, error: '首頁抽唔到商品' };
  }

  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: false, error: '試算表冇資料' };

  // 只讀兩條單欄，唔讀整行（快好多）
  var idCol = sheet.getRange(2, COL_MAP.id, lastRow - 1, 1).getValues();
  var brCol = sheet.getRange(2, COL_MAP.brand, lastRow - 1, 1).getValues();

  var byId = {};
  for (var i = 0; i < idCol.length; i++) {
    var pid = safeStr(idCol[i][0]).trim();
    if (pid) byId[pid] = safeStr(brCol[i][0]).trim();
  }

  var counts = {};
  var matched = 0;
  for (var k = 0; k < ids.length; k++) {
    var b = byId[ids[k]];
    if (b) { counts[b] = (counts[b] || 0) + 1; matched++; }
  }

  var list = Object.keys(counts);
  list.sort(function (x, y) { return counts[y] - counts[x]; });
  list = list.slice(0, PROMO_BRAND_LIMIT);

  var props = PropertiesService.getScriptProperties();
  props.setProperty('PROMO_BRANDS', JSON.stringify(list));
  props.setProperty('PROMO_BRANDS_AT', new Date().toISOString());
  props.setProperty('PROMO_IDS', String(ids.length));

  return { ok: true, homepageProducts: ids.length, matched: matched, brands: list };
}

/** 前端讀取用 */
function getPromoBrands() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty('PROMO_BRANDS');
  var list = [];
  if (raw) { try { list = JSON.parse(raw) || []; } catch (e) { list = []; } }
  return {
    ok: true,
    brands: Array.isArray(list) ? list : [],
    updated: props.getProperty('PROMO_BRANDS_AT') || '',
    sourceProducts: parseInt(props.getProperty('PROMO_IDS') || '0', 10) || 0
  };
}

/** 建立每日 05:30 更新推廣品牌嘅觸發器（跑一次就得） */
function createPromoBrandsTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'updatePromoBrands') { ScriptApp.deleteTrigger(t); }
  });
  ScriptApp.newTrigger('updatePromoBrands').timeBased().everyDays(1).atHour(5).nearMinute(30).create();
  Logger.log('已建立每日 05:30 更新推廣品牌觸發器');
}

// ============================================================
//  GET
// ============================================================
function doGet(e) {
  try {
    var action = e.parameter.action;
    var token  = e.parameter.token;

    if (action === 'ping') {
      return json({ok: true, version: API_VERSION, time: new Date().toISOString()});
    }

    if (!isAuthorized(token)) return json({error: 'Unauthorized'});

    ensureColumns();

    switch (action) {
      case 'list':        return json(listProducts());
      case 'get':         return json(getProduct(e.parameter.id));
      case 'options':     return json(getOptions());
      case 'phashStatus': return json(getPhashStatus());
      case 'phashLock':
        if (!isAdminToken(token)) return json({error: 'Unauthorized admin'});
        return json(acquirePhashLock());
      case 'phashUnlock':
        if (!isAdminToken(token)) return json({error: 'Unauthorized admin'});
        return json(releasePhashLock());
      case 'oosList':     return json(getActiveOOSMap());
      case 'oosHistory':  return json(getProductOOSHistory(e.parameter.id, 5));
      case 'oosReports':  return json(getOOSReports(e.parameter.date || ''));
      case 'orderTotals': return json(getOrderTotals());
      case 'substitutes': return json(getSubstitutes());
      case 'tempIds':     return json(getTempIdProducts());
      case 'holidays':    return json(getHolidayData());
      case 'me':          return json({ ok: true, user: currentUser_(token) });   // #38
      case 'promoBrands': return json(getPromoBrands());                          // #43
      default:            return json({error: 'Unknown action: ' + action, version: API_VERSION});
    }
  } catch (err) {
    return json({error: err.message, version: API_VERSION});
  }
}


// ============================================================
//  POST
// ============================================================
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    var token  = body.token;

    // #38：公開動作（未登入都用得）—— 必須喺權限檢查之前
    if (action === 'register') return json(withLock(function() { return registerUser(body); }));
    // login 亦會寫入（last_login / session / 清理過期），一齊排鎖避免同時刪列出錯
    if (action === 'login')    return json(withLock(function() { return loginUser(body); }));

    if (!isAuthorized(token)) return json({error: 'Unauthorized'});

    ensureColumns();

    // 需要管理員權限的「破壞性 / 維護性」操作
    var ADMIN_ONLY = ['delete', 'clearPhash', 'initConfig', 'initOOSReasons',
                      'initBusyDays', 'prefillBusyDays', 'saveOptions', 'exportNow',
                      // #38 員工管理
                      'listUsers', 'setUserActive', 'setUserRole',
                      'resetUserPassword', 'setInviteCodeNow',
                      // #45 一次改幾百行分類，只限管理員
                      'bulkFixCategory'];
    if (ADMIN_ONLY.indexOf(action) >= 0 && !isAdminToken(token)) {
      return json({error: 'Unauthorized admin'});
    }

    // 所有會寫入的操作一律經過 withLock()，避免並發列位移寫錯產品
    switch (action) {
      case 'create':         return json(withLock(function() { return createProduct(body.data); }));
      case 'update':         return json(withLock(function() { return updateProduct(body.id, body.data); }));
      case 'delete':         return json(withLock(function() { return deleteProduct(body.id); }));
      case 'unlock':         return json(withLock(function() { return unlockProduct(body.id); }));
      case 'uploadImage':    return json(withLock(function() { return uploadProductImage(body); }));
      case 'savePhashBatch': return json(withLock(function() { return savePhashBatch(body); }));
      case 'bulkFixCategory':return json(withLock(function() { return bulkFixCategory(body.items); }));
      case 'clearPhash':     return json(withLock(function() { return clearPhashLibrary(); }));
      case 'prefillBusyDays':return json(withLock(function() { return prefillBusyDays(); }));
      case 'reportOOS':      return json(withLock(function() { return reportOOS(body); }));
      case 'setOOSAction':   return json(withLock(function() { return setOOSAction(body); }));
      case 'updateOOSRecord':return json(withLock(function() { return updateOOSRecord(body); }));
      case 'saveOrderTotals': return json(withLock(function() { return saveOrderTotals(body); }));
      case 'saveSubstitute':  return json(withLock(function() { return saveSubstitute(body); }));
      case 'restockOOS':     return json(withLock(function() { return restockOOS(body); }));
      case 'saveOptions':    return json(withLock(function() { return saveOptions(body); }));

      // #38 登入 / 登出 / 員工管理
      case 'logout':            return json(logoutUser(body));
      case 'listUsers':         return json(listUsers());
      case 'setUserActive':     return json(withLock(function() { return setUserActive(body); }));
      case 'setUserRole':       return json(withLock(function() { return setUserRole(body); }));
      case 'resetUserPassword': return json(withLock(function() { return resetUserPassword(body); }));
      case 'setInviteCodeNow':  return json(withLock(function() { return { ok: true, inviteCode: setInviteCode(body.code) }; }));

      // 只讀 / 純計算，不需要鎖
      case 'searchByPhash':  return json(searchByPhash(body));
      case 'initConfig':     return json(initConfigSheet());
      case 'initOOSReasons': return json(initOOSReasonsSheet());
      case 'initBusyDays':   return json(initBusyDaysSheet());
      case 'exportNow':      return json(manualExport());

      default:               return json({error: 'Unknown action: ' + action, version: API_VERSION});
    }
  } catch (err) {
    return json({error: err.message, version: API_VERSION});
  }
}


// ============================================================
//  List
// ============================================================
function listProducts() {
  var cached = cacheGet(CACHE_KEY);
  if (cached) return cached;

  var lock = LockService.getScriptLock();
  var locked = false;
  try { locked = lock.tryLock(15000); } catch (e) { locked = false; }

  if (!locked) {
    Utilities.sleep(800);
    cached = cacheGet(CACHE_KEY);
    if (cached) return cached;
    return readProductsFromSheet();
  }

  try {
    cached = cacheGet(CACHE_KEY);
    if (cached) return cached;
    var result = readProductsFromSheet();
    cachePut(CACHE_KEY, result, CACHE_TTL);
    return result;
  } finally {
    lock.releaseLock();
  }
}

function readProductsFromSheet() {
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var rows = sheet.getRange(2, 1, lastRow - 1, MAX_COL).getValues();

  var products = rows
    .filter(function(r) { return isValidRow(r); })
    .map(function(r) { return rowToProduct(r); })
    .filter(function(p) { return p && p.id; });

  var oosMap = getActiveOOSMap();
  products.forEach(function(p) {
    var oos = oosMap[String(p.id)];
    if (oos) {
      p._oos = true;
      p._oos_reason = oos.reason;
      p._oos_at = oos.at;
      p._oos_reporter = oos.reporter;
      p._oos_orderDate = oos.orderDate;
      p._oos_orderTail = oos.orderTail;
      p._oos_timeSlot = oos.timeSlot;
      p._oos_orderType = oos.orderType;
      p._oos_qty = oos.qty;
      p._oos_reports = oos.reports;
    }
  });

  return products;
}


// ============================================================
//  Get
// ============================================================
function getProduct(id) {
  if (!id) return {error: 'Missing id'};
  var cached = cacheGet(CACHE_KEY);
  if (cached && Array.isArray(cached)) {
    var found = cached.find(function(p) {
      return p && safeStr(p.id).trim() === String(id).trim();
    });
    if (found) return found;
  }
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {error: 'Not found'};
  var rows = sheet.getRange(2, 1, lastRow - 1, MAX_COL).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (!isValidRow(rows[i])) continue;
    if (safeStr(rows[i][0]).trim() === String(id).trim()) return rowToProduct(rows[i]);
  }
  return {error: 'Not found'};
}


// ============================================================
//  找行
// ============================================================
function findRow(sheet, id, barcode) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var rows = sheet.getRange(2, 1, lastRow - 1, MAX_COL).getValues();
  var barcodeStr = safeStr(barcode).trim();
  var idStr = safeStr(id).trim();

  if (barcodeStr) {
    var codes = barcodeStr.split(/[;,，、]/).map(function(s) { return s.trim(); }).filter(function(s) { return s; });
    for (var i = 0; i < rows.length; i++) {
      if (!isValidRow(rows[i])) continue;
      var cellBc = safeStr(rows[i][1]);
      var cellCodes = cellBc.split(/[;,，、]/).map(function(s) { return s.trim(); });
      for (var k = 0; k < codes.length; k++) {
        if (cellCodes.indexOf(codes[k]) >= 0) return i + 2;
      }
    }
  }
  if (idStr) {
    for (var j = 0; j < rows.length; j++) {
      if (!isValidRow(rows[j])) continue;
      if (safeStr(rows[j][0]).trim() === idStr) return j + 2;
    }
  }
  return -1;
}


// ============================================================
//  CRUD
// ============================================================
function createProduct(data) {
  if (!isObj(data)) return {error: 'Missing or invalid data'};
  if (!data.id) return {error: 'Missing id'};

  var sheet = getSheet();
  var existingRow = findRow(sheet, data.id, data.barcode);
  if (existingRow > 0) {
    var r = updateProductInternal(sheet, existingRow, data, 'user');
    invalidateCache();
    scheduleExport();
    return r;
  }

  var tempId = isTempId(data.id);
  var source = tempId ? SOURCE_SCAN_AUTO : 'user';

  var row = productToRow(data, new Date().toISOString(), source);
  sheet.appendRow(row);

  if (tempId) {
    try {
      var ss = SpreadsheetApp.openById(SHEET_ID);
      var log = ss.getSheetByName('_WeeklyUpdateLog');
      if (!log) {
        log = ss.insertSheet('_WeeklyUpdateLog');
        log.appendRow(['時間', '事件', '詳情']);
        log.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#e8f0fe');
        log.setFrozenRows(1);
      }
      log.appendRow([new Date().toISOString(), 'TEMP_ID_CREATED',
        'id=' + data.id + ' barcode=' + (data.barcode || '') + ' name=' + (data.name_zh || data.name || '')]);
    } catch (e) {}
  }

  invalidateCache();
  scheduleExport();
  return {ok: true, id: data.id, action: 'created', isTempId: tempId};
}

function updateProduct(id, data) {
  if (!id) return {error: 'Missing id'};
  if (!isObj(data)) return {error: 'Missing or invalid data'};

  var sheet = getSheet();
  var existingRow = findRow(sheet, id, data.barcode);
  if (existingRow < 0) return {error: 'Not found: ' + id};
  var r = updateProductInternal(sheet, existingRow, data, 'user');
  invalidateCache();
  scheduleExport();
  return r;
}

function updateProductInternal(sheet, rowNum, data, source) {
  if (!isObj(data)) data = {};

  // #37：一次過讀寫整行，唔再逐個欄位 setValue。
  //      原本 20 個欄位 = 20 次 API 往返（每次幾百 ms），
  //      每次儲存要 5-6 秒而且全程霸住 script lock，
  //      令同時儲存嘅人第 3-4 個就 timeout。
  //      改成 1 讀 + 1 寫 → 約 0.5 秒，寫入吞吐提升約 4-5 倍。
  var rng  = sheet.getRange(rowNum, 1, 1, MAX_COL);
  var row  = rng.getValues()[0];
  var fmls = rng.getFormulas()[0];

  ALL_FIELDS.forEach(function(k) {
    if (data[k] !== undefined && k !== 'id') {
      row[COL_MAP[k] - 1] = data[k];
    }
  });
  row[COL_TIMESTAMP - 1] = new Date().toISOString();
  row[COL_SOURCE - 1] = source;

  // 安全檢查：如果該行有任何公式，就唔可以用 setValues（會將公式壓平成數值），
  // 改回逐格寫。正常情況下唔會有公式，所以行快速路徑。
  var hasFormula = false;
  for (var fi = 0; fi < fmls.length; fi++) {
    if (fmls[fi]) { hasFormula = true; break; }
  }

  if (!hasFormula) {
    rng.setValues([row]);                       // 一次過寫成行
  } else {
    for (var ci = 1; ci <= MAX_COL; ci++) {     // 有公式 → 逐格，但跳過公式格
      if (fmls[ci - 1]) continue;
      sheet.getRange(rowNum, ci).setValue(row[ci - 1]);
    }
  }
  return {ok: true, id: data.id || '', action: 'updated'};
}

function deleteProduct(id) {
  if (!id) return {error: 'Missing id'};
  var sheet = getSheet();
  var rowNum = findRow(sheet, id, null);
  if (rowNum < 0) return {error: 'Not found: ' + id};
  sheet.deleteRow(rowNum);
  invalidateCache();
  scheduleExport();
  return {ok: true, id: id};
}

function unlockProduct(id) {
  if (!id) return {error: 'Missing id'};
  var sheet = getSheet();
  var rowNum = findRow(sheet, id, null);
  if (rowNum < 0) return {error: 'Not found: ' + id};
  sheet.getRange(rowNum, COL_SOURCE).setValue('');
  sheet.getRange(rowNum, COL_TIMESTAMP).setValue(new Date().toISOString());
  invalidateCache();
  scheduleExport();
  return {ok: true, id: id};
}


// ============================================================
//  Upload Image
// ============================================================
function uploadProductImage(body) {
  if (!isObj(body)) return {error: 'Invalid body'};
  var productId = body.productId;
  var fileName  = body.fileName;
  var mimeType  = body.mimeType;
  var base64    = body.base64;
  if (!productId || !base64) return {error: 'Missing productId or base64'};

  var bytes = Utilities.base64Decode(base64);
  var blob = Utilities.newBlob(bytes, mimeType || 'image/jpeg',
    fileName || ('photo_' + Date.now() + '.jpg'));

  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty('IMAGE_FOLDER_ID');
  var folder;
  if (folderId) {
    try { folder = DriveApp.getFolderById(folderId); } catch (e) { folderId = null; }
  }
  if (!folderId) {
    folder = DriveApp.createFolder('ProductImages');
    props.setProperty('IMAGE_FOLDER_ID', folder.getId());
  }

  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  var fileId = file.getId();
  var imageUrl = 'https://lh3.googleusercontent.com/d/' + fileId + '=w800';

  var sheet = getSheet();
  var rowNum = findRow(sheet, productId, null);
  if (rowNum < 0) return {error: '找不到 ID: ' + productId};

  sheet.getRange(rowNum, IMAGE_COL_INDEX).setValue(imageUrl);
  sheet.getRange(rowNum, COL_SOURCE).setValue('user');
  sheet.getRange(rowNum, COL_TIMESTAMP).setValue(new Date().toISOString());
  sheet.getRange(rowNum, COL_PHASH).setValue('');

  try {
    var images = sheet.getImages();
    for (var i = 0; i < images.length; i++) {
      var anchor = images[i].getAnchorCell();
      if (anchor.getRow() === rowNum && anchor.getColumn() === IMAGE_COL_INDEX) images[i].remove();
    }
  } catch (e) {}

  var origW = parseInt(body.imgWidth) || IMG_CELL_W;
  var origH = parseInt(body.imgHeight) || IMG_CELL_H;
  if (origW <= 0) origW = IMG_CELL_W;
  if (origH <= 0) origH = IMG_CELL_H;

  var maxDim = Math.max(IMG_CELL_W, IMG_CELL_H);
  var scale = Math.min(maxDim / origW, maxDim / origH);
  if (scale > 1) scale = 1;
  var w = Math.round(origW * scale);
  var h = Math.round(origH * scale);

  try { sheet.setColumnWidth(IMAGE_COL_INDEX, IMG_CELL_W + 10); } catch (e) {}
  try { sheet.setRowHeight(rowNum, Math.max(h, IMG_CELL_H) + 4); } catch (e) {}

  try {
    var img = sheet.insertImage(imageUrl, IMAGE_COL_INDEX, rowNum);
    img.setWidth(w);
    img.setHeight(h);
  } catch (e) {
    Logger.log('insertImage 失敗：' + e.message);
  }

  invalidateCache();
  scheduleExport();
  return {ok: true, productId: productId, fileId: fileId, imageUrl: imageUrl};
}


// ============================================================
//  pHash
// ============================================================
function savePhashBatch(body) {
  var items = body && body.items;
  if (!items || !items.length) return {error: 'No items'};
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {error: 'No data'};
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var idToRow = {};
  ids.forEach(function(r, i) {
    if (!isValidRow(r)) return;
    var id = safeStr(r[0]).trim();
    if (id) idToRow[id] = i + 2;
  });
  var saved = 0;
  items.forEach(function(item) {
    if (!isObj(item) || !item.id) return;
    var rowNum = idToRow[safeStr(item.id).trim()];
    if (rowNum) {
      sheet.getRange(rowNum, COL_PHASH).setValue(item.phash);
      saved++;
    }
  });
  return {ok: true, saved: saved};
}

function searchByPhash(body) {
  var phash = body && body.phash;
  var topN  = body && body.topN;
  if (!phash) return {error: 'Missing phash'};
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {results: [], total: 0};
  var rows = sheet.getRange(2, 1, lastRow - 1, MAX_COL).getValues();
  var scored = [];
  rows.forEach(function(row, i) {
    if (!isValidRow(row)) return;
    var id = safeStr(row[0]).trim();
    var storedHash = row[COL_PHASH - 1];
    if (!id || !storedHash) return;
    scored.push({rowIndex: i, distance: hammingDistance(phash, safeStr(storedHash))});
  });
  scored.sort(function(a, b) { return a.distance - b.distance; });
  var top = scored.slice(0, topN || 12);
  var results = top
    .filter(function(item) { return isValidRow(rows[item.rowIndex]); })
    .map(function(item) {
      var product = rowToProduct(rows[item.rowIndex]);
      product._distance = item.distance;
      product._similarity = Math.round(100 - (item.distance / 64) * 100);
      return product;
    });
  return {results: results, total: scored.length};
}

function getPhashStatus() {
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {total: 0, withPhash: 0, missing: []};
  var rows = sheet.getRange(2, 1, lastRow - 1, MAX_COL).getValues();
  var withPhash = 0;
  var missing = [];
  rows.forEach(function(row) {
    if (!isValidRow(row)) return;
    var id = safeStr(row[0]).trim();
    var image = row[IMAGE_COL_INDEX - 1];
    var phash = row[COL_PHASH - 1];
    if (phash) withPhash++;
    else if (id && image) missing.push(id);
  });
  return {total: rows.length, withPhash: withPhash, missing: missing};
}

function clearPhashLibrary() {
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {ok: true, cleared: 0};
  sheet.getRange(2, COL_PHASH, lastRow - 1, 1).clearContent();
  return {ok: true, cleared: lastRow - 1};
}


// ============================================================
//  一次性工具
// ============================================================
function resizeAllSheetImages() {
  var sheet = getSheet();
  var images = sheet.getImages();
  Logger.log('共 ' + images.length + ' 張圖片');
  for (var i = 0; i < images.length; i++) {
    try {
      images[i].setWidth(IMG_CELL_W);
      images[i].setHeight(IMG_CELL_H);
    } catch (e) {}
  }
  try { sheet.setColumnWidth(IMAGE_COL_INDEX, IMG_CELL_W + 10); } catch (e) {}
  Logger.log('✅ 已將所有圖片縮到 ' + IMG_CELL_W + '×' + IMG_CELL_H);
  return {ok: true, count: images.length};
}

function checkVersion() {
  Logger.log('API_VERSION = ' + API_VERSION);
  Logger.log('has isTempId = ' + (typeof isTempId === 'function'));
  Logger.log('has getTempIdProducts = ' + (typeof getTempIdProducts === 'function'));
  Logger.log('has getHKHolidaysRaw = ' + (typeof getHKHolidaysRaw === 'function'));
  Logger.log('has getCustomBusyDays = ' + (typeof getCustomBusyDays === 'function'));
  Logger.log('has prefillBusyDays = ' + (typeof prefillBusyDays === 'function'));
  Logger.log('has getWinterSolstice = ' + (typeof getWinterSolstice === 'function'));
  Logger.log('isTempId("20261004001") = ' + isTempId("20261004001"));
}

function initOOSSheet() {
  var sheet = getOOSSheet(true);
  Logger.log('✅ _OOS 分頁已就緒（11 欄）');
  return {ok: true};
}