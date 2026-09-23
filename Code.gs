/**
 * E Ink 贈品庫存管理系統 — Google Apps Script 後端
 * 部署為 Web App（執行身分：我；存取權：所有人）— 可在「設定」分頁切換是否需要登入
 *
 * 管理者設定（僅限試算表擁有者，在編輯器執行）：
 *   1. 在「管理者」分頁新增一列，填「名稱」、「啟用」勾選
 *   2. 編輯器選 generateInitCodes → 執行，會在「初始碼」欄產生一次性代碼
 *   3. 把初始碼私下交給該管理者；對方登入後必須設定自己的密碼，初始碼隨即失效
 *   忘記密碼：清空該列「密碼雜湊」「鹽」→ 再執行 generateInitCodes
 *
 * 登入開關：「設定」分頁的「需要登入」取消勾選 = 免登入（任何知道網址的人都能操作），
 *   此時操作者記為「免登入操作者」欄的名稱。勾選後立即改為需要登入。
 * 密碼開關：「登入需要密碼」取消勾選 = 只要輸入「管理者」名單內、已啟用的名稱即可登入。
 *
 * 異動記錄為帳本：不可刪除，登打錯誤請用「沖銷」產生反向記錄。
 */

const SHEET_ITEMS  = '庫存';
const SHEET_LOGS   = '異動記錄';
const SHEET_CONFIG = '設定';
const SHEET_ADMINS = '管理者';
const ITEM_HEADERS  = ['ID','品名','數量','單位','低庫存警示','備註','最後更新','圖片'];
const LOG_HEADERS   = ['時間','品項ID','品名','變動量','變動後數量','備註','操作者','類型','記錄ID','沖銷對象'];
const CONFIG_HEADERS = ['設定項目','值','說明'];
const ADMIN_HEADERS = ['名稱','啟用','初始碼','密碼雜湊','鹽','最後登入','說明'];

const LIMITS = {
  name: 50, unit: 10, note: 300, img: 300, adminName: 20,
  qtyMax: 1000000,
  imgBytes: 3 * 1024 * 1024,
  passMin: 6, passMax: 64
};
const IMG_TYPES = ['image/jpeg','image/png','image/webp'];
const SESSION_TTL = 8 * 3600;      // 登入有效 8 小時
const FAIL_MAX = 5;                // 連續失敗次數上限
const FAIL_LOCK = 15 * 60;         // 鎖定 15 分鐘
const HASH_ROUNDS = 300;
const REVERSIBLE = ['領用','補貨'];
const CONFIG_DEFAULTS = [
  ['圖片資料夾ID', '', '存放贈品圖片的雲端硬碟資料夾；留空會自動建立'],
  ['需要登入', false, '勾選 = 需要管理者登入；取消勾選 = 任何知道網址的人都能操作'],
  ['免登入操作者', 'Kim', '不需登入時，異動記錄上的操作者名稱'],
  ['登入需要密碼', false, '勾選 = 名稱＋密碼；取消勾選 = 只要輸入「管理者」名單內、已啟用的名稱即可登入']
];

/* ============ 僅限擁有者在編輯器執行 ============ */
function initSheets(){
  assertOwner_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hasData = [SHEET_ITEMS, SHEET_LOGS].some(n => { const sh = ss.getSheetByName(n); return sh && sh.getLastRow() > 1; });
  if (hasData) throw new Error('工作表已有資料，已停止初始化，避免覆蓋');

  let items = ss.getSheetByName(SHEET_ITEMS) || ss.insertSheet(SHEET_ITEMS);
  items.clear();
  items.getRange(1,1,1,ITEM_HEADERS.length).setValues([ITEM_HEADERS]);

  let logs = ss.getSheetByName(SHEET_LOGS) || ss.insertSheet(SHEET_LOGS);
  logs.clear();
  logs.getRange(1,1,1,LOG_HEADERS.length).setValues([LOG_HEADERS]);

  configSheet_();
  adminsSheet_();
}

/** 為「啟用、尚未設密碼、沒有初始碼」的管理者產生一次性初始碼 */
function generateInitCodes(){
  assertOwner_();
  const sh = adminsSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) { console.log('管理者分頁沒有資料'); return; }
  const rows = sh.getRange(2, 1, n, ADMIN_HEADERS.length).getValues();
  let made = 0;
  rows.forEach((r, i) => {
    const name = String(r[0] || '').trim();
    if (!name || r[1] !== true || r[3] || r[2]) return;
    sh.getRange(i + 2, 3).setValue(randomCode_(8));
    made++;
  });
  console.log('已產生 ' + made + ' 組初始碼，請到「管理者」分頁查看，私下交給對方');
}

function assertOwner_(){
  const me = Session.getEffectiveUser().getEmail();
  const owner = DriveApp.getFileById(ss_().getId()).getOwner().getEmail();
  if (!me || me !== owner) throw new Error('只有試算表擁有者可以執行');
}

/* ============ 共用工具 ============ */
function ss_(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function itemsSheet_(){ return ss_().getSheetByName(SHEET_ITEMS); }
function logsSheet_(){  return ss_().getSheetByName(SHEET_LOGS); }

function configSheet_(){
  const ss = ss_();
  let sh = ss.getSheetByName(SHEET_CONFIG);
  if (!sh) {
    sh = ss.insertSheet(SHEET_CONFIG);
    sh.getRange(1,1,1,CONFIG_HEADERS.length).setValues([CONFIG_HEADERS]);
    sh.setColumnWidth(1, 140); sh.setColumnWidth(2, 320); sh.setColumnWidth(3, 360);
  }
  return sh;
}
/** 補上缺少的設定列（不覆蓋已有的值） */
function ensureConfigDefaults_(){
  const sh = configSheet_();
  const keys = sh.getDataRange().getValues().map(r => r[0]);
  CONFIG_DEFAULTS.forEach(d => {
    if (keys.indexOf(d[0]) >= 0) return;
    sh.appendRow(d);
    if (typeof d[1] === 'boolean') sh.getRange(sh.getLastRow(), 2).insertCheckboxes();
  });
}
function loginRequired_(){ return getConfig_('需要登入').toUpperCase() === 'TRUE'; }
function passwordRequired_(){ return getConfig_('登入需要密碼').toUpperCase() === 'TRUE'; }
function getConfig_(key){
  const vals = configSheet_().getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) if (vals[i][0] === key) return String(vals[i][1] || '').trim();
  return '';
}
function setConfig_(key, value){
  const sh = configSheet_();
  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (vals[i][0] === key) { sh.getRange(i + 1, 2).setValue(value); return; }
  }
  sh.appendRow([key, value, '']);
}

function adminsSheet_(){
  const ss = ss_();
  let sh = ss.getSheetByName(SHEET_ADMINS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_ADMINS);
    sh.getRange(1,1,1,ADMIN_HEADERS.length).setValues([ADMIN_HEADERS]);
    sh.getRange(2,1,2,ADMIN_HEADERS.length).setValues([
      ['佩芝', true, '', '', '', '', '系統擁有者'],
      ['Kim',  true, '', '', '', '', '']
    ]);
    sh.getRange(2, 2, 50, 1).insertCheckboxes();
    sh.setColumnWidth(4, 120); sh.setColumnWidth(5, 80); sh.setColumnWidth(7, 240);
  }
  return sh;
}

function imageFolder_(){
  const id = getConfig_('圖片資料夾ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* 失效則重建 */ } }
  const parents = DriveApp.getFileById(ss_().getId()).getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const folder = parent.createFolder('eink贈品圖片');
  setConfig_('圖片資料夾ID', folder.getId());
  return folder;
}

function rowsToObjects_(values){
  const head = values.shift();
  return values.filter(r => r[0] !== '').map(r => {
    const o = {};
    head.forEach((h, i) => o[h] = r[i] instanceof Date ? r[i].toISOString() : r[i]);
    return o;
  });
}
function readItems_(){ return rowsToObjects_(itemsSheet_().getDataRange().getValues()); }

function findRow_(id){
  const sh = itemsSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return -1;
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) { if (String(ids[i][0]) === String(id)) return i + 2; }
  return -1;
}

function nextId_(){
  let max = 0;
  readItems_().forEach(it => {
    const m = String(it.ID).match(/^G(\d+)$/);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return 'G' + String(max + 1).padStart(3, '0');
}

function json_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function randomCode_(len){
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Utilities.getUuid());
  let s = '';
  for (let i = 0; i < len; i++) s += chars[(bytes[i] + 256) % chars.length];
  return s;
}

function hash_(pass, salt){
  let h = salt + '|' + pass;
  for (let i = 0; i < HASH_ROUNDS; i++) {
    h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + '|' + salt, Utilities.Charset.UTF_8));
  }
  return h;
}

/* ============ 輸入驗證 ============ */
function userError_(msg, code){ const e = new Error(msg); e.userFacing = true; if (code) e.code = code; return e; }

/** 文字欄位：去頭尾空白、限長度、擋試算表公式 */
function text_(v, max, label, required){
  const s = String(v == null ? '' : v).trim();
  if (required && !s) throw userError_('請填寫' + label);
  if (s.length > max) throw userError_(label + '最多 ' + max + ' 字');
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

/** 整數欄位：限範圍 */
function int_(v, min, max, label){
  const n = Number(v);
  if (!Number.isInteger(n) || n < min || n > max) throw userError_(label + '必須是 ' + min + '～' + max + ' 的整數');
  return n;
}

function id_(v){
  const s = String(v || '');
  if (!/^G\d{3,6}$/.test(s)) throw userError_('品項編號格式錯誤');
  return s;
}

/** 圖片欄：只接受雲端硬碟檔案 ID、https 網址或舊的 images/ 路徑 */
function img_(v){
  const s = String(v || '').trim();
  if (!s) return '';
  if (s.length > LIMITS.img) throw userError_('圖片網址太長');
  const m = s.match(/drive\.google\.com\/(?:file\/d\/|open\?id=|uc\?(?:export=view&)?id=|thumbnail\?id=)([\w-]{20,})/);
  if (m) return m[1];
  if (/^[\w-]{20,}$/.test(s)) return s;
  if (/^images\/[\w.-]+$/.test(s)) return s;
  if (/^https:\/\/[^\s"'<>]+$/.test(s)) return s;
  throw userError_('圖片請用上傳功能，或貼 https 開頭的網址');
}

/* ============ 認證 ============ */
function findAdmin_(name){
  const sh = adminsSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return null;
  const rows = sh.getRange(2, 1, n, ADMIN_HEADERS.length).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === name) return { row: i + 2, r: rows[i] };
  }
  return null;
}

function newSession_(name, mustChange){
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put('s_' + token, JSON.stringify({ n: name, mc: !!mustChange }), SESSION_TTL);
  return token;
}

/** 驗證 session，回傳 { name, mustChange, token } */
function auth_(token){
  const t = String(token || '');
  if (!/^[a-f0-9]{64}$/.test(t)) throw userError_('請先登入', 'AUTH');
  const raw = CacheService.getScriptCache().get('s_' + t);
  if (!raw) throw userError_('登入已逾時，請重新登入', 'AUTH');
  const s = JSON.parse(raw);
  const a = findAdmin_(s.n);
  if (!a || a.r[1] !== true) { CacheService.getScriptCache().remove('s_' + t); throw userError_('帳號已停用', 'AUTH'); }
  return { name: s.n, mustChange: s.mc && passwordRequired_(), token: t };
}

function login_(body){
  const name = String(body.name || '').trim();
  const pass = String(body.passcode || '');
  const needPass = passwordRequired_();
  if (!name || name.length > LIMITS.adminName || (needPass && !pass) || pass.length > LIMITS.passMax) throw userError_(needPass ? '名稱或密碼錯誤' : '名稱不在管理者名單內');
  const cache = CacheService.getScriptCache();
  const failKey = 'f_' + Utilities.base64EncodeWebSafe(name);
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= FAIL_MAX) throw userError_('錯誤次數過多，請 15 分鐘後再試');

  const a = findAdmin_(name);
  let ok = false, mustChange = false;
  if (a && a.r[1] === true && !needPass) ok = true;
  else if (a && a.r[1] === true) {
    const init = String(a.r[2] || ''), hashed = String(a.r[3] || ''), salt = String(a.r[4] || '');
    if (hashed && salt) ok = hash_(pass, salt) === hashed;
    else if (init) { ok = pass === init; mustChange = ok; }
  }
  if (!ok) {
    cache.put(failKey, String(fails + 1), FAIL_LOCK);
    throw userError_(needPass ? '名稱或密碼錯誤' : '名稱不在管理者名單內');
  }
  cache.remove(failKey);
  adminsSheet_().getRange(a.row, 6).setValue(new Date().toISOString());
  return { success: true, token: newSession_(name, mustChange), name, mustChange };
}

function changePasscode_(s, body){
  if (!passwordRequired_()) throw userError_('目前登入不需要密碼');
  const a = findAdmin_(s.name);
  const next = String(body.newPasscode || '');
  if (next.length < LIMITS.passMin || next.length > LIMITS.passMax) throw userError_('新密碼需 ' + LIMITS.passMin + '～' + LIMITS.passMax + ' 字');
  if (!s.mustChange) {
    const cur = String(body.passcode || '');
    if (hash_(cur, String(a.r[4])) !== String(a.r[3])) throw userError_('目前密碼錯誤');
  }
  if (next === String(a.r[2] || '')) throw userError_('新密碼不可與初始碼相同');
  const salt = Utilities.getUuid();
  const sh = adminsSheet_();
  sh.getRange(a.row, 3, 1, 3).setValues([['', hash_(next, salt), salt]]);
  CacheService.getScriptCache().remove('s_' + s.token);
  return { success: true, token: newSession_(s.name, false), name: s.name, mustChange: false };
}

/* ============ 異動記錄（帳本） ============ */
/** 舊資料升級：補上「類型」「記錄ID」「沖銷對象」欄 */
function ensureLogSchema_(){
  const sh = logsSheet_();
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (head.length >= LOG_HEADERS.length && head[8] === '記錄ID') return;
  sh.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const vals = sh.getRange(2, 1, n, LOG_HEADERS.length).getValues();
  const out = vals.map((r, i) => [
    r[7] || (Number(r[3]) < 0 ? '領用' : '補貨'),
    r[8] || 'L' + String(i + 1).padStart(5, '0'),
    r[9] || ''
  ]);
  sh.getRange(2, 8, n, 3).setValues(out);
}

function nextLogId_(){
  return 'L' + String(logsSheet_().getLastRow()).padStart(5, '0');
}

function log_(type, id, name, delta, after, note, operator, reverseOf){
  const logId = nextLogId_();
  logsSheet_().appendRow([ new Date().toISOString(), id, name, delta, after, note, operator, type, logId, reverseOf || '' ]);
  return logId;
}

function readLogs_(){ return rowsToObjects_(logsSheet_().getDataRange().getValues()); }

/* ============ 動作 ============ */
function applyDelta_(id, delta){
  const row = findRow_(id);
  if (row < 0) throw userError_('找不到品項');
  const sh = itemsSheet_();
  const cur = Number(sh.getRange(row, 3).getValue()) || 0;
  const newQty = cur + delta;
  if (newQty < 0) throw userError_('庫存不足');
  if (newQty > LIMITS.qtyMax) throw userError_('數量超過上限');
  sh.getRange(row, 3).setValue(newQty);
  sh.getRange(row, 7).setValue(new Date().toISOString());
  return { newQty, name: sh.getRange(row, 2).getValue() };
}

const ACTIONS = {
  getAll: (s) => ({ items: readItems_(), user: s.name, openMode: !!s.open, passwordMode: passwordRequired_() }),

  getLogs: (s, b) => {
    let logs = readLogs_();
    const reversed = {};
    logs.forEach(l => { if (l['沖銷對象']) reversed[l['沖銷對象']] = l['記錄ID']; });
    logs.forEach(l => { l['已沖銷'] = reversed[l['記錄ID']] || ''; });
    if (b.itemId) logs = logs.filter(l => String(l['品項ID']) === String(b.itemId));
    logs.reverse();
    return { logs: logs.slice(0, 1000) };
  },

  changePasscode: (s, b) => changePasscode_(s, b),

  logout: (s) => { CacheService.getScriptCache().remove('s_' + s.token); return { success: true }; },

  uploadImage: (s, b) => {
    const m = String(b.data || '').match(/^data:([\w/+.-]+);base64,([A-Za-z0-9+/=]+)$/);
    if (!m || IMG_TYPES.indexOf(m[1]) < 0) throw userError_('只接受 JPG、PNG、WebP 圖片');
    const bytes = Utilities.base64Decode(m[2]);
    if (bytes.length > LIMITS.imgBytes) throw userError_('圖片太大（上限 3MB）');
    const ext = m[1].split('/')[1].replace('jpeg', 'jpg');
    const blob = Utilities.newBlob(bytes, m[1], 'img_' + Date.now() + '.' + ext);
    const file = imageFolder_().createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, id: file.getId() };
  },

  addItem: (s, b) => {
    const name = text_(b.name, LIMITS.name, '品名', true);
    const qty = int_(b.qty, 0, LIMITS.qtyMax, '數量');
    const id = nextId_();
    itemsSheet_().appendRow([ id, name, qty, text_(b.unit, LIMITS.unit, '單位') || '個',
      int_(b.minAlert, 0, LIMITS.qtyMax, '低庫存警示'), text_(b.note, LIMITS.note, '備註'),
      new Date().toISOString(), img_(b.imgUrl) ]);
    log_('新增品項', id, name, qty, qty, '期初數量', s.name);
    return { success: true, id };
  },

  /** 編輯品項資料；數量不可在此修改，只能透過領用/補貨/盤點 */
  updateItem: (s, b) => {
    const id = id_(b.id);
    const row = findRow_(id);
    if (row < 0) throw userError_('找不到品項');
    const sh = itemsSheet_();
    const qty = sh.getRange(row, 3).getValue();
    sh.getRange(row, 1, 1, ITEM_HEADERS.length).setValues([[
      id, text_(b.name, LIMITS.name, '品名', true), qty, text_(b.unit, LIMITS.unit, '單位') || '個',
      int_(b.minAlert, 0, LIMITS.qtyMax, '低庫存警示'), text_(b.note, LIMITS.note, '備註'),
      new Date().toISOString(), img_(b.imgUrl) ]]);
    return { success: true };
  },

  deleteItem: (s, b) => {
    const id = id_(b.id);
    const row = findRow_(id);
    if (row < 0) throw userError_('找不到品項');
    const sh = itemsSheet_();
    const qty = Number(sh.getRange(row, 3).getValue()) || 0;
    if (qty !== 0) throw userError_('庫存還有 ' + qty + '，請先歸零再刪除');
    const name = sh.getRange(row, 2).getValue();
    sh.deleteRow(row);
    log_('刪除品項', id, name, 0, 0, '', s.name);
    return { success: true };
  },

  adjustQty: (s, b) => {
    const id = id_(b.id);
    const delta = int_(b.delta, -LIMITS.qtyMax, LIMITS.qtyMax, '數量');
    if (delta === 0) throw userError_('數量不可為 0');
    const note = text_(b.note, LIMITS.note, '備註');
    const r = applyDelta_(id, delta);
    log_(delta < 0 ? '領用' : '補貨', id, r.name, delta, r.newQty, note, s.name);
    return { success: true, newQty: r.newQty };
  },

  /** 沖銷：對一筆領用/補貨產生反向記錄，並把庫存調回 */
  reverseLog: (s, b) => {
    const logId = String(b.logId || '');
    if (!/^L\d{5,}$/.test(logId)) throw userError_('記錄編號格式錯誤');
    const reason = text_(b.reason, LIMITS.note, '沖銷原因', true);
    const logs = readLogs_();
    const orig = logs.find(l => l['記錄ID'] === logId);
    if (!orig) throw userError_('找不到記錄');
    if (REVERSIBLE.indexOf(orig['類型']) < 0) throw userError_('這類記錄不能沖銷');
    if (logs.some(l => l['沖銷對象'] === logId)) throw userError_('這筆已經沖銷過了');
    const delta = -Number(orig['變動量']);
    const r = applyDelta_(String(orig['品項ID']), delta);
    const newId = log_('沖銷', String(orig['品項ID']), r.name, delta, r.newQty, '沖銷 ' + logId + '：' + reason, s.name, logId);
    return { success: true, newQty: r.newQty, logId: newId };
  }
};

/* ============ GET：不提供資料 ============ */
function doGet(){
  return json_({ error: '請使用贈品庫存網頁' });
}

/* ============ POST ============ */
function doPost(e){
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    let body;
    try { body = JSON.parse(e.postData.contents); } catch (x) { throw userError_('資料格式錯誤'); }
    const action = String(body.action || '');
    ensureLogSchema_();
    ensureConfigDefaults_();

    if (action === 'login') return json_(login_(body));
    if (!Object.prototype.hasOwnProperty.call(ACTIONS, action)) throw userError_('未知的操作');

    let s;
    if (loginRequired_()) s = auth_(body.token);
    else {
      if (action === 'changePasscode' || action === 'logout') throw userError_('目前為免登入模式');
      s = { name: text_(getConfig_('免登入操作者'), LIMITS.adminName, '操作者') || '未登入', mustChange: false, token: '', open: true };
    }
    if (s.mustChange && action !== 'changePasscode' && action !== 'logout') throw userError_('請先設定新密碼', 'MUST_CHANGE');
    return json_(ACTIONS[action](s, body));
  } catch (err) {
    if (err && err.userFacing) {
      const out = { error: err.message, code: err.code || '' };
      if (err.code === 'AUTH') { try { out.password = passwordRequired_(); } catch (x) {} }
      return json_(out);
    }
    console.error(err);
    return json_({ error: '操作失敗，請稍後再試' });
  } finally {
    lock.releaseLock();
  }
}
