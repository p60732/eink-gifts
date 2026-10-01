/**
 * E Ink 贈品庫存管理系統 — Google Apps Script 後端
 * 部署為 Web App（執行身分：我；存取權：所有人）— 可在「設定」分頁切換是否需要登入
 *
 * 人員管理：網頁「設定 → 人員管理」可新增、改名、啟用/停用人員（需有「可管理人員」權限）。
 *   不能停用自己或拿掉自己的管理權限；至少保留一位可管理人員。人員不能刪除，只能停用。
 *
 * 管理者設定（僅限試算表擁有者，在編輯器執行）：
 *   1. 在「管理者」分頁新增一列，填「名稱」、「啟用」勾選
 *   2. 編輯器選 generateInitCodes → 執行，會在「初始碼」欄產生一次性代碼
 *   3. 把初始碼私下交給該管理者；對方登入後必須設定自己的密碼，初始碼隨即失效
 *   忘記密碼：清空該列「密碼雜湊」「鹽」→ 再執行 generateInitCodes
 *
 * 登入開關：「設定」分頁的「需要登入」取消勾選 = 免登入（任何知道網址的人都能操作），
 *   此時操作者記為「免登入操作者」欄的名稱。
 * 密碼開關：「登入需要密碼」取消勾選 = 只要輸入「管理者」名單內、已啟用的名稱即可登入。
 *   打開密碼開關後，先前「只用名稱」的登入會全部失效，需要重新用密碼登入。
 * 直接在試算表修改設定、人員或庫存時，最多 2 分鐘後網頁才會反映（快取）；從網頁操作則立即生效。
 *
 * 異動記錄為帳本：不可刪除，登打錯誤請用「沖銷」產生反向記錄。
 *
 * 盤點：同一時間只能有一張進行中的盤點單。差異 = 實點 − 輸入當下的帳面數量，
 *   所以盤點期間照常領用/補貨也不會算錯。完成盤點時先全部檢查、再一次寫入「盤點調整」。
 */

const SHEET_ITEMS  = '庫存';
const SHEET_LOGS   = '異動記錄';
const SHEET_CONFIG = '設定';
const SHEET_ADMINS = '管理者';
const SHEET_ST     = '盤點';
const SHEET_STD    = '盤點明細';
const ITEM_HEADERS  = ['ID','品名','數量','單位','低庫存警示','備註','最後更新','圖片'];
const LOG_HEADERS   = ['時間','品項ID','品名','變動量','變動後數量','備註','操作者','類型','記錄ID','沖銷對象'];
const CONFIG_HEADERS = ['設定項目','值','說明'];
const ADMIN_HEADERS = ['名稱','啟用','初始碼','密碼雜湊','鹽','最後登入','說明','可管理人員'];
const ST_HEADERS  = ['盤點ID','狀態','開始時間','開始者','完成時間','完成者','備註','品項數','已盤','有差異','調整合計'];
const STD_HEADERS = ['盤點ID','品項ID','品名','帳面數量','實點數量','差異','備註','輸入者','輸入時間','調整記錄ID'];

const LIMITS = {
  name: 50, unit: 10, note: 300, img: 300, adminName: 20,
  qtyMax: 1000000,
  imgBytes: 3 * 1024 * 1024,
  passMin: 6, passMax: 64
};
const IMG_TYPES = ['image/jpeg','image/png','image/webp'];
const SESSION_TTL = 6 * 3600;      // 登入有效 6 小時（CacheService 上限）
const FAIL_MAX = 5;                // 連續失敗次數上限
const FAIL_LOCK = 15 * 60;         // 鎖定 15 分鐘
const HASH_ROUNDS = 300;
const REVERSIBLE = ['領用','補貨'];
const CONFIG_DEFAULTS = [
  ['圖片資料夾ID', '', '存放贈品圖片的雲端硬碟資料夾；留空會自動建立'],
  ['需要登入', false, '勾選 = 需要管理者登入；取消勾選 = 任何知道網址的人都能操作'],
  ['免登入操作者', 'Kim', '不需登入時，異動記錄上的操作者名稱'],
  ['登入需要密碼', false, '勾選 = 名稱＋密碼；取消勾選 = 只要輸入「管理者」名單內、已啟用的名稱即可登入'],
  ['備份資料夾ID', '', '每季自動備份存放的雲端硬碟資料夾；留空會自動建立'],
  ['最後備份', '', '最近一次自動備份的時間（系統自動填寫）'],
  ['備份寄送信箱', '', '每半年（1/1、7/1）自動寄出備份 Excel 的收件信箱，多個用逗號分隔；空白 = 寄給試算表擁有者'],
  ['最後寄送備份', '', '最近一次寄出備份 Excel 的時間（系統自動填寫）']
];
const MAIL_MONTHS = [0, 6];      // 1 月、7 月寄出備份 Excel
const MAIL_MAX_TO = 10;
const BACKUP_KEEP = 8;           // 每季一份，保留約兩年
const BACKUP_MONTHS = [0, 3, 6, 9];   // 1、4、7、10 月
const BACKUP_MIN_GAP_DAYS = 45;   // 同一季只備份一次（舊的每週觸發條件也不會重複備份）
const BACKUP_TAG = '_備份_';

/* ============ 僅限擁有者在編輯器執行 ============ */
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

/** 備份整份試算表到「eink贈品備份」資料夾；超過 BACKUP_KEEP 份的舊備份移到垃圾桶。
 *  觸發條件呼叫時只在 1、4、7、10 月、且距上次備份超過 45 天才備份；在編輯器手動執行則立即備份 */
function backupSpreadsheet(e){
  assertOwner_();
  if (e && e.triggerUid) {
    const now = new Date();
    const last = Date.parse(getConfig_('最後備份'));
    if (BACKUP_MONTHS.indexOf(now.getMonth()) < 0) return;
    if (!isNaN(last) && now.getTime() - last < BACKUP_MIN_GAP_DAYS * 86400000) return;
  }
  const file = DriveApp.getFileById(ss_().getId());
  const folder = backupFolder_();
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Taipei', 'yyyy-MM-dd_HHmm');
  file.makeCopy(file.getName() + BACKUP_TAG + stamp, folder);
  const list = [];
  const it = folder.getFiles();
  while (it.hasNext()) { const f = it.next(); if (f.getName().indexOf(BACKUP_TAG) >= 0) list.push(f); }
  list.sort((a, b) => b.getDateCreated().getTime() - a.getDateCreated().getTime());
  list.slice(BACKUP_KEEP).forEach(f => f.setTrashed(true));   // 垃圾桶內 30 天內仍可救回
  setConfig_('最後備份', new Date().toISOString());
  console.log('已備份，目前保留 ' + Math.min(list.length, BACKUP_KEEP) + ' 份');
}

/** 只需執行一次：建立每季（1/1、4/1、7/1、10/1 早上 7 點）的自動備份，並立刻先備份一份 */
function setupQuarterlyBackup(){
  assertOwner_();
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'backupSpreadsheet')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('backupSpreadsheet').timeBased().onMonthDay(1).atHour(7).create();
  backupSpreadsheet();
  console.log('已設定每季自動備份（1/1、4/1、7/1、10/1）');
}

/** 每半年把備份 Excel（庫存、異動記錄、盤點）寄給「備份寄送信箱」。
 *  觸發條件每月 1 日早上 8 點呼叫，只在 1 月、7 月寄出；在編輯器手動執行則立即寄出 */
function mailBackupExcel(e){
  assertOwner_();
  const now = new Date();
  if (e && e.triggerUid && MAIL_MONTHS.indexOf(now.getMonth()) < 0) return;
  const to = backupRecipients_();
  const tz = Session.getScriptTimeZone() || 'Asia/Taipei';
  const day = Utilities.formatDate(now, tz, 'yyyy/MM/dd');
  const blob = backupXlsx_(Utilities.formatDate(now, tz, 'yyyyMMdd'));
  const nItems = Math.max(itemsSheet_().getLastRow() - 1, 0);
  const nLogs = Math.max(logsSheet_().getLastRow() - 1, 0);
  MailApp.sendEmail({
    to: to.join(','),
    subject: 'E Ink 贈品庫存 定期備份 ' + day,
    name: 'E Ink 贈品庫存',
    body: '附件是 ' + day + ' 的贈品庫存備份 Excel（庫存 ' + nItems + ' 項、異動記錄 ' + nLogs + ' 筆，另含盤點記錄）。\n'
      + '請存到自己的電腦或公司硬碟保存。\n\n'
      + '這封信每年 1/1、7/1 自動寄出。收件人可在贈品庫存網頁「設定 → 備份與安全」修改。',
    attachments: [blob]
  });
  setConfig_('最後寄送備份', now.toISOString());
  console.log('已寄出備份 Excel 給 ' + to.join(', '));
}

/** 只需執行一次：建立每半年寄出備份 Excel 的觸發條件，並立刻寄一封確認 */
function setupHalfYearMail(){
  assertOwner_();
  ensureConfigDefaults_();   // 「設定」分頁補上「備份寄送信箱」一列
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'mailBackupExcel')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('mailBackupExcel').timeBased().onMonthDay(1).atHour(8).create();
  mailBackupExcel();
  console.log('已設定每年 1/1、7/1 早上 8 點自動寄出備份 Excel');
}

/** 把「a@x.com, b@y.com」拆成信箱清單並檢查格式（逗號、分號、空白、全形都可分隔） */
function parseEmails_(raw){
  const list = String(raw || '').split(/[,，;；\s]+/).map(x => x.trim()).filter(Boolean);
  const bad = list.filter(x => x.length > 100 || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x));
  if (bad.length) throw userError_('信箱格式錯誤：' + bad.join(', '));
  if (list.length > MAIL_MAX_TO) throw userError_('備份寄送信箱最多 ' + MAIL_MAX_TO + ' 個');
  return list;
}
function backupRecipients_(){
  const list = parseEmails_(getConfig_('備份寄送信箱'));
  return list.length ? list : [ Session.getEffectiveUser().getEmail() ];
}

/** 只複製庫存、異動記錄、盤點到暫存試算表再轉成 Excel（不含人員名單與密碼欄），用完丟垃圾桶 */
function backupXlsx_(stamp){
  const src = ss_();
  const tmp = SpreadsheetApp.create('eink贈品備份_暫存_' + stamp);
  try {
    const blank = tmp.getSheets()[0];
    [SHEET_ITEMS, SHEET_LOGS, SHEET_ST, SHEET_STD].forEach(n => {
      const sh = src.getSheetByName(n);
      if (sh) sh.copyTo(tmp).setName(n);
    });
    tmp.deleteSheet(blank);
    SpreadsheetApp.flush();
    const res = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + tmp.getId() + '/export?format=xlsx', {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) throw new Error('轉成 Excel 失敗（' + res.getResponseCode() + '）');
    return res.getBlob().setName('贈品庫存_備份_' + stamp + '.xlsx');
  } finally {
    DriveApp.getFileById(tmp.getId()).setTrashed(true);
  }
}

function backupFolder_(){
  const id = getConfig_('備份資料夾ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* 失效則重建 */ } }
  const parents = DriveApp.getFileById(ss_().getId()).getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const folder = parent.createFolder('eink贈品備份');
  setConfig_('備份資料夾ID', folder.getId());
  return folder;
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
/* ============ 快取 ============ */
/** 資料世代：每次寫入換一代，所有資料快取的鍵都帶世代，舊世代的快取自然作廢 */
let GEN_ = null;
function dataGen_(){
  if (GEN_ === null) GEN_ = CacheService.getScriptCache().get('dgen') || '0';
  return GEN_;
}
function bumpDataGen_(){
  GEN_ = Utilities.getUuid().slice(0, 8);
  CacheService.getScriptCache().put('dgen', GEN_, 21600);
  CFG_ = null;
}
function cacheKey_(base){ return base + '_' + dataGen_(); }

/** 設定每個請求只讀一次（doPost 開頭會清掉） */
let CFG_ = null;
const CFG_KEY = 'cfg';
const DATA_TTL = 120;   // 直接在試算表手動修改時，最多 2 分鐘後網頁才看得到
function config_(){
  if (!CFG_) {
    const c = CacheService.getScriptCache();
    const key = cacheKey_(CFG_KEY);
    const raw = c.get(key);
    if (raw) CFG_ = JSON.parse(raw);
    else {
      CFG_ = {};
      configSheet_().getDataRange().getValues().slice(1).forEach(r => { if (r[0] !== '') CFG_[r[0]] = String(r[1] == null ? '' : r[1]).trim(); });
      c.put(key, JSON.stringify(CFG_), DATA_TTL);
    }
  }
  return CFG_;
}
function getConfig_(key){ return config_()[key] || ''; }
function setConfig_(key, value){
  const sh = configSheet_();
  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (vals[i][0] === key) { sh.getRange(i + 1, 2).setValue(value); bumpDataGen_(); return; }
  }
  sh.appendRow([key, value, '']);
  bumpDataGen_();
}

function adminsSheet_(){
  const ss = ss_();
  let sh = ss.getSheetByName(SHEET_ADMINS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_ADMINS);
    sh.getRange(1,1,1,ADMIN_HEADERS.length).setValues([ADMIN_HEADERS]);
    sh.getRange(2,1,2,ADMIN_HEADERS.length).setValues([
      ['佩芝', true, '', '', '', '', '系統擁有者', true],
      ['Kim',  true, '', '', '', '', '', true]
    ]);
    sh.getRange(2, 2, 50, 1).insertCheckboxes();
    sh.getRange(2, 8, 50, 1).insertCheckboxes();
    sh.setColumnWidth(4, 120); sh.setColumnWidth(5, 80); sh.setColumnWidth(7, 240);
  }
  return sh;
}

/** 舊版管理者分頁升級：補上「可管理人員」欄，既有啟用人員預設可管理 */
function ensureAdminSchema_(){
  const sh = adminsSheet_();
  const head = sh.getRange(1, 1, 1, ADMIN_HEADERS.length).getValues()[0];
  if (head[7] === '可管理人員') return;
  sh.getRange(1, 8).setValue('可管理人員');
  const n = sh.getLastRow() - 1;
  if (n > 0) {
    const vals = sh.getRange(2, 1, n, 2).getValues();
    sh.getRange(2, 8, n, 1).setValues(vals.map(r => [String(r[0]).trim() !== '' && r[1] === true]));
  }
  sh.getRange(2, 8, Math.max(n, 1) + 49, 1).insertCheckboxes();
  bumpDataGen_();
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
/** 讀取用的快取：鍵帶資料世代，寫入後立即作廢；單筆上限 100KB，太大就不快取 */
function cached_(base, build){
  const c = CacheService.getScriptCache();
  const key = cacheKey_(base);
  const raw = c.get(key);
  if (raw) return JSON.parse(raw);
  const val = build();
  const json = JSON.stringify(val);
  if (json.length < 90000) c.put(key, json, DATA_TTL);
  return val;
}
function readItemsCached_(){ return cached_('items', readItems_); }

function findRow_(id){
  const sh = itemsSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return -1;
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) { if (String(ids[i][0]) === String(id)) return i + 2; }
  return -1;
}

/** 下一個品項編號：同時參考庫存與異動記錄，刪除過的編號不會再被使用（避免舊記錄對到新品項） */
function nextId_(){
  let max = 0;
  const take = v => { const m = String(v).match(/^G(\d+)$/); if (m) max = Math.max(max, parseInt(m[1], 10)); };
  const ish = itemsSheet_(), lsh = logsSheet_();
  if (ish.getLastRow() > 1) ish.getRange(2, 1, ish.getLastRow() - 1, 1).getValues().forEach(r => take(r[0]));
  if (lsh.getLastRow() > 1) lsh.getRange(2, 2, lsh.getLastRow() - 1, 1).getValues().forEach(r => take(r[0]));
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

/** 人員名稱：限長度、不可用符號開頭（避免公式），只允許文字、數字、空白與 ._- */
function personName_(v){
  const s = String(v == null ? '' : v).trim();
  if (!s) throw userError_('請填寫名稱');
  if (s.length > LIMITS.adminName) throw userError_('名稱最多 ' + LIMITS.adminName + ' 字');
  if (!/^[\p{L}\p{N}][\p{L}\p{N} ._\-]*$/u.test(s)) throw userError_('名稱只能包含文字、數字、空白與 . _ -，且不可用符號開頭');
  return s;
}
function canManage_(s){
  if (!s || s.open) return false;
  const a = adminIndex_()[s.name];
  return !!(a && a[0] === true && a[1] === true);
}
function requireManage_(s){ if (!canManage_(s)) throw userError_('沒有管理人員的權限'); }
function readAdmins_(){
  const sh = adminsSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, ADMIN_HEADERS.length).getValues()
    .map((r, i) => ({ row: i + 2, r }))
    .filter(x => String(x.r[0]).trim() !== '');
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

/** 人員權限索引（名稱 → [啟用, 可管理, 登入世代]），快取；在網頁上改人員時立即作廢 */
function adminIndex_(){
  return cached_('adm', () => {
    const props = PropertiesService.getScriptProperties().getProperties();
    const idx = {};
    readAdmins_().forEach(x => {
      const n = String(x.r[0]).trim();
      idx[n] = [x.r[1] === true, x.r[7] === true, Number(props['sg:' + n] || 0), !!x.r[3]];   // 啟用、可管理、登入世代、已設密碼
    });
    return idx;
  });
}

/** 讓某名稱的所有既有登入失效（重設密碼、改密碼、改名、停用時使用） */
function bumpSessionGen_(name){
  const p = PropertiesService.getScriptProperties();
  const k = 'sg:' + name;
  p.setProperty(k, String(Number(p.getProperty(k) || 0) + 1));
  bumpDataGen_();
}
function sessionGen_(name){ return Number(PropertiesService.getScriptProperties().getProperty('sg:' + name) || 0); }

/** pw：這次登入是否有驗證密碼（密碼模式打開後，沒用密碼的舊登入一律失效） */
function newSession_(name, mustChange, pw){
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put('s_' + token, JSON.stringify({ n: name, mc: !!mustChange, pw: !!pw, g: sessionGen_(name) }), SESSION_TTL);
  return token;
}

/** 驗證 session，回傳 { name, mustChange, token } */
function auth_(token){
  const t = String(token || '');
  if (!/^[a-f0-9]{64}$/.test(t)) throw userError_('請先登入', 'AUTH');
  const raw = CacheService.getScriptCache().get('s_' + t);
  if (!raw) throw userError_('登入已逾時，請重新登入', 'AUTH');
  const s = JSON.parse(raw);
  const a = adminIndex_()[s.n];
  const drop = msg => { CacheService.getScriptCache().remove('s_' + t); throw userError_(msg, 'AUTH'); };
  if (!a || a[0] !== true) drop('帳號已停用');
  if ((a[2] || 0) !== (s.g || 0)) drop('登入已失效，請重新登入');
  const needPass = passwordRequired_() && a[1] === true;   // 只有可管理人員需要密碼
  if (needPass && !s.pw) drop('現在需要密碼，請重新登入');
  return { name: s.n, mustChange: s.mc && needPass, token: t };
}

function login_(body){
  const name = String(body.name || '').trim();
  const pass = String(body.passcode || '');
  const pwMode = passwordRequired_();
  if (!name || name.length > LIMITS.adminName || pass.length > LIMITS.passMax) throw userError_(pwMode ? '名稱或密碼錯誤' : '名稱不在人員名單內');
  const cache = CacheService.getScriptCache();
  const failKey = 'f_' + Utilities.base64EncodeWebSafe(name);
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= FAIL_MAX) throw userError_('錯誤次數過多，請 15 分鐘後再試');

  const a = findAdmin_(name);
  // 密碼模式只套用在可管理人員；一般人員輸入工號（名稱）即可登入
  const needPass = pwMode && !!a && a.r[7] === true;
  if (needPass && !pass && a.r[1] === true) throw userError_('管理人員請輸入密碼', 'NEED_PASS');
  let ok = false, mustChange = false;
  if (a && a.r[1] === true && !needPass) ok = true;
  else if (a && a.r[1] === true) {
    const init = String(a.r[2] || ''), hashed = String(a.r[3] || ''), salt = String(a.r[4] || '');
    if (hashed && salt) ok = hash_(pass, salt) === hashed;
    else if (init) { ok = pass === init; mustChange = ok; }
  }
  if (!ok) {
    cache.put(failKey, String(fails + 1), FAIL_LOCK);
    throw userError_(pwMode ? '名稱或密碼錯誤' : '名稱不在人員名單內');
  }
  cache.remove(failKey);
  adminsSheet_().getRange(a.row, 6).setValue(new Date().toISOString());
  return { success: true, token: newSession_(name, mustChange, needPass), name, mustChange };
}

function changePasscode_(s, body){
  if (!canManage_(s)) throw userError_('一般人員登入不需要密碼');
  const a = findAdmin_(s.name);
  const next = String(body.newPasscode || '');
  if (next.length < LIMITS.passMin || next.length > LIMITS.passMax) throw userError_('新密碼需 ' + LIMITS.passMin + '～' + LIMITS.passMax + ' 字');
  if (!s.mustChange && a.r[3]) {   // 已經有密碼才要驗證目前密碼；第一次設定（密碼模式打開前先設好）不用
    const cur = String(body.passcode || '');
    if (hash_(cur, String(a.r[4])) !== String(a.r[3])) throw userError_('目前密碼錯誤');
  }
  if (next === String(a.r[2] || '')) throw userError_('新密碼不可與初始碼相同');
  const salt = Utilities.getUuid();
  const sh = adminsSheet_();
  sh.getRange(a.row, 3, 1, 3).setValues([['', hash_(next, salt), salt]]);
  bumpSessionGen_(s.name);   // 其他裝置上的舊登入一併失效
  return { success: true, token: newSession_(s.name, false, true), name: s.name, mustChange: false };
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

/** 用記錄ID找那一列：Lxxxxx 的數字＋1 就是列號（新增時依列號編號）；對不上才掃「記錄ID」一欄 */
function findLogRow_(logId){
  const sh = logsSheet_();
  const last = sh.getLastRow();
  const W = LOG_HEADERS.length;
  let row = parseInt(logId.slice(1), 10) + 1;
  let r = row >= 2 && row <= last ? sh.getRange(row, 1, 1, W).getValues()[0] : null;
  if (!r || String(r[8]) !== logId) {
    if (last < 2) return null;
    const i = sh.getRange(2, 9, last - 1, 1).getValues().findIndex(x => String(x[0]) === logId);
    if (i < 0) return null;
    row = i + 2; r = sh.getRange(row, 1, 1, W).getValues()[0];
  }
  return { row, log: rowsToObjects_([LOG_HEADERS, r])[0] };
}

const LOG_TYPES = ['領用','補貨','新增品項','刪除品項','沖銷','盤點調整'];

/** 分頁讀取（新到舊）：limit 預設 20（上限 5000）、offset、itemId、type、fromTs/toTs（ISO 時間）
 *  只讀需要的欄與列：
 *  - 沒有篩選：只讀要顯示的那幾列
 *  - 有篩選：先只讀篩選用的欄（時間／品項／類型各 1 欄）找出符合的列，再只讀這一頁的列
 *  - 「已沖銷」：沖銷一定在原記錄之後，所以只讀這一頁最舊那列到最新的「記錄ID／沖銷對象」兩欄 */
function readLogsPage_(b){
  const limit = b.limit == null ? 20 : int_(b.limit, 1, 5000, '筆數');
  const offset = b.offset == null ? 0 : int_(b.offset, 0, 10000000, '起始位置');
  const itemId = b.itemId ? id_(b.itemId) : '';
  const type = b.type ? String(b.type) : '';
  if (type && LOG_TYPES.indexOf(type) < 0) throw userError_('記錄類型錯誤');
  const fromTs = b.fromTs ? tsMs_(b.fromTs) : null, toTs = b.toTs ? tsMs_(b.toTs) : null;
  const sh = logsSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return { logs: [], total: 0 };
  const W = LOG_HEADERS.length;
  const col = c => sh.getRange(2, c, n, 1).getValues();
  let rows, total, firstRow;
  if (!itemId && !type && fromTs == null && toTs == null) {
    total = n;
    const end = n - offset;                       // 由新到舊：最後一列最新
    if (end < 1) return { logs: [], total };
    const cnt = Math.min(limit, end);
    firstRow = 2 + end - cnt;
    rows = sh.getRange(firstRow, 1, cnt, W).getValues().reverse();
  } else {
    let ok = null;                                // 0 起算的列索引是否符合
    const keep = test => { ok = (ok || new Array(n).fill(true)).map((v, i) => v && test(i)); };
    if (fromTs != null || toTs != null) {
      const ts = col(1);
      keep(i => { const v = ts[i][0]; const t = v instanceof Date ? v.getTime() : Date.parse(v);
        return v !== '' && (fromTs == null || t >= fromTs) && (toTs == null || t <= toTs); });
    }
    if (itemId) { const ids = col(2); keep(i => String(ids[i][0]) === itemId); }
    if (type) { const ty = col(8); keep(i => String(ty[i][0]) === type); }
    const idx = [];
    for (let i = n - 1; i >= 0; i--) if (ok[i]) idx.push(i);
    total = idx.length;
    const page = idx.slice(offset, offset + limit);
    if (!page.length) return { logs: [], total };
    const lo = page[page.length - 1], hi = page[0];
    const span = sh.getRange(lo + 2, 1, hi - lo + 1, W).getValues();
    rows = page.map(i => span[i - lo]);
    firstRow = lo + 2;
  }
  const revBy = {};
  sh.getRange(firstRow, 9, n + 2 - firstRow, 2).getValues().forEach(p => { if (p[1]) revBy[p[1]] = p[0]; });
  const logs = rowsToObjects_([LOG_HEADERS].concat(rows.filter(r => r[0] !== '')));
  logs.forEach(l => { l['已沖銷'] = revBy[l['記錄ID']] || ''; });
  return { logs, total };
}
function tsMs_(v){
  const t = Date.parse(String(v));
  if (isNaN(t)) throw userError_('日期格式錯誤');
  return t;
}

/* ============ 盤點 ============ */
function sheetWithHeaders_(name, headers){
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.getRange(1, 1, 1, headers.length).setValues([headers]); }
  return sh;
}
function stSheet_(){  return sheetWithHeaders_(SHEET_ST, ST_HEADERS); }
function stdSheet_(){ return sheetWithHeaders_(SHEET_STD, STD_HEADERS); }
function readSt_(){ return rowsToObjects_(stSheet_().getDataRange().getValues()); }
function stId_(v){
  const s = String(v || '');
  if (!/^S\d{4,}$/.test(s)) throw userError_('盤點單編號格式錯誤');
  return s;
}
function findStRow_(id){
  const vals = stSheet_().getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) if (String(vals[i][0]) === id) return { row: i + 1, r: vals[i] };
  return null;
}
function openSt_(){
  return readSt_().find(x => x['狀態'] === '進行中') || null;
}
/** 某張盤點單的明細：先只讀「盤點ID」一欄找出範圍，再只讀那幾列 */
function stDetails_(id){
  const sh = stdSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const ids = sh.getRange(2, 1, n, 1).getValues();
  let lo = -1, hi = -1;
  for (let i = 0; i < n; i++) if (String(ids[i][0]) === id) { if (lo < 0) lo = i; hi = i; }
  if (lo < 0) return [];
  const rows = sh.getRange(2 + lo, 1, hi - lo + 1, STD_HEADERS.length).getValues().filter(r => String(r[0]) === id);
  return rowsToObjects_([STD_HEADERS].concat(rows));
}
/** 用已讀進來的明細（含表頭）更新盤點單的「已盤／有差異」，不再重讀一次 */
function updateStProgress_(stRow, id, vals){
  const det = vals.slice(1).filter(r => String(r[0]) === id);
  stSheet_().getRange(stRow, 9, 1, 2).setValues([[ det.length, det.filter(r => Number(r[5]) !== 0).length ]]);
}
function requireOpenSt_(id){
  const f = findStRow_(stId_(id));
  if (!f) throw userError_('找不到盤點單');
  if (f.r[1] !== '進行中') throw userError_('這張盤點單已經' + f.r[1]);
  return f;
}

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
  getAll: (s) => {
    const manage = canManage_(s);
    const out = { items: readItemsCached_(), user: s.name, openMode: !!s.open, passwordMode: passwordRequired_(), canManage: manage };
    if (manage) {
      out.lastBackup = getConfig_('最後備份'); out.lastMail = getConfig_('最後寄送備份');
      out.backupEmails = getConfig_('備份寄送信箱');
      out.hasPassword = !!(adminIndex_()[s.name] || [])[3];
    }
    return out;
  },

  /** 人員清單（不含密碼、初始碼等敏感欄位） */
  listAdmins: (s) => {
    requireManage_(s);
    return { people: readAdmins_().map(x => ({
      name: String(x.r[0]).trim(), enabled: x.r[1] === true, canManage: x.r[7] === true,
      lastLogin: x.r[5] instanceof Date ? x.r[5].toISOString() : String(x.r[5] || ''),
      note: String(x.r[6] || ''), hasPassword: !!x.r[3], pendingInit: !!x.r[2]
    })), me: s.name, passwordMode: passwordRequired_() };
  },

  /** 新增（origName 空白）或修改人員 */
  saveAdmin: (s, b) => {
    requireManage_(s);
    const orig = String(b.origName || '').trim();
    const name = personName_(b.name);
    const enabled = b.enabled === true, manage = b.canManage === true;
    const note = text_(b.note, 50, '說明');
    const all = readAdmins_();
    const dup = all.find(x => String(x.r[0]).trim() === name && String(x.r[0]).trim() !== orig);
    if (dup) throw userError_('已經有叫「' + name + '」的人員');
    const sh = adminsSheet_();
    let code = '';
    if (!orig) {
      const init = passwordRequired_() && manage ? randomCode_(8) : '';   // 一般人員不需要密碼
      sh.appendRow([ name, enabled, init, '', '', '', note, manage ]);
      const last = sh.getLastRow();
      sh.getRange(last, 2).insertCheckboxes(); sh.getRange(last, 8).insertCheckboxes();
      sh.getRange(last, 2).setValue(enabled); sh.getRange(last, 8).setValue(manage);
      code = init;
    } else {
      const t = all.find(x => String(x.r[0]).trim() === orig);
      if (!t) throw userError_('找不到人員');
      if (orig === s.name && !enabled) throw userError_('不能停用自己');
      if (orig === s.name && !manage) throw userError_('不能拿掉自己的管理權限');
      const managers = all.filter(x => x.r[1] === true && x.r[7] === true && String(x.r[0]).trim() !== orig).length;
      if (!(enabled && manage) && managers === 0) throw userError_('至少要保留一位可管理人員');
      sh.getRange(t.row, 1, 1, 2).setValues([[ name, enabled ]]);
      sh.getRange(t.row, 7, 1, 2).setValues([[ note, manage ]]);
      const wasManage = t.r[7] === true;
      // 密碼模式下升為可管理人員、但還沒有密碼：發一組一次性初始碼
      if (manage && !wasManage && passwordRequired_() && !t.r[3] && !t.r[2]) {
        code = randomCode_(8);
        sh.getRange(t.row, 3).setValue(code);
      }
      // 改名、停用或權限改變：舊名稱的登入全部失效（重新登入後才會套用新權限）
      if (name !== orig || !enabled || manage !== wasManage) bumpSessionGen_(orig);
    }
    // 新名稱（新增或改名）若曾被別人用過，先讓那些舊登入失效，避免被新的人繼承
    if (!orig || name !== orig) bumpSessionGen_(name);
    bumpDataGen_();
    const out = { success: true, initCode: code };
    if (orig && orig === s.name && name !== orig) {   // 改自己的名字：換發新的登入
      out.token = newSession_(name, false, passwordRequired_()); out.name = name;
    }
    return out;
  },

  /** 網頁「設定 → 備份與安全」：備份寄送信箱、管理人員登入是否需要密碼 */
  saveSettings: (s, b) => {
    const out = { success: true };
    if (b.backupEmails !== undefined) setConfig_('備份寄送信箱', parseEmails_(b.backupEmails).join(', '));
    if (b.managerPassword !== undefined) {
      const on = b.managerPassword === true, was = passwordRequired_();
      if (on && !was) {
        const me = findAdmin_(s.name);
        if (!me.r[3]) throw userError_('請先在「設定」設定你自己的密碼，再打開這個選項');
        if (hash_(String(b.passcode || ''), String(me.r[4])) !== String(me.r[3])) throw userError_('目前密碼錯誤');
        // 其他還沒有密碼、也沒有初始碼的管理人員：發一次性初始碼，避免被鎖在外面
        const sh = adminsSheet_(), codes = [];
        readAdmins_().forEach(x => {
          if (x.r[1] === true && x.r[7] === true && !x.r[3] && !x.r[2]) {
            const c = randomCode_(8); sh.getRange(x.row, 3).setValue(c);
            codes.push({ name: String(x.r[0]).trim(), code: c });
          }
        });
        setConfig_('登入需要密碼', true);
        out.initCodes = codes;
        out.token = newSession_(s.name, false, true); out.name = s.name;   // 自己剛驗證過密碼，不必重新登入
      } else if (!on && was) {
        setConfig_('登入需要密碼', false);
      }
    }
    return out;
  },

  /** 重設某人的密碼：清除舊密碼並產生新的一次性初始碼（只限可管理人員；密碼模式關閉時先備好，打開後用初始碼登入） */
  resetAdminPasscode: (s, b) => {
    requireManage_(s);
    const name = personName_(b.name);
    const t = readAdmins_().find(x => String(x.r[0]).trim() === name);
    if (!t) throw userError_('找不到人員');
    if (t.r[7] !== true) throw userError_('一般人員登入不需要密碼');
    const code = randomCode_(8);
    adminsSheet_().getRange(t.row, 3, 1, 3).setValues([[ code, '', '' ]]);
    bumpSessionGen_(name);   // 對方目前的登入立即失效
    return { success: true, initCode: code };
  },

  /** 異動記錄（新到舊）。沒有篩選時只讀需要的那幾列 */
  getLogs: (s, b) => {
    // 最常用的「最新 20 筆」有快取；網頁上任何寫入後立即清除
    const isFirst = !b.itemId && !b.type && !b.fromTs && !b.toTs && !b.offset && (b.limit == null || Number(b.limit) === 20);
    if (!isFirst) return readLogsPage_(b);
    return cached_('logs0', () => readLogsPage_(b));
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
    let note = '';
    if (qty !== 0) {
      // 建立錯誤（例如重複新增）：新增後完全沒有其他異動、數量仍等於期初數量，才允許連同數量一起刪除
      const lsh = logsSheet_();
      const ln = lsh.getLastRow() - 1;
      // 只讀「品項ID」到「類型」這幾欄（B～H），不讀備註等其他欄
      const mine = ln < 1 ? [] : lsh.getRange(2, 2, ln, 7).getValues().filter(r => String(r[0]) === id);
      const onlyCreated = mine.length === 1 && mine[0][6] === '新增品項' && Number(mine[0][2]) === qty;
      if (!onlyCreated) throw userError_('庫存還有 ' + qty + '，這個品項已有其他異動，請先用拿出或盤點歸零再刪除');
      note = '建立錯誤，連同期初數量 ' + qty + ' 一起刪除';
    }
    const open = openSt_();
    if (open && stDetails_(String(open['盤點ID'])).some(d => String(d['品項ID']) === id)) {
      throw userError_('這個品項在進行中的盤點單裡已有實點數量，請先清除該項或完成盤點再刪除');
    }
    const name = sh.getRange(row, 2).getValue();
    sh.deleteRow(row);
    log_('刪除品項', id, name, -qty, 0, note, s.name);
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

  /** 盤點：目前進行中的盤點單（含明細）與最近的歷史 */
  getStocktakes: () => cached_('st', () => {   // 有快取；網頁上任何寫入後立即作廢
    const all = readSt_();
    const open = all.find(x => x['狀態'] === '進行中') || null;   // 同 openSt_，但重用已讀的資料
    return { open, details: open ? stDetails_(open['盤點ID']) : [], history: all.filter(x => x['狀態'] !== '進行中').reverse().slice(0, 50) };
  }),

  getStocktakeDetail: (s, b) => {
    const id = stId_(b.stId);
    const f = findStRow_(id);
    if (!f) throw userError_('找不到盤點單');
    return { stocktake: rowsToObjects_([ST_HEADERS, f.r])[0], details: stDetails_(id) };
  },

  startStocktake: (s, b) => {
    if (openSt_()) throw userError_('已有進行中的盤點單，請先完成或取消');
    const sh = stSheet_();
    const id = 'S' + String(sh.getLastRow()).padStart(4, '0');
    const itemCount = Math.max(itemsSheet_().getLastRow() - 1, 0);   // 只要筆數，不讀整份庫存
    const row = [ id, '進行中', new Date().toISOString(), s.name, '', '', text_(b.note, LIMITS.note, '備註'), itemCount, 0, 0, 0 ];
    sh.appendRow(row);
    // 直接回傳新的盤點單，網頁不必再讀一次
    return { success: true, stId: id, open: rowsToObjects_([ST_HEADERS, row])[0] };
  },

  /** 輸入/修改某品項的實點數量；帳面以輸入當下的庫存為準 */
  saveCount: (s, b) => {
    const f = requireOpenSt_(b.stId);
    const id = f.r[0];
    const itemId = id_(b.itemId);
    const counted = int_(b.counted, 0, LIMITS.qtyMax, '實點數量');
    const row = findRow_(itemId);
    if (row < 0) throw userError_('找不到品項');
    const nb = itemsSheet_().getRange(row, 2, 1, 2).getValues()[0];   // 品名、數量一次讀
    const name = nb[0], book = Number(nb[1]) || 0;
    const rec = [ id, itemId, name, book, counted, counted - book, text_(b.note, LIMITS.note, '備註'), s.name, new Date().toISOString(), '' ];
    const sh = stdSheet_();
    const vals = sh.getDataRange().getValues();
    let at = -1;
    for (let i = 1; i < vals.length; i++) if (String(vals[i][0]) === id && String(vals[i][1]) === itemId) { at = i + 1; break; }
    if (at > 0 && vals[at - 1][9]) throw userError_('這項已經寫入盤點調整，不能再修改');
    if (at > 0) { sh.getRange(at, 1, 1, STD_HEADERS.length).setValues([rec]); vals[at - 1] = rec; }
    else { sh.appendRow(rec); vals.push(rec); }
    updateStProgress_(f.row, id, vals);
    return { success: true, book, counted, diff: counted - book };
  },

  /** 清除某品項的實點數量（輸錯想改回「未盤」時用） */
  clearCount: (s, b) => {
    const f = requireOpenSt_(b.stId);
    const id = f.r[0];
    const itemId = id_(b.itemId);
    const sh = stdSheet_();
    const vals = sh.getDataRange().getValues();
    const hit = [];
    for (let i = 1; i < vals.length; i++) if (String(vals[i][0]) === id && String(vals[i][1]) === itemId) hit.push(i);
    if (hit.some(i => vals[i][9])) throw userError_('這項已經寫入盤點調整，不能清除');
    hit.reverse().forEach(i => { sh.deleteRow(i + 1); vals.splice(i, 1); });   // 由下往上刪，列號才不會跑掉
    updateStProgress_(f.row, id, vals);
    return { success: true };
  },

  /** 完成盤點：先檢查全部，全部可行才一次寫入盤點調整 */
  finalizeStocktake: (s, b) => {
    const f = requireOpenSt_(b.stId);
    const id = f.r[0];
    const det = stDetails_(id);
    if (!det.length) throw userError_('還沒有輸入任何實點數量');
    const items = readItems_();
    // 已有調整記錄ID的列代表先前已寫入過（例如中途出錯），不重複調整
    const plan = det.filter(d => Number(d['差異']) !== 0 && !d['調整記錄ID']).map(d => {
      const it = items.find(i => String(i.ID) === String(d['品項ID']));
      if (!it) throw userError_('品項 ' + d['品名'] + ' 已不存在，請先清除該項的實點數量');
      const after = Number(it['數量']) + Number(d['差異']);
      if (after < 0) throw userError_('「' + d['品名'] + '」盤點後會變成負數，請重新確認實點數量');
      if (after > LIMITS.qtyMax) throw userError_('「' + d['品名'] + '」盤點後數量超過上限');
      return d;
    });
    const sh = stdSheet_();
    const vals = sh.getDataRange().getValues();
    let total = det.filter(d => d['調整記錄ID']).reduce((n, d) => n + Number(d['差異']), 0);
    plan.forEach(d => {
      const diff = Number(d['差異']);
      const r = applyDelta_(String(d['品項ID']), diff);
      const logId = log_('盤點調整', String(d['品項ID']), r.name, diff, r.newQty,
        '盤點 ' + id + '：帳面 ' + d['帳面數量'] + ' → 實點 ' + d['實點數量'] + (d['備註'] ? '（' + d['備註'] + '）' : ''), s.name);
      for (let i = 1; i < vals.length; i++) if (String(vals[i][0]) === id && String(vals[i][1]) === String(d['品項ID'])) { sh.getRange(i + 1, 10).setValue(logId); break; }
      total += diff;
    });
    stSheet_().getRange(f.row, 2, 1, 10).setValues([[ '已完成', f.r[2], f.r[3], new Date().toISOString(), s.name, f.r[6], items.length, det.length, det.filter(d => Number(d['差異']) !== 0).length, total ]]);
    return { success: true, adjusted: plan.length, counted: det.length, total };
  },

  cancelStocktake: (s, b) => {
    const f = requireOpenSt_(b.stId);
    stSheet_().getRange(f.row, 2, 1, 5).setValues([[ '已取消', f.r[2], f.r[3], new Date().toISOString(), s.name ]]);
    return { success: true };
  },

  /** 沖銷：對一筆領用/補貨產生反向記錄，並把庫存調回 */
  reverseLog: (s, b) => {
    const logId = String(b.logId || '');
    if (!/^L\d{5,}$/.test(logId)) throw userError_('記錄編號格式錯誤');
    const reason = text_(b.reason, LIMITS.note, '沖銷原因', true);
    const f = findLogRow_(logId);
    if (!f) throw userError_('找不到記錄');
    const orig = f.log;
    if (REVERSIBLE.indexOf(orig['類型']) < 0) throw userError_('這類記錄不能沖銷');
    if (!canManage_(s) && String(orig['操作者']) !== s.name) throw userError_('只能沖銷自己登打的記錄，其他人的請找管理人員');
    // 沖銷一定在原記錄之後：只讀原記錄那列以後的「沖銷對象」一欄
    const lsh = logsSheet_();
    if (lsh.getRange(f.row, 10, lsh.getLastRow() - f.row + 1, 1).getValues().some(x => String(x[0]) === logId)) throw userError_('這筆已經沖銷過了');
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
/** 只讀取、不寫入的動作：不必排隊等鎖 */
const READ_ONLY = ['getAll','getLogs','getStocktakes','getStocktakeDetail','listAdmins'];
/** 只有「可管理人員」能做的動作：編輯品項清單、盤點、人員管理。一般人員只能拿出、補入、沖銷自己的記錄 */
const MANAGE_ONLY = ['uploadImage','addItem','updateItem','deleteItem',
  'getStocktakes','getStocktakeDetail','startStocktake','saveCount','clearCount','finalizeStocktake','cancelStocktake',
  'listAdmins','saveAdmin','resetAdminPasscode','saveSettings'];
const SCHEMA_KEY = 'schema_ok_v4';   // 改版號 = 讓新的設定列（備份寄送信箱）補上

/** 格式檢查（補欄位、補設定列）只在快取過期時做一次 */
function ensureSchemaOnce_(lock){
  const c = CacheService.getScriptCache();
  if (c.get(SCHEMA_KEY)) return;
  if (!lock.hasLock()) lock.waitLock(20000);
  ensureLogSchema_();
  ensureConfigDefaults_();
  ensureAdminSchema_();
  stSheet_(); stdSheet_();   // 盤點分頁先建好，讀取動作就不必建立工作表
  SpreadsheetApp.flush();
  bumpDataGen_();            // 格式有更新：讓之前的快取全部作廢
  c.put(SCHEMA_KEY, '1', 21600);
}

function doPost(e){
  const lock = LockService.getScriptLock();
  CFG_ = null; GEN_ = null;
  let writeAction = false;
  try {
    let body;
    try { body = JSON.parse(e.postData.contents); } catch (x) { throw userError_('資料格式錯誤'); }
    const action = String(body.action || '');
    if (action !== 'login' && !Object.prototype.hasOwnProperty.call(ACTIONS, action)) throw userError_('未知的操作');
    const readOnly = READ_ONLY.indexOf(action) >= 0;
    writeAction = !readOnly && action !== 'login';
    if (!readOnly) lock.waitLock(20000);
    ensureSchemaOnce_(lock);

    if (action === 'login') return json_(login_(body));

    let s;
    if (loginRequired_()) s = auth_(body.token);
    else {
      if (action === 'changePasscode' || action === 'logout') throw userError_('目前為免登入模式');
      s = { name: text_(getConfig_('免登入操作者'), LIMITS.adminName, '操作者') || '未登入', mustChange: false, token: '', open: true };
    }
    if (s.mustChange && action !== 'changePasscode' && action !== 'logout') throw userError_('請先設定新密碼', 'MUST_CHANGE');
    if (MANAGE_ONLY.indexOf(action) >= 0) requireManage_(s);
    // 防重複送出：同一次送出（同一個 rid）若已成功，直接回上次結果，不再執行
    // 檢查在取得鎖之後，所以兩個同時到達的重複請求也只會執行一次
    const rid = writeAction && /^[A-Za-z0-9_-]{12,64}$/.test(String(body.rid || '')) ? 'rid_' + s.name + '_' + body.rid : '';
    if (rid) {
      const prev = CacheService.getScriptCache().get(rid);
      if (prev) { const o = JSON.parse(prev); o.duplicate = true; return json_(o); }
    }
    const result = ACTIONS[action](s, body);
    if (rid) { try { CacheService.getScriptCache().put(rid, JSON.stringify(result), 1800); } catch (x) {} }
    return json_(result);
  } catch (err) {
    if (err && err.userFacing) {
      const out = { error: err.message, code: err.code || '' };
      if (err.code === 'AUTH') { try { out.password = passwordRequired_(); } catch (x) {} }
      return json_(out);
    }
    console.error(err);
    return json_({ error: '操作失敗，請稍後再試' });
  } finally {
    if (lock.hasLock()) {
      try { SpreadsheetApp.flush(); } catch (x) {}   // 寫入確實落地後才放鎖
      if (writeAction) { try { bumpDataGen_(); } catch (x) {} }
      lock.releaseLock();
    }
  }
}
