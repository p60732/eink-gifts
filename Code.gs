/**
 * E Ink 贈品庫存管理系統 — Google Apps Script 後端
 * 部署為 Web App（執行身分：我；存取權：所有人）
 *
 * 首次使用：在編輯器選 initSheets → 執行（已有資料時會自動拒絕，避免覆蓋）
 * 設定集中在「設定」工作表（圖片資料夾 ID 等），交接或移轉時只需檢查該表。
 */

const SHEET_ITEMS  = '庫存';
const SHEET_LOGS   = '異動記錄';
const SHEET_CONFIG = '設定';
const ITEM_HEADERS = ['ID','品名','數量','單位','低庫存警示','備註','最後更新','圖片'];
const LOG_HEADERS  = ['時間','品項ID','品名','變動量','變動後數量','備註','操作者'];
const CONFIG_HEADERS = ['設定項目','值','說明'];

const LIMITS = {
  name: 50, unit: 10, note: 300, operator: 30, img: 300,
  qtyMax: 1000000,
  imgBytes: 3 * 1024 * 1024
};
const IMG_TYPES = ['image/jpeg','image/png','image/webp'];

/* ============ 初始化（僅限編輯器執行） ============ */
function initSheets(){
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
}

/* ============ 圖片搬家：把 GitHub images/ 的舊圖搬進雲端硬碟（僅限編輯器執行，可重複執行） ============ */
function migrateImagesToDrive(){
  const OLD_BASE = 'https://p60732.github.io/eink-gifts/';
  const sh = itemsSheet_();
  const last = sh.getLastRow();
  if (last < 2) return;
  const range = sh.getRange(2, 8, last - 1, 1);
  const vals = range.getValues();
  const folder = imageFolder_();
  let moved = 0, failed = 0;
  vals.forEach((r, i) => {
    const v = String(r[0] || '');
    if (!/^images\//.test(v)) return;
    try {
      const res = UrlFetchApp.fetch(OLD_BASE + v, { muteHttpExceptions: true });
      if (res.getResponseCode() !== 200) { failed++; return; }
      const blob = res.getBlob().setName(v.replace(/^images\//, ''));
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      vals[i][0] = file.getId();
      moved++;
    } catch (err) { console.error(err); failed++; }
  });
  range.setValues(vals);
  console.log('搬移完成：成功 ' + moved + '，失敗 ' + failed);
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
    sh.getRange(2,1,1,3).setValues([['圖片資料夾ID','','存放贈品圖片的雲端硬碟資料夾；留空會自動建立']]);
    sh.setColumnWidth(1, 140); sh.setColumnWidth(2, 320); sh.setColumnWidth(3, 360);
  }
  return sh;
}
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

function imageFolder_(){
  const id = getConfig_('圖片資料夾ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* 失效則重建 */ } }
  const parents = DriveApp.getFileById(ss_().getId()).getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const folder = parent.createFolder('eink贈品圖片');
  setConfig_('圖片資料夾ID', folder.getId());
  return folder;
}

function readItems_(){
  const values = itemsSheet_().getDataRange().getValues();
  const head = values.shift();
  return values.filter(r => r[0] !== '').map(r => {
    const o = {};
    head.forEach((h, i) => o[h] = r[i]);
    return o;
  });
}

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

/* ============ 輸入驗證 ============ */
function userError_(msg){ const e = new Error(msg); e.userFacing = true; return e; }

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

function log_(id, name, delta, after, note, operator){
  logsSheet_().appendRow([ new Date().toISOString(), id, name, delta, after, note, operator ]);
}

/* ============ GET ============ */
function doGet(e){
  try {
    const action = (e && e.parameter && e.parameter.action) || '';
    if (action === 'getAll') return json_({ items: readItems_() });
    if (action === 'getLogs') {
      const values = logsSheet_().getDataRange().getValues();
      const head = values.shift();
      let logs = values.filter(r => r[0] !== '').map(r => { const o = {}; head.forEach((h, i) => o[h] = r[i]); return o; });
      if (e.parameter.itemId) logs = logs.filter(l => String(l['品項ID']) === String(e.parameter.itemId));
      logs.reverse();
      return json_({ logs: logs.slice(0, 500) });
    }
    return json_({ error: '未知的操作' });
  } catch (err) {
    console.error(err);
    return json_({ error: '讀取失敗，請稍後再試' });
  }
}

/* ============ POST ============ */
function doPost(e){
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    let body;
    try { body = JSON.parse(e.postData.contents); } catch (x) { throw userError_('資料格式錯誤'); }
    const action = String(body.action || '');
    const operator = text_(body.operator, LIMITS.operator, '操作者');

    if (action === 'uploadImage') {
      const m = String(body.data || '').match(/^data:([\w/+.-]+);base64,([A-Za-z0-9+/=]+)$/);
      if (!m || IMG_TYPES.indexOf(m[1]) < 0) throw userError_('只接受 JPG、PNG、WebP 圖片');
      const bytes = Utilities.base64Decode(m[2]);
      if (bytes.length > LIMITS.imgBytes) throw userError_('圖片太大（上限 3MB）');
      const ext = m[1].split('/')[1].replace('jpeg', 'jpg');
      const blob = Utilities.newBlob(bytes, m[1], 'img_' + Date.now() + '.' + ext);
      const file = imageFolder_().createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return json_({ success: true, id: file.getId() });
    }

    if (action === 'addItem') {
      const name = text_(body.name, LIMITS.name, '品名', true);
      const qty = int_(body.qty, 0, LIMITS.qtyMax, '數量');
      const id = nextId_();
      itemsSheet_().appendRow([ id, name, qty, text_(body.unit, LIMITS.unit, '單位') || '個',
        int_(body.minAlert, 0, LIMITS.qtyMax, '低庫存警示'), text_(body.note, LIMITS.note, '備註'),
        new Date().toISOString(), img_(body.imgUrl) ]);
      log_(id, name, qty, qty, '新增品項', operator);
      return json_({ success: true, id });
    }

    if (action === 'updateItem') {
      const id = id_(body.id);
      const row = findRow_(id);
      if (row < 0) throw userError_('找不到品項');
      const sh = itemsSheet_();
      const oldQty = Number(sh.getRange(row, 3).getValue()) || 0;
      const name = text_(body.name, LIMITS.name, '品名', true);
      const qty = int_(body.qty, 0, LIMITS.qtyMax, '數量');
      sh.getRange(row, 1, 1, ITEM_HEADERS.length).setValues([[
        id, name, qty, text_(body.unit, LIMITS.unit, '單位') || '個',
        int_(body.minAlert, 0, LIMITS.qtyMax, '低庫存警示'), text_(body.note, LIMITS.note, '備註'),
        new Date().toISOString(), img_(body.imgUrl) ]]);
      if (qty !== oldQty) log_(id, name, qty - oldQty, qty, '編輯品項時修改數量', operator);
      return json_({ success: true });
    }

    if (action === 'deleteItem') {
      const id = id_(body.id);
      const row = findRow_(id);
      if (row < 0) throw userError_('找不到品項');
      const sh = itemsSheet_();
      const name = sh.getRange(row, 2).getValue();
      const qty = Number(sh.getRange(row, 3).getValue()) || 0;
      sh.deleteRow(row);
      log_(id, name, -qty, 0, '刪除品項', operator);
      return json_({ success: true });
    }

    if (action === 'adjustQty') {
      const id = id_(body.id);
      const row = findRow_(id);
      if (row < 0) throw userError_('找不到品項');
      const delta = int_(body.delta, -LIMITS.qtyMax, LIMITS.qtyMax, '數量');
      if (delta === 0) throw userError_('數量不可為 0');
      const sh = itemsSheet_();
      const cur = Number(sh.getRange(row, 3).getValue()) || 0;
      const newQty = cur + delta;
      if (newQty < 0) throw userError_('庫存不足');
      if (newQty > LIMITS.qtyMax) throw userError_('數量超過上限');
      const name = sh.getRange(row, 2).getValue();
      sh.getRange(row, 3).setValue(newQty);
      sh.getRange(row, 7).setValue(new Date().toISOString());
      log_(id, name, delta, newQty, text_(body.note, LIMITS.note, '備註'), operator);
      return json_({ success: true, newQty });
    }

    if (action === 'deleteLog') {
      const time = String(body.time || '');
      if (!/^\d{4}-\d{2}-\d{2}T[\d:.]+Z$/.test(time)) throw userError_('記錄時間格式錯誤');
      const sh = logsSheet_();
      const n = sh.getLastRow();
      if (n < 2) throw userError_('找不到記錄');
      const times = sh.getRange(2, 1, n - 1, 1).getValues();
      for (let i = times.length - 1; i >= 0; i--) {
        const t = times[i][0] instanceof Date ? times[i][0].toISOString() : String(times[i][0]);
        if (t === time) { sh.deleteRow(i + 2); return json_({ success: true }); }
      }
      throw userError_('找不到記錄');
    }

    throw userError_('未知的操作');
  } catch (err) {
    if (err && err.userFacing) return json_({ error: err.message });
    console.error(err);
    return json_({ error: '操作失敗，請稍後再試' });
  } finally {
    lock.releaseLock();
  }
}
