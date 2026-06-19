/**
 * E Ink 贈品庫存管理系統 — Google Apps Script 後端
 * 部署為 Web App（執行身分：我；存取權：所有人）
 *
 * 首次使用：在編輯器選 initSheets 函式 → 執行（會建立工作表 + 預設資料）
 */

const SHEET_ITEMS = '庫存';
const SHEET_LOGS  = '異動記錄';
const ITEM_HEADERS = ['ID','品名','數量','單位','低庫存警示','備註','最後更新','圖片'];
const LOG_HEADERS  = ['時間','品項ID','品名','變動量','變動後數量','備註','操作者'];

/* ============ 初始化 ============ */
function initSheets(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let items = ss.getSheetByName(SHEET_ITEMS);
  if(!items){ items = ss.insertSheet(SHEET_ITEMS); }
  items.clear();
  items.getRange(1,1,1,ITEM_HEADERS.length).setValues([ITEM_HEADERS]);
  const seed = [
    ['G001','便當袋',     400,'個',50,'Love E Ink 印花款',                       new Date().toISOString(),'images/G001.jpg'],
    ['G002','鋼杯',       254,'個',30,'',                                          new Date().toISOString(),'images/G002.jpg'],
    ['G003','拉鏈袋',     210,'個',30,'網格款 We Make Surfaces Smart and Green',   new Date().toISOString(),'images/G003.jpg'],
    ['G004','咖啡渣杯',    12,'個',10,'咖啡渣＋竹纖維製，符合美標FDA/歐標LFGB',     new Date().toISOString(),'images/G004.jpg'],
    ['G005','飲料袋',     160,'個',20,'紅色提把帆布',                              new Date().toISOString(),'images/G005.jpg'],
    ['G006','行李箱綁帶', 616,'個',50,'紅白魔鬼氈捲裝',                            new Date().toISOString(),'images/G006.jpg'],
  ];
  items.getRange(2,1,seed.length,ITEM_HEADERS.length).setValues(seed);

  let logs = ss.getSheetByName(SHEET_LOGS);
  if(!logs){ logs = ss.insertSheet(SHEET_LOGS); }
  logs.clear();
  logs.getRange(1,1,1,LOG_HEADERS.length).setValues([LOG_HEADERS]);
}

/* ============ 共用工具 ============ */
function itemsSheet(){ return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ITEMS); }
function logsSheet(){  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOGS); }

function readItems(){
  const sh = itemsSheet();
  const values = sh.getDataRange().getValues();
  const head = values.shift();
  return values.filter(r=>r[0]!=='').map(r=>{
    const o = {};
    head.forEach((h,i)=> o[h] = r[i]);
    return o;
  });
}

function findRow(id){
  const sh = itemsSheet();
  const ids = sh.getRange(2,1,Math.max(sh.getLastRow()-1,1),1).getValues();
  for(let i=0;i<ids.length;i++){ if(String(ids[i][0])===String(id)) return i+2; }
  return -1;
}

function nextId(){
  const items = readItems();
  let max = 0;
  items.forEach(it=>{
    const m = String(it.ID).match(/^G(\d+)$/);
    if(m) max = Math.max(max, parseInt(m[1],10));
  });
  if(max>0) return 'G'+String(max+1).padStart(3,'0');
  return 'G'+String(Date.now()).slice(-6);
}

function json(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* 將 Google Drive 分享連結轉成可直接顯示的格式 */
function normalizeImg(url){
  if(!url) return '';
  const m = String(url).match(/drive\.google\.com\/file\/d\/([^/]+)/);
  if(m) return 'https://drive.google.com/uc?export=view&id='+m[1];
  return url;
}

/* ============ GET ============ */
function doGet(e){
  try{
    const action = (e.parameter.action)||'';
    if(action==='getAll'){
      return json({ items: readItems() });
    }
    if(action==='getLogs'){
      const sh = logsSheet();
      const values = sh.getDataRange().getValues();
      const head = values.shift();
      let logs = values.filter(r=>r[0]!=='').map(r=>{
        const o={}; head.forEach((h,i)=>o[h]=r[i]); return o;
      });
      if(e.parameter.itemId){ logs = logs.filter(l=>String(l['品項ID'])===String(e.parameter.itemId)); }
      logs.reverse();                 // 最新在前
      return json({ logs: logs.slice(0,100) });
    }
    return json({ error: '未知 action' });
  }catch(err){ return json({ error: String(err) }); }
}

/* ============ POST ============ */
function doPost(e){
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try{
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    if(action==='addItem'){
      const sh = itemsSheet();
      const id = nextId();
      sh.appendRow([ id, body.name, Number(body.qty)||0, body.unit||'個',
        Number(body.minAlert)||0, body.note||'', new Date().toISOString(), normalizeImg(body.imgUrl) ]);
      return json({ success:true, id });
    }

    if(action==='updateItem'){
      const row = findRow(body.id);
      if(row<0) return json({ error:'找不到品項' });
      itemsSheet().getRange(row,1,1,ITEM_HEADERS.length).setValues([[
        body.id, body.name, Number(body.qty)||0, body.unit||'個',
        Number(body.minAlert)||0, body.note||'', new Date().toISOString(), normalizeImg(body.imgUrl) ]]);
      return json({ success:true });
    }

    if(action==='deleteItem'){
      const row = findRow(body.id);
      if(row<0) return json({ error:'找不到品項' });
      itemsSheet().deleteRow(row);
      return json({ success:true });
    }

    if(action==='adjustQty'){
      const row = findRow(body.id);
      if(row<0) return json({ error:'找不到品項' });
      const sh = itemsSheet();
      const cur = Number(sh.getRange(row,3).getValue())||0;
      const delta = Number(body.delta)||0;
      const newQty = cur + delta;
      if(newQty < 0) return json({ error:'庫存不足' });
      const name = sh.getRange(row,2).getValue();
      const now = new Date().toISOString();
      sh.getRange(row,3).setValue(newQty);
      sh.getRange(row,7).setValue(now);
      logsSheet().appendRow([ now, body.id, name, delta, newQty, body.note||'', body.operator||'' ]);
      return json({ success:true, newQty });
    }

    return json({ error:'未知 action' });
  }catch(err){
    return json({ error: String(err) });
  }finally{
    lock.releaseLock();
  }
}
