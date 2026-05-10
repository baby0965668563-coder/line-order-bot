require('dotenv').config();

const express = require('express');
const line = require('@line/bot-sdk');
const { GoogleSpreadsheet } = require('google-spreadsheet');

const app = express();

const lineConfig = {
  channelAccessToken: process.env.CHANNEL_ACCESS_TOKEN,
  channelSecret: process.env.CHANNEL_SECRET
};

app.use('/webhook', line.middleware(lineConfig));
app.use('/api', express.json());
app.use('/admin', express.json());

const client = new line.Client(lineConfig);
const doc = new GoogleSpreadsheet(process.env.SHEET_ID);

// ────────────────────────────────────────────────────────────────
//  全域狀態
// ────────────────────────────────────────────────────────────────
let isOpen = false;
let allText = '';
let autoCloseTimer = null;
let autoCloseAt = null;
const knownUsers = {};

const admins = [
  'U8d9c82446aa9eb90d7de001cfc7ea90f',
  'Ubcfae64b443b9fad21bbc584e991b306',
  'U5c44a04efc62664bd45ec80d77be7d93',
  'Uc669eca67bf477460945f45751edd3e9'
];

function isAdmin(userId) { return admins.includes(userId); }

// ────────────────────────────────────────────────────────────────
//  工具函式
// ────────────────────────────────────────────────────────────────
function now() {
  return new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
}

function todayTW() {
  return new Date().toLocaleDateString('zh-TW', { timeZone: 'Asia/Taipei' });
}

// ────────────────────────────────────────────────────────────────
//  Google Sheets 認證（每次呼叫都重新認證，避免 token 過期）
// ────────────────────────────────────────────────────────────────
async function authSheet() {
  await doc.useServiceAccountAuth({
    client_email: process.env.GOOGLE_CLIENT_EMAIL,
    private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n')
  });
  await doc.loadInfo();
}

// ────────────────────────────────────────────────────────────────
//  菜單讀取
// ────────────────────────────────────────────────────────────────
async function loadMenu() {
  await authSheet();
  const sheet = doc.sheetsByTitle['Menu'];
  if (!sheet) return [];
  const rows = await sheet.getRows();
  return rows
    .map(r => ({
      store: String(r['店家'] || '').trim(),
      item:  String(r['品項'] || '').trim(),
      price: Number(r['價格'] || 0)
    }))
    .filter(r => r.store && r.item && r.price > 0);
}

async function loadOptions() {
  await authSheet();
  const groupsSheet  = doc.sheetsByTitle['OptionGroups'];
  const optionsSheet = doc.sheetsByTitle['Options'];
  if (!groupsSheet || !optionsSheet) return {};

  const groupRows  = await groupsSheet.getRows();
  const optionRows = await optionsSheet.getRows();
  const result = {};

  groupRows.forEach(group => {
    const store    = String(group['店家'] || '').trim();
    const item     = String(group['品項'] || '').trim();
    const category = String(group['分類'] || '').trim();
    if (!store || !item || !category) return;

    const key = store + '||' + item;
    if (!result[key]) result[key] = [];

    const options = optionRows
      .filter(opt =>
        String(opt['店家'] || '').trim() === store &&
        String(opt['品項'] || '').trim() === item &&
        String(opt['分類'] || '').trim() === category
      )
      .map(opt => String(opt['選項'] || '').trim())
      .filter(Boolean);

    result[key].push({
      category,
      required: String(group['必選'] || '').trim() === 'TRUE',
      min: Number(group['最少'] || 0),
      max: Number(group['最多'] || 0),
      options
    });
  });
  return result;
}

// ────────────────────────────────────────────────────────────────
//  Users 寫入
// ────────────────────────────────────────────────────────────────
async function saveUserToSheet(profileName, userId, sourceType, groupId) {
  try {
    await authSheet();
    const sheet = doc.sheetsByTitle['Users'];
    if (!sheet) return;
    await sheet.addRow({
      時間:     now(),
      LINE名稱: profileName,
      userId,
      來源類型: sourceType,
      群組ID:   groupId || '',
      權限:     isAdmin(userId) ? 'admin' : 'user'
    });
  } catch (err) {
    console.error('寫入 Users 失敗：', err.message);
  }
}

// ────────────────────────────────────────────────────────────────
//  Orders CRUD
// ────────────────────────────────────────────────────────────────
async function saveOrderToSheet(order) {
  try {
    await authSheet();
    const sheet = doc.sheetsByTitle['Orders'];
    if (!sheet) return { success: false, reason: 'no_sheet' };

    const rows  = await sheet.getRows();
    const today = todayTW();

    // 防重複：同 userId + 品項 + 規格 + 今日 + 非已刪除
    const dup = rows.find(r =>
      String(r['userId'] || '')    === String(order.userId) &&
      String(r['品項']   || '').trim() === String(order.item  || '').trim() &&
      String(r['規格']   || '').trim() === String(order.spec  || '').trim() &&
      String(r['時間']   || '').startsWith(today) &&
      String(r['狀態']   || '') !== '已刪除'
    );
    if (dup) return { success: false, reason: 'duplicate' };

    const qty   = Number(order.qty   || 1);
    const price = Number(order.price || 0);
    await sheet.addRow({
      時間:     now(),
      LINE名稱: String(order.name  || ''),
      userId:   String(order.userId || ''),
      店家:     String(order.store  || ''),
      品項:     String(order.item   || ''),
      規格:     String(order.spec   || ''),
      備註:     String(order.note   || ''),
      數量:     qty,
      單價:     price,
      總價:     price * qty,
      狀態:     '未付款',
      付款時間: '',
      付款方式: '',
      訂單備註: ''
    });
    return { success: true };
  } catch (err) {
    console.error('寫入 Orders 失敗：', err.message);
    return { success: false, reason: err.message };
  }
}

async function getOrdersByUser(userId) {
  try {
    await authSheet();
    const sheet = doc.sheetsByTitle['Orders'];
    if (!sheet) return [];
    const rows  = await sheet.getRows();
    const today = todayTW();
    return rows
      .filter(r =>
        String(r['userId'] || '') === String(userId) &&
        String(r['時間']   || '').startsWith(today) &&
        String(r['狀態']   || '') !== '已刪除'
      )
      .map(r => ({
        rowIndex: r.rowIndex,
        store:    String(r['店家'] || ''),
        item:     String(r['品項'] || ''),
        spec:     String(r['規格'] || ''),
        note:     String(r['備註'] || ''),
        qty:      Number(r['數量'] || 1),
        price:    Number(r['單價'] || 0),
        total:    Number(r['總價'] || 0),
        status:   String(r['狀態'] || '未付款'),
        time:     String(r['時間'] || '')
      }));
  } catch (err) {
    console.error('讀取我的訂單失敗：', err.message);
    return [];
  }
}

async function deleteOrder(userId, rowIndex) {
  try {
    await authSheet();
    const sheet = doc.sheetsByTitle['Orders'];
    if (!sheet) return false;
    const rows   = await sheet.getRows();
    const target = rows.find(r =>
      Number(r.rowIndex) === Number(rowIndex) &&
      String(r['userId'] || '') === String(userId) &&
      String(r['狀態']   || '') !== '已刪除'
    );
    if (!target) return false;
    target['狀態'] = '已刪除';
    await target.save();
    return true;
  } catch (err) {
    console.error('刪除訂單失敗：', err.message);
    return false;
  }
}

async function updateOrderNote(userId, rowIndex, note) {
  try {
    await authSheet();
    const sheet = doc.sheetsByTitle['Orders'];
    if (!sheet) return false;
    const rows   = await sheet.getRows();
    const target = rows.find(r =>
      Number(r.rowIndex) === Number(rowIndex) &&
      String(r['userId'] || '') === String(userId) &&
      String(r['狀態']   || '') !== '已刪除'
    );
    if (!target) return false;
    target['備註'] = note;
    await target.save();
    return true;
  } catch (err) {
    console.error('修改備註失敗：', err.message);
    return false;
  }
}

async function getAllOrdersToday() {
  try {
    await authSheet();
    const sheet = doc.sheetsByTitle['Orders'];
    if (!sheet) return [];
    const rows  = await sheet.getRows();
    const today = todayTW();
    return rows
      .filter(r => String(r['時間'] || '').startsWith(today))
      .map(r => ({
        rowIndex: r.rowIndex,
        name:     String(r['LINE名稱'] || ''),
        userId:   String(r['userId']   || ''),
        store:    String(r['店家']     || ''),
        item:     String(r['品項']     || ''),
        spec:     String(r['規格']     || ''),
        note:     String(r['備註']     || ''),
        qty:      Number(r['數量']     || 1),
        price:    Number(r['單價']     || 0),
        total:    Number(r['總價']     || 0),
        status:   String(r['狀態']     || '未付款'),
        payTime:  String(r['付款時間'] || ''),
        payType:  String(r['付款方式'] || '')
      }));
  } catch (err) {
    console.error('讀取今日訂單失敗：', err.message);
    return [];
  }
}

async function markPaid(rowIndex, payType) {
  try {
    await authSheet();
    const sheet  = doc.sheetsByTitle['Orders'];
    if (!sheet) return false;
    const rows   = await sheet.getRows();
    const target = rows.find(r => Number(r.rowIndex) === Number(rowIndex));
    if (!target) return false;
    target['狀態']    = '已付款';
    target['付款時間'] = now();
    target['付款方式'] = payType || '現金';
    await target.save();
    return true;
  } catch (err) {
    console.error('標記付款失敗：', err.message);
    return false;
  }
}

async function batchMarkPaid(payType) {
  try {
    await authSheet();
    const sheet = doc.sheetsByTitle['Orders'];
    if (!sheet) return 0;
    const rows  = await sheet.getRows();
    const today = todayTW();
    let count   = 0;
    for (const r of rows) {
      if (
        String(r['時間']  || '').startsWith(today) &&
        String(r['狀態']  || '') === '未付款'
      ) {
        r['狀態']    = '已付款';
        r['付款時間'] = now();
        r['付款方式'] = payType || '現金';
        await r.save();
        count++;
      }
    }
    return count;
  } catch (err) {
    console.error('批次付款失敗：', err.message);
    return 0;
  }
}

async function adminDeleteOrder(rowIndex) {
  try {
    await authSheet();
    const sheet  = doc.sheetsByTitle['Orders'];
    if (!sheet) return false;
    const rows   = await sheet.getRows();
    const target = rows.find(r => Number(r.rowIndex) === Number(rowIndex));
    if (!target) return false;
    target['狀態'] = '已刪除';
    await target.save();
    return true;
  } catch (err) {
    console.error('Admin 刪除訂單失敗：', err.message);
    return false;
  }
}

// ────────────────────────────────────────────────────────────────
//  自動結單
// ────────────────────────────────────────────────────────────────
function scheduleAutoClose(minutes) {
  if (autoCloseTimer) clearTimeout(autoCloseTimer);
  const ms = Math.max(1, Number(minutes)) * 60 * 1000;
  autoCloseAt = new Date(Date.now() + ms).toISOString();
  autoCloseTimer = setTimeout(() => {
    isOpen        = false;
    autoCloseAt   = null;
    autoCloseTimer = null;
    console.log('[自動結單] 已自動結單 at', now());
  }, ms);
}

function cancelAutoClose() {
  if (autoCloseTimer) clearTimeout(autoCloseTimer);
  autoCloseTimer = null;
  autoCloseAt    = null;
}

// ────────────────────────────────────────────────────────────────
//  文字訂單解析（原有保留）
// ────────────────────────────────────────────────────────────────
// ────────────────────────────────────────────────────────────────
//  從 Orders Sheet 統計今日訂單（排除「已刪除」）
// ────────────────────────────────────────────────────────────────

/**
 * 統計報表（LINE 用）
 * 品項數量 / 每人金額 / 未付款名單 / 總金額
 */
async function buildStatReport() {
  const orders = await getAllOrdersToday();
  const active = orders.filter(o => o.status !== '已刪除');
  if (active.length === 0) return '📊 今日尚無訂單';

  const itemCount = {};   // { '品項（規格）': qty }
  const userTotal = {};   // { name: total }
  const unpaidSet = new Set();

  for (const o of active) {
    const key  = o.item + (o.spec ? '（' + o.spec + '）' : '');
    itemCount[key] = (itemCount[key] || 0) + o.qty;

    const name = o.name || o.userId || '未知';
    userTotal[name] = (userTotal[name] || 0) + o.total;

    if (o.status === '未付款') unpaidSet.add(name);
  }

  const grandTotal = Object.values(userTotal).reduce((a, b) => a + b, 0);

  let text = '📊 今日訂餐統計\n';
  text += '─────────────\n';
  text += '【品項數量】\n';
  for (const k in itemCount) text += k + ' x' + itemCount[k] + '\n';

  text += '\n【個人金額】\n';
  for (const name in userTotal) text += name + '：$' + userTotal[name] + '\n';

  text += '\n💰 總金額：$' + grandTotal;

  if (unpaidSet.size > 0) {
    text += '\n\n⚠️ 未付款：' + [...unpaidSet].join('、');
  } else {
    text += '\n\n✅ 所有人已付款';
  }

  return text;
}

/**
 * 店家單（LINE 用）
 * 品項＋規格合併數量
 */
async function buildShopOrder() {
  const orders = await getAllOrdersToday();
  const active = orders.filter(o => o.status !== '已刪除');
  if (active.length === 0) return '今日尚無訂單';

  const itemCount = {};
  for (const o of active) {
    const key = o.item + (o.spec ? '（' + o.spec + '）' : '');
    itemCount[key] = (itemCount[key] || 0) + o.qty;
  }

  let out = '您好，今天訂購如下：\n\n';
  let totalCount = 0;
  for (const k in itemCount) {
    out += k + ' x' + itemCount[k] + '\n';
    totalCount += itemCount[k];
  }

  const totalMoney = active.reduce((a, o) => a + o.total, 0);
  out += '\n總數：' + totalCount + '份';
  out += '\n總金額：' + totalMoney + '元';
  out += '\n\n麻煩您，謝謝～';
  return out;
}

// ────────────────────────────────────────────────────────────────
//  管理員 Middleware
// ────────────────────────────────────────────────────────────────
function adminAuth(req, res, next) {
  const token = req.headers['x-admin-token'] || req.query.token || '';
  if (!process.env.ADMIN_TOKEN || token !== process.env.ADMIN_TOKEN) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  next();
}

// ════════════════════════════════════════════════════════════════
//  路由
// ════════════════════════════════════════════════════════════════

// ── 健康檢查 ──────────────────────────────────────────────────────
app.get('/', (_req, res) => res.send('LINE 訂餐機器人運作中'));

// ── 菜單 API（前端 fetch，不內嵌 JSON）────────────────────────────
app.get('/api/menu', async (_req, res) => {
  try {
    const menu       = await loadMenu();
    const optionData = await loadOptions();
    res.json({ menu, optionData });
  } catch (err) {
    console.error('載入菜單失敗：', err.message);
    res.status(500).json({ menu: [], optionData: {} });
  }
});

// ── 狀態 API ──────────────────────────────────────────────────────
app.get('/api/status', (_req, res) => res.json({ isOpen, autoCloseAt }));

// ── 訂餐 LIFF 頁（純靜態，不帶 JSON）─────────────────────────────
app.get('/order', (_req, res) => {
  const LIFF_ID = process.env.LIFF_ID || '2010025093-yATK02dc';
  // 只插入純英數 LIFF_ID，絕對安全
  const safeLiffId = String(LIFF_ID).replace(/[^a-zA-Z0-9\-]/g, '');
  res.setHeader('Content-Type', 'text/html; charset=utf-8');
  res.end(orderPageHtml(safeLiffId));
});

function orderPageHtml(liffId) {
  return [
    '<!DOCTYPE html>',
    '<html><head>',
    '<meta charset="utf-8">',
    '<title>訂餐小幫手</title>',
    '<meta name="viewport" content="width=device-width, initial-scale=1">',
    '<script src="https://static.line-scdn.net/liff/edge/2/sdk.js"></script>',
    '<style>',
    '*{box-sizing:border-box}',
    'body{margin:0;font-family:Arial,"Microsoft JhengHei",sans-serif;background:#f6f3ee;color:#333}',
    '.header{padding:16px;background:#fff;border-bottom:1px solid #eee;position:sticky;top:0;z-index:10}',
    '.header h1{margin:0;font-size:24px}',
    '.header p{margin:6px 0 0;color:#777;font-size:14px}',
    '.status{margin-top:8px;font-size:14px;color:#06c755}',
    '.countdown{margin-top:4px;font-size:13px;color:#e53935;font-weight:bold}',
    '.tabs{background:#f6f3ee;position:sticky;top:112px;z-index:9;padding:8px 12px;overflow-x:auto;white-space:nowrap;border-bottom:1px solid #eee}',
    '.tab-btn{display:inline-block;margin:3px;padding:8px 14px;border-radius:999px;background:#fff;color:#333;border:1px solid #ddd;font-size:14px;cursor:pointer}',
    '.tab-btn.active{background:#06c755;color:#fff;border-color:#06c755}',
    '.store-title{font-size:20px;font-weight:bold;margin:20px 0 10px;padding:8px 4px}',
    '.card{background:#fff;border-radius:16px;padding:14px;margin-bottom:10px;box-shadow:0 2px 8px rgba(0,0,0,.05)}',
    '.store-label{font-size:12px;color:#999;margin-bottom:4px}',
    '.item-name{font-size:18px;font-weight:bold;margin-bottom:6px}',
    '.price{font-size:16px;margin-bottom:10px}',
    '.btn{width:100%;padding:12px;border:none;border-radius:999px;background:#06c755;color:#fff;font-size:15px;font-weight:bold;cursor:pointer}',
    '.btn:disabled{background:#aaa;cursor:default}',
    '.btn.secondary{background:#999}',
    '.cart-fab{','  position:fixed !important;','  bottom:24px !important;','  right:20px !important;','  background:#06c755;','  color:#fff;','  border:none;','  border-radius:50px;','  padding:14px 18px;','  font-size:15px;','  font-weight:bold;','  box-shadow:0 4px 16px rgba(0,0,0,.2);','  cursor:pointer;','  z-index:99999 !important;','  display:flex !important;','  align-items:center;','  gap:4px;','}' ,'.fab-hidden{display:none !important;}',
    '.cart-fab.order-btn{right:160px !important;}',
    '.badge{background:#e53935;color:#fff;border-radius:50%;font-size:11px;padding:2px 6px;margin-left:6px}',
    '.modal{display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,.5);z-index:9999;justify-content:center;align-items:flex-end}',
    '.modal.center{align-items:center}',
    '.modal-box{background:#fff;width:100%;max-width:480px;max-height:85vh;overflow-y:auto;border-radius:20px 20px 0 0;padding:20px}',
    '.modal.center .modal-box{border-radius:20px;margin:0 12px}',
    '.opt-label{display:block;margin:10px 0;font-size:15px;cursor:pointer}',
    '.qty-box{margin-top:16px}',
    '.qty-box select,.note-input{width:100%;padding:11px;border-radius:12px;border:1px solid #ddd;font-size:15px;margin-top:6px}',
    '.row{display:flex;align-items:center;gap:10px;padding:12px 0;border-bottom:1px solid #f0f0f0}',
    '.row-info{flex:1}',
    '.row-name{font-weight:bold;font-size:15px}',
    '.row-sub{font-size:13px;color:#777;margin-top:2px}',
    '.row-price{font-size:15px;font-weight:bold;color:#06c755;white-space:nowrap}',
    '.icon-btn{background:none;border:none;font-size:20px;cursor:pointer;padding:4px 8px}',
    '.tag{display:inline-block;padding:2px 8px;border-radius:999px;font-size:12px}',
    '.tag.paid{background:#e8f5e9;color:#2e7d32}',
    '.tag.unpaid{background:#fff3e0;color:#e65100}',
    '.empty{text-align:center;color:#999;margin:24px 0;font-size:15px}',
    '.container{padding:14px;padding-bottom:120px}',
    '#loadingMsg{text-align:center;padding:40px 20px;color:#999;font-size:15px}',
    '</style></head><body>',

    '<div class="header">',
    '  <h1>🍱 訂餐小幫手</h1>',
    '  <p id="orderStatus">訂單狀態載入中...</p>',
    '  <div class="status" id="status">正在取得 LINE 使用者資料...</div>',
    '  <div class="countdown" id="countdown"></div>',
    '</div>',
    '<div class="tabs" id="storeTabs"></div>',
    '<div class="container" id="menuBox"><div id="loadingMsg">菜單載入中...</div></div>',

    '<button class="cart-fab fab-hidden" id="cartFab" onclick="openCart()">🛒 購物車<span class="badge" id="cartBadge">0</span></button>',
    '<button class="cart-fab order-btn fab-hidden" id="orderFab" onclick="openModal(\'myOrderModal\')">📋 我的訂單</button>',

    // 選項 Modal
    '<div class="modal" id="optionModal"><div class="modal-box">',
    '  <h2 id="modalTitle" style="margin-top:0"></h2>',
    '  <div id="modalOptions"></div>',
    '  <div class="qty-box"><div style="font-weight:bold">數量</div>',
    '  <select id="qtySelect"><option>1</option><option>2</option><option>3</option><option>4</option><option>5</option><option>6</option><option>7</option><option>8</option><option>9</option><option>10</option></select></div>',
    '  <div class="qty-box"><div style="font-weight:bold">備註（選填）</div>',
    '  <input class="note-input" id="itemNote" placeholder="例：不加辣、少冰..."></div>',
    '  <button class="btn" style="margin-top:16px" onclick="submitOptions()">加入購物車</button>',
    '  <button class="btn secondary" style="margin-top:10px" onclick="closeModal(\'optionModal\')">取消</button>',
    '</div></div>',

    // 購物車 Modal
    '<div class="modal" id="cartModal"><div class="modal-box">',
    '  <h2 style="margin-top:0">🛒 我的購物車</h2>',
    '  <div id="cartList"></div>',
    '  <div id="cartTotal" style="font-size:17px;font-weight:bold;margin:14px 0"></div>',
    '  <button class="btn" id="submitCartBtn" onclick="submitCart()">送出訂單</button>',
    '  <button class="btn secondary" style="margin-top:10px" onclick="closeModal(\'cartModal\')">繼續點餐</button>',
    '</div></div>',

    // 我的訂單 Modal
    '<div class="modal" id="myOrderModal"><div class="modal-box">',
    '  <h2 style="margin-top:0">📋 我的訂單</h2>',
    '  <div id="myOrderList"></div>',
    '  <div id="myOrderTotal" style="font-size:16px;font-weight:bold;margin:12px 0"></div>',
    '  <button class="btn secondary" style="margin-top:10px" onclick="closeModal(\'myOrderModal\')">關閉</button>',
    '</div></div>',

    // 備註修改 Modal
    '<div class="modal center" id="editNoteModal"><div class="modal-box">',
    '  <h3 style="margin-top:0">修改備註</h3>',
    '  <input class="note-input" id="editNoteInput" placeholder="備註內容">',
    '  <button class="btn" style="margin-top:12px" onclick="saveNote()">儲存</button>',
    '  <button class="btn secondary" style="margin-top:8px" onclick="closeModal(\'editNoteModal\')">取消</button>',
    '</div></div>',

    '<script>',
    // 用字串拼接傳入 LIFF_ID，絕對不會有 JSON/template 問題
    'var LIFF_ID="' + liffId + '";',
    'var menu=[];',
    'var optionData={};',
    'var profile=null;',
    'var liffReady=false;',
    'var currentItem=null;',
    'var currentGroups=[];',
    'var cart=[];',
    'var editingRowIndex=null;',

    // loadMenuData
    'async function loadMenuData(){',
    '  try{',
    '    var res=await fetch("/api/menu");',
    '    var data=await res.json();',
    '    menu=data.menu||[];',
    '    optionData=data.optionData||{};',
    '    renderMenu();',
    '  }catch(e){',
    '    var el=document.getElementById("loadingMsg");',
    '    if(el)el.innerText="菜單載入失敗，請重新整理";',
    '  }',
    '}',

    // initLIFF
    'async function initLIFF(){',
    '  try{',
    '    await liff.init({liffId:LIFF_ID});',
    '    if(!liff.isLoggedIn()){liff.login();return;}',
    '    profile=await liff.getProfile();',
    '    liffReady=true;',
    '    document.getElementById("status").innerText="已登入："+profile.displayName;',
    '    enableButtons();',
    '    document.getElementById("cartFab").classList.remove("fab-hidden");',
    '    document.getElementById("orderFab").classList.remove("fab-hidden");',
    '    checkOrderStatus();',
    '    loadMyOrders();',
    '  }catch(err){',
    '    console.error("LIFF error:",err);',
    '    document.getElementById("status").innerText="LIFF 初始化失敗："+err.message;',
    '  }',
    '}',

    // checkOrderStatus
    'async function checkOrderStatus(){',
    '  try{',
    '    var res=await fetch("/api/status");',
    '    var data=await res.json();',
    '    document.getElementById("orderStatus").innerText=data.isOpen?"🟢 目前開放點餐":"🔴 目前未開放點餐";',
    '    if(data.autoCloseAt)startCountdown(data.autoCloseAt);',
    '    else document.getElementById("countdown").innerText="";',
    '  }catch(e){}',
    '}',

    // startCountdown
    'function startCountdown(isoStr){',
    '  var el=document.getElementById("countdown");',
    '  function tick(){',
    '    var diff=new Date(isoStr)-new Date();',
    '    if(diff<=0){el.innerText="";return;}',
    '    var m=Math.floor(diff/60000),s=Math.floor((diff%60000)/1000);',
    '    el.innerText="⏰ 自動結單："+m+"分"+s+"秒後";',
    '    setTimeout(tick,1000);',
    '  }',
    '  tick();',
    '}',

    // renderMenu
    'function renderMenu(){',
    '  var box=document.getElementById("menuBox");',
    '  var tabsBox=document.getElementById("storeTabs");',
    '  var lm=document.getElementById("loadingMsg");',
    '  if(lm)lm.style.display="none";',
    '  if(!menu||menu.length===0){box.innerHTML=\'<div class="empty">目前沒有菜單資料</div>\';return;}',
    '  var stores=[];',
    '  menu.forEach(function(m){if(stores.indexOf(m.store)<0)stores.push(m.store);});',
    '  tabsBox.innerHTML=stores.map(function(s,i){',
    '    return \'<button class="tab-btn" id="tab\'+i+\'" onclick="scrollToStore(\'+i+\')">\'+escH(s)+\'</button>\';',
    '  }).join("");',
    '  var html="",cur="";',
    '  menu.forEach(function(m,idx){',
    '    if(m.store!==cur){',
    '      cur=m.store;',
    '      var si=stores.indexOf(cur);',
    '      html+=\'<div id="store\'+si+\'" class="store-title">\'+escH(cur)+\'</div>\';',
    '    }',
    '    html+=\'<div class="card">\'+',
    '      \'<div class="store-label">\'+escH(m.store)+\'</div>\'+',
    '      \'<div class="item-name">\'+escH(m.item)+\'</div>\'+',
    '      \'<div class="price">$\'+Number(m.price)+\'</div>\'+',
    '      \'<button class="btn" id="btn\'+idx+\'" onclick="addToCart(\'+idx+\')" disabled>登入中...</button>\'+',
    '      \'</div>\';',
    '  });',
    '  box.innerHTML=html;',
    '  if(liffReady)enableButtons();',
    '}',

    // scrollToStore
    'function scrollToStore(idx){',
    '  document.querySelectorAll(".tab-btn").forEach(function(b,i){b.classList.toggle("active",i===idx);});',
    '  var el=document.getElementById("store"+idx);',
    '  if(el)el.scrollIntoView({behavior:"smooth",block:"start"});',
    '}',

    // enableButtons
    'function enableButtons(){',
    '  menu.forEach(function(_,idx){',
    '    var btn=document.getElementById("btn"+idx);',
    '    if(btn){btn.disabled=false;btn.innerText="加入購物車";}',
    '  });',
    '}',

    // addToCart
    'function addToCart(idx){',
    '  if(!liffReady||!profile){alert("尚未取得 LINE 使用者資料");return;}',
    '  currentItem=menu[idx];',
    '  var key=currentItem.store+"||"+currentItem.item;',
    '  currentGroups=optionData[key]||[];',
    '  document.getElementById("modalTitle").innerText=currentItem.item;',
    '  document.getElementById("qtySelect").value="1";',
    '  document.getElementById("itemNote").value="";',
    '  var optBox=document.getElementById("modalOptions");',
    '  optBox.innerHTML="";',
    '  if(currentGroups.length===0){',
    '    optBox.innerHTML=\'<div style="color:#777;margin:10px 0">此商品無需選擇規格</div>\';',
    '  }else{',
    '    currentGroups.forEach(function(group,gi){',
    '      var title=document.createElement("div");',
    '      title.style.marginTop="16px";',
    '      title.innerHTML="<b>"+escH(group.category)+"</b><br>請選 "+group.min+"~"+group.max+" 個";',
    '      optBox.appendChild(title);',
    '      group.options.forEach(function(opt){',
    '        var label=document.createElement("label");',
    '        label.className="opt-label";',
    '        var cb=document.createElement("input");',
    '        cb.type="checkbox"; cb.value=opt; cb.dataset.group=gi;',
    '        cb.addEventListener("change",function(){limitCheck(gi,group.max);});',
    '        label.appendChild(cb);',
    '        label.appendChild(document.createTextNode(" "+opt));',
    '        optBox.appendChild(label);',
    '      });',
    '    });',
    '  }',
    '  openModal("optionModal");',
    '}',

    // limitCheck
    'function limitCheck(gi,max){',
    '  var checked=[].slice.call(document.querySelectorAll(\'input[data-group="\'+gi+\'"]:checked\'));',
    '  [].slice.call(document.querySelectorAll(\'input[data-group="\'+gi+\'"]\'))',
    '    .forEach(function(x){x.disabled=checked.length>=max&&!x.checked;});',
    '}',

    // submitOptions
    'function submitOptions(){',
    '  var specParts=[];',
    '  for(var i=0;i<currentGroups.length;i++){',
    '    var group=currentGroups[i];',
    '    var checked=[].slice.call(document.querySelectorAll(\'input[data-group="\'+i+\'"]:checked\'));',
    '    if(checked.length<group.min||checked.length>group.max){',
    '      alert(group.category+" 需要選 "+group.min+"~"+group.max+" 個");return;',
    '    }',
    '    if(checked.length>0)specParts.push(group.category+"："+checked.map(function(x){return x.value;}).join("、"));',
    '  }',
    '  cart.push({',
    '    store:currentItem.store,item:currentItem.item,',
    '    spec:specParts.join(" "),',
    '    note:document.getElementById("itemNote").value.trim(),',
    '    qty:Number(document.getElementById("qtySelect").value)||1,',
    '    price:currentItem.price',
    '  });',
    '  closeModal("optionModal");',
    '  updateCartBadge();',
    '  showToast("已加入購物車 🎉");',
    '}',

    // updateCartBadge
    'function updateCartBadge(){',
    '  document.getElementById("cartBadge").innerText=cart.reduce(function(a,c){return a+c.qty;},0);',
    '}',

    // openCart
    'function openCart(){renderCart();openModal("cartModal");}',

    // renderCart
    'function renderCart(){',
    '  var listEl=document.getElementById("cartList");',
    '  var totalEl=document.getElementById("cartTotal");',
    '  if(cart.length===0){listEl.innerHTML=\'<div class="empty">購物車是空的</div>\';totalEl.innerText="";return;}',
    '  listEl.innerHTML=cart.map(function(c,i){',
    '    var sub=[c.spec,c.note].filter(Boolean).join("｜");',
    '    return \'<div class="row">\'+',
    '      \'<div class="row-info"><div class="row-name">\'+escH(c.item)+\' x\'+c.qty+\'</div>\'+',
    '      (sub?\'<div class="row-sub">\'+escH(sub)+\'</div>\':"")+\'</div>\'+',
    '      \'<div class="row-price">$\'+(c.price*c.qty)+\'</div>\'+',
    '      \'<button class="icon-btn" onclick="removeCartItem(\'+i+\')">🗑</button></div>\';',
    '  }).join("");',
    '  totalEl.innerText="合計：$"+cart.reduce(function(a,c){return a+c.price*c.qty;},0);',
    '}',

    // removeCartItem
    'function removeCartItem(i){cart.splice(i,1);updateCartBadge();renderCart();}',

    // submitCart
    'async function submitCart(){',
    '  if(cart.length===0){alert("購物車是空的");return;}',
    '  var btn=document.getElementById("submitCartBtn");',
    '  btn.disabled=true;btn.innerText="送出中...";',
    '  var allOk=true;',
    '  for(var i=0;i<cart.length;i++){',
    '    var c=cart[i];',
    '    try{',
    '      var res=await fetch("/api/order",{',
    '        method:"POST",',
    '        headers:{"Content-Type":"application/json"},',
    '        body:JSON.stringify({',
    '          store:c.store,item:c.item,spec:c.spec,note:c.note,qty:c.qty,price:c.price,',
    '          name:profile.displayName,userId:profile.userId',
    '        })',
    '      });',
    '      var result=await res.json();',
    '      if(!result.success){',
    '        if(result.reason==="duplicate")showToast("⚠️ "+c.item+" 已送出過，略過");',
    '        else allOk=false;',
    '      }',
    '    }catch(e){allOk=false;}',
    '  }',
    '  cart=[];updateCartBadge();closeModal("cartModal");',
    '  btn.disabled=false;btn.innerText="送出訂單";',
    '  showToast(allOk?"訂單已送出 ✅":"部分品項送出失敗，請重試");',
    '  loadMyOrders();',
    '}',

    // loadMyOrders
    'async function loadMyOrders(){',
    '  if(!profile)return;',
    '  try{',
    '    var res=await fetch("/api/my-orders?userId="+encodeURIComponent(profile.userId));',
    '    renderMyOrders(await res.json());',
    '  }catch(e){}',
    '}',

    // renderMyOrders
    'function renderMyOrders(orders){',
    '  var listEl=document.getElementById("myOrderList");',
    '  var totalEl=document.getElementById("myOrderTotal");',
    '  if(!orders||orders.length===0){listEl.innerHTML=\'<div class="empty">今天還沒有訂單</div>\';totalEl.innerText="";return;}',
    '  listEl.innerHTML=orders.map(function(o){',
    '    var sub=[o.spec,o.note].filter(Boolean).join("｜");',
    '    var tag=o.status==="已付款"?\'<span class="tag paid">已付款</span>\':\'<span class="tag unpaid">未付款</span>\';',
    '    var canEdit=o.status!=="已付款";',
    '    return \'<div class="row">\'+',
    '      \'<div class="row-info"><div class="row-name">\'+escH(o.item)+\' x\'+o.qty+\' \'+tag+\'</div>\'+',
    '      (sub?\'<div class="row-sub">\'+escH(sub)+\'</div>\':"")+',
    '      \'<div class="row-sub">\'+escH(o.store)+\'</div></div>\'+',
    '      \'<div style="display:flex;flex-direction:column;align-items:flex-end;gap:4px">\'+',
    '      \'<div class="row-price">$\'+o.total+\'</div>\'+',
    '      (canEdit?',
    '        \'<button class="icon-btn" onclick="editNote(\'+o.rowIndex+\',\'+JSON.stringify(o.note||"")+\')" title="修改備註">✏️</button>\'+',
    '        \'<button class="icon-btn" onclick="delOrder(\'+o.rowIndex+\')" title="刪除">🗑</button>\'',
    '      :"")+',
    '      \'</div></div>\';',
    '  }).join("");',
    '  totalEl.innerText="今日合計：$"+orders.reduce(function(a,o){return a+o.total;},0);',
    '}',

    // delOrder
    'async function delOrder(rowIndex){',
    '  if(!confirm("確定要刪除這筆訂單？"))return;',
    '  try{',
    '    var res=await fetch("/api/order/"+rowIndex,{',
    '      method:"DELETE",headers:{"Content-Type":"application/json"},',
    '      body:JSON.stringify({userId:profile.userId})',
    '    });',
    '    var r=await res.json();',
    '    if(r.success){showToast("已刪除");loadMyOrders();}',
    '    else alert("刪除失敗");',
    '  }catch(e){alert("刪除失敗："+e.message);}',
    '}',

    // editNote
    'function editNote(rowIndex,currentNote){',
    '  editingRowIndex=rowIndex;',
    '  document.getElementById("editNoteInput").value=currentNote||"";',
    '  openModal("editNoteModal");',
    '}',

    // saveNote
    'async function saveNote(){',
    '  var note=document.getElementById("editNoteInput").value.trim();',
    '  try{',
    '    var res=await fetch("/api/order/"+editingRowIndex+"/note",{',
    '      method:"PATCH",headers:{"Content-Type":"application/json"},',
    '      body:JSON.stringify({userId:profile.userId,note:note})',
    '    });',
    '    var r=await res.json();',
    '    if(r.success){showToast("備註已更新");closeModal("editNoteModal");loadMyOrders();}',
    '    else alert("更新失敗");',
    '  }catch(e){alert("更新失敗："+e.message);}',
    '}',

    // openModal / closeModal
    'function openModal(id){',
    '  document.getElementById(id).style.display="flex";',
    '  if(id==="myOrderModal")loadMyOrders();',
    '}',
    'function closeModal(id){document.getElementById(id).style.display="none";}',

    // showToast
    'function showToast(msg){',
    '  var t=document.getElementById("_toast");',
    '  if(!t){t=document.createElement("div");t.id="_toast";',
    '  t.style.cssText="position:fixed;bottom:90px;left:50%;transform:translateX(-50%);background:#333;color:#fff;padding:10px 20px;border-radius:999px;font-size:14px;z-index:9999;opacity:0;transition:opacity .3s;pointer-events:none";',
    '  document.body.appendChild(t);}',
    '  t.innerText=msg;t.style.opacity="1";',
    '  setTimeout(function(){t.style.opacity="0";},2200);',
    '}',

    // escH
    'function escH(s){',
    '  return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");',
    '}',

    // init
    'loadMenuData();',
    'initLIFF();',
    'setInterval(checkOrderStatus,30000);',

    '</script></body></html>'
  ].join('\n');
}

// ── 管理後台（純靜態 HTML，token 從 URL 讀，不內嵌到 JS）─────────
app.get('/admin', (req, res) => {
  const token = req.query.token || '';
  if (!process.env.ADMIN_TOKEN || token !== process.env.ADMIN_TOKEN) {
    return res.status(401).send('Unauthorized. 請附上 ?token=YOUR_ADMIN_TOKEN');
  }
  res.setHeader('Content-Type', 'text/html; charset=utf-8');
  res.end(adminPageHtml());
});

function adminPageHtml() {
  return [
    '<!DOCTYPE html><html><head>',
    '<meta charset="utf-8"><title>訂餐管理後台</title>',
    '<meta name="viewport" content="width=device-width, initial-scale=1">',
    '<style>',
    '*{box-sizing:border-box}',
    'body{margin:0;font-family:Arial,"Microsoft JhengHei",sans-serif;background:#f5f5f5;color:#333}',
    '.header{background:#222;color:#fff;padding:16px 20px;position:sticky;top:0;z-index:10}',
    '.header h1{margin:0;font-size:20px}',
    '.header p{margin:4px 0 0;font-size:13px;opacity:.7}',
    '.container{padding:16px;max-width:960px;margin:0 auto}',
    '.card{background:#fff;border-radius:12px;padding:16px;margin-bottom:16px;box-shadow:0 2px 8px rgba(0,0,0,.06)}',
    '.btn{padding:10px 16px;border:none;border-radius:999px;background:#333;color:#fff;font-size:14px;cursor:pointer;margin:4px}',
    '.btn.green{background:#06c755}',
    '.btn.red{background:#e53935}',
    '.btn.orange{background:#f57c00}',
    '.btn.sm{padding:6px 12px;font-size:12px}',
    '.btn:disabled{background:#aaa}',
    '.sbar{display:flex;align-items:center;gap:10px;flex-wrap:wrap}',
    '.dot{width:12px;height:12px;border-radius:50%;background:#aaa;display:inline-block}',
    '.dot.open{background:#06c755}',
    'table{width:100%;border-collapse:collapse;font-size:13px}',
    'th,td{padding:9px 8px;border-bottom:1px solid #eee;text-align:left;vertical-align:middle}',
    'th{background:#f9f9f9;font-weight:bold}',
    '.tag{display:inline-block;padding:2px 8px;border-radius:999px;font-size:12px}',
    '.tag.paid{background:#e8f5e9;color:#2e7d32}',
    '.tag.unpaid{background:#fff3e0;color:#e65100}',
    '.tag.del{background:#fce4ec;color:#c62828}',
    '.summary{display:flex;gap:12px;flex-wrap:wrap}',
    '.sbox{background:#f9f9f9;border-radius:10px;padding:12px 16px;flex:1;min-width:120px}',
    '.sbox .num{font-size:26px;font-weight:bold}',
    '.sbox .lbl{font-size:12px;color:#777;margin-top:2px}',
    'input[type=number]{width:64px;padding:6px;border-radius:8px;border:1px solid #ddd;font-size:14px}',
    'select.psel{padding:6px;border-radius:8px;border:1px solid #ddd;font-size:13px}',
    '.search{width:100%;padding:10px;border-radius:10px;border:1px solid #ddd;font-size:14px;margin-bottom:12px}',
    '@media(max-width:600px){th,td{padding:6px 4px;font-size:12px}}',
    '</style></head><body>',
    '<div class="header"><h1>🍱 訂餐管理後台</h1><p id="aTime"></p></div>',
    '<div class="container">',

    // 開單控制
    '<div class="card">',
    '<h3 style="margin-top:0">📢 開單控制</h3>',
    '<div class="sbar">',
    '  <span class="dot" id="sDot"></span>',
    '  <span id="sTxt" style="font-size:15px;font-weight:bold"></span>',
    '  <button class="btn green" onclick="openOrder()">開單</button>',
    '  <button class="btn red"   onclick="closeOrder()">結單</button>',
    '  <button class="btn orange" onclick="clearState()">清空狀態</button>',
    '</div>',
    '<div style="margin-top:14px;display:flex;align-items:center;gap:8px;flex-wrap:wrap">',
    '  <label style="font-size:14px">自動結單：</label>',
    '  <input type="number" id="autoMin" value="30" min="1" max="480">',
    '  <span style="font-size:14px">分鐘後</span>',
    '  <button class="btn" onclick="setAutoClose()">設定</button>',
    '  <button class="btn" onclick="doCancelAutoClose()">取消自動結單</button>',
    '  <span id="acInfo" style="font-size:13px;color:#e53935;font-weight:bold"></span>',
    '</div></div>',

    // 統計
    '<div class="card"><h3 style="margin-top:0">📊 今日統計</h3>',
    '<div class="summary" id="sumBox">載入中...</div></div>',

    // 訂單列表
    '<div class="card"><h3 style="margin-top:0">📋 今日訂單</h3>',
    '<input class="search" id="search" placeholder="搜尋姓名、品項..." oninput="filterTable()">',
    '<div style="margin-bottom:10px;display:flex;gap:8px;flex-wrap:wrap">',
    '  <button class="btn green" onclick="batchPaid()">✅ 全部標記已付款</button>',
    '  <button class="btn" onclick="loadOrders()">🔄 重新整理</button>',
    '  <button class="btn" onclick="copyShopOrder()">📋 複製店家單</button>',
    '</div>',
    '<div style="overflow-x:auto">',
    '<table><thead><tr>',
    '<th>姓名</th><th>品項</th><th>規格</th><th>備註</th><th>數量</th><th>金額</th><th>狀態</th><th>操作</th>',
    '</tr></thead>',
    '<tbody id="orderBody"><tr><td colspan="8" style="text-align:center;padding:20px;color:#999">載入中...</td></tr></tbody>',
    '</table></div></div>',
    '</div>',

    '<script>',
    // TOKEN 從 URL 讀，不內嵌
    'var TOKEN=(location.search.match(/[?&]token=([^&]*)/)||[])[1]||"";',
    'var ordersCache=[];',
    'var acAt=null;',
    'var acTimer=null;',
    'document.getElementById("aTime").innerText=new Date().toLocaleString("zh-TW");',

    // api helper
    'async function api(path,method,body){',
    '  var opts={method:method||"GET",headers:{"Content-Type":"application/json","x-admin-token":TOKEN}};',
    '  if(body)opts.body=JSON.stringify(body);',
    '  var res=await fetch(path,opts);',
    '  return res.json();',
    '}',

    // loadStatus
    'async function loadStatus(){',
    '  var data=await api("/api/status");',
    '  document.getElementById("sDot").className="dot"+(data.isOpen?" open":"");',
    '  document.getElementById("sTxt").innerText=data.isOpen?"開單中":"已結單";',
    '  if(data.autoCloseAt){acAt=data.autoCloseAt;startAcCountdown();}',
    '  else{acAt=null;document.getElementById("acInfo").innerText="";}',
    '}',

    // startAcCountdown
    'function startAcCountdown(){',
    '  if(acTimer)clearInterval(acTimer);',
    '  acTimer=setInterval(function(){',
    '    if(!acAt){clearInterval(acTimer);return;}',
    '    var diff=new Date(acAt)-new Date();',
    '    if(diff<=0){document.getElementById("acInfo").innerText="已自動結單";clearInterval(acTimer);loadStatus();return;}',
    '    var m=Math.floor(diff/60000),s=Math.floor((diff%60000)/1000);',
    '    document.getElementById("acInfo").innerText="⏰ "+m+"分"+s+"秒後自動結單";',
    '  },1000);',
    '}',

    'async function openOrder(){await api("/api/admin/open","POST");loadStatus();loadOrders();}',
    'async function closeOrder(){await api("/api/admin/close","POST");loadStatus();loadOrders();}',
    'async function clearState(){if(!confirm("確定清空今日開單狀態？不會刪除 Sheet 資料"))return;await api("/api/admin/clear","POST");loadStatus();loadOrders();}',
    'async function setAutoClose(){var m=Number(document.getElementById("autoMin").value)||30;await api("/api/admin/auto-close","POST",{minutes:m});loadStatus();}',
    'async function doCancelAutoClose(){await api("/api/admin/cancel-auto-close","POST");acAt=null;document.getElementById("acInfo").innerText="";if(acTimer)clearInterval(acTimer);}',

    // loadOrders
    'async function loadOrders(){',
    '  var data=await api("/api/admin/orders");',
    '  ordersCache=Array.isArray(data)?data:[];',
    '  renderTable(ordersCache);',
    '  renderSummary(ordersCache);',
    '}',

    // renderSummary
    'function renderSummary(orders){',
    '  var active=orders.filter(function(o){return o.status!=="已刪除";});',
    '  var paid=active.filter(function(o){return o.status==="已付款";});',
    '  var unpaid=active.filter(function(o){return o.status==="未付款";});',
    '  var total=active.reduce(function(a,o){return a+o.total;},0);',
    '  var paidM=paid.reduce(function(a,o){return a+o.total;},0);',
    '  function box(n,l){return \'<div class="sbox"><div class="num">\'+n+\'</div><div class="lbl">\'+l+\'</div></div>\';}',
    '  document.getElementById("sumBox").innerHTML=',
    '    box(active.length,"筆訂單")+box(total,"總金額 $")+box(paid.length,"已付款")+box(unpaid.length,"未付款")+box(paidM,"已收 $");',
    '}',

    // renderTable
    'function renderTable(orders){',
    '  var body=document.getElementById("orderBody");',
    '  if(!orders||orders.length===0){body.innerHTML=\'<tr><td colspan="8" style="text-align:center;padding:20px;color:#999">今日無訂單</td></tr>\';return;}',
    '  body.innerHTML=orders.map(function(o){',
    '    var tc=o.status==="已付款"?"paid":o.status==="已刪除"?"del":"unpaid";',
    '    var ops=o.status==="未付款"?',
    '      \'<select class="psel" id="pt\'+o.rowIndex+\'"><option>現金</option><option>Line Pay</option><option>轉帳</option></select>\'+',
    '      \'<button class="btn green sm" onclick="markPaid(\'+o.rowIndex+\')">付款</button>\'+',
    '      \'<button class="btn red sm" onclick="adminDel(\'+o.rowIndex+\')">刪</button>\'',
    '      :"";',
    '    return \'<tr data-name="\'+esc(o.name)+\'" data-item="\'+esc(o.item)+\'">\'+',
    '      \'<td>\'+esc(o.name)+\'</td><td>\'+esc(o.item)+\'</td>\'+',
    '      \'<td style="color:#777">\'+esc(o.spec)+\'</td><td style="color:#777">\'+esc(o.note)+\'</td>\'+',
    '      \'<td>\'+o.qty+\'</td><td>$\'+o.total+\'</td>\'+',
    '      \'<td><span class="tag \'+tc+\'">\'+esc(o.status)+\'</span></td>\'+',
    '      \'<td>\'+ops+\'</td></tr>\';',
    '  }).join("");',
    '}',

    // filterTable
    'function filterTable(){',
    '  var q=document.getElementById("search").value.toLowerCase();',
    '  [].slice.call(document.querySelectorAll("#orderBody tr[data-name]")).forEach(function(r){',
    '    r.style.display=(!q||(r.dataset.name||"").toLowerCase().indexOf(q)>=0||(r.dataset.item||"").toLowerCase().indexOf(q)>=0)?"":"none";',
    '  });',
    '}',

    // markPaid
    'async function markPaid(rowIndex){',
    '  var sel=document.getElementById("pt"+rowIndex);',
    '  var pt=sel?sel.value:"現金";',
    '  await api("/api/admin/paid","POST",{rowIndex:rowIndex,payType:pt});',
    '  loadOrders();',
    '}',

    // adminDel
    'async function adminDel(rowIndex){',
    '  if(!confirm("確定刪除？"))return;',
    '  await api("/api/admin/delete-order","POST",{rowIndex:rowIndex});',
    '  loadOrders();',
    '}',

    // batchPaid
    'async function batchPaid(){',
    '  if(!confirm("將所有未付款訂單標記為已付款？"))return;',
    '  await api("/api/admin/batch-paid","POST",{payType:"現金"});',
    '  loadOrders();',
    '}',

    // copyShopOrder
    'function copyShopOrder(){',
    '  var active=ordersCache.filter(function(o){return o.status!=="已刪除";});',
    '  var counts={};',
    '  active.forEach(function(o){',
    '    var k=o.item+(o.spec?"（"+o.spec+"）":"");',
    '    counts[k]=(counts[k]||0)+o.qty;',
    '  });',
    '  var text="您好，今天訂購如下：\n\n";',
    '  var total=0;',
    '  Object.keys(counts).forEach(function(k){text+=k+" x"+counts[k]+"\n";total+=counts[k];});',
    '  var money=active.reduce(function(a,o){return a+o.total;},0);',
    '  text+="\n總數："+total+"份\n總金額："+money+"元\n\n麻煩您，謝謝～";',
    '  navigator.clipboard.writeText(text).then(function(){alert("已複製到剪貼板");});',
    '}',

    // esc
    'function esc(s){return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");}',

    // init
    'loadStatus();',
    'loadOrders();',
    'setInterval(loadStatus,30000);',
    'setInterval(loadOrders,60000);',
    '</script></body></html>'
  ].join('\n');
}

// ════════════════════════════════════════════════════════════════
//  使用者 API
// ════════════════════════════════════════════════════════════════
app.post('/api/order', async (req, res) => {
  const result = await saveOrderToSheet(req.body);
  res.json(result);
});

app.get('/api/my-orders', async (req, res) => {
  const { userId } = req.query;
  if (!userId) return res.json([]);
  res.json(await getOrdersByUser(userId));
});

app.delete('/api/order/:rowIndex', async (req, res) => {
  const rowIndex = Number(req.params.rowIndex);
  const { userId } = req.body;
  if (!userId) return res.json({ success: false, reason: 'no_user' });
  res.json({ success: await deleteOrder(userId, rowIndex) });
});

app.patch('/api/order/:rowIndex/note', async (req, res) => {
  const rowIndex = Number(req.params.rowIndex);
  const { userId, note } = req.body;
  if (!userId) return res.json({ success: false, reason: 'no_user' });
  res.json({ success: await updateOrderNote(userId, rowIndex, note) });
});

// ════════════════════════════════════════════════════════════════
//  管理員 API
// ════════════════════════════════════════════════════════════════
app.post('/api/admin/open', adminAuth, (_req, res) => {
  if (!isOpen) { isOpen = true; allText = ''; }
  res.json({ isOpen });
});

app.post('/api/admin/close', adminAuth, (_req, res) => {
  isOpen = false;
  cancelAutoClose();
  res.json({ isOpen });
});

app.post('/api/admin/clear', adminAuth, (_req, res) => {
  isOpen = false;
  allText = '';
  cancelAutoClose();
  res.json({ ok: true });
});

app.post('/api/admin/auto-close', adminAuth, (req, res) => {
  const { minutes } = req.body;
  if (!isOpen) { isOpen = true; allText = ''; }
  scheduleAutoClose(Number(minutes) || 30);
  res.json({ autoCloseAt });
});

app.post('/api/admin/cancel-auto-close', adminAuth, (_req, res) => {
  cancelAutoClose();
  res.json({ ok: true });
});

app.get('/api/admin/orders', adminAuth, async (_req, res) => {
  res.json(await getAllOrdersToday());
});

app.post('/api/admin/paid', adminAuth, async (req, res) => {
  const { rowIndex, payType } = req.body;
  res.json({ success: await markPaid(rowIndex, payType) });
});

app.post('/api/admin/batch-paid', adminAuth, async (req, res) => {
  const count = await batchMarkPaid(req.body.payType || '現金');
  res.json({ success: true, count });
});

app.post('/api/admin/delete-order', adminAuth, async (req, res) => {
  const { rowIndex } = req.body;
  res.json({ success: await adminDeleteOrder(rowIndex) });
});

// ════════════════════════════════════════════════════════════════
//  LINE Webhook
// ════════════════════════════════════════════════════════════════
app.post('/webhook', async (req, res) => {
  try {
    const event = req.body.events?.[0];
    if (!event || event.type !== 'message' || event.message.type !== 'text') {
      return res.sendStatus(200);
    }

    let profileName = '未知使用者';
    try {
      if (event.source.type === 'group') {
        const p = await client.getGroupMemberProfile(event.source.groupId, event.source.userId);
        profileName = p.displayName;
      } else {
        const p = await client.getProfile(event.source.userId);
        profileName = p.displayName;
      }
    } catch (e) { console.error('取得使用者名稱失敗：', e.message); }

    knownUsers[profileName] = event.source.userId;
    // 非同步寫入 Users，不等它
    saveUserToSheet(profileName, event.source.userId, event.source.type, event.source.groupId || '').catch(() => {});

    const userId     = event.source.userId;
    const text       = event.message.text.trim();
    const replyToken = event.replyToken;

    const reply = async (t) => {
      await client.replyMessage(replyToken, { type: 'text', text: t });
    };

    // ── 開單 N 分鐘（要在「開單」完整比對之前，避免被截）
    const autoOpenMatch = text.match(/^開單\s+(\d+)$/);
    if (autoOpenMatch) {
      if (!isAdmin(userId)) { await reply('只有管理員可以開單'); return res.sendStatus(200); }
      isOpen = true; allText = '';
      scheduleAutoClose(Number(autoOpenMatch[1]));
      await reply('已開單！將於 ' + autoOpenMatch[1] + ' 分鐘後自動結單\n點餐頁：' + (process.env.LIFF_URL || ''));
      return res.sendStatus(200);
    }

    if (text === '開單') {
      if (!isAdmin(userId)) { await reply('只有管理員可以開單'); return res.sendStatus(200); }
      if (isOpen) { await reply('目前已開單中'); return res.sendStatus(200); }
      isOpen = true; allText = '';
      await reply('已開單，可以開始點餐 🍱\n點餐頁：' + (process.env.LIFF_URL || ''));
      return res.sendStatus(200);
    }

    if (text === '結單' || text === '收單' || text === '統計') {
      if (!isAdmin(userId)) { await reply('只有管理員可以結單 / 統計'); return res.sendStatus(200); }
      isOpen = false; cancelAutoClose();
      await reply(await buildStatReport());
      return res.sendStatus(200);
    }

    if (text === '店家單') {
      if (!isAdmin(userId)) { await reply('只有管理員可以查看店家單'); return res.sendStatus(200); }
      await reply(await buildShopOrder());
      return res.sendStatus(200);
    }

    if (text === '清空') {
      if (!isAdmin(userId)) { await reply('只有管理員可以清空'); return res.sendStatus(200); }
      isOpen = false; cancelAutoClose();
      await reply('已清空訂單');
      return res.sendStatus(200);
    }

    if (text === '後台') {
      if (!isAdmin(userId)) { await reply('只有管理員可以查看後台'); return res.sendStatus(200); }
      const url = (process.env.APP_URL || '') + '/admin?token=' + (process.env.ADMIN_TOKEN || '');
      await reply('管理後台：' + url);
      return res.sendStatus(200);
    }

    if (text === '狀態') {
      const msg = (isOpen ? '🟢 目前開單中' : '🔴 目前未開單') +
        (autoCloseAt ? '\n⏰ 自動結單：' + new Date(autoCloseAt).toLocaleTimeString('zh-TW', { timeZone: 'Asia/Taipei' }) : '');
      await reply(msg);
      return res.sendStatus(200);
    }

    // allText 已廢棄，訂單統一從 Orders Sheet 讀取

    return res.sendStatus(200);
  } catch (error) {
    console.error('Webhook error:', error);
    return res.sendStatus(200);
  }
});

// ── 全域錯誤處理 ──────────────────────────────────────────────────
app.use((err, _req, res, _next) => {
  console.error('Global error:', err);
  res.sendStatus(200);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log('Server running on port', PORT));
