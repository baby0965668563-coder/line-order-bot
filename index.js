require('dotenv').config();

const express = require('express');
const line = require('@line/bot-sdk');
const { GoogleSpreadsheet } = require('google-spreadsheet');

const app = express();

const config = {
  channelAccessToken: process.env.CHANNEL_ACCESS_TOKEN,
  channelSecret: process.env.CHANNEL_SECRET
};

app.use('/webhook', line.middleware(config));
app.use(express.json());

const client = new line.Client(config);
const doc = new GoogleSpreadsheet(process.env.SHEET_ID);

// ─── 全域狀態 ───────────────────────────────────────────────────
let isOpen = false;
let allText = '';
let autoCloseTimer = null;
let autoCloseAt = null;       // ISO string，給前端顯示用
const knownUsers = {};

const admins = [
  "U8d9c82446aa9eb90d7de001cfc7ea90f",
  "Ubcfae64b443b9fad21bbc584e991b306",
  "U5c44a04efc62664bd45ec80d77be7d93",
  "Uc669eca67bf477460945f45751edd3e9"
];

function isAdmin(userId) {
  return admins.includes(userId);
}

// ─── Google Sheets 認證 ──────────────────────────────────────────
async function authSheet() {
  await doc.useServiceAccountAuth({
    client_email: process.env.GOOGLE_CLIENT_EMAIL,
    private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n')
  });
  await doc.loadInfo();
}

// ─── 菜單 / 選項讀取（原有邏輯保留）────────────────────────────
async function loadMenu() {
  await authSheet();
  const sheet = doc.sheetsByTitle['Menu'];
  if (!sheet) return [];
  const rows = await sheet.getRows();
  return rows
    .map(r => ({
      store: String(r['店家'] || '').trim(),
      item: String(r['品項'] || '').trim(),
      price: Number(r['價格'] || 0)
    }))
    .filter(row => row.store && row.item && row.price > 0);
}

async function loadOptions() {
  await authSheet();
  const groupsSheet = doc.sheetsByTitle['OptionGroups'];
  const optionsSheet = doc.sheetsByTitle['Options'];
  if (!groupsSheet || !optionsSheet) return {};
  const groupRows = await groupsSheet.getRows();
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

// ─── Users ───────────────────────────────────────────────────────
async function saveUserToSheet(profileName, userId, sourceType, groupId) {
  try {
    await authSheet();
    const sheet = doc.sheetsByTitle['Users'];
    if (!sheet) return;
    await sheet.addRow({
      時間: now(),
      LINE名稱: profileName,
      userId,
      來源類型: sourceType,
      群組ID: groupId || '',
      權限: isAdmin(userId) ? 'admin' : 'user'
    });
  } catch (err) {
    console.error('寫入 Users 失敗：', err.message);
  }
}

// ─── Orders CRUD ─────────────────────────────────────────────────
async function saveOrderToSheet(order) {
  try {
    await authSheet();
    const sheet = doc.sheetsByTitle['Orders'];
    if (!sheet) return { success: false };

    // 防重複：同一 userId + item + spec 在同一天不能重複（isOpen 期間）
    const rows = await sheet.getRows();
    const today = new Date().toLocaleDateString('zh-TW', { timeZone: 'Asia/Taipei' });
    const dup = rows.find(r =>
      String(r['userId'] || '') === order.userId &&
      String(r['品項'] || '').trim() === order.item &&
      String(r['規格'] || '').trim() === (order.spec || '') &&
      String(r['時間'] || '').startsWith(today) &&
      String(r['狀態'] || '') !== '已刪除'
    );
    if (dup) return { success: false, reason: 'duplicate' };

    const qty   = Number(order.qty || 1);
    const price = Number(order.price || 0);
    const rowData = {
      時間:    now(),
      LINE名稱: order.name || '',
      userId:  order.userId || '',
      店家:    order.store || '',
      品項:    order.item || '',
      規格:    order.spec || '',
      備註:    order.note || '',
      數量:    qty,
      單價:    price,
      總價:    price * qty,
      狀態:    '未付款',
      // 新增欄位
      付款時間: '',
      付款方式: '',
      訂單備註: ''
    };
    await sheet.addRow(rowData);
    return { success: true };
  } catch (err) {
    console.error('寫入 Orders 失敗：', err.message);
    return { success: false, reason: err.message };
  }
}

async function getOrdersByUser(userId) {
  await authSheet();
  const sheet = doc.sheetsByTitle['Orders'];
  if (!sheet) return [];
  const rows = await sheet.getRows();
  const today = new Date().toLocaleDateString('zh-TW', { timeZone: 'Asia/Taipei' });
  return rows
    .filter(r =>
      String(r['userId'] || '') === userId &&
      String(r['時間'] || '').startsWith(today) &&
      String(r['狀態'] || '') !== '已刪除'
    )
    .map((r, i) => ({
      rowIndex: r.rowIndex,
      store:  String(r['店家'] || ''),
      item:   String(r['品項'] || ''),
      spec:   String(r['規格'] || ''),
      note:   String(r['備註'] || ''),
      qty:    Number(r['數量'] || 1),
      price:  Number(r['單價'] || 0),
      total:  Number(r['總價'] || 0),
      status: String(r['狀態'] || '未付款'),
      time:   String(r['時間'] || '')
    }));
}

async function deleteOrder(userId, rowIndex) {
  try {
    await authSheet();
    const sheet = doc.sheetsByTitle['Orders'];
    if (!sheet) return false;
    const rows = await sheet.getRows();
    const target = rows.find(r =>
      r.rowIndex === rowIndex &&
      String(r['userId'] || '') === userId &&
      String(r['狀態'] || '') !== '已刪除'
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
    const rows = await sheet.getRows();
    const target = rows.find(r =>
      r.rowIndex === rowIndex &&
      String(r['userId'] || '') === userId &&
      String(r['狀態'] || '') !== '已刪除'
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

// ─── Admin：取得今日所有訂單 ─────────────────────────────────────
async function getAllOrdersToday() {
  await authSheet();
  const sheet = doc.sheetsByTitle['Orders'];
  if (!sheet) return [];
  const rows = await sheet.getRows();
  const today = new Date().toLocaleDateString('zh-TW', { timeZone: 'Asia/Taipei' });
  return rows
    .filter(r => String(r['時間'] || '').startsWith(today))
    .map(r => ({
      rowIndex: r.rowIndex,
      name:    String(r['LINE名稱'] || ''),
      userId:  String(r['userId'] || ''),
      store:   String(r['店家'] || ''),
      item:    String(r['品項'] || ''),
      spec:    String(r['規格'] || ''),
      note:    String(r['備註'] || ''),
      qty:     Number(r['數量'] || 1),
      price:   Number(r['單價'] || 0),
      total:   Number(r['總價'] || 0),
      status:  String(r['狀態'] || '未付款'),
      payTime: String(r['付款時間'] || ''),
      payType: String(r['付款方式'] || '')
    }));
}

async function markPaid(rowIndex, payType) {
  try {
    await authSheet();
    const sheet = doc.sheetsByTitle['Orders'];
    if (!sheet) return false;
    const rows = await sheet.getRows();
    const target = rows.find(r => r.rowIndex === rowIndex);
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

async function markPaidByUser(userId, payType) {
  try {
    await authSheet();
    const sheet = doc.sheetsByTitle['Orders'];
    if (!sheet) return 0;
    const rows = await sheet.getRows();
    const today = new Date().toLocaleDateString('zh-TW', { timeZone: 'Asia/Taipei' });
    let count = 0;
    for (const r of rows) {
      if (
        String(r['userId'] || '') === userId &&
        String(r['時間'] || '').startsWith(today) &&
        String(r['狀態'] || '') === '未付款'
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

// ─── 自動結單 ─────────────────────────────────────────────────────
function scheduleAutoClose(minutes) {
  if (autoCloseTimer) clearTimeout(autoCloseTimer);
  const ms = minutes * 60 * 1000;
  autoCloseAt = new Date(Date.now() + ms).toISOString();
  autoCloseTimer = setTimeout(async () => {
    if (!isOpen) return;
    isOpen = false;
    autoCloseAt = null;
    autoCloseTimer = null;
    console.log('[自動結單] 已自動結單');
    // 若有群組 ID 可在此推播通知，目前僅 log
  }, ms);
}

function cancelAutoClose() {
  if (autoCloseTimer) clearTimeout(autoCloseTimer);
  autoCloseTimer = null;
  autoCloseAt = null;
}

// ─── 原有文字解析（保留）────────────────────────────────────────
function clean(text) {
  return String(text || '')
    .replace(/[。.,，、!！?？:：;；"'（）()【】\[\]{}<>《》\s]/g, '')
    .trim();
}

function parseOrders(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  let currentItem = '', currentPrice = 0, pendingOrder = null, itemBuffer = '';
  const priceMap = {}, itemCount = {}, userTotal = {};

  function cleanItemName(t) { return clean(t).replace(/\d+顆/g, ''); }
  function add(item, price, name, qty = 1, note = '') {
    item = cleanItemName(item); name = clean(name);
    if (!item || !name || !qty) return;
    const finalItem = note ? `${item}（${note}）` : item;
    itemCount[finalItem] = (itemCount[finalItem] || 0) + qty;
    userTotal[name] = (userTotal[name] || 0) + price * qty;
  }
  function getPrice(rawQty) {
    if (priceMap[rawQty] !== undefined) return priceMap[rawQty];
    if (rawQty === '半' && priceMap['0.5'] !== undefined) return priceMap['0.5'];
    return currentPrice;
  }

  for (let line of lines) {
    if (/今天有人要訂|收錢|謝謝|下午|早上|晚上|星期/.test(line)) continue;
    const priceTable = line.match(/^(半|0\.5|1)\s*[💰$＄]\s*(\d{1,5})$/);
    if (priceTable) { priceMap[priceTable[1]] = Number(priceTable[2]); continue; }
    if (pendingOrder && !/[+*]/.test(line) && !/\d/.test(line)) {
      add(pendingOrder.item, pendingOrder.price, line, 1); pendingOrder = null; continue;
    }
    const orderMatch = line.match(/^(.+?)[+*]\s*(半|0\.5|\d+)(.*)$/);
    if (orderMatch && currentItem) {
      const name = orderMatch[1], rawQty = orderMatch[2], extra = orderMatch[3] || '';
      const note = extra.includes('辣') ? '辣' : '';
      const price = getPrice(rawQty);
      const qty = rawQty === '半' || rawQty === '0.5' ? 1 : Number(rawQty);
      add(currentItem, price, name, qty, note); continue;
    }
    const inlineSymbol = line.match(/^(.+?)\s*[💰$＄]\s*(\d{1,5})\s*([^\d\s]+)$/);
    if (inlineSymbol) { add(inlineSymbol[1], Number(inlineSymbol[2]), inlineSymbol[3], 1); itemBuffer = ''; continue; }
    const itemSymbol = line.match(/^(.+?)\s*[💰$＄]\s*(\d{1,5})$/);
    if (itemSymbol) { currentItem = itemSymbol[1]; currentPrice = Number(itemSymbol[2]); itemBuffer = ''; continue; }
    const inlineNoSymbol = line.match(/^(.+?)(\d{2,5})([^\d\s]+)$/);
    if (inlineNoSymbol && !/[+*]/.test(line)) { add(inlineNoSymbol[1], Number(inlineNoSymbol[2]), inlineNoSymbol[3], 1); itemBuffer = ''; continue; }
    const noSymbolNoName = line.match(/^(.+?)(\d{2,5})$/);
    if (noSymbolNoName && !/[+*]/.test(line) && !/顆/.test(line)) { pendingOrder = { item: noSymbolNoName[1], price: Number(noSymbolNoName[2]) }; itemBuffer = ''; continue; }
    const priceNameOnly = line.match(/^(\d{2,5})\s*([^\d\s]+)$/);
    if (priceNameOnly && itemBuffer) { add(itemBuffer, Number(priceNameOnly[1]), priceNameOnly[2], 1); itemBuffer = ''; continue; }
    const priceOnly = line.match(/^(\d{2,5})$/);
    if (priceOnly && itemBuffer) { pendingOrder = { item: itemBuffer, price: Number(priceOnly[1]) }; itemBuffer = ''; continue; }
    if (!/[+*]/.test(line)) {
      if (/顆/.test(line)) { currentItem = line; currentPrice = 0; itemBuffer = line; continue; }
      if (currentItem) { add(currentItem, currentPrice, line, 1); continue; }
      currentItem = line; currentPrice = 0; itemBuffer = line; continue;
    }
  }
  return { itemCount, userTotal };
}

function formatResult(itemCount, userTotal) {
  let text = '📊 訂餐統計\n\n【品項數量】\n';
  for (let item in itemCount) if (itemCount[item] > 0) text += `${item} x${itemCount[item]}\n`;
  text += '\n【個人金額】\n';
  for (let user in userTotal) text += `${user}：$${userTotal[user]}\n`;
  text += `\n💰 總金額：$${Object.values(userTotal).reduce((a, b) => a + b, 0)}`;
  return text;
}

function formatShopOrder(itemCount, userTotal) {
  let orderText = '您好，今天訂購如下：\n\n';
  let totalCount = 0;
  for (let item in itemCount) {
    const qty = itemCount[item];
    if (qty > 0) { orderText += `${item} x${qty}\n`; totalCount += qty; }
  }
  const totalMoney = Object.values(userTotal).reduce((a, b) => a + b, 0);
  return orderText + `\n總數：${totalCount}份\n總金額：${totalMoney}元\n\n麻煩您，謝謝～`;
}

// ─── 工具 ─────────────────────────────────────────────────────────
function now() {
  return new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
}

function safeJson(data) {
  return JSON.stringify(data)
    .replace(/</g, '\\u003c')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');
}

// ─── 健康檢查 ─────────────────────────────────────────────────────
app.get('/', (req, res) => {
  res.send('LINE 訂餐統計機器人運作中');
});

// ─── 訂餐 LIFF 頁 ─────────────────────────────────────────────────
app.get('/order', async (req, res) => {
  try {
    const menu = await loadMenu();
    const optionData = await loadOptions();
    const menuJson = safeJson(menu);
    const optionJson = safeJson(optionData);

    const html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>訂餐小幫手</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <script src="https://static.line-scdn.net/liff/edge/2/sdk.js"><\/script>
  <style>
    *{box-sizing:border-box}
    body{margin:0;font-family:Arial,"Microsoft JhengHei",sans-serif;background:#f6f3ee;color:#333}
    .header{padding:16px;background:#fff;border-bottom:1px solid #eee;position:sticky;top:0;z-index:10}
    .header h1{margin:0;font-size:24px}
    .header p{margin:6px 0 0;color:#777;font-size:14px}
    .status{margin-top:8px;font-size:14px;color:#06c755}
    .countdown{margin-top:4px;font-size:13px;color:#e53935;font-weight:bold}
    .tabs{background:#f6f3ee;position:sticky;top:110px;z-index:9;padding:8px 12px;overflow-x:auto;white-space:nowrap;border-bottom:1px solid #eee}
    .tab-btn{display:inline-block;margin:3px;padding:8px 14px;border-radius:999px;background:#fff;color:#333;border:1px solid #ddd;font-size:14px;cursor:pointer}
    .tab-btn.active{background:#06c755;color:#fff;border-color:#06c755}
    .store-title{font-size:20px;font-weight:bold;margin:20px 0 10px;padding:8px 4px}
    .card{background:#fff;border-radius:16px;padding:14px;margin-bottom:10px;box-shadow:0 2px 8px rgba(0,0,0,.05)}
    .store{font-size:12px;color:#999;margin-bottom:4px}
    .item-name{font-size:18px;font-weight:bold;margin-bottom:6px}
    .price{font-size:16px;margin-bottom:10px}
    .btn{width:100%;padding:12px;border:none;border-radius:999px;background:#06c755;color:#fff;font-size:15px;font-weight:bold;cursor:pointer}
    .btn:disabled{background:#aaa}
    .btn.danger{background:#e53935}
    .btn.secondary{background:#999}
    .btn.outline{background:#fff;color:#06c755;border:2px solid #06c755}
    .empty{text-align:center;color:#999;margin-top:60px;font-size:16px}
    /* 購物車 */
    .cart-fab{position:fixed;bottom:24px;right:20px;background:#06c755;color:#fff;border:none;border-radius:50px;padding:14px 20px;font-size:15px;font-weight:bold;box-shadow:0 4px 16px rgba(0,0,0,.2);cursor:pointer;z-index:200}
    .badge{background:#e53935;color:#fff;border-radius:50%;font-size:11px;padding:2px 6px;margin-left:6px}
    /* Modal */
    .modal{display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,.5);z-index:9999;justify-content:center;align-items:flex-end}
    .modal.center{align-items:center}
    .modal-box{background:#fff;width:100%;max-width:480px;max-height:85vh;overflow-y:auto;border-radius:20px 20px 0 0;padding:20px}
    .modal.center .modal-box{border-radius:20px;margin:0 12px}
    .option-label{display:block;margin:10px 0;font-size:15px}
    .qty-box{margin-top:16px}
    .qty-box select,.note-input{width:100%;padding:11px;border-radius:12px;border:1px solid #ddd;font-size:15px;margin-top:6px}
    /* 購物車列表 */
    .cart-item{display:flex;align-items:center;gap:10px;padding:12px 0;border-bottom:1px solid #f0f0f0}
    .cart-info{flex:1}
    .cart-name{font-weight:bold;font-size:15px}
    .cart-sub{font-size:13px;color:#777;margin-top:2px}
    .cart-price{font-size:15px;font-weight:bold;color:#06c755}
    .icon-btn{background:none;border:none;font-size:20px;cursor:pointer;padding:4px 8px;border-radius:8px}
    /* 管理後台 tab */
    .admin-tabs{display:flex;gap:8px;margin-bottom:16px;overflow-x:auto}
    .admin-tab{padding:8px 16px;border-radius:999px;border:1px solid #ddd;background:#fff;cursor:pointer;font-size:14px;white-space:nowrap}
    .admin-tab.active{background:#333;color:#fff;border-color:#333}
    table{width:100%;border-collapse:collapse;font-size:13px}
    th,td{padding:8px 6px;border-bottom:1px solid #eee;text-align:left}
    th{background:#f9f9f9;font-weight:bold}
    .tag{display:inline-block;padding:2px 8px;border-radius:999px;font-size:12px}
    .tag.paid{background:#e8f5e9;color:#2e7d32}
    .tag.unpaid{background:#fff3e0;color:#e65100}
    .tag.del{background:#fce4ec;color:#c62828}
    .container{padding:14px;padding-bottom:100px}
  </style>
</head>
<body>
  <div class="header">
    <h1>🍱 訂餐小幫手</h1>
    <p id="orderStatus">訂單狀態載入中...</p>
    <div class="status" id="status">正在取得 LINE 使用者資料...</div>
    <div class="countdown" id="countdown"></div>
  </div>
  <div class="tabs" id="storeTabs"></div>
  <div class="container" id="menu"></div>

  <!-- 購物車 FAB -->
  <button class="cart-fab" id="cartFab" onclick="openCart()" style="display:none">
    🛒 購物車<span class="badge" id="cartBadge">0</span>
  </button>

  <!-- 商品選項 Modal -->
  <div class="modal" id="optionModal">
    <div class="modal-box">
      <h2 id="modalTitle" style="margin-top:0"></h2>
      <div id="modalOptions"></div>
      <div class="qty-box">
        <div style="font-weight:bold">數量</div>
        <select id="qtySelect">
          ${[1,2,3,4,5,6,7,8,9,10].map(n=>`<option value="${n}">${n}份</option>`).join('')}
        </select>
      </div>
      <div class="qty-box">
        <div style="font-weight:bold">備註</div>
        <input class="note-input" id="itemNote" placeholder="例：不加辣、少冰...">
      </div>
      <button class="btn" style="margin-top:16px" onclick="submitOptions()">加入購物車</button>
      <button class="btn secondary" style="margin-top:10px" onclick="closeModal('optionModal')">取消</button>
    </div>
  </div>

  <!-- 購物車 Modal -->
  <div class="modal" id="cartModal">
    <div class="modal-box">
      <h2 style="margin-top:0">🛒 我的購物車</h2>
      <div id="cartList"></div>
      <div id="cartTotal" style="font-size:17px;font-weight:bold;margin:14px 0"></div>
      <button class="btn" onclick="submitCart()" id="submitCartBtn">送出訂單</button>
      <button class="btn secondary" style="margin-top:10px" onclick="closeModal('cartModal')">繼續點餐</button>
    </div>
  </div>

  <!-- 已送出訂單 Modal -->
  <div class="modal" id="myOrderModal">
    <div class="modal-box">
      <h2 style="margin-top:0">📋 我的訂單</h2>
      <div id="myOrderList"></div>
      <div id="myOrderTotal" style="font-size:16px;font-weight:bold;margin:12px 0"></div>
      <button class="btn secondary" style="margin-top:10px" onclick="closeModal('myOrderModal')">關閉</button>
    </div>
  </div>

  <!-- 備註修改 Modal -->
  <div class="modal center" id="editNoteModal">
    <div class="modal-box">
      <h3 style="margin-top:0">修改備註</h3>
      <input class="note-input" id="editNoteInput" placeholder="備註內容">
      <button class="btn" style="margin-top:12px" onclick="saveNote()">儲存</button>
      <button class="btn secondary" style="margin-top:8px" onclick="closeModal('editNoteModal')">取消</button>
    </div>
  </div>

<script>
const menu = ${menuJson};
const optionData = ${optionJson};
const LIFF_ID = '2010025093-yATK02dc';

let profile = null;
let liffReady = false;
let currentItem = null;
let currentGroups = [];
let cart = [];  // { store, item, spec, note, qty, price }
let editingRowIndex = null;

// ── LIFF 初始化 ──────────────────────────────────────────────────
async function initLIFF() {
  try {
    await liff.init({ liffId: LIFF_ID });
    if (!liff.isLoggedIn()) { liff.login(); return; }
    profile = await liff.getProfile();
    liffReady = true;
    document.getElementById('status').innerText = '已登入：' + profile.displayName;
    enableButtons();
    document.getElementById('cartFab').style.display = '';
    checkOrderStatus();
    loadMyOrders();
  } catch (err) {
    document.getElementById('status').innerText = 'LIFF 初始化失敗，請重新整理';
  }
}

// ── 訂單狀態 ─────────────────────────────────────────────────────
async function checkOrderStatus() {
  try {
    const res = await fetch('/api/status');
    const data = await res.json();
    const el = document.getElementById('orderStatus');
    el.innerText = data.isOpen ? '🟢 目前開放點餐' : '🔴 目前未開放點餐';
    if (data.autoCloseAt) {
      startCountdown(data.autoCloseAt);
    }
  } catch(e) {}
}

function startCountdown(isoStr) {
  const el = document.getElementById('countdown');
  function tick() {
    const diff = new Date(isoStr) - new Date();
    if (diff <= 0) { el.innerText = ''; return; }
    const m = Math.floor(diff / 60000);
    const s = Math.floor((diff % 60000) / 1000);
    el.innerText = '⏰ 自動結單：' + m + '分' + s + '秒後';
    setTimeout(tick, 1000);
  }
  tick();
}

// ── 菜單渲染 ─────────────────────────────────────────────────────
function renderMenu() {
  const box = document.getElementById('menu');
  const tabsBox = document.getElementById('storeTabs');
  if (!menu || menu.length === 0) { box.innerHTML = '<div class="empty">目前沒有菜單資料</div>'; return; }
  const stores = [...new Set(menu.map(m => m.store).filter(Boolean))];
  tabsBox.innerHTML = stores.map((s, i) =>
    '<button class="tab-btn" id="tab-'+i+'" onclick="scrollToStore('+i+')">' + s + '</button>'
  ).join('');
  let html = '', currentStore = '';
  menu.forEach((m, index) => {
    if (m.store !== currentStore) {
      currentStore = m.store;
      const si = stores.indexOf(currentStore);
      html += '<div id="store-'+si+'" class="store-title">'+currentStore+'</div>';
    }
    html += '<div class="card">'
      + '<div class="store">' + (m.store||'') + '</div>'
      + '<div class="item-name">' + (m.item||'') + '</div>'
      + '<div class="price">$' + (m.price||0) + '</div>'
      + '<button class="btn" onclick="addToCart('+index+')" id="btn-'+index+'" disabled>載入中...</button>'
      + '</div>';
  });
  box.innerHTML = html;
}

function scrollToStore(index) {
  document.getElementById('tab-' + index) && document.querySelectorAll('.tab-btn').forEach((b,i)=>{
    b.classList.toggle('active', i === index);
  });
  const el = document.getElementById('store-' + index);
  if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function enableButtons() {
  menu.forEach((_, index) => {
    const btn = document.getElementById('btn-' + index);
    if (btn) { btn.disabled = false; btn.innerText = '加入購物車'; }
  });
}

// ── 加入購物車（開 Modal）────────────────────────────────────────
function addToCart(index) {
  if (!liffReady || !profile) { alert('尚未取得 LINE 使用者資料'); return; }
  currentItem = menu[index];
  const key = currentItem.store + '||' + currentItem.item;
  currentGroups = optionData[key] || [];
  document.getElementById('modalTitle').innerText = currentItem.item;
  document.getElementById('qtySelect').value = '1';
  document.getElementById('itemNote').value = '';
  const optionBox = document.getElementById('modalOptions');
  optionBox.innerHTML = '';
  if (currentGroups.length === 0) {
    optionBox.innerHTML = '<div style="color:#777;margin:10px 0">此商品無需選擇規格</div>';
  } else {
    currentGroups.forEach((group, gi) => {
      const title = document.createElement('div');
      title.style.marginTop = '16px';
      title.innerHTML = '<b>' + group.category + '</b><br>請選 ' + group.min + '~' + group.max + ' 個';
      optionBox.appendChild(title);
      group.options.forEach(opt => {
        const label = document.createElement('label');
        label.className = 'option-label';
        label.innerHTML = '<input type="checkbox" value="' + opt + '" data-group="' + gi + '" onchange="limitCheck(' + gi + ',' + group.max + ')"> ' + opt;
        optionBox.appendChild(label);
      });
    });
  }
  openModal('optionModal');
}

function limitCheck(gi, max) {
  const checked = [...document.querySelectorAll('input[data-group="'+gi+'"]:checked')];
  const all = [...document.querySelectorAll('input[data-group="'+gi+'"]')];
  all.forEach(x => { x.disabled = checked.length >= max && !x.checked; });
}

function submitOptions() {
  let specText = '';
  for (let i = 0; i < currentGroups.length; i++) {
    const group = currentGroups[i];
    const checked = [...document.querySelectorAll('input[data-group="'+i+'"]:checked')];
    if (checked.length < group.min || checked.length > group.max) {
      alert(group.category + ' 需要選 ' + group.min + '~' + group.max + ' 個'); return;
    }
    specText += group.category + '：' + checked.map(x => x.value).join('、') + ' ';
  }
  cart.push({
    store: currentItem.store,
    item:  currentItem.item,
    spec:  specText.trim(),
    note:  document.getElementById('itemNote').value.trim(),
    qty:   Number(document.getElementById('qtySelect').value || 1),
    price: currentItem.price
  });
  closeModal('optionModal');
  updateCartBadge();
  showToast('已加入購物車 🎉');
}

// ── 購物車 ────────────────────────────────────────────────────────
function updateCartBadge() {
  const total = cart.reduce((a, c) => a + c.qty, 0);
  document.getElementById('cartBadge').innerText = total;
}

function openCart() {
  renderCart();
  openModal('cartModal');
}

function renderCart() {
  const listEl = document.getElementById('cartList');
  const totalEl = document.getElementById('cartTotal');
  if (cart.length === 0) {
    listEl.innerHTML = '<div class="empty" style="margin:20px 0">購物車是空的</div>';
    totalEl.innerText = '';
    return;
  }
  let html = '';
  cart.forEach((c, i) => {
    const sub = [c.spec, c.note].filter(Boolean).join('｜');
    html += '<div class="cart-item">'
      + '<div class="cart-info">'
      + '<div class="cart-name">' + c.item + ' x' + c.qty + '</div>'
      + (sub ? '<div class="cart-sub">' + sub + '</div>' : '')
      + '</div>'
      + '<div class="cart-price">$' + (c.price * c.qty) + '</div>'
      + '<button class="icon-btn" onclick="removeCartItem('+i+')" title="刪除">🗑</button>'
      + '</div>';
  });
  listEl.innerHTML = html;
  const total = cart.reduce((a, c) => a + c.price * c.qty, 0);
  totalEl.innerText = '合計：$' + total;
}

function removeCartItem(i) {
  cart.splice(i, 1);
  updateCartBadge();
  renderCart();
}

async function submitCart() {
  if (cart.length === 0) { alert('購物車是空的'); return; }
  document.getElementById('submitCartBtn').disabled = true;
  document.getElementById('submitCartBtn').innerText = '送出中...';
  let allOk = true;
  for (const c of cart) {
    const res = await fetch('/api/order', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ ...c, name: profile.displayName, userId: profile.userId })
    });
    const result = await res.json();
    if (!result.success) {
      if (result.reason === 'duplicate') {
        showToast('⚠️ ' + c.item + ' 已送出過，略過');
      } else {
        allOk = false;
      }
    }
  }
  cart = [];
  updateCartBadge();
  closeModal('cartModal');
  document.getElementById('submitCartBtn').disabled = false;
  document.getElementById('submitCartBtn').innerText = '送出訂單';
  if (allOk) { showToast('訂單已送出 ✅'); loadMyOrders(); }
  else { alert('部分品項送出失敗，請重試'); }
}

// ── 我的訂單 ─────────────────────────────────────────────────────
async function loadMyOrders() {
  if (!profile) return;
  try {
    const res = await fetch('/api/my-orders?userId=' + encodeURIComponent(profile.userId));
    const orders = await res.json();
    renderMyOrders(orders);
  } catch(e) {}
}

function renderMyOrders(orders) {
  const listEl = document.getElementById('myOrderList');
  const totalEl = document.getElementById('myOrderTotal');
  if (!orders || orders.length === 0) {
    listEl.innerHTML = '<div class="empty" style="margin:20px 0">今天還沒有訂單</div>';
    totalEl.innerText = '';
    return;
  }
  let html = '';
  let total = 0;
  orders.forEach(o => {
    total += o.total;
    const sub = [o.spec, o.note].filter(Boolean).join('｜');
    const statusTag = o.status === '已付款'
      ? '<span class="tag paid">已付款</span>'
      : '<span class="tag unpaid">未付款</span>';
    html += '<div class="cart-item">'
      + '<div class="cart-info">'
      + '<div class="cart-name">' + o.item + ' x' + o.qty + ' ' + statusTag + '</div>'
      + (sub ? '<div class="cart-sub">' + sub + '</div>' : '')
      + '<div class="cart-sub">' + o.store + '</div>'
      + '</div>'
      + '<div style="display:flex;flex-direction:column;align-items:flex-end;gap:4px">'
      + '<div class="cart-price">$' + o.total + '</div>'
      + (o.status !== '已付款'
          ? '<button class="icon-btn" onclick="editNote('+o.rowIndex+',\''+escHtml(o.note)+'\')" title="備註">✏️</button>'
          + '<button class="icon-btn" onclick="delOrder('+o.rowIndex+')" title="刪除">🗑</button>'
          : '')
      + '</div>'
      + '</div>';
  });
  listEl.innerHTML = html;
  totalEl.innerText = '今日合計：$' + total;
}

function escHtml(s) { return (s||'').replace(/'/g,"\\'"); }

async function delOrder(rowIndex) {
  if (!confirm('確定要刪除這筆訂單？')) return;
  const res = await fetch('/api/order/' + rowIndex, {
    method: 'DELETE',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ userId: profile.userId })
  });
  const result = await res.json();
  if (result.success) { showToast('已刪除'); loadMyOrders(); }
  else { alert('刪除失敗'); }
}

function editNote(rowIndex, currentNote) {
  editingRowIndex = rowIndex;
  document.getElementById('editNoteInput').value = currentNote || '';
  openModal('editNoteModal');
}

async function saveNote() {
  const note = document.getElementById('editNoteInput').value.trim();
  const res = await fetch('/api/order/' + editingRowIndex + '/note', {
    method: 'PATCH',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ userId: profile.userId, note })
  });
  const result = await res.json();
  if (result.success) { showToast('備註已更新'); closeModal('editNoteModal'); loadMyOrders(); }
  else { alert('更新失敗'); }
}

// ── Modal 控制 ────────────────────────────────────────────────────
function openModal(id) {
  const m = document.getElementById(id);
  m.style.display = 'flex';
  // 若是我的訂單，先刷新
  if (id === 'myOrderModal') loadMyOrders();
}
function closeModal(id) { document.getElementById(id).style.display = 'none'; }

// ── Toast ─────────────────────────────────────────────────────────
function showToast(msg) {
  let t = document.getElementById('toast');
  if (!t) {
    t = document.createElement('div');
    t.id = 'toast';
    t.style.cssText = 'position:fixed;bottom:90px;left:50%;transform:translateX(-50%);background:#333;color:#fff;padding:10px 20px;border-radius:999px;font-size:14px;z-index:9999;opacity:0;transition:opacity .3s';
    document.body.appendChild(t);
  }
  t.innerText = msg;
  t.style.opacity = '1';
  setTimeout(() => { t.style.opacity = '0'; }, 2000);
}

// ── 我的訂單入口（加在購物車旁）─────────────────────────────────
function addMyOrderBtn() {
  const fab = document.getElementById('cartFab');
  const btn = document.createElement('button');
  btn.className = 'cart-fab';
  btn.style.right = '140px';
  btn.innerText = '📋 訂單';
  btn.onclick = () => openModal('myOrderModal');
  document.body.appendChild(btn);
}

// ── Init ──────────────────────────────────────────────────────────
renderMenu();
initLIFF().then(() => { addMyOrderBtn(); });
setInterval(checkOrderStatus, 30000);
<\/script>
</body>
</html>`;

    res.send(html);
  } catch (err) {
    console.error('載入訂餐頁失敗：', err.message);
    res.send('載入訂餐頁失敗，請稍後再試');
  }
});

// ─── 管理後台 ─────────────────────────────────────────────────────
app.get('/admin', async (req, res) => {
  // 簡單 token 驗證（在 .env 設 ADMIN_TOKEN）
  const token = req.query.token || '';
  if (token !== process.env.ADMIN_TOKEN) {
    return res.status(401).send('Unauthorized. 請附上 ?token=YOUR_ADMIN_TOKEN');
  }

  const html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>訂餐管理後台</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>
    *{box-sizing:border-box}
    body{margin:0;font-family:Arial,"Microsoft JhengHei",sans-serif;background:#f5f5f5;color:#333}
    .header{background:#333;color:#fff;padding:16px 20px;position:sticky;top:0;z-index:10}
    .header h1{margin:0;font-size:20px}
    .header p{margin:4px 0 0;font-size:13px;opacity:.7}
    .container{padding:16px;max-width:900px;margin:0 auto}
    .card{background:#fff;border-radius:12px;padding:16px;margin-bottom:16px;box-shadow:0 2px 8px rgba(0,0,0,.06)}
    .btn{padding:10px 18px;border:none;border-radius:999px;background:#333;color:#fff;font-size:14px;cursor:pointer;margin:4px}
    .btn.green{background:#06c755}
    .btn.red{background:#e53935}
    .btn.orange{background:#f57c00}
    .btn:disabled{background:#aaa}
    .status-bar{display:flex;align-items:center;gap:12px;flex-wrap:wrap}
    .dot{width:12px;height:12px;border-radius:50%;background:#aaa;display:inline-block}
    .dot.open{background:#06c755}
    table{width:100%;border-collapse:collapse;font-size:13px}
    th,td{padding:10px 8px;border-bottom:1px solid #eee;text-align:left}
    th{background:#f9f9f9;font-weight:bold}
    .tag{display:inline-block;padding:2px 8px;border-radius:999px;font-size:12px}
    .tag.paid{background:#e8f5e9;color:#2e7d32}
    .tag.unpaid{background:#fff3e0;color:#e65100}
    .tag.del{background:#fce4ec;color:#c62828}
    .summary{display:flex;gap:16px;flex-wrap:wrap}
    .summary-box{background:#f9f9f9;border-radius:10px;padding:12px 18px;flex:1;min-width:130px}
    .summary-box .num{font-size:28px;font-weight:bold;color:#333}
    .summary-box .label{font-size:13px;color:#777;margin-top:2px}
    input[type=number]{width:70px;padding:6px;border-radius:8px;border:1px solid #ddd;font-size:14px}
    select.pay-select{padding:6px;border-radius:8px;border:1px solid #ddd;font-size:13px}
    .search-box{width:100%;padding:10px;border-radius:10px;border:1px solid #ddd;font-size:14px;margin-bottom:12px}
    @media(max-width:600px){th,td{padding:7px 5px;font-size:12px}}
  </style>
</head>
<body>
  <div class="header">
    <h1>🍱 訂餐管理後台</h1>
    <p id="adminTime"></p>
  </div>
  <div class="container">
    <!-- 開單控制 -->
    <div class="card">
      <h3 style="margin-top:0">📢 開單控制</h3>
      <div class="status-bar">
        <span class="dot" id="statusDot"></span>
        <span id="statusText" style="font-size:15px;font-weight:bold"></span>
        <button class="btn green" id="openBtn" onclick="openOrder()">開單</button>
        <button class="btn red" id="closeBtn" onclick="closeOrder()">結單</button>
        <button class="btn orange" onclick="clearOrders()">清空</button>
      </div>
      <div style="margin-top:14px;display:flex;align-items:center;gap:8px;flex-wrap:wrap">
        <label style="font-size:14px">自動結單：</label>
        <input type="number" id="autoMin" value="30" min="1" max="480">
        <span style="font-size:14px">分鐘後</span>
        <button class="btn" onclick="setAutoClose()">設定</button>
        <button class="btn" onclick="cancelAutoClose()">取消自動結單</button>
        <span id="autoCloseInfo" style="font-size:13px;color:#e53935;font-weight:bold"></span>
      </div>
    </div>

    <!-- 統計 -->
    <div class="card" id="summaryCard">
      <h3 style="margin-top:0">📊 今日統計</h3>
      <div class="summary" id="summaryBox">載入中...</div>
    </div>

    <!-- 訂單列表 -->
    <div class="card">
      <h3 style="margin-top:0">📋 今日訂單</h3>
      <input class="search-box" id="searchBox" placeholder="搜尋姓名、品項..." oninput="filterTable()">
      <div style="margin-bottom:10px;display:flex;gap:8px;flex-wrap:wrap">
        <button class="btn green" onclick="batchPaid()">✅ 全部標記已付款</button>
        <button class="btn" onclick="loadOrders()">🔄 重新整理</button>
        <button class="btn" onclick="copyShopOrder()">📋 複製店家單</button>
      </div>
      <div style="overflow-x:auto">
        <table id="orderTable">
          <thead>
            <tr>
              <th>姓名</th><th>品項</th><th>規格</th><th>備註</th>
              <th>數量</th><th>金額</th><th>狀態</th><th>操作</th>
            </tr>
          </thead>
          <tbody id="orderBody">
            <tr><td colspan="8" style="text-align:center;padding:20px;color:#999">載入中...</td></tr>
          </tbody>
        </table>
      </div>
    </div>
  </div>

<script>
const TOKEN = '${process.env.ADMIN_TOKEN}';
let ordersCache = [];
let autoCloseAt = null;
let countdownTimer = null;

document.getElementById('adminTime').innerText = new Date().toLocaleString('zh-TW');

async function api(path, method='GET', body=null) {
  const opts = { method, headers: { 'Content-Type':'application/json','x-admin-token': TOKEN } };
  if (body) opts.body = JSON.stringify(body);
  const res = await fetch(path, opts);
  return res.json();
}

async function loadStatus() {
  const data = await api('/api/status');
  const dot = document.getElementById('statusDot');
  const txt = document.getElementById('statusText');
  dot.className = 'dot' + (data.isOpen ? ' open' : '');
  txt.innerText = data.isOpen ? '開單中' : '已結單';
  if (data.autoCloseAt) {
    autoCloseAt = data.autoCloseAt;
    startCountdown();
  } else {
    document.getElementById('autoCloseInfo').innerText = '';
  }
}

function startCountdown() {
  if (countdownTimer) clearInterval(countdownTimer);
  countdownTimer = setInterval(() => {
    if (!autoCloseAt) { clearInterval(countdownTimer); return; }
    const diff = new Date(autoCloseAt) - new Date();
    if (diff <= 0) {
      document.getElementById('autoCloseInfo').innerText = '已自動結單';
      clearInterval(countdownTimer);
      loadStatus();
      return;
    }
    const m = Math.floor(diff/60000), s = Math.floor((diff%60000)/1000);
    document.getElementById('autoCloseInfo').innerText = '⏰ ' + m + '分' + s + '秒後自動結單';
  }, 1000);
}

async function openOrder() {
  await api('/api/admin/open','POST');
  loadStatus(); loadOrders();
}
async function closeOrder() {
  await api('/api/admin/close','POST');
  loadStatus(); loadOrders();
}
async function clearOrders() {
  if (!confirm('確定清空今日狀態？這只會重設開單狀態，不會刪除 Sheet 資料')) return;
  await api('/api/admin/clear','POST');
  loadStatus(); loadOrders();
}
async function setAutoClose() {
  const min = Number(document.getElementById('autoMin').value || 30);
  await api('/api/admin/auto-close','POST',{ minutes: min });
  loadStatus();
}
async function cancelAutoClose() {
  await api('/api/admin/cancel-auto-close','POST');
  autoCloseAt = null;
  document.getElementById('autoCloseInfo').innerText = '';
  clearInterval(countdownTimer);
}

async function loadOrders() {
  const data = await api('/api/admin/orders');
  ordersCache = data;
  renderTable(data);
  renderSummary(data);
}

function renderSummary(orders) {
  const active = orders.filter(o => o.status !== '已刪除');
  const paid = active.filter(o => o.status === '已付款');
  const unpaid = active.filter(o => o.status === '未付款');
  const totalMoney = active.reduce((a,o) => a + o.total, 0);
  const paidMoney  = paid.reduce((a,o) => a + o.total, 0);
  document.getElementById('summaryBox').innerHTML =
    box(active.length,'筆訂單') + box(totalMoney,'總金額 $') +
    box(paid.length,'已付款') + box(unpaid.length,'未付款') +
    box(paidMoney,'已收金額 $');

  function box(n, label) {
    return '<div class="summary-box"><div class="num">'+n+'</div><div class="label">'+label+'</div></div>';
  }
}

function renderTable(orders) {
  const body = document.getElementById('orderBody');
  if (!orders || orders.length === 0) {
    body.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:20px;color:#999">今日無訂單</td></tr>';
    return;
  }
  body.innerHTML = orders.map(o => {
    const tagCls = o.status === '已付款' ? 'paid' : o.status === '已刪除' ? 'del' : 'unpaid';
    return '<tr data-name="'+esc(o.name)+'" data-item="'+esc(o.item)+'">'
      + '<td>'+esc(o.name)+'</td>'
      + '<td>'+esc(o.item)+'</td>'
      + '<td style="color:#777">'+esc(o.spec)+'</td>'
      + '<td style="color:#777">'+esc(o.note)+'</td>'
      + '<td>'+o.qty+'</td>'
      + '<td>$'+o.total+'</td>'
      + '<td><span class="tag '+tagCls+'">'+o.status+'</span></td>'
      + '<td>'
      + (o.status === '未付款'
          ? '<select class="pay-select" id="pt-'+o.rowIndex+'"><option>現金</option><option>Line Pay</option><option>轉帳</option></select>'
          + '<button class="btn green" style="padding:6px 10px;font-size:12px" onclick="markPaid('+o.rowIndex+')">付款</button>'
          + '<button class="btn red" style="padding:6px 10px;font-size:12px" onclick="delOrderAdmin('+o.rowIndex+')">刪</button>'
          : '')
      + '</td>'
      + '</tr>';
  }).join('');
}

function filterTable() {
  const q = document.getElementById('searchBox').value.toLowerCase();
  const rows = document.querySelectorAll('#orderBody tr[data-name]');
  rows.forEach(r => {
    const name = (r.dataset.name || '').toLowerCase();
    const item = (r.dataset.item || '').toLowerCase();
    r.style.display = (!q || name.includes(q) || item.includes(q)) ? '' : 'none';
  });
}

async function markPaid(rowIndex) {
  const sel = document.getElementById('pt-'+rowIndex);
  const payType = sel ? sel.value : '現金';
  await api('/api/admin/paid','POST',{ rowIndex, payType });
  loadOrders();
}

async function delOrderAdmin(rowIndex) {
  if (!confirm('確定刪除？')) return;
  await api('/api/admin/delete-order','POST',{ rowIndex });
  loadOrders();
}

async function batchPaid() {
  if (!confirm('將所有未付款訂單標記為已付款？')) return;
  await api('/api/admin/batch-paid','POST',{ payType:'現金' });
  loadOrders();
}

function copyShopOrder() {
  const active = ordersCache.filter(o => o.status !== '已刪除');
  const counts = {};
  active.forEach(o => {
    const key = o.item + (o.spec ? '（'+o.spec+'）' : '');
    counts[key] = (counts[key]||0) + o.qty;
  });
  let text = '您好，今天訂購如下：\n\n';
  let total = 0;
  for (const k in counts) { text += k + ' x' + counts[k] + '\n'; total += counts[k]; }
  const money = active.reduce((a,o)=>a+o.total,0);
  text += '\n總數：' + total + '份\n總金額：' + money + '元\n\n麻煩您，謝謝～';
  navigator.clipboard.writeText(text).then(()=>alert('已複製到剪貼板'));
}

function esc(s) { return (s||'').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

// init
loadStatus();
loadOrders();
setInterval(loadStatus, 30000);
setInterval(loadOrders, 60000);
<\/script>
</body>
</html>`;

  res.send(html);
});

// ─── API 路由 ─────────────────────────────────────────────────────

// 中介：管理員 token 驗證
function adminAuth(req, res, next) {
  const token = req.headers['x-admin-token'] || req.query.token;
  if (token !== process.env.ADMIN_TOKEN) return res.status(401).json({ error: 'Unauthorized' });
  next();
}

// 狀態
app.get('/api/status', (req, res) => {
  res.json({ isOpen, autoCloseAt });
});

// 使用者下單
app.post('/api/order', async (req, res) => {
  const result = await saveOrderToSheet(req.body);
  res.json(result);
});

// 我的訂單
app.get('/api/my-orders', async (req, res) => {
  const { userId } = req.query;
  if (!userId) return res.json([]);
  const orders = await getOrdersByUser(userId);
  res.json(orders);
});

// 刪除訂單（使用者）
app.delete('/api/order/:rowIndex', async (req, res) => {
  const rowIndex = Number(req.params.rowIndex);
  const { userId } = req.body;
  const ok = await deleteOrder(userId, rowIndex);
  res.json({ success: ok });
});

// 修改備註
app.patch('/api/order/:rowIndex/note', async (req, res) => {
  const rowIndex = Number(req.params.rowIndex);
  const { userId, note } = req.body;
  const ok = await updateOrderNote(userId, rowIndex, note);
  res.json({ success: ok });
});

// ── 管理員 API ────────────────────────────────────────────────────
app.post('/api/admin/open', adminAuth, (req, res) => {
  if (!isOpen) { isOpen = true; allText = ''; }
  res.json({ isOpen });
});

app.post('/api/admin/close', adminAuth, (req, res) => {
  isOpen = false;
  cancelAutoClose();
  res.json({ isOpen });
});

app.post('/api/admin/clear', adminAuth, (req, res) => {
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

app.post('/api/admin/cancel-auto-close', adminAuth, (req, res) => {
  cancelAutoClose();
  res.json({ ok: true });
});

app.get('/api/admin/orders', adminAuth, async (req, res) => {
  const orders = await getAllOrdersToday();
  res.json(orders);
});

app.post('/api/admin/paid', adminAuth, async (req, res) => {
  const { rowIndex, payType } = req.body;
  const ok = await markPaid(rowIndex, payType);
  res.json({ success: ok });
});

app.post('/api/admin/batch-paid', adminAuth, async (req, res) => {
  // 取得所有今日未付款者
  const orders = await getAllOrdersToday();
  const users = [...new Set(orders.filter(o => o.status === '未付款').map(o => o.userId))];
  let total = 0;
  for (const uid of users) {
    total += await markPaidByUser(uid, req.body.payType || '現金');
  }
  res.json({ success: true, count: total });
});

app.post('/api/admin/delete-order', adminAuth, async (req, res) => {
  const { rowIndex } = req.body;
  try {
    await authSheet();
    const sheet = doc.sheetsByTitle['Orders'];
    const rows = await sheet.getRows();
    const target = rows.find(r => r.rowIndex === rowIndex);
    if (!target) return res.json({ success: false });
    target['狀態'] = '已刪除';
    await target.save();
    res.json({ success: true });
  } catch (e) {
    res.json({ success: false, reason: e.message });
  }
});

// ─── Webhook（原有保留 + 新增指令）─────────────────────────────
app.post('/webhook', async (req, res) => {
  try {
    const event = req.body.events?.[0];
    if (!event || event.type !== 'message') return res.sendStatus(200);

    let profileName = '未知使用者';
    try {
      if (event.source.type === 'group') {
        const p = await client.getGroupMemberProfile(event.source.groupId, event.source.userId);
        profileName = p.displayName;
      } else {
        const p = await client.getProfile(event.source.userId);
        profileName = p.displayName;
      }
    } catch (err) { console.error('取得使用者名稱失敗：', err.message); }

    knownUsers[profileName] = event.source.userId;
    await saveUserToSheet(profileName, event.source.userId, event.source.type, event.source.groupId || '');

    const userId = event.source.userId;
    const text = event.message?.text ? event.message.text.trim() : '';
    const replyToken = event.replyToken;

    const reply = (t) => client.replyMessage(replyToken, { type: 'text', text: t });

    // ── 指令處理 ──
    if (text === '開單') {
      if (!isAdmin(userId)) return reply('只有管理員可以開單') && res.sendStatus(200);
      if (isOpen) return reply('目前已開單中') && res.sendStatus(200);
      isOpen = true; allText = '';
      await reply('已開單，可以開始點餐 🍱\n點餐頁：' + process.env.LIFF_URL);
      return res.sendStatus(200);
    }

    if (text.startsWith('開單')) {
      // 例：「開單 30」→ 30 分鐘自動結單
      const min = Number(text.replace('開單', '').trim());
      if (!isAdmin(userId)) return reply('只有管理員可以開單') && res.sendStatus(200);
      isOpen = true; allText = '';
      if (min > 0) scheduleAutoClose(min);
      await reply('已開單！' + (min > 0 ? `將於 ${min} 分鐘後自動結單` : '') + '\n點餐頁：' + (process.env.LIFF_URL || ''));
      return res.sendStatus(200);
    }

    if (text === '清空') {
      if (!isAdmin(userId)) return reply('只有管理員可以清空') && res.sendStatus(200);
      allText = ''; isOpen = false; cancelAutoClose();
      await reply('已清空訂單');
      return res.sendStatus(200);
    }

    if (text === '店家單') {
      if (!isAdmin(userId)) return reply('只有管理員可以查看店家單') && res.sendStatus(200);
      const result = parseOrders(allText);
      await reply(formatShopOrder(result.itemCount, result.userTotal));
      return res.sendStatus(200);
    }

    if (text === '結單' || text === '收單' || text === '統計') {
      if (!isAdmin(userId)) return reply('只有管理員可以結單 / 統計') && res.sendStatus(200);
      const result = parseOrders(allText);
      isOpen = false; cancelAutoClose();
      await reply(formatResult(result.itemCount, result.userTotal));
      return res.sendStatus(200);
    }

    if (text === '後台') {
      if (!isAdmin(userId)) return reply('只有管理員可以查看後台') && res.sendStatus(200);
      const url = (process.env.APP_URL || '') + '/admin?token=' + (process.env.ADMIN_TOKEN || '');
      await reply('管理後台：' + url);
      return res.sendStatus(200);
    }

    if (isOpen && event.message.type === 'text') {
      allText += '\n' + text;
    }

    return res.sendStatus(200);
  } catch (error) {
    console.error('Webhook error:', error);
    return res.sendStatus(200);
  }
});

app.use((err, req, res, next) => {
  console.error('Global error:', err);
  res.sendStatus(200);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => { console.log('Server running on ' + PORT); });
