require('dotenv').config();

const express = require('express');
const line    = require('@line/bot-sdk');
const { GoogleSpreadsheet } = require('google-spreadsheet');

const app = express();

const lineConfig = {
  channelAccessToken: process.env.CHANNEL_ACCESS_TOKEN,
  channelSecret:      process.env.CHANNEL_SECRET
};

// ★ webhook 必須在 express.json() 之前，且只限路徑掛載
app.use('/webhook', line.middleware(lineConfig));
app.use('/api',    express.json());
app.use('/admin',  express.json());

const client = new line.Client(lineConfig);
const doc    = new GoogleSpreadsheet(process.env.SHEET_ID);

// ════════════════════════════════════════════════════════════════
//  全域狀態
// ════════════════════════════════════════════════════════════════
let isOpen        = false;
let autoCloseTimer = null;
let autoCloseAt   = null;   // ISO string，供前端倒數用
const knownUsers  = {};

const admins = [
  'U8d9c82446aa9eb90d7de001cfc7ea90f',
  'Ubcfae64b443b9fad21bbc584e991b306',
  'U5c44a04efc62664bd45ec80d77be7d93',
  'Uc669eca67bf477460945f45751edd3e9'
];

function isAdmin(uid) { return admins.includes(uid); }

// ════════════════════════════════════════════════════════════════
//  工具
// ════════════════════════════════════════════════════════════════
function nowTW() {
  return new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
}
function todayTW() {
  return new Date().toLocaleDateString('zh-TW', { timeZone: 'Asia/Taipei' });
}
function dateTW(d) {
  return new Date(d).toLocaleDateString('zh-TW', { timeZone: 'Asia/Taipei' });
}

// ════════════════════════════════════════════════════════════════
//  Google Sheets 認證
// ════════════════════════════════════════════════════════════════
async function authSheet() {
  await doc.useServiceAccountAuth({
    client_email: process.env.GOOGLE_CLIENT_EMAIL,
    private_key:  process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n')
  });
  await doc.loadInfo();
}

// ════════════════════════════════════════════════════════════════
//  菜單讀取（支援圖片欄位）
// ════════════════════════════════════════════════════════════════
async function loadMenu() {
  await authSheet();
  const sheet = doc.sheetsByTitle['Menu'];
  if (!sheet) return [];
  const rows = await sheet.getRows();
  return rows
    .map(r => ({
      store: String(r['店家']   || '').trim(),
      item:  String(r['品項']   || '').trim(),
      price: Number(r['價格']   || 0),
      image: String(r['圖片URL'] || '').trim()   // 新增欄位，沒填就空字串
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

// ════════════════════════════════════════════════════════════════
//  Users
// ════════════════════════════════════════════════════════════════
async function saveUserToSheet(name, userId, sourceType, groupId) {
  try {
    await authSheet();
    const sheet = doc.sheetsByTitle['Users'];
    if (!sheet) return;
    await sheet.addRow({
      時間:     nowTW(),
      LINE名稱: name,
      userId,
      來源類型: sourceType,
      群組ID:   groupId || '',
      權限:     isAdmin(userId) ? 'admin' : 'user'
    });
  } catch (e) { console.error('寫入 Users 失敗：', e.message); }
}

// ════════════════════════════════════════════════════════════════
//  Orders CRUD
// ════════════════════════════════════════════════════════════════
async function saveOrderToSheet(order) {
  try {
    await authSheet();
    const sheet = doc.sheetsByTitle['Orders'];
    if (!sheet) return { success: false, reason: 'no_sheet' };

    const rows  = await sheet.getRows();
    const today = todayTW();

    // 防重複
    const dup = rows.find(r =>
      String(r['userId'] || '')     === String(order.userId) &&
      String(r['品項']   || '').trim() === String(order.item || '').trim() &&
      String(r['規格']   || '').trim() === String(order.spec || '').trim() &&
      String(r['時間']   || '').startsWith(today) &&
      String(r['狀態']   || '') !== '已刪除'
    );
    if (dup) return { success: false, reason: 'duplicate' };

    const qty   = Number(order.qty   || 1);
    const price = Number(order.price || 0);
    await sheet.addRow({
      時間:     nowTW(),
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
  } catch (e) {
    console.error('寫入 Orders 失敗：', e.message);
    return { success: false, reason: e.message };
  }
}

async function getOrdersByUser(userId, dateStr) {
  try {
    await authSheet();
    const sheet = doc.sheetsByTitle['Orders'];
    if (!sheet) return [];
    const rows = await sheet.getRows();
    const target = dateStr || todayTW();
    return rows
      .filter(r =>
        String(r['userId'] || '') === String(userId) &&
        String(r['時間']   || '').startsWith(target) &&
        String(r['狀態']   || '') !== '已刪除'
      )
      .map(r => ({
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
  } catch (e) {
    console.error('讀取訂單失敗：', e.message);
    return [];
  }
}

// 歷史訂單日期列表（該 userId 有訂單的日期）
async function getOrderDates(userId) {
  try {
    await authSheet();
    const sheet = doc.sheetsByTitle['Orders'];
    if (!sheet) return [];
    const rows = await sheet.getRows();
    const dates = new Set();
    rows.forEach(r => {
      if (String(r['userId'] || '') === userId && String(r['狀態'] || '') !== '已刪除') {
        const d = String(r['時間'] || '').split(' ')[0];
        if (d) dates.add(d);
      }
    });
    return [...dates].sort().reverse();
  } catch (e) { return []; }
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
  } catch (e) { return false; }
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
  } catch (e) { return false; }
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
  } catch (e) { return []; }
}

// 後台歷史查詢（指定日期）
async function getAllOrdersByDate(dateStr) {
  try {
    await authSheet();
    const sheet = doc.sheetsByTitle['Orders'];
    if (!sheet) return [];
    const rows = await sheet.getRows();
    return rows
      .filter(r => String(r['時間'] || '').startsWith(dateStr))
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
  } catch (e) { return []; }
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
    target['付款時間'] = nowTW();
    target['付款方式'] = payType || '現金';
    await target.save();
    return true;
  } catch (e) { return false; }
}

async function batchMarkPaid(payType) {
  try {
    await authSheet();
    const sheet = doc.sheetsByTitle['Orders'];
    if (!sheet) return 0;
    const rows  = await sheet.getRows();
    const today = todayTW();
    let count = 0;
    for (const r of rows) {
      if (
        String(r['時間']  || '').startsWith(today) &&
        String(r['狀態']  || '') === '未付款'
      ) {
        r['狀態']    = '已付款';
        r['付款時間'] = nowTW();
        r['付款方式'] = payType || '現金';
        await r.save();
        count++;
      }
    }
    return count;
  } catch (e) { return 0; }
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
  } catch (e) { return false; }
}

// ════════════════════════════════════════════════════════════════
//  自動結單
// ════════════════════════════════════════════════════════════════
function scheduleAutoClose(minutes) {
  if (autoCloseTimer) clearTimeout(autoCloseTimer);
  const ms = Math.max(1, Number(minutes)) * 60 * 1000;
  autoCloseAt   = new Date(Date.now() + ms).toISOString();
  autoCloseTimer = setTimeout(() => {
    isOpen        = false;
    autoCloseAt   = null;
    autoCloseTimer = null;
    console.log('[自動結單]', nowTW());
  }, ms);
}

function cancelAutoClose() {
  if (autoCloseTimer) clearTimeout(autoCloseTimer);
  autoCloseTimer = null;
  autoCloseAt    = null;
}

// ════════════════════════════════════════════════════════════════
//  LINE 報表文字
// ════════════════════════════════════════════════════════════════
async function buildStatReport() {
  const orders = await getAllOrdersToday();
  const active = orders.filter(o => o.status !== '已刪除');
  if (!active.length) return '📊 今日尚無訂單';

  const itemCount = {}, userTotal = {};
  const unpaidSet = new Set();

  for (const o of active) {
    const k = o.item + (o.spec ? '（' + o.spec + '）' : '');
    itemCount[k] = (itemCount[k] || 0) + o.qty;
    const n = o.name || o.userId || '未知';
    userTotal[n] = (userTotal[n] || 0) + o.total;
    if (o.status === '未付款') unpaidSet.add(n);
  }

  const grand = Object.values(userTotal).reduce((a, b) => a + b, 0);
  let t = '📊 今日訂餐統計\n─────────────\n【品項數量】\n';
  for (const k in itemCount) t += k + ' x' + itemCount[k] + '\n';
  t += '\n【個人金額】\n';
  for (const n in userTotal) t += n + '：$' + userTotal[n] + '\n';
  t += '\n💰 總金額：$' + grand;
  t += unpaidSet.size ? '\n\n⚠️ 未付款：' + [...unpaidSet].join('、') : '\n\n✅ 所有人已付款';
  return t;
}

async function buildShopOrder() {
  const orders = await getAllOrdersToday();
  const active = orders.filter(o => o.status !== '已刪除');
  if (!active.length) return '今日尚無訂單';

  const itemCount = {};
  for (const o of active) {
    const k = o.item + (o.spec ? '（' + o.spec + '）' : '');
    itemCount[k] = (itemCount[k] || 0) + o.qty;
  }

  let out = '您好，今天訂購如下：\n\n', total = 0;
  for (const k in itemCount) { out += k + ' x' + itemCount[k] + '\n'; total += itemCount[k]; }
  const money = active.reduce((a, o) => a + o.total, 0);
  return out + '\n總數：' + total + '份\n總金額：' + money + '元\n\n麻煩您，謝謝～';
}

// ════════════════════════════════════════════════════════════════
//  Admin middleware
// ════════════════════════════════════════════════════════════════
function adminAuth(req, res, next) {
  const t = req.headers['x-admin-token'] || req.query.token || '';
  if (!process.env.ADMIN_TOKEN || t !== process.env.ADMIN_TOKEN)
    return res.status(401).json({ error: 'Unauthorized' });
  next();
}

// ════════════════════════════════════════════════════════════════
//  靜態路由
// ════════════════════════════════════════════════════════════════
app.get('/', (_q, r) => r.send('LINE 訂餐機器人運作中'));

// ── 菜單 API ─────────────────────────────────────────────────────
app.get('/api/menu', async (_q, res) => {
  try {
    const [menu, optionData] = await Promise.all([loadMenu(), loadOptions()]);
    res.json({ menu, optionData });
  } catch (e) {
    res.status(500).json({ menu: [], optionData: {} });
  }
});

// ── 狀態 API ─────────────────────────────────────────────────────
app.get('/api/status', (_q, res) => res.json({ isOpen, autoCloseAt }));

// ── 歷史日期 ──────────────────────────────────────────────────────
app.get('/api/my-dates', async (req, res) => {
  const { userId } = req.query;
  if (!userId) return res.json([]);
  res.json(await getOrderDates(userId));
});

// ── 使用者訂單 ────────────────────────────────────────────────────
app.post('/api/order', async (req, res) => {
  res.json(await saveOrderToSheet(req.body));
});

app.get('/api/my-orders', async (req, res) => {
  const { userId, date } = req.query;
  if (!userId) return res.json([]);
  res.json(await getOrdersByUser(userId, date || null));
});

app.delete('/api/order/:rowIndex', async (req, res) => {
  const { userId } = req.body;
  if (!userId) return res.json({ success: false });
  res.json({ success: await deleteOrder(userId, Number(req.params.rowIndex)) });
});

app.patch('/api/order/:rowIndex/note', async (req, res) => {
  const { userId, note } = req.body;
  if (!userId) return res.json({ success: false });
  res.json({ success: await updateOrderNote(userId, Number(req.params.rowIndex), note) });
});

// ── 管理員 API ────────────────────────────────────────────────────
app.post('/api/admin/open', adminAuth, (_q, res) => {
  isOpen = true;
  res.json({ isOpen });
});
app.post('/api/admin/close', adminAuth, (_q, res) => {
  isOpen = false; cancelAutoClose();
  res.json({ isOpen });
});
app.post('/api/admin/clear', adminAuth, (_q, res) => {
  isOpen = false; cancelAutoClose();
  res.json({ ok: true });
});
app.post('/api/admin/auto-close', adminAuth, (req, res) => {
  if (!isOpen) isOpen = true;
  scheduleAutoClose(Number(req.body.minutes) || 30);
  res.json({ autoCloseAt });
});
app.post('/api/admin/cancel-auto-close', adminAuth, (_q, res) => {
  cancelAutoClose();
  res.json({ ok: true });
});
app.get('/api/admin/orders', adminAuth, async (req, res) => {
  const { date } = req.query;
  if (date) res.json(await getAllOrdersByDate(date));
  else      res.json(await getAllOrdersToday());
});
app.post('/api/admin/paid', adminAuth, async (req, res) => {
  res.json({ success: await markPaid(req.body.rowIndex, req.body.payType) });
});
app.post('/api/admin/batch-paid', adminAuth, async (req, res) => {
  res.json({ success: true, count: await batchMarkPaid(req.body.payType || '現金') });
});
app.post('/api/admin/delete-order', adminAuth, async (req, res) => {
  res.json({ success: await adminDeleteOrder(req.body.rowIndex) });
});

// ════════════════════════════════════════════════════════════════
//  前台 LIFF 頁面
// ════════════════════════════════════════════════════════════════
app.get('/order', (_q, res) => {
  const LIFF_ID = String(process.env.LIFF_ID || '2010025093-yATK02dc').replace(/[^a-zA-Z0-9\-]/g, '');
  res.setHeader('Content-Type', 'text/html; charset=utf-8');
  res.end(buildOrderPage(LIFF_ID));
});

function buildOrderPage(liffId) {
  const lines = [];
  const p = s => lines.push(s);

  p('<!DOCTYPE html><html lang="zh-TW"><head>');
  p('<meta charset="utf-8">');
  p('<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1">');
  p('<title>訂餐小幫手</title>');
  p('<script src="https://static.line-scdn.net/liff/edge/2/sdk.js"></script>');
  p('<style>');
  p(':root{--green:#06c755;--green-dark:#05a847;--red:#e53935;--gray:#f6f3ee;--card:#fff;--text:#222;--sub:#888;--border:#eee;--radius:16px;--shadow:0 2px 12px rgba(0,0,0,.07)}');
  p('*{box-sizing:border-box;-webkit-tap-highlight-color:transparent}');
  p('body{margin:0;font-family:-apple-system,Arial,"Microsoft JhengHei",sans-serif;background:var(--gray);color:var(--text);min-height:100vh}');
  // header
  p('.hd{background:#fff;border-bottom:1px solid var(--border);position:sticky;top:0;z-index:50;padding:12px 16px}');
  p('.hd-row{display:flex;align-items:center;justify-content:space-between}');
  p('.hd h1{margin:0;font-size:20px;font-weight:800;color:var(--text)}');
  p('.hd-logo{font-size:22px}');
  p('.hd-status{font-size:12px;padding:3px 10px;border-radius:999px;font-weight:600}');
  p('.hd-status.open{background:#e8f5e9;color:#2e7d32}');
  p('.hd-status.closed{background:#fce4ec;color:#b71c1c}');
  p('.hd-user{font-size:13px;color:var(--sub);margin-top:4px}');
  p('.hd-cd{font-size:12px;color:var(--red);font-weight:600;margin-top:2px}');
  // search bar
  p('.search-wrap{padding:10px 16px 4px;background:#fff;border-bottom:1px solid var(--border)}');
  p('.search-wrap input{width:100%;padding:9px 14px;border-radius:999px;border:1.5px solid var(--border);font-size:14px;outline:none;background:var(--gray)}');
  p('.search-wrap input:focus{border-color:var(--green)}');
  // tabs
  p('.tabs{background:#fff;padding:8px 12px 10px;overflow-x:auto;white-space:nowrap;display:flex;gap:6px;border-bottom:1px solid var(--border);position:sticky;top:106px;z-index:40}');
  p('.tab{display:inline-flex;align-items:center;gap:4px;padding:6px 14px;border-radius:999px;border:1.5px solid var(--border);font-size:13px;font-weight:600;cursor:pointer;background:#fff;color:var(--sub);transition:.15s;white-space:nowrap;flex-shrink:0}');
  p('.tab.fav{border-color:#ffc107;color:#f57f17}');
  p('.tab.active,.tab:active{background:var(--green);color:#fff;border-color:var(--green)}');
  p('.tab-fav-icon{font-size:14px}');
  // container
  p('.container{padding:12px 12px 140px}');
  p('.store-hd{font-size:16px;font-weight:800;margin:18px 0 8px 2px;display:flex;align-items:center;gap:8px}');
  p('.store-hd .fav-btn{font-size:18px;cursor:pointer;background:none;border:none;padding:0;line-height:1}');
  // card
  p('.card{background:var(--card);border-radius:var(--radius);margin-bottom:10px;box-shadow:var(--shadow);overflow:hidden;display:flex;flex-direction:column}');
  p('.card-img{width:100%;height:140px;object-fit:cover;display:block}');
  p('.card-img-placeholder{height:0}');
  p('.card-body{padding:12px 14px 14px}');
  p('.card-store{font-size:11px;color:var(--sub);margin-bottom:3px;font-weight:500}');
  p('.card-name{font-size:17px;font-weight:800;margin-bottom:5px;line-height:1.3}');
  p('.card-price{font-size:16px;color:var(--green);font-weight:700;margin-bottom:10px}');
  p('.add-btn{width:100%;padding:11px;border:none;border-radius:999px;background:var(--green);color:#fff;font-size:15px;font-weight:700;cursor:pointer;transition:.15s}');
  p('.add-btn:active{background:var(--green-dark)}');
  p('.add-btn:disabled{background:#ccc;cursor:default}');
  // empty
  p('.empty{text-align:center;padding:48px 20px;color:var(--sub);font-size:15px}');
  p('#loadingMsg{text-align:center;padding:48px 20px;color:var(--sub);font-size:15px}');
  // FAB
  p('.fab-wrap{position:fixed;bottom:24px;right:16px;display:flex;flex-direction:column;gap:10px;z-index:9000}');
  p('.fab{display:flex;align-items:center;gap:6px;padding:13px 18px;border:none;border-radius:999px;font-size:14px;font-weight:700;cursor:pointer;box-shadow:0 4px 16px rgba(0,0,0,.18);transition:.15s;white-space:nowrap}');
  p('.fab.cart{background:var(--green);color:#fff}');
  p('.fab.orders{background:#fff;color:var(--text);border:1.5px solid var(--border)}');
  p('.fab.fab-hidden{display:none!important}');
  p('.badge{background:var(--red);color:#fff;border-radius:50%;font-size:11px;min-width:18px;height:18px;display:inline-flex;align-items:center;justify-content:center;padding:0 3px}');
  // modal backdrop
  p('.modal{display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:9999;align-items:flex-end;justify-content:center}');
  p('.modal.center{align-items:center}');
  p('.modal.open{display:flex}');
  p('.modal-box{background:#fff;width:100%;max-width:500px;max-height:88vh;overflow-y:auto;border-radius:20px 20px 0 0;padding:20px 18px 32px;position:relative}');
  p('.modal.center .modal-box{border-radius:20px;margin:0 14px;max-height:80vh}');
  p('.modal-title{font-size:18px;font-weight:800;margin:0 0 14px}');
  p('.modal-close{position:absolute;top:16px;right:16px;background:var(--gray);border:none;border-radius:50%;width:30px;height:30px;font-size:16px;cursor:pointer;display:flex;align-items:center;justify-content:center}');
  // option
  p('.opt-group-title{font-weight:700;font-size:14px;margin:14px 0 4px;color:var(--sub)}');
  p('.opt-chips{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:4px}');
  p('.opt-chip{padding:7px 14px;border-radius:999px;border:1.5px solid var(--border);font-size:14px;cursor:pointer;transition:.15s;background:#fff}');
  p('.opt-chip.selected{background:var(--green);color:#fff;border-color:var(--green)}');
  p('.opt-chip.disabled{opacity:.35;cursor:default}');
  p('.qty-row{display:flex;align-items:center;gap:12px;margin-top:16px}');
  p('.qty-row label{font-weight:700;font-size:14px;flex:1}');
  p('.qty-ctrl{display:flex;align-items:center;gap:0;border:1.5px solid var(--border);border-radius:999px;overflow:hidden}');
  p('.qty-ctrl button{background:none;border:none;width:36px;height:36px;font-size:20px;cursor:pointer;display:flex;align-items:center;justify-content:center;color:var(--green);font-weight:700}');
  p('.qty-ctrl span{min-width:28px;text-align:center;font-size:16px;font-weight:700}');
  p('.note-input{width:100%;padding:10px 14px;border-radius:12px;border:1.5px solid var(--border);font-size:14px;margin-top:10px;outline:none}');
  p('.note-input:focus{border-color:var(--green)}');
  // buttons
  p('.btn-primary{width:100%;padding:13px;border:none;border-radius:999px;background:var(--green);color:#fff;font-size:16px;font-weight:700;cursor:pointer;margin-top:14px;transition:.15s}');
  p('.btn-primary:active{background:var(--green-dark)}');
  p('.btn-primary:disabled{background:#ccc}');
  p('.btn-secondary{width:100%;padding:12px;border:1.5px solid var(--border);border-radius:999px;background:#fff;color:var(--text);font-size:15px;font-weight:600;cursor:pointer;margin-top:8px}');
  p('.btn-danger{background:var(--red);color:#fff;border:none;border-radius:999px;padding:7px 14px;font-size:13px;cursor:pointer;font-weight:600}');
  // cart / order list
  p('.order-row{display:flex;align-items:flex-start;gap:10px;padding:12px 0;border-bottom:1px solid var(--border)}');
  p('.order-info{flex:1;min-width:0}');
  p('.order-name{font-weight:700;font-size:15px;margin-bottom:2px}');
  p('.order-sub{font-size:13px;color:var(--sub);margin-bottom:2px}');
  p('.order-price{font-size:15px;font-weight:700;color:var(--green);white-space:nowrap}');
  p('.order-actions{display:flex;flex-direction:column;align-items:flex-end;gap:6px}');
  p('.icon-btn{background:none;border:none;font-size:18px;cursor:pointer;padding:4px}');
  p('.tag{display:inline-block;padding:2px 9px;border-radius:999px;font-size:12px;font-weight:600}');
  p('.tag-paid{background:#e8f5e9;color:#2e7d32}');
  p('.tag-unpaid{background:#fff3e0;color:#e65100}');
  p('.total-row{font-size:16px;font-weight:800;margin:12px 0 4px;color:var(--text)}');
  // reorder btn
  p('.reorder-btn{background:var(--gray);border:none;border-radius:999px;padding:5px 12px;font-size:12px;cursor:pointer;font-weight:600;color:var(--green);margin-top:4px}');
  // history tabs
  p('.date-chips{display:flex;gap:8px;overflow-x:auto;padding-bottom:6px;margin-bottom:10px}');
  p('.date-chip{padding:6px 14px;border-radius:999px;border:1.5px solid var(--border);font-size:13px;cursor:pointer;white-space:nowrap;background:#fff}');
  p('.date-chip.active{background:var(--green);color:#fff;border-color:var(--green)}');
  // toast
  p('#toast{position:fixed;bottom:110px;left:50%;transform:translateX(-50%);background:#222;color:#fff;padding:10px 22px;border-radius:999px;font-size:14px;z-index:99999;opacity:0;transition:opacity .25s;pointer-events:none;white-space:nowrap}');
  p('</style>');
  p('</head><body>');

  // ── Header ──────────────────────────────────────────────────────
  p('<div class="hd">');
  p('  <div class="hd-row">');
  p('    <div style="display:flex;align-items:center;gap:8px"><span class="hd-logo">🍱</span><h1>訂餐小幫手</h1></div>');
  p('    <span class="hd-status closed" id="hd-status">未開放</span>');
  p('  </div>');
  p('  <div class="hd-user" id="hd-user">正在取得 LINE 使用者資料...</div>');
  p('  <div class="hd-cd" id="hd-cd"></div>');
  p('</div>');

  // ── Search ──────────────────────────────────────────────────────
  p('<div class="search-wrap">');
  p('  <input id="searchInput" type="search" placeholder="🔍 搜尋品項或店家..." oninput="onSearch()">');
  p('</div>');

  // ── Tabs ────────────────────────────────────────────────────────
  p('<div class="tabs" id="storeTabs"></div>');

  // ── Menu ────────────────────────────────────────────────────────
  p('<div class="container" id="menuBox"><div id="loadingMsg">菜單載入中...</div></div>');

  // ── FABs ────────────────────────────────────────────────────────
  p('<div class="fab-wrap">');
  p('  <button class="fab orders fab-hidden" id="orderFab" onclick="openModal(\'myOrderModal\')">📋 我的訂單</button>');
  p('  <button class="fab cart fab-hidden" id="cartFab" onclick="openModal(\'cartModal\')">🛒 購物車<span class="badge" id="cartBadge">0</span></button>');
  p('</div>');

  // ── Toast ──────────────────────────────────────────────────────
  p('<div id="toast"></div>');

  // ── Modal: 商品選項 ─────────────────────────────────────────────
  p('<div class="modal" id="optionModal">');
  p('  <div class="modal-box">');
  p('    <button class="modal-close" onclick="closeModal(\'optionModal\')">✕</button>');
  p('    <div class="modal-title" id="optModalTitle"></div>');
  p('    <div id="optModalOptions"></div>');
  p('    <div class="qty-row">');
  p('      <label>數量</label>');
  p('      <div class="qty-ctrl">');
  p('        <button onclick="changeQty(-1)">−</button>');
  p('        <span id="qtyVal">1</span>');
  p('        <button onclick="changeQty(1)">＋</button>');
  p('      </div>');
  p('    </div>');
  p('    <input class="note-input" id="itemNote" placeholder="備註（例：不加辣、少冰）">');
  p('    <button class="btn-primary" onclick="submitOptions()">加入購物車 🛒</button>');
  p('    <button class="btn-secondary" onclick="closeModal(\'optionModal\')">取消</button>');
  p('  </div>');
  p('</div>');

  // ── Modal: 購物車 ──────────────────────────────────────────────
  p('<div class="modal" id="cartModal">');
  p('  <div class="modal-box">');
  p('    <button class="modal-close" onclick="closeModal(\'cartModal\')">✕</button>');
  p('    <div class="modal-title">🛒 我的購物車</div>');
  p('    <div id="cartList"></div>');
  p('    <div class="total-row" id="cartTotalRow"></div>');
  p('    <button class="btn-primary" id="submitCartBtn" onclick="submitCart()">送出訂單</button>');
  p('    <button class="btn-secondary" onclick="closeModal(\'cartModal\')">繼續點餐</button>');
  p('  </div>');
  p('</div>');

  // ── Modal: 我的訂單 ────────────────────────────────────────────
  p('<div class="modal" id="myOrderModal">');
  p('  <div class="modal-box">');
  p('    <button class="modal-close" onclick="closeModal(\'myOrderModal\')">✕</button>');
  p('    <div class="modal-title">📋 我的訂單</div>');
  p('    <div class="date-chips" id="dateChips"></div>');
  p('    <div id="myOrderList"></div>');
  p('    <div class="total-row" id="myOrderTotal"></div>');
  p('    <button class="btn-secondary" onclick="closeModal(\'myOrderModal\')">關閉</button>');
  p('  </div>');
  p('</div>');

  // ── Modal: 備註編輯 ────────────────────────────────────────────
  p('<div class="modal center" id="editNoteModal">');
  p('  <div class="modal-box">');
  p('    <button class="modal-close" onclick="closeModal(\'editNoteModal\')">✕</button>');
  p('    <div class="modal-title">✏️ 修改備註</div>');
  p('    <input class="note-input" id="editNoteInput" placeholder="備註內容">');
  p('    <button class="btn-primary" onclick="saveNote()">儲存</button>');
  p('    <button class="btn-secondary" onclick="closeModal(\'editNoteModal\')">取消</button>');
  p('  </div>');
  p('</div>');

  // ════════════════════════════════════════════════════════════════
  //  JavaScript
  // ════════════════════════════════════════════════════════════════
  p('<script>');
  p('var LIFF_ID="' + liffId + '";');
  p('var menu=[],optionData={},favStores=[];');
  p('var profile=null,liffReady=false;');
  p('var currentItem=null,currentGroups=[],currentQty=1;');
  p('var cart=[];');
  p('var editingRowIndex=null;');
  p('var orderDates=[],currentOrderDate=null;');
  p('var searchQ="";');

  // ── loadMenuData ─────────────────────────────────────────────
  p('async function loadMenuData(){');
  p('  try{');
  p('    var r=await fetch("/api/menu");');
  p('    var d=await r.json();');
  p('    menu=d.menu||[];');
  p('    optionData=d.optionData||{};');
  p('    renderMenu();');
  p('  }catch(e){');
  p('    var el=document.getElementById("loadingMsg");');
  p('    if(el)el.innerText="菜單載入失敗，請重新整理";');
  p('  }');
  p('}');

  // ── initLIFF ─────────────────────────────────────────────────
  p('async function initLIFF(){');
  p('  try{');
  p('    await liff.init({liffId:LIFF_ID});');
  p('    if(!liff.isLoggedIn()){liff.login();return;}');
  p('    profile=await liff.getProfile();');
  p('    liffReady=true;');
  p('    document.getElementById("hd-user").innerText="👤 "+profile.displayName;');
  p('    loadFav();');
  p('    enableButtons();');
  p('    document.getElementById("cartFab").classList.remove("fab-hidden");');
  p('    document.getElementById("orderFab").classList.remove("fab-hidden");');
  p('    checkStatus();');
  p('    loadMyOrders(null);');
  p('  }catch(e){');
  p('    document.getElementById("hd-user").innerText="LIFF 初始化失敗："+e.message;');
  p('  }');
  p('}');

  // ── checkStatus ──────────────────────────────────────────────
  p('async function checkStatus(){');
  p('  try{');
  p('    var r=await fetch("/api/status");');
  p('    var d=await r.json();');
  p('    var el=document.getElementById("hd-status");');
  p('    if(d.isOpen){el.innerText="開放點餐";el.className="hd-status open";}');
  p('    else{el.innerText="未開放";el.className="hd-status closed";}');
  p('    if(d.autoCloseAt)startCd(d.autoCloseAt);');
  p('    else document.getElementById("hd-cd").innerText="";');
  p('  }catch(e){}');
  p('}');
  p('function startCd(iso){');
  p('  var el=document.getElementById("hd-cd");');
  p('  function tick(){');
  p('    var diff=new Date(iso)-new Date();');
  p('    if(diff<=0){el.innerText="";return;}');
  p('    el.innerText="⏰ 自動結單："+Math.floor(diff/60000)+"分"+Math.floor((diff%60000)/1000)+"秒後";');
  p('    setTimeout(tick,1000);');
  p('  }tick();');
  p('}');

  // ── 收藏 ─────────────────────────────────────────────────────
  p('function loadFav(){');
  p('  try{var s=localStorage.getItem("favStores_"+(profile?profile.userId:""));');
  p('  favStores=s?JSON.parse(s):[];}catch(e){favStores=[];}');
  p('}');
  p('function saveFav(){');
  p('  try{localStorage.setItem("favStores_"+(profile?profile.userId:""),JSON.stringify(favStores));}catch(e){}');
  p('}');
  p('function toggleFav(store){');
  p('  var i=favStores.indexOf(store);');
  p('  if(i>=0)favStores.splice(i,1); else favStores.push(store);');
  p('  saveFav();renderMenu();');
  p('}');

  // ── renderMenu ───────────────────────────────────────────────
  p('function renderMenu(){');
  p('  var lm=document.getElementById("loadingMsg");');
  p('  if(lm)lm.style.display="none";');
  p('  var box=document.getElementById("menuBox");');
  p('  var tabs=document.getElementById("storeTabs");');
  p('  if(!menu.length){box.innerHTML=\'<div class="empty">目前沒有菜單資料</div>\';tabs.innerHTML="";return;}');
  // filtered by search
  p('  var filtered=menu.filter(function(m){');
  p('    if(!searchQ)return true;');
  p('    return m.item.indexOf(searchQ)>=0||m.store.indexOf(searchQ)>=0;');
  p('  });');
  // stores order: fav first
  p('  var allStores=[];');
  p('  menu.forEach(function(m){if(allStores.indexOf(m.store)<0)allStores.push(m.store);});');
  p('  var favInMenu=favStores.filter(function(s){return allStores.indexOf(s)>=0;});');
  p('  var rest=allStores.filter(function(s){return favStores.indexOf(s)<0;});');
  p('  var stores=favInMenu.concat(rest);');
  // tabs
  p('  tabs.innerHTML=stores.map(function(s,i){');
  p('    var isFav=favStores.indexOf(s)>=0;');
  p('    return \'<div class="tab\'+(isFav?" fav":"")+\'" id="tab\'+i+\'" onclick="scrollToStore(\'+i+\')">\'');
  p('      +(isFav?\'<span>⭐</span>\':\'\')+escH(s)+\'</div>\';');
  p('  }).join("");');
  // menu cards
  p('  var html="",cur="";');
  p('  stores.forEach(function(store,si){');
  p('    var items=filtered.filter(function(m){return m.store===store;});');
  p('    if(!items.length)return;');
  p('    var isFav=favStores.indexOf(store)>=0;');
  p('    html+=\'<div id="store\'+si+\'" class="store-hd">\'');
  p('      +\'<span>\'+escH(store)+\'</span>\'');
  p('      +\'<button class="fav-btn" onclick="toggleFav(\'+(\'"\'+store+\'"\').replace(/&amp;/g,"&")+\')">\'+(isFav?"⭐":"☆")+\'</button>\'');
  p('      +\'</div>\';');
  p('    items.forEach(function(m){');
  p('      var idx=menu.indexOf(m);');
  p('      html+=\'<div class="card">\';');
  p('      if(m.image){html+=\'<img class="card-img" src="\'+escH(m.image)+\'" alt="\'+escH(m.item)+\'" onerror="this.style.display=\'none\'">\'; }');
  p('      html+=\'<div class="card-body">\';');
  p('      html+=\'<div class="card-store">\'+escH(m.store)+\'</div>\';');
  p('      html+=\'<div class="card-name">\'+escH(m.item)+\'</div>\';');
  p('      html+=\'<div class="card-price">$\'+m.price+\'</div>\';');
  p('      html+=\'<button class="add-btn" id="btn\'+idx+\'" onclick="addToCart(\'+idx+\')" disabled>載入中...</button>\';');
  p('      html+=\'</div></div>\';');
  p('    });');
  p('  });');
  p('  box.innerHTML=html||\'<div class="empty">找不到相符的品項</div>\';');
  p('  if(liffReady)enableButtons();');
  p('}');

  p('function scrollToStore(i){');
  p('  document.querySelectorAll(".tab").forEach(function(b,j){b.classList.toggle("active",j===i);});');
  p('  var el=document.getElementById("store"+i);');
  p('  if(el)el.scrollIntoView({behavior:"smooth",block:"start"});');
  p('}');

  p('function onSearch(){searchQ=document.getElementById("searchInput").value.trim();renderMenu();}');

  p('function enableButtons(){');
  p('  menu.forEach(function(_,i){');
  p('    var b=document.getElementById("btn"+i);');
  p('    if(b){b.disabled=false;b.innerText="加入購物車";}');
  p('  });');
  p('}');

  // ── addToCart ────────────────────────────────────────────────
  p('function addToCart(idx){');
  p('  if(!liffReady||!profile){alert("尚未取得 LINE 使用者資料");return;}');
  p('  currentItem=menu[idx];');
  p('  var key=currentItem.store+"||"+currentItem.item;');
  p('  currentGroups=optionData[key]||[];');
  p('  currentQty=1;');
  p('  document.getElementById("optModalTitle").innerText=currentItem.item;');
  p('  document.getElementById("qtyVal").innerText="1";');
  p('  document.getElementById("itemNote").value="";');
  p('  var ob=document.getElementById("optModalOptions");');
  p('  ob.innerHTML="";');
  p('  if(currentGroups.length===0){');
  p('    ob.innerHTML=\'<p style="color:var(--sub);margin:8px 0">此商品無需選擇規格，可直接加入購物車。</p>\';');
  p('  }else{');
  p('    currentGroups.forEach(function(g,gi){');
  p('      var d=document.createElement("div");');
  p('      d.innerHTML=\'<div class="opt-group-title">\'+escH(g.category)+\' (\'+g.min+\'~\'+g.max+\'選)</div><div class="opt-chips" id="chips\'+gi+\'"></div>\';');
  p('      ob.appendChild(d);');
  p('      g.options.forEach(function(opt){');
  p('        var c=document.createElement("div");');
  p('        c.className="opt-chip";c.innerText=opt;');
  p('        c.onclick=function(){toggleChip(gi,g.max,c);};');
  p('        document.getElementById("chips"+gi).appendChild(c);');
  p('      });');
  p('    });');
  p('  }');
  p('  openModal("optionModal");');
  p('}');

  p('function changeQty(d){');
  p('  currentQty=Math.max(1,Math.min(20,currentQty+d));');
  p('  document.getElementById("qtyVal").innerText=currentQty;');
  p('}');

  p('function toggleChip(gi,max,chip){');
  p('  var chips=[].slice.call(document.querySelectorAll("#chips"+gi+" .opt-chip"));');
  p('  var selected=chips.filter(function(c){return c.classList.contains("selected");});');
  p('  if(chip.classList.contains("selected")){chip.classList.remove("selected");}');
  p('  else if(selected.length<max){chip.classList.add("selected");}');
  p('  // disable unselected if max reached');
  p('  var sel=chips.filter(function(c){return c.classList.contains("selected");});');
  p('  chips.forEach(function(c){');
  p('    if(sel.length>=max&&!c.classList.contains("selected"))c.classList.add("disabled");');
  p('    else c.classList.remove("disabled");');
  p('  });');
  p('}');

  p('function submitOptions(){');
  p('  var specParts=[];');
  p('  for(var i=0;i<currentGroups.length;i++){');
  p('    var g=currentGroups[i];');
  p('    var sel=[].slice.call(document.querySelectorAll("#chips"+i+" .opt-chip.selected"));');
  p('    if(sel.length<g.min||sel.length>g.max){alert(g.category+" 需要選 "+g.min+"~"+g.max+" 個");return;}');
  p('    if(sel.length)specParts.push(g.category+"："+sel.map(function(c){return c.innerText;}).join("、"));');
  p('  }');
  p('  cart.push({store:currentItem.store,item:currentItem.item,spec:specParts.join(" "),note:document.getElementById("itemNote").value.trim(),qty:currentQty,price:currentItem.price});');
  p('  closeModal("optionModal");');
  p('  updateBadge();');
  p('  showToast("已加入購物車 🎉");');
  p('}');

  // ── 購物車 ──────────────────────────────────────────────────
  p('function updateBadge(){');
  p('  document.getElementById("cartBadge").innerText=cart.reduce(function(a,c){return a+c.qty;},0);');
  p('}');

  p('function renderCart(){');
  p('  var el=document.getElementById("cartList");');
  p('  var te=document.getElementById("cartTotalRow");');
  p('  if(!cart.length){el.innerHTML=\'<div class="empty" style="padding:24px 0">購物車是空的</div>\';te.innerText="";return;}');
  p('  el.innerHTML=cart.map(function(c,i){');
  p('    var sub=[c.spec,c.note].filter(Boolean).join("｜");');
  p('    return \'<div class="order-row">\'');
  p('      +\'<div class="order-info">\'');
  p('      +\'<div class="order-name">\'+escH(c.item)+\' x\'+c.qty+\'</div>\'');
  p('      +(sub?\'<div class="order-sub">\'+escH(sub)+\'</div>\':"")');
  p('      +\'</div>\'');
  p('      +\'<div class="order-actions">\'');
  p('      +\'<div class="order-price">$\'+(c.price*c.qty)+\'</div>\'');
  p('      +\'<button class="btn-danger" onclick="removeCartItem(\'+i+\')">移除</button>\'');
  p('      +\'</div></div>\';');
  p('  }).join("");');
  p('  te.innerText="合計：$"+cart.reduce(function(a,c){return a+c.price*c.qty;},0);');
  p('}');

  p('function removeCartItem(i){cart.splice(i,1);updateBadge();renderCart();}');

  p('async function submitCart(){');
  p('  if(!cart.length){alert("購物車是空的");return;}');
  p('  var btn=document.getElementById("submitCartBtn");');
  p('  btn.disabled=true;btn.innerText="送出中...";');
  p('  for(var i=0;i<cart.length;i++){');
  p('    var c=cart[i];');
  p('    try{');
  p('      var r=await fetch("/api/order",{method:"POST",headers:{"Content-Type":"application/json"},');
  p('        body:JSON.stringify({store:c.store,item:c.item,spec:c.spec,note:c.note,qty:c.qty,price:c.price,name:profile.displayName,userId:profile.userId})});');
  p('      var d=await r.json();');
  p('      if(!d.success&&d.reason==="duplicate")showToast("⚠️ "+c.item+" 已送出過，略過");');
  p('    }catch(e){}');
  p('  }');
  p('  cart=[];updateBadge();closeModal("cartModal");');
  p('  btn.disabled=false;btn.innerText="送出訂單";');
  p('  showToast("訂單已送出 ✅");');
  p('  loadMyOrders(null);');
  p('}');

  // ── 我的訂單 ────────────────────────────────────────────────
  p('async function loadMyOrders(dateStr){');
  p('  if(!profile)return;');
  p('  currentOrderDate=dateStr;');
  p('  var url="/api/my-orders?userId="+encodeURIComponent(profile.userId);');
  p('  if(dateStr)url+="&date="+encodeURIComponent(dateStr);');
  p('  try{var r=await fetch(url);renderMyOrders(await r.json());}catch(e){}');
  p('}');

  p('async function loadOrderDates(){');
  p('  if(!profile)return;');
  p('  try{');
  p('    var r=await fetch("/api/my-dates?userId="+encodeURIComponent(profile.userId));');
  p('    orderDates=await r.json();');
  p('    renderDateChips();');
  p('  }catch(e){}');
  p('}');

  p('function renderDateChips(){');
  p('  var el=document.getElementById("dateChips");');
  p('  if(!orderDates.length){el.innerHTML="";return;}');
  p('  el.innerHTML=orderDates.map(function(d){');
  p('    return \'<div class="date-chip\'+(d===currentOrderDate?" active":"")+\'" onclick="loadMyOrders(\'+"\'"+\')\">'+'\'+d+\'</div>\';');
  // note: date chips onclick built carefully
  p('  }).join("");');
  p('}');

  p('function renderMyOrders(orders){');
  p('  var el=document.getElementById("myOrderList");');
  p('  var te=document.getElementById("myOrderTotal");');
  p('  if(!orders||!orders.length){el.innerHTML=\'<div class="empty" style="padding:24px 0">這天沒有訂單</div>\';te.innerText="";return;}');
  p('  el.innerHTML=orders.map(function(o){');
  p('    var sub=[o.spec,o.note].filter(Boolean).join("｜");');
  p('    var tag=o.status==="已付款"?\'<span class="tag tag-paid">已付款</span>\':\'<span class="tag tag-unpaid">未付款</span>\';');
  p('    var canEdit=o.status!=="已付款";');
  p('    return \'<div class="order-row">\'');
  p('      +\'<div class="order-info">\'');
  p('      +\'<div class="order-name">\'+escH(o.item)+\' x\'+o.qty+\' \'+tag+\'</div>\'');
  p('      +(sub?\'<div class="order-sub">\'+escH(sub)+\'</div>\':"")');
  p('      +\'<div class="order-sub">\'+escH(o.store)+\'</div>\'');
  p('      +\'<button class="reorder-btn" onclick="reorder(\'+JSON.stringify(o)+\')">🔄 再訂一次</button>\'');
  p('      +\'</div>\'');
  p('      +\'<div class="order-actions">\'');
  p('      +\'<div class="order-price">$\'+o.total+\'</div>\'');
  p('      +(canEdit?\'<button class="icon-btn" onclick="editNote(\'+o.rowIndex+\',\'+JSON.stringify(o.note||"")+\')">✏️</button>\':""  )');
  p('      +(canEdit?\'<button class="icon-btn" onclick="delOrder(\'+o.rowIndex+\')">🗑</button>\':"")');
  p('      +\'</div></div>\';');
  p('  }).join("");');
  p('  te.innerText="合計：$"+orders.reduce(function(a,o){return a+o.total;},0);');
  p('}');

  // ── 再訂一次 ────────────────────────────────────────────────
  p('function reorder(o){');
  p('  // 找到對應 menu 索引');
  p('  var idx=menu.findIndex(function(m){return m.store===o.store&&m.item===o.item;});');
  p('  if(idx<0){showToast("此品項已不在菜單中");return;}');
  p('  cart.push({store:o.store,item:o.item,spec:o.spec||"",note:o.note||"",qty:o.qty||1,price:menu[idx].price});');
  p('  updateBadge();');
  p('  showToast("已加入購物車 🔄");');
  p('}');

  // ── 刪除 / 備註 ─────────────────────────────────────────────
  p('async function delOrder(rowIndex){');
  p('  if(!confirm("確定要刪除這筆訂單？"))return;');
  p('  var r=await fetch("/api/order/"+rowIndex,{method:"DELETE",headers:{"Content-Type":"application/json"},body:JSON.stringify({userId:profile.userId})});');
  p('  var d=await r.json();');
  p('  if(d.success){showToast("已刪除");loadMyOrders(currentOrderDate);}else alert("刪除失敗");');
  p('}');
  p('function editNote(rowIndex,cur){editingRowIndex=rowIndex;document.getElementById("editNoteInput").value=cur||"";openModal("editNoteModal");}');
  p('async function saveNote(){');
  p('  var note=document.getElementById("editNoteInput").value.trim();');
  p('  var r=await fetch("/api/order/"+editingRowIndex+"/note",{method:"PATCH",headers:{"Content-Type":"application/json"},body:JSON.stringify({userId:profile.userId,note:note})});');
  p('  var d=await r.json();');
  p('  if(d.success){showToast("備註已更新");closeModal("editNoteModal");loadMyOrders(currentOrderDate);}else alert("更新失敗");');
  p('}');

  // ── Modal 控制 ───────────────────────────────────────────────
  p('function openModal(id){');
  p('  document.getElementById(id).classList.add("open");');
  p('  if(id==="cartModal")renderCart();');
  p('  if(id==="myOrderModal"){loadOrderDates();loadMyOrders(currentOrderDate);}');
  p('}');
  p('function closeModal(id){document.getElementById(id).classList.remove("open");}');
  // close on backdrop click
  p('document.querySelectorAll(".modal").forEach(function(m){');
  p('  m.addEventListener("click",function(e){if(e.target===m)m.classList.remove("open");});');
  p('});');

  // ── Toast ────────────────────────────────────────────────────
  p('function showToast(msg){');
  p('  var t=document.getElementById("toast");');
  p('  t.innerText=msg;t.style.opacity="1";');
  p('  clearTimeout(t._timer);');
  p('  t._timer=setTimeout(function(){t.style.opacity="0";},2500);');
  p('}');

  // ── escH ─────────────────────────────────────────────────────
  p('function escH(s){return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");}');

  // ── renderDateChips 修正版 ───────────────────────────────────
  // 重新用正確方式渲染 date chips（避免閉包問題）
  p('function renderDateChips(){');
  p('  var el=document.getElementById("dateChips");');
  p('  if(!orderDates.length){el.innerHTML="";return;}');
  p('  el.innerHTML="";');
  p('  orderDates.forEach(function(d){');
  p('    var div=document.createElement("div");');
  p('    div.className="date-chip"+(d===currentOrderDate?" active":"");');
  p('    div.innerText=d;');
  p('    div.onclick=(function(dd){return function(){loadMyOrders(dd);};})(d);');
  p('    el.appendChild(div);');
  p('  });');
  p('}');

  // ── Init ─────────────────────────────────────────────────────
  p('loadMenuData();');
  p('initLIFF();');
  p('setInterval(checkStatus,30000);');
  p('</script>');
  p('</body></html>');

  return lines.join('\n');
}

// ════════════════════════════════════════════════════════════════
//  管理後台
// ════════════════════════════════════════════════════════════════
app.get('/admin', (req, res) => {
  const token = req.query.token || '';
  if (!process.env.ADMIN_TOKEN || token !== process.env.ADMIN_TOKEN)
    return res.status(401).send('Unauthorized. 請附上 ?token=YOUR_ADMIN_TOKEN');
  res.setHeader('Content-Type', 'text/html; charset=utf-8');
  res.end(buildAdminPage());
});

function buildAdminPage() {
  const lines = [];
  const p = s => lines.push(s);

  p('<!DOCTYPE html><html lang="zh-TW"><head>');
  p('<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">');
  p('<title>訂餐管理後台</title>');
  p('<style>');
  p(':root{--green:#06c755;--green-d:#05a847;--red:#e53935;--orange:#f57c00;--bg:#f4f4f8;--card:#fff;--border:#e8e8ef;--text:#1a1a2e;--sub:#888;--r:12px;--sh:0 2px 12px rgba(0,0,0,.07)}');
  p('*{box-sizing:border-box}');
  p('body{margin:0;font-family:-apple-system,Arial,"Microsoft JhengHei",sans-serif;background:var(--bg);color:var(--text)}');
  p('.hd{background:var(--text);color:#fff;padding:14px 20px;position:sticky;top:0;z-index:50;display:flex;align-items:center;justify-content:space-between}');
  p('.hd h1{margin:0;font-size:18px;font-weight:800}');
  p('.hd-time{font-size:12px;opacity:.6}');
  p('.wrap{padding:16px;max-width:1000px;margin:0 auto}');
  p('.card{background:var(--card);border-radius:var(--r);padding:18px;margin-bottom:16px;box-shadow:var(--sh)}');
  p('.card h3{margin:0 0 14px;font-size:16px;font-weight:800;display:flex;align-items:center;gap:6px}');
  // stat boxes
  p('.stats{display:grid;grid-template-columns:repeat(auto-fill,minmax(140px,1fr));gap:10px;margin-bottom:0}');
  p('.stat-box{background:var(--bg);border-radius:10px;padding:14px 16px}');
  p('.stat-num{font-size:26px;font-weight:800;margin-bottom:2px}');
  p('.stat-lbl{font-size:12px;color:var(--sub)}');
  p('.stat-box.green .stat-num{color:var(--green)}');
  p('.stat-box.red .stat-num{color:var(--red)}');
  p('.stat-box.orange .stat-num{color:var(--orange)}');
  // control
  p('.ctrl-row{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:12px}');
  p('.status-dot{width:11px;height:11px;border-radius:50%;background:var(--sub);flex-shrink:0}');
  p('.status-dot.on{background:var(--green)}');
  p('.status-lbl{font-size:15px;font-weight:700}');
  p('.ac-info{font-size:13px;color:var(--red);font-weight:600}');
  // buttons
  p('.btn{padding:9px 16px;border:none;border-radius:999px;font-size:13px;font-weight:700;cursor:pointer;transition:.15s;white-space:nowrap}');
  p('.btn.green{background:var(--green);color:#fff}');
  p('.btn.red{background:var(--red);color:#fff}');
  p('.btn.orange{background:var(--orange);color:#fff}');
  p('.btn.gray{background:#ccc;color:#fff}');
  p('.btn.outline{background:var(--card);border:1.5px solid var(--border);color:var(--text)}');
  p('.btn.sm{padding:6px 12px;font-size:12px}');
  p('.btn:hover{opacity:.88}');
  p('.btn:disabled{background:#ccc;cursor:default}');
  p('input[type=number],input[type=date],select.sel{padding:8px 12px;border-radius:8px;border:1.5px solid var(--border);font-size:13px;background:var(--card)}');
  p('input[type=number]{width:68px}');
  // search
  p('.search{width:100%;padding:9px 14px;border-radius:999px;border:1.5px solid var(--border);font-size:14px;margin-bottom:12px;background:var(--card)}');
  // table
  p('.tbl-wrap{overflow-x:auto}');
  p('table{width:100%;border-collapse:collapse;font-size:13px}');
  p('th{padding:10px 8px;border-bottom:2px solid var(--border);text-align:left;font-weight:700;color:var(--sub);white-space:nowrap;background:var(--bg)}');
  p('td{padding:9px 8px;border-bottom:1px solid var(--border);vertical-align:middle}');
  p('tr:hover td{background:#f9f9fc}');
  p('.store-badge{display:inline-block;padding:2px 8px;border-radius:6px;font-size:11px;font-weight:600;background:#e3f2fd;color:#1565c0;margin-right:4px}');
  p('.tag{display:inline-block;padding:3px 10px;border-radius:999px;font-size:12px;font-weight:600}');
  p('.tag-paid{background:#e8f5e9;color:#2e7d32}');
  p('.tag-unpaid{background:#fff3e0;color:#e65100}');
  p('.tag-del{background:#fce4ec;color:#b71c1c}');
  // group section
  p('.group-sec{margin-bottom:20px}');
  p('.group-title{font-size:15px;font-weight:800;padding:8px 10px;background:var(--bg);border-radius:8px;margin-bottom:6px;display:flex;align-items:center;justify-content:space-between}');
  p('.group-sum{font-size:12px;color:var(--sub)}');
  // date filter
  p('.date-row{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:14px}');
  p('@media(max-width:600px){th,td{padding:7px 5px;font-size:12px}.stats{grid-template-columns:repeat(2,1fr)}}');
  p('</style></head><body>');

  p('<div class="hd">');
  p('  <h1>🍱 訂餐管理後台</h1>');
  p('  <span class="hd-time" id="hdTime"></span>');
  p('</div>');
  p('<div class="wrap">');

  // ── 開單控制 ────────────────────────────────────────────────
  p('<div class="card">');
  p('<h3>📢 開單控制</h3>');
  p('<div class="ctrl-row">');
  p('  <span class="status-dot" id="sDot"></span>');
  p('  <span class="status-lbl" id="sTxt"></span>');
  p('  <button class="btn green" onclick="openOrder()">開單</button>');
  p('  <button class="btn red"   onclick="closeOrder()">結單</button>');
  p('  <button class="btn orange" onclick="clearState()">清空狀態</button>');
  p('</div>');
  p('<div class="ctrl-row" style="margin-top:6px">');
  p('  <span style="font-size:13px;font-weight:600">自動結單：</span>');
  p('  <input type="number" id="autoMin" value="30" min="1" max="480"> 分鐘後');
  p('  <button class="btn outline" onclick="setAutoClose()">設定</button>');
  p('  <button class="btn outline" onclick="doCancelAc()">取消</button>');
  p('  <span class="ac-info" id="acInfo"></span>');
  p('</div>');
  p('</div>');

  // ── 今日統計 ────────────────────────────────────────────────
  p('<div class="card">');
  p('<h3>📊 今日統計</h3>');
  p('<div class="stats" id="statsBox"></div>');
  p('</div>');

  // ── 訂單列表 ────────────────────────────────────────────────
  p('<div class="card">');
  p('<h3>📋 訂單列表');
  p('  <div style="display:flex;gap:6px;flex-wrap:wrap">');
  p('    <button class="btn green sm" onclick="batchPaid()">✅ 全標已付款</button>');
  p('    <button class="btn outline sm" onclick="loadOrders()">🔄 重新整理</button>');
  p('    <button class="btn outline sm" onclick="copyShopOrder()">📋 複製店家單</button>');
  p('  </div>');
  p('</h3>');
  // 日期篩選
  p('<div class="date-row">');
  p('  <span style="font-size:13px;font-weight:600">查詢日期：</span>');
  p('  <input type="date" id="dateFilter" onchange="loadOrders()">');
  p('  <button class="btn outline sm" onclick="document.getElementById(\'dateFilter\').value=\'\';loadOrders()">今日</button>');
  p('</div>');
  p('<input class="search" id="search" placeholder="🔍 搜尋姓名、品項..." oninput="filterTable()">');
  p('<div class="ctrl-row" style="margin-bottom:10px">');
  p('  <label style="font-size:13px;font-weight:600">分組方式：</label>');
  p('  <select class="sel" id="groupBy" onchange="renderOrders()">');
  p('    <option value="none">不分組</option>');
  p('    <option value="store" selected>依店家</option>');
  p('    <option value="person">依人員</option>');
  p('  </select>');
  p('</div>');
  p('<div id="ordersWrap"></div>');
  p('</div>');

  p('</div>'); // .wrap

  // ════════════════════════════════════════════════════════════
  //  Admin JS
  // ════════════════════════════════════════════════════════════
  p('<script>');
  p('var TOKEN=(location.search.match(/[?&]token=([^&]*)/)||[])[1]||"";');
  p('var ordersCache=[],acAt=null,acTimer=null;');
  p('document.getElementById("hdTime").innerText=new Date().toLocaleString("zh-TW");');

  p('async function api(path,method,body){');
  p('  var opts={method:method||"GET",headers:{"Content-Type":"application/json","x-admin-token":TOKEN}};');
  p('  if(body!==undefined)opts.body=JSON.stringify(body);');
  p('  return (await fetch(path,opts)).json();');
  p('}');

  // status
  p('async function loadStatus(){');
  p('  var d=await api("/api/status");');
  p('  document.getElementById("sDot").className="status-dot"+(d.isOpen?" on":"");');
  p('  document.getElementById("sTxt").innerText=d.isOpen?"開單中":"已結單";');
  p('  if(d.autoCloseAt){acAt=d.autoCloseAt;startAcCd();}');
  p('  else{acAt=null;document.getElementById("acInfo").innerText="";}');
  p('}');
  p('function startAcCd(){');
  p('  if(acTimer)clearInterval(acTimer);');
  p('  acTimer=setInterval(function(){');
  p('    if(!acAt){clearInterval(acTimer);return;}');
  p('    var diff=new Date(acAt)-new Date();');
  p('    if(diff<=0){document.getElementById("acInfo").innerText="已自動結單";clearInterval(acTimer);loadStatus();return;}');
  p('    document.getElementById("acInfo").innerText="⏰ "+Math.floor(diff/60000)+"分"+Math.floor((diff%60000)/1000)+"秒後自動結單";');
  p('  },1000);');
  p('}');

  p('async function openOrder(){await api("/api/admin/open","POST",{});loadStatus();loadOrders();}');
  p('async function closeOrder(){await api("/api/admin/close","POST",{});loadStatus();loadOrders();}');
  p('async function clearState(){if(!confirm("確定清空開單狀態？"))return;await api("/api/admin/clear","POST",{});loadStatus();loadOrders();}');
  p('async function setAutoClose(){var m=Number(document.getElementById("autoMin").value)||30;await api("/api/admin/auto-close","POST",{minutes:m});loadStatus();}');
  p('async function doCancelAc(){await api("/api/admin/cancel-auto-close","POST",{});acAt=null;document.getElementById("acInfo").innerText="";clearInterval(acTimer);}');

  // load orders
  p('async function loadOrders(){');
  p('  var date=document.getElementById("dateFilter").value;');
  p('  var url="/api/admin/orders"+(date?"?date="+date:"");');
  p('  ordersCache=await api(url);');
  p('  if(!Array.isArray(ordersCache))ordersCache=[];');
  p('  renderStats(ordersCache);');
  p('  renderOrders();');
  p('}');

  // stats
  p('function renderStats(orders){');
  p('  var active=orders.filter(function(o){return o.status!=="已刪除";});');
  p('  var paid=active.filter(function(o){return o.status==="已付款";});');
  p('  var unpaid=active.filter(function(o){return o.status==="未付款";});');
  p('  var total=active.reduce(function(a,o){return a+o.total;},0);');
  p('  var paidM=paid.reduce(function(a,o){return a+o.total;},0);');
  p('  var unpaidM=unpaid.reduce(function(a,o){return a+o.total;},0);');
  p('  document.getElementById("statsBox").innerHTML=');
  p('    sb(active.length,"筆訂單","")+sb("$"+total,"總金額","green")+sb("$"+paidM,"已收款","green")+sb("$"+unpaidM,"未收款","red")+sb(paid.length,"已付款","green")+sb(unpaid.length,"未付款","orange");');
  p('  function sb(n,l,cls){return \'<div class="stat-box\'+(cls?" "+cls:"\'>")+\'<div class="stat-num">\'+n+\'</div><div class="stat-lbl">\'+l+\'</div></div>\';}');
  p('}');

  // render orders
  p('function renderOrders(){');
  p('  var q=(document.getElementById("search").value||"").toLowerCase();');
  p('  var gb=document.getElementById("groupBy").value;');
  p('  var filtered=ordersCache.filter(function(o){');
  p('    return !q||(o.name||"").toLowerCase().indexOf(q)>=0||(o.item||"").toLowerCase().indexOf(q)>=0;');
  p('  });');
  p('  var el=document.getElementById("ordersWrap");');
  p('  if(!filtered.length){el.innerHTML=\'<div style="text-align:center;padding:24px;color:#aaa">無訂單資料</div>\';return;}');
  p('  if(gb==="none"){el.innerHTML=buildTable(filtered);return;}');
  p('  // group');
  p('  var groups={};');
  p('  filtered.forEach(function(o){');
  p('    var key=gb==="store"?(o.store||"其他"):(o.name||"未知");');
  p('    if(!groups[key])groups[key]=[];');
  p('    groups[key].push(o);');
  p('  });');
  p('  var html="";');
  p('  Object.keys(groups).forEach(function(gk){');
  p('    var items=groups[gk];');
  p('    var subtotal=items.filter(function(o){return o.status!=="已刪除";}).reduce(function(a,o){return a+o.total;},0);');
  p('    html+=\'<div class="group-sec">\'');
  p('      +\'<div class="group-title"><span>\'+(gb==="store"?\'<span class="store-badge">\'+esc(gk)+\'</span>\':esc(gk))+\'</span>\'');
  p('      +\'<span class="group-sum">小計 $\'+subtotal+\' / \'+items.length+\' 筆</span></div>\';');
  p('    html+=buildTable(items);');
  p('    html+="</div>";');
  p('  });');
  p('  el.innerHTML=html;');
  p('}');
  p('function filterTable(){renderOrders();}');

  // buildTable
  p('function buildTable(orders){');
  p('  if(!orders.length)return "";');
  p('  var h=\'<div class="tbl-wrap"><table><thead><tr>\'');
  p('    +\'<th>姓名</th><th>店家</th><th>品項</th><th>規格</th><th>備註</th><th>數量</th><th>金額</th><th>狀態</th><th>操作</th>\'');
  p('    +\'</tr></thead><tbody>\';');
  p('  h+=orders.map(function(o){');
  p('    var tc=o.status==="已付款"?"tag-paid":o.status==="已刪除"?"tag-del":"tag-unpaid";');
  p('    var ops=o.status==="未付款"');
  p('      ?\'<select class="sel" id="pt\'+o.rowIndex+\'" style="padding:4px 8px;font-size:12px">\'');
  p('        +\'<option>現金</option><option>Line Pay</option><option>轉帳</option></select> \'');
  p('        +\'<button class="btn green sm" onclick="markPaid(\'+o.rowIndex+\')">付款</button> \'');
  p('        +\'<button class="btn red sm" onclick="adminDel(\'+o.rowIndex+\')">刪</button>\'');
  p('      :"";');
  p('    return \'<tr>\'');
  p('      +\'<td>\'+esc(o.name)+\'</td>\'');
  p('      +\'<td><span class="store-badge">\'+esc(o.store)+\'</span></td>\'');
  p('      +\'<td>\'+esc(o.item)+\'</td>\'');
  p('      +\'<td style="color:var(--sub)">\'+esc(o.spec)+\'</td>\'');
  p('      +\'<td style="color:var(--sub)">\'+esc(o.note)+\'</td>\'');
  p('      +\'<td>\'+o.qty+\'</td>\'');
  p('      +\'<td style="font-weight:700">$\'+o.total+\'</td>\'');
  p('      +\'<td><span class="tag \'+tc+\'">\'+esc(o.status)+\'</span></td>\'');
  p('      +\'<td>\'+ops+\'</td>\'');
  p('      +\'</tr>\';');
  p('  }).join("");');
  p('  h+="</tbody></table></div>";');
  p('  return h;');
  p('}');

  p('async function markPaid(rowIndex){');
  p('  var sel=document.getElementById("pt"+rowIndex);');
  p('  await api("/api/admin/paid","POST",{rowIndex:rowIndex,payType:sel?sel.value:"現金"});');
  p('  loadOrders();');
  p('}');
  p('async function adminDel(rowIndex){');
  p('  if(!confirm("確定刪除？"))return;');
  p('  await api("/api/admin/delete-order","POST",{rowIndex:rowIndex});');
  p('  loadOrders();');
  p('}');
  p('async function batchPaid(){');
  p('  if(!confirm("將所有未付款訂單標記為已付款？"))return;');
  p('  await api("/api/admin/batch-paid","POST",{payType:"現金"});');
  p('  loadOrders();');
  p('}');
  p('function copyShopOrder(){');
  p('  var active=ordersCache.filter(function(o){return o.status!=="已刪除";});');
  p('  var counts={};');
  p('  active.forEach(function(o){var k=o.item+(o.spec?"（"+o.spec+"）":"");counts[k]=(counts[k]||0)+o.qty;});');
  p('  var txt="您好，今天訂購如下：\n\n",n=0;');
  p('  Object.keys(counts).forEach(function(k){txt+=k+" x"+counts[k]+"\n";n+=counts[k];});');
  p('  var m=active.reduce(function(a,o){return a+o.total;},0);');
  p('  txt+="\n總數："+n+"份\n總金額："+m+"元\n\n麻煩您，謝謝～";');
  p('  navigator.clipboard.writeText(txt).then(function(){alert("已複製到剪貼板");});');
  p('}');
  p('function esc(s){return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");}');

  p('loadStatus();loadOrders();');
  p('setInterval(loadStatus,30000);setInterval(loadOrders,60000);');
  p('</script></body></html>');

  return lines.join('\n');
}

// ════════════════════════════════════════════════════════════════
//  LINE Webhook（body 由 line.middleware 解析，不碰 express.json）
// ════════════════════════════════════════════════════════════════
app.post('/webhook', async (req, res) => {
  try {
    const event = req.body.events?.[0];
    if (!event || event.type !== 'message' || event.message.type !== 'text')
      return res.sendStatus(200);

    let profileName = '未知使用者';
    try {
      if (event.source.type === 'group') {
        const p = await client.getGroupMemberProfile(event.source.groupId, event.source.userId);
        profileName = p.displayName;
      } else {
        const p = await client.getProfile(event.source.userId);
        profileName = p.displayName;
      }
    } catch (e) { console.error('取得名稱失敗：', e.message); }

    knownUsers[profileName] = event.source.userId;
    saveUserToSheet(profileName, event.source.userId, event.source.type, event.source.groupId || '').catch(() => {});

    const uid        = event.source.userId;
    const text       = event.message.text.trim();
    const replyToken = event.replyToken;
    const reply      = t => client.replyMessage(replyToken, { type: 'text', text: t });

    // ── 開單 N 分鐘
    const autoOpen = text.match(/^開單\s+(\d+)$/);
    if (autoOpen) {
      if (!isAdmin(uid)) { await reply('只有管理員可以開單'); return res.sendStatus(200); }
      isOpen = true;
      scheduleAutoClose(Number(autoOpen[1]));
      await reply('已開單！將於 ' + autoOpen[1] + ' 分鐘後自動結單\n點餐頁：' + (process.env.LIFF_URL || ''));
      return res.sendStatus(200);
    }

    if (text === '開單') {
      if (!isAdmin(uid)) { await reply('只有管理員可以開單'); return res.sendStatus(200); }
      if (isOpen) { await reply('目前已開單中'); return res.sendStatus(200); }
      isOpen = true;
      await reply('已開單，可以開始點餐 🍱\n點餐頁：' + (process.env.LIFF_URL || ''));
      return res.sendStatus(200);
    }

    if (text === '結單' || text === '收單' || text === '統計') {
      if (!isAdmin(uid)) { await reply('只有管理員可以結單 / 統計'); return res.sendStatus(200); }
      isOpen = false; cancelAutoClose();
      await reply(await buildStatReport());
      return res.sendStatus(200);
    }

    if (text === '店家單') {
      if (!isAdmin(uid)) { await reply('只有管理員可以查看店家單'); return res.sendStatus(200); }
      await reply(await buildShopOrder());
      return res.sendStatus(200);
    }

    if (text === '清空') {
      if (!isAdmin(uid)) { await reply('只有管理員可以清空'); return res.sendStatus(200); }
      isOpen = false; cancelAutoClose();
      await reply('已清空訂單');
      return res.sendStatus(200);
    }

    if (text === '後台') {
      if (!isAdmin(uid)) { await reply('只有管理員可以查看後台'); return res.sendStatus(200); }
      await reply('管理後台：' + (process.env.APP_URL || '') + '/admin?token=' + (process.env.ADMIN_TOKEN || ''));
      return res.sendStatus(200);
    }

    if (text === '狀態') {
      const msg = (isOpen ? '🟢 目前開單中' : '🔴 目前未開單') +
        (autoCloseAt ? '\n⏰ 自動結單：' + new Date(autoCloseAt).toLocaleTimeString('zh-TW', { timeZone: 'Asia/Taipei' }) : '');
      await reply(msg);
      return res.sendStatus(200);
    }

    return res.sendStatus(200);
  } catch (err) {
    console.error('Webhook error:', err);
    return res.sendStatus(200);
  }
});

app.use((err, _q, res, _n) => { console.error('Global error:', err); res.sendStatus(200); });

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log('Server running on port', PORT));
