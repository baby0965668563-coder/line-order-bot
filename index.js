require('dotenv').config();

const express = require('express');
const line    = require('@line/bot-sdk');
const { GoogleSpreadsheet } = require('google-spreadsheet');

const app = express();

const lineConfig = {
  channelAccessToken: process.env.CHANNEL_ACCESS_TOKEN,
  channelSecret:      process.env.CHANNEL_SECRET
};

// ★ LINE webhook 必須在 express.json() 之前
app.use('/webhook', line.middleware(lineConfig));
app.use('/api',    express.json());
app.use('/admin',  express.json());

const client = new line.Client(lineConfig);
const doc    = new GoogleSpreadsheet(process.env.SHEET_ID);

// ════════════════════════════════════════════════════════════════
//  全域狀態
// ════════════════════════════════════════════════════════════════
let isOpen         = false;
let autoCloseTimer = null;
let autoCloseAt    = null;
const knownUsers   = {};

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
//  菜單
// ════════════════════════════════════════════════════════════════
async function loadMenu() {
  await authSheet();
  const sheet = doc.sheetsByTitle['Menu'];
  if (!sheet) return [];
  const rows = await sheet.getRows();
  return rows
    .map(r => ({
      store: String(r['店家']    || '').trim(),
      item:  String(r['品項']    || '').trim(),
      price: Number(r['價格']    || 0),
      image: String(r['圖片URL'] || '').trim()
    }))
    .filter(r => r.store && r.item && r.price > 0);
}

async function loadOptions() {
  await authSheet();
  const gs = doc.sheetsByTitle['OptionGroups'];
  const os = doc.sheetsByTitle['Options'];
  if (!gs || !os) return {};
  const gr = await gs.getRows();
  const or = await os.getRows();
  const result = {};
  gr.forEach(g => {
    const store = String(g['店家'] || '').trim();
    const item  = String(g['品項'] || '').trim();
    const cat   = String(g['分類'] || '').trim();
    if (!store || !item || !cat) return;
    const key = store + '||' + item;
    if (!result[key]) result[key] = [];
    const opts = or
      .filter(o => String(o['店家']||'').trim()===store && String(o['品項']||'').trim()===item && String(o['分類']||'').trim()===cat)
      .map(o => String(o['選項']||'').trim()).filter(Boolean);
    result[key].push({ category:cat, required:String(g['必選']||'').trim()==='TRUE', min:Number(g['最少']||0), max:Number(g['最多']||0), options:opts });
  });
  return result;
}

// ════════════════════════════════════════════════════════════════
//  Users
// ════════════════════════════════════════════════════════════════
async function saveUserToSheet(name, userId, src, gid) {
  try {
    await authSheet();
    const s = doc.sheetsByTitle['Users'];
    if (!s) return;
    await s.addRow({ 時間:nowTW(), LINE名稱:name, userId, 來源類型:src, 群組ID:gid||'', 權限:isAdmin(userId)?'admin':'user' });
  } catch(e) { console.error('Users write fail:', e.message); }
}

// ════════════════════════════════════════════════════════════════
//  Orders CRUD
// ════════════════════════════════════════════════════════════════
async function saveOrderToSheet(order) {
  try {
    await authSheet();
    const s = doc.sheetsByTitle['Orders'];
    if (!s) return { success:false, reason:'no_sheet' };
    const rows  = await s.getRows();
    const today = todayTW();
    const dup = rows.find(r =>
      String(r['userId']||'')=== String(order.userId) &&
      String(r['品項']  ||'').trim()===String(order.item||'').trim() &&
      String(r['規格']  ||'').trim()===String(order.spec||'').trim() &&
      String(r['時間']  ||'').startsWith(today) &&
      String(r['狀態']  ||'')!=='已刪除'
    );
    if (dup) return { success:false, reason:'duplicate' };
    const qty=Number(order.qty||1), price=Number(order.price||0);
    await s.addRow({
      時間:nowTW(), LINE名稱:String(order.name||''), userId:String(order.userId||''),
      店家:String(order.store||''), 品項:String(order.item||''), 規格:String(order.spec||''),
      備註:String(order.note||''), 數量:qty, 單價:price, 總價:price*qty,
      狀態:'未付款', 付款時間:'', 付款方式:'', 訂單備註:''
    });
    return { success:true };
  } catch(e) { console.error('Orders write fail:', e.message); return { success:false, reason:e.message }; }
}

async function getOrdersByUser(userId, dateStr) {
  try {
    await authSheet();
    const s = doc.sheetsByTitle['Orders'];
    if (!s) return [];
    const rows = await s.getRows();
    const target = dateStr || todayTW();
    return rows
      .filter(r => String(r['userId']||'')===String(userId) && String(r['時間']||'').startsWith(target) && String(r['狀態']||'')!=='已刪除')
      .map(r => ({
        rowIndex:r.rowIndex, store:String(r['店家']||''), item:String(r['品項']||''),
        spec:String(r['規格']||''), note:String(r['備註']||''),
        qty:Number(r['數量']||1), price:Number(r['單價']||0), total:Number(r['總價']||0),
        status:String(r['狀態']||'未付款'), time:String(r['時間']||'')
      }));
  } catch(e) { return []; }
}

async function getOrderDates(userId) {
  try {
    await authSheet();
    const s = doc.sheetsByTitle['Orders'];
    if (!s) return [];
    const rows = await s.getRows();
    const dates = new Set();
    rows.forEach(r => {
      if (String(r['userId']||'')===userId && String(r['狀態']||'')!=='已刪除') {
        const d = String(r['時間']||'').split(' ')[0];
        if (d) dates.add(d);
      }
    });
    return [...dates].sort().reverse().slice(0, 14);
  } catch(e) { return []; }
}

async function deleteOrder(userId, rowIndex) {
  try {
    await authSheet();
    const s = doc.sheetsByTitle['Orders'];
    if (!s) return false;
    const rows = await s.getRows();
    const t = rows.find(r => Number(r.rowIndex)===Number(rowIndex) && String(r['userId']||'')===String(userId) && String(r['狀態']||'')!=='已刪除');
    if (!t) return false;
    t['狀態']='已刪除'; await t.save(); return true;
  } catch(e) { return false; }
}

async function updateOrderNote(userId, rowIndex, note) {
  try {
    await authSheet();
    const s = doc.sheetsByTitle['Orders'];
    if (!s) return false;
    const rows = await s.getRows();
    const t = rows.find(r => Number(r.rowIndex)===Number(rowIndex) && String(r['userId']||'')===String(userId) && String(r['狀態']||'')!=='已刪除');
    if (!t) return false;
    t['備註']=note; await t.save(); return true;
  } catch(e) { return false; }
}

async function getAllOrdersByDate(dateStr) {
  try {
    await authSheet();
    const s = doc.sheetsByTitle['Orders'];
    if (!s) return [];
    const rows = await s.getRows();
    const target = dateStr || todayTW();
    return rows
      .filter(r => String(r['時間']||'').startsWith(target))
      .map(r => ({
        rowIndex:r.rowIndex, name:String(r['LINE名稱']||''), userId:String(r['userId']||''),
        store:String(r['店家']||''), item:String(r['品項']||''), spec:String(r['規格']||''),
        note:String(r['備註']||''), qty:Number(r['數量']||1), price:Number(r['單價']||0),
        total:Number(r['總價']||0), status:String(r['狀態']||'未付款'),
        payTime:String(r['付款時間']||''), payType:String(r['付款方式']||'')
      }));
  } catch(e) { return []; }
}

async function markPaid(rowIndex, payType) {
  try {
    await authSheet();
    const s = doc.sheetsByTitle['Orders'];
    if (!s) return false;
    const rows = await s.getRows();
    const t = rows.find(r => Number(r.rowIndex)===Number(rowIndex));
    if (!t) return false;
    t['狀態']='已付款'; t['付款時間']=nowTW(); t['付款方式']=payType||'現金';
    await t.save(); return true;
  } catch(e) { return false; }
}

async function batchMarkPaid(payType) {
  try {
    await authSheet();
    const s = doc.sheetsByTitle['Orders'];
    if (!s) return 0;
    const rows = await s.getRows();
    const today = todayTW();
    let count = 0;
    for (const r of rows) {
      if (String(r['時間']||'').startsWith(today) && String(r['狀態']||'')==='未付款') {
        r['狀態']='已付款'; r['付款時間']=nowTW(); r['付款方式']=payType||'現金';
        await r.save(); count++;
      }
    }
    return count;
  } catch(e) { return 0; }
}

async function adminDeleteOrder(rowIndex) {
  try {
    await authSheet();
    const s = doc.sheetsByTitle['Orders'];
    if (!s) return false;
    const rows = await s.getRows();
    const t = rows.find(r => Number(r.rowIndex)===Number(rowIndex));
    if (!t) return false;
    t['狀態']='已刪除'; await t.save(); return true;
  } catch(e) { return false; }
}

// ════════════════════════════════════════════════════════════════
//  自動結單
// ════════════════════════════════════════════════════════════════
function scheduleAutoClose(minutes) {
  if (autoCloseTimer) clearTimeout(autoCloseTimer);
  const ms = Math.max(1, Number(minutes)) * 60 * 1000;
  autoCloseAt    = new Date(Date.now() + ms).toISOString();
  autoCloseTimer = setTimeout(() => {
    isOpen = false; autoCloseAt = null; autoCloseTimer = null;
    console.log('[自動結單]', nowTW());
  }, ms);
}
function cancelAutoClose() {
  if (autoCloseTimer) clearTimeout(autoCloseTimer);
  autoCloseTimer = null; autoCloseAt = null;
}

// ════════════════════════════════════════════════════════════════
//  LINE 報表
// ════════════════════════════════════════════════════════════════
async function buildStatReport() {
  const orders = await getAllOrdersByDate(todayTW());
  const active = orders.filter(o => o.status !== '已刪除');
  if (!active.length) return '📊 今日尚無訂單';
  const itemCount={}, userTotal={};
  const unpaid = new Set();
  for (const o of active) {
    const k = o.item+(o.spec?'（'+o.spec+'）':'');
    itemCount[k]=(itemCount[k]||0)+o.qty;
    const n = o.name||o.userId||'未知';
    userTotal[n]=(userTotal[n]||0)+o.total;
    if (o.status==='未付款') unpaid.add(n);
  }
  const grand = Object.values(userTotal).reduce((a,b)=>a+b,0);
  let t = '📊 今日訂餐統計\n─────────────\n【品項數量】\n';
  for (const k in itemCount) t += k+' x'+itemCount[k]+'\n';
  t += '\n【個人金額】\n';
  for (const n in userTotal) t += n+'：$'+userTotal[n]+'\n';
  t += '\n💰 總金額：$'+grand;
  t += unpaid.size ? '\n\n⚠️ 未付款：'+[...unpaid].join('、') : '\n\n✅ 所有人已付款';
  return t;
}

async function buildShopOrder() {
  const orders = await getAllOrdersByDate(todayTW());
  const active = orders.filter(o => o.status !== '已刪除');
  if (!active.length) return '今日尚無訂單';
  const itemCount={};
  for (const o of active) {
    const k = o.item+(o.spec?'（'+o.spec+'）':'');
    itemCount[k]=(itemCount[k]||0)+o.qty;
  }
  let out='您好，今天訂購如下：\n\n', total=0;
  for (const k in itemCount) { out+=k+' x'+itemCount[k]+'\n'; total+=itemCount[k]; }
  const money = active.reduce((a,o)=>a+o.total,0);
  return out+'\n總數：'+total+'份\n總金額：'+money+'元\n\n麻煩您，謝謝～';
}

// ════════════════════════════════════════════════════════════════
//  Admin middleware
// ════════════════════════════════════════════════════════════════
function adminAuth(req, res, next) {
  const t = req.headers['x-admin-token'] || req.query.token || '';
  if (!process.env.ADMIN_TOKEN || t !== process.env.ADMIN_TOKEN)
    return res.status(401).json({ error:'Unauthorized' });
  next();
}

// ════════════════════════════════════════════════════════════════
//  API 路由
// ════════════════════════════════════════════════════════════════
app.get('/', (_q,r) => r.send('LINE 訂餐機器人運作中'));

app.get('/api/menu', async (_q, res) => {
  try {
    const [menu, optionData] = await Promise.all([loadMenu(), loadOptions()]);
    res.json({ menu, optionData });
  } catch(e) { res.status(500).json({ menu:[], optionData:{} }); }
});

app.get('/api/status', (_q, res) => res.json({ isOpen, autoCloseAt }));

app.get('/api/my-dates', async (req, res) => {
  const { userId } = req.query;
  if (!userId) return res.json([]);
  res.json(await getOrderDates(userId));
});

app.post('/api/order', async (req, res) => {
  res.json(await saveOrderToSheet(req.body));
});

app.get('/api/my-orders', async (req, res) => {
  const { userId, date } = req.query;
  if (!userId) return res.json([]);
  res.json(await getOrdersByUser(userId, date||null));
});

app.delete('/api/order/:ri', async (req, res) => {
  const { userId } = req.body;
  if (!userId) return res.json({ success:false });
  res.json({ success: await deleteOrder(userId, Number(req.params.ri)) });
});

app.patch('/api/order/:ri/note', async (req, res) => {
  const { userId, note } = req.body;
  if (!userId) return res.json({ success:false });
  res.json({ success: await updateOrderNote(userId, Number(req.params.ri), note) });
});

// Admin API
app.post('/api/admin/open', adminAuth, (_q,res) => { isOpen=true; res.json({ isOpen }); });
app.post('/api/admin/close', adminAuth, (_q,res) => { isOpen=false; cancelAutoClose(); res.json({ isOpen }); });
app.post('/api/admin/clear', adminAuth, (_q,res) => { isOpen=false; cancelAutoClose(); res.json({ ok:true }); });
app.post('/api/admin/auto-close', adminAuth, (req,res) => {
  if (!isOpen) isOpen=true;
  scheduleAutoClose(Number(req.body.minutes)||30);
  res.json({ autoCloseAt });
});
app.post('/api/admin/cancel-auto-close', adminAuth, (_q,res) => { cancelAutoClose(); res.json({ ok:true }); });
app.get('/api/admin/orders', adminAuth, async (req,res) => {
  res.json(await getAllOrdersByDate(req.query.date || todayTW()));
});
app.post('/api/admin/paid', adminAuth, async (req,res) => {
  res.json({ success: await markPaid(req.body.rowIndex, req.body.payType) });
});
app.post('/api/admin/batch-paid', adminAuth, async (req,res) => {
  res.json({ success:true, count: await batchMarkPaid(req.body.payType||'現金') });
});
app.post('/api/admin/delete-order', adminAuth, async (req,res) => {
  res.json({ success: await adminDeleteOrder(req.body.rowIndex) });
});

// ════════════════════════════════════════════════════════════════
//  前台頁面（Array join，零 template literal，零 JSON 內嵌）
// ════════════════════════════════════════════════════════════════
app.get('/order', (_q, res) => {
  const liffId = String(process.env.LIFF_ID || '2010025093-yATK02dc').replace(/[^a-zA-Z0-9\-]/g,'');
  res.setHeader('Content-Type','text/html; charset=utf-8');
  res.end(buildOrderPage(liffId));
});

function buildOrderPage(liffId) {
  const L = [];
  const p = s => L.push(s);

  /* ── HEAD ─────────────────────────────────────────────────── */
  p('<!DOCTYPE html><html lang="zh-TW"><head>');
  p('<meta charset="utf-8">');
  p('<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1,user-scalable=no">');
  p('<title>訂餐小幫手</title>');
  p('<script src="https://static.line-scdn.net/liff/edge/2/sdk.js"></script>');
  p('<style>');
  p(':root{--g:#06c755;--gd:#05a847;--r:#e53935;--bg:#f5f5f5;--card:#fff;--bdr:#e8e8e8;--txt:#1a1a1a;--sub:#888;--rr:16px;--sh:0 2px 12px rgba(0,0,0,.08);--hd-h:84px;--srch-h:45px}');
  p('*{box-sizing:border-box;-webkit-tap-highlight-color:transparent;margin:0;padding:0}');
  p('html,body{height:auto!important;overflow-x:hidden;overflow-y:auto;-webkit-overflow-scrolling:touch}');
  p('body{font-family:-apple-system,Arial,"Microsoft JhengHei",sans-serif;background:var(--bg);color:var(--txt);min-height:100vh}');

  /* header */
  p('.hd{background:#fff;border-bottom:1px solid var(--bdr);position:sticky;top:0;z-index:100;padding:10px 14px;will-change:transform}');
  p('.hd-top{display:flex;align-items:center;justify-content:space-between;margin-bottom:4px}');
  p('.hd-title{display:flex;align-items:center;gap:8px;font-size:19px;font-weight:800}');
  p('.hd-badge{font-size:11px;padding:2px 9px;border-radius:999px;font-weight:700}');
  p('.hd-badge.open{background:#e8f5e9;color:#2e7d32}');
  p('.hd-badge.closed{background:#fce4ec;color:#c62828}');
  p('.hd-user{font-size:12px;color:var(--sub)}');
  p('.hd-cd{font-size:11px;color:var(--r);font-weight:700;margin-top:2px}');

  /* search */
  p('.srch{padding:8px 14px;background:#fff;border-bottom:1px solid var(--bdr);position:sticky;top:var(--hd-h);z-index:99;will-change:transform;height:var(--srch-h)}');
  p('.srch input{width:100%;padding:9px 14px 9px 36px;border-radius:999px;border:1.5px solid var(--bdr);font-size:14px;background:var(--bg);outline:none;background-image:url("data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'16\' height=\'16\' fill=\'%23888\' viewBox=\'0 0 16 16\'%3E%3Cpath d=\'M11.742 10.344a6.5 6.5 0 1 0-1.397 1.398l3.85 3.85a1 1 0 0 0 1.415-1.415l-3.868-3.833zM12 6.5a5.5 5.5 0 1 1-11 0 5.5 5.5 0 0 1 11 0z\'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:12px center}');
  p('.srch input:focus{border-color:var(--g)}');

  /* tabs */
  p('.tabs{display:flex;flex-wrap:nowrap;gap:8px;padding:8px 12px;overflow-x:auto;overflow-y:hidden;-webkit-overflow-scrolling:touch;scrollbar-width:none;background:#fff;border-bottom:1px solid var(--bdr);position:sticky;top:calc(var(--hd-h) + var(--srch-h));z-index:89}');
  p('.tabs::-webkit-scrollbar{display:none}');
  p('.tab{display:inline-flex;align-items:center;gap:4px;padding:7px 14px;border-radius:999px;border:1.5px solid var(--bdr);font-size:13px;font-weight:600;cursor:pointer;background:#fff;color:var(--sub);flex:0 0 auto;min-width:max-content;white-space:nowrap;transition:all .15s;flex-shrink:0}');
  p('.tab.active{background:var(--g);color:#fff;border-color:var(--g)}');
  p('.tab.starred{border-color:#fdd835;color:#f9a825}');

  /* container */
  p('.ctn{padding:10px 10px 160px;-webkit-overflow-scrolling:touch}');
  p('.store-row{display:flex;align-items:center;justify-content:space-between;padding:6px 4px;margin:14px 0 6px}');
  p('.store-name{font-size:16px;font-weight:800}');
  p('.fav-btn{background:none;border:none;font-size:20px;cursor:pointer;padding:2px 6px;line-height:1;color:var(--sub)}');

  /* card */
  p('.card{background:var(--card);border-radius:var(--rr);margin-bottom:10px;box-shadow:var(--sh);overflow:hidden}');
  p('.card-img{width:100%;height:150px;object-fit:cover;display:block}');
  p('.card-body{padding:12px 14px 14px}');
  p('.card-store{font-size:11px;color:var(--sub);margin-bottom:3px;font-weight:500}');
  p('.card-name{font-size:17px;font-weight:800;margin-bottom:5px;line-height:1.3}');
  p('.card-price{font-size:16px;color:var(--g);font-weight:800;margin-bottom:10px}');
  p('.add-btn{width:100%;padding:11px;border:none;border-radius:999px;background:var(--g);color:#fff;font-size:15px;font-weight:700;cursor:pointer}');
  p('.add-btn:disabled{background:#ccc;cursor:default}');

  /* FABs */
  p('.fab-area{position:fixed;bottom:20px;right:14px;display:flex;flex-direction:column;align-items:flex-end;gap:10px;z-index:8000;pointer-events:none}');
  p('.fab{pointer-events:auto}');
  p('.fab{display:flex;align-items:center;gap:6px;padding:13px 18px;border:none;border-radius:999px;font-size:14px;font-weight:700;cursor:pointer;box-shadow:0 4px 18px rgba(0,0,0,.18);white-space:nowrap}');
  p('.fab-cart{background:var(--g);color:#fff}');
  p('.fab-orders{background:#fff;color:var(--txt);border:1.5px solid var(--bdr)}');
  p('.fab-hidden{display:none!important}');
  p('.badge{background:var(--r);color:#fff;border-radius:50%;min-width:19px;height:19px;font-size:11px;display:inline-flex;align-items:center;justify-content:center;padding:0 3px}');

  /* modal */
  p('.modal{display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:9000;align-items:flex-end;justify-content:center}');
  p('.modal.ctr{align-items:center}');
  p('.modal.show{display:flex}');
  p('.mbox{background:#fff;width:100%;max-width:500px;max-height:88vh;overflow-y:auto;border-radius:22px 22px 0 0;padding:20px 16px 32px;position:relative}');
  p('.modal.ctr .mbox{border-radius:18px;margin:0 12px;max-height:82vh}');
  p('.m-close{position:absolute;top:14px;right:14px;background:var(--bg);border:none;border-radius:50%;width:30px;height:30px;font-size:15px;cursor:pointer;display:flex;align-items:center;justify-content:center;color:var(--sub)}');
  p('.m-title{font-size:18px;font-weight:800;margin-bottom:14px;padding-right:36px}');

  /* options */
  p('.opt-g-title{font-size:13px;font-weight:700;color:var(--sub);margin:14px 0 8px}');
  p('.chips{display:flex;flex-wrap:wrap;gap:8px}');
  p('.chip{padding:7px 16px;border-radius:999px;border:1.5px solid var(--bdr);font-size:14px;cursor:pointer;background:#fff;transition:.12s;user-select:none}');
  p('.chip.on{background:var(--g);color:#fff;border-color:var(--g)}');
  p('.chip.off{opacity:.35;cursor:default}');

  /* qty */
  p('.qty-row{display:flex;align-items:center;justify-content:space-between;margin-top:16px}');
  p('.qty-row label{font-weight:700;font-size:14px}');
  p('.qty-ctrl{display:flex;align-items:center;gap:0;border:1.5px solid var(--bdr);border-radius:999px;overflow:hidden}');
  p('.qty-ctrl button{background:none;border:none;width:38px;height:38px;font-size:22px;cursor:pointer;color:var(--g);font-weight:700;display:flex;align-items:center;justify-content:center}');
  p('.qty-ctrl span{min-width:32px;text-align:center;font-size:16px;font-weight:700}');
  p('.note-input{width:100%;padding:10px 14px;border-radius:12px;border:1.5px solid var(--bdr);font-size:14px;margin-top:12px;outline:none}');
  p('.note-input:focus{border-color:var(--g)}');

  /* buttons */
  p('.btn-p{width:100%;padding:13px;border:none;border-radius:999px;background:var(--g);color:#fff;font-size:16px;font-weight:700;cursor:pointer;margin-top:14px}');
  p('.btn-p:disabled{background:#ccc}');
  p('.btn-s{width:100%;padding:12px;border:1.5px solid var(--bdr);border-radius:999px;background:#fff;color:var(--txt);font-size:15px;font-weight:600;cursor:pointer;margin-top:8px}');
  p('.btn-del{background:var(--r);color:#fff;border:none;border-radius:999px;padding:6px 13px;font-size:12px;cursor:pointer;font-weight:700}');
  p('.btn-reorder{background:var(--bg);border:none;border-radius:999px;padding:5px 12px;font-size:12px;cursor:pointer;font-weight:700;color:var(--g);margin-top:4px}');

  /* order rows */
  p('.orow{display:flex;align-items:flex-start;gap:10px;padding:12px 0;border-bottom:1px solid var(--bg)}');
  p('.oinfo{flex:1;min-width:0}');
  p('.oname{font-weight:700;font-size:15px;margin-bottom:2px}');
  p('.osub{font-size:12px;color:var(--sub);margin-bottom:2px}');
  p('.oprice{font-size:15px;font-weight:800;color:var(--g);white-space:nowrap}');
  p('.oact{display:flex;flex-direction:column;align-items:flex-end;gap:5px}');
  p('.tag-paid{display:inline-block;padding:2px 9px;border-radius:999px;font-size:11px;font-weight:700;background:#e8f5e9;color:#2e7d32}');
  p('.tag-unpaid{display:inline-block;padding:2px 9px;border-radius:999px;font-size:11px;font-weight:700;background:#fff3e0;color:#e65100}');
  p('.icon-btn{background:none;border:none;font-size:18px;cursor:pointer;padding:4px}');

  /* date chips */
  p('.dcwrap{display:flex;gap:7px;overflow-x:auto;padding-bottom:8px;margin-bottom:10px;scrollbar-width:none}');
  p('.dcwrap::-webkit-scrollbar{display:none}');
  p('.dchip{padding:5px 13px;border-radius:999px;border:1.5px solid var(--bdr);font-size:12px;cursor:pointer;white-space:nowrap;background:#fff;flex-shrink:0}');
  p('.dchip.on{background:var(--g);color:#fff;border-color:var(--g)}');

  /* total */
  p('.total-bar{font-size:16px;font-weight:800;margin:12px 0 4px}');

  /* empty */
  p('.empty{text-align:center;padding:40px 20px;color:var(--sub);font-size:15px}');
  p('#lm{text-align:center;padding:48px 20px;color:var(--sub);font-size:15px}');

  /* toast */
  p('#toast{position:fixed;bottom:100px;left:50%;transform:translateX(-50%);background:#111;color:#fff;padding:10px 22px;border-radius:999px;font-size:14px;z-index:99999;opacity:0;transition:opacity .25s;pointer-events:none;white-space:nowrap}');
  p('</style></head><body>');

  /* ── HEADER ──────────────────────────────────────────────── */
  p('<div class="hd">');
  p('  <div class="hd-top">');
  p('    <div class="hd-title"><span>🍱</span><span>訂餐小幫手</span></div>');
  p('    <span class="hd-badge closed" id="hdBadge">未開放</span>');
  p('  </div>');
  p('  <div class="hd-user" id="hdUser">正在取得 LINE 使用者資料...</div>');
  p('  <div class="hd-cd" id="hdCd"></div>');
  p('</div>');

  /* ── SEARCH ──────────────────────────────────────────────── */
  p('<div class="srch"><input id="srchInput" type="search" placeholder="搜尋品項或店家..." oninput="onSearch()"></div>');

  /* ── TABS ────────────────────────────────────────────────── */
  p('<div class="tabs" id="tabsEl"></div>');

  /* ── MENU ────────────────────────────────────────────────── */
  p('<div class="ctn" id="menuEl"><div id="lm">菜單載入中...</div></div>');

  /* ── FABs ────────────────────────────────────────────────── */
  p('<div class="fab-area">');
  p('  <button class="fab fab-orders fab-hidden" id="fabOrders" onclick="openModal(\'moOrders\')">📋 我的訂單</button>');
  p('  <button class="fab fab-cart fab-hidden" id="fabCart" onclick="openModal(\'moCart\')">🛒 購物車<span class="badge" id="cartBadge">0</span></button>');
  p('</div>');

  /* ── TOAST ───────────────────────────────────────────────── */
  p('<div id="toast"></div>');

  /* ── MODAL: 選項 ─────────────────────────────────────────── */
  p('<div class="modal" id="moOpts"><div class="mbox">');
  p('  <button class="m-close" onclick="closeModal(\'moOpts\')">✕</button>');
  p('  <div class="m-title" id="moOptsTitle"></div>');
  p('  <div id="moOptsBody"></div>');
  p('  <div class="qty-row"><label>數量</label>');
  p('    <div class="qty-ctrl">');
  p('      <button onclick="chQty(-1)">−</button>');
  p('      <span id="qtyVal">1</span>');
  p('      <button onclick="chQty(1)">＋</button>');
  p('    </div>');
  p('  </div>');
  p('  <input class="note-input" id="itemNote" placeholder="備註（例：不加辣、少冰）">');
  p('  <button class="btn-p" onclick="submitOpts()">加入購物車 🛒</button>');
  p('  <button class="btn-s" onclick="closeModal(\'moOpts\')">取消</button>');
  p('</div></div>');

  /* ── MODAL: 購物車 ───────────────────────────────────────── */
  p('<div class="modal" id="moCart"><div class="mbox">');
  p('  <button class="m-close" onclick="closeModal(\'moCart\')">✕</button>');
  p('  <div class="m-title">🛒 我的購物車</div>');
  p('  <div id="cartList"></div>');
  p('  <div class="total-bar" id="cartTotal"></div>');
  p('  <button class="btn-p" id="submitBtn" onclick="submitCart()">送出訂單</button>');
  p('  <button class="btn-s" onclick="closeModal(\'moCart\')">繼續點餐</button>');
  p('</div></div>');

  /* ── MODAL: 我的訂單 ─────────────────────────────────────── */
  p('<div class="modal" id="moOrders"><div class="mbox">');
  p('  <button class="m-close" onclick="closeModal(\'moOrders\')">✕</button>');
  p('  <div class="m-title">📋 我的訂單</div>');
  p('  <div class="dcwrap" id="dateChips"></div>');
  p('  <div id="orderList"></div>');
  p('  <div class="total-bar" id="orderTotal"></div>');
  p('  <button class="btn-s" onclick="closeModal(\'moOrders\')">關閉</button>');
  p('</div></div>');

  /* ── MODAL: 備註編輯 ─────────────────────────────────────── */
  p('<div class="modal ctr" id="moNote"><div class="mbox">');
  p('  <button class="m-close" onclick="closeModal(\'moNote\')">✕</button>');
  p('  <div class="m-title">✏️ 修改備註</div>');
  p('  <input class="note-input" id="noteInput" placeholder="備註內容">');
  p('  <button class="btn-p" onclick="saveNote()">儲存</button>');
  p('  <button class="btn-s" onclick="closeModal(\'moNote\')">取消</button>');
  p('</div></div>');

  /* ════════════════════════════════════════════════════════════
     JavaScript — 全部用一般字串，不用 template literal
     無任何 JSON 內嵌，資料全部由 fetch 取得
  ══════════════════════════════════════════════════════════════ */
  p('<script>');

  /* 全域變數 */
  p("var LIFF_ID='" + liffId + "';");
  p('var menu=[],opts={},favs=[];');
  p('var profile=null,ready=false;');
  p('var curItem=null,curGroups=[],curQty=1;');
  p('var cart=[];');
  p('var editRI=null;');
  p('var orderDates=[],curDate=null;');
  p('var searchQ="";');

  /* loadMenuData */
  p('async function loadMenuData(){');
  p('  try{');
  p('    var r=await fetch("/api/menu");');
  p('    var d=await r.json();');
  p('    menu=d.menu||[];');
  p('    opts=d.optionData||{};');
  p('    renderMenu();');
  p('  }catch(e){');
  p('    var lm=document.getElementById("lm");');
  p('    if(lm)lm.textContent="菜單載入失敗，請重新整理";');
  p('  }');
  p('}');

  /* initLIFF */
  p('async function initLIFF(){');
  p('  try{');
  p('    await liff.init({liffId:LIFF_ID});');
  p('    if(!liff.isLoggedIn()){liff.login();return;}');
  p('    profile=await liff.getProfile();');
  p('    ready=true;');
  p('    document.getElementById("hdUser").textContent="👤 "+profile.displayName;');
  p('    loadFavs();');
  p('    enableBtns();');
  p('    document.getElementById("fabCart").classList.remove("fab-hidden");');
  p('    document.getElementById("fabOrders").classList.remove("fab-hidden");');
  p('    checkStatus();');
  p('  }catch(e){');
  p('    document.getElementById("hdUser").textContent="LIFF 初始化失敗："+e.message;');
  p('    console.error("LIFF error",e);');
  p('  }');
  p('}');

  /* checkStatus */
  p('async function checkStatus(){');
  p('  try{');
  p('    var r=await fetch("/api/status");');
  p('    var d=await r.json();');
  p('    var el=document.getElementById("hdBadge");');
  p('    if(d.isOpen){el.textContent="開放點餐";el.className="hd-badge open";}');
  p('    else{el.textContent="未開放";el.className="hd-badge closed";}');
  p('    if(d.autoCloseAt)startCd(d.autoCloseAt);');
  p('    else document.getElementById("hdCd").textContent="";');
  p('  }catch(e){}');
  p('}');

  /* startCd */
  p('function startCd(iso){');
  p('  var el=document.getElementById("hdCd");');
  p('  function tick(){');
  p('    var diff=new Date(iso)-new Date();');
  p('    if(diff<=0){el.textContent="";return;}');
  p('    var m=Math.floor(diff/60000),s=Math.floor((diff%60000)/1000);');
  p('    el.textContent="⏰ 自動結單："+m+"分"+s+"秒後";');
  p('    setTimeout(tick,1000);');
  p('  }');
  p('  tick();');
  p('}');

  /* loadFavs / saveFavs */
  p('function loadFavs(){');
  p('  try{');
  p('    var k="fav_"+(profile?profile.userId:"");');
  p('    var v=localStorage.getItem(k);');
  p('    favs=v?JSON.parse(v):[];');
  p('  }catch(e){favs=[];}');
  p('}');
  p('function saveFavs(){');
  p('  try{localStorage.setItem("fav_"+(profile?profile.userId:""),JSON.stringify(favs));}catch(e){}');
  p('}');
  p('function toggleFav(store){');
  p('  var i=favs.indexOf(store);');
  p('  if(i>=0)favs.splice(i,1);else favs.push(store);');
  p('  saveFavs();');
  p('  renderMenu();');
  p('}');

  /* renderMenu */
  p('function renderMenu(){');
  p('  var lm=document.getElementById("lm");');
  p('  if(lm)lm.style.display="none";');
  p('  var menuEl=document.getElementById("menuEl");');
  p('  var tabsEl=document.getElementById("tabsEl");');
  p('  if(!menu.length){menuEl.innerHTML=\'<div class="empty">目前沒有菜單資料</div>\';tabsEl.innerHTML="";return;}');
  /* 搜尋篩選 */
  p('  var q=searchQ.toLowerCase();');
  p('  var filtered=q?menu.filter(function(m){return m.item.toLowerCase().indexOf(q)>=0||m.store.toLowerCase().indexOf(q)>=0;}):menu;');
  /* 店家排序：收藏優先 */
  p('  var allStores=[];');
  p('  menu.forEach(function(m){if(allStores.indexOf(m.store)<0)allStores.push(m.store);});');
  p('  var favOrd=favs.filter(function(s){return allStores.indexOf(s)>=0;});');
  p('  var rest=allStores.filter(function(s){return favs.indexOf(s)<0;});');
  p('  var stores=favOrd.concat(rest);');
  /* tabs */
  p('  tabsEl.innerHTML=stores.map(function(s,i){');
  p('    var isFav=favs.indexOf(s)>=0;');
  p('    return \'<div class="tab\'+(isFav?" starred":"")+\'" id="tab\'+i+\'" onclick="goStore(\'+i+\')">\'+escH(s)+\'</div>\';');
  p('  }).join("");');
  /* cards */
  p('  var html="";');
  p('  stores.forEach(function(store,si){');
  p('    var items=filtered.filter(function(m){return m.store===store;});');
  p('    if(!items.length)return;');
  p('    var isFav=favs.indexOf(store)>=0;');
  // 重點：fav-btn 的 onclick 用 data-store attribute，完全避免店家名稱的 escape 問題
  p('    html+=\'<div class="store-row" id="store\'+si+\'">\';');
  p('    html+=\'<span class="store-name">\'+escH(store)+\'</span>\';');
  p('    html+=\'<button class="fav-btn" data-store="\'+escH(store)+\'" onclick="handleFavBtn(this)">\'+( isFav?"⭐":"☆")+"</button>";');
  p('    html+=\'</div>\';');
  p('    items.forEach(function(m){');
  p('      var idx=menu.indexOf(m);');
  p('      html+=\'<div class="card">\';');
  p('      if(m.image){html+=\'<img class="card-img" src="\'+escH(m.image)+\'" alt="\'+escH(m.item)+\'" loading="lazy" onerror="this.remove()">\'; }');
  p('      html+=\'<div class="card-body">\';');
  p('      html+=\'<div class="card-store">\'+escH(m.store)+\'</div>\';');
  p('      html+=\'<div class="card-name">\'+escH(m.item)+\'</div>\';');
  p('      html+=\'<div class="card-price">$\'+Number(m.price)+\'</div>\';');
  p('      html+=\'<button class="add-btn" id="btn\'+idx+\'" onclick="addToCart(\'+idx+\')" disabled>載入中...</button>\';');
  p('      html+=\'</div></div>\';');
  p('    });');
  p('  });');
  p('  menuEl.innerHTML=html||\'<div class="empty">找不到相符的品項</div>\';');
  p('  if(ready)enableBtns();');
  p('}');

  /* handleFavBtn — 從 data-store 取店家名，完全避免 escape 問題 */
  p('function handleFavBtn(btn){');
  p('  var store=btn.getAttribute("data-store");');
  p('  if(store)toggleFav(store);');
  p('}');

  p('function goStore(i){');
  p('  document.querySelectorAll(".tab").forEach(function(b,j){b.classList.toggle("active",j===i);});');
  p('  var el=document.getElementById("store"+i);');
  p('  if(el)el.scrollIntoView({behavior:"smooth",block:"start"});');
  p('}');

  p('function onSearch(){searchQ=document.getElementById("srchInput").value.trim();renderMenu();}');

  p('function enableBtns(){');
  p('  menu.forEach(function(_,i){');
  p('    var b=document.getElementById("btn"+i);');
  p('    if(b){b.disabled=false;b.textContent="加入購物車";}');
  p('  });');
  p('}');

  /* addToCart */
  p('function addToCart(idx){');
  p('  if(!ready||!profile){alert("尚未取得 LINE 使用者資料");return;}');
  p('  curItem=menu[idx];');
  p('  var key=curItem.store+"||"+curItem.item;');
  p('  curGroups=opts[key]||[];');
  p('  curQty=1;');
  p('  document.getElementById("moOptsTitle").textContent=curItem.item;');
  p('  document.getElementById("qtyVal").textContent="1";');
  p('  document.getElementById("itemNote").value="";');
  p('  var body=document.getElementById("moOptsBody");');
  p('  body.innerHTML="";');
  p('  if(!curGroups.length){');
  p('    body.innerHTML=\'<p style="color:var(--sub);margin:8px 0;font-size:14px">此商品無需選擇規格。</p>\';');
  p('  }else{');
  p('    curGroups.forEach(function(g,gi){');
  p('      var wrap=document.createElement("div");');
  p('      var title=document.createElement("div");');
  p('      title.className="opt-g-title";');
  p('      title.textContent=g.category+" ("+g.min+"~"+g.max+"選)";');
  p('      wrap.appendChild(title);');
  p('      var chips=document.createElement("div");');
  p('      chips.className="chips";');
  p('      chips.id="chips"+gi;');
  p('      g.options.forEach(function(opt){');
  p('        var c=document.createElement("div");');
  p('        c.className="chip";');
  p('        c.textContent=opt;');
  p('        c.addEventListener("click",function(){toggleChip(gi,g.max,c);});');
  p('        chips.appendChild(c);');
  p('      });');
  p('      wrap.appendChild(chips);');
  p('      body.appendChild(wrap);');
  p('    });');
  p('  }');
  p('  openModal("moOpts");');
  p('}');

  p('function chQty(d){curQty=Math.max(1,Math.min(20,curQty+d));document.getElementById("qtyVal").textContent=curQty;}');

  p('function toggleChip(gi,max,chip){');
  p('  var chips=[].slice.call(document.querySelectorAll("#chips"+gi+" .chip"));');
  p('  var sel=chips.filter(function(c){return c.classList.contains("on");});');
  p('  if(chip.classList.contains("on")){chip.classList.remove("on");}');
  p('  else if(sel.length<max){chip.classList.add("on");}');
  p('  var sel2=chips.filter(function(c){return c.classList.contains("on");});');
  p('  chips.forEach(function(c){');
  p('    if(sel2.length>=max&&!c.classList.contains("on"))c.classList.add("off");');
  p('    else c.classList.remove("off");');
  p('  });');
  p('}');

  p('function submitOpts(){');
  p('  var specParts=[];');
  p('  for(var i=0;i<curGroups.length;i++){');
  p('    var g=curGroups[i];');
  p('    var sel=[].slice.call(document.querySelectorAll("#chips"+i+" .chip.on"));');
  p('    if(sel.length<g.min||sel.length>g.max){alert(g.category+" 需要選 "+g.min+"~"+g.max+" 個");return;}');
  p('    if(sel.length)specParts.push(g.category+"："+sel.map(function(c){return c.textContent;}).join("、"));');
  p('  }');
  p('  cart.push({store:curItem.store,item:curItem.item,spec:specParts.join(" "),note:document.getElementById("itemNote").value.trim(),qty:curQty,price:curItem.price});');
  p('  closeModal("moOpts");');
  p('  updBadge();');
  p('  showToast("已加入購物車 🎉");');
  p('}');

  /* cart */
  p('function updBadge(){document.getElementById("cartBadge").textContent=cart.reduce(function(a,c){return a+c.qty;},0);}');

  p('function renderCart(){');
  p('  var el=document.getElementById("cartList");');
  p('  var te=document.getElementById("cartTotal");');
  p('  if(!cart.length){el.innerHTML=\'<div class="empty" style="padding:20px 0">購物車是空的</div>\';te.textContent="";return;}');
  p('  el.innerHTML=cart.map(function(c,i){');
  p('    var sub=[c.spec,c.note].filter(Boolean).join("｜");');
  p('    return \'<div class="orow">\'');
  p('      +\'<div class="oinfo"><div class="oname">\'+escH(c.item)+\' x\'+c.qty+\'</div>\'');
  p('      +(sub?\'<div class="osub">\'+escH(sub)+\'</div>\':"")');
  p('      +\'</div>\'');
  p('      +\'<div class="oact"><div class="oprice">$\'+(c.price*c.qty)+\'</div>\'');
  p('      +\'<button class="btn-del" onclick="removeCart(\'+i+\')">移除</button></div></div>\';');
  p('  }).join("");');
  p('  te.textContent="合計：$"+cart.reduce(function(a,c){return a+c.price*c.qty;},0);');
  p('}');

  p('function removeCart(i){cart.splice(i,1);updBadge();renderCart();}');

  p('async function submitCart(){');
  p('  if(!cart.length){alert("購物車是空的");return;}');
  p('  var btn=document.getElementById("submitBtn");');
  p('  btn.disabled=true;btn.textContent="送出中...";');
  p('  for(var i=0;i<cart.length;i++){');
  p('    var c=cart[i];');
  p('    try{');
  p('      var r=await fetch("/api/order",{method:"POST",headers:{"Content-Type":"application/json"},');
  p('        body:JSON.stringify({store:c.store,item:c.item,spec:c.spec,note:c.note,qty:c.qty,price:c.price,name:profile.displayName,userId:profile.userId})});');
  p('      var d=await r.json();');
  p('      if(!d.success&&d.reason==="duplicate")showToast("⚠️ "+c.item+" 已送出，略過");');
  p('    }catch(e){console.error("order err",e);}');
  p('  }');
  p('  cart=[];updBadge();closeModal("moCart");');
  p('  btn.disabled=false;btn.textContent="送出訂單";');
  p('  showToast("訂單已送出 ✅");');
  p('}');

  /* 我的訂單 */
  p('async function loadMyOrders(dateStr){');
  p('  curDate=dateStr||null;');
  p('  var url="/api/my-orders?userId="+encodeURIComponent(profile.userId);');
  p('  if(curDate)url+="&date="+encodeURIComponent(curDate);');
  p('  try{var r=await fetch(url);renderOrders(await r.json());}catch(e){}');
  p('}');

  p('async function loadDates(){');
  p('  if(!profile)return;');
  p('  try{');
  p('    var r=await fetch("/api/my-dates?userId="+encodeURIComponent(profile.userId));');
  p('    orderDates=await r.json();');
  p('    renderDates();');
  p('  }catch(e){}');
  p('}');

  /* renderDates — 用 DOM API，完全避免閉包和字串 escape 問題 */
  p('function renderDates(){');
  p('  var el=document.getElementById("dateChips");');
  p('  el.innerHTML="";');
  p('  if(!orderDates.length)return;');
  p('  orderDates.forEach(function(d){');
  p('    var div=document.createElement("div");');
  p('    div.className="dchip"+(d===curDate?" on":"");');
  p('    div.textContent=d;');
  p('    div.addEventListener("click",(function(dd){return function(){loadMyOrders(dd);renderDates();};})(d));');
  p('    el.appendChild(div);');
  p('  });');
  p('}');

  p('function renderOrders(orders){');
  p('  var el=document.getElementById("orderList");');
  p('  var te=document.getElementById("orderTotal");');
  p('  if(!orders||!orders.length){el.innerHTML=\'<div class="empty" style="padding:20px 0">這天沒有訂單</div>\';te.textContent="";return;}');
  p('  el.innerHTML=orders.map(function(o){');
  p('    var sub=[o.spec,o.note].filter(Boolean).join("｜");');
  p('    var tag=o.status==="已付款"?\'<span class="tag-paid">已付款</span>\':\'<span class="tag-unpaid">未付款</span>\';');
  p('    var canEdit=o.status!=="已付款";');
  p('    return \'<div class="orow">\'');
  p('      +\'<div class="oinfo">\'');
  p('      +\'<div class="oname">\'+escH(o.item)+\' x\'+o.qty+\' \'+tag+\'</div>\'');
  p('      +(sub?\'<div class="osub">\'+escH(sub)+\'</div>\':"")');
  p('      +\'<div class="osub">\'+escH(o.store)+\'</div>\'');
  p('      +\'<button class="btn-reorder" onclick="reorder(\'+o.rowIndex+\')">🔄 再訂一次</button>\'');
  p('      +\'</div>\'');
  p('      +\'<div class="oact">\'');
  p('      +\'<div class="oprice">$\'+o.total+\'</div>\'');
  p('      +(canEdit?\'<button class="icon-btn" onclick="openEditNote(\'+o.rowIndex+\',\'+JSON.stringify(o.note||"")+\')">✏️</button>\':""  )');
  p('      +(canEdit?\'<button class="icon-btn" onclick="delOrder(\'+o.rowIndex+\')">🗑</button>\':"")');
  p('      +\'</div></div>\';');
  p('  }).join("");');
  p('  te.textContent="合計：$"+orders.reduce(function(a,o){return a+o.total;},0);');
  p('}');

  /* reorder — 用 rowIndex 找回訂單資訊不夠，改存快取 */
  p('var ordersCache=[];');
  p('function renderOrders(orders){');
  p('  ordersCache=orders||[];');
  p('  var el=document.getElementById("orderList");');
  p('  var te=document.getElementById("orderTotal");');
  p('  if(!ordersCache.length){el.innerHTML=\'<div class="empty" style="padding:20px 0">這天沒有訂單</div>\';te.textContent="";return;}');
  p('  el.innerHTML=ordersCache.map(function(o,oi){');
  p('    var sub=[o.spec,o.note].filter(Boolean).join("｜");');
  p('    var tag=o.status==="已付款"?\'<span class="tag-paid">已付款</span>\':\'<span class="tag-unpaid">未付款</span>\';');
  p('    var canEdit=o.status!=="已付款";');
  p('    return \'<div class="orow">\'');
  p('      +\'<div class="oinfo">\'');
  p('      +\'<div class="oname">\'+escH(o.item)+\' x\'+o.qty+\' \'+tag+\'</div>\'');
  p('      +(sub?\'<div class="osub">\'+escH(sub)+\'</div>\':"")');
  p('      +\'<div class="osub">\'+escH(o.store)+\'</div>\'');
  p('      +\'<button class="btn-reorder" onclick="reorder(\'+oi+\')">🔄 再訂一次</button>\'');
  p('      +\'</div>\'');
  p('      +\'<div class="oact">\'');
  p('      +\'<div class="oprice">$\'+o.total+\'</div>\'');
  p('      +(canEdit?\'<button class="icon-btn" onclick="openEditNote(\'+o.rowIndex+\',\'+JSON.stringify(o.note||"")+\')">✏️</button>\':"")');
  p('      +(canEdit?\'<button class="icon-btn" onclick="delOrder(\'+o.rowIndex+\')">🗑</button>\':"")');
  p('      +\'</div></div>\';');
  p('  }).join("");');
  p('  te.textContent="合計：$"+ordersCache.reduce(function(a,o){return a+o.total;},0);');
  p('}');

  /* reorder */
  p('function reorder(oi){');
  p('  var o=ordersCache[oi];');
  p('  if(!o)return;');
  p('  var idx=menu.findIndex(function(m){return m.store===o.store&&m.item===o.item;});');
  p('  if(idx<0){showToast("此品項已不在菜單中");return;}');
  p('  cart.push({store:o.store,item:o.item,spec:o.spec||"",note:o.note||"",qty:o.qty||1,price:menu[idx].price});');
  p('  updBadge();');
  p('  showToast("已加入購物車 🔄");');
  p('}');

  /* delOrder */
  p('async function delOrder(rowIndex){');
  p('  if(!confirm("確定要刪除這筆訂單？"))return;');
  p('  var r=await fetch("/api/order/"+rowIndex,{method:"DELETE",headers:{"Content-Type":"application/json"},body:JSON.stringify({userId:profile.userId})});');
  p('  var d=await r.json();');
  p('  if(d.success){showToast("已刪除");loadMyOrders(curDate);}else alert("刪除失敗");');
  p('}');

  /* editNote */
  p('function openEditNote(rowIndex,cur){editRI=rowIndex;document.getElementById("noteInput").value=cur||"";openModal("moNote");}');
  p('async function saveNote(){');
  p('  var note=document.getElementById("noteInput").value.trim();');
  p('  var r=await fetch("/api/order/"+editRI+"/note",{method:"PATCH",headers:{"Content-Type":"application/json"},body:JSON.stringify({userId:profile.userId,note:note})});');
  p('  var d=await r.json();');
  p('  if(d.success){showToast("備註已更新");closeModal("moNote");loadMyOrders(curDate);}else alert("更新失敗");');
  p('}');

  /* modal */
  p('function openModal(id){');
  p('  document.getElementById(id).classList.add("show");');
  p('  if(id==="moCart")renderCart();');
  p('  if(id==="moOrders"){loadDates();loadMyOrders(curDate);}');
  p('}');
  p('function closeModal(id){document.getElementById(id).classList.remove("show");}');
  /* 點背景關閉 */
  p('document.querySelectorAll(".modal").forEach(function(m){');
  p('  m.addEventListener("click",function(e){if(e.target===m)m.classList.remove("show");});');
  p('});');

  /* toast */
  p('var _toastTimer=null;');
  p('function showToast(msg){');
  p('  var t=document.getElementById("toast");');
  p('  t.textContent=msg;t.style.opacity="1";');
  p('  clearTimeout(_toastTimer);');
  p('  _toastTimer=setTimeout(function(){t.style.opacity="0";},2500);');
  p('}');

  /* escH */
  p('function escH(s){return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");}');

  /* init */
  p('loadMenuData();');
  p('initLIFF();');
  p('setInterval(checkStatus,30000);');

  p('</script>');
  p('</body></html>');

  return L.join('\n');
}

// ════════════════════════════════════════════════════════════════
//  管理後台
// ════════════════════════════════════════════════════════════════
app.get('/admin', (req, res) => {
  const token = req.query.token || '';
  if (!process.env.ADMIN_TOKEN || token !== process.env.ADMIN_TOKEN)
    return res.status(401).send('Unauthorized. 請附上 ?token=YOUR_ADMIN_TOKEN');
  res.setHeader('Content-Type','text/html; charset=utf-8');
  res.end(buildAdminPage());
});

function buildAdminPage() {
  const L = [];
  const p = s => L.push(s);

  p('<!DOCTYPE html><html lang="zh-TW"><head>');
  p('<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">');
  p('<title>訂餐管理後台</title>');
  p('<style>');
  p(':root{--g:#06c755;--r:#e53935;--o:#f57c00;--bg:#f4f4f8;--card:#fff;--bdr:#e8e8ef;--txt:#1a1a2e;--sub:#888}');
  p('*{box-sizing:border-box;margin:0;padding:0}');
  p('body{font-family:-apple-system,Arial,"Microsoft JhengHei",sans-serif;background:var(--bg);color:var(--txt)}');
  p('.hd{background:#1a1a2e;color:#fff;padding:14px 20px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:50}');
  p('.hd h1{font-size:18px;font-weight:800}');
  p('.hd-time{font-size:12px;opacity:.6}');
  p('.wrap{padding:14px;max-width:1000px;margin:0 auto}');
  p('.card{background:var(--card);border-radius:14px;padding:18px;margin-bottom:14px;box-shadow:0 2px 12px rgba(0,0,0,.07)}');
  p('.card h3{font-size:15px;font-weight:800;margin-bottom:14px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:6px}');
  p('.stats{display:grid;grid-template-columns:repeat(auto-fill,minmax(130px,1fr));gap:8px;margin-bottom:4px}');
  p('.sbox{background:var(--bg);border-radius:10px;padding:12px 14px}');
  p('.snum{font-size:24px;font-weight:800;margin-bottom:2px}');
  p('.slbl{font-size:11px;color:var(--sub)}');
  p('.sbox.g .snum{color:var(--g)}.sbox.r .snum{color:var(--r)}.sbox.o .snum{color:var(--o)}');
  p('.crow{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:10px}');
  p('.sdot{width:10px;height:10px;border-radius:50%;background:var(--sub);flex-shrink:0}');
  p('.sdot.on{background:var(--g)}');
  p('.stxt{font-size:14px;font-weight:700}');
  p('.ac-info{font-size:12px;color:var(--r);font-weight:700}');
  p('.btn{padding:8px 14px;border:none;border-radius:999px;font-size:13px;font-weight:700;cursor:pointer;white-space:nowrap}');
  p('.btn.g{background:var(--g);color:#fff}.btn.r{background:var(--r);color:#fff}.btn.o{background:var(--o);color:#fff}');
  p('.btn.out{background:var(--card);border:1.5px solid var(--bdr);color:var(--txt)}.btn.sm{padding:5px 10px;font-size:12px}');
  p('.btn:hover{opacity:.85}');
  p('input[type=number],input[type=date],select.sel{padding:7px 10px;border-radius:8px;border:1.5px solid var(--bdr);font-size:13px;background:var(--card)}');
  p('input[type=number]{width:64px}');
  p('.srch{width:100%;padding:9px 14px;border-radius:999px;border:1.5px solid var(--bdr);font-size:13px;margin-bottom:10px;background:var(--card)}');
  p('.tbl-w{overflow-x:auto}');
  p('table{width:100%;border-collapse:collapse;font-size:13px}');
  p('th{padding:9px 8px;border-bottom:2px solid var(--bdr);text-align:left;font-weight:700;color:var(--sub);white-space:nowrap;background:var(--bg)}');
  p('td{padding:8px;border-bottom:1px solid var(--bdr);vertical-align:middle}');
  p('tr:hover td{background:#f8f8fc}');
  p('.sbadge{display:inline-block;padding:2px 7px;border-radius:6px;font-size:11px;font-weight:700;background:#e3f2fd;color:#1565c0}');
  p('.tag-p{display:inline-block;padding:2px 9px;border-radius:999px;font-size:11px;font-weight:700;background:#e8f5e9;color:#2e7d32}');
  p('.tag-u{display:inline-block;padding:2px 9px;border-radius:999px;font-size:11px;font-weight:700;background:#fff3e0;color:#e65100}');
  p('.tag-d{display:inline-block;padding:2px 9px;border-radius:999px;font-size:11px;font-weight:700;background:#fce4ec;color:#b71c1c}');
  p('.gsec{margin-bottom:18px}');
  p('.gttl{font-size:14px;font-weight:800;padding:8px 10px;background:var(--bg);border-radius:8px;margin-bottom:6px;display:flex;align-items:center;justify-content:space-between}');
  p('.gsub{font-size:12px;color:var(--sub)}');
  p('@media(max-width:600px){th,td{padding:6px 5px;font-size:12px}.stats{grid-template-columns:repeat(2,1fr)}}');
  p('</style></head><body>');

  p('<div class="hd"><h1>🍱 訂餐管理後台</h1><span class="hd-time" id="hdT"></span></div>');
  p('<div class="wrap">');

  /* 開單控制 */
  p('<div class="card"><h3>📢 開單控制</h3>');
  p('<div class="crow">');
  p('  <span class="sdot" id="sDot"></span><span class="stxt" id="sTxt"></span>');
  p('  <button class="btn g" onclick="doOpen()">開單</button>');
  p('  <button class="btn r" onclick="doClose()">結單</button>');
  p('  <button class="btn o" onclick="doClear()">清空</button>');
  p('</div>');
  p('<div class="crow" style="margin-top:4px">');
  p('  <span style="font-size:13px;font-weight:600">自動結單：</span>');
  p('  <input type="number" id="acMin" value="30" min="1" max="480"> 分鐘');
  p('  <button class="btn out" onclick="doAc()">設定</button>');
  p('  <button class="btn out" onclick="doCancelAc()">取消</button>');
  p('  <span class="ac-info" id="acInfo"></span>');
  p('</div></div>');

  /* 統計 */
  p('<div class="card"><h3>📊 今日統計</h3><div class="stats" id="statsEl">載入中...</div></div>');

  /* 訂單列表 */
  p('<div class="card"><h3>📋 訂單列表');
  p('  <div style="display:flex;gap:6px;flex-wrap:wrap">');
  p('    <button class="btn g sm" onclick="doBatchPaid()">✅ 全標已付</button>');
  p('    <button class="btn out sm" onclick="loadOrders()">🔄 重整</button>');
  p('    <button class="btn out sm" onclick="copyShop()">📋 複製店家單</button>');
  p('  </div>');
  p('</h3>');
  p('<div class="crow">');
  p('  <span style="font-size:13px;font-weight:600">日期：</span>');
  p('  <input type="date" id="dateF" onchange="loadOrders()">');
  p('  <button class="btn out sm" onclick="document.getElementById(\'dateF\').value=\'\';loadOrders()">今日</button>');
  p('  <span style="font-size:13px;font-weight:600;margin-left:8px">分組：</span>');
  p('  <select class="sel" id="gbSel" onchange="renderTable()">');
  p('    <option value="none">不分組</option>');
  p('    <option value="store" selected>依店家</option>');
  p('    <option value="person">依人員</option>');
  p('  </select>');
  p('</div>');
  p('<input class="srch" id="srch" placeholder="搜尋姓名、品項..." oninput="renderTable()">');
  p('<div id="ordersEl"></div>');
  p('</div>');

  p('</div>'); /* .wrap */

  /* Admin JS */
  p('<script>');
  p('var TOKEN=(location.search.match(/[?&]token=([^&]*)/)||[])[1]||"";');
  p('var cache=[],acAt=null,acTimer=null;');
  p('document.getElementById("hdT").textContent=new Date().toLocaleString("zh-TW");');

  p('async function api(path,method,body){');
  p('  var o={method:method||"GET",headers:{"Content-Type":"application/json","x-admin-token":TOKEN}};');
  p('  if(body!==undefined)o.body=JSON.stringify(body);');
  p('  return (await fetch(path,o)).json();');
  p('}');

  p('async function loadStatus(){');
  p('  var d=await api("/api/status");');
  p('  document.getElementById("sDot").className="sdot"+(d.isOpen?" on":"");');
  p('  document.getElementById("sTxt").textContent=d.isOpen?"開單中":"已結單";');
  p('  if(d.autoCloseAt){acAt=d.autoCloseAt;startAcCd();}');
  p('  else{acAt=null;document.getElementById("acInfo").textContent="";}');
  p('}');

  p('function startAcCd(){');
  p('  if(acTimer)clearInterval(acTimer);');
  p('  acTimer=setInterval(function(){');
  p('    if(!acAt){clearInterval(acTimer);return;}');
  p('    var diff=new Date(acAt)-new Date();');
  p('    if(diff<=0){document.getElementById("acInfo").textContent="已自動結單";clearInterval(acTimer);loadStatus();return;}');
  p('    document.getElementById("acInfo").textContent="⏰ "+Math.floor(diff/60000)+"分"+Math.floor((diff%60000)/1000)+"秒";');
  p('  },1000);');
  p('}');

  p('async function doOpen(){await api("/api/admin/open","POST",{});loadStatus();loadOrders();}');
  p('async function doClose(){await api("/api/admin/close","POST",{});loadStatus();loadOrders();}');
  p('async function doClear(){if(!confirm("確定清空狀態？"))return;await api("/api/admin/clear","POST",{});loadStatus();loadOrders();}');
  p('async function doAc(){var m=Number(document.getElementById("acMin").value)||30;await api("/api/admin/auto-close","POST",{minutes:m});loadStatus();}');
  p('async function doCancelAc(){await api("/api/admin/cancel-auto-close","POST",{});acAt=null;document.getElementById("acInfo").textContent="";clearInterval(acTimer);}');

  p('async function loadOrders(){');
  p('  var date=document.getElementById("dateF").value;');
  p('  var url="/api/admin/orders"+(date?"?date="+date:"");');
  p('  var d=await api(url);');
  p('  cache=Array.isArray(d)?d:[];');
  p('  renderStats(cache);');
  p('  renderTable();');
  p('}');

  p('function renderStats(o){');
  p('  var act=o.filter(function(x){return x.status!=="已刪除";});');
  p('  var paid=act.filter(function(x){return x.status==="已付款";});');
  p('  var unpaid=act.filter(function(x){return x.status==="未付款";});');
  p('  var tot=act.reduce(function(a,x){return a+x.total;},0);');
  p('  var pM=paid.reduce(function(a,x){return a+x.total;},0);');
  p('  var uM=unpaid.reduce(function(a,x){return a+x.total;},0);');
  // mkSb 定義在下方
  p('  document.getElementById("statsEl").innerHTML=');
  p('    mkSb(act.length,"筆訂單","")+mkSb("$"+tot,"總金額","g")+mkSb("$"+pM,"已收款","g")+mkSb("$"+uM,"未收款","r")+mkSb(paid.length,"已付款","g")+mkSb(unpaid.length,"未付款","o");');
  p('}');
  p('function mkSb(n,l,cls){');
  p('  return \'<div class="sbox\'+(cls?" "+cls:"")+\'"><div class="snum">\'+n+\'</div><div class="slbl">\'+l+\'</div></div>\';');
  p('}');

  p('function renderTable(){');
  p('  var q=(document.getElementById("srch").value||"").toLowerCase();');
  p('  var gb=document.getElementById("gbSel").value;');
  p('  var data=cache.filter(function(o){return !q||(o.name||"").toLowerCase().indexOf(q)>=0||(o.item||"").toLowerCase().indexOf(q)>=0;});');
  p('  var el=document.getElementById("ordersEl");');
  p('  if(!data.length){el.innerHTML=\'<div style="text-align:center;padding:24px;color:#aaa">無訂單</div>\';return;}');
  p('  if(gb==="none"){el.innerHTML=mkTable(data);return;}');
  p('  var groups={};');
  p('  data.forEach(function(o){var k=gb==="store"?(o.store||"其他"):(o.name||"未知");if(!groups[k])groups[k]=[];groups[k].push(o);});');
  p('  var html="";');
  p('  Object.keys(groups).forEach(function(gk){');
  p('    var items=groups[gk];');
  p('    var sub=items.filter(function(o){return o.status!=="已刪除";}).reduce(function(a,o){return a+o.total;},0);');
  p('    html+=\'<div class="gsec"><div class="gttl"><span>\'+esc(gk)+\'</span><span class="gsub">小計 $\'+sub+\' / \'+items.length+\' 筆</span></div>\'+mkTable(items)+\'</div>\';');
  p('  });');
  p('  el.innerHTML=html;');
  p('}');

  p('function mkTable(rows){');
  p('  if(!rows.length)return "";');
  p('  var h=\'<div class="tbl-w"><table><thead><tr>\'');
  p('    +\'<th>姓名</th><th>店家</th><th>品項</th><th>規格</th><th>備註</th><th>數量</th><th>金額</th><th>狀態</th><th>操作</th>\'');
  p('    +\'</tr></thead><tbody>\';');
  p('  h+=rows.map(function(o){');
  p('    var tc=o.status==="已付款"?"tag-p":o.status==="已刪除"?"tag-d":"tag-u";');
  p('    var ops=o.status==="未付款"');
  p('      ?\'<select class="sel" id="pt\'+o.rowIndex+\'" style="padding:4px 7px;font-size:11px">\'');
  p('        +\'<option>現金</option><option>Line Pay</option><option>轉帳</option></select> \'');
  p('        +\'<button class="btn g sm" onclick="doPaid(\'+o.rowIndex+\')">付款</button> \'');
  p('        +\'<button class="btn r sm" onclick="doDelOrder(\'+o.rowIndex+\')">刪</button>\'');
  p('      :"";');
  p('    return \'<tr>\'');
  p('      +\'<td>\'+esc(o.name)+\'</td><td><span class="sbadge">\'+esc(o.store)+\'</span></td>\'');
  p('      +\'<td>\'+esc(o.item)+\'</td><td style="color:var(--sub)">\'+esc(o.spec)+\'</td>\'');
  p('      +\'<td style="color:var(--sub)">\'+esc(o.note)+\'</td><td>\'+o.qty+\'</td>\'');
  p('      +\'<td style="font-weight:800">$\'+o.total+\'</td>\'');
  p('      +\'<td><span class="\'+tc+\'">\'+esc(o.status)+\'</span></td>\'');
  p('      +\'<td>\'+ops+\'</td></tr>\';');
  p('  }).join("");');
  p('  h+="</tbody></table></div>";');
  p('  return h;');
  p('}');

  p('async function doPaid(ri){var sel=document.getElementById("pt"+ri);await api("/api/admin/paid","POST",{rowIndex:ri,payType:sel?sel.value:"現金"});loadOrders();}');
  p('async function doDelOrder(ri){if(!confirm("確定刪除？"))return;await api("/api/admin/delete-order","POST",{rowIndex:ri});loadOrders();}');
  p('async function doBatchPaid(){if(!confirm("全部標記已付款？"))return;await api("/api/admin/batch-paid","POST",{payType:"現金"});loadOrders();}');

  p('function copyShop(){');
  p('  var act=cache.filter(function(o){return o.status!=="已刪除";});');
  p('  var cnt={};');
  p('  act.forEach(function(o){var k=o.item+(o.spec?"（"+o.spec+"）":"");cnt[k]=(cnt[k]||0)+o.qty;});');
  p('  var txt="您好，今天訂購如下：\n\n",n=0;');
  p('  Object.keys(cnt).forEach(function(k){txt+=k+" x"+cnt[k]+"\n";n+=cnt[k];});');
  p('  var m=act.reduce(function(a,o){return a+o.total;},0);');
  p('  txt+="\n總數："+n+"份\n總金額："+m+"元\n\n麻煩您，謝謝～";');
  p('  navigator.clipboard.writeText(txt).then(function(){alert("已複製");});');
  p('}');

  p('function esc(s){return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");}');

  p('loadStatus();loadOrders();');
  p('setInterval(loadStatus,30000);setInterval(loadOrders,60000);');
  p('</script></body></html>');

  return L.join('\n');
}

// ════════════════════════════════════════════════════════════════
//  LINE Webhook
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
    } catch(e) { console.error('取得名稱失敗:', e.message); }

    knownUsers[profileName] = event.source.userId;
    saveUserToSheet(profileName, event.source.userId, event.source.type, event.source.groupId||'').catch(()=>{});

    const uid   = event.source.userId;
    const text  = event.message.text.trim();
    const reply = t => client.replyMessage(event.replyToken, { type:'text', text:t });

    const autoOpen = text.match(/^開單\s+(\d+)$/);
    if (autoOpen) {
      if (!isAdmin(uid)) { await reply('只有管理員可以開單'); return res.sendStatus(200); }
      isOpen = true;
      scheduleAutoClose(Number(autoOpen[1]));
      await reply('已開單！將於 '+autoOpen[1]+' 分鐘後自動結單\n點餐頁：'+(process.env.LIFF_URL||''));
      return res.sendStatus(200);
    }
    if (text === '開單') {
      if (!isAdmin(uid)) { await reply('只有管理員可以開單'); return res.sendStatus(200); }
      if (isOpen) { await reply('目前已開單中'); return res.sendStatus(200); }
      isOpen = true;
      await reply('已開單，可以開始點餐 🍱\n點餐頁：'+(process.env.LIFF_URL||''));
      return res.sendStatus(200);
    }
    if (text === '結單' || text === '收單' || text === '統計') {
      if (!isAdmin(uid)) { await reply('只有管理員可以結單/統計'); return res.sendStatus(200); }
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
      await reply('管理後台：'+(process.env.APP_URL||'')+'/admin?token='+(process.env.ADMIN_TOKEN||''));
      return res.sendStatus(200);
    }
    if (text === '狀態') {
      const msg = (isOpen?'🟢 目前開單中':'🔴 目前未開單')+(autoCloseAt?'\n⏰ 自動結單：'+new Date(autoCloseAt).toLocaleTimeString('zh-TW',{timeZone:'Asia/Taipei'}):'');
      await reply(msg);
      return res.sendStatus(200);
    }
    return res.sendStatus(200);
  } catch(err) {
    console.error('Webhook error:', err);
    return res.sendStatus(200);
  }
});

app.use((err,_q,res,_n) => { console.error('Global error:',err); res.sendStatus(200); });

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log('Server running on port', PORT));
