require('dotenv').config();

const express = require('express');
const line = require('@line/bot-sdk');
const { GoogleSpreadsheet } = require('google-spreadsheet');

const app = express();

const lineConfig = {
channelAccessToken: process.env.CHANNEL_ACCESS_TOKEN,
channelSecret:      process.env.CHANNEL_SECRET
};

// ★ LINE webhook 必須在 express.json() 之前
app.use('/webhook', line.middleware(lineConfig));
app.use('/api', express.json());
app.use('/admin', express.json());

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
‘U8d9c82446aa9eb90d7de001cfc7ea90f’,
‘Ubcfae64b443b9fad21bbc584e991b306’,
‘U5c44a04efc62664bd45ec80d77be7d93’,
‘Uc669eca67bf477460945f45751edd3e9’
];
function isAdmin(uid) { return admins.includes(uid); }

// ════════════════════════════════════════════════════════════════
//  工具
// ════════════════════════════════════════════════════════════════
function nowTW() {
return new Date().toLocaleString(‘zh-TW’, { timeZone: ‘Asia/Taipei’ });
}
function todayTW() {
return new Date().toLocaleDateString(‘zh-TW’, { timeZone: ‘Asia/Taipei’ });
}

// ════════════════════════════════════════════════════════════════
//  Google Sheets 認證
// ════════════════════════════════════════════════════════════════
async function authSheet() {
await doc.useServiceAccountAuth({
client_email: process.env.GOOGLE_CLIENT_EMAIL,
private_key:  process.env.GOOGLE_PRIVATE_KEY.replace(/\n/g, ‘\n’)
});
await doc.loadInfo();
}

// ════════════════════════════════════════════════════════════════
//  菜單
// ════════════════════════════════════════════════════════════════
async function loadMenu() {
await authSheet();
const sheet = doc.sheetsByTitle[‘Menu’];
if (!sheet) { console.error(‘loadMenu: Menu sheet not found’); return []; }
const rows = await sheet.getRows();
return rows
.map(r => ({
store: String(r[‘店家’]    || ‘’).trim(),
item:  String(r[‘品項’]    || ‘’).trim(),
price: parseFloat(String(r[‘價格’] || ‘0’).replace(/[^0-9.]/g, ‘’)) || 0,
image: String(r[‘圖片URL’] || ‘’).trim()
}))
.filter(r => r.store && r.item && r.price > 0);
}

async function loadOptions() {
await authSheet();
const gs = doc.sheetsByTitle[‘OptionGroups’];
const os = doc.sheetsByTitle[‘Options’];
if (!gs || !os) return {};
const gr = await gs.getRows();
const or = await os.getRows();
const result = {};
gr.forEach(g => {
const store = String(g[‘店家’] || ‘’).trim();
const item  = String(g[‘品項’] || ‘’).trim();
const cat   = String(g[‘分類’] || ‘’).trim();
if (!store || !item || !cat) return;
const key = store + ‘||’ + item;
if (!result[key]) result[key] = [];
const opts = or
.filter(o => String(o[‘店家’]||’’).trim()===store && String(o[‘品項’]||’’).trim()===item && String(o[‘分類’]||’’).trim()===cat)
.map(o => ({
name:  String(o[‘選項’] || ‘’).trim(),
extra: parseFloat(String(o[‘加價’] || ‘0’).replace(/[^0-9.]/g,’’)) || 0
}))
.filter(o => o.name);
result[key].push({ category:cat, required:String(g[‘必選’]||’’).trim()===‘TRUE’, min:Number(g[‘最少’]||0), max:Number(g[‘最多’]||0), options:opts });
});
return result;
}

// ════════════════════════════════════════════════════════════════
//  Users
// ════════════════════════════════════════════════════════════════
async function saveUserToSheet(name, userId, src, gid) {
try {
await authSheet();
const s = doc.sheetsByTitle[‘Users’];
if (!s) return;
await s.addRow({ 時間:nowTW(), LINE名稱:name, userId, 來源類型:src, 群組ID:gid||’’, 權限:isAdmin(userId)?‘admin’:‘user’ });
} catch(e) { console.error(‘Users write fail:’, e.message); }
}

// ════════════════════════════════════════════════════════════════
//  Orders CRUD
// ════════════════════════════════════════════════════════════════
// ════════════════════════════════════════════════════════════════
//  純文字訂餐模式（阿姨版）
//  不需要 LIFF，LINE 裡直接回覆數字即可
// ════════════════════════════════════════════════════════════════

// 今日菜單快取（開團時載入，結單時清除）
let textMenuCache = [];   // [{no, store, item, price}]

// 讀菜單並編號
async function loadTextMenu() {
const menu = await loadMenu();
textMenuCache = menu.map((m, i) => ({ no: i + 1, store: m.store, item: m.item, price: m.price }));
return textMenuCache;
}

// 產生菜單訊息
function buildMenuMessage(menuList) {
const emojiNum = [‘1️⃣’,‘2️⃣’,‘3️⃣’,‘4️⃣’,‘5️⃣’,‘6️⃣’,‘7️⃣’,‘8️⃣’,‘9️⃣’,‘🔟’];
let msg = ‘🍱 今日菜單\n’;
msg += ‘─────────────\n’;
menuList.forEach(m => {
const em = m.no <= 10 ? emojiNum[m.no - 1] : m.no + ‘.’;
msg += em + ’ ’ + m.item + ’　$’ + m.price + ‘\n’;
});
msg += ‘\n請直接回覆數字點餐：\n’;
msg += ‘例：\n2　　　（點2號）\n2不要菜　（加備註）\n2+3　　（點2號，3份）’;
return msg;
}

// 解析文字訂餐：“2不要菜”“3+2加辣”“1”
function parseTextOrder(text, menuList) {
// 支援格式：數字 [+數量] [備註文字]
// e.g. “2”, “2不要菜”, “3+2”, “3+2加辣”, “1+3 飯少一點”
const m = text.match(/^(\d+)(?:+(\d+))?(.*)$/);
if (!m) return null;
const no   = parseInt(m[1], 10);
const qty  = Math.min(20, Math.max(1, parseInt(m[2] || ‘1’, 10)));
const note = (m[3] || ‘’).trim();
const item = menuList.find(x => x.no === no);
if (!item) return null;
return { no, item: item.item, store: item.store, price: item.price, qty, note };
}

// 儲存或覆蓋文字訂單（同一人今日同品項覆蓋，不同品項新增）
async function saveTextOrder(name, userId, parsed) {
try {
await authSheet();
const sheet = doc.sheetsByTitle[‘Orders’];
if (!sheet) return { success: false, reason: ‘no_sheet’ };
const rows  = await sheet.getRows();
const today = todayTW();

```
// 找今日同一人同品項的未刪除訂單
const existing = rows.find(r =>
  String(r['userId'] || '') === String(userId) &&
  String(r['品項']   || '').trim() === String(parsed.item).trim() &&
  String(r['時間']   || '').startsWith(today) &&
  String(r['狀態']   || '') !== '已刪除'
);

if (existing) {
  // 覆蓋：更新數量、備註、總價
  existing['數量'] = parsed.qty;
  existing['備註'] = parsed.note;
  existing['總價'] = parsed.price * parsed.qty;
  await existing.save();
  return { success: true, action: 'updated' };
}

// 新增
await sheet.addRow({
  時間:     nowTW(),
  LINE名稱: name,
  userId:   String(userId),
  店家:     parsed.store,
  品項:     parsed.item,
  規格:     '',
  備註:     parsed.note,
  數量:     parsed.qty,
  單價:     parsed.price,
  總價:     parsed.price * parsed.qty,
  狀態:     '未付款',
  付款時間: '',
  付款方式: '',
  訂單備註: ''
});
return { success: true, action: 'created' };
```

} catch(e) {
console.error(‘saveTextOrder fail:’, e.message);
return { success: false, reason: e.message };
}
}

// 查詢某人今日訂單（給”我的訂單”指令用）
async function getMyTextOrders(userId) {
try {
await authSheet();
const sheet = doc.sheetsByTitle[‘Orders’];
if (!sheet) return [];
const rows  = await sheet.getRows();
const today = todayTW();
return rows.filter(r =>
String(r[‘userId’] || ‘’) === String(userId) &&
String(r[‘時間’]   || ‘’).startsWith(today) &&
String(r[‘狀態’]   || ‘’) !== ‘已刪除’
).map(r => ({
item:   String(r[‘品項’] || ‘’),
note:   String(r[‘備註’] || ‘’),
qty:    Number(r[‘數量’] || 1),
price:  Number(r[‘單價’] || 0),
total:  Number(r[‘總價’] || 0),
status: String(r[‘狀態’] || ‘’)
}));
} catch(e) { return []; }
}

// 標記付款（管理員輸入”XXX 已付款”）
async function markPaidByName(name) {
try {
await authSheet();
const sheet = doc.sheetsByTitle[‘Orders’];
if (!sheet) return 0;
const rows  = await sheet.getRows();
const today = todayTW();
let count = 0;
for (const r of rows) {
if (
String(r[‘LINE名稱’] || ‘’).trim() === name.trim() &&
String(r[‘時間’]     || ‘’).startsWith(today) &&
String(r[‘狀態’]     || ‘’) === ‘未付款’
) {
r[‘狀態’]    = ‘已付款’;
r[‘付款時間’] = nowTW();
r[‘付款方式’] = ‘現金’;
await r.save();
count++;
}
}
return count;
} catch(e) { return 0; }
}

// 產生文字統計報表（更詳細版，給查看訂單指令）
async function buildDetailedReport() {
const orders = await getAllOrdersByDate(todayTW());
const active = orders.filter(o => o.status !== ‘已刪除’);
if (!active.length) return ‘📋 今日尚無訂單’;

// 依人員分組
const byPerson = {};
for (const o of active) {
const n = o.name || ‘未知’;
if (!byPerson[n]) byPerson[n] = { orders: [], total: 0, paid: false };
byPerson[n].orders.push(o);
byPerson[n].total += o.total;
if (o.status === ‘已付款’) byPerson[n].paid = true;
}

let msg = ‘📋 今日訂單明細\n’;
msg += ‘─────────────\n’;
for (const name in byPerson) {
const p = byPerson[name];
const payIcon = p.orders.every(o => o.status === ‘已付款’) ? ‘✅’ : ‘💰’;
msg += payIcon + ’ ’ + name + ‘（$’ + p.total + ‘）\n’;
for (const o of p.orders) {
msg += ’  • ’ + o.item;
if (o.qty > 1) msg += ’ ×’ + o.qty;
if (o.note) msg += ‘（’ + o.note + ‘）’;
msg += ‘\n’;
}
}

const grand = active.reduce((a, o) => a + o.total, 0);
const unpaidNames = Object.keys(byPerson).filter(n =>
byPerson[n].orders.some(o => o.status === ‘未付款’)
);
msg += ‘─────────────\n’;
msg += ‘💰 總金額：$’ + grand + ‘\n’;
if (unpaidNames.length) {
msg += ‘⚠️ 未付款：’ + unpaidNames.join(’、’);
} else {
msg += ‘✅ 全部已付款’;
}
return msg;
}

// 產生匯出格式（CSV 式文字，方便複製）
async function buildExportText() {
const orders = await getAllOrdersByDate(todayTW());
const active = orders.filter(o => o.status !== ‘已刪除’);
if (!active.length) return ‘今日尚無訂單’;
let out = ‘姓名,品項,數量,備註,金額,付款狀態\n’;
for (const o of active) {
out += [o.name, o.item, o.qty, o.note, o.total, o.status].join(’,’) + ‘\n’;
}
return out;
}

async function saveOrderToSheet(order) {
try {
await authSheet();
const s = doc.sheetsByTitle[‘Orders’];
if (!s) return { success:false, reason:‘no_sheet’ };
const rows  = await s.getRows();
const today = todayTW();
const dup = rows.find(r =>
String(r[‘userId’]||’’)=== String(order.userId) &&
String(r[‘品項’]  ||’’).trim()===String(order.item||’’).trim() &&
String(r[‘規格’]  ||’’).trim()===String(order.spec||’’).trim() &&
String(r[‘時間’]  ||’’).startsWith(today) &&
String(r[‘狀態’]  ||’’)!==‘已刪除’
);
if (dup) return { success:false, reason:‘duplicate’ };
const qty=Number(order.qty||1), price=Number(order.price||0);
await s.addRow({
時間:nowTW(), LINE名稱:String(order.name||’’), userId:String(order.userId||’’),
店家:String(order.store||’’), 品項:String(order.item||’’), 規格:String(order.spec||’’),
備註:String(order.note||’’), 數量:qty, 單價:price, 總價:price*qty,
狀態:‘未付款’, 付款時間:’’, 付款方式:’’, 訂單備註:’’
});
return { success:true };
} catch(e) { console.error(‘Orders write fail:’, e.message); return { success:false, reason:e.message }; }
}

async function getOrdersByUser(userId, dateStr) {
try {
await authSheet();
const s = doc.sheetsByTitle[‘Orders’];
if (!s) return [];
const rows = await s.getRows();
const target = dateStr || todayTW();
return rows
.filter(r => String(r[‘userId’]||’’)===String(userId) && String(r[‘時間’]||’’).startsWith(target) && String(r[‘狀態’]||’’)!==‘已刪除’)
.map(r => ({
rowIndex:r.rowIndex, store:String(r[‘店家’]||’’), item:String(r[‘品項’]||’’),
spec:String(r[‘規格’]||’’), note:String(r[‘備註’]||’’),
qty:Number(r[‘數量’]||1), price:Number(r[‘單價’]||0), total:Number(r[‘總價’]||0),
status:String(r[‘狀態’]||‘未付款’), time:String(r[‘時間’]||’’)
}));
} catch(e) { return []; }
}

async function getOrderDates(userId) {
try {
await authSheet();
const s = doc.sheetsByTitle[‘Orders’];
if (!s) return [];
const rows = await s.getRows();
const dates = new Set();
rows.forEach(r => {
if (String(r[‘userId’]||’’)===userId && String(r[‘狀態’]||’’)!==‘已刪除’) {
const d = String(r[‘時間’]||’’).split(’ ’)[0];
if (d) dates.add(d);
}
});
return […dates].sort().reverse().slice(0, 14);
} catch(e) { return []; }
}

async function deleteOrder(userId, rowIndex) {
try {
await authSheet();
const s = doc.sheetsByTitle[‘Orders’];
if (!s) return false;
const rows = await s.getRows();
const t = rows.find(r => Number(r.rowIndex)===Number(rowIndex) && String(r[‘userId’]||’’)===String(userId) && String(r[‘狀態’]||’’)!==‘已刪除’);
if (!t) return false;
t[‘狀態’]=‘已刪除’; await t.save(); return true;
} catch(e) { return false; }
}

async function updateOrderQty(userId, rowIndex, qty) {
try {
await authSheet();
const s = doc.sheetsByTitle[‘Orders’];
if (!s) return false;
const rows = await s.getRows();
const t = rows.find(r =>
Number(r.rowIndex) === Number(rowIndex) &&
String(r[‘userId’] || ‘’) === String(userId) &&
String(r[‘狀態’] || ‘’) === ‘未付款’
);
if (!t) return false;
const q     = Math.min(20, Math.max(1, Number(qty) || 1));
const price = Number(t[‘單價’] || 0);
t[‘數量’] = q;
t[‘總價’] = price * q;
await t.save();
return true;
} catch(e) {
console.error(‘updateOrderQty fail:’, e.message);
return false;
}
}

async function updateOrderNote(userId, rowIndex, note) {
try {
await authSheet();
const s = doc.sheetsByTitle[‘Orders’];
if (!s) return false;
const rows = await s.getRows();
const t = rows.find(r => Number(r.rowIndex)===Number(rowIndex) && String(r[‘userId’]||’’)===String(userId) && String(r[‘狀態’]||’’)!==‘已刪除’);
if (!t) return false;
t[‘備註’]=note; await t.save(); return true;
} catch(e) { return false; }
}

async function getAllOrdersByDate(dateStr) {
try {
await authSheet();
const s = doc.sheetsByTitle[‘Orders’];
if (!s) return [];
const rows = await s.getRows();
const target = dateStr || todayTW();
return rows
.filter(r => String(r[‘時間’]||’’).startsWith(target))
.map(r => ({
rowIndex:r.rowIndex, name:String(r[‘LINE名稱’]||’’), userId:String(r[‘userId’]||’’),
store:String(r[‘店家’]||’’), item:String(r[‘品項’]||’’), spec:String(r[‘規格’]||’’),
note:String(r[‘備註’]||’’), qty:Number(r[‘數量’]||1), price:Number(r[‘單價’]||0),
total:Number(r[‘總價’]||0), status:String(r[‘狀態’]||‘未付款’),
payTime:String(r[‘付款時間’]||’’), payType:String(r[‘付款方式’]||’’)
}));
} catch(e) { return []; }
}

async function markPaid(rowIndex, payType) {
try {
await authSheet();
const s = doc.sheetsByTitle[‘Orders’];
if (!s) return false;
const rows = await s.getRows();
const t = rows.find(r => Number(r.rowIndex)===Number(rowIndex));
if (!t) return false;
t[‘狀態’]=‘已付款’; t[‘付款時間’]=nowTW(); t[‘付款方式’]=payType||‘現金’;
await t.save(); return true;
} catch(e) { return false; }
}

async function batchMarkPaid(payType) {
try {
await authSheet();
const s = doc.sheetsByTitle[‘Orders’];
if (!s) return 0;
const rows = await s.getRows();
const today = todayTW();
let count = 0;
for (const r of rows) {
if (String(r[‘時間’]||’’).startsWith(today) && String(r[‘狀態’]||’’)===‘未付款’) {
r[‘狀態’]=‘已付款’; r[‘付款時間’]=nowTW(); r[‘付款方式’]=payType||‘現金’;
await r.save(); count++;
}
}
return count;
} catch(e) { return 0; }
}

async function adminDeleteOrder(rowIndex) {
try {
await authSheet();
const s = doc.sheetsByTitle[‘Orders’];
if (!s) return false;
const rows = await s.getRows();
const t = rows.find(r => Number(r.rowIndex)===Number(rowIndex));
if (!t) return false;
t[‘狀態’]=‘已刪除’; await t.save(); return true;
} catch(e) { return false; }
}

async function clearTodayOrders() {
try {
await authSheet();
const s = doc.sheetsByTitle[‘Orders’];
if (!s) return { success:false, count:0 };
const rows  = await s.getRows();
const today = todayTW();
let count = 0;
for (const r of rows) {
if (String(r[‘時間’]||’’).startsWith(today) && String(r[‘狀態’]||’’)!==‘已刪除’) {
r[‘狀態’]=‘已刪除’; await r.save(); count++;
}
}
return { success:true, count };
} catch(e) {
console.error(‘clearTodayOrders fail:’, e.message);
return { success:false, count:0 };
}
}

// ════════════════════════════════════════════════════════════════
//  自動結單
// ════════════════════════════════════════════════════════════════
function scheduleAutoClose(minutes) {
if (autoCloseTimer) clearTimeout(autoCloseTimer);
const ms = Math.max(1, Number(minutes)) * 60 * 1000;
autoCloseAt    = new Date(Date.now() + ms).toISOString();

// 提前 5 分鐘推播提醒（只在時間 > 5 分鐘時）
if (ms > 5 * 60 * 1000) {
setTimeout(() => {
if (isOpen) pushToGroup(‘⏰ 距離收單還有 5 分鐘，請把握時間點餐！’);
}, ms - 5 * 60 * 1000);
}

autoCloseTimer = setTimeout(async () => {
isOpen = false; autoCloseAt = null; autoCloseTimer = null;
console.log(’[自動結單]’, nowTW());
const stat = await buildStatReport().catch(() => ‘’);
pushToGroup(‘🔴 已自動結單！\n\n’ + stat);
}, ms);
}
function cancelAutoClose() {
if (autoCloseTimer) clearTimeout(autoCloseTimer);
autoCloseTimer = null; autoCloseAt = null;
}

// ── 推播通知（需設定 LINE_GROUP_ID 環境變數）─────────────────────
async function pushToGroup(text) {
const gid = process.env.LINE_GROUP_ID || ‘’;
if (!gid) return;
try {
await client.pushMessage(gid, { type: ‘text’, text });
} catch(e) {
console.error(’[push] fail:’, e.message);
}
}

// ════════════════════════════════════════════════════════════════
//  LINE 報表
// ════════════════════════════════════════════════════════════════
async function buildStatReport() {
const orders = await getAllOrdersByDate(todayTW());
const active = orders.filter(o => o.status !== ‘已刪除’);
if (!active.length) return ‘📊 今日尚無訂單’;
const itemCount={}, userTotal={};
const unpaid = new Set();
for (const o of active) {
const k = o.item+(o.spec?’（’+o.spec+’）’:’’);
itemCount[k]=(itemCount[k]||0)+o.qty;
const n = o.name||o.userId||‘未知’;
userTotal[n]=(userTotal[n]||0)+o.total;
if (o.status===‘未付款’) unpaid.add(n);
}
const grand = Object.values(userTotal).reduce((a,b)=>a+b,0);
let t = ‘📊 今日訂餐統計\n─────────────\n【品項數量】\n’;
for (const k in itemCount) t += k+’ x’+itemCount[k]+’\n’;
t += ‘\n【個人金額】\n’;
for (const n in userTotal) t += n+’：$’+userTotal[n]+’\n’;
t += ‘\n💰 總金額：$’+grand;
t += unpaid.size ? ‘\n\n⚠️ 未付款：’+[…unpaid].join(’、’) : ‘\n\n✅ 所有人已付款’;
return t;
}

async function buildShopOrder() {
const orders = await getAllOrdersByDate(todayTW());
const active = orders.filter(o => o.status !== ‘已刪除’);
if (!active.length) return ‘今日尚無訂單’;
const itemCount={};
for (const o of active) {
const k = o.item+(o.spec?’（’+o.spec+’）’:’’);
itemCount[k]=(itemCount[k]||0)+o.qty;
}
let out=‘您好，今天訂購如下：\n\n’, total=0;
for (const k in itemCount) { out+=k+’ x’+itemCount[k]+’\n’; total+=itemCount[k]; }
const money = active.reduce((a,o)=>a+o.total,0);
return out+’\n總數：’+total+‘份\n總金額：’+money+‘元\n\n麻煩您，謝謝～’;
}

// ════════════════════════════════════════════════════════════════
//  Admin middleware
// ════════════════════════════════════════════════════════════════
function adminAuth(req, res, next) {
const t = req.headers[‘x-admin-token’] || req.query.token || ‘’;
if (!process.env.ADMIN_TOKEN || t !== process.env.ADMIN_TOKEN)
return res.status(401).json({ error:‘Unauthorized’ });
next();
}

// ════════════════════════════════════════════════════════════════
//  API 路由
// ════════════════════════════════════════════════════════════════
app.get(’/’, (_q,r) => r.send(‘LINE 訂餐機器人運作中’));

app.get(’/api/menu’, async (_q, res) => {
let menu = [], optionData = {};
try {
menu = await loadMenu();
} catch(e) {
console.error(’[/api/menu] loadMenu failed:’, e.message);
}
try {
optionData = await loadOptions();
} catch(e) {
console.error(’[/api/menu] loadOptions failed:’, e.message);
}
res.json({ menu, optionData });
});

app.get(’/api/status’, (_q, res) => res.json({ isOpen, autoCloseAt }));

// 管理員專用：取得後台連結（token 不暴露給前端）
app.get(’/api/admin/link’, (req, res) => {
const uid = req.query.userId || ‘’;
if (!isAdmin(uid)) return res.status(403).json({ error: ‘forbidden’ });
const token = process.env.ADMIN_TOKEN || ‘’;
const base  = process.env.APP_URL || ‘’;
res.json({ url: base + ‘/admin?token=’ + token });
});

app.get(’/api/my-dates’, async (req, res) => {
const { userId } = req.query;
if (!userId) return res.json([]);
res.json(await getOrderDates(userId));
});

app.post(’/api/order’, async (req, res) => {
res.json(await saveOrderToSheet(req.body));
});

app.get(’/api/my-orders’, async (req, res) => {
const { userId, date } = req.query;
if (!userId) return res.json([]);
res.json(await getOrdersByUser(userId, date||null));
});

app.delete(’/api/order/:ri’, async (req, res) => {
const { userId } = req.body;
if (!userId) return res.json({ success:false });
res.json({ success: await deleteOrder(userId, Number(req.params.ri)) });
});

app.patch(’/api/order/:ri/qty’, async (req, res) => {
const { userId, qty } = req.body;
if (!userId) return res.json({ success: false });
res.json({ success: await updateOrderQty(userId, Number(req.params.ri), qty) });
});

app.patch(’/api/order/:ri/note’, async (req, res) => {
const { userId, note } = req.body;
if (!userId) return res.json({ success:false });
res.json({ success: await updateOrderNote(userId, Number(req.params.ri), note) });
});

// Admin API
app.post(’/api/admin/open’, adminAuth, (_q,res) => { isOpen=true; res.json({ isOpen }); });
app.post(’/api/admin/close’, adminAuth, (_q,res) => { isOpen=false; cancelAutoClose(); res.json({ isOpen }); });
app.post(’/api/admin/clear’, adminAuth, (_q,res) => { isOpen=false; cancelAutoClose(); res.json({ ok:true }); });
app.post(’/api/admin/auto-close’, adminAuth, (req,res) => {
if (!isOpen) isOpen=true;
scheduleAutoClose(Number(req.body.minutes)||30);
res.json({ autoCloseAt });
});
app.post(’/api/admin/cancel-auto-close’, adminAuth, (_q,res) => { cancelAutoClose(); res.json({ ok:true }); });
app.get(’/api/admin/orders’, adminAuth, async (req,res) => {
res.json(await getAllOrdersByDate(req.query.date || todayTW()));
});
app.post(’/api/admin/paid’, adminAuth, async (req,res) => {
res.json({ success: await markPaid(req.body.rowIndex, req.body.payType) });
});
app.post(’/api/admin/batch-paid’, adminAuth, async (req,res) => {
res.json({ success:true, count: await batchMarkPaid(req.body.payType||‘現金’) });
});
app.post(’/api/admin/delete-order’, adminAuth, async (req,res) => {
res.json({ success: await adminDeleteOrder(req.body.rowIndex) });
});

app.post(’/api/admin/clear-today’, adminAuth, async (_q, res) => {
res.json(await clearTodayOrders());
});

// ════════════════════════════════════════════════════════════════
//  前台頁面（Array join，零 template literal，零 JSON 內嵌）
// ════════════════════════════════════════════════════════════════
app.get(’/order’, (_q, res) => {
const liffId = String(process.env.LIFF_ID || ‘2010025093-yATK02dc’).replace(/[^a-zA-Z0-9-]/g,’’);
res.setHeader(‘Content-Type’,‘text/html; charset=utf-8’);
res.end(buildOrderPage(liffId));
});

function buildOrderPage(liffId) {
const L = [];
const p = s => L.push(s);

/* ── HEAD ─────────────────────────────────────────────────── */
p(’<!DOCTYPE html><html lang="zh-TW"><head>’);
p(’<meta charset="utf-8">’);
p(’<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1,user-scalable=no">’);
p(’<title>訂餐小幫手</title>’);
p(’<script src="https://static.line-scdn.net/liff/edge/2/sdk.js"></script>’);
p(’<style>’);
p(’:root{–g:#06c755;–gd:#05a847;–r:#e53935;–bg:#f0f0f0;–card:#fff;–bdr:#e2e2e2;–txt:#1a1a1a;–sub:#999;–rr:14px;–sh:0 1px 6px rgba(0,0,0,.07),0 4px 16px rgba(0,0,0,.04)}’);
p(’*{box-sizing:border-box;-webkit-tap-highlight-color:transparent;margin:0;padding:0}’);
p(‘html,body{overflow-x:hidden}’);
p(‘body{font-family:-apple-system,“SF Pro Text”,Arial,“Microsoft JhengHei”,sans-serif;background:var(–bg);color:var(–txt);min-height:100vh}’);
p(’.hd{background:#fff;border-bottom:1px solid var(–bdr);position:sticky;top:0;z-index:100;padding:8px 14px 6px}’);
p(’.hd-top{display:flex;align-items:center;justify-content:space-between;gap:8px}’);
p(’.hd-title{display:flex;align-items:center;gap:7px;font-size:17px;font-weight:800;letter-spacing:-.3px}’);
p(’.hd-logo{width:30px;height:30px;border-radius:8px;background:linear-gradient(135deg,#06c755,#00a846);display:flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0}’);
p(’.hd-badge{font-size:10px;padding:3px 8px;border-radius:999px;font-weight:700;letter-spacing:.3px}’);
p(’.hd-badge.open{background:#dcfce7;color:#15803d}’);
p(’.hd-badge.closed{background:#fee2e2;color:#b91c1c}’);
p(’.hd-user{font-size:11px;color:var(–sub);margin-top:3px}’);
p(’.hd-cd{font-size:10px;color:var(–r);font-weight:700;margin-top:1px}’);
p(’.srch{padding:6px 12px 7px;background:#fff;border-bottom:1px solid var(–bdr)}’);
p(’.srch input{width:100%;padding:8px 14px 8px 34px;border-radius:999px;border:1.5px solid var(–bdr);font-size:13px;background:var(–bg);outline:none;background-image:url(“data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='14' height='14' fill='%23aaa' viewBox='0 0 16 16'%3E%3Cpath d='M11.742 10.344a6.5 6.5 0 1 0-1.397 1.398l3.85 3.85a1 1 0 0 0 1.415-1.415l-3.868-3.833zM12 6.5a5.5 5.5 0 1 1-11 0 5.5 5.5 0 0 1 11 0z'/%3E%3C/svg%3E”);background-repeat:no-repeat;background-position:11px center}’);
p(’.srch input:focus{border-color:var(–g);background:#fff}’);
p(’.tabs{display:flex;flex-wrap:nowrap;gap:6px;padding:6px 12px 8px;overflow-x:auto;overflow-y:hidden;-webkit-overflow-scrolling:touch;touch-action:pan-x;scrollbar-width:none;background:#fff;border-bottom:1px solid var(–bdr)}’);
p(’.tabs::-webkit-scrollbar{display:none}’);
p(’.tab{display:inline-flex;align-items:center;padding:5px 13px;border-radius:999px;border:1.5px solid var(–bdr);font-size:12px;font-weight:600;cursor:pointer;background:#fff;color:var(–sub);flex:0 0 auto;min-width:max-content;white-space:nowrap;flex-shrink:0;transition:background .15s,color .15s,border-color .15s}’);
p(’.tab.active{background:var(–g);color:#fff;border-color:var(–g)}’);
p(’.tab.starred{border-color:#fbbf24;color:#d97706}’);
p(’.tab.active.starred{background:#f59e0b;border-color:#f59e0b;color:#fff}’);
p(’.ctn{padding:8px 10px 140px;-webkit-overflow-scrolling:touch}’);
p(’.store-row{display:flex;align-items:center;justify-content:space-between;padding:4px 2px;margin:12px 0 6px}’);
p(’.store-name{font-size:14px;font-weight:800;color:var(–txt);letter-spacing:-.2px}’);
p(’.fav-btn{background:none;border:none;font-size:17px;cursor:pointer;padding:2px 4px;line-height:1;color:var(–sub);transition:transform .15s}’);
p(’.fav-btn:active{transform:scale(1.3)}’);
p(’.card{background:var(–card);border-radius:var(–rr);margin-bottom:8px;box-shadow:var(–sh);overflow:hidden;display:flex;flex-direction:row;align-items:stretch}’);
p(’.card-img-wrap{width:88px;min-width:88px;max-width:88px;overflow:hidden;border-radius:var(–rr) 0 0 var(–rr)}’);
p(’.card-img{width:100%;height:100%;object-fit:cover;display:block}’);
p(’.card-body{flex:1;min-width:0;padding:10px 12px 10px 12px;display:flex;flex-direction:column;justify-content:space-between}’);
p(’.card-top{flex:1}’);
p(’.card-store{font-size:10px;color:var(–sub);margin-bottom:2px;font-weight:500}’);
p(’.card-name{font-size:14px;font-weight:800;line-height:1.3;margin-bottom:3px;color:var(–txt)}’);
p(’.card-price{font-size:15px;color:var(–g);font-weight:800;margin-bottom:8px}’);
p(’.card-footer{display:flex;align-items:center;justify-content:flex-end}’);
p(’.add-btn{padding:7px 16px;border:none;border-radius:999px;background:var(–g);color:#fff;font-size:13px;font-weight:700;cursor:pointer;transition:background .15s,transform .1s;white-space:nowrap}’);
p(’.add-btn:active{background:var(–gd);transform:scale(.96)}’);
p(’.add-btn:disabled{background:#d1d5db;cursor:default;transform:none}’);
p(’.card.no-img{flex-direction:row}’);
p(’.fab-area{position:fixed;bottom:20px;right:14px;display:flex;flex-direction:column;align-items:flex-end;gap:8px;z-index:8000;pointer-events:none}’);
p(’.fab{pointer-events:auto}’);
p(’.fab{display:flex;align-items:center;gap:6px;padding:12px 18px;border:none;border-radius:999px;font-size:14px;font-weight:700;cursor:pointer;box-shadow:0 4px 20px rgba(0,0,0,.22);white-space:nowrap;transition:transform .15s,box-shadow .15s}’);
p(’.fab:active{transform:scale(.95);box-shadow:0 2px 10px rgba(0,0,0,.15)}’);
p(’.fab-cart{background:var(–g);color:#fff}’);
p(’.fab-orders{background:#fff;color:var(–txt);border:1.5px solid var(–bdr)}’);
p(’.fab-admin{background:#1e293b;color:#fff}’);
p(’.fab-hidden{display:none!important}’);
p(’.badge{background:var(–r);color:#fff;border-radius:50%;min-width:18px;height:18px;font-size:10px;font-weight:800;display:inline-flex;align-items:center;justify-content:center;padding:0 3px;margin-left:2px}’);
p(’.modal{display:none;position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9000;align-items:flex-end;justify-content:center;-webkit-backdrop-filter:blur(2px);backdrop-filter:blur(2px)}’);
p(’.modal.ctr{align-items:center}’);
p(’.modal.show{display:flex}’);
p(‘body.modal-open{overflow:hidden!important}’);
p(’.mbox{background:#fff;width:100%;max-width:500px;max-height:88vh;overflow-y:auto;border-radius:20px 20px 0 0;padding:20px 16px 32px;position:relative;box-shadow:0 -4px 30px rgba(0,0,0,.12)}’);
p(’.modal.ctr .mbox{border-radius:18px;margin:0 12px;max-height:82vh}’);
p(’.m-handle{width:36px;height:4px;border-radius:999px;background:var(–bdr);margin:0 auto 14px}’);
p(’.m-close{position:absolute;top:14px;right:14px;background:var(–bg);border:none;border-radius:50%;width:28px;height:28px;font-size:14px;cursor:pointer;display:flex;align-items:center;justify-content:center;color:var(–sub)}’);
p(’.m-title{font-size:17px;font-weight:800;margin-bottom:14px;padding-right:36px}’);
p(’.opt-g-title{font-size:12px;font-weight:700;color:var(–sub);margin:14px 0 8px;text-transform:uppercase;letter-spacing:.5px}’);
p(’.chips{display:flex;flex-wrap:wrap;gap:7px}’);
p(’.chip{padding:6px 14px;border-radius:999px;border:1.5px solid var(–bdr);font-size:13px;cursor:pointer;background:#fff;transition:background .12s,border-color .12s,color .12s;user-select:none}’);
p(’.chip.on{background:var(–g);color:#fff;border-color:var(–g)}’);
p(’.chip.off{opacity:.3;cursor:default}’);
p(’.qty-row{display:flex;align-items:center;justify-content:space-between;margin-top:16px}’);
p(’.qty-row label{font-weight:700;font-size:14px}’);
p(’.qty-ctrl{display:flex;align-items:center;border:1.5px solid var(–bdr);border-radius:999px;overflow:hidden}’);
p(’.qty-ctrl button{background:none;border:none;width:36px;height:36px;font-size:20px;cursor:pointer;color:var(–g);font-weight:700;display:flex;align-items:center;justify-content:center}’);
p(’.qty-ctrl span{min-width:30px;text-align:center;font-size:15px;font-weight:700}’);
p(’.note-input{width:100%;padding:10px 14px;border-radius:10px;border:1.5px solid var(–bdr);font-size:14px;margin-top:12px;outline:none;font-family:inherit}’);
p(’.note-input:focus{border-color:var(–g)}’);
p(’.btn-p{width:100%;padding:13px;border:none;border-radius:999px;background:var(–g);color:#fff;font-size:15px;font-weight:700;cursor:pointer;margin-top:14px;transition:background .15s}’);
p(’.btn-p:active{background:var(–gd)}’);
p(’.btn-p:disabled{background:#d1d5db;cursor:default}’);
p(’.btn-s{width:100%;padding:12px;border:1.5px solid var(–bdr);border-radius:999px;background:#fff;color:var(–sub);font-size:14px;font-weight:600;cursor:pointer;margin-top:8px}’);
p(’.btn-del{background:var(–r);color:#fff;border:none;border-radius:999px;padding:5px 12px;font-size:12px;cursor:pointer;font-weight:700}’);
p(’.btn-reorder{background:var(–bg);border:none;border-radius:999px;padding:4px 11px;font-size:12px;cursor:pointer;font-weight:700;color:var(–g);margin-top:4px}’);
p(’.orow{display:flex;align-items:flex-start;gap:10px;padding:11px 0;border-bottom:1px solid var(–bdr)}’);
p(’.oinfo{flex:1;min-width:0}’);
p(’.oname{font-weight:700;font-size:14px;margin-bottom:2px}’);
p(’.osub{font-size:11px;color:var(–sub);margin-bottom:2px}’);
p(’.oprice{font-size:14px;font-weight:800;color:var(–g);white-space:nowrap}’);
p(’.oact{display:flex;flex-direction:column;align-items:flex-end;gap:5px}’);
p(’.tag-paid{display:inline-block;padding:2px 8px;border-radius:999px;font-size:10px;font-weight:700;background:#dcfce7;color:#15803d}’);
p(’.tag-unpaid{display:inline-block;padding:2px 8px;border-radius:999px;font-size:10px;font-weight:700;background:#fef3c7;color:#b45309}’);
p(’.icon-btn{background:none;border:none;font-size:17px;cursor:pointer;padding:3px}’);
p(’.qty-inline{display:inline-flex;align-items:center;gap:0;border:1.5px solid var(–bdr);border-radius:999px;overflow:hidden;background:#fff}’);
p(’.qty-inline button{background:none;border:none;width:28px;height:28px;font-size:16px;cursor:pointer;color:var(–g);font-weight:700;display:flex;align-items:center;justify-content:center;line-height:1}’);
p(’.qty-inline span{min-width:24px;text-align:center;font-size:13px;font-weight:700}’);
p(’.dcwrap{display:flex;gap:6px;overflow-x:auto;padding-bottom:8px;margin-bottom:10px;scrollbar-width:none}’);
p(’.dcwrap::-webkit-scrollbar{display:none}’);
p(’.dchip{padding:4px 12px;border-radius:999px;border:1.5px solid var(–bdr);font-size:12px;cursor:pointer;white-space:nowrap;background:#fff;flex-shrink:0}’);
p(’.dchip.on{background:var(–g);color:#fff;border-color:var(–g)}’);
p(’.total-bar{font-size:15px;font-weight:800;margin:12px 0 4px}’);
p(’.empty{text-align:center;padding:32px 20px;color:var(–sub);font-size:14px}’);
p(’#lm{text-align:center;padding:40px 20px;color:var(–sub);font-size:14px}’);
p(’#toast{position:fixed;bottom:90px;left:50%;transform:translateX(-50%);background:rgba(17,17,17,.92);color:#fff;padding:9px 20px;border-radius:999px;font-size:13px;z-index:99999;opacity:0;transition:opacity .25s;pointer-events:none;white-space:nowrap;backdrop-filter:blur(4px)}’);
p(’</style></head><body>’);

/* ── HEADER ──────────────────────────────────────────────── */
p(’<div class="hd">’);
p(’  <div class="hd-top">’);
p(’    <div class="hd-title"><div class="hd-logo">🍱</div><span>訂餐小幫手</span></div>’);
p(’    <span class="hd-badge closed" id="hdBadge">未開放</span>’);
p(’  </div>’);
p(’  <div class="hd-user" id="hdUser">正在取得 LINE 使用者資料…</div>’);
p(’  <div class="hd-cd" id="hdCd"></div>’);
p(’</div>’);

/* ── SEARCH ──────────────────────────────────────────────── */
p(’<div class="srch"><input id="srchInput" type="search" placeholder="搜尋品項或店家..." oninput="onSearch()"></div>’);

/* ── TABS ────────────────────────────────────────────────── */
p(’<div class="tabs" id="tabsEl"></div>’);

/* ── MENU ────────────────────────────────────────────────── */
p(’<div class="ctn" id="menuEl"><div id="lm">菜單載入中…</div></div>’);

/* ── FABs ────────────────────────────────────────────────── */
p(’<div class="fab-area">’);
p(’  <button class="fab fab-admin fab-hidden" id="fabAdmin" onclick="goAdmin()">⚙️ 後台</button>’);
p(’  <button class="fab fab-orders fab-hidden" id="fabOrders" onclick="openModal(\'moOrders\')">📋 我的訂單</button>’);
p(’  <button class="fab fab-cart fab-hidden" id="fabCart" onclick="openModal(\'moCart\')">🛒 購物車<span class="badge" id="cartBadge">0</span></button>’);
p(’</div>’);

/* ── TOAST ───────────────────────────────────────────────── */
p(’<div id="toast"></div>’);

/* ── MODAL: 選項 ─────────────────────────────────────────── */
p(’<div class="modal" id="moOpts"><div class="mbox">’);
p(’  <div class="m-handle"></div>’);
p(’  <button class="m-close" onclick="closeModal(\'moOpts\')">✕</button>’);
p(’  <div class="m-title" id="moOptsTitle"></div>’);
p(’  <div id="moOptsBody"></div>’);
p(’  <div class="qty-row"><label>數量</label>’);
p(’    <div class="qty-ctrl">’);
p(’      <button onclick="chQty(-1)">−</button>’);
p(’      <span id="qtyVal">1</span>’);
p(’      <button onclick="chQty(1)">＋</button>’);
p(’    </div>’);
p(’  </div>’);
p(’  <input class="note-input" id="itemNote" placeholder="備註（例：不加辣、少冰）">’);
p(’  <button class="btn-p" onclick="submitOpts()">加入購物車 🛒</button>’);
p(’  <button class="btn-s" onclick="closeModal(\'moOpts\')">取消</button>’);
p(’</div></div>’);

/* ── MODAL: 購物車 ───────────────────────────────────────── */
p(’<div class="modal" id="moCart"><div class="mbox">’);
p(’  <div class="m-handle"></div>’);
p(’  <button class="m-close" onclick="closeModal(\'moCart\')">✕</button>’);
p(’  <div class="m-title">🛒 我的購物車</div>’);
p(’  <div id="cartList"></div>’);
p(’  <div class="total-bar" id="cartTotal"></div>’);
p(’  <button class="btn-p" id="submitBtn" onclick="submitCart()">送出訂單</button>’);
p(’  <button class="btn-s" onclick="closeModal(\'moCart\')">繼續點餐</button>’);
p(’</div></div>’);

/* ── MODAL: 我的訂單 ─────────────────────────────────────── */
p(’<div class="modal" id="moOrders"><div class="mbox">’);
p(’  <div class="m-handle"></div>’);
p(’  <button class="m-close" onclick="closeModal(\'moOrders\')">✕</button>’);
p(’  <div class="m-title">📋 我的訂單</div>’);
p(’  <div class="dcwrap" id="dateChips"></div>’);
p(’  <div id="orderList"></div>’);
p(’  <div class="total-bar" id="orderTotal"></div>’);
p(’  <button class="btn-s" onclick="closeModal(\'moOrders\')">關閉</button>’);
p(’</div></div>’);

/* ── MODAL: 備註編輯 ─────────────────────────────────────── */
p(’<div class="modal ctr" id="moNote"><div class="mbox">’);
p(’  <button class="m-close" onclick="closeModal(\'moNote\')">✕</button>’);
p(’  <div class="m-title">✏️ 修改備註</div>’);
p(’  <input class="note-input" id="noteInput" placeholder="備註內容">’);
p(’  <button class="btn-p" onclick="saveNote()">儲存</button>’);
p(’  <button class="btn-s" onclick="closeModal(\'moNote\')">取消</button>’);
p(’</div></div>’);

/* ════════════════════════════════════════════════════════════
JavaScript — 全部用一般字串，不用 template literal
無任何 JSON 內嵌，資料全部由 fetch 取得
══════════════════════════════════════════════════════════════ */
p(’<script>’);

/* 全域變數 */
p(“var LIFF_ID=’” + liffId + “’;”);
p(‘var menu=[],opts={},favs=[];’);
p(‘var profile=null,ready=false;’);
p(‘var curItem=null,curGroups=[],curQty=1;’);
p(‘var cart=[];’);
p(‘var editRI=null;’);
p(‘var orderDates=[],curDate=null;’);
p(‘var searchQ=””;’);

/* loadMenuData */
p(‘async function loadMenuData(){’);
p(’  try{’);
p(’    var r=await fetch(”/api/menu”);’);
p(’    var d=await r.json();’);
p(’    menu=d.menu||[];’);
p(’    opts=d.optionData||{};’);
p(’    renderMenu();’);
p(’  }catch(e){’);
p(’    var lm=document.getElementById(“lm”);’);
p(’    if(lm)lm.textContent=“菜單載入失敗，請重新整理”;’);
p(’  }’);
p(’}’);

/* initLIFF */
p(‘async function initLIFF(){’);
p(’  try{’);
p(’    await liff.init({liffId:LIFF_ID});’);
p(’    if(!liff.isLoggedIn()){liff.login();return;}’);
p(’    profile=await liff.getProfile();’);
p(’    ready=true;’);
p(’    document.getElementById(“hdUser”).textContent=“👤 “+profile.displayName;’);
p(’    loadFavs();’);
p(’    renderMenu();’);
p(’    enableBtns();’);
p(’    document.getElementById(“fabCart”).classList.remove(“fab-hidden”);’);
p(’    document.getElementById(“fabOrders”).classList.remove(“fab-hidden”);’);
p(’    checkStatus();’);
p(’    checkAdminLink();’);
p(’  }catch(e){’);
p(’    document.getElementById(“hdUser”).textContent=“LIFF 初始化失敗：”+e.message;’);
p(’    console.error(“LIFF error”,e);’);
p(’  }’);
p(’}’);

/* checkStatus */
p(‘async function checkStatus(){’);
p(’  try{’);
p(’    var r=await fetch(”/api/status”);’);
p(’    var d=await r.json();’);
p(’    var el=document.getElementById(“hdBadge”);’);
p(’    if(d.isOpen){el.textContent=“開放點餐”;el.className=“hd-badge open”;}’);
p(’    else{el.textContent=“未開放”;el.className=“hd-badge closed”;}’);
p(’    if(d.autoCloseAt)startCd(d.autoCloseAt);’);
p(’    else document.getElementById(“hdCd”).textContent=””;’);
p(’    updateOrderBtns(d.isOpen);’);
p(’  }catch(e){}’);
p(’}’);

p(‘function updateOrderBtns(open){’);
p(’  menu.forEach(function(_,i){’);
p(’    var b=document.getElementById(“btn”+i);’);
p(’    if(!b)return;’);
p(’    if(open){b.disabled=false;b.textContent=“加入購物車”;}’);
p(’    else{b.disabled=true;b.textContent=“未開放點餐”;}’);
p(’  });’);
p(’}’);

/* startCd */
p(‘function startCd(iso){’);
p(’  var el=document.getElementById(“hdCd”);’);
p(’  function tick(){’);
p(’    var diff=new Date(iso)-new Date();’);
p(’    if(diff<=0){el.textContent=””;return;}’);
p(’    var m=Math.floor(diff/60000),s=Math.floor((diff%60000)/1000);’);
p(’    el.textContent=“⏰ 自動結單：”+m+“分”+s+“秒後”;’);
p(’    setTimeout(tick,1000);’);
p(’  }’);
p(’  tick();’);
p(’}’);

/* loadFavs / saveFavs */
p(‘function loadFavs(){’);
p(’  try{’);
p(’    var k=“fav_”+(profile?profile.userId:””);’);
p(’    var v=localStorage.getItem(k);’);
p(’    favs=v?JSON.parse(v):[];’);
p(’  }catch(e){favs=[];}’);
p(’}’);
p(‘function saveFavs(){’);
p(’  try{localStorage.setItem(“fav_”+(profile?profile.userId:””),JSON.stringify(favs));}catch(e){}’);
p(’}’);
p(‘function toggleFav(store){’);
p(’  var i=favs.indexOf(store);’);
p(’  if(i>=0)favs.splice(i,1);else favs.push(store);’);
p(’  saveFavs();’);
p(’  renderMenu();’);
p(’}’);

/* renderMenu */
p(‘function renderMenu(){’);
p(’  var lm=document.getElementById(“lm”);’);
p(’  if(lm)lm.style.display=“none”;’);
p(’  var menuEl=document.getElementById(“menuEl”);’);
p(’  var tabsEl=document.getElementById(“tabsEl”);’);
p(’  if(!menu.length){menuEl.innerHTML='<div class="empty">目前沒有菜單資料</div>';tabsEl.innerHTML=””;return;}’);
/* 搜尋篩選 */
p(’  var q=searchQ.toLowerCase();’);
p(’  var filtered=q?menu.filter(function(m){return m.item.toLowerCase().indexOf(q)>=0||m.store.toLowerCase().indexOf(q)>=0;}):menu;’);
/* 店家排序：收藏優先 */
p(’  var allStores=[];’);
p(’  menu.forEach(function(m){if(allStores.indexOf(m.store)<0)allStores.push(m.store);});’);
p(’  var favOrd=favs.filter(function(s){return allStores.indexOf(s)>=0;});’);
p(’  var rest=allStores.filter(function(s){return favs.indexOf(s)<0;});’);
p(’  var stores=favOrd.concat(rest);’);
/* tabs */
p(’  tabsEl.innerHTML=stores.map(function(s,i){’);
p(’    var isFav=favs.indexOf(s)>=0;’);
p(’    return '<div class=“tab'+(isFav?” starred”:””)+'” id=“tab'+i+'” onclick=“goStore('+i+')”>'+escH(s)+'</div>';’);
p(’  }).join(””);’);
/* cards */
p(’  var html=””;’);
p(’  stores.forEach(function(store,si){’);
p(’    var items=filtered.filter(function(m){return m.store===store;});’);
p(’    if(!items.length)return;’);
p(’    var isFav=favs.indexOf(store)>=0;’);
// 重點：fav-btn 的 onclick 用 data-store attribute，完全避免店家名稱的 escape 問題
p(’    html+='<div class="store-row" id="store\'+si+\'">';’);
p(’    html+='<span class="store-name">'+escH(store)+'</span>';’);
p(’    html+='<button class="fav-btn" data-store="\'+escH(store)+\'" onclick="handleFavBtn(this)">'+( isFav?“⭐”:“☆”)+”</button>”;’);
p(’    html+='</div>';’);
p(’    items.forEach(function(m){’);
p(’      var idx=menu.indexOf(m);’);
p(’      html+='<div class=“card”+(m.image?””:” no-img”)+”>';’);
p(’      if(m.image){’);
p(’        html+='<div class="card-img-wrap"><img class="card-img" src="\'+escH(m.image)+\'" alt="\'+escH(m.item)+\'" loading="lazy" onerror="this.parentNode.remove()"></div>';’);
p(’      }’);
p(’      html+='<div class="card-body">';’);
p(’      html+='<div class="card-top">';’);
p(’      html+='<div class="card-store">'+escH(m.store)+'</div>';’);
p(’      html+='<div class="card-name">'+escH(m.item)+'</div>';’);
p(’      html+='<div class="card-price">$'+Number(m.price)+'</div>';’);
p(’      html+='</div>';’);
p(’      html+='<div class="card-footer"><button class="add-btn" id="btn\'+idx+\'" onclick="addToCart(\'+idx+\')" disabled>載入中…</button></div>';’);
p(’      html+='</div></div>';’);
p(’    });’);
p(’  });’);
p(’  menuEl.innerHTML=html||'<div class="empty">找不到相符的品項</div>';’);
p(’  if(ready)enableBtns();’);
p(’}’);

/* handleFavBtn — 從 data-store 取店家名，完全避免 escape 問題 */
p(‘function handleFavBtn(btn){’);
p(’  var store=btn.getAttribute(“data-store”);’);
p(’  if(store)toggleFav(store);’);
p(’}’);

p(‘function goStore(i){’);
p(’  document.querySelectorAll(”.tab”).forEach(function(b,j){b.classList.toggle(“active”,j===i);});’);
p(’  var el=document.getElementById(“store”+i);’);
p(’  if(el)el.scrollIntoView({behavior:“smooth”,block:“start”});’);
p(’}’);

p(‘function onSearch(){searchQ=document.getElementById(“srchInput”).value.trim();renderMenu();}’);

p(‘function enableBtns(){’);
p(’  menu.forEach(function(_,i){’);
p(’    var b=document.getElementById(“btn”+i);’);
p(’    if(b){b.disabled=false;b.textContent=“加入購物車”;}’);
p(’  });’);
p(’}’);

/* addToCart */
p(‘function addToCart(idx){’);
p(’  if(!ready||!profile){alert(“尚未取得 LINE 使用者資料”);return;}’);
p(’  curItem=menu[idx];’);
p(’  var key=curItem.store+”||”+curItem.item;’);
p(’  curGroups=opts[key]||[];’);
p(’  curQty=1;’);
p(’  document.getElementById(“moOptsTitle”).textContent=curItem.item;’);
p(’  document.getElementById(“qtyVal”).textContent=“1”;’);
p(’  document.getElementById(“itemNote”).value=””;’);
p(’  var body=document.getElementById(“moOptsBody”);’);
p(’  body.innerHTML=””;’);
p(’  if(!curGroups.length){’);
p(’    body.innerHTML='<p style="color:var(--sub);margin:8px 0;font-size:14px">此商品無需選擇規格。</p>';’);
p(’  }else{’);
p(’    curGroups.forEach(function(g,gi){’);
p(’      var wrap=document.createElement(“div”);’);
p(’      var title=document.createElement(“div”);’);
p(’      title.className=“opt-g-title”;’);
p(’      title.textContent=g.category+” (”+g.min+”~”+g.max+“選)”;’);
p(’      wrap.appendChild(title);’);
p(’      var chips=document.createElement(“div”);’);
p(’      chips.className=“chips”;’);
p(’      chips.id=“chips”+gi;’);
p(’      g.options.forEach(function(opt){’);
p(’        var c=document.createElement(“div”);’);
p(’        c.className=“chip”;’);
p(’        var optName=typeof opt===“object”?opt.name:opt;’);
p(’        var optExtra=typeof opt===“object”?(opt.extra||0):0;’);
p(’        c.textContent=optName+(optExtra>0?” +$”+optExtra:””);’);
p(’        c.dataset.optname=optName;’);
p(’        c.dataset.extra=optExtra;’);
p(’        c.addEventListener(“click”,function(){toggleChip(gi,g.max,c);});’);
p(’        chips.appendChild(c);’);
p(’      });’);
p(’      wrap.appendChild(chips);’);
p(’      body.appendChild(wrap);’);
p(’    });’);
p(’  }’);
p(’  openModal(“moOpts”);’);
p(’}’);

p(‘function chQty(d){curQty=Math.max(1,Math.min(20,curQty+d));document.getElementById(“qtyVal”).textContent=curQty;}’);

p(‘function toggleChip(gi,max,chip){’);
p(’  var chips=[].slice.call(document.querySelectorAll(”#chips”+gi+” .chip”));’);
p(’  var sel=chips.filter(function(c){return c.classList.contains(“on”);});’);
p(’  if(chip.classList.contains(“on”)){chip.classList.remove(“on”);}’);
p(’  else if(sel.length<max){chip.classList.add(“on”);}’);
p(’  var sel2=chips.filter(function(c){return c.classList.contains(“on”);});’);
p(’  chips.forEach(function(c){’);
p(’    if(sel2.length>=max&&!c.classList.contains(“on”))c.classList.add(“off”);’);
p(’    else c.classList.remove(“off”);’);
p(’  });’);
p(’}’);

p(‘function submitOpts(){’);
p(’  var specParts=[];’);
p(’  for(var i=0;i<curGroups.length;i++){’);
p(’    var g=curGroups[i];’);
p(’    var sel=[].slice.call(document.querySelectorAll(”#chips”+i+” .chip.on”));’);
p(’    if(sel.length<g.min||sel.length>g.max){alert(g.category+” 需要選 “+g.min+”~”+g.max+” 個”);return;}’);
p(’    if(sel.length)specParts.push(g.category+”：”+sel.map(function(c){return c.textContent;}).join(”、”));’);
p(’  }’);
p(’  // 計算所有選項的加價總和’);
p(’  var extraTotal=0;’);
p(’  document.querySelectorAll(”.chip.on”).forEach(function(ch){’);
p(’    extraTotal+=parseFloat(ch.dataset.extra||0);’);
p(’  });’);
p(’  var finalPrice=curItem.price+extraTotal;’);
p(’  // specParts 加上加價標示’);
p(’  var specStr=specParts.join(” “);’);
p(’  if(extraTotal>0)specStr=specStr+(specStr?” “:””)+”(+$”+extraTotal+”)”;’);
p(’  cart.push({store:curItem.store,item:curItem.item,spec:specStr,note:document.getElementById(“itemNote”).value.trim(),qty:curQty,price:finalPrice,basePrice:curItem.price,extra:extraTotal});’);
p(’  closeModal(“moOpts”);’);
p(’  updBadge();’);
p(’  showToast(“已加入購物車 🎉”);’);
p(’}’);

/* cart */
p(‘function updBadge(){document.getElementById(“cartBadge”).textContent=cart.reduce(function(a,c){return a+c.qty;},0);}’);

p(‘function renderCart(){’);
p(’  var el=document.getElementById(“cartList”);’);
p(’  var te=document.getElementById(“cartTotal”);’);
p(’  if(!cart.length){el.innerHTML='<div class="empty" style="padding:20px 0">購物車是空的</div>';te.textContent=””;return;}’);
p(’  el.innerHTML=cart.map(function(c,i){’);
p(’    var sub=[c.spec,c.note].filter(Boolean).join(“｜”);’);
p(’    return '<div class="orow">'’);
p(’      +'<div class="oinfo"><div class="oname">'+escH(c.item)+' x'+c.qty+'</div>'’);
p(’      +(sub?'<div class="osub">'+escH(sub)+'</div>':””)’);
p(’      +'</div>'’);
p(’      +'<div class="oact"><div class="oprice">$'+(c.price*c.qty)+'</div>'’);
p(’      +'<button class="btn-del" onclick="removeCart(\'+i+\')">移除</button></div></div>';’);
p(’  }).join(””);’);
p(’  te.textContent=“合計：$”+cart.reduce(function(a,c){return a+c.price*c.qty;},0);’);
p(’}’);

p(‘function removeCart(i){cart.splice(i,1);updBadge();renderCart();}’);

p(‘async function submitCart(){’);
p(’  if(!cart.length){alert(“購物車是空的”);return;}’);
p(’  // 送單前確認是否仍開放點餐’);
p(’  try{’);
p(’    var sr=await fetch(”/api/status”);’);
p(’    var sd=await sr.json();’);
p(’    if(!sd.isOpen){alert(“目前未開放點餐，無法送出訂單。”);return;}’);
p(’  }catch(e){}’);
p(’  var btn=document.getElementById(“submitBtn”);’);
p(’  btn.disabled=true;btn.textContent=“送出中…”;’);
p(’  for(var i=0;i<cart.length;i++){’);
p(’    var c=cart[i];’);
p(’    try{’);
p(’      var r=await fetch(”/api/order”,{method:“POST”,headers:{“Content-Type”:“application/json”},’);
p(’        body:JSON.stringify({store:c.store,item:c.item,spec:c.spec,note:c.note,qty:c.qty,price:c.price,name:profile.displayName,userId:profile.userId})});’);
p(’      var d=await r.json();’);
p(’      if(!d.success&&d.reason===“duplicate”)showToast(“⚠️ “+c.item+” 已送出，略過”);’);
p(’    }catch(e){console.error(“order err”,e);}’);
p(’  }’);
p(’  cart=[];updBadge();closeModal(“moCart”);’);
p(’  btn.disabled=false;btn.textContent=“送出訂單”;’);
p(’  showToast(“訂單已送出 ✅”);’);
p(’}’);

/* 我的訂單 */
p(‘async function loadMyOrders(dateStr){’);
p(’  curDate=dateStr||null;’);
p(’  var url=”/api/my-orders?userId=”+encodeURIComponent(profile.userId);’);
p(’  if(curDate)url+=”&date=”+encodeURIComponent(curDate);’);
p(’  try{var r=await fetch(url);renderOrders(await r.json());}catch(e){}’);
p(’}’);

p(‘async function loadDates(){’);
p(’  if(!profile)return;’);
p(’  try{’);
p(’    var r=await fetch(”/api/my-dates?userId=”+encodeURIComponent(profile.userId));’);
p(’    orderDates=await r.json();’);
p(’    renderDates();’);
p(’  }catch(e){}’);
p(’}’);

/* renderDates — 用 DOM API，完全避免閉包和字串 escape 問題 */
p(‘function renderDates(){’);
p(’  var el=document.getElementById(“dateChips”);’);
p(’  el.innerHTML=””;’);
p(’  if(!orderDates.length)return;’);
p(’  orderDates.forEach(function(d){’);
p(’    var div=document.createElement(“div”);’);
p(’    div.className=“dchip”+(d===curDate?” on”:””);’);
p(’    div.textContent=d;’);
p(’    div.addEventListener(“click”,(function(dd){return function(){loadMyOrders(dd);renderDates();};})(d));’);
p(’    el.appendChild(div);’);
p(’  });’);
p(’}’);

```
p('var ordersCache=[];');
```

p(‘function renderOrders(orders){’);
p(’  ordersCache=orders||[];’);
p(’  var el=document.getElementById(“orderList”);’);
p(’  var te=document.getElementById(“orderTotal”);’);
p(’  if(!ordersCache.length){el.innerHTML='<div class="empty" style="padding:20px 0">這天沒有訂單</div>';te.textContent=””;return;}’);
p(’  el.innerHTML=ordersCache.map(function(o,oi){’);
p(’    var sub=[o.spec,o.note].filter(Boolean).join(“｜”);’);
p(’    var tag=o.status===“已付款”?'<span class="tag-paid">已付款</span>':'<span class="tag-unpaid">未付款</span>';’);
p(’    var canEdit=o.status!==“已付款”;’);
p(’    return '<div class="orow">'’);
p(’      +'<div class="oinfo">'’);
p(’      +'<div class="oname">'+escH(o.item)+' x'+o.qty+' '+tag+'</div>'’);
p(’      +(sub?'<div class="osub">'+escH(sub)+'</div>':””)’);
p(’      +'<div class="osub">'+escH(o.store)+'</div>'’);
p(’      +'<button class="btn-reorder" onclick="reorder(\'+oi+\')">🔄 再訂一次</button>'’);
p(’      +'</div>'’);
p(’      +'<div class="oact">'’);
p(’      +'<div class="oprice">$'+o.total+'</div>'’);
p(’      +(canEdit?’);
p(’        '<div class="qty-inline">'+’);
p(’        '<button onclick="chOrderQty(\'+o.rowIndex+\',\'+o.qty+\',-1)">−</button>'+’);
p(’        '<span>'+o.qty+'</span>'+’);
p(’        '<button onclick="chOrderQty(\'+o.rowIndex+\',\'+o.qty+\',1)">＋</button>'+’);
p(’        '</div>'’);
p(’      :””)+(canEdit?'<button class=“icon-btn” onclick=“openEditNote('+o.rowIndex+','+JSON.stringify(o.note||””)+')”>✏️</button>':””)’);
p(’      +(canEdit?'<button class="icon-btn" onclick="delOrder(\'+o.rowIndex+\')">🗑</button>':””)’);
p(’      +'</div></div>';’);
p(’  }).join(””);’);
p(’  te.textContent=“合計：$”+ordersCache.reduce(function(a,o){return a+o.total;},0);’);
p(’}’);

/* reorder */
p(‘function reorder(oi){’);
p(’  var o=ordersCache[oi];’);
p(’  if(!o)return;’);
p(’  var idx=menu.findIndex(function(m){return m.store===o.store&&m.item===o.item;});’);
p(’  // 優先用歷史單價（含加價），其次 total/qty，最後才用菜單基礎價’);
p(’  var usePrice=0;’);
p(’  if(o.price&&o.price>0){’);
p(’    usePrice=o.price;’);
p(’  }else if(o.total&&o.qty&&o.qty>0){’);
p(’    usePrice=Math.round(o.total/o.qty);’);
p(’  }else if(idx>=0){’);
p(’    usePrice=menu[idx].price;’);
p(’  }’);
p(’  if(usePrice<=0&&idx<0){showToast(“此品項已不在菜單中”);return;}’);
p(’  cart.push({store:o.store,item:o.item,spec:o.spec||””,note:o.note||””,qty:o.qty||1,price:usePrice});’);
p(’  updBadge();’);
p(’  showToast(“已加入購物車 🔄”);’);
p(’}’);

/* delOrder */
p(‘async function delOrder(rowIndex){’);
p(’  if(!confirm(“確定要刪除這筆訂單？”))return;’);
p(’  var r=await fetch(”/api/order/”+rowIndex,{method:“DELETE”,headers:{“Content-Type”:“application/json”},body:JSON.stringify({userId:profile.userId})});’);
p(’  var d=await r.json();’);
p(’  if(d.success){showToast(“已刪除”);loadMyOrders(curDate);}else alert(“刪除失敗”);’);
p(’}’);

/* editNote */
p(‘async function chOrderQty(rowIndex,curQty,delta){’);
p(’  var newQty=Math.min(20,Math.max(1,(curQty||1)+delta));’);
p(’  if(newQty===curQty)return;’);
p(’  try{’);
p(’    var r=await fetch(”/api/order/”+rowIndex+”/qty”,{’);
p(’      method:“PATCH”,’);
p(’      headers:{“Content-Type”:“application/json”},’);
p(’      body:JSON.stringify({userId:profile.userId,qty:newQty})’);
p(’    });’);
p(’    var d=await r.json();’);
p(’    if(d.success){showToast(“數量已更新”);loadMyOrders(curDate);}’);
p(’    else showToast(“更新失敗，請重試”);’);
p(’  }catch(e){showToast(“更新失敗：”+e.message);}’);
p(’}’);

p(‘function openEditNote(rowIndex,cur){editRI=rowIndex;document.getElementById(“noteInput”).value=cur||””;openModal(“moNote”);}’);
p(‘async function saveNote(){’);
p(’  var note=document.getElementById(“noteInput”).value.trim();’);
p(’  var r=await fetch(”/api/order/”+editRI+”/note”,{method:“PATCH”,headers:{“Content-Type”:“application/json”},body:JSON.stringify({userId:profile.userId,note:note})});’);
p(’  var d=await r.json();’);
p(’  if(d.success){showToast(“備註已更新”);closeModal(“moNote”);loadMyOrders(curDate);}else alert(“更新失敗”);’);
p(’}’);

/* modal */
p(‘function openModal(id){’);
p(’  document.getElementById(id).classList.add(“show”);’);
p(’  document.body.classList.add(“modal-open”);’);
p(’  if(id===“moCart”)renderCart();’);
p(’  if(id===“moOrders”){loadDates();loadMyOrders(curDate);}’);
p(’}’);
p(‘function closeModal(id){’);
p(’  document.getElementById(id).classList.remove(“show”);’);
p(’  var any=document.querySelector(”.modal.show”);’);
p(’  if(!any)document.body.classList.remove(“modal-open”);’);
p(’}’);
/* 點背景關閉 */
p(‘document.querySelectorAll(”.modal”).forEach(function(m){’);
p(’  m.addEventListener(“click”,function(e){if(e.target===m)m.classList.remove(“show”);});’);
p(’});’);

/* toast */
p(‘var _adminUrl=null;’);
p(‘async function checkAdminLink(){’);
p(’  if(!profile)return;’);
p(’  try{’);
p(’    var r=await fetch(”/api/admin/link?userId=”+encodeURIComponent(profile.userId));’);
p(’    if(!r.ok)return;’);
p(’    var d=await r.json();’);
p(’    if(d.url){_adminUrl=d.url;document.getElementById(“fabAdmin”).classList.remove(“fab-hidden”);}’);
p(’  }catch(e){}’);
p(’}’);
p(‘function goAdmin(){if(_adminUrl)window.open(_adminUrl,”_blank”);}’);

p(‘var _toastTimer=null;’);
p(‘function showToast(msg){’);
p(’  var t=document.getElementById(“toast”);’);
p(’  t.textContent=msg;t.style.opacity=“1”;’);
p(’  clearTimeout(_toastTimer);’);
p(’  _toastTimer=setTimeout(function(){t.style.opacity=“0”;},2500);’);
p(’}’);

/* escH */
p(‘function escH(s){return String(s||””).replace(/&/g,”&”).replace(/</g,”<”).replace(/>/g,”>”).replace(/”/g,”"”);}’);

/* init */
p(‘loadMenuData();’);
p(‘initLIFF();’);
p(‘setInterval(checkStatus,30000);’);

p(’</script>’);
p(’</body></html>’);

return L.join(’\n’);
}

// ════════════════════════════════════════════════════════════════
//  管理後台
// ════════════════════════════════════════════════════════════════
app.get(’/admin’, (req, res) => {
const token = req.query.token || ‘’;
if (!process.env.ADMIN_TOKEN || token !== process.env.ADMIN_TOKEN)
return res.status(401).send(‘Unauthorized. 請附上 ?token=YOUR_ADMIN_TOKEN’);
res.setHeader(‘Content-Type’,‘text/html; charset=utf-8’);
res.end(buildAdminPage());
});

function buildAdminPage() {
const L = [];
const p = s => L.push(s);

/* ── HEAD ────────────────────────────────────────────────── */
p(’<!DOCTYPE html><html lang="zh-TW"><head>’);
p(’<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">’);
p(’<title>訂餐管理後台</title>’);
p(’<style>’);
p(’:root{–g:#06c755;–gd:#05a847;–r:#e53935;–o:#f57c00;–b:#2563eb;–bg:#f1f5f9;–card:#fff;–bdr:#e2e8f0;–txt:#0f172a;–sub:#64748b;–sh:0 1px 3px rgba(0,0,0,.08),0 4px 16px rgba(0,0,0,.04)}’);
p(’*{box-sizing:border-box;margin:0;padding:0}’);
p(‘body{font-family:-apple-system,Arial,“Microsoft JhengHei”,sans-serif;background:var(–bg);color:var(–txt);min-height:100vh}’);

/* header */
p(’.hd{background:linear-gradient(135deg,#0f172a,#1e293b);color:#fff;padding:14px 20px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:50;box-shadow:0 2px 8px rgba(0,0,0,.2)}’);
p(’.hd-left{display:flex;align-items:center;gap:10px}’);
p(’.hd-icon{width:34px;height:34px;border-radius:10px;background:linear-gradient(135deg,#06c755,#00a846);display:flex;align-items:center;justify-content:center;font-size:18px;flex-shrink:0}’);
p(’.hd h1{font-size:17px;font-weight:800;letter-spacing:-.3px}’);
p(’.hd-sub{font-size:11px;opacity:.5;margin-top:1px}’);
p(’.hd-time{font-size:11px;opacity:.5;white-space:nowrap}’);

/* layout */
p(’.wrap{padding:16px;max-width:1080px;margin:0 auto}’);
p(’.section{margin-bottom:16px}’);
p(’.section-title{font-size:12px;font-weight:700;color:var(–sub);letter-spacing:.8px;text-transform:uppercase;margin-bottom:8px;padding-left:2px}’);
p(’.card{background:var(–card);border-radius:16px;padding:18px;box-shadow:var(–sh);border:1px solid var(–bdr)}’);
p(’.card+.card{margin-top:12px}’);

/* stat grid */
p(’.stat-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}’);
p(’.stat-grid.six{grid-template-columns:repeat(3,1fr)}’);
p(’@media(min-width:600px){.stat-grid.six{grid-template-columns:repeat(6,1fr)}}’);
p(’.sbox{border-radius:12px;padding:13px 14px;border:1px solid var(–bdr);background:var(–card)}’);
p(’.sbox.hi{border-color:transparent}’);
p(’.sbox.g{background:linear-gradient(135deg,#f0fdf4,#dcfce7);border-color:#bbf7d0}’);
p(’.sbox.r{background:linear-gradient(135deg,#fff1f2,#fee2e2);border-color:#fecaca}’);
p(’.sbox.o{background:linear-gradient(135deg,#fffbeb,#fef3c7);border-color:#fde68a}’);
p(’.sbox.b{background:linear-gradient(135deg,#eff6ff,#dbeafe);border-color:#bfdbfe}’);
p(’.snum{font-size:22px;font-weight:800;letter-spacing:-.5px;margin-bottom:2px;line-height:1}’);
p(’.sbox.g .snum{color:#15803d}.sbox.r .snum{color:#be123c}.sbox.o .snum{color:#b45309}.sbox.b .snum{color:#1d4ed8}’);
p(’.slbl{font-size:11px;color:var(–sub);font-weight:500}’);

/* control area */
p(’.ctrl-row{display:flex;align-items:center;gap:8px;flex-wrap:wrap}’);
p(’.status-pill{display:inline-flex;align-items:center;gap:6px;padding:5px 12px;border-radius:999px;font-size:13px;font-weight:700}’);
p(’.status-pill.on{background:#dcfce7;color:#15803d}’);
p(’.status-pill.off{background:#fee2e2;color:#be123c}’);
p(’.dot{width:8px;height:8px;border-radius:50%;background:currentColor;flex-shrink:0}’);
p(’.ac-row{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-top:10px;padding-top:10px;border-top:1px solid var(–bdr)}’);
p(’.ac-info{font-size:12px;color:var(–r);font-weight:700}’);

/* buttons */
p(’.btn{padding:8px 16px;border:none;border-radius:999px;font-size:13px;font-weight:700;cursor:pointer;white-space:nowrap;transition:opacity .15s,transform .1s;display:inline-flex;align-items:center;gap:5px}’);
p(’.btn:active{transform:scale(.96)}’);
p(’.btn.g{background:var(–g);color:#fff}’);
p(’.btn.r{background:var(–r);color:#fff}’);
p(’.btn.o{background:var(–o);color:#fff}’);
p(’.btn.b{background:var(–b);color:#fff}’);
p(’.btn.ghost{background:var(–card);border:1.5px solid var(–bdr);color:var(–txt)}’);
p(’.btn.sm{padding:5px 12px;font-size:12px}’);
p(’.btn:hover{opacity:.88}’);

/* form controls */
p(‘input[type=number],input[type=date],select.sel{padding:7px 11px;border-radius:8px;border:1.5px solid var(–bdr);font-size:13px;background:var(–card);color:var(–txt);font-family:inherit}’);
p(‘input[type=number]{width:68px}’);
p(‘input[type=date]{color:var(–txt)}’);

/* toolbar */
p(’.toolbar{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:12px}’);
p(’.srch{flex:1;min-width:160px;padding:8px 14px 8px 34px;border-radius:999px;border:1.5px solid var(–bdr);font-size:13px;background:var(–card);outline:none;font-family:inherit;background-image:url(“data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='14' height='14' fill='%2394a3b8' viewBox='0 0 16 16'%3E%3Cpath d='M11.742 10.344a6.5 6.5 0 1 0-1.397 1.398l3.85 3.85a1 1 0 0 0 1.415-1.415l-3.868-3.833zM12 6.5a5.5 5.5 0 1 1-11 0 5.5 5.5 0 0 1 11 0z'/%3E%3C/svg%3E”);background-repeat:no-repeat;background-position:11px center}’);
p(’.srch:focus{border-color:var(–g);outline:none}’);
p(’.filter-row{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:12px}’);
p(’.filter-label{font-size:12px;font-weight:600;color:var(–sub)}’);

/* table */
p(’.tbl-wrap{overflow-x:auto;border-radius:10px;border:1px solid var(–bdr)}’);
p(‘table{width:100%;border-collapse:collapse;font-size:13px}’);
p(‘thead{position:sticky;top:0;z-index:1}’);
p(‘th{padding:10px 10px;border-bottom:1px solid var(–bdr);text-align:left;font-weight:700;font-size:11px;letter-spacing:.5px;text-transform:uppercase;color:var(–sub);background:#f8fafc;white-space:nowrap}’);
p(‘td{padding:10px 10px;border-bottom:1px solid var(–bdr);vertical-align:middle}’);
p(‘tr:last-child td{border-bottom:0}’);
p(‘tr:hover td{background:#f8fafc}’);

/* badges / tags */
p(’.store-badge{display:inline-block;padding:2px 8px;border-radius:6px;font-size:11px;font-weight:700;background:#dbeafe;color:#1e40af;white-space:nowrap}’);
p(’.tag{display:inline-block;padding:3px 9px;border-radius:999px;font-size:11px;font-weight:700;white-space:nowrap}’);
p(’.tag.p{background:#dcfce7;color:#15803d}’);
p(’.tag.u{background:#fef3c7;color:#b45309}’);
p(’.tag.d{background:#fee2e2;color:#be123c}’);

/* group section */
p(’.grp{margin-bottom:16px}’);
p(’.grp-hd{display:flex;align-items:center;justify-content:space-between;padding:9px 12px;background:var(–bg);border-radius:10px;margin-bottom:6px;border:1px solid var(–bdr)}’);
p(’.grp-name{font-size:14px;font-weight:800;display:flex;align-items:center;gap:7px}’);
p(’.grp-meta{font-size:12px;color:var(–sub);display:flex;align-items:center;gap:10px}’);
p(’.grp-meta strong{color:var(–txt);font-weight:700}’);

/* pay select */
p(’.pay-sel{padding:4px 8px;border-radius:7px;border:1.5px solid var(–bdr);font-size:11px;background:var(–card);color:var(–txt);cursor:pointer;font-family:inherit}’);

/* empty / loading */
p(’.placeholder{text-align:center;padding:32px;color:var(–sub);font-size:13px}’);

/* responsive */
p(’@media(max-width:640px){th,td{padding:8px 7px;font-size:12px}.stat-grid{grid-template-columns:repeat(2,1fr)}.stat-grid.six{grid-template-columns:repeat(2,1fr)}}’);
p(’.chart-title{font-size:12px;font-weight:700;color:var(–sub);letter-spacing:.5px;text-transform:uppercase;margin-bottom:10px}’);
p(’.chart-row{display:flex;align-items:center;gap:10px;margin-bottom:8px}’);
p(’.chart-lbl{font-size:12px;font-weight:600;color:var(–txt);width:90px;flex-shrink:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}’);
p(’.chart-bar-wrap{flex:1;background:#f1f5f9;border-radius:999px;height:20px;overflow:hidden;position:relative}’);
p(’.chart-bar{height:100%;border-radius:999px;background:linear-gradient(90deg,var(–g),#00c851);transition:width .4s ease;min-width:2px}’);
p(’.chart-bar.r{background:linear-gradient(90deg,var(–r),#ff6b6b)}’);
p(’.chart-val{font-size:12px;font-weight:700;color:var(–sub);width:54px;text-align:right;flex-shrink:0}’);
p(’.chart-section{margin-bottom:18px}’);
p(’</style></head><body>’);

/* ── HEADER ──────────────────────────────────────────────── */
p(’<div class="hd">’);
p(’  <div class="hd-left">’);
p(’    <div class="hd-icon">🍱</div>’);
p(’    <div><div class="hd h1" style="font-size:17px;font-weight:800">訂餐管理後台</div><div class="hd-sub" id="hdSub">載入中…</div></div>’);
p(’  </div>’);
p(’  <div class="hd-time" id="hdT"></div>’);
p(’</div>’);
p(’<div class="wrap">’);

/* ── 開單控制 ──────────────────────────────────────────────── */
p(’<div class="section">’);
p(’<div class="section-title">開單控制</div>’);
p(’<div class="card">’);
p(’  <div class="ctrl-row">’);
p(’    <span class="status-pill off" id="sPill"><span class="dot"></span><span id="sTxt">已結單</span></span>’);
p(’    <button class="btn g" onclick="doOpen()">▶ 開單</button>’);
p(’    <button class="btn r" onclick="doClose()">■ 結單</button>’);
p(’    <button class="btn ghost" onclick="doClear()">清空狀態</button>’);
p(’  </div>’);
p(’  <div class="ac-row">’);
p(’    <span class="filter-label">自動結單</span>’);
p(’    <input type="number" id="acMin" value="30" min="1" max="480"> 分鐘後’);
p(’    <button class="btn ghost sm" onclick="doAc()">設定</button>’);
p(’    <button class="btn ghost sm" onclick="doCancelAc()">取消</button>’);
p(’    <span class="ac-info" id="acInfo"></span>’);
p(’  </div>’);
p(’</div></div>’);

/* ── 今日統計 ──────────────────────────────────────────────── */
p(’<div class="section">’);
p(’<div class="section-title">今日統計</div>’);
p(’<div class="stat-grid six" id="statsEl">’);
p(’  <div class="sbox"><div class="snum">—</div><div class="slbl">載入中</div></div>’);
p(’</div>’);
p(’<div id="chartEl" style="margin-top:14px"></div>’);
p(’</div>’);

/* ── 訂單列表 ──────────────────────────────────────────────── */
p(’<div class="section">’);
p(’<div class="section-title">訂單列表</div>’);
p(’<div class="card">’);

/* toolbar */
p(’  <div class="toolbar">’);
p(’    <input class="srch" id="srch" placeholder="搜尋姓名、品項..." oninput="renderTable()">’);
p(’    <button class="btn g sm" onclick="doBatchPaid()">✅ 全標已付</button>’);
p(’    <button class="btn b sm" onclick="copyShop()">📋 複製店家單</button>’);
p(’    <button class="btn ghost sm" onclick="loadOrders()">↺ 重整</button>’);
p(’    <button class="btn r sm" onclick="doClearToday()">🗑 清空今日訂單</button>’);
p(’  </div>’);

/* filter row */
p(’  <div class="filter-row">’);
p(’    <span class="filter-label">日期</span>’);
p(’    <input type="date" id="dateF" onchange="loadOrders()">’);
p(’    <button class="btn ghost sm" onclick="document.getElementById(\'dateF\').value=\'\';loadOrders()">今日</button>’);
p(’    <span class="filter-label" style="margin-left:6px">分組</span>’);
p(’    <select class="sel" id="gbSel" onchange="renderTable()">’);
p(’      <option value="store" selected>依店家</option>’);
p(’      <option value="person">依人員</option>’);
p(’      <option value="none">不分組</option>’);
p(’    </select>’);
p(’  </div>’);

p(’  <div id="ordersEl"><div class="placeholder">載入中…</div></div>’);
p(’</div></div>’);
p(’</div>’); /* .wrap */

/* ── SCRIPT ──────────────────────────────────────────────── */
p(’<script>’);
p(’var TOKEN=(location.search.match(/[?&]token=([^&]*)/)||[])[1]||””;’);
p(‘var cache=[],acAt=null,acTimer=null;’);
p(‘document.getElementById(“hdT”).textContent=new Date().toLocaleString(“zh-TW”);’);

/* api helper */
p(‘async function api(path,method,body){’);
p(’  var o={method:method||“GET”,headers:{“Content-Type”:“application/json”,“x-admin-token”:TOKEN}};’);
p(’  if(body!==undefined)o.body=JSON.stringify(body);’);
p(’  return (await fetch(path,o)).json();’);
p(’}’);

/* status */
p(‘async function loadStatus(){’);
p(’  var d=await api(”/api/status”);’);
p(’  var pill=document.getElementById(“sPill”);’);
p(’  var txt=document.getElementById(“sTxt”);’);
p(’  var sub=document.getElementById(“hdSub”);’);
p(’  if(d.isOpen){pill.className=“status-pill on”;txt.textContent=“開單中”;sub.textContent=“目前開放點餐”;}’);
p(’  else{pill.className=“status-pill off”;txt.textContent=“已結單”;sub.textContent=“目前未開放”;}’);
p(’  if(d.autoCloseAt){acAt=d.autoCloseAt;startAcCd();}’);
p(’  else{acAt=null;document.getElementById(“acInfo”).textContent=””;}’);
p(’}’);

p(‘function startAcCd(){’);
p(’  if(acTimer)clearInterval(acTimer);’);
p(’  acTimer=setInterval(function(){’);
p(’    if(!acAt){clearInterval(acTimer);return;}’);
p(’    var diff=new Date(acAt)-new Date();’);
p(’    if(diff<=0){document.getElementById(“acInfo”).textContent=“已自動結單”;clearInterval(acTimer);loadStatus();return;}’);
p(’    document.getElementById(“acInfo”).textContent=“⏰ “+Math.floor(diff/60000)+“分”+Math.floor((diff%60000)/1000)+“秒後自動結單”;’);
p(’  },1000);’);
p(’}’);

/* open / close / clear */
p(‘async function doOpen(){await api(”/api/admin/open”,“POST”,{});loadStatus();loadOrders();}’);
p(‘async function doClose(){await api(”/api/admin/close”,“POST”,{});loadStatus();loadOrders();}’);
p(‘async function doClear(){if(!confirm(“確定清空狀態？”))return;await api(”/api/admin/clear”,“POST”,{});loadStatus();loadOrders();}’);
p(‘async function doAc(){var m=Number(document.getElementById(“acMin”).value)||30;await api(”/api/admin/auto-close”,“POST”,{minutes:m});loadStatus();}’);
p(‘async function doCancelAc(){await api(”/api/admin/cancel-auto-close”,“POST”,{});acAt=null;document.getElementById(“acInfo”).textContent=””;clearInterval(acTimer);}’);

/* load orders */
p(‘async function loadOrders(){’);
p(’  document.getElementById(“ordersEl”).innerHTML='<div class="placeholder">載入中…</div>';’);
p(’  var date=document.getElementById(“dateF”).value;’);
p(’  var d=await api(”/api/admin/orders”+(date?”?date=”+date:””));’);
p(’  cache=Array.isArray(d)?d:[];’);
p(’  renderStats(cache);’);
p(’  renderTable();’);
p(’}’);

/* stats */
p(‘function renderStats(o){’);
p(’  var act=o.filter(function(x){return x.status!==“已刪除”;});’);
p(’  var paid=act.filter(function(x){return x.status===“已付款”;});’);
p(’  var unpaid=act.filter(function(x){return x.status===“未付款”;});’);
p(’  var tot=act.reduce(function(a,x){return a+x.total;},0);’);
p(’  var pM=paid.reduce(function(a,x){return a+x.total;},0);’);
p(’  var uM=unpaid.reduce(function(a,x){return a+x.total;},0);’);
p(’  document.getElementById(“statsEl”).innerHTML=’);
p(’    sb(act.length,“筆訂單”,””)+sb(”$”+tot,“總金額”,“b”)+sb(”$”+pM,“已收款”,“g”)+sb(”$”+uM,“待收款”,“r”)+sb(paid.length,“已付款”,“g”)+sb(unpaid.length,“未付款”,“o”);’);
p(’  renderChart(act);’);
p(’}’);

p(‘function renderChart(act){’);
p(’  var el=document.getElementById(“chartEl”);’);
p(’  if(!el||!act.length){if(el)el.innerHTML=””;return;}’);
p(’  // 各店家統計’);
p(’  var stores={};’);
p(’  act.forEach(function(o){’);
p(’    if(!stores[o.store])stores[o.store]={total:0,qty:0};’);
p(’    stores[o.store].total+=o.total;’);
p(’    stores[o.store].qty+=o.qty;’);
p(’  });’);
p(’  var names=Object.keys(stores);’);
p(’  if(!names.length){el.innerHTML=””;return;}’);
p(’  var maxTotal=Math.max.apply(null,names.map(function(k){return stores[k].total;}));’);
p(’  var maxQty=Math.max.apply(null,names.map(function(k){return stores[k].qty;}));’);
p(’  function mkBar(val,max,cls,unit){’);
p(’    var pct=max>0?Math.round(val/max*100):0;’);
p(’    return '<div class="chart-bar \'+cls+\'" style="width:\'+pct+\'%"></div>';’);
p(’  }’);
p(’  var h='<div class="chart-title">各店家金額</div>';’);
p(’  names.forEach(function(k){’);
p(’    h+='<div class="chart-row">'+’);
p(’      '<div class="chart-lbl" title="\'+esc(k)+\'">'+ esc(k)+'</div>'+’);
p(’      '<div class="chart-bar-wrap">'+ mkBar(stores[k].total,maxTotal,””,”$”)+'</div>'+’);
p(’      '<div class="chart-val">$'+stores[k].total+'</div>'+’);
p(’    '</div>';’);
p(’  });’);
p(’  h+='<div class="chart-title" style="margin-top:14px">各店家份數</div>';’);
p(’  names.forEach(function(k){’);
p(’    h+='<div class="chart-row">'+’);
p(’      '<div class="chart-lbl" title="\'+esc(k)+\'">'+ esc(k)+'</div>'+’);
p(’      '<div class="chart-bar-wrap">'+ mkBar(stores[k].qty,maxQty,“r”,“份”)+'</div>'+’);
p(’      '<div class="chart-val">'+stores[k].qty+'份</div>'+’);
p(’    '</div>';’);
p(’  });’);
p(’  el.innerHTML='<div class="chart-section">'+h+'</div>';’);
p(’}’);

p(‘function sb(n,l,cls){return '<div class=“sbox hi'+(cls?” “+cls:””)+'”><div class="snum">'+n+'</div><div class="slbl">'+l+'</div></div>';}’);

/* render table */
p(‘function renderTable(){’);
p(’  var q=(document.getElementById(“srch”).value||””).toLowerCase();’);
p(’  var gb=document.getElementById(“gbSel”).value;’);
p(’  var data=cache.filter(function(o){’);
p(’    return !q||(o.name||””).toLowerCase().indexOf(q)>=0||(o.item||””).toLowerCase().indexOf(q)>=0||(o.store||””).toLowerCase().indexOf(q)>=0;’);
p(’  });’);
p(’  var el=document.getElementById(“ordersEl”);’);
p(’  if(!data.length){el.innerHTML='<div class="placeholder">無符合訂單</div>';return;}’);
p(’  if(gb===“none”){el.innerHTML=mkTbl(data);return;}’);
p(’  var groups={};’);
p(’  data.forEach(function(o){’);
p(’    var k=gb===“store”?(o.store||“其他”):(o.name||“未知”);’);
p(’    if(!groups[k])groups[k]=[];’);
p(’    groups[k].push(o);’);
p(’  });’);
p(’  var html=””;’);
p(’  Object.keys(groups).forEach(function(gk){’);
p(’    var rows=groups[gk];’);
p(’    var act=rows.filter(function(o){return o.status!==“已刪除”;});’);
p(’    var tot=act.reduce(function(a,o){return a+o.total;},0);’);
p(’    var paidCnt=act.filter(function(o){return o.status===“已付款”;}).length;’);
p(’    var unpaidCnt=act.filter(function(o){return o.status===“未付款”;}).length;’);
p(’    html+='<div class="grp">';’);
p(’    html+='<div class="grp-hd">';’);
p(’    html+='<div class="grp-name"><span class="store-badge">'+esc(gk)+'</span></div>';’);
p(’    html+='<div class="grp-meta">';’);
p(’    html+='<span>合計 <strong>$'+tot+'</strong></span>';’);
p(’    html+='<span>'+act.length+' 筆</span>';’);
p(’    if(unpaidCnt)html+='<span style="color:var(--o)">'+unpaidCnt+' 待付</span>';’);
p(’    if(paidCnt)html+='<span style="color:var(--g)">'+paidCnt+' 已付</span>';’);
p(’    html+='</div></div>';’);
p(’    html+=mkTbl(rows);’);
p(’    html+='</div>';’);
p(’  });’);
p(’  el.innerHTML=html;’);
p(’}’);

/* build table */
p(‘function mkTbl(rows){’);
p(’  if(!rows.length)return “”;’);
p(’  var h='<div class="tbl-wrap"><table><thead><tr>'’);
p(’    +'<th>姓名</th><th>店家</th><th>品項</th><th>規格</th><th>備註</th><th>數量</th><th>金額</th><th>狀態</th><th>操作</th>'’);
p(’    +'</tr></thead><tbody>';’);
p(’  h+=rows.map(function(o){’);
p(’    var tc=o.status===“已付款”?“p”:o.status===“已刪除”?“d”:“u”;’);
p(’    var ops=””;’);
p(’    if(o.status===“未付款”){’);
p(’      ops+='<select class="pay-sel" id="pt\'+o.rowIndex+\'"><option>現金</option><option>Line Pay</option><option>轉帳</option></select> ';’);
p(’      ops+='<button class="btn g sm" onclick="doPaid(\'+o.rowIndex+\')">付款</button> ';’);
p(’      ops+='<button class="btn r sm" onclick="doDelOrder(\'+o.rowIndex+\')">刪</button>';’);
p(’    }’);
p(’    return '<tr>'’);
p(’      +'<td style="font-weight:600">'+esc(o.name)+'</td>'’);
p(’      +'<td><span class="store-badge">'+esc(o.store)+'</span></td>'’);
p(’      +'<td style="font-weight:600">'+esc(o.item)+'</td>'’);
p(’      +'<td style="color:var(--sub);font-size:12px">'+esc(o.spec)+'</td>'’);
p(’      +'<td style="color:var(--sub);font-size:12px">'+esc(o.note)+'</td>'’);
p(’      +'<td style="text-align:center">'+o.qty+'</td>'’);
p(’      +'<td style="font-weight:800;color:var(--g)">$'+o.total+'</td>'’);
p(’      +'<td><span class="tag \'+tc+\'">'+esc(o.status)+'</span></td>'’);
p(’      +'<td>'+ops+'</td></tr>';’);
p(’  }).join(””);’);
p(’  h+=”</tbody></table></div>”;’);
p(’  return h;’);
p(’}’);

/* actions */
p(‘async function doPaid(ri){’);
p(’  var sel=document.getElementById(“pt”+ri);’);
p(’  await api(”/api/admin/paid”,“POST”,{rowIndex:ri,payType:sel?sel.value:“現金”});’);
p(’  loadOrders();’);
p(’}’);
p(‘async function doDelOrder(ri){if(!confirm(“確定刪除？”))return;await api(”/api/admin/delete-order”,“POST”,{rowIndex:ri});loadOrders();}’);
p(‘async function doBatchPaid(){if(!confirm(“將所有未付款標記為已付款？”))return;await api(”/api/admin/batch-paid”,“POST”,{payType:“現金”});loadOrders();}’);

p(‘async function doClearToday(){’);
p(’  var date=document.getElementById(“dateF”).value;’);
p(’  var label=date||“今日”;’);
p(’  if(!confirm(“確定清空 “+label+” 的所有訂單？此操作會將當日訂單標記為已刪除，無法還原。”))return;’);
p(’  if(!confirm(“再次確認：清空 “+label+” 全部訂單？”))return;’);
p(’  var d=await api(”/api/admin/clear-today”,“POST”,{});’);
p(’  if(d.success){alert(“已清空 “+d.count+” 筆訂單”);loadOrders();}’);
p(’  else alert(“清空失敗，請重試”);’);
p(’}’);

/* copy shop order */
p(‘function copyShop(){’);
p(’  var act=cache.filter(function(o){return o.status!==“已刪除”;});’);
p(’  if(!act.length){alert(“今日無訂單”);return;}’);
p(’  var cnt={};’);
p(’  act.forEach(function(o){var k=o.item+(o.spec?”（”+o.spec+”）”:””);cnt[k]=(cnt[k]||0)+o.qty;});’);
p(’  var dl=document.getElementById(“dateF”).value||“今日”;’);
p(’  var lines=[“您好，”+dl+” 訂購如下：”,””];’);
p(’  var n=0;’);
p(’  Object.keys(cnt).forEach(function(k){lines.push(k+” x”+cnt[k]);n+=cnt[k];});’);
p(’  var m=act.reduce(function(a,o){return a+o.total;},0);’);
p(’  lines.push(””,“總數：”+n+“份”,“總金額：”+m+“元”,””,“麻煩您，謝謝～”);’);
p(’  var txt=lines.join(”\\n”);’);
p(’  navigator.clipboard.writeText(txt).then(function(){’);
p(’    var btn=event.target;var orig=btn.textContent;’);
p(’    btn.textContent=“✓ 已複製”;btn.style.background=”#15803d”;’);
p(’    setTimeout(function(){btn.textContent=orig;btn.style.background=””;},1800);’);
p(’  }).catch(function(){alert(txt);});’);
p(’}’);

/* esc */
p(‘function esc(s){return String(s||””).replace(/&/g,”&”).replace(/</g,”<”).replace(/>/g,”>”);}’);

/* init */
p(‘loadStatus();’);
p(‘loadOrders();’);
p(‘setInterval(loadStatus,30000);’);
p(‘setInterval(loadOrders,60000);’);
p(’</script></body></html>’);

return L.join(’\n’);
}

// ════════════════════════════════════════════════════════════════
//  LINE Webhook
// ════════════════════════════════════════════════════════════════
app.post(’/webhook’, async (req, res) => {
try {
const event = req.body.events?.[0];
if (!event || event.type !== ‘message’ || event.message.type !== ‘text’)
return res.sendStatus(200);

```
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
  await loadTextMenu();
  const menuStr = buildMenuMessage(textMenuCache);
  const autoMsg = '🟢 開始點餐！\n⏰ 將於 '+autoOpen[1]+' 分鐘後自動結單\n\n' + menuStr;
  await reply(autoMsg);
  pushToGroup(autoMsg);
  return res.sendStatus(200);
}
if (text === '開單') {
  if (!isAdmin(uid)) { await reply('只有管理員可以開單'); return res.sendStatus(200); }
  if (isOpen) { await reply('目前已開單中'); return res.sendStatus(200); }
  isOpen = true;
  await loadTextMenu();
  const menuStr = buildMenuMessage(textMenuCache);
  const openMsg = '🟢 開始點餐！\n\n' + menuStr;
  await reply(openMsg);
  pushToGroup(openMsg);
  return res.sendStatus(200);
}
if (text === '結單' || text === '收單' || text === '統計') {
  if (!isAdmin(uid)) { await reply('只有管理員可以結單/統計'); return res.sendStatus(200); }
  isOpen = false; cancelAutoClose();
  textMenuCache = [];
  const stat = await buildStatReport();
  await reply(stat);
  pushToGroup('🔴 已收單！\n\n' + stat);
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

// ── 管理員指令：查看訂單 ─────────────────────────────────────
if (text === '查看訂單' || text === '訂單') {
  if (!isAdmin(uid)) { await reply('只有管理員可以查看'); return res.sendStatus(200); }
  await reply(await buildDetailedReport());
  return res.sendStatus(200);
}

// ── 管理員指令：匯出 ────────────────────────────────────────
if (text === '匯出' || text === '匯出訂單') {
  if (!isAdmin(uid)) { await reply('只有管理員可以匯出'); return res.sendStatus(200); }
  await reply(await buildExportText());
  return res.sendStatus(200);
}

// ── 管理員指令：XXX 已付款 ───────────────────────────────────
const paidMatch = text.match(/^(.+?)\s*已付款$/);
if (paidMatch && isAdmin(uid)) {
  const targetName = paidMatch[1].trim();
  const count = await markPaidByName(targetName);
  if (count > 0) {
    await reply(targetName + ' 已標記付款（' + count + ' 筆）✅');
  } else {
    await reply('找不到 ' + targetName + ' 的未付款訂單');
  }
  return res.sendStatus(200);
}

// ── 使用者指令：我的訂單 ────────────────────────────────────
if (text === '我的訂單' || text === '我點了什麼') {
  const myOrders = await getMyTextOrders(uid);
  if (!myOrders.length) {
    await reply('你今天還沒有訂單');
  } else {
    let msg = '📋 你的訂單：\n';
    let total = 0;
    for (const o of myOrders) {
      msg += '• ' + o.item;
      if (o.qty > 1) msg += ' ×' + o.qty;
      if (o.note) msg += '（' + o.note + '）';
      msg += '　$' + o.total;
      if (o.status === '已付款') msg += ' ✅';
      msg += '\n';
      total += o.total;
    }
    msg += '合計：$' + total;
    await reply(msg);
  }
  return res.sendStatus(200);
}

// ── 使用者指令：菜單 ────────────────────────────────────────
if (text === '菜單' || text === '今日菜單') {
  if (!isOpen) {
    await reply('目前尚未開團，請等待管理員開團 🍱');
    return res.sendStatus(200);
  }
  const menu = textMenuCache.length ? textMenuCache : await loadTextMenu();
  await reply(buildMenuMessage(menu));
  return res.sendStatus(200);
}

// ── 使用者指令：取消訂單 ────────────────────────────────────
if (text === '取消' || text === '取消訂單') {
  if (!isOpen) { await reply('目前未開團'); return res.sendStatus(200); }
  try {
    await authSheet();
    const s = doc.sheetsByTitle['Orders'];
    const rows = await s.getRows();
    const today = todayTW();
    let count = 0;
    for (const r of rows) {
      if (String(r['userId']||'')===String(uid) && String(r['時間']||'').startsWith(today) && String(r['狀態']||'')==='未付款') {
        r['狀態']='已刪除'; await r.save(); count++;
      }
    }
    await reply(count > 0 ? '已取消你的訂單（' + count + ' 筆）' : '找不到可取消的訂單');
  } catch(e) {
    await reply('取消失敗，請聯絡管理員');
  }
  return res.sendStatus(200);
}

// ── 開團時：純數字訂餐（e.g. "2", "2不要菜", "3+2加辣"）────
if (isOpen && textMenuCache.length) {
  const parsed = parseTextOrder(text, textMenuCache);
  if (parsed) {
    const result = await saveTextOrder(profileName, uid, parsed);
    if (result.success) {
      const action = result.action === 'updated' ? '已更新' : '已收到';
      let msg = action + ' 你的訂單 ✅\n';
      msg += parsed.item;
      if (parsed.qty > 1) msg += ' ×' + parsed.qty;
      if (parsed.note) msg += '（' + parsed.note + '）';
      msg += '\n金額：$' + (parsed.price * parsed.qty);
      msg += '\n\n輸入"我的訂單"查看，輸入"取消"可取消';
      await reply(msg);
    } else {
      await reply('訂單失敗，請重試或聯絡管理員');
    }
    return res.sendStatus(200);
  }
}

// ── 開團中但看不懂的訊息：給提示 ───────────────────────────
if (isOpen && textMenuCache.length) {
  await reply('請輸入數字點餐，例如"2"或"2不要菜"\n輸入"菜單"可重看菜單');
  return res.sendStatus(200);
}

return res.sendStatus(200);
```

} catch(err) {
console.error(‘Webhook error:’, err);
return res.sendStatus(200);
}
});

app.use((err,_q,res,_n) => { console.error(‘Global error:’,err); res.sendStatus(200); });

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(‘Server running on port’, PORT));
