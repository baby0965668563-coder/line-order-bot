require(‘dotenv’).config();

const express = require(‘express’);
const line    = require(’@line/bot-sdk’);
const { GoogleSpreadsheet } = require(‘google-spreadsheet’);

const app = express();

const lineConfig = {
channelAccessToken: process.env.CHANNEL_ACCESS_TOKEN,
channelSecret:      process.env.CHANNEL_SECRET
};

app.use(’/webhook’, line.middleware(lineConfig));
app.use(express.json());

const client = new line.Client(lineConfig);
const doc    = new GoogleSpreadsheet(process.env.SHEET_ID);

// ════════════════════════════════════════════════════════════════
//  全域狀態
// ════════════════════════════════════════════════════════════════
let isOpen        = false;
let autoCloseTimer = null;
let todayOrders   = [];   // [{ name, qty, note, time }]
let menuText      = ‘’;   // 今日菜單文字（可選）

const ADMINS = (process.env.ADMIN_IDS || ‘’).split(’,’).map(s => s.trim()).filter(Boolean);

function isAdmin(uid) {
return ADMINS.includes(uid);
}

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
//  Google Sheets
// ════════════════════════════════════════════════════════════════
async function authSheet() {
await doc.useServiceAccountAuth({
client_email: process.env.GOOGLE_CLIENT_EMAIL,
private_key:  process.env.GOOGLE_PRIVATE_KEY.replace(/\n/g, ‘\n’)
});
await doc.loadInfo();
}

async function writeOrder(name, qty, note) {
try {
await authSheet();
const sheet = doc.sheetsByTitle[‘Orders’];
if (!sheet) {
console.error(‘Orders sheet not found’);
return;
}
await sheet.addRow({
‘time’:  nowTW(),
‘name’:  name,
‘qty’:   qty,
‘note’:  note || ‘’
});
} catch (e) {
console.error(‘writeOrder fail:’, e.message);
}
}

// ════════════════════════════════════════════════════════════════
//  解析訂單文字
//  支援格式：
//    慧玲+1
//    阿明+2辣
//    慧玲 +1
//    慧玲+1 不辣
//    慧玲 + 1 加蛋少冰
// ════════════════════════════════════════════════════════════════
function parseOrder(text) {
// 去除前後空白
const t = text.trim();

// 格式：姓名 + 數量 [備註]
// 姓名：非數字、非+的文字（至少1字）
// 數量：1-2位數字
// 備註：剩餘文字（可空）
const m = t.match(/^([^\d+＋]+?)\s*[+＋]\s*([0-9]{1,2})\s*(.*)$/);
if (!m) return null;

const name = m[1].trim();
const qty  = parseInt(m[2], 10);
const note = m[3].trim();

if (!name || qty < 1 || qty > 20) return null;

return { name, qty, note };
}

// ════════════════════════════════════════════════════════════════
//  收單動作
// ════════════════════════════════════════════════════════════════
async function doClose(targetGroupId) {
isOpen = false;
if (autoCloseTimer) {
clearTimeout(autoCloseTimer);
autoCloseTimer = null;
}

const gid = targetGroupId || process.env.LINE_GROUP_ID || ‘’;
if (gid) {
try {
await client.pushMessage(gid, {
type: ‘text’,
text: ‘📦 已收單\n感謝大家訂購 🙏’
});
} catch (e) {
console.error(‘doClose push fail:’, e.message);
}
}
}

// ════════════════════════════════════════════════════════════════
//  解析 /開單 指令的時間
//  支援：/開單 19:30  /開單 7:30  /開單 1930
// ════════════════════════════════════════════════════════════════
function parseCloseTime(text) {
const m = text.match(/(\d{1,2})[：:.]?(\d{2})/);
if (!m) return null;
return { hour: parseInt(m[1], 10), minute: parseInt(m[2], 10) };
}

function msUntil(hour, minute) {
const now = new Date(new Date().toLocaleString(‘en-US’, { timeZone: ‘Asia/Taipei’ }));
const target = new Date(now);
target.setHours(hour, minute, 0, 0);
if (target <= now) target.setDate(target.getDate() + 1);
return target - now;
}

// ════════════════════════════════════════════════════════════════
//  格式化查單回覆
// ════════════════════════════════════════════════════════════════
function buildOrderList() {
if (!todayOrders.length) return ‘目前沒有訂單’;

let lines = [‘📋 今日訂單’];
for (const o of todayOrders) {
let line = o.name + ’ +’ + o.qty;
if (o.note) line += ’ ’ + o.note;
lines.push(line);
}
const total = todayOrders.reduce((s, o) => s + o.qty, 0);
lines.push(’’);
lines.push(‘共 ’ + total + ’ 份’);
return lines.join(’\n’);
}

// ════════════════════════════════════════════════════════════════
//  格式化店家單
// ════════════════════════════════════════════════════════════════
function buildShopMsg() {
if (!todayOrders.length) return ‘目前沒有訂單’;

const total = todayOrders.reduce((s, o) => s + o.qty, 0);

// 統計備註
const noteCount = {};
let noNoteQty = 0;
for (const o of todayOrders) {
if (o.note) {
noteCount[o.note] = (noteCount[o.note] || 0) + o.qty;
} else {
noNoteQty += o.qty;
}
}

let lines = [‘您好，今天訂購如下：’, ‘’];
lines.push(‘總共 ’ + total + ’ 份’);
if (noNoteQty > 0) lines.push(‘原味 ’ + noNoteQty + ’ 份’);
for (const note in noteCount) {
lines.push(note + ’ ’ + noteCount[note] + ’ 份’);
}
lines.push(’’);
lines.push(‘謝謝’);
return lines.join(’\n’);
}

// ════════════════════════════════════════════════════════════════
//  Webhook
// ════════════════════════════════════════════════════════════════
app.post(’/webhook’, async (req, res) => {
res.sendStatus(200); // 先回 200，避免 LINE timeout

try {
const events = req.body.events || [];
for (const event of events) {
if (event.type !== ‘message’ || event.message.type !== ‘text’) continue;

```
  const text  = event.message.text.trim();
  const uid   = event.source.userId;
  const srcType = event.source.type;
  const groupId = event.source.groupId || '';

  const replyToken = event.replyToken;

  const reply = async (msg) => {
    try {
      await client.replyMessage(replyToken, { type: 'text', text: msg });
    } catch (e) {
      console.error('reply fail:', e.message);
    }
  };

  // ── 取得發言者名稱 ─────────────────────────────────────────
  let senderName = '';
  try {
    if (srcType === 'group') {
      const profile = await client.getGroupMemberProfile(groupId, uid);
      senderName = profile.displayName;
    } else {
      const profile = await client.getProfile(uid);
      senderName = profile.displayName;
    }
  } catch (e) {
    console.error('getProfile fail:', e.message);
  }

  // ════════════════════════════════════════════════════════════
  //  管理員指令
  // ════════════════════════════════════════════════════════════

  // /開單 19:30
  if (/^\/(開單|開)/.test(text) && isAdmin(uid)) {
    const timeStr = text.replace(/^\/(開單|開)\s*/, '').trim();
    const t = timeStr ? parseCloseTime(timeStr) : null;

    isOpen = true;
    todayOrders = [];

    let replyMsg = '📢 已開放訂購';

    if (t) {
      const ms = msUntil(t.hour, t.minute);
      const closeTimeStr = String(t.hour).padStart(2, '0') + ':' + String(t.minute).padStart(2, '0');
      replyMsg += '\n🕐 將於 ' + closeTimeStr + ' 自動收單';

      if (autoCloseTimer) clearTimeout(autoCloseTimer);
      autoCloseTimer = setTimeout(() => {
        doClose(groupId);
      }, ms);
    }

    if (menuText) replyMsg += '\n\n' + menuText;

    await reply(replyMsg);
    continue;
  }

  // /收單
  if (/^\/(收單|結單|關)/.test(text) && isAdmin(uid)) {
    await doClose(groupId);
    await reply(buildOrderList());
    continue;
  }

  // /查單
  if (/^\/(查單|查|訂單)/.test(text) && isAdmin(uid)) {
    await reply(buildOrderList());
    continue;
  }

  // /店家
  if (/^\/(店家|店)/.test(text) && isAdmin(uid)) {
    await reply(buildShopMsg());
    continue;
  }

  // /清單（清空今日訂單）
  if (/^\/(清單|清空|重置)/.test(text) && isAdmin(uid)) {
    todayOrders = [];
    await reply('已清空今日訂單');
    continue;
  }

  // /菜單 （設定或顯示今日菜單文字）
  if (/^\/(菜單|menu)/.test(text) && isAdmin(uid)) {
    const content = text.replace(/^\/(菜單|menu)\s*/, '').trim();
    if (content) {
      menuText = content;
      await reply('菜單已設定：\n' + menuText);
    } else if (menuText) {
      await reply(menuText);
    } else {
      await reply('今日尚未設定菜單');
    }
    continue;
  }

  // /狀態
  if (/^\/(狀態|status)/.test(text) && isAdmin(uid)) {
    const total = todayOrders.reduce((s, o) => s + o.qty, 0);
    const statusMsg = (isOpen ? '🟢 開單中' : '🔴 已收單') +
      '\n目前訂單：' + todayOrders.length + ' 筆，共 ' + total + ' 份';
    await reply(statusMsg);
    continue;
  }

  // ════════════════════════════════════════════════════════════
  //  一般訂單：姓名+數量[備註]
  // ════════════════════════════════════════════════════════════
  const order = parseOrder(text);
  if (order) {
    if (!isOpen) {
      // 收單後不記錄，靜默不回（避免洗版）
      continue;
    }

    // 同名覆蓋（同一輪開單期間，同名視為修改）
    const existIdx = todayOrders.findIndex(o => o.name === order.name);
    if (existIdx >= 0) {
      todayOrders[existIdx] = { ...order, time: nowTW() };
    } else {
      todayOrders.push({ ...order, time: nowTW() });
    }

    // 寫入 Sheets（非同步，不阻塞回覆）
    writeOrder(order.name, order.qty, order.note).catch(() => {});

    // 回覆確認
    let confirmMsg = '✅ ' + order.name + ' +' + order.qty;
    if (order.note) confirmMsg += ' ' + order.note;
    await reply(confirmMsg);
    continue;
  }

  // ════════════════════════════════════════════════════════════
  //  其他訊息：不回覆（避免打擾群組）
  // ════════════════════════════════════════════════════════════
}
```

} catch (err) {
console.error(‘webhook error:’, err);
}
});

// ════════════════════════════════════════════════════════════════
//  健康檢查
// ════════════════════════════════════════════════════════════════
app.get(’/’, (_req, res) => {
const total = todayOrders.reduce((s, o) => s + o.qty, 0);
res.send(‘LINE 訂餐機器人運作中｜’ + (isOpen ? ‘開單中’ : ‘已收單’) + ‘｜今日 ’ + total + ’ 份’);
});

// ════════════════════════════════════════════════════════════════
//  啟動
// ════════════════════════════════════════════════════════════════
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
console.log(‘Server running on port’, PORT);
});