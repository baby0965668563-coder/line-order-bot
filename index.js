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

let isOpen = false;
let autoCloseTimer = null;
let autoCloseAt = null;
let textMenuCache = [];
let sheetReady = false;

const admins = [
  'U8d9c82446aa9eb90d7de001cfc7ea90f',
  'Ubcfae64b443b9fad21bbc584e991b306',
  'U5c44a04efc62664bd45ec80d77be7d93',
  'Uc669eca67bf477460945f45751edd3e9'
];

function isAdmin(uid) {
  return admins.includes(uid);
}

function nowTW() {
  return new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
}

function todayTW() {
  return new Date().toLocaleDateString('zh-TW', { timeZone: 'Asia/Taipei' });
}

async function authSheet() {
  if (sheetReady) return;

  await doc.useServiceAccountAuth({
    client_email: process.env.GOOGLE_CLIENT_EMAIL,
    private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n')
  });

  await doc.loadInfo();
  sheetReady = true;
}

async function loadMenu() {
  await authSheet();

  const sheet = doc.sheetsByTitle['Menu'];
  if (!sheet) {
    console.error('Menu sheet not found');
    return [];
  }

  const rows = await sheet.getRows();

  return rows
    .map(r => ({
      store: String(r['店家'] || '').trim(),
      item: String(r['品項'] || '').trim(),
      price: Number(String(r['價格'] || '0').replace(/[^0-9.]/g, '')) || 0
    }))
    .filter(r => r.store && r.item && r.price > 0);
}

async function loadTextMenu() {
  const menu = await loadMenu();

  textMenuCache = menu.map((m, i) => ({
    no: i + 1,
    store: m.store,
    item: m.item,
    price: m.price
  }));

  return textMenuCache;
}

function buildMenuMessage(menuList) {
  if (!menuList.length) return '目前沒有菜單資料';

  const emojiNum = ['1️⃣','2️⃣','3️⃣','4️⃣','5️⃣','6️⃣','7️⃣','8️⃣','9️⃣','🔟'];

  let msg = '🍱 今日菜單\n';
  msg += '─────────────\n';

  menuList.forEach(m => {
    const no = m.no <= 10 ? emojiNum[m.no - 1] : m.no + '.';
    msg += no + ' ' + m.item + '　$' + m.price + '\n';
  });

  msg += '\n請直接回覆數字點餐：\n';
  msg += '例：\n';
  msg += '2　　　（點2號）\n';
  msg += '2不要菜　（加備註）\n';
  msg += '2+3　　（點2號，3份）';

  return msg;
}

function parseTextOrder(text, menuList) {
  const m = text.match(/^([0-9]+)(?:[+＋]([0-9]+))?(.*)$/);
  if (!m) return null;

  const no = parseInt(m[1], 10);
  const qty = Math.min(20, Math.max(1, parseInt(m[2] || '1', 10)));
  const note = String(m[3] || '').trim();

  const item = menuList.find(x => x.no === no);
  if (!item) return null;

  return {
    no,
    store: item.store,
    item: item.item,
    price: item.price,
    qty,
    note
  };
}

async function saveUserToSheet(name, userId, sourceType, groupId) {
  try {
    await authSheet();

    const sheet = doc.sheetsByTitle['Users'];
    if (!sheet) return;

    await sheet.addRow({
      時間: nowTW(),
      LINE名稱: name,
      userId: userId,
      來源類型: sourceType,
      群組ID: groupId || '',
      權限: isAdmin(userId) ? 'admin' : 'user'
    });
  } catch (e) {
    console.error('saveUserToSheet fail:', e.message);
  }
}

async function saveTextOrder(name, userId, parsed) {
  try {
    await authSheet();

    const sheet = doc.sheetsByTitle['Orders'];
    if (!sheet) return { success: false, reason: '找不到 Orders 工作表' };

    const rows = await sheet.getRows();
    const today = todayTW();

    const existing = rows.find(r =>
      String(r['userId'] || '') === String(userId) &&
      String(r['品項'] || '').trim() === String(parsed.item).trim() &&
      String(r['時間'] || '').startsWith(today) &&
      String(r['狀態'] || '') !== '已刪除'
    );

    if (existing) {
      existing['數量'] = parsed.qty;
      existing['備註'] = parsed.note;
      existing['總價'] = parsed.price * parsed.qty;
      await existing.save();

      return { success: true, action: 'updated' };
    }

    await sheet.addRow({
      時間: nowTW(),
      LINE名稱: name,
      userId: String(userId),
      店家: parsed.store,
      品項: parsed.item,
      規格: '',
      備註: parsed.note,
      數量: parsed.qty,
      單價: parsed.price,
      總價: parsed.price * parsed.qty,
      狀態: '未付款',
      付款時間: '',
      付款方式: '',
      訂單備註: ''
    });

    return { success: true, action: 'created' };
  } catch (e) {
    console.error('saveTextOrder fail:', e.message);
    return { success: false, reason: e.message };
  }
}

async function getMyTextOrders(userId) {
  try {
    await authSheet();

    const sheet = doc.sheetsByTitle['Orders'];
    if (!sheet) return [];

    const rows = await sheet.getRows();
    const today = todayTW();

    return rows
      .filter(r =>
        String(r['userId'] || '') === String(userId) &&
        String(r['時間'] || '').startsWith(today) &&
        String(r['狀態'] || '') !== '已刪除'
      )
      .map(r => ({
        item: String(r['品項'] || ''),
        note: String(r['備註'] || ''),
        qty: Number(r['數量'] || 1),
        price: Number(r['單價'] || 0),
        total: Number(r['總價'] || 0),
        status: String(r['狀態'] || '')
      }));
  } catch (e) {
    console.error('getMyTextOrders fail:', e.message);
    return [];
  }
}

async function getAllOrdersToday() {
  try {
    await authSheet();

    const sheet = doc.sheetsByTitle['Orders'];
    if (!sheet) return [];

    const rows = await sheet.getRows();
    const today = todayTW();

    return rows
      .filter(r => String(r['時間'] || '').startsWith(today))
      .map(r => ({
        row: r,
        name: String(r['LINE名稱'] || ''),
        userId: String(r['userId'] || ''),
        store: String(r['店家'] || ''),
        item: String(r['品項'] || ''),
        note: String(r['備註'] || ''),
        qty: Number(r['數量'] || 1),
        price: Number(r['單價'] || 0),
        total: Number(r['總價'] || 0),
        status: String(r['狀態'] || '未付款')
      }));
  } catch (e) {
    console.error('getAllOrdersToday fail:', e.message);
    return [];
  }
}

async function buildStatReport() {
  const orders = await getAllOrdersToday();
  const active = orders.filter(o => o.status !== '已刪除');

  if (!active.length) return '📊 今日尚無訂單';

  const itemCount = {};
  const userTotal = {};
  const unpaid = new Set();

  for (const o of active) {
    itemCount[o.item] = (itemCount[o.item] || 0) + o.qty;

    const name = o.name || '未知';
    userTotal[name] = (userTotal[name] || 0) + o.total;

    if (o.status === '未付款') unpaid.add(name);
  }

  let msg = '📊 今日訂餐統計\n';
  msg += '─────────────\n';
  msg += '【品項數量】\n';

  for (const item in itemCount) {
    msg += item + ' x' + itemCount[item] + '\n';
  }

  msg += '\n【個人金額】\n';

  for (const name in userTotal) {
    msg += name + '：$' + userTotal[name] + '\n';
  }

  const grand = active.reduce((sum, o) => sum + o.total, 0);

  msg += '\n💰 總金額：$' + grand;

  if (unpaid.size) {
    msg += '\n\n⚠️ 未付款：' + Array.from(unpaid).join('、');
  } else {
    msg += '\n\n✅ 所有人已付款';
  }

  return msg;
}

async function buildDetailedReport() {
  const orders = await getAllOrdersToday();
  const active = orders.filter(o => o.status !== '已刪除');

  if (!active.length) return '📋 今日尚無訂單';

  const byPerson = {};

  for (const o of active) {
    const name = o.name || '未知';

    if (!byPerson[name]) {
      byPerson[name] = {
        orders: [],
        total: 0
      };
    }

    byPerson[name].orders.push(o);
    byPerson[name].total += o.total;
  }

  let msg = '📋 今日訂單明細\n';
  msg += '─────────────\n';

  for (const name in byPerson) {
    const p = byPerson[name];
    const paid = p.orders.every(o => o.status === '已付款');
    msg += (paid ? '✅ ' : '💰 ') + name + '（$' + p.total + '）\n';

    for (const o of p.orders) {
      msg += '  • ' + o.item;
      if (o.qty > 1) msg += ' ×' + o.qty;
      if (o.note) msg += '（' + o.note + '）';
      msg += '\n';
    }
  }

  const grand = active.reduce((sum, o) => sum + o.total, 0);

  msg += '─────────────\n';
  msg += '💰 總金額：$' + grand;

  return msg;
}

async function buildShopOrder() {
  const orders = await getAllOrdersToday();
  const active = orders.filter(o => o.status !== '已刪除');

  if (!active.length) return '今日尚無訂單';

  const itemCount = {};

  for (const o of active) {
    itemCount[o.item] = (itemCount[o.item] || 0) + o.qty;
  }

  let msg = '您好，今天訂購如下：\n\n';
  let totalQty = 0;

  for (const item in itemCount) {
    msg += item + ' x' + itemCount[item] + '\n';
    totalQty += itemCount[item];
  }

  const money = active.reduce((sum, o) => sum + o.total, 0);

  msg += '\n總數：' + totalQty + '份';
  msg += '\n總金額：' + money + '元';
  msg += '\n\n麻煩您，謝謝～';

  return msg;
}

async function markPaidByName(name) {
  try {
    const orders = await getAllOrdersToday();

    let count = 0;

    for (const o of orders) {
      if (
        o.name.trim() === name.trim() &&
        o.status === '未付款'
      ) {
        o.row['狀態'] = '已付款';
        o.row['付款時間'] = nowTW();
        o.row['付款方式'] = '現金';
        await o.row.save();
        count++;
      }
    }

    return count;
  } catch (e) {
    console.error('markPaidByName fail:', e.message);
    return 0;
  }
}

async function cancelMyOrders(userId) {
  try {
    const orders = await getAllOrdersToday();

    let count = 0;

    for (const o of orders) {
      if (
        String(o.userId) === String(userId) &&
        o.status === '未付款'
      ) {
        o.row['狀態'] = '已刪除';
        await o.row.save();
        count++;
      }
    }

    return count;
  } catch (e) {
    console.error('cancelMyOrders fail:', e.message);
    return 0;
  }
}

function scheduleAutoClose(minutes) {
  if (autoCloseTimer) clearTimeout(autoCloseTimer);

  const ms = Math.max(1, Number(minutes)) * 60 * 1000;
  autoCloseAt = new Date(Date.now() + ms).toISOString();

  autoCloseTimer = setTimeout(async () => {
    isOpen = false;
    autoCloseAt = null;
    autoCloseTimer = null;
    textMenuCache = [];

    const stat = await buildStatReport().catch(() => '統計失敗');
    await pushToGroup('🔴 已自動結單！\n\n' + stat);
  }, ms);
}

function cancelAutoClose() {
  if (autoCloseTimer) clearTimeout(autoCloseTimer);
  autoCloseTimer = null;
  autoCloseAt = null;
}

async function pushToGroup(text) {
  const gid = process.env.LINE_GROUP_ID || '';
  if (!gid) return;

  try {
    await client.pushMessage(gid, {
      type: 'text',
      text
    });
  } catch (e) {
    console.error('pushToGroup fail:', e.message);
  }
}

app.get('/', (_req, res) => {
  res.send('LINE 訂餐機器人運作中 ✅');
});

app.get('/api/status', (_req, res) => {
  res.json({
    isOpen,
    autoCloseAt
  });
});

app.get('/api/menu', async (_req, res) => {
  const menu = await loadMenu();
  res.json({ menu });
});

app.post('/webhook', async (req, res) => {
  try {
    const events = req.body.events || [];

    for (const event of events) {
      if (event.type !== 'message') continue;
      if (!event.message || event.message.type !== 'text') continue;

      const uid = event.source.userId;
      const text = event.message.text.trim();

      console.log('[Webhook]', {
        type: event.source.type,
        uid,
        text
      });

      const reply = async (msg) => {
        return client.replyMessage(event.replyToken, {
          type: 'text',
          text: msg
        });
      };

      let profileName = '未知使用者';

      try {
        if (event.source.type === 'group') {
          const p = await client.getGroupMemberProfile(event.source.groupId, uid);
          profileName = p.displayName;
        } else {
          const p = await client.getProfile(uid);
          profileName = p.displayName;
        }
      } catch (e) {
        console.error('取得名稱失敗:', e.message);
      }

      saveUserToSheet(
        profileName,
        uid,
        event.source.type,
        event.source.groupId || ''
      ).catch(() => {});

      if (text === '測試') {
        await reply('收到訊息了 ✅\n官方 LINE webhook 正常運作中');
        continue;
      }

      if (text === '嗨' || text === '哈囉' || text === 'hi' || text === 'Hi') {
        await reply('嗨～我有收到你的訊息 ✅');
        continue;
      }

      if (text === '狀態') {
        let msg = isOpen ? '🟢 目前開單中' : '🔴 目前未開單';

        if (autoCloseAt) {
          msg += '\n⏰ 自動結單時間：' +
            new Date(autoCloseAt).toLocaleTimeString('zh-TW', {
              timeZone: 'Asia/Taipei'
            });
        }

        await reply(msg);
        continue;
      }

      const autoOpen = text.match(/^開單\s+(\d+)$/);

      if (autoOpen) {
        if (!isAdmin(uid)) {
          await reply('只有管理員可以開單');
          continue;
        }

        isOpen = true;
        scheduleAutoClose(Number(autoOpen[1]));

        const menu = await loadTextMenu();
        const msg =
          '🟢 開始點餐！\n' +
          '⏰ 將於 ' + autoOpen[1] + ' 分鐘後自動結單\n\n' +
          buildMenuMessage(menu);

        await reply(msg);
        continue;
      }

      if (text === '開單') {
        if (!isAdmin(uid)) {
          await reply('只有管理員可以開單');
          continue;
        }

        if (isOpen) {
          await reply('目前已開單中');
          continue;
        }

        isOpen = true;

        const menu = await loadTextMenu();
        await reply('🟢 開始點餐！\n\n' + buildMenuMessage(menu));
        continue;
      }

      if (text === '結單' || text === '收單' || text === '統計') {
        if (!isAdmin(uid)) {
          await reply('只有管理員可以結單/統計');
          continue;
        }

        isOpen = false;
        cancelAutoClose();
        textMenuCache = [];

        const stat = await buildStatReport();
        await reply('🔴 已收單！\n\n' + stat);
        continue;
      }

      if (text === '查看訂單' || text === '訂單') {
        if (!isAdmin(uid)) {
          await reply('只有管理員可以查看訂單');
          continue;
        }

        await reply(await buildDetailedReport());
        continue;
      }

      if (text === '店家單') {
        if (!isAdmin(uid)) {
          await reply('只有管理員可以查看店家單');
          continue;
        }

        await reply(await buildShopOrder());
        continue;
      }

      const paidMatch = text.match(/^(.+?)\s*已付款$/);

      if (paidMatch && isAdmin(uid)) {
        const targetName = paidMatch[1].trim();
        const count = await markPaidByName(targetName);

        if (count > 0) {
          await reply(targetName + ' 已標記付款（' + count + ' 筆）✅');
        } else {
          await reply('找不到 ' + targetName + ' 的未付款訂單');
        }

        continue;
      }

      if (text === '菜單' || text === '今日菜單') {
        if (!isOpen) {
          await reply('目前尚未開團，請等待管理員開團 🍱');
          continue;
        }

        const menu = textMenuCache.length ? textMenuCache : await loadTextMenu();
        await reply(buildMenuMessage(menu));
        continue;
      }

      if (text === '我的訂單' || text === '我點了什麼') {
        const myOrders = await getMyTextOrders(uid);

        if (!myOrders.length) {
          await reply('你今天還沒有訂單');
          continue;
        }

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
        continue;
      }

      if (text === '取消' || text === '取消訂單') {
        if (!isOpen) {
          await reply('目前未開團');
          continue;
        }

        const count = await cancelMyOrders(uid);

        if (count > 0) {
          await reply('已取消你的訂單（' + count + ' 筆）');
        } else {
          await reply('找不到可取消的訂單');
        }

        continue;
      }

      if (isOpen) {
        const menu = textMenuCache.length ? textMenuCache : await loadTextMenu();
        const parsed = parseTextOrder(text, menu);

        if (parsed) {
          const result = await saveTextOrder(profileName, uid, parsed);

          if (result.success) {
            const action = result.action === 'updated' ? '已更新' : '已收到';

            let msg = action + ' 你的訂單 ✅\n';
            msg += parsed.item;
            if (parsed.qty > 1) msg += ' ×' + parsed.qty;
            if (parsed.note) msg += '（' + parsed.note + '）';
            msg += '\n金額：$' + parsed.price * parsed.qty;
            msg += '\n\n輸入「我的訂單」查看，輸入「取消」可取消';

            await reply(msg);
          } else {
            await reply('訂單失敗，請重試或聯絡管理員\n原因：' + result.reason);
          }

          continue;
        }

        await reply('請輸入數字點餐，例如：\n2\n2不要菜\n2+3\n\n輸入「菜單」可重看菜單');
        continue;
      }

      if (event.source.type === 'user') {
        await reply(
          '我有收到你的私訊 ✅\n\n' +
          '你可以輸入：\n' +
          '測試\n' +
          '狀態\n' +
          '菜單\n' +
          '我的訂單'
        );
        continue;
      }
    }

    return res.sendStatus(200);
  } catch (err) {
    console.error('Webhook error:', err);
    return res.sendStatus(200);
  }
});

app.use((err, _req, res, _next) => {
  console.error('Global error:', err);
  res.sendStatus(200);
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log('Server running on port', PORT);
});
