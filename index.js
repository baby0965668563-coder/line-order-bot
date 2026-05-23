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
let sheetReady = false;
let currentGroupId = process.env.LINE_GROUP_ID || '';
let menuText = '';
let menuMap = {};

const hardAdmins = [
  'U8d9c82446aa9eb90d7de001cfc7ea90f',
  'Ubcfae64b443b9fad21bbc584e991b306',
  'U5c44a04efc62664bd45ec80d77be7d93',
  'Uc669eca67bf477460945f45751edd3e9'
];

const envAdmins = (process.env.ADMIN_IDS || '')
  .split(',')
  .map(s => s.trim())
  .filter(Boolean);

const admins = Array.from(new Set([...hardAdmins, ...envAdmins]));

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

async function getProfileName(event, uid) {
  try {
    if (event.source.type === 'group') {
      const p = await client.getGroupMemberProfile(event.source.groupId, uid);
      return p.displayName || '未知使用者';
    }

    if (event.source.type === 'room') {
      const p = await client.getRoomMemberProfile(event.source.roomId, uid);
      return p.displayName || '未知使用者';
    }

    const p = await client.getProfile(uid);
    return p.displayName || '未知使用者';
  } catch (e) {
    return '未知使用者';
  }
}

async function saveUserToSheet(name, userId, sourceType, groupId) {
  try {
    await authSheet();

    let sheet = doc.sheetsByTitle['Users'];

    if (!sheet) {
      sheet = await doc.addSheet({
        title: 'Users',
        headerValues: ['時間', 'LINE名稱', 'userId', '來源類型', '群組ID', '權限']
      });
    }

    await sheet.addRow({
      時間: nowTW(),
      LINE名稱: name,
      userId,
      來源類型: sourceType,
      群組ID: groupId || '',
      權限: isAdmin(userId) ? 'admin' : 'user'
    });
  } catch (e) {
    console.error('saveUserToSheet fail:', e.message);
  }
}

function parseQty(rawQty) {
  const q = String(rawQty || '').trim();
  if (q === '半' || q === '0.5' || q === '.5') return 0.5;
  return Number(q);
}

function formatQty(qty) {
  return qty === 0.5 ? '半' : String(qty);
}

function parseMenuText(text) {
  const lines = String(text || '').split('\n').map(v => v.trim()).filter(Boolean);
  const map = {};

  for (const line of lines) {
    const match = line.match(/^(\d{1,2})[\s、.．。:：-]+(.+)$/);
    if (match) map[Number(match[1])] = match[2].trim();
  }

  return map;
}

function parseNumberOrder(text) {
  const match = String(text || '')
    .trim()
    .match(/^(\d{1,2})\s*[＋+＊*]\s*(半|0\.5|\.5|\d{1,2})\s*(.*)$/);

  if (!match) return null;

  const no = Number(match[1]);
  const qty = parseQty(match[2]);
  const note = String(match[3] || '').trim();

  if (!no || !qty || qty <= 0 || qty > 20) return null;

  return { no, qty, note };
}

function parseSimpleOrder(text) {
  const match = String(text || '')
    .trim()
    .match(/^(.+?)\s*[＋+＊*]\s*(半|0\.5|\.5|\d{1,2})\s*(.*)$/);

  if (!match) return null;

  const name = match[1].trim();
  const qty = parseQty(match[2]);
  const note = String(match[3] || '').trim();

  if (!name || !qty || qty <= 0 || qty > 20) return null;

  return { name, qty, note };
}

function parseBulkOrders(text) {
  const lines = String(text || '')
    .split('\n')
    .map(v => v.trim())
    .filter(Boolean);

  let currentItem = null;
  let currentPrice = 0;
  const results = [];

  for (const line of lines) {
    if (/今天有人要訂|收錢|謝謝|下午|早上|晚上|星期/.test(line)) continue;

    const itemMatch = line.match(/^(.+?)\s*[💰$＄]\s*(\d{1,5})\s*$/);

    if (itemMatch) {
      currentItem = itemMatch[1].trim();
      currentPrice = Number(itemMatch[2]);
      continue;
    }

    const orderMatch = line.match(/^(.+?)\s*[＋+＊*]\s*(半|0\.5|\.5|\d{1,2})\s*(.*)$/);

    if (orderMatch && currentItem && currentPrice > 0) {
      const qty = parseQty(orderMatch[2]);
      if (!qty || qty <= 0) continue;

      results.push({
        name: orderMatch[1].trim(),
        qty,
        note: String(orderMatch[3] || '').trim(),
        item: currentItem,
        price: currentPrice,
        store: '手動輸入'
      });
    }
  }

  return results;
}

async function saveOrderToSheet(order, sourceUserId) {
  try {
    await authSheet();

    let sheet = doc.sheetsByTitle['Orders'];

    if (!sheet) {
      sheet = await doc.addSheet({
        title: 'Orders',
        headerValues: [
          '時間', 'LINE名稱', 'userId', '店家', '品項', '規格', '備註',
          '數量', '單價', '總價', '狀態', '付款時間', '付款方式', '訂單備註'
        ]
      });
    }

    await sheet.addRow({
      時間: nowTW(),
      LINE名稱: order.name || '',
      userId: sourceUserId || '',
      店家: order.store || '手動收單',
      品項: order.item || '未指定品項',
      規格: '',
      備註: order.note || '',
      數量: order.qty || 1,
      單價: order.price || 0,
      總價: Number(order.price || 0) * Number(order.qty || 1),
      狀態: '未付款',
      付款時間: '',
      付款方式: '',
      訂單備註: ''
    });

    return true;
  } catch (e) {
    console.error('saveOrderToSheet fail:', e.message);
    return false;
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
      .filter(r =>
        String(r['時間'] || '').startsWith(today) &&
        String(r['狀態'] || '') !== '已刪除'
      )
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

function isSpicy(note) {
  const n = String(note || '');
  if (n.includes('不辣')) return false;
  return n.includes('辣') || n.includes('🌶');
}

function isNotSpicy(note) {
  return String(note || '').includes('不辣');
}

async function buildDetailedReport() {
  const orders = await getAllOrdersToday();
  if (!orders.length) return '📋 今日尚無訂單';

  let msg = '📋 今日訂單明細\n─────────────\n';
  let totalQty = 0;
  let totalMoney = 0;

  for (const o of orders) {
    totalQty += o.qty;
    totalMoney += o.total;

    msg += `${o.name}+${formatQty(o.qty)}`;
    if (o.note) msg += ` ${o.note}`;
    if (o.item && o.item !== '未指定品項') msg += `｜${o.item}`;
    if (o.total > 0) msg += `｜$${o.total}`;
    if (o.status === '已付款') msg += ' ✅';
    msg += '\n';
  }

  msg += '─────────────\n';
  msg += `共 ${formatQty(totalQty)} 份`;
  if (totalMoney > 0) msg += `\n總金額：$${totalMoney}`;

  return msg;
}

async function buildStatReport() {
  const orders = await getAllOrdersToday();
  if (!orders.length) return '📦 已收單\n\n📊 今日尚無訂單';

  const itemCount = {};
  const spicyCount = {};
  const notSpicyCount = {};
  const userTotal = {};
  const unpaid = new Set();

  let totalQty = 0;
  let totalMoney = 0;

  for (const o of orders) {
    const item = o.item || '未指定品項';

    itemCount[item] = (itemCount[item] || 0) + o.qty;
    totalQty += o.qty;
    totalMoney += o.total;

    if (isSpicy(o.note)) spicyCount[item] = (spicyCount[item] || 0) + o.qty;
    if (isNotSpicy(o.note)) notSpicyCount[item] = (notSpicyCount[item] || 0) + o.qty;

    if (o.total > 0) userTotal[o.name || '未知'] = (userTotal[o.name || '未知'] || 0) + o.total;
    if (o.status === '未付款') unpaid.add(o.name || '未知');
  }

  let msg = '📦 已收單\n感謝大家訂購 🙏\n\n';
  msg += '📊 今日訂餐統計\n─────────────\n';
  msg += '【品項數量】\n';

  for (const item in itemCount) {
    msg += `${item} ×${formatQty(itemCount[item])}\n`;
    if (spicyCount[item]) msg += `　辣 ×${formatQty(spicyCount[item])}\n`;
    if (notSpicyCount[item]) msg += `　不辣 ×${formatQty(notSpicyCount[item])}\n`;
  }

  msg += `\n總份數：${formatQty(totalQty)} 份`;
  if (totalMoney > 0) msg += `\n總金額：$${totalMoney}`;

  if (Object.keys(userTotal).length) {
    msg += '\n\n【個人金額】\n';
    for (const name in userTotal) {
      msg += `${name}：$${userTotal[name]}\n`;
    }
  }

  if (unpaid.size) {
    msg += '\n⚠️ 未付款：' + Array.from(unpaid).join('、');
  }

  return msg;
}

async function buildShopOrder() {
  const orders = await getAllOrdersToday();
  if (!orders.length) return '今日尚無訂單';

  const itemCount = {};
  const spicyCount = {};
  const notSpicyCount = {};

  let totalQty = 0;
  let totalMoney = 0;

  for (const o of orders) {
    const item = o.item || '未指定品項';

    itemCount[item] = (itemCount[item] || 0) + o.qty;
    totalQty += o.qty;
    totalMoney += o.total;

    if (isSpicy(o.note)) spicyCount[item] = (spicyCount[item] || 0) + o.qty;
    if (isNotSpicy(o.note)) notSpicyCount[item] = (notSpicyCount[item] || 0) + o.qty;
  }

  let msg = '您好，今天訂購如下：\n\n';

  for (const item in itemCount) {
    msg += `${item} x${formatQty(itemCount[item])}\n`;
    if (spicyCount[item]) msg += `辣 x${formatQty(spicyCount[item])}\n`;
    if (notSpicyCount[item]) msg += `不辣 x${formatQty(notSpicyCount[item])}\n`;
  }

  msg += `\n總數：${formatQty(totalQty)}份`;
  if (totalMoney > 0) msg += `\n總金額：${totalMoney}元`;
  msg += '\n\n麻煩您，謝謝～';

  return msg;
}

async function markPaidByName(name) {
  const orders = await getAllOrdersToday();
  let count = 0;

  for (const o of orders) {
    if (o.name.trim() === name.trim() && o.status === '未付款') {
      o.row['狀態'] = '已付款';
      o.row['付款時間'] = nowTW();
      o.row['付款方式'] = '現金';
      await o.row.save();
      count++;
    }
  }

  return count;
}

async function clearTodayOrders() {
  const orders = await getAllOrdersToday();
  let count = 0;

  for (const o of orders) {
    o.row['狀態'] = '已刪除';
    await o.row.save();
    count++;
  }

  return count;
}

function parseCloseTime(text) {
  const m = String(text || '').match(/(\d{1,2})[：:.]?(\d{2})/);
  if (!m) return null;

  const hour = Number(m[1]);
  const minute = Number(m[2]);

  if (hour < 0 || hour > 23 || minute < 0 || minute > 59) return null;

  return { hour, minute };
}

function msUntil(hour, minute) {
  const now = new Date();
  const target = new Date();

  target.setHours(hour, minute, 0, 0);
  if (target <= now) target.setDate(target.getDate() + 1);

  return target - now;
}

function cancelAutoClose() {
  if (autoCloseTimer) clearTimeout(autoCloseTimer);
  autoCloseTimer = null;
  autoCloseAt = null;
}

function scheduleAutoClose(hour, minute) {
  cancelAutoClose();

  const ms = msUntil(hour, minute);
  autoCloseAt = new Date(Date.now() + ms).toISOString();

  autoCloseTimer = setTimeout(async () => {
    isOpen = false;
    autoCloseAt = null;
    autoCloseTimer = null;

    const stat = await buildStatReport().catch(() => '📦 已收單\n統計失敗');
    await pushToGroup(stat);
  }, ms);
}

function scheduleAutoCloseByMinutes(minutes) {
  cancelAutoClose();

  const ms = Math.max(1, Number(minutes)) * 60 * 1000;
  autoCloseAt = new Date(Date.now() + ms).toISOString();

  autoCloseTimer = setTimeout(async () => {
    isOpen = false;
    autoCloseAt = null;
    autoCloseTimer = null;

    const stat = await buildStatReport().catch(() => '📦 已收單\n統計失敗');
    await pushToGroup(stat);
  }, ms);
}

async function pushToGroup(text) {
  const gid = currentGroupId || process.env.LINE_GROUP_ID || '';
  if (!gid) return;

  try {
    await client.pushMessage(gid, { type: 'text', text });
  } catch (e) {
    console.error('pushToGroup fail:', e.message);
  }
}

app.get('/', (_req, res) => {
  res.send('LINE 訂餐機器人運作中 ✅');
});

app.get('/api/status', (_req, res) => {
  res.json({ isOpen, autoCloseAt, currentGroupId, menuText, menuMap });
});

app.post('/webhook', async (req, res) => {
  res.sendStatus(200);

  try {
    const events = req.body.events || [];

    for (const event of events) {
      if (event.type !== 'message') continue;
      if (!event.message || event.message.type !== 'text') continue;

      const uid = event.source.userId;
      const text = event.message.text.trim();

      if (event.source.groupId) currentGroupId = event.source.groupId;
      if (event.source.roomId) currentGroupId = event.source.roomId;

      const reply = async (msg) => {
        try {
          await client.replyMessage(event.replyToken, { type: 'text', text: msg });
        } catch (e) {
          console.error('reply fail:', e.message);
        }
      };

      const profileName = await getProfileName(event, uid);

      saveUserToSheet(
        profileName,
        uid,
        event.source.type,
        event.source.groupId || event.source.roomId || ''
      ).catch(() => {});

      if (text === '測試') {
        await reply('收到訊息了 ✅\nLINE webhook 正常運作中');
        continue;
      }

      if (/^\/?(狀態|status)$/.test(text)) {
        if (!isAdmin(uid)) continue;

        let msg = isOpen ? '🟢 目前開單中' : '🔴 目前未開單';

        if (autoCloseAt) {
          msg += '\n⏰ 自動收單時間：' +
            new Date(autoCloseAt).toLocaleTimeString('zh-TW', {
              timeZone: 'Asia/Taipei',
              hour: '2-digit',
              minute: '2-digit'
            });
        }

        const orders = await getAllOrdersToday();
        const total = orders.reduce((sum, o) => sum + o.qty, 0);
        msg += `\n目前訂單：${orders.length} 筆，共 ${formatQty(total)} 份`;

        await reply(msg);
        continue;
      }

      if (/^\/?(菜單|menu)/.test(text)) {
        if (!isAdmin(uid)) continue;

        const content = text.replace(/^\/?(菜單|menu)\s*/, '').trim();

        if (content) {
          menuText = content;
          menuMap = parseMenuText(content);
          await reply('✅ 菜單已設定\n\n' + menuText);
        } else if (menuText) {
          await reply(menuText);
        } else {
          await reply('今日尚未設定菜單');
        }

        continue;
      }

      if (/^\/?(開單|開)/.test(text)) {
        if (!isAdmin(uid)) {
          await reply('只有管理員可以開單');
          continue;
        }

        const timeText = text.replace(/^\/?(開單|開)\s*/, '').trim();
        isOpen = true;

        let msg = '🟢 已開單，可以開始加單';

        if (timeText) {
          const t = parseCloseTime(timeText);

          if (t) {
            scheduleAutoClose(t.hour, t.minute);
            const closeTime = String(t.hour).padStart(2, '0') + ':' + String(t.minute).padStart(2, '0');
            msg += `\n🕐 將於 ${closeTime} 自動收單`;
          } else if (/^\d+$/.test(timeText)) {
            scheduleAutoCloseByMinutes(Number(timeText));
            msg += `\n⏰ 將於 ${timeText} 分鐘後自動收單`;
          } else {
            await reply('時間格式錯誤，請輸入：/開單 19:30 或 /開 1930');
            continue;
          }
        } else {
          cancelAutoClose();
        }

        if (menuText) msg += '\n\n' + menuText;

        await reply(msg);
        continue;
      }

      if (/^\/?(收單|結單|關|統計)$/.test(text)) {
        if (!isAdmin(uid)) {
          await reply('只有管理員可以收單/結單/統計');
          continue;
        }

        isOpen = false;
        cancelAutoClose();

        await reply(await buildStatReport());
        continue;
      }

      if (/^\/?(查單|查|訂單|查看訂單)$/.test(text)) {
        if (!isAdmin(uid)) {
          await reply('只有管理員可以查看訂單');
          continue;
        }

        await reply(await buildDetailedReport());
        continue;
      }

      if (/^\/?(店家|店|店家單)$/.test(text)) {
        if (!isAdmin(uid)) {
          await reply('只有管理員可以查看店家單');
          continue;
        }

        await reply(await buildShopOrder());
        continue;
      }

      if (/^\/?(清單|清空|重置)$/.test(text)) {
        if (!isAdmin(uid)) {
          await reply('只有管理員可以清空');
          continue;
        }

        const count = await clearTodayOrders();
        isOpen = false;
        cancelAutoClose();

        await reply(`已清空今日訂單（${count} 筆）`);
        continue;
      }

      const paidMatch = text.match(/^(.+?)\s*已付款$/);

      if (paidMatch && isAdmin(uid)) {
        const targetName = paidMatch[1].trim();
        const count = await markPaidByName(targetName);

        await reply(count > 0
          ? `${targetName} 已標記付款（${count} 筆）✅`
          : `找不到 ${targetName} 的未付款訂單`
        );

        continue;
      }

      if (isOpen) {
        const bulk = parseBulkOrders(text);

        if (bulk.length) {
          let ok = 0;

          for (const order of bulk) {
            const success = await saveOrderToSheet(order, 'legacy_' + order.name);
            if (success) ok++;
          }

          await reply(`✅ 已匯入 ${ok} 筆訂單`);
          continue;
        }

        const numberOrder = parseNumberOrder(text);

        if (numberOrder) {
          const itemName = menuMap[numberOrder.no];
          if (!itemName) continue;

          await saveOrderToSheet(
            {
              name: profileName,
              qty: numberOrder.qty,
              note: numberOrder.note,
              item: itemName,
              store: '動態菜單',
              price: 0
            },
            uid
          );

          continue;
        }

        const simple = parseSimpleOrder(text);

        if (simple) {
          await saveOrderToSheet(
            {
              name: simple.name,
              qty: simple.qty,
              note: simple.note,
              item: '未指定品項',
              store: '手動收單',
              price: 0
            },
            uid
          );

          continue;
        }

        continue;
      }

      continue;
    }
  } catch (err) {
    console.error('Webhook error:', err);
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
