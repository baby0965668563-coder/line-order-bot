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
    console.error('取得名稱失敗:', e.message);
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
      userId: userId,
      來源類型: sourceType,
      群組ID: groupId || '',
      權限: isAdmin(userId) ? 'admin' : 'user'
    });
  } catch (e) {
    console.error('saveUserToSheet fail:', e.message);
  }
}

function parseSimpleOrder(text) {
  const clean = String(text || '').trim();
  const match = clean.match(/^(.+?)\s*[＋+＊*]\s*(\d+)\s*(.*)$/);

  if (!match) return null;

  const name = match[1].trim();
  const qty = Number(match[2]);
  const note = String(match[3] || '').trim();

  if (!name || !qty || qty <= 0) return null;

  return { name, qty, note };
}

function parseLegacyOrders(text) {
  const lines = String(text || '')
    .split('\n')
    .map(v => v.trim())
    .filter(Boolean);

  let currentItem = null;
  const results = [];

  for (const line of lines) {
    const itemMatch = line.match(/^(.+?)\s*\$(\d+)/);

    if (itemMatch) {
      currentItem = {
        item: itemMatch[1].trim(),
        price: Number(itemMatch[2])
      };
      continue;
    }

    const orderMatch = line.match(/^(.+?)\s*[＋+＊*]\s*(\d+)\s*(.*)$/);

    if (orderMatch && currentItem) {
      results.push({
        name: orderMatch[1].trim(),
        qty: Number(orderMatch[2]),
        note: String(orderMatch[3] || '').trim(),
        item: currentItem.item,
        price: currentItem.price
      });
    }
  }

  return results;
}

async function saveSimpleOrderToSheet(order, sourceUserId) {
  try {
    await authSheet();

    let sheet = doc.sheetsByTitle['Orders'];

    if (!sheet) {
      sheet = await doc.addSheet({
        title: 'Orders',
        headerValues: [
          '時間',
          'LINE名稱',
          'userId',
          '店家',
          '品項',
          '規格',
          '備註',
          '數量',
          '單價',
          '總價',
          '狀態',
          '付款時間',
          '付款方式',
          '訂單備註'
        ]
      });
    }

    await sheet.addRow({
      時間: nowTW(),
      LINE名稱: order.name,
      userId: sourceUserId || '',
      店家: order.store || '手動收單',
      品項: order.item || '未指定品項',
      規格: '',
      備註: order.note || '',
      數量: order.qty,
      單價: order.price || 0,
      總價: order.price ? order.price * order.qty : 0,
      狀態: '未付款',
      付款時間: '',
      付款方式: '',
      訂單備註: ''
    });

    return { success: true };
  } catch (e) {
    console.error('saveSimpleOrderToSheet fail:', e.message);
    return { success: false, reason: e.message };
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

  let msg = '📋 今日訂單明細\n';
  msg += '─────────────\n';

  let totalQty = 0;
  let totalMoney = 0;

  for (const o of orders) {
    totalQty += o.qty;
    totalMoney += o.total;

    msg += `${o.name} +${o.qty}`;
    if (o.note) msg += ` ${o.note}`;
    if (o.item && o.item !== '未指定品項') msg += `｜${o.item}`;
    if (o.total > 0) msg += `｜$${o.total}`;
    if (o.status === '已付款') msg += ' ✅';
    msg += '\n';
  }

  msg += '─────────────\n';
  msg += `共 ${totalQty} 份`;

  if (totalMoney > 0) {
    msg += `\n總金額：$${totalMoney}`;
  }

  return msg;
}

async function buildStatReport() {
  const orders = await getAllOrdersToday();

  if (!orders.length) {
    return '📦 已收單\n\n📊 今日尚無訂單';
  }

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

    if (isSpicy(o.note)) {
      spicyCount[item] = (spicyCount[item] || 0) + o.qty;
    }

    if (isNotSpicy(o.note)) {
      notSpicyCount[item] = (notSpicyCount[item] || 0) + o.qty;
    }

    const name = o.name || '未知';
    userTotal[name] = (userTotal[name] || 0) + o.total;

    if (o.status === '未付款') unpaid.add(name);
  }

  let msg = '📦 已收單\n感謝大家訂購 🙏\n\n';
  msg += '📊 今日訂餐統計\n';
  msg += '─────────────\n';

  msg += '【品項數量】\n';
  for (const item in itemCount) {
    msg += `${item} ×${itemCount[item]}\n`;

    if (spicyCount[item]) {
      msg += `　辣 ×${spicyCount[item]}\n`;
    }

    if (notSpicyCount[item]) {
      msg += `　不辣 ×${notSpicyCount[item]}\n`;
    }
  }

  msg += `\n總份數：${totalQty} 份`;

  if (totalMoney > 0) {
    msg += `\n總金額：$${totalMoney}`;
  }

  if (Object.values(userTotal).some(v => v > 0)) {
    msg += '\n\n【個人金額】\n';

    for (const name in userTotal) {
      if (userTotal[name] > 0) {
        msg += `${name}：$${userTotal[name]}\n`;
      }
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

    if (isSpicy(o.note)) {
      spicyCount[item] = (spicyCount[item] || 0) + o.qty;
    }

    if (isNotSpicy(o.note)) {
      notSpicyCount[item] = (notSpicyCount[item] || 0) + o.qty;
    }
  }

  let msg = '您好，今天訂購如下：\n\n';

  for (const item in itemCount) {
    if (item === '未指定品項') {
      msg += `總共 ${itemCount[item]} 份\n`;
    } else {
      msg += `${item} x${itemCount[item]}\n`;
    }

    if (spicyCount[item]) {
      msg += `辣 x${spicyCount[item]}\n`;
    }

    if (notSpicyCount[item]) {
      msg += `不辣 x${notSpicyCount[item]}\n`;
    }
  }

  msg += `\n總數：${totalQty}份`;

  if (totalMoney > 0) {
    msg += `\n總金額：${totalMoney}元`;
  }

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

function getMillisecondsUntil(timeText) {
  const match = String(timeText || '').match(/^(\d{1,2}):(\d{2})$/);
  if (!match) return null;

  const hour = Number(match[1]);
  const minute = Number(match[2]);

  if (hour < 0 || hour > 23 || minute < 0 || minute > 59) return null;

  const now = new Date();
  const close = new Date();

  close.setHours(hour);
  close.setMinutes(minute);
  close.setSeconds(0);
  close.setMilliseconds(0);

  if (close <= now) {
    close.setDate(close.getDate() + 1);
  }

  return close - now;
}

function scheduleAutoCloseByTime(timeText) {
  if (autoCloseTimer) clearTimeout(autoCloseTimer);

  const ms = getMillisecondsUntil(timeText);
  if (ms === null) return false;

  const closeDate = new Date(Date.now() + ms);
  autoCloseAt = closeDate.toISOString();

  autoCloseTimer = setTimeout(async () => {
    isOpen = false;
    autoCloseAt = null;
    autoCloseTimer = null;

    const stat = await buildStatReport().catch(() => '📦 已收單\n統計失敗');
    await pushToGroup(stat);
  }, ms);

  return true;
}

function scheduleAutoCloseByMinutes(minutes) {
  if (autoCloseTimer) clearTimeout(autoCloseTimer);

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

function cancelAutoClose() {
  if (autoCloseTimer) clearTimeout(autoCloseTimer);
  autoCloseTimer = null;
  autoCloseAt = null;
}

async function pushToGroup(text) {
  const gid = process.env.LINE_GROUP_ID || '';
  if (!gid) {
    console.error('LINE_GROUP_ID 未設定，無法主動推送收單訊息');
    return;
  }

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
  res.send('LINE 訂單小幫手運作中 ✅');
});

app.get('/api/status', (_req, res) => {
  res.json({
    isOpen,
    autoCloseAt
  });
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

      if (text === '狀態') {
        let msg = isOpen ? '🟢 目前開單中' : '🔴 目前未開單';

        if (autoCloseAt) {
          msg += '\n⏰ 自動收單時間：' +
            new Date(autoCloseAt).toLocaleTimeString('zh-TW', {
              timeZone: 'Asia/Taipei',
              hour: '2-digit',
              minute: '2-digit'
            });
        }

        await reply(msg);
        continue;
      }

      const slashOpen = text.match(/^\/開單\s+(\d{1,2}:\d{2})$/);

      if (slashOpen) {
        if (!isAdmin(uid)) {
          await reply('只有管理員可以開單');
          continue;
        }

        const closeTime = slashOpen[1];
        const ok = scheduleAutoCloseByTime(closeTime);

        if (!ok) {
          await reply('時間格式錯誤，請輸入：/開單 19:30');
          continue;
        }

        isOpen = true;

        await reply(
          '📢 已開放訂購\n' +
          `🕢 將於 ${closeTime} 自動收單\n\n` +
          '大家可以直接輸入：\n' +
          '慧玲+1\n' +
          '阿明+2辣\n' +
          '麗麗*1不辣'
        );

        continue;
      }

      const minuteOpen = text.match(/^開單\s+(\d+)$/);

      if (minuteOpen) {
        if (!isAdmin(uid)) {
          await reply('只有管理員可以開單');
          continue;
        }

        isOpen = true;
        scheduleAutoCloseByMinutes(Number(minuteOpen[1]));

        await reply(
          '📢 已開放訂購\n' +
          `⏰ 將於 ${minuteOpen[1]} 分鐘後自動收單\n\n` +
          '大家可以直接輸入：\n' +
          '慧玲+1\n' +
          '阿明+2辣\n' +
          '麗麗*1不辣'
        );

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
        cancelAutoClose();

        await reply(
          '📢 已開放訂購\n\n' +
          '大家可以直接輸入：\n' +
          '慧玲+1\n' +
          '阿明+2辣\n' +
          '麗麗*1不辣'
        );

        continue;
      }

      if (text === '收單' || text === '結單' || text === '統計') {
        if (!isAdmin(uid)) {
          await reply('只有管理員可以收單/結單/統計');
          continue;
        }

        isOpen = false;
        cancelAutoClose();

        const stat = await buildStatReport();
        await reply(stat);
        continue;
      }

      if (text === '查單' || text === '/查單' || text === '查看訂單' || text === '訂單') {
        if (!isAdmin(uid)) {
          await reply('只有管理員可以查看訂單');
          continue;
        }

        await reply(await buildDetailedReport());
        continue;
      }

      if (text === '店家' || text === '/店家' || text === '店家單') {
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
          await reply(`${targetName} 已標記付款（${count} 筆）✅`);
        } else {
          await reply(`找不到 ${targetName} 的未付款訂單`);
        }

        continue;
      }

      if (isOpen) {
        if (text.includes('$')) {
          const parsedList = parseLegacyOrders(text);

          if (parsedList.length) {
            let ok = 0;

            for (const p of parsedList) {
              const result = await saveSimpleOrderToSheet(
                {
                  name: p.name,
                  item: p.item,
                  store: '手動輸入',
                  price: p.price,
                  qty: p.qty,
                  note: p.note
                },
                'legacy_' + p.name
              );

              if (result.success) ok++;
            }

            await reply(`✅ 已匯入 ${ok} 筆訂單`);
            continue;
          }
        }

        const simple = parseSimpleOrder(text);

        if (simple) {
          await saveSimpleOrderToSheet(
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

      const maybeOrder = parseSimpleOrder(text);
      if (maybeOrder) {
        continue;
      }

      continue;
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
