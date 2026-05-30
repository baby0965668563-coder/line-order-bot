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
app.use('/api', express.json());
app.use('/admin', express.json());

const client = new line.Client(config);
const doc = new GoogleSpreadsheet(process.env.SHEET_ID);

let isOpen = false;
let allText = '';
let autoCloseAt = null;
let autoCloseGroupId = '';
let autoClosed = false;
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

function getTaipeiNow() {
  return new Date(new Date().toLocaleString('en-US', { timeZone: 'Asia/Taipei' }));
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
  } catch {
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

function clean(text) {
  return String(text || '')
    .replace(/[。.,，、!！?？:：;；"'（）()【】\[\]{}<>《》\s]/g, '')
    .trim();
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
  const lines = String(text || '')
    .split('\n')
    .map(v => v.trim())
    .filter(Boolean);

  const map = {};

  for (const line of lines) {
    const match = line.match(/^(\d{1,2})[\s、.．。:：-]+(.+)$/);
    if (match) map[Number(match[1])] = match[2].trim();
  }

  return map;
}

function splitNames(text) {
  return String(text || '')
    .split(/[、,，\s]+/)
    .map(v => v.trim())
    .filter(Boolean);
}

function isTrashLine(line) {
  return /今天有人要訂|收錢|謝謝|下午|早上|晚上|星期|明天|湊到|結單|沒加到|不要怪我|截止|滿\d+/.test(line);
}

function parseOrders(text) {
  const lines = String(text || '').split('\n').map(l => l.trim()).filter(Boolean);

  let currentItem = '';
  let currentPrice = 0;
  let lastPrice = 0;
  let pendingInlineItem = '';
  let priceMap = {};

  const itemCount = {};
  const userTotal = {};
  const details = [];

  function cleanItemName(text) {
    return clean(text)
      .replace(/\d+顆/g, '')
      .replace(/^[$＄💰]\d+$/, '')
      .trim();
  }

  function add(item, price, name, qty = 1, note = '') {
    item = cleanItemName(item);
    name = clean(name);

    if (!item || !name || !qty) return;

    const finalItem = note ? `${item}（${note}）` : item;
    const total = Number(price || 0) * Number(qty || 1);

    itemCount[finalItem] = (itemCount[finalItem] || 0) + qty;
    userTotal[name] = (userTotal[name] || 0) + total;

    details.push({
      item: finalItem,
      rawItem: item,
      name,
      qty,
      note,
      price: Number(price || 0),
      total
    });
  }

  function getPrice(rawQty) {
    if (priceMap[rawQty] !== undefined) return priceMap[rawQty];
    if (rawQty === '半' && priceMap['0.5'] !== undefined) return priceMap['0.5'];
    return currentPrice || lastPrice || 0;
  }

  for (let line of lines) {
    if (!line) continue;

    if (isTrashLine(line) && !/[💰$＄]\s*\d/.test(line)) {
      continue;
    }

    const priceOnly = line.match(/^[💰$＄]\s*(\d{1,5})$/);
    if (priceOnly) {
      lastPrice = Number(priceOnly[1]);
      if (currentItem) currentPrice = lastPrice;
      continue;
    }

    const priceTable = line.match(/^(半|0\.5|1)\s*[💰$＄]\s*(\d{1,5})$/);
    if (priceTable) {
      priceMap[priceTable[1]] = Number(priceTable[2]);
      continue;
    }

    const itemSymbol = line.match(/^(.+?)\s*[💰$＄]\s*(\d{1,5})$/);
    if (itemSymbol) {
      currentItem = itemSymbol[1].trim();
      currentPrice = Number(itemSymbol[2]);
      lastPrice = currentPrice;
      pendingInlineItem = '';
      priceMap = {};
      continue;
    }

    const pureItemAfterPrice = line.match(/^([^\d+＋*＊$＄💰]+)$/);
    if (pureItemAfterPrice && lastPrice > 0) {
      const maybeItem = pureItemAfterPrice[1].trim();

      if (!isTrashLine(maybeItem) && maybeItem.length <= 12) {
        currentItem = maybeItem;
        currentPrice = lastPrice;
        pendingInlineItem = '';
        continue;
      }
    }

const inlineGroup = line.match(/^(.+?)\s*[+＋*＊]\s*(半|0\.5|\.5|\d+)\s*(.+)$/);
    if (inlineGroup && (currentPrice > 0 || lastPrice > 0)) {
      const item = inlineGroup[1].trim();
      const qtyRaw = inlineGroup[2];
      const namesText = inlineGroup[3].trim();
      const names = splitNames(namesText);
      const price = getPrice(qtyRaw);

      if (names.length) {
        currentItem = item;
        currentPrice = price || lastPrice || currentPrice;
        lastPrice = currentPrice || lastPrice;

        for (const n of names) {
          add(currentItem, currentPrice, n, 1, '');
        }

        continue;
      }
    }

    const orderMatch = line.match(/^(.+?)[+＋*＊]\s*(半|0\.5|\.5|\d+)(.*)$/);
    if (orderMatch && currentItem) {
      const name = orderMatch[1].trim();
      const rawQty = orderMatch[2];
      const extra = String(orderMatch[3] || '').trim();
      const price = getPrice(rawQty);
      const qty = parseQty(rawQty);

      add(currentItem, price, name, qty, extra);
      continue;
    }

    const numberOrder = line.match(/^(\d{1,2})\s*[+＋*＊]\s*(半|0\.5|\.5|\d{1,2})(.*)$/);
    if (numberOrder && menuMap[Number(numberOrder[1])]) {
      const itemName = menuMap[Number(numberOrder[1])];
      const qty = parseQty(numberOrder[2]);
      const note = String(numberOrder[3] || '').trim();
      add(itemName, 0, '未命名', qty, note);
      continue;
    }

    const inlineSymbol = line.match(/^(.+?)\s*[💰$＄]\s*(\d{1,5})\s*([^\d\s]+)$/);
    if (inlineSymbol) {
      add(inlineSymbol[1], Number(inlineSymbol[2]), inlineSymbol[3], 1);
      pendingInlineItem = '';
      continue;
    }

    const inlineNoSymbol = line.match(/^(.+?)(\d{2,5})([^\d\s]+)$/);
    if (inlineNoSymbol && !/[+＋*＊]/.test(line)) {
      add(inlineNoSymbol[1], Number(inlineNoSymbol[2]), inlineNoSymbol[3], 1);
      pendingInlineItem = '';
      continue;
    }

    const noSymbolNoName = line.match(/^(.+?)(\d{2,5})$/);
    if (noSymbolNoName && !/[+＋*＊]/.test(line) && !/顆/.test(line)) {
      currentItem = noSymbolNoName[1].trim();
      currentPrice = Number(noSymbolNoName[2]);
      lastPrice = currentPrice;
      pendingInlineItem = '';
      continue;
    }

    const priceNameOnly = line.match(/^(\d{2,5})\s*([^\d\s]+)$/);
    if (priceNameOnly && pendingInlineItem) {
      add(pendingInlineItem, Number(priceNameOnly[1]), priceNameOnly[2], 1);
      pendingInlineItem = '';
      continue;
    }

    const priceOnlyPlain = line.match(/^(\d{2,5})$/);
    if (priceOnlyPlain && pendingInlineItem) {
      currentItem = pendingInlineItem;
      currentPrice = Number(priceOnlyPlain[1]);
      lastPrice = currentPrice;
      pendingInlineItem = '';
      continue;
    }

    if (!/[+＋*＊]/.test(line)) {
      if (/顆/.test(line)) {
        currentItem = line;
        currentPrice = lastPrice || 0;
        pendingInlineItem = line;
        continue;
      }

      if (currentItem && (currentPrice > 0 || lastPrice > 0) && !isTrashLine(line)) {
        const names = splitNames(line);

        for (const n of names) {
          add(currentItem, currentPrice || lastPrice, n, 1);
        }

        continue;
      }

      continue;
    }
  }

  return { itemCount, userTotal, details };
}

function formatResult(itemCount, userTotal) {
  let text = '📦 已收單\n感謝大家訂購 🙏\n\n';
  text += '📊 今日訂餐統計\n';
  text += '─────────────\n';

  text += '【品項數量】\n';
  for (let item in itemCount) {
    if (itemCount[item] > 0) {
      text += `${item} x${formatQty(itemCount[item])}\n`;
    }
  }

  text += '\n【個人金額】\n';
  for (let user in userTotal) {
    text += `${user}：$${userTotal[user]}\n`;
  }

  const total = Object.values(userTotal).reduce((a, b) => a + b, 0);
  text += `\n💰 總金額：$${total}`;

  return text;
}

function formatShopOrder(itemCount, userTotal) {
  let orderText = '您好，今天訂購如下：\n\n';
  let totalCount = 0;

  for (let item in itemCount) {
    const qty = itemCount[item];

    if (qty > 0) {
      orderText += `${item} x${formatQty(qty)}\n`;
      totalCount += qty;
    }
  }

  const totalMoney = Object.values(userTotal).reduce((a, b) => a + b, 0);

  orderText += `\n總數：${formatQty(totalCount)}份`;
  orderText += `\n總金額：${totalMoney}元`;
  orderText += '\n\n麻煩您，謝謝～';

  return orderText;
}

async function saveParsedOrdersToSheet(details) {
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

    for (const o of details) {
      await sheet.addRow({
        時間: nowTW(),
        LINE名稱: o.name || '',
        userId: '',
        店家: '文字統計',
        品項: o.rawItem || o.item || '',
        規格: '',
        備註: o.note || '',
        數量: o.qty || 1,
        單價: o.price || 0,
        總價: o.total || 0,
        狀態: '未付款',
        付款時間: '',
        付款方式: '',
        訂單備註: ''
      });
    }
  } catch (e) {
    console.error('saveParsedOrdersToSheet fail:', e.message);
  }
}

function parseCloseTime(text) {
  const m = String(text || '').match(/(\d{1,2})[：:.]?(\d{2})/);
  if (!m) return null;

  const hour = Number(m[1]);
  const minute = Number(m[2]);

  if (hour < 0 || hour > 23 || minute < 0 || minute > 59) return null;

  return { hour, minute };
}

function getNextCloseDate(hour, minute) {
  const now = getTaipeiNow();
  const target = getTaipeiNow();

  target.setHours(hour, minute, 0, 0);

  if (target.getTime() < now.getTime() - 90 * 1000) {
    target.setDate(target.getDate() + 1);
  }

  return target;
}

function formatAutoCloseTime() {
  if (!autoCloseAt) return '';

  return new Date(autoCloseAt).toLocaleString('zh-TW', {
    timeZone: 'Asia/Taipei',
    hour12: false,
    hour: '2-digit',
    minute: '2-digit'
  });
}

function setAutoClose(hour, minute) {
  const target = getNextCloseDate(hour, minute);

  autoCloseAt = target.toISOString();
  autoCloseGroupId = currentGroupId || process.env.LINE_GROUP_ID || '';
  autoClosed = false;

  console.log('auto close set:', {
    autoCloseAt,
    autoCloseGroupId
  });
}

function cancelAutoClose() {
  autoCloseAt = null;
  autoCloseGroupId = '';
  autoClosed = false;
}

async function executeCloseByTimer() {
  if (!isOpen) return;
  if (!autoCloseAt) return;
  if (autoClosed) return;

  const now = getTaipeiNow().getTime();
  const target = new Date(autoCloseAt).getTime();

  if (now < target) return;

  autoClosed = true;
  isOpen = false;

  const result = parseOrders(allText);
  saveParsedOrdersToSheet(result.details).catch(() => {});

  const msg = formatResult(result.itemCount, result.userTotal);
  await pushToGroup(msg);

  autoCloseAt = null;
  autoCloseGroupId = '';
}

async function pushToGroup(text) {
  const gid = autoCloseGroupId || currentGroupId || process.env.LINE_GROUP_ID || '';

  if (!gid) {
    console.error('沒有 LINE_GROUP_ID，無法自動推送收單訊息');
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

setInterval(() => {
  executeCloseByTimer().catch(e => {
    console.error('auto close interval fail:', e.message);
  });
}, 10 * 1000);

app.get('/', (_req, res) => {
  res.send('LINE 訂餐統計機器人運作中 ✅');
});

app.get('/api/status', async (_req, res) => {
  await executeCloseByTimer().catch(() => {});

  res.json({
    isOpen,
    autoCloseAt,
    autoCloseTimeTW: formatAutoCloseTime(),
    autoCloseGroupId,
    autoClosed,
    currentGroupId,
    menuText,
    allTextLength: allText.length
  });
});

app.post('/webhook', async (req, res) => {
  res.sendStatus(200);

  try {
    await executeCloseByTimer();

    const events = req.body.events || [];

    for (const event of events) {
      if (!event || event.type !== 'message') continue;
      if (!event.message || event.message.type !== 'text') continue;

      const uid = event.source.userId;
      const text = event.message.text.trim();

      if (event.source.groupId) currentGroupId = event.source.groupId;
      if (event.source.roomId) currentGroupId = event.source.roomId;

      const reply = async (msg) => {
        try {
          await client.replyMessage(event.replyToken, {
            type: 'text',
            text: msg
          });
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

      if (/^\/?(狀態|status)$/.test(text)) {
        if (!isAdmin(uid)) continue;

        let msg = isOpen ? '🟢 目前開單中' : '🔴 目前未開單';

        if (autoCloseAt) {
          msg += '\n⏰ 自動收單時間：' + formatAutoCloseTime();
        }

        const result = parseOrders(allText);
        const total = Object.values(result.itemCount).reduce((a, b) => a + b, 0);
        msg += `\n目前統計：${formatQty(total)} 份`;

        await reply(msg);
        continue;
      }

      if (/^\/?(開單|開)$|^\/?(開單|開)\s+/.test(text)) {
        if (!isAdmin(uid)) {
          await reply('只有管理員可以開單');
          continue;
        }

        const timeText = text.replace(/^\/?(開單|開)\s*/, '').trim();

        if (isOpen) {
          await reply('目前已開單中，不會清空訂單');
          continue;
        }

        isOpen = true;
        allText = '';
        cancelAutoClose();

        let msg = '🟢 已開單，可以開始加單';

        if (timeText) {
          const t = parseCloseTime(timeText);

          if (t) {
            setAutoClose(t.hour, t.minute);
            const closeTime = String(t.hour).padStart(2, '0') + ':' + String(t.minute).padStart(2, '0');
            msg += `\n🕐 將於 ${closeTime} 自動收單`;
          } else {
            await reply('時間格式錯誤，請輸入：開單 19:30 或 /開 1930');
            continue;
          }
        }

        if (menuText) msg += '\n\n' + menuText;

        await reply(msg);
        continue;
      }

      const autoCloseMatch = text.match(/^\/?(收單|結單)\s*\/?\s*(\d{1,2})[:：]?(\d{2})$/);

      if (autoCloseMatch) {
        if (!isAdmin(uid)) {
          await reply('只有管理員可以設定自動收單');
          continue;
        }

        const hour = Number(autoCloseMatch[2]);
        const minute = Number(autoCloseMatch[3]);

        if (hour < 0 || hour > 23 || minute < 0 || minute > 59) {
          await reply('時間格式錯誤');
          continue;
        }

        setAutoClose(hour, minute);

        const hh = String(hour).padStart(2, '0');
        const mm = String(minute).padStart(2, '0');

        await reply(`⏰ 已設定 ${hh}:${mm} 自動收單`);
        continue;
      }

      if (/^\/?(清空|重置|清單)$/.test(text)) {
        if (!isAdmin(uid)) {
          await reply('只有管理員可以清空');
          continue;
        }

        allText = '';
        isOpen = false;
        cancelAutoClose();

        await reply('已清空訂單');
        continue;
      }

      if (/^\/?(查單|查|訂單|查看訂單)$/.test(text)) {
        if (!isAdmin(uid)) {
          await reply('只有管理員可以查看訂單');
          continue;
        }

        const result = parseOrders(allText);
        await reply(formatResult(result.itemCount, result.userTotal));
        continue;
      }

      if (/^\/?(店家|店|店家單)$/.test(text)) {
        if (!isAdmin(uid)) {
          await reply('只有管理員可以查看店家單');
          continue;
        }

        const result = parseOrders(allText);
        await reply(formatShopOrder(result.itemCount, result.userTotal));
        continue;
      }

      if (/^\/?(結單|收單|統計|關)$/.test(text)) {
        if (!isAdmin(uid)) {
          await reply('只有管理員可以結單 / 統計');
          continue;
        }

        const result = parseOrders(allText);

        isOpen = false;
        cancelAutoClose();

        saveParsedOrdersToSheet(result.details).catch(() => {});
        await reply(formatResult(result.itemCount, result.userTotal));
        continue;
      }

      if (isOpen) {
        allText += '\n' + text;
        continue;
      }

      continue;
    }
  } catch (error) {
    console.error('Webhook error:', error);
  }
});

app.use((err, _req, res, _next) => {
  console.error('Global error:', err);
  res.sendStatus(200);
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log('Server running on ' + PORT);
});