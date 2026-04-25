require('dotenv').config();

const express = require('express');
const line = require('@line/bot-sdk');

const app = express();

const config = {
  channelAccessToken: process.env.CHANNEL_ACCESS_TOKEN,
  channelSecret: process.env.CHANNEL_SECRET
};

const client = new line.Client(config);

// 狀態
let isOpen = false;
let allText = '';

// 清理文字（去標點 空白）
function clean(text) {
  return String(text || '')
    .replace(/[。.,，、!！?？:：;；"'（）()【】\[\]{}<>《》\s]/g, '')
    .trim();
}

// 解析訂單
function parseOrders(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);

  let currentItem = '';
  let currentPrice = 0;
  let pendingOrder = null;
  let itemBuffer = '';
  let priceMap = {};

  const itemCount = {};
  const userTotal = {};

  function cleanItemName(text) {
    return clean(text).replace(/\d+顆/g, '');
  }

  function add(item, price, name, qty = 1, note = '') {
    item = cleanItemName(item);
    name = clean(name);

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
    if (priceTable) {
      priceMap[priceTable[1]] = Number(priceTable[2]);
      continue;
    }

    if (pendingOrder && !/[+*]/.test(line) && !/\d/.test(line)) {
      add(pendingOrder.item, pendingOrder.price, line, 1);
      pendingOrder = null;
      continue;
    }

    const orderMatch = line.match(/^(.+?)[+*]\s*(半|0\.5|\d+)(.*)$/);
    if (orderMatch && currentItem) {
      const name = orderMatch[1];
      const rawQty = orderMatch[2];
      const extra = orderMatch[3] || '';
      const note = extra.includes('辣') ? '辣' : '';
      const price = getPrice(rawQty);
      const qty = rawQty === '半' || rawQty === '0.5' ? 1 : Number(rawQty);

      add(currentItem, price, name, qty, note);
      continue;
    }

    const inlineSymbol = line.match(/^(.+?)\s*[💰$＄]\s*(\d{1,5})\s*([^\d\s]+)$/);
    if (inlineSymbol) {
      add(inlineSymbol[1], Number(inlineSymbol[2]), inlineSymbol[3], 1);
      itemBuffer = '';
      continue;
    }

    const itemSymbol = line.match(/^(.+?)\s*[💰$＄]\s*(\d{1,5})$/);
    if (itemSymbol) {
      currentItem = itemSymbol[1];
      currentPrice = Number(itemSymbol[2]);
      itemBuffer = '';
      continue;
    }

    const inlineNoSymbol = line.match(/^(.+?)(\d{2,5})([^\d\s]+)$/);
    if (inlineNoSymbol && !/[+*]/.test(line)) {
      add(inlineNoSymbol[1], Number(inlineNoSymbol[2]), inlineNoSymbol[3], 1);
      itemBuffer = '';
      continue;
    }

    const noSymbolNoName = line.match(/^(.+?)(\d{2,5})$/);
    if (noSymbolNoName && !/[+*]/.test(line) && !/顆/.test(line)) {
      pendingOrder = {
        item: noSymbolNoName[1],
        price: Number(noSymbolNoName[2])
      };
      itemBuffer = '';
      continue;
    }

    const priceNameOnly = line.match(/^(\d{2,5})\s*([^\d\s]+)$/);
    if (priceNameOnly && itemBuffer) {
      add(itemBuffer, Number(priceNameOnly[1]), priceNameOnly[2], 1);
      itemBuffer = '';
      continue;
    }

    const priceOnly = line.match(/^(\d{2,5})$/);
    if (priceOnly && itemBuffer) {
      pendingOrder = {
        item: itemBuffer,
        price: Number(priceOnly[1])
      };
      itemBuffer = '';
      continue;
    }

    // 純文字：可能是品項，也可能是人名 +1
if (!/[+*]/.test(line)) {
    if (!/[+*]/.test(line)) {
    // 含「顆」通常是品項，例如：海苔6顆
    if (/顆/.test(line)) {
      currentItem = line;
      currentPrice = 0;
      itemBuffer = line;
      continue;
    }

    // 如果前面已經有品項，這行當作人名 +1
    if (currentItem) {
      add(currentItem, currentPrice, line, 1);
      continue;
    }

    // 否則才暫存成品項
    currentItem = line;
    currentPrice = 0;
    itemBuffer = line;
    continue;
  }

}

  return { itemCount, userTotal };
}
function formatResult(itemCount, userTotal) {
  let text = '📊 訂餐統計\n\n';

  text += '【品項數量】\n';
  for (let item in itemCount) {
    if (itemCount[item] > 0) {
      text += `${item} x${itemCount[item]}\n`;
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

// ⭐ 重點：不要用 express.json()
app.post('/webhook', line.middleware(config), async (req, res) => {
  try {
    const event = req.body.events[0];

    if (!event || event.type !== 'message' || event.message.type !== 'text') {
      return res.sendStatus(200);
    }

    const text = event.message.text.trim();

    // 開單（防呆）
    if (text === '開單') {
      if (isOpen) {
        await client.replyMessage(event.replyToken, {
          type: 'text',
          text: '目前已開單中，不會清空訂單'
        });
        return res.sendStatus(200);
      }

      isOpen = true;
      allText = '';

      await client.replyMessage(event.replyToken, {
        type: 'text',
        text: '已開單，可以開始點餐'
      });

      return res.sendStatus(200);
    }

    // 清空
    if (text === '清空') {
      allText = '';
      isOpen = false;

      await client.replyMessage(event.replyToken, {
        type: 'text',
        text: '已清空訂單'
      });

      return res.sendStatus(200);
    }

    // 結單 / 收單 / 統計
    if (text === '結單' || text === '收單' || text === '統計') {
      const result = parseOrders(allText);
      const reply = formatResult(result.itemCount, result.userTotal);

      isOpen = false;

      await client.replyMessage(event.replyToken, {
        type: 'text',
        text: reply
      });

      return res.sendStatus(200);
    }

    // 收集訂單
    if (isOpen) {
      allText += '\n' + text;
    }

    return res.sendStatus(200);

  } catch (error) {
    console.error('Webhook error:', error);
    return res.sendStatus(200);
  }
});

// 測試用首頁
app.get('/', (req, res) => {
  res.send('LINE 訂餐統計機器人運作中');
});

// 錯誤捕捉
app.use((err, req, res, next) => {
  console.error('Global error:', err);
  res.sendStatus(200);
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log('Server running on ' + PORT);
});
