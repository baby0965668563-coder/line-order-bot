require('dotenv').config();

const express = require('express');
const line = require('@line/bot-sdk');

const app = express();

const config = {
  channelAccessToken: process.env.CHANNEL_ACCESS_TOKEN,
  channelSecret: process.env.CHANNEL_SECRET
};

const client = new line.Client(config);

let isOpen = false;
let orders = [];

function cleanName(text) {
  return String(text || '').replace(/[。.,，\s]/g, '').trim();
}

function parseOrders(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  const result = [];

  let currentItem = '';
  let currentPrice = 0;
  let pendingItem = '';
  let pendingQtyLines = [];
  let pendingDrink = null;

  function pushOrder(item, price, user = '', qty = 1, note = '') {
    for (let i = 0; i < qty; i++) {
      result.push({
        item: note ? `${item}（${note}）` : item,
        price,
        user: cleanName(user)
      });
    }
  }

  function flushPendingQty(price) {
    if (!pendingItem || pendingQtyLines.length === 0) return;

    for (const qLine of pendingQtyLines) {
      const qtyMatch = qLine.match(/^(.+?)[+*]\s*(半|\d+)(.*)$/);
      if (!qtyMatch) continue;

      const user = qtyMatch[1];
      const rawQty = qtyMatch[2];
      const qty = rawQty === '半' ? 1 : Number(rawQty);
      const note = rawQty === '半'
        ? `半${qtyMatch[3].trim()}`
        : qtyMatch[3].trim();

      pushOrder(pendingItem, price, user, qty, note);
    }

    pendingQtyLines = [];
    pendingItem = '';
  }

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    if (pendingDrink && !pendingDrink.user && !/\d/.test(line)) {
      pendingDrink.user = cleanName(line);
      result.push(pendingDrink);
      pendingDrink = null;
      continue;
    }

    if (pendingDrink) {
      result.push(pendingDrink);
      pendingDrink = null;
    }

    const priceOnlyMatch = line.match(/^(一個|每個|單價)?\s*\$?\s*(\d{2,3})$/);
    if (priceOnlyMatch && pendingQtyLines.length > 0) {
      flushPendingQty(Number(priceOnlyMatch[2]));
      currentPrice = Number(priceOnlyMatch[2]);
      continue;
    }

    const headerMatch = line.match(/^(.+?)\s*\$?\s*(\d{2,3})$/);
    const isQtyLine = /^.+?[+*]\s*(半|\d+)/.test(line);

    if (headerMatch && !isQtyLine) {
      currentItem = headerMatch[1].trim();
      currentPrice = Number(headerMatch[2]);
      pendingItem = currentItem;
      pendingQtyLines = [];
      continue;
    }

    const qtyMatch = line.match(/^(.+?)[+*]\s*(半|\d+)(.*)$/);

    if (qtyMatch) {
      const user = qtyMatch[1];
      const rawQty = qtyMatch[2];
      const qty = rawQty === '半' ? 1 : Number(rawQty);
      const note = rawQty === '半'
        ? `半${qtyMatch[3].trim()}`
        : qtyMatch[3].trim();

      if (currentItem && currentPrice) {
        pushOrder(currentItem, currentPrice, user, qty, note);
      } else if (pendingItem) {
        pendingQtyLines.push(line);
      }
      continue;
    }

    const drinkMatch = line.match(/^(.+?)\s*(\d{2,3})\s*([^\d\s]+)?$/);

    if (drinkMatch) {
      const item = drinkMatch[1].trim();
      const price = Number(drinkMatch[2]);
      const user = cleanName(drinkMatch[3] || '');

      pendingDrink = { item, price, user };
      continue;
    }

    const nextLine = lines[i + 1] || '';
    const nextPriceMatch = nextLine.match(/^(\d{2,3})\s*([^\d\s]+)?$/);

    if (nextPriceMatch) {
      const item = line.trim();
      const price = Number(nextPriceMatch[1]);
      const user = cleanName(nextPriceMatch[2] || '');

      pendingDrink = { item, price, user };
      i++;
      continue;
    }

    if (!/\d/.test(line)) {
      currentItem = '';
      currentPrice = 0;
      pendingItem = line;
      pendingQtyLines = [];
      continue;
    }
  }

  if (pendingDrink) result.push(pendingDrink);

  return result;
}

app.post('/webhook', line.middleware(config), async (req, res) => {
  const event = req.body.events[0];

  if (!event || event.type !== 'message' || event.message.type !== 'text') {
    return res.sendStatus(200);
  }

  const text = event.message.text.trim();

  if (text === '開單') {
    isOpen = true;
    orders = [];
    await client.replyMessage(event.replyToken, {
      type: 'text',
      text: '已開單，可以開始點餐'
    });
    return res.sendStatus(200);
  }

  if (text === '清空') {
    orders = [];
    await client.replyMessage(event.replyToken, {
      type: 'text',
      text: '已清空訂單'
    });
    return res.sendStatus(200);
  }

  if (text === '結單' || text === '統計') {
    isOpen = false;

    if (orders.length === 0) {
      await client.replyMessage(event.replyToken, {
        type: 'text',
        text: '目前沒有訂單'
      });
      return res.sendStatus(200);
    }

    let total = 0;
    const itemCount = {};
    const userTotal = {};

    orders.forEach(order => {
      total += order.price;
      itemCount[order.item] = (itemCount[order.item] || 0) + 1;

      const user = order.user || '未填';
      userTotal[user] = (userTotal[user] || 0) + order.price;
    });

    let msg = '📊 訂餐統計\n\n【品項數量】\n';

    Object.entries(itemCount).forEach(([item, count]) => {
      msg += `${item} x${count}\n`;
    });

    msg += '\n【個人金額】\n';

    Object.entries(userTotal).forEach(([user, amount]) => {
      msg += `${user}：$${amount}\n`;
    });

    msg += `\n💰 總金額：$${total}`;

    await client.replyMessage(event.replyToken, {
      type: 'text',
      text: msg
    });

    return res.sendStatus(200);
  }

  if (!isOpen) {
    return res.sendStatus(200);
  }

  const parsed = parseOrders(text);
  orders.push(...parsed);

  return res.sendStatus(200);
});

app.get('/', (req, res) => {
  res.send('LINE 訂餐統計機器人運作中');
});

const port = process.env.PORT || 3000;

app.listen(port, () => {
  console.log(`Server running on ${port}`);
});
