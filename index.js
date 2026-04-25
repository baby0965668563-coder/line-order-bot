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

function cleanItem(text) {
  return String(text || '')
    .replace(/\d+\s*顆/g, '')
    .replace(/[。.,，]/g, '')
    .trim();
}

function qtyToCount(rawQty) {
  if (rawQty === '半' || rawQty === '0.5') return 1;
  if (rawQty === '1') return 2;
  return Number(rawQty);
}

function parseOrders(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  const result = [];

  let currentItem = '';
  let currentPrice = 0;
  let priceMap = {};

  function getPrice(rawQty) {
    if (rawQty === '半' && priceMap['0.5']) return priceMap['0.5'];
    if (rawQty === '0.5' && priceMap['半']) return priceMap['半'];
    if (priceMap[rawQty]) return priceMap[rawQty];
    return currentPrice;
  }

  function pushOrder(item, price, user, qty = 1) {
    for (let i = 0; i < qty; i++) {
      result.push({
        item: cleanItem(item),
        price,
        user: cleanName(user)
      });
    }
  }

  for (const rawLine of lines) {
    const line = rawLine.trim();

    if (/今天有人要訂|收錢|謝謝|下午|早上|晚上/.test(line)) {
      continue;
    }

    // 價格表：0.5 $55 / 1 $110 / 半$55
    const priceTable = [...line.matchAll(/(半|0\.5|1)\s*[💰$＄]\s*(\d{1,4})/g)];
    if (priceTable.length > 0) {
      priceMap = {};
      priceTable.forEach(m => {
        priceMap[m[1]] = Number(m[2]);
      });
      continue;
    }

    // 品項 + 金額：原味 $60 / 辣味 💰60
    const itemPriceMatch = line.match(/^(.+?)\s*[💰$＄]\s*(\d{1,4})$/);

    if (itemPriceMatch && !/[+*]\s*(半|0\.5|\d+)/.test(line)) {
      currentItem = cleanItem(itemPriceMatch[1]);
      currentPrice = Number(itemPriceMatch[2]);
      priceMap = {};
      continue;
    }

    // 名字 + 數量：慧玲+2 / 香菇+2 / 瑞琴*半 / 慧明+0.5
    const qtyMatch = line.match(/^(.+?)[+*]\s*(半|0\.5|\d+).*$/);

    if (qtyMatch && currentItem) {
      const user = qtyMatch[1];
      const rawQty = qtyMatch[2];
      const qty = qtyToCount(rawQty);
      const price = getPrice(rawQty);

      if (price) {
        pushOrder(currentItem, price, user, qty);
      }
      continue;
    }

    // 飲料格式：珍珠奶茶微糖微冰 40士豪
    const drinkMatch = line.match(/^(.+?)\s+(\d{1,4})\s*([^\d\s]+)$/);

    if (drinkMatch) {
      pushOrder(
        cleanItem(drinkMatch[1]),
        Number(drinkMatch[2]),
        drinkMatch[3],
        1
      );
      continue;
    }

    // 純品項行：椒鹽 / 馬告 / 原味不加
    if (!/[+*]/.test(line)) {
      const item = cleanItem(line);
      if (!item) continue;

      currentItem = item;
      currentPrice = 0;
      continue;
    }
  }

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

  if (text === '結單' || text === '統計' || text === '收單') {
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
