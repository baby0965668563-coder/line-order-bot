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

  function pushOrder(item, price, user, qty = 1) {
    for (let i = 0; i < qty; i++) {
      result.push({
        item,
        price,
        user: cleanName(user)
      });
    }
  }

  for (const line of lines) {
    if (/收錢|謝謝|下午|早上|晚上/.test(line)) continue;

    // ① 品項 + 金額：原味 $55 / 椒鹽 60 / 蔥肉餅+蛋 $55
    const priceMatch = line.match(/^(.+?)\s*[💰$＄]?\s*(\d{1,4})$/);

    if (priceMatch && !/[+*]\s*(半|\d+)/.test(line)) {
      currentItem = priceMatch[1].trim();
      currentPrice = Number(priceMatch[2]);
      continue;
    }

    // ② 名字 + 數量：慧玲+2 / 姿瑜*1 / 瑞琴*半
    const qtyMatch = line.match(/^(.+?)[+*]\s*(半|\d+).*$/);

    if (qtyMatch && currentItem && currentPrice) {
      const user = qtyMatch[1];
      const qty = qtyMatch[2] === '半' ? 1 : Number(qtyMatch[2]);
      pushOrder(currentItem, currentPrice, user, qty);
      continue;
    }

    // ③ 飲料格式：珍珠奶茶微糖微冰 40士豪
    const drinkMatch = line.match(/^(.+?)\s*(\d{1,4})\s*([^\d\s]+)?$/);

    if (drinkMatch) {
      const item = drinkMatch[1].trim();
      const price = Number(drinkMatch[2]);
      const user = cleanName(drinkMatch[3] || '');

      if (user) {
        pushOrder(item, price, user, 1);
      } else {
        currentItem = item;
        currentPrice = price;
      }
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

  // 加單不回覆，避免群組一直跳通知
  return res.sendStatus(200);
});

app.get('/', (req, res) => {
  res.send('LINE 訂餐統計機器人運作中');
});

const port = process.env.PORT || 3000;

app.listen(port, () => {
  console.log(`Server running on ${port}`);
});
