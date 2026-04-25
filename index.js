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
  return text.replace(/[。.,，\s]/g, '').trim();
}

function parseOrders(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  const result = [];

  let currentItem = '';
  let currentPrice = 0;
  let lastOrder = null;

  for (const line of lines) {
    // 食物格式：品項 $45
    const itemPriceMatch = line.match(/^(.+?)\s*\$?\s*(\d+)$/);

    if (itemPriceMatch && !/[+*]/.test(line)) {
      currentItem = itemPriceMatch[1].trim();
      currentPrice = Number(itemPriceMatch[2]);
      lastOrder = null;
      continue;
    }

    // 食物格式：姓名 +2 辣 / 姓名*2
    const qtyMatch = line.match(/^(.+?)[+*]\s*(\d+)(.*)$/);
    if (qtyMatch && currentItem && currentPrice) {
      const user = cleanName(qtyMatch[1]);
      const qty = Number(qtyMatch[2]);
      const note = qtyMatch[3].trim();

      for (let i = 0; i < qty; i++) {
        result.push({
          item: note ? `${currentItem}（${note}）` : currentItem,
          price: currentPrice,
          user
        });
      }
      lastOrder = null;
      continue;
    }

    // 飲料格式：品項 60 姓名
    const priceMatch = line.match(/(\d+)(?!.*\d)/);

    if (!priceMatch) {
      if (lastOrder && !lastOrder.user && line.length <= 6) {
        lastOrder.user = cleanName(line);
      }
      continue;
    }

    const price = Number(priceMatch[1]);
    const item = line.slice(0, priceMatch.index).trim();
    const user = cleanName(line.slice(priceMatch.index + priceMatch[1].length));

    const order = { item, price, user };
    result.push(order);
    lastOrder = order;
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

  // 加單時不回覆，避免群組一直跳通知
  return res.sendStatus(200);
});

app.get('/', (req, res) => {
  res.send('LINE 訂餐統計機器人運作中');
});

const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`Server running on ${port}`);
});
