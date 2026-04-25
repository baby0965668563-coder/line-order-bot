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

    orders.forEach(order => {
      total += order.price;
      itemCount[order.item] = (itemCount[order.item] || 0) + 1;
    });

    let msg = '📊 訂餐統計\n\n';

    Object.entries(itemCount).forEach(([item, count]) => {
      msg += `${item} x${count}\n`;
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

  const priceMatch = text.match(/(\d+)(?!.*\d)/);
  if (!priceMatch) {
    return res.sendStatus(200);
  }

  const price = Number(priceMatch[1]);
  const item = text.slice(0, priceMatch.index).trim();

  orders.push({
    item,
    price
  });

  // 加單時不回訊息，避免吵
  return res.sendStatus(200);
});

app.get('/', (req, res) => {
  res.send('LINE 訂餐統計機器人運作中');
});

const port = process.env.PORT || 3000;

app.listen(port, () => {
  console.log(`Server running on ${port}`);
});
