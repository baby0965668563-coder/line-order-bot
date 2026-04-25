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
let allText = '';

function cleanText(text) {
  return String(text || '')
    .replace(/[。.,，、!！?？:：;；"'（）()【】\[\]{}<>《》\s]/g, '')
    .trim();
}

function parseOrders(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);

  let currentItem = '';
  let currentPrice = 0;
  let pendingOrder = null;

  const itemCount = {};
  const userTotal = {};

  function addOrder(item, price, name, qty = 1) {
    item = cleanText(item);
    name = cleanText(name);

    if (!item || !price || !name) return;

    itemCount[item] = (itemCount[item] || 0) + qty;
    userTotal[name] = (userTotal[name] || 0) + price * qty;
  }

  for (let line of lines) {
    if (/今天有人要訂|收錢|謝謝|下午|早上|晚上/.test(line)) continue;

    if (pendingOrder && !/[+*]/.test(line)) {
      addOrder(pendingOrder.item, pendingOrder.price, line, 1);
      pendingOrder = null;
      continue;
    }

    const itemPriceSymbol = line.match(/^(.+?)\s*[💰$＄]\s*(\d{1,4})$/);
    if (itemPriceSymbol && !/[+*]/.test(line)) {
      currentItem = itemPriceSymbol[1].trim();
      currentPrice = Number(itemPriceSymbol[2]);
      itemCount[cleanText(currentItem)] = itemCount[cleanText(currentItem)] || 0;
      pendingOrder = { item: currentItem, price: currentPrice };
      continue;
    }

    const itemPriceName = line.match(/^(.+?)(\d{1,4})([^\d\s]+)$/);
    if (itemPriceName && !/[+*]/.test(line)) {
      addOrder(
        itemPriceName[1].trim(),
        Number(itemPriceName[2]),
        itemPriceName[3],
        1
      );
      pendingOrder = null;
      continue;
    }

    const itemPriceNoName = line.match(/^(.+?)(\d{1,4})$/);
    if (itemPriceNoName && !/[+*]/.test(line)) {
      currentItem = itemPriceNoName[1].trim();
      currentPrice = Number(itemPriceNoName[2]);
      itemCount[cleanText(currentItem)] = itemCount[cleanText(currentItem)] || 0;
      pendingOrder = { item: currentItem, price: currentPrice };
      continue;
    }

    const qtyMatch = line.match(/^(.+?)[+*]\s*(\d+)$/);
    if (qtyMatch && currentItem && currentPrice) {
      const name = qtyMatch[1];
      const qty = Number(qtyMatch[2]);
      addOrder(currentItem, currentPrice, name, qty);
      pendingOrder = null;
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

app.post('/webhook', line.middleware(config), async (req, res) => {
  try {
    const event = req.body.events[0];

    if (!event || event.type !== 'message' || event.message.type !== 'text') {
      return res.sendStatus(200);
    }

    const text = event.message.text.trim();

    if (text === '開單') {
      isOpen = true;
      allText = '';

      await client.replyMessage(event.replyToken, {
        type: 'text',
        text: '已開單，可以開始點餐'
      });

      return res.sendStatus(200);
    }

    if (text === '清空') {
      allText = '';

      await client.replyMessage(event.replyToken, {
        type: 'text',
        text: '已清空訂單'
      });

      return res.sendStatus(200);
    }

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

    if (isOpen) {
      allText += '\n' + text;
    }

    return res.sendStatus(200);
  } catch (error) {
    console.error(error);
    return res.sendStatus(200);
  }
});

app.get('/', (req, res) => {
  res.send('LINE 訂餐統計機器人運作中');
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log('Server running on ' + PORT);
});
