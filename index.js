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

app.use(express.json());

function parseOrders(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);

  let currentItem = null;
  let currentPrice = 0;
  let pendingDrink = null;

  const itemCount = {};
  const userTotal = {};

  function addOrder(item, price, name, qty = 1) {
    item = String(item || '').trim();
    name = String(name || '').replace(/[。.,，\s]/g, '').trim();
    if (!item || !price || !name) return;

    itemCount[item] = (itemCount[item] || 0) + qty;
    userTotal[name] = (userTotal[name] || 0) + price * qty;
  }

  for (let line of lines) {

    if (/今天有人要訂|收錢|謝謝|下午|早上|晚上/.test(line)) continue;

    // 👉 接續上一行飲料
    if (pendingDrink && !/\d/.test(line) && !/[+*]/.test(line)) {
      addOrder(pendingDrink.item, pendingDrink.price, line, 1);
      pendingDrink = null;
      continue;
    }

    // 👉 抓「品項 + $」
    const itemMatch = line.match(/^(.+?)\s*[💰$＄]\s*(\d+)/);
    if (itemMatch) {
      pendingDrink = null;
      currentItem = itemMatch[1].trim();
      currentPrice = parseInt(itemMatch[2]);
      if (!itemCount[currentItem]) itemCount[currentItem] = 0;
      continue;
    }

    // 👉 抓「人名+數量」
    const orderMatch = line.match(/^(.+?)[+*](\d+)/);
    if (orderMatch && currentItem) {
      const name = orderMatch[1].trim();
      const qty = parseInt(orderMatch[2]);
      addOrder(currentItem, currentPrice, name, qty);
      continue;
    }

    // 👉 抓飲料
    const drinkMatch = line.match(/^(.+?)\s*(\d{2,4})\s*([^\d\s]+)?$/);
    if (drinkMatch && !/[+*]/.test(line)) {
      const item = drinkMatch[1].trim();
      const price = parseInt(drinkMatch[2]);
      const name = drinkMatch[3] || '';

      if (name) {
        addOrder(item, price, name, 1);
      } else {
        pendingDrink = { item, price };
      }
      continue;
    }
  }

  return { itemCount, userTotal };
}

function formatResult(itemCount, userTotal) {
  let text = '📊 訂餐統計\n\n';

  text += '【品項數量】\n';
  for (let item in itemCount) {
    text += `${item} x${itemCount[item]}\n`;
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
  const event = req.body.events[0];

  if (!event || event.type !== 'message') return res.sendStatus(200);

  const text = event.message.text.trim();

  if (text === '開單') {
    isOpen = true;
    allText = '';
    return client.replyMessage(event.replyToken, { type: 'text', text: '已開單' });
  }

  if (text === '結單' || text === '收單' || text === '統計') {
    const result = parseOrders(allText);
    const reply = formatResult(result.itemCount, result.userTotal);
    isOpen = false;

    return client.replyMessage(event.replyToken, { type: 'text', text: reply });
  }

  if (isOpen) {
    allText += '\n' + text;
  }

  return res.sendStatus(200);
});

app.get('/', (req, res) => res.send('ok'));

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log('running on ' + PORT);
});
