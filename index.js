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

function parseOrders(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);

  let currentItem = null;
  let currentPrice = 0;
  let pendingDrink = null;
  let pendingItemLine = null;
  let priceMap = {};

  const itemCount = {};
  const userTotal = {};

  function cleanName(name) {
    return String(name || '').replace(/[。.,，\s]/g, '').trim();
  }

  function addOrder(item, price, name, qty = 1) {
    item = String(item || '').trim();
    name = cleanName(name);

    if (!item || !price || !name) return;

    itemCount[item] = (itemCount[item] || 0) + qty;
    userTotal[name] = (userTotal[name] || 0) + price * qty;
  }

  function getPrice(rawQty) {
    if (priceMap[rawQty] !== undefined) return priceMap[rawQty];
    if (rawQty === '半' && priceMap['0.5'] !== undefined) return priceMap['0.5'];
    if (rawQty === '0.5' && priceMap['半'] !== undefined) return priceMap['半'];
    return currentPrice;
  }

  function getQty(rawQty) {
    if (rawQty === '半' || rawQty === '0.5') return 1;
    return parseInt(rawQty, 10);
  }

  for (let i = 0; i < lines.length; i++) {
    let line = lines[i];

    if (/今天有人要訂|收錢|謝謝|下午|早上|晚上/.test(line)) continue;

    // 價格表：0.5 $55 / 1 $110 / 半 $55
    const priceTable = line.match(/^(半|0\.5|1)\s*[💰$＄]\s*(\d+)$/);
    if (priceTable) {
      priceMap[priceTable[1]] = parseInt(priceTable[2], 10);
      continue;
    }

    // 飲料：上一筆缺姓名，這行補姓名
    if (pendingDrink && !/\d/.test(line) && !/[+*]/.test(line)) {
      addOrder(pendingDrink.item, pendingDrink.price, line);
      pendingDrink = null;
      continue;
    }

    // 品項 $金額：午餐/下午茶
    const itemMatch = line.match(/^(.+?)\s*[💰$＄]\s*(\d+)/);
    if (itemMatch && !/[+*]/.test(line)) {
      currentItem = itemMatch[1].trim();
      currentPrice = parseInt(itemMatch[2], 10);
      itemCount[currentItem] = itemCount[currentItem] || 0;
      continue;
    }

    // 人名+數量：芷葳+2 / 心玄*3 / 慧明*0.5 / 瑞琴*半
    const orderMatch = line.match(/^(.+?)[+*]\s*(半|0\.5|\d+)/);
    if (orderMatch && currentItem) {
      const name = orderMatch[1];
      const rawQty = orderMatch[2];
      const qty = getQty(rawQty);
      const price = getPrice(rawQty);

      addOrder(currentItem, price, name, qty);
      continue;
    }

    // 飲料：品項40姓名 / 品項 40 姓名
    const drinkFull = line.match(/^(.+?)\s*(\d{2,4})\s*([^\d\s]+)$/);
    if (drinkFull && !/[+*]/.test(line)) {
      addOrder(drinkFull[1].trim(), parseInt(drinkFull[2], 10), drinkFull[3]);
      continue;
    }

    // 飲料：品項40 / 品項 40，姓名下一行
    const drinkNoName = line.match(/^(.+?)\s*(\d{2,4})$/);
    if (drinkNoName && !/[+*]/.test(line)) {
      pendingDrink = {
        item: drinkNoName[1].trim(),
        price: parseInt(drinkNoName[2], 10)
      };
      continue;
    }

    // 飲料：品項一行、價格一行、姓名一行
    if (!/\d/.test(line) && !/[+*]/.test(line)) {
      pendingItemLine = line;
      continue;
    }

    // 價格單獨一行：55
    if (/^\d{2,4}$/.test(line) && pendingItemLine) {
      pendingDrink = {
        item: pendingItemLine,
        price: parseInt(line, 10)
      };
      pendingItemLine = null;
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
