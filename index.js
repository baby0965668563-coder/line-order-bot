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
  let priceMap = {};
  let pendingDrink = null;

  const itemCount = {};
  const userTotal = {};

  function cleanName(name) {
    return String(name || '').replace(/[。.,，\s]/g, '').trim();
  }

  function addOrder(item, price, name, count = 1) {
    item = String(item || '').trim();
    name = cleanName(name);
    if (!item || !name) return;

    itemCount[item] = (itemCount[item] || 0) + count;
    userTotal[name] = (userTotal[name] || 0) + price * count;
  }

  function getPrice(rawQty) {
    if (priceMap[rawQty] !== undefined) return priceMap[rawQty];
    if (rawQty === '半' && priceMap['0.5'] !== undefined) return priceMap['0.5'];
    if (rawQty === '0.5' && priceMap['半'] !== undefined) return priceMap['半'];
    return currentPrice;
  }

  for (let line of lines) {
    if (/今天有人要訂|收錢|謝謝|下午|早上|晚上/.test(line)) continue;

    // 飲料上一行沒名字，這行補名字
    if (pendingDrink && !/\d/.test(line) && !/[+*]/.test(line)) {
      addOrder(pendingDrink.item, pendingDrink.price, line, 1);
      pendingDrink = null;
      continue;
    }

    // 價格表：0.5 $55 / 1 $110 / 半 $55
    const priceTable = [...line.matchAll(/(半|\d+(?:\.\d+)?)\s*[💰$＄]\s*(\d+)/g)];
    if (priceTable.length > 0 && /^[半\d]/.test(line)) {
      priceMap = {};
      priceTable.forEach(m => {
        priceMap[m[1]] = parseInt(m[2], 10);
      });
      continue;
    }

    // 品項 + $價格：蔥肉餅 $45 / 紅豆餅💰25
    const itemMatch = line.match(/^(.+?)\s*[💰$＄]\s*(\d+)/);
    if (itemMatch && !/[+*]/.test(line)) {
      pendingDrink = null;
      currentItem = itemMatch[1].trim();
      currentPrice = parseInt(itemMatch[2], 10);
      priceMap = {};
      itemCount[currentItem] = itemCount[currentItem] || 0;
      continue;
    }

    // 名字 + 數量：慧明*0.5 / 姿瑜*1 / 心玄+2 / 瑞琴*半
    const orderMatch = line.match(/^(.+?)[+*]\s*(半|\d+(?:\.\d+)?)/);
    if (orderMatch && currentItem) {
      const name = orderMatch[1].trim();
      const rawQty = orderMatch[2];
      const price = getPrice(rawQty);

      // 0.5 / 半 算 1 筆；1 算 1 筆；2 算 2 筆
      const count =
        rawQty === '0.5' || rawQty === '半'
          ? 1
          : parseInt(rawQty, 10);

      addOrder(currentItem, price, name, count);
      continue;
    }

    // 飲料：品項40名字 / 品項 40 名字 / 品項55 下一行名字
    const drinkMatch = line.match(/^(.+?)\s*(\d{2,4})\s*([^\d\s]+)?$/);
    if (drinkMatch && !/[+*]/.test(line)) {
      const item = drinkMatch[1].trim();
      const price = parseInt(drinkMatch[2], 10);
      const name = drinkMatch[3] || '';

      if (name) {
        addOrder(item, price, name, 1);
      } else {
        pendingDrink = { item, price };
      }
      continue;
    }

    // 純品項：椒鹽 / 馬告 / 辣味椒鹽
    if (!/[+*]/.test(line)) {
      currentItem = line.trim();
      currentPrice = 0;
      itemCount[currentItem] = itemCount[currentItem] || 0;
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
