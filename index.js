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

function clean(text) {
  return String(text || '')
    .replace(/[^\u4e00-\u9fa5a-zA-Z0-9]/g, '')
    .trim();
}

function parseOrders(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);

  let currentItem = '';
  let currentPrice = 0;

  const itemCount = {};
  const userTotal = {};

  function add(item, price, name, qty = 1) {
    item = clean(item);
    name = clean(name);

    if (!item || !price || !name) return;

    itemCount[item] = (itemCount[item] || 0) + qty;
    userTotal[name] = (userTotal[name] || 0) + price * qty;
  }

  for (let i = 0; i < lines.length; i++) {
    let line = lines[i];

    if (/收錢|謝謝|下午|早上|晚上/.test(line)) continue;

    // 👉 抓價格（有$ / 💰 / 或純數字）
    let priceMatch = line.match(/(\d{2,4})/);

    // 👉 抓數量（+1、*2、1份都吃）
    let qtyMatch = line.match(/(\d+)/);

    // 👉 抓名字（中文）
    let nameMatch = line.match(/[\u4e00-\u9fa5]{2,4}/g);

    // 👉 情境1：品項+價格
    if (/[💰$]/.test(line)) {
      currentPrice = parseInt(priceMatch?.[1] || 0);
      currentItem = clean(line.replace(/[💰$]\s*\d+/, ''));
      continue;
    }

    // 👉 情境2：品項+價格+名字（同一行）
    if (priceMatch && nameMatch && !/[+*]/.test(line)) {
      let item = clean(line.replace(nameMatch[nameMatch.length - 1], '').replace(/\d+/g, ''));
      let name = nameMatch[nameMatch.length - 1];
      let price = parseInt(priceMatch[1]);

      add(item, price, name, 1);
      continue;
    }

    // 👉 情境3：名字+數量
    if (/[+*]/.test(line) && currentItem && currentPrice) {
      let name = nameMatch?.[0];
      let qty = parseInt(qtyMatch?.[1] || 1);

      add(currentItem, currentPrice, name, qty);
      continue;
    }

    // 👉 情境4：價格在下一行
    if (!/[+*]/.test(line) && !priceMatch) {
      let next = lines[i + 1] || '';

      if (/^\d{2,4}$/.test(next)) {
        currentItem = clean(line);
        currentPrice = parseInt(next);
        i++;
        continue;
      }
    }

    // 👉 情境5：名字單獨一行（接上一筆）
    if (!/[+*]/.test(line) && currentItem && currentPrice) {
      let name = nameMatch?.[0];
      if (name) add(currentItem, currentPrice, name, 1);
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

app.use(express.json());

app.post('/webhook', line.middleware(config), async (req, res) => {
  const event = req.body.events[0];

  if (!event || event.type !== 'message') return res.sendStatus(200);

  const text = event.message.text.trim();

  if (text === '開單') {
    isOpen = true;
    allText = '';
    await client.replyMessage(event.replyToken, { type: 'text', text: '已開單' });
    return res.sendStatus(200);
  }

  if (text === '結單' || text === '收單' || text === '統計') {
    const result = parseOrders(allText);
    const reply = formatResult(result.itemCount, result.userTotal);
    isOpen = false;

    await client.replyMessage(event.replyToken, { type: 'text', text: reply });
    return res.sendStatus(200);
  }

  if (isOpen) {
    allText += '\n' + text;
  }

  return res.sendStatus(200);
});

app.get('/', (req, res) => res.send('ok'));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log('running ' + PORT));
