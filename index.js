require('dotenv').config();

const express = require('express');
const line = require('@line/bot-sdk');
const { GoogleSpreadsheet } = require('google-spreadsheet');

const app = express();

const config = {
  channelAccessToken: process.env.CHANNEL_ACCESS_TOKEN,
  channelSecret: process.env.CHANNEL_SECRET
};

const client = new line.Client(config);
const doc = new GoogleSpreadsheet(process.env.SHEET_ID);

async function getSheet() {
  await doc.useServiceAccountAuth({
    client_email: process.env.GOOGLE_CLIENT_EMAIL,
    private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
  });
  await doc.loadInfo();
  return doc.sheetsByIndex[0];
}

function parseOrders(text) {
  const lines = text
    .split('\n')
    .map(l => l.trim())
    .filter(Boolean);

  const orders = [];
  let lastOrder = null;

  for (const line of lines) {
    if (line.includes('有人要訂')) continue;

    const priceMatch = line.match(/(\d+)(?!.*\d)/);

    if (!priceMatch) {
      if (lastOrder && !lastOrder.user && line.length <= 6) {
        lastOrder.user = line.replace(/[。.,，\s]/g, '');
      }
      continue;
    }

    const price = Number(priceMatch[1]);
    const before = line.slice(0, priceMatch.index).trim();
    const after = line.slice(priceMatch.index + priceMatch[1].length).trim();

    const user = after ? after.replace(/[。.,，\s]/g, '') : '';

    const order = {
      item: before,
      price,
      user
    };

    orders.push(order);
    lastOrder = order;
  }

  return orders;
}

async function saveOrders(orders) {
  const sheet = await getSheet();

  for (const order of orders) {
    await sheet.addRow({
      時間: new Date().toLocaleString('zh-TW'),
      使用者: order.user || '未填',
      品項: order.item,
      金額: order.price
    });
  }
}

async function getSummary() {
  const sheet = await getSheet();
  const rows = await sheet.getRows();

  if (rows.length === 0) return '目前沒有訂單';

  let total = 0;
  const itemCount = {};
  const userTotal = {};

  rows.forEach(row => {
    const item = row['品項'];
    const user = row['使用者'] || '未填';
    const price = Number(row['金額']) || 0;

    total += price;
    itemCount[item] = (itemCount[item] || 0) + 1;
    userTotal[user] = (userTotal[user] || 0) + price;
  });

  let msg = '📊 今日訂餐統計\n\n【品項數量】\n';

  Object.entries(itemCount).forEach(([item, count]) => {
    msg += `${item} x${count}\n`;
  });

  msg += '\n【個人金額】\n';

  Object.entries(userTotal).forEach(([user, amount]) => {
    msg += `${user}：$${amount}\n`;
  });

  msg += `\n💰 總金額：$${total}`;

  return msg;
}

app.post('/webhook', line.middleware(config), async (req, res) => {
  try {
    await Promise.all(req.body.events.map(handleEvent));
    res.status(200).end();
  } catch (err) {
    console.error(err);
    res.status(500).end();
  }
});

async function handleEvent(event) {
  if (event.type !== 'message' || event.message.type !== 'text') return null;

  const text = event.message.text.trim();

  if (text === '統計') {
    const summary = await getSummary();
    return client.replyMessage(event.replyToken, {
      type: 'text',
      text: summary
    });
  }

  if (text === '清空') {
    const sheet = await getSheet();
    const rows = await sheet.getRows();
    await Promise.all(rows.map(row => row.delete()));

    return client.replyMessage(event.replyToken, {
      type: 'text',
      text: '已清空訂單'
    });
  }

  const orders = parseOrders(text);

  if (orders.length === 0) return null;

  await saveOrders(orders);

  return client.replyMessage(event.replyToken, {
    type: 'text',
    text: `已記錄 ${orders.length} 筆訂單`
  });
}

app.get('/', (req, res) => {
  res.send('LINE 訂餐統計機器人運作中');
});

const port = process.env.PORT || 3000;

app.listen(port, () => {
  console.log(`Server running on ${port}`);
});
