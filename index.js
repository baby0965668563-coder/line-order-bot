require('dotenv').config();

const express = require('express');
const line = require('@line/bot-sdk');

const app = express();

const config = {
  channelAccessToken: process.env.CHANNEL_ACCESS_TOKEN,
  channelSecret: process.env.CHANNEL_SECRET
};

const client = new line.Client(config);

// 狀態
let isOpen = false;
let allText = '';

// 清理文字（去標點 空白）
function clean(text) {
  return String(text || '')
    .replace(/[。.,，、!！?？:：;；"'（）()【】\[\]{}<>《》\s]/g, '')
    .trim();
}

// 解析訂單
function parseOrders(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);

  let currentItem = '';
  let currentPrice = 0;

  const itemCount = {};
  const userTotal = {};

  function add(item, price, name, qty, note = '') {
    item = clean(item);
    name = clean(name);

    if (!item || !price || !name || !qty) return;

    const finalItem = note ? `${item}（${note}）` : item;

    itemCount[finalItem] = (itemCount[finalItem] || 0) + qty;
    userTotal[name] = (userTotal[name] || 0) + price * qty;
  }

  for (let line of lines) {

    // 過濾垃圾文字
    if (/今天有人要訂|收錢|謝謝|下午|早上|晚上/.test(line)) continue;

    // ⭐ 抓品項 + 金額（不管前後空白）
    const itemMatch = line.match(/(.+?)\s*[💰$＄]\s*(\d+)/);
    
    // ✅ 新增：同一行「品項 + 金額 + 人名」
    const inlineMatch = line.match(/(.+?)[💰$]\s*(\d+)\s*([^\d\s]+)/);

    if (inlineMatch) {
  const item = inlineMatch[1].trim();
  const price = parseInt(inlineMatch[2]);
  const user = inlineMatch[3].trim();

  // 初始化
  if (!itemCount[item]) itemCount[item] = 0;
  if (!userTotal[user]) userTotal[user] = 0;

  itemCount[item] += 1;
  userTotal[user] += price;

  continue; // ⚠️ 很重要！避免被下面邏輯再跑一次
}
    if (itemMatch) {
      currentItem = itemMatch[1];
      currentPrice = Number(itemMatch[2]);
      continue;
    }

    // ⭐ 抓人名 + 數量 + 辣
    const orderMatch = line.match(/(.+?)[+*]\s*(\d+)(.*)/);

    if (orderMatch && currentItem && currentPrice) {
      const name = orderMatch[1];
      const qty = Number(orderMatch[2]);
      const extra = orderMatch[3] || '';

      let note = '';
      if (extra.includes('辣')) note = '辣';

      add(currentItem, currentPrice, name, qty, note);
    }
  }

  return { itemCount, userTotal };
}

// 輸出結果
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

// ⭐ 重點：不要用 express.json()
app.post('/webhook', line.middleware(config), async (req, res) => {
  try {
    const event = req.body.events[0];

    if (!event || event.type !== 'message' || event.message.type !== 'text') {
      return res.sendStatus(200);
    }

    const text = event.message.text.trim();

    // 開單（防呆）
    if (text === '開單') {
      if (isOpen) {
        await client.replyMessage(event.replyToken, {
          type: 'text',
          text: '目前已開單中，不會清空訂單'
        });
        return res.sendStatus(200);
      }

      isOpen = true;
      allText = '';

      await client.replyMessage(event.replyToken, {
        type: 'text',
        text: '已開單，可以開始點餐'
      });

      return res.sendStatus(200);
    }

    // 清空
    if (text === '清空') {
      allText = '';
      isOpen = false;

      await client.replyMessage(event.replyToken, {
        type: 'text',
        text: '已清空訂單'
      });

      return res.sendStatus(200);
    }

    // 結單 / 收單 / 統計
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

    // 收集訂單
    if (isOpen) {
      allText += '\n' + text;
    }

    return res.sendStatus(200);

  } catch (error) {
    console.error('Webhook error:', error);
    return res.sendStatus(200);
  }
});

// 測試用首頁
app.get('/', (req, res) => {
  res.send('LINE 訂餐統計機器人運作中');
});

// 錯誤捕捉
app.use((err, req, res, next) => {
  console.error('Global error:', err);
  res.sendStatus(200);
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log('Server running on ' + PORT);
});
