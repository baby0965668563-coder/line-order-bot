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
  let pendingOrder = null;
  let itemBuffer = '';

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
    if (/今天有人要訂|收錢|謝謝|下午|早上|晚上|星期/.test(line)) continue;

    // 上一行是 品項+價格，這一行是名字
    if (pendingOrder && !/[+*]/.test(line) && !/\d/.test(line)) {
      add(pendingOrder.item, pendingOrder.price, line, 1);
      pendingOrder = null;
      continue;
    }

    // 人名+數量：藝馨+1辣 / 秋蘭+2 / 翊婕*2
    const orderMatch = line.match(/^(.+?)[+*]\s*(\d+)(.*)$/);
    if (orderMatch && currentItem && currentPrice) {
      const name = orderMatch[1];
      const qty = Number(orderMatch[2]);
      const extra = orderMatch[3] || '';
      const note = extra.includes('辣') ? '辣' : '';

      add(currentItem, currentPrice, name, qty, note);
      continue;
    }

    // 品項 + $金額 + 名字：打拋雞/中辣 $120 大芳
    const inlineSymbol = line.match(/^(.+?)\s*[💰$＄]\s*(\d{1,5})\s*([^\d\s]+)$/);
    if (inlineSymbol) {
      add(inlineSymbol[1], Number(inlineSymbol[2]), inlineSymbol[3], 1);
      itemBuffer = '';
      continue;
    }

    // 品項 + $金額：蔥肉餅+蛋 $55
    const itemSymbol = line.match(/^(.+?)\s*[💰$＄]\s*(\d{1,5})$/);
    if (itemSymbol) {
      currentItem = itemSymbol[1];
      currentPrice = Number(itemSymbol[2]);
      itemCount[clean(currentItem)] = itemCount[clean(currentItem)] || 0;
      itemBuffer = '';
      continue;
    }

    // 品項換行 + 價格名字：焙香烏龍拿鐵熱 無糖(L)+仙草凍 / 70逸軒
    const priceNameOnly = line.match(/^(\d{2,5})\s*([^\d\s]+)$/);
    if (priceNameOnly && itemBuffer) {
      add(itemBuffer, Number(priceNameOnly[1]), priceNameOnly[2], 1);
      itemBuffer = '';
      continue;
    }

    // 品項換行 + 價格單獨一行：桂花烏龍茶... / 55 / 心玄
    const priceOnly = line.match(/^(\d{2,5})$/);
    if (priceOnly && itemBuffer) {
      pendingOrder = {
        item: itemBuffer,
        price: Number(priceOnly[1])
      };
      itemBuffer = '';
      continue;
    }

    // 品項金額名字：熱帶水果茶無糖60美卉
    const inlineNoSymbol = line.match(/^(.+?)(\d{2,5})([^\d\s]+)$/);
    if (inlineNoSymbol && !/[+*]/.test(line)) {
      add(inlineNoSymbol[1], Number(inlineNoSymbol[2]), inlineNoSymbol[3], 1);
      itemBuffer = '';
      continue;
    }

    // 品項金額，名字下一行：熱帶水果茶無糖60 / 慧玲
    const noSymbolNoName = line.match(/^(.+?)(\d{2,5})$/);
    if (noSymbolNoName && !/[+*]/.test(line)) {
      pendingOrder = {
        item: noSymbolNoName[1],
        price: Number(noSymbolNoName[2])
      };
      itemBuffer = '';
      continue;
    }

    // 純文字：可能是品項被換到下一行，先暫存
    if (!/[+*]/.test(line)) {
      itemBuffer = itemBuffer ? itemBuffer + line : line;
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
