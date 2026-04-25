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

  const itemCount = {};
  const userTotal = {};

  function addOrder(item, price, name, qty = 1) {
    item = String(item || '').trim();
    name = String(name || '').replace(/[。.,，\s]/g, '').trim();
    if (!item || !price || !name) return;

    itemCount[item] = (itemCount[item] || 0) + qty;
    userTotal[name] = (userTotal[name] || 0) + price * qty;
  }

  for (let rawLine of lines) {
    let line = rawLine.trim();
    if (!line) continue;

    if (/今天有人要訂|收錢|謝謝|下午|早上|晚上/.test(line)) {
      continue;
    }

    // 上一行是飲料但沒名字，這一行當名字
    if (pendingDrink && !/\d/.test(line) && !/[+*]/.test(line)) {
      addOrder(pendingDrink.item, pendingDrink.price, line, 1);
      pendingDrink = null;
      continue;
    }

    // 品項 + $價格：蔥肉餅 $45 / 紅豆餅💰25
    const itemMatch = line.match(/^(.+?)\s*[💰$＄]\s*(\d+)/);
    if (itemMatch) {
      if (pendingDrink) pendingDrink = null;
      currentItem = itemMatch[1].trim();
      currentPrice = parseInt(itemMatch[2]);
      if (!itemCount[currentItem]) itemCount[currentItem] = 0;
      continue;
    }

    // 人名 + 數量：芷葳+2 / 心玄*3
    const orderMatch = line.match(/^(.+?)[+*](\d+)/);
    if (orderMatch && currentItem) {
      const name = orderMatch[1].trim();
      const qty = parseInt(orderMatch[2]);
      addOrder(currentItem, currentPrice, name, qty);
      continue;
    }

    // 飲料：品項 40名字 / 品項40名字 / 品項 55
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
