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

function parseOrders(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  const result = [];

  let currentItem = '';
  let currentPrice = 0;
  let pendingGroups = [];
  let pendingDrink = null;

  function pushOrder(item, price, user = '', qty = 1, note = '') {
    for (let i = 0; i < qty; i++) {
      result.push({
        item: note ? `${item}（${note}）` : item,
        price,
        user: cleanName(user)
      });
    }
  }

  function getGroup(item) {
    let group = pendingGroups.find(g => g.item === item);
    if (!group) {
      group = { item, lines: [] };
      pendingGroups.push(group);
    }
    return group;
  }

  function flushPendingGroups(price) {
    for (const group of pendingGroups) {
      for (const qLine of group.lines) {
        const qtyMatch = qLine.match(/^(.+?)[+*]\s*(半|\d+)(.*)$/);
        if (!qtyMatch) continue;

        const user = qtyMatch[1];
        const rawQty = qtyMatch[2];
        const qty = rawQty === '半' ? 1 : Number(rawQty);
        const note = rawQty === '半'
          ? `半${qtyMatch[3].trim()}`
          : qtyMatch[3].trim();

        pushOrder(group.item, price, user, qty, note);
      }
    }

    pendingGroups = [];
  }

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    if (/收錢|謝謝|下午|早上|晚上/.test(line)) {
      continue;
    }

    if (pendingDrink && !pendingDrink.user && !/\d/.test(line)) {
      pendingDrink.user = cleanName(line);
      result.push(pendingDrink);
      pendingDrink = null;
      continue;
    }

    if (pendingDrink) {
      result.push(pendingDrink);
      pendingDrink = null;
    }

    // 一個 💰 60 / 一個 $60 / 每個60 / 單價60
    const priceOnlyMatch = line.match(/^(一個|每個|單價)?\s*[💰$＄]?\s*(\d{1,4})$/);
    if (priceOnlyMatch && pendingGroups.length > 0) {
      flushPendingGroups(Number(priceOnlyMatch[2]));
      continue;
    }

    // 品項+價格：蔥肉餅+蛋 $55
    const headerMatch = line.match(/^(.+?)\s*[💰$＄]?\s*(\d{2,4})$/);
    const isQtyLine = /^.+?[+*]\s*(半|\d+)/.test(line);

    if (headerMatch && !isQtyLine) {
      currentItem = headerMatch[1].trim();
      currentPrice = Number(headerMatch[2]);
      continue;
    }

    // 姓名+數量：慧玲+2 / 瑞琴*半
    const qtyMatch = line.match(/^(.+?)[+*]\s*(半|\d+)(.*)$/);

    if (qtyMatch) {
      const user = qtyMatch[1];
      const rawQty = qtyMatch[2];
      const qty = rawQty === '半' ? 1 : Number(rawQty);
      const note = rawQty === '半'
        ? `半${qtyMatch[3].trim()}`
        : qtyMatch[3].trim();

      if (currentItem && currentPrice) {
        pushOrder(currentItem, currentPrice, user, qty, note);
      } else if (currentItem) {
        getGroup(currentItem).lines.push(line);
      }
      continue;
    }

    // 飲料：珍珠奶茶微糖微冰 40士豪
    const drinkMatch = line.match(/^(.+?)\s*(\d{2,4})\s*([^\d\s]+)?$/);

    if (drinkMatch) {
      const item = drinkMatch[1].trim();
      const price = Number(drinkMatch[2]);
      const user = cleanName(drinkMatch[3] || '');

      pendingDrink = { item, price, user };
      continue;
    }

    // 沒數字 = 品項名稱：原味 / 辣味 / 肉鬆
    if (!/\d/.test(line)) {
      currentItem = line;
      currentPrice = 0;
      getGroup(currentItem);
      continue;
    }
  }

  if (pendingDrink) result.push(pendingDrink);

  return result;
}
