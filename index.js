function parseOrders(text) {
  const lines = text.split('\n');

  let currentItem = null;
  let currentPrice = 0;

  const itemCount = {};
  const userTotal = {};

  for (let rawLine of lines) {
    let line = rawLine.trim();
    if (!line) continue;

    // ✅ 抓「品項 + 價格」(支援 $ 💰 ＄)
    const itemMatch = line.match(/^(.+?)\s*[💰$＄]\s*(\d+)/);
    if (itemMatch) {
      currentItem = itemMatch[1].trim();
      currentPrice = parseInt(itemMatch[2]);

      if (!itemCount[currentItem]) {
        itemCount[currentItem] = 0;
      }
      continue;
    }

    // ✅ 抓「名字 + 數量」(+ 或 *)
    const orderMatch = line.match(/^(.+?)[+*](\d+)/);
    if (orderMatch && currentItem) {
      const name = orderMatch[1].trim();
      const qty = parseInt(orderMatch[2]);

      // 累加品項數量
      itemCount[currentItem] += qty;

      // 累加個人金額
      if (!userTotal[name]) userTotal[name] = 0;
      userTotal[name] += qty * currentPrice;

      continue;
    }

    // ❌ 其他全部忽略（時間、備註、垃圾字）
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
    text += `${user} : $${userTotal[user]}\n`;
  }

  const total = Object.values(userTotal).reduce((a, b) => a + b, 0);
  text += `\n\n💰 總金額：$${total}`;

  return text;
}
