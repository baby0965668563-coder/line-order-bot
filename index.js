let isOpen = false;
let orders = [];

app.post('/webhook', line.middleware(config), async (req, res) => {
  const event = req.body.events[0];

  if (event.type !== 'message' || event.message.type !== 'text') {
    return res.sendStatus(200);
  }

  const text = event.message.text;
  let replyText = '';

  const sheet = await getSheet();

  // 🟢 開單
  if (text === '開單') {
    isOpen = true;
    orders = [];
    replyText = '已開單，可以開始點餐🐰';
  }

  // 🔴 結單
  else if (text === '結單') {
    isOpen = false;

    for (let order of orders) {
      await sheet.addRow(order);
    }

    replyText = `已結單，共 ${orders.length} 筆訂單🐰`;
  }

  // 📦 點餐
  else {
    if (!isOpen) {
      replyText = '現在沒有開放訂單❌';
    } else {
      orders.push({
        時間: new Date().toLocaleString(),
        使用者: event.source.userId,
        品項: text,
        金額: ''
      });

      replyText = '已加入訂單🐰';
    }
  }

  await client.replyMessage(event.replyToken, {
    type: 'text',
    text: replyText
  });

  res.sendStatus(200);
});
