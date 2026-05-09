require('dotenv').config();

const express = require('express');
const line = require('@line/bot-sdk');
const { GoogleSpreadsheet } = require('google-spreadsheet');

const app = express();

const config = {
  channelAccessToken: process.env.CHANNEL_ACCESS_TOKEN,
  channelSecret: process.env.CHANNEL_SECRET
};

app.use('/webhook', line.middleware(config));
app.use(express.json());

const client = new line.Client(config);
const doc = new GoogleSpreadsheet(process.env.SHEET_ID);

let isOpen = false;
let allText = '';
const knownUsers = {};

const admins = [
  "U8d9c82446aa9eb90d7de001cfc7ea90f",
  "Ubcfae64b443b9fad21bbc584e991b306",
  "U5c44a04efc62664bd45ec80d77be7d93",
  "Uc669eca67bf477460945f45751edd3e9"
];

function isAdmin(userId) {
  return admins.includes(userId);
}

async function authSheet() {
  await doc.useServiceAccountAuth({
    client_email: process.env.GOOGLE_CLIENT_EMAIL,
    private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n')
  });
  await doc.loadInfo();
}

async function loadMenu() {
  await authSheet();

  const sheet = doc.sheetsByTitle['Menu'];
  if (!sheet) return [];

  const rows = await sheet.getRows();

  return rows
    .map(r => ({
      store: String(r['店家'] || '').trim(),
      item: String(r['品項'] || '').trim(),
      price: Number(r['價格'] || 0)
    }))
    .filter(row =>
      row.store &&
      row.item &&
      row.price > 0
    );
}

async function loadOptions() {
  await authSheet();

  const groupsSheet = doc.sheetsByTitle['OptionGroups'];
  const optionsSheet = doc.sheetsByTitle['Options'];

  if (!groupsSheet || !optionsSheet) return {};

  const groupRows = await groupsSheet.getRows();
  const optionRows = await optionsSheet.getRows();

  const result = {};

  groupRows.forEach(group => {
    const store = String(group['店家'] || '').trim();
    const item = String(group['品項'] || '').trim();
    const category = String(group['分類'] || '').trim();

    if (!store || !item || !category) return;

    const key = store + '||' + item;

    if (!result[key]) result[key] = [];

    const options = optionRows
      .filter(opt =>
        String(opt['店家'] || '').trim() === store &&
        String(opt['品項'] || '').trim() === item &&
        String(opt['分類'] || '').trim() === category
      )
      .map(opt => String(opt['選項'] || '').trim())
      .filter(Boolean);

    result[key].push({
      category,
      required: String(group['必選'] || '').trim() === 'TRUE',
      min: Number(group['最少'] || 0),
      max: Number(group['最多'] || 0),
      options
    });
  });

  return result;
}

async function saveUserToSheet(profileName, userId, sourceType, groupId) {
  try {
    await authSheet();

    const sheet = doc.sheetsByTitle['Users'];
    if (!sheet) return;

    await sheet.addRow({
      時間: new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' }),
      LINE名稱: profileName,
      userId: userId,
      來源類型: sourceType,
      群組ID: groupId || '',
      權限: isAdmin(userId) ? 'admin' : 'user'
    });

    console.log('已寫入 Users');
  } catch (err) {
    console.error('寫入 Users 失敗：', err.message);
  }
}

async function saveOrderToSheet(order) {
  try {
    await authSheet();

    const sheet = doc.sheetsByTitle['Orders'];
    if (!sheet) return false;

    const qty = Number(order.qty || 1);
    const price = Number(order.price || 0);

    await sheet.addRow({
      時間: new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' }),
      LINE名稱: order.name || '',
      userId: order.userId || '',
      店家: order.store || '',
      品項: order.item || '',
      規格: order.spec || '',
      備註: order.note || '',
      數量: qty,
      單價: price,
      總價: price * qty,
      狀態: '未付款'
    });

    return true;
  } catch (err) {
    console.error('寫入 Orders 失敗：', err.message);
    return false;
  }
}

function clean(text) {
  return String(text || '')
    .replace(/[。.,，、!！?？:：;；"'（）()【】\[\]{}<>《》\s]/g, '')
    .trim();
}

function parseOrders(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);

  let currentItem = '';
  let currentPrice = 0;
  let pendingOrder = null;
  let itemBuffer = '';
  let priceMap = {};

  const itemCount = {};
  const userTotal = {};

  function cleanItemName(text) {
    return clean(text).replace(/\d+顆/g, '');
  }

  function add(item, price, name, qty = 1, note = '') {
    item = cleanItemName(item);
    name = clean(name);

    if (!item || !name || !qty) return;

    const finalItem = note ? `${item}（${note}）` : item;
    itemCount[finalItem] = (itemCount[finalItem] || 0) + qty;
    userTotal[name] = (userTotal[name] || 0) + price * qty;
  }

  function getPrice(rawQty) {
    if (priceMap[rawQty] !== undefined) return priceMap[rawQty];
    if (rawQty === '半' && priceMap['0.5'] !== undefined) return priceMap['0.5'];
    return currentPrice;
  }

  for (let line of lines) {
    if (/今天有人要訂|收錢|謝謝|下午|早上|晚上|星期/.test(line)) continue;

    const priceTable = line.match(/^(半|0\.5|1)\s*[💰$＄]\s*(\d{1,5})$/);
    if (priceTable) {
      priceMap[priceTable[1]] = Number(priceTable[2]);
      continue;
    }

    if (pendingOrder && !/[+*]/.test(line) && !/\d/.test(line)) {
      add(pendingOrder.item, pendingOrder.price, line, 1);
      pendingOrder = null;
      continue;
    }

    const orderMatch = line.match(/^(.+?)[+*]\s*(半|0\.5|\d+)(.*)$/);
    if (orderMatch && currentItem) {
      const name = orderMatch[1];
      const rawQty = orderMatch[2];
      const extra = orderMatch[3] || '';
      const note = extra.includes('辣') ? '辣' : '';
      const price = getPrice(rawQty);
      const qty = rawQty === '半' || rawQty === '0.5' ? 1 : Number(rawQty);

      add(currentItem, price, name, qty, note);
      continue;
    }

    const inlineSymbol = line.match(/^(.+?)\s*[💰$＄]\s*(\d{1,5})\s*([^\d\s]+)$/);
    if (inlineSymbol) {
      add(inlineSymbol[1], Number(inlineSymbol[2]), inlineSymbol[3], 1);
      itemBuffer = '';
      continue;
    }

    const itemSymbol = line.match(/^(.+?)\s*[💰$＄]\s*(\d{1,5})$/);
    if (itemSymbol) {
      currentItem = itemSymbol[1];
      currentPrice = Number(itemSymbol[2]);
      itemBuffer = '';
      continue;
    }

    const inlineNoSymbol = line.match(/^(.+?)(\d{2,5})([^\d\s]+)$/);
    if (inlineNoSymbol && !/[+*]/.test(line)) {
      add(inlineNoSymbol[1], Number(inlineNoSymbol[2]), inlineNoSymbol[3], 1);
      itemBuffer = '';
      continue;
    }

    const noSymbolNoName = line.match(/^(.+?)(\d{2,5})$/);
    if (noSymbolNoName && !/[+*]/.test(line) && !/顆/.test(line)) {
      pendingOrder = {
        item: noSymbolNoName[1],
        price: Number(noSymbolNoName[2])
      };
      itemBuffer = '';
      continue;
    }

    const priceNameOnly = line.match(/^(\d{2,5})\s*([^\d\s]+)$/);
    if (priceNameOnly && itemBuffer) {
      add(itemBuffer, Number(priceNameOnly[1]), priceNameOnly[2], 1);
      itemBuffer = '';
      continue;
    }

    const priceOnly = line.match(/^(\d{2,5})$/);
    if (priceOnly && itemBuffer) {
      pendingOrder = {
        item: itemBuffer,
        price: Number(priceOnly[1])
      };
      itemBuffer = '';
      continue;
    }

    if (!/[+*]/.test(line)) {
      if (/顆/.test(line)) {
        currentItem = line;
        currentPrice = 0;
        itemBuffer = line;
        continue;
      }

      if (currentItem) {
        add(currentItem, currentPrice, line, 1);
        continue;
      }

      currentItem = line;
      currentPrice = 0;
      itemBuffer = line;
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

function formatShopOrder(itemCount, userTotal) {
  let orderText = '您好，今天訂購如下：\n\n';
  let totalCount = 0;

  for (let item in itemCount) {
    const qty = itemCount[item];

    if (qty > 0) {
      orderText += `${item} x${qty}\n`;
      totalCount += qty;
    }
  }

  const totalMoney = Object.values(userTotal).reduce((a, b) => a + b, 0);

  orderText += `\n總數：${totalCount}份`;
  orderText += `\n總金額：${totalMoney}元`;
  orderText += '\n\n麻煩您，謝謝～';

  return orderText;
}

function safeJson(data) {
  return JSON.stringify(data)
    .replace(/</g, '\\u003c')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');
}

app.get('/', (req, res) => {
  res.send('LINE 訂餐統計機器人運作中');
});

app.get('/order', async (req, res) => {
  try {
    const menu = await loadMenu();
    const optionData = await loadOptions();

    const menuJson = safeJson(menu);
    const optionJson = safeJson(optionData);

    const html = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>訂餐小幫手</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <script src="https://static.line-scdn.net/liff/edge/2/sdk.js"></script>
  <style>
    body {
      margin: 0;
      font-family: Arial, "Microsoft JhengHei", sans-serif;
      background: #f6f3ee;
      color: #333;
    }

    .header {
      padding: 20px;
      background: #fff;
      border-bottom: 1px solid #eee;
      position: sticky;
      top: 0;
      z-index: 10;
    }

    .header h1 {
      margin: 0;
      font-size: 28px;
    }

    .header p {
      margin: 8px 0 0;
      color: #777;
      font-size: 15px;
    }

    .status {
      margin-top: 12px;
      font-size: 15px;
      color: #06c755;
    }

    .container {
      padding: 16px;
    }

    .tabs {
      background: #f6f3ee;
      position: sticky;
      top: 118px;
      z-index: 9;
      padding: 10px 12px;
      overflow-x: auto;
      white-space: nowrap;
      border-bottom: 1px solid #eee;
    }

    .tab-btn {
      width: auto;
      display: inline-block;
      margin: 4px;
      padding: 10px 14px;
      border-radius: 999px;
      background: #fff;
      color: #333;
      border: 1px solid #ddd;
      font-size: 14px;
    }

    .store-title {
      font-size: 22px;
      font-weight: bold;
      margin: 24px 0 12px;
      padding: 10px 4px;
    }

    .card {
      background: #fff;
      border-radius: 16px;
      padding: 16px;
      margin-bottom: 12px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }

    .store {
      font-size: 13px;
      color: #999;
      margin-bottom: 6px;
    }

    .item {
      font-size: 20px;
      font-weight: bold;
      margin-bottom: 8px;
    }

    .price {
      font-size: 17px;
      margin-bottom: 12px;
    }

    button {
      width: 100%;
      padding: 13px;
      border: none;
      border-radius: 999px;
      background: #06c755;
      color: white;
      font-size: 16px;
      font-weight: bold;
    }

    button:disabled {
      background: #aaa;
    }

    .empty {
      text-align: center;
      color: #999;
      margin-top: 60px;
    }

    .modal {
      display: none;
      position: fixed;
      top: 0;
      left: 0;
      width: 100%;
      height: 100%;
      background: rgba(0,0,0,0.5);
      z-index: 9999;
      justify-content: center;
      align-items: center;
    }

    .modal-box {
      background: white;
      width: 90%;
      max-width: 420px;
      max-height: 80vh;
      overflow-y: auto;
      border-radius: 20px;
      padding: 20px;
    }

    .option-label {
      display: block;
      margin: 10px 0;
      font-size: 16px;
    }

    .qty-box {
      margin-top: 16px;
    }

    .qty-box select {
      width: 100%;
      padding: 12px;
      border-radius: 12px;
      border: 1px solid #ddd;
      font-size: 16px;
    }

    .cancel-btn {
      background: #999;
      margin-top: 10px;
    }
  </style>
</head>

<body>
  <div class="header">
    <h1>訂餐小幫手</h1>
    <p>請選擇今天要訂的餐點</p>
    <div class="status" id="status">正在取得 LINE 使用者資料...</div>
  </div>

  <div class="tabs" id="storeTabs"></div>
  <div class="container" id="menu"></div>

  <div class="modal" id="optionModal">
    <div class="modal-box">
      <h2 id="modalTitle"></h2>
      <div id="modalOptions"></div>

      <div class="qty-box">
        <div style="margin-bottom:8px;font-weight:bold;">數量</div>
        <select id="qtySelect">
          <option value="1">1份</option>
          <option value="2">2份</option>
          <option value="3">3份</option>
          <option value="4">4份</option>
          <option value="5">5份</option>
          <option value="6">6份</option>
          <option value="7">7份</option>
          <option value="8">8份</option>
          <option value="9">9份</option>
          <option value="10">10份</option>
        </select>
      </div>

      <button onclick="submitOptions()">確認加入</button>
      <button class="cancel-btn" onclick="closeModal()">取消</button>
    </div>
  </div>

<script>
const menu = ${menuJson};
const optionData = ${optionJson};
const LIFF_ID = '2010025093-yATK02dc';

let profile = null;
let liffReady = false;
let currentItem = null;
let currentGroups = [];

function renderMenu() {
  const box = document.getElementById('menu');
  const tabsBox = document.getElementById('storeTabs');

  if (!menu || menu.length === 0) {
    box.innerHTML = '<div class="empty">目前沒有菜單資料</div>';
    return;
  }

  const stores = [...new Set(menu.map(m => m.store).filter(Boolean))];

  tabsBox.innerHTML = stores.map((store, i) =>
    '<button class="tab-btn" onclick="scrollToStore(' + i + ')">' +
      store +
    '</button>'
  ).join('');

  let html = '';
  let currentStore = '';

  menu.forEach((m, index) => {
    if (m.store !== currentStore) {
      currentStore = m.store;
      const storeIndex = stores.indexOf(currentStore);

      html +=
        '<div id="store-' + storeIndex + '" class="store-title">' +
          currentStore +
        '</div>';
    }

    html +=
      '<div class="card">' +
        '<div class="store">' + (m.store || '') + '</div>' +
        '<div class="item">' + (m.item || '') + '</div>' +
        '<div class="price">$' + (m.price || 0) + '</div>' +
        '<button onclick="addOrder(' + index + ')" id="btn-' + index + '" disabled>' +
          '載入中...' +
        '</button>' +
      '</div>';
  });

  box.innerHTML = html;
}

function scrollToStore(index) {
  const el = document.getElementById('store-' + index);

  if (el) {
    el.scrollIntoView({
      behavior: 'smooth',
      block: 'start'
    });
  }
}

function enableButtons() {
  menu.forEach((_, index) => {
    const btn = document.getElementById('btn-' + index);

    if (btn) {
      btn.disabled = false;
      btn.innerText = '加入訂單';
    }
  });
}

async function initLIFF() {
  try {
    await liff.init({
      liffId: LIFF_ID
    });

    if (!liff.isLoggedIn()) {
      liff.login();
      return;
    }

    profile = await liff.getProfile();
    liffReady = true;

    document.getElementById('status').innerText =
      '已登入：' + (profile.displayName || '');

    enableButtons();

  } catch (err) {
    console.error(err);

    document.getElementById('status').innerText =
      'LIFF 初始化失敗，請重新整理';
  }
}

function closeModal() {
  document.getElementById('optionModal').style.display = 'none';
}

async function addOrder(index) {
  if (!liffReady || !profile) {
    alert('尚未取得 LINE 使用者資料');
    return;
  }

  currentItem = menu[index];

  const key = currentItem.store + '||' + currentItem.item;
  currentGroups = optionData[key] || [];

  document.getElementById('modalTitle').innerText = currentItem.item;
  document.getElementById('qtySelect').value = '1';

  const optionBox = document.getElementById('modalOptions');
  optionBox.innerHTML = '';

  if (currentGroups.length === 0) {
    optionBox.innerHTML =
      '<div style="color:#777;margin:10px 0;">此商品無需選擇規格</div>';
  } else {
    currentGroups.forEach((group, groupIndex) => {
      const title = document.createElement('div');

      title.style.marginTop = '16px';
      title.innerHTML =
        '<b>' +
        group.category +
        '</b><br>' +
        '請選 ' +
        group.min +
        ' ~ ' +
        group.max +
        ' 個';

      optionBox.appendChild(title);

      group.options.forEach(opt => {
        const label = document.createElement('label');
        label.className = 'option-label';

        label.innerHTML =
          '<input ' +
          'type="checkbox" ' +
          'value="' + opt + '" ' +
          'data-group="' + groupIndex + '" ' +
          'onchange="limitCheck(' + groupIndex + ',' + group.max + ')"' +
          '> ' +
          opt;

        optionBox.appendChild(label);
      });
    });
  }

  document.getElementById('optionModal').style.display = 'flex';
}

function limitCheck(groupIndex, max) {
  const checked = [
    ...document.querySelectorAll(
      'input[data-group="' + groupIndex + '"]:checked'
    )
  ];

  const all = [
    ...document.querySelectorAll(
      'input[data-group="' + groupIndex + '"]'
    )
  ];

  if (checked.length >= max) {
    all.forEach(x => {
      if (!x.checked) x.disabled = true;
    });
  } else {
    all.forEach(x => {
      x.disabled = false;
    });
  }
}

async function submitOptions() {
  let specText = '';

  for (let i = 0; i < currentGroups.length; i++) {
    const group = currentGroups[i];

    const checked = [
      ...document.querySelectorAll(
        'input[data-group="' + i + '"]:checked'
      )
    ];

    if (
      checked.length < group.min ||
      checked.length > group.max
    ) {
      alert(
        group.category +
        ' 需要選 ' +
        group.min +
        ' ~ ' +
        group.max +
        ' 個'
      );

      return;
    }

    const values = checked.map(x => x.value);

    specText +=
      group.category +
      '：' +
      values.join('、') +
      ' ';
  }

  document.getElementById('optionModal').style.display = 'none';

  submitFinalOrder(specText.trim());
}

async function submitFinalOrder(specText) {
  const qty = Number(document.getElementById('qtySelect').value || 1);

  const orderData = {
    name: profile.displayName || '',
    userId: profile.userId || '',
    store: currentItem.store || '',
    item: currentItem.item || '',
    spec: specText || '',
    note: '',
    qty: qty,
    price: currentItem.price || 0
  };

  try {
    const res = await fetch('/api/order', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(orderData)
    });

    const result = await res.json();

    if (result.success) {
      alert('已加入訂單');
    } else {
      alert('加入失敗');
    }

  } catch (err) {
    alert('送出失敗：' + err.message);
  }
}

try {
  renderMenu();
} catch (err) {
  console.error('renderMenu error:', err);
}

initLIFF();
</script>

</body>
</html>
`;

    res.send(html);
  } catch (err) {
    console.error('載入訂餐頁失敗：', err.message);
    res.send('載入訂餐頁失敗，請稍後再試');
  }
});

app.post('/api/order', async (req, res) => {
  console.log('收到訂單');
  console.log(req.body);

  const success = await saveOrderToSheet(req.body);

  res.json({ success });
});

app.post('/webhook', async (req, res) => {
  try {
    const event = req.body.events?.[0];

    if (!event || event.type !== 'message') {
      return res.sendStatus(200);
    }

    let profileName = '未知使用者';

    try {
      if (event.source.type === 'group') {
        const profile = await client.getGroupMemberProfile(
          event.source.groupId,
          event.source.userId
        );

        profileName = profile.displayName;
      } else {
        const profile = await client.getProfile(event.source.userId);
        profileName = profile.displayName;
      }
    } catch (err) {
      console.error('取得使用者名稱失敗：', err.message);
    }

    knownUsers[profileName] = event.source.userId;

    await saveUserToSheet(
      profileName,
      event.source.userId,
      event.source.type,
      event.source.groupId || ''
    );

    const userId = event.source.userId;
    const text = event.message.text ? event.message.text.trim() : '';

    if (text === '開單') {
      if (!isAdmin(userId)) {
        await client.replyMessage(event.replyToken, {
          type: 'text',
          text: '只有管理員可以開單'
        });

        return res.sendStatus(200);
      }

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

    if (text === '清空') {
      if (!isAdmin(userId)) {
        await client.replyMessage(event.replyToken, {
          type: 'text',
          text: '只有管理員可以清空'
        });

        return res.sendStatus(200);
      }

      allText = '';
      isOpen = false;

      await client.replyMessage(event.replyToken, {
        type: 'text',
        text: '已清空訂單'
      });

      return res.sendStatus(200);
    }

    if (text === '店家單') {
      if (!isAdmin(userId)) {
        await client.replyMessage(event.replyToken, {
          type: 'text',
          text: '只有管理員可以查看店家單'
        });

        return res.sendStatus(200);
      }

      const result = parseOrders(allText);
      const reply = formatShopOrder(result.itemCount, result.userTotal);

      await client.replyMessage(event.replyToken, {
        type: 'text',
        text: reply
      });

      return res.sendStatus(200);
    }

    if (text === '結單' || text === '收單' || text === '統計') {
      if (!isAdmin(userId)) {
        await client.replyMessage(event.replyToken, {
          type: 'text',
          text: '只有管理員可以結單 / 統計'
        });

        return res.sendStatus(200);
      }

      const result = parseOrders(allText);
      const reply = formatResult(result.itemCount, result.userTotal);

      isOpen = false;

      await client.replyMessage(event.replyToken, {
        type: 'text',
        text: reply
      });

      return res.sendStatus(200);
    }

    if (isOpen && event.message.type === 'text') {
      allText += '\n' + text;
    }

    return res.sendStatus(200);

  } catch (error) {
    console.error('Webhook error:', error);
    return res.sendStatus(200);
  }
});

app.use((err, req, res, next) => {
  console.error('Global error:', err);
  res.sendStatus(200);
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log('Server running on ' + PORT);
});
