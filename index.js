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

  const groupsSheet =
    doc.sheetsByTitle['OptionGroups'];

  const optionsSheet =
    doc.sheetsByTitle['Options'];

  if (!groupsSheet || !optionsSheet) {
    return {};
  }

  const groupRows =
    await groupsSheet.getRows();

  const optionRows =
    await optionsSheet.getRows();

  const result = {};

  groupRows.forEach(group => {

    const store =
      String(group['店家'] || '').trim();

    const item =
      String(group['品項'] || '').trim();

    const category =
      String(group['分類'] || '').trim();

    const min =
      Number(group['最少'] || 0);

    const max =
      Number(group['最多'] || 0);

    const key = store + '||' + item;

    if (!result[key]) {
      result[key] = [];
    }

    const options = optionRows
      .filter(opt =>
        String(opt['店家']).trim() === store &&
        String(opt['品項']).trim() === item &&
        String(opt['分類']).trim() === category
      )
      .map(opt =>
        String(opt['選項']).trim()
      );

    result[key].push({
      category,
      min,
      max,
      options
    });

  });

  return result;

}

async function saveOrderToSheet(order) {

  try {

    await authSheet();

    const sheet =
      doc.sheetsByTitle['Orders'];

    if (!sheet) return false;

    const qty =
      Number(order.qty || 1);

    const price =
      Number(order.price || 0);

    await sheet.addRow({

      時間:
        new Date().toLocaleString(
          'zh-TW',
          { timeZone: 'Asia/Taipei' }
        ),

      LINE名稱:
        order.name || '',

      userId:
        order.userId || '',

      店家:
        order.store || '',

      品項:
        order.item || '',

      規格:
        order.spec || '',

      備註:
        order.note || '',

      數量:
        qty,

      單價:
        price,

      總價:
        qty * price,

      狀態:
        '未付款'

    });

    return true;

  } catch (err) {

    console.error(err);

    return false;

  }

}

app.get('/', (req, res) => {
  res.send('LINE 訂餐系統運作中');
});

app.get('/order', async (req, res) => {

  try {

    const menu =
      await loadMenu();

    const optionData =
      await loadOptions();

    const menuJson =
      JSON.stringify(menu).replace(/</g, '\\u003c');

    const optionJson =
      JSON.stringify(optionData).replace(/</g, '\\u003c');

    res.send(`

<!DOCTYPE html>
<html>

<head>

<meta charset="utf-8">

<title>訂餐小幫手</title>

<meta
  name="viewport"
  content="width=device-width, initial-scale=1"
/>

<script src="https://static.line-scdn.net/liff/edge/2/sdk.js"></script>

<style>

body{
  margin:0;
  background:#f6f3ee;
  font-family:Arial;
}

.header{
  background:white;
  padding:20px;
  position:sticky;
  top:0;
  z-index:10;
}

.container{
  padding:16px;
}

.card{
  background:white;
  border-radius:16px;
  padding:16px;
  margin-bottom:14px;
}

.store{
  color:#999;
  font-size:13px;
}

.item{
  font-size:24px;
  font-weight:bold;
  margin-top:8px;
}

.price{
  font-size:18px;
  margin-top:10px;
}

button{
  width:100%;
  border:none;
  background:#06c755;
  color:white;
  padding:14px;
  border-radius:999px;
  font-size:16px;
  margin-top:14px;
}

.modal{
  display:none;
  position:fixed;
  top:0;
  left:0;
  width:100%;
  height:100%;
  background:rgba(0,0,0,0.5);
  z-index:9999;
  justify-content:center;
  align-items:center;
}

.modal-box{
  background:white;
  width:90%;
  max-width:420px;
  max-height:80vh;
  overflow-y:auto;
  border-radius:20px;
  padding:20px;
}

.option-label{
  display:block;
  margin:10px 0;
}

</style>

</head>

<body>

<div class="header">

<h1>訂餐小幫手</h1>

<p>請選擇今天要訂的餐點</p>

<div id="status">
正在取得 LINE 使用者資料...
</div>

</div>

<div class="container" id="storeTabs"></div>

<div class="container" id="menu"></div>

<div
  class="modal"
  id="optionModal"
>

<div class="modal-box">

<h2 id="modalTitle"></h2>

<div id="modalOptions"></div>
<div style="margin-top:16px;">
  <div style="margin-bottom:8px;font-weight:bold;">
    數量
  </div>

  <select id="qtySelect" style="
    width:100%;
    padding:12px;
    border-radius:12px;
    border:1px solid #ddd;
    font-size:16px;
  ">
    <option value="1">1份</option>
    <option value="2">2份</option>
    <option value="3">3份</option>
    <option value="4">4份</option>
    <option value="5">5份</option>
  </select>
</div>

<button onclick="submitOptions()">
確認加入
</button>

</div>

</div>

<script>

const menu = ${menuJson};

const optionData = ${optionJson};

const LIFF_ID =
'2010025093-yATK02dc';

let profile = null;

let liffReady = false;

let currentItem = null;

let currentGroups = [];

function renderMenu(){

  const box =
    document.getElementById('menu');

  if(!menu.length){

    box.innerHTML =
      '<div class="empty">目前沒有菜單資料</div>';

    return;
  }

  let html = '';
  let currentStore = '';

  menu.forEach((m,index)=>{

    if(m.store !== currentStore){

      currentStore = m.store;

      html +=
        '<div style="' +
        'font-size:22px;' +
        'font-weight:bold;' +
        'margin:24px 0 12px;' +
        'padding:10px 4px;' +
        '">' +
        currentStore +
        '</div>';
    }

    html +=
      '<div class="card">' +

        '<div class="store">' +
          m.store +
        '</div>' +

        '<div class="item">' +
          m.item +
        '</div>' +

        '<div class="price">$' +
          m.price +
        '</div>' +

        '<button ' +
          'onclick="addOrder(' + index + ')" ' +
          'id="btn-' + index + '" ' +
          'disabled' +
        '>' +
          '載入中...' +
        '</button>' +

      '</div>';

  });

  box.innerHTML = html;

}

function enableButtons(){

  menu.forEach((_,index)=>{

    const btn =
      document.getElementById(
        'btn-'+index
      );

    if(btn){

      btn.disabled = false;

      btn.innerText =
        '加入訂單';

    }

  });

}

async function initLIFF(){

  try{

    await liff.init({
      liffId: LIFF_ID
    });

    if(!liff.isLoggedIn()){

      liff.login();

      return;

    }

    profile =
      await liff.getProfile();

    liffReady = true;

    document.getElementById(
      'status'
    ).innerText =
      '已登入：' +
      profile.displayName;

    enableButtons();

  }catch(err){

    console.error(err);

    document.getElementById(
      'status'
    ).innerText =
      'LIFF 初始化失敗';

  }

}

async function addOrder(index){

  if(!liffReady){

    alert('LIFF 尚未完成');

    return;

  }

  currentItem =
    menu[index];

  const key =
    currentItem.store +
    '||' +
    currentItem.item;

  currentGroups =
    optionData[key] || [];

  if(currentGroups.length === 0){

  document.getElementById('modalTitle').innerText =
    currentItem.item;

  document.getElementById('modalOptions').innerHTML =
    '<div style="color:#777;">此商品無需選擇規格</div>';

  document.getElementById('optionModal').style.display =
    'block';

  return;

}

  document.getElementById(
    'modalTitle'
  ).innerText =
    currentItem.item;

  const optionBox =
    document.getElementById(
      'modalOptions'
    );

  optionBox.innerHTML = '';

  currentGroups.forEach(
  (group,i)=>{

    const title =
      document.createElement('div');

    title.style.marginTop =
      '16px';

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

group.options.forEach(opt=>{

  const label =
    document.createElement('label');

  label.className =
    'option-label';

 label.innerHTML =
  '<input ' +
  'type="checkbox" ' +
  'value="' + opt + '" ' +
  'data-group="' + i + '" ' +
  'onchange="limitCheck(' + i + ',' + group.max + ')"' +
  '> ' +
  opt;

  optionBox.appendChild(label);

});

  });

  document.getElementById(
    'optionModal'
  ).style.display = 'block';

}

function limitCheck(groupIndex, max){

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

  if(checked.length >= max){

    all.forEach(x=>{

      if(!x.checked){
        x.disabled = true;
      }

    });

  }else{

    all.forEach(x=>{
      x.disabled = false;
    });

  }

}
async function submitOptions(){

  let specText = '';

  for(
    let i=0;
    i<currentGroups.length;
    i++
  ){

    const group =
      currentGroups[i];

    const checked = [
      ...document.querySelectorAll(
        'input[data-group="'+i+'"]:checked'
      )
    ];

    if(
      checked.length < group.min ||
      checked.length > group.max
    ){

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

    const values =
      checked.map(x=>x.value);

    specText +=
      group.category +
      '：' +
      values.join('、') +
      ' ';

  }

  document.getElementById(
    'optionModal'
  ).style.display = 'none';

  submitFinalOrder(
    specText.trim()
  );

}

async function submitFinalOrder(specText){

  const orderData = {

    name:
      profile.displayName || '',

    userId:
      profile.userId || '',

    store:
      currentItem.store || '',

    item:
      currentItem.item || '',

    spec:
      specText || '',

    note:'',

    qty: Number(document.getElementById('qtySelect').value || 1),

    price:
      currentItem.price || 0

  };

  try{

    const res =
      await fetch('/api/order',{

      method:'POST',

      headers:{
        'Content-Type':
        'application/json'
      },

      body:
        JSON.stringify(orderData)

    });

    const result =
      await res.json();

    if(result.success){

      alert('已加入訂單');

    }else{

      alert('加入失敗');

    }

  }catch(err){

    alert(
      '送出失敗：' +
      err.message
    );

  }

}

renderMenu();

initLIFF();

</script>

</body>
</html>

`);

  } catch (err) {

    console.error(err);

    res.send('載入訂餐頁失敗');

  }

});

app.post('/api/order', async (req, res) => {

  const success =
    await saveOrderToSheet(req.body);

  res.json({
    success
  });

});

app.post('/webhook', async (req, res) => {

  try {

    const event =
      req.body.events?.[0];

    if (!event) {
      return res.sendStatus(200);
    }

    if (event.type !== 'message') {
      return res.sendStatus(200);
    }

    const userId =
      event.source.userId;

    const text =
      event.message.text || '';

    if (
      text === '開單'
    ) {

      if (!isAdmin(userId)) {

        await client.replyMessage(
          event.replyToken,
          {
            type: 'text',
            text: '只有管理員可以開單'
          }
        );

        return res.sendStatus(200);

      }

      isOpen = true;

      allText = '';

      await client.replyMessage(
        event.replyToken,
        {
          type: 'text',
          text: '已開單'
        }
      );

      return res.sendStatus(200);

    }

    if (
      text === '清空'
    ) {

      if (!isAdmin(userId)) {

        await client.replyMessage(
          event.replyToken,
          {
            type: 'text',
            text: '只有管理員可以清空'
          }
        );

        return res.sendStatus(200);

      }

      isOpen = false;

      allText = '';

      await client.replyMessage(
        event.replyToken,
        {
          type: 'text',
          text: '已清空'
        }
      );

      return res.sendStatus(200);

    }

    if (
      isOpen &&
      event.message.type === 'text'
    ) {

      allText += '\n' + text;

    }

    res.sendStatus(200);

  } catch (err) {

    console.error(err);

    res.sendStatus(200);

  }

});

const PORT =
  process.env.PORT || 3000;

app.listen(PORT, () => {

  console.log(
    'Server running on ' + PORT
  );

});
