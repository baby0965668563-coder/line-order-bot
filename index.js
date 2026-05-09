require('dotenv').config();

const express = require('express');
const line = require('@line/bot-sdk');
const { GoogleSpreadsheet } = require('google-spreadsheet');

const app = express();

app.use(express.json());

const config = {
  channelAccessToken: process.env.CHANNEL_ACCESS_TOKEN,
  channelSecret: process.env.CHANNEL_SECRET
};

const client = new line.Client(config);

const doc = new GoogleSpreadsheet(process.env.SHEET_ID);

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

  return rows.map(r => ({
    store: r['店家'],
    item: r['品項'],
    price: r['價格']
  }));
}

async function saveOrderToSheet(order) {

  try {

    await authSheet();

    const sheet = doc.sheetsByTitle['Orders'];

    if (!sheet) {
      console.log('找不到 Orders');
      return false;
    }

    await sheet.addRow({
      時間: new Date().toLocaleString('zh-TW', {
        timeZone: 'Asia/Taipei'
      }),
      LINE名稱: order.name || '',
      userId: order.userId || '',
      店家: order.store || '',
      品項: order.item || '',
      規格: order.spec || '',
      備註: order.note || '',
      數量: order.qty || 1,
      單價: order.price || 0,
      總價: Number(order.price || 0) * Number(order.qty || 1),
      狀態: '未付款'
    });

    console.log('成功寫入 Orders');

    return true;

  } catch (err) {

    console.log(err);

    return false;
  }
}

app.get('/', (req, res) => {
  res.send('LINE 訂餐機器人運作中');
});

app.get('/order', async (req, res) => {

  const menu = await loadMenu();

  let html = `
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>訂餐小幫手</title>

<script src="https://static.line-scdn.net/liff/edge/2/sdk.js"></script>

<style>

body{
  margin:0;
  background:#f6f3ee;
  font-family:Arial;
}

.header{
  padding:20px;
  background:white;
}

.container{
  padding:16px;
}

.card{
  background:white;
  border-radius:16px;
  padding:16px;
  margin-bottom:12px;
}

.store{
  color:#999;
  font-size:13px;
}

.item{
  font-size:20px;
  font-weight:bold;
  margin-top:6px;
}

.price{
  margin-top:8px;
  font-size:18px;
}

button{
  margin-top:12px;
  width:100%;
  border:none;
  border-radius:999px;
  padding:12px;
  background:#06c755;
  color:white;
  font-size:16px;
}

</style>

</head>

<body>

<div class="header">
<h2>訂餐小幫手</h2>
</div>

<div class="container">
`;

  menu.forEach(item => {

    html += `
<div class="card">

<div class="store">${item.store}</div>

<div class="item">${item.item}</div>

<div class="price">$${item.price}</div>

<button onclick="addOrder(
'${item.store}',
'${item.item}',
'${item.price}'
)">
加入訂單
</button>

</div>
`;
  });

  html += `
</div>

<script>

let profile = null;

async function initLIFF(){

  await liff.init({
    liffId:'2010025093-yATK02dc'
  });

  if(!liff.isLoggedIn()){
    liff.login();
    return;
  }

  profile = await liff.getProfile();

  console.log(profile);
}

async function addOrder(store,item,price){

  alert('按鈕有反應');

  const orderData = {
    name: profile?.displayName || '',
    userId: profile?.userId || '',
    store: store,
    item: item,
    spec: '',
    note: '',
    qty: 1,
    price: price
  };

  try{

    const res = await fetch('/api/order',{
      method:'POST',
      headers:{
        'Content-Type':'application/json'
      },
      body:JSON.stringify(orderData)
    });

    const result = await res.json();

    console.log(result);

    if(result.success){
      alert('已加入訂單');
    }else{
      alert('加入失敗');
    }

  }catch(err){

    alert('錯誤：'+err.message);

  }
}

initLIFF();

</script>

</body>
</html>
`;

  res.send(html);
});

app.post('/api/order', async (req, res) => {

  console.log('收到訂單');

  console.log(req.body);

  const success = await saveOrderToSheet(req.body);

  if(success){

    res.json({
      success:true
    });

  }else{

    res.status(500).json({
      success:false
    });

  }
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log('Server running on ' + PORT);
});
