require('dotenv').config(); 

const express = require('express'); const line = require('@line/bot-sdk'); const { GoogleSpreadsheet } = require('google-spreadsheet'); 

const app = express(); 

const lineConfig = { channelAccessToken: process.env.CHANNEL_ACCESS_TOKEN, channelSecret: process.env.CHANNEL_SECRET }; 

// ★ LINE webhook 必須在 express.json() 之前 app.use('/webhook', line.middleware(lineConfig)); app.use('/api', express.json()); app.use('/admin', express.json()); 

const client = new line.Client(lineConfig); const doc = new GoogleSpreadsheet(process.env.SHEET_ID); 

// ════════════════════════════════════════════════════════════════ // 全域狀態 // ════════════════════════════════════════════════════════════════ let isOpen = false; let autoCloseTimer = null; let autoCloseAt = null; const knownUsers = {}; 

const admins = [ 'U8d9c82446aa9eb90d7de001cfc7ea90f', 'Ubcfae64b443b9fad21bbc584e991b306', 'U5c44a04efc62664bd45ec80d77be7d93', 'Uc669eca67bf477460945f45751edd3e9' ]; function isAdmin(uid) { return admins.includes(uid); } 

// ════════════════════════════════════════════════════════════════ // 工具 // ════════════════════════════════════════════════════════════════ function nowTW() { return new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' }); } function todayTW() { return new Date().toLocaleDateString('zh-TW', { timeZone: 'Asia/Taipei' }); } 

// ════════════════════════════════════════════════════════════════ // Google Sheets 認證 // ════════════════════════════════════════════════════════════════ async function authSheet() { await doc.useServiceAccountAuth({ client_email: process.env.GOOGLE_CLIENT_EMAIL, private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\n/g, '\n') }); await doc.loadInfo(); } 

// ════════════════════════════════════════════════════════════════ // 菜單 // ════════════════════════════════════════════════════════════════ async function loadMenu() { await authSheet(); const sheet = doc.sheetsByTitle['Menu']; if (!sheet) { console.error('loadMenu: Menu sheet not found'); return []; } const rows = await sheet.getRows(); return rows .map(r => ({ store: String(r['店家'] || '').trim(), item: String(r['品項'] || '').trim(), price: parseFloat(String(r['價格'] || '0').replace(/[^0-9.]/g, '')) || 0, image: String(r['圖片URL'] || '').trim() })) .filter(r => r.store && r.item && r.price > 0); } 

async function loadOptions() { await authSheet(); const gs = doc.sheetsByTitle['OptionGroups']; const os = doc.sheetsByTitle['Options']; if (!gs || !os) return {}; const gr = await gs.getRows(); const or = await os.getRows(); const result = {}; gr.forEach(g => { const store = String(g['店家'] || '').trim(); const item = String(g['品項'] || '').trim(); const cat = String(g['分類'] || '').trim(); if (!store || !item || !cat) return; const key = store + '||' + item; if (!result[key]) result[key] = []; const opts = or .filter(o => String(o['店家']||'').trim()===store && String(o['品項']||'').trim()===item && String(o['分類']||'').trim()===cat) .map(o => ({ name: String(o['選項'] || '').trim(), extra: parseFloat(String(o['加價'] || '0').replace(/[^0-9.]/g,'')) || 0 })) .filter(o => o.name); result[key].push({ category:cat, required:String(g['必選']||'').trim()==='TRUE', min:Number(g['最少']||0), max:Number(g['最多']||0), options:opts }); }); return result; } 

// ════════════════════════════════════════════════════════════════ // Users // ════════════════════════════════════════════════════════════════ async function saveUserToSheet(name, userId, src, gid) { try { await authSheet(); const s = doc.sheetsByTitle['Users']; if (!s) return; await s.addRow({ 時間:nowTW(), LINE名稱:name, userId, 來源類型:src, 群組ID:gid||'', 權限:isAdmin(userId)?'admin':'user' }); } catch(e) { console.error('Users write fail:', e.message); } } 

// ════════════════════════════════════════════════════════════════ // Orders CRUD // ════════════════════════════════════════════════════════════════ // ════════════════════════════════════════════════════════════════ // 純文字訂餐模式（阿姨版） // 不需要 LIFF，LINE 裡直接回覆數字即可 // ════════════════════════════════════════════════════════════════ 

// 今日菜單快取（開團時載入，結單時清除） let textMenuCache = []; // [{no, store, item, price}] 

// 讀菜單並編號 async function loadTextMenu() { const menu = await loadMenu(); textMenuCache = menu.map((m, i) => ({ no: i + 1, store: m.store, item: m.item, price: m.price })); return textMenuCache; } 

// 產生菜單訊息 function buildMenuMessage(menuList) { const emojiNum = ['1️⃣','2️⃣','3️⃣','4️⃣','5️⃣','6️⃣','7️⃣','8️⃣','9️⃣','🔟']; let msg = '🍱 今日菜單\n'; msg += '─────────────\n'; menuList.forEach(m => { const em = m.no <= 10 ? emojiNum[m.no - 1] : m.no + '.'; msg += em + ' ' + m.item + '　$' + m.price + '\n'; }); msg += '\n請直接回覆數字點餐：\n'; msg += '例：\n2　　　（點2號）\n2不要菜　（加備註）\n2+3　　（點2號，3份）'; return msg; } 

// 解析文字訂餐："2不要菜""3+2加辣""1" function parseTextOrder(text, menuList) { // 支援格式：數字 [+數量] [備註文字] // e.g. "2", "2不要菜", "3+2", "3+2加辣", "1+3 飯少一點" const m = text.match(/^(\d+)(?:+(\d+))?(.*)$/); if (!m) return null; const no = parseInt(m[1], 10); const qty = Math.min(20, Math.max(1, parseInt(m[2] || '1', 10))); const note = (m[3] || '').trim(); const item = menuList.find(x => x.no === no); if (!item) return null; return { no, item: item.item, store: item.store, price: item.price, qty, note }; } 

// 儲存或覆蓋文字訂單（同一人今日同品項覆蓋，不同品項新增） async function saveTextOrder(name, userId, parsed) { try { await authSheet(); const sheet = doc.sheetsByTitle['Orders']; if (!sheet) return { success: false, reason: 'no_sheet' }; const rows = await sheet.getRows(); const today = todayTW(); 

// 找今日同一人同品項的未刪除訂單 
const existing = rows.find(r => 
 String(r['userId'] || '') === String(userId) && 
 String(r['品項']   || '').trim() === String(parsed.item).trim() && 
 String(r['時間']   || '').startsWith(today) && 
 String(r['狀態']   || '') !== '已刪除' 
); 
 
if (existing) { 
 // 覆蓋：更新數量、備註、總價 
 existing['數量'] = parsed.qty; 
 existing['備註'] = parsed.note; 
 existing['總價'] = parsed.price * parsed.qty; 
 await existing.save(); 
 return { success: true, action: 'updated' }; 
} 
 
// 新增 
await sheet.addRow({ 
 時間:     nowTW(), 
 LINE名稱: name, 
 userId:   String(userId), 
 店家:     parsed.store, 
 品項:     parsed.item, 
 規格:     '', 
 備註:     parsed.note, 
 數量:     parsed.qty, 
 單價:     parsed.price, 
 總價:     parsed.price * parsed.qty, 
 狀態:     '未付款', 
 付款時間: '', 
 付款方式: '', 
 訂單備註: '' 
}); 
return { success: true, action: 'created' }; 
 

} catch(e) { console.error('saveTextOrder fail:', e.message); return { success: false, reason: e.message }; } } 

// 查詢某人今日訂單（給"我的訂單"指令用） async function getMyTextOrders(userId) { try { await authSheet(); const sheet = doc.sheetsByTitle['Orders']; if (!sheet) return []; const rows = await sheet.getRows(); const today = todayTW(); return rows.filter(r => String(r['userId'] || '') === String(userId) && String(r['時間'] || '').startsWith(today) && String(r['狀態'] || '') !== '已刪除' ).map(r => ({ item: String(r['品項'] || ''), note: String(r['備註'] || ''), qty: Number(r['數量'] || 1), price: Number(r['單價'] || 0), total: Number(r['總價'] || 0), status: String(r['狀態'] || '') })); } catch(e) { return []; } } 

// 標記付款（管理員輸入"XXX 已付款"） async function markPaidByName(name) { try { await authSheet(); const sheet = doc.sheetsByTitle['Orders']; if (!sheet) return 0; const rows = await sheet.getRows(); const today = todayTW(); let count = 0; for (const r of rows) { if ( String(r['LINE名稱'] || '').trim() === name.trim() && String(r['時間'] || '').startsWith(today) && String(r['狀態'] || '') === '未付款' ) { r['狀態'] = '已付款'; r['付款時間'] = nowTW(); r['付款方式'] = '現金'; await r.save(); count++; } } return count; } catch(e) { return 0; } } 

// 產生文字統計報表（更詳細版，給查看訂單指令） async function buildDetailedReport() { const orders = await getAllOrdersByDate(todayTW()); const active = orders.filter(o => o.status !== '已刪除'); if (!active.length) return '📋 今日尚無訂單'; 

// 依人員分組 const byPerson = {}; for (const o of active) { const n = o.name || '未知'; if (!byPerson[n]) byPerson[n] = { orders: [], total: 0, paid: false }; byPerson[n].orders.push(o); byPerson[n].total += o.total; if (o.status === '已付款') byPerson[n].paid = true; } 

let msg = '📋 今日訂單明細\n'; msg += '─────────────\n'; for (const name in byPerson) { const p = byPerson[name]; const payIcon = p.orders.every(o => o.status === '已付款') ? '✅' : '💰'; msg += payIcon + ' ' + name + '（$' + p.total + '）\n'; for (const o of p.orders) { msg += ' • ' + o.item; if (o.qty > 1) msg += ' ×' + o.qty; if (o.note) msg += '（' + o.note + '）'; msg += '\n'; } } 

const grand = active.reduce((a, o) => a + o.total, 0); const unpaidNames = Object.keys(byPerson).filter(n => byPerson[n].orders.some(o => o.status === '未付款') ); msg += '─────────────\n'; msg += '💰 總金額：$' + grand + '\n'; if (unpaidNames.length) { msg += '⚠️ 未付款：' + unpaidNames.join('、'); } else { msg += '✅ 全部已付款'; } return msg; } 

// 產生匯出格式（CSV 式文字，方便複製） async function buildExportText() { const orders = await getAllOrdersByDate(todayTW()); const active = orders.filter(o => o.status !== '已刪除'); if (!active.length) return '今日尚無訂單'; let out = '姓名,品項,數量,備註,金額,付款狀態\n'; for (const o of active) { out += [o.name, o.item, o.qty, o.note, o.total, o.status].join(',') + '\n'; } return out; } 

async function saveOrderToSheet(order) { try { await authSheet(); const s = doc.sheetsByTitle['Orders']; if (!s) return { success:false, reason:'no_sheet' }; const rows = await s.getRows(); const today = todayTW(); const dup = rows.find(r => String(r['userId']||'')=== String(order.userId) && String(r['品項'] ||'').trim()===String(order.item||'').trim() && String(r['規格'] ||'').trim()===String(order.spec||'').trim() && String(r['時間'] ||'').startsWith(today) && String(r['狀態'] ||'')!=='已刪除' ); if (dup) return { success:false, reason:'duplicate' }; const qty=Number(order.qty||1), price=Number(order.price||0); await s.addRow({ 時間:nowTW(), LINE名稱:String(order.name||''), userId:String(order.userId||''), 店家:String(order.store||''), 品項:String(order.item||''), 規格:String(order.spec||''), 備註:String(order.note||''), 數量:qty, 單價:price, 總價:price*qty, 狀態:'未付款', 付款時間:'', 付款方式:'', 訂單備註:'' }); return { success:true }; } catch(e) { console.error('Orders write fail:', e.message); return { success:false, reason:e.message }; } } 

async function getOrdersByUser(userId, dateStr) { try { await authSheet(); const s = doc.sheetsByTitle['Orders']; if (!s) return []; const rows = await s.getRows(); const target = dateStr || todayTW(); return rows .filter(r => String(r['userId']||'')===String(userId) && String(r['時間']||'').startsWith(target) && String(r['狀態']||'')!=='已刪除') .map(r => ({ rowIndex:r.rowIndex, store:String(r['店家']||''), item:String(r['品項']||''), spec:String(r['規格']||''), note:String(r['備註']||''), qty:Number(r['數量']||1), price:Number(r['單價']||0), total:Number(r['總價']||0), status:String(r['狀態']||'未付款'), time:String(r['時間']||'') })); } catch(e) { return []; } } 

async function getOrderDates(userId) { try { await authSheet(); const s = doc.sheetsByTitle['Orders']; if (!s) return []; const rows = await s.getRows(); const dates = new Set(); rows.forEach(r => { if (String(r['userId']||'')===userId && String(r['狀態']||'')!=='已刪除') { const d = String(r['時間']||'').split(' ')[0]; if (d) dates.add(d); } }); return […dates].sort().reverse().slice(0, 14); } catch(e) { return []; } } 

async function deleteOrder(userId, rowIndex) { try { await authSheet(); const s = doc.sheetsByTitle['Orders']; if (!s) return false; const rows = await s.getRows(); const t = rows.find(r => Number(r.rowIndex)===Number(rowIndex) && String(r['userId']||'')===String(userId) && String(r['狀態']||'')!=='已刪除'); if (!t) return false; t['狀態']='已刪除'; await t.save(); return true; } catch(e) { return false; } } 

async function updateOrderQty(userId, rowIndex, qty) { try { await authSheet(); const s = doc.sheetsByTitle['Orders']; if (!s) return false; const rows = await s.getRows(); const t = rows.find(r => Number(r.rowIndex) === Number(rowIndex) && String(r['userId'] || '') === String(userId) && String(r['狀態'] || '') === '未付款' ); if (!t) return false; const q = Math.min(20, Math.max(1, Number(qty) || 1)); const price = Number(t['單價'] || 0); t['數量'] = q; t['總價'] = price * q; await t.save(); return true; } catch(e) { console.error('updateOrderQty fail:', e.message); return false; } } 

async function updateOrderNote(userId, rowIndex, note) { try { await authSheet(); const s = doc.sheetsByTitle['Orders']; if (!s) return false; const rows = await s.getRows(); const t = rows.find(r => Number(r.rowIndex)===Number(rowIndex) && String(r['userId']||'')===String(userId) && String(r['狀態']||'')!=='已刪除'); if (!t) return false; t['備註']=note; await t.save(); return true; } catch(e) { return false; } } 

async function getAllOrdersByDate(dateStr) { try { await authSheet(); const s = doc.sheetsByTitle['Orders']; if (!s) return []; const rows = await s.getRows(); const target = dateStr || todayTW(); return rows .filter(r => String(r['時間']||'').startsWith(target)) .map(r => ({ rowIndex:r.rowIndex, name:String(r['LINE名稱']||''), userId:String(r['userId']||''), store:String(r['店家']||''), item:String(r['品項']||''), spec:String(r['規格']||''), note:String(r['備註']||''), qty:Number(r['數量']||1), price:Number(r['單價']||0), total:Number(r['總價']||0), status:String(r['狀態']||'未付款'), payTime:String(r['付款時間']||''), payType:String(r['付款方式']||'') })); } catch(e) { return []; } } 

async function markPaid(rowIndex, payType) { try { await authSheet(); const s = doc.sheetsByTitle['Orders']; if (!s) return false; const rows = await s.getRows(); const t = rows.find(r => Number(r.rowIndex)===Number(rowIndex)); if (!t) return false; t['狀態']='已付款'; t['付款時間']=nowTW(); t['付款方式']=payType||'現金'; await t.save(); return true; } catch(e) { return false; } } 

async function batchMarkPaid(payType) { try { await authSheet(); const s = doc.sheetsByTitle['Orders']; if (!s) return 0; const rows = await s.getRows(); const today = todayTW(); let count = 0; for (const r of rows) { if (String(r['時間']||'').startsWith(today) && String(r['狀態']||'')==='未付款') { r['狀態']='已付款'; r['付款時間']=nowTW(); r['付款方式']=payType||'現金'; await r.save(); count++; } } return count; } catch(e) { return 0; } } 

async function adminDeleteOrder(rowIndex) { try { await authSheet(); const s = doc.sheetsByTitle['Orders']; if (!s) return false; const rows = await s.getRows(); const t = rows.find(r => Number(r.rowIndex)===Number(rowIndex)); if (!t) return false; t['狀態']='已刪除'; await t.save(); return true; } catch(e) { return false; } } 

async function clearTodayOrders() { try { await authSheet(); const s = doc.sheetsByTitle['Orders']; if (!s) return { success:false, count:0 }; const rows = await s.getRows(); const today = todayTW(); let count = 0; for (const r of rows) { if (String(r['時間']||'').startsWith(today) && String(r['狀態']||'')!=='已刪除') { r['狀態']='已刪除'; await r.save(); count++; } } return { success:true, count }; } catch(e) { console.error('clearTodayOrders fail:', e.message); return { success:false, count:0 }; } } 

// ════════════════════════════════════════════════════════════════ // 自動結單 // ════════════════════════════════════════════════════════════════ function scheduleAutoClose(minutes) { if (autoCloseTimer) clearTimeout(autoCloseTimer); const ms = Math.max(1, Number(minutes)) * 60 * 1000; autoCloseAt = new Date(Date.now() + ms).toISOString(); 

// 提前 5 分鐘推播提醒（只在時間 > 5 分鐘時） if (ms > 5 * 60 * 1000) { setTimeout(() => { if (isOpen) pushToGroup('⏰ 距離收單還有 5 分鐘，請把握時間點餐！'); }, ms - 5 * 60 * 1000); } 

autoCloseTimer = setTimeout(async () => { isOpen = false; autoCloseAt = null; autoCloseTimer = null; console.log('[自動結單]', nowTW()); const stat = await buildStatReport().catch(() => ''); pushToGroup('🔴 已自動結單！\n\n' + stat); }, ms); } function cancelAutoClose() { if (autoCloseTimer) clearTimeout(autoCloseTimer); autoCloseTimer = null; autoCloseAt = null; } 

// ── 推播通知（需設定 LINE_GROUP_ID 環境變數）───────────────────── async function pushToGroup(text) { const gid = process.env.LINE_GROUP_ID || ''; if (!gid) return; try { await client.pushMessage(gid, { type: 'text', text }); } catch(e) { console.error('[push] fail:', e.message); } } 

// ════════════════════════════════════════════════════════════════ // LINE 報表 // ════════════════════════════════════════════════════════════════ async function buildStatReport() { const orders = await getAllOrdersByDate(todayTW()); const active = orders.filter(o => o.status !== '已刪除'); if (!active.length) return '📊 今日尚無訂單'; const itemCount={}, userTotal={}; const unpaid = new Set(); for (const o of active) { const k = o.item+(o.spec?'（'+o.spec+'）':''); itemCount[k]=(itemCount[k]||0)+o.qty; const n = o.name||o.userId||'未知'; userTotal[n]=(userTotal[n]||0)+o.total; if (o.status==='未付款') unpaid.add(n); } const grand = Object.values(userTotal).reduce((a,b)=>a+b,0); let t = '📊 今日訂餐統計\n─────────────\n【品項數量】\n'; for (const k in itemCount) t += k+' x'+itemCount[k]+'\n'; t += '\n【個人金額】\n'; for (const n in userTotal) t += n+'：$'+userTotal[n]+'\n'; t += '\n💰 總金額：$'+grand; t += unpaid.size ? '\n\n⚠️ 未付款：'+[…unpaid].join('、') : '\n\n✅ 所有人已付款'; return t; } 

async function buildShopOrder() { const orders = await getAllOrdersByDate(todayTW()); const active = orders.filter(o => o.status !== '已刪除'); if (!active.length) return '今日尚無訂單'; const itemCount={}; for (const o of active) { const k = o.item+(o.spec?'（'+o.spec+'）':''); itemCount[k]=(itemCount[k]||0)+o.qty; } let out='您好，今天訂購如下：\n\n', total=0; for (const k in itemCount) { out+=k+' x'+itemCount[k]+'\n'; total+=itemCount[k]; } const money = active.reduce((a,o)=>a+o.total,0); return out+'\n總數：'+total+'份\n總金額：'+money+'元\n\n麻煩您，謝謝～'; } 

// ════════════════════════════════════════════════════════════════ // Admin middleware // ════════════════════════════════════════════════════════════════ function adminAuth(req, res, next) { const t = req.headers['x-admin-token'] || req.query.token || ''; if (!process.env.ADMIN_TOKEN || t !== process.env.ADMIN_TOKEN) return res.status(401).json({ error:'Unauthorized' }); next(); } 

// ════════════════════════════════════════════════════════════════ // API 路由 // ════════════════════════════════════════════════════════════════ app.get('/', (_q,r) => r.send('LINE 訂餐機器人運作中')); 

app.get('/api/menu', async (_q, res) => { let menu = [], optionData = {}; try { menu = await loadMenu(); } catch(e) { console.error('[/api/menu] loadMenu failed:', e.message); } try { optionData = await loadOptions(); } catch(e) { console.error('[/api/menu] loadOptions failed:', e.message); } res.json({ menu, optionData }); }); 

app.get('/api/status', (_q, res) => res.json({ isOpen, autoCloseAt })); 

// 管理員專用：取得後台連結（token 不暴露給前端） app.get('/api/admin/link', (req, res) => { const uid = req.query.userId || ''; if (!isAdmin(uid)) return res.status(403).json({ error: 'forbidden' }); const token = process.env.ADMIN_TOKEN || ''; const base = process.env.APP_URL || ''; res.json({ url: base + '/admin?token=' + token }); }); 

app.get('/api/my-dates', async (req, res) => { const { userId } = req.query; if (!userId) return res.json([]); res.json(await getOrderDates(userId)); }); 

app.post('/api/order', async (req, res) => { res.json(await saveOrderToSheet(req.body)); }); 

app.get('/api/my-orders', async (req, res) => { const { userId, date } = req.query; if (!userId) return res.json([]); res.json(await getOrdersByUser(userId, date||null)); }); 

app.delete('/api/order/:ri', async (req, res) => { const { userId } = req.body; if (!userId) return res.json({ success:false }); res.json({ success: await deleteOrder(userId, Number(req.params.ri)) }); }); 

app.patch('/api/order/:ri/qty', async (req, res) => { const { userId, qty } = req.body; if (!userId) return res.json({ success: false }); res.json({ success: await updateOrderQty(userId, Number(req.params.ri), qty) }); }); 

app.patch('/api/order/:ri/note', async (req, res) => { const { userId, note } = req.body; if (!userId) return res.json({ success:false }); res.json({ success: await updateOrderNote(userId, Number(req.params.ri), note) }); }); 

// Admin API app.post('/api/admin/open', adminAuth, (_q,res) => { isOpen=true; res.json({ isOpen }); }); app.post('/api/admin/close', adminAuth, (_q,res) => { isOpen=false; cancelAutoClose(); res.json({ isOpen }); }); app.post('/api/admin/clear', adminAuth, (_q,res) => { isOpen=false; cancelAutoClose(); res.json({ ok:true }); }); app.post('/api/admin/auto-close', adminAuth, (req,res) => { if (!isOpen) isOpen=true; scheduleAutoClose(Number(req.body.minutes)||30); res.json({ autoCloseAt }); }); app.post('/api/admin/cancel-auto-close', adminAuth, (_q,res) => { cancelAutoClose(); res.json({ ok:true }); }); app.get('/api/admin/orders', adminAuth, async (req,res) => { res.json(await getAllOrdersByDate(req.query.date || todayTW())); }); app.post('/api/admin/paid', adminAuth, async (req,res) => { res.json({ success: await markPaid(req.body.rowIndex, req.body.payType) }); }); app.post('/api/admin/batch-paid', adminAuth, async (req,res) => { res.json({ success:true, count: await batchMarkPaid(req.body.payType||'現金') }); }); app.post('/api/admin/delete-order', adminAuth, async (req,res) => { res.json({ success: await adminDeleteOrder(req.body.rowIndex) }); }); 

app.post('/api/admin/clear-today', adminAuth, async (_q, res) => { res.json(await clearTodayOrders()); }); 

// ════════════════════════════════════════════════════════════════ // 前台頁面（Array join，零 template literal，零 JSON 內嵌） // ════════════════════════════════════════════════════════════════ app.get('/order', (_q, res) => { const liffId = String(process.env.LIFF_ID || '2010025093-yATK02dc').replace(/[^a-zA-Z0-9-]/g,''); res.setHeader('Content-Type','text/html; charset=utf-8'); res.end(buildOrderPage(liffId)); }); 

function buildOrderPage(liffId) { const L = []; const p = s => L.push(s); 

/* ── HEAD ─────────────────────────────────────────────────── / p(''); p(''); p(''); p(''); p(''); p(''); 

/* ── HEADER ──────────────────────────────────────────────── */ p(' 

'); p('  

'); p('  

🍱 

訂餐小幫手 

'); p(' 未開放'); p('  

'); p('  

正在取得 LINE 使用者資料… 

'); p('  

'); p(' 

'); 

/* ── SEARCH ──────────────────────────────────────────────── */ p(' 

'); 

/* ── TABS ────────────────────────────────────────────────── */ p(' 

'); 

/* ── MENU ────────────────────────────────────────────────── */ p(' 

菜單載入中… 

'); 

/* ── FABs ────────────────────────────────────────────────── */ p(' 

'); p(' ⚙️ 後台'); p(' 📋 我的訂單'); p(' 🛒 購物車0'); p(' 

'); 

/* ── TOAST ───────────────────────────────────────────────── */ p(' 

'); 

/* ── MODAL: 選項 ─────────────────────────────────────────── */ p(' 

'); p('  

'); p(' ✕'); p('  

'); p('  

'); p('  

數量 

'); p('  

'); p(' −'); p(' 1'); p(' ＋'); p('  

'); p('  

'); p(' '); p(' 加入購物車 🛒'); p(' 取消'); p(' 

'); 

/* ── MODAL: 購物車 ───────────────────────────────────────── */ p(' 

'); p('  

'); p(' ✕'); p('  

🛒 我的購物車 

'); p('  

'); p('  

'); p(' 送出訂單'); p(' 繼續點餐'); p(' 

'); 

/* ── MODAL: 我的訂單 ─────────────────────────────────────── */ p(' 

'); p('  

'); p(' ✕'); p('  

📋 我的訂單 

'); p('  

'); p('  

'); p('  

'); p(' 關閉'); p(' 

'); 

/* ── MODAL: 備註編輯 ─────────────────────────────────────── */ p(' 

'); p(' ✕'); p('  

✏️ 修改備註 

'); p(' '); p(' 儲存'); p(' 取消'); p(' 

'); 

/* ════════════════════════════════════════════════════════════ JavaScript — 全部用一般字串，不用 template literal 無任何 JSON 內嵌，資料全部由 fetch 取得 ══════════════════════════════════════════════════════════════ */ p(''); p(''); 

return L.join('\n'); } 

// ════════════════════════════════════════════════════════════════ // 管理後台 // ════════════════════════════════════════════════════════════════ app.get('/admin', (req, res) => { const token = req.query.token || ''; if (!process.env.ADMIN_TOKEN || token !== process.env.ADMIN_TOKEN) return res.status(401).send('Unauthorized. 請附上 ?token=YOUR_ADMIN_TOKEN'); res.setHeader('Content-Type','text/html; charset=utf-8'); res.end(buildAdminPage()); }); 

function buildAdminPage() { const L = []; const p = s => L.push(s); 

/* ── HEAD ────────────────────────────────────────────────── / p(''); p(''); p(''); p(''); 

/* ── HEADER ──────────────────────────────────────────────── */ p(' 

'); p('  

'); p('  

🍱 

'); p('  

訂餐管理後台 

載入中… 

'); p('  

'); p('  

'); p(' 

'); p(' 

'); 

/* ── 開單控制 ──────────────────────────────────────────────── */ p(' 

'); p(' 

開單控制 

'); p(' 

'); p('  

'); p(' 已結單'); p(' ▶ 開單'); p(' ■ 結單'); p(' 清空狀態'); p('  

'); p('  

'); p(' 自動結單'); p(' 分鐘後'); p(' 設定'); p(' 取消'); p(' '); p('  

'); p(' 

'); 

/* ── 今日統計 ──────────────────────────────────────────────── */ p(' 

'); p(' 

今日統計 

'); p(' 

'); p('  

— 

載入中 

'); p(' 

'); p(' 

'); p(' 

'); 

/* ── 訂單列表 ──────────────────────────────────────────────── */ p(' 

'); p(' 

訂單列表 

'); p(' 

'); 

/* toolbar */ p('  

'); p(' '); p(' ✅ 全標已付'); p(' 📋 複製店家單'); p(' ↺ 重整'); p(' 🗑 清空今日訂單'); p('  

'); 

/* filter row */ p('  

'); p(' 日期'); p(' '); p(' 今日'); p(' 分組'); p(' '); p('  

依店家 

'); p('  

依人員 

'); p('  

不分組 

'); p(' '); p('  

'); 

p('  

載入中… 

'); p(' 

'); p(' 

'); /* .wrap */ 

/* ── SCRIPT ──────────────────────────────────────────────── / p(''); 

return L.join('\n'); } 

// ════════════════════════════════════════════════════════════════ // LINE Webhook // ════════════════════════════════════════════════════════════════ app.post('/webhook', async (req, res) => { try { const event = req.body.events?.[0]; if (!event || event.type !== 'message' || event.message.type !== 'text') return res.sendStatus(200); 

let profileName = '未知使用者'; 
try { 
 if (event.source.type === 'group') { 
   const p = await client.getGroupMemberProfile(event.source.groupId, event.source.userId); 
   profileName = p.displayName; 
 } else { 
   const p = await client.getProfile(event.source.userId); 
   profileName = p.displayName; 
 } 
} catch(e) { console.error('取得名稱失敗:', e.message); } 
 
knownUsers[profileName] = event.source.userId; 
saveUserToSheet(profileName, event.source.userId, event.source.type, event.source.groupId||'').catch(()=>{}); 
 
const uid   = event.source.userId; 
const text  = event.message.text.trim(); 
const reply = t => client.replyMessage(event.replyToken, { type:'text', text:t }); 
 
const autoOpen = text.match(/^開單\s+(\d+)$/); 
if (autoOpen) { 
 if (!isAdmin(uid)) { await reply('只有管理員可以開單'); return res.sendStatus(200); } 
 isOpen = true; 
 scheduleAutoClose(Number(autoOpen[1])); 
 await loadTextMenu(); 
 const menuStr = buildMenuMessage(textMenuCache); 
 const autoMsg = '🟢 開始點餐！\n⏰ 將於 '+autoOpen[1]+' 分鐘後自動結單\n\n' + menuStr; 
 await reply(autoMsg); 
 pushToGroup(autoMsg); 
 return res.sendStatus(200); 
} 
if (text === '開單') { 
 if (!isAdmin(uid)) { await reply('只有管理員可以開單'); return res.sendStatus(200); } 
 if (isOpen) { await reply('目前已開單中'); return res.sendStatus(200); } 
 isOpen = true; 
 await loadTextMenu(); 
 const menuStr = buildMenuMessage(textMenuCache); 
 const openMsg = '🟢 開始點餐！\n\n' + menuStr; 
 await reply(openMsg); 
 pushToGroup(openMsg); 
 return res.sendStatus(200); 
} 
if (text === '結單' || text === '收單' || text === '統計') { 
 if (!isAdmin(uid)) { await reply('只有管理員可以結單/統計'); return res.sendStatus(200); } 
 isOpen = false; cancelAutoClose(); 
 textMenuCache = []; 
 const stat = await buildStatReport(); 
 await reply(stat); 
 pushToGroup('🔴 已收單！\n\n' + stat); 
 return res.sendStatus(200); 
} 
if (text === '店家單') { 
 if (!isAdmin(uid)) { await reply('只有管理員可以查看店家單'); return res.sendStatus(200); } 
 await reply(await buildShopOrder()); 
 return res.sendStatus(200); 
} 
if (text === '清空') { 
 if (!isAdmin(uid)) { await reply('只有管理員可以清空'); return res.sendStatus(200); } 
 isOpen = false; cancelAutoClose(); 
 await reply('已清空訂單'); 
 return res.sendStatus(200); 
} 
if (text === '後台') { 
 if (!isAdmin(uid)) { await reply('只有管理員可以查看後台'); return res.sendStatus(200); } 
 await reply('管理後台：'+(process.env.APP_URL||'')+'/admin?token='+(process.env.ADMIN_TOKEN||'')); 
 return res.sendStatus(200); 
} 
if (text === '狀態') { 
 const msg = (isOpen?'🟢 目前開單中':'🔴 目前未開單')+(autoCloseAt?'\n⏰ 自動結單：'+new Date(autoCloseAt).toLocaleTimeString('zh-TW',{timeZone:'Asia/Taipei'}):''); 
 await reply(msg); 
 return res.sendStatus(200); 
} 
 
// ── 管理員指令：查看訂單 ───────────────────────────────────── 
if (text === '查看訂單' || text === '訂單') { 
 if (!isAdmin(uid)) { await reply('只有管理員可以查看'); return res.sendStatus(200); } 
 await reply(await buildDetailedReport()); 
 return res.sendStatus(200); 
} 
 
// ── 管理員指令：匯出 ──────────────────────────────────────── 
if (text === '匯出' || text === '匯出訂單') { 
 if (!isAdmin(uid)) { await reply('只有管理員可以匯出'); return res.sendStatus(200); } 
 await reply(await buildExportText()); 
 return res.sendStatus(200); 
} 
 
// ── 管理員指令：XXX 已付款 ─────────────────────────────────── 
const paidMatch = text.match(/^(.+?)\s*已付款$/); 
if (paidMatch && isAdmin(uid)) { 
 const targetName = paidMatch[1].trim(); 
 const count = await markPaidByName(targetName); 
 if (count > 0) { 
   await reply(targetName + ' 已標記付款（' + count + ' 筆）✅'); 
 } else { 
   await reply('找不到 ' + targetName + ' 的未付款訂單'); 
 } 
 return res.sendStatus(200); 
} 
 
// ── 使用者指令：我的訂單 ──────────────────────────────────── 
if (text === '我的訂單' || text === '我點了什麼') { 
 const myOrders = await getMyTextOrders(uid); 
 if (!myOrders.length) { 
   await reply('你今天還沒有訂單'); 
 } else { 
   let msg = '📋 你的訂單：\n'; 
   let total = 0; 
   for (const o of myOrders) { 
     msg += '• ' + o.item; 
     if (o.qty > 1) msg += ' ×' + o.qty; 
     if (o.note) msg += '（' + o.note + '）'; 
     msg += '　$' + o.total; 
     if (o.status === '已付款') msg += ' ✅'; 
     msg += '\n'; 
     total += o.total; 
   } 
   msg += '合計：$' + total; 
   await reply(msg); 
 } 
 return res.sendStatus(200); 
} 
 
// ── 使用者指令：菜單 ──────────────────────────────────────── 
if (text === '菜單' || text === '今日菜單') { 
 if (!isOpen) { 
   await reply('目前尚未開團，請等待管理員開團 🍱'); 
   return res.sendStatus(200); 
 } 
 const menu = textMenuCache.length ? textMenuCache : await loadTextMenu(); 
 await reply(buildMenuMessage(menu)); 
 return res.sendStatus(200); 
} 
 
// ── 使用者指令：取消訂單 ──────────────────────────────────── 
if (text === '取消' || text === '取消訂單') { 
 if (!isOpen) { await reply('目前未開團'); return res.sendStatus(200); } 
 try { 
   await authSheet(); 
   const s = doc.sheetsByTitle['Orders']; 
   const rows = await s.getRows(); 
   const today = todayTW(); 
   let count = 0; 
   for (const r of rows) { 
     if (String(r['userId']||'')===String(uid) && String(r['時間']||'').startsWith(today) && String(r['狀態']||'')==='未付款') { 
       r['狀態']='已刪除'; await r.save(); count++; 
     } 
   } 
   await reply(count > 0 ? '已取消你的訂單（' + count + ' 筆）' : '找不到可取消的訂單'); 
 } catch(e) { 
   await reply('取消失敗，請聯絡管理員'); 
 } 
 return res.sendStatus(200); 
} 
 
// ── 開團時：純數字訂餐（e.g. "2", "2不要菜", "3+2加辣"）──── 
if (isOpen && textMenuCache.length) { 
 const parsed = parseTextOrder(text, textMenuCache); 
 if (parsed) { 
   const result = await saveTextOrder(profileName, uid, parsed); 
   if (result.success) { 
     const action = result.action === 'updated' ? '已更新' : '已收到'; 
     let msg = action + ' 你的訂單 ✅\n'; 
     msg += parsed.item; 
     if (parsed.qty > 1) msg += ' ×' + parsed.qty; 
     if (parsed.note) msg += '（' + parsed.note + '）'; 
     msg += '\n金額：$' + (parsed.price * parsed.qty); 
     msg += '\n\n輸入"我的訂單"查看，輸入"取消"可取消'; 
     await reply(msg); 
   } else { 
     await reply('訂單失敗，請重試或聯絡管理員'); 
   } 
   return res.sendStatus(200); 
 } 
} 
 
// ── 開團中但看不懂的訊息：給提示 ─────────────────────────── 
if (isOpen && textMenuCache.length) { 
 await reply('請輸入數字點餐，例如"2"或"2不要菜"\n輸入"菜單"可重看菜單'); 
 return res.sendStatus(200); 
} 
 
return res.sendStatus(200); 
 

} catch(err) { console.error('Webhook error:', err); return res.sendStatus(200); } }); 

app.use((err,_q,res,_n) => { console.error('Global error:',err); res.sendStatus(200); }); 

const PORT = process.env.PORT || 3000; app.listen(PORT, () => console.log('Server running on port', PORT)); 

 
