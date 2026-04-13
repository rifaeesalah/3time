const APP_VERSION = '3 TIME Apps Script POS v5.0';
const DEFAULT_SHEET_NAME = '3TIME_DB_V4';
const CACHE_KEY_SETTINGS = 'SYSTEM_SETTINGS_V4';

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.page === 'customer') {
    return HtmlService.createTemplateFromFile('Customer')
      .evaluate()
      .setTitle('3 TIME - Customer Ordering')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (p.page === 'kds') {
    return HtmlService.createTemplateFromFile('KitchenDisplay')
      .evaluate()
      .setTitle('3 TIME KDS')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('3 TIME POS V5')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


function doPost(e) {
  try {
    const body = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
    const action = body.action || '';
    let result = {success:false, message:'unknown action'};
    if (action === 'claimNextPrintJobs') result = claimNextPrintJobs(body.printerName, body.bridgeToken, body.limit);
    else if (action === 'acknowledgePrintedJobs') result = acknowledgePrintedJobs(body);
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({success:false, message:err.message || String(err)})).setMimeType(ContentService.MimeType.JSON);
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function setupSystem() {
  const ss = SpreadsheetApp.create(DEFAULT_SHEET_NAME + ' ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss'));
  const sheets = {};
  [
    'settings','users','tables','menu_items','orders','order_items','service_calls','payments','payment_lines','activity_log',
    'stock_items','stock_movements','print_queue','shifts','shift_events','suppliers','purchase_orders','purchase_order_items','goods_receipts','table_transfers','merge_logs','printer_bridge_logs','refunds','void_logs','supplier_payments','stock_counts','stock_count_lines'
  ].forEach(name => sheets[name] = ss.insertSheet(name));
  const defaultSheet = ss.getSheets().find(s => s.getName() === 'Sheet1');
  if (defaultSheet) ss.deleteSheet(defaultSheet);

  writeSheet_(sheets.settings, ['key','value'], [
    ['restaurant_name', '3 TIME Restaurant & Lounge'],
    ['restaurant_phone', '081-000-0000'],
    ['restaurant_address', 'Yala, Thailand'],
    ['currency', 'THB'],
    ['vat_enabled', 'FALSE'],
    ['vat_rate', '0'],
    ['customer_webapp_url', ''],
    ['auto_queue_print', 'TRUE'],
    ['printer_bridge_enabled', 'FALSE'],
    ['printer_bridge_token', ''],
    ['printer_bridge_default_printer', 'Kitchen-1'],
    ['printer_bridge_poll_seconds', '5'],
    ['created_at', new Date().toISOString()],
    ['version', APP_VERSION],
    ['bar_categories', 'เครื่องดื่ม'],
    ['kitchen_categories', 'โรตี,ข้าวต้ม,ก๋วยเตี๋ยว,ซีฟู้ด'],
  ]);

  writeSheet_(sheets.users, ['id','username','password','name','role','active','pin'], [
    ['U001','admin','1234','ผู้ดูแลระบบ','admin','TRUE','1234'],
    ['U002','manager','1234','ผู้จัดการ','manager','TRUE','1234'],
    ['U003','cashier','1234','แคชเชียร์','cashier','TRUE','1234'],
    ['U004','waiter','1234','พนักงานเสิร์ฟ','waiter','TRUE','1234'],
    ['U005','kitchen','1234','ครัว','kitchen','TRUE','1234'],
  ]);

  const tableRows = [];
  for (let i = 1; i <= 16; i++) {
    tableRows.push([
      'T' + Utilities.formatString('%03d', i),
      String(i),
      i <= 8 ? 'A' : 'B',
      i <= 4 ? 2 : 4,
      'available',
      Utilities.getUuid().replace(/-/g,'').slice(0,12),
      '',
      'TRUE'
    ]);
  }
  writeSheet_(sheets.tables, ['id','table_no','zone','seats','status','token','note','active'], tableRows);

  writeSheet_(sheets.menu_items, ['id','name','category','price','description','image_url','is_available','sort_order','recipe_json'], [
    ['M001','โรตีแกงไก่','โรตี',35,'โรตีแป้งหอม เสิร์ฟพร้อมแกงไก่','', 'TRUE',1, JSON.stringify([{stockId:'S001',qty:1},{stockId:'S002',qty:1}])],
    ['M002','โรตีไข่','โรตี',30,'โรตีไข่นุ่มหอม กินง่าย','', 'TRUE',2, JSON.stringify([{stockId:'S001',qty:1},{stockId:'S003',qty:1}])],
    ['M003','ข้าวต้มปลา','ข้าวต้ม',65,'ข้าวต้มร้อน ๆ น้ำซุปกลมกล่อม','', 'TRUE',3, JSON.stringify([{stockId:'S004',qty:1},{stockId:'S005',qty:1}])],
    ['M004','ก๋วยเตี๋ยวเนื้อ','ก๋วยเตี๋ยว',75,'เส้นนุ่ม น้ำซุปเข้ม','', 'TRUE',4, JSON.stringify([{stockId:'S006',qty:1},{stockId:'S007',qty:1}])],
    ['M005','กุ้งถัง','ซีฟู้ด',289,'กุ้งซอสเข้มข้น สำหรับแชร์','', 'TRUE',5, JSON.stringify([{stockId:'S008',qty:1}])],
    ['M006','ชามะนาว','เครื่องดื่ม',35,'ชาหอมสดชื่น','', 'TRUE',6, JSON.stringify([{stockId:'S009',qty:1}])],
    ['M007','ชาร้อน','เครื่องดื่ม',25,'ชาร้อนแบบดั้งเดิม','', 'TRUE',7, JSON.stringify([{stockId:'S010',qty:1}])]
  ]);

  writeSheet_(sheets.orders, ['id','source','table_id','table_no','customer_name','status','subtotal','vat','total_amount','note','staff_id','staff_name','created_at','updated_at','paid_at'], []);
  writeSheet_(sheets.order_items, ['id','order_id','menu_id','menu_name','qty','unit_price','line_total','note'], []);
  writeSheet_(sheets.service_calls, ['id','table_id','table_no','type','message','status','created_at','handled_at','handled_by'], []);
  writeSheet_(sheets.payments, ['id','order_id','table_id','method','subtotal','vat','total_amount','cash_received','change_amount','staff_id','staff_name','paid_at','shift_id'], []);
  writeSheet_(sheets.payment_lines, ['id','payment_id','order_id','method','amount','cash_received','change_amount','created_at'], []);
  writeSheet_(sheets.activity_log, ['id','action','entity_type','entity_id','detail_json','created_at','actor_id','actor_name'], []);
  writeSheet_(sheets.stock_items, ['id','name','category','unit','qty_on_hand','reorder_point','cost_per_unit','active','note','updated_at'], [
    ['S001','แป้งโรตี','วัตถุดิบ','ก้อน',120,30,6,'TRUE','',new Date().toISOString()],
    ['S002','แกงไก่','วัตถุดิบ','ถ้วย',40,10,12,'TRUE','',new Date().toISOString()],
    ['S003','ไข่ไก่','วัตถุดิบ','ฟอง',120,30,4,'TRUE','',new Date().toISOString()],
    ['S004','ข้าวสวย','วัตถุดิบ','ชาม',80,20,5,'TRUE','',new Date().toISOString()],
    ['S005','ปลา','วัตถุดิบ','เสิร์ฟ',40,10,18,'TRUE','',new Date().toISOString()],
    ['S006','เส้นก๋วยเตี๋ยว','วัตถุดิบ','ชาม',90,20,6,'TRUE','',new Date().toISOString()],
    ['S007','เนื้อ','วัตถุดิบ','เสิร์ฟ',50,10,22,'TRUE','',new Date().toISOString()],
    ['S008','กุ้ง','วัตถุดิบ','เสิร์ฟ',30,8,95,'TRUE','',new Date().toISOString()],
    ['S009','ชามะนาวเบส','เครื่องดื่ม','แก้ว',60,15,8,'TRUE','',new Date().toISOString()],
    ['S010','ชาร้อนเบส','เครื่องดื่ม','แก้ว',60,15,5,'TRUE','',new Date().toISOString()]
  ]);
  writeSheet_(sheets.stock_movements, ['id','stock_id','stock_name','movement_type','qty_delta','balance_after','reference_type','reference_id','note','actor_id','actor_name','created_at'], []);
  writeSheet_(sheets.print_queue, ['id','queue_type','reference_id','payload_json','status','created_at','printed_at','printer_name','note'], []);
  writeSheet_(sheets.shifts, ['id','opened_at','opened_by_id','opened_by_name','opening_cash','status','closed_at','closed_by_id','closed_by_name','closing_cash_counted','expected_cash','cash_diff','note'], []);
  writeSheet_(sheets.shift_events, ['id','shift_id','event_type','reference_id','amount','detail_json','created_at'], []);

  writeSheet_(sheets.suppliers, ['id','name','contact_name','phone','line_id','address','payment_term_days','active','note','updated_at'], [
    ['SUP001','ซัพพลายเออร์กลาง 3 TIME','คุณอาลี','081-000-1111','','ยะลา',7,'TRUE','',new Date().toISOString()],
    ['SUP002','ทะเลสดยะลา','คุณฟารีดา','081-000-2222','','ยะลา',3,'TRUE','วัตถุดิบซีฟู้ด',new Date().toISOString()]
  ]);
  writeSheet_(sheets.purchase_orders, ['id','supplier_id','supplier_name','status','po_date','expected_date','subtotal','tax','total_amount','note','created_by_id','created_by_name','approved_by_id','approved_by_name','received_at'], []);
  writeSheet_(sheets.purchase_order_items, ['id','po_id','stock_id','stock_name','qty','unit_cost','line_total','received_qty','note'], []);
  writeSheet_(sheets.goods_receipts, ['id','po_id','supplier_id','supplier_name','received_at','received_by_id','received_by_name','total_amount','reference_no','note'], []);
  writeSheet_(sheets.table_transfers, ['id','from_table_id','from_table_no','to_table_id','to_table_no','order_ids_json','type','created_at','actor_id','actor_name','note'], []);
  writeSheet_(sheets.merge_logs, ['id','target_table_id','target_table_no','source_table_ids_json','source_table_nos','order_ids_json','created_at','actor_id','actor_name','note'], []);
  writeSheet_(sheets.printer_bridge_logs, ['id','queue_id','printer_name','status','message','created_at'], []);
  writeSheet_(sheets.refunds, ['id','order_id','table_id','refund_type','item_ids_json','amount','reason','status','created_at','actor_id','actor_name','payment_id'], []);
  writeSheet_(sheets.void_logs, ['id','order_id','table_id','item_ids_json','amount','reason','created_at','actor_id','actor_name'], []);
  writeSheet_(sheets.supplier_payments, ['id','supplier_id','supplier_name','po_id','amount','method','reference_no','note','paid_at','actor_id','actor_name'], []);
  writeSheet_(sheets.stock_counts, ['id','status','count_date','note','created_at','created_by_id','created_by_name','applied_at','applied_by_id','applied_by_name'], []);
  writeSheet_(sheets.stock_count_lines, ['id','stock_count_id','stock_id','stock_name','system_qty','counted_qty','diff_qty','note'], []);

  styleSheets_(ss);
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  cacheClear_();
  return {
    success: true,
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl(),
    message: 'สร้างระบบ V4 เรียบร้อยแล้ว'
  };
}

function getBootstrapData() {
  const ss = getDb_();
  return {
    success: true,
    hasSetup: !!ss,
    appVersion: APP_VERSION,
    timezone: Session.getScriptTimeZone(),
    settings: getSystemSettings(),
    demoUsers: getUsers_().map(u => ({username:u.username, role:u.role, name:u.name})),
    currentShift: getCurrentShift(),
  };
}

function login(username, password) {
  username = String(username || '').trim();
  password = String(password || '').trim();
  if (!username || !password) throw new Error('กรุณากรอกชื่อผู้ใช้และรหัสผ่าน');
  const user = getUsers_().find(u => u.username === username && u.password === password && isTrue_(u.active));
  if (!user) throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
  logActivity_('login', 'user', user.id, {username:user.username, role:user.role}, user);
  return {
    success: true,
    user: { id:user.id, username:user.username, name:user.name, role:user.role },
    settings: getSystemSettings(),
    currentShift: getCurrentShift(),
  };
}

function getDashboardSummary() {
  const orders = getOrders({});
  const tables = getTables();
  const todayKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const todayPaid = orders.filter(o => o.status === 'paid' && String(o.paidAt || '').slice(0,10) === todayKey);
  const active = orders.filter(o => ['pending','cooking','ready','served'].includes(o.status));
  const inv = getInventorySummary();
  return {
    todayRevenue: sum_(todayPaid, 'totalAmount'),
    todayOrders: todayPaid.length,
    activeOrders: active.length,
    pendingOrders: orders.filter(o => o.status === 'pending').length,
    lowStockCount: inv.lowStockCount,
    openServiceCalls: getServiceCalls({status:'open'}).length,
    tables,
    currentShift: getCurrentShift(),
  };
}

function getTables() {
  return getRows_('tables').filter(r => isTrue_(r.active)).sort((a,b) => Number(a.table_no) - Number(b.table_no)).map(r => ({
    id: r.id,
    tableNo: r.table_no,
    zone: r.zone,
    seats: Number(r.seats || 0),
    status: r.status || 'available',
    token: r.token,
    note: r.note || '',
  }));
}

function saveTable(payload) {
  payload = payload || {};
  getSheet_('tables').appendRow([
    nextId_('T', getSheet_('tables')),
    String(payload.tableNo || ''),
    String(payload.zone || 'A'),
    Number(payload.seats || 4),
    'available',
    Utilities.getUuid().replace(/-/g,'').slice(0,12),
    String(payload.note || ''),
    'TRUE'
  ]);
  return {success:true};
}

function updateTableStatus(tableId, status) {
  updateRowById_('tables', tableId, row => { row.status = status; return row; });
  return {success:true};
}

function getMenuItems(category) {
  const rows = getRows_('menu_items').filter(r => isTrue_(r.is_available)).map(r => ({
    id: r.id,
    name: r.name,
    category: r.category,
    price: Number(r.price || 0),
    description: r.description || '',
    imageUrl: r.image_url || '',
    sortOrder: Number(r.sort_order || 999),
    recipeJson: r.recipe_json || '[]',
  })).sort((a,b) => a.sortOrder - b.sortOrder || a.name.localeCompare(b.name, 'th'));
  return category ? rows.filter(r => r.category === category) : rows;
}

function saveMenuItem(item) {
  item = item || {};
  const sh = getSheet_('menu_items');
  const id = item.id || nextId_('M', sh);
  const rows = getRows_('menu_items');
  const existing = rows.find(r => r.id === id);
  const recipeJson = typeof item.recipeJson === 'string' ? item.recipeJson : JSON.stringify(item.recipeJson || []);
  if (existing) {
    updateRowById_('menu_items', id, row => {
      row.name = item.name;
      row.category = item.category;
      row.price = item.price;
      row.description = item.description || '';
      row.image_url = item.imageUrl || '';
      row.is_available = item.isAvailable === false ? 'FALSE' : 'TRUE';
      row.sort_order = Number(item.sortOrder || 999);
      row.recipe_json = recipeJson;
      return row;
    });
  } else {
    sh.appendRow([id, item.name || '', item.category || 'อื่น ๆ', Number(item.price || 0), item.description || '', item.imageUrl || '', item.isAvailable === false ? 'FALSE' : 'TRUE', Number(item.sortOrder || 999), recipeJson]);
  }
  return {success:true, id};
}

function createOrder(payload) {
  payload = payload || {};
  if (!payload.tableId) throw new Error('ไม่พบโต๊ะ');
  if (!payload.items || !payload.items.length) throw new Error('ไม่มีรายการสินค้า');
  const user = resolveUser_(payload.staffId);
  const settings = getSystemSettings();
  const subtotal = payload.items.reduce((sum, i) => sum + Number(i.price || i.unitPrice || 0) * Number(i.qty || 0), 0);
  const vat = settings.vatEnabled ? round2_(subtotal * Number(settings.vatRate || 0) / 100) : 0;
  const total = round2_(subtotal + vat);
  const orderId = nextId_('O', getSheet_('orders'));
  const now = new Date().toISOString();

  getSheet_('orders').appendRow([
    orderId,
    payload.source || 'staff',
    payload.tableId,
    payload.tableNo || '',
    payload.customerName || '',
    'pending',
    subtotal,
    vat,
    total,
    payload.note || '',
    user ? user.id : (payload.staffId || ''),
    user ? user.name : '',
    now,
    now,
    ''
  ]);

  const itemSheet = getSheet_('order_items');
  payload.items.forEach(it => {
    itemSheet.appendRow([
      nextId_('OI', itemSheet),
      orderId,
      it.menuId || '',
      it.name || it.menuName || '',
      Number(it.qty || 0),
      Number(it.price || it.unitPrice || 0),
      round2_(Number(it.qty || 0) * Number(it.price || it.unitPrice || 0)),
      it.note || ''
    ]);
  });

  deductStockForOrder_(orderId, payload.items, user);

  updateRowById_('tables', payload.tableId, row => { row.status = 'occupied'; return row; });
  logActivity_('create_order', 'order', orderId, payload, user);
  if (getSystemSettings().autoQueuePrint) queueKitchenPrint(orderId);
  return {success:true, orderId};
}

function getOrders(filter) {
  filter = filter || {};
  const rows = getRows_('orders');
  const items = getRows_('order_items');
  let list = rows.map(r => {
    const its = items.filter(i => i.order_id === r.id).map(i => ({
      id: i.id, menuId: i.menu_id, name: i.menu_name, qty: Number(i.qty || 0), unitPrice: Number(i.unit_price || 0), lineTotal: Number(i.line_total || 0), note: i.note || ''
    }));
    return {
      id: r.id,
      source: r.source || 'staff',
      tableId: r.table_id,
      tableNo: r.table_no,
      customerName: r.customer_name || '',
      status: r.status || 'pending',
      subtotal: Number(r.subtotal || 0),
      vat: Number(r.vat || 0),
      totalAmount: Number(r.total_amount || 0),
      note: r.note || '',
      staffId: r.staff_id || '',
      staffName: r.staff_name || '',
      createdAt: r.created_at,
      updatedAt: r.updated_at,
      paidAt: r.paid_at,
      itemsArray: its,
      items: JSON.stringify(its),
    };
  }).sort((a,b) => String(b.createdAt).localeCompare(String(a.createdAt)));
  if (filter.status && filter.status !== 'all') list = list.filter(o => o.status === filter.status);
  if (filter.tableId) list = list.filter(o => String(o.tableId) === String(filter.tableId));
  if (filter.tableNo) list = list.filter(o => String(o.tableNo) === String(filter.tableNo));
  return list;
}

function updateOrderStatus(orderId, status) {
  const order = updateRowById_('orders', orderId, row => { row.status = status; row.updated_at = new Date().toISOString(); return row; });
  if (['paid','cancelled'].includes(status)) releaseTableIfNoActiveOrders_(order.table_id);
  return {success:true};
}

function processPayment(payload) {
  payload = payload || {};
  if (payload.splitLines && payload.splitLines.length) return processSplitPayment(payload);
  const order = getRows_('orders').find(r => r.id === payload.orderId);
  if (!order) throw new Error('ไม่พบออเดอร์');
  const total = Number(order.total_amount || 0);
  const cashReceived = Number(payload.cashReceived || total);
  const method = payload.method || 'cash';
  const changeAmount = method === 'cash' ? Math.max(0, cashReceived - total) : 0;
  const user = resolveUser_(payload.staffId);
  const paidAt = new Date().toISOString();
  const paymentId = nextId_('P', getSheet_('payments'));
  const currentShift = getCurrentShift();

  updateRowById_('orders', payload.orderId, row => { row.status = 'paid'; row.updated_at = paidAt; row.paid_at = paidAt; return row; });
  getSheet_('payments').appendRow([
    paymentId, payload.orderId, order.table_id, method, Number(order.subtotal || 0), Number(order.vat || 0), total, cashReceived, changeAmount,
    user ? user.id : '', user ? user.name : '', paidAt, currentShift ? currentShift.id : ''
  ]);
  getSheet_('payment_lines').appendRow([nextId_('PL', getSheet_('payment_lines')), paymentId, payload.orderId, method, total, cashReceived, changeAmount, paidAt]);
  if (currentShift) appendShiftEvent_(currentShift.id, 'payment', payload.orderId, total, {method});
  releaseTableIfNoActiveOrders_(order.table_id);
  logActivity_('payment', 'order', payload.orderId, payload, user);
  return {success:true, receiptData: buildReceipt_(payload.orderId), paymentId};
}

function processSplitPayment(payload) {
  payload = payload || {};
  const order = getRows_('orders').find(r => r.id === payload.orderId);
  if (!order) throw new Error('ไม่พบออเดอร์');
  const total = Number(order.total_amount || 0);
  const lines = (payload.splitLines || []).map(l => ({
    method: String(l.method || 'cash'),
    amount: Number(l.amount || 0),
    cashReceived: Number(l.cashReceived || l.amount || 0)
  })).filter(l => l.amount > 0);
  if (!lines.length) throw new Error('กรุณาระบุรายการชำระ');
  const sumLines = round2_(lines.reduce((s,l)=>s+l.amount,0));
  if (Math.abs(sumLines - total) > 0.01) throw new Error('ยอด split bill ต้องรวมเท่ากับยอดบิล');
  const user = resolveUser_(payload.staffId);
  const paidAt = new Date().toISOString();
  const paymentId = nextId_('P', getSheet_('payments'));
  const totalCashReceived = lines.reduce((s,l)=>s+l.cashReceived,0);
  const totalChange = round2_(lines.reduce((s,l)=>s + (l.method === 'cash' ? Math.max(0, l.cashReceived - l.amount) : 0),0));
  const currentShift = getCurrentShift();

  updateRowById_('orders', payload.orderId, row => { row.status = 'paid'; row.updated_at = paidAt; row.paid_at = paidAt; return row; });
  getSheet_('payments').appendRow([
    paymentId, payload.orderId, order.table_id, 'split', Number(order.subtotal || 0), Number(order.vat || 0), total, totalCashReceived, totalChange,
    user ? user.id : '', user ? user.name : '', paidAt, currentShift ? currentShift.id : ''
  ]);
  const linesSheet = getSheet_('payment_lines');
  lines.forEach(line => linesSheet.appendRow([
    nextId_('PL', linesSheet), paymentId, payload.orderId, line.method, line.amount, line.cashReceived,
    line.method === 'cash' ? Math.max(0, line.cashReceived - line.amount) : 0, paidAt
  ]));
  if (currentShift) appendShiftEvent_(currentShift.id, 'payment_split', payload.orderId, total, {lines});
  releaseTableIfNoActiveOrders_(order.table_id);
  logActivity_('payment_split', 'order', payload.orderId, {lines}, user);
  return {success:true, receiptData: buildReceipt_(payload.orderId), paymentId};
}

function getPublicOrderingContext(tableId, token) {
  const table = getRows_('tables').find(r => r.id === tableId && r.token === token && isTrue_(r.active));
  if (!table) throw new Error('ลิงก์โต๊ะไม่ถูกต้อง');
  return {
    success: true,
    table: { id: table.id, tableNo: table.table_no, zone: table.zone, seats: Number(table.seats || 0) },
    settings: getSystemSettings(),
    menuItems: getMenuItems(),
    activeOrders: getActiveOrdersForTable_(table.id),
  };
}

function placePublicOrder(payload) {
  payload = payload || {};
  const table = getRows_('tables').find(r => r.id === payload.tableId && r.token === payload.token && isTrue_(r.active));
  if (!table) throw new Error('ลิงก์โต๊ะไม่ถูกต้อง');
  return createOrder({
    source: 'customer', tableId: table.id, tableNo: table.table_no, customerName: payload.customerName || '', note: payload.note || '', items: payload.items || [], staffId: ''
  });
}

function getCustomerActiveOrders(tableId, token) {
  const table = getRows_('tables').find(r => r.id === tableId && r.token === token && isTrue_(r.active));
  if (!table) throw new Error('ลิงก์โต๊ะไม่ถูกต้อง');
  return {success:true, orders:getActiveOrdersForTable_(table.id)};
}

function createServiceCall(payload) {
  payload = payload || {};
  const table = getRows_('tables').find(r => r.id === payload.tableId && r.token === payload.token && isTrue_(r.active));
  if (!table) throw new Error('ลิงก์โต๊ะไม่ถูกต้อง');
  getSheet_('service_calls').appendRow([
    nextId_('SC', getSheet_('service_calls')), table.id, table.table_no, payload.type || 'waiter', payload.message || '', 'open', new Date().toISOString(), '', ''
  ]);
  return {success:true};
}

function getServiceCalls(filter) {
  filter = filter || {};
  let rows = getRows_('service_calls').map(r => ({
    id: r.id, tableId: r.table_id, tableNo: r.table_no, type: r.type, message: r.message || '', status: r.status || 'open', createdAt: r.created_at, handledAt: r.handled_at, handledBy: r.handled_by || ''
  })).sort((a,b) => String(b.createdAt).localeCompare(String(a.createdAt)));
  if (filter.status) rows = rows.filter(r => r.status === filter.status);
  return rows;
}

function handleServiceCall(callId, actorId) {
  const user = resolveUser_(actorId);
  updateRowById_('service_calls', callId, row => { row.status = 'handled'; row.handled_at = new Date().toISOString(); row.handled_by = user ? user.name : ''; return row; });
  return {success:true};
}

function getCustomerOrderingLinks() {
  const appUrl = ScriptApp.getService().getUrl() || getSystemSettings().customerWebappUrl || '';
  const links = getTables().map(t => ({
    tableId: t.id,
    tableNo: t.tableNo,
    zone: t.zone,
    url: appUrl ? `${appUrl}?page=customer&tableId=${encodeURIComponent(t.id)}&token=${encodeURIComponent(t.token)}` : ''
  }));
  return {success:true, appUrl, links};
}

function getInventorySummary() {
  const items = getRows_('stock_items').filter(r => isTrue_(r.active)).map(r => ({
    id: r.id,
    name: r.name,
    category: r.category,
    unit: r.unit,
    qtyOnHand: Number(r.qty_on_hand || 0),
    reorderPoint: Number(r.reorder_point || 0),
    costPerUnit: Number(r.cost_per_unit || 0),
    note: r.note || '',
    updatedAt: r.updated_at || ''
  }));
  return {
    items,
    lowStockCount: items.filter(i => i.qtyOnHand <= i.reorderPoint).length,
    totalStockValue: round2_(items.reduce((s,i)=>s + i.qtyOnHand * i.costPerUnit,0))
  };
}

function saveStockItem(payload) {
  payload = payload || {};
  const sh = getSheet_('stock_items');
  const id = payload.id || nextId_('S', sh);
  const rows = getRows_('stock_items');
  const exists = rows.find(r => r.id === id);
  if (exists) {
    updateRowById_('stock_items', id, row => {
      row.name = payload.name;
      row.category = payload.category || 'อื่น ๆ';
      row.unit = payload.unit || 'หน่วย';
      row.reorder_point = Number(payload.reorderPoint || 0);
      row.cost_per_unit = Number(payload.costPerUnit || 0);
      row.note = payload.note || '';
      row.active = payload.active === false ? 'FALSE' : 'TRUE';
      row.updated_at = new Date().toISOString();
      return row;
    });
  } else {
    sh.appendRow([id, payload.name || '', payload.category || 'อื่น ๆ', payload.unit || 'หน่วย', Number(payload.qtyOnHand || 0), Number(payload.reorderPoint || 0), Number(payload.costPerUnit || 0), payload.active === false ? 'FALSE' : 'TRUE', payload.note || '', new Date().toISOString()]);
  }
  return {success:true, id};
}

function adjustStock(payload) {
  payload = payload || {};
  const qtyDelta = Number(payload.qtyDelta || 0);
  if (!payload.stockId) throw new Error('ไม่พบ stock item');
  if (!qtyDelta) throw new Error('กรุณาระบุจำนวนที่ปรับ');
  const user = resolveUser_(payload.actorId);
  let newBalance = 0;
  let stockName = '';
  updateRowById_('stock_items', payload.stockId, row => {
    const current = Number(row.qty_on_hand || 0);
    const next = round2_(current + qtyDelta);
    if (next < 0) throw new Error('สต๊อกคงเหลือไม่พอ');
    row.qty_on_hand = next;
    row.updated_at = new Date().toISOString();
    stockName = row.name;
    newBalance = next;
    return row;
  });
  getSheet_('stock_movements').appendRow([
    nextId_('SM', getSheet_('stock_movements')), payload.stockId, stockName, payload.movementType || (qtyDelta > 0 ? 'in' : 'out'), qtyDelta, newBalance,
    payload.referenceType || 'manual', payload.referenceId || '', payload.note || '', user ? user.id : '', user ? user.name : '', new Date().toISOString()
  ]);
  return {success:true, balance:newBalance};
}

function getStockMovements(limit) {
  const rows = getRows_('stock_movements').map(r => ({
    id:r.id, stockId:r.stock_id, stockName:r.stock_name, movementType:r.movement_type, qtyDelta:Number(r.qty_delta || 0), balanceAfter:Number(r.balance_after || 0), note:r.note || '', createdAt:r.created_at
  })).sort((a,b)=>String(b.createdAt).localeCompare(String(a.createdAt)));
  return rows.slice(0, Number(limit || 100));
}

function getPrintQueue(status) {
  let rows = getRows_('print_queue').map(r => ({
    id:r.id, queueType:r.queue_type, referenceId:r.reference_id, payloadJson:r.payload_json || '{}', status:r.status || 'queued', createdAt:r.created_at, printedAt:r.printed_at || '', printerName:r.printer_name || '', note:r.note || ''
  })).sort((a,b)=>String(b.createdAt).localeCompare(String(a.createdAt)));
  if (status && status !== 'all') rows = rows.filter(r => r.status === status);
  return rows;
}

function queueKitchenPrint(orderId) {
  const order = getOrders({}).find(o => o.id === orderId);
  if (!order) throw new Error('ไม่พบออเดอร์');
  const qId = nextId_('Q', getSheet_('print_queue'));
  getSheet_('print_queue').appendRow([
    qId, 'kitchen_ticket', orderId, JSON.stringify({order}), 'queued', new Date().toISOString(), '', '', 'สร้างคิวพิมพ์ครัวอัตโนมัติ'
  ]);
  return {success:true, queueId:qId};
}

function markPrintQueueDone(queueId, printerName) {
  updateRowById_('print_queue', queueId, row => { row.status = 'printed'; row.printed_at = new Date().toISOString(); row.printer_name = printerName || 'manual'; return row; });
  return {success:true};
}

function getCurrentShift() {
  const row = getRows_('shifts').filter(r => r.status === 'open').sort((a,b)=>String(b.opened_at).localeCompare(String(a.opened_at)))[0];
  return row ? mapShift_(row) : null;
}

function openShift(payload) {
  payload = payload || {};
  if (getCurrentShift()) throw new Error('มีรอบกะที่เปิดอยู่แล้ว');
  const user = resolveUser_(payload.actorId);
  const shiftId = nextId_('SH', getSheet_('shifts'));
  const now = new Date().toISOString();
  getSheet_('shifts').appendRow([
    shiftId, now, user ? user.id : '', user ? user.name : '', Number(payload.openingCash || 0), 'open', '', '', '', '', '', '', payload.note || ''
  ]);
  appendShiftEvent_(shiftId, 'open_shift', shiftId, Number(payload.openingCash || 0), {note:payload.note || ''});
  return {success:true, shift:getCurrentShift()};
}

function closeShift(payload) {
  payload = payload || {};
  const current = getCurrentShift();
  if (!current) throw new Error('ไม่มีกะที่เปิดอยู่');
  const user = resolveUser_(payload.actorId);
  const countedCash = Number(payload.countedCash || 0);
  const expectedCash = getShiftExpectedCash_(current.id);
  const diff = round2_(countedCash - expectedCash);
  updateRowById_('shifts', current.id, row => {
    row.status = 'closed';
    row.closed_at = new Date().toISOString();
    row.closed_by_id = user ? user.id : '';
    row.closed_by_name = user ? user.name : '';
    row.closing_cash_counted = countedCash;
    row.expected_cash = expectedCash;
    row.cash_diff = diff;
    row.note = payload.note || row.note || '';
    return row;
  });
  appendShiftEvent_(current.id, 'close_shift', current.id, countedCash, {expectedCash, diff, note:payload.note || ''});
  return {success:true, shiftSummary:getShiftSummary(current.id)};
}

function getShiftSummary(shiftId) {
  const shifts = getRows_('shifts');
  const row = shiftId ? shifts.find(s => s.id === shiftId) : shifts.sort((a,b)=>String(b.opened_at).localeCompare(String(a.opened_at)))[0];
  if (!row) return null;
  const shift = mapShift_(row);
  const payments = getRows_('payments').filter(p => p.shift_id === shift.id);
  const lines = getRows_('payment_lines').filter(pl => payments.some(p => p.id === pl.payment_id));
  const methodTotals = {};
  lines.forEach(l => methodTotals[l.method] = round2_((methodTotals[l.method] || 0) + Number(l.amount || 0)));
  return {
    shift,
    totalSales: round2_(payments.reduce((s,p)=>s+Number(p.total_amount || 0),0)),
    totalPayments: payments.length,
    expectedCash: getShiftExpectedCash_(shift.id),
    paymentMethods: methodTotals,
    payments: payments.map(p => ({id:p.id, orderId:p.order_id, method:p.method, totalAmount:Number(p.total_amount || 0), paidAt:p.paid_at}))
  };
}

function getAllShifts(limit) {
  return getRows_('shifts').map(mapShift_).sort((a,b)=>String(b.openedAt).localeCompare(String(a.openedAt))).slice(0, Number(limit || 30));
}



function mergeTableBills(payload) {
  payload = payload || {};
  const sourceTableIds = (payload.sourceTableIds || []).map(String).filter(Boolean);
  const targetTableId = String(payload.targetTableId || '');
  if (!sourceTableIds.length || !targetTableId) throw new Error('กรุณาระบุโต๊ะต้นทางและโต๊ะปลายทาง');
  if (sourceTableIds.includes(targetTableId)) throw new Error('โต๊ะปลายทางห้ามซ้ำกับโต๊ะต้นทาง');
  const tables = getRows_('tables');
  const target = tables.find(t => String(t.id) === targetTableId && isTrue_(t.active));
  if (!target) throw new Error('ไม่พบโต๊ะปลายทาง');
  const targetOrdersBefore = getActiveOrdersForTable_(targetTableId);
  const movedOrderIds = [];
  const sourceNos = [];
  sourceTableIds.forEach(tableId => {
    const table = tables.find(t => String(t.id) === tableId && isTrue_(t.active));
    if (!table) return;
    const orders = getActiveOrdersForTable_(tableId);
    if (!orders.length) return;
    sourceNos.push(table.table_no);
    orders.forEach(order => {
      updateRowById_('orders', order.id, row => {
        row.table_id = target.id;
        row.table_no = target.table_no;
        row.updated_at = new Date().toISOString();
        row.note = [row.note || '', '[MERGED FROM TABLE ' + table.table_no + ']'].filter(Boolean).join(' ');
        return row;
      });
      movedOrderIds.push(order.id);
    });
    updateRowById_('tables', table.id, row => { row.status = 'available'; return row; });
  });
  if (!movedOrderIds.length) throw new Error('ไม่พบออเดอร์ค้างของโต๊ะที่เลือก');
  updateRowById_('tables', target.id, row => { row.status = 'occupied'; return row; });
  const user = resolveUser_(payload.actorId);
  getSheet_('merge_logs').appendRow([
    nextId_('MG', getSheet_('merge_logs')), target.id, target.table_no, JSON.stringify(sourceTableIds), sourceNos.join(', '), JSON.stringify(movedOrderIds),
    new Date().toISOString(), user ? user.id : '', user ? user.name : '', payload.note || ''
  ]);
  logActivity_('merge_bill', 'table', target.id, {sourceTableIds, targetTableId, movedOrderIds, targetOrdersBefore:targetOrdersBefore.map(o=>o.id)}, user);
  return {success:true, movedOrderIds, targetTableId};
}

function transferTable(payload) {
  payload = payload || {};
  const fromTableId = String(payload.fromTableId || '');
  const toTableId = String(payload.toTableId || '');
  if (!fromTableId || !toTableId) throw new Error('กรุณาระบุโต๊ะต้นทางและปลายทาง');
  if (fromTableId === toTableId) throw new Error('โต๊ะต้นทางและปลายทางต้องไม่ซ้ำกัน');
  const tables = getRows_('tables');
  const fromTable = tables.find(t => String(t.id) === fromTableId && isTrue_(t.active));
  const toTable = tables.find(t => String(t.id) === toTableId && isTrue_(t.active));
  if (!fromTable || !toTable) throw new Error('ไม่พบโต๊ะ');
  const targetActive = getActiveOrdersForTable_(toTableId);
  if (targetActive.length) throw new Error('โต๊ะปลายทางมีออเดอร์ค้างอยู่ ใช้รวมบิลแทน');
  const orders = getActiveOrdersForTable_(fromTableId);
  if (!orders.length) throw new Error('โต๊ะต้นทางไม่มีออเดอร์ค้าง');
  const movedOrderIds = [];
  orders.forEach(order => {
    updateRowById_('orders', order.id, row => {
      row.table_id = toTable.id;
      row.table_no = toTable.table_no;
      row.updated_at = new Date().toISOString();
      row.note = [row.note || '', '[TRANSFERRED FROM TABLE ' + fromTable.table_no + ']'].filter(Boolean).join(' ');
      return row;
    });
    movedOrderIds.push(order.id);
  });
  updateRowById_('tables', fromTable.id, row => { row.status = 'available'; return row; });
  updateRowById_('tables', toTable.id, row => { row.status = 'occupied'; return row; });
  const user = resolveUser_(payload.actorId);
  getSheet_('table_transfers').appendRow([
    nextId_('TT', getSheet_('table_transfers')), fromTable.id, fromTable.table_no, toTable.id, toTable.table_no, JSON.stringify(movedOrderIds),
    'transfer', new Date().toISOString(), user ? user.id : '', user ? user.name : '', payload.note || ''
  ]);
  logActivity_('transfer_table', 'table', fromTable.id, {fromTableId, toTableId, movedOrderIds}, user);
  return {success:true, movedOrderIds};
}

function getSuppliers() {
  return getRows_('suppliers').filter(r => isTrue_(r.active)).map(r => ({
    id:r.id, name:r.name, contactName:r.contact_name || '', phone:r.phone || '', lineId:r.line_id || '', address:r.address || '',
    paymentTermDays:Number(r.payment_term_days || 0), note:r.note || '', updatedAt:r.updated_at || ''
  })).sort((a,b)=>a.name.localeCompare(b.name,'th'));
}

function saveSupplier(payload) {
  payload = payload || {};
  const sh = getSheet_('suppliers');
  const id = payload.id || nextId_('SUP', sh);
  const rows = getRows_('suppliers');
  const exists = rows.find(r => r.id === id);
  if (exists) {
    updateRowById_('suppliers', id, row => {
      row.name = payload.name || '';
      row.contact_name = payload.contactName || '';
      row.phone = payload.phone || '';
      row.line_id = payload.lineId || '';
      row.address = payload.address || '';
      row.payment_term_days = Number(payload.paymentTermDays || 0);
      row.note = payload.note || '';
      row.active = payload.active === false ? 'FALSE' : 'TRUE';
      row.updated_at = new Date().toISOString();
      return row;
    });
  } else {
    sh.appendRow([id, payload.name || '', payload.contactName || '', payload.phone || '', payload.lineId || '', payload.address || '', Number(payload.paymentTermDays || 0), payload.active === false ? 'FALSE' : 'TRUE', payload.note || '', new Date().toISOString()]);
  }
  return {success:true, id};
}

function createPurchaseOrder(payload) {
  payload = payload || {};
  if (!payload.supplierId) throw new Error('กรุณาเลือกซัพพลายเออร์');
  const items = (payload.items || []).map(i => ({
    stockId:String(i.stockId || ''),
    qty:Number(i.qty || 0),
    unitCost:Number(i.unitCost || 0),
    note:String(i.note || '')
  })).filter(i => i.stockId && i.qty > 0);
  if (!items.length) throw new Error('กรุณาระบุรายการสั่งซื้อ');
  const supplier = getRows_('suppliers').find(r => r.id === payload.supplierId);
  if (!supplier) throw new Error('ไม่พบซัพพลายเออร์');
  const stockMap = {};
  getRows_('stock_items').forEach(s => stockMap[s.id] = s);
  const poId = nextId_('PO', getSheet_('purchase_orders'));
  const now = new Date().toISOString();
  const subtotal = round2_(items.reduce((s,i)=>s + i.qty*i.unitCost,0));
  const tax = 0;
  const total = round2_(subtotal + tax);
  const user = resolveUser_(payload.actorId);
  getSheet_('purchase_orders').appendRow([
    poId, supplier.id, supplier.name, payload.status || 'ordered', String(payload.poDate || now.slice(0,10)), String(payload.expectedDate || ''), subtotal, tax, total,
    payload.note || '', user ? user.id : '', user ? user.name : '', '', '', ''
  ]);
  const itemSh = getSheet_('purchase_order_items');
  items.forEach(i => {
    const stock = stockMap[i.stockId];
    itemSh.appendRow([nextId_('POI', itemSh), poId, i.stockId, stock ? stock.name : i.stockId, i.qty, i.unitCost, round2_(i.qty*i.unitCost), 0, i.note || '']);
  });
  logActivity_('create_po', 'purchase_order', poId, {supplierId:payload.supplierId, items}, user);
  return {success:true, poId};
}

function getPurchaseOrders(status) {
  const items = getRows_('purchase_order_items');
  let rows = getRows_('purchase_orders').map(r => {
    const poItems = items.filter(i => i.po_id === r.id).map(i => ({
      id:i.id, stockId:i.stock_id, stockName:i.stock_name, qty:Number(i.qty||0), unitCost:Number(i.unit_cost||0), lineTotal:Number(i.line_total||0), receivedQty:Number(i.received_qty||0), note:i.note||''
    }));
    return {
      id:r.id, supplierId:r.supplier_id, supplierName:r.supplier_name, status:r.status || 'ordered', poDate:r.po_date || '', expectedDate:r.expected_date || '',
      subtotal:Number(r.subtotal || 0), tax:Number(r.tax || 0), totalAmount:Number(r.total_amount || 0), note:r.note || '',
      createdByName:r.created_by_name || '', receivedAt:r.received_at || '', items:poItems
    };
  }).sort((a,b)=>String(b.poDate).localeCompare(String(a.poDate)) || b.id.localeCompare(a.id));
  if (status && status !== 'all') rows = rows.filter(r => r.status === status);
  return rows;
}

function receivePurchaseOrder(payload) {
  payload = payload || {};
  const poId = String(payload.poId || '');
  if (!poId) throw new Error('ไม่พบ PO');
  const po = getRows_('purchase_orders').find(r => r.id === poId);
  if (!po) throw new Error('ไม่พบ PO');
  const receiptItems = (payload.items || []).map(i => ({
    stockId:String(i.stockId || ''),
    receivedQty:Number(i.receivedQty || 0),
    unitCost:Number(i.unitCost || 0)
  })).filter(i => i.stockId && i.receivedQty > 0);
  if (!receiptItems.length) throw new Error('กรุณาระบุจำนวนรับเข้า');
  const user = resolveUser_(payload.actorId);
  let totalReceived = 0;
  const itemRows = getRows_('purchase_order_items').filter(r => r.po_id === poId);
  receiptItems.forEach(it => {
    const poItem = itemRows.find(r => r.stock_id === it.stockId);
    if (!poItem) return;
    const nextReceived = Number(poItem.received_qty || 0) + it.receivedQty;
    updateRowById_('purchase_order_items', poItem.id, row => { row.received_qty = nextReceived; return row; });
    adjustStock({
      stockId: it.stockId, qtyDelta: it.receivedQty, movementType: 'po_receive', referenceType: 'purchase_order',
      referenceId: poId, note: 'รับเข้าจาก PO ' + poId, actorId: user ? user.id : ''
    });
    totalReceived += round2_(it.receivedQty * (it.unitCost || Number(poItem.unit_cost || 0)));
  });
  const allItems = getRows_('purchase_order_items').filter(r => r.po_id === poId);
  const fullyReceived = allItems.every(r => Number(r.received_qty || 0) >= Number(r.qty || 0));
  updateRowById_('purchase_orders', poId, row => {
    row.status = fullyReceived ? 'received' : 'partial_received';
    row.received_at = new Date().toISOString();
    return row;
  });
  const supplier = getRows_('suppliers').find(r => r.id === po.supplier_id);
  getSheet_('goods_receipts').appendRow([
    nextId_('GR', getSheet_('goods_receipts')), poId, po.supplier_id, supplier ? supplier.name : po.supplier_name, new Date().toISOString(),
    user ? user.id : '', user ? user.name : '', totalReceived, payload.referenceNo || '', payload.note || ''
  ]);
  logActivity_('receive_po', 'purchase_order', poId, {receiptItems}, user);
  return {success:true, poId, status: fullyReceived ? 'received' : 'partial_received'};
}

function getGoodsReceipts(limit) {
  return getRows_('goods_receipts').map(r => ({
    id:r.id, poId:r.po_id, supplierName:r.supplier_name || '', receivedAt:r.received_at, receivedByName:r.received_by_name || '', totalAmount:Number(r.total_amount || 0), referenceNo:r.reference_no || '', note:r.note || ''
  })).sort((a,b)=>String(b.receivedAt).localeCompare(String(a.receivedAt))).slice(0, Number(limit || 50));
}

function savePrinterBridgeConfig(payload) {
  payload = payload || {};
  const current = getSystemSettings();
  return saveSystemSettings({
    restaurantName: current.restaurantName,
    restaurantPhone: current.restaurantPhone,
    restaurantAddress: current.restaurantAddress,
    currency: current.currency || 'THB',
    vatEnabled: current.vatEnabled,
    vatRate: current.vatRate,
    customerWebappUrl: current.customerWebappUrl,
    autoQueuePrint: current.autoQueuePrint,
    printerBridgeEnabled: !!payload.printerBridgeEnabled,
    printerBridgeToken: payload.printerBridgeToken || current.printerBridgeToken || '',
    printerBridgeDefaultPrinter: payload.printerBridgeDefaultPrinter || current.printerBridgeDefaultPrinter || 'Kitchen-1',
    printerBridgePollSeconds: Number(payload.printerBridgePollSeconds || current.printerBridgePollSeconds || 5)
  });
}

function claimNextPrintJobs(printerName, bridgeToken, limit) {
  const settings = getSystemSettings();
  if (!settings.printerBridgeEnabled) throw new Error('ยังไม่เปิดใช้งาน printer bridge');
  if (!settings.printerBridgeToken || bridgeToken !== settings.printerBridgeToken) throw new Error('Bridge token ไม่ถูกต้อง');
  const rows = getRows_('print_queue').filter(r => (r.status || 'queued') === 'queued').slice(0, Number(limit || 5));
  const jobs = [];
  rows.forEach(r => {
    updateRowById_('print_queue', r.id, row => {
      row.status = 'claimed';
      row.printer_name = printerName || settings.printerBridgeDefaultPrinter || 'Kitchen-1';
      row.note = [row.note || '', 'claimed by bridge'].filter(Boolean).join(' | ');
      return row;
    });
    jobs.push({
      queueId:r.id,
      queueType:r.queue_type,
      referenceId:r.reference_id,
      payload: JSON.parse(r.payload_json || '{}'),
      printerName: printerName || settings.printerBridgeDefaultPrinter || 'Kitchen-1'
    });
  });
  return {success:true, jobs};
}

function acknowledgePrintedJobs(payload) {
  payload = payload || {};
  const settings = getSystemSettings();
  if (!settings.printerBridgeToken || payload.bridgeToken !== settings.printerBridgeToken) throw new Error('Bridge token ไม่ถูกต้อง');
  const items = payload.items || [];
  items.forEach(item => {
    updateRowById_('print_queue', item.queueId, row => {
      row.status = item.status === 'error' ? 'error' : 'printed';
      row.printed_at = new Date().toISOString();
      row.printer_name = item.printerName || row.printer_name || settings.printerBridgeDefaultPrinter || 'Kitchen-1';
      row.note = [row.note || '', item.message || ''].filter(Boolean).join(' | ');
      return row;
    });
    getSheet_('printer_bridge_logs').appendRow([nextId_('PBL', getSheet_('printer_bridge_logs')), item.queueId, item.printerName || '', item.status || 'printed', item.message || '', new Date().toISOString()]);
  });
  return {success:true};
}

function getCloseShiftPrintableData(shiftId) {
  const summary = getShiftSummary(shiftId);
  if (!summary) throw new Error('ไม่พบข้อมูลกะ');
  const lowStock = getInventorySummary().items.filter(i => i.qtyOnHand <= i.reorderPoint);
  const receipts = getGoodsReceipts(20).filter(r => !shiftId || String(r.receivedAt || '').slice(0,10) >= String(summary.shift.openedAt || '').slice(0,10));
  return {
    success:true,
    summary:summary,
    lowStock:lowStock,
    recentReceipts:receipts,
    printableHtml: buildShiftPrintableHtml_(summary, lowStock, receipts)
  };
}

function getSystemSettings() {
  const cached = CacheService.getScriptCache().get(CACHE_KEY_SETTINGS);
  if (cached) return JSON.parse(cached);
  const rows = getRows_('settings');
  const obj = {};
  rows.forEach(r => obj[r.key] = r.value);
  const settings = {
    restaurantName: obj.restaurant_name || '3 TIME Restaurant & Lounge',
    restaurantPhone: obj.restaurant_phone || '',
    restaurantAddress: obj.restaurant_address || '',
    currency: obj.currency || 'THB',
    vatEnabled: isTrue_(obj.vat_enabled),
    vatRate: Number(obj.vat_rate || 0),
    customerWebappUrl: obj.customer_webapp_url || '',
    autoQueuePrint: isTrue_(obj.auto_queue_print),
    printerBridgeEnabled: isTrue_(obj.printer_bridge_enabled),
    printerBridgeToken: obj.printer_bridge_token || '',
    printerBridgeDefaultPrinter: obj.printer_bridge_default_printer || 'Kitchen-1',
    printerBridgePollSeconds: Number(obj.printer_bridge_poll_seconds || 5),
    version: obj.version || APP_VERSION,
  };
  CacheService.getScriptCache().put(CACHE_KEY_SETTINGS, JSON.stringify(settings), 300);
  return settings;
}

function saveSystemSettings(settings) {
  settings = settings || {};
  const map = {
    restaurant_name: settings.restaurantName,
    restaurant_phone: settings.restaurantPhone,
    restaurant_address: settings.restaurantAddress,
    currency: settings.currency || 'THB',
    vat_enabled: settings.vatEnabled ? 'TRUE' : 'FALSE',
    vat_rate: String(settings.vatRate || 0),
    customer_webapp_url: settings.customerWebappUrl || '',
    auto_queue_print: settings.autoQueuePrint ? 'TRUE' : 'FALSE',
    printer_bridge_enabled: settings.printerBridgeEnabled ? 'TRUE' : 'FALSE',
    printer_bridge_token: settings.printerBridgeToken || '',
    printer_bridge_default_printer: settings.printerBridgeDefaultPrinter || 'Kitchen-1',
    printer_bridge_poll_seconds: String(settings.printerBridgePollSeconds || 5),
    version: APP_VERSION,
  };
  const rows = getRows_('settings');
  const sh = getSheet_('settings');
  const headers = getHeaders_(sh);
  Object.keys(map).forEach(key => {
    const idx = rows.findIndex(r => r.key === key);
    if (idx >= 0) sh.getRange(idx + 2, headers.indexOf('value') + 1).setValue(map[key]);
    else sh.appendRow([key, map[key]]);
  });
  cacheClear_();
  return {success:true};
}

function getKitchenTicketData(orderId) {
  const receipt = buildReceipt_(orderId);
  return { success:true, ticket:{ orderId:receipt.orderId, tableNo:receipt.tableNo, items:receipt.items, note:receipt.note, createdAt:receipt.createdAt, shopName:receipt.restaurantName } };
}

function getSalesReport(fromDate, toDate) {
  const rows = getRows_('payments').map(r => ({
    id: r.id, orderId: r.order_id, totalAmount: Number(r.total_amount || 0), method: r.method, paidAt: r.paid_at, staffName: r.staff_name || '', shiftId:r.shift_id || ''
  }));
  return rows.filter(r => { const d = String(r.paidAt || '').slice(0,10); return (!fromDate || d >= fromDate) && (!toDate || d <= toDate); });
}

function getDbInfo() {
  const ss = getDb_();
  return { spreadsheetId:ss.getId(), spreadsheetUrl:ss.getUrl(), name:ss.getName() };
}

function resetDemoData() {
  const dbInfo = getDbInfo();
  PropertiesService.getScriptProperties().deleteProperty('SPREADSHEET_ID');
  setupSystem();
  return {success:true, previous:dbInfo};
}

function getPrinterBridgeConfig() {
  const s = getSystemSettings();
  return {success:true, enabled:s.printerBridgeEnabled, bridgeToken:s.printerBridgeToken, defaultPrinter:s.printerBridgeDefaultPrinter, pollSeconds:s.printerBridgePollSeconds, note:'bridge แบบ polling queue พร้อมเชื่อม external agent'};
}

function buildShiftPrintableHtml_(summary, lowStock, receipts) {
  const s = getSystemSettings();
  const methods = Object.keys(summary.paymentMethods || {}).map(k => '<tr><td>' + k + '</td><td style="text-align:right">฿' + moneyText_(summary.paymentMethods[k]) + '</td></tr>').join('');
  const payments = (summary.payments || []).map(p => '<tr><td>' + p.orderId + '</td><td>' + (p.method || '-') + '</td><td style="text-align:right">฿' + moneyText_(p.totalAmount) + '</td><td>' + (p.paidAt || '') + '</td></tr>').join('');
  const low = (lowStock || []).map(i => '<tr><td>' + i.name + '</td><td>' + i.qtyOnHand + ' ' + i.unit + '</td><td>' + i.reorderPoint + '</td></tr>').join('');
  const rec = (receipts || []).map(r => '<tr><td>' + r.poId + '</td><td>' + r.supplierName + '</td><td style="text-align:right">฿' + moneyText_(r.totalAmount) + '</td><td>' + r.receivedAt + '</td></tr>').join('');
  return '<html><head><meta charset="utf-8"><style>body{font-family:Sarabun,Arial,sans-serif;padding:20px;color:#111} h1,h2,h3{margin:0 0 10px} table{width:100%;border-collapse:collapse;margin:10px 0 18px} th,td{border:1px solid #ddd;padding:8px;font-size:12px} th{background:#f5f5f5;text-align:left} .sum{display:grid;grid-template-columns:repeat(2,minmax(220px,1fr));gap:10px;margin-bottom:16px}.box{border:1px solid #ddd;padding:12px;border-radius:10px}.right{text-align:right}</style></head><body>' +
  '<h1>' + s.restaurantName + '</h1><h2>รายงานปิดกะ</h2><div>กะ: ' + summary.shift.id + ' | เปิด: ' + (summary.shift.openedAt||'-') + ' | ปิด: ' + (summary.shift.closedAt||'-') + '</div>' +
  '<div class="sum"><div class="box"><strong>ยอดขายรวม</strong><div>฿' + moneyText_(summary.totalSales) + '</div></div><div class="box"><strong>เงินสดที่ควรมี</strong><div>฿' + moneyText_(summary.expectedCash) + '</div></div><div class="box"><strong>เงินสดนับจริง</strong><div>฿' + moneyText_(summary.shift.closingCashCounted) + '</div></div><div class="box"><strong>ส่วนต่าง</strong><div>฿' + moneyText_(summary.shift.cashDiff) + '</div></div></div>' +
  '<h3>สรุปตามวิธีชำระ</h3><table><thead><tr><th>วิธีชำระ</th><th class="right">ยอด</th></tr></thead><tbody>' + methods + '</tbody></table>' +
  '<h3>รายการชำระเงิน</h3><table><thead><tr><th>ออเดอร์</th><th>วิธีชำระ</th><th class="right">ยอด</th><th>เวลา</th></tr></thead><tbody>' + payments + '</tbody></table>' +
  '<h3>วัตถุดิบใกล้หมด</h3><table><thead><tr><th>รายการ</th><th>คงเหลือ</th><th>จุดสั่งซื้อ</th></tr></thead><tbody>' + (low || '<tr><td colspan="3">ไม่มี</td></tr>') + '</tbody></table>' +
  '<h3>รับเข้าล่าสุด</h3><table><thead><tr><th>PO</th><th>ซัพพลายเออร์</th><th class="right">ยอดรับเข้า</th><th>เวลา</th></tr></thead><tbody>' + (rec || '<tr><td colspan="4">ไม่มี</td></tr>') + '</tbody></table>' +
  '</body></html>';
}

// ===== Helpers =====
function getDb_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) return null;
  try { return SpreadsheetApp.openById(id); } catch (err) { throw new Error('ไม่สามารถเปิดฐานข้อมูล Google Sheets ได้'); }
}
function getSheet_(name) {
  const ss = getDb_();
  if (!ss) throw new Error('ยังไม่ได้ตั้งค่าระบบ กรุณารัน setupSystem() ก่อน');
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('ไม่พบชีต ' + name);
  return sh;
}
function getHeaders_(sh) { return sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String); }
function getRows_(sheetName) {
  const sh = getSheet_(sheetName);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const values = sh.getRange(1,1,lastRow,sh.getLastColumn()).getValues();
  const headers = values[0].map(String);
  return values.slice(1).filter(r => r.join('') !== '').map(r => { const obj = {}; headers.forEach((h,i) => obj[h] = r[i]); return obj; });
}
function writeSheet_(sh, headers, rows) {
  sh.clear();
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  if (rows && rows.length) sh.getRange(2,1,rows.length,headers.length).setValues(rows);
}
function styleSheets_(ss) {
  ss.getSheets().forEach(sh => {
    const lastCol = Math.max(2, sh.getLastColumn());
    sh.getRange(1,1,1,lastCol).setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, lastCol);
  });
}
function nextId_(prefix, sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return prefix + '001';
  const vals = sh.getRange(2,1,lastRow-1,1).getValues().flat().filter(String);
  let max = 0;
  vals.forEach(v => { const n = parseInt(String(v).replace(/\D/g,''),10); if (!isNaN(n) && n > max) max = n; });
  return prefix + Utilities.formatString('%03d', max + 1);
}
function updateRowById_(sheetName, id, mutator) {
  const sh = getSheet_(sheetName);
  const headers = getHeaders_(sh);
  const idCol = headers.indexOf('id') + 1;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) throw new Error('ไม่พบข้อมูล');
  const vals = sh.getRange(2,1,lastRow-1,headers.length).getValues();
  const idx = vals.findIndex(r => String(r[idCol - 1]) === String(id));
  if (idx === -1) throw new Error('ไม่พบข้อมูล id ' + id);
  const rowObj = {}; headers.forEach((h,i) => rowObj[h] = vals[idx][i]);
  const updated = mutator(rowObj) || rowObj;
  sh.getRange(idx + 2, 1, 1, headers.length).setValues([headers.map(h => updated[h])]);
  return updated;
}
function resolveUser_(userId) { return userId ? (getUsers_().find(u => u.id === userId) || null) : null; }
function getUsers_() { return getRows_('users'); }
function releaseTableIfNoActiveOrders_(tableId) {
  const active = getRows_('orders').filter(r => String(r.table_id) === String(tableId) && ['pending','cooking','ready','served'].includes(r.status));
  if (!active.length) updateRowById_('tables', tableId, row => { row.status = 'available'; return row; });
}
function buildReceipt_(orderId) {
  const settings = getSystemSettings();
  const order = getRows_('orders').find(r => r.id === orderId);
  if (!order) throw new Error('ไม่พบออเดอร์');
  const items = getRows_('order_items').filter(r => r.order_id === orderId).map(r => ({ name:r.menu_name, qty:Number(r.qty || 0), unitPrice:Number(r.unit_price || 0), lineTotal:Number(r.line_total || 0) }));
  const paymentLines = getRows_('payment_lines').filter(r => r.order_id === orderId).map(r => ({ method:r.method, amount:Number(r.amount||0) }));
  return {
    orderId, tableNo: order.table_no, createdAt: order.created_at, paidAt: order.paid_at || '', restaurantName: settings.restaurantName, restaurantPhone: settings.restaurantPhone, restaurantAddress: settings.restaurantAddress,
    note: order.note || '', subtotal:Number(order.subtotal || 0), vat:Number(order.vat || 0), totalAmount:Number(order.total_amount || 0), items, paymentLines
  };
}
function logActivity_(action, entityType, entityId, detail, actor) {
  try {
    getSheet_('activity_log').appendRow([nextId_('L', getSheet_('activity_log')), action, entityType, entityId, JSON.stringify(detail || {}), new Date().toISOString(), actor ? actor.id : '', actor ? actor.name : '']);
  } catch(err) {}
}
function deductStockForOrder_(orderId, items, user) {
  const menuMap = {};
  getRows_('menu_items').forEach(m => menuMap[m.id] = m);
  items.forEach(it => {
    const menu = menuMap[it.menuId || ''];
    if (!menu) return;
    let recipe = [];
    try { recipe = JSON.parse(menu.recipe_json || '[]'); } catch(err) { recipe = []; }
    recipe.forEach(r => {
      const delta = -1 * Number(r.qty || 0) * Number(it.qty || 0);
      if (!delta) return;
      adjustStock({
        stockId: r.stockId, qtyDelta: delta, movementType: 'out', referenceType: 'order', referenceId: orderId,
        note: `ตัดจากออเดอร์ ${orderId} / ${it.name || menu.name}`, actorId: user ? user.id : ''
      });
    });
  });
}
function getActiveOrdersForTable_(tableId) { return getOrders({tableId:tableId}).filter(o => ['pending','cooking','ready','served'].includes(o.status)); }
function appendShiftEvent_(shiftId, eventType, referenceId, amount, detail) { getSheet_('shift_events').appendRow([nextId_('SE', getSheet_('shift_events')), shiftId, eventType, referenceId || '', Number(amount || 0), JSON.stringify(detail || {}), new Date().toISOString()]); }
function getShiftExpectedCash_(shiftId) {
  const shift = getRows_('shifts').find(s => s.id === shiftId);
  if (!shift) return 0;
  const opening = Number(shift.opening_cash || 0);
  const payments = getRows_('payment_lines').filter(pl => {
    const p = getRows_('payments').find(pp => pp.id === pl.payment_id);
    return p && p.shift_id === shiftId && pl.method === 'cash';
  });
  return round2_(opening + payments.reduce((s,p)=>s+Number(p.amount || 0),0));
}
function mapShift_(r) {
  return {
    id:r.id, openedAt:r.opened_at, openedByName:r.opened_by_name || '', openingCash:Number(r.opening_cash || 0), status:r.status || 'closed', closedAt:r.closed_at || '',
    closedByName:r.closed_by_name || '', closingCashCounted:Number(r.closing_cash_counted || 0), expectedCash:Number(r.expected_cash || 0), cashDiff:Number(r.cash_diff || 0), note:r.note || ''
  };
}
function sum_(rows, field) { return round2_(rows.reduce((s, r) => s + Number(r[field] || 0), 0)); }
function round2_(n) { return Math.round(Number(n || 0) * 100) / 100; }
function isTrue_(v) { return v === true || v === 1 || String(v).toUpperCase() === 'TRUE' || String(v) === '1'; }
function moneyText_(n){ return Number(n||0).toLocaleString('th-TH',{minimumFractionDigits:2,maximumFractionDigits:2}); }
function cacheClear_() { CacheService.getScriptCache().remove(CACHE_KEY_SETTINGS); }


function getOrderItemsRaw_(orderId) {
  return getRows_('order_items').filter(r => r.order_id === orderId).map(r => ({
    id: r.id,
    orderId: r.order_id,
    menuId: r.menu_id,
    name: r.menu_name,
    qty: Number(r.qty || 0),
    unitPrice: Number(r.unit_price || 0),
    lineTotal: Number(r.line_total || 0),
    note: r.note || ''
  }));
}

function getOrderPaymentSummary(orderId) {
  const order = getRows_('orders').find(r => r.id === orderId);
  if (!order) throw new Error('ไม่พบออเดอร์');
  const items = getOrderItemsRaw_(orderId);
  const paymentLines = getRows_('payment_lines').filter(r => r.order_id === orderId).map(r => ({
    id:r.id, method:r.method, amount:Number(r.amount||0), cashReceived:Number(r.cash_received||0), changeAmount:Number(r.change_amount||0), createdAt:r.created_at
  }));
  const refunds = getRows_('refunds').filter(r => r.order_id === orderId).map(r => ({
    id:r.id, refundType:r.refund_type, amount:Number(r.amount||0), reason:r.reason||'', createdAt:r.created_at, status:r.status||'approved'
  }));
  const voids = getRows_('void_logs').filter(r => r.order_id === orderId).map(r => ({
    id:r.id, amount:Number(r.amount||0), reason:r.reason||'', createdAt:r.created_at
  }));
  return {
    success:true,
    order:getOrders({}).find(o => o.id === orderId),
    items:items,
    paymentLines:paymentLines,
    refunds:refunds,
    voids:voids
  };
}

function processPayment(payload) {
  payload = payload || {};
  if (payload.personSplits && payload.personSplits.length) return processPersonItemSplitPayment(payload);
  if (payload.splitLines && payload.splitLines.length) return processSplitPayment(payload);
  const order = getRows_('orders').find(r => r.id === payload.orderId);
  if (!order) throw new Error('ไม่พบออเดอร์');
  const total = Number(order.total_amount || 0);
  const cashReceived = Number(payload.cashReceived || total);
  const method = payload.method || 'cash';
  const changeAmount = method === 'cash' ? Math.max(0, cashReceived - total) : 0;
  const user = resolveUser_(payload.staffId);
  const paidAt = new Date().toISOString();
  const paymentId = nextId_('P', getSheet_('payments'));
  const currentShift = getCurrentShift();

  updateRowById_('orders', payload.orderId, row => { row.status = 'paid'; row.updated_at = paidAt; row.paid_at = paidAt; return row; });
  getSheet_('payments').appendRow([
    paymentId, payload.orderId, order.table_id, method, Number(order.subtotal || 0), Number(order.vat || 0), total, cashReceived, changeAmount,
    user ? user.id : '', user ? user.name : '', paidAt, currentShift ? currentShift.id : ''
  ]);
  getSheet_('payment_lines').appendRow([nextId_('PL', getSheet_('payment_lines')), paymentId, payload.orderId, method, total, cashReceived, changeAmount, paidAt]);
  if (currentShift) appendShiftEvent_(currentShift.id, 'payment', payload.orderId, total, {method});
  releaseTableIfNoActiveOrders_(order.table_id);
  logActivity_('payment', 'order', payload.orderId, payload, user);
  return {success:true, receiptData: buildReceipt_(payload.orderId), paymentId};
}

function processPersonItemSplitPayment(payload) {
  payload = payload || {};
  const order = getRows_('orders').find(r => r.id === payload.orderId);
  if (!order) throw new Error('ไม่พบออเดอร์');
  const items = getOrderItemsRaw_(payload.orderId).filter(i => i.lineTotal > 0);
  const itemMap = {};
  items.forEach(i => itemMap[i.id] = i);
  const splits = (payload.personSplits || []).map(s => ({
    person: String(s.person || ''),
    itemIds: (s.itemIds || []).map(String).filter(Boolean),
    method: String(s.method || 'cash'),
    cashReceived: Number(s.cashReceived || 0)
  })).filter(s => s.itemIds.length);
  if (!splits.length) throw new Error('กรุณาเลือกรายการต่อคน');
  const picked = [];
  splits.forEach(s => s.itemIds.forEach(id => picked.push(id)));
  const unique = Array.from(new Set(picked));
  if (unique.length !== picked.length) throw new Error('มีรายการซ้ำระหว่างคน');
  const missing = items.filter(i => !unique.includes(i.id));
  if (missing.length) throw new Error('ยังมีบางรายการที่ยังไม่ถูกจัดให้คนชำระ');
  const lines = splits.map(s => {
    const amount = round2_(s.itemIds.reduce((sum,id)=>sum + Number((itemMap[id]||{}).lineTotal || 0),0));
    return {
      method:s.method,
      amount:amount,
      cashReceived:s.method === 'cash' ? Number(s.cashReceived || amount) : amount,
      note:s.person
    };
  }).filter(l => l.amount > 0);
  return processSplitPayment({orderId:payload.orderId, splitLines:lines, staffId:payload.staffId});
}

function voidOrderItems(payload) {
  payload = payload || {};
  const order = getRows_('orders').find(r => r.id === payload.orderId);
  if (!order) throw new Error('ไม่พบออเดอร์');
  if (String(order.status || '') === 'paid') throw new Error('บิลนี้ชำระแล้ว ใช้ partial refund แทน');
  const itemIds = (payload.itemIds || []).map(String).filter(Boolean);
  if (!itemIds.length) throw new Error('กรุณาเลือกรายการที่จะ void');
  const items = getOrderItemsRaw_(payload.orderId).filter(i => itemIds.includes(i.id) && i.lineTotal > 0);
  if (!items.length) throw new Error('ไม่พบรายการที่เลือก');
  const amount = round2_(items.reduce((s,i)=>s+i.lineTotal,0));
  items.forEach(it => {
    updateRowById_('order_items', it.id, row => {
      row.note = [row.note || '', '[VOID]', payload.reason || ''].filter(Boolean).join(' ');
      row.line_total = 0;
      row.qty = 0;
      return row;
    });
  });
  recalcOrderTotals_(payload.orderId);
  const user = resolveUser_(payload.actorId);
  getSheet_('void_logs').appendRow([
    nextId_('VD', getSheet_('void_logs')), payload.orderId, order.table_id, JSON.stringify(itemIds), amount, payload.reason || '', new Date().toISOString(), user ? user.id : '', user ? user.name : ''
  ]);
  const currentShift = getCurrentShift();
  if (currentShift) appendShiftEvent_(currentShift.id, 'void', payload.orderId, -amount, {itemIds:itemIds, reason:payload.reason || ''});
  logActivity_('void_items', 'order', payload.orderId, {itemIds:itemIds, amount:amount, reason:payload.reason || ''}, user);
  return {success:true, amount:amount};
}

function partialRefund(payload) {
  payload = payload || {};
  const order = getRows_('orders').find(r => r.id === payload.orderId);
  if (!order) throw new Error('ไม่พบออเดอร์');
  if (String(order.status || '') !== 'paid') throw new Error('ทำ partial refund ได้เฉพาะบิลที่จ่ายแล้ว');
  const itemIds = (payload.itemIds || []).map(String).filter(Boolean);
  if (!itemIds.length) throw new Error('กรุณาเลือกรายการที่จะ refund');
  const items = getOrderItemsRaw_(payload.orderId).filter(i => itemIds.includes(i.id));
  if (!items.length) throw new Error('ไม่พบรายการที่จะ refund');
  const amount = round2_(items.reduce((s,i)=>s+i.lineTotal,0));
  if (amount <= 0) throw new Error('ยอด refund ต้องมากกว่า 0');
  const user = resolveUser_(payload.actorId);
  const refundId = nextId_('RF', getSheet_('refunds'));
  getSheet_('refunds').appendRow([
    refundId, payload.orderId, order.table_id, 'partial_item_refund', JSON.stringify(itemIds), amount, payload.reason || '', 'approved', new Date().toISOString(), user ? user.id : '', user ? user.name : '', payload.paymentId || ''
  ]);
  const currentShift = getCurrentShift();
  if (currentShift) appendShiftEvent_(currentShift.id, 'refund', payload.orderId, -amount, {itemIds:itemIds, reason:payload.reason || ''});
  logActivity_('partial_refund', 'order', payload.orderId, {itemIds:itemIds, amount:amount, reason:payload.reason || ''}, user);
  return {success:true, refundId:refundId, amount:amount};
}

function recalcOrderTotals_(orderId) {
  const items = getOrderItemsRaw_(orderId);
  const subtotal = round2_(items.reduce((s,i)=>s + Number(i.lineTotal || 0),0));
  const settings = getSystemSettings();
  const vat = settings.vatEnabled ? round2_(subtotal * Number(settings.vatRate || 0) / 100) : 0;
  const total = round2_(subtotal + vat);
  updateRowById_('orders', orderId, row => {
    row.subtotal = subtotal;
    row.vat = vat;
    row.total_amount = total;
    row.updated_at = new Date().toISOString();
    if (total <= 0 && row.status !== 'paid') row.status = 'cancelled';
    return row;
  });
  const order = getRows_('orders').find(r => r.id === orderId);
  if (order && (String(order.status || '') === 'cancelled' || total <= 0)) releaseTableIfNoActiveOrders_(order.table_id);
  return {subtotal:subtotal, vat:vat, total:total};
}

function getRefunds(limit) {
  return getRows_('refunds').map(r => ({
    id:r.id, orderId:r.order_id, tableId:r.table_id, refundType:r.refund_type, amount:Number(r.amount||0), reason:r.reason||'', status:r.status||'', createdAt:r.created_at, actorName:r.actor_name||''
  })).sort((a,b)=>String(b.createdAt).localeCompare(String(a.createdAt))).slice(0, Number(limit || 100));
}

function saveSupplierPayment(payload) {
  payload = payload || {};
  if (!payload.supplierId) throw new Error('กรุณาเลือกซัพพลายเออร์');
  const amount = Number(payload.amount || 0);
  if (amount <= 0) throw new Error('กรุณาระบุจำนวนเงิน');
  const supplier = getRows_('suppliers').find(r => r.id === payload.supplierId);
  if (!supplier) throw new Error('ไม่พบซัพพลายเออร์');
  const user = resolveUser_(payload.actorId);
  getSheet_('supplier_payments').appendRow([
    nextId_('SP', getSheet_('supplier_payments')), supplier.id, supplier.name, payload.poId || '', amount, payload.method || 'transfer', payload.referenceNo || '', payload.note || '', new Date().toISOString(), user ? user.id : '', user ? user.name : ''
  ]);
  logActivity_('supplier_payment', 'supplier', supplier.id, payload, user);
  return {success:true};
}

function getSupplierPayments(limit) {
  return getRows_('supplier_payments').map(r => ({
    id:r.id, supplierId:r.supplier_id, supplierName:r.supplier_name, poId:r.po_id || '', amount:Number(r.amount||0), method:r.method||'', referenceNo:r.reference_no||'', note:r.note||'', paidAt:r.paid_at, actorName:r.actor_name||''
  })).sort((a,b)=>String(b.paidAt).localeCompare(String(a.paidAt))).slice(0, Number(limit || 100));
}

function getSupplierBalances() {
  const suppliers = getSuppliers();
  const receipts = getGoodsReceipts(1000);
  const poMap = {};
  getPurchaseOrders('all').forEach(po => poMap[po.id] = po);
  const payments = getSupplierPayments(1000);
  return suppliers.map(s => {
    const totalReceived = round2_(receipts.filter(r => r.supplierName === s.name).reduce((sum,r)=>sum + Number(r.totalAmount||0),0));
    const totalPaid = round2_(payments.filter(p => p.supplierId === s.id).reduce((sum,p)=>sum + Number(p.amount||0),0));
    return {
      supplierId:s.id,
      supplierName:s.name,
      totalReceived:totalReceived,
      totalPaid:totalPaid,
      balance:round2_(totalReceived - totalPaid)
    };
  });
}

function createStockCountSession(payload) {
  payload = payload || {};
  const user = resolveUser_(payload.actorId);
  const id = nextId_('SCN', getSheet_('stock_counts'));
  const now = new Date().toISOString();
  getSheet_('stock_counts').appendRow([id, 'draft', payload.countDate || now.slice(0,10), payload.note || '', now, user ? user.id : '', user ? user.name : '', '', '', '']);
  const lineSh = getSheet_('stock_count_lines');
  getInventorySummary().items.forEach(item => {
    lineSh.appendRow([nextId_('SCL', lineSh), id, item.id, item.name, item.qtyOnHand, '', '', '']);
  });
  return {success:true, stockCountId:id};
}

function getStockCountSessions(limit) {
  return getRows_('stock_counts').map(r => ({
    id:r.id, status:r.status || 'draft', countDate:r.count_date || '', note:r.note || '', createdAt:r.created_at, createdByName:r.created_by_name || '', appliedAt:r.applied_at || '', appliedByName:r.applied_by_name || ''
  })).sort((a,b)=>String(b.createdAt).localeCompare(String(a.createdAt))).slice(0, Number(limit || 50));
}

function getStockCountDetail(stockCountId) {
  const head = getRows_('stock_counts').find(r => r.id === stockCountId);
  if (!head) throw new Error('ไม่พบรอบนับสต๊อก');
  const lines = getRows_('stock_count_lines').filter(r => r.stock_count_id === stockCountId).map(r => ({
    id:r.id, stockId:r.stock_id, stockName:r.stock_name, systemQty:Number(r.system_qty||0), countedQty:r.counted_qty === '' ? '' : Number(r.counted_qty||0), diffQty:r.diff_qty === '' ? '' : Number(r.diff_qty||0), note:r.note||''
  }));
  return {success:true, session:{
    id:head.id, status:head.status||'draft', countDate:head.count_date||'', note:head.note||'', createdAt:head.created_at
  }, lines:lines};
}

function applyStockCount(payload) {
  payload = payload || {};
  const stockCountId = String(payload.stockCountId || '');
  const user = resolveUser_(payload.actorId);
  const head = getRows_('stock_counts').find(r => r.id === stockCountId);
  if (!head) throw new Error('ไม่พบรอบนับสต๊อก');
  if (String(head.status || '') === 'applied') throw new Error('รอบนี้ถูก apply แล้ว');
  const lineInputs = payload.lines || [];
  if (!lineInputs.length) throw new Error('ไม่มีรายการนับสต๊อก');
  lineInputs.forEach(l => {
    const counted = Number(l.countedQty || 0);
    const line = getRows_('stock_count_lines').find(r => r.id === l.id && r.stock_count_id === stockCountId);
    if (!line) return;
    const systemQty = Number(line.system_qty || 0);
    const diff = round2_(counted - systemQty);
    updateRowById_('stock_count_lines', line.id, row => {
      row.counted_qty = counted;
      row.diff_qty = diff;
      row.note = l.note || '';
      return row;
    });
    if (Math.abs(diff) > 0.00001) {
      adjustStock({
        stockId: line.stock_id,
        qtyDelta: diff,
        movementType: 'stock_take',
        referenceType: 'stock_count',
        referenceId: stockCountId,
        note: l.note || ('stock take ' + stockCountId),
        actorId: user ? user.id : ''
      });
    }
  });
  updateRowById_('stock_counts', stockCountId, row => {
    row.status = 'applied';
    row.applied_at = new Date().toISOString();
    row.applied_by_id = user ? user.id : '';
    row.applied_by_name = user ? user.name : '';
    return row;
  });
  logActivity_('apply_stock_take', 'stock_count', stockCountId, {lineCount:lineInputs.length}, user);
  return {success:true};
}

function classifyMenuStation_(item) {
  const category = String(item.category || '').trim();
  if (category === 'เครื่องดื่ม') return 'bar';
  return 'kitchen';
}

function queueKitchenPrint(orderId) {
  const order = getOrders({}).find(o => o.id === orderId);
  if (!order) throw new Error('ไม่พบออเดอร์');
  const byStation = {kitchen:[], bar:[]};
  (order.itemsArray || []).forEach(it => {
    const menu = getRows_('menu_items').find(m => m.id === it.menuId);
    const station = classifyMenuStation_(menu || {category:''});
    byStation[station].push(it);
  });
  const sh = getSheet_('print_queue');
  const created = [];
  Object.keys(byStation).forEach(station => {
    if (!byStation[station].length) return;
    const qId = nextId_('Q', sh);
    sh.appendRow([
      qId, station + '_ticket', orderId, JSON.stringify({order:{id:order.id, tableNo:order.tableNo, createdAt:order.createdAt, itemsArray:byStation[station], note:order.note || ''}, station:station}), 'queued', new Date().toISOString(), '', station === 'bar' ? 'BAR-1' : 'KITCHEN-1', 'สร้างคิวพิมพ์อัตโนมัติ'
    ]);
    created.push(qId);
  });
  return {success:true, queueIds:created};
}

function getPrintQueue(status, station) {
  let rows = getRows_('print_queue').map(r => ({
    id:r.id, queueType:r.queue_type, referenceId:r.reference_id, payloadJson:r.payload_json || '{}', status:r.status || 'queued', createdAt:r.created_at, printedAt:r.printed_at || '', printerName:r.printer_name || '', note:r.note || ''
  })).sort((a,b)=>String(b.createdAt).localeCompare(String(a.createdAt)));
  if (status && status !== 'all') rows = rows.filter(r => r.status === status);
  if (station && station !== 'all') rows = rows.filter(r => String(r.queueType || '').indexOf(station) === 0);
  return rows;
}

function getKitchenDisplayData(station) {
  station = station || 'kitchen';
  const orders = getOrders({status:'all'}).filter(o => ['pending','cooking','ready','served'].includes(o.status));
  const display = orders.map(o => {
    const items = (o.itemsArray || []).filter(it => {
      const menu = getRows_('menu_items').find(m => m.id === it.menuId);
      return classifyMenuStation_(menu || {category:''}) === station;
    });
    return {
      id:o.id,
      tableNo:o.tableNo,
      status:o.status,
      createdAt:o.createdAt,
      updatedAt:o.updatedAt,
      note:o.note || '',
      items:items
    };
  }).filter(o => o.items.length);
  const queue = getPrintQueue('all', station);
  return {success:true, station:station, orders:display, printQueue:queue, settings:getSystemSettings()};
}

function getAdvancedShiftDashboard(shiftId) {
  const shift = shiftId ? getAllShifts(200).find(s => s.id === shiftId) : getCurrentShift();
  if (!shift) throw new Error('ไม่พบกะ');
  const summary = getShiftSummary(shift.id);
  const paymentLines = getRows_('payment_lines').filter(r => {
    const paid = String(r.created_at || '');
    return paid >= String(shift.openedAt || '') && (!shift.closedAt || paid <= String(shift.closedAt || '9999'));
  });
  const byMethod = {};
  paymentLines.forEach(r => byMethod[r.method] = round2_((byMethod[r.method] || 0) + Number(r.amount || 0)));
  const refunds = getRows_('refunds').filter(r => String(r.created_at || '') >= String(shift.openedAt || '') && (!shift.closedAt || String(r.created_at || '') <= String(shift.closedAt || '9999')));
  const voids = getRows_('void_logs').filter(r => String(r.created_at || '') >= String(shift.openedAt || '') && (!shift.closedAt || String(r.created_at || '') <= String(shift.closedAt || '9999')));
  const supplierPayments = getRows_('supplier_payments').filter(r => String(r.paid_at || '') >= String(shift.openedAt || '') && (!shift.closedAt || String(r.paid_at || '') <= String(shift.closedAt || '9999')));
  const orderItems = getRows_('order_items');
  const orders = getRows_('orders').filter(r => String(r.created_at || '') >= String(shift.openedAt || '') && (!shift.closedAt || String(r.created_at || '') <= String(shift.closedAt || '9999')));
  const topMap = {};
  orders.forEach(o => {
    orderItems.filter(i => i.order_id === o.id).forEach(it => {
      if (Number(it.qty || 0) <= 0) return;
      topMap[it.menu_name] = (topMap[it.menu_name] || 0) + Number(it.qty || 0);
    });
  });
  const topItems = Object.keys(topMap).map(name => ({name:name, qty:topMap[name]})).sort((a,b)=>b.qty-a.qty).slice(0,10);
  return {
    success:true,
    shift:shift,
    summary:summary,
    byMethod:byMethod,
    refunds:{count:refunds.length, total:round2_(refunds.reduce((s,r)=>s+Number(r.amount||0),0))},
    voids:{count:voids.length, total:round2_(voids.reduce((s,r)=>s+Number(r.amount||0),0))},
    supplierPayments:{count:supplierPayments.length, total:round2_(supplierPayments.reduce((s,r)=>s+Number(r.amount||0),0))},
    topItems:topItems
  };
}

function buildShiftPrintableHtml_(summary, lowStock, receipts) {
  const adv = getAdvancedShiftDashboard(summary.shift.id);
  const methodRows = Object.keys(adv.byMethod).map(k => '<tr><td>'+k+'</td><td style="text-align:right">฿'+money_(adv.byMethod[k])+'</td></tr>').join('');
  const topRows = (adv.topItems || []).map(r => '<tr><td>'+r.name+'</td><td style="text-align:right">'+r.qty+'</td></tr>').join('');
  return '<!DOCTYPE html><html><head><meta charset="utf-8"><title>Shift Report</title><style>body{font-family:Arial,sans-serif;padding:24px;color:#111}h1,h2,h3{margin:0 0 10px}table{width:100%;border-collapse:collapse;margin-top:10px}th,td{border-bottom:1px solid #ddd;padding:8px;text-align:left} .grid{display:grid;grid-template-columns:1fr 1fr;gap:16px} .card{border:1px solid #ddd;border-radius:12px;padding:14px}</style></head><body>'+
    '<h1>3 TIME - Close Shift Report</h1>'+
    '<div>Shift: '+summary.shift.id+' | Opened: '+summary.shift.openedAt+' | Closed: '+(summary.shift.closedAt || '-')+'</div>'+
    '<div class="grid" style="margin-top:16px">'+
      '<div class="card"><h3>ภาพรวมกะ</h3><table>'+
        '<tr><td>ยอดขาย</td><td style="text-align:right">฿'+money_(summary.totalSales)+'</td></tr>'+
        '<tr><td>เงินเปิดกะ</td><td style="text-align:right">฿'+money_(summary.shift.openingCash)+'</td></tr>'+
        '<tr><td>เงินสดคาดหวัง</td><td style="text-align:right">฿'+money_(summary.expectedCash)+'</td></tr>'+
        '<tr><td>เงินสดนับจริง</td><td style="text-align:right">฿'+money_(summary.shift.closingCashCounted || 0)+'</td></tr>'+
        '<tr><td>ส่วนต่าง</td><td style="text-align:right">฿'+money_(summary.cashDiff || 0)+'</td></tr>'+
        '<tr><td>จำนวนบิล</td><td style="text-align:right">'+summary.payments.length+'</td></tr>'+
      '</table></div>'+
      '<div class="card"><h3>ตามวิธีชำระ</h3><table>'+methodRows+'</table></div>'+
    '</div>'+
    '<div class="grid" style="margin-top:16px">'+
      '<div class="card"><h3>Refund / Void</h3><table>'+
        '<tr><td>Refund</td><td style="text-align:right">'+adv.refunds.count+' รายการ / ฿'+money_(adv.refunds.total)+'</td></tr>'+
        '<tr><td>Void</td><td style="text-align:right">'+adv.voids.count+' รายการ / ฿'+money_(adv.voids.total)+'</td></tr>'+
        '<tr><td>จ่ายเจ้าหนี้</td><td style="text-align:right">'+adv.supplierPayments.count+' รายการ / ฿'+money_(adv.supplierPayments.total)+'</td></tr>'+
      '</table></div>'+
      '<div class="card"><h3>ขายดีในกะ</h3><table>'+topRows+'</table></div>'+
    '</div>'+
    '<div class="card" style="margin-top:16px"><h3>ของใกล้หมด</h3><table>'+
      (lowStock || []).map(i => '<tr><td>'+i.name+'</td><td>'+i.qtyOnHand+' '+i.unit+'</td><td>ROP '+i.reorderPoint+'</td></tr>').join('') +
    '</table></div>'+
    '<div class="card" style="margin-top:16px"><h3>รับสินค้าเข้าล่าสุด</h3><table>'+
      (receipts || []).map(r => '<tr><td>'+r.receivedAt+'</td><td>'+r.poId+'</td><td>'+r.supplierName+'</td><td style="text-align:right">฿'+money_(r.totalAmount)+'</td></tr>').join('') +
    '</table></div>'+
  '</body></html>';
}

function money_(n) {
  return Utilities.formatString('%,.2f', Number(n || 0));
}


// ===== V5 PATCH =====
function ensureV5Schema_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) return;
  const ss = SpreadsheetApp.openById(id);
  const requiredSheets = {
    receipt_reprints:['id','order_id','payment_id','reprinted_at','actor_id','actor_name','note'],
    receipt_cancellations:['id','order_id','payment_id','amount','reason','created_at','actor_id','actor_name']
  };
  Object.keys(requiredSheets).forEach(function(name) {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      writeSheet_(sh, requiredSheets[name], []);
    } else if (sh.getLastRow() < 1) {
      writeSheet_(sh, requiredSheets[name], []);
    }
  });
  const ensureCols = [
    ['orders',['payment_status','paid_amount','balance_due','last_payment_at']],
    ['payment_lines',['note','is_cancelled']],
    ['payments',['is_cancelled','cancelled_at','cancel_reason']]
  ];
  ensureCols.forEach(function(pair) {
    const sh = ss.getSheetByName(pair[0]);
    if (!sh) return;
    const headers = getHeaders_(sh);
    let changed = false;
    pair[1].forEach(function(col) {
      if (headers.indexOf(col) === -1) {
        sh.insertColumnAfter(sh.getLastColumn());
        sh.getRange(1, sh.getLastColumn()).setValue(col);
        changed = true;
      }
    });
    if (changed) sh.autoResizeColumns(1, sh.getLastColumn());
  });
}

function getDb_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) return null;
  try {
    const ss = SpreadsheetApp.openById(id);
    ensureV5Schema_();
    return ss;
  } catch (err) {
    throw new Error('ไม่สามารถเปิดฐานข้อมูล Google Sheets ได้');
  }
}

function getActiveCancelledPaymentIds_() {
  const cancelled = {};
  getRows_('receipt_cancellations').forEach(function(r) { if (r.payment_id) cancelled[String(r.payment_id)] = true; });
  return cancelled;
}

function getOrderFinancials_(orderId) {
  const order = getRows_('orders').find(function(r){ return String(r.id) === String(orderId); });
  if (!order) throw new Error('ไม่พบออเดอร์');
  const cancelledMap = getActiveCancelledPaymentIds_();
  const payLines = getRows_('payment_lines').filter(function(r){ return String(r.order_id) === String(orderId) && !cancelledMap[String(r.payment_id)]; });
  const paidAmount = round2_(payLines.reduce(function(s,r){ return s + Number(r.amount||0); },0));
  const refunds = getRows_('refunds').filter(function(r){ return String(r.order_id) === String(orderId) && String(r.status||'approved') !== 'cancelled'; });
  const refundedAmount = round2_(refunds.reduce(function(s,r){ return s + Number(r.amount||0); },0));
  const cancelledPayments = getRows_('receipt_cancellations').filter(function(r){ return String(r.order_id) === String(orderId); });
  const cancelledPaymentAmount = round2_(cancelledPayments.reduce(function(s,r){ return s + Number(r.amount||0); },0));
  const total = Number(order.total_amount || 0);
  const balanceDue = round2_(Math.max(0, total - paidAmount));
  let paymentStatus = 'unpaid';
  if (paidAmount > 0 && balanceDue > 0) paymentStatus = 'partial_paid';
  if (balanceDue <= 0.009) paymentStatus = 'paid';
  return { total:total, paidAmount:paidAmount, balanceDue:balanceDue, refundedAmount:refundedAmount, cancelledPaymentAmount:cancelledPaymentAmount, paymentStatus:paymentStatus };
}

function getOrderCogs_(orderId) {
  const items = getOrderItemsRaw_(orderId);
  const menuMap = {};
  getRows_('menu_items').forEach(function(m){ menuMap[m.id] = m; });
  const stockMap = {};
  getRows_('stock_items').forEach(function(s){ stockMap[s.id] = s; });
  let total = 0;
  items.forEach(function(it){
    const menu = menuMap[it.menuId || ''];
    let recipe = [];
    try { recipe = JSON.parse((menu && menu.recipe_json) || '[]'); } catch (err) { recipe = []; }
    recipe.forEach(function(r){
      const stock = stockMap[r.stockId || ''];
      total += Number(it.qty || 0) * Number(r.qty || 0) * Number((stock && stock.cost_per_unit) || 0);
    });
  });
  return round2_(total);
}

function updateOrderPaymentState_(orderId) {
  const fin = getOrderFinancials_(orderId);
  updateRowById_('orders', orderId, function(row) {
    row.payment_status = fin.paymentStatus;
    row.paid_amount = fin.paidAmount;
    row.balance_due = fin.balanceDue;
    row.last_payment_at = new Date().toISOString();
    if (fin.paymentStatus === 'paid') {
      row.status = 'paid';
      if (!row.paid_at) row.paid_at = new Date().toISOString();
    } else if (fin.paymentStatus === 'partial_paid') {
      if (['cancelled','paid'].indexOf(String(row.status||'')) === -1) row.status = 'served';
    } else if (fin.paymentStatus === 'unpaid' && String(row.status||'') === 'paid') {
      row.status = 'served';
      row.paid_at = '';
    }
    row.updated_at = new Date().toISOString();
    return row;
  });
  const order = getRows_('orders').find(function(r){ return String(r.id) === String(orderId); });
  if (order && (fin.paymentStatus === 'paid' || String(order.status||'') === 'cancelled')) releaseTableIfNoActiveOrders_(order.table_id);
  return fin;
}

function getOrders(filter) {
  filter = filter || {};
  const rows = getRows_('orders');
  const items = getRows_('order_items');
  let list = rows.map(function(r) {
    const its = items.filter(function(i){ return i.order_id === r.id; }).map(function(i){
      return {id:i.id, menuId:i.menu_id, name:i.menu_name, qty:Number(i.qty||0), unitPrice:Number(i.unit_price||0), lineTotal:Number(i.line_total||0), note:i.note||''};
    });
    const fin = getOrderFinancials_(r.id);
    return {
      id:r.id, source:r.source||'staff', tableId:r.table_id, tableNo:r.table_no, customerName:r.customer_name||'', status:r.status||'pending',
      subtotal:Number(r.subtotal||0), vat:Number(r.vat||0), totalAmount:Number(r.total_amount||0), note:r.note||'', staffId:r.staff_id||'', staffName:r.staff_name||'',
      createdAt:r.created_at, updatedAt:r.updated_at, paidAt:r.paid_at, itemsArray:its, items:JSON.stringify(its),
      paidAmount:fin.paidAmount, balanceDue:fin.balanceDue, refundedAmount:fin.refundedAmount, paymentStatus:fin.paymentStatus
    };
  }).sort(function(a,b){ return String(b.createdAt).localeCompare(String(a.createdAt)); });
  if (filter.status && filter.status !== 'all') list = list.filter(function(o){ return o.status === filter.status || o.paymentStatus === filter.status; });
  if (filter.tableId) list = list.filter(function(o){ return String(o.tableId) === String(filter.tableId); });
  if (filter.tableNo) list = list.filter(function(o){ return String(o.tableNo) === String(filter.tableNo); });
  return list;
}

function processPayment(payload) {
  payload = payload || {};
  if (payload.personUnitSplits && payload.personUnitSplits.length) return processPersonUnitSplitPayment(payload);
  if (payload.personSplits && payload.personSplits.length) return processPersonItemSplitPayment(payload);
  if (payload.splitLines && payload.splitLines.length) return processSplitPayment(payload);
  const order = getRows_('orders').find(function(r){ return r.id === payload.orderId; });
  if (!order) throw new Error('ไม่พบออเดอร์');
  const fin = getOrderFinancials_(payload.orderId);
  if (fin.balanceDue <= 0.009) throw new Error('บิลนี้ชำระครบแล้ว');
  const amount = round2_(Number(payload.amount || payload.cashReceived || fin.balanceDue));
  if (amount <= 0) throw new Error('จำนวนเงินต้องมากกว่า 0');
  if (amount - fin.balanceDue > 0.01) throw new Error('ยอดรับเกินยอดคงค้าง');
  const cashReceived = Number(payload.cashReceived || amount);
  const method = payload.method || 'cash';
  const changeAmount = method === 'cash' ? Math.max(0, cashReceived - amount) : 0;
  const user = resolveUser_(payload.staffId);
  const paidAt = new Date().toISOString();
  const paymentId = nextId_('P', getSheet_('payments'));
  const currentShift = getCurrentShift();
  getSheet_('payments').appendRow([paymentId, payload.orderId, order.table_id, method, Number(order.subtotal||0), Number(order.vat||0), amount, cashReceived, changeAmount, user ? user.id : '', user ? user.name : '', paidAt, currentShift ? currentShift.id : '', 'FALSE', '', '']);
  getSheet_('payment_lines').appendRow([nextId_('PL', getSheet_('payment_lines')), paymentId, payload.orderId, method, amount, cashReceived, changeAmount, paidAt, payload.note || '', 'FALSE']);
  const updatedFin = updateOrderPaymentState_(payload.orderId);
  if (currentShift) appendShiftEvent_(currentShift.id, updatedFin.paymentStatus === 'paid' ? 'payment' : 'payment_partial', payload.orderId, amount, {method:method});
  logActivity_('payment', 'order', payload.orderId, payload, user);
  return {success:true, receiptData:buildReceipt_(payload.orderId), paymentId:paymentId, financials:updatedFin};
}

function processSplitPayment(payload) {
  payload = payload || {};
  const order = getRows_('orders').find(function(r){ return r.id === payload.orderId; });
  if (!order) throw new Error('ไม่พบออเดอร์');
  const fin = getOrderFinancials_(payload.orderId);
  const limit = fin.balanceDue;
  const lines = (payload.splitLines || []).map(function(l){
    return {method:String(l.method||'cash'), amount:Number(l.amount||0), cashReceived:Number(l.cashReceived || l.amount || 0), note:String(l.note||'')};
  }).filter(function(l){ return l.amount > 0; });
  if (!lines.length) throw new Error('กรุณาระบุรายการชำระ');
  const sumLines = round2_(lines.reduce(function(s,l){ return s + l.amount; },0));
  if (sumLines - limit > 0.01) throw new Error('ยอด split bill เกินยอดคงค้าง');
  const user = resolveUser_(payload.staffId);
  const paidAt = new Date().toISOString();
  const paymentId = nextId_('P', getSheet_('payments'));
  const totalCashReceived = lines.reduce(function(s,l){ return s + l.cashReceived; },0);
  const totalChange = round2_(lines.reduce(function(s,l){ return s + (l.method === 'cash' ? Math.max(0, l.cashReceived - l.amount) : 0); },0));
  const currentShift = getCurrentShift();
  getSheet_('payments').appendRow([paymentId, payload.orderId, order.table_id, 'split', Number(order.subtotal||0), Number(order.vat||0), sumLines, totalCashReceived, totalChange, user ? user.id : '', user ? user.name : '', paidAt, currentShift ? currentShift.id : '', 'FALSE', '', '']);
  const linesSheet = getSheet_('payment_lines');
  lines.forEach(function(line){ linesSheet.appendRow([nextId_('PL', linesSheet), paymentId, payload.orderId, line.method, line.amount, line.cashReceived, line.method === 'cash' ? Math.max(0, line.cashReceived - line.amount) : 0, paidAt, line.note || '', 'FALSE']); });
  const updatedFin = updateOrderPaymentState_(payload.orderId);
  if (currentShift) appendShiftEvent_(currentShift.id, updatedFin.paymentStatus === 'paid' ? 'payment_split' : 'payment_partial_split', payload.orderId, sumLines, {lines:lines});
  logActivity_('payment_split', 'order', payload.orderId, {lines:lines}, user);
  return {success:true, receiptData:buildReceipt_(payload.orderId), paymentId:paymentId, financials:updatedFin};
}

function processPersonUnitSplitPayment(payload) {
  payload = payload || {};
  const order = getRows_('orders').find(function(r){ return r.id === payload.orderId; });
  if (!order) throw new Error('ไม่พบออเดอร์');
  const items = getOrderItemsRaw_(payload.orderId).filter(function(i){ return Number(i.qty||0) > 0 && Number(i.lineTotal||0) > 0; });
  const unitMap = {};
  items.forEach(function(i){ for (var n = 1; n <= Number(i.qty || 0); n++) unitMap[i.id + '#' + n] = { itemId:i.id, amount:Number(i.unitPrice||0), name:i.name + ' #' + n }; });
  const splits = (payload.personUnitSplits || []).map(function(s){ return {person:String(s.person||''), unitKeys:(s.unitKeys||[]).map(String).filter(Boolean), method:String(s.method||'cash'), cashReceived:Number(s.cashReceived||0)}; }).filter(function(s){ return s.unitKeys.length; });
  if (!splits.length) throw new Error('กรุณาเลือกจำนวนต่อคน');
  const picked = [];
  splits.forEach(function(s){ s.unitKeys.forEach(function(k){ picked.push(k); }); });
  const unique = Array.from(new Set(picked));
  if (unique.length !== picked.length) throw new Error('มีจำนวนสินค้าซ้ำระหว่างคน');
  const allKeys = Object.keys(unitMap);
  const missing = allKeys.filter(function(k){ return unique.indexOf(k) === -1; });
  if (missing.length) throw new Error('ยังมีจำนวนสินค้าบางชิ้นที่ยังไม่ถูกจัดให้คนชำระ');
  const lines = splits.map(function(s){
    const amount = round2_(s.unitKeys.reduce(function(sum,k){ return sum + Number((unitMap[k]||{}).amount || 0); },0));
    return {method:s.method, amount:amount, cashReceived:s.method === 'cash' ? Number(s.cashReceived || amount) : amount, note:s.person};
  });
  return processSplitPayment({orderId:payload.orderId, splitLines:lines, staffId:payload.staffId});
}

function getOrderPaymentSummary(orderId) {
  const order = getRows_('orders').find(function(r){ return r.id === orderId; });
  if (!order) throw new Error('ไม่พบออเดอร์');
  const items = getOrderItemsRaw_(orderId);
  const cancelledMap = getActiveCancelledPaymentIds_();
  const paymentLines = getRows_('payment_lines').filter(function(r){ return r.order_id === orderId; }).map(function(r){
    return {id:r.id, paymentId:r.payment_id, method:r.method, amount:Number(r.amount||0), cashReceived:Number(r.cash_received||0), changeAmount:Number(r.change_amount||0), createdAt:r.created_at, note:r.note||'', isCancelled:!!cancelledMap[String(r.payment_id)]};
  });
  const refunds = getRows_('refunds').filter(function(r){ return r.order_id === orderId; }).map(function(r){ return {id:r.id, refundType:r.refund_type, amount:Number(r.amount||0), reason:r.reason||'', createdAt:r.created_at, status:r.status||'approved'}; });
  const voids = getRows_('void_logs').filter(function(r){ return r.order_id === orderId; }).map(function(r){ return {id:r.id, amount:Number(r.amount||0), reason:r.reason||'', createdAt:r.created_at}; });
  const unitItems = [];
  items.forEach(function(i){ for (var n = 1; n <= Number(i.qty || 0); n++) unitItems.push({unitKey:i.id + '#' + n, itemId:i.id, name:i.name, unitNo:n, amount:Number(i.unitPrice||0)}); });
  return {success:true, order:getOrders({}).find(function(o){ return o.id === orderId; }), items:items, unitItems:unitItems, paymentLines:paymentLines, refunds:refunds, voids:voids, financials:getOrderFinancials_(orderId)};
}

function cancelReceipt(payload) {
  payload = payload || {};
  const paymentId = String(payload.paymentId || '');
  if (!paymentId) throw new Error('ไม่พบ payment id');
  const cancelledMap = getActiveCancelledPaymentIds_();
  if (cancelledMap[paymentId]) throw new Error('ใบเสร็จนี้ถูกยกเลิกแล้ว');
  const payment = getRows_('payments').find(function(r){ return String(r.id) === paymentId; });
  if (!payment) throw new Error('ไม่พบรายการชำระ');
  const amount = Number(payment.total_amount || 0);
  const user = resolveUser_(payload.actorId);
  getSheet_('receipt_cancellations').appendRow([nextId_('RC', getSheet_('receipt_cancellations')), payment.order_id, paymentId, amount, payload.reason || '', new Date().toISOString(), user ? user.id : '', user ? user.name : '']);
  updateRowById_('payments', paymentId, function(row){ row.is_cancelled='TRUE'; row.cancelled_at=new Date().toISOString(); row.cancel_reason=payload.reason||''; return row; });
  const fin = updateOrderPaymentState_(payment.order_id);
  const currentShift = getCurrentShift();
  if (currentShift) appendShiftEvent_(currentShift.id, 'receipt_cancel', payment.order_id, -amount, {paymentId:paymentId, reason:payload.reason || ''});
  logActivity_('cancel_receipt', 'payment', paymentId, payload, user);
  return {success:true, paymentId:paymentId, financials:fin};
}

function reprintReceipt(payload) {
  payload = payload || {};
  const orderId = String(payload.orderId || '');
  if (!orderId) throw new Error('ไม่พบ order id');
  const user = resolveUser_(payload.actorId);
  getSheet_('receipt_reprints').appendRow([nextId_('RP', getSheet_('receipt_reprints')), orderId, payload.paymentId || '', new Date().toISOString(), user ? user.id : '', user ? user.name : '', payload.note || 'reprint']);
  logActivity_('reprint_receipt', 'order', orderId, payload, user);
  return {success:true, receiptData:buildReceipt_(orderId), printableHtml:buildReceiptPrintableHtml_(orderId)};
}

function getReceiptAuditLogs(limit) {
  const reprints = getRows_('receipt_reprints').map(function(r){ return {type:'reprint', id:r.id, orderId:r.order_id, paymentId:r.payment_id||'', amount:0, reason:r.note||'', createdAt:r.reprinted_at, actorName:r.actor_name||''}; });
  const cancels = getRows_('receipt_cancellations').map(function(r){ return {type:'cancel', id:r.id, orderId:r.order_id, paymentId:r.payment_id||'', amount:Number(r.amount||0), reason:r.reason||'', createdAt:r.created_at, actorName:r.actor_name||''}; });
  return reprints.concat(cancels).sort(function(a,b){ return String(b.createdAt).localeCompare(String(a.createdAt)); }).slice(0, Number(limit || 100));
}

function buildReceiptPrintableHtml_(orderId) {
  const r = buildReceipt_(orderId);
  const rows = (r.items || []).map(function(i){ return '<tr><td>'+i.name+'</td><td style="text-align:right">'+i.qty+'</td><td style="text-align:right">฿'+money_(i.lineTotal)+'</td></tr>'; }).join('');
  const pays = (r.paymentLines || []).map(function(i){ return '<tr><td>'+i.method+'</td><td style="text-align:right">฿'+money_(i.amount)+'</td></tr>'; }).join('');
  return '<!DOCTYPE html><html><head><meta charset="utf-8"><title>Receipt</title><style>body{font-family:Arial,sans-serif;padding:18px;color:#111}table{width:100%;border-collapse:collapse}td,th{padding:6px;border-bottom:1px solid #ddd;text-align:left}</style></head><body>'+
    '<h2>'+r.restaurantName+'</h2><div>โต๊ะ '+(r.tableNo||'-')+'</div><div>Order '+r.orderId+'</div><div>'+r.restaurantPhone+'</div><div>'+r.restaurantAddress+'</div><hr>'+
    '<table><thead><tr><th>รายการ</th><th style="text-align:right">จำนวน</th><th style="text-align:right">ยอด</th></tr></thead><tbody>'+rows+'</tbody></table>'+
    '<div style="margin-top:10px">Subtotal ฿'+money_(r.subtotal)+'</div><div>Total ฿'+money_(r.totalAmount)+'</div>'+
    '<h4>Payment</h4><table><tbody>'+pays+'</tbody></table>'+
    '<div style="margin-top:10px">พิมพ์เมื่อ '+new Date().toLocaleString('th-TH')+'</div></body></html>';
}

function getSupplierAging() {
  const today = new Date();
  const pos = getPurchaseOrders('all');
  const payments = getSupplierPayments(5000);
  const bySupplier = {};
  pos.forEach(function(po){
    if (!po.supplierId) return;
    const paid = round2_(payments.filter(function(p){ return String(p.poId||'') === String(po.id); }).reduce(function(s,p){ return s + Number(p.amount||0); },0));
    const outstanding = round2_(Number(po.totalAmount || 0) - paid);
    if (outstanding <= 0.009) return;
    const base = new Date(po.expectedDate || po.poDate || new Date());
    const ageDays = Math.max(0, Math.floor((today - base) / 86400000));
    if (!bySupplier[po.supplierId]) bySupplier[po.supplierId] = {supplierId:po.supplierId, supplierName:po.supplierName, current:0, d1_30:0, d31_60:0, d61_90:0, d90p:0, total:0, items:[]};
    const rec = bySupplier[po.supplierId];
    if (ageDays <= 0) rec.current += outstanding;
    else if (ageDays <= 30) rec.d1_30 += outstanding;
    else if (ageDays <= 60) rec.d31_60 += outstanding;
    else if (ageDays <= 90) rec.d61_90 += outstanding;
    else rec.d90p += outstanding;
    rec.total += outstanding;
    rec.items.push({poId:po.id, ageDays:ageDays, outstanding:outstanding});
  });
  return Object.keys(bySupplier).map(function(k){
    const r = bySupplier[k];
    r.current=round2_(r.current); r.d1_30=round2_(r.d1_30); r.d31_60=round2_(r.d31_60); r.d61_90=round2_(r.d61_90); r.d90p=round2_(r.d90p); r.total=round2_(r.total); return r;
  }).sort(function(a,b){ return b.total-a.total; });
}

function getDailyGrossProfit(days) {
  days = Number(days || 14);
  const start = new Date(); start.setDate(start.getDate() - days + 1); start.setHours(0,0,0,0);
  const cancelledMap = getActiveCancelledPaymentIds_();
  const orders = getRows_('orders');
  const orderMap = {}; orders.forEach(function(o){ orderMap[o.id] = o; });
  const orderCogs = {}; orders.forEach(function(o){ orderCogs[o.id] = getOrderCogs_(o.id); });
  const buckets = {};
  for (var i=0;i<days;i++) {
    const d = new Date(start); d.setDate(start.getDate()+i);
    const key = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    buckets[key] = {date:key, revenue:0, cogs:0, refunds:0, grossProfit:0, orders:0};
  }
  getRows_('payment_lines').forEach(function(pl){
    const key = String(pl.created_at || '').slice(0,10);
    if (!buckets[key]) return;
    if (cancelledMap[String(pl.payment_id)]) return;
    const order = orderMap[pl.order_id]; if (!order) return;
    const total = Number(order.total_amount || 0) || 1;
    const amount = Number(pl.amount || 0);
    buckets[key].revenue += amount;
    buckets[key].cogs += orderCogs[pl.order_id] * (amount / total);
    buckets[key].orders += 1;
  });
  getRows_('refunds').forEach(function(r){
    const key = String(r.created_at || '').slice(0,10);
    if (!buckets[key]) return;
    if (String(r.status||'approved') === 'cancelled') return;
    buckets[key].refunds += Number(r.amount || 0);
  });
  return Object.keys(buckets).sort().map(function(k){
    const b = buckets[k];
    b.revenue = round2_(b.revenue); b.cogs = round2_(b.cogs); b.refunds = round2_(b.refunds); b.grossProfit = round2_(b.revenue - b.cogs - b.refunds); b.grossMarginPct = b.revenue > 0 ? round2_(100 * b.grossProfit / b.revenue) : 0; return b;
  });
}

function getCloseShiftPrintableData(shiftId) {
  const summary = getShiftSummary(shiftId);
  const lowStock = getInventorySummary().items.filter(function(i){ return i.qtyOnHand <= i.reorderPoint; });
  const receipts = getGoodsReceipts(20).filter(function(r){ return !shiftId || String(r.receivedAt || '').slice(0,10) >= String(summary.shift.openedAt || '').slice(0,10); });
  const gross = getDailyGrossProfit(7);
  return {success:true, shift:summary.shift, summary:summary, lowStockItems:lowStock, recentReceipts:receipts, printableHtml: buildShiftPrintableHtmlV5_(summary, lowStock, receipts, gross)};
}

function buildShiftPrintableHtmlV5_(summary, lowStock, receipts, gross) {
  const aging = getSupplierAging().slice(0,10);
  const gpRows = (gross || []).map(function(r){ return '<tr><td>'+r.date+'</td><td style="text-align:right">฿'+money_(r.revenue)+'</td><td style="text-align:right">฿'+money_(r.cogs)+'</td><td style="text-align:right">฿'+money_(r.refunds)+'</td><td style="text-align:right">฿'+money_(r.grossProfit)+'</td></tr>'; }).join('');
  const agingRows = aging.map(function(r){ return '<tr><td>'+r.supplierName+'</td><td style="text-align:right">฿'+money_(r.current)+'</td><td style="text-align:right">฿'+money_(r.d1_30)+'</td><td style="text-align:right">฿'+money_(r.d31_60)+'</td><td style="text-align:right">฿'+money_(r.d61_90+r.d90p)+'</td><td style="text-align:right">฿'+money_(r.total)+'</td></tr>'; }).join('');
  return buildShiftPrintableHtml_(summary, lowStock, receipts).replace('</body></html>', '<div class="card" style="margin-top:16px"><h3>กำไรขั้นต้นรายวัน</h3><table><thead><tr><th>วันที่</th><th class="right">Revenue</th><th class="right">COGS</th><th class="right">Refund</th><th class="right">Gross Profit</th></tr></thead><tbody>'+gpRows+'</tbody></table></div><div class="card" style="margin-top:16px"><h3>Supplier Aging</h3><table><thead><tr><th>ซัพพลายเออร์</th><th class="right">Current</th><th class="right">1-30</th><th class="right">31-60</th><th class="right">61+ </th><th class="right">รวม</th></tr></thead><tbody>'+agingRows+'</tbody></table></div></body></html>');
}
