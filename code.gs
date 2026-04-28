// --- 設定: データ用スプレッドシート ---
const SPREADSHEET_ID = '1TpkECkY_JUu_g7SogBTz1yjG7PusBn_-vqKk_UJf0S4';
const SPREADSHEET_ID_ORDERS = '15Z3DsmG1K0YkYp-v09W-71EHesHoS6IOQIgVSSzTy-0'; // 注文履歴専用
const SHEET_CUSTOMER = '顧客マスタ';
const SHEET_PRODUCT = '商品マスタ';
const SHEET_CARRIER_MASTER = '配送会社マスタ';
const SHEET_STOCK = '品目倉庫別在庫'; // ←追加: 倉庫在庫シート
const SHEET_UNIT_MASTER = '単位マスタ';
const SHEET_ORDER_PARENT = '注文履歴_親';
const SHEET_ORDER_CHILD = '注文履歴_子';

const CACHE_KEY = 'preprocessed_order_data_v36';

// --- 設定: 帳票テンプレート用スプレッドシート ---
const TEMPLATE_SS_ID = '1L0h9G4gO_gQQwgSJxC3jrNdu59djvhQBSxpn3bj6x4A';
const SHEET_TEMPLATE_INVOICE = '送り状';
const SHEET_TEMPLATE_DELIVERY = '納品書';
const SHEET_TEMPLATE_INSTRUCTION = '出荷指示書';

// --- 設定: 納入先台帳用スプレッドシート ---
const DELIVERY_SS_ID = '1TpkECkY_JUu_g7SogBTz1yjG7PusBn_-vqKk_UJf0S4';
const SHEET_DELIVERY_MASTER = '納入先マスタ';

function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
      .setTitle('受注管理システム')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- ヘルパー関数 ---
function getSheetDataAsObjects_(sheetName, spreadSheetId = SPREADSHEET_ID) {
  try {
    const sheet = SpreadsheetApp.openById(spreadSheetId).getSheetByName(sheetName);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const headers = data.shift().map(h => String(h));
    return data.map(row => {
      const obj = {};
      headers.forEach((header, i) => obj[header] = row[i]);
      return obj;
    });
  } catch (e) {
    Logger.log(e.message);
    return [];
  }
}

// getValues() と getDisplayValues() を混合し、idColumns に指定した列だけ表示値（0埋め文字列）を使う
function getSheetDataWithDisplayIds_(sheetName, idColumns, spreadSheetId = SPREADSHEET_ID) {
  try {
    const sheet = SpreadsheetApp.openById(spreadSheetId).getSheetByName(sheetName);
    if (!sheet) return [];
    const range = sheet.getDataRange();
    const values = range.getValues();
    const displayValues = range.getDisplayValues();
    if (values.length < 2) return [];
    const headers = values[0].map(h => String(h));
    const idColSet = new Set(idColumns.map(col => headers.indexOf(col)).filter(i => i !== -1));
    values.shift();
    displayValues.shift();
    return values.map((row, ri) => {
      const obj = {};
      headers.forEach((header, i) => {
        obj[header] = idColSet.has(i) ? String(displayValues[ri][i]) : row[i];
      });
      return obj;
    });
  } catch (e) {
    Logger.log(e.message);
    return [];
  }
}

function getTodayString_() {
  return Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd');
}

function generateNumericOrderId_() {
  return Utilities.formatDate(new Date(), 'JST', 'yyyyMMddHHmmss');
}

function generateRandomId_() {
  return Math.random().toString(36).substr(2, 9);
}

// 日付オブジェクトを yyyyMMdd 文字列に変換
function formatDateToYMD_(date) {
  if (!date) return '';
  return Utilities.formatDate(new Date(date), 'JST', 'yyyyMMdd');
}

// yyyyMMdd文字列 / yyyy/MM/dd文字列 / Date型 を yyyy-MM-dd に変換（HTML表示用）
function parseYMDString_(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'JST', 'yyyy-MM-dd');
  }
  const s = String(val).trim();
  // yyyyMMdd形式
  if (s.length === 8 && !isNaN(s)) {
    return `${s.substring(0, 4)}-${s.substring(4, 6)}-${s.substring(6, 8)}`;
  }
  // yyyy/MM/dd形式
  if (/^\d{4}\/\d{2}\/\d{2}$/.test(s)) {
    return s.replace(/\//g, '-');
  }
  // yyyy-MM-dd形式（そのまま返す）
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    return s;
  }
  try {
    return Utilities.formatDate(new Date(s), 'JST', 'yyyy-MM-dd');
  } catch(e) {
    return '';
  }
}

function convertTaxTypeToCode_(text) {
  if (text === '外税') return 0;
  if (text === '内税') return 1;
  if (text === '非課税') return 2;
  return 0;
}

function getDailyIncrementedFilename_(baseName) {
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), 'JST', 'yyyy年MM月dd日');
  const keyDate = `DATE_${baseName.replace(/\s+/g, '')}`;
  const keyCount = `COUNT_${baseName.replace(/\s+/g, '')}`;

  const lastDate = props.getProperty(keyDate);
  let count = 1;
  if (lastDate === today) {
    const savedCount = props.getProperty(keyCount);
    if (savedCount) count = Number(savedCount) + 1;
  }

  props.setProperties({
    [keyDate]: today,
    [keyCount]: String(count)
  });

  return `${today} ${baseName} ${count}.xlsx`;
}

function getInitialDates() {
  const calendarId = 'ja.japanese#holiday@group.v.calendar.google.com';
  const calendar = CalendarApp.getCalendarById(calendarId);
  const today = new Date();
  const shipDateObj = new Date(today);
  shipDateObj.setDate(shipDateObj.getDate() + 1);
  const shipDate = Utilities.formatDate(shipDateObj, 'JST', 'yyyy-MM-dd');
  
  let deliveryDateObj = new Date(today);
  let businessDaysAdded = 0;
  while (businessDaysAdded < 2) {
    deliveryDateObj.setDate(deliveryDateObj.getDate() + 1);
    const day = deliveryDateObj.getDay();
    const isWeekend = (day === 0 || day === 6);
    const hasEvents = calendar.getEventsForDay(deliveryDateObj).length > 0;
    if (!isWeekend && !hasEvents) {
      businessDaysAdded++;
    }
  }
  const deliveryDate = Utilities.formatDate(deliveryDateObj, 'JST', 'yyyy-MM-dd');
  
  return { shipDate, deliveryDate };
}

// --- フロントエンド連携API ---

function getAllCustomers() {
  const customers = getSheetDataAsObjects_(SHEET_CUSTOMER);
  const carriers = getSheetDataAsObjects_(SHEET_CARRIER_MASTER);
  const deliveries = getSheetDataAsObjects_(SHEET_DELIVERY_MASTER, DELIVERY_SS_ID);
  const carrierNameMap = new Map();
  carriers.forEach(c => {
    if (c.配送会社名 && c.配送会社コード) {
      carrierNameMap.set(c.配送会社名, String(c.配送会社コード));
    }
  });
  const deliveryCarrierMap = new Map();
  deliveries.forEach(d => {
    const custId = String(d.取引先コード);
    const cName = d.配送会社名 || d.配送会社 || ''; 
    if (custId && cName && !deliveryCarrierMap.has(custId)) {
      deliveryCarrierMap.set(custId, cName);
    }
  });
  return customers.map(c => {
    let targetCarrierName = c.配送会社名 || ''; 
    let determinedCode = '';

    if (!targetCarrierName) {
      const custId = String(c.取引先コード);
      if (deliveryCarrierMap.has(custId)) {
        targetCarrierName = deliveryCarrierMap.get(custId);
      }
    }

    if (targetCarrierName && carrierNameMap.has(targetCarrierName)) {
      determinedCode = carrierNameMap.get(targetCarrierName);
    }

    return { 
      id: c.取引先コード, 
      name: c.取引先名1,
      kana: c.ふりがな || c.取引先名カナ1 || '',
      defaultCarrier: determinedCode, 
      address: (c.住所1 || '') + (c.住所2 || '')
    };
  }).sort((a,b) => a.name.localeCompare(b.name, 'ja'));
}

function getAllCarriers() {
  const carriers = getSheetDataAsObjects_(SHEET_CARRIER_MASTER);
  return carriers.map(c => ({
    id: c.配送会社コード,
    name: c.配送会社名,
    kana: c.配送会社名カナ || ''
  })).sort((a,b) => a.name.localeCompare(b.name, 'ja'));
}

function getAllProducts() {
   const products = getSheetDataWithDisplayIds_(SHEET_PRODUCT, ['品目コード']);
   return products.map(p => {
     // 単位と換算数量をオブジェクト化
     const unitMap = {};
     for(let i=1; i<=5; i++) {
         const u = p[`単位${i}`];
         const r = p[`換算数量${i}`];
         if (u && String(u).trim() !== '') {
             unitMap[String(u).trim()] = Number(r) || 1;
         }
     }

     return {
       id: p.品目コード,
       name: p.品目名,
       standard: p.規格 || '',
       brand: p.ブランド名 || '',
       kana: p.品目名カナ || '',
       unit: p.単位1 || '',
       unitMap: unitMap,
       zaikoUnitName: p.在庫単位名 || '',
       tradeUnitCode: String(p.取引単位コード ?? ''),
       tradeUnitName: p.取引単位名 || p.単位1 || '',
       price: Number(p.標準売上単価) || 0,
       caseQty: Number(p.入数) || 1,
       gosu: Number(p.合数) || 0,
       tempType: (p.在庫管理区分 === 'する' || p.在庫管理区分 === '冷凍') ? 'frozen' : 'chilled'
     };
   });
}

function getAllUnitMaster() {
  const units = getSheetDataAsObjects_(SHEET_UNIT_MASTER);
  return units.map(u => ({
    code: String(u['単位コード'] ?? u[Object.keys(u)[0]] ?? ''),
    name: String(u['単位名'] ?? u[Object.keys(u)[1]] ?? '')
  }));
}

function getAllDeliveryDestinations() {
  const destinations = getSheetDataAsObjects_(SHEET_DELIVERY_MASTER, DELIVERY_SS_ID);
  return destinations.map(d => ({
    id: d.納入先コード,
    name: d.納入先名1,
    customerId: String(d.取引先コード) 
  }));
}

function getAppInitialData() {
  return {
    dashboard: getDashboardData(),
    units: getAllUnitMaster(),
    customers: getAllCustomers(),
    products: getAllProducts(),
    carriers: getAllCarriers(),
    destinations: getAllDeliveryDestinations(),
    dates: getInitialDates(),
    orders: getRecentOrders()
  };
}

function getDashboardData() {
  const todayStr = getTodayString_();
  const parents = getSheetDataAsObjects_(SHEET_ORDER_PARENT, SPREADSHEET_ID_ORDERS);
  const children = getSheetDataAsObjects_(SHEET_ORDER_CHILD, SPREADSHEET_ID_ORDERS);
  const todaysParents = parents.filter(p => {
    if(!p.注文日) return false;
    const pDateStr = parseYMDString_(p.注文日); 
    return pDateStr === todayStr;
  });
  const total = todaysParents.length;
  const confirmed = todaysParents.filter(p => p.ステータス === '確定済み').length;
  const pending = total - confirmed;
  const todayShipChildren = children.filter(c => c.出荷日 && parseYMDString_(c.出荷日) === todayStr);
  const todayShipIds = new Set(todayShipChildren.map(c => String(c.注文ID)));
  const shipCount = todayShipIds.size;
  const sales = todayShipChildren.reduce((sum, c) => sum + (Number(c.金額) || 0), 0);

  const alertCustomers = [];
  return {
    total, confirmed, pending, sales, shipCount, alertCustomers
  };
}

function getRecentOrders() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEY);
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }

  const today = new Date();
  const dateStrings = [];
  for (let i = 0; i < 3; i++) {
    const d = new Date();
    d.setDate(today.getDate() + i);
    dateStrings.push(Utilities.formatDate(d, 'JST', 'yyyy-MM-dd'));
  }
  
  const parents = getSheetDataAsObjects_(SHEET_ORDER_PARENT, SPREADSHEET_ID_ORDERS);
  const children = getSheetDataAsObjects_(SHEET_ORDER_CHILD, SPREADSHEET_ID_ORDERS);
  const products = getSheetDataWithDisplayIds_(SHEET_PRODUCT, ['品目コード']);
  
  const carriers = getSheetDataAsObjects_(SHEET_CARRIER_MASTER);
  const carrierMap = new Map(carriers.map(c => [String(c.配送会社コード), c.配送会社名]));
  const deliveries = getSheetDataAsObjects_(SHEET_DELIVERY_MASTER, DELIVERY_SS_ID);
  const deliveryMap = new Map(deliveries.map(d => [String(d.納入先コード), d.納入先名1]));
  const productInfoMap = new Map(products.map(p => [
    String(p.品目コード), 
    { 
      name: p.品目名, 
      std: p.規格 || '', 
      unit: p.単位1 || '', 
      price: Number(p.標準売上単価) || 0, 
      caseQty: Number(p.入数) || 1,
      gosu: Number(p.合数) || 0
    }
  ]));
  const parentIds = new Set(parents.map(p => String(p.注文ID)));
  const relatedChildren = children.filter(c => parentIds.has(String(c.注文ID)));

  const childrenMap = {};
  relatedChildren.forEach(c => {
    const pid = String(c.注文ID);
    if (!childrenMap[pid]) childrenMap[pid] = [];
    childrenMap[pid].push(c);
  });
  const result = [];

  parents.forEach(parent => {
    const pid = String(parent.注文ID);
    const myChildren = childrenMap[pid] || [];
    
    const parentOrderDateStr = parseYMDString_(parent.注文日);

    const cCode = String(parent.運送会社 || '');
    const dCode = String(parent.納入先 || '');
    
    const cName = carrierMap.get(cCode) || cCode;
    const dName = deliveryMap.get(dCode) || dCode;
    const custId = parent.取引先コード || parent.得意先コード || parent.顧客ID;

    let hasRecentItem = false;

    const items = myChildren.map(c => {
      const pId = c.品目コード || c.商品ID;
      const pInfo = productInfoMap.get(String(pId)) || {};
      
      const rawDate = c['納品日'] || parent['納品日'] || parentOrderDateStr;
      const itemDateStr = parseYMDString_(rawDate);

      const rawShipDate = c['出荷日'] || parent['出荷日'] || '';
      const shipDateStr = parseYMDString_(rawShipDate);

      if (dateStrings.includes(shipDateStr)) {
        hasRecentItem = true;
      }

      const qty = Number(c.数量) || 0;
      let unitPrice = Number(c.単価) || pInfo.price || 0;
      let amount = Number(c.金額) || (unitPrice * qty);

      return {
        orderId: pid, 
        detailId: c.内訳ID, 
        productId: pId,
        productName: pInfo.name || '不明',
        standard: c.規格 || pInfo.std, 
        unit: c.単位コード || pInfo.unit, 
        unitPrice: unitPrice, 
        caseQty: Number(c.入数) || pInfo.caseQty,
        gosu: Number(c.合数) || pInfo.gosu,
        caseCount: Number(c.箱数),
        quantity: qty,
        price: amount,
        awaseId: c.合わせ || '', 
        shipDate: shipDateStr,
        deliveryDate: itemDateStr, 
        remarks: c.摘要 || c.備考,
        orderQty: c.注文時数量 || qty,
        orderUnit: c.注文時単位 || c.単位コード || pInfo.unit,
        kg: c.kg || 0,
        shippingQty: c.出荷数 || 0,
        shipUnit: c.出荷単位 || '',
        warehouseCode: c.倉庫コード || '', // ←追加
        lotNumber: c.ロット番号 || ''      // ←追加
      };
    });
    if (hasRecentItem) {
      const total = items.reduce((sum, item) => sum + item.price, 0);
      const custName = parent.取引先名 || parent.取引先名1 || parent.得意先名;
      result.push({
        customerId: custId,
        customerName: custName,
        deliveryDestinationCode: dCode, 
        deliveryDestination: dName,     
        carrierCode: cCode,             
        carrierName: cName,             
        orderDate: parentOrderDateStr, 
        status: parent.ステータス || '未確定',
        isPrinted: (!!parent.納品書出力フラグ || !!parent.送り状出力フラグ),
        isCsvExported: !!parent['CSV出力フラグ'],
        orderId: pid, 
        items: items,
        totalPrice: total
      });
    }
  });

  cache.put(CACHE_KEY, JSON.stringify(result), 300);
  return result;
}

function deleteOrder(detailId, orderId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID_ORDERS);
  const childSheet = ss.getSheetByName(SHEET_ORDER_CHILD);
  const data = childSheet.getDataRange().getValues();
  
  const detailIdCol = data[0].indexOf('内訳ID');
  let deleted = false;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][detailIdCol]) === String(detailId)) {
      childSheet.deleteRow(i + 1);
      deleted = true;
      break;
    }
  }
  
  if(deleted) {
    CacheService.getScriptCache().remove(CACHE_KEY);
    checkAndDeleteEmptyParent_(orderId);
    return { status: 'success' };
  } else {
    throw new Error('削除対象の明細が見つかりませんでした');
  }
}

function checkAndDeleteEmptyParent_(orderId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID_ORDERS);
  const childSheet = ss.getSheetByName(SHEET_ORDER_CHILD);
  const parentSheet = ss.getSheetByName(SHEET_ORDER_PARENT);
  
  const children = childSheet.getDataRange().getValues();
  const oidColC = children[0].indexOf('注文ID');
  const hasChild = children.some((row, i) => i > 0 && String(row[oidColC]) === String(orderId));
  if (!hasChild) {
    const parents = parentSheet.getDataRange().getValues();
    const oidColP = parents[0].indexOf('注文ID');
    for(let i=1; i<parents.length; i++) {
      if(String(parents[i][oidColP]) === String(orderId)) {
        parentSheet.deleteRow(i+1);
        break;
      }
    }
  }
}

function updateAndConfirmOrder(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID_ORDERS);
  const parentSheet = ss.getSheetByName(SHEET_ORDER_PARENT);
  const childSheet = ss.getSheetByName(SHEET_ORDER_CHILD);

  let parentCustId = '';

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const pData = parentSheet.getDataRange().getValues();
    const pHeaders = pData[0];
    const pIdx = {
      id: pHeaders.indexOf('注文ID'),
      custId: pHeaders.findIndex(h => h === '取引先コード' || h === '得意先コード'),
      custName: pHeaders.findIndex(h => h === '取引先名' || h === '取引先名1' || h === '得意先名'),
      delivery: pHeaders.indexOf('納入先'), 
      carrier: pHeaders.indexOf('運送会社'),
      delDate: pHeaders.indexOf('納品日'), 
      status: pHeaders.indexOf('ステータス')
    };
    let targetOrderId = null;
    if(data.items.length > 0) {
      targetOrderId = data.items[0].orderId;
    }

    if (!targetOrderId) throw new Error('注文IDが特定できません');

    let parentRowIndex = -1;
    for(let i=1; i<pData.length; i++) {
      if(String(pData[i][pIdx.id]) === String(targetOrderId)) {
        parentRowIndex = i + 1;
        break;
      }
    }

    let parentDelivery = '';
    const parentDeliveryDateStr = (data.items.length > 0 && data.items[0].deliveryDate) ? formatDateToYMD_(data.items[0].deliveryDate) : '';
    if(parentRowIndex !== -1) {
      parentCustId = pData[parentRowIndex - 1][pIdx.custId];
      parentDelivery = pData[parentRowIndex - 1][pIdx.delivery];
      parentSheet.getRange(parentRowIndex, pIdx.status + 1).setValue('確定済み');
      
      if (pIdx.delDate !== -1 && parentDeliveryDateStr) {
        parentSheet.getRange(parentRowIndex, pIdx.delDate + 1).setValue(parentDeliveryDateStr);
      }
      
      if (data.newDeliveryDestination && pIdx.delivery !== -1) {
        parentSheet.getRange(parentRowIndex, pIdx.delivery + 1).setValue(data.newDeliveryDestination);
        parentDelivery = data.newDeliveryDestination; 
      }
      if (data.newCarrierName && pIdx.carrier !== -1) {
        parentSheet.getRange(parentRowIndex, pIdx.carrier + 1).setValue(data.newCarrierName);
      }
      if (data.newCustomerId) {
         parentSheet.getRange(parentRowIndex, pIdx.custId + 1).setValue(data.newCustomerId);
         parentSheet.getRange(parentRowIndex, pIdx.custName + 1).setValue(data.newCustomerName);
         parentCustId = data.newCustomerId; 
      }
    } else {
      throw new Error('親注文データが見つかりません');
    }

    const products = getSheetDataWithDisplayIds_(SHEET_PRODUCT, ['品目コード']);
    const productMap = new Map(products.map(p => [
        String(p.品目コード), 
        { 
            price: Number(p.標準売上単価) || 0,
            caseQty: Number(p.入数) || 1,
            gosu: Number(p.合数) || 0,
            taxRate: p.新税率 || 0,
            taxType: convertTaxTypeToCode_(p.消費税課税区分)
        }
    ]));
    const cData = childSheet.getDataRange().getValues();
    const cHeaders = cData[0];
    
    // 既存行の更新用インデックス
    const cIdx = {
      detailId: cHeaders.indexOf('内訳ID'),
      orderId: cHeaders.indexOf('注文ID'),
      shipDate: cHeaders.indexOf('出荷日'),
      delDate: cHeaders.indexOf('納品日'),
      prodId: cHeaders.indexOf('品目コード'),
      unit: cHeaders.indexOf('単位コード'), 
      caseQty: cHeaders.indexOf('入数'),
      gosu: cHeaders.indexOf('合数'),
      caseCount: cHeaders.indexOf('箱数'),
      qty: cHeaders.indexOf('数量'),
      price: cHeaders.indexOf('単価'),
      amount: cHeaders.indexOf('金額'),
      taxRate: cHeaders.indexOf('消費税率'),
      taxType: cHeaders.indexOf('課税区分'),
      remarks: cHeaders.indexOf('摘要'),
      standard: cHeaders.indexOf('規格'),
      awase: cHeaders.indexOf('合わせ'),
      customer: cHeaders.indexOf('得意先名'), 
      delivery: cHeaders.indexOf('納入先'),
      orderUnit: cHeaders.indexOf('注文時単位'),
      orderQty: cHeaders.indexOf('注文時数量'),
      kg: cHeaders.indexOf('kg'),
      shippingQty: cHeaders.indexOf('出荷数'),
      shipUnit: cHeaders.indexOf('出荷単位'),
      warehouseCode: cHeaders.indexOf('倉庫コード'), // ←追加
      lotNumber: cHeaders.indexOf('ロット番号')      // ←追加
    };
    const detailIdMap = new Map();
    for(let i=1; i<cData.length; i++) {
      if(String(cData[i][cIdx.orderId]) === String(targetOrderId)) {
        detailIdMap.set(String(cData[i][cIdx.detailId]), i + 1);
      }
    }

    const newRows = [];
    data.items.forEach((item, index) => {
      const isNew = String(item.orderId).startsWith('NEW-') || !item.detailId; 
      const pInfo = productMap.get(String(item.productId)) || {price:0, caseQty:1, gosu:0, taxRate:0, taxType:0};
      
      const itemShipDateObj = item.shipDate ? new Date(item.shipDate) : ''; 
      const itemDeliveryDateObj = item.deliveryDate ? new Date(item.deliveryDate) : '';
      
      const finalCaseCount = Number(item.caseCount) || 0;
      const finalQuantity = Number(item.quantity) || 0; // 請求用数量

      const unitPrice = (Number(item.price) > 0) ? Number(item.price) : pInfo.price;
      const amount = finalQuantity * unitPrice;

      const remarks = item.newRemarks || item.remarks;
      const standard = item.standard || ''; 
      const unit = item.unit || '';
      const awaseVal = item.awaseId || ''; 

      if (!isNew && detailIdMap.has(String(item.detailId))) {
        const rowNum = detailIdMap.get(String(item.detailId));
        // コード系フィールドは先頭「'」でテキスト強制（先頭ゼロ落ち防止）
        if(cIdx.prodId !== -1) childSheet.getRange(rowNum, cIdx.prodId + 1).setValue("'" + String(item.productId || ''));
        if(cIdx.unit !== -1) childSheet.getRange(rowNum, cIdx.unit + 1).setValue("'" + String(unit || ''));
        if(cIdx.caseQty !== -1) childSheet.getRange(rowNum, cIdx.caseQty + 1).setValue(pInfo.caseQty);
        if(cIdx.gosu !== -1) childSheet.getRange(rowNum, cIdx.gosu + 1).setValue(pInfo.gosu);
        if(cIdx.caseCount !== -1) childSheet.getRange(rowNum, cIdx.caseCount + 1).setValue(finalCaseCount);
        if(cIdx.qty !== -1) childSheet.getRange(rowNum, cIdx.qty + 1).setValue(finalQuantity);
        if(cIdx.price !== -1) childSheet.getRange(rowNum, cIdx.price + 1).setValue(unitPrice);
        if(cIdx.amount !== -1) childSheet.getRange(rowNum, cIdx.amount + 1).setValue(amount);
        if(cIdx.taxRate !== -1) childSheet.getRange(rowNum, cIdx.taxRate + 1).setValue(pInfo.taxRate);
        if(cIdx.taxType !== -1) childSheet.getRange(rowNum, cIdx.taxType + 1).setValue(pInfo.taxType);
        if(cIdx.shipDate !== -1) childSheet.getRange(rowNum, cIdx.shipDate + 1).setValue(itemShipDateObj);
        if(cIdx.delDate !== -1) childSheet.getRange(rowNum, cIdx.delDate + 1).setValue(itemDeliveryDateObj);
        if(cIdx.remarks !== -1) childSheet.getRange(rowNum, cIdx.remarks + 1).setValue(remarks);
        if(cIdx.standard !== -1) childSheet.getRange(rowNum, cIdx.standard + 1).setValue("'" + String(standard || ''));
        if(cIdx.awase !== -1) childSheet.getRange(rowNum, cIdx.awase + 1).setValue(awaseVal);
        if(cIdx.customer !== -1) childSheet.getRange(rowNum, cIdx.customer + 1).setValue("'" + String(parentCustId || ''));
        if(cIdx.delivery !== -1) childSheet.getRange(rowNum, cIdx.delivery + 1).setValue("'" + String(parentDelivery || ''));
        if(cIdx.orderUnit !== -1) childSheet.getRange(rowNum, cIdx.orderUnit + 1).setValue("'" + String(item.orderUnit || ''));
        if(cIdx.orderQty !== -1) childSheet.getRange(rowNum, cIdx.orderQty + 1).setValue(item.orderQty || 0);
        if(cIdx.kg !== -1) childSheet.getRange(rowNum, cIdx.kg + 1).setValue(item.kg || 0);
        if(cIdx.shippingQty !== -1) childSheet.getRange(rowNum, cIdx.shippingQty + 1).setValue(item.shippingQty || 0);
        if(cIdx.shipUnit !== -1) childSheet.getRange(rowNum, cIdx.shipUnit + 1).setValue(item.shipUnit || '');
        if(cIdx.warehouseCode !== -1) childSheet.getRange(rowNum, cIdx.warehouseCode + 1).setValue("'" + String(item.warehouseCode || ''));
        if(cIdx.lotNumber !== -1) childSheet.getRange(rowNum, cIdx.lotNumber + 1).setValue("'" + String(item.lotNumber || ''));
      } else {
        const newDetailId = generateRandomId_();
        // コード系フィールドは先頭「'」でテキスト強制（先頭ゼロ落ち防止）
        const row = [
          newDetailId,                               // 1. 内訳ID
          targetOrderId,                             // 2. 注文ID
          '',                                        // 3. 伝票番号
          itemShipDateObj,                           // 4. 出荷日
          itemDeliveryDateObj,                       // 5. 納品日
          index + 1,                                 // 6. 行番号
          "'" + String(item.productId    || ''),     // 7. 品目コード
          "'" + String(unit              || ''),     // 8. 単位コード
          pInfo.caseQty,                             // 9. 入数
          pInfo.gosu,                                // 10. 合数
          finalCaseCount,                            // 11. 箱数
          "'" + String(item.warehouseCode || ''),    // 12. 倉庫コード
          "'" + String(item.lotNumber     || ''),    // 13. ロット番号
          finalQuantity,                             // 14. 数量
          unitPrice,                                 // 15. 単価
          amount,                                    // 16. 金額
          pInfo.taxRate,                             // 17. 消費税率
          pInfo.taxType,                             // 18. 課税区分
          remarks,                                   // 19. 摘要
          "'" + String(standard           || ''),    // 20. 規格
          awaseVal,                                  // 21. 合わせ
          "'" + String(parentCustId       || ''),    // 22. 得意先名
          "'" + String(parentDelivery     || ''),    // 23. 納入先
          "'" + String(item.orderUnit     || ''),    // 24. 注文時単位
          item.orderQty    || 0,                     // 25. 注文時数量
          item.kg          || 0,                     // 26. kg
          item.shippingQty || 0,                     // 27. 出荷数
          item.shipUnit    || ''                      // 28. 出荷単位
        ];
        newRows.push(row);
      }
    });
    if(newRows.length > 0) {
      const startRow = childSheet.getLastRow() + 1;
      childSheet.getRange(startRow, 1, newRows.length, newRows[0].length).setValues(newRows);
    }

    CacheService.getScriptCache().remove(CACHE_KEY);
  } finally {
    lock.releaseLock();
  }

  return { status: 'success' };
}

function exportDataAsCsv(startDateStr, endDateStr) { return null; }

function createSplitRows_(items) {
  const finalRows = [];
  const fractionGroups = {};
  items.forEach(item => {
    const originalQty = Number(item.quantity) || 0;
    const originalAmount = Number(item.price) || 0;
    const integerPart = Math.floor(originalQty);     
    const decimalPart = originalQty - integerPart;   

    let unitPrice = 0;
    if (originalQty > 0) {
      unitPrice = originalAmount / originalQty;
    }

    if (integerPart > 0) {
      const intAmount = Math.round(unitPrice * integerPart);
      finalRows.push({
        ...item,
        productName: item.productName,
        quantity: integerPart,          
        price: intAmount,               
        displayBox: integerPart,        
        isFraction: false
      });
    }

    if (decimalPart > 0) {
      const decAmount = originalAmount - (Math.round(unitPrice * integerPart));
      const fracRow = {
        ...item,
        productName: item.productName,
        quantity: decimalPart,          
        price: decAmount,               
        displayBox: '',       
        isFraction: true
      };
      if (item.awaseId) {
        if (!fractionGroups[item.awaseId]) fractionGroups[item.awaseId] = [];
        fractionGroups[item.awaseId].push(fracRow);
      } else {
        fracRow.displayBox = 1; 
        finalRows.push(fracRow);
      }
    }
  });

  Object.keys(fractionGroups).forEach(awaseId => {
    const group = fractionGroups[awaseId];
    group.forEach((row, index) => {
      if (index === 0) {
        row.displayBox = 1; 
      } else {
        row.displayBox = ''; 
      }
      finalRows.push(row);
    });
  });
  return finalRows;
}

function generateBatchPdfs(targetList) {
  const todaysOrders = getRecentOrders(); 
  const customers = getSheetDataAsObjects_(SHEET_CUSTOMER);
  const customerMap = new Map(customers.map(c => [String(c.取引先コード), c])); 

  const rawTargetOrders = [];
  targetList.forEach(req => {
    const matchedOrder = todaysOrders.find(o => String(o.customerId) === String(req.customerId));
    if (matchedOrder) {
      const cInfo = customerMap.get(String(matchedOrder.customerId));
      if (cInfo) {
          matchedOrder.customerAddress = (cInfo.住所1 || '') + (cInfo.住所2 || ''); 
      }
      rawTargetOrders.push(matchedOrder);
    }
  });
  if (rawTargetOrders.length === 0) throw new Error('出力対象のデータが見つかりませんでした。');

  const groupedShipments = {};
  rawTargetOrders.forEach(order => {
    order.items.forEach(item => {
      const dDate = item.deliveryDate ? item.deliveryDate : order.orderDate;
      const key = `${order.customerId}_${dDate}`;

      if (!groupedShipments[key]) {
        groupedShipments[key] = {
          customerId: order.customerId,
          customerName: order.customerName,
          customerAddress: order.customerAddress,
          deliveryDate: dDate,
          shipDate: item.shipDate || '',
          deliveryDestination: order.deliveryDestination,
          deliveryDestinationCode: order.deliveryDestinationCode,
          carrierName: order.carrierName,
          orderIds: new Set(),
          items: []
        };
      }
      groupedShipments[key].orderIds.add(order.orderId);
      groupedShipments[key].items.push(item);
    });
  });
  const templateFile = DriveApp.getFileById(TEMPLATE_SS_ID);
  const tempFolder = DriveApp.getRootFolder(); 
  const newFile = templateFile.makeCopy('TEMP_PRINT_' + new Date().getTime(), tempFolder);
  const newSs = SpreadsheetApp.open(newFile);
  try {
    const templateInvoice = newSs.getSheetByName(SHEET_TEMPLATE_INVOICE);
    const templateDelivery = newSs.getSheetByName(SHEET_TEMPLATE_DELIVERY);

    if (!templateInvoice || !templateDelivery) throw new Error('テンプレートシートが見つかりません');
    const generatedSheetIds = [];
    const shipmentsArray = Object.values(groupedShipments);
    const processedOrderIds = [];

    // 単位マスタ読み込み（単位コード→単位名変換用）
    const unitMasterRows = getSheetDataAsObjects_(SHEET_UNIT_MASTER);
    const unitCodeToName = new Map(unitMasterRows.map(u => [
      String(u['単位コード'] ?? u[Object.keys(u)[0]] ?? ''),
      String(u['単位名']   ?? u[Object.keys(u)[1]] ?? '')
    ]));

    const ssMain = SpreadsheetApp.openById(SPREADSHEET_ID_ORDERS);
    const childSheetMain = ssMain.getSheetByName(SHEET_ORDER_CHILD);
    const cData = childSheetMain.getDataRange().getValues();
    const slipNoIdx = cData[0].indexOf('伝票番号');
    
    let maxSlipNo = 0;
    if (slipNoIdx !== -1) {
      for (let i = 1; i < cData.length; i++) {
        const val = Number(cData[i][slipNoIdx]);
        if (!isNaN(val) && val > maxSlipNo) {
          maxSlipNo = val;
        }
      }
    }

    let slipCounter = maxSlipNo + 1;
    shipmentsArray.forEach(shipment => {
      const slipNo = String(slipCounter).padStart(9, '0');
      shipment.slipNumber = slipNo;
      shipment.items.forEach(item => item.slipNumber = slipNo);
      slipCounter++;
    });
    const INV_MAX_ROWS = 14;
    const DEL_MAX_ROWS = 10;

    shipmentsArray.forEach((shipment, index) => {
      shipment.orderIds.forEach(id => processedOrderIds.push(id));

      const splitItems = createSplitRows_(shipment.items);

      // 送り状: 14行ずつページ分割
      const invChunks = [];
      for (let i = 0; i < splitItems.length; i += INV_MAX_ROWS) {
        invChunks.push(splitItems.slice(i, i + INV_MAX_ROWS));
      }
      if (invChunks.length === 0) invChunks.push([]);

      // 納品書: 10行ずつページ分割
      const delChunks = [];
      for (let i = 0; i < splitItems.length; i += DEL_MAX_ROWS) {
        delChunks.push(splitItems.slice(i, i + DEL_MAX_ROWS));
      }
      if (delChunks.length === 0) delChunks.push([]);

      if (templateInvoice) {
        const invTotal = invChunks.length;
        invChunks.forEach((chunk, pageIdx) => {
          const sheetInv = templateInvoice.copyTo(newSs).setName(`Inv_${index}_${pageIdx}`);
          fillInvoiceSheet_(sheetInv, { ...shipment, items: chunk }, unitCodeToName, pageIdx + 1, invTotal);
          generatedSheetIds.push(sheetInv.getSheetId());
        });
      }

      if (templateDelivery) {
        delChunks.forEach((chunk, pageIdx) => {
          const sheetDel = templateDelivery.copyTo(newSs).setName(`Del_${index}_${pageIdx}`);
          fillDeliverySheet_(sheetDel, { ...shipment, items: chunk });
          generatedSheetIds.push(sheetDel.getSheetId());
        });
      }
    });
    const allSheets = newSs.getSheets();
    allSheets.forEach(sheet => {
      if (!generatedSheetIds.includes(sheet.getSheetId())) {
        try { newSs.deleteSheet(sheet); } catch(e) {}
      }
    });
    SpreadsheetApp.flush();

    const filename = getDailyIncrementedFilename_('送り状＆納品書');
    const xlsxBlob = getBlobFromWholeSs_(newSs, filename);
    const base64 = Utilities.base64Encode(xlsxBlob.getBytes());

    const uniqueOrderIds = [...new Set(processedOrderIds)];
    const slipUpdates = [];
    shipmentsArray.forEach(shipment => {
      shipment.items.forEach(item => {
        if(item.detailId) slipUpdates.push({ detailId: item.detailId, slipNumber: shipment.slipNumber });
      });
    });
    updateSlipNumbersAndFlags_(slipUpdates, uniqueOrderIds);

    return {
      status: 'success',
      base64: base64,
      filename: xlsxBlob.getName(),
      message: `${rawTargetOrders.length}件の注文から、${shipmentsArray.length}件の出荷伝票（計${generatedSheetIds.length}枚）を生成しました。`
    };
  } catch(e) {
    throw e;
  } finally {
    newFile.setTrashed(true);
  }
}

function updateSlipNumbersAndFlags_(slipUpdates, orderIds) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID_ORDERS);
  
  const parentSheet = ss.getSheetByName(SHEET_ORDER_PARENT);
  const pData = parentSheet.getDataRange().getValues();
  const idIdx = pData[0].indexOf('注文ID');
  const deliveryFlagIdx = pData[0].indexOf('納品書出力フラグ');
  const invoiceFlagIdx = pData[0].indexOf('送り状出力フラグ');
  const now = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss');
  const targetSet = new Set(orderIds.map(String));
  if (idIdx !== -1) {
    for (let i = 1; i < pData.length; i++) {
      if (targetSet.has(String(pData[i][idIdx]))) {
        if (deliveryFlagIdx !== -1) parentSheet.getRange(i + 1, deliveryFlagIdx + 1).setValue(now);
        if (invoiceFlagIdx !== -1) parentSheet.getRange(i + 1, invoiceFlagIdx + 1).setValue(now);
      }
    }
  }

  const childSheet = ss.getSheetByName(SHEET_ORDER_CHILD);
  const cData = childSheet.getDataRange().getValues();
  const detailIdIdx = cData[0].indexOf('内訳ID');
  const slipNoIdx = cData[0].indexOf('伝票番号');
  
  if (detailIdIdx !== -1 && slipNoIdx !== -1) {
    const slipMap = new Map();
    slipUpdates.forEach(u => slipMap.set(String(u.detailId), u.slipNumber));
    
    for (let i = 1; i < cData.length; i++) {
      const dId = String(cData[i][detailIdIdx]);
      if (slipMap.has(dId)) {
        childSheet.getRange(i + 1, slipNoIdx + 1).setValue(`'${slipMap.get(dId)}`);
      }
    }
  }
}

function fillInvoiceSheet_(sheet, shipment, unitCodeToName, currentPage, totalPages) {
  sheet.getRange('A7').setValue(shipment.customerAddress || '');
  sheet.getRange('A9').setValue(shipment.customerName + ' 御中');

  sheet.getRange('D3').setValue(shipment.shipDate || getTodayString_());
  sheet.getRange('D4').setValue(shipment.deliveryDate);
  sheet.getRange('I1').setValue(shipment.slipNumber || '');
  sheet.getRange('I2').setValue(`${currentPage || 1}   /   ${totalPages || 1}`);
  sheet.getRange('B28').setValue(shipment.carrierName || '');

  // 明細: A=品目名, B=規格, C=単価, D=入数, E=合数, F=箱数, G=数量, H=単位名, I=摘要/備考
  sheet.getRange(13, 1, 14, 9).clearContent();
  const rows = shipment.items.map(item => {
    const unitName = (unitCodeToName && unitCodeToName.get(String(item.unit || ''))) || item.unit || '';
    return [
      item.productName,
      item.standard || '-',
      '',
      item.caseQty,
      item.gosu || 0,
      item.displayBox,
      item.quantity,
      unitName,
      item.remarks || ''
    ];
  });
  if (rows.length > 0) {
    sheet.getRange(13, 1, rows.length, 9).setValues(rows);
  }

  let totalBox = 0;
  shipment.items.forEach(item => {
    if (typeof item.displayBox === 'number') {
      totalBox += item.displayBox;
    }
  });
  sheet.getRange('E27').setValue(totalBox);
}

function fillDeliverySheet_(sheet, shipment) {
  sheet.getRange('A4').setValue(shipment.customerAddress || ''); 
  sheet.getRange('C4').setValue(shipment.deliveryDate);
  sheet.getRange('A6').setValue(shipment.customerName + ' 御中');
  
  sheet.getRange('A9').setValue("お客様コード " + shipment.customerId);
  sheet.getRange('H2').setValue(shipment.slipNumber || ''); 
  
  const lastRow = sheet.getLastRow();
  if (lastRow >= 11) {
    sheet.getRange(11, 1, lastRow - 10, 8).clearContent();
  }

  const rows = shipment.items.map(item => [
    `${item.productName} / ${item.standard || '-'}`,
    '',
    `${item.caseQty}/${item.gosu || 0}`,
    item.displayBox,
    item.quantity, 
    item.unitPrice, 
    item.price,
    item.remarks
  ]);
  if (rows.length > 0) {
    sheet.getRange(11, 1, rows.length, 8).setValues(rows);
  }

  let totalCase = 0;
  let totalQty = 0;
  let totalAmount = 0;

  shipment.items.forEach(item => {
    totalQty += (Number(item.quantity) || 0);
    totalAmount += (Number(item.price) || 0);
    
    if (typeof item.displayBox === 'number') {
      totalCase += item.displayBox;
    }
  });
  const tax = Math.floor(totalAmount * 0.1); 
  const totalIncTax = totalAmount + tax;

  sheet.getRange('A21').setValue("納入先");
  sheet.getRange('B21').setValue(shipment.deliveryDestinationCode || shipment.deliveryDestination || '');

  sheet.getRange('D21').setValue(totalCase); 
  sheet.getRange('E21').setValue(totalQty);
  sheet.getRange('F21').setValue("合計:" + totalAmount.toLocaleString()); 
  sheet.getRange('G21').setValue("消費税:" + tax.toLocaleString()); 
  sheet.getRange('H21').setValue("金額:" + totalIncTax.toLocaleString()); 
}

function getBlobFromWholeSs_(ss, filename) {
  const url = ss.getUrl().replace(/\/edit.*$/, '') + '/export?format=xlsx';
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token }
  });
  return response.getBlob().setName(filename);
}

function importDeliveryMasterCsv(csvContent) {
  const ss = SpreadsheetApp.openById(DELIVERY_SS_ID);
  let sheet = ss.getSheetByName(SHEET_DELIVERY_MASTER);
  if (!sheet) sheet = ss.insertSheet(SHEET_DELIVERY_MASTER);
  let csvData;
  try { csvData = Utilities.parseCsv(csvContent); } catch (e) { throw new Error('CSV読込失敗');
  }
  if (csvData.length === 0) return { status: 'success', message: 'データなし' };

  const lastRow = sheet.getLastRow();
  let existingValues = [];
  if (lastRow > 1) existingValues = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const codeMap = new Map();
  existingValues.forEach((row, index) => { if (String(row[0])) codeMap.set(String(row[0]), index); });

  const updates = [];
  const appends = [];
  let startIndex = (csvData.length > 0 && (csvData[0][0] === '納入先コード' || csvData[0][0] === 'コード')) ?
1 : 0;

  for (let i = startIndex; i < csvData.length; i++) {
    const row = csvData[i].slice(0, 6);
    while (row.length < 6) row.push('');
    const code = String(row[0]);
    if (!code) continue;
    if (codeMap.has(code)) {
      updates.push({ rowNum: codeMap.get(code) + 2, data: row });
    } else {
      appends.push(row);
    }
  }
  updates.forEach(update => sheet.getRange(update.rowNum, 1, 1, 6).setValues([update.data]));
  if (appends.length > 0) sheet.getRange(lastRow + 1, 1, appends.length, 6).setValues(appends);

  return { status: 'success', message: `完了 更新:${updates.length} 新規:${appends.length}` };
}

function generateShippingInstructions(targetList) {
  return {
    status: 'success',
    url: 'https://docs.google.com/spreadsheets/d/1GnhQRHvBZ7yL-cgHeaCNDCt8OtM-m09zI060vxw03K8/edit?pli=1&gid=423721715#gid=423721715',
    message: '出荷指示書のシートを開きます。'
  };
}

function generateFrozenShippingInstructions(targetList) {
  return {
    status: 'success',
    url: 'https://docs.google.com/spreadsheets/d/144Vdqlf92eh_jUXyNyvuc0ognOL5rXuwLye1xo0g1Kk/edit?gid=0#gid=0',
    message: '冷凍品出荷指示書のシートを開きます。'
  };
}

// --- 全ロット在庫一覧（フロント連携API） ---
function getAllStocksWithProduct() {
  const stocks = getSheetDataAsObjects_(SHEET_STOCK);
  const products = getSheetDataWithDisplayIds_(SHEET_PRODUCT, ['品目コード']);

  const productMap = new Map(products.map(p => [String(p.品目コード), p]));

  return stocks.map(s => {
    const p = productMap.get(String(s.品目コード)) || {};

    let expDate = '';
    if (s.賞味期限) {
      expDate = (s.賞味期限 instanceof Date)
        ? Utilities.formatDate(s.賞味期限, 'JST', 'yyyy-MM-dd')
        : parseYMDString_(String(s.賞味期限));
    }

    let mfgDate = '';
    if (s.製造日) {
      mfgDate = (s.製造日 instanceof Date)
        ? Utilities.formatDate(s.製造日, 'JST', 'yyyy-MM-dd')
        : parseYMDString_(String(s.製造日));
    }

    const tempType = (p.在庫管理区分 === 'する' || p.在庫管理区分 === '冷凍') ? 'frozen' : 'chilled';

    return {
      productId: String(s.品目コード || ''),
      productName: p.品目名 || '',
      kana: p.品目名カナ || '',
      brand: p.ブランド名 || '',
      standard: p.規格 || '',
      tempType: tempType,
      warehouseCode: String(s.倉庫コード || ''),
      warehouseName: String(s.倉庫名 || ''),
      lotNumber: String(s.ロット番号 || ''),
      expirationDate: expDate,
      manufactureDate: mfgDate,
      stockQty: Number(s.在庫数量) || 0,
      unit: String(s.単位 || '')
    };
  });
}

// --- 在庫引当（フロント連携API） ---
function getAvailableStocks(productId) {
  const stocks = getSheetDataAsObjects_(SHEET_STOCK);
  // 指定された品目コードと一致する在庫情報を抽出
  return stocks.filter(s => String(s.品目コード) === String(productId)).map(s => {
    let expDate = '';
    if (s.賞味期限) {
      expDate = (s.賞味期限 instanceof Date) ? Utilities.formatDate(s.賞味期限, 'JST', 'yyyy/MM/dd') : s.賞味期限;
    }
    return {
      warehouseCode: s.倉庫コード || '',
      warehouseName: s.倉庫名 || '',
      lotNumber: s.ロット番号 || '',
      expirationDate: expDate,
      stockQty: s.在庫数量 || 0,
      unit: s.単位 || ''
    };
  });
}

function debugCheckColumns() {
  const parents = getSheetDataAsObjects_(SHEET_ORDER_PARENT, SPREADSHEET_ID_ORDERS);
  const children = getSheetDataAsObjects_(SHEET_ORDER_CHILD, SPREADSHEET_ID_ORDERS);

  Logger.log('=== 注文履歴_親 ===');
  Logger.log('行数: ' + parents.length);
  if (parents.length > 0) {
    Logger.log('カラム名: ' + JSON.stringify(Object.keys(parents[0])));
    Logger.log('先頭行: ' + JSON.stringify(parents[0]));
  }

  Logger.log('=== 注文履歴_子 ===');
  Logger.log('行数: ' + children.length);
  if (children.length > 0) {
    Logger.log('カラム名: ' + JSON.stringify(Object.keys(children[0])));
    Logger.log('先頭行の出荷日: ' + children[0]['出荷日']);
    Logger.log('先頭行の納品日: ' + children[0]['納品日']);
    Logger.log('先頭行の注文ID: ' + children[0]['注文ID']);
  } else {
    Logger.log('子テーブルが空またはシート名不一致');
  }
}
