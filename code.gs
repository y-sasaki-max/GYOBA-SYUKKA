// --- 設定: データ用スプレッドシート ---
const SPREADSHEET_ID = '1TpkECkY_JUu_g7SogBTz1yjG7PusBn_-vqKk_UJf0S4'; 
const SHEET_CUSTOMER = '顧客マスタ'; 
const SHEET_PRODUCT = '商品マスタ';
const SHEET_CARRIER_MASTER = '配送会社マスタ';
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

// yyyyMMdd文字列 または Date型 を yyyy-MM-dd に変換（HTML表示用）
function parseYMDString_(val) {
  if (!val) return '';
  const s = String(val);
  if (s.length === 8 && !isNaN(s)) {
    const y = s.substring(0, 4);
    const m = s.substring(4, 6);
    const d = s.substring(6, 8);
    return `${y}-${m}-${d}`;
  }
  try {
    return Utilities.formatDate(new Date(val), 'JST', 'yyyy-MM-dd');
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

  return `${today} ${baseName} ${count}.pdf`;
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
   const products = getSheetDataAsObjects_(SHEET_PRODUCT);
   return products.map(p => {
     // ★追加: 単位1〜5 と 換算数量1〜5 をオブジェクト（辞書）化
     const unitMap = {};
     for(let i=1; i<=5; i++) {
         const u = p[`単位${i}`];
         const r = p[`換算数量${i}`];
         if (u && String(u).trim() !== '') {
             // 換算数量が空欄の場合は1として扱う
             unitMap[u] = Number(r) || 1;
         }
     }

     return { 
       id: p.品目コード,             
       name: p.品目名,                
       standard: p.規格 || '',        
       brand: p.ブランド名 || '',      
       kana: p.品目名カナ || '',      
       unit: p.単位1 || '',          
       unitMap: unitMap, // ★ここに追加                 
       price: Number(p.標準売上単価) || 0, 
       caseQty: Number(p.入数) || 1, 
       gosu: Number(p.合数) || 0,
       tempType: (p.在庫管理区分 === 'する' || p.在庫管理区分 === '冷凍') ? 'frozen' : 'chilled' 
     };
   });
}

function getAllDeliveryDestinations() {
  const destinations = getSheetDataAsObjects_(SHEET_DELIVERY_MASTER, DELIVERY_SS_ID);
  return destinations.map(d => ({
    id: d.納入先コード,
    name: d.納入先名1,
    customerId: String(d.取引先コード) 
  }));
}

function getDashboardData() {
  const todayStr = getTodayString_();
  const parents = getSheetDataAsObjects_(SHEET_ORDER_PARENT);
  const children = getSheetDataAsObjects_(SHEET_ORDER_CHILD);
  const todaysParents = parents.filter(p => {
    if(!p.注文日) return false;
    const pDateStr = parseYMDString_(p.注文日); 
    return pDateStr === todayStr;
  });
  const total = todaysParents.length;
  const confirmed = todaysParents.filter(p => p.ステータス === '確定済み').length;
  const pending = total - confirmed;
  const parentIds = new Set(todaysParents.map(p => String(p.注文ID)));
  const relatedChildren = children.filter(c => parentIds.has(String(c.注文ID)));
  const sales = relatedChildren.reduce((sum, c) => sum + (Number(c.金額) || 0), 0);
  
  const alertCustomers = [];
  return {
    total, confirmed, pending, sales, alertCustomers
  };
}

function getRecentOrders() {
  const today = new Date();
  const dateStrings = [];
  
  for (let i = 0; i < 3; i++) {
    const d = new Date();
    d.setDate(today.getDate() + i);
    dateStrings.push(Utilities.formatDate(d, 'JST', 'yyyy-MM-dd'));
  }
  
  const parents = getSheetDataAsObjects_(SHEET_ORDER_PARENT);
  const children = getSheetDataAsObjects_(SHEET_ORDER_CHILD);
  const products = getSheetDataAsObjects_(SHEET_PRODUCT);
  
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

    let hasRecentItem = false;

    const items = myChildren.map(c => {
      const pId = c.品目コード || c.商品ID;
      const pInfo = productInfoMap.get(String(pId)) || {};
      
      const rawDate = c['納品日'] || parent['納品日'] || parentOrderDateStr;
      const itemDateStr = parseYMDString_(rawDate);

      const rawShipDate = c['出荷日'] || parent['出荷日'] || '';
      const shipDateStr = parseYMDString_(rawShipDate);

      if (dateStrings.includes(itemDateStr)) {
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
        remarks: c.摘要 || c.備考
      };
    });

    if (hasRecentItem) {
      const total = items.reduce((sum, item) => sum + item.price, 0);
      const custId = parent.取引先コード || parent.得意先コード || parent.顧客ID;
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
        orderId: pid, 
        items: items,
        totalPrice: total
      });
    }
  });

  return result;
}

function deleteOrder(detailId, orderId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
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
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
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
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const parentSheet = ss.getSheetByName(SHEET_ORDER_PARENT);
  const childSheet = ss.getSheetByName(SHEET_ORDER_CHILD);

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

    let parentCustId = '';
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

    const products = getSheetDataAsObjects_(SHEET_PRODUCT);
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
      delivery: cHeaders.indexOf('納入先')    
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
      
      const inputCaseCount = Number(item.caseCount) || 0;
      const inputQuantity = Number(item.quantity) || 0;

      const finalCaseCount = inputCaseCount;
      const finalQuantity = finalCaseCount > 0 ? (finalCaseCount * pInfo.caseQty) : inputQuantity;

      const unitPrice = pInfo.price;
      const amount = finalQuantity * unitPrice; 

      const remarks = item.newRemarks || item.remarks;
      const standard = item.standard || ''; 
      const unit = item.unit || '';
      const awaseVal = item.awaseId || ''; 

      if (!isNew && detailIdMap.has(String(item.detailId))) {
        const rowNum = detailIdMap.get(String(item.detailId));
        if(cIdx.prodId !== -1) childSheet.getRange(rowNum, cIdx.prodId + 1).setValue(item.productId);
        if(cIdx.unit !== -1) childSheet.getRange(rowNum, cIdx.unit + 1).setValue(unit);
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
        if(cIdx.standard !== -1) childSheet.getRange(rowNum, cIdx.standard + 1).setValue(standard); 
        if(cIdx.awase !== -1) childSheet.getRange(rowNum, cIdx.awase + 1).setValue(awaseVal);
        if(cIdx.customer !== -1) childSheet.getRange(rowNum, cIdx.customer + 1).setValue(parentCustId);
        if(cIdx.delivery !== -1) childSheet.getRange(rowNum, cIdx.delivery + 1).setValue(parentDelivery);
      } else {
        const newDetailId = generateRandomId_();
        const row = [
          newDetailId,          
          targetOrderId,        
          '',                   
          itemShipDateObj,      
          itemDeliveryDateObj,  
          index + 1,            
          item.productId,       
          unit,                 
          pInfo.caseQty,        
          pInfo.gosu,           
          finalCaseCount,       
          '',                   
          '',                   
          finalQuantity,        
          unitPrice,            
          amount,               
          pInfo.taxRate,        
          pInfo.taxType,        
          remarks,              
          standard,             
          awaseVal,             
          parentCustId,         
          parentDelivery        
        ];
        newRows.push(row);
      }
    });

    if(newRows.length > 0) {
      childSheet.getRange(childSheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    }

    CacheService.getScriptCache().remove(CACHE_KEY);
    return { status: 'success' };
  } finally {
    lock.releaseLock();
  }
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
          deliveryDestination: order.deliveryDestination,
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

    const ssMain = SpreadsheetApp.openById(SPREADSHEET_ID);
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

    shipmentsArray.forEach((shipment, index) => {
      const suffix = `_${index}`;
      shipment.orderIds.forEach(id => processedOrderIds.push(id));
      
      const splitItems = createSplitRows_(shipment.items);
      const processedShipment = { ...shipment, items: splitItems };

      if (templateInvoice) {
        const sheetInv = templateInvoice.copyTo(newSs).setName(`Inv${suffix}`);
        fillInvoiceSheet_(sheetInv, processedShipment);
        generatedSheetIds.push(sheetInv.getSheetId()); 
      }

      if (templateDelivery) {
        const sheetDel = templateDelivery.copyTo(newSs).setName(`Del${suffix}`);
        fillDeliverySheet_(sheetDel, processedShipment); 
        generatedSheetIds.push(sheetDel.getSheetId()); 
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
    const pdfBlob = getBlobFromWholeSs_(newSs, filename);
    const base64 = Utilities.base64Encode(pdfBlob.getBytes());

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
      filename: pdfBlob.getName(),
      message: `${rawTargetOrders.length}件の注文から、${shipmentsArray.length}件の出荷伝票（計${generatedSheetIds.length}枚）を生成しました。`
    };
  } catch(e) {
    throw e;
  } finally {
    newFile.setTrashed(true);
  }
}

function updateSlipNumbersAndFlags_(slipUpdates, orderIds) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
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

function fillInvoiceSheet_(sheet, shipment) {
  sheet.getRange('A7').setValue(shipment.customerAddress || ''); 
  sheet.getRange('A9').setValue(shipment.customerName + ' 御中');

  sheet.getRange('D3').setValue(getTodayString_());
  sheet.getRange('C4').setValue(shipment.deliveryDate); 
  
  sheet.getRange('I1').setValue(shipment.slipNumber || ''); 
  sheet.getRange('B28').setValue(shipment.carrierName || ''); 

  sheet.getRange(13, 1, 14, 8).clearContent();
  const rows = shipment.items.map(item => [
    item.productName,       
    item.standard || '-',   
    item.caseQty,           
    item.unit || '',        
    item.unitPrice,         
    item.displayBox,        
    item.quantity,          
    item.price   
  ]);
  if (rows.length > 0) {
    sheet.getRange(13, 1, rows.length, 8).setValues(rows);
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
    item.caseQty,
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
  sheet.getRange('B21').setValue(shipment.deliveryDestination || ''); 

  sheet.getRange('D21').setValue(totalCase); 
  sheet.getRange('E21').setValue(totalQty);
  sheet.getRange('F21').setValue("合計:" + totalAmount.toLocaleString()); 
  sheet.getRange('G21').setValue("消費税:" + tax.toLocaleString()); 
  sheet.getRange('H21').setValue("金額:" + totalIncTax.toLocaleString()); 
}

function getBlobFromWholeSs_(ss, filename) {
  const url = ss.getUrl().replace(/\/edit.*$/, '') 
    + '/export?exportFormat=pdf&format=pdf'
    + '&size=A4'
    + '&portrait=false' 
    + '&fitw=true' 
    + '&gridlines=false';
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
  try { csvData = Utilities.parseCsv(csvContent); } catch (e) { throw new Error('CSV読込失敗'); }
  if (csvData.length === 0) return { status: 'success', message: 'データなし' };

  const lastRow = sheet.getLastRow();
  let existingValues = [];
  if (lastRow > 1) existingValues = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const codeMap = new Map();
  existingValues.forEach((row, index) => { if (String(row[0])) codeMap.set(String(row[0]), index); });

  const updates = [];
  const appends = [];
  let startIndex = (csvData.length > 0 && (csvData[0][0] === '納入先コード' || csvData[0][0] === 'コード')) ? 1 : 0;

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
