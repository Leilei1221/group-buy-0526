// 幾隻貓團購 5/26 - Google Apps Script
// 每次修改後請重新部署新版本

const PRODUCT_LIST = [
  { id: 'p01', name: '厚切排骨', unit: '片', price: 125 },
  { id: 'p02', name: '金饌雞腿', unit: '片', price: 140 },
  { id: 'p03', name: '鹹酥雞腿肉', unit: '包', price: 250 },
];

function doPost(e) {
  try {
    const params = e.parameter;
    const name = params.name || '';
    const summary = params.summary || '';
    const total = parseInt(params.total || '0', 10);
    const note = params.note || '';

    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheets()[0];

    // 確保標題列存在
    if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).getValue() === '') {
      setupHeaders(sheet);
    }

    // 建立資料列
    const row = [new Date(), name];
    PRODUCT_LIST.forEach(p => {
      const q = parseInt(params[p.id] || '0', 10);
      row.push(q > 0 ? q : '');
    });
    row.push(summary, total, '');

    const lastRow = sheet.getLastRow() + 1;
    sheet.getRange(lastRow, 1, 1, row.length).setValues([row]);

    // 格式化新列
    const isEven = (lastRow % 2 === 0);
    sheet.getRange(lastRow, 1, 1, row.length)
      .setBackground(isEven ? '#e8f5ea' : '#ffffff');

    // 數量欄置中
    sheet.getRange(lastRow, 3, 1, PRODUCT_LIST.length)
      .setHorizontalAlignment('center');

    // 訂購摘要欄 wrap
    sheet.getRange(lastRow, 3 + PRODUCT_LIST.length).setWrap(true);

    // 產生備註及回覆欄（自動產生，非使用者備註）
    const itemLines = [];
    PRODUCT_LIST.forEach(p => {
      const q = parseInt(params[p.id] || '0', 10);
      if (q > 0) itemLines.push(`${p.name} ${q}${p.unit} 單價${p.price}元`);
    });
    const replyMsg = `${name}你好，你所訂購的產品如下：${itemLines.join('、')}、總金額：${total}元${note ? '\n備註：' + note : ''}`;
    const replyCol = 3 + PRODUCT_LIST.length + 2;
    sheet.getRange(lastRow, replyCol).setValue(replyMsg).setWrap(true);

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', msg: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('SPREADSHEET_ID');
  if (ssId) {
    try { return SpreadsheetApp.openById(ssId); } catch(e) {}
  }
  const ss = SpreadsheetApp.create('幾隻貓團購 5/26 訂單');
  ssId = ss.getId();
  props.setProperty('SPREADSHEET_ID', ssId);
  setupHeaders(ss.getSheets()[0]);
  return ss;
}

function setupHeaders(sheet) {
  const headers = ['時間戳記', '姓名'];
  PRODUCT_LIST.forEach(p => headers.push(`${p.name}\n$${p.price}`));
  headers.push('訂購摘要', '總金額（元）', '備註及回覆');

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1a5c2a')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setWrap(true);
  sheet.setRowHeight(1, 48);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);

  // 欄寬
  sheet.setColumnWidth(1, 155);
  sheet.setColumnWidth(2, 75);
  for (let i = 0; i < PRODUCT_LIST.length; i++) {
    sheet.setColumnWidth(3 + i, 78);
  }
  const baseCol = 3 + PRODUCT_LIST.length;
  sheet.setColumnWidth(baseCol, 220);
  sheet.setColumnWidth(baseCol + 1, 90);
  sheet.setColumnWidth(baseCol + 2, 180);
}

// 查詢試算表連結（在 GAS 編輯器執行後看 Logger）
function getSpreadsheetUrl() {
  const props = PropertiesService.getScriptProperties();
  const ssId = props.getProperty('SPREADSHEET_ID');
  if (!ssId) { Logger.log('尚無試算表，請先送出一筆訂單'); return; }
  Logger.log(SpreadsheetApp.openById(ssId).getUrl());
}
