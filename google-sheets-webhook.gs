const SHEET_NAME = '订单';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getOrderSheet_();
    const data = JSON.parse(e.postData.contents || '{}');
    const row = [
      data.createdAt || new Date().toISOString(),
      data.orderId || '',
      data.name || '',
      data.phone || '',
      data.address || '',
      data.items || '',
      data.total || 0,
      data.note || '',
      data.status || '待处理',
      data.source || ''
    ];
    sheet.appendRow(row);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function getOrderSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  const headers = [
    '下单时间',
    '订单号',
    '收货人',
    '电话',
    '地址',
    '商品明细',
    '总金额',
    '备注',
    '状态',
    '来源页面'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}
