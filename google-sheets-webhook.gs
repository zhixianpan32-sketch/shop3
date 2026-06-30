var SHEET_NAME = '\u8ba2\u5355';

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    var sheet = getOrderSheet_();
    var raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    var data = JSON.parse(raw);

    sheet.appendRow([
      data.createdAt || new Date().toISOString(),
      data.orderId || '',
      data.name || '',
      data.phone || '',
      data.address || '',
      data.items || '',
      data.total || 0,
      data.note || '',
      data.status || '\u5f85\u5904\u7406',
      data.source || ''
    ]);

    return jsonOutput_({ ok: true });
  } catch (err) {
    return jsonOutput_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  getOrderSheet_();
  return jsonOutput_({ ok: true, message: 'ready' });
}

function getOrderSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  var headers = [
    '\u4e0b\u5355\u65f6\u95f4',
    '\u8ba2\u5355\u53f7',
    '\u6536\u8d27\u4eba',
    '\u7535\u8bdd',
    '\u5730\u5740',
    '\u5546\u54c1\u660e\u7ec6',
    '\u603b\u91d1\u989d',
    '\u5907\u6ce8',
    '\u72b6\u6001',
    '\u6765\u6e90\u9875\u9762'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
