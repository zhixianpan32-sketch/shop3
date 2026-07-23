var ORDER_SHEET_NAME = '\u8ba2\u5355';
var ITEM_SHEET_NAME = '\u8ba2\u5355\u660e\u7ec6';
var SUMMARY_SHEET_NAME = '\u5546\u54c1\u6c47\u603b';
var SPREADSHEET_PROPERTY = 'SHOP3_ORDER_SPREADSHEET_ID';
var SPREADSHEET_TITLE = 'Shop3 \u81ea\u52a8\u8ba2\u5355\u8868';

var ORDER_HEADERS = [
  '\u4e0b\u5355\u65f6\u95f4',
  '\u8ba2\u5355\u53f7',
  '\u6536\u8d27\u4eba',
  '\u7535\u8bdd',
  '\u5730\u5740',
  '\u5546\u54c1\u660e\u7ec6',
  '\u5546\u54c1\u91d1\u989d',
  '\u6253\u8d4f\u9a91\u624b',
  '\u6253\u8d4f\u5546\u5bb6',
  '\u603b\u91d1\u989d',
  '\u5907\u6ce8',
  '\u72b6\u6001',
  '\u6765\u6e90\u9875\u9762'
];

var ITEM_HEADERS = [
  '\u4e0b\u5355\u65f6\u95f4',
  '\u8ba2\u5355\u53f7',
  '\u5546\u54c1ID',
  '\u5546\u54c1\u540d\u79f0',
  '\u5e97\u94fa',
  '\u6570\u91cf',
  '\u5355\u4ef7',
  '\u5c0f\u8ba1'
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  var locked = false;

  try {
    lock.waitLock(10000);
    locked = true;

    var raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    if (raw.length > 100000) throw new Error('Payload too large');

    var data = JSON.parse(raw);
    var orderId = safeText_(data.orderId, 80);
    if (!orderId) throw new Error('Missing orderId');

    var spreadsheet = getOrderSpreadsheet_();
    var schema = ensureSchema_(spreadsheet);

    if (hasOrder_(schema.orders, orderId)) {
      return jsonOutput_({
        ok: true,
        duplicate: true,
        orderId: orderId,
        spreadsheetUrl: spreadsheet.getUrl()
      });
    }

    var createdAt = dateValue_(data.createdAt);
    var orderRow = schema.orders.getLastRow() + 1;
    schema.orders.appendRow([
      createdAt,
      orderId,
      safeText_(data.name, 120),
      safeText_(data.phone, 80),
      safeText_(data.address, 500),
      safeText_(data.items, 5000),
      nonNegativeNumber_(data.goodsTotal || data.total),
      nonNegativeNumber_(data.tipRider),
      nonNegativeNumber_(data.tipMerchant),
      nonNegativeNumber_(data.total),
      safeText_(data.note, 1000),
      safeText_(data.status || '\u5f85\u5904\u7406', 40),
      safeText_(data.source, 500)
    ]);
    schema.orders.getRange(orderRow, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    schema.orders.getRange(orderRow, 7, 1, 4).setNumberFormat('0.00');

    appendItemRows_(schema.items, data.itemRows, createdAt, orderId);
    SpreadsheetApp.flush();

    return jsonOutput_({
      ok: true,
      duplicate: false,
      orderId: orderId,
      spreadsheetUrl: spreadsheet.getUrl()
    });
  } catch (err) {
    return jsonOutput_({ ok: false, error: String(err) });
  } finally {
    if (locked) lock.releaseLock();
  }
}

function doGet() {
  try {
    var spreadsheet = getOrderSpreadsheet_();
    var schema = ensureSchema_(spreadsheet);
    return jsonOutput_({
      ok: true,
      message: 'ready',
      spreadsheetName: spreadsheet.getName(),
      spreadsheetUrl: spreadsheet.getUrl(),
      orderCount: Math.max(0, schema.orders.getLastRow() - 1),
      itemCount: Math.max(0, schema.items.getLastRow() - 1),
      sheets: [ORDER_SHEET_NAME, ITEM_SHEET_NAME, SUMMARY_SHEET_NAME]
    });
  } catch (err) {
    return jsonOutput_({ ok: false, error: String(err) });
  }
}

function getOrderSpreadsheet_() {
  var properties = PropertiesService.getScriptProperties();
  var savedId = properties.getProperty(SPREADSHEET_PROPERTY);

  if (savedId) {
    try {
      return SpreadsheetApp.openById(savedId);
    } catch (err) {
      properties.deleteProperty(SPREADSHEET_PROPERTY);
    }
  }

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    spreadsheet = SpreadsheetApp.create(SPREADSHEET_TITLE);
  }

  properties.setProperty(SPREADSHEET_PROPERTY, spreadsheet.getId());
  return spreadsheet;
}

function ensureSchema_(spreadsheet) {
  var orders = ensureSheet_(spreadsheet, ORDER_SHEET_NAME, ORDER_HEADERS);
  var items = ensureSheet_(spreadsheet, ITEM_SHEET_NAME, ITEM_HEADERS);
  migrateLegacyOrders_(orders, items);
  ensureSummarySheet_(spreadsheet);

  return { orders: orders, items: items };
}

function ensureSheet_(spreadsheet, name, headers) {
  var sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  } else {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#FFF1ED')
    .setFontColor('#E8492D');
  sheet.setFrozenRows(1);
  return sheet;
}

function ensureSummarySheet_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(SUMMARY_SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SUMMARY_SHEET_NAME);

  sheet.getRange('A1:C1')
    .setValues([['\u5546\u54c1', '\u9500\u91cf', '\u9500\u552e\u989d']])
    .setFontWeight('bold')
    .setBackground('#FFF1ED')
    .setFontColor('#E8492D');
  sheet.setFrozenRows(1);

  var formulaCell = sheet.getRange('A2');
  if (!formulaCell.getFormula() && !formulaCell.getDisplayValue()) {
    formulaCell.setFormula(
      '=QUERY(\'' + ITEM_SHEET_NAME + '\'!D2:H,' +
      '"select D, sum(F), sum(H) where D is not null group by D ' +
      'label D \'\', sum(F) \'\', sum(H) \'\'",0)'
    );
    sheet.getRange('B2:C1000').setNumberFormat('0.00');
  }
}

function hasOrder_(sheet, orderId) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  return sheet
    .getRange(2, 2, lastRow - 1, 1)
    .createTextFinder(orderId)
    .matchEntireCell(true)
    .findNext() !== null;
}

function appendItemRows_(sheet, itemRows, createdAt, orderId) {
  if (!Array.isArray(itemRows) || itemRows.length === 0) return;

  var rows = [];
  itemRows.slice(0, 200).forEach(function(item) {
    var quantity = Math.max(1, Math.floor(nonNegativeNumber_(item.qty) || 1));
    var unitPrice = nonNegativeNumber_(item.unitPrice);
    rows.push([
      createdAt,
      orderId,
      safeText_(item.productId, 80),
      safeText_(item.name || '\u5546\u54c1', 300),
      safeText_(item.store, 100),
      quantity,
      unitPrice,
      nonNegativeNumber_(item.subtotal || unitPrice * quantity)
    ]);
  });

  if (rows.length) {
    var firstRow = sheet.getLastRow() + 1;
    sheet.getRange(firstRow, 1, rows.length, ITEM_HEADERS.length).setValues(rows);
    sheet.getRange(firstRow, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sheet.getRange(firstRow, 6, rows.length, 3).setNumberFormat('0.00');
  }
}

function migrateLegacyOrders_(ordersSheet, itemsSheet) {
  var orderRowCount = ordersSheet.getLastRow() - 1;
  if (orderRowCount <= 0) return;

  var orderRange = ordersSheet.getRange(2, 1, orderRowCount, ORDER_HEADERS.length);
  var orderRows = orderRange.getValues();
  var existingItemOrders = {};
  var itemRowCount = itemsSheet.getLastRow() - 1;

  if (itemRowCount > 0) {
    itemsSheet.getRange(2, 2, itemRowCount, 1).getValues().forEach(function(row) {
      if (row[0]) existingItemOrders[String(row[0])] = true;
    });
  }

  var changed = false;
  var historicalItems = [];

  orderRows.forEach(function(row) {
    var legacyStatus = String(row[8] || '').trim();
    var legacySource = String(row[9] || '').trim();
    var isLegacy = !row[11] && !row[12] && (
      legacyStatus === '\u5f85\u5904\u7406' ||
      legacyStatus === '\u5df2\u5b8c\u6210' ||
      legacySource ||
      row[7]
    );

    if (isLegacy) {
      var oldTotal = nonNegativeNumber_(row[6]);
      var oldNote = row[7];
      var oldStatus = row[8] || '\u5f85\u5904\u7406';
      var oldSource = row[9];

      row[7] = 0;
      row[8] = 0;
      row[9] = oldTotal;
      row[10] = oldNote;
      row[11] = oldStatus;
      row[12] = oldSource;
      changed = true;
    }

    var orderId = String(row[1] || '').trim();
    if (!orderId || existingItemOrders[orderId]) return;

    var parsedItems = parseLegacyItems_(row[5]);
    parsedItems.forEach(function(item) {
      historicalItems.push([
        dateValue_(row[0]),
        safeText_(orderId, 80),
        '',
        safeText_(item.name, 300),
        '',
        item.qty,
        item.unitPrice,
        item.subtotal
      ]);
    });

    if (parsedItems.length) existingItemOrders[orderId] = true;
  });

  if (changed) {
    orderRange.setValues(orderRows);
    ordersSheet.getRange(2, 1, orderRowCount, 1)
      .setNumberFormat('yyyy-mm-dd hh:mm:ss');
    ordersSheet.getRange(2, 7, orderRowCount, 4)
      .setNumberFormat('0.00');
  }

  if (historicalItems.length) {
    var firstItemRow = itemsSheet.getLastRow() + 1;
    itemsSheet
      .getRange(firstItemRow, 1, historicalItems.length, ITEM_HEADERS.length)
      .setValues(historicalItems);
    itemsSheet
      .getRange(firstItemRow, 1, historicalItems.length, 1)
      .setNumberFormat('yyyy-mm-dd hh:mm:ss');
    itemsSheet
      .getRange(firstItemRow, 6, historicalItems.length, 3)
      .setNumberFormat('0.00');
  }
}

function parseLegacyItems_(value) {
  var text = String(value || '').trim();
  if (!text) return [];

  var parsed = [];
  text.split('\uff1b').forEach(function(part) {
    var match = part.trim().match(/^(.*?)\s+x(\d+)\s*=\s*([0-9]+(?:\.[0-9]+)?)/i);
    if (!match) return;

    var quantity = Math.max(1, Math.floor(nonNegativeNumber_(match[2]) || 1));
    var subtotal = nonNegativeNumber_(match[3]);
    parsed.push({
      name: match[1].trim(),
      qty: quantity,
      unitPrice: subtotal / quantity,
      subtotal: subtotal
    });
  });

  return parsed;
}

function safeText_(value, maxLength) {
  var text = value === null || value === undefined ? '' : String(value).trim();
  if (text.length > maxLength) text = text.substring(0, maxLength);
  if (/^[=+\-@]/.test(text)) text = "'" + text;
  return text;
}

function nonNegativeNumber_(value) {
  var number = Number(value);
  return isFinite(number) ? Math.max(0, number) : 0;
}

function dateValue_(value) {
  var date = value ? new Date(value) : new Date();
  return isNaN(date.getTime()) ? new Date() : date;
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
