const SHEET_NAME = 'Bookings';
const SHEET_ID = '';
const DATA_START_ROW = 5;
const DATA_START_COL = 2;
const DATA_COL_COUNT = 34;

const COL = {
  TRAVEL_DATE: 2,
  FIRST_NAME: 3,
  LAST_NAME: 4,
  PHONE: 6,
  NUMBER_OF_GUEST: 10,
  TOUR: 11,
  COUPON_CODE: 18,
  DISCOUNT_AMOUNT: 19,
  PAYMENT_METHOD: 20,
  TOTAL_AMOUNT: 22,
  DOWNPAYMENT: 24,
  RECEIPT_DOWNPAYMENT: 25,
  TOTAL_BALANCE: 26,
  RECEIPT_FULL_PAYMENT: 29,
  ADDITIONAL_AD: 30,
  WHATSAPP_NEW: 32,
  REMARKS_NEW: 33,
  WHATSAPP_CONFIRMED: 34,
  REMARKS_CONFIRMED: 35
};

function doGet() {
  try {
    var sheet = getSheet_();
    var lastRow = sheet.getLastRow();

    if (lastRow < DATA_START_ROW) {
      return jsonResponse_([]);
    }

    var rowCount = lastRow - DATA_START_ROW + 1;
    var rows = sheet.getRange(DATA_START_ROW, DATA_START_COL, rowCount, DATA_COL_COUNT).getValues();
    var waNewRich = sheet.getRange(DATA_START_ROW, COL.WHATSAPP_NEW, rowCount, 1).getRichTextValues();
    var waConfirmedRich = sheet.getRange(DATA_START_ROW, COL.WHATSAPP_CONFIRMED, rowCount, 1).getRichTextValues();

    var payload = rows.map(function (row, idx) {
      var waNew = waNewRich[idx] && waNewRich[idx][0] ? waNewRich[idx][0].getLinkUrl() : '';
      var waConfirmed = waConfirmedRich[idx] && waConfirmedRich[idx][0] ? waConfirmedRich[idx][0].getLinkUrl() : '';

      return {
        rowIndex: idx + DATA_START_ROW,
        name: (safeCell_(cell_(row, COL.FIRST_NAME)) + ' ' + safeCell_(cell_(row, COL.LAST_NAME))).trim(),
        phone: safeCell_(cell_(row, COL.PHONE)),
        tour: safeCell_(cell_(row, COL.TOUR)),
        numberOfGuest: safeCell_(cell_(row, COL.NUMBER_OF_GUEST)),
        couponCode: safeCell_(cell_(row, COL.COUPON_CODE)),
        discountAmount: safeCell_(cell_(row, COL.DISCOUNT_AMOUNT)),
        paymentMethod: safeCell_(cell_(row, COL.PAYMENT_METHOD)),
        totalAmount: safeCell_(cell_(row, COL.TOTAL_AMOUNT)),
        downpaymentAmount: safeCell_(cell_(row, COL.DOWNPAYMENT)),
        receiptDownpayment: safeCell_(cell_(row, COL.RECEIPT_DOWNPAYMENT)),
        totalBalance: safeCell_(cell_(row, COL.TOTAL_BALANCE)),
        receiptFullPayment: safeCell_(cell_(row, COL.RECEIPT_FULL_PAYMENT)),
        adData: safeCell_(cell_(row, COL.ADDITIONAL_AD)),
        date: formatDateCell_(cell_(row, COL.TRAVEL_DATE)),
        whatsappNew: waNew || safeCell_(cell_(row, COL.WHATSAPP_NEW)),
        remarksNew: safeCell_(cell_(row, COL.REMARKS_NEW)),
        whatsappConfirmed: waConfirmed || safeCell_(cell_(row, COL.WHATSAPP_CONFIRMED)),
        remarksConfirmed: safeCell_(cell_(row, COL.REMARKS_CONFIRMED))
      };
    });

    return jsonResponse_(payload);
  } catch (error) {
    return jsonResponse_({ status: 'error', message: error.message });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Missing request body.');
    }

    var body = JSON.parse(e.postData.contents);
    var rowIndex = Number(body.rowIndex);

    if (!rowIndex || rowIndex < DATA_START_ROW) {
      throw new Error('Invalid rowIndex');
    }

    var sheet = getSheet_();
    if (rowIndex > sheet.getLastRow()) {
      throw new Error('rowIndex out of range');
    }

    var action = safeCell_(body.action || 'mark_sent_new').toLowerCase();

    if (action === 'mark_sent_new') {
      sheet.getRange(rowIndex, COL.REMARKS_NEW).setValue('sent');
      return jsonResponse_({ status: 'success' });
    }

    if (action === 'mark_sent_confirmed') {
      sheet.getRange(rowIndex, COL.REMARKS_CONFIRMED).setValue('sent');
      return jsonResponse_({ status: 'success' });
    }

    if (action === 'save_receipt') {
      return saveReceipt_(sheet, rowIndex, body);
    }

    throw new Error('Unsupported action: ' + action);
  } catch (error) {
    return jsonResponse_({ status: 'error', message: error.message });
  }
}

function saveReceipt_(sheet, rowIndex, body) {
  var paymentType = safeCell_(body.paymentType || 'downpayment').toLowerCase();
  var invoiceNumber = safeCell_(body.invoiceNumber);
  var encodedDownpaymentAmount = safeCell_(body.encodedDownpaymentAmount);
  var encodedFullPaymentAmount = safeCell_(body.encodedFullPaymentAmount);

  if (!invoiceNumber) {
    throw new Error('invoiceNumber is required');
  }

  var targetCol = paymentType === 'full' ? COL.RECEIPT_FULL_PAYMENT : COL.RECEIPT_DOWNPAYMENT;
  var targetCell = sheet.getRange(rowIndex, targetCol);
  targetCell.setValue(invoiceNumber);

  if (paymentType === 'downpayment' && encodedDownpaymentAmount) {
    var normalized = normalizeAmount_(encodedDownpaymentAmount);
    if (normalized === null) {
      throw new Error('Invalid encodedDownpaymentAmount');
    }
    var currentDownpayment = normalizeAmount_(sheet.getRange(rowIndex, COL.DOWNPAYMENT).getValue()) || 0;
    if (normalized < currentDownpayment) {
      throw new Error('Downpayment should start at the amount encoded in column X');
    }
    sheet.getRange(rowIndex, COL.DOWNPAYMENT).setValue(normalized);
  }

  if (paymentType === 'full') {
    var normalizedFull = normalizeAmount_(encodedFullPaymentAmount);
    if (normalizedFull === null) {
      throw new Error('encodedFullPaymentAmount is required for full payment');
    }
    var currentBalance = normalizeAmount_(sheet.getRange(rowIndex, COL.TOTAL_BALANCE).getValue()) || 0;
    if (normalizedFull !== currentBalance) {
      throw new Error('Full payment amount must exactly match column Z');
    }
  }

  var imageBase64 = safeCell_(body.receiptImageBase64);
  if (imageBase64) {
    var mimeType = safeCell_(body.receiptMimeType) || 'image/png';
    var filename = safeCell_(body.receiptFileName) || ('receipt_' + rowIndex + '_' + Date.now() + '.png');

    var bytes = Utilities.base64Decode(imageBase64);
    var blob = Utilities.newBlob(bytes, mimeType, filename);
    var file = DriveApp.createFile(blob);
    targetCell.setNote('Receipt image: ' + file.getUrl());
    return jsonResponse_({ status: 'success', message: 'Invoice saved with image link note.' });
  }

  return jsonResponse_({ status: 'success', message: 'Invoice saved.' });
}

function doOptions() {
  return jsonResponse_({ status: 'ok' });
}

function getSheet_() {
  var spreadsheet = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet not found: ' + SHEET_NAME);
  }
  return sheet;
}

function jsonResponse_(payload) {
  var output = ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);

  if (typeof output.setHeader === 'function') {
    output.setHeader('Access-Control-Allow-Origin', '*');
    output.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
    output.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  }

  return output;
}

function safeCell_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function formatDateCell_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return safeCell_(value);
}

function cell_(row, absoluteCol) {
  var offset = absoluteCol - DATA_START_COL;
  return row[offset];
}

function normalizeAmount_(value) {
  var cleaned = String(value).replace(/[^0-9.-]/g, '').trim();
  if (!cleaned) return null;
  var numeric = Number(cleaned);
  if (!isFinite(numeric) || numeric < 0) return null;
  return Number(numeric.toFixed(2));
}
