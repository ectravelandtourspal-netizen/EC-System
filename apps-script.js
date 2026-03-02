const SHEET_NAME = 'Bookings';
const SHEET_ID = '';
const COUPON_SHEET_NAME = 'COUPONS';
const EXCLUDED_COUPON_CELL = 'A6';
const REPORT_SHEET_NAME = 'INFLUENCER PAYMENT';
const REPORT_START_ROW = 10;
const DATA_START_ROW = 5;
const DATA_START_COL = 2;
const DATA_COL_COUNT = 37;

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
  COMMISSION_AMOUNT_AA: 27,
  RECEIPT_FULL_PAYMENT: 29,
  ADDITIONAL_AD: 30,
  WHATSAPP_NEW: 32,
  REMARKS_NEW: 33,
  WHATSAPP_CONFIRMED: 34,
  REMARKS_CONFIRMED: 35,
  RECEIPT_INVOICE_AK: 37,
  INFLUENCER_STATUS_AL: 38
};

function doGet() {
  try {
    var sheet = getSheet_();
    var excludedCouponCode = getExcludedCouponCode_();
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
        excludedCouponCode: excludedCouponCode,
        discountAmount: safeCell_(cell_(row, COL.DISCOUNT_AMOUNT)),
        paymentMethod: safeCell_(cell_(row, COL.PAYMENT_METHOD)),
        totalAmount: safeCell_(cell_(row, COL.TOTAL_AMOUNT)),
        downpaymentAmount: safeCell_(cell_(row, COL.DOWNPAYMENT)),
        receiptDownpayment: safeCell_(cell_(row, COL.RECEIPT_DOWNPAYMENT)),
        totalBalance: safeCell_(cell_(row, COL.TOTAL_BALANCE)),
        commissionAmountAA: safeCell_(cell_(row, COL.COMMISSION_AMOUNT_AA)),
        influencerInvoiceAK: safeCell_(cell_(row, COL.RECEIPT_INVOICE_AK)),
        influencerStatusAL: safeCell_(cell_(row, COL.INFLUENCER_STATUS_AL)),
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
    var action = safeCell_(body.action || 'mark_sent_new').toLowerCase();

    if (action === 'save_commission_receipt_bulk') {
      return saveCommissionReceiptBulk_(getSheet_(), body);
    }

    var rowIndex = Number(body.rowIndex);

    if (!rowIndex || rowIndex < DATA_START_ROW) {
      throw new Error('Invalid rowIndex');
    }

    var sheet = getSheet_();
    if (rowIndex > sheet.getLastRow()) {
      throw new Error('rowIndex out of range');
    }

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

function saveCommissionReceiptBulk_(sheet, body) {
  var travelDatesRaw = Array.isArray(body.travelDates) ? body.travelDates : [body.travelDate];
  var targetDates = [];
  travelDatesRaw.forEach(function (value) {
    var normalizedDate = normalizeDateKey_(value);
    if (normalizedDate && targetDates.indexOf(normalizedDate) === -1) {
      targetDates.push(normalizedDate);
    }
  });
  var couponCode = safeCell_(body.couponCode);
  var invoiceNumber = safeCell_(body.invoiceNumber);
  var encodedCommissionAmount = normalizeAmount_(body.encodedCommissionAmount);

  if (!targetDates.length) {
    throw new Error('travelDates is required');
  }
  if (!couponCode) {
    throw new Error('couponCode is required');
  }
  if (!invoiceNumber) {
    throw new Error('invoiceNumber is required');
  }
  if (encodedCommissionAmount === null) {
    throw new Error('encodedCommissionAmount is required');
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) {
    throw new Error('No booking rows found.');
  }

  var rowCount = lastRow - DATA_START_ROW + 1;
  var rows = sheet.getRange(DATA_START_ROW, DATA_START_COL, rowCount, DATA_COL_COUNT).getValues();

  var matchedRows = [];
  var sumAA = 0;
  var dateSummary = {};

  rows.forEach(function (row, idx) {
    var rowDate = normalizeDateKey_(cell_(row, COL.TRAVEL_DATE));
    var rowCoupon = safeCell_(cell_(row, COL.COUPON_CODE));
    var rowStatusAL = safeCell_(cell_(row, COL.INFLUENCER_STATUS_AL)).toUpperCase();

    if (targetDates.indexOf(rowDate) !== -1 && rowCoupon === couponCode && rowStatusAL !== 'PAID') {
      var sheetRow = idx + DATA_START_ROW;
      var amountAA = normalizeAmount_(cell_(row, COL.COMMISSION_AMOUNT_AA)) || 0;
      matchedRows.push(sheetRow);
      sumAA += amountAA;

      if (!dateSummary[rowDate]) {
        dateSummary[rowDate] = { bookings: 0, totalAmount: 0 };
      }
      dateSummary[rowDate].bookings += 1;
      dateSummary[rowDate].totalAmount = Number((dateSummary[rowDate].totalAmount + amountAA).toFixed(2));
    }
  });

  sumAA = Number(sumAA.toFixed(2));

  if (!matchedRows.length) {
    throw new Error('No matching bookings found for selected travel date and coupon code.');
  }

  if (sumAA !== encodedCommissionAmount) {
    throw new Error('Encoded amount must exactly match SUM of column AA (' + sumAA.toFixed(2) + ').');
  }

  var imageBase64 = safeCell_(body.receiptImageBase64);
  var imageUrl = '';

  if (imageBase64) {
    var mimeType = safeCell_(body.receiptMimeType) || 'image/png';
    var filename = safeCell_(body.receiptFileName) || ('commission_receipt_' + Date.now() + '.png');
    var bytes = Utilities.base64Decode(imageBase64);
    var blob = Utilities.newBlob(bytes, mimeType, filename);
    var file = DriveApp.createFile(blob);
    imageUrl = file.getUrl();
  }

  var noteLines = [
    'Commissions receipt encoded',
    'Travel Dates: ' + targetDates.join(', '),
    'Coupon Code: ' + couponCode,
    'Invoice Number: ' + invoiceNumber,
    'Encoded Amount: ' + encodedCommissionAmount.toFixed(2),
    'Matched Rows: ' + matchedRows.join(', '),
    'Saved At: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
  ];

  if (imageUrl) {
    noteLines.push('Receipt image: ' + imageUrl);
  }

  var noteText = noteLines.join('\n');
  matchedRows.forEach(function (sheetRow) {
    sheet.getRange(sheetRow, COL.RECEIPT_INVOICE_AK).setValue(invoiceNumber);
    sheet.getRange(sheetRow, COL.COMMISSION_AMOUNT_AA).setNote(noteText);
  });

  appendInfluencerPaymentReport_(targetDates, dateSummary, invoiceNumber, couponCode);

  return jsonResponse_({
    status: 'success',
    message: 'Commissions receipt saved for ' + matchedRows.length + ' booking(s).',
    matchedRows: matchedRows,
    sumAA: sumAA
  });
}

function appendInfluencerPaymentReport_(targetDates, dateSummary, invoiceNumber, couponCode) {
  var spreadsheet = getSpreadsheet_();
  var reportSheet = spreadsheet.getSheetByName(REPORT_SHEET_NAME);
  if (!reportSheet) {
    reportSheet = spreadsheet.insertSheet(REPORT_SHEET_NAME);
  }

  var startCol = getCouponBlockStartCol_(reportSheet, couponCode);
  var startRow = getNextReportStartRowForBlock_(reportSheet, startCol);
  var endCol = startCol + 3;
  var titleRow = startRow;
  var headerRow = startRow + 1;

  reportSheet.getRange(titleRow, startCol, 1, 4).breakApart();
  reportSheet.getRange(titleRow, startCol, 1, 4).merge();
  reportSheet.getRange(titleRow, startCol).setValue('SUMMARY OF PAYMENT WITH INVOICE # ' + invoiceNumber);
  reportSheet.getRange(titleRow, startCol).setFontWeight('bold');

  reportSheet.getRange(headerRow, startCol, 1, 4).setValues([['DATES', 'NO. OF BOOKINGS', 'AMOUNT', 'TOTAL']]);
  reportSheet.getRange(headerRow, startCol, 1, 4).setFontWeight('bold');

  var sortedDates = targetDates.slice().sort();
  var dataRows = sortedDates
    .filter(function (dateKey) { return Boolean(dateSummary[dateKey]); })
    .map(function (dateKey) {
      var info = dateSummary[dateKey];
      var bookings = info.bookings || 0;
      var totalAmount = Number(info.totalAmount || 0);
      var amount = bookings ? Number((totalAmount / bookings).toFixed(2)) : 0;
      return [dateKey, bookings, amount];
    });

  if (!dataRows.length) {
    return;
  }

  var dataStartRow = startRow + 2;
  reportSheet.getRange(dataStartRow, startCol, dataRows.length, 3).setValues(dataRows);

  var formulas = dataRows.map(function (_, idx) {
    var rowNumber = dataStartRow + idx;
    var bookingsCol = columnToLetter_(startCol + 1);
    var amountCol = columnToLetter_(startCol + 2);
    return ['=' + bookingsCol + rowNumber + '*' + amountCol + rowNumber];
  });
  reportSheet.getRange(dataStartRow, endCol, dataRows.length, 1).setFormulas(formulas);

  var grandTotalRow = dataStartRow + dataRows.length;
  var totalColLetter = columnToLetter_(endCol);
  reportSheet.getRange(grandTotalRow, startCol + 2).setValue('GRAND TOTAL');
  reportSheet.getRange(grandTotalRow, endCol).setFormula('=SUM(' + totalColLetter + dataStartRow + ':' + totalColLetter + (grandTotalRow - 1) + ')');
  reportSheet.getRange(grandTotalRow, startCol + 2, 1, 2).setFontWeight('bold');
}

function getCouponBlockStartCol_(reportSheet, couponCode) {
  var normalizedCoupon = safeCell_(couponCode);
  var blockWidth = 4;
  var blockGap = 2;
  var stride = blockWidth + blockGap;

  var lastCol = Math.max(reportSheet.getLastColumn(), 1);
  for (var col = 1; col <= lastCol; col += stride) {
    var marker = safeCell_(reportSheet.getRange(1, col).getValue());
    if (marker === normalizedCoupon) {
      return col;
    }
  }

  for (var assignCol = 1; assignCol <= lastCol + stride; assignCol += stride) {
    var markerAssign = safeCell_(reportSheet.getRange(1, assignCol).getValue());
    if (!markerAssign) {
      reportSheet.getRange(1, assignCol).setValue(normalizedCoupon);
      return assignCol;
    }
  }

  reportSheet.getRange(1, 1).setValue(normalizedCoupon);
  return 1;
}

function getNextReportStartRowForBlock_(reportSheet, startCol) {
  var blockWidth = 4;
  var lastRow = Math.max(reportSheet.getLastRow(), REPORT_START_ROW);
  if (lastRow < REPORT_START_ROW) return REPORT_START_ROW;

  var rowCount = lastRow - REPORT_START_ROW + 1;
  if (rowCount <= 0) return REPORT_START_ROW;

  var values = reportSheet.getRange(REPORT_START_ROW, startCol, rowCount, blockWidth).getDisplayValues();
  var lastUsedOffset = -1;

  values.forEach(function (row, idx) {
    var hasContent = row.some(function (cell) {
      return String(cell || '').trim() !== '';
    });
    if (hasContent) {
      lastUsedOffset = idx;
    }
  });

  if (lastUsedOffset === -1) {
    return REPORT_START_ROW;
  }

  var lastUsedRow = REPORT_START_ROW + lastUsedOffset;
  return lastUsedRow + 3;
}

function columnToLetter_(columnNumber) {
  var col = Number(columnNumber);
  var letter = '';
  while (col > 0) {
    var temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = Math.floor((col - temp - 1) / 26);
  }
  return letter || 'A';
}

function doOptions() {
  return jsonResponse_({ status: 'ok' });
}

function getSheet_() {
  var spreadsheet = getSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet not found: ' + SHEET_NAME);
  }
  return sheet;
}

function getSpreadsheet_() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function getExcludedCouponCode_() {
  try {
    var ss = getSpreadsheet_();
    var couponSheet = ss.getSheetByName(COUPON_SHEET_NAME);
    if (!couponSheet) return '';
    return safeCell_(couponSheet.getRange(EXCLUDED_COUPON_CELL).getValue());
  } catch (error) {
    return '';
  }
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

function normalizeDateKey_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  var text = safeCell_(value);
  if (!text) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;

  var parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  return text;
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
