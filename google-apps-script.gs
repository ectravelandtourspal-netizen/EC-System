/**
 * GOOGLE APPS SCRIPT BACKEND FOR BOOKING DASHBOARD
 *
 * DEPLOYMENT:
 * 1) Open the Google Sheet -> Extensions -> Apps Script.
 * 2) Paste this code.
 * 3) Deploy -> New deployment -> Web app.
 * 4) Execute as: Me
 * 5) Who has access: Anyone
 * 6) Deploy and copy the /exec URL to script.js API_URL.
 */

const SHEET_NAME = 'Bookings';
const DATA_START_ROW = 5; // Row 4 is header
const DATA_START_COL = 2; // B
const DATA_COL_COUNT = 34; // B:AI

const COL = {
  TRAVEL_DATE: 2, // B
  FIRST_NAME: 3, // C
  LAST_NAME: 4, // D
  PHONE: 6, // F
  NUMBER_OF_GUEST: 10, // J
  TOUR: 11, // K
  COUPON_CODE: 18, // R
  DISCOUNT_AMOUNT: 19, // S
  PAYMENT_METHOD: 20, // T
  TOTAL_AMOUNT: 22, // V
  DOWNPAYMENT: 24, // X
  RECEIPT_DOWNPAYMENT: 25, // Y
  TOTAL_BALANCE: 26, // Z
  RECEIPT_FULL_PAYMENT: 29, // AC
  ADDITIONAL_AD: 30, // AD
  WHATSAPP_NEW: 32, // AF
  REMARKS_NEW: 33, // AG
  WHATSAPP_CONFIRMED: 34, // AH
  REMARKS_CONFIRMED: 35 // AI
};

function doGet() {
  try {
    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();

    if (lastRow < DATA_START_ROW) {
      return jsonResponse_([]);
    }

    const rowCount = lastRow - DATA_START_ROW + 1;
    const rows = sheet.getRange(DATA_START_ROW, DATA_START_COL, rowCount, DATA_COL_COUNT).getValues();
    const waNewRich = sheet.getRange(DATA_START_ROW, COL.WHATSAPP_NEW, rowCount, 1).getRichTextValues();
    const waConfirmedRich = sheet.getRange(DATA_START_ROW, COL.WHATSAPP_CONFIRMED, rowCount, 1).getRichTextValues();

    const payload = rows.map((row, idx) => {
      const waNew = waNewRich[idx] && waNewRich[idx][0] ? waNewRich[idx][0].getLinkUrl() : '';
      const waConfirmed = waConfirmedRich[idx] && waConfirmedRich[idx][0] ? waConfirmedRich[idx][0].getLinkUrl() : '';

      return {
        rowIndex: idx + DATA_START_ROW,
        name: `${safeText_(cell_(row, COL.FIRST_NAME))} ${safeText_(cell_(row, COL.LAST_NAME))}`.trim(),
        phone: safeText_(cell_(row, COL.PHONE)),
        tour: safeText_(cell_(row, COL.TOUR)),
        numberOfGuest: safeText_(cell_(row, COL.NUMBER_OF_GUEST)),
        couponCode: safeText_(cell_(row, COL.COUPON_CODE)),
        discountAmount: safeText_(cell_(row, COL.DISCOUNT_AMOUNT)),
        paymentMethod: safeText_(cell_(row, COL.PAYMENT_METHOD)),
        totalAmount: safeText_(cell_(row, COL.TOTAL_AMOUNT)),
        downpaymentAmount: safeText_(cell_(row, COL.DOWNPAYMENT)),
        receiptDownpayment: safeText_(cell_(row, COL.RECEIPT_DOWNPAYMENT)),
        totalBalance: safeText_(cell_(row, COL.TOTAL_BALANCE)),
        receiptFullPayment: safeText_(cell_(row, COL.RECEIPT_FULL_PAYMENT)),
        adData: safeText_(cell_(row, COL.ADDITIONAL_AD)),
        date: formatDate_(cell_(row, COL.TRAVEL_DATE)),
        whatsappNew: waNew || safeText_(cell_(row, COL.WHATSAPP_NEW)),
        remarksNew: safeText_(cell_(row, COL.REMARKS_NEW)),
        whatsappConfirmed: waConfirmed || safeText_(cell_(row, COL.WHATSAPP_CONFIRMED)),
        remarksConfirmed: safeText_(cell_(row, COL.REMARKS_CONFIRMED))
      };
    });

    return jsonResponse_(payload);
  } catch (err) {
    return jsonResponse_({ status: 'error', message: err.message });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Missing request body.');
    }

    const body = JSON.parse(e.postData.contents);
    const rowIndex = Number(body.rowIndex);

    if (!rowIndex || rowIndex < DATA_START_ROW) {
      throw new Error('Invalid rowIndex');
    }

    const sheet = getSheet_();
    if (rowIndex > sheet.getLastRow()) {
      throw new Error('rowIndex out of range');
    }

    const action = safeText_(body.action || 'mark_sent_new').toLowerCase();

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
  } catch (err) {
    return jsonResponse_({ status: 'error', message: err.message });
  }
}

function saveReceipt_(sheet, rowIndex, body) {
  const paymentType = safeText_(body.paymentType || 'downpayment').toLowerCase();
  const invoiceNumber = safeText_(body.invoiceNumber);
  const encodedDownpaymentAmount = safeText_(body.encodedDownpaymentAmount);
  const encodedFullPaymentAmount = safeText_(body.encodedFullPaymentAmount);

  if (!invoiceNumber) {
    throw new Error('invoiceNumber is required');
  }

  const targetCol = paymentType === 'full' ? COL.RECEIPT_FULL_PAYMENT : COL.RECEIPT_DOWNPAYMENT;
  const targetCell = sheet.getRange(rowIndex, targetCol);
  targetCell.setValue(invoiceNumber);

  if (paymentType === 'downpayment' && encodedDownpaymentAmount) {
    const normalized = normalizeAmount_(encodedDownpaymentAmount);
    if (normalized === null) {
      throw new Error('Invalid encodedDownpaymentAmount');
    }
    const currentDownpayment = normalizeAmount_(sheet.getRange(rowIndex, COL.DOWNPAYMENT).getValue()) || 0;
    if (normalized < currentDownpayment) {
      throw new Error('Downpayment should start at the amount encoded in column X');
    }
    sheet.getRange(rowIndex, COL.DOWNPAYMENT).setValue(normalized);
  }

  if (paymentType === 'full') {
    const normalizedFull = normalizeAmount_(encodedFullPaymentAmount);
    if (normalizedFull === null) {
      throw new Error('encodedFullPaymentAmount is required for full payment');
    }
    const currentBalance = normalizeAmount_(sheet.getRange(rowIndex, COL.TOTAL_BALANCE).getValue()) || 0;
    if (normalizedFull !== currentBalance) {
      throw new Error('Full payment amount must exactly match column Z');
    }
  }

  const imageBase64 = safeText_(body.receiptImageBase64);
  if (imageBase64) {
    const mimeType = safeText_(body.receiptMimeType) || 'image/png';
    const filename = safeText_(body.receiptFileName) || `receipt_${rowIndex}_${Date.now()}.png`;

    const bytes = Utilities.base64Decode(imageBase64);
    const blob = Utilities.newBlob(bytes, mimeType, filename);
    const file = DriveApp.createFile(blob);
    const url = file.getUrl();

    targetCell.setNote(`Receipt image: ${url}`);
    return jsonResponse_({ status: 'success', message: 'Invoice saved with image link note.' });
  }

  return jsonResponse_({ status: 'success', message: 'Invoice saved.' });
}

function doOptions() {
  return jsonResponse_({ status: 'ok' });
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error(`Sheet not found: ${SHEET_NAME}`);
  }
  return sheet;
}

function jsonResponse_(payload) {
  const out = ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);

  if (typeof out.setHeader === 'function') {
    out.setHeader('Access-Control-Allow-Origin', '*');
    out.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
    out.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  }

  return out;
}

function safeText_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function formatDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return safeText_(value);
}

function cell_(row, absoluteCol) {
  const offset = absoluteCol - DATA_START_COL;
  return row[offset];
}

function normalizeAmount_(value) {
  const cleaned = String(value).replace(/[^0-9.-]/g, '').trim();
  if (!cleaned) return null;
  const numeric = Number(cleaned);
  if (!isFinite(numeric) || numeric < 0) return null;
  return Number(numeric.toFixed(2));
}
