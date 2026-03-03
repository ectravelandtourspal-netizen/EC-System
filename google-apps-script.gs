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
const SYSTEM_ACC_SPREADSHEET_ID = '1qEvVWIuVHrFqz4LSVqETfwRGjSS5cwaqSjcYy4ExNNg';
const SYSTEM_ACC_SHEET_NAME = 'SYSTEM ACC WORKING';
const SYSTEM_ACC_START_ROW = 4;
const SYSTEM_ACC_COL_USERNAME = 4; // D
const SYSTEM_ACC_COL_FIRST_PASSWORD = 5; // E
const SYSTEM_ACC_COL_CURRENT_PASSWORD = 6; // F
const SYSTEM_ACC_ACTIVITY_START_ROW = 4;
const SYSTEM_ACC_ACTIVITY_START_COL = 7; // G
const SYSTEM_ACC_ACTIVITY_COL_COUNT = 4; // G:J
const AUTH_PASSWORD_KEY_PREFIX = 'auth_pwd_';
const COUPON_SHEET_NAME = 'COUPONS';
const EXCLUDED_COUPON_CELL = 'A6';
const REPORT_SHEET_NAME = 'INFLUENCER PAYMENT';
const REPORT_START_ROW = 10;
const REPORT_BLOCK_WIDTH = 5;
const REPORT_BLOCK_GAP = 2;
const DATA_START_ROW = 5; // Row 4 is header
const DATA_START_COL = 2; // B
const DATA_COL_COUNT = 40; // B:AO
const PAYMENT_RECEIPT_ROOT_FOLDER_ID = '1gIU8ZtOgCW_Yl40hz1t6ykm2N95Gw030';
const COMMISSION_RECEIPT_ROOT_FOLDER_ID = '1AswkuPdtvGllaKxzdX_cAVThsa3GXZFc';
const FORCE_MYDRIVE_FALLBACK = true;
const PASSWORD_RESET_OWNER_EMAIL = 'micah.tianga.work@gmail.com';
const RESET_TOKEN_KEY_PREFIX = 'reset_token_';
const RESET_TOKEN_TTL_MS = 30 * 60 * 1000;

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
  COMMISSION_AMOUNT_AA: 27, // AA
  RECEIPT_FULL_PAYMENT: 29, // AC
  ADDITIONAL_AD: 30, // AD
  WHATSAPP_NEW: 32, // AF
  REMARKS_NEW: 33, // AG
  WHATSAPP_CONFIRMED: 34, // AH
  REMARKS_CONFIRMED: 35, // AI
  RECEIPT_INVOICE_AK: 37, // AK
  INFLUENCER_STATUS_AL: 38, // AL
  RECEIPT_URL_AM: 39, // AM (Downpayment image URL)
  RECEIPT_URL_AN: 40, // AN (Full payment image URL)
  CANCEL_REASON_AO: 41 // AO (Cancel reason)
};

function doGet(e) {
  try {
    const action = safeText_(e && e.parameter ? e.parameter.action : '').toLowerCase();
    if (action === 'confirm_reset') {
      return confirmPasswordResetFromEmail_(e && e.parameter ? e.parameter : {});
    }

    if (action === 'reject_reset') {
      return rejectPasswordResetFromEmail_(e && e.parameter ? e.parameter : {});
    }

    const sheet = getSheet_();
    const excludedCouponCode = getExcludedCouponCode_();
    const lastRow = sheet.getLastRow();

    if (lastRow < DATA_START_ROW) {
      return jsonResponse_([]);
    }

    const rowCount = lastRow - DATA_START_ROW + 1;
    const bookingDates = sheet.getRange(DATA_START_ROW, 1, rowCount, 1).getValues();
    const rows = sheet.getRange(DATA_START_ROW, DATA_START_COL, rowCount, DATA_COL_COUNT).getValues();
    const waNewRich = sheet.getRange(DATA_START_ROW, COL.WHATSAPP_NEW, rowCount, 1).getRichTextValues();
    const waConfirmedRich = sheet.getRange(DATA_START_ROW, COL.WHATSAPP_CONFIRMED, rowCount, 1).getRichTextValues();

    const payload = rows.map((row, idx) => {
      const waNew = waNewRich[idx] && waNewRich[idx][0] ? waNewRich[idx][0].getLinkUrl() : '';
      const waConfirmed = waConfirmedRich[idx] && waConfirmedRich[idx][0] ? waConfirmedRich[idx][0].getLinkUrl() : '';

      return {
        rowIndex: idx + DATA_START_ROW,
        bookingDate: formatDate_(bookingDates[idx] ? bookingDates[idx][0] : ''),
        name: `${safeText_(cell_(row, COL.FIRST_NAME))} ${safeText_(cell_(row, COL.LAST_NAME))}`.trim(),
        phone: safeText_(cell_(row, COL.PHONE)),
        tour: safeText_(cell_(row, COL.TOUR)),
        numberOfGuest: safeText_(cell_(row, COL.NUMBER_OF_GUEST)),
        couponCode: safeText_(cell_(row, COL.COUPON_CODE)),
        excludedCouponCode: excludedCouponCode,
        discountAmount: safeText_(cell_(row, COL.DISCOUNT_AMOUNT)),
        paymentMethod: safeText_(cell_(row, COL.PAYMENT_METHOD)),
        totalAmount: safeText_(cell_(row, COL.TOTAL_AMOUNT)),
        downpaymentAmount: safeText_(cell_(row, COL.DOWNPAYMENT)),
        receiptDownpayment: safeText_(cell_(row, COL.RECEIPT_DOWNPAYMENT)),
        totalBalance: safeText_(cell_(row, COL.TOTAL_BALANCE)),
        commissionAmountAA: safeText_(cell_(row, COL.COMMISSION_AMOUNT_AA)),
        influencerInvoiceAK: safeText_(cell_(row, COL.RECEIPT_INVOICE_AK)),
        influencerStatusAL: safeText_(cell_(row, COL.INFLUENCER_STATUS_AL)),
        receiptUrlAM: safeText_(cell_(row, COL.RECEIPT_URL_AM)),
        receiptUrlAN: safeText_(cell_(row, COL.RECEIPT_URL_AN)),
        cancelReasonAO: safeText_(cell_(row, COL.CANCEL_REASON_AO)),
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
    return jsonResponse_({
      status: 'error',
      message: err.message,
      where: 'doGet',
      stack: String(err.stack || '')
    });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Missing request body.');
    }

    const body = JSON.parse(e.postData.contents);
    const action = safeText_(body.action || 'mark_sent_new').toLowerCase();

    if (action === 'save_commission_receipt_bulk') {
      const sheet = getSheet_();
      const result = saveCommissionReceiptBulk_(sheet, body);
      logSystemActivity_(body, `Saved commissions receipt bulk (Coupon: ${safeText_(body.couponCode)}, Invoice: ${safeText_(body.invoiceNumber)})`);
      return result;
    }

    if (action === 'login_user') {
      const result = loginUser_(body);
      logSystemActivity_(body, 'Logged in');
      return result;
    }

    if (action === 'set_new_password') {
      const result = setNewPassword_(body);
      logSystemActivity_(body, 'Changed password');
      return result;
    }

    if (action === 'request_password_reset') {
      const result = requestPasswordReset_(body);
      logSystemActivity_(body, 'Requested password reset');
      return result;
    }

    if (action === 'get_commission_payment_summaries') {
      return getCommissionPaymentSummaries_();
    }

    const rowIndex = Number(body.rowIndex);

    if (!rowIndex || rowIndex < DATA_START_ROW) {
      throw new Error('Invalid rowIndex');
    }

    const sheet = getSheet_();
    if (rowIndex > sheet.getLastRow()) {
      throw new Error('rowIndex out of range');
    }

    if (action === 'mark_sent_new') {
      sheet.getRange(rowIndex, COL.REMARKS_NEW).setValue('sent');
      logSystemActivity_(body, `Sent WhatsApp (New) for row ${rowIndex}`);
      return jsonResponse_({ status: 'success' });
    }

    if (action === 'mark_sent_confirmed') {
      sheet.getRange(rowIndex, COL.REMARKS_CONFIRMED).setValue('sent');
      logSystemActivity_(body, `Sent WhatsApp (Confirmed) for row ${rowIndex}`);
      return jsonResponse_({ status: 'success' });
    }

    if (action === 'save_receipt') {
      const result = saveReceipt_(sheet, rowIndex, body);
      const paymentType = safeText_(body.paymentType || 'downpayment').toUpperCase();
      logSystemActivity_(body, `Saved ${paymentType} receipt (row ${rowIndex}, Invoice: ${safeText_(body.invoiceNumber)})`);
      return result;
    }

    if (action === 'save_cancel_reason') {
      const result = saveCancelReason_(sheet, rowIndex, body);
      logSystemActivity_(body, `Saved cancel reason (row ${rowIndex})`);
      return result;
    }

    throw new Error('Unsupported action: ' + action);
  } catch (err) {
    return jsonResponse_({
      status: 'error',
      message: err.message,
      where: 'doPost',
      stack: String(err.stack || '')
    });
  }
}

function saveReceipt_(sheet, rowIndex, body) {
  const paymentType = safeText_(body.paymentType || 'downpayment').toLowerCase();
  const invoiceNumber = safeText_(body.invoiceNumber);
  const encodedDownpaymentAmount = safeText_(body.encodedDownpaymentAmount);
  const encodedFullPaymentAmount = safeText_(body.encodedFullPaymentAmount);
  const receiptImageUrl = safeText_(body.receiptImageUrl);

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

  if (paymentType === 'downpayment') {
    sheet.getRange(rowIndex, COL.RECEIPT_URL_AM).setValue(receiptImageUrl);
  }

  if (paymentType === 'full') {
    sheet.getRange(rowIndex, COL.RECEIPT_URL_AN).setValue(receiptImageUrl);
  }

  return jsonResponse_({ status: 'success', message: 'Invoice saved.' });
}

function saveCancelReason_(sheet, rowIndex, body) {
  const cancelReason = safeText_(body.cancelReason);
  if (!cancelReason) {
    throw new Error('cancelReason is required');
  }

  sheet.getRange(rowIndex, COL.CANCEL_REASON_AO).setValue(cancelReason);
  return jsonResponse_({ status: 'success', message: 'Cancel reason saved.' });
}

function saveCommissionReceiptBulk_(sheet, body) {
  const travelDatesRaw = Array.isArray(body.travelDates) ? body.travelDates : [body.travelDate];
  const targetDates = [...new Set(travelDatesRaw.map((value) => normalizeDateKey_(value)).filter(Boolean))];
  const couponCode = safeText_(body.couponCode);
  const invoiceNumber = safeText_(body.invoiceNumber);
  const encodedCommissionAmount = normalizeAmount_(body.encodedCommissionAmount);
  const receiptImageUrl = safeText_(body.receiptImageUrl);

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

  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) {
    throw new Error('No booking rows found.');
  }

  const rowCount = lastRow - DATA_START_ROW + 1;
  const rows = sheet.getRange(DATA_START_ROW, DATA_START_COL, rowCount, DATA_COL_COUNT).getValues();

  let matchedRows = [];
  let sumAA = 0;
  const dateSummary = new Map();

  rows.forEach((row, idx) => {
    const rowDate = normalizeDateKey_(cell_(row, COL.TRAVEL_DATE));
    const rowCoupon = safeText_(cell_(row, COL.COUPON_CODE));
    const rowStatusAL = safeText_(cell_(row, COL.INFLUENCER_STATUS_AL)).toUpperCase();
    const rowStatusAD = safeText_(cell_(row, COL.ADDITIONAL_AD)).toUpperCase();

    if (targetDates.includes(rowDate) && rowCoupon === couponCode && rowStatusAL !== 'PAID' && isCommissionBookingCompleted_(rowStatusAD) && !isBlockedCommissionStatus_(rowStatusAD)) {
      const sheetRow = idx + DATA_START_ROW;
      const amountAA = normalizeAmount_(cell_(row, COL.COMMISSION_AMOUNT_AA)) || 0;
      const guestCount = normalizeGuestCount_(cell_(row, COL.NUMBER_OF_GUEST));
      const rowTotal = Number((amountAA * guestCount).toFixed(2));
      matchedRows.push(sheetRow);
      sumAA += rowTotal;

      if (!dateSummary.has(rowDate)) {
        dateSummary.set(rowDate, { bookings: 0, totalAmount: 0 });
      }
      const current = dateSummary.get(rowDate);
      current.bookings += guestCount;
      current.totalAmount = Number((current.totalAmount + rowTotal).toFixed(2));
    }
  });

  sumAA = Number(sumAA.toFixed(2));

  if (!matchedRows.length) {
    throw new Error('No matching bookings found for selected travel date and coupon code.');
  }

  if (sumAA !== encodedCommissionAmount) {
    throw new Error(`Encoded amount must exactly match Total Amount (AA × Guests) (${sumAA.toFixed(2)}).`);
  }

  const noteLines = [
    `Commissions receipt encoded`,
    `Travel Dates: ${targetDates.join(', ')}`,
    `Coupon Code: ${couponCode}`,
    `Invoice Number: ${invoiceNumber}`,
    `Encoded Amount: ${encodedCommissionAmount.toFixed(2)}`,
    `Matched Rows: ${matchedRows.join(', ')}`,
    `Saved At: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')}`
  ];

  if (receiptImageUrl) {
    noteLines.push(`Receipt URL: ${receiptImageUrl}`);
  }

  const noteText = noteLines.join('\n');
  matchedRows.forEach((sheetRow) => {
    sheet.getRange(sheetRow, COL.RECEIPT_INVOICE_AK).setValue(invoiceNumber);
    sheet.getRange(sheetRow, COL.COMMISSION_AMOUNT_AA).setNote(noteText);
  });

  appendInfluencerPaymentReport_(targetDates, dateSummary, invoiceNumber, couponCode, receiptImageUrl);

  return jsonResponse_({
    status: 'success',
    message: `Commissions receipt saved for ${matchedRows.length} booking(s).`,
    matchedRows: matchedRows,
    sumAA: sumAA
  });
}

function loginUser_(body) {
  const username = safeText_(body.username).toUpperCase();
  const password = safeText_(body.password);

  if (!username || !password) {
    throw new Error('username and password are required');
  }

  const account = findSystemAccount_(username);
  if (!account) {
    throw new Error('Invalid username or password.');
  }

  const storedHash = getStoredPasswordHash_(username);
  const hashMatched = storedHash && storedHash === hashPassword_(password);

  if (storedHash && !hashMatched) {
    throw new Error('Invalid username or password.');
  }

  if (!storedHash) {
    const legacyCurrentPassword = safeText_(account.legacyCurrentPassword);
    const matchedFirst = password === account.firstPassword;
    const matchedLegacy = legacyCurrentPassword && password === legacyCurrentPassword;

    if (!matchedFirst && !matchedLegacy) {
      throw new Error('Invalid username or password.');
    }

    if (matchedLegacy && password !== account.firstPassword) {
      setStoredPasswordHash_(username, password);
      clearLegacyCurrentPassword_(account);
    }
  }

  const mustChangePassword = !storedHash && password === account.firstPassword;
  return jsonResponse_({
    status: 'success',
    username: username,
    mustChangePassword: mustChangePassword
  });
}

function setNewPassword_(body) {
  const username = safeText_(body.username).toUpperCase();
  const currentPassword = safeText_(body.currentPassword);
  const newPassword = safeText_(body.newPassword);

  if (!username || !currentPassword || !newPassword) {
    throw new Error('username, currentPassword, and newPassword are required');
  }

  if (newPassword.length < 4) {
    throw new Error('New password must be at least 4 characters.');
  }

  const account = findSystemAccount_(username);
  if (!account) {
    throw new Error('User not found.');
  }

  if (!validateCurrentPassword_(account, currentPassword)) {
    throw new Error('Current password is incorrect.');
  }

  if (newPassword === account.firstPassword) {
    throw new Error('New password must be different from first password.');
  }

  setStoredPasswordHash_(username, newPassword);
  clearLegacyCurrentPassword_(account);
  return jsonResponse_({ status: 'success', message: 'New password saved.' });
}

function requestPasswordReset_(body) {
  const username = safeText_(body.username).toUpperCase();
  if (!username) {
    throw new Error('username is required');
  }

  const account = findSystemAccount_(username);
  if (!account) {
    throw new Error('User not found.');
  }

  const token = Utilities.getUuid();
  const expiresAt = Date.now() + RESET_TOKEN_TTL_MS;
  const payload = JSON.stringify({ username: username, expiresAt: expiresAt });
  PropertiesService.getScriptProperties().setProperty(RESET_TOKEN_KEY_PREFIX + token, payload);

  const webAppUrl = ScriptApp.getService().getUrl();
  if (!webAppUrl) {
    throw new Error('Cannot build reset URL. Redeploy as Web App first.');
  }

  const confirmUrl = `${webAppUrl}?action=confirm_reset&username=${encodeURIComponent(username)}&token=${encodeURIComponent(token)}`;
  const rejectUrl = `${webAppUrl}?action=reject_reset&username=${encodeURIComponent(username)}&token=${encodeURIComponent(token)}`;

  const ownerEmail = safeText_(PASSWORD_RESET_OWNER_EMAIL) || safeText_(Session.getEffectiveUser().getEmail());
  if (!ownerEmail) {
    throw new Error('Owner email is not available.');
  }

  try {
    MailApp.sendEmail({
      to: ownerEmail,
      subject: `[EC APP] Password Reset Confirmation for ${username}`,
      htmlBody: `
        <p>Password reset requested for user: <strong>${username}</strong>.</p>
        <p>Choose one action below (valid for 30 minutes):</p>
        <p>
          <a href="${confirmUrl}" style="display:inline-block;padding:10px 14px;border-radius:8px;background:#0f9f57;color:#fff;text-decoration:none;font-weight:600;margin-right:8px;">Confirm Reset</a>
          <a href="${rejectUrl}" style="display:inline-block;padding:10px 14px;border-radius:8px;background:#dc2626;color:#fff;text-decoration:none;font-weight:600;">Reject Reset</a>
        </p>
        <p>If buttons do not work, use links:</p>
        <p>Confirm: <a href="${confirmUrl}">${confirmUrl}</a></p>
        <p>Reject: <a href="${rejectUrl}">${rejectUrl}</a></p>
        <p>After confirmation, password will be set back to first password and user must create a new password on next login.</p>
      `
    });
  } catch (err) {
    throw new Error(
      `Reset email permission is not authorized yet. In Apps Script editor, run authorizeMailAccess() once and approve access, then redeploy Web App. Details: ${safeText_(err && err.message ? err.message : err)}`
    );
  }

  return jsonResponse_({ status: 'success', message: 'Reset request sent to owner email.' });
}

function confirmPasswordResetFromEmail_(params) {
  const username = safeText_(params.username).toUpperCase();
  const token = safeText_(params.token);

  if (!username || !token) {
    return HtmlService.createHtmlOutput('<h3>Invalid reset request.</h3>');
  }

  const key = RESET_TOKEN_KEY_PREFIX + token;
  const payloadText = PropertiesService.getScriptProperties().getProperty(key);
  if (!payloadText) {
    return HtmlService.createHtmlOutput('<h3>Reset link is invalid or expired.</h3>');
  }

  let payload;
  try {
    payload = JSON.parse(payloadText);
  } catch {
    PropertiesService.getScriptProperties().deleteProperty(key);
    return HtmlService.createHtmlOutput('<h3>Reset request data is invalid.</h3>');
  }

  if (safeText_(payload.username).toUpperCase() !== username) {
    return HtmlService.createHtmlOutput('<h3>Reset link does not match this user.</h3>');
  }

  if (Number(payload.expiresAt || 0) < Date.now()) {
    PropertiesService.getScriptProperties().deleteProperty(key);
    return HtmlService.createHtmlOutput('<h3>Reset link expired.</h3>');
  }

  const account = findSystemAccount_(username);
  if (!account) {
    return HtmlService.createHtmlOutput('<h3>User not found.</h3>');
  }

  clearStoredPasswordHash_(username);
  clearLegacyCurrentPassword_(account);
  PropertiesService.getScriptProperties().deleteProperty(key);

  return HtmlService.createHtmlOutput('<h3>Password reset confirmed.</h3><p>User password is set back to first password and must create a new password at next login.</p>');
}

function rejectPasswordResetFromEmail_(params) {
  const username = safeText_(params.username).toUpperCase();
  const token = safeText_(params.token);

  if (!username || !token) {
    return HtmlService.createHtmlOutput('<h3>Invalid reject request.</h3>');
  }

  const key = RESET_TOKEN_KEY_PREFIX + token;
  const payloadText = PropertiesService.getScriptProperties().getProperty(key);
  if (!payloadText) {
    return HtmlService.createHtmlOutput('<h3>Reset link is invalid or expired.</h3>');
  }

  let payload;
  try {
    payload = JSON.parse(payloadText);
  } catch {
    PropertiesService.getScriptProperties().deleteProperty(key);
    return HtmlService.createHtmlOutput('<h3>Reset request data is invalid.</h3>');
  }

  if (safeText_(payload.username).toUpperCase() !== username) {
    return HtmlService.createHtmlOutput('<h3>Reset link does not match this user.</h3>');
  }

  PropertiesService.getScriptProperties().deleteProperty(key);
  return HtmlService.createHtmlOutput('<h3>Password reset rejected.</h3><p>No changes were made.</p>');
}

function appendInfluencerPaymentReport_(targetDates, dateSummary, invoiceNumber, couponCode, receiptImageUrl) {
  const spreadsheet = getSpreadsheet_();
  let reportSheet = spreadsheet.getSheetByName(REPORT_SHEET_NAME);
  if (!reportSheet) {
    reportSheet = spreadsheet.insertSheet(REPORT_SHEET_NAME);
  }

  backfillInfluencerPaymentReport_(reportSheet);

  const startCol = getCouponBlockStartCol_(reportSheet, couponCode);
  const sortedDates = [...targetDates].sort((a, b) => a.localeCompare(b));
  const dataRows = sortedDates
    .filter((dateKey) => dateSummary.has(dateKey))
    .map((dateKey) => {
      const info = dateSummary.get(dateKey);
      const bookings = info.bookings || 0;
      const totalAmount = Number(info.totalAmount || 0);
      const amount = bookings ? Number((totalAmount / bookings).toFixed(2)) : 0;
      return [dateKey, bookings, amount];
    });

  if (!dataRows.length) {
    return;
  }

  const startRow = REPORT_START_ROW;
  const rowsToInsert = dataRows.length + 5;
  shiftReportBlockDown_(reportSheet, startCol, rowsToInsert);
  const datePaidText = formatReportDatePaid_(new Date());

  const endCol = startCol + 3;
  const urlCol = startCol + 4;
  const titleRow = startRow;
  const headerRow = startRow + 1;

  reportSheet.getRange(titleRow, startCol, 1, REPORT_BLOCK_WIDTH).breakApart();
  reportSheet.getRange(titleRow, startCol, 1, REPORT_BLOCK_WIDTH).merge();
  reportSheet.getRange(titleRow, startCol).setValue(`SUMMARY OF PAYMENT WITH INVOICE # ${invoiceNumber}`);
  reportSheet.getRange(titleRow, startCol).setFontWeight('bold');

  reportSheet.getRange(headerRow, startCol, 1, REPORT_BLOCK_WIDTH).setValues([['TRAVEL DATES', 'NO. OF BOOKINGS', 'AMOUNT', 'TOTAL', 'DATE PAID/RECEIPT']]);
  reportSheet.getRange(headerRow, startCol, 1, REPORT_BLOCK_WIDTH).setFontWeight('bold');

  const dataStartRow = startRow + 2;
  reportSheet.getRange(dataStartRow, startCol, dataRows.length, 3).setValues(dataRows);

  const totalFormulas = dataRows.map((_, idx) => {
    const row = dataStartRow + idx;
    const bookingsCol = columnToLetter_(startCol + 1);
    const amountCol = columnToLetter_(startCol + 2);
    return [`=${bookingsCol}${row}*${amountCol}${row}`];
  });
  reportSheet.getRange(dataStartRow, endCol, dataRows.length, 1).setFormulas(totalFormulas);
  reportSheet.getRange(dataStartRow, urlCol, dataRows.length, 1).clearContent();
  reportSheet.getRange(dataStartRow, urlCol).setValue(datePaidText);

  const grandTotalRow = dataStartRow + dataRows.length;
  const totalColLetter = columnToLetter_(endCol);
  reportSheet.getRange(grandTotalRow, startCol + 2).setValue('GRAND TOTAL');
  reportSheet.getRange(grandTotalRow, endCol).setFormula(`=SUM(${totalColLetter}${dataStartRow}:${totalColLetter}${grandTotalRow - 1})`);
  reportSheet.getRange(grandTotalRow, startCol + 2, 1, 2).setFontWeight('bold');
  reportSheet.getRange(grandTotalRow, urlCol).setValue(safeText_(receiptImageUrl));
}

function shiftReportBlockDown_(reportSheet, startCol, rowsToInsert) {
  const rowsToShift = Number(rowsToInsert) || 0;
  if (rowsToShift <= 0) {
    return;
  }

  const lastUsedRow = getLastUsedRowForBlock_(reportSheet, startCol);
  if (lastUsedRow < REPORT_START_ROW) {
    return;
  }

  const rowCount = lastUsedRow - REPORT_START_ROW + 1;
  const sourceRange = reportSheet.getRange(REPORT_START_ROW, startCol, rowCount, REPORT_BLOCK_WIDTH);
  const existingValues = sourceRange.getValues();

  sourceRange.clearContent();

  const requiredLastRow = REPORT_START_ROW + rowsToShift + rowCount - 1;
  const maxRows = reportSheet.getMaxRows();
  if (requiredLastRow > maxRows) {
    reportSheet.insertRowsAfter(maxRows, requiredLastRow - maxRows);
  }

  reportSheet.getRange(REPORT_START_ROW + rowsToShift, startCol, rowCount, REPORT_BLOCK_WIDTH).setValues(existingValues);
}

function getLastUsedRowForBlock_(reportSheet, startCol) {
  const lastRow = Math.max(reportSheet.getLastRow(), REPORT_START_ROW);
  if (lastRow < REPORT_START_ROW) {
    return REPORT_START_ROW - 1;
  }

  const rowCount = lastRow - REPORT_START_ROW + 1;
  if (rowCount <= 0) {
    return REPORT_START_ROW - 1;
  }

  const values = reportSheet.getRange(REPORT_START_ROW, startCol, rowCount, REPORT_BLOCK_WIDTH).getDisplayValues();
  for (let i = values.length - 1; i >= 0; i -= 1) {
    if (values[i].some((cell) => String(cell || '').trim() !== '')) {
      return REPORT_START_ROW + i;
    }
  }

  return REPORT_START_ROW - 1;
}

function getCommissionPaymentSummaries_() {
  const spreadsheet = getSpreadsheet_();
  const reportSheet = spreadsheet.getSheetByName(REPORT_SHEET_NAME);

  if (!reportSheet) {
    return jsonResponse_({ status: 'success', groups: [] });
  }

  backfillInfluencerPaymentReport_(reportSheet);

  const lastCol = Math.max(reportSheet.getLastColumn(), 1);
  const lastRow = Math.max(reportSheet.getLastRow(), REPORT_START_ROW);
  const rowCount = Math.max(lastRow - REPORT_START_ROW + 1, 1);
  const markerRow = reportSheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const markerCols = [];
  markerRow.forEach((value, idx) => {
    if (safeText_(value)) {
      markerCols.push(idx + 1);
    }
  });

  const groups = [];

  markerCols.forEach((startCol) => {
    const couponCode = safeText_(reportSheet.getRange(1, startCol).getValue());
    if (!couponCode) return;

    const values = reportSheet.getRange(REPORT_START_ROW, startCol, rowCount, REPORT_BLOCK_WIDTH).getDisplayValues();
    const payments = [];

    for (let i = 0; i < values.length; i += 1) {
      const title = safeText_(values[i][0]);
      if (!title.startsWith('SUMMARY OF PAYMENT WITH INVOICE #')) continue;

      const invoiceNumber = safeText_(title.replace('SUMMARY OF PAYMENT WITH INVOICE #', ''));
      const rows = [];
      let grandTotal = '';
      let receiptUrl = '';

      for (let j = i + 2; j < values.length; j += 1) {
        const line = values[j];
        const dateText = safeText_(line[0]);
        const bookingsText = safeText_(line[1]);
        const amountText = safeText_(line[2]);
        const totalText = safeText_(line[3]);
        const receiptText = safeText_(line[4]);

        if (amountText.toUpperCase() === 'GRAND TOTAL') {
          grandTotal = totalText;
          receiptUrl = receiptText;
          i = j;
          break;
        }

        if (dateText || bookingsText || amountText || totalText) {
          rows.push({
            date: dateText,
            bookings: bookingsText,
            amount: amountText,
            total: totalText
          });
        }
      }

      payments.push({
        title,
        invoiceNumber,
        rows,
        grandTotal,
        receiptUrl
      });
    }

    groups.push({
      couponCode,
      payments
    });
  });

  groups.sort((a, b) => a.couponCode.localeCompare(b.couponCode));
  return jsonResponse_({ status: 'success', groups });
}

function getCouponBlockStartCol_(reportSheet, couponCode) {
  const normalizedCoupon = safeText_(couponCode);
  const blockWidth = REPORT_BLOCK_WIDTH;
  const blockGap = REPORT_BLOCK_GAP;
  const stride = blockWidth + blockGap;

  const lastCol = Math.max(reportSheet.getLastColumn(), 1);
  for (let col = 1; col <= lastCol; col += 1) {
    const marker = safeText_(reportSheet.getRange(1, col).getValue());
    if (marker === normalizedCoupon) {
      return col;
    }
  }

  const markerRow = reportSheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const occupiedCols = [];
  markerRow.forEach((value, idx) => {
    if (safeText_(value)) {
      occupiedCols.push(idx + 1);
    }
  });

  const nextCol = occupiedCols.length ? Math.max(...occupiedCols) + stride : 1;
  reportSheet.getRange(1, nextCol).setValue(normalizedCoupon);
  return nextCol;
}

function backfillInfluencerPaymentReport_(reportSheet) {
  if (!reportSheet) return;

  const lastCol = Math.max(reportSheet.getLastColumn(), 1);
  const lastRow = Math.max(reportSheet.getLastRow(), REPORT_START_ROW);
  if (lastRow < REPORT_START_ROW) return;

  const markerRow = reportSheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const markerCols = [];
  markerRow.forEach((value, idx) => {
    if (safeText_(value)) {
      markerCols.push(idx + 1);
    }
  });
  if (!markerCols.length) return;

  const rowCount = Math.max(lastRow - REPORT_START_ROW + 1, 1);
  const todayText = formatReportDatePaid_(new Date());

  markerCols.forEach((startCol) => {
    const values = reportSheet.getRange(REPORT_START_ROW, startCol, rowCount, REPORT_BLOCK_WIDTH).getDisplayValues();

    for (let i = 0; i < values.length; i += 1) {
      const title = safeText_(values[i][0]);
      if (!title.startsWith('SUMMARY OF PAYMENT WITH INVOICE #')) continue;

      const headerSheetRow = REPORT_START_ROW + i + 1;
      const currentHeader = reportSheet.getRange(headerSheetRow, startCol, 1, REPORT_BLOCK_WIDTH).getDisplayValues()[0];
      const normalizedHeader = currentHeader.map((cell) => safeText_(cell).toUpperCase());
      const needsHeaderUpdate = normalizedHeader[0] !== 'TRAVEL DATES' || normalizedHeader[4] !== 'DATE PAID/RECEIPT';
      if (needsHeaderUpdate) {
        reportSheet.getRange(headerSheetRow, startCol, 1, REPORT_BLOCK_WIDTH).setValues([['TRAVEL DATES', 'NO. OF BOOKINGS', 'AMOUNT', 'TOTAL', 'DATE PAID/RECEIPT']]);
        reportSheet.getRange(headerSheetRow, startCol, 1, REPORT_BLOCK_WIDTH).setFontWeight('bold');
      }

      const dataStartOffset = i + 2;
      if (dataStartOffset >= values.length) {
        continue;
      }

      for (let j = dataStartOffset; j < values.length; j += 1) {
        const line = values[j];
        const amountText = safeText_(line[2]).toUpperCase();
        if (amountText === 'GRAND TOTAL') {
          i = j;
          break;
        }

        const hasData = safeText_(line[0]) || safeText_(line[1]) || safeText_(line[2]) || safeText_(line[3]);
        if (!hasData) {
          continue;
        }

        const datePaidCell = safeText_(line[4]);
        if (!datePaidCell) {
          reportSheet.getRange(REPORT_START_ROW + j, startCol + 4).setValue(todayText);
        }
        break;
      }
    }
  });
}

function getNextReportStartRowForBlock_(reportSheet, startCol) {
  const blockWidth = REPORT_BLOCK_WIDTH;
  const lastRow = Math.max(reportSheet.getLastRow(), REPORT_START_ROW);
  if (lastRow < REPORT_START_ROW) return REPORT_START_ROW;

  const rowCount = lastRow - REPORT_START_ROW + 1;
  if (rowCount <= 0) return REPORT_START_ROW;

  const values = reportSheet.getRange(REPORT_START_ROW, startCol, rowCount, blockWidth).getDisplayValues();
  let lastUsedOffset = -1;

  values.forEach((row, idx) => {
    if (row.some((cell) => String(cell || '').trim() !== '')) {
      lastUsedOffset = idx;
    }
  });

  if (lastUsedOffset === -1) {
    return REPORT_START_ROW;
  }

  const lastUsedRow = REPORT_START_ROW + lastUsedOffset;
  return lastUsedRow + 3;
}

function columnToLetter_(columnNumber) {
  let col = Number(columnNumber);
  let letter = '';
  while (col > 0) {
    const temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = Math.floor((col - temp - 1) / 26);
  }
  return letter || 'A';
}

function doOptions() {
  return jsonResponse_({ status: 'ok' });
}

function getSheet_() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error(`Sheet not found: ${SHEET_NAME}`);
  }
  return sheet;
}

function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSystemAccSheet_() {
  const ss = SpreadsheetApp.openById(SYSTEM_ACC_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SYSTEM_ACC_SHEET_NAME);
  if (!sheet) {
    throw new Error(`Sheet not found: ${SYSTEM_ACC_SHEET_NAME}`);
  }
  return sheet;
}

function getActorUsernameFromBody_(body) {
  const actor = safeText_(body && body.actorUsername ? body.actorUsername : (body && body.username ? body.username : ''));
  return actor ? actor.toUpperCase() : 'UNKNOWN';
}

function logSystemActivity_(body, activityText) {
  try {
    const username = getActorUsernameFromBody_(body || {});
    prependSystemActivityLog_(username, activityText);
  } catch (err) {
    Logger.log(`Activity log failed: ${safeText_(err && err.message ? err.message : err)}`);
  }
}

function prependSystemActivityLog_(username, activityText) {
  const sheet = getSystemAccSheet_();
  const startRow = SYSTEM_ACC_ACTIVITY_START_ROW;
  const startCol = SYSTEM_ACC_ACTIVITY_START_COL;
  const width = SYSTEM_ACC_ACTIVITY_COL_COUNT;

  const lastRow = Math.max(sheet.getLastRow(), startRow);
  const existingRowCount = Math.max(lastRow - startRow + 1, 0);
  const existing = existingRowCount > 0
    ? sheet.getRange(startRow, startCol, existingRowCount, width).getValues()
    : [];

  const now = new Date();
  const dateText = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const timeText = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
  const actor = safeText_(username).toUpperCase() || 'UNKNOWN';
  const activity = safeText_(activityText) || 'Activity';

  const valuesToWrite = [[dateText, timeText, actor, activity], ...existing];

  const requiredLastRow = startRow + valuesToWrite.length - 1;
  const maxRows = sheet.getMaxRows();
  if (requiredLastRow > maxRows) {
    sheet.insertRowsAfter(maxRows, requiredLastRow - maxRows);
  }

  sheet.getRange(startRow, startCol, valuesToWrite.length, width).setValues(valuesToWrite);
}

function findSystemAccount_(username) {
  const normalized = safeText_(username).toUpperCase();
  if (!normalized) return null;

  const sheet = getSystemAccSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < SYSTEM_ACC_START_ROW) return null;

  const rowCount = lastRow - SYSTEM_ACC_START_ROW + 1;
  const values = sheet.getRange(SYSTEM_ACC_START_ROW, SYSTEM_ACC_COL_USERNAME, rowCount, 3).getDisplayValues();
  for (let idx = 0; idx < values.length; idx += 1) {
    const row = values[idx];
    const rowUsername = safeText_(row[0]).toUpperCase();
    if (rowUsername !== normalized) continue;

    const firstPassword = safeText_(row[1]);
    const legacyCurrentPassword = safeText_(row[2]);
    return {
      sheet,
      row: idx + SYSTEM_ACC_START_ROW,
      username: rowUsername,
      firstPassword,
      legacyCurrentPassword
    };
  }

  return null;
}

function getAuthPasswordKey_(username) {
  return `${AUTH_PASSWORD_KEY_PREFIX}${safeText_(username).toUpperCase()}`;
}

function hashPassword_(password) {
  const text = safeText_(password);
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return bytes.map((b) => {
    const normalized = b < 0 ? b + 256 : b;
    return normalized.toString(16).padStart(2, '0');
  }).join('');
}

function getStoredPasswordHash_(username) {
  return safeText_(PropertiesService.getScriptProperties().getProperty(getAuthPasswordKey_(username)));
}

function setStoredPasswordHash_(username, password) {
  PropertiesService.getScriptProperties().setProperty(getAuthPasswordKey_(username), hashPassword_(password));
}

function clearStoredPasswordHash_(username) {
  PropertiesService.getScriptProperties().deleteProperty(getAuthPasswordKey_(username));
}

function clearLegacyCurrentPassword_(account) {
  if (!account || !account.sheet || !account.row) return;
  account.sheet.getRange(account.row, SYSTEM_ACC_COL_CURRENT_PASSWORD).setValue('');
}

function validateCurrentPassword_(account, currentPassword) {
  const username = safeText_(account && account.username);
  const supplied = safeText_(currentPassword);
  if (!username || !supplied) return false;

  const storedHash = getStoredPasswordHash_(username);
  if (storedHash) {
    return storedHash === hashPassword_(supplied);
  }

  const legacyCurrentPassword = safeText_(account.legacyCurrentPassword);
  if (legacyCurrentPassword && supplied === legacyCurrentPassword) {
    return true;
  }

  return supplied === safeText_(account.firstPassword);
}

function getExcludedCouponCode_() {
  try {
    const ss = getSpreadsheet_();
    const couponSheet = ss.getSheetByName(COUPON_SHEET_NAME);
    if (!couponSheet) return '';
    return safeText_(couponSheet.getRange(EXCLUDED_COUPON_CELL).getValue());
  } catch {
    return '';
  }
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

function formatReportDatePaid_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'MMM. dd, yyyy').toUpperCase();
  }

  const parsed = new Date(value);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'MMM. dd, yyyy').toUpperCase();
  }

  return '';
}

function normalizeDateKey_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  const text = safeText_(value);
  if (!text) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;

  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  return text;
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

function normalizeGuestCount_(value) {
  const cleaned = String(value || '').replace(/[^0-9]/g, '').trim();
  const numeric = Number(cleaned);
  if (!isFinite(numeric) || numeric <= 0) return 1;
  return Math.floor(numeric);
}

function isBlockedCommissionStatus_(statusText) {
  const normalized = safeText_(statusText).toUpperCase();
  return normalized === 'CANCEL' || normalized === 'CANCELED' || normalized === 'CANCELLED' || normalized === 'PENDING PAYMENT';
}

function isCommissionBookingCompleted_(statusText) {
  const normalized = safeText_(statusText).toUpperCase();
  return normalized === 'BOOKING COMPLETED' || normalized === 'COMPLETED';
}

function getOrCreateChildFolder_(parentFolder, folderName) {
  const cleanName = safeText_(folderName);
  if (!cleanName) {
    throw new Error('Folder name is required');
  }

  const existing = parentFolder.getFoldersByName(cleanName);
  if (existing.hasNext()) {
    return existing.next();
  }

  return parentFolder.createFolder(cleanName);
}

function getMonthFolderName_(travelDateValue) {
  const dateKey = normalizeDateKey_(travelDateValue);
  if (/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) {
    return dateKey.slice(0, 7);
  }

  const parsed = new Date(dateKey || travelDateValue);
  const finalDate = !isNaN(parsed.getTime()) ? parsed : new Date();
  return Utilities.formatDate(finalDate, Session.getScriptTimeZone(), 'yyyy-MM');
}

function extensionFromMimeType_(mimeType, originalFileName) {
  const name = safeText_(originalFileName);
  const dotIndex = name.lastIndexOf('.');
  if (dotIndex > -1 && dotIndex < name.length - 1) {
    return name.slice(dotIndex + 1).toLowerCase();
  }

  const normalized = safeText_(mimeType).toLowerCase();
  const map = {
    'image/jpeg': 'jpg',
    'image/jpg': 'jpg',
    'image/png': 'png',
    'image/gif': 'gif',
    'image/webp': 'webp',
    'image/bmp': 'bmp',
    'image/tiff': 'tiff',
    'image/heic': 'heic',
    'image/heif': 'heif'
  };

  return map[normalized] || 'png';
}

function sanitizeFileName_(value) {
  const text = safeText_(value);
  if (!text) return '';
  return text.replace(/[\\/:*?"<>|]+/g, '_').replace(/\s+/g, ' ').trim();
}

function resolveRootFolderWithFallback_(preferredFolderId, fallbackFolderName) {
  if (FORCE_MYDRIVE_FALLBACK) {
    const fallback = getOrCreateChildFolder_(DriveApp.getRootFolder(), fallbackFolderName);
    return {
      folder: fallback,
      usedFallback: true,
      warning: `FORCE_MYDRIVE_FALLBACK is enabled. Using fallback folder: ${fallback.getName()}.`
    };
  }

  try {
    const preferred = DriveApp.getFolderById(safeText_(preferredFolderId));
    return {
      folder: preferred,
      usedFallback: false,
      warning: ''
    };
  } catch (err) {
    const fallback = getOrCreateChildFolder_(DriveApp.getRootFolder(), fallbackFolderName);
    return {
      folder: fallback,
      usedFallback: true,
      warning: `No access to configured folder ${preferredFolderId}. Using fallback folder: ${fallback.getName()}.`
    };
  }
}

function ensureDriveAccess_() {
  try {
    DriveApp.getRootFolder().getId();
  } catch (err) {
    throw new Error(`Drive authorization required: ${err.message}. Open Apps Script editor, run authorizeDriveAccess_ once, approve permissions, then redeploy web app.`);
  }
}

function authorizeDriveAccess_() {
  const rootId = DriveApp.getRootFolder().getId();
  const paymentFolder = DriveApp.getFolderById(PAYMENT_RECEIPT_ROOT_FOLDER_ID).getName();
  const commissionFolder = DriveApp.getFolderById(COMMISSION_RECEIPT_ROOT_FOLDER_ID).getName();
  return {
    rootId,
    paymentFolder,
    commissionFolder,
    status: 'Drive access authorized'
  };
}

function authorizeDriveAccess() {
  return authorizeDriveAccess_();
}

function authorizeMailAccess_() {
  const quota = MailApp.getRemainingDailyQuota();
  return {
    quota,
    status: 'Mail access authorized'
  };
}

function authorizeMailAccess() {
  return authorizeMailAccess_();
}

function diagnoseDriveFolderAccess() {
  const result = {
    effectiveUser: '',
    paymentReceiptFolderId: PAYMENT_RECEIPT_ROOT_FOLDER_ID,
    commissionReceiptFolderId: COMMISSION_RECEIPT_ROOT_FOLDER_ID,
    paymentReceiptFolder: null,
    commissionReceiptFolder: null,
    status: 'ok'
  };

  try {
    result.effectiveUser = Session.getEffectiveUser().getEmail() || '';
  } catch (err) {
    result.effectiveUser = '';
  }

  try {
    const folder = DriveApp.getFolderById(PAYMENT_RECEIPT_ROOT_FOLDER_ID);
    result.paymentReceiptFolder = {
      ok: true,
      name: folder.getName(),
      url: folder.getUrl()
    };
  } catch (err) {
    result.paymentReceiptFolder = {
      ok: false,
      error: String(err.message || err)
    };
    result.status = 'error';
  }

  try {
    const folder = DriveApp.getFolderById(COMMISSION_RECEIPT_ROOT_FOLDER_ID);
    result.commissionReceiptFolder = {
      ok: true,
      name: folder.getName(),
      url: folder.getUrl()
    };
  } catch (err) {
    result.commissionReceiptFolder = {
      ok: false,
      error: String(err.message || err)
    };
    result.status = 'error';
  }

  if (result.status === 'error') {
    result.nextStep = 'Share both folders to the web app owner account, run authorizeDriveAccess, then redeploy the web app as Execute as: Me.';
  }

  return result;
}
