const API_URL = "https://script.google.com/macros/s/AKfycbzVCOri-Rgw1NFxGYKsoDoCsDZEokXyvy1OgMn85CM3InntUF7PnEkZrQrlMHFplcgq/exec";
const CLOUDINARY_CLOUD_NAME = "dkd5mtpid";
const CLOUDINARY_UPLOAD_PRESET = "EC_APP";
const CLOUDINARY_FOLDER = "ec-app-receipts";
const CACHE_KEY = "booking_dashboard_cache_v2";
const AUTH_SESSION_KEY = "ec_app_auth_session_v1";
const AUTH_SESSION_USER_KEY = "ec_app_auth_session_user_v1";
const DEBUG = true;

const PAGE = document.body.dataset.page || "home";

const state = {
  bookings: [],
  filtered: [],
  isAuthenticated: false,
  authResolve: null,
  authUsername: "",
  authCurrentPassword: "",
  sendingRows: new Set(),
  activeReceiptRowIndex: null,
  activeCancelRowIndex: null,
  activeDownpaymentMinimum: 0,
  activeFullPaymentExact: 0,
  activeCommissionsSumAA: 0,
  activeCommissionsRows: [],
  commissionsSelectedDates: [],
  commissionsDateMenuOpen: false,
  commissionsExpandedDate: "",
  sendConfirmResolver: null,
  commissionSummaryGroups: [],
  selectedCommissionSummaryCoupon: ""
};

const elements = {
  searchInput: document.getElementById("searchInput"),
  tourFilter: document.getElementById("tourFilter"),
  travelDateFilter: document.getElementById("travelDateFilter"),
  refreshBtn: document.getElementById("refreshBtn"),
  bookingsBody: document.getElementById("bookingsBody"),
  tableHead: document.getElementById("tableHead"),
  pendingCount: document.getElementById("pendingCount"),
  loadingOverlay: document.getElementById("loadingOverlay"),
  toast: document.getElementById("toast"),
  sendConfirmModal: document.getElementById("sendConfirmModal"),
  sendConfirmMessage: document.getElementById("sendConfirmMessage"),
  sendConfirmOk: document.getElementById("sendConfirmOk"),
  sendConfirmCancel: document.getElementById("sendConfirmCancel"),
  receiptModal: document.getElementById("receiptModal"),
  receiptBookingInfo: document.getElementById("receiptBookingInfo"),
  receiptDownValue: document.getElementById("receiptDownValue"),
  receiptFullValue: document.getElementById("receiptFullValue"),
  receiptTypeSelect: document.getElementById("receiptTypeSelect"),
  downpaymentAmountGroup: document.getElementById("downpaymentAmountGroup"),
  downpaymentAmountInput: document.getElementById("downpaymentAmountInput"),
  downpaymentCurrentValue: document.getElementById("downpaymentCurrentValue"),
  downpaymentValidationNote: document.getElementById("downpaymentValidationNote"),
  fullPaymentAmountGroup: document.getElementById("fullPaymentAmountGroup"),
  fullPaymentAmountInput: document.getElementById("fullPaymentAmountInput"),
  fullPaymentCurrentValue: document.getElementById("fullPaymentCurrentValue"),
  fullPaymentValidationNote: document.getElementById("fullPaymentValidationNote"),
  receiptInvoiceInput: document.getElementById("receiptInvoiceInput"),
  receiptFileInput: document.getElementById("receiptFileInput"),
  receiptCancelBtn: document.getElementById("receiptCancelBtn"),
  receiptSaveBtn: document.getElementById("receiptSaveBtn"),
  receiptSavingIndicator: document.getElementById("receiptSavingIndicator"),
  receiptImageViewer: document.getElementById("receiptImageViewer"),
  receiptImageViewerImg: document.getElementById("receiptImageViewerImg"),
  receiptImageViewerClose: document.getElementById("receiptImageViewerClose"),
  cancelReasonModal: document.getElementById("cancelReasonModal"),
  cancelReasonBookingInfo: document.getElementById("cancelReasonBookingInfo"),
  cancelReasonSlot: document.getElementById("cancelReasonSlot"),
  cancelReasonInputGroup: document.getElementById("cancelReasonInputGroup"),
  cancelReasonInput: document.getElementById("cancelReasonInput"),
  cancelReasonCancelBtn: document.getElementById("cancelReasonCancelBtn"),
  cancelReasonSaveBtn: document.getElementById("cancelReasonSaveBtn"),
  commissionsBookingCount: document.getElementById("commissionsBookingCount"),
  commissionsTravelDateDropdown: document.getElementById("commissionsTravelDateDropdown"),
  commissionsTravelDateToggle: document.getElementById("commissionsTravelDateToggle"),
  commissionsTravelDateLabel: document.getElementById("commissionsTravelDateLabel"),
  commissionsTravelDateMenu: document.getElementById("commissionsTravelDateMenu"),
  commissionsTravelDateOptions: document.getElementById("commissionsTravelDateOptions"),
  commissionsCouponCode: document.getElementById("commissionsCouponCode"),
  commissionsSumAA: document.getElementById("commissionsSumAA"),
  commissionsReceiptAmount: document.getElementById("commissionsReceiptAmount"),
  commissionsAmountValidation: document.getElementById("commissionsAmountValidation"),
  commissionsInvoiceInput: document.getElementById("commissionsInvoiceInput"),
  commissionsReceiptFile: document.getElementById("commissionsReceiptFile"),
  commissionsSaveBtn: document.getElementById("commissionsSaveBtn"),
  commissionsSavingIndicator: document.getElementById("commissionsSavingIndicator"),
  commissionsUnpaidDatesCount: document.getElementById("commissionsUnpaidDatesCount"),
  commissionsUnpaidDatesBody: document.getElementById("commissionsUnpaidDatesBody"),
  commissionSummaryButtons: document.getElementById("commissionSummaryButtons"),
  commissionSummaryList: document.getElementById("commissionSummaryList"),
  commissionSummaryEmpty: document.getElementById("commissionSummaryEmpty"),
  travelDateSummaryList: document.getElementById("travelDateSummaryList"),
  travelDateSummaryEmpty: document.getElementById("travelDateSummaryEmpty"),
  homeNewPendingCount: document.getElementById("homeNewPendingCount"),
  homeConfirmedPendingCount: document.getElementById("homeConfirmedPendingCount"),
  homeDashboardPendingCount: document.getElementById("homeDashboardPendingCount"),
  homePaymentNoInvoiceCount: document.getElementById("homePaymentNoInvoiceCount"),
  homeCommissionsUnpaidDateCount: document.getElementById("homeCommissionsUnpaidDateCount")
};

const PAGE_CONFIG = {
  home: { needsData: false },
  "new-bookings": {
    needsData: true,
    countLabel: "Total Pending",
    columns: ["name", "date", "tour", "guests", "coupon", "discount", "paymentMethod", "totalAmount", "downpayment", "phone", "remarksNew", "actionNew"],
    filterFn: (row) => row.remarksNew.toLowerCase() !== "sent" && !isExcludedCouponRow(row)
  },
  "confirmed-whatsapp": {
    needsData: true,
    countLabel: "Total Pending",
    columns: ["name", "date", "tour", "totalBalance", "phone", "remarksConfirmed", "actionConfirmed"],
    filterFn: (row) => {
      const statusAD = String(row.adData || "").trim().toUpperCase();
      return row.remarksConfirmed.toLowerCase() !== "sent" && !isExcludedCouponRow(row) && statusAD !== "CANCELED";
    }
  },
  "bookings-dashboard": {
    needsData: true,
    countLabel: "Total Rows",
    columns: ["name", "date", "tour", "guests", "paymentMethod", "totalAmount", "downpayment", "totalBalance", "phone", "remarksNew", "remarksConfirmed", "adData"],
    filterFn: () => true
  },
  "payment-receipt-encoding": {
    needsData: true,
    countLabel: "Total Rows",
    columns: ["name", "date", "tour", "totalBalance", "paymentMethod", "receiptDown", "receiptFull"],
    filterFn: (row) => String(row.adData || "").trim().toUpperCase() !== "CANCELED"
  },
  "commissions-payment": {
    needsData: true,
    countLabel: "Matching Bookings",
    columns: [],
    filterFn: (row) => String(row.adData || "").trim().toUpperCase() !== "CANCELED"
  },
  "commissions-payment-summary": {
    needsData: false,
    countLabel: "",
    columns: [],
    filterFn: () => true
  }
};

const COLUMN_LABELS = {
  name: "Name",
  date: "Travel Date",
  tour: "Tour",
  guests: "# Guests",
  coupon: "Coupon",
  discount: "Discount",
  paymentMethod: "Payment Method",
  totalAmount: "Total Amount",
  downpayment: "Downpayment",
  totalBalance: "Total Balance",
  adData: "STATUS",
  phone: "Phone",
  remarksNew: "AG Remarks",
  remarksConfirmed: "AI Remarks",
  actionNew: "Action",
  actionConfirmed: "Action",
  receiptDown: "Y (Downpayment Invoice)",
  receiptFull: "AC (Full Invoice)",
  receiptAction: "Encode Receipt"
};

init();

async function init() {
  setupGlobalErrorHandlers();
  const ok = await ensureAuthenticated();
  if (!ok) return;

  if (PAGE === "home") {
    ensureLogoutControl_();
    fetchHomeCounts();
    return;
  }

  if (PAGE === "commissions-payment-summary") {
    bindEvents();
    fetchCommissionPaymentSummaries();
    return;
  }

  bindEvents();
  hydrateFromCache();
  fetchBookings();
}

async function ensureAuthenticated() {
  if (state.isAuthenticated) return true;

  if (hasSessionAuthentication_()) {
    state.isAuthenticated = true;
    state.authUsername = getSessionAuthUsername_() || "UNKNOWN";
    return true;
  }

  ensureAuthModal_();
  showAuthStep_("login");

  return await new Promise((resolve) => {
    state.authResolve = resolve;
  });
}

function ensureAuthModal_() {
  if (document.getElementById("authModal")) return;

  const wrapper = document.createElement("div");
  wrapper.id = "authModal";
  wrapper.className = "auth-modal";
  wrapper.innerHTML = `
    <div class="auth-modal-card" role="dialog" aria-modal="true" aria-labelledby="authTitle">
      <h3 id="authTitle">Login</h3>
      <p id="authSubtitle" class="modal-subtext">Enter your username and password.</p>

      <div id="authLoginFields">
        <label>
          <span class="modal-label">Username</span>
          <input id="authUsername" type="text" autocomplete="username" />
        </label>
        <label>
          <span class="modal-label">Password</span>
          <div class="auth-password-wrap">
            <input id="authPassword" type="password" autocomplete="current-password" />
            <button id="authPasswordToggle" class="auth-password-toggle" type="button" aria-label="Show password" aria-pressed="false">Show</button>
          </div>
          <p id="authLoginError" class="auth-login-error" hidden></p>
        </label>
      </div>

      <div id="authNewPasswordFields" hidden>
        <label>
          <span class="modal-label">New Password</span>
          <div class="auth-password-wrap">
            <input id="authNewPassword" type="password" autocomplete="new-password" />
            <button id="authNewPasswordToggle" class="auth-password-toggle" type="button" aria-label="Show password" aria-pressed="false">Show</button>
          </div>
        </label>
        <label>
          <span class="modal-label">Confirm New Password</span>
          <div class="auth-password-wrap">
            <input id="authConfirmPassword" type="password" autocomplete="new-password" />
            <button id="authConfirmPasswordToggle" class="auth-password-toggle" type="button" aria-label="Show password" aria-pressed="false">Show</button>
          </div>
        </label>
      </div>

      <div id="authResetFields" hidden>
        <label>
          <span class="modal-label">Username</span>
          <input id="authResetUsername" type="text" autocomplete="username" />
        </label>
      </div>

      <div class="auth-modal-actions">
        <button id="authBackBtn" class="btn btn-secondary" type="button" hidden>Back</button>
        <button id="authForgotBtn" class="btn btn-secondary" type="button">Forgot Password</button>
        <button id="authSubmitBtn" class="btn btn-primary" type="button">Login</button>
      </div>
    </div>
  `;

  document.body.appendChild(wrapper);

  wrapper.querySelector("#authSubmitBtn")?.addEventListener("click", handleAuthSubmit_);
  wrapper.querySelector("#authForgotBtn")?.addEventListener("click", () => showAuthStep_("reset"));
  wrapper.querySelector("#authBackBtn")?.addEventListener("click", () => showAuthStep_("login"));
  wrapper.querySelector("#authPasswordToggle")?.addEventListener("click", () => toggleAuthPasswordVisibility_("authPassword", "authPasswordToggle"));
  wrapper.querySelector("#authNewPasswordToggle")?.addEventListener("click", () => toggleAuthPasswordVisibility_("authNewPassword", "authNewPasswordToggle"));
  wrapper.querySelector("#authConfirmPasswordToggle")?.addEventListener("click", () => toggleAuthPasswordVisibility_("authConfirmPassword", "authConfirmPasswordToggle"));
  wrapper.querySelector("#authUsername")?.addEventListener("input", () => setAuthLoginError_(""));
  wrapper.querySelector("#authPassword")?.addEventListener("input", () => setAuthLoginError_(""));
}

function setAuthLoginError_(message) {
  const modal = document.getElementById("authModal");
  if (!modal) return;
  const errorEl = modal.querySelector("#authLoginError");
  if (!errorEl) return;

  const text = String(message || "").trim();
  if (!text) {
    errorEl.hidden = true;
    errorEl.textContent = "";
    return;
  }

  errorEl.textContent = text;
  errorEl.hidden = false;
}

function toggleAuthPasswordVisibility_(inputId, toggleId) {
  const modal = document.getElementById("authModal");
  if (!modal) return;

  const passwordInput = modal.querySelector(`#${inputId}`);
  const toggleBtn = modal.querySelector(`#${toggleId}`);
  if (!passwordInput || !toggleBtn) return;

  const shouldShow = passwordInput.type === "password";
  passwordInput.type = shouldShow ? "text" : "password";
  toggleBtn.textContent = shouldShow ? "Hide" : "Show";
  toggleBtn.setAttribute("aria-label", shouldShow ? "Hide password" : "Show password");
  toggleBtn.setAttribute("aria-pressed", shouldShow ? "true" : "false");
}

function showAuthStep_(step) {
  const modal = document.getElementById("authModal");
  if (!modal) return;

  modal.hidden = false;
  modal.dataset.step = step;

  const title = modal.querySelector("#authTitle");
  const subtitle = modal.querySelector("#authSubtitle");
  const loginFields = modal.querySelector("#authLoginFields");
  const newPasswordFields = modal.querySelector("#authNewPasswordFields");
  const resetFields = modal.querySelector("#authResetFields");
  const submitBtn = modal.querySelector("#authSubmitBtn");
  const forgotBtn = modal.querySelector("#authForgotBtn");
  const backBtn = modal.querySelector("#authBackBtn");
  const authPasswordInput = modal.querySelector("#authPassword");
  const authPasswordToggle = modal.querySelector("#authPasswordToggle");
  const newPasswordInput = modal.querySelector("#authNewPassword");
  const newPasswordToggle = modal.querySelector("#authNewPasswordToggle");
  const confirmPasswordInput = modal.querySelector("#authConfirmPassword");
  const confirmPasswordToggle = modal.querySelector("#authConfirmPasswordToggle");

  setAuthLoginError_("");

  if (authPasswordInput) {
    authPasswordInput.type = "password";
  }
  if (authPasswordToggle) {
    authPasswordToggle.textContent = "Show";
    authPasswordToggle.setAttribute("aria-label", "Show password");
    authPasswordToggle.setAttribute("aria-pressed", "false");
  }
  if (newPasswordInput) {
    newPasswordInput.type = "password";
  }
  if (newPasswordToggle) {
    newPasswordToggle.textContent = "Show";
    newPasswordToggle.setAttribute("aria-label", "Show password");
    newPasswordToggle.setAttribute("aria-pressed", "false");
  }
  if (confirmPasswordInput) {
    confirmPasswordInput.type = "password";
  }
  if (confirmPasswordToggle) {
    confirmPasswordToggle.textContent = "Show";
    confirmPasswordToggle.setAttribute("aria-label", "Show password");
    confirmPasswordToggle.setAttribute("aria-pressed", "false");
  }

  loginFields.hidden = step !== "login";
  newPasswordFields.hidden = step !== "new-password";
  resetFields.hidden = step !== "reset";

  if (step === "login") {
    title.textContent = "Login";
    subtitle.textContent = "Enter your username and password.";
    submitBtn.textContent = "Login";
    forgotBtn.hidden = false;
    backBtn.hidden = true;
  } else if (step === "new-password") {
    title.textContent = "Create New Password";
    subtitle.textContent = "First login detected. You must set a new password.";
    submitBtn.textContent = "Save Password";
    forgotBtn.hidden = true;
    backBtn.hidden = true;
  } else {
    title.textContent = "Reset Password";
    subtitle.textContent = "Send reset request email to owner for confirmation.";
    submitBtn.textContent = "Send Reset";
    forgotBtn.hidden = true;
    backBtn.hidden = false;
  }
}

function completeAuthentication_() {
  state.isAuthenticated = true;
  setSessionAuthentication_(state.authUsername);
  const modal = document.getElementById("authModal");
  if (modal) {
    modal.hidden = true;
  }

  if (typeof state.authResolve === "function") {
    state.authResolve(true);
    state.authResolve = null;
  }
}

function forceRelogin_(message) {
  state.isAuthenticated = false;
  state.authUsername = "";
  state.authCurrentPassword = "";
  clearSessionAuthentication_();

  ensureAuthModal_();
  showAuthStep_("login");

  const modal = document.getElementById("authModal");
  if (modal) {
    const passwordInput = modal.querySelector("#authPassword");
    if (passwordInput) {
      passwordInput.value = "";
    }
  }

  if (message) {
    showToast(message);
  }
}

function ensureLogoutControl_() {
  if (PAGE !== "home") return;

  const topbar = document.querySelector(".topbar");
  if (!topbar) return;

  if (!topbar.style.position) {
    topbar.style.position = "relative";
  }

  let logoutBtn = document.getElementById("logoutBtn");
  if (!logoutBtn) {
    logoutBtn = document.createElement("button");
    logoutBtn.id = "logoutBtn";
    logoutBtn.type = "button";
    logoutBtn.className = "btn btn-secondary";
    logoutBtn.textContent = "Logout";
    logoutBtn.style.position = "absolute";
    logoutBtn.style.top = "0";
    logoutBtn.style.right = "0";
    topbar.appendChild(logoutBtn);
  }

  if (logoutBtn.dataset.bound === "1") return;
  logoutBtn.dataset.bound = "1";
  logoutBtn.addEventListener("click", () => {
    state.isAuthenticated = false;
    state.authUsername = "";
    state.authCurrentPassword = "";
    clearSessionAuthentication_();
    window.location.reload();
  });
}

function hasSessionAuthentication_() {
  try {
    return sessionStorage.getItem(AUTH_SESSION_KEY) === "1";
  } catch {
    return false;
  }
}

function getSessionAuthUsername_() {
  try {
    return String(sessionStorage.getItem(AUTH_SESSION_USER_KEY) || "").trim().toUpperCase();
  } catch {
    return "";
  }
}

function setSessionAuthentication_(username) {
  try {
    sessionStorage.setItem(AUTH_SESSION_KEY, "1");
    const value = String(username || "").trim().toUpperCase();
    if (value) {
      sessionStorage.setItem(AUTH_SESSION_USER_KEY, value);
    }
  } catch {
    // Ignore storage errors and keep in-memory auth only.
  }
}

function clearSessionAuthentication_() {
  try {
    sessionStorage.removeItem(AUTH_SESSION_KEY);
    sessionStorage.removeItem(AUTH_SESSION_USER_KEY);
  } catch {
    // Ignore storage errors.
  }
}

function getActorUsername_() {
  return String(state.authUsername || "UNKNOWN").trim().toUpperCase();
}

async function handleAuthSubmit_() {
  const modal = document.getElementById("authModal");
  if (!modal) return;

  const step = String(modal.dataset.step || "login");
  const submitBtn = modal.querySelector("#authSubmitBtn");
  const previousText = submitBtn?.textContent || "";
  if (submitBtn) {
    submitBtn.disabled = true;
    submitBtn.textContent = "Please wait...";
  }

  try {
    if (step === "login") {
      setAuthLoginError_("");
      const username = String(modal.querySelector("#authUsername")?.value || "").trim().toUpperCase();
      const password = String(modal.querySelector("#authPassword")?.value || "").trim();

      if (!username || !password) {
        setAuthLoginError_("Username and password are required.");
        return;
      }

      const response = await fetchWithTimeout(API_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({ action: "login_user", username, password })
      }, 15000);

      const result = await parseJsonSafe(response);
      if (!response.ok || String(result?.status || "").toLowerCase() === "error") {
        throw new Error(result?.message || `Login failed: ${response.status}`);
      }

      state.authUsername = username;
      state.authCurrentPassword = password;

      if (result.mustChangePassword) {
        showAuthStep_("new-password");
        return;
      }

      completeAuthentication_();
      return;
    }

    if (step === "new-password") {
      const newPassword = String(modal.querySelector("#authNewPassword")?.value || "").trim();
      const confirmPassword = String(modal.querySelector("#authConfirmPassword")?.value || "").trim();

      if (!newPassword || !confirmPassword) {
        showToast("New password and confirmation are required.", true);
        return;
      }
      if (newPassword !== confirmPassword) {
        showToast("New password and confirmation do not match.", true);
        return;
      }

      const response = await fetchWithTimeout(API_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({
          action: "set_new_password",
          username: state.authUsername,
          currentPassword: state.authCurrentPassword,
          newPassword
        })
      }, 15000);

      const result = await parseJsonSafe(response);
      if (!response.ok || String(result?.status || "").toLowerCase() === "error") {
        throw new Error(result?.message || `Failed to set new password: ${response.status}`);
      }

      forceRelogin_("Password changed. Please login again.");
      return;
    }

    const resetUsername = String(modal.querySelector("#authResetUsername")?.value || "").trim().toUpperCase();
    if (!resetUsername) {
      showToast("Username is required.", true);
      return;
    }

    const response = await fetchWithTimeout(API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ action: "request_password_reset", username: resetUsername })
    }, 15000);

    const result = await parseJsonSafe(response);
    if (!response.ok || String(result?.status || "").toLowerCase() === "error") {
      throw new Error(result?.message || `Failed to request reset: ${response.status}`);
    }

    showToast(result?.message || "Reset request sent.");
    showAuthStep_("login");
  } catch (error) {
    debugError("handleAuthSubmit", error, {});
    if (step === "login") {
      setAuthLoginError_(error.message || "Authentication failed.");
    }
    showToast(error.message || "Authentication failed.", true);
  } finally {
    if (submitBtn) {
      submitBtn.disabled = false;
      submitBtn.textContent = previousText || "Submit";
    }
  }
}

function bindEvents() {
  elements.searchInput?.addEventListener("input", applyFilters);
  elements.tourFilter?.addEventListener("change", applyFilters);
  elements.travelDateFilter?.addEventListener("change", applyFilters);
  if (PAGE === "commissions-payment-summary") {
    elements.refreshBtn?.addEventListener("click", fetchCommissionPaymentSummaries);
  } else {
    elements.refreshBtn?.addEventListener("click", fetchBookings);
  }
  elements.sendConfirmOk?.addEventListener("click", () => resolveSendConfirm(true));
  elements.sendConfirmCancel?.addEventListener("click", () => resolveSendConfirm(false));
  elements.sendConfirmModal?.addEventListener("click", (event) => {
    if (event.target === elements.sendConfirmModal) {
      resolveSendConfirm(false);
    }
  });
  elements.receiptCancelBtn?.addEventListener("click", closeReceiptModal);
  elements.receiptSaveBtn?.addEventListener("click", handleSaveReceipt);
  elements.receiptDownValue?.addEventListener("click", () => handleModalInvoiceClick("downpayment"));
  elements.receiptFullValue?.addEventListener("click", () => handleModalInvoiceClick("full"));
  elements.receiptImageViewerClose?.addEventListener("click", closeReceiptImageViewer);
  elements.cancelReasonCancelBtn?.addEventListener("click", closeCancelReasonModal);
  elements.cancelReasonSaveBtn?.addEventListener("click", handleSaveCancelReason);
  elements.cancelReasonSlot?.addEventListener("change", updateCancelReasonInputVisibility);
  elements.receiptTypeSelect?.addEventListener("change", updateDownpaymentAmountVisibility);
  elements.downpaymentAmountInput?.addEventListener("input", validateDownpaymentAmount);
  elements.fullPaymentAmountInput?.addEventListener("input", validateFullPaymentAmount);
  elements.commissionsTravelDateToggle?.addEventListener("click", toggleCommissionsDateMenu);
  elements.commissionsCouponCode?.addEventListener("change", handleCommissionsCouponChange);
  elements.commissionsReceiptAmount?.addEventListener("input", validateCommissionsReceiptAmount);
  elements.commissionsSaveBtn?.addEventListener("click", handleSaveCommissionsReceipt);
  document.addEventListener("click", handleCommissionsDocumentClick);
  elements.receiptModal?.addEventListener("click", (event) => {
    if (event.target === elements.receiptModal) {
      closeReceiptModal();
    }
  });
  elements.receiptImageViewer?.addEventListener("click", (event) => {
    if (event.target === elements.receiptImageViewer) {
      closeReceiptImageViewer();
    }
  });
  elements.cancelReasonModal?.addEventListener("click", (event) => {
    if (event.target === elements.cancelReasonModal) {
      closeCancelReasonModal();
    }
  });
}

async function refreshCurrentPageData() {
  if (PAGE === "home") {
    await fetchHomeCounts();
    return;
  }

  if (PAGE === "commissions-payment-summary") {
    await fetchCommissionPaymentSummaries();
    return;
  }

  if (PAGE_CONFIG[PAGE]?.needsData) {
    await fetchBookings();
  }
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, Number(ms) || 0));
}

async function refreshCurrentPageDataWithDelay(ms = 300) {
  await delay(ms);
  await refreshCurrentPageData();
}

async function fetchBookings() {
  if (!isApiConfigured()) {
    showToast("Set your API URL in script.js", true);
    return;
  }

  setLoading(true);
  const startedAt = performance.now();

  try {
    const response = await fetchWithTimeout(API_URL, { method: "GET" }, 20000);
    if (!response.ok) {
      throw new Error(`Fetch failed: ${response.status}`);
    }

    const data = await parseJsonSafe(response);
    const rows = Array.isArray(data) ? data : data.data;

    if (!Array.isArray(rows)) {
      throw new Error("Invalid API response format");
    }

    state.bookings = rows
      .filter((item) => item && item.rowIndex)
      .map(normalizeRow);

    saveCache(state.bookings);
    populateTourFilter(state.bookings);
    populateTravelDateFilter(state.bookings);
    applyFilters();

    debugInfo("fetchBookings:success", {
      page: PAGE,
      count: state.bookings.length,
      durationMs: Math.round(performance.now() - startedAt)
    });
  } catch (error) {
    debugError("fetchBookings", error, { page: PAGE });
    showToast(error.message || "Failed to load bookings.", true);
  } finally {
    setLoading(false);
  }
}

function normalizeRow(row) {
  return {
    rowIndex: Number(row.rowIndex),
    bookingDate: String(row.bookingDate || "").trim(),
    name: String(row.name || "").trim(),
    date: String(row.date || "").trim(),
    tour: String(row.tour || "").trim(),
    phone: String(row.phone || "").trim(),
    numberOfGuest: String(row.numberOfGuest || "").trim(),
    couponCode: String(row.couponCode || "").trim(),
    excludedCouponCode: String(row.excludedCouponCode || "").trim(),
    discountAmount: String(row.discountAmount || "").trim(),
    paymentMethod: String(row.paymentMethod || "").trim(),
    totalAmount: String(row.totalAmount || "").trim(),
    downpaymentAmount: String(row.downpaymentAmount || "").trim(),
    commissionAmountAA: String(row.commissionAmountAA || "").trim(),
    influencerInvoiceAK: String(row.influencerInvoiceAK || "").trim(),
    influencerStatusAL: String(row.influencerStatusAL || "").trim(),
    totalBalance: String(row.totalBalance || "").trim(),
    adData: String(row.adData || "").trim(),
    whatsappNew: String(row.whatsappNew || row.whatsapp || "").trim(),
    remarksNew: String(row.remarksNew || row.remarks || "").trim(),
    whatsappConfirmed: String(row.whatsappConfirmed || "").trim(),
    remarksConfirmed: String(row.remarksConfirmed || "").trim(),
    receiptDownpayment: String(row.receiptDownpayment || "").trim(),
    receiptFullPayment: String(row.receiptFullPayment || "").trim(),
    receiptUrlAM: String(row.receiptUrlAM || "").trim(),
    receiptUrlAN: String(row.receiptUrlAN || "").trim(),
    cancelReasonAO: String(row.cancelReasonAO || "").trim()
  };
}

function isExcludedCouponRow(row) {
  const rowCoupon = String(row?.couponCode || "").trim().toUpperCase();
  const excludedCoupon = String(row?.excludedCouponCode || "").trim().toUpperCase();
  if (!rowCoupon || !excludedCoupon) return false;
  return rowCoupon === excludedCoupon;
}

function isCommissionRowEligible(row) {
  const coupon = String(row?.couponCode || "").trim();
  const status = String(row?.influencerStatusAL || "").trim().toUpperCase();
  const bookingStatus = String(row?.adData || "").trim().toUpperCase();
  const isBookingCompleted = bookingStatus === "BOOKING COMPLETED" || bookingStatus === "COMPLETED";
  const isBlockedStatus =
    bookingStatus === "CANCEL" ||
    bookingStatus === "CANCELED" ||
    bookingStatus === "CANCELLED" ||
    bookingStatus === "PENDING PAYMENT";
  return Boolean(coupon) && status !== "PAID" && isBookingCompleted && !isBlockedStatus;
}

function applyFilters() {
  const cfg = PAGE_CONFIG[PAGE];
  if (!cfg || !cfg.needsData) return;

  if (PAGE === "commissions-payment") {
    state.filtered = state.bookings.slice();
    populateCommissionsCouponFilter(state.bookings);
    renderCommissionsUnpaidDatesDashboard(state.bookings);
    handleCommissionsCouponChange();
    return;
  }

  const query = (elements.searchInput?.value || "").trim().toLowerCase();
  const selectedTour = elements.tourFilter?.value || "";
  const selectedTravelDate = elements.travelDateFilter?.value || "";

  const scoped = state.bookings.filter(cfg.filterFn);

  state.filtered = scoped.filter((booking) => {
    const matchesName = !query || booking.name.toLowerCase().includes(query);
    const matchesTour = !selectedTour || booking.tour === selectedTour;
    const matchesDate = !selectedTravelDate || normalizeDateKey(booking.date) === selectedTravelDate;
    return matchesName && matchesTour && matchesDate;
  });

  if (PAGE === "new-bookings" || PAGE === "confirmed-whatsapp") {
    state.filtered.sort(compareByLatestBookingDate);
  }

  if (PAGE === "bookings-dashboard" || PAGE === "payment-receipt-encoding") {
    state.filtered.sort(compareByNearestUpcomingTravelDate);
  }

  renderTable(state.filtered, cfg.columns);
  if (PAGE === "bookings-dashboard") {
    const pendingPaymentCount = scoped.filter((row) => String(row.adData || "").trim().toUpperCase() === "PENDING PAYMENT").length;
    updatePendingCount(pendingPaymentCount);
    renderUpcomingTravelDateSummary(state.bookings);
    return;
  }

  if (PAGE === "payment-receipt-encoding") {
    const noInvoiceCount = scoped.filter((row) => !String(row.receiptFullPayment || "").trim()).length;
    updatePendingCount(noInvoiceCount);
    return;
  }

  updatePendingCount(scoped.length);
}

function renderUpcomingTravelDateSummary(rows) {
  if (!elements.travelDateSummaryList || !elements.travelDateSummaryEmpty) return;

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const grouped = new Map();

  rows.forEach((row) => {
    const dateKey = normalizeDateKey(row?.date);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) return;

    const travelDate = toDateOrNull(dateKey);
    if (!travelDate || travelDate.getTime() < today.getTime()) return;

    if (!grouped.has(dateKey)) {
      grouped.set(dateKey, {
        dateKey,
        total: 0,
        completed: 0,
        pending: 0
      });
    }

    const item = grouped.get(dateKey);
    item.total += 1;

    const statusText = String(row?.adData || "").trim().toUpperCase();
    if (statusText === "PENDING PAYMENT") {
      item.pending += 1;
    } else {
      item.completed += 1;
    }
  });

  const upcomingDates = [...grouped.values()]
    .sort((a, b) => a.dateKey.localeCompare(b.dateKey))
    .slice(0, 4);

  if (!upcomingDates.length) {
    elements.travelDateSummaryList.innerHTML = "";
    elements.travelDateSummaryEmpty.hidden = false;
    return;
  }

  elements.travelDateSummaryEmpty.hidden = true;

  elements.travelDateSummaryList.innerHTML = upcomingDates
    .map((item) => {
      const total = Number(item.total || 0);
      const completed = Number(item.completed || 0);
      const pending = Number(item.pending || 0);
      const completedAngle = total > 0 ? Math.round((completed / total) * 360) : 0;

      return `
        <article class="travel-date-donut-item" aria-label="${escapeHtml(formatDate(item.dateKey))}">
          <div class="travel-date-donut" style="--completed-angle:${completedAngle}deg;">
            <div class="travel-date-donut-inner">${escapeHtml(String(total))}</div>
          </div>
          <p class="travel-date-donut-date">${escapeHtml(formatDate(item.dateKey))}</p>
          <p class="travel-date-donut-meta">Completed ${escapeHtml(String(completed))} • Pending ${escapeHtml(String(pending))}</p>
        </article>
      `;
    })
    .join("");
}

function compareByLatestBookingDate(a, b) {
  const aDate = toDateOrNull(a.bookingDate);
  const bDate = toDateOrNull(b.bookingDate);

  if (aDate && bDate) {
    return bDate.getTime() - aDate.getTime();
  }
  if (aDate) return -1;
  if (bDate) return 1;

  return Number(b.rowIndex || 0) - Number(a.rowIndex || 0);
}

function compareByNearestUpcomingTravelDate(a, b) {
  const aDate = toDateOrNull(a.date);
  const bDate = toDateOrNull(b.date);

  if (aDate && bDate) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const aUpcoming = aDate.getTime() >= today.getTime();
    const bUpcoming = bDate.getTime() >= today.getTime();

    if (aUpcoming && bUpcoming) {
      return aDate.getTime() - bDate.getTime();
    }
    if (aUpcoming && !bUpcoming) return -1;
    if (!aUpcoming && bUpcoming) return 1;

    return bDate.getTime() - aDate.getTime();
  }

  if (aDate) return -1;
  if (bDate) return 1;
  return Number(a.rowIndex || 0) - Number(b.rowIndex || 0);
}

function toDateOrNull(value) {
  const parsed = new Date(String(value || "").trim());
  if (Number.isNaN(parsed.getTime())) return null;
  parsed.setHours(0, 0, 0, 0);
  return parsed;
}

async function fetchHomeCounts() {
  if (!isApiConfigured()) return;

  try {
    const response = await fetchWithTimeout(API_URL, { method: "GET" }, 20000);
    if (!response.ok) {
      throw new Error(`Fetch failed: ${response.status}`);
    }

    const data = await parseJsonSafe(response);
    const rows = Array.isArray(data) ? data : data.data;
    if (!Array.isArray(rows)) {
      throw new Error("Invalid API response format");
    }

    const bookings = rows
      .filter((item) => item && item.rowIndex)
      .map(normalizeRow);

    const newPending = bookings.filter((row) => row.remarksNew.toLowerCase() !== "sent" && !isExcludedCouponRow(row)).length;
    const confirmedPending = bookings.filter((row) => row.remarksConfirmed.toLowerCase() !== "sent" && !isExcludedCouponRow(row)).length;
    const dashboardPending = bookings.filter((row) => String(row.adData || "").trim().toUpperCase() === "PENDING PAYMENT").length;
    const paymentNoInvoice = bookings.filter((row) => !String(row.receiptFullPayment || "").trim()).length;
    const commissionsUnpaidDateCount = new Set(
      bookings
        .filter(isCommissionRowEligible)
        .map((row) => normalizeDateKey(row.date))
        .filter(Boolean)
    ).size;

    if (elements.homeNewPendingCount) {
      elements.homeNewPendingCount.textContent = String(newPending);
      elements.homeNewPendingCount.hidden = false;
    }
    if (elements.homeConfirmedPendingCount) {
      elements.homeConfirmedPendingCount.textContent = String(confirmedPending);
      elements.homeConfirmedPendingCount.hidden = false;
    }
    if (elements.homeDashboardPendingCount) {
      elements.homeDashboardPendingCount.textContent = String(dashboardPending);
      elements.homeDashboardPendingCount.hidden = false;
    }
    if (elements.homePaymentNoInvoiceCount) {
      elements.homePaymentNoInvoiceCount.textContent = String(paymentNoInvoice);
      elements.homePaymentNoInvoiceCount.hidden = false;
    }
    if (elements.homeCommissionsUnpaidDateCount) {
      elements.homeCommissionsUnpaidDateCount.textContent = String(commissionsUnpaidDateCount);
      elements.homeCommissionsUnpaidDateCount.hidden = false;
    }
  } catch (error) {
    debugError("fetchHomeCounts", error, { page: PAGE });
  }
}

async function fetchCommissionPaymentSummaries() {
  if (!isApiConfigured()) {
    showToast("Set your API URL in script.js", true);
    return;
  }

  setLoading(true);

  try {
    const response = await fetchWithTimeout(API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ action: "get_commission_payment_summaries" })
    }, 20000);

    if (!response.ok) {
      throw new Error(`Fetch failed: ${response.status}`);
    }

    const result = await parseJsonSafe(response);
    if (result && typeof result === "object" && String(result.status || "").toLowerCase() === "error") {
      throw new Error(result.message || "Failed to load payment summaries.");
    }

    state.commissionSummaryGroups = Array.isArray(result?.groups) ? result.groups : [];

    const firstCoupon = state.commissionSummaryGroups[0]?.couponCode || "";
    if (!state.selectedCommissionSummaryCoupon || !state.commissionSummaryGroups.some((g) => g.couponCode === state.selectedCommissionSummaryCoupon)) {
      state.selectedCommissionSummaryCoupon = firstCoupon;
    }

    renderCommissionSummaryButtons();
    renderCommissionSummaryList();
  } catch (error) {
    debugError("fetchCommissionPaymentSummaries", error);
    showToast(error.message || "Failed to load payment summaries.", true);
  } finally {
    setLoading(false);
  }
}

function renderCommissionSummaryButtons() {
  if (!elements.commissionSummaryButtons) return;

  const groups = state.commissionSummaryGroups;
  if (!groups.length) {
    elements.commissionSummaryButtons.innerHTML = "";
    return;
  }

  elements.commissionSummaryButtons.innerHTML = groups
    .map((group) => {
      const active = group.couponCode === state.selectedCommissionSummaryCoupon;
      return `<button class="btn ${active ? "btn-primary" : "btn-secondary"} summary-coupon-btn" data-action="select-summary-coupon" data-coupon="${escapeHtmlAttr(group.couponCode)}">${escapeHtml(group.couponCode)}</button>`;
    })
    .join("");

  elements.commissionSummaryButtons.querySelectorAll("button[data-action='select-summary-coupon']").forEach((btn) => {
    btn.addEventListener("click", () => {
      state.selectedCommissionSummaryCoupon = String(btn.dataset.coupon || "").trim();
      renderCommissionSummaryButtons();
      renderCommissionSummaryList();
    });
  });
}

function renderCommissionSummaryList() {
  if (!elements.commissionSummaryList || !elements.commissionSummaryEmpty) return;

  const selectedCoupon = state.selectedCommissionSummaryCoupon;
  const group = state.commissionSummaryGroups.find((item) => item.couponCode === selectedCoupon);

  if (!group || !Array.isArray(group.payments) || !group.payments.length) {
    elements.commissionSummaryList.innerHTML = "";
    elements.commissionSummaryEmpty.hidden = false;
    elements.commissionSummaryEmpty.textContent = "No payment summaries found for this coupon code.";
    return;
  }

  elements.commissionSummaryEmpty.hidden = true;

  elements.commissionSummaryList.innerHTML = group.payments
    .map((payment) => {
      const lineRows = Array.isArray(payment.rows) ? payment.rows : [];
      const lineRowsHtml = lineRows.length
        ? lineRows
            .map((line) => `
              <tr>
                <td>${escapeHtml(formatDate(line.date || ""))}</td>
                <td>${escapeHtml(String(line.bookings || "-"))}</td>
                <td>${escapeHtml(String(line.amount || "-"))}</td>
                <td>${escapeHtml(String(line.total || "-"))}</td>
              </tr>
            `)
            .join("")
        : '<tr><td colspan="4" class="empty-state">No line items.</td></tr>';

      const receiptUrl = getValidHttpUrl(payment.receiptUrl);
      const receiptBtn = receiptUrl
        ? `<button class="btn btn-secondary" data-action="open-summary-receipt" data-url="${escapeHtmlAttr(receiptUrl)}" data-label="Commission Receipt">See Receipt Image</button>`
        : `<button class="btn btn-secondary" type="button" disabled>No Receipt Image</button>`;

      return `
        <section class="summary-card">
          <div class="summary-card-header">
            <div>
              <h3>${escapeHtml(payment.title || `Invoice # ${payment.invoiceNumber || "-"}`)}</h3>
              <p class="subtitle">Coupon Code: ${escapeHtml(selectedCoupon)}</p>
            </div>
            <div>${receiptBtn}</div>
          </div>
          <div class="table-scroll">
            <table>
              <thead>
                <tr>
                  <th>Travel Date</th>
                  <th>No. of Bookings</th>
                  <th>Amount</th>
                  <th>Total</th>
                </tr>
              </thead>
              <tbody>
                ${lineRowsHtml}
                <tr class="summary-grand-total-row">
                  <td colspan="3"><strong>GRAND TOTAL</strong></td>
                  <td><strong>${escapeHtml(String(payment.grandTotal || "-"))}</strong></td>
                </tr>
              </tbody>
            </table>
          </div>
        </section>
      `;
    })
    .join("");

  elements.commissionSummaryList.querySelectorAll("button[data-action='open-summary-receipt']").forEach((btn) => {
    btn.addEventListener("click", () => {
      const url = getValidHttpUrl(btn.dataset.url || "");
      if (!url) {
        showToast("Receipt image URL is missing or invalid.", true);
        return;
      }
      openReceiptImageViewer(url, btn.dataset.label || "Receipt Image");
    });
  });
}

function renderCommissionsUnpaidDatesDashboard(rows) {
  if (!elements.commissionsUnpaidDatesBody) return;

  const summary = new Map();

  rows
    .filter(isCommissionRowEligible)
    .forEach((row) => {
      const dateKey = normalizeDateKey(row.date);
      if (!dateKey) return;

      if (!summary.has(dateKey)) {
        summary.set(dateKey, {
          count: 0,
          totalAA: 0,
          coupons: new Set()
        });
      }

      const item = summary.get(dateKey);
      item.count += 1;
      const amount = normalizeAmountInput(row.commissionAmountAA);
      const guestCount = normalizeGuestCountInput(row.numberOfGuest);
      item.totalAA += amount === null ? 0 : amount * guestCount;

      const couponCode = String(row.couponCode || "").trim();
      if (couponCode) {
        item.coupons.add(couponCode);
      }
    });

  const travelDates = [...summary.keys()].sort((a, b) => a.localeCompare(b));

  if (state.commissionsExpandedDate && !travelDates.includes(state.commissionsExpandedDate)) {
    state.commissionsExpandedDate = "";
  }

  if (elements.commissionsUnpaidDatesCount) {
    elements.commissionsUnpaidDatesCount.textContent = String(travelDates.length);
  }

  if (!travelDates.length) {
    elements.commissionsUnpaidDatesBody.innerHTML = '<tr><td colspan="4" class="empty-state">No unpaid travel dates found.</td></tr>';
    return;
  }

  elements.commissionsUnpaidDatesBody.innerHTML = travelDates.map((dateKey) => {
    const item = summary.get(dateKey);
    const couponText = [...item.coupons].sort((a, b) => a.localeCompare(b)).join(", ") || "-";
    const isExpanded = state.commissionsExpandedDate === dateKey;

    const detailsRows = rows
      .filter(isCommissionRowEligible)
      .filter((row) => normalizeDateKey(row.date) === dateKey)
      .map((row) => {
        const guestCount = normalizeGuestCountInput(row.numberOfGuest);
        const amountAA = normalizeAmountInput(row.commissionAmountAA) || 0;
        const totalAmount = Number((amountAA * guestCount).toFixed(2));
        const packageText = String(row.tour || "-").trim() || "-";
        const couponText = String(row.couponCode || "-").trim() || "-";
        const packageWithCoupon = `${packageText} (${couponText})`;

        return `
          <tr>
            <td>${escapeHtml(formatDate(row.date))}</td>
            <td>${escapeHtml(row.name || "-")}</td>
            <td>${escapeHtml(packageWithCoupon)}</td>
            <td>${guestCount}</td>
            <td>${formatMoney(amountAA)}</td>
            <td>${formatMoney(totalAmount)}</td>
          </tr>
        `;
      })
      .join("");

    const detailsHtml = isExpanded
      ? `
        <tr class="commissions-unpaid-detail-row">
          <td colspan="4">
            <div class="commissions-unpaid-detail-wrap">
              <table class="commissions-unpaid-detail-table">
                <thead>
                  <tr>
                    <th>Travel Date</th>
                    <th>Guest Name</th>
                    <th>Package</th>
                    <th>No. of Guests</th>
                    <th>AA</th>
                    <th>Total (AA × Guests)</th>
                  </tr>
                </thead>
                <tbody>
                  ${detailsRows || '<tr><td colspan="6" class="empty-state">No applicable rows.</td></tr>'}
                </tbody>
              </table>
            </div>
          </td>
        </tr>
      `
      : "";

    return `
      <tr class="receipt-click-row commissions-unpaid-click-row" data-action="toggle-commission-unpaid-date" data-date-key="${escapeHtmlAttr(dateKey)}">
        <td>${escapeHtml(formatDate(dateKey))}</td>
        <td>${item.count}</td>
        <td>${escapeHtml(couponText)}</td>
        <td>${formatMoney(item.totalAA)}</td>
      </tr>
      ${detailsHtml}
    `;
  }).join("");

  elements.commissionsUnpaidDatesBody.querySelectorAll("tr[data-action='toggle-commission-unpaid-date']").forEach((rowEl) => {
    rowEl.addEventListener("click", () => {
      const dateKey = String(rowEl.dataset.dateKey || "").trim();
      state.commissionsExpandedDate = state.commissionsExpandedDate === dateKey ? "" : dateKey;
      renderCommissionsUnpaidDatesDashboard(rows);
    });
  });
}

function populateCommissionsTravelDateFilter(rows) {
  if (!elements.commissionsTravelDateOptions) return;

  const selectedCoupon = String(elements.commissionsCouponCode?.value || "").trim();
  const previousValues = new Set(state.commissionsSelectedDates);
  const eligibleRows = rows
    .filter(isCommissionRowEligible)
    .filter((row) => String(row.couponCode || "").trim() === selectedCoupon);
  const uniqueDateKeys = [...new Set(eligibleRows.map((row) => normalizeDateKey(row.date)).filter(Boolean))]
    .sort((a, b) => a.localeCompare(b));

  const validSet = new Set(uniqueDateKeys);
  state.commissionsSelectedDates = state.commissionsSelectedDates.filter((dateKey) => validSet.has(dateKey));

  elements.commissionsTravelDateOptions.innerHTML = "";

  if (elements.commissionsTravelDateToggle) {
    elements.commissionsTravelDateToggle.disabled = !selectedCoupon || !uniqueDateKeys.length;
  }

  if (!selectedCoupon) {
    state.commissionsSelectedDates = [];
    elements.commissionsTravelDateOptions.innerHTML = '<div class="multi-date-empty">Select coupon code first.</div>';
    updateCommissionsTravelDateLabel();
    closeCommissionsDateMenu();
    return;
  }

  if (!uniqueDateKeys.length) {
    state.commissionsSelectedDates = [];
    elements.commissionsTravelDateOptions.innerHTML = '<div class="multi-date-empty">No travel dates found.</div>';
    updateCommissionsTravelDateLabel();
    return;
  }

  uniqueDateKeys.forEach((dateKey) => {
    const optionBtn = document.createElement("button");
    optionBtn.type = "button";
    optionBtn.className = "multi-date-option";
    optionBtn.dataset.value = dateKey;
    optionBtn.setAttribute("role", "option");
    optionBtn.setAttribute("aria-selected", previousValues.has(dateKey) || state.commissionsSelectedDates.includes(dateKey) ? "true" : "false");
    optionBtn.innerHTML = `<span class="multi-date-circle" aria-hidden="true"></span><span>${escapeHtml(formatDate(dateKey))}</span>`;
    optionBtn.addEventListener("click", () => {
      toggleCommissionTravelDateSelection(dateKey);
    });

    const isSelected = state.commissionsSelectedDates.includes(dateKey) || (state.commissionsSelectedDates.length === 0 && previousValues.has(dateKey));
    optionBtn.classList.toggle("selected", isSelected);
    optionBtn.setAttribute("aria-selected", isSelected ? "true" : "false");
    elements.commissionsTravelDateOptions.appendChild(optionBtn);
  });

  if (!state.commissionsSelectedDates.length && previousValues.size) {
    state.commissionsSelectedDates = uniqueDateKeys.filter((dateKey) => previousValues.has(dateKey));
  }

  renderCommissionsTravelDateSelectionState();
  updateCommissionsTravelDateLabel();
}

function getSelectedCommissionTravelDates() {
  return state.commissionsSelectedDates.slice();
}

function toggleCommissionTravelDateSelection(dateKey) {
  const selected = new Set(state.commissionsSelectedDates);
  if (selected.has(dateKey)) {
    selected.delete(dateKey);
  } else {
    selected.add(dateKey);
  }

  state.commissionsSelectedDates = [...selected].sort((a, b) => a.localeCompare(b));
  renderCommissionsTravelDateSelectionState();
  updateCommissionsTravelDateLabel();
  handleCommissionsDateChange();
}

function renderCommissionsTravelDateSelectionState() {
  if (!elements.commissionsTravelDateOptions) return;

  const selected = new Set(state.commissionsSelectedDates);
  elements.commissionsTravelDateOptions.querySelectorAll(".multi-date-option").forEach((btn) => {
    const value = String(btn.dataset.value || "");
    const isSelected = selected.has(value);
    btn.classList.toggle("selected", isSelected);
    btn.setAttribute("aria-selected", isSelected ? "true" : "false");
  });
}

function updateCommissionsTravelDateLabel() {
  if (!elements.commissionsTravelDateLabel) return;

  const selectedCoupon = String(elements.commissionsCouponCode?.value || "").trim();
  if (!selectedCoupon) {
    elements.commissionsTravelDateLabel.textContent = "Select coupon code first";
    return;
  }

  const count = state.commissionsSelectedDates.length;
  if (!count) {
    elements.commissionsTravelDateLabel.textContent = "Select travel date(s)";
    return;
  }

  if (count === 1) {
    elements.commissionsTravelDateLabel.textContent = formatDate(state.commissionsSelectedDates[0]);
    return;
  }

  elements.commissionsTravelDateLabel.textContent = `${count} travel dates selected`;
}

function toggleCommissionsDateMenu(event) {
  event?.stopPropagation();
  state.commissionsDateMenuOpen = !state.commissionsDateMenuOpen;
  if (elements.commissionsTravelDateMenu) {
    elements.commissionsTravelDateMenu.hidden = !state.commissionsDateMenuOpen;
  }
  if (elements.commissionsTravelDateToggle) {
    elements.commissionsTravelDateToggle.setAttribute("aria-expanded", state.commissionsDateMenuOpen ? "true" : "false");
  }
}

function closeCommissionsDateMenu() {
  state.commissionsDateMenuOpen = false;
  if (elements.commissionsTravelDateMenu) {
    elements.commissionsTravelDateMenu.hidden = true;
  }
  if (elements.commissionsTravelDateToggle) {
    elements.commissionsTravelDateToggle.setAttribute("aria-expanded", "false");
  }
}

function handleCommissionsDocumentClick(event) {
  if (!state.commissionsDateMenuOpen || !elements.commissionsTravelDateDropdown) return;
  if (!elements.commissionsTravelDateDropdown.contains(event.target)) {
    closeCommissionsDateMenu();
  }
}

function handleCommissionsDateChange() {
  refreshCommissionsSelection();
}

function handleCommissionsCouponChange() {
  const previous = new Set(state.commissionsSelectedDates);
  populateCommissionsTravelDateFilter(state.bookings);
  if (!String(elements.commissionsCouponCode?.value || "").trim()) {
    state.commissionsSelectedDates = [];
  } else {
    const validDates = new Set(
      state.bookings
        .filter(isCommissionRowEligible)
        .filter((row) => String(row.couponCode || "").trim() === String(elements.commissionsCouponCode?.value || "").trim())
        .map((row) => normalizeDateKey(row.date))
        .filter(Boolean)
    );
    state.commissionsSelectedDates = [...previous].filter((dateKey) => validDates.has(dateKey));
  }
  renderCommissionsTravelDateSelectionState();
  updateCommissionsTravelDateLabel();
  refreshCommissionsSelection();
}

function populateCommissionsCouponFilter(rows) {
  if (!elements.commissionsCouponCode) return;

  const previousCoupon = String(elements.commissionsCouponCode.value || "").trim();
  const coupons = [...new Set(
    rows
      .filter(isCommissionRowEligible)
      .map((row) => String(row.couponCode || "").trim())
      .filter(Boolean)
  )].sort((a, b) => a.localeCompare(b));

  elements.commissionsCouponCode.innerHTML = '<option value="">Select coupon code</option>';
  coupons.forEach((coupon) => {
    const option = document.createElement("option");
    option.value = coupon;
    option.textContent = coupon;
    elements.commissionsCouponCode.appendChild(option);
  });

  elements.commissionsCouponCode.disabled = !coupons.length;
  if (previousCoupon && coupons.includes(previousCoupon)) {
    elements.commissionsCouponCode.value = previousCoupon;
  } else {
    elements.commissionsCouponCode.value = "";
  }
}

function refreshCommissionsSelection() {
  const selectedDates = getSelectedCommissionTravelDates();
  if (!elements.commissionsCouponCode) return;

  const selectedCoupon = String(elements.commissionsCouponCode.value || "").trim();
  const matches = state.bookings.filter((row) => {
    if (!isCommissionRowEligible(row)) return false;
    const sameDate = selectedDates.includes(normalizeDateKey(row.date));
    const sameCoupon = String(row.couponCode || "").trim() === selectedCoupon;
    return sameDate && sameCoupon;
  });

  const sumAA = matches.reduce((acc, row) => {
    const amount = normalizeAmountInput(row.commissionAmountAA);
    const guestCount = normalizeGuestCountInput(row.numberOfGuest);
    return acc + (amount === null ? 0 : amount * guestCount);
  }, 0);

  state.activeCommissionsRows = matches;
  state.activeCommissionsSumAA = Number(sumAA.toFixed(2));

  if (elements.commissionsBookingCount) {
    elements.commissionsBookingCount.textContent = String(matches.length);
  }
  if (elements.commissionsSumAA) {
    elements.commissionsSumAA.textContent = formatMoney(state.activeCommissionsSumAA);
  }

  const canEncode = Boolean(selectedDates.length && selectedCoupon && matches.length > 0);

  if (elements.commissionsReceiptAmount) {
    elements.commissionsReceiptAmount.disabled = !canEncode;
    if (!canEncode) {
      elements.commissionsReceiptAmount.value = "";
    }
  }
  if (elements.commissionsInvoiceInput) {
    elements.commissionsInvoiceInput.disabled = !canEncode;
    if (!canEncode) {
      elements.commissionsInvoiceInput.value = "";
    }
  }
  if (elements.commissionsReceiptFile) {
    elements.commissionsReceiptFile.disabled = !canEncode;
    if (!canEncode) {
      elements.commissionsReceiptFile.value = "";
    }
  }
  if (elements.commissionsSaveBtn) {
    elements.commissionsSaveBtn.disabled = !canEncode;
  }

  validateCommissionsReceiptAmount();
}

function validateCommissionsReceiptAmount() {
  if (!elements.commissionsAmountValidation || !elements.commissionsReceiptAmount) return true;

  const raw = String(elements.commissionsReceiptAmount.value || "").trim();
  if (!raw) {
    elements.commissionsAmountValidation.hidden = true;
    elements.commissionsAmountValidation.textContent = "";
    return false;
  }

  const parsed = normalizeAmountInput(raw);
  if (parsed === null) {
    elements.commissionsAmountValidation.hidden = false;
    elements.commissionsAmountValidation.textContent = "Invalid amount format.";
    return false;
  }

  if (parsed !== state.activeCommissionsSumAA) {
    elements.commissionsAmountValidation.hidden = false;
    elements.commissionsAmountValidation.textContent = `Entered amount must exactly match Total Amount (AA × Guests): ${formatMoney(state.activeCommissionsSumAA)}.`;
    return false;
  }

  elements.commissionsAmountValidation.hidden = true;
  elements.commissionsAmountValidation.textContent = "";
  return true;
}

async function handleSaveCommissionsReceipt() {
  const selectedDates = getSelectedCommissionTravelDates();
  const couponCode = String(elements.commissionsCouponCode?.value || "").trim();
  const rawAmount = String(elements.commissionsReceiptAmount?.value || "").trim();
  const invoiceNumber = String(elements.commissionsInvoiceInput?.value || "").trim();
  const matchedRowIndexes = new Set(state.activeCommissionsRows.map((row) => Number(row.rowIndex)));

  if (!selectedDates.length) {
    showToast("Select at least one travel date.", true);
    return;
  }
  if (!couponCode) {
    showToast("Coupon code is required.", true);
    return;
  }
  if (!rawAmount) {
    showToast("Enter the receipt amount.", true);
    return;
  }
  if (!invoiceNumber) {
    showToast("Invoice number is required.", true);
    return;
  }
  if (!validateCommissionsReceiptAmount()) {
    showToast("Entered amount must exactly match Total Amount (AA × Guests).", true);
    return;
  }

  const parsedAmount = normalizeAmountInput(rawAmount);
  if (parsedAmount === null) {
    showToast("Invalid receipt amount.", true);
    return;
  }

  let receiptImageUrl = "";

  setCommissionsSaving(true);

  try {
    const file = elements.commissionsReceiptFile?.files?.[0];
    if (file) {
      receiptImageUrl = await uploadReceiptToCloudinary(file, "commission", `coupon_${couponCode || "na"}`);
    }

    const response = await fetchWithTimeout(API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({
        action: "save_commission_receipt_bulk",
        actorUsername: getActorUsername_(),
        travelDates: selectedDates,
        couponCode,
        invoiceNumber,
        encodedCommissionAmount: String(parsedAmount),
        receiptImageUrl
      })
    }, 20000);

    if (!response.ok) {
      throw new Error(`POST failed: ${response.status}`);
    }

    const result = await parseJsonSafe(response);
    if (result && typeof result === "object" && String(result.status || "").toLowerCase() === "error") {
      throw new Error(result.message || "Failed to save commissions receipt.");
    }

    if (matchedRowIndexes.size) {
      state.bookings = state.bookings.map((row) => {
        if (!matchedRowIndexes.has(Number(row.rowIndex))) return row;
        return {
          ...row,
          influencerInvoiceAK: invoiceNumber,
          influencerStatusAL: "PAID"
        };
      });
      saveCache(state.bookings);
      applyFilters();
    }

    showToast(result.message || "Commissions receipt saved.");

    if (elements.commissionsReceiptAmount) {
      elements.commissionsReceiptAmount.value = "";
    }
    if (elements.commissionsInvoiceInput) {
      elements.commissionsInvoiceInput.value = "";
    }
    if (elements.commissionsReceiptFile) {
      elements.commissionsReceiptFile.value = "";
    }
    validateCommissionsReceiptAmount();
    await refreshCurrentPageDataWithDelay(300);
  } catch (error) {
    debugError("handleSaveCommissionsReceipt", error, { selectedDates, couponCode });
    showToast(error.message || "Failed to save commissions receipt.", true);
  } finally {
    setCommissionsSaving(false);
  }
}

function setCommissionsSaving(isSaving) {
  if (elements.commissionsSaveBtn) {
    elements.commissionsSaveBtn.disabled = isSaving || !state.activeCommissionsRows.length;
    elements.commissionsSaveBtn.textContent = isSaving ? "Saving..." : "Save Commissions Receipt";
  }
  if (elements.commissionsReceiptAmount) {
    elements.commissionsReceiptAmount.disabled = isSaving || !state.activeCommissionsRows.length;
  }
  if (elements.commissionsInvoiceInput) {
    elements.commissionsInvoiceInput.disabled = isSaving || !state.activeCommissionsRows.length;
  }
  if (elements.commissionsReceiptFile) {
    elements.commissionsReceiptFile.disabled = isSaving || !state.activeCommissionsRows.length;
  }
  if (elements.commissionsTravelDateToggle) {
    const hasCouponSelected = Boolean(String(elements.commissionsCouponCode?.value || "").trim());
    elements.commissionsTravelDateToggle.disabled = isSaving || !hasCouponSelected;
  }
  if (isSaving) {
    closeCommissionsDateMenu();
  }
  if (elements.commissionsCouponCode) {
    elements.commissionsCouponCode.disabled = isSaving;
  }
  if (elements.commissionsSavingIndicator) {
    elements.commissionsSavingIndicator.hidden = !isSaving;
  }
}

function renderTable(rows, columns) {
  if (!elements.tableHead || !elements.bookingsBody) return;

  elements.tableHead.innerHTML = `<tr>${columns.map((col) => `<th>${COLUMN_LABELS[col]}</th>`).join("")}</tr>`;

  if (!rows.length) {
    elements.bookingsBody.innerHTML = `
      <tr>
        <td colspan="${columns.length}" class="empty-state">No records found.</td>
      </tr>
    `;
    return;
  }

  elements.bookingsBody.innerHTML = rows
    .map((row) => {
      const hasFullInvoice = Boolean(String(row.receiptFullPayment || "").trim());
      const rowClass = PAGE === "payment-receipt-encoding" && hasFullInvoice ? "invoice-row-complete" : "";
      const statusClass = getStatusRowClass(row);
      const isPendingPaymentRow = String(row.adData || "").trim().toUpperCase() === "PENDING PAYMENT";
      const isRowClickable = PAGE === "payment-receipt-encoding" || (PAGE === "bookings-dashboard" && isPendingPaymentRow);
      const clickable = isRowClickable ? "receipt-click-row" : "";
      const dataAttr = isRowClickable ? `data-row-index="${row.rowIndex}"` : "";
      return `<tr class="${rowClass} ${statusClass} ${clickable}" ${dataAttr}>${columns.map((col) => `<td>${cellHtml(col, row)}</td>`).join("")}</tr>`;
    })
    .join("");

  bindRowActions();
}

function cellHtml(col, row) {
  switch (col) {
    case "name": return escapeHtml(row.name || "-");
    case "date": return escapeHtml(formatDate(row.date));
    case "tour": return escapeHtml(row.tour || "-");
    case "guests": return escapeHtml(row.numberOfGuest || "-");
    case "coupon": return escapeHtml(row.couponCode || "-");
    case "discount": return escapeHtml(formatMoney(row.discountAmount));
    case "paymentMethod": return escapeHtml(row.paymentMethod || "-");
    case "totalAmount": {
      const amountText = formatMoney(row.totalAmount);
      return escapeHtml(amountText);
    }
    case "downpayment": {
      const amountText = formatMoney(row.downpaymentAmount);
      if (PAGE !== "bookings-dashboard") return escapeHtml(amountText);
      const hasDownInvoice = Boolean(String(row.receiptDownpayment || "").trim());
      if (!hasDownInvoice) return "-";
      const url = getValidHttpUrl(row.receiptUrlAM);
      if (!url) return escapeHtml(amountText);
      return `<button class="receipt-link-btn" data-action="open-receipt-image" data-url="${escapeHtmlAttr(url)}" data-label="Downpayment Receipt">${escapeHtml(amountText)}</button>`;
    }
    case "totalBalance": {
      const amountText = formatMoney(row.totalBalance);
      if (PAGE !== "bookings-dashboard") return escapeHtml(amountText);
      const hasFullInvoice = Boolean(String(row.receiptFullPayment || "").trim());
      if (!hasFullInvoice) return escapeHtml(amountText);
      const url = getValidHttpUrl(row.receiptUrlAN);
      if (!url) return escapeHtml(amountText);
      return `<button class="receipt-link-btn" data-action="open-receipt-image" data-url="${escapeHtmlAttr(url)}" data-label="Full Payment Receipt">${escapeHtml(amountText)}</button>`;
    }
    case "adData": return escapeHtml(row.adData || "-");
    case "phone": return escapeHtml(row.phone || "-");
    case "remarksNew": return escapeHtml(row.remarksNew || "-");
    case "remarksConfirmed": return escapeHtml(row.remarksConfirmed || "-");
    case "receiptDown": {
      const hasValue = Boolean(String(row.receiptDownpayment || "").trim());
      const cls = hasValue ? "invoice-cell invoice-filled" : "invoice-cell";
      return `<span class="${cls}">${escapeHtml(row.receiptDownpayment || "-")}</span>`;
    }
    case "receiptFull": {
      const hasValue = Boolean(String(row.receiptFullPayment || "").trim());
      const cls = hasValue ? "invoice-cell invoice-filled" : "invoice-cell";
      return `<span class="${cls}">${escapeHtml(row.receiptFullPayment || "-")}</span>`;
    }
    case "actionNew": {
      const disabled = state.sendingRows.has(row.rowIndex) ? "disabled" : "";
      const label = state.sendingRows.has(row.rowIndex) ? "Sending..." : "Send";
      return `<button class="send-btn" data-action="send-new" data-row-index="${row.rowIndex}" data-url="${escapeHtmlAttr(row.whatsappNew)}" ${disabled}>${label}</button>`;
    }
    case "actionConfirmed": {
      const disabled = state.sendingRows.has(row.rowIndex) ? "disabled" : "";
      const label = state.sendingRows.has(row.rowIndex) ? "Sending..." : "Send";
      return `<button class="send-btn" data-action="send-confirmed" data-row-index="${row.rowIndex}" data-url="${escapeHtmlAttr(row.whatsappConfirmed)}" ${disabled}>${label}</button>`;
    }
    default:
      return "-";
  }
}

function bindRowActions() {
  document.querySelectorAll("button[data-action='send-new']").forEach((btn) => {
    btn.addEventListener("click", async () => {
      const rowIndex = Number(btn.dataset.rowIndex);
      const url = btn.dataset.url || "";
      await handleSend(rowIndex, url, "new");
    });
  });

  document.querySelectorAll("button[data-action='send-confirmed']").forEach((btn) => {
    btn.addEventListener("click", async () => {
      const rowIndex = Number(btn.dataset.rowIndex);
      const url = btn.dataset.url || "";
      await handleSend(rowIndex, url, "confirmed");
    });
  });

  document.querySelectorAll("button[data-action='save-receipt']").forEach((btn) => {
    btn.addEventListener("click", async () => {
      const rowIndex = Number(btn.dataset.rowIndex);
      await handleSaveReceipt(rowIndex);
    });
  });

  document.querySelectorAll("button[data-action='open-receipt-image']").forEach((btn) => {
    btn.addEventListener("click", (event) => {
      event.stopPropagation();
      const url = getValidHttpUrl(btn.dataset.url || "");
      if (!url) {
        showToast("Receipt image URL is missing or invalid.", true);
        return;
      }
      openReceiptImageViewer(url, btn.dataset.label || "Receipt Image");
    });
  });

  if (PAGE === "payment-receipt-encoding") {
    document.querySelectorAll("tr[data-row-index]").forEach((rowEl) => {
      rowEl.addEventListener("click", () => {
        const rowIndex = Number(rowEl.dataset.rowIndex);
        openReceiptModal(rowIndex);
      });
    });
    return;
  }

  if (PAGE === "bookings-dashboard") {
    document.querySelectorAll("tr[data-row-index]").forEach((rowEl) => {
      rowEl.addEventListener("click", () => {
        const rowIndex = Number(rowEl.dataset.rowIndex);
        openCancelReasonModal(rowIndex);
      });
    });
  }
}

function openCancelReasonModal(rowIndex) {
  if (!elements.cancelReasonModal) return;

  const booking = state.bookings.find((row) => Number(row.rowIndex) === Number(rowIndex));
  if (!booking) return;

  state.activeCancelRowIndex = Number(rowIndex);

  if (elements.cancelReasonBookingInfo) {
    elements.cancelReasonBookingInfo.textContent = `${booking.name || "-"} • ${formatDate(booking.date)} • ${booking.tour || "-"}`;
  }

  if (elements.cancelReasonSlot) {
    elements.cancelReasonSlot.value = "cancel";
  }

  if (elements.cancelReasonInput) {
    elements.cancelReasonInput.value = String(booking.cancelReasonAO || "").trim();
    elements.cancelReasonInput.focus();
  }

  updateCancelReasonInputVisibility();
  elements.cancelReasonModal.hidden = false;
}

function closeCancelReasonModal() {
  if (!elements.cancelReasonModal) return;
  elements.cancelReasonModal.hidden = true;
  state.activeCancelRowIndex = null;
}

function updateCancelReasonInputVisibility() {
  const slot = String(elements.cancelReasonSlot?.value || "").trim().toLowerCase();
  const shouldShow = slot === "cancel";
  if (elements.cancelReasonInputGroup) {
    elements.cancelReasonInputGroup.hidden = !shouldShow;
  }
}

async function handleSaveCancelReason() {
  const rowIndex = Number(state.activeCancelRowIndex);
  if (!rowIndex) {
    closeCancelReasonModal();
    return;
  }

  const slot = String(elements.cancelReasonSlot?.value || "").trim().toLowerCase();
  if (slot !== "cancel") {
    closeCancelReasonModal();
    return;
  }

  const cancelReason = String(elements.cancelReasonInput?.value || "").trim();
  if (!cancelReason) {
    showToast("Reason is required for canceled booking.", true);
    return;
  }

  if (elements.cancelReasonSaveBtn) {
    elements.cancelReasonSaveBtn.disabled = true;
    elements.cancelReasonSaveBtn.textContent = "Saving...";
  }

  try {
    const response = await fetchWithTimeout(API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({
        action: "save_cancel_reason",
        actorUsername: getActorUsername_(),
        rowIndex,
        cancelReason
      })
    }, 15000);

    if (!response.ok) {
      throw new Error(`POST failed: ${response.status}`);
    }

    const result = await parseJsonSafe(response);
    if (result && typeof result === "object" && String(result.status || "").toLowerCase() === "error") {
      throw new Error(result.message || "Failed to save cancel reason.");
    }

    state.bookings = state.bookings.map((item) => {
      if (Number(item.rowIndex) !== rowIndex) return item;
      return {
        ...item,
        cancelReasonAO: cancelReason
      };
    });

    saveCache(state.bookings);
    applyFilters();
    closeCancelReasonModal();
    showToast(result?.message || "Cancel reason saved.");
    await refreshCurrentPageDataWithDelay(300);
  } catch (error) {
    debugError("handleSaveCancelReason", error, { rowIndex });
    showToast(error.message || "Failed to save cancel reason.", true);
  } finally {
    if (elements.cancelReasonSaveBtn) {
      elements.cancelReasonSaveBtn.disabled = false;
      elements.cancelReasonSaveBtn.textContent = "Save";
    }
  }
}

function handleModalInvoiceClick(slot) {
  const rowIndex = Number(state.activeReceiptRowIndex);
  if (!rowIndex) return;

  const booking = state.bookings.find((row) => Number(row.rowIndex) === rowIndex);
  if (!booking) return;

  const isFull = slot === "full";
  const invoiceCode = String(isFull ? booking.receiptFullPayment : booking.receiptDownpayment || "").trim();
  if (!invoiceCode) return;

  const url = getValidHttpUrl(isFull ? booking.receiptUrlAN : booking.receiptUrlAM);
  if (!url) {
    showToast(`Invoice exists but no receipt image URL found in column ${isFull ? "AN" : "AM"}.`, true);
    return;
  }

  openReceiptImageViewer(url, isFull ? "Full Payment Receipt" : "Downpayment Receipt");
}

function openReceiptImageViewer(url, label) {
  if (!elements.receiptImageViewer || !elements.receiptImageViewerImg) {
    window.open(url, "_blank", "noopener,noreferrer");
    return;
  }

  const title = document.getElementById("receiptImageViewerTitle");
  if (title) {
    title.textContent = label || "Receipt Image";
  }

  elements.receiptImageViewerImg.src = url;
  elements.receiptImageViewer.hidden = false;
}

function closeReceiptImageViewer() {
  if (!elements.receiptImageViewer) return;
  if (elements.receiptImageViewerImg) {
    elements.receiptImageViewerImg.removeAttribute("src");
  }
  elements.receiptImageViewer.hidden = true;
}

async function handleSend(rowIndex, whatsapp, mode) {
  const url = getValidWhatsAppUrl(whatsapp);
  if (!url) {
    showToast(`Missing or invalid WhatsApp link in ${mode === "new" ? "AF" : "AH"}.`, true);
    return;
  }

  const confirmed = await askSendConfirmation("Open WhatsApp and mark this row as sent?");
  if (!confirmed) return;

  window.open(url, "_blank", "noopener,noreferrer");

  state.sendingRows.add(rowIndex);
  applyFilters();

  try {
    const response = await fetchWithTimeout(API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({
        action: mode === "new" ? "mark_sent_new" : "mark_sent_confirmed",
        rowIndex,
        actorUsername: getActorUsername_()
      })
    }, 12000);

    if (!response.ok) {
      throw new Error(`POST failed: ${response.status}`);
    }

    const sendResult = await parseJsonSafe(response);
    if (sendResult && typeof sendResult === "object" && String(sendResult.status || "").toLowerCase() === "error") {
      throw new Error(sendResult.message || "Failed to mark as sent.");
    }

    state.bookings = state.bookings.map((item) => {
      if (item.rowIndex !== rowIndex) return item;
      return {
        ...item,
        remarksNew: mode === "new" ? "sent" : item.remarksNew,
        remarksConfirmed: mode === "confirmed" ? "sent" : item.remarksConfirmed
      };
    });

    saveCache(state.bookings);
    applyFilters();
    showToast("Marked as sent.");
    await refreshCurrentPageDataWithDelay(300);
  } catch (error) {
    debugError("handleSend", error, { rowIndex, mode });
    showToast(error.message || "Failed to mark as sent.", true);
  } finally {
    state.sendingRows.delete(rowIndex);
    applyFilters();
  }
}

async function handleSaveReceipt() {
  const rowIndex = Number(state.activeReceiptRowIndex);
  if (!rowIndex) return;

  const paymentType = elements.receiptTypeSelect?.value || "downpayment";
  const invoiceNumber = String(elements.receiptInvoiceInput?.value || "").trim();
  const existing = state.bookings.find((row) => row.rowIndex === rowIndex);
  const encodedDownpaymentRaw = String(elements.downpaymentAmountInput?.value || "").trim();
  const encodedFullPaymentRaw = String(elements.fullPaymentAmountInput?.value || "").trim();
  let encodedDownpaymentAmount = "";
  let encodedFullPaymentAmount = "";

  if (existing) {
    const hasDown = Boolean(String(existing.receiptDownpayment || "").trim());
    const hasFull = Boolean(String(existing.receiptFullPayment || "").trim());

    if (hasFull) {
      showToast("Full payment invoice already exists in column AC. Encoding is locked.", true);
      return;
    }

    if (hasDown && hasFull) {
      showToast("Both Downpayment and Full Payment invoices are already encoded.", true);
      return;
    }

    if (paymentType === "downpayment" && hasDown) {
      showToast("Downpayment invoice already exists in column Y.", true);
      return;
    }

    if (paymentType === "full" && hasFull) {
      showToast("Full payment invoice already exists in column AC.", true);
      return;
    }
  }

  if (!invoiceNumber) {
    showToast("Invoice number is required.", true);
    return;
  }

  if (paymentType === "downpayment" && encodedDownpaymentRaw) {
    const parsed = normalizeAmountInput(encodedDownpaymentRaw);
    if (parsed === null) {
      showToast("Invalid downpayment amount.", true);
      return;
    }
    if (parsed < state.activeDownpaymentMinimum) {
      showValidationNote(`Downpayment should start at ${formatMoney(state.activeDownpaymentMinimum)} (column X).`);
      showToast("Downpayment should start at the amount encoded in column X.", true);
      return;
    }
    encodedDownpaymentAmount = String(parsed);
  }

  if (paymentType === "downpayment" && !encodedDownpaymentRaw) {
    showValidationNote(`Downpayment should start at ${formatMoney(state.activeDownpaymentMinimum)} (column X).`);
    showToast("Enter a downpayment amount.", true);
    return;
  }

  if (paymentType === "full") {
    if (!encodedFullPaymentRaw) {
      showFullValidationNote(`Full payment amount must be exactly ${formatMoney(state.activeFullPaymentExact)} (column Z).`);
      showToast("Enter full payment amount.", true);
      return;
    }

    const parsedFull = normalizeAmountInput(encodedFullPaymentRaw);
    if (parsedFull === null) {
      showFullValidationNote("Invalid full payment amount.");
      showToast("Invalid full payment amount.", true);
      return;
    }

    if (parsedFull !== state.activeFullPaymentExact) {
      showFullValidationNote(`Full payment amount must be exactly ${formatMoney(state.activeFullPaymentExact)} (column Z).`);
      showToast("Full payment amount must exactly match column Z.", true);
      return;
    }

    encodedFullPaymentAmount = String(parsedFull);
  }

  let receiptImageUrl = "";

  setReceiptSaving(true);

  try {
    const file = elements.receiptFileInput?.files?.[0];
    if (file) {
      receiptImageUrl = await uploadReceiptToCloudinary(file, paymentType, rowIndex);
    }

    const requestPayload = {
      action: "save_receipt",
      actorUsername: getActorUsername_(),
      rowIndex,
      paymentType,
      invoiceNumber,
      encodedDownpaymentAmount,
      encodedFullPaymentAmount,
      receiptImageUrl
    };

    const response = await fetchWithTimeout(API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(requestPayload)
    }, 20000);

    const rawBody = await response.text();
    let result = null;

    try {
      result = rawBody ? JSON.parse(rawBody) : null;
    } catch {
      const parseError = new Error(`API returned invalid JSON (HTTP ${response.status}).`);
      parseError.httpStatus = response.status;
      parseError.responseBodyPreview = String(rawBody).slice(0, 1200);
      parseError.requestPayload = requestPayload;
      throw parseError;
    }

    if (!response.ok) {
      const httpError = new Error(result?.message || `POST failed: ${response.status}`);
      httpError.httpStatus = response.status;
      httpError.responseData = result;
      httpError.responseBodyPreview = String(rawBody).slice(0, 1200);
      httpError.requestPayload = requestPayload;
      throw httpError;
    }

    if (result && typeof result === "object" && String(result.status || "").toLowerCase() === "error") {
      const apiError = new Error(result.message || "Failed to save receipt encoding.");
      apiError.httpStatus = response.status;
      apiError.responseData = result;
      apiError.responseBodyPreview = String(rawBody).slice(0, 1200);
      apiError.requestPayload = requestPayload;
      throw apiError;
    }

    if (result?.driveWarning) {
      debugInfo("handleSaveReceipt:driveWarning", {
        rowIndex,
        paymentType,
        driveWarning: result.driveWarning
      });
    }

    state.bookings = state.bookings.map((item) => {
      if (item.rowIndex !== rowIndex) return item;
      return {
        ...item,
        downpaymentAmount: paymentType === "downpayment" && encodedDownpaymentAmount
          ? encodedDownpaymentAmount
          : item.downpaymentAmount,
        receiptDownpayment: paymentType === "downpayment" ? invoiceNumber : item.receiptDownpayment,
        receiptFullPayment: paymentType === "full" ? invoiceNumber : item.receiptFullPayment,
        receiptUrlAM: paymentType === "downpayment" && receiptImageUrl ? receiptImageUrl : item.receiptUrlAM,
        receiptUrlAN: paymentType === "full" && receiptImageUrl ? receiptImageUrl : item.receiptUrlAN
      };
    });

    saveCache(state.bookings);
    applyFilters();
    closeReceiptModal();
    showToast(result?.message || "Receipt encoded successfully.");
    await refreshCurrentPageDataWithDelay(300);
  } catch (error) {
    debugError("handleSaveReceipt", error, {
      rowIndex,
      paymentType,
      apiUrl: API_URL,
      httpStatus: error?.httpStatus || null,
      responseData: error?.responseData || null,
      responseBodyPreview: error?.responseBodyPreview || null,
      requestPayload: error?.requestPayload || null
    });
    showToast(error.message || "Failed to save receipt encoding.", true);
  } finally {
    setReceiptSaving(false);
  }
}

function openReceiptModal(rowIndex) {
  const existing = state.bookings.find((row) => row.rowIndex === rowIndex);
  if (!existing || !elements.receiptModal) return;

  state.activeReceiptRowIndex = rowIndex;

  const hasDown = Boolean(String(existing.receiptDownpayment || "").trim());
  const hasFull = Boolean(String(existing.receiptFullPayment || "").trim());
  const lockAll = hasFull;
  state.activeDownpaymentMinimum = getDownpaymentMinimum(existing.downpaymentAmount);
  state.activeFullPaymentExact = getDownpaymentMinimum(existing.totalBalance);

  if (elements.receiptBookingInfo) {
    elements.receiptBookingInfo.textContent = `${existing.name || "-"} • ${existing.tour || "-"} • Balance ${formatMoney(existing.totalBalance)}`;
  }
  if (elements.receiptDownValue) {
    elements.receiptDownValue.textContent = existing.receiptDownpayment || "-";
    elements.receiptDownValue.classList.toggle("invoice-filled", hasDown);
    elements.receiptDownValue.style.cursor = hasDown ? "pointer" : "default";
    elements.receiptDownValue.title = hasDown
      ? (getValidHttpUrl(existing.receiptUrlAM) ? "Open downpayment receipt image" : "No image URL in column AM")
      : "";
  }
  if (elements.receiptFullValue) {
    elements.receiptFullValue.textContent = existing.receiptFullPayment || "-";
    elements.receiptFullValue.classList.toggle("invoice-filled", hasFull);
    elements.receiptFullValue.style.cursor = hasFull ? "pointer" : "default";
    elements.receiptFullValue.title = hasFull
      ? (getValidHttpUrl(existing.receiptUrlAN) ? "Open full payment receipt image" : "No image URL in column AN")
      : "";
  }

  if (elements.receiptTypeSelect) {
    const downOpt = elements.receiptTypeSelect.querySelector("option[value='downpayment']");
    const fullOpt = elements.receiptTypeSelect.querySelector("option[value='full']");
    if (downOpt) downOpt.disabled = hasDown;
    if (fullOpt) fullOpt.disabled = hasFull;
    elements.receiptTypeSelect.value = hasDown && !hasFull ? "full" : "downpayment";
    elements.receiptTypeSelect.disabled = lockAll;
  }

  if (elements.downpaymentCurrentValue) {
    elements.downpaymentCurrentValue.textContent = formatMoney(existing.downpaymentAmount);
  }
  if (elements.downpaymentAmountInput) {
    elements.downpaymentAmountInput.value = "";
    elements.downpaymentAmountInput.disabled = lockAll;
  }
  if (elements.fullPaymentCurrentValue) {
    elements.fullPaymentCurrentValue.textContent = formatMoney(existing.totalBalance);
  }
  if (elements.fullPaymentAmountInput) {
    elements.fullPaymentAmountInput.value = "";
    elements.fullPaymentAmountInput.disabled = lockAll;
  }
  hideValidationNote();
  hideFullValidationNote();

  if (elements.receiptInvoiceInput) {
    elements.receiptInvoiceInput.value = "";
    elements.receiptInvoiceInput.disabled = lockAll;
  }
  if (elements.receiptFileInput) {
    elements.receiptFileInput.value = "";
    elements.receiptFileInput.disabled = lockAll;
  }
  if (elements.receiptSaveBtn) {
    elements.receiptSaveBtn.disabled = lockAll;
    elements.receiptSaveBtn.textContent = lockAll ? "Completed" : "Save";
  }

  if (elements.receiptSavingIndicator) {
    elements.receiptSavingIndicator.hidden = true;
  }

  updateDownpaymentAmountVisibility();

  elements.receiptModal.hidden = false;
}

function closeReceiptModal() {
  if (elements.receiptModal) {
    elements.receiptModal.hidden = true;
  }
  state.activeReceiptRowIndex = null;
  state.activeDownpaymentMinimum = 0;
  state.activeFullPaymentExact = 0;
  hideValidationNote();
  hideFullValidationNote();
  setReceiptSaving(false);
}

function getStatusRowClass(row) {
  if (PAGE !== "bookings-dashboard") return "";

  const status = String(row.adData || "").trim().toUpperCase();
  if (status === "BOOKING COMPLETE" || status === "BOOKING COMPLETED") {
    return "status-complete-row";
  }
  if (status === "PENDING PAYMENT") {
    return "status-pending-row";
  }
  return "";
}

function setReceiptSaving(isSaving) {
  if (elements.receiptSaveBtn) {
    elements.receiptSaveBtn.disabled = isSaving;
    elements.receiptSaveBtn.textContent = isSaving ? "Saving..." : "Save";
  }
  if (elements.receiptCancelBtn) {
    elements.receiptCancelBtn.disabled = isSaving;
  }
  if (elements.receiptSavingIndicator) {
    elements.receiptSavingIndicator.hidden = !isSaving;
  }
}

async function askSendConfirmation(message) {
  if (!elements.sendConfirmModal || !elements.sendConfirmOk || !elements.sendConfirmCancel) {
    return window.confirm(message);
  }

  if (elements.sendConfirmMessage) {
    elements.sendConfirmMessage.textContent = message;
  }

  elements.sendConfirmModal.hidden = false;

  return new Promise((resolve) => {
    state.sendConfirmResolver = resolve;
  });
}

function resolveSendConfirm(result) {
  if (elements.sendConfirmModal) {
    elements.sendConfirmModal.hidden = true;
  }

  if (typeof state.sendConfirmResolver === "function") {
    state.sendConfirmResolver(Boolean(result));
  }
  state.sendConfirmResolver = null;
}

function updateDownpaymentAmountVisibility() {
  if (!elements.receiptTypeSelect) return;

  const showDown = elements.receiptTypeSelect.value === "downpayment" && !elements.receiptTypeSelect.disabled;
  const showFull = elements.receiptTypeSelect.value === "full" && !elements.receiptTypeSelect.disabled;

  if (elements.downpaymentAmountGroup) {
    elements.downpaymentAmountGroup.hidden = !showDown;
  }
  if (elements.fullPaymentAmountGroup) {
    elements.fullPaymentAmountGroup.hidden = !showFull;
  }

  if (showDown) {
    validateDownpaymentAmount();
  } else {
    hideValidationNote();
  }

  if (showFull) {
    validateFullPaymentAmount();
  } else {
    hideFullValidationNote();
  }
}

function validateDownpaymentAmount() {
  if (!elements.receiptTypeSelect || elements.receiptTypeSelect.value !== "downpayment") {
    hideValidationNote();
    return true;
  }

  const raw = String(elements.downpaymentAmountInput?.value || "").trim();
  if (!raw) {
    hideValidationNote();
    return true;
  }

  const parsed = normalizeAmountInput(raw);
  if (parsed === null) {
    showValidationNote("Invalid amount format.");
    return false;
  }

  if (parsed < state.activeDownpaymentMinimum) {
    showValidationNote(`Downpayment should start at ${formatMoney(state.activeDownpaymentMinimum)} (column X).`);
    return false;
  }

  if (parsed === state.activeFullPaymentExact) {
    const fullOption = elements.receiptTypeSelect.querySelector("option[value='full']");
    if (fullOption && !fullOption.disabled) {
      elements.receiptTypeSelect.value = "full";
      if (elements.fullPaymentAmountInput) {
        elements.fullPaymentAmountInput.value = String(parsed);
      }
      updateDownpaymentAmountVisibility();
      showToast("Amount matches column Z. Switched to Full Payment.");
      hideValidationNote();
      return true;
    }
  }

  hideValidationNote();
  return true;
}

function showValidationNote(message) {
  if (!elements.downpaymentValidationNote) return;
  elements.downpaymentValidationNote.textContent = message;
  elements.downpaymentValidationNote.hidden = false;
}

function hideValidationNote() {
  if (!elements.downpaymentValidationNote) return;
  elements.downpaymentValidationNote.hidden = true;
  elements.downpaymentValidationNote.textContent = "";
}

function validateFullPaymentAmount() {
  if (!elements.receiptTypeSelect || elements.receiptTypeSelect.value !== "full") {
    hideFullValidationNote();
    return true;
  }

  const raw = String(elements.fullPaymentAmountInput?.value || "").trim();
  if (!raw) {
    hideFullValidationNote();
    return true;
  }

  const parsed = normalizeAmountInput(raw);
  if (parsed === null) {
    showFullValidationNote("Invalid full payment amount.");
    return false;
  }

  if (parsed !== state.activeFullPaymentExact) {
    showFullValidationNote(`Full payment amount must be exactly ${formatMoney(state.activeFullPaymentExact)} (column Z).`);
    return false;
  }

  hideFullValidationNote();
  return true;
}

function showFullValidationNote(message) {
  if (!elements.fullPaymentValidationNote) return;
  elements.fullPaymentValidationNote.textContent = message;
  elements.fullPaymentValidationNote.hidden = false;
}

function hideFullValidationNote() {
  if (!elements.fullPaymentValidationNote) return;
  elements.fullPaymentValidationNote.hidden = true;
  elements.fullPaymentValidationNote.textContent = "";
}

function populateTourFilter(rows) {
  if (!elements.tourFilter) return;

  const previousValue = elements.tourFilter.value;
  const tours = [...new Set(rows.map((row) => row.tour).filter(Boolean))].sort((a, b) => a.localeCompare(b));

  elements.tourFilter.innerHTML = '<option value="">All Tours</option>';
  tours.forEach((tour) => {
    const option = document.createElement("option");
    option.value = tour;
    option.textContent = tour;
    elements.tourFilter.appendChild(option);
  });

  if (tours.includes(previousValue)) {
    elements.tourFilter.value = previousValue;
  }
}

function populateTravelDateFilter(rows) {
  if (!elements.travelDateFilter) return;

  const previousValue = elements.travelDateFilter.value;
  const uniqueDateKeys = [...new Set(rows.map((row) => normalizeDateKey(row.date)).filter(Boolean))]
    .sort((a, b) => a.localeCompare(b));

  elements.travelDateFilter.innerHTML = '<option value="">All Dates</option>';

  uniqueDateKeys.forEach((dateKey) => {
    const option = document.createElement("option");
    option.value = dateKey;
    option.textContent = formatDate(dateKey);
    elements.travelDateFilter.appendChild(option);
  });

  if (uniqueDateKeys.includes(previousValue)) {
    elements.travelDateFilter.value = previousValue;
  }
}

function updatePendingCount(count) {
  if (!elements.pendingCount) return;
  elements.pendingCount.textContent = String(count);
}

function setLoading(isLoading) {
  if (elements.loadingOverlay) {
    elements.loadingOverlay.hidden = !isLoading;
  }
  if (elements.refreshBtn) {
    elements.refreshBtn.disabled = Boolean(isLoading);
    elements.refreshBtn.textContent = isLoading ? "Refreshing..." : "Refresh";
  }
}

function showToast(message, isError = false) {
  if (!elements.toast) return;

  elements.toast.textContent = message;
  elements.toast.style.background = isError ? "var(--danger)" : "#111827";
  elements.toast.hidden = false;

  clearTimeout(showToast.timeout);
  showToast.timeout = setTimeout(() => {
    elements.toast.hidden = true;
  }, 2500);
}

function hydrateFromCache() {
  try {
    const raw = localStorage.getItem(`${CACHE_KEY}_${PAGE}`);
    if (!raw) return;
    const cached = JSON.parse(raw);
    if (!Array.isArray(cached)) return;

    state.bookings = cached.map(normalizeRow);
    populateTourFilter(state.bookings);
    applyFilters();
  } catch (error) {
    debugError("hydrateFromCache", error);
  }
}

function saveCache(rows) {
  try {
    localStorage.setItem(`${CACHE_KEY}_${PAGE}`, JSON.stringify(rows));
  } catch (error) {
    debugError("saveCache", error);
  }
}

async function fetchWithTimeout(url, options, timeoutMs) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), timeoutMs);

  try {
    return await fetch(url, { ...options, signal: controller.signal });
  } catch (error) {
    if (error.name === "AbortError") {
      throw new Error("API timeout. Check Apps Script deployment and internet.");
    }
    throw new Error("Cannot reach API. Check Web App URL and deployment access (Anyone).");
  } finally {
    clearTimeout(timeoutId);
  }
}

async function parseJsonSafe(response) {
  const text = await response.text();
  try {
    return JSON.parse(text);
  } catch {
    debugError("parseJsonSafe", new Error("Invalid JSON response"), {
      status: response.status,
      bodyPreview: String(text).slice(0, 280)
    });
    throw new Error("API did not return JSON.");
  }
}

function isApiConfigured() {
  const url = String(API_URL || "").trim();
  return url && url !== "PASTE_YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE" && /^https:\/\//i.test(url);
}

function getValidWhatsAppUrl(value) {
  const url = String(value || "").trim();
  if (!url) return "";
  if (!/^https?:\/\//i.test(url)) return "";
  if (!/wa\.me|whatsapp\.com/i.test(url)) return "";
  return url;
}

function getValidHttpUrl(value) {
  const url = String(value || "").trim();
  if (!url) return "";
  if (!/^https?:\/\//i.test(url)) return "";
  return url;
}

function formatDate(value) {
  if (!value) return "-";
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return value;
  return parsed.toLocaleDateString(undefined, { year: "numeric", month: "short", day: "numeric" });
}

function normalizeDateKey(value) {
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return String(value || "").trim();
  const year = parsed.getFullYear();
  const month = String(parsed.getMonth() + 1).padStart(2, "0");
  const day = String(parsed.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function formatMoney(value) {
  const text = String(value || "").trim();
  if (!text) return "-";
  const numeric = Number(text.replace(/[^0-9.-]/g, ""));
  if (!Number.isFinite(numeric)) return text;
  return `₱${numeric.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

function normalizeAmountInput(value) {
  const cleaned = String(value).replace(/[^0-9.-]/g, "").trim();
  if (!cleaned) return null;
  const numeric = Number(cleaned);
  if (!Number.isFinite(numeric) || numeric < 0) return null;
  return Number(numeric.toFixed(2));
}

function normalizeGuestCountInput(value) {
  const cleaned = String(value || "").replace(/[^0-9]/g, "").trim();
  const parsed = Number(cleaned);
  if (!Number.isFinite(parsed) || parsed <= 0) return 1;
  return Math.floor(parsed);
}

function getDownpaymentMinimum(value) {
  const parsed = normalizeAmountInput(value);
  return parsed === null ? 0 : parsed;
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function escapeHtmlAttr(value) {
  return escapeHtml(value).replace(/`/g, "");
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = String(reader.result || "");
      const base64 = result.includes(",") ? result.split(",")[1] : result;
      resolve({ base64, mimeType: file.type || "application/octet-stream" });
    };
    reader.onerror = () => reject(new Error("Failed to read receipt image."));
    reader.readAsDataURL(file);
  });
}

async function uploadReceiptToCloudinary(file, paymentType, rowIndex) {
  const cloudName = String(CLOUDINARY_CLOUD_NAME || "").trim();
  const uploadPreset = String(CLOUDINARY_UPLOAD_PRESET || "").trim();

  if (!cloudName || cloudName === "PASTE_CLOUDINARY_CLOUD_NAME") {
    throw new Error("Cloudinary cloud name is not configured in script.js");
  }

  if (!uploadPreset || uploadPreset === "PASTE_CLOUDINARY_UNSIGNED_UPLOAD_PRESET") {
    throw new Error("Cloudinary upload preset is not configured in script.js");
  }

  const formData = new FormData();
  formData.append("file", file);
  formData.append("upload_preset", uploadPreset);
  formData.append("folder", CLOUDINARY_FOLDER);
  formData.append("public_id", `row_${rowIndex}_${paymentType}_${Date.now()}`);

  const endpoint = `https://api.cloudinary.com/v1_1/${encodeURIComponent(cloudName)}/image/upload`;
  const response = await fetchWithTimeout(endpoint, {
    method: "POST",
    body: formData
  }, 30000);

  const rawBody = await response.text();
  let payload;

  try {
    payload = rawBody ? JSON.parse(rawBody) : null;
  } catch {
    throw new Error(`Cloudinary returned invalid JSON (HTTP ${response.status}).`);
  }

  if (!response.ok) {
    const cloudinaryMessage = payload?.error?.message || payload?.message || `HTTP ${response.status}`;
    throw new Error(`Cloudinary upload failed: ${cloudinaryMessage}`);
  }

  const secureUrl = String(payload?.secure_url || "").trim();
  if (!secureUrl) {
    throw new Error("Cloudinary upload succeeded but secure_url is missing.");
  }

  return secureUrl;
}

function setupGlobalErrorHandlers() {
  window.addEventListener("error", (event) => {
    if (isExtensionNoise(event.filename, event.message)) return;
    debugError("window:error", event.error || new Error(event.message), {
      source: event.filename,
      line: event.lineno,
      column: event.colno
    });
  });

  window.addEventListener("unhandledrejection", (event) => {
    const reason = event.reason instanceof Error ? event.reason : new Error(String(event.reason));
    if (isExtensionNoise("", reason.message)) return;
    debugError("window:unhandledrejection", reason);
  });
}

function isExtensionNoise(source, message) {
  const src = String(source || "").toLowerCase();
  const msg = String(message || "").toLowerCase();
  return (
    src.startsWith("chrome-extension://") ||
    src.startsWith("moz-extension://") ||
    msg.includes("metamask") ||
    msg.includes("ses removing unpermitted intrinsics")
  );
}

function debugInfo(scope, details) {
  if (!DEBUG) return;
  console.info(`[BookingDashboard] ${scope}`, details || {});
}

function debugError(scope, error, details) {
  console.error(`[BookingDashboard] ${scope}`, {
    message: error?.message || String(error),
    stack: error?.stack || null,
    ...(details || {})
  });
}
