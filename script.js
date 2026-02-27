const API_URL = "https://script.google.com/macros/s/AKfycbx5RPjxcPVPhiVPLfHGuuskp8hmqTHOLtQpp7lSkBCYGnqPsChHToXoCG7CeRJUW0TR/exec";
const CACHE_KEY = "booking_dashboard_cache_v2";
const DEBUG = true;

const PAGE = document.body.dataset.page || "home";

const state = {
  bookings: [],
  filtered: [],
  sendingRows: new Set(),
  activeReceiptRowIndex: null,
  activeDownpaymentMinimum: 0,
  activeFullPaymentExact: 0,
  sendConfirmResolver: null
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
  receiptSavingIndicator: document.getElementById("receiptSavingIndicator")
};

const PAGE_CONFIG = {
  home: { needsData: false },
  "new-bookings": {
    needsData: true,
    countLabel: "Total Pending",
    columns: ["name", "date", "tour", "guests", "coupon", "discount", "paymentMethod", "totalAmount", "downpayment", "phone", "remarksNew", "actionNew"],
    filterFn: (row) => row.remarksNew.toLowerCase() !== "sent"
  },
  "confirmed-whatsapp": {
    needsData: true,
    countLabel: "Total Pending",
    columns: ["name", "date", "tour", "totalBalance", "phone", "remarksConfirmed", "actionConfirmed"],
    filterFn: (row) => row.remarksConfirmed.toLowerCase() !== "sent"
  },
  "bookings-dashboard": {
    needsData: true,
    countLabel: "Total Rows",
    columns: ["name", "date", "tour", "guests", "paymentMethod", "totalBalance", "phone", "remarksNew", "remarksConfirmed", "adData"],
    filterFn: () => true
  },
  "payment-receipt-encoding": {
    needsData: true,
    countLabel: "Total Rows",
    columns: ["name", "date", "tour", "totalBalance", "paymentMethod", "receiptDown", "receiptFull"],
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

function init() {
  setupGlobalErrorHandlers();
  if (PAGE === "home") return;

  bindEvents();
  hydrateFromCache();
  fetchBookings();
}

function bindEvents() {
  elements.searchInput?.addEventListener("input", applyFilters);
  elements.tourFilter?.addEventListener("change", applyFilters);
  elements.travelDateFilter?.addEventListener("change", applyFilters);
  elements.refreshBtn?.addEventListener("click", fetchBookings);
  elements.sendConfirmOk?.addEventListener("click", () => resolveSendConfirm(true));
  elements.sendConfirmCancel?.addEventListener("click", () => resolveSendConfirm(false));
  elements.sendConfirmModal?.addEventListener("click", (event) => {
    if (event.target === elements.sendConfirmModal) {
      resolveSendConfirm(false);
    }
  });
  elements.receiptCancelBtn?.addEventListener("click", closeReceiptModal);
  elements.receiptSaveBtn?.addEventListener("click", handleSaveReceipt);
  elements.receiptTypeSelect?.addEventListener("change", updateDownpaymentAmountVisibility);
  elements.downpaymentAmountInput?.addEventListener("input", validateDownpaymentAmount);
  elements.fullPaymentAmountInput?.addEventListener("input", validateFullPaymentAmount);
  elements.receiptModal?.addEventListener("click", (event) => {
    if (event.target === elements.receiptModal) {
      closeReceiptModal();
    }
  });
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
    name: String(row.name || "").trim(),
    date: String(row.date || "").trim(),
    tour: String(row.tour || "").trim(),
    phone: String(row.phone || "").trim(),
    numberOfGuest: String(row.numberOfGuest || "").trim(),
    couponCode: String(row.couponCode || "").trim(),
    discountAmount: String(row.discountAmount || "").trim(),
    paymentMethod: String(row.paymentMethod || "").trim(),
    totalAmount: String(row.totalAmount || "").trim(),
    downpaymentAmount: String(row.downpaymentAmount || "").trim(),
    totalBalance: String(row.totalBalance || "").trim(),
    adData: String(row.adData || "").trim(),
    whatsappNew: String(row.whatsappNew || row.whatsapp || "").trim(),
    remarksNew: String(row.remarksNew || row.remarks || "").trim(),
    whatsappConfirmed: String(row.whatsappConfirmed || "").trim(),
    remarksConfirmed: String(row.remarksConfirmed || "").trim(),
    receiptDownpayment: String(row.receiptDownpayment || "").trim(),
    receiptFullPayment: String(row.receiptFullPayment || "").trim()
  };
}

function applyFilters() {
  const cfg = PAGE_CONFIG[PAGE];
  if (!cfg || !cfg.needsData) return;

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

  renderTable(state.filtered, cfg.columns);
  updatePendingCount(scoped.length, cfg.countLabel);
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
      const clickable = PAGE === "payment-receipt-encoding" ? "receipt-click-row" : "";
      const dataAttr = PAGE === "payment-receipt-encoding" ? `data-row-index="${row.rowIndex}"` : "";
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
    case "totalAmount": return escapeHtml(formatMoney(row.totalAmount));
    case "downpayment": return escapeHtml(formatMoney(row.downpaymentAmount));
    case "totalBalance": return escapeHtml(formatMoney(row.totalBalance));
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

  if (PAGE === "payment-receipt-encoding") {
    document.querySelectorAll("tr[data-row-index]").forEach((rowEl) => {
      rowEl.addEventListener("click", () => {
        const rowIndex = Number(rowEl.dataset.rowIndex);
        openReceiptModal(rowIndex);
      });
    });
  }
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
      body: JSON.stringify({ action: mode === "new" ? "mark_sent_new" : "mark_sent_confirmed", rowIndex })
    }, 12000);

    if (!response.ok) {
      throw new Error(`POST failed: ${response.status}`);
    }

    await parseJsonSafe(response);

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

  let receiptImageBase64 = "";
  let receiptMimeType = "";
  let receiptFileName = "";

  setReceiptSaving(true);

  try {
    const file = elements.receiptFileInput?.files?.[0];
    if (file) {
      const encoded = await fileToBase64(file);
      receiptImageBase64 = encoded.base64;
      receiptMimeType = encoded.mimeType;
      receiptFileName = file.name;
    }

    const response = await fetchWithTimeout(API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({
        action: "save_receipt",
        rowIndex,
        paymentType,
        invoiceNumber,
        encodedDownpaymentAmount,
        encodedFullPaymentAmount,
        receiptImageBase64,
        receiptMimeType,
        receiptFileName
      })
    }, 20000);

    if (!response.ok) {
      throw new Error(`POST failed: ${response.status}`);
    }

    const result = await parseJsonSafe(response);

    state.bookings = state.bookings.map((item) => {
      if (item.rowIndex !== rowIndex) return item;
      return {
        ...item,
        downpaymentAmount: paymentType === "downpayment" && encodedDownpaymentAmount
          ? encodedDownpaymentAmount
          : item.downpaymentAmount,
        receiptDownpayment: paymentType === "downpayment" ? invoiceNumber : item.receiptDownpayment,
        receiptFullPayment: paymentType === "full" ? invoiceNumber : item.receiptFullPayment
      };
    });

    saveCache(state.bookings);
    applyFilters();
    closeReceiptModal();
    showToast(result.message || "Receipt encoded successfully.");
  } catch (error) {
    debugError("handleSaveReceipt", error, { rowIndex, paymentType });
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
  }
  if (elements.receiptFullValue) {
    elements.receiptFullValue.textContent = existing.receiptFullPayment || "-";
    elements.receiptFullValue.classList.toggle("invoice-filled", hasFull);
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
