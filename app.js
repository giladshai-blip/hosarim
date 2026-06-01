const STORAGE_KEYS = {
  shipments: "gilad_shipments_v4",
  ignored: "gilad_ignored_v4",
};

const state = {
  inventoryData: [],
  ordersData: [],
  inventoryFileName: "",
  ordersFileName: "",
  lastUpdated: "",
  manualShipments: loadJson(STORAGE_KEYS.shipments, {}),
  ignoredSkus: loadJson(STORAGE_KEYS.ignored, []),
  currentSearch: "",
  filterMode: "shortages",
  numericFilters: {
    ordersMin: "",
    ordersMax: "",
    daysMin: "",
    daysMax: "",
    inventoryMin: "",
    netMax: "",
  },
  archiveOpen: false,
  activeSku: null,
  recentlyUpdatedSku: null,
  preparedItems: [],
};

const columnAliases = {
  sku: ["מק'ט", 'מק"ט', "מקט", "SKU", "sku", "Item", "Item Code", "קוד פריט"],
  inventoryQty: ["מלאי מרלוג ביח'", "מלאי מרלוג ביחידות", "מלאי", "Quantity", "Qty", "כמות במלאי"],
  inventoryDays: ["ימי מלאי", "ימי מלאי 0", "ימי מלאי0", "Days Inventory", "Inventory Days", "Days of Inventory", "DOH", "DOS"],
  orderQty: ["כמות בהזמנות", "הזמנות", "כמות", "Quantity", "Qty", "כמות פתוחה"],
  name: ["תאור מוצר", "תיאור מוצר", "שם מוצר", "Description", "Item Name", "Product Name"],
};

const LOW_INVENTORY_DAYS = 7;

const elements = {
  inputBoth: document.getElementById("inputBoth"),
  inputInventory: document.getElementById("inputInventory"),
  inputOrders: document.getElementById("inputOrders"),
  dropBoth: document.getElementById("dropBoth"),
  dropInventory: document.getElementById("dropInventory"),
  dropOrders: document.getElementById("dropOrders"),
  nameBoth: document.getElementById("nameBoth"),
  nameInventory: document.getElementById("nameInventory"),
  nameOrders: document.getElementById("nameOrders"),
  fileStatus: document.getElementById("fileStatus"),
  inventoryFileLabel: document.getElementById("inventoryFileLabel"),
  ordersFileLabel: document.getElementById("ordersFileLabel"),
  lastUpdatedLabel: document.getElementById("lastUpdatedLabel"),
  swapReportsBtn: document.getElementById("swapReportsBtn"),
  statusBadge: document.getElementById("statusBadge"),
  emptyState: document.getElementById("emptyState"),
  appBody: document.getElementById("appBody"),
  searchInput: document.getElementById("searchInput"),
  searchMeta: document.getElementById("searchMeta"),
  clearSearchBtn: document.getElementById("clearSearchBtn"),
  filterMode: document.getElementById("filterMode"),
  ordersMinFilter: document.getElementById("ordersMinFilter"),
  ordersMaxFilter: document.getElementById("ordersMaxFilter"),
  daysMinFilter: document.getElementById("daysMinFilter"),
  daysMaxFilter: document.getElementById("daysMaxFilter"),
  inventoryMinFilter: document.getElementById("inventoryMinFilter"),
  netMaxFilter: document.getElementById("netMaxFilter"),
  clearFiltersBtn: document.getElementById("clearFiltersBtn"),
  resetFilesBtn: document.getElementById("resetFilesBtn"),
  resetManualBtn: document.getElementById("resetManualBtn"),
  exportBtn: document.getElementById("exportBtn"),
  archiveBtn: document.getElementById("archiveBtn"),
  archiveCount: document.getElementById("archiveCount"),
  hiddenPanel: document.getElementById("hiddenPanel"),
  hiddenCount: document.getElementById("hiddenCount"),
  hiddenList: document.getElementById("hiddenList"),
  tableBody: document.getElementById("shortageTableBody"),
  shortageCount: document.getElementById("shortageCount"),
  shippedCount: document.getElementById("shippedCount"),
  visibleCount: document.getElementById("visibleCount"),
  messageArea: document.getElementById("messageArea"),
  dialog: document.getElementById("shipmentDialog"),
  modalDesc: document.getElementById("modalDesc"),
  modalInput: document.getElementById("modalInput"),
  modalSave: document.getElementById("modalSave"),
  modalClose: document.getElementById("modalClose"),
};

document.documentElement.dataset.xlsxNative = "DecompressionStream" in window ? "yes" : "no";

setupBulkDropZone();
setupDropZone(elements.dropInventory, elements.inputInventory, "inventory");
setupDropZone(elements.dropOrders, elements.inputOrders, "orders");

elements.searchInput.addEventListener("input", (event) => {
  state.currentSearch = event.target.value.trim().toLowerCase();
  processAndRender();
});

elements.clearSearchBtn.addEventListener("click", () => {
  elements.searchInput.value = "";
  state.currentSearch = "";
  elements.searchInput.focus();
  processAndRender();
});

elements.filterMode.addEventListener("change", (event) => {
  state.filterMode = event.target.value;
  processAndRender();
});

setupNumericFilters();

elements.clearFiltersBtn.addEventListener("click", () => {
  clearNumericFilters();
  processAndRender();
});

elements.swapReportsBtn.addEventListener("click", swapReports);

elements.resetFilesBtn.addEventListener("click", () => {
  const confirmed = window.confirm("לאפס את שני הקבצים שהועלו?");
  if (!confirmed) return;
  resetFiles();
});

elements.resetManualBtn.addEventListener("click", () => {
  const confirmed = window.confirm("לאפס את כל דיווחי המשלוחים מהמפעל?");
  if (!confirmed) return;

  state.manualShipments = {};
  state.recentlyUpdatedSku = null;
  saveLocalState();
  processAndRender();
  showMessage("דיווחי המשלוחים אופסו.");
});

elements.exportBtn.addEventListener("click", exportVisibleRows);

elements.archiveBtn.addEventListener("click", () => {
  state.archiveOpen = !state.archiveOpen;
  renderHiddenItems();
});

elements.modalSave.addEventListener("click", () => {
  saveShipment();
  elements.dialog.close("save");
});

elements.modalClose.addEventListener("click", () => {
  elements.dialog.close("cancel");
});

elements.modalInput.addEventListener("keydown", (event) => {
  if (event.key !== "Enter") return;
  event.preventDefault();
  saveShipment();
  elements.dialog.close("save");
});

function setupNumericFilters() {
  const filterInputs = [
    [elements.ordersMinFilter, "ordersMin"],
    [elements.ordersMaxFilter, "ordersMax"],
    [elements.daysMinFilter, "daysMin"],
    [elements.daysMaxFilter, "daysMax"],
    [elements.inventoryMinFilter, "inventoryMin"],
    [elements.netMaxFilter, "netMax"],
  ];

  filterInputs.forEach(([input, key]) => {
    input.addEventListener("input", (event) => {
      state.numericFilters[key] = event.target.value;
      processAndRender();
    });
  });
}

function setupDropZone(zone, input, type) {
  const button = zone.querySelector(".drop-card");

  button.addEventListener("click", () => input.click());
  input.addEventListener("change", () => {
    if (input.files.length > 0) handleFile(input.files[0], type);
  });

  zone.addEventListener("dragover", (event) => {
    event.preventDefault();
    zone.classList.add("dragover");
  });

  zone.addEventListener("dragleave", () => {
    zone.classList.remove("dragover");
  });

  zone.addEventListener("drop", (event) => {
    event.preventDefault();
    zone.classList.remove("dragover");
    const file = event.dataTransfer.files[0];
    if (file) handleFile(file, type);
  });
}

function setupBulkDropZone() {
  const button = elements.dropBoth.querySelector(".drop-card");

  button.addEventListener("click", () => elements.inputBoth.click());
  elements.inputBoth.addEventListener("change", () => {
    if (elements.inputBoth.files.length > 0) handleBulkFiles(Array.from(elements.inputBoth.files));
  });

  elements.dropBoth.addEventListener("dragover", (event) => {
    event.preventDefault();
    elements.dropBoth.classList.add("dragover");
  });

  elements.dropBoth.addEventListener("dragleave", () => {
    elements.dropBoth.classList.remove("dragover");
  });

  elements.dropBoth.addEventListener("drop", (event) => {
    event.preventDefault();
    elements.dropBoth.classList.remove("dragover");
    handleBulkFiles(Array.from(event.dataTransfer.files));
  });
}

async function handleBulkFiles(files) {
  const acceptedFiles = files.filter(isSupportedFile);
  if (acceptedFiles.length < 2) {
    showMessage("צריך לבחור שני קבצים: מלאי והזמנות.");
    return;
  }

  try {
    const reports = await Promise.all(acceptedFiles.map(async (file) => {
      const rows = await readFileRows(file);
      return {
        file,
        rows,
        type: detectReportType(rows, file.name),
      };
    }));

    assignMissingReportTypes(reports);

    const inventoryReport = reports.find((report) => report.type === "inventory");
    const ordersReport = reports.find((report) => report.type === "orders");

    if (!inventoryReport || !ordersReport) {
      showMessage("לא הצלחתי לזהות אוטומטית איזה קובץ הוא מלאי ואיזה קובץ הוא הזמנות. אפשר להעלות אותם באזורים הנפרדים.");
      return;
    }

    validateRows(inventoryReport.rows, "inventory", inventoryReport.file.name);
    validateRows(ordersReport.rows, "orders", ordersReport.file.name);
    setReport("inventory", inventoryReport.rows, inventoryReport.file.name);
    setReport("orders", ordersReport.rows, ordersReport.file.name);
    state.lastUpdated = formatDateTime(new Date());

    markFileLoaded(elements.dropInventory, elements.nameInventory, inventoryReport.file.name);
    markFileLoaded(elements.dropOrders, elements.nameOrders, ordersReport.file.name);
    markFileLoaded(elements.dropBoth, elements.nameBoth, `${inventoryReport.file.name} + ${ordersReport.file.name}`);
    updateFileStatus();
    checkAndProcess();
  } catch (error) {
    showMessage(`לא הצלחתי לקרוא את הקבצים: ${error.message}`);
  }
}

async function handleFile(file, type) {
  try {
    const rows = await readFileRows(file);
    validateRows(rows, type, file.name);
    setReport(type, rows, file.name);

    if (type === "inventory") {
      markFileLoaded(elements.dropInventory, elements.nameInventory, file.name);
    } else {
      markFileLoaded(elements.dropOrders, elements.nameOrders, file.name);
    }
    updateFileStatus();
    checkAndProcess();
  } catch (error) {
    showMessage(`לא הצלחתי לקרוא את הקובץ: ${error.message}`);
  }
}

function setReport(type, rows, fileName) {
  state.lastUpdated = formatDateTime(new Date());

  if (type === "inventory") {
    state.inventoryData = rows;
    state.inventoryFileName = fileName;
    return;
  }

  state.ordersData = rows;
  state.ordersFileName = fileName;
}

function readFileRows(file) {
  const lowerName = file.name.toLowerCase();
  const isCsv = lowerName.endsWith(".csv");
  const isXlsx = lowerName.endsWith(".xlsx");

  if (!isCsv && !isXlsx) {
    showMessage("אפשר להעלות קובצי XLSX או CSV בלבד.");
    return Promise.reject(new Error("אפשר להעלות קובצי XLSX או CSV בלבד."));
  }

  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = async () => {
      try {
        const rows = isXlsx
          ? await parseXlsx(reader.result)
          : parseCsv(String(reader.result || ""));
        resolve(rows);
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = () => reject(new Error("לא הצלחתי לקרוא את הקובץ. כדאי לנסות לשמור אותו שוב."));

    if (isXlsx) {
      reader.readAsArrayBuffer(file);
    } else {
      reader.readAsText(file, "UTF-8");
    }
  });
}

function isSupportedFile(file) {
  const lowerName = file.name.toLowerCase();
  return lowerName.endsWith(".csv") || lowerName.endsWith(".xlsx");
}

function detectReportType(rows, fileName) {
  const headers = new Set(Object.keys(rows[0] || {}).map(normalizeHeader));
  const inventoryScore = scoreHeaders(headers, columnAliases.inventoryQty);
  const ordersScore = scoreHeaders(headers, columnAliases.orderQty);

  if (inventoryScore > ordersScore) return "inventory";
  if (ordersScore > inventoryScore) return "orders";

  const lowerName = fileName.toLowerCase();
  if (/wms|inventory|stock|מלאי|מחסן|מרלוג/.test(lowerName)) return "inventory";
  if (/order|orders|sales|הזמנות|לקוחות|צבר/.test(lowerName)) return "orders";
  return "unknown";
}

function scoreHeaders(headers, aliases) {
  return aliases.reduce((score, alias) => {
    const normalized = normalizeHeader(alias);
    if (!headers.has(normalized)) return score;
    return score + (["Quantity", "Qty", "כמות"].includes(alias) ? 1 : 4);
  }, 0);
}

function assignMissingReportTypes(reports) {
  const inventoryReports = reports.filter((report) => report.type === "inventory");
  const ordersReports = reports.filter((report) => report.type === "orders");
  const unknownReports = reports.filter((report) => report.type === "unknown");

  if (reports.length !== 2 || unknownReports.length !== 1) return;
  if (inventoryReports.length === 1 && ordersReports.length === 0) unknownReports[0].type = "orders";
  if (ordersReports.length === 1 && inventoryReports.length === 0) unknownReports[0].type = "inventory";
}

function validateRows(rows, type, fileName) {
  if (!rows.length) {
    throw new Error(`${fileName}: הקובץ ריק או שאין בו שורות נתונים.`);
  }

  const headers = new Set(Object.keys(rows[0] || {}).map(normalizeHeader));
  const missing = [];

  if (!hasAnyHeader(headers, columnAliases.sku)) missing.push("מק\"ט");
  if (type === "inventory" && !hasAnyHeader(headers, columnAliases.inventoryQty)) missing.push("מלאי");
  if (type === "orders" && !hasAnyHeader(headers, columnAliases.orderQty)) missing.push("כמות בהזמנות");

  if (missing.length) {
    throw new Error(`${fileName}: חסרות עמודות חובה - ${missing.join(", ")}.`);
  }
}

function hasAnyHeader(headers, aliases) {
  return aliases.some((alias) => headers.has(normalizeHeader(alias)));
}

function updateFileStatus() {
  const hasInventory = Boolean(state.inventoryFileName);
  const hasOrders = Boolean(state.ordersFileName);

  elements.fileStatus.hidden = !hasInventory && !hasOrders;
  elements.inventoryFileLabel.textContent = state.inventoryFileName || "לא נטען";
  elements.ordersFileLabel.textContent = state.ordersFileName || "לא נטען";
  elements.lastUpdatedLabel.textContent = state.lastUpdated || "לא עודכן";
  elements.swapReportsBtn.disabled = !(hasInventory && hasOrders);
}

function swapReports() {
  const inventoryData = state.inventoryData;
  const inventoryFileName = state.inventoryFileName;

  state.inventoryData = state.ordersData;
  state.inventoryFileName = state.ordersFileName;
  state.ordersData = inventoryData;
  state.ordersFileName = inventoryFileName;

  markFileLoaded(elements.dropInventory, elements.nameInventory, state.inventoryFileName);
  markFileLoaded(elements.dropOrders, elements.nameOrders, state.ordersFileName);
  markFileLoaded(elements.dropBoth, elements.nameBoth, `${state.inventoryFileName} + ${state.ordersFileName}`);
  updateFileStatus();
  checkAndProcess();
  showMessage("החלפתי בין קובץ המלאי וקובץ ההזמנות.");
}

function resetFiles() {
  state.inventoryData = [];
  state.ordersData = [];
  state.inventoryFileName = "";
  state.ordersFileName = "";
  state.lastUpdated = "";
  state.preparedItems = [];
  state.recentlyUpdatedSku = null;

  clearLoadedFile(elements.dropInventory, elements.nameInventory);
  clearLoadedFile(elements.dropOrders, elements.nameOrders);
  clearLoadedFile(elements.dropBoth, elements.nameBoth);

  elements.inputInventory.value = "";
  elements.inputOrders.value = "";
  elements.inputBoth.value = "";
  elements.tableBody.replaceChildren();
  elements.searchInput.value = "";
  state.currentSearch = "";
  clearNumericFilters();
  updateSearchMeta(0);
  elements.shortageCount.textContent = "0";
  elements.shippedCount.textContent = "0";
  elements.visibleCount.textContent = "0";
  elements.appBody.hidden = true;
  elements.emptyState.hidden = false;
  updateFileStatus();
  setStatus("waiting", "ממתין להעלאת קבצים");
  showMessage("הקבצים אופסו. דיווחי המשלוחים נשמרו.");
}

function clearLoadedFile(zone, nameElement) {
  zone.classList.remove("loaded");
  nameElement.textContent = "";
  nameElement.hidden = true;
}

function markFileLoaded(zone, nameElement, fileName) {
  zone.classList.add("loaded");
  nameElement.textContent = fileName;
  nameElement.hidden = false;
}

function checkAndProcess() {
  const hasInventory = state.inventoryData.length > 0;
  const hasOrders = state.ordersData.length > 0;

  if (hasInventory && hasOrders) {
    elements.emptyState.hidden = true;
    elements.appBody.hidden = false;
    setStatus("ready", "הנתונים סונכרנו בהצלחה");
    processAndRender();
    return;
  }

  if (hasInventory || hasOrders) {
    setStatus("partial", "ממתין לקובץ השני");
  }
}

function processAndRender() {
  const ordersBySku = groupOrdersBySku(state.ordersData);
  const inventoryBySku = mapInventoryBySku(state.inventoryData);
  const allSkus = new Set([...Object.keys(ordersBySku), ...Object.keys(inventoryBySku)]);

  let items = Array.from(allSkus)
    .map((sku) => {
      const inventory = inventoryBySku[sku] || { qty: 0, days: null, name: "לא מופיע במלאי (רק בהזמנות)" };
      const orders = ordersBySku[sku] || 0;
      const manual = Number(state.manualShipments[sku] || 0);
      const expectedInventory = inventory.qty + manual;
      const balance = expectedInventory - orders;

      return {
        sku,
        name: inventory.name,
        inventory: inventory.qty,
        inventoryDays: inventory.days,
        orders,
        manual,
        expectedInventory,
        balance,
        status: getItemStatus(balance, manual, inventory.days),
      };
    })
    .filter((item) => item.sku && !state.ignoredSkus.includes(item.sku));

  if (state.currentSearch) {
    state.recentlyUpdatedSku = null;
    items = items.filter((item) => {
      return item.sku.toLowerCase().includes(state.currentSearch) ||
        item.name.toLowerCase().includes(state.currentSearch);
    });
  }

  if (state.filterMode === "shortages") {
    items = items.filter((item) => {
      return item.balance < 0 || item.manual > 0 || item.status.key === "risk";
    });
  }

  if (state.filterMode === "shipped") {
    items = items.filter((item) => item.manual > 0);
  }

  items = applyNumericFilters(items);

  items.sort((a, b) => {
    if (a.balance !== b.balance) return a.balance - b.balance;
    return a.name.localeCompare(b.name, "he");
  });

  state.preparedItems = items;
  renderRows(items);
  updateKpis(items);
  updateSearchMeta(items.length);
  renderHiddenItems();
}

function getItemStatus(balance, manual, inventoryDays) {
  if (balance < 0) return { key: "shortage", label: "חסר" };
  if (manual > 0) return { key: "shipped", label: "נסגר במשלוח" };
  if (isLowInventoryDays(inventoryDays)) return { key: "risk", label: "בסיכון" };
  if (balance > 0) return { key: "surplus", label: "בעודף" };
  return { key: "balanced", label: "מאוזן" };
}

function isLowInventoryDays(inventoryDays) {
  return Number.isFinite(inventoryDays) && inventoryDays >= 0 && inventoryDays <= LOW_INVENTORY_DAYS;
}

function groupOrdersBySku(rows) {
  return rows.reduce((accumulator, row) => {
    const sku = getValue(row, columnAliases.sku);
    const qty = parseNumber(getValue(row, columnAliases.orderQty));

    if (sku && Number.isFinite(qty)) {
      accumulator[sku] = (accumulator[sku] || 0) + qty;
    }

    return accumulator;
  }, {});
}

function mapInventoryBySku(rows) {
  return rows.reduce((accumulator, row) => {
    const sku = getValue(row, columnAliases.sku);
    const qty = parseNumber(getValue(row, columnAliases.inventoryQty));
    const days = parseOptionalNumber(getValue(row, columnAliases.inventoryDays));
    const name = getValue(row, columnAliases.name) || "מוצר לא מזוהה";

    if (sku && Number.isFinite(qty)) {
      accumulator[sku] = {
        qty: (accumulator[sku]?.qty || 0) + qty,
        days: Number.isFinite(days) ? days : accumulator[sku]?.days ?? null,
        name,
      };
    }

    return accumulator;
  }, {});
}

function renderRows(items) {
  elements.tableBody.replaceChildren();

  if (items.length === 0) {
    const row = document.createElement("tr");
    const cell = document.createElement("td");
    cell.colSpan = document.querySelectorAll("thead th").length;
    cell.className = "number-cell";
    cell.textContent = "אין שורות להצגה לפי הסינון הנוכחי";
    row.append(cell);
    elements.tableBody.append(row);
    return;
  }

  const fragment = document.createDocumentFragment();
  items.forEach((item) => fragment.append(createTableRow(item)));
  elements.tableBody.append(fragment);
}

function createTableRow(item) {
  const row = document.createElement("tr");
  if (item.manual > 0 && item.balance >= 0) {
    row.classList.add("manually-shipped");
  }
  if (item.sku === state.recentlyUpdatedSku) {
    row.classList.add("recently-updated");
  }

  row.append(
    createCell(item.sku, "sticky-col sku-cell"),
    createNameCell(item),
    createCell(formatNumber(item.inventory), "number-cell"),
    createInventoryDaysCell(item.inventoryDays),
    createCell(formatNumber(item.orders), "number-cell"),
    createManualCell(item),
    createExpectedInventoryCell(item),
    createBalanceCell(item.balance),
    createStatusCell(item.status),
    createActionsCell(item)
  );

  return row;
}

function createCell(text, className) {
  const cell = document.createElement("td");
  cell.className = className;
  cell.textContent = text;
  return cell;
}

function createNameCell(item) {
  const cell = document.createElement("td");
  cell.className = "sticky-col name-cell";

  const wrapper = document.createElement("div");
  wrapper.className = "name-with-action";

  const text = document.createElement("span");
  text.className = "name-text";
  text.title = item.name;
  text.textContent = item.name;

  const hideButton = document.createElement("button");
  hideButton.type = "button";
  hideButton.className = "inline-hide-button";
  hideButton.title = "הסתר רשומה";
  hideButton.setAttribute("aria-label", `הסתר את ${item.sku}`);
  hideButton.innerHTML = iconMarkup("hide");
  hideButton.addEventListener("click", () => ignoreSku(item.sku));

  wrapper.append(text, hideButton);
  cell.append(wrapper);
  return cell;
}

function createInventoryDaysCell(days) {
  const cell = document.createElement("td");
  cell.className = "number-cell";

  if (!Number.isFinite(days)) {
    cell.textContent = "-";
    cell.classList.add("muted-dash");
    return cell;
  }

  const badge = document.createElement("span");
  badge.className = `days-badge ${isLowInventoryDays(days) ? "low" : "ok"}`;
  badge.textContent = formatNumber(days);
  cell.append(badge);

  return cell;
}

function createManualCell(item) {
  const cell = document.createElement("td");
  cell.className = "number-cell";

  const input = document.createElement("input");
  input.className = "manual-input";
  input.type = "number";
  input.min = "0";
  input.step = "1";
  input.inputMode = "numeric";
  input.value = item.manual > 0 ? String(item.manual) : "";
  input.placeholder = "0";
  input.setAttribute("aria-label", `נשלח היום עבור ${item.sku}`);
  input.addEventListener("change", () => {
    updateManualShipment(item.sku, input.value, "המשלוח עודכן והיתרה חושבה מחדש.");
  });
  input.addEventListener("keydown", (event) => {
    if (event.key !== "Enter") return;
    event.preventDefault();
    input.blur();
  });

  cell.append(input);

  return cell;
}

function createExpectedInventoryCell(item) {
  const cell = document.createElement("td");
  cell.className = "number-cell";

  const badge = document.createElement("span");
  badge.className = `expected-badge ${item.manual > 0 ? "updated" : "current"}`;
  badge.textContent = formatNumber(item.expectedInventory);
  badge.title = `מלאי WMS ${formatNumber(item.inventory)} + נשלח היום ${formatNumber(item.manual)}`;
  cell.append(badge);
  return cell;
}

function createBalanceCell(balance) {
  const cell = document.createElement("td");
  cell.className = "number-cell net-cell";

  const badge = document.createElement("span");
  badge.className = `balance-badge ${balance < 0 ? "negative" : "positive"}`;
  badge.textContent = `${balance > 0 ? "+" : ""}${formatNumber(balance)}`;
  cell.append(badge);

  return cell;
}

function createStatusCell(status) {
  const cell = document.createElement("td");
  cell.className = "number-cell";

  const badge = document.createElement("span");
  badge.className = `status-chip ${status.key}`;
  badge.textContent = status.label;
  cell.append(badge);

  return cell;
}

function createActionsCell(item) {
  const cell = document.createElement("td");
  const actions = document.createElement("div");
  actions.className = "row-actions";

  const shipmentButton = document.createElement("button");
  shipmentButton.type = "button";
  shipmentButton.className = "icon-button ship";
  shipmentButton.title = "דווח על משלוח";
  shipmentButton.setAttribute("aria-label", `דווח על משלוח עבור ${item.sku}`);
  shipmentButton.innerHTML = iconMarkup("shipment");
  shipmentButton.addEventListener("click", () => openShipmentDialog(item));

  actions.append(shipmentButton);
  cell.append(actions);
  return cell;
}

function openShipmentDialog(item) {
  state.activeSku = item.sku;
  elements.modalDesc.textContent = `${item.name} · ${item.sku}`;
  elements.modalInput.value = state.manualShipments[item.sku] || "";
  elements.dialog.showModal();
  elements.modalInput.focus();
}

function saveShipment() {
  if (!state.activeSku) return;
  updateManualShipment(state.activeSku, elements.modalInput.value, "המשלוח עודכן והיתרה חושבה מחדש.");
}

function updateManualShipment(sku, rawValue, message) {
  const value = Math.max(0, parseInt(rawValue || "0", 10) || 0);

  if (value > 0) {
    state.manualShipments[sku] = value;
  } else {
    delete state.manualShipments[sku];
  }

  state.recentlyUpdatedSku = sku;
  saveLocalState();
  processAndRender();
  if (message) showMessage(message);
}

function ignoreSku(sku) {
  if (!state.ignoredSkus.includes(sku)) {
    state.ignoredSkus.push(sku);
    state.archiveOpen = false;
    saveLocalState();
    processAndRenderWithoutJump();
    showMessage("הפריט עבר לארכיון. אפשר להחזיר אותו דרך כפתור פריטים בארכיון.");
  }
}

function processAndRenderWithoutJump() {
  const scrollX = window.scrollX;
  const scrollY = window.scrollY;
  processAndRender();
  requestAnimationFrame(() => window.scrollTo(scrollX, scrollY));
}

function restoreSku(sku) {
  state.ignoredSkus = state.ignoredSkus.filter((ignoredSku) => ignoredSku !== sku);
  saveLocalState();
  processAndRender();
  showMessage("הפריט הוחזר לטבלה.");
}

function clearHiddenSkus() {
  state.ignoredSkus = [];
  saveLocalState();
  processAndRender();
  showMessage("כל הפריטים המוסתרים הוחזרו.");
}

function applyNumericFilters(items) {
  return items.filter((item) => {
    return isInRange(item.orders, state.numericFilters.ordersMin, state.numericFilters.ordersMax) &&
      isInRange(item.inventoryDays, state.numericFilters.daysMin, state.numericFilters.daysMax) &&
      isInRange(item.inventory, state.numericFilters.inventoryMin, "") &&
      isInRange(item.balance, "", state.numericFilters.netMax);
  });
}

function isInRange(value, minValue, maxValue, allowMissing = false) {
  if (!hasNumberFilter(minValue) && !hasNumberFilter(maxValue)) return true;
  if (!Number.isFinite(value)) return allowMissing;

  const min = hasNumberFilter(minValue) ? Number(minValue) : -Infinity;
  const max = hasNumberFilter(maxValue) ? Number(maxValue) : Infinity;

  return value >= min && value <= max;
}

function hasNumberFilter(value) {
  return value !== "" && Number.isFinite(Number(value));
}

function clearNumericFilters() {
  Object.keys(state.numericFilters).forEach((key) => {
    state.numericFilters[key] = "";
  });

  elements.ordersMinFilter.value = "";
  elements.ordersMaxFilter.value = "";
  elements.daysMinFilter.value = "";
  elements.daysMaxFilter.value = "";
  elements.inventoryMinFilter.value = "";
  elements.netMaxFilter.value = "";
  elements.clearFiltersBtn.disabled = true;
}

function hasActiveNumericFilters() {
  return Object.values(state.numericFilters).some(hasNumberFilter);
}

function updateSearchMeta(visibleCount) {
  const hasSearch = state.currentSearch.length > 0;
  const hasFilters = hasActiveNumericFilters();
  elements.clearSearchBtn.disabled = !hasSearch;
  elements.clearFiltersBtn.disabled = !hasFilters;

  if (!hasSearch && !hasFilters) {
    elements.searchMeta.textContent = "הקלד כדי לסנן את הטבלה";
    return;
  }

  if (hasSearch) {
    const filterSuffix = hasFilters ? " ולפי הסינונים" : "";
    elements.searchMeta.textContent = `נמצאו ${formatNumber(visibleCount)} פריטים עבור "${elements.searchInput.value.trim()}"${filterSuffix}`;
    return;
  }

  elements.searchMeta.textContent = `נמצאו ${formatNumber(visibleCount)} פריטים לפי הסינונים`;
}

function updateKpis(items) {
  const totalMissingUnits = items.reduce((sum, item) => {
    return item.balance < 0 ? sum + Math.abs(item.balance) : sum;
  }, 0);

  const totalShippedUnits = items.reduce((sum, item) => sum + item.manual, 0);

  elements.shortageCount.textContent = formatNumber(totalMissingUnits);
  elements.shippedCount.textContent = formatNumber(totalShippedUnits);
  elements.visibleCount.textContent = formatNumber(items.length);
}

function renderHiddenItems() {
  elements.hiddenCount.textContent = formatNumber(state.ignoredSkus.length);
  elements.archiveCount.textContent = formatNumber(state.ignoredSkus.length);
  elements.archiveBtn.disabled = state.ignoredSkus.length === 0;
  elements.archiveBtn.classList.toggle("active", state.archiveOpen);
  elements.hiddenList.replaceChildren();
  elements.hiddenPanel.hidden = state.ignoredSkus.length === 0 || !state.archiveOpen;

  if (state.ignoredSkus.length === 0 || !state.archiveOpen) return;

  const inventoryBySku = mapInventoryBySku(state.inventoryData);
  const fragment = document.createDocumentFragment();

  state.ignoredSkus.forEach((sku) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "hidden-item";
    button.textContent = `${sku} · ${inventoryBySku[sku]?.name || "פריט מוסתר"}`;
    button.title = "לחץ להחזרה לטבלה";
    button.addEventListener("click", () => restoreSku(sku));
    fragment.append(button);
  });

  const restoreAll = document.createElement("button");
  restoreAll.type = "button";
  restoreAll.className = "hidden-item restore-all";
  restoreAll.textContent = "החזר הכל";
  restoreAll.addEventListener("click", clearHiddenSkus);
  fragment.append(restoreAll);

  elements.hiddenList.append(fragment);
}

function exportVisibleRows() {
  if (state.preparedItems.length === 0) {
    showMessage("אין שורות לייצוא כרגע.");
    return;
  }

  const headers = ["מק\"ט", "תיאור מוצר", "מלאי WMS", "ימי מלאי", "הזמנות", "נשלח היום", "מלאי צפוי לאחר משלוח", "חוסר נטו", "סטטוס"];
  const lines = [
    headers,
    ...state.preparedItems.map((item) => [
      item.sku,
      item.name,
      item.inventory,
      Number.isFinite(item.inventoryDays) ? item.inventoryDays : "",
      item.orders,
      item.manual,
      item.expectedInventory,
      item.balance,
      item.status.label,
    ]),
  ].map((row) => row.map(escapeCsvCell).join(","));

  const blob = new Blob(["\uFEFF" + lines.join("\n")], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "shortages-net.csv";
  link.click();
  URL.revokeObjectURL(url);
}

function parseCsv(text) {
  const cleanText = text.replace(/^\uFEFF/, "");
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;

  for (let index = 0; index < cleanText.length; index += 1) {
    const char = cleanText[index];
    const next = cleanText[index + 1];

    if (char === '"' && inQuotes && next === '"') {
      cell += '"';
      index += 1;
      continue;
    }

    if (char === '"') {
      inQuotes = !inQuotes;
      continue;
    }

    if (char === "," && !inQuotes) {
      row.push(cell);
      cell = "";
      continue;
    }

    if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") index += 1;
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
      continue;
    }

    cell += char;
  }

  if (cell || row.length > 0) {
    row.push(cell);
    rows.push(row);
  }

  const [headersRow, ...dataRows] = rows.filter((currentRow) => {
    return currentRow.some((value) => value.trim() !== "");
  });

  if (!headersRow) return [];

  const headers = headersRow.map((header) => normalizeHeader(header));
  return dataRows.map((dataRow) => {
    return headers.reduce((record, header, index) => {
      if (header) record[header] = (dataRow[index] || "").trim();
      return record;
    }, {});
  });
}

async function parseXlsx(arrayBuffer) {
  if (!("DecompressionStream" in window)) {
    throw new Error("הדפדפן לא תומך בקריאת XLSX מקומית. אפשר להשתמש ב-CSV כחלופה.");
  }

  const archive = await readZipArchive(arrayBuffer);
  const workbookXml = await archive.text("xl/workbook.xml");
  const workbookRelsXml = await archive.text("xl/_rels/workbook.xml.rels");
  const sheetPath = getFirstSheetPath(workbookXml, workbookRelsXml);
  const sharedStringsXml = archive.has("xl/sharedStrings.xml")
    ? await archive.text("xl/sharedStrings.xml")
    : "";
  const sharedStrings = parseSharedStrings(sharedStringsXml);
  const sheetXml = await archive.text(sheetPath);
  const rows = parseWorksheet(sheetXml, sharedStrings);

  if (rows.length === 0) return [];

  const [headersRow, ...dataRows] = rows;
  const headers = headersRow.map((header) => normalizeHeader(header));

  return dataRows
    .filter((row) => row.some((value) => String(value || "").trim() !== ""))
    .map((row) => {
      return headers.reduce((record, header, index) => {
        if (header) record[header] = String(row[index] ?? "").trim();
        return record;
      }, {});
    });
}

async function readZipArchive(arrayBuffer) {
  const bytes = new Uint8Array(arrayBuffer);
  const entries = parseZipEntries(bytes);
  const fileMap = new Map(entries.map((entry) => [entry.name, entry]));

  return {
    has(path) {
      return fileMap.has(path);
    },
    async text(path) {
      const entry = fileMap.get(path);
      if (!entry) throw new Error(`חסר קובץ פנימי באקסל: ${path}`);
      const data = await inflateZipEntry(bytes, entry);
      return new TextDecoder("utf-8").decode(data);
    },
  };
}

function parseZipEntries(bytes) {
  const eocdOffset = findEndOfCentralDirectory(bytes);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const centralDirectoryOffset = view.getUint32(eocdOffset + 16, true);
  const totalEntries = view.getUint16(eocdOffset + 10, true);
  const decoder = new TextDecoder("utf-8");
  const entries = [];
  let offset = centralDirectoryOffset;

  for (let index = 0; index < totalEntries; index += 1) {
    if (view.getUint32(offset, true) !== 0x02014b50) {
      throw new Error("מבנה קובץ XLSX לא תקין.");
    }

    const compression = view.getUint16(offset + 10, true);
    const compressedSize = view.getUint32(offset + 20, true);
    const uncompressedSize = view.getUint32(offset + 24, true);
    const fileNameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const localHeaderOffset = view.getUint32(offset + 42, true);
    const fileNameStart = offset + 46;
    const fileNameEnd = fileNameStart + fileNameLength;
    const name = decoder.decode(bytes.slice(fileNameStart, fileNameEnd));

    entries.push({
      name,
      compression,
      compressedSize,
      uncompressedSize,
      localHeaderOffset,
    });

    offset = fileNameEnd + extraLength + commentLength;
  }

  return entries;
}

function findEndOfCentralDirectory(bytes) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const minOffset = Math.max(0, bytes.length - 66000);

  for (let offset = bytes.length - 22; offset >= minOffset; offset -= 1) {
    if (view.getUint32(offset, true) === 0x06054b50) return offset;
  }

  throw new Error("זה לא נראה כמו קובץ XLSX תקין.");
}

async function inflateZipEntry(bytes, entry) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const localOffset = entry.localHeaderOffset;

  if (view.getUint32(localOffset, true) !== 0x04034b50) {
    throw new Error("מבנה קובץ XLSX לא תקין.");
  }

  const fileNameLength = view.getUint16(localOffset + 26, true);
  const extraLength = view.getUint16(localOffset + 28, true);
  const dataStart = localOffset + 30 + fileNameLength + extraLength;
  const compressed = bytes.slice(dataStart, dataStart + entry.compressedSize);

  if (entry.compression === 0) return compressed;
  if (entry.compression !== 8) {
    throw new Error("סוג דחיסה לא נתמך בקובץ XLSX.");
  }

  const stream = new Blob([compressed]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  const buffer = await new Response(stream).arrayBuffer();
  const inflated = new Uint8Array(buffer);

  if (entry.uncompressedSize && inflated.length !== entry.uncompressedSize) {
    throw new Error("לא הצלחתי לפתוח את אחד מחלקי קובץ האקסל.");
  }

  return inflated;
}

function getFirstSheetPath(workbookXml, workbookRelsXml) {
  const workbook = parseXml(workbookXml);
  const rels = parseXml(workbookRelsXml);
  const firstSheet = workbook.querySelector("sheet");
  if (!firstSheet) throw new Error("לא נמצא גיליון בקובץ האקסל.");

  const relationshipId = firstSheet.getAttribute("r:id");
  const relationship = Array.from(rels.querySelectorAll("Relationship")).find((rel) => {
    return rel.getAttribute("Id") === relationshipId;
  });

  if (!relationship) throw new Error("לא נמצא קישור לגיליון הראשון בקובץ.");

  const target = relationship.getAttribute("Target") || "worksheets/sheet1.xml";
  if (target.startsWith("/")) return target.slice(1);
  return normalizeZipPath(`xl/${target}`);
}

function parseSharedStrings(xmlText) {
  if (!xmlText) return [];

  const xml = parseXml(xmlText);
  return Array.from(xml.querySelectorAll("si")).map((item) => {
    const textNodes = Array.from(item.querySelectorAll("t"));
    return textNodes.map((node) => node.textContent || "").join("");
  });
}

function parseWorksheet(xmlText, sharedStrings) {
  const xml = parseXml(xmlText);
  const rows = [];

  xml.querySelectorAll("sheetData row").forEach((rowNode) => {
    const row = [];

    rowNode.querySelectorAll("c").forEach((cellNode) => {
      const reference = cellNode.getAttribute("r") || "";
      const columnIndex = columnNameToIndex(reference.replace(/\d+/g, ""));
      const type = cellNode.getAttribute("t");
      const valueNode = cellNode.querySelector("v");
      const inlineNode = cellNode.querySelector("is t");
      let value = "";

      if (type === "s" && valueNode) {
        value = sharedStrings[Number(valueNode.textContent)] || "";
      } else if (type === "inlineStr") {
        value = inlineNode?.textContent || "";
      } else {
        value = valueNode?.textContent || "";
      }

      row[columnIndex] = value;
    });

    rows.push(row.map((value) => value ?? ""));
  });

  return rows;
}

function parseXml(xmlText) {
  const xml = new DOMParser().parseFromString(xmlText, "application/xml");
  const error = xml.querySelector("parsererror");
  if (error) throw new Error("לא הצלחתי לקרוא את מבנה קובץ האקסל.");
  return xml;
}

function columnNameToIndex(columnName) {
  if (!columnName) return 0;

  return columnName.split("").reduce((index, char) => {
    return index * 26 + char.toUpperCase().charCodeAt(0) - 64;
  }, 0) - 1;
}

function normalizeZipPath(path) {
  const parts = [];
  path.split("/").forEach((part) => {
    if (!part || part === ".") return;
    if (part === "..") {
      parts.pop();
      return;
    }
    parts.push(part);
  });
  return parts.join("/");
}

function getValue(row, aliases) {
  const normalizedRow = Object.entries(row).reduce((accumulator, [key, value]) => {
    accumulator[normalizeHeader(key)] = value;
    return accumulator;
  }, {});

  for (const alias of aliases) {
    const value = normalizedRow[normalizeHeader(alias)];
    if (value !== undefined && String(value).trim() !== "") return String(value).trim();
  }

  return "";
}

function parseNumber(value) {
  if (value === null || value === undefined || value === "") return 0;
  const normalized = String(value).replace(/,/g, "").replace(/[^\d.-]/g, "");
  const number = Number(normalized);
  return Number.isFinite(number) ? number : NaN;
}

function parseOptionalNumber(value) {
  if (value === null || value === undefined || String(value).trim() === "") return NaN;
  const text = String(value).trim();
  if (text.includes("/") && !/[.,]/.test(text)) return NaN;
  return parseNumber(text);
}

function normalizeHeader(value) {
  return String(value || "")
    .replace(/^\uFEFF/, "")
    .replace(/\s+/g, " ")
    .trim();
}

function formatNumber(value) {
  return Number(value || 0).toLocaleString("he-IL");
}

function formatDateTime(date) {
  return new Intl.DateTimeFormat("he-IL", {
    dateStyle: "short",
    timeStyle: "short",
  }).format(date);
}

function setStatus(status, text) {
  elements.statusBadge.className = `status-badge ${status}`;
  elements.statusBadge.innerHTML = `<span class="status-dot"></span>${text}`;
}

function showMessage(text) {
  elements.messageArea.textContent = text;
  elements.messageArea.hidden = false;
  window.setTimeout(() => {
    elements.messageArea.hidden = true;
  }, 4500);
}

function saveLocalState() {
  localStorage.setItem(STORAGE_KEYS.shipments, JSON.stringify(state.manualShipments));
  localStorage.setItem(STORAGE_KEYS.ignored, JSON.stringify(state.ignoredSkus));
}

function loadJson(key, fallback) {
  try {
    return JSON.parse(localStorage.getItem(key) || JSON.stringify(fallback));
  } catch {
    return fallback;
  }
}

function escapeCsvCell(value) {
  const text = String(value ?? "");
  return `"${text.replace(/"/g, '""')}"`;
}

function iconMarkup(type) {
  if (type === "shipment") {
    return `
      <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
        <path d="M8 7h12m0 0-4-4m4 4-4 4" />
        <path d="M16 17H4m0 0 4 4m-4-4 4-4" />
      </svg>
    `;
  }

  return `
    <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path d="M3 3l18 18" />
      <path d="M10.6 10.7a2 2 0 0 0 2.7 2.7" />
      <path d="M9.9 4.2A10.9 10.9 0 0 1 12 4c5 0 8.5 4.1 10 8a13.2 13.2 0 0 1-2.3 3.8" />
      <path d="M6.6 6.7A13.4 13.4 0 0 0 2 12c1.5 3.9 5 8 10 8a10.6 10.6 0 0 0 4.1-.8" />
    </svg>
  `;
}

window.shortageApp = {
  state,
  parseCsv,
  parseXlsx,
  readFileRows,
  handleBulkFiles,
  processAndRender,
};
