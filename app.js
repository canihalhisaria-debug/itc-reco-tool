const FIELD_DEFS = [
  { key: "GSTIN", label: "GSTIN", optional: true },
  { key: "Supplier Name", label: "Supplier Name", optional: true },
  { key: "Invoice No", label: "Invoice No", optional: true },
  { key: "Invoice Date", label: "Invoice Date", optional: true },
  { key: "Taxable Value", label: "Taxable Value", optional: true },
  { key: "Total Tax", label: "Total Tax", optional: true },
  { key: "IGST", label: "IGST", optional: true },
  { key: "CGST", label: "CGST", optional: true },
  { key: "SGST", label: "SGST", optional: true },
  { key: "CESS", label: "CESS", optional: true },
];

const DEFAULT_SETTINGS = {
  taxableTolerance: 1,
  taxTolerance: 1,
  monthOnly: true,
};

const REMARKS = {
  MATCHED: "Matched",
  VALUE_DIFFERENCE: "Value Difference",
  NOT_IN_2B: "Not in 2B",
  NOT_IN_PR: "Not in PR",
};

const EXPORT_COLUMNS = ["GSTIN", "SupplierName", "InvoiceNo", "InvoiceDate", "TaxableValue", "IGST", "CGST", "SGST", "CESS", "ComputedTotalTax", "Remark"];

const state = {
  raw2b: [],
  rawBooks: [],
  headers2b: [],
  headersBooks: [],
  mapped2b: {},
  mappedBooks: {},
  settings: { ...DEFAULT_SETTINGS },
  results: {
    Matched: [],
    "Value Difference": [],
    "Not in 2B": [],
    "Not in PR": [],
    PurchaseRegisterExport: [],
    GSTR2BExport: [],
  },
  activeTab: "Matched",
  resultSearch: "",
};

const file2bInput = document.getElementById("file2b");
const fileBooksInput = document.getElementById("fileBooks");
const status2b = document.getElementById("status2b");
const statusBooks = document.getElementById("statusBooks");
const map2bEl = document.getElementById("map2b");
const mapBooksEl = document.getElementById("mapBooks");
const reconcileBtn = document.getElementById("reconcileBtn");
const tabButtons = document.getElementById("tabButtons");
const resultTable = document.getElementById("resultTable");
const exportPrBtn = document.getElementById("exportPrBtn");
const export2bBtn = document.getElementById("export2bBtn");
const taxableToleranceInput = document.getElementById("taxableTolerance");
const taxToleranceInput = document.getElementById("taxTolerance");
const monthOnlyInput = document.getElementById("monthOnly");
const resultSearchInput = document.getElementById("resultSearch");

file2bInput.addEventListener("change", (e) => handleFile(e.target.files[0], "2b"));
fileBooksInput.addEventListener("change", (e) => handleFile(e.target.files[0], "books"));
reconcileBtn.addEventListener("click", reconcile);
exportPrBtn.addEventListener("click", () => exportDataset("PR"));
export2bBtn.addEventListener("click", () => exportDataset("2B"));

[taxableToleranceInput, taxToleranceInput].forEach((input) => {
  input.addEventListener("change", syncSettingsFromUi);
  input.addEventListener("input", syncSettingsFromUi);
});

if (monthOnlyInput) {
  monthOnlyInput.addEventListener("change", syncSettingsFromUi);
}

resultSearchInput.addEventListener("input", () => {
  state.resultSearch = normalizeText(resultSearchInput.value).toLowerCase();
  renderTable();
});

syncSettingsToUi();

async function handleFile(file, type) {
  if (!file) return;
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: "array", cellDates: false });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(firstSheet, { defval: "", raw: false });

  if (type === "2b") {
    state.raw2b = rows;
    state.headers2b = rows.length ? Object.keys(rows[0]) : [];
    state.mapped2b = {};
    status2b.textContent = `2B: Loaded ${rows.length} rows`;
    status2b.classList.add("ok");
    renderMapping("2b");
  } else {
    state.rawBooks = rows;
    state.headersBooks = rows.length ? Object.keys(rows[0]) : [];
    state.mappedBooks = {};
    statusBooks.textContent = `Books: Loaded ${rows.length} rows`;
    statusBooks.classList.add("ok");
    renderMapping("books");
  }

  updateReconcileButtonState();
}

function renderMapping(type) {
  const target = type === "2b" ? map2bEl : mapBooksEl;
  const headers = type === "2b" ? state.headers2b : state.headersBooks;
  const mapped = type === "2b" ? state.mapped2b : state.mappedBooks;

  target.innerHTML = "";
  FIELD_DEFS.forEach((field) => {
    const row = document.createElement("div");
    row.className = "mapping-row";

    const label = document.createElement("label");
    label.innerHTML = `${escapeHtml(field.label)} ${field.optional ? '<span class="optional-tag">Optional</span>' : '<span class="required-tag">Required</span>'}`;

    const select = document.createElement("select");
    select.innerHTML = `<option value="">Not mapped</option>${headers
      .map((h) => `<option value="${escapeHtml(h)}">${escapeHtml(h)}</option>`)
      .join("")}`;

    select.value = mapped[field.key] || "";
    select.addEventListener("change", () => {
      mapped[field.key] = select.value;
      updateReconcileButtonState();
    });

    row.appendChild(label);
    row.appendChild(select);
    target.appendChild(row);
  });
}

function hasBaseGstSource(mapping) {
  return Boolean(mapping["IGST"]) || Boolean(mapping["CGST"]) || Boolean(mapping["SGST"]);
}

function hasMinimumMapping(mapping) {
  const hasTaxable = Boolean(mapping["Taxable Value"]);
  const hasTaxComponents = hasBaseGstSource(mapping) || Boolean(mapping["Total Tax"]);
  return hasTaxable && hasTaxComponents;
}

function updateReconcileButtonState() {
  const ready = state.raw2b.length > 0 && state.rawBooks.length > 0 && hasMinimumMapping(state.mapped2b) && hasMinimumMapping(state.mappedBooks);
  reconcileBtn.disabled = !ready;
}

function syncSettingsToUi() {
  taxableToleranceInput.value = String(state.settings.taxableTolerance);
  taxToleranceInput.value = String(state.settings.taxTolerance);
  if (monthOnlyInput) monthOnlyInput.checked = Boolean(state.settings.monthOnly);
}

function syncSettingsFromUi() {
  state.settings.taxableTolerance = toNumber(taxableToleranceInput.value);
  state.settings.taxTolerance = toNumber(taxToleranceInput.value);
  if (monthOnlyInput) state.settings.monthOnly = monthOnlyInput.checked;
}

function normalizeRow(row, mapping, sourceIndex, sourceType) {
  const taxableValue = roundTo2(toNumber(getMapped(row, mapping, "Taxable Value")));
  const igst = roundTo2(toNumber(getMapped(row, mapping, "IGST")));
  const cgst = roundTo2(toNumber(getMapped(row, mapping, "CGST")));
  const sgst = roundTo2(toNumber(getMapped(row, mapping, "SGST")));
  const cess = mapping.CESS ? roundTo2(toNumber(getMapped(row, mapping, "CESS"))) : 0;
  const dateSource = getMapped(row, mapping, "Invoice Date");
  const invoiceDate = normalizeDate(dateSource);
  const computedTotalTax = roundTo2(computeTotalTax({ igst, cgst, sgst, cess }));

  return {
    sourceType,
    sourceIndex,
    gstin: normalizeGSTIN(getMapped(row, mapping, "GSTIN")),
    supplierName: normalizeText(getMapped(row, mapping, "Supplier Name")),
    invoiceNo: normalizeText(getMapped(row, mapping, "Invoice No")),
    invoiceNoNorm: normalizeInvoiceNo(getMapped(row, mapping, "Invoice No")),
    invoiceDate,
    invoiceMonth: getInvoiceMonth(invoiceDate),
    taxableValue,
    igst,
    cgst,
    sgst,
    cess,
    computedTotalTax,
    lineCount: 1,
  };
}

function getMapped(row, mapping, key) {
  const column = mapping[key];
  return column ? row[column] : "";
}

function getInvoiceMonth(invoiceDate) {
  if (!invoiceDate) return "NA";
  return invoiceDate.slice(0, 7);
}

function buildCanonicalKey(row) {
  if (!row.invoiceNoNorm) return "";
  const datePart = state.settings.monthOnly ? row.invoiceMonth : row.invoiceDate || "NA";
  return `${normalizeGSTIN(row.gstin)}||${normalizeInvoiceNo(row.invoiceNo)}||${datePart}`;
}

function aggregateRows(rows) {
  const aggregated = [];
  const invoiceMap = new Map();

  rows.forEach((row) => {
    const key = buildCanonicalKey(row);
    if (!key) {
      aggregated.push({ ...row, canonicalKey: `NOINV||${row.sourceType}||${row.sourceIndex}` });
      return;
    }

    if (!invoiceMap.has(key)) {
      invoiceMap.set(key, {
        ...row,
        canonicalKey: key,
        lineCount: 0,
      });
    }

    const target = invoiceMap.get(key);
    target.taxableValue = roundTo2(target.taxableValue + row.taxableValue);
    target.igst = roundTo2(target.igst + row.igst);
    target.cgst = roundTo2(target.cgst + row.cgst);
    target.sgst = roundTo2(target.sgst + row.sgst);
    target.cess = roundTo2(target.cess + row.cess);
    target.computedTotalTax = roundTo2(target.computedTotalTax + row.computedTotalTax);
    target.lineCount += 1;
    if (!target.supplierName && row.supplierName) target.supplierName = row.supplierName;
    if ((!target.invoiceDate || row.invoiceDate < target.invoiceDate) && row.invoiceDate) {
      target.invoiceDate = row.invoiceDate;
      target.invoiceMonth = row.invoiceMonth;
    }
  });

  return aggregated.concat(Array.from(invoiceMap.values()));
}

function reconcile() {
  syncSettingsFromUi();

  const prRows = aggregateRows(state.rawBooks.map((row, idx) => normalizeRow(row, state.mappedBooks, idx, "PR")));
  const twoBRows = aggregateRows(state.raw2b.map((row, idx) => normalizeRow(row, state.mapped2b, idx, "2B")));

  const twoBByKey = new Map(twoBRows.map((row) => [row.canonicalKey, row]));
  const matchedKeys = new Set();

  const matched = [];
  const valueDifference = [];
  const notIn2B = [];
  const notInPR = [];

  const prExportRows = [];
  const twoBExportRows = [];

  prRows.forEach((prRow) => {
    const twoBRow = twoBByKey.get(prRow.canonicalKey);
    if (!twoBRow || prRow.canonicalKey.startsWith("NOINV||")) {
      const row = toDisplayRow(prRow, REMARKS.NOT_IN_2B);
      notIn2B.push(row);
      prExportRows.push(row);
      return;
    }

    matchedKeys.add(prRow.canonicalKey);
    const taxableDiff = Math.abs(roundTo2(prRow.taxableValue - twoBRow.taxableValue));
    const taxDiff = Math.abs(roundTo2(prRow.computedTotalTax - twoBRow.computedTotalTax));
    const remark = taxableDiff <= state.settings.taxableTolerance && taxDiff <= state.settings.taxTolerance ? REMARKS.MATCHED : REMARKS.VALUE_DIFFERENCE;

    const prDisplay = toDisplayRow(prRow, remark);
    const twoBDisplay = toDisplayRow(twoBRow, remark);
    if (remark === REMARKS.MATCHED) matched.push(prDisplay);
    else valueDifference.push(prDisplay);
    prExportRows.push(prDisplay);
    twoBExportRows.push(twoBDisplay);
  });

  twoBRows.forEach((twoBRow) => {
    if (twoBRow.canonicalKey.startsWith("NOINV||") || !matchedKeys.has(twoBRow.canonicalKey)) {
      const row = toDisplayRow(twoBRow, REMARKS.NOT_IN_PR);
      notInPR.push(row);
      twoBExportRows.push(row);
    }
  });

  state.results = {
    Matched: matched.sort(compareBusinessExportRows),
    "Value Difference": valueDifference.sort(compareBusinessExportRows),
    "Not in 2B": notIn2B.sort(compareBusinessExportRows),
    "Not in PR": notInPR.sort(compareBusinessExportRows),
    PurchaseRegisterExport: prExportRows.sort(compareBusinessExportRows),
    GSTR2BExport: twoBExportRows.sort(compareBusinessExportRows),
  };

  document.getElementById("totalBooks").textContent = String(prRows.length);
  document.getElementById("total2b").textContent = String(twoBRows.length);
  document.getElementById("matchedCount").textContent = String(matched.length);
  document.getElementById("missing2bCount").textContent = String(notIn2B.length);
  document.getElementById("missingBooksCount").textContent = String(notInPR.length);
  document.getElementById("valueDiffCount").textContent = String(valueDifference.length);

  state.activeTab = "Matched";
  renderTabs();
  renderTable();
  exportPrBtn.disabled = false;
  export2bBtn.disabled = false;
}

function toDisplayRow(baseRow, remark) {
  return {
    GSTIN: baseRow?.gstin || "",
    SupplierName: baseRow?.supplierName || "",
    InvoiceNo: baseRow?.invoiceNo || "",
    InvoiceDate: baseRow?.invoiceDate || "",
    TaxableValue: baseRow?.taxableValue ?? "",
    IGST: baseRow?.igst ?? "",
    CGST: baseRow?.cgst ?? "",
    SGST: baseRow?.sgst ?? "",
    CESS: baseRow?.cess ?? 0,
    ComputedTotalTax: baseRow?.computedTotalTax ?? 0,
    Remark: remark,
  };
}

function renderTabs() {
  const tabs = ["Matched", "Value Difference", "Not in 2B", "Not in PR"];
  tabButtons.innerHTML = "";

  tabs.forEach((tab) => {
    const btn = document.createElement("button");
    btn.className = `tab-btn ${state.activeTab === tab ? "active" : ""}`;
    btn.textContent = `${tab} (${state.results[tab].length})`;
    btn.addEventListener("click", () => {
      state.activeTab = tab;
      renderTabs();
      renderTable();
    });
    tabButtons.appendChild(btn);
  });
}

function renderTable() {
  const rows = getActiveRows();
  if (!rows.length) {
    resultTable.innerHTML = "<tr><td>No records found.</td></tr>";
    return;
  }

  const thead = `<thead><tr>${EXPORT_COLUMNS.map((col) => `<th class="${getColumnClass(col, true)}">${escapeHtml(col)}</th>`).join("")}</tr></thead>`;
  const tbody = rows.map((row) => `<tr>${EXPORT_COLUMNS.map((col) => renderCell(row, col)).join("")}</tr>`).join("");
  resultTable.innerHTML = thead + `<tbody>${tbody}</tbody>`;
}

function renderCell(row, col) {
  const className = getColumnClass(col, false);
  if (col === "Remark") {
    return `<td class="${className}"><span class="remark-badge ${getRemarkClass(row.Remark)}">${escapeHtml(formatCell(row[col]))}</span></td>`;
  }
  return `<td class="${className}">${escapeHtml(formatCell(row[col]))}</td>`;
}

function getColumnClass(col, isHeader) {
  const lower = String(col).toLowerCase();
  const classes = [];
  if (lower.includes("date")) classes.push("col-date");
  if (["taxablevalue", "igst", "cgst", "sgst", "cess", "computedtotaltax"].includes(lower.replace(/[^a-z0-9]/g, ""))) classes.push("col-num");
  if (lower.includes("supplier")) classes.push("col-supplier");
  if (lower.includes("invoice") && lower.includes("no")) classes.push("col-invoice");
  if (isHeader) classes.push("col-head");
  return classes.join(" ");
}

function getRemarkClass(remark) {
  if (!remark) return "";
  return `remark-${String(remark)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "")}`;
}

function exportDataset(type) {
  const sourceRows = type === "PR" ? state.results.PurchaseRegisterExport : state.results.GSTR2BExport;
  const rows = sourceRows.map((row) => formatExportRow(row));
  const sheet = XLSX.utils.json_to_sheet(rows.length ? rows : [createEmptyExportRow()], { header: EXPORT_COLUMNS });
  applyWorksheetTemplate(sheet, rows.length + 1);

  const workbook = XLSX.utils.book_new();
  const tabName = `${type}_${state.activeTab.replace(/\s+/g, "_")}`;
  XLSX.utils.book_append_sheet(workbook, sheet, tabName.slice(0, 31));
  XLSX.writeFile(workbook, `${tabName}_${getTodayStamp()}.xlsx`);
}

function formatExportRow(row) {
  return {
    GSTIN: row.GSTIN || "",
    SupplierName: row.SupplierName || "",
    InvoiceNo: row.InvoiceNo || "",
    InvoiceDate: row.InvoiceDate || "",
    TaxableValue: row.TaxableValue ?? "",
    IGST: row.IGST ?? "",
    CGST: row.CGST ?? "",
    SGST: row.SGST ?? "",
    CESS: row.CESS ?? 0,
    ComputedTotalTax: row.ComputedTotalTax ?? 0,
    Remark: sanitizeRemark(row.Remark),
  };
}

function createEmptyExportRow() {
  return EXPORT_COLUMNS.reduce((acc, col) => {
    acc[col] = ["CESS", "ComputedTotalTax"].includes(col) ? 0 : "";
    return acc;
  }, {});
}

function sanitizeRemark(remark) {
  const allowed = new Set(Object.values(REMARKS));
  return allowed.has(remark) ? remark : "";
}

function getTodayStamp() {
  const now = new Date();
  const yyyy = now.getFullYear();
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  const dd = String(now.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function applyWorksheetTemplate(sheet, rowCount) {
  sheet["!autofilter"] = { ref: `A1:K1` };
  sheet["!freeze"] = { xSplit: 0, ySplit: 1, topLeftCell: "A2", activePane: "bottomLeft", state: "frozen" };
  sheet["!cols"] = [{ wch: 16 }, { wch: 32 }, { wch: 18 }, { wch: 14 }, { wch: 14 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 16 }, { wch: 18 }];

  const headerStyle = {
    font: { bold: true, color: { rgb: "1F2937" } },
    fill: { fgColor: { rgb: "E5E7EB" } },
    alignment: { horizontal: "center", vertical: "center" },
  };

  for (let c = 0; c < EXPORT_COLUMNS.length; c += 1) {
    const cell = XLSX.utils.encode_cell({ r: 0, c });
    if (sheet[cell]) sheet[cell].s = headerStyle;
  }

  for (let r = 1; r < rowCount; r += 1) {
    [4, 5, 6, 7, 8, 9].forEach((c) => {
      const ref = XLSX.utils.encode_cell({ r, c });
      if (sheet[ref]) sheet[ref].s = { alignment: { horizontal: "right", vertical: "center" } };
    });
    const dateRef = XLSX.utils.encode_cell({ r, c: 3 });
    if (sheet[dateRef]) sheet[dateRef].s = { alignment: { horizontal: "center", vertical: "center" } };

    const remarkRef = XLSX.utils.encode_cell({ r, c: 10 });
    if (sheet[remarkRef]) sheet[remarkRef].s = { fill: { fgColor: { rgb: getRemarkFill(sheet[remarkRef].v) } }, alignment: { horizontal: "center", vertical: "center" } };
  }
}

function getRemarkFill(remark) {
  if (remark === REMARKS.MATCHED) return "DCFCE7";
  if (remark === REMARKS.NOT_IN_2B || remark === REMARKS.NOT_IN_PR) return "FEE2E2";
  if (remark === REMARKS.VALUE_DIFFERENCE) return "FEF3C7";
  return "FFFFFF";
}

function compareBusinessExportRows(a, b) {
  const supplierA = normalizeSupplierName(a.SupplierName);
  const supplierB = normalizeSupplierName(b.SupplierName);
  const keyA = supplierA || normalizeGSTIN(a.GSTIN);
  const keyB = supplierB || normalizeGSTIN(b.GSTIN);
  if (keyA !== keyB) return keyA.localeCompare(keyB);
  const dateA = normalizeText(a.InvoiceDate || "9999-99-99");
  const dateB = normalizeText(b.InvoiceDate || "9999-99-99");
  if (dateA !== dateB) return dateA.localeCompare(dateB);
  return normalizeInvoiceNo(a.InvoiceNo).localeCompare(normalizeInvoiceNo(b.InvoiceNo));
}

function getActiveRows() {
  const activeRows = state.results[state.activeTab] || [];
  if (!state.resultSearch) return activeRows;
  return activeRows.filter((row) => {
    const haystack = [row.SupplierName, row.GSTIN, row.InvoiceNo].map((v) => String(v || "").toLowerCase()).join(" ");
    return haystack.includes(state.resultSearch);
  });
}

function normalizeDate(value) {
  if (!value && value !== 0) return "";

  if (typeof value === "number") {
    const date = XLSX.SSF.parse_date_code(value);
    if (date) return `${date.y}-${pad2(date.m)}-${pad2(date.d)}`;
  }

  const str = normalizeText(value);
  if (!str) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;
  if (/^\d{2}[/-]\d{2}[/-]\d{4}$/.test(str)) {
    const [d, m, y] = str.split(/[/-]/);
    return `${y}-${pad2(m)}-${pad2(d)}`;
  }

  const date = new Date(str);
  if (!Number.isNaN(date.getTime())) {
    return `${date.getUTCFullYear()}-${pad2(date.getUTCMonth() + 1)}-${pad2(date.getUTCDate())}`;
  }

  return "";
}

function normalizeInvoiceNo(value) {
  return normalizeText(value)
    .toUpperCase()
    .replace(/[\s/_-]+/g, "")
    .replace(/[^A-Z0-9]/g, "");
}

function normalizeGSTIN(value) {
  return normalizeText(value).toUpperCase();
}

function normalizeText(value) {
  return String(value ?? "")
    .replace(/\u00A0/g, " ")
    .trim()
    .replace(/\s+/g, " ");
}

function normalizeSupplierName(value) {
  return normalizeText(value)
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function formatCell(value) {
  return typeof value === "number" ? value.toFixed(2) : value ?? "";
}

function toNumber(value) {
  const text = normalizeText(value)
    .replace(/,/g, "")
    .replace(/^\((.*)\)$/, "-$1");
  if (!text || text === "-") return 0;
  const n = parseFloat(text);
  return Number.isFinite(n) ? n : 0;
}

function computeTotalTax(row) {
  return roundTo2(toNumber(row.igst) + toNumber(row.cgst) + toNumber(row.sgst) + toNumber(row.cess));
}

function roundTo2(value) {
  return Math.round((Number(value) + Number.EPSILON) * 100) / 100;
}

function pad2(n) {
  return String(n).padStart(2, "0");
}

function escapeHtml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
