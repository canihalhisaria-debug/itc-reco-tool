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
  aggregate2b: true,
  aggregatePr: true,
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
const aggregate2bInput = document.getElementById("aggregate2b");
const aggregatePrInput = document.getElementById("aggregatePr");
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

[aggregate2bInput, aggregatePrInput].forEach((input) => {
  input.addEventListener("change", syncSettingsFromUi);
});

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
  aggregate2bInput.checked = Boolean(state.settings.aggregate2b);
  aggregatePrInput.checked = Boolean(state.settings.aggregatePr);
}

function syncSettingsFromUi() {
  state.settings.taxableTolerance = toNumber(taxableToleranceInput.value);
  state.settings.taxTolerance = toNumber(taxToleranceInput.value);
  state.settings.aggregate2b = aggregate2bInput.checked;
  state.settings.aggregatePr = aggregatePrInput.checked;
}

function normalizeRow(row, mapping, sourceIndex, sourceType) {
  const taxableValue = roundTo2(toNumber(getMapped(row, mapping, "Taxable Value")));
  const igst = roundTo2(toNumber(getMapped(row, mapping, "IGST")));
  const cgst = roundTo2(toNumber(getMapped(row, mapping, "CGST")));
  const sgst = roundTo2(toNumber(getMapped(row, mapping, "SGST")));
  const cess = mapping.CESS ? roundTo2(toNumber(getMapped(row, mapping, "CESS"))) : 0;
  const baseGST = roundTo2(igst + cgst + sgst);
  const computedTotalTax = roundTo2(baseGST + cess);

  return {
    sourceType,
    sourceIndex,
    gstin: normalizeGSTIN(getMapped(row, mapping, "GSTIN")),
    supplierName: normalizeText(getMapped(row, mapping, "Supplier Name")),
    invoiceNo: normalizeText(getMapped(row, mapping, "Invoice No")),
    invoiceNoNorm: normalizeInvoiceNo(getMapped(row, mapping, "Invoice No")),
    invoiceDate: normalizeDate(getMapped(row, mapping, "Invoice Date")),
    taxableValue,
    igst,
    cgst,
    sgst,
    cess,
    baseGST,
    computedTotalTax,
    sourceRowNos: [sourceIndex],
  };
}

function buildGroupKey(row) {
  if (!row.invoiceNoNorm) return "";
  if (row.gstin && row.invoiceDate) return `GSTIN_DATE||${row.gstin}||${row.invoiceNoNorm}||${row.invoiceDate}`;
  if (row.gstin && !row.invoiceDate) return `GSTIN_ONLY||${row.gstin}||${row.invoiceNoNorm}`;
  if (row.supplierName && row.invoiceDate) return `SUPPLIER_DATE||${normalizeSupplierName(row.supplierName)}||${row.invoiceNoNorm}||${row.invoiceDate}`;
  if (row.supplierName) return `SUPPLIER_ONLY||${normalizeSupplierName(row.supplierName)}||${row.invoiceNoNorm}`;
  return "";
}

function aggregateRows(rows, shouldAggregate) {
  if (!shouldAggregate) {
    return rows.map((row) => ({ ...row, groupKey: `ROW||${row.sourceType}||${row.sourceIndex}` }));
  }

  const grouped = [];
  const map = new Map();

  rows.forEach((row) => {
    const key = buildGroupKey(row);
    if (!key) {
      grouped.push({ ...row, groupKey: `NOINV||${row.sourceType}||${row.sourceIndex}` });
      return;
    }

    if (!map.has(key)) {
      map.set(key, {
        ...row,
        groupKey: key,
        sourceRowNos: [],
      });
    }

    const target = map.get(key);
    target.taxableValue = roundTo2(target.taxableValue + row.taxableValue);
    target.igst = roundTo2(target.igst + row.igst);
    target.cgst = roundTo2(target.cgst + row.cgst);
    target.sgst = roundTo2(target.sgst + row.sgst);
    target.cess = roundTo2(target.cess + row.cess);
    target.baseGST = roundTo2(target.baseGST + row.baseGST);
    target.computedTotalTax = roundTo2(target.computedTotalTax + row.computedTotalTax);
    target.sourceRowNos.push(...row.sourceRowNos);
  });

  return grouped.concat(Array.from(map.values()));
}

function reconcile() {
  syncSettingsFromUi();

  const normalizedPrRows = state.rawBooks.map((row, idx) => normalizeRow(row, state.mappedBooks, idx, "PR"));
  const normalized2bRows = state.raw2b.map((row, idx) => normalizeRow(row, state.mapped2b, idx, "2B"));
  const prGroups = aggregateRows(normalizedPrRows, state.settings.aggregatePr);
  const b2Groups = aggregateRows(normalized2bRows, state.settings.aggregate2b);

  const b2ByKey = new Map(b2Groups.map((row) => [row.groupKey, row]));
  const matchedKeys = new Set();

  const prRemarks = new Array(normalizedPrRows.length).fill(REMARKS.NOT_IN_2B);
  const b2Remarks = new Array(normalized2bRows.length).fill(REMARKS.NOT_IN_PR);

  prGroups.forEach((prGroup) => {
    const b2Group = b2ByKey.get(prGroup.groupKey);
    if (!b2Group || prGroup.groupKey.startsWith("NOINV||")) {
      prGroup.sourceRowNos.forEach((idx) => {
        prRemarks[idx] = REMARKS.NOT_IN_2B;
      });
      return;
    }

    matchedKeys.add(prGroup.groupKey);
    const taxableDiff = Math.abs(roundTo2(prGroup.taxableValue - b2Group.taxableValue));
    const taxDiff = Math.abs(roundTo2(prGroup.baseGST - b2Group.baseGST));
    const cessDiff = Math.abs(roundTo2(prGroup.cess - b2Group.cess));
    const withinTolerance = taxableDiff <= state.settings.taxableTolerance && taxDiff <= state.settings.taxTolerance && cessDiff <= state.settings.taxTolerance;
    const remark = withinTolerance ? REMARKS.MATCHED : REMARKS.VALUE_DIFFERENCE;

    prGroup.sourceRowNos.forEach((idx) => {
      prRemarks[idx] = remark;
    });

    b2Group.sourceRowNos.forEach((idx) => {
      b2Remarks[idx] = remark;
    });
  });

  b2Groups.forEach((b2Group) => {
    if (b2Group.groupKey.startsWith("NOINV||") || !matchedKeys.has(b2Group.groupKey)) {
      b2Group.sourceRowNos.forEach((idx) => {
        b2Remarks[idx] = REMARKS.NOT_IN_PR;
      });
    }
  });

  const prExportRows = normalizedPrRows.map((row, idx) => toDisplayRow(row, prRemarks[idx]));
  const b2ExportRows = normalized2bRows.map((row, idx) => toDisplayRow(row, b2Remarks[idx]));

  state.results = {
    Matched: prExportRows.filter((r) => r.Remark === REMARKS.MATCHED).sort(compareBusinessExportRows),
    "Value Difference": prExportRows.filter((r) => r.Remark === REMARKS.VALUE_DIFFERENCE).sort(compareBusinessExportRows),
    "Not in 2B": prExportRows.filter((r) => r.Remark === REMARKS.NOT_IN_2B).sort(compareBusinessExportRows),
    "Not in PR": b2ExportRows.filter((r) => r.Remark === REMARKS.NOT_IN_PR).sort(compareBusinessExportRows),
    PurchaseRegisterExport: prExportRows.sort(compareBusinessExportRows),
    GSTR2BExport: b2ExportRows.sort(compareBusinessExportRows),
  };

  document.getElementById("totalBooks").textContent = String(normalizedPrRows.length);
  document.getElementById("total2b").textContent = String(normalized2bRows.length);
  document.getElementById("matchedCount").textContent = String(state.results.Matched.length);
  document.getElementById("missing2bCount").textContent = String(state.results["Not in 2B"].length);
  document.getElementById("missingBooksCount").textContent = String(state.results["Not in PR"].length);
  document.getElementById("valueDiffCount").textContent = String(state.results["Value Difference"].length);

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
  const tabName = `${type}_with_Remarks`;
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

function getMapped(row, mapping, key) {
  const column = mapping[key];
  return column ? row[column] : "";
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
