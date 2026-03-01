const FIELD_DEFS = [
  { key: "GSTIN", label: "GSTIN", optional: true },
  { key: "Supplier Name", label: "Supplier Name", optional: true },
  { key: "Invoice No", label: "Invoice No", optional: true },
  { key: "Invoice Date", label: "Invoice Date", optional: true },
  { key: "Taxable Value", label: "Taxable Value", optional: false },
  { key: "Total Tax", label: "Total Tax", optional: true },
  { key: "IGST", label: "IGST", optional: true },
  { key: "CGST", label: "CGST", optional: true },
  { key: "SGST", label: "SGST", optional: true },
  { key: "CESS", label: "CESS", optional: true },
];

const DEFAULT_SETTINGS = {
  taxableTolerance: 1,
  taxTolerance: 1,
};

const REMARKS = {
  MATCHED: "Matched",
  VALUE_DIFFERENCE: "Value Difference",
  NOT_IN_2B: "Not in 2B",
  NOT_IN_PR: "Not in PR",
};

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
const exportBtn = document.getElementById("exportBtn");
const taxableToleranceInput = document.getElementById("taxableTolerance");
const taxToleranceInput = document.getElementById("taxTolerance");

file2bInput.addEventListener("change", (e) => handleFile(e.target.files[0], "2b"));
fileBooksInput.addEventListener("change", (e) => handleFile(e.target.files[0], "books"));
reconcileBtn.addEventListener("click", reconcile);
exportBtn.addEventListener("click", exportResults);

[taxableToleranceInput, taxToleranceInput].forEach((input) => {
  input.addEventListener("change", syncSettingsFromUi);
  input.addEventListener("input", syncSettingsFromUi);
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

function hasTaxSource(mapping) {
  return (
    Boolean(mapping["Total Tax"]) ||
    Boolean(mapping["IGST"]) ||
    Boolean(mapping["CGST"]) ||
    Boolean(mapping["SGST"]) ||
    Boolean(mapping["CESS"])
  );
}

function updateReconcileButtonState() {
  const ready =
    state.raw2b.length > 0 &&
    state.rawBooks.length > 0 &&
    Boolean(state.mapped2b["Taxable Value"]) &&
    Boolean(state.mappedBooks["Taxable Value"]) &&
    hasTaxSource(state.mapped2b) &&
    hasTaxSource(state.mappedBooks);

  reconcileBtn.disabled = !ready;
}

function syncSettingsToUi() {
  taxableToleranceInput.value = String(state.settings.taxableTolerance);
  taxToleranceInput.value = String(state.settings.taxTolerance);
}

function syncSettingsFromUi() {
  state.settings.taxableTolerance = toNumber(taxableToleranceInput.value);
  state.settings.taxTolerance = toNumber(taxToleranceInput.value);
}

function normalizeRow(row, mapping, sourceIndex) {
  const taxableValue = roundTo2(toNumber(getMapped(row, mapping, "Taxable Value")));
  const igst = roundTo2(toNumber(getMapped(row, mapping, "IGST")));
  const cgst = roundTo2(toNumber(getMapped(row, mapping, "CGST")));
  const sgst = roundTo2(toNumber(getMapped(row, mapping, "SGST")));
  const cess = roundTo2(toNumber(getMapped(row, mapping, "CESS")));
  const totalTaxMapped = roundTo2(toNumber(getMapped(row, mapping, "Total Tax")));

  const totalTax = mapping["Total Tax"] ? totalTaxMapped : roundTo2(igst + cgst + sgst + cess);

  const normalized = {
    original: { ...row },
    sourceIndex,
    gstin: String(getMapped(row, mapping, "GSTIN")).trim().toUpperCase(),
    supplierName: String(getMapped(row, mapping, "Supplier Name")).trim(),
    supplierNorm: normalizeSupplierName(getMapped(row, mapping, "Supplier Name")),
    invoiceNo: String(getMapped(row, mapping, "Invoice No")).trim(),
    invoiceNoNorm: normalizeInvoiceNo(getMapped(row, mapping, "Invoice No")),
    invoiceDate: normalizeDate(getMapped(row, mapping, "Invoice Date")),
    taxableValue,
    igst,
    cgst,
    sgst,
    cess,
    totalTax,
    used: false,
  };

  normalized.invoiceMonth = monthPart(normalized.invoiceDate);
  return normalized;
}

function getMapped(row, mapping, key) {
  const column = mapping[key];
  return column ? row[column] : "";
}

function monthPart(date) {
  return date ? date.slice(0, 7) : "";
}

function diffSummary(prRow, twoBRow) {
  const taxableDiff = roundTo2(Math.abs(prRow.taxableValue - twoBRow.taxableValue));
  const taxDiff = roundTo2(Math.abs(prRow.totalTax - twoBRow.totalTax));
  const taxableTolerance = roundTo2(state.settings.taxableTolerance);
  const taxTolerance = roundTo2(state.settings.taxTolerance);

  return {
    taxableDiff,
    taxDiff,
    withinTolerance: isWithinTolerance(taxableDiff, taxableTolerance) && isWithinTolerance(taxDiff, taxTolerance),
    combinedDiff: roundTo2(taxableDiff + taxDiff),
  };
}

function findBestByAmount(prRow, candidates) {
  if (!candidates.length) return null;
  const evaluated = candidates.map((row) => ({ row, ...diffSummary(prRow, row) }));
  const withinTolerance = evaluated.filter((item) => item.withinTolerance);
  const pool = withinTolerance.length ? withinTolerance : evaluated;
  pool.sort((a, b) => a.combinedDiff - b.combinedDiff || a.row.sourceIndex - b.row.sourceIndex);
  return {
    ...pool[0],
    matchType: withinTolerance.length ? "MATCHED" : "VALUE_DIFFERENCE",
  };
}

function findCandidate(prRow, twoBRows) {
  const available = twoBRows.filter((row) => !row.used);
  if (!available.length) return null;

  if (prRow.gstin && prRow.invoiceNoNorm && prRow.invoiceDate) {
    const strict = available.find(
      (row) => row.gstin === prRow.gstin && row.invoiceNoNorm === prRow.invoiceNoNorm && row.invoiceDate === prRow.invoiceDate
    );
    if (strict) {
      const diff = diffSummary(prRow, strict);
      return { row: strict, ...diff, matchType: diff.withinTolerance ? "MATCHED" : "VALUE_DIFFERENCE", mode: "Strict" };
    }
  }

  if (prRow.gstin) {
    const gstPool = available.filter((row) => row.gstin === prRow.gstin);
    const amountMatch = findBestByAmount(prRow, gstPool);
    if (amountMatch) {
      return { ...amountMatch, mode: "Amount-Based" };
    }
  }

  const fallbackPool = available.filter((row) => {
    if (!prRow.supplierNorm || !row.supplierNorm) return false;
    return supplierSimilarity(prRow.supplierNorm, row.supplierNorm) >= 0.7;
  });

  const fallbackMatch = findBestByAmount(prRow, fallbackPool);
  if (fallbackMatch) {
    return { ...fallbackMatch, mode: "Amount-Only Fallback" };
  }

  return null;
}

function reconcile() {
  syncSettingsFromUi();

  const prRows = state.rawBooks.map((row, idx) => normalizeRow(row, state.mappedBooks, idx));
  const twoBRows = state.raw2b.map((row, idx) => normalizeRow(row, state.mapped2b, idx));

  const matched = [];
  const valueDifference = [];
  const notIn2B = [];

  const prExportRows = [];
  const twoBExportRows = new Array(twoBRows.length);

  prRows.forEach((prRow) => {
    const candidate = findCandidate(prRow, twoBRows);

    if (!candidate) {
      notIn2B.push(buildOutcomeRow(prRow, null, REMARKS.NOT_IN_2B));
      prExportRows.push(buildPrExportRow(prRow, null, REMARKS.NOT_IN_2B));
      return;
    }

    candidate.row.used = true;
    const remark = REMARKS[candidate.matchType];
    const outcome = buildOutcomeRow(prRow, candidate.row, remark, candidate.taxableDiff, candidate.taxDiff, candidate.mode);

    if (candidate.matchType === "MATCHED") {
      matched.push(outcome);
    } else {
      valueDifference.push(outcome);
    }

    prExportRows.push(buildPrExportRow(prRow, candidate.row, remark, candidate.taxableDiff, candidate.taxDiff, candidate.mode));
    twoBExportRows[candidate.row.sourceIndex] = build2BExportRow(
      candidate.row,
      prRow,
      remark,
      candidate.taxableDiff,
      candidate.taxDiff,
      candidate.mode
    );
  });

  const notInPR = [];
  twoBRows.forEach((twoBRow) => {
    if (twoBRow.used) return;
    const remark = REMARKS.NOT_IN_PR;
    const outcome = buildOutcomeRow(null, twoBRow, remark);
    notInPR.push(outcome);
    twoBExportRows[twoBRow.sourceIndex] = build2BExportRow(twoBRow, null, remark);
  });

  state.results = {
    Matched: matched,
    "Value Difference": valueDifference,
    "Not in 2B": notIn2B,
    "Not in PR": notInPR,
    PurchaseRegisterExport: prExportRows,
    GSTR2BExport: twoBExportRows.filter(Boolean),
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
  exportBtn.disabled = false;
}

function buildOutcomeRow(prRow, twoBRow, remark, taxableDiff = "", taxDiff = "", mode = "") {
  return {
    Remark: remark,
    MatchMode: mode,
    PR_GSTIN: prRow?.gstin || "",
    TwoB_GSTIN: twoBRow?.gstin || "",
    PR_SupplierName: prRow?.supplierName || "",
    TwoB_SupplierName: twoBRow?.supplierName || "",
    PR_InvoiceNo: prRow?.invoiceNo || "",
    TwoB_InvoiceNo: twoBRow?.invoiceNo || "",
    PR_InvoiceDate: prRow?.invoiceDate || "",
    TwoB_InvoiceDate: twoBRow?.invoiceDate || "",
    PR_TaxableValue: prRow?.taxableValue ?? "",
    TwoB_TaxableValue: twoBRow?.taxableValue ?? "",
    PR_TotalTax: prRow?.totalTax ?? "",
    TwoB_TotalTax: twoBRow?.totalTax ?? "",
    TaxableDiff: taxableDiff,
    TaxDiff: taxDiff,
  };
}

function buildPrExportRow(prRow, twoBRow, remark, taxableDiff = "", taxDiff = "", mode = "") {
  return {
    ...prRow.original,
    Remark: remark,
    MatchMode: mode,
    TotalTax: prRow.totalTax,
    TwoB_TotalTax: twoBRow?.totalTax ?? "",
    TwoB_TaxableValue: twoBRow?.taxableValue ?? "",
    TaxableDiff: taxableDiff,
    TaxDiff: taxDiff,
  };
}

function build2BExportRow(twoBRow, prRow, remark, taxableDiff = "", taxDiff = "", mode = "") {
  return {
    ...twoBRow.original,
    Remark: remark,
    MatchMode: mode,
    TotalTax: twoBRow.totalTax,
    PR_TotalTax: prRow?.totalTax ?? "",
    PR_TaxableValue: prRow?.taxableValue ?? "",
    TaxableDiff: taxableDiff,
    TaxDiff: taxDiff,
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
  const rows = state.results[state.activeTab] || [];
  if (!rows.length) {
    resultTable.innerHTML = "<tr><td>No records found.</td></tr>";
    return;
  }

  const columns = Object.keys(rows[0]);
  const thead = `<thead><tr>${columns.map((col) => `<th>${escapeHtml(col)}</th>`).join("")}</tr></thead>`;
  const tbody = `<tbody>${rows
    .map((row) => {
      return `<tr>${columns
        .map((col) => {
          if (col === "Remark") {
            return `<td><span class="remark-badge ${getRemarkClass(row.Remark)}">${escapeHtml(formatCell(row[col]))}</span></td>`;
          }
          return `<td>${escapeHtml(formatCell(row[col]))}</td>`;
        })
        .join("")}</tr>`;
    })
    .join("")}</tbody>`;

  resultTable.innerHTML = thead + tbody;
}

function getRemarkClass(remark) {
  if (!remark) return "";
  return `remark-${String(remark)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "")}`;
}

function exportResults() {
  const workbook = XLSX.utils.book_new();
  const prRows = state.results.PurchaseRegisterExport;
  const twoBRows = state.results.GSTR2BExport;

  const prSheetRows = prRows.length ? prRows : [{ Info: "No records found" }];
  const twoBSheetRows = twoBRows.length ? twoBRows : [{ Info: "No records found" }];

  const prSheet = XLSX.utils.json_to_sheet(prSheetRows);
  const twoBSheet = XLSX.utils.json_to_sheet(twoBSheetRows);

  styleExportSheet(prSheet, prSheetRows);
  styleExportSheet(twoBSheet, twoBSheetRows);

  XLSX.utils.book_append_sheet(workbook, prSheet, "PR_With_Remarks");
  XLSX.utils.book_append_sheet(workbook, twoBSheet, "GSTR2B_With_Remarks");
  XLSX.writeFile(workbook, "reconciliation_results.xlsx");
}

function styleExportSheet(worksheet, rows) {
  const range = XLSX.utils.decode_range(worksheet["!ref"] || "A1");
  const headers = rows.length ? Object.keys(rows[0]) : [];

  const headerStyle = {
    font: { bold: true, sz: 12 },
    fill: { patternType: "solid", fgColor: { rgb: "FFEFEFEF" } },
  };

  const amountStyle = { numFmt: "#,##0.00" };
  const dateStyle = { numFmt: "dd-mmm-yyyy" };

  const remarkFills = {
    [REMARKS.MATCHED]: { patternType: "solid", fgColor: { rgb: "FFD9EAD3" } },
    [REMARKS.NOT_IN_2B]: { patternType: "solid", fgColor: { rgb: "FFF4CCCC" } },
    [REMARKS.NOT_IN_PR]: { patternType: "solid", fgColor: { rgb: "FFF4CCCC" } },
    [REMARKS.VALUE_DIFFERENCE]: { patternType: "solid", fgColor: { rgb: "FFFFF2CC" } },
  };

  const amountColumns = new Set(
    headers
      .map((header, index) => ({ header, index }))
      .filter(({ header }) => /taxable|igst|cgst|sgst|cess|totaltax|taxdiff|taxablediff|total tax/i.test(header))
      .map(({ index }) => index)
  );

  const dateColumns = new Set(
    headers
      .map((header, index) => ({ header, index }))
      .filter(({ header }) => /date/i.test(header))
      .map(({ index }) => index)
  );

  const remarkIndex = headers.findIndex((header) => /^remark$/i.test(header));

  worksheet["!autofilter"] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: 0, c: range.e.c } }) };
  worksheet["!freeze"] = { xSplit: 0, ySplit: 1, topLeftCell: "A2", activePane: "bottomLeft", state: "frozen" };

  worksheet["!cols"] = headers.map((header) => ({ wch: getColumnWidth(header) }));

  for (let col = range.s.c; col <= range.e.c; col += 1) {
    const address = XLSX.utils.encode_cell({ r: 0, c: col });
    if (!worksheet[address]) continue;
    worksheet[address].s = { ...(worksheet[address].s || {}), ...headerStyle };
  }

  for (let rowIndex = 1; rowIndex <= range.e.r; rowIndex += 1) {
    const remarkCellAddress = remarkIndex >= 0 ? XLSX.utils.encode_cell({ r: rowIndex, c: remarkIndex }) : "";
    const remarkValue = remarkCellAddress && worksheet[remarkCellAddress] ? worksheet[remarkCellAddress].v : "";
    const remarkFill = remarkFills[String(remarkValue || "").trim()];

    for (let col = range.s.c; col <= range.e.c; col += 1) {
      const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: col });
      const cell = worksheet[cellAddress];
      if (!cell) continue;

      let style = { ...(cell.s || {}) };

      if (amountColumns.has(col) && typeof cell.v === "number") {
        style = { ...style, ...amountStyle };
      }

      if (dateColumns.has(col) && cell.v) {
        const parsedDate = normalizeDate(cell.v);
        if (parsedDate) {
          const [year, month, day] = parsedDate.split("-").map(Number);
          cell.v = new Date(Date.UTC(year, month - 1, day));
          cell.t = "d";
          style = { ...style, ...dateStyle };
        }
      }

      if (remarkFill) {
        style = { ...style, fill: remarkFill };
      }

      cell.s = style;
    }
  }
}

function getColumnWidth(header) {
  const normalized = String(header || "").toLowerCase();

  if (normalized.includes("gstin")) return 18;
  if (normalized.includes("supplier")) return 30;
  if (normalized.includes("invoice") && normalized.includes("no")) return 18;
  if (normalized.includes("date")) return 14;
  if (normalized.includes("taxable")) return 16;
  if (normalized.includes("igst") || normalized.includes("cgst") || normalized.includes("sgst") || normalized.includes("cess")) return 14;
  if (normalized.includes("totaltax") || normalized.includes("total tax")) return 14;
  if (normalized.includes("diff")) return 14;
  if (normalized.includes("remark")) return 16;

  return 16;
}

function normalizeDate(value) {
  if (!value) return "";

  if (typeof value === "number") {
    const date = XLSX.SSF.parse_date_code(value);
    if (date) return `${date.y}-${pad2(date.m)}-${pad2(date.d)}`;
  }

  const str = String(value).trim();
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
  return String(value || "")
    .toUpperCase()
    .replace(/[\s\/\-‐‑‒–—―]+/g, "")
    .trim();
}

function normalizeSupplierName(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function supplierSimilarity(a, b) {
  const aTokens = new Set(a.split(" ").filter(Boolean));
  const bTokens = new Set(b.split(" ").filter(Boolean));
  if (!aTokens.size || !bTokens.size) return 0;

  let intersection = 0;
  aTokens.forEach((token) => {
    if (bTokens.has(token)) intersection += 1;
  });
  const union = new Set([...aTokens, ...bTokens]).size;
  return intersection / union;
}

function formatCell(value) {
  return typeof value === "number" ? value.toFixed(2) : value ?? "";
}

function toNumber(value) {
  const n = parseFloat(String(value ?? "").replace(/,/g, "").trim());
  return Number.isFinite(n) ? n : 0;
}

function roundTo2(value) {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

function isWithinTolerance(diff, tolerance) {
  return Math.abs(diff) <= tolerance + Number.EPSILON;
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
