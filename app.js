const FIELD_DEFS = [
  { key: "GSTIN", label: "GSTIN", optional: false },
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
  return Boolean(mapping["Total Tax"]) || Boolean(mapping["IGST"]) || Boolean(mapping["CGST"]) || Boolean(mapping["SGST"]);
}

function hasMinimumMapping(mapping) {
  return Boolean(mapping.GSTIN) && Boolean(mapping["Taxable Value"]) && hasTaxSource(mapping);
}

function updateReconcileButtonState() {
  const ready =
    state.raw2b.length > 0 &&
    state.rawBooks.length > 0 &&
    hasMinimumMapping(state.mapped2b) &&
    hasMinimumMapping(state.mappedBooks);

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
  const dateSource = getMapped(row, mapping, "Invoice Date");
  const invoiceDate = normalizeDate(dateSource);

  const normalized = {
    original: { ...row },
    sourceIndex,
    gstin: normalizeText(getMapped(row, mapping, "GSTIN")).toUpperCase(),
    supplierName: normalizeText(getMapped(row, mapping, "Supplier Name")),
    supplierNorm: normalizeSupplierName(getMapped(row, mapping, "Supplier Name")),
    invoiceNo: normalizeText(getMapped(row, mapping, "Invoice No")),
    invoiceNoNorm: normalizeInvoiceNo(getMapped(row, mapping, "Invoice No")),
    invoiceDate,
    hasInvoiceDate: Boolean(invoiceDate),
    dateParseError: Boolean(normalizeText(dateSource)) && !invoiceDate,
    taxableValue,
    igst,
    cgst,
    sgst,
    cess,
    totalTax,
    used: false,
  };

  return normalized;
}

function getMapped(row, mapping, key) {
  const column = mapping[key];
  return column ? row[column] : "";
}

function diffSummary(prRow, twoBRow) {
  const taxableDiff = roundTo2(Math.abs(prRow.taxableValue - twoBRow.taxableValue));
  const taxDiff = roundTo2(Math.abs(prRow.totalTax - twoBRow.totalTax));
  const taxableTolerance = roundTo2(state.settings.taxableTolerance);
  const taxTolerance = roundTo2(state.settings.taxTolerance);

  return {
    taxableDiff,
    taxDiff,
    taxableWithin: isWithinTolerance(taxableDiff, taxableTolerance),
    taxWithin: isWithinTolerance(taxDiff, taxTolerance),
    withinTolerance: isWithinTolerance(taxableDiff, taxableTolerance) && isWithinTolerance(taxDiff, taxTolerance),
    combinedDiff: roundTo2(taxableDiff + taxDiff),
  };
}

function compareForBest(a, b) {
  if (a.combinedDiff !== b.combinedDiff) return a.combinedDiff - b.combinedDiff;
  return a.row.sourceIndex - b.row.sourceIndex;
}

function findAmountToleranceMatch(prRow, candidateRows) {
  const pool = candidateRows
    .map((row) => ({ row, ...diffSummary(prRow, row) }))
    .filter((item) => item.withinTolerance)
    .sort(compareForBest);
  return pool.length ? pool[0] : null;
}

function findClosestDifference(prRow, candidateRows) {
  const evaluated = candidateRows.map((row) => ({ row, ...diffSummary(prRow, row) })).sort(compareForBest);
  return evaluated.length ? evaluated[0] : null;
}

function datesMatchWhenPresent(prRow, twoBRow) {
  if (prRow.hasInvoiceDate && twoBRow.hasInvoiceDate) {
    return prRow.invoiceDate === twoBRow.invoiceDate;
  }
  return true;
}

function reconcile() {
  syncSettingsFromUi();

  const prRows = state.rawBooks.map((row, idx) => normalizeRow(row, state.mappedBooks, idx));
  const twoBRows = state.raw2b.map((row, idx) => normalizeRow(row, state.mapped2b, idx));

  const matched = [];
  const valueDifference = [];
  const notIn2B = [];
  const notInPR = [];

  const prExportRows = [];
  const twoBExportRows = new Array(twoBRows.length);

  const matchModes = [
    {
      mode: "Exact",
      getCandidates: (prRow) =>
        twoBRows.filter(
          (row) => !row.used && row.gstin === prRow.gstin && row.invoiceNoNorm && row.invoiceNoNorm === prRow.invoiceNoNorm && datesMatchWhenPresent(prRow, row)
        ),
    },
    {
      mode: "InvoiceBased",
      getCandidates: (prRow) =>
        twoBRows.filter((row) => !row.used && row.gstin === prRow.gstin && row.invoiceNoNorm && row.invoiceNoNorm === prRow.invoiceNoNorm),
    },
    {
      mode: "AmountBased",
      getCandidates: (prRow) =>
        twoBRows.filter(
          (row) =>
            !row.used &&
            row.gstin === prRow.gstin &&
            // Amount-based fallback should not override available invoice-based matching.
            (!prRow.invoiceNoNorm || !row.invoiceNoNorm)
        ),
    },
  ];

  for (const strategy of matchModes) {
    prRows.forEach((prRow) => {
      if (prRow.used) return;
      const match = findAmountToleranceMatch(prRow, strategy.getCandidates(prRow));
      if (!match) return;

      prRow.used = true;
      match.row.used = true;

      const reason = buildMatchReason(strategy.mode, prRow, match.row);
      const outcome = buildOutcomeRow(prRow, match.row, REMARKS.MATCHED, match.taxableDiff, match.taxDiff, strategy.mode, reason);
      matched.push(outcome);

      prExportRows.push(buildPrExportRow(prRow, match.row, REMARKS.MATCHED, match.taxableDiff, match.taxDiff, strategy.mode, reason));
      twoBExportRows[match.row.sourceIndex] = build2BExportRow(
        match.row,
        prRow,
        REMARKS.MATCHED,
        match.taxableDiff,
        match.taxDiff,
        strategy.mode,
        reason
      );
    });
  }

  prRows.forEach((prRow) => {
    if (prRow.used) return;

    const gstMatches = twoBRows.filter((row) => !row.used && row.gstin === prRow.gstin);
    const invoiceMatches = gstMatches.filter((row) => row.invoiceNoNorm && row.invoiceNoNorm === prRow.invoiceNoNorm);
    const invoiceOnlyMatches = prRow.invoiceNoNorm
      ? twoBRows.filter((row) => !row.used && row.invoiceNoNorm && row.invoiceNoNorm === prRow.invoiceNoNorm)
      : [];
    const gstAmountFallbackMatches = gstMatches.filter((row) => !prRow.invoiceNoNorm || !row.invoiceNoNorm);

    let fallbackPool = [];
    let mode = "AmountBased";

    if (invoiceMatches.length) {
      fallbackPool = invoiceMatches;
      mode = "InvoiceBased";
    } else if (invoiceOnlyMatches.length) {
      fallbackPool = invoiceOnlyMatches;
      mode = "InvoiceOnly";
    } else if (gstAmountFallbackMatches.length) {
      fallbackPool = gstAmountFallbackMatches;
      mode = "AmountBased";
    }

    const closest = findClosestDifference(prRow, fallbackPool);

    if (!closest) {
      const reason = buildNotFoundReason(prRow);
      notIn2B.push(buildOutcomeRow(prRow, null, REMARKS.NOT_IN_2B, "", "", "", reason));
      prExportRows.push(buildPrExportRow(prRow, null, REMARKS.NOT_IN_2B, "", "", "", reason));
      return;
    }

    prRow.used = true;
    closest.row.used = true;
    const reason = buildDiffReason(prRow, closest.row, closest, mode);
    valueDifference.push(buildOutcomeRow(prRow, closest.row, REMARKS.VALUE_DIFFERENCE, closest.taxableDiff, closest.taxDiff, mode, reason));

    prExportRows.push(buildPrExportRow(prRow, closest.row, REMARKS.VALUE_DIFFERENCE, closest.taxableDiff, closest.taxDiff, mode, reason));
    twoBExportRows[closest.row.sourceIndex] = build2BExportRow(
      closest.row,
      prRow,
      REMARKS.VALUE_DIFFERENCE,
      closest.taxableDiff,
      closest.taxDiff,
      mode,
      reason
    );
  });

  twoBRows.forEach((twoBRow) => {
    if (twoBRow.used) return;
    const reason = buildNotFoundReason(null, twoBRow);
    notInPR.push(buildOutcomeRow(null, twoBRow, REMARKS.NOT_IN_PR, "", "", "", reason));
    twoBExportRows[twoBRow.sourceIndex] = build2BExportRow(twoBRow, null, REMARKS.NOT_IN_PR, "", "", "", reason);
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

function buildMatchReason(mode, prRow, twoBRow) {
  const reasons = ["Matched within taxable/tax tolerance"];
  if (mode === "InvoiceBased" && prRow.hasInvoiceDate && twoBRow.hasInvoiceDate && prRow.invoiceDate !== twoBRow.invoiceDate) {
    reasons.push("date mismatch ignored in InvoiceBased mode");
  }
  if (mode === "AmountBased" && prRow.invoiceNoNorm !== twoBRow.invoiceNoNorm) {
    reasons.push("invoice mismatch accepted in AmountBased fallback");
  }
  if (prRow.dateParseError || twoBRow.dateParseError) {
    reasons.push("date parse fail");
  }
  return reasons.join("; ");
}

function buildDiffReason(prRow, twoBRow, diff, mode) {
  const reasons = [];
  if (!diff.taxableWithin) reasons.push(`taxable diff ${diff.taxableDiff.toFixed(2)} > tolerance`);
  if (!diff.taxWithin) reasons.push(`tax diff ${diff.taxDiff.toFixed(2)} > tolerance`);
  if (prRow.invoiceNoNorm !== twoBRow.invoiceNoNorm) reasons.push("invoice mismatch");
  if (prRow.hasInvoiceDate && twoBRow.hasInvoiceDate && prRow.invoiceDate !== twoBRow.invoiceDate) reasons.push("date mismatch");
  if (prRow.dateParseError || twoBRow.dateParseError) reasons.push("date parse fail");
  if (!reasons.length) reasons.push(`closest ${mode} candidate but outside tolerance`);
  return reasons.join("; ");
}

function buildNotFoundReason(prRow, twoBRow = null) {
  const row = prRow || twoBRow;
  const side = prRow ? "2B" : "PR";
  const reasons = [`No ${side} row available for GSTIN`];
  if (row && row.dateParseError) reasons.push("date parse fail");
  return reasons.join("; ");
}

function buildOutcomeRow(prRow, twoBRow, remark, taxableDiff = "", taxDiff = "", mode = "", reason = "") {
  return {
    Remark: remark,
    MatchMode: mode,
    Reason: reason,
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

function buildPrExportRow(prRow, twoBRow, remark, taxableDiff = "", taxDiff = "", mode = "", reason = "") {
  return {
    ...prRow.original,
    Remark: remark,
    MatchMode: mode,
    Reason: reason,
    TotalTax: prRow.totalTax,
    TwoB_TotalTax: twoBRow?.totalTax ?? "",
    TwoB_TaxableValue: twoBRow?.taxableValue ?? "",
    TaxableDiff: taxableDiff,
    TaxDiff: taxDiff,
  };
}

function build2BExportRow(twoBRow, prRow, remark, taxableDiff = "", taxDiff = "", mode = "", reason = "") {
  return {
    ...twoBRow.original,
    Remark: remark,
    MatchMode: mode,
    Reason: reason,
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
  const prRows = state.results.PurchaseRegisterExport;
  const twoBRows = state.results.GSTR2BExport;

  downloadCsv("PR_with_Remarks.csv", prRows.length ? prRows : [{ Info: "No records found" }]);
  downloadCsv("2B_with_Remarks.csv", twoBRows.length ? twoBRows : [{ Info: "No records found" }]);
}

function downloadCsv(fileName, rows) {
  const sheet = XLSX.utils.json_to_sheet(rows);
  const csv = XLSX.utils.sheet_to_csv(sheet);
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(link.href);
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
    .replace(/[^A-Z0-9]/g, "");
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
  const text = normalizeText(value).replace(/,/g, "");
  if (!text || text === "-") return 0;
  const n = parseFloat(text);
  return Number.isFinite(n) ? n : 0;
}

function roundTo2(value) {
  return Math.round((Number(value) + Number.EPSILON) * 100) / 100;
}

function isWithinTolerance(diff, tolerance) {
  return roundTo2(Math.abs(diff)) <= roundTo2(tolerance) + Number.EPSILON;
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
