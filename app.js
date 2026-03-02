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
    "Party Summary": [],
    PurchaseRegisterExport: [],
    GSTR2BExport: [],
  },
  activeTab: "Matched",
  partySummarySearch: "",
  onlyMismatchParties: false,
  resultSearch: "",
  groupByParty: true,
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
const partySearchInput = document.getElementById("partySearch");
const mismatchPartiesOnlyInput = document.getElementById("onlyMismatchParties");
const partyFilters = document.getElementById("partySummaryFilters");
const resultSearchInput = document.getElementById("resultSearch");
const groupByPartyToggle = document.getElementById("groupByPartyToggle");

file2bInput.addEventListener("change", (e) => handleFile(e.target.files[0], "2b"));
fileBooksInput.addEventListener("change", (e) => handleFile(e.target.files[0], "books"));
reconcileBtn.addEventListener("click", reconcile);
exportBtn.addEventListener("click", exportResults);

[taxableToleranceInput, taxToleranceInput].forEach((input) => {
  input.addEventListener("change", syncSettingsFromUi);
  input.addEventListener("input", syncSettingsFromUi);
});

partySearchInput.addEventListener("input", () => {
  state.partySummarySearch = normalizeText(partySearchInput.value).toLowerCase();
  if (state.activeTab === "Party Summary") renderTable();
});

mismatchPartiesOnlyInput.addEventListener("change", () => {
  state.onlyMismatchParties = mismatchPartiesOnlyInput.checked;
  if (state.activeTab === "Party Summary") renderTable();
});

resultSearchInput.addEventListener("input", () => {
  state.resultSearch = normalizeText(resultSearchInput.value).toLowerCase();
  renderTable();
});

groupByPartyToggle.addEventListener("change", () => {
  state.groupByParty = groupByPartyToggle.checked;
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
  const cess = mapping.CESS ? roundTo2(toNumber(getMapped(row, mapping, "CESS"))) : 0;
  const totalTaxMapped = roundTo2(toNumber(getMapped(row, mapping, "Total Tax")));
  const totalTax = computeTotalTax({ igst, cgst, sgst, cess, totalTaxMapped, hasTotalTaxColumn: Boolean(mapping["Total Tax"]) });
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
    computedCESS: cess,
    totalTax,
    computedTotalTax: totalTax,
    used: false,
  };

  return normalized;
}

function getMapped(row, mapping, key) {
  const column = mapping[key];
  return column ? row[column] : "";
}

function diffSummary(prRow, twoBRow) {
  const taxableDiff = roundTo2(prRow.taxableValue - twoBRow.taxableValue);
  const taxDiff = roundTo2(prRow.totalTax - twoBRow.totalTax);
  const cessDiff = roundTo2(prRow.cess - twoBRow.cess);
  const cessComparable = Boolean(state.mapped2b.CESS) && Boolean(state.mappedBooks.CESS);
  const taxableTolerance = roundTo2(state.settings.taxableTolerance);
  const taxTolerance = roundTo2(state.settings.taxTolerance);
  const taxableWithin = isWithinTolerance(taxableDiff, taxableTolerance);
  const taxWithin = isWithinTolerance(taxDiff, taxTolerance);
  const cessWithin = cessComparable ? isWithinTolerance(cessDiff, taxTolerance) : true;
  const withinTolerance = taxableWithin && taxWithin && cessWithin;
  const combinedDiff = roundTo2(Math.abs(taxableDiff) + Math.abs(taxDiff) + (cessComparable ? Math.abs(cessDiff) : 0));

  return {
    taxableDiff,
    taxDiff,
    cessDiff,
    cessComparable,
    taxableWithin,
    taxWithin,
    cessWithin,
    withinTolerance,
    combinedDiff,
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

function getInvoiceMatches(prRow, twoBRows) {
  if (!prRow.invoiceNoNorm) return [];
  return twoBRows.filter((row) => !row.used && row.invoiceNoNorm && row.invoiceNoNorm === prRow.invoiceNoNorm);
}

function getAmountBasedMatches(prRow, twoBRows) {
  return twoBRows.filter((row) => !row.used && !row.hasStrongInvoiceMatch);
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
  markStrongInvoiceRows(prRows, twoBRows);

  const matched = [];
  const valueDifference = [];
  const notIn2B = [];
  const notInPR = [];

  const prExportRows = [];
  const twoBExportRows = new Array(twoBRows.length);

  prRows.forEach((prRow) => {
    if (prRow.used) return;

    const invoiceMatches = getInvoiceMatches(prRow, twoBRows);
    const invoiceToleranceMatch = findAmountToleranceMatch(prRow, invoiceMatches);
    if (invoiceToleranceMatch) {
      prRow.used = true;
      invoiceToleranceMatch.row.used = true;
      const mode = "Invoice+Amount";
      const reason = buildMatchReason(mode, prRow, invoiceToleranceMatch.row);
      const outcome = buildOutcomeRow(
        prRow,
        invoiceToleranceMatch.row,
        REMARKS.MATCHED,
        invoiceToleranceMatch.taxableDiff,
        invoiceToleranceMatch.taxDiff,
        invoiceToleranceMatch.cessDiff,
        mode,
        reason
      );
      matched.push(outcome);
      prExportRows.push(
        buildPrExportRow(prRow, invoiceToleranceMatch.row, REMARKS.MATCHED, invoiceToleranceMatch.taxableDiff, invoiceToleranceMatch.taxDiff, invoiceToleranceMatch.cessDiff, mode, reason)
      );
      twoBExportRows[invoiceToleranceMatch.row.sourceIndex] = build2BExportRow(
        invoiceToleranceMatch.row,
        prRow,
        REMARKS.MATCHED,
        invoiceToleranceMatch.taxableDiff,
        invoiceToleranceMatch.taxDiff,
        invoiceToleranceMatch.cessDiff,
        mode,
        reason
      );
      return;
    }

    const amountToleranceMatch = findAmountToleranceMatch(prRow, getAmountBasedMatches(prRow, twoBRows));
    if (amountToleranceMatch) {
      prRow.used = true;
      amountToleranceMatch.row.used = true;
      const mode = "AmountBased";
      const reason = buildMatchReason(mode, prRow, amountToleranceMatch.row);
      const outcome = buildOutcomeRow(
        prRow,
        amountToleranceMatch.row,
        REMARKS.MATCHED,
        amountToleranceMatch.taxableDiff,
        amountToleranceMatch.taxDiff,
        amountToleranceMatch.cessDiff,
        mode,
        reason
      );
      matched.push(outcome);
      prExportRows.push(
        buildPrExportRow(prRow, amountToleranceMatch.row, REMARKS.MATCHED, amountToleranceMatch.taxableDiff, amountToleranceMatch.taxDiff, amountToleranceMatch.cessDiff, mode, reason)
      );
      twoBExportRows[amountToleranceMatch.row.sourceIndex] = build2BExportRow(
        amountToleranceMatch.row,
        prRow,
        REMARKS.MATCHED,
        amountToleranceMatch.taxableDiff,
        amountToleranceMatch.taxDiff,
        amountToleranceMatch.cessDiff,
        mode,
        reason
      );
      return;
    }

    const fallbackPool = invoiceMatches.length ? invoiceMatches : getAmountBasedMatches(prRow, twoBRows);
    const closest = findClosestDifference(prRow, fallbackPool);

    if (!closest) {
      const reason = buildNotFoundReason(prRow);
      notIn2B.push(buildOutcomeRow(prRow, null, REMARKS.NOT_IN_2B, "", "", "", "", reason));
      prExportRows.push(buildPrExportRow(prRow, null, REMARKS.NOT_IN_2B, "", "", "", "", reason));
      return;
    }

    prRow.used = true;
    closest.row.used = true;
    const mode = invoiceMatches.length ? "InvoiceOnly" : "AmountBased";
    const toleranceResult = diffSummary(prRow, closest.row);
    const remark = toleranceResult.withinTolerance ? REMARKS.MATCHED : REMARKS.VALUE_DIFFERENCE;
    const reason = toleranceResult.withinTolerance ? buildMatchReason(mode, prRow, closest.row) : buildDiffReason(prRow, closest.row, toleranceResult, mode);
    const bucket = remark === REMARKS.MATCHED ? matched : valueDifference;

    bucket.push(buildOutcomeRow(prRow, closest.row, remark, toleranceResult.taxableDiff, toleranceResult.taxDiff, toleranceResult.cessDiff, mode, reason));
    prExportRows.push(buildPrExportRow(prRow, closest.row, remark, toleranceResult.taxableDiff, toleranceResult.taxDiff, toleranceResult.cessDiff, mode, reason));
    twoBExportRows[closest.row.sourceIndex] = build2BExportRow(closest.row, prRow, remark, toleranceResult.taxableDiff, toleranceResult.taxDiff, toleranceResult.cessDiff, mode, reason);
  });

  twoBRows.forEach((twoBRow) => {
    if (twoBRow.used) return;
    const reason = buildNotFoundReason(null, twoBRow);
    notInPR.push(buildOutcomeRow(null, twoBRow, REMARKS.NOT_IN_PR, "", "", "", "", reason));
    twoBExportRows[twoBRow.sourceIndex] = build2BExportRow(twoBRow, null, REMARKS.NOT_IN_PR, "", "", "", "", reason);
  });

  state.results = {
    Matched: matched,
    "Value Difference": valueDifference,
    "Not in 2B": notIn2B,
    "Not in PR": notInPR,
    "Party Summary": buildPartySummaryRows(matched, valueDifference, notIn2B, notInPR),
    PurchaseRegisterExport: prExportRows.sort(compareBusinessExportRows),
    GSTR2BExport: twoBExportRows.filter(Boolean).sort(compareBusinessExportRows),
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
  const reasons = ["Matched within taxable/total tax tolerance"];
  if (mode === "AmountBased" && prRow.invoiceNoNorm !== twoBRow.invoiceNoNorm) {
    reasons.push("invoice mismatch accepted in AmountBased fallback");
  }
  if (prRow.gstin && twoBRow.gstin && prRow.gstin !== twoBRow.gstin) reasons.push("GSTIN mismatch ignored");
  if (prRow.dateParseError || twoBRow.dateParseError) {
    reasons.push("date parse fail");
  }
  return reasons.join("; ");
}

function buildDiffReason(prRow, twoBRow, diff, mode) {
  const reasons = [];
  if (!diff.taxableWithin) reasons.push(`taxable diff ${Math.abs(diff.taxableDiff).toFixed(2)} > tolerance`);
  if (!diff.taxWithin) reasons.push(`Total tax diff ${Math.abs(diff.taxDiff).toFixed(2)} > tolerance`);
  if (diff.cessComparable && !diff.cessWithin) reasons.push(`CESS diff ${Math.abs(diff.cessDiff).toFixed(2)} > tolerance`);
  if (prRow.invoiceNoNorm !== twoBRow.invoiceNoNorm) reasons.push("invoice mismatch");
  if (prRow.gstin && twoBRow.gstin && prRow.gstin !== twoBRow.gstin) reasons.push("GSTIN mismatch");
  if (prRow.hasInvoiceDate && twoBRow.hasInvoiceDate && prRow.invoiceDate !== twoBRow.invoiceDate) reasons.push("date mismatch");
  if (prRow.dateParseError || twoBRow.dateParseError) reasons.push("date parse fail");
  if (!reasons.length) reasons.push(`closest ${mode} candidate but outside tolerance`);
  return reasons.join("; ");
}

function buildNotFoundReason(prRow, twoBRow = null) {
  const row = prRow || twoBRow;
  const side = prRow ? "2B" : "PR";
  const reasons = [`No ${side} row available for matching`];
  if (row && row.dateParseError) reasons.push("date parse fail");
  return reasons.join("; ");
}

function buildOutcomeRow(prRow, twoBRow, remark, taxableDiff = "", gstDiff = "", cessDiff = "", mode = "", reason = "") {
  const gstinMismatch = Boolean(prRow?.gstin && twoBRow?.gstin && prRow.gstin !== twoBRow.gstin);
  return {
    Remark: remark,
    MatchMode: mode,
    GSTINMismatch: gstinMismatch,
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
    PR_CESS: prRow?.computedCESS ?? 0,
    TwoB_CESS: twoBRow?.computedCESS ?? 0,
    TaxableDiff: taxableDiff,
    TaxDiff: gstDiff,
    CessDiff: cessDiff,
  };
}

function buildPrExportRow(prRow, twoBRow, remark) {
  const base = prRow || twoBRow;
  return buildBusinessExportRow(base, remark);
}

function build2BExportRow(twoBRow, prRow, remark) {
  const base = twoBRow || prRow;
  return buildBusinessExportRow(base, remark);
}

function buildBusinessExportRow(baseRow, remark) {
  return {
    GSTIN: baseRow?.gstin || "",
    SupplierName: baseRow?.supplierName || "",
    InvoiceNo: baseRow?.invoiceNo || "",
    InvoiceDate: baseRow?.invoiceDate || "",
    TaxableValue: baseRow?.taxableValue ?? "",
    IGST: baseRow?.igst ?? "",
    CGST: baseRow?.cgst ?? "",
    SGST: baseRow?.sgst ?? "",
    CESS: baseRow?.computedCESS ?? 0,
    ComputedTotalTax: baseRow?.computedTotalTax ?? baseRow?.totalTax ?? "",
    Remark: remark,
  };
}

function renderTabs() {
  const tabs = ["Matched", "Value Difference", "Not in 2B", "Not in PR", "Party Summary"];
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
  partyFilters.hidden = state.activeTab !== "Party Summary";
  if (!rows.length) {
    resultTable.innerHTML = "<tr><td>No records found.</td></tr>";
    return;
  }

  const columns = Object.keys(rows[0]);
  const thead = `<thead><tr>${columns.map((col) => `<th class="${getColumnClass(col, true)}">${escapeHtml(col)}</th>`).join("")}</tr></thead>`;

  let bodyHtml = "";
  if (state.activeTab !== "Party Summary" && state.groupByParty) {
    const groups = groupRowsByParty(rows);
    bodyHtml = groups
      .map((group) => {
        const groupHeader = `<tr class="party-group-header"><td colspan="${columns.length}"><div class="party-group-head"><span class="party-id">${escapeHtml(group.partyLabel)}</span><span class="party-meta">Matched: ${group.counts.matched} | Not in 2B: ${group.counts.notIn2B} | Not in PR: ${group.counts.notInPR} | Value Difference: ${group.counts.valueDifference}</span><span class="party-totals">Taxable: ${group.totals.taxable.toFixed(2)} | Total Tax: ${group.totals.tax.toFixed(2)}</span></div></td></tr>`;
        const rowsHtml = group.rows
          .map((row) => `<tr>${columns.map((col) => renderCell(row, col)).join("")}</tr>`)
          .join("");
        return groupHeader + rowsHtml;
      })
      .join("");
  } else {
    bodyHtml = rows.map((row) => `<tr>${columns.map((col) => renderCell(row, col)).join("")}</tr>`).join("");
  }

  resultTable.innerHTML = thead + `<tbody>${bodyHtml}</tbody>`;
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
  if (["taxablevalue", "igst", "cgst", "sgst", "cess", "computedtotaltax", "prtaxablevalue", "twobtaxablevalue", "prtotaltax", "twobtotaltax", "taxablediff", "taxdiff", "cessdiff"].includes(lower.replace(/[^a-z0-9]/g, ""))) classes.push("col-num");
  if (lower.includes("supplier")) classes.push("col-supplier");
  if (lower.includes("invoice") && lower.includes("no")) classes.push("col-invoice");
  if (isHeader) classes.push("col-head");
  return classes.join(" ");
}

function groupRowsByParty(rows) {
  const grouped = new Map();
  rows.forEach((row) => {
    const key = getPartySortKeyFromDisplayRow(row);
    if (!grouped.has(key.key)) {
      grouped.set(key.key, {
        partyLabel: `${key.supplier || "Unknown Supplier"} (${key.gstin || "No GSTIN"})`,
        rows: [],
        counts: { matched: 0, notIn2B: 0, notInPR: 0, valueDifference: 0 },
        totals: { taxable: 0, tax: 0 },
      });
    }
    const grp = grouped.get(key.key);
    grp.rows.push(row);
    if (row.Remark === REMARKS.MATCHED) grp.counts.matched += 1;
    if (row.Remark === REMARKS.NOT_IN_2B) grp.counts.notIn2B += 1;
    if (row.Remark === REMARKS.NOT_IN_PR) grp.counts.notInPR += 1;
    if (row.Remark === REMARKS.VALUE_DIFFERENCE) grp.counts.valueDifference += 1;
    grp.totals.taxable = roundTo2(grp.totals.taxable + getRowTaxable(row));
    grp.totals.tax = roundTo2(grp.totals.tax + getRowTotalTax(row));
  });
  return Array.from(grouped.values());
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

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(prRows.length ? prRows : [{ Info: "No records found" }]), "PR_with_Remarks");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(twoBRows.length ? twoBRows : [{ Info: "No records found" }]), "2B_with_Remarks");
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.json_to_sheet(state.results["Party Summary"].length ? state.results["Party Summary"] : [{ Info: "No records found" }]),
    "Party_Summary"
  );
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(buildSummaryRows()), "Summary");
  XLSX.writeFile(workbook, "ITC_Reco_Output.xlsx");
}


function compareBusinessExportRows(a, b) {
  const supplierA = normalizeSupplierName(a.SupplierName);
  const supplierB = normalizeSupplierName(b.SupplierName);
  const keyA = supplierA || normalizeText(a.GSTIN).toUpperCase();
  const keyB = supplierB || normalizeText(b.GSTIN).toUpperCase();
  if (keyA !== keyB) return keyA.localeCompare(keyB);
  const dateA = normalizeText(a.InvoiceDate || "9999-99-99");
  const dateB = normalizeText(b.InvoiceDate || "9999-99-99");
  if (dateA !== dateB) return dateA.localeCompare(dateB);
  return normalizeInvoiceNo(a.InvoiceNo).localeCompare(normalizeInvoiceNo(b.InvoiceNo));
}

function buildSummaryRows() {
  return [
    { Metric: "Total Books Invoices", Value: document.getElementById("totalBooks").textContent },
    { Metric: "Total 2B Invoices", Value: document.getElementById("total2b").textContent },
    { Metric: "Matched", Value: document.getElementById("matchedCount").textContent },
    { Metric: "Not in 2B", Value: document.getElementById("missing2bCount").textContent },
    { Metric: "Not in PR", Value: document.getElementById("missingBooksCount").textContent },
    { Metric: "Value Difference", Value: document.getElementById("valueDiffCount").textContent },
    { Metric: "Taxable Tolerance", Value: String(state.settings.taxableTolerance) },
    { Metric: "Tax Tolerance", Value: String(state.settings.taxTolerance) },
  ];
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
    .replace(/[\\/_]+/g, "-")
    .replace(/\s*[-.]+\s*/g, "-")
    .replace(/-+/g, "-")
    .replace(/[^A-Z0-9-]/g, "")
    .replace(/^-|-$/g, "");
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
  const igst = toNumber(row.igst);
  const cgst = toNumber(row.cgst);
  const sgst = toNumber(row.sgst);
  const cess = toNumber(row.cess);
  const componentSum = roundTo2(igst + cgst + sgst + cess);
  if (row.hasTotalTaxColumn) {
    const totalTaxMapped = toNumber(row.totalTaxMapped);
    if (Math.abs(totalTaxMapped - componentSum) <= 0.01) return componentSum;
  }
  return componentSum;
}

function markStrongInvoiceRows(prRows, twoBRows) {
  const prSet = new Set(prRows.filter((row) => row.invoiceNoNorm).map((row) => row.invoiceNoNorm));
  const b2Set = new Set(twoBRows.filter((row) => row.invoiceNoNorm).map((row) => row.invoiceNoNorm));
  prRows.forEach((row) => {
    row.hasStrongInvoiceMatch = Boolean(row.invoiceNoNorm && b2Set.has(row.invoiceNoNorm));
  });
  twoBRows.forEach((row) => {
    row.hasStrongInvoiceMatch = Boolean(row.invoiceNoNorm && prSet.has(row.invoiceNoNorm));
  });
}

function getActiveRows() {
  if (state.activeTab === "Party Summary") {
    return state.results["Party Summary"].filter((row) => {
      const matchesSearch = !state.partySummarySearch || String(row.PartyKey).toLowerCase().includes(state.partySummarySearch);
      const mismatchTotal =
        Math.abs(row.NotInPR_Taxable) +
        Math.abs(row.NotIn2B_Taxable) +
        Math.abs(row.ValueDiffImpact_Taxable) +
        Math.abs(row.NotInPR_TotalTax) +
        Math.abs(row.NotIn2B_TotalTax) +
        Math.abs(row.ValueDiffImpact_TotalTax);
      return matchesSearch && (!state.onlyMismatchParties || mismatchTotal > 0);
    });
  }

  const activeRows = (state.results[state.activeTab] || []).filter((row) => {
    if (!state.resultSearch) return true;
    const haystack = [row.PR_SupplierName, row.TwoB_SupplierName, row.PR_GSTIN, row.TwoB_GSTIN, row.PR_InvoiceNo, row.TwoB_InvoiceNo]
      .map((v) => String(v || "").toLowerCase())
      .join(" ");
    return haystack.includes(state.resultSearch);
  });

  return activeRows.sort(compareDisplayRows);
}

function getPartySortKeyFromDisplayRow(row) {
  const supplier = normalizeText(row.PR_SupplierName || row.TwoB_SupplierName || "");
  const gstin = normalizeText(row.PR_GSTIN || row.TwoB_GSTIN || "").toUpperCase();
  const sortSupplier = normalizeSupplierName(supplier);
  const sortKey = sortSupplier || gstin || "zzzz";
  return { key: `${sortKey}|${gstin}`, supplier, gstin, sortKey };
}

function compareDisplayRows(a, b) {
  const aParty = getPartySortKeyFromDisplayRow(a);
  const bParty = getPartySortKeyFromDisplayRow(b);
  if (aParty.sortKey !== bParty.sortKey) return aParty.sortKey.localeCompare(bParty.sortKey);
  if (aParty.gstin !== bParty.gstin) return aParty.gstin.localeCompare(bParty.gstin);
  const aDate = normalizeText(a.PR_InvoiceDate || a.TwoB_InvoiceDate || "9999-99-99");
  const bDate = normalizeText(b.PR_InvoiceDate || b.TwoB_InvoiceDate || "9999-99-99");
  if (aDate !== bDate) return aDate.localeCompare(bDate);
  const aInv = normalizeInvoiceNo(a.PR_InvoiceNo || a.TwoB_InvoiceNo || "");
  const bInv = normalizeInvoiceNo(b.PR_InvoiceNo || b.TwoB_InvoiceNo || "");
  return aInv.localeCompare(bInv);
}

function getRowTaxable(row) {
  return toNumber(row.PR_TaxableValue || row.TwoB_TaxableValue || row.TaxableValue);
}

function getRowTotalTax(row) {
  return toNumber(row.PR_TotalTax || row.TwoB_TotalTax || row.ComputedTotalTax);
}

function buildPartySummaryRows(matchedRows, valueDiffRows, notIn2BRows, notInPRRows) {
  const partyMap = new Map();
  const ensure = (row) => {
    const gstin = row.TwoB_GSTIN || row.PR_GSTIN || "";
    const supplier = row.TwoB_SupplierName || row.PR_SupplierName || "Unknown Supplier";
    const partyKey = gstin || supplier;
    if (!partyMap.has(partyKey)) {
      partyMap.set(partyKey, {
        PartyKey: partyKey,
        GSTIN: gstin,
        SupplierName: supplier,
        AsPer2B_Taxable: 0,
        AsPer2B_TotalTax: 0,
        AsPer2B_CESS: 0,
        NotInPR_Taxable: 0,
        NotInPR_TotalTax: 0,
        NotInPR_CESS: 0,
        NotIn2B_Taxable: 0,
        NotIn2B_TotalTax: 0,
        NotIn2B_CESS: 0,
        ValueDiffImpact_Taxable: 0,
        ValueDiffImpact_TotalTax: 0,
        ValueDiffImpact_CESS: 0,
      });
    }
    return partyMap.get(partyKey);
  };

  [...matchedRows, ...valueDiffRows, ...notInPRRows].forEach((row) => {
    const party = ensure(row);
    party.AsPer2B_Taxable = roundTo2(party.AsPer2B_Taxable + toNumber(row.TwoB_TaxableValue));
    party.AsPer2B_TotalTax = roundTo2(party.AsPer2B_TotalTax + toNumber(row.TwoB_TotalTax));
    party.AsPer2B_CESS = roundTo2(party.AsPer2B_CESS + toNumber(row.TwoB_CESS));
  });

  notInPRRows.forEach((row) => {
    const party = ensure(row);
    party.NotInPR_Taxable = roundTo2(party.NotInPR_Taxable + toNumber(row.TwoB_TaxableValue));
    party.NotInPR_TotalTax = roundTo2(party.NotInPR_TotalTax + toNumber(row.TwoB_TotalTax));
    party.NotInPR_CESS = roundTo2(party.NotInPR_CESS + toNumber(row.TwoB_CESS));
  });

  notIn2BRows.forEach((row) => {
    const party = ensure(row);
    party.NotIn2B_Taxable = roundTo2(party.NotIn2B_Taxable + toNumber(row.PR_TaxableValue));
    party.NotIn2B_TotalTax = roundTo2(party.NotIn2B_TotalTax + toNumber(row.PR_TotalTax));
    party.NotIn2B_CESS = roundTo2(party.NotIn2B_CESS + toNumber(row.PR_CESS));
  });

  valueDiffRows.forEach((row) => {
    const party = ensure(row);
    party.ValueDiffImpact_Taxable = roundTo2(party.ValueDiffImpact_Taxable + toNumber(row.TaxableDiff));
    party.ValueDiffImpact_TotalTax = roundTo2(party.ValueDiffImpact_TotalTax + toNumber(row.TaxDiff));
    party.ValueDiffImpact_CESS = roundTo2(party.ValueDiffImpact_CESS + toNumber(row.CessDiff));
  });

  return Array.from(partyMap.values())
    .map((party) => ({
      ...party,
      DerivedPR_Taxable: roundTo2(party.AsPer2B_Taxable - party.NotInPR_Taxable + party.NotIn2B_Taxable + party.ValueDiffImpact_Taxable),
      DerivedPR_TotalTax: roundTo2(party.AsPer2B_TotalTax - party.NotInPR_TotalTax + party.NotIn2B_TotalTax + party.ValueDiffImpact_TotalTax),
      DerivedPR_CESS: roundTo2(party.AsPer2B_CESS - party.NotInPR_CESS + party.NotIn2B_CESS + party.ValueDiffImpact_CESS),
    }))
    .sort((a, b) => a.SupplierName.localeCompare(b.SupplierName));
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
