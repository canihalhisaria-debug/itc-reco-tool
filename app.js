const REQUIRED_FIELDS = [
  "GSTIN",
  "Supplier Name",
  "Invoice No",
  "Invoice Date",
  "Taxable Value",
  "IGST",
  "CGST",
  "SGST",
];

const DEFAULT_SETTINGS = {
  taxableTolerance: 50,
  taxTolerance: 10,
  requireSameMonth: true,
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
const requireSameMonthInput = document.getElementById("requireSameMonth");

file2bInput.addEventListener("change", (e) => handleFile(e.target.files[0], "2b"));
fileBooksInput.addEventListener("change", (e) => handleFile(e.target.files[0], "books"));
reconcileBtn.addEventListener("click", reconcile);
exportBtn.addEventListener("click", exportResults);

[taxableToleranceInput, taxToleranceInput].forEach((input) => {
  input.addEventListener("change", syncSettingsFromUi);
  input.addEventListener("input", syncSettingsFromUi);
});
requireSameMonthInput.addEventListener("change", syncSettingsFromUi);

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
  REQUIRED_FIELDS.forEach((field) => {
    const row = document.createElement("div");
    row.className = "mapping-row";

    const label = document.createElement("label");
    label.textContent = field;

    const select = document.createElement("select");
    select.innerHTML = `<option value="">Select column...</option>${headers
      .map((h) => `<option value="${escapeHtml(h)}">${escapeHtml(h)}</option>`)
      .join("")}`;

    select.value = mapped[field] || "";
    select.addEventListener("change", () => {
      mapped[field] = select.value;
      updateReconcileButtonState();
    });

    row.appendChild(label);
    row.appendChild(select);
    target.appendChild(row);
  });
}

function updateReconcileButtonState() {
  const ready =
    state.raw2b.length > 0 &&
    state.rawBooks.length > 0 &&
    REQUIRED_FIELDS.every((f) => state.mapped2b[f]) &&
    REQUIRED_FIELDS.every((f) => state.mappedBooks[f]);

  reconcileBtn.disabled = !ready;
}

function syncSettingsToUi() {
  taxableToleranceInput.value = String(state.settings.taxableTolerance);
  taxToleranceInput.value = String(state.settings.taxTolerance);
  requireSameMonthInput.checked = state.settings.requireSameMonth;
}

function syncSettingsFromUi() {
  state.settings.taxableTolerance = toNumber(taxableToleranceInput.value);
  state.settings.taxTolerance = toNumber(taxToleranceInput.value);
  state.settings.requireSameMonth = Boolean(requireSameMonthInput.checked);
}

function normalizeRow(row, mapping, sourceIndex) {
  const normalized = {
    original: { ...row },
    sourceIndex,
    gstin: String(row[mapping["GSTIN"]] || "").trim().toUpperCase(),
    supplierName: String(row[mapping["Supplier Name"]] || "").trim(),
    invoiceNo: String(row[mapping["Invoice No"]] || "").trim(),
    invoiceDate: normalizeDate(row[mapping["Invoice Date"]]),
    taxableValue: toNumber(row[mapping["Taxable Value"]]),
    igst: toNumber(row[mapping["IGST"]]),
    cgst: toNumber(row[mapping["CGST"]]),
    sgst: toNumber(row[mapping["SGST"]]),
    used: false,
  };

  normalized.totalGST = normalized.igst + normalized.cgst + normalized.sgst;
  normalized.invoiceMonth = monthPart(normalized.invoiceDate);
  return normalized;
}

function monthPart(date) {
  return date ? date.slice(0, 7) : "";
}

function diffSummary(prRow, twoBRow) {
  const taxableDiff = Math.abs(prRow.taxableValue - twoBRow.taxableValue);
  const gstDiff = Math.abs(prRow.totalGST - twoBRow.totalGST);
  return {
    taxableDiff,
    gstDiff,
    withinTolerance:
      taxableDiff <= state.settings.taxableTolerance && gstDiff <= state.settings.taxTolerance,
    combinedDiff: taxableDiff + gstDiff,
  };
}

function findCandidate(prRow, twoBRows) {
  const sameGstinRows = twoBRows.filter((row) => !row.used && row.gstin === prRow.gstin);
  if (!sameGstinRows.length) return null;

  const baseCandidates = state.settings.requireSameMonth
    ? sameGstinRows.filter((row) => row.invoiceMonth && row.invoiceMonth === prRow.invoiceMonth)
    : sameGstinRows;

  if (!baseCandidates.length) return null;

  const evaluated = baseCandidates.map((row) => ({ row, ...diffSummary(prRow, row) }));
  const withinTolerance = evaluated.filter((item) => item.withinTolerance);

  const pool = withinTolerance.length ? withinTolerance : evaluated;
  pool.sort((a, b) => a.combinedDiff - b.combinedDiff || a.sourceIndex - b.sourceIndex);
  return {
    ...pool[0],
    matchType: withinTolerance.length ? "MATCHED" : "VALUE_DIFFERENCE",
  };
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
    const outcome = buildOutcomeRow(prRow, candidate.row, remark, candidate.taxableDiff, candidate.gstDiff);

    if (candidate.matchType === "MATCHED") {
      matched.push(outcome);
    } else {
      valueDifference.push(outcome);
    }

    prExportRows.push(buildPrExportRow(prRow, candidate.row, remark, candidate.taxableDiff, candidate.gstDiff));
    twoBExportRows[candidate.row.sourceIndex] = build2BExportRow(
      candidate.row,
      prRow,
      remark,
      candidate.taxableDiff,
      candidate.gstDiff
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

function buildOutcomeRow(prRow, twoBRow, remark, taxableDiff = "", gstDiff = "") {
  return {
    Remark: remark,
    PR_GSTIN: prRow?.gstin || "",
    TwoB_GSTIN: twoBRow?.gstin || "",
    PR_InvoiceDate: prRow?.invoiceDate || "",
    TwoB_InvoiceDate: twoBRow?.invoiceDate || "",
    PR_Taxable: prRow?.taxableValue ?? "",
    TwoB_Taxable: twoBRow?.taxableValue ?? "",
    PR_TotalGST: prRow?.totalGST ?? "",
    TwoB_TotalGST: twoBRow?.totalGST ?? "",
    Taxable_Diff: taxableDiff,
    GST_Diff: gstDiff,
  };
}

function buildPrExportRow(prRow, twoBRow, remark, taxableDiff = "", gstDiff = "") {
  return {
    ...prRow.original,
    Remark: remark,
    "2B_Taxable": twoBRow?.taxableValue ?? "",
    "2B_TotalGST": twoBRow?.totalGST ?? "",
    Taxable_Diff: taxableDiff,
    GST_Diff: gstDiff,
  };
}

function build2BExportRow(twoBRow, prRow, remark, taxableDiff = "", gstDiff = "") {
  return {
    ...twoBRow.original,
    Remark: remark,
    PR_Taxable: prRow?.taxableValue ?? "",
    PR_TotalGST: prRow?.totalGST ?? "",
    Taxable_Diff: taxableDiff,
    GST_Diff: gstDiff,
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
      const remarkClass = getRemarkClass(row.Remark);
      return `<tr class="${remarkClass}">${columns
        .map((col) => `<td>${escapeHtml(formatCell(row[col]))}</td>`)
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

  const prSheet = XLSX.utils.json_to_sheet(prRows.length ? prRows : [{ Info: "No records found" }]);
  const twoBSheet = XLSX.utils.json_to_sheet(twoBRows.length ? twoBRows : [{ Info: "No records found" }]);

  XLSX.utils.book_append_sheet(workbook, prSheet, "Purchase Register");
  XLSX.utils.book_append_sheet(workbook, twoBSheet, "GSTR-2B");
  XLSX.writeFile(workbook, "reconciliation_results.xlsx");
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
    return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
  }

  return "";
}

function formatCell(value) {
  return typeof value === "number" ? value.toFixed(2) : value ?? "";
}

function toNumber(value) {
  const n = parseFloat(String(value ?? "").replace(/,/g, "").trim());
  return Number.isFinite(n) ? n : 0;
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
