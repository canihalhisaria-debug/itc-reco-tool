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

const MATCH_SETTINGS = {
  taxableTolerance: 50,
  taxTolerance: 10,
  smartMatch: true,
  monthOnly: true,
};

const state = {
  raw2b: [],
  rawBooks: [],
  headers2b: [],
  headersBooks: [],
  mapped2b: {},
  mappedBooks: {},
  results: {
    Matched: [],
    "Missing in 2B": [],
    "Missing in Books": [],
    "Value Difference": [],
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

file2bInput.addEventListener("change", (e) => handleFile(e.target.files[0], "2b"));
fileBooksInput.addEventListener("change", (e) => handleFile(e.target.files[0], "books"));
reconcileBtn.addEventListener("click", reconcile);
exportBtn.addEventListener("click", exportCurrentTab);

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

    row.append(label, select);
    target.appendChild(row);
  });
}

function updateReconcileButtonState() {
  const canRun =
    state.raw2b.length > 0 &&
    state.rawBooks.length > 0 &&
    REQUIRED_FIELDS.every((f) => state.mapped2b[f]) &&
    REQUIRED_FIELDS.every((f) => state.mappedBooks[f]);
  reconcileBtn.disabled = !canRun;
}

function normalizeRow(row, mapping) {
  const normalized = {
    gstin: String(row[mapping["GSTIN"]] || "").trim().toUpperCase(),
    supplierName: String(row[mapping["Supplier Name"]] || "").trim().toUpperCase(),
    invoiceNo: String(row[mapping["Invoice No"]] || "").trim().toUpperCase(),
    invoiceDate: normalizeDate(row[mapping["Invoice Date"]]),
    taxableValue: toNumber(row[mapping["Taxable Value"]]),
    igst: toNumber(row[mapping["IGST"]]),
    cgst: toNumber(row[mapping["CGST"]]),
    sgst: toNumber(row[mapping["SGST"]]),
  };
  normalized.totalTax = normalized.igst + normalized.cgst + normalized.sgst;
  normalized.key = `${normalized.gstin}|${normalized.supplierName}|${normalized.invoiceNo}|${normalized.invoiceDate}`;
  return normalized;
}

function monthPart(date) {
  return date ? date.slice(0, 7) : "";
}

function invoiceSuffix(invoiceNo) {
  return invoiceNo.length <= 4 ? invoiceNo : invoiceNo.slice(-4);
}

function isSameMonth(dateA, dateB) {
  return Boolean(monthPart(dateA) && monthPart(dateA) === monthPart(dateB));
}

function isAmountMatch(bookRow, twoBRow) {
  const taxableDiff = Math.abs(bookRow.taxableValue - twoBRow.taxableValue);
  const taxDiff = Math.abs(bookRow.totalTax - twoBRow.totalTax);
  return {
    taxableDiff,
    taxDiff,
    withinTolerance:
      taxableDiff <= MATCH_SETTINGS.taxableTolerance && taxDiff <= MATCH_SETTINGS.taxTolerance,
  };
}

function findCandidate(bookRow, candidates) {
  const unusedCandidates = candidates.filter((candidate) => !candidate.used);
  const exact = unusedCandidates.find(
    (candidate) =>
      candidate.gstin === bookRow.gstin &&
      candidate.supplierName === bookRow.supplierName &&
      candidate.invoiceNo === bookRow.invoiceNo &&
      candidate.invoiceDate === bookRow.invoiceDate
  );
  if (exact) return { candidate: exact, matchType: "EXACT" };

  if (!MATCH_SETTINGS.smartMatch) return { candidate: null, matchType: "" };

  const monthMatch = unusedCandidates.find(
    (candidate) =>
      candidate.gstin === bookRow.gstin &&
      candidate.supplierName === bookRow.supplierName &&
      candidate.invoiceNo === bookRow.invoiceNo &&
      (!MATCH_SETTINGS.monthOnly || isSameMonth(candidate.invoiceDate, bookRow.invoiceDate))
  );
  if (monthMatch) return { candidate: monthMatch, matchType: "MONTH_MATCH" };

  const suffixMatch = unusedCandidates.find(
    (candidate) =>
      candidate.gstin === bookRow.gstin &&
      candidate.supplierName === bookRow.supplierName &&
      invoiceSuffix(candidate.invoiceNo) === invoiceSuffix(bookRow.invoiceNo) &&
      (!MATCH_SETTINGS.monthOnly || isSameMonth(candidate.invoiceDate, bookRow.invoiceDate))
  );
  if (suffixMatch) return { candidate: suffixMatch, matchType: "SUFFIX_MATCH" };

  return { candidate: null, matchType: "" };
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

function pad2(n) {
  return String(n).padStart(2, "0");
}

function toNumber(value) {
  const n = parseFloat(String(value ?? "").replace(/,/g, "").trim());
  return Number.isFinite(n) ? n : 0;
}

function reconcile() {
  const books = state.rawBooks.map((r) => normalizeRow(r, state.mappedBooks));
  const twoB = state.raw2b.map((r) => normalizeRow(r, state.mapped2b));

  const twoBMap = new Map();
  twoB.forEach((row, idx) => {
    const groupKey = `${row.gstin}|${row.supplierName}`;
    const arr = twoBMap.get(groupKey) || [];
    arr.push({ ...row, sourceIndex: idx, used: false });
    twoBMap.set(groupKey, arr);
  });

  const matched = [];
  const missingIn2B = [];
  const valueDiff = [];

  books.forEach((bookRow, index) => {
    const groupKey = `${bookRow.gstin}|${bookRow.supplierName}`;
    const candidates = twoBMap.get(groupKey) || [];
    const { candidate, matchType } = findCandidate(bookRow, candidates);

    if (!candidate) {
      missingIn2B.push({ Type: "Missing in 2B", ...flattenRow(bookRow, "Books", index) });
      return;
    }

    const { taxableDiff, taxDiff, withinTolerance } = isAmountMatch(bookRow, candidate);
    candidate.used = true;

    if (withinTolerance) {
      matched.push({
        GSTIN: bookRow.gstin,
        SupplierName: bookRow.supplierName,
        InvoiceNo: bookRow.invoiceNo,
        InvoiceDate: bookRow.invoiceDate,
        MatchType: matchType,
        BooksTaxable: bookRow.taxableValue,
        TwoBTaxable: candidate.taxableValue,
        "Taxable Difference": taxableDiff,
        BooksTotalTax: bookRow.totalTax,
        TwoBTotalTax: candidate.totalTax,
        "Tax Difference": taxDiff,
      });
    } else {
      valueDiff.push({
        GSTIN: bookRow.gstin,
        SupplierName: bookRow.supplierName,
        InvoiceNo: bookRow.invoiceNo,
        InvoiceDate: bookRow.invoiceDate,
        MatchType: matchType,
        BooksTaxable: bookRow.taxableValue,
        TwoBTaxable: candidate.taxableValue,
        "Taxable Difference": taxableDiff,
        BooksTotalTax: bookRow.totalTax,
        TwoBTotalTax: candidate.totalTax,
        "Tax Difference": taxDiff,
      });
    }
  });

  const missingInBooks = [];
  twoBMap.forEach((rows) => {
    rows.forEach((r, index) => {
      if (!r.used) {
        missingInBooks.push({ Type: "Missing in Books", ...flattenRow(r, "2B", index) });
      }
    });
  });

  state.results = {
    Matched: matched,
    "Missing in 2B": missingIn2B,
    "Missing in Books": missingInBooks,
    "Value Difference": valueDiff,
  };

  document.getElementById("totalBooks").textContent = books.length;
  document.getElementById("total2b").textContent = twoB.length;
  document.getElementById("matchedCount").textContent = matched.length;
  document.getElementById("missing2bCount").textContent = missingIn2B.length;
  document.getElementById("missingBooksCount").textContent = missingInBooks.length;

  renderTabs();
  state.activeTab = "Matched";
  renderTable();
  exportBtn.disabled = false;
}

function flattenRow(row, source, index) {
  return {
    Source: source,
    RowNo: index + 1,
    GSTIN: row.gstin,
    SupplierName: row.supplierName,
    InvoiceNo: row.invoiceNo,
    InvoiceDate: row.invoiceDate,
    TaxableValue: row.taxableValue,
    IGST: row.igst,
    CGST: row.cgst,
    SGST: row.sgst,
    TotalTax: row.totalTax,
  };
}

function renderTabs() {
  tabButtons.innerHTML = "";
  Object.keys(state.results).forEach((tab) => {
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
  const thead = `<thead><tr>${columns.map((c) => `<th>${escapeHtml(c)}</th>`).join("")}</tr></thead>`;
  const tbody = `<tbody>${rows
    .map(
      (row) =>
        `<tr>${columns
          .map((col) => `<td>${escapeHtml(formatCell(row[col]))}</td>`)
          .join("")}</tr>`
    )
    .join("")}</tbody>`;

  resultTable.innerHTML = thead + tbody;
}

function formatCell(value) {
  if (typeof value === "number") return value.toFixed(2);
  return value ?? "";
}

function exportCurrentTab() {
  const rows = state.results[state.activeTab] || [];
  if (!rows.length) return;

  const ws = XLSX.utils.json_to_sheet(rows);
  const csv = XLSX.utils.sheet_to_csv(ws);
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = `${state.activeTab.replace(/\s+/g, "_")}.csv`;
  a.click();

  URL.revokeObjectURL(url);
}

function escapeHtml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
