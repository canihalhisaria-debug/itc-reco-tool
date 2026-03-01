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
};

const MATCH_TYPES = {
  MATCHED_AUTO: "MATCHED_AUTO",
  MATCH_REVIEW: "MATCH_REVIEW",
  NOT_IN_2B: "NOT_IN_2B",
  NOT_IN_PR: "NOT_IN_PR",
};

const REMARKS = {
  [MATCH_TYPES.MATCHED_AUTO]: "Auto matched by GSTIN + Amount",
  [MATCH_TYPES.MATCH_REVIEW]: "Review - amount close",
  [MATCH_TYPES.NOT_IN_2B]: "Not reflecting in 2B",
  [MATCH_TYPES.NOT_IN_PR]: "Not recorded in Purchase Register",
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
    "Not in 2B": [],
    "Not in PR": [],
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
  normalized.invoiceMonth = monthPart(normalized.invoiceDate);
  return normalized;
}

function monthPart(date) {
  return date ? date.slice(0, 7) : "";
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

function findBestCandidate(bookRow, candidates) {
  const sameMonthCandidates = candidates.filter(
    (candidate) => !candidate.used && candidate.gstin === bookRow.gstin && isSameMonth(candidate.invoiceDate, bookRow.invoiceDate)
  );

  if (!sameMonthCandidates.length) {
    return { candidate: null, taxableDiff: 0, taxDiff: 0, withinTolerance: false };
  }

  let bestCandidate = null;
  let bestDiff = Number.POSITIVE_INFINITY;
  let bestAmountMatch = null;

  sameMonthCandidates.forEach((candidate) => {
    const amountMatch = isAmountMatch(bookRow, candidate);
    const combinedDiff = amountMatch.taxableDiff + amountMatch.taxDiff;
    if (combinedDiff < bestDiff) {
      bestDiff = combinedDiff;
      bestCandidate = candidate;
      bestAmountMatch = amountMatch;
    }
  });

  return { candidate: bestCandidate, ...bestAmountMatch };
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
    const groupKey = row.gstin;
    const arr = twoBMap.get(groupKey) || [];
    arr.push({ ...row, sourceIndex: idx, used: false });
    twoBMap.set(groupKey, arr);
  });

  const matched = [];
  const notIn2B = [];

  books.forEach((bookRow, index) => {
    const groupKey = bookRow.gstin;
    const candidates = twoBMap.get(groupKey) || [];
    const { candidate, taxableDiff, taxDiff, withinTolerance } = findBestCandidate(bookRow, candidates);

    if (!candidate) {
      notIn2B.push({
        ...flattenRow(bookRow, "Books", index),
        MatchType: MATCH_TYPES.NOT_IN_2B,
        Remark: REMARKS[MATCH_TYPES.NOT_IN_2B],
      });
      return;
    }
    candidate.used = true;


    const matchType = withinTolerance ? MATCH_TYPES.MATCHED_AUTO : MATCH_TYPES.MATCH_REVIEW;
    matched.push({
      GSTIN: bookRow.gstin,
      BooksSupplierName: bookRow.supplierName,
      TwoBSupplierName: candidate.supplierName,
      BooksInvoiceNo: bookRow.invoiceNo,
      TwoBInvoiceNo: candidate.invoiceNo,
      BooksInvoiceDate: bookRow.invoiceDate,
      TwoBInvoiceDate: candidate.invoiceDate,
      InvoiceMonth: bookRow.invoiceMonth,
      MatchType: matchType,
      Remark: REMARKS[matchType],
      BooksTaxable: bookRow.taxableValue,
      TwoBTaxable: candidate.taxableValue,
      "Taxable Difference": taxableDiff,
      BooksTotalTax: bookRow.totalTax,
      TwoBTotalTax: candidate.totalTax,
      "Tax Difference": taxDiff,
    });
  });

  const notInPR = [];
  twoBMap.forEach((rows) => {
    rows.forEach((r, index) => {
      if (!r.used) {
        notInPR.push({
          ...flattenRow(r, "2B", index),
          MatchType: MATCH_TYPES.NOT_IN_PR,
          Remark: REMARKS[MATCH_TYPES.NOT_IN_PR],
        });
      }
    });
  });

  state.results = {
    Matched: matched,
    "Not in 2B": notIn2B,
    "Not in PR": notInPR,
  };

  document.getElementById("totalBooks").textContent = books.length;
  document.getElementById("total2b").textContent = twoB.length;
  document.getElementById("matchedCount").textContent = matched.length;
  document.getElementById("missing2bCount").textContent = notIn2B.length;
  document.getElementById("missingBooksCount").textContent = notInPR.length;

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
  const workbook = XLSX.utils.book_new();
  const sheetsToExport = ["Matched", "Not in 2B", "Not in PR"];

  sheetsToExport.forEach((sheetName) => {
    const rows = state.results[sheetName] || [];
    const ws = XLSX.utils.json_to_sheet(rows.length ? rows : [{ Info: "No records found" }]);
    XLSX.utils.book_append_sheet(workbook, ws, sheetName);
  });

  XLSX.writeFile(workbook, "reconciliation_results.xlsx");
}

function escapeHtml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
