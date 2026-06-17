// ─── COLUMN MAPPER ──────────────────────────────────────────────────────────
// Converts raw Excel rows (any of the supported source formats) into the
// dashboard's internal PO record shape. Two formats are recognised:
//
//   1. "Real" PP Auto format (e.g. PO_Data_2026-27.xlsx) — line-item level,
//      columns like "P.O Date.", "P.O No.", "GST Type" (text like
//      "SGST 9%-CGST 9%"), "Status of PO" (Open/Closed/PO Cancelled), etc.
//   2. The dashboard's own simplified mock format — poDate, poNo, supplier,
//      item, qty, rate, status (Complete/Partial/Pending), etc.
//
// detectFormat() looks at the header keys present and picks the right mapper.

// Trims whitespace from every key in a row object (the real export has
// headers like " Taxable Amt " with leading/trailing spaces).
function normalizeKeys(row) {
  const out = {};
  for (const k of Object.keys(row)) {
    out[k.trim()] = row[k];
  }
  return out;
}

function excelDateToISO(value) {
  if (value == null || value === "") return "";
  // Already a JS Date (xlsx with cellDates) or a parseable date string
  if (value instanceof Date && !isNaN(value)) {
    return value.toISOString().slice(0, 10);
  }
  if (typeof value === "string") {
    const d = new Date(value);
    if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
    return "";
  }
  if (typeof value === "number") {
    // Excel serial date (days since 1899-12-30)
    const d = new Date(Math.round((value - 25569) * 86400 * 1000));
    if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
  }
  return "";
}

function parseGstType(raw) {
  // Examples: "SGST 9%-CGST 9%", "IGST 18%", "IGST 5%", "GST Free", "" / undefined
  if (!raw || typeof raw !== "string") return { label: "N/A", pct: null };
  const cleaned = raw.trim();
  if (/free/i.test(cleaned)) return { label: "GST Free", pct: 0 };
  const nums = cleaned.match(/(\d+(\.\d+)?)/g);
  if (!nums) return { label: cleaned, pct: null };
  if (/igst/i.test(cleaned)) {
    return { label: "IGST" + nums[0] + "%", pct: parseFloat(nums[0]) };
  }
  // SGST + CGST split — total is sum of both halves
  const total = nums.reduce((s, n) => s + parseFloat(n), 0);
  return { label: "GST" + Math.round(total), pct: total };
}

function mapStatus(rawStatus, qty, receivedQty) {
  if (!rawStatus) return "N/A";
  const s = String(rawStatus).trim().toLowerCase();
  if (s === "po cancelled" || s === "cancelled") return "Cancelled";
  if (s === "closed") return "Complete";
  if (s === "open") {
    const q = Number(qty) || 0;
    const r = Number(receivedQty) || 0;
    if (r <= 0) return "Pending";
    if (r < q) return "Partial";
    return "Complete";
  }
  // Already in dashboard's own vocabulary
  if (["complete", "partial", "pending"].includes(s)) {
    return s.charAt(0).toUpperCase() + s.slice(1);
  }
  return "N/A";
}

// Detects which known header set a sheet's first row matches.
function detectFormat(row) {
  if (!row) return "unknown";
  const keys = Object.keys(row);
  if (keys.includes("P.O No.") || keys.includes("Status of PO") || keys.includes("GST Type")) {
    return "real";
  }
  if (keys.includes("poNo") || keys.includes("status")) {
    return "mock";
  }
  return "unknown";
}

// Maps one row of the "real" PP Auto export format.
function mapRealRow(rawRow, idx) {
  const row = normalizeKeys(rawRow);
  const supplier = (row["Supplier Name"] || "").toString().trim();
  const item = (row["Items"] || "").toString().trim();
  // Skip fully blank line items (e.g. cancelled "Blank PO" placeholder rows)
  if (!supplier && !item) return null;

  const qty = Number(row["QTY"]) || 0;
  const rate = Number(row["Rate"]) || 0;
  const receivedQty = Number(row["Received QTY"]) || 0;
  const gst = parseGstType(row["GST Type"]);

  return {
    _rowId: idx,
    poDate: excelDateToISO(row["P.O Date."]),
    poNo: "PO-" + row["P.O No."],
    supplier,
    item,
    qty,
    rate,
    deliveryType: (row["Delivery Type"] && row["Delivery Type"] !== 0) ? row["Delivery Type"] : "Regular",
    gstType: gst.label,
    gstPct: gst.pct,
    taxableAmt: Number(row["Taxable Amt"]) || 0,
    totalAmt: Number(row["Total Rate With GST"]) || 0,
    chargesType: row["Charges Type"] || "",
    chargesAmt: Number(row["Charges AMT"]) || 0,
    description: row["Description"] || "",
    preparedBy: row["Prepared By"] || "N/A",
    requestedBy: row["Requested By"] || "N/A",
    // Not yet captured in the source process — placeholders for future use.
    authorisedBy: row["Authorised by"] || "N/A",
    sentToParty: row["Sent to Party"] != null ? row["Sent to Party"] : "N/A",
    receivedQty,
    shortExcessQty: row["Short & Excess QTY"] != null ? Number(row["Short & Excess QTY"]) : (receivedQty - qty),
    status: mapStatus(row["Status of PO"], qty, receivedQty),
    remarks: row["Remarks of PO"] || "",
    internalRemarks: row["Internal Remarks"] || "",
  };
}

// Maps one row of the dashboard's own mock/simple format (pass-through with
// light normalisation so both formats end up with identical field names).
function mapMockRow(row, idx) {
  return {
    _rowId: idx,
    poDate: row.poDate || "",
    poNo: row.poNo || "",
    supplier: row.supplier || "",
    item: row.item || "",
    qty: Number(row.qty) || 0,
    rate: Number(row.rate) || 0,
    deliveryType: row.deliveryType || "Regular",
    gstType: row.gstType || "N/A",
    gstPct: null,
    taxableAmt: Number(row.taxableAmt) || 0,
    totalAmt: Number(row.totalAmt) || 0,
    chargesType: "",
    chargesAmt: 0,
    description: "",
    preparedBy: row.preparedBy || "N/A",
    requestedBy: row.requestedBy || "N/A",
    authorisedBy: row.authorisedBy || "N/A",
    sentToParty: row.sentToParty != null ? row.sentToParty : "N/A",
    receivedQty: Number(row.receivedQty) || 0,
    shortExcessQty: (Number(row.receivedQty) || 0) - (Number(row.qty) || 0),
    status: row.status || "N/A",
    remarks: row.remarks || "",
    internalRemarks: row.internalRemarks || "",
  };
}

// ─── RECEIPTS MAPPER ────────────────────────────────────────────────────────
// Handles both the dashboard's own simple Receipts sheet (date, poNo,
// voucherNo, supplier, item, receivedQty, rate) and the real export's
// "Recvd_Item" sheet, which comes from a pivot table and has irregular
// headers (a stray numeric '0' key, 'Vou' for voucher, '__EMPTY' junk column).
export function mapReceiptRows(rawRows) {
  if (!Array.isArray(rawRows) || rawRows.length === 0) return [];
  return rawRows
    .map((rawRow, idx) => {
      const row = normalizeKeys(rawRow);
      const isReal = "PO No." in row || "Vou" in row || "Recvd QTY" in row;
      if (isReal) {
        const poNo = row["PO No."];
        if (poNo == null) return null;
        return {
          _rowId: idx,
          date: excelDateToISO(row["Date"]),
          poNo: "PO-" + poNo,
          voucherNo: row["Vou"] != null ? "GRN-" + row["Vou"] : "N/A",
          supplier: (row["Supplier"] || "").toString().trim(),
          item: (row["Item"] || "").toString().trim(),
          receivedQty: Number(row["Recvd QTY"]) || 0,
          rate: 0, // not present in this sheet; left for cross-reference with PO data if needed
        };
      }
      // Dashboard's own mock shape
      if (!row.poNo && !row.supplier) return null;
      return {
        _rowId: idx,
        date: row.date || "",
        poNo: row.poNo || "",
        voucherNo: row.voucherNo || "N/A",
        supplier: row.supplier || "",
        item: row.item || "",
        receivedQty: Number(row.receivedQty) || 0,
        rate: Number(row.rate) || 0,
      };
    })
    .filter(Boolean);
}

// ─── RATE HISTORY MAPPER ────────────────────────────────────────────────────
// Real export's "Rate Update" sheet already matches the dashboard's mock
// shape almost exactly (Date, Supplier, Item, New Rate, Old Rate) — just
// needs date normalisation and key renaming.
export function mapRateRows(rawRows) {
  if (!Array.isArray(rawRows) || rawRows.length === 0) return [];
  return rawRows
    .map((rawRow, idx) => {
      const row = normalizeKeys(rawRow);
      const isReal = "New Rate" in row || "Old Rate" in row;
      if (isReal) {
        if (!row["Supplier"] && !row["Item"]) return null;
        return {
          _rowId: idx,
          date: excelDateToISO(row["Date"]),
          supplier: (row["Supplier"] || "").toString().trim(),
          item: (row["Item"] || "").toString().trim(),
          newRate: Number(row["New Rate"]) || 0,
          oldRate: Number(row["Old Rate"]) || 0,
        };
      }
      if (!row.supplier && !row.item) return null;
      return {
        _rowId: idx,
        date: row.date || "",
        supplier: row.supplier || "",
        item: row.item || "",
        newRate: Number(row.newRate) || 0,
        oldRate: Number(row.oldRate) || 0,
      };
    })
    .filter(Boolean);
}

// Public entry point: takes raw sheet_to_json() rows, returns normalised
// dashboard-shaped PO records. Drops unparseable/fully-blank rows.
export function mapPORows(rawRows) {
  if (!Array.isArray(rawRows) || rawRows.length === 0) return [];
  const format = detectFormat(normalizeKeys(rawRows[0]));
  const mapper = format === "real" ? mapRealRow : mapMockRow;
  return rawRows
    .map((row, i) => mapper(row, i))
    .filter(Boolean);
}

export { detectFormat, parseGstType, mapStatus, excelDateToISO };
