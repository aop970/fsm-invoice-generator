import * as XLSX from 'xlsx';
import type { InvoiceSummary, WeeklySummary, InvoiceRow, MgmtRow } from './types';
import { MGMT_TABLE } from './constants';
import { buildCoverPeriodStr, getWeekEndDate } from './transform';

// ── Formatting helpers ─────────────────────────────────────────────────────

function fmt(n: number): string {
  return n.toFixed(2);
}

function pctFmt(n: number): string {
  return (n * 100).toFixed(2) + '%';
}

// ── Labor tab builder (FSM I / FSM II) ────────────────────────────────────
//
// Reference column structure (20 cols for FSM I, 19 cols for FSM II):
//   A  Week
//   B  Employee Name
//   C  Associate ID
//   D  Associate Type
//   E  Employee Region
//   F  Employee District
//   G  Employee Market
//   H  Store ID
//   I  Store
//   J  Associate State
//   K  Store State
//   L  Zip Code
//   M  Visit Date
//   N  Time Hours
//   O  Base Pay Rate
//   P  MU
//   Q  Pay Rate Total
//   R  Bill
//   S  Comments
//   T  (empty — FSM I only)
//
// Row 1 (metadata):  C=count, N=total OT hours, O=OT rate label, R=total Bill
// Row 2 (headers)
// Row 3+ (data)

function buildLaborTab(
  rows: InvoiceRow[],
  tabTotal: number,
  otLabel: string,        // e.g. "OT-15.53" or "OT-17.75"
  extraEmptyCol: boolean, // true = FSM I (20 cols), false = FSM II (19 cols)
): unknown[][] {
  const numCols = extraEmptyCol ? 20 : 19;

  // Count of unique Associate IDs (mirrors the COUNTA(UNIQUE(FILTER(...))) formula in the reference)
  const uniqueAssociateCount = new Set(rows.map(r => r.associateId.trim()).filter(Boolean)).size;

  // Total OT hours = sum of rows where Base Pay Rate is the OT rate
  const otHours = rows
    .filter(r => r.basePayRate < 32) // OT rates are 15.53 and 17.75 (both < 32)
    .reduce((s, r) => s + r.timeHours, 0);

  // Row 1: metadata
  const metaRow: unknown[] = new Array(numCols).fill('');
  metaRow[2]  = uniqueAssociateCount; // col C — unique associate count
  metaRow[13] = otHours;              // col N — total OT hours
  metaRow[14] = otLabel;              // col O — OT rate label string
  metaRow[17] = tabTotal;             // col R — total Bill

  // Row 2: headers
  const headerRow: unknown[] = [
    'Week',             // A
    'Employee Name',    // B
    'Associate ID',     // C
    'Associate Type',   // D
    'Employee Region',  // E
    'Employee District',// F
    'Employee Market',  // G
    'Store ID',         // H
    'Store',            // I
    'Associate State',  // J
    'Store State',      // K
    'Zip Code',         // L
    'Visit Date',       // M
    'Time Hours',       // N
    'Base Pay Rate',    // O
    'MU',               // P
    'Pay Rate Total',   // Q
    'Bill',             // R
    'Comments',         // S
  ];
  if (extraEmptyCol) headerRow.push(''); // T (empty)

  const dataRows = rows.map(r => {
    const row: unknown[] = [
      r.week,             // A
      r.employeeName,     // B
      r.associateId,      // C
      r.associateType,    // D
      r.region,           // E
      r.district,         // F
      r.market,           // G
      r.storeId,          // H
      r.storeName,        // I
      r.associateState,   // J
      r.storeState,       // K
      r.zipCode,          // L
      r.visitDate,        // M
      r.timeHours,        // N
      fmt(r.basePayRate), // O
      fmt(r.mu),          // P
      fmt(r.payRateTotal),// Q
      fmt(r.bill),        // R
      r.comments,         // S
    ];
    if (extraEmptyCol) row.push(''); // T
    return row;
  });

  return [metaRow, headerRow, ...dataRows];
}

// ── Management Detail tab builder ──────────────────────────────────────────
//
// Reference column structure (13 cols):
//   A  Week
//   B  Associate Name
//   C  Associate ID
//   D  Title
//   E  Associate State
//   F  Hours
//   G  Hourly Rate
//   H  Total
//   I  % Allocation
//   J  Total Bill
//   K  Comments
//   L  (empty)
//   M  (empty)
//
// Row 1 (metadata): C=count of data rows, J=total of all Total Bill values
// Row 2 (headers)
// Row 3+ (data)

function buildMgmtTab(rows: MgmtRow[]): unknown[][] {
  const numCols = 13;

  const grandTotalBill = rows.reduce((s, r) => s + r.totalBill, 0);

  // Row 1: metadata
  const metaRow: unknown[] = new Array(numCols).fill('');
  metaRow[2]  = rows.length; // col C — count of data rows (28 per week × number of weeks)
  metaRow[9]  = grandTotalBill;  // col J — total bill

  // Row 2: headers
  const headerRow: unknown[] = [
    'Week',           // A
    'Associate Name', // B
    'Associate ID',   // C
    'Title',          // D
    'Associate State',// E
    'Hours',          // F
    'Hourly Rate',    // G
    'Total',          // H
    '% Allocation',   // I
    'Total Bill',     // J
    'Comments',       // K
    '',               // L
    '',               // M
  ];

  const dataRows = rows.map(r => [
    r.week,                    // A
    r.associateName,           // B
    r.associateId,             // C
    r.title,                   // D
    r.associateState,          // E
    r.hours,                   // F
    fmt(r.hourlyRate),         // G
    fmt(r.total),              // H
    pctFmt(r.allocationPct),   // I
    fmt(r.totalBill),          // J
    '',                        // K (Comments — blank)
    '',                        // L
    '',                        // M
  ]);

  return [metaRow, headerRow, ...dataRows];
}

// ── Cover Tab builder ──────────────────────────────────────────────────────
//
// Layout (columns A–E, rows 7–end):
//   Rows 7–11:   Static company header + remit-to info
//   Rows 13–18:  Invoice metadata (bill-to, invoice #, dates, PO#)
//   Rows 20–21:  Total Due label + amount
//   Row 24:      Line-item table headers
//   Rows 25+:    Manager rows (non-zero allocation), then FSM I + FSM II rows
//   Subtotal row
//   2 rows gap
//   Secondary section header
//   Secondary section rows (Cloud Services, Remote Management Software)
//   Grand Totals row

interface CoverParams {
  invoiceName: string;       // e.g. "FSM26-W15" or "FSM26-W15-16"
  periodStr: string;         // e.g. "04/06/2026 - 04/12/2026"
  weekEndDate: Date | null;  // last date in the period (used to compute invoice/due dates)
  mgmtRows: MgmtRow[];
  fsmIRows: InvoiceRow[];
  fsmIIRows: InvoiceRow[];
  fsmITotal: number;
  fsmIITotal: number;
  mgmtTotal: number;
}

function buildCoverTab(p: CoverParams): unknown[][] {
  // We build the sheet as a sparse array-of-arrays (row index 0-based, col index 0-based).
  // Column mapping: A=0, B=1, C=2, D=3, E=4
  // Row mapping:    row N in spec → index N-1 in array

  const NUM_COLS = 5;
  const rows: unknown[][] = [];

  const setCell = (rowIdx: number, colIdx: number, value: unknown) => {
    while (rows.length <= rowIdx) rows.push(new Array(NUM_COLS).fill(''));
    rows[rowIdx][colIdx] = value;
  };

  // ── Static company header (rows 7–11, idx 6–10) ──────────────────────────
  setCell(6,  0, '2020 Companies, Inc');
  setCell(6,  4, 'Remit To Information:');
  setCell(7,  0, '1900 W. Kirkwood Blvd.');
  setCell(7,  4, 'Wells Fargo Bank N.A.');
  setCell(8,  0, 'Suite 1300B');
  setCell(8,  4, 'Acct# 4255732695');
  setCell(9,  0, 'Southlake, Texas 76092');
  setCell(9,  4, 'Routing #  121000248');
  setCell(10, 0, 'FEIN: 26-2002404');

  // ── Invoice metadata (rows 13–18, idx 12–17) ─────────────────────────────
  setCell(12, 0, 'Bill To:');
  setCell(12, 3, 'Invoice #');
  setCell(12, 4, p.invoiceName);

  setCell(13, 0, 'Samsung Electronics America, LLC');
  setCell(13, 3, 'Date of Activity:');
  setCell(13, 4, p.periodStr);

  setCell(14, 0, '700 Sylvan Avenue');
  setCell(14, 3, 'Invoice Date:');
  // Invoice Date = week-end date + 2 days
  if (p.weekEndDate) {
    const invDate = new Date(p.weekEndDate);
    invDate.setDate(invDate.getDate() + 2);
    // Store as Excel serial date so Excel formats it as a date
    const invSerial = XLSX.SSF.parse_date_code
      ? dateToExcelSerial(invDate)
      : formatDateMDY(invDate);
    setCell(14, 4, invSerial);
  }

  setCell(15, 0, 'Englewood Cliffs, NJ 07632');
  setCell(15, 3, 'Due Date:');
  // Due Date = Invoice Date + 30 days
  if (p.weekEndDate) {
    const dueDate = new Date(p.weekEndDate);
    dueDate.setDate(dueDate.getDate() + 2 + 30);
    const dueSerial = dateToExcelSerial(dueDate);
    setCell(15, 4, dueSerial);
  }

  setCell(16, 0, 'Attn: Michael Compean');
  setCell(16, 3, 'PO#:');
  setCell(16, 4, 'T26C31H000162');

  setCell(17, 0, 'm.compean@partner.sea.samsung.com');

  // ── Total Due (rows 20–21, idx 19–20) ────────────────────────────────────
  setCell(19, 4, 'Total Due');

  // Grand total = mgmt + FSM I + FSM II
  const grandTotal = p.mgmtTotal + p.fsmITotal + p.fsmIITotal;
  setCell(20, 4, grandTotal);

  // ── Line items table header (row 24, idx 23) ─────────────────────────────
  setCell(23, 0, 'Description');
  setCell(23, 1, 'State');
  setCell(23, 2, 'Qty');
  setCell(23, 3, 'Rate');
  setCell(23, 4, 'Grand Total');

  // ── Management line items (rows 25+, idx 24+) ────────────────────────────
  // One row per manager with totalBill > 0, in MGMT_TABLE order.
  // Col A: title — but only if different from the previous row's title.
  // Col B: associateState
  // Col C: 1
  // Col D: totalBill (= rate for that manager)
  // Col E: totalBill

  // Aggregate by associateId across all weeks (sum totalBill)
  const mgmtByAssoc = new Map<string, MgmtRow>();
  for (const r of p.mgmtRows) {
    const existing = mgmtByAssoc.get(r.associateId);
    if (!existing) {
      mgmtByAssoc.set(r.associateId, { ...r });
    } else {
      existing.totalBill = Math.round((existing.totalBill + r.totalBill) * 100) / 100;
    }
  }

  let dataRowIdx = 24; // row 25, 0-based
  let prevTitle = '';

  for (const m of MGMT_TABLE) {
    const agg = mgmtByAssoc.get(m.associateId);
    if (!agg || agg.totalBill === 0) continue; // skip zero-allocation managers

    const titleCell = m.title !== prevTitle ? m.title : '';
    setCell(dataRowIdx, 0, titleCell);
    setCell(dataRowIdx, 1, m.associateState);
    setCell(dataRowIdx, 2, 1);
    setCell(dataRowIdx, 3, agg.totalBill);
    setCell(dataRowIdx, 4, agg.totalBill);
    prevTitle = m.title;
    dataRowIdx++;
  }

  // ── FSM I row ─────────────────────────────────────────────────────────────
  const fsmIHours = p.fsmIRows.reduce((s, r) => s + r.timeHours, 0);
  setCell(dataRowIdx, 0, 'FSM I');
  setCell(dataRowIdx, 1, 'ALL');
  setCell(dataRowIdx, 2, fsmIHours);
  setCell(dataRowIdx, 3, 0);
  setCell(dataRowIdx, 4, p.fsmITotal);
  dataRowIdx++;

  // ── FSM II row ────────────────────────────────────────────────────────────
  const fsmIIHours = p.fsmIIRows.reduce((s, r) => s + r.timeHours, 0);
  setCell(dataRowIdx, 0, 'FSM II');
  setCell(dataRowIdx, 1, 'ALL');
  setCell(dataRowIdx, 2, fsmIIHours);
  setCell(dataRowIdx, 3, 0);
  setCell(dataRowIdx, 4, p.fsmIITotal);
  dataRowIdx++;

  // ── Subtotal row ──────────────────────────────────────────────────────────
  setCell(dataRowIdx, 2, fsmIHours + fsmIIHours);
  setCell(dataRowIdx, 4, grandTotal);
  dataRowIdx++;

  // ── Gap (2 rows) ──────────────────────────────────────────────────────────
  dataRowIdx += 2;

  // ── Secondary section header ──────────────────────────────────────────────
  setCell(dataRowIdx, 0, 'Description');
  setCell(dataRowIdx, 1, 'Description');
  setCell(dataRowIdx, 2, 'QTY');
  setCell(dataRowIdx, 3, 'Rate');
  setCell(dataRowIdx, 4, 'Grand Total');
  dataRowIdx++;

  // ── Cloud Services — MGR row ──────────────────────────────────────────────
  setCell(dataRowIdx, 0, 'Cloud Services');
  setCell(dataRowIdx, 1, 'MGR');
  setCell(dataRowIdx, 3, 'various');
  dataRowIdx++;

  // ── Cloud Services — Field row ────────────────────────────────────────────
  setCell(dataRowIdx, 1, 'Field');
  setCell(dataRowIdx, 3, 40);
  setCell(dataRowIdx, 4, 0);
  dataRowIdx++;

  // ── Remote Management Software row ───────────────────────────────────────
  setCell(dataRowIdx, 0, 'Remote Management Software');
  setCell(dataRowIdx, 1, 'Jan');
  setCell(dataRowIdx, 3, 42.22);
  setCell(dataRowIdx, 4, 0);
  dataRowIdx++;

  // ── Grand Totals row ──────────────────────────────────────────────────────
  setCell(dataRowIdx, 0, 'Grand Totals');
  setCell(dataRowIdx, 2, 0);
  setCell(dataRowIdx, 4, 0);

  return rows;
}

// ── Date helpers ──────────────────────────────────────────────────────────

/** Convert a JS Date to an Excel serial number (days since 1900-01-01, with Lotus 1900 leap-year bug). */
function dateToExcelSerial(d: Date): number {
  // Excel epoch is Jan 1 1900 = serial 1 (with the 1900 leap-year bug: serial 60 is Feb 29 1900, which didn't exist)
  const epoch = new Date(Date.UTC(1899, 11, 30)); // Dec 30 1899
  const utcDate = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate());
  return (utcDate - epoch.getTime()) / 86400000;
}

function formatDateMDY(d: Date): string {
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${mm}/${dd}/${d.getFullYear()}`;
}

// ── Invoice naming ────────────────────────────────────────────────────────

/** Build the invoice name string: "FSM26-W15" (single) or "FSM26-W15-16" (bi-weekly) */
function buildInvoiceName(weeks: number[]): string {
  if (weeks.length === 0) return 'FSM-Invoice';
  const year = new Date().getFullYear().toString().slice(-2);
  const w1 = String(weeks[0]).padStart(2, '0');
  if (weeks.length === 1) return `FSM${year}-W${w1}`;
  const w2 = String(weeks[weeks.length - 1]).padStart(2, '0');
  return `FSM${year}-W${w1}-${w2}`;
}

// ── Weekly workbook (Mode 1) ────────────────────────────────────────────────

export function buildWeeklyWorkbook(summary: WeeklySummary): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Collect visitDates from labor rows for cover tab date computations
  const allRows = [...summary.fsmI, ...summary.fsmII];
  const periodStr = buildCoverPeriodStr(allRows);
  const weekEndDate = getWeekEndDate(allRows);
  const invoiceName = buildInvoiceName(summary.weeks);

  // Cover tab — first tab, named after the invoice
  const coverData = buildCoverTab({
    invoiceName,
    periodStr,
    weekEndDate,
    mgmtRows: summary.mgmt,
    fsmIRows: summary.fsmI,
    fsmIIRows: summary.fsmII,
    fsmITotal: summary.fsmITotal,
    fsmIITotal: summary.fsmIITotal,
    mgmtTotal: summary.mgmtTotal,
  });
  const wsCover = XLSX.utils.aoa_to_sheet(coverData);
  applyColumnWidths(wsCover, [42, 10, 14, 20, 20]);
  // Mark date cells in E15 and E16 as date format
  applyCoverDateFormats(wsCover, coverData);
  XLSX.utils.book_append_sheet(wb, wsCover, invoiceName);

  // FSM I — second tab
  const fsmIData = buildLaborTab(summary.fsmI, summary.fsmITotal, 'OT-15.53', true);
  const wsFsmI = XLSX.utils.aoa_to_sheet(fsmIData);
  applyColumnWidths(wsFsmI, [8, 30, 14, 14, 20, 22, 28, 10, 40, 14, 12, 10, 14, 12, 14, 12, 14, 12, 20, 6]);
  XLSX.utils.book_append_sheet(wb, wsFsmI, 'FSM I');

  // FSM II — third tab
  const fsmIIData = buildLaborTab(summary.fsmII, summary.fsmIITotal, 'OT-17.75', false);
  const wsFsmII = XLSX.utils.aoa_to_sheet(fsmIIData);
  applyColumnWidths(wsFsmII, [8, 30, 14, 14, 20, 22, 28, 10, 40, 14, 12, 10, 14, 12, 14, 12, 14, 12, 20]);
  XLSX.utils.book_append_sheet(wb, wsFsmII, 'FSM II');

  // Management Detail — fourth tab
  const mgmtData = buildMgmtTab(summary.mgmt);
  const wsMgmt = XLSX.utils.aoa_to_sheet(mgmtData);
  applyColumnWidths(wsMgmt, [8, 32, 14, 35, 14, 8, 12, 14, 14, 12, 20, 6, 6]);
  XLSX.utils.book_append_sheet(wb, wsMgmt, 'Management Detail Hours');

  return wb;
}

/** Generate weekly filename: FSM[YY]-W[WW]_Weekly.xlsx */
export function buildWeeklyFilename(weeks: number[]): string {
  if (weeks.length === 0) return 'FSM-Weekly.xlsx';
  const year = new Date().getFullYear().toString().slice(-2);
  const w = String(weeks[0]).padStart(2, '0');
  return `FSM${year}-W${w}_Weekly.xlsx`;
}

// ── Client invoice workbook (Mode 2) ──────────────────────────────────────

export function buildWorkbook(summary: InvoiceSummary): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  const allRows = [...summary.fsmI, ...summary.fsmII];
  const periodStr = buildCoverPeriodStr(allRows);
  const weekEndDate = getWeekEndDate(allRows);
  const invoiceName = buildInvoiceName(summary.weeks);

  // Cover tab — first tab
  const coverData = buildCoverTab({
    invoiceName,
    periodStr,
    weekEndDate,
    mgmtRows: summary.mgmt,
    fsmIRows: summary.fsmI,
    fsmIIRows: summary.fsmII,
    fsmITotal: summary.fsmITotal,
    fsmIITotal: summary.fsmIITotal,
    mgmtTotal: summary.mgmtTotal,
  });
  const wsCover = XLSX.utils.aoa_to_sheet(coverData);
  applyColumnWidths(wsCover, [42, 10, 14, 20, 20]);
  applyCoverDateFormats(wsCover, coverData);
  XLSX.utils.book_append_sheet(wb, wsCover, invoiceName);

  // FSM I — second tab (20 columns with empty col T)
  const fsmIData = buildLaborTab(summary.fsmI, summary.fsmITotal, 'OT-15.53', true);
  const wsFsmI = XLSX.utils.aoa_to_sheet(fsmIData);
  applyColumnWidths(wsFsmI, [8, 30, 14, 14, 20, 22, 28, 10, 40, 14, 12, 10, 14, 12, 14, 12, 14, 12, 20, 6]);
  XLSX.utils.book_append_sheet(wb, wsFsmI, 'FSM I');

  // FSM II — third tab (19 columns, no col T)
  const fsmIIData = buildLaborTab(summary.fsmII, summary.fsmIITotal, 'OT-17.75', false);
  const wsFsmII = XLSX.utils.aoa_to_sheet(fsmIIData);
  applyColumnWidths(wsFsmII, [8, 30, 14, 14, 20, 22, 28, 10, 40, 14, 12, 10, 14, 12, 14, 12, 14, 12, 20]);
  XLSX.utils.book_append_sheet(wb, wsFsmII, 'FSM II');

  // Management Detail Hours — fourth tab (13 columns)
  const mgmtData = buildMgmtTab(summary.mgmt);
  const wsMgmt = XLSX.utils.aoa_to_sheet(mgmtData);
  applyColumnWidths(wsMgmt, [8, 32, 14, 35, 14, 8, 12, 14, 14, 12, 20, 6, 6]);
  XLSX.utils.book_append_sheet(wb, wsMgmt, 'Management Detail Hours');

  return wb;
}

function applyColumnWidths(ws: XLSX.WorkSheet, widths: number[]) {
  ws['!cols'] = widths.map(w => ({ wch: w }));
}

/**
 * After aoa_to_sheet, mark the invoice date (E15, row idx 14) and due date (E16, row idx 15)
 * cells with Excel date number format so they render as dates in Excel.
 * We locate them by scanning the cover data for numeric values in column E rows 14–15.
 */
function applyCoverDateFormats(ws: XLSX.WorkSheet, data: unknown[][]) {
  // Row 14 (0-based) = spec row 15 (Invoice Date), Row 15 = spec row 16 (Due Date)
  const dateRowIdxs = [14, 15];
  for (const rowIdx of dateRowIdxs) {
    const row = data[rowIdx];
    if (!row) continue;
    const val = row[4]; // col E
    if (typeof val === 'number') {
      const cellRef = XLSX.utils.encode_cell({ r: rowIdx, c: 4 });
      if (ws[cellRef]) {
        ws[cellRef].t = 'n';
        ws[cellRef].z = 'mm/dd/yyyy';
      }
    }
  }
}

/** Generate filename: FSM[YY]-W[WW].xlsx (single) or FSM[YY]-W[WW]-[WW].xlsx (bi-weekly) */
export function buildFilename(weeks: number[]): string {
  if (weeks.length === 0) return 'FSM-Invoice.xlsx';
  return buildInvoiceName(weeks) + '.xlsx';
}

/** Trigger browser download of the workbook */
export function downloadWorkbook(wb: XLSX.WorkBook, filename: string): void {
  XLSX.writeFile(wb, filename);
}
