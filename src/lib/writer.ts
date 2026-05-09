import * as XLSX from 'xlsx';
import type { InvoiceSummary, WeeklySummary, InvoiceRow, MgmtRow } from './types';
import { MGMT_TABLE } from './constants';
import { buildCoverPeriodStr, getWeekEndDate } from './transform';

// ── Number format strings ──────────────────────────────────────────────────
const FMT_ACCOUNTING = '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)';
const FMT_2DP        = '0.00';
const FMT_DATE       = 'mm/dd/yyyy';
const FMT_PCT        = '0.00%';
const FMT_DOLLAR     = '"$"#,##0.00';

// ── Style builders ─────────────────────────────────────────────────────────

type CellStyle = {
  font?: { bold?: boolean; sz?: number; name?: string };
  fill?: { fgColor?: { rgb: string }; patternType?: string };
  numFmt?: string;
  alignment?: { horizontal?: string; vertical?: string; wrapText?: boolean };
};

function boldStyle(): CellStyle {
  return { font: { bold: true } };
}

function boldAccounting(): CellStyle {
  return { font: { bold: true }, numFmt: FMT_ACCOUNTING };
}

function yellowStyle(): CellStyle {
  return { fill: { fgColor: { rgb: 'FFFF00' }, patternType: 'solid' } };
}

function yellowAccountingStyle(): CellStyle {
  return {
    fill: { fgColor: { rgb: 'FFFF00' }, patternType: 'solid' },
    numFmt: FMT_ACCOUNTING,
  };
}

function numFmtStyle(fmt: string): CellStyle {
  return { numFmt: fmt };
}

/** Apply a style object to a specific cell ref in a worksheet. */
function styleCell(ws: XLSX.WorkSheet, cellRef: string, style: CellStyle) {
  const cell = ws[cellRef];
  if (!cell) return;
  // Merge — preserve existing z (number format) only if we are not overriding
  cell.s = { ...(cell.s ?? {}), ...style };
  // If numFmt is provided, also set the z property (SheetJS uses both)
  if (style.numFmt) {
    cell.z = style.numFmt;
  }
}

/** Apply a style to a range of cells (e.g. row 2, cols A–T). */
function styleRow(ws: XLSX.WorkSheet, rowIdx: number, colStart: number, colEnd: number, style: CellStyle) {
  for (let c = colStart; c <= colEnd; c++) {
    const ref = XLSX.utils.encode_cell({ r: rowIdx, c });
    styleCell(ws, ref, style);
  }
}

// ── Formatting helpers (string) ────────────────────────────────────────────

function fmt(n: number): string {
  return n.toFixed(2);
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
  metaRow[17] = tabTotal;             // col R — total Bill (number, styled with accounting fmt)

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
      r.visitDate,        // M — kept as string; styled with date fmt below
      r.timeHours,        // N
      r.basePayRate,      // O — number, styled 0.00
      fmt(r.mu),          // P
      r.payRateTotal,     // Q — number, styled 0.00
      r.bill,             // R — number, styled accounting
      r.comments,         // S
    ];
    if (extraEmptyCol) row.push(''); // T
    return row;
  });

  return [metaRow, headerRow, ...dataRows];
}

/** Apply formatting to an FSM I or FSM II worksheet after aoa_to_sheet. */
function applyLaborStyles(ws: XLSX.WorkSheet, numDataRows: number, hasExtraCol: boolean) {
  const lastCol = hasExtraCol ? 19 : 18; // T=19, S=18 (0-based)

  // Row 1 (idx 0) metadata cells:
  //   N1 (col 13): 0.00
  //   O1 (col 14): already a string label — skip
  //   R1 (col 17): accounting $
  styleCell(ws, XLSX.utils.encode_cell({ r: 0, c: 13 }), numFmtStyle(FMT_2DP));
  styleCell(ws, XLSX.utils.encode_cell({ r: 0, c: 17 }), numFmtStyle(FMT_ACCOUNTING));

  // Row 2 (idx 1) headers: bold all columns
  styleRow(ws, 1, 0, lastCol, boldStyle());

  // Data rows (idx 2+):
  for (let i = 0; i < numDataRows; i++) {
    const r = i + 2;
    // M (col 12): date format
    styleCell(ws, XLSX.utils.encode_cell({ r, c: 12 }), numFmtStyle(FMT_DATE));
    // O (col 14): 0.00 — base pay rate (number)
    styleCell(ws, XLSX.utils.encode_cell({ r, c: 14 }), numFmtStyle(FMT_2DP));
    // Q (col 16): 0.00 — pay rate total
    styleCell(ws, XLSX.utils.encode_cell({ r, c: 16 }), numFmtStyle(FMT_2DP));
    // R (col 17): accounting $
    styleCell(ws, XLSX.utils.encode_cell({ r, c: 17 }), numFmtStyle(FMT_ACCOUNTING));
  }
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
  metaRow[2]  = rows.length;      // col C — count of data rows
  metaRow[9]  = grandTotalBill;   // col J — total bill (number, styled accounting + yellow)

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
    r.hours,                   // F — number, styled 0.00
    r.hourlyRate,              // G — number, styled "$"#,##0.00
    r.total,                   // H — number, styled accounting
    r.allocationPct,           // I — number, styled 0.00%
    r.totalBill,               // J — number, styled accounting
    '',                        // K (Comments — blank)
    '',                        // L
    '',                        // M
  ]);

  return [metaRow, headerRow, ...dataRows];
}

/** Apply formatting to the Management Detail worksheet. */
function applyMgmtStyles(ws: XLSX.WorkSheet, numDataRows: number) {
  // Row 1 (idx 0): J1 (col 9) — yellow fill + accounting
  styleCell(ws, XLSX.utils.encode_cell({ r: 0, c: 9 }), yellowAccountingStyle());

  // Row 2 (idx 1): bold all headers (13 cols)
  styleRow(ws, 1, 0, 12, boldStyle());

  // Data rows (idx 2+)
  for (let i = 0; i < numDataRows; i++) {
    const r = i + 2;
    styleCell(ws, XLSX.utils.encode_cell({ r, c: 5 }),  numFmtStyle(FMT_2DP));        // F: Hours
    styleCell(ws, XLSX.utils.encode_cell({ r, c: 6 }),  numFmtStyle(FMT_DOLLAR));     // G: Hourly Rate
    styleCell(ws, XLSX.utils.encode_cell({ r, c: 7 }),  numFmtStyle(FMT_ACCOUNTING)); // H: Total
    styleCell(ws, XLSX.utils.encode_cell({ r, c: 8 }),  numFmtStyle(FMT_PCT));        // I: % Allocation
    styleCell(ws, XLSX.utils.encode_cell({ r, c: 9 }),  numFmtStyle(FMT_ACCOUNTING)); // J: Total Bill
  }
}

// ── New Hire Fee params ────────────────────────────────────────────────────

export interface NhfParams {
  count: number;
  rate: number; // 285 or 395
  label: string; // e.g. "2020 Transfer", "In Talent Network", "New Hire"
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
//   Secondary section rows (Cloud Services MGR, Cloud Services FT, NHF if present)
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
  nhf: NhfParams | null;     // null = no new hires
}

function buildCoverTab(p: CoverParams): { data: unknown[][]; subtotalRowIdx: number; secondaryHeaderRowIdx: number; lastDataRowIdx: number } {
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
    const invSerial = dateToExcelSerial(invDate);
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

  // ── Compute secondary section totals ─────────────────────────────────────
  // Cloud Services — MGR: count distinct managers with totalBill > 0
  const mgmtByAssocForCloud = new Map<string, number>();
  for (const r of p.mgmtRows) {
    mgmtByAssocForCloud.set(r.associateId, (mgmtByAssocForCloud.get(r.associateId) ?? 0) + r.totalBill);
  }
  const mgrCount = [...mgmtByAssocForCloud.values()].filter(v => v > 0).length;
  const mgrRate = 54;
  const mgrCloudTotal = mgrCount * mgrRate;

  // Cloud Services — FT: count unique FT associate IDs across fsmI + fsmII
  const ftIds = new Set<string>();
  for (const r of [...p.fsmIRows, ...p.fsmIIRows]) {
    if (r.associateType === 'FT') {
      ftIds.add(r.associateId.trim());
    }
  }
  const ftCount = ftIds.size;
  const ftRate = 40;
  const ftCloudTotal = ftCount * ftRate;

  // NHF total
  const nhfTotal = p.nhf ? p.nhf.count * p.nhf.rate : 0;

  // ── Total Due = mgmt + FSM I + FSM II + cloud services + NHF ─────────────
  const primaryTotal = p.mgmtTotal + p.fsmITotal + p.fsmIITotal;
  const grandTotal = primaryTotal + mgrCloudTotal + ftCloudTotal + nhfTotal;

  // ── Total Due (rows 20–21, idx 19–20) ────────────────────────────────────
  setCell(19, 4, 'Total Due');
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
  const subtotalRowIdx = dataRowIdx;
  setCell(dataRowIdx, 2, fsmIHours + fsmIIHours);
  setCell(dataRowIdx, 4, primaryTotal);
  dataRowIdx++;

  // ── Gap (2 rows) ──────────────────────────────────────────────────────────
  dataRowIdx += 2;

  // ── Secondary section header ──────────────────────────────────────────────
  const secondaryHeaderRowIdx = dataRowIdx;
  setCell(dataRowIdx, 0, 'Description');
  setCell(dataRowIdx, 1, 'Description');
  setCell(dataRowIdx, 2, 'QTY');
  setCell(dataRowIdx, 3, 'Rate');
  setCell(dataRowIdx, 4, 'Grand Total');
  dataRowIdx++;

  // ── Cloud Services — MGR row ──────────────────────────────────────────────
  setCell(dataRowIdx, 0, 'Cloud Services');
  setCell(dataRowIdx, 1, 'MGR');
  setCell(dataRowIdx, 2, mgrCount);
  setCell(dataRowIdx, 3, mgrRate);
  setCell(dataRowIdx, 4, mgrCloudTotal);
  dataRowIdx++;

  // ── Cloud Services — FT row ───────────────────────────────────────────────
  setCell(dataRowIdx, 0, 'Cloud Services');
  setCell(dataRowIdx, 1, 'FT');
  setCell(dataRowIdx, 2, ftCount);
  setCell(dataRowIdx, 3, ftRate);
  setCell(dataRowIdx, 4, ftCloudTotal);
  dataRowIdx++;

  // ── New Hire Fee row (only if NHF present) ────────────────────────────────
  if (p.nhf) {
    setCell(dataRowIdx, 0, 'New Hire Fee');
    setCell(dataRowIdx, 1, p.nhf.label);
    setCell(dataRowIdx, 2, p.nhf.count);
    setCell(dataRowIdx, 3, p.nhf.rate);
    setCell(dataRowIdx, 4, nhfTotal);
    dataRowIdx++;
  }

  // ── Grand Totals row ──────────────────────────────────────────────────────
  const lastDataRowIdx = dataRowIdx;
  setCell(dataRowIdx, 0, 'Grand Totals');
  setCell(dataRowIdx, 2, ftCount + mgrCount + (p.nhf ? p.nhf.count : 0));
  setCell(dataRowIdx, 4, grandTotal);

  return { data: rows, subtotalRowIdx, secondaryHeaderRowIdx, lastDataRowIdx };
}

/** Apply formatting to the Cover worksheet. */
function applyCoverStyles(
  ws: XLSX.WorkSheet,
  data: unknown[][],
  subtotalRowIdx: number,
  secondaryHeaderRowIdx: number,
  lastDataRowIdx: number,
) {
  // A7 (idx 6): bold — company name
  styleCell(ws, XLSX.utils.encode_cell({ r: 6, c: 0 }), boldStyle());

  // D13–D17 (idx 12–16, col 3): bold — label column
  for (let r = 12; r <= 16; r++) {
    styleCell(ws, XLSX.utils.encode_cell({ r, c: 3 }), boldStyle());
  }

  // E13–E17 (idx 12–16, col 4): yellow fill — dynamic fields
  for (let r = 12; r <= 16; r++) {
    styleCell(ws, XLSX.utils.encode_cell({ r, c: 4 }), yellowStyle());
  }

  // E20 (idx 19): bold — "Total Due"
  styleCell(ws, XLSX.utils.encode_cell({ r: 19, c: 4 }), boldStyle());

  // E21 (idx 20): bold + accounting — grand total amount
  styleCell(ws, XLSX.utils.encode_cell({ r: 20, c: 4 }), boldAccounting());

  // Row 24 (idx 23) headers: bold all 5 cols
  styleRow(ws, 23, 0, 4, boldStyle());

  // Rows 25+ through subtotal row: bold all (management line items, FSM I, FSM II, subtotal)
  for (let r = 24; r <= subtotalRowIdx; r++) {
    styleRow(ws, r, 0, 4, boldStyle());
  }

  // Secondary section header: bold all 5 cols
  styleRow(ws, secondaryHeaderRowIdx, 0, 4, boldStyle());

  // Secondary section data rows: Cloud Services + NHF (rows between header and grand totals)
  for (let r = secondaryHeaderRowIdx + 1; r <= lastDataRowIdx; r++) {
    styleRow(ws, r, 0, 4, boldStyle());
  }

  // Grand Totals row: bold
  styleRow(ws, lastDataRowIdx, 0, 4, boldStyle());

  // Date cells: E15 and E16 (idx 14, 15) get date format (preserve yellow fill)
  const dateRowIdxs = [14, 15];
  for (const rowIdx of dateRowIdxs) {
    const row = data[rowIdx];
    if (!row) continue;
    const val = row[4]; // col E
    if (typeof val === 'number') {
      const cellRef = XLSX.utils.encode_cell({ r: rowIdx, c: 4 });
      const cell = ws[cellRef];
      if (cell) {
        cell.t = 'n';
        cell.z = FMT_DATE;
        cell.s = { ...(cell.s ?? {}), numFmt: FMT_DATE };
      }
    }
  }
}

// ── Date helpers ──────────────────────────────────────────────────────────

/** Convert a JS Date to an Excel serial number (days since 1900-01-01, with Lotus 1900 leap-year bug). */
function dateToExcelSerial(d: Date): number {
  // Excel epoch is Jan 1 1900 = serial 1 (with the 1900 leap-year bug: serial 60 is Feb 29 1900, which didn't exist)
  const epoch = new Date(Date.UTC(1899, 11, 30)); // Dec 30 1899
  const utcDate = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate());
  return (utcDate - epoch.getTime()) / 86400000;
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

// ── Column width constants ─────────────────────────────────────────────────
// Exact widths from reference invoice
const FSM_I_WIDTHS  = [8.29, 24.86, 13.43, 15.71, 29.43, 36.43, 49.43, 30.71, 53.14, 16.14, 12.86, 10.43, 22.0, 12.43, 16.14, 8.43, 16.43, 13.0, 21.0, 6];
const FSM_II_WIDTHS = FSM_I_WIDTHS.slice(0, 19);
const MGMT_WIDTHS   = [8, 25.29, 11.71, 33.14, 16.86, 8.14, 8.14, 18.0, 11.0, 13.43, 13.43, 6, 6];

// ── Weekly workbook (Mode 1) ────────────────────────────────────────────────

export function buildWeeklyWorkbook(summary: WeeklySummary, nhf: NhfParams | null = null): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Collect visitDates from labor rows for cover tab date computations
  const allRows = [...summary.fsmI, ...summary.fsmII];
  const periodStr = buildCoverPeriodStr(allRows);
  const weekEndDate = getWeekEndDate(allRows);
  const invoiceName = buildInvoiceName(summary.weeks);

  // Cover tab — first tab, named after the invoice
  const { data: coverData, subtotalRowIdx, secondaryHeaderRowIdx, lastDataRowIdx } = buildCoverTab({
    invoiceName,
    periodStr,
    weekEndDate,
    mgmtRows: summary.mgmt,
    fsmIRows: summary.fsmI,
    fsmIIRows: summary.fsmII,
    fsmITotal: summary.fsmITotal,
    fsmIITotal: summary.fsmIITotal,
    mgmtTotal: summary.mgmtTotal,
    nhf,
  });
  const wsCover = XLSX.utils.aoa_to_sheet(coverData);
  applyColumnWidths(wsCover, [42, 10, 14, 20, 20]);
  applyCoverStyles(wsCover, coverData, subtotalRowIdx, secondaryHeaderRowIdx, lastDataRowIdx);
  XLSX.utils.book_append_sheet(wb, wsCover, invoiceName);

  // FSM I — second tab
  const fsmIData = buildLaborTab(summary.fsmI, summary.fsmITotal, 'OT-15.53', true);
  const wsFsmI = XLSX.utils.aoa_to_sheet(fsmIData);
  applyColumnWidths(wsFsmI, FSM_I_WIDTHS);
  applyLaborStyles(wsFsmI, summary.fsmI.length, true);
  XLSX.utils.book_append_sheet(wb, wsFsmI, 'FSM I');

  // FSM II — third tab
  const fsmIIData = buildLaborTab(summary.fsmII, summary.fsmIITotal, 'OT-17.75', false);
  const wsFsmII = XLSX.utils.aoa_to_sheet(fsmIIData);
  applyColumnWidths(wsFsmII, FSM_II_WIDTHS);
  applyLaborStyles(wsFsmII, summary.fsmII.length, false);
  XLSX.utils.book_append_sheet(wb, wsFsmII, 'FSM II');

  // Management Detail — fourth tab
  const mgmtData = buildMgmtTab(summary.mgmt);
  const wsMgmt = XLSX.utils.aoa_to_sheet(mgmtData);
  applyColumnWidths(wsMgmt, MGMT_WIDTHS);
  applyMgmtStyles(wsMgmt, summary.mgmt.length);
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

export function buildWorkbook(summary: InvoiceSummary, nhf: NhfParams | null = null): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  const allRows = [...summary.fsmI, ...summary.fsmII];
  const periodStr = buildCoverPeriodStr(allRows);
  const weekEndDate = getWeekEndDate(allRows);
  const invoiceName = buildInvoiceName(summary.weeks);

  // Cover tab — first tab
  const { data: coverData, subtotalRowIdx, secondaryHeaderRowIdx, lastDataRowIdx } = buildCoverTab({
    invoiceName,
    periodStr,
    weekEndDate,
    mgmtRows: summary.mgmt,
    fsmIRows: summary.fsmI,
    fsmIIRows: summary.fsmII,
    fsmITotal: summary.fsmITotal,
    fsmIITotal: summary.fsmIITotal,
    mgmtTotal: summary.mgmtTotal,
    nhf,
  });
  const wsCover = XLSX.utils.aoa_to_sheet(coverData);
  applyColumnWidths(wsCover, [42, 10, 14, 20, 20]);
  applyCoverStyles(wsCover, coverData, subtotalRowIdx, secondaryHeaderRowIdx, lastDataRowIdx);
  XLSX.utils.book_append_sheet(wb, wsCover, invoiceName);

  // FSM I — second tab (20 columns with empty col T)
  const fsmIData = buildLaborTab(summary.fsmI, summary.fsmITotal, 'OT-15.53', true);
  const wsFsmI = XLSX.utils.aoa_to_sheet(fsmIData);
  applyColumnWidths(wsFsmI, FSM_I_WIDTHS);
  applyLaborStyles(wsFsmI, summary.fsmI.length, true);
  XLSX.utils.book_append_sheet(wb, wsFsmI, 'FSM I');

  // FSM II — third tab (19 columns, no col T)
  const fsmIIData = buildLaborTab(summary.fsmII, summary.fsmIITotal, 'OT-17.75', false);
  const wsFsmII = XLSX.utils.aoa_to_sheet(fsmIIData);
  applyColumnWidths(wsFsmII, FSM_II_WIDTHS);
  applyLaborStyles(wsFsmII, summary.fsmII.length, false);
  XLSX.utils.book_append_sheet(wb, wsFsmII, 'FSM II');

  // Management Detail Hours — fourth tab (13 columns)
  const mgmtData = buildMgmtTab(summary.mgmt);
  const wsMgmt = XLSX.utils.aoa_to_sheet(mgmtData);
  applyColumnWidths(wsMgmt, MGMT_WIDTHS);
  applyMgmtStyles(wsMgmt, summary.mgmt.length);
  XLSX.utils.book_append_sheet(wb, wsMgmt, 'Management Detail Hours');

  return wb;
}

function applyColumnWidths(ws: XLSX.WorkSheet, widths: number[]) {
  ws['!cols'] = widths.map(w => ({ wch: w }));
}

/** Generate filename: FSM[YY]-W[WW].xlsx (single) or FSM[YY]-W[WW]-[WW].xlsx (bi-weekly) */
export function buildFilename(weeks: number[]): string {
  if (weeks.length === 0) return 'FSM-Invoice.xlsx';
  return buildInvoiceName(weeks) + '.xlsx';
}

/** Trigger browser download of the workbook */
export function downloadWorkbook(wb: XLSX.WorkBook, filename: string): void {
  XLSX.writeFile(wb, filename, { cellStyles: true });
}
