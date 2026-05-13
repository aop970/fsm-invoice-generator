/**
 * filler.ts — xlsx-populate template fill logic
 * Loads FSM_Invoice_TEMPLATE.xlsx, fills placeholders, writes data rows, returns buffer.
 */

import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';
import XlsxPopulate from 'xlsx-populate';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const TEMPLATE_GITHUB_URL =
  'https://raw.githubusercontent.com/aop970/fsm-invoice-generator/main/templates/FSM_Invoice_TEMPLATE.xlsx';

async function loadWorkbook(): Promise<unknown> {
  const localPath = path.resolve(__dirname, '../templates/FSM_Invoice_TEMPLATE.xlsx');
  if (fs.existsSync(localPath)) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (XlsxPopulate as any).fromFileAsync(localPath);
  }
  // Local file absent — fetch from GitHub (requires GITHUB_TOKEN for private repo)
  const headers: Record<string, string> = {};
  if (process.env.GITHUB_TOKEN) {
    headers['Authorization'] = `token ${process.env.GITHUB_TOKEN}`;
  }
  const res = await fetch(TEMPLATE_GITHUB_URL, { headers });
  if (!res.ok) {
    throw new Error(`Template fetch failed: ${res.status} ${res.statusText}`);
  }
  const buf = Buffer.from(await res.arrayBuffer());
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  return (XlsxPopulate as any).fromDataAsync(buf);
}

// ── Types (mirror src/lib/types.ts — no imports across browser/server boundary) ──

interface InvoiceRow {
  week: number;
  employeeName: string;
  associateId: string;
  associateType: string;
  region: string;
  district: string;
  market: string;
  storeId: string;
  storeName: string;
  associateState: string;
  storeState: string;
  zipCode: string;
  visitDate: string;
  timeHours: number;
  basePayRate: number;
  mu: number;
  payRateTotal: number;
  bill: number;
  comments: string;
}

interface MgmtRow {
  week: number;
  associateName: string;
  associateId: string;
  title: string;
  associateState: string;
  hours: number;
  hourlyRate: number;
  total: number;
  allocationPct: number;
  totalBill: number;
}

export interface NhfParams {
  count: number;
  rate: number;
  label: string;
}

export interface GeneratePayload {
  mode: 'weekly' | 'invoice';
  invoiceName: string;     // e.g. "FSM26-W15"
  periodStr: string;       // e.g. "04/06/2026 - 04/12/2026"
  invoiceDate: string;     // ISO date string e.g. "2026-04-14"
  dueDate: string;         // ISO date string e.g. "2026-05-14"
  poNumber: string;
  fsmI: InvoiceRow[];
  fsmII: InvoiceRow[];
  mgmt: MgmtRow[];
  fsmITotal: number;
  fsmIITotal: number;
  mgmtTotal: number;
  nhf: NhfParams | null;
  weeks: number[];
}

// ── INVOICE_TEMPLATE management row → associateId mapping ───────────────────
// Template rows 25-49 correspond to non-zero-allocation managers in MGMT_TABLE order.
// This array maps templateRow (index 0 = row 25) → associateId.
const MGMT_TEMPLATE_ROW_MAP: string[] = [
  'MK1002A',   // row 25 — Account Director (Mike King)
  'CP1010A',   // row 26 — National Manager (Corey Purdin, TX)
  'SS1183I',   // row 27 — National Manager, FSM (Shawn Scialo, AZ)
  'EE016130',  // row 28 — Operations Manager (Mike Coronado)
  'KE1021I',   // row 29 — Project Manager (Kristen Eberline)
  'MV1046I',   // row 30 — Operations Coordinator (Mason Vanmeter, AZ)
  'LJ1011I',   // row 31 — Operations Coordinator (Lisa Ghattas, MO)
  'JM1347I',   // row 32 — Data Analyst (George Macias)
  'MF1010I',   // row 33 — Training Manager (Mary Beth French)
  'EE004278',  // row 34 — Onboarding Specialist, FSM (Amanda Bradshaw)
  'HW1020I',   // row 35 — Field Operations Manager I (Hunter West, AZ)
  'EE010650',  // row 36 — Field Operations Manager I (Kenneth Hitt, AZ)
  'DE1016I',   // row 37 — Field Operations Manager I (David Exum, CA)
  'SW1147C',   // row 38 — Field Operations Manager I (Sara Wali, CA)
  'EE003625',  // row 39 — Field Operations Manager I (David Akom, GA)
  'EE010554',  // row 40 — Field Operations Manager I (Rebecca Tejeda, IL)
  'EL1011I',   // row 41 — Field Operations Manager I (Eric Lopez, KY)
  'EE010525',  // row 42 — Field Operations Manager I (Donald Scarfo, SC)
  'YS1013I',   // row 43 — Field Operations Manager I (Yamil Saade, TX)
  'EE003697',  // row 44 — Field Operations Manager I (Steffanie Molina-Frybarger, TX)
  'MW1112I',   // row 45 — Field Operations Manager I (Matthew Wickham, VA)
  'MA1183C',   // row 46 — Field Operations Manager I (Miguel Aguilar, VA)
  'EE010255',  // row 47 — Field Operations Manager I (Belinda Chi, WA)
  'EE002390',  // row 48 — Inventory Specialist (Brian Conner)
  'WM1054A',   // row 49 — Inventory/IT Specialist (Billy MacDonald)
];

// ── Date helpers ─────────────────────────────────────────────────────────────

/** Convert ISO date string (YYYY-MM-DD) to Excel serial number. */
function isoToExcelSerial(isoStr: string): number {
  const [year, month, day] = isoStr.split('-').map(Number);
  const epoch = new Date(Date.UTC(1899, 11, 30)); // Dec 30 1899
  const utcDate = Date.UTC(year, month - 1, day);
  return (utcDate - epoch.getTime()) / 86400000;
}

// ── Round helper ──────────────────────────────────────────────────────────────

function fmt2(n: number): string {
  return n.toFixed(2);
}

// ── Main filler ───────────────────────────────────────────────────────────────

export async function fillTemplate(payload: GeneratePayload): Promise<Buffer> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const wb = await loadWorkbook() as any;

  // ── 1. INVOICE_TEMPLATE placeholders ──────────────────────────────────────
  const inv = wb.sheet('INVOICE_TEMPLATE');

  inv.cell('E13').value(payload.invoiceName);
  inv.cell('E14').value(payload.periodStr);
  // Invoice Date as Excel serial (date format already applied in template)
  inv.cell('E15').value(isoToExcelSerial(payload.invoiceDate));
  // E16 has formula =E15+30 — leave it; it will auto-calculate
  inv.cell('E17').value(payload.poNumber);

  // FSM I totals
  const fsmIHours = payload.fsmI.reduce((s, r) => s + r.timeHours, 0);
  inv.cell('C50').value(fsmIHours);
  inv.cell('E50').value(payload.fsmITotal);

  // FSM II totals
  const fsmIIHours = payload.fsmII.reduce((s, r) => s + r.timeHours, 0);
  inv.cell('C51').value(fsmIIHours);
  inv.cell('E51').value(payload.fsmIITotal);

  // Cloud Services — Field Agent Count = unique FT associate IDs across FSM I + FSM II
  const ftIds = new Set<string>();
  for (const r of [...payload.fsmI, ...payload.fsmII]) {
    if (r.associateType === 'FT') {
      ftIds.add(r.associateId.trim());
    }
  }
  const ftCount = ftIds.size;
  inv.cell('C56').value(ftCount);
  inv.cell('C57').value(ftCount); // Remote SW qty = same as field agent count

  // ── 2. Management bill amounts in INVOICE_TEMPLATE rows 25-49 ─────────────
  // Aggregate totalBill by associateId across all mgmt rows (bi-weekly has 2 rows per manager)
  const mgmtBillMap = new Map<string, number>();
  for (const r of payload.mgmt) {
    mgmtBillMap.set(r.associateId, (mgmtBillMap.get(r.associateId) ?? 0) + r.totalBill);
  }

  MGMT_TEMPLATE_ROW_MAP.forEach((associateId, idx) => {
    const templateRow = 25 + idx;
    const totalBill = mgmtBillMap.get(associateId) ?? 0;
    inv.cell(templateRow, 4).value(totalBill); // col D = 4
  });

  // ── 3. FSM I tab — write metadata (row 1) and data rows (3+) ─────────────
  fillLaborTab(wb.sheet('FSM I'), payload.fsmI, payload.fsmITotal, 'OT-15.53', true);

  // ── 4. FSM II tab — write metadata (row 1) and data rows (3+) ────────────
  fillLaborTab(wb.sheet('FSM II'), payload.fsmII, payload.fsmIITotal, 'OT-17.75', false);

  // ── 5. Management Detail Hours tab ───────────────────────────────────────
  fillMgmtTab(wb.sheet('Management Detail Hours'), payload.mgmt);

  // ── 6. Cloud Services tab — update Allocation (col H) for each person ────
  // The Cloud Services tab has a static roster. We update col F (Quantity=1)
  // and col H (Allocation) based on whether a manager/FT associate appears
  // in this billing period. The formula =F*G*H auto-computes the total.
  // For simplicity: set F=1, H=allocationPct for mgmt; F=1, H=1 for FT field agents.
  // We leave the tab as-is — the template's existing data and formulas handle totals.
  // The Tie-Out sheet references Cloud Services!I1 = SUM(I3:I1048576) which is a formula.
  // We don't need to modify Cloud Services beyond keeping existing data intact.

  // ── 7. Return buffer ──────────────────────────────────────────────────────
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const buf: Buffer = await (wb as any).outputAsync();
  return buf;
}

// ── Labor tab filler (FSM I / FSM II) ─────────────────────────────────────

function fillLaborTab(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  sheet: any,
  rows: InvoiceRow[],
  tabTotal: number,
  otLabel: string,
  hasExtraCol: boolean, // true = FSM I (col T), false = FSM II
): void {
  const numCols = hasExtraCol ? 20 : 19;

  // Count unique associates
  const uniqueCount = new Set(rows.map(r => r.associateId.trim()).filter(Boolean)).size;

  // Total OT hours
  const otHours = rows
    .filter(r => r.basePayRate < 32) // OT rates: 15.53 and 17.75
    .reduce((s, r) => s + r.timeHours, 0);

  // Row 1 — metadata
  sheet.cell(1, 3).value(uniqueCount);  // C1 — unique associate count
  sheet.cell(1, 14).value(otHours);     // N1 — total OT hours
  sheet.cell(1, 15).value(otLabel);     // O1 — OT rate label
  sheet.cell(1, 18).value(tabTotal);    // R1 — total bill

  // Clear existing data rows (row 3+)
  // Find the last used row by scanning for data, then clear
  // We clear a generous range to handle templates with pre-existing data
  const clearLimit = 6000;
  for (let r = 3; r <= clearLimit; r++) {
    const cellVal = sheet.cell(r, 1).value();
    if (cellVal === null || cellVal === undefined || cellVal === '') {
      // Check a few more rows to handle gaps
      let hasMore = false;
      for (let lookahead = r + 1; lookahead <= r + 3; lookahead++) {
        if (sheet.cell(lookahead, 1).value()) { hasMore = true; break; }
      }
      if (!hasMore) break;
    }
    // Clear this row
    for (let c = 1; c <= numCols; c++) {
      sheet.cell(r, c).value(null);
    }
  }

  // Write fresh data rows starting at row 3
  rows.forEach((row, idx) => {
    const r = 3 + idx;
    sheet.cell(r, 1).value(row.week);          // A — Week
    sheet.cell(r, 2).value(row.employeeName);  // B — Employee Name
    sheet.cell(r, 3).value(row.associateId);   // C — Associate ID
    sheet.cell(r, 4).value(row.associateType); // D — Associate Type
    sheet.cell(r, 5).value(row.region);        // E — Employee Region
    sheet.cell(r, 6).value(row.district);      // F — Employee District
    sheet.cell(r, 7).value(row.market);        // G — Employee Market
    sheet.cell(r, 8).value(row.storeId);       // H — Store ID
    sheet.cell(r, 9).value(row.storeName);     // I — Store
    sheet.cell(r, 10).value(row.associateState); // J — Associate State
    sheet.cell(r, 11).value(row.storeState);   // K — Store State
    sheet.cell(r, 12).value(row.zipCode);      // L — Zip Code
    sheet.cell(r, 13).value(row.visitDate);    // M — Visit Date (string)
    sheet.cell(r, 14).value(row.timeHours);    // N — Time Hours
    sheet.cell(r, 15).value(row.basePayRate);  // O — Base Pay Rate
    sheet.cell(r, 16).value(parseFloat(fmt2(row.mu))); // P — MU
    sheet.cell(r, 17).value(row.payRateTotal); // Q — Pay Rate Total
    sheet.cell(r, 18).value(row.bill);         // R — Bill
    sheet.cell(r, 19).value(row.comments);     // S — Comments
    if (hasExtraCol) {
      sheet.cell(r, 20).value('');             // T — empty (FSM I only)
    }
  });
}

// ── Management Detail tab filler ──────────────────────────────────────────

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function fillMgmtTab(sheet: any, rows: MgmtRow[]): void {
  const numCols = 13;
  const grandTotalBill = rows.reduce((s, r) => s + r.totalBill, 0);

  // Row 1 — metadata
  sheet.cell(1, 3).value(rows.length);         // C1 — count of data rows
  sheet.cell(1, 10).value(grandTotalBill);     // J1 — total bill

  // Clear existing data rows
  const clearLimit = 500;
  for (let r = 3; r <= clearLimit; r++) {
    const cellVal = sheet.cell(r, 1).value();
    if (cellVal === null || cellVal === undefined || cellVal === '') {
      let hasMore = false;
      for (let lookahead = r + 1; lookahead <= r + 3; lookahead++) {
        if (sheet.cell(lookahead, 1).value()) { hasMore = true; break; }
      }
      if (!hasMore) break;
    }
    for (let c = 1; c <= numCols; c++) {
      sheet.cell(r, c).value(null);
    }
  }

  // Write fresh data rows starting at row 3
  rows.forEach((row, idx) => {
    const r = 3 + idx;
    sheet.cell(r, 1).value(row.week);           // A — Week
    sheet.cell(r, 2).value(row.associateName);  // B — Associate Name
    sheet.cell(r, 3).value(row.associateId);    // C — Associate ID
    sheet.cell(r, 4).value(row.title);          // D — Title
    sheet.cell(r, 5).value(row.associateState); // E — Associate State
    sheet.cell(r, 6).value(row.hours);          // F — Hours
    sheet.cell(r, 7).value(row.hourlyRate);     // G — Hourly Rate
    sheet.cell(r, 8).value(row.total);          // H — Total
    sheet.cell(r, 9).value(row.allocationPct);  // I — % Allocation
    sheet.cell(r, 10).value(row.totalBill);     // J — Total Bill
    sheet.cell(r, 11).value('');               // K — Comments (blank)
    sheet.cell(r, 12).value('');               // L
    sheet.cell(r, 13).value('');               // M
  });
}
