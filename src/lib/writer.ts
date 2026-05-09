/**
 * writer.ts — Client-side workbook builder (template-fill approach)
 *
 * Packages data into a JSON payload and POSTs to /generate (proxied to the
 * Node server at localhost:3001). The server fills FSM_Invoice_TEMPLATE.xlsx
 * using xlsx-populate and returns the buffer; this module triggers a browser
 * download from the response blob.
 *
 * Exported signatures are identical to the previous SheetJS implementation
 * so App.tsx requires zero changes.
 */

import type { InvoiceSummary, WeeklySummary } from './types';
import { buildCoverPeriodStr, getWeekEndDate } from './transform';

// ── Public types ──────────────────────────────────────────────────────────────

export interface NhfParams {
  count: number;
  rate: number;   // 285 or 395
  label: string;  // e.g. "2020 Transfer", "In Talent Network", "New Hire"
}

/**
 * Opaque payload object returned by buildWorkbook / buildWeeklyWorkbook.
 * App.tsx receives this and passes it directly to downloadWorkbook — it never
 * inspects the contents, so the concrete shape is an implementation detail.
 */
export interface WorkbookPayload {
  _payload: GeneratePayload;
  _filename: string;
}

// ── Internal payload shape (matches server/filler.ts GeneratePayload) ─────────

interface GeneratePayload {
  mode: 'weekly' | 'invoice';
  invoiceName: string;
  periodStr: string;
  invoiceDate: string;   // ISO date string
  dueDate: string;       // ISO date string
  poNumber: string;
  fsmI: unknown[];
  fsmII: unknown[];
  mgmt: unknown[];
  fsmITotal: number;
  fsmIITotal: number;
  mgmtTotal: number;
  nhf: NhfParams | null;
  weeks: number[];
}

// ── Helpers ───────────────────────────────────────────────────────────────────

const PO_NUMBER = 'T26C31H000162';

/** Build the invoice name: "FSM26-W15" (single) or "FSM26-W15-16" (bi-weekly) */
function buildInvoiceName(weeks: number[]): string {
  if (weeks.length === 0) return 'FSM-Invoice';
  const year = new Date().getFullYear().toString().slice(-2);
  const w1 = String(weeks[0]).padStart(2, '0');
  if (weeks.length === 1) return `FSM${year}-W${w1}`;
  const w2 = String(weeks[weeks.length - 1]).padStart(2, '0');
  return `FSM${year}-W${w1}-${w2}`;
}

/** Format a Date as ISO date string (YYYY-MM-DD) in local time. */
function toISODate(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

// ── Public API ────────────────────────────────────────────────────────────────

/** Generate weekly filename: FSM[YY]-W[WW]_Weekly.xlsx */
export function buildWeeklyFilename(weeks: number[]): string {
  if (weeks.length === 0) return 'FSM-Weekly.xlsx';
  const year = new Date().getFullYear().toString().slice(-2);
  const w = String(weeks[0]).padStart(2, '0');
  return `FSM${year}-W${w}_Weekly.xlsx`;
}

/** Generate invoice filename: FSM[YY]-W[WW].xlsx */
export function buildFilename(weeks: number[]): string {
  if (weeks.length === 0) return 'FSM-Invoice.xlsx';
  return buildInvoiceName(weeks) + '.xlsx';
}

/**
 * Build the weekly workbook payload (Mode 1 — Finance Report).
 * Returns an opaque WorkbookPayload; pass to downloadWorkbook to trigger download.
 */
export function buildWeeklyWorkbook(
  summary: WeeklySummary,
  nhf: NhfParams | null = null,
): WorkbookPayload {
  const allRows = [...summary.fsmI, ...summary.fsmII];
  const periodStr = buildCoverPeriodStr(allRows);
  const weekEndDate = getWeekEndDate(allRows);
  const invoiceName = buildInvoiceName(summary.weeks);

  let invoiceDate = '';
  let dueDate = '';
  if (weekEndDate) {
    const inv = new Date(weekEndDate);
    inv.setDate(inv.getDate() + 2);
    const due = new Date(inv);
    due.setDate(due.getDate() + 30);
    invoiceDate = toISODate(inv);
    dueDate = toISODate(due);
  }

  const payload: GeneratePayload = {
    mode: 'weekly',
    invoiceName,
    periodStr,
    invoiceDate,
    dueDate,
    poNumber: PO_NUMBER,
    fsmI: summary.fsmI,
    fsmII: summary.fsmII,
    mgmt: summary.mgmt,
    fsmITotal: summary.fsmITotal,
    fsmIITotal: summary.fsmIITotal,
    mgmtTotal: summary.mgmtTotal,
    nhf,
    weeks: summary.weeks,
  };

  return { _payload: payload, _filename: buildWeeklyFilename(summary.weeks) };
}

/**
 * Build the invoice workbook payload (Mode 2 — Client Invoice).
 * Returns an opaque WorkbookPayload; pass to downloadWorkbook to trigger download.
 */
export function buildWorkbook(
  summary: InvoiceSummary,
  nhf: NhfParams | null = null,
): WorkbookPayload {
  const allRows = [...summary.fsmI, ...summary.fsmII];
  const periodStr = buildCoverPeriodStr(allRows);
  const weekEndDate = getWeekEndDate(allRows);
  const invoiceName = buildInvoiceName(summary.weeks);

  let invoiceDate = '';
  let dueDate = '';
  if (weekEndDate) {
    const inv = new Date(weekEndDate);
    inv.setDate(inv.getDate() + 2);
    const due = new Date(inv);
    due.setDate(due.getDate() + 30);
    invoiceDate = toISODate(inv);
    dueDate = toISODate(due);
  }

  const payload: GeneratePayload = {
    mode: 'invoice',
    invoiceName,
    periodStr,
    invoiceDate,
    dueDate,
    poNumber: PO_NUMBER,
    fsmI: summary.fsmI,
    fsmII: summary.fsmII,
    mgmt: summary.mgmt,
    fsmITotal: summary.fsmITotal,
    fsmIITotal: summary.fsmIITotal,
    mgmtTotal: summary.mgmtTotal,
    nhf,
    weeks: summary.weeks,
  };

  return { _payload: payload, _filename: buildFilename(summary.weeks) };
}

/**
 * POST the payload to /generate and trigger a browser download from the response blob.
 * Returns a Promise<void> — the caller in App.tsx does not await it, but TypeScript
 * treats Promise<void> as assignable to void so no signature change is needed.
 */
export function downloadWorkbook(wb: WorkbookPayload, filename: string): Promise<void> {
  return (async () => {
    const res = await fetch('/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(wb._payload),
    });

    if (!res.ok) {
      const msg = await res.text().catch(() => res.statusText);
      throw new Error(`Server error ${res.status}: ${msg}`);
    }

    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  })();
}
