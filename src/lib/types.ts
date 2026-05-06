// ── Types ──────────────────────────────────────────────────────────────────

export interface ActivityRow {
  week: number;
  employeeName: string;
  associateId: string;
  associateType: string; // FT / PT
  region: string;
  district: string;
  market: string;
  storeId: string;       // numeric store identifier (Client Store ID)
  storeName: string;     // text name of store
  associateState: string; // employee's home state
  storeState: string;    // state where store is located
  zipCode: string;
  visitDate: string;
  timeHours: number;
  comments: string;
  program: string; // FSM | FSM II | FSM II Street | etc.
}

export interface TimeOffRow {
  associateId: string;
  employeeName: string;
  visitDate: string;
  timeHours: number;
  comments: string; // Time Off Type
  program: string;
  status: string;
}

export interface TermedPtoRow {
  associateId: string;
  employeeName: string;
  timeHours: number;
  program: string;
}

export interface InvoiceRow {
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

export interface MgmtRow {
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

/** Mode 1 — Weekly Finance Report (FSM I, FSM II, Management Detail) */
export interface WeeklySummary {
  mode: 'weekly';
  fsmI: InvoiceRow[];
  fsmII: InvoiceRow[];
  mgmt: MgmtRow[];
  fsmITotal: number;
  fsmIITotal: number;
  mgmtTotal: number;
  fieldLaborTotal: number;
  billingPeriod: string;
  weeks: number[];
}

/** Mode 2 — Bi-Weekly Client Invoice (includes Management Detail) */
export interface InvoiceSummary {
  mode: 'invoice';
  fsmI: InvoiceRow[];
  fsmII: InvoiceRow[];
  mgmt: MgmtRow[];
  fsmITotal: number;
  fsmIITotal: number;
  mgmtTotal: number;
  grandTotal: number;
  billingPeriod: string;
  weeks: number[];
}

export type AnyReport = WeeklySummary | InvoiceSummary;
