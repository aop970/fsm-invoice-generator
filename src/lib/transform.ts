import type { ActivityRow, TimeOffRow, TermedPtoRow, InvoiceRow, MgmtRow, InvoiceSummary, WeeklySummary } from './types';
import { RATES, MARKUP, OT_COMMENT_PATTERNS, MGMT_TABLE } from './constants';

// ── Helpers ────────────────────────────────────────────────────────────────

function round2(n: number): number {
  return Math.round(n * 100) / 100;
}

/**
 * Calculate ISO 8601 week number for a given date.
 * Weeks run Monday–Sunday. Jan 4 is always in week 1.
 * April 12, 2026 (Sunday) → 15; April 19, 2026 (Sunday) → 16.
 */
function getISOWeek(date: Date): number {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7; // treat Sunday as 7
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d.getTime() - yearStart.getTime()) / 86400000) + 1) / 7);
}

/**
 * Derive ISO week number from a YYYY-MM-DD string.
 * Returns 0 if the string is empty/invalid.
 */
function isoWeekFromDateStr(dateStr: string): number {
  if (!dateStr) return 0;
  const parts = dateStr.split('-').map(Number);
  if (parts.length < 3 || parts.some(isNaN)) return 0;
  return getISOWeek(new Date(parts[0], parts[1] - 1, parts[2]));
}

function isOT(comments: string): boolean {
  const c = comments.toLowerCase().trim();
  return OT_COMMENT_PATTERNS.some(p => c.includes(p));
}

/** Map program string to 'FSM_I' | 'FSM_II' */
function programToTab(program: string): 'FSM_I' | 'FSM_II' | null {
  const p = program.toLowerCase().trim();
  if (p === 'fsm') return 'FSM_I';
  if (p === 'fsm ii' || p === 'fsm ii street') return 'FSM_II';
  return null;
}

function calcRow(
  baseRate: number,
  associateType: string,
  timeHours: number,
): { mu: number; payRateTotal: number; bill: number } {
  const muPct = associateType.toUpperCase().startsWith('PT') ? MARKUP.PT : MARKUP.FT;
  const mu = round2(baseRate * muPct);
  const payRateTotal = round2(baseRate + mu);
  const bill = round2(timeHours * payRateTotal);
  return { mu, payRateTotal, bill };
}

// ── Transform ──────────────────────────────────────────────────────────────

/** Mode 1: single-week Finance Report — FSM I, FSM II, and Management Detail (one row per manager) */
export function buildWeeklyReport(
  activityRows: ActivityRow[],
  timeOffRows: TimeOffRow[],
): WeeklySummary {
  // Re-stamp ISO week numbers from visitDate (ignore whatever week value came from the source file)
  for (const r of activityRows) {
    const w = isoWeekFromDateStr(r.visitDate);
    if (w > 0) r.week = w;
  }
  for (const t of timeOffRows) {
    const w = isoWeekFromDateStr(t.visitDate);
    if (w > 0) (t as ActivityRow & TimeOffRow).week = w;
  }

  // Build associate → tab map
  const assocTabMap = new Map<string, 'FSM_I' | 'FSM_II'>();
  for (const r of activityRows) {
    const tab = programToTab(r.program);
    if (tab) assocTabMap.set(r.associateId.trim(), tab);
  }

  // Detect weeks from ISO-stamped rows
  const weekNums = [...new Set(activityRows.map(r => r.week))].filter(w => w > 0).sort((a, b) => a - b);

  const fsmIRows: InvoiceRow[] = [];
  const fsmIIRows: InvoiceRow[] = [];

  // Activity rows
  for (const r of activityRows) {
    const tab = programToTab(r.program);
    if (!tab) continue;
    const rates = RATES[tab];
    const baseRate = isOT(r.comments) ? rates.ot : rates.regular;
    const { mu, payRateTotal, bill } = calcRow(baseRate, r.associateType, r.timeHours);
    const row: InvoiceRow = {
      week: r.week,
      employeeName: r.employeeName,
      associateId: r.associateId,
      associateType: r.associateType,
      region: r.region,
      district: r.district,
      market: r.market,
      storeId: r.storeId,
      storeName: r.storeName,
      associateState: r.associateState,
      storeState: r.storeState,
      zipCode: r.zipCode,
      visitDate: r.visitDate,
      timeHours: r.timeHours,
      basePayRate: baseRate,
      mu,
      payRateTotal,
      bill,
      comments: r.comments,
    };
    if (tab === 'FSM_I') fsmIRows.push(row);
    else fsmIIRows.push(row);
  }

  // Time Off rows — use ISO week from visitDate
  const weekDateMap = buildWeekDateMap(activityRows);
  for (const t of timeOffRows) {
    if (!t.associateId && !t.employeeName) continue;
    const tab = assocTabMap.get(t.associateId.trim()) ?? 'FSM_I';
    const rates = RATES[tab];
    const baseRate = rates.regular;
    const actRow = activityRows.find(a => a.associateId.trim() === t.associateId.trim());
    const associateType = actRow?.associateType ?? 'FT';
    const { mu, payRateTotal, bill } = calcRow(baseRate, associateType, t.timeHours);
    // Prefer direct ISO week from date; fall back to range matching
    let week = isoWeekFromDateStr(t.visitDate);
    if (week === 0) week = assignWeek(t.visitDate, weekDateMap, weekNums);
    const row: InvoiceRow = {
      week,
      employeeName: t.employeeName || actRow?.employeeName || '',
      associateId: t.associateId,
      associateType,
      region: actRow?.region ?? '',
      district: actRow?.district ?? '',
      market: actRow?.market ?? '',
      storeId: actRow?.storeId ?? '',
      storeName: actRow?.storeName ?? '',
      associateState: actRow?.associateState ?? '',
      storeState: actRow?.storeState ?? '',
      zipCode: actRow?.zipCode ?? '',
      visitDate: t.visitDate,
      timeHours: t.timeHours,
      basePayRate: baseRate,
      mu,
      payRateTotal,
      bill,
      comments: t.comments || 'Time Off',
    };
    if (tab === 'FSM_I') fsmIRows.push(row);
    else fsmIIRows.push(row);
  }

  const sortRows = (rows: InvoiceRow[]) =>
    rows.sort((a, b) => {
      if (a.week !== b.week) return a.week - b.week;
      return a.employeeName.localeCompare(b.employeeName);
    });

  sortRows(fsmIRows);
  sortRows(fsmIIRows);

  const fsmITotal = round2(fsmIRows.reduce((s, r) => s + r.bill, 0));
  const fsmIITotal = round2(fsmIIRows.reduce((s, r) => s + r.bill, 0));

  // ── Management Detail (Mode 1 — one row per manager, 40 hrs fixed) ────────
  const weekDateMapForMgmt = buildWeekDateMap(activityRows);
  const mgmtTimeOffDeductions = buildMgmtTimeOffDeductions(timeOffRows, weekNums, weekDateMapForMgmt);

  // In Mode 1 there is exactly one week per manager (use the single week, or weekNums[0])
  const singleWeek = weekNums.length > 0 ? weekNums[0] : 1;
  const mgmtRows: MgmtRow[] = MGMT_TABLE.map(m => {
    const deduction = mgmtTimeOffDeductions.get(`${m.associateId}:${singleWeek}`) ?? 0;
    const hours = Math.max(0, 40 - deduction);
    const total = round2(hours * m.hourlyRate);
    const totalBill = round2(total * m.allocationPct);
    return {
      week: singleWeek,
      associateName: m.name,
      associateId: m.associateId,
      title: m.title,
      associateState: m.associateState,
      hours,
      hourlyRate: m.hourlyRate,
      total,
      allocationPct: m.allocationPct,
      totalBill,
    };
  });

  const mgmtTotal = round2(mgmtRows.reduce((s, r) => s + r.totalBill, 0));
  const fieldLaborTotal = round2(fsmITotal + fsmIITotal);
  const billingPeriod = buildBillingPeriodStr(weekNums, activityRows);

  return {
    mode: 'weekly',
    fsmI: fsmIRows,
    fsmII: fsmIIRows,
    mgmt: mgmtRows,
    fsmITotal,
    fsmIITotal,
    mgmtTotal,
    fieldLaborTotal,
    billingPeriod,
    weeks: weekNums,
  };
}

/** Mode 2: bi-weekly Client Invoice — FSM I, FSM II, Management Detail */
export function buildInvoice(
  activityW1: ActivityRow[],
  activityW2: ActivityRow[],
  timeOffW1: TimeOffRow[],
  timeOffW2: TimeOffRow[],
  termedPto: TermedPtoRow[],
): InvoiceSummary {
  const allActivity = [...activityW1, ...activityW2];

  // Re-stamp ISO week numbers from visitDate on all activity rows
  for (const r of allActivity) {
    const w = isoWeekFromDateStr(r.visitDate);
    if (w > 0) r.week = w;
  }

  // Build associate → tab map from activity rows
  const assocTabMap = new Map<string, 'FSM_I' | 'FSM_II'>();
  for (const r of allActivity) {
    const tab = programToTab(r.program);
    if (tab) assocTabMap.set(r.associateId.trim(), tab);
  }

  // Detect weeks from ISO-stamped rows
  const weekNums = [...new Set(allActivity.map(r => r.week))].filter(w => w > 0).sort((a, b) => a - b);
  const lastWeek = weekNums.length > 0 ? weekNums[weekNums.length - 1] : 1;

  // ── 1. Activity rows ───────────────────────────────────────────────────
  const fsmIRows: InvoiceRow[] = [];
  const fsmIIRows: InvoiceRow[] = [];

  for (const r of allActivity) {
    const tab = programToTab(r.program);
    if (!tab) continue;

    const rates = RATES[tab];
    const baseRate = isOT(r.comments) ? rates.ot : rates.regular;
    const { mu, payRateTotal, bill } = calcRow(baseRate, r.associateType, r.timeHours);

    const row: InvoiceRow = {
      week: r.week,
      employeeName: r.employeeName,
      associateId: r.associateId,
      associateType: r.associateType,
      region: r.region,
      district: r.district,
      market: r.market,
      storeId: r.storeId,
      storeName: r.storeName,
      associateState: r.associateState,
      storeState: r.storeState,
      zipCode: r.zipCode,
      visitDate: r.visitDate,
      timeHours: r.timeHours,
      basePayRate: baseRate,
      mu,
      payRateTotal,
      bill,
      comments: r.comments,
    };

    if (tab === 'FSM_I') fsmIRows.push(row);
    else fsmIIRows.push(row);
  }

  // ── 2. Time Off rows ───────────────────────────────────────────────────
  const allTimeOff = [...timeOffW1, ...timeOffW2];

  // Build date→week map for fallback assignment
  const weekDateMap = buildWeekDateMap(allActivity);

  for (const t of allTimeOff) {
    if (!t.associateId && !t.employeeName) continue;

    // Determine tab for this employee
    const tab = assocTabMap.get(t.associateId.trim()) ?? 'FSM_I';
    const rates = RATES[tab];
    const baseRate = rates.regular;

    // Lookup associate type from activity rows
    const actRow = allActivity.find(a => a.associateId.trim() === t.associateId.trim());
    const associateType = actRow?.associateType ?? 'FT';

    const { mu, payRateTotal, bill } = calcRow(baseRate, associateType, t.timeHours);

    // Assign week number: prefer direct ISO week from date, fall back to range matching
    let week = isoWeekFromDateStr(t.visitDate);
    if (week === 0) week = assignWeek(t.visitDate, weekDateMap, weekNums);

    const row: InvoiceRow = {
      week,
      employeeName: t.employeeName || actRow?.employeeName || '',
      associateId: t.associateId,
      associateType,
      region: actRow?.region ?? '',
      district: actRow?.district ?? '',
      market: actRow?.market ?? '',
      storeId: actRow?.storeId ?? '',
      storeName: actRow?.storeName ?? '',
      associateState: actRow?.associateState ?? '',
      storeState: actRow?.storeState ?? '',
      zipCode: actRow?.zipCode ?? '',
      visitDate: t.visitDate,
      timeHours: t.timeHours,
      basePayRate: baseRate,
      mu,
      payRateTotal,
      bill,
      comments: t.comments || 'Time Off',
    };

    if (tab === 'FSM_I') fsmIRows.push(row);
    else fsmIIRows.push(row);
  }

  // ── 3. Termed PTO rows ─────────────────────────────────────────────────
  for (const t of termedPto) {
    if (!t.employeeName && !t.associateId) continue;

    const tab = assocTabMap.get(t.associateId.trim()) ?? 'FSM_I';
    const rates = RATES[tab];
    const baseRate = rates.regular;

    const actRow = allActivity.find(a => a.associateId.trim() === t.associateId.trim());
    const associateType = actRow?.associateType ?? 'FT';

    const { mu, payRateTotal, bill } = calcRow(baseRate, associateType, t.timeHours);

    const row: InvoiceRow = {
      week: lastWeek,
      employeeName: t.employeeName,
      associateId: t.associateId,
      associateType,
      region: actRow?.region ?? '',
      district: actRow?.district ?? '',
      market: actRow?.market ?? '',
      storeId: actRow?.storeId ?? '',
      storeName: actRow?.storeName ?? '',
      associateState: actRow?.associateState ?? '',
      storeState: actRow?.storeState ?? '',
      zipCode: actRow?.zipCode ?? '',
      visitDate: '',
      timeHours: t.timeHours,
      basePayRate: baseRate,
      mu,
      payRateTotal,
      bill,
      comments: 'Termed PTO',
    };

    if (tab === 'FSM_I') fsmIRows.push(row);
    else fsmIIRows.push(row);
  }

  // ── 4. Sort ────────────────────────────────────────────────────────────
  const sortRows = (rows: InvoiceRow[]) =>
    rows.sort((a, b) => {
      if (a.week !== b.week) return a.week - b.week;
      return a.employeeName.localeCompare(b.employeeName);
    });

  sortRows(fsmIRows);
  sortRows(fsmIIRows);

  // ── 5. Management Detail ───────────────────────────────────────────────
  // Build time-off deductions for management
  const mgmtTimeOffDeductions = buildMgmtTimeOffDeductions(allTimeOff, weekNums, weekDateMap);

  const mgmtRows: MgmtRow[] = [];
  for (const wk of weekNums.length > 0 ? weekNums : [1]) {
    for (const m of MGMT_TABLE) {
      const deduction = mgmtTimeOffDeductions.get(`${m.associateId}:${wk}`) ?? 0;
      const hours = Math.max(0, 40 - deduction);
      const total = round2(hours * m.hourlyRate);
      const totalBill = round2(total * m.allocationPct);
      mgmtRows.push({
        week: wk,
        associateName: m.name,
        associateId: m.associateId,
        title: m.title,
        associateState: m.associateState,
        hours,
        hourlyRate: m.hourlyRate,
        total,
        allocationPct: m.allocationPct,
        totalBill,
      });
    }
  }

  // ── 6. Totals ──────────────────────────────────────────────────────────
  const fsmITotal = round2(fsmIRows.reduce((s, r) => s + r.bill, 0));
  const fsmIITotal = round2(fsmIIRows.reduce((s, r) => s + r.bill, 0));
  const mgmtTotal = round2(mgmtRows.reduce((s, r) => s + r.totalBill, 0));
  const grandTotal = round2(fsmITotal + fsmIITotal + mgmtTotal);

  // ── 7. Billing period string ───────────────────────────────────────────
  const billingPeriod = buildBillingPeriodStr(weekNums, allActivity);

  return {
    mode: 'invoice' as const,
    fsmI: fsmIRows,
    fsmII: fsmIIRows,
    mgmt: mgmtRows,
    fsmITotal,
    fsmIITotal,
    mgmtTotal,
    grandTotal,
    billingPeriod,
    weeks: weekNums,
  };
}

// ── Utilities ──────────────────────────────────────────────────────────────

interface WeekDateMap {
  [week: number]: { min: string; max: string };
}

function buildWeekDateMap(activity: ActivityRow[]): WeekDateMap {
  const map: WeekDateMap = {};
  for (const r of activity) {
    if (!r.week || !r.visitDate) continue;
    if (!map[r.week]) map[r.week] = { min: r.visitDate, max: r.visitDate };
    if (r.visitDate < map[r.week].min) map[r.week].min = r.visitDate;
    if (r.visitDate > map[r.week].max) map[r.week].max = r.visitDate;
  }
  return map;
}

function assignWeek(date: string, weekDateMap: WeekDateMap, weekNums: number[]): number {
  if (!date) return weekNums[0] ?? 1;
  for (const wk of weekNums) {
    const range = weekDateMap[wk];
    if (!range) continue;
    if (date >= range.min && date <= range.max) return wk;
  }
  // If date is before first week, assign to first week; after last, assign to last
  if (weekNums.length > 0) {
    const firstRange = weekDateMap[weekNums[0]];
    if (firstRange && date < firstRange.min) return weekNums[0];
    return weekNums[weekNums.length - 1];
  }
  return 1;
}

function buildMgmtTimeOffDeductions(
  allTimeOff: TimeOffRow[],
  weekNums: number[],
  weekDateMap: WeekDateMap,
): Map<string, number> {
  const deductions = new Map<string, number>();
  for (const t of allTimeOff) {
    // Check if this is a management person by name matching
    const mgmtMember = MGMT_TABLE.find(m =>
      m.name.toLowerCase() === t.employeeName.toLowerCase() ||
      m.associateId.toLowerCase() === t.associateId.toLowerCase()
    );
    if (!mgmtMember) continue;

    const wk = assignWeek(t.visitDate, weekDateMap, weekNums);
    const key = `${mgmtMember.associateId}:${wk}`;
    deductions.set(key, (deductions.get(key) ?? 0) + t.timeHours);
  }
  return deductions;
}

function buildBillingPeriodStr(weekNums: number[], activity: ActivityRow[]): string {
  if (weekNums.length === 0) return 'Unknown Period';

  const w1 = weekNums[0];
  const w2 = weekNums[weekNums.length - 1];
  const weekLabel = w1 === w2 ? `W${w1}` : `W${w1}-${w2}`;

  // Get date range from activity
  const dates = activity.map(r => r.visitDate).filter(Boolean).sort();
  if (dates.length === 0) return weekLabel;

  const startStr = dates[0];
  const endStr = dates[dates.length - 1];

  const fmtLong = (iso: string) => {
    const [year, month, day] = iso.split('-').map(Number);
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                    'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    return `${months[month - 1]} ${day}, ${year}`;
  };

  return `${weekLabel} | ${fmtLong(startStr)} – ${fmtLong(endStr)}`;
}

/**
 * Return the billing period as a display string suitable for the cover tab
 * (e.g. "04/06/2026 - 04/12/2026"), derived from the min/max visitDate across all rows.
 * Accepts any array of objects that have a `visitDate` string field.
 */
export function buildCoverPeriodStr(rows: { visitDate: string }[]): string {
  const dates = rows.map(r => r.visitDate).filter(Boolean).sort();
  if (dates.length === 0) return '';
  const fmtMDY = (iso: string) => {
    const [year, month, day] = iso.split('-').map(Number);
    return `${String(month).padStart(2, '0')}/${String(day).padStart(2, '0')}/${year}`;
  };
  return `${fmtMDY(dates[0])} - ${fmtMDY(dates[dates.length - 1])}`;
}

/**
 * Return the week-end date as a Date object for the latest visitDate in the given rows.
 * Used by the cover tab to compute Invoice Date (latest date + 2 days) and Due Date (+30 days).
 * Accepts any array of objects that have a `visitDate` string field.
 */
export function getWeekEndDate(rows: { visitDate: string }[]): Date | null {
  const dates = rows.map(r => r.visitDate).filter(Boolean).sort();
  if (dates.length === 0) return null;
  const latestStr = dates[dates.length - 1];
  const [year, month, day] = latestStr.split('-').map(Number);
  return new Date(year, month - 1, day);
}
