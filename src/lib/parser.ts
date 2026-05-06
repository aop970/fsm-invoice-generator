import * as XLSX from 'xlsx';
import type { ActivityRow, TimeOffRow, TermedPtoRow } from './types';
import { FSM_PROGRAM_PATTERNS } from './constants';

// ── Helpers ────────────────────────────────────────────────────────────────

function toStr(v: unknown): string {
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

function toNum(v: unknown): number {
  if (v === null || v === undefined || v === '') return 0;
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

/** Convert an Excel serial date to YYYY-MM-DD string, or pass through a string date. */
function excelDateToStr(v: unknown): string {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number') {
    const d = XLSX.SSF.parse_date_code(v);
    if (!d) return String(v);
    const mm = String(d.m).padStart(2, '0');
    const dd = String(d.d).padStart(2, '0');
    return `${d.y}-${mm}-${dd}`;
  }
  return toStr(v);
}

/** Normalize header names: lowercase, collapse whitespace, strip special chars */
function normalizeHeader(h: string): string {
  return h.toLowerCase().replace(/[\s\-_\/]+/g, ' ').trim();
}

/** Find a column index by trying multiple possible header names */
function findCol(headers: string[], candidates: string[]): number {
  for (const c of candidates) {
    const idx = headers.indexOf(c);
    if (idx !== -1) return idx;
  }
  // partial match fallback
  for (const c of candidates) {
    const idx = headers.findIndex(h => h.includes(c));
    if (idx !== -1) return idx;
  }
  return -1;
}

function readSheet(file: File): Promise<XLSX.WorkBook> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = e.target?.result;
      if (!data) { reject(new Error('No data')); return; }
      const wb = XLSX.read(data, { type: 'array', cellDates: false });
      resolve(wb);
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
}

// ── Activity Report Parser ─────────────────────────────────────────────────

export async function parseActivityReport(file: File): Promise<ActivityRow[]> {
  const wb = await readSheet(file);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rawRows = (XLSX.utils.sheet_to_json(ws, { defval: '', raw: true, header: 1 }) as unknown) as unknown[][];

  if (rawRows.length < 2) return [];

  // Find header row — look for a row containing "Employee Name" or "Associate ID"
  let headerRowIdx = 0;
  for (let i = 0; i < Math.min(10, rawRows.length); i++) {
    const rowStr = rawRows[i].map(c => toStr(c).toLowerCase()).join(' ');
    if (rowStr.includes('employee name') || rowStr.includes('associate id')) {
      headerRowIdx = i;
      break;
    }
  }

  const headers = (rawRows[headerRowIdx] as unknown[]).map(c => normalizeHeader(toStr(c)));
  const dataRows = rawRows.slice(headerRowIdx + 1);

  const colWeek         = findCol(headers, ['week', 'week number']);
  const colName         = findCol(headers, ['employee name', 'worker name', 'worker']);
  const colId           = findCol(headers, ['associate id', 'employee id', 'worker id']);
  const colType1        = findCol(headers, ['employee type 1', 'associate type', 'type 1', 'ft pt', 'ft/pt']);
  const colType3        = findCol(headers, ['employee type 3', 'program type', 'type 3', 'program']);
  const colRegion       = findCol(headers, ['employee region', 'region']);
  const colDistrict     = findCol(headers, ['employee district', 'district']);
  const colMarket       = findCol(headers, ['market name', 'employee market', 'market']);
  const colStoreId      = findCol(headers, ['client store id', 'store id', 'location id', 'store location']);
  const colStoreName    = findCol(headers, ['store', 'store name']);
  const colAssocState   = findCol(headers, ['employee state', 'associate state', 'state']);
  const colStoreState   = findCol(headers, ['store state']);
  const colZip          = findCol(headers, ['store zip', 'zip code', 'zip']);
  const colDate         = findCol(headers, ['date in', 'date', 'visit date', 'work date', 'activity date']);
  const colHours        = findCol(headers, ['time hours', 'time in hours', 'hours', 'total hours', 'reg hours']);
  const colComments     = findCol(headers, ['time type', 'pay code comments', 'pay code / comments', 'pay code', 'comments', 'comment', 'paycode']);

  const rows: ActivityRow[] = [];

  for (const raw of dataRows) {
    const r = raw as unknown[];
    const program = toStr(colType3 >= 0 ? r[colType3] : '');
    if (!program) continue;

    const programNorm = program.toLowerCase().trim();
    // Only process FSM-type rows
    if (!['fsm', 'fsm ii', 'fsm ii street'].includes(programNorm)) continue;

    const hours = toNum(colHours >= 0 ? r[colHours] : 0);
    // Skip zero-hour rows
    if (hours === 0) continue;

    rows.push({
      week:           toNum(colWeek >= 0 ? r[colWeek] : 0),
      employeeName:   toStr(colName >= 0 ? r[colName] : ''),
      associateId:    toStr(colId >= 0 ? r[colId] : ''),
      associateType:  toStr(colType1 >= 0 ? r[colType1] : 'FT').toUpperCase(),
      region:         toStr(colRegion >= 0 ? r[colRegion] : ''),
      district:       toStr(colDistrict >= 0 ? r[colDistrict] : ''),
      market:         toStr(colMarket >= 0 ? r[colMarket] : ''),
      storeId:        toStr(colStoreId >= 0 ? r[colStoreId] : ''),
      storeName:      toStr(colStoreName >= 0 ? r[colStoreName] : ''),
      associateState: toStr(colAssocState >= 0 ? r[colAssocState] : ''),
      storeState:     toStr(colStoreState >= 0 ? r[colStoreState] : ''),
      zipCode:        toStr(colZip >= 0 ? r[colZip] : ''),
      visitDate:      excelDateToStr(colDate >= 0 ? r[colDate] : ''),
      timeHours:      hours,
      comments:       toStr(colComments >= 0 ? r[colComments] : ''),
      program:        programNorm,
    });
  }

  return rows;
}

// ── Time Off Parser ────────────────────────────────────────────────────────

export async function parseTimeOff(file: File): Promise<TimeOffRow[]> {
  const wb = await readSheet(file);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rawRows = (XLSX.utils.sheet_to_json(ws, { defval: '', raw: true, header: 1 }) as unknown) as unknown[][];

  if (rawRows.length < 2) return [];

  let headerRowIdx = 0;
  for (let i = 0; i < Math.min(10, rawRows.length); i++) {
    const rowStr = (rawRows[i] as unknown[]).map(c => toStr(c).toLowerCase()).join(' ');
    if (rowStr.includes('worker') || rowStr.includes('program') || rowStr.includes('associate id')) {
      headerRowIdx = i;
      break;
    }
  }

  const headers = (rawRows[headerRowIdx] as unknown[]).map(c => normalizeHeader(toStr(c)));
  const dataRows = rawRows.slice(headerRowIdx + 1);

  const colId      = findCol(headers, ['associate id', 'employee id', 'worker id', 'id']);
  const colWorker  = findCol(headers, ['worker', 'employee name', 'name', 'worker name']);
  const colDate    = findCol(headers, ['time off date', 'date', 'visit date']);
  const colHours   = findCol(headers, ['total hours', 'hours', 'time hours', 'approved hours']);
  const colType    = findCol(headers, ['time off type', 'type', 'leave type', 'absence type']);
  const colProgram = findCol(headers, ['program', 'program name', 'client program']);
  const colStatus  = findCol(headers, ['status', 'approval status', 'state']);

  const rows: TimeOffRow[] = [];

  for (const raw of dataRows) {
    const r = raw as unknown[];
    const program = toStr(colProgram >= 0 ? r[colProgram] : '');
    const programNorm = program.toLowerCase().trim();

    const isFSM = FSM_PROGRAM_PATTERNS.some(p => programNorm.includes(p));
    if (!isFSM) continue;

    const status = toStr(colStatus >= 0 ? r[colStatus] : 'Approved');
    if (status.toLowerCase() !== 'approved') continue;

    const hours = toNum(colHours >= 0 ? r[colHours] : 0);
    if (hours <= 0) continue;

    rows.push({
      associateId:  toStr(colId >= 0 ? r[colId] : ''),
      employeeName: toStr(colWorker >= 0 ? r[colWorker] : ''),
      visitDate:    excelDateToStr(colDate >= 0 ? r[colDate] : ''),
      timeHours:    hours,
      comments:     toStr(colType >= 0 ? r[colType] : 'Time Off'),
      program:      programNorm,
      status,
    });
  }

  return rows;
}

// ── Termed PTO Parser ──────────────────────────────────────────────────────

export async function parseTermedPto(file: File): Promise<TermedPtoRow[]> {
  const wb = await readSheet(file);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rawRows = (XLSX.utils.sheet_to_json(ws, { defval: '', raw: true, header: 1 }) as unknown) as unknown[][];

  if (rawRows.length < 2) return [];

  let headerRowIdx = 0;
  for (let i = 0; i < Math.min(10, rawRows.length); i++) {
    const rowStr = (rawRows[i] as unknown[]).map(c => toStr(c).toLowerCase()).join(' ');
    if (rowStr.includes('worker') || rowStr.includes('employee id') || rowStr.includes('hours')) {
      headerRowIdx = i;
      break;
    }
  }

  const headers = (rawRows[headerRowIdx] as unknown[]).map(c => normalizeHeader(toStr(c)));
  const dataRows = rawRows.slice(headerRowIdx + 1);

  const colId      = findCol(headers, ['employee id', 'associate id', 'worker id', 'id']);
  const colWorker  = findCol(headers, ['worker', 'employee name', 'name', 'worker name']);
  const colHours   = findCol(headers, ['hours', 'total hours', 'pto hours', 'time hours']);
  const colProgram = findCol(headers, ['program', 'program name', 'client program']);

  const rows: TermedPtoRow[] = [];

  for (const raw of dataRows) {
    const r = raw as unknown[];
    const program = toStr(colProgram >= 0 ? r[colProgram] : '');
    const programNorm = program.toLowerCase().trim();

    const isFSM = FSM_PROGRAM_PATTERNS.some(p => programNorm.includes(p));
    if (!isFSM) continue;

    const hours = toNum(colHours >= 0 ? r[colHours] : 0);
    if (hours <= 0) continue;

    rows.push({
      associateId:  toStr(colId >= 0 ? r[colId] : ''),
      employeeName: toStr(colWorker >= 0 ? r[colWorker] : ''),
      timeHours:    hours,
      program:      programNorm,
    });
  }

  return rows;
}
