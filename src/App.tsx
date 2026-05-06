import { useState, useCallback, useRef } from 'react';
import { DropZone } from './components/DropZone';
import { StatusLog } from './components/StatusLog';
import type { LogEntry } from './components/StatusLog';
import { SummaryPanel } from './components/SummaryPanel';
import { parseActivityReport, parseTimeOff, parseTermedPto } from './lib/parser';
import { buildWeeklyReport, buildInvoice } from './lib/transform';
import {
  buildWeeklyWorkbook, buildWeeklyFilename,
  buildWorkbook, buildFilename,
  downloadWorkbook,
} from './lib/writer';
import type { AnyReport } from './lib/types';

type AppMode = 'weekly' | 'invoice';
type Step = 'idle' | 'processing' | 'done' | 'error';

let logCounter = 0;

/** ISO 8601 week number — weeks run Mon–Sun; Jan 4 is always in week 1. */
function getISOWeek(date: Date): number {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7; // treat Sunday as 7
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d.getTime() - yearStart.getTime()) / 86400000) + 1) / 7);
}

/** Extract "WE MM.DD" from a filename and compute ISO week number */
function detectWeekFromFilename(name: string): { label: string; weekNum: number | null } | null {
  const match = name.match(/WE\s+(\d{2})\.(\d{2})/i);
  if (!match) return null;
  const label = `${match[1]}/${match[2]}`;
  const year = new Date().getFullYear();
  const d = new Date(year, parseInt(match[1]) - 1, parseInt(match[2]));
  return { label, weekNum: getISOWeek(d) };
}

function PeriodPreview({ actW1, actW2, mode }: {
  actW1: File | null;
  actW2: File | null;
  mode: AppMode;
}) {
  const info1 = actW1 ? detectWeekFromFilename(actW1.name) : null;
  const info2 = actW2 ? detectWeekFromFilename(actW2.name) : null;

  if (mode === 'weekly') {
    if (!info1) return null;
    const wLabel = info1.weekNum ? `W${String(info1.weekNum).padStart(2, '0')}` : '';
    return (
      <div className="text-xs text-slate-400 px-1">
        Detected week ending: <span className="text-blue-300">{info1.label}</span>
        {wLabel && <span className="ml-1 text-slate-500">({wLabel})</span>}
      </div>
    );
  }

  // Invoice mode
  if (!info1 && !info2) return null;
  const parts: string[] = [];
  if (info1) {
    const w = info1.weekNum ? ` (W${String(info1.weekNum).padStart(2, '0')})` : '';
    parts.push(`Week 1: ${info1.label}${w}`);
  }
  if (info2) {
    const w = info2.weekNum ? ` (W${String(info2.weekNum).padStart(2, '0')})` : '';
    parts.push(`Week 2: ${info2.label}${w}`);
  }
  return (
    <div className="text-xs text-slate-400 px-1 space-y-0.5">
      {parts.map((p, i) => (
        <div key={i}>
          Detected <span className="text-blue-300">{p}</span>
        </div>
      ))}
    </div>
  );
}

export default function App() {
  const [mode, setMode] = useState<AppMode>('weekly');

  // Weekly mode files
  const [wActW1, setWActW1] = useState<File | null>(null);
  const [wToW1, setWToW1]   = useState<File | null>(null);

  // Invoice mode files
  const [iActW1, setIActW1] = useState<File | null>(null);
  const [iActW2, setIActW2] = useState<File | null>(null);
  const [iToW1, setIToW1]   = useState<File | null>(null);
  const [iToW2, setIToW2]   = useState<File | null>(null);
  const [iPto,  setIPto]    = useState<File | null>(null);

  const [step, setStep]       = useState<Step>('idle');
  const [logs, setLogs]       = useState<LogEntry[]>([]);
  const [summary, setSummary] = useState<AnyReport | null>(null);
  const [filename, setFilename] = useState<string>('FSM-Report.xlsx');

  const summaryRef = useRef<HTMLDivElement | null>(null);

  const log = useCallback((text: string, type: LogEntry['type'] = 'info') => {
    setLogs(prev => [...prev, { id: ++logCounter, text, type }]);
  }, []);

  const canGenerateWeekly  = wActW1 !== null && wToW1 !== null;
  const canGenerateInvoice = iActW1 !== null && iActW2 !== null && iToW1 !== null && iToW2 !== null;
  const canGenerate = mode === 'weekly' ? canGenerateWeekly : canGenerateInvoice;

  const handleModeChange = (m: AppMode) => {
    if (m === mode) return;
    setMode(m);
    setStep('idle');
    setLogs([]);
    setSummary(null);
  };

  const handleGenerate = async () => {
    if (!canGenerate) return;

    setStep('processing');
    setSummary(null);
    setLogs([]);

    try {
      if (mode === 'weekly') {
        log('Parsing Activity Report…', 'progress');
        const aw1 = await parseActivityReport(wActW1!);
        log(`Activity: ${aw1.length} FSM rows found`, 'success');

        log('Parsing Time Off Report…', 'progress');
        const tw1 = await parseTimeOff(wToW1!);
        log(`Time Off: ${tw1.length} approved FSM rows found`, 'success');

        log('Applying business rules and computing rates…', 'progress');
        const rpt = buildWeeklyReport(aw1, tw1);
        log(`FSM I: ${rpt.fsmI.length} rows  |  FSM II: ${rpt.fsmII.length} rows`, 'info');

        log('Building Excel workbook…', 'progress');
        const wb = buildWeeklyWorkbook(rpt);
        const fn = buildWeeklyFilename(rpt.weeks);
        setFilename(fn);
        void wb;

        log(`Ready: ${fn}`, 'success');
        setStep('done');
        setSummary(rpt);

      } else {
        log('Parsing Activity Report — Week 1…', 'progress');
        const aw1 = await parseActivityReport(iActW1!);
        log(`Activity W1: ${aw1.length} FSM rows found`, 'success');

        log('Parsing Activity Report — Week 2…', 'progress');
        const aw2 = await parseActivityReport(iActW2!);
        log(`Activity W2: ${aw2.length} FSM rows found`, 'success');

        log('Parsing Time Off — Week 1…', 'progress');
        const tw1 = await parseTimeOff(iToW1!);
        log(`Time Off W1: ${tw1.length} approved FSM rows found`, 'success');

        log('Parsing Time Off — Week 2…', 'progress');
        const tw2 = await parseTimeOff(iToW2!);
        log(`Time Off W2: ${tw2.length} approved FSM rows found`, 'success');

        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        let ptoRows: any[] = [];
        if (iPto) {
          log('Parsing Termed PTO…', 'progress');
          ptoRows = await parseTermedPto(iPto);
          log(`Termed PTO: ${ptoRows.length} FSM rows found`, 'success');
        }

        log('Applying business rules and computing rates…', 'progress');
        const inv = buildInvoice(aw1, aw2, tw1, tw2, ptoRows);
        log(`FSM I: ${inv.fsmI.length} rows  |  FSM II: ${inv.fsmII.length} rows  |  Mgmt: ${inv.mgmt.length} rows`, 'info');

        log('Building Excel workbook…', 'progress');
        const wb = buildWorkbook(inv);
        const fn = buildFilename(inv.weeks);
        setFilename(fn);
        void wb;

        log(`Ready: ${fn}`, 'success');
        setStep('done');
        setSummary(inv);
      }

      setTimeout(() => summaryRef.current?.scrollIntoView({ behavior: 'smooth' }), 100);
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      log(`Error: ${msg}`, 'error');
      setStep('error');
    }
  };

  const handleDownload = useCallback(() => {
    if (!summary) return;
    if (summary.mode === 'weekly') {
      const wb = buildWeeklyWorkbook(summary);
      downloadWorkbook(wb, filename);
    } else {
      const wb = buildWorkbook(summary);
      downloadWorkbook(wb, filename);
    }
  }, [summary, filename]);

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 font-sans">
      {/* Header */}
      <header className="border-b border-slate-800 bg-slate-900/80 backdrop-blur sticky top-0 z-10">
        <div className="max-w-5xl mx-auto px-6 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 rounded-lg bg-blue-600 flex items-center justify-center text-sm font-bold">F</div>
            <div>
              <h1 className="text-base font-bold text-white leading-none">FSM Invoice Generator</h1>
              <p className="text-xs text-slate-400 mt-0.5">Samsung FSM Program</p>
            </div>
          </div>
          <div className="text-xs text-slate-500">v3.0</div>
        </div>
      </header>

      {/* Main */}
      <main className="max-w-5xl mx-auto px-6 py-8 space-y-8">

        {/* Mode Toggle */}
        <section>
          <div className="flex rounded-xl overflow-hidden border border-slate-700 w-full sm:w-auto sm:inline-flex">
            <button
              onClick={() => handleModeChange('weekly')}
              className={[
                'flex-1 sm:flex-none px-6 py-3 text-sm font-semibold transition-colors',
                mode === 'weekly'
                  ? 'bg-blue-600 text-white'
                  : 'bg-slate-800 text-slate-400 hover:text-slate-200 hover:bg-slate-700',
              ].join(' ')}
            >
              Weekly Report
            </button>
            <button
              onClick={() => handleModeChange('invoice')}
              className={[
                'flex-1 sm:flex-none px-6 py-3 text-sm font-semibold transition-colors border-l border-slate-700',
                mode === 'invoice'
                  ? 'bg-blue-600 text-white'
                  : 'bg-slate-800 text-slate-400 hover:text-slate-200 hover:bg-slate-700',
              ].join(' ')}
            >
              Client Invoice
            </button>
          </div>
          <p className="text-xs text-slate-500 mt-2">
            {mode === 'weekly'
              ? 'Single-week Finance Report — FSM I and FSM II tabs only.'
              : 'Bi-weekly Client Invoice — FSM I, FSM II, and Management Detail tabs.'}
          </p>
        </section>

        {/* Upload Section */}
        <section>
          <h2 className="text-sm font-semibold uppercase tracking-widest text-slate-400 mb-4">
            Upload Source Files
          </h2>

          {mode === 'weekly' ? (
            /* ── Mode 1: Weekly — 2 upload zones ── */
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              <div>
                <div className="text-xs font-medium text-blue-400 uppercase tracking-wider mb-2">Activity Report</div>
                <DropZone
                  label="Activity Report"
                  subLabel="SAMFSM Activity Report WE [date]"
                  file={wActW1}
                  onFile={setWActW1}
                />
              </div>
              <div>
                <div className="text-xs font-medium text-indigo-400 uppercase tracking-wider mb-2">Time Off Report</div>
                <DropZone
                  label="Time Off Report"
                  subLabel="Time Off Taken (Samsung) SCHED"
                  file={wToW1}
                  onFile={setWToW1}
                />
              </div>
            </div>
          ) : (
            /* ── Mode 2: Invoice — 5 upload zones ── */
            <>
              {/* Activity Reports */}
              <div className="mb-4">
                <div className="text-xs font-medium text-blue-400 uppercase tracking-wider mb-2">Activity Reports</div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <DropZone
                    label="Activity Report — Week 1"
                    subLabel="SAMFSM Activity Report WE [date]"
                    file={iActW1}
                    onFile={setIActW1}
                  />
                  <DropZone
                    label="Activity Report — Week 2"
                    subLabel="SAMFSM Activity Report WE [date]"
                    file={iActW2}
                    onFile={setIActW2}
                  />
                </div>
              </div>

              {/* Time Off Reports */}
              <div className="mb-4">
                <div className="text-xs font-medium text-indigo-400 uppercase tracking-wider mb-2">Time Off Reports</div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <DropZone
                    label="Time Off Report — Week 1"
                    subLabel="Time Off Taken (Samsung) SCHED"
                    file={iToW1}
                    onFile={setIToW1}
                  />
                  <DropZone
                    label="Time Off Report — Week 2"
                    subLabel="Time Off Taken (Samsung) SCHED"
                    file={iToW2}
                    onFile={setIToW2}
                  />
                </div>
              </div>

              {/* Termed PTO */}
              <div>
                <div className="text-xs font-medium text-purple-400 uppercase tracking-wider mb-2">Termed PTO</div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <DropZone
                    label="Termed PTO"
                    subLabel="Termed PTO.xlsx"
                    file={iPto}
                    onFile={setIPto}
                    optional
                  />
                </div>
              </div>
            </>
          )}
        </section>

        {/* Period Preview */}
        <PeriodPreview
          actW1={mode === 'weekly' ? wActW1 : iActW1}
          actW2={mode === 'weekly' ? null : iActW2}
          mode={mode}
        />

        {/* Generate Button */}
        <section>
          <button
            onClick={handleGenerate}
            disabled={!canGenerate || step === 'processing'}
            className={[
              'w-full rounded-xl py-4 font-bold text-sm tracking-wide transition-all duration-150',
              canGenerate && step !== 'processing'
                ? 'bg-blue-600 hover:bg-blue-500 active:bg-blue-700 text-white cursor-pointer'
                : 'bg-slate-700 text-slate-500 cursor-not-allowed',
            ].join(' ')}
          >
            {step === 'processing'
              ? 'Processing…'
              : mode === 'weekly'
                ? 'Generate Weekly Report'
                : 'Generate Client Invoice'}
          </button>
          {!canGenerate && (
            <p className="text-xs text-slate-500 text-center mt-2">
              {mode === 'weekly'
                ? 'Upload Activity Report and Time Off Report to enable.'
                : 'Upload all 4 required files (Activity W1, Activity W2, Time Off W1, Time Off W2) to enable.'}
            </p>
          )}
        </section>

        {/* Status Log */}
        <StatusLog entries={logs} />

        {/* Summary */}
        {summary && step === 'done' && (
          <div ref={summaryRef}>
            <SummaryPanel
              summary={summary}
              filename={filename}
              onDownload={handleDownload}
            />
          </div>
        )}
      </main>
    </div>
  );
}
