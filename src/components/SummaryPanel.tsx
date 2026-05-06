import React from 'react';
import type { AnyReport } from '../lib/types';

interface Props {
  summary: AnyReport;
  filename: string;
  onDownload: () => void;
}

const fmt = (n: number) =>
  n.toLocaleString('en-US', { style: 'currency', currency: 'USD' });

interface StatCardProps {
  label: string;
  rows: number;
  total: number;
  color: string;
}

const StatCard: React.FC<StatCardProps> = ({ label, rows, total, color }) => (
  <div className={`rounded-xl border ${color} p-4 flex flex-col gap-1`}>
    <div className="text-xs font-semibold uppercase tracking-wider text-slate-400">{label}</div>
    <div className="text-2xl font-bold text-white">{fmt(total)}</div>
    <div className="text-xs text-slate-400">{rows.toLocaleString()} row{rows !== 1 ? 's' : ''}</div>
  </div>
);

export const SummaryPanel: React.FC<Props> = ({ summary, filename, onDownload }) => {
  const isWeekly = summary.mode === 'weekly';

  return (
    <div className="mt-6 rounded-2xl border border-slate-600 bg-slate-800/50 p-6 space-y-5">
      <div>
        <h2 className="text-lg font-bold text-white">
          {isWeekly ? 'Weekly Report Generated' : 'Invoice Generated'}
        </h2>
        <p className="text-sm text-slate-400 mt-0.5">
          Billing Period:{' '}
          <span className="text-blue-300">{summary.billingPeriod}</span>
        </p>
      </div>

      {isWeekly ? (
        /* ── Mode 1: Weekly — 4 cards (FSM I, FSM II, Management, Field Labor Total) ── */
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
          <StatCard
            label="FSM I"
            rows={summary.fsmI.length}
            total={summary.fsmITotal}
            color="border-blue-700/50 bg-blue-950/20"
          />
          <StatCard
            label="FSM II"
            rows={summary.fsmII.length}
            total={summary.fsmIITotal}
            color="border-indigo-700/50 bg-indigo-950/20"
          />
          <StatCard
            label="Management"
            rows={summary.mgmt.length}
            total={summary.mgmtTotal}
            color="border-purple-700/50 bg-purple-950/20"
          />
          <div className="rounded-xl border border-emerald-600/50 bg-emerald-950/20 p-4 flex flex-col gap-1">
            <div className="text-xs font-semibold uppercase tracking-wider text-slate-400">Field Labor Total</div>
            <div className="text-2xl font-bold text-emerald-300">{fmt(summary.fieldLaborTotal)}</div>
            <div className="text-xs text-slate-400">
              {(summary.fsmI.length + summary.fsmII.length).toLocaleString()} total rows
            </div>
          </div>
        </div>
      ) : (
        /* ── Mode 2: Invoice — 4 cards (FSM I, FSM II, Mgmt, Grand Total) ── */
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
          <StatCard
            label="FSM I"
            rows={summary.fsmI.length}
            total={summary.fsmITotal}
            color="border-blue-700/50 bg-blue-950/20"
          />
          <StatCard
            label="FSM II"
            rows={summary.fsmII.length}
            total={summary.fsmIITotal}
            color="border-indigo-700/50 bg-indigo-950/20"
          />
          <StatCard
            label="Management"
            rows={summary.mgmt.length}
            total={summary.mgmtTotal}
            color="border-purple-700/50 bg-purple-950/20"
          />
          <div className="rounded-xl border border-emerald-600/50 bg-emerald-950/20 p-4 flex flex-col gap-1">
            <div className="text-xs font-semibold uppercase tracking-wider text-slate-400">Grand Total</div>
            <div className="text-2xl font-bold text-emerald-300">{fmt(summary.grandTotal)}</div>
            <div className="text-xs text-slate-400">
              {(summary.fsmI.length + summary.fsmII.length + summary.mgmt.length).toLocaleString()} total rows
            </div>
          </div>
        </div>
      )}

      <button
        onClick={onDownload}
        className="w-full flex items-center justify-center gap-2 rounded-xl bg-emerald-600 hover:bg-emerald-500 active:bg-emerald-700 text-white font-semibold py-3 text-sm transition-colors"
      >
        <span>&#x2B07;</span>
        Download {filename}
      </button>
    </div>
  );
};
