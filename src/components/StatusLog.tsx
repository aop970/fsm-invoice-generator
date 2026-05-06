import React from 'react';

export interface LogEntry {
  id: number;
  text: string;
  type: 'info' | 'success' | 'error' | 'progress';
}

interface Props {
  entries: LogEntry[];
}

const iconFor = (type: LogEntry['type']) => {
  switch (type) {
    case 'success':  return '✓';
    case 'error':    return '✗';
    case 'progress': return '⟳';
    default:         return '·';
  }
};

const colorFor = (type: LogEntry['type']) => {
  switch (type) {
    case 'success':  return 'text-emerald-400';
    case 'error':    return 'text-red-400';
    case 'progress': return 'text-blue-400 animate-spin';
    default:         return 'text-slate-400';
  }
};

export const StatusLog: React.FC<Props> = ({ entries }) => {
  if (entries.length === 0) return null;

  return (
    <div className="rounded-xl bg-slate-900 border border-slate-700 p-4 mt-4 font-mono text-xs space-y-1 max-h-48 overflow-y-auto">
      {entries.map(e => (
        <div key={e.id} className="flex gap-2 items-start">
          <span className={`inline-block w-4 text-center shrink-0 ${colorFor(e.type)}`}>
            {iconFor(e.type)}
          </span>
          <span className={e.type === 'error' ? 'text-red-300' : 'text-slate-300'}>{e.text}</span>
        </div>
      ))}
    </div>
  );
};
