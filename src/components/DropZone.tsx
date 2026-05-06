import React, { useCallback, useState } from 'react';

interface Props {
  label: string;
  subLabel?: string;
  file: File | null;
  onFile: (f: File | null) => void;
  optional?: boolean;
  accept?: string;
}

export const DropZone: React.FC<Props> = ({
  label,
  subLabel,
  file,
  onFile,
  optional = false,
  accept = '.xlsx,.xls',
}) => {
  const [dragOver, setDragOver] = useState(false);

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setDragOver(false);
      const dropped = e.dataTransfer.files[0];
      if (dropped) onFile(dropped);
    },
    [onFile],
  );

  const handleChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const picked = e.target.files?.[0] ?? null;
      onFile(picked);
      e.target.value = '';
    },
    [onFile],
  );

  const inputId = `dz-${label.replace(/\s+/g, '-').toLowerCase()}`;

  return (
    <label
      htmlFor={inputId}
      onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
      onDragLeave={() => setDragOver(false)}
      onDrop={handleDrop}
      className={[
        'flex flex-col items-center justify-center gap-1',
        'rounded-xl border-2 cursor-pointer transition-all duration-150',
        'px-4 py-5 text-center select-none',
        dragOver
          ? 'border-blue-400 bg-blue-950/30'
          : file
            ? 'border-emerald-500 bg-emerald-950/20'
            : 'border-slate-600 bg-slate-800/40 hover:border-blue-500 hover:bg-slate-800/60',
      ].join(' ')}
    >
      <input
        id={inputId}
        type="file"
        accept={accept}
        className="sr-only"
        onChange={handleChange}
      />

      {/* Icon */}
      <div className={[
        'text-2xl mb-1',
        file ? 'text-emerald-400' : 'text-slate-400',
      ].join(' ')}>
        {file ? '✓' : '📂'}
      </div>

      {/* Labels */}
      <div className="font-semibold text-sm text-slate-200 leading-tight">{label}</div>
      {subLabel && (
        <div className="text-xs text-slate-400">{subLabel}</div>
      )}
      {optional && !file && (
        <div className="text-xs text-slate-500 italic">Optional</div>
      )}

      {/* File name */}
      {file ? (
        <div className="mt-1 text-xs text-emerald-400 truncate max-w-full" title={file.name}>
          {file.name}
        </div>
      ) : (
        <div className="mt-1 text-xs text-slate-500">Drop file or click to browse</div>
      )}

      {/* Clear button */}
      {file && (
        <button
          type="button"
          onClick={(e) => { e.preventDefault(); e.stopPropagation(); onFile(null); }}
          className="mt-1 text-xs text-red-400 hover:text-red-300 underline"
        >
          Remove
        </button>
      )}
    </label>
  );
};
