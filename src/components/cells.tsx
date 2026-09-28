import { useEffect, useState, type KeyboardEvent } from 'react';
import { isValidISO } from '../engine/dates';

/** Commits on Enter / blur only, so half-typed dates don't reshuffle the plan. */
export function DateInput(props: {
  value: string | null;
  onCommit: (iso: string | null) => void;
  allowEmpty?: boolean;
  disabled?: boolean;
  className?: string;
  title?: string;
  'aria-label'?: string;
}) {
  const [v, setV] = useState(props.value ?? '');
  useEffect(() => setV(props.value ?? ''), [props.value]);
  const commit = () => {
    if (v === (props.value ?? '')) return;
    if (v === '' && props.allowEmpty) props.onCommit(null);
    else if (isValidISO(v)) props.onCommit(v);
    else setV(props.value ?? '');
  };
  return (
    <input
      type="date"
      className={props.className}
      value={v}
      disabled={props.disabled}
      title={props.title}
      aria-label={props['aria-label']}
      onChange={(e) => setV(e.target.value)}
      onBlur={commit}
      onKeyDown={(e: KeyboardEvent<HTMLInputElement>) => {
        if (e.key === 'Enter') (e.target as HTMLInputElement).blur();
        if (e.key === 'Escape') {
          setV(props.value ?? '');
          (e.target as HTMLInputElement).blur();
        }
      }}
    />
  );
}

export function NumberInput(props: {
  value: number | null | undefined;
  onCommit: (n: number | null) => void;
  min?: number;
  disabled?: boolean;
  className?: string;
  title?: string;
  'aria-label'?: string;
}) {
  const shown = props.value == null ? '' : String(props.value);
  const [v, setV] = useState(shown);
  useEffect(() => setV(shown), [shown]);
  const commit = () => {
    if (v === shown) return;
    if (v.trim() === '') return props.onCommit(null);
    const n = Number(v);
    if (Number.isFinite(n) && n >= (props.min ?? 0)) props.onCommit(Math.round(n));
    else setV(shown);
  };
  return (
    <input
      type="number"
      inputMode="numeric"
      className={props.className}
      value={v}
      min={props.min ?? 0}
      disabled={props.disabled}
      title={props.title}
      aria-label={props['aria-label']}
      onChange={(e) => setV(e.target.value)}
      onBlur={commit}
      onKeyDown={(e) => {
        if (e.key === 'Enter') (e.target as HTMLInputElement).blur();
        if (e.key === 'Escape') {
          setV(shown);
          (e.target as HTMLInputElement).blur();
        }
      }}
    />
  );
}

export function TextInput(props: {
  value: string;
  onCommit: (s: string) => void;
  className?: string;
  placeholder?: string;
  'aria-label'?: string;
}) {
  const [v, setV] = useState(props.value);
  useEffect(() => setV(props.value), [props.value]);
  return (
    <input
      type="text"
      className={props.className}
      value={v}
      placeholder={props.placeholder}
      aria-label={props['aria-label']}
      onChange={(e) => setV(e.target.value)}
      onBlur={() => v !== props.value && props.onCommit(v)}
      onKeyDown={(e) => {
        if (e.key === 'Enter') (e.target as HTMLInputElement).blur();
        if (e.key === 'Escape') {
          setV(props.value);
          (e.target as HTMLInputElement).blur();
        }
      }}
    />
  );
}
