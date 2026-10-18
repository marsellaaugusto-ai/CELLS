import type { ReactNode } from 'react';
import type { Peer } from '../collab/presence';

export type ModeId = 'archivo' | 'datos' | 'formato' | 'formulas' | 'graficos' | 'conectar';

export interface Mode {
  id: ModeId;
  label: string;
  /** Triangle and section-title color. */
  color: string;
  /** Header fill; dark enough for white text. */
  strong: string;
  /** Active pill / selected background. */
  tint: string;
}

export const MODES: Mode[] = [
  { id: 'archivo', label: 'Archivo', color: '#4b9ff0', strong: '#2a6fc0', tint: '#1d3047' },
  { id: 'datos', label: 'Datos', color: '#5cc26a', strong: '#2f8f3a', tint: '#1f3a27' },
  { id: 'formato', label: 'Formato', color: '#2bb5c9', strong: '#1c7c8a', tint: '#163a40' },
  { id: 'formulas', label: 'Fórmulas', color: '#e8743b', strong: '#b0501e', tint: '#40261a' },
  { id: 'graficos', label: 'Gráficos', color: '#e9b326', strong: '#8a640b', tint: '#3d3217' },
  { id: 'conectar', label: 'Conectar', color: '#a77de0', strong: '#6b44a8', tint: '#2e2440' },
];

export function Triangle({ color }: { color: string }) {
  return (
    <svg width="22" height="20" viewBox="0 0 22 20" aria-hidden="true">
      <path d="M11 2 L20 18 L2 18 Z" fill={color} />
    </svg>
  );
}

export function Rail({ mode, onMode }: { mode: ModeId; onMode(m: ModeId): void }) {
  return (
    <nav className="rail" aria-label="Modos">
      <div className="brand">
        <img src="./favicon.svg" width="28" height="28" alt="" />
        Cells
      </div>
      {MODES.map((m) => (
        <button
          key={m.id}
          type="button"
          className="rail-btn"
          aria-current={m.id === mode ? 'page' : undefined}
          onClick={() => onMode(m.id)}
        >
          <Triangle color={m.color} />
          {m.label}
        </button>
      ))}
    </nav>
  );
}

export function ModeHead({ index, label }: { index: number; label: string }) {
  return (
    <div className="mode-head">
      <span className="num">{String(index).padStart(2, '0')}</span>
      <h1>{label}</h1>
    </div>
  );
}

export function Card({ title, children }: { title: string; children: ReactNode }) {
  return (
    <section className="card">
      <h2>{title}</h2>
      {children}
    </section>
  );
}

export function Toggle({ on, onChange, children }: { on: boolean; onChange(v: boolean): void; children: ReactNode }) {
  return (
    <button type="button" className="toggle" aria-pressed={on} onClick={() => onChange(!on)}>
      <span className="track">
        <span className="knob" />
      </span>
      {children}
    </button>
  );
}

export function PlayIcon() {
  return (
    <svg width="10" height="12" viewBox="0 0 10 12" aria-hidden="true">
      <path d="M1 1 L9 6 L1 11 Z" fill="currentColor" />
    </svg>
  );
}

/** Who has this workbook open right now, for the sheet's header strip. */
export function People({ me, peers }: { me: Peer; peers: Peer[] }) {
  return (
    <ul className="people" aria-label="Personas en este libro">
      <li>
        <span className="dot" style={{ background: me.color }} />
        Vos
      </li>
      {peers.map((p) => (
        <li key={p.id}>
          <span className="dot" style={{ background: p.color }} />
          {p.name}
        </li>
      ))}
    </ul>
  );
}
