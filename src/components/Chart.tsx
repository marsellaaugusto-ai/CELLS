import { rectKeys, toKey, type Rect } from '../model/address';
import type { Value } from '../model/types';
import { formatValue } from '../engine/format';

/**
 * Bars for the numbers in the selection, in the same style as the reference design:
 * blue above zero, red below. Each bar is labelled with the nearest text to its left
 * in the same row (e.g. the product name), or with its address.
 */
export function Chart({ values, rect }: { values: Record<string, Value>; rect: Rect }) {
  const points: { label: string; v: number }[] = [];
  for (const key of rectKeys(rect)) {
    const v = values[key];
    if (typeof v !== 'number') continue;
    const r = Number(key.replace(/^[A-Z]+/, '')) - 1;
    let label = key;
    for (let c = rect.c0 - 1; c >= 0; c--) {
      const t = values[toKey(r, c)];
      if (typeof t === 'string' && t.trim()) {
        label = t;
        break;
      }
    }
    points.push({ label, v });
  }
  const shown = points.slice(0, 40);

  if (!shown.length) {
    return <p className="empty">Seleccioná celdas con números para ver el gráfico.</p>;
  }

  const maxPos = Math.max(0, ...shown.map((p) => p.v));
  const maxNeg = Math.max(0, ...shown.map((p) => -p.v));
  const scale = Math.max(maxPos, maxNeg) || 1;
  const hasNeg = maxNeg > 0;

  return (
    <div className="chart" role="img" aria-label={`Gráfico de barras de ${shown.length} valores`}>
      {shown.map((p, i) => (
        <div className="bar-col" key={i} title={`${p.label}: ${formatValue(p.v)}`}>
          <div className="bar-area">
            <div className="half pos">
              {p.v > 0 && <div className="bar pos" style={{ height: `${(p.v / scale) * 100}%` }} />}
            </div>
            {hasNeg && (
              <div className="half neg">
                {p.v < 0 && <div className="bar neg" style={{ height: `${(-p.v / scale) * 100}%` }} />}
              </div>
            )}
          </div>
          <span className="bar-label">{p.label}</span>
        </div>
      ))}
    </div>
  );
}
