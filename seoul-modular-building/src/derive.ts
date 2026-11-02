// VERSION: v1.0.0 — 2026-11-02 — Initial implementation: floor levels / bands / entrance zone derived from config
/**
 * Pure functions that turn BuildingConfig into derived dimensions (no three.js, unit-testable).
 * Every level is computed from the floor heights so the config stays the single source of truth.
 */
import type { BuildingConfig } from './config';

export const EPS = 1e-6;

/** Round to micrometers to strip float noise (e.g. 11.2 + 3.6 = 14.799999999999999) and -0. */
export function r6(v: number): number {
  const r = Math.round(v * 1e6) / 1e6;
  return Object.is(r, -0) ? 0 : r;
}

export type GlazingKind = 'storefront' | 'window';

/** A horizontal strip of glazing (one per floor) on every face. */
export interface GlazingRow {
  floorIndex: number;
  floorLabel: string;
  bottom: number;
  top: number;
  kind: GlazingKind;
}

/** A horizontal solid frame band between glazing rows (straddles a floor or the roof level). */
export interface SolidBand {
  label: string;
  /** Index of the floor level this band straddles (1..count-1), or `count` for the roof level. */
  floorIndex: number;
  bottom: number;
  top: number;
}

export interface DerivedDimensions {
  floorLabels: string[];
  floorHeights: number[];
  /** Top-of-slab level of each floor, index 0 = GF. */
  floorLevels: number[];
  roofLevel: number;
  parapetTop: number;
  baysX: number;
  baysZ: number;
  /** Width of an interior vertical fin = bay spacing - window width. */
  finWidth: number;
  /** Corner frame width on each face = half a fin, so windows stay centered in their bay. */
  cornerWidth: number;
  /** Half-extents of the body line = back face of the 0.35 m frames = slab edges / parapet inner face. */
  bodyHalfX: number;
  bodyHalfZ: number;
  /**
   * Glass plane set back from the frame's outer face. Derived as min(cornerWidth, projectionDepth):
   * with 0.30 m corners and 0.35 m deep frames, a 0.35 m setback would make the two corner panes of
   * adjacent faces cross each other's reveal; 0.30 m makes them meet exactly at the corner pier.
   */
  glassSetback: number;
  /** Half-extents of the glass plane (back face of the glass). */
  glassHalfX: number;
  glassHalfZ: number;
  glazing: GlazingRow[];
  bands: SolidBand[];
}

export function floorLabel(index: number): string {
  return index === 0 ? 'GF' : `${index + 1}F`;
}

export function deriveDimensions(cfg: BuildingConfig): DerivedDimensions {
  const { count, groundFloorHeight, typicalFloorHeight } = cfg.floors;
  const floorHeights: number[] = [];
  const floorLevels: number[] = [];
  const floorLabels: string[] = [];
  let level = 0;
  for (let i = 0; i < count; i++) {
    const h = i === 0 ? groundFloorHeight : typicalFloorHeight;
    floorLevels.push(r6(level));
    floorHeights.push(h);
    floorLabels.push(floorLabel(i));
    level += h;
  }
  const roofLevel = r6(level);
  const parapetTop = r6(roofLevel + cfg.roofZoneHeight);

  const spacing = cfg.bay.spacing;
  const finWidth = r6(spacing - cfg.facade.windowModule.width);
  const cornerWidth = r6(finWidth / 2);

  const glazing: GlazingRow[] = floorLevels.map((L, i) =>
    i === 0
      ? { floorIndex: 0, floorLabel: floorLabels[0] ?? 'GF', bottom: r6(L), top: r6(L + cfg.facade.groundFloorGlazing.height), kind: 'storefront' as const }
      : {
          floorIndex: i,
          floorLabel: floorLabels[i] ?? floorLabel(i),
          bottom: r6(L + cfg.facade.windowModule.sillHeight),
          top: r6(L + cfg.facade.windowModule.sillHeight + cfg.facade.windowModule.height),
          kind: 'window' as const,
        },
  );

  // Solid bands = complement of glazing rows inside [0, roofLevel].
  const bands: SolidBand[] = [];
  let cursor = 0;
  glazing.forEach((row, i) => {
    if (row.bottom > cursor + EPS) {
      bands.push({ label: `${floorLabel(i - 1)}/${floorLabel(i)}`, floorIndex: i, bottom: r6(cursor), top: row.bottom });
    }
    cursor = Math.max(cursor, row.top);
  });
  if (roofLevel > cursor + EPS) {
    bands.push({ label: `${floorLabel(count - 1)}/RF`, floorIndex: count, bottom: r6(cursor), top: roofLevel });
  }

  const proj = cfg.facade.projectionDepth;
  const glassSetback = r6(Math.min(cornerWidth, proj));
  return {
    floorLabels,
    floorHeights,
    floorLevels,
    roofLevel,
    parapetTop,
    baysX: Math.round(cfg.envelope.width / spacing),
    baysZ: Math.round(cfg.envelope.depth / spacing),
    finWidth,
    cornerWidth,
    bodyHalfX: r6(cfg.envelope.width / 2 - proj),
    bodyHalfZ: r6(cfg.envelope.depth / 2 - proj),
    glassSetback,
    glassHalfX: r6(cfg.envelope.width / 2 - glassSetback),
    glassHalfZ: r6(cfg.envelope.depth / 2 - glassSetback),
    glazing,
    bands,
  };
}

/** Bays on the front face that the central entrance cuts into, and the clear opening they leave. */
export interface EntranceZone {
  bayIndices: number[];
  /** Bay-line extents of the zone (u = X on the front face). */
  uStart: number;
  uEnd: number;
  /** Clear glazing extents inside the zone (between the bordering fins). */
  clearStart: number;
  clearEnd: number;
}

export function computeEntranceZone(cfg: BuildingConfig, d: DerivedDimensions): EntranceZone {
  const L = cfg.envelope.width;
  const s = cfg.bay.spacing;
  const half = cfg.entrance.width / 2;
  const bayIndices: number[] = [];
  for (let k = 0; k < d.baysX; k++) {
    const b0 = -L / 2 + k * s;
    const b1 = b0 + s;
    if (b1 > -half + EPS && b0 < half - EPS) bayIndices.push(k);
  }
  const first = bayIndices[0] ?? 0;
  const last = bayIndices[bayIndices.length - 1] ?? 0;
  const uStart = r6(-L / 2 + first * s);
  const uEnd = r6(-L / 2 + (last + 1) * s);
  // Corner width == half fin width, so the clear inset is identical at a face edge or interior bay line.
  return { bayIndices, uStart, uEnd, clearStart: r6(uStart + d.cornerWidth), clearEnd: r6(uEnd - d.cornerWidth) };
}
