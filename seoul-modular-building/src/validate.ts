// VERSION: v1.0.0 — 2026-11-02 — Initial implementation: config sanity checks + layout envelope/overlap audit
/**
 * Pure validation (no three.js). Returns structured issues instead of throwing so the app can
 * log every finding at the VALIDATE stage and keep rendering whatever is valid.
 */
import type { BuildingConfig } from './config';
import { computeEntranceZone, EPS, r6, type DerivedDimensions } from './derive';
import type { BoxSpec, BuildingLayout } from './layout';

export type IssueLevel = 'ERROR' | 'WARN' | 'INFO';

export interface Issue {
  level: IssueLevel;
  code: string;
  message: string;
  data?: Record<string, string | number | boolean>;
}

export interface Bounds {
  min: [number, number, number];
  max: [number, number, number];
}

const near = (a: number, b: number, tol = 1e-4): boolean => Math.abs(a - b) <= tol;

export function boundsOf(specs: readonly BoxSpec[]): Bounds {
  const min: [number, number, number] = [Infinity, Infinity, Infinity];
  const max: [number, number, number] = [-Infinity, -Infinity, -Infinity];
  for (const b of specs) {
    for (let a = 0; a < 3; a++) {
      const lo = (b.center[a] as number) - (b.size[a] as number) / 2;
      const hi = (b.center[a] as number) + (b.size[a] as number) / 2;
      if (lo < (min[a] as number)) min[a] = r6(lo);
      if (hi > (max[a] as number)) max[a] = r6(hi);
    }
  }
  return { min, max };
}

/** Positive-volume intersection (touching faces do NOT count). */
export function boxesOverlap(a: BoxSpec, b: BoxSpec, tol = 1e-5): boolean {
  for (let i = 0; i < 3; i++) {
    const aLo = (a.center[i] as number) - (a.size[i] as number) / 2;
    const aHi = (a.center[i] as number) + (a.size[i] as number) / 2;
    const bLo = (b.center[i] as number) - (b.size[i] as number) / 2;
    const bHi = (b.center[i] as number) + (b.size[i] as number) / 2;
    if (Math.min(aHi, bHi) - Math.max(aLo, bLo) <= tol) return false;
  }
  return true;
}

export function validateConfig(cfg: BuildingConfig, d: DerivedDimensions): Issue[] {
  const issues: Issue[] = [];
  const push = (level: IssueLevel, code: string, message: string, data?: Issue['data']): void => {
    issues.push(data ? { level, code, message, data } : { level, code, message });
  };

  // 1. Floor heights + roof zone must add up to the overall height.
  push(near(d.parapetTop, cfg.envelope.height) ? 'INFO' : 'ERROR', 'HEIGHT_SUM', 'GF + typical floors + roof zone must equal envelope height', {
    computed: d.parapetTop,
    expected: cfg.envelope.height,
  });

  // 2. Bay grid must tile the footprint exactly.
  for (const [axis, len, bays] of [['X', cfg.envelope.width, d.baysX], ['Z', cfg.envelope.depth, d.baysZ]] as const) {
    const ok = near(bays * cfg.bay.spacing, len);
    push(ok ? 'INFO' : 'ERROR', `BAY_GRID_${axis}`, `bay spacing must tile the ${axis} extent`, { length: len, spacing: cfg.bay.spacing, bays });
  }

  // 3. Window module must fit inside a bay with a positive fin.
  if (!(d.finWidth > EPS)) push('ERROR', 'WINDOW_FITS_BAY', 'window width must be smaller than bay spacing', { window: cfg.facade.windowModule.width, bay: cfg.bay.spacing });

  // 4. Every glazing row must stay below the next one and below the roof level.
  d.glazing.forEach((row, i) => {
    const nextBottom = d.glazing[i + 1]?.bottom ?? d.roofLevel;
    if (row.top > nextBottom + EPS) push('ERROR', 'GLAZING_ROW_OVERLAP', `glazing of ${row.floorLabel} runs into the row above`, { top: row.top, nextBottom });
  });
  d.bands.forEach((b) => {
    if (!(b.top - b.bottom > EPS)) push('ERROR', 'BAND_NON_POSITIVE', `band ${b.label} has no height`, { bottom: b.bottom, top: b.top });
  });

  // 5. Entrance must fit inside the bays it replaces and under the GF band.
  const zone = computeEntranceZone(cfg, d);
  const clear = r6(zone.clearEnd - zone.clearStart);
  push(cfg.entrance.width <= clear + EPS ? 'INFO' : 'ERROR', 'ENTRANCE_FITS', 'entrance width must fit the clear opening of the bays it replaces', {
    entranceWidth: cfg.entrance.width,
    clearWidth: clear,
    bays: zone.bayIndices.map((k) => k + 1).join('+'),
  });
  if (cfg.entrance.height > cfg.facade.groundFloorGlazing.height + EPS) {
    push('ERROR', 'ENTRANCE_HEIGHT', 'entrance must not be taller than the GF glazing row (it would cut into the GF band)', {
      entranceHeight: cfg.entrance.height,
      glazingHeight: cfg.facade.groundFloorGlazing.height,
    });
  }
  const e = cfg.entrance;
  if (e.doorHeight + e.transomBarHeight > e.height - e.frameWidth - EPS) {
    push('ERROR', 'ENTRANCE_DOOR_HEIGHT', 'door + transom bar must stay under the entrance head', { doorHeight: e.doorHeight, head: e.height - e.frameWidth });
  }
  if (cfg.canopy.width < cfg.entrance.width - EPS) push('WARN', 'CANOPY_NARROW', 'canopy is narrower than the entrance', { canopy: cfg.canopy.width, entrance: cfg.entrance.width });

  // 6. Core must stay inside the body line.
  const k = cfg.core;
  if (Math.abs(k.centerX) + k.width / 2 > d.bodyHalfX + EPS || Math.abs(k.centerZ) + k.depth / 2 > d.bodyHalfZ + EPS) {
    push('ERROR', 'CORE_OUTSIDE', 'core footprint crosses the facade body line');
  }

  // 6b. Glass + mullions must fit inside the frame depth (they sit in the reveal).
  const inReveal = cfg.facade.glassThickness + cfg.facade.mullion.depth;
  push(inReveal <= d.glassSetback + EPS ? 'INFO' : 'ERROR', 'GLASS_IN_REVEAL', 'glass + mullion depth must fit in front of the glass setback', {
    glassSetback: d.glassSetback,
    projectionDepth: cfg.facade.projectionDepth,
    cornerWidth: d.cornerWidth,
    glassPlusMullion: r6(inReveal),
  });

  // 7. Roof equipment must sit inside the parapet and (by default) below the parapet top.
  const innerX = cfg.envelope.width / 2 - cfg.parapet.thickness;
  const innerZ = cfg.envelope.depth / 2 - cfg.parapet.thickness;
  for (const eq of cfg.roofEquipment) {
    if (Math.abs(eq.centerX) + eq.sizeX / 2 > innerX + EPS || Math.abs(eq.centerZ) + eq.sizeZ / 2 > innerZ + EPS) {
      push('ERROR', 'EQUIPMENT_OUTSIDE', `${eq.id} crosses the parapet`, { id: eq.id });
    }
    const top = r6(d.roofLevel + eq.height);
    if (top > d.parapetTop + EPS) {
      push('WARN', 'EQUIPMENT_ABOVE_ENVELOPE', `${eq.id} top exceeds the ${cfg.envelope.height} m envelope height`, { id: eq.id, top, envelope: d.parapetTop });
    }
  }
  return issues;
}

export interface LayoutReport {
  issues: Issue[];
  massBounds: Bounds;
  canopyBounds: Bounds;
  glassOpaqueOverlaps: number;
  counts: Record<string, number>;
}

export function validateLayout(layout: BuildingLayout, cfg: BuildingConfig): LayoutReport {
  const issues: Issue[] = [];
  const ff = layout.facadeFrames;
  const opaque: BoxSpec[] = [
    ...layout.core,
    ...layout.slabs,
    ...ff.verticalFins,
    ...ff.floorBands,
    ...ff.mullions,
    ...ff.cornerPiers,
    ...layout.entrance.frame,
    ...layout.canopy,
    ...layout.parapet,
    ...layout.roofEquipment,
  ];
  const glass = [...layout.glassPanels, ...layout.entrance.glass];

  // Envelope: everything except the canopy must fit 18.0 x 12.0 and top out at 24.0.
  const mass = [...opaque.filter((b) => b.tags.role !== 'canopy'), ...glass];
  const massBounds = boundsOf(mass);
  const W = cfg.envelope.width;
  const D = cfg.envelope.depth;
  const H = cfg.envelope.height;
  const envOk =
    near(massBounds.min[0], -W / 2) && near(massBounds.max[0], W / 2) && near(massBounds.min[2], -D / 2) && near(massBounds.max[2], D / 2) && near(massBounds.max[1], H);
  issues.push({
    level: envOk ? 'INFO' : 'ERROR',
    code: 'ENVELOPE',
    message: 'building mass (excl. canopy) must measure W x D x H exactly',
    data: {
      sizeX: r6(massBounds.max[0] - massBounds.min[0]),
      sizeZ: r6(massBounds.max[2] - massBounds.min[2]),
      topY: massBounds.max[1],
      belowGradeY: massBounds.min[1],
      expected: `${W}x${D}x${H}`,
    },
  });

  const canopyBounds = boundsOf(layout.canopy);
  issues.push({
    level: 'INFO',
    code: 'CANOPY_PROJECTION',
    message: 'canopy projects beyond the front envelope by design',
    data: { frontFaceZ: D / 2, canopyMaxZ: canopyBounds.max[2], projection: r6(canopyBounds.max[2] - D / 2) },
  });

  // Window module: every 2F-6F panel must be exactly the configured module.
  const win = cfg.facade.windowModule;
  const windows = layout.glassPanels.filter((g) => g.tags.kind === 'window');
  const offModule = windows.filter((g) => {
    const horizontal = Math.max(g.size[0], g.size[2]);
    return !near(horizontal, win.width) || !near(g.size[1], win.height);
  });
  issues.push({
    level: offModule.length === 0 ? 'INFO' : 'ERROR',
    code: 'WINDOW_MODULE',
    message: `2F-6F glass panels must be ${win.width} x ${win.height} m`,
    data: { panels: windows.length, offModule: offModule.length },
  });

  // Entrance outer extents.
  const eb = boundsOf(layout.entrance.frame);
  const eOk = near(eb.max[0] - eb.min[0], cfg.entrance.width) && near(eb.max[1] - eb.min[1], cfg.entrance.height) && near((eb.max[0] + eb.min[0]) / 2, 0);
  issues.push({
    level: eOk ? 'INFO' : 'ERROR',
    code: 'ENTRANCE_SIZE',
    message: 'entrance frame outer extents must equal the configured opening, centered on X=0',
    data: { width: r6(eb.max[0] - eb.min[0]), height: r6(eb.max[1] - eb.min[1]), centerX: r6((eb.max[0] + eb.min[0]) / 2) },
  });

  // Transparent vs opaque intersections cause sorting/z-fighting artifacts -> must be zero.
  let glassOpaqueOverlaps = 0;
  const examples: string[] = [];
  for (const g of glass) {
    for (const o of opaque) {
      if (boxesOverlap(g, o)) {
        glassOpaqueOverlaps++;
        if (examples.length < 3) examples.push(`${g.id}~${o.id}`);
      }
    }
  }
  issues.push({
    level: glassOpaqueOverlaps === 0 ? 'INFO' : 'ERROR',
    code: 'GLASS_OPAQUE_OVERLAP',
    message: 'glass panels must not intersect opaque parts',
    data: { overlaps: glassOpaqueOverlaps, examples: examples.join(',') || 'none' },
  });

  // Opaque/opaque overlaps are allowed only where intentional (same material, hidden faces):
  // floor bands of adjacent faces share a 5 x 5 cm corner joint, and floor slabs pass through the core.
  const opaqueOverlaps: Record<string, number> = {};
  for (let i = 0; i < opaque.length; i++) {
    for (let j = i + 1; j < opaque.length; j++) {
      const a = opaque[i] as BoxSpec;
      const b = opaque[j] as BoxSpec;
      if (boxesOverlap(a, b)) {
        const key = [a.tags.role, b.tags.role].sort().join('~');
        opaqueOverlaps[key] = (opaqueOverlaps[key] ?? 0) + 1;
      }
    }
  }
  const allowed = new Set(['floorBand~floorBand', 'core~slab']);
  const unexpected = Object.keys(opaqueOverlaps).filter((k) => !allowed.has(k));
  issues.push({
    level: unexpected.length === 0 ? 'INFO' : 'WARN',
    code: 'OPAQUE_OVERLAP',
    message: 'opaque overlaps (only band corner joints and slab-through-core are expected)',
    data: { ...opaqueOverlaps, unexpected: unexpected.join(',') || 'none' },
  });

  const counts: Record<string, number> = {
    core: layout.core.length,
    slabs: layout.slabs.length,
    verticalFins: ff.verticalFins.length,
    floorBands: ff.floorBands.length,
    mullions: ff.mullions.length,
    cornerPiers: ff.cornerPiers.length,
    glassPanels: layout.glassPanels.length,
    entranceFrame: layout.entrance.frame.length,
    entranceGlass: layout.entrance.glass.length,
    canopy: layout.canopy.length,
    parapet: layout.parapet.length,
    roofEquipment: layout.roofEquipment.length,
  };
  return { issues, massBounds, canopyBounds, glassOpaqueOverlaps, counts };
}
