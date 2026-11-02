// VERSION: v1.0.0 — 2026-11-02 — Initial implementation: procedural box layout for all 8 building groups
/**
 * Pure layout: converts BuildingConfig into axis-aligned box specs (center + size, world meters)
 * grouped by the eight BuildingRoot groups. No three.js here, so the geometry can be unit-tested
 * and audited (ids/tags on every box) independently of rendering.
 *
 * Facades are generated once in a face-local frame and mapped to world space:
 *   u = horizontal coordinate along the face, left -> right as seen from OUTSIDE
 *   v = height (world Y)
 *   n = distance from the building center along the face's outward normal
 */
import type { BuildingConfig } from './config';
import { computeEntranceZone, deriveDimensions, EPS, r6, type DerivedDimensions, type EntranceZone, type GlazingRow } from './derive';

export type Vec3 = [number, number, number];
export type FaceName = 'front' | 'right' | 'rear' | 'left';

export interface BoxTags {
  role: string;
  face?: FaceName | 'corner';
  floorIndex?: number;
  floor?: string;
  bay?: string;
  kind?: string;
}

export interface BoxSpec {
  id: string;
  center: Vec3;
  size: Vec3;
  tags: BoxTags;
}

export interface BuildingLayout {
  derived: DerivedDimensions;
  entranceZone: EntranceZone;
  core: BoxSpec[];
  slabs: BoxSpec[];
  facadeFrames: { verticalFins: BoxSpec[]; floorBands: BoxSpec[]; mullions: BoxSpec[]; cornerPiers: BoxSpec[] };
  glassPanels: BoxSpec[];
  entrance: { frame: BoxSpec[]; glass: BoxSpec[] };
  canopy: BoxSpec[];
  parapet: BoxSpec[];
  roofEquipment: BoxSpec[];
}

interface FaceFrame {
  name: FaceName;
  normal: Vec3;
  uAxis: Vec3;
  /** Distance from center to the outer face along the normal. */
  halfExtent: number;
  /** Face length along u. */
  length: number;
  bays: number;
}

export const FACE_NORMALS: Record<FaceName, Vec3> = {
  front: [0, 0, 1],
  right: [1, 0, 0],
  rear: [0, 0, -1],
  left: [-1, 0, 0],
};

function faceFrames(cfg: BuildingConfig, d: DerivedDimensions): FaceFrame[] {
  const W = cfg.envelope.width;
  const D = cfg.envelope.depth;
  return [
    { name: 'front', normal: FACE_NORMALS.front, uAxis: [1, 0, 0], halfExtent: D / 2, length: W, bays: d.baysX },
    { name: 'right', normal: FACE_NORMALS.right, uAxis: [0, 0, -1], halfExtent: W / 2, length: D, bays: d.baysZ },
    { name: 'rear', normal: FACE_NORMALS.rear, uAxis: [-1, 0, 0], halfExtent: D / 2, length: W, bays: d.baysX },
    { name: 'left', normal: FACE_NORMALS.left, uAxis: [0, 0, 1], halfExtent: W / 2, length: D, bays: d.baysZ },
  ];
}

/** Box from face-local extents [u0,u1] x [v0,v1] x [n0,n1] mapped to world space. */
function faceBox(f: FaceFrame, u0: number, u1: number, v0: number, v1: number, n0: number, n1: number, id: string, tags: BoxTags): BoxSpec {
  const uc = (u0 + u1) / 2;
  const nc = (n0 + n1) / 2;
  const su = u1 - u0;
  const sn = n1 - n0;
  return {
    id,
    center: [r6(f.uAxis[0] * uc + f.normal[0] * nc), r6((v0 + v1) / 2), r6(f.uAxis[2] * uc + f.normal[2] * nc)],
    size: [
      r6(Math.abs(f.uAxis[0]) * su + Math.abs(f.normal[0]) * sn),
      r6(v1 - v0),
      r6(Math.abs(f.uAxis[2]) * su + Math.abs(f.normal[2]) * sn),
    ],
    tags: { ...tags, face: tags.face ?? f.name },
  };
}

function worldBox(x0: number, x1: number, y0: number, y1: number, z0: number, z1: number, id: string, tags: BoxTags): BoxSpec {
  return {
    id,
    center: [r6((x0 + x1) / 2), r6((y0 + y1) / 2), r6((z0 + z1) / 2)],
    size: [r6(x1 - x0), r6(y1 - y0), r6(z1 - z0)],
    tags,
  };
}

const bayLabel = (k: number): string => `B${String(k + 1).padStart(2, '0')}`;

export function computeLayout(cfg: BuildingConfig): BuildingLayout {
  const d = deriveDimensions(cfg);
  const zone = computeEntranceZone(cfg, d);
  const W = cfg.envelope.width;
  const D = cfg.envelope.depth;
  const s = cfg.bay.spacing;
  const proj = cfg.facade.projectionDepth;
  const win = cfg.facade.windowModule;
  const t = cfg.facade.glassThickness;
  const mul = cfg.facade.mullion;

  const verticalFins: BoxSpec[] = [];
  const floorBands: BoxSpec[] = [];
  const mullions: BoxSpec[] = [];
  const cornerPiers: BoxSpec[] = [];
  const glassPanels: BoxSpec[] = [];

  // ---- Facades (FacadeFrames + GlassPanels) ------------------------------------------------
  for (const f of faceFrames(cfg, d)) {
    const nOuter = f.halfExtent;
    const nBody = r6(nOuter - proj); // frames: [nBody, nOuter] = 0.35 m deep
    const nGlass = r6(nOuter - d.glassSetback); // glass back face
    const uEdge = r6(f.length / 2 - d.cornerWidth); // bands run between the corner piers

    for (const band of d.bands) {
      floorBands.push(
        faceBox(f, -uEdge, uEdge, band.bottom, band.top, nBody, nOuter, `band.${f.name}.${band.label}`, {
          role: 'floorBand',
          floorIndex: band.floorIndex,
          floor: band.label,
        }),
      );
    }

    const addOpening = (u0: number, u1: number, row: GlazingRow, bay: string, kind: 'window' | 'storefront' | 'sidelight'): void => {
      const tags: BoxTags = { role: 'glass', kind, floorIndex: row.floorIndex, floor: row.floorLabel, bay };
      glassPanels.push(faceBox(f, u0, u1, row.bottom, row.top, nGlass, nGlass + t, `glass.${f.name}.${row.floorLabel}.${bay}`, tags));

      // Mullions sit just outside the glass, well inside the 0.35 m reveal.
      const n0 = nGlass + t;
      const n1 = n0 + mul.depth;
      const mTags: BoxTags = { role: 'mullion', floorIndex: row.floorIndex, floor: row.floorLabel, bay };
      const verticals: number[] = [];
      if (kind !== 'sidelight') {
        for (let i = 1; i < mul.divisionsX; i++) verticals.push(u0 + (i * (u1 - u0)) / mul.divisionsX);
      }
      verticals.forEach((u, i) =>
        mullions.push(faceBox(f, u - mul.width / 2, u + mul.width / 2, row.bottom, row.top, n0, n1, `mullion.${f.name}.${row.floorLabel}.${bay}.v${i + 1}`, mTags)),
      );
      const horizontals: number[] = [];
      if (kind === 'window') {
        for (let j = 1; j < mul.divisionsY; j++) horizontals.push(row.bottom + (j * (row.top - row.bottom)) / mul.divisionsY);
      } else {
        const transom = row.bottom + cfg.facade.groundFloorGlazing.transomHeight;
        if (transom > row.bottom + EPS && transom < row.top - EPS) horizontals.push(transom);
      }
      // Horizontal mullions are split around the verticals so no two mullions overlap.
      const breaks = [u0, ...verticals.flatMap((u) => [u - mul.width / 2, u + mul.width / 2]), u1];
      horizontals.forEach((v, j) => {
        for (let b = 0; b + 1 < breaks.length; b += 2) {
          const a = breaks[b] as number;
          const c = breaks[b + 1] as number;
          mullions.push(faceBox(f, a, c, v - mul.width / 2, v + mul.width / 2, n0, n1, `mullion.${f.name}.${row.floorLabel}.${bay}.h${j + 1}.${b / 2 + 1}`, mTags));
        }
      });
    };

    for (const row of d.glazing) {
      const entranceRow = f.name === 'front' && row.kind === 'storefront';

      // Interior vertical fins on the bay lines, one segment per glazing row (bands are continuous).
      for (let j = 1; j < f.bays; j++) {
        const u = -f.length / 2 + j * s;
        if (entranceRow && u > zone.uStart + EPS && u < zone.uEnd - EPS) continue; // fin removed for the entrance
        verticalFins.push(
          faceBox(f, u - d.finWidth / 2, u + d.finWidth / 2, row.bottom, row.top, nBody, nOuter, `fin.${f.name}.${row.floorLabel}.L${j}`, {
            role: 'verticalFin',
            floorIndex: row.floorIndex,
            floor: row.floorLabel,
          }),
        );
      }

      for (let k = 0; k < f.bays; k++) {
        if (entranceRow && zone.bayIndices.includes(k)) continue;
        const uc = -f.length / 2 + (k + 0.5) * s;
        addOpening(uc - win.width / 2, uc + win.width / 2, row, bayLabel(k), row.kind);
      }

      if (entranceRow) {
        const half = cfg.entrance.width / 2;
        if (zone.clearStart < -half - EPS) addOpening(zone.clearStart, -half, row, 'SL-L', 'sidelight');
        if (half < zone.clearEnd - EPS) addOpening(half, zone.clearEnd, row, 'SL-R', 'sidelight');
      }
    }
  }

  // Concrete corner piers (cornerWidth on each face), full height up to the roof level.
  const c = d.cornerWidth;
  for (const [sx, sz] of [[1, 1], [1, -1], [-1, -1], [-1, 1]] as const) {
    const x0 = sx > 0 ? W / 2 - c : -W / 2;
    const z0 = sz > 0 ? D / 2 - c : -D / 2;
    cornerPiers.push(worldBox(x0, x0 + c, 0, d.roofLevel, z0, z0 + c, `pier.${sx > 0 ? 'E' : 'W'}${sz > 0 ? 'S' : 'N'}`, { role: 'cornerPier', face: 'corner' }));
  }

  // ---- Slabs: one per floor (top = floor level) + roof slab; edges on the body line ---------
  const slabs: BoxSpec[] = d.floorLevels.map((L, i) =>
    worldBox(-d.bodyHalfX, d.bodyHalfX, L - cfg.slabs.floorThickness, L, -d.bodyHalfZ, d.bodyHalfZ, `slab.${d.floorLabels[i]}`, {
      role: 'slab',
      floorIndex: i,
      floor: d.floorLabels[i] ?? '',
    }),
  );
  slabs.push(
    worldBox(-d.bodyHalfX, d.bodyHalfX, d.roofLevel - cfg.slabs.roofThickness, d.roofLevel, -d.bodyHalfZ, d.bodyHalfZ, 'slab.RF', {
      role: 'slab',
      floorIndex: cfg.floors.count,
      floor: 'RF',
    }),
  );

  // ---- Core: concrete box from ground to the underside of the roof slab ---------------------
  const k = cfg.core;
  const core = [
    worldBox(k.centerX - k.width / 2, k.centerX + k.width / 2, 0, d.roofLevel - cfg.slabs.roofThickness, k.centerZ - k.depth / 2, k.centerZ + k.depth / 2, 'core', {
      role: 'core',
    }),
  ];

  // ---- Entrance (front face, centered on X = 0) --------------------------------------------
  const front = faceFrames(cfg, d)[0] as FaceFrame;
  const e = cfg.entrance;
  const nGlassFront = r6(D / 2 - d.glassSetback);
  const nF0 = r6(D / 2 - proj); // entrance frame starts at the body line like every other frame
  const nF1 = nF0 + e.frameDepth;
  const eHalf = e.width / 2;
  const innerHalf = eHalf - e.frameWidth;
  const frameTags = (part: string): BoxTags => ({ role: 'entranceFrame', kind: part, floorIndex: 0, floor: 'GF' });
  const glassTags = (part: string): BoxTags => ({ role: 'entranceGlass', kind: part, floorIndex: 0, floor: 'GF' });
  const entranceFrame: BoxSpec[] = [
    faceBox(front, -eHalf, -innerHalf, 0, e.height, nF0, nF1, 'entrance.jamb.L', frameTags('jamb')),
    faceBox(front, innerHalf, eHalf, 0, e.height, nF0, nF1, 'entrance.jamb.R', frameTags('jamb')),
    faceBox(front, -innerHalf, innerHalf, e.height - e.frameWidth, e.height, nF0, nF1, 'entrance.head', frameTags('head')),
    faceBox(front, -innerHalf, innerHalf, e.doorHeight, e.doorHeight + e.transomBarHeight, nF0, nF1, 'entrance.transomBar', frameTags('transomBar')),
  ];
  const entranceGlass: BoxSpec[] = [];
  const innerWidth = 2 * innerHalf;
  const leafWidth = (innerWidth - (e.leafCount - 1) * e.leafDividerWidth) / e.leafCount;
  let u = -innerHalf;
  for (let i = 0; i < e.leafCount; i++) {
    entranceGlass.push(faceBox(front, u, u + leafWidth, 0, e.doorHeight, nGlassFront, nGlassFront + t, `entrance.leaf.${i + 1}`, glassTags('doorLeaf')));
    u += leafWidth;
    if (i < e.leafCount - 1) {
      entranceFrame.push(faceBox(front, u, u + e.leafDividerWidth, 0, e.doorHeight, nF0, nF1, `entrance.stile.${i + 1}`, frameTags('stile')));
      u += e.leafDividerWidth;
    }
  }
  entranceGlass.push(
    faceBox(front, -innerHalf, innerHalf, e.doorHeight + e.transomBarHeight, e.height - e.frameWidth, nGlassFront, nGlassFront + t, 'entrance.transomGlass', glassTags('transom')),
  );

  // ---- Canopy: projects from the front outer face, underside on the entrance head ----------
  const cn = cfg.canopy;
  const canopy = [faceBox(front, -cn.width / 2, cn.width / 2, e.height, e.height + cn.thickness, D / 2, D / 2 + cn.projection, 'canopy', { role: 'canopy' })];

  // ---- Parapet: perimeter walls from the roof level to the parapet top ---------------------
  const p = cfg.parapet.thickness;
  const y0 = d.roofLevel;
  const y1 = d.parapetTop;
  const parapet = [
    worldBox(-W / 2, W / 2, y0, y1, D / 2 - p, D / 2, 'parapet.front', { role: 'parapet', face: 'front' }),
    worldBox(W / 2 - p, W / 2, y0, y1, -D / 2 + p, D / 2 - p, 'parapet.right', { role: 'parapet', face: 'right' }),
    worldBox(-W / 2, W / 2, y0, y1, -D / 2, -D / 2 + p, 'parapet.rear', { role: 'parapet', face: 'rear' }),
    worldBox(-W / 2, -W / 2 + p, y0, y1, -D / 2 + p, D / 2 - p, 'parapet.left', { role: 'parapet', face: 'left' }),
  ];

  // ---- Roof equipment: simple boxes standing on the roof slab --------------------------------
  const roofEquipment = cfg.roofEquipment.map((eq) =>
    worldBox(eq.centerX - eq.sizeX / 2, eq.centerX + eq.sizeX / 2, d.roofLevel, d.roofLevel + eq.height, eq.centerZ - eq.sizeZ / 2, eq.centerZ + eq.sizeZ / 2, `roofEquipment.${eq.id}`, {
      role: 'roofEquipment',
      kind: eq.label,
    }),
  );

  return {
    derived: d,
    entranceZone: zone,
    core,
    slabs,
    facadeFrames: { verticalFins, floorBands, mullions, cornerPiers },
    glassPanels,
    entrance: { frame: entranceFrame, glass: entranceGlass },
    canopy,
    parapet,
    roofEquipment,
  };
}

/** Every box in the layout (all groups), e.g. for whole-building bounds. */
export function allBoxes(l: BuildingLayout): BoxSpec[] {
  const ff = l.facadeFrames;
  return [
    ...l.core,
    ...l.slabs,
    ...ff.verticalFins,
    ...ff.floorBands,
    ...ff.mullions,
    ...ff.cornerPiers,
    ...l.glassPanels,
    ...l.entrance.frame,
    ...l.entrance.glass,
    ...l.canopy,
    ...l.parapet,
    ...l.roofEquipment,
  ];
}
