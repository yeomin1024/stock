// VERSION: v1.0.0 — 2026-11-02 — Initial implementation: single source of truth for every dimension
/**
 * SEOUL MODULAR BUILDING — configuration.
 *
 * Coordinate system
 *   - 1 unit = 1 meter, Y-up.
 *   - Origin (0, 0, 0) = ground-level center of the footprint.
 *   - +Z = FRONT elevation (central entrance), +X = RIGHT elevation (blueprint axis gizmo).
 *
 * The values in BUILDING_CONFIG are the FINAL reference values. The blueprint image is a visual
 * reference only; if the image looks different, these numbers win. Floor levels are NOT typed in
 * by hand — they are derived from the floor heights in `derive.ts`, so changing one height moves
 * every level above it consistently.
 *
 * Values marked [assumed] are not given in the brief; they were estimated from the blueprint's
 * proportions and are isolated here so they can be adjusted without touching geometry code.
 */

export const APP_VERSION = '1.0.0';

export interface RoofEquipmentSpec {
  id: string;
  label: string;
  /** Plan center (X, Z) of the unit, meters. */
  centerX: number;
  centerZ: number;
  /** Plan size along X / Z, meters. */
  sizeX: number;
  sizeZ: number;
  /** Height above the roof slab top, meters. */
  height: number;
}

export interface BuildingConfig {
  /** Overall envelope = outer face of facade frames / parapet top. */
  envelope: { width: number; depth: number; height: number };
  floors: { count: number; groundFloorHeight: number; typicalFloorHeight: number };
  /** Roof slab top -> parapet top. */
  roofZoneHeight: number;
  /** Structural bay spacing, used along both X and Z. */
  bay: { spacing: number };
  facade: {
    /**
     * Facade projection: depth of every frame/fin/band from the body line (slab edge) to the outer
     * face. The glass plane is set back by the corner width (0.30 m, see derive.ts) inside that depth.
     */
    projectionDepth: number;
    /** Repeated window module on floors 2F-6F (glass panel size). */
    windowModule: { width: number; height: number; sillHeight: number };
    /** Ground-floor storefront glazing: same bay width, head aligned with the entrance head. */
    groundFloorGlazing: { height: number; transomHeight: number };
    /** [assumed] Thin mullions in front of each glass panel (blueprint shows divided lites). */
    mullion: { width: number; depth: number; divisionsX: number; divisionsY: number };
    glassThickness: number;
  };
  /** [assumed] thicknesses — slabs hang below their floor level (top of slab = floor level). */
  slabs: { floorThickness: number; roofThickness: number };
  /** [assumed] Concrete core, aligned to the 3.0 m grid, centered behind the entrance. */
  core: { width: number; depth: number; centerX: number; centerZ: number };
  /** Central entrance opening on the front face (outer extents of the entrance frame). */
  entrance: {
    width: number;
    height: number;
    frameWidth: number;
    frameDepth: number;
    doorHeight: number;
    transomBarHeight: number;
    leafCount: number;
    leafDividerWidth: number;
  };
  /** [assumed] Canopy slab; underside sits on the entrance head (entrance.height). */
  canopy: { width: number; projection: number; thickness: number };
  parapet: { thickness: number };
  /** [assumed] Positions/sizes traced from the blueprint TOP PLAN; kept below the parapet top. */
  roofEquipment: RoofEquipmentSpec[];
}

export const BUILDING_CONFIG: BuildingConfig = {
  envelope: { width: 18.0, depth: 12.0, height: 24.0 },
  floors: { count: 6, groundFloorHeight: 4.0, typicalFloorHeight: 3.6 },
  roofZoneHeight: 2.0,
  bay: { spacing: 3.0 },
  facade: {
    projectionDepth: 0.35,
    windowModule: { width: 2.4, height: 2.4, sillHeight: 0.6 },
    groundFloorGlazing: { height: 3.2, transomHeight: 2.4 },
    mullion: { width: 0.06, depth: 0.08, divisionsX: 2, divisionsY: 2 },
    glassThickness: 0.024,
  },
  slabs: { floorThickness: 0.25, roofThickness: 0.3 },
  core: { width: 6.0, depth: 3.0, centerX: 0.0, centerZ: 0.0 },
  entrance: {
    width: 4.0,
    height: 3.2,
    frameWidth: 0.12,
    frameDepth: 0.2,
    doorHeight: 2.4,
    transomBarHeight: 0.08,
    leafCount: 4,
    leafDividerWidth: 0.06,
  },
  canopy: { width: 4.8, projection: 1.6, thickness: 0.2 },
  parapet: { thickness: 0.35 },
  roofEquipment: [
    { id: 'RTU-01', label: 'Air handling unit', centerX: -4.9, centerZ: -2.0, sizeX: 2.0, sizeZ: 3.0, height: 1.6 },
    { id: 'RTU-02', label: 'Exhaust unit', centerX: -2.9, centerZ: -1.4, sizeX: 1.2, sizeZ: 2.8, height: 1.4 },
    { id: 'RTU-03A', label: 'Condenser A', centerX: 4.15, centerZ: -2.4, sizeX: 1.3, sizeZ: 2.4, height: 1.5 },
    { id: 'RTU-03B', label: 'Condenser B', centerX: 5.55, centerZ: -2.4, sizeX: 1.3, sizeZ: 2.4, height: 1.5 },
    { id: 'RTU-04', label: 'Cooling fan unit', centerX: 5.2, centerZ: 2.6, sizeX: 2.0, sizeZ: 2.0, height: 1.2 },
    { id: 'RTU-05', label: 'Small service unit', centerX: -6.1, centerZ: 4.0, sizeX: 1.0, sizeZ: 1.4, height: 0.9 },
  ],
};

/** sRGB hex colors; three.js converts them to linear internally (ColorManagement). */
export const MATERIAL_CONFIG = {
  concrete: { color: '#b9bbba', roughness: 0.9, metalness: 0.0 },
  metal: { color: '#4d545c', roughness: 0.5, metalness: 0.35 },
  glass: { color: '#536c7c', roughness: 0.05, metalness: 0.15, opacity: 0.8, envMapIntensity: 1.2 },
  ground: { color: '#d3d5d8', roughness: 1.0, metalness: 0.0 },
} as const;

export const RENDER_CONFIG = {
  background: '#e3e4e6',
  fog: { near: 90, far: 260 },
  toneMappingExposure: 1.0,
  maxPixelRatio: 2,
  environmentIntensity: 0.5,
  hemisphere: { sky: '#f3f6f9', ground: '#9b978f', intensity: 0.7 },
  sun: {
    color: '#fff9f0',
    intensity: 2.4,
    position: [-18, 38, 26] as [number, number, number],
    target: [0, 8, 0] as [number, number, number],
    shadowMapSize: 2048,
    shadowRadius: 4,
    shadowExtent: 26,
    shadowBias: -0.0004,
    shadowNormalBias: 0.03,
  },
  ground: { size: 400 },
  camera: { fov: 40, near: 0.1, far: 600, fitMargin: 1.08 },
  controls: { minDistance: 8, maxDistance: 160, maxPolarAngle: Math.PI / 2 - 0.02 },
} as const;

export const UI_CONFIG = {
  /** Dev GridHelper: size 60 m with 20 divisions = 3.0 m spacing, aligned with the bay grid. */
  grid: { size: 60, divisions: 20, centerColor: '#7f858c', lineColor: '#aeb3b9', visibleByDefault: true },
} as const;
