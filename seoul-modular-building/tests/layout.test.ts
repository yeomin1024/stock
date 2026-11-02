// VERSION: v1.0.0 — 2026-11-02 — Initial implementation: dimension, module and overlap checks
import { describe, expect, it } from 'vitest';
import { BUILDING_CONFIG, type BuildingConfig } from '../src/config';
import { deriveDimensions } from '../src/derive';
import { computeLayout } from '../src/layout';
import { boundsOf, validateConfig, validateLayout } from '../src/validate';

const clone = (): BuildingConfig => structuredClone(BUILDING_CONFIG);

describe('derived dimensions', () => {
  const d = deriveDimensions(BUILDING_CONFIG);

  it('derives floor levels from the floor heights', () => {
    expect(d.floorLabels).toEqual(['GF', '2F', '3F', '4F', '5F', '6F']);
    expect(d.floorLevels).toEqual([0, 4, 7.6, 11.2, 14.8, 18.4]);
    expect(d.roofLevel).toBe(22);
    expect(d.parapetTop).toBe(24);
  });

  it('tiles the footprint with 3.0 m bays and 0.6 m fins / 0.3 m corners', () => {
    expect([d.baysX, d.baysZ]).toEqual([6, 4]);
    expect(d.finWidth).toBe(0.6);
    expect(d.cornerWidth).toBe(0.3);
  });

  it('places glazing rows and the solid bands between them', () => {
    expect(d.glazing.map((g) => [g.bottom, g.top])).toEqual([
      [0, 3.2],
      [4.6, 7],
      [8.2, 10.6],
      [11.8, 14.2],
      [15.4, 17.8],
      [19, 21.4],
    ]);
    expect(d.bands.map((b) => [b.label, b.bottom, b.top])).toEqual([
      ['GF/2F', 3.2, 4.6],
      ['2F/3F', 7, 8.2],
      ['3F/4F', 10.6, 11.8],
      ['4F/5F', 14.2, 15.4],
      ['5F/6F', 17.8, 19],
      ['6F/RF', 21.4, 22],
    ]);
  });

  it('default config passes every config check', () => {
    const bad = validateConfig(BUILDING_CONFIG, d).filter((i) => i.level !== 'INFO');
    expect(bad).toEqual([]);
  });
});

describe('layout', () => {
  const layout = computeLayout(BUILDING_CONFIG);
  const report = validateLayout(layout, BUILDING_CONFIG);

  it('has no validation errors or warnings', () => {
    expect(report.issues.filter((i) => i.level !== 'INFO')).toEqual([]);
  });

  it('fits the 18.0 x 12.0 x 24.0 m envelope exactly (canopy excluded)', () => {
    expect(report.massBounds.min[0]).toBe(-9);
    expect(report.massBounds.max[0]).toBe(9);
    expect(report.massBounds.min[2]).toBe(-6);
    expect(report.massBounds.max[2]).toBe(6);
    expect(report.massBounds.max[1]).toBe(24);
    // Only the GF slab sits below grade (top of slab = +0.00).
    expect(report.massBounds.min[1]).toBe(-BUILDING_CONFIG.slabs.floorThickness);
  });

  it('builds the expected module counts', () => {
    // 2F-6F: 5 floors x (6 + 6 + 4 + 4) windows; GF: 4 front + 6 rear + 8 side storefronts + 2 sidelights.
    expect(layout.glassPanels.filter((g) => g.tags.kind === 'window')).toHaveLength(100);
    expect(layout.glassPanels.filter((g) => g.tags.kind === 'storefront')).toHaveLength(18);
    expect(layout.glassPanels.filter((g) => g.tags.kind === 'sidelight')).toHaveLength(2);
    // Fins: front 5 lines x 6 rows - 1 (removed at the entrance), rear 30, sides 3 x 6 each.
    expect(layout.facadeFrames.verticalFins).toHaveLength(29 + 30 + 18 + 18);
    expect(layout.facadeFrames.floorBands).toHaveLength(6 * 4);
    expect(layout.facadeFrames.cornerPiers).toHaveLength(4);
    expect(layout.slabs).toHaveLength(7);
    expect(layout.parapet).toHaveLength(4);
    expect(layout.roofEquipment).toHaveLength(BUILDING_CONFIG.roofEquipment.length);
  });

  it('every 2F-6F glass panel is a 2.4 x 2.4 m module centered in its 3.0 m bay', () => {
    for (const g of layout.glassPanels.filter((p) => p.tags.kind === 'window')) {
      expect(Math.max(g.size[0], g.size[2])).toBeCloseTo(2.4, 6);
      expect(g.size[1]).toBeCloseTo(2.4, 6);
      const u = g.tags.face === 'front' || g.tags.face === 'rear' ? g.center[0] : g.center[2];
      expect(Math.abs(((u % 3) + 3) % 3)).toBeCloseTo(1.5, 6); // bay centers at +-1.5, +-4.5, +-7.5
    }
  });

  it('frames project exactly 0.35 m from the glass line', () => {
    for (const f of [...layout.facadeFrames.verticalFins, ...layout.facadeFrames.floorBands]) {
      const depth = f.tags.face === 'front' || f.tags.face === 'rear' ? f.size[2] : f.size[0];
      expect(depth).toBeCloseTo(0.35, 6);
    }
  });

  it('entrance is 4.0 x 3.2 m, centered on the front face, with the canopy above it', () => {
    const eb = boundsOf(layout.entrance.frame);
    expect(eb.max[0] - eb.min[0]).toBeCloseTo(4.0, 6);
    expect(eb.max[1] - eb.min[1]).toBeCloseTo(3.2, 6);
    expect((eb.max[0] + eb.min[0]) / 2).toBeCloseTo(0, 6);
    const cb = boundsOf(layout.canopy);
    expect(cb.min[1]).toBeCloseTo(3.2, 6);
    expect(cb.min[2]).toBeCloseTo(6.0, 6);
  });

  it('glass never intersects opaque geometry', () => {
    expect(report.glassOpaqueOverlaps).toBe(0);
  });

  it('corner panes of adjacent faces meet exactly at the corner pier (no gap, no crossing)', () => {
    const d = layout.derived;
    expect(d.glassSetback).toBe(0.3);
    const frontB06 = layout.glassPanels.find((g) => g.id === 'glass.front.3F.B06');
    const rightB01 = layout.glassPanels.find((g) => g.id === 'glass.right.3F.B01');
    const pier = layout.facadeFrames.cornerPiers.find((p) => p.id === 'pier.ES');
    if (!frontB06 || !rightB01 || !pier) throw new Error('missing corner parts');
    const frontMaxX = frontB06.center[0] + frontB06.size[0] / 2;
    const frontBackZ = frontB06.center[2] - frontB06.size[2] / 2;
    const rightMaxZ = rightB01.center[2] + rightB01.size[2] / 2;
    const rightBackX = rightB01.center[0] - rightB01.size[0] / 2;
    const pierMinX = pier.center[0] - pier.size[0] / 2;
    const pierMinZ = pier.center[2] - pier.size[2] / 2;
    expect(frontMaxX).toBeCloseTo(rightBackX, 6); // front pane ends on the side glass plane
    expect(rightMaxZ).toBeCloseTo(frontBackZ, 6); // side pane ends on the front glass plane
    expect(frontMaxX).toBeCloseTo(pierMinX, 6); // ...which is the pier's inner face
    expect(rightMaxZ).toBeCloseTo(pierMinZ, 6);
  });

  it('only intentional opaque overlaps exist', () => {
    const overlap = report.issues.find((i) => i.code === 'OPAQUE_OVERLAP');
    expect(overlap?.level).toBe('INFO');
    expect(overlap?.data?.['floorBand~floorBand']).toBe(6 * 4); // 6 bands x 4 corners
    expect(overlap?.data?.['core~slab']).toBe(5); // 2F-6F slabs pass through the core
  });

  it('roof equipment stays inside the parapet and below 24.0 m', () => {
    const b = boundsOf(layout.roofEquipment);
    expect(b.min[1]).toBe(22);
    expect(b.max[1]).toBeLessThanOrEqual(24);
    expect(Math.max(-b.min[0], b.max[0])).toBeLessThanOrEqual(9 - 0.35);
    expect(Math.max(-b.min[2], b.max[2])).toBeLessThanOrEqual(6 - 0.35);
  });
});

describe('guards catch inconsistent configs', () => {
  it('flags floor heights that no longer add up to 24.0 m', () => {
    const cfg = clone();
    cfg.floors.typicalFloorHeight = 3.8;
    const issues = validateConfig(cfg, deriveDimensions(cfg));
    expect(issues.find((i) => i.code === 'HEIGHT_SUM')?.level).toBe('ERROR');
  });

  it('flags an entrance wider than the bays it replaces', () => {
    const cfg = clone();
    cfg.entrance.width = 6.0;
    const issues = validateConfig(cfg, deriveDimensions(cfg));
    expect(issues.find((i) => i.code === 'ENTRANCE_FITS')?.level).toBe('ERROR');
  });

  it('warns when roof equipment would poke above the envelope', () => {
    const cfg = clone();
    (cfg.roofEquipment[0] as { height: number }).height = 2.5;
    const issues = validateConfig(cfg, deriveDimensions(cfg));
    expect(issues.some((i) => i.code === 'EQUIPMENT_ABOVE_ENVELOPE' && i.level === 'WARN')).toBe(true);
  });
});
