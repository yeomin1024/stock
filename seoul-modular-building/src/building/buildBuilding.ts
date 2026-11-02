// VERSION: v1.0.0 — 2026-11-02 — Initial implementation: BuildingRoot with the 8 named groups
import { Group, InstancedMesh, Mesh } from 'three';
import type { BuildingLayout } from '../layout';
import type { Logger } from '../logger';
import type { BuildingMaterials } from '../materials';
import { createBoxMesh, createInstancedBoxes, GLASS_PART, TRIANGLES_PER_BOX } from './instancing';

export const GROUP_NAMES = ['Core', 'Slabs', 'FacadeFrames', 'GlassPanels', 'Entrance', 'Canopy', 'Parapet', 'RoofEquipment'] as const;
export type GroupName = (typeof GROUP_NAMES)[number];

export interface GroupStats {
  meshes: number;
  instances: number;
  triangles: number;
  drawCalls: number;
}

export interface BuiltBuilding {
  root: Group;
  groups: Record<GroupName, Group>;
  stats: Record<GroupName, GroupStats> & { total: GroupStats };
}

function statsOf(group: Group): GroupStats {
  const st: GroupStats = { meshes: 0, instances: 0, triangles: 0, drawCalls: 0 };
  group.traverse((o) => {
    if (o instanceof InstancedMesh) {
      st.meshes++;
      st.instances += o.count;
      st.triangles += o.count * TRIANGLES_PER_BOX;
      st.drawCalls++;
    } else if (o instanceof Mesh) {
      st.meshes++;
      st.instances++;
      st.triangles += TRIANGLES_PER_BOX;
      st.drawCalls++;
    }
  });
  return st;
}

export function buildBuilding(layout: BuildingLayout, mats: BuildingMaterials, log: Logger): BuiltBuilding {
  const root = new Group();
  root.name = 'BuildingRoot';
  const groups = {} as Record<GroupName, Group>;
  for (const name of GROUP_NAMES) {
    const g = new Group();
    g.name = name;
    root.add(g);
    groups[name] = g;
  }

  const [coreSpec] = layout.core;
  const [canopySpec] = layout.canopy;
  if (!coreSpec || !canopySpec) throw new Error('layout is missing the core or canopy box');

  groups.Core.add(createBoxMesh('Core.Walls', coreSpec, mats.concrete));
  groups.Slabs.add(createInstancedBoxes('Slabs.Floors', layout.slabs, mats.concrete));

  const ff = layout.facadeFrames;
  groups.FacadeFrames.add(
    createInstancedBoxes('FacadeFrames.VerticalFins', ff.verticalFins, mats.metal),
    createInstancedBoxes('FacadeFrames.FloorBands', ff.floorBands, mats.metal),
    createInstancedBoxes('FacadeFrames.Mullions', ff.mullions, mats.metal),
    createInstancedBoxes('FacadeFrames.CornerPiers', ff.cornerPiers, mats.concrete),
  );

  groups.GlassPanels.add(createInstancedBoxes('GlassPanels.Glazing', layout.glassPanels, mats.glass, GLASS_PART));

  groups.Entrance.add(
    createInstancedBoxes('Entrance.Frame', layout.entrance.frame, mats.metal),
    createInstancedBoxes('Entrance.Glass', layout.entrance.glass, mats.glass, GLASS_PART),
  );

  groups.Canopy.add(createBoxMesh('Canopy.Slab', canopySpec, mats.metal));
  groups.Parapet.add(createInstancedBoxes('Parapet.Walls', layout.parapet, mats.concrete));
  groups.RoofEquipment.add(createInstancedBoxes('RoofEquipment.Units', layout.roofEquipment, mats.metal));

  const stats = {} as BuiltBuilding['stats'];
  const total: GroupStats = { meshes: 0, instances: 0, triangles: 0, drawCalls: 0 };
  for (const name of GROUP_NAMES) {
    const st = statsOf(groups[name]);
    stats[name] = st;
    total.meshes += st.meshes;
    total.instances += st.instances;
    total.triangles += st.triangles;
    total.drawCalls += st.drawCalls;
    log.info('BUILD', 'group_built', {
      group: name,
      children: groups[name].children.map((c) => c.name),
      instances: st.instances,
      triangles: st.triangles,
      drawCalls: st.drawCalls,
    });
  }
  stats.total = total;
  log.info('BUILD', 'building_built', { root: root.name, groups: GROUP_NAMES.length, ...total });
  return { root, groups, stats };
}
