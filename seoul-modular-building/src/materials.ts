// VERSION: v1.0.0 — 2026-11-02 — Initial implementation: concrete / dark metal / blue-gray glass PBR materials
import { MeshStandardMaterial } from 'three';
import { MATERIAL_CONFIG } from './config';

export interface BuildingMaterials {
  concrete: MeshStandardMaterial;
  metal: MeshStandardMaterial;
  glass: MeshStandardMaterial;
}

export function createMaterials(mc = MATERIAL_CONFIG): BuildingMaterials {
  const concrete = new MeshStandardMaterial({
    name: 'Concrete',
    color: mc.concrete.color,
    roughness: mc.concrete.roughness,
    metalness: mc.concrete.metalness,
  });
  const metal = new MeshStandardMaterial({
    name: 'DarkMetal',
    color: mc.metal.color,
    roughness: mc.metal.roughness,
    metalness: mc.metal.metalness,
  });
  // Glass lives in its own meshes (never merged with opaque parts). depthWrite=false lets the
  // near and far facades blend instead of the first-drawn pane hiding the other one.
  const glass = new MeshStandardMaterial({
    name: 'BlueGrayGlass',
    color: mc.glass.color,
    roughness: mc.glass.roughness,
    metalness: mc.glass.metalness,
    transparent: true,
    opacity: mc.glass.opacity,
    depthWrite: false,
    envMapIntensity: mc.glass.envMapIntensity,
  });
  return { concrete, metal, glass };
}
