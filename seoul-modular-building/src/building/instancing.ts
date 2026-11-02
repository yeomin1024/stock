// VERSION: v1.0.0 — 2026-11-02 — Initial implementation: BoxSpec[] -> InstancedMesh / Mesh
import { BoxGeometry, InstancedMesh, Matrix4, Mesh, Quaternion, Vector3, type Material } from 'three';
import type { BoxSpec } from '../layout';

/** One shared 1 x 1 x 1 box (12 triangles). Every part is this box scaled per instance. */
export const UNIT_BOX = new BoxGeometry(1, 1, 1);
export const TRIANGLES_PER_BOX = (UNIT_BOX.index?.count ?? 36) / 3;

export interface PartOptions {
  castShadow: boolean;
  receiveShadow: boolean;
  renderOrder?: number;
}

const OPAQUE: PartOptions = { castShadow: true, receiveShadow: true };
/**
 * Glass: drawn after opaque parts. It casts a (full) shadow: with non-casting glass the sun shines
 * through both facades and speckles the ground shadow with window-shaped light patches.
 */
export const GLASS_PART: PartOptions = { castShadow: true, receiveShadow: false, renderOrder: 1 };

function tagUserData(target: Mesh, specs: readonly BoxSpec[]): void {
  target.userData = {
    ids: specs.map((s) => s.id),
    tags: specs.map((s) => s.tags),
  };
}

/** Repeated parts -> a single InstancedMesh (one draw call per material). */
export function createInstancedBoxes(name: string, specs: readonly BoxSpec[], material: Material, opts: PartOptions = OPAQUE): InstancedMesh {
  const mesh = new InstancedMesh(UNIT_BOX, material, specs.length);
  const m = new Matrix4();
  const q = new Quaternion();
  const p = new Vector3();
  const s = new Vector3();
  specs.forEach((spec, i) => {
    p.fromArray(spec.center);
    s.fromArray(spec.size);
    mesh.setMatrixAt(i, m.compose(p, q, s));
  });
  mesh.instanceMatrix.needsUpdate = true;
  mesh.computeBoundingBox();
  mesh.computeBoundingSphere();
  mesh.name = name;
  mesh.castShadow = opts.castShadow;
  mesh.receiveShadow = opts.receiveShadow;
  mesh.renderOrder = opts.renderOrder ?? 0;
  tagUserData(mesh, specs);
  return mesh;
}

/** Single parts (core, canopy) -> plain Mesh using the same unit box. */
export function createBoxMesh(name: string, spec: BoxSpec, material: Material, opts: PartOptions = OPAQUE): Mesh {
  const mesh = new Mesh(UNIT_BOX, material);
  mesh.position.fromArray(spec.center);
  mesh.scale.fromArray(spec.size);
  mesh.name = name;
  mesh.castShadow = opts.castShadow;
  mesh.receiveShadow = opts.receiveShadow;
  mesh.renderOrder = opts.renderOrder ?? 0;
  tagUserData(mesh, [spec]);
  return mesh;
}
