// VERSION: v1.0.0 — 2026-11-02 — Initial implementation: deterministic assembly + turntable timeline
/**
 * Assembly demo. `seek(t)` is a pure function of time (no accumulated state), so the same
 * timeline drives real-time playback in the UI and frame-exact video capture.
 * Parts fly in per instance (staggered by floor) and are restored to their exact base matrices
 * at the end, so the demo can never leave the model in a modified state.
 */
import { InstancedMesh, Mesh, Vector3, type Box3, type Object3D } from 'three';
import { GROUP_NAMES, type BuiltBuilding, type GroupName } from './building/buildBuilding';
import type { BuildingConfig } from './config';
import { FACE_NORMALS, type BoxTags, type BuildingLayout } from './layout';
import type { Logger } from './logger';
import { directionFrom, VIEW_PRESETS, type Stage } from './stage';

type Direction = 'grow' | 'outward' | 'above' | 'front';

export interface DemoStep {
  group: GroupName;
  title: string;
  detail: string;
  start: number;
  duration: number;
  direction: Direction;
  distance: number;
  stagger: number;
  staggerBy: 'floor' | 'index' | 'none';
}

export interface AssemblyDemo {
  readonly duration: number;
  readonly steps: readonly DemoStep[];
  isPlaying(): boolean;
  play(): void;
  stop(): void;
  seek(t: number): void;
  update(nowMs: number): void;
}

interface Animator {
  apply(local: number, step: DemoStep): void;
  reset(): void;
}

const clamp01 = (v: number): number => (v < 0 ? 0 : v > 1 ? 1 : v);
const easeOutCubic = (p: number): number => 1 - (1 - p) ** 3;
const easeInOutSine = (p: number): number => -(Math.cos(Math.PI * p) - 1) / 2;
const m = (v: number): string => (Math.abs(v * 10 - Math.round(v * 10)) < 1e-9 ? v.toFixed(1) : String(v));

function directionFor(direction: Direction, tags: BoxTags | undefined, x: number, z: number, out: Vector3): Vector3 {
  if (direction === 'above') return out.set(0, 1, 0);
  if (direction === 'front') return out.set(0, 0, 1);
  if (direction === 'outward') {
    const face = tags?.face;
    if (face && face !== 'corner') return out.fromArray(FACE_NORMALS[face]);
    const len = Math.hypot(x, z);
    return len > 1e-6 ? out.set(x / len, 0, z / len) : out.set(0, 0, 1);
  }
  return out.set(0, 0, 0);
}

function instancedAnimator(mesh: InstancedMesh): Animator {
  const arr = mesh.instanceMatrix.array as Float32Array;
  const base = Float32Array.from(arr);
  const tags = (mesh.userData['tags'] ?? []) as BoxTags[];
  const dir = new Vector3();
  return {
    apply(local, step) {
      for (let i = 0; i < mesh.count; i++) {
        const o = i * 16;
        const key = step.staggerBy === 'floor' ? (tags[i]?.floorIndex ?? 0) : step.staggerBy === 'index' ? i : 0;
        const p = clamp01((local - key * step.stagger) / step.duration);
        arr.set(base.subarray(o, o + 16), o);
        if (p <= 0) {
          arr[o] = 0; // scale -> 0 hides the instance (degenerate triangles are not rasterized)
          arr[o + 5] = 0;
          arr[o + 10] = 0;
        } else if (p < 1) {
          directionFor(step.direction, tags[i], base[o + 12] as number, base[o + 14] as number, dir);
          const k = (1 - easeOutCubic(p)) * step.distance;
          arr[o + 12] = (base[o + 12] as number) + dir.x * k;
          arr[o + 13] = (base[o + 13] as number) + dir.y * k;
          arr[o + 14] = (base[o + 14] as number) + dir.z * k;
        }
      }
      mesh.instanceMatrix.needsUpdate = true;
    },
    reset() {
      arr.set(base);
      mesh.instanceMatrix.needsUpdate = true;
    },
  };
}

function meshAnimator(mesh: Mesh): Animator {
  const basePos = mesh.position.clone();
  const baseScale = mesh.scale.clone();
  const dir = new Vector3();
  return {
    apply(local, step) {
      const p = clamp01(local / step.duration);
      mesh.visible = p > 0;
      mesh.position.copy(basePos);
      mesh.scale.copy(baseScale);
      if (p <= 0 || p >= 1) return;
      const e = easeOutCubic(p);
      if (step.direction === 'grow') {
        // Grow upward from the slab-on-grade: keep the bottom face fixed.
        mesh.scale.y = baseScale.y * e;
        mesh.position.y = basePos.y - baseScale.y / 2 + mesh.scale.y / 2;
      } else {
        directionFor(step.direction, undefined, basePos.x, basePos.z, dir);
        mesh.position.addScaledVector(dir, (1 - e) * step.distance);
      }
    },
    reset() {
      mesh.visible = true;
      mesh.position.copy(basePos);
      mesh.scale.copy(baseScale);
    },
  };
}

export function buildSteps(cfg: BuildingConfig, layout: BuildingLayout): DemoStep[] {
  const d = layout.derived;
  const maxTop = Math.max(...cfg.roofEquipment.map((e) => e.height)) + d.roofLevel;
  return [
    { group: 'Core', title: 'Core', detail: `Concrete core ${m(cfg.core.width)} × ${m(cfg.core.depth)} m`, start: 0.6, duration: 1.0, direction: 'grow', distance: 0, stagger: 0, staggerBy: 'none' },
    {
      group: 'Slabs',
      title: 'Slabs',
      detail: `GF ${m(cfg.floors.groundFloorHeight)} m · 2F–6F ${m(cfg.floors.typicalFloorHeight)} m · Roof +${m(d.roofLevel)} m`,
      start: 1.3,
      duration: 0.6,
      direction: 'above',
      distance: 6,
      stagger: 0.12,
      staggerBy: 'floor',
    },
    {
      group: 'FacadeFrames',
      title: 'Facade frames',
      detail: `Bay ${m(cfg.bay.spacing)} m · Fin ${m(d.finWidth)} m · Depth ${m(cfg.facade.projectionDepth)} m`,
      start: 2.3,
      duration: 0.6,
      direction: 'outward',
      distance: 4,
      stagger: 0.14,
      staggerBy: 'floor',
    },
    {
      group: 'GlassPanels',
      title: 'Glass panels',
      detail: `${layout.glassPanels.length} panels · Module ${m(cfg.facade.windowModule.width)} × ${m(cfg.facade.windowModule.height)} m`,
      start: 3.4,
      duration: 0.5,
      direction: 'outward',
      distance: 3,
      stagger: 0.12,
      staggerBy: 'floor',
    },
    { group: 'Entrance', title: 'Entrance', detail: `Central entrance ${m(cfg.entrance.width)} × ${m(cfg.entrance.height)} m`, start: 4.5, duration: 0.6, direction: 'front', distance: 4, stagger: 0, staggerBy: 'none' },
    {
      group: 'Canopy',
      title: 'Canopy',
      detail: `${m(cfg.canopy.width)} × ${m(cfg.canopy.projection)} m · underside +${m(cfg.entrance.height)} m`,
      start: 5.0,
      duration: 0.6,
      direction: 'front',
      distance: 5,
      stagger: 0,
      staggerBy: 'none',
    },
    { group: 'Parapet', title: 'Roof / parapet', detail: `Roof zone ${m(cfg.roofZoneHeight)} m · Top +${m(d.parapetTop)} m`, start: 5.5, duration: 0.6, direction: 'above', distance: 5, stagger: 0, staggerBy: 'none' },
    {
      group: 'RoofEquipment',
      title: 'Roof equipment',
      detail: `${cfg.roofEquipment.length} units · top ≤ +${m(maxTop)} m`,
      start: 6.0,
      duration: 0.5,
      direction: 'above',
      distance: 5,
      stagger: 0.08,
      staggerBy: 'index',
    },
  ];
}

export interface DemoCallbacks {
  onCaption(text: string | null): void;
  onFinish(): void;
}

export function createAssemblyDemo(
  built: BuiltBuilding,
  layout: BuildingLayout,
  cfg: BuildingConfig,
  stage: Stage,
  bounds: Box3,
  log: Logger,
  cb: DemoCallbacks,
): AssemblyDemo {
  const steps = buildSteps(cfg, layout);
  const duration = 16.0;
  const lastStep = steps[steps.length - 1] as DemoStep;
  const assemblyEnd = lastStep.start + lastStep.duration + lastStep.stagger * (cfg.roofEquipment.length - 1);

  const animators = new Map<GroupName, Animator[]>();
  for (const name of GROUP_NAMES) {
    const list: Animator[] = [];
    built.groups[name].traverse((o: Object3D) => {
      if (o instanceof InstancedMesh) list.push(instancedAnimator(o));
      else if (o instanceof Mesh) list.push(meshAnimator(o));
    });
    animators.set(name, list);
  }

  // Camera path. Phase A (assembly): swing inside the front-right quadrant so the entrance side
  // stays in view while parts arrive. Phase B: full turntable back to the 3/4 start azimuth.
  const az0 = VIEW_PRESETS.perspective.azimuthDeg;
  const swing = 25;
  const swingEnd = assemblyEnd + 0.4;
  const elStart = 30;
  const elEnd = 22;
  const cameraAt = (t: number): { az: number; el: number } => {
    if (t <= swingEnd) {
      const p = easeInOutSine(t / swingEnd);
      return { az: az0 - swing + 2 * swing * p, el: elStart + (elEnd - elStart) * p };
    }
    return { az: az0 + swing + (360 - 2 * swing) * easeInOutSine((t - swingEnd) / (duration - swingEnd)), el: elEnd };
  };
  // Constant orbit radius that frames the whole building from every azimuth/elevation on the path.
  let radius = 0;
  for (let a = 0; a < 360; a += 15) {
    for (const el of [elEnd, (elStart + elEnd) / 2, elStart]) radius = Math.max(radius, stage.computeFit(bounds, az0 + a, el).distance);
  }
  const target = bounds.getCenter(new Vector3());

  const W = cfg.envelope.width;
  const D = cfg.envelope.depth;
  const H = cfg.envelope.height;
  const captionAt = (t: number): string | null => {
    const first = steps[0] as DemoStep;
    if (t < first.start) return `00 · Site — origin (0, 0, 0) at ground center · footprint ${m(W)} × ${m(D)} m`;
    if (t < assemblyEnd + 0.4) {
      let idx = 0;
      steps.forEach((s, i) => {
        if (s.start <= t) idx = i;
      });
      const s = steps[idx] as DemoStep;
      return `${String(idx + 1).padStart(2, '0')}/${String(steps.length).padStart(2, '0')} · ${s.title} — ${s.detail}`;
    }
    return `Complete — ${m(W)} × ${m(D)} × ${m(H)} m · ${cfg.floors.count} floors · bay ${m(cfg.bay.spacing)} m`;
  };

  let playing = false;
  let startMs = 0;
  let lastCaption: string | null = null;

  const seek = (t: number): void => {
    const time = Math.min(Math.max(t, 0), duration);
    for (const step of steps) {
      for (const a of animators.get(step.group) ?? []) a.apply(time - step.start, step);
    }
    const cam = cameraAt(time);
    stage.camera.position.copy(target).addScaledVector(directionFrom(cam.az, cam.el), radius);
    stage.controls.target.copy(target);
    stage.camera.lookAt(target);
    const caption = captionAt(time);
    if (caption !== lastCaption) {
      lastCaption = caption;
      cb.onCaption(caption);
      log.info('DEMO', 'caption', { t: time, text: caption });
    }
    stage.requestRender();
  };

  const finish = (reason: string): void => {
    for (const list of animators.values()) list.forEach((a) => a.reset());
    playing = false;
    lastCaption = null;
    cb.onCaption(null);
    stage.controls.enabled = true;
    stage.controls.update();
    stage.requestRender();
    cb.onFinish();
    log.info('DEMO', 'finished', { reason });
  };

  log.info('DEMO', 'timeline_ready', {
    duration,
    assemblyEnd: Math.round(assemblyEnd * 100) / 100,
    steps: steps.map((s) => `${s.group}@${s.start}s`),
    orbitRadius: Math.round(radius * 100) / 100,
    swingEnd: Math.round(swingEnd * 100) / 100,
    path: `az ${az0 - swing}->${az0 + swing} (assembly), ${az0 + swing}->${az0 + 360} (turntable); el ${elStart}->${elEnd}`,
  });

  return {
    duration,
    steps,
    isPlaying: () => playing,
    play() {
      if (playing) return;
      playing = true;
      startMs = performance.now();
      stage.controls.enabled = false;
      log.info('DEMO', 'play', { duration });
      seek(0);
    },
    stop() {
      if (playing) finish('stopped');
    },
    seek,
    update(nowMs) {
      if (!playing) return;
      const t = (nowMs - startMs) / 1000;
      if (t >= duration) finish('completed');
      else seek(t);
    },
  };
}
