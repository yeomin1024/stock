// VERSION: v1.0.0 — 2026-11-02 — Initial implementation: renderer, lights, ground, grid, camera fit, controls
import {
  Box3,
  Color,
  DirectionalLight,
  Fog,
  GridHelper,
  HemisphereLight,
  Mesh,
  MeshStandardMaterial,
  NeutralToneMapping,
  PCFShadowMap,
  PerspectiveCamera,
  PlaneGeometry,
  PMREMGenerator,
  REVISION,
  Scene,
  SRGBColorSpace,
  Vector3,
  WebGLRenderer,
} from 'three';
import { OrbitControls } from 'three/addons/controls/OrbitControls.js';
import { RoomEnvironment } from 'three/addons/environments/RoomEnvironment.js';
import { MATERIAL_CONFIG, RENDER_CONFIG, UI_CONFIG } from './config';
import type { Logger } from './logger';

export type ViewName = 'perspective' | 'front' | 'right' | 'rear' | 'top';

/** Azimuth measured from +Z (front) toward +X (right); elevation above the horizon. */
export const VIEW_PRESETS: Record<ViewName, { label: string; azimuthDeg: number; elevationDeg: number }> = {
  perspective: { label: '3/4', azimuthDeg: 35, elevationDeg: 20 },
  front: { label: 'Front', azimuthDeg: 0, elevationDeg: 3 },
  right: { label: 'Right', azimuthDeg: 90, elevationDeg: 3 },
  rear: { label: 'Rear', azimuthDeg: 180, elevationDeg: 3 },
  top: { label: 'Top', azimuthDeg: 0, elevationDeg: 89.5 },
};

export interface Stage {
  renderer: WebGLRenderer;
  scene: Scene;
  camera: PerspectiveCamera;
  controls: OrbitControls;
  grid: GridHelper;
  requestRender(): void;
  /** Renders if something changed since the last frame (on-demand rendering). Returns true if it drew. */
  renderIfNeeded(): boolean;
  renderNow(): void;
  /** Exact framing of a box for a direction: no part of the box leaves the frustum. */
  computeFit(box: Box3, azimuthDeg: number, elevationDeg: number, margin?: number): { position: Vector3; target: Vector3; distance: number };
  setView(view: ViewName, box: Box3): void;
  setGridVisible(visible: boolean): void;
}

const DEG = Math.PI / 180;

export function directionFrom(azimuthDeg: number, elevationDeg: number): Vector3 {
  const az = azimuthDeg * DEG;
  const el = elevationDeg * DEG;
  return new Vector3(Math.sin(az) * Math.cos(el), Math.sin(el), Math.cos(az) * Math.cos(el));
}

export function createStage(container: HTMLElement, log: Logger, opts: { preserveDrawingBuffer: boolean }): Stage {
  const rc = RENDER_CONFIG;
  const renderer = new WebGLRenderer({ antialias: true, preserveDrawingBuffer: opts.preserveDrawingBuffer, powerPreference: 'high-performance' });
  renderer.setPixelRatio(Math.min(window.devicePixelRatio, rc.maxPixelRatio));
  renderer.setSize(container.clientWidth, container.clientHeight);
  renderer.outputColorSpace = SRGBColorSpace;
  renderer.toneMapping = NeutralToneMapping; // Khronos PBR Neutral: keeps material base colors faithful
  renderer.toneMappingExposure = rc.toneMappingExposure;
  renderer.shadowMap.enabled = true;
  renderer.shadowMap.type = PCFShadowMap; // r186: PCF + shadow.radius = soft shadows (PCFSoft was removed)
  container.appendChild(renderer.domElement);

  const scene = new Scene();
  scene.background = new Color(rc.background);
  scene.fog = new Fog(rc.background, rc.fog.near, rc.fog.far);

  // Subtle image-based lighting so the metal frames and glass get reflections (no external files).
  const pmrem = new PMREMGenerator(renderer);
  scene.environment = pmrem.fromScene(new RoomEnvironment(), 0.04).texture;
  scene.environmentIntensity = rc.environmentIntensity;
  pmrem.dispose();

  const hemi = new HemisphereLight(rc.hemisphere.sky, rc.hemisphere.ground, rc.hemisphere.intensity);
  hemi.name = 'HemisphereLight';
  scene.add(hemi);

  const sun = new DirectionalLight(rc.sun.color, rc.sun.intensity);
  sun.name = 'SunLight';
  sun.position.fromArray(rc.sun.position);
  sun.target.position.fromArray(rc.sun.target);
  sun.castShadow = true;
  sun.shadow.mapSize.set(rc.sun.shadowMapSize, rc.sun.shadowMapSize);
  sun.shadow.radius = rc.sun.shadowRadius;
  sun.shadow.bias = rc.sun.shadowBias;
  sun.shadow.normalBias = rc.sun.shadowNormalBias;
  const sc = sun.shadow.camera;
  sc.left = -rc.sun.shadowExtent;
  sc.right = rc.sun.shadowExtent;
  sc.top = rc.sun.shadowExtent;
  sc.bottom = -rc.sun.shadowExtent;
  sc.near = 1;
  sc.far = 140;
  sc.updateProjectionMatrix();
  scene.add(sun, sun.target);

  // Ground: receives the building's shadow. polygonOffset pushes it back so the GF slab top
  // (also at y = 0) always wins the depth test instead of z-fighting.
  const groundMat = new MeshStandardMaterial({
    name: 'Ground',
    color: MATERIAL_CONFIG.ground.color,
    roughness: MATERIAL_CONFIG.ground.roughness,
    metalness: MATERIAL_CONFIG.ground.metalness,
    polygonOffset: true,
    polygonOffsetFactor: 1,
    polygonOffsetUnits: 4,
  });
  const ground = new Mesh(new PlaneGeometry(rc.ground.size, rc.ground.size), groundMat);
  ground.name = 'Ground';
  ground.rotation.x = -Math.PI / 2;
  ground.receiveShadow = true;
  scene.add(ground);

  const gc = UI_CONFIG.grid;
  const grid = new GridHelper(gc.size, gc.divisions, gc.centerColor, gc.lineColor);
  grid.name = 'DevGridHelper';
  grid.position.y = 0.01;
  grid.visible = gc.visibleByDefault;
  scene.add(grid);

  const camera = new PerspectiveCamera(rc.camera.fov, container.clientWidth / Math.max(1, container.clientHeight), rc.camera.near, rc.camera.far);
  const controls = new OrbitControls(camera, renderer.domElement);
  controls.enableDamping = true;
  controls.dampingFactor = 0.08;
  controls.minDistance = rc.controls.minDistance;
  controls.maxDistance = rc.controls.maxDistance;
  controls.maxPolarAngle = rc.controls.maxPolarAngle;

  let needsRender = true;
  const requestRender = (): void => {
    needsRender = true;
  };
  controls.addEventListener('change', requestRender);

  const renderNow = (): void => {
    renderer.render(scene, camera);
    needsRender = false;
  };

  const resize = (): void => {
    const w = container.clientWidth;
    const h = Math.max(1, container.clientHeight);
    camera.aspect = w / h;
    camera.updateProjectionMatrix();
    renderer.setSize(w, h);
    requestRender();
    log.debug('RENDER', 'resize', { width: w, height: h, pixelRatio: renderer.getPixelRatio() });
  };
  new ResizeObserver(resize).observe(container);

  const computeFit: Stage['computeFit'] = (box, azimuthDeg, elevationDeg, margin = rc.camera.fitMargin) => {
    const target = box.getCenter(new Vector3());
    const dir = directionFrom(azimuthDeg, elevationDeg); // target -> camera
    const forward = dir.clone().negate();
    const right = new Vector3().crossVectors(forward, new Vector3(0, 1, 0)).normalize();
    const up = new Vector3().crossVectors(right, forward).normalize();
    const tanV = Math.tan((camera.fov * DEG) / 2);
    const tanH = tanV * camera.aspect;
    let distance = 0;
    const corner = new Vector3();
    for (let i = 0; i < 8; i++) {
      corner.set(i & 1 ? box.max.x : box.min.x, i & 2 ? box.max.y : box.min.y, i & 4 ? box.max.z : box.min.z).sub(target);
      const towardCamera = corner.dot(dir);
      // A point at lateral offset l and depth d toward the camera fits if l / (D - d) <= tan(fov/2).
      distance = Math.max(distance, (Math.abs(corner.dot(right)) * margin) / tanH + towardCamera, (Math.abs(corner.dot(up)) * margin) / tanV + towardCamera);
    }
    return { position: target.clone().addScaledVector(dir, distance), target, distance };
  };

  const setView: Stage['setView'] = (view, box) => {
    const p = VIEW_PRESETS[view];
    const fit = computeFit(box, p.azimuthDeg, p.elevationDeg);
    camera.position.copy(fit.position);
    controls.target.copy(fit.target);
    camera.lookAt(fit.target);
    controls.update();
    requestRender();
    log.info('CAMERA', 'view_set', {
      view,
      azimuthDeg: p.azimuthDeg,
      elevationDeg: p.elevationDeg,
      distance: fit.distance,
      target: fit.target.toArray().map((v) => Math.round(v * 100) / 100),
      aspect: camera.aspect,
    });
  };

  const setGridVisible = (visible: boolean): void => {
    grid.visible = visible;
    requestRender();
  };

  log.info('RENDER', 'renderer_ready', {
    three: REVISION,
    toneMapping: 'NeutralToneMapping',
    exposure: renderer.toneMappingExposure,
    outputColorSpace: renderer.outputColorSpace,
    shadowType: 'PCFShadowMap',
    shadowMap: rc.sun.shadowMapSize,
    shadowRadius: rc.sun.shadowRadius,
    pixelRatio: renderer.getPixelRatio(),
    size: `${container.clientWidth}x${container.clientHeight}`,
  });

  return {
    renderer,
    scene,
    camera,
    controls,
    grid,
    requestRender,
    renderIfNeeded: () => {
      if (!needsRender) return false;
      renderNow();
      return true;
    },
    renderNow,
    computeFit,
    setView,
    setGridVisible,
  };
}
