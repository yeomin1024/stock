// VERSION: v1.0.0 — 2026-11-02 — Initial implementation: boot pipeline CONFIG -> LAYOUT -> VALIDATE -> BUILD -> RENDER
import './style.css';
import { Box3, type PerspectiveCamera, type Scene } from 'three';
import { buildBuilding, GROUP_NAMES, type GroupName } from './building/buildBuilding';
import { APP_VERSION, BUILDING_CONFIG, UI_CONFIG } from './config';
import { createAssemblyDemo } from './demo';
import { deriveDimensions } from './derive';
import { allBoxes, computeLayout } from './layout';
import { createLogger, formatLine, parseLogLevel, type Logger, type LogStage } from './logger';
import { createMaterials } from './materials';
import { createStage, VIEW_PRESETS, type ViewName } from './stage';
import { createUI } from './ui';
import { boundsOf, validateConfig, validateLayout, type Issue } from './validate';

/** Debug / capture handle (used by scripts/capture.mjs). */
export interface SeoulBuildingApi {
  version: string;
  ready: boolean;
  scene: Scene;
  camera: PerspectiveCamera;
  setView(view: ViewName): void;
  setGrid(visible: boolean): void;
  setLayer(group: GroupName, visible: boolean): void;
  renderNow(): void;
  demo: { duration: number; seek(t: number): void; play(): void; stop(): void };
}

declare global {
  interface Window {
    __SEOUL_BUILDING__?: SeoulBuildingApi;
  }
}

function reportIssues(log: Logger, stage: LogStage, issues: Issue[]): void {
  for (const i of issues) {
    const fields = { message: i.message, ...(i.data ?? {}) };
    if (i.level === 'ERROR') log.error(stage, i.code, fields);
    else if (i.level === 'WARN') log.warn(stage, i.code, fields);
    else log.info(stage, i.code, fields);
  }
}

function webgl2Available(): boolean {
  try {
    return document.createElement('canvas').getContext('webgl2') !== null;
  } catch {
    return false;
  }
}

function boot(log: Logger, params: URLSearchParams): void {
  const t0 = performance.now();
  const captureMode = params.get('capture') === '1';
  const app = document.getElementById('app');
  if (!app) throw new Error('#app container missing from index.html');
  log.info('BOOT', 'start', { version: APP_VERSION, captureMode, logLevel: log.level });

  // ---- 1. CONFIG -----------------------------------------------------------------------
  const cfg = BUILDING_CONFIG;
  const d = deriveDimensions(cfg);
  log.info('CONFIG', 'envelope', { width: cfg.envelope.width, depth: cfg.envelope.depth, height: cfg.envelope.height, units: 'm', up: 'Y', origin: 'ground center (0,0,0)', front: '+Z' });
  log.info('CONFIG', 'floors', { labels: d.floorLabels, heights: d.floorHeights, levels: d.floorLevels, roofLevel: d.roofLevel, parapetTop: d.parapetTop });
  log.info('CONFIG', 'facade', {
    bay: cfg.bay.spacing,
    baysX: d.baysX,
    baysZ: d.baysZ,
    window: `${cfg.facade.windowModule.width}x${cfg.facade.windowModule.height}`,
    sill: cfg.facade.windowModule.sillHeight,
    finWidth: d.finWidth,
    cornerWidth: d.cornerWidth,
    projectionDepth: cfg.facade.projectionDepth,
    glassSetback: d.glassSetback,
  });
  log.debug('CONFIG', 'glazing_rows', { rows: d.glazing.map((g) => `${g.floorLabel}:${g.bottom}-${g.top}`) });
  log.debug('CONFIG', 'bands', { bands: d.bands.map((b) => `${b.label}:${b.bottom}-${b.top}`) });
  reportIssues(log, 'CONFIG', validateConfig(cfg, d));

  // ---- 2. LAYOUT + VALIDATE ------------------------------------------------------------
  const tLayout = performance.now();
  const layout = computeLayout(cfg);
  const report = validateLayout(layout, cfg);
  log.info('LAYOUT', 'computed', { ...report.counts, ms: performance.now() - tLayout });
  reportIssues(log, 'VALIDATE', report.issues);

  // ---- 3. RENDER setup + BUILD ---------------------------------------------------------
  if (!webgl2Available()) throw new Error('WebGL2 is not available in this browser');
  const stage = createStage(app, log, { preserveDrawingBuffer: captureMode });
  const tBuild = performance.now();
  const built = buildBuilding(layout, createMaterials(), log);
  stage.scene.add(built.root);
  built.root.updateMatrixWorld(true);
  log.info('PERF', 'build_done', { ms: performance.now() - tBuild });

  // Cross-check: the three.js scene bounds must equal the pure-layout bounds (instancing sanity).
  const box = new Box3().setFromObject(built.root);
  const expected = boundsOf(allBoxes(layout));
  const sceneMatches = box.min.toArray().every((v, i) => Math.abs(v - (expected.min[i] as number)) < 1e-4) && box.max.toArray().every((v, i) => Math.abs(v - (expected.max[i] as number)) < 1e-4);
  (sceneMatches ? log.info : log.error)('VALIDATE', 'scene_bounds', {
    min: box.min.toArray().map((v) => Math.round(v * 1000) / 1000),
    max: box.max.toArray().map((v) => Math.round(v * 1000) / 1000),
    matchesLayout: sceneMatches,
  });

  stage.setView('perspective', box);
  const gridParam = params.get('grid');
  if (gridParam !== null) stage.setGridVisible(gridParam === '1');

  // ---- 4. UI + demo --------------------------------------------------------------------
  const m = (v: number): string => v.toFixed(1);
  const summary = [
    `${m(cfg.envelope.width)} × ${m(cfg.envelope.depth)} × ${m(cfg.envelope.height)} m (W × D × H)`,
    `GF ${m(cfg.floors.groundFloorHeight)} · 2F–6F ${m(cfg.floors.typicalFloorHeight)} · Roof ${m(cfg.roofZoneHeight)} m`,
    `Bay ${m(cfg.bay.spacing)} · Window ${m(cfg.facade.windowModule.width)} × ${m(cfg.facade.windowModule.height)} · Depth ${cfg.facade.projectionDepth} m`,
    `Entrance ${m(cfg.entrance.width)} × ${m(cfg.entrance.height)} m · ${built.stats.total.triangles.toLocaleString('en-US')} tris · ${built.stats.total.drawCalls} draws`,
  ];
  let demoRef: ReturnType<typeof createAssemblyDemo> | null = null;
  const ui = createUI(
    document.body,
    {
      version: APP_VERSION,
      groups: GROUP_NAMES,
      views: (Object.keys(VIEW_PRESETS) as ViewName[]).map((name) => ({ name, label: VIEW_PRESETS[name].label })),
      gridVisible: stage.grid.visible,
      summary,
      captureMode,
    },
    {
      onGridToggle: (v) => {
        stage.setGridVisible(v);
        log.info('UI', 'grid_toggled', { visible: v });
      },
      onLayerToggle: (g, v) => {
        built.groups[g].visible = v;
        stage.requestRender();
        log.info('UI', 'layer_toggled', { group: g, visible: v });
      },
      onView: (v) => {
        demoRef?.stop();
        stage.setView(v, box);
      },
      onDemo: () => {
        if (!demoRef) return;
        if (demoRef.isPlaying()) demoRef.stop();
        else {
          demoRef.play();
          ui.setDemoPlaying(true);
        }
      },
    },
  );
  const demo = createAssemblyDemo(built, layout, cfg, stage, box, log, {
    onCaption: (t) => ui.setCaption(t),
    onFinish: () => ui.setDemoPlaying(false),
  });
  demoRef = demo;
  ui.setGridChecked(stage.grid.visible);
  log.info('UI', 'ready', { gridVisible: stage.grid.visible, gridSpacing: UI_CONFIG.grid.size / UI_CONFIG.grid.divisions, layers: GROUP_NAMES });

  // ---- 5. Render loop (on demand: draws only when the camera/scene changed) -------------
  const loop = (now: number): void => {
    requestAnimationFrame(loop);
    demo.update(now);
    if (!demo.isPlaying()) stage.controls.update();
    stage.renderIfNeeded();
  };
  stage.renderNow();
  log.info('PERF', 'first_frame', { bootMs: performance.now() - t0 });
  requestAnimationFrame(loop);

  window.__SEOUL_BUILDING__ = {
    version: APP_VERSION,
    ready: true,
    scene: stage.scene,
    camera: stage.camera,
    setView: (v) => stage.setView(v, box),
    setGrid: (v) => {
      stage.setGridVisible(v);
      ui.setGridChecked(v);
    },
    setLayer: (g, v) => {
      built.groups[g].visible = v;
      stage.requestRender();
    },
    renderNow: () => stage.renderNow(),
    demo: { duration: demo.duration, seek: demo.seek, play: demo.play, stop: demo.stop },
  };
}

const params = new URLSearchParams(window.location.search);
const log = createLogger(parseLogLevel(params.get('log')));
try {
  boot(log, params);
} catch (err) {
  const message = err instanceof Error ? err.message : String(err);
  console.error(formatLine('ERROR', 'BOOT', 'boot_failed', { message, hint: 'Open in a WebGL2 browser (recent Chrome/Edge/Firefox/Safari); run with ?log=debug for stage details' }));
  const box = document.createElement('div');
  box.className = 'error';
  box.textContent = `3D 뷰를 시작하지 못했습니다: ${message}. WebGL2를 지원하는 최신 브라우저에서 열어주세요.`;
  document.body.append(box);
}
