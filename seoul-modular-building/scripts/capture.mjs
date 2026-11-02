// VERSION: v1.0.0 — 2026-11-02 — Initial implementation: headless stills + frame-exact demo video
/**
 * Optional tooling (not needed to run the app): renders stills of each view preset and a
 * frame-exact MP4 of the assembly demo by calling window.__SEOUL_BUILDING__.demo.seek(t) per
 * frame, so the video never depends on the headless browser's real-time frame rate.
 *
 * Requirements: Playwright + Chromium, and ffmpeg (with libx264) on PATH.
 *   npm i -D playwright && npx playwright install chromium
 *
 * Usage (with `npm run build && npm run preview` running on :4173):
 *   node scripts/capture.mjs [--url http://localhost:4173/] [--out captures] [--fps 30]
 *                            [--width 1280] [--height 720] [--only stills|video]
 */
import { execFileSync } from 'node:child_process';
import { mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { join, resolve } from 'node:path';

const require = createRequire(import.meta.url);
const args = Object.fromEntries(
  process.argv.slice(2).reduce((acc, a, i, all) => (a.startsWith('--') ? [...acc, [a.slice(2), all[i + 1]]] : acc), []),
);
const URL_BASE = args.url ?? 'http://localhost:4173/';
const OUT = resolve(args.out ?? 'captures');
const FPS = Number(args.fps ?? 30);
const WIDTH = Number(args.width ?? 1280);
const HEIGHT = Number(args.height ?? 720);
const ONLY = args.only ?? 'all';

const log = (stage, event, fields = {}) =>
  console.log(`[${stage}] [${new Date().toISOString()}] level=INFO event=${event} ${Object.entries(fields).map(([k, v]) => `${k}=${v}`).join(' ')}`);

let chromium;
try {
  ({ chromium } = require('playwright'));
} catch (err) {
  console.error(`[CAPTURE] level=ERROR event=playwright_missing message="${err.message}" hint="npm i -D playwright && npx playwright install chromium"`);
  process.exit(1);
}

mkdirSync(OUT, { recursive: true });
const consoleLines = [];

async function openPage(browser, query) {
  const page = await browser.newPage({ viewport: { width: WIDTH, height: HEIGHT }, deviceScaleFactor: 1 });
  page.on('console', (msg) => consoleLines.push(msg.text()));
  page.on('pageerror', (err) => consoleLines.push(`[PAGEERROR] ${err.message}`));
  const url = `${URL_BASE}?${query}`;
  try {
    await page.goto(url, { waitUntil: 'load' });
    await page.waitForFunction(() => window.__SEOUL_BUILDING__?.ready === true, null, { timeout: 60_000 });
  } catch (err) {
    console.error(`[CAPTURE] level=ERROR event=page_not_ready url=${url} message="${err.message}" hint="is 'npm run preview' running?"`);
    throw err;
  }
  return page;
}

async function captureStills(browser) {
  const page = await openPage(browser, 'log=debug&grid=0');
  for (const view of ['perspective', 'front', 'right', 'rear', 'top']) {
    await page.evaluate((v) => {
      window.__SEOUL_BUILDING__.setView(v);
      window.__SEOUL_BUILDING__.renderNow();
    }, view);
    const path = join(OUT, `view-${view}.png`);
    await page.screenshot({ path });
    log('CAPTURE', 'still_saved', { view, path });
  }
  await page.close();
}

async function captureVideo(browser) {
  const page = await openPage(browser, 'capture=1&grid=1');
  const duration = await page.evaluate(() => window.__SEOUL_BUILDING__.demo.duration);
  const frameDir = join(OUT, 'frames');
  rmSync(frameDir, { recursive: true, force: true });
  mkdirSync(frameDir, { recursive: true });
  const total = Math.round(duration * FPS) + FPS; // +1 s hold on the finished building
  const t0 = Date.now();
  for (let i = 0; i < total; i++) {
    const t = Math.min(i / FPS, duration);
    await page.evaluate((time) => {
      window.__SEOUL_BUILDING__.demo.seek(time);
      window.__SEOUL_BUILDING__.renderNow();
    }, t);
    await page.screenshot({ path: join(frameDir, `f${String(i).padStart(4, '0')}.png`) });
    if (i % FPS === 0) log('CAPTURE', 'frame', { index: i, of: total, t: t.toFixed(2), elapsedS: ((Date.now() - t0) / 1000).toFixed(1) });
  }
  await page.close();
  const mp4 = join(OUT, 'seoul-modular-building-demo.mp4');
  try {
    execFileSync('ffmpeg', ['-y', '-loglevel', 'error', '-framerate', String(FPS), '-i', join(frameDir, 'f%04d.png'), '-c:v', 'libx264', '-pix_fmt', 'yuv420p', '-crf', '20', '-preset', 'slow', '-movflags', '+faststart', mp4], { stdio: 'inherit' });
  } catch (err) {
    console.error(`[CAPTURE] level=ERROR event=ffmpeg_failed message="${err.message}" hint="install ffmpeg with libx264; frames kept in ${frameDir}"`);
    throw err;
  }
  log('CAPTURE', 'video_saved', { path: mp4, frames: total, fps: FPS, seconds: (total / FPS).toFixed(1), renderS: ((Date.now() - t0) / 1000).toFixed(1) });
}

const browser = await chromium.launch({ args: ['--use-angle=swiftshader', '--enable-unsafe-swiftshader', '--ignore-gpu-blocklist'] });
try {
  if (ONLY === 'all' || ONLY === 'stills') await captureStills(browser);
  if (ONLY === 'all' || ONLY === 'video') await captureVideo(browser);
} finally {
  await browser.close();
  writeFileSync(join(OUT, 'browser-console.log'), consoleLines.join('\n') + '\n');
  log('CAPTURE', 'console_saved', { path: join(OUT, 'browser-console.log'), lines: consoleLines.length });
}
