// VERSION: v2.0.0 — 2026-10-06 — 스토리보드 still 렌더 (가이드 v2 6-5) + manifest.json
// 사용: node scripts/render-storyboard.mjs [S05 S12 ...]
// - 장면마다 "모든 요소가 다 나온 시점(장면 끝 10프레임 전)" → out/storyboard/S01.png ~ S30.png
// - S14, S16~S19, S25~S29 는 자막 단위로 1장씩 더 → S25-66.png 처럼
// - 번들은 한 번만 만들고 renderStill 로 PNG 를 쓴다. 카드 문구는 manifest.json 에 모은다 (index.html 생성용).
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {bundle} from '@remotion/bundler';
import {renderStill, selectComposition} from '@remotion/renderer';
import {log} from './log.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const OUT = path.join(root, 'out', 'storyboard');
fs.mkdirSync(OUT, {recursive: true});

const tl = await import(path.join(root, 'src', 'data', 'timeline.ts'));
const {SUB_WHY} = await import(path.join(root, 'src', 'data', 'storyboard.ts'));
const only = process.argv.slice(2).filter((a) => /^S\d\d$/.test(a));
const shots = tl.storyboardShots().filter((s) => !only.length || only.includes(s.scene));

const t0 = Date.now();
log.info('BUNDLE', {entry: 'src/index.ts', shots: shots.length});
let serveUrl;
try {
	serveUrl = await bundle({entryPoint: path.join(root, 'src', 'index.ts')});
} catch (err) {
	log.error('BUNDLE', {error: err.message, next: 'npx tsc --noEmit 로 타입 오류 확인'});
	process.exit(1);
}
const composition = await selectComposition({serveUrl, id: 'Molbbang'});
log.info('BUNDLE', {ms: Date.now() - t0, durationInFrames: composition.durationInFrames});

let failed = 0;
for (const s of shots) {
	const t1 = Date.now();
	try {
		await renderStill({composition, serveUrl, frame: s.frame, output: path.join(OUT, s.file), imageFormat: 'png', overwrite: true});
		log.info('STILL', {file: s.file, frame: s.frame, sec: (s.frame / tl.FPS).toFixed(2), ms: Date.now() - t1});
	} catch (err) {
		failed++;
		log.error('STILL', {file: s.file, frame: s.frame, error: err.message});
	}
}

// manifest: 전체 still 목록 (부분 렌더여도 전체 목록을 쓴다)
const manifest = tl.storyboardShots().map((s) => {
	const def = tl.SCENES.find((x) => x.id === s.scene);
	const {start, end} = tl.sceneRange(s.scene);
	const isSub = s.file.includes('-');
	return {
		file: s.file,
		scene: s.scene,
		frame: s.frame,
		time: `${Math.floor(s.frame / tl.FPS / 60)}:${String(Math.floor((s.frame / tl.FPS) % 60)).padStart(2, '0')}`,
		sceneSubs: def.subs.length ? `${def.subs[0]}–${def.subs[def.subs.length - 1]}` : '고지 카드 6초',
		subs: s.subs.map((n) => ({n, text: tl.subText(n)})),
		why: def.perSub ? SUB_WHY[s.subs[0]] ?? def.why : def.why,
		sceneWhy: def.why,
		isSub,
		enter: def.enter,
		sceneStart: start,
		sceneEnd: end,
	};
});
fs.writeFileSync(path.join(OUT, 'manifest.json'), JSON.stringify(manifest, null, 1));
log.info('DONE', {stills: shots.length, failed, totalMs: Date.now() - t0, dir: path.relative(root, OUT)});
if (failed) process.exit(1);
