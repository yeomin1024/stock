// VERSION: v1.0.0 — 2026-10-05 — 검수용 still 렌더 (가이드 6번 작업 순서 6)
// 사용: node scripts/render-stills.mjs [S01 S05 ...] [--extra=프레임,프레임]
// 장면마다 대표 프레임 2장: a = 첫 자막 시작 + 1.5초(장면 길이 안으로 제한), b = 마지막 자막 끝 직전(다음 장면 와이프 시작 전).
// 번들은 한 번만 만들고 renderStill 로 프레임별 PNG 를 out/stills/ 에 쓴다.
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {bundle} from '@remotion/bundler';
import {renderStill, selectComposition} from '@remotion/renderer';
import {log} from './log.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const OUT = path.join(root, 'out', 'stills');
fs.mkdirSync(OUT, {recursive: true});

const tl = await import(path.join(root, 'src', 'data', 'timeline.ts'));
const args = process.argv.slice(2);
const only = args.filter((a) => /^S\d\d$/.test(a));
const extra = (args.find((a) => a.startsWith('--extra=')) ?? '--extra=').slice(8).split(',').filter(Boolean).map(Number);

const jobs = [];
for (const s of tl.SCENES) {
	if (only.length && !only.includes(s.id)) continue;
	const {start, end} = tl.sceneRange(s.id);
	const lastEnd = s.subs.length ? tl.subEnd(s.subs.at(-1)) : end;
	const firstStart = s.subs.length ? tl.subStart(s.subs[0]) : start;
	const a = Math.min(firstStart + 45, end - 1);
	const next = tl.SCENES[tl.SCENES.indexOf(s) + 1];
	const preroll = next && next.enter === 'wipe' ? tl.WIPE.pre : 0;
	const b = Math.max(a + 1, Math.min(lastEnd, end) - 8 - preroll);
	jobs.push({name: `${s.id}_a`, frame: a}, {name: `${s.id}_b`, frame: b});
}
for (const fr of extra) jobs.push({name: `extra_${String(fr).padStart(5, '0')}`, frame: fr});

const t0 = Date.now();
log.info('BUNDLE', {entry: 'src/index.ts'});
let serveUrl;
try {
	serveUrl = await bundle({entryPoint: path.join(root, 'src', 'index.ts')});
} catch (err) {
	log.error('BUNDLE', {error: err.message, next: 'npx tsc --noEmit 로 타입 오류 확인'});
	process.exit(1);
}
const composition = await selectComposition({serveUrl, id: 'Molbbang'});
log.info('BUNDLE', {ms: Date.now() - t0, durationInFrames: composition.durationInFrames, jobs: jobs.length});

for (const j of jobs) {
	const t1 = Date.now();
	const output = path.join(OUT, `${j.name}.png`);
	try {
		await renderStill({composition, serveUrl, frame: j.frame, output, imageFormat: 'png', overwrite: true});
		log.info('STILL', {name: j.name, frame: j.frame, sec: (j.frame / tl.FPS).toFixed(2), ms: Date.now() - t1});
	} catch (err) {
		log.error('STILL', {name: j.name, frame: j.frame, error: err.message});
	}
}
log.info('DONE', {stills: jobs.length, totalMs: Date.now() - t0, dir: path.relative(root, OUT)});
