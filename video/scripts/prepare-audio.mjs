// VERSION: v2.5.0 — 2026-10-08 — 내레이션/BGM 준비: input/ → public/ 복사 + src/data/audio.ts 생성
// 사용: node scripts/prepare-audio.mjs   (npm run prepare-data 에 포함)
// - input/narration.mp3 (또는 .wav) 가 있으면 public/ 으로 복사하고 길이(초)를 기록한다. 없으면 NARRATION = null (무음 영상).
// - Remotion 은 public/ 의 파일만 staticFile() 로 읽으므로 복사본을 쓴다 (public/ 사본은 git 에 넣지 않음).
// - 가이드 2번: bgm.mp3 가 있으면 내레이션보다 약 -18dB 낮게 → BGM.volumeDb = -18.
import fs from 'node:fs';
import path from 'node:path';
import {spawnSync} from 'node:child_process';
import {fileURLToPath} from 'node:url';
import {log} from './log.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const IN = path.join(root, 'input');
const PUB = path.join(root, 'public');
const OUT = path.join(root, 'src', 'data', 'audio.ts');

const probe = (file) => {
	const args = ['-v', 'error', '-show_entries', 'format=duration', '-of', 'csv=p=0', file];
	let r = spawnSync('ffprobe', args, {encoding: 'utf8'});
	if (r.error) r = spawnSync('npx', ['remotion', 'ffprobe', ...args], {cwd: root, encoding: 'utf8'});
	const sec = parseFloat(String(r.stdout).trim());
	if (r.status !== 0 || !Number.isFinite(sec)) throw new Error(`ffprobe 실패: ${file} ${String(r.stderr).slice(-200)}`);
	return sec;
};

const pick = (base) => ['.mp3', '.wav', '.m4a'].map((e) => path.join(IN, base + e)).find((p) => fs.existsSync(p));
const entry = (base) => {
	const src = pick(base);
	if (!src) {
		log.info('AUDIO', {track: base, status: 'none', note: `input/${base}.mp3 없음`});
		return null;
	}
	fs.mkdirSync(PUB, {recursive: true});
	const file = path.basename(src);
	fs.copyFileSync(src, path.join(PUB, file));
	const durationSec = probe(src);
	log.info('AUDIO', {track: base, src: path.relative(root, src), public: `public/${file}`, durationSec: durationSec.toFixed(3), bytes: fs.statSync(src).size});
	return {file, durationSec: Math.round(durationSec * 1000) / 1000};
};

let narration;
let bgm;
try {
	narration = entry('narration');
	bgm = entry('bgm');
} catch (err) {
	log.error('AUDIO', {error: err.message, next: 'ffprobe(ffmpeg) 설치 또는 오디오 파일 확인'});
	process.exit(1);
}
const ts = `// 자동 생성 파일 — 직접 수정하지 말 것. 생성: node scripts/prepare-audio.mjs
// 원본: input/narration.*, input/bgm.* → public/ 사본을 staticFile() 로 읽는다. null 이면 그 트랙 없음.
export type AudioTrack = {readonly file: string; readonly durationSec: number};
export const NARRATION: AudioTrack | null = ${narration ? JSON.stringify(narration) : 'null'};
export const BGM: (AudioTrack & {readonly volumeDb: number}) | null = ${bgm ? JSON.stringify({...bgm, volumeDb: -18}) : 'null'};
`;
fs.writeFileSync(OUT, ts);
log.info('WRITE', {file: path.relative(root, OUT), narration: Boolean(narration), bgm: Boolean(bgm)});
