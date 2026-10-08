// VERSION: v2.5.0 — 2026-10-08 — 내레이션 오디오에 SRT 맞추기 (오디오가 기준 시계)
// 사용: node scripts/align-srt-to-audio.mjs [--audio input/narration.mp3] [--srt input/archive/대본_완성본_v2.4_영상시간.srt]
// 배경: 사용자가 v2.4 SRT 로 만든 내레이션은 자막 칸보다 길게 읽힌 줄이 쌓여 끝에서 약 15초 늦다(486.5초 vs 영상 471.1초).
//       오디오를 바꾸지 않고, 자막·장면 시간을 오디오에 맞춘다.
// 방법
//  1) ffmpeg silencedetect(-38dB, 0.08초 이상)로 말소리 구간을 찾는다.
//  2) 자막 117줄을 말소리 구간 묶음에 순서대로 배정 (동적 계획법):
//     - 줄 i 가 말소리 구간 j..k-1 을 차지, 길이 ≈ 글자 수 × r (r = 전체 말소리 길이 / 전체 글자 수)
//     - 밀림 D_i = 오디오 시작 − SRT 시작 은 거의 줄지 않는다(음성 도구가 넘친 줄만큼 뒤를 민다): D_i ≥ D_{i-1} − 0.2초, D_i ≥ −0.15초
//  3) 독립 검증: "넘친 만큼 민다" 모델 D_{i+1} = D_i + max(0, 말소리_i + pad − 칸_i) 의 잔차를 본다.
//  4) 새 SRT: 시작 = 말 시작 − LEAD(0.1초), 끝 = 다음 줄 시작(이어짐). 원래 SRT 에서 띄어 있던 곳(고지 카드)은
//     끝 = 말 끝 + TAIL(0.3초) 로 두어 공백을 남긴다. 마지막 줄 끝 = 말 끝 + TAIL.
import fs from 'node:fs';
import path from 'node:path';
import {spawnSync} from 'node:child_process';
import {fileURLToPath} from 'node:url';
import {log} from './log.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const arg = (k, d) => {
	const i = process.argv.indexOf(k);
	return i > 0 ? process.argv[i + 1] : d;
};
const AUDIO = path.resolve(root, arg('--audio', 'input/narration.mp3'));
const SRC = path.resolve(root, arg('--srt', 'input/archive/대본_완성본_v2.4_영상시간.srt'));
const OUT = path.join(root, 'input', '대본_완성본.srt');
const REPORT = path.join(root, 'out', 'audio-align.json');
const NOISE = '-38dB';
const MIN_SIL = 0.08;
const LEAD = 0.1;
const TAIL = 0.3;

// ---- 1. 읽기 -------------------------------------------------------------------
const toSec = (t) => {
	const m = /^(\d+):(\d+):(\d+),(\d+)$/.exec(t.trim());
	if (!m) throw new Error(`시간 형식 오류: ${t}`);
	return (+m[1] * 60 + +m[2]) * 60 + +m[3] + +m[4] / 1000;
};
const fmt = (sec) => {
	const ms = Math.round(sec * 1000);
	const h = Math.floor(ms / 3600000);
	const m = Math.floor((ms % 3600000) / 60000);
	const s = Math.floor((ms % 60000) / 1000);
	return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')},${String(ms % 1000).padStart(3, '0')}`;
};
let cues;
try {
	cues = fs
		.readFileSync(SRC, 'utf8')
		.replace(/\r\n/g, '\n')
		.trim()
		.split(/\n\n+/)
		.map((b) => {
			const l = b.split('\n');
			const [a, z] = l[1].split(' --> ');
			return {id: +l[0], start: toSec(a), end: toSec(z), lines: l.slice(2)};
		});
	if (!fs.existsSync(AUDIO)) throw new Error(`오디오 없음: ${AUDIO}`);
} catch (err) {
	log.error('LOAD', {error: err.message, next: '--audio / --srt 경로 확인'});
	process.exit(1);
}

// ffmpeg: 시스템 ffmpeg 이 없으면 Remotion 에 들어 있는 ffmpeg(npx remotion ffmpeg)
const ffmpeg = (args) => {
	const sys = spawnSync('ffmpeg', ['-nostdin', ...args], {encoding: 'utf8', maxBuffer: 64 << 20});
	if (!sys.error) return sys;
	return spawnSync('npx', ['remotion', 'ffmpeg', '-nostdin', ...args], {cwd: root, encoding: 'utf8', maxBuffer: 64 << 20});
};
const det = ffmpeg(['-i', AUDIO, '-af', `silencedetect=noise=${NOISE}:d=${MIN_SIL}`, '-f', 'null', '-']);
const errText = det.stderr ?? '';
const durM = /Duration: (\d+):(\d+):([\d.]+)/.exec(errText);
if (det.status !== 0 || !durM) {
	log.error('SILENCE', {status: det.status, error: errText.slice(-400), next: 'ffmpeg 설치 또는 오디오 파일 확인'});
	process.exit(1);
}
const audioDur = +durM[1] * 3600 + +durM[2] * 60 + +durM[3];
const sil = [];
let cur = null;
for (const m of errText.matchAll(/silence_(start|end): ([\d.]+)/g)) {
	if (m[1] === 'start') cur = +m[2];
	else sil.push([cur, +m[2]]);
}
if (cur !== null && (!sil.length || sil.at(-1)[0] !== cur)) sil.push([cur, audioDur]);
const speech = [];
let prevEnd = 0;
for (const [s, e] of sil) {
	if (s > prevEnd + 1e-3) speech.push([prevEnd, s]);
	prevEnd = e;
}
if (prevEnd < audioDur - 0.05) speech.push([prevEnd, audioDur]);
log.info('LOAD', {audio: path.relative(root, AUDIO), audioSec: audioDur.toFixed(3), srt: path.relative(root, SRC), cues: cues.length, speechSegments: speech.length, noise: NOISE, minSilence: MIN_SIL});

// ---- 2. 정렬 (동적 계획법) ----------------------------------------------------------
const chars = cues.map((c) => [...c.lines.join('')].filter((ch) => !/\s/.test(ch) && !'.,?!…·'.includes(ch)).length);
const N = cues.length;
const M = speech.length;
const r = speech.reduce((a, [s, e]) => a + (e - s), 0) / chars.reduce((a, b) => a + b, 0);
const segCost = (i, j, k) => {
	const d = speech[k - 1][1] - speech[j][0];
	const exp = r * chars[i];
	return ((d - exp) / (0.35 + 0.2 * exp)) ** 2;
};
const INF = Infinity;
const dp = Array.from({length: N + 1}, () => new Float64Array(M + 1).fill(INF));
const back = Array.from({length: N + 1}, () => new Int32Array(M + 1).fill(-1));
const Dv = Array.from({length: N + 1}, () => new Float64Array(M + 1));
dp[0][0] = 0;
Dv[0][0] = speech[0][0] - cues[0].start;
for (let i = 0; i < N; i++) {
	for (let j = 0; j < M; j++) {
		if (dp[i][j] === INF) continue;
		const prevD = Dv[i][j];
		for (let k = j + 1; k <= Math.min(M, j + 8); k++) {
			let c = dp[i][j] + segCost(i, j, k);
			if (i === N - 1) {
				if (k === M && c < dp[N][M]) {
					dp[N][M] = c;
					back[N][M] = j;
				}
				continue;
			}
			if (k >= M) continue;
			const D = speech[k][0] - cues[i + 1].start;
			if (D < -0.15 || D < prevD - 0.2) continue;
			c += 0.3 * Math.max(0, D - prevD);
			if (c < dp[i + 1][k]) {
				dp[i + 1][k] = c;
				back[i + 1][k] = j;
				Dv[i + 1][k] = D;
			}
		}
	}
}
if (dp[N][M] === INF) {
	log.error('ALIGN', {issue: 'no_alignment', next: 'silencedetect 값(NOISE, MIN_SIL) 조정 또는 오디오가 이 SRT 로 만든 것인지 확인'});
	process.exit(1);
}
const startSeg = new Array(N);
startSeg[N - 1] = back[N][M];
for (let i = N - 1; i > 0; i--) startSeg[i - 1] = back[i][startSeg[i]];
const items = cues.map((c, i) => {
	const j = startSeg[i];
	const k = i + 1 < N ? startSeg[i + 1] : M;
	const on = speech[j][0];
	const off = speech[k - 1][1];
	return {id: c.id, srtStart: c.start, on, off, dur: off - on, exp: r * chars[i], drift: on - c.start, segs: k - j, text: c.lines.join(' ')};
});

// ---- 3. 검증 ---------------------------------------------------------------------
const ratios = items.map((x) => x.dur / x.exp);
const outliers = items.filter((x) => x.dur / x.exp < 0.65 || x.dur / x.exp > 1.5);
const driftDrops = items.slice(1).filter((x, i) => x.drift < items[i].drift - 0.15);
let fit = null;
for (let p = 0; p <= 1.2; p += 0.01) {
	const res = items.slice(0, -1).map((x, i) => items[i + 1].drift - (x.drift + Math.max(0, x.dur + p - (cues[i + 1].start - cues[i].start))));
	const err = res.reduce((a, b) => a + Math.abs(b), 0) / res.length;
	if (!fit || err < fit.err) fit = {p, err, big: res.map((v, i) => [items[i + 1].id, v]).filter(([, v]) => Math.abs(v) > 0.35)};
}
log.info('ALIGN', {
	cost: dp[N][M].toFixed(2),
	secPerChar: r.toFixed(4),
	ratioMin: Math.min(...ratios).toFixed(2),
	ratioMax: Math.max(...ratios).toFixed(2),
	outliers: outliers.map((x) => x.id).join(',') || 'none',
	driftFirst: items[0].drift.toFixed(3),
	driftLast: items.at(-1).drift.toFixed(3),
	driftDrops: driftDrops.length,
	pushModelPad: fit.p.toFixed(2),
	pushModelMeanAbsErr: fit.err.toFixed(3),
	pushModelBig: fit.big.map(([id, v]) => `${id}:${v.toFixed(2)}`).join(',') || 'none',
});
if (driftDrops.length || fit.err > 0.15) {
	log.error('ALIGN', {issue: 'alignment_not_trustworthy', driftDrops: driftDrops.length, meanAbsErr: fit.err.toFixed(3), next: 'out/audio-align.json 의 줄별 결과 확인'});
	process.exit(1);
}

// ---- 4. 새 SRT ---------------------------------------------------------------------
const out = items.map((x, i) => ({start: i === 0 ? Math.max(0, x.on - LEAD) : Math.max(x.on - LEAD, items[i - 1].off), lines: cues[i].lines}));
out.forEach((e, i) => {
	const gapInSource = i + 1 < N && cues[i + 1].start - cues[i].end > 0.5; // 원래 SRT 에서 띄어 있던 곳 = 고지 카드
	e.end = i + 1 < N ? (gapInSource ? Math.min(items[i].off + TAIL, out[i + 1].start) : out[i + 1].start) : items[i].off + TAIL;
	if (gapInSource) log.info('GAP', {afterSrt: i + 1, subtitleEnd: fmt(e.end), nextStart: fmt(out[i + 1].start), gapSec: (out[i + 1].start - e.end).toFixed(3), sourceGapSec: (cues[i + 1].start - cues[i].end).toFixed(3)});
});
for (let i = 1; i < out.length; i++) {
	if (out[i].start < out[i - 1].end - 1e-6 || out[i].end <= out[i].start) {
		log.error('VALIDATE', {issue: 'overlap_or_empty', at: i + 1});
		process.exit(1);
	}
}
fs.writeFileSync(OUT, out.map((e, i) => `${i + 1}\n${fmt(e.start)} --> ${fmt(e.end)}\n${e.lines.join('\n')}`).join('\n\n') + '\n');
fs.mkdirSync(path.dirname(REPORT), {recursive: true});
fs.writeFileSync(REPORT, JSON.stringify({audio: path.relative(root, AUDIO), audioSec: audioDur, secPerChar: r, lead: LEAD, tail: TAIL, pushModel: fit, items}, null, 1));
log.info('WRITE', {file: path.relative(root, OUT), entries: out.length, firstStart: fmt(out[0].start), lastEnd: fmt(out.at(-1).end), audioSec: audioDur.toFixed(3), report: path.relative(root, REPORT)});
