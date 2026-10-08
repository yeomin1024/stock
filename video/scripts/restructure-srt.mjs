// VERSION: v2.4.0 — 2026-10-08 — 자막 한 줄화 + 읽기 속도 조정 + 고지 카드 3.5초 공백 (사용자 요청, 일회성 변환 도구)
// v2.4.0: SRT 를 영상 시간과 똑같이 — 문장 18 뒤에 고지 카드 공백을 넣는다 (전에는 영상 코드에서만 밀었음). 사용자 요청으로 6초 → 3.5초.
// 사용: node scripts/restructure-srt.mjs [--in input/archive/대본_완성본_v2.2_두줄.srt] [--speed 1.1]
// 규칙
//  1) 두 줄 자막을 이어 붙여 한 줄 폭(아래 MAX_TEXT_W)에 들어가면 → 한 줄로 합친다.
//  2) 들어가지 않으면 → 원래 줄바꿈 위치에서 자막 두 개로 나눈다. 시간은 글자 수(공백 제외) 비율로 나눈다.
//  3) 예외: "앞의 사연과 같이" 다음의 혼잣말(다음 자막이 "이런 식으로"로 시작하기 전까지)은 원래 모양(두 줄) 그대로.
//  4) 모든 시간을 SPEED 배 빠르게 (t' = t / SPEED).
//  5) 고지 카드: 입력 자막 NOTICE_AFTER(=18, 가이드 "자막 18번이 끝난 직후") 다음 자막부터 NOTICE_SEC(3.5초, 가이드 6초를 사용자 요청으로 줄임) 뒤로 민다 → SRT 시간 = 영상 시간.
//     (속도 조정 뒤에 더하므로 고지 카드는 1.1배와 상관없이 3.5초.) 영상 끝 여유 1초만 timeline.ts 에서 더한다.
// 폭 측정: 영상 자막과 같은 글꼴(Noto Sans KR 700, 46px)을 헤드리스 Chromium(Playwright)에서 canvas measureText 로 잰다.
// MAX_TEXT_W = 자막 띠 최대 폭 1500 − 좌우 안쪽 여백 34×2 (src/components/Subtitles.tsx 의 MAX_W, PAD_X).
import fs from 'node:fs';
import path from 'node:path';
import {createRequire} from 'node:module';
import {execSync} from 'node:child_process';
import {fileURLToPath} from 'node:url';
import {log} from './log.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const arg = (k, d) => {
	const i = process.argv.indexOf(k);
	return i > 0 ? process.argv[i + 1] : d;
};
const IN = path.resolve(root, arg('--in', 'input/archive/대본_완성본_v2.2_두줄.srt'));
const OUT = path.join(root, 'input', '대본_완성본.srt');
const SPEED = Number(arg('--speed', '1.1'));
const MAX_TEXT_W = 1500 - 34 * 2;
const NOTICE_AFTER = Number(arg('--notice-after', '18')); // 입력 SRT(두 줄 v2.2)의 자막 번호 = 대본 문장 번호
const NOTICE_SEC = Number(arg('--notice-sec', '3.5'));
const FONT = "700 46px 'Noto Sans KR'";

// ---- 1. 읽기 ------------------------------------------------------------------
const toMs = (t) => {
	const m = /^(\d+):(\d+):(\d+),(\d+)$/.exec(t.trim());
	if (!m) throw new Error(`시간 형식 오류: ${t}`);
	return ((+m[1] * 60 + +m[2]) * 60 + +m[3]) * 1000 + +m[4];
};
const fmt = (ms) => {
	const h = Math.floor(ms / 3600000);
	const m = Math.floor((ms % 3600000) / 60000);
	const s = Math.floor((ms % 60000) / 1000);
	const x = ms % 1000;
	return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')},${String(x).padStart(3, '0')}`;
};
let cues;
try {
	cues = fs
		.readFileSync(IN, 'utf8')
		.replace(/\r\n/g, '\n')
		.trim()
		.split(/\n\n+/)
		.map((b) => {
			const l = b.split('\n');
			const [a, z] = l[1].split(' --> ');
			return {id: +l[0], startMs: toMs(a), endMs: toMs(z), lines: l.slice(2)};
		});
} catch (err) {
	log.error('LOAD', {file: path.relative(root, IN), error: err.message, next: '입력 SRT 경로(--in) 확인'});
	process.exit(1);
}
log.info('LOAD', {file: path.relative(root, IN), cues: cues.length, twoLine: cues.filter((c) => c.lines.length === 2).length, speed: SPEED, maxTextW: MAX_TEXT_W, noticeAfter: NOTICE_AFTER, noticeSec: NOTICE_SEC});

// 예외: "앞의 사연과 같이" 다음 혼잣말
const anchor = cues.findIndex((c) => c.lines.join(' ') === '앞의 사연과 같이');
const keep = new Set();
for (let i = anchor + 1; i < cues.length && !cues[i].lines[0].startsWith('이런 식으로'); i++) keep.add(cues[i].id);
if (anchor < 0 || keep.size === 0) {
	log.error('RULE', {issue: '혼잣말 예외 구간을 찾지 못함', next: 'SRT 에 "앞의 사연과 같이" 자막이 있는지 확인'});
	process.exit(1);
}
log.info('RULE', {exceptionCues: [...keep].join(','), reason: '앞의 사연 다음 혼잣말은 원래 모양 유지'});

// ---- 2. 폭 측정 (Playwright Chromium) --------------------------------------------
let chromium;
try {
	const require = createRequire(import.meta.url);
	({chromium} = require(path.join(execSync('npm root -g').toString().trim(), 'playwright')));
} catch (err) {
	log.error('MEASURE', {error: err.message, next: 'npm i -g playwright (브라우저는 /opt/pw-browsers 사용)'});
	process.exit(1);
}
const texts = cues.map((c) => c.lines.join(' '));
const browser = await chromium.launch();
const page = await browser.newPage();
await page.setContent(`<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@700&display=block">`);
const widths = await page.evaluate(
	async ({texts, FONT}) => {
		await document.fonts.load(FONT, texts.join(''));
		await document.fonts.ready;
		if (!document.fonts.check(FONT, texts.join(''))) throw new Error('Noto Sans KR 700 로드 실패');
		const c = document.createElement('canvas').getContext('2d');
		c.font = FONT;
		return texts.map((t) => Math.round(c.measureText(t).width));
	},
	{texts, FONT},
);
await browser.close();

// ---- 3. 합치기 / 나누기 / 유지 + 속도 ----------------------------------------------
const scale = (ms) => Math.round(ms / SPEED);
const chars = (s) => s.replace(/\s/g, '').length;
const out = [];
const stats = {join: 0, split: 0, keep: 0, single: 0};
cues.forEach((c, i) => {
	const w = widths[i];
	if (keep.has(c.id) || c.lines.length === 1) {
		const kind = keep.has(c.id) ? 'keep' : 'single';
		stats[kind]++;
		out.push({from: c.id, startMs: scale(c.startMs), endMs: scale(c.endMs), lines: c.lines});
		log.debug('CUE', {id: c.id, action: kind, width: w});
		return;
	}
	if (w <= MAX_TEXT_W) {
		stats.join++;
		out.push({from: c.id, startMs: scale(c.startMs), endMs: scale(c.endMs), lines: [c.lines.join(' ')]});
		log.debug('CUE', {id: c.id, action: 'join', width: w});
		return;
	}
	stats.split++;
	const [a, b] = c.lines;
	const mid = c.startMs + Math.round(((c.endMs - c.startMs) * chars(a)) / (chars(a) + chars(b)));
	out.push({from: c.id, startMs: scale(c.startMs), endMs: scale(mid), lines: [a]});
	out.push({from: c.id, startMs: scale(mid), endMs: scale(c.endMs), lines: [b]});
	log.info('CUE', {id: c.id, action: 'split', width: w, line1: a, line2: b, midSec: (scale(mid) / 1000).toFixed(3)});
});

// 고지 카드 공백: 입력 자막 NOTICE_AFTER 보다 뒤에서 나온 자막은 전부 NOTICE_SEC 뒤로
for (const e of out) {
	if (e.from > NOTICE_AFTER) {
		e.startMs += NOTICE_SEC * 1000;
		e.endMs += NOTICE_SEC * 1000;
	}
}
const lastBefore = out.filter((e) => e.from <= NOTICE_AFTER).at(-1);
const firstAfter = out.find((e) => e.from > NOTICE_AFTER);
log.info('NOTICE', {afterCue: NOTICE_AFTER, gapFrom: fmt(lastBefore.endMs), gapTo: fmt(firstAfter.startMs), gapMs: firstAfter.startMs - lastBefore.endMs});

// 검증: 시간 단조 증가·겹침 없음
for (let i = 1; i < out.length; i++) {
	if (out[i].startMs < out[i - 1].endMs || out[i].endMs <= out[i].startMs) {
		log.error('VALIDATE', {issue: 'overlap_or_empty', at: i + 1, prevEnd: out[i - 1].endMs, start: out[i].startMs});
		process.exit(1);
	}
}
const srt = out.map((e, i) => `${i + 1}\n${fmt(e.startMs)} --> ${fmt(e.endMs)}\n${e.lines.join('\n')}`).join('\n\n') + '\n';
fs.writeFileSync(OUT, srt);
log.info('WRITE', {
	file: path.relative(root, OUT),
	entries: out.length,
	...stats,
	lastEndBefore: fmt(cues.at(-1).endMs),
	lastEndAfter: fmt(out.at(-1).endMs),
});
