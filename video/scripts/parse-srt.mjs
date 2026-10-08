// VERSION: v2.2.0 — 2026-10-08 — SRT → src/data/subtitles.ts 생성 + 타임라인 검증/출력
// v2.2.0: 대본에서 엔론 5문장을 지워 기대 개수 91, 잘라 내기 출력 삭제 (v2.1.0 의 잘라 내기 설정은 없어짐)
// v2.0.0: 가이드 v2 — 자막 96개 / 고지 6초 / 약 9분 8초 확인 출력
// 사용: node scripts/parse-srt.mjs
// 1) input/대본_완성본.srt 를 @remotion/captions parseSrt 로 파싱
// 2) 무결성 검사(개수, 단조 증가, 겹침, 공백 구간) 후 로그
// 3) subtitles.ts 를 쓰고, timeline.ts 를 불러와 고지 카드를 반영한 전체 길이와 장면표를 출력
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {parseSrt} from '@remotion/captions';
import {log} from './log.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const SRT = path.join(root, 'input', '대본_완성본.srt');
const OUT = path.join(root, 'src', 'data', 'subtitles.ts');
const EXPECTED_COUNT = 91; // 가이드 v2 0번 표의 96개 − 엔론 5개 (v2.2.0 대본 수정)

let raw;
try {
	raw = fs.readFileSync(SRT, 'utf8');
} catch (err) {
	log.error('LOAD', {file: SRT, error: err.message, next: 'input/ 폴더에 대본_완성본.srt 를 넣고 다시 실행'});
	process.exit(1);
}

const {captions} = parseSrt({input: raw});
// parseSrt 는 SRT 번호를 돌려주지 않으므로, 원본 블록의 번호를 따로 읽어 순서를 검증한다.
const ids = raw
	.replace(/\r/g, '')
	.split(/\n\n+/)
	.map((b) => b.trim())
	.filter(Boolean)
	.map((b) => Number(b.split('\n')[0]));

log.info('LOAD', {file: path.relative(root, SRT), entries: captions.length, firstMs: captions[0].startMs, lastMs: captions.at(-1).endMs});

// ---- 무결성 검사 -----------------------------------------------------------
let problems = 0;
if (captions.length !== EXPECTED_COUNT) {
	log.warn('VALIDATE', {issue: 'count', expected: EXPECTED_COUNT, got: captions.length});
	problems++;
}
ids.forEach((id, i) => {
	if (id !== i + 1) {
		log.warn('VALIDATE', {issue: 'id_order', index: i, id});
		problems++;
	}
});
captions.forEach((c, i) => {
	if (!(c.endMs > c.startMs)) {
		log.warn('VALIDATE', {issue: 'non_positive_duration', id: i + 1, startMs: c.startMs, endMs: c.endMs});
		problems++;
	}
	const prev = captions[i - 1];
	if (prev && c.startMs < prev.endMs) {
		log.warn('VALIDATE', {issue: 'overlap', id: i + 1, prevEndMs: prev.endMs, startMs: c.startMs});
		problems++;
	}
	if (prev && c.startMs > prev.endMs) {
		log.info('VALIDATE', {note: 'gap', beforeId: i + 1, gapMs: c.startMs - prev.endMs});
	}
	if (c.text.split('\n').length > 2) {
		log.warn('VALIDATE', {issue: 'more_than_2_lines', id: i + 1});
	}
});
log.info('VALIDATE', {problems});

// ---- subtitles.ts 쓰기 ------------------------------------------------------
const body = captions
	.map(
		(c, i) =>
			`\t{id: ${i + 1}, startMs: ${c.startMs}, endMs: ${c.endMs}, text: ${JSON.stringify(c.text)}},`,
	)
	.join('\n');
const ts = `// 자동 생성 파일 — 직접 수정하지 말 것. 생성: node scripts/parse-srt.mjs
// 원본: input/대본_완성본.srt (자막 ${captions.length}개). 텍스트와 줄바꿈은 SRT 그대로.
export type Subtitle = {
	readonly id: number;
	readonly startMs: number;
	readonly endMs: number;
	readonly text: string;
};

export const SUBTITLES: readonly Subtitle[] = [
${body}
];
`;
fs.writeFileSync(OUT, ts);
log.info('WRITE', {file: path.relative(root, OUT), entries: captions.length});

// ---- 타임라인 출력 (timeline.ts 를 그대로 사용 — 계산 로직 중복 없음) --------
const tl = await import(path.join(root, 'src', 'data', 'timeline.ts'));
const sec = (f) => (f / tl.FPS).toFixed(2);
log.info('TIMELINE', {
	fps: tl.FPS,
	lastSubtitleEndFrame: tl.subEnd(captions.length),
	disclaimerStart: tl.DISCLAIMER_START,
	disclaimerFrames: tl.DISCLAIMER_FRAMES,
	totalFrames: tl.TOTAL_FRAMES,
	totalSec: sec(tl.TOTAL_FRAMES),
	totalHuman: `${Math.floor(tl.TOTAL_FRAMES / tl.FPS / 60)}분 ${(tl.TOTAL_FRAMES / tl.FPS % 60).toFixed(2)}초`,
});
for (const s of tl.SCENES) {
	const {start, end} = tl.sceneRange(s.id);
	log.info('SCENE', {id: s.id, subs: s.subs.length ? `${s.subs[0]}-${s.subs.at(-1)}` : 'notice', start, end, frames: end - start, startSec: sec(start), endSec: sec(end), enter: s.enter});
}
const totalSec = tl.TOTAL_FRAMES / tl.FPS;
const mmss = `${Math.floor(totalSec / 60)}분 ${(totalSec % 60).toFixed(2)}초`;
console.log(`\n자막 개수: ${captions.length}개 (기대 ${EXPECTED_COUNT}개) → ${captions.length === EXPECTED_COUNT ? 'OK' : '불일치'}`);
console.log(`고지 카드: 자막 ${tl.DISCLAIMER_AFTER_ID}번 끝 ${sec(tl.DISCLAIMER_START)} s 에 ${tl.DISCLAIMER_SEC} s 삽입, 자막 ${tl.DISCLAIMER_AFTER_ID + 1}번부터 ${tl.DISCLAIMER_SEC} s 뒤로`);
console.log(`전체 길이: ${tl.TOTAL_FRAMES} frames = ${sec(tl.TOTAL_FRAMES)} s = ${mmss} (SRT 마지막 자막 끝 ${(captions.at(-1).endMs / 1000).toFixed(3)} s + 고지 ${tl.DISCLAIMER_SEC} s + 여유 ${tl.TAIL_SEC} s) — 엔론 5문장을 지운 대본 기준 (원본 약 9분 8초 − 28.154초)`);
