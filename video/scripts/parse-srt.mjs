// VERSION: v2.4.0 — 2026-10-08 — SRT(한 줄 자막, 영상 시간) → src/data/subtitles.ts, 대본 TXT 문장 단위로 묶기 + 검증/출력
// v2.4.0: SRT 시간 = 영상 시간. 고지 카드 공백(문장 18 끝 ~ 19 시작)이 timeline.DISCLAIMER_SEC 와 같은지 검사
// v2.3.0: 자막을 한 줄씩 나누면서 SRT 번호와 장면 코드가 따로 놀지 않게, SRT 조각을 대본 TXT 의 문장(줄)에 묶는다.
//         장면 코드의 sub(n) 은 "대본 문장 n" (TXT 문장 줄 순서, 제목·파트 머리·[장면] 고지 블록 제외) 기준.
// v2.2.0: 대본에서 엔론 5문장을 지워 기대 개수 91, 잘라 내기 출력 삭제
// v2.0.0: 가이드 v2 — 자막 96개 / 고지 6초 / 약 9분 8초 확인 출력
// 사용: node scripts/parse-srt.mjs
// 1) input/대본_완성본.srt 를 @remotion/captions parseSrt 로 파싱, input/대본_완성본.txt 에서 문장 줄을 읽는다
// 2) 무결성 검사(번호 순서, 단조 증가, 겹침, 한 줄 규칙) + SRT 조각을 이어 붙인 글이 대본 문장과 같은지 검사
// 3) subtitles.ts(SUBTITLES: SRT 조각, SENTENCES: 문장 → 조각 범위)를 쓰고, timeline.ts 로 전체 길이와 장면표를 출력
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {parseSrt} from '@remotion/captions';
import {log} from './log.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const SRT = path.join(root, 'input', '대본_완성본.srt');
const TXT = path.join(root, 'input', '대본_완성본.txt');
const OUT = path.join(root, 'src', 'data', 'subtitles.ts');
const EXPECTED_SENTENCES = 91; // 가이드 v2 0번 표의 96개 − 엔론 5개. SRT 조각 수는 한 줄 나누기에 따라 달라진다

let raw;
let txt;
try {
	raw = fs.readFileSync(SRT, 'utf8');
	txt = fs.readFileSync(TXT, 'utf8');
} catch (err) {
	log.error('LOAD', {file: err.path ?? SRT, error: err.message, next: 'input/ 폴더에 대본_완성본.srt 와 대본_완성본.txt 를 넣고 다시 실행'});
	process.exit(1);
}

// 대본 문장 줄: 빈 줄로 나눈 블록 중 [장면] 고지 블록은 빼고, 제목 줄과 파트 머리(-로 시작)도 뺀다
const sentences = txt
	.replace(/\r/g, '')
	.split(/\n\s*\n/)
	.filter((b) => !b.trim().startsWith('[장면]'))
	.flatMap((b) => b.split('\n'))
	.map((l) => l.trim())
	.filter((l) => l && !l.startsWith('-') && !l.startsWith('제목:'));

const {captions} = parseSrt({input: raw});
// parseSrt 는 SRT 번호를 돌려주지 않으므로, 원본 블록의 번호를 따로 읽어 순서를 검증한다.
const ids = raw
	.replace(/\r/g, '')
	.split(/\n\n+/)
	.map((b) => b.trim())
	.filter(Boolean)
	.map((b) => Number(b.split('\n')[0]));

log.info('LOAD', {file: path.relative(root, SRT), entries: captions.length, firstMs: captions[0].startMs, lastMs: captions.at(-1).endMs, sentences: sentences.length});

// ---- 무결성 검사 -----------------------------------------------------------
let problems = 0;
if (sentences.length !== EXPECTED_SENTENCES) {
	log.warn('VALIDATE', {issue: 'sentence_count', expected: EXPECTED_SENTENCES, got: sentences.length});
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
		log.debug('VALIDATE', {note: 'gap', beforeId: i + 1, gapMs: c.startMs - prev.endMs});
	}
});

// ---- SRT 조각 → 대본 문장 묶기 ------------------------------------------------
// 조각 글(줄바꿈은 공백으로)을 차례로 이어 붙여 문장과 정확히 같아질 때까지 묶는다.
const flat = (t) => t.replace(/\n/g, ' ').trim();
const groups = [];
let k = 0;
for (let n = 0; n < sentences.length; n++) {
	const first = k;
	let acc = '';
	while (k < captions.length && acc.length < sentences[n].length) {
		acc = acc ? `${acc} ${flat(captions[k].text)}` : flat(captions[k].text);
		k++;
	}
	if (acc !== sentences[n]) {
		log.error('VALIDATE', {issue: 'srt_txt_mismatch', sentence: n + 1, txt: sentences[n], srt: acc, next: 'SRT 와 TXT 의 문장이 같은지 확인'});
		process.exit(1);
	}
	groups.push({n: n + 1, first: first + 1, last: k});
}
if (k !== captions.length) {
	log.error('VALIDATE', {issue: 'srt_leftover', usedEntries: k, entries: captions.length});
	process.exit(1);
}

// 한 줄 규칙 (v2.3.0 사용자 요청): 두 줄 자막은 "앞의 사연과 같이" 다음 혼잣말(다음 문장이 "이런 식으로"로 시작하기 전까지)만 허용
const anchor = sentences.indexOf('앞의 사연과 같이');
const monologue = new Set();
for (let n = anchor + 1; anchor >= 0 && n < sentences.length && !sentences[n].startsWith('이런 식으로'); n++) monologue.add(n + 1);
const lineCount = {1: 0, 2: 0};
for (const g of groups) {
	for (let e = g.first; e <= g.last; e++) {
		const lines = captions[e - 1].text.split('\n').length;
		lineCount[lines] = (lineCount[lines] ?? 0) + 1;
		if (lines > 1 && !monologue.has(g.n)) {
			log.warn('VALIDATE', {issue: 'multi_line_subtitle', entry: e, sentence: g.n});
			problems++;
		}
	}
}
log.info('VALIDATE', {problems, sentences: groups.length, entries: captions.length, oneLine: lineCount[1], twoLine: lineCount[2] ?? 0, monologueSentences: [...monologue].join(',')});

// ---- subtitles.ts 쓰기 ------------------------------------------------------
const sentenceOf = new Map(groups.flatMap((g) => Array.from({length: g.last - g.first + 1}, (_, i) => [g.first + i, g.n])));
const body = captions
	.map((c, i) => `\t{id: ${i + 1}, sentence: ${sentenceOf.get(i + 1)}, startMs: ${c.startMs}, endMs: ${c.endMs}, text: ${JSON.stringify(c.text)}},`)
	.join('\n');
const sentBody = groups.map((g, i) => `\t{n: ${g.n}, first: ${g.first}, last: ${g.last}, text: ${JSON.stringify(sentences[i])}},`).join('\n');
const ts = `// 자동 생성 파일 — 직접 수정하지 말 것. 생성: node scripts/parse-srt.mjs
// 원본: input/대본_완성본.srt (자막 ${captions.length}개) + input/대본_완성본.txt (문장 ${groups.length}개).
// SUBTITLES = 화면에 나오는 SRT 자막(텍스트·줄바꿈 그대로), SENTENCES = 대본 문장 → SRT 자막 범위(first–last).
export type Subtitle = {
	readonly id: number;
	readonly sentence: number;
	readonly startMs: number;
	readonly endMs: number;
	readonly text: string;
};
export type Sentence = {
	readonly n: number;
	readonly first: number;
	readonly last: number;
	readonly text: string;
};

export const SUBTITLES: readonly Subtitle[] = [
${body}
];

export const SENTENCES: readonly Sentence[] = [
${sentBody}
];
`;
fs.writeFileSync(OUT, ts);
log.info('WRITE', {file: path.relative(root, OUT), entries: captions.length, sentences: groups.length});

// ---- 타임라인 출력 (timeline.ts 를 그대로 사용 — 계산 로직 중복 없음) --------
const tl = await import(path.join(root, 'src', 'data', 'timeline.ts'));
const sec = (f) => (f / tl.FPS).toFixed(2);
log.info('TIMELINE', {
	fps: tl.FPS,
	lastSubtitleEndFrame: tl.subEnd(groups.length),
	disclaimerStart: tl.DISCLAIMER_START,
	disclaimerFrames: tl.DISCLAIMER_FRAMES,
	totalFrames: tl.TOTAL_FRAMES,
	totalSec: sec(tl.TOTAL_FRAMES),
	totalHuman: `${Math.floor(tl.TOTAL_FRAMES / tl.FPS / 60)}분 ${((tl.TOTAL_FRAMES / tl.FPS) % 60).toFixed(2)}초`,
});
for (const s of tl.SCENES) {
	const {start, end} = tl.sceneRange(s.id);
	log.info('SCENE', {id: s.id, sentences: s.subs.length ? `${s.subs[0]}-${s.subs.at(-1)}` : 'notice', start, end, frames: end - start, startSec: sec(start), endSec: sec(end), enter: s.enter});
}
const totalSec = tl.TOTAL_FRAMES / tl.FPS;
const mmss = `${Math.floor(totalSec / 60)}분 ${(totalSec % 60).toFixed(2)}초`;
console.log(`\n대본 문장: ${groups.length}개 (기대 ${EXPECTED_SENTENCES}개) → ${groups.length === EXPECTED_SENTENCES ? 'OK' : '불일치'} / SRT 자막: ${captions.length}개 (한 줄 ${lineCount[1]}, 두 줄 ${lineCount[2] ?? 0} — 혼잣말 예외)`);
// 고지 카드 공백 검사: SRT 의 문장 18 끝 ~ 문장 19 시작 = DISCLAIMER_SEC (영상 시간 = SRT 시간)
const g18 = groups[tl.DISCLAIMER_AFTER_ID - 1];
const g19 = groups[tl.DISCLAIMER_AFTER_ID];
const gapMs = captions[g19.first - 1].startMs - captions[g18.last - 1].endMs;
const gapOk = Math.abs(gapMs - tl.DISCLAIMER_SEC * 1000) <= 1;
(gapOk ? log.info : log.warn)('VALIDATE', {check: 'notice_gap', afterSentence: tl.DISCLAIMER_AFTER_ID, srtEntries: `${g18.last}->${g19.first}`, gapMs, expectedMs: tl.DISCLAIMER_SEC * 1000, frames: tl.DISCLAIMER_FRAMES, ok: gapOk});
console.log(`고지 카드: SRT ${g18.last}번 끝 ${(captions[g18.last - 1].endMs / 1000).toFixed(3)} s ~ ${g19.first}번 시작 ${(captions[g19.first - 1].startMs / 1000).toFixed(3)} s 공백 ${(gapMs / 1000).toFixed(3)} s (기대 ${tl.DISCLAIMER_SEC} s) → ${gapOk ? 'OK' : '불일치'}`);
console.log(`전체 길이: ${tl.TOTAL_FRAMES} frames = ${sec(tl.TOTAL_FRAMES)} s = ${mmss} (SRT 마지막 자막 끝 ${(captions.at(-1).endMs / 1000).toFixed(3)} s + 여유 ${tl.TAIL_SEC} s, SRT 시간 = 영상 시간)`);
