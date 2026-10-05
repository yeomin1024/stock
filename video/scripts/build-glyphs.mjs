// VERSION: v1.0.0 — 2026-10-05 — 사용 글자만 담은 Google Fonts 청크 목록 생성
// 사용: node scripts/build-glyphs.mjs
// Noto Sans/Serif KR 는 unicode-range 로 쪼개진 청크가 굵기마다 124개다. 전부 받으면 렌더 탭마다
// 수백 개를 내려받으므로, 자막·화면 텍스트에 실제로 쓰인 글자를 포함한 청크만 고른다.
// 대상: input/대본_완성본.srt + src/**/*.ts(x) 의 문자열 (주석 제외).
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {createRequire} from 'node:module';
import {log} from './log.mjs';

const require = createRequire(import.meta.url);
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const OUT = path.join(root, 'src', 'design', 'font-subsets.ts');

const walk = (dir) =>
	fs.readdirSync(dir, {withFileTypes: true}).flatMap((e) => {
		const p = path.join(dir, e.name);
		if (e.isDirectory()) return walk(p);
		return /\.(ts|tsx)$/.test(e.name) && p !== OUT ? [p] : [];
	});

const stripComments = (src) =>
	src.replace(/\/\*[\s\S]*?\*\//g, '').replace(/(^|[^:'"`])\/\/.*$/gm, '$1');

const files = [path.join(root, 'input', '대본_완성본.srt'), ...walk(path.join(root, 'src'))];
const chars = new Set();
for (const f of files) {
	const text = f.endsWith('.srt') ? fs.readFileSync(f, 'utf8') : stripComments(fs.readFileSync(f, 'utf8'));
	for (const ch of text) {
		if (ch.codePointAt(0) > 0x20) chars.add(ch);
	}
}
log.info('GLYPHS', {files: files.length, uniqueChars: chars.size, hangul: [...chars].filter((c) => /[가-힣]/.test(c)).length});

const parseRanges = (s) =>
	s.split(',').map((r) => {
		const [a, b] = r.trim().replace(/^U\+/i, '').split('-');
		const lo = parseInt(a, 16);
		return [lo, b ? parseInt(b, 16) : lo];
	});

const pick = (modName) => {
	const info = require(`@remotion/google-fonts/${modName}`).getInfo();
	const ranges = Object.entries(info.unicodeRanges).map(([k, v]) => [k, parseRanges(v)]);
	const chosen = new Set();
	const uncovered = [];
	for (const ch of chars) {
		const cp = ch.codePointAt(0);
		const hit = ranges.find(([, rs]) => rs.some(([lo, hi]) => cp >= lo && cp <= hi));
		if (hit) chosen.add(hit[0]);
		else uncovered.push(ch);
	}
	if (uncovered.length) {
		log.warn('GLYPHS', {font: modName, uncovered: uncovered.join(''), note: '이 글자는 대체 폰트로 그려짐'});
	}
	const list = [...chosen].sort();
	log.info('GLYPHS', {font: modName, subsets: list.length, of: ranges.length});
	return list;
};

const sans = pick('NotoSansKR');
const serif = pick('NotoSerifKR');

fs.writeFileSync(
	OUT,
	`// 자동 생성 파일 — 직접 수정하지 말 것. 생성: node scripts/build-glyphs.mjs
// 화면/자막에 쓰인 글자를 포함한 Google Fonts 청크만 로드한다.
export const SANS_SUBSETS: string[] = ${JSON.stringify(sans)};
export const SERIF_SUBSETS: string[] = ${JSON.stringify(serif)};
`,
);
log.info('WRITE', {file: path.relative(root, OUT), sans: sans.length, serif: serif.length});
