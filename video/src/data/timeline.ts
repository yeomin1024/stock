// VERSION: v1.0.0 — 2026-10-05 — 타임라인: 모든 장면/자막 타이밍을 SRT 자막 번호로 계산
// 규칙: 프레임 = 초 × 30 (반올림). 초를 장면 코드에 직접 하드코딩하지 않는다.
// 자막 19번 직후 고지 카드 3초를 끼우고, 자막 20번 이후는 전부 3초 뒤로 민다.
// (이 파일은 Node 스크립트에서도 그대로 import 하므로 .ts 확장자 import 와 순수 TS 문법만 쓴다.)
import {SUBTITLES} from './subtitles.ts';

export const FPS = 30;
export const WIDTH = 1920;
export const HEIGHT = 1080;

export const DISCLAIMER_AFTER_ID = 19;
export const DISCLAIMER_SEC = 3;
export const TAIL_SEC = 1;

export const secToFrame = (sec: number): number => Math.round(sec * FPS);

const shiftSec = (id: number): number => (id > DISCLAIMER_AFTER_ID ? DISCLAIMER_SEC : 0);

const subtitle = (id: number) => {
	const s = SUBTITLES[id - 1];
	if (!s || s.id !== id) {
		throw new Error(`[TIMELINE] 자막 ${id}번을 찾을 수 없음 (총 ${SUBTITLES.length}개)`);
	}
	return s;
};

/** 최종 타임라인(고지 카드 반영) 기준 자막 시작 프레임 */
export const subStart = (id: number): number => secToFrame(subtitle(id).startMs / 1000 + shiftSec(id));
/** 최종 타임라인(고지 카드 반영) 기준 자막 끝 프레임 */
export const subEnd = (id: number): number => secToFrame(subtitle(id).endMs / 1000 + shiftSec(id));

export const DISCLAIMER_START = subEnd(DISCLAIMER_AFTER_ID);
export const DISCLAIMER_FRAMES = secToFrame(DISCLAIMER_SEC);
export const TOTAL_FRAMES = secToFrame(
	SUBTITLES[SUBTITLES.length - 1].endMs / 1000 + DISCLAIMER_SEC + TAIL_SEC,
);

/**
 * 읽기 속도 계수. 무음 영상은 시청자가 자막을 "읽으므로" 발화보다 빠르게(≈1.8배) 내용을 따라간다.
 * 그래서 단어 위치 비율에 이 계수를 곱해, 자막 앞쪽 55% 안에서 순서대로 드러나게 한다.
 * (내레이션을 넣는 경우 1.0 으로 바꾸면 발화 속도 기준이 된다.)
 */
export const READ_PACE = 0.55;

/**
 * 자막 안에서 특정 단어가 읽히는 시점 추정(프레임) = 자막 시작 + 글자 위치 비율 × READ_PACE × 자막 길이.
 * 단어가 없으면 즉시 예외 — 대본이 바뀌었는데 연출이 그대로인 상황을 조용히 넘기지 않기 위함.
 */
export const wordAt = (id: number, word: string): number => {
	const text = subtitle(id).text.replace(/\n/g, ' ');
	const idx = text.indexOf(word);
	if (idx < 0) {
		throw new Error(`[TIMELINE] 자막 ${id}번에 "${word}" 없음: ${text}`);
	}
	const a = subStart(id);
	const b = subEnd(id);
	return Math.round(a + (idx / text.length) * READ_PACE * (b - a));
};

// ---------------------------------------------------------------------------
// 장면 표 (가이드 5번). enter = 이 장면으로 들어올 때의 전환 방식.
// 찢어진 종이 와이프와 컷을 번갈아 쓰되, 앞 장면과 그림이 이어지는 곳(S04, S11, S18)은 컷.
// shot 이 같은 장면들은 카메라 줌(1.00→1.04)을 이어서 쓴다.
// ---------------------------------------------------------------------------
export type Enter = 'cut' | 'wipe';
export type Tone = 'cream' | 'navy';
export type SceneDef = {
	readonly id: string;
	readonly subs: readonly number[];
	readonly enter: Enter;
	readonly tone: Tone;
	readonly shot?: string;
};

export const SCENES: readonly SceneDef[] = [
	{id: 'S01', subs: [1, 2], enter: 'cut', tone: 'cream'},
	{id: 'S02', subs: [3], enter: 'cut', tone: 'cream'},
	{id: 'S03', subs: [4, 5], enter: 'wipe', tone: 'cream'},
	{id: 'S04', subs: [6], enter: 'cut', tone: 'cream'},
	{id: 'S05', subs: [7], enter: 'cut', tone: 'navy'},
	{id: 'S06', subs: [8], enter: 'wipe', tone: 'navy'},
	{id: 'S07', subs: [9], enter: 'cut', tone: 'navy'},
	{id: 'S08', subs: [10, 11], enter: 'wipe', tone: 'navy'},
	{id: 'S09', subs: [12], enter: 'wipe', tone: 'cream'},
	{id: 'S10', subs: [13, 14, 15], enter: 'cut', tone: 'cream', shot: 'thoughts'},
	{id: 'S11', subs: [16], enter: 'cut', tone: 'cream', shot: 'thoughts'},
	{id: 'S12', subs: [17, 18, 19], enter: 'wipe', tone: 'cream'},
	{id: 'S13', subs: [], enter: 'cut', tone: 'cream'},
	{id: 'S14', subs: [20, 21], enter: 'wipe', tone: 'cream'},
	{id: 'S15', subs: [22, 23], enter: 'cut', tone: 'cream'},
	{id: 'S16', subs: [24, 25, 26], enter: 'wipe', tone: 'cream'},
	{id: 'S17', subs: [27, 28, 29], enter: 'cut', tone: 'cream', shot: 'bessembinder'},
	{id: 'S18', subs: [30], enter: 'cut', tone: 'cream', shot: 'bessembinder'},
	{id: 'S19', subs: [31, 32], enter: 'wipe', tone: 'cream'},
	{id: 'S20', subs: [33], enter: 'cut', tone: 'cream'},
	{id: 'S21', subs: [34, 35], enter: 'wipe', tone: 'cream'},
	{id: 'S22', subs: [36, 37, 38], enter: 'cut', tone: 'cream'},
	{id: 'S23', subs: [39], enter: 'wipe', tone: 'cream'},
	{id: 'S24', subs: [40, 41, 42, 43], enter: 'cut', tone: 'cream'},
];

/** 찢어진 종이 와이프: 경계 B 기준 [B - pre, B - pre + len] 동안 진행. 자막 시작 직후 끝나도록 앞당김. */
export const WIPE = {pre: 14, len: 18} as const;

const sceneIndex = (id: string): number => {
	const i = SCENES.findIndex((s) => s.id === id);
	if (i < 0) throw new Error(`[TIMELINE] 장면 ${id} 없음`);
	return i;
};

const rawStart = (i: number): number => {
	if (i === 0) return 0;
	const s = SCENES[i];
	return s.subs.length ? subStart(s.subs[0]) : DISCLAIMER_START;
};

export const sceneRange = (id: string): {start: number; end: number} => {
	const i = sceneIndex(id);
	const start = rawStart(i);
	const end = i + 1 < SCENES.length ? rawStart(i + 1) : TOTAL_FRAMES;
	return {start, end};
};

/** 카메라 줌 구간: 같은 shot 끼리 묶은 범위 */
export const shotRange = (id: string): {start: number; end: number} => {
	const s = SCENES[sceneIndex(id)];
	if (!s.shot) return sceneRange(id);
	const members = SCENES.filter((x) => x.shot === s.shot).map((x) => sceneRange(x.id));
	return {start: Math.min(...members.map((m) => m.start)), end: Math.max(...members.map((m) => m.end))};
};

/** 장면 안에서 쓰는 로컬 타이밍 (장면 시작 = 0) */
export const sceneTimes = (id: string) => {
	const {start, end} = sceneRange(id);
	return {
		start,
		end,
		dur: end - start,
		/** 자막 n 시작 (로컬 프레임) */
		sub: (n: number) => subStart(n) - start,
		/** 자막 n 끝 (로컬 프레임) */
		subEnd: (n: number) => subEnd(n) - start,
		/** 자막 n 안에서 단어가 나오는 추정 시점 (로컬 프레임) */
		word: (n: number, w: string) => wordAt(n, w) - start,
	};
};

// S04 내부에서 네이비로 찢어 넘기는 시점(로컬): 채도 빠짐(0–36f)이 끝난 뒤, 자막 6 길이의 45% 지점.
export const S04_NAVY_WIPE_AT = (): number =>
	subStart(6) - sceneRange('S04').start + Math.round((subEnd(6) - subStart(6)) * 0.45);

/** 자막 띠 테마: 네이비 배경 구간이면 1, 아니면 0 (와이프 중간에서 전환) */
export const navyAmount = (frame: number): number => {
	const toNavy = sceneRange('S04').start + S04_NAVY_WIPE_AT() + WIPE.len / 2;
	const toCream = sceneRange('S09').start - WIPE.pre + WIPE.len / 2;
	const k = 4;
	if (frame <= toNavy - k || frame >= toCream + k) return 0;
	if (frame >= toNavy + k && frame <= toCream - k) return 1;
	if (frame < toNavy + k) return (frame - (toNavy - k)) / (2 * k);
	return 1 - (frame - (toCream - k)) / (2 * k);
};
