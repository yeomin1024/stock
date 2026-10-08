// VERSION: v2.2.0 — 2026-10-08 — 대본(SRT/TXT)에서 엔론 5문장을 지워 자막 91개 → 잘라 내기 설정(CUT_SUBS) 삭제, 장면표 자막 번호 −5
// v2.1.0 — 2026-10-07 — 엔론 사례(자막 55–59, S23) 제외: 자막 60번부터 28.154초 앞당김, 장면 29개
// v2.0.0 — 2026-10-06 — 가이드 v2: 자막 96개, 고지 카드 6초(자막 18 뒤), 장면 30개
// 규칙: 프레임 = 초 × 30 (반올림). 장면 코드에 초를 하드코딩하지 않는다.
// 가이드 v2 "각 요소는 그 내용을 말하는 자막이 시작될 때 등장" → 요소 시점은 sub(n) (+ 4–6f 순차) 만 쓴다.
// (Node 스크립트에서도 그대로 import 하므로 .ts 확장자 import 와 순수 TS 문법만 쓴다.)
import {SUBTITLES} from './subtitles.ts';

export const FPS = 30;
export const WIDTH = 1920;
export const HEIGHT = 1080;

export const DISCLAIMER_AFTER_ID = 18;
export const DISCLAIMER_SEC = 6;
export const TAIL_SEC = 1;
export const EXPECTED_SUBTITLES = 91; // v2.2.0: 원본 96개 − 엔론 5개 (원본은 input/archive/)

export const secToFrame = (sec: number): number => Math.round(sec * FPS);

const subtitle = (id: number) => {
	const s = SUBTITLES[id - 1];
	if (!s || s.id !== id) {
		throw new Error(`[TIMELINE] 자막 ${id}번을 찾을 수 없음 (총 ${SUBTITLES.length}개)`);
	}
	return s;
};

const shiftSec = (id: number): number => (id > DISCLAIMER_AFTER_ID ? DISCLAIMER_SEC : 0);

export const subText = (id: number): string => subtitle(id).text;
/** 최종 타임라인(고지 카드 반영) 기준 자막 시작 프레임 */
export const subStart = (id: number): number => secToFrame(subtitle(id).startMs / 1000 + shiftSec(id));
/** 최종 타임라인(고지 카드 반영) 기준 자막 끝 프레임 */
export const subEnd = (id: number): number => secToFrame(subtitle(id).endMs / 1000 + shiftSec(id));

export const DISCLAIMER_START = subEnd(DISCLAIMER_AFTER_ID);
export const DISCLAIMER_FRAMES = secToFrame(DISCLAIMER_SEC);
const LAST = SUBTITLES[SUBTITLES.length - 1];
export const TOTAL_FRAMES = secToFrame(LAST.endMs / 1000 + shiftSec(LAST.id) + TAIL_SEC);

// ---------------------------------------------------------------------------
// 장면 표 (가이드 v2 5번). enter = 들어올 때 전환 (찢어진 종이 와이프 / 컷 번갈아).
// v2.1.0: S23(엔론) 삭제. 장면 번호는 가이드 표와 맞추려고 바꾸지 않는다(S22 다음이 S24).
// v2.2.0: 대본에서 엔론 5문장을 지워 S24 이후 자막 번호가 가이드 표보다 5 작다 (가이드 60 → 대본 55 … 가이드 96 → 대본 91).
// S22·S24 가 모두 와이프로 들어와 와이프가 두 번 이어진다 — S24 는 '문제 → 해결책' 파트 경계라 와이프를 유지.
// 앞 장면 그림을 이어 쓰는 곳(S03 계좌 카드, S04 색 빠짐, S09 사연 축소)은 컷.
// why = 가이드 "연결 근거" (스토리보드 index.html 에 그대로 표시)
// perSub = 한 장면 안에서 그림이 크게 바뀌어 자막 단위 still 을 더 만드는 장면 (가이드 6-5)
// ---------------------------------------------------------------------------
export type Enter = 'cut' | 'wipe';
export type Tone = 'cream' | 'navy' | 'bright';
export type SceneDef = {
	readonly id: string;
	readonly subs: readonly number[];
	readonly enter: Enter;
	readonly tone: Tone;
	readonly why: string;
	readonly perSub?: boolean;
	readonly shot?: string;
	/** 장면 끝이 페이드아웃이면 대표 still 을 이 자막 끝 10프레임 전으로 (S30) */
	readonly mainAtSub?: number;
};

export const SCENES: readonly SceneDef[] = [
	{id: 'S01', subs: [1, 2], enter: 'cut', tone: 'cream', why: '종목을 알게 된 경위, 고른 이유 3가지'},
	{id: 'S02', subs: [3], enter: 'cut', tone: 'cream', why: '전 재산을 한 종목에 넣음', shot: 'account'},
	{id: 'S03', subs: [4, 5], enter: 'cut', tone: 'cream', why: '수익, 더 오를 것 같은 기대', shot: 'account'},
	{id: 'S04', subs: [6], enter: 'cut', tone: 'cream', why: '반전 예고'},
	{id: 'S05', subs: [7], enter: 'cut', tone: 'navy', why: '계좌에 찍힌 손실'},
	{id: 'S06', subs: [8], enter: 'wipe', tone: 'navy', why: '손실의 원인이 된 소식'},
	{id: 'S07', subs: [9], enter: 'cut', tone: 'navy', why: '고점에서 내려오던 주가가 소식에 폭락'},
	{id: 'S08', subs: [10, 11], enter: 'wipe', tone: 'navy', why: '물타기 할 돈이 없음, 한탄'},
	{id: 'S09', subs: [12, 13, 14, 15], enter: 'cut', tone: 'cream', why: '몰빵하는 사람들의 생각 3가지'},
	{id: 'S10', subs: [16, 17, 18], enter: 'wipe', tone: 'cream', why: '이런 사람이 많다, 과연 수익을 낼까', perSub: true},
	{id: 'S11', subs: [], enter: 'cut', tone: 'cream', why: '투자 유의 사항'},
	{id: 'S12', subs: [19, 20], enter: 'wipe', tone: 'cream', why: '주식은 앞일을 모른다'},
	{id: 'S13', subs: [21, 22], enter: 'cut', tone: 'cream', why: '장기 상승해도 단기 악재가 온다, 누구나 손해'},
	{id: 'S14', subs: [23, 24, 25, 26, 27], enter: 'wipe', tone: 'cream', why: '빅테크 급락 사례', perSub: true},
	{id: 'S15', subs: [28, 29], enter: 'cut', tone: 'cream', why: '전문가도 예측 못 함의 근거'},
	{id: 'S16', subs: [30, 31, 32, 33], enter: 'wipe', tone: 'cream', why: '몰빵 상태에서는 선택지가 없음', perSub: true},
	{id: 'S17', subs: [34, 35, 36, 37, 38], enter: 'cut', tone: 'cream', why: '손실의 비대칭', perSub: true},
	{id: 'S18', subs: [39, 40, 41, 42, 43, 44], enter: 'wipe', tone: 'cream', why: '베셈바인더 연구', perSub: true},
	{id: 'S19', subs: [45, 46, 47], enter: 'cut', tone: 'cream', why: 'JP모건 분석', perSub: true},
	{id: 'S20', subs: [48], enter: 'wipe', tone: 'cream', why: '생존자 편향'},
	{id: 'S21', subs: [49, 50, 51], enter: 'cut', tone: 'cream', why: '손실 회피'},
	{id: 'S22', subs: [52, 53, 54], enter: 'wipe', tone: 'cream', why: '고통이 매일 찾아와 판단이 흐려짐', perSub: true},
	{id: 'S24', subs: [55, 56, 57], enter: 'wipe', tone: 'bright', why: '해결책 도입', perSub: true},
	{id: 'S25', subs: [58, 59, 60, 61, 62, 63], enter: 'cut', tone: 'cream', why: '비중 상한과 분산 계산', perSub: true},
	{id: 'S26', subs: [64, 65, 66, 67, 68, 69], enter: 'wipe', tone: 'cream', why: '같은 테마는 함께 무너짐', perSub: true},
	{id: 'S27', subs: [70, 71, 72, 73, 74], enter: 'cut', tone: 'cream', why: '지수 ETF로 쉽게 분산', perSub: true},
	{id: 'S28', subs: [75, 76, 77, 78, 79], enter: 'wipe', tone: 'cream', why: '분할 매수와 현금의 역할', perSub: true},
	{id: 'S29', subs: [80, 81, 82, 83, 84, 85], enter: 'cut', tone: 'cream', why: '수익 난 종목의 비중 관리', perSub: true},
	{id: 'S30', subs: [86, 87, 88, 89, 90, 91], enter: 'wipe', tone: 'cream', why: '마무리 당부', perSub: true, mainAtSub: 89},
];

/** 찢어진 종이 와이프: 경계 B 기준 [B - pre, B - pre + len]. 자막 시작 직후 끝나도록 앞당김 (새 장면 요소는 아직 안 나옴). */
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

export const nextEntersWithWipe = (id: string): boolean => {
	const next = SCENES[sceneIndex(id) + 1];
	return Boolean(next && next.enter === 'wipe');
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
	};
};

// S04 안에서 네이비로 찢어 넘기는 시점(로컬): 색 빠짐(0–30f) 뒤, 자막 6 길이의 45% 지점
export const S04_NAVY_WIPE_AT = (): number =>
	subStart(6) - sceneRange('S04').start + Math.round((subEnd(6) - subStart(6)) * 0.45);

/** 자막 띠 테마: 네이비 배경 구간이면 1 (S04 와이프 중간 ~ S09 시작) */
export const navyAmount = (frame: number): number => {
	const toNavy = sceneRange('S04').start + S04_NAVY_WIPE_AT() + WIPE.len / 2;
	const toCream = sceneRange('S09').start;
	const k = 4;
	if (frame <= toNavy - k || frame >= toCream) return 0;
	if (frame >= toNavy + k) return 1;
	return (frame - (toNavy - k)) / (2 * k);
};

// ---------------------------------------------------------------------------
// 스토리보드 still (가이드 6-5): 장면마다 "모든 요소가 다 나온 시점 = 장면 끝 10프레임 전".
// 다음 장면이 와이프로 들어오면 와이프가 덮기 전 프레임으로 당긴다.
// perSub 장면은 마지막 자막을 뺀 자막마다 "그 자막 끝 10프레임 전" still 을 더 만든다 (S25-61.png — 대본 자막 번호).
// ---------------------------------------------------------------------------
// 가이드 목록(S14, S16~S19, S25~S29) 외에 S10, S22, S24, S30 도 자막마다 그림이 바뀌어 자막 단위 still 을 더 만든다.
// S04 는 전환 장면이라 장면 끝이 빈 네이비 → 색이 빠진 카드와 찢어지는 종이가 함께 보이는 S04-wipe.png 를 더 만든다.
export type StoryShot = {readonly file: string; readonly scene: string; readonly frame: number; readonly subs: readonly number[]};

export const storyboardShots = (): StoryShot[] => {
	const out: StoryShot[] = [];
	for (const s of SCENES) {
		const {start, end} = sceneRange(s.id);
		const endFrame = end - 10 - (nextEntersWithWipe(s.id) ? WIPE.pre : 0);
		const mainIdx = s.mainAtSub ? s.subs.indexOf(s.mainAtSub) : s.subs.length - 1;
		const mainFrame = s.mainAtSub ? subEnd(s.mainAtSub) - 10 : endFrame;
		if (s.id === 'S04') {
			out.push({file: 'S04-wipe.png', scene: s.id, frame: start + S04_NAVY_WIPE_AT() + Math.round(WIPE.len * 0.45), subs: [6]});
		}
		if (s.perSub) {
			for (const n of s.subs.slice(0, mainIdx)) {
				out.push({file: `${s.id}-${n}.png`, scene: s.id, frame: subEnd(n) - 10, subs: [n]});
			}
			out.push({file: `${s.id}.png`, scene: s.id, frame: mainFrame, subs: s.subs.slice(mainIdx)});
		} else {
			out.push({file: `${s.id}.png`, scene: s.id, frame: mainFrame, subs: s.subs});
		}
	}
	return out;
};
