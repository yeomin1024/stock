// VERSION: v2.0.0 — 2026-10-06 — 가이드 v2 3-4 모션 상수/헬퍼 (v2: 카메라 1.00→1.03)
// 등장 12–18f (spring ease-out, 바운스 없음), 퇴장 8–10f, 순차 간격 4–6f,
// 형광펜 10f, 카운터 20–30f, 카메라 1.00→1.03.
import {Easing, interpolate, spring} from 'remotion';
import {FPS} from '../data/timeline';

export const M = {
	enter: 15,
	exit: 9,
	stagger: 5,
	highlight: 10,
	counter: 26,
	draw: 18,
	cameraZoom: 0.03,
} as const;

const clampOpts = {extrapolateLeft: 'clamp', extrapolateRight: 'clamp'} as const;

/** 등장 진행도 0→1 (spring, damping 200 = 과한 바운스 없음) */
export const enterP = (frame: number, at: number, dur: number = M.enter): number =>
	spring({frame: frame - at, fps: FPS, config: {damping: 200}, durationInFrames: dur});

/** 퇴장 진행도 1→0 (ease-in) */
export const exitP = (frame: number, at: number, dur: number = M.exit): number =>
	interpolate(frame, [at, at + dur], [1, 0], {...clampOpts, easing: Easing.bezier(0.55, 0, 0.75, 0.3)});

/** 등장·퇴장을 함께 고려한 가시도 (퇴장 시점이 없으면 등장만) */
export const visP = (frame: number, at: number, exitAt?: number, dur: number = M.enter, exitDur: number = M.exit): number =>
	Math.min(enterP(frame, at, dur), exitAt === undefined ? 1 : exitP(frame, exitAt, exitDur));

/** 일반 구간 진행도 (clamp + ease) */
export const prog = (
	frame: number,
	a: number,
	b: number,
	easing: (t: number) => number = Easing.bezier(0.33, 1, 0.68, 1),
): number => interpolate(frame, [a, b], [0, 1], {...clampOpts, easing});

/** 선형 진행도 (clamp) */
export const lin = (frame: number, a: number, b: number): number => interpolate(frame, [a, b], [0, 1], clampOpts);

export const easeInOut = Easing.bezier(0.65, 0, 0.35, 1);
export const easeOut = Easing.bezier(0.16, 1, 0.3, 1);
export const easeIn = Easing.bezier(0.55, 0, 1, 0.45);

/** a→b 선형 보간 */
export const lerp = (a: number, b: number, t: number): number => a + (b - a) * t;
