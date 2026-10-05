// VERSION: v1.0.0 — 2026-10-05 — 금액/% 카운트업·카운트다운 (20–30f, tabular-nums 로 흔들림 없음)
import React from 'react';
import {Easing, interpolate} from 'remotion';
import {M} from '../design/motion';
import {T} from '../design/type';
import {useSceneFrame} from './Scene';

/** from→to 값 (at 에서 시작, dur 프레임 동안). steps 가 있으면 그 단위로 끊어서 보여준다. */
export const useCount = (from: number, to: number, at: number, dur: number = M.counter, steps?: number): number => {
	const f = useSceneFrame();
	const v = interpolate(f, [at, at + dur], [from, to], {
		extrapolateLeft: 'clamp',
		extrapolateRight: 'clamp',
		easing: Easing.bezier(0.25, 0.8, 0.35, 1),
	});
	if (!steps) return v;
	const n = Math.round((v - from) / steps) * steps + from;
	return n;
};

export const Counter: React.FC<{
	readonly from: number;
	readonly to: number;
	readonly at: number;
	readonly dur?: number;
	readonly steps?: number;
	readonly format: (v: number) => string;
	readonly color?: string | ((v: number) => string);
	readonly style?: React.CSSProperties;
}> = ({from, to, at, dur = M.counter, steps, format, color, style}) => {
	const v = useCount(from, to, at, dur, steps);
	const c = typeof color === 'function' ? color(v) : color;
	return <span style={{...T.number, color: c, ...style}}>{format(v)}</span>;
};
