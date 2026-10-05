// VERSION: v1.0.0 — 2026-10-05 — S10 (자막 13–15) 생각 말풍선 3개가 각 자막 시작에 맞춰 하나씩 등장
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {sceneTimes} from '../data/timeline';
import {C} from '../design/colors';
import {SANS} from '../design/fonts';
import {enterP, prog} from '../design/motion';
import {DrawPath, Svg} from '../components/Draw';
import {cloudPath, personPath} from '../components/Icons';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';

const t = sceneTimes('S10');

export type BubbleDef = {
	readonly cx: number;
	readonly cy: number;
	readonly rx: number;
	readonly ry: number;
	readonly text: string;
	readonly seed: number;
	readonly at: number;
	/** 머리 쪽으로 이어지는 작은 원 [x, y, r] */
	readonly dots: readonly (readonly [number, number, number])[];
};

export const PERSON = {cx: 960, cy: 652, r: 46};

export const BUBBLES: readonly BubbleDef[] = [
	{cx: 470, cy: 420, rx: 320, ry: 112, text: '몰빵하면 많이 오르겠지?', seed: 1, at: t.sub(13), dots: [[870, 600, 11], [790, 566, 17]]},
	{cx: 960, cy: 210, rx: 290, ry: 100, text: '나누면 수익이 줄어', seed: 2, at: t.sub(14), dots: [[960, 560, 11], [960, 500, 17]]},
	{cx: 1450, cy: 420, rx: 320, ry: 112, text: '떨어져도 금방 회복', seed: 3, at: t.sub(15), dots: [[1050, 600, 11], [1130, 566, 17]]},
];

/** 말풍선 하나. p = 등장(0→1), scale/translate/textOpacity 는 S11 합치기에서 사용 */
export const ThoughtBubble: React.FC<{
	readonly b: BubbleDef;
	readonly p: number;
	readonly dotsP: number;
	readonly scale?: number;
	readonly dx?: number;
	readonly dy?: number;
	readonly textOpacity?: number;
	readonly opacity?: number;
}> = ({b, p, dotsP, scale = 1, dx = 0, dy = 0, textOpacity = 1, opacity = 1}) => {
	if (p <= 0.001 && dotsP <= 0.001) return null;
	const k = (0.7 + 0.3 * p) * scale;
	return (
		<>
			<Svg>
				<g opacity={opacity}>
					{b.dots.map(([x, y, r], i) => (
						<circle key={i} cx={x} cy={y} r={r * Math.min(1, Math.max(0, dotsP * 2 - i))} fill={C.white} stroke={C.ink} strokeWidth={4} />
					))}
					<g
						opacity={p}
						style={{transformBox: 'view-box', transformOrigin: `${b.cx}px ${b.cy}px`, scale: String(k), translate: `${dx}px ${dy}px`}}
					>
						<path d={cloudPath(b.cx, b.cy, b.rx, b.ry, 11, b.seed)} fill={C.white} stroke={C.ink} strokeWidth={5} strokeLinejoin="round" />
					</g>
				</g>
			</Svg>
			<div
				style={{
					position: 'absolute',
					left: b.cx - b.rx,
					top: b.cy - 40,
					width: b.rx * 2,
					height: 80,
					display: 'flex',
					alignItems: 'center',
					justifyContent: 'center',
					fontFamily: SANS,
					fontWeight: 700,
					fontSize: 44,
					color: C.ink,
					whiteSpace: 'nowrap',
					opacity: p * textOpacity * opacity,
					scale: String(k),
					translate: `${dx}px ${dy}px`,
				}}
			>
				{b.text}
			</div>
		</>
	);
};

export const Person: React.FC<{readonly p: number; readonly opacity?: number}> = ({p, opacity = 1}) => {
	const {head, body} = personPath(PERSON.cx, PERSON.cy, PERSON.r);
	return (
		<Svg>
			<g opacity={opacity}>
				<DrawPath d={head} p={p} width={6} />
				<DrawPath d={body} p={p} width={6} />
			</g>
		</Svg>
	);
};

export const S10: React.FC = () => {
	const f = useSceneFrame();
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer depth="bg">
				<Person p={prog(f, 0, 18)} />
			</Layer>
			<Layer depth="mid">
				{BUBBLES.map((b) => (
					<ThoughtBubble key={b.seed} b={b} dotsP={prog(f, b.at, b.at + 8)} p={enterP(f, b.at + 4, 15)} />
				))}
			</Layer>
		</AbsoluteFill>
	);
};
