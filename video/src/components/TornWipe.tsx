// VERSION: v2.0.0 — 2026-10-06 — 찢어진 종이 와이프 / 찢어진 띠 (가벼운 폴리곤 clip-path)
// v2: 가우시안 블러 그림자 제거(속도 규칙). 가장자리 = 흰 종이 섬유 + 얇은 반투명 그림자 띠(필터 없음).
import React from 'react';
import {AbsoluteFill, random} from 'remotion';
import {C} from '../design/colors';
import {easeInOut, lerp} from '../design/motion';
import {Svg} from './Draw';
import {Pt, ptsToPath, ptsToPolygon, tornEdge, tornEdgeH} from './hand';

const SLANT = 170;

export const TornWipe: React.FC<{
	/** 0→1 진행도 */
	readonly p: number;
	readonly seed: string;
	readonly children: React.ReactNode;
}> = ({p, seed, children}) => {
	if (p >= 1) return <>{children}</>;
	if (p <= 0) return null;
	const x0 = lerp(-SLANT - 60, 1920 + 40, easeInOut(p));
	const edge = tornEdge(seed, x0, -40, 1120, {slant: SLANT, jag: 6, wave: 14});
	const fiber: Pt[] = edge.map(([x, y], i) => [x + 4 + random(`${seed}-f${i}`) * 9, y]);
	const shade: Pt[] = fiber.map(([x, y]) => [x + 10, y]);
	return (
		<>
			<AbsoluteFill style={{clipPath: `polygon(-60px -40px, ${ptsToPolygon(edge)}, -60px 1120px)`}}>{children}</AbsoluteFill>
			<AbsoluteFill style={{pointerEvents: 'none'}}>
				<Svg>
					<path d={ptsToPath([...fiber, ...shade.slice().reverse()])} fill="rgba(0,0,0,0.12)" />
					<path d={ptsToPath([...edge, ...fiber.slice().reverse()])} fill={C.fiber} />
				</Svg>
			</AbsoluteFill>
		</>
	);
};

export const TornBand: React.FC<{
	/** 펼쳐짐 0→1 */
	readonly p: number;
	readonly y: number;
	readonly h: number;
	readonly seed: string;
	readonly color?: string;
	readonly x0?: number;
	readonly x1?: number;
	readonly rotate?: number;
	readonly children?: React.ReactNode;
}> = ({p, y, h, seed, color = C.yellow, x0 = -40, x1 = 1960, rotate = -1, children}) => {
	if (p <= 0) return null;
	const top = tornEdgeH(`${seed}-t`, x0, x1, y, {jag: 4, wave: 7, step: 16});
	const bot = tornEdgeH(`${seed}-b`, x0, x1, y + h, {jag: 4, wave: 7, step: 16});
	const poly = [...top, ...bot.slice().reverse()];
	const reveal = lerp(x0, x1, easeInOut(Math.min(1, p)));
	return (
		<AbsoluteFill style={{rotate: `${rotate}deg`, transformOrigin: `960px ${y + h / 2}px`}}>
			<AbsoluteFill style={{clipPath: `inset(-200px ${1920 - reveal}px -200px -200px)`}}>
				<Svg>
					<path d={ptsToPath(poly.map(([a, b]) => [a + 5, b + 7] as Pt))} fill="rgba(30,30,30,0.10)" />
					<path d={ptsToPath(poly)} fill={color} />
				</Svg>
				{children}
			</AbsoluteFill>
		</AbsoluteFill>
	);
};
