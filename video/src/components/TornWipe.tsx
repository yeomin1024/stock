// VERSION: v1.0.0 — 2026-10-05 — 찢어진 종이 와이프 / 찢어진 노랑 띠
// TornWipe: 들어오는 장면을 불규칙한 세로 가장자리로 잘라 왼쪽→오른쪽으로 덮는다.
//           가장자리에는 흰 종이 섬유 + 이전 장면 위로 떨어지는 그림자.
// TornBand: 위/아래가 찢어진 가로 종이 띠가 왼쪽→오른쪽으로 펼쳐진다.
import React from 'react';
import {AbsoluteFill, random} from 'remotion';
import {C} from '../design/colors';
import {easeInOut, lerp} from '../design/motion';
import {Svg, useSafeId} from './Draw';
import {Pt, ptsToPath, ptsToPolygon, tornEdge, tornEdgeH} from './hand';

const SLANT = 170;

export const TornWipe: React.FC<{
	/** 0→1 진행도 */
	readonly p: number;
	readonly seed: string;
	readonly children: React.ReactNode;
}> = ({p, seed, children}) => {
	const blur = useSafeId('wipe-shadow');
	if (p >= 1) return <>{children}</>;
	if (p <= 0) return null;
	const x0 = lerp(-SLANT - 60, 1920 + 40, easeInOut(p));
	const edge = tornEdge(seed, x0, -40, 1120, {slant: SLANT, jag: 6, wave: 14});
	const fiber: Pt[] = edge.map(([x, y], i) => [x + 4 + random(`${seed}-f${i}`) * 9, y]);
	const fiberPoly = [...edge, ...fiber.slice().reverse()];
	const shadow = fiber.map(([x, y]) => [x + 8, y] as Pt);
	return (
		<>
			<AbsoluteFill style={{clipPath: `polygon(-60px -40px, ${ptsToPolygon(edge)}, -60px 1120px)`}}>
				{children}
			</AbsoluteFill>
			<AbsoluteFill style={{pointerEvents: 'none'}}>
				<Svg>
					<defs>
						<filter id={blur} x="-50%" y="-10%" width="200%" height="120%">
							<feGaussianBlur stdDeviation={7} />
						</filter>
					</defs>
					<path d={ptsToPath(shadow, false)} stroke="rgba(0,0,0,0.28)" strokeWidth={18} fill="none" filter={`url(#${blur})`} />
					<path d={ptsToPath(fiberPoly)} fill={C.fiber} />
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
}> = ({p, y, h, seed, color = C.yellow, x0 = -40, x1 = 1960, rotate = -1.2, children}) => {
	const shadowId = useSafeId('band-shadow');
	if (p <= 0) return null;
	const top = tornEdgeH(`${seed}-t`, x0, x1, y, {jag: 4, wave: 7, step: 16});
	const bot = tornEdgeH(`${seed}-b`, x0, x1, y + h, {jag: 4, wave: 7, step: 16});
	const poly = [...top, ...bot.slice().reverse()];
	const reveal = lerp(x0, x1, easeInOut(Math.min(1, p)));
	return (
		<AbsoluteFill style={{rotate: `${rotate}deg`, transformOrigin: `960px ${y + h / 2}px`}}>
			<AbsoluteFill style={{clipPath: `inset(-200px ${1920 - reveal}px -200px -200px)`}}>
				<Svg>
					<defs>
						<filter id={shadowId} x="-10%" y="-30%" width="120%" height="160%">
							<feGaussianBlur stdDeviation={6} />
						</filter>
					</defs>
					<path d={ptsToPath(poly.map(([a, b]) => [a + 6, b + 10] as Pt))} fill="rgba(30,30,30,0.16)" filter={`url(#${shadowId})`} />
					<path d={ptsToPath(poly)} fill={color} />
				</Svg>
				{children}
			</AbsoluteFill>
		</AbsoluteFill>
	);
};
