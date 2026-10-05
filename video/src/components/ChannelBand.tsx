// VERSION: v1.0.0 — 2026-10-05 — 찢어진 노랑 종이 띠 + 큰 세리프 문구 (S09 채널명, S13 고지, S24 엔딩 공용)
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {C} from '../design/colors';
import {SERIF} from '../design/fonts';
import {enterP, exitP, lin} from '../design/motion';
import {useSceneFrame} from './Scene';
import {TornBand} from './TornWipe';

export const BandMessage: React.FC<{
	readonly at: number;
	readonly text: string;
	readonly seed: string;
	readonly exitAt?: number;
	readonly fontSize?: number;
	readonly y?: number;
	readonly h?: number;
}> = ({at, text, seed, exitAt, fontSize = 96, y = 330, h = 210}) => {
	const f = useSceneFrame();
	const band = lin(f, at, at + 16);
	const txt = enterP(f, at + 8, 15);
	const out = exitAt === undefined ? 1 : exitP(f, exitAt, 9);
	if (band <= 0 || out <= 0) return null;
	return (
		<AbsoluteFill style={{opacity: out}}>
			<TornBand p={band} y={y} h={h} seed={seed} color={C.yellow} />
			<div
				style={{
					position: 'absolute',
					left: 0,
					right: 0,
					top: y,
					height: h,
					display: 'flex',
					alignItems: 'center',
					justifyContent: 'center',
					rotate: '-1.2deg',
					fontFamily: SERIF,
					fontWeight: 900,
					fontSize,
					color: C.ink,
					whiteSpace: 'nowrap',
					opacity: txt,
					scale: String(0.92 + 0.08 * txt),
				}}
			>
				{text}
			</div>
		</AbsoluteFill>
	);
};
