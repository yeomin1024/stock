// VERSION: v2.4.0 — 2026-10-08 — 고지 카드 3.5초(사용자 요청)에 맞춰 등장을 앞당김: 띠 0–12f, 문구 8/12/16f 시작·12f → 28f 에 모두 보임
// v2.0.0 — 2026-10-06 — S11 (고지 카드, 문장 18 직후) 노랑 띠 카드에 대본 [장면] 고지 문구 3문장이 순서대로
// 연결 근거: 투자 유의 사항
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {NOTICE} from '../data/facts';
import {C} from '../design/colors';
import {SANS} from '../design/fonts';
import {enterP, lin} from '../design/motion';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {TornBand} from '../components/TornWipe';

export const S11: React.FC = () => {
	const f = useSceneFrame();
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer>
				<TornBand p={lin(f, 0, 12)} y={220} h={540} seed="S11-notice" rotate={-0.6} />
				<div
					style={{
						position: 'absolute',
						left: 260,
						right: 260,
						top: 220,
						height: 540,
						display: 'flex',
						flexDirection: 'column',
						justifyContent: 'center',
						gap: 34,
						rotate: '-0.6deg',
					}}
				>
					{NOTICE.map((line, i) => {
						const p = enterP(f, 8 + i * 4, 12);
						return (
							<div
								key={i}
								style={{
									fontFamily: SANS,
									fontWeight: 700,
									fontSize: 42,
									lineHeight: 1.5,
									color: C.ink,
									textAlign: 'center',
									wordBreak: 'keep-all',
									opacity: p,
									translate: `0px ${(1 - p) * 20}px`,
								}}
							>
								{line}
							</div>
						);
					})}
				</div>
			</Layer>
		</AbsoluteFill>
	);
};
