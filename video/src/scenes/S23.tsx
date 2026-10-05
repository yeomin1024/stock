// VERSION: v1.0.0 — 2026-10-05 — S23 (자막 39) 번개가 블록 하나에 떨어져도 나머지 9개는 그대로, 선으로 그린 방패가 감싼다 → "대처할 여유"
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {FACTS} from '../data/facts';
import {sceneTimes} from '../data/timeline';
import {C, alpha} from '../design/colors';
import {easeInOut, enterP, lin, prog} from '../design/motion';
import {T} from '../design/type';
import {DrawPath, Svg} from '../components/Draw';
import {Highlight} from '../components/Highlight';
import {Lightning, shieldPath} from '../components/Icons';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';

const t = sceneTimes('S23');
const N = FACTS.diversify.stocks;
const LOSS_FRAC = Math.abs(FACTS.story.lossPct) / 100;
const BOLT_AT = t.word(39, '악재') - 6;
const SHIELD_AT = t.word(39, '대처할') - 4;
const HIT = 2;

const BS = 120;
const GAP = 16;
const GX = 700 - (5 * BS + 4 * GAP) / 2;
const GY = 340;
const cell = (i: number) => ({x: GX + (i % 5) * (BS + GAP), y: GY + Math.floor(i / 5) * (BS + GAP)});

export const S23: React.FC = () => {
	const f = useSceneFrame();
	const hit = cell(HIT);
	const strike = lin(f, BOLT_AT, BOLT_AT + 4);
	const flash = lin(f, BOLT_AT + 3, BOLT_AT + 5) * (1 - lin(f, BOLT_AT + 6, BOLT_AT + 14));
	const shrink = easeInOut(prog(f, BOLT_AT + 4, BOLT_AT + 18, (x) => x));
	const boltOut = 1 - lin(f, BOLT_AT + 16, BOLT_AT + 26);
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer depth="mid">
				<Svg>
					{Array.from({length: N}, (_, i) => {
						const {x, y} = cell(i);
						const appear = enterP(f, 4 + i * 2, 14);
						if (i !== HIT) {
							return <rect key={i} x={x} y={y} width={BS} height={BS} rx={12} fill={C.paper} stroke={C.ink} strokeWidth={4} opacity={appear} />;
						}
						const lossH = BS * LOSS_FRAC * shrink;
						return (
							<g key={i} opacity={appear}>
								{shrink > 0 ? (
									<rect x={x} y={y} width={BS} height={lossH} rx={10} fill={alpha(C.blue, 0.12)} stroke={C.blue} strokeWidth={4} strokeDasharray="9 7" />
								) : null}
								<rect x={x} y={y + lossH} width={BS} height={BS - lossH} rx={12} fill={shrink > 0 ? alpha(C.blue, 0.2) : C.paper} stroke={shrink > 0 ? C.blue : C.ink} strokeWidth={4} />
							</g>
						);
					})}
					{/* 번개 */}
					<g opacity={boltOut}>
						<Lightning x={hit.x + BS / 2 + 10} y={GY - 250} h={240} reveal={strike} rotate={4} />
					</g>
					<circle cx={hit.x + BS / 2} cy={hit.y} r={90} fill={C.yellow} opacity={flash * 0.5} />
					{/* 방패 */}
					<DrawPath d={shieldPath(700, 180, 940, 600)} p={prog(f, SHIELD_AT, SHIELD_AT + 26)} width={9} stroke={C.ink} />
				</Svg>
			</Layer>
			<Layer depth="fg">
				<Reveal at={SHIELD_AT + 6} from="right" style={{left: 1250, top: 360}}>
					<div style={{...T.headline, fontSize: 96}}>
						<Highlight at={SHIELD_AT + 16}>대처할 여유</Highlight>
					</div>
				</Reveal>
			</Layer>
		</AbsoluteFill>
	);
};
