// VERSION: v2.2.0 — 2026-10-08 — 자막 번호 −5(대본에서 엔론 5문장 삭제) · S24 (자막 55–57) 화면이 밝아지며 "올바른 방법"
// 56: 큰 블록 하나가 여러 작은 블록으로 나뉨 "욕심 대신 나눠 담기" → 57: 1~5번 빈 목록 "5가지". 연결 근거: 해결책 도입
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {sceneTimes} from '../data/timeline';
import {C} from '../design/colors';
import {SANS, SERIF} from '../design/fonts';
import {easeInOut, enterP, exitP, lerp, prog} from '../design/motion';
import {T} from '../design/type';
import {DrawPath, Svg} from '../components/Draw';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {handLine} from '../components/hand';

const t = sceneTimes('S24');
const [S55, S56, S57] = [55, 56, 57].map((n) => t.sub(n));
const BIG = {x: 610, y: 300, w: 700, h: 170};
const PIECES = 6;
const PW = 180;
const PGAP = 40;
const PX0 = (1920 - (PIECES * PW + (PIECES - 1) * PGAP)) / 2;

export const S24: React.FC = () => {
	const f = useSceneFrame();
	const bigIn = enterP(f, S56, 14);
	const split = easeInOut(prog(f, S56 + 16, S56 + 40, (x) => x));
	const out56 = exitP(f, S57 - 2, 9);
	return (
		<AbsoluteFill>
			<SceneBg tone="bright" from="cream" at={S55} dur={24} />
			<Layer>
				<Reveal at={S55 + 6} from="up" style={{left: 0, right: 0, top: 128, textAlign: 'center'}}>
					<div style={{fontFamily: SERIF, fontWeight: 900, fontSize: 88, color: C.ink, lineHeight: 1.2}}>올바른 방법</div>
				</Reveal>
				<Svg>
					<DrawPath d={handLine(770, 250, 1150, 246, 's24ul', 2)} p={prog(f, S55 + 16, S55 + 28)} width={8} />
					{/* 61: 큰 블록 → 작은 블록 6개 */}
					<g opacity={Math.min(bigIn, out56)}>
						{Array.from({length: PIECES}, (_, i) => {
							const x = lerp(BIG.x + (i * BIG.w) / PIECES, PX0 + i * (PW + PGAP), split);
							const w = lerp(BIG.w / PIECES, PW, split);
							const h = lerp(BIG.h, 130, split);
							const y = lerp(BIG.y, 330, split);
							return <rect key={i} x={x} y={y} width={w} height={h} rx={lerp(0, 14, split)} fill={C.paper} stroke={C.ink} strokeWidth={4} />;
						})}
						{split < 0.02 ? <rect x={BIG.x} y={BIG.y} width={BIG.w} height={BIG.h} rx={16} fill={C.paper} stroke={C.ink} strokeWidth={5} /> : null}
					</g>
				</Svg>
				<Reveal at={S56 + 34} exitAt={S57 - 2} from="up" style={{left: 0, right: 0, top: 520, textAlign: 'center'}}>
					<div style={{...T.label, fontSize: 44}}>욕심 대신 나눠 담기</div>
				</Reveal>
				{/* 62: 1~5 빈 목록 + 5가지 */}
				{[1, 2, 3, 4, 5].map((n, i) => {
					const p = enterP(f, S57 + 4 + i * 4, 12);
					const y = 300 + i * 92;
					return p > 0.001 ? (
						<div key={n} style={{position: 'absolute', left: 820, top: y, opacity: p, display: 'flex', alignItems: 'center', gap: 26}}>
							<div style={{width: 64, height: 64, borderRadius: 32, border: `4px solid ${C.ink}`, boxSizing: 'border-box', display: 'flex', alignItems: 'center', justifyContent: 'center', fontFamily: SANS, fontWeight: 900, fontSize: 36, color: C.ink}}>
								{n}
							</div>
							<div style={{width: 520, height: 0, borderBottom: `4px dashed ${C.gray}`}} />
						</div>
					) : null;
				})}
				<Reveal at={S57 + 6} from="left" style={{left: 380, top: 470}}>
					<div style={{...T.headline, fontSize: 84}}>5가지</div>
				</Reveal>
			</Layer>
		</AbsoluteFill>
	);
};
