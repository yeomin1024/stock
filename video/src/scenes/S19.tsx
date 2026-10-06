// VERSION: v2.0.0 — 2026-10-06 — S19 (자막 45–47) "JP모건 · 1980년 이후" 카드
// 46: 10칸 중 4칸이 파랑 + 각 칸에 고점에서 -70% 떨어진 뒤 회복 못 하는 미니 선 → "40% 이상 / 고점 대비 -70% · 회복 못 함"
// 47: 칸 3개 중 2개에 "시장보다 못함" → "3번 중 2번". 연결 근거: JP모건 분석
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {FACTS, SRC, formatPct} from '../data/facts';
import {sceneTimes} from '../data/timeline';
import {C, mix} from '../design/colors';
import {SANS, SERIF} from '../design/fonts';
import {enterP, exitP, lin, prog} from '../design/motion';
import {T} from '../design/type';
import {ConceptTag, SourceCaption, cardStyle} from '../components/Bits';
import {Counter} from '../components/Counter';
import {DrawPath, Svg} from '../components/Draw';
import {Highlight} from '../components/Highlight';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {roundRect, smoothPath, Pt} from '../components/hand';

const t = sceneTimes('S19');
const J = FACTS.jpm;
const S45 = t.sub(45);
const S46 = t.sub(46);
const S47 = t.sub(47);

const N = 10;
const X0 = 166;
const CW = 148;
const GAP = 12;
const CY = 300;
const CH = 210;
const LOSERS = [1, 3, 6, 8]; // 10칸 중 4칸 = 40%

// 미니 선 모양 (값 0~1, 개념도)
const normalShape = (i: number): number[] =>
	Array.from({length: 9}, (_, k) => {
		const s = k / 8;
		return 0.22 + 0.55 * s + 0.12 * Math.sin(s * 9 + i * 1.3) - 0.1 * Math.exp(-((s - 0.55) ** 2) / 0.01);
	});
const loserShape = [0.25, 0.48, 0.72, 1, 0.62, 0.3, 0.27, 0.31, 0.28]; // 고점 1 → 0.3 (-70%) 뒤 제자리
const toPts = (vals: number[], cx: number): Pt[] => vals.map((v, k) => [cx + 16 + (k / (vals.length - 1)) * (CW - 32), CY + CH - 24 - v * (CH - 56)] as Pt);

export const S19: React.FC = () => {
	const f = useSceneFrame();
	const out46 = exitP(f, S47 - 2, 9);
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer>
				{/* 자막 45: 카드 */}
				<Reveal at={S45} from="left" style={{left: 96, top: 128}}>
					<div style={{...cardStyle(), position: 'relative', padding: '18px 36px 20px', display: 'flex', alignItems: 'baseline', gap: 22}}>
						<span style={{fontFamily: SERIF, fontWeight: 900, fontSize: 72, color: C.ink}}>JP모건</span>
						<span style={{...T.label, fontSize: 44}}>· {J.sinceYear}년 이후</span>
					</div>
				</Reveal>

				{/* 자막 46: 10칸 + 미니 선 */}
				{f >= S46 && out46 > 0.001 ? (
					<Svg>
						<g opacity={out46}>
							{Array.from({length: N}, (_, i) => {
								const cx = X0 + i * (CW + GAP);
								const li = LOSERS.indexOf(i);
								const hit = li >= 0 ? lin(f, S46 + 10 + li * 6, S46 + 18 + li * 6) : 0;
								const appear = enterP(f, S46 + i * 2, 12);
								const pts = toPts(li >= 0 ? loserShape : normalShape(i), cx);
								return (
									<g key={i} opacity={appear}>
										<rect x={cx} y={CY} width={CW} height={CH} rx={14} fill={mix(C.paper, C.blue, 0.14 * hit)} />
										<path d={roundRect(cx, CY, CW, CH, 14)} fill="none" stroke={mix(C.ink, C.blue, hit)} strokeWidth={3 + 2 * hit} />
										<DrawPath d={smoothPath(pts)} p={prog(f, S46 + 4 + i * 2, S46 + 28 + i * 2)} width={4 + 2 * hit} stroke={li >= 0 ? mix(C.gray, C.blue, hit) : C.gray} />
										{li >= 0 ? (
											<DrawPath d={`M ${pts[3][0]} ${pts[3][1]} L ${cx + CW - 12} ${pts[3][1]}`} p={prog(f, S46 + 24 + li * 4, S46 + 34 + li * 4)} dash="6 6" width={3} stroke={C.blue} linecap="butt" />
										) : null}
									</g>
								);
							})}
						</g>
					</Svg>
				) : null}
				<Reveal at={S46 + 10} exitAt={S47 - 2} from="up" style={{left: X0 - 6, top: 560}}>
					<div style={{display: 'flex', alignItems: 'baseline', gap: 18}}>
						<Counter from={0} to={J.noRecoverPct} at={S46 + 10} format={(v) => formatPct(v, 0, false)} color={C.blue} style={{fontSize: 160}} />
						<span style={{fontFamily: SANS, fontWeight: 900, fontSize: 56, color: C.blue}}>이상</span>
					</div>
				</Reveal>
				<Reveal at={S46 + 30} exitAt={S47 - 2} from="right" style={{left: 900, top: 620}}>
					<div style={{...T.label, fontSize: 44}}>
						고점 대비 <span style={{color: C.blue, fontWeight: 900}}>{formatPct(J.drawdownPct)}</span> · 회복 못 함
					</div>
				</Reveal>
				<ConceptTag at={S46 + 6} exitAt={S47 - 2} x={1666} y={CY + CH + 12} />

				{/* 자막 47: 칸 3개 중 2개 "시장보다 못함" + 3번 중 2번 */}
				{[0, 1, 2].map((i) => {
					const p = enterP(f, S47 + 6 + i * 5, 14);
					const bad = i < 2;
					const mark = bad ? lin(f, S47 + 22 + i * 6, S47 + 30 + i * 6) : 0;
					const cx = 520 + i * 440;
					return p > 0.001 ? (
						<div
							key={i}
							style={{
								...cardStyle(),
								left: cx - 190,
								top: 310,
								width: 380,
								height: 230,
								opacity: p,
								background: mix(C.paper, C.blue, mark),
								borderColor: mix(C.ink, C.blue, mark),
								display: 'flex',
								alignItems: 'center',
								justifyContent: 'center',
							}}
						>
							<span style={{...T.label, fontSize: 40, color: C.white, opacity: mark}}>시장보다 못함</span>
						</div>
					) : null;
				})}
				<Reveal at={S47 + 34} from="up" style={{left: 0, right: 0, top: 580, textAlign: 'center'}}>
					<div style={{...T.number, fontSize: 160, color: C.ink}}>
						<Highlight at={S47 + 42} top={0.5}>
							{J.worseThanMarket}
						</Highlight>
					</div>
				</Reveal>
			</Layer>
			<SourceCaption text={SRC.jpm} at={S46 + 10} x={1824} align="right" />
		</AbsoluteFill>
	);
};
