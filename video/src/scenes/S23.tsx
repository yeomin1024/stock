// VERSION: v2.0.0 — 2026-10-06 — S23 (자막 55–59) "엔론 · 2001 · 미국 에너지 기업 · 파산" 카드
// 56: 퇴직연금 막대의 62%가 자사주 색(잉크 = 한 종목)으로 → 57: 그 부분에 "노후 자금 절반 넘게 = 한 종목"
// 58: 하락하는 주가선(개념도) 위 80달러·70달러 지점에 "팔 수 있었다" → 59: "손실 10억 달러+" (파랑)
// 연결 근거: 확신 때문에 팔지 못한 사례
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {FACTS, SRC} from '../data/facts';
import {sceneTimes} from '../data/timeline';
import {C} from '../design/colors';
import {SERIF} from '../design/fonts';
import {enterP, exitP, prog} from '../design/motion';
import {T} from '../design/type';
import {ConceptTag, SourceCaption, cardStyle} from '../components/Bits';
import {DrawPath, Svg} from '../components/Draw';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {smoothPath, Pt} from '../components/hand';

const t = sceneTimes('S23');
const E = FACTS.enron;
const [S55, S56, S57, S58, S59] = [55, 56, 57, 58, 59].map((n) => t.sub(n));

const BAR = {x: 260, y: 560, w: 1400, h: 120};
const LINE_A: Pt[] = [[300, 360], [420, 384], [540, 410], [650, 432], [760, 454], [870, 478], [980, 540]];
const LINE_B: Pt[] = [[980, 540], [1100, 610], [1220, 668], [1340, 712], [1460, 744], [1580, 760]];
const P80 = LINE_A[3];
const P70 = LINE_A[5];

export const S23: React.FC = () => {
	const f = useSceneFrame();
	const barOut = exitP(f, S58 - 2, 9);
	const fill = prog(f, S56 + 10, S56 + 36);
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer>
				{/* 55: 카드 */}
				<Reveal at={S55} from="left" style={{left: 96, top: 124}}>
					<div style={{...cardStyle(), position: 'relative', padding: '16px 36px 20px'}}>
						<div style={{fontFamily: SERIF, fontWeight: 900, fontSize: 72, color: C.ink, lineHeight: 1.1}}>엔론</div>
						<div style={{...T.label, fontSize: 36, marginTop: 6}}>
							{E.year} · {E.desc} · <span style={{color: C.blue}}>파산</span>
						</div>
					</div>
				</Reveal>

				{/* 56–57: 퇴직연금 막대 */}
				{f >= S56 && barOut > 0.001 ? (
					<div style={{position: 'absolute', inset: 0, opacity: barOut}}>
						<Reveal at={S56} from="up" style={{left: BAR.x, top: 330}}>
							<div style={{...T.label, color: C.gray}}>퇴직연금의</div>
							<div style={{display: 'flex', alignItems: 'baseline', gap: 22, marginTop: 4}}>
								<span style={{...T.number, fontSize: 170, color: C.ink}}>{E.ownStockPct}%</span>
								<span style={{...T.label, fontSize: 44}}>= 자사주</span>
							</div>
						</Reveal>
						<div style={{position: 'absolute', left: BAR.x, top: BAR.y, width: BAR.w, height: BAR.h, border: `4px solid ${C.ink}`, borderRadius: 12, boxSizing: 'border-box', background: C.paper, overflow: 'hidden', opacity: enterP(f, S56 + 4, 12)}}>
							<div style={{position: 'absolute', left: 0, top: 0, bottom: 0, width: `${(E.ownStockPct * fill)}%`, background: C.ink}} />
							<div
								style={{
									position: 'absolute',
									left: 0,
									width: `${E.ownStockPct}%`,
									top: 0,
									bottom: 0,
									display: 'flex',
									alignItems: 'center',
									justifyContent: 'center',
									...T.label,
									color: C.light,
									opacity: enterP(f, S57, 12),
								}}
							>
								노후 자금 절반 넘게 = 한 종목
							</div>
						</div>
					</div>
				) : null}

				{/* 58–59: 하락하는 주가선 (개념도) */}
				{f >= S58 ? (
					<>
						<Svg>
							<DrawPath d={smoothPath(LINE_A)} p={prog(f, S58 + 2, S58 + 30)} width={7} stroke={C.ink} />
							<DrawPath d={smoothPath(LINE_B)} p={prog(f, S59, S59 + 30)} width={8} stroke={C.blue} />
							{[P80, P70].map(([x, y], i) => (
								<circle key={i} cx={x} cy={y} r={13 * enterP(f, S58 + 18 + i * 6, 10)} fill={C.ink} stroke={C.paper} strokeWidth={4} />
							))}
						</Svg>
						<Reveal at={S58 + 18} from="down" dist={14} style={{left: P80[0] - 90, width: 180, top: P80[1] + 24, textAlign: 'center'}}>
							<div style={{...T.label}}>{E.prices[0]}</div>
						</Reveal>
						<Reveal at={S58 + 24} from="down" dist={14} style={{left: P70[0] - 90, width: 180, top: P70[1] + 24, textAlign: 'center'}}>
							<div style={{...T.label}}>{E.prices[1]}</div>
						</Reveal>
						<Reveal at={S58 + 32} from="up" dist={14} style={{left: 560, width: 420, top: 300, textAlign: 'center'}}>
							<div style={{display: 'inline-block', ...T.label, fontSize: 38, border: `3px solid ${C.ink}`, borderRadius: 30, padding: '2px 22px 4px', background: C.paper}}>팔 수 있었다</div>
						</Reveal>
						<ConceptTag at={S58 + 6} x={1690} y={772} />
					</>
				) : null}
				<Reveal at={S59 + 20} from="right" style={{left: 900, top: 150}}>
					<div style={{display: 'flex', alignItems: 'baseline', gap: 20}}>
						<span style={{...T.label, fontSize: 44, color: C.blue}}>손실</span>
						<span style={{...T.number, fontSize: 160, color: C.blue}}>{E.lossLabel}</span>
					</div>
				</Reveal>
			</Layer>
			<SourceCaption text={SRC.press} at={S55 + 6} exitAt={S56 - 2} />
			<SourceCaption text={SRC.enronCrs} at={S56 + 6} exitAt={S58 - 2} />
			<SourceCaption text={SRC.cnn} at={S58 + 6} exitAt={S59 + 18} />
			<SourceCaption text={SRC.wharton} at={S59 + 20} />
		</AbsoluteFill>
	);
};
