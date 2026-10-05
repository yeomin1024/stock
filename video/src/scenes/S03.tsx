// VERSION: v1.0.0 — 2026-10-05 — S03 (자막 4–5) 평가손익 0 → +1,000만 원(빨강), 상승 라인(개념도), 전고점 점선 + 위쪽 점선 화살표
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {FACTS, formatManwon} from '../data/facts';
import {sceneTimes} from '../data/timeline';
import {C, pnlColor} from '../design/colors';
import {SANS} from '../design/fonts';
import {enterP, prog} from '../design/motion';
import {T} from '../design/type';
import {ConceptTag, cardStyle} from '../components/Bits';
import {Counter} from '../components/Counter';
import {DrawPath, Svg} from '../components/Draw';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {arrowHead, handLine, smoothPath, Pt} from '../components/hand';

const t = sceneTimes('S03');
const GAIN_AT = t.word(4, '천만');
const PEAK_AT = t.sub(5);

// 차트 (개념도: 눈금 없음). 원래 960–1680 폭으로 잡은 모양을 1120–1760 으로 옮긴다.
const AX = 1100;
const AY0 = 130;
const AY1 = 740;
const AX1 = 1800;
const sx = (x: number) => 1120 + (x - 960) * 0.889;
const BEFORE_RAW: Pt[] = [
	[960, 470],
	[1010, 400],
	[1050, 330],
	[1090, 262],
	[1130, 300],
	[1175, 380],
	[1215, 360],
	[1260, 470],
	[1305, 560],
	[1345, 540],
	[1390, 640],
];
const AFTER_RAW: Pt[] = [
	[1390, 640],
	[1440, 600],
	[1485, 615],
	[1530, 540],
	[1575, 555],
	[1625, 470],
	[1680, 410],
];
const BEFORE: Pt[] = BEFORE_RAW.map(([x, y]) => [sx(x), y] as Pt);
const AFTER: Pt[] = AFTER_RAW.map(([x, y]) => [sx(x), y] as Pt);
const PEAK_Y = 262;
const END = AFTER[AFTER.length - 1];

export const S03Content: React.FC = () => {
	const f = useSceneFrame();
	const cardIn = enterP(f, 0, 15);
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer depth="bg">
				<Svg>
					<DrawPath d={handLine(AX, AY0, AX, AY1, 's03ay', 1.2)} p={prog(f, 0, 14)} width={4} stroke={C.ink} />
					<DrawPath d={handLine(AX, AY1, AX1, AY1, 's03ax', 1.2)} p={prog(f, 4, 20)} width={4} stroke={C.ink} />
				</Svg>
			</Layer>
			<Layer depth="mid">
				<Svg>
					<DrawPath d={smoothPath(BEFORE)} p={prog(f, 2, Math.min(20, GAIN_AT - 2))} stroke={C.gray} width={7} />
					<DrawPath d={smoothPath(AFTER)} p={prog(f, GAIN_AT - 2, GAIN_AT + 26)} stroke={C.red} width={8} />
					{f >= GAIN_AT + 24 ? <circle cx={END[0]} cy={END[1]} r={12} fill={C.red} opacity={enterP(f, GAIN_AT + 24, 8)} /> : null}
					{/* 전고점: 이전 고점에서 오른쪽으로 점선 */}
					<DrawPath d={`M ${BEFORE[3][0]} ${PEAK_Y} L ${AX1 - 20} ${PEAK_Y}`} p={prog(f, PEAK_AT, PEAK_AT + 18)} dash="16 14" width={4} stroke={C.ink} linecap="butt" />
					{/* 위로 향하는 점선 화살표 */}
					<DrawPath d={`M ${END[0]} ${END[1] - 26} L ${END[0]} ${PEAK_Y + 22}`} p={prog(f, PEAK_AT + 12, PEAK_AT + 28)} dash="12 10" width={6} stroke={C.red} linecap="butt" />
					<DrawPath d={arrowHead([END[0], PEAK_Y + 14], -Math.PI / 2, 26)} p={prog(f, PEAK_AT + 26, PEAK_AT + 32)} width={6} stroke={C.red} />
				</Svg>
				<Reveal at={PEAK_AT + 4} from="down" dist={20} style={{left: AX1 - 170, top: PEAK_Y - 66}}>
					<div style={{...T.label, fontSize: 42}}>전고점</div>
				</Reveal>
				<ConceptTag at={10} x={AX1 - 100} y={AY1 + 14} />
			</Layer>
			<Layer depth="fg">
				<div style={{...cardStyle(), left: 90, top: 170, width: 960, height: 350, opacity: cardIn, translate: `${(1 - cardIn) * -40}px 0px`, padding: '36px 48px', boxSizing: 'border-box'}}>
					<div style={{display: 'flex', justifyContent: 'space-between', alignItems: 'baseline'}}>
						<div style={{...T.label, fontSize: 42}}>평가손익</div>
						<div style={{fontFamily: SANS, fontWeight: 700, fontSize: 34, color: C.gray}}>{FACTS.mdb.ticker}</div>
					</div>
					<div style={{marginTop: 46}}>
						<Counter
							from={0}
							to={FACTS.story.gainManwon}
							at={GAIN_AT}
							steps={50}
							format={(v) => formatManwon(v, true)}
							color={(v) => (v === 0 ? C.ink : pnlColor(v))}
							style={{fontSize: 160}}
						/>
					</div>
				</div>
			</Layer>
		</AbsoluteFill>
	);
};

export const S03: React.FC = () => <S03Content />;
