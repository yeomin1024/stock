// VERSION: v2.0.0 — 2026-10-06 — S03 (자막 4–5) 같은 계좌 카드의 평가손익 0 → +1,000만 원(빨강) + 작은 상승선(개념도)
// 자막 5: 위쪽 "전고점" 점선 + 위로 향하는 점선 화살표. 연결 근거: 수익, 더 오를 것 같은 기대
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {FACTS} from '../data/facts';
import {sceneTimes} from '../data/timeline';
import {C, mix} from '../design/colors';
import {enterP, exitP, prog} from '../design/motion';
import {T} from '../design/type';
import {AccountCard} from '../components/Account';
import {ConceptTag} from '../components/Bits';
import {useCount} from '../components/Counter';
import {DrawPath, Svg} from '../components/Draw';
import {Highlight} from '../components/Highlight';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {arrowHead, handLine, smoothPath, Pt} from '../components/hand';
import {CARD_POS} from './S02';

const t = sceneTimes('S03');
const S4 = t.sub(4);
const S5 = t.sub(5);
const MODE_AT = S4 + 4;
const GAIN_AT = S4 + 14;

// 작은 상승선 (개념도: 눈금 없음)
const AX = 1410;
const AY0 = 230;
const AY1 = 690;
const AX1 = 1820;
const PEAK: Pt = [1462, 300];
const HISTORY: Pt[] = [[1428, 372], PEAK, [1500, 372], [1540, 430], [1580, 528], [1622, 600]];
const RISE: Pt[] = [[1622, 600], [1662, 566], [1700, 522], [1742, 470], [1790, 420]];
const END = RISE[RISE.length - 1];

/** S04 에서 색이 빠질 때(desat) 재사용하므로 분리 */
export const S03Content: React.FC<{readonly desat?: number}> = ({desat = 0}) => {
	const f = useSceneFrame();
	const red = mix(C.red, C.gray, desat);
	const pnl = useCount(0, FACTS.story.gainManwon, GAIN_AT, 26, 50);
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer>
				<AccountCard x={CARD_POS.x} y={CARD_POS.y} pnlMode={prog(f, MODE_AT, MODE_AT + 12)} totalManwon={FACTS.story.principalManwon} pnlManwon={pnl} bar={1} desat={desat} />
				{/* S02 에서 이어진 "몰빵" 은 퇴장 */}
				{exitP(f, S4, 9) > 0.001 ? (
					<div style={{position: 'absolute', left: 1440, top: 380, opacity: exitP(f, S4, 9)}}>
						<div style={{...T.headline, fontSize: 88}}>
							<Highlight at={-100}>몰빵</Highlight>
						</div>
					</div>
				) : null}
				<Svg>
					<DrawPath d={handLine(AX, AY0, AX, AY1, 's03ay', 1)} p={prog(f, S4 + 8, S4 + 20)} width={4} />
					<DrawPath d={handLine(AX, AY1, AX1, AY1, 's03ax', 1)} p={prog(f, S4 + 12, S4 + 24)} width={4} />
					<DrawPath d={smoothPath(HISTORY)} p={prog(f, S4 + 10, GAIN_AT)} stroke={C.gray} width={6} />
					<DrawPath d={smoothPath(RISE)} p={prog(f, GAIN_AT, GAIN_AT + 26)} stroke={red} width={7} />
					{f >= GAIN_AT + 24 ? <circle cx={END[0]} cy={END[1]} r={11} fill={red} opacity={enterP(f, GAIN_AT + 24, 8)} /> : null}
					{/* 자막 5: 전고점 점선 + 위로 향하는 점선 화살표 */}
					<DrawPath d={`M ${PEAK[0]} ${PEAK[1]} L ${AX1} ${PEAK[1]}`} p={prog(f, S5, S5 + 16)} dash="14 12" width={4} linecap="butt" />
					<DrawPath d={`M ${END[0]} ${END[1] - 22} L ${END[0]} ${PEAK[1] + 22}`} p={prog(f, S5 + 6, S5 + 22)} dash="10 9" width={6} stroke={red} linecap="butt" />
					<DrawPath d={arrowHead([END[0], PEAK[1] + 14], -Math.PI / 2, 24)} p={prog(f, S5 + 20, S5 + 26)} width={6} stroke={red} />
				</Svg>
				<Reveal at={S5 + 4} from="down" dist={18} style={{left: 1560, top: PEAK[1] - 62}}>
					<div style={{...T.label}}>전고점</div>
				</Reveal>
				<ConceptTag at={S4 + 12} x={AX1 - 104} y={AY1 + 12} />
			</Layer>
		</AbsoluteFill>
	);
};

export const S03: React.FC = () => <S03Content />;
