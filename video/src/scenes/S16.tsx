// VERSION: v2.0.0 — 2026-10-06 — S16 (자막 30–33) 선택지 카드 3장(물타기, 손절, 손절 라인)
// 자막 31: 물타기 ✕ "현금 0원", 손절 "= 손실 확정". 자막 32–33: 세 번째 카드가 커지며 미니 차트 —
// 손절 라인 점선 위에 있던 주가선이 하루를 건너뛰고 점선보다 훨씬 아래에서 다시 시작 (개념도). 연결 근거: 몰빵 상태에서는 선택지가 없음
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {sceneTimes} from '../data/timeline';
import {C} from '../design/colors';
import {SERIF} from '../design/fonts';
import {easeInOut, enterP, exitP, lerp, prog} from '../design/motion';
import {T} from '../design/type';
import {ConceptTag, cardStyle} from '../components/Bits';
import {DrawPath, Svg} from '../components/Draw';
import {Lightning} from '../components/Icons';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {handLine, smoothPath, Pt} from '../components/hand';

const t = sceneTimes('S16');
const S30 = t.sub(30);
const S31 = t.sub(31);
const S32 = t.sub(32);
const S33 = t.sub(33);

const CARD = {w: 470, h: 300, y: 240};
const CX = [400, 960, 1520];
const BIG = {x: 300, y: 140, w: 1320, h: 640};
// 자막 31의 ✕: "물타기" 제목 위만 덮는다 (아래 "현금 0원"은 가리지 않음)
const XM = {x: CX[0], y: CARD.y + 98, w: 250, h: 120};

// 미니 차트 (큰 카드 안, 개념도)
const STOP_Y = 470;
const DAY1: Pt[] = [[380, 330], [470, 352], [560, 340], [650, 380], [740, 372], [830, 410], [900, 424]];
const DAY2: Pt[] = [[1040, 660], [1140, 676], [1240, 652], [1340, 690], [1460, 672], [1550, 694]];

const Option: React.FC<{readonly i: number; readonly title: string; readonly children?: React.ReactNode; readonly out: number}> = ({i, title, children, out}) => {
	const f = useSceneFrame();
	const p = enterP(f, S30 + i * 6, 15);
	if (p <= 0.001 || out <= 0.001) return null;
	return (
		<div style={{...cardStyle(), left: CX[i] - CARD.w / 2, top: CARD.y, width: CARD.w, height: CARD.h, opacity: Math.min(p, out), translate: `0px ${(1 - p) * 30}px`}}>
			<div style={{position: 'absolute', left: 0, right: 0, top: 54, textAlign: 'center', fontFamily: SERIF, fontWeight: 900, fontSize: 72, color: C.ink}}>{title}</div>
			{children}
		</div>
	);
};

export const S16: React.FC = () => {
	const f = useSceneFrame();
	const out12 = exitP(f, S32, 9);
	const grow = easeInOut(prog(f, S32 + 2, S32 + 22, (x) => x));
	const c3 = enterP(f, S30 + 12, 15);
	const x = lerp(CX[2] - CARD.w / 2, BIG.x, grow);
	const y = lerp(CARD.y, BIG.y, grow);
	const w = lerp(CARD.w, BIG.w, grow);
	const h = lerp(CARD.h, BIG.h, grow);
	const xa = handLine(XM.x - XM.w / 2, XM.y - XM.h / 2, XM.x + XM.w / 2, XM.y + XM.h / 2, 's16xa', 2);
	const xb = handLine(XM.x + XM.w / 2, XM.y - XM.h / 2, XM.x - XM.w / 2, XM.y + XM.h / 2, 's16xb', 2);
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer>
				<Option i={0} title="물타기" out={out12}>
					<Reveal at={S31 + 8} from="none" style={{left: 0, right: 0, top: 190, textAlign: 'center'}}>
						<div style={{...T.label}}>현금 0원</div>
					</Reveal>
				</Option>
				<Option i={1} title="손절" out={out12}>
					<Reveal at={S31 + 14} from="none" style={{left: 0, right: 0, top: 190, textAlign: 'center'}}>
						<div style={{...T.label, color: C.blue}}>= 손실 확정</div>
					</Reveal>
				</Option>
				<Svg>
					<g opacity={out12}>
						<DrawPath d={xa} p={prog(f, S31, S31 + 8)} width={12} />
						<DrawPath d={xb} p={prog(f, S31 + 4, S31 + 12)} width={12} />
					</g>
				</Svg>
				{/* 세 번째 카드: 손절 라인 → 커지며 미니 차트 */}
				{c3 > 0.001 ? (
					<div style={{...cardStyle(), left: x, top: y, width: w, height: h, opacity: c3, translate: `0px ${(1 - c3) * 30}px`}}>
						<div
							style={{
								position: 'absolute',
								left: lerp(0, 36, grow),
								width: lerp(CARD.w - 8, 420, grow),
								top: lerp(54, 26, grow),
								textAlign: lerp(0, 1, grow) > 0.5 ? 'left' : 'center',
								fontFamily: SERIF,
								fontWeight: 900,
								fontSize: lerp(72, 64, grow),
								color: C.ink,
								whiteSpace: 'nowrap',
							}}
						>
							손절 라인
						</div>
					</div>
				) : null}
				{grow > 0.98 ? (
					<>
						<Svg>
							<DrawPath d={`M 360 ${STOP_Y} L 1560 ${STOP_Y}`} p={prog(f, S32 + 22, S32 + 40)} dash="18 12" width={5} stroke={C.ink} linecap="butt" />
							<DrawPath d={smoothPath(DAY1)} p={prog(f, S32 + 30, S32 + 60)} width={7} stroke={C.ink} />
							<line x1={960} y1={250} x2={960} y2={740} stroke={C.gray} strokeWidth={3} strokeDasharray="6 10" opacity={prog(f, S32 + 56, S32 + 66)} />
							<Lightning x={965} y={290} h={150} reveal={prog(f, S33, S33 + 5, (x) => x)} />
							<DrawPath d={`M ${DAY1[DAY1.length - 1][0]} ${DAY1[DAY1.length - 1][1]} L ${DAY2[0][0]} ${DAY2[0][1]}`} p={prog(f, S33 + 8, S33 + 18)} dash="8 10" width={4} stroke={C.blue} linecap="butt" />
							<DrawPath d={smoothPath(DAY2)} p={prog(f, S33 + 14, S33 + 40)} width={7} stroke={C.blue} />
						</Svg>
						<Reveal at={S33 + 8} from="left" dist={20} style={{left: 1040, top: 300}}>
							<div style={{...T.label}}>장 시작 전 악재</div>
						</Reveal>
						<ConceptTag at={S32 + 24} x={1470} y={720} />
					</>
				) : null}
			</Layer>
		</AbsoluteFill>
	);
};
