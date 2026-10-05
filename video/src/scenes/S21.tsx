// VERSION: v1.0.0 — 2026-10-05 — S21 (자막 34–35) "올바른 방법": 1억 원 블록이 10개로 쪼개져 10칸에 배치(분산) → 비중 막대가 같은 높이로(비중 조절)
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {FACTS, formatManwon} from '../data/facts';
import {sceneTimes} from '../data/timeline';
import {C} from '../design/colors';
import {SANS} from '../design/fonts';
import {easeInOut, enterP, lerp, prog} from '../design/motion';
import {T} from '../design/type';
import {ConceptTag} from '../components/Bits';
import {DrawPath, Svg} from '../components/Draw';
import {Highlight} from '../components/Highlight';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {handLine, roundRect} from '../components/hand';

const t = sceneTimes('S21');
const N = FACTS.diversify.stocks;
const SPLIT_AT = t.word(34, '현금을');
const DIV_AT = t.word(34, '분산');
const EQ_AT = t.sub(35);

const BIG = {x: 580, y: 270, w: 760, h: 180};
const BASE_Y = 700;
const SLOT_X0 = 166;
const SLOT_W = 148;
const SLOT_GAP = 12;
const H_UNIT = 130;
// 쪼갠 직후의 들쭉날쭉한 비중 (개념도 — 숫자 표시 없음)
const UNEVEN = [1.9, 0.6, 1.4, 0.5, 1.0, 0.8, 1.6, 0.7, 0.9, 0.5];

export const S21: React.FC = () => {
	const f = useSceneFrame();
	const bigIn = enterP(f, 8, 15);
	const bigOut = 1 - prog(f, SPLIT_AT, SPLIT_AT + 6);
	const eq = easeInOut(prog(f, EQ_AT + 6, EQ_AT + 28, (x) => x));
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer depth="bg">
				<Reveal at={0} from="left" style={{left: 160, top: 96}}>
					<div style={{...T.headline, fontSize: 84}}>올바른 방법</div>
				</Reveal>
				<Svg>
					<DrawPath d={handLine(164, 214, 640, 208, 's21ul', 3)} p={prog(f, 8, 22)} width={8} />
				</Svg>
			</Layer>
			<Layer depth="mid">
				<Svg>
					{/* 1억 원 블록 */}
					{bigOut > 0.001 ? (
						<g opacity={Math.min(bigIn, bigOut)}>
							<rect x={BIG.x} y={BIG.y} width={BIG.w} height={BIG.h} rx={18} fill={C.paper} stroke={C.ink} strokeWidth={5} />
						</g>
					) : null}
					{/* 쪼개진 10개 블록 */}
					{f >= SPLIT_AT
						? Array.from({length: N}, (_, i) => {
								const mv = easeInOut(prog(f, SPLIT_AT + 4 + i * 2, SPLIT_AT + 26 + i * 2, (x) => x));
								const h1 = lerp(UNEVEN[i] * H_UNIT, H_UNIT, eq);
								const x = lerp(BIG.x + (i * BIG.w) / N, SLOT_X0 + i * (SLOT_W + SLOT_GAP), mv);
								const w = lerp(BIG.w / N, SLOT_W, mv);
								const h = lerp(BIG.h, h1, mv);
								const y = lerp(BIG.y, BASE_Y - h1, mv);
								return <path key={i} d={roundRect(x, y, w, h, Math.min(12, w / 4))} fill={C.paper} stroke={C.ink} strokeWidth={4} />;
							})
						: null}
					{/* 같은 높이 기준선 */}
					<DrawPath d={`M ${SLOT_X0 - 20} ${BASE_Y - H_UNIT} L ${SLOT_X0 + N * (SLOT_W + SLOT_GAP) + 8} ${BASE_Y - H_UNIT}`} p={prog(f, EQ_AT + 24, EQ_AT + 40)} dash="14 12" width={4} stroke={C.ink} linecap="butt" />
					<DrawPath d={handLine(SLOT_X0 - 20, BASE_Y + 3, SLOT_X0 + N * (SLOT_W + SLOT_GAP) + 8, BASE_Y + 3, 's21base', 1)} p={prog(f, SPLIT_AT + 4, SPLIT_AT + 20)} width={4} />
				</Svg>
				{bigOut > 0.001 ? (
					<div
						style={{
							position: 'absolute',
							left: BIG.x,
							top: BIG.y,
							width: BIG.w,
							height: BIG.h,
							display: 'flex',
							alignItems: 'center',
							justifyContent: 'center',
							fontFamily: SANS,
							fontWeight: 900,
							fontSize: 160,
							color: C.ink,
							opacity: Math.min(bigIn, bigOut),
						}}
					>
						{formatManwon(FACTS.story.principalManwon)}
					</div>
				) : null}
				<ConceptTag at={SPLIT_AT + 20} x={1650} y={BASE_Y + 16} />
			</Layer>
			<Layer depth="fg">
				<Reveal at={DIV_AT - 4} from="up" style={{left: 160, top: 300}}>
					<div style={{...T.headline, fontSize: 96}}>
						<Highlight at={DIV_AT + 4}>분산</Highlight>
					</div>
				</Reveal>
				<Reveal at={EQ_AT} from="up" style={{left: 470, top: 300}}>
					<div style={{...T.headline, fontSize: 96}}>
						<span style={{color: C.gray, marginRight: 40}}>+</span>
						<Highlight at={EQ_AT + 10}>비중 조절</Highlight>
					</div>
				</Reveal>
			</Layer>
		</AbsoluteFill>
	);
};
