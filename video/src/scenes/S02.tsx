// VERSION: v1.0.0 — 2026-10-05 — S02 (자막 3) 지폐 블록 10개가 종목 카드로 빨려 들어가 쌓임, 0원 → 1억 원
import React from 'react';
import {AbsoluteFill, random} from 'remotion';
import {FACTS, formatManwon} from '../data/facts';
import {sceneTimes} from '../data/timeline';
import {C} from '../design/colors';
import {SANS, SERIF} from '../design/fonts';
import {easeInOut, enterP, lerp, prog} from '../design/motion';
import {T} from '../design/type';
import {Svg} from '../components/Draw';
import {Highlight} from '../components/Highlight';
import {Bill} from '../components/Icons';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {cardStyle} from '../components/Bits';

const t = sceneTimes('S02');
const N = FACTS.diversify.stocks; // 지폐 블록 10개 = 1,000만 원 × 10
const PER = FACTS.diversify.perStockManwon;
const BW = 170;
const BH = 86;
const FLY_AT = t.word(3, '전 재산');
const FLY_STEP = 3;
const FLY_DUR = 16;
// 지폐가 모두 쌓인 뒤 (날아가는 지폐와 글자가 겹치지 않게)
const MOLBBANG_AT = Math.max(t.word(3, '몰빵'), FLY_AT + (N - 1) * FLY_STEP + FLY_DUR);

// 카드(트레이) 위치
const CX = 1120;
const CY = 560;
const CW = 600;
const CH = 200;

const scatter = Array.from({length: N}, (_, i) => {
	const col = i % 5;
	const row = Math.floor(i / 5);
	return {
		x: 150 + col * 158 + (random(`s02x${i}`) - 0.5) * 30,
		y: 470 + row * 140 + (random(`s02y${i}`) - 0.5) * 40,
		rot: (random(`s02r${i}`) - 0.5) * 22,
	};
});
const stackPos = (i: number) => ({x: CX + CW / 2 - BW / 2 + (i % 2 ? 6 : -6), y: CY - BH - 6 - i * 17, rot: (i % 3) - 1});

export const S02: React.FC = () => {
	const f = useSceneFrame();
	const landed = Array.from({length: N}, (_, i) => f >= FLY_AT + i * FLY_STEP + FLY_DUR).filter(Boolean).length;
	const cardIn = enterP(f, 0, 15);

	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer depth="mid">
				{/* 종목 카드 */}
				<div style={{...cardStyle(), left: CX, top: CY, width: CW, height: CH, opacity: cardIn, translate: `${(1 - cardIn) * 40}px 0px`, padding: '30px 40px', boxSizing: 'border-box'}}>
					<div style={{fontFamily: SANS, fontWeight: 900, fontSize: 80, lineHeight: 1, color: C.ink}}>{FACTS.mdb.ticker}</div>
					<div style={{...T.label, marginTop: 14}}>{FACTS.mdb.nameKo}</div>
				</div>
				{/* 지폐 블록 */}
				<Svg>
					{scatter.map((s, i) => {
						const a = FLY_AT + i * FLY_STEP;
						const p = easeInOut(prog(f, a, a + FLY_DUR, (x) => x));
						const appear = enterP(f, 2 + i * 2, 12);
						const d = stackPos(i);
						const x = lerp(s.x, d.x, p);
						// 낮은 호로 날아가 왼쪽 위 금액 글자(y≤366)와 겹치지 않게
						const y = lerp(s.y, d.y, p) - Math.sin(Math.PI * p) * 30;
						return <Bill key={i} x={x} y={y} w={BW} h={BH} rotate={lerp(s.rot, d.rot, p)} opacity={appear} />;
					})}
				</Svg>
			</Layer>
			<Layer depth="fg">
				<Reveal at={FLY_AT - 8} from="left" style={{left: 150, top: 130}}>
					<div style={{...T.label, fontSize: 44}}>전 재산</div>
				</Reveal>
				<Reveal at={FLY_AT - 6} from="left" style={{left: 140, top: 196}}>
					<div style={{...T.number, fontSize: 170, color: C.ink}}>{formatManwon(landed * PER)}</div>
				</Reveal>
				<Reveal at={MOLBBANG_AT - 4} from="up" style={{left: 150, top: 430}}>
					<div style={{fontFamily: SERIF, fontWeight: 900, fontSize: 96, lineHeight: 1.2, color: C.ink}}>
						<Highlight at={MOLBBANG_AT + 4}>몰빵</Highlight>
					</div>
				</Reveal>
			</Layer>
		</AbsoluteFill>
	);
};
