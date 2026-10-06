// VERSION: v2.0.0 — 2026-10-06 — S26 (자막 69–74) "둘째 · 업종 분산"
// 70: 카드 3장(AI 반도체, AI 소프트웨어, AI 전력주)이 한 바구니 테두리 안으로 모이고 "같은 테마 = 사실상 한 종목"
// 71: AI 반도체 카드 NVDA·AVGO -17% 막대 → 72: AI 전력주 카드 비스트라 -28% 막대
// 73: 서로 떨어진 업종 칩 4개(반도체, 헬스케어, 소비재, 금융) → 74: 번개가 한 칩에만, 나머지 3개는 그대로
// 연결 근거: 같은 테마는 함께 무너짐
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {FACTS, SRC} from '../data/facts';
import {sceneTimes} from '../data/timeline';
import {C, mix} from '../design/colors';
import {SANS} from '../design/fonts';
import {easeInOut, enterP, exitP, lerp, lin, prog} from '../design/motion';
import {T} from '../design/type';
import {SceneTitle, SourceCaption} from '../components/Bits';
import {DropCard} from '../components/Common';
import {DrawPath, Svg} from '../components/Draw';
import {Highlight} from '../components/Highlight';
import {Lightning} from '../components/Icons';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {roundRect} from '../components/hand';

const t = sceneTimes('S26');
const [S69, S70, S71, S72, S73, S74] = [69, 70, 71, 72, 73, 74].map((n) => t.sub(n));
const TH = FACTS.theme;
const FROM = [380, 960, 1540];
const TO = [520, 960, 1400];
const CARD_TOP = 250;
const CARD_H = 150;
const CHIPS = [
	{label: TH.sectors[0], x: 1460, y: 330},
	{label: TH.sectors[1], x: 470, y: 380},
	{label: TH.sectors[2], x: 760, y: 650},
	{label: TH.sectors[3], x: 1300, y: 640},
];
const HIT = 0; // 번개가 떨어지는 칩: 반도체

export const S26: React.FC = () => {
	const f = useSceneFrame();
	const gather = easeInOut(prog(f, S70 + 20, S70 + 44, (x) => x));
	const out72 = exitP(f, S73 - 2, 9);
	const basket = prog(f, S70 + 36, S70 + 56);
	const strike = prog(f, S74 + 10, S74 + 15, (x) => x);
	const hit = lin(f, S74 + 14, S74 + 20);
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer>
				<SceneTitle text="둘째 · 업종 분산" at={S69} />
				{/* 70–72: 같은 테마 카드 3장 + 바구니 */}
				<Svg>
					<g opacity={out72}>
						<DrawPath d={roundRect(270, CARD_TOP - 24, 1380, CARD_H + 48, 30)} p={basket} width={6} dash="18 12" stroke={C.ink} linecap="butt" />
					</g>
				</Svg>
				{TH.cards.map((label, i) => (
					<DropCard
						key={label}
						cx={lerp(FROM[i], TO[i], gather)}
						top={CARD_TOP}
						w={400}
						h={CARD_H}
						title={label}
						titleSize={50}
						line={i === 0 ? TH.semi.names : i === 2 ? TH.power.names : undefined}
						reserveLine
						at={S70 + i * 6}
						dropAt={i === 0 ? S71 : i === 2 ? S72 : undefined}
						pct={i === 0 ? TH.semi.pct : i === 2 ? TH.power.pct : undefined}
						pxPerPct={7}
						numberSize={160}
						opacity={out72}
					/>
				))}
				<Reveal at={S70 + 44} exitAt={S73 - 2} from="up" dist={16} style={{right: 96, top: 140}}>
					<div style={{...T.label, fontSize: 44}}>
						같은 테마 = <Highlight at={S70 + 52}>사실상 한 종목</Highlight>
					</div>
				</Reveal>
				{/* 73–74: 서로 다른 업종 칩 */}
				{CHIPS.map((c, i) => {
					const p = enterP(f, S73 + 4 + i * 5, 14);
					const h = i === HIT ? hit : 0;
					return p > 0.001 ? (
						<div
							key={c.label}
							style={{
								position: 'absolute',
								left: c.x - 170,
								top: c.y - 60,
								width: 340,
								height: 120,
								borderRadius: 60,
								border: `5px solid ${mix(C.ink, C.blue, h)}`,
								background: mix(C.paper, C.blue, h),
								boxSizing: 'border-box',
								display: 'flex',
								alignItems: 'center',
								justifyContent: 'center',
								fontFamily: SANS,
								fontWeight: 700,
								fontSize: 44,
								color: mix(C.ink, C.white, h),
								opacity: p,
								translate: `0px ${(1 - p) * 24}px`,
							}}
						>
							{c.label}
						</div>
					) : null;
				})}
				<Svg>
					<Lightning x={CHIPS[HIT].x + 10} y={CHIPS[HIT].y - 196} h={140} reveal={strike} />
				</Svg>
			</Layer>
			<SourceCaption text={SRC.reuters} at={S71} exitAt={S73 - 2} />
		</AbsoluteFill>
	);
};

