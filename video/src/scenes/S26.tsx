// VERSION: v2.2.0 — 2026-10-08 — 자막 번호 −5(대본에서 엔론 5문장 삭제) · S26 (자막 64–69) "둘째 · 업종 분산"
// 65: 카드 3장(AI 반도체, AI 소프트웨어, AI 전력주)이 한 바구니 테두리 안으로 모이고 "같은 테마 = 사실상 한 종목"
// 66: AI 반도체 카드 NVDA·AVGO -17% 막대 → 67: AI 전력주 카드 비스트라 -28% 막대
// 68: 서로 떨어진 업종 칩 4개(반도체, 헬스케어, 소비재, 금융) → 69: 번개가 한 칩에만, 나머지 3개는 그대로
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
const [S64, S65, S66, S67, S68, S69] = [64, 65, 66, 67, 68, 69].map((n) => t.sub(n));
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
	const gather = easeInOut(prog(f, S65 + 20, S65 + 44, (x) => x));
	const out67 = exitP(f, S68 - 2, 9);
	const basket = prog(f, S65 + 36, S65 + 56);
	const strike = prog(f, S69 + 10, S69 + 15, (x) => x);
	const hit = lin(f, S69 + 14, S69 + 20);
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer>
				<SceneTitle text="둘째 · 업종 분산" at={S64} />
				{/* 70–72: 같은 테마 카드 3장 + 바구니 */}
				<Svg>
					<g opacity={out67}>
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
						at={S65 + i * 6}
						dropAt={i === 0 ? S66 : i === 2 ? S67 : undefined}
						pct={i === 0 ? TH.semi.pct : i === 2 ? TH.power.pct : undefined}
						pxPerPct={7}
						numberSize={160}
						opacity={out67}
					/>
				))}
				<Reveal at={S65 + 44} exitAt={S68 - 2} from="up" dist={16} style={{right: 96, top: 140}}>
					<div style={{...T.label, fontSize: 44}}>
						같은 테마 = <Highlight at={S65 + 52}>사실상 한 종목</Highlight>
					</div>
				</Reveal>
				{/* 73–74: 서로 다른 업종 칩 */}
				{CHIPS.map((c, i) => {
					const p = enterP(f, S68 + 4 + i * 5, 14);
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
			<SourceCaption text={SRC.reuters} at={S66} exitAt={S68 - 2} />
		</AbsoluteFill>
	);
};

