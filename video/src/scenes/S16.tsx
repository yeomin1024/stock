// VERSION: v1.0.0 — 2026-10-05 — S16 (자막 24–26) CDNS·META·GOOGL 카드 + 파랑 하락 막대 카운트다운. 자막 26: 업종 칩 6개가 차례로 파랑으로 깜빡임
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {FACTS, SRC, formatPct} from '../data/facts';
import {sceneTimes} from '../data/timeline';
import {C, alpha, mix} from '../design/colors';
import {SANS} from '../design/fonts';
import {enterP, exitP, lin} from '../design/motion';
import {SourceCaption, cardStyle} from '../components/Bits';
import {useCount} from '../components/Counter';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';

const t = sceneTimes('S16');
const SECTOR_AT = t.sub(26);
const SECTORS = ['기술', '헬스케어', '금융', '소비재', '에너지', '산업재'];

const CARD_W = 480;
const CARD_Y = 110;
const CARD_H = 160;
const BASE_Y = CARD_Y + CARD_H + 14;
const PX_PER_PCT = 13; // 26% → 338px: 숫자(160px)까지 y<800 안에 들어오게

// 카드 등장: 자막 24 시작에 순차(5f 간격) / 막대: 자막 25에서 이름이 읽힐 때
const COLS = [
	{fact: FACTS.drops[0], cx: 400, show: t.sub(24), drop: t.word(25, 'CDNS')},
	{fact: FACTS.drops[1], cx: 960, show: t.sub(24) + 5, drop: t.word(25, '메타')},
	{fact: FACTS.drops[2], cx: 1520, show: t.sub(24) + 10, drop: t.word(25, '구글')},
];

const DropColumn: React.FC<{readonly c: (typeof COLS)[number]; readonly out: number}> = ({c, out}) => {
	const f = useSceneFrame();
	const card = enterP(f, c.show, 15);
	const v = useCount(0, c.fact.pct, c.drop, 26);
	const h = Math.abs(v) * PX_PER_PCT;
	const decimals = Number.isInteger(c.fact.pct) ? 0 : 1;
	if (card <= 0.001) return null;
	return (
		<div style={{opacity: Math.min(card, out)}}>
			<div
				style={{
					...cardStyle(),
					left: c.cx - CARD_W / 2,
					top: CARD_Y,
					width: CARD_W,
					height: CARD_H,
					translate: `0px ${(1 - card) * -40}px`,
					display: 'flex',
					flexDirection: 'column',
					alignItems: 'center',
					justifyContent: 'center',
				}}
			>
				<div style={{fontFamily: SANS, fontWeight: 900, fontSize: 74, lineHeight: 1, color: C.ink}}>{c.fact.ticker}</div>
				<div style={{fontFamily: SANS, fontWeight: 500, fontSize: 26, color: C.gray, marginTop: 10}}>
					{c.fact.when}
				</div>
			</div>
			{/* 하락 막대: 카드 아래에서 아래로 자란다 */}
			<div style={{position: 'absolute', left: c.cx - 80, top: BASE_Y, width: 160, height: h, background: C.blue, borderRadius: '0 0 10px 10px'}} />
			{f >= c.drop ? (
				<div
					style={{
						position: 'absolute',
						left: c.cx - 280,
						width: 560,
						top: BASE_Y + h + 8,
						textAlign: 'center',
						fontFamily: SANS,
						fontWeight: 900,
						fontSize: 160,
						lineHeight: 1,
						color: C.blue,
						fontVariantNumeric: 'tabular-nums',
						whiteSpace: 'nowrap',
					}}
				>
					{formatPct(v, decimals)}
					{c.fact.note === '장중' ? <span style={{fontSize: 44, fontWeight: 700, marginLeft: 8}}>(장중)</span> : null}
				</div>
			) : null}
		</div>
	);
};

export const S16: React.FC = () => {
	const f = useSceneFrame();
	const out = exitP(f, SECTOR_AT - 2, 9);
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer depth="mid">
				{COLS.map((c) => (
					<DropColumn key={c.fact.ticker} c={c} out={out} />
				))}
				{/* 자막 26: 업종 칩 그리드 */}
				{SECTORS.map((s, i) => {
					const col = i % 3;
					const row = Math.floor(i / 3);
					const at = SECTOR_AT + 8 + i * 5;
					const blinkAt = SECTOR_AT + 40 + i * 10;
					const blink = lin(f, blinkAt, blinkAt + 3) * (1 - lin(f, blinkAt + 7, blinkAt + 11));
					const after = lin(f, blinkAt + 7, blinkAt + 11);
					return (
						<Reveal key={s} at={at} from="up" dist={30} style={{left: 290 + col * 460, top: 230 + row * 190}}>
							<div
								style={{
									width: 420,
									height: 130,
									borderRadius: 65,
									boxSizing: 'border-box',
									border: `5px solid ${mix(C.ink, C.blue, Math.max(blink, after))}`,
									background: mix(C.paper, C.blue, blink),
									display: 'flex',
									alignItems: 'center',
									justifyContent: 'center',
									fontFamily: SANS,
									fontWeight: 700,
									fontSize: 44,
									color: mix(mix(C.ink, C.blue, after), C.white, blink),
									boxShadow: `8px 8px 0 ${alpha(C.ink, 0.1)}`,
								}}
							>
								{s}
							</div>
						</Reveal>
					);
				})}
			</Layer>
			<SourceCaption text={SRC.market} at={COLS[1].drop} exitAt={SECTOR_AT - 2} />
		</AbsoluteFill>
	);
};
