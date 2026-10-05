// VERSION: v1.0.0 — 2026-10-05 — S01 (자막 1–2) 스마트폰 검색 → MDB 결과 카드 → 체크 3개
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {FACTS} from '../data/facts';
import {sceneTimes} from '../data/timeline';
import {C} from '../design/colors';
import {SANS} from '../design/fonts';
import {enterP, prog} from '../design/motion';
import {T} from '../design/type';
import {DrawPath, Svg} from '../components/Draw';
import {checkPath, Magnifier} from '../components/Icons';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {roundRect} from '../components/hand';

const t = sceneTimes('S01');
const QUERY = Array.from('AI 관련주');

// 폰 프레임
const PX = 730;
const PY = 70;
const PW = 460;
const PH = 720;
const TYPE_AT = 24;
const TYPE_STEP = 3;
const CARD_AT = t.word(1, '몽고디비');
const CHECKS = [
	{label: '전망', at: t.word(2, '전망')},
	{label: '실적', at: t.word(2, '실적')},
	{label: '상승세', at: t.word(2, '상승세')},
];

export const S01: React.FC = () => {
	const f = useSceneFrame();
	const phoneP = prog(f, 0, 20);
	const fillP = prog(f, 12, 24);
	const typed = Math.max(0, Math.min(QUERY.length, Math.floor((f - TYPE_AT) / TYPE_STEP) + 1));
	const cursorOn = f < CARD_AT && Math.floor(f / 8) % 2 === 0 && f >= TYPE_AT - 6;
	const card = enterP(f, CARD_AT, 16);

	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer depth="mid">
				<Svg>
					{/* 폰 화면 바탕 */}
					<rect x={PX} y={PY} width={PW} height={PH} rx={58} fill={C.paper} opacity={fillP} />
					<DrawPath d={roundRect(PX, PY, PW, PH, 58)} p={phoneP} width={7} />
					<rect x={PX + PW / 2 - 50} y={PY + 22} width={100} height={14} rx={7} fill={C.ink} opacity={fillP} />
					{/* 검색창 */}
					<rect x={PX + 34} y={PY + 70} width={PW - 68} height={78} rx={39} fill={C.white} opacity={fillP} />
					<DrawPath d={roundRect(PX + 34, PY + 70, PW - 68, 78, 39)} p={prog(f, 10, 26)} width={4} />
					<Magnifier cx={PX + 78} cy={PY + 104} r={14} p={prog(f, 16, 28)} />
				</Svg>
				{/* 타이핑 */}
				<div
					style={{
						position: 'absolute',
						left: PX + 116,
						top: PY + 84,
						...T.label,
						fontSize: 40,
						lineHeight: '50px',
						display: 'flex',
						alignItems: 'center',
					}}
				>
					{QUERY.slice(0, typed).join('')}
					<span style={{display: 'inline-block', width: 4, height: 42, marginLeft: 4, background: C.ink, opacity: cursorOn ? 1 : 0}} />
				</div>
				{/* 결과 카드 */}
				{card > 0.001 ? (
					<div
						style={{
							position: 'absolute',
							left: PX + 34,
							top: PY + 180,
							width: PW - 68,
							height: 190,
							background: C.white,
							border: `4px solid ${C.ink}`,
							borderRadius: 22,
							boxShadow: '8px 8px 0 rgba(30,30,30,0.13)',
							opacity: card,
							translate: `0px ${(1 - card) * 50}px`,
							padding: '26px 30px',
							boxSizing: 'border-box',
						}}
					>
						<div style={{fontFamily: SANS, fontWeight: 900, fontSize: 78, lineHeight: 1, color: C.ink}}>
							{FACTS.mdb.ticker}
						</div>
						<div style={{...T.label, marginTop: 16, fontSize: 40}}>{FACTS.mdb.nameKo}</div>
					</div>
				) : null}
				{/* 체크 3개 */}
				{CHECKS.map((c, i) => {
					const y = PY + 420 + i * 92;
					const boxP = enterP(f, c.at - 4, 12);
					return (
						<React.Fragment key={c.label}>
							<Svg>
								<rect
									x={PX + 46}
									y={y}
									width={56}
									height={56}
									rx={10}
									fill={C.white}
									stroke={C.ink}
									strokeWidth={4}
									opacity={boxP}
								/>
								<DrawPath d={checkPath(PX + 52, y + 4, 50)} p={prog(f, c.at, c.at + 10)} stroke={C.ink} width={8} />
							</Svg>
							<div
								style={{
									position: 'absolute',
									left: PX + 128,
									top: y + 2,
									...T.label,
									fontSize: 42,
									lineHeight: '52px',
									opacity: boxP,
									translate: `${(1 - boxP) * -16}px 0px`,
								}}
							>
								{c.label}
							</div>
						</React.Fragment>
					);
				})}
			</Layer>
		</AbsoluteFill>
	);
};
