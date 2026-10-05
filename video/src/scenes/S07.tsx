// VERSION: v1.0.0 — 2026-10-05 — S07 (자막 9) $410.44 → 약 $300 급락선(실데이터 3점), 최대 -26% 카운트다운, 종가 기준 -18.5%
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {FACTS, SRC, formatPct} from '../data/facts';
import {sceneTimes} from '../data/timeline';
import {C, alpha} from '../design/colors';
import {SANS} from '../design/fonts';
import {enterP, prog} from '../design/motion';
import {T} from '../design/type';
import {SourceCaption} from '../components/Bits';
import {Counter} from '../components/Counter';
import {DrawPath, Svg} from '../components/Draw';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';

const t = sceneTimes('S07');
const M = FACTS.mdb;
const PCT_AT = t.word(9, '최대');
const CLOSE_AT = t.word(9, '폭락');

// 가격 → y (실제 값 3개만 찍으므로 비율이 실제와 같다)
const P_HI = 420;
const P_LO = 280;
const Y_TOP = 210;
const Y_BOT = 690;
const py = (usd: number) => Y_TOP + ((P_HI - usd) / (P_HI - P_LO)) * (Y_BOT - Y_TOP);
const PA = {x: 200, y: py(M.prevClose.usd)};
const PB = {x: 620, y: py(M.intradayLow.usd)};
const PC = {x: 860, y: py(M.close.usd)};

const Dot: React.FC<{x: number; y: number; p: number; color: string}> = ({x, y, p, color}) =>
	p > 0.001 ? <circle cx={x} cy={y} r={13 * p} fill={color} stroke={C.navy} strokeWidth={4} /> : null;

const PriceLabel: React.FC<{x: number; y: number; at: number; price: string; caption: string; align?: 'left' | 'right'}> = ({
	x,
	y,
	at,
	price,
	caption,
	align = 'left',
}) => (
	<Reveal at={at} from="none" style={{left: align === 'left' ? x : undefined, right: align === 'right' ? 1920 - x : undefined, top: y, textAlign: align}}>
		<div style={{fontFamily: SANS, fontWeight: 900, fontSize: 44, color: C.white, lineHeight: 1.1, fontVariantNumeric: 'tabular-nums'}}>{price}</div>
		<div style={{fontFamily: SANS, fontWeight: 500, fontSize: 26, color: alpha(C.cream, 0.7), marginTop: 4}}>{caption}</div>
	</Reveal>
);

export const S07: React.FC = () => {
	const f = useSceneFrame();
	return (
		<AbsoluteFill>
			<SceneBg tone="navy" />
			<Layer depth="mid">
				<Svg>
					{/* 직전 종가 기준선 */}
					<DrawPath d={`M ${PA.x} ${PA.y} L ${PC.x + 60} ${PA.y}`} p={prog(f, 4, 20)} dash="12 12" width={3} stroke={alpha(C.cream, 0.5)} linecap="butt" />
					<DrawPath d={`M ${PA.x} ${PA.y} L ${PB.x} ${PB.y}`} p={prog(f, 8, 26)} width={9} stroke={C.blue} />
					<DrawPath d={`M ${PB.x} ${PB.y} L ${PC.x} ${PC.y}`} p={prog(f, CLOSE_AT, CLOSE_AT + 14)} width={7} stroke={C.blue} />
					<Dot x={PA.x} y={PA.y} p={enterP(f, 0, 10)} color={C.cream} />
					<Dot x={PB.x} y={PB.y} p={enterP(f, 24, 10)} color={C.blue} />
					<Dot x={PC.x} y={PC.y} p={enterP(f, CLOSE_AT + 12, 10)} color={C.blue} />
				</Svg>
				<PriceLabel x={PA.x - 20} y={PA.y - 112} at={2} price={M.prevClose.label} caption={`${M.prevClose.date} 종가`} />
				<PriceLabel x={PB.x - 70} y={PB.y + 30} at={26} price={M.intradayLow.label} caption={`${M.intradayLow.date} 장중 저점`} />
				<PriceLabel x={PC.x + 30} y={PC.y + 26} at={CLOSE_AT + 14} price={M.close.label} caption={`${M.close.date} 종가`} />
			</Layer>
			<Layer depth="fg">
				<Reveal at={PCT_AT - 4} from="right" style={{left: 1180, top: 190}}>
					<div style={{...T.label, color: C.cream, fontSize: 44}}>최대</div>
				</Reveal>
				<Reveal at={PCT_AT - 4} from="right" style={{left: 1160, top: 250}}>
					<Counter from={0} to={M.intradayLow.pct} at={PCT_AT} format={(v) => formatPct(v)} color={C.blue} style={{fontSize: 230}} />
				</Reveal>
				<Reveal at={CLOSE_AT + 14} from="up" dist={20} style={{left: 1186, top: 530}}>
					<div style={{...T.label, color: C.cream, fontSize: 42}}>
						종가 기준 <span style={{color: C.blue, fontWeight: 900, fontVariantNumeric: 'tabular-nums'}}>{formatPct(M.close.pct, 1)}</span>
					</div>
				</Reveal>
			</Layer>
			<SourceCaption text={SRC.market} at={6} color={alpha(C.cream, 0.6)} />
		</AbsoluteFill>
	);
};
