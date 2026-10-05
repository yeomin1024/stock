// VERSION: v1.0.0 — 2026-10-05 — S22 (자막 36–38) 화면 반분할: 몰빵 1억 블록 -30% → -3,000만 원 / 분산 10블록 중 1개 -30% → -300만 원(3%).
// 자막 38: 블록이 비중 막대로 바뀌고 15% 상한 점선.
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {FACTS, formatManwon, formatPct} from '../data/facts';
import {sceneTimes} from '../data/timeline';
import {C, alpha} from '../design/colors';
import {SANS} from '../design/fonts';
import {easeInOut, enterP, lerp, prog} from '../design/motion';
import {T} from '../design/type';
import {Counter} from '../components/Counter';
import {DrawPath, Svg} from '../components/Draw';
import {Highlight} from '../components/Highlight';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {handLine} from '../components/hand';

const t = sceneTimes('S22');
const D = FACTS.diversify;
const S = FACTS.story;
const LOSS_AT = t.sub(37) + 4;
const CAP_AT = t.sub(38) + 4;
const LOSS_FRAC = Math.abs(S.lossPct) / 100; // 같은 폭락 -30%
const HIT = 2; // 분산 쪽에서 폭락한 1개 블록

type Rect = {x: number; y: number; w: number; h: number};
const lerpRect = (a: Rect, b: Rect, k: number): Rect => ({x: lerp(a.x, b.x, k), y: lerp(a.y, b.y, k), w: lerp(a.w, b.w, k), h: lerp(a.h, b.h, k)});

// 블록 배치 (자막 36–37)
const BIG: Rect = {x: 300, y: 400, w: 360, h: 360};
const SMALL = (i: number): Rect => ({x: 1083 + (i % 5) * 146, y: 440 + Math.floor(i / 5) * 146, w: 130, h: 130});
// 비중 막대 배치 (자막 38): 바닥 760, 100% = 380px
const BASE = 760;
const FULL = 380;
const BIG_BAR: Rect = {x: 380, y: BASE - FULL, w: 200, h: FULL};
const SMALL_BAR = (i: number): Rect => {
	const h = (FULL * 100) / D.stocks / 100;
	return {x: 1117 + i * 66, y: BASE - h, w: 52, h};
};
const CAP_Y = BASE - (FULL * D.capPct) / 100;

const Block: React.FC<{readonly a: Rect; readonly b: Rect; readonly morph: number; readonly loss: number; readonly appear: number; readonly label: string; readonly labelSize: number}> = ({
	a,
	b,
	morph,
	loss,
	appear,
	label,
	labelSize,
}) => {
	// 손실: 위쪽 30% 가 파랑 점선 자리로 남고 실물은 줄어든다
	const lossH = a.h * LOSS_FRAC * loss;
	const solidA: Rect = {x: a.x, y: a.y + lossH, w: a.w, h: a.h - lossH};
	const r = lerpRect(solidA, b, morph);
	const ghost = loss * (1 - morph);
	if (appear <= 0.001) return null;
	return (
		<g opacity={appear}>
			{ghost > 0.001 ? (
				<rect x={a.x} y={a.y} width={a.w} height={lossH} rx={10} fill={alpha(C.blue, 0.12 * ghost)} stroke={C.blue} strokeWidth={4} strokeDasharray="10 8" opacity={ghost} />
			) : null}
			<rect x={r.x} y={r.y} width={r.w} height={r.h} rx={Math.min(12, r.w / 4)} fill={loss > 0 && morph < 1 ? mixFill(loss * (1 - morph)) : C.paper} stroke={C.ink} strokeWidth={4} />
			<foreignObject x={a.x} y={a.y + lossH} width={a.w} height={a.h - lossH} opacity={1 - prog(morph, 0, 0.3)}>
				<div style={{width: '100%', height: '100%', display: 'flex', alignItems: 'center', justifyContent: 'center', fontFamily: SANS, fontWeight: 900, fontSize: labelSize, color: C.ink, whiteSpace: 'nowrap'}}>
					{label}
				</div>
			</foreignObject>
		</g>
	);
};
const mixFill = (k: number) => (k > 0 ? alpha(C.blue, 0.1 + 0.1 * k) : C.paper);

export const S22: React.FC = () => {
	const f = useSceneFrame();
	const loss = easeInOut(prog(f, LOSS_AT, LOSS_AT + 20, (x) => x));
	const morphAt = (i: number) => easeInOut(prog(f, CAP_AT + i, CAP_AT + 22 + i, (x) => x));
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer depth="bg">
				<Svg>
					<DrawPath d={handLine(960, 110, 960, 790, 's22div', 2)} p={prog(f, 0, 16)} width={3} stroke={C.gray} />
				</Svg>
				<Reveal at={0} from="up" style={{left: 0, width: 960, top: 90, textAlign: 'center'}}>
					<div style={{...T.headline, fontSize: 80}}>몰빵</div>
				</Reveal>
				<Reveal at={4} from="up" style={{left: 960, width: 960, top: 90, textAlign: 'center'}}>
					<div style={{...T.headline, fontSize: 80}}>분산</div>
				</Reveal>
			</Layer>
			<Layer depth="mid">
				<Svg>
					<Block a={BIG} b={BIG_BAR} morph={morphAt(0)} loss={loss} appear={enterP(f, 8, 15)} label={formatManwon(S.principalManwon)} labelSize={64} />
					{Array.from({length: D.stocks}, (_, i) => (
						<Block
							key={i}
							a={SMALL(i)}
							b={SMALL_BAR(i)}
							morph={morphAt(2 + i)}
							loss={i === HIT ? loss : 0}
							appear={enterP(f, 14 + i * 2, 14)}
							label={formatManwon(D.perStockManwon).replace(' 원', '')}
							labelSize={30}
						/>
					))}
					{/* 15% 상한 점선 */}
					<DrawPath d={`M 120 ${CAP_Y} L 1800 ${CAP_Y}`} p={prog(f, CAP_AT + 24, CAP_AT + 44)} dash="18 12" width={5} stroke={C.ink} linecap="butt" />
				</Svg>
				<Reveal at={CAP_AT + 38} from="up" dist={16} style={{left: 1650, top: CAP_Y - 92}}>
					<div style={{...T.number, fontSize: 72, color: C.ink}}>
						<Highlight at={CAP_AT + 44} top={0.5}>
							{formatPct(D.capPct, 0, false)}
						</Highlight>
					</div>
				</Reveal>
			</Layer>
			<Layer depth="fg">
				<Reveal at={LOSS_AT - 4} from="up" dist={20} style={{left: 0, width: 960, top: 196, textAlign: 'center'}}>
					<Counter from={0} to={S.lossManwon} at={LOSS_AT} format={(v) => formatManwon(v, true)} color={(v) => (v === 0 ? C.ink : C.blue)} style={{fontSize: 160}} />
				</Reveal>
				<Reveal at={LOSS_AT - 4} from="up" dist={20} style={{left: 960, width: 960, top: 196, textAlign: 'center'}}>
					<Counter from={0} to={D.lossManwon} at={LOSS_AT} format={(v) => formatManwon(v, true)} color={(v) => (v === 0 ? C.ink : C.blue)} style={{fontSize: 160}} />
				</Reveal>
				<Reveal at={LOSS_AT + 22} from="up" dist={16} style={{left: 960, width: 960, top: 368, textAlign: 'center'}}>
					<div style={{...T.label, fontSize: 40}}>
						계좌의 <span style={{color: C.blue, fontWeight: 900}}>{formatPct(D.lossPctOfAccount, 0)}</span>
					</div>
				</Reveal>
			</Layer>
		</AbsoluteFill>
	);
};
