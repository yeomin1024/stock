// VERSION: v2.2.0 — 2026-10-08 — 자막 번호 −5(대본에서 엔론 5문장 삭제) · S25 (자막 58–63) "첫째 · 비중 상한"
// 58: 사연자의 계좌 막대(MDB 100%, 잉크) → 59: 15% 상한 점선, MDB 칸이 15%까지로 줄고 "15% = 1,050만 원"
// 60: 막대가 10칸(700만 원씩)으로 나뉨 "10종목 × 700만 원"
// 61: 반분할 — 왼쪽(몰빵) 계좌 전체 -2,000만 원 / 오른쪽(분산) MDB 칸 하나만 줄어 -200만 원(-3%), 동시에 카운트다운
// 62: 오른쪽 10칸이 모두 조금씩 내려감 "시장 전체 하락은 함께" → 63: 다시 한 칸만 크게 줄고 나머지 9칸 그대로 "계좌 붕괴는 막는다"
// 연결 근거: 비중 상한과 분산 계산. 막대 높이 = 금액 비율(전체 340px = 7,000만 원).
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {FACTS, formatManwon, formatPct} from '../data/facts';
import {sceneTimes} from '../data/timeline';
import {C} from '../design/colors';
import {SERIF} from '../design/fonts';
import {easeInOut, enterP, exitP, lerp, prog} from '../design/motion';
import {T} from '../design/type';
import {SceneTitle} from '../components/Bits';
import {Seg, StackBar} from '../components/Account';
import {Counter} from '../components/Counter';
import {DrawPath, Svg} from '../components/Draw';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {handLine} from '../components/hand';

const t = sceneTimes('S25');
const [S58, S59, S60, S61, S62, S63] = [58, 59, 60, 61, 62, 63].map((n) => t.sub(n));
const D = FACTS.diversify;
const CAP = FACTS.cap;
const LOSS = Math.abs(FACTS.story.lossPct) / 100; // 같은 폭락 -29%

const BOTTOM = 780;
const FULL = 340;
const W = 200;
const UNIT = FULL / D.stocks; // 700만 원 = 34px
const CAP_H = (FULL * CAP.pct) / 100;
const LCX = 524; // 왼쪽(몰빵) 열 중심: "-2,000만 원" 160px 폭 ≈ 820px → 114–934 (여백 96 유지)
const MARKET_DIP = 0.25; // 자막 62: 시장 전체가 빠질 때 칸마다 줄어드는 비율 (개념 표현, 숫자 표시 안 함)

export const S25: React.FC = () => {
	const f = useSceneFrame();
	const a59 = easeInOut(prog(f, S59 + 6, S59 + 30, (x) => x));
	const a60 = easeInOut(prog(f, S60 + 4, S60 + 24, (x) => x));
	const move = easeInOut(prog(f, S61, S61 + 20, (x) => x));
	const loss = easeInOut(prog(f, S61 + 14, S61 + 40, (x) => x));
	const dip = easeInOut(prog(f, S62 + 6, S62 + 26, (x) => x)) * (1 - easeInOut(prog(f, S63 + 4, S63 + 24, (x) => x)));
	const barIn = enterP(f, S58 + 6, 15);
	const outPre = exitP(f, S61 - 2, 9);

	// 오른쪽(분산) 막대: 58–60 에는 가운데, 61 부터 오른쪽으로
	const cxR = lerp(960, 1600, move);
	const mdbH = lerp(lerp(FULL, CAP_H, a59), UNIT, a60);
	const restH = FULL - mdbH;
	const divided = a60 > 0.999;
	const segsR: Seg[] = divided
		? [
				{h: UNIT, kind: 'mdb', lost: UNIT * LOSS * loss},
				...Array.from({length: D.stocks - 1}, () => ({h: UNIT, kind: 'stock' as const, lost: UNIT * MARKET_DIP * dip})),
			]
		: [{h: mdbH, kind: 'mdb'}, ...(restH > 0.5 ? [{h: restH, kind: 'stock' as const}] : [])];
	const leftIn = enterP(f, S61 + 4, 15);
	const counterDim = 1 - 0.7 * dip;
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer>
				<SceneTitle text="첫째 · 비중 상한" at={S58} exitAt={S61 - 2} />
				<Svg>
					<StackBar x={cxR - W / 2} bottom={BOTTOM} w={W} segs={segsR} opacity={barIn} />
					{/* 65: 10칸 나눔선 (나뉘는 동안) */}
					{!divided && a60 > 0
						? Array.from({length: D.stocks - 1}, (_, k) => {
								const y = BOTTOM - mdbH - (k + 1) * (restH / (D.stocks - 1));
								return k < D.stocks - 2 ? <line key={k} x1={cxR - W / 2} x2={cxR + W / 2} y1={y} y2={y} stroke={C.ink} strokeWidth={4} opacity={a60} /> : null;
							})
						: null}
					{/* 64: 15% 상한 점선 */}
					<g opacity={outPre}>
						<DrawPath d={`M ${cxR - W / 2 - 70} ${BOTTOM - CAP_H} L ${cxR + W / 2 + 70} ${BOTTOM - CAP_H}`} p={prog(f, S59, S59 + 16)} dash="14 10" width={5} stroke={C.ink} linecap="butt" />
					</g>
					{/* 66: 왼쪽(몰빵) 막대 + 가운데 나눔선 */}
					<StackBar x={LCX - W / 2} bottom={BOTTOM} w={W} segs={[{h: FULL, kind: 'mdb', lost: FULL * LOSS * loss}]} opacity={leftIn} />
					<DrawPath d={handLine(960, 140, 960, 790, 's25div', 2)} p={prog(f, S61 + 4, S61 + 20)} width={3} stroke={C.gray} />
				</Svg>
				<Reveal at={S59 + 14} exitAt={S61 - 2} from="left" dist={20} style={{left: 1150, top: BOTTOM - CAP_H - 30}}>
					<div style={{...T.label}}>
						{CAP.pct}% = {formatManwon(CAP.maxManwon)}
					</div>
				</Reveal>
				<Reveal at={S60 + 20} exitAt={S61 - 2} from="left" dist={20} style={{left: 1150, top: 520}}>
					<div style={{...T.label}}>
						{D.stocks}종목 × {formatManwon(D.perStockManwon)}
					</div>
				</Reveal>
				{/* 66: 머리글 + 동시 카운트다운 */}
				<Reveal at={S61 + 4} from="up" style={{left: LCX - 460, width: 920, top: 128, textAlign: 'center'}}>
					<div style={{fontFamily: SERIF, fontWeight: 900, fontSize: 72, color: C.ink}}>몰빵</div>
				</Reveal>
				<Reveal at={S61 + 8} from="up" style={{left: 960, width: 960, top: 128, textAlign: 'center'}}>
					<div style={{fontFamily: SERIF, fontWeight: 900, fontSize: 72, color: C.ink}}>분산</div>
				</Reveal>
				<Reveal at={S61 + 12} from="up" dist={20} style={{left: LCX - 460, width: 920, top: 236, textAlign: 'center'}}>
					<Counter from={0} to={FACTS.story.lossManwon} at={S61 + 14} format={(v) => formatManwon(v, true)} color={(v) => (v === 0 ? C.ink : C.blue)} style={{fontSize: 160}} />
				</Reveal>
				<Reveal at={S61 + 12} from="up" dist={20} style={{left: 960, width: 960, top: 236, textAlign: 'center', opacity: counterDim}}>
					<Counter from={0} to={D.lossManwon} at={S61 + 14} format={(v) => formatManwon(v, true)} color={(v) => (v === 0 ? C.ink : C.blue)} style={{fontSize: 160}} />
					<div style={{...T.label, color: C.blue, marginTop: 6, fontVariantNumeric: 'tabular-nums'}}>({formatPct(D.lossPctOfAccount)})</div>
				</Reveal>
				{/* 67 / 68 라벨 */}
				<Reveal at={S62 + 4} exitAt={S63 - 2} from="right" dist={20} style={{left: 1000, width: 460, top: 600, textAlign: 'right'}}>
					<div style={{...T.label, fontSize: 38}}>시장 전체 하락은 함께</div>
				</Reveal>
				<Reveal at={S63 + 4} from="right" dist={20} style={{left: 1000, width: 460, top: 600, textAlign: 'right'}}>
					<div style={{...T.label, fontSize: 38}}>계좌 붕괴는 막는다</div>
				</Reveal>
			</Layer>
		</AbsoluteFill>
	);
};
