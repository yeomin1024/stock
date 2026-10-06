// VERSION: v2.0.0 — 2026-10-06 — S29 (자막 85–90) "다섯째 · 수익 나면 비중 확인"
// 86: S03의 +1,000만 원 계좌 카드가 작게 다시 + 생각 말풍선 "더 오를 것 같은데…"
// 87: 10칸 중 한 칸이 2배로 커지며 비중 10% → 18% → 88: 그 칸에 경고 "몰빵에 가까워짐"
// 89: 15% 상한 점선 위로 튀어나온 부분만 잘려 다른 칸들로 나뉘어 들어감 "15% 넘는 만큼만 덜기"
// 90: 균형 잡힌 10칸 중 한 칸에 번개가 떨어져도 나머지가 버팀 "대처할 여유". 연결 근거: 수익 난 종목의 비중 관리
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {FACTS} from '../data/facts';
import {sceneTimes} from '../data/timeline';
import {C} from '../design/colors';
import {easeInOut, enterP, exitP, lerp, prog} from '../design/motion';
import {T} from '../design/type';
import {SceneTitle} from '../components/Bits';
import {AccountCard} from '../components/Account';
import {ThoughtBubble} from '../components/Common';
import {Counter} from '../components/Counter';
import {DrawPath, Svg} from '../components/Draw';
import {Highlight} from '../components/Highlight';
import {Lightning, warnPaths} from '../components/Icons';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';

const t = sceneTimes('S29');
const [S85, S86, S87, S88, S89, S90] = [85, 86, 87, 88, 89, 90].map((n) => t.sub(n));
const RB = FACTS.rebalance;
const N = 10;
const BW = 100;
const GAP = 30;
const X0 = (1920 - (N * BW + (N - 1) * GAP)) / 2;
const BOTTOM = 770;
const UNIT = 120; // 700만 원
const FOCUS = 3;
const STRUCK = 7;
// 15% 상한 높이: 2배가 된 뒤 계좌 = 11칸 분량 → 15% = 1.65칸
const CAP_H = UNIT * N * 1.1 * (FACTS.cap.pct / 100);
const EXCESS = UNIT * 2 - CAP_H;
const LOSS = Math.abs(FACTS.story.lossPct) / 100;

export const S29: React.FC = () => {
	const f = useSceneFrame();
	const out86 = exitP(f, S87 - 2, 9);
	const grow = easeInOut(prog(f, S87 + 10, S87 + 34, (x) => x));
	const trim = easeInOut(prog(f, S89 + 20, S89 + 50, (x) => x));
	const strike = prog(f, S90 + 12, S90 + 17, (x) => x);
	const lost = easeInOut(prog(f, S90 + 16, S90 + 34, (x) => x)) * UNIT * LOSS;
	const barsIn = (i: number) => enterP(f, S87 + i * 2, 12);
	const hFocus = UNIT * (1 + grow) - EXCESS * trim;
	const hOther = UNIT + (EXCESS / (N - 1)) * trim;
	const fx = X0 + FOCUS * (BW + GAP);
	const w = warnPaths(fx + BW / 2, BOTTOM - UNIT * 2 - 70, 80);
	const capY = BOTTOM - CAP_H;
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer>
				<SceneTitle text="다섯째 · 수익 나면 비중 확인" at={S85} />
				{/* 86: S03 계좌 카드(작게) + 말풍선 */}
				<AccountCard x={96} y={250} scale={0.8} pnlMode={1} totalManwon={FACTS.story.principalManwon} pnlManwon={FACTS.story.gainManwon} bar={1} opacity={Math.min(enterP(f, S86, 15), out86)} />
				<ThoughtBubble cx={1480} cy={440} rx={320} ry={120} text="더 오를 것 같은데…" at={S86 + 10} tail="left" opacity={out86} />

				{/* 87–90: 10칸 */}
				{f >= S87 ? (
					<Svg>
						{Array.from({length: N}, (_, i) => {
							const x = X0 + i * (BW + GAP);
							const isF = i === FOCUS;
							const h = isF ? hFocus : hOther;
							const l = i === STRUCK ? lost : 0;
							return (
								<g key={i} opacity={barsIn(i)}>
									{l > 0.5 ? <rect x={x} y={BOTTOM - h} width={BW} height={l} fill="rgba(45,108,223,0.12)" stroke={C.blue} strokeWidth={3} strokeDasharray="8 6" /> : null}
									<rect x={x} y={BOTTOM - h + l} width={BW} height={h - l} rx={6} fill={isF ? C.ink : C.paper} stroke={C.ink} strokeWidth={4} />
								</g>
							);
						})}
						{/* 89: 잘려 나가는 윗부분 → 9칸으로 나뉘어 들어감 */}
						{f >= S89 + 20 && trim < 1
							? Array.from({length: N - 1}, (_, k) => {
									const target = k < FOCUS ? k : k + 1;
									const tx = X0 + target * (BW + GAP);
									const sx = fx;
									const sliceH = EXCESS / (N - 1);
									const sy = capY - EXCESS + k * sliceH;
									const ty = BOTTOM - hOther;
									return (
										<rect key={k} x={lerp(sx, tx, trim)} y={lerp(sy, ty, trim) - (1 - Math.abs(2 * trim - 1)) * 60} width={BW} height={sliceH} fill={C.ink} opacity={1 - trim * 0.3} />
									);
								})
							: null}
						{/* 88: 경고 */}
						<g opacity={Math.min(enterP(f, S88, 12), exitP(f, S89 - 2, 9))}>
							<path d={w.tri} fill={C.paper} stroke={C.ink} strokeWidth={6} strokeLinejoin="round" />
							<path d={w.bar} stroke={C.ink} strokeWidth={8} strokeLinecap="round" />
							<circle cx={w.dot[0]} cy={w.dot[1]} r={5} fill={C.ink} />
						</g>
						{/* 89: 15% 상한 점선 */}
						<DrawPath d={`M ${X0 - 30} ${capY} L ${X0 + N * (BW + GAP)} ${capY}`} p={prog(f, S89 + 4, S89 + 20)} dash="14 10" width={5} linecap="butt" />
						<Lightning x={X0 + STRUCK * (BW + GAP) + BW / 2 + 6} y={BOTTOM - hOther - 170} h={150} reveal={strike} />
					</Svg>
				) : null}
				<Reveal at={S87 + 6} exitAt={S90 - 2} from="right" style={{left: 1180, top: 136}}>
					<div style={{display: 'flex', alignItems: 'baseline', gap: 18}}>
						<span style={{...T.label}}>비중</span>
						<Counter
							from={RB.fromPct}
							to={RB.toPct}
							at={S87 + 10}
							dur={24}
							steps={1}
							format={(v) => `${Math.round(v)}%`}
							color={C.ink}
							style={{fontSize: 160}}
						/>
					</div>
				</Reveal>
				<Reveal at={S88 + 6} exitAt={S89 - 2} from="left" style={{left: fx + BW + 40, top: BOTTOM - UNIT * 2 - 96}}>
					<div style={{...T.label}}>몰빵에 가까워짐</div>
				</Reveal>
				<Reveal at={S89 + 10} exitAt={S90 - 2} from="left" style={{left: X0 + 5 * (BW + GAP), top: capY - 76}}>
					<div style={{...T.label}}>{FACTS.cap.pct}% 넘는 만큼만 덜기</div>
				</Reveal>
				<Reveal at={S90 + 30} from="right" style={{left: 1180, top: 160}}>
					<div style={{...T.headline, fontSize: 80}}>
						<Highlight at={S90 + 40}>대처할 여유</Highlight>
					</div>
				</Reveal>
			</Layer>
		</AbsoluteFill>
	);
};
