// VERSION: v2.0.0 — 2026-10-06 — S27 (자막 75–79) "셋째 · 지수 ETF"
// 76: 상자가 열리며 작은 점 500개 그리드가 퍼짐 "S&P 500 = 대표 기업 500곳" → 77: 그중 몇 개의 큰 점이 빛남 "승자 기업도 함께"
// 78: 도넛 90% "S&P 500 지수 펀드"(나머지 10% 회색 "단기 국채") "버핏 90%" → 79: 사연자의 계좌 막대가 절반 ETF + 절반 개별 종목 칸
// 연결 근거: 지수 ETF로 쉽게 분산
import React from 'react';
import {AbsoluteFill, random} from 'remotion';
import {FACTS, SRC} from '../data/facts';
import {sceneTimes} from '../data/timeline';
import {C} from '../design/colors';
import {SERIF} from '../design/fonts';
import {easeInOut, enterP, exitP, lin, prog} from '../design/motion';
import {T} from '../design/type';
import {SceneTitle, SourceCaption} from '../components/Bits';
import {EtfPatternDefs, StackBar} from '../components/Account';
import {Svg, useSafeId} from '../components/Draw';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {roundRect} from '../components/hand';

const t = sceneTimes('S27');
const [S75, S76, S77, S78, S79] = [75, 76, 77, 78, 79].map((n) => t.sub(n));
const E = FACTS.etf;

const COLS = 25;
const ROWS = 20; // 25 × 20 = 500
const GX = 640;
const GY = 252; // 제목(아래끝 ≈ 210)과 띄움. 마지막 줄 252 + 19×28 = 784
const SP = 28;
const BOX = {x: 870, y: 650, w: 180, h: 120};
const ORIGIN = {x: BOX.x + BOX.w / 2, y: BOX.y + 30};
const WINNERS = new Set(Array.from({length: 12}, (_, k) => Math.floor(random(`s27w${k}`) * 500)));
const DONUT = {cx: 800, cy: 490, r: 250, sw: 70}; // 안쪽 지름 430 > "90%" 160px 폭(≈350)
const BAR = {x: 860, bottom: 780, w: 200, full: 500};

export const S27: React.FC = () => {
	const f = useSceneFrame();
	const patId = useSafeId('etf-dots');
	const lid = easeInOut(prog(f, S76 + 6, S76 + 22, (x) => x));
	const out77 = exitP(f, S78 - 2, 9);
	const out78 = exitP(f, S79 - 2, 9);
	const circ = 2 * Math.PI * DONUT.r;
	const inkArc = prog(f, S78 + 8, S78 + 36) * (E.buffettPct / 100);
	const grayArc = prog(f, S78 + 36, S78 + 46) * ((100 - E.buffettPct) / 100);
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer>
				<SceneTitle text="셋째 · 지수 ETF" at={S75} />
				{/* 76–77: 상자 → 500개 점 */}
				{f >= S76 && out77 > 0.001 ? (
					<Svg>
						<g opacity={out77}>
							<g opacity={1 - prog(f, S76 + 40, S76 + 56)}>
								<rect x={BOX.x} y={BOX.y} width={BOX.w} height={BOX.h} rx={8} fill={C.paper} stroke={C.ink} strokeWidth={5} opacity={enterP(f, S76, 10)} />
								<g style={{transformBox: 'view-box', transformOrigin: `${BOX.x - 6}px ${BOX.y}px`, rotate: `${-115 * lid}deg`}} opacity={enterP(f, S76, 10)}>
									<rect x={BOX.x - 10} y={BOX.y - 28} width={BOX.w + 20} height={28} rx={6} fill={C.paper} stroke={C.ink} strokeWidth={5} />
								</g>
							</g>
							{Array.from({length: COLS * ROWS}, (_, i) => {
								const tx = GX + (i % COLS) * SP;
								const ty = GY + Math.floor(i / COLS) * SP;
								const dist = Math.hypot(tx - ORIGIN.x, ty - ORIGIN.y) / 700;
								const p = easeInOut(prog(f, S76 + 16 + dist * 24, S76 + 40 + dist * 24, (x) => x));
								if (p <= 0) return null;
								const win = WINNERS.has(i) ? lin(f, S77 + 4, S77 + 14) : 0;
								return (
									<circle
										key={i}
										cx={ORIGIN.x + (tx - ORIGIN.x) * p}
										cy={ORIGIN.y + (ty - ORIGIN.y) * p}
										r={5.5 + 7 * win}
										fill={win > 0.5 ? C.yellow : C.ink}
										stroke={win > 0.5 ? C.ink : 'none'}
										strokeWidth={3}
									/>
								);
							})}
						</g>
					</Svg>
				) : null}
				<Reveal at={S76 + 30} exitAt={S77 - 2} from="left" style={{left: 104, top: 380}}>
					<div style={{fontFamily: SERIF, fontWeight: 900, fontSize: 72, color: C.ink}}>S&amp;P 500</div>
					<div style={{...T.label, marginTop: 8}}>= 대표 기업 {E.companies}곳</div>
				</Reveal>
				<Reveal at={S77 + 8} exitAt={S78 - 2} from="left" style={{left: 104, top: 420}}>
					<div style={{...T.label, fontSize: 44}}>승자 기업도 함께</div>
				</Reveal>

				{/* 78: 도넛 */}
				{f >= S78 && out78 > 0.001 ? (
					<>
						<Svg>
							<g opacity={out78} transform={`rotate(-90 ${DONUT.cx} ${DONUT.cy})`}>
								<circle cx={DONUT.cx} cy={DONUT.cy} r={DONUT.r} fill="none" stroke="rgba(30,30,30,0.07)" strokeWidth={DONUT.sw} />
								<circle cx={DONUT.cx} cy={DONUT.cy} r={DONUT.r} fill="none" stroke={C.ink} strokeWidth={DONUT.sw} strokeDasharray={`${inkArc * circ} ${circ}`} />
								<circle
									cx={DONUT.cx}
									cy={DONUT.cy}
									r={DONUT.r}
									fill="none"
									stroke={C.gray}
									strokeWidth={DONUT.sw}
									strokeDasharray={`${grayArc * circ} ${circ}`}
									strokeDashoffset={-(E.buffettPct / 100) * circ}
								/>
							</g>
						</Svg>
						<Reveal at={S78 + 20} exitAt={S79 - 2} from="none" style={{left: DONUT.cx - 180, width: 360, top: DONUT.cy - 120, textAlign: 'center'}}>
							<div style={{...T.label}}>버핏</div>
							<div style={{...T.number, fontSize: 160, color: C.ink, marginTop: 4}}>{E.buffettPct}%</div>
						</Reveal>
						<Reveal at={S78 + 30} exitAt={S79 - 2} from="right" style={{left: 1160, top: 410}}>
							<div style={{...T.label, fontSize: 44, display: 'flex', alignItems: 'center', gap: 16}}>
								<span style={{display: 'inline-block', width: 34, height: 34, borderRadius: 6, background: C.ink}} />
								S&amp;P 500 지수 펀드
							</div>
							<div style={{...T.label, color: C.gray, marginTop: 24, display: 'flex', alignItems: 'center', gap: 16}}>
								<span style={{display: 'inline-block', width: 34, height: 34, borderRadius: 6, background: C.gray}} />
								단기 국채
							</div>
						</Reveal>
					</>
				) : null}

				{/* 79: 절반 ETF + 절반 개별 종목 */}
				{f >= S79 ? (
					<>
						<Svg>
							<EtfPatternDefs id={patId} />
							<StackBar
								x={BAR.x}
								bottom={BAR.bottom}
								w={BAR.w}
								etfPatternId={patId}
								opacity={enterP(f, S79 + 4, 14)}
								segs={[{h: BAR.full / 2, kind: 'etf'}, ...Array.from({length: 5}, () => ({h: BAR.full / 10, kind: 'stock' as const}))]}
							/>
							<path d={roundRect(BAR.x, BAR.bottom - BAR.full, BAR.w, BAR.full, 4)} fill="none" stroke={C.ink} strokeWidth={5} opacity={enterP(f, S79 + 4, 14)} />
						</Svg>
						<Reveal at={S79 + 4} from="none" style={{left: BAR.x, top: BAR.bottom - BAR.full - 62}}>
							<div style={{...T.label, color: C.gray}}>예:</div>
						</Reveal>
						<Reveal at={S79 + 12} from="left" style={{left: BAR.x + BAR.w + 30, top: BAR.bottom - BAR.full * 0.75 - 28}}>
							<div style={{...T.label}}>절반 개별 종목</div>
						</Reveal>
						<Reveal at={S79 + 18} from="left" style={{left: BAR.x + BAR.w + 30, top: BAR.bottom - BAR.full * 0.25 - 28}}>
							<div style={{...T.label}}>절반 지수 ETF</div>
						</Reveal>
					</>
				) : null}
			</Layer>
			<SourceCaption text={SRC.berkshire} at={S78 + 20} exitAt={S79 - 2} />
		</AbsoluteFill>
	);
};
