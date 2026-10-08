// VERSION: v2.3.0 — 2026-10-08 — 한 줄 자막: 뒷줄 내용 요소는 t.line(n, 2) · 자막 번호 −5(대본에서 엔론 5문장 삭제) · S27 (자막 70–74) "셋째 · 지수 ETF"
// 71: 상자가 열리며 작은 점 500개 그리드가 퍼짐 "S&P 500 = 대표 기업 500곳" → 72: 그중 몇 개의 큰 점이 빛남 "승자 기업도 함께"
// 73: 도넛 90% "S&P 500 지수 펀드"(나머지 10% 회색 "단기 국채") "버핏 90%" → 74: 사연자의 계좌 막대가 절반 ETF + 절반 개별 종목 칸
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
const [S70, S71, S72, S73, S74] = [70, 71, 72, 73, 74].map((n) => t.sub(n));
const S71L2 = t.line(71, 2); // 뒷줄 "미국 대표 기업 500곳에 한 번에 나눠 투자하는 셈이죠."
const S73L2 = t.line(73, 2); // 뒷줄 "S&P 500 지수 펀드에 넣으라고 했습니다."
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
	const lid = easeInOut(prog(f, S71 + 6, S71 + 22, (x) => x));
	const out72 = exitP(f, S73 - 2, 9);
	const out73 = exitP(f, S74 - 2, 9);
	const circ = 2 * Math.PI * DONUT.r;
	const inkArc = prog(f, S73 + 8, S73 + 36) * (E.buffettPct / 100);
	const grayArc = prog(f, S73L2 + 8, S73L2 + 18) * ((100 - E.buffettPct) / 100);
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer>
				<SceneTitle text="셋째 · 지수 ETF" at={S70} />
				{/* 76–77: 상자 → 500개 점 */}
				{f >= S71 && out72 > 0.001 ? (
					<Svg>
						<g opacity={out72}>
							<g opacity={1 - prog(f, S71L2 + 24, S71L2 + 40)}>
								<rect x={BOX.x} y={BOX.y} width={BOX.w} height={BOX.h} rx={8} fill={C.paper} stroke={C.ink} strokeWidth={5} opacity={enterP(f, S71, 10)} />
								<g style={{transformBox: 'view-box', transformOrigin: `${BOX.x - 6}px ${BOX.y}px`, rotate: `${-115 * lid}deg`}} opacity={enterP(f, S71, 10)}>
									<rect x={BOX.x - 10} y={BOX.y - 28} width={BOX.w + 20} height={28} rx={6} fill={C.paper} stroke={C.ink} strokeWidth={5} />
								</g>
							</g>
							{Array.from({length: COLS * ROWS}, (_, i) => {
								const tx = GX + (i % COLS) * SP;
								const ty = GY + Math.floor(i / COLS) * SP;
								const dist = Math.hypot(tx - ORIGIN.x, ty - ORIGIN.y) / 700;
								const p = easeInOut(prog(f, S71L2 + dist * 24, S71L2 + 24 + dist * 24, (x) => x));
								if (p <= 0) return null;
								const win = WINNERS.has(i) ? lin(f, S72 + 4, S72 + 14) : 0;
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
				<Reveal at={S71 + 12} exitAt={S72 - 2} from="left" style={{left: 104, top: 380}}>
					<div style={{fontFamily: SERIF, fontWeight: 900, fontSize: 72, color: C.ink}}>S&amp;P 500</div>
				</Reveal>
				<Reveal at={S71L2 + 20} exitAt={S72 - 2} from="left" style={{left: 104, top: 476}}>
					<div style={{...T.label}}>= 대표 기업 {E.companies}곳</div>
				</Reveal>
				<Reveal at={S72 + 8} exitAt={S73 - 2} from="left" style={{left: 104, top: 420}}>
					<div style={{...T.label, fontSize: 44}}>승자 기업도 함께</div>
				</Reveal>

				{/* 78: 도넛 */}
				{f >= S73 && out73 > 0.001 ? (
					<>
						<Svg>
							<g opacity={out73} transform={`rotate(-90 ${DONUT.cx} ${DONUT.cy})`}>
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
						<Reveal at={S73 + 20} exitAt={S74 - 2} from="none" style={{left: DONUT.cx - 180, width: 360, top: DONUT.cy - 120, textAlign: 'center'}}>
							<div style={{...T.label}}>버핏</div>
							<div style={{...T.number, fontSize: 160, color: C.ink, marginTop: 4}}>{E.buffettPct}%</div>
						</Reveal>
						<Reveal at={S73L2 + 4} exitAt={S74 - 2} from="right" style={{left: 1160, top: 410}}>
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
				{f >= S74 ? (
					<>
						<Svg>
							<EtfPatternDefs id={patId} />
							<StackBar
								x={BAR.x}
								bottom={BAR.bottom}
								w={BAR.w}
								etfPatternId={patId}
								opacity={enterP(f, S74 + 4, 14)}
								segs={[{h: BAR.full / 2, kind: 'etf'}, ...Array.from({length: 5}, () => ({h: BAR.full / 10, kind: 'stock' as const}))]}
							/>
							<path d={roundRect(BAR.x, BAR.bottom - BAR.full, BAR.w, BAR.full, 4)} fill="none" stroke={C.ink} strokeWidth={5} opacity={enterP(f, S74 + 4, 14)} />
						</Svg>
						<Reveal at={S74 + 4} from="none" style={{left: BAR.x, top: BAR.bottom - BAR.full - 62}}>
							<div style={{...T.label, color: C.gray}}>예:</div>
						</Reveal>
						<Reveal at={S74 + 12} from="left" style={{left: BAR.x + BAR.w + 30, top: BAR.bottom - BAR.full * 0.75 - 28}}>
							<div style={{...T.label}}>절반 개별 종목</div>
						</Reveal>
						<Reveal at={S74 + 18} from="left" style={{left: BAR.x + BAR.w + 30, top: BAR.bottom - BAR.full * 0.25 - 28}}>
							<div style={{...T.label}}>절반 지수 ETF</div>
						</Reveal>
					</>
				) : null}
			</Layer>
			<SourceCaption text={SRC.berkshire} at={S73 + 20} exitAt={S74 - 2} />
		</AbsoluteFill>
	);
};
