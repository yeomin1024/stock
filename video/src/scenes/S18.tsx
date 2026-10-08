// VERSION: v2.3.0 — 2026-10-08 — 한 줄 자막: 뒷줄 내용 요소는 t.line(n, 2) · S18 (자막 39–44) "03 확률 자체가 낮다" → 40: 90년 타임라인 + 약 2만 6천 개
// 41: 10×10 와플(100칸 = 전체 종목) 중 58칸 파랑 + "58% 국채보다 못함" → 42: 한 줄 10칸 중 6칸 강조 "열 중 여섯"
// 43: 4칸만 노랑, 나머지 흐림 "상위 4%" → 44: 커서가 칸 위를 헤매다 노랑 4칸을 찾음 "4%를 맞혀야 하는 게임"
// 연결 근거: 베셈바인더 연구. 파랑 58칸은 줄마다 6칸(마지막 두 줄은 5칸) → 어느 줄을 봐도 "열 중 여섯".
import React from 'react';
import {AbsoluteFill, random} from 'remotion';
import {FACTS, SRC, formatPct, formatThousandStocks} from '../data/facts';
import {sceneTimes} from '../data/timeline';
import {C, alpha} from '../design/colors';
import {SERIF} from '../design/fonts';
import {easeOut, lin, prog} from '../design/motion';
import {T} from '../design/type';
import {NumberTitle, SourceCaption} from '../components/Bits';
import {Counter} from '../components/Counter';
import {DrawPath, Svg} from '../components/Draw';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {Waffle, waffleCellXY} from '../components/Waffle';
import {handLine, roundRect} from '../components/hand';

const t = sceneTimes('S18');
const B = FACTS.bessembinder;
const S = [39, 40, 41, 42, 43, 44].map((n) => t.sub(n));
const [S39, S40, S41, S42, S43, S44] = S;
const S40L2 = t.line(40, 2); // 뒷줄 "약 2만 6천 개의 90년치 기록을…"
const S43L2 = t.line(43, 2); // 뒷줄 "전부 상위 4% 종목에서 나왔습니다."

const G = {x: 220, y: 150, cell: 54, gap: 6};
const blueCols = (row: number) => (row < 8 ? 6 : 5); // 8×6 + 2×5 = 58
const isBlue = (i: number) => i % 10 < blueCols(Math.floor(i / 10));
const BLUE_ORDER = Array.from({length: 100}, (_, i) => i).filter(isBlue);
const TARGETS = [18, 36, 69, 97]; // 상위 4% (파랑이 아닌 칸)
const EMPTY = alpha(C.ink, 0.09);

// 자막 44 커서 경로 (결정적): 정해진 이동 번호에서만 노랑 칸, 나머지는 노랑이 아닌 칸
const HOP = 4;
const HITS = [8, 16, 24, 31];
const PATH = Array.from({length: HITS[HITS.length - 1] + 1}, (_, k) => {
	const h = HITS.indexOf(k);
	if (h >= 0) return TARGETS[h];
	let c = Math.floor(random(`s18c${k}`) * 100);
	while (TARGETS.includes(c)) c = (c + 7) % 100;
	return c;
});
const foundAt = (ti: number) => S44 + 6 + HITS[ti] * HOP;

export const S18: React.FC = () => {
	const f = useSceneFrame();
	const fillAt = S41 + 8;
	const rowFocus = prog(f, S42, S42 + 10) * (1 - prog(f, S43, S43 + 8));
	const dimAll = prog(f, S43L2, S43L2 + 10);
	const k = Math.min(PATH.length - 1, Math.max(0, Math.floor((f - S44 - 6) / HOP)));
	const sub = (f - S44 - 6 - k * HOP) / HOP;
	const cur = waffleCellXY(PATH[k], G.x, G.y, G.cell, G.gap);
	const prev = waffleCellXY(PATH[Math.max(0, k - 1)], G.x, G.y, G.cell, G.gap);
	const tw = easeOut(Math.min(1, Math.max(0, sub * 1.6)));
	const cursorOn = f >= S44 + 6 && f < foundAt(3) + 12;
	const waffleIn = prog(f, S41, S41 + 12);
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer>
				<NumberTitle num="03" title="확률 자체가 낮다" mark="확률 자체가 낮다" at={S39} markAt={S39 + 24} exitAt={S40 - 2} />

				{/* 자막 40: 90년 타임라인 + 약 2만 6천 개 */}
				<Svg>
					<g opacity={1 - prog(f, S41 - 6, S41 + 2)}>
						<DrawPath d={handLine(460, 330, 1460, 330, 's18tl', 1.5)} p={prog(f, S40L2 + 4, S40L2 + 26)} width={7} />
						<DrawPath d="M 460 304 L 460 356" p={prog(f, S40L2, S40L2 + 6)} width={7} />
						<DrawPath d="M 1460 304 L 1460 356" p={prog(f, S40L2 + 24, S40L2 + 30)} width={7} />
					</g>
				</Svg>
				<Reveal at={S40L2 + 10} exitAt={S41 - 6} from="up" dist={16} style={{left: 760, width: 400, top: 236, textAlign: 'center'}}>
					<div style={{...T.label, fontSize: 44}}>{B.years}년</div>
				</Reveal>
				<Reveal at={S40L2 + 6} exitAt={S41 - 6} from="up" style={{left: 0, right: 0, top: 410, textAlign: 'center'}}>
					<Counter from={0} to={B.stocksThousand} at={S40L2 + 6} format={formatThousandStocks} color={C.ink} style={{fontSize: 160}} />
				</Reveal>

				{/* 자막 41–44: 와플 */}
				{f >= S41 ? (
					<>
						<Waffle
							{...G}
							cellStyle={(i) => {
								const row = Math.floor(i / 10);
								const blueIdx = BLUE_ORDER.indexOf(i);
								const blue = blueIdx >= 0 ? lin(f, fillAt + blueIdx * 0.5, fillAt + blueIdx * 0.5 + 4) : 0;
								const ti = TARGETS.indexOf(i);
								const gold = ti >= 0 ? lin(f, S43L2 + 6 + ti * 4, S43L2 + 14 + ti * 4) : 0;
								const rowDim = row === 0 ? 1 : 1 - 0.7 * rowFocus;
								if (gold > 0) {
									const found = f >= foundAt(ti) ? 1 : 0;
									return {fill: C.yellow, ring: found, stroke: C.ink, strokeWidth: 3, opacity: waffleIn};
								}
								const base = blue > 0 ? C.blue : EMPTY;
								return {fill: base, opacity: waffleIn * rowDim * (1 - 0.72 * dimAll) * (blue > 0 ? 0.35 + 0.65 * blue : 1)};
							}}
						/>
						<Svg>
							<DrawPath d={roundRect(G.x - 10, G.y - 10, 10 * G.cell + 9 * G.gap + 20, G.cell + 20, 12)} p={prog(f, S42 + 4, S42 + 18)} opacity={1 - prog(f, S43, S43 + 8)} width={6} />
							{cursorOn ? (
								<rect x={prev[0] + (cur[0] - prev[0]) * tw - 8} y={prev[1] + (cur[1] - prev[1]) * tw - 8} width={G.cell + 16} height={G.cell + 16} rx={10} fill="none" stroke={C.ink} strokeWidth={6} />
							) : null}
						</Svg>
					</>
				) : null}

				{/* 오른쪽 텍스트: 자막마다 하나씩 교체 */}
				<Reveal at={S41 + 10} exitAt={S42 - 2} from="right" style={{left: 920, top: 190}}>
					<Counter from={0} to={B.belowTbillPct} at={fillAt} dur={29} format={(v) => formatPct(v, 0, false)} color={C.blue} style={{fontSize: 200, display: 'block'}} />
					<div style={{...T.label, fontSize: 44, marginTop: 34, display: 'flex', alignItems: 'center', gap: 16}}>
						<span style={{display: 'inline-block', width: 34, height: 34, borderRadius: 6, background: C.blue}} />
						국채보다 못함
					</div>
				</Reveal>
				<Reveal at={S42 + 4} exitAt={S43 - 2} from="right" style={{left: 920, top: 140}}>
					<div style={{fontFamily: SERIF, fontWeight: 900, fontSize: 88, color: C.ink, lineHeight: 1.2}}>{B.sixOfTen}</div>
				</Reveal>
				<Reveal at={S43L2 + 10} exitAt={S44 - 2} from="right" style={{left: 920, top: 230}}>
					<div style={{display: 'flex', alignItems: 'baseline', gap: 24}}>
						<div style={{...T.headline, fontSize: 80}}>상위</div>
						<div style={{...T.number, fontSize: 200, color: C.ink}}>{formatPct(B.topPct, 0, false)}</div>
					</div>
				</Reveal>
				<Reveal at={S44 + 4} from="right" style={{left: 920, top: 250}}>
					<div style={{...T.headline, fontSize: 84, lineHeight: 1.3}}>
						{formatPct(B.topPct, 0, false)}를 맞혀야
						<br />
						하는 게임
					</div>
				</Reveal>
			</Layer>
			<SourceCaption text={SRC.bessembinder} at={S40 + 6} x={1824} align="right" />
		</AbsoluteFill>
	);
};
