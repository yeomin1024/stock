// VERSION: v1.0.0 — 2026-10-05 — S20 (자막 33) 100칸 위를 커서가 빠르게 헤매며 숨은 4칸을 하나씩 찾는다 → "4%를 맞혀야 하는 게임"
import React from 'react';
import {AbsoluteFill, random} from 'remotion';
import {FACTS} from '../data/facts';
import {C, alpha} from '../design/colors';
import {easeOut, lin, prog} from '../design/motion';
import {T} from '../design/type';
import {Svg} from '../components/Draw';
import {Highlight} from '../components/Highlight';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {Waffle, waffleCellXY} from '../components/Waffle';

const G = {x: 180, y: 110, cell: 56, gap: 6};
const TARGETS = [17, 42, 68, 93]; // 4칸 = 4%
const HOP_START = 6;
const HOP = 3;
const HITS = [12, 24, 35, 46]; // 몇 번째 이동에서 찾는지
const HOPS = HITS[HITS.length - 1] + 1;

// 결정적 커서 경로: 지정된 이동 번호에서만 목표 칸, 나머지는 목표가 아닌 칸
const PATH: number[] = Array.from({length: HOPS}, (_, k) => {
	const hit = HITS.indexOf(k);
	if (hit >= 0) return TARGETS[hit];
	let c = Math.floor(random(`s20c${k}`) * 100);
	while (TARGETS.includes(c)) c = (c + 7) % 100;
	return c;
});
const foundAt = (ti: number) => HOP_START + HITS[ti] * HOP;

export const S20: React.FC = () => {
	const f = useSceneFrame();
	const appear = prog(f, 0, 10);
	const k = Math.min(HOPS - 1, Math.max(0, Math.floor((f - HOP_START) / HOP)));
	const sub = (f - HOP_START - k * HOP) / HOP;
	const cur = waffleCellXY(PATH[k], G.x, G.y, G.cell, G.gap);
	const prev = waffleCellXY(PATH[Math.max(0, k - 1)], G.x, G.y, G.cell, G.gap);
	const tw = easeOut(Math.min(1, Math.max(0, sub * 1.6)));
	const cx = prev[0] + (cur[0] - prev[0]) * tw;
	const cy = prev[1] + (cur[1] - prev[1]) * tw;
	const cursorOn = f >= HOP_START && f < foundAt(TARGETS.length - 1) + 10;
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer depth="mid">
				<Waffle
					{...G}
					cellStyle={(i) => {
						const ti = TARGETS.indexOf(i);
						if (ti >= 0 && f >= foundAt(ti)) {
							const g = lin(f, foundAt(ti), foundAt(ti) + 6);
							return {fill: C.yellow, glow: g, stroke: C.ink, strokeWidth: 3, scale: 1 + 0.12 * (1 - lin(f, foundAt(ti) + 4, foundAt(ti) + 12))};
						}
						// 지나간 칸은 잠깐 회색으로 깜빡임 (빗나감)
						const visits = PATH.map((c, h) => (c === i ? HOP_START + h * HOP : -1)).filter((x) => x >= 0 && x <= f);
						const last = visits.length ? visits[visits.length - 1] : -100;
						const miss = 1 - lin(f, last + 1, last + 9);
						return {fill: miss > 0.01 ? alpha(C.gray, 0.25 + 0.5 * miss) : alpha(C.ink, 0.09), opacity: appear};
					}}
				/>
				{cursorOn ? (
					<Svg>
						<rect x={cx - 7} y={cy - 7} width={G.cell + 14} height={G.cell + 14} rx={10} fill="none" stroke={C.ink} strokeWidth={6} />
					</Svg>
				) : null}
				<Reveal at={0} from="right" style={{left: 900, top: 230}}>
					<div style={{...T.headline, fontSize: 96, lineHeight: 1.3}}>
						<Highlight at={10}>{FACTS.bessembinder.topPct}%</Highlight>를 맞혀야
						<br />
						하는 게임
					</div>
				</Reveal>
			</Layer>
		</AbsoluteFill>
	);
};
