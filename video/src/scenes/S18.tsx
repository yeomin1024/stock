// VERSION: v1.0.0 — 2026-10-05 — S18 (자막 30) 와플 차트에서 4칸만 노랑으로 빛나고 나머지는 흐려짐 → 상위 4%
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {FACTS, SRC, formatPct} from '../data/facts';
import {C} from '../design/colors';
import {SANS} from '../design/fonts';
import {exitP, lin, prog} from '../design/motion';
import {T} from '../design/type';
import {SourceCaption} from '../components/Bits';
import {Highlight} from '../components/Highlight';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {Waffle} from '../components/Waffle';
import {EMPTY_CELL, WAFFLE} from './S17';

const B = FACTS.bessembinder;
// 장면 시작 = 자막 30 시작
const GLOW_AT = 12;
// 상위 4% = 국채 미달(파랑) 칸이 아닌 마지막 줄 오른쪽 4칸
const TOP = Array.from({length: B.topPct}, (_, i) => 100 - B.topPct + i);

export const S18: React.FC = () => {
	const f = useSceneFrame();
	const dim = prog(f, 4, 16);
	const prevOut = exitP(f, 0, 9);
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer depth="mid">
				<Waffle
					{...WAFFLE}
					cellStyle={(i) => {
						const k = TOP.indexOf(i);
						if (k >= 0) {
							const g = lin(f, GLOW_AT + k * 4, GLOW_AT + k * 4 + 8);
							return {fill: g > 0 ? C.yellow : EMPTY_CELL, glow: g, scale: 1 + 0.12 * g * (1 - lin(f, GLOW_AT + k * 4 + 8, GLOW_AT + k * 4 + 16)), stroke: g > 0 ? C.ink : undefined, strokeWidth: 3 * g};
						}
						const base = i < B.belowTbillPct ? C.blue : EMPTY_CELL;
						return {fill: base, opacity: 1 - 0.75 * dim};
					}}
				/>
				{/* S17 에서 이어진 "58% 국채보다 못함" 은 퇴장 */}
				{prevOut > 0.001 ? (
					<>
						<div style={{position: 'absolute', left: 900, top: 150, opacity: prevOut, ...T.number, fontSize: 240, color: C.blue}}>
							{formatPct(B.belowTbillPct, 0, false)}
						</div>
						<div style={{position: 'absolute', left: 912, top: 430, opacity: prevOut, display: 'flex', alignItems: 'center', gap: 18}}>
							<div style={{width: 36, height: 36, borderRadius: 6, background: C.blue}} />
							<div style={{fontFamily: SANS, fontWeight: 700, fontSize: 44, color: C.ink, whiteSpace: 'nowrap'}}>국채보다 못함</div>
						</div>
					</>
				) : null}
				<Reveal at={GLOW_AT} from="right" style={{left: 908, top: 170}}>
					<div style={{...T.label, fontSize: 44, color: C.ink}}>초과 수익을 만든 종목</div>
				</Reveal>
				<Reveal at={GLOW_AT + 4} from="right" style={{left: 900, top: 240}}>
					<div style={{display: 'flex', alignItems: 'baseline', gap: 26}}>
						<div style={{...T.headline, fontSize: 96}}>상위</div>
						<div style={{...T.number, fontSize: 240, color: C.ink}}>
							<Highlight at={GLOW_AT + 14} top={0.5}>
								{formatPct(B.topPct, 0, false)}
							</Highlight>
						</div>
					</div>
				</Reveal>
			</Layer>
			<SourceCaption text={SRC.bessembinder} at={-30} />
		</AbsoluteFill>
	);
};
