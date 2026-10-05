// VERSION: v1.0.0 — 2026-10-05 — S19 (자막 31–32) 10칸 중 4칸 이상이 파랑: 고점에서 -70% 떨어진 뒤 회복 못 한 미니 라인
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {FACTS, SRC, formatPct} from '../data/facts';
import {sceneTimes} from '../data/timeline';
import {C, mix} from '../design/colors';
import {SANS} from '../design/fonts';
import {enterP, lin, prog} from '../design/motion';
import {T} from '../design/type';
import {ConceptTag, SourceCaption} from '../components/Bits';
import {Counter} from '../components/Counter';
import {DrawPath, Svg} from '../components/Draw';
import {Highlight} from '../components/Highlight';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {roundRect, smoothPath, Pt} from '../components/hand';

const t = sceneTimes('S19');
const J = FACTS.jpm;
const HIT_AT = t.sub(32);
const DD_AT = t.word(32, '70%');
const NR_AT = t.word(32, '회복');

const N = 10;
const X0 = 166;
const CW = 148;
const GAP = 12;
const CY = 190;
const CH = 220;
// 10칸 중 4칸 = 40%
const LOSERS = [1, 3, 6, 8];

// 미니 라인 모양 (값 0~1). 실제 데이터가 아닌 모양 예시라 숫자 눈금 없음
const normalShape = (i: number): number[] => {
	const ph = i * 1.3;
	return Array.from({length: 9}, (_, k) => {
		const s = k / 8;
		return 0.22 + 0.55 * s + 0.12 * Math.sin(s * 9 + ph) - 0.1 * Math.exp(-((s - 0.55) ** 2) / 0.01);
	});
};
const PEAK = 1;
const loserShape = [0.25, 0.48, 0.72, PEAK, 0.62, PEAK * 0.3, 0.27, 0.31, 0.28];

const toPts = (vals: number[], cx: number): Pt[] =>
	vals.map((v, k) => [cx + 16 + (k / (vals.length - 1)) * (CW - 32), CY + CH - 26 - v * (CH - 60)] as Pt);

export const S19: React.FC = () => {
	const f = useSceneFrame();
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer depth="mid">
				<Reveal at={0} from="left" style={{left: X0, top: 96}}>
					<div style={{...T.label, fontSize: 44}}>
						<Highlight at={8}>{J.sinceYear}년 이후</Highlight>
					</div>
				</Reveal>
				<Svg>
					{Array.from({length: N}, (_, i) => {
						const cx = X0 + i * (CW + GAP);
						const isLoser = LOSERS.includes(i);
						const hitAt = HIT_AT + LOSERS.indexOf(i) * 6;
						const hit = isLoser ? lin(f, hitAt, hitAt + 8) : 0;
						const appear = enterP(f, i * 3, 12);
						const vals = isLoser ? loserShape : normalShape(i);
						const pts = toPts(vals, cx);
						const peakPt = pts[3];
						const stroke = mix(C.ink, C.blue, hit);
						return (
							<g key={i} opacity={appear}>
								<rect x={cx} y={CY} width={CW} height={CH} rx={14} fill={mix(C.paper, C.blue, 0.14 * hit)} />
								<path d={roundRect(cx, CY, CW, CH, 14)} fill="none" stroke={stroke} strokeWidth={3 + 2 * hit} />
								<DrawPath d={smoothPath(pts)} p={prog(f, 8 + i * 3, 34 + i * 3)} width={4 + 2 * hit} stroke={isLoser ? mix(C.gray, C.blue, hit) : C.gray} />
								{isLoser ? (
									<DrawPath
										d={`M ${peakPt[0]} ${peakPt[1]} L ${cx + CW - 12} ${peakPt[1]}`}
										p={prog(f, DD_AT + LOSERS.indexOf(i) * 4, DD_AT + 10 + LOSERS.indexOf(i) * 4)}
										dash="6 6"
										width={3}
										stroke={C.blue}
										linecap="butt"
									/>
								) : null}
							</g>
						);
					})}
				</Svg>
				<Reveal at={HIT_AT} from="up" style={{left: X0 - 6, top: 460}}>
					<div style={{display: 'flex', alignItems: 'baseline', gap: 18}}>
						<Counter from={0} to={J.noRecoverPct} at={HIT_AT} format={(v) => formatPct(v, 0, false)} color={C.blue} style={{fontSize: 200}} />
						<span style={{fontFamily: SANS, fontWeight: 900, fontSize: 64, color: C.blue}}>이상</span>
					</div>
				</Reveal>
				<Reveal at={DD_AT} from="right" style={{left: 1060, top: 486}}>
					<div style={{...T.label, fontSize: 44}}>
						고점 대비 <span style={{color: C.blue, fontWeight: 900, fontVariantNumeric: 'tabular-nums'}}>{formatPct(J.drawdownPct, 0)}</span>
					</div>
				</Reveal>
				<Reveal at={NR_AT} from="right" style={{left: 1060, top: 576}}>
					<div style={{...T.label, fontSize: 44}}>
						<Highlight at={NR_AT + 6}>회복 못 함</Highlight>
					</div>
				</Reveal>
			</Layer>
			<ConceptTag at={30} x={X0 + N * (CW + GAP) - GAP - 74} y={CY + CH + 12} />
			<SourceCaption text={SRC.jpm} at={6} x={X0} />
		</AbsoluteFill>
	);
};
