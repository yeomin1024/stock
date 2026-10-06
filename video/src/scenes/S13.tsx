// VERSION: v2.0.0 — 2026-10-06 — S13 (자막 21–22) 장기 상승 점선 추세선 위로 주가선 → 번개와 함께 급락 (개념도)
// 자막 22: "최근 매수"·"기존 수익" 표시가 둘 다 손실 구간으로 떨어진다. 텍스트 덩어리 3개 이하를 위해 이때 "장기 상승" 라벨은 퇴장.
// 연결 근거: 장기 상승해도 단기 악재가 온다, 누구나 손해
import React from 'react';
import {AbsoluteFill, random} from 'remotion';
import {sceneTimes} from '../data/timeline';
import {C, mix} from '../design/colors';
import {SANS} from '../design/fonts';
import {easeIn, enterP, lerp, prog} from '../design/motion';
import {T} from '../design/type';
import {ConceptTag} from '../components/Bits';
import {DrawPath, Svg} from '../components/Draw';
import {Lightning} from '../components/Icons';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {handLine, smoothPath, Pt} from '../components/hand';

const t = sceneTimes('S13');
const S21 = t.sub(21);
const S22 = t.sub(22);
const BOLT_AT = S21 + 70;
const FALL_AT = S22 + 30;

const X0 = 220;
const X1 = 1700;
const trendY = (x: number) => 700 - ((x - X0) * 440) / (X1 - X0);
const PEAK_X = 1380;
const RISE: Pt[] = [];
for (let x = X0; x <= PEAK_X; x += 46) {
	const n = Math.sin(x / 61) * 0.6 + (random(`s13n${x}`) - 0.5) * 1.1;
	RISE.push([x, trendY(x) - 38 + n * 30]);
}
RISE.push([PEAK_X, trendY(PEAK_X) - 64]);
const CRASH: Pt[] = [[PEAK_X, trendY(PEAK_X) - 64], [1408, 440], [1436, 650], [1470, 700], [1520, 680], [1572, 710], [1640, 698]];
const LOW_Y = 702;

const yAt = (x: number): number => {
	for (let i = 0; i < RISE.length - 1; i++) {
		const [ax, ay] = RISE[i];
		const [bx, by] = RISE[i + 1];
		if (x >= ax && x <= bx) return lerp(ay, by, (x - ax) / (bx - ax));
	}
	return RISE[RISE.length - 1][1];
};

const MARKERS = [
	{label: '최근 매수', x: 1290, at: S22, fall: FALL_AT},
	{label: '기존 수익', x: 560, at: S22 + 6, fall: FALL_AT + 4},
].map((m) => ({...m, y: yAt(m.x)}));

export const S13: React.FC = () => {
	const f = useSceneFrame();
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer>
				<Svg>
					<DrawPath d={handLine(180, 130, 180, 760, 's13ay', 1)} p={prog(f, S21, S21 + 12)} width={4} />
					<DrawPath d={handLine(180, 760, 1760, 760, 's13ax', 1)} p={prog(f, S21 + 4, S21 + 18)} width={4} />
					<DrawPath d={`M ${X0} ${trendY(X0)} L ${X1} ${trendY(X1)}`} p={prog(f, S21 + 6, S21 + 30)} dash="18 14" width={5} stroke={C.gray} linecap="butt" />
					<DrawPath d={smoothPath(RISE)} p={prog(f, S21 + 10, BOLT_AT - 4)} width={7} stroke={C.red} />
					<DrawPath d={smoothPath(CRASH)} p={prog(f, BOLT_AT + 6, BOLT_AT + 22)} width={8} stroke={C.blue} />
					<Lightning x={1405} y={124} h={180} reveal={prog(f, BOLT_AT, BOLT_AT + 5, (x) => x)} />
					{MARKERS.map((m) => {
						const appear = enterP(f, m.at, 12);
						const fall = easeIn(prog(f, m.fall, m.fall + 18, (x) => x));
						const y = lerp(m.y, LOW_Y, fall);
						return appear > 0.001 ? (
							<g key={m.label}>
								{fall > 0 ? <line x1={m.x} y1={m.y} x2={m.x} y2={y} stroke={C.blue} strokeWidth={4} strokeDasharray="8 8" /> : null}
								<circle cx={m.x} cy={m.y} r={9} fill="none" stroke={C.gray} strokeWidth={3} opacity={fall} />
								<circle cx={m.x} cy={y} r={14 * appear} fill={mix(C.ink, C.blue, fall)} stroke={C.paper} strokeWidth={4} />
							</g>
						) : null;
					})}
				</Svg>
				{MARKERS.map((m) => {
					const fall = easeIn(prog(f, m.fall, m.fall + 18, (x) => x));
					const col = mix(C.ink, C.blue, fall);
					return (
						<Reveal key={m.label} at={m.at} from="down" dist={16} style={{left: m.x - 110, top: lerp(m.y, LOW_Y, fall) - 96, width: 220, textAlign: 'center'}}>
							<div
								style={{
									display: 'inline-block',
									fontFamily: SANS,
									fontWeight: 700,
									fontSize: 36,
									color: col,
									background: C.paper,
									border: `3px solid ${col}`,
									borderRadius: 30,
									padding: '2px 20px 4px',
									whiteSpace: 'nowrap',
								}}
							>
								{m.label}
							</div>
						</Reveal>
					);
				})}
				<Reveal at={S21 + 14} exitAt={S22} from="down" dist={16} style={{left: 960, top: trendY(980) + 24}}>
					<div style={{...T.label, color: C.gray}}>장기 상승</div>
				</Reveal>
				<Reveal at={BOLT_AT + 4} from="left" dist={20} style={{left: 1480, top: 168}}>
					<div style={{...T.label, fontSize: 42}}>단기 악재</div>
				</Reveal>
				<ConceptTag at={S21 + 8} x={1650} y={772 - 2} />
			</Layer>
		</AbsoluteFill>
	);
};
