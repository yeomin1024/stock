// VERSION: v1.0.0 — 2026-10-05 — S15 (자막 22–23) 장기 상승 추세선(점선, 개념도) 위 주가선 → 번개와 함께 급락.
// 자막 23: "최근 매수"·"기존 수익" 마커가 둘 다 손실 구간으로 떨어진다.
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

const t = sceneTimes('S15');
const BOLT_AT = t.word(22, '악재') - 8;
const RECENT_AT = t.sub(23);
const OLD_AT = t.word(23, '기존에');
const FALL_AT = t.word(23, '큰 손해') - 10;

const X0 = 220;
const X1 = 1700;
const trendY = (x: number) => 690 - ((x - X0) * 440) / (X1 - X0);
const PEAK_X = 1380;
const RISE: Pt[] = [];
for (let x = X0; x <= PEAK_X; x += 46) {
	const n = Math.sin(x / 61) * 0.6 + (random(`s15n${x}`) - 0.5) * 1.1;
	RISE.push([x, trendY(x) - 38 + n * 30]);
}
RISE.push([PEAK_X, trendY(PEAK_X) - 64]);
const CRASH: Pt[] = [
	[PEAK_X, trendY(PEAK_X) - 64],
	[1408, 430],
	[1436, 640],
	[1470, 690],
	[1520, 668],
	[1572, 700],
	[1640, 688],
];
const LOW_Y = 692;

// 가격선 위의 y 를 x 로 찾기 (선형 보간)
const yAt = (x: number): number => {
	for (let i = 0; i < RISE.length - 1; i++) {
		const [ax, ay] = RISE[i];
		const [bx, by] = RISE[i + 1];
		if (x >= ax && x <= bx) return lerp(ay, by, (x - ax) / (bx - ax));
	}
	return RISE[RISE.length - 1][1];
};

const MARKERS = [
	{label: '기존 수익', x: 560, at: OLD_AT, fall: FALL_AT + 4},
	{label: '최근 매수', x: 1290, at: RECENT_AT, fall: FALL_AT},
].map((m) => ({...m, y: yAt(m.x)}));

export const S15: React.FC = () => {
	const f = useSceneFrame();
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer depth="bg">
				<Svg>
					<DrawPath d={handLine(180, 120, 180, 750, 's15ay', 1)} p={prog(f, 0, 14)} width={4} />
					<DrawPath d={handLine(180, 750, 1760, 750, 's15ax', 1)} p={prog(f, 4, 20)} width={4} />
				</Svg>
				<ConceptTag at={8} x={1650} y={112} />
			</Layer>
			<Layer depth="mid">
				<Svg>
					<DrawPath d={`M ${X0} ${trendY(X0)} L ${X1} ${trendY(X1)}`} p={prog(f, 4, 34)} dash="18 14" width={5} stroke={C.gray} linecap="butt" />
					<DrawPath d={smoothPath(RISE)} p={prog(f, 12, BOLT_AT - 6)} width={7} stroke={C.red} />
					<DrawPath d={smoothPath(CRASH)} p={prog(f, BOLT_AT + 6, BOLT_AT + 22)} width={8} stroke={C.blue} />
					<Lightning x={1405} y={104} h={190} reveal={prog(f, BOLT_AT, BOLT_AT + 5, (x) => x)} />
					{/* 마커: 매수 지점 → 낙하 */}
					{MARKERS.map((m) => {
						const appear = enterP(f, m.at, 12);
						const fall = easeIn(prog(f, m.fall, m.fall + 18, (x) => x));
						const y = lerp(m.y, LOW_Y, fall);
						const col = mix(C.ink, C.blue, fall);
						return appear > 0.001 ? (
							<g key={m.label}>
								<DrawPath d={`M ${m.x} ${m.y} L ${m.x} ${y}`} p={fall > 0 ? 1 : 0} dash="8 8" width={4} stroke={C.blue} linecap="butt" />
								<circle cx={m.x} cy={m.y} r={9} fill="none" stroke={C.gray} strokeWidth={3} opacity={fall} />
								<circle cx={m.x} cy={y} r={14 * appear} fill={col} stroke={C.paper} strokeWidth={4} />
							</g>
						) : null;
					})}
				</Svg>
				{MARKERS.map((m) => {
					const fall = easeIn(prog(f, m.fall, m.fall + 18, (x) => x));
					const y = lerp(m.y, LOW_Y, fall);
					const col = mix(C.ink, C.blue, fall);
					return (
						<Reveal key={m.label} at={m.at} from="down" dist={16} style={{left: m.x - 100, top: y - 92, width: 200, textAlign: 'center'}}>
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
									padding: '4px 20px 6px',
									whiteSpace: 'nowrap',
								}}
							>
								{m.label}
							</div>
						</Reveal>
					);
				})}
				<Reveal at={18} from="down" dist={16} style={{left: 960, top: trendY(980) + 24}}>
					<div style={{...T.label, color: C.gray, fontSize: 40}}>장기 상승</div>
				</Reveal>
				<Reveal at={BOLT_AT + 4} from="left" dist={20} style={{left: 1490, top: 150}}>
					<div style={{...T.label, fontSize: 42}}>단기 악재</div>
				</Reveal>
			</Layer>
		</AbsoluteFill>
	);
};
