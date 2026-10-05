// VERSION: v1.0.0 — 2026-10-05 — S14 (자막 20–21) "01" + 헤드라인. 자막 21: 앞으로 뻗은 점선 길이 안개 그라데이션 속으로 사라짐
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {sceneTimes} from '../data/timeline';
import {C, alpha} from '../design/colors';
import {enterP, prog} from '../design/motion';
import {T} from '../design/type';
import {DrawPath, Svg, useSafeId} from '../components/Draw';
import {Highlight} from '../components/Highlight';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useCameraP, useSceneFrame} from '../components/Scene';
import {smoothPath} from '../components/hand';

const t = sceneTimes('S14');
const HL_AT = t.word(20, '대처할');
const ROAD_AT = t.sub(21);
const ROAD = smoothPath([
	[230, 730],
	[520, 700],
	[820, 640],
	[1100, 610],
	[1380, 560],
	[1640, 540],
	[1900, 520],
]);

export const S14: React.FC = () => {
	const f = useSceneFrame();
	const cam = useCameraP();
	const gradId = useSafeId('fog');
	const blurId = useSafeId('fog-blur');
	const fog = prog(f, ROAD_AT - 4, ROAD_AT + 20);
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer depth="mid">
				<Reveal at={0} from="up" style={{left: 150, top: 96}}>
					<div style={{...T.number, fontSize: 230, color: C.ink, letterSpacing: '-0.04em'}}>01</div>
				</Reveal>
				<Reveal at={6} from="left" style={{left: 160, top: 340}}>
					<div style={{...T.headline, fontSize: 92}}>
						하락하면 <Highlight at={HL_AT}>대처할 방법이 없다</Highlight>
					</div>
				</Reveal>
				<Svg>
					<circle cx={230} cy={730} r={16 * enterP(f, ROAD_AT, 10)} fill={C.ink} />
					<DrawPath d={ROAD} p={prog(f, ROAD_AT + 4, ROAD_AT + 64)} width={9} dash="2 22" linecap="round" />
				</Svg>
			</Layer>
			{/* 안개: 오른쪽으로 갈수록 크림색이 짙어져 길이 사라진다 (그래픽 영역 안) */}
			<Layer depth="fg">
				<Svg>
					<defs>
						<linearGradient id={gradId} x1="0" y1="0" x2="1" y2="0">
							<stop offset="0" stopColor={C.cream} stopOpacity={0} />
							<stop offset="0.55" stopColor={C.cream} stopOpacity={0.85} />
							<stop offset="1" stopColor={C.cream} stopOpacity={1} />
						</linearGradient>
						<filter id={blurId} x="-50%" y="-50%" width="200%" height="200%">
							<feGaussianBlur stdDeviation={28} />
						</filter>
					</defs>
					<g opacity={fog}>
						<rect x={900} y={470} width={1100} height={330} fill={`url(#${gradId})`} filter={`url(#${blurId})`} />
						{[
							[1250, 560, 260, 70],
							[1520, 600, 300, 90],
							[1760, 540, 260, 100],
						].map(([x, y, rx, ry], i) => (
							<ellipse key={i} cx={x - cam * (20 + i * 12)} cy={y} rx={rx} ry={ry} fill={alpha(C.paper, 0.9)} filter={`url(#${blurId})`} />
						))}
					</g>
				</Svg>
			</Layer>
		</AbsoluteFill>
	);
};
