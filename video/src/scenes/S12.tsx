// VERSION: v2.0.0 — 2026-10-06 — S12 (자막 19–20) "01" + "하락하면 대처할 방법이 없다"
// 자막 20: 앞으로 뻗은 점선 길이 위로 갈수록 연하게 사라진다 (그라데이션, blur 없음) + "앞을 알 수 없다"
// 연결 근거: 주식은 앞일을 모른다
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {sceneTimes} from '../data/timeline';
import {C} from '../design/colors';
import {enterP, prog} from '../design/motion';
import {T} from '../design/type';
import {NumberTitle} from '../components/Bits';
import {DrawPath, Svg, useSafeId} from '../components/Draw';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {smoothPath} from '../components/hand';

const t = sceneTimes('S12');
const S19 = t.sub(19);
const S20 = t.sub(20);
const PATH = smoothPath([
	[880, 790],
	[990, 700],
	[1110, 610],
	[1240, 512],
	[1370, 418],
	[1490, 334],
	[1600, 262],
]);

export const S12: React.FC = () => {
	const f = useSceneFrame();
	const gradId = useSafeId('road-fade');
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer>
				<NumberTitle num="01" title="하락하면 대처할 방법이 없다" mark="대처할 방법이 없다" at={S19} markAt={S19 + 24} />
				<Svg>
					<defs>
						<linearGradient id={gradId} gradientUnits="userSpaceOnUse" x1={0} y1={790} x2={0} y2={262}>
							<stop offset="0" stopColor={C.ink} stopOpacity={1} />
							<stop offset="0.55" stopColor={C.ink} stopOpacity={0.45} />
							<stop offset="1" stopColor={C.ink} stopOpacity={0} />
						</linearGradient>
					</defs>
					<circle cx={880} cy={790 - 2} r={14 * enterP(f, S20, 10)} fill={C.ink} />
					<DrawPath d={PATH} p={prog(f, S20 + 4, S20 + 54)} width={10} dash="2 24" stroke={`url(#${gradId})`} />
				</Svg>
				<Reveal at={S20 + 30} from="none" style={{left: 1470, top: 168}}>
					<div style={{...T.label, color: C.gray, fontSize: 44}}>앞을 알 수 없다</div>
				</Reveal>
			</Layer>
		</AbsoluteFill>
	);
};
