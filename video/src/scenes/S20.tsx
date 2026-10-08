// VERSION: v2.3.0 — 2026-10-08 — 한 줄 자막: 뒷줄 내용 요소는 t.line(n, 2) · S20 (자막 48) 점 수십 개 중 몇 개만 노랑으로 빛나며 위로 떠오르고(성공담),
// 나머지 회색 점은 소리 없이 아래로 흐려져 사라진다. 연결 근거: 생존자 편향
import React from 'react';
import {AbsoluteFill, random} from 'remotion';
import {sceneTimes} from '../data/timeline';
import {C, mix} from '../design/colors';
import {easeInOut, enterP, lerp, prog} from '../design/motion';
import {T} from '../design/type';
import {Svg} from '../components/Draw';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';

const t = sceneTimes('S20');
const S48 = t.sub(48);
const SPLIT_AT = S48 + 30;
const LEAVE_AT = t.line(48, 2); // 뒷줄 "크게 잃은 사람들은 조용히 떠나기 때문이죠."
const N = 54;
const WINNERS = [7, 19, 30, 41, 48];
const DOTS = Array.from({length: N}, (_, i) => ({
	x: 560 + random(`s20x${i}`) * 800,
	y: 330 + random(`s20y${i}`) * 260,
	r: 11 + random(`s20r${i}`) * 5,
	d: random(`s20d${i}`) * 20,
}));

export const S20: React.FC = () => {
	const f = useSceneFrame();
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer>
				<Reveal at={S48} from="left" style={{left: 104, top: 128}}>
					<div style={{...T.headline, fontSize: 80}}>성공담만 들리는 이유</div>
				</Reveal>
				<Svg>
					{DOTS.map((d, i) => {
						const appear = enterP(f, S48 + Math.floor(i / 6), 12);
						const win = WINNERS.indexOf(i);
						if (win >= 0) {
							const up = easeInOut(prog(f, SPLIT_AT + win * 4, SPLIT_AT + 40 + win * 4, (x) => x));
							const gold = prog(f, SPLIT_AT + win * 4, SPLIT_AT + 10 + win * 4);
							return (
								<circle
									key={i}
									cx={lerp(d.x, 820 + win * 70, up)}
									cy={lerp(d.y, 268, up)}
									r={d.r * appear * (1 + 0.5 * gold)}
									fill={mix(C.gray, C.yellow, gold)}
									stroke={gold > 0.5 ? C.ink : 'none'}
									strokeWidth={3}
								/>
							);
						}
						const fall = prog(f, LEAVE_AT + d.d, LEAVE_AT + 78 + d.d);
						return <circle key={i} cx={d.x} cy={d.y + 150 * fall} r={d.r * appear} fill={C.gray} opacity={1 - fall} />;
					})}
				</Svg>
				<Reveal at={LEAVE_AT + 10} from="none" style={{left: 0, right: 0, top: 690, textAlign: 'center'}}>
					<div style={{...T.label, color: C.gray, fontSize: 44}}>크게 잃은 사람은 조용히 떠난다</div>
				</Reveal>
			</Layer>
		</AbsoluteFill>
	);
};
