// VERSION: v2.0.0 — 2026-10-06 — S07 (자막 9) "고점"에서 완만히 내려오던 선이 마지막에 거의 수직으로 떨어짐(개념도) → 떨어진 지점에 최대 -26%(파랑)
// 연결 근거: 고점에서 내려오던 주가가 소식에 폭락
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {FACTS, SRC, formatPct} from '../data/facts';
import {sceneTimes} from '../data/timeline';
import {C, alpha} from '../design/colors';
import {enterP, prog} from '../design/motion';
import {T} from '../design/type';
import {ConceptTag, SourceCaption} from '../components/Bits';
import {Counter} from '../components/Counter';
import {DrawPath, Svg} from '../components/Draw';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {smoothPath, Pt} from '../components/hand';

const t = sceneTimes('S07');
const S9 = t.sub(9);
const DROP_AT = S9 + 40;
const PEAK: Pt = [250, 250];
const SLOPE: Pt[] = [PEAK, [420, 282], [560, 300], [700, 330], [840, 352], [980, 388]];
const DROP: Pt[] = [[980, 388], [1010, 470], [1030, 640], [1046, 712]];
const LOW = DROP[DROP.length - 1];

export const S07: React.FC = () => {
	const f = useSceneFrame();
	return (
		<AbsoluteFill>
			<SceneBg tone="navy" />
			<Layer>
				<Svg>
					<circle cx={PEAK[0]} cy={PEAK[1]} r={12 * enterP(f, S9, 10)} fill={C.light} />
					<DrawPath d={smoothPath(SLOPE)} p={prog(f, S9 + 6, DROP_AT)} stroke={C.blue} width={7} />
					<DrawPath d={smoothPath(DROP)} p={prog(f, DROP_AT, DROP_AT + 10)} stroke={C.blue} width={10} />
					<circle cx={LOW[0]} cy={LOW[1]} r={14 * enterP(f, DROP_AT + 8, 8)} fill={C.blue} stroke={C.navy} strokeWidth={4} />
				</Svg>
				<Reveal at={S9 + 2} from="down" dist={16} style={{left: PEAK[0] - 40, top: PEAK[1] - 76}}>
					<div style={{...T.label, color: C.light}}>고점</div>
				</Reveal>
				<Reveal at={DROP_AT + 6} from="left" dist={30} style={{left: LOW[0] + 70, top: 410}}>
					<div style={{...T.label, color: C.light, fontSize: 44}}>최대</div>
					<Counter from={0} to={FACTS.story.maxDropPct} at={DROP_AT + 6} format={(v) => formatPct(v)} color={C.blue} style={{fontSize: 210, display: 'block', marginTop: 6}} />
				</Reveal>
				<ConceptTag at={S9 + 6} x={1690} y={136} color={alpha(C.light, 0.6)} />
			</Layer>
			<SourceCaption text={SRC.press} at={DROP_AT + 6} color={alpha(C.light, 0.65)} />
		</AbsoluteFill>
	);
};
