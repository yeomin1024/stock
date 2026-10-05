// VERSION: v1.0.0 — 2026-10-05 — S11 (자막 16) 말풍선 3개가 가운데로 모여 합쳐지며 큰 세리프 "확신" + 노랑 하이라이트
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {C} from '../design/colors';
import {SERIF} from '../design/fonts';
import {easeInOut, enterP, lerp, prog} from '../design/motion';
import {Svg} from '../components/Draw';
import {Highlight} from '../components/Highlight';
import {cloudPath} from '../components/Icons';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {BUBBLES, Person, ThoughtBubble} from './S10';

const CENTER = {x: 960, y: 380};
const GATHER = 20;

export const S11: React.FC = () => {
	const f = useSceneFrame();
	const g = easeInOut(prog(f, 0, GATHER, (x) => x));
	const merged = enterP(f, GATHER - 8, 14);
	const mergedOut = 1 - prog(f, GATHER + 12, GATHER + 24);
	const word = enterP(f, GATHER + 10, 16);
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer depth="bg">
				<Person p={1} opacity={1 - prog(f, 0, 14)} />
			</Layer>
			<Layer depth="mid">
				{BUBBLES.map((b) => (
					<ThoughtBubble
						key={b.seed}
						b={b}
						p={1}
						dotsP={1 - g}
						scale={lerp(1, 0.4, g)}
						dx={(CENTER.x - b.cx) * g}
						dy={(CENTER.y - b.cy) * g}
						textOpacity={1 - prog(f, 0, 8)}
						opacity={1 - prog(f, GATHER - 4, GATHER + 2)}
					/>
				))}
				{merged > 0.001 && mergedOut > 0.001 ? (
					<Svg>
						<g
							opacity={Math.min(merged, mergedOut)}
							style={{transformBox: 'view-box', transformOrigin: `${CENTER.x}px ${CENTER.y}px`, scale: String(0.55 + 0.45 * merged)}}
						>
							<path d={cloudPath(CENTER.x, CENTER.y, 420, 190, 13, 5)} fill={C.white} stroke={C.ink} strokeWidth={6} strokeLinejoin="round" />
						</g>
					</Svg>
				) : null}
			</Layer>
			<Layer depth="fg">
				{word > 0.001 ? (
					<div
						style={{
							position: 'absolute',
							left: 0,
							right: 0,
							top: CENTER.y - 150,
							textAlign: 'center',
							fontFamily: SERIF,
							fontWeight: 900,
							fontSize: 240,
							lineHeight: 1.2,
							color: C.ink,
							opacity: word,
							scale: String(0.9 + 0.1 * word),
						}}
					>
						<Highlight at={GATHER + 24}>확신</Highlight>
					</div>
				) : null}
			</Layer>
		</AbsoluteFill>
	);
};
