// VERSION: v1.0.0 — 2026-10-05 — S12 (자막 17–19) 흩어진 점 100개 대부분이 한 칸으로 몰림(개념도) → 큰 물음표 "큰 수익?" → "문제 분석"
import React from 'react';
import {AbsoluteFill, random} from 'remotion';
import {sceneTimes} from '../data/timeline';
import {C, alpha} from '../design/colors';
import {SERIF} from '../design/fonts';
import {easeInOut, enterP, exitP, lerp, prog} from '../design/motion';
import {T} from '../design/type';
import {ConceptTag} from '../components/Bits';
import {DrawPath, Svg} from '../components/Draw';
import {Highlight} from '../components/Highlight';
import {questionPath} from '../components/Icons';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {handLine, roundRect} from '../components/hand';

const t = sceneTimes('S12');
const RUSH_AT = t.word(17, '전 재산') - 16;
const Q_AT = t.sub(18);
const TITLE_AT = t.sub(19);

const BOX_W = 220;
const BOX_H = 150;
const BOX_Y = 560;
const BOXES = [310, 580, 850, 1120, 1390];
const MAIN = 2;
const N = 100;
const OTHERS = 12; // 나머지 칸으로 가는 점 (각 3개)

const R = 7;
const dots = Array.from({length: N}, (_, i) => {
	const from: [number, number] = [200 + random(`d12x${i}`) * 1520, 120 + random(`d12y${i}`) * 350];
	let to: [number, number];
	if (i < N - OTHERS) {
		const col = i % 11;
		const row = Math.floor(i / 11);
		to = [BOXES[MAIN] + 22 + col * 17.6, BOX_Y + 24 + row * 13.5];
	} else {
		const k = i - (N - OTHERS);
		const box = [0, 1, 3, 4][k % 4];
		const slot = Math.floor(k / 4);
		to = [BOXES[box] + 80 + slot * 30, BOX_Y + BOX_H - 30];
	}
	return {from, to, delay: random(`d12d${i}`) * 26};
});

export const S12: React.FC = () => {
	const f = useSceneFrame();
	const dim = 1 - prog(f, Q_AT, Q_AT + 12) * 0.78;
	const out1 = exitP(f, TITLE_AT, 9);
	const filled = dots.filter((d, i) => i < N - OTHERS && f >= RUSH_AT + d.delay + 20).length;
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer depth="mid">
				<Svg>
					<g opacity={dim * out1}>
						{BOXES.map((x, i) => (
							<React.Fragment key={x}>
								<rect
									x={x}
									y={BOX_Y}
									width={BOX_W}
									height={BOX_H}
									rx={16}
									fill={i === MAIN ? alpha(C.yellow, 0.25 + 0.5 * (filled / (N - OTHERS))) : C.paper}
									opacity={prog(f, 6 + i * 4, 18 + i * 4)}
								/>
								<DrawPath d={roundRect(x, BOX_Y, BOX_W, BOX_H, 16)} p={prog(f, 6 + i * 4, 22 + i * 4)} width={i === MAIN ? 6 : 4} />
							</React.Fragment>
						))}
						{dots.map((d, i) => {
							const a = RUSH_AT + d.delay;
							const p = easeInOut(prog(f, a, a + 20, (x) => x));
							const pop = enterP(f, 2 + i * 0.25, 10);
							return (
								<circle
									key={i}
									cx={lerp(d.from[0], d.to[0], p)}
									cy={lerp(d.from[1], d.to[1], p) - Math.sin(Math.PI * p) * 50}
									r={R * pop}
									fill={C.ink}
								/>
							);
						})}
					</g>
				</Svg>
				<ConceptTag at={10} exitAt={TITLE_AT} x={1650} y={BOX_Y + BOX_H + 18} />
			</Layer>
			<Layer depth="fg">
				{/* 자막 18: 큰 물음표 */}
				<Svg>
					<g opacity={out1}>
						<DrawPath d={questionPath(1400, 150, 360).hook} p={prog(f, Q_AT + 4, Q_AT + 22)} width={30} />
						<circle cx={1400} cy={150 + 0.94 * 360} r={22 * enterP(f, Q_AT + 20, 8)} fill={C.ink} />
					</g>
				</Svg>
				<Reveal at={Q_AT + 8} exitAt={TITLE_AT} from="left" style={{left: 300, top: 250}}>
					<div style={{...T.headline, fontSize: 96}}>
						<Highlight at={Q_AT + 16}>큰 수익</Highlight>?
					</div>
				</Reveal>
				{/* 자막 19: 문제 분석 타이틀 */}
				<Reveal at={TITLE_AT + 8} from="up" style={{left: 0, right: 0, top: 250, textAlign: 'center'}}>
					<div style={{fontFamily: SERIF, fontWeight: 900, fontSize: 96, color: C.ink, lineHeight: 1.2}}>문제 분석</div>
				</Reveal>
				<Svg>
					<DrawPath d={handLine(730, 388, 1190, 382, 's12ul', 3)} p={prog(f, TITLE_AT + 18, TITLE_AT + 30)} width={10} />
				</Svg>
			</Layer>
		</AbsoluteFill>
	);
};
