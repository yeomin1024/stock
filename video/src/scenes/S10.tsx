// VERSION: v2.3.0 — 2026-10-08 — 한 줄 자막: 뒷줄 내용 요소는 t.line(n, 2) · S10 (자막 16–18) 작은 계좌 카드 12개가 깔리고 모두 한 종목으로 꽉 찬 비중 막대
// 자막 17: 가운데 큰 물음표 + "큰 수익?". 자막 18: "무엇이 문제일까?" 타이틀. 연결 근거: 이런 사람이 많다, 과연 수익을 낼까
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {sceneTimes} from '../data/timeline';
import {C, alpha} from '../design/colors';
import {SERIF} from '../design/fonts';
import {easeInOut, enterP, exitP, prog} from '../design/motion';
import {T} from '../design/type';
import {DrawPath, Svg} from '../components/Draw';
import {Highlight} from '../components/Highlight';
import {questionPath} from '../components/Icons';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {handLine, roundRect} from '../components/hand';

const t = sceneTimes('S10');
const S16 = t.sub(16);
const S17 = t.sub(17);
const S18 = t.sub(18);
const S16L2 = t.line(16, 2); // 뒷줄 "전 재산을 몰빵하는 분들…" → 비중 막대가 한 종목으로 꽉 참
const S17L2 = t.line(17, 2); // 뒷줄 "생각대로 큰 수익을 얻을 수 있을까요?" → 물음표, 큰 수익?

const CW = 230;
const CH = 150;
const GX = 40;
const GY = 44;
const X0 = (1920 - (6 * CW + 5 * GX)) / 2;
const Y0 = 270;

/** 계좌 카드(S02)의 축소판: 잉크 테두리 + 크림 채움 + 오른쪽 비중 막대(한 종목 100% = 잉크) */
const MiniAccount: React.FC<{readonly x: number; readonly y: number; readonly p: number; readonly fill: number}> = ({x, y, p, fill}) => (
	<g opacity={p} transform={`translate(0 ${(1 - p) * 24})`}>
		<path d={roundRect(x, y, CW, CH, 16)} fill={C.paper} stroke={C.ink} strokeWidth={4} />
		<rect x={x + 22} y={y + 30} width={92} height={12} rx={6} fill={alpha(C.ink, 0.22)} />
		<rect x={x + 22} y={y + 62} width={120} height={26} rx={8} fill={alpha(C.ink, 0.35)} />
		<rect x={x + CW - 66} y={y + 20} width={44} height={CH - 40} rx={6} fill={C.paper} stroke={C.ink} strokeWidth={3} />
		<rect x={x + CW - 66} y={y + 20 + (CH - 40) * (1 - fill)} width={44} height={(CH - 40) * fill} rx={6} fill={C.ink} />
	</g>
);

export const S10: React.FC = () => {
	const f = useSceneFrame();
	const dim = 1 - 0.75 * prog(f, S17, S17 + 12);
	const out = exitP(f, S18, 9);
	const q = questionPath(960, 170, 330);
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer>
				<Svg>
					<g opacity={dim * out}>
						{Array.from({length: 12}, (_, i) => (
							<MiniAccount key={i} x={X0 + (i % 6) * (CW + GX)} y={Y0 + Math.floor(i / 6) * (CH + GY)} p={enterP(f, S16 + i * 3, 14)} fill={easeInOut(prog(f, S16L2 + i * 2, S16L2 + 16 + i * 2, (x) => x))} />
						))}
					</g>
					<g opacity={out}>
						<DrawPath d={q.hook} p={prog(f, S17L2, S17L2 + 18)} width={28} />
						<circle cx={q.dot[0]} cy={q.dot[1]} r={21 * enterP(f, S17L2 + 16, 8)} fill={C.ink} />
					</g>
				</Svg>
				<Reveal at={S17L2 + 6} exitAt={S18} from="up" style={{left: 0, right: 0, top: 560, textAlign: 'center'}}>
					<div style={{...T.headline, fontSize: 88}}>
						<Highlight at={S17L2 + 14}>큰 수익</Highlight>?
					</div>
				</Reveal>
				<Reveal at={S18 + 8} from="up" style={{left: 0, right: 0, top: 330, textAlign: 'center'}}>
					<div style={{fontFamily: SERIF, fontWeight: 900, fontSize: 88, color: C.ink, lineHeight: 1.2}}>무엇이 문제일까?</div>
				</Reveal>
				<Svg>
					<DrawPath d={handLine(700, 460, 1220, 454, 's10ul', 3)} p={prog(f, S18 + 18, S18 + 30)} width={9} />
				</Svg>
			</Layer>
		</AbsoluteFill>
	);
};
