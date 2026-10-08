// VERSION: v2.3.0 — 2026-10-08 — 한 줄 자막: 뒷줄 내용 요소는 t.line(n, 2) · S21 (자막 49–51) "04" + "판단력까지 무너진다"
// 자막 51: 천칭 저울 — 왼쪽(나쁜 쪽) "잃을 때 고통 ×2" 가 오른쪽 "얻을 때 기쁨" 보다 무겁게 기운다. 연결 근거: 손실 회피
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {FACTS, SRC} from '../data/facts';
import {sceneTimes} from '../data/timeline';
import {C} from '../design/colors';
import {easeInOut, enterP, prog} from '../design/motion';
import {T} from '../design/type';
import {NumberTitle, SourceCaption} from '../components/Bits';
import {DrawPath, Svg} from '../components/Draw';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';

const t = sceneTimes('S21');
const S49 = t.sub(49);
const S51 = t.sub(51);
const S51L2 = t.line(51, 2); // 뒷줄 "같은 금액을 얻을 때의 기쁨보다 두 배 정도 큽니다."
const TILT_AT = S51L2 + 16;
const BLOCK_AT = [S51 + 14, S51L2 + 10, S51L2 + 4]; // 파랑 1(고통) / 파랑 2(두 배) / 빨강(기쁨)

const PX = 1160; // 받침점
const PY = 380;
const HALF = 400;
const HANG = 150;
const POST_BOTTOM = 770;

export const S21: React.FC = () => {
	const f = useSceneFrame();
	const appear = enterP(f, S51, 16);
	const tilt = easeInOut(prog(f, TILT_AT, TILT_AT + 24, (x) => x)) * 11; // 왼쪽이 내려감
	const rad = (tilt * Math.PI) / 180;
	const L = {x: PX - HALF * Math.cos(rad), y: PY + HALF * Math.sin(rad)};
	const R = {x: PX + HALF * Math.cos(rad), y: PY - HALF * Math.sin(rad)};
	const blocksIn = (k: number) => enterP(f, BLOCK_AT[k], 12);
	const pan = (x: number, y: number) => `M ${x - 130} ${y + HANG} Q ${x} ${y + HANG + 60} ${x + 130} ${y + HANG}`;
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer>
				<NumberTitle num="04" title="판단력까지 무너진다" mark="무너진다" at={S49} markAt={S49 + 24} />
				{appear > 0.001 ? (
					<Svg>
						<g opacity={appear}>
							<DrawPath d={`M ${PX} ${PY} L ${PX} ${POST_BOTTOM}`} p={1} width={8} />
							<DrawPath d={`M ${PX - 110} ${POST_BOTTOM} L ${PX + 110} ${POST_BOTTOM}`} p={1} width={8} />
							<line x1={L.x} y1={L.y} x2={R.x} y2={R.y} stroke={C.ink} strokeWidth={9} strokeLinecap="round" />
							<circle cx={PX} cy={PY} r={14} fill={C.ink} />
							{[L, R].map((e, i) => (
								<g key={i}>
									<line x1={e.x} y1={e.y} x2={e.x - 120} y2={e.y + HANG} stroke={C.ink} strokeWidth={3} />
									<line x1={e.x} y1={e.y} x2={e.x + 120} y2={e.y + HANG} stroke={C.ink} strokeWidth={3} />
									<path d={pan(e.x, e.y)} fill={C.paper} stroke={C.ink} strokeWidth={6} />
								</g>
							))}
							{/* 왼쪽: 고통 블록 2개(파랑) / 오른쪽: 기쁨 블록 1개(빨강) — "두 배" */}
							{[0, 1].map((k) => (
								<rect key={k} x={L.x - 92 + k * 96} y={L.y + HANG - 70} width={86} height={70} rx={8} fill={C.blue} stroke={C.ink} strokeWidth={3} opacity={blocksIn(k)} />
							))}
							<rect x={R.x - 43} y={R.y + HANG - 70} width={86} height={70} rx={8} fill={C.red} stroke={C.ink} strokeWidth={3} opacity={blocksIn(2)} />
						</g>
					</Svg>
				) : null}
				<Reveal at={S51 + 18} from="up" dist={16} style={{left: L.x - 210, width: 420, top: L.y + HANG + 50, textAlign: 'center'}}>
					<div style={{...T.label, fontSize: 44, color: C.blue}}>
						잃을 때 고통 <span style={{opacity: prog(f, S51L2 + 14, S51L2 + 20)}}>×{FACTS.lossAversion}</span>
					</div>
				</Reveal>
				<Reveal at={S51L2 + 6} from="up" dist={16} style={{left: R.x - 210, width: 420, top: R.y + HANG + 50, textAlign: 'center'}}>
					<div style={{...T.label, fontSize: 44, color: C.red}}>얻을 때 기쁨</div>
				</Reveal>
			</Layer>
			<SourceCaption text={SRC.behavioral} at={S51 + 18} />
		</AbsoluteFill>
	);
};
