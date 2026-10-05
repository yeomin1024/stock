// VERSION: v1.0.0 — 2026-10-05 — S06 (자막 8) 텍스트로만 만든 신문 헤드라인 카드 + 회사 A→B 마커 화살표
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {FACTS} from '../data/facts';
import {sceneTimes} from '../data/timeline';
import {C, alpha} from '../design/colors';
import {SANS, SERIF} from '../design/fonts';
import {easeInOut, enterP, lerp, prog} from '../design/motion';
import {DrawPath, Svg} from '../components/Draw';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {handArrow, roundRect} from '../components/hand';

const t = sceneTimes('S06');
const CEO_AT = t.word(8, 'CEO가');
const MOVE_AT = t.word(8, '이적');

const NX = 300;
const NY = 96;
const NW = 1320;
const NH = 380;

// 회사 블록
const A = {x: 470, y: 560, w: 320, h: 190};
const B = {x: 1130, y: 560, w: 320, h: 190};
const ARROW = handArrow([A.x + A.w + 24, A.y + 70], [B.x - 24, B.y + 70], 's06arrow', -0.22, 30);

export const S06: React.FC = () => {
	const f = useSceneFrame();
	const paper = enterP(f, 0, 16);
	const blocks = enterP(f, 14, 15);
	const token = enterP(f, CEO_AT - 6, 12);
	const move = easeInOut(prog(f, MOVE_AT + 4, MOVE_AT + 26, (x) => x));
	const tokenX = lerp(A.x + A.w / 2, B.x + B.w / 2, move);
	// 신문 카드 아래(y≥486)를 지나도록 낮은 호
	const tokenY = A.y - 6 - Math.sin(Math.PI * move) * 30;
	const bSolid = prog(f, MOVE_AT + 22, MOVE_AT + 30);
	return (
		<AbsoluteFill>
			<SceneBg tone="navy" />
			<Layer depth="mid">
				{/* 신문 카드 */}
				<div
					style={{
						position: 'absolute',
						left: NX,
						top: NY,
						width: NW,
						height: NH,
						background: C.paper,
						boxShadow: '0 16px 0 rgba(0,0,0,0.35)',
						rotate: '-1deg',
						opacity: paper,
						translate: `${(1 - paper) * -120}px 0px`,
						padding: '34px 48px',
						boxSizing: 'border-box',
					}}
				>
					<div style={{borderTop: `5px solid ${C.ink}`, borderBottom: `2px solid ${C.ink}`, height: 8}} />
					<div style={{fontFamily: SANS, fontWeight: 700, fontSize: 36, color: C.ink, marginTop: 16}}>
						{FACTS.mdb.ceoNewsDate}
					</div>
					<div style={{fontFamily: SERIF, fontWeight: 900, fontSize: 96, lineHeight: 1.2, color: C.ink, marginTop: 6, whiteSpace: 'nowrap'}}>
						CEO, 다른 회사로 이적
					</div>
					{/* 본문 자리 (글자 대신 회색 줄) */}
					<div style={{display: 'flex', gap: 36, marginTop: 22}}>
						{[0, 1, 2].map((c) => (
							<div key={c} style={{flex: 1, display: 'flex', flexDirection: 'column', gap: 12}}>
								{[0, 1, 2].map((r) => (
									<div key={r} style={{height: 10, borderRadius: 5, background: alpha(C.ink, 0.16), width: `${[100, 92, 70][(r + c) % 3]}%`}} />
								))}
							</div>
						))}
					</div>
				</div>
			</Layer>
			<Layer depth="fg">
				<Svg>
					<g opacity={blocks}>
						<rect x={A.x} y={A.y} width={A.w} height={A.h} rx={18} fill={alpha(C.cream, 0.08)} stroke={C.cream} strokeWidth={5} />
						<path d={roundRect(B.x, B.y, B.w, B.h, 18)} fill="none" stroke={C.cream} strokeWidth={5} strokeDasharray="18 14" opacity={1 - bSolid} />
						<path d={roundRect(B.x, B.y, B.w, B.h, 18)} fill={alpha(C.cream, 0.08)} stroke={C.cream} strokeWidth={5} opacity={bSolid} />
					</g>
					<DrawPath d={ARROW.shaft} p={prog(f, MOVE_AT - 14, MOVE_AT + 4)} stroke={C.yellow} width={9} />
					<DrawPath d={ARROW.head} p={prog(f, MOVE_AT + 2, MOVE_AT + 8)} stroke={C.yellow} width={9} />
				</Svg>
				<div style={{position: 'absolute', left: A.x, top: A.y + 110, width: A.w, textAlign: 'center', fontFamily: SANS, fontWeight: 900, fontSize: 56, lineHeight: 1, color: C.cream, opacity: blocks}}>
					{FACTS.mdb.ticker}
				</div>
				{token > 0.001 ? (
					<div
						style={{
							position: 'absolute',
							left: tokenX - 80,
							top: tokenY - 30,
							width: 160,
							height: 74,
							borderRadius: 37,
							background: C.yellow,
							border: `4px solid ${C.ink}`,
							boxSizing: 'border-box',
							display: 'flex',
							alignItems: 'center',
							justifyContent: 'center',
							fontFamily: SANS,
							fontWeight: 900,
							fontSize: 40,
							color: C.ink,
							opacity: token,
							scale: String(0.8 + 0.2 * token),
						}}
					>
						CEO
					</div>
				) : null}
			</Layer>
		</AbsoluteFill>
	);
};
