// VERSION: v2.3.0 — 2026-10-08 — 한 줄 자막: 뒷줄 내용 요소는 t.line(n, 2) · S01 (자막 1–2) 스마트폰 검색 "AI 관련주" → 결과 카드 "몽고디비 MDB" → 자막 2: 체크 3개
// 연결 근거: 종목을 알게 된 경위(자막 1), 고른 이유 3가지(자막 2)
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {FACTS} from '../data/facts';
import {sceneTimes} from '../data/timeline';
import {C} from '../design/colors';
import {SANS} from '../design/fonts';
import {enterP, prog} from '../design/motion';
import {T} from '../design/type';
import {DrawPath, Svg} from '../components/Draw';
import {checkPath, Magnifier, phoneFrame} from '../components/Icons';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {roundRect} from '../components/hand';

const t = sceneTimes('S01');
const QUERY = Array.from('AI 관련주');
const S1 = t.sub(1);
const TYPE_AT = S1 + 12;
const TYPE_STEP = 3;
const CARD_AT = Math.max(TYPE_AT + QUERY.length * TYPE_STEP + 12, t.line(1, 2)); // 뒷줄 "몽고디비라는 종목을…"
const CHECKS = ['전망', '실적', '상승세'].map((label, i) => ({label, at: t.sub(2) + i * 6}));

export const PHONE = {x: 720, y: 130, w: 480, h: 670} as const;

export const S01: React.FC = () => {
	const f = useSceneFrame();
	const {x: PX, y: PY, w: PW, h: PH} = PHONE;
	const frame = phoneFrame(PX, PY, PW, PH);
	const phoneP = prog(f, S1, S1 + 18);
	const fillP = prog(f, S1 + 8, S1 + 20);
	const typed = Math.max(0, Math.min(QUERY.length, Math.floor((f - TYPE_AT) / TYPE_STEP) + 1));
	const cursorOn = f >= TYPE_AT - 4 && f < CARD_AT && Math.floor(f / 8) % 2 === 0;
	const card = enterP(f, CARD_AT, 16);
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer>
				<Svg>
					<path d={frame.body} fill={C.paper} opacity={fillP} />
					<DrawPath d={frame.body} p={phoneP} width={7} />
					<rect x={frame.notch[0]} y={frame.notch[1]} width={frame.notch[2]} height={frame.notch[3]} rx={7} fill={C.ink} opacity={fillP} />
					<rect x={PX + 34} y={PY + 64} width={PW - 68} height={80} rx={40} fill={C.white} opacity={fillP} />
					<DrawPath d={roundRect(PX + 34, PY + 64, PW - 68, 80, 40)} p={prog(f, S1 + 6, S1 + 20)} width={4} />
					<Magnifier cx={PX + 80} cy={PY + 100} r={15} p={prog(f, S1 + 10, S1 + 22)} />
				</Svg>
				<div style={{position: 'absolute', left: PX + 120, top: PY + 76, ...T.label, fontSize: 40, lineHeight: '56px', display: 'flex', alignItems: 'center'}}>
					{QUERY.slice(0, typed).join('')}
					<span style={{display: 'inline-block', width: 4, height: 44, marginLeft: 4, background: C.ink, opacity: cursorOn ? 1 : 0}} />
				</div>
				{card > 0.001 ? (
					<div
						style={{
							position: 'absolute',
							left: PX + 34,
							top: PY + 176,
							width: PW - 68,
							height: 176,
							background: C.white,
							border: `4px solid ${C.ink}`,
							borderRadius: 22,
							opacity: card,
							translate: `0px ${(1 - card) * 50}px`,
							padding: '26px 30px',
							boxSizing: 'border-box',
						}}
					>
						<div style={{fontFamily: SANS, fontWeight: 900, fontSize: 56, lineHeight: 1.1, color: C.ink}}>{FACTS.story.nameKo}</div>
						<div style={{...T.label, color: C.gray, marginTop: 8}}>{FACTS.story.ticker}</div>
					</div>
				) : null}
				{CHECKS.map((c, i) => {
					const y = PY + 392 + i * 88;
					const boxP = enterP(f, c.at, 12);
					return (
						<React.Fragment key={c.label}>
							<Svg>
								<rect x={PX + 50} y={y} width={58} height={58} rx={10} fill={C.white} stroke={C.ink} strokeWidth={4} opacity={boxP} />
								<DrawPath d={checkPath(PX + 56, y + 4, 52)} p={prog(f, c.at + 2, c.at + 12)} width={8} />
							</Svg>
							<div style={{position: 'absolute', left: PX + 132, top: y + 2, ...T.label, fontSize: 42, lineHeight: '54px', opacity: boxP}}>{c.label}</div>
						</React.Fragment>
					);
				})}
			</Layer>
		</AbsoluteFill>
	);
};
