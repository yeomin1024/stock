// VERSION: v2.3.0 — 2026-10-08 — 한 줄 자막: 뒷줄 내용 요소는 t.line(n, 2) · 자막 번호 −5(대본에서 엔론 5문장 삭제) · S30 (자막 86–91)
// 86: 왼쪽 "위험 4가지", 오른쪽 "방법 5가지" 두 열 → 87: "늦지 않았습니다"
// 88: 스마트폰 화면 "내 계좌 최대 비중 종목 __%" 칸이 깜빡임 → 89: "큰 수익보다 대처할 수 있는 투자" 노랑 형광펜
// 90–91: 천천히 페이드아웃 (그래픽 → 크림, 마지막 1초 여유 동안 잉크로). 연결 근거: 마무리 당부
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {sceneTimes} from '../data/timeline';
import {C} from '../design/colors';
import {SANS, SERIF} from '../design/fonts';
import {enterP, exitP, lin} from '../design/motion';
import {T} from '../design/type';
import {Svg} from '../components/Draw';
import {Highlight} from '../components/Highlight';
import {phoneFrame} from '../components/Icons';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {PHONE} from './S01';

const t = sceneTimes('S30');
const [S86, S87, S88, S89, S90] = [86, 87, 88, 89, 90].map((n) => t.sub(n));
const END91 = t.subEnd(91);
const S86L2 = t.line(86, 2); // 뒷줄 "분산 투자의 중요성을 알아보았습니다." → 오른쪽 열
const RISKS = ['대처 불가', '회복 어려움', '낮은 확률', '판단력 붕괴'];
const WAYS = ['비중 상한', '업종 분산', '지수 ETF', '나눠 사기·현금', '비중 확인'];

const Column: React.FC<{readonly x: number; readonly title: string; readonly items: readonly string[]; readonly bullet: string}> = ({x, title, items, bullet}) => (
	<>
		<Reveal at={x > 900 ? S86L2 : S86} exitAt={S87 - 2} from="up" style={{left: x, top: 140}}>
			<div style={{fontFamily: SERIF, fontWeight: 900, fontSize: 72, color: C.ink}}>{title}</div>
		</Reveal>
		{items.map((it, i) => (
			<Reveal key={it} at={(x > 900 ? S86L2 : S86) + 12 + i * 4} exitAt={S87 - 2} from="left" dist={20} style={{left: x + 6, top: 270 + i * 88}}>
				<div style={{...T.label, fontSize: 44, display: 'flex', alignItems: 'center', gap: 20}}>
					<span style={{display: 'inline-block', width: 18, height: 18, borderRadius: 4, background: bullet}} />
					{it}
				</div>
			</Reveal>
		))}
	</>
);

export const S30: React.FC = () => {
	const f = useSceneFrame();
	const frame = phoneFrame(PHONE.x, PHONE.y, PHONE.w, PHONE.h);
	const phoneP = Math.min(enterP(f, S88, 15), exitP(f, S89 - 2, 9));
	const blink = f >= S88 + 20 && Math.floor((f - S88 - 20) / 10) % 2 === 0;
	const fadeOut = 1 - lin(f, S90, END91);
	const toInk = lin(f, END91, t.dur - 1);
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer>
				<Column x={240} title="위험 4가지" items={RISKS} bullet={C.blue} />
				<Column x={1060} title="방법 5가지" items={WAYS} bullet={C.ink} />
				<Svg>
					<line x1={960} y1={160} x2={960} y2={700} stroke={C.gray} strokeWidth={3} opacity={Math.min(enterP(f, S86 + 4, 12), exitP(f, S87 - 2, 9))} />
				</Svg>
				<Reveal at={S87 + 4} exitAt={S88 - 2} from="up" style={{left: 0, right: 0, top: 370, textAlign: 'center'}}>
					<div style={{fontFamily: SERIF, fontWeight: 900, fontSize: 88, color: C.ink, lineHeight: 1.2}}>늦지 않았습니다</div>
				</Reveal>
				{/* 93: 폰 화면 (S01 과 같은 폰) */}
				{phoneP > 0.001 ? (
					<div style={{position: 'absolute', inset: 0, opacity: phoneP}}>
						<Svg>
							<path d={frame.body} fill={C.paper} stroke={C.ink} strokeWidth={7} />
							<rect x={frame.notch[0]} y={frame.notch[1]} width={frame.notch[2]} height={frame.notch[3]} rx={7} fill={C.ink} />
						</Svg>
						<div style={{position: 'absolute', left: PHONE.x, width: PHONE.w, top: PHONE.y + 110, textAlign: 'center'}}>
							<div style={{...T.label}}>내 계좌</div>
							<div style={{...T.label, marginTop: 6}}>최대 비중 종목</div>
							<div
								style={{
									margin: '40px auto 0',
									width: 300,
									height: 200,
									borderRadius: 20,
									border: `6px solid ${blink ? C.ink : C.gray}`,
									boxSizing: 'border-box',
									display: 'flex',
									alignItems: 'center',
									justifyContent: 'center',
									fontFamily: SANS,
									fontWeight: 900,
									fontSize: 120,
									color: C.ink,
								}}
							>
								__%
							</div>
						</div>
					</div>
				) : null}
				<div style={{position: 'absolute', inset: 0, opacity: fadeOut}}>
					<Reveal at={S89 + 4} from="up" style={{left: 0, right: 0, top: 380, textAlign: 'center'}}>
						<div style={{...T.headline, fontSize: 80}}>
							큰 수익보다 <Highlight at={S89 + 16}>대처할 수 있는 투자</Highlight>
						</div>
					</Reveal>
				</div>
			</Layer>
			<AbsoluteFill style={{backgroundColor: C.ink, opacity: toInk}} />
		</AbsoluteFill>
	);
};
