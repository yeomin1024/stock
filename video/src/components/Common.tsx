// VERSION: v2.0.0 — 2026-10-06 — 공통 컴포넌트: 생각 말풍선, 헤드라인 카드, 기업 하락 카드
import React from 'react';
import {C} from '../design/colors';
import {SANS, SERIF} from '../design/fonts';
import {enterP} from '../design/motion';
import {T} from '../design/type';
import {formatPct} from '../data/facts';
import {useCount} from './Counter';
import {Svg} from './Draw';
import {Highlight} from './Highlight';
import {cloudPath} from './Icons';
import {cardStyle} from './Bits';
import {useSceneFrame} from './Scene';

/** 생각 말풍선 (S09, S29 재사용). tail = 꼬리 원이 향하는 쪽 */
export const ThoughtBubble: React.FC<{
	readonly cx: number;
	readonly cy: number;
	readonly rx: number;
	readonly ry: number;
	readonly text: string;
	readonly at: number;
	readonly seed?: number;
	readonly tail?: 'left' | 'right';
	readonly fontSize?: number;
	readonly opacity?: number;
}> = ({cx, cy, rx, ry, text, at, seed = 1, tail = 'left', fontSize = 44, opacity = 1}) => {
	const f = useSceneFrame();
	const p = enterP(f, at, 15);
	if (p <= 0.001 || opacity <= 0.001) return null;
	const k = 0.75 + 0.25 * p;
	const sx = tail === 'left' ? -1 : 1;
	const dots: [number, number, number][] = [
		[cx + sx * rx * 0.55, cy + ry * 1.18, 16],
		[cx + sx * rx * 0.78, cy + ry * 1.5, 10],
	];
	return (
		<div style={{position: 'absolute', inset: 0, opacity: p * opacity}}>
			<Svg>
				<g style={{transformBox: 'view-box', transformOrigin: `${cx}px ${cy}px`, scale: String(k)}}>
					{dots.map(([x, y, r], i) => (
						<circle key={i} cx={x} cy={y} r={r} fill={C.white} stroke={C.ink} strokeWidth={4} />
					))}
					<path d={cloudPath(cx, cy, rx, ry, 11, seed)} fill={C.white} stroke={C.ink} strokeWidth={5} strokeLinejoin="round" />
				</g>
			</Svg>
			<div
				style={{
					position: 'absolute',
					left: cx - rx,
					top: cy - 50,
					width: rx * 2,
					height: 100,
					display: 'flex',
					alignItems: 'center',
					justifyContent: 'center',
					fontFamily: SANS,
					fontWeight: 700,
					fontSize,
					color: C.ink,
					whiteSpace: 'nowrap',
					scale: String(k),
				}}
			>
				{text}
			</div>
		</div>
	);
};

/** 텍스트로만 만든 뉴스 헤드라인 카드 (S06) */
export const HeadlineCard: React.FC<{
	readonly x: number;
	readonly y: number;
	readonly w: number;
	readonly date: string;
	readonly title: string;
	readonly p: number;
}> = ({x, y, w, date, title, p}) =>
	p <= 0.001 ? null : (
		<div
			style={{
				position: 'absolute',
				left: x,
				top: y,
				width: w,
				background: C.paper,
				boxShadow: '0 8px 0 rgba(0,0,0,0.35)',
				rotate: '-1deg',
				opacity: p,
				translate: `${(1 - p) * -140}px 0px`,
				padding: '36px 56px 44px',
				boxSizing: 'border-box',
			}}
		>
			<div style={{borderTop: `6px solid ${C.ink}`, borderBottom: `2px solid ${C.ink}`, height: 8}} />
			<div style={{fontFamily: SANS, fontWeight: 700, fontSize: 40, color: C.ink, marginTop: 22, fontVariantNumeric: 'tabular-nums'}}>{date}</div>
			<div style={{fontFamily: SERIF, fontWeight: 900, fontSize: 88, lineHeight: 1.2, color: C.ink, marginTop: 10, whiteSpace: 'nowrap'}}>{title}</div>
		</div>
	);

/**
 * 기업(또는 업종) 하락 카드 + 아래로 자라는 파랑 막대 + 카운트다운 숫자 (S14, S26 재사용).
 * 카드 제목은 at 에, 부제(연도·원인)와 막대는 dropAt 에 나온다.
 */
export const DropCard: React.FC<{
	readonly cx: number;
	readonly top: number;
	readonly w: number;
	readonly h: number;
	readonly title: string;
	readonly titleSize?: number;
	readonly line?: string;
	/** 같은 줄의 다른 카드에 보조 줄이 있으면 true: 보조 줄 자리를 비워 제목 높이를 맞춘다 */
	readonly reserveLine?: boolean;
	readonly at: number;
	readonly dropAt?: number;
	readonly pct?: number;
	readonly note?: string;
	readonly pxPerPct: number;
	readonly numberSize?: number;
	readonly mark?: {readonly text: string; readonly at: number};
	readonly opacity?: number;
	readonly dx?: number;
}> = ({cx, top, w, h, title, titleSize = 60, line, reserveLine = false, at, dropAt, pct, note, pxPerPct, numberSize = 160, mark, opacity = 1, dx = 0}) => {
	const f = useSceneFrame();
	const card = enterP(f, at, 15);
	const lineP = dropAt === undefined ? 0 : enterP(f, dropAt, 12);
	const v = useCount(0, pct ?? 0, dropAt ?? 1e9, 26);
	if (card <= 0.001 || opacity <= 0.001) return null;
	const barTop = top + h + 12;
	const barH = Math.abs(v) * pxPerPct;
	const decimals = pct !== undefined && !Number.isInteger(pct) ? 1 : 0;
	return (
		<div style={{position: 'absolute', inset: 0, opacity: Math.min(card, opacity), translate: `${dx}px 0px`}}>
			<div
				style={{
					...cardStyle(),
					left: cx - w / 2,
					top,
					width: w,
					height: h,
					translate: `0px ${(1 - card) * -30}px`,
					display: 'flex',
					flexDirection: 'column',
					alignItems: 'center',
					justifyContent: 'center',
					gap: 8,
				}}
			>
				<div style={{fontFamily: SANS, fontWeight: 900, fontSize: titleSize, lineHeight: 1.05, color: C.ink, whiteSpace: 'nowrap'}}>{title}</div>
				{line || reserveLine ? <div style={{...T.label, fontSize: 32, color: C.ink, opacity: lineP, lineHeight: 1.2}}>{line || '\u00A0'}</div> : null}
			</div>
			{pct !== undefined && dropAt !== undefined && f >= dropAt ? (
				<>
					<div style={{position: 'absolute', left: cx - 70, top: barTop, width: 140, height: barH, background: C.blue, borderRadius: '0 0 10px 10px'}} />
					<div
						style={{
							position: 'absolute',
							left: cx - 320,
							width: 640,
							top: barTop + barH + 10,
							textAlign: 'center',
							...T.number,
							fontSize: numberSize,
							color: C.blue,
						}}
					>
						{formatPct(v, decimals)}
						{note ? <span style={{fontSize: 36, fontWeight: 700, marginLeft: 8}}>({note})</span> : null}
					</div>
					{mark ? (
						<div
							style={{
								position: 'absolute',
								left: cx - 320,
								width: 640,
								top: barTop + barH + numberSize + 22,
								textAlign: 'center',
								...T.label,
								fontSize: 36,
								opacity: enterP(f, mark.at, 12),
							}}
						>
							<Highlight at={mark.at + 4}>{mark.text}</Highlight>
						</div>
					) : null}
				</>
			) : null}
		</div>
	);
};
