// VERSION: v2.0.0 — 2026-10-06 — 사연자의 계좌 카드 / 계좌 막대 (영상 내내 같은 모양·같은 색)
// AccountCard: 잉크 테두리 + 크림 채움 카드. 왼쪽 = 금액, 오른쪽 = MDB 비중 막대(잉크 채움).
//   pnlMode 0 → S02 배치(큰 총액), 1 → S03 이후 배치(평가손익 큰 숫자).
// StackBar: 세로 계좌 막대(높이 = 금액). 칸(종목)을 아래부터 쌓고, 잃은 부분은 칸 위쪽 파랑 점선(고스트).
import React from 'react';
import {C, mix, pnlColor} from '../design/colors';
import {SANS} from '../design/fonts';
import {T} from '../design/type';
import {formatManwon} from '../data/facts';
import {cardStyle} from './Bits';

export const ACCOUNT_CARD = {w: 1240, h: 480} as const;

export const AccountCard: React.FC<{
	readonly x: number;
	readonly y: number;
	readonly scale?: number;
	readonly pnlMode: number;
	readonly totalManwon: number;
	readonly pnlManwon?: number;
	readonly bar: number;
	readonly desat?: number;
	readonly dark?: boolean;
	readonly opacity?: number;
	readonly dx?: number;
	readonly dy?: number;
}> = ({x, y, scale = 1, pnlMode, totalManwon, pnlManwon = 0, bar, desat = 0, dark = false, opacity = 1, dx = 0, dy = 0}) => {
	if (opacity <= 0.001) return null;
	const {w, h} = ACCOUNT_CARD;
	const pad = 44;
	const barW = 150;
	const barH = h - pad * 2 - 64;
	const pnlC = mix(pnlManwon === 0 ? C.ink : pnlColor(pnlManwon), C.gray, desat);
	const a = 1 - pnlMode;
	const b = pnlMode;
	return (
		<div
			style={{
				...cardStyle(dark),
				left: x,
				top: y,
				width: w,
				height: h,
				opacity,
				translate: `${dx}px ${dy}px`,
				scale: String(scale),
				transformOrigin: '0 0',
			}}
		>
			{/* 배치 A: S02 — 내 계좌 / 7,000만 원 (큰 숫자) */}
			<div style={{position: 'absolute', left: pad, top: pad - 6, opacity: a}}>
				<div style={{...T.label}}>내 계좌</div>
				<div style={{...T.number, fontSize: 180, color: C.ink, marginTop: 30}}>{formatManwon(totalManwon)}</div>
			</div>
			{/* 배치 B: S03 이후 — 내 계좌 / 평가손익 (큰 숫자). 잔고 금액은 데이터 시트에 없는 값(8,000만 등)이 되므로 쓰지 않는다 */}
			<div style={{position: 'absolute', left: pad, top: pad - 6, opacity: b}}>
				<div style={{...T.label}}>내 계좌</div>
				<div style={{...T.label, color: C.gray, marginTop: 34}}>평가손익</div>
				<div style={{...T.number, fontSize: 180, color: pnlC, marginTop: 10}}>{formatManwon(pnlManwon, true)}</div>
			</div>
			{/* MDB 비중 막대 */}
			<div
				style={{
					position: 'absolute',
					right: pad,
					top: pad,
					width: barW,
					height: barH,
					border: `4px solid ${C.ink}`,
					borderRadius: 10,
					boxSizing: 'border-box',
					overflow: 'hidden',
					background: C.paper,
				}}
			>
				<div style={{position: 'absolute', left: 0, right: 0, bottom: 0, height: `${bar * 100}%`, background: C.ink}} />
			</div>
			<div
				style={{
					position: 'absolute',
					right: pad - 30,
					width: barW + 60,
					top: pad + barH + 14,
					textAlign: 'center',
					fontFamily: SANS,
					fontWeight: 700,
					fontSize: 36,
					color: C.ink,
					fontVariantNumeric: 'tabular-nums',
					opacity: bar > 0.02 ? 1 : 0,
				}}
			>
				MDB {Math.round(bar * 100)}%
			</div>
		</div>
	);
};

// ---------------------------------------------------------------------------
export type SegKind = 'mdb' | 'stock' | 'etf';
export type Seg = {
	/** 칸 원래 높이(px) */
	readonly h: number;
	readonly kind: SegKind;
	/** 잃은 높이(px): 칸 위쪽이 파랑 점선 고스트로 바뀐다 */
	readonly lost?: number;
	/** 늘어난 높이(px): 칸 위로 더 쌓인다(같은 색) */
	readonly opacity?: number;
};

/** 지수 ETF 칸 채움용 점 패턴 (S27: "대표 기업 500곳" 점과 같은 무늬) */
export const EtfPatternDefs: React.FC<{readonly id: string}> = ({id}) => (
	<defs>
		<pattern id={id} width={14} height={14} patternUnits="userSpaceOnUse">
			<rect width={14} height={14} fill={C.paper} />
			<circle cx={7} cy={7} r={3.2} fill={C.ink} />
		</pattern>
	</defs>
);

/** 세로 계좌 막대 (SVG 안에서 사용). bottom = 바닥 y, 칸은 아래부터 위로 쌓인다. */
export const StackBar: React.FC<{
	readonly x: number;
	readonly bottom: number;
	readonly w: number;
	readonly segs: readonly Seg[];
	readonly etfPatternId?: string;
	readonly opacity?: number;
}> = ({x, bottom, w, segs, etfPatternId, opacity = 1}) => {
	let y = bottom;
	return (
		<g opacity={opacity}>
			{segs.map((s, i) => {
				const top = y - s.h;
				y = top;
				const lost = Math.min(s.h, Math.max(0, s.lost ?? 0));
				const solidH = s.h - lost;
				const fill = s.kind === 'mdb' ? C.ink : s.kind === 'etf' ? `url(#${etfPatternId})` : C.paper;
				return (
					<g key={i} opacity={s.opacity ?? 1}>
						{lost > 0.5 ? (
							<rect x={x} y={top} width={w} height={lost} fill="rgba(45,108,223,0.12)" stroke={C.blue} strokeWidth={3} strokeDasharray="9 7" />
						) : null}
						{solidH > 0.5 ? <rect x={x} y={top + lost} width={w} height={solidH} fill={fill} stroke={C.ink} strokeWidth={4} /> : null}
					</g>
				);
			})}
		</g>
	);
};
