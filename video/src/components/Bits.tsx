// VERSION: v1.0.0 — 2026-10-05 — 출처 캡션, "개념도" 태그, 카드 스타일, 큰 번호+헤드라인
import React from 'react';
import {C} from '../design/colors';
import {SANS} from '../design/fonts';
import {T} from '../design/type';
import {Reveal} from './Reveal';

/** 출처 캡션 (22px, 웜 그레이). 그래픽 영역(y<820) 안쪽 하단 기본 위치 */
export const SourceCaption: React.FC<{
	readonly text: string;
	readonly at: number;
	readonly exitAt?: number;
	readonly x?: number;
	readonly y?: number;
	readonly align?: 'left' | 'right';
	readonly color?: string;
}> = ({text, at, exitAt, x = 120, y = 772, align = 'left', color}) => (
	<Reveal
		at={at}
		exitAt={exitAt}
		from="none"
		style={align === 'left' ? {left: x, top: y} : {right: 1920 - x, top: y}}
	>
		<div style={{...T.source, color: color ?? T.source.color}}>{text}</div>
	</Reveal>
);

/** 실제 데이터가 아닌 개념 그래프 표시 */
export const ConceptTag: React.FC<{
	readonly at: number;
	readonly exitAt?: number;
	readonly x: number;
	readonly y: number;
	readonly color?: string;
}> = ({at, exitAt, x, y, color = C.gray}) => (
	<Reveal at={at} exitAt={exitAt} from="none" style={{left: x, top: y}}>
		<div
			style={{
				fontFamily: SANS,
				fontWeight: 500,
				fontSize: 22,
				color,
				border: `1.5px solid ${color}`,
				borderRadius: 6,
				padding: '1px 10px 2px',
				whiteSpace: 'nowrap',
			}}
		>
			개념도
		</div>
	</Reveal>
);

export const cardStyle = (dark = false): React.CSSProperties => ({
	position: 'absolute',
	background: C.paper,
	border: `4px solid ${C.ink}`,
	borderRadius: 22,
	boxShadow: dark ? '0 16px 0 rgba(0,0,0,0.35)' : '12px 12px 0 rgba(30,30,30,0.13)',
});

/** 섹션 번호("01")와 헤드라인 — S14, S17 공용 */
export const SectionTitle: React.FC<{
	readonly num: string;
	readonly title: React.ReactNode;
	readonly at: number;
	readonly x?: number;
	readonly y?: number;
}> = ({num, title, at, x = 160, y = 220}) => (
	<>
		<Reveal at={at} from="up" style={{left: x, top: y}}>
			<div style={{...T.number, fontSize: 220, color: C.ink, letterSpacing: '-0.04em'}}>{num}</div>
		</Reveal>
		<Reveal at={at + 6} from="left" style={{left: x + 6, top: y + 250}}>
			<div style={{...T.headline, fontSize: 92}}>{title}</div>
		</Reveal>
	</>
);
