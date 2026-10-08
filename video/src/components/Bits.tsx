// VERSION: v2.3.0 — 2026-10-08 — NumberTitle titleAt 추가 · 출처 캡션, "개념도", 카드 스타일, 번호 타이틀, 장면 타이틀
import React from 'react';
import {C} from '../design/colors';
import {LAYOUT, SERIF, TYPE} from '../design/fonts';
import {T} from '../design/type';
import {Highlight} from './Highlight';
import {Reveal} from './Reveal';

/** 출처 캡션 (웜 그레이). 기본 위치: 그래픽 영역 왼쪽 아래 */
export const SourceCaption: React.FC<{
	readonly text: string;
	readonly at: number;
	readonly exitAt?: number;
	readonly x?: number;
	readonly y?: number;
	readonly align?: 'left' | 'right';
	readonly color?: string;
}> = ({text, at, exitAt, x = LAYOUT.margin, y = LAYOUT.bottom - 36, align = 'left', color}) => (
	<Reveal at={at} exitAt={exitAt} from="none" style={align === 'left' ? {left: x, top: y} : {right: 1920 - x, top: y}}>
		<div style={{...T.caption, color: color ?? T.caption.color}}>{text}</div>
	</Reveal>
);

/** 실제 데이터가 아닌 개념 그래프 표시 (구석에 작게) */
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
				...T.caption,
				color,
				border: `2px solid ${color}`,
				borderRadius: 6,
				padding: '0px 10px 2px',
				lineHeight: 1.3,
			}}
		>
			개념도
		</div>
	</Reveal>
);

/** 카드: 잉크 테두리 + 크림 채움. 그림자는 블러 없는 작은 단색 오프셋 (속도 규칙) */
export const cardStyle = (dark = false): React.CSSProperties => ({
	position: 'absolute',
	background: C.paper,
	border: `4px solid ${C.ink}`,
	borderRadius: 22,
	boxShadow: dark ? '0 8px 0 rgba(0,0,0,0.35)' : '6px 6px 0 rgba(30,30,30,0.12)',
	boxSizing: 'border-box',
});

/** 큰 번호 타이틀 ("01" + 헤드라인). 왼쪽 위에 두고 아래 영역을 내용에 쓴다. */
export const NumberTitle: React.FC<{
	readonly num: string;
	readonly title: string;
	/** 형광펜을 그을 부분 (title 안의 부분 문자열) */
	readonly mark?: string;
	readonly at: number;
	/** 제목이 번호보다 늦게(뒷줄 자막에) 나올 때 — 없으면 at + 5 */
	readonly titleAt?: number;
	readonly markAt?: number;
	readonly exitAt?: number;
}> = ({num, title, mark, at, titleAt, markAt, exitAt}) => {
	const i = mark ? title.indexOf(mark) : -1;
	return (
		<>
			<Reveal at={at} exitAt={exitAt} from="up" style={{left: LAYOUT.margin + 8, top: LAYOUT.top + 8}}>
				<div style={{...T.headline, fontSize: 88, color: C.ink, lineHeight: 1}}>{num}</div>
			</Reveal>
			<Reveal at={titleAt ?? at + 5} exitAt={exitAt} from="left" style={{left: LAYOUT.margin + 8, top: LAYOUT.top + 118}}>
				<div style={{...T.headline, fontSize: 80}}>
					{i >= 0 && mark ? (
						<>
							{title.slice(0, i)}
							<Highlight at={markAt ?? (titleAt ?? at) + 20}>{mark}</Highlight>
							{title.slice(i + mark.length)}
						</>
					) : (
						title
					)}
				</div>
			</Reveal>
		</>
	);
};

/** 장면 타이틀 ("첫째 · 비중 상한" 같은 머리글). 왼쪽 위 */
export const SceneTitle: React.FC<{readonly text: string; readonly at: number; readonly exitAt?: number; readonly size?: number}> = ({
	text,
	at,
	exitAt,
	size = 76,
}) => (
	<Reveal at={at} exitAt={exitAt} from="left" style={{left: LAYOUT.margin + 8, top: LAYOUT.top + 4}}>
		<div style={{fontFamily: SERIF, fontWeight: 900, fontSize: size, color: C.ink, lineHeight: 1.15, whiteSpace: 'nowrap'}}>{text}</div>
	</Reveal>
);

export const MIN_TEXT = TYPE.min;
