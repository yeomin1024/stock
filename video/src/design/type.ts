// VERSION: v2.0.0 — 2026-10-06 — 텍스트 스타일 프리셋 (가이드 v2 3-2)
import type React from 'react';
import {C} from './colors';
import {SANS, SERIF, TYPE} from './fonts';

export const T = {
	/** 장면 헤드라인·번호 타이틀: Noto Serif KR 900, 72–88px */
	headline: {
		fontFamily: SERIF,
		fontWeight: 900,
		fontSize: TYPE.headline,
		color: C.ink,
		lineHeight: 1.2,
		letterSpacing: '-0.01em',
		whiteSpace: 'nowrap',
	},
	/** 라벨: Noto Sans KR 700, 36–44px */
	label: {
		fontFamily: SANS,
		fontWeight: 700,
		fontSize: TYPE.label,
		color: C.ink,
		lineHeight: 1.3,
		whiteSpace: 'nowrap',
	},
	/** 핵심 숫자: Noto Sans KR 900, 160–220px, tabular-nums */
	number: {
		fontFamily: SANS,
		fontWeight: 900,
		fontSize: TYPE.number,
		lineHeight: 1,
		letterSpacing: '-0.02em',
		fontVariantNumeric: 'tabular-nums',
		whiteSpace: 'nowrap',
	},
	/** 출처·개념도 캡션: 웜 그레이 */
	caption: {
		fontFamily: SANS,
		fontWeight: 500,
		fontSize: TYPE.caption,
		color: C.gray,
		whiteSpace: 'nowrap',
	},
} satisfies Record<string, React.CSSProperties>;
