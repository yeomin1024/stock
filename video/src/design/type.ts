// VERSION: v1.0.0 — 2026-10-05 — 텍스트 스타일 프리셋 (가이드 3-3)
import type React from 'react';
import {C} from './colors';
import {SANS, SERIF, TYPE} from './fonts';

export const T = {
	headline: {
		fontFamily: SERIF,
		fontWeight: 900,
		fontSize: TYPE.headline,
		color: C.ink,
		lineHeight: 1.18,
		letterSpacing: '-0.01em',
		whiteSpace: 'nowrap',
	},
	label: {
		fontFamily: SANS,
		fontWeight: 700,
		fontSize: TYPE.label,
		color: C.ink,
		lineHeight: 1.25,
		whiteSpace: 'nowrap',
	},
	number: {
		fontFamily: SANS,
		fontWeight: 900,
		fontSize: TYPE.number,
		lineHeight: 1,
		letterSpacing: '-0.02em',
		fontVariantNumeric: 'tabular-nums',
		whiteSpace: 'nowrap',
	},
	source: {
		fontFamily: SANS,
		fontWeight: 500,
		fontSize: TYPE.source,
		color: C.gray,
		whiteSpace: 'nowrap',
	},
} satisfies Record<string, React.CSSProperties>;
