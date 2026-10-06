// VERSION: v2.0.0 — 2026-10-06 — 10×10 와플 그리드 (S18 공용, 칸 단위 스타일 함수)
// v2: 블러 글로우 제거(속도 규칙) → "빛남"은 칸 바깥 노랑 테두리(ring)로 표현.
import React from 'react';
import {C} from '../design/colors';
import {Svg} from './Draw';

export type CellStyle = {
	readonly fill: string;
	readonly opacity?: number;
	readonly scale?: number;
	/** 0~1, 칸 바깥 노랑 테두리 */
	readonly ring?: number;
	readonly stroke?: string;
	readonly strokeWidth?: number;
};

export const waffleCellXY = (i: number, x: number, y: number, cell: number, gap: number, cols = 10): [number, number] => {
	const c = i % cols;
	const r = Math.floor(i / cols);
	return [x + c * (cell + gap), y + r * (cell + gap)];
};

export const Waffle: React.FC<{
	readonly x: number;
	readonly y: number;
	readonly cell: number;
	readonly gap: number;
	readonly n?: number;
	readonly cols?: number;
	readonly cellStyle: (i: number) => CellStyle;
}> = ({x, y, cell, gap, n = 100, cols = 10, cellStyle}) => {
	const cells = Array.from({length: n}, (_, i) => ({i, s: cellStyle(i), xy: waffleCellXY(i, x, y, cell, gap, cols)}));
	return (
		<Svg>
			{cells.map(({i, s, xy: [cx, cy]}) => {
				const k = s.scale ?? 1;
				const w = cell * k;
				const ring = s.ring ?? 0;
				return (
					<g key={i}>
						{ring > 0 ? (
							<rect x={cx - 7} y={cy - 7} width={cell + 14} height={cell + 14} rx={10} fill="none" stroke={C.yellow} strokeWidth={6} opacity={ring} />
						) : null}
						<rect
							x={cx + (cell - w) / 2}
							y={cy + (cell - w) / 2}
							width={w}
							height={w}
							rx={6}
							fill={s.fill}
							opacity={s.opacity ?? 1}
							stroke={s.stroke}
							strokeWidth={s.strokeWidth}
						/>
					</g>
				);
			})}
		</Svg>
	);
};
