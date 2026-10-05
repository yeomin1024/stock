// VERSION: v1.0.0 — 2026-10-05 — 10×10 와플 차트 (S17, S18, S20 공용)
import React from 'react';
import {C} from '../design/colors';
import {Svg, useSafeId} from './Draw';

export type CellStyle = {
	readonly fill: string;
	readonly opacity?: number;
	readonly scale?: number;
	/** 0~1, 노랑 빛 번짐 */
	readonly glow?: number;
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
	readonly glowColor?: string;
}> = ({x, y, cell, gap, n = 100, cols = 10, cellStyle, glowColor = C.yellow}) => {
	const glowId = useSafeId('waffle-glow');
	const cells = Array.from({length: n}, (_, i) => ({i, s: cellStyle(i), xy: waffleCellXY(i, x, y, cell, gap, cols)}));
	return (
		<Svg>
			<defs>
				<filter id={glowId} x="-100%" y="-100%" width="300%" height="300%">
					<feGaussianBlur stdDeviation={10} />
				</filter>
			</defs>
			{cells
				.filter(({s}) => (s.glow ?? 0) > 0)
				.map(({i, s, xy: [cx, cy]}) => (
					<rect
						key={`g${i}`}
						x={cx - 8}
						y={cy - 8}
						width={cell + 16}
						height={cell + 16}
						rx={10}
						fill={glowColor}
						opacity={0.85 * (s.glow ?? 0)}
						filter={`url(#${glowId})`}
					/>
				))}
			{cells.map(({i, s, xy: [cx, cy]}) => {
				const k = s.scale ?? 1;
				const w = cell * k;
				return (
					<rect
						key={i}
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
				);
			})}
		</Svg>
	);
};
