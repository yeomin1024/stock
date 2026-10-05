// VERSION: v1.0.0 — 2026-10-05 — stroke-dashoffset 로 그려지는 선 (실선/점선 공용)
// 점선은 dasharray 가 그리기 애니메이션과 충돌하므로, 실선 마스크를 그려 점선을 드러내는 방식.
import React, {useId} from 'react';
import {C} from '../design/colors';

/** url(#id) 에 안전한 id (React useId 의 콜론 제거) */
export const useSafeId = (prefix: string): string => `${prefix}-${useId().replace(/[^a-zA-Z0-9_-]/g, '')}`;

export const DrawPath: React.FC<{
	readonly d: string;
	/** 0→1 그려진 비율 */
	readonly p: number;
	readonly stroke?: string;
	readonly width?: number;
	/** 점선 패턴 (예: "14 12"). 있으면 마스크 방식으로 그린다 */
	readonly dash?: string;
	readonly linecap?: 'round' | 'butt' | 'square';
	readonly opacity?: number;
}> = ({d, p, stroke = C.ink, width = 4, dash, linecap = 'round', opacity = 1}) => {
	const id = useSafeId('draw');
	if (p <= 0.001) return null;
	const k = Math.min(1, p);
	if (!dash) {
		return (
			<path
				d={d}
				pathLength={1}
				strokeDasharray="1 2"
				strokeDashoffset={1 - k}
				stroke={stroke}
				strokeWidth={width}
				strokeLinecap={linecap}
				strokeLinejoin="round"
				fill="none"
				opacity={opacity}
			/>
		);
	}
	return (
		<>
			<defs>
				<mask id={id} maskUnits="userSpaceOnUse" x={-4000} y={-4000} width={12000} height={12000}>
					<path
						d={d}
						pathLength={1}
						strokeDasharray="1 2"
						strokeDashoffset={1 - k}
						stroke="#fff"
						strokeWidth={width + 10}
						strokeLinecap="butt"
						fill="none"
					/>
				</mask>
			</defs>
			<path
				d={d}
				stroke={stroke}
				strokeWidth={width}
				strokeDasharray={dash}
				strokeLinecap={linecap}
				fill="none"
				mask={`url(#${id})`}
				opacity={opacity}
			/>
		</>
	);
};

/** 1920×1080 전체를 덮는 SVG 캔버스 (그래픽 좌표 = 화면 좌표) */
export const Svg: React.FC<{readonly children: React.ReactNode; readonly style?: React.CSSProperties}> = ({
	children,
	style,
}) => (
	<svg
		width={1920}
		height={1080}
		viewBox="0 0 1920 1080"
		style={{position: 'absolute', left: 0, top: 0, overflow: 'visible', ...style}}
	>
		{children}
	</svg>
);
