// VERSION: v1.0.0 — 2026-10-05 — 노랑 형광펜: 글자 뒤에서 왼쪽→오른쪽으로 10프레임 동안 그어짐
import React from 'react';
import {C} from '../design/colors';
import {M, prog} from '../design/motion';
import {useSceneFrame} from './Scene';

export const Highlight: React.FC<{
	readonly at: number;
	readonly children: React.ReactNode;
	readonly color?: string;
	readonly dur?: number;
	/** 글자 높이 중 형광펜이 시작되는 위치(위에서부터 비율) */
	readonly top?: number;
	readonly style?: React.CSSProperties;
}> = ({at, children, color = C.yellow, dur = M.highlight, top = 0.42, style}) => {
	const f = useSceneFrame();
	const p = prog(f, at, at + dur, (t) => t);
	return (
		<span style={{position: 'relative', display: 'inline-block', isolation: 'isolate', ...style}}>
			<svg
				viewBox="0 0 100 100"
				preserveAspectRatio="none"
				style={{
					position: 'absolute',
					left: '-0.12em',
					width: 'calc(100% + 0.24em)',
					top: `${top * 100}%`,
					height: `${(0.98 - top) * 100}%`,
					zIndex: -1,
					overflow: 'visible',
					clipPath: `inset(-10% ${(1 - p) * 100}% -10% 0)`,
				}}
			>
				<path d="M 1 14 C 25 6, 60 12, 99 4 L 100 90 C 70 97, 35 92, 0 97 Z" fill={color} />
			</svg>
			{children}
		</span>
	);
};
