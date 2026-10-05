// VERSION: v1.0.0 — 2026-10-05 — 공통 등장/퇴장 래퍼 (spring ease-out 12–18f, 퇴장 8–10f)
import React from 'react';
import {M, enterP, exitP} from '../design/motion';
import {useSceneFrame} from './Scene';

type From = 'up' | 'down' | 'left' | 'right' | 'scale' | 'none';

export const Reveal: React.FC<{
	readonly at: number;
	readonly exitAt?: number;
	readonly dur?: number;
	readonly exitDur?: number;
	readonly from?: From;
	readonly dist?: number;
	readonly style?: React.CSSProperties;
	readonly children: React.ReactNode;
}> = ({at, exitAt, dur = M.enter, exitDur = M.exit, from = 'up', dist = 40, style, children}) => {
	const f = useSceneFrame();
	const e = enterP(f, at, dur);
	const x = exitAt === undefined ? 1 : exitP(f, exitAt, exitDur);
	const v = Math.min(e, x);
	if (v <= 0.001) return null;
	const k = 1 - e;
	const out = 1 - x;
	const dx = from === 'left' ? -dist * k : from === 'right' ? dist * k : 0;
	const dy = from === 'up' ? dist * k - out * dist * 0.4 : from === 'down' ? -dist * k + out * dist * 0.4 : 0;
	const scale = from === 'scale' ? 0.85 + 0.15 * e : 1;
	return (
		<div
			style={{
				position: 'absolute',
				opacity: v,
				translate: `${dx}px ${dy}px`,
				scale: String(scale * (1 - out * 0.04)),
				...style,
			}}
		>
			{children}
		</div>
	);
};
