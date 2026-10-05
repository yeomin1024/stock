// VERSION: v1.0.0 — 2026-10-05 — 장면 컨텍스트 / 카메라 줌 / 패럴랙스 레이어
// - useSceneFrame(): 장면 로컬 프레임 (0 = 장면의 첫 자막 시작). 와이프 프리롤 동안은 음수.
// - <Layer depth>: 카메라 1.00→1.04 줌. bg/mid/fg 가 서로 다른 속도로 움직여 패럴랙스를 만든다.
//   줌 중심은 그래픽 영역(y<820)의 중심(960, 410) — 자막 영역 쪽으로 그래픽이 밀려나지 않게.
import React, {createContext, useContext} from 'react';
import {AbsoluteFill, useCurrentFrame} from 'remotion';
import {C} from '../design/colors';
import {M} from '../design/motion';
import type {Tone} from '../data/timeline';

type SceneCtxValue = {
	/** 장면 Sequence 시작이 장면 시작보다 몇 프레임 앞서는지 (와이프 프리롤) */
	readonly pre: number;
	/** 카메라 샷 시작 기준으로 이 장면이 몇 프레임 뒤에 시작하는지 */
	readonly camFrom: number;
	/** 카메라 샷 전체 길이 */
	readonly camDur: number;
};

const SceneCtx = createContext<SceneCtxValue>({pre: 0, camFrom: 0, camDur: 1});

export const SceneProvider: React.FC<SceneCtxValue & {readonly children: React.ReactNode}> = ({
	children,
	...value
}) => <SceneCtx.Provider value={value}>{children}</SceneCtx.Provider>;

/** 장면 로컬 프레임 (자막 기준 0). */
export const useSceneFrame = (): number => {
	const frame = useCurrentFrame();
	const {pre} = useContext(SceneCtx);
	return frame - pre;
};

/** 카메라 진행도 0→1 (샷 단위, 선형 = "천천히") */
export const useCameraP = (): number => {
	const f = useSceneFrame();
	const {camFrom, camDur} = useContext(SceneCtx);
	return Math.min(1, Math.max(0, (f + camFrom) / camDur));
};

const DEPTH = {bg: 0.5, mid: 1, fg: 1.5} as const;
const DRIFT = {bg: 3, mid: 0, fg: -6} as const;

export const Layer: React.FC<{
	readonly depth?: keyof typeof DEPTH;
	readonly children: React.ReactNode;
	readonly style?: React.CSSProperties;
}> = ({depth = 'mid', children, style}) => {
	const p = useCameraP();
	return (
		<AbsoluteFill
			style={{
				transformOrigin: '960px 410px',
				scale: String(1 + M.cameraZoom * DEPTH[depth] * p),
				translate: `0px ${DRIFT[depth] * p}px`,
				...style,
			}}
		>
			{children}
		</AbsoluteFill>
	);
};

/** 장면 바탕 (불투명). 와이프 시 들어오는 장면이 이전 장면을 덮도록 각 장면이 직접 칠한다. */
export const SceneBg: React.FC<{readonly tone: Tone}> = ({tone}) => (
	<AbsoluteFill style={{backgroundColor: tone === 'navy' ? C.navy : C.cream}} />
);
