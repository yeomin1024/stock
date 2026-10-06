// VERSION: v2.0.0 — 2026-10-06 — 장면 컨텍스트 / 카메라 줌
// - useSceneFrame(): 장면 로컬 프레임 (0 = 장면의 첫 자막 시작). 와이프 프리롤 동안은 음수.
// - <Layer>: 장면마다 1.00 → 1.03 으로 천천히 줌인 (가이드 v2 3-4). 줌 중심 = 그래픽 영역(120~800) 중심.
// - <SceneBg>: 크림 / 네이비 / 밝은 크림(S24). from 을 주면 색이 바뀌는 시점을 애니메이션.
import React, {createContext, useContext} from 'react';
import {AbsoluteFill, useCurrentFrame} from 'remotion';
import {C, mix} from '../design/colors';
import {M, lin} from '../design/motion';
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

export const SceneProvider: React.FC<SceneCtxValue & {readonly children: React.ReactNode}> = ({children, ...value}) => (
	<SceneCtx.Provider value={value}>{children}</SceneCtx.Provider>
);

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

export const Layer: React.FC<{readonly children: React.ReactNode; readonly style?: React.CSSProperties}> = ({children, style}) => {
	const p = useCameraP();
	return (
		<AbsoluteFill style={{transformOrigin: '960px 460px', scale: String(1 + M.cameraZoom * p), ...style}}>{children}</AbsoluteFill>
	);
};

export const toneColor = (tone: Tone): string => (tone === 'navy' ? C.navy : tone === 'bright' ? C.paper : C.cream);

/** 장면 바탕 (불투명). 와이프로 들어오는 장면이 이전 장면을 덮도록 각 장면이 직접 칠한다. */
export const SceneBg: React.FC<{readonly tone: Tone; readonly from?: Tone; readonly at?: number; readonly dur?: number}> = ({
	tone,
	from,
	at = 0,
	dur = 20,
}) => {
	const f = useSceneFrame();
	const color = from ? mix(toneColor(from), toneColor(tone), lin(f, at, at + dur)) : toneColor(tone);
	return <AbsoluteFill style={{backgroundColor: color}} />;
};
