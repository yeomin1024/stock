// VERSION: v2.0.0 — 2026-10-06 — S09 (자막 12–15) "앞의 사연과 같이": 사연 장면(S08 마지막 화면)이 작게 축소되어 왼쪽으로 빠지고,
// 자막 13·14·15 시작에 맞춰 생각 말풍선 3개가 하나씩. 인물 그림 없음. 연결 근거: 몰빵하는 사람들의 생각 3가지
import React from 'react';
import {AbsoluteFill, Freeze} from 'remotion';
import {sceneRange, sceneTimes, shotRange} from '../data/timeline';
import {C} from '../design/colors';
import {easeInOut, prog} from '../design/motion';
import {ThoughtBubble} from '../components/Common';
import {Layer, SceneBg, SceneProvider} from '../components/Scene';
import {useSceneFrame} from '../components/Scene';
import {S08} from './S08';

const t = sceneTimes('S09');
const s08 = sceneRange('S08');
const s08shot = shotRange('S08');
const S12 = t.sub(12);

export const BUBBLES = [
	{cx: 560, cy: 290, rx: 330, ry: 112, text: '몰빵하면 많이 오르겠지?', at: t.sub(13), tail: 'left' as const, seed: 1},
	{cx: 1360, cy: 330, rx: 320, ry: 112, text: '나누면 수익이 줄어', at: t.sub(14), tail: 'right' as const, seed: 2},
	{cx: 760, cy: 620, rx: 330, ry: 110, text: '떨어져도 금방 회복', at: t.sub(15), tail: 'left' as const, seed: 3},
];

export const S09: React.FC = () => {
	const f = useSceneFrame();
	const shrink = easeInOut(prog(f, S12, S12 + 20, (x) => x));
	const slide = easeInOut(prog(f, S12 + 18, S12 + 44, (x) => x));
	const scale = 1 - 0.6 * shrink;
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			{slide < 1 ? (
				<AbsoluteFill
					style={{
						scale: String(scale),
						translate: `${-1500 * slide}px ${-40 * shrink}px`,
						borderRadius: 40 * shrink,
						overflow: 'hidden',
						outline: shrink > 0.01 ? `${8 / Math.max(scale, 0.01)}px solid ${C.ink}` : undefined,
					}}
				>
					<SceneProvider pre={0} camFrom={s08.start - s08shot.start} camDur={s08shot.end - s08shot.start}>
						<Freeze frame={s08.end - s08.start - 1}>
							<S08 />
						</Freeze>
					</SceneProvider>
				</AbsoluteFill>
			) : null}
			<Layer>
				{BUBBLES.map((b) => (
					<ThoughtBubble key={b.seed} {...b} />
				))}
			</Layer>
		</AbsoluteFill>
	);
};
