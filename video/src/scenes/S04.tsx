// VERSION: v1.0.0 — 2026-10-05 — S04 (자막 6) S03 마지막 화면의 채도가 빠지며 흐려지고 → 찢어진 종이 와이프로 네이비
import React from 'react';
import {AbsoluteFill, Freeze} from 'remotion';
import {S04_NAVY_WIPE_AT, WIPE, sceneRange, shotRange} from '../data/timeline';
import {lin, prog} from '../design/motion';
import {SceneBg, SceneProvider, useSceneFrame} from '../components/Scene';
import {TornWipe} from '../components/TornWipe';
import {S03Content} from './S03';

const s03 = sceneRange('S03');
const s03shot = shotRange('S03');
const WIPE_AT = S04_NAVY_WIPE_AT();

export const S04: React.FC = () => {
	const f = useSceneFrame();
	const d = prog(f, 0, 36);
	return (
		<AbsoluteFill>
			<AbsoluteFill style={{filter: `grayscale(${d}) blur(${(d * 3).toFixed(2)}px) brightness(${1 - d * 0.06})`}}>
				<SceneProvider pre={0} camFrom={s03.start - s03shot.start} camDur={s03shot.end - s03shot.start}>
					<Freeze frame={s03.end - s03.start - 1}>
						<S03Content />
					</Freeze>
				</SceneProvider>
			</AbsoluteFill>
			<TornWipe p={lin(f, WIPE_AT, WIPE_AT + WIPE.len)} seed="S04-navy">
				<SceneBg tone="navy" />
			</TornWipe>
		</AbsoluteFill>
	);
};
