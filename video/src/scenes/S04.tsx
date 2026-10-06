// VERSION: v2.0.0 — 2026-10-06 — S04 (자막 6) S03 마지막 화면의 색이 빠지고 → 찢어진 종이 와이프로 네이비
// v2 속도 규칙: 화면 전체 CSS filter(grayscale/blur) 금지 → 빨강을 회색으로 "색 값 자체"를 보간한다.
// 연결 근거: 반전 예고
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
	const desat = prog(f, 0, 30);
	return (
		<AbsoluteFill>
			<SceneProvider pre={0} camFrom={s03.start - s03shot.start} camDur={s03shot.end - s03shot.start}>
				<Freeze frame={s03.end - s03.start - 1}>
					<S03Content desat={desat} />
				</Freeze>
			</SceneProvider>
			<TornWipe p={lin(f, WIPE_AT, WIPE_AT + WIPE.len)} seed="S04-navy">
				<SceneBg tone="navy" />
			</TornWipe>
		</AbsoluteFill>
	);
};
