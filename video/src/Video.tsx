// VERSION: v2.1.0 — 2026-10-07 — 메인 타임라인: S01–S30 중 S23(엔론) 제외 29장면 + 자막 + 종이 질감
// 모든 장면의 시작/끝은 data/timeline.ts 에서 SRT 자막 번호로 계산된다 (초 하드코딩 없음).
// 와이프로 들어오는 장면은 WIPE.pre 프레임 먼저 시작해 찢어진 종이로 이전 장면을 덮고(새 장면 요소는 자막 시작 뒤에 나옴),
// 그 다음 장면이 와이프면 현재 장면은 와이프가 끝날 때까지 남아 있는다.
import React from 'react';
import {AbsoluteFill, Sequence, useCurrentFrame, useVideoConfig} from 'remotion';
import {SCENES, WIPE, sceneRange, shotRange} from './data/timeline';
import {C} from './design/colors';
import {PaperTexture} from './components/PaperTexture';
import {SceneProvider} from './components/Scene';
import {Subtitles} from './components/Subtitles';
import {TornWipe} from './components/TornWipe';
import {S01} from './scenes/S01';
import {S02} from './scenes/S02';
import {S03} from './scenes/S03';
import {S04} from './scenes/S04';
import {S05} from './scenes/S05';
import {S06} from './scenes/S06';
import {S07} from './scenes/S07';
import {S08} from './scenes/S08';
import {S09} from './scenes/S09';
import {S10} from './scenes/S10';
import {S11} from './scenes/S11';
import {S12} from './scenes/S12';
import {S13} from './scenes/S13';
import {S14} from './scenes/S14';
import {S15} from './scenes/S15';
import {S16} from './scenes/S16';
import {S17} from './scenes/S17';
import {S18} from './scenes/S18';
import {S19} from './scenes/S19';
import {S20} from './scenes/S20';
import {S21} from './scenes/S21';
import {S22} from './scenes/S22';
import {S24} from './scenes/S24';
import {S25} from './scenes/S25';
import {S26} from './scenes/S26';
import {S27} from './scenes/S27';
import {S28} from './scenes/S28';
import {S29} from './scenes/S29';
import {S30} from './scenes/S30';

const WipeIn: React.FC<{readonly seed: string; readonly children: React.ReactNode}> = ({seed, children}) => {
	const f = useCurrentFrame();
	return (
		<TornWipe p={f / WIPE.len} seed={seed}>
			{children}
		</TornWipe>
	);
};

const SceneSlot: React.FC<{readonly id: string; readonly children: React.ReactNode}> = ({id, children}) => {
	const {fps} = useVideoConfig();
	const idx = SCENES.findIndex((s) => s.id === id);
	const def = SCENES[idx];
	const next = SCENES[idx + 1];
	const {start, end} = sceneRange(id);
	const shot = shotRange(id);
	const pre = def.enter === 'wipe' ? WIPE.pre : 0;
	const post = next && next.enter === 'wipe' ? WIPE.len - WIPE.pre : 0;
	const from = start - pre;
	return (
		<Sequence name={id} from={from} durationInFrames={end + post - from} premountFor={fps}>
			<SceneProvider pre={pre} camFrom={start - shot.start} camDur={shot.end - shot.start}>
				{def.enter === 'wipe' ? <WipeIn seed={`${id}-wipe`}>{children}</WipeIn> : children}
			</SceneProvider>
		</Sequence>
	);
};

export const Video: React.FC = () => (
	<AbsoluteFill style={{backgroundColor: C.cream}}>
		<SceneSlot id="S01"><S01 /></SceneSlot>
		<SceneSlot id="S02"><S02 /></SceneSlot>
		<SceneSlot id="S03"><S03 /></SceneSlot>
		<SceneSlot id="S04"><S04 /></SceneSlot>
		<SceneSlot id="S05"><S05 /></SceneSlot>
		<SceneSlot id="S06"><S06 /></SceneSlot>
		<SceneSlot id="S07"><S07 /></SceneSlot>
		<SceneSlot id="S08"><S08 /></SceneSlot>
		<SceneSlot id="S09"><S09 /></SceneSlot>
		<SceneSlot id="S10"><S10 /></SceneSlot>
		<SceneSlot id="S11"><S11 /></SceneSlot>
		<SceneSlot id="S12"><S12 /></SceneSlot>
		<SceneSlot id="S13"><S13 /></SceneSlot>
		<SceneSlot id="S14"><S14 /></SceneSlot>
		<SceneSlot id="S15"><S15 /></SceneSlot>
		<SceneSlot id="S16"><S16 /></SceneSlot>
		<SceneSlot id="S17"><S17 /></SceneSlot>
		<SceneSlot id="S18"><S18 /></SceneSlot>
		<SceneSlot id="S19"><S19 /></SceneSlot>
		<SceneSlot id="S20"><S20 /></SceneSlot>
		<SceneSlot id="S21"><S21 /></SceneSlot>
		<SceneSlot id="S22"><S22 /></SceneSlot>
		<SceneSlot id="S24"><S24 /></SceneSlot>
		<SceneSlot id="S25"><S25 /></SceneSlot>
		<SceneSlot id="S26"><S26 /></SceneSlot>
		<SceneSlot id="S27"><S27 /></SceneSlot>
		<SceneSlot id="S28"><S28 /></SceneSlot>
		<SceneSlot id="S29"><S29 /></SceneSlot>
		<SceneSlot id="S30"><S30 /></SceneSlot>
		<PaperTexture />
		<Subtitles />
	</AbsoluteFill>
);
