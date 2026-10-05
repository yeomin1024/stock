// VERSION: v1.0.0 — 2026-10-05 — S09 (자막 12) 찢어진 노랑 종이 띠가 펼쳐지고 채널명 등장
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {useChannelName} from '../data/channel';
import {BandMessage} from '../components/ChannelBand';
import {Layer, SceneBg} from '../components/Scene';

export const S09: React.FC = () => {
	const name = useChannelName();
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer depth="mid">
				<BandMessage at={0} text={name} seed="S09-band" fontSize={96} />
			</Layer>
		</AbsoluteFill>
	);
};
