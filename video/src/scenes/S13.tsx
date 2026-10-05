// VERSION: v1.0.0 — 2026-10-05 — S13 (고지 카드 3초, 자막 19 직후) 노랑 띠 카드에 "투자 권유가 아닙니다"
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {BandMessage} from '../components/ChannelBand';
import {Layer, SceneBg} from '../components/Scene';

export const S13: React.FC = () => (
	<AbsoluteFill>
		<SceneBg tone="cream" />
		<Layer depth="mid">
			<BandMessage at={0} text="투자 권유가 아닙니다" seed="S13-band" fontSize={96} />
		</Layer>
	</AbsoluteFill>
);
