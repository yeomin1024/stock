// VERSION: v1.0.0 — 2026-10-05 — 종이 질감(feTurbulence 노이즈 6%) + 아주 약한 비네팅
// 화면에 고정된 정적 레이어 — 내용이 바뀌지 않으므로 컴포지터 레이어로 올려 매 프레임 재래스터를 피한다.
import React from 'react';
import {AbsoluteFill} from 'remotion';

export const PaperTexture: React.FC = () => (
	<AbsoluteFill style={{pointerEvents: 'none', willChange: 'transform'}}>
		<svg width={1920} height={1080} viewBox="0 0 1920 1080" style={{position: 'absolute', inset: 0}}>
			<defs>
				<filter id="paper-grain" x="0" y="0" width="100%" height="100%" colorInterpolationFilters="sRGB">
					<feTurbulence type="fractalNoise" baseFrequency="0.85" numOctaves={3} seed={7} stitchTiles="stitch" />
					<feColorMatrix
						type="matrix"
						values="0.33 0.33 0.33 0 0  0.33 0.33 0.33 0 0  0.33 0.33 0.33 0 0  0 0 0 0 1"
					/>
				</filter>
				<filter id="paper-mottle" x="0" y="0" width="100%" height="100%" colorInterpolationFilters="sRGB">
					<feTurbulence type="fractalNoise" baseFrequency="0.006 0.012" numOctaves={2} seed={3} />
					<feColorMatrix
						type="matrix"
						values="0.33 0.33 0.33 0 0  0.33 0.33 0.33 0 0  0.33 0.33 0.33 0 0  0 0 0 0 1"
					/>
				</filter>
				<radialGradient id="paper-vignette" cx="50%" cy="46%" r="72%">
					<stop offset="0.62" stopColor="#000" stopOpacity={0} />
					<stop offset="1" stopColor="#000" stopOpacity={0.13} />
				</radialGradient>
			</defs>
			<rect width={1920} height={1080} filter="url(#paper-grain)" opacity={0.06} />
			<rect width={1920} height={1080} filter="url(#paper-mottle)" opacity={0.035} />
			<rect width={1920} height={1080} fill="url(#paper-vignette)" />
		</svg>
	</AbsoluteFill>
);
