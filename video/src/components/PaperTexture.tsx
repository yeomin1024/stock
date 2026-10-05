// VERSION: v1.1.0 — 2026-10-05 — 종이 질감(feTurbulence 노이즈 6%) + 아주 약한 비네팅
// v1.1.0: 같은 feTurbulence SVG 를 인라인 요소 대신 CSS 배경(data URI)으로 쓴다.
//   인라인 필터는 매 프레임 다시 계산돼 4K 에서 프레임당 약 1초가 들었다(90f 벤치: 121s → 텍스처 없이 30s).
//   배경 이미지로 쓰면 브라우저가 한 번 래스터한 결과를 재사용한다. 노이즈 파라미터(주파수/옥타브/시드/불투명도)는 그대로.
//   외부 이미지 파일이 아니라 코드에서 만든 SVG 문자열이다.
import React from 'react';
import {AbsoluteFill} from 'remotion';

const GRAY = '0.33 0.33 0.33 0 0  0.33 0.33 0.33 0 0  0.33 0.33 0.33 0 0  0 0 0 0 1';

const noiseSvg = (w: number, h: number, freq: string, octaves: number, seed: number, stitch: boolean): string =>
	`<svg xmlns='http://www.w3.org/2000/svg' width='${w}' height='${h}' viewBox='0 0 ${w} ${h}'>` +
	`<filter id='n' x='0' y='0' width='100%' height='100%' color-interpolation-filters='sRGB'>` +
	`<feTurbulence type='fractalNoise' baseFrequency='${freq}' numOctaves='${octaves}' seed='${seed}'${stitch ? " stitchTiles='stitch'" : ''}/>` +
	`<feColorMatrix type='matrix' values='${GRAY}'/></filter>` +
	`<rect width='${w}' height='${h}' filter='url(#n)'/></svg>`;

const asUrl = (svg: string): string => `url("data:image/svg+xml;utf8,${encodeURIComponent(svg)}")`;

// 고주파 결: 512px 타일 반복 (stitchTiles 로 이음매 없음)
const GRAIN = asUrl(noiseSvg(512, 512, '0.85', 3, 7, true));
// 저주파 얼룩: 화면 한 장 크기 (반복하면 눈에 띄므로 타일링하지 않음)
const MOTTLE = asUrl(noiseSvg(1920, 1080, '0.006 0.012', 2, 3, false));

export const PaperTexture: React.FC = () => (
	<AbsoluteFill style={{pointerEvents: 'none'}}>
		<AbsoluteFill style={{backgroundImage: GRAIN, backgroundSize: '512px 512px', opacity: 0.06}} />
		<AbsoluteFill style={{backgroundImage: MOTTLE, backgroundSize: '1920px 1080px', opacity: 0.035}} />
		<AbsoluteFill
			style={{
				backgroundImage: 'radial-gradient(ellipse 72% 72% at 50% 46%, rgba(0,0,0,0) 62%, rgba(0,0,0,0.13) 100%)',
			}}
		/>
	</AbsoluteFill>
);
