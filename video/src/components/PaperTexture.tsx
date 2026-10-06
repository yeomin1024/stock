// VERSION: v2.0.0 — 2026-10-06 — 종이 질감 (가이드 v2 2번 속도 규칙)
// feTurbulence 금지 → 고정된 작은 점 패턴(CSS radial-gradient 타일) + 아주 약한 비네팅 그라데이션.
// 둘 다 정적이고 필터가 없어 프레임마다 다시 계산할 것이 거의 없다.
import React from 'react';
import {AbsoluteFill} from 'remotion';

export const PaperTexture: React.FC = () => (
	<AbsoluteFill style={{pointerEvents: 'none'}}>
		{/* 점 두 겹: 7px / 11px 간격으로 엇갈려 반복 무늬가 덜 보이게 */}
		<AbsoluteFill
			style={{
				backgroundImage:
					'radial-gradient(circle at 1px 1px, rgba(30,30,30,0.055) 0.9px, transparent 1.3px), radial-gradient(circle at 5px 6px, rgba(255,255,255,0.10) 1px, transparent 1.5px)',
				backgroundSize: '7px 7px, 11px 11px',
			}}
		/>
		<AbsoluteFill style={{backgroundImage: 'radial-gradient(ellipse 72% 72% at 50% 46%, rgba(0,0,0,0) 62%, rgba(0,0,0,0.11) 100%)'}} />
	</AbsoluteFill>
);
