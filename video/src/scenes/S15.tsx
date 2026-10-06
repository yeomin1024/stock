// VERSION: v2.0.0 — 2026-10-06 — S15 (자막 28–29) 헤드라인 "전문가도 미리 알기 어렵다" → 자막 29: 큰 숫자 "2명" 카운트업 + 라벨
// 연결 근거: 전문가도 예측 못 함의 근거
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {FACTS, SRC} from '../data/facts';
import {sceneTimes} from '../data/timeline';
import {C} from '../design/colors';
import {T} from '../design/type';
import {SourceCaption} from '../components/Bits';
import {Counter} from '../components/Counter';
import {Highlight} from '../components/Highlight';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg} from '../components/Scene';

const t = sceneTimes('S15');
const S28 = t.sub(28);
const S29 = t.sub(29);

export const S15: React.FC = () => (
	<AbsoluteFill>
		<SceneBg tone="cream" />
		<Layer>
			<Reveal at={S28} from="up" style={{left: 0, right: 0, top: 170, textAlign: 'center'}}>
				<div style={{...T.headline, fontSize: 84}}>전문가도 미리 알기 어렵다</div>
			</Reveal>
			<Reveal at={S29} from="up" style={{left: 0, right: 0, top: 340, textAlign: 'center'}}>
				<div style={{...T.number, fontSize: 220, color: C.ink}}>
					<Highlight at={S29 + 26} top={0.5}>
						<Counter from={0} to={FACTS.sellAnalysts} at={S29 + 4} dur={20} steps={1} format={(v) => `${Math.round(v)}명`} color={C.ink} style={{fontSize: 220}} />
					</Highlight>
				</div>
			</Reveal>
			<Reveal at={S29 + 6} from="up" style={{left: 0, right: 0, top: 600, textAlign: 'center'}}>
				<div style={{...T.label, fontSize: 44}}>메타 폭락 전 &apos;팔아라&apos; 의견</div>
			</Reveal>
		</Layer>
		<SourceCaption text={SRC.bloomberg} at={S29 + 6} />
	</AbsoluteFill>
);
