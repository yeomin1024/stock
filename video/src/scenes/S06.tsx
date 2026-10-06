// VERSION: v2.0.0 — 2026-10-06 — S06 (자막 8) 텍스트로만 만든 뉴스 헤드라인 카드가 슬라이드 인
// 연결 근거: 손실의 원인이 된 소식
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {FACTS, SRC} from '../data/facts';
import {sceneTimes} from '../data/timeline';
import {C, alpha} from '../design/colors';
import {enterP} from '../design/motion';
import {SourceCaption} from '../components/Bits';
import {HeadlineCard} from '../components/Common';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';

const t = sceneTimes('S06');
const S8 = t.sub(8);

export const S06: React.FC = () => {
	const f = useSceneFrame();
	return (
		<AbsoluteFill>
			<SceneBg tone="navy" />
			<Layer>
				<HeadlineCard x={310} y={250} w={1300} date={FACTS.story.ceoNewsDate} title="CEO, 다른 회사로 이적" p={enterP(f, S8, 18)} />
			</Layer>
			<SourceCaption text={SRC.press} at={S8 + 10} color={alpha(C.light, 0.65)} />
		</AbsoluteFill>
	);
};
