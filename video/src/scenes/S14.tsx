// VERSION: v2.3.0 — 2026-10-08 — 한 줄 자막: 뒷줄 내용 요소는 t.line(n, 2) · S14 (자막 23–27) 기업 카드 3개(META, GOOGL, NVDA). 자막 24·25·26 시작에 각 카드의 연도·원인 + 파랑 하락 막대.
// 자막 27: NVDA 카드 아래 "당시 하루 최대 시총 손실" 에 노랑 형광펜. 연결 근거: 빅테크 급락 사례
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {FACTS, SRC} from '../data/facts';
import {sceneTimes} from '../data/timeline';
import {SourceCaption} from '../components/Bits';
import {DropCard} from '../components/Common';
import {Layer, SceneBg} from '../components/Scene';

const t = sceneTimes('S14');
const S23 = t.sub(23);
const DROP_AT = [t.sub(24), t.sub(25), t.line(26, 2)]; // 26 뒷줄 "딥시크의 등장으로 하루에 17%…"
const CX = [400, 960, 1520];

export const S14: React.FC = () => (
	<AbsoluteFill>
		<SceneBg tone="cream" />
		<Layer>
			{FACTS.bigtech.map((c, i) => (
				<DropCard
					key={c.ticker}
					cx={CX[i]}
					top={130}
					w={500}
					h={196}
					title={c.ticker}
					line={`${c.year} · ${c.cause}`}
					at={S23 + i * 6}
					dropAt={DROP_AT[i]}
					pct={c.pct}
					note={c.note || undefined}
					pxPerPct={9}
					numberSize={160}
					mark={c.ticker === 'NVDA' ? {text: FACTS.nvdaRecord, at: t.sub(27)} : undefined}
				/>
			))}
		</Layer>
		<SourceCaption text={SRC.pressBloomberg} at={DROP_AT[0]} />
	</AbsoluteFill>
);
