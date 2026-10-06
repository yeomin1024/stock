// VERSION: v2.0.0 — 2026-10-06 — S02 (자막 3) 사연자의 계좌 카드 등장 → 비중 막대가 MDB 한 종목으로 100% → "몰빵"(노랑)
// 연결 근거: 전 재산(7,000만 원)을 한 종목에 넣음
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {FACTS} from '../data/facts';
import {sceneTimes} from '../data/timeline';
import {enterP, prog} from '../design/motion';
import {T} from '../design/type';
import {AccountCard} from '../components/Account';
import {useCount} from '../components/Counter';
import {Highlight} from '../components/Highlight';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';

const t = sceneTimes('S02');
export const CARD_POS = {x: 96, y: 200} as const;
const S3 = t.sub(3);
const FILL_AT = S3 + 10;
const MOLBBANG_AT = FILL_AT + 30;

export const S02: React.FC = () => {
	const f = useSceneFrame();
	const cardIn = enterP(f, S3, 16);
	const total = useCount(0, FACTS.story.principalManwon, FILL_AT, 28, 100);
	const bar = prog(f, FILL_AT, FILL_AT + 28);
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer>
				<AccountCard x={CARD_POS.x} y={CARD_POS.y} pnlMode={0} totalManwon={total} bar={bar} opacity={cardIn} dy={(1 - cardIn) * 40} />
				<Reveal at={MOLBBANG_AT} from="up" style={{left: 1440, top: 380}}>
					<div style={{...T.headline, fontSize: 88}}>
						<Highlight at={MOLBBANG_AT + 6}>몰빵</Highlight>
					</div>
				</Reveal>
			</Layer>
		</AbsoluteFill>
	);
};
