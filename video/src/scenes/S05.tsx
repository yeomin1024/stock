// VERSION: v2.0.0 — 2026-10-06 — S05 (자막 7) 같은 계좌 카드: +1,000만 → -2,000만 원, 빨강 → 파랑, 3프레임 흔들림
// 연결 근거: 계좌에 찍힌 손실
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {FACTS} from '../data/facts';
import {sceneTimes} from '../data/timeline';
import {enterP} from '../design/motion';
import {ACCOUNT_CARD, AccountCard} from '../components/Account';
import {useCount} from '../components/Counter';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';

const t = sceneTimes('S05');
const S7 = t.sub(7);
const DROP_AT = S7 + 16;
const DROP_DUR = 30;
const SHAKE = [-16, 13, -8]; // 3프레임
const SHAKE_AT = DROP_AT + DROP_DUR;

export const S05: React.FC = () => {
	const f = useSceneFrame();
	const cardIn = enterP(f, S7, 16);
	const pnl = useCount(FACTS.story.gainManwon, FACTS.story.lossManwon, DROP_AT, DROP_DUR, 100);
	const si = f - SHAKE_AT;
	const shake = si >= 0 && si < SHAKE.length ? SHAKE[si] : 0;
	return (
		<AbsoluteFill>
			<SceneBg tone="navy" />
			<Layer>
				<AccountCard
					x={(1920 - ACCOUNT_CARD.w) / 2}
					y={200}
					pnlMode={1}
					totalManwon={FACTS.story.principalManwon}
					pnlManwon={pnl}
					bar={1}
					dark
					opacity={cardIn}
					dx={shake}
					dy={(1 - cardIn) * 40}
				/>
			</Layer>
		</AbsoluteFill>
	);
};
