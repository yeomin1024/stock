// VERSION: v1.0.0 — 2026-10-05 — S05 (자막 7) 계좌 카드 +1,000만 → -3,000만 원 카운트다운, 빨강→파랑, 3프레임 흔들림
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {FACTS, formatManwon} from '../data/facts';
import {sceneTimes} from '../data/timeline';
import {C, pnlColor} from '../design/colors';
import {SANS} from '../design/fonts';
import {enterP} from '../design/motion';
import {T} from '../design/type';
import {cardStyle} from '../components/Bits';
import {Counter} from '../components/Counter';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';

const t = sceneTimes('S05');
const DROP_AT = t.word(7, '갑자기');
const DROP_DUR = 30;
const SHAKE = [-16, 13, -8]; // 3프레임
const SHAKE_AT = DROP_AT + DROP_DUR;

export const S05: React.FC = () => {
	const f = useSceneFrame();
	const cardIn = enterP(f, 0, 16);
	const si = f - SHAKE_AT;
	const shake = si >= 0 && si < SHAKE.length ? SHAKE[si] : 0;
	return (
		<AbsoluteFill>
			<SceneBg tone="navy" />
			<Layer depth="mid">
				<div
					style={{
						...cardStyle(true),
						left: 420,
						top: 170,
						width: 1080,
						height: 470,
						opacity: cardIn,
						translate: `${shake}px ${(1 - cardIn) * 50}px`,
						padding: '48px 64px',
						boxSizing: 'border-box',
					}}
				>
					<div style={{display: 'flex', justifyContent: 'space-between', alignItems: 'baseline'}}>
						<div style={{...T.label, fontSize: 44}}>평가손익</div>
						<div style={{fontFamily: SANS, fontWeight: 700, fontSize: 36, color: C.gray}}>{FACTS.mdb.ticker}</div>
					</div>
					<div style={{marginTop: 96, textAlign: 'center'}}>
						<Counter
							from={FACTS.story.gainManwon}
							to={FACTS.story.lossManwon}
							at={DROP_AT}
							dur={DROP_DUR}
							steps={100}
							format={(v) => formatManwon(v, true)}
							color={(v) => (v === 0 ? C.ink : pnlColor(v))}
							style={{fontSize: 168}}
						/>
					</div>
				</div>
			</Layer>
		</AbsoluteFill>
	);
};
