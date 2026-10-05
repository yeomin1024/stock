// VERSION: v1.0.0 — 2026-10-05 — S24 (자막 40–43) 핵심 키워드 카드 3개 → "대처가 가능한 투자" 하이라이트 → 채널명 엔딩 → 페이드아웃
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {useChannelName} from '../data/channel';
import {sceneTimes} from '../data/timeline';
import {C} from '../design/colors';
import {SANS} from '../design/fonts';
import {easeInOut, enterP, exitP, lerp, lin, prog} from '../design/motion';
import {T} from '../design/type';
import {cardStyle} from '../components/Bits';
import {BandMessage} from '../components/ChannelBand';
import {Svg} from '../components/Draw';
import {Highlight} from '../components/Highlight';
import {Lightning} from '../components/Icons';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';

const t = sceneTimes('S24');
const SUM_AT = t.sub(41);
const END_AT = t.sub(42);
const FADE_AT = t.subEnd(43);

const CW = 500;
const CH = 320;
const CY = 180;

const IconRisk: React.FC<{readonly cx: number}> = ({cx}) => (
	<Svg>
		<rect x={cx - 60} y={CY + 140} width={120} height={64} rx={10} fill={C.paper} stroke={C.ink} strokeWidth={4} />
		<Lightning x={cx + 6} y={CY + 22} h={120} rotate={6} />
	</Svg>
);
const IconSpread: React.FC<{readonly cx: number}> = ({cx}) => (
	<Svg>
		{[0, 1, 2, 3, 4].map((i) => (
			<rect key={i} x={cx - 150 + i * 62} y={CY + 110} width={50} height={50} rx={8} fill={C.paper} stroke={C.ink} strokeWidth={4} />
		))}
	</Svg>
);
const IconWeight: React.FC<{readonly cx: number}> = ({cx}) => (
	<Svg>
		{[0, 1, 2, 3, 4].map((i) => (
			<rect key={i} x={cx - 140 + i * 58} y={CY + 130} width={42} height={74} rx={6} fill={C.paper} stroke={C.ink} strokeWidth={4} />
		))}
		<line x1={cx - 170} y1={CY + 108} x2={cx + 170} y2={CY + 108} stroke={C.ink} strokeWidth={4} strokeDasharray="12 9" />
	</Svg>
);

const CARDS = [
	{label: '몰빵의 위험', cx: 380, at: t.sub(40) + 2, Icon: IconRisk},
	{label: '분산', cx: 960, at: t.word(40, '분산'), Icon: IconSpread},
	{label: '비중 조절', cx: 1540, at: t.word(40, '분산') + 16, Icon: IconWeight},
];

export const S24: React.FC = () => {
	const f = useSceneFrame();
	const name = useChannelName();
	const lift = easeInOut(prog(f, SUM_AT, SUM_AT + 18, (x) => x));
	const out = exitP(f, END_AT, 9);
	const fade = lin(f, FADE_AT, FADE_AT + 28);
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer depth="mid">
				{out > 0.001 ? (
					<AbsoluteFill style={{opacity: out}}>
						{CARDS.map((c) => {
							const p = enterP(f, c.at, 15);
							if (p <= 0.001) return null;
							return (
								<AbsoluteFill
									key={c.label}
									style={{
										opacity: p,
										translate: `0px ${(1 - p) * 40 + lerp(0, -76, lift)}px`,
										scale: String(lerp(1, 0.76, lift)),
										transformOrigin: `${c.cx}px ${CY}px`,
									}}
								>
									<div style={{...cardStyle(), left: c.cx - CW / 2, top: CY, width: CW, height: CH}} />
									<c.Icon cx={c.cx} />
									<div
										style={{
											position: 'absolute',
											left: c.cx - CW / 2,
											width: CW,
											top: CY + 236,
											textAlign: 'center',
											fontFamily: SANS,
											fontWeight: 900,
											fontSize: 44,
											color: C.ink,
										}}
									>
										{c.label}
									</div>
								</AbsoluteFill>
							);
						})}
						<Reveal at={SUM_AT + 10} from="up" style={{left: 0, right: 0, top: 470, textAlign: 'center'}}>
							<div style={{...T.headline, fontSize: 96}}>
								<Highlight at={t.word(41, '대처가')}>대처가 가능한 투자</Highlight>
							</div>
						</Reveal>
					</AbsoluteFill>
				) : null}
				<BandMessage at={END_AT + 4} text={name} seed="S24-band" fontSize={96} />
			</Layer>
			{/* 엔딩 페이드아웃 */}
			<AbsoluteFill style={{backgroundColor: C.ink, opacity: fade}} />
		</AbsoluteFill>
	);
};
