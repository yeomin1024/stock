// VERSION: v2.0.0 — 2026-10-06 — S08 (자막 10–11) 선으로 그린 지갑이 열리면 비어 있음 → "물타기 할 현금 0원"
// 자막 11: 화면이 한 단계 더 어두워지고 가는 선이 비처럼 내린다 (자막 영역 위까지만). 연결 근거: 물타기 할 돈이 없음, 한탄
import React from 'react';
import {AbsoluteFill, random} from 'remotion';
import {sceneTimes} from '../data/timeline';
import {C} from '../design/colors';
import {easeInOut, prog} from '../design/motion';
import {T} from '../design/type';
import {Svg, useSafeId} from '../components/Draw';
import {Wallet} from '../components/Icons';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';

const t = sceneTimes('S08');
const S10 = t.sub(10);
const S11 = t.sub(11);
const OPEN_AT = S10 + 18;

export const WALLET = {x: 330, y: 230, w: 560, h: 330} as const;

const DROPS = Array.from({length: 90}, (_, i) => ({
	x: random(`rx${i}`) * 2100 - 80,
	speed: 26 + random(`rs${i}`) * 12,
	len: 50 + random(`rl${i}`) * 70,
	phase: random(`rp${i}`) * 1000,
	w: 1.4 + random(`rw${i}`) * 1.4,
}));

export const S08: React.FC = () => {
	const f = useSceneFrame();
	const maskId = useSafeId('rain-mask');
	const gradId = useSafeId('rain-grad');
	const open = easeInOut(prog(f, OPEN_AT, OPEN_AT + 22, (x) => x));
	const rainIn = prog(f, S11, S11 + 20);
	const dark = prog(f, S11, S11 + 45) * 0.42;
	const rf = f - S11;
	return (
		<AbsoluteFill>
			<SceneBg tone="navy" />
			<Layer>
				<Svg>
					<Wallet {...WALLET} open={open} draw={prog(f, S10, S10 + 18)} line={C.light} bg={C.navy} />
				</Svg>
				<Reveal at={OPEN_AT + 14} from="right" style={{left: 1030, top: 270}}>
					<div style={{...T.label, color: C.light, fontSize: 44}}>물타기 할 현금</div>
				</Reveal>
				<Reveal at={OPEN_AT + 18} from="right" style={{left: 1016, top: 340}}>
					<div style={{...T.number, color: C.light, fontSize: 220}}>0원</div>
				</Reveal>
			</Layer>
			<AbsoluteFill style={{backgroundColor: '#000', opacity: dark}} />
			{rainIn > 0 ? (
				<Svg>
					<defs>
						<linearGradient id={gradId} x1="0" y1="0" x2="0" y2="1">
							<stop offset="0.6" stopColor="#fff" />
							<stop offset="0.74" stopColor="#fff" stopOpacity={0} />
						</linearGradient>
						<mask id={maskId} maskUnits="userSpaceOnUse" x={0} y={0} width={1920} height={1080}>
							<rect x={0} y={0} width={1920} height={1080} fill={`url(#${gradId})`} />
						</mask>
					</defs>
					<g mask={`url(#${maskId})`} opacity={rainIn * 0.5}>
						{DROPS.map((d, i) => {
							const y = ((rf * d.speed + d.phase) % 1100) - 140;
							return <line key={i} x1={d.x} y1={y} x2={d.x - d.len * 0.16} y2={y + d.len} stroke={C.light} strokeWidth={d.w} strokeLinecap="round" />;
						})}
					</g>
				</Svg>
			) : null}
		</AbsoluteFill>
	);
};
