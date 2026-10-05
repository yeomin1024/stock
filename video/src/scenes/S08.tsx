// VERSION: v1.0.0 — 2026-10-05 — S08 (자막 10–11) 선으로 그린 지갑이 열리면 비어 있음 → "물타기 자금 0원". 자막 11: 비 + 더 어둡게
import React from 'react';
import {AbsoluteFill, random} from 'remotion';
import {sceneTimes} from '../data/timeline';
import {C, alpha} from '../design/colors';
import {easeInOut, prog} from '../design/motion';
import {T} from '../design/type';
import {DrawPath, Svg, useSafeId} from '../components/Draw';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {roundRect} from '../components/hand';

const t = sceneTimes('S08');
const OPEN_AT = t.word(10, '물타기') + 12;
const RAIN_AT = t.sub(11);

const W = {x: 330, y: 230, w: 560, h: 330, r: 30};
const HINGE = W.y + W.h;

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
	const draw = prog(f, 0, 20);
	const open = easeInOut(prog(f, OPEN_AT, OPEN_AT + 22, (x) => x));
	const flapScale = 1 - open * 1.62; // 1 → -0.62 (아래로 젖혀짐)
	const rainIn = prog(f, RAIN_AT, RAIN_AT + 20);
	const dark = prog(f, RAIN_AT, RAIN_AT + 45) * 0.42;
	const rf = f - RAIN_AT;

	return (
		<AbsoluteFill>
			<SceneBg tone="navy" />
			<Layer depth="mid">
				<Svg>
					{/* 지갑 뒷판 + 안쪽 주머니 (비어 있음) */}
					<path d={roundRect(W.x, W.y, W.w, W.h, W.r)} fill={alpha(C.cream, 0.05)} stroke="none" />
					<DrawPath d={roundRect(W.x, W.y, W.w, W.h, W.r)} p={draw} stroke={C.cream} width={6} />
					<g opacity={open}>
						<path d={`M ${W.x + 36} ${W.y + 120} Q ${W.x + W.w / 2} ${W.y + 150} ${W.x + W.w - 36} ${W.y + 120}`} stroke={C.cream} strokeWidth={4} fill="none" />
						<path d={`M ${W.x + 36} ${W.y + 200} Q ${W.x + W.w / 2} ${W.y + 230} ${W.x + W.w - 36} ${W.y + 200}`} stroke={C.cream} strokeWidth={4} fill="none" />
						<rect x={W.x + 22} y={W.y + 22} width={W.w - 44} height={W.h - 44} rx={18} fill="none" stroke={alpha(C.cream, 0.45)} strokeWidth={2} strokeDasharray="8 10" />
					</g>
					{/* 앞 덮개: 아래쪽 경첩을 축으로 젖혀진다 */}
					<g transform={`translate(0 ${HINGE}) scale(1 ${flapScale}) translate(0 ${-HINGE})`}>
						<path d={roundRect(W.x, W.y, W.w, W.h, W.r)} fill={C.navy} stroke="none" opacity={draw > 0.95 ? 1 : 0} />
						<DrawPath d={roundRect(W.x, W.y, W.w, W.h, W.r)} p={draw} stroke={C.cream} width={6} />
						<rect x={W.x + W.w - 120} y={W.y + W.h / 2 - 40} width={140} height={80} rx={20} fill={C.navy} stroke={C.cream} strokeWidth={5} opacity={prog(f, 12, 22)} />
						<circle cx={W.x + W.w - 40} cy={W.y + W.h / 2} r={9} fill={C.cream} opacity={prog(f, 16, 24)} />
					</g>
				</Svg>
			</Layer>
			<Layer depth="fg">
				<Reveal at={OPEN_AT + 14} from="right" style={{left: 1030, top: 270}}>
					<div style={{...T.label, color: C.cream, fontSize: 44}}>물타기 자금</div>
				</Reveal>
				<Reveal at={OPEN_AT + 18} from="right" style={{left: 1016, top: 340}}>
					<div style={{...T.number, color: C.white, fontSize: 240}}>0원</div>
				</Reveal>
			</Layer>
			{/* 자막 11: 화면이 더 어두워지고 가는 선이 비처럼 내린다 (자막 영역 위까지만) */}
			<AbsoluteFill style={{backgroundColor: '#000', opacity: dark}} />
			{rainIn > 0 ? (
				<Svg>
					<defs>
						<linearGradient id={gradId} x1="0" y1="0" x2="0" y2="1">
							<stop offset="0.6" stopColor="#fff" />
							<stop offset="0.76" stopColor="#fff" stopOpacity={0} />
						</linearGradient>
						<mask id={maskId} maskUnits="userSpaceOnUse" x={0} y={0} width={1920} height={1080}>
							<rect x={0} y={0} width={1920} height={1080} fill={`url(#${gradId})`} />
						</mask>
					</defs>
					<g mask={`url(#${maskId})`} opacity={rainIn * 0.5}>
						{DROPS.map((d, i) => {
							const y = ((rf * d.speed + d.phase) % 1100) - 140;
							return (
								<line key={i} x1={d.x} y1={y} x2={d.x - d.len * 0.16} y2={y + d.len} stroke={C.cream} strokeWidth={d.w} strokeLinecap="round" />
							);
						})}
					</g>
				</Svg>
			) : null}
		</AbsoluteFill>
	);
};
