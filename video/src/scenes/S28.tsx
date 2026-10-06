// VERSION: v2.0.0 — 2026-10-06 — S28 (자막 80–84) "넷째 · 나눠 사고 현금 남기기"
// 81: 700만 원 블록이 3조각 "700만 원 → 3번" → 82: 첫 조각만 매수된 뒤 번개, 남은 2조각(현금)에서 "더 싸게 사기"·"기다리기" 화살표
// 83: S08의 빈 지갑이 작게 다시 나와 "현금 0원" → 84: 현금 2조각에 "현금 = 위기 때 쓸 선택지"
// 연결 근거: 분할 매수와 현금의 역할
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {FACTS, formatManwon} from '../data/facts';
import {sceneTimes} from '../data/timeline';
import {C, mix} from '../design/colors';
import {SANS} from '../design/fonts';
import {easeInOut, enterP, exitP, lerp, lin, prog} from '../design/motion';
import {T} from '../design/type';
import {SceneTitle} from '../components/Bits';
import {DrawPath, Svg} from '../components/Draw';
import {Highlight} from '../components/Highlight';
import {Lightning, Wallet} from '../components/Icons';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {handArrow} from '../components/hand';

const t = sceneTimes('S28');
const [S80, S81, S82, S83, S84] = [80, 81, 82, 83, 84].map((n) => t.sub(n));
const SP = FACTS.split;

const BLOCK = {x: 610, y: 380, w: 700, h: 140};
const PW = 210;
const PG = 35;
const PIECE_X = [0, 1, 2].map((i) => BLOCK.x + i * (PW + PG));
const LOSS = Math.abs(FACTS.story.lossPct) / 100;
const ARROW_CHEAP = handArrow([PIECE_X[1] + PW / 2, BLOCK.y + BLOCK.h + 14], [PIECE_X[1] + PW / 2 - 70, BLOCK.y + BLOCK.h + 120], 's28a', 0.2, 22);
const ARROW_WAIT = handArrow([PIECE_X[2] + PW / 2, BLOCK.y + BLOCK.h + 14], [PIECE_X[2] + PW / 2 + 90, BLOCK.y + BLOCK.h + 120], 's28b', -0.2, 22);
const PANEL = {x: 1440, y: 250, w: 380, h: 300};

export const S28: React.FC = () => {
	const f = useSceneFrame();
	const blockIn = enterP(f, S81, 14);
	const split = easeInOut(prog(f, S81 + 20, S81 + 40, (x) => x));
	const bought = lin(f, S82 + 4, S82 + 12);
	const strike = prog(f, S82 + 22, S82 + 27, (x) => x);
	const lost = easeInOut(prog(f, S82 + 26, S82 + 42, (x) => x)) * BLOCK.h * LOSS;
	const panel = Math.min(enterP(f, S83 + 4, 15), exitP(f, S84 - 2, 9));
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer>
				<SceneTitle text="넷째 · 나눠 사고 현금 남기기" at={S80} />
				<Svg>
					<g opacity={blockIn}>
						{PIECE_X.map((px, i) => {
							const x = lerp(BLOCK.x + (i * BLOCK.w) / 3, px, split);
							const w = lerp(BLOCK.w / 3, PW, split);
							const isFirst = i === 0;
							const fill = isFirst ? mix(C.paper, C.ink, bought) : C.paper;
							const l = isFirst ? lost : 0;
							return (
								<g key={i}>
									{l > 0.5 ? <rect x={x} y={BLOCK.y} width={w} height={l} fill="rgba(45,108,223,0.12)" stroke={C.blue} strokeWidth={3} strokeDasharray="9 7" /> : null}
									<rect x={x} y={BLOCK.y + l} width={w} height={BLOCK.h - l} rx={split > 0.5 ? 12 : 0} fill={fill} stroke={C.ink} strokeWidth={5} />
								</g>
							);
						})}
					</g>
					<Lightning x={PIECE_X[0] + PW / 2 + 8} y={BLOCK.y - 150} h={140} reveal={strike} opacity={1 - prog(f, S83 - 10, S83)} />
					<DrawPath d={ARROW_CHEAP.shaft} p={prog(f, S82 + 40, S82 + 54)} width={6} />
					<DrawPath d={ARROW_CHEAP.head} p={prog(f, S82 + 52, S82 + 58)} width={6} />
					<DrawPath d={ARROW_WAIT.shaft} p={prog(f, S82 + 46, S82 + 60)} width={6} />
					<DrawPath d={ARROW_WAIT.head} p={prog(f, S82 + 58, S82 + 64)} width={6} />
				</Svg>
				{/* 블록 위 글자: 81 "700만 원" (나뉘기 전) / 82 이후 현금 조각 "현금" */}
				<div style={{position: 'absolute', left: BLOCK.x, top: BLOCK.y, width: BLOCK.w, height: BLOCK.h, display: 'flex', alignItems: 'center', justifyContent: 'center', fontFamily: SANS, fontWeight: 900, fontSize: 60, color: C.ink, opacity: blockIn * (1 - split)}}>
					{formatManwon(SP.perStockManwon)}
				</div>
				{[1, 2].map((i) => (
					<div key={i} style={{position: 'absolute', left: PIECE_X[i], top: BLOCK.y, width: PW, height: BLOCK.h, display: 'flex', alignItems: 'center', justifyContent: 'center', ...T.label, opacity: lin(f, S82 + 10, S82 + 20)}}>
						현금
					</div>
				))}
				<Reveal at={S81 + 34} exitAt={S82 - 2} from="up" style={{left: 0, right: 0, top: 560, textAlign: 'center'}}>
					<div style={{...T.label, fontSize: 44}}>
						{formatManwon(SP.perStockManwon)} → {SP.times}번
					</div>
				</Reveal>
				<Reveal at={S82 + 54} from="up" dist={14} style={{left: PIECE_X[1] + PW / 2 - 330, width: 300, top: BLOCK.y + BLOCK.h + 120, textAlign: 'right'}}>
					<div style={{...T.label}}>더 싸게 사기</div>
				</Reveal>
				<Reveal at={S82 + 60} from="up" dist={14} style={{left: PIECE_X[2] + PW / 2 + 100, top: BLOCK.y + BLOCK.h + 120}}>
					<div style={{...T.label}}>기다리기</div>
				</Reveal>
				{/* 83: S08 의 빈 지갑 (같은 모양·같은 색: 네이비 위 밝은 선) */}
				{panel > 0.001 ? (
					<div style={{position: 'absolute', left: PANEL.x, top: PANEL.y, width: PANEL.w, height: PANEL.h, borderRadius: 24, background: C.navy, opacity: panel, scale: String(0.9 + 0.1 * panel), overflow: 'hidden'}}>
						<Svg style={{left: -PANEL.x, top: -PANEL.y}}>
							<Wallet x={PANEL.x + 85} y={PANEL.y + 30} w={210} h={110} open={1} line={C.light} bg={C.navy} sw={5} />
						</Svg>
						<div style={{position: 'absolute', left: 0, right: 0, bottom: 22, textAlign: 'center', ...T.label, color: C.light}}>현금 0원</div>
					</div>
				) : null}
				<Reveal at={S84 + 6} from="up" style={{left: PIECE_X[1] - 40, width: 2 * PW + PG + 80, top: 290, textAlign: 'center'}}>
					<div style={{...T.label, fontSize: 44}}>
						현금 = 위기 때 쓸 <Highlight at={S84 + 16}>선택지</Highlight>
					</div>
				</Reveal>
			</Layer>
		</AbsoluteFill>
	);
};
