// VERSION: v2.0.0 — 2026-10-06 — S17 (자막 34–38) "02" + "버텨도 회복이 어렵다"
// 자막 36: 사연자의 계좌 막대 7,000만 → 5,000만 (잃은 부분 파랑 점선) + 원래 높이로 올라가는 화살표 "+40% 필요"
// 자막 37: 막대 3개(사연 -29%, 반토막 -50%, -70%)와 필요한 상승(+40%, 2배, 3.3배). 자막 38: 막대 전체 테두리 "계좌 전체"
// 연결 근거: 손실의 비대칭. 막대 높이는 실제 비율(71% / 50% / 30%).
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {FACTS, formatPct} from '../data/facts';
import {sceneTimes} from '../data/timeline';
import {C} from '../design/colors';
import {easeInOut, enterP, prog} from '../design/motion';
import {T} from '../design/type';
import {NumberTitle} from '../components/Bits';
import {StackBar} from '../components/Account';
import {DrawPath, Svg} from '../components/Draw';
import {Highlight} from '../components/Highlight';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {arrowHead, roundRect} from '../components/hand';

const t = sceneTimes('S17');
const S34 = t.sub(34);
const S36 = t.sub(36);
const S37 = t.sub(37);
const S38 = t.sub(38);

const BOTTOM = 742;
const FULL = 440;
const W = 170;
const CXS = [640, 1060, 1480];
const R = FACTS.recovery;

const Column: React.FC<{readonly i: number; readonly at: number; readonly dropAt: number; readonly showCaption: boolean}> = ({i, at, dropAt, showCaption}) => {
	const f = useSceneFrame();
	const r = R[i];
	const appear = enterP(f, at, 14);
	if (appear <= 0.001) return null;
	const drop = easeInOut(prog(f, dropAt, dropAt + 22, (x) => x));
	const lost = (FULL * Math.abs(r.dropPct) * drop) / 100;
	const remainTop = BOTTOM - FULL + lost;
	const arrowP = prog(f, dropAt + 24, dropAt + 40);
	const cx = CXS[i];
	return (
		<>
			<Svg>
				<g opacity={appear}>
					<StackBar x={cx - W / 2} bottom={BOTTOM} w={W} segs={[{h: FULL, kind: 'mdb', lost}]} />
					{lost > 4 ? (
						<>
							<DrawPath d={`M ${cx + W / 2 + 34} ${remainTop - 6} L ${cx + W / 2 + 34} ${BOTTOM - FULL + 24}`} p={arrowP} dash="10 8" width={6} stroke={C.red} linecap="butt" />
							<DrawPath d={arrowHead([cx + W / 2 + 34, BOTTOM - FULL + 14], -Math.PI / 2, 22)} p={prog(f, dropAt + 38, dropAt + 44)} width={6} stroke={C.red} />
						</>
					) : null}
				</g>
			</Svg>
			<Reveal at={dropAt + 26} from="up" dist={16} style={{left: cx - 230, width: 460, top: 118, textAlign: 'center'}}>
				<div style={{...T.number, fontSize: 160, color: C.red}}>{r.needLabel}</div>
			</Reveal>
			{showCaption ? (
				<Reveal at={at + 4} from="none" style={{left: cx - 200, width: 400, top: BOTTOM + 8, textAlign: 'center'}}>
					<div style={{...T.label, fontSize: 36}}>
						{r.caption ? `${r.caption} ` : ''}
						<span style={{color: C.blue, fontVariantNumeric: 'tabular-nums'}}>{formatPct(r.dropPct)}</span>
					</div>
				</Reveal>
			) : null}
		</>
	);
};

export const S17: React.FC = () => {
	const f = useSceneFrame();
	const border = prog(f, S38, S38 + 20);
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer>
				<NumberTitle num="02" title="버텨도 회복이 어렵다" mark="회복이 어렵다" at={S34} markAt={S34 + 24} exitAt={S36 - 2} />
				{/* 자막 36: 사연자의 계좌 막대 (자막 37 부터 아래 라벨 "사연 -29%") */}
				<Column i={0} at={S36} dropAt={S36 + 12} showCaption={f >= S37} />
				<Reveal at={S36 + 40} exitAt={S37} from="none" style={{left: CXS[0] + 236, top: 196}}>
					<div style={{...T.label, color: C.red}}>필요</div>
				</Reveal>
				<Column i={1} at={S37} dropAt={S37 + 6} showCaption />
				<Column i={2} at={S37 + 6} dropAt={S37 + 12} showCaption />
				{/* 자막 38: 막대 전체를 감싸는 테두리 + "계좌 전체" */}
				<Svg>
					<DrawPath d={roundRect(CXS[0] - W / 2 - 60, BOTTOM - FULL - 14, CXS[2] - CXS[0] + W + 170, 800 - (BOTTOM - FULL - 14), 24)} p={border} width={6} stroke={C.ink} />
				</Svg>
				<Reveal at={S38 + 12} from="right" style={{left: 150, top: 470}}>
					<div style={{...T.label, fontSize: 44}}>
						<Highlight at={S38 + 20}>계좌 전체</Highlight>
					</div>
				</Reveal>
			</Layer>
		</AbsoluteFill>
	);
};
