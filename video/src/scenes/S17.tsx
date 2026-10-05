// VERSION: v1.0.0 — 2026-10-05 — S17 (자막 27–29) "02 확률 자체가 낮다" → 1926→2016 타임라인 + 약 2만 6천 개 → 와플 58칸 파랑
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {FACTS, SRC, formatPct, formatThousandStocks} from '../data/facts';
import {sceneTimes} from '../data/timeline';
import {C, alpha} from '../design/colors';
import {SANS} from '../design/fonts';
import {enterP, lin, prog} from '../design/motion';
import {T} from '../design/type';
import {SourceCaption} from '../components/Bits';
import {Counter} from '../components/Counter';
import {DrawPath, Svg} from '../components/Draw';
import {Highlight} from '../components/Highlight';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {Waffle} from '../components/Waffle';
import {handLine} from '../components/hand';

const t = sceneTimes('S17');
const B = FACTS.bessembinder;
const TL_AT = t.sub(28);
const WAFFLE_AT = t.sub(29);
export const WAFFLE = {x: 200, y: 140, cell: 50, gap: 6} as const;
export const FILL_AT_S17 = WAFFLE_AT + 10;
export const EMPTY_CELL = alpha(C.ink, 0.09);

export const WafflePct: React.FC<{readonly fillAt: number; readonly appearAt: number}> = ({fillAt, appearAt}) => {
	const f = useSceneFrame();
	return (
		<Waffle
			{...WAFFLE}
			cellStyle={(i) => {
				const appear = enterP(f, appearAt + Math.floor(i / 10) * 1.5, 10);
				const blue = i < B.belowTbillPct ? lin(f, fillAt + i * 0.5, fillAt + i * 0.5 + 4) : 0;
				return {fill: blue > 0 ? C.blue : EMPTY_CELL, opacity: appear * (blue > 0 ? 0.3 + 0.7 * blue : 1), scale: 0.86 + 0.14 * appear};
			}}
		/>
	);
};

export const S17: React.FC = () => {
	const f = useSceneFrame();
	const lineP = prog(f, TL_AT + 4, TL_AT + 28);
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer depth="mid">
				{/* 자막 27: 섹션 번호 + 헤드라인 */}
				<Reveal at={0} exitAt={TL_AT - 4} from="up" style={{left: 150, top: 96}}>
					<div style={{...T.number, fontSize: 230, color: C.ink, letterSpacing: '-0.04em'}}>02</div>
				</Reveal>
				<Reveal at={6} exitAt={TL_AT - 4} from="left" style={{left: 160, top: 340}}>
					<div style={{...T.headline, fontSize: 92}}>
						<Highlight at={t.word(27, '확률')}>확률 자체가 낮다</Highlight>
					</div>
				</Reveal>

				{/* 자막 28: 1926 → 2016 타임라인, 약 2만 6천 개 */}
				{f >= TL_AT && f < WAFFLE_AT + 12 ? (
					<>
						<Svg>
							<g opacity={1 - prog(f, WAFFLE_AT - 4, WAFFLE_AT + 5)}>
								<DrawPath d={handLine(320, 330, 1600, 330, 's17tl', 1.5)} p={lineP} width={7} />
								<DrawPath d="M 320 306 L 320 354" p={prog(f, TL_AT, TL_AT + 6)} width={7} />
								<DrawPath d="M 1600 306 L 1600 354" p={prog(f, TL_AT + 26, TL_AT + 32)} width={7} />
							</g>
						</Svg>
					</>
				) : null}
				<Reveal at={TL_AT} exitAt={WAFFLE_AT - 4} from="up" dist={20} style={{left: 320 - 110, top: 200, width: 220, textAlign: 'center'}}>
					<div style={{...T.number, fontSize: 64, color: C.ink}}>{B.fromYear}</div>
				</Reveal>
				<Reveal at={TL_AT + 26} exitAt={WAFFLE_AT - 4} from="up" dist={20} style={{left: 1600 - 110, top: 200, width: 220, textAlign: 'center'}}>
					<div style={{...T.number, fontSize: 64, color: C.ink}}>{B.toYear}</div>
				</Reveal>
				<Reveal at={TL_AT + 14} exitAt={WAFFLE_AT - 4} from="up" dist={20} style={{left: 960 - 120, top: 214, width: 240, textAlign: 'center'}}>
					<div style={{...T.label, fontSize: 44}}>
						<Highlight at={TL_AT + 22}>{B.years}년</Highlight>
					</div>
				</Reveal>
				<Reveal at={t.word(28, '미국 주식') - 6} exitAt={WAFFLE_AT - 4} from="up" style={{left: 0, right: 0, top: 420, textAlign: 'center'}}>
					<Counter from={0} to={B.stocksThousand} at={t.word(28, '미국 주식')} format={formatThousandStocks} color={C.ink} style={{fontSize: 168}} />
					<div style={{...T.label, color: C.gray, marginTop: 22}}>분석 종목 수</div>
				</Reveal>

				{/* 자막 29: 와플 58칸 */}
				{f >= WAFFLE_AT ? <WafflePct fillAt={FILL_AT_S17} appearAt={WAFFLE_AT} /> : null}
				<Reveal at={WAFFLE_AT + 10} from="right" style={{left: 900, top: 150}}>
					<Counter from={0} to={B.belowTbillPct} at={FILL_AT_S17} dur={29} format={(v) => formatPct(v, 0, false)} color={C.blue} style={{fontSize: 240}} />
				</Reveal>
				<Reveal at={WAFFLE_AT + 18} from="right" style={{left: 912, top: 430}}>
					<div style={{display: 'flex', alignItems: 'center', gap: 18}}>
						<div style={{width: 36, height: 36, borderRadius: 6, background: C.blue}} />
						<div style={{fontFamily: SANS, fontWeight: 700, fontSize: 44, color: C.ink, whiteSpace: 'nowrap'}}>국채보다 못함</div>
					</div>
				</Reveal>
			</Layer>
			<SourceCaption text={SRC.bessembinder} at={TL_AT + 6} />
		</AbsoluteFill>
	);
};
