// VERSION: v2.3.0 — 2026-10-08 — 한 줄 자막: 뒷줄 내용 요소는 t.line(n, 2) · S22 (자막 52–54)
// 52: 달력 칸이 하루씩 넘어가며 같은 파랑 표시가 매일 찍힌다 "매일"
// 53: 스마트폰이 여러 번 깜빡이고 "하루 수십 번", 달 "밤에도"
// 54: "팔기"·"기다리기" 버튼이 흔들리며 둘 다 흐려진다 + "팔아야 할 때 못 팔고 / 기다려야 할 때 못 기다림"
// 연결 근거: 고통이 매일 찾아와 판단이 흐려짐
import React from 'react';
import {AbsoluteFill} from 'remotion';
import {sceneTimes} from '../data/timeline';
import {C} from '../design/colors';
import {SANS, SERIF} from '../design/fonts';
import {easeInOut, enterP, exitP, lerp, prog} from '../design/motion';
import {T} from '../design/type';
import {Svg} from '../components/Draw';
import {moonPath, phoneFrame} from '../components/Icons';
import {Reveal} from '../components/Reveal';
import {Layer, SceneBg, useSceneFrame} from '../components/Scene';
import {roundRect} from '../components/hand';

const t = sceneTimes('S22');
const S52 = t.sub(52);
const S53 = t.sub(53);
const S53L2 = t.line(53, 2); // 뒷줄 "밤에는 미국 장을 지켜보느라 잠을 설치게 됩니다."
const S54 = t.sub(54);

const CAL = {x: 300, y: 260, cw: 120, ch: 100, gap: 14};

export const S22: React.FC = () => {
	const f = useSceneFrame();
	const out53 = exitP(f, S54 - 2, 9);
	const dock = easeInOut(prog(f, S53, S53 + 18, (x) => x)); // 달력이 왼쪽 위로 작게
	const calScale = lerp(1, 0.5, dock);
	const calDx = lerp(0, 96 - CAL.x, dock);
	const calDy = lerp(0, 150 - CAL.y, dock);
	const blink = f >= S53 + 10 && f < S53 + 70 && Math.floor((f - S53 - 10) / 6) % 2 === 0;
	const phone = phoneFrame(640, 330, 240, 420);
	const shake = f >= S54 + 16 && f < S54 + 46 ? Math.sin((f - S54) * 1.6) * 4 : 0;
	const fade = 1 - 0.6 * prog(f, S54 + 46, S54 + 66);
	return (
		<AbsoluteFill>
			<SceneBg tone="cream" />
			<Layer>
				{/* 52: 달력 + 매일 */}
				<div style={{position: 'absolute', inset: 0, opacity: out53, transformOrigin: `${CAL.x}px ${CAL.y}px`, scale: String(calScale), translate: `${calDx}px ${calDy}px`}}>
					<Svg>
						{Array.from({length: 14}, (_, i) => {
							const x = CAL.x + (i % 7) * (CAL.cw + CAL.gap);
							const y = CAL.y + Math.floor(i / 7) * (CAL.ch + CAL.gap);
							const p = enterP(f, S52 + i, 10);
							const mark = prog(f, S52 + 12 + i * 4, S52 + 18 + i * 4);
							return (
								<g key={i} opacity={p}>
									<path d={roundRect(x, y, CAL.cw, CAL.ch, 10)} fill={C.paper} stroke={C.ink} strokeWidth={4} />
									<path d={`M ${x + 36} ${y + 34} L ${x + 84} ${y + 34} L ${x + 60} ${y + 72} Z`} fill={C.blue} opacity={mark} />
								</g>
							);
						})}
					</Svg>
					<Reveal at={S52 + 8} from="right" style={{left: 1290, top: 300}}>
						<div style={{fontFamily: SERIF, fontWeight: 900, fontSize: 84, color: C.ink}}>매일</div>
					</Reveal>
				</div>

				{/* 53: 폰 깜빡임 + 하루 수십 번 / 달 + 밤에도 */}
				{f >= S53 && out53 > 0.001 ? (
					<div style={{position: 'absolute', inset: 0, opacity: out53}}>
						<Svg>
							<g opacity={enterP(f, S53 + 4, 14)}>
								<path d={phone.body} fill={blink ? C.ink : C.paper} stroke={C.ink} strokeWidth={7} />
								<rect x={phone.notch[0]} y={phone.notch[1]} width={phone.notch[2]} height={phone.notch[3]} rx={6} fill={blink ? C.paper : C.ink} />
							</g>
							<path d={moonPath(1460, 440, 90)} fill={C.paper} stroke={C.ink} strokeWidth={6} opacity={enterP(f, S53L2 + 4, 14)} />
						</Svg>
						<Reveal at={S53 + 10} from="left" style={{left: 910, top: 510}}>
							<div style={{...T.label, fontSize: 44}}>하루 수십 번</div>
						</Reveal>
						<Reveal at={S53L2 + 8} from="up" style={{left: 1360, width: 220, top: 570, textAlign: 'center'}}>
							<div style={{...T.label, fontSize: 44}}>밤에도</div>
						</Reveal>
					</div>
				) : null}

				{/* 54: 팔기 / 기다리기 */}
				{[
					{label: '팔기', cap: '팔아야 할 때 못 팔고', cx: 620},
					{label: '기다리기', cap: '기다려야 할 때 못 기다림', cx: 1300},
				].map((b, i) => {
					const p = enterP(f, S54 + 4 + i * 5, 14);
					return p > 0.001 ? (
						<React.Fragment key={b.label}>
							<div
								style={{
									position: 'absolute',
									left: b.cx - 190,
									top: 300,
									width: 380,
									height: 150,
									borderRadius: 75,
									border: `5px solid ${C.ink}`,
									background: C.paper,
									boxSizing: 'border-box',
									display: 'flex',
									alignItems: 'center',
									justifyContent: 'center',
									fontFamily: SANS,
									fontWeight: 900,
									fontSize: 60,
									color: C.ink,
									opacity: p * fade,
									rotate: `${shake * (i ? -1 : 1)}deg`,
								}}
							>
								{b.label}
							</div>
							<Reveal at={S54 + 20 + i * 5} from="up" style={{left: b.cx - 300, width: 600, top: 500, textAlign: 'center'}}>
								<div style={{...T.label, fontSize: 40}}>{b.cap}</div>
							</Reveal>
						</React.Fragment>
					) : null;
				})}
			</Layer>
		</AbsoluteFill>
	);
};
