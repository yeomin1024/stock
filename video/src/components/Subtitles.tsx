// VERSION: v1.0.0 — 2026-10-05 — 자막 트랙 (가이드 3-4)
// - 화면 하단 중앙, 아래 여백 80px, 최대 폭 1500px, Noto Sans KR 700 46px
// - 크림 장면: 잉크 글자 + 크림 띠(90%), 네이비 장면: 흰 글자 + 네이비 띠(80%), 모서리 12px
// - SRT 줄바꿈 그대로, 등장/퇴장 4프레임 페이드
// - 이어지는 자막 사이에는 띠를 유지한 채 폭/높이만 부드럽게 바꿔 빈 화면이 생기지 않게 한다
import React, {useEffect, useState} from 'react';
import {AbsoluteFill, useCurrentFrame, useDelayRender} from 'remotion';
import {measureText} from '@remotion/layout-utils';
import {SUBTITLES} from '../data/subtitles';
import {navyAmount, subEnd, subStart} from '../data/timeline';
import {C, alpha, mix} from '../design/colors';
import {SANS, TYPE, fontsReady} from '../design/fonts';
import {lerp, lin, prog} from '../design/motion';

const FADE = 4;
const PAD_X = 34;
const PAD_Y = 14;
const LINE_H = Math.round(TYPE.subtitle * 1.42);
const MAX_W = 1500;
const BOTTOM = 80;

type Box = {w: number; h: number};

const SUBS = SUBTITLES.map((s) => ({id: s.id, text: s.text, start: subStart(s.id), end: subEnd(s.id)}));

export const Subtitles: React.FC = () => {
	const frame = useCurrentFrame();
	const {delayRender, continueRender, cancelRender} = useDelayRender();
	const [handle] = useState(() => delayRender('자막 폭 측정 (폰트 로드 대기)'));
	const [boxes, setBoxes] = useState<Box[] | null>(null);

	useEffect(() => {
		fontsReady()
			.then(() => {
				const measured = SUBS.map((s) => {
					const lines = s.text.split('\n');
					const w = Math.max(
						...lines.map(
							(line) =>
								measureText({text: line, fontFamily: SANS, fontSize: TYPE.subtitle, fontWeight: '700'}).width,
						),
					);
					return {w: Math.min(MAX_W, Math.ceil(w) + PAD_X * 2), h: lines.length * LINE_H + PAD_Y * 2};
				});
				setBoxes(measured);
				continueRender(handle);
			})
			.catch((err) => cancelRender(err));
	}, [continueRender, cancelRender, handle]);

	if (!boxes) return null;

	const i = SUBS.findIndex((s) => frame >= s.start && frame < s.end);
	if (i < 0) return null;
	const s = SUBS[i];
	const prev = SUBS[i - 1];
	const next = SUBS[i + 1];
	const joinedPrev = prev !== undefined && prev.end === s.start;
	const joinedNext = next !== undefined && next.start === s.end;

	// 글자: 4프레임 페이드 인/아웃
	const textIn = lin(frame, s.start, s.start + FADE);
	const textOut = 1 - lin(frame, s.end - FADE, s.end);
	const textOpacity = Math.min(textIn, textOut);

	// 띠: 이어지는 자막이면 유지, 아니면 4프레임 페이드
	const bandIn = joinedPrev ? 1 : textIn;
	const bandOut = joinedNext ? 1 : textOut;
	const bandOpacity = Math.min(bandIn, bandOut);

	// 띠 크기: 이전 자막 크기에서 6프레임 동안 이어 바꾼다
	const cur = boxes[i];
	const from = joinedPrev ? boxes[i - 1] : cur;
	const k = prog(frame, s.start, s.start + 6);
	const w = lerp(from.w, cur.w, k);
	const h = lerp(from.h, cur.h, k);

	const navy = navyAmount(frame);
	const bandColor = alpha(mix(C.cream, C.navy, navy), lerp(0.9, 0.8, navy));
	const textColor = mix(C.ink, C.white, navy);

	return (
		<AbsoluteFill style={{pointerEvents: 'none'}}>
			<div
				style={{
					position: 'absolute',
					left: (1920 - w) / 2,
					top: 1080 - BOTTOM - h,
					width: w,
					height: h,
					borderRadius: 12,
					backgroundColor: bandColor,
					opacity: bandOpacity,
				}}
			/>
			<div
				style={{
					position: 'absolute',
					left: (1920 - MAX_W) / 2,
					width: MAX_W,
					top: 1080 - BOTTOM - cur.h + PAD_Y,
					fontFamily: SANS,
					fontWeight: 700,
					fontSize: TYPE.subtitle,
					lineHeight: `${LINE_H}px`,
					color: textColor,
					textAlign: 'center',
					whiteSpace: 'pre-line',
					opacity: textOpacity,
				}}
			>
				{s.text}
			</div>
		</AbsoluteFill>
	);
};
