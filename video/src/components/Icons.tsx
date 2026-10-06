// VERSION: v2.0.0 — 2026-10-06 — 코드로 그린 벡터 도형 (이미지/이모지/로고/인물 그림 없음)
// v2: 사람 실루엣 제거(인물 그림 금지). 지갑·폰·달·경고·X 추가 — 모두 해당 자막의 말에서 나오는 물건만.
import React from 'react';
import {C} from '../design/colors';
import {DrawPath} from './Draw';
import {handLine, roundRect, smoothPath, Pt} from './hand';

/** 번개 도형 = "악재". 잉크 채움 — 노랑은 한 화면 한 곳의 핵심 강조에만 쓰므로(색 의미 고정) 쓰지 않는다. reveal 0→1 위에서 아래로 */
export const Lightning: React.FC<{
	readonly x: number;
	readonly y: number;
	readonly h: number;
	readonly reveal?: number;
	readonly fill?: string;
	readonly stroke?: string;
	readonly rotate?: number;
	readonly opacity?: number;
}> = ({x, y, h, reveal = 1, fill = C.ink, stroke = C.ink, rotate = 8, opacity = 1}) => {
	if (reveal <= 0 || opacity <= 0) return null;
	const w = h * 0.62;
	const pts: Pt[] = [
		[0.62, 0],
		[0.12, 0.56],
		[0.46, 0.56],
		[0.28, 1],
		[0.9, 0.38],
		[0.55, 0.38],
		[0.8, 0],
	].map(([a, b]) => [x - w / 2 + a * w, y + b * h] as Pt);
	const d = pts.map(([a, b], i) => `${i ? 'L' : 'M'} ${a.toFixed(1)} ${b.toFixed(1)}`).join(' ') + ' Z';
	return (
		<g
			opacity={opacity}
			style={{
				clipPath: `inset(0 0 ${(1 - Math.min(1, reveal)) * 100}% 0)`,
				transformBox: 'fill-box',
				transformOrigin: 'center top',
				rotate: `${rotate}deg`,
			}}
		>
			<path d={d} fill={fill} stroke={stroke} strokeWidth={Math.max(3, h * 0.025)} strokeLinejoin="round" />
		</g>
	);
};

/** 물음표 (획) — 위 y, 크기 s */
export const questionPath = (cx: number, y: number, s: number): {hook: string; dot: Pt} => ({
	hook: `M ${cx - 0.3 * s} ${y + 0.26 * s} C ${cx - 0.3 * s} ${y - 0.02 * s}, ${cx + 0.34 * s} ${y - 0.04 * s}, ${cx + 0.32 * s} ${y + 0.28 * s} C ${cx + 0.3 * s} ${y + 0.5 * s}, ${cx} ${y + 0.5 * s}, ${cx} ${y + 0.74 * s}`,
	dot: [cx, y + 0.94 * s],
});

/** 체크 표시 (손그림) — 왼쪽 위 (x, y), 크기 s */
export const checkPath = (x: number, y: number, s: number): string =>
	smoothPath([
		[x, y + 0.5 * s],
		[x + 0.2 * s, y + 0.7 * s],
		[x + 0.36 * s, y + 0.92 * s],
		[x + 0.6 * s, y + 0.5 * s],
		[x + 1.02 * s, y - 0.02 * s],
	]);

/** 생각 말풍선(구름) path — 타원 둘레를 따라 바깥으로 볼록한 호 n개 */
export const cloudPath = (cx: number, cy: number, rx: number, ry: number, n = 11, seed = 0): string => {
	const pts: Pt[] = [];
	for (let i = 0; i < n; i++) {
		const a = (i / n) * Math.PI * 2 + Math.sin(i * 2.3 + seed) * 0.06;
		pts.push([cx + Math.cos(a) * rx, cy + Math.sin(a) * ry]);
	}
	let d = `M ${pts[0][0].toFixed(1)} ${pts[0][1].toFixed(1)}`;
	for (let i = 0; i < n; i++) {
		const [ax, ay] = pts[i];
		const [bx, by] = pts[(i + 1) % n];
		const rr = Math.hypot(bx - ax, by - ay) * (0.58 + 0.06 * Math.sin(i * 1.7 + seed));
		d += ` A ${rr.toFixed(1)} ${rr.toFixed(1)} 0 0 1 ${bx.toFixed(1)} ${by.toFixed(1)}`;
	}
	return `${d} Z`;
};

/** 돋보기 (선) */
export const Magnifier: React.FC<{readonly cx: number; readonly cy: number; readonly r: number; readonly p?: number; readonly color?: string}> = ({
	cx,
	cy,
	r,
	p = 1,
	color = C.ink,
}) => (
	<>
		<DrawPath d={`M ${cx + r} ${cy} A ${r} ${r} 0 1 1 ${cx - r} ${cy} A ${r} ${r} 0 1 1 ${cx + r} ${cy}`} p={p} stroke={color} width={5} />
		<DrawPath d={handLine(cx + r * 0.72, cy + r * 0.72, cx + r * 1.45, cy + r * 1.45, 'mag', 0.5)} p={p} stroke={color} width={6} />
	</>
);

/**
 * 선으로 그린 지갑 (S08, S28 재사용 — 같은 모양). (x, y) 왼쪽 위, w×h.
 * open 0→1: 앞 덮개가 아래 경첩을 축으로 젖혀져 빈 안쪽이 보인다. draw 0→1: 선 그리기.
 */
export const Wallet: React.FC<{
	readonly x: number;
	readonly y: number;
	readonly w: number;
	readonly h: number;
	readonly open: number;
	readonly draw?: number;
	readonly line: string;
	readonly bg: string;
	readonly sw?: number;
}> = ({x, y, w, h, open, draw = 1, line, bg, sw = 6}) => {
	const r = h * 0.09;
	const hinge = y + h;
	const flap = 1 - open * 1.62;
	return (
		<g>
			<DrawPath d={roundRect(x, y, w, h, r)} p={draw} stroke={line} width={sw} />
			<g opacity={open}>
				<path d={`M ${x + w * 0.065} ${y + h * 0.36} Q ${x + w / 2} ${y + h * 0.45} ${x + w * 0.935} ${y + h * 0.36}`} stroke={line} strokeWidth={sw * 0.66} fill="none" />
				<path d={`M ${x + w * 0.065} ${y + h * 0.6} Q ${x + w / 2} ${y + h * 0.69} ${x + w * 0.935} ${y + h * 0.6}`} stroke={line} strokeWidth={sw * 0.66} fill="none" />
			</g>
			<g transform={`translate(0 ${hinge}) scale(1 ${flap}) translate(0 ${-hinge})`}>
				<path d={roundRect(x, y, w, h, r)} fill={bg} stroke="none" opacity={draw > 0.95 ? 1 : 0} />
				<DrawPath d={roundRect(x, y, w, h, r)} p={draw} stroke={line} width={sw} />
				<rect x={x + w - w * 0.21} y={y + h / 2 - h * 0.12} width={w * 0.25} height={h * 0.24} rx={h * 0.06} fill={bg} stroke={line} strokeWidth={sw * 0.83} opacity={draw > 0.95 ? 1 : 0} />
			</g>
		</g>
	);
};

/** 스마트폰 외곽 (S01, S22, S30 재사용 — 같은 모양) */
export const phoneFrame = (x: number, y: number, w: number, h: number): {body: string; notch: [number, number, number, number]} => ({
	body: roundRect(x, y, w, h, Math.min(w, h) * 0.12),
	notch: [x + w / 2 - w * 0.11, y + h * 0.035, w * 0.22, Math.max(10, h * 0.02)],
});

/** 초승달 (S22 "밤에도") */
export const moonPath = (cx: number, cy: number, r: number): string =>
	`M ${cx + r * 0.3} ${cy - r} A ${r} ${r} 0 1 0 ${cx + r * 0.3} ${cy + r} A ${r * 0.78} ${r * 0.78} 0 1 1 ${cx + r * 0.3} ${cy - r} Z`;

/** 경고 삼각형 (S29 "몰빵에 가까워짐") — 중심, 크기 */
export const warnPaths = (cx: number, cy: number, s: number): {tri: string; bar: string; dot: Pt} => ({
	tri: `M ${cx} ${cy - s * 0.5} L ${cx + s * 0.55} ${cy + s * 0.45} L ${cx - s * 0.55} ${cy + s * 0.45} Z`,
	bar: `M ${cx} ${cy - s * 0.18} L ${cx} ${cy + s * 0.14}`,
	dot: [cx, cy + s * 0.29],
});
