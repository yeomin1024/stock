// VERSION: v1.0.0 — 2026-10-05 — 코드로 그린 벡터 도형 (이미지/이모지/로고 사용 안 함)
import React from 'react';
import {C} from '../design/colors';
import {DrawPath} from './Draw';
import {handLine, smoothPath, Pt} from './hand';

/** 번개 도형 (노랑 + 잉크 외곽선). (x, y) = 위쪽 끝, h = 높이. reveal 0→1 위에서 아래로 내리친다 */
export const Lightning: React.FC<{
	readonly x: number;
	readonly y: number;
	readonly h: number;
	readonly reveal?: number;
	readonly fill?: string;
	readonly stroke?: string;
	readonly rotate?: number;
}> = ({x, y, h, reveal = 1, fill = C.yellow, stroke = C.ink, rotate = 8}) => {
	if (reveal <= 0) return null;
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

/** 방패 외곽선 path (상단 가운데 cx, 위 y, 폭 w, 높이 h) */
export const shieldPath = (cx: number, y: number, w: number, h: number): string => {
	const L = cx - w / 2;
	const R = cx + w / 2;
	return [
		`M ${L} ${y + 0.1 * h}`,
		`C ${cx - w / 4} ${y + 0.1 * h}, ${cx - 0.08 * w} ${y + 0.04 * h}, ${cx} ${y}`,
		`C ${cx + 0.08 * w} ${y + 0.04 * h}, ${cx + w / 4} ${y + 0.1 * h}, ${R} ${y + 0.1 * h}`,
		`L ${R} ${y + 0.52 * h}`,
		`C ${R} ${y + 0.8 * h}, ${cx + 0.22 * w} ${y + 0.93 * h}, ${cx} ${y + h}`,
		`C ${cx - 0.22 * w} ${y + 0.93 * h}, ${L} ${y + 0.8 * h}, ${L} ${y + 0.52 * h}`,
		'Z',
	].join(' ');
};

/** 머리+어깨 사람 실루엣 (선) */
export const personPath = (cx: number, cy: number, r: number): {head: string; body: string} => ({
	head: `M ${cx + r} ${cy} A ${r} ${r} 0 1 1 ${cx - r} ${cy} A ${r} ${r} 0 1 1 ${cx + r} ${cy}`,
	body: `M ${cx - 1.9 * r} ${cy + 3.1 * r} C ${cx - 1.9 * r} ${cy + 1.7 * r}, ${cx - 1 * r} ${cy + 1.3 * r}, ${cx} ${cy + 1.3 * r} C ${cx + 1 * r} ${cy + 1.3 * r}, ${cx + 1.9 * r} ${cy + 1.7 * r}, ${cx + 1.9 * r} ${cy + 3.1 * r}`,
});

/** 물음표 (획) — 위 y, 크기 s. 점은 별도 */
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

/** 단순화한 지폐 블록 */
export const Bill: React.FC<{readonly x: number; readonly y: number; readonly w?: number; readonly h?: number; readonly opacity?: number; readonly rotate?: number}> = ({
	x,
	y,
	w = 150,
	h = 76,
	opacity = 1,
	rotate = 0,
}) => (
	<g opacity={opacity} style={{transformBox: 'fill-box', transformOrigin: 'center', rotate: `${rotate}deg`}}>
		<rect x={x} y={y} width={w} height={h} rx={8} fill={C.paper} stroke={C.ink} strokeWidth={3.5} />
		<rect x={x + 9} y={y + 9} width={w - 18} height={h - 18} rx={4} fill="none" stroke={C.gray} strokeWidth={2} />
		<circle cx={x + w / 2} cy={y + h / 2} r={h * 0.22} fill="none" stroke={C.ink} strokeWidth={3} />
	</g>
);
