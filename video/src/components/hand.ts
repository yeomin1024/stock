// VERSION: v1.0.0 — 2026-10-05 — 손그림 느낌의 흔들리는 path 생성기 (마커 동그라미, 밑줄, 화살표, 찢어진 가장자리)
// remotion random(seed) 로 결정적(매 프레임 동일) — 렌더마다 모양이 바뀌지 않는다.
import {random} from 'remotion';

export type Pt = readonly [number, number];

const r = (seed: string): number => random(seed) * 2 - 1; // -1..1

/** Catmull-Rom 스플라인 → 부드러운 cubic bezier path */
export const smoothPath = (pts: readonly Pt[], closed = false): string => {
	if (pts.length < 2) return '';
	const p = closed ? [pts[pts.length - 1], ...pts, pts[0], pts[1]] : [pts[0], ...pts, pts[pts.length - 1]];
	let d = `M ${p[1][0].toFixed(1)} ${p[1][1].toFixed(1)}`;
	for (let i = 1; i < p.length - 2; i++) {
		const [x0, y0] = p[i - 1];
		const [x1, y1] = p[i];
		const [x2, y2] = p[i + 1];
		const [x3, y3] = p[i + 2];
		const c1x = x1 + (x2 - x0) / 6;
		const c1y = y1 + (y2 - y0) / 6;
		const c2x = x2 - (x3 - x1) / 6;
		const c2y = y2 - (y3 - y1) / 6;
		d += ` C ${c1x.toFixed(1)} ${c1y.toFixed(1)}, ${c2x.toFixed(1)} ${c2y.toFixed(1)}, ${x2.toFixed(1)} ${y2.toFixed(1)}`;
	}
	return closed ? `${d} Z` : d;
};

/** 손으로 그은 직선 (양 끝은 거의 고정, 가운데가 살짝 흔들림) */
export const handLine = (x1: number, y1: number, x2: number, y2: number, seed: string, amp = 3): string => {
	const len = Math.hypot(x2 - x1, y2 - y1);
	const n = Math.max(3, Math.round(len / 70));
	const nx = -(y2 - y1) / len;
	const ny = (x2 - x1) / len;
	const pts: Pt[] = [];
	for (let i = 0; i <= n; i++) {
		const t = i / n;
		const o = r(`${seed}-${i}`) * amp * Math.sin(Math.PI * t);
		pts.push([x1 + (x2 - x1) * t + nx * o, y1 + (y2 - y1) * t + ny * o]);
	}
	return smoothPath(pts);
};

/** 손으로 그린 마커 동그라미 (한 바퀴보다 조금 더 돌아 끝이 겹친다) */
export const handCircle = (cx: number, cy: number, rx: number, ry: number, seed: string, turns = 1.12): string => {
	const n = 36;
	const a0 = -Math.PI * 0.6 + r(`${seed}-a0`) * 0.3;
	const pts: Pt[] = [];
	for (let i = 0; i <= n; i++) {
		const t = i / n;
		const a = a0 + t * Math.PI * 2 * turns;
		const k = 1 + r(`${seed}-${i}`) * 0.025 + (t - 0.5) * 0.06;
		pts.push([cx + Math.cos(a) * rx * k, cy + Math.sin(a) * ry * k]);
	}
	return smoothPath(pts);
};

/** 화살촉 (끝점 tip, 진행 방향 각도 ang) */
export const arrowHead = (tip: Pt, ang: number, size = 26, spread = 0.5): string => {
	const [x, y] = tip;
	const a1 = ang + Math.PI - spread;
	const a2 = ang + Math.PI + spread;
	return `M ${(x + Math.cos(a1) * size).toFixed(1)} ${(y + Math.sin(a1) * size).toFixed(1)} L ${x} ${y} L ${(x + Math.cos(a2) * size).toFixed(1)} ${(y + Math.sin(a2) * size).toFixed(1)}`;
};

/** 손으로 그린 곡선 화살표: 시작→끝, bend 만큼 휘어짐. {shaft, head} */
export const handArrow = (from: Pt, to: Pt, seed: string, bend = 0.18, head = 26): {shaft: string; head: string} => {
	const [x1, y1] = from;
	const [x2, y2] = to;
	const mx = (x1 + x2) / 2;
	const my = (y1 + y2) / 2;
	const len = Math.hypot(x2 - x1, y2 - y1);
	const nx = -(y2 - y1) / len;
	const ny = (x2 - x1) / len;
	const n = 8;
	const pts: Pt[] = [];
	for (let i = 0; i <= n; i++) {
		const t = i / n;
		// 2차 베지어 위의 점 + 약간의 손떨림
		const cx = mx + nx * len * bend;
		const cy = my + ny * len * bend;
		const bx = (1 - t) * (1 - t) * x1 + 2 * (1 - t) * t * cx + t * t * x2;
		const by = (1 - t) * (1 - t) * y1 + 2 * (1 - t) * t * cy + t * t * y2;
		const o = r(`${seed}-${i}`) * 2.2 * Math.sin(Math.PI * t);
		pts.push([bx + nx * o, by + ny * o]);
	}
	const [px, py] = pts[pts.length - 2];
	const ang = Math.atan2(y2 - py, x2 - px);
	return {shaft: smoothPath(pts), head: arrowHead(to, ang, head)};
};

/**
 * 찢어진 종이 가장자리: x0 를 기준으로 위(y0)→아래(y1)로 내려가는 불규칙한 점들.
 * 큰 물결(wave) + 잔 톱니(jag). slant 만큼 아래쪽이 오른쪽으로 기운다.
 */
export const tornEdge = (
	seed: string,
	x0: number,
	y0: number,
	y1: number,
	opts: {step?: number; jag?: number; wave?: number; slant?: number} = {},
): Pt[] => {
	const {step = 13, jag = 6, wave = 16, slant = 0} = opts;
	const n = Math.ceil((y1 - y0) / step);
	const ph1 = random(`${seed}-ph1`) * Math.PI * 2;
	const ph2 = random(`${seed}-ph2`) * Math.PI * 2;
	const pts: Pt[] = [];
	for (let i = 0; i <= n; i++) {
		const t = i / n;
		const y = y0 + (y1 - y0) * t;
		const x =
			x0 +
			slant * t +
			Math.sin(t * 7.1 + ph1) * wave * 0.6 +
			Math.sin(t * 17.3 + ph2) * wave * 0.4 +
			r(`${seed}-j${i}`) * jag;
		pts.push([x, y]);
	}
	return pts;
};

/** 가로 방향 찢어진 가장자리 (x0→x1, 기준 y) */
export const tornEdgeH = (
	seed: string,
	x0: number,
	x1: number,
	y: number,
	opts: {step?: number; jag?: number; wave?: number} = {},
): Pt[] => tornEdge(seed, y, x0, x1, opts).map(([a, b]) => [b, a] as Pt);

export const ptsToPolygon = (pts: readonly Pt[]): string =>
	pts.map(([x, y]) => `${x.toFixed(1)}px ${y.toFixed(1)}px`).join(', ');

export const ptsToPath = (pts: readonly Pt[], close = true): string =>
	pts.map(([x, y], i) => `${i === 0 ? 'M' : 'L'} ${x.toFixed(1)} ${y.toFixed(1)}`).join(' ') + (close ? ' Z' : '');

/** 둥근 사각형 path (그리기 애니메이션용 — 왼쪽 위 모서리 다음에서 시작해 시계방향) */
export const roundRect = (x: number, y: number, w: number, h: number, r: number): string =>
	[
		`M ${x + r} ${y}`,
		`L ${x + w - r} ${y}`,
		`A ${r} ${r} 0 0 1 ${x + w} ${y + r}`,
		`L ${x + w} ${y + h - r}`,
		`A ${r} ${r} 0 0 1 ${x + w - r} ${y + h}`,
		`L ${x + r} ${y + h}`,
		`A ${r} ${r} 0 0 1 ${x} ${y + h - r}`,
		`L ${x} ${y + r}`,
		`A ${r} ${r} 0 0 1 ${x + r} ${y}`,
		'Z',
	].join(' ');
