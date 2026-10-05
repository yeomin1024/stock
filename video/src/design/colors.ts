// VERSION: v1.0.0 — 2026-10-05 — 가이드 3-1 색상표
// 수익 = 빨강, 손실 = 파랑 (국내 증권 앱 관례). 이 파일 밖에서 HEX 를 새로 만들지 않는다.
export const C = {
	cream: '#F2EBDD', // 기본 배경 (크림 종이)
	ink: '#1E1E1E', // 선, 글자
	yellow: '#FFD400', // 강조 하이라이트
	red: '#E63B2E', // 상승·수익
	blue: '#2D6CDF', // 하락·손실
	gray: '#9A968E', // 보조 (웜 그레이)
	navy: '#14213D', // 위기 장면 배경
	// 파생 톤 (위 색에서 밝기만 조정)
	paper: '#FBF7EE', // 카드 종이 (크림보다 밝게)
	fiber: '#FFFDF7', // 찢어진 종이 단면의 흰 섬유
	white: '#FFFFFF',
} as const;

/** 수익/손실 색: 0 이상이면 빨강, 음수면 파랑 */
export const pnlColor = (v: number): string => (v < 0 ? C.blue : C.red);

/** HEX 두 색을 t(0~1)로 섞는다 */
export const mix = (a: string, b: string, t: number): string => {
	const pa = [1, 3, 5].map((i) => parseInt(a.slice(i, i + 2), 16));
	const pb = [1, 3, 5].map((i) => parseInt(b.slice(i, i + 2), 16));
	const k = Math.min(1, Math.max(0, t));
	const out = pa.map((v, i) => Math.round(v + (pb[i] - v) * k));
	return `#${out.map((v) => v.toString(16).padStart(2, '0')).join('')}`;
};

/** HEX + 불투명도 → rgba() */
export const alpha = (hex: string, a: number): string => {
	const [r, g, b] = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16));
	return `rgba(${r}, ${g}, ${b}, ${a})`;
};
