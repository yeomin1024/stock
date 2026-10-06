// VERSION: v2.0.0 — 2026-10-06 — 가이드 v2 3-1 색 (의미 고정)
// 수익·상승 = 빨강, 손실·하락 = 파랑, 노랑 = 핵심 강조(한 화면에 한 곳만), 웜 그레이 = 비교 대상·보조.
// 사연자의 계좌·종목(MDB) = 잉크 테두리 + 크림 채움 카드, MDB 한 종목 비중 = 잉크 채움.
export const C = {
	cream: '#F2EBDD', // 기본 배경
	navy: '#14213D', // 위기 장면 배경 (S04~S08)
	ink: '#1E1E1E', // 글자·선
	light: '#F7F3EA', // 네이비 배경 위 글자·선
	red: '#E63B2E', // 수익·상승
	blue: '#2D6CDF', // 손실·하락
	yellow: '#FFD400', // 핵심 강조
	gray: '#9A968E', // 비교 대상·보조 정보
	// 파생 톤 (위 색에서 밝기만 조정)
	paper: '#FBF7EE', // 카드 채움 / S24 "밝은 크림"
	fiber: '#FFFDF7', // 찢어진 종이 단면
	white: '#FFFFFF',
} as const;

/** 수익/손실 색: 0 이상 빨강, 음수 파랑 */
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
