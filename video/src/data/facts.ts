// VERSION: v1.0.0 — 2026-10-05 — 가이드 4번 데이터 시트. 화면 숫자는 여기 값만 사용한다.
// 숫자(value)는 계산·카운터용, 표시 문자열(label)은 화면용. 새 숫자를 만들지 않는다.

/** 사용자가 채워야 하는 값 — 비어 있으면 대본 표기 그대로 "[채널명]" 으로 나간다. */
export const CHANNEL_NAME_DEFAULT = '[채널명]';

export const SRC = {
	market: '출처: 시장 데이터',
	bessembinder: '출처: Bessembinder (2018)',
	jpm: '출처: J.P. Morgan Asset Management',
} as const;

export const FACTS = {
	story: {
		principalManwon: 10000, // 사연 투자 금액 1억 원 (만 원 단위)
		gainManwon: 1000, // 사연 수익 +1,000만 원
		lossManwon: -3000, // 사연 손실 -3,000만 원
		lossPct: -30, // 원금 대비 -30%
	},
	mdb: {
		ticker: 'MDB',
		nameKo: '몽고디비',
		prevClose: {date: '2026.09.25', usd: 410.44, label: '$410.44'},
		intradayLow: {date: '2026.09.28', usd: 300, label: '약 $300', pct: -26},
		close: {date: '2026.09.28', usd: 334.68, label: '$334.68', pct: -18.5},
		ceoNewsDate: '2026.09.28',
	},
	drops: [
		// 자막 24 순서대로 카드 배치: CDNS → META → GOOGL
		{ticker: 'CDNS', pct: -10.4, label: '-10.4%', when: '2025년 5월', note: '하루'},
		{ticker: 'META', pct: -26, label: '-26%', when: '2022.02.03', note: '하루'},
		{ticker: 'GOOGL', pct: -9, label: '-9%', when: '2023.02.08', note: '장중'},
	],
	bessembinder: {
		fromYear: 1926,
		toYear: 2016,
		years: 90,
		stocksThousand: 26, // 약 2만 6천 개
		belowTbillPct: 58,
		topPct: 4,
	},
	jpm: {
		sinceYear: 1980,
		noRecoverPct: 40, // 40% 이상
		drawdownPct: -70, // 고점 대비 -70% 이상
	},
	diversify: {
		stocks: 10,
		perStockManwon: 1000, // 10종목 × 1,000만 원
		lossManwon: -300, // 약 -300만 원
		lossPctOfAccount: -3, // 계좌의 -3%
		capPct: 15, // 한 종목 비중 상한 조언
	},
} as const;

// ---- 표시 형식 -------------------------------------------------------------
const comma = (n: number): string => Math.round(Math.abs(n)).toLocaleString('en-US');

/** 만 원 단위 금액 → "1억 원" / "3,000만 원" / "0원" (부호 옵션) */
export const formatManwon = (manwon: number, signed = false): string => {
	const r = Math.round(manwon);
	const sign = signed ? (r > 0 ? '+' : r < 0 ? '-' : '') : r < 0 ? '-' : '';
	const abs = Math.abs(r);
	if (abs === 0) return '0원';
	if (abs >= 10000) {
		const eok = Math.floor(abs / 10000);
		const rest = abs % 10000;
		return rest === 0 ? `${sign}${eok}억 원` : `${sign}${eok}억 ${comma(rest)}만 원`;
	}
	return `${sign}${comma(abs)}만 원`;
};

/** 천 개 단위 → "약 2만 6천 개" */
export const formatThousandStocks = (thousand: number): string => {
	const t = Math.round(thousand);
	const man = Math.floor(t / 10);
	const cheon = t % 10;
	if (man === 0) return `약 ${cheon}천 개`;
	return cheon === 0 ? `약 ${man}만 개` : `약 ${man}만 ${cheon}천 개`;
};

/** 퍼센트 표시: 소수 자릿수 고정 */
export const formatPct = (v: number, digits = 0, signed = true): string => {
	const s = Math.abs(v).toFixed(digits);
	const sign = signed ? (v > 0 ? '+' : v < 0 ? '-' : '') : v < 0 ? '-' : '';
	return `${sign}${s}%`;
};
