// VERSION: v2.0.0 — 2026-10-06 — 가이드 v2 4번 데이터 시트. 화면 숫자는 여기 값만 사용한다.
// value = 계산·카운터용, label = 화면 표시용. 새 숫자를 만들지 않는다.

export const SRC = {
	press: '출처: 시장 보도',
	pressBloomberg: '출처: 시장 보도, Bloomberg',
	bloomberg: '출처: Bloomberg',
	bessembinder: '출처: Bessembinder (2018)',
	jpm: '출처: J.P. Morgan',
	behavioral: '출처: 행동경제학 연구',
	enronCrs: '출처: 미 의회조사국·CNN',
	cnn: '출처: CNN',
	wharton: '출처: Wharton',
	reuters: '출처: Reuters',
	berkshire: '출처: 버크셔 해서웨이',
} as const;

/** 대본 [장면] 고지 문구 3문장 (input/대본_완성본.txt 그대로) */
export const NOTICE = [
	'본 영상은 특정 종목의 매수나 매도를 권유하지 않습니다.',
	'과거의 주가 흐름과 통계가 미래의 수익을 보장하지 않으며, 주식 투자는 원금 손실이 발생할 수 있습니다.',
	'모든 투자 판단과 그에 따른 책임은 투자자 본인에게 있습니다.',
] as const;

export const FACTS = {
	story: {
		ticker: 'MDB',
		nameKo: '몽고디비',
		principalManwon: 7000, // 사연 투자 금액 7,000만 원 (전 재산, 1종목)
		gainManwon: 1000, // +1,000만 원
		lossManwon: -2000, // -2,000만 원
		afterLossManwon: 5000, // 계좌 5,000만 원
		lossPct: -29, // 원금 대비 약 -29%
		recoverPct: 40, // 5,000만 → 7,000만 회복에 +40%
		ceoNewsDate: '2026.09.28',
		maxDropPct: -26, // 장중 최대 하락
	},
	recovery: [
		// 손실의 비대칭 (S17): 하락 → 필요한 상승
		{key: 'story', dropPct: -29, needLabel: '+40%', caption: '사연'},
		{key: 'half', dropPct: -50, needLabel: '2배', caption: '반토막'},
		{key: 'seventy', dropPct: -70, needLabel: '3.3배', caption: ''},
	],
	bigtech: [
		{ticker: 'META', year: '2022', cause: '실적 실망', pct: -26, label: '-26%', note: ''},
		{ticker: 'GOOGL', year: '2023', cause: 'AI 챗봇 광고 오답', pct: -9, label: '-9%', note: '장중'},
		{ticker: 'NVDA', year: '2025', cause: '딥시크 등장', pct: -17, label: '-17%', note: ''},
	],
	nvdaRecord: '당시 하루 최대 시총 손실',
	sellAnalysts: 2, // 메타 폭락 전 '팔아라' 의견 애널리스트 2명
	bessembinder: {stocksThousand: 26, years: 90, belowTbillPct: 58, sixOfTen: '열 중 여섯', topPct: 4},
	jpm: {sinceYear: 1980, noRecoverPct: 40, drawdownPct: -70, worseThanMarket: '3번 중 2번'},
	lossAversion: 2, // 손실 고통 ≈ 이익 기쁨의 2배
	enron: {year: 2001, desc: '미국 에너지 기업', ownStockPct: 62, prices: ['80달러', '70달러'], lossLabel: '10억 달러+'},
	cap: {pct: 15, maxManwon: 1050},
	diversify: {stocks: 10, perStockManwon: 700, lossManwon: -200, lossPctOfAccount: -3},
	theme: {
		cards: ['AI 반도체', 'AI 소프트웨어', 'AI 전력주'],
		semi: {names: 'NVDA·AVGO', pct: -17},
		power: {names: '비스트라', pct: -28},
		sectors: ['반도체', '헬스케어', '소비재', '금융'],
	},
	etf: {companies: 500, buffettPct: 90},
	split: {perStockManwon: 700, times: 3},
	rebalance: {fromPct: 10, toPct: 18},
} as const;

// ---- 표시 형식 -------------------------------------------------------------
const comma = (n: number): string => Math.round(Math.abs(n)).toLocaleString('en-US');

/** 만 원 단위 금액 → "7,000만 원" / "1억 원" / "0원" (signed 면 + 부호) */
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

/** 퍼센트 표시 */
export const formatPct = (v: number, digits = 0, signed = true): string => {
	const s = Math.abs(v).toFixed(digits);
	const sign = signed ? (v > 0 ? '+' : v < 0 ? '-' : '') : v < 0 ? '-' : '';
	return `${sign}${s}%`;
};
