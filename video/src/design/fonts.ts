// VERSION: v1.0.0 — 2026-10-05 — 가이드 2/3-3 폰트: @remotion/google-fonts 로 로드
// Noto Sans KR (500/700/900): 본문, 숫자, 자막 / Noto Serif KR (700/900): 장면 헤드라인
// 청크 목록은 scripts/build-glyphs.mjs 가 실제 사용 글자로 생성한다.
import {loadFont as loadSans} from '@remotion/google-fonts/NotoSansKR';
import {loadFont as loadSerif} from '@remotion/google-fonts/NotoSerifKR';
import {SANS_SUBSETS, SERIF_SUBSETS} from './font-subsets';

type SansSubsets = NonNullable<NonNullable<Parameters<typeof loadSans>[1]>['subsets']>;
type SerifSubsets = NonNullable<NonNullable<Parameters<typeof loadSerif>[1]>['subsets']>;

const sans = loadSans('normal', {
	weights: ['500', '700', '900'],
	subsets: SANS_SUBSETS as SansSubsets,
	ignoreTooManyRequestsWarning: true,
});
const serif = loadSerif('normal', {
	weights: ['700', '900'],
	subsets: SERIF_SUBSETS as SerifSubsets,
	ignoreTooManyRequestsWarning: true,
});

export const SANS = sans.fontFamily;
export const SERIF = serif.fontFamily;

/** 자막 폭 측정처럼 폰트가 실제로 로드된 뒤에 해야 하는 작업용 */
export const fontsReady = (): Promise<unknown> => Promise.all([sans.waitUntilDone(), serif.waitUntilDone()]);

/** 가이드 3-3 타이포 스케일 (1080p 기준 px) */
export const TYPE = {
	number: 200, // 강조 숫자 160–240
	headline: 84, // 장면 헤드라인 72–96
	label: 40, // 라벨 36–44
	source: 22, // 출처 캡션
	subtitle: 46, // 자막
} as const;
