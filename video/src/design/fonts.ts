// VERSION: v2.0.0 — 2026-10-06 — 폰트: @remotion/google-fonts (Noto Sans KR 500/700/900, Noto Serif KR 700/900)
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

/** 가이드 v2 3-2 크기 (1080p 기준 px) */
export const TYPE = {
	number: 180, // 핵심 숫자 160–220
	headline: 84, // 장면 헤드라인·번호 타이틀 72–88
	label: 40, // 라벨 36–44
	caption: 28, // 출처·"개념도" 캡션 — 표의 24px 대신 "모든 글자 최소 28px" 규칙을 따름
	min: 28, // 화면의 모든 글자 최소
	subtitle: 46, // 자막
} as const;

/** 가이드 v2 3-2 배치 */
export const LAYOUT = {
	margin: 96, // 바깥 여백
	top: 120, // 그래픽 영역 위
	bottom: 800, // 그래픽 영역 아래
} as const;
