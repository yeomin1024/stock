# =============================================================================
#  sector_rotation.py
#  VERSION: v0.3.0 - 2026-09-05 - IMPROVEMENT_PLAN_SECTOR_v0.3.md §1.A~§1.E,§1.G 반영(⚠ §1.B/§1.C는
#    신호/사이징 파라미터 변경 — 아래 CHANGELOG 참조). §1.F(순환매 계층)는 조건부/생략(미구현).
#
#  목적:
#    market_regime_trader.py(이하 M)가 SPY에 대해 하는 일을 11개 SPDR 섹터 ETF(XLK/XLV/XLY/
#    XLP/XLF/XLE/XLI/XLB/XLU/XLC/XLRE) 각각에 그대로 반복한다 — 2018년부터 매 거래일, 섹터별로
#    "상승(위험선호)/중립/하락(위험회피)" 국면과 목표비중을 산출하고, M과 같은 구성의 시트로
#    기록한다. 예측 대상은 v0.1의 "SPY 대비 상대수익"이 아니라 **각 섹터의 절대 수익/낙폭**이다.
#
#  섹터 하나의 파이프라인(전부 M의 함수를 무수정 재사용):
#    후보지표 = M의 후보지표 전부(A.변동성/B.신용/C.매크로/D.크로스에셋/E.추세/F·G 자동생성 매크로·
#              크로스에셋, 실데이터 287개 — 그 섹터의 미래수익에 대해 다시 검증됨)
#            + 섹터 자체 기술지표 8종(200일선 이격도·12-1모멘텀·52주낙폭·변동성조정모멘텀·RSI·
#              MACD·50/200이격·실현변동성비 — M의 E.추세 블록과 동일 산식을 섹터 가격에 적용)
#            + SPY대비 상대강도 10종(v0.1 가족A + v0.3 REL_MA200_SLOPE) + 섹터별 매크로 8종(v0.3 §1.D 확장)
#            + SPY 계층 결과(M 번들에서 로드: 복합점수·위험점수의 '마스킹 전' 백분위, 베타 상호작용)
#            + [v0.3 §1.E] 잔차모멘텀(베타중립 RESID_MOM_12_1) + 섹터폭(SECTOR_BREADTH_200, 11섹터 공통 1회 계산)
#    → M.validate_indicators(6기준: IC/NW-t/5분위/하락AUC/4구간 부호안정성/커버리지, 시간감쇠 가중)
#    → M.build_walkforward_weights(월 재추정, 보조채택, 추세트랙캡, 기저시리즈캡, 위험(H)트랙)
#    → M.composite_score / score_percentile (수익점수·위험점수 백분위)
#    → M.generate_signals(규칙 ⓪ 급락트리거 ~ ⑩ 과열헤어컷까지 M의 국면 규칙 전부, 이력현상)
#    → M.run_backtest(t일 종가 신호 → t+1일 시가 체결, 비용, 현금레그) + 200일선 벤치마크
#    → M.event_study / drawdown_episodes / extract_trades / threshold_sensitivity / 룩어헤드 감사
#
#  M의 실행 결과는 M v1.22.0이 저장하는 결과 번들(market_regime_result.pkl.gz)에서 읽는다 —
#  M을 다시 돌리지 않는다. run(res_or_path, M): res dict(M.run 반환값) 또는 번들 경로.
#
#  market_regime_trader.py는 이 파일에서 **한 줄도 수정하지 않는다**. 유일한 런타임 의존은
#  (1) M.YAHOO_EXPECTED_START에 섹터 티커를 update()하는 것, (2) M.INDICATOR_SPECS/SPEC_BY_KEY를
#  "M의 스펙 + 섹터 스펙"으로 일시 교체하는 것(_indicator_spec_override, self_test()와 같은 패턴,
#  try/finally 원복)이다.
#
#  CHANGELOG
#  ---------------------------------------------------------------------------
#  v0.3.0 | 2026-09-05 | IMPROVEMENT_PLAN_SECTOR_v0.3.md 이행(사용자 요청 "알려준 개선사항대로 수정해서
#    코드 알려줘"). 작업순서 A→B→G→C→D→E(§4 권장, F는 순환매 계층으로 조건부/생략 — 미구현).
#    [§1.A 다음 거래일 예측 — 표시 전용, 재계산 없음] M v1.24.0 build_next_day_prediction()을 그대로
#      재사용(신규 함수 없음). 영향: run_sector()→build_sector_sheets()가 각 섹터 01_일별_티커 마지막에
#      예측 1행 추가(_append_sector_next_day_row, bt/성과 시트는 무관·불변). run()이 SPY 자체 예측
#      nd_spy를 계산해 build_prediction_matrix()에 전달 → 01Z_섹터일별예측에 '구분'(실적/예측) 열이
#      생기고 맨 끝에 예측 1행 추가. build_sector_report()의 '최근 예측'(up_line)은 구분=="실적"만
#      사용하도록 수정(예측 행이 섞여 미확정 값이 나오는 결함 방지) + '다음 거래일 예측 - *' meta 13줄
#      (SPY 1 + 섹터 11 + 안내 1) 신설, 버전 다음에 삽입.
#    [§1.B ⚠ HAZARD_SOURCE — 신호 파라미터 변경, 기본값이 v0.2.0에서 바뀜] SectorConfig.HAZARD_SOURCE:
#      "spy"(신규 기본값)/"sector"(v0.2.0 현행 — 이 값이어야 v0.2.0과 비트 동일)/"max". v0.2.0은 섹터
#      자체 haz_score의 백분위를 그대로 신호에 썼다("sector"에 해당) — v0.3.0은 기본값을 "spy"로
#      바꿔 M이 SPY에서 이미 검증·교정한 H를 재사용한다(근거: §0.3, 섹터 자체 H는 표본이 짧고 SPY의
#      크레딧/변동성 기반 H가 조기경보로 이미 검증됨). run_sector()에서 haz_pct_sector(섹터 자체 H)는
#      HAZARD_SOURCE 값과 무관하게 항상 계산해 01시트에 '위험점수백분위(H,섹터자체)'로 진단 표시.
#    [§1.C ⚠ 사이징 오버레이 변동성 정규화 — 사이징 파라미터 변경] SectorConfig.VOL_SCALE_OVERLAYS=True.
#      _sector_vol_scale(): 섹터 최초 ~252거래일 실현변동성/SPY 동기간 실현변동성 비율을 [0.7, 2.0]로
#      클립한 고정 스칼라(격자 최적화 아님, 섹터당 1회, run() 시점에 결정). _vol_scaled_cfg():
#      dataclasses.replace로 EXTENSION_HAIRCUT_STEPS 임계값(캡은 불변)·RECOVERY_CONFIRM_PCT·
#      DEEP_RECOVERY_DD·STRUCT_BOTTOM_DD에만 곱해 적용 — 매수/매도 게이트(과열 캡, 신호 임계 자체)는
#      불변. validate_and_weight_sector()의 캐시키는 조정 전 cfg_i로 계산(사이징 임계값은 검증에
#      관여하지 않으므로 캐시 재사용에 영향 없음). SIZING 로그로 vol_scale·조정된 임계값 확인 가능.
#    [§1.D 섹터별 매크로 사전방향 표 확장] SECTOR_MACRO_TABLE 11개 섹터 전부 2~4행 → 8행(공통 축: 금리
#      NOM2Y/NOM10Y, 실질금리, 기대인플레이션(BEI5Y), 달러, 유가, 구리, 신용스프레드, 커브, 클레임,
#      센티먼트 중 섹터군에 맞는 8개), 섹터군별 사전부호(XLK/XLC/XLY 금리−; XLP/XLV 금리−·H+; XLF
#      커브+·금리+; XLE 금리+·BEI+·달러−·WTI+; XLI/XLB 구리+·달러−; XLU/XLRE 금리−·H+). 신규 소스:
#      DGS2(2년물), T5YIE(5년 BEI, 이제 더 넓게 사용), DTWEXBGS(광의 달러지수, XLK), BAMLC0A0CM
#      (IG OAS). 검증 대상 후보군만 넓어짐 — 채택 여부는 워크포워드가 매월 판단(신호 로직 불변).
#    [§1.E 신규 후보지표 3종] residual_momentum_values()(베타 중립 잔차 모멘텀, RESID_MOM_12_1 —
#      rolling_beta(lag=1)로 인과적 베타 추정 후 잔차누적 21일~252일 구간), sector_breadth_200
#      (SECTOR_BREADTH_200, 11섹터 200일선 상회 비율 — run()에서 1회만 계산해 ctx로 전달, 섹터별
#      재계산 없음), REL_MA200_SLOPE(SPY대비 상대MA200의 20일 변화율). build_sector_candidates()
#      시그니처에 breadth 인자 추가, sector_lookahead_audit()도 breadth 절단·재계산 커버.
#    [§1.G 진단 시트] build_rule_contribution(): 규칙(위험회피진입/급락트리거/중립감축/과열헤어컷/
#      추세승격/회복승격)별 발동일수·발동일 익일 B&H평균수익 + [전체] 비중=0인 날의 상승일 미탑승/
#      하락일 회피 B&H수익 합(%p) — 09b_규칙별기여 신설 시트(build_sector_report), 12_섹터요약에 동일
#      값을 인용한 3열(H진입일 익일평균수익(%)/상승 미탑승(%p)/하락 회피(%p), _rule_contrib_value()로
#      단일 소스 인용·재계산 없음) 추가. 07_연도별성과에 SPY평균비중(같은 해 M의 평균 목표비중) 병기 —
#      §0.2의 "강세장 과소투자" 진단표를 매 실행 자동 재현.
#    [영향받지 않음] validate_indicators/build_walkforward_weights/composite_score/generate_signals
#      (§1.B 기본값 "spy" 경로)/run_backtest 자체 로직은 무수정 — M도 한 줄도 수정하지 않음(기존 원칙
#      유지). 격자탐색/최적화 없음(§1.C는 1회 고정 스칼라, 그리드서치 아님).
#  v0.2.0 | 2026-09-05 | 사용자 요청 "market_regime_trader.py처럼 2018년부터 일별로 11개 섹터 상승·하락 예측,
#    그 코드에서 예측에 쓰이는 모든 기술을 똑같이 사용, 시트도 똑같이, M 결과 파일을 섹터 예측에 참고".
#    [재설계] 예측 대상을 상대수익(P_i/SPY)에서 섹터 절대수익으로 변경. v0.1의 횡단면 순위·tilt 배분·
#      Σtarget=E_t 포트폴리오·수용기준 ①~⑤는 이 방향과 맞지 않아 제거(v0.1.0 파일은 프로젝트에 보관).
#    [v0.1 실측 리포트(2026-09-04)에서 발견한 결함 — 전부 이번 재설계에서 해소]
#      (a) 섹터 계층이 M의 전체 달력(1993~)에서 돌아 1993~2001년(섹터 편입 전)에 E_t>0인데 배분 대상이
#          없어 불변식 Σtarget=E_t 위반(최대오차 1.0) — v0.2는 평가창을 M과 동일한 SIGNAL_START(2018-01-02)
#          이후로 통일하고, 섹터별 신호는 그 섹터의 실제 이력에서만 산출.
#      (b) 17_섹터성과: 전략/B2/B3(1993~, 33.6년)와 B0/B1(2018~, 8.7년)의 기간이 달라 비교 무효 — v0.2의
#          모든 성과 비교는 같은 창(SIGNAL_START~).
#      (c) 가족 C(국면상호작용)가 "가중 커버리지 31%"로 전멸 — M의 score_pct/haz_pct는 SIGNAL_START 이후만
#          값이 있는 '마스킹된' 시리즈였다. v0.2는 마스킹 전 res["score"]/res["haz_score"]에
#          M.score_percentile을 다시 적용한 전체이력 백분위(인과: expanding rank)를 쓴다.
#      (d) 시대별 스프레드 시대1 NaN 등 상대수익 전용 진단은 제거.
#    [구성] 섹터별 후보 = M 후보 전부(+섹터 지표 20여 종). 검증·가중·점수·신호·백테스트·이벤트·구간·
#      민감도·감사 전부 M 함수 재사용. 시트: 00 실행요약 / 01 섹터별 일별기록(11시트, M 01과 동일 컬럼) /
#      01Z 섹터일별예측 매트릭스(날짜×11섹터 상승·중립·하락) / 02~11 M과 동일 시트에 '티커' 열 /
#      12 섹터요약(섹터별 1행) / 13 섹터분산전략(11섹터 균등분산 vs 균등 B&H vs SPY 전략 — 참고용).
#    [성능] 섹터 1개 = M 워크포워드 1회(실데이터 약 15분) → 11개 순차 2.5시간 이상. (1) 섹터별 중간결과
#      디스크 캐시(CACHE_DIR, 입력 해시 키 — 같은 데이터로 재실행 시 즉시), (2) fork 기반 프로세스 병렬
#      (MAX_WORKERS=0 → min(CPU, 4); Colab 2 vCPU 기준 약 2배), (3) SECTOR_REWEIGHT_FREQ="A"(연 재추정)
#      옵션 — 기본은 None(M과 같은 월 재추정). M의 _reestimation_boundaries는 D/W/M/A만 지원.
#    [파라미터 — 명시] SECTOR_TRAIN_MIN_YEARS=3(M은 5): XLRE(2015-10 상장)·XLC(2018-06 상장)의 신호 시작을
#      각각 2018-10·2021-07로 당기기 위함(5년이면 2020-11·2023-07). v0.1 스펙 §1.3의 편입 규칙(실제이력+3년)과
#      동일값. 나머지 9개 섹터(1998-12 상장)는 어느 값이든 2018 이전에 워밍업이 끝나 영향 없음. 3년 학습창의
#      N_eff(반감기 913일 기준 약 715) > N_EFF_MIN(500)이라 M의 유효표본 안전장치와 충돌하지 않음.
#  v0.1.0 | 2026-09-04 | 최초 구현(상대수익 로테이션). 상세는 CHANGELOG_SECTOR.md v0.1.0 항목.
#
#  ※ 본 코드는 연구/교육용 도구이며 투자 자문이 아니다. (Not financial advice)
# =============================================================================
from __future__ import annotations

import os
import sys
import time
import math
import hashlib
import traceback
import dataclasses
import contextlib
import multiprocessing as mp
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple, Any

import numpy as np
import pandas as pd

VERSION = "v0.3.0"
VERSION_DATE = "2026-09-05"

# =============================================================================
# [0] 섹터 유니버스
# =============================================================================
SECTORS: Tuple[str, ...] = (
    "XLK", "XLV", "XLY", "XLP", "XLF", "XLE", "XLI", "XLB", "XLU", "XLC", "XLRE",
)

SECTOR_NAME_KR: Dict[str, str] = {
    "XLK": "정보기술", "XLV": "헬스케어", "XLY": "경기소비재", "XLP": "필수소비재",
    "XLF": "금융", "XLE": "에너지", "XLI": "산업재", "XLB": "소재", "XLU": "유틸리티",
    "XLC": "커뮤니케이션", "XLRE": "부동산",
}

# Yahoo YAHOO_EXPECTED_START 형식과 동일하게 "상장 다음 달 1일" 단위로 표기.
SECTOR_EXPECTED_START: Dict[str, str] = {
    "XLK": "1999-01-01", "XLV": "1999-01-01", "XLY": "1999-01-01", "XLP": "1999-01-01",
    "XLF": "1999-01-01", "XLE": "1999-01-01", "XLI": "1999-01-01", "XLB": "1999-01-01",
    "XLU": "1999-01-01",
    "XLRE": "2015-11-01",   # 2015-10-07 상장
    "XLC": "2018-07-01",    # 2018-06-18 상장
}

STATE_KR = {"RISK_ON": "상승(위험선호)", "NEUTRAL": "중립", "RISK_OFF": "하락(위험회피)",
            "TREND_ONLY_IN": "추세필터-보유(지표부족)", "TREND_ONLY_OUT": "추세필터-현금(지표부족)",
            "NO_SIGNAL": "신호없음"}
STATE_SHORT = {"RISK_ON": "상승", "NEUTRAL": "중립", "RISK_OFF": "하락",
               "TREND_ONLY_IN": "추세보유", "TREND_ONLY_OUT": "추세현금", "NO_SIGNAL": "신호없음"}


def _ensure_yahoo_expected_start(M) -> None:
    """M.YAHOO_EXPECTED_START(모듈 전역 dict)에 섹터 티커를 런타임에 추가한다. .get(ticker)
    조회만 쓰는 기존 코드 경로는 새 키가 추가되어도 다른 티커에 영향받지 않는다. 멱등."""
    M.YAHOO_EXPECTED_START.update(SECTOR_EXPECTED_START)


# =============================================================================
# [1] 설정
# =============================================================================
@dataclass
class SectorConfig:
    SECTORS: Tuple[str, ...] = SECTORS
    ADJ_CLOSE_STALE_DAYS: int = 5              # Adj Close가 Close보다 이만큼 이상 늦게 끊기면 지연 판정·Close 수익률로 이어붙임
    # ---- 후보지표 구성 -------------------------------------------------------
    USE_MARKET_CANDIDATES: bool = True         # M의 후보지표 전부를 섹터 후보에 포함(그 섹터 수익에 대해 재검증)
    USE_SECTOR_TECHNICAL: bool = True          # 섹터 자체 기술지표 8종(M E.추세 블록과 동일 산식)
    USE_RELATIVE_STRENGTH: bool = True         # SPY대비 상대강도 9종(v0.1 가족A) + REL_MA200_SLOPE(v0.3.0 §1.E-3)
    USE_SECTOR_MACRO: bool = True              # 섹터별 매크로 표(v0.3.0 §1.D — 섹터당 8~10개로 확장)
    USE_SPY_LAYER_FEATURES: bool = True        # SPY 계층 결과(마스킹 전 점수/위험 백분위·베타 상호작용)
    # [v0.3.0 §1.E] 신규 후보 3종(저비용) — 전부 기존 6기준 게이트를 그대로 통과해야 채택(조용한 채택 없음).
    USE_RESID_MOMENTUM: bool = True            # §1.E-1: 잔차모멘텀(베타중립 12-1개월) — Blitz·Huij·Martens 2011
    RESID_MOM_WINDOW: int = 756                # 잔차 추정 롤링 회귀창(거래일, 약 36개월) — t-1까지의 데이터만 사용(인과)
    USE_SECTOR_BREADTH: bool = True            # §1.E-2: 섹터 폭(11섹터 중 자기 200일선 상회 비율) — 시장 내부 지표, 전 섹터 공통
    # ---- 학습/재추정 ----------------------------------------------------------
    SECTOR_TRAIN_MIN_YEARS: int = 3            # 헤더 CHANGELOG [파라미터] 참조(M은 5)
    SECTOR_REWEIGHT_FREQ: Optional[str] = None # None=M과 동일(REWEIGHT_FREQ, 기본 "M"). "A"=연 1회(⚠ 약 5배 빠름, 신호 달라짐)
    # ---- 위험(H)트랙 소스 [v0.3.0 §1.B ⚠ 위험 파라미터] ------------------------
    # 근거(IMPROVEMENT_PLAN_SECTOR_v0.3.md §0.3): 섹터 자체 H(현행 "sector")는 SPY와 같은 매크로 지표를
    # '섹터 낙폭 라벨'로 재검증한 결과라 레벨 효과에 지배되고, 실측 리포트에서 9/11 섹터가 H발동일 익일
    # 평균수익이 오히려 양(+)이었다(예측력 없음 — '위험 예측'이 아니라 '고금리·강달러 레벨 감지기'가 됨).
    # M이 SPY에서 이미 검증·교정(규칙 ①~⑨)한 H를 그대로 쓰는 "spy"를 기본값으로 한다.
    HAZARD_SOURCE: str = "spy"                 # "spy"(기본, M 번들의 마스킹 전 haz_pct 재사용) | "sector"(v0.2.0 현행) | "max"(둘 중 큰 값)
    # ---- 사이징 오버레이 변동성 정규화 [v0.3.0 §1.C ⚠ 사이징 파라미터] -----------
    # 근거(§0.5): 과열헤어컷·회복확인폭·깊은낙폭 임계값이 SPY(연변동성 ~18%) 기준 절대값이라 고변동
    # 섹터(XLK·XLE 25~30%)에서 훨씬 자주/일찍 걸린다. 섹터 자기 변동성/SPY 변동성 배율(훈련구간 초기
    # 252일, 인과적으로 고정)을 곱해 임계값을 섹터 변동성 스케일에 맞춘다. 신호 임계값 자체(게이트
    # 통과기준)는 변경하지 않는다 — 사이징 오버레이 임계값에만 적용.
    VOL_SCALE_OVERLAYS: bool = True
    VOL_SCALE_MIN: float = 0.7
    VOL_SCALE_MAX: float = 2.0
    VOL_SCALE_WINDOW: int = 252                # 배율 산정에 쓰는 창(훈련구간 첫 N거래일, 인과적으로 고정 — 격자 최적화 아님)
    # ---- 진단 단계 -------------------------------------------------------------
    RUN_SELFTEST: bool = True                  # 실데이터 전에 합성데이터 판별력 자기검사(FAIL이면 중단)
    SELFTEST_MIN_TRUE_ADOPTED: int = 2
    SELFTEST_MAX_NOISE_ADOPTED: int = 1
    RUN_THRESHOLD_SENSITIVITY: bool = True     # 06c(섹터별 약 10초)
    RUN_LOOKAHEAD_AUDIT: bool = True           # 11 룩어헤드감사(절단재계산)
    AUDIT_SAMPLE: int = 6                      # 섹터별 감사 표본 날짜 수
    # ---- 성능 -----------------------------------------------------------------
    MAX_WORKERS: int = 0                       # 0=자동(min(CPU수,4)), 1=순차. fork 불가 환경은 자동 순차
    USE_CACHE: bool = True                     # 섹터별 검증/워크포워드 결과 디스크 캐시
    CACHE_DIR: str = "./cache_sector"
    # ---- 출력 -----------------------------------------------------------------
    OUT_XLSX: str = "sector_regime_report.xlsx"
    EXPORT_DAILY_CSV: bool = True              # 01Z 매트릭스를 CSV로도 저장
    DAILY_CSV_PATH: str = "sector_regime_daily.csv"
    RANDOM_SEED: int = 20260905


CFG = SectorConfig()


def log(stage: str, msg: str, level: str = "info", M=None) -> None:
    """M(로드된 market_regime_trader 모듈)이 주어지면 그 로거로, 아니면 print로 남긴다."""
    if M is not None:
        M.log(f"SECTOR_{stage}", msg, level)
    else:
        print(f"[{level.upper()}][SECTOR_{stage}] {msg}")


def kv(**kwargs) -> str:
    parts = []
    for k, v in kwargs.items():
        if isinstance(v, float):
            v = f"{v:.4f}" if abs(v) < 1e4 else f"{v:.2f}"
        parts.append(f"{k}={v}")
    return " ".join(parts)


# =============================================================================
# [2] M.INDICATOR_SPECS / M.SPEC_BY_KEY 일시 교체 컨텍스트 매니저
# =============================================================================
@contextlib.contextmanager
def _indicator_spec_override(M, specs: List[Any]):
    """M.validate_indicators/M.build_walkforward_weights/M.build_reason_text는 지표 목록을
    인자로 받지 않고 모듈 전역 M.INDICATOR_SPECS(및 M.SPEC_BY_KEY)를 직접 참조한다 — M의
    self_test()가 쓰는 것과 같은 패턴(교체 후 try/finally 원복)을 모듈 밖에서 재현한다."""
    orig_specs = M.INDICATOR_SPECS
    orig_by_key = M.SPEC_BY_KEY
    M.INDICATOR_SPECS = list(specs)
    M.SPEC_BY_KEY = {s.key: s for s in specs}
    try:
        yield
    finally:
        M.INDICATOR_SPECS = orig_specs
        M.SPEC_BY_KEY = orig_by_key


# =============================================================================
# [3] 데이터 — 섹터 ETF 수집·무결성·Adj Close 지연 감지
# =============================================================================
def fetch_sector_prices(res: dict, M, quality_rows: List[dict],
                        sector_px_override: Optional[Dict[str, pd.DataFrame]] = None
                        ) -> Tuple[Dict[str, Optional[pd.DataFrame]], List[dict]]:
    """11개 섹터 ETF를 M.fetch_all_yahoo로 수집(퇴화수집 게이트·지연캐시 포함)하고, SPY까지
    포함해 M.validate_price_data로 무결성(시작일·교차오염) 검사한다. SPY는 res["px_dict"]["SPY"]
    (이미 검증된 프레임)를 재사용하므로 SPY 하드요건은 자연히 통과한다.
    sector_px_override: 티커→OHLC 프레임을 직접 주면 수집을 건너뛴다(합성데이터 테스트·오프라인용).
    이 경우에도 무결성 검사는 동일하게 거친다."""
    _ensure_yahoo_expected_start(M)
    cfg = res["cfg"]
    yahoo_diag: List[dict] = []
    if sector_px_override is not None:
        px = {t: sector_px_override.get(t) for t in SECTORS}
        log("DATA", kv(event="sector_px_override", tickers=sum(1 for v in px.values() if v is not None),
                       note="수집 생략 — 직접 주입된 프레임 사용(합성/오프라인)"), M=M, level="warning")
    else:
        px = M.fetch_all_yahoo(list(SECTORS), cfg, diag=yahoo_diag)
    for row in yahoo_diag:
        quality_rows.append({"시리즈": f"[Yahoo:{row.get('종류', '?')}] {row.get('시리즈', '?')}",
                             "행수": row.get("행수", 0), "시작": row.get("시작", "-"), "종료": row.get("종료", "-"),
                             "무결성판정": row.get("사유", "-")})
    combined = dict(px)
    combined["SPY"] = res["px_dict"]["SPY"]
    validated = M.validate_price_data(combined, cfg, quality_rows)
    sector_px = {t: validated.get(t) for t in SECTORS}
    n_ok = sum(1 for v in sector_px.values() if v is not None and len(v) > 0)
    log("DATA", kv(event="sector_fetch_done", tickers=len(SECTORS), ok=n_ok,
                   missing=",".join(t for t in SECTORS if sector_px.get(t) is None) or "-"), M=M)
    return sector_px, yahoo_diag


def adj_close_lag_check(df: Optional[pd.DataFrame], ticker: str, cfg_min_stale_days: int
                        ) -> Tuple[bool, int, Optional[pd.Timestamp], Optional[pd.Timestamp]]:
    """Adj Close 최종 유효일이 Close 최종 유효일보다 cfg_min_stale_days거래일 이상 앞서면
    '지연'으로 판정한다(report23의 ^VIX3M Adj Close 사례와 같은 종류 — 배당 있는 섹터 ETF는
    Adj Close가 필수라 이 검사가 더 중요하다). 반환: (지연여부, 지연일수, Adj 마지막일, Close 마지막일)."""
    if df is None or len(df) == 0 or "Close" not in df.columns:
        return False, 0, None, None
    close_last = df["Close"].dropna().index.max() if df["Close"].notna().any() else None
    if "Adj Close" not in df.columns:
        return False, 0, None, close_last
    adj_s = df["Adj Close"].dropna()
    adj_last = adj_s.index.max() if len(adj_s) else None
    if adj_last is None or close_last is None:
        return False, 0, adj_last, close_last
    gap_trading_days = int(((df.index > adj_last) & (df.index <= close_last)).sum())
    return gap_trading_days >= cfg_min_stale_days, gap_trading_days, adj_last, close_last


def build_total_return_close(df: pd.DataFrame, cal: pd.DatetimeIndex,
                             stale_days_threshold: int) -> Tuple[pd.Series, bool]:
    """Adj Close(배당 포함 총수익) 시계열을 cal에 정렬해 반환한다. Adj Close가 지연돼 있으면
    지연 구간만 Close의 일간수익률(배당 미포함 근사)을 마지막 정상 Adj Close 위에 이어붙인다.
    반환: (시계열, 대체적용여부). 상장 전 구간은 NaN으로 남긴다(ffill로 채우지 않음)."""
    close = df["Close"].reindex(cal) if "Close" in df.columns else None
    if close is None:
        return None, False
    first = df.index.min()
    close = close.where(cal >= first).ffill()
    if "Adj Close" not in df.columns:
        return close, False
    adj = df["Adj Close"].reindex(cal)
    stale, gap, adj_last, close_last = adj_close_lag_check(df, "", stale_days_threshold)
    if not stale or adj_last is None:
        return adj.where(cal >= first).ffill(), False
    adj_head = adj.loc[:adj_last].where(cal[cal <= adj_last] >= first).ffill()
    base = float(adj_head.iloc[-1])
    close_tail_ret = close.loc[adj_last:].pct_change().fillna(0.0)
    tail = base * (1.0 + close_tail_ret).cumprod()
    out = pd.concat([adj_head.iloc[:-1], tail]).reindex(cal)
    return out.where(cal >= first).ffill(), True


def sector_price_frame(df: pd.DataFrame, cal: pd.DatetimeIndex, scfg: SectorConfig
                       ) -> Tuple[pd.DataFrame, pd.DatetimeIndex, dict]:
    """섹터 OHLC 프레임을 M의 달력(cal)에 정렬하고 상장일 이후 구간(idx_i)만 잘라 반환한다.
    'Adj Close' 컬럼은 build_total_return_close()의 총수익 종가로 교체(지연 시 이어붙임).
    M.run_backtest는 Open/Close/Adj Close를, 지표는 Adj Close(총수익)와 Close(원시)를 쓴다."""
    df = df[~df.index.duplicated(keep="last")].sort_index()
    first = df.index.min()
    idx_i = cal[cal >= first]
    price = df.reindex(idx_i)
    n_gap = int(price["Close"].isna().sum())
    for c in ("Open", "High", "Low", "Close"):
        if c in price.columns:
            price[c] = price[c].astype(float).ffill()
    tr, spliced = build_total_return_close(df, cal, scfg.ADJ_CLOSE_STALE_DAYS)
    price["Adj Close"] = tr.reindex(idx_i).astype(float)
    if "Open" not in price.columns:
        price["Open"] = price["Close"]
    stale, gap, adj_last, close_last = adj_close_lag_check(df, "", scfg.ADJ_CLOSE_STALE_DAYS)
    info = {"실제데이터시작": str(first.date()), "행수": int(len(df)), "달력정렬결측(ffill)": n_gap,
            "AdjClose지연": bool(stale), "지연거래일수": int(gap),
            "AdjClose마지막일": str(adj_last.date()) if adj_last is not None else "-",
            "Close마지막일": str(close_last.date()) if close_last is not None else "-",
            "Close수익률대체적용": bool(spliced)}
    return price, idx_i, info


# =============================================================================
# [4] 섹터 후보지표 — 섹터 기술지표 / SPY대비 상대강도 / 섹터별 매크로 / SPY 계층 결과
#     key는 "{티커}__{접미사}"로 M의 키와 절대 충돌하지 않게 한다(M에 DXY_MOM 등이 이미 있음).
# =============================================================================
@dataclass
class _RawSpec:
    suffix: str
    name_kr: str
    category: str
    prior_sign: int
    rationale: str
    lead_mechanism: str
    source: str
    trend_track: bool = False
    eval_horizon: Optional[int] = None
    base_series: str = ""


def _eval_h(window: int) -> Optional[int]:
    """M._generate_universe_indicators()의 _eval_h와 동일 규칙: 60일 이상 관측창은 63일 평가지평."""
    return 63 if window >= 60 else None


def rolling_beta(sector_ret: pd.Series, spy_ret: pd.Series, window: int = 252, lag: int = 1) -> pd.Series:
    """과거 window일 롤링 베타를 lag일 지연시켜 '어제까지 확정된 베타'를 쓴다(인과)."""
    cov = sector_ret.rolling(window, min_periods=window // 2).cov(spy_ret)
    var = spy_ret.rolling(window, min_periods=window // 2).var()
    return (cov / var.replace(0, np.nan)).shift(lag)


# ---- (a) 섹터 자체 기술지표 — M.build_indicators()의 E.추세 블록·RVOL_RATIO와 동일 산식 ----------
def sector_technical_specs() -> List[_RawSpec]:
    return [
        _RawSpec("TREND_200", "섹터 200일선 대비 이격도", "E2.섹터추세", +1,
                 "대형 하락장의 대부분은 200일선 하회 구간에서 발생 — 섹터에도 같은 국면 필터",
                 "장기 추세 이탈 = 하락장 대부분 구간", "섹터 총수익종가", trend_track=True, eval_horizon=63),
        _RawSpec("MOM_12_1", "섹터 12-1개월 모멘텀", "E2.섹터추세", +1,
                 "시계열 모멘텀 프리미엄(최근 1개월 제외로 단기 반전 제거)", "시계열 모멘텀 프리미엄",
                 "섹터 총수익종가", trend_track=True, eval_horizon=63),
        _RawSpec("DD_FROM_252H", "섹터 52주 고점 대비 낙폭", "E2.섹터추세", +1,
                 "고점 대비 낙폭이 커질수록 추세 훼손", "추세 훼손 정도", "섹터 총수익종가",
                 trend_track=True, eval_horizon=63),
        _RawSpec("VOL_ADJ_MOM", "섹터 변동성조정 3개월 모멘텀", "E2.섹터추세", +1,
                 "같은 상승률이라도 변동성이 낮을 때 추세의 질이 높다", "추세의 질(risk-adjusted)",
                 "섹터 총수익종가", trend_track=True, eval_horizon=63),
        _RawSpec("RSI_14", "섹터 RSI(14)", "E2.섹터추세", +1,
                 "RSI가 50 위에 머무는 구간은 상승 지속, 40 아래 반복 이탈은 추세 훼손", "모멘텀 지속성(RSI 레짐)",
                 "섹터 총수익종가", trend_track=True, eval_horizon=63),
        _RawSpec("MACD_HIST", "섹터 MACD 히스토그램(12-26-9, 가격정규화)", "E2.섹터추세", +1,
                 "히스토그램 양수 확대=상승 가속, 음수 확대=하락 가속", "중기 추세 가속/감속",
                 "섹터 총수익종가", trend_track=True, eval_horizon=63),
        _RawSpec("MA_50_200_SPREAD", "섹터 50-200일선 이격 스프레드", "E2.섹터추세", +1,
                 "골든/데드크로스의 연속형 버전", "중장기 추세선 구조", "섹터 총수익종가",
                 trend_track=True, eval_horizon=63),
        _RawSpec("RVOL_RATIO", "섹터 실현변동성 확장비(20일/100일)", "A2.섹터변동성", -1,
                 "단기 실현변동성이 장기 대비 확장되면 변동성 군집 시작 — 하락과 동반·지속",
                 "변동성 군집 = 하락 지속 구간 진입", "섹터 총수익종가"),
    ]


def sector_technical_values(adj: pd.Series) -> pd.DataFrame:
    """M.build_indicators()의 SPY(spy_a=Adj Close) 산식을 섹터 총수익종가에 그대로 적용."""
    out = pd.DataFrame(index=adj.index)
    ma200 = adj.rolling(200, min_periods=150).mean()
    out["TREND_200"] = adj / ma200.replace(0, np.nan) - 1.0
    out["MOM_12_1"] = adj.shift(21) / adj.shift(252) - 1.0
    out["DD_FROM_252H"] = adj / adj.rolling(252, min_periods=120).max() - 1.0
    r63 = adj / adj.shift(63) - 1.0
    logr = np.log(adj.replace(0, np.nan)).diff()
    vol63 = logr.rolling(63).std() * np.sqrt(252)
    out["VOL_ADJ_MOM"] = r63 / vol63.replace(0, np.nan)
    _delta = adj.diff()
    _gain = _delta.clip(lower=0).ewm(alpha=1 / 14, adjust=False).mean()
    _loss = (-_delta.clip(upper=0)).ewm(alpha=1 / 14, adjust=False).mean()
    out["RSI_14"] = 100.0 - 100.0 / (1.0 + _gain / _loss.replace(0, np.nan))
    _ema12 = adj.ewm(span=12, adjust=False).mean()
    _ema26 = adj.ewm(span=26, adjust=False).mean()
    _macd = _ema12 - _ema26
    _sig9 = _macd.ewm(span=9, adjust=False).mean()
    out["MACD_HIST"] = (_macd - _sig9) / adj.replace(0, np.nan)
    _ma50 = adj.rolling(50, min_periods=40).mean()
    out["MA_50_200_SPREAD"] = _ma50 / ma200.replace(0, np.nan) - 1.0
    out["RVOL_RATIO"] = logr.rolling(20).std() / logr.rolling(100).std().replace(0, np.nan)
    return out.replace([np.inf, -np.inf], np.nan)


# ---- (b) SPY대비 상대강도(v0.1 가족 A) ------------------------------------------------------------
def relative_strength_specs() -> List[_RawSpec]:
    return [
        _RawSpec("REL_MOM_21", "SPY대비 21일 상대모멘텀", "H.SPY대비상대", -1,
                 "1개월 내 상대강도는 단기 과매수/과매도로 역전되는 경향(섹터 반전 효과)",
                 "최근 1개월 급등한 섹터는 단기 차익실현으로 반락", "P_i/SPY 비율"),
        _RawSpec("REL_MOM_63", "SPY대비 63일 상대모멘텀", "H.SPY대비상대", +1,
                 "3개월 상대모멘텀은 지속되는 경향(섹터 모멘텀 효과)", "자금 흐름의 관성",
                 "P_i/SPY 비율", trend_track=True, eval_horizon=63),
        _RawSpec("REL_MOM_126", "SPY대비 126일 상대모멘텀", "H.SPY대비상대", +1,
                 "6개월 상대모멘텀 지속(모멘텀 효과의 표준 관측창)", "중기 자금흐름 관성",
                 "P_i/SPY 비율", trend_track=True, eval_horizon=63),
        _RawSpec("REL_MOM_12_1", "SPY대비 12-1개월 상대모멘텀", "H.SPY대비상대", +1,
                 "최근 1개월을 제외한 12개월 상대모멘텀(전통적 모멘텀 팩터 정의)",
                 "최근월 반전 효과를 걸러낸 순수 모멘텀", "P_i/SPY 비율", trend_track=True, eval_horizon=63),
        _RawSpec("REL_MA_50_200", "상대가격 50/200일선 이격", "H.SPY대비상대", +1,
                 "상대가격(P_i/SPY)의 골든/데드크로스는 상대추세 전환의 연속형 지표",
                 "이동평균 교차는 추세추종 자금의 진입/이탈 신호", "P_i/SPY 비율", trend_track=True, eval_horizon=63),
        _RawSpec("REL_DD_252H", "상대가격 52주 고점대비 낙폭", "H.SPY대비상대", +1,
                 "상대 신고가 근접(낙폭 작음)은 상대강세 지속과 연관", "주도주 지위 유지 경향",
                 "P_i/SPY 비율", trend_track=True, eval_horizon=63),
        _RawSpec("REL_RSI_14", "상대가격 RSI(14)", "H.SPY대비상대", -1,
                 "상대가격의 단기 과매수(RSI 높음)는 역전되는 경향", "기술적 과열은 단기 조정을 부른다", "P_i/SPY 비율"),
        _RawSpec("REL_VOL_RATIO", "섹터/SPY 실현변동성 비율(z)", "H.SPY대비상대", -1,
                 "섹터 변동성이 SPY 대비 급등하면 이후 열위(저변동성 효과)", "패닉 매도 국면, 회복이 느림", "섹터·SPY 종가"),
        _RawSpec("REL_EXT_200", "섹터-SPY 200일선 이격도 차", "H.SPY대비상대", -1,
                 "섹터 자체 200일선 이격도가 SPY보다 훨씬 높으면(과열) 이후 열위",
                 "M 규칙 ⑩의 근거(이격도가 가장 강한 조기경보)를 섹터 상대판으로 재사용", "섹터·SPY 종가"),
        # [v0.3.0 §1.E-3] REL_MA_50_200(상대가격 이동평균의 '수준')과 달리 상대가격 200일선 자체의
        # '기울기'(추세 가속/감속) 정보 — REL_MOM류(가격 모멘텀)와도 다른 축.
        _RawSpec("REL_MA200_SLOPE", "상대가격(P_i/SPY) 200일선의 20일 기울기", "H.SPY대비상대", +1,
                 "상대가격 200일 이동평균 자체가 우상향으로 가속되는 것은 상대추세 전환이 구조적으로 굳어지는 신호",
                 "장기 상대추세의 방향 전환(가속도)", "P_i/SPY 비율의 200일 이동평균", trend_track=True, eval_horizon=63),
    ]


def relative_strength_values(sector_tr: pd.Series, spy_tr: pd.Series,
                             sector_raw: pd.Series, spy_raw: pd.Series, M) -> pd.DataFrame:
    rel = (sector_tr / spy_tr.replace(0, np.nan)).astype(float)
    out = pd.DataFrame(index=rel.index)
    out["REL_MOM_21"] = M._mom(rel, 21)
    out["REL_MOM_63"] = M._mom(rel, 63)
    out["REL_MOM_126"] = M._mom(rel, 126)
    out["REL_MOM_12_1"] = rel.shift(21) / rel.shift(252) - 1.0
    ma50 = rel.rolling(50, min_periods=40).mean()
    ma200 = rel.rolling(200, min_periods=150).mean()
    out["REL_MA_50_200"] = ma50 / ma200.replace(0, np.nan) - 1.0
    out["REL_DD_252H"] = rel / rel.rolling(252, min_periods=120).max() - 1.0
    _delta = rel.diff()
    _gain = _delta.clip(lower=0).ewm(alpha=1 / 14, adjust=False).mean()
    _loss = (-_delta.clip(upper=0)).ewm(alpha=1 / 14, adjust=False).mean()
    out["REL_RSI_14"] = 100.0 - 100.0 / (1.0 + _gain / _loss.replace(0, np.nan))
    r_i = np.log(sector_tr.replace(0, np.nan)).diff()
    r_spy = np.log(spy_tr.replace(0, np.nan)).diff()
    vol_ratio = r_i.rolling(20).std() / r_spy.rolling(20).std().replace(0, np.nan)
    out["REL_VOL_RATIO"] = M._z(vol_ratio, 100)
    ext_i = sector_raw / sector_raw.rolling(200, min_periods=150).mean().replace(0, np.nan) - 1.0
    ext_spy = spy_raw / spy_raw.rolling(200, min_periods=150).mean().replace(0, np.nan) - 1.0
    out["REL_EXT_200"] = ext_i - ext_spy
    # [v0.3.0 §1.E-3] 상대가격 200일선'기울기' — REL_MA_50_200(수준)과 다른 정보.
    out["REL_MA200_SLOPE"] = ma200 / ma200.shift(20) - 1.0
    return out.replace([np.inf, -np.inf], np.nan)


# ---- (c) SPY 계층 결과 + 베타 상호작용(v0.1 가족 C, 마스킹 전 백분위로 교체) ----------------------
def spy_layer_series(res: dict, M) -> Dict[str, pd.Series]:
    """M 결과 번들에서 '마스킹 전' 전체이력 백분위를 만든다. res["score_pct"]/["haz_pct"]는
    SIGNAL_START 이후만 값이 있는 리포트용 마스킹 시리즈라(v0.1 결함 (c)) 그대로 쓰면 커버리지
    게이트에서 전멸한다. res["score"]/["haz_score"]는 워크포워드 가중치가 생긴 시점부터 값이 있고
    M.score_percentile은 expanding rank(인과)이므로, 그날까지의 정보만으로 만든 백분위가 된다.
    급락트리거 백분위도 M.run()과 같은 산식(원시값×(-사전방향)의 expanding 백분위)으로 마스킹 없이 재계산."""
    out: Dict[str, pd.Series] = {}
    out["SCORE_PCT"] = M.score_percentile(res["score"])
    out["HAZ_PCT"] = M.score_percentile(res["haz_score"])
    cfg = res["cfg"]
    ft_key = getattr(cfg, "FAST_TRIGGER_INDICATOR", "VIX_TERM")
    ft_spec = next((s for s in M.INDICATOR_SPECS if s.key == ft_key), None)
    ind = res["ind"]
    if ft_spec is not None and ft_key in ind.columns and ind[ft_key].notna().any():
        out["FAST_PCT"] = M.score_percentile(ind[ft_key] * (-ft_spec.prior_sign))
    else:
        out["FAST_PCT"] = pd.Series(np.nan, index=ind.index)
    return out


def spy_layer_specs() -> List[_RawSpec]:
    return [
        _RawSpec("SPY_SCORE_PCT", "SPY 복합점수 백분위(M 계층, 마스킹 전)", "I.SPY계층", +1,
                 "시장 전체 강세 확신이 높을수록 섹터도 상승 확률이 높다(섹터는 시장 베타를 공유)",
                 "SPY 계층 확정 출력의 재사용", "M.res[score] → score_percentile"),
        _RawSpec("SPY_HAZ_PCT", "SPY 위험점수(H) 백분위(M 계층, 마스킹 전)", "I.SPY계층", -1,
                 "시장 위험(해저드)이 높을수록 섹터 하락 확률이 높다", "SPY 계층 확정 출력의 재사용",
                 "M.res[haz_score] → score_percentile"),
        _RawSpec("BETA_X_SCORE", "(베타-1)×SPY복합점수백분위", "C.국면상호작용", +1,
                 "강세 확신이 높을수록 고베타 섹터가 더 크게 상승", "강세장에서 고베타 우위",
                 "SPY_SCORE_PCT × 롤링베타"),
        _RawSpec("BETA_X_H", "(베타-1)×SPY위험점수백분위", "C.국면상호작용", -1,
                 "시장 위험이 높을수록 고베타 섹터가 더 크게 하락", "위험 프리미엄 확대 시 고베타 할인",
                 "SPY_HAZ_PCT × 롤링베타"),
        _RawSpec("BETA_X_DH", "(베타-1)×SPY위험점수 15일변화", "C.국면상호작용", -1,
                 "위험이 가속(ΔH 급등)하는 국면에서 고베타가 더 취약", "위험 '가속도'에 대한 고베타 민감도",
                 "SPY_HAZ_PCT.diff(15) × 롤링베타"),
        _RawSpec("BETA_X_FT", "(베타-1)×급락트리거백분위", "C.국면상호작용", -1,
                 "급성 변동성 급등 국면에서 고베타가 즉각 열위", "옵션 헤지수요 급증은 고베타 자산부터 매도",
                 "FAST_PCT × 롤링베타"),
    ]


def spy_layer_values(spy_series: Dict[str, pd.Series], beta: pd.Series) -> pd.DataFrame:
    idx = beta.index
    out = pd.DataFrame(index=idx)
    sp = spy_series["SCORE_PCT"].reindex(idx)
    hp = spy_series["HAZ_PCT"].reindex(idx)
    fp = spy_series["FAST_PCT"].reindex(idx)
    bx = beta - 1.0
    out["SPY_SCORE_PCT"] = sp
    out["SPY_HAZ_PCT"] = hp
    out["BETA_X_SCORE"] = bx * sp
    out["BETA_X_H"] = bx * hp
    out["BETA_X_DH"] = bx * hp.diff(15)
    out["BETA_X_FT"] = bx * fp
    return out.replace([np.inf, -np.inf], np.nan)


# ---- (d) 섹터별 매크로·크로스에셋(v0.1 가족 B 표) ---------------------------------------------------
# (suffix, name_kr, kind, source_id, transform, window, prior_sign, rationale, mechanism)
#   kind: fred_rate(레벨 변화=diff) | fred_index(변화율) | fred_level(수준 z) | yahoo_mom | ind_direct
_FamilyBRow = Tuple[str, str, str, str, str, int, int, str, str]

# [v0.3.0 §1.D] IMPROVEMENT_PLAN_SECTOR_v0.3.md §0.4/§1.D: M 후보의 사전방향은 '시장(SPY) 기준'이라
# 에너지·금융·유틸리티처럼 매크로 반응이 시장 평균과 반대인 섹터에서는 방향 불일치로 대부분 FAIL한다
# (XLE 국면검증 FAIL·엄격 PASS 13개, XLK 점수가 사실상 금리·달러 '단일 팩터'). 게이트를 낮추는 대신
# 공통 축(10Y/2Y 금리·실질금리·BEI·달러·유가·구리·신용스프레드·커브·실업청구·소비심리)을 섹터별
# 경제적 사전방향으로 8~10개씩 명시 확장한다 — 데이터에서 부호를 학습(첫 훈련창 IC로 결정)하지 않는다.
SECTOR_MACRO_TABLE: Dict[str, List[_FamilyBRow]] = {
    # 그룹: 금리-·실질금리-·달러- (장기 듀레이션 성장주 밸류에이션 채널)
    "XLK": [
        ("REALRATE10", "10년 실질금리 60일 변화", "fred_rate", "DFII10", "chg", 60, -1,
         "장기 듀레이션 성장주는 실질금리 상승에 밸류에이션이 눌린다", "할인율 채널"),
        ("NOM10Y", "10년 국채금리 60일 변화", "fred_rate", "DGS10", "chg", 60, -1,
         "명목금리 상승도 동일한 밸류에이션 압박", "할인율 채널"),
        ("NOM2Y_CHG", "2년 국채금리 60일 변화", "fred_rate", "DGS2", "chg", 60, -1,
         "단기금리 상승은 성장주 할인율 상승과 긴축 기대를 동시에 반영", "할인율/긴축기대 채널"),
        ("BEI5Y_CHG", "5년 기대인플레이션 60일 변화", "fred_rate", "T5YIE", "chg", 60, -1,
         "기대인플레 상승은 명목금리 상승 압력을 더해 장기 듀레이션 자산에 불리", "할인율 채널"),
        ("DXY_MOM", "달러인덱스 60일 모멘텀", "yahoo_mom", "DX-Y.NYB", "mom", 60, -1,
         "해외매출 비중이 높아 달러 강세가 환산이익을 깎는다", "환율 채널"),
        ("BROADUSD_CHG", "무역가중 달러지수(광의) 60일 변화율", "fred_index", "DTWEXBGS", "chg", 60, -1,
         "DXY_MOM과 다른 바스켓(광의) — 같은 환율 채널의 교차확인", "환율 채널"),
        ("IG_OAS_CHG", "투자등급 스프레드 20일 변화", "fred_rate", "BAMLC0A0CM", "chg", 20, -1,
         "성장주는 신주·전환사채 등 외부자금조달 의존도가 높아 신용스프레드 확대에 불리", "자금조달비용 채널"),
        ("RSP_MOM", "동일가중/시총가중 60일 상대강도", "ind_direct", "RSP_SPY_MOM", "", 60, -1,
         "동일가중 우위는 대형 기술주(시총상위) 열위를 시사", "지수 구성 효과"),
    ],
    # 그룹: 금리-·실질금리-·달러- (광고·구독 플랫폼도 성장주 논리 공유)
    "XLC": [
        ("REALRATE10", "10년 실질금리 60일 변화", "fred_rate", "DFII10", "chg", 60, -1,
         "광고·구독 플랫폼도 성장주 밸류에이션 논리를 공유", "할인율 채널"),
        ("NOM10Y", "10년 국채금리 60일 변화", "fred_rate", "DGS10", "chg", 60, -1,
         "명목금리 상승도 동일한 밸류에이션 압박", "할인율 채널"),
        ("NOM2Y_CHG", "2년 국채금리 60일 변화", "fred_rate", "DGS2", "chg", 60, -1,
         "단기금리 상승은 성장주 할인율 상승·긴축기대를 반영", "할인율/긴축기대 채널"),
        ("BEI5Y_CHG", "5년 기대인플레이션 60일 변화", "fred_rate", "T5YIE", "chg", 60, -1,
         "기대인플레 상승은 장기 듀레이션 자산에 불리", "할인율 채널"),
        ("DXY_MOM", "달러인덱스 60일 모멘텀", "yahoo_mom", "DX-Y.NYB", "mom", 60, -1,
         "해외 광고·구독 매출 비중이 높아 달러 강세에 환산손실", "환율 채널"),
        ("IG_OAS_CHG", "투자등급 스프레드 20일 변화", "fred_rate", "BAMLC0A0CM", "chg", 20, -1,
         "성장주 자금조달비용 채널(XLK와 동일 논리)", "자금조달비용 채널"),
        ("UMCSENT_Z", "미시간대 소비자심리지수(z)", "fred_level", "UMCSENT", "", 252, +1,
         "소비심리 개선은 광고 지출·구독 수요 확대로 연결", "광고 수요 채널"),
        ("RSP_MOM", "동일가중/시총가중 60일 상대강도", "ind_direct", "RSP_SPY_MOM", "", 60, -1,
         "메가캡 비중이 큰 섹터 특성상 XLK와 동일 논리", "지수 구성 효과"),
    ],
    # 그룹: 금리-·실질금리-·달러- (재량소비 — 소비여력·신용 채널이 추가로 강하게 작동)
    "XLY": [
        ("REALRATE10", "10년 실질금리 60일 변화", "fred_rate", "DFII10", "chg", 60, -1,
         "내구재·주택 관련 소비는 실질금리(할부·모기지 실질부담) 상승에 민감", "할인율/신용비용 채널"),
        ("NOM10Y", "10년 국채금리 60일 변화", "fred_rate", "DGS10", "chg", 60, -1,
         "장기금리 상승은 내구재 할부·모기지 금리에 전가", "신용비용 채널"),
        ("DXY_MOM", "달러인덱스 60일 모멘텀", "yahoo_mom", "DX-Y.NYB", "mom", 60, -1,
         "달러 강세는 수입 소비재 원가·해외매출 환산이익에 역풍", "환율 채널"),
        ("UMCSENT_CHG", "미시간대 소비자심리 60일 변화", "fred_rate", "UMCSENT", "chg", 60, +1,
         "소비심리 개선은 재량소비 지출 확대로 이어진다", "소비여력 채널"),
        ("ICSA_CHG", "신규 실업수당청구 60일 변화율", "fred_index", "ICSA", "chg", 60, -1,
         "고용 악화는 재량소비부터 위축시킨다", "고용-소비 채널"),
        ("HY_OAS_CHG", "하이일드 스프레드 20일 변화", "fred_rate", "BAMLH0A0HYM2", "chg", 20, -1,
         "신용스프레드 확대는 소비자·기업 신용여건 악화를 시사, 재량소비 위축", "소비자신용 채널"),
        ("OIL_MOM", "WTI 원유 60일 모멘텀", "yahoo_mom", "CL=F", "mom", 60, -1,
         "유가 상승은 가처분소득을 압박해 재량소비에 불리", "유가-소비 채널"),
        ("MORT30_CHG", "30년 모기지금리 60일 변화", "fred_rate", "MORTGAGE30US", "chg", 60, -1,
         "모기지금리 상승은 주택·내구재 관련 소비수요를 위축", "금리-내구재 채널"),
    ],
    # 그룹: 금리-·실질금리-·H+ (방어 로테이션 수혜)
    "XLP": [
        ("H_PCT", "SPY 위험점수백분위(H)", "ind_direct", "__HAZ_PCT__", "", 0, +1,
         "필수소비재는 방어 로테이션의 전형적 수혜 섹터", "방어 로테이션"),
        ("NOM10Y", "10년 국채금리 60일 변화", "fred_rate", "DGS10", "chg", 60, -1,
         "채권 대용 성격(안정배당)이라 금리 상승에 상대적으로 불리", "채권대용 채널"),
        ("REALRATE10", "10년 실질금리 60일 변화", "fred_rate", "DFII10", "chg", 60, -1,
         "채권대용 채널의 실질금리판 — 배당수익률 매력도가 실질금리에 더 민감", "채권대용 채널"),
        ("NOM2Y_CHG", "2년 국채금리 60일 변화", "fred_rate", "DGS2", "chg", 60, -1,
         "단기금리 상승은 배당주 대비 단기채의 상대매력을 높인다", "채권대용 채널"),
        ("BEI5Y_CHG", "5년 기대인플레이션 60일 변화", "fred_rate", "T5YIE", "chg", 60, -1,
         "고정배당의 실질가치를 인플레가 잠식", "채권대용 채널"),
        ("DXY_MOM", "달러인덱스 60일 모멘텀", "yahoo_mom", "DX-Y.NYB", "mom", 60, -1,
         "다국적 매출 비중이 높아 달러 강세가 환산이익을 깎는다", "환율 채널"),
        ("CURVE", "수익률곡선 10년-2년 60일 변화", "fred_rate", "T10Y2Y", "chg", 60, -1,
         "커브 가팔라짐(경기낙관)은 위험선호를 자극해 방어주 상대열위", "위험선호-방어주 채널"),
        ("HY_OAS_CHG", "하이일드 스프레드 20일 변화", "fred_rate", "BAMLH0A0HYM2", "chg", 20, +1,
         "신용스프레드 확대(위험회피)는 방어 섹터로의 자금 이동과 동반", "방어 로테이션"),
    ],
    # XLV: 방어적 성격 + 성장(바이오텍) 성격 혼재 — 방어 로테이션과 할인율 채널 동시 적용
    "XLV": [
        ("H_PCT", "SPY 위험점수백분위(H)", "ind_direct", "__HAZ_PCT__", "", 0, +1,
         "헬스케어는 경기방어적 수요(질병·처방)로 방어 로테이션 수혜", "방어 로테이션"),
        ("NOM10Y", "10년 국채금리 60일 변화", "fred_rate", "DGS10", "chg", 60, -1,
         "배당·성장이 혼재해 금리 상승에 약하게 불리", "할인율 채널(약함)"),
        ("REALRATE10", "10년 실질금리 60일 변화", "fred_rate", "DFII10", "chg", 60, -1,
         "바이오텍 등 장기 듀레이션 성장 하위섹터의 밸류에이션 압박", "할인율 채널"),
        ("NOM2Y_CHG", "2년 국채금리 60일 변화", "fred_rate", "DGS2", "chg", 60, -1,
         "단기금리 상승은 성장형 바이오텍 자금조달 여건을 악화", "할인율 채널(약함)"),
        ("BEI5Y_CHG", "5년 기대인플레이션 60일 변화", "fred_rate", "T5YIE", "chg", 60, -1,
         "장기 듀레이션 성장주 밸류에이션 압박(약함)", "할인율 채널(약함)"),
        ("DXY_MOM", "달러인덱스 60일 모멘텀", "yahoo_mom", "DX-Y.NYB", "mom", 60, -1,
         "대형 제약사 해외매출 비중이 높아 달러 강세가 환산이익을 깎는다", "환율 채널"),
        ("IG_OAS_CHG", "투자등급 스프레드 20일 변화", "fred_rate", "BAMLC0A0CM", "chg", 20, -1,
         "임상단계 바이오텍은 외부자금조달 의존도가 높아 신용스프레드 확대에 불리", "자금조달비용 채널"),
        ("ICSA_CHG", "신규 실업수당청구 60일 변화율", "fred_index", "ICSA", "chg", 60, +1,
         "고용 악화기에는 방어적 수요(처방·진료 유지)로 상대적 자금 유입", "방어 로테이션"),
    ],
    # 그룹: 커브+·금리+·HY스프레드- (순이자마진·대손비용 채널)
    "XLF": [
        ("CURVE", "수익률곡선 10년-2년 60일 변화", "fred_rate", "T10Y2Y", "chg", 60, +1,
         "커브가 가팔라지면 예대마진(순이자마진) 개선 기대", "순이자마진 채널"),
        ("NOM10Y", "10년 국채금리 60일 변화", "fred_rate", "DGS10", "chg", 60, +1,
         "금리 상승 국면에서 은행 이자수익 개선 기대", "순이자마진 채널"),
        ("NOM2Y_CHG", "2년 국채금리 60일 변화", "fred_rate", "DGS2", "chg", 60, +1,
         "단기금리 상승은 예대마진에 직접 반영되는 조달금리 채널", "순이자마진 채널"),
        ("HY_OAS_CHG", "하이일드 스프레드 20일 변화", "fred_rate", "BAMLH0A0HYM2", "chg", 20, -1,
         "신용스프레드 확대는 대손비용 증가·자금조달비용 상승을 시사", "신용비용 채널"),
        ("IG_OAS_CHG", "투자등급 스프레드 20일 변화", "fred_rate", "BAMLC0A0CM", "chg", 20, -1,
         "투자등급 스프레드 확대는 은행 자체 자금조달비용 상승", "신용비용 채널"),
        ("HOUST_CHG", "주택착공건수 120일 변화율", "fred_index", "HOUST", "chg", 120, +1,
         "주택시장 활황은 모기지·소비자대출 수요 확대", "대출수요 채널"),
        ("ICSA_CHG", "신규 실업수당청구 60일 변화율", "fred_index", "ICSA", "chg", 60, -1,
         "고용 악화는 대출 연체·손실충당 우려로 은행주에 불리", "대손비용 채널"),
        ("UMCSENT_CHG", "미시간대 소비자심리 60일 변화", "fred_rate", "UMCSENT", "chg", 60, +1,
         "소비자 신뢰 개선은 소비자대출 수요·상환여력 개선과 연결", "대출수요 채널"),
    ],
    # 그룹: 금리+·BEI+·달러-·WTI+ (원자재 직접수익·인플레헤지 채널)
    "XLE": [
        ("OIL_MOM20", "WTI 원유 20일 모멘텀", "yahoo_mom", "CL=F", "mom", 20, +1,
         "유가는 에너지 섹터 이익의 직접적 동인", "직접 수익 채널"),
        ("OIL_MOM60", "WTI 원유 60일 모멘텀", "yahoo_mom", "CL=F", "mom", 60, +1,
         "동일 논리의 중기 관측창", "직접 수익 채널"),
        ("BEI5Y_CHG", "5년 기대인플레이션 60일 변화", "fred_rate", "T5YIE", "chg", 60, +1,
         "기대인플레 상승기에 원자재·에너지가 인플레 헤지 수요를 흡수", "인플레 헤지 채널"),
        ("DXY_MOM", "달러인덱스 60일 모멘텀", "yahoo_mom", "DX-Y.NYB", "mom", 60, -1,
         "달러 강세는 달러표시 원자재 가격에 역풍", "환율-원자재 채널"),
        ("NOM10Y", "10년 국채금리 60일 변화", "fred_rate", "DGS10", "chg", 60, +1,
         "금리 상승은 흔히 강한 성장·리플레이션 국면과 동반돼 에너지 수요 기대를 반영(시장 평균과 반대 방향)",
         "리플레이션 채널"),
        ("NOM2Y_CHG", "2년 국채금리 60일 변화", "fred_rate", "DGS2", "chg", 60, +1,
         "단기금리 상승도 동일한 리플레이션/긴축 사이클 신호", "리플레이션 채널"),
        ("COPPER_MOM", "구리 60일 모멘텀", "yahoo_mom", "HG=F", "mom", 60, +1,
         "원자재 복합체 동조 — 구리 강세는 글로벌 에너지 수요 확대와 동반", "원자재 복합체 채널"),
        ("HY_OAS_CHG", "하이일드 스프레드 20일 변화", "fred_rate", "BAMLH0A0HYM2", "chg", 20, -1,
         "신용스프레드 확대(경기둔화 우려)는 글로벌 원유수요 둔화 우려와 동반", "수요둔화 채널"),
    ],
    # 그룹: 구리+·달러-·산업생산+
    "XLI": [
        ("INDPRO_CHG", "산업생산지수 60일 변화율", "fred_index", "INDPRO", "chg", 60, +1,
         "산업생산 확대는 산업재 수요와 직결", "산업수요 채널"),
        ("DGORDER_CHG", "내구재 신규주문 60일 변화율", "fred_index", "DGORDER", "chg", 60, +1,
         "내구재 신규주문 증가는 산업재 기업의 향후 매출 선행지표", "수주 채널"),
        ("COPPER_MOM", "구리 60일 모멘텀", "yahoo_mom", "HG=F", "mom", 60, +1,
         "구리는 전통적인 글로벌 경기 바로미터", "경기 바로미터 채널"),
        ("EFA_MOM", "선진국(미국제외) 주식 60일 모멘텀", "yahoo_mom", "EFA", "mom", 60, +1,
         "글로벌 산업 사이클과 동조", "글로벌 사이클 채널"),
        ("DXY_MOM", "달러인덱스 60일 모멘텀", "yahoo_mom", "DX-Y.NYB", "mom", 60, -1,
         "달러 강세는 미국 제조·수출 경쟁력과 해외매출 환산이익에 역풍", "환율-수출 채널"),
        ("REALRATE10", "10년 실질금리 60일 변화", "fred_rate", "DFII10", "chg", 60, -1,
         "실질금리 상승은 설비투자(capex)의 자본비용을 높인다", "자본비용 채널"),
        ("ICSA_CHG", "신규 실업수당청구 60일 변화율", "fred_index", "ICSA", "chg", 60, -1,
         "고용 악화는 산업재 최종수요 위축의 선행 신호", "산업수요 채널"),
        ("EEM_MOM", "신흥국 주식 60일 모멘텀", "yahoo_mom", "EEM", "mom", 60, +1,
         "신흥국(중국 등) 인프라·제조 수요가 글로벌 산업 사이클의 핵심 동인", "글로벌 사이클 채널"),
    ],
    # 그룹: 구리+·달러-·산업생산+ (원자재 성격이 XLI보다 강함)
    "XLB": [
        ("COPPER_MOM", "구리 60일 모멘텀", "yahoo_mom", "HG=F", "mom", 60, +1,
         "구리는 소재 섹터 수요의 직접 프록시", "직접 수요 채널"),
        ("DXY_MOM", "달러인덱스 60일 모멘텀", "yahoo_mom", "DX-Y.NYB", "mom", 60, -1,
         "달러 강세는 소재·원자재 가격에 역풍", "환율-원자재 채널"),
        ("EEM_MOM", "신흥국 주식 60일 모멘텀", "yahoo_mom", "EEM", "mom", 60, +1,
         "신흥국(중국 등) 수요가 소재 섹터의 핵심 동인", "신흥국 수요 채널"),
        ("BEI5Y_CHG", "5년 기대인플레이션 60일 변화", "fred_rate", "T5YIE", "chg", 60, +1,
         "인플레 기대 상승기 원자재 관련주 상대 강세", "인플레 헤지 채널"),
        ("INDPRO_CHG", "산업생산지수 60일 변화율", "fred_index", "INDPRO", "chg", 60, +1,
         "산업생산 확대는 원자재 투입수요와 직결", "산업수요 채널"),
        ("OIL_MOM60", "WTI 원유 60일 모멘텀", "yahoo_mom", "CL=F", "mom", 60, +1,
         "에너지·화학 하위섹터 비중이 있어 원자재 복합체와 동조", "원자재 복합체 채널"),
        ("REALRATE10", "10년 실질금리 60일 변화", "fred_rate", "DFII10", "chg", 60, -1,
         "실질금리 상승은 채광·화학 설비투자의 자본비용을 높인다", "자본비용 채널"),
        ("EFA_MOM", "선진국(미국제외) 주식 60일 모멘텀", "yahoo_mom", "EFA", "mom", 60, +1,
         "글로벌 산업 사이클과 동조", "글로벌 사이클 채널"),
    ],
    # 그룹: 금리-·실질금리-·H+ (채권대용·방어 로테이션)
    "XLU": [
        ("NOM10Y", "10년 국채금리 60일 변화", "fred_rate", "DGS10", "chg", 60, -1,
         "채권 대용 자산(고배당)이라 금리 상승에 가장 직접적으로 불리", "채권대용 채널"),
        ("REALRATE10", "10년 실질금리 60일 변화", "fred_rate", "DFII10", "chg", 60, -1,
         "동일 논리의 실질금리판", "채권대용 채널"),
        ("H_PCT", "SPY 위험점수백분위(H)", "ind_direct", "__HAZ_PCT__", "", 0, +1,
         "유틸리티는 대표적 방어 섹터", "방어 로테이션"),
        ("NOM2Y_CHG", "2년 국채금리 60일 변화", "fred_rate", "DGS2", "chg", 60, -1,
         "단기채 대비 배당수익률 상대매력이 단기금리 상승에 약화", "채권대용 채널"),
        ("BEI5Y_CHG", "5년 기대인플레이션 60일 변화", "fred_rate", "T5YIE", "chg", 60, -1,
         "고정배당의 실질가치를 인플레가 잠식", "채권대용 채널"),
        ("CURVE", "수익률곡선 10년-2년 60일 변화", "fred_rate", "T10Y2Y", "chg", 60, -1,
         "커브 가팔라짐(위험선호)은 저베타 방어주 상대열위와 동반", "위험선호-방어주 채널"),
        ("IG_OAS_CHG", "투자등급 스프레드 20일 변화", "fred_rate", "BAMLC0A0CM", "chg", 20, -1,
         "유틸리티는 부채비율이 높아 신용스프레드 확대(자금조달비용 상승)에 취약", "신용비용 채널"),
        ("HY_OAS_CHG", "하이일드 스프레드 20일 변화", "fred_rate", "BAMLH0A0HYM2", "chg", 20, +1,
         "신용스프레드 확대(위험회피)는 방어 섹터로의 자금 이동과 동반", "방어 로테이션"),
    ],
    # 그룹: 금리-·실질금리-·H+ (채권대용·자금조달 채널)
    "XLRE": [
        ("NOM10Y", "10년 국채금리 60일 변화", "fred_rate", "DGS10", "chg", 60, -1,
         "리츠는 배당 자산으로 금리 상승에 밸류에이션이 눌린다", "할인율 채널"),
        ("REALRATE10", "10년 실질금리 60일 변화", "fred_rate", "DFII10", "chg", 60, -1,
         "동일 논리의 실질금리판 — 배당수익률(cap rate) 매력도 채널", "할인율 채널"),
        ("MORT30_CHG", "30년 모기지금리 60일 변화", "fred_rate", "MORTGAGE30US", "chg", 60, -1,
         "모기지금리 상승은 부동산 거래·자금조달 비용에 직접 불리", "자금조달 채널"),
        ("HY_OAS_CHG", "하이일드 스프레드 20일 변화", "fred_rate", "BAMLH0A0HYM2", "chg", 20, -1,
         "리츠는 부채비율이 높아 신용스프레드 확대(자금조달비용 상승)에 취약", "신용비용 채널"),
        ("IG_OAS_CHG", "투자등급 스프레드 20일 변화", "fred_rate", "BAMLC0A0CM", "chg", 20, -1,
         "투자등급 회사채 조달 비중이 높은 대형 리츠에 직접 영향", "신용비용 채널"),
        ("HPI_CHG", "케이스실러 주택가격지수 120일 변화율", "fred_index", "CSUSHPISA", "chg", 120, +1,
         "기초자산 가치 상승은 리츠 순자산가치(NAV)에 긍정적", "자산가치 채널"),
        ("H_PCT", "SPY 위험점수백분위(H)", "ind_direct", "__HAZ_PCT__", "", 0, +1,
         "채권대용 고배당 자산 성격상 방어 로테이션에서도 일부 수혜", "방어 로테이션"),
        ("NOM2Y_CHG", "2년 국채금리 60일 변화", "fred_rate", "DGS2", "chg", 60, -1,
         "단기금리 상승은 리츠의 단기 차환(리파이낸싱) 비용을 직접 높인다", "자금조달 채널"),
    ],
}


def sector_macro_values(ticker: str, res: dict, M, idx: pd.DatetimeIndex,
                        spy_series: Dict[str, pd.Series]) -> pd.DataFrame:
    """source series는 전부 res["fred"](발표지연 적용 완료)/res["px_dict"]/res["ind"]/
    spy_series(마스킹 전 H 백분위)에서 가져온다 — 새로 수집하는 원천 데이터는 없다."""
    out = pd.DataFrame(index=idx)
    for suffix, name_kr, kind, sid, transform, window, prior_sign, rationale, mech in \
            SECTOR_MACRO_TABLE.get(ticker, []):
        col = None
        if kind in ("fred_rate", "fred_index", "fred_level"):
            src = res["fred"].get(sid)
            if src is not None:
                s = src.reindex(idx).ffill()
                if kind == "fred_rate":
                    col = s - s.shift(window)
                elif kind == "fred_index":
                    col = s / s.shift(window) - 1.0
                else:
                    col = M._z(s, window or 252)
        elif kind == "yahoo_mom":
            d = res["px_dict"].get(sid)
            if d is not None:
                price_col = "Adj Close" if "Adj Close" in d.columns else "Close"
                col = M._mom(d[price_col].reindex(idx).ffill(), window)
        elif kind == "ind_direct":
            if sid == "__HAZ_PCT__":
                col = spy_series["HAZ_PCT"].reindex(idx)
            else:
                ind = res.get("ind")
                col = ind[sid].reindex(idx) if (ind is not None and sid in ind.columns) else None
        out[suffix] = col if col is not None else np.nan
    return out.replace([np.inf, -np.inf], np.nan)


# ---- (e) [v0.3.0 §1.E-1] 잔차 모멘텀(residual momentum, 베타중립) --------------------------------
def residual_momentum_values(adj_tr: pd.Series, spy_tr: pd.Series, window: int = 756) -> pd.Series:
    """섹터 일간(로그)수익을 SPY에 window일(기본 756≈36개월) 롤링 회귀한 잔차의 12-1개월 누적.
    베타는 rolling_beta()로 t-1까지의 데이터만으로 추정해 lag=1 적용(인과) — 그 날의 잔차 계산에
    당일 확정 베타가 아니라 '어제까지 확정된' 베타를 쓴다. 잔차 누적합(로그가산)의 21일 전 값 -
    252일 전 값 = 최근 1개월을 제외한 11개월 구간의 순수(베타중립) 초과수익.
    시장 베타 성분을 제거한 순수 섹터 모멘텀(Blitz·Huij·Martens 2011) — 상대모멘텀(REL_MOM_12_1,
    시장 대비 '가격비율'의 모멘텀이라 베타가 섞여 있음)보다 베타 노출에 강건하다는 문헌 근거."""
    r_i = np.log(adj_tr.replace(0, np.nan)).diff()
    r_spy = np.log(spy_tr.replace(0, np.nan)).diff()
    beta = rolling_beta(r_i, r_spy, window=window, lag=1)
    resid = r_i - beta * r_spy
    resid_cum = resid.cumsum()
    return (resid_cum.shift(21) - resid_cum.shift(252)).replace([np.inf, -np.inf], np.nan)


def resid_momentum_specs() -> List[_RawSpec]:
    return [
        _RawSpec("RESID_MOM_12_1", "잔차모멘텀(베타중립) 12-1개월", "E2.섹터추세", +1,
                 "시장 베타 성분을 제거한 순수 섹터 모멘텀 — 상대모멘텀보다 베타 노출에 강건하다는 문헌 근거"
                 "(Blitz·Huij·Martens 2011)", "베타중립 모멘텀 프리미엄",
                 "섹터·SPY 총수익종가 36개월 롤링회귀 잔차(베타 t-1 lag)", trend_track=True, eval_horizon=63),
    ]


# ---- (f) [v0.3.0 §1.E-2] 섹터 폭(breadth) — 11개 섹터 공통 시장 내부 지표 ------------------------
def sector_breadth_specs() -> List[_RawSpec]:
    return [
        _RawSpec("SECTOR_BREADTH_200", "섹터 폭(자기 200일선 상회 섹터 비율)", "J.시장폭", +1,
                 "200일선 위에 있는 섹터가 많을수록 상승이 소수 주도주에 국한되지 않고 시장 내부로 넓게 "
                 "확산돼 있다는 신호(breadth) — 전 섹터 공통(그날 상장돼 있는 섹터만으로 계산)",
                 "시장 내부 참여도(breadth) 확산", "11개 섹터 각자의 자기 200일선 상회 여부(res['run']에서 1회 계산)",
                 trend_track=True, eval_horizon=63),
    ]


def _mk_spec(M, ticker: str, raw: _RawSpec):
    return M.IndicatorSpec(
        key=f"{ticker}__{raw.suffix}", name_kr=f"[{ticker}] {raw.name_kr}", category=raw.category,
        prior_sign=raw.prior_sign, rationale=raw.rationale, source=raw.source,
        lead_mechanism=raw.lead_mechanism, trend_track=raw.trend_track,
        eval_horizon=raw.eval_horizon, base_series=raw.base_series)


def build_sector_candidates(ticker: str, res: dict, M, scfg: SectorConfig,
                            adj_tr: pd.Series, close_raw: pd.Series,
                            spy_tr: pd.Series, spy_raw: pd.Series,
                            spy_series: Dict[str, pd.Series], idx: pd.DatetimeIndex,
                            breadth: Optional[pd.Series] = None,
                            ) -> Tuple[pd.DataFrame, List[Any]]:
    """한 섹터의 후보지표 전체(값 DataFrame, M.IndicatorSpec 목록)를 만든다.
    = [M 후보 전부(res["ind"] 재사용, 재계산 없음)] + [섹터 기술 8] + [상대강도 10(v0.3.0: REL_MA200_SLOPE 추가)]
      + [섹터 매크로 8~10(v0.3.0 §1.D 확장)] + [SPY 계층 6] + [v0.3.0 §1.E 신규 2~3: 잔차모멘텀·섹터폭].
    모든 입력은 idx(섹터 상장일 이후 달력)로 정렬된다.
    breadth: [v0.3.0 §1.E-2] run()에서 11섹터 전체로 1회 계산한 SECTOR_BREADTH_200(공통 시장 내부 지표) —
    None이면(단일 섹터 테스트 등) 해당 후보를 건너뛴다."""
    frames: List[pd.DataFrame] = []
    specs: List[Any] = []
    if scfg.USE_MARKET_CANDIDATES:
        m_ind = res["ind"].reindex(idx)
        frames.append(m_ind)
        specs.extend(list(M.INDICATOR_SPECS))
    sec = pd.DataFrame(index=idx)
    if scfg.USE_SECTOR_TECHNICAL:
        vals = sector_technical_values(adj_tr.reindex(idx))
        for raw in sector_technical_specs():
            sec[f"{ticker}__{raw.suffix}"] = vals[raw.suffix]
            specs.append(_mk_spec(M, ticker, raw))
    if scfg.USE_RELATIVE_STRENGTH:
        vals = relative_strength_values(adj_tr.reindex(idx), spy_tr.reindex(idx),
                                        close_raw.reindex(idx), spy_raw.reindex(idx), M)
        for raw in relative_strength_specs():
            sec[f"{ticker}__{raw.suffix}"] = vals[raw.suffix]
            specs.append(_mk_spec(M, ticker, raw))
    if scfg.USE_SECTOR_MACRO:
        vals = sector_macro_values(ticker, res, M, idx, spy_series)
        for suffix, name_kr, kind, sid, transform, window, prior_sign, rationale, mech in \
                SECTOR_MACRO_TABLE.get(ticker, []):
            sec[f"{ticker}__{suffix}"] = vals[suffix] if suffix in vals.columns else np.nan
            specs.append(M.IndicatorSpec(
                key=f"{ticker}__{suffix}", name_kr=f"[{ticker}] {name_kr}", category="B2.섹터매크로",
                prior_sign=prior_sign, rationale=rationale, source=f"{kind}:{sid}",
                lead_mechanism=mech, trend_track=False, eval_horizon=_eval_h(window),
                base_series=(sid if kind != "ind_direct" else "")))
    if scfg.USE_SPY_LAYER_FEATURES:
        r_i = np.log(adj_tr.reindex(idx).replace(0, np.nan)).diff()
        r_spy = np.log(spy_tr.reindex(idx).replace(0, np.nan)).diff()
        beta = rolling_beta(r_i, r_spy, window=252, lag=1)
        vals = spy_layer_values(spy_series, beta)
        for raw in spy_layer_specs():
            sec[f"{ticker}__{raw.suffix}"] = vals[raw.suffix]
            specs.append(_mk_spec(M, ticker, raw))
    # [v0.3.0 §1.E-1] 잔차모멘텀(베타중립) — 기존 6기준 게이트를 그대로 통과해야 채택.
    if getattr(scfg, "USE_RESID_MOMENTUM", True):
        rv = residual_momentum_values(adj_tr.reindex(idx), spy_tr.reindex(idx),
                                      window=getattr(scfg, "RESID_MOM_WINDOW", 756))
        for raw in resid_momentum_specs():
            sec[f"{ticker}__{raw.suffix}"] = rv
            specs.append(_mk_spec(M, ticker, raw))
    # [v0.3.0 §1.E-2] 섹터 폭 — run()에서 11섹터 공통으로 1회 계산된 값을 그대로 재사용(재계산 없음).
    if getattr(scfg, "USE_SECTOR_BREADTH", True) and breadth is not None:
        bv = breadth.reindex(idx)
        for raw in sector_breadth_specs():
            sec[f"{ticker}__{raw.suffix}"] = bv
            specs.append(_mk_spec(M, ticker, raw))
    frames.append(sec)
    ind = pd.concat(frames, axis=1).replace([np.inf, -np.inf], np.nan)
    keys = [s.key for s in specs]
    assert len(keys) == len(set(keys)), "후보지표 key 중복"
    ind = ind[keys]
    return ind, specs


# =============================================================================
# [5] 판별력 자기검사 — 실데이터 해석 전에 반드시 먼저 통과해야 함
#     합성 '섹터 가격'에 선행신호를 심은 지표 3개 + 순수 잡음 5개를 M.validate_indicators에
#     넣어 전자만 PASS하는지 확인한다(검증 레이어가 고장 나면 어떤 결과도 신뢰할 수 없으므로).
# =============================================================================
def build_selftest_case(M, scfg: SectorConfig, n: int = 1500
                        ) -> Tuple[pd.Series, pd.DataFrame, List[Any]]:
    """M._build_hazard_selftest_case()와 같은 설계 원칙 — 대부분의 날은 순수 잡음이고, 간격을 둔
    소수의 '사건' 구간에서만 지표가 사건 직전에 스파이크하고 가격이 그 방향으로 크게 움직인다."""
    rng = np.random.default_rng(scfg.RANDOM_SEED)
    idx = pd.bdate_range("2010-01-04", periods=n)
    true_meta = [("TRUE_A", +1), ("TRUE_B", -1), ("TRUE_C", +1)]
    noise_keys = ["NOISE_A", "NOISE_B", "NOISE_C", "NOISE_D", "NOISE_E"]
    ret = rng.normal(0.0002, 0.006, n)
    ind_data: Dict[str, np.ndarray] = {k: rng.normal(0.0, 1.0, n) for k, _ in true_meta}
    for k in noise_keys:
        ind_data[k] = rng.normal(0.0, 1.0, n)
    move_days, spacing = 5, 45
    for p in range(spacing, n - spacing - 10, spacing):
        direction = 1 if rng.uniform() < 0.5 else -1
        magnitude = direction * rng.uniform(0.04, 0.08)
        per_day = math.log(1.0 + magnitude) / move_days
        for j in range(move_days):
            ret[p + j] += per_day
        for key, sign in true_meta:
            for off in range(-3, 3):
                t = p + off
                if 0 <= t < n:
                    ind_data[key][t] += sign * direction * 5.0 * (1.0 - abs(off) / 4.0)
    px = pd.Series(100.0 * np.exp(np.cumsum(ret)), index=idx)
    ind = pd.DataFrame({k: pd.Series(v, index=idx) for k, v in ind_data.items()})
    specs: List[Any] = []
    for key, sign in true_meta:
        specs.append(M.IndicatorSpec(key=key, name_kr=f"[자체테스트] {key}", category="TEST", prior_sign=sign,
                                     rationale="자기검사 전용 합성 신호", source="합성데이터", lead_mechanism="주입된 선행신호"))
    for key in noise_keys:
        specs.append(M.IndicatorSpec(key=key, name_kr=f"[자체테스트] {key}", category="TEST", prior_sign=1,
                                     rationale="자기검사 전용 순수 잡음", source="합성데이터", lead_mechanism="없음(잡음)"))
    return px, ind, specs


def run_selftest(M, scfg: SectorConfig = CFG) -> Dict[str, Any]:
    """판별력 자기검사. run()이 실데이터 전에 자동 호출하고 FAIL이면 중단한다. 합성 사건의 크기
    (±4~8%/5일)에 맞춰 낙폭 라벨 -3%를 쓴다(픽스처 파라미터 — 실데이터 검증은 M과 같은 -5%)."""
    px, ind, specs = build_selftest_case(M, scfg)
    cfg_t = dataclasses.replace(M.CFG, DD_LABEL_THRESHOLD=-0.03, USE_HAZARD_TRACK=False,
                                DATA_START=str(px.index[0].date()), SIGNAL_START=str(px.index[0].date()),
                                TRAIN_MIN_YEARS=3, MIN_INDICATORS=3, REWEIGHT_FREQ="M")
    with _indicator_spec_override(M, specs):
        vt = M.validate_indicators(ind, px, cfg_t, verbose=False, compute_quintiles=False)
    v = vt.set_index("지표코드")
    true_keys = ["TRUE_A", "TRUE_B", "TRUE_C"]
    noise_keys = ["NOISE_A", "NOISE_B", "NOISE_C", "NOISE_D", "NOISE_E"]
    n_true = int((v.loc[true_keys, "판정"] == "PASS").sum())
    n_noise = int((v.loc[noise_keys, "판정"] == "PASS").sum())
    passed = n_true >= scfg.SELFTEST_MIN_TRUE_ADOPTED and n_noise <= scfg.SELFTEST_MAX_NOISE_ADOPTED
    log("SELFTEST", kv(true_pass=f"{n_true}/3", noise_pass=f"{n_noise}/5", verdict="PASS" if passed else "FAIL"),
        M=M, level="info" if passed else "error")
    return {"passed": passed, "n_true_pass": n_true, "n_noise_pass": n_noise, "vt": vt}


# =============================================================================
# [6] 섹터 1개 파이프라인 = M.run()의 3)전체표본검증 ~ 7)감사 단계를 섹터 가격에 적용
# =============================================================================
def sector_cfg_for(M_cfg, scfg: SectorConfig, ticker: str, idx_i: pd.DatetimeIndex):
    """M cfg를 베이스로 섹터별 사본. 신호·검증 파라미터(임계값·지평·게이트·규칙 스위치)는 전부
    M 값을 그대로 두고, 티커/이력 시작/학습최소연수/재추정주기/감사표본/시드/출력만 바꾼다."""
    fields = {f.name for f in dataclasses.fields(M_cfg)}
    upd = dict(TRADE_TICKER=ticker, DATA_START=str(idx_i[0].date()),
               TRAIN_MIN_YEARS=scfg.SECTOR_TRAIN_MIN_YEARS,
               REWEIGHT_FREQ=scfg.SECTOR_REWEIGHT_FREQ or M_cfg.REWEIGHT_FREQ,
               RUN_LOOKAHEAD_AUDIT=scfg.RUN_LOOKAHEAD_AUDIT, AUDIT_SAMPLE=scfg.AUDIT_SAMPLE,
               RANDOM_SEED=scfg.RANDOM_SEED, OUT_XLSX=f"sector_{ticker}.xlsx",
               EXPORT_RESULT_BUNDLE=False, EXPORT_DAILY_CSV=False)
    return dataclasses.replace(M_cfg, **{k: v for k, v in upd.items() if k in fields})


def _sector_vol_scale(adj_tr: pd.Series, spy_tr: pd.Series, scfg: SectorConfig) -> float:
    """[v0.3.0 §1.C ⚠] 섹터 자기 변동성 / SPY 변동성 배율 k — 그 섹터 이력의 '첫' VOL_SCALE_WINDOW
    거래일(기본 252일)만으로 1회 고정 계산한다. 어떤 신호 계산 시점 t와도 무관한 구조적 상수라
    격자탐색이 아니며(§3 프로토콜의 '격자 금지' 원칙), 이후 어떤 날짜에도 미래 구간의 변동성 정보가
    섞여 들어가지 않는다(강한 의미의 인과성 — t-1이 아니라 애초에 t에 의존하지 않음).
    표본이 부족하거나(신규상장 등) 계산 불가하면 안전하게 1.0(배율 없음)을 반환한다."""
    w = scfg.VOL_SCALE_WINDOW
    r_i = np.log(adj_tr.replace(0, np.nan)).diff().dropna()
    if len(r_i) == 0:
        return 1.0
    r_i = r_i.iloc[:w]
    r_spy_full = np.log(spy_tr.replace(0, np.nan)).diff()
    r_spy = r_spy_full.reindex(r_i.index)
    min_n = max(20, w // 2)
    if len(r_i) < min_n or int(r_spy.notna().sum()) < min_n:
        return 1.0
    sigma_i = float(r_i.std())
    sigma_spy = float(r_spy.std())
    if not np.isfinite(sigma_i) or not np.isfinite(sigma_spy) or sigma_spy <= 0:
        return 1.0
    return float(np.clip(sigma_i / sigma_spy, scfg.VOL_SCALE_MIN, scfg.VOL_SCALE_MAX))


def _vol_scaled_cfg(cfg_i, vol_scale: float):
    """[v0.3.0 §1.C] cfg_i 사본 — EXTENSION_HAIRCUT_STEPS(임계값만, 상한은 불변)·RECOVERY_CONFIRM_PCT·
    DEEP_RECOVERY_DD·STRUCT_BOTTOM_DD에 vol_scale을 곱한다. 다른 필드는 전부 그대로(신호·검증
    게이트 자체는 변경하지 않음 — 사이징 오버레이 임계값에만 적용)."""
    if vol_scale == 1.0:
        return cfg_i
    steps = tuple((round(thr * vol_scale, 4), cap) for thr, cap in cfg_i.EXTENSION_HAIRCUT_STEPS)
    return dataclasses.replace(
        cfg_i, EXTENSION_HAIRCUT_STEPS=steps,
        RECOVERY_CONFIRM_PCT=round(cfg_i.RECOVERY_CONFIRM_PCT * vol_scale, 4),
        DEEP_RECOVERY_DD=round(cfg_i.DEEP_RECOVERY_DD * vol_scale, 4),
        STRUCT_BOTTOM_DD=round(cfg_i.STRUCT_BOTTOM_DD * vol_scale, 4))


# 검증/워크포워드 결과에 영향을 주지 않는 것이 확인된(M 소스 정적 검사: validate_indicators/build_walkforward_weights/
# _select_and_weight_*가 참조하는 cfg 필드 목록에 없음) 리포트·I/O·감사·수집 전용 필드만 캐시 키에서 제외한다.
# 나머지 필드는 전부 키에 포함(보수적 — 불필요한 재계산은 있어도 오래된 캐시로 인한 오답은 없다).
_CACHE_KEY_IGNORE_FIELDS = frozenset({
    "AUDIT_SAMPLE", "RUN_LOOKAHEAD_AUDIT", "RUN_HALF_LIFE_SENSITIVITY", "HL_SENS_REWEIGHT_FREQ", "ENSEMBLE_HALF_LIVES",
    "OUT_XLSX", "LOG_LEVEL", "EXPORT_RESULT_BUNDLE", "RESULT_BUNDLE_PATH", "EXPORT_DAILY_CSV", "DAILY_CSV_PATH",
    "RANDOM_SEED", "CACHE_DIR", "FRED_API_KEY", "FETCH_TIMEOUT_CONNECT", "FETCH_TIMEOUT_READ", "FETCH_RETRIES",
    "FETCH_MAX_WORKERS", "DRAWDOWN_EPISODE_THRESHOLD",
})


def _cache_key(ticker: str, cfg_i, ind: pd.DataFrame, px_adj: pd.Series, M) -> str:
    """검증/워크포워드 캐시 키 — 코드 버전·M 버전·섹터 cfg(리포트/I/O 전용 필드 제외)·지표 열 목록·인덱스 범위·
    값 체크섬. 이 중 하나라도 다르면 다른 키(재계산). 값 체크섬은 nan을 제외한 합/제곱합/결측수."""
    import json
    h = hashlib.sha1()
    h.update(f"{VERSION}|{getattr(M, 'BUNDLE_VERSION', '?')}|{ticker}".encode())
    cfg_d = {k: v for k, v in dataclasses.asdict(cfg_i).items() if k not in _CACHE_KEY_IGNORE_FIELDS}
    h.update(json.dumps(cfg_d, sort_keys=True, default=str).encode())
    h.update("|".join(map(str, ind.columns)).encode())
    h.update(f"{ind.index[0]}|{ind.index[-1]}|{len(ind)}".encode())
    arr = ind.to_numpy(dtype=float)
    h.update(np.array([np.nansum(arr), np.nansum(arr * arr), float(np.isnan(arr).sum())]).tobytes())
    pa = px_adj.to_numpy(dtype=float)
    h.update(np.array([np.nansum(pa), np.nansum(pa * pa), float(np.isnan(pa).sum())]).tobytes())
    return h.hexdigest()


def validate_and_weight_sector(ticker: str, ind_i: pd.DataFrame, adj_i: pd.Series, cfg_i,
                               specs: List[Any], M, scfg: SectorConfig) -> Dict[str, Any]:
    """[M.run() 3)+4)] 전체표본 검증표(03시트용, asof=마지막일 감쇠가중) + 워크포워드 가중치
    (수익 W·위험 W_haz·재추정 로그). 비용 지배 단계 — 입력 해시 키로 디스크 캐시한다."""
    t0 = time.time()
    key = _cache_key(ticker, cfg_i, ind_i, adj_i, M)
    path = os.path.join(scfg.CACHE_DIR, f"{ticker}_{key[:16]}.pkl.gz")
    if scfg.USE_CACHE and os.path.exists(path):
        try:
            cached = pd.read_pickle(path, compression="gzip")
            if cached.get("key") == key and all(k in cached for k in ("val_full", "W", "wlog", "W_haz")):
                log("CACHE", kv(ticker=ticker, event="cache_hit", file=os.path.basename(path),
                                periods=len(cached["wlog"]), elapsed_s=round(time.time() - t0, 2)), M=M)
                cached["cache_hit"] = True
                return cached
        except Exception as e:
            log("CACHE", kv(ticker=ticker, event="cache_read_failed", err=type(e).__name__), M=M, level="warning")
    idx = ind_i.index
    with _indicator_spec_override(M, specs):
        w_report = M.decay_weights(idx, asof=idx[-1], half_life_days=cfg_i.HALF_LIFE_DAYS)
        t1 = time.time()
        val_full = M.validate_indicators(ind_i, adj_i, cfg_i, weights=w_report, verbose=False)
        t2 = time.time()
        W, wlog, W_haz = M.build_walkforward_weights(ind_i, adj_i, cfg_i)
        t3 = time.time()
    out = {"key": key, "val_full": val_full, "W": W, "wlog": wlog, "W_haz": W_haz, "cache_hit": False,
           "t_validate": round(t2 - t1, 2), "t_walkforward": round(t3 - t2, 2)}
    log("VALIDATE", kv(ticker=ticker, candidates=len(specs), strict_pass=int((val_full["판정"] == "PASS").sum()),
                       reestimations=len(wlog), t_validate_s=out["t_validate"], t_walkforward_s=out["t_walkforward"]), M=M)
    if scfg.USE_CACHE:
        try:
            os.makedirs(scfg.CACHE_DIR, exist_ok=True)
            tmp = f"{path}.tmp{os.getpid()}"
            pd.to_pickle(out, tmp, compression="gzip", protocol=4)
            os.replace(tmp, path)     # 원자적 교체 — 중단된 실행이 잘린 캐시 파일을 남기지 않게
            log("CACHE", kv(ticker=ticker, event="cache_saved", file=os.path.basename(path),
                            size_mb=round(os.path.getsize(path) / 1e6, 1)), M=M)
        except Exception as e:
            log("CACHE", kv(ticker=ticker, event="cache_write_failed", err=type(e).__name__), M=M, level="warning")
    return out


def _score_with_weights(Z_row: pd.Series, w_row: pd.Series) -> float:
    """M.lookahead_audit()과 동일한 한 날짜 점수 재계산식."""
    avail = Z_row.notna() & (w_row != 0)
    den = float((w_row.abs() * avail).sum())
    return float((Z_row * w_row).where(avail).sum() / den) if den > 0 else np.nan


def sector_lookahead_audit(ticker: str, res: dict, M, scfg: SectorConfig, cfg_i,
                           raw_df: pd.DataFrame, spy_raw_df: pd.DataFrame,
                           W: pd.DataFrame, W_haz: pd.DataFrame,
                           score_full: pd.Series, haz_full: pd.Series, n_dates: int,
                           breadth: Optional[pd.Series] = None) -> pd.DataFrame:
    """[M.lookahead_audit과 같은 방법] 무작위 검사일 d마다 '섹터·SPY 원시가격, M 지표값, M 점수'를
    d까지로 잘라 섹터 후보지표를 처음부터 다시 만들고(z-score 포함) 그날 가중치(W.loc[d])로 점수를
    재계산해 전체계산 점수와 비교한다. 섹터 지표 구성(롤링·ewm·베타 지연·상대가격·발표지연 시리즈)
    어디에도 d 이후 정보가 섞이지 않았음을 확인하는 감사. M 지표 자체의 인과성은 M의 11시트가 담당."""
    rng = np.random.default_rng(cfg_i.RANDOM_SEED)
    cal = res["cal"]
    valid = score_full.dropna().index
    valid = valid[(valid >= pd.Timestamp(cfg_i.SIGNAL_START))]
    if len(valid) > 15:
        valid = valid[:-15]     # 마지막 15거래일은 Adj Close 지연 이어붙임 구간과 겹칠 수 있어 제외
    if len(valid) < 30:
        return pd.DataFrame([{"티커": ticker, "결과": "감사 생략(표본 부족)"}])
    picks = sorted(rng.choice(valid, size=min(n_dates, len(valid)), replace=False))
    rows: List[dict] = []
    for d in picks:
        d = pd.Timestamp(d)
        cal_t = cal[cal <= d]
        sec_t = raw_df.loc[raw_df.index <= d]
        spy_t = spy_raw_df.loc[spy_raw_df.index <= d]
        price_t, idx_t, _ = sector_price_frame(sec_t, cal_t, scfg)
        spy_tr_t, _ = build_total_return_close(spy_t, cal_t, scfg.ADJ_CLOSE_STALE_DAYS)
        spy_raw_t = spy_t["Close"].reindex(cal_t).ffill()
        res_t = {"ind": res["ind"].loc[res["ind"].index <= d],
                 "fred": {k: (v.loc[v.index <= d] if v is not None else None) for k, v in res["fred"].items()},
                 "px_dict": {k: (v.loc[v.index <= d] if v is not None else None) for k, v in res["px_dict"].items()},
                 "score": res["score"].loc[res["score"].index <= d],
                 "haz_score": res["haz_score"].loc[res["haz_score"].index <= d],
                 "cfg": res["cfg"]}
        spy_series_t = spy_layer_series(res_t, M)
        # [v0.3.0 §1.E-2] 섹터 폭도 d까지로 절단해 재계산 감사에 포함 — d 이후 다른 섹터의 정보가
        # 섞여 있지 않은지 확인(다른 섹터 지표와 동일한 인과성 기준 적용).
        breadth_t = breadth.loc[breadth.index <= d] if breadth is not None else None
        ind_t, _ = build_sector_candidates(ticker, res_t, M, scfg, price_t["Adj Close"], price_t["Close"],
                                           spy_tr_t, spy_raw_t, spy_series_t, idx_t, breadth=breadth_t)
        Z_t = pd.DataFrame({k: M.expanding_zscore(ind_t[k]) for k in ind_t.columns}, index=ind_t.index)
        z_row = Z_t.loc[d]
        for kind, Wm, full in (("복합점수", W, score_full), ("위험점수(H)", W_haz, haz_full)):
            w_row = Wm.loc[d].reindex(z_row.index).fillna(0.0)
            s_t = _score_with_weights(z_row, w_row)
            s_full = float(full.loc[d]) if d in full.index else np.nan
            if kind == "위험점수(H)" and pd.isna(s_t) and pd.isna(s_full):
                rows.append({"티커": ticker, "검사일": str(d.date()), "전체계산 점수": np.nan, "절단재계산 점수": np.nan,
                             "차이": np.nan, "일치": "N/A(위험지표 미채택)", "감사종류": kind})
                continue
            diff = abs(s_t - s_full) if pd.notna(s_t) and pd.notna(s_full) else np.nan
            ok = bool(pd.notna(diff) and diff < 1e-8)
            rows.append({"티커": ticker, "검사일": str(d.date()), "전체계산 점수": round(s_full, 8) if pd.notna(s_full) else np.nan,
                         "절단재계산 점수": round(s_t, 8) if pd.notna(s_t) else np.nan,
                         "차이": diff, "일치": "OK" if ok else "불일치", "감사종류": kind})
            log("AUDIT", kv(ticker=ticker, kind=kind, date=str(d.date()), full=s_full if pd.notna(s_full) else -99,
                            truncated=s_t if pd.notna(s_t) else -99, result="OK" if ok else "MISMATCH"), M=M,
                level="info" if ok else "error")
    return pd.DataFrame(rows)


def run_sector(ticker: str, ctx: Dict[str, Any]) -> Dict[str, Any]:
    """섹터 1개 전체 파이프라인. 반환 dict는 pandas/기본형만 담는다(프로세스 경계 통과 —
    Config/IndicatorSpec 인스턴스 없음). 시트 조각(01~11)도 여기서 만들어 부모는 조립만 한다."""
    M, res, scfg = ctx["M"], ctx["res"], ctx["scfg"]
    M_cfg = res["cfg"]
    price_i, idx_i, info = ctx["frames"][ticker]
    raw_df = ctx["raw"][ticker]
    spy_tr, spy_raw, spy_series, rf = ctx["spy_tr"], ctx["spy_raw"], ctx["spy_series"], ctx["rf_daily"]
    cfg_i = sector_cfg_for(M_cfg, scfg, ticker, idx_i)
    np.random.seed(cfg_i.RANDOM_SEED)
    t_all = time.time()
    timing: Dict[str, float] = {}
    log("START", kv(ticker=ticker, name=SECTOR_NAME_KR.get(ticker, ""), data_start=str(idx_i[0].date()),
                    data_end=str(idx_i[-1].date()), rows=len(idx_i), signal_start=cfg_i.SIGNAL_START,
                    reweight_freq=cfg_i.REWEIGHT_FREQ, train_min_years=cfg_i.TRAIN_MIN_YEARS), M=M)

    adj_i = price_i["Adj Close"].astype(float)
    close_i = price_i["Close"].astype(float)
    sig_mask = pd.Series(idx_i >= pd.Timestamp(cfg_i.SIGNAL_START), index=idx_i)

    # ---- 지표 ---- [v0.3.0 §1.E-2] breadth: run()에서 11섹터 공통으로 1회 계산된 값(재계산 없음).
    t0 = time.time()
    breadth = ctx.get("sector_breadth_200")
    ind_i, specs = build_sector_candidates(ticker, res, M, scfg, adj_i, close_i, spy_tr, spy_raw, spy_series, idx_i,
                                           breadth=breadth)
    trend200 = sector_technical_values(adj_i)["TREND_200"]
    timing["04_지표생성"] = round(time.time() - t0, 2)
    log("INDICATOR", kv(ticker=ticker, candidates=len(specs),
                        market=sum(1 for s in specs if not s.key.startswith(f"{ticker}__")),
                        sector=sum(1 for s in specs if s.key.startswith(f"{ticker}__")),
                        with_data=int((ind_i.notna().sum() > 0).sum())), M=M)

    # ---- 검증 + 워크포워드 (캐시) ---- cfg_i(미조정, §1.C 배율 적용 전) 기준 — 캐시키는 아래 §1.C
    # 사이징 오버레이 임계값 변경과 무관해야 한다(그 필드들은 검증/워크포워드에 관여하지 않음).
    t0 = time.time()
    hv = validate_and_weight_sector(ticker, ind_i, adj_i, cfg_i, specs, M, scfg)
    val_full, W, wlog, W_haz = hv["val_full"], hv["W"], hv["wlog"], hv["W_haz"]
    timing["05_06_검증+워크포워드"] = round(time.time() - t0, 2)

    # ---- [v0.3.0 §1.C ⚠] 사이징 오버레이 변동성 정규화 — 이 지점부터 cfg_i는 '실제 적용된' cfg를
    # 가리킨다(신호·검증 게이트는 불변, 과열헤어컷/회복확인폭/깊은낙폭 임계값만 섹터 변동성에 맞춰 조정).
    vol_scale = _sector_vol_scale(adj_i, spy_tr, scfg) if scfg.VOL_SCALE_OVERLAYS else 1.0
    cfg_i = _vol_scaled_cfg(cfg_i, vol_scale)
    log("SIZING", kv(ticker=ticker, vol_scale=round(vol_scale, 3),
                     haircut_steps=str(cfg_i.EXTENSION_HAIRCUT_STEPS),
                     recovery_confirm_pct=cfg_i.RECOVERY_CONFIRM_PCT, deep_recovery_dd=cfg_i.DEEP_RECOVERY_DD,
                     struct_bottom_dd=cfg_i.STRUCT_BOTTOM_DD, hazard_source=scfg.HAZARD_SOURCE), M=M)

    # ---- 복합점수 / 위험점수 ---- [v0.3.0 §1.B ⚠] HAZARD_SOURCE에 따라 신호에 쓰이는 haz_pct의
    # 소스를 전환한다(§0.3 근거). 섹터 자체 haz_score/haz_pct_sector는 그대로 계산해 01시트에
    # 진단용으로 나란히 남긴다 — '신호에 쓰이는지'만 바뀐다.
    t0 = time.time()
    with _indicator_spec_override(M, specs):
        score, contrib, n_used = M.composite_score(ind_i, W, cfg_i)
        haz_score, haz_contrib, _ = M.composite_score(ind_i, W_haz, cfg_i)
    score_pct = M.score_percentile(score).where(sig_mask)
    haz_pct_sector = M.score_percentile(haz_score).where(sig_mask)
    haz_pct_spy = res["haz_pct"].reindex(idx_i).where(sig_mask)
    if scfg.HAZARD_SOURCE == "sector":
        haz_pct = haz_pct_sector
    elif scfg.HAZARD_SOURCE == "max":
        haz_pct = pd.concat([haz_pct_sector, haz_pct_spy], axis=1).max(axis=1, skipna=True).where(sig_mask)
    else:  # "spy"(기본) — M이 SPY에서 이미 검증·교정한 H를 그대로 재사용(재계산 없음)
        haz_pct = haz_pct_spy
    fast_pct = None
    if cfg_i.USE_FAST_TRIGGER and res.get("fast_pct") is not None:
        fast_pct = res["fast_pct"].reindex(idx_i)      # M과 동일 지표(VIX_TERM) — 시장 급락트리거를 그대로 공유
    recov_conf = deep_recov = struct_dd = None
    if cfg_i.USE_RECOVERY_FLOOR:
        roll_low = close_i.rolling(cfg_i.RECOVERY_LOW_WINDOW, min_periods=20).min()
        recov_conf = close_i >= roll_low * (1.0 + cfg_i.RECOVERY_CONFIRM_PCT)
        struct_dd = M.deep_drawdown_flag(close_i, cfg_i.STRUCT_BOTTOM_DD, cfg_i.RECOVERY_LOW_WINDOW, mode="peak_to_trough")
        if cfg_i.USE_DEEP_RECOVERY_BOOST:
            deep = M.deep_drawdown_flag(close_i, cfg_i.DEEP_RECOVERY_DD, cfg_i.RECOVERY_LOW_WINDOW, mode=cfg_i.DEEP_RECOVERY_DD_MODE)
            deep_recov = recov_conf & deep

    # ---- 신호 / 백테스트 ----
    sig = M.generate_signals(score_pct, trend200, cfg_i, score=score, haz_pct=haz_pct, fast_pct=fast_pct,
                             recov_conf=recov_conf, deep_recov=deep_recov, struct_dd=struct_dd)
    with _indicator_spec_override(M, specs):
        reason = M.build_reason_text(contrib, sig["state"], score)
    bt = M.run_backtest(price_i, sig["target_pos"], cfg_i, rf)
    bt = bt.loc[bt.index >= pd.Timestamp(cfg_i.SIGNAL_START)]
    ma_pos = (trend200 > 0).astype(float).where(sig_mask, 0.0)
    bt_ma = M.run_backtest(price_i, ma_pos, cfg_i, rf)
    bt_ma = bt_ma.loc[bt_ma.index >= pd.Timestamp(cfg_i.SIGNAL_START)]
    timing["07_08_신호+백테스트"] = round(time.time() - t0, 2)

    # ---- 민감도 / 감사 / 이벤트 / 거래 / 구간 ----
    t0 = time.time()
    sens = pd.DataFrame()
    if scfg.RUN_THRESHOLD_SENSITIVITY:
        sens = M.threshold_sensitivity(score_pct, trend200, price_i, cfg_i, rf, haz_pct=haz_pct, fast_pct=fast_pct,
                                       recov_conf=recov_conf, deep_recov=deep_recov, struct_dd=struct_dd)
    timing["09_임계값민감도"] = round(time.time() - t0, 2)
    t0 = time.time()
    audit = pd.DataFrame()
    if scfg.RUN_LOOKAHEAD_AUDIT:
        audit = sector_lookahead_audit(ticker, res, M, scfg, cfg_i, raw_df, ctx["spy_raw_df"], W, W_haz,
                                       score, haz_score, scfg.AUDIT_SAMPLE, breadth=breadth)
    adopted = sorted({k for k in W.columns if (W[k] != 0).any()})
    events = M.event_study(ind_i, adj_i, bt, adopted, cfg_i, haz_pct=haz_pct)
    trades = M.extract_trades(bt, reason, sig["state"])
    episodes = M.drawdown_episodes(bt, cfg_i)
    timing["11_감사+이벤트+구간"] = round(time.time() - t0, 2)

    # [v0.3.0 §1.G-2] SPY(M) 같은 해 평균 목표비중 — 07_연도별성과에 병기(§0.2 표 자동 재현).
    spy_yearly_pos = res["bt"]["pos_exec"].groupby(res["bt"].index.year).mean()

    # ---- 시트 조각 ----
    sheets = build_sector_sheets(M, ticker, cfg_i, specs, price_i, bt, bt_ma, sig, score, score_pct, n_used,
                                 haz_score, haz_pct, fast_pct, recov_conf, reason, ind_i, contrib, W, W_haz,
                                 wlog, val_full, adopted, trades, episodes, events, sens, audit,
                                 haz_pct_sector=haz_pct_sector, spy_yearly_pos=spy_yearly_pos)
    timing["12_run_sector()합계"] = round(time.time() - t_all, 2)
    first_signal = score_pct.dropna().index[0] if score_pct.notna().any() else None
    log("DONE", kv(ticker=ticker, adopted_ever=len(adopted), strict_pass=int((val_full["판정"] == "PASS").sum()),
                   first_signal=str(first_signal.date()) if first_signal is not None else "-",
                   hazard_source=scfg.HAZARD_SOURCE, vol_scale=round(vol_scale, 3),
                   cache_hit=hv.get("cache_hit", False), elapsed_s=timing["12_run_sector()합계"]), M=M)
    return {"ticker": ticker, "info": info, "cfg_dict": dataclasses.asdict(cfg_i), "timing": timing,
            "cache_hit": bool(hv.get("cache_hit", False)), "n_candidates": len(specs),
            "first_signal": (str(first_signal.date()) if first_signal is not None else None),
            "adopted": adopted, "sheets": sheets,
            "hazard_source": scfg.HAZARD_SOURCE, "vol_scale": vol_scale,   # [v0.3.0 §1.B/§1.C]
            # 통합 시트용 소형 시리즈
            "state": sig["state"].loc[sig.index >= pd.Timestamp(cfg_i.SIGNAL_START)],
            "target_pos": sig["target_pos"].loc[sig.index >= pd.Timestamp(cfg_i.SIGNAL_START)],
            "score_pct": score_pct.loc[sig_mask], "haz_pct": haz_pct.loc[sig_mask],
            "haz_pct_sector": haz_pct_sector.loc[sig_mask],   # [v0.3.0 §1.B] 진단용(신호 미관여)
            "strategy_ret": bt["strategy_ret"], "bh_ret": bt["bh_ret"], "pos_exec": bt["pos_exec"],
            "ma_ret": bt_ma["strategy_ret"]}


# =============================================================================
# [7] 섹터별 시트 조각 — M.build_report()의 01~11 시트 생성 로직을 섹터에 그대로 적용(+'티커' 열)
# =============================================================================
# [v0.3.0 §1.G-1] 09b_규칙별기여 — 원인 진단을 시트에서 바로 확인(§0.2/§0.3 표를 매 실행 자동 산출).
_CONTRIB_RULES: List[Tuple[str, str]] = [
    ("hazard_entry", "위험회피진입①(H)"), ("fast_trigger", "급락트리거⓪(FT)"),
    ("neutral_risk_cut", "중립감축⑨"), ("extension_haircut", "과열헤어컷⑩"),
    ("trend_promotion", "추세승격⑥"), ("recovery_floor", "회복승격⑤"),
]


def build_rule_contribution(ticker: str, sig: pd.DataFrame, bt: pd.DataFrame) -> pd.DataFrame:
    """[v0.3.0 §1.G-1] 규칙별 발동일수·발동일의 '익일' B&H수익 평균(그 규칙이 비중을 줄이거나
    늘린 날 다음날 실제로 시장이 어느 방향으로 움직였는지 — §0.3 H진입일 검증과 같은 방법),
    비중=0(현금)인 날의 B&H 수익 합(=그 구간 동안 놓친/피한 수익, §0.2)을 상승일/하락일로
    나누어 자동 산출한다. 12_섹터요약의 신규 3열이 이 표에서 값을 가져다 쓴다(단일 소스)."""
    idx = bt.index
    nxt = bt["bh_ret"].shift(-1)
    rows: List[dict] = []
    for col, label in _CONTRIB_RULES:
        if col not in sig.columns:
            continue
        fired = sig[col].reindex(idx).fillna(False).astype(bool)
        n = int(fired.sum())
        avg_next = float(nxt.where(fired).mean()) if n > 0 else np.nan
        rows.append({"티커": ticker, "규칙": label, "발동일수": n,
                     "발동일 익일B&H평균수익(%)": round(avg_next * 100, 4) if pd.notna(avg_next) else np.nan,
                     "비중=0인 날 B&H수익 합(%p)": np.nan})
    zero_mask = bt["pos_exec"] <= 1e-9
    up_mask = zero_mask & (bt["bh_ret"] > 0)
    dn_mask = zero_mask & (bt["bh_ret"] < 0)
    rows.append({"티커": ticker, "규칙": "[전체] 비중=0(현금)인 날", "발동일수": int(zero_mask.sum()),
                 "발동일 익일B&H평균수익(%)": np.nan,
                 "비중=0인 날 B&H수익 합(%p)": round(float(bt["bh_ret"].where(zero_mask).sum()) * 100, 4)})
    rows.append({"티커": ticker, "규칙": "[전체] 상승일 미탑승(비중=0 & B&H>0, %p)", "발동일수": int(up_mask.sum()),
                 "발동일 익일B&H평균수익(%)": np.nan,
                 "비중=0인 날 B&H수익 합(%p)": round(float(bt["bh_ret"].where(up_mask).sum()) * 100, 4)})
    rows.append({"티커": ticker, "규칙": "[전체] 하락일 회피(비중=0 & B&H<0, %p — 음수=회피한 손실)", "발동일수": int(dn_mask.sum()),
                 "발동일 익일B&H평균수익(%)": np.nan,
                 "비중=0인 날 B&H수익 합(%p)": round(float(bt["bh_ret"].where(dn_mask).sum()) * 100, 4)})
    return pd.DataFrame(rows)


def _append_sector_next_day_row(daily: pd.DataFrame, nd: dict) -> pd.DataFrame:
    """[v0.3.0 §1.A] M._append_next_day_row()와 같은 패턴 — S의 01_일별_{티커} 컬럼명(섹터상황/예측,
    M은 시장상황 하나뿐)에 맞춘 버전. 새 계산 없음: M.build_next_day_prediction()이 만든 nd를 표시만."""
    row = {c: ("" if daily[c].dtype == object else np.nan) for c in daily.columns}
    if "날짜" in row:
        row["날짜"] = nd["다음거래일"].date()
    if "섹터상황" in row:
        row["섹터상황"] = nd["확정국면"]
    if "예측" in row:
        row["예측"] = STATE_SHORT.get(nd["확정국면_원시"], nd["확정국면_원시"])
    if "복합점수" in row:
        row["복합점수"] = nd["복합점수"]
    if "복합점수백분위" in row:
        row["복합점수백분위"] = nd["복합점수백분위"]
    if "목표비중" in row:
        row["목표비중"] = nd["목표비중"]
    if "체결비중" in row:
        row["체결비중"] = np.nan
    if "매매행동" in row:
        row["매매행동"] = nd["예상행동_en"]
    if "위험점수백분위(H)" in row:
        row["위험점수백분위(H)"] = nd["위험점수백분위(H)"]
    if "근거요약" in row:
        row["근거요약"] = (f"다음 거래일 예측(전일 종가 신호). 발동 규칙: {nd['발동규칙']}. "
                          f"{nd['근거요약']}{nd['기준일_경과주의']}")
    row_df = pd.DataFrame([row])[list(daily.columns)]
    return pd.concat([daily, row_df], ignore_index=True)


def build_sector_sheets(M, ticker, cfg_i, specs, price_i, bt, bt_ma, sig, score, score_pct, n_used,
                        haz_score, haz_pct, fast_pct, recov_conf, reason, ind_i, contrib, W, W_haz,
                        wlog, val_full, adopted, trades, episodes, events, sens, audit,
                        haz_pct_sector: Optional[pd.Series] = None,
                        spy_yearly_pos: Optional[pd.Series] = None) -> Dict[str, pd.DataFrame]:
    idx = bt.index
    name_map = {s.key: s.name_kr for s in specs}
    _flag = lambda col: (sig[col].reindex(idx).map({True: "발동", False: ""}) if col in sig.columns else "")

    # ---- 01 일별기록 (M 01과 동일 컬럼) ----
    daily = pd.DataFrame(index=idx)
    daily["날짜"] = [d.date() for d in idx]
    daily["시가"] = bt["Open"].round(2)
    daily["고가"] = price_i["High"].reindex(idx).round(2) if "High" in price_i.columns else np.nan
    daily["저가"] = price_i["Low"].reindex(idx).round(2) if "Low" in price_i.columns else np.nan
    daily["종가"] = bt["Close"].round(2)
    daily["일간등락률"] = (bt["ret_cc"] * 100).round(3)
    daily["섹터상황"] = sig["state"].reindex(idx).map(STATE_KR).fillna("-")
    daily["예측"] = sig["state"].reindex(idx).map(STATE_SHORT).fillna("-")
    daily["복합점수"] = score.reindex(idx).round(4)
    daily["복합점수백분위"] = score_pct.reindex(idx).round(4)
    daily["사용지표수"] = n_used.reindex(idx)
    daily["목표비중"] = sig["target_pos"].reindex(idx).round(2)
    daily["체결비중"] = bt["pos_exec"].round(2)
    daily["매매행동"] = M.action_labels(bt["pos_exec"])
    daily["추세오버라이드"] = sig["trend_override"].reindex(idx).map({True: "발동", False: ""})
    daily["위험점수(H)"] = haz_score.reindex(idx).round(4)
    daily["위험점수백분위(H)"] = haz_pct.reindex(idx).round(4)
    # [v0.3.0 §1.B] 신호에 실제로 쓰인 H(HAZARD_SOURCE에 따라 spy/sector/max)와 별개로, 섹터 자체
    # H는 진단 목적으로 표시만 나란히 남긴다(§1.B.2 — 신호에는 관여하지 않음).
    if haz_pct_sector is not None:
        daily["위험점수백분위(H,섹터자체)"] = haz_pct_sector.reindex(idx).round(4)
    daily["위험회피진입(H)"] = _flag("hazard_entry")
    daily["위험선호차단(H)"] = _flag("hazard_block")
    daily["위험회피해제안전판(H)"] = _flag("hazard_floor")
    daily["매수보류게이트(H)"] = _flag("buy_hold_gate")
    daily["급락트리거백분위"] = fast_pct.reindex(idx).round(4) if fast_pct is not None else np.nan
    daily["급락트리거(FT)"] = _flag("fast_trigger")
    daily["회복확인(저점대비)"] = (recov_conf.reindex(idx).map({True: "확인", False: ""}) if recov_conf is not None else "")
    daily["회복승격(R)"] = _flag("recovery_floor")
    daily["추세승격(T)"] = _flag("trend_promotion")
    daily["속도경보(ΔH)"] = _flag("hazard_velocity")
    daily["중립감축(H)"] = _flag("neutral_risk_cut")
    if "extension_haircut" in sig.columns:
        _eh = sig["extension_haircut"].reindex(idx).fillna(False).astype(bool)
        _ec = sig["ext_cap"].reindex(idx)
        daily["과열헤어컷(E)"] = np.where(_eh, "상한 " + _ec.round(1).astype(str), "")
    else:
        daily["과열헤어컷(E)"] = ""
    daily["레버리지(L)"] = _flag("leverage")
    daily["깊은낙폭회복(D)"] = _flag("deep_recovery")
    daily["전략일간수익"] = (bt["strategy_ret"] * 100).round(3)
    daily["전략자산곡선"] = (bt["equity"] / bt["equity"].iloc[0]).round(4)
    daily["섹터자산곡선"] = (bt["bh_equity"] / bt["bh_equity"].iloc[0]).round(4)
    daily["전략낙폭"] = (bt["equity"] / bt["equity"].cummax() - 1).round(4)
    daily["섹터낙폭"] = (bt["bh_equity"] / bt["bh_equity"].cummax() - 1).round(4)
    daily["근거요약"] = reason.reindex(idx)
    for k in adopted:
        daily[f"[값]{name_map.get(k, k)}"] = ind_i[k].reindex(idx).round(4)
        daily[f"[기여]{name_map.get(k, k)}"] = contrib[k].reindex(idx).round(4)
    daily = daily.reset_index(drop=True)

    # [v0.3.0 §1.A] 다음 거래일 예측 1행 — M.build_next_day_prediction()을 그대로 재사용(재계산 없음,
    # 새 계산이 아니라 sig/bt의 마지막 값을 표시만 재구성). bt/성과 시트는 무관 — 이 변수는 daily에만 반영.
    next_day = None
    try:
        if hasattr(M, "build_next_day_prediction"):
            next_day = M.build_next_day_prediction(
                {"bt": bt, "sig": sig, "score": score, "score_pct": score_pct, "haz_pct": haz_pct, "reason": reason},
                cfg_i)
            daily = _append_sector_next_day_row(daily, next_day)
    except Exception as e:
        log("REPORT", kv(ticker=ticker, event="next_day_row_failed", err=type(e).__name__, msg=str(e)[:160]),
            M=M, level="warning")

    # [v0.3.0 §1.G-1] 09b_규칙별기여 — bt(SIGNAL_START 이후)와 실제 신호에 쓰인 sig 기준(§1.B 반영).
    rule_contrib = build_rule_contribution(ticker, sig, bt)

    # ---- 03 지표검증 / 04 채택근거상세 ----
    val_cols = ["지표코드", "지표명", "카테고리", "검증트랙", "평가지평", "판정", "판정사유",
                "위험트랙판정", "위험트랙AUC엣지", "위험트랙AUC방향안정성", "위험트랙판정사유",
                "사용데이터", "커버리지", "커버리지(가중)", "표본수", "N_eff", "사전방향", "실증방향",
                "사전방향일치", "IC_평가지평", "IC_5d", "IC_21d", "IC_63d", "NW_t", "NW_p",
                "Q1수익률", "Q2수익률", "Q3수익률", "Q4수익률", "Q5수익률", "Q5-Q1", "단조성",
                "하락AUC", "AUC엣지", "부호일치구간수",
                "짧은이력완화적용", "유효최소커버리지", "유효최소구간일치수", "위험트랙유효최소구간일치수",
                "구간1IC", "구간2IC", "구간3IC", "구간4IC", "선행메커니즘"]
    val_sheet = val_full[[c for c in val_cols if c in val_full.columns]].copy()
    val_sheet["채택이력(워크포워드)"] = val_sheet["지표코드"].map(lambda k: "채택된 적 있음" if k in adopted else "")
    val_sheet = val_sheet.sort_values(["판정", "AUC엣지"], ascending=[True, False])
    detail_rows = []
    for _, r in val_full.iterrows():
        is_trend = bool(r.get("추세트랙", False))
        detail_rows.append({
            "지표코드": r["지표코드"], "지표명": r["지표명"], "카테고리": r["카테고리"], "검증트랙": r.get("검증트랙", "-"),
            "판정": r["판정"], "왜 이 지표를 후보로 넣었나(경제적 근거)": r["경제적근거"], "선행 메커니즘": r["선행메커니즘"],
            "사용 데이터": r["사용데이터"],
            "검증 결과 요약": (("[추세트랙 - IC/NW-t 게이트 면제] " if is_trend else "[표준트랙] ")
                          + f"평가지평={r.get('평가지평', cfg_i.VAL_PRIMARY_H)}일, 지평별 IC={r.get('IC_평가지평')}, "
                          + ("" if is_trend else f"Newey-West t={r.get('NW_t')} (p={r.get('NW_p')}), ")
                          + f"5분위 스프레드={r.get('Q5-Q1')}, 단조성={r.get('단조성')}, 하락판별 AUC={r.get('하락AUC')}, "
                          f"시대별 부호일치={r.get('부호일치구간수')}, 커버리지={r.get('커버리지')}"),
            "최종 판정 사유": r["판정사유"]})
    detail = pd.DataFrame(detail_rows).sort_values(["판정", "카테고리"])

    # ---- 06 성과 / 06b 운용통계 / 07 연도별 ----
    perf = pd.DataFrame([M.perf_metrics(bt["strategy_ret"], f"{ticker} 복합지표 전략"),
                         M.perf_metrics(bt["bh_ret"], f"{ticker} 단순보유(Buy&Hold)"),
                         M.perf_metrics(bt_ma["strategy_ret"], f"{ticker} 200일선 단독(벤치마크)")])
    extra = pd.DataFrame([
        {"항목": "총 거래(포지션 변경) 횟수", "값": int((bt["turnover"] > 1e-9).sum())},
        {"항목": "연평균 회전율(편도)", "값": round(float(bt["turnover"].sum() / (len(bt) / 252)), 2)},
        {"항목": "누적 거래비용(수익률 차감분)", "값": round(float(bt["cost"].sum()), 4)},
        {"항목": "시장 투자 시간 비율", "값": round(float((bt["pos_exec"] > 0).mean()), 4)},
        {"항목": "평균 비중", "값": round(float(bt["pos_exec"].mean()), 4)},
        {"항목": "채택 지표 수(기간 중 1회 이상)", "값": len(adopted)},
        {"항목": "채택 지표", "값": ", ".join(name_map.get(k, k) for k in adopted)},
    ])
    yr = pd.DataFrame({"연도": [d.year for d in idx], "s": bt["strategy_ret"].values, "b": bt["bh_ret"].values,
                       "pos": bt["pos_exec"].values, "turn": bt["turnover"].values})
    ann = yr.groupby("연도").agg(
        전략수익률=("s", lambda x: round(float((1 + x).prod() - 1), 4)),
        섹터수익률=("b", lambda x: round(float((1 + x).prod() - 1), 4)),
        평균비중=("pos", lambda x: round(float(x.mean()), 3)),
        거래횟수=("turn", lambda x: int((x > 1e-9).sum()))).reset_index()
    ann["초과수익"] = (ann["전략수익률"] - ann["섹터수익률"]).round(4)
    dd_y = bt.groupby(bt.index.year).apply(
        lambda g: round(float(((1 + g["strategy_ret"]).cumprod() / (1 + g["strategy_ret"]).cumprod().cummax() - 1).min()), 4))
    ann["전략연중최대낙폭"] = ann["연도"].map(dd_y)
    # [v0.3.0 §1.G-2] SPY(M) 같은 해 평균 목표비중 병기 — §0.2 표를 매 실행 자동 재현(강세장 과소투자 진단).
    if spy_yearly_pos is not None:
        ann["SPY평균비중"] = ann["연도"].map(spy_yearly_pos).round(3)

    # ---- 08 워크포워드 가중치 ----
    wf = pd.DataFrame(wlog)
    if len(wf):
        rename_map = {}
        for c in wf.columns:
            if c.startswith("[해저드]") and c[5:] in name_map:
                rename_map[c] = f"[해저드]{name_map[c[5:]]}"
            elif c in name_map:
                rename_map[c] = name_map[c]
        wf = wf.rename(columns=rename_map)

    # ---- 09 국면통계 ----
    st = sig["state"].reindex(idx)
    nxt = bt["bh_ret"].shift(-1)
    reg = pd.DataFrame({"국면": st, "익일섹터수익": nxt}).dropna()
    regime_stats = reg.groupby("국면")["익일섹터수익"].agg(
        일수="count", 평균익일수익률="mean", 익일상승확률=lambda x: (x > 0).mean(), 변동성="std").reset_index()
    regime_stats["평균익일수익률"] = (regime_stats["평균익일수익률"] * 100).round(4)
    regime_stats["익일상승확률"] = regime_stats["익일상승확률"].round(4)
    regime_stats["변동성"] = (regime_stats["변동성"] * 100).round(4)
    regime_stats["연율화수익률"] = (regime_stats["평균익일수익률"] / 100 * 252).round(4)
    regime_stats["국면"] = regime_stats["국면"].map(STATE_KR).fillna(regime_stats["국면"])
    regime_validity: Optional[bool] = None
    try:
        r_on = float(regime_stats.loc[regime_stats["국면"] == STATE_KR["RISK_ON"], "평균익일수익률"].iloc[0])
        r_off = float(regime_stats.loc[regime_stats["국면"] == STATE_KR["RISK_OFF"], "평균익일수익률"].iloc[0])
        regime_validity = bool(r_on > r_off)
        verdict = (f"국면 정의 유효 (상승국면 익일평균 {r_on:.4f}% > 하락국면 {r_off:.4f}%)" if regime_validity
                   else f"주의: 하락국면의 익일수익률이 더 높음 ({r_off:.4f}% >= {r_on:.4f}%)")
    except Exception:
        verdict = "판정 불가(상승 또는 하락 국면 표본 없음)"
    regime_stats["국면정의 검증"] = verdict

    for df in (val_sheet, detail, perf, extra, ann, wf, regime_stats, trades, episodes, events, sens, audit,
              rule_contrib):
        if isinstance(df, pd.DataFrame) and len(df) and "티커" not in df.columns:
            df.insert(0, "티커", ticker)
    return {"daily": daily, "trades": trades, "val_sheet": val_sheet, "detail": detail, "events": events,
            "episodes": episodes, "perf": perf, "extra": extra, "sens": sens, "annual": ann, "wf": wf,
            "regime_stats": regime_stats, "audit": audit, "regime_validity": regime_validity,
            "regime_verdict": verdict, "n_strict_pass": int((val_full["판정"] == "PASS").sum()),
            "latest_adopted": (int(pd.to_numeric(pd.DataFrame(wlog)["채택지표수"], errors="coerce").dropna().iloc[-1])
                               if len(wlog) else 0),
            # [v0.3.0 §1.A/§1.G-1]
            "rule_contrib": rule_contrib, "next_day": next_day}


# =============================================================================
# [8] 병렬 실행 — fork 기반. 자식은 부모 메모리(res·프레임)를 복사-쓰기로 공유하고 결과만
#     큐로 돌려준다(함수 피클 없음 → 모듈이 sys.modules에 등록돼 있지 않아도 동작).
# =============================================================================
_CTX: Dict[str, Any] = {}


def _child_entry(ticker: str, q) -> None:
    try:
        out = run_sector(ticker, _CTX)
        q.put((ticker, "ok", out))
    except Exception as e:  # noqa
        q.put((ticker, "err", f"{type(e).__name__}: {e}\n{traceback.format_exc()[-3000:]}"))


def _resolve_workers(scfg: SectorConfig, n_tasks: int) -> int:
    if scfg.MAX_WORKERS and scfg.MAX_WORKERS > 0:
        n = scfg.MAX_WORKERS
    else:
        n = min(os.cpu_count() or 1, 4)
    return max(1, min(n, n_tasks))


def run_sectors(tickers: List[str], ctx: Dict[str, Any], scfg: SectorConfig, M
                ) -> Tuple[Dict[str, Dict[str, Any]], Dict[str, str]]:
    """섹터들을 순차 또는 fork 병렬로 실행. 반환: (성공 결과 dict, 실패 티커→에러 문자열)."""
    global _CTX
    results: Dict[str, Dict[str, Any]] = {}
    failed: Dict[str, str] = {}
    n_workers = _resolve_workers(scfg, len(tickers))
    can_fork = n_workers > 1 and sys.platform.startswith("linux")
    mpctx = None
    if can_fork:
        try:
            mpctx = mp.get_context("fork")
        except ValueError:
            can_fork = False
    if not can_fork:
        log("RUN", kv(event="sequential", n_sectors=len(tickers),
                      reason=("MAX_WORKERS=1" if n_workers == 1 else "fork 불가")), M=M)
        for t in tickers:
            try:
                results[t] = run_sector(t, ctx)
            except Exception as e:  # noqa
                failed[t] = f"{type(e).__name__}: {e}\n{traceback.format_exc()[-3000:]}"
                log("RUN", kv(ticker=t, event="sector_failed", err=type(e).__name__, msg=str(e)[:200]), M=M, level="error")
        return results, failed

    _CTX = ctx
    log("RUN", kv(event="parallel_fork", n_sectors=len(tickers), workers=n_workers), M=M)
    q = mpctx.Queue()
    pending = list(tickers)
    running: Dict[str, Any] = {}
    t_start: Dict[str, float] = {}
    while pending or running:
        while pending and len(running) < n_workers:
            t = pending.pop(0)
            p = mpctx.Process(target=_child_entry, args=(t, q), name=f"sector-{t}")
            p.start()
            running[t] = p
            t_start[t] = time.time()
        try:
            t, status, out = q.get(timeout=60)
        except Exception:
            # 60초 동안 결과가 없으면 큐에 아무것도 넣지 못하고 죽은 자식(OOM 등)이 있는지 확인
            for t_dead in [t for t, p in list(running.items()) if not p.is_alive() and p.exitcode not in (None, 0)]:
                p = running.pop(t_dead)
                failed[t_dead] = f"child exited with code {p.exitcode} (메모리 부족 가능 — MAX_WORKERS=1로 재시도)"
                log("RUN", kv(ticker=t_dead, event="child_died", exitcode=p.exitcode), M=M, level="error")
            continue
        p = running.pop(t)
        p.join(timeout=30)
        if status == "ok":
            results[t] = out
            log("RUN", kv(ticker=t, event="sector_done", elapsed_s=round(time.time() - t_start[t], 1),
                          remaining=len(pending) + len(running)), M=M)
        else:
            failed[t] = out
            log("RUN", kv(ticker=t, event="sector_failed", msg=str(out).splitlines()[0][:200]), M=M, level="error")
    _CTX = {}
    return results, failed


# =============================================================================
# [9] 최상위 실행 — run(res_or_path, M) → sres ; build_sector_report(sres) → xlsx
# =============================================================================
def _resolve_res(res_or_path, M) -> dict:
    if isinstance(res_or_path, str):
        if not hasattr(M, "load_result_bundle"):
            raise RuntimeError("market_regime_trader.py v1.22.0 이상이 필요합니다(load_result_bundle 없음) — "
                               "GitHub의 market_regime_trader.py를 최신으로 올린 뒤 다시 받으세요.")
        return M.load_result_bundle(res_or_path)
    res = res_or_path
    missing = [k for k in ("cfg", "ind", "score", "haz_score", "sig", "bt", "cal", "px_dict", "fred") if k not in res]
    if missing:
        raise KeyError(f"res에 필수 키가 없습니다: {missing} — M.run()의 반환값 또는 M.load_result_bundle() 결과를 넘기세요")
    return res


def run(res_or_path, M, scfg: Optional[SectorConfig] = None,
        sector_px_override: Optional[Dict[str, pd.DataFrame]] = None) -> Dict[str, Any]:
    """섹터 11개 전체 실행. res_or_path: M.run() 반환 dict 또는 결과 번들 경로(market_regime_result.pkl.gz).
    sector_px_override: 합성/오프라인 테스트용 가격 주입(fetch_sector_prices 참조)."""
    scfg = scfg or CFG
    t_all = time.time()
    res = _resolve_res(res_or_path, M)
    M_cfg = res["cfg"]
    cal = res["cal"]
    log("START", kv(version=VERSION, m_bundle=getattr(M, "BUNDLE_VERSION", "?"), sectors=len(scfg.SECTORS),
                    signal_start=M_cfg.SIGNAL_START, cal=f"{cal[0].date()}~{cal[-1].date()}",
                    reweight_freq=scfg.SECTOR_REWEIGHT_FREQ or M_cfg.REWEIGHT_FREQ,
                    train_min_years=scfg.SECTOR_TRAIN_MIN_YEARS, workers=_resolve_workers(scfg, len(scfg.SECTORS)),
                    use_cache=scfg.USE_CACHE, seed=scfg.RANDOM_SEED), M=M)
    stage_timing: Dict[str, float] = {}

    # ---- 0) 자기검사 ----
    t0 = time.time()
    st: Dict[str, Any] = {"passed": None, "n_true_pass": None, "n_noise_pass": None}
    if scfg.RUN_SELFTEST:
        st = run_selftest(M, scfg)
        if not st["passed"]:
            log("RUN", kv(event="abort_selftest_failed", true_pass=st["n_true_pass"], noise_pass=st["n_noise_pass"]),
                M=M, level="error")
            return {"selftest": st, "aborted": True, "sectors": {}, "failed": {}, "scfg": scfg}
    stage_timing["00_자기검사"] = round(time.time() - t0, 2)

    # ---- 1) 데이터 ----
    t0 = time.time()
    quality: List[dict] = []
    sector_px, yahoo_diag = fetch_sector_prices(res, M, quality, sector_px_override=sector_px_override)
    spy_df = res["px_dict"]["SPY"]
    spy_df = spy_df[~spy_df.index.duplicated(keep="last")].sort_index()
    spy_tr, _ = build_total_return_close(spy_df, cal, scfg.ADJ_CLOSE_STALE_DAYS)
    spy_raw = spy_df["Close"].reindex(cal).ffill()
    spy_series = spy_layer_series(res, M)
    rf_daily = None
    dgs = res["fred"].get("DGS3MO")
    if dgs is not None and dgs.notna().sum() > 100:
        rf_daily = (dgs / 100.0 / 252.0).reindex(cal).ffill().fillna(0.0)
    frames: Dict[str, Tuple[pd.DataFrame, pd.DatetimeIndex, dict]] = {}
    raw: Dict[str, pd.DataFrame] = {}
    universe_rows: List[dict] = []
    for t in scfg.SECTORS:
        df = sector_px.get(t)
        if df is None or len(df) == 0:
            universe_rows.append({"티커": t, "섹터명": SECTOR_NAME_KR.get(t, ""), "상태": "수집 실패(제외)"})
            continue
        price_i, idx_i, info = sector_price_frame(df, cal, scfg)
        if len(idx_i) < 252:
            universe_rows.append({"티커": t, "섹터명": SECTOR_NAME_KR.get(t, ""), "상태": f"이력 부족({len(idx_i)}일, 제외)", **info})
            continue
        frames[t] = (price_i, idx_i, info)
        raw[t] = df
        universe_rows.append({"티커": t, "섹터명": SECTOR_NAME_KR.get(t, ""), "상태": "정상", **info})
        if info["AdjClose지연"]:
            log("DATA", kv(ticker=t, event="adj_close_stale", gap_days=info["지연거래일수"],
                           adj_last=info["AdjClose마지막일"], close_last=info["Close마지막일"],
                           action="지연 구간은 Close 수익률로 이어붙임(배당 미포함 근사)"), M=M, level="warning")
    universe = pd.DataFrame(universe_rows)
    stage_timing["01_섹터데이터"] = round(time.time() - t0, 2)
    log("DATA", kv(event="universe_ready", ok=len(frames), excluded=len(scfg.SECTORS) - len(frames),
                   spy_score_pct_first=(str(spy_series["SCORE_PCT"].dropna().index[0].date())
                                        if spy_series["SCORE_PCT"].notna().any() else "-"),
                   rf=rf_daily is not None), M=M)

    # [v0.3.0 §1.E-2] 섹터 폭(SECTOR_BREADTH_200) — 11섹터 공통 후보라 여기서 1회만 계산해 ctx로
    # 전달한다(섹터별로 재계산하지 않음). 각 섹터의 TREND_200(그날 자기 200일선 이격도, 인과)이
    # >0인지만 보므로 다른 섹터의 미래 정보가 섞일 여지가 없다.
    sector_breadth_200 = None
    if scfg.USE_SECTOR_BREADTH and frames:
        trend200_by_t = {t: sector_technical_values(frames[t][0]["Adj Close"].astype(float))["TREND_200"]
                         for t in frames}
        breadth_df = pd.DataFrame({t: (v > 0) for t, v in trend200_by_t.items()}).reindex(cal)
        sector_breadth_200 = breadth_df.mean(axis=1, skipna=True)
        log("DATA", kv(event="sector_breadth_built", sectors=len(trend200_by_t),
                       last=round(float(sector_breadth_200.dropna().iloc[-1]), 3)
                       if sector_breadth_200.notna().any() else None), M=M)

    # ---- 2) 섹터별 파이프라인 ----
    t0 = time.time()
    ctx = {"M": M, "res": res, "scfg": scfg, "frames": frames, "raw": raw, "spy_tr": spy_tr, "spy_raw": spy_raw,
           "spy_raw_df": spy_df, "spy_series": spy_series, "rf_daily": rf_daily,
           "sector_breadth_200": sector_breadth_200}
    results, failed = run_sectors(list(frames.keys()), ctx, scfg, M)
    stage_timing["02_섹터파이프라인"] = round(time.time() - t0, 2)

    # ---- 3) 통합 ----
    t0 = time.time()
    sig_start = pd.Timestamp(M_cfg.SIGNAL_START)
    eval_idx = cal[cal >= sig_start]
    # [v0.3.0 §1.A] SPY의 다음 거래일 예측 — M.build_next_day_prediction()을 그대로 재사용(재계산 없음).
    # 구버전 M(v1.24.0 미만) 번들에도 안전하게 동작하도록 가드.
    nd_spy = None
    try:
        if hasattr(M, "build_next_day_prediction"):
            nd_spy = M.build_next_day_prediction(res, M_cfg)
    except Exception as e:
        log("RUN", kv(event="next_day_spy_failed", err=type(e).__name__, msg=str(e)[:160]), M=M, level="warning")
    matrix = build_prediction_matrix(results, res, eval_idx, scfg, nd_spy=nd_spy)
    portfolio_perf, portfolio_curve = build_portfolio_reference(results, res, eval_idx, M)
    summary = build_sector_summary(results, failed, universe)
    stage_timing["03_통합시트"] = round(time.time() - t0, 2)
    stage_timing["04_run()합계"] = round(time.time() - t_all, 2)
    log("DONE", kv(event="sector_pipeline_complete", ok=len(results), failed=len(failed),
                   elapsed_s=stage_timing["04_run()합계"]), M=M)
    return {"sectors": results, "failed": failed, "selftest": st, "universe": universe, "quality": pd.DataFrame(quality),
            "matrix": matrix, "portfolio_perf": portfolio_perf, "portfolio_curve": portfolio_curve,
            "summary": summary, "stage_timing": stage_timing, "scfg": scfg, "M_cfg": M_cfg,
            "signal_start": str(sig_start.date()), "cal_end": str(cal[-1].date()), "aborted": False,
            "m_bundle_meta": res.get("bundle_meta", {}), "nd_spy": nd_spy}


# =============================================================================
# [10] 통합 시트 — 01Z 섹터일별예측 매트릭스 / 12 섹터요약 / 13 섹터분산전략(참고)
# =============================================================================
def build_prediction_matrix(results: Dict[str, Dict[str, Any]], res: dict, eval_idx: pd.DatetimeIndex,
                            scfg: SectorConfig, nd_spy: Optional[dict] = None) -> pd.DataFrame:
    """날짜 × 11섹터: 예측(상승/중립/하락/추세보유/추세현금/신호없음)·목표비중, SPY 시장상황, 상승·하락 섹터 수.
    섹터 상장 전 날짜는 공란. 매 행이 그날 종가 기준 확정(다음날 시가 체결)이라는 M의 규칙 그대로.
    [v0.3.0 §1.A] nd_spy가 주어지면(run()에서 M.build_next_day_prediction()으로 재계산 없이 만든 SPY 다음날
    예측) 맨 끝에 '예측' 행 1개를 추가한다 — '구분' 열로 실적/예측을 구분(불변식·성과 계산은 '실적'만 사용,
    이 매트릭스는 참고용 시트라 별도 재계산 로직이 없어 여기서만 구분하면 됨)."""
    out = pd.DataFrame(index=eval_idx)
    out["날짜"] = [d.date() for d in eval_idx]
    spy_state = res["sig"]["state"].reindex(eval_idx)
    out["SPY 시장상황"] = spy_state.map(STATE_SHORT).fillna("-")
    out["SPY 목표비중"] = res["sig"]["target_pos"].reindex(eval_idx).round(2)
    up = pd.Series(0, index=eval_idx)
    down = pd.Series(0, index=eval_idx)
    for t in scfg.SECTORS:
        sr = results.get(t)
        if sr is None:
            out[f"{t} 예측"] = "실패/제외"
            out[f"{t} 목표비중"] = np.nan
            continue
        st = sr["state"].reindex(eval_idx)
        out[f"{t} 예측"] = st.map(STATE_SHORT).fillna("")
        out[f"{t} 목표비중"] = sr["target_pos"].reindex(eval_idx).round(2)
        up = up + (st == "RISK_ON").astype(int)
        down = down + (st == "RISK_OFF").astype(int)
    out.insert(1, "구분", "실적")
    out.insert(4, "상승예측 섹터수", up.values)
    out.insert(5, "하락예측 섹터수", down.values)
    out = out.reset_index(drop=True)

    if nd_spy is not None:
        row = {c: ("" if out[c].dtype == object else np.nan) for c in out.columns}
        row["날짜"] = nd_spy["다음거래일"].date()
        row["구분"] = "예측"
        row["SPY 시장상황"] = STATE_SHORT.get(nd_spy["확정국면_원시"], nd_spy["확정국면_원시"])
        row["SPY 목표비중"] = nd_spy["목표비중"]
        n_up = n_down = 0
        for t in scfg.SECTORS:
            sr = results.get(t)
            nd_t = sr["sheets"].get("next_day") if sr is not None else None
            if nd_t is None:
                row[f"{t} 예측"] = "실패/제외" if sr is None else ""
                row[f"{t} 목표비중"] = np.nan
                continue
            st_raw = nd_t["확정국면_원시"]
            row[f"{t} 예측"] = STATE_SHORT.get(st_raw, st_raw)
            row[f"{t} 목표비중"] = nd_t["목표비중"]
            n_up += int(st_raw == "RISK_ON")
            n_down += int(st_raw == "RISK_OFF")
        row["상승예측 섹터수"] = n_up
        row["하락예측 섹터수"] = n_down
        row_df = pd.DataFrame([row])[list(out.columns)]
        out = pd.concat([out, row_df], ignore_index=True)
    return out


def build_portfolio_reference(results: Dict[str, Dict[str, Any]], res: dict, eval_idx: pd.DatetimeIndex, M
                              ) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """[참고용 — 매매 권고 아님] 같은 평가창(SIGNAL_START~)에서
    (a) 11섹터 균등분산 전략(각 섹터 자기 신호, 그날 백테스트가 있는 섹터끼리 균등)  (b) 11섹터 균등 B&H
    (c) SPY 국면전략(M)  (d) SPY B&H — 성과표와 자산곡선. v0.1 결함 (b)의 교정: 네 곡선 모두 같은 날짜 범위."""
    if not results:
        return pd.DataFrame(), pd.DataFrame()
    strat = pd.DataFrame({t: sr["strategy_ret"] for t, sr in results.items()}).reindex(eval_idx)
    bh = pd.DataFrame({t: sr["bh_ret"] for t, sr in results.items()}).reindex(eval_idx)
    ew_strat = strat.mean(axis=1, skipna=True).fillna(0.0)
    ew_bh = bh.mean(axis=1, skipna=True).fillna(0.0)
    spy_bt = res["bt"].reindex(eval_idx)
    spy_strat = spy_bt["strategy_ret"].fillna(0.0)
    spy_bh = spy_bt["bh_ret"].fillna(0.0)
    perf = pd.DataFrame([M.perf_metrics(ew_strat, "11섹터 균등분산 전략(참고)"),
                         M.perf_metrics(ew_bh, "11섹터 균등 단순보유(참고)"),
                         M.perf_metrics(spy_strat, "SPY 국면전략(M)"),
                         M.perf_metrics(spy_bh, "SPY 단순보유")])
    perf.insert(1, "평가창", f"{eval_idx[0].date()}~{eval_idx[-1].date()}")
    curve = pd.DataFrame(index=eval_idx)
    curve["날짜"] = [d.date() for d in eval_idx]
    curve["11섹터균등분산전략"] = (1 + ew_strat).cumprod().round(4)
    curve["11섹터균등B&H"] = (1 + ew_bh).cumprod().round(4)
    curve["SPY국면전략"] = (1 + spy_strat).cumprod().round(4)
    curve["SPY B&H"] = (1 + spy_bh).cumprod().round(4)
    curve["신호있는섹터수"] = strat.notna().sum(axis=1).values
    return perf, curve.reset_index(drop=True)


def _rule_contrib_value(rule_contrib: pd.DataFrame, label: str, col: str) -> Any:
    """[v0.3.0 §1.G-1] 09b_규칙별기여 표에서 (규칙 라벨, 열) 하나를 뽑는다 — 12_섹터요약 신규 3열이
    이 표를 유일한 소스로 삼도록(값 재계산 없이 그대로 인용)."""
    if rule_contrib is None or not len(rule_contrib) or "규칙" not in rule_contrib.columns:
        return np.nan
    hit = rule_contrib.loc[rule_contrib["규칙"] == label, col]
    return hit.iloc[0] if len(hit) else np.nan


def build_sector_summary(results: Dict[str, Dict[str, Any]], failed: Dict[str, str], universe: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for t in SECTORS:
        base = {"티커": t, "섹터명": SECTOR_NAME_KR.get(t, "")}
        if t in failed:
            rows.append({**base, "상태": "실패", "비고": str(failed[t]).splitlines()[0][:160]})
            continue
        sr = results.get(t)
        if sr is None:
            urow = universe.loc[universe["티커"] == t] if len(universe) else pd.DataFrame()
            rows.append({**base, "상태": (urow["상태"].iloc[0] if len(urow) else "제외")})
            continue
        sh = sr["sheets"]
        perf = sh["perf"].set_index("전략") if len(sh["perf"]) else pd.DataFrame()
        p_s = perf.iloc[0] if len(perf) > 0 else {}
        p_b = perf.iloc[1] if len(perf) > 1 else {}
        p_m = perf.iloc[2] if len(perf) > 2 else {}
        dist = sr["state"].value_counts().to_dict()
        aud = sh["audit"]
        if len(aud) and "일치" in aud.columns:
            chk = aud[~aud["일치"].astype(str).str.startswith("N/A")]
            audit_ok = "전체 통과" if (len(chk) == 0 or chk["일치"].astype(str).str.startswith("OK").all()) else "불일치 발생"
        else:
            audit_ok = "미실행"
        rc = sh.get("rule_contrib")
        rows.append({**base, "상태": "정상", "실제데이터시작": sr["info"]["실제데이터시작"],
                     "신호시작일(복합점수)": sr["first_signal"] or "-",
                     "후보지표수": sr["n_candidates"], "엄격채택(전체표본)": sh["n_strict_pass"],
                     "최신재추정 채택수": sh["latest_adopted"], "기간중 채택된 지표수": len(sr["adopted"]),
                     "상승일수": dist.get("RISK_ON", 0), "중립일수": dist.get("NEUTRAL", 0), "하락일수": dist.get("RISK_OFF", 0),
                     "추세필터일수": dist.get("TREND_ONLY_IN", 0) + dist.get("TREND_ONLY_OUT", 0),
                     "신호없음일수": dist.get("NO_SIGNAL", 0),
                     "전략CAGR": p_s.get("CAGR"), "전략샤프": p_s.get("샤프"), "전략MDD": p_s.get("최대낙폭(MDD)"),
                     "B&H CAGR": p_b.get("CAGR"), "B&H샤프": p_b.get("샤프"), "B&H MDD": p_b.get("최대낙폭(MDD)"),
                     "200일선벤치CAGR": p_m.get("CAGR"),
                     "국면정의검증": ("PASS" if sh["regime_validity"] else ("FAIL" if sh["regime_validity"] is False else "판정불가")),
                     "룩어헤드감사": audit_ok, "캐시사용": sr["cache_hit"],
                     "실행시간(초)": sr["timing"].get("12_run_sector()합계"),
                     "AdjClose지연": sr["info"]["AdjClose지연"],
                     # [v0.3.0 §1.G-1] 09b_규칙별기여에서 그대로 인용(단일 소스, 재계산 없음).
                     "H진입일 익일평균수익(%)": _rule_contrib_value(rc, "위험회피진입①(H)", "발동일 익일B&H평균수익(%)"),
                     "상승 미탑승(%p)": _rule_contrib_value(
                         rc, "[전체] 상승일 미탑승(비중=0 & B&H>0, %p)", "비중=0인 날 B&H수익 합(%p)"),
                     "하락 회피(%p)": _rule_contrib_value(
                         rc, "[전체] 하락일 회피(비중=0 & B&H<0, %p — 음수=회피한 손실)", "비중=0인 날 B&H수익 합(%p)")})
    return pd.DataFrame(rows)


# =============================================================================
# [11] 리포트 — M.write_excel과 같은 서식(제목만 섹터용). 시트 구성은 M과 동일 + 통합 시트.
# =============================================================================
def _concat(results: Dict[str, Dict[str, Any]], key: str) -> pd.DataFrame:
    parts = [sr["sheets"][key] for t, sr in results.items()
             if isinstance(sr["sheets"].get(key), pd.DataFrame) and len(sr["sheets"][key])]
    return pd.concat(parts, ignore_index=True, sort=False) if parts else pd.DataFrame()


def build_sector_report(sres: Dict[str, Any], M=None, path: Optional[str] = None) -> str:
    t0 = time.time()
    scfg: SectorConfig = sres.get("scfg", CFG)
    path = path or scfg.OUT_XLSX
    if sres.get("aborted"):
        st = sres.get("selftest", {})
        meta = [("버전", f"{VERSION} ({VERSION_DATE})"),
                ("판별력 자기검사", f"FAIL (참신호 {st.get('n_true_pass')}/3, 잡음오채택 {st.get('n_noise_pass')}/5) — 실데이터 실행 중단"),
                ("면책", "본 산출물은 연구·교육 목적의 백테스트 결과이며 투자 자문이 아닙니다.")]
        write_sector_excel(path, {}, meta, M=M)
        return path
    results = sres["sectors"]
    M_cfg = sres["M_cfg"]
    st = sres["selftest"]
    summary = sres["summary"]
    ok_t = [t for t in scfg.SECTORS if t in results]

    sheets: Dict[str, pd.DataFrame] = {}
    sheets["01Z_섹터일별예측"] = sres["matrix"]
    sheets["12_섹터요약"] = summary
    sheets["13_섹터분산전략(참고)"] = sres["portfolio_perf"]
    sheets["13b_분산전략자산곡선"] = sres["portfolio_curve"]
    for t in ok_t:
        sheets[f"01_일별_{t}"] = results[t]["sheets"]["daily"]
    sheets["02_거래내역"] = _concat(results, "trades")
    sheets["03_지표검증"] = _concat(results, "val_sheet")
    sheets["04_채택근거상세"] = _concat(results, "detail")
    sheets["05_이벤트스터디"] = _concat(results, "events")
    sheets["05b_하락상승구간"] = _concat(results, "episodes")
    sheets["06_성과요약"] = _concat(results, "perf")
    sheets["06b_운용통계"] = _concat(results, "extra")
    sheets["06c_임계값민감도"] = _concat(results, "sens")
    sheets["07_연도별성과"] = _concat(results, "annual")
    sheets["08_워크포워드가중치"] = _concat(results, "wf")
    sheets["09_국면통계"] = _concat(results, "regime_stats")
    sheets["09b_규칙별기여"] = _concat(results, "rule_contrib")  # [v0.3.0 §1.G-1]
    q = sres.get("quality", pd.DataFrame())
    u = sres.get("universe", pd.DataFrame())
    sheets["10_데이터품질"] = pd.concat([u, q], ignore_index=True, sort=False) if len(q) else u
    sheets["11_룩어헤드감사"] = _concat(results, "audit")

    # ---- 00 실행요약 ----
    n_ok = len(ok_t)
    audit_all = sheets["11_룩어헤드감사"]
    if len(audit_all) and "일치" in audit_all.columns:
        chk = audit_all[~audit_all["일치"].astype(str).str.startswith("N/A")]
        audit_line = ("전체 통과" if (len(chk) == 0 or chk["일치"].astype(str).str.startswith("OK").all())
                      else "불일치 발생 - 확인 필요") + f" ({len(chk)}건 검사)"
    else:
        audit_line = "미실행"
    regime_pass = int(sum(1 for t in ok_t if results[t]["sheets"]["regime_validity"]))
    # [v0.3.0 §1.A] '구분'==실적만 사용 — build_prediction_matrix()가 맨 끝에 붙이는 '예측' 행이
    # 여기 섞이면 "최근 예측"이 확정되지 않은 다음 거래일 값으로 잘못 표시된다.
    mtx = sres["matrix"]
    mtx_actual = mtx[mtx["구분"] == "실적"] if "구분" in mtx.columns else mtx
    up_line = ""
    if len(mtx_actual):
        last = mtx_actual.iloc[-1]
        preds = [f"{t}:{last.get(f'{t} 예측', '')}" for t in scfg.SECTORS]
        up_line = f"{last['날짜']} 기준 — SPY:{last.get('SPY 시장상황', '')} | " + ", ".join(preds)
    pp = sres["portfolio_perf"].set_index("전략") if len(sres["portfolio_perf"]) else pd.DataFrame()

    def _pf(name: str) -> str:
        if name in pp.index:
            r = pp.loc[name]
            return f"CAGR {r.get('CAGR')} / 샤프 {r.get('샤프')} / MDD {r.get('최대낙폭(MDD)')}"
        return "-"

    n_cand = results[ok_t[0]]["n_candidates"] if ok_t else "-"

    # [v0.3.0 §1.A] '다음 거래일 예측' meta 블록 — M의 build_report() 패턴 그대로(새 계산 없음, t일
    # 종가로 이미 확정된 target_pos를 표시만 재구성). SPY 1줄 + 섹터 11줄 + 안내 1줄.
    nd_spy = sres.get("nd_spy")
    nd_rows: List[Tuple[str, str]] = []
    if nd_spy is not None:
        nd_rows.append(("다음 거래일 예측 - 기준일(데이터)", f"{nd_spy['기준일'].date()}{nd_spy['기준일_경과주의']}"))
        nd_rows.append(("다음 거래일 예측 - 대상일", f"{nd_spy['다음거래일'].date()} (다음 영업일 기준 — 미국 공휴일 "
                                                 "미반영, 실제 휴장이면 그 다음 개장일에 체결)"))
        nd_rows.append(("다음 거래일 예측 - SPY", f"{nd_spy['확정국면']} / 목표비중 {nd_spy['목표비중']:.2f} / "
                                               f"{nd_spy['예상행동_kr']}"))
        for t in scfg.SECTORS:
            nd_t = results.get(t, {}).get("sheets", {}).get("next_day") if t in results else None
            if nd_t is None:
                nd_rows.append((f"다음 거래일 예측 - {t}", "실패/제외" if t not in results else "판정불가(재구성 실패)"))
                continue
            nd_rows.append((f"다음 거래일 예측 - {t}", f"{nd_t['확정국면']} / 목표비중 {nd_t['목표비중']:.2f} / "
                                                     f"{nd_t['예상행동_kr']}"))
        nd_rows.append(("다음 거래일 예측 - 안내", "t일 종가로 확정된 target_pos를 t+1일 시가에 체결하는 기존 체결 "
                                               "규칙을 표시만 재구성한 것 — 새 계산이 아니며 06/07 등 백테스트 성과 "
                                               "시트에는 영향 없음. 01Z_섹터일별예측 마지막 행(구분=예측)·01_일별_티커 "
                                               "마지막 행에도 같은 값이 있음"))
    else:
        nd_rows.append(("다음 거래일 예측", "미제공(M 번들이 v1.24.0 미만이거나 계산 실패 — '최근 예측'/09b_규칙별기여 참고)"))

    meta = [
        ("버전", f"sector_rotation.py {VERSION} ({VERSION_DATE}) — market_regime_trader.py 번들 "
                f"{sres.get('m_bundle_meta', {}).get('bundle_version', '직접 res')}"),
        ("예측 대상", "11개 SPDR 섹터 ETF 각각의 절대 상승/하락 국면(SPY와 동일 파이프라인을 섹터 가격에 적용)"),
        ("신호/백테스트 기간", f"{sres['signal_start']} ~ {sres['cal_end']} (M의 SIGNAL_START와 동일 — 모든 성과 비교는 같은 창)"),
        ("체결 규칙", "t일 종가에 신호 확정 → t+1일 시가 체결 (룩어헤드 구조적 차단, M과 동일)"),
        ("거래비용", f"편도 {M_cfg.COST_BPS:.0f}bp (M과 동일)"),
        ("판별력 자기검사", (f"PASS (참신호 {st['n_true_pass']}/3, 잡음오채택 {st['n_noise_pass']}/5)" if st.get("passed")
                        else "미실행(RUN_SELFTEST=False)")),
        ("섹터 실행 결과", f"{n_ok}/{len(scfg.SECTORS)} 정상" + (f", 실패: {', '.join(sres['failed'].keys())}" if sres["failed"] else "")),
        ("최근 예측", up_line),
        ("후보지표", f"섹터당 {n_cand}개 = M 후보 전부(변동성/신용/매크로/크로스에셋/추세/자동생성) "
                    f"+ 섹터 기술 8 + SPY대비 상대강도 10(§1.E REL_MA200_SLOPE 포함) + 섹터 매크로 8(§1.D 확장) "
                    f"+ SPY 계층 6(마스킹 전 점수·위험 백분위, 베타 상호작용) + 잔차모멘텀 1 + 섹터폭 1(§1.E)"),
        ("검증·가중·신호 규칙", "M과 동일: 6기준 검증(IC/NW-t/5분위/하락AUC/4구간 부호안정성/커버리지, 시간감쇠), 워크포워드 "
                           f"재추정({scfg.SECTOR_REWEIGHT_FREQ or M_cfg.REWEIGHT_FREQ}), 보조채택, 추세트랙캡 {M_cfg.TREND_TRACK_WEIGHT_CAP:.0%}, "
                           f"기저시리즈캡 {M_cfg.SERIES_WEIGHT_CAP:.0%}, 위험(H)트랙, 규칙 ⓪ 급락트리거(M의 VIX_TERM 공유) ~ ⑩ 과열헤어컷, "
                           f"이력현상 {M_cfg.HYSTERESIS_DAYS}일/최소보유 {M_cfg.MIN_HOLD_DAYS}일, 국면 임계 {M_cfg.PCT_RISK_OFF:.0%}/{M_cfg.PCT_RISK_ON:.0%}"),
        ("섹터 학습 최소연수", f"{scfg.SECTOR_TRAIN_MIN_YEARS}년 (M은 {M_cfg.TRAIN_MIN_YEARS}년 — XLRE/XLC의 신호 시작을 당기기 위한 "
                          f"명시적 파라미터, 12_섹터요약 '신호시작일' 참조)"),
        ("국면정의 검증(수용기준)", f"{regime_pass}/{n_ok} 섹터 PASS (상승국면 익일평균수익 > 하락국면) — 09_국면통계"),
        ("룩어헤드 감사", f"섹터별 무작위 {scfg.AUDIT_SAMPLE}개 날짜 절단 재계산: {audit_line} — 11_룩어헤드감사"),
        ("11섹터 균등분산 전략(참고)", _pf("11섹터 균등분산 전략(참고)")),
        ("11섹터 균등 단순보유(참고)", _pf("11섹터 균등 단순보유(참고)")),
        ("SPY 국면전략(M)", _pf("SPY 국면전략(M)")),
        ("SPY 단순보유", _pf("SPY 단순보유")),
        ("난수 시드", str(scfg.RANDOM_SEED)),
        ("시트 안내", "01Z 섹터일별예측(날짜×11섹터 상승/중립/하락·목표비중, SPY 시장상황 병기, '구분' 열=실적/예측 — 맨 끝 1행이 "
                   "다음 거래일 예측) / 12 섹터요약(섹터별 1행, H진입일 익일평균수익·상승 미탑승·하락 회피 포함) / "
                   "13 섹터분산전략(참고 성과)·13b 자산곡선 / 01_일별_티커(섹터별 M 01시트와 동일 컬럼 + 위험점수백분위(H,섹터자체) 진단열, "
                   "마지막 행이 다음 거래일 예측) / "
                   "02 거래내역 / 03 지표검증 / 04 채택근거상세 / 05 이벤트스터디 / 05b 하락상승구간 / 06 성과·06b 운용통계·06c 임계값민감도 / "
                   "07 연도별(SPY평균비중 병기) / 08 워크포워드가중치 / 09 국면통계 / 09b 규칙별기여(규칙 발동일수·익일평균수익, 상승 미탑승/하락 "
                   "회피) / 10 데이터품질(섹터 유니버스·Adj Close 지연) / 11 룩어헤드감사 — 02~11(09b 포함)은 전부 '티커' 열로 구분"),
        ("면책", "본 산출물은 연구·교육 목적의 백테스트 결과이며 투자 자문이 아닙니다. 과거 성과는 미래 수익을 보장하지 않습니다."),
    ]
    meta = [meta[0]] + nd_rows + meta[1:]  # [v0.3.0 §1.A] 버전 다음에 '다음 거래일 예측' 블록 삽입(M과 동일 패턴)
    for k, v in sorted(sres.get("stage_timing", {}).items()):
        meta.append((f"실행시간 - {k}", f"{v:.1f}초"))
    for t in ok_t:
        tm = results[t]["timing"]
        meta.append((f"실행시간 - 섹터 {t}", f"{tm.get('12_run_sector()합계', 0):.1f}초 (검증+워크포워드 "
                     f"{tm.get('05_06_검증+워크포워드', 0):.1f}초{', 캐시' if results[t]['cache_hit'] else ''})"))

    write_sector_excel(path, sheets, meta, M=M)
    if scfg.EXPORT_DAILY_CSV:
        try:
            sres["matrix"].to_csv(scfg.DAILY_CSV_PATH, index=False, encoding="utf-8-sig")
            log("REPORT", kv(event="daily_csv_saved", file=scfg.DAILY_CSV_PATH, rows=len(sres["matrix"])), M=M)
        except Exception as e:
            log("REPORT", kv(event="daily_csv_failed", err=type(e).__name__), M=M, level="warning")
    log("REPORT", kv(event="report_ready", file=path, sheets=len(sheets) + 2, sectors=n_ok,
                     elapsed_s=round(time.time() - t0, 2)), M=M)
    return path


def write_sector_excel(path: str, sheets: Dict[str, pd.DataFrame], meta: List[Tuple[str, str]], M=None) -> None:
    """M.write_excel과 같은 서식·조건부서식·자산곡선 차트(13b 시트 기준). 00시트 제목만 섹터용."""
    t0 = time.time()
    with pd.ExcelWriter(path, engine="xlsxwriter", datetime_format="yyyy-mm-dd", date_format="yyyy-mm-dd") as xl:
        wb = xl.book
        f_title = wb.add_format({"bold": True, "font_size": 14, "font_color": "#1F3864"})
        f_head = wb.add_format({"bold": True, "bg_color": "#1F3864", "font_color": "white", "border": 1,
                                "align": "center", "valign": "vcenter", "text_wrap": True})
        f_wrap = wb.add_format({"text_wrap": True, "valign": "top", "border": 1})
        f_key = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1})
        f_val = wb.add_format({"border": 1, "text_wrap": True})
        f_pass = wb.add_format({"bg_color": "#C6EFCE", "font_color": "#006100", "bold": True})
        f_fail = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})
        f_up = wb.add_format({"bg_color": "#C6EFCE"})
        f_down = wb.add_format({"bg_color": "#FFC7CE"})

        ws = wb.add_worksheet("00_실행요약")
        xl.sheets["00_실행요약"] = ws
        ws.set_column(0, 0, 30); ws.set_column(1, 1, 110)
        ws.write(0, 0, "미국 11개 섹터 국면(상승/하락) 예측 & 섹터별 매매 시스템 — SPY 파이프라인 재적용", f_title)
        r = 2
        for k, v in meta:
            ws.write(r, 0, str(k), f_key); ws.write(r, 1, str(v), f_val); r += 1

        for name, df in sheets.items():
            if df is None or len(df) == 0:
                continue
            df.to_excel(xl, sheet_name=name, index=False)
            w = xl.sheets[name]
            w.freeze_panes(1, 1)
            cols = list(df.columns)
            for j, col in enumerate(cols):
                w.write(0, j, str(col), f_head)
                sample = df[col].map(str).head(300)
                width = min(max(12, int(sample.str.len().max() if len(sample) else 12) + 2), 60)
                if col in ("경제적근거", "판정사유", "근거요약", "원인", "진입 근거", "청산 근거", "선행메커니즘",
                           "왜 이 지표를 후보로 넣었나(경제적 근거)", "검증 결과 요약", "최종 판정 사유"):
                    w.set_column(j, j, 60, f_wrap)
                else:
                    w.set_column(j, j, width)
            w.autofilter(0, 0, len(df), len(cols) - 1)
            if "판정" in cols:
                cj = cols.index("판정")
                w.conditional_format(1, cj, len(df), cj, {"type": "cell", "criteria": "==", "value": '"PASS"', "format": f_pass})
                w.conditional_format(1, cj, len(df), cj, {"type": "cell", "criteria": "==", "value": '"FAIL"', "format": f_fail})
            if "복합점수" in cols:
                cj = cols.index("복합점수")
                w.conditional_format(1, cj, len(df), cj, {"type": "3_color_scale", "min_color": "#F8696B",
                                                          "mid_color": "#FFEB84", "max_color": "#63BE7B"})
            if "목표비중" in cols:
                cj = cols.index("목표비중")
                w.conditional_format(1, cj, len(df), cj, {"type": "data_bar", "bar_color": "#638EC6"})
            if name.startswith("01Z") or name.startswith("01_일별_"):
                for j, col in enumerate(cols):
                    if col.endswith(" 예측") or col in ("SPY 시장상황", "예측"):
                        w.conditional_format(1, j, len(df), j, {"type": "cell", "criteria": "==", "value": '"상승"', "format": f_up})
                        w.conditional_format(1, j, len(df), j, {"type": "cell", "criteria": "==", "value": '"하락"', "format": f_down})

        if "13b_분산전략자산곡선" in sheets and len(sheets["13b_분산전략자산곡선"]) > 0:
            d = sheets["13b_분산전략자산곡선"]
            cols = list(d.columns)
            ch = wb.add_chart({"type": "line"})
            n = len(d)
            for cname, color in [("11섹터균등분산전략", "#1F3864"), ("11섹터균등B&H", "#7F7F7F"),
                                 ("SPY국면전략", "#2E75B6"), ("SPY B&H", "#C00000")]:
                if cname in cols:
                    ci = cols.index(cname)
                    ch.add_series({"name": cname, "categories": ["13b_분산전략자산곡선", 1, 0, n, 0],
                                   "values": ["13b_분산전략자산곡선", 1, ci, n, ci], "line": {"color": color, "width": 1.25}})
            ch.set_title({"name": "11섹터 균등분산 전략 vs 균등 B&H vs SPY (누적 1.0 기준, 같은 평가창)"})
            ch.set_y_axis({"log_base": 10, "name": "누적수익(로그)"})
            ch.set_size({"width": 1100, "height": 460})
            cw = wb.add_worksheet("99_자산곡선")
            cw.insert_chart("B2", ch)
    log("REPORT", kv(event="excel_written", path=path, sheets=len(sheets) + 2, elapsed_s=round(time.time() - t0, 2)), M=M)


def main(res_or_path, M, scfg: Optional[SectorConfig] = None,
         sector_px_override: Optional[Dict[str, pd.DataFrame]] = None) -> str:
    """M.main()과 같은 역할: run → build_sector_report → Colab이면 자동 다운로드."""
    scfg = scfg or CFG
    sres = run(res_or_path, M, scfg, sector_px_override=sector_px_override)
    out = build_sector_report(sres, M=M)
    if hasattr(M, "maybe_colab_download"):
        M.maybe_colab_download(out)
        if scfg.EXPORT_DAILY_CSV and os.path.exists(scfg.DAILY_CSV_PATH) and not sres.get("aborted"):
            M.maybe_colab_download(scfg.DAILY_CSV_PATH)
    return out


# =============================================================================
# [12] 합성 섹터가격(오프라인 배관 검사용) + 모듈 단독 실행
# =============================================================================
def make_synthetic_sector_prices(res: dict, seed: int = 20260905) -> Dict[str, pd.DataFrame]:
    """M.make_synthetic_data()의 SPY 위에 섹터별 고유 잡음을 얹은 합성 OHLC. 상장일은 실제와 같게
    (1998-12-22 / XLRE 2015-10-08 / XLC 2018-06-19) 두어 편입·워밍업·신호시작 규칙이 실데이터와 같은
    경로를 타게 한다. 검증 레이어의 결과 해석 목적이 아니라 배관(파이프라인·리포트) 검사 전용."""
    rng = np.random.default_rng(seed)
    spy = res["px_dict"]["SPY"]
    spy = spy[~spy.index.duplicated(keep="last")].sort_index()
    spy_ret = np.log(spy["Close"].astype(float)).diff().fillna(0.0)
    starts = {t: "1998-12-22" for t in SECTORS}
    starts.update({"XLRE": "2015-10-08", "XLC": "2018-06-19"})
    out: Dict[str, pd.DataFrame] = {}
    for t in SECTORS:
        idx = spy.index[spy.index >= pd.Timestamp(starts[t])]
        beta = rng.uniform(0.7, 1.3)
        idio = rng.normal(0.0, 0.006, len(idx))
        r = beta * spy_ret.reindex(idx).values + idio
        close = 50.0 * np.exp(np.cumsum(r))
        open_ = close * (1.0 + rng.normal(0.0, 0.002, len(idx)))
        high = np.maximum(open_, close) * (1.0 + np.abs(rng.normal(0.0, 0.003, len(idx))))
        low = np.minimum(open_, close) * (1.0 - np.abs(rng.normal(0.0, 0.003, len(idx))))
        div_factor = np.exp(np.cumsum(np.full(len(idx), 0.00005)))    # 배당 재투자 근사(Adj > Close)
        out[t] = pd.DataFrame({"Open": open_, "High": high, "Low": low, "Close": close,
                               "Adj Close": close * div_factor / div_factor[-1], "Volume": 1_000_000.0}, index=idx)
    return out


if __name__ == "__main__":
    sys.path.insert(0, ".")
    import market_regime_trader as _M  # type: ignore
    print(f"sector_rotation.py {VERSION}: import 확인(M {getattr(_M, 'BUNDLE_VERSION', '?')}). "
          f"실행은 sector_rotation.main('market_regime_result.pkl.gz', M) 또는 main(res, M). SECTORS={SECTORS}")
