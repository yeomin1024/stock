# =============================================================================
#  industry_rotation.py
#  VERSION: v0.1.3 - 2026-09-11 - main()에 Colab 자동다운로드 추가(S.main()과 동일 패턴, 기존엔 누락)
#
#  목적:
#    market_regime_trader.py(M, v1.49.0)가 SPY 국면(E_t)을, sector_rotation.py(S, v0.37.0)가
#    S★(9개 섹터 배분)를 정했다. 이 파일(I)은 S★의 섹터 비중을 "그 섹터에 속한 산업 ETF들로
#    한 단계 더 나눠 담는다" — 총 노출은 절대 바꾸지 않는다(Σ산업 + 부모ETF = S★의 그 날 그
#    섹터 비중, 14_계층정합 시트가 매일 이 등식을 검사한다).
#
#    M -> S -> I. S가 "SPY 자리에 섹터를, SPY 계층 특징에 M의 점수/위험을" 넣었듯,
#    I는 "섹터 자리에 산업을, 섹터 계층 특징에 S의 점수/위험을" 넣는다. 새 통계 기법은 없다 —
#    S/M에서 검증된 함수를 그대로 호출하고, SPY가 하드코딩된 곳만 '부모 섹터'로 일반화한다.
#
#  설계 문서: claude/INDUSTRY_LAYER_SPEC_v0.1.md (프로젝트 "미국 주식 매수, 매도 프로그램
#    만들기2"). 본 파일은 그 문서의 §1~§11을 구현하되, §14 마일스톤 1~5(골격/단일산업파이프라인/
#    병렬통합/풀링순위/계층배분+격자+수용기준)에 해당하는 범위를 v0.1.0으로 먼저 낸다 — S 자신도
#    v0.1.0(sector_rotation_v0.1.0.py, ~2500행)에서 시작해 37개 버전에 걸쳐 지금의 형태(13d~13o
#    전체 진단시트, 06c 민감도, 11 룩어헤드감사 시트, 01Y 등)로 자랐다. 이번 v0.1.0도 같은 방식
#    으로 시작한다 — 핵심 파이프라인(수집->후보지표->검증->워크포워드->신호->백테스트->풀링순위
#    검증->계층배분->수용기준)은 전부 실동작하며 합성데이터 E2E로 검증됐지만, 아래는 v0.2 이후로
#    명시적으로 미룬다(조용히 생략하지 않고 여기 적는다 — 투명성 원칙):
#      (a) MAX_WORKERS 병렬 실행 — v0.1.0은 산업별 순차 실행만 지원(S의 fork 병렬 패턴은 복제하지
#          않음, 정확성 우선 — 직렬 19개 산업은 병렬 대비 느리지만 결과는 항상 재현 가능하다).
#      (b) 회복플로어/깊은낙폭/구조적바닥 승격 규칙(M 규칙⑧⑨⑬) — M.deep_drawdown_flag의 정확한
#          반환 개수·인자 순서를 이번 라운드에 재검증하지 못해 recov_conf/deep_recov/struct_dd를
#          항상 None으로 둔다(해당 규칙 비활성 = 안전한 축소, 잘못된 인자로 호출해 오답을 내는
#          것보다 낫다). score_pct/haz_pct/fast_pct 기반 규칙(①~⑦,⑩~⑫,⑭,⑮)은 전부 정상 동작.
#      (c) 06c_임계값민감도(M.threshold_sensitivity), 11_룩어헤드감사 시트, 13d~13o(13/13b/13c
#          제외) 시트군, 01Y_산업예측정확도, 13h_외부검증FF49(FF49 네트워크 요구), "다음 거래일
#          예측" 행(M.build_next_day_prediction은 SPY 형태의 res 번들을 요구해 산업 단위로 바로
#          재사용 불가) — 이번 버전은 대신 14_계층정합·15_산업대섹터귀속(설계서 신규 시트 2종)과
#          03_지표검증·10_데이터품질·01_일별_<IND>×19·12_산업요약·13/13b/13c·99_자산곡선을 낸다.
#          룩어헤드 안전성은 이 파일 대신 test_industry_lookahead_cut.py(절단재계산 회귀)가 독립
#          검증한다(사용자 표준 "룩어헤드 편향 방지" 요구 — 시트 유무와 무관하게 항상 우선).
#
#    market_regime_trader.py·sector_rotation.py는 이 파일에서 **한 줄도 수정하지 않는다**.
#
#  CHANGELOG
#  ---------------------------------------------------------------------------
#  v0.1.3 | 2026-09-11 | 사용자 지시: main()이 리포트만 만들고 Colab 자동다운로드를 하지 않던
#    것을 고침(M.main()/S.main()은 이미 하는데 I.main()만 빠져 있었음 — v0.1.0~v0.1.2 범위
#    축소 목록에 없던 순수 누락). S.main()(v0.9.2)과 동일한 패턴 재사용: 엑셀(OUT_XLSX)+일별
#    CSV(DAILY_CSV_PATH)+배분CSV(ALLOC_CSV_PATH, alloc 있을 때만) 중 실제로 존재하는 파일을
#    모아 M.maybe_colab_download_many()로 zip 1개에 묶어 한 번에 다운로드(브라우저의 "같은
#    세션 다중 자동다운로드 차단" 회피 — S가 v0.9.2에서 겪고 고친 문제와 동일). 그 함수가 없는
#    구버전 M과 조합 시 M.maybe_colab_download()를 파일별로 개별 호출하는 것으로 자동 대체.
#    Colab이 아니면(로컬/서버) M.maybe_colab_download 자체가 조용히 건너뜀 — 파이프라인 성공에
#    영향 없음. 영향 함수: main()(§8). import os 추가. 신호·배분·리포트 내용에는 무변경.
#  v0.1.2 | 2026-09-11 | 신규 함수 industry_lookahead_audit 추가(§2b, build_industry_candidates
#    바로 아래) — S.sector_lookahead_audit과 동일한 절단재계산 방법론을 산업층에 적용. 무작위
#    검사일 d마다 산업/부모/SPY 원시가격 + M/S의 이미 계산된 점수를 d까지 잘라 후보지표를 처음부터
#    다시 만들고 그날 가중치로 점수를 재계산해 전체계산과 비교한다. sres의 부모 score_pct/haz_pct는
#    S를 재실행하지 않고 d까지 추가로 자르는 방식으로 재사용(S 자신의 지표 인과성은 S 자체 감사가
#    담당). run()/run_industry()/build_industry_report() 등 기존 파이프라인에는 배선하지 않음(11_
#    룩어헤드감사 시트는 여전히 v0.2 예정) — test_industry_lookahead_cut.py 전용 호출부이며 기존
#    동작에는 아무 영향도 주지 않는 순수 추가. 영향 함수: industry_lookahead_audit(신규).
#  v0.1.1 | 2026-09-11 | 버그 수정: _leader_and_gate에서 그 부모 소속 산업이 전부 비활성(NaN)인
#    날(n_ok==0) masked.idxmax(axis=1)가 pandas "Encountered all NA values"로 예외 발생 ->
#    build_industry_allocation/run() 전체가 실패(합성 E2E 테스트 [4/9]에서 최초 발견).
#    수정: masked.fillna(-np.inf).idxmax(axis=1) 후 기존 계획대로 .where(n_ok>=1)로 널처리 —
#    leader 최종값 자체는 원래 의도(n_ok==0인 날은 리더 없음=NaN)와 동일, idxmax만 안전화.
#    영향 함수: _leader_and_gate(약 line 728). 다른 신호/배분 로직 변경 없음.
#  v0.1.0 | 2026-09-11 | 최초 구현. INDUSTRY_LAYER_SPEC_v0.1.md §1~§11 구현(위 범위 설명 참조).
#    §1 재사용: S.fetch_sector_prices/build_total_return_close/sector_price_frame/
#       sector_technical_specs·values/relative_strength_specs·values/spy_layer_series/
#       residual_momentum_values/sector_cfg_for/_sector_vol_scale/_vol_scaled_cfg/
#       validate_and_weight_sector/_indicator_spec_override/rotation_walkforward_select/
#       _cs_rank_ic/_nw_mean_tstat/portfolio_backtest/build_sector_sheets/write_sector_excel/
#       run_selftest를 그대로 호출(재계산 없음). M.composite_score/score_percentile/
#       generate_signals/run_backtest/event_study/extract_trades/drawdown_episodes/
#       build_reason_text/perf_metrics를 S와 동일한 순서로 호출.
#    §2 유니버스: 19개 산업 ETF(등급A) x 8개 부모(XLU 제외, 산업ETF 없음) — INDUSTRIES 표.
#       유니버스 무결성: 부모가 S.SECTOR_EXCLUDE에 있거나 sres["sectors"]에 없으면 그 산업은
#       "제외(부모 미실행)"로 12_산업요약에 적고 건너뛴다(run()의 활성 유니버스 계산부 참조).
#    §4 후보지표: (a) M 후보 321(재사용) (b) 산업 자체기술 8 (c) 부모대비상대 10(P_ 접두)
#       (d) SPY대비상대 10(S_ 접두) (e) SPY계층 6 (f) 부모계층 6(신규 parent_layer_*, PARENT_*)
#       (g) 매크로 8(부모 표 상속)+최대2(산업별 덮어쓰기, §4.5) (h) 잔차모멘텀 2(부모·SPY)
#       (i) 폭 2(산업폭·부모폭).
#    §6 산업 순환매: 풀링(활성 산업 전부) + 부모초과수익(비율 P_i/P_parent의 일간수익률을
#       ret_cc_full로 사용 — S의 상대강도 산식과 동일한 비율기반 정의, §6.2의 "산업수익-부모수익"
#       근사를 곱셈적으로 정확히 구현). S.rotation_walkforward_select를 SPY후보 없이(cand=산업만)
#       그대로 호출.
#    §7 배분: build_industry_allocation 신규 구현(부모비중 안 산업리더 집중·하락시 부모, ⚠
#       INDUSTRY_LEADER_CAP=0.5·INDUSTRY_FALLBACK_SHARE=0.5) + [산업집중격자]/[폴백격자] +
#       대조군A(=S★ 비트동일, 100% 부모ETF). 비용은 산업 10bp/부모 5bp를 회전율별로 분리
#       적용(portfolio_backtest를 cost_bps=0으로 호출한 뒤 자산군별 비용을 직접 차감 — §7.5).
#    §8 수용기준 ①~⑤ PASS/FAIL 표(13f_산업수용기준). §14_계층정합(위반일수·최대오차).
#       §15_산업대섹터귀속(I★ vs S★ 일별 초과수익).
#    ※ 본 코드는 연구/교육용 도구이며 투자 자문이 아니다.
# =============================================================================
from __future__ import annotations

import os
import time
import traceback
import dataclasses
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd

VERSION = "v0.1.3"
VERSION_DATE = "2026-09-11"

# =============================================================================
# [0] 산업 유니버스 — INDUSTRY_LAYER_SPEC_v0.1.md §2
#     (산업티커, 부모섹터, FF49매핑열) 3튜플. 등급A(기본 활성) 19개 · 8개 부모(XLU 제외).
# =============================================================================
INDUSTRIES: Tuple[Tuple[str, str, str], ...] = (
    ("SOXX", "XLK", "Chips"), ("IGV", "XLK", "Softw"),
    ("IBB", "XLV", "Drugs"), ("IHE", "XLV", "Drugs"), ("IHI", "XLV", "MedEq"), ("IHF", "XLV", "Hlth"),
    ("XRT", "XLY", "Rtail"), ("XHB", "XLY", "Cnstr"), ("PEJ", "XLY", "Fun"),
    ("PBJ", "XLP", "Food"),
    ("KBE", "XLF", "Banks"), ("KRE", "XLF", "Banks"), ("KIE", "XLF", "Insur"), ("KCE", "XLF", "Fin"),
    ("ITA", "XLI", "Aero"), ("IYT", "XLI", "Trans"),
    ("IYZ", "XLC", "Telcm"), ("FDN", "XLC", "Fun"),
    ("REZ", "XLRE", "RlEst"),
)

# ⚠ 등급B(이력 짧음/유동성 낮음) — 기본 꺼짐. IndustryConfig.USE_TIER_B=True로만 켠다(§15#7).
INDUSTRIES_OPTIONAL: Tuple[Tuple[str, str, str], ...] = (
    ("CARZ", "XLY", "Autos"), ("JETS", "XLI", "Trans"), ("SOCL", "XLC", ""),
    ("SRVR", "XLRE", "RlEst"), ("INDS", "XLRE", "RlEst"),
)

INDUSTRY_NAME_KR: Dict[str, str] = {
    "SOXX": "반도체", "IGV": "소프트웨어", "IBB": "바이오텍", "IHE": "제약", "IHI": "의료기기",
    "IHF": "헬스케어서비스·보험", "XRT": "소매", "XHB": "주택건설", "PEJ": "레저·엔터테인먼트",
    "PBJ": "식음료", "KBE": "은행", "KRE": "지역은행", "KIE": "보험", "KCE": "자본시장",
    "ITA": "항공우주·방산", "IYT": "운송", "IYZ": "통신", "FDN": "인터넷·인터랙티브미디어",
    "REZ": "주거·헬스케어리츠",
    "CARZ": "자동차", "JETS": "항공사", "SOCL": "소셜미디어", "SRVR": "데이터센터리츠", "INDS": "산업용리츠",
}

# ⚠ 상장연도는 근사치(설계서 §2). 코드는 실측 첫 유효일을 10_데이터품질에 적고,
#   INDUSTRY_TRAIN_MIN_YEARS 미만이면 신호 시작을 자동으로 늦춘다(S의 워밍업 규칙과 동일).
INDUSTRY_EXPECTED_START: Dict[str, str] = {
    "SOXX": "2001-08-01", "IGV": "2001-08-01", "IBB": "2001-03-01", "IHE": "2006-06-01",
    "IHI": "2006-06-01", "IHF": "2006-06-01", "XRT": "2006-06-01", "XHB": "2006-06-01",
    "PEJ": "2005-06-01", "PBJ": "2005-06-01", "KBE": "2005-11-01", "KRE": "2006-06-01",
    "KIE": "2005-11-01", "KCE": "2005-11-01", "ITA": "2006-05-01", "IYT": "2003-11-01",
    "IYZ": "2000-05-01", "FDN": "2006-06-01", "REZ": "2007-05-01",
    "CARZ": "2011-08-01", "JETS": "2015-04-01", "SOCL": "2011-11-01", "SRVR": "2018-06-01", "INDS": "2018-06-01",
}

PARENTS: Tuple[str, ...] = tuple(sorted({p for _, p, _ in INDUSTRIES}))  # 8개(XLU 제외)


def _ensure_yahoo_expected_start_industries(M) -> None:
    """[§2] M.YAHOO_EXPECTED_START(모듈 전역 dict)에 산업 ETF 티커를 런타임에 추가한다.
    .get(ticker) 조회만 쓰는 기존 경로는 새 키가 추가돼도 다른 티커에 영향받지 않는다 —
    market_regime_trader.py 소스는 한 글자도 바뀌지 않는다(S의 _ensure_yahoo_expected_start와
    동일 패턴). 부모 섹터 티커는 S.SECTOR_EXPECTED_START가 이미 다룬다(S.fetch_sector_prices가
    내부에서 S._ensure_yahoo_expected_start(M)을 호출)."""
    M.YAHOO_EXPECTED_START.update(INDUSTRY_EXPECTED_START)


# =============================================================================
# [1] 설정 — SectorConfig와 필드명을 의도적으로 동일하게 맞춘 필드가 많다(§ 재사용 지도).
#     S/M의 제네릭 함수(rotation_walkforward_select·sector_cfg_for·_sector_vol_scale·
#     validate_and_weight_sector 등)는 scfg.FIELD를 getattr로 읽으므로, 이 IndustryConfig를
#     그 함수들에 scfg 자리에 그대로 넘기면(덕타이핑) 재구현 없이 재사용된다.
# =============================================================================
@dataclass
class IndustryConfig:
    INDUSTRIES: Tuple[Tuple[str, str, str], ...] = INDUSTRIES
    USE_TIER_B: bool = False                    # ⚠ 등급B 산업(CARZ·JETS·SOCL·SRVR·INDS) 기본 꺼짐(§15#7)
    ADJ_CLOSE_STALE_DAYS: int = 5

    # ---- 후보지표 구성 스위치 ----
    USE_MARKET_CANDIDATES: bool = True
    USE_INDUSTRY_TECHNICAL: bool = True
    USE_PARENT_RELATIVE: bool = True
    USE_SPY_RELATIVE: bool = True
    USE_SPY_LAYER_FEATURES: bool = True
    SPY_LAYER_BETA_CONDITIONAL: bool = True
    USE_PARENT_LAYER: bool = True
    PARENT_LAYER_BETA_CONDITIONAL: bool = True
    USE_INDUSTRY_MACRO: bool = True
    USE_RESID_MOMENTUM: bool = True
    RESID_MOM_WINDOW: int = 756
    USE_INDUSTRY_BREADTH: bool = True

    # ---- 부모 계층 시리즈 소스(§4.3) ----
    PARENT_LAYER_SOURCE: str = "masked_extend"   # "masked_extend"(기본) | "recompute"(v0.2 예정, 현재 masked_extend로 폴백)

    # ---- 학습/재추정(S와 동일 이름 — sector_cfg_for가 이 이름으로 읽는다) ----
    SECTOR_TRAIN_MIN_YEARS: int = 3              # = INDUSTRY_TRAIN_MIN_YEARS
    SECTOR_REWEIGHT_FREQ: Optional[str] = None   # None = M_cfg.REWEIGHT_FREQ("M")
    # ⚠ [§12] "Q"로 바꾸면 재추정이 ~3배 빨라지지만 신호가 달라진다 — 기본은 M과 동일하게 유지.

    # ---- 위험(H) 트랙 소스 ----
    HAZARD_SOURCE: str = "spy"                   # "spy"(기본, M의 H) | "parent"(S의 H) | "max"

    # ---- 사이징 오버레이 변동성 정규화(부모 대비, S는 SPY 대비) ----
    VOL_SCALE_OVERLAYS: bool = True
    VOL_SCALE_MIN: float = 0.7
    VOL_SCALE_MAX: float = 2.0                   # ⚠ SOXX(연 35%대) 같은 고변동 산업은 상한에 닿을 수 있다 — 12_산업요약 vol_scale 열로 노출
    VOL_SCALE_WINDOW: int = 252

    # ---- 풀링 횡단면 순환매(§6) — S.rotation_walkforward_select가 scfg.ROTATION_*를 그대로 읽는다 ----
    USE_ROTATION: bool = True
    ROTATION_SIGNALS: Tuple[str, ...] = ("PARENT_SCORE_PCT", "P_REL_MOM_126", "P_REL_MOM_12_1",
                                         "P_REL_MOM_21", "P_REL_EXT_200", "RESID_MOM_12_1_PARENT",
                                         "PBETA_X_SCORE")
    ROTATION_SELECT_T: float = 2.0
    ROTATION_SELECT_MIN_DAYS: int = 1000
    ROTATION_SELECT_HORIZON: int = 21
    ROTATION_DECAY_WEIGHTED: bool = True
    ROTATION_DECAY_HALF_LIFE_DAYS: Optional[int] = None
    ROTATION_SELECT_STAT: str = "top1"
    ROTATION_SELECT_MODE: str = "best_available"
    ROTATION_BEST_N: int = 2
    ROTATION_SELECT_T_MIN: float = 1.0
    ROTATION_SMOOTH_DAYS: int = 21
    ROTATION_MIN_AGREE: int = 2
    ROTATION_LEADER_MARGIN_STEPS: float = 1.0
    ROTATION_EVIDENCE_TIER: bool = True
    ROTATION_DEDUP_COMPOSITE: bool = True
    ROTATION_AVOID_VALIDATE: bool = True
    ROTATION_INCLUDE_SPY_CANDIDATE: bool = False  # [§6.1] 풀링 횡단면 = 활성 산업 전부, SPY/부모는 후보 아님
    ROTATION_IC_HORIZON: int = 21
    USE_EXTERNAL_VALIDATION: bool = False          # v0.2 예정(§6.4) — FF49 네트워크 필요, 이번 버전은 꺼둠
    EXTERNAL_T: float = 2.0

    # ---- 산업 배분(§7) ----
    INDUSTRY_LEADER_CAP: float = 0.5              # ⚠ 리더 산업에 주는 섹터비중 몫(기본 "절반만 산업으로")
    INDUSTRY_FALLBACK_SHARE: float = 0.5           # ⚠ 폴백일 적격산업 균등배분 몫
    INDUSTRY_EXCLUDE_STATES: Tuple[str, ...] = ("RISK_OFF",)
    COST_BPS_INDUSTRY: float = 10.0                # ⚠ 산업 ETF 편도(S·M의 5bp보다 큼 — 스프레드 반영)
    PARENT_COST_BPS: float = 5.0                   # ⚠ 부모 ETF 다리 편도(S·M과 동일)

    # ---- 수용기준(§8, 사전 고정) ----
    ACCEPT_CAGR_GAIN: float = 0.01                 # ⚠ I★ CAGR ≥ S★ + 1.0%p
    ACCEPT_MDD_WORSE: float = 0.01                 # ⚠ MDD 악화 ≤ 1.0%p
    ACCEPT_TOP1_T: float = 2.0                     # 풀링 상위1 스프레드 NW-t ≥ 2.0
    ACCEPT_SPREAD_PERIODS_RATIO: float = 2.0 / 3.0 # 상위3-하위3 스프레드 양수 구간 ≥ 2/3

    # ---- 격자 채택 기준(S v1.46.0~과 동일, ①②③④) ----
    GRID_CAGR_LOSS_TOL: float = 0.005
    GRID_MDD_WORSE_TOL: float = 0.005

    # ---- 진단 ----
    RUN_SELFTEST: bool = True
    SELFTEST_MIN_TRUE_ADOPTED: int = 2
    SELFTEST_MAX_NOISE_ADOPTED: int = 1
    RUN_LOOKAHEAD_AUDIT: bool = True               # sector_cfg_for가 참조(M_cfg 전파용) — I 자체 11시트는 v0.2
    AUDIT_SAMPLE: int = 3

    # ---- 성능 ----
    USE_CACHE: bool = True
    CACHE_DIR: str = "./cache_industry"
    MAX_WORKERS: int = 1                           # ⚠ v0.1.0은 순차 실행만(§12, 정확성 우선 — 병렬은 v0.2 예정)

    # ---- 출력 ----
    OUT_XLSX: str = "industry_regime_report.xlsx"
    EXPORT_DAILY_CSV: bool = True
    DAILY_CSV_PATH: str = "industry_daily.csv"
    ALLOC_CSV_PATH: str = "industry_allocation_daily.csv"

    RANDOM_SEED: int = 20260905


CFG = IndustryConfig()


def log(stage: str, msg: str, level: str = "info", M=None) -> None:
    if M is not None:
        M.log(f"INDUSTRY_{stage}", msg, level)
    else:
        print(f"[{level.upper()}][INDUSTRY_{stage}] {msg}")


def kv(**kwargs) -> str:
    parts = []
    for k, v in kwargs.items():
        if isinstance(v, float):
            v = f"{v:.4f}" if abs(v) < 1e4 else f"{v:.2f}"
        parts.append(f"{k}={v}")
    return " ".join(parts)


def active_industries(icfg: IndustryConfig, sres: dict) -> List[Tuple[str, str, str]]:
    """[유니버스 무결성 규칙, §2] 부모가 S.SECTOR_EXCLUDE에 있거나 sres["sectors"]에 없으면
    그 산업은 제외한다. 부모 없는 산업은 존재할 수 없다(assert)."""
    table = icfg.INDUSTRIES + (INDUSTRIES_OPTIONAL if icfg.USE_TIER_B else ())
    for ind, parent, _ in table:
        assert parent, f"{ind}: 부모 섹터 없이 존재하는 산업"
    ok_parents = set(sres.get("sectors", {}).keys())
    return [(ind, parent, ff49) for (ind, parent, ff49) in table if parent in ok_parents]


# =============================================================================
# [2] 후보지표 원료 — _RawSpec (S와 동일 구조)
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
    return 63 if window >= 60 else None


def _mk_ispec(M, ind_ticker: str, raw: _RawSpec):
    return M.IndicatorSpec(
        key=f"{ind_ticker}__{raw.suffix}", name_kr=f"[{ind_ticker}] {raw.name_kr}", category=raw.category,
        prior_sign=raw.prior_sign, rationale=raw.rationale, source=raw.source,
        lead_mechanism=raw.lead_mechanism, trend_track=raw.trend_track,
        eval_horizon=raw.eval_horizon, base_series=raw.base_series)


# ---- 부모 계층 시리즈(§4.3) — masked_extend: S의 마스킹된 score_pct/haz_pct를 그대로 쓴다 ----
def parent_layer_series(sres: dict, parent: str, M, mode: str = "masked_extend") -> Dict[str, pd.Series]:
    """[§4.3] sres["sectors"][parent]는 S의 run_sector() 반환 dict — score_pct/haz_pct는
    이미 '그 섹터의 SIGNAL_START 이후'만 값이 있는 마스킹 시리즈다(S가 그렇게 만들었다 — S의
    run_sector 07~08단계: score_pct = M.score_percentile(score).where(sig_mask)). masked_extend
    모드는 이 마스킹 시리즈를 그대로 후보값으로 쓴다 — coverage 게이트가 그 산업의 평가창에서
    부모 신호가 충분히 있는지를 자연히 판정하므로(짧으면 FAIL로 자동 배제), 잘못 채택되기보다
    안전하게 배제되는 쪽으로 수렴한다. recompute 모드(캐시에서 마스킹 전 score 복원)는 v0.2
    예정 — 이번 버전은 항상 masked_extend로 동작(경고 로그)."""
    sr = sres["sectors"].get(parent, {})
    out: Dict[str, pd.Series] = {}
    out["SCORE_PCT"] = sr.get("score_pct", pd.Series(dtype=float))
    out["HAZ_PCT"] = sr.get("haz_pct", pd.Series(dtype=float))
    if mode != "masked_extend":
        log("DATA", kv(event="parent_layer_source_fallback", parent=parent, requested=mode,
                       used="masked_extend", note="recompute는 v0.2 예정"), M=M, level="warning")
    return out


def parent_layer_specs(beta_med: Optional[float] = None, beta_conditional: bool = False) -> List[_RawSpec]:
    """[§4.4] S.spy_layer_specs를 부모 버전으로 복제 — SCORE_PCT/HAZ_PCT 사전방향을 베타(부모
    대비) 중앙값이 1 미만이면 뒤집는다(S v0.13.0과 같은 CAPM 논리). PBETA_X_FT는 산업/부모
    자체의 급락트리거가 없으므로 SPY의 FAST_PCT(M 계층)를 그대로 재사용한다."""
    defensive = bool(beta_conditional and beta_med is not None and np.isfinite(beta_med) and beta_med < 1.0)
    sc_sign = -1 if defensive else +1
    hz_sign = +1 if defensive else -1
    return [
        _RawSpec("PARENT_SCORE_PCT", "부모섹터 복합점수 백분위(S 계층, 마스킹)", "F.부모계층", sc_sign,
                 "부모 섹터 강세 확신이 높을수록 산업도 상승 확률이 높다(산업은 부모의 베타를 공유)"
                 if not defensive else "베타(부모대비)<1인 산업은 부모 강세 국면에서 상대적으로 뒤진다",
                 "S 계층 확정 출력의 재사용", "sres[sectors][parent][score_pct]"),
        _RawSpec("PARENT_HAZ_PCT", "부모섹터 위험점수 백분위(S 계층, 마스킹)", "F.부모계층", hz_sign,
                 "부모 섹터 위험이 높을수록 산업 하락 확률이 높다"
                 if not defensive else "베타(부모대비)<1인 산업은 부모 위험 상승기에 상대적으로 유리",
                 "S 계층 확정 출력의 재사용", "sres[sectors][parent][haz_pct]"),
        _RawSpec("PBETA_X_SCORE", "(부모베타-1)×부모복합점수백분위", "C.국면상호작용", +1,
                 "부모 강세 확신이 높을수록 고베타(부모대비) 산업이 더 크게 상승", "강세장에서 고베타 우위",
                 "PARENT_SCORE_PCT × 롤링베타(부모대비)"),
        _RawSpec("PBETA_X_H", "(부모베타-1)×부모위험점수백분위", "C.국면상호작용", -1,
                 "부모 위험이 높을수록 고베타(부모대비) 산업이 더 크게 하락", "위험 프리미엄 확대 시 고베타 할인",
                 "PARENT_HAZ_PCT × 롤링베타(부모대비)"),
        _RawSpec("PBETA_X_DH", "(부모베타-1)×부모위험점수 15일변화", "C.국면상호작용", -1,
                 "부모 위험이 가속(ΔH 급등)하는 국면에서 고베타(부모대비) 산업이 더 취약",
                 "위험 '가속도'에 대한 고베타 민감도", "PARENT_HAZ_PCT.diff(15) × 롤링베타(부모대비)"),
        _RawSpec("PBETA_X_FT", "(부모베타-1)×SPY급락트리거백분위", "C.국면상호작용", -1,
                 "급성 변동성 급등 국면에서 고베타(부모대비) 산업이 즉각 열위",
                 "산업·부모 자체 급락트리거가 없어 M의 SPY FAST_PCT를 재사용",
                 "res[fast_pct] × 롤링베타(부모대비)"),
    ]


def parent_layer_values(parent_series: Dict[str, pd.Series], beta_vs_parent: pd.Series,
                        fast_pct_spy: Optional[pd.Series] = None) -> pd.DataFrame:
    idx = beta_vs_parent.index
    out = pd.DataFrame(index=idx)
    sp = parent_series["SCORE_PCT"].reindex(idx)
    hp = parent_series["HAZ_PCT"].reindex(idx)
    bx = beta_vs_parent - 1.0
    out["PARENT_SCORE_PCT"] = sp
    out["PARENT_HAZ_PCT"] = hp
    out["PBETA_X_SCORE"] = bx * sp
    out["PBETA_X_H"] = bx * hp
    out["PBETA_X_DH"] = bx * hp.diff(15)
    fp = fast_pct_spy.reindex(idx) if fast_pct_spy is not None else pd.Series(np.nan, index=idx)
    out["PBETA_X_FT"] = bx * fp
    return out.replace([np.inf, -np.inf], np.nan)


# ---- 산업별 매크로 덮어쓰기(§4.5) — 소수만, 경제적 근거 필수 ----
_IndMacroRow = Tuple[str, str, str, str, str, int, int, str, str]
INDUSTRY_MACRO_OVERRIDE: Dict[str, List[_IndMacroRow]] = {
    "XHB": [("MORT30_CHG20", "30년 모기지금리 20일 변화", "fred_rate", "MORTGAGE30US", "chg", 20, -1,
             "주택 수요 채널의 단기(20일) 버전 — 부모(XLY) 표의 60일 버전보다 반응이 빠르다", "금리-주택 채널")],
    "KBE": [("CURVE_CHG20", "수익률곡선 10년-2년 20일 변화", "fred_rate", "T10Y2Y", "chg", 20, +1,
             "순이자마진 채널의 단기 버전", "순이자마진 채널")],
    "KRE": [("CURVE_CHG20", "수익률곡선 10년-2년 20일 변화", "fred_rate", "T10Y2Y", "chg", 20, +1,
             "지역은행은 예대마진 민감도가 대형은행보다 커 단기 커브 변화에도 더 민감하게 반응", "순이자마진 채널")],
    "IYT": [("OIL_MOM20", "WTI 원유 20일 모멘텀", "yahoo_mom", "CL=F", "mom", 20, -1,
             "연료비는 운송업 영업비용의 핵심 변수 — 부모(XLI) 표에는 없는 채널", "연료비 채널")],
    "REZ": [("NOM10Y_CHG20", "10년 국채금리 20일 변화", "fred_rate", "DGS10", "chg", 20, -1,
             "리츠 할인율 채널의 단기 버전", "할인율 채널")],
}


def _industry_macro_values(rows: List[_IndMacroRow], res: dict, M, idx: pd.DatetimeIndex,
                           haz_pct_spy: pd.Series) -> pd.DataFrame:
    """S.sector_macro_values와 동일한 kind-스위치를 산업별 rows 목록에 적용한다(S 함수는
    SECTOR_MACRO_TABLE[ticker]를 내부에서 직접 읽어 파라미터화가 안 되므로 얇게 복제)."""
    out = pd.DataFrame(index=idx)
    for suffix, name_kr, kind, sid, transform, window, prior_sign, rationale, mech in rows:
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
                col = haz_pct_spy.reindex(idx)
        out[suffix] = col if col is not None else np.nan
    return out.replace([np.inf, -np.inf], np.nan)


def build_industry_candidates(ind_ticker: str, parent: str, res: dict, M, icfg: IndustryConfig, S,
                              adj_tr: pd.Series, close_raw: pd.Series,
                              spy_tr: pd.Series, spy_raw: pd.Series,
                              parent_tr: pd.Series, parent_raw: pd.Series,
                              spy_series: Dict[str, pd.Series], parent_series: Dict[str, pd.Series],
                              idx: pd.DatetimeIndex,
                              industry_breadth: Optional[pd.Series] = None,
                              parent_breadth: Optional[pd.Series] = None,
                              ) -> Tuple[pd.DataFrame, List[Any]]:
    """[§4] 한 산업의 후보지표 전체 = M후보(재사용) + 산업기술8 + 부모대비상대10(P_) +
    SPY대비상대10(S_) + 부모계층6 + 산업매크로(8~10) + 잔차모멘텀2(부모·SPY) + 폭2."""
    frames: List[pd.DataFrame] = []
    specs: List[Any] = []

    if icfg.USE_MARKET_CANDIDATES:
        frames.append(res["ind"].reindex(idx))
        specs.extend(list(M.INDICATOR_SPECS))

    sec = pd.DataFrame(index=idx)

    if icfg.USE_INDUSTRY_TECHNICAL:
        vals = S.sector_technical_values(adj_tr.reindex(idx))
        for raw in S.sector_technical_specs():
            sec[f"{ind_ticker}__{raw.suffix}"] = vals[raw.suffix]
            specs.append(_mk_ispec(M, ind_ticker, raw))

    if icfg.USE_PARENT_RELATIVE:
        vals = S.relative_strength_values(adj_tr.reindex(idx), parent_tr.reindex(idx),
                                          close_raw.reindex(idx), parent_raw.reindex(idx), M)
        for raw in S.relative_strength_specs():
            key = f"{ind_ticker}__P_{raw.suffix}"
            sec[key] = vals[raw.suffix]
            specs.append(M.IndicatorSpec(
                key=key, name_kr=f"[{ind_ticker}] 부모({parent})대비 {raw.name_kr}", category="H2.부모대비상대",
                prior_sign=raw.prior_sign, rationale=raw.rationale, source=raw.source,
                lead_mechanism=raw.lead_mechanism, trend_track=raw.trend_track,
                eval_horizon=raw.eval_horizon, base_series=""))

    if icfg.USE_SPY_RELATIVE:
        vals = S.relative_strength_values(adj_tr.reindex(idx), spy_tr.reindex(idx),
                                          close_raw.reindex(idx), spy_raw.reindex(idx), M)
        for raw in S.relative_strength_specs():
            key = f"{ind_ticker}__S_{raw.suffix}"
            sec[key] = vals[raw.suffix]
            specs.append(M.IndicatorSpec(
                key=key, name_kr=f"[{ind_ticker}] SPY대비 {raw.name_kr}", category="H.SPY대비상대",
                prior_sign=raw.prior_sign, rationale=raw.rationale, source=raw.source,
                lead_mechanism=raw.lead_mechanism, trend_track=raw.trend_track,
                eval_horizon=raw.eval_horizon, base_series=""))

    if icfg.USE_SPY_LAYER_FEATURES:
        r_i = np.log(adj_tr.reindex(idx).replace(0, np.nan)).diff()
        r_spy = np.log(spy_tr.reindex(idx).replace(0, np.nan)).diff()
        beta_spy = S.rolling_beta(r_i, r_spy, window=252, lag=1)
        _bmed = float(beta_spy.median()) if beta_spy.notna().any() else np.nan
        vals = S.spy_layer_values(spy_series, beta_spy)
        for raw in S.spy_layer_specs(beta_med=_bmed, beta_conditional=bool(icfg.SPY_LAYER_BETA_CONDITIONAL)):
            sec[f"{ind_ticker}__{raw.suffix}"] = vals[raw.suffix]
            specs.append(_mk_ispec(M, ind_ticker, raw))

    beta_parent = None
    if icfg.USE_PARENT_LAYER:
        r_i = np.log(adj_tr.reindex(idx).replace(0, np.nan)).diff()
        r_p = np.log(parent_tr.reindex(idx).replace(0, np.nan)).diff()
        beta_parent = S.rolling_beta(r_i, r_p, window=252, lag=1)
        _bmed_p = float(beta_parent.median()) if beta_parent.notna().any() else np.nan
        vals = parent_layer_values(parent_series, beta_parent, fast_pct_spy=spy_series.get("FAST_PCT"))
        for raw in parent_layer_specs(beta_med=_bmed_p, beta_conditional=bool(icfg.PARENT_LAYER_BETA_CONDITIONAL)):
            sec[f"{ind_ticker}__{raw.suffix}"] = vals[raw.suffix]
            specs.append(_mk_ispec(M, ind_ticker, raw))

    if icfg.USE_INDUSTRY_MACRO:
        rows = list(S.SECTOR_MACRO_TABLE.get(parent, [])) + list(INDUSTRY_MACRO_OVERRIDE.get(ind_ticker, []))
        vals = _industry_macro_values(rows, res, M, idx, spy_series.get("HAZ_PCT", pd.Series(index=idx, dtype=float)))
        for suffix, name_kr, kind, sid, transform, window, prior_sign, rationale, mech in rows:
            sec[f"{ind_ticker}__{suffix}"] = vals[suffix] if suffix in vals.columns else np.nan
            specs.append(M.IndicatorSpec(
                key=f"{ind_ticker}__{suffix}", name_kr=f"[{ind_ticker}] {name_kr}", category="B3.산업매크로",
                prior_sign=prior_sign, rationale=rationale, source=f"{kind}:{sid}",
                lead_mechanism=mech, trend_track=False, eval_horizon=_eval_h(window),
                base_series=(sid if kind != "ind_direct" else "")))

    if icfg.USE_RESID_MOMENTUM:
        rv_p = S.residual_momentum_values(adj_tr.reindex(idx), parent_tr.reindex(idx), window=icfg.RESID_MOM_WINDOW)
        sec[f"{ind_ticker}__RESID_MOM_12_1_PARENT"] = rv_p
        specs.append(M.IndicatorSpec(
            key=f"{ind_ticker}__RESID_MOM_12_1_PARENT", name_kr=f"[{ind_ticker}] 잔차모멘텀(부모중립) 12-1개월",
            category="E2.산업추세", prior_sign=+1,
            rationale="부모 베타 성분을 제거한 순수 산업 모멘텀(Blitz·Huij·Martens 2011)",
            source="산업·부모 총수익종가 36개월 롤링회귀 잔차(베타 t-1 lag)",
            lead_mechanism="베타중립 모멘텀 프리미엄", trend_track=True, eval_horizon=63))
        rv_s = S.residual_momentum_values(adj_tr.reindex(idx), spy_tr.reindex(idx), window=icfg.RESID_MOM_WINDOW)
        sec[f"{ind_ticker}__RESID_MOM_12_1_SPY"] = rv_s
        specs.append(M.IndicatorSpec(
            key=f"{ind_ticker}__RESID_MOM_12_1_SPY", name_kr=f"[{ind_ticker}] 잔차모멘텀(SPY중립) 12-1개월",
            category="E2.산업추세", prior_sign=+1,
            rationale="시장 베타 성분을 제거한 순수 산업 모멘텀",
            source="산업·SPY 총수익종가 36개월 롤링회귀 잔차(베타 t-1 lag)",
            lead_mechanism="베타중립 모멘텀 프리미엄", trend_track=True, eval_horizon=63))

    if icfg.USE_INDUSTRY_BREADTH:
        if industry_breadth is not None:
            sec[f"{ind_ticker}__IND_BREADTH_200"] = industry_breadth.reindex(idx)
            specs.append(M.IndicatorSpec(
                key=f"{ind_ticker}__IND_BREADTH_200", name_kr=f"[{ind_ticker}] 산업 폭(200일선 상회 비율)",
                category="J.시장폭", prior_sign=+1,
                rationale="200일선 위 산업이 많을수록 상승이 소수 산업에 국한되지 않고 넓게 확산돼 있다는 신호",
                source="활성 산업 각자의 자기 200일선 상회 여부", lead_mechanism="시장 내부 참여도 확산",
                trend_track=True, eval_horizon=63))
        if parent_breadth is not None:
            sec[f"{ind_ticker}__PARENT_BREADTH_200"] = parent_breadth.reindex(idx)
            specs.append(M.IndicatorSpec(
                key=f"{ind_ticker}__PARENT_BREADTH_200", name_kr=f"[{ind_ticker}] 부모 섹터 폭(200일선 상회 비율)",
                category="J.시장폭", prior_sign=+1,
                rationale="S의 섹터 폭을 그대로 재사용(재계산 없음)",
                source="8개 활성 부모 섹터 각자의 자기 200일선 상회 여부", lead_mechanism="상위 계층 내부 참여도",
                trend_track=True, eval_horizon=63))

    frames.append(sec)
    ind = pd.concat(frames, axis=1).replace([np.inf, -np.inf], np.nan)
    keys = [s.key for s in specs]
    assert len(keys) == len(set(keys)), f"{ind_ticker}: 후보지표 key 중복"
    ind = ind[keys]
    return ind, specs


# =============================================================================
# [2b] 산업층 룩어헤드 절단재계산 감사 — S.sector_lookahead_audit(§ sector_rotation.py 3191행)과
#     동일한 방법론을 산업 후보지표 구성에 적용한 것이다. 11_룩어헤드감사 시트는 v0.2로 미뤘지만
#     (파일 헤더 CHANGELOG (c) 참조), 이 함수는 그 안전성 검증을 시트가 아니라
#     test_industry_lookahead_cut.py라는 독립 회귀 테스트로 대체한다 — run()/run_industry()의
#     기존 파이프라인에는 배선하지 않는다(순수 추가, 기존 동작 변경 없음).
#
#     검사 방법: 무작위 검사일 d마다 산업·부모·SPY 원시가격과 M/S의 이미 계산된 점수 시리즈를
#     d까지로 잘라 build_industry_candidates()를 처음부터 다시 호출하고(z-score 포함), 그날
#     가중치(W.loc[d])로 점수를 재계산해 전체계산 점수와 비교한다. S 자신의 지표 인과성은 S의
#     자체 룩어헤드감사가 이미 담당하므로, 여기서는 parent_layer_series가 읽는 sres의 score_pct/
#     haz_pct를 "S가 이미 SIGNAL_START로 마스킹해 놓은 값"을 d까지 추가로 자르는 방식으로 재사용
#     한다(S를 재실행하지 않음 — sector_lookahead_audit이 M의 res["score"]를 재실행 없이 자르는
#     것과 동일한 관행, §2b 근거).
# =============================================================================
def industry_lookahead_audit(ind_ticker: str, parent: str, res: dict, sres: dict, M, icfg: IndustryConfig, S,
                             raw_df: pd.DataFrame, parent_raw_df: pd.DataFrame, spy_raw_df: pd.DataFrame,
                             W: pd.DataFrame, W_haz: pd.DataFrame,
                             score_full: pd.Series, haz_full: pd.Series, n_dates: int,
                             industry_breadth: Optional[pd.Series] = None,
                             parent_breadth: Optional[pd.Series] = None) -> pd.DataFrame:
    """[test_industry_lookahead_cut.py 전용] 반환: 티커/검사일/전체계산 점수/절단재계산 점수/차이/
    일치("OK"|"불일치"|"N/A(위험지표 미채택)")/감사종류(복합점수|위험점수(H)) 열을 가진 DataFrame.
    표본 부족(<30일)이면 '감사 생략(표본 부족)' 한 행을 반환한다(S와 동일한 관행)."""
    rng = np.random.default_rng(icfg.RANDOM_SEED)
    cal = res["cal"]
    valid = score_full.dropna().index
    valid = valid[(valid >= pd.Timestamp(res["cfg"].SIGNAL_START))]
    if len(valid) > 15:
        valid = valid[:-15]     # 마지막 15거래일은 Adj Close 지연 이어붙임 구간과 겹칠 수 있어 제외(S와 동일)
    if len(valid) < 30:
        return pd.DataFrame([{"티커": ind_ticker, "결과": "감사 생략(표본 부족)"}])
    picks = sorted(rng.choice(valid, size=min(n_dates, len(valid)), replace=False))
    rows: List[dict] = []
    for d in picks:
        d = pd.Timestamp(d)
        cal_t = cal[cal <= d]
        ind_t_raw = raw_df.loc[raw_df.index <= d]
        par_t_raw = parent_raw_df.loc[parent_raw_df.index <= d]
        spy_t_raw = spy_raw_df.loc[spy_raw_df.index <= d]
        price_t, idx_t, _ = S.sector_price_frame(ind_t_raw, cal_t, icfg)
        parent_tr_t, _ = S.build_total_return_close(par_t_raw, cal_t, icfg.ADJ_CLOSE_STALE_DAYS)
        parent_raw_t = par_t_raw["Close"].reindex(cal_t).ffill()
        spy_tr_t, _ = S.build_total_return_close(spy_t_raw, cal_t, icfg.ADJ_CLOSE_STALE_DAYS)
        spy_raw_t = spy_t_raw["Close"].reindex(cal_t).ffill()
        res_t = {"ind": res["ind"].loc[res["ind"].index <= d],
                 "fred": {k: (v.loc[v.index <= d] if v is not None else None) for k, v in res["fred"].items()},
                 "px_dict": {k: (v.loc[v.index <= d] if v is not None else None) for k, v in res["px_dict"].items()},
                 "score": res["score"].loc[res["score"].index <= d],
                 "haz_score": res["haz_score"].loc[res["haz_score"].index <= d],
                 "cfg": res["cfg"]}
        spy_series_t = S.spy_layer_series(res_t, M)
        # [§2b] parent_layer_series가 읽는 것은 sres["sectors"][parent]의 마스킹된 score_pct/haz_pct
        # 뿐이다 — S를 재실행하지 않고 그 두 시리즈를 d까지 추가로 잘라 그대로 재사용한다.
        sr = sres["sectors"].get(parent, {})
        sres_t = {"sectors": {parent: {
            "score_pct": sr.get("score_pct", pd.Series(dtype=float)).loc[:d],
            "haz_pct": sr.get("haz_pct", pd.Series(dtype=float)).loc[:d]}}}
        parent_series_t = parent_layer_series(sres_t, parent, M, mode=icfg.PARENT_LAYER_SOURCE)
        breadth_t = industry_breadth.loc[industry_breadth.index <= d] if industry_breadth is not None else None
        pbreadth_t = parent_breadth.loc[parent_breadth.index <= d] if parent_breadth is not None else None
        ind_i_t, _ = build_industry_candidates(
            ind_ticker, parent, res_t, M, icfg, S, price_t["Adj Close"].astype(float), price_t["Close"].astype(float),
            spy_tr_t, spy_raw_t, parent_tr_t, parent_raw_t, spy_series_t, parent_series_t, idx_t,
            industry_breadth=breadth_t, parent_breadth=pbreadth_t)
        Z_t = pd.DataFrame({k: M.expanding_zscore(ind_i_t[k]) for k in ind_i_t.columns}, index=ind_i_t.index)
        if d not in Z_t.index:
            continue
        z_row = Z_t.loc[d]
        for kind, Wm, full in (("복합점수", W, score_full), ("위험점수(H)", W_haz, haz_full)):
            w_row = Wm.loc[d].reindex(z_row.index).fillna(0.0) if d in Wm.index else pd.Series(0.0, index=z_row.index)
            s_t = S._score_with_weights(z_row, w_row)
            s_full = float(full.loc[d]) if d in full.index else np.nan
            if kind == "위험점수(H)" and pd.isna(s_t) and pd.isna(s_full):
                rows.append({"티커": ind_ticker, "검사일": str(d.date()), "전체계산 점수": np.nan,
                            "절단재계산 점수": np.nan, "차이": np.nan, "일치": "N/A(위험지표 미채택)", "감사종류": kind})
                continue
            diff = abs(s_t - s_full) if pd.notna(s_t) and pd.notna(s_full) else np.nan
            ok = bool(pd.notna(diff) and diff < 1e-8)
            rows.append({"티커": ind_ticker, "검사일": str(d.date()),
                        "전체계산 점수": round(s_full, 8) if pd.notna(s_full) else np.nan,
                        "절단재계산 점수": round(s_t, 8) if pd.notna(s_t) else np.nan,
                        "차이": diff, "일치": "OK" if ok else "불일치", "감사종류": kind})
            log("AUDIT", kv(ticker=ind_ticker, kind=kind, date=str(d.date()),
                            full=s_full if pd.notna(s_full) else -99, truncated=s_t if pd.notna(s_t) else -99,
                            result="OK" if ok else "MISMATCH"), M=M, level="info" if ok else "error")
    return pd.DataFrame(rows)


# =============================================================================
# [3] 단일 산업 파이프라인 — run_industry(ind_ticker, ctx). S.run_sector를 일반화 복제.
# =============================================================================
def run_industry(ind_ticker: str, ctx: Dict[str, Any]) -> Dict[str, Any]:
    """[§1.3, run_sector -> run_industry] ctx 필수 키: M, res, sres, S, icfg,
    frames(dict[ticker -> (price_i, idx_i, info)]), raw(dict[ticker -> raw df]),
    parent(str), parent_frames/parent_raw(부모 가격), spy_tr/spy_raw/spy_series,
    parent_series(dict[parent -> {SCORE_PCT,HAZ_PCT}]), rf_daily,
    industry_breadth_200/parent_breadth_200(선택)."""
    t0 = time.time()
    M, res, sres, S, icfg = ctx["M"], ctx["res"], ctx["sres"], ctx["S"], ctx["icfg"]
    M_cfg = res["cfg"]
    parent = ctx["parent_of"][ind_ticker]
    price_i, idx_i, info = ctx["frames"][ind_ticker]
    adj_i = price_i["Adj Close"].astype(float)
    close_i = price_i["Close"].astype(float)
    spy_tr, spy_raw, spy_series, rf = ctx["spy_tr"], ctx["spy_raw"], ctx["spy_series"], ctx["rf_daily"]
    parent_tr = ctx["parent_tr"][parent]
    parent_raw = ctx["parent_raw"][parent]
    parent_series = ctx["parent_series"][parent]
    industry_breadth = ctx.get("industry_breadth_200")
    parent_breadth = ctx.get("parent_breadth_200")

    timing: Dict[str, float] = {}
    cfg_i = S.sector_cfg_for(M_cfg, icfg, ind_ticker, idx_i)
    t1 = time.time()

    ind_i, specs = build_industry_candidates(
        ind_ticker, parent, res, M, icfg, S, adj_i, close_i, spy_tr, spy_raw, parent_tr, parent_raw,
        spy_series, parent_series, idx_i, industry_breadth=industry_breadth, parent_breadth=parent_breadth)
    trend200 = S.sector_technical_values(adj_i)["TREND_200"]
    t2 = time.time()
    timing["01_지표생성"] = round(t2 - t1, 2)

    hv = S.validate_and_weight_sector(ind_ticker, ind_i, adj_i, cfg_i, specs, M, icfg)
    val_full, W, wlog, W_haz = hv["val_full"], hv["W"], hv["wlog"], hv["W_haz"]
    t3 = time.time()
    timing["02_검증+워크포워드"] = round(t3 - t2, 2)

    vol_scale = S._sector_vol_scale(adj_i, parent_tr, icfg) if icfg.VOL_SCALE_OVERLAYS else 1.0
    cfg_i = S._vol_scaled_cfg(cfg_i, vol_scale)

    with S._indicator_spec_override(M, specs):
        score, contrib, n_used = M.composite_score(ind_i, W, cfg_i)
        haz_score, haz_contrib, _ = M.composite_score(ind_i, W_haz, cfg_i)
        score_pct_full = M.score_percentile(score)
        haz_pct_industry_full = M.score_percentile(haz_score)

    sig_mask = idx_i >= pd.Timestamp(cfg_i.SIGNAL_START)
    score_pct = score_pct_full.where(sig_mask)
    haz_pct_industry = haz_pct_industry_full.where(sig_mask)

    if icfg.HAZARD_SOURCE == "sector" or icfg.HAZARD_SOURCE == "parent":
        pr_haz = ctx["parent_series"][parent].get("HAZ_PCT", pd.Series(index=idx_i, dtype=float)).reindex(idx_i)
        haz_pct = pr_haz
    elif icfg.HAZARD_SOURCE == "max":
        spy_h = res.get("haz_pct", pd.Series(index=idx_i, dtype=float)).reindex(idx_i)
        pr_haz = ctx["parent_series"][parent].get("HAZ_PCT", pd.Series(index=idx_i, dtype=float)).reindex(idx_i)
        haz_pct = pd.concat([spy_h, pr_haz], axis=1).max(axis=1)
    else:
        haz_pct = res.get("haz_pct", pd.Series(index=idx_i, dtype=float)).reindex(idx_i)

    fast_pct = res.get("fast_pct", pd.Series(index=idx_i, dtype=float)).reindex(idx_i) if getattr(cfg_i, "USE_FAST_TRIGGER", False) else None
    # [v0.1.0 §범위] 회복플로어/깊은낙폭/구조적바닥 규칙은 비활성(recov_conf/deep_recov/struct_dd=None) — 헤더 CHANGELOG (b) 참조.
    recov_conf = deep_recov = struct_dd = None

    with S._indicator_spec_override(M, specs):
        sig = M.generate_signals(score_pct, trend200, cfg_i, score=score, haz_pct=haz_pct, fast_pct=fast_pct,
                                 recov_conf=recov_conf, deep_recov=deep_recov, struct_dd=struct_dd)
        reason = M.build_reason_text(contrib, sig["state"], score)
    t4 = time.time()
    timing["03_신호생성"] = round(t4 - t3, 2)

    bt = M.run_backtest(price_i, sig["target_pos"], cfg_i, rf)
    bt = bt.loc[bt.index >= pd.Timestamp(cfg_i.SIGNAL_START)]
    ma_pos = (trend200 > 0).astype(float).where(sig_mask, 0.0)
    bt_ma = M.run_backtest(price_i, ma_pos, cfg_i, rf)
    t5 = time.time()
    timing["04_백테스트"] = round(t5 - t4, 2)

    adopted = sorted({k for k in W.columns if (W[k] != 0).any()})
    try:
        events = M.event_study(ind_i, adj_i, bt, adopted, cfg_i, haz_pct=haz_pct)
    except Exception as e:
        events = pd.DataFrame()
        log("PIPE", kv(event="event_study_failed", ticker=ind_ticker, err=str(e)[:120]), M=M, level="warning")
    trades = M.extract_trades(bt, reason, sig["state"])
    episodes = M.drawdown_episodes(bt, cfg_i)
    t6 = time.time()
    timing["05_이벤트+거래+구간"] = round(t6 - t5, 2)

    sheets = S.build_sector_sheets(
        M, ind_ticker, cfg_i, specs, price_i, bt, bt_ma, sig, score, score_pct, n_used,
        haz_score, haz_pct, fast_pct, recov_conf, reason, ind_i, contrib, W, W_haz,
        wlog, val_full, adopted, trades, episodes, events, pd.DataFrame(), pd.DataFrame(),
        haz_pct_sector=haz_pct_industry, spy_yearly_pos=None, daily_indicator_detail=False)

    # 풀링 순환매용 원자료(§6.2) — 이미 계산된 후보열에서 부모/SPY 상대 신호를 그대로 뽑아 재사용(재계산 없음).
    rot_cols = {
        "PARENT_SCORE_PCT": f"{ind_ticker}__PARENT_SCORE_PCT", "PBETA_X_SCORE": f"{ind_ticker}__PBETA_X_SCORE",
        "P_REL_MOM_126": f"{ind_ticker}__P_REL_MOM_126", "P_REL_MOM_12_1": f"{ind_ticker}__P_REL_MOM_12_1",
        "P_REL_MOM_21": f"{ind_ticker}__P_REL_MOM_21", "P_REL_EXT_200": f"{ind_ticker}__P_REL_EXT_200",
        "RESID_MOM_12_1_PARENT": f"{ind_ticker}__RESID_MOM_12_1_PARENT",
    }
    rot_raw = pd.DataFrame(index=idx_i)
    for name, col in rot_cols.items():
        rot_raw[name] = ind_i[col] if col in ind_i.columns else np.nan
    # [§6.2] 풀링 순환매용 '부모초과수익' 일간수익률 — rel=산업/부모 가격비율(S의 상대강도 산식과
    # 동일 정의)의 일간수익률. 전체(마스킹 전) 이력 — rotation_walkforward_select가 SIGNAL_START
    # 이전 구간도 학습에 쓴다(S와 동일 관행).
    rel_px = adj_i / parent_tr.reindex(idx_i).replace(0, np.nan)
    rot_raw["REL_RET"] = rel_px.pct_change()
    ret_cc_full = adj_i.pct_change()

    timing["06_run_industry합계"] = round(time.time() - t0, 2)
    first_signal = score_pct.dropna().index.min()

    return {
        "ticker": ind_ticker, "parent": parent, "info": info, "cfg_dict": dataclasses.asdict(cfg_i),
        "timing": timing, "cache_hit": bool(hv.get("cache_hit", False)), "n_candidates": len(specs),
        "first_signal": (str(first_signal.date()) if pd.notna(first_signal) else None),
        "adopted": adopted, "sheets": sheets, "hazard_source": icfg.HAZARD_SOURCE, "vol_scale": vol_scale,
        "state": sig["state"].loc[sig.index >= pd.Timestamp(cfg_i.SIGNAL_START)],
        "target_pos": sig["target_pos"].loc[sig.index >= pd.Timestamp(cfg_i.SIGNAL_START)],
        "score_pct": score_pct.loc[sig_mask], "haz_pct": haz_pct.loc[sig_mask],
        "haz_pct_industry": haz_pct_industry.loc[sig_mask],
        "strategy_ret": bt["strategy_ret"], "bh_ret": bt["bh_ret"], "pos_exec": bt["pos_exec"],
        "ma_ret": bt_ma["strategy_ret"], "ret_co": bt["ret_co"], "ret_oc": bt["ret_oc"],
        "rot_raw": rot_raw, "px_open": price_i["Open"].astype(float).loc[sig_mask],
        "px_close": close_i.loc[sig_mask], "ret_cc_full": ret_cc_full,
    }


def run_industries(tickers: List[str], ctx: Dict[str, Any], icfg: IndustryConfig, M
                   ) -> Tuple[Dict[str, Dict[str, Any]], Dict[str, str]]:
    """[§12 ⚠ v0.1.0 범위] 순차 실행만 지원한다(S의 fork 병렬은 이번 버전에 이식하지 않음 —
    헤더 CHANGELOG (a) 참조). 19개 산업 직렬 실행은 병렬 대비 느리지만 항상 재현 가능하다."""
    results: Dict[str, Dict[str, Any]] = {}
    failed: Dict[str, str] = {}
    for i, t in enumerate(tickers):
        try:
            log("PIPE", kv(event="industry_progress", done=i, total=len(tickers), ticker=t), M=M)
            results[t] = run_industry(t, ctx)
        except Exception as e:
            failed[t] = f"{type(e).__name__}: {e}\n{traceback.format_exc()[-3000:]}"
            log("PIPE", kv(event="industry_failed", ticker=t, err=str(e)[:200]), M=M, level="error")
    return results, failed


# =============================================================================
# [4] 풀링 횡단면 산업 순환매 검증(§6) — S.rotation_walkforward_select를 그대로 호출.
# =============================================================================
def build_pooled_rotation(results: Dict[str, Dict[str, Any]], eval_idx: pd.DatetimeIndex,
                          icfg: IndustryConfig, M, S) -> Dict[str, Any]:
    """[§6.2] 횡단면 = 활성 산업 전부(SPY/부모는 후보 아님, ROTATION_INCLUDE_SPY_CANDIDATE=False).
    ret_cc_full은 '산업가격/부모가격' 비율의 일간수익률 — S가 상대강도 전반에 쓰는 것과 같은
    비율기반 정의로, 그 비율의 향후 h일 수익률이 정확히 '산업 h일 수익률 - 부모 h일 수익률'의
    곱셈적(로그) 근사가 된다(§6.2 "부모 초과수익"의 구현)."""
    cols = list(results.keys())
    full_idx = None
    for t in cols:
        idx = results[t]["ret_cc_full"].index
        full_idx = idx if full_idx is None else full_idx.union(idx)
    full_idx = full_idx.sort_values()

    # ret_cc_full(풀링 순환매용) = rot_raw["REL_RET"](산업/부모 가격비율의 일간수익률, run_industry에서
    # 이미 전체이력으로 계산됨) — 재계산 없음.
    ret_cc_full = pd.DataFrame({t: results[t]["rot_raw"].get("REL_RET", pd.Series(dtype=float)) for t in cols})
    ret_cc_full = ret_cc_full.reindex(full_idx)
    listed_full = ret_cc_full.notna()

    sig_full: Dict[str, pd.DataFrame] = {}
    for name in icfg.ROTATION_SIGNALS:
        mat = pd.DataFrame({t: results[t]["rot_raw"].get(name, pd.Series(dtype=float)) for t in cols})
        sig_full[name] = mat.reindex(full_idx)

    half_life = icfg.ROTATION_DECAY_HALF_LIFE_DAYS
    wf = S.rotation_walkforward_select(sig_full, ret_cc_full, listed_full, eval_idx, icfg, M,
                                       external=None, half_life_days=half_life)
    log("ROT", kv(event="pooled_rank_ready", n=len(cols), horizon=wf.get("horizon"),
                  mode=wf.get("mode"), stat=wf.get("stat")), M=M)
    return wf


# =============================================================================
# [5] 계층 배분(§7) — build_industry_allocation. S의 build_sector_allocation과 달리 SPY가
#     하드코딩된 12번째 후보/폴백 자산 구조가 아니라, "부모 비중 안에서 산업 vs 부모ETF"를
#     결정하는 계층적 배분이라 일반화 복제가 아닌 신규 구현이다(설계서 §7.3 근거).
# =============================================================================
def _leader_and_gate(sub_composite: pd.DataFrame, sub_eligible: pd.DataFrame, margin_steps: float
                     ) -> Tuple[pd.Series, pd.Series, pd.Series]:
    """[§7.2 '여유 게이트'] S의 컨빅션게이트(순위 마진)를 산업-부모 부분집합에 적용한 것.
    rank_pct(0..1, 1=최고)의 1위-2위 격차가 margin_steps/(n_ok-1) 이상이면 '리더' 통과.
    n_ok<=1이면 격차 계산이 불가하므로 자동 통과(있는 게 곧 1위)."""
    n_ok = sub_eligible.sum(axis=1)
    masked = sub_composite.where(sub_eligible)
    rank_pct = masked.rank(axis=1, ascending=True, pct=True)
    top1 = rank_pct.max(axis=1)
    is_top = rank_pct.eq(top1, axis=0) & sub_eligible
    # 동률(예: n_ok==1)이면 top2가 없다 — margin=1.0(항상 통과)로 취급.
    rank_pct_wo_top = rank_pct.where(~is_top)
    top2 = rank_pct_wo_top.max(axis=1)
    margin = (top1 - top2).fillna(1.0)
    step = (margin_steps / (n_ok - 1).clip(lower=1)).astype(float)
    gate_pass = (n_ok >= 1) & (margin >= step - 1e-9)
    # v0.1.1 수정: n_ok==0인 날(그 부모 소속 산업이 전부 NaN)에는 masked 행 전체가 NaN이 되어
    # .idxmax(axis=1)가 "Encountered all NA values"로 예외를 던진다(원래 다음 줄의
    # .where(n_ok>=1)로 널처리할 계획이었으나 idxmax 자체가 먼저 죽어서 도달하지 못했다).
    # -inf 센티널로 전NaN 행을 채워 idxmax가 항상 값을 반환하게 한 뒤, 그 '가짜' 리더를
    # 기존 계획대로 n_ok>=1 마스크로 널처리한다 — 최종 leader 값 자체는 변경 없음.
    leader = masked.fillna(-np.inf).idxmax(axis=1)
    leader = leader.where(n_ok >= 1)
    return leader, gate_pass, n_ok


def _build_industry_frac(icfg: IndustryConfig, leader_cap: float, fallback_share: float,
                         cols: List[str], active_parents: List[str], parent_of: Dict[str, str],
                         composite_all_s: pd.DataFrame, eligible: pd.DataFrame,
                         ) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """[§7.2] 반환: frac(date x industry, 그날 w_s[parent]에 곱할 비율), parent_frac(date x parent,
    그날 w_s[parent]에 곱할 부모ETF 잔여비율). frac.sum(그 부모 소속) + parent_frac[parent] == 1.0
    (적격 산업이 있는 날) 또는 parent_frac==1.0·frac==0(적격 산업이 없는 날 — 전부 부모ETF)."""
    idx = composite_all_s.index
    frac = pd.DataFrame(0.0, index=idx, columns=cols)
    parent_frac = pd.DataFrame(1.0, index=idx, columns=list(active_parents))
    for parent in active_parents:
        inds = [c for c in cols if parent_of[c] == parent]
        if not inds:
            continue
        leader, gate_pass, n_ok = _leader_and_gate(composite_all_s[inds], eligible[inds],
                                                   icfg.ROTATION_LEADER_MARGIN_STEPS)
        has_elig = n_ok >= 1
        lead_day = has_elig & gate_pass
        fb_day = has_elig & ~gate_pass
        if lead_day.any():
            for d in idx[lead_day]:
                frac.loc[d, leader.loc[d]] = leader_cap
            parent_frac.loc[lead_day, parent] = 1.0 - leader_cap
        if fb_day.any():
            elig_mat = eligible.loc[fb_day, inds]
            n_fb = elig_mat.sum(axis=1).replace(0, np.nan)
            share = (fallback_share / n_fb)
            for c in inds:
                frac.loc[fb_day, c] = np.where(elig_mat[c].values, share.values, 0.0)
            parent_frac.loc[fb_day, parent] = 1.0 - fallback_share
    return frac, parent_frac


def _industry_portfolio_backtest(S, target_w: pd.DataFrame, ret_co: pd.DataFrame, ret_oc: pd.DataFrame,
                                 rf_daily: Optional[pd.Series], industry_cols: List[str],
                                 parent_cols: List[str], cost_ind_bps: float, cost_par_bps: float,
                                 init_exec: Optional[pd.Series] = None, init_prev: Optional[pd.Series] = None
                                 ) -> pd.DataFrame:
    """[§7.5] S.portfolio_backtest는 비용을 단일 스칼라(cost_bps)로만 받는다. 산업(10bp)·부모(5bp)
    비용률을 각각 반영하기 위해 cost_bps=0으로 총수익(비용 미차감)을 얻은 뒤, 자산군별 회전율에
    각자의 비용률을 곱해 직접 차감한다(이중차감 없음 — gross 쪽 cost 열은 항상 0)."""
    gross = S.portfolio_backtest(target_w, ret_co, ret_oc, rf_daily, cost_bps=0.0,
                                 init_exec=init_exec, init_prev=init_prev)
    exec_w = target_w.shift(1)
    if init_exec is not None and len(target_w.index):
        exec_w.iloc[0] = init_exec.reindex(target_w.columns).fillna(0.0)
    exec_w = exec_w.fillna(0.0)
    prev_w = exec_w.shift(1)
    if init_prev is not None and len(target_w.index):
        prev_w.iloc[0] = init_prev.reindex(target_w.columns).fillna(0.0)
    prev_w = prev_w.fillna(0.0)
    turn_ind = (exec_w[industry_cols] - prev_w[industry_cols]).abs().sum(axis=1)
    turn_par = (exec_w[parent_cols] - prev_w[parent_cols]).abs().sum(axis=1)
    cost = turn_ind * (cost_ind_bps / 1e4) + turn_par * (cost_par_bps / 1e4)
    out = gross.copy()
    out["turnover"] = turn_ind + turn_par
    out["cost"] = cost
    out["strategy_ret"] = (gross["strategy_ret"] - cost).fillna(0.0)
    out["equity"] = (1.0 + out["strategy_ret"]).cumprod()
    out["dd"] = out["equity"] / out["equity"].cummax() - 1.0
    return out


def build_industry_allocation(results: Dict[str, Dict[str, Any]], sres: dict, res: dict,
                              eval_idx: pd.DatetimeIndex, icfg: IndustryConfig, M, S,
                              wf: Dict[str, Any], rf_daily: Optional[pd.Series] = None
                              ) -> Dict[str, Any]:
    cols = list(results.keys())
    if not cols:
        return {}
    parent_of = {t: results[t]["parent"] for t in cols}
    active_parents = sorted({parent_of[t] for t in cols})
    all_cols = cols + active_parents

    state = pd.DataFrame({t: results[t]["state"] for t in cols}).reindex(eval_idx)
    listed = state.notna()
    eligible = listed & ~state.isin(icfg.INDUSTRY_EXCLUDE_STATES)

    composite_all_s = wf["composite_all"].rolling(icfg.ROTATION_SMOOTH_DAYS, min_periods=1).mean()
    composite_all_s = composite_all_s.reindex(index=eval_idx, columns=cols)

    w_s_all = sres.get("alloc", {}).get("target_w", pd.DataFrame())
    w_s = w_s_all.reindex(index=eval_idx, columns=active_parents).fillna(0.0)

    ret_co = pd.DataFrame(index=eval_idx, columns=all_cols, dtype=float)
    ret_oc = pd.DataFrame(index=eval_idx, columns=all_cols, dtype=float)
    for t in cols:
        ret_co[t] = results[t]["ret_co"].reindex(eval_idx)
        ret_oc[t] = results[t]["ret_oc"].reindex(eval_idx)
    for p in active_parents:
        pr = sres["sectors"].get(p, {})
        ret_co[p] = pr.get("ret_co", pd.Series(dtype=float)).reindex(eval_idx)
        ret_oc[p] = pr.get("ret_oc", pd.Series(dtype=float)).reindex(eval_idx)
    ret_co = ret_co.fillna(0.0)
    ret_oc = ret_oc.fillna(0.0)

    def _mk_target_w(leader_cap: float, fallback_share: float) -> pd.DataFrame:
        frac, parent_frac = _build_industry_frac(icfg, leader_cap, fallback_share, cols, active_parents,
                                                  parent_of, composite_all_s, eligible)
        tw = pd.DataFrame(0.0, index=eval_idx, columns=all_cols)
        for t in cols:
            p = parent_of[t]
            tw[t] = frac[t].reindex(eval_idx).fillna(0.0) * w_s[p]
        for p in active_parents:
            tw[p] = parent_frac[p].reindex(eval_idx).fillna(1.0) * w_s[p]
        return tw

    def _bt(tw: pd.DataFrame) -> pd.DataFrame:
        return _industry_portfolio_backtest(S, tw, ret_co, ret_oc, rf_daily, cols, active_parents,
                                            icfg.COST_BPS_INDUSTRY, icfg.PARENT_COST_BPS)

    live_cap, live_fb = float(icfg.INDUSTRY_LEADER_CAP), float(icfg.INDUSTRY_FALLBACK_SHARE)
    label_star = f"부모비중 안 산업리더 {live_cap:.0%}·하락 시 부모 ★"
    target_ws: Dict[str, pd.DataFrame] = {label_star: _mk_target_w(live_cap, live_fb)}

    for cv in (0.25, 0.5, 0.75, 1.0):
        if abs(cv - live_cap) < 1e-9:
            continue
        target_ws[f"산업리더 상한 {cv:.0%} [산업집중격자]"] = _mk_target_w(cv, live_fb)
    for fv in (0.0, 0.25, 0.75, 1.0):
        if abs(fv - live_fb) < 1e-9:
            continue
        target_ws[f"폴백 균등배분 {fv:.0%} [폴백격자]"] = _mk_target_w(live_cap, fv)

    label_ctrl_a = "대조군A: S★ 그대로(산업 미사용)"
    ctrl_tw = pd.DataFrame(0.0, index=eval_idx, columns=all_cols)
    for p in active_parents:
        ctrl_tw[p] = w_s[p]
    target_ws[label_ctrl_a] = ctrl_tw

    bts: Dict[str, pd.DataFrame] = {}
    star_label = next((c for c in sres.get("alloc", {}).get("bts", {}) if str(c).endswith("★")), None)
    for label, tw in target_ws.items():
        if label == label_ctrl_a and star_label is not None:
            # [§7.4] "= S★ 비트 동일" — S가 이미 계산한 바로 그 일간수익을 그대로 쓴다(재계산으로
            # 인한 미세한 비용/반올림 차이조차 없게).
            s_ret = sres["alloc"]["bts"][star_label]["strategy_ret"].reindex(eval_idx).fillna(0.0)
            eq = (1.0 + s_ret).cumprod()
            bts[label] = pd.DataFrame({"strategy_ret": s_ret, "exposure": tw.sum(axis=1),
                                       "turnover": 0.0, "cost": 0.0, "equity": eq,
                                       "dd": eq / eq.cummax() - 1.0}, index=eval_idx)
        else:
            bts[label] = _bt(tw)

    order = [l for l in target_ws if l != label_ctrl_a] + [label_ctrl_a]
    perf = pd.DataFrame([M.perf_metrics(bts[l]["strategy_ret"], l) for l in order])
    for col_extra, fn in (("평가창", lambda l: f"{eval_idx[0].date()}~{eval_idx[-1].date()}"),
                          ("평균노출", lambda l: round(float(target_ws[l].sum(axis=1).mean()), 4)),
                          ("누적비용(%p)", lambda l: round(float(bts[l]["cost"].sum()) * 100, 3))):
        perf[col_extra] = [fn(l) for l in order]

    base_r = bts[label_star]["strategy_ret"]
    excess_list, robust_list = [], []
    for l in order:
        r = bts[l]["strategy_ret"]
        exc = float((r - base_r).mean() * 252 * 100) if l != label_star else 0.0
        if l == label_star:
            robust = "★"
        else:
            top5 = (r - base_r).sort_values(ascending=False).index[:5]
            exc_wo5 = float((r.drop(top5) - base_r.drop(top5)).mean() * 252 * 100)
            robust = "통과" if np.sign(exc_wo5 + 1e-12) == np.sign(exc + 1e-12) or abs(exc) < 1e-9 else "⚠ 5일 의존"
        excess_list.append(round(exc, 3))
        robust_list.append(robust)
    perf["★대비 초과(연율%p)"] = excess_list
    perf["강건성(기준④)"] = robust_list

    curve = pd.DataFrame(index=eval_idx)
    curve["날짜"] = eval_idx.date
    for l in order:
        curve[l] = bts[l]["equity"].values
    curve["E_t(부모비중 합)"] = w_s.sum(axis=1).values

    return {
        "target_w": target_ws[label_star], "target_ws": target_ws, "bts": bts, "perf": perf,
        "curve": curve.reset_index(drop=True), "cols": cols, "active_parents": active_parents,
        "parent_of": parent_of, "all_cols": all_cols, "eligible": eligible, "state": state,
        "w_s": w_s, "composite_all_s": composite_all_s, "label_star": label_star,
        "label_ctrl_a": label_ctrl_a, "ret_co": ret_co, "ret_oc": ret_oc,
        "cost_bps_industry": icfg.COST_BPS_INDUSTRY, "cost_bps_parent": icfg.PARENT_COST_BPS,
        "rf_daily": rf_daily,
    }


# =============================================================================
# [6] 수용기준(§8) · 계층정합(14) · 산업대섹터귀속(15)
# =============================================================================
def build_industry_acceptance(alloc: Dict[str, Any], wf: Dict[str, Any], icfg: IndustryConfig,
                              M, S) -> pd.DataFrame:
    """[§8] I★ vs S★ 5기준. ③은 풀링 채택신호들의 (상위1-하위1)스프레드 평균에 NW-HAC t(설계서의
    '풀링 상위1 스프레드'를 상위1-하위1 차분으로 근사 — S는 top3-bottom3 스프레드용 헬퍼를 별도로
    노출하지 않아 top1/bottom1로 근사한다, v0.1.0 §범위). ④는 연도별 그 스프레드 평균의 부호로
    '양수 구간' 비율을 잰다."""
    perf = alloc.get("perf", pd.DataFrame())

    def _get(label, col):
        try:
            r = perf[perf["전략"] == label]
            if len(r):
                return float(r.iloc[0][col])
        except Exception:
            pass
        return float("nan")

    star, ctrl = alloc.get("label_star"), alloc.get("label_ctrl_a")
    cagr_i, cagr_s = _get(star, "CAGR"), _get(ctrl, "CAGR")
    mdd_i, mdd_s = _get(star, "최대낙폭(MDD)"), _get(ctrl, "최대낙폭(MDD)")
    calmar_i, calmar_s = _get(star, "칼마(CAGR/MDD)"), _get(ctrl, "칼마(CAGR/MDD)")

    adopted = sorted({s for lst in wf.get("selected_by_year", {}).values() for s in lst})
    spreads = []
    for s in adopted:
        t1 = wf.get("top1_full", {}).get(s)
        b1 = wf.get("bottom1_full", {}).get(s)
        if t1 is not None and b1 is not None:
            spreads.append(t1 - b1)
    spread_avg = pd.concat(spreads, axis=1).mean(axis=1) if spreads else pd.Series(dtype=float)
    if len(spread_avg.dropna()) >= 30:
        _, t3, n3 = S._nw_mean_tstat(spread_avg.dropna(), lag=21)
    else:
        t3, n3 = float("nan"), int(spread_avg.notna().sum())
    yearly = spread_avg.dropna().groupby(spread_avg.dropna().index.year).mean() if len(spread_avg.dropna()) else pd.Series(dtype=float)
    pos_ratio = float((yearly > 0).mean()) if len(yearly) else float("nan")

    rows = [
        {"기준": "① CAGR ≥ S★+기준", "I★": round(cagr_i, 4) if cagr_i == cagr_i else None,
         "S★": round(cagr_s, 4) if cagr_s == cagr_s else None, "기준값": icfg.ACCEPT_CAGR_GAIN,
         "실측": round(cagr_i - cagr_s, 4) if (cagr_i == cagr_i and cagr_s == cagr_s) else None,
         "판정": "PASS" if (cagr_i == cagr_i and cagr_s == cagr_s and cagr_i - cagr_s >= icfg.ACCEPT_CAGR_GAIN) else "FAIL"},
        {"기준": "② MDD 악화 ≤ 기준", "I★": round(mdd_i, 4) if mdd_i == mdd_i else None,
         "S★": round(mdd_s, 4) if mdd_s == mdd_s else None, "기준값": icfg.ACCEPT_MDD_WORSE,
         "실측": round(mdd_s - mdd_i, 4) if (mdd_i == mdd_i and mdd_s == mdd_s) else None,
         "판정": "PASS" if (mdd_i == mdd_i and mdd_s == mdd_s and (mdd_s - mdd_i) <= icfg.ACCEPT_MDD_WORSE) else "FAIL"},
        {"기준": "③ 풀링 상위1-하위1 스프레드 NW-t ≥ 기준", "I★": round(t3, 3) if t3 == t3 else None,
         "S★": None, "기준값": icfg.ACCEPT_TOP1_T, "실측": n3,
         "판정": "PASS" if (t3 == t3 and t3 >= icfg.ACCEPT_TOP1_T) else "FAIL"},
        {"기준": "④ 스프레드 양수 구간 비율 ≥ 기준", "I★": round(pos_ratio, 3) if pos_ratio == pos_ratio else None,
         "S★": None, "기준값": round(icfg.ACCEPT_SPREAD_PERIODS_RATIO, 3), "실측": len(yearly),
         "판정": "PASS" if (pos_ratio == pos_ratio and pos_ratio >= icfg.ACCEPT_SPREAD_PERIODS_RATIO) else "FAIL"},
        {"기준": "⑤ 칼마 ≥ S★ 칼마(동률 불가)", "I★": round(calmar_i, 3) if calmar_i == calmar_i else None,
         "S★": round(calmar_s, 3) if calmar_s == calmar_s else None, "기준값": 0.0,
         "실측": round(calmar_i - calmar_s, 3) if (calmar_i == calmar_i and calmar_s == calmar_s) else None,
         "판정": "PASS" if (calmar_i == calmar_i and calmar_s == calmar_s and calmar_i > calmar_s) else "FAIL"},
    ]
    return pd.DataFrame(rows)


def build_hierarchy_check(alloc: Dict[str, Any]) -> pd.DataFrame:
    """[14_계층정합, §7.1] 매일 Σ산업 + 부모ETF = S★의 그 섹터비중(허용오차 1e-9)."""
    tw, cols, parent_of, active_parents, w_s = (alloc["target_w"], alloc["cols"], alloc["parent_of"],
                                                 alloc["active_parents"], alloc["w_s"])
    rows, max_err_all, viol_all = [], 0.0, 0
    for p in active_parents:
        inds = [c for c in cols if parent_of[c] == p]
        total = tw[inds].sum(axis=1) + tw[p]
        err = (total - w_s[p]).abs()
        max_err = float(err.max()) if len(err) else 0.0
        n_viol = int((err > 1e-9).sum())
        max_err_all = max(max_err_all, max_err)
        viol_all += n_viol
        rows.append({"부모섹터": p, "산업수": len(inds), "최대오차": max_err, "위반일수(>1e-9)": n_viol})
    rows.append({"부모섹터": "전체", "산업수": len(cols), "최대오차": max_err_all, "위반일수(>1e-9)": viol_all})
    return pd.DataFrame(rows)


def build_industry_vs_sector_attribution(alloc: Dict[str, Any]) -> pd.DataFrame:
    """[15_산업대섹터귀속] I★(부모비중 안에서 산업으로 나눠 담은 것)이 S★(부모ETF 100%) 대비
    무엇을 얻었는지 — 일별 초과수익 누적·IR·꼬리·연도별."""
    star = alloc["bts"][alloc["label_star"]]["strategy_ret"]
    ctrl = alloc["bts"][alloc["label_ctrl_a"]]["strategy_ret"]
    idx = star.index.intersection(ctrl.index)
    excess = (star.reindex(idx) - ctrl.reindex(idx)).dropna()
    rows: List[dict] = []
    if len(excess):
        ir = float(excess.mean() / excess.std() * np.sqrt(252)) if excess.std() > 0 else float("nan")
        rows.append({"구분": "A. 총괄", "일평균초과(연율%)": round(float(excess.mean() * 252 * 100), 3),
                    "초과변동성(연율%)": round(float(excess.std() * np.sqrt(252) * 100), 3),
                    "정보비율(IR)": round(ir, 3) if ir == ir else None,
                    "누적초과(%)": round(float(((1.0 + excess).prod() - 1.0) * 100), 3)})
        for d, v in excess.sort_values(ascending=False).head(5).items():
            rows.append({"구분": "B. 큰 양(+)초과일", "날짜": str(pd.Timestamp(d).date()), "초과(%)": round(float(v * 100), 4)})
        for d, v in excess.sort_values().head(5).items():
            rows.append({"구분": "B. 큰 음(-)초과일", "날짜": str(pd.Timestamp(d).date()), "초과(%)": round(float(v * 100), 4)})
        yr = excess.groupby(excess.index.year).apply(lambda s: (1.0 + s).prod() - 1.0)
        for y, v in yr.items():
            rows.append({"구분": "C. 연도별", "연도": int(y), "초과수익(%)": round(float(v * 100), 3)})
    return pd.DataFrame(rows)


def build_industry_prediction_matrix(results: Dict[str, Dict[str, Any]], eval_idx: pd.DatetimeIndex,
                                     icfg: IndustryConfig, S) -> pd.DataFrame:
    """[01Z_산업일별예측] '다음 거래일 예측' 행은 v0.2 예정(헤더 CHANGELOG (c)) — 실적행만 낸다."""
    df = pd.DataFrame(index=eval_idx)
    df["날짜"] = eval_idx.date
    n_up = pd.Series(0, index=eval_idx)
    n_down = pd.Series(0, index=eval_idx)
    for t, r in results.items():
        st = r["state"].reindex(eval_idx)
        tp = r["target_pos"].reindex(eval_idx)
        df[f"{t} 예측"] = st.map(lambda x: S.STATE_SHORT.get(x, "-") if pd.notna(x) else "-")
        df[f"{t} 목표비중"] = tp.round(4)
        n_up = n_up.add((st == "RISK_ON").astype(int), fill_value=0)
        n_down = n_down.add((st == "RISK_OFF").astype(int), fill_value=0)
    df.insert(1, "상승예측 산업수", n_up.astype(int).values)
    df.insert(2, "하락예측 산업수", n_down.astype(int).values)
    return df.reset_index(drop=True)


def build_industry_summary(results: Dict[str, Dict[str, Any]], failed: Dict[str, str],
                           active_table: List[Tuple[str, str, str]], icfg: IndustryConfig) -> pd.DataFrame:
    rows = []
    for ind, parent, ff49 in active_table:
        if ind in results:
            r = results[ind]
            rows.append({"티커": ind, "산업명": INDUSTRY_NAME_KR.get(ind, ind), "부모섹터": parent,
                        "FF49매핑": ff49, "첫유효일": r["info"].get("실제데이터시작"),
                        "신호시작일": r.get("first_signal"), "후보지표수": r["n_candidates"],
                        "채택지표수": len(r["adopted"]), "vol_scale": round(float(r["vol_scale"]), 3),
                        "위험소스": r.get("hazard_source"), "상태": "성공"})
        else:
            rows.append({"티커": ind, "산업명": INDUSTRY_NAME_KR.get(ind, ind), "부모섹터": parent,
                        "FF49매핑": ff49, "상태": f"실패: {failed.get(ind, '데이터없음')[:100]}"})
    return pd.DataFrame(rows)


# =============================================================================
# [7] 최상위 오케스트레이터 — run(sres, res, M, S, icfg)
# =============================================================================
def run(sres: dict, res: dict, M, S, icfg: Optional[IndustryConfig] = None,
       industry_px_override: Optional[Dict[str, pd.DataFrame]] = None,
       parent_px_override: Optional[Dict[str, pd.DataFrame]] = None) -> Dict[str, Any]:
    t0 = time.time()
    icfg = icfg or CFG

    if sres.get("aborted"):
        log("START", kv(event="s_aborted"), M=M, level="error")
        return {"aborted": True, "industries": {}, "failed": {}, "icfg": icfg,
                "note": "S(sector_rotation) 결과가 aborted — 산업 계층을 실행할 수 없습니다."}

    table = active_industries(icfg, sres)
    industry_tickers = [t for t, _, _ in table]
    parent_of = {t: p for t, p, _ in table}
    active_parents = sorted(set(parent_of.values()))
    log("START", kv(event="industry_universe_ready", active=len(industry_tickers),
                    parents=len(active_parents), excluded_optional=(0 if icfg.USE_TIER_B else len(INDUSTRIES_OPTIONAL))), M=M)

    if icfg.RUN_SELFTEST:
        st = S.run_selftest(M, icfg)
        if not st.get("passed"):
            log("START", kv(event="selftest_fail", **{k: v for k, v in st.items() if k != "passed"}), M=M, level="error")
            return {"selftest": st, "aborted": True, "industries": {}, "failed": {}, "icfg": icfg}
    else:
        st = {"passed": None}

    quality: List[dict] = []
    _ensure_yahoo_expected_start_industries(M)
    cal = res["cal"]
    ind_px, _ = S.fetch_sector_prices(res, M, quality, sector_px_override=industry_px_override,
                                      tickers=tuple(industry_tickers))
    par_px, _ = S.fetch_sector_prices(res, M, quality, sector_px_override=parent_px_override,
                                      tickers=tuple(active_parents))

    spy_df = res["px_dict"]["SPY"]
    spy_df = spy_df[~spy_df.index.duplicated(keep="last")].sort_index()
    spy_tr, _ = S.build_total_return_close(spy_df, cal, icfg.ADJ_CLOSE_STALE_DAYS)
    spy_raw = spy_df["Close"].reindex(cal).ffill()
    spy_series = S.spy_layer_series(res, M)

    rf_series = res.get("fred", {}).get("DGS3MO")
    rf_daily = None
    if rf_series is not None and rf_series.notna().sum() > 100:
        rf_daily = (rf_series / 100.0 / 252.0).reindex(cal).ffill()

    frames: Dict[str, Any] = {}
    raw: Dict[str, Any] = {}
    universe_rows: List[dict] = []
    for t in industry_tickers:
        parent = parent_of[t]
        df = ind_px.get(t)
        if df is None or len(df) == 0:
            universe_rows.append({"티커": t, "산업명": INDUSTRY_NAME_KR.get(t, t), "부모섹터": parent,
                                  "상태": "제외(데이터없음)"})
            continue
        price_i, idx_i, info = S.sector_price_frame(df, cal, icfg)
        if len(idx_i) < 252:
            universe_rows.append({"티커": t, "산업명": INDUSTRY_NAME_KR.get(t, t), "부모섹터": parent,
                                  "상태": "제외(이력부족<1년)", **info})
            continue
        frames[t] = (price_i, idx_i, info)
        raw[t] = df
        universe_rows.append({"티커": t, "산업명": INDUSTRY_NAME_KR.get(t, t), "부모섹터": parent,
                              "상태": "포함", **info})

    parent_tr: Dict[str, pd.Series] = {}
    parent_raw: Dict[str, pd.Series] = {}
    for p in active_parents:
        df = par_px.get(p)
        if df is None or len(df) == 0:
            log("DATA", kv(event="parent_price_missing", parent=p), M=M, level="error")
            continue
        tr, _ = S.build_total_return_close(df, cal, icfg.ADJ_CLOSE_STALE_DAYS)
        parent_tr[p] = tr
        parent_raw[p] = df["Close"].reindex(cal).ffill()
    active_parents = [p for p in active_parents if p in parent_tr]
    frames = {t: v for t, v in frames.items() if parent_of[t] in parent_tr}

    parent_series = {p: parent_layer_series(sres, p, M, mode=icfg.PARENT_LAYER_SOURCE) for p in active_parents}

    industry_breadth_200 = None
    if icfg.USE_INDUSTRY_BREADTH and frames:
        flags = []
        for t, (price_i, idx_i, info) in frames.items():
            trend = S.sector_technical_values(price_i["Adj Close"].astype(float))["TREND_200"]
            flags.append((trend > 0).reindex(cal))
        industry_breadth_200 = pd.concat(flags, axis=1).mean(axis=1)
    parent_breadth_200 = None
    if icfg.USE_INDUSTRY_BREADTH and parent_tr:
        flags = [(S.sector_technical_values(tr)["TREND_200"] > 0).reindex(cal) for tr in parent_tr.values()]
        parent_breadth_200 = pd.concat(flags, axis=1).mean(axis=1)

    ctx = {"M": M, "res": res, "sres": sres, "S": S, "icfg": icfg, "frames": frames, "raw": raw,
          "parent_of": parent_of, "parent_tr": parent_tr, "parent_raw": parent_raw,
          "spy_tr": spy_tr, "spy_raw": spy_raw, "spy_series": spy_series, "parent_series": parent_series,
          "rf_daily": rf_daily, "industry_breadth_200": industry_breadth_200,
          "parent_breadth_200": parent_breadth_200}

    results, failed = run_industries(list(frames.keys()), ctx, icfg, M)
    log("PIPE", kv(event="industries_done", ok=len(results), failed=len(failed)), M=M)

    eval_idx = cal[cal >= pd.Timestamp(res["cfg"].SIGNAL_START)]
    wf: Dict[str, Any] = {}
    alloc: Dict[str, Any] = {}
    accept_df = pd.DataFrame()
    hier_df = pd.DataFrame()
    attrib_df = pd.DataFrame()
    if icfg.USE_ROTATION and results:
        try:
            wf = build_pooled_rotation(results, eval_idx, icfg, M, S)
            alloc = build_industry_allocation(results, sres, res, eval_idx, icfg, M, S, wf, rf_daily=rf_daily)
            if alloc:
                accept_df = build_industry_acceptance(alloc, wf, icfg, M, S)
                hier_df = build_hierarchy_check(alloc)
                attrib_df = build_industry_vs_sector_attribution(alloc)
                log("ACCEPT", kv(event="acceptance", **{f"c{i+1}": accept_df.iloc[i]["판정"] for i in range(len(accept_df))}), M=M)
                _viol = int(hier_df.iloc[-1]["위반일수(>1e-9)"]) if len(hier_df) else -1
                log("HIER", kv(event="hierarchy_check", violations=_viol), M=M,
                    level=("warning" if _viol else "info"))
        except Exception as e:
            log("START", kv(event="allocation_failed", err=str(e)[:200]), M=M, level="error")
            wf, alloc = {}, {}

    universe = pd.DataFrame(universe_rows)
    matrix = build_industry_prediction_matrix(results, eval_idx, icfg, S)
    summary = build_industry_summary(results, failed, table, icfg)
    stage_timing = {"00_전체": round(time.time() - t0, 1)}

    return {
        "industries": results, "failed": failed, "selftest": st, "universe": universe,
        "quality": pd.DataFrame(quality), "matrix": matrix, "summary": summary,
        "wf": wf, "alloc": alloc, "acceptance": accept_df, "hierarchy": hier_df,
        "attribution": attrib_df, "icfg": icfg,
        "signal_start": (str(eval_idx[0].date()) if len(eval_idx) else "-"),
        "cal_end": str(cal[-1].date()), "aborted": False, "stage_timing": stage_timing,
        "active_table": table,
    }


# =============================================================================
# [8] 리포트
# =============================================================================
def build_industry_report(ires: Dict[str, Any], M=None, S=None, path: Optional[str] = None) -> str:
    icfg = ires.get("icfg", CFG)
    path = path or icfg.OUT_XLSX
    if ires.get("aborted"):
        meta = [("버전", f"industry_rotation.py {VERSION} ({VERSION_DATE})"),
               ("판정", ires.get("note", "자기검사 실패 — 산업 계층 리포트를 낼 수 없습니다.")),
               ("면책", "본 산출물은 연구·교육 목적의 백테스트 결과이며 투자 자문이 아닙니다.")]
        S.write_sector_excel(path, {}, meta, M=M)
        return path

    results = ires["industries"]
    alloc = ires.get("alloc", {})
    accept_df = ires.get("acceptance", pd.DataFrame())
    passed = bool(len(accept_df) and (accept_df["판정"] == "PASS").all())

    sheets: Dict[str, pd.DataFrame] = {}
    sheets["01Z_산업일별예측"] = ires["matrix"]
    sheets["12_산업요약"] = ires["summary"]
    if alloc:
        sheets["13_산업배분전략"] = alloc["perf"]
        sheets["13b_배분전략자산곡선"] = alloc["curve"]
        tw = alloc["target_w"].round(4).copy()
        tw.insert(0, "날짜", alloc["target_w"].index.date)
        sheets["13c_일별배분비중"] = tw.reset_index(drop=True)
        sheets["13f_산업수용기준"] = accept_df
        sheets["14_계층정합"] = ires["hierarchy"]
        sheets["15_산업대섹터귀속"] = ires["attribution"]
    for t, r in results.items():
        sheets[f"01_일별_{t}"] = r["sheets"].get("daily", pd.DataFrame())
    val_frames = [r["sheets"]["val_sheet"] for r in results.values() if r["sheets"].get("val_sheet") is not None
                 and len(r["sheets"]["val_sheet"])]
    if val_frames:
        sheets["03_지표검증"] = pd.concat(val_frames, ignore_index=True, sort=False)
    q, u = ires.get("quality", pd.DataFrame()), ires.get("universe", pd.DataFrame())
    sheets["10_데이터품질"] = pd.concat([u, q], ignore_index=True, sort=False) if len(q) else u

    verdict = ("산업 계층이 S★를 이긴다(수용기준 ①~⑤ 전부 PASS)" if passed else
              "산업 계층은 S★를 이기지 못함 — 운용은 S★ 그대로(설계서 §8 관행, 진단용으로만 유지)")
    n_fail_ind = len(ires.get("failed", {}))
    meta = [
        ("버전", f"industry_rotation.py {VERSION} ({VERSION_DATE}) — sector_rotation.py {getattr(S, 'VERSION', '?')} — "
                f"market_regime_trader.py {getattr(M, 'BUNDLE_VERSION', '?')}"),
        ("판정", verdict),
        ("예측 대상", f"{len(results)}개 산업 ETF(부모섹터 하위) — 부모는 S.SECTOR_EXCLUDE(XLB·XLE 등) 제외 후 산업ETF가 있는 섹터만"),
        ("실패 산업", f"{n_fail_ind}개" if n_fail_ind else "없음"),
        ("신호/백테스트 기간", f"{ires.get('signal_start')} ~ {ires.get('cal_end')}"),
        ("체결 규칙", "t일 종가에 신호 확정 → t+1일 시가 체결(M·S와 동일, 룩어헤드 구조적 차단)"),
        ("거래비용", f"산업 ETF 편도 {icfg.COST_BPS_INDUSTRY:.0f}bp · 부모 ETF 편도 {icfg.PARENT_COST_BPS:.0f}bp"),
        ("총 노출 불변식", "Σ산업비중 + 부모ETF비중 = S★의 그 섹터비중 — 14_계층정합 시트가 매일 이 등식을 검사(위반 0일이어야 함)"),
        ("v0.1.0 범위(⚠ 명시적 축소)",
         "MAX_WORKERS 병렬 실행·06c 임계값민감도·11 룩어헤드감사 시트·01Y 예측정확도·13d~13o(13/13b/13c 제외)·"
         "13h FF49외부검증(네트워크 필요)·회복플로어/깊은낙폭/구조적바닥 규칙(M 규칙⑧⑨⑬)·'다음 거래일 예측' 행은 "
         "v0.2 예정(파일 헤더 CHANGELOG (a)(b)(c) 참조). 풀링 순환매의 ③④ 수용기준은 상위1-하위1 스프레드로 "
         "근사한다(top3-bottom3 헬퍼가 S에 별도 노출되어 있지 않음). 룩어헤드 안전성은 이 리포트와 별개로 "
         "test_industry_lookahead_cut.py(절단재계산 회귀)가 항상 독립적으로 검증한다."),
        ("면책", "본 산출물은 연구·교육 목적의 백테스트 결과이며 투자 자문이 아닙니다. 과거 성과는 미래 수익을 보장하지 않습니다."),
    ]
    for k, v in ires.get("stage_timing", {}).items():
        meta.append((f"실행시간 - {k}", f"{v:.1f}초"))

    S.write_sector_excel(path, sheets, meta, M=M)
    if icfg.EXPORT_DAILY_CSV:
        try:
            ires["matrix"].to_csv(icfg.DAILY_CSV_PATH, index=False, encoding="utf-8-sig")
            if alloc:
                sheets["13c_일별배분비중"].to_csv(icfg.ALLOC_CSV_PATH, index=False, encoding="utf-8-sig")
        except Exception as e:
            log("REPORT", kv(event="csv_export_failed", err=str(e)[:150]), M=M, level="warning")
    return path


def main(sres: dict, res: dict, M, S, icfg: Optional[IndustryConfig] = None,
        industry_px_override: Optional[Dict[str, pd.DataFrame]] = None,
        parent_px_override: Optional[Dict[str, pd.DataFrame]] = None) -> str:
    """S.main()과 같은 역할: run -> build_industry_report -> Colab이면 자동 다운로드.
    [v0.1.3] 이전 버전까지는 리포트 경로만 반환하고 다운로드를 하지 않았다(S.main()/M.main()과
    달리 이 단계가 빠져 있었음 — 사용자 지시로 추가). S.main()과 동일하게 엑셀+CSV 2개를
    M.maybe_colab_download_many()로 zip 1개에 묶어 한 번에 내려받는다(브라우저가 같은 세션에서
    자동 다운로드가 여러 번 이어지면 첫 번째 이후를 조용히 차단하는 문제를 피하기 위함 — S가
    v0.9.2에서 이미 겪고 고친 문제라 같은 해법을 그대로 재사용한다). M이 구버전이라 그 함수가
    없으면 파일별로 M.maybe_colab_download()를 개별 호출하는 것으로 자동 대체한다."""
    icfg = icfg or CFG
    ires = run(sres, res, M, S, icfg, industry_px_override=industry_px_override,
              parent_px_override=parent_px_override)
    out = build_industry_report(ires, M=M, S=S)
    dl_paths = [out]
    if icfg.EXPORT_DAILY_CSV and os.path.exists(icfg.DAILY_CSV_PATH) and not ires.get("aborted"):
        dl_paths.append(icfg.DAILY_CSV_PATH)
    if icfg.EXPORT_DAILY_CSV and os.path.exists(icfg.ALLOC_CSV_PATH) and not ires.get("aborted") and ires.get("alloc"):
        dl_paths.append(icfg.ALLOC_CSV_PATH)
    if hasattr(M, "maybe_colab_download_many"):
        M.maybe_colab_download_many(dl_paths, zip_name=os.path.splitext(icfg.OUT_XLSX)[0] + "_bundle.zip")
    elif hasattr(M, "maybe_colab_download"):
        for p in dl_paths:
            M.maybe_colab_download(p)
    return out


# =============================================================================
# [9] 합성데이터(테스트용, 네트워크 불필요) — 부모 총수익(S가 이미 계산해 둔 bh_ret)을 시장요인으로
#     쓰는 단일요인 모델. §11 test_industry_e2e_synth.py가 이걸로 파이프라인 전체를 배관검사한다.
# =============================================================================
def make_synthetic_industry_prices(sres: dict, res: dict, icfg: Optional[IndustryConfig] = None,
                                   seed: int = 20260905) -> Dict[str, pd.DataFrame]:
    icfg = icfg or CFG
    rng = np.random.default_rng(seed)
    cal = res["cal"]
    out: Dict[str, pd.DataFrame] = {}
    for ind, parent, _ in active_industries(icfg, sres):
        pr = sres["sectors"].get(parent, {})
        parent_bh = pr.get("bh_ret")
        if parent_bh is None or len(parent_bh) == 0:
            continue
        parent_ret_full = parent_bh.reindex(cal).fillna(0.0)
        start = pd.Timestamp(INDUSTRY_EXPECTED_START.get(ind, "2005-01-01"))
        idx = cal[cal >= start]
        if len(idx) < 300:
            continue
        beta = rng.uniform(0.8, 1.2)
        idio = rng.normal(0.0, 0.008, len(idx))
        r = beta * parent_ret_full.reindex(idx).values + idio
        close = 50.0 * np.exp(np.cumsum(r))
        open_ = close * (1.0 + rng.normal(0.0, 0.002, len(idx)))
        high = np.maximum(open_, close) * (1.0 + np.abs(rng.normal(0.0, 0.003, len(idx))))
        low = np.minimum(open_, close) * (1.0 - np.abs(rng.normal(0.0, 0.003, len(idx))))
        div_factor = np.exp(np.cumsum(np.full(len(idx), 0.00006)))
        out[ind] = pd.DataFrame({"Open": open_, "High": high, "Low": low, "Close": close,
                                 "Adj Close": close * div_factor / div_factor[-1], "Volume": 1_000_000.0}, index=idx)
    return out


if __name__ == "__main__":
    print(f"industry_rotation.py {VERSION} ({VERSION_DATE}) — 산업 계층 순환매/배분. "
         "run(sres, res, M, S, icfg) 또는 main(sres, res, M, S, icfg)로 사용. 연구/교육용, 투자 자문 아님.")
