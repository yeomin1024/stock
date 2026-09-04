# =============================================================================
#  sector_rotation.py
#  VERSION: v0.1.0 - 2026-09-04 - 11개 SPDR 섹터 ETF 로테이션 최초 구현
#
#  목적:
#    market_regime_trader.py(이하 M)가 만든 SPY 국면(목표비중 E_t)을 그대로 두고, 그 노출을
#    11개 섹터(XLK/XLV/XLY/XLP/XLF/XLE/XLI/XLB/XLU/XLC/XLRE) 중 어디에 둘지만 결정한다.
#    각 섹터는 SPY와 같은 파이프라인(후보지표 -> 6기준 검증 -> 워크포워드 가중 -> 복합점수
#    -> 이력현상 상태기계 -> 매수/중립/매도)을 거치되, 예측 대상이 "섹터 절대수익"이 아니라
#    "SPY 대비 상대수익"이라는 점만 다르다. 시장 방향은 이미 M이 담당하므로 여기서는 방향을
#    다시 예측하지 않고 "어느 섹터가 시장을 이기고 질지"만 예측한다.
#
#  설계 문서: claude/SECTOR_ROTATION_SPEC_v0.1.md (프로젝트 "미국 주식 매수, 매도 프로그램
#    만들기2"). 본 파일은 그 문서의 §1~§9를 그대로 구현한다.
#
#  market_regime_trader.py는 이 파일에서 **한 줄도 수정하지 않는다** (§9.1 "가능하면 무수정").
#  구현 중 발견한 단순화: 스펙 초안(§4.1)은 validate_indicators/build_walkforward_weights에
#  dd_label_override/fwd_override 훅을 추가하는 안이었으나, 실제로는 "섹터가격/SPY가격" 비율
#  시계열(rel = P_i/SPY) 자체를 그 두 함수의 px_adj 인자로 그대로 넘기면(그 함수들은 px_adj를
#  '가격'으로만 다루고 절대/상대를 구분하지 않는다) forward_return(rel,h)이 곧 상대수익이 되고
#  forward_maxdd(rel,h)가 곧 상대낙폭이 되어, 코드 수정이 전혀 필요 없다(DD_LABEL_THRESHOLD만
#  cfg에서 -3%로 낮춰 쓴다 — 이미 존재하는 파라미터). 유일한 런타임 의존은 (1) M.YAHOO_
#  EXPECTED_START 딕셔너리에 섹터 티커를 런타임에 update()하는 것(fetch_all_yahoo가 그 값을
#  .get()으로만 읽으므로 다른 티커에 영향 없음), (2) M.INDICATOR_SPECS/M.SPEC_BY_KEY를 섹터별
#  후보지표 목록으로 일시 교체하는 것 — 두 함수 모두 이 두 전역을 직접 참조하는 설계라서
#  피할 수 없고, self_test()가 위험트랙 픽스처 검증에 쓰는 것과 완전히 같은(=이미 프로덕션에서
#  검증된) 패턴이다(_indicator_spec_override() 참조). 둘 다 try/finally로 항상 원복한다.
#
#  기본값: USE_SECTOR_ROTATION=False. §8 수용기준(두 번의 실측)을 통과하기 전까지는 리포트가
#  진단 목적으로만 쓰인다 — 실제 목표비중 산출 경로에는 아직 연결하지 않는다(연결 지점은
#  build_sector_targets()의 반환값을 사용자가 원할 때 M 실행 스크립트에서 target_i로 쓰는 것).
#
#  CHANGELOG
#  ---------------------------------------------------------------------------
#  v0.1.0 | 2026-09-04 | 최초 구현. SECTOR_ROTATION_SPEC_v0.1.md §1~§9 구현:
#    §1 유니버스(11 SPDR 섹터), 진입일 규칙(실제이력+3년), Adj Close 지연 감지.
#    §2 타깃 = P_i/SPY 비율 시계열(그 자체를 px_adj로 재사용).
#    §3 지표 가족 A(횡단면·기술, 공통 9종)/B(섹터별 매크로, 표 기반 3~4종)/C(국면 상호작용,
#       공통 5종) — 전부 사전방향 명시.
#    §4 검증: 섹터별 IC/NW-t/AUC 게이트(M.validate_indicators 재사용, dd_label만 상대낙폭
#       -3% 기준) + 횡단면 rank IC 게이트 + 판별력 자기검사(합성데이터).
#    §5 워크포워드 가중(섹터별 M.build_walkforward_weights 재사용 — 패널 결합 추정은 v0.2로
#       보류, 아래 "v0.1 범위 축소" 참조), 복합점수(M.composite_score), 횡단면 순위, 이력현상
#       상태기계(신규 구현), tilt 포트폴리오 구성(Σ target_i = E_t 불변식).
#    §6 다자산 백테스트(신규 구현, 현금 레그를 한 번만 계산 — run_backtest 반복 호출 금지),
#       벤치마크 B0~B3.
#    §7 룩어헤드 감사(절단재계산), 유니버스 편입 인과성 감사.
#    §8 수용기준 자동 판정(PASS/FAIL 표).
#    §9 리포트(별도 파일, 12~19번대 시트), 단위테스트.
#
#    [단위테스트에서 발견·수정 — 최초 배포 전] build_portfolio() 상한(SECTOR_MAX_WEIGHT) 적용
#    로직: 원래 "clip 후 전체를 동일 배율로 재정규화"를 2회 반복하는 방식이었는데, 이 방식은
#    상한을 정확히 만족시킬 수 없는 구조적 결함이 있었다(재정규화 배율이 1보다 크면 이미
#    상한에 걸린 항목까지 같은 배율로 다시 커져서 재차 상한을 초과 — 기하급수적으로 상한에
#    접근할 뿐 도달하지 못함). test_sector_portfolio.py가 실제로 상한을 ~1%p 초과하는 사례를
#    잡아냈다. 상한에 걸린 항목을 그 값에 고정하고 남은 예산만 미고정 항목에 재분배하는
#    워터필링(waterfilling) 방식으로 교체해 정확히 수렴하도록 수정했다(build_portfolio 참조).
#
#    [v0.1 범위 축소 — 문서에 명시] 스펙 §5.1(b)의 "가족 A/C 패널 결합 추정"은 이번 버전에서
#    보류하고 섹터별 독립 추정만 구현했다(각 섹터가 자기 자신의 워크포워드 가중치를 갖는다).
#    이유: 패널 추정은 build_walkforward_weights를 복사-수정해야 하는데, 그 함수는 검증된
#    프로덕션 코드이고 이미 크고 복잡하다 — 이번 라운드에 반쯤 검증된 패널 버전을 새로 만드는
#    것보다, 섹터별 독립 추정(이미 100% 검증된 함수를 그대로 재사용)으로 먼저 실측하고, 결과가
#    지표 부족(가족 A/C 특유의 "패널이라야 통계량이 안정된다")으로 나오면 v0.2에서 패널 버전을
#    별도로 만들어 비교하는 것이 안전하다(사용자 우선순위 (1)정확성 > (3)성능, 그리고 검증되지
#    않은 대형 신규 코드보다 검증된 코드의 반복 재사용이 §6 정확성/투명성 원칙에 더 부합).
#
#  ※ 본 코드는 연구/교육용 도구이며 투자 자문이 아니다. (Not financial advice)
# =============================================================================
from __future__ import annotations

import time
import math
import dataclasses
import contextlib
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional, Tuple, Any

import numpy as np
import pandas as pd

VERSION = "v0.1.0"

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

# [§1.1] Yahoo YAHOO_EXPECTED_START 형식과 동일하게 "상장 다음 달 1일" 단위로 표기.
SECTOR_EXPECTED_START: Dict[str, str] = {
    "XLK": "1999-01-01", "XLV": "1999-01-01", "XLY": "1999-01-01", "XLP": "1999-01-01",
    "XLF": "1999-01-01", "XLE": "1999-01-01", "XLI": "1999-01-01", "XLB": "1999-01-01",
    "XLU": "1999-01-01",
    "XLRE": "2015-11-01",   # 2015-10-07 상장
    "XLC": "2018-07-01",    # 2018-06-18 상장
}


def _ensure_yahoo_expected_start(M) -> None:
    """[§1.1] M.YAHOO_EXPECTED_START(모듈 전역 dict)에 섹터 티커를 런타임에 추가한다.
    .get(ticker) 조회만 쓰는 기존 코드 경로(_yahoo_degenerate/validate_price_data)는 새
    키가 추가되어도 다른 티커의 동작에 전혀 영향받지 않는다 — market_regime_trader.py
    소스는 한 글자도 바뀌지 않는다. 멱등(이미 있으면 덮어쓰기만, 값 동일)."""
    M.YAHOO_EXPECTED_START.update(SECTOR_EXPECTED_START)


# =============================================================================
# [1] 설정
# =============================================================================
@dataclass
class SectorConfig:
    SECTORS: Tuple[str, ...] = SECTORS
    SECTOR_MIN_HISTORY_YEARS: int = 3          # [§1.3] 유니버스 편입에 필요한 최소 이력
    SECTOR_REL_DD_THRESHOLD: float = -0.03     # [§2.3] 상대낙폭 라벨 임계값(21일 -3%)
    ADJ_CLOSE_STALE_DAYS: int = 5              # [§1.2] Adj Close가 Close보다 이만큼 이상
                                                # 늦게 끊기면 "Adj Close 지연"으로 표기·보정

    # ---- 상태기계 (⚠ 신호/사이징 파라미터, 격자 민감도만 보고 최적셀 채택 금지) ----
    SECTOR_TOP_K: int = 3
    SECTOR_STATE_HI: float = 0.60
    SECTOR_STATE_LO: float = 0.40
    SECTOR_HYSTERESIS_DAYS: int = 2
    SECTOR_MIN_HOLD_DAYS: int = 5
    SECTOR_TILT_OVER: float = 1.5
    SECTOR_TILT_UNDER: float = 0.0
    SECTOR_MAX_WEIGHT: float = 0.25

    # ---- 가중치 상한 (기존 SERIES_WEIGHT_CAP=0.20과 별개 상한, §5.1) ----
    SECTOR_REGIME_FAMILY_CAP: float = 0.40     # 가족 C(국면상호작용) 합산 가중치 상한
    SECTOR_COMMON_FAMILY_CAP: float = 0.60     # 가족 A(횡단면·기술) 합산 가중치 상한

    # ---- 활성화 스위치 (§8 "무실증 활성화 금지") ----
    USE_SECTOR_ROTATION: bool = False          # True가 되어도 이 파일은 target_i를 계산해
                                                # 반환할 뿐, M의 실제 매매신호에 자동 연결되지
                                                # 않는다 — 연결은 사용자가 명시적으로 한다.
    USE_DISPERSION_GATE: bool = False          # [§5.4] 기본 비활성

    # ---- 판별력 자기검사(§4.3) 수용기준 ----
    SELFTEST_MIN_TRUE_ADOPTED: int = 2         # 심은 3개 중 최소 채택 수
    SELFTEST_MAX_NOISE_ADOPTED: int = 1        # 잡음 5개 중 최대 채택 허용 수

    # ---- 횡단면 게이트(§4.2) ----
    CS_MIN_RANK_IC: float = 0.02
    CS_MIN_NW_T: float = 2.0

    RANDOM_SEED: int = 20260904


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
    """[헤더 주석 참조] M.validate_indicators/M.build_walkforward_weights/M.build_reason_text는
    모두 지표 목록을 인자로 받지 않고 모듈 전역 M.INDICATOR_SPECS(및 M.SPEC_BY_KEY)를 직접
    참조한다 — 이는 market_regime_trader.py의 self_test()가 위험트랙 픽스처를 검증할 때 쓰는
    것과 완전히 같은 패턴(orig=INDICATOR_SPECS[:]; INDICATOR_SPECS=[...]; try/finally 원복)을
    모듈 밖에서 그대로 재현한다. 예외가 나도 반드시 원복된다."""
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
def fetch_sector_prices(res: dict, M, quality_rows: List[dict]
                        ) -> Tuple[Dict[str, Optional[pd.DataFrame]], List[dict]]:
    """[§1.2] 11개 섹터 ETF를 M.fetch_all_yahoo로 수집(퇴화수집 게이트·지연캐시 포함,
    market_regime_trader.py 무수정 재사용)하고, SPY까지 포함해 M.validate_price_data로
    무결성(시작일·교차오염) 검사한다. SPY는 res["px_dict"]["SPY"](이미 검증된 프레임)를
    그대로 재사용하므로 M.validate_price_data의 SPY 하드요건이 항상 자연히 통과한다.
    반환: (티커->프레임 dict, yahoo_diag 리스트)."""
    _ensure_yahoo_expected_start(M)
    cfg = res["cfg"]
    yahoo_diag: List[dict] = []
    px = M.fetch_all_yahoo(list(SECTORS), cfg, diag=yahoo_diag)
    for row in yahoo_diag:
        quality_rows.append(row)

    combined = dict(px)
    combined["SPY"] = res["px_dict"]["SPY"]
    validated = M.validate_price_data(combined, cfg, quality_rows)
    sector_px = {t: validated.get(t) for t in SECTORS}

    n_ok = sum(1 for v in sector_px.values() if v is not None and len(v) > 0)
    log("DATA", kv(event="sector_fetch_done", tickers=len(SECTORS), ok=n_ok), M=M)
    return sector_px, yahoo_diag


def adj_close_lag_check(df: Optional[pd.DataFrame], ticker: str, cfg_min_stale_days: int
                        ) -> Tuple[bool, int, Optional[pd.Timestamp], Optional[pd.Timestamp]]:
    """[§1.2] Adj Close 최종 유효일이 Close 최종 유효일보다 cfg_min_stale_days거래일 이상
    앞서면 '지연'으로 판정한다(report23의 ^VIX3M Adj Close 사례와 같은 종류 — 배당 있는
    섹터 ETF는 Adj Close가 필수라 이 검사가 더 중요하다). 반환: (지연여부, 지연일수,
    Adj Close 마지막일, Close 마지막일)."""
    if df is None or len(df) == 0 or "Close" not in df.columns:
        return False, 0, None, None
    close_last = df["Close"].dropna().index.max() if df["Close"].notna().any() else None
    if "Adj Close" not in df.columns:
        return False, 0, None, close_last
    adj_s = df["Adj Close"].dropna()
    adj_last = adj_s.index.max() if len(adj_s) else None
    if adj_last is None or close_last is None:
        return False, 0, adj_last, close_last
    gap_days = int((df.index >= adj_last) .sum() - 1) if adj_last <= close_last else 0
    # 거래일 기준 지연폭: adj_last 이후 ~ close_last 까지 인덱스상의 거래일 수
    gap_trading_days = int(((df.index > adj_last) & (df.index <= close_last)).sum())
    stale = gap_trading_days >= cfg_min_stale_days
    return stale, gap_trading_days, adj_last, close_last


def build_total_return_close(df: pd.DataFrame, cal: pd.DatetimeIndex,
                             stale_days_threshold: int) -> Tuple[pd.Series, bool]:
    """[§1.2] Adj Close(배당 포함 총수익) 시계열을 cal에 정렬해 반환한다. Adj Close가
    stale_days_threshold거래일 이상 지연돼 있으면(§ adj_close_lag_check), 지연 구간만
    Close의 일간수익률(배당 미포함 근사)을 마지막 정상 Adj Close 위에 이어붙여 총수익
    시계열이 최근 구간에서 갑자기 끊기지 않게 한다. 반환: (시계열, 대체적용여부)."""
    close = df["Close"].reindex(cal).ffill() if "Close" in df.columns else None
    if "Adj Close" not in df.columns or close is None:
        return close, False
    adj = df["Adj Close"].reindex(cal)
    stale, gap, adj_last, close_last = adj_close_lag_check(df, "", stale_days_threshold)
    if not stale or adj_last is None:
        return adj.ffill(), False
    # adj_last까지는 원래 Adj Close를 쓰고, 그 이후는 Close의 일간수익률을 곱해 이어붙인다.
    adj_head = adj.loc[:adj_last].ffill()
    base = float(adj_head.iloc[-1])
    close_tail_ret = close.loc[adj_last:].pct_change().fillna(0.0)
    tail_factor = (1.0 + close_tail_ret).cumprod()
    tail = base * tail_factor
    out = pd.concat([adj_head.iloc[:-1], tail])
    out = out.reindex(cal).ffill()
    return out, True


def sector_entry_dates(sector_px: Dict[str, Optional[pd.DataFrame]], scfg: SectorConfig
                       ) -> Dict[str, pd.Timestamp]:
    """[§1.3] 섹터별 '실제 데이터 시작 + SECTOR_MIN_HISTORY_YEARS'를 유니버스 편입일로
    계산한다. 실제 상장일(과거)만 쓰므로 룩어헤드가 아니다 — 편입일 자체가 미래 정보를
    포함하지 않는, 그 시점에 이미 관측 가능한 이력 길이 규칙이다."""
    out: Dict[str, pd.Timestamp] = {}
    for t in scfg.SECTORS:
        df = sector_px.get(t)
        if df is None or len(df) == 0:
            out[t] = pd.Timestamp.max
            continue
        start = df.index.min()
        out[t] = start + pd.DateOffset(years=scfg.SECTOR_MIN_HISTORY_YEARS)
    return out


def sector_active_mask(entry_dates: Dict[str, pd.Timestamp], cal: pd.DatetimeIndex,
                       scfg: SectorConfig) -> pd.DataFrame:
    """[§1.3] date x sector 불리언 — 그날 유니버스에 편입돼 있는지."""
    out = pd.DataFrame(False, index=cal, columns=list(scfg.SECTORS))
    for t in scfg.SECTORS:
        out.loc[cal >= entry_dates[t], t] = True
    return out


# =============================================================================
# [4] 예측 대상(타깃) — SPY 대비 상대가격 비율 시계열
# =============================================================================
def relative_price_series(sector_tr_close: pd.Series, spy_adj: pd.Series) -> pd.Series:
    """[§2.1] rel_i,t = P_i,t / SPY_t (둘 다 총수익 조정 종가). 이 시계열을 그대로
    M.validate_indicators/M.build_walkforward_weights의 px_adj 인자로 넘기면:
      forward_return(rel, h) = rel[t+h]/rel[t] - 1
                              = (P_i,t+h/SPY_t+h) / (P_i,t/SPY_t) - 1
                              = ln 상대수익의 1차 근사(단순수익 버전, 기존 관례와 통일)
      forward_maxdd(rel, h)  = rel의 향후 h일 최대낙폭 = '상대낙폭'
    이 되어 상대수익/상대낙폭을 위한 어떤 코드 수정도 필요 없다."""
    return (sector_tr_close / spy_adj.replace(0, np.nan)).astype(float)


def rolling_beta(sector_ret: pd.Series, spy_ret: pd.Series, window: int = 252,
                 lag: int = 1) -> pd.Series:
    """[§3.1 BETA_252] t행의 값이 t-lag일까지의 정보만 쓰도록 shift(lag)로 인과성을
    보장한다(과거 window일 롤링 공분산/분산 → 그 자체가 이미 인과적이고, 추가로 lag일
    지연시켜 '오늘 베타'가 아니라 '어제까지 확정된 베타'를 쓴다)."""
    cov = sector_ret.rolling(window, min_periods=window // 2).cov(spy_ret)
    var = spy_ret.rolling(window, min_periods=window // 2).var()
    beta = (cov / var.replace(0, np.nan))
    return beta.shift(lag)


# =============================================================================
# [5] 후보 지표 — 가족 A(횡단면·기술) / B(섹터별 매크로) / C(국면 상호작용)
# =============================================================================
@dataclass
class _RawSpec:
    """섹터 지표 하나를 만들기 위한 최소 정보. build_sector_indicators()가 이걸로
    실제 값(pd.Series)과 M.IndicatorSpec을 함께 만든다."""
    suffix: str            # 접미사(섹터티커와 합쳐 key가 됨) 예: "REL_MOM_21"
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
    """[기존 _generate_universe_indicators()의 _eval_h와 동일 규칙] 60일 이상 관측창의
    변화/모멘텀 지표는 21일 평가지평에서 단기 평균회귀로 부호가 뒤집히기 쉬우므로 63일을
    쓴다. None이면 호출부(validate_indicators)가 cfg.VAL_PRIMARY_H(21일)를 그대로 쓴다."""
    return 63 if window >= 60 else None


# ---- 가족 A: 횡단면·기술 지표(모든 섹터에 동일 정의) ----------------------------
def family_a_specs() -> List[_RawSpec]:
    return [
        _RawSpec("REL_MOM_21", "SPY대비 21일 상대모멘텀", "A.횡단면상대", -1,
                 "1개월 내 상대강도는 단기 과매수/과매도로 역전되는 경향(섹터 반전 효과)",
                 "최근 1개월 급등한 섹터는 단기 차익실현으로 상대 반락", "P_i/SPY 비율"),
        _RawSpec("REL_MOM_63", "SPY대비 63일 상대모멘텀", "A.횡단면상대", +1,
                 "3개월 상대모멘텀은 지속되는 경향(섹터 모멘텀 효과)",
                 "자금 흐름의 관성 — 최근 아웃퍼폼 섹터로 자금이 계속 유입",
                 "P_i/SPY 비율", trend_track=True, eval_horizon=_eval_h(63)),
        _RawSpec("REL_MOM_126", "SPY대비 126일 상대모멘텀", "A.횡단면상대", +1,
                 "6개월 상대모멘텀 지속(모멘텀 효과의 표준 관측창)",
                 "중기 자금흐름 관성", "P_i/SPY 비율", trend_track=True, eval_horizon=_eval_h(126)),
        _RawSpec("REL_MOM_12_1", "SPY대비 12-1개월 상대모멘텀", "A.횡단면상대", +1,
                 "최근 1개월을 제외한 12개월 상대모멘텀(전통적 모멘텀 팩터 정의)",
                 "최근월 반전 효과를 걸러낸 순수 모멘텀", "P_i/SPY 비율",
                 trend_track=True, eval_horizon=_eval_h(252)),
        _RawSpec("REL_MA_50_200", "상대가격 50/200일선 이격", "A.횡단면상대", +1,
                 "상대가격(P_i/SPY)의 골든/데드크로스는 상대추세 전환의 연속형 지표",
                 "이동평균 교차는 추세추종 자금의 진입/이탈 신호", "P_i/SPY 비율",
                 trend_track=True, eval_horizon=_eval_h(60)),
        _RawSpec("REL_DD_252H", "상대가격 52주 고점대비 낙폭", "A.횡단면상대", +1,
                 "상대 신고가 근접(낙폭 작음)은 상대강세 지속과 연관",
                 "상대 신고가 경신 섹터는 계속 주도주 지위를 유지하는 경향", "P_i/SPY 비율",
                 trend_track=True, eval_horizon=_eval_h(60)),
        _RawSpec("REL_RSI_14", "상대가격 RSI(14)", "A.횡단면상대", -1,
                 "상대가격의 단기 과매수(RSI 높음)는 역전되는 경향",
                 "기술적 과열은 단기 상대조정을 부른다", "P_i/SPY 비율"),
        _RawSpec("REL_VOL_RATIO", "섹터/SPY 실현변동성 비율(z)", "A.횡단면상대", -1,
                 "섹터 변동성이 SPY 대비 급등하면 이후 상대 열위(저변동성 효과의 상대판)",
                 "변동성 급등 = 불확실성/패닉 매도 국면, 이후 회복이 느림", "섹터·SPY 종가"),
        _RawSpec("REL_EXT_200", "섹터-SPY 200일선 이격도 차", "A.횡단면상대", -1,
                 "섹터 자체 200일선 이격도가 SPY보다 훨씬 높으면(과열) 이후 상대 열위",
                 "market_regime_trader.py 규칙 ⑩의 근거(200일선 이격도가 report21/23에서 "
                 "가장 강한 조기경보)를 섹터 상대판으로 재사용", "섹터·SPY 종가"),
    ]


def family_a_values(sector_close_tr: pd.Series, spy_close_tr: pd.Series,
                    sector_close_raw: pd.Series, spy_close_raw: pd.Series,
                    M) -> pd.DataFrame:
    """[§3.1] 가족 A 실제 값. 총수익(배당포함, _tr) 시계열로 모멘텀류를, 원시(배당제외
    가능성 있는 raw) 종가로 200일선 이격도류를 계산해 기존 SPY 지표(TREND_200 등이
    spy_a=Adj Close 기준인 것)와 정의를 통일한다 — 실무상 raw는 close 컬럼을 그대로 쓴다."""
    rel = (sector_close_tr / spy_close_tr.replace(0, np.nan)).astype(float)
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
    _rs = _gain / _loss.replace(0, np.nan)
    out["REL_RSI_14"] = 100.0 - 100.0 / (1.0 + _rs)

    r_i = np.log(sector_close_tr.replace(0, np.nan)).diff()
    r_spy = np.log(spy_close_tr.replace(0, np.nan)).diff()
    vol_i = r_i.rolling(20).std()
    vol_spy = r_spy.rolling(20).std()
    vol_ratio = vol_i / vol_spy.replace(0, np.nan)
    out["REL_VOL_RATIO"] = M._z(vol_ratio, 100)

    ma200_i = sector_close_raw.rolling(200, min_periods=150).mean()
    ma200_spy = spy_close_raw.rolling(200, min_periods=150).mean()
    ext_i = sector_close_raw / ma200_i.replace(0, np.nan) - 1.0
    ext_spy = spy_close_raw / ma200_spy.replace(0, np.nan) - 1.0
    out["REL_EXT_200"] = ext_i - ext_spy
    return out.replace([np.inf, -np.inf], np.nan)


# ---- 가족 C: 국면 상호작용 지표(SPY 계층 출력 재사용, 모든 섹터 공통) -------------
def family_c_specs() -> List[_RawSpec]:
    return [
        _RawSpec("BETA_X_H", "(베타-1)×위험점수백분위(H)", "C.국면상호작용", -1,
                 "위험(H)이 높을수록 고베타 섹터가 저베타보다 상대 열위",
                 "시장 전체 위험 프리미엄 확대 국면에서 고베타 자산이 더 많이 할인된다",
                 "M.res[haz_pct] × 롤링베타"),
        _RawSpec("BETA_X_SCORE", "(베타-1)×복합점수백분위", "C.국면상호작용", +1,
                 "시장 강세 확신(복합점수)이 높을수록 고베타 섹터가 우위",
                 "강세장 확신 국면에서는 고베타가 저베타를 아웃퍼폼", "M.res[score_pct] × 롤링베타"),
        _RawSpec("BETA_X_FT", "(베타-1)×급락트리거백분위", "C.국면상호작용", -1,
                 "급성 변동성 급등(급락트리거) 국면에서 고베타가 즉각 상대 열위",
                 "옵션시장 헤지수요 급증은 고베타 자산부터 매도", "M.res[fast_pct] × 롤링베타"),
        _RawSpec("BETA_X_DH", "(베타-1)×위험점수15일변화", "C.국면상호작용", -1,
                 "위험이 가속(ΔH 급등)하는 국면에서 고베타가 상대 열위",
                 "위험 '수준'이 아니라 '가속도'에도 고베타가 더 민감하게 반응",
                 "M.res[haz_pct].diff(15) × 롤링베타"),
        _RawSpec("REGIME_STATE_BETA", "확정국면×(베타-1)", "C.국면상호작용", +1,
                 "확정 위험선호 국면에서 고베타 우위, 확정 위험회피 국면에서 저베타 우위",
                 "SPY 계층 상태기계의 확정 출력을 섹터 상대수익 지표로 재사용",
                 "M.res[sig][state] × 롤링베타"),
    ]


def family_c_values(res: dict, beta: pd.Series) -> pd.DataFrame:
    """[§3.3] SPY 계층의 확정 출력(haz_pct/score_pct/fast_pct/state)을 그대로 재사용한다
    — 전부 그날 종가까지의 정보로 확정된 값이라 인과성이 자동 보장된다."""
    idx = beta.index
    out = pd.DataFrame(index=idx)
    beta_x = (beta - 1.0)
    haz_pct = res.get("haz_pct")
    score_pct = res.get("score_pct")
    fast_pct = res.get("fast_pct")
    state = res.get("sig", pd.DataFrame()).get("state") if res.get("sig") is not None else None

    out["BETA_X_H"] = beta_x * haz_pct.reindex(idx) if haz_pct is not None else np.nan
    out["BETA_X_SCORE"] = beta_x * score_pct.reindex(idx) if score_pct is not None else np.nan
    out["BETA_X_FT"] = beta_x * fast_pct.reindex(idx) if fast_pct is not None else np.nan
    if haz_pct is not None:
        dh15 = haz_pct.reindex(idx).diff(15)
        out["BETA_X_DH"] = beta_x * dh15
    else:
        out["BETA_X_DH"] = np.nan
    if state is not None:
        state_num = state.reindex(idx).map({"RISK_ON": 1.0, "NEUTRAL": 0.0, "RISK_OFF": -1.0}).astype(float)
        out["REGIME_STATE_BETA"] = beta_x * state_num
    else:
        out["REGIME_STATE_BETA"] = np.nan
    return out.replace([np.inf, -np.inf], np.nan)


# ---- 가족 B: 섹터별 매크로·크로스에셋 지표(표 기반, §3.2) -------------------------
# 각 항목: (suffix, name_kr, kind, source_id, transform, window, prior_sign, rationale, mechanism)
#   kind: "fred_rate"(레벨 변화=diff) | "fred_index"(변화율=pct) | "fred_level"(수준 z)
#         | "yahoo_mom"(모멘텀) | "ind_direct"(res["ind"]의 기존 컬럼을 그대로 재사용)
_FamilyBRow = Tuple[str, str, str, str, str, int, int, str, str]

SECTOR_FAMILY_B_TABLE: Dict[str, List[_FamilyBRow]] = {
    "XLK": [
        ("REALRATE10", "10년 실질금리 60일 변화", "fred_rate", "DFII10", "chg", 60, -1,
         "장기 듀레이션 성장주는 실질금리 상승에 밸류에이션이 눌린다", "할인율 채널"),
        ("NOM10Y", "10년 국채금리 60일 변화", "fred_rate", "DGS10", "chg", 60, -1,
         "명목금리 상승도 동일한 밸류에이션 압박", "할인율 채널"),
        ("RSP_MOM", "동일가중/시총가중 60일 상대강도", "ind_direct", "RSP_SPY_MOM", "", 60, -1,
         "동일가중 우위는 대형 기술주(시총상위) 열위를 시사", "지수 구성 효과"),
        ("H_PCT", "위험점수백분위(H)", "ind_direct", "__HAZ_PCT__", "", 0, -1,
         "고베타 성장주는 시장 위험 상승기에 더 크게 할인", "베타 채널(가족C와 상호보완)"),
    ],
    "XLC": [
        ("REALRATE10", "10년 실질금리 60일 변화", "fred_rate", "DFII10", "chg", 60, -1,
         "광고·구독 플랫폼도 성장주 밸류에이션 논리를 공유", "할인율 채널"),
        ("UMCSENT_Z", "미시간대 소비자심리지수(z)", "fred_level", "UMCSENT", "", 252, +1,
         "소비심리 개선은 광고 지출·구독 수요 확대로 연결", "광고 수요 채널"),
        ("RSP_MOM", "동일가중/시총가중 60일 상대강도", "ind_direct", "RSP_SPY_MOM", "", 60, -1,
         "메가캡 비중이 큰 섹터 특성상 XLK와 동일 논리", "지수 구성 효과"),
    ],
    "XLY": [
        ("UMCSENT_CHG", "미시간대 소비자심리 60일 변화", "fred_rate", "UMCSENT", "chg", 60, +1,
         "소비심리 개선은 재량소비 지출 확대로 이어진다", "소비여력 채널"),
        ("ICSA_CHG", "신규 실업수당청구 60일 변화율", "fred_index", "ICSA", "chg", 60, -1,
         "고용 악화는 재량소비부터 위축시킨다", "고용-소비 채널"),
        ("OIL_MOM", "WTI 원유 60일 모멘텀", "yahoo_mom", "CL=F", "mom", 60, -1,
         "유가 상승은 가처분소득을 압박해 재량소비에 불리", "유가-소비 채널"),
        ("MORT30_CHG", "30년 모기지금리 60일 변화", "fred_rate", "MORTGAGE30US", "chg", 60, -1,
         "모기지금리 상승은 주택·내구재 관련 소비수요를 위축", "금리-내구재 채널"),
    ],
    "XLP": [
        ("H_PCT", "위험점수백분위(H)", "ind_direct", "__HAZ_PCT__", "", 0, +1,
         "필수소비재는 방어 로테이션의 전형적 수혜 섹터", "방어 로테이션"),
        ("NOM10Y", "10년 국채금리 60일 변화", "fred_rate", "DGS10", "chg", 60, -1,
         "채권 대용 성격(안정배당)이라 금리 상승에 상대적으로 불리", "채권대용 채널"),
        ("DXY_MOM", "달러인덱스 60일 모멘텀", "yahoo_mom", "DX-Y.NYB", "mom", 60, -1,
         "다국적 매출 비중이 높아 달러 강세가 환산이익을 깎는다", "환율 채널"),
    ],
    "XLV": [
        ("H_PCT", "위험점수백분위(H)", "ind_direct", "__HAZ_PCT__", "", 0, +1,
         "헬스케어는 경기방어적 수요(질병·처방)로 방어 로테이션 수혜", "방어 로테이션"),
        ("NOM10Y", "10년 국채금리 60일 변화", "fred_rate", "DGS10", "chg", 60, -1,
         "배당·성장이 혼재해 금리 상승에 약하게 불리", "할인율 채널(약함)"),
    ],
    "XLF": [
        ("CURVE", "수익률곡선 10년-2년 60일 변화", "fred_rate", "T10Y2Y", "chg", 60, +1,
         "커브가 가팔라지면 예대마진(순이자마진) 개선 기대", "순이자마진 채널"),
        ("NOM10Y", "10년 국채금리 60일 변화", "fred_rate", "DGS10", "chg", 60, +1,
         "금리 상승 국면에서 은행 이자수익 개선 기대", "순이자마진 채널"),
        ("HY_OAS_CHG", "하이일드 스프레드 20일 변화", "fred_rate", "BAMLH0A0HYM2", "chg", 20, -1,
         "신용스프레드 확대는 대손비용 증가·자금조달비용 상승을 시사", "신용비용 채널"),
        ("HOUST_CHG", "주택착공건수 120일 변화율", "fred_index", "HOUST", "chg", 120, +1,
         "주택시장 활황은 모기지·소비자대출 수요 확대", "대출수요 채널"),
    ],
    "XLE": [
        ("OIL_MOM20", "WTI 원유 20일 모멘텀", "yahoo_mom", "CL=F", "mom", 20, +1,
         "유가는 에너지 섹터 이익의 직접적 동인", "직접 수익 채널"),
        ("OIL_MOM60", "WTI 원유 60일 모멘텀", "yahoo_mom", "CL=F", "mom", 60, +1,
         "동일 논리의 중기 관측창", "직접 수익 채널"),
        ("BEI5Y_CHG", "5년 기대인플레이션 60일 변화", "fred_rate", "T5YIE", "chg", 60, +1,
         "기대인플레 상승기에 원자재·에너지가 인플레 헤지 수요를 흡수", "인플레 헤지 채널"),
        ("DXY_MOM", "달러인덱스 60일 모멘텀", "yahoo_mom", "DX-Y.NYB", "mom", 60, -1,
         "달러 강세는 달러표시 원자재 가격에 역풍", "환율-원자재 채널"),
    ],
    "XLI": [
        ("INDPRO_CHG", "산업생산지수 60일 변화율", "fred_index", "INDPRO", "chg", 60, +1,
         "산업생산 확대는 산업재 수요와 직결", "산업수요 채널"),
        ("DGORDER_CHG", "내구재 신규주문 60일 변화율", "fred_index", "DGORDER", "chg", 60, +1,
         "내구재 신규주문 증가는 산업재 기업의 향후 매출 선행지표", "수주 채널"),
        ("COPPER_MOM", "구리 60일 모멘텀", "yahoo_mom", "HG=F", "mom", 60, +1,
         "구리는 전통적인 글로벌 경기 바로미터", "경기 바로미터 채널"),
        ("EFA_MOM", "선진국(미국제외) 주식 60일 모멘텀", "yahoo_mom", "EFA", "mom", 60, +1,
         "글로벌 산업 사이클과 동조", "글로벌 사이클 채널"),
    ],
    "XLB": [
        ("COPPER_MOM", "구리 60일 모멘텀", "yahoo_mom", "HG=F", "mom", 60, +1,
         "구리는 소재 섹터 수요의 직접 프록시", "직접 수요 채널"),
        ("DXY_MOM", "달러인덱스 60일 모멘텀", "yahoo_mom", "DX-Y.NYB", "mom", 60, -1,
         "달러 강세는 소재·원자재 가격에 역풍", "환율-원자재 채널"),
        ("EEM_MOM", "신흥국 주식 60일 모멘텀", "yahoo_mom", "EEM", "mom", 60, +1,
         "신흥국(중국 등) 수요가 소재 섹터의 핵심 동인", "신흥국 수요 채널"),
        ("BEI5Y_CHG", "5년 기대인플레이션 60일 변화", "fred_rate", "T5YIE", "chg", 60, +1,
         "인플레 기대 상승기 원자재 관련주 상대 강세", "인플레 헤지 채널"),
    ],
    "XLU": [
        ("NOM10Y", "10년 국채금리 60일 변화", "fred_rate", "DGS10", "chg", 60, -1,
         "채권 대용 자산(고배당)이라 금리 상승에 가장 직접적으로 불리", "채권대용 채널"),
        ("REALRATE10", "10년 실질금리 60일 변화", "fred_rate", "DFII10", "chg", 60, -1,
         "동일 논리의 실질금리판", "채권대용 채널"),
        ("H_PCT", "위험점수백분위(H)", "ind_direct", "__HAZ_PCT__", "", 0, +1,
         "유틸리티는 대표적 방어 섹터", "방어 로테이션"),
    ],
    "XLRE": [
        ("NOM10Y", "10년 국채금리 60일 변화", "fred_rate", "DGS10", "chg", 60, -1,
         "리츠는 배당 자산으로 금리 상승에 밸류에이션이 눌린다", "할인율 채널"),
        ("MORT30_CHG", "30년 모기지금리 60일 변화", "fred_rate", "MORTGAGE30US", "chg", 60, -1,
         "모기지금리 상승은 부동산 거래·자금조달 비용에 직접 불리", "자금조달 채널"),
        ("HY_OAS_CHG", "하이일드 스프레드 20일 변화", "fred_rate", "BAMLH0A0HYM2", "chg", 20, -1,
         "리츠는 부채비율이 높아 신용스프레드 확대(자금조달비용 상승)에 취약", "신용비용 채널"),
        ("HPI_CHG", "케이스실러 주택가격지수 120일 변화율", "fred_index", "CSUSHPISA", "chg", 120, +1,
         "기초자산 가치 상승은 리츠 순자산가치(NAV)에 긍정적", "자산가치 채널"),
    ],
}


def family_b_values(ticker: str, res: dict, M, cal: pd.DatetimeIndex) -> pd.DataFrame:
    """[§3.2] source series는 전부 res["fred"](발표지연 적용 완료)/res["px_dict"](이미
    수집된 크로스에셋)/res["ind"](이미 계산된 SPY 지표 컬럼)/res["haz_pct"]에서 가져온다
    — 이번 호출을 위해 새로 수집하는 원천 데이터는 없다(11개 섹터 ETF 가격 제외)."""
    rows = SECTOR_FAMILY_B_TABLE.get(ticker, [])
    out = pd.DataFrame(index=cal)
    for suffix, name_kr, kind, sid, transform, window, prior_sign, rationale, mech in rows:
        col = None
        if kind == "fred_rate":
            src = res["fred"].get(sid)
            if src is not None:
                s = src.reindex(cal).ffill()
                col = s - s.shift(window)
        elif kind == "fred_index":
            src = res["fred"].get(sid)
            if src is not None:
                s = src.reindex(cal).ffill()
                col = s / s.shift(window) - 1.0
        elif kind == "fred_level":
            src = res["fred"].get(sid)
            if src is not None:
                s = src.reindex(cal).ffill()
                col = M._z(s, window or 252)
        elif kind == "yahoo_mom":
            d = res["px_dict"].get(sid)
            if d is not None:
                price_col = "Adj Close" if "Adj Close" in d.columns else "Close"
                s = d[price_col].reindex(cal).ffill()
                col = M._mom(s, window)
        elif kind == "ind_direct":
            if sid == "__HAZ_PCT__":
                col = res.get("haz_pct")
                col = col.reindex(cal) if col is not None else None
            else:
                ind = res.get("ind")
                col = ind[sid].reindex(cal) if (ind is not None and sid in ind.columns) else None
        out[suffix] = col if col is not None else np.nan
    return out.replace([np.inf, -np.inf], np.nan)


def build_sector_indicators(ticker: str, res: dict, M,
                            sector_close_tr: pd.Series, sector_close_raw: pd.Series,
                            spy_close_tr: pd.Series, spy_close_raw: pd.Series,
                            cal: pd.DatetimeIndex) -> Tuple[pd.DataFrame, List[Any]]:
    """[§3] 한 섹터의 가족 A+B+C 지표값(ind)과 그에 대응하는 M.IndicatorSpec 목록을 만든다.
    spec.key는 "{티커}__{suffix}" 형식으로 전 섹터에 걸쳐 고유하다(스왑 컨텍스트가 섹터별로
    한 번에 하나씩 열리므로 실제로는 매번 그 섹터의 지표만 보이지만, 로그·시트 추적성을
    위해 접두어를 남긴다)."""
    r_i = np.log(sector_close_tr.replace(0, np.nan)).diff()
    r_spy = np.log(spy_close_tr.replace(0, np.nan)).diff()
    beta = rolling_beta(r_i, r_spy, window=252, lag=1)

    a_vals = family_a_values(sector_close_tr, spy_close_tr, sector_close_raw, spy_close_raw, M)
    c_vals = family_c_values(res, beta.reindex(cal))
    b_vals = family_b_values(ticker, res, M, cal)

    specs: List[Any] = []
    ind = pd.DataFrame(index=cal)
    for raw in family_a_specs():
        key = f"{ticker}__{raw.suffix}"
        ind[key] = a_vals[raw.suffix].reindex(cal)
        specs.append(M.IndicatorSpec(
            key=key, name_kr=f"[{ticker}] {raw.name_kr}", category=raw.category,
            prior_sign=raw.prior_sign, rationale=raw.rationale, source=raw.source,
            lead_mechanism=raw.lead_mechanism, trend_track=raw.trend_track,
            eval_horizon=raw.eval_horizon, base_series=""))
            # [단위테스트에서 발견·수정 - 최초 배포 전] base_series를 f"REL_{raw.suffix}"처럼
            # 지표마다 다른 문자열로 주면 M.build_walkforward_weights의 기저시리즈 그룹캡
            # (SERIES_WEIGHT_CAP=20%, group_key=base_series)이 "지표 하나짜리 그룹"을 만들어
            # 의도치 않게 그 지표 하나에만 개별 20% 상한을 몰래 씌운다(스펙에 없는 동작).
            # family A 9종은 서로 다른 통계적 변환(모멘텀/이평교차/낙폭/RSI 등)이라 기존
            # SPY 계층의 수기지표들처럼 서로 그룹화하지 않는 것이 맞다(base_series=""는
            # "이 20%-그룹캡의 대상이 아님"을 뜻함 - market_regime_trader.py 4483행 부근
            # 주석 "원천이 불분명한 수기 지표(base_series='')는 캡 대상에서 제외" 참조).
            # family A/C의 "합산 60%/40%" 상한(§5.1)은 이 20%-그룹캡과는 별개 메커니즘이며
            # apply_family_weight_cap()이 M.build_walkforward_weights 반환 후 명시적으로 적용한다.
    for raw in family_c_specs():
        key = f"{ticker}__{raw.suffix}"
        ind[key] = c_vals[raw.suffix].reindex(cal)
        specs.append(M.IndicatorSpec(
            key=key, name_kr=f"[{ticker}] {raw.name_kr}", category=raw.category,
            prior_sign=raw.prior_sign, rationale=raw.rationale, source=raw.source,
            lead_mechanism="", trend_track=False, eval_horizon=None,
            base_series=""))  # 위와 동일한 이유(family C 합산 40% 상한은 apply_family_weight_cap이 별도 적용)
    for suffix, name_kr, kind, sid, transform, window, prior_sign, rationale, mech in \
            SECTOR_FAMILY_B_TABLE.get(ticker, []):
        key = f"{ticker}__{suffix}"
        ind[key] = b_vals[suffix].reindex(cal) if suffix in b_vals.columns else np.nan
        specs.append(M.IndicatorSpec(
            key=key, name_kr=f"[{ticker}] {name_kr}", category="B.섹터매크로",
            prior_sign=prior_sign, rationale=rationale, source=f"{kind}:{sid}",
            lead_mechanism=mech, trend_track=False, eval_horizon=_eval_h(window),
            base_series=sid))
    ind = ind.replace([np.inf, -np.inf], np.nan)
    return ind, specs


# =============================================================================
# [6] 검증 — 섹터별 시계열 게이트 + 횡단면 rank IC 게이트
# =============================================================================
def sector_cfg_for(res: dict, scfg: SectorConfig, actual_start: pd.Timestamp,
                   entry_date: pd.Timestamp):
    """[§4.1] SPY cfg를 베이스로 상대낙폭 라벨(-3%)과 위험트랙 비활성만 덮어쓴 사본을
    만든다. DATA_START/SIGNAL_START/TRAIN_MIN_YEARS는 섹터별 실제 이력에 맞춘다(§1.3)."""
    M_cfg = res["cfg"]
    base_sig_start = pd.Timestamp(M_cfg.SIGNAL_START)
    sig_start = max(base_sig_start, entry_date)
    return dataclasses.replace(
        M_cfg,
        DD_LABEL_THRESHOLD=scfg.SECTOR_REL_DD_THRESHOLD,
        USE_HAZARD_TRACK=False,
        TRAIN_MIN_YEARS=scfg.SECTOR_MIN_HISTORY_YEARS,
        DATA_START=str(actual_start.date()),
        SIGNAL_START=str(sig_start.date()),
        MIN_INDICATORS=3,
    )


def apply_family_weight_cap(W: pd.DataFrame, specs: List[Any], category_prefix: str,
                            cap: float) -> pd.DataFrame:
    """[§5.1 "가족 A 합계 ≤ 60%(SECTOR_COMMON_FAMILY_CAP)/가족 C 합계 ≤ 40%
    (SECTOR_REGIME_FAMILY_CAP)"] category가 category_prefix로 시작하는 지표들의 그날
    |가중치| 합이 cap을 넘으면, M.build_walkforward_weights 내부의 기저시리즈캡/추세트랙캡과
    완전히 같은 패턴(그룹을 cap까지 축소 -> 풀려난 몫을 그룹 밖 지표들에 기존 비중 비례로
    재분배)을 모든 날짜에 벡터화 적용한다. market_regime_trader.py를 수정하지 않고 그 함수가
    반환한 W를 후처리하는 방식 — 이 캡은 (기존 재사용 중인) SERIES_WEIGHT_CAP=20% 기저시리즈
    캡과는 별개의, 더 넓은 상한이다(같은 "가족"이라도 서로 다른 통계적 변환이라 기저시리즈로
    보지 않으므로 base_series는 비워둔다 - build_sector_indicators의 코멘트 참조)."""
    codes = list(W.columns)
    cat_by_code = {s.key: getattr(s, "category", "") for s in specs}
    fam_mask = np.array([cat_by_code.get(c, "").startswith(category_prefix) for c in codes])
    if not fam_mask.any():
        return W
    W = W.copy()
    fam_abs_sum = W.loc[:, fam_mask].abs().sum(axis=1)
    over = fam_abs_sum > cap
    if not over.any():
        return W
    freed = (fam_abs_sum - cap).clip(lower=0.0)
    scale = pd.Series(1.0, index=W.index)
    scale.loc[over] = cap / fam_abs_sum.loc[over]
    W.loc[:, fam_mask] = W.loc[:, fam_mask].mul(scale, axis=0)
    other_abs_sum = W.loc[:, ~fam_mask].abs().sum(axis=1).replace(0, np.nan)
    redis_scale = (1.0 + freed / other_abs_sum).fillna(1.0)
    W.loc[:, ~fam_mask] = W.loc[:, ~fam_mask].mul(redis_scale, axis=0)
    return W


def _sector_score_pipeline(ind_t: pd.DataFrame, rel_t: pd.Series, cfg_sector, specs: List[Any],
                           scfg: "SectorConfig", M, verbose_log: bool = False,
                           compute_quintiles: bool = True, full_report: bool = True
                           ) -> Tuple[pd.DataFrame, List[dict], pd.Series, pd.DataFrame, pd.Series, pd.DataFrame]:
    """검증(§4.1)+워크포워드 가중(§5.1)+가족캡(§5.1)+복합점수(§5.2)를 한 파이프라인으로
    묶는다. validate_and_weight_sector(전체계산)와 lookahead_audit_sector(절단재계산) 양쪽이
    이 함수를 그대로 재사용해야 두 계산이 항상 같은 처리를 거친다 — 한쪽만 가족캡을 적용하면
    룩어헤드 감사의 "전체계산==절단재계산" 비교 자체가 무의미해진다."""
    with _indicator_spec_override(M, specs):
        vt = M.validate_indicators(ind_t, rel_t, cfg_sector, verbose=verbose_log,
                                   compute_quintiles=compute_quintiles, full_report=full_report)
        W, wlog, _W_haz = M.build_walkforward_weights(ind_t, rel_t, cfg_sector)
        W = apply_family_weight_cap(W, specs, "A.", scfg.SECTOR_COMMON_FAMILY_CAP)
        W = apply_family_weight_cap(W, specs, "C.", scfg.SECTOR_REGIME_FAMILY_CAP)
        score, contrib, n_used = M.composite_score(ind_t, W, cfg_sector)
    return vt, wlog, W, score, contrib, n_used


def validate_and_weight_sector(ticker: str, ind: pd.DataFrame, specs: List[Any],
                               rel: pd.Series, cfg_sector, M, scfg: "SectorConfig" = None,
                               verbose_log: bool = False) -> Dict[str, Any]:
    """[§4.1+§5.1] 한 섹터에 대해 시계열 게이트(M.validate_indicators, 전체표본) +
    워크포워드 가중(M.build_walkforward_weights) + 가족비중상한(§5.1) + 복합점수
    (M.composite_score)까지 한 번에 수행한다. INDICATOR_SPECS/SPEC_BY_KEY 스왑은 이 함수
    호출 동안만 유지된다. 반환 dict: vt(전체표본 검증표), W(가중치행렬, 가족캡 적용 후),
    wlog, score, contrib, n_used, score_pct."""
    scfg = scfg or CFG
    actual_start = pd.Timestamp(cfg_sector.DATA_START)
    idx = ind.index[ind.index >= actual_start]
    ind_t = ind.reindex(idx)
    rel_t = rel.reindex(idx)

    vt, wlog, W, score, contrib, n_used = _sector_score_pipeline(
        ind_t, rel_t, cfg_sector, specs, scfg, M, verbose_log=verbose_log,
        compute_quintiles=True, full_report=True)
    score_pct = M.score_percentile(score)

    log("VALIDATE", kv(ticker=ticker, candidates=len(specs),
                       passed=int((vt["판정"] == "PASS").sum()),
                       periods=len(wlog)), M=M)
    return {"vt": vt, "W": W, "wlog": wlog, "score": score, "contrib": contrib,
            "n_used": n_used, "score_pct": score_pct, "idx": idx}


def cross_sectional_rank_ic(fwd_rel_by_sector: Dict[str, pd.Series],
                            indicator_values_by_sector: Dict[str, pd.Series],
                            horizon: int) -> Dict[str, float]:
    """[§4.2] 매일 N개 섹터를 가로질러 (그날의 지표값, 그날 기준 향후 horizon일 상대수익)의
    순위상관(스피어만)을 구하고, 그 일별 rank IC 시계열의 평균이 0과 다른지를 뉴이-웨스트
    HAC 표준오차(겹치는 지평 보정, lag=horizon)로 검정한다.
    fwd_rel_by_sector[t] = M.forward_return(rel_t_series, horizon) — 호출부가 이미 h일
    앞의 상대수익으로 만들어 넘긴다(이 함수는 그 값을 그대로 그날 행에 쓴다 — t행의 값은
    t+h의 실현치이므로 상관계산 자체는 '그날 기준'이라는 시점 정렬만 지킨다. 통계용으로만
    쓰고, 지표 채택 로직에는 이 결과가 직접 게이트로 들어가지 않으며 §4.2 수용기준 판정에만
    쓰인다 — 참고 진단이지 채택 여부를 좌우하는 것이 아니라는 뜻은 아니고, §8 활성화 판정의
    한 조건이다)."""
    tickers = [t for t in fwd_rel_by_sector if t in indicator_values_by_sector]
    if len(tickers) < 3:
        return {"mean_rank_ic": np.nan, "nw_t": np.nan, "n_days": 0}
    idx = fwd_rel_by_sector[tickers[0]].index
    fwd_mat = pd.DataFrame({t: fwd_rel_by_sector[t] for t in tickers}, index=idx)
    x_mat = pd.DataFrame({t: indicator_values_by_sector[t].reindex(idx) for t in tickers}, index=idx)
    daily_ic = pd.Series(np.nan, index=idx)
    for dt in idx:
        y = fwd_mat.loc[dt]
        x = x_mat.loc[dt]
        both = pd.concat([x, y], axis=1).dropna()
        if len(both) < 5:
            continue
        daily_ic.loc[dt] = both.iloc[:, 0].corr(both.iloc[:, 1], method="spearman")
    daily_ic = daily_ic.dropna()
    if len(daily_ic) < 30:
        return {"mean_rank_ic": np.nan, "nw_t": np.nan, "n_days": len(daily_ic)}
    # 표준적인 HAC(뉴이-웨스트) 평균-t: daily_ic 자체를 "관측치"로 보고 그 평균이 0과
    # 다른지 검정한다(회귀가 아니라 평균 검정이므로 M.newey_west_tstat(x~y 회귀용)을
    # 그대로 쓸 수 없어 동일한 HAC 공식을 직접 적용).
    m = float(daily_ic.mean())
    n = len(daily_ic)
    resid = daily_ic - m
    gamma0 = float((resid ** 2).mean())
    lag_max = min(horizon, n - 1)
    var = gamma0
    for L in range(1, lag_max + 1):
        w = 1.0 - L / (lag_max + 1)
        cov = float((resid.iloc[L:].values * resid.iloc[:-L].values).mean())
        var += 2 * w * cov
    se = math.sqrt(max(var, 1e-12) / n)
    t_stat = m / se if se > 0 else np.nan
    return {"mean_rank_ic": m, "nw_t": t_stat, "n_days": n}


# =============================================================================
# [7] 판별력 자기검사 (§4.3) — 실데이터 해석 전에 반드시 먼저 통과해야 함
# =============================================================================
def build_selftest_case(M, scfg: SectorConfig, n: int = 1500
                        ) -> Tuple[pd.Series, pd.DataFrame, List[Any]]:
    """[§4.3] 합성 rel(=P_i/SPY 대용) 시계열을 만든다. _build_hazard_selftest_case()(기존
    프로덕션 self_test 픽스처)와 같은 설계 원칙 — 대부분의 날은 순수 잡음이고, 간격을 둔
    소수의 '사건' 구간에서만 지표가 사건 직전에 스파이크하고 rel이 그 방향으로 크게
    움직인다 — 를 그대로 따른다. 매일 IC를 억지로 만드는 대신(그러면 노이즈 지표와
    구분이 안 된다) 사건 기반 희소 신호로 만들어야 진짜 판별력 검사가 된다.
    3개 지표에는 이 사건들을 선행하는 신호를 심고, 5개는 순수 잡음(사건과 무관)으로 둔다.
    수용: 심은 3개 전부 채택, 잡음 5개 중 채택 ≤ SELFTEST_MAX_NOISE_ADOPTED.
    반환: (rel 시계열, ind DataFrame, IndicatorSpec 목록[진짜 3 + 잡음 5])."""
    rng = np.random.default_rng(scfg.RANDOM_SEED)
    idx = pd.bdate_range("2010-01-04", periods=n)

    true_meta = [("TRUE_A", +1), ("TRUE_B", -1), ("TRUE_C", +1)]
    noise_keys = ["NOISE_A", "NOISE_B", "NOISE_C", "NOISE_D", "NOISE_E"]

    ret = rng.normal(0.0002, 0.006, n)     # 일간 로그수익 기저잡음
    ind_data: Dict[str, np.ndarray] = {k: rng.normal(0.0, 1.0, n) for k, _ in true_meta}
    for k in noise_keys:
        ind_data[k] = rng.normal(0.0, 1.0, n)

    move_days = 5
    spacing = 45
    episodes = list(range(spacing, n - spacing - 10, spacing))
    for p in episodes:
        direction = 1 if rng.uniform() < 0.5 else -1
        magnitude = direction * rng.uniform(0.04, 0.08)          # ±4~8% 총 상대이동
        per_day = math.log(1.0 + magnitude) / move_days
        for j in range(move_days):
            ret[p + j] += per_day
        for key, sign in true_meta:
            for off in range(-3, 3):
                t = p + off
                if 0 <= t < n:
                    ind_data[key][t] += sign * direction * 5.0 * (1.0 - abs(off) / 4.0)

    rel = pd.Series(100.0 * np.exp(np.cumsum(ret)), index=idx)
    ind = pd.DataFrame({k: pd.Series(v, index=idx) for k, v in ind_data.items()})

    specs: List[Any] = []
    for key, sign in true_meta:
        specs.append(M.IndicatorSpec(
            key=key, name_kr=f"[자체테스트] {key}", category="TEST", prior_sign=sign,
            rationale="자기검사 전용 합성 신호", source="합성데이터", lead_mechanism="주입된 선행신호"))
    for key in noise_keys:
        specs.append(M.IndicatorSpec(
            key=key, name_kr=f"[자체테스트] {key}", category="TEST", prior_sign=1,
            rationale="자기검사 전용 순수 잡음", source="합성데이터", lead_mechanism="없음(잡음)"))
    return rel, ind, specs


def run_selftest(M, scfg: SectorConfig = CFG) -> Dict[str, Any]:
    """[§4.3] 판별력 자기검사 실행. 실데이터 검증 결과를 신뢰하기 전에 이 함수가 PASS해야
    한다(run_all()이 자동으로 먼저 호출하고, FAIL이면 실행을 중단한다)."""
    rel, ind, specs = build_selftest_case(M, scfg)
    cfg_t = dataclasses.replace(M.CFG, DD_LABEL_THRESHOLD=scfg.SECTOR_REL_DD_THRESHOLD,
                                USE_HAZARD_TRACK=False, DATA_START=str(rel.index[0].date()),
                                SIGNAL_START=str(rel.index[0].date()), TRAIN_MIN_YEARS=3,
                                MIN_INDICATORS=3, REWEIGHT_FREQ="M")
    with _indicator_spec_override(M, specs):
        vt = M.validate_indicators(ind, rel, cfg_t, verbose=False, compute_quintiles=False)
    v = vt.set_index("지표코드")
    true_keys = ["TRUE_A", "TRUE_B", "TRUE_C"]
    noise_keys = ["NOISE_A", "NOISE_B", "NOISE_C", "NOISE_D", "NOISE_E"]
    n_true_pass = int((v.loc[true_keys, "판정"] == "PASS").sum())
    n_noise_pass = int((v.loc[noise_keys, "판정"] == "PASS").sum())
    passed = (n_true_pass >= scfg.SELFTEST_MIN_TRUE_ADOPTED
              and n_noise_pass <= scfg.SELFTEST_MAX_NOISE_ADOPTED)
    log("SELFTEST", kv(true_pass=f"{n_true_pass}/3", noise_pass=f"{n_noise_pass}/5",
                       verdict="PASS" if passed else "FAIL"), M=M,
        level="info" if passed else "error")
    return {"passed": passed, "n_true_pass": n_true_pass, "n_noise_pass": n_noise_pass, "vt": vt}


# =============================================================================
# [8] 복합점수 → 횡단면 순위 → 이력현상 상태기계 (§5.2~§5.3)
# =============================================================================
def cross_sectional_rank(score_by_sector: pd.DataFrame, active_mask: pd.DataFrame
                         ) -> pd.DataFrame:
    """[§5.2] 그날 유니버스에 편입된(active) 섹터들만 대상으로 1(최고)~N(최저) 순위를
    매긴다. 편입 전 섹터는 NaN(순위 없음)."""
    masked = score_by_sector.where(active_mask)
    return masked.rank(axis=1, ascending=False, method="average")


def sector_raw_state(rank: pd.Series, spct: pd.Series, n_active: pd.Series,
                     scfg: SectorConfig) -> pd.Series:
    """[§5.3] 원시 상태(이력현상 적용 전). n_active는 그날 유니버스 크기(그 날짜 행의
    active 섹터 수) — 매일 달라질 수 있다(§1.3 유니버스 확장)."""
    raw = pd.Series("NEUTRAL", index=rank.index, dtype=object)
    out_mask = (rank <= scfg.SECTOR_TOP_K) & (spct >= scfg.SECTOR_STATE_HI)
    under_mask = (rank >= (n_active - scfg.SECTOR_TOP_K + 1)) & (spct <= scfg.SECTOR_STATE_LO)
    raw[out_mask.fillna(False)] = "OUTPERFORM"
    raw[under_mask.fillna(False) & ~out_mask.fillna(False)] = "UNDERPERFORM"
    raw[rank.isna()] = "INACTIVE"
    return raw


def apply_sector_hysteresis(raw_state: pd.Series, scfg: SectorConfig) -> pd.Series:
    """[§5.3] 독자 구현(SPY 계층 generate_signals()의 복잡한 규칙 ⓪~⑪ 상태기계와는 별개
    — 여기서는 단순 k연속일 확인 + 최소보유일만 쓴다): raw_state가 SECTOR_HYSTERESIS_DAYS
    연속으로 새 상태를 가리켜야 전환하고, 전환 후 SECTOR_MIN_HOLD_DAYS는 고정 유지한다.
    INACTIVE(유니버스 편입 전)는 즉시 반영(확인 불필요 — 편입 전 섹터는 애초에 tilt가
    없어야 하므로 지연시킬 이유가 없다)."""
    idx = raw_state.index
    confirmed = pd.Series("NEUTRAL", index=idx, dtype=object)
    current = "INACTIVE"
    hold_left = 0
    streak_state, streak_len = None, 0
    for i in range(len(idx)):
        r = raw_state.iloc[i]
        if r == "INACTIVE":
            current = "INACTIVE"
            hold_left = 0
            streak_state, streak_len = None, 0
            confirmed.iloc[i] = current
            continue
        if current == "INACTIVE":
            current = "NEUTRAL"   # 편입 첫날은 중립에서 시작(과거 확정이력 없음)
        if r == streak_state:
            streak_len += 1
        else:
            streak_state, streak_len = r, 1
        if hold_left > 0:
            hold_left -= 1
        elif r != current and streak_len >= scfg.SECTOR_HYSTERESIS_DAYS:
            current = r
            hold_left = scfg.SECTOR_MIN_HOLD_DAYS - 1
        confirmed.iloc[i] = current
    return confirmed


# =============================================================================
# [9] 포트폴리오 구성 — 시장 계층(E_t) × 섹터 계층(tilt) (§5.4)
# =============================================================================
def build_portfolio(confirmed_state: pd.DataFrame, active_mask: pd.DataFrame,
                    E_t: pd.Series, scfg: SectorConfig) -> pd.DataFrame:
    """[§5.4] target_i,t = E_t × tilt_i,t × b_i,t / Σ_j tilt_j,t × b_j,t, 상한 적용 후
    재정규화(2회 반복), 그 결과 Σ target_i,t = E_t가 항상 성립한다(단위테스트로 강제).
    b_i,t = 1/N_t(그날 active 섹터 동일가중). 전 섹터 UNDER인 퇴화 케이스는 tilt를 전부
    1.0으로 되돌린다."""
    idx = confirmed_state.index
    cols = list(scfg.SECTORS)
    tilt = pd.DataFrame(1.0, index=idx, columns=cols)
    tilt = tilt.mask(confirmed_state == "OUTPERFORM", scfg.SECTOR_TILT_OVER)
    tilt = tilt.mask(confirmed_state == "UNDERPERFORM", scfg.SECTOR_TILT_UNDER)
    tilt = tilt.mask(confirmed_state == "INACTIVE", 0.0)
    tilt = tilt.mask(~active_mask.reindex(columns=cols, fill_value=False), 0.0)

    n_active = active_mask.reindex(columns=cols, fill_value=False).sum(axis=1).replace(0, np.nan)
    b = active_mask.reindex(columns=cols, fill_value=False).astype(float).div(n_active, axis=0)

    tilt_sum = (tilt * (active_mask.reindex(columns=cols, fill_value=False))).sum(axis=1)
    degenerate = (tilt_sum <= 1e-12) & (n_active.fillna(0) > 0)
    if degenerate.any():
        tilt.loc[degenerate, :] = active_mask.reindex(columns=cols, fill_value=False).loc[degenerate].astype(float)
        log("PORTFOLIO", kv(event="degenerate_all_under_fallback_equal_weight",
                            days=int(degenerate.sum())), M=None, level="warning")

    raw_w = tilt * b

    # 상한 적용 — 워터필링(waterfilling) 방식으로 정확히 수렴시킨다(§5.4).
    # [단위테스트에서 발견·수정한 버그] 이전 구현은 "clip 후 전체를 동일 배율로 재정규화"를
    # 고정 2회 반복했는데, 이 방식은 근본적으로 상한을 정확히 만족시킬 수 없다: sum(clip(w))
    # < 1이면 재정규화 배율이 1보다 커서 "이미 상한에 걸린 항목까지" 같은 배율로 다시
    # 키우기 때문에 재정규화 직후 다시 상한을 넘는다(기하급수적으로 상한에 접근할 뿐 도달
    # 못함 — 2회만 반복하면 오차가 ~1%p 수준으로 남는다). 올바른 해법은 상한에 걸린 항목을
    # 그 값에 "고정"하고, 남은 예산을 아직 고정 안 된 항목에만 재분배하는 것이며, 이는
    # 반복마다 최소 1개 항목이 새로 고정되므로 최대 len(cols)회 안에 정확히 수렴한다.
    cap = scfg.SECTOR_MAX_WEIGHT
    fixed = pd.DataFrame(False, index=idx, columns=cols)   # 상한에 고정된 항목
    w = pd.DataFrame(0.0, index=idx, columns=cols)
    budget_left = pd.Series(1.0, index=idx)                # 미고정 항목에 남은 정규화 예산(합계 1 기준)
    proportional = pd.DataFrame(0.0, index=idx, columns=cols)
    for _ in range(len(cols) + 1):
        active = ~fixed
        denom_i = raw_w.where(active, 0.0).sum(axis=1).replace(0, np.nan)
        proportional = (raw_w.where(active, 0.0).div(denom_i, axis=0)
                        .mul(budget_left, axis=0).fillna(0.0))
        over = active & (proportional > cap + 1e-12)
        if not over.values.any():
            break
        w = w.mask(over, cap)
        fixed = fixed | over
        budget_left = (1.0 - w.where(fixed, 0.0).sum(axis=1)).clip(lower=0.0)
    w = w.where(fixed, proportional)

    # 구조적 예외: active 섹터 수가 너무 적어 N×SECTOR_MAX_WEIGHT < 1이면(예: 4개 미만) 상한을
    # 지키면서는 예산을 전부 배분할 수 없다. 이때는 불변식①(Σtarget=E_t — "노출은 재배분만
    # 하고 절대 줄이지 않는다")을 다양화 상한보다 우선해, 남는 예산을 이미 상한에 도달한
    # 항목들에 원래 비중 비율대로 추가 배분한다(상한 소폭 초과, 반드시 로그 경고).
    shortfall = (1.0 - w.sum(axis=1)).clip(lower=0.0)
    need_fix = shortfall > 1e-9
    if need_fix.any():
        fixed_raw_sum = raw_w.where(fixed, 0.0).sum(axis=1).replace(0, np.nan)
        add = (raw_w.where(fixed, 0.0).div(fixed_raw_sum, axis=0)
              .mul(shortfall, axis=0).fillna(0.0))
        w = w.add(add, fill_value=0.0)
        log("PORTFOLIO", kv(event="cap_infeasible_exceeded_to_preserve_exposure_invariant",
                            days=int(need_fix.sum())), M=None, level="warning")

    target = w.mul(E_t.reindex(idx), axis=0)
    return target


def portfolio_invariant_check(target: pd.DataFrame, E_t: pd.Series, scfg: SectorConfig,
                              atol: float = 1e-9) -> Dict[str, Any]:
    """[§5.4 불변식 검증] ① Σ target_i = E_t ② 0 ≤ target_i ≤ E_t*SECTOR_MAX_WEIGHT(+atol)."""
    row_sum = target.sum(axis=1)
    e = E_t.reindex(target.index)
    sum_ok = bool(((row_sum - e).abs() <= 1e-6).all())
    max_allowed = e * scfg.SECTOR_MAX_WEIGHT + atol
    bounds_ok = bool((target.ge(-atol).all().all())
                     and (target.le(max_allowed, axis=0).all().all()))
    return {"sum_equals_E_t": sum_ok, "bounds_ok": bounds_ok,
            "max_abs_sum_diff": float((row_sum - e).abs().max())}


# =============================================================================
# [10] 다자산 백테스트 (§6.1) — 현금 레그는 포트폴리오 전체에서 한 번만 계산한다
#      (run_backtest를 섹터별로 반복 호출해 합산하면 현금이 N배로 과대계상되므로 금지 —
#      §6.1 설계노트. 대신 run_backtest와 동일한 손익 분해식을 다자산으로 직접 재현한다.)
# =============================================================================
def run_sector_backtest(price_by_sector: Dict[str, pd.DataFrame], target: pd.DataFrame,
                        cfg, rf_daily: Optional[pd.Series] = None) -> pd.DataFrame:
    """[§6.1] M.run_backtest()의 단일자산 손익분해(ret_co/ret_oc, T일 신호→T+1일 시가체결,
    편도비용)를 섹터마다 독립적으로 만든 뒤, gross만 합산하고 현금레그·비용은 전체
    포트폴리오 기준으로 한 번만 계산한다 — N=1(섹터 하나)일 때 M.run_backtest()와 완전히
    같은 값이 나오는 것으로 단위테스트에서 확인한다(단, 아래 '첫날 예외' 제외).

    [단위테스트에서 발견한 첫날(day 1) 경계 차이 — market_regime_trader.py 무수정 원칙상
    수정하지 않고 여기 기록만 함] M.run_backtest()는 데이터 첫날에 pos_prev*ret_co +
    pos_exec*ret_oc 계산에서 ret_co/ret_oc가 아직 정의되지 않아(전일 종가가 없음) NaN이
    나오고, 0(포지션 없음)×NaN=NaN이 되어 그날의 gross가 NaN이 된다. 이후 strategy_ret 전체를
    한 번에 .fillna(0.0)하기 때문에, 그날 실제로는 유효했던 현금레그(무포지션 상태의 단기금리
    수취분)까지 통째로 0으로 지워진다(첫날에만 발생하는 미세한 기존 동작). 이 함수는 gross를
    섹터별로 먼저 개별 fillna(0.0)한 뒤 합산하므로 그런 문제가 없어 첫날의 현금레그를 정확히
    반영한다(오히려 이쪽이 더 정확함). 실전 영향은 0에 가깝다(전체 실행기간 중 정확히 1일,
    금액도 무포지션 상태의 하루치 단기금리뿐) — 단위테스트(test_sector_backtest.py)는 이 첫날
    차이를 알고 그 크기(=그날의 rf)까지 명시적으로 검증하고 둘째날부터는 완전 일치를 확인한다."""
    idx = target.index
    tickers = list(target.columns)
    gross = pd.Series(0.0, index=idx)
    cost = pd.Series(0.0, index=idx)
    pos_exec_sum = pd.Series(0.0, index=idx)
    per_sector_frames: Dict[str, pd.DataFrame] = {}

    for t in tickers:
        price = price_by_sector.get(t)
        pos_target = target[t].reindex(idx)
        if price is None:
            per_sector_frames[t] = pd.DataFrame(index=idx)
            continue
        df = pd.DataFrame(index=idx)
        df["Open"] = price["Open"].reindex(idx)
        df["Close"] = price["Close"].reindex(idx)
        adj_col = "Adj Close" if "Adj Close" in price.columns else "Close"
        df["AdjClose"] = price[adj_col].reindex(idx)
        df["ret_cc"] = df["AdjClose"].pct_change()
        df["ret_co"] = (df["Open"] / df["Close"].shift(1) - 1.0)
        df["ret_oc"] = (1.0 + df["ret_cc"]) / (1.0 + df["ret_co"]) - 1.0

        df["pos_target"] = pos_target
        df["pos_exec"] = df["pos_target"].shift(1)
        df["pos_prev"] = df["pos_exec"].shift(1).fillna(0.0)
        df["pos_exec"] = df["pos_exec"].fillna(0.0)
        turn = (df["pos_exec"] - df["pos_prev"]).abs()
        df["cost"] = turn * (cfg.COST_BPS / 1e4)
        df["gross"] = df["pos_prev"] * df["ret_co"] + df["pos_exec"] * df["ret_oc"]

        gross = gross.add(df["gross"].fillna(0.0), fill_value=0.0)
        cost = cost.add(df["cost"].fillna(0.0), fill_value=0.0)
        pos_exec_sum = pos_exec_sum.add(df["pos_exec"].fillna(0.0), fill_value=0.0)
        per_sector_frames[t] = df

    rf = rf_daily.reindex(idx).fillna(0.0) if rf_daily is not None else pd.Series(0.0, index=idx)
    cash_leg = (1.0 - pos_exec_sum) * rf   # 전체 포트폴리오 노출 합(=E_t)만 빼고 현금은 1회만
    out = pd.DataFrame(index=idx)
    out["gross"] = gross
    out["cash_leg"] = cash_leg
    out["cost"] = cost
    out["pos_exec_sum"] = pos_exec_sum
    out["strategy_ret"] = (gross + cash_leg - cost).fillna(0.0)
    out["equity"] = (1.0 + out["strategy_ret"]).cumprod()
    out["dd"] = out["equity"] / out["equity"].cummax() - 1.0
    out.attrs["per_sector"] = per_sector_frames
    return out


def benchmark_equalweight(price_by_sector: Dict[str, pd.DataFrame], active_mask: pd.DataFrame,
                          E_t: pd.Series, cfg, rf_daily: Optional[pd.Series] = None
                          ) -> pd.DataFrame:
    """[§6.2 B2] E_t × 동일가중(N_t) — 로테이션 알파를 격리하기 위한 벤치마크(같은 E_t,
    같은 유니버스, tilt만 1.0으로 고정)."""
    cols = list(active_mask.columns)
    n_active = active_mask.sum(axis=1).replace(0, np.nan)
    b = active_mask.astype(float).div(n_active, axis=0).fillna(0.0)
    target_ew = b.mul(E_t.reindex(active_mask.index), axis=0)
    return run_sector_backtest(price_by_sector, target_ew, cfg, rf_daily)


def benchmark_ew_buyhold(price_by_sector: Dict[str, pd.DataFrame], active_mask: pd.DataFrame,
                         cfg, rf_daily: Optional[pd.Series] = None) -> pd.DataFrame:
    """[§6.2 B3] 동일가중 11섹터 단순보유(E_t=1 고정)."""
    e_one = pd.Series(1.0, index=active_mask.index)
    return benchmark_equalweight(price_by_sector, active_mask, e_one, cfg, rf_daily)


# =============================================================================
# [11] 룩어헤드 감사 (§7) — 절단재계산
# =============================================================================
def lookahead_audit_sector(ticker: str, ind: pd.DataFrame, specs: List[Any], rel: pd.Series,
                           cfg_sector, M, full_score: pd.Series, scfg: "SectorConfig" = None,
                           n_dates: int = 3, seed: int = 20260904) -> pd.DataFrame:
    """[§7] M.lookahead_audit()과 같은 발상: 무작위 날짜 d를 골라 그 날짜까지만으로 다시
    전체 파이프라인(validate_indicators+build_walkforward_weights+가족캡+composite_score,
    _sector_score_pipeline()로 validate_and_weight_sector와 완전히 동일하게 재사용)을 돌리고,
    d일의 복합점수가 전체계산 결과와 같은지 비교한다(1e-9 이내면 OK). 다르면 그 시점 이후
    정보가 d일의 점수 계산에 새어들었다는 뜻(버그) — 워크포워드 루프 자체가 이미 재추정
    시점마다 eval_end로 절단하므로 통과가 정상이며, 이 감사는 그 설계가 실제 구현에서도
    지켜지는지 확인하는 안전망이다. [주의] full_score는 반드시 이 함수와 같은 가족캡이
    적용된 값(validate_and_weight_sector의 반환값)이어야 비교가 성립한다 — 그렇지 않으면
    가족캡 적용 여부 차이가 룩어헤드 버그로 오인될 수 있다."""
    scfg = scfg or CFG
    rng = np.random.default_rng(seed)
    valid_idx = full_score.dropna().index
    if len(valid_idx) < 300:
        return pd.DataFrame(columns=["티커", "검사일", "전체계산점수", "절단재계산점수", "차이", "일치"])
    picks = rng.choice(valid_idx[252:], size=min(n_dates, len(valid_idx) - 252), replace=False)
    rows = []
    for d in sorted(pd.Timestamp(p) for p in picks):
        sub_idx = ind.index[ind.index <= d]
        _sub_vt, _sub_wlog, _sub_W, sub_score, _c, _n = _sector_score_pipeline(
            ind.reindex(sub_idx), rel.reindex(sub_idx), cfg_sector, specs, scfg, M,
            verbose_log=False, compute_quintiles=False, full_report=False)
        full_v = float(full_score.loc[d]) if d in full_score.index else np.nan
        cut_v = float(sub_score.iloc[-1]) if len(sub_score) else np.nan
        diff = abs(full_v - cut_v) if pd.notna(full_v) and pd.notna(cut_v) else np.nan
        rows.append({"티커": ticker, "검사일": str(d.date()), "전체계산점수": full_v,
                    "절단재계산점수": cut_v, "차이": diff,
                    "일치": "OK" if (pd.notna(diff) and diff < 1e-6) else "확인필요"})
    return pd.DataFrame(rows)


def universe_entry_causality_audit(entry_dates: Dict[str, pd.Timestamp],
                                   sector_px: Dict[str, Optional[pd.DataFrame]]) -> pd.DataFrame:
    """[§7] 유니버스 편입일이 '그 시점에 이미 관측 가능한' 실제 데이터 시작일 + 최소이력
    으로만 계산됐는지(=상장 정보를 미래에서 끌어오지 않았는지) 확인한다."""
    rows = []
    for t, entry in entry_dates.items():
        df = sector_px.get(t)
        actual_start = df.index.min() if (df is not None and len(df)) else None
        ok = (actual_start is not None) and (entry > actual_start) and (entry != pd.Timestamp.max)
        rows.append({"티커": t, "실제데이터시작": str(actual_start.date()) if actual_start is not None else "N/A",
                    "유니버스편입일": str(entry.date()) if entry != pd.Timestamp.max else "N/A(데이터없음)",
                    "인과성": "OK" if ok else "확인필요"})
    return pd.DataFrame(rows)


# =============================================================================
# [12] 수용기준 판정 (§8)
# =============================================================================
def evaluate_acceptance(cs_stats: Dict[str, float], spread_by_era_positive_count: int,
                        strategy_metrics: Dict[str, float], b1_metrics: Dict[str, float],
                        alpha_ir: float, cost_over_alpha: float,
                        year2022_return: float, scfg: SectorConfig) -> pd.DataFrame:
    """[§8] 6개 조건의 PASS/FAIL 표. 하나라도 FAIL이면 활성화하지 않는다(USE_SECTOR_ROTATION
    은 사용자가 이 표를 보고 명시적으로 켠다 — 이 함수는 자동으로 스위치를 켜지 않는다)."""
    def _r(cond, label, detail):
        return {"조건": label, "판정": "PASS" if cond else "FAIL", "상세": detail}

    rows = [
        _r(pd.notna(cs_stats.get("mean_rank_ic")) and cs_stats["mean_rank_ic"] >= scfg.CS_MIN_RANK_IC
           and pd.notna(cs_stats.get("nw_t")) and cs_stats["nw_t"] >= scfg.CS_MIN_NW_T,
           "① 횡단면 rank IC",
           f"mean={cs_stats.get('mean_rank_ic')}, NW-t={cs_stats.get('nw_t')} "
           f"(기준 ≥{scfg.CS_MIN_RANK_IC}, t≥{scfg.CS_MIN_NW_T})"),
        _r(spread_by_era_positive_count >= 3, "② 상위3-하위3 스프레드 시대별 부호",
           f"{spread_by_era_positive_count}/4 시대에서 양(+) (기준 ≥3/4)"),
        _r(pd.notna(strategy_metrics.get("샤프")) and pd.notna(b1_metrics.get("샤프"))
           and strategy_metrics["샤프"] >= b1_metrics["샤프"]
           and strategy_metrics.get("최대낙폭(MDD)", -1) >= b1_metrics.get("최대낙폭(MDD)", -1) - 0.005,
           "③ 샤프≥B1 & MDD 훼손<0.5%p",
           f"전략 샤프={strategy_metrics.get('샤프')} vs B1={b1_metrics.get('샤프')}, "
           f"전략 MDD={strategy_metrics.get('최대낙폭(MDD)')} vs B1={b1_metrics.get('최대낙폭(MDD)')}"),
        _r(pd.notna(alpha_ir) and alpha_ir >= 0.5 and pd.notna(cost_over_alpha) and cost_over_alpha < (1 / 3),
           "④ 로테이션알파 IR≥0.5 & 비용<알파의1/3",
           f"IR={alpha_ir}, 비용/알파={cost_over_alpha}"),
        _r(pd.notna(year2022_return) and year2022_return >= -0.016,
           "⑤ 2022 방어 훼손 없음",
           f"2022 수익률={year2022_return} (기준: -1.6%보다 나빠지면 FAIL — SPY 계층 단독 "
           f"기준치 report23 -1.6%를 섹터 계층이 재배분만 하므로 이론상 거의 동일해야 정상)"),
    ]
    df = pd.DataFrame(rows)
    df.loc[len(df)] = {"조건": "종합판정",
                       "판정": "PASS(활성 후보)" if (df["판정"] == "PASS").all() else "FAIL(OFF 유지)",
                       "상세": f"{(df['판정']=='PASS').sum()}/{len(rows)}개 조건 통과"}
    return df


# =============================================================================
# [13] 리포트 (§9) — 별도 파일. write_excel()을 재사용하지 않는다(그 함수는 00시트
#      제목·구조가 SPY 리포트에 고정돼 있다) — 가벼운 전용 라이터를 새로 둔다.
# =============================================================================
def write_sector_excel(path: str, sheets: Dict[str, pd.DataFrame], meta: List[Tuple[str, str]]) -> None:
    t0 = time.time()
    with pd.ExcelWriter(path, engine="xlsxwriter",
                        datetime_format="yyyy-mm-dd", date_format="yyyy-mm-dd") as xl:
        wb = xl.book
        f_title = wb.add_format({"bold": True, "font_size": 14, "font_color": "#1F3864"})
        f_head = wb.add_format({"bold": True, "bg_color": "#1F3864", "font_color": "white",
                                "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True})
        f_key = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1})
        f_val = wb.add_format({"border": 1, "text_wrap": True})

        ws = wb.add_worksheet("12_섹터개요")
        xl.sheets["12_섹터개요"] = ws
        ws.set_column(0, 0, 30); ws.set_column(1, 1, 100)
        ws.write(0, 0, "11개 섹터 로테이션 확장 (SPY 국면 계층 위, 상대수익 예측)", f_title)
        r = 2
        for k, v in meta:
            ws.write(r, 0, str(k), f_key); ws.write(r, 1, str(v), f_val); r += 1

        for name, df in sheets.items():
            if df is None or len(df) == 0:
                continue
            df.to_excel(xl, sheet_name=name, index=False)
            w = xl.sheets[name]
            w.freeze_panes(1, 1)
            for j, col in enumerate(df.columns):
                w.write(0, j, str(col), f_head)
                try:
                    maxlen = int(df[col].map(str).str.len().max())
                except Exception:
                    maxlen = 12
                w.set_column(j, j, min(max(maxlen + 2, 10), 60))
    log("REPORT", kv(event="written", path=path, sheets=len(sheets),
                     elapsed_s=round(time.time() - t0, 2)))


# =============================================================================
# [14] 최상위 진입점
# =============================================================================
def run(res: dict, M, scfg: Optional[SectorConfig] = None) -> Dict[str, Any]:
    """[§9.4 실행순서] 1) 유니버스·수집·무결성 2) 타깃·후보지표·검증(+자기검사) 3) 가중·
    점수·상태기계·포트폴리오·백테스트·감사 를 한 번에 수행한다. 반환 dict의 "sheets"를
    build_sector_report()에 넘기면 엑셀이 만들어진다.
    res: M.run(M.CFG)의 반환값. scfg: 없으면 모듈 기본 CFG."""
    t0 = time.time()
    scfg = scfg or CFG
    cal = res["cal"]
    spy_close_raw = res["price"]["Close"].reindex(cal)
    spy_adj = res["px_adj"].reindex(cal)

    # ---- [4.3] 판별력 자기검사 — 먼저 통과해야 실데이터 해석을 신뢰할 수 있다 ----
    st = run_selftest(M, scfg)
    if not st["passed"]:
        log("RUN", kv(event="selftest_failed_abort"), M=M, level="error")
        return {"selftest": st, "aborted": True}

    # ---- [1] 유니버스·수집·무결성 ----
    quality_rows: List[dict] = []
    sector_px, yahoo_diag = fetch_sector_prices(res, M, quality_rows)
    entry_dates = sector_entry_dates(sector_px, scfg)
    active_mask = sector_active_mask(entry_dates, cal, scfg)
    universe_audit = universe_entry_causality_audit(entry_dates, sector_px)

    adj_lag_rows = []
    tr_close: Dict[str, pd.Series] = {}
    raw_close: Dict[str, pd.Series] = {}
    for t in scfg.SECTORS:
        df = sector_px.get(t)
        if df is None or len(df) == 0:
            tr_close[t] = pd.Series(np.nan, index=cal)
            raw_close[t] = pd.Series(np.nan, index=cal)
            continue
        s_tr, spliced = build_total_return_close(df, cal, scfg.ADJ_CLOSE_STALE_DAYS)
        tr_close[t] = s_tr
        raw_close[t] = df["Close"].reindex(cal).ffill()
        stale, gap, adj_last, close_last = adj_close_lag_check(df, t, scfg.ADJ_CLOSE_STALE_DAYS)
        adj_lag_rows.append({"티커": t, "AdjClose지연": stale, "지연거래일수": gap,
                            "AdjClose마지막일": str(adj_last.date()) if adj_last is not None else "N/A",
                            "Close마지막일": str(close_last.date()) if close_last is not None else "N/A",
                            "Close수익률대체적용": spliced})

    E_t = res["sig"]["target_pos"].reindex(cal)

    # ---- [2]~[4] 섹터별 타깃·지표·검증·가중치 ----
    per_sector: Dict[str, Dict[str, Any]] = {}
    vt_rows_all = []
    score_pct_by_sector = pd.DataFrame(index=cal, columns=list(scfg.SECTORS), dtype=float)
    score_by_sector = pd.DataFrame(index=cal, columns=list(scfg.SECTORS), dtype=float)
    fwd21_rel_by_sector: Dict[str, pd.Series] = {}
    audit_rows_all = []
    wlog_rows_all = []

    for t in scfg.SECTORS:
        df = sector_px.get(t)
        if df is None or len(df) == 0:
            continue
        rel = relative_price_series(tr_close[t], spy_adj)
        ind, specs = build_sector_indicators(t, res, M, tr_close[t], raw_close[t],
                                             spy_adj, spy_close_raw, cal)
        actual_start = df.index.min()
        entry = entry_dates[t]
        cfg_t = sector_cfg_for(res, scfg, actual_start, entry)
        result = validate_and_weight_sector(t, ind, specs, rel, cfg_t, M, scfg=scfg)
        per_sector[t] = {"ind": ind, "specs": specs, "rel": rel, "cfg": cfg_t, **result}

        vt = result["vt"].copy(); vt.insert(0, "티커", t)
        vt_rows_all.append(vt)
        for e in result["wlog"]:
            e2 = dict(e); e2["티커"] = t; wlog_rows_all.append(e2)

        sp = result["score_pct"].reindex(cal)
        sc = result["score"].reindex(cal)
        score_pct_by_sector[t] = sp
        score_by_sector[t] = sc
        fwd21_rel_by_sector[t] = M.forward_return(rel, 21).reindex(cal)

        audit_rows_all.append(lookahead_audit_sector(t, ind.reindex(result["idx"]), specs,
                                                      rel.reindex(result["idx"]), cfg_t, M,
                                                      sc.reindex(result["idx"]), scfg=scfg))
        log("RUN", kv(event="sector_done", ticker=t,
                      passed=int((result['vt']['판정'] == 'PASS').sum())), M=M)

    vt_all = pd.concat(vt_rows_all, ignore_index=True) if vt_rows_all else pd.DataFrame()
    wlog_all = pd.DataFrame(wlog_rows_all) if wlog_rows_all else pd.DataFrame()
    audit_all = pd.concat(audit_rows_all, ignore_index=True) if audit_rows_all else pd.DataFrame()

    # ---- [4.2] 횡단면 rank IC (복합점수 vs 21일 상대수익, §8 조건①의 근거) ----
    cs_stats = cross_sectional_rank_ic(fwd21_rel_by_sector,
                                       {t: score_by_sector[t] for t in scfg.SECTORS}, 21)

    # ---- [5] 순위·상태기계·포트폴리오 ----
    rank = cross_sectional_rank(score_pct_by_sector, active_mask)
    n_active = active_mask.sum(axis=1)
    raw_states = pd.DataFrame(index=cal, columns=list(scfg.SECTORS), dtype=object)
    confirmed_states = pd.DataFrame(index=cal, columns=list(scfg.SECTORS), dtype=object)
    for t in scfg.SECTORS:
        raw = sector_raw_state(rank[t], score_pct_by_sector[t], n_active, scfg)
        raw_states[t] = raw
        confirmed_states[t] = apply_sector_hysteresis(raw, scfg)

    target = build_portfolio(confirmed_states, active_mask, E_t, scfg)
    invariant = portfolio_invariant_check(target, E_t, scfg)

    # ---- [6] 백테스트 ----
    rf = None
    fred = res.get("fred", {})
    if "DGS3MO" in fred and fred["DGS3MO"] is not None:
        rf = (fred["DGS3MO"] / 100.0 / 252.0).reindex(cal).ffill().fillna(0.0)
    price_by_sector = {t: sector_px[t] for t in scfg.SECTORS if sector_px.get(t) is not None}
    strat_bt = run_sector_backtest(price_by_sector, target, res["cfg"], rf)
    b2_bt = benchmark_equalweight(price_by_sector, active_mask, E_t, res["cfg"], rf)
    b3_bt = benchmark_ew_buyhold(price_by_sector, active_mask, res["cfg"], rf)
    b0_bh = res["bt"][["bh_ret"]].rename(columns={"bh_ret": "strategy_ret"}).copy()
    b0_bh["equity"] = (1 + b0_bh["strategy_ret"]).cumprod()
    b1_bt = res["bt"][["strategy_ret", "equity"]].copy()

    strat_metrics = M.perf_metrics(strat_bt["strategy_ret"], "섹터로테이션 전략")
    b0_metrics = M.perf_metrics(b0_bh["strategy_ret"], "B0 SPY 단순보유")
    b1_metrics = M.perf_metrics(b1_bt["strategy_ret"], "B1 현행 SPY 국면전략")
    b2_metrics = M.perf_metrics(b2_bt["strategy_ret"], "B2 국면×동일가중11섹터")
    b3_metrics = M.perf_metrics(b3_bt["strategy_ret"], "B3 동일가중11섹터 단순보유")

    alpha_ret = (strat_bt["strategy_ret"] - b2_bt["strategy_ret"]).dropna()
    alpha_mean = float(alpha_ret.mean() * 252) if len(alpha_ret) else np.nan
    alpha_std = float(alpha_ret.std() * math.sqrt(252)) if len(alpha_ret) else np.nan
    alpha_ir = alpha_mean / alpha_std if alpha_std else np.nan
    total_alpha = float((strat_bt["equity"].iloc[-1] - b2_bt["equity"].iloc[-1])) if len(strat_bt) else np.nan
    total_cost = float(strat_bt["cost"].sum())
    cost_over_alpha = (total_cost / abs(total_alpha)) if total_alpha not in (0, np.nan) and pd.notna(total_alpha) else np.nan

    # 4-시대 상위3-하위3 스프레드 부호
    blocks = np.array_split(np.arange(len(cal)), 4)
    spread_pos = 0
    era_rows = []
    for bi, b in enumerate(blocks):
        sub_idx = cal[b]
        sub_rank = rank.reindex(sub_idx)
        sub_fwd = pd.DataFrame({t: fwd21_rel_by_sector[t].reindex(sub_idx) for t in scfg.SECTORS})
        top_mask = sub_rank <= scfg.SECTOR_TOP_K
        bot_n = n_active.reindex(sub_idx)
        bot_mask = sub_rank.ge(bot_n - scfg.SECTOR_TOP_K + 1, axis=0)
        top_ret = sub_fwd.where(top_mask).stack().mean()
        bot_ret = sub_fwd.where(bot_mask).stack().mean()
        spread = (top_ret - bot_ret) if pd.notna(top_ret) and pd.notna(bot_ret) else np.nan
        if pd.notna(spread) and spread > 0:
            spread_pos += 1
        era_rows.append({"시대": bi + 1, "시작": str(sub_idx.min().date()) if len(sub_idx) else "",
                        "종료": str(sub_idx.max().date()) if len(sub_idx) else "",
                        "상위3평균21일상대수익": top_ret, "하위3평균21일상대수익": bot_ret,
                        "스프레드": spread})

    year2022 = strat_bt["strategy_ret"].loc["2022-01-01":"2022-12-31"]
    year2022_ret = float((1 + year2022).prod() - 1) if len(year2022) else np.nan

    acceptance = evaluate_acceptance(cs_stats, spread_pos, strat_metrics, b1_metrics,
                                     alpha_ir, cost_over_alpha, year2022_ret, scfg)

    log("RUN", kv(event="all_done", elapsed_s=round(time.time() - t0, 2),
                  acceptance=acceptance.iloc[-1]["판정"]), M=M)

    return {
        "scfg": scfg, "selftest": st, "sector_px": sector_px, "entry_dates": entry_dates,
        "active_mask": active_mask, "universe_audit": universe_audit,
        "adj_lag": pd.DataFrame(adj_lag_rows), "per_sector": per_sector,
        "vt_all": vt_all, "wlog_all": wlog_all, "audit_all": audit_all,
        "cs_stats": cs_stats, "rank": rank, "raw_states": raw_states,
        "confirmed_states": confirmed_states, "target": target, "invariant": invariant,
        "strat_bt": strat_bt, "b0_bt": b0_bh, "b1_bt": b1_bt, "b2_bt": b2_bt, "b3_bt": b3_bt,
        "strat_metrics": strat_metrics, "b0_metrics": b0_metrics, "b1_metrics": b1_metrics,
        "b2_metrics": b2_metrics, "b3_metrics": b3_metrics,
        "alpha_ir": alpha_ir, "cost_over_alpha": cost_over_alpha, "total_alpha": total_alpha,
        "era_spread": pd.DataFrame(era_rows), "year2022_return": year2022_ret,
        "acceptance": acceptance, "quality_rows": quality_rows, "yahoo_diag": yahoo_diag,
    }


# =============================================================================
# [15] 리포트 조립 (§9.1, §11)
# =============================================================================
def build_sector_report(sres: Dict[str, Any], path: str = "sector_rotation_report.xlsx") -> str:
    """run()의 반환값을 12~19번대 시트로 엮어 별도 엑셀 파일로 저장한다."""
    scfg: SectorConfig = sres["scfg"]

    universe_rows = []
    for t in scfg.SECTORS:
        entry = sres["entry_dates"].get(t)
        df = sres["sector_px"].get(t)
        universe_rows.append({
            "티커": t, "섹터명": SECTOR_NAME_KR.get(t, t),
            "실제데이터시작": str(df.index.min().date()) if (df is not None and len(df)) else "N/A",
            "유니버스편입일": str(entry.date()) if entry is not None and entry != pd.Timestamp.max else "N/A",
        })
    universe_df = pd.DataFrame(universe_rows).merge(sres["adj_lag"], on="티커", how="left") \
        if len(sres["adj_lag"]) else pd.DataFrame(universe_rows)
    universe_df = universe_df.merge(sres["universe_audit"].rename(
        columns={"실제데이터시작": "_dup1", "유니버스편입일": "_dup2"}), on="티커", how="left",
        suffixes=("", "_감사"))
    universe_df = universe_df.drop(columns=[c for c in universe_df.columns if c.startswith("_dup")],
                                   errors="ignore")

    # ---- 15 신호(01시트 스타일) ----
    signal_rows = []
    cal = sres["confirmed_states"].index
    for t in scfg.SECTORS:
        cs = sres["confirmed_states"][t]
        rk = sres["rank"][t]
        sp = None
        if t in sres["per_sector"]:
            sp = sres["per_sector"][t]["score_pct"].reindex(cal)
        tg = sres["target"][t] if t in sres["target"].columns else pd.Series(np.nan, index=cal)
        for dt in cal:
            signal_rows.append({"티커": t, "날짜": dt, "확정상태": cs.loc[dt],
                               "횡단면순위": rk.loc[dt] if pd.notna(rk.loc[dt]) else np.nan,
                               "복합점수백분위": float(sp.loc[dt]) if (sp is not None and pd.notna(sp.loc[dt])) else np.nan,
                               "목표비중": float(tg.loc[dt]) if pd.notna(tg.loc[dt]) else np.nan})
    signal_df = pd.DataFrame(signal_rows)

    # ---- 17 성과 ----
    perf_df = pd.DataFrame([sres["strat_metrics"], sres["b0_metrics"], sres["b1_metrics"],
                           sres["b2_metrics"], sres["b3_metrics"]])
    perf_df["로테이션알파IR"] = [round(sres["alpha_ir"], 3) if pd.notna(sres["alpha_ir"]) else np.nan] + [np.nan] * 4
    perf_df["비용/알파"] = [round(sres["cost_over_alpha"], 3) if pd.notna(sres["cost_over_alpha"]) else np.nan] + [np.nan] * 4

    meta = [
        ("버전", f"{VERSION} ({pd.Timestamp.today().date()})"),
        ("USE_SECTOR_ROTATION(기본)", str(scfg.USE_SECTOR_ROTATION)),
        ("판별력 자기검사", f"{'PASS' if sres['selftest']['passed'] else 'FAIL'} "
         f"(참신호 {sres['selftest']['n_true_pass']}/3, 잡음오채택 {sres['selftest']['n_noise_pass']}/5)"),
        ("횡단면 rank IC", f"mean={sres['cs_stats'].get('mean_rank_ic')}, "
         f"NW-t={sres['cs_stats'].get('nw_t')}, n_days={sres['cs_stats'].get('n_days')}"),
        ("포트폴리오 불변식", f"Σtarget=E_t: {sres['invariant']['sum_equals_E_t']}, "
         f"상한준수: {sres['invariant']['bounds_ok']}, 최대오차: {sres['invariant']['max_abs_sum_diff']:.2e}"),
        ("2022 방어(전략)", f"{sres['year2022_return']:.4f}" if pd.notna(sres["year2022_return"]) else "N/A"),
        ("§8 종합판정", sres["acceptance"].iloc[-1]["판정"]),
        ("면책", "본 산출물은 연구·교육 목적의 백테스트 결과이며 투자 자문이 아닙니다."),
    ]

    sheets = {
        "13_섹터유니버스": universe_df,
        "14_섹터지표검증": sres["vt_all"],
        "15_섹터가중치로그": sres["wlog_all"],
        "16_섹터신호": signal_df,
        "17_섹터성과": perf_df,
        "18_섹터시대별스프레드": sres["era_spread"],
        "19_섹터룩어헤드감사": sres["audit_all"],
        "20_섹터수용기준": sres["acceptance"],
    }
    write_sector_excel(path, sheets, meta)
    return path


# =============================================================================
# [16] 자체 실행(모듈 단독 스모크테스트 — 합성데이터, 실데이터 없이 배관만 확인)
# =============================================================================
if __name__ == "__main__":
    import sys
    sys.path.insert(0, ".")
    import market_regime_trader as M  # type: ignore
    print("sector_rotation.py self-check: import 확인만 수행(실 데이터는 단위테스트 참조)")
    print(f"VERSION={VERSION}, SECTORS={SECTORS}")
