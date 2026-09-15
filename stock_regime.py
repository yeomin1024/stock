# =============================================================================
#  stock_regime.py
#  VERSION: v0.1.0 - 2026-09-15 - [★ 신규 파일 — 4번째 계층: 개별 주식]
#    사용자 지시(REPORT59): "이제 가장 마지막 층인 개별 주식도 똑같이 예측하도록 해 일단 샘플로
#    각 산업별 대표 티커 하나씩하고 펀더멘탈, 어닝 같은것도 보도록해"
#    사용자 선택: 신규 파일로 만든다(industry_rotation.py는 이미 8,500줄) · 데이터는 yfinance.
#
#  계층 구조:  M(시장·SPY) → S(섹터 11) → I(산업 29) → **K(개별 주식 29 · 이 파일)**
#    K는 I의 각 산업 ETF마다 대표 티커 1개를 붙인 표본이다(사용자 지시 "일단 샘플로").
#
#  ── 이 파일의 v0.1.0 자리매김(중요) ────────────────────────────────────────────────────
#    M·S·I가 겪은 실패를 반복하지 않기 위해, v0.1.0은 **배분·백테스트가 아니라 예측 채점**이 본체다.
#    라이브 규칙은 의도적으로 가장 단순한 것 하나(200일선 위/아래)로 고정하고, 펀더멘탈·어닝 후보들은
#    **격자가 아니라 채점표(03 시트)에서 소수 클래스 잣대로** 먼저 측정한다. 승격은 다음 라운드 판정이다.
#    이유: I 계층에서 "그럴듯한 근거로 라이브를 바꿨다가 세 라운드 연속 실패"한 이력이 있다(v0.15~v0.17).
#
#  ── 판정 잣대(M·S·I와 동일하게 적용한다) ──────────────────────────────────────────────
#    (1) 소수 클래스 기준 — 정밀도는 **항상 같은 행의 기저율과 함께** 읽는다.
#    (2) 연도 k/n이 1급 판정 — 풀샘플 단독 금지(상관 높은 29계열은 유효표본이 작다).
#    (3) 룩어헤드 금지 — 절단재계산 감사(11 시트)로 매 실행 확인한다.
#    (4) 단면 비교는 **자기이력 백분위**로만 한다(정적 전이력 값은 정체성을 줄세울 뿐이다).
#
#  ── ★★ 펀더멘탈의 인과 처리(이 파일에서 가장 조심한 부분) ──────────────────────────────
#    yfinance의 quarterly_* 재무제표는 **기간말(period end)** 로 색인된다. 기간말 값을 그날부터 쓰면
#    실제로는 몇 주 뒤에 공시된 숫자를 미리 쓰는 것이고, 그것이 전형적인 펀더멘탈 룩어헤드다.
#    이 파일은 두 단계로 막는다:
#      (a) **실제 발표일 우선** — get_earnings_dates()로 받은 발표일이 있으면 그 분기 값의 유효일로 쓴다.
#      (b) 발표일이 없으면 **기간말 + FUND_PUBLISH_LAG_DAYS(기본 60일)** 를 유효일로 쓴다.
#          (미국 대형주 10-Q 기한이 40일, 10-K가 60일이므로 60일은 보수적인 상한이다.)
#    그리고 as-of 조인으로 "유효일 ≤ t 인 가장 최근 분기"만 t일 특성에 넣는다.
#    ⚠ '다음 어닝이 N일 안에 있다' 류의 특성은 **넣지 않는다**. 발표 예정일은 오늘 기준으로만 알 수 있어
#      과거 시점에 그것을 알았다고 가정하면 룩어헤드다(yfinance가 주는 미래 일정은 과거에 없었다).
#      쓰는 것은 '마지막 발표 이후 경과일'뿐이며 이것은 언제나 과거 정보다.
#
#  ── 대표 티커 선정 기준(사용자가 바꿀 수 있다) ──────────────────────────────────────────
#    각 산업 ETF의 성격을 대표하는 **유동성 큰 대형주 1개**. 구성 변경에 흔들리지 않도록 ETF의 현재
#    최상위 보유가 아니라 '그 산업을 설명할 때 먼저 떠오르는 이름'을 골랐다. 바꾸려면 STOCK_UNIVERSE만 고친다.
#
#  [CHANGELOG]
#  v0.1.0 (2026-09-15) 신규:
#    (K1) StockConfig · STOCK_UNIVERSE(29) · yfinance 다운로드 + 로컬 캐시
#    (K2) 가격 특성 — 전부 롤링·자기이력 expanding 백분위(ext200·mom21/63/126·vol21·dd63·park5)
#    (K3) ★ 펀더멘탈 as-of 특성 — 발표일 우선 + 60일 보수 지연, 매출/EPS YoY·마진·성장 가속
#    (K4) ★ 어닝 특성 — 서프라이즈%·발표 후 경과일·PEAD 창(발표 후 1/5/21일 수익)
#    (K5) 라이브 규칙 = ext200 > 0 (단순 기준선) · 목표비중 1/0
#    (K6) 03_예측규칙정확도 — 규칙 × 티커 × 지평(1/5/21/63) 소수 클래스 + 연도 k/n
#    (K7) 11_룩어헤드감사 — 절단재계산(가격 특성 + 펀더멘탈 as-of 둘 다)
#    (K8) 19_상승하락구간 — industry_rotation.build_up_down_segments 재사용(계층 예산 = 1.0)
# =============================================================================
from __future__ import annotations

import dataclasses
import json
import math
import os
import sys
import time
import warnings
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

VERSION = "v0.1.0"
VERSION_DATE = "2026-09-15"

# ---- 산업 ETF → 대표 티커(사용자 지시 "각 산업별 대표 티커 하나씩") ----
#   왼쪽이 I 계층의 산업 ETF, 오른쪽이 이 파일이 예측하는 개별 주식이다.
STOCK_UNIVERSE: Dict[str, str] = {
    "SOXX": "NVDA",   # 반도체
    "IGV":  "MSFT",   # 소프트웨어
    "SKYY": "ORCL",   # 클라우드
    "HACK": "CRWD",   # 사이버보안
    "IBB":  "AMGN",   # 바이오테크(대형)
    "XBI":  "INCY",   # 바이오테크(중소)
    "IHE":  "LLY",    # 제약
    "IHI":  "ISRG",   # 의료기기
    "IHF":  "UNH",    # 헬스케어 서비스
    "XRT":  "TGT",    # 소매
    "XHB":  "DHI",    # 주택건설
    "PEJ":  "BKNG",   # 레저·여행
    "CARZ": "TSLA",   # 자동차
    "PBJ":  "PEP",    # 식음료
    "KBE":  "JPM",    # 은행
    "KRE":  "USB",    # 지역은행
    "KIE":  "PGR",    # 보험
    "KCE":  "GS",     # 자본시장
    "ITA":  "RTX",    # 항공우주·방산
    "IYT":  "UNP",    # 운송
    "JETS": "DAL",    # 항공사
    "IYZ":  "VZ",     # 통신
    "FDN":  "GOOGL",  # 인터넷
    "SOCL": "META",   # 소셜미디어
    "REZ":  "AVB",    # 주거 리츠
    "XOP":  "EOG",    # 석유·가스 E&P
    "XES":  "SLB",    # 오일 서비스
    "XME":  "FCX",    # 금속·광업
    "GDX":  "NEM",    # 금광
}
STOCK_NAME_KR: Dict[str, str] = {
    "NVDA": "엔비디아", "MSFT": "마이크로소프트", "ORCL": "오라클", "CRWD": "크라우드스트라이크",
    "AMGN": "암젠", "INCY": "인사이트", "LLY": "일라이릴리", "ISRG": "인튜이티브서지컬",
    "UNH": "유나이티드헬스", "TGT": "타깃", "DHI": "DR호턴", "BKNG": "부킹홀딩스",
    "TSLA": "테슬라", "PEP": "펩시코", "JPM": "JP모간", "USB": "US뱅코프", "PGR": "프로그레시브",
    "GS": "골드만삭스", "RTX": "RTX", "UNP": "유니온퍼시픽", "DAL": "델타항공", "VZ": "버라이즌",
    "GOOGL": "알파벳", "META": "메타", "AVB": "아발론베이", "EOG": "EOG리소시스",
    "SLB": "슐럼버거", "FCX": "프리포트맥모란", "NEM": "뉴몬트",
}


@dataclass
class StockConfig:
    # ---- 데이터 ----
    START: str = "2015-01-01"          # 특성 워밍업(200일선·자기이력 백분위)에 3년 여유를 둔다
    EVAL_START: str = "2018-01-01"     # 채점 시작 — M·S·I와 같은 창
    END: Optional[str] = None          # None이면 최신
    CACHE_DIR: str = "./_stock_cache"
    CACHE_DAYS: int = 1                # 캐시 유효기간(일) — 같은 날 재실행은 재다운로드하지 않는다
    UNIVERSE: Dict[str, str] = field(default_factory=lambda: dict(STOCK_UNIVERSE))

    # ---- ★ 펀더멘탈 인과 처리(파일 헤더 참조) ----
    FUND_PUBLISH_LAG_DAYS: int = 60    # 발표일을 모를 때 기간말에 더하는 보수 지연(10-K 기한 60일)
    USE_EARNINGS_DATES: bool = True    # 실제 발표일을 우선 사용
    EARNINGS_LIMIT: int = 60           # get_earnings_dates(limit=) — 과거 분기를 최대한 받는다

    # ---- 특성 ----
    OWN_PCT_MIN_HIST: int = 250        # 자기이력 백분위 최소 관측(1년)
    HORIZONS: Tuple[int, ...] = (1, 5, 21, 63)
    VOL_WIN: int = 21
    DD_WIN: int = 63
    MA_LONG: int = 200

    # ---- 라이브 규칙(v0.1.0은 의도적으로 가장 단순한 것 하나) ----
    LIVE_RULE: str = "ext200"          # "ext200"(200일선 위) | "always"(항상 보유 · 대조군)
    #   ⚠ 펀더멘탈·어닝 후보를 라이브에 넣지 않는다 — 03 시트에서 먼저 측정한다(파일 헤더 참조).

    # ---- 19_상승하락구간 ----
    SHOW_UPDOWN_SEGMENTS: bool = True
    SEG_MIN_MOVE: float = 0.10         # 개별 주식은 변동성이 더 커서 10%로 시작한다
    SEG_MIN_DAYS: int = 3
    SEG_DETAIL_MAX: int = 6000
    SEG_RESPONSE_EPS: float = 0.02

    # ---- 감사 ----
    AUDIT_DATES: int = 6               # 절단재계산 감사 표본 날짜 수(티커당)
    AUDIT_TICKERS: int = 6             # 감사할 티커 수(전수는 느리다)

    # ---- 출력 ----
    OUT_XLSX: str = "stock_regime_report.xlsx"
    LOG_LEVEL: str = "INFO"


CFG = StockConfig()


# =============================================================================
# [0] 로깅 — M·S·I와 같은 형식 `[STAGE] [TIMESTAMP] key=value …`
# =============================================================================
_T0 = time.time()


def kv(**kw) -> str:
    out = []
    for k, v in kw.items():
        if isinstance(v, float):
            out.append(f"{k}={v:.4f}")
        else:
            out.append(f"{k}={v}")
    return " ".join(out)


def log(stage: str, msg: str, level: str = "info") -> None:
    lv = str(level).upper()
    order = {"DEBUG": 10, "INFO": 20, "WARNING": 30, "ERROR": 40}
    if order.get(lv, 20) < order.get(str(CFG.LOG_LEVEL).upper(), 20):
        return
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{lv}][STOCK:{stage}] [{ts}] (+{time.time() - _T0:6.1f}s) {msg}", flush=True)


# =============================================================================
# [1] 데이터 — yfinance 다운로드 + 로컬 캐시
# =============================================================================
def _cache_path(cfg: StockConfig, name: str) -> str:
    os.makedirs(cfg.CACHE_DIR, exist_ok=True)
    return os.path.join(cfg.CACHE_DIR, name)


def _cache_fresh(path: str, days: int) -> bool:
    if not os.path.exists(path):
        return False
    age = (time.time() - os.path.getmtime(path)) / 86400.0
    return age <= max(float(days), 0.0)


def download_prices(tickers: List[str], cfg: StockConfig) -> Dict[str, pd.DataFrame]:
    """OHLCV + 배당·분할 반영 종가. 실패한 티커는 건너뛰고 **이유를 남긴다**(조용히 빠지지 않게)."""
    out: Dict[str, pd.DataFrame] = {}
    failed: Dict[str, str] = {}
    cp = _cache_path(cfg, f"px_{cfg.START}_{cfg.END or 'now'}.pkl")
    if _cache_fresh(cp, cfg.CACHE_DAYS):
        try:
            out = pd.read_pickle(cp)
            log("DATA", kv(event="prices_from_cache", tickers=len(out), path=os.path.basename(cp)))
            return out
        except Exception as e:
            log("DATA", kv(event="cache_read_failed", err=str(e)[:120], action="재다운로드"), level="warning")
    try:
        import yfinance as yf
    except Exception as e:
        log("DATA", kv(event="yfinance_import_failed", err=str(e)[:160],
                       suggest="pip install yfinance"), level="error")
        raise
    for t in tickers:
        try:
            df = yf.Ticker(t).history(start=cfg.START, end=cfg.END, auto_adjust=True,
                                      actions=True, raise_errors=False)
            if df is None or not len(df):
                failed[t] = "빈 응답"
                continue
            df = df.rename(columns=str.title)
            df.index = pd.to_datetime(df.index).tz_localize(None).normalize()
            keep = [c for c in ["Open", "High", "Low", "Close", "Volume"] if c in df.columns]
            df = df[keep].dropna(subset=["Close"])
            if len(df) < 300:
                failed[t] = f"관측 부족({len(df)})"
                continue
            out[t] = df
        except Exception as e:
            failed[t] = f"{type(e).__name__}: {str(e)[:80]}"
    log("DATA", kv(event="prices_downloaded", ok=len(out), failed=len(failed),
                   detail=(";".join(f"{k}={v}" for k, v in failed.items()) or "-"),
                   start=cfg.START, end=(cfg.END or "now")))
    if out:
        try:
            pd.to_pickle(out, cp)
        except Exception as e:
            log("DATA", kv(event="cache_write_failed", err=str(e)[:100]), level="warning")
    return out


def download_fundamentals(tickers: List[str], cfg: StockConfig) -> Dict[str, Dict[str, Any]]:
    """분기 재무제표 + **실제 발표일**. 반환 구조:
      {ticker: {"q": DataFrame(기간말 × 항목), "earn": DataFrame(발표일 index, EPS·서프라이즈)}}
    ⚠ 여기서는 **아무 것도 가공하지 않는다** — 인과 처리는 build_fundamental_asof가 전담한다."""
    out: Dict[str, Dict[str, Any]] = {}
    cp = _cache_path(cfg, "fund.pkl")
    if _cache_fresh(cp, max(cfg.CACHE_DAYS, 7)):     # 재무제표는 분기마다 바뀌므로 1주 캐시
        try:
            out = pd.read_pickle(cp)
            log("DATA", kv(event="fundamentals_from_cache", tickers=len(out)))
            return out
        except Exception as e:
            log("DATA", kv(event="fund_cache_read_failed", err=str(e)[:120]), level="warning")
    import yfinance as yf
    n_q = n_e = 0
    problems: Dict[str, str] = {}
    for t in tickers:
        rec: Dict[str, Any] = {"q": pd.DataFrame(), "earn": pd.DataFrame()}
        tk = yf.Ticker(t)
        # (1) 분기 손익계산서 — 항목명이 버전·티커마다 달라 방어적으로 고른다
        try:
            q = None
            for attr in ("quarterly_income_stmt", "quarterly_financials"):
                try:
                    cand = getattr(tk, attr)
                except Exception:
                    cand = None
                if isinstance(cand, pd.DataFrame) and len(cand.columns):
                    q = cand
                    break
            if isinstance(q, pd.DataFrame) and len(q.columns):
                qq = q.T
                qq.index = pd.to_datetime(qq.index, errors="coerce")
                qq = qq[qq.index.notna()].sort_index()
                rec["q"] = qq
                n_q += 1
        except Exception as e:
            problems[t] = f"q:{type(e).__name__}"
        # (2) 실제 발표일 + 서프라이즈
        if cfg.USE_EARNINGS_DATES:
            try:
                e = tk.get_earnings_dates(limit=int(cfg.EARNINGS_LIMIT))
                if isinstance(e, pd.DataFrame) and len(e):
                    e = e.copy()
                    e.index = pd.to_datetime(e.index, errors="coerce")
                    try:
                        e.index = e.index.tz_localize(None)
                    except Exception:
                        pass
                    e = e[e.index.notna()].sort_index()
                    e.index = e.index.normalize()
                    rec["earn"] = e
                    n_e += 1
            except Exception as ex:
                problems[t] = problems.get(t, "") + f" earn:{type(ex).__name__}"
        out[t] = rec
    log("DATA", kv(event="fundamentals_downloaded", tickers=len(out), with_quarterly=n_q,
                   with_earnings_dates=n_e,
                   problems=(";".join(f"{k}={v.strip()}" for k, v in problems.items()) or "-"),
                   note="발표일이 없는 티커는 기간말 + FUND_PUBLISH_LAG_DAYS로 인과 처리한다"))
    try:
        pd.to_pickle(out, cp)
    except Exception as e:
        log("DATA", kv(event="fund_cache_write_failed", err=str(e)[:100]), level="warning")
    return out


# =============================================================================
# [2] 특성 — 전부 t일까지의 정보로만 만든다
#     규약: 단면 비교가 필요한 값은 **자기이력 expanding 백분위**로 정규화한다
#           (정적 전이력 값으로 줄세우면 '정체성'을 줄세우는 셈 — I 계층 라운드53의 교훈).
# =============================================================================
def _own_pct(s: pd.Series, min_hist: int) -> pd.Series:
    """자기이력 expanding 백분위 — t일까지의 자기 분포에서 오늘 값의 위치(0~1). 룩어헤드 없음."""
    return pd.Series(s).astype(float).expanding(min_periods=int(min_hist)).rank(pct=True)


def build_price_features(df: pd.DataFrame, cfg: StockConfig) -> pd.DataFrame:
    """[K2] 가격 특성. 입력은 한 티커의 OHLCV, 출력은 같은 색인의 특성 표.
    모든 열이 **그날까지의 값만** 쓴다(rolling·expanding·shift만 사용 · center=False)."""
    px = pd.to_numeric(df["Close"], errors="coerce")
    hi = pd.to_numeric(df.get("High", px), errors="coerce")
    lo = pd.to_numeric(df.get("Low", px), errors="coerce")
    r = px.pct_change()
    mh = int(cfg.OWN_PCT_MIN_HIST)
    out = pd.DataFrame(index=df.index)
    out["종가"] = px
    out["일간수익"] = r
    ma = px.rolling(int(cfg.MA_LONG), min_periods=int(cfg.MA_LONG * 0.75)).mean()
    out["ext200"] = px / ma - 1.0
    out["ext200_pct"] = _own_pct(out["ext200"], mh)
    for w in (21, 63, 126):
        out[f"mom{w}"] = px / px.shift(w) - 1.0
        out[f"mom{w}_pct"] = _own_pct(out[f"mom{w}"], mh)
    vol = r.rolling(int(cfg.VOL_WIN), min_periods=int(cfg.VOL_WIN * 0.75)).std() * math.sqrt(252.0)
    out["vol21"] = vol
    out["vol21_pct"] = _own_pct(vol, mh)
    roll_max = px.rolling(int(cfg.DD_WIN), min_periods=int(cfg.DD_WIN * 0.5)).max()
    out["dd63"] = px / roll_max - 1.0
    out["dd63_pct"] = _own_pct(out["dd63"], mh)
    # Parkinson 변동성(고가-저가) — 종가 변동성이 못 보는 장중 폭을 잡는다
    with np.errstate(invalid="ignore", divide="ignore"):
        park = np.log(hi / lo) ** 2
    park5 = pd.Series(park, index=df.index).rolling(5, min_periods=3).mean()
    out["park5"] = np.sqrt(park5 / (4.0 * math.log(2.0)) * 252.0)
    out["park5_pct"] = _own_pct(out["park5"], mh)
    out["거래대금21_pct"] = _own_pct((px * pd.to_numeric(df.get("Volume", np.nan), errors="coerce"))
                                  .rolling(21, min_periods=15).mean(), mh)
    return out


# ---- 손익계산서 항목명 후보(yfinance 버전·티커마다 다르다 → 순서대로 찾는다) ----
_REV_KEYS = ("Total Revenue", "TotalRevenue", "Operating Revenue", "Revenues")
_OPI_KEYS = ("Operating Income", "OperatingIncome", "Total Operating Income As Reported")
_NI_KEYS = ("Net Income", "NetIncome", "Net Income Common Stockholders",
            "Net Income Continuous Operations")
_EPS_KEYS = ("Diluted EPS", "DilutedEPS", "Basic EPS", "BasicEPS")


def _pick(df: pd.DataFrame, keys: Tuple[str, ...], fuzzy: Tuple[str, ...] = ()
          ) -> Tuple[Optional[pd.Series], str]:
    """항목을 고르고 **무엇을 골랐는지 이름까지 돌려준다**(02 시트에 남겨 첫 실행에서 확인할 수 있게).

    ⚠ 왜 퍼지 매칭까지 두는가: yfinance 버전·티커마다 손익계산서 항목 이름이 다르다
      ("Total Revenue" / "TotalRevenue" / "Operating Revenue" …). 정확 매칭만 두면 어떤 티커에서
      **조용히 NaN**이 되고, 그러면 펀더멘탈 규칙이 '측정 불가'인데 '효과 없음'처럼 보인다.
      그래서 (1) 정확 매칭 → (2) 부분 문자열 매칭 순으로 찾고 고른 이름을 로그·시트에 남긴다."""
    for k in keys:
        if k in df.columns:
            s = pd.to_numeric(df[k], errors="coerce")
            if s.notna().sum() >= 4:
                return s, str(k)
    for pat in (fuzzy or ()):
        for c in df.columns:
            if str(pat).lower() in str(c).lower():
                s = pd.to_numeric(df[c], errors="coerce")
                if s.notna().sum() >= 4:
                    return s, f"{c}(퍼지:{pat})"
    return None, "-"


def build_fundamental_asof(rec: Dict[str, Any], idx: pd.DatetimeIndex, cfg: StockConfig
                           ) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """[K3 ★] 분기 재무를 **유효일(발표일) 기준 as-of**로 일별 특성으로 펼친다.

    반환 (일별 특성, 분기 원장). 분기 원장에는 기간말·유효일·유효일 출처가 남아 감사가 가능하다.

    ★ 인과 처리(이 함수가 이 파일의 핵심 안전장치다):
      유효일 = (a) 그 분기에 대응하는 **실제 발표일**이 있으면 그 날
               (b) 없으면 **기간말 + FUND_PUBLISH_LAG_DAYS**
      그리고 as-of 조인으로 '유효일 ≤ t'인 가장 최근 분기만 t일에 넣는다.
      (a)의 대응 규칙: 기간말 이후 **가장 가까운** 발표일. 기간말보다 이른 발표일은 그 분기 것이 아니다.
      ⚠ 'YoY'는 4분기 전 대비다 — 분기 수가 5개 미만이면 성장률 열은 전부 NaN으로 남긴다(추정하지 않는다)."""
    q = rec.get("q")
    earn = rec.get("earn")
    empty = pd.DataFrame(index=idx)
    ledger = pd.DataFrame()
    # ⚠ 방향 방어 — download_fundamentals가 이미 전치해서 (기간말 × 항목)으로 넘기지만, yfinance 버전에
    #   따라 이미 그 방향으로 오는 경우가 있다. 색인이 날짜처럼 보이지 않으면 한 번 더 전치한다.
    if isinstance(q, pd.DataFrame) and len(q):
        _idx_dt = pd.to_datetime(pd.Series(q.index), errors="coerce").notna().mean()
        _col_dt = pd.to_datetime(pd.Series(q.columns), errors="coerce").notna().mean()
        if _col_dt > 0.8 and _idx_dt < 0.5:
            q = q.T
            q.index = pd.to_datetime(q.index, errors="coerce")
            q = q[q.index.notna()].sort_index()
    if not isinstance(q, pd.DataFrame) or not len(q):
        for c in ["매출_YoY", "EPS_YoY", "영업이익률", "영업이익률_변화", "매출성장_가속",
                  "펀더멘탈_분기수"]:
            empty[c] = np.nan
        return empty, ledger
    rev, _n_rev = _pick(q, _REV_KEYS, ("revenue", "sales"))
    opi, _n_opi = _pick(q, _OPI_KEYS, ("operating income",))
    ni, _n_ni = _pick(q, _NI_KEYS, ("net income",))
    eps, _n_eps = _pick(q, _EPS_KEYS, ("eps", "per share"))
    picked = f"매출={_n_rev} · 영업이익={_n_opi} · 순이익={_n_ni} · EPS={_n_eps}"
    L = pd.DataFrame(index=q.index)
    L["매출"] = rev if rev is not None else np.nan
    L["영업이익"] = opi if opi is not None else np.nan
    L["순이익"] = ni if ni is not None else np.nan
    L["EPS"] = eps if eps is not None else np.nan
    L = L.sort_index()
    L["기간말"] = L.index
    # ---- 유효일 ----
    lag = pd.Timedelta(days=int(cfg.FUND_PUBLISH_LAG_DAYS))
    eff: List[pd.Timestamp] = []
    src: List[str] = []
    e_idx = (earn.index if isinstance(earn, pd.DataFrame) and len(earn) else pd.DatetimeIndex([]))
    for pe in L.index:
        cand = e_idx[e_idx >= pe]
        if len(cand):
            d = pd.Timestamp(cand[0])
            # 발표일이 기간말보다 150일 넘게 뒤면 짝이 잘못 맞은 것으로 보고 지연 규칙을 쓴다
            if (d - pe).days <= 150:
                eff.append(d); src.append("발표일")
                continue
        eff.append(pd.Timestamp(pe) + lag); src.append(f"기간말+{int(cfg.FUND_PUBLISH_LAG_DAYS)}일")
    L["유효일"] = pd.DatetimeIndex(eff)
    L["유효일 출처"] = src
    L["선택 항목"] = picked          # 어떤 열을 골랐는지 원장에 남긴다(02 시트가 이걸 요약한다)
    # ---- 분기 파생(원장 안에서 계산 — 전부 과거 분기만 쓴다) ----
    L["매출_YoY"] = L["매출"] / L["매출"].shift(4) - 1.0
    L["EPS_YoY"] = L["EPS"] / L["EPS"].shift(4).abs().replace(0, np.nan) - 1.0
    with np.errstate(invalid="ignore", divide="ignore"):
        L["영업이익률"] = L["영업이익"] / L["매출"].replace(0, np.nan)
    L["영업이익률_변화"] = L["영업이익률"] - L["영업이익률"].shift(4)
    L["매출성장_가속"] = L["매출_YoY"] - L["매출_YoY"].shift(1)
    # ---- as-of 조인 ----
    feats = ["매출_YoY", "EPS_YoY", "영업이익률", "영업이익률_변화", "매출성장_가속"]
    LL = L.dropna(subset=["유효일"]).sort_values("유효일")
    base = pd.DataFrame({"날짜": pd.DatetimeIndex(idx)}).sort_values("날짜")
    merged = pd.merge_asof(base, LL[["유효일"] + feats].rename(columns={"유효일": "날짜"}),
                           on="날짜", direction="backward")
    merged = merged.set_index("날짜").reindex(idx)
    merged["펀더멘탈_분기수"] = pd.merge_asof(
        base, pd.DataFrame({"날짜": LL["유효일"].values,
                            "펀더멘탈_분기수": np.arange(1, len(LL) + 1)}).sort_values("날짜"),
        on="날짜", direction="backward").set_index("날짜").reindex(idx)["펀더멘탈_분기수"]
    # 분기 수가 5개 미만인 구간은 YoY를 만들 수 없다 — 값이 있으면 지운다(추정 금지)
    thin = merged["펀더멘탈_분기수"].fillna(0) < 5
    for c in ("매출_YoY", "EPS_YoY", "영업이익률_변화"):
        merged.loc[thin, c] = np.nan
    mh = int(cfg.OWN_PCT_MIN_HIST)
    for c in feats:
        merged[f"{c}_pct"] = _own_pct(merged[c], mh)
    return merged, L


def build_earnings_features(rec: Dict[str, Any], idx: pd.DatetimeIndex, px: pd.Series,
                            cfg: StockConfig) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """[K4] 어닝 이벤트 특성 + 이벤트 원장.
    쓰는 것: **서프라이즈%**(직전 발표), **발표 후 경과일**, 발표일 당일 갭.
    ⚠ '다음 발표까지 남은 일수'는 쓰지 않는다 — 과거 시점에 미래 일정을 알았다고 가정하는 룩어헤드다.
      (yfinance가 주는 예정일은 '오늘' 기준이며 과거에 그 값이 있었다는 보장이 없다.)"""
    out = pd.DataFrame(index=idx)
    for c in ("어닝_서프라이즈%", "어닝_경과일", "어닝_갭%", "어닝_서프라이즈%_pct"):
        out[c] = np.nan
    earn = rec.get("earn")
    ledger = pd.DataFrame()
    if not isinstance(earn, pd.DataFrame) or not len(earn):
        return out, ledger
    e = earn.copy()
    scol = next((c for c in e.columns if "Surprise" in str(c)), None)
    rcol = next((c for c in e.columns if "Reported" in str(c)), None)
    ecol = next((c for c in e.columns if "Estimate" in str(c)), None)
    L = pd.DataFrame(index=e.index)
    L["발표일"] = e.index
    L["예상EPS"] = pd.to_numeric(e[ecol], errors="coerce") if ecol else np.nan
    L["실제EPS"] = pd.to_numeric(e[rcol], errors="coerce") if rcol else np.nan
    L["서프라이즈%"] = pd.to_numeric(e[scol], errors="coerce") if scol else np.nan
    # 발표일이 **과거**인 행만 이벤트로 본다(미래 예정 행은 원장에만 남기고 특성에는 안 쓴다)
    L["과거여부"] = L.index <= (idx[-1] if len(idx) else L.index[-1])
    p = pd.Series(px).reindex(idx).astype(float)
    # 발표일 당일 갭(전일 종가 → 당일 종가) — 발표가 장전/장후인지 모르므로 '당일 수익'으로 근사한다
    day_ret = p.pct_change()
    gaps = {}
    for d in L.index:
        if d in day_ret.index and pd.notna(day_ret.get(d)):
            gaps[d] = float(day_ret.loc[d])
    L["발표일 수익%"] = [100.0 * gaps.get(d, np.nan) for d in L.index]
    # PEAD 창 — 발표 후 1/5/21일 누적 수익(원장 전용 진단. 특성이 아니다)
    for h in (1, 5, 21):
        vals = []
        for d in L.index:
            if d not in p.index:
                vals.append(np.nan); continue
            i = p.index.get_loc(d)
            j = min(i + h, len(p) - 1)
            vals.append(100.0 * float(p.iloc[j] / p.iloc[i] - 1.0) if i < len(p) - 1 else np.nan)
        L[f"발표후{h}일%"] = vals
    ledger = L.reset_index(drop=True)
    past = L[L["과거여부"]].sort_index()
    if len(past):
        base = pd.DataFrame({"날짜": pd.DatetimeIndex(idx)}).sort_values("날짜")
        m = pd.merge_asof(base,
                          pd.DataFrame({"날짜": pd.DatetimeIndex(past.index),
                                        "어닝_서프라이즈%": past["서프라이즈%"].values,
                                        "_last": pd.DatetimeIndex(past.index)}).sort_values("날짜"),
                          on="날짜", direction="backward").set_index("날짜").reindex(idx)
        out["어닝_서프라이즈%"] = m["어닝_서프라이즈%"]
        _last = m["_last"]
        out["어닝_경과일"] = (pd.Series(idx, index=idx) - _last).dt.days
        out["어닝_갭%"] = [100.0 * gaps.get(d, np.nan) if d in gaps else np.nan for d in idx]
        out["어닝_갭%"] = pd.Series(out["어닝_갭%"].values, index=idx).ffill()
        out["어닝_서프라이즈%_pct"] = _own_pct(out["어닝_서프라이즈%"], int(cfg.OWN_PCT_MIN_HIST))
    return out, ledger


# =============================================================================
# [3] 예측 규칙 · 소수 클래스 채점
#     ⚠ v0.1.0에서 라이브는 ext200 하나뿐이다. 나머지는 전부 **후보**이고 03 시트가 채점만 한다.
# =============================================================================
def _fwd_ret(px: pd.Series, h: int) -> pd.Series:
    """t일 확정 → t+1 시가 체결 관행에 맞춘 향후 h일 수익(M·S·I와 같은 정의)."""
    p = pd.Series(px).astype(float)
    return p.shift(-(h + 1)) / p.shift(-1) - 1.0


def _minority(pred: pd.Series, real: pd.Series) -> Dict[str, Any]:
    """소수 클래스 잣대 — 정밀도는 **반드시 기저율과 함께** 본다."""
    p = pd.Series(pred).astype(bool)
    y = pd.Series(real).astype(bool)
    m = p.index.intersection(y.index)
    p, y = p.loc[m], y.loc[m]
    tp = int((p & y).sum()); fp = int((p & ~y).sum())
    tn = int((~p & ~y).sum()); fn = int((~p & y).sum())
    den = float((tp + fp) * (tp + fn) * (tn + fp) * (tn + fn))
    rec = tp / (tp + fn) if (tp + fn) else np.nan
    spc = tn / (tn + fp) if (tn + fp) else np.nan
    return {"n": len(p), "n_pred": int(p.sum()), "base": (float(y.mean()) if len(y) else np.nan),
            "prec": (tp / (tp + fp) if (tp + fp) else np.nan), "rec": rec,
            "bal": (0.5 * (rec + spc) if pd.notna(rec) and pd.notna(spc) else np.nan),
            "mcc": ((tp * tn - fp * fn) / math.sqrt(den) if den > 0 else np.nan)}


def _yearly_kn(pred: pd.Series, real: pd.Series, min_days: int = 40) -> Tuple[int, int]:
    """연도별 '정밀도 > 기저' 횟수 — 1급 판정 지표(풀샘플 단독 금지)."""
    df = pd.DataFrame({"p": pd.Series(pred).astype(bool), "y": pd.Series(real).astype(bool)}).dropna()
    w = n = 0
    for _, g in df.groupby(df.index.year):
        if len(g) < min_days or not bool(g["p"].any()):
            continue
        n += 1
        w += int(float(g.loc[g["p"], "y"].mean()) > float(g["y"].mean()))
    return w, n


def rule_matrix(feat: pd.DataFrame, cfg: StockConfig) -> Dict[str, pd.Series]:
    """예측 규칙 후보 — **하락**을 맞추려는 규칙 집합(사용자 요구가 하락 정확도다).
    각 규칙은 '앞으로 내릴 것 같다 = True'를 뜻한다. 전부 t일까지의 정보다.

    가격 후보(P)는 I 계층에서 이미 측정된 것을 이식한다 — 특히 **vol21 하위 1/3**은
    I의 13p A3에서 정밀−기저 +0.14를 낸 유일한 산업별 하락 신호였다.
    펀더멘탈·어닝 후보(F·E)가 이 파일의 새로운 부분이다."""
    R: Dict[str, pd.Series] = {}
    g = lambda c: (pd.to_numeric(feat[c], errors="coerce") if c in feat.columns
                   else pd.Series(np.nan, index=feat.index))
    # ---- 가격 ----
    R["P1 200일선 아래"] = g("ext200") < 0
    R["P2 vol21 자기이력 하위1/3"] = g("vol21_pct") <= 1.0 / 3.0
    R["P3 vol21 자기이력 상위1/5"] = g("vol21_pct") >= 0.8
    R["P4 mom63 하위1/3"] = g("mom63_pct") <= 1.0 / 3.0
    R["P5 park5 상위1/5"] = g("park5_pct") >= 0.8
    R["P6 P1 & P2"] = (g("ext200") < 0) & (g("vol21_pct") <= 1.0 / 3.0)
    # ---- 펀더멘탈 ----
    R["F1 매출 YoY < 0"] = g("매출_YoY") < 0
    R["F2 EPS YoY < 0"] = g("EPS_YoY") < 0
    R["F3 영업이익률 4분기 악화"] = g("영업이익률_변화") < 0
    R["F4 매출성장 감속"] = g("매출성장_가속") < 0
    R["F5 F1 & F3"] = (g("매출_YoY") < 0) & (g("영업이익률_변화") < 0)
    R["F6 P1 & F1(추세+펀더)"] = (g("ext200") < 0) & (g("매출_YoY") < 0)
    # ---- 어닝 ----
    R["E1 직전 서프라이즈 < 0"] = g("어닝_서프라이즈%") < 0
    R["E2 서프라이즈 자기이력 하위1/3"] = g("어닝_서프라이즈%_pct") <= 1.0 / 3.0
    R["E3 발표 후 21일 이내 & 서프라이즈<0"] = (g("어닝_경과일") <= 21) & (g("어닝_서프라이즈%") < 0)
    R["E4 발표일 수익 < −3%"] = g("어닝_갭%") < -3.0
    R["E5 E1 & P1"] = (g("어닝_서프라이즈%") < 0) & (g("ext200") < 0)
    R["E6 E1 & F1"] = (g("어닝_서프라이즈%") < 0) & (g("매출_YoY") < 0)
    return {k: pd.Series(v).astype("boolean") for k, v in R.items()}


def build_rule_accuracy(panel: Dict[str, pd.DataFrame], cfg: StockConfig) -> pd.DataFrame:
    """[03_예측규칙정확도] 규칙 × 티커 × 지평의 소수 클래스 표 + 규칙별 29종목 집계.
    '정밀도'는 늘 '기저 실현하락률'과 같은 행에 있다(사용자 규칙 1)."""
    rows: List[dict] = []
    rows.append({"블록": "── 읽는 법 ──", "규칙": "",
                 "판독": ("각 규칙은 '앞으로 내릴 것 같다'를 뜻한다. **정밀도를 반드시 같은 행의 "
                        "'기저 실현하락률'과 비교**할 것 — 기저보다 높아야 예측력이 있다. 판정은 "
                        "**연도 k/n**으로 한다(풀샘플은 상관 높은 29계열을 한 덩어리로 세어 부풀려진다). "
                        "P는 가격(I 계층에서 이식), F는 펀더멘탈, E는 어닝 후보다. "
                        "⚠ v0.1.0에서 라이브는 P1(200일선)만이고 나머지는 전부 후보다 — "
                        "승격은 연도 k/n과 대조군을 갖춘 다음 라운드의 일이다.")})
    agg: Dict[Tuple[str, int], Dict[str, List[float]]] = {}
    for t, feat in panel.items():
        rules = rule_matrix(feat, cfg)
        px = pd.to_numeric(feat["종가"], errors="coerce")
        for h in cfg.HORIZONS:
            fwd = _fwd_ret(px, int(h))
            real = (fwd < 0)
            ok = fwd.notna()
            for name, pr in rules.items():
                p = pr.reindex(feat.index).fillna(False).astype(bool) & ok
                if int(p.sum()) < 20:
                    continue
                m = _minority(p[ok], real[ok])
                w, n = _yearly_kn(p[ok], real[ok])
                gap = (float(fwd[ok & p].mean()) - float(fwd[ok & ~p].mean())) if int(p.sum()) else np.nan
                rows.append({"블록": "A. 티커별", "규칙": name, "티커": t,
                             "이름": STOCK_NAME_KR.get(t, ""), "지평(일)": int(h),
                             "예측일수": m["n_pred"], "표본일수": m["n"],
                             "기저 실현하락률": round(m["base"], 4),
                             "하락 정밀도": (round(m["prec"], 4) if pd.notna(m["prec"]) else None),
                             "정밀−기저": (round(m["prec"] - m["base"], 4)
                                       if pd.notna(m["prec"]) and pd.notna(m["base"]) else None),
                             "하락 재현율": (round(m["rec"], 4) if pd.notna(m["rec"]) else None),
                             "균형정확도": (round(m["bal"], 4) if pd.notna(m["bal"]) else None),
                             "MCC": (round(m["mcc"], 4) if pd.notna(m["mcc"]) else None),
                             "연도 정밀도>기저": f"{w}/{n}",
                             "연도비율": (round(w / n, 3) if n else None),
                             "향후수익 격차(%p)": (round(gap * 100, 3) if pd.notna(gap) else None)})
                k = (name, int(h))
                a = agg.setdefault(k, {"pmb": [], "base": [], "prec": [], "w": [], "n": [],
                                       "mcc": [], "np": [], "gap": []})
                if pd.notna(m["prec"]) and pd.notna(m["base"]):
                    a["pmb"].append(m["prec"] - m["base"])
                a["base"].append(m["base"]); a["prec"].append(m["prec"])
                a["w"].append(w); a["n"].append(n); a["mcc"].append(m["mcc"])
                a["np"].append(m["n_pred"])
                if pd.notna(gap):
                    a["gap"].append(gap)
    for (name, h), a in sorted(agg.items(), key=lambda x: (x[0][1], x[0][0])):
        if not a["pmb"]:
            continue
        w, n = int(np.nansum(a["w"])), int(np.nansum(a["n"]))
        rows.append({"블록": "B. 규칙별 29종목 집계", "규칙": name, "티커": f"{len(a['pmb'])}종목",
                     "지평(일)": int(h), "예측일수": int(np.nanmean(a["np"])),
                     "기저 실현하락률": round(float(np.nanmean(a["base"])), 4),
                     "하락 정밀도": round(float(np.nanmean(a["prec"])), 4),
                     "정밀−기저": round(float(np.nanmean(a["pmb"])), 4),
                     "MCC": round(float(np.nanmean(a["mcc"])), 4),
                     "정밀>기저 종목": f"{int(sum(1 for v in a['pmb'] if v > 0))}/{len(a['pmb'])}",
                     "연도 정밀도>기저": f"{w}/{n}",
                     "연도비율": (round(w / n, 3) if n else None),
                     "향후수익 격차(%p)": (round(float(np.nanmean(a["gap"])) * 100, 3) if a["gap"] else None)})
    rows.append({"블록": "── 판정(사전등록) ──", "규칙": "",
                 "판독": ("v0.2.0 승격 조건 — 후보가 **셋을 동시에** 만족해야 라이브에 올린다: "
                        "(1) B블록 연도비율 ≥ 0.60  (2) 정밀>기저 종목 ≥ 20/29  "
                        "(3) 향후수익 격차 < 0(경고일의 실제 수익이 더 낮아야 쓸 값이 있다). "
                        "⚠ 연구·교육용이며 투자 조언이 아니다.")})
    return pd.DataFrame(rows)


# =============================================================================
# [4] 라이브 규칙 · 목표비중
# =============================================================================
def build_positions(feat: pd.DataFrame, cfg: StockConfig) -> pd.DataFrame:
    """[K5] 라이브 규칙은 v0.1.0에서 **ext200 > 0** 하나다(가장 단순한 기준선).
    체결 규칙은 M·S·I와 같다 — t일 확정, t+1일 집행(exec_w(t) = target_w(t−1))."""
    out = pd.DataFrame(index=feat.index)
    rule = str(getattr(cfg, "LIVE_RULE", "ext200")).lower()
    ext = pd.to_numeric(feat.get("ext200"), errors="coerce")
    if rule == "always":
        tgt = pd.Series(1.0, index=feat.index)
        state = pd.Series("상승(항상보유)", index=feat.index)
    else:
        up = (ext > 0)
        tgt = up.astype(float).where(ext.notna(), np.nan).ffill().fillna(0.0)
        state = pd.Series(np.where(up.fillna(False), "상승", "하락"), index=feat.index)
    out["확정국면"] = state
    out["목표비중"] = tgt
    out["집행비중"] = tgt.shift(1).fillna(0.0)
    out["전략일간수익"] = out["집행비중"] * pd.to_numeric(feat["일간수익"], errors="coerce").fillna(0.0)
    return out


# =============================================================================
# [5] 룩어헤드 감사 — 절단재계산(M·S·I와 같은 방식)
# =============================================================================
def build_lookahead_audit(prices: Dict[str, pd.DataFrame], fund: Dict[str, Dict[str, Any]],
                          cfg: StockConfig) -> pd.DataFrame:
    """[11_룩어헤드감사] 검사일 t까지만 잘라서 다시 계산한 값이 전체 계산과 같은지 확인한다.
    가격 특성과 **펀더멘탈 as-of** 둘 다 검사한다 — 후자가 이 파일에서 가장 위험한 지점이다."""
    rows: List[dict] = []
    tks = list(prices)[: int(cfg.AUDIT_TICKERS)]
    cols = ["ext200_pct", "vol21_pct", "dd63_pct", "mom63_pct"]
    fcols = ["매출_YoY", "EPS_YoY", "영업이익률"]
    for t in tks:
        df = prices[t]
        full = build_price_features(df, cfg)
        ffull, _ = build_fundamental_asof(fund.get(t, {}), df.index, cfg)
        idx = full.index[full.index >= pd.Timestamp(cfg.EVAL_START)]
        if len(idx) < 50:
            continue
        picks = list(pd.Series(idx).iloc[np.linspace(10, len(idx) - 1, num=int(cfg.AUDIT_DATES),
                                                    dtype=int)])
        for d in picks:
            cut = df.loc[:d]
            tr = build_price_features(cut, cfg)
            ftr, _ = build_fundamental_asof(fund.get(t, {}), cut.index, cfg)
            for c in cols:
                a = float(full[c].get(d, np.nan)); b = float(tr[c].get(d, np.nan))
                if pd.isna(a) and pd.isna(b):
                    continue
                rows.append({"티커": t, "검사일": d, "항목": c, "전체계산": a, "절단재계산": b,
                             "차이": (abs(a - b) if pd.notna(a) and pd.notna(b) else np.nan),
                             "일치": ("OK" if (pd.notna(a) and pd.notna(b) and abs(a - b) < 1e-9)
                                    else ("OK(둘 다 NaN)" if pd.isna(a) and pd.isna(b) else "⚠ 불일치")),
                             "감사종류": "가격 특성"})
            for c in fcols:
                if c not in ffull.columns or c not in ftr.columns:
                    continue
                a = float(ffull[c].get(d, np.nan)); b = float(ftr[c].get(d, np.nan))
                if pd.isna(a) and pd.isna(b):
                    rows.append({"티커": t, "검사일": d, "항목": c, "전체계산": np.nan,
                                 "절단재계산": np.nan, "차이": np.nan, "일치": "OK(둘 다 NaN)",
                                 "감사종류": "펀더멘탈 as-of"})
                    continue
                rows.append({"티커": t, "검사일": d, "항목": c, "전체계산": a, "절단재계산": b,
                             "차이": (abs(a - b) if pd.notna(a) and pd.notna(b) else np.nan),
                             "일치": ("OK" if (pd.notna(a) and pd.notna(b) and abs(a - b) < 1e-9)
                                    else "⚠ 불일치"),
                             "감사종류": "펀더멘탈 as-of"})
    out = pd.DataFrame(rows)
    bad = int((out["일치"].astype(str).str.startswith("⚠")).sum()) if len(out) else 0
    log("AUDIT", kv(event="lookahead_audit", checks=len(out), mismatch=bad,
                    note="펀더멘탈 as-of는 발표일(또는 기간말+지연) 기준이라 절단해도 값이 같아야 한다"),
        level=("warning" if bad else "info"))
    return out


# =============================================================================
# [6] 실행 · 리포트
# =============================================================================
def run(cfg: Optional[StockConfig] = None, s_overrides: Optional[Dict[str, Any]] = None
        ) -> Dict[str, Any]:
    """[K 계층 본체] 데이터 → 특성 → 규칙 채점 → 포지션 → 감사 → 리포트 dict."""
    cfg = cfg or CFG
    if s_overrides:
        cfg = dataclasses.replace(cfg, **s_overrides)
        log("RUN", kv(event="overrides", detail=json.dumps(s_overrides, ensure_ascii=False)[:300]))
    uni = dict(cfg.UNIVERSE)
    tickers = sorted(set(uni.values()))
    log("RUN", kv(event="start", version=VERSION, tickers=len(tickers),
                  eval_start=cfg.EVAL_START, live_rule=cfg.LIVE_RULE,
                  fund_lag_days=cfg.FUND_PUBLISH_LAG_DAYS))
    prices = download_prices(tickers, cfg)
    if not prices:
        log("RUN", kv(event="aborted", reason="가격 데이터 0건"), level="error")
        return {"aborted": True, "note": "가격 데이터를 하나도 받지 못했다", "cfg": cfg}
    fund = download_fundamentals(list(prices), cfg)

    panel: Dict[str, pd.DataFrame] = {}
    pos: Dict[str, pd.DataFrame] = {}
    fund_ledgers: List[pd.DataFrame] = []
    earn_ledgers: List[pd.DataFrame] = []
    quality: List[dict] = []
    parent_of = {v: k for k, v in uni.items()}
    for t, df in prices.items():
        pf = build_price_features(df, cfg)
        ff, fl = build_fundamental_asof(fund.get(t, {}), df.index, cfg)
        ef, el = build_earnings_features(fund.get(t, {}), df.index, pf["종가"], cfg)
        feat = pd.concat([pf, ff, ef], axis=1)
        feat = feat.loc[feat.index >= pd.Timestamp(cfg.EVAL_START)]
        if len(feat) < 250:
            quality.append({"티커": t, "항목": "제외", "값": f"평가창 관측 {len(feat)} < 250"})
            continue
        panel[t] = feat
        pos[t] = build_positions(feat, cfg)
        if len(fl):
            fl2 = fl.copy(); fl2.insert(0, "티커", t); fund_ledgers.append(fl2.reset_index(drop=True))
        if len(el):
            el2 = el.copy(); el2.insert(0, "티커", t); earn_ledgers.append(el2)
        quality.append({"티커": t, "산업ETF": parent_of.get(t, ""), "이름": STOCK_NAME_KR.get(t, ""),
                        "관측일(평가창)": len(feat),
                        "시작": str(feat.index[0].date()), "종료": str(feat.index[-1].date()),
                        "펀더멘탈 분기수": (int(pd.to_numeric(feat.get("펀더멘탈_분기수"),
                                                        errors="coerce").max())
                                    if "펀더멘탈_분기수" in feat.columns
                                    and pd.to_numeric(feat["펀더멘탈_분기수"],
                                                      errors="coerce").notna().any() else 0),
                        "유효일 출처": (";".join(sorted(set(fl["유효일 출처"].astype(str)))) if len(fl) else "-"),
                        "선택 항목": (str(fl["선택 항목"].iloc[0]) if (len(fl) and "선택 항목" in fl.columns)
                                  else "-"),
                        "어닝 이벤트수": (int(len(el)) if len(el) else 0),
                        "결측 비율(ext200)": round(float(feat["ext200"].isna().mean()), 4)})
    if not panel:
        return {"aborted": True, "note": "평가창 관측이 충분한 티커가 없다", "cfg": cfg}
    log("FEAT", kv(event="features_built", tickers=len(panel),
                    with_fundamentals=int(sum(1 for t in panel
                                              if pd.to_numeric(panel[t].get("매출_YoY"),
                                                               errors="coerce").notna().any())),
                    with_earnings=int(sum(1 for t in panel
                                          if pd.to_numeric(panel[t].get("어닝_서프라이즈%"),
                                                           errors="coerce").notna().any()))))
    acc = build_rule_accuracy(panel, cfg)
    audit = build_lookahead_audit(prices, fund, cfg)

    # ---- 성과 요약(라이브 규칙 기준 · 단일 종목 롱/플랫) ----
    perf: List[dict] = []
    for t, p in pos.items():
        r = pd.to_numeric(p["전략일간수익"], errors="coerce").fillna(0.0)
        b = pd.to_numeric(panel[t]["일간수익"], errors="coerce").fillna(0.0)
        for lbl, s in (("전략(ext200)", r), ("매수보유", b)):
            cur = (1.0 + s).cumprod()
            yrs = max(len(s) / 252.0, 1e-9)
            cagr = float(cur.iloc[-1]) ** (1.0 / yrs) - 1.0
            mdd = float((cur / cur.cummax() - 1.0).min())
            perf.append({"티커": t, "이름": STOCK_NAME_KR.get(t, ""),
                         "산업ETF": parent_of.get(t, ""), "전략": lbl,
                         "총수익배수": round(float(cur.iloc[-1]), 4), "CAGR": round(cagr, 4),
                         "최대낙폭(MDD)": round(mdd, 4),
                         "칼마(CAGR/MDD)": (round(cagr / abs(mdd), 3) if mdd < -1e-9 else None),
                         "연변동성": round(float(s.std() * math.sqrt(252.0)), 4),
                         "샤프": (round(float(s.mean() / s.std() * math.sqrt(252.0)), 3)
                                if float(s.std()) > 0 else None),
                         "일간승률": round(float((s > 0).mean()), 4),
                         "평균노출": round(float(pd.to_numeric(p["집행비중"],
                                                           errors="coerce").mean()), 4)})
    log("PERF", kv(event="performance", rows=len(perf),
                    strat_median_calmar=round(float(pd.DataFrame(perf).query("전략.str.startswith('전략')",
                                                                            engine='python')
                                                    ["칼마(CAGR/MDD)"].median(skipna=True)), 3)
                    if perf else None))
    return {"cfg": cfg, "panel": panel, "pos": pos, "prices": prices, "fund": fund,
            "parent_of": parent_of, "accuracy": acc, "audit": audit,
            "quality": pd.DataFrame(quality), "perf": pd.DataFrame(perf),
            "fund_ledger": (pd.concat(fund_ledgers, ignore_index=True) if fund_ledgers
                            else pd.DataFrame()),
            "earn_ledger": (pd.concat(earn_ledgers, ignore_index=True) if earn_ledgers
                            else pd.DataFrame())}


def build_report(res: Dict[str, Any], path: Optional[str] = None, I=None) -> str:
    """[리포트] M·S·I와 같은 시트 어휘를 쓴다. I 모듈이 주어지면 19_상승하락구간을 그 함수로 만든다."""
    cfg: StockConfig = res.get("cfg", CFG)
    path = path or cfg.OUT_XLSX
    sheets: Dict[str, pd.DataFrame] = {}
    if res.get("aborted"):
        sheets["00_실행요약"] = pd.DataFrame([{"항목": "판정", "값": res.get("note", "실행 실패")}])
        _write(path, sheets)
        return path
    panel: Dict[str, pd.DataFrame] = res["panel"]
    pos: Dict[str, pd.DataFrame] = res["pos"]
    parent_of = res["parent_of"]

    # ---- 01Z 일별 예측 매트릭스(전 종목 한 시트) ----
    idx = sorted(set().union(*[set(v.index) for v in panel.values()]))
    idx = pd.DatetimeIndex(idx)
    Z = pd.DataFrame(index=idx)
    for t in sorted(panel):
        Z[f"{t} 확정국면"] = pos[t]["확정국면"].reindex(idx)
        Z[f"{t} 목표비중"] = pos[t]["목표비중"].reindex(idx)
    Z.insert(0, "상승 종목수", sum((pos[t]["목표비중"].reindex(idx).fillna(0) > 0).astype(int)
                                for t in sorted(panel)))
    sheets["01Z_주식일별예측"] = Z.reset_index().rename(columns={"index": "날짜"})

    # ---- 00 실행요약 ----
    meta: List[Tuple[str, Any]] = [
        ("제목", "미국 개별 주식 국면 예측 — M(시장)→S(섹터)→I(산업)→**K(개별 주식)** 4계층의 마지막 층"),
        ("버전", f"stock_regime.py {VERSION} ({VERSION_DATE})"),
        ("⚠ 실매매 적용 여부",
         "아니오 — 진단·연구용이다. 실매매 주문 근거는 market_regime_report.xlsx의 ★ SPY 국면전략이며, "
         "이 주식 계층은 v0.1.0(첫 버전)으로 **예측 채점**이 본체다."),
        ("표본", f"산업 ETF {len(cfg.UNIVERSE)}개마다 대표 티커 1개 = {len(panel)}종목 "
                f"(사용자 지시 '일단 샘플로 각 산업별 대표 티커 하나씩')"),
        ("평가창", f"{cfg.EVAL_START} ~ {str(idx[-1].date()) if len(idx) else '-'} "
                 f"({len(idx)}거래일) · 특성 워밍업 {cfg.START}부터"),
        ("★ 라이브 규칙(v0.1.0)",
         f"{cfg.LIVE_RULE} — 200일선 위면 비중 1, 아래면 0. **의도적으로 가장 단순한 기준선**이며 "
         "펀더멘탈·어닝 후보는 03_예측규칙정확도에서 먼저 채점한다(I 계층에서 '그럴듯한 근거로 라이브를 "
         "바꿨다가 세 라운드 연속 실패'한 이력 때문). 승격 조건은 03 시트 마지막 행에 사전등록해 두었다."),
        ("★ 펀더멘탈 인과 처리",
         f"분기 재무는 **기간말이 아니라 유효일** 기준 as-of로만 쓴다 — 실제 발표일 우선, 없으면 "
         f"기간말 + {cfg.FUND_PUBLISH_LAG_DAYS}일(10-K 기한 60일 기준의 보수 상한). "
         "'다음 어닝까지 며칠' 류는 과거 시점에 알 수 없으므로 **특성에 넣지 않았다**. "
         "04_펀더멘탈원장의 '유효일 출처' 열에서 티커별로 확인할 수 있고, 11_룩어헤드감사가 "
         "절단재계산으로 매 실행 검증한다."),
        ("판정 잣대", "소수 클래스 기준(정밀도는 항상 기저율과 함께) · 연도 k/n이 1급 · 자기이력 백분위로만 단면 비교"),
        ("면책", "본 산출물은 연구·교육 목적의 백테스트이며 투자 자문이 아니다."),
    ]
    aud = res.get("audit", pd.DataFrame())
    if len(aud):
        bad = int((aud["일치"].astype(str).str.startswith("⚠")).sum())
        meta.append(("룩어헤드 감사", f"검사 {len(aud)}건 · 불일치 **{bad}건**"
                                 + (" — 통과" if bad == 0 else " ⚠ 조사 필요")))
    pf = res.get("perf", pd.DataFrame())
    if len(pf):
        st = pf[pf["전략"].astype(str).str.startswith("전략")]
        bh = pf[pf["전략"].astype(str).eq("매수보유")]
        meta.append(("라이브 규칙 성과(중위)",
                     f"전략 CAGR {st['CAGR'].median():.4f} · MDD {st['최대낙폭(MDD)'].median():.4f} · "
                     f"칼마 {st['칼마(CAGR/MDD)'].median(skipna=True):.3f} vs "
                     f"매수보유 CAGR {bh['CAGR'].median():.4f} · MDD {bh['최대낙폭(MDD)'].median():.4f} · "
                     f"칼마 {bh['칼마(CAGR/MDD)'].median(skipna=True):.3f} "
                     f"(29종목 중위값 — 개별 종목 표는 06 시트)"))
    # 다음 거래일 예측
    if len(idx):
        last = idx[-1]
        nd = []
        for t in sorted(panel):
            p = pos[t]
            if last in p.index:
                nd.append(f"{t}({STOCK_NAME_KR.get(t, '')}) {p['확정국면'].loc[last]}/"
                          f"{float(p['목표비중'].loc[last]):.0f}")
        meta.append((f"다음 거래일 예측({last.date()} 확정 → 익일 집행)", " · ".join(nd)))
    sheets["00_실행요약"] = pd.DataFrame(meta, columns=["항목", "값"])
    sheets["02_티커요약"] = res["quality"]
    sheets["03_예측규칙정확도"] = res["accuracy"]
    if len(res.get("fund_ledger", pd.DataFrame())):
        sheets["04_펀더멘탈원장"] = res["fund_ledger"]
    if len(res.get("earn_ledger", pd.DataFrame())):
        sheets["05_어닝이벤트"] = res["earn_ledger"]
    sheets["06_성과요약"] = pf
    if len(aud):
        sheets["11_룩어헤드감사"] = aud

    # ---- 19_상승하락구간 (I 모듈 함수 재사용) ----
    if bool(getattr(cfg, "SHOW_UPDOWN_SEGMENTS", True)) and I is not None:
        try:
            ret = pd.DataFrame({t: pd.to_numeric(panel[t]["일간수익"], errors="coerce")
                                for t in sorted(panel)}).reindex(idx)
            cur = (1.0 + ret.fillna(0.0)).cumprod()
            ex = pd.DataFrame({t: pd.to_numeric(pos[t]["집행비중"], errors="coerce")
                               for t in sorted(panel)}).reindex(idx).fillna(0.0)
            seg = I.build_up_down_segments(
                cur, ex, ret, cfg,
                name_map={t: STOCK_NAME_KR.get(t, "") for t in sorted(panel)},
                parent_map={t: parent_of.get(t, "") for t in sorted(panel)},
                bench_w=None,          # 주식 계층의 비중 예산은 1.0이므로 완전 참여가 도달 가능한 벤치다
                layer="개별주식", M=None)
            if isinstance(seg, pd.DataFrame) and len(seg):
                sheets["19_상승하락구간"] = seg
        except Exception as e:
            log("REPORT", kv(event="updown_segments_failed", err=str(e)[:200]), level="warning")
    for t in sorted(panel):
        d = pd.concat([panel[t], pos[t]], axis=1)
        sheets[f"01_일별_{t}"] = d.reset_index().rename(columns={"index": "날짜"})
    _write(path, sheets)
    _mb = (os.path.getsize(path) / 1e6) if os.path.exists(path) else float("nan")
    log("REPORT", kv(event="written", path=path, sheets=len(sheets), size_mb=round(_mb, 2),
                     note="사용자 제약: 리포트 파일 전체 30MB 이하"))
    if _mb > 30.0:
        log("REPORT", kv(event="size_over_limit", size_mb=round(_mb, 2),
                         suggest="SEG_DETAIL_MAX를 줄이거나 01_일별_* 시트를 끄는 것을 검토"),
            level="warning")
    return path


def _write(path: str, sheets: Dict[str, pd.DataFrame]) -> None:
    with pd.ExcelWriter(path, engine="xlsxwriter") as xw:
        for name, df in sheets.items():
            d = df if isinstance(df, pd.DataFrame) else pd.DataFrame(df)
            d.to_excel(xw, sheet_name=str(name)[:31], index=False)


def main(s_overrides: Optional[Dict[str, Any]] = None, I=None) -> str:
    """단독 실행 진입점. run_pipeline이 I 모듈을 넘겨 주면 19 시트가 함께 나온다."""
    res = run(CFG, s_overrides)
    return build_report(res, I=I)


if __name__ == "__main__":
    print(f"stock_regime.py {VERSION} ({VERSION_DATE}) — 개별 주식 계층(K). "
          f"산업별 대표 티커 {len(STOCK_UNIVERSE)}종 · 펀더멘탈·어닝 포함. "
          f"실행: main() 또는 run_pipeline.py")
