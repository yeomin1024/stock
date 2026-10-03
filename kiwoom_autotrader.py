# %% [markdown]
# ## 1. 공통 유틸
# 미국 장 시간은 **뉴욕 시간(ET)** 으로 계산합니다(서머타임 자동 반영). 로그에는 한국 시간과 뉴욕 시간을 함께 찍습니다.

# %%
"""
미국주식 저점매수·고점매도 — 가상계좌(페이퍼) 자동매매 + 과거 실시간 시뮬레이션
- 실시간: 키움 웹소켓(FE) 또는 야후 1분봉 폴링 → 같은 엔진 → 가상계좌 체결 (실제 주문 API는 잠겨 있음)
- 과거: 분봉/시간봉을 시간 순서대로 하나씩 흘려 같은 엔진으로 가상 거래 → 결과 파일
- 일일 필터: 시장 국면(M) · 섹터(S) · 산업(I) 예측 결과를 '전날까지 확정된 값'으로만 사용(미래 정보 차단)
"""
import asyncio
import csv
import datetime as dt
import json
import math
import os
import threading
import time
from collections import deque
from dataclasses import dataclass, field, asdict
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import requests

KST = ZoneInfo("Asia/Seoul")
ET = ZoneInfo("America/New_York")
ALLOW_ORDERS = False      # ⚠ 실제/모의 주문 잠금. 가상계좌 검증이 끝나기 전에는 True로 바꾸지 마세요.


def now_et() -> dt.datetime:
    return dt.datetime.now(ET)


def hhmm(t: dt.datetime) -> str:
    return t.strftime("%H:%M")


def to_num(x, default=0.0) -> float:
    if x is None:
        return default
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip().replace(",", "")
    if s in ("", "+", "-"):
        return default
    try:
        return float(s)
    except ValueError:
        return default


def to_price(x) -> float:
    """현재가의 +/- 는 전일대비 표시라서 가격은 절댓값."""
    return abs(to_num(x))


def usd(x, sign=False) -> str:
    return f"{'+' if sign and x >= 0 else '-' if x < 0 else ''}${abs(x):,.2f}"


def log(*args):
    n = dt.datetime.now(KST)
    print(f"[KST {n:%m-%d %H:%M:%S} | ET {n.astimezone(ET):%H:%M:%S}]", *args, flush=True)


def et_session_in_kst(day: dt.date = None) -> str:
    day = day or now_et().date()
    o = dt.datetime.combine(day, dt.time(9, 30), ET).astimezone(KST)
    c = dt.datetime.combine(day, dt.time(16, 0), ET).astimezone(KST)
    return f"{o:%H:%M}~{c:%H:%M} KST"


def bar_bucket(ts: dt.datetime, n: int) -> dt.datetime:
    """09:30 기준 n분 봉의 시작 시각(야후 시간봉 09:30·10:30… 과 맞춤)."""
    o = ts.replace(hour=9, minute=30, second=0, microsecond=0)
    return o + dt.timedelta(minutes=(int((ts - o).total_seconds() // 60) // n) * n)


# %% [markdown]
# ## 2. 설정 (Config)
# 시간은 **뉴욕 시간(ET)**, 금액은 **달러**. 종목은 `{"티커": "거래소"}` — `ND` 나스닥, `NY` 뉴욕, `NA` 아멕스
# (거래소 코드는 키움 실시간 시세를 받을 때만 씁니다).

# %%
# 이전 종목 예측 코드(stock_regime.py)의 유니버스: 종목 → 산업 ETF, 산업 예측 코드(industry_rotation.py): 산업 → 섹터
PRIOR_STOCK_TO_INDUSTRY = {
    "NVDA": "SOXX", "MSFT": "IGV", "ORCL": "SKYY", "CRWD": "HACK", "AMGN": "IBB", "INCY": "XBI", "LLY": "IHE",
    "ISRG": "IHI", "UNH": "IHF", "TGT": "XRT", "DHI": "XHB", "BKNG": "PEJ", "TSLA": "CARZ", "PEP": "PBJ",
    "JPM": "KBE", "USB": "KRE", "PGR": "KIE", "GS": "KCE", "RTX": "ITA", "UNP": "IYT", "DAL": "JETS",
    "VZ": "IYZ", "GOOGL": "FDN", "META": "SOCL", "AVB": "REZ", "EOG": "XOP", "SLB": "XES", "FCX": "XME",
    "NEM": "GDX",
    "CRDO": "SOXX", "AVGO": "SOXX", "MU": "SOXX", "LRCX": "SOXX", "MRVL": "SOXX", "AMD": "SOXX", "INTC": "SOXX",
    "SNDK": "SOXX", "CRM": "IGV", "NOW": "IGV", "CDNS": "IGV", "SHOP": "IGV", "DDOG": "SKYY", "MDB": "SKYY",
    "NET": "SKYY", "AAPL": "SKYY", "WDC": "SKYY", "STX": "SKYY", "PANW": "HACK", "FTNT": "HACK", "BTSG": "IHF",
    "ROST": "XRT", "CROX": "XRT", "RL": "XRT", "AMZN": "XRT", "DASH": "PEJ", "KO": "PBJ", "MNST": "PBJ",
    "DAVE": "KBE", "SHIP": "IYT",
}
PRIOR_INDUSTRY_TO_SECTOR = {
    "SOXX": "XLK", "IGV": "XLK", "SKYY": "XLK", "HACK": "XLK", "IBB": "XLV", "XBI": "XLV", "IHE": "XLV",
    "IHI": "XLV", "IHF": "XLV", "XRT": "XLY", "XHB": "XLY", "PEJ": "XLY", "CARZ": "XLY", "PBJ": "XLP",
    "KBE": "XLF", "KRE": "XLF", "KIE": "XLF", "KCE": "XLF", "ITA": "XLI", "IYT": "XLI", "JETS": "XLI",
    "IYZ": "XLC", "FDN": "XLC", "SOCL": "XLC", "REZ": "XLRE", "XOP": "XLE", "XES": "XLE", "XME": "XLB",
    "GDX": "XLB",
}
# 위 유니버스의 상장 거래소(키움 stex_tp). 2026-10 야후 거래소 정보(NMS·NGM·NCM→ND, NYQ→NY)로 만든 표.
PRIOR_STOCK_EXCHANGE = {
    "AAPL": "ND", "AMD": "ND", "AMGN": "ND", "AMZN": "ND", "AVB": "NY", "AVGO": "ND", "BKNG": "ND", "BTSG": "ND",
    "CDNS": "ND", "CRDO": "ND", "CRM": "NY", "CROX": "ND", "CRWD": "ND", "DAL": "NY", "DASH": "ND", "DAVE": "ND",
    "DDOG": "ND", "DHI": "NY", "EOG": "NY", "FCX": "NY", "FTNT": "ND", "GOOGL": "ND", "GS": "NY", "INCY": "ND",
    "INTC": "ND", "ISRG": "ND", "JPM": "NY", "KO": "NY", "LLY": "NY", "LRCX": "ND", "MDB": "ND", "META": "ND",
    "MNST": "ND", "MRVL": "ND", "MSFT": "ND", "MU": "ND", "NEM": "NY", "NET": "NY", "NOW": "NY", "NVDA": "ND",
    "ORCL": "NY", "PANW": "ND", "PEP": "ND", "PGR": "NY", "RL": "NY", "ROST": "ND", "RTX": "NY", "SHIP": "ND",
    "SHOP": "ND", "SLB": "NY", "SNDK": "ND", "STX": "ND", "TGT": "NY", "TSLA": "ND", "UNH": "NY", "UNP": "NY",
    "USB": "NY", "VZ": "NY", "WDC": "ND",
}


DEFAULT_UNIVERSE = {k: v for k, v in PRIOR_STOCK_EXCHANGE.items() if k != "AVB"}   # AVB: 야후 데이터 없음


@dataclass
class Config:
    """기본값 = 과거 실시간 시뮬레이션에서 채택한 C1(시간봉 · 일봉 50일 추세 · 5종목×20% · M 국면 목표비중 ·
    S 섹터 배분 필터) + 실제 수수료 0.07%로 다시 고른 최대 보유 5거래일(R10)."""
    # --- 계정 / 시세 ---
    appkey: str = ""
    secretkey: str = ""
    mock: bool = True                 # 키움 시세를 받을 서버(모의/실전). 주문과는 무관
    quote_source: str = "yfinance"    # 'kiwoom' = 웹소켓 실시간(앱키·IP 등록 필요) / 'yfinance' = 1분봉 폴링(앱키 불필요)
    # --- 가상계좌 ---
    paper_cash: float = 10000.0       # 시작 가상자산($)
    paper_state: str = "paper_account.json"   # 다음 실행에 이어 쓰는 가상계좌 상태 파일
    # --- 대상 종목 {티커: 거래소} — 기본: 이전 종목 예측 코드(K)의 유니버스 57종 ---
    symbols: dict = field(default_factory=lambda: dict(DEFAULT_UNIVERSE))
    # --- 자금 / 리스크 ---
    position_pct: float = 20.0        # 1회 매수 = 가상자산 × 20% × 국면 노출도
    max_positions: int = 5
    max_trades_per_symbol: int = 1
    daily_loss_pct: float = 2.0       # 당일 실현손실이 시작자산의 2%면 신규 매수 중단
    stop_loss_pct: float = 4.0
    # --- 저점 매수 ---
    bar_minutes: int = 60             # 지표를 계산하는 봉(분). 시세는 1분·틱으로 받아도 이 길이로 묶음
    rsi_period: int = 14
    rsi_buy: float = 30.0
    bb_period: int = 20
    bb_k: float = 2.0
    rebound_pct: float = 0.5
    max_chase_pct: float = 1.0
    arm_timeout_min: int = 60
    no_reentry_after_stop: bool = True
    # --- 고점 매도 ---
    min_profit_pct: float = 3.0
    trail_pct: float = 1.0
    rsi_sell: float = 70.0
    # --- 추세·보유기간 ---
    trend_ma_days: int = 50           # >0이면 전일 종가가 일봉 N일 평균 위인 종목만 매수(상승 추세의 눌림목만)
    hold_overnight: bool = True       # True면 장마감에 팔지 않고 최대 max_hold_days 거래일 보유(스윙)
    max_hold_days: int = 5            # 수수료 0.07% 재검증(R10)에서 IS 샤프 최고
    # --- 시간 (ET) ---
    entry_start: str = "09:35"
    entry_end: str = "15:30"
    force_exit: str = "15:50"
    session_end: str = "16:01"
    cooldown_sec: int = 300
    market_hours_check: bool = True
    # --- 비용 (%) ---
    fee_pct: float = 0.07             # 사용자 계좌 미국주식 수수료(매수·매도 각각)
    sell_fee_pct: float = 0.003
    slippage_pct: float = 0.05
    # --- 일일 필터 (이전 예측 코드) ---
    signals_dir: str = "signals"      # 이전 예측 결과 폴더(market_regime_daily.csv · sector_allocation_daily.csv …)
    regime: str = "M"                 # 'M' = 시장 국면 모델 목표비중(없으면 SPY로 대체) / 'spy' = SPY 200일선·60일 모멘텀 / 'off'
    sector_filter: bool = True        # S: 섹터 배분 0인 섹터의 종목은 신규 매수 금지
    industry_filter: bool = False     # I: 산업 배분 0인 산업의 종목은 신규 매수 금지(그날 산업 배분이 전부 0이면 적용 안 함)
    rank_col: str = ""                # I 점수로 종목 순위: 'P(상승,ML)' / 'P(부모초과)' / '목표비중' / '' = 안 씀
    watchlist_size: int = 0           # 순위 상위 N종만 그날 매수 대상(0 = 전부)
    alloc_map: dict = field(default_factory=lambda: dict(PRIOR_STOCK_TO_INDUSTRY))
    # --- 기타 ---
    status_every_min: int = 10
    trade_log: str = "paper_trades.csv"
    equity_log: str = "paper_equity.csv"


# %% [markdown]
# ## 3. 키움 REST API 클라이언트 (미국주식 시세)
# 토큰 `POST /oauth2/token`, 분봉 `usa06011`, 실시간 체결 `FE`(`wss://…:10000/api/us/websocket`).
# 주문 함수(`ust20000`/`ust20001`)는 남겨 두었지만 **`ALLOW_ORDERS=False`인 동안 호출하면 바로 오류**가 납니다.

# %%
class KiwoomError(RuntimeError):
    pass


def _parse_expires(s) -> dt.datetime:
    s = str(s or "").strip()
    for fmt in ("%Y%m%d%H%M%S", "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S"):
        try:
            return dt.datetime.strptime(s[:19], fmt).replace(tzinfo=KST)
        except ValueError:
            pass
    return dt.datetime.now(KST) + dt.timedelta(hours=12)


class KiwoomUSAPI:
    def __init__(self, appkey, secretkey, mock=True, max_req_per_sec=4.0):
        host = "mockapi.kiwoom.com" if mock else "api.kiwoom.com"
        self.base = f"https://{host}"
        self.ws_url = f"wss://{host}:10000/api/us/websocket"
        self.appkey, self.secretkey = appkey, secretkey
        self.token, self.token_exp = None, None
        self._min_interval = 1.0 / max_req_per_sec
        self._per_tr = 1.05 if mock else 0.0     # 모의투자: API 하나당 1초 1회
        self._last_req, self._last_tr = 0.0, {}
        self._lock = threading.Lock()
        self.session = requests.Session()

    def _throttle(self, api_id=""):
        with self._lock:
            now = time.monotonic()
            wait = max(self._last_req + self._min_interval, self._last_tr.get(api_id, 0.0) + self._per_tr) - now
            if wait > 0:
                time.sleep(wait)
            self._last_req = self._last_tr[api_id] = time.monotonic()

    def get_token(self, force=False) -> str:
        if not force and self.token and dt.datetime.now(KST) < self.token_exp - dt.timedelta(minutes=10):
            return self.token
        if not self.appkey or not self.secretkey:
            raise KiwoomError("앱키/시크릿키가 비어 있습니다.")
        self._throttle("token")
        r = self.session.post(self.base + "/oauth2/token",
                              json={"grant_type": "client_credentials",
                                    "appkey": self.appkey, "secretkey": self.secretkey},
                              headers={"Content-Type": "application/json;charset=UTF-8"}, timeout=10)
        try:
            j = r.json()
        except ValueError:
            raise KiwoomError(f"토큰 응답 파싱 실패 HTTP {r.status_code}: {r.text[:200]}")
        if not j.get("token"):
            raise KiwoomError(f"토큰 발급 실패: {j.get('return_code')} {j.get('return_msg')}\n"
                              "→ openapi.kiwoom.com 에서 지금 이 런타임의 공인 IP가 등록됐는지, "
                              "모의/실전 앱키를 맞게 넣었는지 확인하세요.")
        self.token, self.token_exp = j["token"], _parse_expires(j.get("expires_dt"))
        return self.token

    def request(self, path, api_id, body, cont_yn="N", next_key=""):
        for attempt in range(3):
            self._throttle(api_id)
            headers = {"Content-Type": "application/json;charset=UTF-8",
                       "authorization": f"Bearer {self.get_token()}",
                       "api-id": api_id, "cont-yn": cont_yn, "next-key": next_key}
            r = self.session.post(self.base + path, json=body, headers=headers, timeout=10)
            if r.status_code == 429:
                time.sleep(1.0 + attempt)
                continue
            try:
                j = r.json()
            except ValueError:
                raise KiwoomError(f"[{api_id}] HTTP {r.status_code}: {r.text[:200]}")
            rc, msg = str(j.get("return_code", "0")), str(j.get("return_msg", ""))
            if rc == "0" and r.status_code < 400:
                return j, r.headers
            if rc == "5" or "초과" in msg:
                time.sleep(1.0 + attempt)
                continue
            if "8005" in msg or "토큰" in msg or "Token" in msg:
                self.get_token(force=True)
                continue
            raise KiwoomError(f"[{api_id}] return_code={rc} {msg}")
        raise KiwoomError(f"[{api_id}] 재시도 한도 초과")

    def minute_closes(self, code, exch, minutes=1, pages=1) -> list:
        """분봉 종가(과거→현재). 키움 차트는 최신 봉이 먼저 온다(영업일자로 확인)."""
        rows, cont, key = [], "N", ""
        body = {"stex_tp": exch, "stk_cd": code, "strt_dt": now_et().strftime("%Y%m%d"),
                "tic_scope": str(minutes), "upd_stkpc_tp": "1", "exrt_appl_tp": "0"}
        for _ in range(pages):
            j, h = self.request("/api/us/chart", "usa06011", body, cont, key)
            rows += j.get("result_list") or []
            if h.get("cont-yn") != "Y":
                break
            cont, key = "Y", h.get("next-key", "")
        if not rows:
            return []
        days = [str(r.get("bus_dt", "")) for r in rows]
        newest_first = days[0] > days[-1] if days[0] != days[-1] else True
        closes = [to_price(r.get("cur_prc")) for r in rows]
        closes = closes[::-1] if newest_first else closes
        return [c for c in closes if c > 0]

    def minute_history(self, code, exch, minutes=5, strt_dt=None, max_pages=300, log_first=False):
        """분봉 원자료를 쪽(page)마다 이어 받아 그대로 돌려준다. (rows, 쪽수)"""
        body = {"stex_tp": exch, "stk_cd": code, "strt_dt": strt_dt or now_et().strftime("%Y%m%d"),
                "tic_scope": str(minutes), "upd_stkpc_tp": "1", "exrt_appl_tp": "0"}
        rows, pages, cont, key = [], 0, "N", ""
        while pages < max_pages:
            j, h = self.request("/api/us/chart", "usa06011", body, cont, key)
            page = j.get("result_list") or []
            if log_first and pages == 0 and page:
                log(f"  {code} 첫 쪽 {len(page)}봉 | 첫 행 {page[0]} | 마지막 행 {page[-1]}")
            rows += page
            pages += 1
            if h.get("cont-yn") != "Y" or not page:
                break
            cont, key = "Y", h.get("next-key", "")
        return rows, pages

    def order(self, side, code, exch, qty) -> dict:
        if not ALLOW_ORDERS:
            raise KiwoomError("주문이 잠겨 있습니다(ALLOW_ORDERS=False). 지금은 가상계좌로만 거래합니다.")
        api_id = "ust20000" if side == "BUY" else "ust20001"
        body = {"stex_tp": exch, "stk_cd": code, "ord_qty": str(int(qty)), "trde_tp": "03", "ord_uv": ""}
        j, _ = self.request("/api/us/ordr", api_id, body)
        return j


# %% [markdown]
# ## 4. 저점매수·고점매도 전략
# | 단계 | 조건 |
# |---|---|
# | 저점 구간 진입 | 봉 RSI ≤ `rsi_buy` **또는** 현재가 ≤ 볼린저 하단 |
# | 매수 | 구간 최저가 대비 `rebound_pct` 반등 (떨어지는 칼날 회피) |
# | 감시 해제 | 볼린저 중심선 회복, 저점 대비 `max_chase_pct` 이상 상승(추격 금지), `arm_timeout_min` 경과 |
# | 익절 | 고점 수익률이 `min_profit_pct`를 찍은 뒤 고점 대비 `trail_pct` 하락(과열이면 폭 절반) |
# | 손절 / 장마감 | 매수가 대비 `stop_loss_pct` 하락 / `force_exit` 이후 전량 |

# %%
def rsi_wilder(closes, period) -> float:
    if len(closes) < period + 1:
        return math.nan
    d = np.diff(np.asarray(closes, float))
    up, dn = np.clip(d, 0, None), np.clip(-d, 0, None)
    ag, al = up[:period].mean(), dn[:period].mean()
    for u, w in zip(up[period:], dn[period:]):
        ag = (ag * (period - 1) + u) / period
        al = (al * (period - 1) + w) / period
    if al == 0:
        return 100.0 if ag > 0 else 50.0
    return 100.0 - 100.0 / (1.0 + ag / al)


@dataclass
class SymState:
    code: str
    closes: deque = field(default_factory=lambda: deque(maxlen=400))
    rsi: float = math.nan
    mid: float = math.nan
    upper: float = math.nan
    lower: float = math.nan
    bar_t: dt.datetime = None
    bar_end: dt.datetime = None
    bar_c: float = 0.0
    daily: deque = field(default_factory=lambda: deque(maxlen=300))   # 일봉 종가(추세 필터)
    trend_ok: bool = True
    ag: float = math.nan         # RSI(와일더) 평균 상승폭 — 봉마다 갱신
    al: float = math.nan
    armed: bool = False
    low: float = math.inf
    armed_at: dt.datetime = None
    qty: int = 0
    entry: float = 0.0
    peak: float = 0.0
    entry_time: dt.datetime = None
    buys_today: int = 0
    stopped_today: bool = False
    cooldown_until: dt.datetime = None
    last_price: float = 0.0
    last_ts: dt.datetime = None
    bad_ticks: int = 0


class LowHighStrategy:
    def __init__(self, cfg: Config):
        self.cfg = cfg

    def recompute(self, st: SymState):
        """봉이 하나 닫힐 때(또는 워밍업으로 종가를 한꺼번에 넣은 직후) 호출 — RSI는 와일더 방식으로 이어서 갱신."""
        c, p, n = self.cfg, self.cfg.rsi_period, len(st.closes)
        if math.isnan(st.ag):
            if n >= p + 1:
                d = np.diff(np.asarray(st.closes, float))
                ag, al = np.clip(d[:p], 0, None).mean(), np.clip(-d[:p], 0, None).mean()
                for x in d[p:]:
                    ag, al = (ag * (p - 1) + max(x, 0.0)) / p, (al * (p - 1) + max(-x, 0.0)) / p
                st.ag, st.al = ag, al
        else:
            x = st.closes[-1] - st.closes[-2]
            st.ag, st.al = (st.ag * (p - 1) + max(x, 0.0)) / p, (st.al * (p - 1) + max(-x, 0.0)) / p
        if not math.isnan(st.ag):
            st.rsi = (100.0 if st.ag > 0 else 50.0) if st.al == 0 else 100.0 - 100.0 / (1.0 + st.ag / st.al)
        if n >= c.bb_period:
            w = [st.closes[-k] for k in range(1, c.bb_period + 1)]
            m = sum(w) / len(w)
            s = math.sqrt(sum((v - m) ** 2 for v in w) / len(w))
            st.mid, st.upper, st.lower = m, m + c.bb_k * s, m - c.bb_k * s

    def update_bar(self, st: SymState, ts: dt.datetime, px: float) -> bool:
        if st.bar_end is not None and st.bar_t <= ts < st.bar_end:      # 같은 봉(가장 흔한 경우)
            st.bar_c = px
            return False
        b = bar_bucket(ts, self.cfg.bar_minutes)
        if st.bar_t is None:
            st.bar_t, st.bar_end, st.bar_c = b, b + dt.timedelta(minutes=self.cfg.bar_minutes), px
            return False
        if b > st.bar_t:
            st.closes.append(st.bar_c)
            self.recompute(st)
            st.bar_t, st.bar_end, st.bar_c = b, b + dt.timedelta(minutes=self.cfg.bar_minutes), px
            return True
        st.bar_c = px
        return False

    def decide(self, st: SymState, px: float, now: dt.datetime):
        c = self.cfg
        if st.qty > 0:
            st.peak = max(st.peak, px)
            gain = px / st.entry - 1
            if px <= st.entry * (1 - c.stop_loss_pct / 100):
                return "SELL", f"손절 {gain * 100:+.2f}%"
            if hhmm(now) >= c.force_exit:
                if not c.hold_overnight:
                    return "SELL", f"장마감 청산 {gain * 100:+.2f}%"
                if np.busday_count(st.entry_time.date(), now.date()) >= c.max_hold_days - 1:
                    return "SELL", f"보유기한 청산 {gain * 100:+.2f}%"
            if (st.peak / st.entry - 1) * 100 >= c.min_profit_pct:
                hot = st.rsi >= c.rsi_sell or (not math.isnan(st.upper) and st.peak >= st.upper)
                trail = c.trail_pct * (0.5 if hot else 1.0)
                if px <= st.peak * (1 - trail / 100):
                    return "SELL", f"고점 {usd(st.peak)} 대비 -{trail:.2f}% → 익절 {gain * 100:+.2f}%"
            return None, ""
        if math.isnan(st.lower) or math.isnan(st.rsi):
            return None, "워밍업 중"
        if not st.armed:
            if st.rsi <= c.rsi_buy or px <= st.lower:
                st.armed, st.low, st.armed_at = True, px, now
                return None, f"저점 구간 진입 (RSI {st.rsi:.0f}, 하단 {usd(st.lower)})"
            return None, ""
        st.low = min(st.low, px)
        if (px >= st.mid or px >= st.low * (1 + c.max_chase_pct / 100)
                or (now - st.armed_at).total_seconds() > c.arm_timeout_min * 60):
            st.armed, st.low = False, math.inf
            return None, "저점 감시 해제"
        if px >= st.low * (1 + c.rebound_pct / 100):
            return "BUY", f"저점 {usd(st.low)} 대비 +{(px / st.low - 1) * 100:.2f}% 반등 (RSI {st.rsi:.0f})"
        return None, ""


# %% [markdown]
# ## 5. 가상계좌 (Book)
# 가상 현금·보유·실현손익·일별 자산을 관리합니다. 매수는 가상 현금 안에서만 되고, 수수료·슬리피지가 현금에서 빠집니다.
# 실시간 실행에서는 `paper_account.json`에 저장돼 다음 날 이어집니다.

# %%
class Book:
    def __init__(self, cfg: Config, trade_log: str = None, verbose=True):
        self.cfg, self.trade_log, self.verbose = cfg, trade_log, verbose
        self.st: dict[str, SymState] = {}
        self.start_cash = self.cash = float(cfg.paper_cash)
        self.realized = self.realized_today = self.fees = 0.0
        self.day_start_equity = self.cash
        self.halted = False
        self.exposure = 1.0
        self.trades, self.equity_rows = [], []
        self.day = None

    def get(self, code) -> SymState:
        if code not in self.st:
            self.st[code] = SymState(code)
        return self.st[code]

    def equity(self) -> float:
        return self.cash + sum(s.qty * s.last_price for s in self.st.values() if s.qty)

    def open_count(self) -> int:
        return sum(1 for s in self.st.values() if s.qty > 0)

    def new_day(self, day):
        self.day = day
        self.realized_today, self.halted = 0.0, False
        self.day_start_equity = self.equity()
        for s in self.st.values():
            s.buys_today, s.cooldown_until, s.armed, s.low, s.stopped_today = 0, None, False, math.inf, False

    def can_enter(self, code, now, eligible=None) -> tuple:
        c, s = self.cfg, self.get(code)
        if not (c.entry_start <= hhmm(now) < c.entry_end):
            return False, "매수 시간대 아님"
        if self.halted:
            return False, "일일 손실한도"
        if self.exposure <= 0:
            return False, "국면 노출 0%"
        if eligible is not None and code not in eligible:
            return False, "오늘 매수 대상 아님(섹터/산업/순위)"
        if s.qty > 0:
            return False, "보유 중"
        if s.buys_today >= c.max_trades_per_symbol:
            return False, "종목당 매수 횟수"
        if c.trend_ma_days and not s.trend_ok:
            return False, "일봉 추세 아님"
        if c.no_reentry_after_stop and s.stopped_today:
            return False, "오늘 손절한 종목"
        if s.cooldown_until and now < s.cooldown_until:
            return False, "재매수 대기"
        if self.open_count() >= c.max_positions:
            return False, "동시 보유 한도"
        return True, ""

    def order_qty(self, px) -> int:
        if px <= 0:
            return 0
        budget = self.equity() * self.cfg.position_pct / 100 * self.exposure
        budget = min(budget, self.cash / (1 + self.cfg.fee_pct / 100))
        return int(budget // px)

    def sim_price(self, side, px) -> float:
        s = self.cfg.slippage_pct / 100
        return px * (1 + s) if side == "BUY" else px * (1 - s)

    def apply_fill(self, code, side, qty, px, now, reason=""):
        c, s = self.cfg, self.get(code)
        fee, sfee = c.fee_pct / 100, c.sell_fee_pct / 100
        pnl = None
        if side == "BUY":
            cost = qty * px * (1 + fee)
            self.cash -= cost
            self.fees += qty * px * fee
            s.entry = (s.entry * s.qty + px * qty) / (s.qty + qty)
            s.peak = max(s.peak, px) if s.qty > 0 else px
            s.qty += qty
            s.entry_time = now
            s.last_price = s.last_price or px
        else:
            qty = min(qty, s.qty)
            if qty <= 0:
                return None
            self.cash += qty * px * (1 - fee - sfee)
            self.fees += qty * px * (fee + sfee)
            pnl = qty * (px * (1 - fee - sfee) - s.entry * (1 + fee))
            self.realized += pnl
            self.realized_today += pnl
            s.qty -= qty
            if reason.startswith("손절"):
                s.stopped_today = True
            if s.qty == 0:
                s.cooldown_until = now + dt.timedelta(seconds=c.cooldown_sec)
                s.peak = 0.0
        rec = {"시각(ET)": now.strftime("%Y-%m-%d %H:%M:%S"), "종목": code, "구분": "매수" if side == "BUY" else "매도",
               "수량": qty, "가격($)": round(px, 4),
               "손익($)": None if pnl is None else round(pnl, 2),
               "수익률(%)": None if pnl is None else round(pnl / (qty * s.entry) * 100 if s.entry else 0, 3),
               "가상현금($)": round(self.cash, 2), "사유": reason}
        if side == "SELL" and s.qty == 0:
            s.entry = 0.0
        self.trades.append(rec)
        if self.verbose:
            extra = "" if pnl is None else f" 손익 {usd(pnl, True)} (당일 {usd(self.realized_today, True)})"
            log(f"✅ 가상체결 {rec['구분']} {code} {qty}주 @ {usd(px)}{extra} | 현금 {usd(self.cash)} | {reason}")
        if self.trade_log:
            _append_csv(self.trade_log, rec)
        limit = self.day_start_equity * c.daily_loss_pct / 100
        if pnl is not None and self.realized_today <= -limit and not self.halted:
            self.halted = True
            if self.verbose:
                log(f"⛔ 당일 실현손실 {usd(self.realized_today)} (한도 {usd(limit)}) → 신규 매수 중단")
        return pnl

    def snapshot(self, day=None):
        row = {"날짜": str(day or self.day), "자산($)": round(self.equity(), 2), "현금($)": round(self.cash, 2),
               "당일실현($)": round(self.realized_today, 2), "누적실현($)": round(self.realized, 2),
               "누적수수료($)": round(self.fees, 2), "국면노출": round(self.exposure, 2),
               "보유종목": ",".join(f"{k}:{s.qty}" for k, s in self.st.items() if s.qty)}
        self.equity_rows.append(row)
        return row

    # ---- 저장/불러오기 (실시간 가상계좌) ----
    def save_state(self, path):
        state = {"cash": self.cash, "start_cash": self.start_cash, "realized": self.realized, "fees": self.fees,
                 "positions": {k: {"qty": s.qty, "entry": s.entry, "peak": s.peak, "last": s.last_price,
                                   "entry_time": s.entry_time.isoformat() if s.entry_time else None}
                               for k, s in self.st.items() if s.qty},
                 "saved_at": dt.datetime.now(KST).isoformat()}
        with open(path, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=1)

    def load_state(self, path) -> bool:
        if not path or not os.path.exists(path):
            return False
        with open(path, encoding="utf-8") as f:
            st = json.load(f)
        self.cash, self.start_cash = float(st["cash"]), float(st.get("start_cash", st["cash"]))
        self.realized, self.fees = float(st.get("realized", 0)), float(st.get("fees", 0))
        for k, p in (st.get("positions") or {}).items():
            s = self.get(k)
            s.qty, s.entry, s.peak, s.last_price = int(p["qty"]), float(p["entry"]), float(p["peak"]), float(p["last"])
            et = p.get("entry_time")
            s.entry_time = dt.datetime.fromisoformat(et) if et else now_et()
        return True


def _append_csv(path, rec):
    new = not os.path.exists(path)
    with open(path, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(rec))
        if new:
            w.writeheader()
        w.writerow(rec)


# %% [markdown]
# ## 6. 일일 필터 — 이전 국면·섹터·산업·종목 예측 결과 연결
# 날짜 D의 매매에는 **D보다 앞선 날짜의 행만** 씁니다(그날 종가로 만든 신호는 다음 거래일에 체결 — 미래 정보 차단).
# - **M 국면**: `market_regime_daily.csv`의 `목표비중`(0~1) → 1회 매수 금액에 곱함. 0이면 신규 매수 없음
# - **S 섹터**: `sector_allocation_daily.csv`의 `XLK 배분비중` 등 → 0인 섹터의 종목 제외
# - **I 산업**: `industry_allocation_daily.csv`의 산업 비중 → 0인 산업의 종목 제외(그날 산업 배분이 전부 0이면 미적용)
# - **I 점수 순위**: `industry_daily.csv`의 `SOXX P(상승,ML)` 등 → 상위 `watchlist_size`종만 매수 대상
# - **K 종목**: 유니버스와 종목→산업 매핑(`PRIOR_STOCK_TO_INDUSTRY`). 종목 확률은 선택 정보가 없어(IC≈0) 쓰지 않음
# - 대안 국면 `spy`: SPY 종가>200일선 & 60일 수익률>0 → 100%, 둘 다 반대 → 0%, 섞이면 50%

# %%
@dataclass
class DaySignal:
    exposure: float = 1.0
    eligible: set = None        # None = 전부 매수 대상
    note: str = ""


def _dated_csv(path) -> pd.DataFrame:
    df = pd.read_csv(path, encoding="utf-8-sig", low_memory=False)
    df.index = pd.to_datetime(df.iloc[:, 0], errors="coerce")
    df = df[df.index.notna()].sort_index()
    return df[~df.index.duplicated(keep="last")]


def spy_regime_series(start="2015-01-01") -> pd.Series:
    """날짜 t의 값 = t 종가까지 보고 정한 노출도(→ t 다음 거래일에 사용)."""
    import yfinance as yf
    c = yf.Ticker("SPY").history(start=start, interval="1d", auto_adjust=True)["Close"]
    c.index = c.index.tz_localize(None).normalize() if c.index.tz is not None else c.index
    ma, mom = c.rolling(200).mean(), c / c.shift(60) - 1
    e = np.where((c > ma) & (mom > 0), 1.0, np.where((c < ma) & (mom < 0), 0.0, 0.5))
    return pd.Series(e, index=c.index)[ma.notna()]


class DailySignals:
    """strict=False(실시간)면 파일이 없을 때 경고하고 대체: M 없음 → SPY 국면, S·I 없음 → 그 필터 끔."""
    def __init__(self, cfg: Config, codes=None, spy_series: pd.Series = None, strict=True):
        self.cfg = cfg
        self.codes = [c.upper() for c in (codes or cfg.symbols)]
        self.warn = []
        d = cfg.signals_dir

        def need(name):
            p = os.path.join(d, name)
            if not (d and os.path.exists(p)):
                if strict:
                    raise FileNotFoundError(f"{name} 없음 — signals_dir({d!r})에 이전 파이프라인 결과를 두세요")
                self.warn.append(f"{name} 없음")
                return None
            return _dated_csv(p)
        regime = cfg.regime
        self.m = need("market_regime_daily.csv") if regime == "M" else None
        if regime == "M" and self.m is None:
            regime = "spy"
            self.warn.append("→ 국면은 SPY 200일선 규칙으로 대체")
        self.spy = None
        if regime == "spy":
            try:
                self.spy = spy_series if spy_series is not None else spy_regime_series()
            except Exception as e:
                if strict:
                    raise
                self.warn.append(f"SPY 국면 계산 실패({e}) → 노출 100%")
        self.s = need("sector_allocation_daily.csv") if cfg.sector_filter else None
        self.i = need("industry_allocation_daily.csv") if cfg.industry_filter else None
        self.ip = need("industry_daily.csv") if cfg.rank_col else None

    def staleness(self, day) -> list:
        out = []
        for name, df in (("M", self.m), ("S", self.s), ("I", self.i)):
            if df is not None and len(df):
                age = np.busday_count(df.index[-1].date(), day)
                if age > 3:
                    out.append(f"{name} 신호 {age}거래일 지남(마지막 {df.index[-1].date()})")
        return out

    @staticmethod
    def _asof(obj, day):
        sub = obj[obj.index < pd.Timestamp(day)]
        return None if len(sub) == 0 else sub.iloc[-1]

    def for_day(self, day) -> DaySignal:
        c, notes = self.cfg, []
        exp = 1.0
        if self.m is not None:
            r = self._asof(self.m, day)
            if r is not None and not pd.isna(r.get("목표비중")):
                exp = float(np.clip(to_num(r["목표비중"]), 0, 1))
            notes.append(f"M {exp:.2f}")
        elif self.spy is not None:
            r = self._asof(self.spy, day)
            exp = 1.0 if r is None else float(r)
            notes.append(f"SPY국면 {exp:.2f}")
        elig = set(self.codes)
        ind = {k: c.alloc_map.get(k) for k in self.codes}
        if self.s is not None:
            r = self._asof(self.s, day)
            if r is not None:
                keep = {k for k in elig if ind[k] and
                        to_num(r.get(f"{PRIOR_INDUSTRY_TO_SECTOR.get(ind[k])} 배분비중"), 0) > 0}
                notes.append(f"S {len(keep)}/{len(elig)}")
                elig = keep
        if self.i is not None:
            r = self._asof(self.i, day)
            if r is not None and to_num(r.get("산업배분 합계"), 0) > 0:
                keep = {k for k in elig if ind[k] and to_num(r.get(ind[k]), 0) > 0}
                notes.append(f"I {len(keep)}/{len(elig)}")
                elig = keep
            else:
                notes.append("I 배분없음→미적용")
        if self.ip is not None and c.watchlist_size > 0:
            r = self._asof(self.ip, day)
            if r is not None:
                score = {k: to_num(r.get(f"{ind[k]} {c.rank_col}"), -1) for k in elig if ind[k]}
                top = sorted(score, key=lambda k: -score[k])[:c.watchlist_size]
                notes.append(f"순위 상위{len(top)}")
                elig = set(top)
        return DaySignal(exp, elig, " · ".join(notes))


# %% [markdown]
# ## 7. 엔진 — 실시간과 과거 시뮬레이션이 함께 쓰는 판단·체결 코드
# 가격이 들어올 때마다 `on_price` → 봉 갱신 → 전략 판단 → 가상계좌 체결. 날짜가 바뀌면 그날의 필터를 다시 계산합니다.

# %%
class Engine:
    def __init__(self, cfg: Config, book: Book = None, signals: DailySignals = None, verbose=True):
        self.cfg = cfg
        self.codes = [c.upper() for c in cfg.symbols]
        self.book = book or Book(cfg, verbose=verbose)
        self.signals = signals
        self.strat = LowHighStrategy(cfg)
        self.verbose = verbose
        self.day = None
        self.today = DaySignal()
        self.entry_filter = None      # (code, ts, SymState) -> bool : 지표 모델로 매수 신호 거르기(없으면 전부 통과)
        self.entry_signal = None      # (code, ts, SymState) -> float|None : 지표 예측 모델 점수. 있으면 규칙 대신 이걸로 매수
        self.entry_thr = 0.0

    def start_day(self, day):
        self.day = day
        self.book.new_day(day)
        self.today = self.signals.for_day(day) if self.signals else DaySignal()
        self.book.exposure = self.today.exposure
        self.refresh_trend()
        if self.verbose:
            n = len(self.codes) if self.today.eligible is None else len(self.today.eligible)
            log(f"📅 {day} | 국면 노출 {self.today.exposure:.0%} | 매수 대상 {n}/{len(self.codes)}종 | "
                f"가상자산 {usd(self.book.equity())} | {self.today.note}")

    def refresh_trend(self):
        n = self.cfg.trend_ma_days
        if n:
            for s in self.book.st.values():
                d = s.daily
                s.trend_ok = len(d) >= n and d[-1] > sum(d[-k] for k in range(1, n + 1)) / n

    def end_day(self):
        if self.day is not None:
            for s in self.book.st.values():          # 그날 마지막 가격 = 일봉 종가(추세 필터용)
                if s.last_ts is not None and s.last_ts.date() == self.day:
                    s.daily.append(s.last_price)
            row = self.book.snapshot(self.day)
            if self.verbose:
                log(f"🧾 {self.day} 마감 | 자산 {usd(row['자산($)'])} | 당일실현 {usd(row['당일실현($)'], True)}")
            return row

    def on_price(self, code, px, ts):
        if ts.date() != self.day:
            self.end_day()
            self.start_day(ts.date())
        s = self.book.get(code)
        ref = s.closes[-1] if s.closes else s.last_price
        if ref and not (ref / 1.3 < px < ref * 1.3):
            s.bad_ticks += 1
            if self.verbose and s.bad_ticks in (1, 100):
                log(f"⚠️ {code} 이상 틱 무시: {px} (기준 {ref:.2f})")
            return
        s.last_price, s.last_ts = px, ts
        closed = self.strat.update_bar(s, ts, px)
        if self.entry_signal is not None:                  # 모델 매수 모드: 봉이 닫힐 때 한 번 판단, 매도는 규칙(손절·익절·기한)
            if s.qty > 0:
                action, reason = self.strat.decide(s, px, ts)
                if action == "SELL":
                    self._fill(code, "SELL", px, ts, reason)
            elif closed and self.book.can_enter(code, ts, self.today.eligible)[0]:
                score = self.entry_signal(code, ts, s)
                if score is not None and score >= self.entry_thr:
                    self._fill(code, "BUY", px, ts, f"지표 모델 점수 {score:.4f}")
            return
        was = s.armed
        action, reason = self.strat.decide(s, px, ts)
        if self.verbose and s.armed and not was:
            log(f"👀 {code} {reason} @ {usd(px)}")
        if action == "BUY" and self.book.can_enter(code, ts, self.today.eligible)[0]:
            if self.entry_filter is not None and not self.entry_filter(code, ts, s):
                s.armed, s.low = False, math.inf          # 모델이 거른 신호 → 다음 저점 구간을 기다림
                return
            self._fill(code, "BUY", px, ts, reason)
        elif action == "SELL" and s.qty > 0:
            self._fill(code, "SELL", px, ts, reason)

    def on_clock(self, ts):
        """틱이 뜸해도 손절·장마감 청산이 되게 주기적으로 호출."""
        for code, s in self.book.st.items():
            if s.qty > 0 and s.last_price > 0:
                action, reason = self.strat.decide(s, s.last_price, ts)
                if action == "SELL":
                    self._fill(code, "SELL", s.last_price, ts, reason)

    def _fill(self, code, side, px, ts, reason):
        s, b = self.book.get(code), self.book
        if side == "BUY":
            fill = b.sim_price("BUY", px)
            qty = b.order_qty(fill)
            s.armed, s.low = False, math.inf
            if qty < 1:
                return
            s.buys_today += 1
            b.apply_fill(code, "BUY", qty, fill, ts, reason)
        else:
            b.apply_fill(code, "SELL", s.qty, b.sim_price("SELL", px), ts, reason)

    def status_df(self) -> pd.DataFrame:
        rows = []
        for code in self.codes:
            s = self.book.get(code)
            nan = math.isnan
            rows.append({"종목": code, "현재가($)": round(s.last_price, 2), "RSI": None if nan(s.rsi) else round(s.rsi, 1),
                         "볼린저하단": None if nan(s.lower) else round(s.lower, 2), "저점감시": "●" if s.armed else "",
                         "보유": s.qty, "매수가": round(s.entry, 2),
                         "평가(%)": round((s.last_price / s.entry - 1) * 100, 2) if s.qty else None,
                         "금일매수": s.buys_today,
                         "대상": "" if self.today.eligible is None or code in self.today.eligible else "제외"})
        return pd.DataFrame(rows)


# %% [markdown]
# ## 8. 실시간 가상거래 (LiveRunner)
# 시세만 실시간으로 받고, 체결은 전부 가상계좌에서 합니다(주문 API 호출 없음).
# - `quote_source="kiwoom"`: 키움 웹소켓 실시간 체결(앱키·IP 등록 필요)
# - `quote_source="yfinance"`: 야후 1분봉을 1분마다 받아 씀(앱키 불필요, 1~2분 지연)

# %%
class LiveRunner:
    def __init__(self, cfg: Config, api: KiwoomUSAPI = None, signals: DailySignals = None, clock=None):
        self.cfg = cfg
        self.api = api
        self.clock = clock or now_et
        self.exch = {str(k).upper(): str(v).upper() for k, v in cfg.symbols.items()}
        needs = cfg.regime in ("M", "spy") or cfg.sector_filter or cfg.industry_filter or cfg.rank_col
        if signals is None and needs:
            signals = DailySignals(cfg, strict=False)
            for w in signals.warn:
                log(f"⚠️ 일일 필터: {w}")
            for w in signals.staleness(self.clock().date()):
                log(f"⚠️ {w} — 이전 예측 파이프라인을 다시 돌려 signals_dir 파일을 갱신하세요")
        self.engine = Engine(cfg, Book(cfg, trade_log=cfg.trade_log), signals)
        self._stop = False
        self._ws = None
        self._tasks = []
        self._seen = {}

    @property
    def book(self):
        return self.engine.book

    async def run(self):
        c = self.cfg
        log(f"가상거래 시작 [시세: {c.quote_source}] 종목 {list(self.exch)} | 오늘 정규장 {et_session_in_kst()}")
        now = self.clock()
        if c.market_hours_check and (now.weekday() >= 5 or hhmm(now) >= c.session_end):
            log(f"지금은 미국 정규장 시간이 아닙니다. 정규장은 {et_session_in_kst()} — 그 전에는 과거 시뮬레이션을 이용하세요.")
            return
        if self.book.load_state(c.paper_state):
            log(f"가상계좌 불러옴: 현금 {usd(self.book.cash)} · 누적실현 {usd(self.book.realized, True)}")
        else:
            log(f"새 가상계좌: {usd(self.book.cash)}")
        self.engine.start_day(now.date())
        await self._warmup()
        feed = self._kiwoom_loop if c.quote_source == "kiwoom" else self._yf_loop
        self._tasks = [asyncio.create_task(feed()), asyncio.create_task(self._housekeeping())]
        try:
            await asyncio.gather(*self._tasks)
        except (asyncio.CancelledError, KeyboardInterrupt):
            log("중지 요청")
        finally:
            self._stop = True
            for t in self._tasks:
                t.cancel()
            await asyncio.gather(*self._tasks, return_exceptions=True)
            self.close_day()

    def close_day(self):
        row = self.engine.end_day()
        self.engine.day = None
        if row and self.cfg.equity_log:
            _append_csv(self.cfg.equity_log, row)
        if self.cfg.paper_state:
            self.book.save_state(self.cfg.paper_state)
        print(self.engine.status_df().to_string(index=False))
        log(f"가상자산 {usd(self.book.equity())} (시작 {usd(self.book.start_cash)}) | 누적실현 {usd(self.book.realized, True)}")

    def stop(self):
        self._stop = True
        if self._ws is not None:
            asyncio.ensure_future(self._ws.close())
        for t in self._tasks:
            if not t.done():
                t.cancel()

    async def _warmup(self):
        n = self.cfg.bar_minutes
        for code in list(self.exch):
            closes, src = [], self.cfg.quote_source
            if src == "kiwoom" and self.api is not None:
                try:
                    closes = (await asyncio.to_thread(self.api.minute_closes, code, self.exch[code], n, 1))[:-1]
                except Exception as e:
                    log(f"{code} 키움 분봉 실패({e}) → 야후로 대체")
            if not closes:
                try:
                    iv = n if n in (1, 2, 5, 15, 30, 60) else 1
                    df = await asyncio.to_thread(load_bars_yf, code, iv, "1mo" if iv >= 15 else "5d")
                    cur = bar_bucket(self.clock(), n).replace(tzinfo=None)
                    closes, src = list(df.loc[df["ts"] < cur, "close"]), "yfinance"
                except Exception as e:
                    log(f"{code} 워밍업 실패({e})")
            s = self.book.get(code)
            s.closes.extend(closes[-400:])
            self.engine.strat.recompute(s)
            if len(s.closes) and not s.last_price:
                s.last_price = s.closes[-1]
            log(f"{code} 워밍업 {len(s.closes)}봉[{src}] RSI {s.rsi:.1f}")
        if self.cfg.trend_ma_days:
            try:
                daily = await asyncio.to_thread(_daily_closes_yf, list(self.exch), self.clock().date())
                for code, closes in daily.items():
                    self.book.get(code).daily.extend(closes[-300:])
                self.engine.refresh_trend()
                ok = [k for k in self.exch if self.book.get(k).trend_ok]
                log(f"일봉 {self.cfg.trend_ma_days}일선 위(매수 가능): {ok}")
            except Exception as e:
                log(f"일봉 추세 데이터 실패({e}) → 오늘은 추세 필터 통과 종목 없음")

    # ---- 시세: 키움 웹소켓 ----
    async def _kiwoom_loop(self):
        import websockets
        backoff, fails = 1, 0
        items = [{"jmcode": k, "stex_tp": v} for k, v in self.exch.items()]
        while not self._stop:
            try:
                token = await asyncio.to_thread(self.api.get_token)
                async with websockets.connect(self.api.ws_url, ping_interval=None, open_timeout=15, max_size=None) as ws:
                    self._ws = ws
                    await ws.send(json.dumps({"trnm": "LOGIN", "token": token}))
                    async for raw in ws:
                        msg = json.loads(raw)
                        trnm = str(msg.get("trnm", "")).upper()
                        if trnm == "PING":
                            await ws.send(raw if isinstance(raw, str) else json.dumps(msg))
                        elif trnm == "LOGIN":
                            if str(msg.get("return_code")) != "0":
                                fails += 1
                                raise KiwoomError(f"웹소켓 로그인 실패: {msg.get('return_msg')}")
                            fails, backoff = 0, 1
                            await ws.send(json.dumps({"trnm": "REG", "grp_no": "1", "refresh": "1",
                                                      "data": [{"item": items, "type": ["FE"]}]}))
                            log(f"웹소켓 로그인 OK → 실시간 등록 {list(self.exch)}")
                        elif trnm == "REG" and str(msg.get("return_code")) != "0":
                            log(f"실시간 등록 실패: {msg.get('return_msg')} (미국 실시간 시세 이용 신청 확인)")
                        elif trnm == "REAL":
                            for d in msg.get("data") or []:
                                if d.get("type") == "FE":
                                    code = self._match(d.get("item"), d.get("values") or {})
                                    px = to_price((d.get("values") or {}).get("10"))
                                    if code and px > 0:
                                        self.engine.on_price(code, px, self.clock())
            except asyncio.CancelledError:
                raise
            except Exception as e:
                if self._stop:
                    break
                if fails >= 3:
                    log("웹소켓 로그인 3회 연속 실패 → 정지")
                    self._stop = True
                    break
                log(f"웹소켓 끊김: {e!r} → {backoff}초 후 재접속")
                await asyncio.sleep(backoff)
                backoff = min(backoff * 2, 60)
            finally:
                self._ws = None

    def _match(self, item, vals) -> str:
        cands = [item.get("jmcode") if isinstance(item, dict) else item, vals.get("9001")]
        for c in cands:
            s = str(c or "").strip().upper()
            if s in self.exch:
                return s
            hits = [k for k in self.exch if s.endswith(k)]
            if hits:
                return max(hits, key=len)
        return ""

    # ---- 시세: 야후 1분봉 폴링 ----
    def _yf_fetch(self):
        import yfinance as yf
        return yf.download(list(self.exch), period="1d", interval="1m", prepost=False, group_by="ticker",
                           auto_adjust=False, progress=False, threads=True)

    async def _yf_loop(self):
        while not self._stop:
            try:
                raw = await asyncio.to_thread(self._yf_fetch)
                now = self.clock()
                for code in self.exch:
                    try:
                        df = raw[code].dropna(subset=["Close"]) if isinstance(raw.columns, pd.MultiIndex) else raw.dropna()
                    except KeyError:
                        continue
                    idx = df.index.tz_convert(ET) if df.index.tz is not None else df.index.tz_localize(ET)
                    for t, r in zip(idx, df.itertuples()):
                        t = t.to_pydatetime()
                        if t + dt.timedelta(minutes=1) > now or (code in self._seen and t <= self._seen[code]):
                            continue                                   # 진행 중인 봉 / 이미 본 봉
                        self._seen[code] = t
                        path = _bar_path(r.Open, r.High, r.Low, r.Close)
                        for k, px in enumerate(path):
                            self.engine.on_price(code, float(px), t + dt.timedelta(seconds=60 * k / len(path)))
            except asyncio.CancelledError:
                raise
            except Exception as e:
                log(f"야후 시세 실패: {e!r}")
            await asyncio.sleep(max(5, 65 - self.clock().second))      # 매분 5초쯤에 갱신

    async def _housekeeping(self):
        last_status = time.monotonic()
        while not self._stop:
            await asyncio.sleep(2)
            now = self.clock()
            self.engine.on_clock(now)
            if time.monotonic() - last_status > self.cfg.status_every_min * 60:
                last_status = time.monotonic()
                print(self.engine.status_df().to_string(index=False))
                log(f"가상자산 {usd(self.book.equity())} | 현금 {usd(self.book.cash)} | 당일실현 {usd(self.book.realized_today, True)}")
            if self.cfg.market_hours_check and hhmm(now) >= self.cfg.session_end:
                log("미국 장 종료 → 정지")
                self.stop()


# %% [markdown]
# ## 9. 과거 실시간 시뮬레이션
# 과거 분봉/시간봉을 **시간 순서대로 한 봉씩**, 봉 안의 가격 경로(시가→저가→고가→종가, 음봉이면 고가 먼저)로 쪼개
# 실시간과 똑같은 `Engine`에 넣습니다. 매일 아침 그날의 국면·섹터·산업 필터를 다시 계산합니다.
# 결과: 거래내역 · 일별 자산 · 요약 지표(수익률·샤프·MDD·승률·손익비)와 SPY 단순보유 비교.

# %%
def load_bars_yf(code, minutes=1, period="7d", start=None, end=None) -> pd.DataFrame:
    """정규장 봉(뉴욕 시간, tz 없음). 야후 한도: 1분 30일, 5분 60일, 60분 730일."""
    import yfinance as yf
    kw = dict(interval=f"{minutes}m" if minutes < 60 else "1h", prepost=False, auto_adjust=False)
    df = yf.Ticker(code).history(start=start, end=end, **kw) if start else yf.Ticker(code).history(period=period, **kw)
    if df.empty:
        raise ValueError(f"{code}: 야후 데이터 없음")
    idx = df.index.tz_convert(ET).tz_localize(None) if df.index.tz is not None else df.index
    out = pd.DataFrame({"ts": idx, "open": df["Open"].values, "high": df["High"].values, "low": df["Low"].values,
                        "close": df["Close"].values, "volume": df["Volume"].values})
    t = out["ts"].dt.time
    return out[(t >= dt.time(9, 30)) & (t < dt.time(16, 0))].dropna().reset_index(drop=True)


def _daily_closes_yf(codes, before: dt.date) -> dict:
    """일봉 종가(before 날짜 전까지)."""
    import yfinance as yf
    raw = yf.download(list(codes), period="2y", interval="1d", group_by="ticker", auto_adjust=False,
                      progress=False, threads=True)
    out = {}
    for c in codes:
        try:
            s = raw[c]["Close"].dropna() if isinstance(raw.columns, pd.MultiIndex) else raw["Close"].dropna()
        except KeyError:
            continue
        idx = s.index.tz_localize(None) if s.index.tz is not None else s.index
        out[c] = [float(v) for d, v in zip(idx, s.values) if d.date() < before]
    return out


def _bar_path(o, h, l, c, steps=3, max_step=0.002, order="auto", rng=None):
    """봉 안의 가격 경로. 한 걸음이 0.2%를 넘지 않게 잘게 나눠 손절·익절 가격을 건너뛰지 않게 한다.
    order: 'auto' = 양봉은 저가 먼저·음봉은 고가 먼저(일반적 가정) / 'high_first' = 항상 고가 먼저 / 'random'"""
    low_first = (c >= o) if order == "auto" else False if order == "high_first" else bool(rng.random() < 0.5)
    pts = (o, l, h, c) if low_first else (o, h, l, c)
    out = [pts[0]]
    for a, b in zip(pts, pts[1:]):
        n = max(steps, math.ceil(abs(b - a) / (a * max_step))) if a > 0 else steps
        out += [a + (b - a) * k / n for k in range(1, n + 1)]
    return out


def simulate(cfg: Config, bars: dict, signals: DailySignals = None, verbose=False, steps=3, seed: dict = None,
             path_order="auto", path_seed=0, entry_filter=None, entry_signal=None, entry_thr=0.0):
    """bars: {티커: DataFrame(ts, open, high, low, close)} — 봉 간격은 cfg.bar_minutes 이하여야 함.
    seed: {티커: {"closes": [전략 봉 종가…], "daily": [일봉 종가…]}} — 시작 전 지표 워밍업(실시간 워밍업과 같은 역할)."""
    frames = []
    for code, df in bars.items():
        d = df[["ts", "open", "high", "low", "close"]].copy()
        d["code"] = code.upper()
        frames.append(d)
    allb = pd.concat(frames, ignore_index=True).sort_values(["ts", "code"])
    gaps = allb["ts"].drop_duplicates().sort_values().diff().dt.total_seconds().div(60)
    interval = int(gaps[gaps > 0].min()) if (gaps > 0).any() else cfg.bar_minutes
    if cfg.bar_minutes < interval:
        raise ValueError(f"데이터가 {interval}분봉인데 전략 봉이 {cfg.bar_minutes}분 — bar_minutes를 {interval} 이상으로")
    cfg_codes = {c.upper() for c in cfg.symbols}
    eng = Engine(cfg, Book(cfg, trade_log=None, verbose=verbose), signals, verbose=verbose)
    eng.entry_filter = entry_filter
    eng.entry_signal, eng.entry_thr = entry_signal, entry_thr
    for code, sd in (seed or {}).items():
        s = eng.book.get(code.upper())
        s.closes.extend(sd.get("closes", [])[-400:])
        s.daily.extend(sd.get("daily", [])[-300:])
        eng.strat.recompute(s)
    rng = np.random.default_rng(path_seed)
    ts_arr = allb["ts"].to_numpy()
    cut = np.flatnonzero(ts_arr[1:] != ts_arr[:-1]) + 1
    o, h, l, c, codes = (allb[k].to_numpy() for k in ("open", "high", "low", "close", "code"))
    for a, b in zip(np.r_[0, cut], np.r_[cut, len(allb)]):
        t0 = pd.Timestamp(ts_arr[a]).to_pydatetime().replace(tzinfo=ET)
        dur = min(interval, (t0.replace(hour=16, minute=0) - t0).total_seconds() / 60)
        pts = []
        for i in range(a, b):
            if codes[i] in cfg_codes:
                p = _bar_path(o[i], h[i], l[i], c[i], steps, order=path_order, rng=rng)
                pts += [(k / len(p), codes[i], v) for k, v in enumerate(p)]
        if not pts:
            continue
        if entry_signal is not None:                       # 같은 봉에 신호가 겹치면 점수 높은 종목부터(알파벳 순 편향 방지)
            sc = {cd: (entry_signal(cd, t0, None) or -1e9) for cd in {p[1] for p in pts}}
            pts.sort(key=lambda x: (x[0], -sc[x[1]]))
        else:
            pts.sort(key=lambda x: x[0])                   # 여러 종목의 경로를 시간 순서로 섞음
        for frac, code, v in pts:
            eng.on_price(code, float(v), t0 + dt.timedelta(minutes=dur * frac))
        eng.on_clock(t0 + dt.timedelta(minutes=dur) - dt.timedelta(seconds=1))
    eng.end_day()
    return {"trades": pd.DataFrame(eng.book.trades), "equity": pd.DataFrame(eng.book.equity_rows),
            "book": eng.book, "interval": interval}


def sim_metrics(res, start=None, end=None, spy_close: pd.Series = None, start_cash=None) -> dict:
    eq = res["equity"].copy()
    if eq.empty:
        return {"거래일": 0}
    eq["날짜"] = pd.to_datetime(eq["날짜"])
    s = eq.set_index("날짜")["자산($)"]
    base = start_cash if start_cash is not None else res["book"].start_cash
    if start is not None:
        before = s[s.index < pd.Timestamp(start)]
        base = float(before.iloc[-1]) if len(before) else base
        s = s[s.index >= pd.Timestamp(start)]
    if end is not None:
        s = s[s.index <= pd.Timestamp(end)]
    if s.empty:
        return {"거래일": 0}
    r = s.pct_change()
    r.iloc[0] = s.iloc[0] / base - 1
    total = s.iloc[-1] / base - 1
    days = len(s)
    curve = pd.concat([pd.Series([base]), s.reset_index(drop=True)])
    mdd = float((curve / curve.cummax() - 1).min())
    tr = res["trades"]
    if len(tr):
        t = pd.to_datetime(tr["시각(ET)"])
        tr = tr[(t >= s.index[0]) & (t < s.index[-1] + pd.Timedelta(days=1))]
    sells = tr[tr["구분"] == "매도"] if len(tr) else tr
    gains, losses = (sells["손익($)"].clip(lower=0).sum(), -sells["손익($)"].clip(upper=0).sum()) if len(sells) else (0, 0)
    f = lambda x, n=2: None if x is None or (isinstance(x, float) and math.isnan(x)) else round(float(x), n)
    out = {"시작": str(s.index[0].date()), "끝": str(s.index[-1].date()), "거래일": days,
           "수익률(%)": f(total * 100),
           "연환산(%)": f(((1 + total) ** (252 / days) - 1) * 100),
           "샤프": f(r.mean() / r.std() * math.sqrt(252)) if r.std() > 0 else 0.0,
           "MDD(%)": f(mdd * 100), "청산횟수": int(len(sells)),
           "승률(%)": f((sells["손익($)"] > 0).mean() * 100, 1) if len(sells) else None,
           "손익비": f(gains / losses) if losses > 0 else None,
           "평균거래(%)": f(sells["수익률(%)"].mean(), 3) if len(sells) else None}
    if spy_close is not None:
        sc = spy_close[(spy_close.index >= s.index[0]) & (spy_close.index <= s.index[-1])]
        prev = spy_close[spy_close.index < s.index[0]]
        if len(sc) and len(prev):
            sr = sc.pct_change()
            sr.iloc[0] = sc.iloc[0] / prev.iloc[-1] - 1
            spy_curve = pd.concat([pd.Series([prev.iloc[-1]]), sc.reset_index(drop=True)])
            out.update({"SPY수익률(%)": f((sc.iloc[-1] / prev.iloc[-1] - 1) * 100),
                        "SPY샤프": f(sr.mean() / sr.std() * math.sqrt(252)),
                        "SPY MDD(%)": f(float((spy_curve / spy_curve.cummax() - 1).min()) * 100)})
    return out


def save_sim(res, out_dir, name, metrics: dict = None, cfg: Config = None):
    os.makedirs(out_dir, exist_ok=True)
    res["trades"].to_csv(os.path.join(out_dir, f"{name}_trades.csv"), index=False, encoding="utf-8-sig")
    res["equity"].to_csv(os.path.join(out_dir, f"{name}_equity.csv"), index=False, encoding="utf-8-sig")
    with open(os.path.join(out_dir, f"{name}_summary.json"), "w", encoding="utf-8") as f:
        json.dump({"metrics": metrics or {}, "config": {k: v for k, v in asdict(cfg).items()
                                                         if k not in ("appkey", "secretkey", "alloc_map")} if cfg else {}},
                  f, ensure_ascii=False, indent=1, default=str)


# %% [markdown]
# ## 10. 분봉 기록 모으기 (분봉 예측 모델 학습용)
# 야후는 5분봉 60일·1분봉 30일까지만 줍니다. 분봉 모델을 제대로 학습하려면 몇 달 이상이 필요합니다.
# - `collect_kiwoom_minutes`: 키움 `usa06011`로 받을 수 있는 만큼 받아 종목별 CSV로 저장.
#   과거를 몇 쪽까지 주는지, 시간 표기가 뉴욕/한국 중 무엇인지 자동 판별하고, 야후 5분봉과 겹치는 구간을 대조해 보고합니다.
# - `archive_yf_minutes`: 야후 1분봉(최근 7일)을 받아 기존 보관 파일에 이어 붙임 — 매주 한 번 돌리면 기록이 계속 쌓입니다.

# %%
def kiwoom_rows_to_bars(rows):
    """키움 분봉 원자료 → (DataFrame[ts, open, high, low, close, volume] 뉴욕 시간, '시간표기')"""
    df = pd.DataFrame(rows)
    if df.empty or "cntr_tm" not in df:
        return pd.DataFrame(columns=["ts", "open", "high", "low", "close", "volume"]), "?"
    tm = df["cntr_tm"].astype(str).str.strip()
    bd = (df["bus_dt"] if "bus_dt" in df else pd.Series([""] * len(df))).astype(str).str.strip()
    full = tm.str.len() >= 14
    raw = np.where(full, tm.str[:14], bd.str[:8] + tm.str[-6:].str.zfill(6))
    ts = pd.Series(pd.to_datetime(raw, format="%Y%m%d%H%M%S", errors="coerce"))
    hrs = ts.dt.hour
    kst = bool(((hrs >= 21) | (hrs <= 7)).mean() > 0.5)       # 미국 정규장이 한국 시간으로 찍힌 경우
    if kst:
        if not full.all():                                       # 영업일자 + 한국 시각: 자정 넘은 봉은 다음 날
            ts = ts + pd.to_timedelta(np.where(hrs <= 12, 1, 0), unit="D")
        ts = ts.dt.tz_localize(KST).dt.tz_convert(ET).dt.tz_localize(None)
    out = pd.DataFrame({"ts": ts, "open": df["open_pric"].map(to_price), "high": df["high_pric"].map(to_price),
                        "low": df["low_pric"].map(to_price), "close": df["cur_prc"].map(to_price),
                        "volume": df["trde_qty"].map(lambda v: abs(to_num(v)))}).dropna(subset=["ts"])
    t = out["ts"].dt.time
    out = out[(t >= dt.time(9, 30)) & (t < dt.time(16, 0)) & (out["close"] > 0)]
    return out.drop_duplicates("ts").sort_values("ts").reset_index(drop=True), ("한국시간" if kst else "뉴욕시간")


def collect_kiwoom_minutes(api: KiwoomUSAPI, symbols: dict, minutes=5, out_dir="kiwoom_minute", days_back=400,
                           max_pages=300, compare_yf=3):
    """종목별로 받을 수 있는 만큼 분봉을 받아 out_dir/<종목>_<분>m.csv.gz 로 저장하고 범위표를 돌려준다."""
    os.makedirs(out_dir, exist_ok=True)
    rows_out = []
    early = (now_et() - dt.timedelta(days=days_back)).strftime("%Y%m%d")
    for i, (code, exch) in enumerate(symbols.items()):
        t0 = time.time()
        try:
            best = None
            for sd in (None, early):          # 시작일자를 '오늘'·'먼 과거' 두 가지로 받아 더 긴 쪽을 씀(의미가 문서에 없어서)
                rows, pages = api.minute_history(code, exch, minutes, sd, max_pages, log_first=(i == 0))
                bars, tz = kiwoom_rows_to_bars(rows)
                n_days = bars["ts"].dt.date.nunique() if len(bars) else 0
                if best is None or n_days > best[3]:
                    best = (bars, tz, pages, n_days, sd or "오늘")
                if n_days >= 60:
                    break
            bars, tz, pages, n_days, sd = best
            bars.to_csv(os.path.join(out_dir, f"{code}_{minutes}m.csv.gz"), index=False)
            rec = {"종목": code, "봉": len(bars), "거래일": n_days, "처음": bars["ts"].min() if len(bars) else None,
                   "마지막": bars["ts"].max() if len(bars) else None, "쪽수": pages, "시간표기": tz, "시작일자": sd,
                   "초": round(time.time() - t0, 1)}
        except Exception as e:
            rec = {"종목": code, "오류": str(e)[:200]}
        if i < compare_yf and rec.get("봉"):
            try:                                  # 야후 5분봉과 겹치는 구간 종가 대조
                y = load_bars_yf(code, minutes, "60d").set_index("ts")["close"]
                k = bars.set_index("ts")["close"]
                both = pd.concat([k, y], axis=1, join="inner").dropna()
                rec["야후대조_겹친봉"] = len(both)
                rec["야후대조_일치율(%)"] = round(((both.iloc[:, 0] / both.iloc[:, 1] - 1).abs() < 0.002).mean() * 100, 1) if len(both) else None
            except Exception as e:
                rec["야후대조_일치율(%)"] = f"실패 {e}"
        rows_out.append(rec)
        log(f"[{i + 1}/{len(symbols)}] {rec}")
    cov = pd.DataFrame(rows_out)
    cov.to_csv(os.path.join(out_dir, "_coverage.csv"), index=False, encoding="utf-8-sig")
    return cov


def archive_yf_minutes(codes, out_dir="yf_minute_archive"):
    """야후 1분봉 최근 7일을 받아 종목별 보관 파일에 이어 붙임(중복 제거). 매주 실행하면 기록이 쌓인다."""
    import yfinance as yf
    os.makedirs(out_dir, exist_ok=True)
    raw = yf.download(list(codes), period="7d", interval="1m", prepost=False, group_by="ticker", auto_adjust=False,
                      progress=False, threads=True)
    rep = {}
    for c in codes:
        try:
            d = raw[c].dropna(subset=["Close"])
        except KeyError:
            continue
        idx = d.index.tz_convert(ET).tz_localize(None)
        new = pd.DataFrame({"ts": idx, "open": d["Open"].values, "high": d["High"].values, "low": d["Low"].values,
                            "close": d["Close"].values, "volume": d["Volume"].values})
        p = os.path.join(out_dir, f"{c}_1m.csv.gz")
        old = pd.read_csv(p, parse_dates=["ts"]) if os.path.exists(p) else new.iloc[:0]
        allb = pd.concat([old, new]).drop_duplicates("ts").sort_values("ts")
        allb.to_csv(p, index=False)
        rep[c] = (len(allb), str(allb["ts"].min())[:10], str(allb["ts"].max())[:10])
    return rep


# %% [local-only]
async def _main():
    cfg = Config(appkey=os.environ.get("KIWOOM_APPKEY", ""), secretkey=os.environ.get("KIWOOM_SECRETKEY", ""))
    api = KiwoomUSAPI(cfg.appkey, cfg.secretkey, cfg.mock) if cfg.quote_source == "kiwoom" else None
    await LiveRunner(cfg, api).run()


if __name__ == "__main__":
    asyncio.run(_main())
