# VERSION: v1.1.0 — 2026-10-07 — K v0.36 숏(실제 −1배 ETF) 따라가기 · AVB→MAA · 2% 미만 K 비중 섹터 ETF로(선택) · M v1.86·S v1.02·I v0.66
# CHANGELOG
#  v1.1.0 (2026-10-07) 리포트 M v1.86 · S v1.02 · I v0.66 · K v0.36 대응
#    - _k_col: K 13c '숏:NVDD' 열 → 티커 'NVDD'(−1배 ETF를 사서 보유 = 숏 노출, 공매도·신용 없음). 이전에는 정규식에 안 맞아 조용히 버려졌음
#    - INVERSE_1X_ETFS(K v0.36 R148 목록, 기초 종목 매핑) · 거래소 표 추가 · −1.5~−2배 ETF(TSLQ·NVDS·CONI·SMCZ·TSLZ 등)는 LEVERAGED_ETFS에
#    - DailySignals.for_day: earn_avoid면 −1배 ETF도 '기초 종목' 실적 발표일에 회피 · 노트에 'K 숏 …' 표시
#    - Config.k_small_to_etf(기본 False): K 비중 < k_min_weight인 종목 몫을 그 섹터 ETF로 모음(00K 풀 섹터표 · 핵심은 산업→섹터)
#      k_small_pool_only=True면 핵심 58종 밖 풀 종목 몫만 옮김(K v0.34 R146 바구니를 원래 ETF 다리로 되돌리는 것과 같음)
#    - REPORT_SHEETS += stock_pool_sector.csv(00K_S&P500풀 티커→섹터 ETF)
#    - PRIOR_STOCK_TO_INDUSTRY/EXCHANGE += MAA(REZ 대표 · K v0.36이 AVB 대신 씀), ESS·AES·AWK·ATO 거래소
#  v1.0.0 (2026-10-05) P8(R1→P8) — M v1.85 · S v1.00 · I v0.65 · K v0.33+ 대응(이 파일에 버전 표기를 붙이기 전 마지막 상태)
# 연구·교육용 가상거래 도구이며 투자 자문이 아닙니다. 실제 주문은 잠겨 있습니다(ALLOW_ORDERS=False).

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
import re
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
    "VZ": "IYZ", "GOOGL": "FDN", "META": "SOCL", "AVB": "REZ", "MAA": "REZ", "EOG": "XOP", "SLB": "XES", "FCX": "XME",
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
    "USB": "NY", "VZ": "NY", "WDC": "ND", "MAA": "NY",   # MAA: K v0.36부터 AVB 대신 주거 리츠(REZ) 대표
}


DEFAULT_UNIVERSE = {k: v for k, v in PRIOR_STOCK_EXCHANGE.items() if k != "AVB"}   # AVB: 야후 데이터 없음
# K v0.33+는 S&P 500 전 종목에서 고름 → 58종 밖 종목도 배분. 2026-10 야후 거래소 정보로 만든 표(없는 종목은 실행 중 조회)
PRIOR_STOCK_EXCHANGE.update({
    "ABBV": "NY", "APH": "NY", "ANET": "NY", "CLX": "NY", "COST": "ND", "DVN": "NY", "EXC": "ND", "FSLR": "ND",
    "IT": "NY", "GDDY": "NY", "IBM": "NY", "JBL": "NY", "KEYS": "NY", "KLAC": "ND", "KR": "NY", "LITE": "ND",
    "MCK": "NY", "MPWR": "ND", "NFLX": "ND", "NWS": "ND", "NEE": "NY", "ON": "ND", "PLTR": "ND", "PM": "NY",
    "SWKS": "ND", "TER": "ND", "VRSN": "ND", "VST": "NY", "ESS": "NY", "AES": "NY", "AWK": "NY", "ATO": "NY",
    **{e: "NA" for e in ("XLB", "XLC", "XLE", "XLF", "XLI", "XLK", "XLP", "XLRE", "XLU", "XLV", "XLY")},   # NYSE Arca(SPY와 같게)
})
# K v0.36(R148) 숏 = 실제 −1배 ETF를 사서 보유(공매도·신용 아님). {ETF: 기초 종목} — 실적 발표 회피는 기초 종목 날짜로
INVERSE_1X_ETFS = {"AAPD": "AAPL", "AMZD": "AMZN", "GGLS": "GOOGL", "MSFD": "MSFT", "NVDD": "NVDA", "TSLS": "TSLA",
                   "METD": "META", "AMDD": "AMD", "PLTD": "PLTR", "MUD": "MU", "NFXS": "NFLX", "ORCS": "ORCL", "SEF": "XLF"}
PRIOR_STOCK_EXCHANGE.update({e: "ND" for e in INVERSE_1X_ETFS if e != "SEF"})   # 2026-10 야후: NGM(나스닥)
PRIOR_STOCK_EXCHANGE["SEF"] = "NA"                                               # PCX(NYSE Arca)
SECTOR_ETFS = ("XLB", "XLC", "XLE", "XLF", "XLI", "XLK", "XLP", "XLRE", "XLU", "XLV", "XLY")
# 레버리지 상품(사용 금지 — 실행기에서 이 종목이 들어오거나 leverage ≠ 1이면 멈춤)
LEVERAGED_ETFS = {"TQQQ", "QLD", "UPRO", "SSO", "SPXL", "TNA", "SOXL", "USD", "TECL", "ROM", "LABU", "FAS", "NVDL",
                  "TSLL", "CONL", "SQQQ", "SOXS", "SPXU", "SDS", "TZA", "FNGU", "MSTU", "MSTX", "NVDU", "TSLT",
                  "TSLQ", "NVDS", "CONI", "SMCZ", "TSLZ", "TSDD", "NVD", "QID", "PSQ2"}   # −1.4배 이상 인버스(K v0.35가 제외한 것)도 금지


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
    leverage: float = 1.0             # 연구용 가상 신용(매수 가능 금액 = 현금 + 자산×(L−1)). 사용 금지 → 1 고정(현금 안에서만)
    margin_rate_pct: float = 7.0      # 빌린 금액(음수 현금)에 붙는 연 이자율 — 거래일마다 1/252씩
    sym_scale: dict = field(default_factory=dict)   # {티커: k} 레버리지 ETF 등: 손절·익절·트레일·반등 폭을 k배
    # --- 낙폭 제어 · 비중 · 물타기 ---
    dd_soft_pct: float = 0.0          # >0: 자산이 고점 대비 이만큼 빠지면 새 매수 금액 × dd_soft_scale
    dd_soft_scale: float = 0.5
    dd_hard_pct: float = 0.0          # >0: 고점 대비 이만큼 빠지면 전부 팔고 dd_cool_days 거래일 쉼(쉬고 나면 고점 다시 잡음)
    dd_cool_days: int = 3
    risk_pct: float = 0.0             # >0: 1회 매수 금액 ≤ 자산 × risk_pct ÷ 손절폭(손절 시 자산 손실이 risk_pct%가 되게)
    add_drop_pct: float = 0.0         # >0: 저점매수 보유 중 첫 매수가 대비 이만큼씩 더 빠지면 추가 매수(물타기)
    add_max: int = 0                  # 물타기 최대 횟수
    add_frac: float = 1.0             # 물타기 1회 수량 = 첫 매수 수량 × add_frac
    off_scale: float = 0.0            # >0: 국면 0인 날에도 저점매수(금액 × off_scale), 그날 장마감에 청산(당일 거래)
    earn_avoid: bool = False          # 실적 발표 당일·전날: 그 종목 새로 안 사고, 들고 있으면 장마감 전에 판다(발표 갭 회피)
    mkt_symbol: str = "SPY"           # 시장 급락 판단용(거래 안 함)
    mkt_stop_pct: float = 0.0         # >0: 시장(SPY)이 전일 종가 대비 이만큼 빠져 있으면 저점매수 새로 안 함
    mkt_flat_pct: float = 0.0         # >0: 시장이 이만큼 빠지면 저점매수 보유분을 팔고 그날 매수 중단
    day_stop_pct: float = 0.0         # >0: 자산(평가 포함)이 그날 시작 대비 이만큼 빠지면 저점매수 보유분을 팔고 그날 매수 중단
    halt_all: bool = False            # True: 위 두 경우에 K·로테이션 보유분도 팔고 그날은 다시 안 삼(다음 날 목표대로 다시 삼)
    max_invest_pct: float = 0.0       # >0: 보유 평가액 합계가 자산의 이 %를 넘지 않게 매수 금액을 줄임(현금 비중 확보)
    decide_on_bar: bool = False       # True: 저점매수 판단·체결을 봉이 바뀌는 첫 가격(시가)과 봉 마감(종가)에서만 — 봉 안 가격 순서 가정과 무관
    stop_cool_days: int = 0           # >0: 손절한 종목은 이 거래일 동안 다시 안 삼
    max_stops_day: int = 0            # >0: 하루 손절이 이 횟수에 닿으면 그날 저점매수 중단
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
    mom_top: int = 0                  # >0이면 그날 매수 대상 중 일봉 mom_days일 수익률 상위 N종만(강한 종목의 눌림목만)
    mom_days: int = 60
    # --- 돌파 매수(추세 추종) ---
    entry_mode: str = "dip"           # 'dip' = 저점매수 / 'brk' = 일봉 brk_days일 최고 종가 돌파 매수 / 'both' = 둘 다
                                      # 'rot' = 모멘텀 로테이션 / 'rot+dip' = 로테이션 + 남는 자리에 저점매수
    brk_days: int = 20
    brk_stop_pct: float = 8.0         # 돌파 매수 손절(매수가 대비)
    brk_trail_pct: float = 10.0       # 돌파 매수 추적 청산(보유 중 고점 대비)
    brk_hold_days: int = 40           # 돌파 매수 최대 보유 거래일
    # --- 모멘텀 로테이션: rot_every거래일마다 일봉 rot_days일 수익률 상위 rot_top종을 같은 비중으로 보유 ---
    rot_top: int = 5
    rot_days: int = 63
    rot_every: int = 5
    rot_pct: float = 20.0             # 로테이션 1종목 비중(가상자산 %)
    rot_stop_pct: float = 0.0         # >0이면 보유 중 고점 대비 이만큼 빠지면 팔고 다음 리밸런싱까지 제외
    rot_keep: int = 0                 # >rot_top이면 들고 있는 종목은 순위가 이 안에 있는 동안 계속 보유(잦은 교체 방지)
    rot_score: str = "ret"            # 'ret' = N일 수익률 / 'sharpe' = N일 로그수익률 ÷ 변동성(꾸준히 오른 종목 우선)
    rot_skip: int = 0                 # 모멘텀 계산에서 최근 k일을 뺌(단기 되돌림 회피)
    rot_regime_exit: bool = True      # True = 국면 0이면 로테이션 종목 전부 매도 / False = 새로 사지만 않고 보유는 유지
    rot_at_open: bool = False         # True = 로테이션 매수도 장 시작 첫 가격(09:30)에 — 매수 시간대(entry_start) 제한 없음
    rot_vol_target: float = 0.0       # >0: 로테이션 1종목 비중 × min(1, 목표 연변동성% ÷ 그 종목 최근 20일 연변동성%)
    k_rot_scale: float = 1.0          # entry_mode 'k+rot': K 따라가기 몫 배율(로테이션은 rot_pct)
    k_min_weight: float = 0.0         # K 몫 새 매수는 K 비중이 이 값(0~1) 이상일 때만(작은 비중 잦은 매매 줄이기)
    k_exit_days: int = 1              # K 비중이 이 거래일 연속 0이어야 K 몫을 팖(1 = 0 되는 날 바로)
    k_top: int = 0                    # >0: 그날 K 비중 상위 N종만 따라감
    k_equal_pct: float = 0.0          # >0: K 몫 종목마다 K 비중과 상관없이 자산의 이 %만큼(k_top과 함께 — K 상위 N종 동일비중)
    k_small_to_etf: bool = False      # True: K 비중 < k_min_weight인 종목 몫을 그 섹터 ETF로 모음(K v0.34+ 작은 바구니 → 섹터 노출 유지)
    k_small_pool_only: bool = False   # True: 위 옮기기를 58종(핵심) 밖 S&P 500 풀 종목에만(핵심 종목 작은 몫은 그대로 = 안 삼)
    k_drop1_max: float = 0.0          # >0: K '다음날 하락확률(%)'이 이 값 이상인 종목은 K 몫으로 새로 사지 않음(K v0.33+)
    rot_base_only: bool = False       # True: 모멘텀 로테이션 후보는 기본 58종만(K가 더한 S&P 500 종목은 K 몫으로만)
    k_fit: bool = False               # True: K 몫 배율을 min(k_rot_scale, (1 − 모멘텀 몫) ÷ 그날 K 비중 합)으로 — 현금 넘침 없이
    batch_buys: bool = False          # True: 같은 순간 들어온 시세는 매도 먼저, 매수는 모아서 모멘텀 → K 비중 큰 순(체결 순서 영향 제거)
    k_universe: bool = False          # True: 실행기가 K 배분 종목(58종 밖 S&P 500·섹터 ETF 포함, k_symbols)을 거래 대상에 더함
    k_mom_days: int = 0               # >0: K 몫 새 매수는 그 종목 일봉 N일 수익률이 양수일 때만(전일 종가까지)
    hold_loser_days: int = 0          # >0: K·로테이션 매도 신호 때 손실 중이면 본전(수수료 포함) 회복을 최대 N거래일 기다림
                                      #     (국면 0·실적 발표 회피 매도는 바로)
    hold_loser_riskoff: bool = False  # True: 국면 0으로 정리할 때도 손실 중이면 본전 대기(hold_loser_days 안에서)
    hold_loser_stop_pct: float = 0.0  # >0: 본전 대기 중이라도 매수가 대비 이만큼 빠지면 바로 팖
    earn_hold_loser: bool = False     # True: 실적 발표 회피 청산 때 손실 중인 보유분은 팔지 않고 둠
    k_take_profit_pct: float = 0.0    # >0: K 몫이 매수가 대비 이만큼 오르면 이익 실현, 그 종목은 k_tp_cool_days 거래일 동안 K 몫으로 다시 안 삼
    k_tp_cool_days: int = 5
    # --- 변동성 맞춤 폭: >0이면 종목의 최근 20일 일간 변동성 ÷ 이 값(%)만큼 손절·익절·트레일·반등 폭을 늘이고 줄임(0.5~3배) ---
    vol_ref_pct: float = 0.0
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
    regime_scale: bool = True         # True = 목표비중만큼 매수 금액을 줄임 / False = 목표비중 > 0이면 전액(켜기·끄기로만 씀)
    sector_filter: bool = True        # S: 섹터 배분 0인 섹터의 종목은 신규 매수 금지
    industry_filter: bool = False     # I: 산업 배분 0인 산업의 종목은 신규 매수 금지(그날 산업 배분이 전부 0이면 적용 안 함)
    stock_filter: bool = False        # K: 주식층 배분비중이 0인 종목은 신규 매수 금지(그날 K 배분이 전부 0이면 적용 안 함)
    k_scale: float = 1.0              # entry_mode 'k'·'k+dip': K 배분비중 × k_scale 만큼 그 종목을 보유
    k_rebalance_pct: float = 0.0      # >0: 보유 중인 K 종목이 그날 목표 금액과 이만큼(%) 넘게 차이 나면 장 시작에 맞춰 사고팖
    k_strict: bool = False            # True: K 배분이 전부 0(현금)인 날은 매수 대상 없음(stock_filter와 함께)
    rot_daily_filter: bool = False    # True: 로테이션 보유 종목이 그날 매수 대상(S·I·K 필터)에서 빠지면 다음 리밸런싱을 기다리지 않고 팖
    rank_col: str = ""                # I 점수로 종목 순위: 'P(상승,ML)' / 'P(부모초과)' / '목표비중' / '' = 안 씀
    watchlist_size: int = 0           # 순위 상위 N종만 그날 매수 대상(0 = 전부)
    alloc_map: dict = field(default_factory=lambda: dict(PRIOR_STOCK_TO_INDUSTRY))
    # --- 기타 ---
    status_every_min: int = 10
    trade_log: str = "paper_trades.csv"
    equity_log: str = "paper_equity.csv"


# 미리 정한 전략 묶음(모두 레버리지 없음: 현금 안에서 · 주식만) — Config(**PRESETS["N1"], ...) 처럼 씀
# 숫자는 과거 실시간 시뮬레이션(시간봉 2023-11 ~ 2026-10, 수수료 0.07%, 58종)
PRESETS = {
    "C0": {},                                                          # 이전 기본값: 시간봉 저점매수
    "A": {"max_trades_per_symbol": 2, "daily_loss_pct": 100.0},        # 저점매수 개선: 종목당 하루 2회 · 일일 손실한도 끔
    "N2": {"entry_mode": "rot+dip", "rot_top": 5, "rot_pct": 20.0, "rot_days": 126, "rot_every": 5,   # 1000% 목표
           "max_positions": 8, "max_trades_per_symbol": 2, "daily_loss_pct": 100.0, "regime_scale": False,
           "trend_ma_days": 0, "sector_filter": False},
    "R1_5000": {"entry_mode": "rot", "rot_top": 1, "rot_pct": 100.0, "max_positions": 1, "rot_keep": 2,  # 5000% 목표(1종목)
                "rot_days": 126, "rot_skip": 21, "rot_every": 5, "rot_at_open": True, "daily_loss_pct": 100.0,
                "regime_scale": False, "max_trades_per_symbol": 2, "trend_ma_days": 0, "sector_filter": False},
    # MDD −5% 이내 · 거의 매일 거래 · 수익 최대(2026-10-04): K 주식층 비중 따라가기 + 저점매수 + 낙폭 제어
    # P1: 봉 안에서 판단 — 시간봉 시뮬레이션은 +943%였지만 5분봉 점검에서 그 수익이 대부분 사라짐(봉 안 가격 순서 가정 효과)
    "P1": {"entry_mode": "k+dip", "max_positions": 30, "position_pct": 10.0, "max_trades_per_symbol": 3,
           "rot_at_open": True, "regime_scale": False, "sector_filter": False, "trend_ma_days": 0, "daily_loss_pct": 100.0,
           "off_scale": 0.5, "mkt_flat_pct": 1.5, "mkt_stop_pct": 1.0, "max_invest_pct": 70.0,
           "stop_cool_days": 3, "max_stops_day": 3},
    # P2(추천): 저점매수는 봉 시가·종가에서만 판단 → 봉 안 가격 순서·데이터 해상도(5분봉)와 무관하게 같은 결과
    "P2": {"entry_mode": "k+dip", "decide_on_bar": True, "k_scale": 0.7, "k_rebalance_pct": 10.0, "rot_at_open": True,
           "max_positions": 30, "position_pct": 4.0, "max_trades_per_symbol": 3, "off_scale": 0.5,
           "regime_scale": False, "sector_filter": False, "trend_ma_days": 0, "daily_loss_pct": 100.0,
           "mkt_flat_pct": 1.5, "mkt_stop_pct": 1.0, "max_invest_pct": 55.0, "dd_soft_pct": 3.0,
           "stop_cool_days": 3, "max_stops_day": 3},
    # P3(추천, 2026-10-04 2차): MDD −10% 이내에서 수익 최대 — K 따라가기 몫 + 6-1 모멘텀 상위 2종 몫, 전부 장 시작 체결
    "P3": {"entry_mode": "k+rot", "k_rot_scale": 1.2, "rot_top": 2, "rot_keep": 4, "rot_pct": 18.0, "rot_days": 126,
           "rot_skip": 21, "rot_every": 5, "rot_at_open": True, "max_positions": 30, "earn_avoid": True,
           "regime_scale": False, "sector_filter": False, "trend_ma_days": 0, "daily_loss_pct": 100.0,
           "max_trades_per_symbol": 2},
    # P4(추천, 2026-10-04 3차): P3에서 K 몫 잦은 매매 줄임(비중 2% 미만은 안 삼 · 비중 0이 3거래일 이어져야 팖) → 승률·수익 ↑
    "P4": {"entry_mode": "k+rot", "k_rot_scale": 0.9, "k_min_weight": 0.02, "k_exit_days": 3,
           "rot_top": 2, "rot_keep": 4, "rot_pct": 15.0, "rot_days": 126, "rot_skip": 21, "rot_every": 5,
           "rot_at_open": True, "max_positions": 30, "earn_avoid": True, "regime_scale": False, "sector_filter": False,
           "trend_ma_days": 0, "daily_loss_pct": 100.0, "max_trades_per_symbol": 2},
    # P4H(수익 우선 대안): P4 + 손실 중 매도 신호면 본전까지 최대 20거래일 대기(국면 1일 때만)
    "P4H": {"entry_mode": "k+rot", "k_rot_scale": 0.9, "k_min_weight": 0.02, "k_exit_days": 3,
            "rot_top": 2, "rot_keep": 4, "rot_pct": 15.0, "rot_days": 126, "rot_skip": 21, "rot_every": 5,
            "rot_at_open": True, "max_positions": 30, "earn_avoid": True, "regime_scale": False, "sector_filter": False,
            "trend_ma_days": 0, "daily_loss_pct": 100.0, "max_trades_per_symbol": 2, "hold_loser_days": 20},
    # P5(추천, 2026-10-04 4차, 승률 80% 목표): 손실 중 매도 신호면 본전까지 최대 20거래일 대기(국면 0·실적 발표 때도,
    #     손실 7% 넘으면 바로 매도) + MDD −10% 안에 들도록 K 배율 0.7 · 모멘텀 각 12%
    "P5": {"entry_mode": "k+rot", "k_rot_scale": 0.7, "k_min_weight": 0.02, "k_exit_days": 3,
           "rot_top": 2, "rot_keep": 4, "rot_pct": 12.0, "rot_days": 126, "rot_skip": 21, "rot_every": 5,
           "rot_at_open": True, "max_positions": 30, "earn_avoid": True, "regime_scale": False, "sector_filter": False,
           "trend_ma_days": 0, "daily_loss_pct": 100.0, "max_trades_per_symbol": 2,
           "hold_loser_days": 20, "hold_loser_riskoff": True, "hold_loser_stop_pct": 7.0, "earn_hold_loser": True},
    # P6(수익 우선, 2026-10-04 5차): P4의 모멘텀 몫을 '168일(1달 건너뜀) 모멘텀 1위 1종목 20%'로 — MDD −10% 안 최고 수익
    "P6": {"entry_mode": "k+rot", "k_rot_scale": 0.9, "k_min_weight": 0.02, "k_exit_days": 3,
           "rot_top": 1, "rot_keep": 2, "rot_pct": 20.0, "rot_days": 168, "rot_skip": 21, "rot_every": 5,
           "rot_at_open": True, "max_positions": 30, "earn_avoid": True, "regime_scale": False, "sector_filter": False,
           "trend_ma_days": 0, "daily_loss_pct": 100.0, "max_trades_per_symbol": 2},
    # P7(추천, 2026-10-04 6차, MDD −15%까지 허용 · 승률 80% 유지 · 수익 최대): K 비중 × 1.5 + 168일 모멘텀 1위 50%
    #     + 손실 중 매도 신호면 본전까지 최대 20거래일 대기(국면 0·실적 발표 때도, 손실 8% 넘으면 바로 매도)
    "P7": {"entry_mode": "k+rot", "k_rot_scale": 1.5, "k_min_weight": 0.02, "k_exit_days": 3,
           "rot_top": 1, "rot_keep": 2, "rot_pct": 50.0, "rot_days": 168, "rot_skip": 21, "rot_every": 5,
           "rot_at_open": True, "max_positions": 30, "earn_avoid": True, "regime_scale": False, "sector_filter": False,
           "trend_ma_days": 0, "daily_loss_pct": 100.0, "max_trades_per_symbol": 2,
           "hold_loser_days": 20, "hold_loser_riskoff": True, "hold_loser_stop_pct": 8.0, "earn_hold_loser": True},
    # P8(추천, 2026-10-05, M v1.85 · S v1.00 · I v0.65 · K v0.33+ 대응 · 2026-10-07 v1.1.0: K v0.36 숏 −1배 ETF도 같은 규칙으로 — 값 그대로가 v4에서도 최고)
    #     : K가 S&P 500에서 고른 종목·섹터 ETF까지 따라감
    #     (모멘텀 후보는 58종) + 같은 순간 매도 먼저·매수는 정해진 순서(체결 순서 영향 제거) + K×1.5 + 168일 모멘텀 1위 35%
    "P8": {"entry_mode": "k+rot", "k_rot_scale": 1.5, "k_min_weight": 0.02, "k_exit_days": 3,
           "rot_top": 1, "rot_keep": 2, "rot_pct": 35.0, "rot_days": 168, "rot_skip": 21, "rot_every": 5,
           "rot_at_open": True, "max_positions": 30, "earn_avoid": True, "regime_scale": False, "sector_filter": False,
           "trend_ma_days": 0, "daily_loss_pct": 100.0, "max_trades_per_symbol": 2,
           "hold_loser_days": 20, "hold_loser_riskoff": True, "hold_loser_stop_pct": 8.0, "earn_hold_loser": True,
           "k_universe": True, "rot_base_only": True, "batch_buys": True},
}
PRESET_ALIAS = {"R1": "P8"}           # Kaggle 실행 셀(STRATEGY="R1")을 바꾸지 않아도 최신 추천 전략으로 실행
PRESET_NOTES = {
    "C0": "시간봉 저점매수 · M 비중·S 필터 — 과거 +274% (MDD −7%)",
    "A": "C0 + 종목당 2회 · 일일 손실한도 끔 — 과거 +369% (MDD −7%)",
    "N2": "126일 모멘텀 상위 5종 로테이션(5거래일마다) + 남는 현금 시간봉 저점매수, M 국면 켜기·끄기 — "
          "과거 +1,724% (MDD −21%), 2023년 S&P100에선 +159%",
    "R1_5000": "6-1 모멘텀 1위 1종목에 전액 — 과거 +8,661% (MDD −35%, 이익의 75%가 SNDK), 2023년 S&P100에선 +737%",
    "P1": "K 주식층 따라가기 + 봉 안 저점매수 — 시간봉 과거 +943%(MDD −4.2%)지만 5분봉 점검에서 수익 대부분 사라짐(참고용)",
    "P2": "K 주식층(v0.31) 비중 × 0.7 매일 맞춤 + 봉 시가·종가 저점매수(4%) + 국면 0인 날 당일거래, 시장 급락 스위치·"
          "투자 상한 55%·낙폭 3%부터 매수 축소·손절 3일 쉬기 — 과거 +274% (MDD −4.9%, 86% 날 거래, 5분봉과 결과 같음)",
    "P3": "K 주식층(v0.31) 비중 × 1.2 따라가기 + 6-1 모멘텀 상위 2종 각 18%, 실적 발표 회피, 장 시작 체결 — "
          "과거 +1,828% (MDD −9.9%, 샤프 3.71). M 국면 의존 큼(SPY 규칙이면 MDD −24%)",
    "P4": "P3 개선: K 비중 × 0.9 따라가기(비중 2% 미만 안 삼 · 비중 0이 3거래일 이어지면 팖) + 6-1 모멘텀 상위 2종 각 15%, "
          "실적 발표 회피, 장 시작 체결 — 과거 +2,043% (MDD −9.65%, 승률 64.9%, 샤프 3.35)",
    "P4H": "P4 + 손실 중 매도 신호면 본전까지 최대 20거래일 대기(국면 1일 때만) — 과거 +2,044% (MDD −9.63%, 승률 70.2%)",
    "P5": "P4 + 손실 중 매도 신호면 본전까지 최대 20거래일 대기(국면 0·실적 발표 때도, 손실 7% 넘으면 바로 매도), "
          "K 배율 0.7 · 모멘텀 각 12% — 과거 +1,134% (MDD −9.91%, 승률 81.5%, 샤프 3.28)",
    "P6": "K 주식층 비중 × 0.9 따라가기 + 168일(1달 건너뜀) 모멘텀 1위 1종목 20%, 실적 발표 회피, 장 시작 체결 — "
          "과거 +2,240% (MDD −9.87%, 승률 65.2%, 샤프 3.40). 모멘텀 기간에 민감(189일이면 MDD −11.2%)",
    "P7": "K 주식층 비중 × 1.5 + 168일(1달 건너뜀) 모멘텀 1위 1종목 50%, 손실 중 매도 신호면 본전까지 최대 20거래일 대기"
          "(국면 0·실적 발표 때도, 손실 8% 넘으면 바로 매도) — 과거 +6,667% (MDD −13.92%, 승률 82.7%, 샤프 3.20). "
          "모멘텀 기간에 민감(189일이면 MDD −20%). ⚠️ 현금이 모자랄 때 티커 알파벳 순서로 먼저 사는 덕을 봄 — "
          "순서 영향을 없애면(batch_buys) MDD −20.7%",
    "P8": "M v1.86 · S v1.02 · I v0.66 · K v0.36 대응: K가 S&P 500에서 고른 종목·섹터 ETF·숏(실제 −1배 ETF)까지 K 비중 × 1.5로 따라감 + "
          "168일 모멘텀 1위(59종 중) 35%, 손실 중 매도 신호면 본전까지 최대 20거래일 대기(손실 8% 넘으면 매도), "
          "같은 순간 매도 먼저·매수는 모멘텀 → K 비중 큰 순 — 과거 +4,732% (MDD −14.61%, 승률 83.6%, 샤프 3.18) · "
          "K v0.33 신호로는 +5,419%(감소분은 K v0.34 예측 바구니 — 2% 미만 몫은 안 삼)",
}


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
    brk_level: float = math.inf  # 돌파 기준가(전일까지 일봉 N일 최고 종가)
    vk: float = 1.0              # 변동성 맞춤 배율(vol_ref_pct 사용 시) — 매일 갱신
    vk_hold: float = 1.0         # 보유 중 포지션은 매수할 때의 배율로 고정
    mode: str = ""               # 보유 중인 포지션의 진입 방식 'dip' / 'brk' / 'rot'
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
    gap_px: float = 0.0          # ±30% 밖 가격이 같은 수준(±3%)으로 3번 연속 오면 진짜 갭으로 받아들임
    gap_n: int = 0
    adds: int = 0                # 물타기 횟수
    first_entry: float = 0.0     # 첫 매수가(물타기 기준)
    init_qty: int = 0            # 첫 매수 수량
    intraday: bool = False       # 국면 0인 날 산 당일 거래 → 장마감 청산
    rebal_day: dt.date = None    # K 비중 맞추기를 한 날(하루 한 번)
    stop_until: dt.date = None   # 손절 뒤 다시 사지 않는 마지막 날(stop_cool_days)
    k_zero: int = 0              # K 몫 보유 중 K 비중이 연속 0인 거래일 수
    exit_wait: dt.date = None    # 손실 중 매도 신호를 처음 받은 날(hold_loser_days)
    tp_until: dt.date = None     # 이익 실현 뒤 K 몫으로 다시 사지 않는 마지막 날


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
        k = (c.sym_scale.get(st.code, 1.0) if c.sym_scale else 1.0) * (st.vk_hold if st.qty > 0 else st.vk)
        if st.qty > 0 and st.mode in ("rot", "kf"):
            st.peak = max(st.peak, px)
            if c.rot_stop_pct and px <= st.peak * (1 - c.rot_stop_pct * k / 100):
                return "SELL", f"로테이션 손절: 고점 {usd(st.peak)} 대비 -{c.rot_stop_pct * k:.1f}% ({(px / st.entry - 1) * 100:+.2f}%)"
            return None, ""
        if st.qty > 0 and st.mode == "brk":
            st.peak = max(st.peak, px)
            gain = px / st.entry - 1
            if px <= st.entry * (1 - c.brk_stop_pct * k / 100):
                return "SELL", f"돌파 손절 {gain * 100:+.2f}%"
            if px <= st.peak * (1 - c.brk_trail_pct * k / 100):
                return "SELL", f"돌파 추적청산: 고점 {usd(st.peak)} 대비 -{c.brk_trail_pct * k:.1f}% ({gain * 100:+.2f}%)"
            if hhmm(now) >= c.force_exit and np.busday_count(st.entry_time.date(), now.date()) >= c.brk_hold_days - 1:
                return "SELL", f"돌파 보유기한 청산 {gain * 100:+.2f}%"
            return None, ""
        if st.qty == 0 and c.entry_mode in ("brk", "both") and px > st.brk_level:
            return "BUY", f"돌파매수: 일봉 {c.brk_days}일 최고 종가 {usd(st.brk_level)} 돌파"
        if st.qty == 0 and c.entry_mode == "brk":
            return None, ""
        if st.qty > 0:
            st.peak = max(st.peak, px)
            gain = px / st.entry - 1
            stop_px = st.entry * (1 - c.stop_loss_pct * k / 100)
            if (c.add_drop_pct and st.adds < c.add_max and st.first_entry and px > stop_px
                    and px <= st.first_entry * (1 - c.add_drop_pct * k * (st.adds + 1) / 100)):
                return "ADD", f"물타기 {st.adds + 1}회: 첫 매수가 대비 {(px / st.first_entry - 1) * 100:+.2f}%"
            if px <= stop_px:
                return "SELL", f"손절 {gain * 100:+.2f}%"
            if hhmm(now) >= c.force_exit:
                if not c.hold_overnight or st.intraday:
                    return "SELL", f"장마감 청산 {gain * 100:+.2f}%"
                if np.busday_count(st.entry_time.date(), now.date()) >= c.max_hold_days - 1:
                    return "SELL", f"보유기한 청산 {gain * 100:+.2f}%"
            if (st.peak / st.entry - 1) * 100 >= c.min_profit_pct * k:
                hot = st.rsi >= c.rsi_sell or (not math.isnan(st.upper) and st.peak >= st.upper)
                trail = c.trail_pct * k * (0.5 if hot else 1.0)
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
        if (px >= st.mid or px >= st.low * (1 + c.max_chase_pct * k / 100)
                or (now - st.armed_at).total_seconds() > c.arm_timeout_min * 60):
            st.armed, st.low = False, math.inf
            return None, "저점 감시 해제"
        if px >= st.low * (1 + c.rebound_pct * k / 100):
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
        self.realized = self.realized_today = self.fees = self.interest = 0.0
        self.day_start_equity = self.cash
        self.halted = False
        self.exposure = 1.0
        self.rot_target, self.rot_out = set(), set()     # 모멘텀 로테이션: 지금 들고 있어야 할 종목 / 손절돼 쉬는 종목
        self.k_weights = {}                              # K 주식층 따라가기: 오늘 목표 비중 {티커: 0~1}
        self.k_scale_eff = None                          # k_fit: 오늘 K 몫 실제 배율(모멘텀 몫을 뺀 나머지에 맞춤)
        self.peak_eq = self.cash                         # 낙폭 제어용 자산 고점
        self.stops_today = 0                             # 오늘 손절 횟수(max_stops_day)
        self.pause_until = None                          # 낙폭 한도로 쉬는 마지막 날
        self.trades, self.equity_rows = [], []
        self.day = None

    def drawdown(self, eq=None) -> float:
        eq = self.equity() if eq is None else eq
        self.peak_eq = max(self.peak_eq, eq)
        return 1 - eq / self.peak_eq if self.peak_eq > 0 else 0.0

    def get(self, code) -> SymState:
        if code not in self.st:
            self.st[code] = SymState(code)
        return self.st[code]

    def equity(self) -> float:
        return self.cash + sum(s.qty * s.last_price for s in self.st.values() if s.qty)

    def open_count(self) -> int:
        return sum(1 for s in self.st.values() if s.qty > 0)

    def new_day(self, day):
        if self.cash < 0:                 # 가상 신용 이자(빌린 금액 × 연이율/252) — 실시간은 매일 새로 켜도 하루 한 번
            i = -self.cash * self.cfg.margin_rate_pct / 100 / 252
            self.cash -= i
            self.interest += i
        self.day = day
        self.realized_today, self.halted, self.stops_today = 0.0, False, 0
        self.day_start_equity = self.equity()
        for s in self.st.values():
            s.buys_today, s.cooldown_until, s.armed, s.low, s.stopped_today = 0, None, False, math.inf, False

    def can_enter(self, code, now, eligible=None, any_time=False) -> tuple:
        c, s = self.cfg, self.get(code)
        if not any_time and not (c.entry_start <= hhmm(now) < c.entry_end):
            return False, "매수 시간대 아님"
        if self.halted:
            return False, "일일 손실한도"
        if self.pause_until is not None and self.day is not None and self.day <= self.pause_until:
            return False, "낙폭 한도로 쉬는 중"
        if self.exposure <= 0 and c.off_scale <= 0:
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
        if s.stop_until is not None and self.day is not None and self.day <= s.stop_until:
            return False, "손절 뒤 쉬는 종목"
        if s.cooldown_until and now < s.cooldown_until:
            return False, "재매수 대기"
        if self.open_count() >= c.max_positions:
            return False, "동시 보유 한도"
        return True, ""

    def order_qty(self, px, pct=None) -> int:
        if px <= 0:
            return 0
        c, eq = self.cfg, self.equity()
        exp = self.exposure if self.exposure > 0 else c.off_scale        # 국면 0: 당일 거래 비중
        budget = eq * (c.position_pct if pct is None else pct) / 100 * exp
        if c.dd_soft_pct and self.drawdown(eq) * 100 >= c.dd_soft_pct:
            budget *= c.dd_soft_scale
        if c.max_invest_pct:                                             # 보유 평가액 상한
            held = sum(s.qty * s.last_price for s in self.st.values() if s.qty)
            budget = min(budget, max(eq * c.max_invest_pct / 100 - held, 0.0))
        power = self.cash + eq * max(self.cfg.leverage - 1, 0.0)      # 레버리지 1이면 현금만
        budget = min(budget, power / (1 + self.cfg.fee_pct / 100))
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
            if s.qty == 0:                                   # 첫 매수(물타기는 보유기간·기준가를 바꾸지 않음)
                s.entry_time, s.first_entry, s.init_qty, s.adds = now, px, qty, 0
            s.qty += qty
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
            if reason.startswith(("손절", "돌파 손절", "로테이션 손절")):
                s.stopped_today = True
                self.stops_today += 1
                if c.stop_cool_days and self.day is not None:
                    s.stop_until = (pd.Timestamp(self.day) + pd.offsets.BDay(c.stop_cool_days)).date()
            if s.qty == 0:
                s.cooldown_until = now + dt.timedelta(seconds=c.cooldown_sec)
                s.peak, s.adds, s.first_entry, s.init_qty, s.intraday, s.k_zero = 0.0, 0, 0.0, 0, False, 0
                s.exit_wait = None
        rec = {"시각(ET)": now.strftime("%Y-%m-%d %H:%M:%S"), "종목": code, "구분": "매수" if side == "BUY" else "매도",
               "수량": qty, "가격($)": round(px, 4),
               "손익($)": None if pnl is None else round(pnl, 2),
               "수익률(%)": None if pnl is None else round(pnl / (qty * s.entry) * 100 if s.entry else 0, 3),
               "가상현금($)": round(self.cash, 2), "사유": reason}
        if side == "SELL" and s.qty == 0:
            s.entry, s.mode = 0.0, ""
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
               "누적수수료($)": round(self.fees, 2), "누적이자($)": round(self.interest, 2), "국면노출": round(self.exposure, 2),
               "보유종목": ",".join(f"{k}:{s.qty}" for k, s in self.st.items() if s.qty)}
        self.equity_rows.append(row)
        return row

    # ---- 저장/불러오기 (실시간 가상계좌) ----
    def save_state(self, path):
        state = {"cash": self.cash, "start_cash": self.start_cash, "realized": self.realized, "fees": self.fees,
                 "interest": self.interest, "rot_target": sorted(self.rot_target), "rot_out": sorted(self.rot_out),
                 "peak_eq": self.peak_eq, "pause_until": self.pause_until.isoformat() if self.pause_until else None,
                 "until": {k: {"stop": s.stop_until.isoformat() if s.stop_until else None,      # 다시 안 사는 기간(손절·이익 실현 뒤)
                               "tp": s.tp_until.isoformat() if s.tp_until else None}
                           for k, s in self.st.items() if s.stop_until or s.tp_until},
                 "positions": {k: {"qty": s.qty, "entry": s.entry, "peak": s.peak, "last": s.last_price, "mode": s.mode,
                                   "vk": s.vk_hold, "adds": s.adds, "first_entry": s.first_entry, "init_qty": s.init_qty,
                                   "intraday": s.intraday, "k_zero": s.k_zero,
                                   "exit_wait": s.exit_wait.isoformat() if s.exit_wait else None,
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
        self.interest = float(st.get("interest", 0))
        self.rot_target, self.rot_out = set(st.get("rot_target", [])), set(st.get("rot_out", []))
        pu = st.get("pause_until")
        self.peak_eq, self.pause_until = float(st.get("peak_eq", self.cash)), dt.date.fromisoformat(pu) if pu else None
        for k, p in (st.get("positions") or {}).items():
            s = self.get(k)
            s.qty, s.entry, s.peak, s.last_price = int(p["qty"]), float(p["entry"]), float(p["peak"]), float(p["last"])
            s.mode, s.vk_hold = p.get("mode", "dip"), float(p.get("vk", 1.0))
            s.adds, s.first_entry = int(p.get("adds", 0)), float(p.get("first_entry", p["entry"]))
            s.init_qty, s.intraday = int(p.get("init_qty", p["qty"])), bool(p.get("intraday", False))
            s.k_zero = int(p.get("k_zero", 0))
            s.exit_wait = dt.date.fromisoformat(p["exit_wait"]) if p.get("exit_wait") else None
        for k, u in (st.get("until") or {}).items():
            s = self.get(k)
            s.stop_until = dt.date.fromisoformat(u["stop"]) if u.get("stop") else None
            s.tp_until = dt.date.fromisoformat(u["tp"]) if u.get("tp") else None
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
    kw: dict = None             # K 주식층 배분비중 {티커: 0~1} (stock_allocation_daily.csv를 쓸 때)
    blackout: set = field(default_factory=set)   # 오늘·다음 거래일에 실적 발표가 있는 종목(earn_avoid)
    kp1: dict = None            # K v0.33+ 종목별 '다음날 하락확률(%)' {티커: %} (열이 있는 종목만)


# 파이프라인 리포트(results/reports/<날짜>/*.xlsx) → 일일 필터 CSV. 열 이름은 이전 CSV와 같음
REPORT_SHEETS = {
    "market_regime_daily.csv": ("market_regime_report_v*.xlsx", "01_일별기록"),
    "sector_allocation_daily.csv": ("sector_regime_report_v*.xlsx", "13c_일별배분비중"),
    "industry_allocation_daily.csv": ("industry_regime_report_v*.xlsx", "13c_일별배분비중"),
    "industry_daily.csv": ("industry_regime_report_v*.xlsx", "01Z_산업일별예측"),
    "stock_allocation_daily.csv": ("stock_regime_report_v*.xlsx", "13c_일별배분비중"),
    "stock_daily.csv": ("stock_regime_report_v*.xlsx", "01Z_주식일별예측"),
    "earnings_dates.csv": ("stock_regime_report_v*.xlsx", "05_어닝이벤트"),   # 종목별 실적 발표일(앞으로 예정된 날 포함)
    "stock_pool_sector.csv": ("stock_regime_report_v*.xlsx", "00K_S&P500풀"),   # K v0.33+ 풀 종목 → 섹터 ETF(k_small_to_etf)
}


def _ver_key(path):
    import re
    m = re.search(r"_v(\d+)\.(\d+)\.(\d+)", os.path.basename(path))
    return tuple(int(x) for x in m.groups()) if m else (0, 0, 0)


def latest_report_dir(root) -> str:
    """root 아래 results/reports/<YYYY-MM-DD>/ 중 가장 최근 폴더(없으면 '')."""
    import glob
    dirs = sorted(d for d in glob.glob(os.path.join(root, "**", "reports", "20??-??-??"), recursive=True) if os.path.isdir(d))
    return dirs[-1] if dirs else ""


def signals_from_reports(report_dir, out_dir) -> dict:
    """최신 국면(M)·섹터(S)·산업(I)·주식(K) 리포트에서 일별 시트를 읽어 CSV로 저장. {파일: '버전 · 마지막 날짜'}"""
    import glob, warnings
    os.makedirs(out_dir, exist_ok=True)
    done = {}
    for name, (pat, sheet) in REPORT_SHEETS.items():
        files = sorted(glob.glob(os.path.join(report_dir, pat)), key=_ver_key)
        if not files:
            continue
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            df = pd.read_excel(files[-1], sheet_name=sheet)
        df.to_csv(os.path.join(out_dir, name), index=False, encoding="utf-8-sig")
        dcol = "발표일" if "발표일" in df.columns else df.columns[0]          # 실적 발표일 표는 첫 열이 티커
        last = pd.to_datetime(df[dcol], errors="coerce", format="mixed").max() if dcol != "티커" else pd.NaT
        done[name] = os.path.basename(files[-1]) + (f" · ~{last:%Y-%m-%d}" if pd.notna(last) else "")
    return done


def _k_col(c):
    """K 13c 열 이름 → 티커. 'ETF_XLK' → 'XLK' (v0.33+), '숏:NVDD' → 'NVDD' (v0.36+ 실제 −1배 ETF, 확률 열도 같은 규칙)."""
    c = str(c)
    if c.startswith("숏:"):
        c = c[2:]
    return c[4:] if c.startswith("ETF_") else c


def _pool_sector(signals_dir) -> dict:
    """{티커: 섹터 ETF} — K 00K 풀 표(있으면) + 핵심 종목(산업 → 섹터)."""
    out = {t: PRIOR_INDUSTRY_TO_SECTOR.get(i) for t, i in PRIOR_STOCK_TO_INDUSTRY.items() if PRIOR_INDUSTRY_TO_SECTOR.get(i)}
    p = os.path.join(signals_dir or "", "stock_pool_sector.csv")
    if os.path.exists(p):
        df = pd.read_csv(p, encoding="utf-8-sig")
        if {"티커", "섹터 ETF"} <= set(df.columns):
            out.update({str(t).upper(): str(e) for t, e in zip(df["티커"], df["섹터 ETF"]) if str(e) in SECTOR_ETFS})
    return out


def k_small_to_etf(w: dict, min_weight: float, sector: dict, keep=()) -> tuple:
    """K 비중 dict에서 min_weight 미만 '종목' 몫을 그 섹터 ETF로 옮김(ETF·−1배 ETF·섹터 모르는 종목은 그대로). → (새 dict, 옮긴 합)"""
    out, moved = dict(w), 0.0
    for t, v in w.items():
        e = sector.get(t)
        if 0 < v < min_weight and e and t not in SECTOR_ETFS and t not in INVERSE_1X_ETFS and t not in keep:
            out[e] = out.get(e, 0.0) + v
            out.pop(t)
            moved += v
    return out, moved


def k_symbols(signals_dir, base=None, since="2023-11-01", min_weight=0.0, small_to_etf=False) -> dict:
    """K 배분표에서 since 이후 비중이 한 번이라도 min_weight 이상인 종목을 base(기본 58종)에 더함 → {티커: 거래소}.
    K v0.33+는 S&P 500 전 종목에서 고르므로 58종 밖 종목·섹터 ETF도 따라가려면 거래 대상에 넣어야 함."""
    out = dict(DEFAULT_UNIVERSE if base is None else base)
    p = os.path.join(signals_dir or "", "stock_allocation_daily.csv")
    if not os.path.exists(p):
        return out
    k = _dated_csv(p)
    k = k[k.index >= pd.Timestamp(since)] if since else k
    k.columns = [_k_col(c) for c in k.columns]
    skip = {"날짜", "구분", "★ 합계", "현금"}
    if small_to_etf:                                  # 작은 몫을 섹터 ETF로 모으면 그 ETF들도 거래 대상
        out.update({e: PRIOR_STOCK_EXCHANGE.get(e, "NA") for e in SECTOR_ETFS})
    for t in k.columns:
        if t in skip or "확률" in t or not re.fullmatch(r"[A-Z][A-Z.\-]{0,6}", t):
            continue
        w = pd.to_numeric(k[t], errors="coerce")
        if (w >= max(min_weight, 1e-9)).any() and t not in out and t not in LEVERAGED_ETFS:
            out[t] = PRIOR_STOCK_EXCHANGE.get(t) or _yf_exchange(t)
    return out


def _yf_exchange(t) -> str:
    """야후 거래소 코드 → 키움 stex_tp(NMS·NGM·NCM→ND, NYQ→NY, PCX→NA). 조회 실패면 'NY'(시세 출처가 yfinance면 쓰이지 않음)."""
    try:
        import yfinance as yf
        e = str(yf.Ticker(t).fast_info.get("exchange") or "")
    except Exception:
        e = ""
    return {"NMS": "ND", "NGM": "ND", "NCM": "ND", "NYQ": "NY", "PCX": "NA", "ASE": "AM"}.get(e, "NY")


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
        self.k = need("stock_allocation_daily.csv") if (cfg.stock_filter or cfg.entry_mode.startswith("k")) else None
        self.sector = {}
        if self.k is not None:                                   # K v0.33+: 'ETF_XLK' → 'XLK' · v0.36+: '숏:NVDD' → 'NVDD'
            self.k.columns = [_k_col(c) for c in self.k.columns]
            self.k = self.k.loc[:, ~self.k.columns.duplicated()]
            if cfg.k_small_to_etf:
                self.sector = _pool_sector(d)
                self._kcols = [t for t in self.k.columns if t not in ("날짜", "구분", "★ 합계", "현금") and "확률" not in t]
        self.earn = {}
        if cfg.earn_avoid:
            p = os.path.join(d, "earnings_dates.csv") if d else ""
            if p and os.path.exists(p):
                e = pd.read_csv(p, encoding="utf-8-sig")
                e["발표일"] = pd.to_datetime(e["발표일"], errors="coerce").dt.date
                for t, g in e.dropna(subset=["발표일"]).groupby(e["티커"].astype(str).str.upper()):
                    self.earn[t] = set(g["발표일"])
            elif strict:
                raise FileNotFoundError("earnings_dates.csv 없음 — 주식층 리포트의 05_어닝이벤트에서 만드세요")
            else:
                self.warn.append("earnings_dates.csv 없음 → 실적 발표 회피 안 함")

    def staleness(self, day) -> list:
        out = []
        for name, df in (("M", self.m), ("S", self.s), ("I", self.i), ("K", self.k)):
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
                keep = {k for k in elig if ind[k] == "MARKET" or ind[k] and         # MARKET = 지수 ETF(섹터 필터 없음)
                        to_num(r.get(f"{PRIOR_INDUSTRY_TO_SECTOR.get(ind[k])} 배분비중"), 0) > 0}
                notes.append(f"S {len(keep)}/{len(elig)}")
                elig = keep
        if self.i is not None:
            r = self._asof(self.i, day)
            if r is not None and to_num(r.get("산업배분 합계"), 0) > 0:
                keep = {k for k in elig if ind[k] == "MARKET" or ind[k] and to_num(r.get(ind[k]), 0) > 0}
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
        kw, kp1 = None, None
        if self.k is not None:
            r = self._asof(self.k, day)
            if r is not None and c.k_small_to_etf and c.k_min_weight > 0:    # 작은 몫 → 섹터 ETF(전체 열에서 먼저 모음)
                full = {t: to_num(r.get(t), 0) for t in self._kcols}
                full, moved = k_small_to_etf({t: v for t, v in full.items() if v > 0}, c.k_min_weight, self.sector,
                                             keep=DEFAULT_UNIVERSE if c.k_small_pool_only else ())
                kw = {t: v for t, v in full.items() if t in self.codes}
                notes.append(f"K 작은몫→ETF {moved:.1%}")
            else:
                kw = {} if r is None else {t: to_num(r.get(t), 0) for t in self.codes if to_num(r.get(t), 0) > 0}
            shorts = {t: v for t, v in kw.items() if t in INVERSE_1X_ETFS}
            if shorts:
                notes.append("K 숏 " + ",".join(f"{t}({INVERSE_1X_ETFS[t]}) {v:.1%}" for t, v in shorts.items()))
            kp1 = {} if r is None else {t: to_num(r.get(f"{t} 다음날 하락확률(%)"), np.nan) for t in self.codes
                                        if f"{t} 다음날 하락확률(%)" in r.index}
            notes.append(f"K {len(kw)}종 {sum(kw.values()):.0%}")
            if c.stock_filter and (kw or c.k_strict):
                elig = {t for t in elig if t in kw}
        if not c.regime_scale:
            exp = 1.0 if exp > 0 else 0.0
        black = set()
        if self.earn:
            d0 = pd.Timestamp(day).date()
            d1 = (pd.Timestamp(day) + pd.offsets.BDay(1)).date()
            black = {t for t in self.codes if self.earn.get(INVERSE_1X_ETFS.get(t, t), set()) & {d0, d1}}   # −1배 ETF = 기초 종목 날짜
            if black:
                notes.append(f"실적발표 회피 {','.join(sorted(black))}")
        return DaySignal(exp, elig, " · ".join(notes), kw, black, kp1)


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
        self.base_eligible = None     # 모멘텀 상위로 줄이기 전의 그날 매수 대상
        self.mkt_px = self.mkt_prev = 0.0   # 시장(SPY) 현재가·전일 종가(급락 판단)
        self.mkt_halt = False               # 오늘 시장 급락·자산 손실로 저점매수 중단
        self.entry_filter = None      # (code, ts, SymState) -> bool : 지표 모델로 매수 신호 거르기(없으면 전부 통과)
        self.entry_signal = None      # (code, ts, SymState) -> float|None : 지표 예측 모델 점수. 있으면 규칙 대신 이걸로 매수
        self.entry_thr = 0.0
        self.pending = []             # batch_buys: 아직 체결 안 한 매수 [(우선순위, 티커, 가격, 시각, 사유)]

    def mkt_change(self) -> float:
        return self.mkt_px / self.mkt_prev - 1 if self.mkt_px and self.mkt_prev else 0.0

    def start_day(self, day):
        self.day = day
        b = self.book
        b.new_day(day)
        self.mkt_halt = False
        if b.pause_until is not None and day > b.pause_until:          # 낙폭 한도 휴식 끝 → 고점을 지금 자산으로
            b.pause_until, b.peak_eq = None, b.equity()
        self.today = self.signals.for_day(day) if self.signals else DaySignal()
        b.exposure = self.today.exposure
        if b.exposure <= 0 and self.cfg.off_scale > 0:   # 국면 0인 날 당일 거래: S·I·K도 '전부 현금'이라 필터 없이 전 종목
            self.today.eligible = None
        self.base_eligible = self.today.eligible
        self.refresh_trend()
        if self.cfg.mom_top > 0:
            self.today.eligible = self.momentum_top(self.base_eligible)
        if self.cfg.entry_mode == "k+rot":                               # K 따라가기 + 모멘텀 로테이션 두 몫
            self.rotation_update(day)
            kw = dict(self.today.kw or {}) if self.today.exposure > 0 else {}
            if self.cfg.k_top > 0:
                kw = dict(sorted(kw.items(), key=lambda x: -x[1])[:self.cfg.k_top])
            b.k_weights = kw
            if self.cfg.k_fit:                       # K 몫 + 모멘텀 몫이 100%를 넘지 않게 K 배율을 줄임(체결 순서와 무관)
                room = max(0.0, 1.0 - self.cfg.rot_pct / 100 * self.cfg.rot_top)
                tot = sum(w for w in kw.values() if w >= self.cfg.k_min_weight)
                b.k_scale_eff = min(self.cfg.k_rot_scale, room / tot) if tot > 0 else self.cfg.k_rot_scale
            for code, s in b.st.items():                                 # K 비중 연속 0일 수(팔지 판단용)
                if s.qty > 0 and s.mode == "kf":
                    s.k_zero = 0 if code in kw else s.k_zero + 1
        elif self.cfg.entry_mode.startswith("rot"):
            self.rotation_update(day)
        elif self.cfg.entry_mode.startswith("k"):                        # K 주식층 비중 따라가기(매일)
            b.k_weights = dict(self.today.kw or {}) if self.today.exposure > 0 else {}
            b.rot_target, b.rot_out = set(b.k_weights), set()
        if self.today.blackout:                                          # 실적 발표 회피: 저점매수 대상에서도 뺌
            base = set(self.codes) if self.today.eligible is None else set(self.today.eligible)
            self.today.eligible = base - self.today.blackout
        if self.verbose:
            n = len(self.codes) if self.today.eligible is None else len(self.today.eligible)
            log(f"📅 {day} | 국면 노출 {self.today.exposure:.0%} | 매수 대상 {n}/{len(self.codes)}종 | "
                f"가상자산 {usd(self.book.equity())} | {self.today.note}")

    def momentum_top(self, eligible):
        """매수 대상 중 일봉 N일 수익률 상위 종목(전일 종가까지 — 그날 정보 없음)."""
        n, base = self.cfg.mom_days, set(self.codes) if eligible is None else set(eligible)
        sc = {}
        for code in base:
            d = self.book.get(code).daily
            if len(d) > n and d[-1 - n] > 0:
                sc[code] = d[-1] / d[-1 - n] - 1
        top = set(sorted(sc, key=lambda k: -sc[k])[:self.cfg.mom_top])
        self.today.note += f" · 모멘텀 상위{len(top)}"
        return top

    def rotation_update(self, day):
        """리밸런싱 날(날짜로 정해 실시간·시뮬레이션이 같음)이거나, 국면이 다시 켜졌는데 목표가 비어 있으면 다시 고른다."""
        b = self.book
        if self.today.exposure <= 0 and not self.cfg.rot_regime_exit:
            return                                  # 국면 0 동안 목표를 그대로 둠(보유 유지, 새 매수 없음)
        due = np.busday_count(dt.date(2000, 1, 3), day) % max(self.cfg.rot_every, 1) == 0
        if due or (self.today.exposure > 0 and not b.rot_target):
            b.rot_target, b.rot_out = self.rotation_pick(), set()

    def rotation_pick(self):
        """전일 종가까지의 일봉 rot_days일 수익률 상위(양수만, 추세 필터 통과 종목만)."""
        if self.today.exposure <= 0:
            return set()
        n, k0 = self.cfg.rot_days, self.cfg.rot_skip
        sc = {}
        for code in (self.codes if self.base_eligible is None else self.base_eligible):
            s = self.book.st.get(code)
            if s is None or (self.cfg.rot_base_only and code not in DEFAULT_UNIVERSE):
                continue
            d = s.daily
            if len(d) > n + k0 and d[-1 - n - k0] > 0 and (not self.cfg.trend_ma_days or s.trend_ok):
                r = d[-1 - k0] / d[-1 - n - k0] - 1
                if r > 0:
                    if self.cfg.rot_score == "k":                     # K 주식층 비중이 큰 순서
                        sc[code] = (self.today.kw or {}).get(code, 0.0)
                    elif self.cfg.rot_score == "sharpe":
                        lr = np.diff(np.log(np.asarray([d[-k - k0] for k in range(n + 1, 0, -1)], float)))
                        sd = float(lr.std())
                        sc[code] = float(lr.sum()) / (sd * math.sqrt(n)) if sd > 0 else 0.0
                    else:
                        sc[code] = r
        ranked = sorted(sc, key=lambda k: -sc[k])
        top = ranked[:self.cfg.rot_top]
        if self.cfg.rot_keep > self.cfg.rot_top:                # 버퍼: 순위 rot_keep 안의 보유 종목은 유지
            inner = set(ranked[:self.cfg.rot_keep])
            keep = [c for c in ranked if c in inner and self.book.st[c].qty > 0 and self.book.st[c].mode == "rot"]
            keep = keep[:self.cfg.rot_top]
            top = keep + [c for c in ranked if c not in keep][:self.cfg.rot_top - len(keep)]
        self.today.note += f" · 로테이션 {','.join(top) or '없음'}"
        return set(top)

    def refresh_trend(self):
        n, nb, vr = self.cfg.trend_ma_days, self.cfg.brk_days, self.cfg.vol_ref_pct
        brk = self.cfg.entry_mode in ("brk", "both")
        for s in self.book.st.values():
            d = s.daily
            if n:
                s.trend_ok = len(d) >= n and d[-1] > sum(d[-k] for k in range(1, n + 1)) / n
            if brk:
                s.brk_level = max(d[-k] for k in range(1, nb + 1)) if len(d) >= nb else math.inf
            if vr > 0:
                if len(d) >= 21:
                    sd = float(np.std([d[-i] / d[-i - 1] - 1 for i in range(1, 21)])) * 100
                    s.vk = min(max(sd / vr, 0.5), 3.0)
                else:
                    s.vk = 1.0

    def end_day(self):
        if self.mkt_px:
            self.mkt_prev = self.mkt_px                # 시장 전일 종가
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
        if code == self.cfg.mkt_symbol and code not in self.cfg.symbols:   # 시장 기준(거래 안 함)
            if not self.mkt_px or 0.7 < px / self.mkt_px < 1.3:
                self.mkt_px = px
            return
        s = self.book.get(code)
        ref = s.last_price or (s.closes[-1] if s.closes else 0)          # 마지막으로 받아들인 가격 기준
        if ref and not (ref / 1.3 < px < ref * 1.3):
            if s.gap_px and abs(px / s.gap_px - 1) < 0.03:
                s.gap_n += 1
            else:
                s.gap_px, s.gap_n = px, 1
            if s.gap_n < 3:                                              # 한두 번 튄 값은 잘못된 틱으로 무시
                s.bad_ticks += 1
                if self.verbose and s.bad_ticks in (1, 100):
                    log(f"⚠️ {code} 이상 틱 무시: {px} (기준 {ref:.2f})")
                return
            if self.verbose:
                log(f"↕️ {code} {px / ref - 1:+.1%} 갭이 이어져 새 가격 수준으로 받아들임 ({ref:.2f} → {px:.2f})")
        s.gap_px, s.gap_n = 0.0, 0
        s.last_price, s.last_ts = px, ts
        closed = self.strat.update_bar(s, ts, px)
        if self.cfg.entry_mode.startswith(("rot", "k")) and self._rotation(code, s, px, ts):
            return
        if self.cfg.decide_on_bar and not closed:          # 봉 안에서는 판단 안 함(봉 시작 가격·봉 마감 on_clock에서만)
            return
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
            c = self.cfg
            if (self.mkt_halt or (c.mkt_stop_pct and self.mkt_change() <= -c.mkt_stop_pct / 100)
                    or (c.max_stops_day and self.book.stops_today >= c.max_stops_day)):
                s.armed, s.low = False, math.inf          # 시장 급락 중에는 개별 종목 저점매수 안 함
                return
            if self.entry_filter is not None and not self.entry_filter(code, ts, s):
                s.armed, s.low = False, math.inf          # 모델이 거른 신호 → 다음 저점 구간을 기다림
                return
            self._fill(code, "BUY", px, ts, reason)
        elif action == "SELL" and s.qty > 0:
            self._fill(code, "SELL", px, ts, reason)
        elif action == "ADD" and s.qty > 0:
            self._fill(code, "ADD", px, ts, reason)

    def _krot(self, code, s, px, ts) -> bool:
        """'k+rot': 모멘텀 로테이션 몫(rot_pct) + K 주식층 따라가기 몫(K 비중 × k_rot_scale). 둘 다 장 시작에 체결."""
        b, c = self.book, self.cfg
        rt = b.rot_target if self.today.exposure > 0 else set()
        kw = b.k_weights
        if s.qty > 0:
            why = ""
            if s.mode == "rot" and code not in rt:
                why = "로테이션 제외(모멘텀 순위 밖 또는 국면 0)"
            elif s.mode == "kf" and code not in kw and (s.k_zero >= c.k_exit_days or self.today.exposure <= 0):
                why = "K몫 제외(주식층 비중 0)"
            elif (s.mode == "kf" and c.k_take_profit_pct and s.entry
                  and px >= s.entry * (1 + c.k_take_profit_pct / 100)):
                s.tp_until = (pd.Timestamp(ts.date()) + pd.offsets.BDay(c.k_tp_cool_days)).date()
                self._fill(code, "SELL", px, ts, f"K몫 이익 실현 {(px / s.entry - 1) * 100:+.1f}%")
                return True
            if why:
                be = s.entry * (1 + 2 * c.fee_pct / 100 + c.sell_fee_pct / 100 + 2 * c.slippage_pct / 100)
                deep = c.hold_loser_stop_pct and px <= s.entry * (1 - c.hold_loser_stop_pct / 100)
                if (c.hold_loser_days and (self.today.exposure > 0 or c.hold_loser_riskoff) and px < be
                        and not deep):                                       # 손실 중이면 본전까지 기다림
                    if s.exit_wait is None:
                        s.exit_wait = ts.date()
                    if np.busday_count(s.exit_wait, ts.date()) < c.hold_loser_days:
                        return True
                    why += f" · 본전 대기 {c.hold_loser_days}일 끝"
                elif deep and c.hold_loser_days:
                    why += f" · 손실 {c.hold_loser_stop_pct:g}% 넘음"
                self._fill(code, "SELL", px, ts, why)
            else:
                s.exit_wait = None
            return True
        if code in b.rot_out or code in self.today.blackout or (self.mkt_halt and c.halt_all):
            return True
        k_ok = code in kw and kw[code] >= c.k_min_weight and not (s.tp_until and ts.date() <= s.tp_until)
        if k_ok and c.k_drop1_max:                                      # K가 내일 하락 가능성을 높게 본 종목은 안 삼
            p1 = (self.today.kp1 or {}).get(code)
            k_ok = not (p1 is not None and p1 == p1 and p1 >= c.k_drop1_max)
        if k_ok and c.k_mom_days:                                       # 상승 추세인 K 종목만
            d, n = s.daily, c.k_mom_days
            k_ok = len(d) > n and d[-1 - n] > 0 and d[-1] > d[-1 - n]
        if (code in rt or k_ok) and b.can_enter(code, ts, None, any_time=c.rot_at_open)[0]:
            if code in rt:
                why, prio = f"로테이션 매수: {c.rot_days}일 모멘텀 상위 {c.rot_top}", (0, 0.0)
            else:
                why, prio = f"K몫 매수: 주식층 비중 {kw[code]:.1%}", (1, -kw[code])
            if c.batch_buys:                       # 같은 순간의 매도를 먼저 다 처리한 뒤 정해진 순서로 삼(flush_pending)
                if all(p[1] != code for p in self.pending):
                    self.pending.append((prio, code, px, ts, why))
            else:
                self._fill(code, "BUY", px, ts, why)
        return True

    def flush_pending(self):
        """batch_buys: 모아 둔 매수를 모멘텀 몫 → K 비중 큰 순서로 체결(시세가 들어온 순서·알파벳 순서와 무관)."""
        if not self.pending:
            return
        todo, self.pending = sorted(self.pending, key=lambda p: (p[0], p[1])), []
        for _, code, px, ts, why in todo:
            s = self.book.get(code)
            if s.qty == 0 and self.book.can_enter(code, ts, None, any_time=self.cfg.rot_at_open)[0]:
                self._fill(code, "BUY", px, ts, why)

    def _rotation(self, code, s, px, ts) -> bool:
        """모멘텀 로테이션 처리. True면 이 틱은 끝(저점매수 판단 안 함)."""
        if self.cfg.entry_mode == "k+rot":
            return self._krot(code, s, px, ts)
        b = self.book
        if self.today.exposure > 0:
            tgt = b.rot_target
        else:                                       # 국면 0: 전부 팔거나(기본) 들고 있는 것만 유지(새로 안 삼)
            tgt = set() if self.cfg.rot_regime_exit else {c for c in b.rot_target if b.st.get(c) and b.st[c].qty > 0}
        if s.qty > 0 and s.mode == "rot":
            if code not in tgt:
                self._fill(code, "SELL", px, ts, "로테이션 제외(모멘텀 순위 밖 또는 국면 0)")
                return True
            if self.cfg.rot_daily_filter and self.base_eligible is not None and code not in self.base_eligible:
                self._fill(code, "SELL", px, ts, "로테이션 제외(오늘 S·I·K 필터 밖)")
                b.rot_out.add(code)
                return True
            c = self.cfg
            if c.k_rebalance_pct > 0 and c.entry_mode.startswith("k") and s.rebal_day != ts.date():
                s.rebal_day = ts.date()                     # 그날 첫 가격에 K 목표 금액으로 맞춤
                want = b.equity() * b.k_weights.get(code, 0) * c.k_scale * (b.exposure if b.exposure > 0 else 0)
                have = s.qty * px
                if want > 0 and abs(have - want) / want * 100 > c.k_rebalance_pct:
                    dq = int(abs(want - have) // px)
                    if want > have:
                        fill = b.sim_price("BUY", px)
                        dq = min(dq, int(max(b.cash, 0) / (fill * (1 + c.fee_pct / 100))))
                        if dq >= 1:
                            b.apply_fill(code, "BUY", dq, fill, ts, f"K 비중 맞춤(늘림 {b.k_weights.get(code, 0):.1%})")
                    elif 1 <= dq < s.qty:
                        b.apply_fill(code, "SELL", dq, b.sim_price("SELL", px), ts,
                                     f"K 비중 맞춤(줄임 {b.k_weights.get(code, 0):.1%})")
                return True
            action, reason = self.strat.decide(s, px, ts)
            if action == "SELL":
                self._fill(code, "SELL", px, ts, reason)
            return True
        if (s.qty == 0 and code in tgt and code not in b.rot_out and code not in self.today.blackout
                and not (self.mkt_halt and self.cfg.halt_all)
                and not (self.cfg.rot_daily_filter and self.base_eligible is not None and code not in self.base_eligible)
                and b.can_enter(code, ts, None, any_time=self.cfg.rot_at_open)[0]):
            if self.cfg.entry_mode.startswith("k"):
                why = f"K 따라 매수: 주식층 비중 {b.k_weights.get(code, 0):.1%}"
            else:
                why = f"로테이션 매수: {self.cfg.rot_days}일 모멘텀 상위 {self.cfg.rot_top}"
            self._fill(code, "BUY", px, ts, why)
            return True
        return self.cfg.entry_mode in ("rot", "k")

    def on_clock(self, ts):
        """틱이 뜸해도 손절·장마감 청산이 되게 주기적으로 호출. 낙폭 한도(dd_hard_pct)도 여기서 확인."""
        b, c = self.book, self.cfg
        if c.dd_hard_pct and b.pause_until is None and b.drawdown() * 100 >= c.dd_hard_pct:
            dd = b.drawdown()
            for code, s in list(b.st.items()):
                if s.qty > 0:
                    self._fill(code, "SELL", s.last_price, ts, f"낙폭 한도 청산(고점 대비 -{dd * 100:.1f}%)")
            b.pause_until = (pd.Timestamp(ts.date()) + pd.offsets.BDay(max(c.dd_cool_days, 0))).date()
            if self.verbose:
                log(f"⛔ 자산 고점 대비 -{dd * 100:.1f}% → 전부 팔고 {b.pause_until}까지 쉼")
            return
        if not self.mkt_halt and (c.mkt_flat_pct or c.day_stop_pct):
            hit = ""
            if c.mkt_flat_pct and self.mkt_change() <= -c.mkt_flat_pct / 100:
                hit = f"시장 급락 {self.mkt_change() * 100:+.1f}%"
            elif c.day_stop_pct and b.equity() <= b.day_start_equity * (1 - c.day_stop_pct / 100):
                hit = f"당일 자산 -{c.day_stop_pct}% 도달"
            if hit:
                self.mkt_halt = True
                for code, s in list(b.st.items()):
                    if s.qty > 0 and (s.mode not in ("rot", "kf") or c.halt_all) and s.last_price > 0:
                        self._fill(code, "SELL", s.last_price, ts, f"{hit} → 청산")
                if self.verbose:
                    log(f"⛔ {hit} → 저점매수 보유분 청산, 오늘 저점매수 중단")
        if self.today.blackout and hhmm(ts) >= c.force_exit:          # 실적 발표 전 장마감 청산
            for code in self.today.blackout:
                s = b.st.get(code)
                if s is not None and s.qty > 0 and s.last_price > 0:
                    if c.earn_hold_loser and s.last_price < s.entry:      # 손실 중이면 발표를 안고 감
                        continue
                    self._fill(code, "SELL", s.last_price, ts, "실적 발표 전 청산")
        for code, s in b.st.items():
            if s.qty > 0 and s.last_price > 0:
                action, reason = self.strat.decide(s, s.last_price, ts)
                if action == "SELL":
                    self._fill(code, "SELL", s.last_price, ts, reason)
                elif action == "ADD":
                    self._fill(code, "ADD", s.last_price, ts, reason)

    def _fill(self, code, side, px, ts, reason):
        s, b, c = self.book.get(code), self.book, self.cfg
        if side == "ADD":                                              # 물타기: 첫 매수 수량 × add_frac (현금 안에서)
            fill = b.sim_price("BUY", px)
            qty = min(int(s.init_qty * c.add_frac), int(max(b.cash, 0) / (fill * (1 + c.fee_pct / 100))))
            if qty >= 1:
                s.adds += 1
                b.apply_fill(code, "BUY", qty, fill, ts, reason)
            else:
                s.adds = c.add_max                                      # 살 돈이 없으면 이번 보유에서는 그만
            return
        if side == "BUY":
            rot = reason.startswith(("로테이션", "K 따라", "K몫"))
            fill = b.sim_price("BUY", px)
            if reason.startswith("K 따라"):
                pct = b.k_weights.get(code, 0) * 100 * c.k_scale
            elif reason.startswith("K몫"):
                scale = b.k_scale_eff if (c.k_fit and b.k_scale_eff is not None) else c.k_rot_scale
                pct = c.k_equal_pct or b.k_weights.get(code, 0) * 100 * scale
            elif rot:
                pct = c.rot_pct
                d = s.daily
                if c.rot_vol_target and len(d) >= 21:                   # 변동성 맞춤: 출렁임이 큰 종목은 비중 축소
                    vol = float(np.std([d[-i] / d[-i - 1] - 1 for i in range(1, 21)])) * math.sqrt(252) * 100
                    if vol > 0:
                        pct *= min(1.0, c.rot_vol_target / vol)
            else:
                pct = c.position_pct
                if c.risk_pct:                                         # 손절 시 자산 손실 = risk_pct%
                    k = (c.sym_scale.get(code, 1.0) if c.sym_scale else 1.0) * s.vk
                    pct = min(pct, c.risk_pct / (c.stop_loss_pct * k) * 100)
            qty = b.order_qty(fill, pct)
            s.armed, s.low = False, math.inf
            if qty < 1:
                return
            s.buys_today += 1
            s.mode = "kf" if reason.startswith("K몫") else "rot" if rot else "brk" if reason.startswith("돌파매수") else "dip"
            s.vk_hold = s.vk
            s.intraday = (not rot) and b.exposure <= 0                 # 국면 0인 날 산 저점매수 = 당일 거래
            b.apply_fill(code, "BUY", qty, fill, ts, reason)
        else:
            if reason.startswith("로테이션 손절"):
                b.rot_out.add(code)
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
        if cfg.mkt_stop_pct or cfg.mkt_flat_pct:
            self.exch.setdefault(cfg.mkt_symbol, "NA")          # 시장 급락 판단용 시세(거래 안 함)
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
        c = self.cfg
        if c.trend_ma_days or c.mom_top or c.entry_mode != "dip" or c.mkt_stop_pct or c.mkt_flat_pct:
            try:
                daily = await asyncio.to_thread(_daily_closes_yf, list(self.exch), self.clock().date())
                for code, closes in daily.items():
                    self.book.get(code).daily.extend(closes[-300:])
                if daily.get(c.mkt_symbol):
                    self.engine.mkt_prev = daily[c.mkt_symbol][-1]       # 시장 전일 종가
                    log(f"시장 기준 {c.mkt_symbol} 전일 종가 {self.engine.mkt_prev:.2f}")
                self.engine.refresh_trend()
                if self.cfg.mom_top:
                    self.engine.today.eligible = self.engine.momentum_top(self.engine.base_eligible)
                if "rot" in self.cfg.entry_mode:
                    self.engine.rotation_update(self.clock().date())
                    log(f"모멘텀 로테이션 목표: {sorted(self.book.rot_target)}")
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
                            self.engine.flush_pending()
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
                self.engine.flush_pending()                       # 이번에 받은 모든 종목 시세를 본 뒤 매수(매도 먼저)
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
            self.engine.flush_pending()
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


def download_hourly_yf(codes, period="730d") -> dict:
    """야후 시간봉(정규장) {티커: DataFrame(ts, open, high, low, close, volume)}.
    상장한 지 730일이 안 된 종목(예: BTSG)은 야후가 기간 요청을 '상장일부터'로 바꿔
    'must be within the last 730 days' 오류로 거부하므로, 빠진 종목만 최근 729일 날짜 범위로 다시 받는다."""
    import logging, yfinance as yf
    logging.getLogger("yfinance").setLevel(logging.CRITICAL)      # 위 거부 메시지는 아래에서 다시 받으므로 숨김

    def clean(d):
        if isinstance(d.columns, pd.MultiIndex):
            d = d.copy()
            d.columns = d.columns.get_level_values(0)
        d = d.dropna(subset=["Close"])
        idx = d.index.tz_convert(ET).tz_localize(None) if d.index.tz is not None else d.index
        df = pd.DataFrame({"ts": idx, "open": d["Open"].values, "high": d["High"].values, "low": d["Low"].values,
                           "close": d["Close"].values, "volume": d["Volume"].values})
        t = df["ts"].dt.time
        return df[(t >= dt.time(9, 30)) & (t < dt.time(16, 0))].drop_duplicates("ts").reset_index(drop=True)

    codes = list(codes)
    raw = yf.download(codes, period=period, interval="1h", prepost=False, group_by="ticker", auto_adjust=False,
                      progress=False, threads=True)
    out = {}
    for c in codes:
        try:
            d = clean(raw[c] if isinstance(raw.columns, pd.MultiIndex) else raw)
        except KeyError:
            continue
        if len(d):
            out[c] = d
    start = (dt.date.today() - dt.timedelta(days=729)).isoformat()
    for c in [c for c in codes if c not in out]:
        try:
            d = clean(yf.download(c, start=start, interval="1h", prepost=False, auto_adjust=False, progress=False))
        except Exception as e:
            log(f"{c}: 시간봉 받기 실패 — 시뮬레이션에서 제외 ({str(e)[:120]})")
            continue
        if len(d):
            out[c] = d
            log(f"{c}: 상장 기간이 짧아 최근 729일로 다시 받음 — {d['ts'].min():%Y-%m-%d}부터 {len(d)}봉")
        else:
            log(f"{c}: 시간봉 없음 — 시뮬레이션에서 제외")
    return out


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
             path_order="auto", path_seed=0, entry_filter=None, entry_signal=None, entry_thr=0.0, tick_order="alpha"):
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
    if cfg.mkt_stop_pct or cfg.mkt_flat_pct:
        cfg_codes.add(cfg.mkt_symbol)                  # 시장 급락 판단용 시세(거래 안 함)
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
            if tick_order == "reverse":                    # 같은 순간 시세의 처리 순서 점검용(기본은 티커 알파벳 순)
                pts.reverse()
            elif tick_order == "random":
                order = rng.permutation(len(pts))
                pts = [pts[i] for i in order]
            pts.sort(key=lambda x: x[0])                   # 여러 종목의 경로를 시간 순서로 섞음
        cur = None
        for frac, code, v in pts:
            if cur is not None and frac != cur:
                eng.flush_pending()                        # 같은 순간의 시세를 다 본 뒤 모아 둔 매수 체결(batch_buys)
            cur = frac
            eng.on_price(code, float(v), t0 + dt.timedelta(minutes=dur * frac))
        eng.flush_pending()
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
