# =============================================================================
#  stock_regime.py
#  VERSION: v0.2.1 - 2026-09-15 - [★★ 내 설계 오류 정정: 점화 빈도를 보지 않고 감축 규칙을 골랐다]
#    ── 회귀 테스트가 v0.2.0을 **출하 전에** 반증했다 ────────────────────────────────────
#      v0.2.0은 CUT_RULES에 세 규칙을 OR로 넣었다 — E3 · P2(vol21 자기이력 **하위 1/3**) ·
#      E2(서프라이즈 자기이력 **하위 1/3**). 정밀−기저가 전부 양수여서 고른 것이다.
#      ★ 그런데 **점화 빈도를 보지 않았다.** 분위 버킷 규칙은 **정의상** 33%의 날에 켜진다.
#        합집합 ≈ 1 − 0.67×0.67×0.85 ≈ **0.6** → **60%의 날을 잘라낸다.**
#      합성 데이터 실측: 평균노출 **0.4942**, 즉 대체하려던 ext200(**0.5866**)보다 **더 낮다.**
#      사용자 요구는 "상승을 타서 비중 최대로"인데 v0.2.0은 **요구와 반대 방향**이었다.
#      교훈(일반화): **정밀−기저는 "이 상태에서 하락이 조금 더 흔하다"만 말한다. "지금 나가라"가
#        아니다.** 분위 버킷은 위험 이벤트가 아니라 **상태 배분기**이므로 라이브 감축 규칙이 될 수
#        없다. 라이브 감축 규칙은 **이벤트 모양**(시작과 끝이 있는 짧은 구간)이어야 한다.
#
#    (K7) ★ 라이브 CUT_RULES를 **E3 하나**로 줄였다. E3은 정밀−기저 **+0.0481**(측정된 것 중 최고,
#         16/25종목)이고 점화가 **발표 후 21거래일**로 한정된 **이벤트**다. 합성 실측 평균노출
#         **0.8075**(ext200 0.6900 → +0.1175) — 사용자 요구와 방향이 맞다.
#    (K8) ★ 신규 `CUT_MIN_AGREE`(기본 1 = OR). **동의 개수**로 감축을 정한다. S 계층의
#         ROTATION_MIN_AGREE=2와 같은 과점화 억제 장치다. 합성 실측 평균노출이 단조롭게 움직인다:
#           동의1 **0.5487** → 동의2 **0.7825** → 동의3 **0.9313**.
#         ⇒ P2·E2는 **퇴출이 아니라 격자로 내렸다**(아래 K9). 판정은 엔진 시트가 한다.
#    (K9) [노출격자] 3행 추가 — `3규칙 OR(구 라이브안)` · `3규칙 동의2` · `빠른방어 동의2`.
#         각 행에 **달력 대조군**이 자동으로 붙는다. ⚠ 내 오프라인 합성 계산에서는 3규칙 OR의 칼마가
#         가장 높게 나왔지만 **그것을 근거로 쓰지 않는다** — 합성 난수에서 주기적 감축이 낙폭을
#         우연히 피한 것이다. 이 원칙을 세 번 어겨서 세 번 틀렸다(R58·R59).
#    (K10) ★ `CUT_EXPOSURE_FLOOR_WARN=0.70` — 라이브 평균노출이 이 선 아래면 **경고 로그**를 남긴다.
#          값은 **자동으로 바꾸지 않는다**(리스크 파라미터를 조용히 건드리지 않는다는 원칙).
#          POS 단계 로그에 규칙별 점화율(`per_rule=E3=0.193;P2=0.333;...`)을 함께 남겨
#          같은 실수가 다시 일어나면 **리포트만 보고** 알 수 있게 했다.
#    (K11) 매 감축 계산에 `감축득표` 열을 남긴다(몇 개 규칙이 동의했는가 · 감사용).
#
#    [검증] ⑫-b v0.2.1 사전등록(⑫는 아래 v0.2.0 항목에 그대로 유효):
#      (g) 라이브(E3만) 평균노출이 **0.70 이상**인가(exposure_below_floor 경고가 없는가).
#      (h) [노출격자]에서 `E3만` vs `3규칙 OR` vs `동의2` 중 **칼마 최고**가 무엇인가 — 그것이
#          다음 라운드의 라이브 후보다. 단 **달력 대조군을 이긴 경우에만** 승격한다.
#      (i) 규칙별 점화율 로그에서 P2·E2가 실제로 0.33 근처인가(내 진단이 실데이터에서도 맞는가).
#
#  VERSION: v0.2.0 - 2026-09-15 - [★★ 라이브 규칙 전환 · 어닝 기반 펀더멘탈 · 빠른 방어 후보 · 격자]
#    REPORT60. 사용자 지시: "buy and hold랑 비교했을 때 하락 기간을 피했는지(비중 최대로 감축) 상승을 타서
#    제대로 수익을 최대한 많이 냈는지(비중 최대로) 판단해서 그렇지 못한 기간에서 문제 원인 찾아내서 개선해"
#
#    ── ★ 첫 실행(리포트: stock_regime_report.xlsx v0.1.0) 판정: 라이브 규칙이 **틀렸다** ──────────
#      전략(200일선) vs 매수보유, 28종목 중위:
#        CAGR **0.0685 vs 0.1400** — BH보다 높은 종목 **1/28**(SLB만)
#        MDD  −0.4270 vs −0.5273  — BH보다 좋은 종목 22/28
#        칼마 **0.165 vs 0.255**  — BH보다 높은 종목 **7/28**
#      ⇒ 하락은 어느 정도 피했지만(MDD 22/28) **상승을 훨씬 더 많이 놓쳐서 칼마마저 나빠졌다.**
#
#      03 시트가 원인을 정확히 짚는다 — **라이브 규칙 P1(200일선 아래)이 측정상 반대 방향이다**:
#        P1 정밀−기저 **−0.0198** · 정밀>기저 종목 **8/28** · 연도비율 **0.147** · 향후수익 격차 **+0.994%p**
#        (200일선 아래일 때 향후 21일 수익이 **더 높다** = 하락 예측력이 없고 반대다)
#      같은 계열 전부 음수: P4 mom63 하위1/3 −0.0271 · P3 vol21 상위1/5 −0.0345 · P5 park5 상위1/5 −0.0391.
#      **추세·고변동성 계열은 개별 주식에서 하락을 예측하지 못한다** — I 계층의 "고변동성·깊은 낙폭은
#      바닥"과 같은 구조다(v0.15.0 브레이크 실패와 동일 교훈).
#      양수인 것은 넷뿐: **E3 발표 후 21일 이내 & 서프라이즈<0 +0.0481**(16/25종목 · 격차 −1.450) ·
#        **P2 vol21 자기이력 하위1/3 +0.0278**(16/28 · 연도비율 0.548 · 격차 −0.876) ·
#        E1 직전 서프라이즈<0 +0.0243(15/26) · **E2 서프라이즈 자기이력 하위1/3 +0.0193(20/28 종목)**.
#
#      ── 19 시트가 메커니즘을 보여준다: **200일선은 양쪽에서 다 늦는다** ──────────────────
#        상승구간 786개 중 **미참여(평균비중≈0) 190개(24%)** · 나머지도 **중위 대응 지연 12.5거래일**
#          CRWD 2020-03-16~07-09 **+255.8%** → 저점 후 **23거래일**·상승분 38% 지점 진입, 벤치대비 −190.8%p
#          TSLA 2019-08-23~2020-02-04 +319.6% → 저점 후 36거래일, −189.2%p
#          TSLA 2023-01-03~02-15 +98.2% → **미참여**, −98.2%p
#          META 2022-11-03~2023-02-07 +115.5% → 저점 후 **61거래일**·상승분 **97% 지점**(다 끝난 뒤), −90.7%p
#        하락도 같은 지연: TSLA 2020-03-04~03-18 **−51.8%를 비중 1.00으로 그대로** 맞았고(10거래일이라
#          200일선이 반응할 시간이 없다) ORCL 2025-10 −37.1%는 감축 지연 27거래일, TGT 2019-12~2020-03은 43거래일.
#      ⇒ 급락 후 반등에서 가격이 200일선을 되찾는 데 수십 거래일이 걸리고, 그때는 이미 상승분의 40~97%가
#        지나 있다. 2020년 3월 코로나 저점이 교과서적 사례다.
#
#    ── ★★ 두 번째 결함: 펀더멘탈이 사실상 측정 불가였다 ────────────────────────────────
#      02 시트 `펀더멘탈 분기수`가 **5~7개**뿐이다. yfinance의 quarterly_income_stmt가 최근 5~6분기만
#      주기 때문이고, YoY는 4분기 전과 비교하므로 거의 전부 NaN이 된다.
#      결과: 03 B블록에 **F 계열이 1종목밖에 없다**(F1·F4·F5·F6은 아예 등장 못 함).
#      ★ 해결책이 리포트 안에 있었다 — **05_어닝이벤트에 실제EPS가 100분기(2002년부터) 있다.**
#        get_earnings_dates()는 **실제 발표일 + 실제 EPS + 서프라이즈**를 25년치 준다(티커당 58~100행).
#      ⇒ (K1) EPS YoY·가속·서프라이즈 추세를 **어닝 원장에서** 만든다. 재무제표는 마진 등 '최근 전용'
#        진단으로만 남긴다(열 이름에 '(최근)'을 붙여 혼동을 막는다).
#
#    (K1) ★ **어닝 기반 펀더멘탈** build_earnings_fundamentals() 신설 — 실제 발표일 색인이므로
#         as-of 조인만 하면 인과가 보장된다(발표일 ≤ t). 만드는 것:
#           EPS_YoY_E(4분기 전 대비) · EPS_가속_E · EPS_4분기합_YoY_E(연환산 · 계절성 제거) ·
#           서프라이즈_4분기평균 · 서프라이즈_연속음수 · 어닝_경과일
#         ⚠ EPS 부호가 바뀌는 경우(적자→흑자)는 YoY가 무의미하므로 **분모에 abs()를 쓰고 |기준|<0.01은 NaN**.
#    (K2) ★★ **라이브 규칙 전환** — `LIVE_RULE: "ext200" → "base1_cut"`.
#         뜻: **기본 비중 1.0**(상승 참여를 최대화 — 사용자 요구 "비중 최대로")에서 시작해
#           **측정상 값이 있는 하락 신호일 때만** 감축한다. 200일선은 쓰지 않는다(측정상 반대 방향).
#         감축 신호(03 시트에서 정밀−기저 > 0인 것만 · CUT_RULES로 교체 가능):
#           E3(발표 후 21일 이내 & 서프라이즈<0) · P2(vol21 자기이력 하위1/3) · E2(서프라이즈 자기이력 하위1/3)
#         감축 폭 CUT_WEIGHT=0.0(전량) · 여러 신호가 겹치면 그대로 0.
#         ⚠ 이것은 **B&H에 가까운 쪽으로 가는** 전환이다. 노출이 0.664 → 0.9 근처로 올라가므로 MDD는
#           나빠질 수 있다. 그 상충을 **격자와 대조군이 판정한다**(아래 K4). 기대: CAGR이 B&H에 붙고
#           칼마가 0.165에서 오른다. 되돌리기: k_overrides={"LIVE_RULE": "ext200"}
#    (K3) **빠른 방어 후보 6종 신규**(200일선이 늦다는 진단의 직접 대응 — 전부 03 시트에서 채점만):
#           P7 20일선 아래 · P8 5일 수익 하위 10% · P9 낙폭 −10% 돌파 · P10 ATR 트레일링 이탈 ·
#           P11 20일선 **하향 교차 당일**(이벤트) · P12 P7 & P2
#         ⚠ 라이브에 넣지 않는다 — v0.1.0의 자리매김(채점 먼저)을 지킨다. 승격 조건은 03 시트 마지막 행.
#    (K4) ★ 신규 격자 **[노출격자]** — 라이브 변형 + 대조군을 06 시트에 나란히 싣는다:
#           `항상 1.0(= B&H)` · `200일선(v0.1.0 라이브)` · `base1_cut(라이브)` ·
#           `base1_cut 절반 감축` · **`달력 대조(같은 감축일수·신호 없음)`**
#         달력 대조가 있어야 "신호가 좋은 날을 골랐나"와 "그냥 노출을 줄여 MDD가 좋아졌나"를 가른다
#         (I 계층에서 이 대조군이 세 번 설계를 살렸다).
#    (K5) **19 블록 C 신설** — 29종목 **동일가중 전략 포트폴리오** vs **동일가중 B&H**. 블록 A가 종목별
#         구간을 보는 동안 블록 C는 "이 계층 전체가 B&H 대비 하락을 피했고 상승을 탔나"를 3행으로 답한다.
#    (K6) AVB(REZ 대표) 누락 조사 — 02 시트에 28종목만 나왔다. 제외 사유를 02 시트 '상태' 열에 남긴다.
#
#    [검증] ⑫ v0.2.0 사전등록(다음 리포트에서 판정):
#      (a) base1_cut의 **CAGR 중위가 B&H의 90% 이상**(0.1400 × 0.9 = 0.126)인가. 미달이면 감축 신호가
#          너무 많이 자르는 것이므로 CUT_RULES를 E3 하나로 줄인다.
#      (b) base1_cut **칼마 중위 > 0.255(B&H)** 인가. 미달이면 라이브를 `항상 1.0`으로 두고(= B&H 인정)
#          감축은 격자에서만 계속 검정한다 — 못 이기는 규칙을 라이브에 두지 않는다.
#      (c) [노출격자]에서 base1_cut이 **달력 대조군**을 칼마로 이기는가. 지면 감축 신호에 값이 없다.
#      (d) F 계열(K1 어닝 기반)이 03 B블록에 **20종목 이상**으로 등장하는가. 미달이면 어닝 원장 파싱을 고친다.
#      (e) 19 블록 C ★★ 종합 행의 부호가 "하락 +, 상승 −"(방어형)에서 개선됐는가.
#      (f) 02 시트에 29종목이 모두(제외면 사유와 함께) 나오는가.
#
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

VERSION = "v0.2.1"
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
    #   [v0.2.1] 60 → 100. K1(어닝 기반 펀더멘탈)이 EPS YoY·TTM YoY·가속을 이 원장으로 만들므로
    #   분기 수가 곧 **측정 가능 구간**이다. 리포트 05 시트에서 실측된 상한이 100분기(2002년부터)였다.
    #   60이면 약 14년, 100이면 약 25년 — 자기이력 분위의 워밍업이 평가창 시작 전에 끝난다.
    #   ⚠ 되돌리기: k_overrides={"EARNINGS_LIMIT": 60}
    EARNINGS_LIMIT: int = 100          # get_earnings_dates(limit=) — 과거 분기를 최대한 받는다

    # ---- 특성 ----
    OWN_PCT_MIN_HIST: int = 250        # 자기이력 백분위 최소 관측(1년)
    HORIZONS: Tuple[int, ...] = (1, 5, 21, 63)
    VOL_WIN: int = 21
    DD_WIN: int = 63
    MA_LONG: int = 200

    # ---- [v0.2.0 K2 ★★ 라이브 규칙 전환] ----
    #   "ext200"(v0.1.0) → **"base1_cut"**. 이유는 파일 헤더 참조 — 200일선은 03 시트 측정에서
    #   정밀−기저 −0.0198 · 8/28종목 · 연도비율 0.147 · 향후수익 격차 **+0.994%p**로 **반대 방향**이었고,
    #   19 시트에서 상승구간 786개 중 미참여 190개 · 중위 대응 지연 12.5거래일로 **양쪽에서 다 늦었다**.
    #   base1_cut의 뜻: **기본 비중 1.0**에서 시작해(사용자 요구 "상승을 타서 비중 최대로")
    #     CUT_RULES의 하락 신호가 켜진 날만 CUT_WEIGHT로 감축한다.
    #   선택지: "base1_cut"(라이브) | "ext200"(v0.1.0) | "always"(항상 1.0 = B&H 대조군)
    #   ⚠ 되돌리기: k_overrides={"LIVE_RULE": "ext200"} 또는 {"LIVE_RULE": "always"}
    LIVE_RULE: str = "base1_cut"
    #   감축에 쓰는 규칙 — **03 시트에서 정밀−기저 > 0으로 측정된 것만** 넣는다(rule_matrix의 키).
    #   리포트 실측(h=21 · 29종목 집계): E3 +0.0481(16/25종목) · P2 +0.0278(16/28 · 연도비율 0.548) ·
    #     E2 +0.0193(20/28종목 — 종목 폭이 가장 넓다). E1(+0.0243)은 E2·E3와 겹쳐 제외.
    #   ⚠ 여기에 P1·P3·P4·P5를 넣지 않는다 — 전부 측정상 음수다(추세·고변동성은 개별 주식 하락을 못 맞춘다).
    #
    #   ★★ v0.2.1 정정 — **점화 빈도를 보지 않고 규칙을 고른 것이 내 설계 오류였다.**
    #     회귀 테스트가 잡았다: 위 3규칙을 OR로 묶으면 평균노출이 **0.4942**로, 대체하려던
    #     ext200(0.5866)보다 **오히려 낮다.** 원인은 명백하다 —
    #       P2 = "vol21 자기이력 **하위 1/3**"  → 정의상 약 33%의 날에 점화
    #       E2 = "서프라이즈 자기이력 **하위 1/3**" → 정의상 약 33%의 날에 점화
    #     세 신호의 합집합은 1 − 0.67×0.67×0.85 ≈ **0.6**, 즉 **60%의 날을 잘라낸다.**
    #     분위 버킷 규칙은 '위험 이벤트'가 아니라 **상태 배분기**다. 정밀−기저가 양수라는 것은
    #     "이 상태에서 하락이 조금 더 흔하다"는 뜻일 뿐, "지금 나가라"는 뜻이 아니다.
    #     사용자 요구가 "상승을 타서 비중 최대로"인데 기본 1.0에서 출발한 규칙이 40%만 들고 있으면
    #     **요구와 반대 방향**이다. 그래서 라이브에는 **이벤트 모양의 규칙만** 남긴다.
    #   라이브 = E3 하나. 이유 세 가지:
    #     (1) 정밀−기저 **+0.0481**로 측정된 것 중 **가장 높다**(16/25종목).
    #     (2) 점화 구간이 **발표 후 21거래일**로 한정된 **이벤트**다 — 분위 버킷과 성질이 다르다.
    #     (3) 향후수익 격차 **−1.450%p**로 방향이 맞다(하락 쪽으로 치우친다).
    #   P2·E2는 **퇴출이 아니라 격자로 내린다** — CUT_MIN_AGREE=2(동시 2개)로 묶으면 점화가
    #     급격히 줄어 라이브 후보가 될 수 있다. 그 판정은 다음 리포트의 [노출격자]가 한다.
    CUT_RULES: Tuple[str, ...] = ("E3 발표 후 21일 이내 & 서프라이즈<0",)
    #   여러 규칙을 쓸 때 **몇 개가 동시에 켜져야** 감축하는가. 1 = OR(합집합) · 2 = 2개 이상 동의.
    #   S 계층의 ROTATION_MIN_AGREE=2와 같은 장치다(그쪽에서 과점화를 막은 것이 검증된 방식).
    CUT_MIN_AGREE: int = 1
    #   ★ 노출 하한 감시 — 라이브 평균노출이 이 값 미만이면 **경고를 남긴다**(자동 변경은 하지 않는다).
    #     사용자 요구가 "비중 최대로"이므로, 감축 규칙이 이 선을 깨면 그 자체가 결함 신호다.
    #     리스크 파라미터를 조용히 바꾸지 않기 위해 경고만 하고 값은 그대로 둔다.
    CUT_EXPOSURE_FLOOR_WARN: float = 0.70
    CUT_WEIGHT: float = 0.0            # 감축일 목표비중(0.0 = 전량 회피 · 0.5 = 절반)
    # [v0.2.0 K4] 노출 격자 — (라벨, LIVE_RULE, CUT_WEIGHT[, CUT_RULES, CUT_MIN_AGREE]).
    #   달력 대조군은 코드가 자동으로 붙인다. 4·5번째 원소는 생략 가능(생략 시 CFG 값).
    #   ⚠ 되돌리기: k_overrides={"EXPOSURE_GRID": ()}
    EXPOSURE_GRID: Tuple[Tuple, ...] = (
        ("항상 1.0(= B&H)", "always", 1.0),
        ("200일선(v0.1.0 라이브)", "ext200", 0.0),
        ("base1_cut E3만 전량(=라이브)", "base1_cut", 0.0),
        ("base1_cut E3만 절반", "base1_cut", 0.5),
        # ★ v0.2.1 신규 — 퇴출된 3규칙 OR을 **격자로 내려** 실제로 나쁜지 리포트가 판정한다.
        #   내 오프라인 계산이 아니라 엔진 시트가 판정한다(이 원칙을 세 번 어겨서 세 번 틀렸다).
        ("base1_cut 3규칙 OR(구 라이브안)", "base1_cut", 0.0,
         ("E3 발표 후 21일 이내 & 서프라이즈<0", "P2 vol21 자기이력 하위1/3",
          "E2 서프라이즈 자기이력 하위1/3"), 1),
        # ★ v0.2.1 신규 — 같은 3규칙을 **동시 2개 동의**로 묶는다(과점화 억제 장치).
        ("base1_cut 3규칙 동의2", "base1_cut", 0.0,
         ("E3 발표 후 21일 이내 & 서프라이즈<0", "P2 vol21 자기이력 하위1/3",
          "E2 서프라이즈 자기이력 하위1/3"), 2),
        # ★ v0.2.1 신규 — 빠른 방어 후보(K3) 중 이벤트 모양 2종을 동의2로. 라이브 아님 · 채점용.
        ("base1_cut 빠른방어 동의2", "base1_cut", 0.0,
         ("P9 낙폭 −10% 돌파", "P11 20일선 하향교차 당일",
          "E3 발표 후 21일 이내 & 서프라이즈<0"), 2),
    )

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
    # ---- [v0.2.0 K3] 빠른 방어 특성 — 200일선이 양쪽에서 늦다는 진단의 대응(전부 t일까지의 정보) ----
    ma20 = px.rolling(20, min_periods=15).mean()
    out["ma20_ext"] = px / ma20 - 1.0
    _above = (px > ma20)
    out["ma20_cross_dn"] = (_above.shift(1).fillna(False) & ~_above).astype(float)   # 하향 교차 당일만 1
    out["ret5"] = px / px.shift(5) - 1.0
    out["ret5_pct"] = _own_pct(out["ret5"], mh)
    # ATR(14) 트레일링 스톱 — 최근 63일 고가에서 3 ATR 아래로 내려온 날. 전부 과거 값만 쓴다.
    _tr = pd.concat([(hi - lo), (hi - px.shift(1)).abs(), (lo - px.shift(1)).abs()], axis=1).max(axis=1)
    atr = _tr.rolling(14, min_periods=10).mean()
    out["atr14"] = atr
    _peak = px.rolling(63, min_periods=30).max()
    out["atr_trail_break"] = ((px < (_peak - 3.0 * atr)) & atr.notna()).astype(float)
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


def build_earnings_fundamentals(rec: Dict[str, Any], idx: pd.DatetimeIndex, cfg: StockConfig
                                ) -> pd.DataFrame:
    """[K1 ★ v0.2.0 신규] **어닝 원장의 실제 EPS로** 펀더멘탈 성장 특성을 만든다.

    왜 이것이 필요한가(v0.1.0 첫 실행의 결함):
      yfinance의 quarterly_income_stmt는 **최근 5~6분기**만 준다(02 시트 '펀더멘탈 분기수' 5~7).
      YoY는 4분기 전과 비교하므로 거의 전부 NaN이 되고, 그래서 03 시트 B블록에 F 계열이
      **1종목**밖에 등장하지 못했다 — '효과 없음'이 아니라 **측정 불가**였다.
      반면 get_earnings_dates()는 **실제 발표일 + 실제 EPS + 서프라이즈를 25년치**(티커당 58~100분기) 준다.

    ★ 인과: 색인이 **실제 발표일**이므로 as-of(발표일 ≤ t) 조인만으로 보장된다. 기간말 추정이 필요 없다 —
      quarterly 재무제표보다 오히려 안전하다.
    ⚠ EPS 부호가 바뀌는 구간(적자→흑자)에서 YoY는 무의미하다. 분모에 abs()를 쓰고
      **|기준 EPS| < 0.01이면 NaN**으로 둔다(추정하지 않는다).
    ⚠ 4분기 합(TTM) YoY도 함께 만든다 — 계절성이 큰 종목(소매·항공)에서 단일 분기 YoY보다 안정적이다."""
    out = pd.DataFrame(index=idx)
    cols = ["EPS_YoY_E", "EPS_가속_E", "EPS_TTM_YoY_E", "서프라이즈_4분기평균", "서프라이즈_연속음수"]
    for c in cols:
        out[c] = np.nan
    earn = rec.get("earn")
    if not isinstance(earn, pd.DataFrame) or not len(earn):
        return out
    e = earn.copy()
    e.index = pd.to_datetime(e.index, errors="coerce")
    e = e[e.index.notna()].sort_index()
    rcol = next((c for c in e.columns if "Reported" in str(c)), None)
    scol = next((c for c in e.columns if "Surprise" in str(c)), None)
    if rcol is None:
        return out
    L = pd.DataFrame(index=e.index)
    L["eps"] = pd.to_numeric(e[rcol], errors="coerce")
    L["sur"] = pd.to_numeric(e[scol], errors="coerce") if scol else np.nan
    # 발표일이 **미래**인 행(다음 예정 실적)은 특성에서 제외한다 — 과거에 알 수 없었다.
    L = L[L.index <= (idx[-1] if len(idx) else L.index[-1])]
    L = L[L["eps"].notna()]
    if len(L) < 6:
        return out
    base = L["eps"].shift(4)
    denom = base.abs()
    L["EPS_YoY_E"] = (L["eps"] - base) / denom.where(denom >= 0.01)
    L["EPS_가속_E"] = L["EPS_YoY_E"] - L["EPS_YoY_E"].shift(1)
    ttm = L["eps"].rolling(4, min_periods=4).sum()
    tbase = ttm.shift(4)
    tden = tbase.abs()
    L["EPS_TTM_YoY_E"] = (ttm - tbase) / tden.where(tden >= 0.04)
    L["서프라이즈_4분기평균"] = L["sur"].rolling(4, min_periods=2).mean()
    _neg = (L["sur"] < 0).astype(float)
    # 연속 음수 카운트(직전까지) — 누적합 차이로 만든다(그날까지의 정보만)
    _grp = (_neg == 0).cumsum()
    L["서프라이즈_연속음수"] = _neg.groupby(_grp).cumsum()
    keep = [c for c in cols if c in L.columns]
    LL = L[keep].copy()
    LL["날짜"] = LL.index
    base_df = pd.DataFrame({"날짜": pd.DatetimeIndex(idx)}).sort_values("날짜")
    m = pd.merge_asof(base_df, LL.sort_values("날짜"), on="날짜", direction="backward")
    m = m.set_index("날짜").reindex(idx)
    for c in keep:
        out[c] = m[c]
    mh = int(cfg.OWN_PCT_MIN_HIST)
    for c in ("EPS_YoY_E", "EPS_TTM_YoY_E", "서프라이즈_4분기평균"):
        out[f"{c}_pct"] = _own_pct(out[c], mh)
    return out


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
    # ---- [v0.2.0 K3] 빠른 방어 후보 — "200일선은 양쪽에서 다 늦다"는 진단의 직접 대응 ----
    #   19 시트 실측: 상승구간 786개 중 미참여 190개 · 중위 대응 지연 **12.5거래일**.
    #   하락에서도 TSLA 2020-03(−51.8%, 10거래일)은 200일선이 반응할 시간이 아예 없었다.
    #   그래서 **더 짧은 창**의 후보를 넣는다. ⚠ 라이브가 아니라 03 시트 채점용이다.
    R["P7 20일선 아래"] = g("ma20_ext") < 0
    R["P8 5일수익 하위10%"] = g("ret5_pct") <= 0.10
    R["P9 낙폭 −10% 돌파"] = g("dd63") <= -0.10
    R["P10 ATR 트레일링 이탈"] = g("atr_trail_break") > 0
    R["P11 20일선 하향교차 당일"] = g("ma20_cross_dn") > 0
    R["P12 P7 & P2"] = (g("ma20_ext") < 0) & (g("vol21_pct") <= 1.0 / 3.0)
    # ---- [v0.2.0 K1] 어닝 기반 펀더멘탈 후보 — quarterly 재무제표(5분기)로는 측정 불가였던 것 ----
    R["G1 EPS YoY(어닝) < 0"] = g("EPS_YoY_E") < 0
    R["G2 EPS TTM YoY(어닝) < 0"] = g("EPS_TTM_YoY_E") < 0
    R["G3 EPS 성장 감속(어닝)"] = g("EPS_가속_E") < 0
    R["G4 서프라이즈 4분기평균 < 0"] = g("서프라이즈_4분기평균") < 0
    R["G5 서프라이즈 2회 연속 음수"] = g("서프라이즈_연속음수") >= 2
    R["G6 G2 & P2"] = (g("EPS_TTM_YoY_E") < 0) & (g("vol21_pct") <= 1.0 / 3.0)
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
def build_positions(feat: pd.DataFrame, cfg: StockConfig,
                    rule: Optional[str] = None, cut_weight: Optional[float] = None,
                    cut_rules: Optional[Tuple[str, ...]] = None,
                    calendar_control: bool = False,
                    cut_min_agree: Optional[int] = None,
                    ticker: str = "") -> pd.DataFrame:
    """[K2 ★ v0.2.0] 목표비중. 체결 규칙은 M·S·I와 같다 — t일 확정, t+1일 집행(exec_w(t)=target_w(t−1)).

    규칙 3종:
      "base1_cut" (라이브) — **기본 1.0**에서 시작해 CUT_RULES의 하락 신호가 켜진 날만 CUT_WEIGHT로 감축.
        사용자 요구("상승을 타서 비중 최대로")를 기본값으로 삼고, **측정상 값이 있는 신호로만** 깎는다.
      "ext200"    (v0.1.0) — 200일선 위 1 / 아래 0. ⚠ 03 시트 측정에서 정밀−기저 −0.0198로 **반대 방향**이었다.
      "always"    (대조군) — 항상 1.0 = 사실상 B&H. 격자의 기준선이다.

    calendar_control=True면 **같은 감축 일수만큼 신호 없이 균등 간격으로** 감축한다(달력 대조군).
      이게 있어야 "신호가 좋은 날을 골랐나"와 "그냥 노출을 줄여 MDD가 좋아졌나"를 가른다 —
      I 계층에서 이 대조군이 세 번 설계를 살렸다(v0.15.0 브레이크·v0.16.0 경고격자·v0.18.0 확신캡).
    """
    out = pd.DataFrame(index=feat.index)
    rule = str(rule if rule is not None else getattr(cfg, "LIVE_RULE", "base1_cut")).lower()
    cw = float(cut_weight if cut_weight is not None else getattr(cfg, "CUT_WEIGHT", 0.0))
    ext = pd.to_numeric(feat.get("ext200"), errors="coerce")
    if rule == "always":
        tgt = pd.Series(1.0, index=feat.index)
        state = pd.Series("보유(항상 1.0)", index=feat.index)
    elif rule == "ext200":
        up = (ext > 0)
        tgt = up.astype(float).where(ext.notna(), np.nan).ffill().fillna(0.0)
        state = pd.Series(np.where(up.fillna(False), "상승", "하락"), index=feat.index)
    else:   # base1_cut
        names = tuple(cut_rules if cut_rules is not None
                      else getattr(cfg, "CUT_RULES", ()) or ())
        agree = int(cut_min_agree if cut_min_agree is not None
                    else getattr(cfg, "CUT_MIN_AGREE", 1))
        agree = max(1, min(agree, max(1, len(names))))
        R = rule_matrix(feat, cfg)
        # ★ v0.2.1 — OR 대신 **동의 개수**로 센다. agree=1이면 OR과 동일하다.
        #   왜 바뀌었나: 분위 버킷 규칙(하위 1/3)은 정의상 33%의 날에 점화되므로 OR로 3개를 묶으면
        #   약 60%의 날을 잘라 평균노출이 0.49까지 떨어졌다 — 대체하려던 ext200(0.59)보다 낮았다.
        #   점화 빈도를 **규칙마다 로그로 남긴다** — 다음에 같은 실수를 반복하지 않기 위해서다.
        votes = pd.Series(0, index=feat.index, dtype=int)
        used: List[str] = []; missing: List[str] = []; freqs: List[str] = []
        for nm in names:
            if nm in R:
                _b = R[nm].reindex(feat.index).fillna(False).astype(bool)
                votes = votes + _b.astype(int)
                used.append(nm); freqs.append(f"{str(nm).split()[0]}={float(_b.mean()):.3f}")
            else:
                missing.append(nm)
        cut = (votes >= agree)
        if missing:
            log("POS", kv(event="cut_rule_not_found", ticker=ticker,
                           missing=";".join(missing),
                           suggest="rule_matrix의 키와 CUT_RULES의 문자열이 정확히 같아야 한다"),
                level="warning")
        if not used:
            log("POS", kv(event="cut_rules_empty", ticker=ticker,
                           note="감축 규칙이 하나도 적용되지 않아 비중이 항상 1.0이다(= B&H)"),
                level="warning")
        log("POS", kv(event="cut_signal", ticker=ticker, rules=len(used), min_agree=agree,
                       fire_rate=round(float(cut.mean()), 4), per_rule=";".join(freqs),
                       calendar_control=int(bool(calendar_control)),
                       note="fire_rate가 0.3을 넘으면 '이벤트'가 아니라 상태 배분기다 — 라이브 부적격"),
            level="debug")
        if calendar_control:
            # 같은 날수를 신호 없이 균등 간격으로 — 난수 없음(재현 가능)
            k = int(cut.sum())
            cal = pd.Series(False, index=feat.index)
            if 0 < k < len(feat):
                pos = np.unique(np.linspace(0, len(feat) - 1, num=k, dtype=int))
                cal.iloc[pos] = True
            cut = cal
        tgt = pd.Series(1.0, index=feat.index).where(~cut, cw)
        state = pd.Series(np.where(cut, "감축", "보유"), index=feat.index)
        out["감축신호"] = cut.astype(int)
        out["감축득표"] = votes.astype(int)       # 몇 개 규칙이 동의했는가(감사용)
        # ★ 노출 하한 감시 — 자동으로 값을 바꾸지 않는다(리스크 파라미터를 조용히 건드리지 않기 위해).
        #   ★ **라이브 설정에만** 경고한다(cut_rules/cut_min_agree를 명시한 호출은 격자다).
        #     격자 행은 일부러 나쁜 설정을 검정하는 것이므로 경고하면 로그가 무의미해진다 —
        #     실제로 첫 실행에서 격자의 '3규칙 OR' 행이 티커마다 경고를 뿜어 라이브 경고를 묻었다.
        _fl = float(getattr(cfg, "CUT_EXPOSURE_FLOOR_WARN", 0.0))
        _mx = float(tgt.mean())
        _is_live = (cut_rules is None and cut_min_agree is None and not calendar_control)
        if _fl > 0.0 and _mx < _fl and _is_live:
            log("POS", kv(event="exposure_below_floor", ticker=ticker,
                           mean_target_w=round(_mx, 4), floor=_fl,
                           fire_rate=round(float(cut.mean()), 4), min_agree=agree,
                           suggest=("CUT_MIN_AGREE를 올리거나 분위 버킷 규칙(하위 1/3 계열)을 "
                                    "CUT_RULES에서 빼고 이벤트형 규칙만 남길 것 — "
                                    "사용자 요구는 '상승을 타서 비중 최대로'다")),
                level="warning")
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
        gf = build_earnings_fundamentals(fund.get(t, {}), df.index, cfg)   # [v0.2.0 K1]
        feat = pd.concat([pf, ff, ef, gf], axis=1)
        feat = feat.loc[feat.index >= pd.Timestamp(cfg.EVAL_START)]
        if len(feat) < 250:
            # [v0.2.0 K6] 제외 사유를 02 시트에 **같은 열 구성으로** 남긴다 — v0.1.0은 열이 달라
            #   AVB가 표에서 통째로 사라졌고, 왜 빠졌는지 리포트만 보고는 알 수 없었다.
            quality.append({"티커": t, "산업ETF": parent_of.get(t, ""), "이름": STOCK_NAME_KR.get(t, ""),
                            "관측일(평가창)": len(feat), "상태": f"⚠ 제외 — 평가창 관측 {len(feat)} < 250",
                            "시작": (str(feat.index[0].date()) if len(feat) else "-"),
                            "종료": (str(feat.index[-1].date()) if len(feat) else "-")})
            log("FEAT", kv(event="ticker_excluded", ticker=t, obs=len(feat),
                           reason="평가창 관측 250일 미만"), level="warning")
            continue
        panel[t] = feat
        pos[t] = build_positions(feat, cfg, ticker=t)
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
                        # [v0.2.0 K1] 어닝 기반 EPS YoY가 실제로 몇 %의 날에 있는가 —
                        #   v0.1.0에서 재무제표 기반 F 계열이 측정 불가였던 것을 여기서 바로 확인한다.
                        "EPS YoY(어닝) 유효비율": (round(float(pd.to_numeric(feat.get("EPS_YoY_E"),
                                                                     errors="coerce").notna().mean()), 4)
                                            if "EPS_YoY_E" in feat.columns else 0.0),
                        "매출 YoY(재무제표) 유효비율": (round(float(pd.to_numeric(feat.get("매출_YoY"),
                                                                        errors="coerce").notna().mean()), 4)
                                               if "매출_YoY" in feat.columns else 0.0),
                        "상태": "정상",
                        "결측 비율(ext200)": round(float(feat["ext200"].isna().mean()), 4)})
    # [v0.2.0 K6] 다운로드 자체가 실패한 티커도 02 시트에 남긴다 — prices에 없으면 위 루프를
    #   아예 통과하지 못해 v0.1.0에서는 '29개 요청 → 28행'의 차이를 리포트로 설명할 수 없었다.
    for t in tickers:
        if t not in prices:
            quality.append({"티커": t, "산업ETF": parent_of.get(t, ""),
                            "이름": STOCK_NAME_KR.get(t, ""), "관측일(평가창)": 0,
                            "상태": "⚠ 제외 — 가격 다운로드 실패(로그 DATA 단계 참조)",
                            "시작": "-", "종료": "-"})
            log("FEAT", kv(event="ticker_excluded", ticker=t, obs=0,
                           reason="가격 다운로드 실패"), level="warning")
    if not panel:
        return {"aborted": True, "note": "평가창 관측이 충분한 티커가 없다", "cfg": cfg,
                "quality": pd.DataFrame(quality)}
    log("FEAT", kv(event="features_built", tickers=len(panel),
                    with_fundamentals=int(sum(1 for t in panel
                                              if pd.to_numeric(panel[t].get("매출_YoY"),
                                                               errors="coerce").notna().any())),
                    with_earnings=int(sum(1 for t in panel
                                          if pd.to_numeric(panel[t].get("어닝_서프라이즈%"),
                                                           errors="coerce").notna().any()))))
    acc = build_rule_accuracy(panel, cfg)
    audit = build_lookahead_audit(prices, fund, cfg)

    # ---- 성과 요약 + [v0.2.0 K4 노출격자] ----
    #   ★ 달력 대조군이 함께 실린다 — "신호가 좋은 날을 골랐나"와 "그냥 노출을 줄여 MDD가 좋아졌나"를
    #     가르는 유일한 방법이다(I 계층에서 이 대조군이 세 번 설계를 살렸다).
    perf: List[dict] = []
    _grid = tuple(getattr(cfg, "EXPOSURE_GRID", ()) or ())
    # (K6 ★) 라벨별로 (티커, 일간수익, 집행비중)을 모아 **동일가중 포트폴리오** 행을 만든다.
    #   왜: 사용자 질문("B&H 대비 하락을 피했고 상승을 탔나")의 단위는 종목이 아니라 계층이다.
    #   종목별 중위 칼마는 각 종목의 좋은 해가 서로 다른 해라서 포트 칼마와 전혀 다른 값이 된다.
    _agg: Dict[str, List[Tuple[str, pd.Series, pd.Series]]] = {}
    for t, p in pos.items():
        r = pd.to_numeric(p["전략일간수익"], errors="coerce").fillna(0.0)
        b = pd.to_numeric(panel[t]["일간수익"], errors="coerce").fillna(0.0)
        _lw = pd.to_numeric(p["집행비중"], errors="coerce")
        # ★ (lbl, 일간수익, **그 행의 집행비중**) — v0.2.0 이전에는 모든 격자 행이 라이브 노출을
        #   찍어서 "평균노출" 열이 전부 같은 값이었다. 격자의 핵심 지표가 노출이므로 치명적이었다.
        _rows: List[Tuple[str, pd.Series, pd.Series]] = [
            (f"★ 라이브({cfg.LIVE_RULE})", r, _lw),
            ("매수보유", b, pd.Series(1.0, index=b.index))]
        for _g in _grid:
            # (라벨, 규칙, 감축비중[, CUT_RULES, CUT_MIN_AGREE]) — 뒤 2개는 생략 가능
            _lbl, _rule, _cw = _g[0], _g[1], _g[2]
            _cr = _g[3] if len(_g) > 3 else None
            _ma = _g[4] if len(_g) > 4 else None
            try:
                _pp = build_positions(panel[t], cfg, rule=_rule, cut_weight=_cw,
                                      cut_rules=_cr, cut_min_agree=_ma, ticker=t)
                _rows.append((f"[노출격자] {_lbl}",
                              pd.to_numeric(_pp["전략일간수익"], errors="coerce").fillna(0.0),
                              pd.to_numeric(_pp["집행비중"], errors="coerce")))
                if str(_rule).lower() == "base1_cut":
                    _pc = build_positions(panel[t], cfg, rule=_rule, cut_weight=_cw,
                                          cut_rules=_cr, cut_min_agree=_ma,
                                          calendar_control=True, ticker=t)
                    _rows.append((f"[노출격자·대조] {_lbl} 같은날수 달력(신호없음)",
                                  pd.to_numeric(_pc["전략일간수익"], errors="coerce").fillna(0.0),
                                  pd.to_numeric(_pc["집행비중"], errors="coerce")))
            except Exception as e:
                log("PERF", kv(event="exposure_grid_row_failed", ticker=t, row=str(_lbl),
                               err=str(e)[:100]), level="warning")
        for lbl, s, _w in _rows:
            _agg.setdefault(lbl, []).append((t, s, _w))
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
                         "평균노출": round(float(_w.mean()), 4)})
    # ---- (K6 ★ 신규) 동일가중 포트폴리오 행 — 06 시트 맨 앞에 붙는다 ----
    _port: List[dict] = []
    for lbl, items in _agg.items():
        if not items:
            continue
        _R = pd.concat([sr.rename(tk) for tk, sr, _ in items], axis=1).sort_index()
        _W = pd.concat([wr.rename(tk) for tk, _, wr in items], axis=1).sort_index()
        _n = max(1, _R.shape[1])
        s = _R.fillna(0.0).mean(axis=1)          # 동일가중(미관측 종목은 현금 0%)
        cur = (1.0 + s).cumprod()
        yrs = max(len(s) / 252.0, 1e-9)
        cagr = float(cur.iloc[-1]) ** (1.0 / yrs) - 1.0
        mdd = float((cur / cur.cummax() - 1.0).min())
        _port.append({"티커": "★ 포트(동일가중)", "이름": f"{_n}종목", "산업ETF": "",
                      "전략": lbl, "총수익배수": round(float(cur.iloc[-1]), 4),
                      "CAGR": round(cagr, 4), "최대낙폭(MDD)": round(mdd, 4),
                      "칼마(CAGR/MDD)": (round(cagr / abs(mdd), 3) if mdd < -1e-9 else None),
                      "연변동성": round(float(s.std() * math.sqrt(252.0)), 4),
                      "샤프": (round(float(s.mean() / s.std() * math.sqrt(252.0)), 3)
                             if float(s.std()) > 0 else None),
                      "일간승률": round(float((s > 0).mean()), 4),
                      "평균노출": round(float(_W.mean(axis=1).mean()), 4)})
    perf = _port + perf
    _pdf = pd.DataFrame(perf) if perf else pd.DataFrame(columns=["티커", "전략", "칼마(CAGR/MDD)"])
    _live = f"★ 라이브({cfg.LIVE_RULE})"
    def _pick_calmar(lbl: str, tick: str = "★ 포트(동일가중)"):
        _m = _pdf[(_pdf["전략"] == lbl) & (_pdf["티커"] == tick)]["칼마(CAGR/MDD)"]
        return (round(float(_m.iloc[0]), 3) if len(_m) and pd.notna(_m.iloc[0]) else None)
    def _med(lbl: str):
        _m = _pdf[(_pdf["전략"] == lbl) & (_pdf["티커"] != "★ 포트(동일가중)")]["칼마(CAGR/MDD)"]
        return (round(float(_m.median(skipna=True)), 3) if len(_m) else None)
    log("PERF", kv(event="performance", rows=len(perf), port_rows=len(_port),
                    live_rule=cfg.LIVE_RULE,
                    port_calmar_live=_pick_calmar(_live), port_calmar_bh=_pick_calmar("매수보유"),
                    median_calmar_live=_med(_live), median_calmar_bh=_med("매수보유"),
                    note="포트 행은 동일가중 · 중위 행은 종목별 — 두 값은 일치하지 않는 것이 정상이다"))
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
            # ---- (K5 ★ 신규) 블록 C — 포트폴리오 구간(동일가중 B&H 대비) ----
            # 왜 필요한가: 블록 A는 "그 **종목**의 상승을 탔나"를 묻는다. 주식 계층은 종목마다
            # 예산이 1.0이라 블록 A도 도달 가능한 벤치를 쓰지만, "이 **계층 전체**가 B&H 대비
            # 하락을 피했고 상승을 탔나"는 별개의 질문이며 그것이 사용자가 묻는 단위다.
            #   · 전략 포트 = 29종목 **동일가중**, 각 종목은 자기 집행비중만큼 참여
            #   · 벤치      = 29종목 **동일가중 B&H**(항상 1.0) — 같은 종목·같은 가중이므로 도달 가능
            #   · 총노출    = 집행비중의 종목 평균(0~1). 1.0이면 전량 보유, 0이면 전량 현금.
            _n = max(1, ret.shape[1])
            _bh_r = ret.fillna(0.0).mean(axis=1)                    # 동일가중 B&H 일간수익
            _st_r = (ret.fillna(0.0) * ex).sum(axis=1) / float(_n)  # 전략 일간수익(미보유=현금 0%)
            _expo = ex.mean(axis=1)                                 # 포트 총노출(0~1)
            segc = I.build_portfolio_segments(
                (1.0 + _st_r).cumprod(), _expo, (1.0 + _bh_r).cumprod(), cfg,
                bench_label=f"{ret.shape[1]}종목 동일가중 B&H", layer="개별주식",
                port_label="★ 전략(동일가중)", M=None)
            if isinstance(segc, pd.DataFrame) and len(segc):
                seg = pd.concat([seg, segc], ignore_index=True) if isinstance(seg, pd.DataFrame) else segc
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
