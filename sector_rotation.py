# =============================================================================
#  sector_rotation.py
#  VERSION: v0.7.0 - 2026-09-06 - [검증 잣대를 전략에 맞춤 ⚠] 신호 채택 통계를 '상위1 스프레드'(평활 순위 1위 섹터의 향후 21일
#    수익 − 상장 평균, NW-t)로 변경(ROTATION_SELECT_STAT="top1"; "ic"=종전) + 수용기준 ⑤ '목표: CAGR ≥ SPY 국면전략(M)' + 비교 변형
#    '리더 자체 목표비중' + MACRO_BETA_FCST 팩터 커버리지 결함 수정 + REL_EXT_200 사전방향 −1→+1(외부 FF49 근거). 3단계 규칙·E_t·M 무변경.
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
#  [v0.4.0] 그 위의 섹터 집중 배분 계층([10b]): 11개 섹터의 절대 국면·점수를 매일 횡단면으로 비교해
#    "상승 확률이 높은 섹터에 비중을 많이"(순위가중, 상한 25%, 하락 국면 제외) 배분한다. 총 노출 E_t는
#    SPY M의 목표비중을 따르고, 순위 신호의 정보 유무는 사전 고정 수용기준(§1.F.3)으로 같은 리포트에서 판정한다.
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
#  v0.7.0 | 2026-09-06 | 사용자 요청(v0.6.0 실데이터 리포트 확인 후): "왜 국면 판단 비중대로 안 따라가? 국면 비중대로 따라가고 그중 상승
#    확률이 가장 높은 섹터를 사야 하는 것 아닌가. 섹터 예측을 잘한다면 SPY 국면전략보다 수익률이 높아야 한다 — 그걸 목표로 개선, 다른
#    문제도 전체 점검".
#    [진단 — 실측 v0.6.0] (1) 비중은 E_t를 정확히 따르고 있었다(13c 섹터+SPY 합계 == E_t, 전 거래일 차이 0). 문제는 '어느 섹터'였다.
#      (2) 채택 신호 RESID_MOM_12_1(외부검증 지원)은 rank IC는 있으나 1위가 저베타 방어섹터(XLP·XLU·XLV)에 몰려 상승장에서 SPY보다
#      뒤졌다(리더 CAGR 16.56%/MDD −9.3% vs SPY M 18.65%/−7.1%). (3) 근본 원인 = 검증 잣대: 평균 rank IC는 '11개 전체 순서'의
#      정확도인데 전략은 '1위 하나'만 산다. 시스템의 상승 확률 SCORE_PCT는 rank IC t≈0.5로 탈락했지만, 같은 리포트 데이터로 잰
#      '상위1 스프레드'(21일 평활 순위 1위의 향후 21일 수익 − 적격 평균)는 +1.25%/21일·NW-t 3.2(무작위 0.4~1.9)였고, SCORE_PCT 1위
#      ×E_t 근사 백테스트(표본내·비용 제외)는 CAGR ≈19.8% vs SPY×E_t 15.5%, 9년 중 8년 우위, 보유/평활 창 변경에도 18.6~21.4%.
#      (4) 전체 점검 결함: MACRO_BETA_FCST 학습 관측일이 36일(팩터 하나의 이력이 짧아 '모든 열 유효' 조건이 그 이후로만 성립),
#      REL_EXT_200 사전방향(−, 평균회귀)이 외부 FF49 전 학습창에서 t≈−8.2로 정반대(= 지속 +8.2).
#    [⚠ 선택 통계 ROTATION_SELECT_STAT="top1"(기본)] rotation_walkforward_select()/_top1_spread_series(): 후보별 0~1 순위(사전방향
#      적용)를 배분과 같은 21일 평활 후 그날 1위 섹터의 향후 h일 수익 − 상장 섹터 평균 → 일별 시계열의 NW-HAC t(lag=h)를 채택 판정에
#      사용. 임계값(엄격 2.0 / 최소 1.0 / 외부 2.0)·모드·최소 관측일·학습창 마감(연초−35일)은 그대로. rank IC t도 계속 계산해 13g에
#      'NW-HAC t(IC)'·'NW-HAC t(상위1)'·'선택 통계' 열로 병기('NW-HAC t' 열 = 채택에 쓴 통계). 외부검증(FF49)도 같은 통계로 계산
#      (13h에 두 t 병기). "ic"로 두면 v0.6 방식. 잡음 신호의 상위1 t는 여전히 ≈N(0,1)이라 오채택 크기는 IC 때와 같다(단위테스트).
#    [수용기준] ③ = 배분에 쓰는 평활 복합순위의 상위1 스프레드 NW-t ≥ 2.0(top1일 때; 참고로 rank IC t 병기). ⑤ 신설 '목표: CAGR ≥
#      SPY 국면전략(M)'(ROTATION_ACCEPT_VS_SPY_M=0 여유) — ①②④가 '균등 대조군 대비 정보 유무', ③이 '1위 판별력', ⑤가 '이 계층을
#      둘 이유'(M만 쓰는 것보다 나은가). 13d에 후보별 상위1 스프레드·t 열, 00시트 '목표 대비(⑤)' 줄 신설. 전 기간 폴백이면 ⑤는 자동
#      PASS(주 전략 = SPY M)이며 ③④는 N/A.
#    [비교 변형] '집중배분(리더 자체 목표비중) [비교]': 리더 보유일 노출을 E_t 대신 그 섹터의 M식 target_pos로(E_t=0인 날은 현금 유지).
#      13 시트 비교용 — 실행 비중(target_w)은 사용자 요청대로 E_t를 따른다. ROTATION_ALT_LEADER_OWN_POS로 끔.
#    [결함 수정] build_rotation_raw_signals(): 팩터 채택을 '관측 500일 초과'에서 '섹터 이력 대비 커버리지 ≥ MACRO_BETA_MIN_COVERAGE
#      (0.8)'로 — 이력 짧은 팩터가 회귀 표본을 잠식하던 문제 해소, 사용/제외 팩터를 로그(macro_beta_factors). 단위테스트: 짧은 팩터 포함
#      시 유효일 198일 → 제외 시 6691일.
#    [⚠ 사전방향] REL_EXT_200 −1 → +1: 근거는 11섹터 표본이 아니라 외부 표본(FF49 1926~, 전 학습창 t≈+8)과 업종 모멘텀 문헌 —
#      사후 데이터 맞춤과 구별해 ROTATION_SIGNAL_SPECS 주석에 명시. 채택 여부는 여전히 워크포워드가 결정.
#    [검증] test_sector_rotation_v070.py 6항목 PASS(손계산·지속형 심은 신호 전 연도 채택/잡음 미채택·인과성·커버리지·비교 변형/⑤·
#      외부 top1). 기존 스위트(v040/v050/v060 — 픽스처가 IC 통계로 보정된 v050/v060은 ROTATION_SELECT_STAT="ic" 명시)·후보·매트릭스·
#      자기검사·v0.3 검증·E2E(13 성과표 10행, 수용기준 6행) PASS.
#    [기대치 — 정직하게] 실데이터에서 SCORE_PCT의 1999~2017 학습창 상위1 t가 2.0(또는 최선 가용 1.0)을 넘는지는 Colab 실행 후 13g로
#      확정된다. 넘으면 리더가 SCORE_PCT 1위(고베타·추세 섹터)로 바뀌어 ⑤ 달성 가능성이 있지만 MDD는 SPY M보다 깊어질 수 있다
#      (집중의 대가). 넘지 못하면 v0.6과 같은 신호가 남거나 폴백(=SPY M)이 된다 — 그 경우 ⑤는 '동률'로 PASS지만 순환매 예측은 못 한 것.
#  v0.6.0 | 2026-09-05 | 사용자 지적(v0.5.1 후): "SPY가 오르면 모든 섹터가 오르나? 아니잖아. 순환매를 판단하도록 해야지 — 왜 더
#    못 하나". 맞는 지적 — 연도별 최고·최저 섹터 수익 차이는 2022년 102%p(XLE +64/XLC −38), 2020년 76%p, 2023년 63%p로 거대하다.
#    v0.5의 문제는 '분산이 없다'가 아니라 검정 문턱이었다: 11개 섹터의 일별 순위상관은 표준편차 ≈0.32로 매우 시끄러워 t≥2.0은
#    IC≥0.036을 요구하고, 실제 섹터 순환매 신호의 힘은 IC 0.02~0.03(잔차모멘텀 t 1.3~1.7)이라 전부 문턱에 걸려 0/9 연도 채택 →
#    전 기간 폴백. 리포트 수익으로 12-1 모멘텀 상위1×E_t를 억지로 돌리면(표본내·비용 제외) CAGR 16.8%(SPY×E_t 15.5%)이지만
#    MDD −12%(−10%)·샤프 1.31(1.61) — 약한 신호를 100% 집중하면 위험이 더 커질 수 있다는 것도 사실.
#    [⚠ 채택 모드] ROTATION_SELECT_MODE="best_available"(기본): 엄격(t≥2.0) 통과 신호가 있으면 그것, 없으면 방향 일치·
#      t≥ROTATION_SELECT_T_MIN(1.0)인 후보 중 t 상위 ROTATION_BEST_N(2)개를 그 해에 반드시 채택 → 순환매 판단이 항상 나온다.
#      13g '채택 근거' 열에 등급(엄격/외부검증 지원/최선 가용)을 기록하고 00시트 '순환매 예측 판정'에 등급 분포를 요약한다.
#      "strict"로 두면 v0.5 방식. 2개 채택 + 과반 규칙 = 둘이 같은 1위를 지목할 때만 100% 집중(약한 신호일 때 자연스러운 보호).
#    [외부 검증] fetch_ff49_daily()/parse_ff49_daily_csv()/ff49_signal_matrices()/external_validation_ff49(): Kenneth
#      French 49업종 일별 포트폴리오(무료, 1926~)에서 같은 정의의 후보(상대모멘텀 3·잔차모멘텀·상대과열·국면×베타 2)를 만들어
#      연도별 학습창(연초−35일 이전)으로 rank IC NW-t를 구한다(13h_외부검증FF49). 외부 t≥EXTERNAL_T(2.0)이고 11섹터 로컬
#      t≥1.0(방향 일치)이면 '외부검증 지원'으로 채택 — 좁은 횡단면(11)의 검정력 부족을 넓은 횡단면(49)으로 보완. 다운로드 실패
#      (샌드박스 403 등)면 생략·로그(리포트는 그대로 완성). 파서는 헤더 위치 변화에 강건(8자리 날짜 행만 읽음), 캐시 FF49_CACHE.
#    [새 후보 MACRO_BETA_FCST] rolling_ols_forecast(): 섹터 21일 상대(로그)수익을 팩터(DGS10·DX-Y.NYB·CL=F·BAMLH0A0HYM2·HG=F)
#      21일 변화에 3년 롤링 OLS(누적 외적합 차분으로 O(n)) → 예측 = Σ β̂·최근 21일 팩터 변화. 모든 입력이 t까지 확정(인과 단위
#      테스트). 사전부호 고정(MACRO_TAILWIND)보다 정밀한 매크로 순풍. ALL_COMPOSITE 구성원에 포함(기본 10종).
#    [검증] test_sector_rotation_v060.py 4항목 PASS(FF49 파서·외부 배관·채택 모드 3분기·롤링 OLS 정확성/인과성), 기존 스위트
#      전부 PASS(e2e 채택로그 108행). 실데이터 결과(어느 신호가 어느 등급으로 채택되고 리더가 누구인지)는 Colab 실행 후 13g/13h/00.
#  v0.5.1 | 2026-09-05 | 사용자 요청(v0.5.0 실데이터 리포트 확인 후): "섹터 순환매를 잘 예측하고 있는지 판단, 문제·개선점 확인".
#    [판정 — 실측 v0.5.0] 13g: 후보 9종 어느 것도 1999~2025 어느 학습창에서도 기준(rank IC NW-t ≥ 2.0)을 넘지 못함(0/9 연도
#      채택). 최강 RESID_MOM_12_1 t 1.25~1.68, MACRO_TAILWIND 0.2~0.85, BETA_X_MSCORE 대부분 음(−0.6~+0.2), REL_MOM_21은 9/9
#      연도 음(t≈−1.2, 섹터 수준 단기 반전). → 3단계 규칙은 전 기간 폴백(SPY) → 성과 = SPY M 전략(CAGR 18.54%/샤프 1.84/
#      MDD −7.1%; 대조군A 균등11 15.2%). 결론: 순환매를 '예측하지 않는다'는 것이 시스템의 판단이며 코드 결함이 아니라 데이터의
#      성질(11개 섹터 상대수익은 무료 일별 신호로 유의하게 예측되지 않음 — 2000년 이후 업종 모멘텀 약화 문헌과 일치).
#    [개선 1 — 검정력, 사전 고정] ROTATION_COMPOSITE_MEMBERS: MOM_COMPOSITE(잔차12-1·상대126·상대12-1 순위 평균),
#      ALL_COMPOSITE(기본 9종 사전방향 적용 순위 평균) 두 후보를 rotation_walkforward_select()에 추가 — 개별 신호가 약해도
#      오류가 독립이면 결합 순위의 IC는 올라간다(Asness·Moskowitz·Pedersen 2013). 구성은 사전 고정(데이터로 구성원을 고르지
#      않음), 같은 기준(t ≥ 2.0)으로 심사, 13d/13g에 표시. 단위테스트: 심은 신호를 포함한 ALL_COMPOSITE 채택·잡음만인
#      MOM_COMPOSITE 미채택.
#    [개선 2 — 리포트] 채택 신호가 한 해도 없으면 수용기준 ③④를 'N/A(해당 없음)'로, 종합을 '판정 불가(N/A)'로 표기(이전엔
#      '계산불가 FAIL'). ①②에는 '폴백 SPY vs 균등 11 비교 — 순위 신호 무관' 명시. 00시트 '순환매 예측 판정' 줄(예측 못함/
#      부분적/예측함 + 상위 후보 3종의 학습 t 최대·최소·평균) 신설 — rot_val["prediction_verdict"].
#    [하지 않은 것] 임계값 완화(t≥1.65), REL_MOM_21 부호 반전 — 데이터에 맞춘 사후 조정은 §3 원칙 위반. 다음 단계 후보:
#      Kenneth French 49업종 포트폴리오(무료)로 같은 신호를 외부 검증해 검정력을 높인 뒤 통과 신호만 11 섹터에 적용(미구현).
#  v0.5.0 | 2026-09-05 | 사용자 요청(v0.4.0 실데이터 리포트 확인 후): "문제 있으면 개선 — 지표 검증이 문제인지 추가 지표가
#    필요한지 원인 분석, 뉴스 참고, 국면 판단을 최대한 활용하는지 확인, 순환매를 잘 예측: 상승 확률 가장 높은 섹터 하나
#    매수, 하락 확률 가장 높은 건 피하고, 거의 차이가 없을 때의 전략도 생각해서 적용".
#    [원인 분석 — 실측 리포트 v0.4.0] 13d: 순위 신호 3종 전부 횡단면 rank IC≈0(NW-t 0.1~0.9), 집중배분(상위4) 13.5% <
#      대조군A(M노출×균등11) 15.1%, 2024~26 상위3−하위3 −10.9%/년. 근본 원인 = **지표 검증 방식**: 섹터 지표는 '그 섹터
#      절대수익'에 대해 시계열로만 검증되고(M 방식) 11섹터가 같은 매크로 지표를 공유해 횡단면 차이는 잡음 — "어느 섹터가
#      더 나은가"에 대한 검증이 없었음. 국면 판단 활용: 대조군A(E_t=M 목표비중)가 섹터 자체 타이밍 균등(11.2%)보다 훨씬
#      좋고 SPY M(18.5%)에 근접 → 성과의 원천은 M의 국면 판단이지만 '어느 섹터'에는 전혀 쓰이지 않았음. 리포트 데이터로
#      즉석 검사한 β_i×(M 국면)의 횡단면 IC는 NW-t≈2.0(5일)·1.5(21일)로 유일하게 의미 있는 후보. 역대 순환매 뉴스(2020
#      팬데믹 방어주→2021 리오프닝 에너지·금융→2022 금리충격 에너지↑기술↓→2023-24 AI 기술·통신)는 전부 "시장 국면×베타"와
#      "금리·유가·달러·신용의 방향"으로 설명되는 사례 → 두 축을 후보로 추가.
#    [순환매 후보 순위신호 9종 — ROTATION_SIGNAL_SPECS(사전방향·근거)] SCORE_PCT(마스킹 전 전체 이력), BETA_X_MSCORE
#      = β_i×(SPY 복합점수백분위−0.5)(신설: CAPM — 시장 기대수익 양이면 고베타, 음이면 저베타가 앞선다 = M 국면 판단의
#      횡단면 활용), BETA_X_MHAZ = β_i×(0.5−SPY H)(신설: 방어 로테이션), MACRO_TAILWIND(신설: 섹터 매크로표 §1.D의
#      사전부호×252일 z의 평균 — 매크로 방향이 그 섹터에 순풍인 정도), RESID_MOM_12_1, REL_MOM_126, REL_MOM_12_1,
#      REL_MOM_21(Moskowitz·Grinblatt 1개월 업종 모멘텀), REL_EXT_200(−, 상대 과열 평균회귀). run_sector()가
#      build_rotation_raw_signals()로 전체 이력(1999~)을 반환(rot_raw: 마스킹 없음, BETA 포함) + ret_cc_full.
#    [횡단면 워크포워드 검증 — rotation_walkforward_select()] 매년 1월 1일(2018~)에 '연초−35일' 이전 데이터만으로 각
#      후보의 일별 rank IC(지평 21일, 사전방향 적용 vs 향후 수익) 평균의 NW-HAC t를 구해 t ≥ ROTATION_SELECT_T(2.0) &
#      관측일 ≥ 1000인 신호만 그 해에 채택(조용한 채택 없음, 13g 시트에 연도별 전 후보 로그). 채택 신호의 0~1 순위 평균
#      (21일 평활) = 그 해 복합순위. 채택 0개인 해는 순위를 쓰지 않는다(→ 폴백). 인과성·판별력 단위테스트.
#    [3단계 배분 규칙 — ROTATION_TILT="leader3" ⚠] ① 채택 신호 과반이 같은 섹터를 1위로 지목(명확한 1위) → 그 섹터
#      하나에 E_t×ROTATION_LEADER_WEIGHT(1.0) ② 1위는 불명확하나 과반이 같은 섹터를 꼴찌로 지목 → 그 섹터를 뺀 적격
#      균등(≥3개) ③ 둘 다 아니면(차이 없음) ROTATION_FALLBACK="spy": SPY에 E_t(=검증된 M 전략; 균등 11 바스켓은 이
#      기간 SPY보다 3.4%p/년 뒤졌고 근거 없는 액티브 베팅). 리더 최소보유 ROTATION_MIN_HOLD_DAYS=21(월 리밸런스 관행),
#      적격 상실(하락 국면 등)은 즉시 청산, E_t=0이면 현금. '과반'은 채택 신호 수로 정해지는 규칙(임계값 격자 없음).
#      상위4·순위가중은 비교용으로 계속 산출(순위 없는 날은 같은 폴백). 포트폴리오에 'SPY' 열 추가(M bt의 ret_co/ret_oc
#      재사용) — 전 기간 폴백이면 M.run_backtest(SPY, E_t)와 비트 동일(단위테스트).
#    [리포트] 13d에 후보 종류·사전방향·설명·경제적 근거 열, 13g_순환매신호채택(연도별 학습 IC·t·채택) 신설, 13c에 판단·
#      1위·회피 섹터·채택신호수·득표·SPY 배분비중 열, 12_섹터요약에 리더 보유일·꼴찌 회피일 비율, 00시트 '다음 거래일 배분 -
#      판단'(왜 리더/회피/폴백/현금인지)·'판단 분포'·SPY M 성과 병기. 수용기준 판정문에 워크포워드 채택 현황 추가.
#    [합성 11섹터 검증] 섹터=β×SPY+잡음(M 자기검사 점수가 SPY를 예측하는 데이터)에서 워크포워드가 BETA_X_MSCORE·
#      BETA_X_MHAZ만 전 연도 채택(t 3.1·2.4), 나머지 7종 미채택 — 검증 배관이 구조를 찾아내고 잡음은 거른다. 리더 100%
#      집중은 CAGR +2.8%p(vs 대조군A)이지만 MDD 2.1%p 악화·OOS 복합 IC t=1.38로 수용기준 ②③ FAIL — 집중의 대가가
#      리포트에 그대로 드러남. 실데이터 판정은 Colab 실행 후 13f/13g 참조.
#  v0.4.0 | 2026-09-05 | 사용자 요청(v0.3.0 실데이터 리포트 확인 후): "11개 섹터 중 가장 상승 확률 높은 쪽에 비중을 많이
#    넣는 방식이지? 11개 균등하게 비중 넣으면 안 돼 — 그렇게 하도록 하고 문제 있으면 개선". 답: v0.3.0까지는 그런
#    방식이 아니었다 — 11개 섹터가 각자 자기 국면만 판단하고 13시트는 그 11개를 균등하게 섞은 참고 성과였다.
#    실측(2018-01-02~2026-09-04): 균등분산 CAGR 11.2%/샤프 1.28 vs SPY M 18.5%/1.83, 섹터 목표비중 간 평균 상관 0.64
#    (= '시장 타이밍 11벌', 순환매 판단 없음 — 개선계획 §0.4 그대로). XLK 24.2%가 XLE 4.7%(MDD -43%)·XLV 6.3%
#    (국면검증 FAIL)에 희석. → IMPROVEMENT_PLAN_SECTOR_v0.3.md §1.F(순환매 계층)를 구현.
#    [§1.F 집중 배분 계층 ⚠ 배분 파라미터 신설 — 섹터별 절대 국면 신호는 무변경]
#      build_sector_allocation(): 매일 11개 섹터를 횡단면 순위로 세운다. 복합순위 = ROTATION_SIGNALS(SCORE_PCT 복합점수
#      백분위 = 시스템의 '상승 확률' 척도, RESID_MOM_12_1 베타중립 잔차모멘텀, REL_MOM_126 상대모멘텀)의 그날 0~1 순위
#      평균(동일가중)을 ROTATION_SMOOTH_DAYS=21일 후행 평균으로 평활(일별 순위 뒤집힘의 회전율 억제 — 합성 11섹터에서
#      연 28회→16회; 적격 여부·E_t는 평활 없이 즉시 반영). 비중 = E_t × tilt(순위)/Σtilt — 주 전략 tilt는 상위 K=4 균등
#      (ROTATION_TILT="topk": 사용자 요청과 문헌의 상위 3분위 구성; 대안 "linear" 순위가중은 13시트에 나란히 산출),
#      E_t = SPY M 목표비중(검증된 시장 타이밍 — 총노출은 M이 정하고 배분 계층은 '어느 섹터'만 정한다),
#      ROTATION_EXCLUDE_STATES(RISK_OFF/TREND_ONLY_OUT/NO_SIGNAL) 섹터 제외, 섹터 상한 ROTATION_MAX_WEIGHT=25%
#      (워터필링 — v0.1 build_portfolio의 검증된 로직 이식; ROTATION_CAP_STRICT=True: 적격 섹터 <4개면 부족분은 현금,
#      상한 엄수 → 적격 3개면 75%·E_t, 2개면 50%·E_t만 투자). 값은 전부 사전 고정(격자 최적화 없음 — §3 프로토콜).
#      portfolio_backtest(): 다자산 T+1 시가 체결 — M.run_backtest 산식(전일종가→시가 pos_prev, 시가→종가 pos_exec,
#      현금레그 rf, 편도 비용 5bp)을 섹터별로 확장. 단일 섹터·target=E_t 퇴화 케이스에서 M.run_backtest와 비트 동일(단위테스트).
#      대조군: A "M노출×균등11(제외·캡 없음)" — 순환매 없이 M 타이밍만 / B "M노출×균등(하락제외·캡)" — 주 전략과 같은
#      집합, tilt만 균등(순위의 순수 효과 분리). 13시트 = 배분 4개 + 기존 참고 4개(균등분산·균등B&H·SPY M·SPY B&H), 같은 평가창.
#      rotation_validation(): 횡단면 rank IC(신호별·복합, 지평 5/21/63, 뉴이-웨스트 HAC 평균-t — v0.1
#      cross_sectional_rank_ic 이식·벡터화) + 상위3−하위3 일별 스프레드(3구간) + 사전 고정 수용기준(§1.F.3: 대조군A 대비
#      CAGR +1.5%p·MDD 악화 ≤1%p·복합 IC NW-t ≥2.0·스프레드 3구간 중 2 양수) PASS/FAIL. 미래수익은 이 진단에서만 쓰이고
#      배분비중 계산에는 되돌아가지 않는다(인과성 단위테스트: d 이후 교란 → d 이하 비중·수익 불변).
#      ⚠ 수용기준 FAIL이어도 사용자 요청에 따라 집중배분을 유지·표시한다(ROTATION 판정을 00시트에 명시, 대조군 성과 병기).
#    [리포트] 13_섹터배분전략(13_섹터분산전략(참고) 개명) / 13b_배분전략자산곡선(개명, 배분 4곡선 추가) / 13c_일별배분비중
#      (날짜×섹터 목표비중·순위·E_t·합계·현금·적격수, 맨 끝 '예측' 행 = 다음 거래일 체결 비중) / 13d_횡단면IC / 13e_순위스프레드
#      / 13f_배분수용기준 신설. 00시트 '다음 거래일 배분 - *'(방식·요약·섹터 11줄·합계/현금·수용기준·대조군 비교) 블록을
#      '다음 거래일 예측' 아래에 추가. 01Z에 '집중배분 합계'·'집중배분(다음날 체결, 상위5)' 열, 12_섹터요약에 '배분 평균비중'·
#      '배분 1위 빈도'·'배분 적격일 비율' 열. sector_allocation_daily.csv(ALLOC_CSV_PATH) 추가 저장·다운로드.
#    [반환] run_sector(): ret_co/ret_oc(백테스트 수익 분해)·rot_raw(순위 신호 원시값) 추가. run(): alloc/rot_val/alloc_sheet.
#    [검증] test_sector_allocation_v040.py(5항목 PASS), 기존 스위트 재실행 PASS(e2e는 13~13f 시트·Σw≤E_t·상한·제외 검사 추가).
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

VERSION = "v0.7.0"
VERSION_DATE = "2026-09-06"

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
    # ---- 섹터 집중 배분 계층 [v0.4.0 §1.F ⚠ 배분(사이징) 파라미터] ---------------
    # 사용자 요청(2026-09-05 실데이터 리포트 확인 후): "11개 균등 비중은 안 됨 — 상승 확률이 가장 높은 섹터에
    # 비중을 많이". 매일 11개 섹터를 횡단면 순위(ROTATION_SIGNALS의 그날 순위 평균)로 세워 순위가 높을수록 큰
    # 비중을 주고(순위가중), 하락 국면 섹터는 제외, 총 노출 E_t는 검증된 SPY M의 목표비중을 따른다(§1.F.2).
    # 값은 전부 사전 고정(격자 최적화 없음) — 수용기준(§1.F.3)으로 순위 신호의 정보 유무를 같은 리포트에서 판정.
    USE_ROTATION: bool = True
    # [v0.5.0] 순환매 후보 순위신호 — 전부 ROTATION_SIGNAL_SPECS(사전방향·근거)에 등록된 이름. v0.4.0 실측에서 v0.4의 3종은
    #   횡단면 rank IC≈0(NW-t 0.1~0.9)이었다: 섹터 지표가 '그 섹터 절대수익'에 대해 시계열로만 검증됐고 11섹터가 같은 매크로
    #   지표를 공유해 횡단면 차이는 잡음이었기 때문. 그래서 v0.5.0은 (1) 후보를 넓히고(국면×베타, 매크로 순풍 — 역대 순환매
    #   뉴스(2020 팬데믹 방어주→2021 리오프닝 에너지·금융→2022 금리충격→2023-24 AI 기술)가 전부 이 두 축으로 설명됨)
    #   (2) 횡단면 워크포워드 검증을 통과한 신호만 그 해에 쓴다(ROTATION_SELECT_*).
    ROTATION_SIGNALS: Tuple[str, ...] = ("SCORE_PCT", "BETA_X_MSCORE", "BETA_X_MHAZ", "MACRO_TAILWIND",
                                         "RESID_MOM_12_1", "REL_MOM_126", "REL_MOM_12_1", "REL_MOM_21", "REL_EXT_200",
                                         "MACRO_BETA_FCST",                   # [v0.6.0] 롤링 회귀 민감도 × 최근 매크로 변화
                                         "MOM_COMPOSITE", "ALL_COMPOSITE")   # [v0.5.1] 사전 고정 복합 후보 2종
    MACRO_BETA_FACTORS: Tuple[str, ...] = ("DGS10", "DX-Y.NYB", "CL=F", "BAMLH0A0HYM2", "HG=F")   # 금리·달러·유가·신용·구리
    MACRO_BETA_WINDOW: int = 756               # 민감도 추정 롤링 창(거래일, 약 3년)
    MACRO_BETA_MIN_COVERAGE: float = 0.8       # [v0.7.0] 팩터 채택 최소 커버리지(섹터 이력 대비 관측 비율) — v0.6 실측 '학습 36일' 결함 수정
    ROTATION_SELECT_T: float = 2.0             # '엄격' 채택 기준: 학습창 선택통계(아래 STAT)의 NW-HAC t ≥ 2.0 & 부호 = 사전방향
    ROTATION_SELECT_MIN_DAYS: int = 1000       # 학습창 최소 관측일(약 4년 — 그 미만이면 그 신호는 아직 판정 불가=미채택)
    ROTATION_SELECT_HORIZON: int = 21          # 검증 지평(거래일) — 월 리밸런스 관행과 동일
    # [v0.7.0 ⚠ 선택 통계] 사용자 목표("국면 비중을 따르며 그중 상승 확률이 가장 높은 섹터 하나를 사서 SPY 국면전략보다 높은 수익").
    #   v0.6 실측 진단: 평균 rank IC는 '11개 전체 순서'를 재는 잣대여서 '1위 하나를 고르는' 전략과 맞지 않았다 — 시스템의 상승 확률
    #   SCORE_PCT는 rank IC t≈0.5로 탈락했지만 그 1위 섹터의 21일 초과수익(상위1 스프레드)은 +1.25%/21일·NW-t 3.2였고, 채택된
    #   잔차모멘텀은 IC는 있어도 1위가 저베타 방어섹터로 몰려 SPY보다 뒤졌다(리더 16.6% < SPY M 18.7%).
    #   "top1": 선택통계 = 상위1 스프레드(그날 평활 순위 1위 섹터의 향후 h일 수익 − 상장 섹터 평균)의 NW-HAC t — 전략이 실제로 얻는
    #   것을 직접 검증. "ic": v0.5~0.6 방식(평균 rank IC의 t). 두 통계 모두 13g에 기록되며 임계값(t≥2.0/1.0)은 그대로.
    ROTATION_SELECT_STAT: str = "top1"         # "top1"(기본) | "ic"
    # [v0.6.0 ⚠] 채택 모드 — 사용자 요청("SPY가 오른다고 모든 섹터가 오르는 게 아니다 — 순환매 판단을 하도록 하라").
    #   v0.5 실측: 후보 전부 t 1.3~1.7로 '엄격' 문턱(2.0)에 걸려 0/9 연도 채택 → 전 기간 폴백(SPY). 11개 섹터의 일별 순위상관은
    #   표준편차 ≈0.32로 매우 시끄러워 t≥2.0은 IC≥0.036을 요구하는데, 문헌상 섹터 순환매 신호의 힘은 IC 0.02~0.03 수준이다.
    #   "best_available": 엄격 통과 신호가 있으면 그것을, 없으면 방향이 맞고 t ≥ ROTATION_SELECT_T_MIN인 후보 중 t 상위
    #   ROTATION_BEST_N개를 그 해에 반드시 채택 → 순환매 판단이 항상 나오되, 13g '채택 근거'에 등급(엄격/외부검증 지원/최선 가용)을
    #   같이 기록한다. "strict": v0.5 방식(엄격 + 외부검증 지원만). 약한 신호로 한 섹터 100% 집중은 위험이 커질 수 있음 —
    #   13시트에서 대조군·SPY M과 나란히 확인할 것.
    ROTATION_SELECT_MODE: str = "best_available"   # "best_available" | "strict"
    ROTATION_BEST_N: int = 2                   # best_available에서 강제 채택하는 최대 후보 수(과반 규칙과 맞물림: 2개면 둘이 같은 1위여야 리더)
    ROTATION_SELECT_T_MIN: float = 1.0         # best_available/외부검증 지원 채택의 최소 로컬 t(방향 일치 필수)
    # [v0.6.0] 외부 검증 — Kenneth French 49업종 일별 포트폴리오(무료)로 같은 신호의 rank IC를 더 넓은 횡단면에서 검증.
    #   외부 t ≥ EXTERNAL_T이고 11섹터 로컬 t ≥ ROTATION_SELECT_T_MIN(방향 일치)이면 '외부검증 지원'으로 채택.
    #   다운로드 실패(네트워크 없음)면 생략하고 로그만 남긴다(리포트는 그대로 완성).
    USE_EXTERNAL_VALIDATION: bool = True
    EXTERNAL_T: float = 2.0
    FF49_URL: str = "https://mba.tuck.dartmouth.edu/pages/faculty/ken.french/ftp/49_Industry_Portfolios_daily_CSV.zip"
    FF49_CACHE: str = "./cache_sector/ff49_daily.csv"
    ROTATION_SMOOTH_DAYS: int = 21             # 복합순위점수의 후행 이동평균 창(거래일, 인과). 1=평활 없음. 근거: 일별 순위 뒤집힘이
    #   그대로 비중 변화가 되면 회전율이 연 28회(합성 11섹터 실측)로 치솟아 비용이 성과를 잠식 — 문헌·§0.6의 검증 단위인
    #   '월 1회 리밸런스'에 맞춘 1개월(21일) 평활. 적격 여부(하락 제외)와 E_t 변화는 평활 없이 당일 즉시 반영된다.
    # [v0.5.0] 주 전략 "leader3" — 사용자 요청 그대로의 3단계 규칙:
    #   ① 그 해 채택된 신호의 '과반'이 같은 섹터를 1위로 지목(명확한 1위) → 그 섹터 하나에 E_t × ROTATION_LEADER_WEIGHT
    #   ② 1위는 불명확하지만 과반이 같은 섹터를 꼴찌로 지목(명확한 꼴찌) → 그 섹터를 뺀 적격 섹터 균등(≥3개일 때)
    #   ③ 둘 다 아니면(거의 차이가 없음) → ROTATION_FALLBACK: "spy"(검증된 M 전략 그대로 SPY에 E_t) | "equal"(적격 균등)
    #   '과반'은 임계값이 아니라 채택 신호 수로 정해지는 규칙(임계값 격자 없음). 채택 신호가 0개인 해는 항상 ③.
    ROTATION_TILT: str = "leader3"             # "leader3"(기본) | "topk"(상위 K 균등) | "linear"(순위가중). 나머지 둘은 13시트 비교용으로 항상 산출
    ROTATION_LEADER_WEIGHT: float = 1.0        # ⚠ 명확한 1위 섹터에 주는 E_t 대비 비중(1.0 = 그 섹터 하나만 매수 — 사용자 요청)
    ROTATION_FALLBACK: str = "spy"             # "spy" | "equal". 기본 spy: 균등 11섹터 바스켓은 이 기간 SPY보다 3.4%p/년 뒤졌고(대조군A
    #   15.1% vs SPY M 18.5%) 근거 없는 액티브 베팅(동일가중 vs 시총가중)이므로, 차이가 없을 때는 검증된 M 전략을 그대로 든다
    ROTATION_MIN_HOLD_DAYS: int = 21           # 1위 섹터 최소 보유(거래일) — 월 리밸런스 관행. 적격 상실(하락 국면 등)은 즉시 청산
    ROTATION_TOP_K: int = 4                    # "topk" 모드의 K(11개의 약 1/3 — 상위 3분위). 상한 25%와 함께 K×25%=100%
    ROTATION_MAX_WEIGHT: float = 0.25          # 섹터 상한(E_t 대비 비율) — 상한 초과분은 워터필링으로 다른 적격 섹터에 재배분
    ROTATION_CAP_STRICT: bool = True           # True: 적격 섹터가 적어(<1/상한) 예산을 다 못 채우면 남는 노출은 현금(상한 엄수)
    ROTATION_EXCLUDE_STATES: Tuple[str, ...] = ("RISK_OFF", "TREND_ONLY_OUT", "NO_SIGNAL")  # 배분 제외 국면
    ROTATION_IC_HORIZON: int = 21              # 횡단면 rank IC 검정 지평(거래일) — 수용기준 판정용(다른 지평은 참고 표시)
    # 수용기준(§1.F.3, 사전 고정) — 대조군 "M노출×균등11" 대비
    ROTATION_ACCEPT_CAGR_GAIN: float = 0.015   # CAGR +1.5%p 이상
    ROTATION_ACCEPT_MDD_WORSE: float = 0.01    # MDD 악화 1%p 이하
    ROTATION_ACCEPT_IC_T: float = 2.0          # ③ 복합순위의 검정 t ≥ 2.0 — [v0.7.0] ROTATION_SELECT_STAT="top1"이면 상위1 스프레드 t, "ic"면 rank IC t
    ROTATION_ACCEPT_SPREAD_PERIODS: int = 2    # 상위3−하위3 스프레드가 3구간(18-20/21-23/24-26) 중 양수인 구간 수 ≥ 2
    # [v0.7.0] ⑤ 목표 기준 — 사용자 목표 그대로: 주 전략 CAGR ≥ SPY 국면전략(M) CAGR + 아래 여유(0 = 같기만 해도 PASS).
    #   ①~④가 '대조군A(균등 11) 대비 정보 유무'라면 ⑤는 '이 계층을 둘 이유가 있는가'(M만 쓰는 것보다 나은가)다.
    ROTATION_ACCEPT_VS_SPY_M: float = 0.0
    ROTATION_ALT_LEADER_OWN_POS: bool = True   # [v0.7.0] 비교 변형 '리더 자체 목표비중'(리더 섹터의 M식 target_pos를 노출로 사용)을 13시트에 산출
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
    ALLOC_CSV_PATH: str = "sector_allocation_daily.csv"   # [v0.4.0] 13c 일별 배분비중 CSV(EXPORT_DAILY_CSV와 함께 저장)
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


# ---- [v0.5.0] 순환매 후보 순위신호(횡단면) — 사전방향·근거 표. 채택은 rotation_walkforward_select()가 매년 과거 데이터로만 판정.
#   (prior_sign, name_kr, rationale)  — 사전방향은 "값이 클수록 다른 섹터보다 앞선다(+)/뒤진다(−)"
ROTATION_SIGNAL_SPECS: Dict[str, Tuple[int, str, str]] = {
    "SCORE_PCT": (+1, "섹터 복합점수 백분위(시계열 검증 결과)",
                  "그 섹터 자체 상승 확신이 높을수록 다른 섹터보다 앞설 것이라는 가설 — v0.4.0 실측 IC≈0(시계열 검증이 횡단면 차이를 보장하지 않음)"),
    "BETA_X_MSCORE": (+1, "국면×베타: β_i × (SPY 복합점수백분위 − 0.5)",
                      "CAPM: 시장 기대수익이 양이면 고베타 섹터가, 음이면 저베타(방어) 섹터가 앞선다 — M의 검증된 국면 판단을 '어느 섹터'에 직접 활용"),
    "BETA_X_MHAZ": (+1, "국면×베타(위험): β_i × (0.5 − SPY 위험점수백분위 H)",
                    "시장 위험(H)이 높을수록 저베타·방어 섹터(필수소비·유틸·헬스)로 자금 이동 — 방어 로테이션"),
    "MACRO_TAILWIND": (+1, "매크로 순풍: 섹터 매크로표(§1.D) 사전부호 × z(60일 변화)의 평균",
                       "금리·달러·유가·구리·신용·커브의 최근 방향이 그 섹터의 경제적 사전방향과 맞을수록 앞선다 — 역대 순환매 뉴스"
                       "(2021 리오프닝 에너지·금융, 2022 금리충격 에너지↑기술↓)의 공통 설명변수"),
    "RESID_MOM_12_1": (+1, "베타중립 잔차모멘텀 12-1개월", "Blitz·Huij·Martens 2011 — 시장 성분을 제거한 순수 섹터 모멘텀"),
    "REL_MOM_126": (+1, "SPY대비 6개월 상대모멘텀", "중기 상대강도 지속(섹터 모멘텀)"),
    "REL_MOM_12_1": (+1, "SPY대비 12-1개월 상대모멘텀", "전통적 모멘텀 팩터 정의(최근 1개월 제외)"),
    "REL_MOM_21": (+1, "SPY대비 1개월 상대모멘텀", "Moskowitz·Grinblatt 1999 — 업종 모멘텀은 1개월 지평에서 가장 강함(개별주와 달리 단기 반전 없음)"),
    # [v0.7.0 ⚠ 사전방향 변경 −1 → +1] v0.5의 '상대 과열 평균회귀' 가설은 근거가 약한 자체 가설이었고, 외부 데이터(Ken French 49업종
    #   1926~, 13h)에서 같은 정의의 신호가 전 학습창에서 강한 양(+)의 지속(t≈+8)을 보였다 — 200일선 대비 상대 이격은 사실상 '상대강도'
    #   이며 업종 수준에서는 지속(Moskowitz·Grinblatt 1999)이 문헌의 기본 결과다. 11섹터 표본이 아닌 외부 표본·문헌에 근거한 변경
    #   (사후 데이터 맞춤과 구별). 채택 여부는 여전히 워크포워드가 결정.
    "REL_EXT_200": (+1, "섹터−SPY 200일선 이격도 차(상대 추세 강도)",
                    "업종 수준 상대강도는 지속 — 외부(FF49) 검증에서 양(+) 지속 확인, 평균회귀 가설(v0.5)은 기각"),
    # [v0.6.0] 학습된 매크로 민감도 × 최근 변화 — 사전부호 고정(MACRO_TAILWIND)보다 정밀. 민감도는 t-21까지의 완결 창으로만 추정(인과).
    "MACRO_BETA_FCST": (+1, "매크로 베타 예측: Σ_f β̂_{i,f}(3년 롤링) × 최근 21일 팩터 변화(금리·달러·유가·신용·구리)",
                        "섹터별 매크로 민감도가 다르므로 같은 매크로 변화가 섹터마다 다른 상대수익을 만든다 — 2021 리오프닝(유가·금리↑→에너지·"
                        "금융), 2022 금리충격(→기술 열위)의 채널을 데이터로 추정한 민감도로 잡음"),
    # [v0.5.1] 사전 고정 복합 후보 — 개별 신호가 약해도 오류가 독립이면 결합 순위의 IC는 올라간다(Asness·Moskowitz·Pedersen
    #   2013의 순위 결합). 구성은 사전에 고정(모멘텀 가족 3종 / 기본 후보 9종 전체, 동일가중) — 데이터로 구성원을 고르지 않는다.
    #   같은 워크포워드 기준(t ≥ 2.0)으로 심사되며, 채택되면 다른 채택 신호와 함께 순위 평균에 들어간다.
    "MOM_COMPOSITE": (+1, "모멘텀 복합(잔차12-1 · 상대126 · 상대12-1 순위 평균)",
                      "세 모멘텀 정의의 공통 성분만 남겨 잡음을 상쇄 — v0.5.0 실측에서 셋 다 양(+)이지만 개별로는 t 0.5~1.7"),
    "ALL_COMPOSITE": (+1, "전체 복합(기본 후보 10종의 사전방향 적용 순위 평균)",
                      "서로 다른 경제 채널(국면×베타·매크로·모멘텀·과열)의 결합 — 사전 고정 동일가중"),
}
ROTATION_COMPOSITE_MEMBERS: Dict[str, Tuple[str, ...]] = {
    "MOM_COMPOSITE": ("RESID_MOM_12_1", "REL_MOM_126", "REL_MOM_12_1"),
    "ALL_COMPOSITE": ("SCORE_PCT", "BETA_X_MSCORE", "BETA_X_MHAZ", "MACRO_TAILWIND", "RESID_MOM_12_1",
                      "REL_MOM_126", "REL_MOM_12_1", "REL_MOM_21", "REL_EXT_200", "MACRO_BETA_FCST"),
}


def _macro_factor_changes(res: dict, idx: pd.DatetimeIndex, factors: Tuple[str, ...], h: int = 21) -> pd.DataFrame:
    """[v0.6.0] 팩터별 h일 변화(금리·스프레드는 레벨 차이 %p, 가격은 로그수익). res["fred"]/res["px_dict"]에서만 읽는다(재수집 없음)."""
    out = pd.DataFrame(index=idx)
    for f in factors:
        if f in (res.get("fred") or {}) and res["fred"][f] is not None:
            s = res["fred"][f].reindex(idx).ffill()
            out[f] = s - s.shift(h)
        elif f in (res.get("px_dict") or {}) and res["px_dict"][f] is not None:
            d = res["px_dict"][f]
            d = d[~d.index.duplicated(keep="last")].sort_index()
            col = "Adj Close" if "Adj Close" in d.columns else "Close"
            s = d[col].astype(float).reindex(idx).ffill()
            out[f] = np.log(s.replace(0, np.nan)).diff(h)
    return out.replace([np.inf, -np.inf], np.nan)


def rolling_ols_forecast(y: pd.Series, X: pd.DataFrame, window: int = 756, min_obs: int = 504) -> pd.Series:
    """[v0.6.0] 롤링 OLS(절편 포함) y_s ~ X_s (s ≤ t, 창 window) → 예측 = Σ_f β̂_{f,t}·X_{f,t}(절편 제외). 모든 입력이 t까지의
    확정치라 인과적. 관측 부족(min_obs)·특이행렬이면 NaN. 누적 외적합의 차분으로 X'X, X'y를 O(n)로 구하고 날짜별 (k+1)×(k+1)
    선형계를 푼다(릿지 1e-10은 수치 안정용)."""
    df = pd.concat([y.rename("_y"), X], axis=1)
    valid = df.notna().all(axis=1).values
    Z = df.where(pd.Series(valid, index=df.index), 0.0)
    Xa = np.column_stack([np.ones(len(Z)), Z[list(X.columns)].values.astype(float)])
    ya = Z["_y"].values.astype(float)
    k1 = Xa.shape[1]
    XX = np.einsum("ni,nj->nij", Xa, Xa) * valid[:, None, None]
    Xy = Xa * ya[:, None] * valid[:, None]
    cXX = np.cumsum(XX, axis=0); cXy = np.cumsum(Xy, axis=0); cN = np.cumsum(valid.astype(int))
    out = np.full(len(Z), np.nan)
    eye = 1e-10 * np.eye(k1)
    for t in range(len(Z)):
        if not valid[t]:
            continue
        lo = t - window
        sXX = cXX[t] - (cXX[lo] if lo >= 0 else 0.0)
        sXy = cXy[t] - (cXy[lo] if lo >= 0 else 0.0)
        n = cN[t] - (cN[lo] if lo >= 0 else 0)
        if n < min_obs:
            continue
        try:
            beta = np.linalg.solve(sXX + eye, sXy)
        except np.linalg.LinAlgError:
            continue
        out[t] = float(np.dot(beta[1:], Xa[t, 1:]))
    return pd.Series(out, index=y.index)


def build_rotation_raw_signals(ticker: str, ind_i: pd.DataFrame, score: pd.Series, adj_i: pd.Series,
                               spy_tr: pd.Series, idx_i: pd.DatetimeIndex, M,
                               res: Optional[dict] = None, scfg: Optional["SectorConfig"] = None) -> pd.DataFrame:
    """[v0.5.0] 순환매 후보 순위신호의 '원시값'(전체 이력, 인과) — 국면×베타 두 신호는 SPY 계층 시리즈가 필요해
    build_sector_allocation()에서 β와 결합한다(여기서는 BETA만 계산). 없는 열은 NaN(워크포워드에서 자동 미채택).
    [v0.6.0] MACRO_BETA_FCST: 섹터 21일 상대(로그)수익을 팩터 21일 변화에 3년 롤링 회귀한 민감도 × 최근 21일 팩터 변화."""
    out = pd.DataFrame(index=idx_i)
    out["SCORE_PCT"] = M.score_percentile(score).reindex(idx_i)          # 마스킹 전(SIGNAL_START 이전에도 값)
    for sfx in ("RESID_MOM_12_1", "REL_MOM_126", "REL_MOM_12_1", "REL_MOM_21", "REL_EXT_200"):
        key = f"{ticker}__{sfx}"
        out[sfx] = ind_i[key].reindex(idx_i) if key in ind_i.columns else np.nan
    r_i = np.log(adj_i.replace(0, np.nan)).diff()
    r_spy = np.log(spy_tr.reindex(idx_i).replace(0, np.nan)).diff()
    out["BETA"] = rolling_beta(r_i, r_spy, window=252, lag=1)
    # 매크로 순풍: 섹터 매크로표 각 행의 값(이미 발표지연·변화율 적용)을 252일 롤링 z로 표준화해 사전부호를 곱한 평균
    parts = []
    for suffix, name_kr, kind, sid, transform, window, prior_sign, rationale, mech in SECTOR_MACRO_TABLE.get(ticker, []):
        key = f"{ticker}__{suffix}"
        if key not in ind_i.columns:
            continue
        s = ind_i[key].reindex(idx_i)
        mu = s.rolling(252, min_periods=120).mean()
        sd = s.rolling(252, min_periods=120).std().replace(0, np.nan)
        parts.append(((s - mu) / sd).clip(-3, 3) * prior_sign)
    out["MACRO_TAILWIND"] = pd.concat(parts, axis=1).mean(axis=1, skipna=True) if parts else np.nan
    # [v0.6.0] 매크로 베타 예측 — 팩터 데이터가 res에 있을 때만(없으면 NaN → 워크포워드 미채택)
    out["MACRO_BETA_FCST"] = np.nan
    if res is not None and scfg is not None and getattr(scfg, "MACRO_BETA_FACTORS", None):
        X = _macro_factor_changes(res, idx_i, tuple(scfg.MACRO_BETA_FACTORS), h=21)
        # [v0.7.0 결함 수정] v0.6 실측에서 학습 관측일이 36일뿐이었다(→ '표본부족'): 팩터를 '관측 500일 초과'로만 걸러 이력이 짧은
        #   팩터(예: 늦게 시작하는 시리즈)가 들어오면 rolling OLS의 '모든 열 유효' 조건이 그 팩터의 시작 이후로만 성립하기 때문.
        #   → 섹터 이력(idx_i) 대비 커버리지 ≥ MACRO_BETA_MIN_COVERAGE인 팩터만 쓰고, 사용/제외 팩터를 로그에 남긴다.
        min_cov = float(getattr(scfg, "MACRO_BETA_MIN_COVERAGE", 0.8))
        cov = X.notna().mean(axis=0) if X.shape[1] else pd.Series(dtype=float)
        keep = [c for c in X.columns if cov[c] >= min_cov]
        drop = {c: round(float(cov[c]), 3) for c in X.columns if c not in keep}
        X = X[keep]
        if X.shape[1] >= 2:
            y_rel = (r_i - r_spy).rolling(21).sum()          # 21일 상대 로그수익(창 [t-20, t], t까지 확정)
            out["MACRO_BETA_FCST"] = rolling_ols_forecast(y_rel, X, window=int(scfg.MACRO_BETA_WINDOW), min_obs=504)
            log("ROTATION", kv(event="macro_beta_factors", ticker=ticker, used=",".join(keep), dropped_low_coverage=drop or "-",
                               valid_days=int(out["MACRO_BETA_FCST"].notna().sum())), M=M)
        else:
            log("ROTATION", kv(event="macro_beta_skipped", ticker=ticker, reason=f"커버리지≥{min_cov:.0%} 팩터 {len(keep)}개(<2)",
                               coverage={c: round(float(v), 3) for c, v in cov.items()}), M=M, level="warning")
    return out.replace([np.inf, -np.inf], np.nan)


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
    # [v0.5.0] 순환매 후보 순위신호 — '전체 이력'(SIGNAL_START 마스킹 없음: 횡단면 워크포워드 검증이 2018 이전 이력으로
    # 학습해야 하므로). 전부 그 섹터 후보지표 프레임/점수에서 그대로 꺼내거나 인과적으로 계산(재계산·미래 정보 없음).
    rot_raw = build_rotation_raw_signals(ticker, ind_i, score, adj_i, spy_tr, idx_i, M, res=res, scfg=scfg)
    ret_cc_full = adj_i.pct_change()   # 전체 이력 총수익(검증용 미래수익 산출 — 배분에는 미사용)
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
            "ma_ret": bt_ma["strategy_ret"],
            # [v0.4.0 §1.F] 포트폴리오 백테스트용 수익 분해(M.run_backtest와 동일 정의: 전일종가→시가, 시가→종가) + 순위 신호
            "ret_co": bt["ret_co"], "ret_oc": bt["ret_oc"], "rot_raw": rot_raw,
            "ret_cc_full": ret_cc_full}   # [v0.5.0] 전체 이력(횡단면 워크포워드 검증의 미래수익 산출용)


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
    portfolio_perf, portfolio_curve = build_portfolio_reference(results, res, eval_idx, M)

    # [v0.4.0 §1.F] 섹터 집중 배분 계층 — 섹터별 절대 국면 위에 횡단면 순위로 비중을 기울인다(E_t = SPY M 목표비중).
    alloc: Dict[str, Any] = {}
    rot_val: Dict[str, Any] = {}
    alloc_sheet = pd.DataFrame()
    if scfg.USE_ROTATION and results:
        t1 = time.time()
        try:
            alloc = build_sector_allocation(results, res, eval_idx, scfg, M, rf_daily=rf_daily, spy_series=spy_series)
            rot_val = rotation_validation(alloc, scfg, M)
            alloc_sheet = build_allocation_sheet(alloc, nd_spy, scfg)
            # 13 성과표/13b 자산곡선: 배분 전략 4개를 앞에, 기존 참고 4개(균등분산·균등B&H·SPY M·SPY B&H)를 뒤에
            portfolio_perf = pd.concat([alloc["perf"], portfolio_perf], ignore_index=True, sort=False)
            portfolio_curve = alloc["curve"].merge(portfolio_curve, on="날짜", how="left")
        except Exception as e:
            log("ROTATION", kv(event="allocation_failed", err=type(e).__name__, msg=str(e)[:200],
                               action="배분 계층 없이 리포트 계속(13 시트는 기존 참고 4개만)"), M=M, level="error")
            alloc, rot_val, alloc_sheet = {}, {}, pd.DataFrame()
        stage_timing["03a_집중배분"] = round(time.time() - t1, 2)
    matrix = build_prediction_matrix(results, res, eval_idx, scfg, nd_spy=nd_spy, alloc=alloc)
    summary = build_sector_summary(results, failed, universe, alloc=alloc)
    stage_timing["03_통합시트"] = round(time.time() - t0, 2)
    stage_timing["04_run()합계"] = round(time.time() - t_all, 2)
    log("DONE", kv(event="sector_pipeline_complete", ok=len(results), failed=len(failed),
                   rotation=("on" if alloc else "off"), elapsed_s=stage_timing["04_run()합계"]), M=M)
    return {"sectors": results, "failed": failed, "selftest": st, "universe": universe, "quality": pd.DataFrame(quality),
            "matrix": matrix, "portfolio_perf": portfolio_perf, "portfolio_curve": portfolio_curve,
            "summary": summary, "stage_timing": stage_timing, "scfg": scfg, "M_cfg": M_cfg,
            "signal_start": str(sig_start.date()), "cal_end": str(cal[-1].date()), "aborted": False,
            "m_bundle_meta": res.get("bundle_meta", {}), "nd_spy": nd_spy,
            # [v0.4.0 §1.F]
            "alloc": alloc, "rot_val": rot_val, "alloc_sheet": alloc_sheet}


# =============================================================================
# [10] 통합 시트 — 01Z 섹터일별예측 매트릭스 / 12 섹터요약 / 13 섹터분산전략(참고)
# =============================================================================
def build_prediction_matrix(results: Dict[str, Dict[str, Any]], res: dict, eval_idx: pd.DatetimeIndex,
                            scfg: SectorConfig, nd_spy: Optional[dict] = None,
                            alloc: Optional[Dict[str, Any]] = None) -> pd.DataFrame:
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
    # [v0.4.0 §1.F] 집중 배분 요약(그날 종가 확정 → 다음 거래일 시가 체결 비중) — 상세는 13c_일별배분비중
    if alloc:
        out.insert(6, "집중배분 합계", alloc["target_w"].sum(axis=1).reindex(eval_idx).round(4).values)
        out.insert(7, "집중배분(다음날 체결, 상위5)", [allocation_summary_text(alloc, d) for d in eval_idx])
    out = out.reset_index(drop=True)

    if nd_spy is not None:
        row = {c: ("" if out[c].dtype == object else np.nan) for c in out.columns}
        row["날짜"] = nd_spy["다음거래일"].date()
        row["구분"] = "예측"
        row["SPY 시장상황"] = STATE_SHORT.get(nd_spy["확정국면_원시"], nd_spy["확정국면_원시"])
        row["SPY 목표비중"] = nd_spy["목표비중"]
        if alloc and len(eval_idx):
            # 마지막 실적 행의 배분비중이 곧 다음 거래일에 체결할 비중(새 계산 없음)
            row["집중배분 합계"] = round(float(alloc["target_w"].iloc[-1].sum()), 4)
            row["집중배분(다음날 체결, 상위5)"] = allocation_summary_text(alloc, eval_idx[-1])
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


# =============================================================================
# [10b] [v0.4.0 §1.F] 섹터 집중 배분 계층 — "어느 섹터가 다른 섹터보다 나은가"의 횡단면 판단
#   섹터별 절대 국면(위 [6]~[10])은 그대로 두고, 그 위에 매일 11개 섹터의 횡단면 순위로 비중을 기울인다.
#   w_i,t = E_t × tilt(rank_i,t) / Σ tilt,  E_t = SPY M 목표비중(검증된 시장 타이밍),  하락 국면 섹터 제외,
#   섹터 상한(워터필링).  t일 종가 확정 → t+1일 시가 체결, 비용·현금레그는 M.run_backtest와 같은 산식.
#   순위 신호의 '정보 유무'는 rotation_validation()이 사전 고정 수용기준(§1.F.3)으로 같은 리포트에서 판정.
# =============================================================================
ROT_LABEL_PRIMARY = "집중배분(순위가중·캡{cap:.0%}) ★"
ROT_LABEL_TOPK = "집중배분(상위{k} 균등·캡{cap:.0%})"
ROT_LABEL_CTRL_A = "대조군A: M노출×균등11(제외·캡 없음)"
ROT_LABEL_CTRL_B = "대조군B: M노출×균등(하락제외·캡)"


def _cs_rank01(mat: pd.DataFrame) -> pd.DataFrame:
    """행(날짜)별 횡단면 순위를 0(최저)~1(최고)로 정규화. 그날 값이 1개뿐이면 0.5, NaN은 순위에서 제외."""
    rk = mat.rank(axis=1, method="average")
    n = mat.notna().sum(axis=1)
    out = (rk - 1.0).div((n - 1.0).replace(0, np.nan), axis=0)
    single = mat.notna().mul((n == 1).astype(int), axis=0).astype(bool)
    out = out.mask(single, 0.5)
    return out.where(mat.notna())


def _nanmean_frames(frames: List[pd.DataFrame], index: pd.Index, columns: List[str]) -> pd.DataFrame:
    """같은 축의 DataFrame들을 셀 단위로 평균(NaN 제외). 하나만 있으면 그대로."""
    if len(frames) == 1:
        return frames[0].reindex(index=index, columns=columns)
    stacked = pd.concat([f.reindex(index=index, columns=columns) for f in frames])
    return stacked.groupby(level=0).mean().reindex(index=index, columns=columns)


def _cap_waterfill(raw: pd.DataFrame, cap: float, strict: bool) -> Tuple[pd.DataFrame, pd.Series]:
    """raw(비음, 행합 1 또는 0)에 섹터 상한 cap을 워터필링으로 적용한다(v0.1 build_portfolio의 검증된 로직 이식):
    상한에 걸린 항목은 그 값에 고정하고 남은 예산을 아직 고정 안 된 항목에만 재분배 — 반복마다 최소 1개가
    새로 고정되므로 열 수 이내에 정확히 수렴. strict=True면 적격 섹터 수 × cap < 1이라 예산을 다 못 채울 때
    남는 몫을 현금으로 둔다(상한 엄수, 집중도 위험 관리). strict=False면 v0.1처럼 고정 항목에 비례 추가 배분.
    반환: (w, shortfall) — shortfall은 그날 현금으로 남은 비율(0~1)."""
    idx, cols = raw.index, raw.columns
    fixed = pd.DataFrame(False, index=idx, columns=cols)
    w = pd.DataFrame(0.0, index=idx, columns=cols)
    budget_left = pd.Series(1.0, index=idx)
    proportional = pd.DataFrame(0.0, index=idx, columns=cols)
    has_any = raw.sum(axis=1) > 1e-12
    for _ in range(len(cols) + 1):
        active = ~fixed
        denom = raw.where(active, 0.0).sum(axis=1).replace(0, np.nan)
        proportional = raw.where(active, 0.0).div(denom, axis=0).mul(budget_left, axis=0).fillna(0.0)
        over = active & (proportional > cap + 1e-12)
        if not over.values.any():
            break
        w = w.mask(over, cap)
        fixed = fixed | over
        budget_left = (1.0 - w.where(fixed, 0.0).sum(axis=1)).clip(lower=0.0)
    w = w.where(fixed, proportional)
    shortfall = (1.0 - w.sum(axis=1)).clip(lower=0.0).where(has_any, 0.0)
    if not strict:
        need = shortfall > 1e-9
        if need.any():
            fixed_raw = raw.where(fixed, 0.0).sum(axis=1).replace(0, np.nan)
            add = raw.where(fixed, 0.0).div(fixed_raw, axis=0).mul(shortfall, axis=0).fillna(0.0)
            w = w.add(add, fill_value=0.0)
            shortfall = pd.Series(0.0, index=idx)
    return w, shortfall


def portfolio_backtest(target_w: pd.DataFrame, ret_co: pd.DataFrame, ret_oc: pd.DataFrame,
                       rf_daily: Optional[pd.Series], cost_bps: float,
                       init_exec: Optional[pd.Series] = None, init_prev: Optional[pd.Series] = None,
                       lev_spread_bps: float = 0.0) -> pd.DataFrame:
    """다자산 T+1 시가 체결 백테스트 — M.run_backtest()의 단일자산 산식을 섹터별로 그대로 확장:
    exec_w(t) = target_w(t-1), prev_w(t) = exec_w(t-1);
    gross = Σ_i prev_w·ret_co + Σ_i exec_w·ret_oc ; 현금레그 = (1-Σ exec_w)·rf ; 비용 = Σ_i |exec_w-prev_w|·bps.
    단일 섹터·target_w=E_t인 퇴화 케이스에서 M.run_backtest(price, E_t)의 strategy_ret과 비트 동일(단위테스트).
    [v0.7.0] init_exec/init_prev: 평가창 첫날의 체결 비중(= 전날 목표)·전전날 체결 비중 — M은 SIGNAL_START 이전부터 이어진 포지션을
    들고 평가창에 들어오므로, 이를 주면 첫 2일도 M과 동일(전 기간 폴백 SPY = SPY M 비트 동일). None이면 첫날 현금(종전).
    lev_spread_bps: M v1.21 §C와 같은 초과노출(Σexec_w>1) 레버리지 스프레드(연 bp) — E_t>1인 날의 비용 정합."""
    idx, cols = target_w.index, list(target_w.columns)
    exec_w = target_w.shift(1)
    if init_exec is not None:
        exec_w.iloc[0] = init_exec.reindex(cols).fillna(0.0).values
    exec_w = exec_w.fillna(0.0)
    prev_w = exec_w.shift(1)
    if init_prev is not None:
        prev_w.iloc[0] = init_prev.reindex(cols).fillna(0.0).values
    prev_w = prev_w.fillna(0.0)
    co = ret_co.reindex(index=idx, columns=cols).fillna(0.0)
    oc = ret_oc.reindex(index=idx, columns=cols).fillna(0.0)
    gross = (prev_w * co).sum(axis=1) + (exec_w * oc).sum(axis=1)
    expo = exec_w.sum(axis=1)
    rf = rf_daily.reindex(idx).fillna(0.0) if rf_daily is not None else pd.Series(0.0, index=idx)
    cash_leg = (1.0 - expo) * rf
    turn = (exec_w - prev_w).abs().sum(axis=1)
    cost = turn * (cost_bps / 1e4)
    lev_cost = np.maximum(expo - 1.0, 0.0) * (float(lev_spread_bps) / 1e4 / 252.0)
    out = pd.DataFrame(index=idx)
    out["strategy_ret"] = (gross + cash_leg - cost - lev_cost).fillna(0.0)
    out["exposure"] = expo
    out["turnover"] = turn
    out["cost"] = cost
    out["equity"] = (1.0 + out["strategy_ret"]).cumprod()
    out["dd"] = out["equity"] / out["equity"].cummax() - 1.0
    return out


def rotation_walkforward_select(sig_full: Dict[str, pd.DataFrame], ret_cc_full: pd.DataFrame, listed_full: pd.DataFrame,
                                eval_idx: pd.DatetimeIndex, scfg: SectorConfig, M,
                                external: Optional[Dict[int, Dict[str, float]]] = None) -> Dict[str, Any]:
    """[v0.5.0] 순환매 신호의 횡단면 워크포워드 검증. 매년 1월 1일(평가 첫해부터)에 '그 이전' 데이터만으로 각 후보의
    일별 rank IC(지평 ROTATION_SELECT_HORIZON, 사전방향을 곱한 값 vs 향후 수익) 평균의 NW-HAC t를 구해
    t ≥ ROTATION_SELECT_T(+부호)이고 관측일 ≥ ROTATION_SELECT_MIN_DAYS인 신호만 그 해에 채택한다.
    채택된 신호들의 그날 횡단면 0~1 순위 평균 = 그 해의 복합순위(채택 0개면 NaN → 배분 규칙은 폴백).
    미래수익은 IC 산출에만 쓰이고, 학습창 마감(연초 − 35일 → 지평 21일의 실현 수익이 연초 전에 전부 확정)으로
    선택 자체도 인과적이다. 반환: composite_all(eval_idx×섹터), selection_log(DataFrame), selected_by_year, ic_full.
    [v0.7.0] 선택 통계 ROTATION_SELECT_STAT: "top1" = 상위1 스프레드(_top1_spread_series — 평활 순위 1위 섹터의 향후 h일 수익 −
    상장 평균)의 NW-HAC t, "ic" = 종전 rank IC의 t. 두 통계를 모두 계산해 13g에 기록하고 채택 판정에는 선택된 하나만 쓴다."""
    h = int(scfg.ROTATION_SELECT_HORIZON)
    stat = str(getattr(scfg, "ROTATION_SELECT_STAT", "top1")).lower()
    if stat not in ("top1", "ic"):
        log("ROTATION", kv(event="unknown_select_stat", value=stat, action="top1로 대체"), M=M, level="warning")
        stat = "top1"
    smooth = int(getattr(scfg, "ROTATION_SMOOTH_DAYS", 21) or 1)
    logr = np.log1p(ret_cc_full.fillna(0.0)).where(ret_cc_full.notna())
    fwd = (np.exp(logr.rolling(h).sum().shift(-h)) - 1.0).where(ret_cc_full.notna())
    ic_full: Dict[str, pd.Series] = {}
    top1_full: Dict[str, pd.Series] = {}
    rank_full: Dict[str, pd.DataFrame] = {}
    for name, mat in sig_full.items():
        if name in ROTATION_COMPOSITE_MEMBERS:
            continue   # 복합 후보는 아래에서 구성원 순위로 만든다
        sign = ROTATION_SIGNAL_SPECS.get(name, (+1, "", ""))[0]
        x = (mat * sign).where(listed_full)
        ic_full[name] = _cs_rank_ic(x, fwd)
        rank_full[name] = _cs_rank01(x)
        top1_full[name] = _top1_spread_series(rank_full[name], listed_full, fwd, smooth=smooth)
    # [v0.5.1] 사전 고정 복합 후보 — 구성원(사전방향 적용 0~1 순위)의 동일가중 평균. 구성원이 2개 미만이면 만들지 않는다.
    cols_all = list(listed_full.columns)
    for cname, members in ROTATION_COMPOSITE_MEMBERS.items():
        if cname not in scfg.ROTATION_SIGNALS and cname not in sig_full:
            continue
        avail = [m for m in members if m in rank_full]
        if len(avail) < 2:
            log("ROTATION", kv(event="composite_candidate_skipped", signal=cname, available=len(avail)), M=M, level="warning")
            continue
        comp_m = _nanmean_frames([rank_full[m] for m in avail], listed_full.index, cols_all).where(listed_full)
        rank_full[cname] = _cs_rank01(comp_m)           # 복합값을 다시 그날 순위(0~1)로 — 다른 후보와 같은 척도
        ic_full[cname] = _cs_rank_ic(comp_m, fwd)
        top1_full[cname] = _top1_spread_series(rank_full[cname], listed_full, fwd, smooth=smooth)
        log("ROTATION", kv(event="composite_candidate_built", signal=cname, members=",".join(avail)), M=M)
    years = sorted(set(eval_idx.year))
    rows: List[dict] = []
    selected_by_year: Dict[int, List[str]] = {}
    basis_by_year: Dict[int, Dict[str, str]] = {}
    mode = str(getattr(scfg, "ROTATION_SELECT_MODE", "strict")).lower()
    t_min = float(getattr(scfg, "ROTATION_SELECT_T_MIN", 1.0))
    best_n = int(getattr(scfg, "ROTATION_BEST_N", 2))
    ext_t_thr = float(getattr(scfg, "EXTERNAL_T", 2.0))
    external = external or {}
    stat_kr = "상위1 스프레드" if stat == "top1" else "rank IC"
    for y in years:
        cutoff = pd.Timestamp(year=y, month=1, day=1) - pd.Timedelta(days=35)
        stats_ic: Dict[str, Tuple[float, float, int]] = {}
        stats_top: Dict[str, Tuple[float, float, int]] = {}
        for name, ic in ic_full.items():
            stats_ic[name] = _nw_mean_tstat(ic.loc[ic.index < cutoff], lag=h)
            sp = top1_full[name]
            stats_top[name] = _nw_mean_tstat(sp.loc[sp.index < cutoff], lag=h)
        stats = stats_top if stat == "top1" else stats_ic     # [v0.7.0] 채택 판정에 쓰는 통계
        ext_y = external.get(y, {}) if isinstance(external, dict) else {}
        # ① 엄격: t ≥ ROTATION_SELECT_T
        strict = [n for n, (m, t, nn) in stats.items()
                  if nn >= scfg.ROTATION_SELECT_MIN_DAYS and pd.notna(t) and t >= scfg.ROTATION_SELECT_T]
        # ② 외부검증 지원: FF49 외부 t ≥ EXTERNAL_T & 로컬 t ≥ T_MIN(방향 일치)
        ext_sup = [n for n, (m, t, nn) in stats.items()
                   if n not in strict and nn >= scfg.ROTATION_SELECT_MIN_DAYS and pd.notna(t) and t >= t_min
                   and pd.notna(ext_y.get(n, np.nan)) and float(ext_y.get(n)) >= ext_t_thr]
        basis = {n: f"엄격(t≥{scfg.ROTATION_SELECT_T:.1f})" for n in strict}
        basis.update({n: f"외부검증 지원(FF49 t≥{ext_t_thr:.1f} & 로컬 t≥{t_min:.1f})" for n in ext_sup})
        sel = strict + ext_sup
        # ③ [v0.6.0] best_available: 아무것도 없으면 방향 일치·t ≥ T_MIN 중 상위 BEST_N을 강제 채택
        if mode == "best_available" and not sel:
            cands = sorted([(t, n) for n, (m, t, nn) in stats.items()
                            if nn >= scfg.ROTATION_SELECT_MIN_DAYS and pd.notna(t) and t >= t_min], reverse=True)[:best_n]
            sel = [n for t, n in cands]
            basis.update({n: f"최선 가용(t≥{t_min:.1f}, 상위{best_n})" for n in sel})
        for name, (m, t, nn) in stats.items():
            ext_val = ext_y.get(name, np.nan)
            m_ic, t_ic, n_ic = stats_ic[name]
            m_tp, t_tp, n_tp = stats_top[name]
            # 'NW-HAC t' 열 = 채택 판정에 쓴 통계의 t(선택 통계 열 참조). IC·상위1 두 통계의 t는 각각의 열에 항상 함께 기록.
            rows.append({"적용연도": y, "학습창 마감": cutoff.date(), "신호": name,
                         "사전방향": ROTATION_SIGNAL_SPECS.get(name, (+1, "", ""))[0], "선택 통계": stat_kr,
                         "학습 관측일": nn, "학습 평균 rank IC(사전방향 적용)": round(m_ic, 4) if pd.notna(m_ic) else np.nan,
                         "NW-HAC t(IC)": round(t_ic, 2) if pd.notna(t_ic) else np.nan,
                         f"학습 상위1 스프레드(%/{h}일)": round(m_tp * 100, 3) if pd.notna(m_tp) else np.nan,
                         "NW-HAC t(상위1)": round(t_tp, 2) if pd.notna(t_tp) else np.nan,
                         "NW-HAC t": round(t, 2) if pd.notna(t) else np.nan,
                         "외부 t(FF49)": round(float(ext_val), 2) if pd.notna(ext_val) else np.nan,
                         "채택": "채택" if name in sel else ("표본부족" if nn < scfg.ROTATION_SELECT_MIN_DAYS else "미채택"),
                         "채택 근거": basis.get(name, "")})
        selected_by_year[y] = sel
        basis_by_year[y] = basis
        log("ROTATION", kv(event="walkforward_select", year=y, train_end=str(cutoff.date()), mode=mode, stat=stat,
                           n_selected=len(sel), selected=",".join(f"{n}[{basis[n].split('(')[0]}]" for n in sel) if sel else "-",
                           top_t=";".join(f"{n}={stats[n][1]:.2f}" for n in sorted(stats, key=lambda k: -(stats[k][1] if pd.notna(stats[k][1]) else -99))[:3])), M=M)
    # 연도별 복합순위(채택 신호의 0~1 순위 평균) — eval_idx 범위
    comp = pd.DataFrame(np.nan, index=eval_idx, columns=list(listed_full.columns))
    for y in years:
        sel = selected_by_year[y]
        rows_y = eval_idx[eval_idx.year == y]
        if not sel or len(rows_y) == 0:
            continue
        comp.loc[rows_y] = _nanmean_frames([rank_full[s] for s in sel], rows_y, list(listed_full.columns)).values
    sel_log = pd.DataFrame(rows)
    return {"composite_all": comp, "selection_log": sel_log, "selected_by_year": selected_by_year,
            "basis_by_year": basis_by_year, "ic_full": ic_full, "top1_full": top1_full, "rank_full": rank_full,
            "horizon": h, "external": external, "mode": mode, "stat": stat, "smooth": smooth}


# ---- [v0.6.0] 외부 검증: Kenneth French 49업종 일별 포트폴리오 --------------------------------------------
def parse_ff49_daily_csv(text: str) -> pd.DataFrame:
    """Ken French '49_Industry_Portfolios_daily.CSV' 텍스트에서 첫 블록(Average Value Weighted Returns -- Daily)을
    DataFrame(index=날짜, 열=49업종, 값=일수익 소수)으로 파싱. -99.99(결측)는 NaN. 형식이 바뀌어도 '날짜 8자리로 시작하는 행'만
    읽으므로 헤더 위치 변화에 강건하다."""
    lines = text.splitlines()
    start = None
    for i, ln in enumerate(lines):
        if "Average Value Weighted Returns" in ln and "Daily" in ln:
            start = i
            break
    if start is None:
        # 헤더 문구가 없으면 첫 8자리 날짜 행 앞의 헤더를 찾는다
        for i, ln in enumerate(lines):
            if ln.strip()[:8].isdigit() and len(ln.strip()) > 8:
                start = i - 2
                break
    if start is None:
        raise ValueError("FF49 CSV: 데이터 블록을 찾을 수 없음")
    header = None
    data_rows: List[List[str]] = []
    for ln in lines[start + 1:]:
        st = ln.strip()
        if not st:
            if data_rows:
                break
            continue
        parts = [x.strip() for x in st.split(",")]
        if parts[0][:8].isdigit() and len(parts[0]) == 8:
            data_rows.append(parts)
        elif header is None and not parts[0]:
            header = parts[1:]
    if header is None or not data_rows:
        raise ValueError("FF49 CSV: 헤더/데이터 행 파싱 실패")
    df = pd.DataFrame([r[1:1 + len(header)] for r in data_rows], columns=header,
                      index=pd.to_datetime([r[0] for r in data_rows], format="%Y%m%d"))
    df = df.apply(pd.to_numeric, errors="coerce")
    df = df.mask(df <= -99.0) / 100.0
    df.index.name = "date"
    return df.sort_index()


def fetch_ff49_daily(scfg: SectorConfig, M=None) -> Optional[pd.DataFrame]:
    """캐시(FF49_CACHE) → 없으면 FF49_URL의 zip을 내려받아 파싱·캐시. 네트워크/파싱 실패면 None(외부 검증 생략, 로그)."""
    t0 = time.time()
    cache = getattr(scfg, "FF49_CACHE", "")
    if cache and os.path.exists(cache):
        try:
            df = pd.read_csv(cache, index_col=0, parse_dates=True)
            log("EXTERNAL", kv(event="ff49_cache_hit", file=cache, rows=len(df), cols=df.shape[1],
                               span=f"{df.index[0].date()}~{df.index[-1].date()}"), M=M)
            return df
        except Exception as e:
            log("EXTERNAL", kv(event="ff49_cache_read_failed", err=type(e).__name__), M=M, level="warning")
    try:
        import io, zipfile, urllib.request
        req = urllib.request.Request(scfg.FF49_URL, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=60) as r:
            blob = r.read()
        zf = zipfile.ZipFile(io.BytesIO(blob))
        name = [n for n in zf.namelist() if n.lower().endswith(".csv")][0]
        text = zf.read(name).decode("latin-1")
        df = parse_ff49_daily_csv(text)
        if cache:
            os.makedirs(os.path.dirname(cache) or ".", exist_ok=True)
            df.to_csv(cache)
        log("EXTERNAL", kv(event="ff49_downloaded", rows=len(df), cols=df.shape[1],
                           span=f"{df.index[0].date()}~{df.index[-1].date()}", elapsed_s=round(time.time() - t0, 1)), M=M)
        return df
    except Exception as e:
        log("EXTERNAL", kv(event="ff49_unavailable", err=type(e).__name__, msg=str(e)[:120],
                           action="외부 검증 생략 — 채택은 로컬(11섹터) 기준만 적용"), M=M, level="warning")
        return None


def ff49_signal_matrices(ff: pd.DataFrame, spy_series: Optional[Dict[str, pd.Series]] = None,
                         resid_window: int = 756) -> Dict[str, pd.DataFrame]:
    """49업종 수익행렬에서 11섹터 후보와 '같은 정의'의 순위신호를 만든다(계산 가능한 것만): 상대모멘텀 3종·잔차모멘텀·상대과열,
    (spy_series가 있으면) 국면×베타 2종. 시장 대용 = 그날 업종 동일가중 평균수익. 전부 t까지의 데이터만 사용."""
    r = ff.copy()
    mkt = r.mean(axis=1, skipna=True)
    logp = np.log1p(r).cumsum()
    logm = np.log1p(mkt).cumsum()
    out: Dict[str, pd.DataFrame] = {}
    out["REL_MOM_21"] = (logp - logp.shift(21)).sub(logm - logm.shift(21), axis=0)
    out["REL_MOM_126"] = (logp - logp.shift(126)).sub(logm - logm.shift(126), axis=0)
    out["REL_MOM_12_1"] = (logp.shift(21) - logp.shift(252)).sub(logm.shift(21) - logm.shift(252), axis=0)
    # 잔차모멘텀: 롤링 베타(lag 1) 잔차 누적 21~252
    lr = np.log1p(r); lm = np.log1p(mkt)
    cov = lr.rolling(resid_window, min_periods=252).cov(lm)
    var = lm.rolling(resid_window, min_periods=252).var()
    beta = cov.div(var.replace(0, np.nan), axis=0).shift(1)
    resid = lr.sub(beta.mul(lm, axis=0))
    rc = resid.cumsum()
    out["RESID_MOM_12_1"] = rc.shift(21) - rc.shift(252)
    P = np.exp(logp); Pm = np.exp(logm)
    ext_i = P / P.rolling(200, min_periods=150).mean() - 1.0
    ext_m = Pm / Pm.rolling(200, min_periods=150).mean() - 1.0
    out["REL_EXT_200"] = ext_i.sub(ext_m, axis=0)
    if spy_series is not None:
        b252 = lr.rolling(252, min_periods=120).cov(lm).div(lm.rolling(252, min_periods=120).var().replace(0, np.nan), axis=0).shift(1)
        sc = spy_series["SCORE_PCT"].reindex(r.index)
        hz = spy_series["HAZ_PCT"].reindex(r.index)
        out["BETA_X_MSCORE"] = b252.mul(sc - 0.5, axis=0)
        out["BETA_X_MHAZ"] = b252.mul(0.5 - hz, axis=0)
    return {k: v.replace([np.inf, -np.inf], np.nan) for k, v in out.items()}


def external_validation_ff49(scfg: SectorConfig, M, years: List[int], spy_series: Optional[Dict[str, pd.Series]] = None
                             ) -> Tuple[Dict[int, Dict[str, float]], pd.DataFrame]:
    """[v0.6.0] 연도별(적용연도 Y, 학습창 마감 = Y-01-01 − 35일 이전) 49업종 횡단면 rank IC(지평 ROTATION_SELECT_HORIZON)의
    NW-HAC t. 반환: ({Y: {신호: t}}, 로그 DataFrame). 실패/미사용이면 ({}, 빈 DF).
    [v0.7.0] 로컬과 같은 선택 통계(ROTATION_SELECT_STAT)를 외부에도 적용: "top1"이면 49업종 중 평활 순위 1위 업종의 향후 h일 수익 −
    49업종 평균(상위1 스프레드)의 t를 반환값으로, "ic"면 rank IC의 t. 두 통계 모두 로그(13h)에 기록."""
    if not getattr(scfg, "USE_EXTERNAL_VALIDATION", False):
        return {}, pd.DataFrame()
    ff = fetch_ff49_daily(scfg, M)
    if ff is None or ff.shape[1] < 20:
        return {}, pd.DataFrame()
    t0 = time.time()
    h = int(scfg.ROTATION_SELECT_HORIZON)
    stat = str(getattr(scfg, "ROTATION_SELECT_STAT", "top1")).lower()
    smooth = int(getattr(scfg, "ROTATION_SMOOTH_DAYS", 21) or 1)
    sigs = ff49_signal_matrices(ff, spy_series)
    logr = np.log1p(ff)
    fwd = (np.exp(logr.rolling(h).sum().shift(-h)) - 1.0).where(ff.notna())
    ext: Dict[int, Dict[str, float]] = {}
    rows: List[dict] = []
    ics: Dict[str, pd.Series] = {}
    tops: Dict[str, pd.Series] = {}
    avail = ff.notna()
    for name, mat in sigs.items():
        sign = ROTATION_SIGNAL_SPECS.get(name, (+1, "", ""))[0]
        x = (mat * sign).where(avail)
        ics[name] = _cs_rank_ic(x, fwd, min_n=20)
        tops[name] = _top1_spread_series(_cs_rank01(x), avail, fwd, smooth=smooth, min_n=20)
    for y in years:
        cutoff = pd.Timestamp(year=y, month=1, day=1) - pd.Timedelta(days=35)
        ext[y] = {}
        for name in sigs:
            ic, sp = ics[name], tops[name]
            m_ic, t_ic, n_ic = _nw_mean_tstat(ic.loc[ic.index < cutoff], lag=h)
            m_tp, t_tp, n_tp = _nw_mean_tstat(sp.loc[sp.index < cutoff], lag=h)
            t_use = t_tp if stat == "top1" else t_ic
            ext[y][name] = float(t_use) if pd.notna(t_use) else np.nan
            rows.append({"적용연도": y, "학습창 마감": cutoff.date(), "신호": name, "선택 통계": ("상위1 스프레드" if stat == "top1" else "rank IC"),
                         "외부 관측일": n_ic, "외부 평균 rank IC": round(m_ic, 4) if pd.notna(m_ic) else np.nan,
                         "외부 NW-HAC t(IC)": round(t_ic, 2) if pd.notna(t_ic) else np.nan,
                         f"외부 상위1 스프레드(%/{h}일)": round(m_tp * 100, 3) if pd.notna(m_tp) else np.nan,
                         "외부 NW-HAC t(상위1)": round(t_tp, 2) if pd.notna(t_tp) else np.nan,
                         "외부 NW-HAC t": round(t_use, 2) if pd.notna(t_use) else np.nan,
                         "외부 유의(t≥기준)": "Y" if (pd.notna(t_use) and t_use >= float(scfg.EXTERNAL_T)) else "N"})
    log("EXTERNAL", kv(event="ff49_validation_done", industries=ff.shape[1], signals=len(sigs), years=len(years), stat=stat,
                       span=f"{ff.index[0].date()}~{ff.index[-1].date()}", elapsed_s=round(time.time() - t0, 1)), M=M)
    return ext, pd.DataFrame(rows)


def build_sector_allocation(results: Dict[str, Dict[str, Any]], res: dict, eval_idx: pd.DatetimeIndex,
                            scfg: SectorConfig, M, rf_daily: Optional[pd.Series] = None,
                            spy_series: Optional[Dict[str, pd.Series]] = None) -> Dict[str, Any]:
    """[v0.4.0 §1.F → v0.5.0] 일별 섹터 배분비중(목표, t일 종가 확정 → t+1 시가 체결)과 배분 전략들의 백테스트.
    v0.5.0: (1) 후보 순위신호를 전체 이력으로 만들고 rotation_walkforward_select()로 매년 채택 (2) 주 전략 "leader3" =
    명확한 1위 100% / 명확한 꼴찌 회피 균등 / 폴백(SPY 또는 균등) (3) 'SPY' 열을 포트폴리오에 포함(폴백 매수 대상).
    반환 dict: target_w(주 전략, E_t 곱한 실제 비중 — 열에 SPY 포함), composite(적격 마스킹·평활 복합순위), composite_all
    (상장 전체·평활 — 검증용), tier/leader/laggard(일별 판단), eligible, listed, E, n_eligible, rank_pos, bts/perf/curve,
    diag, signals(평가창 원시 신호 행렬), wf(워크포워드 결과), ret_cc, state, cols."""
    t0 = time.time()
    cols = [t for t in scfg.SECTORS if t in results]
    cap = float(scfg.ROTATION_MAX_WEIGHT)
    E = res["sig"]["target_pos"].reindex(eval_idx).fillna(0.0).astype(float)
    if not cols:
        return {}
    full_idx = res["cal"]
    state = pd.DataFrame({t: results[t]["state"].reindex(eval_idx) for t in cols})
    listed = state.notna()
    eligible = listed & ~state.isin(list(scfg.ROTATION_EXCLUDE_STATES))
    ret_co = pd.DataFrame({t: results[t]["ret_co"].reindex(eval_idx) for t in cols})
    ret_oc = pd.DataFrame({t: results[t]["ret_oc"].reindex(eval_idx) for t in cols})
    ret_cc = pd.DataFrame({t: results[t]["bh_ret"].reindex(eval_idx) for t in cols})
    # SPY(폴백 매수 대상) 수익 분해 — M의 bt(run_backtest 산출)에서 그대로
    spy_bt = res["bt"]
    ret_co["SPY"] = spy_bt["ret_co"].reindex(eval_idx)
    ret_oc["SPY"] = spy_bt["ret_oc"].reindex(eval_idx)
    all_cols = cols + ["SPY"]

    # ---- 전체 이력 원시 신호 행렬(날짜×섹터) + 국면×베타 결합 ----
    ret_cc_full = pd.DataFrame({t: results[t]["ret_cc_full"].reindex(full_idx) for t in cols})
    listed_full = ret_cc_full.notna()
    raw_full: Dict[str, pd.DataFrame] = {}
    for name in scfg.ROTATION_SIGNALS:
        if name not in ROTATION_SIGNAL_SPECS:
            log("ROTATION", kv(event="unknown_rotation_signal", signal=name), M=M, level="warning")
            continue
        if name in ROTATION_COMPOSITE_MEMBERS:
            continue   # [v0.5.1] 복합 후보는 원시 행렬이 없다 — rotation_walkforward_select()가 구성원 순위로 만든다
        if name in ("BETA_X_MSCORE", "BETA_X_MHAZ"):
            if spy_series is None:
                log("ROTATION", kv(event="signal_unavailable", signal=name, reason="spy_series 없음"), M=M, level="warning")
                continue
            beta = pd.DataFrame({t: results[t]["rot_raw"]["BETA"].reindex(full_idx) for t in cols
                                 if isinstance(results[t].get("rot_raw"), pd.DataFrame) and "BETA" in results[t]["rot_raw"].columns})
            if beta.empty:
                continue
            if name == "BETA_X_MSCORE":
                f = (spy_series["SCORE_PCT"].reindex(full_idx) - 0.5)
            else:
                f = (0.5 - spy_series["HAZ_PCT"].reindex(full_idx))
            raw_full[name] = beta.mul(f, axis=0).reindex(columns=cols)
            continue
        parts = {}
        for t in cols:
            rr = results[t].get("rot_raw")
            if isinstance(rr, pd.DataFrame) and name in rr.columns and rr[name].notna().any():
                parts[t] = rr[name].reindex(full_idx)
        if parts:
            raw_full[name] = pd.DataFrame(parts).reindex(columns=cols)
        else:
            log("ROTATION", kv(event="signal_unavailable_all_sectors", signal=name), M=M, level="warning")
    if not raw_full:
        raise RuntimeError("ROTATION_SIGNALS 중 사용 가능한 신호가 없습니다 — 배분 계층을 만들 수 없음")
    signals = {k: v.reindex(eval_idx) for k, v in raw_full.items()}   # 평가창(13d 진단용)

    # ---- 횡단면 워크포워드 검증 → 연도별 채택 신호 → 복합순위 ----
    # [v0.6.0] 외부 검증(FF49) — 실패/미사용이면 빈 dict(로컬 기준만 적용)
    external: Dict[int, Dict[str, float]] = {}
    ext_log = pd.DataFrame()
    try:
        external, ext_log = external_validation_ff49(scfg, M, sorted(set(eval_idx.year)), spy_series)
    except Exception as e:
        log("EXTERNAL", kv(event="ff49_validation_failed", err=type(e).__name__, msg=str(e)[:160]), M=M, level="warning")
    wf = rotation_walkforward_select(raw_full, ret_cc_full, listed_full, eval_idx, scfg, M, external=external)
    wf["external_log"] = ext_log
    for cname in ROTATION_COMPOSITE_MEMBERS:          # [v0.5.1] 13d 진단표에도 복합 후보를 올린다(평가창 순위 행렬)
        if cname in wf["rank_full"]:
            signals[cname] = wf["rank_full"][cname].reindex(eval_idx)
    composite_all = wf["composite_all"].where(listed)
    smooth = int(getattr(scfg, "ROTATION_SMOOTH_DAYS", 21) or 1)
    composite_all_s = composite_all.rolling(smooth, min_periods=1).mean().where(listed) if smooth > 1 else composite_all
    composite = composite_all_s.where(eligible)
    rankable = composite.notna()
    n_elig = rankable.sum(axis=1)
    rank_pos = composite.rank(axis=1, ascending=False, method="first")   # 1 = 그날 최고
    # 채택 신호별(평활) 순위 — 과반 투표용
    rank_sel_s: Dict[str, pd.DataFrame] = {}
    for name, rk in wf["rank_full"].items():
        r = rk.reindex(eval_idx).where(listed)
        rank_sel_s[name] = (r.rolling(smooth, min_periods=1).mean().where(listed) if smooth > 1 else r).where(eligible)

    # ---- tilt → 비율(E_t 대비) → 상한 ----
    def _frac_from_tilt(tilt: pd.DataFrame) -> Tuple[pd.DataFrame, pd.Series]:
        s = tilt.sum(axis=1).replace(0, np.nan)
        raw = tilt.div(s, axis=0).fillna(0.0)
        w, short = _cap_waterfill(raw, cap, scfg.ROTATION_CAP_STRICT)
        w["SPY"] = 0.0
        return w.reindex(columns=all_cols), short

    tilt_linear = composite.rank(axis=1, ascending=True, method="average").fillna(0.0)
    tilt_topk = ((rank_pos <= scfg.ROTATION_TOP_K) & rankable).astype(float)
    frac_lin, short_lin = _frac_from_tilt(tilt_linear)
    frac_topk, short_topk = _frac_from_tilt(tilt_topk)
    frac_ctrl_b, _ = _frac_from_tilt(eligible.astype(float))
    # [v0.5.0] 비교용 상위K/순위가중도 '그 해 채택 신호가 없어 순위가 없는 날'은 주 전략과 같은 폴백을 쓴다(비교 공정성)
    no_rank = n_elig == 0
    if no_rank.any():
        if str(scfg.ROTATION_FALLBACK).lower() == "spy":
            for fr in (frac_lin, frac_topk):
                fr.loc[no_rank, :] = 0.0
                fr.loc[no_rank, "SPY"] = 1.0
        else:
            eq = eligible.astype(float).div(eligible.sum(axis=1).replace(0, np.nan), axis=0).fillna(0.0)
            for fr in (frac_lin, frac_topk):
                fr.loc[no_rank, cols] = eq.loc[no_rank, cols].values
    n_listed = listed.sum(axis=1).replace(0, np.nan)
    frac_ctrl_a = listed.astype(float).div(n_listed, axis=0).fillna(0.0)
    frac_ctrl_a["SPY"] = 0.0
    frac_ctrl_a = frac_ctrl_a.reindex(columns=all_cols)

    # ---- [v0.5.0] 주 전략 leader3: 명확한 1위 / 명확한 꼴찌 회피 / 폴백 ----
    years_arr = eval_idx.year
    sel_by_year = wf["selected_by_year"]
    frac_leader = pd.DataFrame(0.0, index=eval_idx, columns=all_cols)
    tier = pd.Series("현금", index=eval_idx, dtype=object)
    leader_s = pd.Series("", index=eval_idx, dtype=object)
    laggard_s = pd.Series("", index=eval_idx, dtype=object)
    votes_leader = pd.Series(0, index=eval_idx, dtype=int)
    votes_laggard = pd.Series(0, index=eval_idx, dtype=int)
    n_sel_s = pd.Series(0, index=eval_idx, dtype=int)
    comp_vals = composite.values
    elig_vals = rankable.values
    col_arr = np.array(cols)
    # 신호별 argmax/argmin(적격 내) 미리 계산 — 전부 NaN인 행은 None(pandas idxmax는 all-NA 행에서 예외)
    def _row_arg(r: pd.DataFrame, fn: str) -> pd.Series:
        out = pd.Series([None] * len(r), index=r.index, dtype=object)
        has = r.notna().any(axis=1)
        if has.any():
            out[has] = getattr(r[has], fn)(axis=1).astype(object)
        return out
    am = {n: _row_arg(r, "idxmax") for n, r in rank_sel_s.items()}
    an = {n: _row_arg(r, "idxmin") for n, r in rank_sel_s.items()}
    cur_leader: Optional[str] = None
    held = 0
    min_hold = int(scfg.ROTATION_MIN_HOLD_DAYS)
    lw = float(scfg.ROTATION_LEADER_WEIGHT)
    fallback_spy = str(scfg.ROTATION_FALLBACK).lower() == "spy"
    switches = 0
    for i, d in enumerate(eval_idx):
        sel = sel_by_year.get(int(years_arr[i]), [])
        K = len(sel)
        n_sel_s.iloc[i] = K
        row = comp_vals[i]
        ok = elig_vals[i] & ~np.isnan(row)
        n_ok = int(ok.sum())
        leader = laggard = None
        v_lead = v_lag = 0
        if n_ok >= 1 and K > 0:
            leader = col_arr[np.nanargmax(np.where(ok, row, -np.inf))]
            v_lead = sum(1 for s in sel if s in am and am[s].iloc[i] == leader)
            if n_ok >= 4:
                laggard = col_arr[np.nanargmin(np.where(ok, row, np.inf))]
                v_lag = sum(1 for s in sel if s in an and an[s].iloc[i] == laggard)
        clear_leader = leader is not None and v_lead * 2 > K
        clear_laggard = laggard is not None and v_lag * 2 > K
        votes_leader.iloc[i], votes_laggard.iloc[i] = v_lead, v_lag
        # 최소보유 상태기계(적격 상실 즉시 청산)
        if cur_leader is not None and not (ok[cols.index(cur_leader)] if cur_leader in cols else False):
            cur_leader = None
        if clear_leader and leader != cur_leader:
            if cur_leader is None or held >= min_hold:
                cur_leader = leader
                held = 0
                switches += 1
        elif not clear_leader and cur_leader is not None and held >= min_hold:
            cur_leader = None
        if cur_leader is not None:
            frac_leader.iat[i, all_cols.index(cur_leader)] = lw
            tier.iloc[i] = "리더"
            leader_s.iloc[i] = cur_leader
            held += 1
        elif clear_laggard:
            basket = [c for j, c in enumerate(cols) if ok[j] and c != laggard]
            for c in basket:
                frac_leader.iat[i, all_cols.index(c)] = 1.0 / len(basket)
            tier.iloc[i] = "회피"
            laggard_s.iloc[i] = laggard
        else:
            if fallback_spy:
                frac_leader.iat[i, all_cols.index("SPY")] = 1.0
            elif n_ok > 0:
                for j, c in enumerate(cols):
                    if ok[j]:
                        frac_leader.iat[i, all_cols.index(c)] = 1.0 / n_ok
            tier.iloc[i] = "폴백"
        if E.iloc[i] <= 1e-12:
            tier.iloc[i] = "현금"
    fallback_txt = "SPY" if fallback_spy else "균등"
    label_leader = f"집중배분(명확1위 {lw:.0%}·꼴찌회피·폴백{fallback_txt}) ★"
    label_topk = ROT_LABEL_TOPK.format(k=scfg.ROTATION_TOP_K, cap=cap)
    label_lin = ROT_LABEL_PRIMARY.format(cap=cap).replace(" ★", "")
    mode = str(scfg.ROTATION_TILT).lower()
    if mode == "topk":
        label_primary, frac_primary = label_topk + " ★", frac_topk
        label_leader = label_leader.replace(" ★", "")
    elif mode == "linear":
        label_primary, frac_primary = label_lin + " ★", frac_lin
        label_leader = label_leader.replace(" ★", "")
    else:
        label_primary, frac_primary = label_leader, frac_leader
    variants = {label_primary: frac_primary}
    for lab, fr in ((label_leader, frac_leader), (label_topk, frac_topk), (label_lin, frac_lin)):
        if lab != label_primary:
            variants[lab] = fr
    variants[ROT_LABEL_CTRL_A] = frac_ctrl_a
    variants[ROT_LABEL_CTRL_B] = frac_ctrl_b

    cost_bps = float(res["cfg"].COST_BPS)
    lev_bps = float(getattr(res["cfg"], "LEVERAGE_SPREAD_BPS", 0.0) or 0.0)   # [v0.7.0] M v1.21 §C 초과노출 스프레드 정합
    # [v0.7.0] 평가창 진입 시 초기 포지션 = SPY M이 그 전날·전전날 들고 있던 SPY 비중(M은 SIGNAL_START 이전부터 포지션을 이어 온다).
    #   모든 변형에 같은 초기값 → 전 기간 폴백(SPY)이면 SPY M과 첫날부터 비트 동일, 대조군도 첫 2일만 M 포지션에서 출발.
    E_full = res["sig"]["target_pos"].reindex(full_idx).fillna(0.0).astype(float)
    i0 = int(full_idx.get_indexer([eval_idx[0]])[0]) if len(eval_idx) else -1
    init_exec = pd.Series(0.0, index=all_cols); init_prev = pd.Series(0.0, index=all_cols)
    if i0 >= 1:
        init_exec["SPY"] = float(E_full.iloc[i0 - 1])
    if i0 >= 2:
        init_prev["SPY"] = float(E_full.iloc[i0 - 2])
    bt_kw = dict(rf_daily=rf_daily, cost_bps=cost_bps, init_exec=init_exec, init_prev=init_prev, lev_spread_bps=lev_bps)
    listed_all = listed.copy()
    listed_all["SPY"] = True
    listed_all = listed_all.reindex(columns=all_cols)
    bts: Dict[str, pd.DataFrame] = {}
    target_ws: Dict[str, pd.DataFrame] = {}
    for label, frac in variants.items():
        tw = frac.reindex(columns=all_cols).mul(E, axis=0)
        bad = (tw.abs() > 1e-12) & ~listed_all
        if bad.values.any():
            log("ROTATION", kv(event="weight_on_unlisted_sector_zeroed", strategy=label, cells=int(bad.values.sum())),
                M=M, level="warning")
            tw = tw.where(~bad, 0.0)
        target_ws[label] = tw
        bts[label] = portfolio_backtest(tw, ret_co, ret_oc, **bt_kw)
    # [v0.7.0] 비교 변형 '리더 자체 목표비중': 리더 보유일에는 E_t 대신 그 섹터의 M식 target_pos(섹터 자체 국면 판단 크기)를
    #   노출로 쓴다(E_t=0인 날은 주 전략과 같이 현금 — M의 위험회피 게이트는 유지). 회피·폴백일은 주 전략과 동일. 13시트 비교용이며
    #   기본 실행 비중(target_w)은 여전히 E_t를 따른다(사용자 요청: 국면 판단 비중대로).
    label_own = f"집중배분(리더 자체 목표비중·꼴찌회피·폴백{fallback_txt}) [비교]"
    if getattr(scfg, "ROTATION_ALT_LEADER_OWN_POS", True):
        own_pos = pd.DataFrame({t: results[t]["target_pos"].reindex(eval_idx) for t in cols}).fillna(0.0).astype(float)
        tw_own = target_ws[label_leader].copy()
        is_leader_day = (tier == "리더").values
        for i in np.flatnonzero(is_leader_day):
            c = leader_s.iloc[i]
            if c in cols:
                tw_own.iat[i, all_cols.index(c)] = float(own_pos.iat[i, cols.index(c)]) * lw
        target_ws[label_own] = tw_own
        bts[label_own] = portfolio_backtest(tw_own, ret_co, ret_oc, **bt_kw)
        variants[label_own] = tw_own   # 성과표 순서용(비중 자체는 위에서 계산 완료)
    # [v0.7.0] 참조: SPY 국면전략(M) 성과(같은 평가창, M의 bt 그대로) — 수용기준 ⑤(목표: CAGR ≥ SPY M)에 사용
    spy_m_ret = res["bt"]["strategy_ret"].reindex(eval_idx).fillna(0.0)
    spy_m_pm = M.perf_metrics(spy_m_ret, "SPY 국면전략(M)")
    spy_m = {"CAGR": float(spy_m_pm["CAGR"]), "MDD": float(spy_m_pm["최대낙폭(MDD)"]), "샤프": spy_m_pm["샤프"]}

    # 13 성과표 순서: 주 전략 → 대안·비교 변형 → 대조군 A/B
    order = [l for l in variants if l not in (ROT_LABEL_CTRL_A, ROT_LABEL_CTRL_B)] + [ROT_LABEL_CTRL_A, ROT_LABEL_CTRL_B]
    perf = pd.DataFrame([M.perf_metrics(bts[l]["strategy_ret"], l) for l in order])
    perf.insert(1, "평가창", f"{eval_idx[0].date()}~{eval_idx[-1].date()}")
    perf["평균노출"] = [round(float(bts[l]["exposure"].mean()), 4) for l in order]
    perf["연평균회전율(편도)"] = [round(float(bts[l]["turnover"].sum() / (len(eval_idx) / 252)), 2) for l in order]
    perf["누적비용"] = [round(float(bts[l]["cost"].sum()), 4) for l in order]
    curve = pd.DataFrame(index=eval_idx)
    curve["날짜"] = [d.date() for d in eval_idx]
    for l in order:
        curve[l] = bts[l]["equity"].round(4)
    curve["E_t(SPY목표비중)"] = E.round(2)
    curve["적격섹터수"] = n_elig.astype(int).values

    target_w = target_ws[label_primary]
    tier_counts = tier.value_counts().to_dict()
    diag = {
        "label_primary": label_primary, "label_leader": label_leader, "label_topk": label_topk, "label_linear": label_lin,
        "label_alt": (label_topk if mode == "leader3" else label_leader),
        "label_own": (label_own if label_own in bts else None), "spy_m": spy_m,          # [v0.7.0]
        "select_stat": wf.get("stat", "top1"),
        "tilt": mode, "cap": cap, "cap_strict": bool(scfg.ROTATION_CAP_STRICT),
        "signals": list(raw_full.keys()), "smooth_days": smooth, "min_hold_days": min_hold, "leader_weight": lw,
        "fallback": fallback_txt, "select_t": float(scfg.ROTATION_SELECT_T), "select_horizon": wf["horizon"],
        "selected_by_year": {int(k): v for k, v in sel_by_year.items()},
        "avg_n_eligible": round(float(n_elig.mean()), 2), "days_no_eligible": int((n_elig == 0).sum()),
        "days_leader": int(tier_counts.get("리더", 0)), "days_avoid": int(tier_counts.get("회피", 0)),
        "days_fallback": int(tier_counts.get("폴백", 0)), "days_cash": int(tier_counts.get("현금", 0)),
        "leader_switches": switches,
        "avg_exposure_primary": round(float(bts[label_primary]["exposure"].mean()), 4),
        "avg_E": round(float(E.mean()), 4),
        "top_holding_freq": {t: round(float((rank_pos[t] == 1).mean()), 4) for t in cols},
        "leader_freq": {t: round(float((leader_s == t).mean()), 4) for t in cols},
    }
    log("ROTATION", kv(event="allocation_built",
                       **{k: v for k, v in diag.items() if k not in ("top_holding_freq", "leader_freq", "selected_by_year", "spy_m")},
                       spy_m_cagr=spy_m["CAGR"], spy_m_mdd=spy_m["MDD"],
                       elapsed_s=round(time.time() - t0, 2)), M=M)
    for l in variants:
        pr = perf.set_index("전략").loc[l]
        log("ROTATION", kv(event="allocation_perf", strategy=l, cagr=pr["CAGR"], sharpe=pr["샤프"],
                           mdd=pr["최대낙폭(MDD)"], avg_exposure=pr["평균노출"], turnover_yr=pr["연평균회전율(편도)"]), M=M)
    return {"target_w": target_w, "frac_w": frac_primary, "composite": composite, "composite_all": composite_all,
            "composite_all_smooth": composite_all_s, "tier": tier, "leader": leader_s, "laggard": laggard_s,
            "votes_leader": votes_leader, "votes_laggard": votes_laggard, "n_selected": n_sel_s,
            "eligible": eligible, "listed": listed, "E": E, "n_eligible": n_elig,
            "shortfall": pd.Series(0.0, index=eval_idx),
            "rank_pos": rank_pos, "bts": bts, "target_ws": target_ws, "perf": perf, "curve": curve.reset_index(drop=True),
            "diag": diag, "signals": signals, "wf": wf, "ret_cc": ret_cc, "state": state, "cols": cols,
            "all_cols": all_cols}


def _nw_mean_tstat(x: pd.Series, lag: int) -> Tuple[float, float, int]:
    """일별 IC 시계열의 평균이 0과 다른지 — 뉴이-웨스트 HAC(바틀렛 가중, lag) 표준오차의 평균-t(v0.1 이식)."""
    s = x.dropna()
    n = len(s)
    if n < 30:
        return np.nan, np.nan, n
    m = float(s.mean())
    resid = (s - m).values
    gamma0 = float((resid ** 2).mean())
    lag_max = min(lag, n - 1)
    var = gamma0
    for L in range(1, lag_max + 1):
        w = 1.0 - L / (lag_max + 1)
        var += 2 * w * float((resid[L:] * resid[:-L]).mean())
    se = math.sqrt(max(var, 1e-12) / n)
    return m, (m / se if se > 0 else np.nan), n


def _cs_rank_ic(x_mat: pd.DataFrame, fwd_mat: pd.DataFrame, min_n: int = 5) -> pd.Series:
    """행(날짜)별 스피어만 순위상관(그날 신호 vs 그날 기준 향후 수익) 시계열. 두 값이 모두 있는 섹터 ≥ min_n인 날만."""
    both = x_mat.notna() & fwd_mat.notna()
    xr = x_mat.where(both).rank(axis=1)
    yr = fwd_mat.where(both).rank(axis=1)
    n = both.sum(axis=1)
    xm = xr.sub(xr.mean(axis=1), axis=0)
    ym = yr.sub(yr.mean(axis=1), axis=0)
    cov = (xm * ym).sum(axis=1)
    den = np.sqrt((xm ** 2).sum(axis=1) * (ym ** 2).sum(axis=1))
    ic = (cov / den.replace(0, np.nan)).where(n >= min_n)
    return ic


def _top1_spread_series(rank_mat: pd.DataFrame, avail: pd.DataFrame, fwd_mat: pd.DataFrame,
                        smooth: int = 1, min_n: int = 5) -> pd.Series:
    """[v0.7.0] '상위1 스프레드' 일별 시계열 = (그날 평활 순위가 가장 높은 1개 자산의 향후 h일 수익) − (그날 순위·수익이 모두
    있는 자산 전체의 평균 향후 h일 수익). '1위 하나에 집중'하는 전략이 실제로 얻는 초과수익의 직접 추정치(비용 제외)로, 평균
    rank IC('전체 순서'의 정확도)와 달리 1위 판별력만 잰다. rank_mat: 0~1 횡단면 순위(사전방향 적용), avail: 순위 대상 마스크
    (상장), fwd_mat: t 기준 향후 h일 수익(통계 산출에만 쓰이며 배분에는 되돌아가지 않음). 평활은 배분에 쓰는 것과 같은
    후행 이동평균(인과). 동률이면 첫 열. 자산 수 < min_n인 날은 NaN."""
    cols = list(rank_mat.columns)
    r = rank_mat.where(avail.reindex(index=rank_mat.index, columns=cols).fillna(False))
    rs = r.rolling(smooth, min_periods=1).mean().where(r.notna()) if smooth > 1 else r
    fwd = fwd_mat.reindex(index=rank_mat.index, columns=cols)
    both = rs.notna() & fwd.notna()
    n = both.sum(axis=1)
    has = both.any(axis=1).values
    pos = rs.where(both).fillna(-np.inf).values.argmax(axis=1)
    fv = fwd.values
    r_top = np.where(has, fv[np.arange(len(fv)), pos], np.nan)
    bench = fwd.where(both).mean(axis=1)
    spread = pd.Series(r_top, index=rank_mat.index) - bench
    return spread.where(n >= min_n)


def rotation_validation(alloc: Dict[str, Any], scfg: SectorConfig, M,
                        perf_ctrl_label: str = ROT_LABEL_CTRL_A) -> Dict[str, Any]:
    """[v0.4.0 §1.F.1/§1.F.3] 횡단면 rank IC(신호별·복합, 지평 5/21/63) + 상위3−하위3 일별 스프레드(3구간) +
    사전 고정 수용기준 판정. 전부 진단 — 배분비중 계산에는 어떤 값도 되돌아가지 않는다(미래수익은 여기서만 사용)."""
    if not alloc:
        return {}
    t0 = time.time()
    ret_cc = alloc["ret_cc"]
    eval_idx = ret_cc.index
    h_main = int(scfg.ROTATION_IC_HORIZON)
    horizons = sorted({5, h_main, 63})
    logr = np.log1p(ret_cc.fillna(0.0)).where(ret_cc.notna())

    def _fwd(h: int) -> pd.DataFrame:
        return (np.exp(logr.rolling(h).sum().shift(-h)) - 1.0).where(ret_cc.notna())

    ic_rows: List[dict] = []
    # [v0.5.0] 원시 신호는 사전방향을 곱해 평가(+면 '앞선다' 방향이 맞다는 뜻). 복합순위는 워크포워드 채택 신호로만 만든
    # 것이라 평가창(2018~) IC는 사실상 표본외 추정치다.
    sig_mats: Dict[str, pd.DataFrame] = {}
    for name, mat in alloc["signals"].items():
        sign = ROTATION_SIGNAL_SPECS.get(name, (+1, "", ""))[0]
        sig_mats[name] = mat * sign
    smooth_days = int(alloc.get("diag", {}).get("smooth_days", 1))
    KEY_RAW = "복합순위(워크포워드 채택, 원시, 상장 전체)"
    KEY_USED = f"복합순위(워크포워드 채택, {smooth_days}일 평활 — 배분 사용, 상장 전체)"
    KEY_ELIG = "복합순위(평활, 적격만)"
    sig_mats[KEY_RAW] = alloc["composite_all"]
    sig_mats[KEY_USED] = alloc.get("composite_all_smooth", alloc["composite_all"])
    sig_mats[KEY_ELIG] = alloc["composite"]
    # [v0.7.0] 선택 통계에 맞춘 ③: "top1"이면 배분에 쓰는 평활 복합순위의 '상위1 스프레드' t, "ic"면 종전 rank IC t.
    #   13d에는 두 통계를 후보별·지평별로 모두 기록(상위1 스프레드는 원시 후보에도 같은 21일 평활을 적용해 배분 조건과 맞춤).
    stat = str((alloc.get("wf") or {}).get("stat", getattr(scfg, "ROTATION_SELECT_STAT", "top1"))).lower()
    listed = alloc["listed"]
    composite_t = np.nan          # ③에 쓰는 t(선택 통계 기준)
    composite_ic_t = np.nan
    composite_top1_t = np.nan
    composite_top1_mean = np.nan
    for name, xm in sig_mats.items():
        spec = ROTATION_SIGNAL_SPECS.get(name)
        is_raw = spec is not None
        for h in horizons:
            fwd_h = _fwd(h)
            ic = _cs_rank_ic(xm, fwd_h)
            m, tstat, n = _nw_mean_tstat(ic, lag=h)
            # 상위1 스프레드: 원시 후보는 0~1 순위 후 평활(배분과 동일 조건), 복합순위 3종은 이미 순위·평활 상태라 그대로
            rk = _cs_rank01(xm.where(listed)) if is_raw else xm
            sp = _top1_spread_series(rk, listed, fwd_h, smooth=(smooth_days if is_raw else 1))
            m_tp, t_tp, n_tp = _nw_mean_tstat(sp, lag=h)
            t_sel = t_tp if stat == "top1" else tstat
            ic_rows.append({"신호": name, "종류": ("후보(원시, 사전방향 적용)" if spec else "복합순위"),
                            "사전방향": (spec[0] if spec else "+1"), "지평(일)": h,
                            "평균 rank IC": round(m, 4) if pd.notna(m) else np.nan,
                            "NW-HAC t": round(tstat, 2) if pd.notna(tstat) else np.nan, "검사일수": n,
                            "IC>0 비율": round(float((ic.dropna() > 0).mean()), 4) if n else np.nan,
                            "상위1 스프레드(%/지평)": round(m_tp * 100, 3) if pd.notna(m_tp) else np.nan,
                            "NW-HAC t(상위1)": round(t_tp, 2) if pd.notna(t_tp) else np.nan,
                            "판정": ("유의(+)" if (pd.notna(t_sel) and t_sel >= scfg.ROTATION_ACCEPT_IC_T)
                                   else ("유의(−)" if (pd.notna(t_sel) and t_sel <= -scfg.ROTATION_ACCEPT_IC_T) else "비유의")),
                            "판정 통계": ("상위1 스프레드" if stat == "top1" else "rank IC"),
                            "설명": (spec[1] if spec else ""), "경제적 근거": (spec[2] if spec else "")})
            if name == KEY_USED and h == h_main:
                composite_ic_t, composite_top1_t, composite_top1_mean = tstat, t_tp, m_tp
                composite_t = t_sel
    ic_table = pd.DataFrame(ic_rows)

    # ---- 상위3−하위3 일별 스프레드(배분에 쓰는 평활 복합순위, 상장 전체; 익일 종가수익, 비용 없음 — 진단) ----
    comp = alloc.get("composite_all_smooth", alloc["composite_all"])
    nxt = ret_cc.shift(-1)
    rk = comp.rank(axis=1, ascending=False, method="first")
    n_avail = comp.notna().sum(axis=1)
    top = nxt.where(rk <= 3).mean(axis=1)
    bot = nxt.where(rk.gt(n_avail - 3, axis=0) & comp.notna()).mean(axis=1)
    spread = (top - bot).where(n_avail >= 6)
    periods = [("2018-01-01", "2020-12-31"), ("2021-01-01", "2023-12-31"), ("2024-01-01", "2026-12-31")]
    sp_rows: List[dict] = []
    n_pos = 0
    for a, b in periods:
        s = spread.loc[(spread.index >= pd.Timestamp(a)) & (spread.index <= pd.Timestamp(b))].dropna()
        if len(s) < 60:
            sp_rows.append({"구간": f"{a[:4]}~{b[:4]}", "일수": len(s), "일평균 스프레드(%)": np.nan, "연환산(%)": np.nan, "양수": "표본부족"})
            continue
        mean_d = float(s.mean())
        n_pos += int(mean_d > 0)
        sp_rows.append({"구간": f"{a[:4]}~{b[:4]}", "일수": len(s), "일평균 스프레드(%)": round(mean_d * 100, 4),
                        "연환산(%)": round(mean_d * 252 * 100, 2), "양수": "Y" if mean_d > 0 else "N"})
    m_all, t_all, n_all = _nw_mean_tstat(spread, lag=1)
    sp_rows.append({"구간": "전체", "일수": n_all, "일평균 스프레드(%)": round(m_all * 100, 4) if pd.notna(m_all) else np.nan,
                    "연환산(%)": round(m_all * 252 * 100, 2) if pd.notna(m_all) else np.nan,
                    "양수": f"NW-t {t_all:.2f}" if pd.notna(t_all) else "-"})
    spread_table = pd.DataFrame(sp_rows)

    # ---- 수용기준(§1.F.3, 사전 고정) — 주 전략 vs 대조군A ----
    perf = alloc["perf"].set_index("전략")
    lp = alloc["diag"]["label_primary"]
    p_pri, p_ctl = perf.loc[lp], perf.loc[perf_ctrl_label]
    cagr_gain = float(p_pri["CAGR"] - p_ctl["CAGR"])
    mdd_worse = float(p_ctl["최대낙폭(MDD)"] - p_pri["최대낙폭(MDD)"])   # 양수 = 주 전략의 MDD가 더 깊음
    if stat == "top1":
        crit3_name = f"③ 복합순위 상위1 스프레드 NW-t ≥ {scfg.ROTATION_ACCEPT_IC_T:.1f} (지평 {h_main}일, 평활 순위 1위 − 상장 평균)"
        crit3_val = (f"t = {composite_t:.2f} (평균 {composite_top1_mean * 100:+.3f}%/{h_main}일; 참고 rank IC t = {composite_ic_t:.2f})"
                     if pd.notna(composite_t) else "계산불가")
    else:
        crit3_name = f"③ 복합순위 rank IC NW-t ≥ {scfg.ROTATION_ACCEPT_IC_T:.1f} (지평 {h_main}일)"
        crit3_val = (f"t = {composite_t:.2f} (참고 상위1 스프레드 t = {composite_top1_t:.2f})" if pd.notna(composite_t) else "계산불가")
    # [v0.7.0] ⑤ 목표: 주 전략 CAGR ≥ SPY 국면전략(M) CAGR + 여유 — 이 계층이 M만 쓰는 것보다 나은가(사용자 목표 그대로)
    spy_m = alloc.get("diag", {}).get("spy_m", {})
    spy_m_cagr = spy_m.get("CAGR", np.nan)
    vs_spy = float(p_pri["CAGR"] - spy_m_cagr) if pd.notna(spy_m_cagr) else np.nan
    margin5 = float(getattr(scfg, "ROTATION_ACCEPT_VS_SPY_M", 0.0))
    crit = [
        {"기준": f"① CAGR ≥ 대조군A + {scfg.ROTATION_ACCEPT_CAGR_GAIN:.1%}p",
         "측정값": f"{cagr_gain:+.2%}p (주 {p_pri['CAGR']:.2%} vs 대조군A {p_ctl['CAGR']:.2%})",
         "판정": "PASS" if cagr_gain >= scfg.ROTATION_ACCEPT_CAGR_GAIN else "FAIL"},
        {"기준": f"② MDD 악화 ≤ {scfg.ROTATION_ACCEPT_MDD_WORSE:.0%}p",
         "측정값": f"{mdd_worse:+.2%}p (주 {p_pri['최대낙폭(MDD)']:.2%} vs 대조군A {p_ctl['최대낙폭(MDD)']:.2%})",
         "판정": "PASS" if mdd_worse <= scfg.ROTATION_ACCEPT_MDD_WORSE else "FAIL"},
        {"기준": crit3_name, "측정값": crit3_val,
         "판정": "PASS" if (pd.notna(composite_t) and composite_t >= scfg.ROTATION_ACCEPT_IC_T) else "FAIL"},
        {"기준": f"④ 상위3−하위3 스프레드 양수 구간 ≥ {scfg.ROTATION_ACCEPT_SPREAD_PERIODS}/3",
         "측정값": f"{n_pos}/3 구간 양수",
         "판정": "PASS" if n_pos >= scfg.ROTATION_ACCEPT_SPREAD_PERIODS else "FAIL"},
        {"기준": f"⑤ 목표: CAGR ≥ SPY 국면전략(M)" + (f" + {margin5:.1%}p" if margin5 > 0 else ""),
         "측정값": (f"{vs_spy:+.2%}p (주 {p_pri['CAGR']:.2%} vs SPY M {spy_m_cagr:.2%}; MDD 주 {p_pri['최대낙폭(MDD)']:.2%} vs "
                  f"SPY M {spy_m.get('MDD', np.nan):.2%}, 샤프 주 {p_pri['샤프']} vs SPY M {spy_m.get('샤프', np.nan)})"
                  if pd.notna(vs_spy) else "계산불가(SPY M 성과 없음)"),
         "판정": "PASS" if (pd.notna(vs_spy) and vs_spy >= margin5) else "FAIL"},
    ]
    # [v0.5.0] 워크포워드 채택 현황 — 채택 신호가 한 해도 없으면 배분은 전 기간 폴백(=사실상 M 전략)이었다는 뜻
    sel_by_year = alloc.get("diag", {}).get("selected_by_year", {})
    years_with = [y for y, s in sel_by_year.items() if s]
    sel_items = ["%s:%s" % (y, "+".join(s)) for y, s in sorted(sel_by_year.items()) if s]
    sel_txt = (f"워크포워드 채택: {len(years_with)}/{len(sel_by_year)}개 연도에서 신호 채택"
               + ((" (" + ", ".join(sel_items) + ")") if years_with else " — 전 기간 미채택(폴백만 적용)"))
    # [v0.5.1] 채택 신호가 한 해도 없으면 순위 자체가 없으므로 ③④는 '해당 없음(N/A)' — FAIL로 표기하지 않는다.
    #   이때 ①②는 '폴백(SPY) vs 대조군A(균등 11)'의 비교일 뿐 순위 신호와 무관함을 명시.
    no_signal_all = len(sel_by_year) > 0 and not years_with
    if no_signal_all:
        for c in crit[2:4]:
            c["판정"] = "N/A"
            c["측정값"] = "해당 없음 — 채택 신호가 없어 순위·복합순위 자체가 만들어지지 않음"
        for c in crit[:2]:
            c["측정값"] += " [폴백 SPY vs 균등 11 비교 — 순위 신호 무관]"
        crit[4]["측정값"] += " [전 기간 폴백 = SPY M 자체 — 차이는 첫 행 처리·반올림뿐]"
    passed = all(c["판정"] == "PASS" for c in crit)
    goal_txt = (f"목표(⑤ CAGR ≥ SPY M) {'달성' if crit[4]['판정'] == 'PASS' else '미달'}: {vs_spy:+.2%}p"
                if pd.notna(vs_spy) else "목표(⑤) 계산불가")
    if no_signal_all:
        verdict = ("판정 불가(N/A) — 순환매 신호 없음: 후보 전부 워크포워드 기준 미달 → 배분은 전 기간 폴백(SPY에 E_t = M 전략). "
                   "이 상태에서는 '섹터를 고르지 않는 것'이 시스템의 판단이며, ①②는 폴백이 균등 바스켓보다 나았다는 뜻일 뿐이다. " + sel_txt)
    else:
        verdict = (("PASS — 순위 신호에 정보가 있고 목표를 달성(사전 고정 5기준 전부 충족). " if passed else
                    "FAIL — " + ", ".join(c["기준"].split(" ")[0] for c in crit if c["판정"] != "PASS")
                    + " 미달. " + goal_txt + ". 집중배분(3단계 규칙)은 사용자 요청으로 유지되나, 미달 기준이 뜻하는 바(①②④ 균등 대조군 대비 "
                    "정보 유무 / ③ 1위 판별력 / ⑤ M만 쓰는 것보다 나은가)를 13시트에서 확인할 것. ")
                   + sel_txt)
    overall = "N/A" if no_signal_all else ("PASS" if passed else "FAIL")
    crit_table = pd.concat([pd.DataFrame(crit),
                            pd.DataFrame([{"기준": "종합", "측정값": verdict, "판정": overall}])],
                           ignore_index=True)
    # [v0.5.1] 순환매 예측 판정 요약(00시트용) — 최강 후보와 그 최대 학습 t(13g), 결론 한 줄
    wf_log = (alloc.get("wf") or {}).get("selection_log", pd.DataFrame())
    best_txt = "-"
    if isinstance(wf_log, pd.DataFrame) and len(wf_log) and "NW-HAC t" in wf_log.columns:
        agg = wf_log.groupby("신호")["NW-HAC t"].agg(["max", "min", "mean"]).sort_values("max", ascending=False)
        top3 = agg.head(3)
        best_txt = "; ".join(f"{n}: 학습 t 최대 {r['max']:.2f}(최소 {r['min']:.2f}, 평균 {r['mean']:.2f})" for n, r in top3.iterrows())
    mode_txt = f"채택 모드={str(getattr(scfg, 'ROTATION_SELECT_MODE', 'strict'))}, 선택 통계={'상위1 스프레드' if stat == 'top1' else 'rank IC'}"
    basis_cnt = {}
    if isinstance(wf_log, pd.DataFrame) and "채택 근거" in wf_log.columns:
        for b in wf_log.loc[wf_log["채택"] == "채택", "채택 근거"]:
            k = str(b).split("(")[0]
            basis_cnt[k] = basis_cnt.get(k, 0) + 1
    basis_txt = ", ".join(f"{k} {v}건" for k, v in basis_cnt.items()) if basis_cnt else "없음"
    if no_signal_all:
        pred_verdict = (f"예측 못함 — 후보 {wf_log['신호'].nunique() if len(wf_log) else 0}종 전부 기준 미달({mode_txt}: "
                        f"best_available이면 t≥{getattr(scfg, 'ROTATION_SELECT_T_MIN', 1.0):.1f}도 없음) → 섹터를 고르지 않고 M 전략(SPY)을 유지. "
                        f"상위 후보: {best_txt}")
    elif passed:
        pred_verdict = f"예측함(수용기준 PASS, {goal_txt}) — {sel_txt}. 채택 근거 등급(연도×신호): {basis_txt}. {mode_txt}. 상위 후보: {best_txt}"
    else:
        pred_verdict = (f"판단은 내리지만 근거는 약함 — 신호가 채택됐으나 수용기준 미달("
                        f"{', '.join(c['기준'].split(' ')[0] for c in crit if c['판정'] not in ('PASS', 'N/A'))}); {goal_txt}. "
                        f"채택 근거 등급(연도×신호): {basis_txt}. {mode_txt}. {sel_txt}. 상위 후보: {best_txt}")
    log("ROTATION", kv(event="validation_done", stat=stat, cagr_gain_pp=round(cagr_gain * 100, 2), mdd_worse_pp=round(mdd_worse * 100, 2),
                       composite_t=round(composite_t, 2) if pd.notna(composite_t) else None,
                       composite_ic_t=round(composite_ic_t, 2) if pd.notna(composite_ic_t) else None,
                       composite_top1_t=round(composite_top1_t, 2) if pd.notna(composite_top1_t) else None, spread_pos_periods=n_pos,
                       vs_spy_m_pp=round(vs_spy * 100, 2) if pd.notna(vs_spy) else None,
                       years_with_signal=len(years_with), verdict="PASS" if passed else "FAIL",
                       elapsed_s=round(time.time() - t0, 2)),
        M=M, level="info" if passed else "warning")
    wf = alloc.get("wf") or {}
    return {"ic_table": ic_table, "spread_table": spread_table, "criteria": crit_table, "passed": passed,
            "verdict": verdict, "composite_ic_t": composite_ic_t, "composite_top1_t": composite_top1_t, "composite_t": composite_t,
            "select_stat": stat, "cagr_gain": cagr_gain, "mdd_worse": mdd_worse, "vs_spy_m": vs_spy, "goal_txt": goal_txt,
            "spread_pos_periods": n_pos, "selection_log": wf.get("selection_log", pd.DataFrame()),
            "external_log": wf.get("external_log", pd.DataFrame()), "select_mode": wf.get("mode", "strict"),
            "selected_by_year": sel_by_year, "no_signal_all": no_signal_all, "prediction_verdict": pred_verdict}


def build_allocation_sheet(alloc: Dict[str, Any], nd_spy: Optional[dict], scfg: SectorConfig) -> pd.DataFrame:
    """13c_일별배분비중: 날짜 × (E_t, 배분합계, 현금, 적격섹터수, 섹터별 배분비중·순위). 각 행의 비중은 그날 종가로
    확정돼 다음 거래일 시가에 체결할 목표비중(M 규칙 그대로). nd_spy가 있으면 맨 끝에 '예측' 행(=마지막 실적 행의
    비중을 다음 영업일 날짜로 재표기 — 새 계산 없음)을 붙인다."""
    if not alloc:
        return pd.DataFrame()
    tw, cols, idx = alloc["target_w"], alloc["cols"], alloc["target_w"].index
    out = pd.DataFrame(index=idx)
    out["날짜"] = [d.date() for d in idx]
    out["구분"] = "실적"
    out["E_t(SPY목표비중)"] = alloc["E"].round(2)
    out["배분합계"] = tw.sum(axis=1).round(4)
    out["현금"] = (1.0 - tw.sum(axis=1)).clip(lower=0.0).round(4)
    out["적격섹터수"] = alloc["n_eligible"].astype(int)
    # [v0.5.0] 3단계 판단(리더/회피/폴백/현금)·리더·꼴찌·채택 신호 수·과반 투표
    if "tier" in alloc:
        out["판단"] = alloc["tier"]
        out["1위 섹터"] = alloc["leader"]
        out["회피 섹터"] = alloc["laggard"]
        out["채택신호수"] = alloc["n_selected"]
        out["1위 득표"] = alloc["votes_leader"]
        out["꼴찌 득표"] = alloc["votes_laggard"]
    if "SPY" in tw.columns:
        out["SPY 배분비중"] = tw["SPY"].round(4)
    for t in cols:
        out[f"{t} 배분비중"] = tw[t].round(4)
    for t in cols:
        out[f"{t} 순위"] = alloc["rank_pos"][t]
    out = out.reset_index(drop=True)
    if nd_spy is not None and len(out):
        row = out.iloc[-1].copy()
        row["날짜"] = nd_spy["다음거래일"].date()
        row["구분"] = "예측"
        out = pd.concat([out, pd.DataFrame([row])], ignore_index=True)
    return out


def allocation_summary_text(alloc: Dict[str, Any], d: pd.Timestamp, top_n: int = 5) -> str:
    """그날 배분(E_t 곱한 실제 비중) 상위 top_n을 'XLK 25.0% · XLF 20.0% (합 63%, 현금 37%)' 형태로."""
    if not alloc or d not in alloc["target_w"].index:
        return ""
    w = alloc["target_w"].loc[d]
    w = w[w > 1e-9].sort_values(ascending=False)
    total = float(w.sum())
    if total <= 1e-9:
        return f"전량 현금 (E_t={float(alloc['E'].loc[d]):.2f})"
    parts = [f"{t} {v:.1%}" for t, v in w.head(top_n).items()]
    more = f" +{len(w) - top_n}개" if len(w) > top_n else ""
    return " · ".join(parts) + more + f" (합 {total:.0%}, 현금 {max(0.0, 1 - total):.0%})"


def _rule_contrib_value(rule_contrib: pd.DataFrame, label: str, col: str) -> Any:
    """[v0.3.0 §1.G-1] 09b_규칙별기여 표에서 (규칙 라벨, 열) 하나를 뽑는다 — 12_섹터요약 신규 3열이
    이 표를 유일한 소스로 삼도록(값 재계산 없이 그대로 인용)."""
    if rule_contrib is None or not len(rule_contrib) or "규칙" not in rule_contrib.columns:
        return np.nan
    hit = rule_contrib.loc[rule_contrib["규칙"] == label, col]
    return hit.iloc[0] if len(hit) else np.nan


def build_sector_summary(results: Dict[str, Dict[str, Any]], failed: Dict[str, str], universe: pd.DataFrame,
                         alloc: Optional[Dict[str, Any]] = None) -> pd.DataFrame:
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
        # [v0.4.0 §1.F] 집중 배분에서 이 섹터가 받은 평균 비중(상장 기간 기준)·그날 1위였던 날 비율·적격일 비율
        alloc_cols: Dict[str, Any] = {}
        if alloc and t in alloc.get("cols", []):
            lst = alloc["listed"][t]
            tw = alloc["target_w"][t].where(lst)
            alloc_cols = {"배분 평균비중": round(float(tw.mean()), 4) if lst.any() else np.nan,
                          "배분 1위 빈도": round(float((alloc["rank_pos"][t] == 1).where(lst).mean()), 4) if lst.any() else np.nan,
                          "배분 적격일 비율": round(float(alloc["eligible"][t].where(lst).mean()), 4) if lst.any() else np.nan}
            if "leader" in alloc:   # [v0.5.0] 실제로 '명확한 1위'로 보유된 날 비율 / '명확한 꼴찌'로 회피된 날 비율
                alloc_cols["리더 보유일 비율"] = round(float((alloc["leader"] == t).where(lst).mean()), 4) if lst.any() else np.nan
                alloc_cols["꼴찌 회피일 비율"] = round(float((alloc["laggard"] == t).where(lst).mean()), 4) if lst.any() else np.nan
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
                         rc, "[전체] 하락일 회피(비중=0 & B&H<0, %p — 음수=회피한 손실)", "비중=0인 날 B&H수익 합(%p)"),
                     **alloc_cols})
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

    alloc = sres.get("alloc") or {}
    rot_val = sres.get("rot_val") or {}
    sheets: Dict[str, pd.DataFrame] = {}
    sheets["01Z_섹터일별예측"] = sres["matrix"]
    sheets["12_섹터요약"] = summary
    sheets["13_섹터배분전략"] = sres["portfolio_perf"]
    sheets["13b_배분전략자산곡선"] = sres["portfolio_curve"]
    if alloc:   # [v0.4.0 §1.F]
        sheets["13c_일별배분비중"] = sres.get("alloc_sheet", pd.DataFrame())
        sheets["13d_횡단면IC"] = rot_val.get("ic_table", pd.DataFrame())
        sheets["13e_순위스프레드"] = rot_val.get("spread_table", pd.DataFrame())
        sheets["13f_배분수용기준"] = rot_val.get("criteria", pd.DataFrame())
        sheets["13g_순환매신호채택"] = rot_val.get("selection_log", pd.DataFrame())   # [v0.5.0] 워크포워드 연도별 채택 로그
        sheets["13h_외부검증FF49"] = rot_val.get("external_log", pd.DataFrame())     # [v0.6.0] Ken French 49업종 외부 검증
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

    # [v0.4.0 §1.F] '다음 거래일 배분' — 마지막 확정일의 집중배분 비중(= 다음 거래일 시가에 체결할 비중, 새 계산 없음)
    # + 사전 고정 수용기준 판정. 사용자가 실제로 실행할 숫자이므로 예측 블록 바로 아래에 둔다.
    if alloc:
        tw_last = alloc["target_w"].iloc[-1]
        d_last = alloc["target_w"].index[-1]
        e_last = float(alloc["E"].iloc[-1])
        total = float(tw_last.sum())
        dg = alloc["diag"]
        yr_last = int(d_last.year)
        sel_last = dg.get("selected_by_year", {}).get(yr_last, [])
        basis_last = (alloc.get("wf") or {}).get("basis_by_year", {}).get(yr_last, {})
        sel_last_txt = "+".join(f"{n}[{basis_last.get(n, '').split('(')[0]}]" for n in sel_last) if sel_last else "없음(→ 폴백)"
        stat_txt = "상위1 스프레드(1위 섹터 향후 21일 수익 − 상장 평균)" if str(dg.get("select_stat", "top1")) == "top1" else "rank IC"
        nd_rows.append(("다음 거래일 배분 - 방식", f"{dg['label_primary']} — 매년 횡단면 워크포워드 검증(선택 통계 {stat_txt}; 엄격 t ≥ {dg.get('select_t', 2.0):.1f} / "
                                             f"외부검증(FF49) 지원 / 최선 가용 t ≥ {getattr(scfg, 'ROTATION_SELECT_T_MIN', 1.0):.1f}; 모드 "
                                             f"{getattr(scfg, 'ROTATION_SELECT_MODE', 'strict')}, 지평 {dg.get('select_horizon', 21)}일, 사전방향 일치) → "
                                             f"{yr_last}년 채택: {sel_last_txt}. 규칙: ① 채택 신호 과반이 같은 섹터를 1위로 지목 → "
                                             f"그 섹터 하나에 E_t×{dg.get('leader_weight', 1.0):.0%} ② 과반이 같은 섹터를 꼴찌로 지목 → 그 섹터 뺀 적격 균등 "
                                             f"③ 둘 다 아니면 폴백 {dg.get('fallback', 'SPY')}. E_t = SPY M 목표비중({e_last:.2f}), 제외 국면 "
                                             f"{'/'.join(scfg.ROTATION_EXCLUDE_STATES)}, 리더 최소보유 {dg.get('min_hold_days', 21)}일"))
        if "tier" in alloc:
            tier_last = str(alloc["tier"].iloc[-1])
            lead_last = str(alloc["leader"].iloc[-1]) or "-"
            lag_last = str(alloc["laggard"].iloc[-1]) or "-"
            n_sel = int(alloc["n_selected"].iloc[-1])
            v_l, v_g = int(alloc["votes_leader"].iloc[-1]), int(alloc["votes_laggard"].iloc[-1])
            if tier_last == "현금":
                why = f"E_t=0(SPY M이 현금) → 섹터 무관하게 전량 현금"
            elif tier_last == "리더":
                why = f"명확한 1위 {lead_last}(채택 신호 {n_sel}개 중 {v_l}개가 1위로 지목) → {lead_last} 하나에 E_t×{dg.get('leader_weight', 1.0):.0%}"
            elif tier_last == "회피":
                why = f"1위 불명확(최다 득표 {v_l}/{n_sel}) · 명확한 꼴찌 {lag_last}({v_g}/{n_sel}) → {lag_last} 제외 적격 균등"
            else:
                why = (f"채택 신호 없음 → 폴백 {dg.get('fallback', 'SPY')}" if n_sel == 0 else
                       f"1위·꼴찌 모두 과반 미달(1위 최다 {v_l}/{n_sel}, 꼴찌 최다 {v_g}/{n_sel}) — 섹터 간 차이 없음 → 폴백 {dg.get('fallback', 'SPY')}")
            nd_rows.append(("다음 거래일 배분 - 판단", f"[{tier_last}] {why}"))
        nd_rows.append(("다음 거래일 배분 - 요약", f"{d_last.date()} 종가 확정 → 다음 거래일 시가 체결: {allocation_summary_text(alloc, d_last, top_n=12)}"))
        order = tw_last.sort_values(ascending=False).index.tolist()
        for t in order:
            w = float(tw_last[t])
            if t == "SPY":
                if w > 1e-9:
                    nd_rows.append(("다음 거래일 배분 - SPY", f"{w:.1%}  (폴백: 섹터 간 차이 없음 → 시장 전체 = M 전략)"))
                continue
            st_t = alloc["state"][t].iloc[-1]
            st_txt = STATE_SHORT.get(st_t, st_t) if isinstance(st_t, str) else "미상장"
            rp = alloc["rank_pos"][t].iloc[-1]
            comp_t = alloc["composite"][t].iloc[-1]
            rp_txt = (f"순위 {int(rp)}/{int(alloc['n_eligible'].iloc[-1])}, 복합순위점수 {comp_t:.2f}" if pd.notna(rp)
                      else ("순위 없음(제외 국면/신호 없음)" if sel_last else "순위 없음(올해 채택 신호 없음)"))
            nd_rows.append((f"다음 거래일 배분 - {t}", f"{w:.1%}  ({st_txt}, {rp_txt})"))
        nd_rows.append(("다음 거래일 배분 - 합계/현금", f"섹터+SPY 합계 {total:.1%} / 현금 {max(0.0, 1 - total):.1%} (E_t={e_last:.2f})"))
        tc = {k: dg.get(k, 0) for k in ("days_leader", "days_avoid", "days_fallback", "days_cash")}
        nd_rows.append(("집중배분 - 판단 분포(평가창)", f"리더 {tc['days_leader']}일 / 회피 {tc['days_avoid']}일 / 폴백 {tc['days_fallback']}일 / "
                                                f"현금 {tc['days_cash']}일, 리더 교체 {dg.get('leader_switches', 0)}회 — 13c '판단' 열, 13g 연도별 채택"))
        if rot_val:
            nd_rows.append(("순환매 예측 판정", rot_val.get("prediction_verdict", "-")))   # [v0.5.1]
            nd_rows.append(("집중배분 수용기준(§1.F.3, 사전 고정)", rot_val["verdict"]))
            pf = sres["portfolio_perf"].set_index("전략")

            def _pf2(name: str) -> str:
                if name in pf.index:
                    r = pf.loc[name]
                    ex = r.get("평균노출")
                    return (f"CAGR {r.get('CAGR'):.2%} / 샤프 {r.get('샤프')} / MDD {r.get('최대낙폭(MDD)'):.2%}"
                            + (f" / 평균노출 {ex}" if pd.notna(ex) else ""))
                return "-"
            alts = [l for l in (dg.get("label_leader"), dg.get("label_topk"), dg.get("label_linear"), dg.get("label_own"))
                    if l and l != dg["label_primary"]]
            # [v0.7.0] 목표 대비 — 사용자 목표(국면 비중을 따르며 상승 확률 1위 섹터 매수 → SPY 국면전략보다 높은 수익) 달성 여부를 한 줄로
            vs = rot_val.get("vs_spy_m", np.nan)
            own_txt = f" | [비교] 리더 자체 목표비중: {_pf2(dg['label_own'])}" if dg.get("label_own") else ""
            nd_rows.append(("목표 대비(⑤ CAGR ≥ SPY 국면전략 M)",
                            (f"{'달성' if (pd.notna(vs) and vs >= float(getattr(scfg, 'ROTATION_ACCEPT_VS_SPY_M', 0.0))) else '미달'} — "
                             f"주 전략 {_pf2(dg['label_primary'])} vs SPY M {_pf2('SPY 국면전략(M)')} (CAGR 차 {vs:+.2%}p)" if pd.notna(vs)
                             else "계산불가") + own_txt + " — 13f ⑤"))
            nd_rows.append(("집중배분 vs 대조군", f"[주] {dg['label_primary']}: {_pf2(dg['label_primary'])} | "
                                              + " | ".join(f"[대안] {l}: {_pf2(l)}" for l in alts) + " | "
                                              f"[대조군A] {_pf2(ROT_LABEL_CTRL_A)} | [대조군B] {_pf2(ROT_LABEL_CTRL_B)} | "
                                              f"[SPY 국면전략(M)] {_pf2('SPY 국면전략(M)')} — 13/13b~13g 시트"))
    elif scfg.USE_ROTATION:
        nd_rows.append(("다음 거래일 배분", "계산 실패 또는 섹터 결과 없음 — 로그의 SECTOR_ROTATION allocation_failed 참조"))

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
                   "다음 거래일 예측, 집중배분 합계·상위5 요약 열) / 12 섹터요약(섹터별 1행, H진입일 익일평균수익·상승 미탑승·하락 회피·"
                   "배분 평균비중·1위 빈도 포함) / "
                   "13 섹터배분전략(집중배분 주·대안 + 대조군A·B + 기존 참고 4개, 같은 평가창)·13b 자산곡선·13c 일별배분비중(판단·1위·"
                   "회피 섹터·SPY/섹터별 목표비중·순위, 맨 끝 예측 행)·13d 횡단면 rank IC(후보 사전방향·근거 포함)·13e 상위3−하위3 "
                   "스프레드·13f 수용기준 판정(①~④ 대조군A 대비, ⑤ 목표 CAGR ≥ SPY M)·13g 순환매 신호 워크포워드 연도별 채택 로그"
                   "(rank IC t·상위1 스프레드 t 병기)·13h Ken French 49업종 외부검증 / "
                   "01_일별_티커(섹터별 M 01시트와 동일 컬럼 + 위험점수백분위(H,섹터자체) 진단열, 마지막 행이 다음 거래일 예측) / "
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
        alloc_sheet = sres.get("alloc_sheet")
        if isinstance(alloc_sheet, pd.DataFrame) and len(alloc_sheet):   # [v0.4.0 §1.F]
            try:
                alloc_sheet.to_csv(scfg.ALLOC_CSV_PATH, index=False, encoding="utf-8-sig")
                log("REPORT", kv(event="alloc_csv_saved", file=scfg.ALLOC_CSV_PATH, rows=len(alloc_sheet)), M=M)
            except Exception as e:
                log("REPORT", kv(event="alloc_csv_failed", err=type(e).__name__), M=M, level="warning")
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
            if name.startswith("13c"):   # [v0.4.0] 섹터별 배분비중 데이터바 — 어느 섹터에 몰렸는지 한눈에
                for j, col in enumerate(cols):
                    if col.endswith(" 배분비중") or col in ("배분합계", "현금"):
                        w.conditional_format(1, j, len(df), j, {"type": "data_bar", "bar_color": "#638EC6", "min_type": "num",
                                                              "min_value": 0, "max_type": "num", "max_value": 1})
            if name.startswith("13f") and "판정" in cols:
                cj = cols.index("판정")
                w.conditional_format(1, cj, len(df), cj, {"type": "cell", "criteria": "==", "value": '"PASS"', "format": f_pass})
                w.conditional_format(1, cj, len(df), cj, {"type": "cell", "criteria": "==", "value": '"FAIL"', "format": f_fail})

        curve_name = "13b_배분전략자산곡선"
        if curve_name in sheets and len(sheets[curve_name]) > 0:
            d = sheets[curve_name]
            cols = list(d.columns)
            ch = wb.add_chart({"type": "line"})
            n = len(d)
            # [v0.4.0] 집중배분 전략(★ 주 전략 굵게)·대조군·기존 참고 4개를 전부 그린다 — 보조 열(날짜/E_t/섹터수)은 제외
            palette = ["#1F3864", "#2E75B6", "#7F7F7F", "#A5A5A5", "#548235", "#BF9000", "#C00000", "#7030A0", "#00B0F0"]
            skip = {"날짜", "E_t(SPY목표비중)", "적격섹터수", "신호있는섹터수"}
            series_cols = [c for c in cols if c not in skip]
            for k, cname in enumerate(series_cols):
                ci = cols.index(cname)
                width = 2.25 if "★" in str(cname) else 1.25
                ch.add_series({"name": str(cname), "categories": [curve_name, 1, 0, n, 0],
                               "values": [curve_name, 1, ci, n, ci],
                               "line": {"color": palette[k % len(palette)], "width": width}})
            ch.set_title({"name": "섹터 집중배분(★) vs 대조군 vs 균등분산 vs SPY (누적 1.0 기준, 같은 평가창)"})
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
        if scfg.EXPORT_DAILY_CSV and os.path.exists(scfg.ALLOC_CSV_PATH) and not sres.get("aborted") and sres.get("alloc"):
            M.maybe_colab_download(scfg.ALLOC_CSV_PATH)   # [v0.4.0] 13c 일별 배분비중
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
