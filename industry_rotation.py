# =============================================================================
#  industry_rotation.py
#  VERSION: v0.3.0 - 2026-09-12 - IMPROVEMENT_PLAN_INDUSTRY_v0.3 전량 구현 — 부모 계층 표본 복원(A1),
#                     다음 거래일 예측(A2), 13c 산업 전용 분리·잔여 모드(A3), ⚠ 폴백 0%·2D 격자(B1),
#                     부모국면 리더 게이트(B2), 산업 자기점수 신호(B3), 역방향 회피(B4), 추종필터(B5),
#                     베타중립 진단(B7), 13l·01Y 진단 시트(C1)
#
#  목적:
#    market_regime_trader.py(M, v1.50.0)가 SPY 국면(E_t)을, sector_rotation.py(S, v0.38.0)가
#    S★(9개 섹터 배분)를 정했다. 이 파일(I)은 S★의 섹터 비중을 "그 섹터에 속한 산업 ETF들로
#    한 단계 더 나눠 담는다" — 총 노출은 절대 바꾸지 않는다(Σ산업 + 부모ETF = S★의 그 날 그
#    섹터 비중, 14_계층정합 시트가 매일 이 등식을 검사한다).
#
#    M -> S -> I. S가 "SPY 자리에 섹터를, SPY 계층 특징에 M의 점수/위험을" 넣었듯,
#    I는 "섹터 자리에 산업을, 섹터 계층 특징에 S의 점수/위험을" 넣는다. 새 통계 기법은 없다 —
#    S/M에서 검증된 함수를 그대로 호출하고, SPY가 하드코딩된 곳만 '부모 섹터'로 일반화한다.
#
#  설계 문서: claude/INDUSTRY_LAYER_SPEC_v0.1.md (프로젝트 "미국 주식 매수, 매도 프로그램
#    만들기2"). 본 파일은 그 문서의 §1~§11을 구현한다. v0.1.x는 핵심 파이프라인만 먼저 냈고(S 자신도
#    v0.1.0에서 시작해 37개 버전에 걸쳐 자랐다), v0.2.0에서 첫 실데이터 리포트 판독으로 드러난 배분 결함을
#    고치고(헤더 CHANGELOG v0.2.0 (A)(B)(C)) 미뤘던 항목 대부분을 채웠다. **아직 v0.3 이후로 남긴 것**(조용히
#    생략하지 않고 여기 적는다 — 투명성 원칙):
#      · 06c_임계값민감도(M.threshold_sensitivity)·01Y_산업예측정확도·13d/13e/13h~13o 진단시트군·
#        13h_외부검증FF49(FF49 네트워크 요구, USE_EXTERNAL_VALIDATION=False)·"다음 거래일 예측" 행
#        (M.build_next_day_prediction은 SPY 형태의 res 번들을 요구해 산업 단위로 바로 재사용 불가).
#      · 수용기준 ③④는 상위1−하위1 스프레드로 근사(top3−bottom3 헬퍼가 S에 별도 노출되어 있지 않음).
#    룩어헤드 안전성은 11_룩어헤드감사 시트(v0.2.0 배선) + test_industry_lookahead_cut.py(절단재계산 회귀)가
#    이중으로 검증한다(사용자 표준 "룩어헤드 편향 방지" 요구 — 항상 우선).
#
#    이 파일은 market_regime_trader.py·sector_rotation.py를 런타임에 **수정하지 않는다**(호출만 한다).
#    (S.write_sector_excel(title=)·M.runtime_env()는 각자의 v0.38.0/v1.50.0에서 추가된 것을 쓴다 — 구버전이면 자동 폴백.)
#
#  CHANGELOG
#  ---------------------------------------------------------------------------
#  v0.4.0 | 2026-09-12 | [⚠ 검증 대상·사전방향·후보·격자 변경 — 리포트2(I v0.3.1, 29산업) 판독.
#    설계는 IMPROVEMENT_PLAN_S0.40_I0.4.md. 전제: sector_rotation.py v0.40.0 · M v1.51.0.
#
#    [판독 — 리더가 0일이었다] 13_산업배분전략의 ★가 S★와 모든 숫자가 같고(산업평균노출 0.0000),
#    13l 'C. 리더 에피소드: 총 0회'. v0.3.0의 §A1(표본 복원)은 성공했지만(PBETA_X_SCORE 1,938~3,357일,
#    리포트41의 0~896일에서), 13g에서 t ≥ 1.0인 신호가 한 해에 2개 이상 나오지 않아 교차확인이 불가했다.
#    프로토콜이 **옳게** 리더를 안 뽑은 것이다. 그러나 그 판정의 근거였던 통계 자체가 틀린 것을 재고 있었다.
#
#    (I1) ⚠⚠ **검증 대상 교정** — 채택 통계를 결정과 일치시킨다(신규 within_parent_walkforward_select).
#         v0.3.1까지는 S.rotation_walkforward_select를 그대로 썼는데 그 함수는 순위를 **전 산업 풀링**으로
#         매기고 상위1 스프레드를 **상장 산업 전체 평균** 대비로 잰다. 실제 결정(leader3_group)은
#         **부모 그룹 안에서** 1위를 고르고 그 대안은 **부모 ETF**다 — 둘이 다르면 검증은 결정과 다른 것을 잰다.
#         실측: P_REL_MOM_12_1의 상위1 t가 풀링에서 9개 학습창 전부 −2.4~−3.4였는데, 같은 신호를 부모
#         안에서 재면 t ≈ +0.5로 사라진다. 그 '강한 반전'은 **부모(섹터) 간** 효과였고 그건 S의 영역이다.
#         새 함수는 매일 부모마다 그 신호의 부모 안 1위를 뽑아 (그 산업 − 그 부모 ETF)의 향후 h일 수익을
#         모으고(날짜당 1관측), 학습창(cutoff 이전, 전방창도 cutoff 이내)에서 NW-HAC t를 낸다. 채택 규칙·
#         표본 게이트·등급은 S와 동일. 반환 키가 S와 같아 leader3_group은 무변경으로 받는다.
#         풀링 통계는 13g에 'NW-HAC t(상위1, 풀링·종전)'로 **진단 병기**한다.
#         ⚠ 되돌리기: IndustryConfig(ROTATION_VALIDATION_MODE="pooled").
#    (I1-4) v0.3.1 §D1의 부작용 수정 — 산업이 1개뿐인 부모(XLP=PBJ, XLRE=REZ)는 어떤 신호든 부모 안 순위가
#         상수라 '무변동'으로 집계돼 00시트에 11개 신호 전부 2,185일로 찍혔고(표시 결함), 더 나쁘게는 그
#         유일 산업을 폴백 바스켓에서도 빼 버렸다(행동 변화). 단일 산업은 여유 게이트를 구조적으로 통과할
#         수 없어 리더가 될 수 없으므로 마스킹할 이유가 없다 → 산업 2개 이상 부모에서만 판정한다.
#    (I1-5) 13l B블록 라벨 '확신 게이트 해당없음' → '채택 신호 없음(리더 불가)' — 그 해 sel_eff가 비어
#         리더 자체가 불가능했던 날들이라 '게이트 미달'과 뜻이 다르다.
#    (I1-6) ⚠⚠ **사전방향 레지스트리 신설**(industry_rotation_signal_specs) — 산업 순환매 신호 이름은
#         S.ROTATION_SIGNAL_SPECS에 **하나도 등록돼 있지 않아** 전부 기본값 +1('클수록 좋다')로 검정됐다.
#         후보지표 쪽에는 정확한 사전방향이 이미 있었는데(build_industry_candidates가 prior_sign을 넘긴다)
#         순환매 쪽만 무시한 것이다. 실제로 틀린 부호로 검정되던 것: P_REL_RSI_14(−1: 상대 과매수 반전) ·
#         P_REL_MOM_21(−1) · P_REL_EXT_200(−1) · P_REL_VOL_RATIO(−1). 하필 v0.3.1에서 **유일하게 채택된
#         신호가 P_REL_RSI_14**였다. 새 숫자를 만들지 않고 S의 _RawSpec 표를 그대로 읽어 P_ 접두로 옮긴다
#         (단일 출처). 부모 계층·자기 점수 계열만 경제적 근거와 함께 명시하며, 그중 SCORE_PCT·
#         SCORE_MINUS_PARENT는 ⚠ **S(섹터, +1)와 반대인 −1**이다 — 부모 안에서 재면 자기 점수가 가장 높은
#         산업이 다음 달 동료에 진다(부모 안 rank IC t −2.5~−2.8, h=21·63 모두 음수).
#    (I2) ⚠ 국면게이트 격자 신설 [국면게이트격자] {("NEUTRAL",), ("RISK_ON","NEUTRAL"), None} —
#         기본값 ("RISK_ON",)은 그대로 두고 격자가 판정한다. 근거: 부모 **중립**일 때 고베타 1위의 부모초과가
#         +0.81%/21일(t 1.94, 6/8년)로 가장 컸고 부모 상승(+0.25, t 0.73)·하락(−0.21)은 약했다 —
#         v0.3.0이 켠 RISK_ON은 그 반대편이다. 13l B블록의 '국면 게이트 차단일' 1위>부모 0.531·+0.29% vs
#         '해당없음' 0.471·−0.21%와도 맞는다.
#    (I3) ⚠ 격자 2종 추가 — [중립바스켓격자] {25%, 50%}(부모 안 '자기국면 중립' 산업 균등: +0.46%/21일,
#         t 2.03, 6/9년 — 기존 폴백의 '적격=상승'과 다른 집합) · [분산게이트격자] {상위50%, 상위25%}
#         (부모 안 63일 상대수익 횡단면 σ가 롤링 252일 상위 분위인 날만 리더 인정, 1일 지연 — 인과).
#    (I4) ⚠ 순위 후보 11 → 16 — 새로 계산하는 값은 PARENT_BETA_252 하나뿐이고 나머지 4개
#         (P_REL_DD_252H·P_REL_VOL_RATIO·P_REL_MA_50_200·P_REL_MA200_Z)는 **이미 후보지표로 만들어 두고도
#         순환매 후보에만 안 올렸던 것들**이다. 부모 안 실험(h=21) top1−부모ETF t: REL_NEAR_HIGH
#         (=P_REL_DD_252H) 1.36 · REL_MA200(=P_REL_MA200_Z) 0.98 · LOW_RVOL(=−P_REL_VOL_RATIO) 1.26(h=5)로
#         기존 채택 신호들보다 높았는데 워크포워드가 볼 기회조차 없었다.
#    (I5) 산업 정보를 **위로** — 산업폭(INDUSTRY_BREADTH_200)은 sector_rotation.py v0.40.0 §S5가 S의 후보로
#         받는다. 산업의 진짜 정보는 시계열 자기국면인데(29산업 중 20개가 상승 판정 +2~8%p) 그걸 부모 안
#         횡단면 선택에 쓰면 역방향이다 — 쓸 곳은 부모 섹터의 국면 예측이다.
#    (I6) INDUSTRY_LAYER_FROZEN(기본 False) — 두 라운드 연속 13f ①⑤ FAIL + 격자 전 행 불통과면 사람이
#         켜서 배분·격자·수용기준을 생략하고 산업 자기국면 시트만 낸다(실행시간 대부분 절감).
#
#    영향 함수: IndustryConfig(⚠ 필드 6개·ROTATION_SIGNALS 16) · industry_rotation_signal_specs(신규) ·
#    within_parent_walkforward_select(신규) · run_industry(rot_raw 5열) · leader3_group(단일산업 예외) ·
#    build_industry_allocation(_run_groups·_mk_target_w 인자 2개·격자 3종) · build_industry_report(13g 병기·
#    00시트 2줄) · run(검증 모드 분기·동결 스위치).
#    회귀: test_industry_within_parent.py(신규) · test_industry_* 6종 갱신.
#    ⚠ 실데이터 판정은 다음 실행의 13g(부모 안 t)·13f·13l이 낸다.
#
#  v0.3.1 | 2026-09-12 | [⚠ 결함 수정(신호 판단에 영향) — v0.3.0 회귀 검증 중 발견. 격자·채택값 무변경]
#    (D1) ⚠ '부모 안 무변동 신호'가 판단을 오염시키고 있었다. leader3_group은 신호별 풀링 순위를 부모 그룹으로
#         잘라 S._cs_rank01로 다시 0~1을 매긴다. 그런데 **부모 계층 신호(PARENT_SCORE_PCT)는 한 부모 아래
#         모든 산업이 같은 값**이라 재정규화 결과가 전원 정확히 0.5가 된다(실측 확인: 4산업 부모 → A~D 전부 0.500).
#         그대로 두면 두 가지가 조용히 망가진다:
#           (1) 복합 평균에 0.5가 섞여 실제 신호를 가운데로 **희석** → 1위−2위 여유(margin)가 줄어 리더가 덜 나온다.
#           (2) _row_arg(idxmax)가 동률 행에서 **첫 열을 임의로** 1위로 뽑아 교차확인 투표에 가짜 표를 넣는다
#               — v0.1의 '임의 리더' 결함(v0.2.0 §0(B))과 같은 종류가 부모 계층 신호 경로에 남아 있었다.
#         수정: rank_g를 만들 때 그 행의 부모 안 최대−최소 ≤ 1e-12면 **그 행만** NaN으로 비운다(행 단위 판정 —
#         값이 실제로 갈리는 날은 전혀 건드리지 않는다). _nanmean_frames는 그 행에서 신호를 무시하고, _row_arg는
#         전부 NaN이면 None을 돌려 표도 주지 않는다. 로그 `inert_signal_masked`, 00시트 '⚠ 부모 안 무변동 신호' 1줄.
#    (D2) 함께 확인한 사실 — §A1은 **PBETA_X_SCORE를 복원했고, PARENT_SCORE_PCT는 원래 채택될 수 없는 신호였다.**
#         13g의 '학습 관측일'은 상위1 스프레드 시계열의 표본 수인데, `_top1_spread_series`가 여유 게이트
#         (margin_steps=1.0)를 적용하므로 **1위가 동률인 날은 표본에서 빠진다**. 부모 계층 신호는 같은 부모의
#         산업들이 전부 동률이라 거의 모든 날이 빠진다(합성 E2E 실측: PARENT_SCORE_PCT 39일 고정 — 연도가 늘어도
#         안 늘어난다). 반면 §A1 수정 후 PBETA_X_SCORE는 1,831 → 3,128일로 커버리지가 회복됐다
#         (리포트41 실측에서는 0~896일로 전 연도 '표본부족'이었다 — 6년간 리더 0일의 원인).
#         ⚠ 따라서 PARENT_SCORE_PCT는 후보 목록에 남기되(임의 제거는 사전등록 격자 프로토콜 위반) **구조적으로
#         채택되지 않는 것이 정상**이며, 위 (D1)로 이제 '채택되더라도 해를 끼치지 않는다'. 다음 라운드에서
#         후보 목록 정리 여부를 격자로 판정할 것.
#    영향 함수: `leader3_group`(rank_g 구성 + 반환 inert_days) · `build_industry_allocation`(로그 1개 필드) ·
#    `build_industry_report`(00시트 1줄). 회귀: `test_industry_e2e_synth.py`(§A1 검사를 PBETA_X_SCORE 기준으로
#    정정 + 무변동 가드 검사) · `test_industry_reverse_avoid.py`(무변동 신호가 투표·복합평균에서 빠지는지).
#
#  v0.3.0 | 2026-09-12 | [⚠ 신호·배분 변경 — 사용자 지시 "개선사항대로 코드 수정". 근거는 전부
#    REPORT41_63_READOUT_v1.50_v0.38_I0.2.0.md(실데이터 25산업 리포트 재계산), 설계는
#    IMPROVEMENT_PLAN_INDUSTRY_v0.3.md. 실데이터 성과·수용기준 판정은 다음 실행이 낸다.]
#
#    [A. 필수 — 사용자 지시 + 결함]
#    (A1) ⚠ 부모 계층 신호 표본 복원 — v0.2까지 parent_layer_series가 S의 **마스킹된** score_pct(2018~)를
#         써서 13g의 PARENT_SCORE_PCT·PBETA_X_SCORE가 2018~2023 내내 '표본부족'(학습 관측일 0~896 < 1000)
#         이었고, 그 6년은 채택 신호가 P_REL_MOM_21 하나뿐이라 교차확인 불가 → **리더가 단 하루도 없었다**
#         (전부 폴백 균등). S 자신은 SPY 계층에 M의 마스킹 전 점수를 쓴다(spy_layer_series) — I만 잘려 있었다.
#         수정: S v0.39.0이 내보내는 score_full/haz_score_full을 M.score_percentile(expanding rank, 인과)로
#         백분위화해 쓴다. PARENT_LAYER_SOURCE 기본값 "masked_extend"→"full"(구버전 S면 자동 폴백+경고).
#         industry_lookahead_audit도 score_full을 d까지 잘라 감사가 성립하게 고쳤다.
#    (A2) 다음 거래일 예측 — 00시트에 산업별 예측 줄 + '다음 거래일 배분(I★)'(산업/잔여/부모별 판단) +
#         격자 수렴 상태, 01Z·13c에 '구분'(실적/예측) 열과 예측 행 1개. **새 계산 없음** — S.build_sector_sheets가
#         이미 만들어 둔 results[t]["sheets"]["next_day"](01_일별_<산업> 마지막 행)를 표시만 재구성(신규 industry_next_day).
#    (A3) 사용자 지시 "일별배분비중에 섹터가 왜 포함되어 있어 산업만 배분하라" — 13c는 이제 **산업 열만**
#         싣고, 부모ETF·SPY·XLU(= S★가 준 비중 중 산업으로 나누지 못한 잔여)는 13c2_잔여다리로 분리했다.
#         잔여의 목적지를 고르는 ⚠ INDUSTRY_ONLY_MODE 신설: "parent"(기본, 잔여=부모ETF — 잔여가 정확히 S★로
#         환원되어 I★−S★가 순수한 산업 판단 기여가 된다) | "industries"(잔여도 적격 산업 균등 = 문자 그대로
#         '산업만 배분') | "cash"(잔여=현금, 총노출 ≤ S★ — 14_계층정합이 '≤'로 자동 완화). 세 모드 전부
#         [잔여격자] 행으로 상설 측정한다 — ⚠ 리포트41 실측에서 잔여 전량 산업은 MDD −15.04%·칼마 2.385로
#         가장 나빴다(기본값을 "industries"로 두지 않은 이유. 켜려면 숫자를 보고 결정할 것).
#
#    [B. 정확도 개선 — 전부 사전등록 격자가 판정, 채택은 다음 라운드]
#    (B1) ⚠ INDUSTRY_FALLBACK_SHARE 기본값 0.5 → **0.0**. 리포트41 격자에서 '폴백 0%'가 CAGR 37.64%·
#         MDD −9.93%(=S★)·칼마 3.790·강건 '통과'로 ①②③④를 전부 통과한 유일한 행이었다. 원인(§3.3(3)):
#         동일가중 니치 산업 바스켓은 시총가중 부모(메가캡 주도)에 구조적으로 진다(산업↑·부모↑ 23,000일
#         표본 향후 21일 −0.17%, 승률 0.47). "확신이 있을 때만 산업, 아니면 부모."
#         격자도 1D 사다리 두 개 → **2D(cap×fb 4×4)**로 바꿨다 — (1.0, 0.0) 같은 조합이 종전엔 측정된 적이 없었다.
#         13_산업배분전략에 ①②③④ 판정 열과 '격자판정'(채택후보) 열, 00시트에 격자 수렴 1줄(M/S와 같은 형식).
#    (B2) ⚠ INDUSTRY_LEADER_REGIMES = ("RISK_ON",) 신설(기본 켬) — 그날 **부모 섹터 자기 국면**이 상승일 때만
#         리더를 인정하고, 보유 중 벗어나면 최소보유를 기다리지 않고 즉시 청산(S의 §1.D 이식).
#         근거: 리더 진입 시 부모 상승 30회 승률 0.50·+1.55% vs 중립 8회 **0.25·−0.86%** vs 하락 1회 −5.64%.
#         v0.2의 리더는 사실상 '부모 상승국면 베타 틸트'였고(PBETA_X_SCORE가 부모 상승 시 베타 순위와 같아져
#         XLI=JETS만 247일, XLK=SOXX만 292일), 횡보장(2024) 승률 0.18로 무너졌다 — 그 구간을 규칙으로 뺀다.
#    (B3) ⚠ 순위 후보에 산업 **자기** 점수 추가 — SCORE_PCT·SCORE_MINUS_PARENT(산업−부모 백분위) 및
#         P_REL_MOM_63·P_REL_RSI_14. S의 순환매에서 가장 자주 엄격 채택된 것이 '섹터 자기 SCORE_PCT'인데
#         I에는 부모 것만 있고 자기 점수가 후보에 없었다. 채택은 워크포워드가 판정(코드가 고르지 않는다).
#    (B4) ⚠ 역방향 회피(ROTATION_REVERSE_AVOID, 기본 끔) — 상위1 스프레드 NW-t ≤ −T인 신호는 "그 신호의
#         1위를 **피하라**"는 뜻이다(13g 실측: P_REL_MOM_12_1이 9개 학습창 전부 t −2.4~−3.0, 하위1은 ≈0 —
#         살 것은 없고 피할 것만 있다). 켜면 그 신호가 부모 안에서 지목한 1위를 그날 후보에서 제외한다.
#    (B5) 추종필터(INDUSTRY_LEADER_MIN_CORR, 기본 끔) — 롤링 252일(1일 지연) 산업-부모 상관이 문턱 미만인
#         산업은 **리더 후보에서만** 제외(바스켓엔 남김). 근거: 리더 상관≥0.8 17회 승률 0.53 vs <0.8 22회 0.36.
#    (B7) 베타중립 진단(ROTATION_TARGET) — 타깃을 ind−β₂₅₂·parent로 바꿔 한 번 더 워크포워드를 돌려
#         13g에 상위1 t를 나란히 싣는다. "리더 우위가 선택력인가 베타 틸트인가"를 리포트가 직접 답한다.
#
#    [C. 진단 시트 — 리포트41을 손으로 만들던 표를 리포트가 직접 낸다]
#    (C1) 13l_산업리더적중률(지평별 리더>부모 비율·타산업 기준선·폴백 바스켓 / 게이트별 / 에피소드 전수와
#         부모·연도·리더·진입국면별 승률) · 01Y_산업예측정확도(익일·지평별·상태 지속성·연도별, 전부 **기저 대비**
#         + 부모 대비 bp). 13c에 국면 게이트·역방향회피·추종필터 열 추가.
#
#    영향 함수: parent_layer_series·industry_lookahead_audit·run_industry(rot_raw)·build_pooled_rotation·
#    leader3_group·build_industry_allocation·build_hierarchy_check·build_industry_leader_columns·
#    build_industry_prediction_matrix·build_industry_report·run + 신규 industry_next_day·
#    build_industry_leader_accuracy·build_industry_prediction_accuracy·_fwd_ret + IndustryConfig 9개 필드.
#    회귀: test_industry_* 5종 + test_run_pipeline_local 갱신 후 전부 그린(합성데이터, 네트워크 불필요).
#    ⚠ 이번 라운드는 **신호·배분이 실제로 바뀐다**(A1·B1·B2·B3). 실데이터 판정은 다음 실행의 13f/13g/13l이 낸다.
#
#  v0.2.0 | 2026-09-12 | [⚠ 신호·배분 변경 — 첫 실데이터 리포트(industry_regime_report.xlsx, 19산업, 8,185초)
#    판독 결과 I★가 S★에 -2.81%p CAGR·MDD -1.24%p로 뒤졌고(수용기준 5/5 FAIL), 원인 대부분이 산업 판단이 아니라
#    배분 코드 결함이었다. 사용자 질문 "섹터 순환매 방식대로 한 거 맞아?"에 대한 정직한 답: v0.1은 아니었다.]
#    (A) 결함 1 — S★의 '산업 없는 다리' 누락: build_industry_allocation이 sres["alloc"]["target_w"]를 산업이 있는
#        8개 부모 열로만 reindex해 SPY 폴백 열·XLU 열을 버렸다. S★가 SPY 100%였던 2022-11-04~09(전섹터하락SPY) 등
#        17일과 XLU 보유 657일에 I★는 그 몫이 현금 → 15시트 2022 초과 -17.2%p, 2022-11-10 하루 -5.41%p의 주범.
#        수정: S★ target_w의 모든 열을 받아 산업이 있는 부모만 나누고 나머지(passthrough_cols)는 그대로 통과.
#        14_계층정합에 '총노출 Σ(I★)=Σ(S★)'·'S★ 재현(I 엔진, 산업 0%) = 대조군A 비트동일' 검사 행 추가, 13c에
#        통과 열 포함, 13b에 'S★ 총노출' 곡선. ret_co/ret_oc/init_exec/init_prev/rf_daily를 sres["alloc"]에서 그대로 사용.
#    (B) 결함 2 — 여유 게이트가 구조적으로 절대 통과 불가: 복합순위를 rank(pct=True)(간격 1/n)로 다시 매겨
#        1위−2위 간격이 문턱 1/(n_ok−1)보다 항상 작았다(2산업 부모: 0.5<1.0, 4산업: 0.25<0.333). 즉 v0.1 I★의
#        '리더' 배분은 한 번도 정당하게 나오지 않았고, 복합순위가 전부 NaN인 날(채택 신호 없는 해)만 margin=1.0
#        대체값으로 '통과'해 첫 열이 임의 리더가 됐다(13c 실측: 2022 SOXX 24일 vs IGV 1일). 나머지 날은 전부
#        '폴백 50% 균등' — 산업 순위가 배분에 쓰인 적이 없다.
#        수정: S.build_sector_allocation의 _run_leader3 상태기계를 부모 그룹에 그대로 이식(leader3_group, 신규):
#        신호별 부모 내 0~1 순위(S._cs_rank01 — 간격이 정확히 1/(n−1)) → sel_eff 교차확인 투표(등급별 need·과반)
#        → 복합순위 1위−2위 여유 ≥ margin_steps/(n_ok−1) → 최소보유 ROTATION_MIN_HOLD_DAYS=21(신규 필드, S와 동일)
#        상태기계(청산은 순위·투표 기준 1위 상실, S v0.10.1 규칙) → 꼴찌 회피(n_ok≥4, 회피자격 신호 투표).
#        전부 NaN인 날은 리더 없음. 검증은 풀링(§6.2), 판단은 부모 안(§7.2) — "S가 M에게 한 것"과 정확히 대응.
#        13c에 부모별 판단/리더/회피/1위 여유/여유 문턱/확신 게이트/적격산업수 열, 로그 industry_leader_applied.
#    (C) 결함 3 — 회복플로어/깊은낙폭회복/구조적저점 규칙(⑤⑦⑧) 비활성: S.run_sector 3318~3325행과 동일 호출로
#        활성화(M.deep_drawdown_flag 시그니처를 S 소스에서 확인). 산업 국면 신호의 규칙 집합이 섹터와 완전히 같아짐.
#    (D) 유니버스 19→29(사용자 지시 "산업이 더 많은데 누락 확인"): SKYY·HACK(XLK) XBI(XLV) CARZ(XLY) JETS(XLI)
#        SOCL(XLC) XOP·XES(XLE) XME·GDX(XLB). XLE·XLB 소속은 S.SECTOR_EXCLUDE가 살아 있는 동안 자동 비활성(표에는
#        남김 — 제외를 되돌리면 그대로 활성). 중복·단명 ETF는 의도적으로 제외(파일 [0] 주석에 목록). 등급B: BJK·SRVR·INDS.
#    (E) MAX_WORKERS fork 병렬(run_industries, S.run_sectors와 동일 구조·전멸 시 순차 재시도) — 기본 0(자동 min(CPU,4)).
#        실데이터 19산업 순차 8,185초 → Kaggle 4CPU 기준 약 1/3~1/4 예상. 캐시 히트 시 산업당 ~15초.
#    (F) 11_룩어헤드감사 시트: industry_lookahead_audit(v0.1.2)을 run_industry에 배선(AUDIT_SAMPLE=3/산업).
#    (G) 신규 진단: 16_산업부모추종(build_parent_following_analysis — 베타·상관·상승/하락 추종률·포착률·부모 국면별
#        초과·상승국면 월승률·국면 일치율·유형 분류) — 사용자 질문 "산업별로 섹터를 따라가는 게 있고 아닌 게 있다"에
#        대한 정량 답. 13g_산업순환매신호채택(wf selection_log). 대조군B(순위 미사용·적격 균등) 행 추가.
#    (H) 00시트 '⚠ 실매매 적용 여부' 1줄(M·S와 동일 문구), '룩어헤드 감사'·'계층정합' 요약 줄, 00시트 제목 산업용
#        (S.write_sector_excel(title=) — S v0.38.0). run() allocation 실패 시 트레이스백 꼬리를 로그에 남김.
#    영향 함수: build_industry_allocation(재작성)·leader3_group(신규)·_industry_portfolio_backtest(열별 비용)·
#    build_hierarchy_check·build_parent_following_analysis(신규)·build_industry_leader_columns(신규)·run_industry·
#    run_industries(재작성)·run·build_industry_report·IndustryConfig(ROTATION_MIN_HOLD_DAYS, MAX_WORKERS=0)·INDUSTRIES.
#    회귀: test_industry_e2e_synth.py(갱신)·test_industry_hierarchy.py(갱신, 통과 다리·S★ 재현 검사 추가)·
#    test_industry_config_guard.py(갱신)·test_industry_grid_labels.py·test_industry_lookahead_cut.py 전부 그린.
#    ⚠ 실데이터 성과·수용기준은 다음 실행(Kaggle/Colab)이 판정한다 — 이 샌드박스는 네트워크가 없다.
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
import sys
import time
import traceback
import dataclasses
import multiprocessing as mp
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd

VERSION = "v0.4.0"
VERSION_DATE = "2026-09-12"

# =============================================================================
# [0] 산업 유니버스 — INDUSTRY_LAYER_SPEC_v0.1.md §2 + v0.2.0 확장(사용자 지시 "산업이 더 많은데 누락 확인")
#     (산업티커, 부모섹터, FF49매핑열) 3튜플. 등급A(기본 활성) 29개 · 부모 10개.
#     부모가 S.SECTOR_EXCLUDE(현재 XLB·XLE)에 있으면 그 산업은 자동 비활성(유니버스 무결성 규칙) —
#     표에는 남겨 두어 제외를 되돌리면 그대로 살아난다. XLU는 산업 ETF가 없어 부모 통과 다리로만 존재.
#     ⚠ v0.2.0 추가 10개: SKYY·HACK(XLK) XBI(XLV) CARZ(XLY) JETS(XLI) SOCL(XLC) XOP·XES(XLE) XME·GDX(XLB).
#       선정 기준 = 유동성 있는 단일 산업 ETF · 2018 신호시작 전 ≥3년 이력(HACK 2014-11, JETS 2015-04는
#       경계 — SECTOR_TRAIN_MIN_YEARS=3 규칙이 신호 시작을 자동으로 늦춘다) · 기존 산업과 중복 아님.
#       의도적 제외(중복): SMH/XSD(=SOXX) XSW(=IGV) XPH(=IHE) ITB(=XHB) IAK(=KIE) IAI(=KCE) IAT(=KRE)
#       PPA/XAR(=ITA) XTN(=IYT) XTL(=IYZ) OIH(=XES). 의도적 제외(이력 짧음/청산 위험): SRVR·INDS(2018),
#       PAVE(2017), PBS(청산 위험). BJK(카지노, 2008)는 등급B.
# =============================================================================
INDUSTRIES: Tuple[Tuple[str, str, str], ...] = (
    ("SOXX", "XLK", "Chips"), ("IGV", "XLK", "Softw"), ("SKYY", "XLK", "Softw"), ("HACK", "XLK", "Softw"),
    ("IBB", "XLV", "Drugs"), ("XBI", "XLV", "Drugs"), ("IHE", "XLV", "Drugs"), ("IHI", "XLV", "MedEq"), ("IHF", "XLV", "Hlth"),
    ("XRT", "XLY", "Rtail"), ("XHB", "XLY", "Cnstr"), ("PEJ", "XLY", "Fun"), ("CARZ", "XLY", "Autos"),
    ("PBJ", "XLP", "Food"),
    ("KBE", "XLF", "Banks"), ("KRE", "XLF", "Banks"), ("KIE", "XLF", "Insur"), ("KCE", "XLF", "Fin"),
    ("ITA", "XLI", "Aero"), ("IYT", "XLI", "Trans"), ("JETS", "XLI", "Trans"),
    ("IYZ", "XLC", "Telcm"), ("FDN", "XLC", "Fun"), ("SOCL", "XLC", "Fun"),
    ("REZ", "XLRE", "RlEst"),
    ("XOP", "XLE", "Oil"), ("XES", "XLE", "Oil"),
    ("XME", "XLB", "Mines"), ("GDX", "XLB", "Gold"),
)

# ⚠ 등급B(이력 짧음/유동성 낮음/구조 변경) — 기본 꺼짐. IndustryConfig.USE_TIER_B=True로만 켠다(§15#7).
INDUSTRIES_OPTIONAL: Tuple[Tuple[str, str, str], ...] = (
    ("BJK", "XLY", "Fun"), ("SRVR", "XLRE", "RlEst"), ("INDS", "XLRE", "RlEst"),
)

INDUSTRY_NAME_KR: Dict[str, str] = {
    "SOXX": "반도체", "IGV": "소프트웨어", "SKYY": "클라우드컴퓨팅", "HACK": "사이버보안",
    "IBB": "바이오텍(대형)", "XBI": "바이오텍(동일가중·중소형)", "IHE": "제약", "IHI": "의료기기",
    "IHF": "헬스케어서비스·보험", "XRT": "소매", "XHB": "주택건설", "PEJ": "레저·엔터테인먼트", "CARZ": "자동차·미래차",
    "PBJ": "식음료", "KBE": "은행", "KRE": "지역은행", "KIE": "보험", "KCE": "자본시장",
    "ITA": "항공우주·방산", "IYT": "운송", "JETS": "항공사", "IYZ": "통신", "FDN": "인터넷·인터랙티브미디어",
    "SOCL": "소셜미디어", "REZ": "주거·헬스케어리츠",
    "XOP": "석유가스 탐사·생산", "XES": "석유가스 장비·서비스", "XME": "금속·광업", "GDX": "금광",
    "BJK": "카지노·게임", "SRVR": "데이터센터리츠", "INDS": "산업용리츠",
}

# ⚠ 상장연도는 근사치(설계서 §2). 코드는 실측 첫 유효일을 10_데이터품질에 적고,
#   INDUSTRY_TRAIN_MIN_YEARS 미만이면 신호 시작을 자동으로 늦춘다(S의 워밍업 규칙과 동일).
INDUSTRY_EXPECTED_START: Dict[str, str] = {
    "SOXX": "2001-08-01", "IGV": "2001-08-01", "SKYY": "2011-07-05", "HACK": "2014-11-11",
    "IBB": "2001-03-01", "XBI": "2006-02-01", "IHE": "2006-06-01", "IHI": "2006-06-01", "IHF": "2006-06-01",
    "XRT": "2006-06-01", "XHB": "2006-06-01", "PEJ": "2005-06-01", "CARZ": "2011-05-09",
    "PBJ": "2005-06-01", "KBE": "2005-11-01", "KRE": "2006-06-01", "KIE": "2005-11-01", "KCE": "2005-11-01",
    "ITA": "2006-05-01", "IYT": "2003-11-01", "JETS": "2015-04-30", "IYZ": "2000-05-01", "FDN": "2006-06-01",
    "SOCL": "2011-11-14", "REZ": "2007-05-01",
    "XOP": "2006-06-19", "XES": "2006-06-19", "XME": "2006-06-19", "GDX": "2006-05-16",
    "BJK": "2008-01-22", "SRVR": "2018-06-01", "INDS": "2018-06-01",
}

PARENTS: Tuple[str, ...] = tuple(sorted({p for _, p, _ in INDUSTRIES}))  # 10개(XLU 제외 — 산업 ETF 없음)


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
    USE_TIER_B: bool = False                    # ⚠ 등급B 산업(BJK·SRVR·INDS) 기본 꺼짐(§15#7)
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
    PARENT_LAYER_SOURCE: str = "full"            # ⚠ [v0.3.0 §A1] "full"(기본, S v0.39.0의 마스킹 전 점수) |
    #   "masked_extend"(종전 — 부모 신호가 2018~로 잘려 2018~2023 리더 0일이 됐던 동작, A/B 비교용)

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
    # ⚠ [v0.3.0 §B3] 순위 후보 신호. SCORE_PCT·SCORE_MINUS_PARENT 신규 — S의 순환매에서 가장 자주
    #   엄격 채택된 것이 '섹터 자기 SCORE_PCT'인데(리포트38 13g) I에는 부모 것만 있고 산업 자기 점수가
    #   후보에 없었다. P_REL_MOM_63/P_REL_RSI_14도 후보에 올린다(S의 상대강도 가족 중 I가 안 쓰던 것).
    #   채택은 워크포워드(S.rotation_walkforward_select)가 판정 — 코드가 미리 고르지 않는다.
    # ⚠ [v0.4.0 §I4] 11 → 16. 새로 계산하는 값은 PARENT_BETA_252 하나뿐이고 나머지 4개는 이미
    #   후보지표로 만들어 두고 순환매 후보에만 안 올렸던 것들이다(P_REL_DD_252H·P_REL_VOL_RATIO·
    #   P_REL_MA_50_200·P_REL_MA200_Z). 채택은 워크포워드가 판정한다 — 코드가 고르지 않는다.
    ROTATION_SIGNALS: Tuple[str, ...] = ("PARENT_SCORE_PCT", "P_REL_MOM_126", "P_REL_MOM_12_1",
                                         "P_REL_MOM_21", "P_REL_MOM_63", "P_REL_EXT_200",
                                         "P_REL_RSI_14", "RESID_MOM_12_1_PARENT",
                                         "PBETA_X_SCORE", "SCORE_PCT", "SCORE_MINUS_PARENT",
                                         "P_REL_DD_252H", "P_REL_VOL_RATIO", "P_REL_MA_50_200",
                                         "P_REL_MA200_Z", "PARENT_BETA_252")
    # ⚠ [v0.3.0 §B7] 순환매 검증 타깃. "ratio"(기본) = 산업/부모 가격비율의 수익(베타 1 가정, 종전과 비트 동일)
    #   | "beta_neutral" = ind − β₂₅₂·parent. 반대쪽은 항상 진단으로 함께 계산해 13g에 병기(ROTATION_TARGET_DIAGNOSTIC).
    ROTATION_TARGET: str = "ratio"
    ROTATION_TARGET_DIAGNOSTIC: bool = True
    # ⚠ [v0.3.0 §B4] 역방향(반전) 회피 — 상위1 스프레드 NW-t ≤ −ROTATION_SELECT_T인 신호의 '1위'를 회피 후보로.
    #   근거(리포트41 §3.3(4)): P_REL_MOM_12_1의 상위1 t가 9개 학습창 전부 −2.4~−3.0 — 살 것은 없고 피할 것만 있다.
    #   기본 꺼짐(격자 [역회피격자]가 판정).
    ROTATION_REVERSE_AVOID: bool = False
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
    ROTATION_MIN_HOLD_DAYS: int = 21               # [v0.2.0] 리더 최소보유(거래일) — S와 동일(월 리밸런스 관행)
    ROTATION_EVIDENCE_TIER: bool = True
    ROTATION_DEDUP_COMPOSITE: bool = True
    ROTATION_AVOID_VALIDATE: bool = True
    ROTATION_INCLUDE_SPY_CANDIDATE: bool = False  # [§6.1] 풀링 횡단면 = 활성 산업 전부, SPY/부모는 후보 아님
    ROTATION_IC_HORIZON: int = 21
    # ---- [v0.4.0 §I1 ⚠ 검증 대상] 채택 통계를 무엇으로 재는가 ----------------------
    # "within_parent"(기본): 부모 안 1위 − **부모 ETF** 향후수익. leader3_group의 실제 결정과 1:1.
    # "pooled": v0.3.1까지의 방식(전 산업 풀링 순위 · 상장평균 벤치) — 부모(섹터) 간 효과를 같이 잰다.
    # 근거: 같은 신호(P_REL_MOM_12_1)가 풀링에서 t −2.4~−3.4인데 부모 안에서는 t ≈ +0.5로 사라졌다
    #   → 그 '반전'은 섹터 간 효과였고 S의 영역이다(IMPROVEMENT_PLAN_S0.40_I0.4 §3.3).
    # ⚠ 되돌리기: IndustryConfig(ROTATION_VALIDATION_MODE="pooled")
    ROTATION_VALIDATION_MODE: str = "within_parent"
    # [v0.4.0 §I3] 사전등록 격자 — 배분층이라 하나의 신호를 공유한다(격자로 싣는 것이 옳다).
    #   국면게이트: 리더를 인정하는 부모 자기국면 집합. 근거(§3.4 H1): 부모 **중립**일 때 고베타 1위의
    #     21일 부모초과가 +0.81%(t 1.94, 6/8년)로 가장 컸고, 부모 상승(+0.25, t 0.73)·하락(−0.21)은 약했다.
    #     v0.3.0이 켠 ("RISK_ON",)은 그 반대편이다 — 기본값은 유지하고 격자가 판정한다.
    INDUSTRY_LEADER_REGIME_GRID: Tuple[Optional[Tuple[str, ...]], ...] = (
        ("NEUTRAL",), ("RISK_ON", "NEUTRAL"), None)
    #   중립바스켓: 부모 안에서 '자기국면 중립'인 산업 균등(§3.4 H2: +0.46%/21일, t 2.03, 6/9년).
    #     기존 폴백은 '적격(=상승)' 바스켓이라 서로 다른 집합이다.
    INDUSTRY_NEUTRAL_BASKET_GRID: Tuple[float, ...] = (0.25, 0.5)
    #   분산게이트: 부모 안 63일 상대수익 횡단면 σ가 롤링 252일 상위 q분위일 때만 리더 인정(§3.4 H3).
    INDUSTRY_DISPERSION_GATE_GRID: Tuple[float, ...] = (0.5, 0.75)
    # ---- [v0.4.0 §I6] 동결 스위치 ------------------------------------------------
    # I★가 두 라운드 연속 13f ①⑤ FAIL이고 격자 전 행이 ①②③④ 불통과면 진단 전용으로 동결한다는
    # 규칙의 코드 쪽 손잡이. True면 배분·격자·수용기준을 만들지 않고 **산업 자기국면 시트만** 낸다
    # (실행 2,420초 → 대폭 단축). 판정은 사람이 내리고, 이 플래그는 그 결정을 실행에 반영할 뿐이다.
    INDUSTRY_LAYER_FROZEN: bool = False
    USE_EXTERNAL_VALIDATION: bool = False          # v0.2 예정(§6.4) — FF49 네트워크 필요, 이번 버전은 꺼둠
    EXTERNAL_T: float = 2.0

    # ---- 산업 배분(§7) ----
    INDUSTRY_LEADER_CAP: float = 0.5              # ⚠ 리더 산업에 주는 섹터비중 몫(기본 "절반만 산업으로")
    # ⚠ [v0.3.0 §B1 기본값 변경 0.5 → 0.0] 리더가 없는 날(폴백) 적격 산업에 균등배분하는 몫.
    #   근거(리포트41 §3.1 격자 실측): 폴백 0% CAGR 37.64%·MDD −9.93%(=S★)·칼마 3.790·강건 '통과'로
    #   ①②③④를 전부 통과한 유일한 행이었다. 반대로 폴백 100%(=잔여까지 전부 산업)는 MDD −15.04%·칼마 2.385로 최악.
    #   이유(§3.3(3)): 동일가중·니치 산업 ETF 바스켓은 시총가중 부모(메가캡 주도)에 구조적으로 진다
    #   (산업↑·부모↑ 23,000일 표본에서 향후 21일 −0.17%, 승률 0.47). "확신이 있을 때만 산업, 아니면 부모".
    INDUSTRY_FALLBACK_SHARE: float = 0.0
    INDUSTRY_EXCLUDE_STATES: Tuple[str, ...] = ("RISK_OFF",)
    # ⚠ [v0.3.0 §B2] 리더 인정 국면 제약 — 그날 **부모 섹터의 자기 국면**이 이 집합에 있을 때만 리더를 인정한다
    #   (S의 ROTATION_LEADER_REGIMES를 부모 국면 기준으로 이식). None이면 제약 없음(v0.2.0 동작).
    #   근거(리포트41 §3.2): 리더 진입 시 부모가 상승국면 30회 승률 0.50·평균초과 +1.55% vs 중립 8회 0.25·−0.86%
    #   vs 하락 1회 −5.64%. 보유 중 부모 국면이 허용 밖으로 바뀌면 최소보유를 기다리지 않고 즉시 청산(S §1.D와 동일).
    INDUSTRY_LEADER_REGIMES: Optional[Tuple[str, ...]] = ("RISK_ON",)
    # ⚠ [v0.3.0 §B5] 리더 후보를 '부모를 잘 따라가는 산업'으로 제한(롤링 252일 상관 기준). None이면 제약 없음.
    #   근거(리포트41 §3.2): 리더 상관≥0.8 17회 승률 0.53 vs <0.8 22회 0.36. 기본 꺼짐 — [추종필터격자]가 판정.
    INDUSTRY_LEADER_MIN_CORR: Optional[float] = None
    INDUSTRY_FOLLOW_CORR_WINDOW: int = 252
    # ⚠ [v0.3.0 §A3 사용자 지시 "산업만 배분하라"] 부모 비중 중 산업으로 배분되지 않은 '잔여'를 어디에 두는가.
    #   "parent"(기본) = 부모 ETF — 잔여가 정확히 S★로 환원되므로 I★−S★가 순수하게 산업 판단의 기여가 된다(§7.3).
    #   "industries" = 잔여도 그 부모의 적격 산업 균등으로(= 사실상 cap·fb 1.0, 산업만 보유. ⚠ 리포트41 실측
    #                  MDD −15.04%·칼마 2.385로 가장 나쁨 — 켜기 전에 13_산업배분전략의 해당 행을 볼 것).
    #   "cash"    = 잔여를 현금으로(총노출이 S★보다 작아진다 — 14_계층정합의 '총노출=S★' 검사가 '≤'로 완화됨).
    INDUSTRY_ONLY_MODE: str = "parent"
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
    RUN_LOOKAHEAD_AUDIT: bool = True               # [v0.2.0] 산업별 절단재계산 감사(11_룩어헤드감사 시트) — industry_lookahead_audit
    AUDIT_SAMPLE: int = 3                          # 산업당 무작위 검사일 수

    # ---- 성능 ----
    USE_CACHE: bool = True
    CACHE_DIR: str = "./cache_industry"
    MAX_WORKERS: int = 0                           # [v0.2.0] 0=자동(min(CPU수,4), S와 동일 fork 병렬), 1=순차. fork 불가 환경은 자동 순차

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
def parent_layer_series(sres: dict, parent: str, M, mode: str = "full") -> Dict[str, pd.Series]:
    """[§4.3 · v0.3.0 §A1 ⚠ 신호 변경] 부모 섹터의 '계층 특징' 시리즈(SCORE_PCT·HAZ_PCT)를 만든다.

    [v0.1~v0.2의 결함 — 리포트41 §3.3(1)] sres["sectors"][parent]["score_pct"]는 S의 run_sector가
    `M.score_percentile(score).where(sig_mask)`로 만든 **리포트용 마스킹 시리즈**(그 섹터의
    SIGNAL_START=2018-01-02 이후만 값)다. 그것을 그대로 후보값으로 쓰니 산업층 워크포워드의
    커버리지 게이트(ROTATION_SELECT_MIN_DAYS=1000)에서 PARENT_SCORE_PCT·PBETA_X_SCORE가
    2018~2023 내내 '표본부족'으로 떨어졌고(13g 실측: 학습 관측일 0/34/53/72/165/317/317/320/406),
    그 6년 동안 채택 신호가 P_REL_MOM_21 하나뿐이라 교차확인(ROTATION_MIN_AGREE=2)이 불가능해
    **리더가 단 하루도 나오지 않았다**(전부 폴백 균등). S 자신은 SPY 계층에 M의 마스킹 전
    res["score"]를 쓴다(S.spy_layer_series) — I만 잘려 있었던 것이다.

    [v0.3.0 수정] mode="full"(기본): S v0.39.0이 내보내는 마스킹 전 원시 점수
    sr["score_full"]/["haz_score_full"]를 M.score_percentile로 백분위화해 쓴다. S.spy_layer_series와
    **완전히 같은 산식**이며 expanding rank라 인과적이다(그날까지의 정보만 사용 — 룩어헤드 없음,
    11_룩어헤드감사가 절단재계산으로 매 실행 검증).
    mode="masked_extend": 종전 동작(구버전 S와의 호환·A/B 비교용).
    S가 v0.39.0 미만이라 score_full이 없으면 자동으로 masked_extend로 폴백하고 경고를 남긴다."""
    sr = sres["sectors"].get(parent, {})
    out: Dict[str, pd.Series] = {}
    use_full = (mode == "full") and isinstance(sr.get("score_full"), pd.Series) and len(sr.get("score_full", [])) > 0
    if use_full:
        out["SCORE_PCT"] = M.score_percentile(sr["score_full"])
        hz = sr.get("haz_score_full")
        out["HAZ_PCT"] = M.score_percentile(hz) if isinstance(hz, pd.Series) and len(hz) else sr.get("haz_pct", pd.Series(dtype=float))
        _first = out["SCORE_PCT"].dropna()
        log("DATA", kv(event="parent_layer_ready", parent=parent, source="full(마스킹 전)",
                       first_valid=(str(_first.index[0].date()) if len(_first) else "-"), n=len(_first)), M=M)
    else:
        out["SCORE_PCT"] = sr.get("score_pct", pd.Series(dtype=float))
        out["HAZ_PCT"] = sr.get("haz_pct", pd.Series(dtype=float))
        if mode == "full":
            log("DATA", kv(event="parent_layer_source_fallback", parent=parent, requested="full",
                           used="masked_extend",
                           note="S가 v0.39.0 미만(score_full 없음) — 부모 계층 신호가 2018~로 잘려 커버리지 게이트에서 탈락할 수 있음"),
                M=M, level="warning")
        else:
            log("DATA", kv(event="parent_layer_ready", parent=parent, source="masked_extend"), M=M)
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
        # [§2b] parent_layer_series가 읽는 것은 sres["sectors"][parent]의 점수 시리즈뿐이다 —
        # S를 재실행하지 않고 그 시리즈들을 d까지 추가로 잘라 그대로 재사용한다.
        # [v0.3.0 §A1] PARENT_LAYER_SOURCE="full"이면 score_full/haz_score_full을 읽으므로 그것도 잘라야
        # 감사가 성립한다(자르지 않으면 d 이후 정보로 만든 백분위가 섞여 '불일치'가 아니라 거짓 통과가 된다).
        sr = sres["sectors"].get(parent, {})
        _cut = lambda k: (sr[k].loc[:d] if isinstance(sr.get(k), pd.Series) and len(sr.get(k, [])) else pd.Series(dtype=float))
        sres_t = {"sectors": {parent: {
            "score_pct": _cut("score_pct"), "haz_pct": _cut("haz_pct"),
            "score_full": _cut("score_full"), "haz_score_full": _cut("haz_score_full")}}}
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
    # [v0.2.0 ⚠ 신호 변경] 회복플로어(⑤)/깊은낙폭회복(⑦)/구조적저점(⑧) 규칙을 S.run_sector(3318~3325행)와
    # **동일한 호출**로 활성화 — v0.1은 M.deep_drawdown_flag의 시그니처를 재검증하지 못해 None으로 뒀었다(헤더 (b)).
    # S 소스로 확인: deep_drawdown_flag(close, dd, window, mode=...) → bool Series 1개. 산업 국면 신호가 이제
    # 섹터 국면 신호와 규칙 집합까지 완전히 같다("섹터 순환매 방식 그대로").
    recov_conf = deep_recov = struct_dd = None
    if getattr(cfg_i, "USE_RECOVERY_FLOOR", False):
        roll_low = close_i.rolling(cfg_i.RECOVERY_LOW_WINDOW, min_periods=20).min()
        recov_conf = close_i >= roll_low * (1.0 + cfg_i.RECOVERY_CONFIRM_PCT)
        struct_dd = M.deep_drawdown_flag(close_i, cfg_i.STRUCT_BOTTOM_DD, cfg_i.RECOVERY_LOW_WINDOW, mode="peak_to_trough")
        if getattr(cfg_i, "USE_DEEP_RECOVERY_BOOST", False):
            deep = M.deep_drawdown_flag(close_i, cfg_i.DEEP_RECOVERY_DD, cfg_i.RECOVERY_LOW_WINDOW, mode=cfg_i.DEEP_RECOVERY_DD_MODE)
            deep_recov = recov_conf & deep

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

    # [v0.2.0] 11_룩어헤드감사 — industry_lookahead_audit(§2b)을 파이프라인에 배선(S.run_sector와 같은 자리).
    audit = pd.DataFrame()
    if icfg.RUN_LOOKAHEAD_AUDIT and ctx.get("raw_parent") is not None and ctx.get("spy_raw_df") is not None:
        try:
            audit = industry_lookahead_audit(ind_ticker, parent, res, sres, M, icfg, S, ctx["raw"][ind_ticker],
                                             ctx["raw_parent"][parent], ctx["spy_raw_df"], W, W_haz, score, haz_score,
                                             icfg.AUDIT_SAMPLE, industry_breadth=industry_breadth,
                                             parent_breadth=parent_breadth)
        except Exception as e:
            audit = pd.DataFrame([{"티커": ind_ticker, "결과": f"감사 실패: {type(e).__name__}: {str(e)[:120]}"}])
            log("AUDIT", kv(event="lookahead_audit_failed", ticker=ind_ticker, err=str(e)[:120]), M=M, level="warning")
    timing["05b_룩어헤드감사"] = round(time.time() - t6, 2)

    sheets = S.build_sector_sheets(
        M, ind_ticker, cfg_i, specs, price_i, bt, bt_ma, sig, score, score_pct, n_used,
        haz_score, haz_pct, fast_pct, recov_conf, reason, ind_i, contrib, W, W_haz,
        wlog, val_full, adopted, trades, episodes, events, pd.DataFrame(), audit,
        haz_pct_sector=haz_pct_industry, spy_yearly_pos=None, daily_indicator_detail=False)

    # 풀링 순환매용 원자료(§6.2) — 이미 계산된 후보열에서 부모/SPY 상대 신호를 그대로 뽑아 재사용(재계산 없음).
    rot_cols = {
        "PARENT_SCORE_PCT": f"{ind_ticker}__PARENT_SCORE_PCT", "PBETA_X_SCORE": f"{ind_ticker}__PBETA_X_SCORE",
        "P_REL_MOM_126": f"{ind_ticker}__P_REL_MOM_126", "P_REL_MOM_12_1": f"{ind_ticker}__P_REL_MOM_12_1",
        "P_REL_MOM_21": f"{ind_ticker}__P_REL_MOM_21", "P_REL_EXT_200": f"{ind_ticker}__P_REL_EXT_200",
        "RESID_MOM_12_1_PARENT": f"{ind_ticker}__RESID_MOM_12_1_PARENT",
        "P_REL_MOM_63": f"{ind_ticker}__P_REL_MOM_63", "P_REL_RSI_14": f"{ind_ticker}__P_REL_RSI_14",
        # [v0.4.0 §I4] 이미 후보지표로 만들어 두고 순환매 후보에는 안 올렸던 것들 — 새로 계산하지 않는다.
        #   근거(IMPROVEMENT_PLAN_S0.40_I0.4 §3.3 부모 안 실험 h=21): 부모 안 top1−부모ETF t가
        #   REL_NEAR_HIGH(=P_REL_DD_252H) 1.36 · REL_MA200(=P_REL_MA200_Z) 0.98 · LOW_RVOL(=−P_REL_VOL_RATIO) 1.26(h=5)로
        #   기존 채택 신호들보다 높았는데 후보 목록에 없어서 워크포워드가 볼 기회조차 없었다.
        "P_REL_DD_252H": f"{ind_ticker}__P_REL_DD_252H", "P_REL_VOL_RATIO": f"{ind_ticker}__P_REL_VOL_RATIO",
        "P_REL_MA_50_200": f"{ind_ticker}__P_REL_MA_50_200", "P_REL_MA200_Z": f"{ind_ticker}__P_REL_MA200_Z",
    }
    rot_raw = pd.DataFrame(index=idx_i)
    for name, col in rot_cols.items():
        rot_raw[name] = ind_i[col] if col in ind_i.columns else np.nan
    # [v0.3.0 §B3 ⚠ 신규 순위 신호] 산업 '자기' 국면 점수 — S의 순환매에서 가장 자주 엄격 채택된 신호가
    # 섹터 자기 SCORE_PCT인데(리포트38 13g: 2019·2020·2024~26), I에는 부모 것(PARENT_SCORE_PCT)만 있고
    # 산업 자기 점수가 순위 후보에 없었다. 마스킹 전 백분위(expanding rank, 인과)를 그대로 싣는다.
    #   SCORE_MINUS_PARENT = 산업 점수 백분위 − 부모 점수 백분위 = "부모보다 자기 국면이 강한 정도".
    # 채택 여부는 워크포워드(S.rotation_walkforward_select)가 판정한다 — 코드가 미리 고르지 않는다.
    rot_raw["SCORE_PCT"] = score_pct_full.reindex(idx_i)
    rot_raw["HAZ_PCT"] = haz_pct_industry_full.reindex(idx_i)
    _p_score = parent_series.get("SCORE_PCT", pd.Series(dtype=float)).reindex(idx_i)
    rot_raw["SCORE_MINUS_PARENT"] = score_pct_full.reindex(idx_i) - _p_score
    # [§6.2] 풀링 순환매용 '부모초과수익' 일간수익률 — rel=산업/부모 가격비율(S의 상대강도 산식과
    # 동일 정의)의 일간수익률. 전체(마스킹 전) 이력 — rotation_walkforward_select가 SIGNAL_START
    # 이전 구간도 학습에 쓴다(S와 동일 관행).
    _par_tr_i = parent_tr.reindex(idx_i)
    rel_px = adj_i / _par_tr_i.replace(0, np.nan)
    rot_raw["REL_RET"] = rel_px.pct_change()
    # [v0.3.0 §B7 진단] 베타중립 초과수익 — 비율(REL_RET)은 베타 1을 가정하므로 고베타 산업이 부모 상승
    # 국면에서 구조적으로 앞선다(리포트41 §3.3(2): 리더가 사실상 '최고베타 산업'으로 수렴). 베타를 빼고도
    # 남는 선택력을 재보려고 ind − β_252·parent(β는 1일 지연 롤링, 인과)를 병기한다. ROTATION_TARGET이
    # "beta_neutral"일 때만 검증 타깃으로 쓰이고, 기본값에서는 13g 진단 열로만 나간다.
    _r_i = adj_i.pct_change()
    _r_p = _par_tr_i.pct_change()
    _beta_bn = S.rolling_beta(np.log1p(_r_i), np.log1p(_r_p), window=252, lag=1)
    rot_raw["REL_RET_BN"] = _r_i - _beta_bn * _r_p
    # [v0.4.0 §I4] 부모 대비 순수 롤링 베타 — PBETA_X_SCORE(베타×점수)는 이미 후보였지만 베타 '자체'는
    #   아니었다. §3.4 H1 실측: 부모 **중립** 국면에서 고베타 1위의 부모초과가 +0.81%/21일(t 1.94, 6/8년)로
    #   가장 컸다(부모 상승 +0.25/t 0.73 · 하락 −0.21). 1일 지연 롤링이라 인과적이며 재계산도 없다.
    rot_raw["PARENT_BETA_252"] = _beta_bn
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
        "px_close": close_i.loc[sig_mask], "ret_cc_full": ret_cc_full, "audit": audit,
    }


_CTX: Dict[str, Any] = {}


def _child_entry(ticker: str, q) -> None:
    try:
        out = run_industry(ticker, _CTX)
        q.put((ticker, "ok", out))
    except Exception as e:  # noqa
        q.put((ticker, "err", f"{type(e).__name__}: {e}\n{traceback.format_exc()[-3000:]}"))


def _resolve_workers(icfg: IndustryConfig, n_tasks: int) -> int:
    if icfg.MAX_WORKERS and icfg.MAX_WORKERS > 0:
        n = icfg.MAX_WORKERS
    else:
        n = min(os.cpu_count() or 1, 4)
    return max(1, min(n, n_tasks))


def run_industries(tickers: List[str], ctx: Dict[str, Any], icfg: IndustryConfig, M
                   ) -> Tuple[Dict[str, Dict[str, Any]], Dict[str, str]]:
    """[v0.2.0] 산업들을 순차 또는 fork 병렬로 실행 — S.run_sectors(sector_rotation.py [8])와 동일 구조
    (자식은 부모 메모리를 복사-쓰기로 공유, 결과만 큐로 반환, 100% 전멸 시 순차 1회 자동 재시도).
    v0.1은 순차만 지원해 실데이터 19개 산업이 8,185초(2.3시간) 걸렸다(리포트 00_실행요약 실측)."""
    global _CTX
    results: Dict[str, Dict[str, Any]] = {}
    failed: Dict[str, str] = {}
    n_workers = _resolve_workers(icfg, len(tickers))
    can_fork = n_workers > 1 and sys.platform.startswith("linux")
    mpctx = None
    if can_fork:
        try:
            mpctx = mp.get_context("fork")
        except ValueError:
            can_fork = False
    if not can_fork:
        log("PIPE", kv(event="sequential", n_industries=len(tickers),
                       reason=("MAX_WORKERS=1" if n_workers == 1 else "fork 불가")), M=M)
        for i, t in enumerate(tickers):
            try:
                log("PIPE", kv(event="industry_progress", done=i, total=len(tickers), ticker=t), M=M)
                results[t] = run_industry(t, ctx)
            except Exception as e:  # noqa
                failed[t] = f"{type(e).__name__}: {e}\n{traceback.format_exc()[-3000:]}"
                log("PIPE", kv(event="industry_failed", ticker=t, err=str(e)[:200]), M=M, level="error")
        return results, failed

    _CTX = ctx
    log("PIPE", kv(event="parallel_fork", n_industries=len(tickers), workers=n_workers), M=M)
    q = mpctx.Queue()
    pending = list(tickers)
    running: Dict[str, Any] = {}
    t_start: Dict[str, float] = {}
    while pending or running:
        while pending and len(running) < n_workers:
            t = pending.pop(0)
            p = mpctx.Process(target=_child_entry, args=(t, q), name=f"industry-{t}")
            p.start()
            running[t] = p
            t_start[t] = time.time()
        try:
            t, status, out = q.get(timeout=60)
        except Exception:
            for t_dead in [t for t, p in list(running.items()) if not p.is_alive() and p.exitcode not in (None, 0)]:
                p = running.pop(t_dead)
                failed[t_dead] = f"child exited with code {p.exitcode} (메모리 부족 가능 — MAX_WORKERS=1로 재시도)"
                log("PIPE", kv(ticker=t_dead, event="child_died", exitcode=p.exitcode), M=M, level="error")
            continue
        p = running.pop(t)
        p.join(timeout=30)
        if status == "ok":
            results[t] = out
            log("PIPE", kv(ticker=t, event="industry_done", elapsed_s=round(time.time() - t_start[t], 1),
                           remaining=len(pending) + len(running)), M=M)
        else:
            failed[t] = out
            log("PIPE", kv(ticker=t, event="industry_failed", msg=str(out).splitlines()[0][:200]), M=M, level="error")
    _CTX = {}
    if failed and not results and len(tickers) > 1:
        log("PIPE", kv(event="parallel_all_failed_retry_sequential", n=len(failed),
                       sample_err=next(iter(failed.values())).splitlines()[0][:200]), M=M, level="warning")
        parallel_failed = dict(failed)
        failed = {}
        for t in tickers:
            try:
                results[t] = run_industry(t, ctx)
                log("PIPE", kv(ticker=t, event="sequential_retry_recovered"), M=M)
            except Exception as e:  # noqa
                failed[t] = (f"[순차 재시도도 실패 — 산업 고유 문제일 수 있음] {type(e).__name__}: {e}\n"
                             f"{traceback.format_exc()[-3000:]}\n\n[참고: 병렬 실행 시 원본 오류]\n{parallel_failed.get(t, '')}")
                log("PIPE", kv(ticker=t, event="sequential_retry_failed", err=str(e)[:200]), M=M, level="error")
    # 결과 순서를 요청 순서(tickers)로 고정 — 병렬은 완료 순으로 들어오므로, 그대로 두면 열 순서·
    # 동률(argmax 첫 인덱스) 처리가 실행마다 달라져 배분이 재현되지 않는다(캐시 2회차 비트동일 회귀).
    results = {t: results[t] for t in tickers if t in results}
    return results, failed


# =============================================================================
# [4] 풀링 횡단면 산업 순환매 검증(§6) — S.rotation_walkforward_select를 그대로 호출.
# =============================================================================
def build_pooled_rotation(results: Dict[str, Dict[str, Any]], eval_idx: pd.DatetimeIndex,
                          icfg: IndustryConfig, M, S) -> Dict[str, Any]:
    """[§6.2] 횡단면 = 활성 산업 전부(SPY/부모는 후보 아님, ROTATION_INCLUDE_SPY_CANDIDATE=False).
    타깃(ret_cc_full) = rot_raw["REL_RET"] — '산업가격/부모가격' 비율의 일간수익률. S가 상대강도
    전반에 쓰는 것과 같은 비율기반 정의로, 그 비율의 향후 h일 수익률이 '산업 h일 수익률 − 부모 h일
    수익률'의 곱셈적(로그) 근사가 된다(§6.2 "부모 초과수익"의 구현).
    [v0.3.0 §B7] icfg.ROTATION_TARGET="beta_neutral"이면 타깃을 REL_RET_BN(ind − β₂₅₂·parent)으로 바꾼다
    (기본 "ratio"는 종전과 비트 동일). 어느 쪽이든 두 타깃 모두로 워크포워드를 돌려 13g에 나란히 싣는다 —
    "리더 우위가 선택력인가 베타 틸트인가"를 리포트가 직접 답하게 하기 위함(리포트41 §3.3(2)).
    [v0.3.0 §B4] 반환에 reverse_avoid_by_year 추가 — 상위1 스프레드 NW-t ≤ −ROTATION_SELECT_T인 신호는
    "그 신호의 1위를 **피하라**"는 뜻이다(13g 실측: P_REL_MOM_12_1이 9개 학습창 전부 t −2.4~−3.0).
    S.rotation_walkforward_select는 손대지 않고, 그 selection_log에서 뽑아 leader3_group이 회피에 쓴다."""
    cols = list(results.keys())
    full_idx = None
    for t in cols:
        idx = results[t]["ret_cc_full"].index
        full_idx = idx if full_idx is None else full_idx.union(idx)
    full_idx = full_idx.sort_values()

    def _mat(key: str) -> pd.DataFrame:
        return pd.DataFrame({t: results[t]["rot_raw"].get(key, pd.Series(dtype=float)) for t in cols}).reindex(full_idx)

    ret_ratio = _mat("REL_RET")
    ret_bn = _mat("REL_RET_BN")
    target = str(getattr(icfg, "ROTATION_TARGET", "ratio")).lower()
    if target not in ("ratio", "beta_neutral"):
        log("ROT", kv(event="unknown_rotation_target", value=target, action="ratio로 대체"), M=M, level="warning")
        target = "ratio"
    ret_cc_full = ret_bn if target == "beta_neutral" else ret_ratio
    listed_full = ret_ratio.notna()          # 상장 여부는 타깃과 무관(비율 기준) — 두 실행의 표본을 같게 유지
    ret_cc_full = ret_cc_full.where(listed_full)

    sig_full: Dict[str, pd.DataFrame] = {}
    for name in icfg.ROTATION_SIGNALS:
        mat = _mat(name)
        if mat.notna().any().any():
            sig_full[name] = mat
        else:
            log("ROT", kv(event="signal_unavailable_all_industries", signal=name), M=M, level="warning")
    if not sig_full:
        raise RuntimeError("ROTATION_SIGNALS 중 사용 가능한 신호가 없습니다 — 배분 계층을 만들 수 없음")

    half_life = icfg.ROTATION_DECAY_HALF_LIFE_DAYS
    wf = S.rotation_walkforward_select(sig_full, ret_cc_full, listed_full, eval_idx, icfg, M,
                                       external=None, half_life_days=half_life)
    wf["target"] = target

    # [§B7 진단] 반대쪽 타깃으로도 한 번 더 — 채택에는 쓰지 않고 13g 병기용(상위1 t 비교).
    wf["alt_log"] = pd.DataFrame()
    if bool(getattr(icfg, "ROTATION_TARGET_DIAGNOSTIC", True)):
        try:
            alt = S.rotation_walkforward_select(sig_full, (ret_ratio if target == "beta_neutral" else ret_bn).where(listed_full),
                                                listed_full, eval_idx, icfg, M, external=None, half_life_days=half_life)
            lg = alt.get("selection_log")
            if isinstance(lg, pd.DataFrame) and len(lg):
                keep = [c for c in ("적용연도", "신호", "NW-HAC t", "NW-HAC t(상위1)", "학습 관측일") if c in lg.columns]
                wf["alt_log"] = lg[keep].rename(columns={
                    "NW-HAC t": f"NW-HAC t({'비율' if target=='beta_neutral' else '베타중립'})",
                    "NW-HAC t(상위1)": f"NW-HAC t(상위1, {'비율' if target=='beta_neutral' else '베타중립'})",
                    "학습 관측일": f"학습 관측일({'비율' if target=='beta_neutral' else '베타중립'})"})
        except Exception as e:  # noqa
            log("ROT", kv(event="alt_target_diag_failed", err=str(e)[:150]), M=M, level="warning")

    # [§B4] 역방향(반전) 신호 — 상위1 스프레드 t ≤ −T 인 연도×신호. "그 신호의 1위를 피하라".
    rev: Dict[int, List[str]] = {}
    lg = wf.get("selection_log")
    tcut = float(icfg.ROTATION_SELECT_T)
    if isinstance(lg, pd.DataFrame) and len(lg) and "NW-HAC t(상위1)" in lg.columns:
        for _, r in lg.iterrows():
            tv = r.get("NW-HAC t(상위1)")
            nn = r.get("학습 관측일", 0)
            if pd.notna(tv) and float(tv) <= -tcut and float(nn or 0) >= icfg.ROTATION_SELECT_MIN_DAYS:
                rev.setdefault(int(r["적용연도"]), []).append(str(r["신호"]))
    wf["reverse_avoid_by_year"] = rev
    log("ROT", kv(event="pooled_rank_ready", n=len(cols), horizon=wf.get("horizon"),
                  mode=wf.get("mode"), stat=wf.get("stat"), target=target,
                  signals=len(sig_full),
                  reverse_avoid=";".join(f"{y}:{'+'.join(v)}" for y, v in sorted(rev.items())) or "-"), M=M)
    return wf


def industry_rotation_signal_specs(S) -> Dict[str, Tuple[int, str, str]]:
    """[v0.4.0 §I1-6 ⚠ 결함 수정] 산업 순환매 후보의 **사전방향 레지스트리**.

    무엇이 틀렸었나: `S.rotation_walkforward_select`와 `leader3_group`은 신호 이름을
    `S.ROTATION_SIGNAL_SPECS`에서 찾아 사전방향(prior_sign)을 곱한다. 그런데 산업 계층이 쓰는
    이름(P_REL_*·PARENT_*·PBETA_*·SCORE_MINUS_PARENT)은 그 표에 **하나도 등록돼 있지 않아**
    전부 기본값 +1("값이 클수록 좋다")로 취급됐다. 후보지표 쪽에는 정확한 사전방향이 이미 있는데도
    (build_industry_candidates가 `prior_sign=raw.prior_sign`으로 넘긴다) 순환매 쪽만 무시한 것이다.
    영향이 큰 것들: P_REL_RSI_14(진짜 −1: 상대 과매수는 되돌아온다) · P_REL_MOM_21(−1: 1개월 반전) ·
    P_REL_EXT_200(−1: 상대 과열) · P_REL_VOL_RATIO(−1: 변동성 급등은 열위).
    v0.3.1까지 **유일하게 채택된 신호가 P_REL_RSI_14**였는데, 그 채택이 경제적 사전방향과 반대
    부호로 이뤄지고 있었다 — 13g의 t 부호를 읽을 때 이 점을 몰랐으면 계속 오독했을 것이다.

    고치는 방법: 새 숫자를 만들지 않고 **S의 _RawSpec 표를 그대로 읽어** P_ 접두 이름으로 옮긴다
    (단일 출처 — S가 사전방향을 바꾸면 여기도 따라 바뀐다). 부모 계층·자기 점수 계열만 여기서
    경제적 근거와 함께 명시한다."""
    out: Dict[str, Tuple[int, str, str]] = {}
    for raw in S.relative_strength_specs():
        out[f"P_{raw.suffix}"] = (int(raw.prior_sign), f"부모 대비 {raw.name_kr}", raw.rationale)
    out.update({
        "PARENT_SCORE_PCT": (+1, "부모 섹터 복합점수 백분위",
                             "부모 국면이 강할수록 그 산업도 오른다 — 단, 한 부모 아래 모든 산업이 같은 값이라 "
                             "'부모 안 선택' 정보는 0이다(v0.3.1 §D1에서 무변동 마스킹)"),
        "PARENT_HAZ_PCT": (-1, "부모 섹터 위험점수 백분위", "부모 위험이 높을수록 산업도 약하다"),
        "PBETA_X_SCORE": (+1, "(부모대비 베타−1) × 부모 복합점수백분위",
                          "부모가 오를 국면이면 부모 대비 고베타 산업이 더 오른다(CAPM)"),
        "PARENT_BETA_252": (+1, "부모 대비 롤링 베타(252일, 1일 지연)",
                            "부모 상승·중립 국면에서 고베타 산업이 앞선다 — IMPROVEMENT_PLAN_S0.40_I0.4 §3.4 H1 "
                            "실측(부모 중립일 때 부모초과 +0.81%/21일, t 1.94, 6/8년)"),
        "SCORE_PCT": (-1, "산업 자기 복합점수 백분위(부모 안 횡단면)",
                      "⚠ 사전방향이 S(섹터, +1)와 **반대**다. 근거: 부모 안에서 재면 자기 점수가 가장 높은 산업이 "
                      "다음 달 동료에 진다(부모 안 rank IC t −2.5 ~ −2.8, h=21/63 모두 음수). 국면 점수는 "
                      "시계열로는 상승 예측력이 있지만(29산업 중 20개 +2~8%p) 산업 횡단면에서는 평균회귀 지표다"),
        "SCORE_MINUS_PARENT": (-1, "산업 점수 백분위 − 부모 점수 백분위",
                               "SCORE_PCT와 같은 이유로 부모 안에서는 평균회귀(부모 안 IC t −2.78)"),
        "RESID_MOM_12_1_PARENT": (+1, "부모 베타중립 12−1개월 잔차모멘텀",
                                  "베타를 뺀 순수 상대 모멘텀(전통 모멘텀 팩터의 산업 버전)"),
    })
    return out


def within_parent_walkforward_select(results: Dict[str, Dict[str, Any]], eval_idx: pd.DatetimeIndex,
                                     icfg: IndustryConfig, M, S, pooled: Dict[str, Any]) -> Dict[str, Any]:
    """[v0.4.0 §I1 ⚠ 검증 대상 교정] **부모 안 순위 vs 부모 ETF** 기준 워크포워드 채택.

    무엇이 틀렸었나(IMPROVEMENT_PLAN_S0.40_I0.4 §3.3·§3.5(1)):
      v0.3.1까지는 S.rotation_walkforward_select를 그대로 썼는데, 그 함수는
        (가) 순위를 **전 산업 풀링** 횡단면으로 매기고,
        (나) 상위1 스프레드를 '그날 상장 산업 전체 평균' 대비로 잰다.
      그런데 실제 결정(leader3_group)은 **부모 그룹 안에서** 1위를 고르고, 그 대안은 **부모 ETF**다.
      둘이 다르면 검증은 결정과 다른 것을 재게 된다 — 실제로 13g의 P_REL_MOM_12_1 상위1 t가
      9개 학습창 전부 −2.4~−3.4였는데, 같은 신호를 부모 안에서 재면 t ≈ +0.5로 사라졌다.
      즉 그 '강한 반전'은 **부모(섹터) 간** 효과였고, 그건 S의 영역이지 I의 영역이 아니다.

    이 함수가 재는 것(결정과 1:1):
      매일, 부모마다 → 그 신호의 부모 안 1위 산업 → 그 산업의 향후 h일 수익 − **그 부모 ETF**의 향후 h일 수익
      → 부모 평균(날짜당 1관측) → 학습창(cutoff 이전, 전방창도 cutoff 이내)에서 NW-HAC(lag h) 평균-t.
      꼴찌(하위1)도 대칭으로 재서 '회피 자격'을 준다(S와 같은 관행).
    채택 규칙은 S와 같다: 엄격(t ≥ ROTATION_SELECT_T) → 없으면 최선 가용(t ≥ ROTATION_SELECT_T_MIN, 상위 N).
    표본 게이트도 같다(ROTATION_SELECT_MIN_DAYS).

    반환 dict는 S.rotation_walkforward_select와 **같은 키 집합**이라 leader3_group이 그대로 받는다:
      rank_full(부모 안 0~1 순위) · selected_by_year · selected_eff_by_year · tier_by_year ·
      avoid_by_year · selection_log · top1_full · bottom1_full · horizon · reverse_avoid_by_year.
    pooled(기존 풀링 결과)은 13g에 진단으로 병기하기 위해 그대로 받아 둔다."""
    _SPECS = industry_rotation_signal_specs(S)      # [§I1-6] 사전방향 레지스트리(단일 출처)
    h = int(getattr(icfg, "ROTATION_SELECT_HORIZON", 21) or 21)
    t_str = float(icfg.ROTATION_SELECT_T)
    t_min = float(getattr(icfg, "ROTATION_SELECT_T_MIN", 1.0))
    best_n = int(getattr(icfg, "ROTATION_BEST_N", 2))
    min_days = int(icfg.ROTATION_SELECT_MIN_DAYS)
    need = max(int(getattr(icfg, "ROTATION_MIN_AGREE", 2)), 1)
    cols = list(results.keys())
    parent_of = {t: results[t]["parent"] for t in cols}
    groups: Dict[str, List[str]] = {}
    for t in cols:
        groups.setdefault(parent_of[t], []).append(t)
    groups = {p: v for p, v in groups.items() if len(v) >= 2}   # 산업 1개 부모는 '선택'이 성립하지 않는다

    full_idx = None
    for t in cols:
        ix = results[t]["ret_cc_full"].index
        full_idx = ix if full_idx is None else full_idx.union(ix)
    full_idx = full_idx.sort_values()

    def _mat(key: str) -> pd.DataFrame:
        return pd.DataFrame({t: results[t]["rot_raw"].get(key, pd.Series(dtype=float)) for t in cols}).reindex(full_idx)

    # 산업·부모의 '누적 총수익 지수'로 향후 h일 수익을 만든다(일간수익 → 누적 → shift; 인과).
    ind_ret = pd.DataFrame({t: results[t]["ret_cc_full"] for t in cols}).reindex(full_idx)
    rel_ret = _mat("REL_RET")
    # 부모 일간수익 복원: parent = (1+ind)/(1+rel) − 1 (rot_raw REL_RET 정의의 역산 — 재수집 없음)
    par_ret = (1.0 + ind_ret) / (1.0 + rel_ret) - 1.0
    C_i = (1.0 + ind_ret.fillna(0.0)).cumprod().where(ind_ret.notna())
    C_p = (1.0 + par_ret.fillna(0.0)).cumprod().where(par_ret.notna())
    fwd_i = C_i.shift(-h) / C_i - 1.0
    fwd_p = C_p.shift(-h) / C_p - 1.0
    exc = (fwd_i - fwd_p)                      # 산업 − 부모ETF (결정 관련량)
    listed = rel_ret.notna()

    # 부모 안 0~1 순위(leader3_group이 쓰는 것과 같은 정의 — S._cs_rank01을 부모 그룹별로)
    rank_full: Dict[str, pd.DataFrame] = {}
    top1_full: Dict[str, pd.Series] = {}
    bottom1_full: Dict[str, pd.Series] = {}
    for name in icfg.ROTATION_SIGNALS:
        mat = _mat(name)
        if not mat.notna().any().any():
            log("ROT", kv(event="signal_unavailable_all_industries", signal=name), M=M, level="warning")
            continue
        sign = _SPECS.get(name, S.ROTATION_SIGNAL_SPECS.get(name, (+1, "", "")))[0]
        x = (mat * sign).where(listed)
        rk = pd.DataFrame(np.nan, index=full_idx, columns=cols)
        tp_parts: List[pd.Series] = []
        bt_parts: List[pd.Series] = []
        for p, inds in groups.items():
            sub = S._cs_rank01(x[inds])
            rk[inds] = sub
            ok = sub.notna() & exc[inds].notna()
            n_ok = ok.sum(axis=1)
            v = sub.where(ok)
            e = exc[inds].where(ok)
            top = v.fillna(-np.inf).idxmax(axis=1).where(n_ok >= 2)
            bot = v.fillna(np.inf).idxmin(axis=1).where(n_ok >= 2)
            tp_parts.append(pd.Series([e.at[i, c] if isinstance(c, str) else np.nan for i, c in top.items()],
                                      index=full_idx))
            bt_parts.append(pd.Series([e.at[i, c] if isinstance(c, str) else np.nan for i, c in bot.items()],
                                      index=full_idx))
        rank_full[name] = rk
        top1_full[name] = pd.concat(tp_parts, axis=1).mean(axis=1) if tp_parts else pd.Series(dtype=float)
        bottom1_full[name] = pd.concat(bt_parts, axis=1).mean(axis=1) if bt_parts else pd.Series(dtype=float)
    if not rank_full:
        raise RuntimeError("ROTATION_SIGNALS 중 사용 가능한 신호가 없습니다 — 배분 계층을 만들 수 없음")

    years = sorted({int(d.year) for d in eval_idx})
    selected_by_year: Dict[int, List[str]] = {}
    selected_eff_by_year: Dict[int, List[str]] = {}
    tier_by_year: Dict[int, str] = {}
    avoid_by_year: Dict[int, List[str]] = {}
    rev_by_year: Dict[int, List[str]] = {}
    rows: List[dict] = []
    for y in years:
        cutoff = pd.Timestamp(year=y, month=1, day=1) - pd.Timedelta(days=35)
        stats: Dict[str, Tuple[float, float, int]] = {}
        stats_bot: Dict[str, Tuple[float, float, int]] = {}
        for name in rank_full:
            # 전방창이 cutoff를 넘지 않는 날만 — 룩어헤드 차단(S의 관행과 동일한 취지, 여기서는 명시적으로).
            sp = top1_full[name]
            ix = sp.index[sp.index < cutoff]
            if len(ix) > h:
                ix = ix[:-h]
            stats[name] = S._nw_mean_tstat(sp.reindex(ix).dropna(), lag=h)
            sb = bottom1_full[name]
            stats_bot[name] = S._nw_mean_tstat(sb.reindex(ix).dropna(), lag=h)
        strict = [n for n, (m, tv, nn) in stats.items()
                  if nn >= min_days and pd.notna(tv) and tv >= t_str]
        basis = {n: f"엄격(t≥{t_str:.1f})" for n in strict}
        sel = list(strict)
        if len(sel) < need:
            cands = sorted([(tv, n) for n, (m, tv, nn) in stats.items()
                            if n not in sel and nn >= min_days and pd.notna(tv) and tv >= t_min], reverse=True)
            add = [n for tv, n in cands[:max(max(best_n, need) - len(sel), 0)]]
            basis.update({n: f"최선 가용(t≥{t_min:.1f}, 상위{best_n})" for n in add})
            sel = sel + add
        tier = "엄격" if strict else ("최선 가용" if sel else "없음")
        sel_eff = list(strict) if strict else list(sel)
        usable = len(sel_eff) >= (1 if tier == "엄격" else need)
        if not usable:
            sel_eff = []
        avoid_ok = [n for n, (m, tv, nn) in stats_bot.items()
                    if n in sel_eff and nn >= min_days and pd.notna(tv) and tv <= -t_min]
        rev = [n for n, (m, tv, nn) in stats.items()
               if nn >= min_days and pd.notna(tv) and tv <= -t_str]
        selected_by_year[y] = sel
        selected_eff_by_year[y] = sel_eff
        tier_by_year[y] = tier
        avoid_by_year[y] = avoid_ok or sel_eff
        rev_by_year[y] = rev
        for name in rank_full:
            m, tv, nn = stats[name]
            mb, tb, _ = stats_bot[name]
            rows.append({"적용연도": y, "학습창 마감": str(cutoff.date()), "신호": name,
                         "사전방향": _SPECS.get(name, S.ROTATION_SIGNAL_SPECS.get(name, (0, "", "")))[0],
                         "사전방향 근거": _SPECS.get(name, (0, "", ""))[2][:120],
                         "선택 통계": "부모 안 상위1 − 부모ETF",
                         "학습 관측일": int(nn),
                         f"학습 상위1−부모(%/{h}일)": (round(m * 100, 3) if pd.notna(m) else np.nan),
                         "NW-HAC t(상위1)": (round(float(tv), 2) if pd.notna(tv) else np.nan),
                         "NW-HAC t": (round(float(tv), 2) if pd.notna(tv) else np.nan),
                         f"학습 하위1−부모(%/{h}일)": (round(mb * 100, 3) if pd.notna(mb) else np.nan),
                         "NW-HAC t(하위1)": (round(float(tb), 2) if pd.notna(tb) else np.nan),
                         "채택": ("채택" if name in sel else
                                ("표본부족" if nn < min_days else "미채택")),
                         "채택 근거": basis.get(name, ""),
                         "리더 판단 사용": ("사용" if name in sel_eff else
                                      ("미사용(교차확인 불가 — 사용 신호 %d개 < 필요 %d개)" % (len(sel), need)
                                       if (name in sel and not sel_eff) else "")),
                         "회피 자격": ("Y" if name in avoid_ok else ("N" if name in sel_eff else "")),
                         "역방향(1위 회피 후보)": ("Y" if name in rev else "")})
        log("ROT", kv(event="within_parent_select", year=y, train_end=str(cutoff.date()), tier=tier,
                      horizon=h, n_parents=len(groups),
                      selected=",".join(sel) or "-", used=",".join(sel_eff) or "-",
                      avoid_ok=",".join(avoid_ok) or "-", reverse=",".join(rev) or "-",
                      top_t=";".join(f"{n}={stats[n][1]:.2f}" for n in
                                     sorted(stats, key=lambda k: -(stats[k][1] if pd.notna(stats[k][1]) else -99))[:3])), M=M)
    out = {"rank_full": rank_full, "top1_full": top1_full, "bottom1_full": bottom1_full,
           "selected_by_year": selected_by_year, "selected_eff_by_year": selected_eff_by_year,
           "tier_by_year": tier_by_year, "avoid_by_year": avoid_by_year,
           "reverse_avoid_by_year": rev_by_year, "selection_log": pd.DataFrame(rows),
           "horizon": h, "mode": "within_parent", "stat": "top1_vs_parent",
           "smooth": int(icfg.ROTATION_SMOOTH_DAYS or 1),
           "target": "excess_vs_parent", "n_groups": len(groups),
           "pooled": pooled}
    return out


# =============================================================================
# [5] 계층 배분(§7) — build_industry_allocation. S의 build_sector_allocation과 달리 SPY가
#     하드코딩된 12번째 후보/폴백 자산 구조가 아니라, "부모 비중 안에서 산업 vs 부모ETF"를
#     결정하는 계층적 배분이라 일반화 복제가 아닌 신규 구현이다(설계서 §7.3 근거).
#
#     [v0.2.0 — 리포트(industry_regime_report.xlsx, 2026-09-11 실데이터) 검토로 드러난 v0.1 결함 3개 수정]
#     (1) S★의 '산업이 없는 다리'(SPY 폴백·XLU 등)를 통째로 버렸다 → I★가 S★보다 노출이 작았다
#         (2022-11 CPI 랠리 등 S★가 SPY 100%였던 17일에 I★는 현금 → 2022 초과수익 -17.2%p의 주범).
#         이제 S★ target_w의 **모든 열**을 받아 산업이 있는 부모만 나누고 나머지는 그대로 통과시킨다
#         (§7.1 "Σ 전체 = E_t" 정확 구현). 14_계층정합에 '총노출=S★' 행과 'S★ 재현' 행을 추가.
#     (2) 여유 게이트가 구조적으로 절대 통과할 수 없었다: 복합순위를 rank(pct=True)(=k/n 간격)로 다시
#         매겨 1위−2위 간격이 항상 1/n < 문턱 1/(n−1)이었다. 즉 '리더' 배분은 한 번도 정당하게 나오지
#         않았고, 복합순위가 전부 NaN인 날(채택 신호 없는 해)만 margin=1.0 대체값으로 '통과'해 첫 열
#         (SOXX 등)이 임의 리더가 됐다(2022 SOXX 24일 vs IGV 1일). → S의 leader3 상태기계(_run_leader3)를
#         부모 그룹에 **그대로 이식**: 신호별 부모 내 0~1 순위(S._cs_rank01) → 교차확인 투표(sel_eff,
#         등급별 need·과반) → 복합순위 1위−2위 여유 ≥ margin_steps/(n_ok−1) → 최소보유 21일 상태기계 →
#         꼴찌 회피(n_ok≥4). 전부 NaN인 날은 리더 없음(폴백/부모ETF).
#     (3) 검증(풀링)과 판단(부모 안)의 분리를 명시: 워크포워드 채택·통계는 풀링 횡단면(§6.2), 일별
#         리더 판단은 부모 그룹 안 순위(§7.2) — "S가 M에게 한 것(시장 안 섹터 순위)을 I가 S에게(섹터 안
#         산업 순위)"와 정확히 대응한다.
# =============================================================================
def _row_arg(r: pd.DataFrame, fn: str) -> pd.Series:
    """행별 idxmax/idxmin — 전부 NaN인 행은 None(pandas idxmax는 all-NA 행에서 예외). S와 동일."""
    out = pd.Series([None] * len(r), index=r.index, dtype=object)
    has = r.notna().any(axis=1)
    if has.any():
        out[has] = getattr(r[has], fn)(axis=1).astype(object)
    return out


def leader3_group(parent: str, inds: List[str], eval_idx: pd.DatetimeIndex, rank_full: Dict[str, pd.DataFrame],
                  wf: Dict[str, Any], eligible: pd.DataFrame, listed: pd.DataFrame,
                  icfg: IndustryConfig, S, parent_state: Optional[pd.Series] = None,
                  follow_corr: Optional[pd.DataFrame] = None) -> Dict[str, Any]:
    """[§7.2 · v0.2.0] S.build_sector_allocation의 _run_leader3 상태기계를 부모 그룹(inds)에 그대로 적용.
    입력 rank_full = 풀링 워크포워드(wf["rank_full"], 신호별 풀링 0~1 순위) — 부모 그룹으로 잘라
    S._cs_rank01로 다시 0~1 정규화하면 인접 순위 간격이 정확히 1/(n−1)이 되어 S의 여유 게이트
    (step = margin_steps/(n_ok−1))와 같은 척도가 된다. 반환:
      leader_ind(date×inds, 그날 보유 리더 1.0 원핫), basket_ind(date×inds, 폴백/회피 바스켓 균등 — 합 1.0),
      tier/leader/laggard/votes_leader/votes_laggard/margin/step/gate/n_ok/composite, switches.
    비중은 호출부가 leader_ind×INDUSTRY_LEADER_CAP + basket_ind×INDUSTRY_FALLBACK_SHARE로 만든다
    (판단은 캡/폴백 비율과 무관 — 격자 변형이 같은 판단을 공유한다).

    [v0.3.0 ⚠ 신호 변경 3개 — 전부 리포트41 §3.2/§3.3 실측 근거]
    (§B2 국면 게이트) parent_state가 주어지고 icfg.INDUSTRY_LEADER_REGIMES가 None이 아니면, 그날 **부모
        섹터의 자기 국면**이 그 집합에 있을 때만 리더를 인정한다(S의 _run_leader3 regime_ok 이식). 보유 중
        국면이 허용 밖으로 바뀌면 최소보유를 기다리지 않고 즉시 청산(S §1.D와 동일 — 위험 축소가 목적).
        근거: 리더 진입 시 부모 상승 30회 승률 0.50·+1.55% vs 중립 8회 0.25·−0.86% vs 하락 1회 −5.64%.
    (§B4 역방향 회피) wf["reverse_avoid_by_year"]의 신호(상위1 스프레드 t ≤ −T — "그 신호의 1위를 피하라")가
        부모 안에서 지목한 1위는 그날 후보에서 통째로 제외한다(리더·바스켓 모두). 반전 신호는 이미 엄격
        문턱의 반대편을 통과했으므로 단독으로 회피 자격이 있다(S가 엄격 신호의 단독 리더를 허용하는 것과
        같은 논리). icfg.ROTATION_REVERSE_AVOID=False(기본)면 이 블록은 완전히 비활성.
    (§B5 추종필터) follow_corr(롤링 252일 산업-부모 상관, 1일 지연 — 인과)가 주어지고
        icfg.INDUSTRY_LEADER_MIN_CORR이 설정되면 상관이 문턱 미만인 산업은 **리더 후보에서만** 뺀다
        (바스켓에는 남긴다). 근거: 리더 상관≥0.8 17회 승률 0.53 vs <0.8 22회 0.36."""
    smooth = int(icfg.ROTATION_SMOOTH_DAYS or 1)
    min_agree = max(1, int(icfg.ROTATION_MIN_AGREE))
    margin_steps = float(icfg.ROTATION_LEADER_MARGIN_STEPS or 0.0)
    min_hold = int(getattr(icfg, "ROTATION_MIN_HOLD_DAYS", 21))
    sel_by_year = wf.get("selected_by_year", {})
    sel_eff_by_year = wf.get("selected_eff_by_year", sel_by_year)
    tier_by_year = wf.get("tier_by_year", {})
    avoid_by_year = wf.get("avoid_by_year", sel_eff_by_year)
    lst = listed.reindex(index=eval_idx, columns=inds).fillna(False).astype(bool)
    elg = eligible.reindex(index=eval_idx, columns=inds).fillna(False).astype(bool)
    # [§B2] 부모 국면 배열 — 없으면 제약 없음(v0.2.0 동작).
    leader_regimes = getattr(icfg, "INDUSTRY_LEADER_REGIMES", None)
    leader_regimes = tuple(leader_regimes) if leader_regimes else None
    if parent_state is not None:
        p_state_arr = parent_state.reindex(eval_idx).astype(object).where(parent_state.reindex(eval_idx).notna(), "-").values
    else:
        p_state_arr = np.array(["-"] * len(eval_idx), dtype=object)
        if leader_regimes:
            log("ROT", kv(event="leader_regime_unavailable", parent=parent, action="국면 제약 무시"), M=None, level="warning")
            leader_regimes = None
    # [§B4] 역방향 회피 연도별 신호
    use_reverse = bool(getattr(icfg, "ROTATION_REVERSE_AVOID", False))
    rev_by_year = wf.get("reverse_avoid_by_year", {}) if use_reverse else {}
    # [§B5] 추종(상관) 필터
    min_corr = getattr(icfg, "INDUSTRY_LEADER_MIN_CORR", None)
    corr_g = follow_corr.reindex(index=eval_idx, columns=inds) if (follow_corr is not None and min_corr is not None) else None

    # 신호별 '부모 안' 0~1 순위 — 풀링 순위를 잘라 다시 정규화(단조 변환이라 원시값 순위와 동일)
    rank_g: Dict[str, pd.DataFrame] = {}
    # [v0.4.0 §I1-4] 산업이 1개뿐인 부모(XLP=PBJ, XLRE=REZ)는 어떤 신호든 부모 안 순위가 상수다.
    #   v0.3.1은 그것까지 '무변동'으로 세어 00시트에 11개 신호 전부 2,185일로 찍혔고(표시 결함),
    #   더 나쁘게는 그 부모의 유일 산업을 폴백 바스켓에서도 빼 버렸다(행동 변화). 단일 산업은 여유
    #   게이트(margin ≥ step)를 구조적으로 통과할 수 없어 리더가 될 수 없으므로 마스킹할 이유가 없다.
    #   → 산업 2개 이상인 부모에서만 무변동 판정을 한다.
    _single = len(inds) < 2
    inert_days: Dict[str, int] = {}
    for name, rk in rank_full.items():
        r = rk.reindex(index=eval_idx, columns=inds).where(lst)
        g = S._cs_rank01(r)
        # [v0.3.1 ⚠ 결함 수정] '부모 안에서 값이 전부 같은 날'은 그 신호가 이 부모의 산업을 고를 정보가 0인 날이다.
        #   대표 사례가 부모 계층 신호(PARENT_SCORE_PCT — 한 부모 아래 모든 산업이 같은 값). _cs_rank01은 그런 행에
        #   전원 0.5를 준다. 그대로 두면 두 가지가 조용히 망가진다:
        #     (1) 복합 평균에 0.5가 섞여 실제 신호를 가운데로 **희석** → 1위−2위 여유(margin)가 줄어 리더가 덜 나온다.
        #     (2) _row_arg(idxmax)가 동률 행에서 **첫 열을 임의로** 1위로 뽑아 교차확인 투표에 가짜 표를 넣는다
        #         (v0.1의 '임의 리더' 결함과 같은 종류 — v0.2.0 §0(B)에서 고친 것이 여기 남아 있었다).
        #   해당 행만 NaN으로 비운다: _nanmean_frames는 그 행에서 이 신호를 무시하고, _row_arg는 전부 NaN이면
        #   None을 돌려주어 표도 안 준다. 값이 실제로 변하는 날은 전혀 건드리지 않는다(행 단위 판정).
        flat = (g.max(axis=1) - g.min(axis=1)).fillna(0.0) <= 1e-12
        n_flat = 0 if _single else int((flat & g.notna().any(axis=1)).sum())
        if n_flat:
            g = g.mask(flat)
            inert_days[name] = n_flat
        rank_g[name] = g
    if inert_days:
        log("ROT", kv(event="inert_signal_masked", parent=parent, n_days=len(eval_idx),
                      detail=";".join(f"{k}={v}" for k, v in sorted(inert_days.items(), key=lambda x: -x[1])),
                      note="부모 안 무변동일 — 복합평균·투표에서 제외(희석·임의1위 방지)"), M=None)
    # 연도별 복합순위 = 그 해 리더 판단에 쓰는 신호(sel_eff)의 부모 안 순위 평균 → 평활 → 적격 마스킹(S와 동일 순서)
    comp = pd.DataFrame(np.nan, index=eval_idx, columns=inds)
    for y in sorted(set(eval_idx.year)):
        sel = [s for s in sel_eff_by_year.get(y, []) if s in rank_g]
        rows_y = eval_idx[eval_idx.year == y]
        if not sel or len(rows_y) == 0:
            continue
        comp.loc[rows_y] = S._nanmean_frames([rank_g[s] for s in sel], rows_y, inds).values
    comp_s = comp.rolling(smooth, min_periods=1).mean().where(lst) if smooth > 1 else comp
    composite = comp_s.where(elg)
    rank_sel_s: Dict[str, pd.DataFrame] = {}
    for name, r in rank_g.items():
        rs = r.rolling(smooth, min_periods=1).mean().where(lst) if smooth > 1 else r
        rank_sel_s[name] = rs.where(elg)
    am = {n: _row_arg(r, "idxmax") for n, r in rank_sel_s.items()}
    an = {n: _row_arg(r, "idxmin") for n, r in rank_sel_s.items()}

    comp_vals = composite.values
    elig_vals = elg.values
    col_arr = np.array(inds)
    idx_of = {c: j for j, c in enumerate(inds)}
    years_arr = eval_idx.year
    n = len(eval_idx)
    leader_ind = pd.DataFrame(0.0, index=eval_idx, columns=inds)
    basket_ind = pd.DataFrame(0.0, index=eval_idx, columns=inds)
    tier_ = pd.Series("부모ETF", index=eval_idx, dtype=object)
    leader_ = pd.Series("", index=eval_idx, dtype=object)
    laggard_ = pd.Series("", index=eval_idx, dtype=object)
    v_lead_ = pd.Series(0, index=eval_idx, dtype=int)
    v_lag_ = pd.Series(0, index=eval_idx, dtype=int)
    margin_ = pd.Series(np.nan, index=eval_idx, dtype=float)
    step_ = pd.Series(np.nan, index=eval_idx, dtype=float)
    gate_ = pd.Series("해당없음", index=eval_idx, dtype=object)
    regime_gate_ = pd.Series("해당없음", index=eval_idx, dtype=object)   # [§B2] 통과/차단/해당없음
    rev_avoid_ = pd.Series("", index=eval_idx, dtype=object)             # [§B4] 그날 역방향 회피된 산업
    corr_block_ = pd.Series("", index=eval_idx, dtype=object)            # [§B5] 상관 문턱 미달로 리더에서 빠진 산업
    n_ok_ = pd.Series(0, index=eval_idx, dtype=int)
    cur_leader: Optional[str] = None
    held = 0
    switches = 0
    for i in range(n):
        yr = int(years_arr[i])
        sel = [s for s in sel_eff_by_year.get(yr, []) if s in rank_g]
        K = len(sel)
        need = 1 if tier_by_year.get(yr) == "엄격" else min_agree
        avoid_ok = [s for s in avoid_by_year.get(yr, sel) if s in rank_g]
        rev_sel = [s for s in rev_by_year.get(yr, []) if s in rank_g]
        row = comp_vals[i]
        ok = (elig_vals[i] & ~np.isnan(row)).copy()
        # [§B4] 역방향 회피 — 반전 신호(상위1 t ≤ −T)가 부모 안에서 지목한 1위를 그날 후보에서 제외.
        if rev_sel:
            _rev_hit = []
            for sname in rev_sel:
                c = am[sname].iloc[i]
                if c is not None and c in idx_of and ok[idx_of[c]]:
                    ok[idx_of[c]] = False
                    _rev_hit.append(c)
            if _rev_hit:
                rev_avoid_.iloc[i] = ",".join(sorted(set(_rev_hit)))
        n_ok = int(ok.sum())
        n_ok_.iloc[i] = n_ok
        # [§B5] 추종(상관) 필터 — 리더 후보 집합만 좁힌다(바스켓은 ok 그대로).
        ok_lead = ok
        if corr_g is not None and n_ok >= 1:
            cvals = corr_g.iloc[i].values.astype(float)
            bad = ok & ~(cvals >= float(min_corr))
            if bad.any():
                ok_lead = ok & ~bad
                corr_block_.iloc[i] = ",".join(col_arr[bad].tolist())
        n_lead_ok = int(ok_lead.sum())
        leader = laggard = None
        v_lead = v_lag = 0
        margin_i = step_i = np.nan
        gate_ok = False
        if n_lead_ok >= 1 and K > 0:
            j1 = int(np.nanargmax(np.where(ok_lead, row, -np.inf)))
            leader = col_arr[j1]
            v_lead = sum(1 for s in sel if s in am and am[s].iloc[i] == leader)
            others = np.where(ok_lead, row, -np.inf).astype(float)
            others[j1] = -np.inf
            second = float(others.max()) if n_lead_ok >= 2 else np.nan
            margin_i = float(row[j1] - second) if np.isfinite(second) else np.nan
            step_i = margin_steps / max(n_lead_ok - 1, 1)
            gate_ok = (margin_steps <= 0) or (np.isfinite(margin_i) and margin_i >= step_i - 1e-9)
        if n_ok >= 4 and K > 0:
            laggard = col_arr[int(np.nanargmin(np.where(ok, row, np.inf)))]
            v_lag = sum(1 for s in avoid_ok if s in an and an[s].iloc[i] == laggard)
        margin_.iloc[i] = margin_i
        step_.iloc[i] = step_i
        lead_ok_votes = leader is not None and v_lead * 2 > K and v_lead >= need
        # [§B2] 국면 게이트 — 부모 자기 국면이 허용 집합에 있을 때만 리더 인정.
        regime_ok = (leader_regimes is None) or (p_state_arr[i] in leader_regimes)
        if leader_regimes is None:
            regime_gate_.iloc[i] = "해당없음"
        else:
            regime_gate_.iloc[i] = "통과" if regime_ok else f"차단({p_state_arr[i]})"
        clear_leader = lead_ok_votes and gate_ok and regime_ok
        gate_blocked = lead_ok_votes and not gate_ok
        gate_.iloc[i] = "통과" if (lead_ok_votes and gate_ok) else ("미달" if gate_blocked else "해당없음")
        leader_lost = not (lead_ok_votes and leader == cur_leader)
        clear_laggard = laggard is not None and v_lag * 2 > K and v_lag >= need
        v_lead_.iloc[i], v_lag_.iloc[i] = v_lead, v_lag
        # 최소보유 상태기계(S v0.10.1 청산 규칙 분리 그대로: 게이트는 진입 전용, 청산은 순위·투표 기준 1위 상실)
        if cur_leader is not None and not ok[idx_of[cur_leader]]:
            cur_leader = None
        # [§B2] 국면 제약은 '적격 상실'과 같은 즉시청산(S §1.D) — min_hold를 기다리지 않는다.
        if cur_leader is not None and leader_regimes is not None and p_state_arr[i] not in leader_regimes:
            cur_leader = None
        if clear_leader and leader != cur_leader:
            if cur_leader is None or held >= min_hold:
                cur_leader = leader
                held = 0
                switches += 1
        elif leader_lost and cur_leader is not None and held >= min_hold:
            cur_leader = None
        if cur_leader is not None:
            leader_ind.iat[i, idx_of[cur_leader]] = 1.0
            tier_.iloc[i] = "리더"
            leader_.iloc[i] = cur_leader
            held += 1
        elif clear_laggard:
            basket = [c for j, c in enumerate(inds) if ok[j] and c != laggard]
            if basket:
                for c in basket:
                    basket_ind.iat[i, idx_of[c]] = 1.0 / len(basket)
                tier_.iloc[i] = "회피"
                laggard_.iloc[i] = laggard
            else:
                tier_.iloc[i] = "부모ETF"
        elif n_ok >= 1:
            for j, c in enumerate(inds):
                if ok[j]:
                    basket_ind.iat[i, j] = 1.0 / n_ok
            tier_.iloc[i] = "폴백(여유부족)" if gate_blocked else "폴백"
        else:
            tier_.iloc[i] = "부모ETF"
    return {"parent": parent, "inds": inds, "leader_ind": leader_ind, "basket_ind": basket_ind,
            "tier": tier_, "leader": leader_, "laggard": laggard_, "votes_leader": v_lead_, "votes_laggard": v_lag_,
            "margin": margin_, "step": step_, "gate": gate_, "n_ok": n_ok_, "composite": composite,
            "regime_gate": regime_gate_, "rev_avoid": rev_avoid_, "corr_block": corr_block_,   # [v0.3.0 §B2/§B4/§B5]
            "parent_state": pd.Series(p_state_arr, index=eval_idx),
            "inert_days": inert_days,                                  # [v0.3.1] 부모 안 무변동일 수(신호별)
            "switches": switches}


def _industry_portfolio_backtest(S, target_w: pd.DataFrame, ret_co: pd.DataFrame, ret_oc: pd.DataFrame,
                                 rf_daily: Optional[pd.Series], cost_bps_by_col: Dict[str, float],
                                 init_exec: Optional[pd.Series] = None, init_prev: Optional[pd.Series] = None
                                 ) -> pd.DataFrame:
    """[§7.5] S.portfolio_backtest는 비용을 단일 스칼라(cost_bps)로만 받는다. 산업(10bp)·부모/통과 다리(5bp)
    비용률을 열별로 반영하기 위해 cost_bps=0으로 총수익(비용 미차감)을 얻은 뒤, 열별 회전율에 각자의
    비용률을 곱해 직접 차감한다(이중차감 없음 — gross 쪽 cost 열은 항상 0). exec/prev 비중 정의는
    S.portfolio_backtest와 동일(exec_w(t)=target_w(t−1), init_exec/init_prev 지원)."""
    gross = S.portfolio_backtest(target_w, ret_co, ret_oc, rf_daily, cost_bps=0.0,
                                 init_exec=init_exec, init_prev=init_prev)
    cols = list(target_w.columns)
    exec_w = target_w.shift(1)
    if init_exec is not None and len(target_w.index):
        exec_w.iloc[0] = init_exec.reindex(cols).fillna(0.0).values
    exec_w = exec_w.fillna(0.0)
    prev_w = exec_w.shift(1)
    if init_prev is not None and len(target_w.index):
        prev_w.iloc[0] = init_prev.reindex(cols).fillna(0.0).values
    prev_w = prev_w.fillna(0.0)
    turn = (exec_w - prev_w).abs()
    bps = pd.Series({c: float(cost_bps_by_col.get(c, 0.0)) for c in cols})
    cost = (turn * (bps / 1e4)).sum(axis=1)
    out = gross.copy()
    out["turnover"] = turn.sum(axis=1)
    out["cost"] = cost
    out["strategy_ret"] = (gross["strategy_ret"] - cost).fillna(0.0)
    out["equity"] = (1.0 + out["strategy_ret"]).cumprod()
    out["dd"] = out["equity"] / out["equity"].cummax() - 1.0
    return out


def build_industry_allocation(results: Dict[str, Dict[str, Any]], sres: dict, res: dict,
                              eval_idx: pd.DatetimeIndex, icfg: IndustryConfig, M, S,
                              wf: Dict[str, Any], rf_daily: Optional[pd.Series] = None
                              ) -> Dict[str, Any]:
    """[§7 · v0.2.0] S★ target_w(모든 열: 9섹터 + SPY)를 받아 '산업이 있는 부모'만 부모 안에서 나누고
    나머지 열(SPY 폴백·XLU 등)은 그대로 통과 → 총노출 = S★ 총노출(매일, 14_계층정합). 반환 dict 키는
    v0.1과 호환(target_w/target_ws/bts/perf/curve/cols/active_parents/parent_of/all_cols/eligible/state/
    w_s/label_star/label_ctrl_a/ret_co/ret_oc/rf_daily …) + passthrough_cols/groups/w_s_all/label_repro."""
    t0 = time.time()
    cols = list(results.keys())
    if not cols:
        return {}
    s_alloc = sres.get("alloc", {}) or {}
    if not s_alloc or "target_w" not in s_alloc:
        log("ROTATION", kv(event="s_alloc_missing", note="sres['alloc']['target_w'] 없음 — S가 배분을 만들지 않았다"),
            M=M, level="error")
        return {}
    parent_of = {t: results[t]["parent"] for t in cols}
    s_all_cols = [c for c in (s_alloc.get("all_cols") or list(s_alloc["target_w"].columns))]
    active_parents = sorted({parent_of[t] for t in cols if parent_of[t] in s_all_cols})
    dropped = sorted({parent_of[t] for t in cols} - set(active_parents))
    if dropped:
        log("ROTATION", kv(event="parent_not_in_s_alloc", parents=",".join(dropped)), M=M, level="warning")
        cols = [t for t in cols if parent_of[t] in active_parents]
    passthrough_cols = [c for c in s_all_cols if c not in active_parents]
    all_cols = cols + active_parents + passthrough_cols

    w_s_all = s_alloc["target_w"].reindex(index=eval_idx, columns=s_all_cols).fillna(0.0).astype(float)
    w_s = w_s_all[active_parents]

    state = pd.DataFrame({t: results[t]["state"] for t in cols}).reindex(eval_idx)
    listed = state.notna()
    eligible = listed & ~state.isin(list(icfg.INDUSTRY_EXCLUDE_STATES))

    # 수익 분해: 산업은 M.run_backtest 산출(ret_co/ret_oc), S 열(부모·통과)은 S가 자기 백테스트에 쓴 바로 그 값
    ret_co = pd.DataFrame(index=eval_idx, columns=all_cols, dtype=float)
    ret_oc = pd.DataFrame(index=eval_idx, columns=all_cols, dtype=float)
    for t in cols:
        ret_co[t] = results[t]["ret_co"].reindex(eval_idx)
        ret_oc[t] = results[t]["ret_oc"].reindex(eval_idx)
    s_co = s_alloc.get("ret_co")
    s_oc = s_alloc.get("ret_oc")
    for c in active_parents + passthrough_cols:
        if s_co is not None and c in s_co.columns:
            ret_co[c] = s_co[c].reindex(eval_idx)
            ret_oc[c] = s_oc[c].reindex(eval_idx)
        else:   # 구버전 S 호환 — 섹터 결과에서 직접
            pr = sres["sectors"].get(c, {})
            ret_co[c] = pr.get("ret_co", pd.Series(dtype=float)).reindex(eval_idx)
            ret_oc[c] = pr.get("ret_oc", pd.Series(dtype=float)).reindex(eval_idx)
    ret_co = ret_co.fillna(0.0)
    ret_oc = ret_oc.fillna(0.0)
    if rf_daily is None:
        rf_daily = s_alloc.get("rf_daily")
    init_exec = s_alloc.get("init_exec")
    init_prev = s_alloc.get("init_prev")
    init_exec = init_exec.reindex(all_cols).fillna(0.0) if init_exec is not None else None
    init_prev = init_prev.reindex(all_cols).fillna(0.0) if init_prev is not None else None
    s_cost = float(s_alloc.get("cost_bps", icfg.PARENT_COST_BPS))
    if abs(s_cost - float(icfg.PARENT_COST_BPS)) > 1e-9:
        log("ROTATION", kv(event="parent_cost_differs_from_s", s_cost_bps=s_cost, parent_cost_bps=icfg.PARENT_COST_BPS,
                           note="S★ 재현 행은 S의 비용률, ★/격자의 부모·통과 다리는 PARENT_COST_BPS 사용"), M=M, level="warning")
    cost_map = {c: float(icfg.COST_BPS_INDUSTRY) for c in cols}
    cost_map.update({c: float(icfg.PARENT_COST_BPS) for c in active_parents + passthrough_cols})
    cost_map_repro = dict(cost_map)
    cost_map_repro.update({c: s_cost for c in active_parents + passthrough_cols})

    # ---- [v0.3.0 §B2] 부모 자기 국면(S의 run_sector 결과) — 리더 인정 국면 제약에 쓴다 ----
    parent_state = {p: sres["sectors"].get(p, {}).get("state", pd.Series(dtype=object)).reindex(eval_idx)
                    for p in active_parents}
    # ---- [v0.3.0 §B5] 산업-부모 롤링 상관(252일, 1일 지연 — 인과). 리더 후보 제한에만 쓴다 ----
    follow_corr = None
    if getattr(icfg, "INDUSTRY_LEADER_MIN_CORR", None) is not None:
        w_corr = int(getattr(icfg, "INDUSTRY_FOLLOW_CORR_WINDOW", 252))
        cc = {}
        for t in cols:
            p = parent_of[t]
            r_i = results[t]["ret_cc_full"]
            r_p = sres["sectors"].get(p, {}).get("bh_ret", pd.Series(dtype=float)).reindex(r_i.index)
            cc[t] = r_i.rolling(w_corr, min_periods=w_corr // 2).corr(r_p).shift(1)
        follow_corr = pd.DataFrame(cc).reindex(eval_idx)
        log("ROTATION", kv(event="follow_corr_ready", window=w_corr, min_corr=icfg.INDUSTRY_LEADER_MIN_CORR,
                           median=round(float(follow_corr.stack().median()), 3) if follow_corr.notna().any().any() else -99), M=M)

    # ---- 부모 그룹별 leader3 판단(캡·폴백 비율과 무관 — 격자 변형이 공유) ----
    rank_full = wf.get("rank_full", {})

    def _run_groups(cfg_over: Optional[Dict[str, Any]] = None) -> Dict[str, Dict[str, Any]]:
        """[v0.4.0 §I2·§I3] 같은 신호(rank_full)로 '판단 규칙'만 바꿔 다시 돌린다 — 격자 행용.
        신호를 재계산하지 않으므로 배분층 격자의 전제(하나의 신호를 공유)를 지킨다."""
        _c = dataclasses.replace(icfg, **cfg_over) if cfg_over else icfg
        out: Dict[str, Dict[str, Any]] = {}
        for p_ in active_parents:
            inds_ = [t for t in cols if parent_of[t] == p_]
            out[p_] = leader3_group(p_, inds_, eval_idx, rank_full, wf, eligible, listed, _c, S,
                                    parent_state=parent_state.get(p_), follow_corr=follow_corr)
        return out

    groups: Dict[str, Dict[str, Any]] = {}
    for p in active_parents:
        inds = [t for t in cols if parent_of[t] == p]
        groups[p] = leader3_group(p, inds, eval_idx, rank_full, wf, eligible, listed, icfg, S,
                                  parent_state=parent_state.get(p), follow_corr=follow_corr)
        g = groups[p]
        tc = g["tier"].value_counts()
        log("ROTATION", kv(event="industry_leader_applied", parent=p, n_ind=len(inds),
                           days_leader=int(tc.get("리더", 0)), days_avoid=int(tc.get("회피", 0)),
                           days_fallback=int(tc.get("폴백", 0)), days_fallback_gate=int(tc.get("폴백(여유부족)", 0)),
                           days_parent_only=int(tc.get("부모ETF", 0)), switches=g["switches"],
                           gate_pass_rate=round(float((g["gate"] == "통과").mean()), 4),
                           regime_block=int((g["regime_gate"].astype(str).str.startswith("차단")).sum()),
                           rev_avoid_days=int((g["rev_avoid"] != "").sum()),
                           corr_block_days=int((g["corr_block"] != "").sum()),
                           inert=";".join(f"{k}={v}" for k, v in sorted((g.get("inert_days") or {}).items(),
                                                                        key=lambda x: -x[1])) or "-",
                           leaders=";".join(f"{k}:{v}" for k, v in g["leader"][g["leader"] != ""].value_counts().items()) or "-"),
            M=M)

    def _mk_target_w(leader_cap: float, fallback_share: float, use_rank: bool = True,
                     only_mode: Optional[str] = None, groups_over: Optional[Dict[str, Dict[str, Any]]] = None,
                     neutral_share: float = 0.0) -> pd.DataFrame:
        """[§7.2 + v0.3.0 §A3] 부모 비중 w_s[p]를 산업/부모ETF로 나눈다. only_mode는 '잔여'(산업으로
        배분되지 않은 몫)의 목적지: "parent"(기본, 잔여=부모ETF — 잔여가 정확히 S★로 환원) |
        "industries"(잔여도 그 부모의 적격 산업 균등 — 사용자 지시 '산업만 배분'. 적격 0개면 부모ETF) |
        "cash"(잔여=현금 — 총노출이 S★보다 작아진다)."""
        mode = (only_mode or getattr(icfg, "INDUSTRY_ONLY_MODE", "parent")).lower()
        _G = groups_over or groups
        ind_state = state
        tw = pd.DataFrame(0.0, index=eval_idx, columns=all_cols)
        for p in active_parents:
            g = _G[p]
            inds = g["inds"]
            if use_rank:
                frac = g["leader_ind"] * leader_cap + g["basket_ind"] * fallback_share
                # [v0.4.0 §I3 중립바스켓] 부모 안 '자기국면 중립' 산업 균등 — 기존 폴백('적격=상승')과
                #   다른 집합이다. §3.4 H2 실측: 중립 바스켓 − 부모 +0.46%/21일(t 2.03, 6/9년).
                #   리더가 있는 날에는 그 몫을 침범하지 않도록 (1−리더몫) 안에서만 준다.
                if neutral_share > 0:
                    st_i = ind_state[inds].reindex(eval_idx) if ind_state is not None else None
                    if st_i is not None:
                        nm = st_i.astype(str).eq("NEUTRAL") & listed[inds].reindex(eval_idx).fillna(False)
                        nf = nm.astype(float)
                        share = nf.div(nf.sum(axis=1).replace(0, np.nan), axis=0).fillna(0.0)
                        room = (1.0 - frac.sum(axis=1)).clip(lower=0.0)
                        frac = frac + share.mul(room * float(neutral_share), axis=0).fillna(0.0)
            else:   # 대조군B: 순위 미사용 — 적격 산업 균등 × fallback_share(리더 개념 없음)
                e = eligible[inds].astype(float)
                frac = e.div(e.sum(axis=1).replace(0, np.nan), axis=0).fillna(0.0) * fallback_share
            if mode == "industries":
                resid = (1.0 - frac.sum(axis=1)).clip(lower=0.0)
                e = eligible[inds].astype(float)
                share = e.div(e.sum(axis=1).replace(0, np.nan), axis=0)     # 적격 0개인 날은 NaN → 0(부모로 남음)
                frac = frac + share.mul(resid, axis=0).fillna(0.0)
            for t in inds:
                tw[t] = frac[t].values * w_s[p].values
            if mode == "cash":
                tw[p] = 0.0                                                  # 잔여를 현금으로(총노출 ≤ S★)
            else:
                tw[p] = (1.0 - frac.sum(axis=1)).values * w_s[p].values
        for c in passthrough_cols:
            tw[c] = w_s_all[c].values
        return tw

    def _bt(tw: pd.DataFrame, cmap: Dict[str, float]) -> pd.DataFrame:
        return _industry_portfolio_backtest(S, tw, ret_co, ret_oc, rf_daily, cmap,
                                            init_exec=init_exec, init_prev=init_prev)

    live_cap, live_fb = float(icfg.INDUSTRY_LEADER_CAP), float(icfg.INDUSTRY_FALLBACK_SHARE)
    live_mode = str(getattr(icfg, "INDUSTRY_ONLY_MODE", "parent")).lower()
    label_star = f"부모비중 안 산업리더 {live_cap:.0%}·폴백 {live_fb:.0%}·잔여 {live_mode} ★"
    target_ws: Dict[str, pd.DataFrame] = {label_star: _mk_target_w(live_cap, live_fb)}
    # [v0.3.0 §B1] 2D 사전등록 격자(cap × fb) — 1D 사다리 두 개로는 (1.0, 0.0) 같은 조합을 못 본다.
    #   리포트41 격자에서 ①②③④를 전부 통과한 행이 '폴백 0%'였고 '리더캡 100%'도 강건 통과였는데,
    #   그 둘의 조합은 측정된 적이 없었다. 라벨은 S의 _is_cap_grid 관행대로 여는 대괄호 접두로 매칭한다.
    for cv in (0.25, 0.5, 0.75, 1.0):
        for fv in (0.0, 0.25, 0.5, 1.0):
            if abs(cv - live_cap) < 1e-9 and abs(fv - live_fb) < 1e-9:
                continue                                  # 라이브 조합은 ★ 행이 이미 있다
            target_ws[f"리더 {cv:.0%}·폴백 {fv:.0%} [산업집중격자]"] = _mk_target_w(cv, fv)
    # [v0.3.0 §A3] 잔여 처리 모드 격자 — "산업만 배분"이 실제로 무엇을 바꾸는지 숫자로 보여 준다.
    for md in ("parent", "industries", "cash"):
        if md == live_mode:
            continue
        target_ws[f"잔여 {md} [잔여격자]"] = _mk_target_w(live_cap, live_fb, only_mode=md)
    # [v0.4.0 §I2 국면게이트격자] 리더를 인정하는 부모 자기국면 집합을 바꿔 다시 판단한다.
    #   근거(§3.4 H1): 부모 **중립**일 때 고베타 1위의 부모초과가 +0.81%/21일(t 1.94, 6/8년)로 가장 컸고
    #   부모 상승(+0.25, t 0.73)·하락(−0.21, t −0.32)은 약했다. v0.3.0이 켠 ("RISK_ON",)은 그 반대편이다.
    #   13l B블록 실측과도 맞는다: '국면 게이트 차단일'의 1위>부모 비율 0.531·+0.29% vs '해당없음' 0.471·−0.21%.
    _live_reg = getattr(icfg, "INDUSTRY_LEADER_REGIMES", None)
    for rg in (getattr(icfg, "INDUSTRY_LEADER_REGIME_GRID", ()) or ()):
        if (tuple(rg) if rg else None) == (tuple(_live_reg) if _live_reg else None):
            continue
        lbl = "제약없음" if not rg else "+".join(rg)
        try:
            target_ws[f"리더국면 {lbl} [국면게이트격자]"] = _mk_target_w(
                live_cap, live_fb, groups_over=_run_groups({"INDUSTRY_LEADER_REGIMES": (tuple(rg) if rg else None)}))
        except Exception as e:   # noqa
            log("ROTATION", kv(event="regime_grid_failed", value=lbl, err=str(e)[:120]), M=M, level="warning")
    # [v0.4.0 §I3 중립바스켓격자] 부모 안 '자기국면 중립' 산업 균등 — 기존 폴백('적격=상승')과 다른 집합.
    #   §3.4 H2: 중립 바스켓 − 부모 +0.46%/21일(t 2.03, 6/9년). 부모 상승·중립 국면 모두에서 양수였다.
    for ns in (getattr(icfg, "INDUSTRY_NEUTRAL_BASKET_GRID", ()) or ()):
        target_ws[f"중립바스켓 {float(ns):.0%} [중립바스켓격자]"] = _mk_target_w(live_cap, live_fb, neutral_share=float(ns))
    # [v0.4.0 §I3 분산게이트격자] 부모 안 63일 상대수익 횡단면 σ가 롤링 252일 상위 q분위인 날만 리더 인정.
    #   §3.4 H3: 분산 Q4에서 REL_NEAR_HIGH 1위 +0.41%/21일 vs Q1 +0.02% — '고를 값어치가 있을 때만 고른다'.
    #   S의 분산게이트(v0.36.0)와 같은 취지이며, 여기서는 부모별로 계산한다.
    _disp_grid = tuple(getattr(icfg, "INDUSTRY_DISPERSION_GATE_GRID", ()) or ())
    if _disp_grid:
        _rel63 = {}
        for p_ in active_parents:
            inds_ = [t for t in cols if parent_of[t] == p_]
            if len(inds_) < 2:
                continue
            _rr = pd.DataFrame({t: results[t]["rot_raw"].get("REL_RET", pd.Series(dtype=float)) for t in inds_})
            _cum = (1.0 + _rr.fillna(0.0)).cumprod().where(_rr.notna())
            _rel63[p_] = (_cum / _cum.shift(63) - 1.0).std(axis=1).reindex(eval_idx)
        for q in _disp_grid:
            gq = {}
            for p_ in active_parents:
                d = _rel63.get(p_)
                g0 = groups[p_]
                if d is None or d.notna().sum() < 252:
                    gq[p_] = g0
                    continue
                thr = d.rolling(252, min_periods=126).quantile(float(q)).shift(1)   # 1일 지연 — 인과
                keep = (d >= thr).fillna(False)
                g2 = dict(g0)
                g2["leader_ind"] = g0["leader_ind"].mul(keep.astype(float), axis=0)
                g2["basket_ind"] = g0["basket_ind"]
                gq[p_] = g2
            target_ws[f"분산게이트 상위{1 - float(q):.0%} [분산게이트격자]"] = _mk_target_w(live_cap, live_fb, groups_over=gq)
    label_ctrl_b = "대조군B: 부모비중 안 적격산업 균등 50%(순위 미사용)"
    target_ws[label_ctrl_b] = _mk_target_w(0.0, 0.5, use_rank=False, only_mode="parent")
    label_repro = "S★ 재현(I 백테스트 엔진, 산업 0% — 대조군A와 비트 동일해야 함)"
    target_ws[label_repro] = _mk_target_w(0.0, 0.0, only_mode="parent")
    label_ctrl_a = "대조군A: S★ 그대로(산업 미사용)"
    target_ws[label_ctrl_a] = _mk_target_w(0.0, 0.0, only_mode="parent")

    bts: Dict[str, pd.DataFrame] = {}
    star_label = next((c for c in s_alloc.get("bts", {}) if str(c).endswith("★")), None)
    for label, tw in target_ws.items():
        if label == label_ctrl_a and star_label is not None:
            # [§7.4] "= S★ 비트 동일" — S가 이미 계산한 바로 그 일간수익을 그대로 쓴다(재계산으로
            # 인한 미세한 비용/반올림 차이조차 없게).
            s_ret = s_alloc["bts"][star_label]["strategy_ret"].reindex(eval_idx).fillna(0.0)
            eq = (1.0 + s_ret).cumprod()
            bts[label] = pd.DataFrame({"strategy_ret": s_ret, "exposure": tw.sum(axis=1),
                                       "turnover": 0.0, "cost": 0.0, "equity": eq,
                                       "dd": eq / eq.cummax() - 1.0}, index=eval_idx)
        elif label == label_repro:
            bts[label] = _bt(tw, cost_map_repro)
        else:
            bts[label] = _bt(tw, cost_map)
    repro_diff = float((bts[label_repro]["strategy_ret"] - bts[label_ctrl_a]["strategy_ret"]).abs().max())
    log("ROTATION", kv(event="s_star_reproduction", max_abs_daily_diff=repro_diff,
                       verdict=("OK" if repro_diff < 1e-10 else "MISMATCH")), M=M,
        level=("info" if repro_diff < 1e-10 else "error"))

    order = [l for l in target_ws if l not in (label_ctrl_a, label_repro)] + [label_repro, label_ctrl_a]
    perf = pd.DataFrame([M.perf_metrics(bts[l]["strategy_ret"], l) for l in order])
    for col_extra, fn in (("평가창", lambda l: f"{eval_idx[0].date()}~{eval_idx[-1].date()}"),
                          ("평균노출", lambda l: round(float(target_ws[l].sum(axis=1).mean()), 4)),
                          ("산업평균노출", lambda l: round(float(target_ws[l][cols].sum(axis=1).mean()), 4)),
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
    # [v0.3.0 §B1] 격자 수렴 판정 — M v1.49.0 _grid_convergence_line·S와 같은 4기준을 열로 직접 싣는다.
    #   ① CAGR 손실 ≤ GRID_CAGR_LOSS_TOL · ② ① 통과자 중 칼마가 ★보다 높은가 · ③ MDD 악화 ≤ GRID_MDD_WORSE_TOL
    #   · ④ 강건성 '통과'. 넷을 다 통과한 행만 다음 라운드 ★ 후보다(사후 최고 선택 금지 — 사전 고정 규칙).
    _star_row = perf.set_index("전략").loc[label_star]
    _c0, _m0, _k0 = float(_star_row["CAGR"]), float(_star_row["최대낙폭(MDD)"]), float(_star_row["칼마(CAGR/MDD)"])
    g1, g2, g3, g4, gall = [], [], [], [], []
    for l in order:
        r = perf.set_index("전략").loc[l]
        c1 = bool(float(r["CAGR"]) >= _c0 - float(icfg.GRID_CAGR_LOSS_TOL))
        c2 = bool(float(r["칼마(CAGR/MDD)"]) > _k0)
        c3 = bool(float(r["최대낙폭(MDD)"]) >= _m0 - float(icfg.GRID_MDD_WORSE_TOL))
        c4 = bool(str(r["강건성(기준④)"]) == "통과")
        is_star = (l == label_star)
        g1.append("★" if is_star else ("O" if c1 else "X")); g2.append("★" if is_star else ("O" if c2 else "X"))
        g3.append("★" if is_star else ("O" if c3 else "X")); g4.append("★" if is_star else ("O" if c4 else "X"))
        gall.append("★" if is_star else ("채택후보" if (c1 and c2 and c3 and c4) else ""))
    perf["①CAGR손실"], perf["②칼마정점"], perf["③MDD악화"], perf["④강건성"] = g1, g2, g3, g4
    perf["격자판정"] = gall
    _cand = [l for l, v in zip(order, gall) if v == "채택후보"]
    grid_line = (f"격자 {len(order)}행 — ① 통과 {sum(1 for v in g1 if v=='O')} · ④ 통과 {sum(1 for v in g4 if v=='O')} · "
                 f"**①②③④ 전부 통과 {len(_cand)}행** (후보: {', '.join(_cand) if _cand else '없음'}) | "
                 f"★ CAGR {_c0:.2%} 칼마 {_k0:.3f} MDD {_m0:.2%} · 격자 칼마 정점 "
                 f"{max(float(perf.set_index('전략').loc[l]['칼마(CAGR/MDD)']) for l in order):.3f} — 0행이면 이번 실행에 채택 후보 없음")
    log("ROTATION", kv(event="grid_convergence", rows=len(order), pass_all=len(_cand),
                       candidates=";".join(_cand) or "-"), M=M)

    curve = pd.DataFrame(index=eval_idx)
    curve["날짜"] = eval_idx.date
    for l in order:
        curve[l] = bts[l]["equity"].values
    curve["S★ 총노출"] = w_s_all.sum(axis=1).values
    curve["부모(산업 보유) 비중 합"] = w_s.sum(axis=1).values
    for l in order:
        pr = perf.set_index("전략").loc[l]
        log("ROTATION", kv(event="allocation_perf", strategy=l, cagr=pr["CAGR"], sharpe=pr["샤프"],
                           mdd=pr["최대낙폭(MDD)"], avg_exposure=pr["평균노출"], ind_exposure=pr["산업평균노출"]), M=M)
    log("ROTATION", kv(event="allocation_built", n_industries=len(cols), n_parents=len(active_parents),
                       passthrough=",".join(passthrough_cols) or "-", variants=len(target_ws),
                       elapsed_s=round(time.time() - t0, 2)), M=M)
    return {
        "target_w": target_ws[label_star], "target_ws": target_ws, "bts": bts, "perf": perf,
        "curve": curve.reset_index(drop=True), "cols": cols, "active_parents": active_parents,
        "parent_of": parent_of, "all_cols": all_cols, "passthrough_cols": passthrough_cols,
        "eligible": eligible, "listed": listed, "state": state, "w_s": w_s, "w_s_all": w_s_all,
        "grid_line": grid_line, "grid_candidates": _cand, "only_mode": live_mode,   # [v0.3.0 §B1/§A3]
        "parent_state": parent_state, "follow_corr": follow_corr,
        "groups": groups, "label_star": label_star, "label_ctrl_a": label_ctrl_a, "label_ctrl_b": label_ctrl_b,
        "label_repro": label_repro, "repro_max_diff": repro_diff,
        "ret_co": ret_co, "ret_oc": ret_oc, "cost_bps_industry": icfg.COST_BPS_INDUSTRY,
        "cost_bps_parent": icfg.PARENT_COST_BPS, "rf_daily": rf_daily,
    }


# =============================================================================
# [6] 수용기준(§8) · 계층정합(14) · 산업대섹터귀속(15) · 산업부모추종(16, v0.2.0 신규)
# =============================================================================
def build_industry_acceptance(alloc: Dict[str, Any], wf: Dict[str, Any], icfg: IndustryConfig,
                              M, S) -> pd.DataFrame:
    """[§8 사전 고정 5기준] ① I★ CAGR ≥ S★+ACCEPT_CAGR_GAIN ② MDD 악화 ≤ ACCEPT_MDD_WORSE
    ③ 풀링 상위1−하위1 스프레드 NW-t ≥ ACCEPT_TOP1_T(v0.1: 상위3−하위3 근사) ④ 스프레드 연도별 양수 비율 ≥ 기준
    ⑤ 칼마 ≥ S★ 칼마. S★ = 대조군A(비트 동일)."""
    perf = alloc["perf"]

    def _get(label: str, col: str) -> float:
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
    """[14_계층정합, §7.1] (a) 부모별: 매일 Σ산업 + 부모ETF = S★의 그 섹터비중(허용오차 1e-9)
    (b) 총노출: 매일 Σ(I★ 전체 열) = Σ(S★ 전체 열) — v0.2.0 신규(통과 다리 포함 여부 검사)
    (c) S★ 재현: I 백테스트 엔진에 '산업 0%'를 넣은 일간수익이 대조군A(S★ 그대로)와 비트 동일한지."""
    tw, cols, parent_of, active_parents, w_s = (alloc["target_w"], alloc["cols"], alloc["parent_of"],
                                                 alloc["active_parents"], alloc["w_s"])
    # [v0.3.0 §A3] 잔여를 현금으로 두는 모드에서는 '='이 아니라 '≤'가 옳은 불변식이다(총노출이 S★보다 작아진다).
    mode = str(alloc.get("only_mode", "parent")).lower()
    _eq = (mode != "cash")
    _lbl = "부모별 Σ산업+부모ETF = S★ 섹터비중" if _eq else "부모별 Σ산업+부모ETF ≤ S★ 섹터비중(잔여=현금)"
    rows, max_err_all, viol_all = [], 0.0, 0
    for p in active_parents:
        inds = [c for c in cols if parent_of[c] == p]
        total = tw[inds].sum(axis=1) + tw[p]
        err = (total - w_s[p]).abs() if _eq else (total - w_s[p]).clip(lower=0.0)
        max_err = float(err.max()) if len(err) else 0.0
        n_viol = int((err > 1e-9).sum())
        max_err_all = max(max_err_all, max_err)
        viol_all += n_viol
        rows.append({"검사": _lbl, "부모섹터": p, "산업수": len(inds),
                     "최대오차": max_err, "위반일수(>1e-9)": n_viol})
    rows.append({"검사": _lbl, "부모섹터": "전체", "산업수": len(cols),
                 "최대오차": max_err_all, "위반일수(>1e-9)": viol_all})
    w_s_all = alloc.get("w_s_all")
    if w_s_all is not None:
        diff = tw.sum(axis=1) - w_s_all.sum(axis=1)
        err_tot = diff.abs() if _eq else diff.clip(lower=0.0)
        rows.append({"검사": ("총노출 Σ(I★ 전체 열) = Σ(S★ 전체 열)" if _eq else "총노출 Σ(I★) ≤ Σ(S★)(잔여=현금)"),
                     "부모섹터": "전체(통과 다리 " + (",".join(alloc.get("passthrough_cols", [])) or "없음") + " 포함)",
                     "산업수": len(cols), "최대오차": float(err_tot.max()) if len(err_tot) else 0.0,
                     "위반일수(>1e-9)": int((err_tot > 1e-9).sum())})
    if "repro_max_diff" in alloc:
        d = float(alloc["repro_max_diff"])
        rows.append({"검사": "S★ 재현(I 엔진, 산업 0%) 일간수익 = 대조군A", "부모섹터": "전체", "산업수": 0,
                     "최대오차": d, "위반일수(>1e-9)": int(d > 1e-9)})
    return pd.DataFrame(rows)


def build_parent_following_analysis(results: Dict[str, Dict[str, Any]], sres: dict, eval_idx: pd.DatetimeIndex,
                                    icfg: IndustryConfig) -> pd.DataFrame:
    """[16_산업부모추종, v0.2.0 신규 — 사용자 질문 "산업별로도 해당 섹터가 오르면 따라가는게 있고 아닌게 있다"]
    산업이 부모 섹터의 움직임을 얼마나 '따라가는지'를 산업마다 정량화한다(진단 전용, 신호에 미사용).
    부모 일간수익은 재수집 없이 산업 일간수익과 산업/부모 비율수익(rot_raw REL_RET)에서 복원:
    parent = (1+ind)/(1+rel) − 1. 열:
      베타(전체이력)·상관(전체이력)·베타(평가창) — 회귀 기울기/상관
      상승일 추종률 = P(산업>0 | 부모>0), 하락일 추종률 = P(산업<0 | 부모<0)
      상승포착 = 부모 상승일 산업 평균수익 / 부모 평균수익, 하락포착 = 같은 정의(부모 하락일)
      부모 상승국면 초과(연율%) = 부모 자기국면(S의 state)이 RISK_ON인 날 (산업−부모) 평균×252, 중립/하락 동일
      상승국면 월승률 = 부모 RISK_ON인 달 중 산업 월수익 > 부모 월수익인 달의 비율
      국면 일치율 = P(산업 자기국면==부모 자기국면) (평가창)
      유형 = 상관·베타로 분류: 고베타 추종형(β≥1.15·ρ≥0.8) / 저베타 추종형(β≤0.85·ρ≥0.8) / 동행형(ρ≥0.8) /
            부분추종(0.6≤ρ<0.8) / 독립형(ρ<0.6)"""
    rows: List[dict] = []
    for t, r in results.items():
        p = r["parent"]
        ind_ret = r.get("ret_cc_full")
        rel = r.get("rot_raw", pd.DataFrame()).get("REL_RET") if isinstance(r.get("rot_raw"), pd.DataFrame) else None
        if ind_ret is None or rel is None:
            continue
        df = pd.DataFrame({"ind": ind_ret, "rel": rel}).dropna()
        df["par"] = (1.0 + df["ind"]) / (1.0 + df["rel"]) - 1.0
        df = df[(df["ind"].abs() < 0.5) & (df["par"].abs() < 0.5)]
        if len(df) < 252:
            continue
        x, y = df["par"].values, df["ind"].values
        vx = float(np.var(x))
        beta_full = float(np.cov(x, y)[0, 1] / vx) if vx > 0 else np.nan
        corr_full = float(np.corrcoef(x, y)[0, 1]) if vx > 0 else np.nan
        ev = df.loc[df.index.intersection(eval_idx)]
        if len(ev) >= 60:
            vxe = float(np.var(ev["par"].values))
            beta_eval = float(np.cov(ev["par"].values, ev["ind"].values)[0, 1] / vxe) if vxe > 0 else np.nan
        else:
            beta_eval = np.nan
        up = df["par"] > 0
        dn = df["par"] < 0
        follow_up = float((df.loc[up, "ind"] > 0).mean()) if up.any() else np.nan
        follow_dn = float((df.loc[dn, "ind"] < 0).mean()) if dn.any() else np.nan
        cap_up = float(df.loc[up, "ind"].mean() / df.loc[up, "par"].mean()) if up.any() and df.loc[up, "par"].mean() != 0 else np.nan
        cap_dn = float(df.loc[dn, "ind"].mean() / df.loc[dn, "par"].mean()) if dn.any() and df.loc[dn, "par"].mean() != 0 else np.nan
        # 부모 자기국면(S) 조건부 초과수익 — 평가창
        p_state = sres.get("sectors", {}).get(p, {}).get("state")
        i_state = r.get("state")
        exc_on = exc_neu = exc_off = np.nan
        win_on = np.nan
        agree = np.nan
        if p_state is not None and len(ev):
            ps = p_state.reindex(ev.index)
            ex = (ev["ind"] - ev["par"])
            for key, st in (("on", "RISK_ON"), ("neu", "NEUTRAL"), ("off", "RISK_OFF")):
                m = ps == st
                val = float(ex[m].mean() * 252 * 100) if m.sum() >= 20 else np.nan
                if key == "on":
                    exc_on = val
                elif key == "neu":
                    exc_neu = val
                else:
                    exc_off = val
            m_on = ps == "RISK_ON"
            if m_on.sum() >= 40:
                mo = pd.DataFrame({"ind": ev["ind"][m_on], "par": ev["par"][m_on]})
                mret = mo.groupby(mo.index.to_period("M")).apply(lambda g: pd.Series({"ind": (1 + g["ind"]).prod() - 1, "par": (1 + g["par"]).prod() - 1}))
                if len(mret) >= 6:
                    win_on = float((mret["ind"] > mret["par"]).mean())
            if i_state is not None:
                both = pd.DataFrame({"i": i_state.reindex(ev.index), "p": ps}).dropna()
                agree = float((both["i"] == both["p"]).mean()) if len(both) else np.nan
        if corr_full != corr_full:
            kind = "산정불가"
        elif corr_full >= 0.8 and beta_full >= 1.15:
            kind = "고베타 추종형(부모 상승 시 더 오름·하락 시 더 빠짐)"
        elif corr_full >= 0.8 and beta_full <= 0.85:
            kind = "저베타 추종형(방향은 같되 진폭 작음)"
        elif corr_full >= 0.8:
            kind = "동행형(부모와 거의 같이 움직임)"
        elif corr_full >= 0.6:
            kind = "부분추종(부모 외 요인 큼)"
        else:
            kind = "독립형(부모를 잘 안 따라감)"
        rows.append({"티커": t, "산업명": INDUSTRY_NAME_KR.get(t, t), "부모섹터": p, "관측일(전체)": int(len(df)),
                     "베타(전체이력)": round(beta_full, 3) if beta_full == beta_full else None,
                     "상관(전체이력)": round(corr_full, 3) if corr_full == corr_full else None,
                     "베타(평가창)": round(beta_eval, 3) if beta_eval == beta_eval else None,
                     "상승일 추종률": round(follow_up, 3) if follow_up == follow_up else None,
                     "하락일 추종률": round(follow_dn, 3) if follow_dn == follow_dn else None,
                     "상승포착": round(cap_up, 3) if cap_up == cap_up else None,
                     "하락포착": round(cap_dn, 3) if cap_dn == cap_dn else None,
                     "부모 상승국면 초과(연율%)": round(exc_on, 2) if exc_on == exc_on else None,
                     "부모 중립국면 초과(연율%)": round(exc_neu, 2) if exc_neu == exc_neu else None,
                     "부모 하락국면 초과(연율%)": round(exc_off, 2) if exc_off == exc_off else None,
                     "상승국면 월승률(vs 부모)": round(win_on, 3) if win_on == win_on else None,
                     "국면 일치율(산업=부모)": round(agree, 3) if agree == agree else None,
                     "유형": kind})
    out = pd.DataFrame(rows)
    if len(out):
        out = out.sort_values(["부모섹터", "티커"]).reset_index(drop=True)
    return out


# =============================================================================
# [6b] 진단 시트 — 13l_산업리더적중률 · 01Y_산업예측정확도 (v0.3.0 §C1 신설)
#      리포트41 §3.2/부록A를 사람이 손으로 재계산해 만들었던 표를, 이제 리포트가 직접 낸다.
#      둘 다 진단 전용 — 신호·배분에 전혀 관여하지 않는다.
# =============================================================================
def _fwd_ret(r: pd.DataFrame, h: int) -> pd.DataFrame:
    """t일 확정 → t+1 시가 체결 규칙에 맞춘 향후 h일 누적수익(로그합 → 되돌림). r은 일간 단순수익."""
    lr = np.log1p(r)
    return np.exp(lr.shift(-1)[::-1].rolling(h, min_periods=h).sum()[::-1]) - 1.0


def build_industry_leader_accuracy(alloc: Dict[str, Any], results: Dict[str, Dict[str, Any]],
                                   sres: dict, icfg: IndustryConfig) -> pd.DataFrame:
    """[13l_산업리더적중률, v0.3.0 §C1] "산업 순환매가 실제로 맞았나"에 리포트가 직접 답한다.
    채점 기준은 전부 **부모 대비**다(이 계층의 일은 '어느 섹터가 오르나'가 아니라 '그 섹터 안에서
    어느 산업이 앞서나'이므로 — §6.2와 같은 정의).
      A. 지평별(h=1/5/21/63): 리더가 부모를 이긴 비율·평균초과 / 같은 부모의 다른 산업 평균초과(기준선) /
         폴백 바스켓 초과. 리더 열이 타산업 열보다 높아야 '고른 것'에 값이 있다.
      B. 확신 게이트·국면 게이트 효과(h=21): 통과/미달/해당없음, 차단된 날.
      C. 리더 에피소드 전수(진입일 기준): 보유일·보유기간 초과 — 부모별/연도별/리더별 승률.
      D. 진입 시 부모 국면별 승률.
    한 줄이라도 계산할 수 없으면(리더 0일 등) 그 블록은 안내 행만 남긴다."""
    groups = alloc.get("groups", {})
    if not groups:
        return pd.DataFrame([{"블록": "안내", "구분": "배분 결과 없음 — 13l 생략"}])
    eval_idx = alloc["target_w"].index
    cols, parent_of = alloc["cols"], alloc["parent_of"]
    R = pd.DataFrame({t: results[t]["ret_cc_full"].reindex(eval_idx) for t in cols})
    RP = pd.DataFrame({p: sres["sectors"].get(p, {}).get("bh_ret", pd.Series(dtype=float)).reindex(eval_idx)
                       for p in alloc["active_parents"]})
    rows: List[dict] = []

    # ---- A. 지평별 ----
    rows.append({"블록": "A. 지평별 리더 적중률", "구분": "── 리더가 '부모'를 이긴 비율 ──",
                 "설명": "이 계층의 일은 섹터 선택이 아니라 섹터 안 산업 선택 — 그래서 기준선은 부모 ETF다. "
                         "'타산업 평균초과'보다 '리더 평균초과'가 높아야 순위에 값이 있다."})
    F = {h: _fwd_ret(R, h) for h in (1, 5, 21, 63)}
    FP = {h: _fwd_ret(RP, h) for h in (1, 5, 21, 63)}
    for h in (1, 5, 21, 63):
        lead_ex, other_ex, bask_ex, n_lead, n_fb = [], [], [], 0, 0
        for p, g in groups.items():
            inds, tier, lead = g["inds"], g["tier"], g["leader"]
            if p not in FP[h].columns:
                continue
            m = tier == "리더"
            n_lead += int(m.sum())
            for d in eval_idx[m]:
                l = lead.at[d]
                if not l or pd.isna(F[h].at[d, l]) or pd.isna(FP[h].at[d, p]):
                    continue
                lead_ex.append(F[h].at[d, l] - FP[h].at[d, p])
                rest = [c for c in inds if c != l]
                if rest:
                    v = F[h].loc[d, rest].mean()
                    if pd.notna(v):
                        other_ex.append(v - FP[h].at[d, p])
            fb = tier.isin(["폴백", "폴백(여유부족)", "회피"])
            n_fb += int(fb.sum())
            bw = g["basket_ind"].loc[fb]
            for d in bw.index:
                w = bw.loc[d]
                w = w[w > 1e-9]
                if len(w) and pd.notna(FP[h].at[d, p]):
                    v = F[h].loc[d, list(w.index)].mean()
                    if pd.notna(v):
                        bask_ex.append(v - FP[h].at[d, p])
        def _agg(a):
            a = np.array([x for x in a if pd.notna(x)], dtype=float)
            return (round(float((a > 0).mean()), 4), round(float(a.mean() * 100), 4), len(a)) if len(a) else (np.nan, np.nan, 0)
        w1, m1, k1 = _agg(lead_ex); w2, m2, k2 = _agg(other_ex); w3, m3, k3 = _agg(bask_ex)
        rows.append({"블록": "A. 지평별 리더 적중률", "구분": f"h={h}일", "리더일": n_lead,
                     "리더>부모 비율": w1, f"리더 평균초과(%/{h}일)": m1,
                     "타산업>부모 비율": w2, f"타산업 평균초과(%/{h}일)": m2,
                     "폴백일": n_fb, "폴백바스켓>부모 비율": w3, f"폴백바스켓 평균초과(%/{h}일)": m3,
                     "표본(리더)": k1})

    # ---- B. 게이트 효과(h=21) ----
    rows.append({"블록": "B. 게이트 효과(h=21일)", "구분": "── 게이트가 좋은 날만 통과시키나 ──",
                 "설명": "통과일의 '1위>부모 비율'이 미달·해당없음보다 높아야 게이트가 제 일을 하는 것이다. "
                         "국면 차단일은 리더 후보가 있었는데 부모 국면 때문에 막은 날이다(v0.3.0 §B2)."})
    for key, sel in (("확신 게이트 통과", lambda g: g["gate"] == "통과"),
                     ("확신 게이트 미달", lambda g: g["gate"] == "미달"),
                     ("채택 신호 없음(리더 불가)", lambda g: g["gate"] == "해당없음"),
                     ("국면 게이트 차단", lambda g: g["regime_gate"].astype(str).str.startswith("차단"))):
        ex, n = [], 0
        for p, g in groups.items():
            if p not in FP[21].columns or key.startswith("국면") and "regime_gate" not in g:
                continue
            m = sel(g)
            n += int(m.sum())
            comp = g["composite"]
            for d in eval_idx[m]:
                r = comp.loc[d].dropna()
                if not len(r) or pd.isna(FP[21].at[d, p]):
                    continue
                top = r.idxmax()
                if pd.notna(F[21].at[d, top]):
                    ex.append(F[21].at[d, top] - FP[21].at[d, p])
        a = np.array(ex, dtype=float)
        rows.append({"블록": "B. 게이트 효과(h=21일)", "구분": key, "리더일": n,
                     "리더>부모 비율": round(float((a > 0).mean()), 4) if len(a) else np.nan,
                     "리더 평균초과(%/21일)": round(float(a.mean() * 100), 4) if len(a) else np.nan,
                     "표본(리더)": len(a)})

    # ---- C. 리더 에피소드 전수 ----
    eps: List[dict] = []
    for p, g in groups.items():
        tier, lead = g["tier"], g["leader"]
        if p not in RP.columns:
            continue
        prev = lead.shift(1)
        entry = (tier == "리더") & (lead != prev)
        pst = g.get("parent_state")
        for d in eval_idx[entry]:
            l = lead.at[d]
            j = eval_idx.get_loc(d)
            k = j
            while k + 1 < len(eval_idx) and tier.iloc[k + 1] == "리더" and lead.iloc[k + 1] == l:
                k += 1
            hold_i = R[l].iloc[j + 1:k + 2]
            hold_p = RP[p].iloc[j + 1:k + 2]
            exc = float(((1 + hold_i).prod() - (1 + hold_p).prod()) * 100) if len(hold_i) else np.nan
            eps.append({"부모": p, "진입일": str(d.date()), "리더": l, "보유일": k - j + 1,
                        "보유기간 초과(%)": round(exc, 3) if pd.notna(exc) else None,
                        "진입시 부모국면": (str(pst.at[d]) if pst is not None else "-"),
                        "연도": int(d.year)})
    E = pd.DataFrame(eps)
    rows.append({"블록": "C. 리더 에피소드", "구분": "── 진입일 기준 전수(보유기간 산업−부모) ──",
                 "설명": f"총 {len(E)}회" + ("" if len(E) else " — 리더가 한 번도 나오지 않았다(신호 채택·게이트·국면 제약 확인)")})
    if len(E):
        ok = E["보유기간 초과(%)"].notna()
        rows.append({"블록": "C. 리더 에피소드", "구분": "전체", "에피소드": len(E),
                     "승률(부모대비)": round(float((E.loc[ok, "보유기간 초과(%)"] > 0).mean()), 4),
                     "평균초과(%)": round(float(E.loc[ok, "보유기간 초과(%)"].mean()), 3),
                     "중앙값초과(%)": round(float(E.loc[ok, "보유기간 초과(%)"].median()), 3),
                     "평균보유일": round(float(E["보유일"].mean()), 1)})
        for key, gcol in (("부모별", "부모"), ("연도별", "연도"), ("리더별", "리더"), ("진입시 부모국면별", "진입시 부모국면")):
            for kk, sub in E.groupby(gcol):
                o = sub["보유기간 초과(%)"].notna()
                rows.append({"블록": f"C. 리더 에피소드 — {key}", "구분": str(kk), "에피소드": len(sub),
                             "승률(부모대비)": round(float((sub.loc[o, "보유기간 초과(%)"] > 0).mean()), 4) if o.any() else np.nan,
                             "평균초과(%)": round(float(sub.loc[o, "보유기간 초과(%)"].mean()), 3) if o.any() else np.nan})
        for _, r in E.sort_values("진입일").iterrows():
            rows.append({"블록": "D. 에피소드 전수", "구분": f"{r['부모']} · {r['진입일']} · {r['리더']}",
                         "보유일": int(r["보유일"]), "보유기간 초과(%)": r["보유기간 초과(%)"],
                         "진입시 부모국면": r["진입시 부모국면"]})
    return pd.DataFrame(rows)


def build_industry_prediction_accuracy(results: Dict[str, Dict[str, Any]], sres: dict,
                                       eval_idx: pd.DatetimeIndex, icfg: IndustryConfig) -> pd.DataFrame:
    """[01Y_산업예측정확도, v0.3.0 §C1] S의 01Y_섹터예측정확도와 같은 형식의 산업 버전.
    ⚠ 적중률은 0.5가 아니라 **그 산업 자신의 기저 상승률**과 비교해야 한다(기저가 55%인 자산에서
    적중 54%는 정보가 없는 것이다). 블록:
      A. 익일 채점 — 상태[t] → 수익[t+1] (체결 t+1 시가). 상승/하락 예측 각각 기저 대비 %p.
      A2. 부모 대비 — 상승예측일 다음날 (산업 − 부모) 평균 bp. 이 계층이 '부모보다 나은 산업'을 골랐는지.
      B. 지평별(h=1/5/21/63) 상승·하락 기저 대비.
      B2. 상태 지속성 — 상태[t]가 h일 뒤까지 유지되는 비율(낮으면 B의 h열은 실제 운용과 다르다).
      C. 연도별 상승예측 기저대비(%p) — 어느 해에 흔들렸나."""
    rows: List[dict] = []
    rows.append({"블록": "A. 익일 채점", "구분": "── 상태[t] → 수익[t+1] ──",
                 "설명": "상승은 기저 상승률과, 하락은 기저 하락률과 비교한다 — 0.5와 비교하면 오독한다"})
    tickers = list(results.keys())
    per_year: Dict[str, Dict[int, float]] = {}
    for t in tickers:
        r = results[t]["ret_cc_full"].reindex(eval_idx)
        st = results[t]["state"].reindex(eval_idx)
        p = results[t]["parent"]
        rp = sres["sectors"].get(p, {}).get("bh_ret", pd.Series(dtype=float)).reindex(eval_idx)
        nxt = r.shift(-1)
        ok = nxt.notna() & st.notna()
        base_up = float((nxt[ok] > 0).mean()) if ok.any() else np.nan
        up, dn = (st == "RISK_ON") & ok, (st == "RISK_OFF") & ok
        acc_up = float((nxt[up] > 0).mean()) if up.any() else np.nan
        acc_dn = float((nxt[dn] < 0).mean()) if dn.any() else np.nan
        ex = (nxt - rp.shift(-1))
        rows.append({"블록": "A. 익일 채점", "구분": t, "부모": p, "산업명": INDUSTRY_NAME_KR.get(t, t),
                     "표본일수": int(ok.sum()), "기저 상승": round(base_up, 4), "기저 하락": round(1 - base_up, 4),
                     "상승예측 일수": int(up.sum()), "상승예측 적중": round(acc_up, 4) if pd.notna(acc_up) else None,
                     "상승 기저대비(%p)": round((acc_up - base_up) * 100, 2) if pd.notna(acc_up) else None,
                     "상승 정보": ("있음" if pd.notna(acc_up) and acc_up > base_up else ("없음" if pd.notna(acc_up) and abs(acc_up - base_up) < 0.002 else "역방향")),
                     "중립 일수": int(((st == "NEUTRAL") & ok).sum()),
                     "하락예측 일수": int(dn.sum()), "하락예측 적중": round(acc_dn, 4) if pd.notna(acc_dn) else None,
                     "하락 기저대비(%p)": round((acc_dn - (1 - base_up)) * 100, 2) if pd.notna(acc_dn) else None,
                     "상승예측일 다음날(bp)": round(float(nxt[up].mean() * 1e4), 1) if up.any() else None,
                     "상승예측일 다음날 부모대비(bp)": round(float(ex[up].mean() * 1e4), 1) if up.any() and ex[up].notna().any() else None,
                     "하락예측일 다음날 부모대비(bp)": round(float(ex[dn].mean() * 1e4), 1) if dn.any() and ex[dn].notna().any() else None})
        yr: Dict[int, float] = {}
        for y, sub in pd.DataFrame({"n": nxt, "s": st}).dropna().groupby(eval_idx[ok].year if False else lambda d: d.year):
            b = float((sub["n"] > 0).mean())
            u = sub[sub["s"] == "RISK_ON"]
            if len(u) >= 20:
                yr[int(y)] = round((float((u["n"] > 0).mean()) - b) * 100, 1)
        per_year[t] = yr

    rows.append({"블록": "B. 지평별", "구분": "── 예측이 사는 지평 ──", "설명": "기저 대비가 양수인 칸만 정보가 있다"})
    for t in tickers:
        r = results[t]["ret_cc_full"].reindex(eval_idx)
        st = results[t]["state"].reindex(eval_idx)
        for h in (1, 5, 21, 63):
            f = _fwd_ret(r.to_frame("x"), h)["x"]
            ok = f.notna() & st.notna()
            if not ok.any():
                continue
            b = float((f[ok] > 0).mean())
            up, dn = (st == "RISK_ON") & ok, (st == "RISK_OFF") & ok
            au = float((f[up] > 0).mean()) if up.any() else np.nan
            ad = float((f[dn] < 0).mean()) if dn.any() else np.nan
            rows.append({"블록": "B. 지평별", "구분": t, "지평": f"h={h}일", "표본일수": int(ok.sum()),
                         "기저 상승": round(b, 4),
                         "상승 기저대비(%p)": round((au - b) * 100, 2) if pd.notna(au) else None,
                         "하락 기저대비(%p)": round((ad - (1 - b)) * 100, 2) if pd.notna(ad) else None})
    rows.append({"블록": "B2. 상태 지속성", "구분": "── B의 h열을 얼마나 믿을 수 있나 ──",
                 "설명": "상태[t]가 h일 뒤까지 그대로인 비율. 낮으면 B의 h열은 실제 운용(매일 재판단)과 다르다"})
    for t in tickers:
        st = results[t]["state"].reindex(eval_idx)
        rec = {"블록": "B2. 상태 지속성", "구분": t}
        for h in (1, 5, 21, 63):
            rec[f"h={h}일 유지율"] = round(float((st == st.shift(-h)).where(st.notna() & st.shift(-h).notna()).mean()), 4)
        rows.append(rec)
    rows.append({"블록": "C. 연도별 상승예측 기저대비(%p)", "구분": "── 어느 해에 흔들렸나 ──",
                 "설명": "값은 '그 해 상승예측 적중률 − 그 해 기저 상승률'(%p). 음수면 그 해엔 정보가 없었다"})
    years = sorted({y for v in per_year.values() for y in v})
    for t in tickers:
        rec = {"블록": "C. 연도별 상승예측 기저대비(%p)", "구분": t}
        rec.update({str(y): per_year[t].get(y) for y in years})
        rows.append(rec)
    return pd.DataFrame(rows)


def build_industry_leader_columns(alloc: Dict[str, Any]) -> pd.DataFrame:
    """[13c 보강, v0.2.0] 부모별 일별 판단(리더/회피/폴백/폴백(여유부족)/부모ETF)·리더·여유·게이트를 13c에 나란히 싣는다."""
    eval_idx = alloc["target_w"].index
    out = pd.DataFrame(index=eval_idx)
    for p, g in alloc.get("groups", {}).items():
        out[f"{p} 판단"] = g["tier"].values
        out[f"{p} 리더"] = g["leader"].values
        out[f"{p} 회피"] = g["laggard"].values
        out[f"{p} 1위 여유"] = g["margin"].round(4).values
        out[f"{p} 여유 문턱"] = g["step"].round(4).values
        out[f"{p} 확신 게이트"] = g["gate"].values
        out[f"{p} 적격산업수"] = g["n_ok"].values
        # [v0.3.0] 신규 게이트 3종 — 왜 리더가 안 나왔는지를 13c에서 바로 읽을 수 있게
        if "regime_gate" in g:
            out[f"{p} 부모국면"] = g["parent_state"].values
            out[f"{p} 국면 게이트"] = g["regime_gate"].values
        if "rev_avoid" in g:
            out[f"{p} 역방향회피"] = g["rev_avoid"].values
        if "corr_block" in g:
            out[f"{p} 추종필터 제외"] = g["corr_block"].values
    return out

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


def industry_next_day(results: Dict[str, Dict[str, Any]]) -> Dict[str, dict]:
    """[v0.3.0 §A2] 산업별 '다음 거래일 예측' dict 모음. **새 계산이 전혀 없다** — S.build_sector_sheets가
    이미 M.build_next_day_prediction으로 만들어 둔 results[t]["sheets"]["next_day"]를 그대로 모을 뿐이다
    (01_일별_<산업> 마지막 행에 들어 있는 바로 그 값). 없으면 그 산업은 빠진다."""
    out: Dict[str, dict] = {}
    for t, r in results.items():
        nd = (r.get("sheets") or {}).get("next_day")
        if isinstance(nd, dict) and nd.get("다음거래일") is not None:
            out[t] = nd
    return out


def build_industry_prediction_matrix(results: Dict[str, Dict[str, Any]], eval_idx: pd.DatetimeIndex,
                                     icfg: IndustryConfig, S, nd_map: Optional[Dict[str, dict]] = None,
                                     alloc: Optional[Dict[str, Any]] = None) -> pd.DataFrame:
    """[01Z_산업일별예측] 날짜 × 산업 예측·목표비중. [v0.3.0 §A2] nd_map이 주어지면 '구분'(실적/예측) 열을
    넣고 맨 끝에 **예측 행 1개**를 붙인다 — S.build_prediction_matrix와 같은 관행(재계산 없음, t일 종가로
    이미 확정된 값을 표시만 재구성). 성과·불변식 계산은 '실적' 행만 쓴다(이 시트는 참고용)."""
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
    df.insert(1, "구분", "실적")
    if alloc:
        df.insert(4, "산업배분 합계", alloc["target_w"][alloc["cols"]].sum(axis=1).reindex(eval_idx).round(4).values)
    df = df.reset_index(drop=True)
    if nd_map:
        nxt = max(nd["다음거래일"] for nd in nd_map.values())
        row = {c: ("" if df[c].dtype == object else np.nan) for c in df.columns}
        row["날짜"] = pd.Timestamp(nxt).date()
        row["구분"] = "예측"
        n_u = n_d = 0
        for t, nd in nd_map.items():
            if f"{t} 예측" in row:
                row[f"{t} 예측"] = nd["확정국면"]
                row[f"{t} 목표비중"] = nd["목표비중"]
                n_u += int(nd.get("확정국면_원시") == "RISK_ON")
                n_d += int(nd.get("확정국면_원시") == "RISK_OFF")
        row["상승예측 산업수"], row["하락예측 산업수"] = n_u, n_d
        if alloc and len(alloc["target_w"].index):
            # 마지막 확정일의 산업 비중이 곧 다음 거래일에 체결할 비중(새 계산 없음)
            row["산업배분 합계"] = round(float(alloc["target_w"][alloc["cols"]].iloc[-1].sum()), 4)
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    return df


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
          "raw_parent": {p: par_px[p] for p in active_parents},       # [v0.2.0] 부모 원시 OHLC — 룩어헤드 감사 절단용
          "spy_raw_df": spy_df,                                        # [v0.2.0] SPY 원시 OHLC — 룩어헤드 감사 절단용
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
    leader_cols = pd.DataFrame()
    frozen = bool(getattr(icfg, "INDUSTRY_LAYER_FROZEN", False))
    if frozen:
        # [v0.4.0 §I6] 동결 — 배분·격자·수용기준을 만들지 않고 산업 자기국면 시트만 낸다.
        #   판정은 사람이 내린다(두 라운드 연속 13f ①⑤ FAIL + 격자 전 행 불통과). 이 플래그는 그
        #   결정을 실행에 반영할 뿐이며, 켜면 실행시간 대부분(배분·격자 백테스트)이 사라진다.
        log("START", kv(event="industry_layer_frozen", note="INDUSTRY_LAYER_FROZEN=True — 배분·격자·수용기준 생략, "
                                                            "산업 자기국면 예측 시트만 산출"), M=M, level="warning")
    if icfg.USE_ROTATION and results and not frozen:
        try:
            # [v0.4.0 §I1 ⚠] 채택 판정의 '측정 대상'을 결정과 일치시킨다.
            #   pooled  = 종전 방식(전 산업 풀링 순위 · 상장평균 벤치) — 이제 **진단으로만** 13g에 병기.
            #   within  = 부모 안 순위 · 부모 ETF 벤치 — leader3_group이 실제로 하는 판단과 1:1.
            #   ROTATION_VALIDATION_MODE로 되돌릴 수 있다("pooled"면 v0.3.1과 동일 경로).
            pooled_wf = build_pooled_rotation(results, eval_idx, icfg, M, S)
            _vmode = str(getattr(icfg, "ROTATION_VALIDATION_MODE", "within_parent")).lower()
            if _vmode == "within_parent":
                wf = within_parent_walkforward_select(results, eval_idx, icfg, M, S, pooled_wf)
                wf["alt_log"] = pooled_wf.get("alt_log", pd.DataFrame())
            else:
                wf = pooled_wf
            alloc = build_industry_allocation(results, sres, res, eval_idx, icfg, M, S, wf, rf_daily=rf_daily)
            if alloc:
                accept_df = build_industry_acceptance(alloc, wf, icfg, M, S)
                hier_df = build_hierarchy_check(alloc)
                attrib_df = build_industry_vs_sector_attribution(alloc)
                leader_cols = build_industry_leader_columns(alloc)
                log("ACCEPT", kv(event="acceptance", **{f"c{i+1}": accept_df.iloc[i]["판정"] for i in range(len(accept_df))}), M=M)
                _viol = int(hier_df["위반일수(>1e-9)"].sum()) if len(hier_df) else -1
                log("HIER", kv(event="hierarchy_check", violations=_viol,
                               max_err=float(hier_df["최대오차"].max()) if len(hier_df) else -1), M=M,
                    level=("warning" if _viol else "info"))
        except Exception as e:
            # 전체 트레이스백 꼬리를 남긴다 — v0.1 개발 중 str(e)만 남겨 원인 추적에 별도 스크립트가 필요했던 교훈.
            log("START", kv(event="allocation_failed", err=str(e)[:200],
                            trace=traceback.format_exc()[-1200:].replace("\n", " | ")), M=M, level="error")
            wf, alloc = {}, {}

    # [v0.2.0] 16_산업부모추종 — 배분과 무관하게 항상 산출(진단 전용)
    following_df = pd.DataFrame()
    if results:
        try:
            following_df = build_parent_following_analysis(results, sres, eval_idx, icfg)
            if len(following_df):
                log("DIAG", kv(event="parent_following_ready", n=len(following_df),
                               types=";".join(f"{k.split('(')[0]}:{v}" for k, v in following_df["유형"].value_counts().items())), M=M)
        except Exception as e:
            log("DIAG", kv(event="parent_following_failed", err=str(e)[:200]), M=M, level="warning")

    # [v0.3.0 §C1] 진단 시트 2종 — 13l_산업리더적중률 · 01Y_산업예측정확도(배분 실패해도 01Y는 나온다)
    leader_acc = pd.DataFrame()
    pred_acc = pd.DataFrame()
    if results:
        try:
            pred_acc = build_industry_prediction_accuracy(results, sres, eval_idx, icfg)
        except Exception as e:
            log("DIAG", kv(event="prediction_accuracy_failed", err=str(e)[:200]), M=M, level="warning")
    if alloc:
        try:
            leader_acc = build_industry_leader_accuracy(alloc, results, sres, icfg)
            _a = leader_acc[leader_acc["구분"] == "전체"] if "구분" in leader_acc.columns else pd.DataFrame()
            if len(_a):
                log("DIAG", kv(event="leader_accuracy_ready", episodes=int(_a.iloc[0].get("에피소드", 0) or 0),
                               win=_a.iloc[0].get("승률(부모대비)"), mean_excess=_a.iloc[0].get("평균초과(%)")), M=M)
        except Exception as e:
            log("DIAG", kv(event="leader_accuracy_failed", err=str(e)[:200],
                           trace=traceback.format_exc()[-600:].replace("\n", " | ")), M=M, level="warning")

    # [v0.3.0 §A2] 다음 거래일 예측 — 재계산 없음(01_일별_<산업> 마지막 행과 같은 값)
    nd_map = industry_next_day(results)
    nd_spy = sres.get("nd_spy")
    if nd_map:
        log("REPORT", kv(event="next_day_ready", n=len(nd_map),
                         target=str(max(nd["다음거래일"] for nd in nd_map.values()).date()),
                         up=sum(1 for nd in nd_map.values() if nd.get("확정국면_원시") == "RISK_ON"),
                         down=sum(1 for nd in nd_map.values() if nd.get("확정국면_원시") == "RISK_OFF")), M=M)

    universe = pd.DataFrame(universe_rows)
    matrix = build_industry_prediction_matrix(results, eval_idx, icfg, S, nd_map=nd_map, alloc=alloc)
    summary = build_industry_summary(results, failed, table, icfg)
    audit_all = pd.concat([r["audit"] for r in results.values() if isinstance(r.get("audit"), pd.DataFrame) and len(r["audit"])],
                          ignore_index=True) if results else pd.DataFrame()
    stage_timing = {"00_전체": round(time.time() - t0, 1)}

    return {
        "industries": results, "failed": failed, "selftest": st, "universe": universe,
        "quality": pd.DataFrame(quality), "matrix": matrix, "summary": summary,
        "wf": wf, "alloc": alloc, "acceptance": accept_df, "hierarchy": hier_df,
        "attribution": attrib_df, "following": following_df, "leader_cols": leader_cols,
        "leader_accuracy": leader_acc, "prediction_accuracy": pred_acc,   # [v0.3.0 §C1]
        "next_day": nd_map, "nd_spy": nd_spy,                              # [v0.3.0 §A2]
        "audit": audit_all, "icfg": icfg,
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
        # [v0.3.0 §A3 사용자 지시 "일별배분비중에 섹터가 왜 포함되어 있어 산업만 배분하라"]
        #   13c는 이제 **산업 열만** 싣는다. 부모ETF·SPY·XLU 열은 I가 새로 산 게 아니라 S★가 준 비중 중
        #   산업으로 나누지 못한 '잔여'이고, 총노출 불변식(Σ=S★)을 눈으로 확인하려면 어딘가엔 있어야 하므로
        #   13c2_잔여다리로 분리했다. 잔여를 실제로 산업에 밀어 넣으려면 INDUSTRY_ONLY_MODE="industries"
        #   (⚠ 리포트41 실측 MDD −15.04%·칼마 2.385 — 13_산업배분전략의 [잔여격자] 행에서 숫자 확인 후 결정).
        _ind_cols, _res_cols = alloc["cols"], (alloc["active_parents"] + alloc.get("passthrough_cols", []))
        tw_ind = alloc["target_w"][_ind_cols].round(4).copy()
        tw_ind.insert(0, "날짜", alloc["target_w"].index.date)
        tw_ind.insert(1, "구분", "실적")
        tw_ind.insert(2, "산업배분 합계", alloc["target_w"][_ind_cols].sum(axis=1).round(4).values)
        lc = ires.get("leader_cols", pd.DataFrame())
        if isinstance(lc, pd.DataFrame) and len(lc):        # 부모별 판단·리더·여유·게이트·국면·회피 열
            tw_ind = pd.concat([tw_ind.reset_index(drop=True), lc.reset_index(drop=True)], axis=1)
        tw_ind = tw_ind.reset_index(drop=True)
        # [§A2] 예측 행 — 마지막 확정일의 비중이 곧 다음 거래일 시가에 체결할 비중(새 계산 없음)
        nd_map = ires.get("next_day") or {}
        if nd_map and len(tw_ind):
            _nxt = max(nd["다음거래일"] for nd in nd_map.values())
            _row = tw_ind.iloc[-1].copy()
            _row["날짜"] = pd.Timestamp(_nxt).date()
            _row["구분"] = "예측(다음 거래일 체결)"
            tw_ind = pd.concat([tw_ind, _row.to_frame().T], ignore_index=True)
        sheets["13c_일별배분비중"] = tw_ind
        tw_res = alloc["target_w"][_res_cols].round(4).copy()
        tw_res.insert(0, "날짜", alloc["target_w"].index.date)
        tw_res.insert(1, "잔여 합계", alloc["target_w"][_res_cols].sum(axis=1).round(4).values)
        tw_res.insert(2, "산업 합계", alloc["target_w"][_ind_cols].sum(axis=1).round(4).values)
        tw_res.insert(3, "총노출(=S★)", alloc["target_w"].sum(axis=1).round(4).values)
        tw_res.insert(4, "S★ 총노출", alloc["w_s_all"].sum(axis=1).round(4).values)
        sheets["13c2_잔여다리"] = tw_res.reset_index(drop=True)
        sheets["13f_산업수용기준"] = accept_df
        wf = ires.get("wf", {}) or {}
        if isinstance(wf.get("selection_log"), pd.DataFrame) and len(wf["selection_log"]):
            _g = wf["selection_log"]
            # [v0.4.0 §I1] 풀링(종전 방식) 통계를 같은 행에 진단으로 병기 — "무엇이 달라졌나"를 리포트가 보여 준다.
            _pl = (wf.get("pooled") or {}).get("selection_log")
            if isinstance(_pl, pd.DataFrame) and len(_pl) and "NW-HAC t(상위1)" in _pl.columns:
                _pk = _pl[["적용연도", "신호", "NW-HAC t(상위1)", "학습 관측일"]].rename(columns={
                    "NW-HAC t(상위1)": "NW-HAC t(상위1, 풀링·종전)", "학습 관측일": "학습 관측일(풀링)"})
                _g = _g.merge(_pk, on=["적용연도", "신호"], how="left")
            alt = wf.get("alt_log")
            if isinstance(alt, pd.DataFrame) and len(alt):    # [v0.3.0 §B7] 반대쪽 타깃 t 병기
                _g = _g.merge(alt, on=["적용연도", "신호"], how="left")
            sheets["13g_산업순환매신호채택"] = _g
        la = ires.get("leader_accuracy", pd.DataFrame())
        if isinstance(la, pd.DataFrame) and len(la):
            sheets["13l_산업리더적중률"] = la                              # [v0.3.0 §C1]
        sheets["14_계층정합"] = ires["hierarchy"]
        sheets["15_산업대섹터귀속"] = ires["attribution"]
    fol = ires.get("following", pd.DataFrame())
    if isinstance(fol, pd.DataFrame) and len(fol):
        sheets["16_산업부모추종"] = fol                                  # [v0.2.0] 산업이 부모 섹터를 얼마나 따라가는가
    pa = ires.get("prediction_accuracy", pd.DataFrame())
    if isinstance(pa, pd.DataFrame) and len(pa):
        sheets["01Y_산업예측정확도"] = pa                                # [v0.3.0 §C1]
    for t, r in results.items():
        sheets[f"01_일별_{t}"] = r["sheets"].get("daily", pd.DataFrame())
    val_frames = [r["sheets"]["val_sheet"] for r in results.values() if r["sheets"].get("val_sheet") is not None
                 and len(r["sheets"]["val_sheet"])]
    if val_frames:
        sheets["03_지표검증"] = pd.concat(val_frames, ignore_index=True, sort=False)
    q, u = ires.get("quality", pd.DataFrame()), ires.get("universe", pd.DataFrame())
    sheets["10_데이터품질"] = pd.concat([u, q], ignore_index=True, sort=False) if len(q) else u
    au = ires.get("audit", pd.DataFrame())
    if isinstance(au, pd.DataFrame) and len(au):
        sheets["11_룩어헤드감사"] = au                                   # [v0.2.0] industry_lookahead_audit 결과(산업별 절단재계산)

    verdict = ("산업 계층이 S★를 이긴다(수용기준 ①~⑤ 전부 PASS)" if passed else
              "산업 계층은 S★를 이기지 못함 — 운용은 S★ 그대로(설계서 §8 관행, 진단용으로만 유지)")
    n_fail_ind = len(ires.get("failed", {}))
    au = ires.get("audit", pd.DataFrame())
    if isinstance(au, pd.DataFrame) and len(au) and "일치" in au.columns:
        n_mis = int((au["일치"] == "불일치").sum())
        n_ok = int((au["일치"] == "OK").sum())
        audit_line = f"산업별 무작위 절단 재계산 {n_ok + n_mis}건 중 불일치 {n_mis}건 — " + ("전체 통과" if n_mis == 0 else "⚠ 불일치 있음(11시트)")
    else:
        audit_line = "미실행(RUN_LOOKAHEAD_AUDIT=False 또는 표본 부족)"
    # [v0.3.1] 부모 안에서 값이 전부 같아 '산업을 고를 정보가 0'인 날이 있었던 신호(대표: PARENT_SCORE_PCT).
    #   그런 날은 복합평균·투표에서 빠진다(희석·임의 1위 방지) — 리포트가 그 사실을 숨기지 않고 한 줄로 보고한다.
    _inert: Dict[str, int] = {}
    _n_eval_days = 0
    for _p, _g in ((ires.get("alloc") or {}).get("groups") or {}).items():
        _n_eval_days = max(_n_eval_days, len(_g.get("tier", [])))
        for _k, _v in (_g.get("inert_days") or {}).items():
            _inert[_k] = max(_inert.get(_k, 0), int(_v))
    if _inert:
        inert_line = ("; ".join(f"{k} {v}일" for k, v in sorted(_inert.items(), key=lambda x: -x[1]))
                      + f" (부모별 최대, 평가 {_n_eval_days}일 중) — 해당 일자는 복합순위 평균·교차확인 투표에서 제외. "
                        "부모 계층 신호는 한 부모 아래 모든 산업이 같은 값이라 '부모 안 순위'를 만들 수 없다"
                        "(13g에서도 상위1 여유 게이트에 걸려 '표본부족'으로 남는다 — 채택되지 않는 것이 정상).")
    else:
        inert_line = "없음 — 모든 채택 신호가 부모 안에서 산업별로 값이 갈렸다"
    hier = ires.get("hierarchy", pd.DataFrame())
    hier_line = "-"
    if isinstance(hier, pd.DataFrame) and len(hier):
        hier_line = (f"위반 {int(hier['위반일수(>1e-9)'].sum())}일 · 최대오차 {float(hier['최대오차'].max()):.2e} "
                     f"(부모별 등식 + 총노출=S★ + S★ 재현 비트동일 — 14시트)")
    # [v0.3.0 §A2 사용자 지시 "다음날 예측이 없어 추가"] — 새 계산 없음. results[t]["sheets"]["next_day"]는
    #   S.build_sector_sheets가 이미 M.build_next_day_prediction으로 만들어 둔 값(01_일별_<산업> 마지막 행).
    nd_map = ires.get("next_day") or {}
    nd_spy = ires.get("nd_spy")
    nd_rows: List[Tuple[str, str]] = []
    if nd_map:
        _nxt = max(nd["다음거래일"] for nd in nd_map.values())
        _base = max(nd["기준일"] for nd in nd_map.values())
        _stale = next((nd.get("기준일_경과주의", "") for nd in nd_map.values() if nd.get("기준일_경과주의")), "")
        nd_rows.append(("다음 거래일 예측 - 기준일(데이터)", f"{pd.Timestamp(_base).date()}{_stale}"))
        nd_rows.append(("다음 거래일 예측 - 대상일", f"{pd.Timestamp(_nxt).date()} (NYSE 정규 휴장일 반영 — 임시 휴장은 미반영)"))
        if nd_spy is not None:
            nd_rows.append(("다음 거래일 예측 - SPY(M ★ 실매매 근거)",
                            f"{nd_spy['확정국면']} / 목표비중 {nd_spy['목표비중']:.2f} / {nd_spy['예상행동_kr']}"))
        for t in sorted(nd_map, key=lambda x: (results[x]["parent"], x)):
            nd = nd_map[t]
            nd_rows.append((f"다음 거래일 예측 - {t}({results[t]['parent']}, {INDUSTRY_NAME_KR.get(t, t)})",
                            f"{nd['확정국면']} / 목표비중 {nd['목표비중']:.2f} / {nd['예상행동_kr']}"))
        nd_rows.append(("다음 거래일 예측 - 안내",
                        "t일 종가로 확정된 target_pos를 t+1일 시가에 체결하는 기존 체결 규칙을 표시만 재구성한 것 — "
                        "새 계산이 아니며 13/15 등 성과 시트에는 영향 없음. 01Z_산업일별예측 마지막 행(구분=예측)·"
                        "01_일별_<산업> 마지막 행·13c 마지막 행에도 같은 값이 있음"))
    else:
        nd_rows.append(("다음 거래일 예측", "미제공(M 번들이 v1.24.0 미만이거나 계산 실패)"))
    # [§A2] '다음 거래일 배분(I★)' — 마지막 확정일의 산업 비중 = 다음 거래일 시가에 체결할 비중(새 계산 없음).
    if alloc:
        _tw_last = alloc["target_w"].iloc[-1]
        _ind_last = _tw_last[alloc["cols"]]
        _ind_last = _ind_last[_ind_last > 1e-9].sort_values(ascending=False)
        _res_cols2 = alloc["active_parents"] + alloc.get("passthrough_cols", [])
        _res_last = _tw_last[_res_cols2]
        _res_last = _res_last[_res_last > 1e-9].sort_values(ascending=False)
        _tier_txt = "; ".join(f"{p}:{g['tier'].iloc[-1]}" + (f"({g['leader'].iloc[-1]})" if g["leader"].iloc[-1] else "")
                              for p, g in alloc.get("groups", {}).items())
        nd_rows.append(("다음 거래일 배분(I★) - 산업",
                        (", ".join(f"{t} {v:.2%}" for t, v in _ind_last.items()) if len(_ind_last) else "없음(전액 잔여 다리)")
                        + f" — 산업 합계 {float(_ind_last.sum()):.2%}"))
        nd_rows.append(("다음 거래일 배분(I★) - 잔여 다리",
                        (", ".join(f"{t} {v:.2%}" for t, v in _res_last.items()) if len(_res_last) else "없음")
                        + f" — 잔여 합계 {float(_res_last.sum()):.2%} · 총노출 {float(_tw_last.sum()):.2%}"
                        f"(S★ {float(alloc['w_s_all'].iloc[-1].sum()):.2%}) · 잔여 처리 모드 {alloc.get('only_mode', 'parent')}"))
        nd_rows.append(("다음 거래일 배분(I★) - 부모별 판단", _tier_txt or "-"))
        nd_rows.append(("격자 수렴 상태(①②③④)", alloc.get("grid_line", "-")))

    meta = [
        ("버전", f"industry_rotation.py {VERSION} ({VERSION_DATE}) — sector_rotation.py {getattr(S, 'VERSION', '?')} — "
                f"market_regime_trader.py {getattr(M, 'BUNDLE_VERSION', '?')}"),
        # [v0.2.0 사용자 지시 "실제 매매에서 사용하는 전략이 뭔지 확실히 표시"] 세 리포트 공통 문구.
        ("⚠ 실매매 적용 여부", "아니오 — 이 산업 계층 리포트는 진단·연구용이며 실매매 주문에 반영되지 않는다. "
                          "실매매 주문 근거는 market_regime_report.xlsx의 ★ SPY 국면전략(00_실행요약 '다음 거래일 예측' 행). "
                          "섹터(S★)·산업(I★) 계층은 그 M 노출을 나눠 담는 연구 전략이고, 수용기준을 통과해도 사용자가 "
                          "명시적으로 전환하기 전에는 실매매에 쓰지 않는다."),
        ("판정", verdict),
        *nd_rows,
        ("룩어헤드 감사", audit_line),
        ("계층정합", hier_line),
        ("예측 대상", f"{len(results)}개 산업 ETF(부모섹터 하위) — 부모는 S.SECTOR_EXCLUDE 제외 후 산업ETF가 있는 섹터만"),
        ("실패 산업", f"{n_fail_ind}개" if n_fail_ind else "없음"),
        ("신호/백테스트 기간", f"{ires.get('signal_start')} ~ {ires.get('cal_end')}"),
        ("체결 규칙", "t일 종가에 신호 확정 → t+1일 시가 체결(M·S와 동일, 룩어헤드 구조적 차단)"),
        ("거래비용", f"산업 ETF 편도 {icfg.COST_BPS_INDUSTRY:.0f}bp · 부모 ETF 편도 {icfg.PARENT_COST_BPS:.0f}bp"),
        ("총 노출 불변식", "Σ산업비중 + 부모ETF비중 = S★의 그 섹터비중 — 14_계층정합 시트가 매일 이 등식을 검사(위반 0일이어야 함)"),
        ("⚠ 산업 배분 규칙(v0.3.0)",
         f"리더 산업에 부모비중의 {icfg.INDUSTRY_LEADER_CAP:.0%}(INDUSTRY_LEADER_CAP) · 리더 없는 날 적격 산업 균등 "
         f"{icfg.INDUSTRY_FALLBACK_SHARE:.0%}(INDUSTRY_FALLBACK_SHARE) · 잔여는 {icfg.INDUSTRY_ONLY_MODE} · "
         f"리더 인정 부모국면 {icfg.INDUSTRY_LEADER_REGIMES or '제약없음'} · 최소보유 {icfg.ROTATION_MIN_HOLD_DAYS}일 · "
         f"역방향회피 {'켬' if icfg.ROTATION_REVERSE_AVOID else '끔'} · 추종필터 {icfg.INDUSTRY_LEADER_MIN_CORR or '끔'} · "
         f"검증 타깃 {getattr(icfg, 'ROTATION_TARGET', 'ratio')}. 각 손잡이의 격자 변형은 13_산업배분전략 참조."),
        # [v0.4.0 §I1 ⚠] 채택 통계를 무엇으로 재는가 — 이번 라운드의 핵심 변경.
        ("⚠ 채택 검증 대상(v0.4.0 §I1)",
         (f"{getattr(icfg, 'ROTATION_VALIDATION_MODE', 'within_parent')} — "
          + ("**부모 안 1위 − 부모 ETF** 향후 21일 수익(leader3_group의 실제 결정과 1:1). "
             "v0.3.1까지는 '전 산업 풀링 순위 · 상장평균 벤치'였는데 그것은 부모(섹터) 간 효과를 같이 쟀다 — "
             "같은 신호가 풀링에서 t −2.4~−3.4인데 부모 안에서는 t≈+0.5로 사라졌다(그 반전은 S의 영역). "
             "13g에 두 통계를 나란히 싣는다. 되돌리기: IndustryConfig(ROTATION_VALIDATION_MODE=\"pooled\")."
             if str(getattr(icfg, "ROTATION_VALIDATION_MODE", "within_parent")).lower() == "within_parent"
             else "풀링(v0.3.1 방식) — 부모 간 효과가 섞인다"))),
        ("⚠ 순위 후보·사전방향(v0.4.0 §I1-6·§I4)",
         (f"후보 {len(icfg.ROTATION_SIGNALS)}종. **사전방향 레지스트리 신설** — v0.3.1까지 산업 순환매 신호는 "
          f"S.ROTATION_SIGNAL_SPECS에 하나도 등록돼 있지 않아 전부 +1('클수록 좋다')로 취급됐다. "
          f"P_REL_RSI_14·P_REL_MOM_21·P_REL_EXT_200·P_REL_VOL_RATIO는 경제적 사전방향이 −1인데도 반대 부호로 "
          f"검정되고 있었고, 하필 v0.3.1에서 유일하게 채택된 신호가 P_REL_RSI_14였다. 이제 S의 _RawSpec 표를 "
          f"그대로 읽어 쓴다(새 숫자 없음, 단일 출처). 13g '사전방향'·'사전방향 근거' 열에서 확인."),),
        ("⚠ 부모 안 무변동 신호(v0.3.1)", inert_line),
        ("13c 읽는 법", "13c_일별배분비중은 **산업 열만** 싣는다(사용자 지시 2026-09-12). 부모ETF·SPY·XLU 열은 "
                     "I가 새로 산 것이 아니라 S★가 준 비중 중 산업으로 나누지 못한 '잔여'이며 13c2_잔여다리에 있다. "
                     "잔여까지 산업으로 밀어 넣으려면 INDUSTRY_ONLY_MODE=\"industries\" — 그 성과는 13_산업배분전략의 "
                     "[잔여격자] 행에서 먼저 확인할 것(리포트41 실측: 잔여 전량 산업 = MDD가 가장 깊었다)."),
        ("v0.3.0 범위(⚠ 명시적 축소 — 남은 것)",
         "06c 임계값민감도·13d/13e/13h~13o 진단시트군·13h FF49외부검증(네트워크 필요)은 v0.4 예정. "
         "v0.3.0에서 해소: 부모 계층 표본 복원(§A1 — 2018~2023 리더 0일의 원인), 다음 거래일 예측(§A2), "
         "13c 산업 전용 분리·잔여 모드(§A3), 폴백 기본 0%·cap×fb 2D 격자(§B1), 부모국면 리더 게이트(§B2), "
         "산업 자기 SCORE_PCT 순위 신호(§B3), 역방향 회피(§B4), 추종필터(§B5), 베타중립 진단(§B7), "
         "13l_산업리더적중률·01Y_산업예측정확도(§C1). 수용기준 ③④는 여전히 상위1-하위1 스프레드 근사."),
        ("면책", "본 산출물은 연구·교육 목적의 백테스트 결과이며 투자 자문이 아닙니다. 과거 성과는 미래 수익을 보장하지 않습니다."),
    ]
    for k, v in ires.get("stage_timing", {}).items():
        meta.append((f"실행시간 - {k}", f"{v:.1f}초"))

    _title = "미국 산업(업종) ETF 국면 예측 & 부모 섹터 안 산업 배분 — S(섹터)→I(산업) 계층 [진단·연구용, 실매매 미적용]"
    try:
        import inspect as _inspect
        if "title" in _inspect.signature(S.write_sector_excel).parameters:
            S.write_sector_excel(path, sheets, meta, M=M, title=_title)
        else:   # 구버전 S(v0.37 이하) 호환
            S.write_sector_excel(path, sheets, meta, M=M)
    except TypeError:
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
