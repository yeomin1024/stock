# =============================================================================
#  run_pipeline.py
#  VERSION: v1.8.0 - 2026-09-14 - [문서 + 실행 레시피 — 실행 로직 무변경] S v0.47.0 → v0.48.0 · I v0.11.0 →
#                    v0.12.0 · M v1.53.1(무변경). REPORT51 구현분 — S: L1 13p 블록B·E 지평열(h=42·63·126
#                    진단) · L2 [득표상실청산격자] · L3 [보유기간격자] · L4 00시트 리더/하락국면리더 귀속
#                    분리(⚠결함수정) · L5 13c '득표 상실 중 보유' 열 · L6 11b_순환매신호감사(rot_raw 절단
#                    재계산, S·I 공통 신설). I: M1 ABS_MOM_12_1 중복 후보 제거(⚠결함수정) · M2 13p 블록P
#                    최빈 부모 몫 분모 수정(⚠결함수정) · M3 v0.11.0 CHANGELOG 정정(rot_raw는 감사 대상이
#                    아니었다) · 경로 A(REPORT51 §7 — 산업 계층 재동결, 아래 참조).
#
#  ⚠ 기본값(신호·채택·사이징) 변경 1건 — 경로 A: I.INDUSTRY_LAYER_FROZEN 기본값 **False → True**(재동결).
#    REPORT51 §3.1 K7 사전 고정 문턱 둘 다 미달(① 순서·크기를 동시에 가진 풀링 전용 후보 없음 ② 슬리브가
#    형성일 일치 대조 2행과 구별 불가)이었다 — 사전 고정 규칙대로면 No-Go(경로 A)다. 보고서 §7이 권고한
#    경로 B(사후 규칙 변경으로 게이트를 순서 우선으로 바꿔 한 번 더 시도)는 시뮬레이션조차 ★를 넘지 못했고
#    (§3.3-c: f=0.1 CAGR 35.4/칼마 3.71 vs ★ 35.9/3.76) 보고서 스스로 "성과 개선이 아니라 순서 정보의 존재를
#    기록하는 것"이라 적었다 — 사용자가 결정을 위임했고("몰라 더 나은 개선방법으로 해"), 사전 등록 원칙을
#    깨면서까지 택할 값어치가 아니라고 판단해 경로 A를 골랐다(근거 전문은 industry_rotation.py
#    IndustryConfig.INDUSTRY_LAYER_FROZEN 필드 주석). 동결이어도 13g(풀링 5열+n)·13p·16은 그대로 나온다 —
#    생략되는 것은 배분·격자·수용기준(13·13b·13c·13c2·13f·13j·13l·14·15)뿐이다(≈40분 절약).
#    되돌리기: i_overrides={"INDUSTRY_LAYER_FROZEN": False}
#    나머지(S L1~L6, I M1~M3)는 전부 진단 열·배분층 격자(라이브 기본값 `()`로 꺼짐)·문서 정정·감사 신설이라
#    ★ 라이브 성과에는 닿지 않는다 — S★·I★(within-parent)는 v0.47.0/v0.11.0과 비트 동일해야 한다.
#
#  실행 순서(REPORT51 §5 "실행 순서" — 이번에도 한 번이면 된다, 신호·사이징 기본값 변경은 위 경로 A
#    하나뿐이고 그것도 산업 배분층을 끄는 방향이라 별도 대조 실행이 필요 없다):
#        out = RP.main()
#    (참고 — 개념상 순서는 S: L4+L5+L6(측정 정합) → L1(진단) → L2+L3(격자) → I: M1+M2+M3 → 경로A, 이지만
#     전부 같은 실행 안에서 함수 호출 순서로 이미 반영돼 있다 — 별도로 여러 번 돌릴 필요는 없다.)
#
#  ★ 이번 실행에서 볼 것(REPORT51 §6 "다음 실행에서 볼 것" 표 — 항목 7·8은 경로 B 전용이라 경로 A에서는
#    해당 없음, 대신 "경로 A 확인"으로 대체):
#    1) **L1 — 13p 블록 B·E 지평 열**: 교차 리더 611일 판단을 h=21 옆에 h=42·63·126으로 재면
#       h=21 0.43 / h=63 ≈0.53 / h=126 ≈0.64 부근으로 재현되는가(±0.02). 00시트 순환매 요약 줄에 h=21·63이
#       나란히 찍히는가.
#    2) **L2 — [득표상실청산격자] 블록 E 행**: 리더 판단일 ≈586·적중 ≥ 0.44, 격자 4~5기준 통과 여부(라이브는
#       기본 꺼짐 — 판정만 본다). 13c '득표 상실 중 보유' 열(L5)이 새로 보이는가.
#    3) **L3 — [보유기간격자] 42/63 행**: 블록 E h=63 적중이 현행(21일 min_hold 기준) 0.53보다 나아지는가,
#       리더 교체 횟수가 34회보다 줄어드는가, 격자 4~5기준.
#    4) **L4 확인(⚠ 결함수정 재현)** — 00시트 순환매 요약 줄이 이제 '리더(교차)'와 '하락국면리더'로 갈려서
#       찍히는가(종전엔 두 유형의 합산값이 한 줄로 나가 귀속이 섞여 있었다).
#    5) **L6 — 11b_순환매신호감사 불일치 0건(S·I 둘 다) — 필수.** S는 build_rotation_raw_signals() 직접
#       입력만 절단재계산, I는 build_industry_rotation_raw_signals() 직접 입력만 절단재계산(둘 다 이 세션에서
#       합성데이터 단위테스트 통과 — 실제 이력으로는 미검증, 아래 [검증] 참조).
#    6) **M1 확인 — 13g 산업 풀링 열에 ABS_MOM_12_1이 더 이상 없음**(풀링 전용 후보 7종 → 6종, 학습창당
#       9→8행 또는 22→21행 부근으로 줄어드는지 확인 — 정확한 행수는 그 해 유효표본에 따라 달라질 수 있다).
#    7) **경로 A 확인** — 산업 리포트 00시트에 "⚠ 산업 계층 동결(v0.5.0 I-A)"가 **동결(기본값)**으로 찍히고,
#       13_산업배분전략·13b·13c·13c2·13f·13j·13l·14·15 시트가 **없어야** 한다(13g·13p·16·01Y·11·11b는 유지).
#    8) **★ 성과 세 계층 비트 동일 — 필수(REPORT51 §6-9)**: S★ CAGR 35.87% / 샤프 2.213 / MDD −9.55% ·
#       M★ CAGR 25.50% / 샤프 2.236 / MDD −7.07%(전부 무변경 — S L1~L6·I M1~M3+경로A는 신호·사이징을
#       건드리지 않는다). 어긋나면 즉시 알려 줄 것.
#
#  [검증] 이 세션은 실제 market_regime_trader.py(M)·시장데이터·프로젝트 정식 회귀 스위트
#    (test_sector_v0*.py·test_industry_v0*.py)가 없어 위 항목을 실제로 실행해 확인하지 못했다. 대신
#    (1) 세 파일 모두 ast.parse+import 성공 (2) sector_rotation.py 신규 로직 3종(L2 vote_loss 상태기계·L6
#    절단재계산 비교·L1 지평 열 집계)과 industry_rotation.py 신규 로직 2종(L6 builder·audit)을 합성데이터로
#    자체 단위테스트했다(세션 로컬 scratch, 모두 통과 — 인과 입력 0건 불일치, 의도적 룩어헤드 주입 시 정상
#    검출) (3) grep/코드 추적으로 M2 분모 수정이 13p 블록 P의 유일한 소스에 적용됐음을 확인했다. **실제
#    배포 전 로컬(Kaggle/Colab)에서 반드시 필요한 것**: 위 §6 표 전 항목의 실측 재현, 기존 회귀 스위트
#    전종 재실행, 스위트에 신규 기본값 가드(INDUSTRY_LAYER_FROZEN=True·ROTATION_EXIT_ON_VOTE_LOSS_GRID=()·
#    ROTATION_MIN_HOLD_GRID=()) 반영, test_sector_v046_block_e에 신규 격자 2종 존재 확인 추가.
#
#  ---- 이전 라운드 레시피(참고) ----
#  VERSION: v1.7.0 - 2026-09-13 - [문서 + 실행 레시피 — 실행 로직 무변경] S v0.47.0 · I v0.11.0 · M v1.53.1(무변경).
#                    REPORT50 구현분: J1 ⚠결함수정 13p 블록E like-for-like(라이브 611일 vs 격자 611일 분리) ·
#                    J2 [단독회피격자] 퇴역 · J5 블록B 최빈 리더(집중도) 열 · K1 ⚠결함수정 [산업슬리브격자]
#                    부모 계층 신호 구조적 제외 · K2 풀링 전용 후보 7종(S_REL_*·ABS_MOM_12_1·HAZ_PCT_OWN) ·
#                    K3 ⚠결함수정 [산업슬리브격자·대조] like-for-like(형성일 마스크) · K4 13g '풀링 학습
#                    관측일(n)' 열 · K5 13p 블록P 연도별 채택신호·최빈 부모 몫 열(xlk_share 대체) ·
#                    K6 문서 정정(REPORT49 §3.4 칼마/MDD 인공물 — 역사 기록은 보존, 정정만 병기).
#                    K7(Go/No-Go 동결 게이트, INDUSTRY_LAYER_FROZEN)은 **사용자 승인 전에는 코드로 넣지
#                    않는다** — REPORT50 §6 "결정요청(사용자)"가 문턱을 승인한 뒤 다음 라운드에서 별도 반영.
#
#  ⚠ 기본값(신호·채택·사이징) 변경 1건 — J2: S.ROTATION_AVOID_STANDALONE_GRID 기본값
#    ("order","size+order") → ()(격자 퇴역). 독립 측정 3회(13g·13p 블록E 회피→하위3 0.15·초과 +0.99%/21일,
#    전부 무작위보다 나쁨 — REPORT50 §2.3) 모두 "꼴찌를 피하라" 방향이 성립하지 않음을 확인했다. ⚠ 라이브
#    배분에는 원래도 이 격자가 관여하지 않았다(진단/격자 전용 스위치) — S★·I★ 성과는 이번에도 비트 동일.
#    되돌리기: s_overrides={"ROTATION_AVOID_STANDALONE_GRID": ("order", "size+order")}
#    나머지(J1·J5·K1·K3·K4·K5)는 전부 like-for-like 버그수정 또는 순수 진단 열 추가라 끄고 켜는 스위치가
#    없다(J1·K3은 상시 적용되는 수정, K1은 구조적 제외라 상시 적용 — 되돌리려면 코드 되돌리기만 가능).
#    K2(풀링 전용 후보 7종)만 예외적으로 끌 수 있다: i_overrides={"INDUSTRY_POOLED_ONLY_SIGNALS": ()}
#    (비우면 [산업슬리브격자]의 풀링 채택 후보가 K1 이전 상태로 좁아진다 — within-parent I★는 원래도 무관).
#
#  실행은 이번에도 한 번이면 된다(신호·사이징 기본값 변경은 격자 전용 J2 하나뿐):  out = RP.main()
#
#  ★ 이번 실행에서 볼 것(REPORT50 §5 "다음 실행에서 볼 것" 8항목):
#    1) **J1 — 13p 블록 E**: '라이브 기준' 행이 이제 611일(리더)로만 잡히고, 새로 분리된
#       '[★ 하락국면리더일]' 행이 264일·적중 0.500으로 따로 나오는가. 종전엔 라이브 행이 875일
#       (611+264 뒤섞임)로 찍혀 격자 변형 행(611일)과 like-for-like가 아니었다 — 이제 맞는가 확인.
#       그리고 블록 B '리더(교차) 판단일' 2026년 값이 **0일**로 바뀌는가(종전 "16일→1.000"은 전부
#       하락국면리더 override였다 — REPORT47~49가 인용한 그 숫자가 실은 교차 리더가 아니었다는 뜻).
#    2) **J2 확인** — 13_섹터배분전략에 [단독회피격자] 행이 **없어야** 한다(격자 행수 35 → **33행** 안팎).
#       13g '단독회피 자격(순서)' 열과 avoid_standalone_by_year 계산은 진단으로 그대로 남는다(제거 대상 아님).
#    3) **K1 — 로그 event=sleeve_excluded_parent_level**에 PARENT_SCORE_PCT 등이 찍히는가. 13g에 새 열
#       '부모 계층 제외(K1, 슬리브 채택 불가)'가 해당 이름에 Y로 표시되는가(같은 신호가 within-parent
#       채택 후보에서는 계속 그대로 쓰인다 — I★ 비트 동일 확인은 §8 참조).
#    4) **K2 — 13g 풀링 진단 5열**에 신규 후보 7종(S_REL_MOM_12_1·126·63·ABS_MOM_12_1·S_REL_DD_252H·
#       S_REL_VOL_RATIO·HAZ_PCT_OWN)이 '풀링 학습 관측일(n)'(K4)과 함께 t·부호 판정과 같이 실리는가.
#       이 신호들이 [산업슬리브격자]의 연도별 채택 후보에도 실제로 뽑히는지 로그 event=industry_sleeve_grid
#       의 'years' 필드에서 확인할 것(뽑히지 않으면 K2가 무효화됐던 종전 결함이 재발했다는 뜻).
#    5) **K3 — 13_산업배분전략 [산업슬리브격자·대조] 라벨**에 '형성일 n/N'이 찍히고, 그 n이 슬리브
#       본행([산업슬리브격자])의 실제 보유일수와 일치하는가(같은 형성일 마스크를 썼다는 증거).
#       종전(매일 f) 대조 행과 칼마·MDD를 비교해 격차가 줄어드는지 확인 — REPORT50 §3.2 기준선:
#       종전 대조 칼마가 실제보다 부풀려져 있었다(K6 참조).
#    6) **K5 — 13p 블록 P**: 연도별 행에 '채택 신호(그 해)' 열, '── 전체 ──' 행에 '형성일/전체'·
#       '최빈 부모 몫' 열이 새로 보이는가. REPORT50 §3.1 기준선(엔진 재현, PARENT_SCORE_PCT 시절):
#       형성 655일 중 IYZ(XLC) 570일·FDN(XLC) 549일·SOCL(XLC) 444일 — 셋 다 XLC 소속이라 '최빈 부모 몫'이
#       **크게** 나와야 종전 xlk_share=0.0(이 셋을 놓쳤던 결함, §4 E6)이 재현·수정됐다는 뜻이다. K1·K2
#       적용 후에는 채택 후보가 바뀌므로 이 구체 수치 자체가 재현될 필요는 없다 — '몫이 여전히 한 부모에
#       쏠려 있는가'만 확인. 상위3 적중이 무작위(3/29≈0.103)를 **넘는지**가 판정 기준 ②(K7 사전 조건 중 하나).
#    7) **11_룩어헤드감사 0건** — K2로 새로 추가된 rot_raw 키(S_REL_*·ABS_MOM_12_1·HAZ_PCT_OWN)도 절단재계산
#       대상에 자동 등록되는가(신규 항목이 감사망 밖에 있으면 안 된다).
#    8) **★ 성과 세 계층 비트 동일** — M v1.53.1(무변경) · S★ CAGR 35.87%/샤프 2.213/MDD −9.55%(무변경,
#       J1·J2·J5 전부 진단/격자 전용) · I★ 무변경(INDUSTRY_SLEEVE_SHARE=0.0 그대로, K1~K5는 슬리브 채택
#       후보군·진단열만 건드린다 — within_parent_walkforward_select의 rank_full·ROTATION_SIGNALS는
#       손대지 않았다). 어긋나면 즉시 알려 줄 것.
#
#  ⚠ K7(Go/No-Go 동결 게이트) — REPORT50 §6 "결정 요청(사용자)". **이번 실행이 자동으로 판정하지 않는다.**
#    이번 실행 결과를 아래 문턱과 사람이 대조해 승인한 뒤에만, 다음 라운드에서 INDUSTRY_LAYER_FROZEN 반영을
#    코드로 넣는다(사후 적합 방지 — 문턱은 이 라운드 실행 전에 사전 고정한다):
#      ① 풀링 전용 후보(K2) 중 풀링 t ≥ 2.0이 9개 학습창 중 ≥ 5개 **그리고** 풀링 1위→실현 상위3 ≥ 0.15인
#         신호가 1개 이상 있는가 (13g 풀링 5열 + K4 'n' 열로 판독)
#      ② [산업슬리브격자] 후보 행이 **형성일 일치 대조 2행 모두**(K3)에 칼마·MDD로 이기고, 블록 P(K5)
#         상위3 적중 > 무작위(3/29≈0.103)인가
#    ①②를 둘 다 만족하지 못하면 `INDUSTRY_LAYER_FROZEN=True`(13g·13p·16 진단 시트는 유지, 배분·격자만
#    생략 — 실행시간 ≈40분 절약)로 되돌리고 산업 배분 격자 라운드를 멈춘다(사용자 승인 필요 — REPORT50 §6
#    질문 1·2). 동결 되돌리기: `i_overrides={"INDUSTRY_LAYER_FROZEN": False}`.
#
#  ---- 이전 라운드 레시피(참고) ----
#  VERSION: v1.6.0 - 2026-09-13 - [문서 + 실행 레시피 — 실행 로직 무변경] S v0.46.0 · I v0.10.0 · M v1.53.1.
#                    REPORT49 구현분: H1 ★★ 독립 산업 슬리브 · H2 산업 13g 풀링 진단 5열 ·
#                    H3 within-parent 격하 · H4 13p 블록 E(격자 변형별) · H5 [대피처격자] 퇴역 ·
#                    H7 M 00시트 버전 문자열.
#
#  ★ 이번 라운드의 본체는 **산업 계층의 구조 전환(H1)**이다. 지금까지 산업은 "S★가 준 섹터 비중 **안에서**
#    부모 ETF를 대체"하는 새장이었고(S★가 XLK에 90%를 주므로 무대는 사실상 'SOXX vs XLK' 하나),
#    그래서 세 라운드 연속 배분 0일이었다. 이제 산업에 **자체 예산 f × S★ 총노출**을 주고 29산업을
#    **부모 무관 공동 순위**로 고른다. ⚠ 라이브는 꺼져 있어(f=0.0) I★는 v0.9.0과 비트 동일 — 격자만 늘어난다.
#    실행은 한 번이면 된다(⚠ 신호·사이징 기본값 변경 0건):  out = RP.main()
#
#  ★ 이번 실행에서 볼 것(판정 기준선은 리포트45/8 실측 + REPORT49 §3.4 근사):
#    1) **H1 — 13_산업배분전략 [산업슬리브격자] 4행 + [산업슬리브격자·대조] 6행.**
#       ⚠ **대조군과 짝지어 읽는 것이 전부다.** 같은 f를 SPY에 넣은 행보다 칼마가 높고 MDD가 낮아야 채택
#       후보다 — 산업 슬리브의 값은 '수익'이 아니라 '★와 다른 것을 든다'(일간 상관 0.63)에서 온다.
#       REPORT49 §3.4 근사 기준선: f=0.3 → 칼마 3.06 vs SPY 혼합 대조 2.35 · MDD −9.6% vs −12.1%.
#       (실제 엔진 값은 다를 수 있다 — 방향만 같으면 된다. ★ 재현이 31.4%였던 근사라서.)
#       로그: event=industry_sleeve_grid share=… k=… holdings=GDX(XLB)…;XME(XLB)… xlk_share=…
#             **xlk_share가 크면 S★ XLK 편향의 재탕**이지 새 정보가 아니다(REPORT49 §3.4 실측 0.189).
#    2) **H1 판정 ② — 13p 블록 P(풀링 슬리브 소수클래스).**
#       '슬리브 → 실현 상위k 비율'이 무작위(k/N ≈ 0.103)를 넘어야 한다. 넘지 못하면 CAGR이 좋아도
#       그건 분산 효과일 뿐이므로 채택하지 않는다(REPORT49 §3.3: 풀링 상위3 적중 0.17 — 넘긴 하지만 약하다).
#    3) **H2 — 13g 풀링 진단 5열, 특히 '풀링 사전방향 판정'.**
#       v0.9.0 I-F가 연 P_REL_VOL_RATIO는 산업 풀링에서 **부호가 뒤집힌 신호**다(저변동 1위의 상위3 적중
#       0.050 · vs SPY −0.87%/21일 t −4.1) — 그래서 졌다. 이번 실행에서 그 신호가 **'반전(사전방향과 반대)'**
#       으로 찍히는지 확인할 것. 찍히면 H2가 제 일을 한 것이고, 앞으로 그런 신호는 채택 전에 걸러진다.
#       로그: event=within_parent_select … pooled_reversed=P_REL_VOL_RATIO,…
#    4) **H4 — 13p 블록 E(격자 변형별 소수 클래스).**
#       리포트45에서 [단독회피격자]는 CAGR로만 기각됐고 '회피 판단이 무작위를 넘었는가'는 **잴 수 없었다**
#       (블록 B가 라이브 배분만 세는데 라이브는 회피 0일). 이제 블록 E가 격자 행의 회피일·적중·집중도를 낸다.
#       볼 것: [단독회피격자] 행의 '회피 섹터 → 실현 하위3'이 무작위(0.273)를 넘는가, '최빈 회피(집중도)'가
#       작은가. 넘고 분산돼 있으면 **CAGR이 져도 '꼴찌 예측은 성립'으로 기록**하고 다음 설계에 쓴다.
#    5) **H5 확인** — 13 시트에 [대피처격자] 행이 **없어야** 한다(격자 38 → 33행 안팎). 13p 블록 D·13m 블록 D는 유지.
#    6) **H3 확인** — 13_산업배분전략이 ★ + 대조군 + [산업슬리브격자]만 남아 짧아진다(within-parent 28행 생략).
#       ⚠ 계산 생략일 뿐이라 **I★ 성과는 무변경**이어야 한다.
#    7) **H7 확인** — market_regime_report 00시트 '버전' 줄이 **v1.53.1**로 찍힌다(리포트66은 v1.51.0로 틀렸다).
#    8) **★ 성과 불변** — S★ CAGR 35.87% / 샤프 2.213 / MDD −9.55%, I★ 동일. 어긋나면 즉시 알려 줄 것.
#
#  ⚠ v0.46.0 / v0.10.0 되돌리기 한 줄:
#    i_overrides={"INDUSTRY_SLEEVE_GRID": ()}                           # H1 슬리브 격자 끄기
#    i_overrides={"INDUSTRY_SLEEVE_SHARE": 0.3, "INDUSTRY_SLEEVE_K": 3} # ⚠⚠ H1 라이브 전환(판정 통과 뒤에만)
#    i_overrides={"INDUSTRY_GRID": True}                                # H3 within-parent 격자 복원
#    i_overrides={"INDUSTRY_LEADER_STANDALONE_GRID": (0.03, 0.07)}      # I-F 격자 복원(권하지 않음 — 반전 신호)
#    s_overrides={"ROTATION_SHELTER_RULE_GRID": ("votes2","votes2_nonup","leader_only")}   # H5 복원
#    s_overrides={"GRID_CRITERION_PRIMARY_BIAS": False}                 # 격자 ⑤ 끄기(종전 4기준)
#
#  ---- 이전 라운드 레시피(참고) ----
#  VERSION: v1.5.0 - 2026-09-13 - [문서 + 단계별 실행 레시피 — 실행 로직 무변경] S v0.45.0 · I v0.9.0 · M v1.53.0.
#                    REPORT48 구현분: G1 격자 ⑤ 버그수정 · G2 13p 블록 D(득표×국면) · G3 [대피처격자] ·
#                    G4′ 유효표본 게이트 사전등록(판독서 G4는 철회) · G5 [단독회피격자] · G6 F1 퇴역 ·
#                    G7 [국면상속격자] 퇴역 · G9 M 리포트 동반 업로드.
#                    + I-F(사용자 지시 "왜 industry는 수정 안했어?"): 산업 [단독리더격자] — 산업 배분이
#                      세 라운드 연속 0일인 원인이 신호 품질이 아니라 **정족수**였다(채택 0~1개 vs 교차확인 2).
#                      순서 잣대 여유로 교차확인을 면제하는 격자. 라이브는 꺼짐 → I★ 비트 동일.
#
#  ★ 이번 라운드는 ⚠ 기본값(신호·채택·사이징) 변경이 **0건**이다 — 라이브 ★는 v0.44.0과 비트 동일하고,
#    늘어난 것은 13 시트의 격자 행과 13p/13m/13f/00의 진단뿐이다. 그래서 단계를 쪼갤 필요 없이 한 번 돌린다:
#        out = RP.main()
#    그리고 **market_regime_report.xlsx도 같이 업로드**할 것(G9) — F5(M 13p)가 두 라운드째 미판정이다.
#
#  ★ 이번 실행에서 볼 것(판정 기준선은 리포트44 실측):
#    1) **G1 — 00시트 '격자 수렴 상태(①②③④⑤)' 줄.**
#       v0.44.0은 ★ 라벨을 alloc에서 찾았는데 그 키는 alloc["diag"] 안에 있어 ⑤가 **한 번도 작동하지 않았다**
#       (리포트44: "⑤ 통과 38/38" = 전부 통과가 아니라 전부 미검사). 이번 실행에서
#         · "⑤ 통과 n"이 **38 미만**이어야 하고
#         · "⑤에서 걸러진 행"에 [상한격자] XLK 100%·M헤어컷 E9/E11 같은 '주력을 더·오래' 행이 나와야 하며
#         · 13_섹터배분전략에 **'주력 평균비중' 열**이 새로 보여야 한다.
#       끄기(종전 4기준): s_overrides={"GRID_CRITERION_PRIMARY_BIAS": False}
#    2) **G2 — 13p 블록 D(득표 × M 국면)** · 00시트 '소수 클래스 정확도' 줄 끝 · 13f ⑪.
#       리포트44 재계산 기준선: 0표 0.182 / 1표 0.296 / 2표 0.396 / 3표 0.821 ·
#       득표≥2·중립 0.565 vs 상승 0.381 · **득표≤1·상승 0.227(초과 −0.71%/21일, 644일)**.
#       이 표가 [대피처격자] 판정의 근거다 — 득표≤1·상승이 무작위(0.27) 아래면 현행 슬리브 규칙이 손해다.
#    3) **G3 — 13_섹터배분전략 [대피처격자] 3행 + 13m 블록 D.**
#       기준선(리포트44 13m): 슬리브 실제 배분 10.64% vs 비주력 균등 10.38%(1,507일) — 현행 판단이 균등과 같다.
#       판정: 격자 4기준(⑤ 포함) **그리고** 13m 블록 D에서 그 규칙의 누적 기여가 '비주력 균등'을 넘는가.
#       CAGR 기대치는 작다(누적 +1.4%p 안팎) — 목적은 수익이 아니라 '판단을 정보 있는 날로 제한'이다.
#       끄기: s_overrides={"ROTATION_SHELTER_RULE_GRID": ()}
#    4) **G5 — [단독회피격자] + 13g '단독회피 자격(순서)' 열 + 13p ⑧.**
#       ⚠ 사전 예상 정정: 리포트44 13g 값으로 미리 돌려 보니 "order" 자격이 **2021·2023·2024·2026** 네 해에
#       선다(판독서는 2024·2026만 예상했다). 2021·2023은 모멘텀 계열의 '최근 252일 꼴찌→하위3'이 0.44~0.63인데
#       그 창은 **2020년 XLE가 1년 내내 바닥이던 구간**이다 — REPORT48 §2.3이 "한 섹터 고정이면 순환매 예측이
#       아니다"라고 배제한 바로 그 인공물이다. 문턱은 사후 적합을 피하려 **그대로 두고**, 대신 로그에
#       집중도 진단을 넣었다:
#         event=avoid_standalone_grid variant=order ... avoid_sectors=XLE=203;... top_share=0.62
#       **top_share가 크면 그 해는 '그 섹터를 빼는 규칙'**이다. 판정은 13p ⑧(발동일 꼴찌→실현 하위3이
#       무작위 0.27을 넘는가)과 top_share를 **같이** 본다. 넘지 못하면 하위 쪽은 이 시스템에서 닫는다.
#       끄기: s_overrides={"ROTATION_AVOID_STANDALONE_GRID": ()}
#    5) **G6·G7 퇴역 확인** — 13 시트에 [역방향회피격자]·[국면상속격자] 행이 **없어야** 한다(격자 38 → 33행 안팎).
#       13g '역방향 회피 자격' 열과 13p 블록 A 'M상속' 6열은 진단으로 그대로 남는다.
#    6) **I-F(산업) — 13_산업배분전략 [단독리더격자] 2행 + 13g '단독리더 여유(순서−무작위)' 열.**
#       배경: 산업 13g 연도별 채택 신호 수가 2018:1 2019:1 2020:0 2021:1 2022:1 2023:0 2024:0 2025:0 2026:1로
#       0~1개뿐이라 교차확인(ROTATION_MIN_AGREE=2)이 구조적으로 불가능했고, 그래서 산업 배분이 세 라운드
#       연속 **0일**이며 13p 블록 B가 **1행**(잴 것이 없음)이었다.
#       채택 신호 P_REL_VOL_RATIO의 순서 잣대 여유는 다섯 해 모두 +0.069~+0.101(무작위 0.327 대비)이다.
#       **판정은 CAGR이 아니다** — 13p 블록 B('부모 안 리더 → 부모 안 실현 1위')가 무작위(1/부모안 산업수)를
#       넘느냐, 그리고 로그 top_share(한 산업 쏠림)가 작으냐를 본다:
#         event=leader_standalone_grid edge=0.03 standalone_leader_days=… leaders=SOXX=…;… top_share=…
#       솔직한 기대치는 낮다(REPORT48 §5: 평가창 실현 적중 0.35 vs 무작위 픽 0.33 · 부모 초과 t ≥2 없음).
#       그래도 재는 이유는 **0일이면 영원히 판정이 안 나기 때문**이다 — 못 넘으면 그때 근거를 갖고 닫는다.
#       끄기: i_overrides={"INDUSTRY_LEADER_STANDALONE_GRID": ()}
#    7) **G4′** — 13g '유효표본 n_eff' 열을 볼 것. 2019 REL_DD_252H가 427.7(채택 33행 중 최소, 차순위 802.8)이다.
#       ⚠ 다만 **빼도 2019는 나아지지 않는다**(원자료 재검: 상위3 0.198→0.202, 하위3 0.294→0.397,
#       초과 −0.37→−0.72%) — 판독서 G4(최소 학습창)는 그래서 철회했다. 굳이 재 보려면:
#         s_overrides={"ROTATION_MIN_N_EFF": 600.0}
#
#  ⚠ v0.45.0 되돌리기 한 줄(전부 격자·진단이라 ★ 성과에는 닿지 않는다):
#    s_overrides={"ROTATION_SHELTER_RULE_GRID": ()}                     # G3 대피처 격자 끄기
#    s_overrides={"ROTATION_SHELTER_RULE": "votes2"}                    # ⚠ G3 라이브 전환(4기준 통과 뒤에만)
#    s_overrides={"ROTATION_AVOID_STANDALONE_GRID": ()}                 # G5 단독회피 격자 끄기
#    s_overrides={"ROTATION_AVOID_STANDALONE_HIT": (0.40, 0.45)}        # G5 문턱 올리기(인공물 배제 실험)
#    s_overrides={"ROTATION_REVERSE_AVOID_GRID": (2.0, 2.5)}            # G6 부활(권하지 않음 — 근거가 뒤집혔다)
#    s_overrides={"ROTATION_REGIME_INHERIT_GRID": ("m_off", "m_up")}    # G7 부활
#    s_overrides={"GRID_CRITERION_PRIMARY_BIAS": False}                 # G1 ⑤ 끄기(종전 4기준)
#    s_overrides={"ROTATION_MIN_N_EFF": 600.0}                          # G4′ 유효표본 게이트 켜기
#    i_overrides={"INDUSTRY_LEADER_STANDALONE_GRID": ()}                # I-F 단독리더 격자 끄기
#    i_overrides={"INDUSTRY_LEADER_STANDALONE_EDGE": 0.03}              # ⚠ I-F 라이브 전환(4기준 통과 뒤에만)
#
#  ---- 이전 라운드 레시피(참고) ----
#  VERSION: v1.4.0 - 2026-09-13 - [문서 + 단계별 실행 레시피 — 실행 로직 무변경] S v0.44.0 · I v0.8.0 · M v1.53.0.
#                    REPORT47 구현분: F1 역방향 회피 격자 부활 · F2 ⚠ 최선가용을 순서 잣대로 · F3 M 상속
#                    (13p 비교열 + [국면상속격자]) · F4 ★ 13q 하락확률보정 · F5 ★ M 13p · F6 산업 벤치 진단·격자 플래그 ·
#                    F7 격자 기준 ⑤(주력 편향).
#
#  ★ 단계별 실행 레시피(REPORT47 §7 — 한 번에 하나):
#    1) **F2 단독 판정**(⚠ 채택층): out = RP.main()
#       이번 기본 실행에서 ⚠는 F2 하나다(ROTATION_BEST_AVAILABLE_RANK_BY="hit"). 리포트43 13g로 재현한
#       영향 범위는 **2020년 한 해**뿐이다(엄격 신호가 없어 최선가용이 발동하는 유일한 해 중 선택이 갈리는 해):
#         v0.43 t 순 → REL_DD_252H(t 1.46·순서 0.309) + SCORE_CS_Z  → 리더 12일 · 초과 −0.7%p
#         v0.44 순서 순 → SCORE_PCT(0.319) + SCORE_CS_Z(= v0.42.1 조합) → 리더 161일 · 초과 +10.5%p(기대)
#       볼 것: 13p 블록 B **2020 리더 판단일·리더→상위3**(기준선 12일/0.083) · 13f ③(1.62 → 2.0 복귀하는가) ·
#              2026은 무변경이어야 한다(16일/1.000 유지 — 그 해는 최선가용이 발동하지 않는다).
#       되돌리기 대조: out = RP.main(s_overrides={"ROTATION_BEST_AVAILABLE_RANK_BY": "t"})
#    2) 같은 실행에서 함께 나오는 진단·격자(신호 무변경 — 따로 돌릴 필요 없다):
#       · 13p 블록 B '역방향 회피 대상 → 실현 상위3'(기준선: 상대변동성 최저 섹터 0.147) + 13_섹터배분전략
#         **[역방향회피격자] 2행** → 4기준(+⑤) 통과하면 그때 라이브 스위치를 켠다(ROTATION_REVERSE_AVOID).
#       · 13p 블록 A 'M상속(M현금/M상승아님) MCC'와 'M상속 우위(MCC차)'(기준선: 자기 0.055 · M현금 0.079 ·
#         M상승아님 0.088) + **[국면상속격자] 2행**.
#       · **13q_하락확률보정** — 신뢰도 기울기 ≥ 0.10이면 방향 정보가 있다는 뜻이고, 그때 비로소 문턱을
#         기저율에 맞춰 재현율 한계(라벨 방식 ~0.074)를 풀 수 있다. **배분 미사용**.
#       · M 리포트 **13p_소수클래스정확도**(기준선: SPY 현금 MCC +0.160 · 상승아님 +0.187).
#       · 격자 수렴 줄이 **①②③④⑤**로 — ⑤에서 걸러진 행(주력 편향을 키우는 행)을 함께 보여 준다.
#       · I 13g 진단 4열(중앙값 벤치 t · 부모 안 실현1위 비율) — 부모 초과 타깃의 구조적 편향(0.45~0.49)이
#         원인인지, 산업 순위 자체에 정보가 없는지를 가른다.
#    3) 그 다음 라운드 후보(각각 단독):
#       s_overrides={"ROTATION_STRICT_REQUIRE_UNIFORM_T": 1.0}      # R4 균등 t 하한
#       s_overrides={"SECTOR_HAZARD_CONFIRM_GATE": True}            # S-N
#       i_overrides={"ROTATION_VALIDATION_BENCH": "group_median"}   # F6(c) 타깃 전환
#
#  ⚠ v0.44.0/v0.8.0 기본값과 되돌리기 한 줄:
#    S: s_overrides={"ROTATION_BEST_AVAILABLE_RANK_BY": "t"}   # ⚠ F2 되돌리기(v0.43.0 동작)
#       s_overrides={"ROTATION_BEST_AVAILABLE_MIN_HIT": 0.27}  # F2-b 켜기(기본 None)
#       s_overrides={"ROTATION_REVERSE_AVOID_GRID": ()}        # F1 격자 끄기
#       s_overrides={"ROTATION_REVERSE_AVOID": True}           # ⚠ F1 라이브 적용(격자 통과 뒤에만)
#       s_overrides={"RUN_DOWN_PROB": False}                   # F4 13q 끄기(실행시간 절약)
#       s_overrides={"GRID_CRITERION_PRIMARY_BIAS": False}     # F7 끄기(종전 4기준)
#       s_overrides={"ROTATION_RECENT_MIN_T": None}            # ⚠ R2 되돌리기
#    I: i_overrides={"INDUSTRY_GRID": False}                   # F6(a) 격자 28행 생략(13·13c·14는 그대로)
#       i_overrides={"ROTATION_VALIDATION_BENCH": "group_median"}  # ⚠ F6(c) 검증 타깃 전환
#       i_overrides={"INDUSTRY_LAYER_FROZEN": True}            # ⚠ 동결(배분 시트 생략)
#
#  VERSION: v1.0.1 - 2026-09-12 - [문서만 변경 — 실행 로직·기본값 무변경] S v0.39.0에서 SECTOR_EXCLUDE
#                    기본값이 ("XLB","XLE") → ()(11섹터 전부 예측)로 돌아갔다(사용자 지시 "sector는 다시
#                    XLE, XLB 같이 예측"). 아래 사용 예시의 기본값 설명과 되돌리기 한 줄을 그에 맞게 고쳤다.
#                    main(sector_exclude=...)의 동작 자체는 그대로다(None이면 S 기본값 사용).
#  VERSION: v1.0.0 - 2026-09-12 - M(시장국면) → S(섹터) → I(산업) 한 번에 실행하는 노트북 셀용 러너.
#                    Colab / Kaggle / 로컬 공용 — 실행 환경을 M.runtime_env()로 판별해 출력·캐시 경로를 정한다.
#
#  왜 별도 파일인가: 실행 셀에 흩어져 있던 "wget → 모듈 로드 → run → report → download"를 한 곳에 모아
#    (1) Kaggle에서는 /kaggle/working 아래(노트북 Persistence(Files) 설정 시 세션 간 유지, 커밋 시 Output 저장)에
#        리포트·캐시를 두고 날짜 폴더(reports/YYYY-MM-DD/)로 이력을 쌓는다 — 사용자 지시 "엑셀 결과도 kaggle
#        영구 저장소에 보관".
#    (2) Colab에서는 종전처럼 zip 1개로 자동 다운로드.
#    (3) 실행 끝에 **실매매 적용 전략 배너**(M ★ SPY 국면전략의 다음 거래일 목표비중·예상 행동)를 찍고,
#        S★·I★는 진단용임을 함께 표시한다 — 사용자 지시 "실제 매매에서 사용하는 전략이 뭔지 확실히 표시".
#    (4) 캐시가 세션 간 살아남으면(Kaggle Persistence) 두 번째 실행부터 S·I가 산업/섹터당 수 초로 끝난다.
#
#  사용(노트북 셀):
#      !wget -q -O run_pipeline.py https://raw.githubusercontent.com/yeomin1024/stock/main/run_pipeline.py
#      ... (market_regime_trader.py / sector_rotation.py / industry_rotation.py 도 같은 방식으로 받는다)
#      import run_pipeline as RP
#      out = RP.main()                                    # 기본: 11섹터 전부 예측(SECTOR_EXCLUDE=()), 산업 계층은
#                                                         #   v0.5.0부터 **동결**(진단 시트만 — 배분·격자 생략)
#      out = RP.main(sector_exclude=("XLB", "XLE"))       # ⚠ 되돌리기: 9섹터로 다시 좁힌다(S v0.37~v0.38 상태)
#      out = RP.main(sector_exclude=("XLB",))             # ⚠ XLB만 제외(XLE는 예측)
#      out = RP.main(run_industry_layer=False)            # 산업 계층 자체를 생략(M+S만 — 리포트 2개)
#      out = RP.main(i_overrides={"INDUSTRY_LAYER_FROZEN": False})   # ⚠ 산업 배분·격자 되살리기(I-B 실험 포함)
#
#  Kaggle 노트북 설정(오른쪽 패널): Internet = On, Persistence = "Files only"(또는 Variables & Files),
#    Accelerator = None(CPU 4코어면 fork 병렬 4워커). Environment는 최신(Pin to original 아님) 권장.
#
#  ※ 본 코드는 연구/교육용 도구이며 투자 자문이 아니다.
# =============================================================================
from __future__ import annotations

import os
import sys
import shutil
import time
import dataclasses
import datetime as dt
import importlib.util
from typing import Any, Dict, Optional, Tuple

VERSION = "v1.8.0"
VERSION_DATE = "2026-09-14"

MODULE_FILES = {
    "market_regime_trader": "market_regime_trader.py",
    "sector_rotation": "sector_rotation.py",
    "industry_rotation": "industry_rotation.py",
}


def load_module(name: str, path: Optional[str] = None):
    """이전 세션에서 캐싱된 낡은 모듈을 버리고 파일에서 새로 로드한다. ★ exec 전에 sys.modules에 등록 —
    dataclass/피클/fork 자식이 모듈 이름으로 다시 찾을 때 실패하지 않게(사용자 셀의 기존 패턴 그대로)."""
    path = path or MODULE_FILES[name]
    if not os.path.exists(path):
        raise FileNotFoundError(f"{path} 가 없습니다 — 셀 상단의 wget(GitHub raw)이 실패했는지 확인하세요")
    sys.modules.pop(name, None)
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


def detect_env(M=None) -> str:
    if M is not None and hasattr(M, "runtime_env"):
        return M.runtime_env()
    if os.environ.get("KAGGLE_KERNEL_RUN_TYPE") or os.path.isdir("/kaggle/working"):
        return "kaggle"
    try:
        import google.colab  # type: ignore  # noqa: F401
        return "colab"
    except Exception:
        return "local"


def base_dir_for(env: str) -> str:
    """리포트·캐시가 놓일 기준 폴더. Kaggle은 /kaggle/working(영구 저장 대상), Colab은 /content, 로컬은 cwd."""
    if env == "kaggle":
        return "/kaggle/working"
    if env == "colab" and os.path.isdir("/content"):
        return "/content"
    return os.getcwd()


def _banner(lines) -> None:
    width = max(len(l) for l in lines) + 4
    print("=" * width)
    for l in lines:
        print(f"  {l}")
    print("=" * width)


def main(sector_exclude: Optional[Tuple[str, ...]] = None, run_industry_layer: bool = True,
         download: bool = True, keep_history: bool = True,
         m_overrides: Optional[Dict[str, Any]] = None, s_overrides: Optional[Dict[str, Any]] = None,
         i_overrides: Optional[Dict[str, Any]] = None, base_dir: Optional[str] = None,
         _hooks: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """M → S → I 실행 + 리포트 + (Colab) 다운로드 / (Kaggle) 영구 보존 + 실매매 배너.
    sector_exclude: None이면 sector_rotation.py의 기본 그대로 — v0.39.0부터 기본은 ()(11섹터 전부 예측).
        ⚠ 9섹터로 되돌리려면 sector_exclude=("XLB","XLE"). 제외는 신호·배분·성과를 바꾸는 설정이다.
    *_overrides: 각 Config 필드 덮어쓰기(dataclasses.replace) — 예: s_overrides={"MAX_WORKERS": 2}.
    _hooks: 테스트 전용 — {"m_run": f(mcfg)->res, "s_run": f(res,M,scfg)->sres, "i_run": f(sres,res,M,S,icfg)->ires}
            (네트워크 없는 샌드박스에서 합성데이터로 러너 전체 경로를 검증하기 위한 주입점)."""
    t_all = time.time()
    M = load_module("market_regime_trader")
    S = load_module("sector_rotation")
    I = load_module("industry_rotation") if run_industry_layer else None
    print("M.VERSION:", getattr(M, "BUNDLE_VERSION", "??"), "| S.VERSION:", getattr(S, "VERSION", "??"),
          "| I.VERSION:", (getattr(I, "VERSION", "??") if I is not None else "(생략)"), "| runner:", VERSION)
    assert hasattr(S, "run"), "S.run이 없음 - GitHub에 올린 sector_rotation.py를 다시 확인하세요"
    if I is not None:
        assert hasattr(I, "run"), "I.run이 없음 - GitHub에 올린 industry_rotation.py를 다시 확인하세요"

    env = detect_env(M)
    base = base_dir or base_dir_for(env)
    os.makedirs(base, exist_ok=True)
    print(f"[runner] 실행 환경 = {env} | 기준 폴더 = {base}")

    # ---- 경로 배치: 리포트·CSV·번들·캐시 전부 기준 폴더 아래(Kaggle이면 영구 저장 대상) ----
    mcfg = dataclasses.replace(
        M.CFG, CACHE_DIR=os.path.join(base, "cache_market_data"),
        OUT_XLSX=os.path.join(base, "market_regime_report.xlsx"),
        RESULT_BUNDLE_PATH=os.path.join(base, "market_regime_result.pkl.gz"),
        DAILY_CSV_PATH=os.path.join(base, "market_regime_daily.csv"),
        **(m_overrides or {}))
    s_kw = dict(CACHE_DIR=os.path.join(base, "cache_sector"),
                OUT_XLSX=os.path.join(base, "sector_regime_report.xlsx"),
                DAILY_CSV_PATH=os.path.join(base, "sector_regime_daily.csv"),
                ALLOC_CSV_PATH=os.path.join(base, "sector_allocation_daily.csv"))
    if sector_exclude is not None:
        s_kw["SECTOR_EXCLUDE"] = tuple(sector_exclude)
    s_kw.update(s_overrides or {})
    scfg = dataclasses.replace(S.CFG, **s_kw)
    icfg = None
    if I is not None:
        i_kw = dict(CACHE_DIR=os.path.join(base, "cache_industry"),
                    OUT_XLSX=os.path.join(base, "industry_regime_report.xlsx"),
                    DAILY_CSV_PATH=os.path.join(base, "industry_daily.csv"),
                    ALLOC_CSV_PATH=os.path.join(base, "industry_allocation_daily.csv"))
        i_kw.update(i_overrides or {})
        icfg = dataclasses.replace(I.CFG, **i_kw)
    print(f"[runner] 섹터 제외(SECTOR_EXCLUDE) = {tuple(getattr(scfg, 'SECTOR_EXCLUDE', ()) or ()) or '없음'}")

    # ---- 실행 ----
    hk = _hooks or {}
    res = (hk.get("m_run") or M.run)(mcfg)
    path = M.build_report(res, mcfg)
    sres = (hk.get("s_run") or S.run)(res, M, scfg)
    path2 = S.build_sector_report(sres, M=M)
    ires, path3 = None, None
    if I is not None:
        ires = (hk.get("i_run") or I.run)(sres, res, M, S, icfg)
        path3 = I.build_industry_report(ires, M=M, S=S)
    paths = [p for p in (path, path2, path3) if p and os.path.exists(p)]
    print("생성 완료:", " / ".join(paths))

    # ---- 이력 보관(reports/YYYY-MM-DD/) — Kaggle Persistence·커밋 Output에 그대로 남는다 ----
    hist_dir = None
    if keep_history:
        stamp = str(res["cal"][-1].date()) if isinstance(res, dict) and "cal" in res else dt.date.today().isoformat()
        hist_dir = os.path.join(base, "reports", stamp)
        os.makedirs(hist_dir, exist_ok=True)
        for p in paths:
            shutil.copy2(p, os.path.join(hist_dir, os.path.basename(p)))
        for extra in (mcfg.DAILY_CSV_PATH, scfg.DAILY_CSV_PATH, scfg.ALLOC_CSV_PATH,
                      (icfg.DAILY_CSV_PATH if icfg else None), (icfg.ALLOC_CSV_PATH if icfg else None)):
            if extra and os.path.exists(extra):
                shutil.copy2(extra, os.path.join(hist_dir, os.path.basename(extra)))
        print(f"[runner] 이력 보관: {hist_dir}")

    # ---- 실매매 적용 전략 배너 ----
    try:
        nd = M.build_next_day_prediction(res, mcfg)
        lines = ["실매매 적용 전략 = ★ M: SPY 국면전략 (market_regime_report.xlsx)",
                 f"기준일 {nd['기준일'].date()}{nd.get('기준일_경과주의', '')}",
                 f"다음 거래일 {nd['다음거래일'].date()} : 확정 국면 {nd['확정국면']} / 목표비중 {nd['목표비중']:.2f} "
                 f"/ 현재 체결비중 {nd['체결비중']:.2f} / 예상 행동 {nd['예상행동_kr']}",
                 f"발동 규칙 {nd['발동규칙']} | 복합점수 백분위 {nd['복합점수백분위']} | H {nd['위험점수백분위(H)']}",
                 "S★(섹터)·I★(산업) 리포트는 진단·연구용 — 실매매 주문에 반영되지 않음"]
        _banner(lines)
    except Exception as e:  # noqa
        print(f"[runner] 배너 생성 실패(리포트는 정상): {type(e).__name__}: {e}")

    # ---- Colab 다운로드 / Kaggle 보존 안내 ----
    if download and env == "colab":
        if hasattr(M, "maybe_colab_download_many"):
            M.maybe_colab_download_many(paths, zip_name=os.path.join(base, "regime_reports_bundle.zip"))
        else:
            for p in paths:
                M.maybe_colab_download(p)
    elif env == "kaggle":
        print("[runner] Kaggle: 리포트는 /kaggle/working 에 보존됨 — 노트북 'Output' 탭에서 다운로드, "
              "Persistence(Files) 설정이면 다음 세션에도 캐시·리포트가 그대로 남음. 커밋(Save & Run All)하면 버전별 Output으로 저장.")
    print(f"[runner] 총 소요 {time.time() - t_all:.0f}초")
    return {"env": env, "base": base, "paths": paths, "history_dir": hist_dir, "res": res, "sres": sres, "ires": ires,
            "mcfg": mcfg, "scfg": scfg, "icfg": icfg}


if __name__ == "__main__":
    main()
