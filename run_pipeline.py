# =============================================================================
#  run_pipeline.py
#  VERSION: v1.2.0 - 2026-09-13 - [문서 + 단계별 실행 레시피 — 실행 로직 무변경] S v0.42.1 · I v0.6.0 ·
#                    M v1.52.0. REPORT45 §6 구현분: S-I(채택 최근성 진단) · S-J(횡단면 후보 2종) ·
#                    S-K([중립상한격자]) · S-L(13l 실보유 정의) · S-M(게이트 감쇠 t) · S-N(H자기확인, 기본 꺼짐) ·
#                    S-O(06c 하락경과 행) · I-D(산업 리포트 완성: 13g 버그·13개 시트·13j·12 확장·00 정리) ·
#                    S-G/I-B 폐기. 러너는 그대로 S.run/I.run에 cfg를 넘기므로 s_overrides/i_overrides로 전부 제어한다.
#
#  ★ 단계별 실행 레시피(REPORT45 §7 — 한 번에 하나씩 판정하려면 이대로):
#    1) **이번 라운드 기본 실행**: out = RP.main()
#       기본값으로 켜져 있는 것: S-I①(13g 최근창 진단 4열, 판정에는 미반영) · S-J(후보 2종 — 채택은
#         매년 워크포워드가 정한다) · S-K([중립상한격자] 3행) · S-L(13l 실보유 정의 + 블록 F) ·
#         S-M(09c 판정을 감쇠 t로) · S-O(06c 진단 4행) · I-D(리포트 시트 완성).
#       기본값으로 **꺼져 있는** 것: S-I②(ROTATION_RECENT_MIN_T=None) · S-N(H자기확인) · I-C′ · I 동결 해제.
#       이번 실행에서 볼 것(REPORT45 §7):
#         · 13g: SCORE_PCT의 '최근 504일 t'와 '최근 252일 1위 실현상위3 비율'이 2024~26에 0/무작위(≈0.09)
#                 근처인가 — 확장창 t 3.6~3.9와의 간극이 S-I②를 켤 근거다.
#         · 13g: SCORE_CS_Z·SCORE_PCT_CHG_21이 채택되는 해가 있는가(포화하지 않는 척도가 실제로 이기는가).
#         · 13_섹터배분전략 [중립상한격자] 3행의 4기준(①CAGR ②칼마 ③MDD ④강건성).
#         · 09c: 감쇠 t로 XLE 2025·XLU 2022가 '유지'로 돌아오는가.
#         · 09e: **비어 있는 것이 정상**(S-N이 꺼져 있다). 13l: 블록 A/B/E가 실보유 기준으로 바뀐 값.
#         · I: 13g_산업순환매신호채택이 **처음으로** 리포트에 실린다(동결에서도) — 하위1 t를 여기서 본다.
#              02~09d 13개 시트 + 12_산업요약 확장 열도 이번부터.
#    2) S-N 단독 판정(⚠ 신호층 — 1)의 결과를 본 **뒤에** 따로 돌린다):
#         out = RP.main(s_overrides={"SECTOR_HAZARD_CONFIRM_GATE": True})
#         확인: 09e_H자기확인게이트에서 XLE가 '자기확인 요구'로, XLK가 '유지'로 갈리는가 ·
#               12시트 'H진입 익일평균'이 XLE에서 음수로 돌아서는가 · XLV 국면정의검증이 PASS로 가는가 ·
#               01_일별의 'H자기확인(차단)' 열에 실제 차단일이 서는가. ⚠ 되돌리기: 오버라이드를 빼면 끝.
#    3) S-I② 최근성 조건(2)의 결과와 무관하게 1)의 13g를 보고 결정):
#         out = RP.main(s_overrides={"ROTATION_RECENT_MIN_T": 0.0})
#         확인: 2024~26 리더일이 줄고 13l '실보유 리더 > SPY'가 오르는가. ⚠ 되돌리기: None.
#    4) I 동결 해제(13j_배분거래내역은 이때만 채워진다 — 배분이 없으면 실제 거래도 없다):
#         out = RP.main(i_overrides={"INDUSTRY_LAYER_FROZEN": False})   # 배분·격자 재실행 ≈ 38분
#         확인: 13j에서 산업이 **부모 ETF를 이긴 비율**(00시트 요약 1줄) · 14_계층정합 위반 0일.
#    5) I-C′(산업 데이터를 S 후보로 — ⚠ 캐시 무효화, 단독으로):
#         out = RP.main(s_overrides={"USE_INDUSTRY_EXTRA_CANDIDATES": True})
#    (프로젝트 규칙 "한 번에 하나"를 지키려면 1) → 2) → 3) 순서. 2)와 3)을 같이 켜면 13l 변화의
#     원인 귀속이 흐려진다.)
#
#  ⚠ v0.42.1/v0.6.0의 기본값과 되돌리기 한 줄(신호·판정층은 ⚠ 표시):
#    S: s_overrides={"SECTOR_HAZARD_CONFIRM_GATE": True}    # ⚠ S-N H진입 자기확인(기본 False — 사전등록만)
#       s_overrides={"ROTATION_RECENT_MIN_T": 0.0}          # ⚠ S-I② 최근성 강등(기본 None = 진단만)
#       s_overrides={"SECTOR_REGIME_GATE_DECAY": False}     # ⚠ S-M 되돌리기: 게이트 판정을 균등가중 t로(v0.41.0)
#       s_overrides={"ROTATION_NEUTRAL_CAP_GRID": ()}       # S-K 격자 끔(라이브 적용은 ROTATION_NEUTRAL_CAP=0.4 등)
#       s_overrides={"ROTATION_NEUTRAL_LEADER_GRID": (0.25, 0.5, 1.0)}  # 폐기된 [중립리더격자] 되살리기
#       s_overrides={"ROTATION_REVERSE_AVOID_GRID": (2.0, 2.5)}         # 폐기된 S-G 격자 되살리기
#       s_overrides={"SECTOR_OFF_AGE_SENSITIVITY": False}   # S-O 06c 진단 4행 끔(실행시간 미미)
#       s_overrides={"ROTATION_SIGNALS": (...)}             # S-J 후보 2종 제외하려면 목록에서 빼면 된다
#       s_overrides={"SECTOR_MARKET_BLOCK_CAP": 0.5, "SECTOR_MACRO_T_MIN": 2.5}  # ⚠ S-A 되살리기(캐시 무효화)
#       s_overrides={"SECTOR_REGIME_GATE_STAT": "fwd21", "REGIME_ACCEPT_STAT": "fwd21"}  # 판정 잣대 v0.40.0
#       s_overrides={"SECTOR_REGIME_GATE": False}           # 하락 정보 게이트 자체를 끔
#       s_overrides={"USE_INDUSTRY_BREADTH": False}         # 산업폭 후보 끔(산업 가격 수집 생략)
#       s_overrides={"USE_INDUSTRY_EXTRA_CANDIDATES": True} # ⚠ I-C′ 확장 후보 2종 켬(단독으로, 캐시 무효화)
#    I: i_overrides={"INDUSTRY_LAYER_FROZEN": False}        # ⚠ 동결 해제(배분·격자·13j·14·15 재실행 ≈ 38분)
#       i_overrides={"INDUSTRY_RUN_SENSITIVITY": True}      # I-D(D) 06c_임계값민감도 켬(산업당 수 초 × 29)
#       i_overrides={"INDUSTRY_AVOID_STANDALONE_GRID": (0.5, 1.0)}  # 폐기된 I-B 격자 되살리기
#       i_overrides={"ROTATION_VALIDATION_MODE": "pooled"}  # 채택 검증을 v0.3.1 방식으로
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

VERSION = "v1.2.0"
VERSION_DATE = "2026-09-13"

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
