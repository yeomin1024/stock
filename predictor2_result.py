#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
predictor2_result.py  — [코드2] 결과 생성 전용
────────────────────────────────────────────────────────────────────────
공유 코어(predictor_core.py)를 그대로 사용해, 코드1이 만든 '가장 최근 날짜의
지표풀 엑셀'(pool_ensemble_<티커>_<날짜>.xlsx)을 참고하여, 기존과 '동일한'
전체 결과 엑셀(ensemble_search_<티커>_<날짜>.xlsx)을 생성한다.

· 재탐색 없이 저장된 풀을 그대로 사용(재현 모드) → 일별 백테스트·KL 순신호·
  카운트0 별도풀·검증_예측로직 등 기존 시트를 동일하게 생성.
· 지표풀 엑셀이 없으면 해당 티커는 건너뜀(먼저 코드1을 실행해야 함).

사용법 (Colab):
    from google.colab import drive; drive.mount('/content/drive')
    import predictor_core as core
    import predictor2_result as p2
    p2.RUN_TICKERS = ['THC', 'HUM', 'TRV']
    p2.main()
────────────────────────────────────────────────────────────────────────
"""
import predictor_core as core

# ── 설정 ──
RUN_TICKERS = None              # None이면 풀 엑셀이 존재하는 모든 티커 자동 탐색
POOL_PREFIX = 'pool_ensemble'   # 코드1과 동일해야 함
POOL_DIR    = None              # 지표풀 엑셀 폴더 (None=core.ENSEMBLE_MIRROR_DIR)
OUT_DIR     = None              # 결과 엑셀 저장 폴더 (None=core.OUTPUT_DIR, 로컬)
MIRROR_RESULT_TO_ENSEMBLE = True  # 결과 엑셀도 ensemble 경로에 복사


def _discover_pool_tickers(pool_dir):
    """pool_dir에서 pool_ensemble_<티커>_<날짜>.xlsx 파일들의 티커 집합을 찾음."""
    import glob, os, re
    _rx = re.compile(rf'^{re.escape(POOL_PREFIX)}_(.+)_(\d{{4}}-\d{{2}}-\d{{2}})\.xlsx$')
    found = {}
    for p in glob.glob(os.path.join(pool_dir, f"{POOL_PREFIX}_*.xlsx")):
        m = _rx.match(os.path.basename(p))
        if m:
            tk, dt = m.group(1), m.group(2)
            if tk not in found or dt > found[tk]:
                found[tk] = dt
    return sorted(found.keys())


def main():
    pool_dir = POOL_DIR or getattr(core, 'ENSEMBLE_MIRROR_DIR', None) or getattr(core, 'OUTPUT_DIR', '.')
    out_dir = OUT_DIR or getattr(core, 'OUTPUT_DIR', '.')

    tickers = list(RUN_TICKERS) if RUN_TICKERS else _discover_pool_tickers(pool_dir)
    if not tickers:
        print(f"✗ 지표풀 엑셀({POOL_PREFIX}_*.xlsx)을 {pool_dir} 에서 찾지 못했습니다. 코드1을 먼저 실행하세요.")
        return []
    if not RUN_TICKERS:
        print(f"  ℹ RUN_TICKERS 미지정 → {pool_dir} 안의 모든 지표풀 파일을 자동 탐색했습니다 "
              f"(과거에 코드1로 만들어둔 티커까지 전부 포함될 수 있음). "
              f"특정 티커만 하려면 p2.RUN_TICKERS = ['TICKER1', ...] 로 지정하세요.")
    print(f"[코드2] 결과 생성 — {len(tickers)}개 티커")
    print(f"  풀 폴더: {pool_dir}")
    print(f"  대상: {tickers}")

    made = []
    for tk in tickers:
        print(f"\n{'='*60}\n[{tk}] 결과 생성\n{'='*60}")
        try:
            # ★ (수정) 드라이브 미러링은 이제 core.build_result_excel_from_pool 내부에서
            #   항상 수행됨 (mirror_to_ensemble=True 기본값) — 여기서 또 호출하면 같은
            #   파일이 두 번 복사되므로 제거. MIRROR_RESULT_TO_ENSEMBLE=False로 두면 끌 수 있음.
            out = core.build_result_excel_from_pool(
                tk, pool_dir=pool_dir, out_dir=out_dir, pool_prefix=POOL_PREFIX,
                mirror_to_ensemble=MIRROR_RESULT_TO_ENSEMBLE)
            if out:
                made.append(out)
        except Exception as e:
            print(f"  ✗ {tk} 실패: {e}")
            import traceback; traceback.print_exc()

    print(f"\n✅ [코드2] 완료 — 결과 엑셀 {len(made)}개 생성")
    for p in made:
        print(f"   · {p}")
    # 자동 다운로드 (코어 설정 따름)
    if getattr(core, 'AUTO_DOWNLOAD_EXCEL', False) and made and hasattr(core, '_auto_download_excels'):
        try:
            core._auto_download_excels(made)
        except Exception:
            pass
    return made


if __name__ == '__main__':
    main()
