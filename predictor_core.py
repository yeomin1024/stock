#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
predictor1_pool.py  — [코드1] 결과 직접 생성 (풀 파일 안 만듦)
────────────────────────────────────────────────────────────────────────
★★★ (요청) 이제 지표풀 엑셀(pool_ensemble_*)을 만들지 않는다. 임시 전체파일을
만들었다가 풀 시트만 뽑고 버리는 대신, 그 전체 결과를 곧바로 최종 결과 엑셀
'ensemble_search_<티커>_<날짜>.xlsx'로 저장한다.

· 코드2(predictor2_result.py)와의 2단계 분리(풀 선출 → 결과 재현) 없이,
  이 스크립트 하나로 티커별 전체 분석 결과가 바로 나온다.
· 자동 다운로드 + 드라이브(ENSEMBLE_MIRROR_DIR) 미러링 모두 그대로 지원.

사용법 (Colab):
    from google.colab import drive; drive.mount('/content/drive')
    import predictor_core as core
    import predictor1_pool as p1
    p1.RUN_TICKERS = ['THC', 'HUM', 'TRV']   # 원하는 티커
    p1.main()
────────────────────────────────────────────────────────────────────────
"""
import predictor_core as core

# ── 설정 ──
RUN_TICKERS = None          # None이면 코어의 기본 티커 목록(core.RUN_TICKERS 또는 MULTI_TICKERS) 사용
OUT_DIR     = None          # None이면 core.OUTPUT_DIR(로컬) — 드라이브 미러는 별도로 자동 수행됨


def _resolve_tickers():
    if RUN_TICKERS:
        return list(RUN_TICKERS)
    for _name in ('RUN_TICKERS', 'MULTI_TICKERS', 'TICKERS'):
        _v = getattr(core, _name, None)
        if _v:
            return list(_v)
    _t = getattr(core, 'TICKER', None)
    return [_t] if _t else []


def main():
    tickers = _resolve_tickers()
    if not tickers:
        print("✗ 대상 티커가 없습니다. p1.RUN_TICKERS = ['THC', ...] 로 지정하세요.")
        return []
    out_dir = OUT_DIR or getattr(core, 'OUTPUT_DIR', '.')
    print(f"[코드1] 결과 직접 생성(풀 파일 없이 바로 ensemble_search) — {len(tickers)}개 티커 → {out_dir}")
    print(f"  대상: {tickers}")
    saved = []
    for tk in tickers:
        print(f"\n{'='*60}\n[{tk}] 결과 생성\n{'='*60}")
        try:
            # ★★★ (요청) build_pool_excel_for_ticker(풀만 뽑고 버림) 대신
            #   build_ensemble_search_direct를 써서 곧바로 최종 결과 엑셀을 만든다.
            p = core.build_ensemble_search_direct(tk, out_dir=out_dir)
            if p:
                saved.append(p)
        except Exception as e:
            print(f"  ✗ {tk} 실패: {e}")
            import traceback; traceback.print_exc()
    print(f"\n✅ [코드1] 완료 — 결과 엑셀 {len(saved)}개 생성")
    for p in saved:
        print(f"   · {p}")
    # ★ 자동 다운로드 — build_ensemble_search_direct 내부에서 run_ensemble_search가
    #   호출될 때 이미 자동다운로드가 걸리지만(caller_name 제외 목록에 없음), 혹시
    #   빠뜨린 경우를 대비해 배치로 한 번 더 시도(이미 받은 파일은 브라우저가 자체적으로
    #   무시하거나 재확인만 하므로 중복 문제 없음).
    if getattr(core, 'AUTO_DOWNLOAD_EXCEL', False) and saved and hasattr(core, '_auto_download_excels'):
        try:
            core._auto_download_excels(saved)
        except Exception as e:
            print(f"  ⚠ 자동 다운로드 실패(무시): {e}")
    return saved


if __name__ == '__main__':
    main()
