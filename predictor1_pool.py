#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
predictor1_pool.py  — [코드1] 지표풀 선출 전용
────────────────────────────────────────────────────────────────────────
공유 코어(predictor_core.py)를 그대로 사용해, 각 티커의 '지표풀'만 선출하여
ensemble 경로에 '지표풀 엑셀'(pool_ensemble_<티커>_<날짜>.xlsx)로 저장한다.

· 기존 지표풀 엑셀이 있으면 → 재탐색 없이 그 풀을 그대로 쓰고 '현재까지 데이터'만 반영해 갱신.
· 없으면 → 전체 탐색 후 풀 저장.
· 파일명이 'ensemble_search_*'(코드2 결과물)와 구분되도록 'pool_ensemble_*' 접두사 사용.

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
POOL_PREFIX = 'pool_ensemble'   # 지표풀 엑셀 접두사 (결과물과 구분)
OUT_DIR     = None          # None이면 core.ENSEMBLE_MIRROR_DIR(내 드라이브 ensemble 경로)


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
    out_dir = OUT_DIR or getattr(core, 'ENSEMBLE_MIRROR_DIR', None) or getattr(core, 'OUTPUT_DIR', '.')
    print(f"[코드1] 지표풀 선출 — {len(tickers)}개 티커 → {out_dir}")
    print(f"  대상: {tickers}")
    saved = []
    for tk in tickers:
        print(f"\n{'='*60}\n[{tk}] 지표풀 선출\n{'='*60}")
        try:
            p = core.build_pool_excel_for_ticker(tk, out_dir=out_dir, pool_prefix=POOL_PREFIX)
            if p:
                saved.append(p)
        except Exception as e:
            print(f"  ✗ {tk} 실패: {e}")
            import traceback; traceback.print_exc()
    print(f"\n✅ [코드1] 완료 — 지표풀 엑셀 {len(saved)}개 생성")
    for p in saved:
        print(f"   · {p}")
    # ★ (버그 수정) 자동 다운로드 — 이전엔 이 호출이 아예 없어서 코드1 결과가 다운로드되지
    #   않았음. run_ensemble_search 내부 자동다운로드도 build_pool_excel_for_ticker 호출
    #   경로에서는 꺼져있으므로(임시 전체파일을 잘못 받으려던 문제 방지), 여기서 최종
    #   풀 엑셀만 한번에 배치로 다운로드한다 (코드2와 동일한 패턴).
    if getattr(core, 'AUTO_DOWNLOAD_EXCEL', False) and saved and hasattr(core, '_auto_download_excels'):
        try:
            core._auto_download_excels(saved)
        except Exception as e:
            print(f"  ⚠ 자동 다운로드 실패(무시): {e}")
    return saved


if __name__ == '__main__':
    main()
