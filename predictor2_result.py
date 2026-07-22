#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
predictor2_result.py  — [코드2] 결과 생성 전용
────────────────────────────────────────────────────────────────────────
공유 코어(predictor_core.py)를 그대로 사용해, 코드1이 만든 '지표풀 엑셀'
(pool_ensemble_<티커>_<날짜>.xlsx)들을 참고하여, 기존과 '동일한' 전체 결과
엑셀(ensemble_search_<티커>_<날짜>.xlsx)을 생성한다.

★★ (요청) 동작 방식 변경 — RUN_TICKERS를 더 이상 보지 않는다.
    1. 풀 폴더(POOL_DIR)에서 발견되는 모든 티커를 매번 다시 스캔한다
       (도중에 코드1이 새 풀을 추가해도 놓치지 않음).
    2. 그중 '오늘 날짜' ensemble_search 결과 파일이 아직 없는 티커를 하나 고른다.
    3. 그 자리에 '빈 파일'을 먼저 만들어 클레임(작업 중/완료 표시)해 둔 뒤,
       실제 결과를 생성한다 — 도중에 세션이 끊겨 재실행해도 이미 만든(또는
       만들고 있는) 티커는 건너뛰고 이어서 다음 티커부터 진행된다.
    4. 다시 1번으로 돌아가 '오늘 결과 파일이 없는' 다음 티커를 찾아 반복한다.
    5. 더 이상 대상이 없으면 종료.

· 재탐색 없이 저장된 풀을 그대로 사용(재현 모드) → 일별 백테스트·KL 순신호·
  카운트0 별도풀·검증_예측로직 등 기존 시트를 동일하게 생성.
· 지표풀 엑셀이 하나도 없으면 아무 것도 하지 않음(먼저 코드1을 실행해야 함).

사용법 (Colab):
    from google.colab import drive; drive.mount('/content/drive')
    import predictor_core as core
    import predictor2_result as p2
    p2.main()
────────────────────────────────────────────────────────────────────────
"""
import os
from datetime import datetime

import predictor_core as core

# ── 설정 ──
POOL_PREFIX = 'pool_ensemble'   # 코드1과 동일해야 함
POOL_DIR    = None              # 지표풀 엑셀 폴더 (None=core.ENSEMBLE_MIRROR_DIR)
OUT_DIR     = None              # 결과 엑셀 저장 폴더 (None=core.OUTPUT_DIR, 로컬)


def _discover_pool_tickers(pool_dir):
    """pool_dir에서 pool_ensemble_<티커>_<날짜>.xlsx 파일들의 티커 집합을 찾음."""
    import glob, re
    _rx = re.compile(rf'^{re.escape(POOL_PREFIX)}_(.+)_(\d{{4}}-\d{{2}}-\d{{2}})\.xlsx$')
    found = {}
    for p in glob.glob(os.path.join(pool_dir, f"{POOL_PREFIX}_*.xlsx")):
        m = _rx.match(os.path.basename(p))
        if m:
            tk, dt = m.group(1), m.group(2)
            if tk not in found or dt > found[tk]:
                found[tk] = dt
    return sorted(found.keys())


def _find_next_pending_ticker(pool_dir, mirror_dir, today_str, exclude=()):
    """pool_dir에서 발견되는 티커들을 순서대로 보면서, mirror_dir(드라이브)에
       '오늘 날짜' ensemble_search 결과 파일이 아직 없는 첫 티커를 반환.
       exclude에 있는 티커는 건너뜀(이번 실행에서 이미 실패 처리한 것 — 무한재시도 방지).
       (없으면 (None, None)) — 매번 새로 스캔하므로 그 사이 늘어난 풀 파일도 반영됨."""
    for tk in _discover_pool_tickers(pool_dir):
        if tk in exclude:
            continue
        expected = os.path.join(mirror_dir, f"ensemble_search_{tk}_{today_str}.xlsx")
        if not os.path.exists(expected):
            return tk, expected
    return None, None


def main():
    pool_dir = POOL_DIR or getattr(core, 'ENSEMBLE_MIRROR_DIR', None) or getattr(core, 'OUTPUT_DIR', '.')
    out_dir = OUT_DIR or getattr(core, 'OUTPUT_DIR', '.')
    # ★ 클레임/완료 확인은 항상 드라이브(mirror_dir) 기준 — 로컬(out_dir)은 Colab
    #   세션이 끊기면 사라지므로, 재실행 시 '이미 했는지' 판단할 수 있는 유일한
    #   영속 저장소는 드라이브뿐이다.
    mirror_dir = getattr(core, 'ENSEMBLE_MIRROR_DIR', None) or pool_dir
    today_str = datetime.now().strftime('%Y-%m-%d')

    try:
        os.makedirs(mirror_dir, exist_ok=True)
    except Exception as e:
        print(f"  ⚠ 드라이브 폴더 생성 실패: {e}")

    print(f"[코드2] 결과 생성 — 풀 폴더를 계속 재탐색하며 '오늘 결과 없는' 티커부터 순차 처리")
    print(f"  풀 폴더: {pool_dir}")
    print(f"  결과 확인(드라이브): {mirror_dir}")
    print(f"  기준 날짜: {today_str}")
    print(f"  ℹ 참고: 계산 도중 세션이 끊기면 0바이트짜리 빈 클레임 파일이 드라이브에 남을 수 "
          f"있습니다. 이 경우 자동 재시도가 안 되니, 다시 하고 싶은 티커의 빈 파일만 드라이브에서 "
          f"직접 지운 뒤 재실행하세요.")

    made = []
    failed_this_run = set()   # ★ 이번 실행에서 이미 실패한 티커 — 무한재시도 방지용(메모리에만 존재)
    while True:
        tk, expected_path = _find_next_pending_ticker(pool_dir, mirror_dir, today_str,
                                                       exclude=failed_this_run)
        if tk is None:
            print("\n  ℹ 오늘 날짜 결과가 없는 티커를 더 찾지 못했습니다 — 종료.")
            break

        print(f"\n{'='*60}\n[{tk}] 결과 생성 (오늘자 결과 없음 → 진행)\n{'='*60}")

        # ★ (요청) 빈 파일을 먼저 만들어 '작업 중'으로 표시 — 세션이 끊겨 재실행해도
        #   이미 시작(또는 완료)한 티커는 건너뛰고 다음 티커로 넘어가게 한다.
        #   (주의: 계산 도중 실패하면 빈 파일이 남을 수 있음 → 아래 except에서 정리)
        try:
            with open(expected_path, 'wb'):
                pass
        except Exception as e:
            print(f"  ⚠ 클레임 파일 생성 실패(무시하고 진행): {e}")

        try:
            out = core.build_result_excel_from_pool(
                tk, pool_dir=pool_dir, out_dir=out_dir, pool_prefix=POOL_PREFIX,
                mirror_to_ensemble=True)   # ★ 항상 True — 클레임 파일을 실제 결과로 덮어써야 함
            if out:
                made.append(out)
            else:
                # 실패(None 반환) — 빈 클레임 파일을 지워서 '다음번 실행 때' 다시 시도되게 함.
                # (이번 실행에서는 failed_this_run에 넣어 무한 재시도를 막는다.)
                if os.path.exists(expected_path) and os.path.getsize(expected_path) == 0:
                    os.remove(expected_path)
                failed_this_run.add(tk)
        except Exception as e:
            print(f"  ✗ {tk} 실패: {e}")
            import traceback; traceback.print_exc()
            try:
                if os.path.exists(expected_path) and os.path.getsize(expected_path) == 0:
                    os.remove(expected_path)
            except Exception:
                pass
            failed_this_run.add(tk)

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
