#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
predictor2_result.py  — [코드2] 결과 생성 전용
────────────────────────────────────────────────────────────────────────
공유 코어(predictor_core.py)를 그대로 사용해, 코드1이 만든 '지표풀 엑셀'
(pool_ensemble_<티커>_<날짜>.xlsx)들을 참고하여, 기존과 '동일한' 전체 결과
엑셀(ensemble_search_<티커>_<날짜>.xlsx)을 생성한다.

★★ 동작 방식 — RUN_TICKERS를 보지 않는다.
    1. 풀 폴더(POOL_DIR)에서 발견되는 모든 티커를 매번 다시 스캔한다
       (도중에 코드1이 새 풀을 추가해도 놓치지 않음).
    2. 그중 '오늘 결과가 없거나(또는 장마감 전 낡은 결과인) 티커'를 하나 고른다.
    3. 그 자리를 '원자적으로' 클레임(exclusive 파일 생성)한 뒤 실제 결과를 생성한다.
       세션이 끊겨 재실행해도 이미 끝난 티커는 건너뛰고 이어서 다음 티커부터 진행되고,
       여러 인스턴스를 동시에 돌려도 서로 같은 티커를 중복 작업하지 않는다.
    4. 다시 1번으로 돌아가 반복, 더 이상 대상이 없으면 종료.

★★ (요청) 미국장 마감(한국시간 새벽 5시) 인식 — 그 시각 이후 실행이면, '오늘 날짜'로 된
   결과 파일이 이미 있어도 그게 '오늘 새벽 5시 이전'(=장마감 전, 데이터가 하루 전 것까지만
   반영된 상태)에 만들어진 거라면 낡은 것으로 보고 새로 돌린다. 5시 이후에 만들어진 파일은
   이미 최신이므로 그대로 스킵.

★★ (요청) 여러 인스턴스 동시 실행 대비 — 클레임을 'x'(exclusive create) 모드로 만들어
   원자적으로 처리. 이미 다른 인스턴스가 막 클레임한 티커는 즉시 실패를 감지해 건너뛰고
   바로 '다른' 대상 티커로 넘어간다 (같은 티커 중복 작업 방지).

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
from datetime import datetime, timedelta

import predictor_core as core

try:
    from zoneinfo import ZoneInfo
    _KST = ZoneInfo('Asia/Seoul')
except Exception:
    _KST = None   # zoneinfo 없으면(구버전 파이썬 등) 시스템 로컬시간을 그대로 사용

# ── 설정 ──
POOL_PREFIX = 'pool_ensemble'   # 코드1과 동일해야 함
POOL_DIR    = None              # 지표풀 엑셀 폴더 (None=core.ENSEMBLE_MIRROR_DIR)
OUT_DIR     = None              # 결과 엑셀 저장 폴더 (None=core.OUTPUT_DIR, 로컬)
MARKET_CLOSE_HOUR_KST = 5        # ★ (요청) 미국장 마감 = 한국시간 새벽 5시 (서머타임 무관 근사)


def _now_kst():
    """★ Colab 기본 시스템 시간대는 UTC라, datetime.now()를 그냥 쓰면 '새벽 5시' 판정이
       틀어진다. zoneinfo로 명시적으로 한국시간(KST)을 구한다."""
    if _KST is not None:
        return datetime.now(_KST).replace(tzinfo=None)   # 이후 비교 편의를 위해 naive로
    return datetime.now()   # 폴백(정확도 낮음 — zoneinfo 미지원 환경)


def _today_str_kst():
    return _now_kst().strftime('%Y-%m-%d')


def _latest_market_close_kst():
    """★ (요청) '가장 최근에 지나간' 장마감(한국시간 새벽 MARKET_CLOSE_HOUR_KST시) 시각.
       지금이 그 시각 이전이면 어제의 마감시각, 이후면 오늘의 마감시각을 반환."""
    now = _now_kst()
    today_close = now.replace(hour=MARKET_CLOSE_HOUR_KST, minute=0, second=0, microsecond=0)
    if now < today_close:
        return today_close - timedelta(days=1)
    return today_close


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


def _is_stale(path, cutoff_dt):
    """★ (요청) path의 결과 파일이 '가장 최근 장마감 시각(cutoff_dt)' 이전에 만들어졌으면
       True(낡음 — 장마감 전 데이터라 다시 만들어야 함). 파일이 없으면 True 취급(=필요함)."""
    if not os.path.exists(path):
        return True
    try:
        mtime = datetime.fromtimestamp(os.path.getmtime(path))
        return mtime < cutoff_dt
    except Exception:
        return True   # mtime 확인 실패 시 안전하게 '낡음'으로 간주 → 다시 만듦


def _find_next_pending_ticker(pool_dir, mirror_dir, today_str, cutoff_dt, *, exclude=()):
    """pool_dir에서 발견되는 티커들을 순서대로 보면서, 다음 조건을 만족하는 첫 티커를 반환:
       - '오늘 날짜' ensemble_search 결과 파일이 아예 없거나,
       - 있어도 '가장 최근 장마감(cutoff_dt)' 이전에 만들어진 낡은 것.
       exclude에 있는 티커는 건너뜀(이번 실행에서 이미 실패/다른 인스턴스가 선점한 것 — 재시도 방지).
       (없으면 (None, None)) — 매번 새로 스캔하므로 그 사이 늘어난 풀 파일도 반영됨."""
    for tk in _discover_pool_tickers(pool_dir):
        if tk in exclude:
            continue
        expected = os.path.join(mirror_dir, f"ensemble_search_{tk}_{today_str}.xlsx")
        if _is_stale(expected, cutoff_dt):
            return tk, expected
    return None, None


def main():
    pool_dir = POOL_DIR or getattr(core, 'ENSEMBLE_MIRROR_DIR', None) or getattr(core, 'OUTPUT_DIR', '.')
    out_dir = OUT_DIR or getattr(core, 'OUTPUT_DIR', '.')
    # ★ 클레임/완료 확인은 항상 드라이브(mirror_dir) 기준 — 로컬(out_dir)은 Colab
    #   세션이 끊기면 사라지므로, 재실행 시 '이미 했는지' 판단할 수 있는 유일한
    #   영속 저장소는 드라이브뿐이다.
    mirror_dir = getattr(core, 'ENSEMBLE_MIRROR_DIR', None) or pool_dir
    today_str = _today_str_kst()
    cutoff_dt = _latest_market_close_kst()

    try:
        os.makedirs(mirror_dir, exist_ok=True)
    except Exception as e:
        print(f"  ⚠ 드라이브 폴더 생성 실패: {e}")

    print(f"[코드2] 결과 생성 — 풀 폴더를 계속 재탐색하며 '오늘 결과 없거나 낡은' 티커부터 순차 처리")
    print(f"  풀 폴더: {pool_dir}")
    print(f"  결과 확인(드라이브): {mirror_dir}")
    print(f"  기준 날짜(KST): {today_str}")
    print(f"  ★ 장마감 기준(KST): {cutoff_dt.strftime('%Y-%m-%d %H:%M')} 이전에 만들어진 "
          f"결과는 낡은 것으로 보고 다시 만듭니다 (미국장 마감=한국시간 새벽 "
          f"{MARKET_CLOSE_HOUR_KST}시 기준).")
    print(f"  ℹ 여러 인스턴스를 동시에 돌려도 안전합니다 — 클레임이 원자적이라 서로 다른 "
          f"티커를 자동으로 나눠 맡습니다.")

    made = []
    skip_this_run = set()   # ★ 이번 실행에서 실패했거나 다른 인스턴스가 선점한 티커 — 무한재시도 방지
    while True:
        tk, expected_path = _find_next_pending_ticker(pool_dir, mirror_dir, today_str, cutoff_dt,
                                                       exclude=skip_this_run)
        if tk is None:
            print("\n  ℹ 처리할 티커를 더 찾지 못했습니다 — 종료.")
            break

        # ★★ (요청) 원자적 클레임 — 'x'(exclusive) 모드는 파일이 이미 있으면 즉시 예외를 던진다.
        #   check-then-create가 아니라 '생성 자체'가 원자적 검사라, 다른 인스턴스가 그 사이
        #   먼저 같은 티커를 클레임했더라도 여기서 확실히 걸러진다 → 중복 작업 없이 바로
        #   '다른' 대상으로 넘어간다.
        #   (낡은 파일을 다시 만드는 경우엔 기존 낡은 파일을 먼저 지우고 exclusive 생성한다.)
        try:
            if os.path.exists(expected_path):
                os.remove(expected_path)   # 낡은(장마감 전) 결과 — 새로 클레임하기 전에 제거
            with open(expected_path, 'xb'):
                pass
        except FileExistsError:
            print(f"  ⏭ [{tk}] 다른 인스턴스가 방금 선점함 — 건너뛰고 다른 티커로 진행")
            skip_this_run.add(tk)
            continue
        except Exception as e:
            print(f"  ⚠ [{tk}] 클레임 파일 생성 실패(무시하고 진행): {e}")

        print(f"\n{'='*60}\n[{tk}] 결과 생성\n{'='*60}")

        try:
            out = core.build_result_excel_from_pool(
                tk, pool_dir=pool_dir, out_dir=out_dir, pool_prefix=POOL_PREFIX,
                mirror_to_ensemble=True)   # ★ 항상 True — 클레임 파일을 실제 결과로 덮어써야 함
            if out:
                made.append(out)
            else:
                # 실패(None 반환) — 빈 클레임 파일을 지워서 '다음번 실행 때' 다시 시도되게 함.
                if os.path.exists(expected_path) and os.path.getsize(expected_path) == 0:
                    os.remove(expected_path)
                skip_this_run.add(tk)
        except Exception as e:
            print(f"  ✗ {tk} 실패: {e}")
            import traceback; traceback.print_exc()
            try:
                if os.path.exists(expected_path) and os.path.getsize(expected_path) == 0:
                    os.remove(expected_path)
            except Exception:
                pass
            skip_this_run.add(tk)

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
