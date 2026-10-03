# 미국주식 가상계좌 자동매매 — GitHub에서 실행기를 wget으로 받아 실행 (실제 주문 없음 · 레버리지 없음)
# 준비: Settings → Internet 켜기 / Add-ons → Secrets: GITHUB_TOKEN(결과 올리기), KIWOOM_APPKEY·KIWOOM_SECRETKEY(키움 쓸 때)
STRATEGY    = "N2"      # 'N2' 1000% 목표(모멘텀 로테이션+저점매수) / 'A' 저점매수 개선 / 'C0' 이전 기본값
RUN_SIM     = True      # 과거 실시간 시뮬레이션 → GitHub results/kaggle/sim/<날짜>_<전략>/
RUN_PAPER   = False     # 실시간 가상거래(한국 18~22시에 Save & Run All) → GitHub results/kaggle/paper/<전략>/
RUN_COLLECT = False     # 키움 5분봉 기록 모으기(한 번) → GitHub data/kiwoom_minute/
KIWOOM_MOCK = True      # 키움 키가 모의투자 키면 True, 실전 키면 False
!wget -q --no-cache -O /tmp/kaggle_runner.py "https://raw.githubusercontent.com/yeomin1024/stock/main/kaggle_runner_onecell.py?$(date +%s)"
import ast, asyncio
_run = eval(compile(open("/tmp/kaggle_runner.py", encoding="utf-8").read(), "kaggle_runner.py", "exec",
                    flags=ast.PyCF_ALLOW_TOP_LEVEL_AWAIT), globals())
if asyncio.iscoroutine(_run):
    await _run
