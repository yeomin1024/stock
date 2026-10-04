# ============================================================================================
#  미국주식 가상계좌 자동매매 — Kaggle 한 셀 실행기 (실제 주문 없음: 가상계좌 $10,000, 주문 잠금)
#  준비: Settings → Internet 켜기 / Add-ons → Secrets 연결
#        KIWOOM_APPKEY, KIWOOM_SECRETKEY (키움 분봉 수집·키움 시세 쓸 때)
#        GITHUB_TOKEN (결과를 GitHub에 올릴 때 — 이 저장소 Contents 읽기·쓰기 권한)
# ============================================================================================
# ===== 설정 — wget으로 받아 실행하는 셀에서 같은 이름으로 먼저 정하면 그 값이 우선합니다 =====
for _k, _v in dict(
    REPO="https://github.com/yeomin1024/stock", BRANCH="main",
    STRATEGY="P4",            # 'P4'(='R1') MDD −10%·승률↑ / 'P3' / 'P2' MDD −5% / 'R1_5000' 1종목 모멘텀 / 'N2' / 'A' / 'C0'
    RUN_COLLECT=False,        # 키움 5분봉 기록 모으기(한 번) → GitHub data/kiwoom_minute/
    KIWOOM_MOCK=True,         # 연결한 키움 키가 모의투자 키면 True, 실전 키면 False(조회만 함)
    RUN_SIM=True,             # 과거 실시간 시뮬레이션 → GitHub results/kaggle/sim/<날짜>_<전략>/
    RUN_PAPER=False,          # 실시간 가상거래(미국 장중, 밤에 Save & Run All) → GitHub results/kaggle/paper/<전략>/
    QUOTE_SOURCE="yfinance",  # 가상거래 시세: 'yfinance'(앱키·IP 불필요) / 'kiwoom'(틱 실시간, IP 등록 필요)
    PAPER_CASH=10000,
    PUSH_RESULTS=True,        # 결과 파일을 GitHub에 올림(GITHUB_TOKEN 필요, 없으면 건너뜀)
    RESULT_DIR="results/kaggle",
).items():
    globals().setdefault(_k, _v)
# ===========================
import subprocess, sys
subprocess.run([sys.executable, "-m", "pip", "-q", "install", "websockets", "yfinance"], check=False)
import os, glob, json, shutil, time, asyncio, datetime as dt

SRC, PUSH_DIR, OUT = "/tmp/kiwoom_src", "/tmp/kiwoom_push", "/kaggle/working"


def _secret(name):
    try:
        from kaggle_secrets import UserSecretsClient
        return UserSecretsClient().get_secret(name) or ""
    except Exception:
        return os.environ.get(name, "")


TOKEN = _secret("GITHUB_TOKEN")


def _git(args, cwd=None):
    url = REPO.rstrip("/").removesuffix(".git") + ".git"
    args = [a.replace("@@REPO@@", url.replace("https://", f"https://x-access-token:{TOKEN}@", 1) if TOKEN else url)
            for a in args]
    r = subprocess.run(["git"] + args, cwd=cwd, capture_output=True, text=True)
    msg = (r.stdout or "") + (r.stderr or "")
    return r.returncode, (msg.replace(TOKEN, "***") if TOKEN else msg)


def _clone(dst):
    shutil.rmtree(dst, ignore_errors=True)
    return _git(["clone", "--depth", "1", "-b", BRANCH, "@@REPO@@", dst])


def push_results(files: dict, message: str):
    """files = {저장소 안 경로: Kaggle 파일/폴더}. 새로 받아 복사 → commit → push (충돌 나면 다시 받아 3번까지)."""
    if not PUSH_RESULTS:
        return
    if not TOKEN:
        print("⚠️ GITHUB_TOKEN이 없어 결과를 GitHub에 올리지 않습니다(Kaggle Output 탭에는 남아 있음)")
        return
    for attempt in range(3):
        c, m = _clone(PUSH_DIR)
        if c:
            print("❌ 올리기용 clone 실패:", m[-300:]); return
        _git(["config", "user.name", "kaggle-runner"], PUSH_DIR)
        _git(["config", "user.email", "kaggle-runner@users.noreply.github.com"], PUSH_DIR)
        n = 0
        for rel, src in files.items():
            dst = os.path.join(PUSH_DIR, rel)
            if os.path.isdir(src):
                shutil.copytree(src, dst, dirs_exist_ok=True); n += 1
            elif os.path.exists(src):
                os.makedirs(os.path.dirname(dst), exist_ok=True); shutil.copy2(src, dst); n += 1
        if n == 0:
            print("올릴 파일 없음"); return
        _git(["add", "-A"], PUSH_DIR)
        c, m = _git(["commit", "-m", message], PUSH_DIR)
        if c and "nothing to commit" in m:
            print("바뀐 내용 없음 — 올리지 않음"); return
        c, m = _git(["push", "origin", BRANCH], PUSH_DIR)
        if c == 0:
            print(f"✅ GitHub에 올림: {message}", flush=True); return
        print("push 실패, 다시 시도:", m[-200:], flush=True); time.sleep(5)
    print("❌ GitHub push 3번 실패 — Kaggle Output 탭의 파일은 남아 있습니다")


# 1) 코드·신호·이전 가상계좌 받기
c, m = _clone(SRC)
if c:
    raise RuntimeError("git clone 실패 — REPO 주소·Internet 설정을 확인하세요\n" + m)
py = glob.glob(f"{SRC}/**/kiwoom_autotrader.py", recursive=True)
if not py:
    raise FileNotFoundError(f"저장소에 kiwoom_autotrader.py가 없습니다 → {REPO}/upload/{BRANCH}")
sys.modules.pop("kiwoom_autotrader", None)
sys.path.insert(0, os.path.dirname(py[0]))
from kiwoom_autotrader import *
assert not ALLOW_ORDERS, "주문 잠금이 풀려 있습니다"
if STRATEGY in PRESET_ALIAS:                    # 실행 셀을 바꾸지 않아도 최신 추천 전략으로
    print(f"STRATEGY '{STRATEGY}' → 최신 추천 '{PRESET_ALIAS[STRATEGY]}' 로 실행")
    STRATEGY = PRESET_ALIAS[STRATEGY]


SIG = os.path.join(OUT, "signals"); os.makedirs(SIG, exist_ok=True)
REP = latest_report_dir(SRC)          # 저장소 results/reports/<날짜>/ = 국면·섹터·산업·주식층 최신 버전 리포트
built = signals_from_reports(REP, SIG) if REP else {}
for n in REPORT_SHEETS:               # 우선순위: Kaggle 입력 데이터셋 > 최신 리포트 > 저장소 signals/
    hits = sorted(glob.glob(f"/kaggle/input/**/{n}", recursive=True), key=os.path.getmtime, reverse=True)
    if hits:
        shutil.copy(hits[0], os.path.join(SIG, n)); src = hits[0]
    elif n in built:
        src = f"{os.path.relpath(REP, SRC)}/{built[n]}"
    elif os.path.exists(f"{SRC}/signals/{n}"):
        shutil.copy(f"{SRC}/signals/{n}", os.path.join(SIG, n)); src = "저장소 signals/ (이전 버전)"
    else:
        src = "없음"
    print(f"신호 {n:30s} ← {src}")
PAPER_DIR = f"{RESULT_DIR}/paper/{STRATEGY}"                             # 전략마다 따로 쓰는 가상계좌
for n in ("paper_account.json", "paper_trades.csv", "paper_equity.csv"):    # 가상계좌 이어 쓰기(GitHub에 올라간 것)
    p = os.path.join(SRC, PAPER_DIR, n)
    if os.path.exists(p) and not os.path.exists(os.path.join(OUT, n)):
        shutil.copy(p, os.path.join(OUT, n)); print("이어 쓰기:", n)

cfg = Config(**PRESETS[STRATEGY], paper_cash=PAPER_CASH, quote_source=QUOTE_SOURCE, signals_dir=SIG, mock=KIWOOM_MOCK,
             paper_state=f"{OUT}/paper_account.json", trade_log=f"{OUT}/paper_trades.csv", equity_log=f"{OUT}/paper_equity.csv")
assert cfg.leverage == 1.0 and not (set(cfg.symbols) & LEVERAGED_ETFS), "레버리지(신용·레버리지 ETF) 사용 금지"
TODAY = now_et().strftime("%Y-%m-%d")
print(f"전략 {STRATEGY}: {PRESET_NOTES[STRATEGY]}")
print(f"오늘 미국 정규장 {et_session_in_kst()} | 대상 {len(cfg.symbols)}종 | 수수료 {cfg.fee_pct}% | 레버리지 없음 | "
      f"국면 {cfg.regime} | 섹터필터 {cfg.sector_filter}")


def kiwoom_api_with_ip_wait(minutes_wait=30):
    import requests
    ip = requests.get("https://api.ipify.org", timeout=10).text
    print(f"이 세션의 공인 IP: {ip} → openapi.kiwoom.com {'모의투자 ' if KIWOOM_MOCK else ''}App Key 관리의 허용 IP에 등록", flush=True)
    api = KiwoomUSAPI(_secret("KIWOOM_APPKEY"), _secret("KIWOOM_SECRETKEY"), KIWOOM_MOCK)
    for _ in range(minutes_wait * 2):
        try:
            api.get_token(); print("✅ 키움 토큰 OK", flush=True); return api
        except Exception as e:
            print("⏳ 토큰 실패 — IP 등록을 기다립니다:", str(e).splitlines()[0], flush=True); time.sleep(30)
    raise RuntimeError("토큰을 받지 못했습니다 — IP 등록·모의/실전 키(KIWOOM_MOCK)를 확인하세요")


# 2) 키움 분봉 기록 모으기
if RUN_COLLECT:
    api = kiwoom_api_with_ip_wait()
    syms = dict(cfg.symbols); syms.update({"QQQ": "ND", "SPY": "NA"})
    cov = collect_kiwoom_minutes(api, syms, 5, f"{OUT}/kiwoom_minute")
    spy = cov[cov["종목"] == "SPY"].to_dict("records")
    if not spy or not (spy[0].get("봉") or 0) > 0:                        # SPY 거래소 코드가 'NA'가 아니면 'NY'로
        print("SPY 재시도(NY):", collect_kiwoom_minutes(api, {"SPY": "NY"}, 5, f"{OUT}/kiwoom_minute_spy", compare_yf=0).to_dict("records"))
        if os.path.exists(f"{OUT}/kiwoom_minute_spy/SPY_5m.csv.gz"):
            shutil.copy(f"{OUT}/kiwoom_minute_spy/SPY_5m.csv.gz", f"{OUT}/kiwoom_minute/SPY_5m.csv.gz")
    display(cov)
    print("받은 거래일 중앙값:", cov.get("거래일", pd.Series(dtype=float)).median())
    push_results({"data/kiwoom_minute": f"{OUT}/kiwoom_minute"}, f"키움 5분봉 기록 {TODAY}")

# 3) 과거 실시간 시뮬레이션
if RUN_SIM:
    import yfinance as yf, matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    want = list(cfg.symbols) + ([cfg.mkt_symbol] if cfg.mkt_stop_pct or cfg.mkt_flat_pct else [])   # + 시장 급락 판단용 SPY
    bars = download_hourly_yf(want)               # 상장 730일 미만 종목(BTSG 등)은 날짜 범위로 다시 받음
    print(len(bars), "/", len(cfg.symbols), "종목 시간봉")
    sig = DailySignals(cfg, strict=False)
    for w in sig.warn:
        print("⚠️", w)
    res = simulate(cfg, bars, sig)
    spy_c = yf.Ticker("SPY").history(period="5y", interval="1d", auto_adjust=True)["Close"]
    spy_c.index = spy_c.index.tz_localize(None).normalize()
    met = sim_metrics(res, spy_close=spy_c)
    print(json.dumps(met, ensure_ascii=False, indent=1))
    sim_dir = f"{OUT}/sim"
    save_sim(res, sim_dir, "sim", met, cfg)
    eq = res["equity"].assign(날짜=lambda d: pd.to_datetime(d["날짜"])).set_index("날짜")["자산($)"]
    s2 = spy_c.reindex(eq.index).ffill(); s2 = s2 / s2.iloc[0] * cfg.paper_cash
    ax = eq.plot(label=f"paper ({STRATEGY})", figsize=(11, 4), logy=True); s2.plot(ax=ax, label="SPY buy&hold")
    ax.set_xlabel(""); ax.legend(); ax.grid(alpha=.3); ax.set_title(f"{STRATEGY}: {met.get('수익률(%)')}%, MDD {met.get('MDD(%)')}% (log scale)")
    plt.savefig(f"{sim_dir}/sim_equity.png", dpi=110); plt.close("all")
    push_results({f"{RESULT_DIR}/sim/{TODAY}_{STRATEGY}": sim_dir}, f"과거 시뮬레이션 {TODAY} {STRATEGY}")

# 4) 실시간 가상거래 (밤새: Save & Run All)
if RUN_PAPER:
    nw = now_et()
    close = nw.replace(hour=16, minute=0, second=0, microsecond=0)
    hrs = None if (nw >= close or nw.weekday() >= 5) else (close - nw).total_seconds() / 3600
    if hrs is None:
        print(f"오늘(ET) 미국 장이 끝났거나 주말입니다. 다음 정규장({et_session_in_kst()}) 전에 다시 실행하세요.")
    elif hrs > 11:
        raise RuntimeError(f"장 마감까지 {hrs:.1f}시간 — Kaggle 세션(최대 12시간)이 장중에 끊길 수 있어 시작하지 않습니다.")
    else:
        api = kiwoom_api_with_ip_wait() if QUOTE_SOURCE == "kiwoom" else None
        runner = LiveRunner(cfg, api)
        paper_files = {f"{PAPER_DIR}/{n}": f"{OUT}/{n}" for n in ("paper_account.json", "paper_trades.csv", "paper_equity.csv")}

        async def _hourly_push():                    # 세션이 끊겨도 잃지 않게 1시간마다 중간 저장
            while True:
                await asyncio.sleep(3600)
                runner.book.save_state(cfg.paper_state)
                await asyncio.to_thread(push_results, paper_files, f"가상거래 중간 저장 {now_et():%Y-%m-%d %H:%M} ET")

        _t = asyncio.create_task(_hourly_push())
        try:
            await runner.run()
        finally:
            _t.cancel()
            push_results(paper_files, f"가상거래 {TODAY} {STRATEGY} 마감 — 자산 {usd(runner.book.equity())}")

# 5) 기록 보기
for f in (cfg.trade_log, cfg.equity_log):
    if os.path.exists(f):
        display(pd.read_csv(f, encoding="utf-8-sig").tail(20))
