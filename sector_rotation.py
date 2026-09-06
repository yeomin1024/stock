# test_sector_rotation_v090.py — [v0.9.0] 검증 항목
#   1) 증거 등급(ROTATION_EVIDENCE_TIER): 엄격 신호 + 약한(외부검증 지원) 신호가 같이 채택된 해 → 엄격 신호만으로 리더 판단(단독,
#      교차확인 불필요), 13g '리더 판단 사용' 열, 판단 등급 '엄격'. 끄면(v0.8) 두 신호 일치가 필요해 리더일이 줄어든다.
#      엄격 신호가 없는 해(약한 신호만)는 v0.8 교차확인 그대로(≥2 일치).
#   2) 회피 자격(ROTATION_AVOID_VALIDATE): (a) 하위1 스프레드 통계 — 상/하 모두 정보가 있는 신호는 t ≤ −2(자격 Y), 상위만 정보가 있고
#      하위는 잡음인 신호는 자격 N(채택은 됨). (b) 배관 — 회피가 나오는 픽스처에서 avoid_by_year를 비우면 회피일 0(리더 판단은 불변).
#   3) 13j 배분거래내역: 손계산 픽스처(3자산×12일)에서 구간 분할·진입/청산일·시가 체결가·자산수익(oc·co 연쇄)·전략기여(비용 포함)·
#      진입 판단/청산 사유가 정확한지 + 전략기여 합 == 백테스트 총수익 − 현금레그(비트 수준).
import sys, time
import numpy as np, pandas as pd
sys.path.insert(0, "/root/work/mrt")
import market_regime_trader as M, sector_rotation as S

t_all = time.time()
res = pd.read_pickle("/tmp/res_selftest.pkl")
cal = res["cal"]
eval_idx = cal[cal >= pd.Timestamp(res["cfg"].SIGNAL_START)]
sectors = ("XLK", "XLV", "XLY", "XLP", "XLF", "XLE", "XLI", "XLB", "XLU", "XLC", "XLRE")
full_idx = cal[cal >= pd.Timestamp("2000-01-03")]
nf = len(full_idx)
rng = np.random.default_rng(90)
drift = np.zeros((nf, 11)); z = rng.normal(0, 1, 11) * 0.0006
for i in range(nf):
    z = 0.995 * z + rng.normal(0, 1, 11) * 0.00006
    drift[i] = z
ret_mat = pd.DataFrame(rng.normal(0.0003, 0.012, (nf, 11)) + drift, index=full_idx, columns=sectors)
score_like = pd.DataFrame(np.clip(0.5 + drift * 150 + rng.normal(0, 0.05, (nf, 11)), 0, 1), index=full_idx, columns=sectors)   # 상·하 모두 정보
rng_w = np.random.default_rng(91)
weak = pd.DataFrame(drift * 100 * 0.03 + rng_w.normal(0, 0.06, (nf, 11)), index=full_idx, columns=sectors)                     # 약한 지속형(상위1 t 0.9~1.6)
# 상위만 정보: 그날 drift 최대 섹터만 +1(잡음 위), 나머지는 순수 잡음 → 1위 판별력은 강하고 꼴찌 판별력은 없음
_to = rng.normal(0, 1, (nf, 11))
_to[np.arange(nf), np.argmax(drift, axis=1)] += 6.0
top_only = pd.DataFrame(_to, index=full_idx, columns=sectors)


def make_results(sig_map):
    out = {}
    for t in sectors:
        rr = ret_mat[t]; co = rr * 0.3; oc = (1 + rr) / (1 + co) - 1
        rot = pd.DataFrame({k: v[t] for k, v in sig_map.items()}, index=full_idx)
        rot["BETA"] = 1.0
        out[t] = {"state": pd.Series("RISK_ON", index=eval_idx), "target_pos": pd.Series(0.8, index=eval_idx),
                  "score_pct": pd.Series(0.5, index=eval_idx), "bh_ret": rr.reindex(eval_idx), "ret_co": co.reindex(eval_idx),
                  "ret_oc": oc.reindex(eval_idx), "rot_raw": rot, "ret_cc_full": rr, "strategy_ret": rr.reindex(eval_idx) * 0.5,
                  "pos_exec": pd.Series(0.5, index=eval_idx),
                  "px_open": pd.Series(100.0, index=eval_idx), "px_close": pd.Series(100.0, index=eval_idx)}
    return out


res_fake = dict(res); res_fake["sig"] = res["sig"].assign(target_pos=pd.Series(1.0, index=res["sig"].index))
_spy_bt = S.portfolio_backtest(pd.DataFrame({"SPY": 1.0}, index=eval_idx), res["bt"][["ret_co"]].rename(columns={"ret_co": "SPY"}),
                               res["bt"][["ret_oc"]].rename(columns={"ret_oc": "SPY"}), None, float(res["cfg"].COST_BPS),
                               init_exec=pd.Series({"SPY": 1.0}), init_prev=pd.Series({"SPY": 1.0}),
                               lev_spread_bps=float(getattr(res["cfg"], "LEVERAGE_SPREAD_BPS", 0.0)))
res_fake["bt"] = res["bt"].copy(); res_fake["bt"].loc[eval_idx, "strategy_ret"] = _spy_bt["strategy_ret"].values

# ---- 1) 증거 등급 ----
# 약한 신호(RESID_MOM_12_1 자리)가 '외부검증 지원'으로 채택되도록 외부검증을 몽키패치(FF49 없이 t=5.0 고정)
_orig_ext = S.external_validation_ff49
S.external_validation_ff49 = lambda scfg, M_, years, *a, **k: ({y: {"RESID_MOM_12_1": 5.0} for y in years}, pd.DataFrame())
try:
    results = make_results({"SCORE_PCT": score_like, "RESID_MOM_12_1": weak})
    base = dict(SECTORS=sectors, ROTATION_SELECT_MODE="strict", USE_EXTERNAL_VALIDATION=True, ROTATION_INCLUDE_SPY_CANDIDATE=False,
                ROTATION_SIGNALS=("SCORE_PCT", "RESID_MOM_12_1"), ROTATION_MIN_AGREE=2, ROTATION_AVOID_VALIDATE=False)
    a_ev = S.build_sector_allocation(results, res_fake, eval_idx, S.SectorConfig(ROTATION_EVIDENCE_TIER=True, **base), M)
    a_v8 = S.build_sector_allocation(results, res_fake, eval_idx, S.SectorConfig(ROTATION_EVIDENCE_TIER=False, **base), M)
finally:
    S.external_validation_ff49 = _orig_ext
wf = a_ev["wf"]
yrs = sorted(wf["selected_by_year"])
grades = {y: {n: wf["basis_by_year"][y][n].split("(")[0] for n in wf["selected_by_year"][y]} for y in yrs}
assert all(grades[y].get("SCORE_PCT") == "엄격" for y in yrs), grades
n_weak_years = sum(1 for y in yrs if grades[y].get("RESID_MOM_12_1") == "외부검증 지원")
assert n_weak_years >= 5, f"약한 신호가 외부검증 지원으로 채택된 해가 {n_weak_years}개뿐(픽스처 t 범위 확인)"
assert all(wf["tier_by_year"][y] == "엄격" and wf["selected_eff_by_year"][y] == ["SCORE_PCT"] for y in yrs), (wf["tier_by_year"], wf["selected_eff_by_year"])
sl = wf["selection_log"]
assert (sl.loc[sl["신호"] == "SCORE_PCT", "리더 판단 사용"] == "사용(엄격 단독)").all()
weak_rows = sl[(sl["신호"] == "RESID_MOM_12_1") & (sl["채택"] == "채택")]
assert (weak_rows["리더 판단 사용"] == "미사용(엄격 신호 우선)").all(), weak_rows["리더 판단 사용"].value_counts()
lead_ev = (a_ev["tier"] == "리더"); lead_v8 = (a_v8["tier"] == "리더")
assert lead_ev.sum() > lead_v8.sum() > 0, (lead_ev.sum(), lead_v8.sum())
# 증거 등급: 리더 판단 신호는 1개(엄격 단독) → 리더 시작일 득표 1/1, 리더는 SCORE_PCT 평활 순위 1위
assert (a_ev["n_selected"] == 1).all()
starts = lead_ev & ((a_ev["tier"].shift(1) != "리더") | (a_ev["leader"] != a_ev["leader"].shift(1)))
assert (a_ev["votes_leader"][starts] == 1).all()
sc_rank_s = S._cs_rank01(score_like.reindex(eval_idx)).rolling(21, min_periods=1).mean()
top_sc = sc_rank_s.idxmax(axis=1)
assert (a_ev["leader"][starts] == top_sc[starts]).all()
# v0.8 방식에서는 두 신호가 같은 1위를 지목한 날만 리더 시작(득표 2)
starts8 = lead_v8 & ((a_v8["tier"].shift(1) != "리더") | (a_v8["leader"] != a_v8["leader"].shift(1)))
assert (a_v8["votes_leader"][starts8] == 2).all()
# 엄격 신호가 없는 해(약한 신호만): 교차확인 유지 — 약한 신호 2개(둘 다 외부 지원)만 있는 픽스처
S.external_validation_ff49 = lambda scfg, M_, years, *a, **k: ({y: {"RESID_MOM_12_1": 5.0, "REL_MOM_126": 5.0} for y in years}, pd.DataFrame())
try:
    weak2 = weak + pd.DataFrame(np.random.default_rng(92).normal(0, 0.02, (nf, 11)), index=full_idx, columns=sectors)   # weak와 상관 높은 약한 신호(t 1~1.7)
    results_w = make_results({"RESID_MOM_12_1": weak, "REL_MOM_126": weak2})
    base_w = dict(base); base_w["ROTATION_SIGNALS"] = ("RESID_MOM_12_1", "REL_MOM_126")
    a_w = S.build_sector_allocation(results_w, res_fake, eval_idx, S.SectorConfig(ROTATION_EVIDENCE_TIER=True, **base_w), M)
finally:
    S.external_validation_ff49 = _orig_ext
wfw = a_w["wf"]
weak_years = [y for y in yrs if wfw["tier_by_year"][y] == "약함"]
assert len(weak_years) >= 5, wfw["tier_by_year"]
starts_w = (a_w["tier"] == "리더") & ((a_w["tier"].shift(1) != "리더") | (a_w["leader"] != a_w["leader"].shift(1)))
starts_w = starts_w & a_w["tier"].index.year.isin(weak_years)
assert (a_w["votes_leader"][starts_w] >= 2).all()
print(f"[1/3] 증거 등급: 엄격+약한 채택 해 → 엄격 단독 판단(리더 {int(lead_ev.sum())}일, 시작일 득표 1/1, 리더=SCORE 1위) vs v0.8 교차확인(리더 "
      f"{int(lead_v8.sum())}일, 시작일 득표 2/2); 약한 신호만인 해 {len(weak_years)}개는 교차확인 유지 OK")

# ---- 2) 회피 자격 ----
# (a) 통계: 상·하 정보 신호(score_like) → 하위1 t ≤ −2(자격 Y); 상위만 정보(top_only) → 채택되지만 자격 N
results_b = make_results({"SCORE_PCT": score_like, "REL_MOM_126": top_only})
cfg_b = S.SectorConfig(SECTORS=sectors, ROTATION_SELECT_MODE="strict", USE_EXTERNAL_VALIDATION=False, ROTATION_INCLUDE_SPY_CANDIDATE=False,
                       ROTATION_SIGNALS=("SCORE_PCT", "REL_MOM_126"), ROTATION_MIN_AGREE=1, ROTATION_EVIDENCE_TIER=False, ROTATION_AVOID_VALIDATE=True)
a_b = S.build_sector_allocation(results_b, res_fake, eval_idx, cfg_b, M)
slb = a_b["wf"]["selection_log"]
sc_rows = slb[slb["신호"] == "SCORE_PCT"]; to_rows = slb[slb["신호"] == "REL_MOM_126"]
assert (sc_rows["채택"] == "채택").all() and (to_rows["채택"] == "채택").all(), (sc_rows["채택"].unique(), to_rows["채택"].unique())
assert (sc_rows["NW-HAC t(하위1)"] <= -2.0).all() and (sc_rows["회피 자격"] == "Y").all(), sc_rows[["적용연도", "NW-HAC t(하위1)", "회피 자격"]]
assert (to_rows["NW-HAC t(상위1)"] >= 2.0).all() and (to_rows["NW-HAC t(하위1)"] > -2.0).all() and (to_rows["회피 자격"] == "N").all(), \
    to_rows[["적용연도", "NW-HAC t(상위1)", "NW-HAC t(하위1)", "회피 자격"]]
assert all(a_b["wf"]["avoid_by_year"][y] == ["SCORE_PCT"] for y in yrs)
# 13d에도 하위1 통계·판정 열
ic_tab = S.rotation_validation(a_b, cfg_b, M)["ic_table"]
assert "하위1 스프레드(%/지평)" in ic_tab.columns and "NW-HAC t(하위1)" in ic_tab.columns and "꼴찌 판별(하위1 t≤−기준)" in ic_tab.columns
r21 = ic_tab[(ic_tab["신호"] == "SCORE_PCT") & (ic_tab["지평(일)"] == 21)].iloc[0]
assert r21["꼴찌 판별(하위1 t≤−기준)"] == "유의(−)"
# (b) 배관: v0.8 픽스처(SCORE_PCT + 잔차모멘텀, MIN_AGREE=2)에서 회피가 나오는 설정 → avoid_by_year를 비우면 회피 0일
rng2 = np.random.default_rng(80)
drift2 = np.zeros((nf, 11)); z2 = rng2.normal(0, 1, 11) * 0.0006
for i in range(nf):
    z2 = 0.995 * z2 + rng2.normal(0, 1, 11) * 0.00006
    drift2[i] = z2
ret2 = pd.DataFrame(rng2.normal(0.0003, 0.012, (nf, 11)) + drift2, index=full_idx, columns=sectors)
score2 = pd.DataFrame(np.clip(0.5 + drift2 * 150 + rng2.normal(0, 0.05, (nf, 11)), 0, 1), index=full_idx, columns=sectors)
mom2 = pd.DataFrame(drift2 * 100 + rng2.normal(0, 0.06, (nf, 11)), index=full_idx, columns=sectors)
results2 = {}
for t in sectors:
    rr = ret2[t]; co = rr * 0.3; oc = (1 + rr) / (1 + co) - 1
    results2[t] = {"state": pd.Series("RISK_ON", index=eval_idx), "target_pos": pd.Series(0.8, index=eval_idx),
                   "score_pct": pd.Series(0.5, index=eval_idx), "bh_ret": rr.reindex(eval_idx), "ret_co": co.reindex(eval_idx),
                   "ret_oc": oc.reindex(eval_idx), "ret_cc_full": rr, "strategy_ret": rr.reindex(eval_idx) * 0.5, "pos_exec": pd.Series(0.5, index=eval_idx),
                   "rot_raw": pd.DataFrame({"SCORE_PCT": score2[t], "RESID_MOM_12_1": mom2[t], "BETA": 1.0}, index=full_idx),
                   "px_open": pd.Series(100.0, index=eval_idx), "px_close": pd.Series(100.0, index=eval_idx)}
cfg2 = S.SectorConfig(SECTORS=sectors, ROTATION_SELECT_MODE="strict", USE_EXTERNAL_VALIDATION=False, ROTATION_INCLUDE_SPY_CANDIDATE=False,
                      ROTATION_SIGNALS=("SCORE_PCT", "RESID_MOM_12_1"), ROTATION_MIN_AGREE=2, ROTATION_EVIDENCE_TIER=False, ROTATION_AVOID_VALIDATE=True)
a_on = S.build_sector_allocation(results2, res_fake, eval_idx, cfg2, M)
n_avoid_on = int((a_on["tier"] == "회피").sum())
assert n_avoid_on > 0, a_on["tier"].value_counts()       # 두 신호 모두 자격 Y → 회피 발생(v0.8과 동일 35일)
_orig_wf = S.rotation_walkforward_select
def _wf_no_avoid(*a, **k):
    out = _orig_wf(*a, **k)
    out["avoid_by_year"] = {y: [] for y in out["avoid_by_year"]}
    return out
S.rotation_walkforward_select = _wf_no_avoid
try:
    a_off = S.build_sector_allocation(results2, res_fake, eval_idx, cfg2, M)
finally:
    S.rotation_walkforward_select = _orig_wf
assert int((a_off["tier"] == "회피").sum()) == 0 and (a_off["votes_laggard"] == 0).all()
assert int((a_off["tier"] == "리더").sum()) == int((a_on["tier"] == "리더").sum())   # 리더 판단은 영향 없음
print(f"[2/3] 회피 자격: 상·하 정보 신호 하위1 t {sc_rows['NW-HAC t(하위1)'].max():.1f}(Y) / 상위만 정보 신호 상위1 t {to_rows['NW-HAC t(상위1)'].min():.1f}·"
      f"하위1 t {to_rows['NW-HAC t(하위1)'].min():.2f}(N, 채택은 됨), 13d 하위1 열; 자격 신호 없으면 회피 {n_avoid_on}일→0일·리더 불변 OK")

# ---- 3) 13j 배분거래내역 손계산 ----
idx = pd.bdate_range("2024-01-01", periods=12)
cols3 = ["A", "B", "SPY"]
tw3 = pd.DataFrame(0.0, index=idx, columns=cols3)
# 결정일 기준 목표비중: d0~d2 SPY 1.0 / d3~d5 A 0.8 / d6 A 0.6 / d7~d8 현금(E=0) / d9~d10 B 0.5 / d11 B 0.5(마지막 — 청산 없음)
tw3.loc[idx[0:3], "SPY"] = 1.0
tw3.loc[idx[3:6], "A"] = 0.8
tw3.loc[idx[6], "A"] = 0.6
tw3.loc[idx[9:12], "B"] = 0.5
E3 = tw3.sum(axis=1)
oc3 = pd.DataFrame(0.01, index=idx, columns=cols3); oc3["B"] = 0.02; oc3["SPY"] = 0.005
co3 = pd.DataFrame(0.002, index=idx, columns=cols3); co3["B"] = -0.001
tier3 = pd.Series("폴백", index=idx, dtype=object); tier3.iloc[3:7] = "리더"; tier3.iloc[7:9] = "현금"; tier3.iloc[9:] = "리더"
leader3 = pd.Series("", index=idx, dtype=object); leader3.iloc[3:7] = "A"; leader3.iloc[9:] = "B"
alloc3 = {"target_w": tw3, "cols": ["A", "B"], "ret_co": co3, "ret_oc": oc3, "cost_bps": 5.0, "init_exec": None, "init_prev": None,
          "tier": tier3, "leader": leader3, "laggard": pd.Series("", index=idx, dtype=object),
          "votes_leader": pd.Series(1, index=idx), "n_selected": pd.Series(1, index=idx), "E": E3,
          "state": pd.DataFrame({"A": "RISK_ON", "B": "RISK_ON"}, index=idx), "spy_m_ret": pd.Series(0.0, index=idx), "rf_daily": None}
results3 = {"A": {"px_open": pd.Series(np.arange(10, 22, dtype=float), index=idx), "px_close": pd.Series(np.arange(10, 22, dtype=float) + 0.5, index=idx)},
            "B": {"px_open": pd.Series(50.0, index=idx), "px_close": pd.Series(51.0, index=idx)}}
res3 = {"px_dict": {"SPY": pd.DataFrame({"Open": 400.0, "Close": 401.0}, index=idx)}}
tr3, summ3 = S.build_allocation_trades(alloc3, results3, res3)
assert len(tr3) == 3 and list(tr3["자산"]) == ["SPY", "A", "B"], tr3[["자산", "진입일(시가 체결)", "청산일(시가 체결)"]]
spy_r, a_r, b_r = tr3.iloc[0], tr3.iloc[1], tr3.iloc[2]
# SPY: 결정 d0 → 체결 d1, 보유 d1~d3(exec_w>0: d1,d2,d3), 청산 d4 시가
assert spy_r["진입일(시가 체결)"] == str(idx[1].date()) and spy_r["청산일(시가 체결)"] == str(idx[4].date()) and spy_r["보유거래일"] == 3
assert abs(spy_r["자산수익률(구간)"] - ((1.005 * 1.002) ** 3 - 1)) < 1e-4          # oc×co 연쇄 3회(청산일 야간 포함)
assert abs(spy_r["전략기여(%p)"] - 100 * (3 * 0.005 + 3 * 0.002 - (1.0 + 1.0) * 5e-4)) < 1e-6   # 진입·청산 비용 각 1.0×5bp
assert spy_r["진입가(시가)"] == 400.0 and spy_r["청산가(시가)"] == 400.0
assert spy_r["진입 판단"].startswith("폴백") and spy_r["청산 사유"].startswith("리더 교체 → A")
# A: 결정 d3 → 체결 d4, 보유 d4~d7(비중 0.8,0.8,0.8,0.6), 청산 d8 시가(현금 E=0)
assert a_r["진입일(시가 체결)"] == str(idx[4].date()) and a_r["청산일(시가 체결)"] == str(idx[8].date()) and a_r["보유거래일"] == 4
assert abs(a_r["평균비중"] - 0.75) < 1e-9 and a_r["진입비중"] == 0.8 and a_r["최대비중"] == 0.8
gross_a = 0.8 * 0.01 + 0.8 * 0.002 + 0.8 * 0.01 + 0.8 * 0.002 + 0.8 * 0.01 + 0.8 * 0.002 + 0.6 * 0.01 + 0.6 * 0.002   # d4: oc; d5,d6,d7: co(전일 비중)+oc; d8: co(0.6)
# 정확히: d4 exec .8 oc, d5 prev .8 co + exec .8 oc, d6 prev .8 co + exec .8 oc, d7 prev .8 co + exec .6 oc, d8 prev .6 co
gross_a = 0.8 * 0.01 + (0.8 * 0.002 + 0.8 * 0.01) + (0.8 * 0.002 + 0.8 * 0.01) + (0.8 * 0.002 + 0.6 * 0.01) + 0.6 * 0.002
cost_a = (0.8 + 0.2 + 0.6) * 5e-4
assert abs(a_r["전략기여(%p)"] - 100 * (gross_a - cost_a)) < 1e-6, (a_r["전략기여(%p)"], 100 * (gross_a - cost_a))
assert abs(a_r["자산수익률(구간)"] - ((1.01 * 1.002) ** 4 - 1)) < 1e-4
assert a_r["진입가(시가)"] == 14.0 and a_r["청산가(시가)"] == 18.0
assert a_r["진입 판단"].startswith("리더 집중") and a_r["청산 사유"].startswith("현금(E_t=0")
# B: 결정 d9 → 체결 d10, 보유 d10~d11(마지막 날, 청산 없음 → 보유중, 청산가 = 마지막 종가)
assert b_r["진입일(시가 체결)"] == str(idx[10].date()) and b_r["청산일(시가 체결)"] == "보유중" and b_r["보유거래일"] == 2
assert abs(b_r["자산수익률(구간)"] - (1.02 * (1 - 0.001) * 1.02 - 1)) < 1e-4      # d10 oc, d11 co+oc(마지막 종가까지)
assert b_r["청산가(시가)"] == 51.0 and b_r["청산 사유"].startswith("보유중")
# 전략기여 합 == 백테스트 총수익(rf 없음) — 비트 수준
bt3 = S.portfolio_backtest(tw3, co3, oc3, None, 5.0)
assert abs(tr3["전략기여(%p)"].sum() / 100 - bt3["strategy_ret"].sum()) < 2e-5   # 행별 3자리 반올림
assert summ3["n_trades"] == 3 and summ3["max_concurrent"] == 1 and abs(summ3["max_gross"] - 1.0) < 1e-9
print(f"[3/3] 13j 손계산: 구간 3개(SPY 3일·A 4일·B 보유중), 시가 체결가·oc/co 연쇄 수익·전략기여(비용 포함)·판단/사유 정확, "
      f"기여 합 == 백테스트 {bt3['strategy_ret'].sum():.5f} OK")
print(f"ALL PASS ({time.time() - t_all:.1f}s)")
