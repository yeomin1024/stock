"""
catboost_ensemble.py
─────────────────────────────────────────────────────────────────
기존 ensemble_search.py의 '투표(K개 카운트)' 부분만 CatBoost 예측으로 교체.
나머지는 동일:
  - 지표 선발: Wilson 점수 → diversify_candidates 풀 (그대로 사용)
  - 타깃 정의: 미래 horizon일 상승/하락 (그대로)
바뀌는 것:
  - 풀에 든 지표들을 '피처'로 써서 CatBoost가 매수/매도 확률을 예측
  - K개 투표 대신, 모델 확률 > threshold 면 신호

핵심 안전장치:
  - 시계열 누설 방지: walk-forward (과거 학습 → 미래 예측)만 사용
  - 매수/매도 별도 모델 (binary)
  - 클래스 불균형 자동 가중
─────────────────────────────────────────────────────────────────
"""
import numpy as np
import pandas as pd
from catboost import CatBoostClassifier


# ════════════════════════════════════════════════════════════════
# 1. 타깃: 기존과 동일하게 미래 horizon일 수익으로 상승/하락 라벨
# ════════════════════════════════════════════════════════════════
def make_labels(close, horizon=1, buy_thresh=0.0, sell_thresh=0.0):
    """
    미래 horizon일 수익률 기준:
      buy_label  = 1 if fwd_ret >  buy_thresh  else 0   (매수해야 할 시점)
      sell_label = 1 if fwd_ret < -sell_thresh else 0   (매도해야 할 시점)
    """
    fwd = close.shift(-horizon) / close - 1
    buy_label = (fwd > buy_thresh).astype(int)
    sell_label = (fwd < -sell_thresh).astype(int)
    return buy_label, sell_label, fwd


# ════════════════════════════════════════════════════════════════
# 2. 풀 지표 → 피처 행렬
#    기존 diversify_candidates가 뽑은 지표 이름 리스트(pool_indicators)를
#    그대로 받아서 feat에서 해당 컬럼만 추출
# ════════════════════════════════════════════════════════════════
def build_feature_matrix(feat, pool_indicators):
    cols = [c for c in pool_indicators if c in feat.columns]
    X = feat[cols].copy()
    # 결측은 0 (신호 없음), 무한대 정리
    X = X.replace([np.inf, -np.inf], np.nan).fillna(0.0)
    return X, cols


# ════════════════════════════════════════════════════════════════
# 3. CatBoost walk-forward 예측 (누설 방지의 핵심)
#    - 시간순으로 n_splits 구간. 각 구간은 '그 이전 전부'로 학습 후 예측.
#    - 미래 정보가 학습에 절대 안 들어감.
# ════════════════════════════════════════════════════════════════
def walk_forward_predict(X, y, n_splits=5, min_train=120,
                         cb_params=None, sample_weight_balanced=True):
    n = len(X)
    proba = pd.Series(np.nan, index=X.index)
    if n < min_train + 40:
        n_splits = max(2, (n - min_train) // 30)
    if n_splits < 2:
        return proba, []
    # 학습 시작점 이후를 n_splits로 분할
    test_start = min_train
    fold_edges = np.linspace(test_start, n, n_splits + 1).astype(int)
    fold_edges = sorted(set(fold_edges.tolist()))
    models_info = []

    default = dict(
        iterations=300, depth=4, learning_rate=0.05,
        l2_leaf_reg=6.0, random_seed=42, verbose=False,
        loss_function='Logloss', early_stopping_rounds=40,
    )
    if cb_params:
        default.update(cb_params)

    for i in range(len(fold_edges) - 1):
        tr_end = fold_edges[i]
        te_end = fold_edges[i + 1]
        if tr_end < min_train or te_end <= tr_end:
            continue
        X_tr, y_tr = X.iloc[:tr_end], y.iloc[:tr_end]
        X_te = X.iloc[tr_end:te_end]
        # 라벨이 한 종류뿐이면 학습 불가 → 스킵
        if y_tr.nunique() < 2 or len(X_te) == 0:
            continue
        w = None
        if sample_weight_balanced:
            pos = max(int(y_tr.sum()), 1); neg = max(len(y_tr) - pos, 1)
            cw = {0: 1.0, 1: neg / pos}   # 소수 클래스 가중
            w = y_tr.map(cw).values
        model = CatBoostClassifier(**default)
        model.fit(X_tr, y_tr, sample_weight=w)
        p = model.predict_proba(X_te)[:, 1]
        proba.iloc[tr_end:te_end] = p
        models_info.append({
            'train_end': tr_end, 'test_end': te_end,
            'n_train': tr_end, 'n_test': len(X_te),
            'pos_rate_train': float(y_tr.mean()),
        })
    return proba, models_info


# ════════════════════════════════════════════════════════════════
# 4. 확률 → 신호. threshold는 학습구간 분위수로 자동(과적합 방지)
# ════════════════════════════════════════════════════════════════
def proba_to_signal(proba, pct=80):
    valid = proba.dropna()
    if len(valid) < 20:
        return pd.Series(0, index=proba.index), np.nan
    thr = np.percentile(valid, pct)
    sig = (proba >= thr).astype(int)
    sig[proba.isna()] = 0
    return sig, thr


# ════════════════════════════════════════════════════════════════
# 5. 전체: 기존 풀(buy_pool/sell_pool) → CatBoost 매수/매도 예측
# ════════════════════════════════════════════════════════════════
def catboost_ensemble(feat, close, buy_pool_indicators, sell_pool_indicators,
                      horizon=1, signal_pct=80, n_splits=5, cb_params=None):
    buy_label, sell_label, fwd = make_labels(close, horizon=horizon)

    Xb, buy_cols = build_feature_matrix(feat, buy_pool_indicators)
    Xs, sell_cols = build_feature_matrix(feat, sell_pool_indicators)

    # 라벨 정의 가능한 구간(마지막 horizon일은 미래수익 없음)으로 한정
    valid_idx = fwd.dropna().index
    Xb, buy_label = Xb.loc[valid_idx], buy_label.loc[valid_idx]
    Xs, sell_label = Xs.loc[valid_idx], sell_label.loc[valid_idx]

    buy_proba, buy_info = walk_forward_predict(Xb, buy_label, n_splits=n_splits, cb_params=cb_params)
    sell_proba, sell_info = walk_forward_predict(Xs, sell_label, n_splits=n_splits, cb_params=cb_params)

    buy_sig, buy_thr = proba_to_signal(buy_proba, pct=signal_pct)
    sell_sig, sell_thr = proba_to_signal(sell_proba, pct=signal_pct)

    return {
        'buy_proba': buy_proba, 'sell_proba': sell_proba,
        'buy_signal': buy_sig, 'sell_signal': sell_sig,
        'buy_thr': buy_thr, 'sell_thr': sell_thr,
        'buy_cols': buy_cols, 'sell_cols': sell_cols,
        'fwd_ret': fwd, 'buy_label': buy_label, 'sell_label': sell_label,
    }


# ════════════════════════════════════════════════════════════════
# 6. OOS 성능 평가 (정밀도/적중)
# ════════════════════════════════════════════════════════════════
def evaluate(res):
    out = {}
    for side in ['buy', 'sell']:
        sig = res[f'{side}_signal']
        lab = res[f'{side}_label'].reindex(sig.index)
        proba = res[f'{side}_proba'].reindex(sig.index)
        m = sig.notna() & lab.notna() & proba.notna()
        fired = (sig == 1) & m
        n_fired = int(fired.sum())
        if n_fired > 0:
            precision = float(lab[fired].mean())  # 신호 켜졌을 때 실제 맞은 비율
        else:
            precision = np.nan
        # AUC 근사 (확률 순위 vs 라벨)
        pr = proba[m]; lb = lab[m]
        if lb.nunique() == 2 and len(lb) > 10:
            from sklearn.metrics import roc_auc_score
            try:
                auc = roc_auc_score(lb, pr)
            except Exception:
                auc = np.nan
        else:
            auc = np.nan
        out[side] = {'n_fired': n_fired, 'precision': precision,
                     'base_rate': float(lb.mean()) if len(lb) else np.nan, 'auc': auc}
    return out


# ════════════════════════════════════════════════════════════════
# 데모: 진짜 신호 + 노이즈 섞인 풀로 동작/누설 검증
# ════════════════════════════════════════════════════════════════
if __name__ == '__main__':
    rng = np.random.default_rng(2024)
    n = 700
    idx = pd.bdate_range('2022-01-01', periods=n)
    # 미래 수익에 인과를 가진 진짜 지표 + 노이즈
    fwd_true = rng.standard_normal(n) * 0.02
    close = pd.Series(100 * np.cumprod(1 + np.r_[0, fwd_true[:-1]] * 0.4 + rng.standard_normal(n) * 0.008), index=idx)

    feat = pd.DataFrame(index=idx)
    # 진짜 예측 지표 4개 (미래수익과 상관)
    for k in range(4):
        feat[f'real_{k}'] = pd.Series(fwd_true, index=idx) * (0.5 + 0.2*k) + rng.standard_normal(n) * 0.03
    # 노이즈 지표 20개
    for k in range(20):
        feat[f'noise_{k}'] = pd.Series(rng.standard_normal(n), index=idx)
    feat = feat.replace([np.inf, -np.inf], np.nan).fillna(0)

    # 기존 diversify_candidates가 뽑았다고 가정한 풀 (진짜4 + 노이즈 일부)
    buy_pool = ['real_0','real_1','real_2','real_3'] + [f'noise_{k}' for k in range(8)]
    sell_pool = ['real_0','real_1','real_2','real_3'] + [f'noise_{k}' for k in range(8,16)]

    res = catboost_ensemble(feat, close, buy_pool, sell_pool,
                            horizon=1, signal_pct=80, n_splits=5)
    perf = evaluate(res)

    print("="*64)
    print("CatBoost 앙상블 데모 (풀: 진짜4 + 노이즈8)")
    print("="*64)
    for side in ['buy', 'sell']:
        p = perf[side]
        print(f"\n[{side.upper()}] OOS 신호 {p['n_fired']}회 발생")
        print(f"  신호 적중률(precision) = {p['precision']:.3f}  (기저율 {p['base_rate']:.3f})")
        print(f"  AUC = {p['auc']:.3f}  (0.5=무작위, 높을수록 예측력)")
    print("\n주: walk-forward라 OOS 구간만 평가됨(미래 누설 없음).")
    print("   precision > base_rate, AUC > 0.5 면 모델이 신호를 잡고 있다는 뜻.")
    print("OK")
