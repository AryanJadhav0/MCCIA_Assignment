"""
Supplier Blindspot - Phase 2: attribute the 140 returns with a blank supplier_id_traced (35% of score)

MODEL  (Bayesian, three evidence sources, each tested by ablation in cross-validation)
  P(s | return) ∝  prior_s  x  P(material | group(s))  x  P(qty band | group(s))   [x date-proximity term, tested and rejected]
    prior_s   = lambda_s * MT_s : supplier's return rate per MT received (shrunk to the panel rate) x its volume.
                lambda_s learned only from known returns.
    group(s)  = "high-return-rate" if lambda_s >= 3 x median lambda (learned from data; NOT from is_underperformer)
    P(material | group), P(qty band | group): Laplace-smoothed frequencies from known returns
    date term = batch-exposure: MT of that material received from s in the W days before the return, exp-decayed
  is_underperformer and years_of_relationship are never used.
Validation: repeated 5-fold CV on the 240 known returns (priors/likelihoods re-learned inside each training fold).
Run:  python phase2_attribution.py <dir_with_earlier_outputs>
"""
import sys, os, numpy as np, pandas as pd
from math import erf
D = sys.argv[1] if len(sys.argv) > 1 else "/mnt/user-data/outputs"
rng = np.random.default_rng(42)
f  = pd.read_csv(f"{D}/po_fact_table.csv", parse_dates=["po_date", "receipt_date"]).sort_values("receipt_date").reset_index(drop=True)
cr = pd.read_csv(f"{D}/returns_clean.csv", parse_dates=["return_date"]).reset_index(drop=True)
cv_prev = pd.read_csv(f"{D}/returns_confirmed_valued.csv")
S = sorted(f.supplier_id.unique()); sidx = {s: i for i, s in enumerate(S)}; nS = len(S)
f["si"] = f.supplier_id.map(sidx)
X = f.groupby("si").quantity_received.sum().reindex(range(nS)).values
known = cr.supplier_known.values
true_si = np.array([sidx[s] if isinstance(s, str) else -1 for s in cr.supplier_id_traced])
K_idx = np.where(known)[0]; U_idx = np.where(~known)[0]
mats = sorted(f.material_id.unique()); midx = {m: i for i, m in enumerate(mats)}
mi = cr.material_id.map(midx).values; qb = np.minimum(np.floor(cr.quantity_returned.values).astype(int), 4)   # 5 qty bands: <1,1-2,2-3,3-4,4-5
ALPHA, BETA, GROUP_MULT = 0.5, 1.0, 3.0   # ALPHA chosen for aggregate calibration (rupee totals); logloss differs by only 0.02 across 0.5-3.0

# ---------- date-proximity (exposure) matrices, used for the ablation and the proximity test ----------
DATE_SETTINGS = {"W60_tau30": (60, 30), "W90_tau90": (90, 90), "W180_tau90": (180, 90)}
E = {k: np.zeros((len(cr), nS)) for k in DATE_SETTINGS}; E90flat = np.zeros((len(cr), nS)); last_age = np.full((len(cr), nS), 1e9)
by_mat = {m: g for m, g in f.groupby("material_id")}
for i, r in cr.iterrows():
    g = by_mat[r.material_id]; g = g[g.receipt_date <= r.return_date]
    age = (r.return_date - g.receipt_date).dt.days.values; si = g.si.values; q = g.quantity_received.values
    for k, (W, t) in DATE_SETTINGS.items():
        mk = age <= W; E[k][i] = np.bincount(si[mk], weights=q[mk] * np.exp(-age[mk] / t), minlength=nS)
    mk = age <= 90; E90flat[i] = np.bincount(si[mk], weights=q[mk], minlength=nS)
    for s_, a_ in zip(si, age): last_age[i, s_] = min(last_age[i, s_], a_)

# ---------- model ----------
def learn(train):
    n = np.bincount(true_si[train], minlength=nS).astype(float); lam0 = n.sum() / X.sum()
    lam = (n + ALPHA) / (X + ALPHA / lam0); grp = lam >= GROUP_MULT * np.median(lam)
    gt = grp[true_si[train]].astype(int)
    Cm = np.full((2, len(mats)), BETA); Cq = np.full((2, 5), BETA)
    for g_, m_, q_ in zip(gt, mi[train], qb[train]): Cm[g_, m_] += 1; Cq[g_, q_] += 1
    return {"lam": lam, "grp": grp, "Lm": Cm / Cm.sum(1, keepdims=True), "Lq": Cq / Cq.sum(1, keepdims=True)}

def predict(rows, M, use_mat=False, use_qty=False, date=None, eps=0.2):
    g = M["grp"].astype(int); w = np.tile(M["lam"] * X, (len(rows), 1))
    if use_mat: w = w * M["Lm"][g][:, mi[rows]].T
    if use_qty: w = w * M["Lq"][g][:, qb[rows]].T
    if date:
        Em = E[date][rows]; mean = Em.mean(1, keepdims=True); mean[mean == 0] = 1
        w = w * (Em / mean + eps)
    return w / w.sum(1, keepdims=True)

VARIANTS = {"A rate x volume only (no material/date/qty)": dict(),
            "B + material": dict(use_mat=True),
            "C + quantity band": dict(use_qty=True),
            "D + material + quantity  [CHOSEN MODEL]": dict(use_mat=True, use_qty=True),
            "E D + date proximity (W60/tau30)": dict(use_mat=True, use_qty=True, date="W60_tau30"),
            "F D + date proximity (W90/tau90)": dict(use_mat=True, use_qty=True, date="W90_tau90"),
            "G D + date proximity (W180/tau90)": dict(use_mat=True, use_qty=True, date="W180_tau90")}
def logloss(P, y): return -np.mean(np.log(np.clip(P[np.arange(len(y)), y], 1e-12, None)))
def top1(P, y):
    return np.mean([(P[i, y[i]] == P[i].max()) / (P[i] == P[i].max()).sum() for i in range(len(y))])
def topk(P, y, k): o = np.argsort(-P, axis=1)[:, :k]; return np.mean([y[i] in o[i] for i in range(len(y))])
def folds(idx, k, seed): return np.array_split(np.random.default_rng(seed).permutation(idx), k)

# ---------- repeated 5-fold CV ----------
REPS = 5; names = list(VARIANTS) + ["Baseline: random among suppliers with a batch in 90d", "Baseline: latest receipt wins"]
acc = {n: [] for n in names}; oof_last = {}; grp_acc = []; grp_base = []
for rep in range(REPS):
    P = {n: np.zeros((len(cr), nS)) for n in names}; pg = np.zeros(len(cr)); ing = np.zeros(len(cr), bool)
    for te in folds(K_idx, 5, 100 + rep):
        tr = np.setdiff1d(K_idx, te); M = learn(tr)
        for n, kw in VARIANTS.items(): P[n][te] = predict(te, M, **kw)
        cand = (E90flat[te] > 0).astype(float); cand[cand.sum(1) == 0] = 1; P[names[-2]][te] = cand / cand.sum(1, keepdims=True)
        lr = -last_age[te]; L = (lr == lr.max(1, keepdims=True)).astype(float); P[names[-1]][te] = L / L.sum(1, keepdims=True)
        chosen = P["D + material + quantity  [CHOSEN MODEL]".replace("D", "D", 1)] if False else P["D + material + quantity  [CHOSEN MODEL]"]
        pg[te] = chosen[te][:, M["grp"]].sum(1); ing[te] = M["grp"][true_si[te]]
    y = true_si[K_idx]
    for n in names:
        Pk = P[n][K_idx]; acc[n].append((top1(Pk, y), topk(Pk, y, 3), logloss(Pk, y) if "latest" not in n else np.nan))
    grp_acc.append(np.mean((pg[K_idx] >= 0.5) == ing[K_idx])); grp_base.append(max(ing[K_idx].mean(), 1 - ing[K_idx].mean()))
    oof_last = dict(P=P, pg=pg, ing=ing)
res = pd.DataFrame([{"method": n, "top1_acc": np.mean([a[0] for a in acc[n]]), "top3_acc": np.mean([a[1] for a in acc[n]]),
                     "logloss": np.nanmean([a[2] for a in acc[n]]) if "latest" not in n else np.nan} for n in names]).round(3)
print("CROSS-VALIDATION on 240 known returns (mean of %d repeats of 5-fold; lower logloss is better)" % REPS); print(res.to_string(index=False))
print("\nHigh-return-rate GROUP membership: accuracy %.3f vs always-guess-majority %.3f" % (np.mean(grp_acc), np.mean(grp_base)))
res.to_csv(f"{D}/attribution_validation.csv", index=False)

CH = "D + material + quantity  [CHOSEN MODEL]"; P_ = oof_last["P"][CH][K_idx]; y = true_si[K_idx]
agg = pd.DataFrame({"supplier": S, "actual_returns": np.bincount(y, minlength=nS), "model_expected": P_.sum(0).round(1)})
print("\nAggregate recovery of return counts (out-of-fold, last repeat):"); print(agg.sort_values("actual_returns", ascending=False).head(4).to_string(index=False))
oth = agg.drop(agg.sort_values("actual_returns", ascending=False).head(3).index)
print("Other 31 suppliers: mean abs error %.2f returns each" % np.mean(np.abs(oth.actual_returns - oth.model_expected)))
pgk, ingk = oof_last["pg"][K_idx], oof_last["ing"][K_idx]
bins = pd.cut(pgk, [0, .3, .45, .6, .75, 1.01], include_lowest=True)
print("\nReliability of P(high-return-rate group) (out-of-fold):")
print(pd.DataFrame({"bin": bins, "truly_in_group": ingk, "p": pgk}).groupby("bin", observed=True).agg(n=("p", "size"), mean_pred=("p", "mean"), observed=("truly_in_group", "mean")).round(3).to_string())

# date-proximity test
tl = np.array([last_age[i, true_si[i]] for i in K_idx if last_age[i, true_si[i]] < 1e8])
ol = np.concatenate([last_age[i, np.arange(nS) != true_si[i]][last_age[i, np.arange(nS) != true_si[i]] < 1e8] for i in K_idx])
r_ = pd.Series(np.concatenate([tl, ol])).rank().values; U = r_[:len(tl)].sum() - len(tl) * (len(tl) + 1) / 2
mu = len(tl) * len(ol) / 2; sd = np.sqrt(len(tl) * len(ol) * (len(tl) + len(ol) + 1) / 12); z = (U - mu) / sd; pv = 2 * (1 - 0.5 * (1 + erf(abs(z) / np.sqrt(2))))
print("\nDATE-PROXIMITY TEST: days from return back to the supplier's latest batch of that material")
print("  true supplier median %.0f d (n=%d) vs other suppliers median %.0f d: Mann-Whitney z=%.2f, p=%.2f -> no signal" % (np.median(tl), len(tl), np.median(ol), z, pv))
print("  traced supplier had NO prior batch of that material for %d of %d known returns (hard date/material filter would wrongly exclude them)" % (len(K_idx) - len(tl), len(K_idx)))

# ---------- final fit on all 240 known; attribute the 140 unknown ----------
M = learn(K_idx); Pu = predict(U_idx, M, use_mat=True, use_qty=True); grp = M["grp"]
print("\nLearned high-return-rate group:", [S[i] for i in np.where(grp)[0]], "| lambda ratio to median: %s" % np.round((M["lam"] / np.median(M["lam"]))[grp], 1))
P_oof_p1 = P_.max(1); hit = P_.argmax(1) == y
def tier(p): return np.where(p >= 0.40, "High", np.where(p >= 0.20, "Medium", "Low"))
rel = pd.DataFrame({"tier": tier(P_oof_p1), "hit": hit, "p1": P_oof_p1}).groupby("tier").agg(n=("hit", "size"), top1_accuracy=("hit", "mean"), mean_p1=("p1", "mean")).round(3)
print("\nTop-1 supplier reliability by tier (out-of-fold):"); print(rel.to_string())
rows = []
for j, i in enumerate(U_idx):
    p = Pu[j]; o = np.argsort(-p)[:3]
    rows.append({"return_id": cr.return_id[i], "return_date": cr.return_date[i].date(), "client_id": cr.client_id[i], "material_id": cr.material_id[i],
                 "quantity_returned": cr.quantity_returned[i], "reason": cr.reason[i],
                 "top1_supplier": S[o[0]], "top1_prob": round(p[o[0]], 3), "top2_supplier": S[o[1]], "top2_prob": round(p[o[1]], 3),
                 "top3_supplier": S[o[2]], "top3_prob": round(p[o[2]], 3), "prob_high_return_group": round(p[grp].sum(), 3),
                 "supplier_level_tier": str(tier(np.array([p[o[0]]]))[0]),
                 "group_level_call": "High-return-rate group" if p[grp].sum() >= 0.5 else "Other suppliers",
                 "group_level_confidence": "Strong" if abs(p[grp].sum() - 0.5) >= 0.25 else ("Moderate" if abs(p[grp].sum() - 0.5) >= 0.12 else "Weak")})
att = pd.DataFrame(rows); att.to_csv(f"{D}/returns_attributed.csv", index=False)
print("\nUnknown returns: mean P(group)=%.2f (known share %.2f) | group call: %s | strength: %s" % (att.prob_high_return_group.mean(), grp[true_si[K_idx]].mean(),
      att.group_level_call.value_counts().to_dict(), att.group_level_confidence.value_counts().to_dict()))

# ---------- value inferred returns (probability-weighted, batch-capped; consistent with Phase 4) ----------
cap = (f.quantity_received - f.rejection_qty).values.astype(float); pos = {p: i for i, p in enumerate(f.po_id)}
for r in cv_prev.itertuples():
    if isinstance(r.matched_po, str): cap[pos[r.matched_po]] -= r.qty_counted
cap = np.maximum(cap, 0); V = np.zeros((len(U_idx), nS)); long = []
for j in np.argsort(cr.return_date.values[U_idx]):
    i = U_idx[j]; r = cr.loc[i]
    for s in np.where(Pu[j] > 1e-4)[0]:
        c = f[(f.si == s) & (f.material_id == r.material_id) & (f.receipt_date <= r.return_date)]
        if c.empty: c = f[(f.si == s) & (f.material_id == r.material_id)]
        if c.empty: continue
        bi = c.index[-1]; q = min(r.quantity_returned, cap[bi]); cap[bi] -= Pu[j, s] * q; V[j, s] = q * f.unit_price_quoted[bi]
        long.append({"return_id": r.return_id, "supplier_id": S[s], "prob": round(Pu[j, s], 4), "matched_po": f.po_id[bi], "qty_if_true": round(q, 3),
                     "price": f.unit_price_quoted[bi], "value_if_true_inr": round(V[j, s], 2), "expected_value_inr": round(Pu[j, s] * V[j, s], 2)})
pd.DataFrame(long).to_csv(f"{D}/returns_inferred_valued.csv", index=False)
exp_val = (Pu * V).sum(0); exp_cnt = Pu.sum(0); draws = np.zeros((4000, nS))
for j in range(len(U_idx)):
    ch = rng.choice(nS, size=4000, p=Pu[j] / Pu[j].sum())
    for s in np.unique(ch): draws[:, s] += (ch == s) * V[j, s]
lo, hi = np.percentile(draws, [10, 90], axis=0)

imp = pd.read_csv(f"{D}/supplier_rupee_impact.csv", index_col=0)
imp["D_inferred_returns_expected_inr"] = pd.Series(exp_val, index=S).round(0)
imp["D_inferred_returns_p10_inr"] = pd.Series(lo, index=S).round(0); imp["D_inferred_returns_p90_inr"] = pd.Series(hi, index=S).round(0)
imp["D_inferred_return_count_expected"] = pd.Series(exp_cnt, index=S).round(1)
imp["HEADLINE_LOSS_v2_inr"] = imp.A_excess_gap_inr + imp.B_rejection_loss_inr + imp.C_confirmed_returns_inr + imp.D_inferred_returns_expected_inr
imp["returns_per_1000MT_all"] = ((imp.C_confirmed_return_count + imp.D_inferred_return_count_expected) / pd.Series(X, index=S) * 1000).round(2)
imp = imp.sort_values("HEADLINE_LOSS_v2_inr", ascending=False); imp["rank_v2"] = range(1, len(imp) + 1)
imp.to_csv(f"{D}/supplier_rupee_impact_v2.csv")
if os.path.exists(f"{D}/attribution_cv_grid.csv"): os.remove(f"{D}/attribution_cv_grid.csv")
pd.set_option("display.width", 250, "display.max_columns", 30)
print("\nTOP 5 (Rs): confirmed vs inferred returns, with 80% range on the inferred part")
print(imp[["supplier_name", "C_confirmed_returns_inr", "D_inferred_returns_expected_inr", "D_inferred_returns_p10_inr", "D_inferred_returns_p90_inr", "HEADLINE_LOSS_v2_inr"]].head(5).round(0).to_string())
print("\nReconciliation: expected inferred value Rs %.0f | to high-rate group %.1f%% | all returns (confirmed + inferred) Rs %.0f" %
      (exp_val.sum(), exp_val[grp].sum() / exp_val.sum() * 100, imp.C_confirmed_returns_inr.sum() + exp_val.sum()))
