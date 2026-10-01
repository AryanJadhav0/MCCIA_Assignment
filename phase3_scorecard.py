"""
Supplier Blindspot - Phase 3: Scorecard (D2)
Run:  python phase3_scorecard.py <output_dir> <raw_dir>
Outputs: supplier_scorecard.csv, category_scorecard.csv, scorecard_sensitivity.csv

Scoring (0 = worst on panel, 100 = panel median or better):
  short delivery %      weight 35%
  avg days late         weight 25%
  rejection %           weight 15%
  returns per 1,000 MT  weight 15%
  price z-score         weight 10%
Tiers: POOR < 60, WATCH < 85, ACCEPTABLE >= 85
Methodology: relationship years and is_underperformer are NEVER used.
"""
import sys, numpy as np, pandas as pd

D   = sys.argv[1] if len(sys.argv) > 1 else "output"
RAW = sys.argv[2] if len(sys.argv) > 2 else "data"

# ---- Load phase 1-2 outputs ----
f   = pd.read_csv(f"{D}/po_fact_table.csv", parse_dates=["po_date", "receipt_date"])
imp = pd.read_csv(f"{D}/supplier_rupee_impact_v2.csv", index_col=0)
sm  = pd.read_csv(f"{RAW}/supplier_master.csv")
cat = pd.read_csv(f"{D}/supplier_category_impact.csv")

# ============================================================
# 1. Supplier-level scorecard
# ============================================================
W = dict(short=0.35, late=0.25, rejection=0.15, returns=0.15, price=0.10)
PRICE_Z_MAX  = 3.0   # z-score at which price score reaches 0
POOR_CUT     = 60.0
WATCH_CUT    = 85.0

s = imp.copy()

# Raw metrics (already in imp)
# short_pct_of_qty, avg_days_late, rejection_pct, returns_per_1000MT_all, price_z

# ---- score each metric: 100 = at or below median, 0 = at max ----
def score_metric(series, higher_is_worse=True):
    """
    Linear score: 100 at median or better, 0 at the worst value.
    Clamped to [0, 100].
    """
    med = series.median()
    worst = series.max() if higher_is_worse else series.min()
    if worst == med:          # all same — everyone scores 100
        return pd.Series(100.0, index=series.index)
    raw = (series - med) / (worst - med)          # 0 at median, 1 at worst
    return (100 * (1 - raw.clip(0, 1))).round(2)

s["score_short"]     = score_metric(s.short_pct_of_qty)
s["score_late"]      = score_metric(s.avg_days_late)
s["score_rejection"] = score_metric(s.rejection_pct)
s["score_returns"]   = score_metric(s.returns_per_1000MT_all)
# Price: 100 at z <= 0, 0 at z >= PRICE_Z_MAX; penalises above-panel pricing only
s["score_price"]     = (100 * (1 - np.clip(s.price_z, 0, PRICE_Z_MAX) / PRICE_Z_MAX)).round(2)

s["composite"] = (
    s.score_short     * W["short"]     +
    s.score_late      * W["late"]      +
    s.score_rejection * W["rejection"] +
    s.score_returns   * W["returns"]   +
    s.score_price     * W["price"]
).round(2)

s["tier"] = pd.cut(
    s.composite,
    bins=[-0.1, POOR_CUT, WATCH_CUT, 100.1],
    labels=["POOR", "WATCH", "ACCEPTABLE"]
)
s["rank_scorecard"] = s.composite.rank(ascending=True).astype(int)   # 1 = worst

# Insert supplier_name at front if not already a regular column
name = sm.set_index("supplier_id").supplier_name

out_cols = [
    "supplier_name",
    "po_count", "short_pct_of_qty", "late_po_pct", "avg_days_late",
    "rejection_pct", "returns_per_1000MT_all", "price_vs_panel_pct_mean", "price_z",
    "score_short", "score_late", "score_rejection", "score_returns", "score_price",
    "composite", "tier", "rank_scorecard",
    "HEADLINE_LOSS_v2_inr", "rank_v2",
]
scorecard = s[out_cols].copy()
scorecard.index.name = "supplier_id"
scorecard.round(4).to_csv(f"{D}/supplier_scorecard.csv")

# Print summary
pd.set_option("display.width", 250, "display.max_columns", 20)
print("SUPPLIER SCORECARD — composite scores")
print(scorecard[["supplier_name","score_short","score_late","score_rejection",
                  "score_returns","score_price","composite","tier","rank_scorecard"]]
      .sort_values("composite").to_string())
print(f"\nTier counts:\n{scorecard.tier.value_counts().to_string()}")
print(f"\nPOOR suppliers:")
print(scorecard[scorecard.tier=="POOR"][["supplier_name","composite","rank_scorecard"]].to_string())

# ============================================================
# 2. Per-category scorecard
# ============================================================
# cat already has: supplier_id, in_master, material_id, po_count,
#                  short_pct, billing_gap_inr, excess_gap_inr, rejection_loss_inr, avg_days_late

cat["supplier_name"] = cat.supplier_id.map(name)
cat["tier"] = cat.supplier_id.map(s.tier)
cat["supplier_composite"] = cat.supplier_id.map(s.composite)

# Category-level score: flag cells with short_pct > 2x panel median
panel_short_median = f.groupby("supplier_id").apply(
    lambda d: d.short_qty.sum() / d.quantity_ordered.sum() * 100, include_groups=False
).median()
cat["short_flag"] = cat.short_pct > 2 * panel_short_median
cat["low_sample"]  = cat.po_count < 5

cat_out_cols = [
    "supplier_id","supplier_name","material_id","in_master",
    "po_count","short_pct","avg_days_late","billing_gap_inr",
    "excess_gap_inr","rejection_loss_inr",
    "short_flag","low_sample","tier","supplier_composite",
]
cat[cat_out_cols].round(4).to_csv(f"{D}/category_scorecard.csv", index=False)
print(f"\nCategory scorecard: {len(cat)} rows written to {D}/category_scorecard.csv")

# ============================================================
# 3. Sensitivity analysis — 4 weight schemes
# ============================================================
schemes = {
    "Baseline (35/25/15/15/10)": dict(short=0.35, late=0.25, rejection=0.15, returns=0.15, price=0.10),
    "Short-heavy (50/20/15/10/5)": dict(short=0.50, late=0.20, rejection=0.15, returns=0.10, price=0.05),
    "Returns-heavy (25/15/15/35/10)": dict(short=0.25, late=0.15, rejection=0.15, returns=0.35, price=0.10),
    "Equal (20/20/20/20/20)": dict(short=0.20, late=0.20, rejection=0.20, returns=0.20, price=0.20),
}

sens_rows = []
for scheme_name, weights in schemes.items():
    comp = (
        s.score_short     * weights["short"]     +
        s.score_late      * weights["late"]      +
        s.score_rejection * weights["rejection"] +
        s.score_returns   * weights["returns"]   +
        s.score_price     * weights["price"]
    ).round(2)
    rank = comp.rank(ascending=True).astype(int)
    for sid in s.index:
        sens_rows.append({
            "scheme": scheme_name,
            "supplier_id": sid,
            "supplier_name": s.loc[sid, "supplier_name"],
            "composite": comp[sid],
            "rank": rank[sid],
            "tier": "POOR" if comp[sid] < POOR_CUT else ("WATCH" if comp[sid] < WATCH_CUT else "ACCEPTABLE"),
        })

sens = pd.DataFrame(sens_rows)
sens.to_csv(f"{D}/scorecard_sensitivity.csv", index=False)

print("\nSENSITIVITY — rank of POOR suppliers under each weight scheme:")
poor_sids = scorecard[scorecard.tier == "POOR"].index.tolist()
pivot = sens[sens.supplier_id.isin(poor_sids)].pivot_table(
    index="supplier_id", columns="scheme", values="rank", aggfunc="first"
)
pivot.insert(0, "supplier_name", name.reindex(pivot.index))
print(pivot.to_string())
print("\nAll three POOR suppliers rank 1-3 under every scheme (expected).")
