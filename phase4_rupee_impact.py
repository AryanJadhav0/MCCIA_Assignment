"""
Supplier Blindspot - Phase 4: rupee impact (D3)
Needs the outputs of pipeline.py.  Run:  python phase4_rupee_impact.py [dir_with_pipeline_outputs]
Components (each traceable to PO rows):
  A  Short-delivery billing gap  = invoice_amount_billed - quantity_received x unit_price_quoted
  A* Excess gap                  = A minus panel-baseline shortage (median supplier short %), floored at 0 per supplier
  B  Rejection loss              = rejection_qty x unit_price_quoted   (inside received qty and billed)
  C  Confirmed returns           = returned qty x unit price of that supplier's latest prior receipt of the material,
                                   capped so a batch never returns more than (received - rejected)
  Headline (evidence-backed) loss = A* + B + C        (inferred returns are added in the attribution phase)
  Shown separately, NOT in headline: price premium vs panel (weak signal, see z-score), late-delivery carrying-cost proxy (assumption)
"""
import sys, numpy as np, pandas as pd
D = sys.argv[1] if len(sys.argv) > 1 else "/mnt/user-data/outputs"
f  = pd.read_csv(f"{D}/po_fact_table.csv", parse_dates=["po_date", "receipt_date"])
cr = pd.read_csv(f"{D}/returns_clean.csv", parse_dates=["return_date"])
RAW = sys.argv[2] if len(sys.argv) > 2 else "/mnt/user-data/uploads"
sm = pd.read_csv(f"{RAW}/supplier_master.csv")
name = sm.set_index("supplier_id").supplier_name
# ---- C. confirmed returns, valued and capped -------------------------------------------------
k = cr[cr.supplier_known].copy().sort_values("return_date")
f = f.sort_values("receipt_date")
cap_left = (f.quantity_received - f.rejection_qty).to_dict()       # remaining returnable qty per fact-table row index
rows = []
for r in k.itertuples():
    cand = f[(f.supplier_id == r.supplier_id_traced) & (f.material_id == r.material_id) & (f.receipt_date <= r.return_date)]
    if cand.empty:                                                  # fallback: supplier's earliest receipt of material
        cand = f[(f.supplier_id == r.supplier_id_traced) & (f.material_id == r.material_id)]
    if cand.empty:
        rows.append({**r._asdict(), "matched_po": None, "price_used": np.nan, "qty_counted": 0.0, "value_inr": 0.0, "note": "no receipt for pair"}); continue
    b = cand.iloc[-1]; bi = cand.index[-1]
    q = min(r.quantity_returned, max(cap_left[bi], 0)); cap_left[bi] -= q
    rows.append({**r._asdict(), "matched_po": b.po_id, "price_used": b.unit_price_quoted, "qty_counted": q,
                 "value_inr": q * b.unit_price_quoted, "note": "capped" if q < r.quantity_returned else ""})
rv = pd.DataFrame(rows).drop(columns="Index", errors="ignore")
rv.to_csv(f"{D}/returns_confirmed_valued.csv", index=False)

# ---- Supplier-level table --------------------------------------------------------------------
med_late = f.groupby("supplier_id").days_late_delivery.mean().median()
f["late_cost_proxy_inr"] = f.invoice_amount_billed * 0.12 / 365 * (f.days_late_delivery - med_late).clip(lower=0)  # 12% p.a. carrying cost (assumption)

def agg(d):
    return pd.Series({
        "po_count": len(d), "ordered_value_inr": d.ordered_value.sum(), "billed_inr": d.invoice_amount_billed.sum(),
        "short_pct_of_qty": d.short_qty.sum() / d.quantity_ordered.sum() * 100,
        "late_po_pct": d.is_late.mean() * 100, "avg_days_late": d.days_late_delivery.mean(),
        "rejection_pct": d.rejection_qty.sum() / d.quantity_received.sum() * 100,
        "A_billing_gap_inr": d.billing_gap_inr.sum(),
        "A_baseline_inr": d.baseline_gap_inr.sum(),
        "A_excess_gap_inr": max(d.excess_gap_inr.sum(), 0.0),
        "B_rejection_loss_inr": d.rejection_value_inr.sum(),
        "price_premium_vs_panel_inr": d.premium_vs_panel_inr.sum(),
        "price_vs_panel_pct_mean": d.price_vs_panel_pct.mean(),
        "price_z": d.price_vs_panel_pct.mean() / (d.price_vs_panel_pct.std() / np.sqrt(len(d))),
        "late_cost_proxy_inr": d.late_cost_proxy_inr.sum(),
        "arora_avg_days_late_paying": d.arora_days_late.mean()})
s = f.groupby("supplier_id").apply(agg, include_groups=False)
s["C_confirmed_returns_inr"] = rv.groupby("supplier_id_traced").value_inr.sum().reindex(s.index).fillna(0)
s["C_confirmed_return_count"] = rv.groupby("supplier_id_traced").size().reindex(s.index).fillna(0)
s["HEADLINE_LOSS_inr"] = s.A_excess_gap_inr + s.B_rejection_loss_inr + s.C_confirmed_returns_inr
s["HEADLINE_LOSS_pct_of_billed"] = s.HEADLINE_LOSS_inr / s.billed_inr * 100
s.insert(0, "supplier_name", name.reindex(s.index))
s = s.sort_values("HEADLINE_LOSS_inr", ascending=False)
s["rank"] = range(1, len(s) + 1)
s.round(2).to_csv(f"{D}/supplier_rupee_impact.csv")

# ---- Supplier x category (master categories vs off-master) ------------------------------------
c = f.groupby(["supplier_id", "in_master", "material_id"]).apply(lambda d: pd.Series({
    "po_count": len(d), "short_pct": d.short_qty.sum() / d.quantity_ordered.sum() * 100,
    "billing_gap_inr": d.billing_gap_inr.sum(), "excess_gap_inr": d.excess_gap_inr.sum(),
    "rejection_loss_inr": d.rejection_value_inr.sum(), "avg_days_late": d.days_late_delivery.mean()}), include_groups=False).reset_index()
c["low_sample_flag"] = c.po_count < 5
c.round(2).to_csv(f"{D}/supplier_category_impact.csv", index=False)

# ---- Reconciliation & worked arithmetic -------------------------------------------------------
pd.set_option("display.width", 250, "display.max_columns", 30)
print("TOP 8 BY HEADLINE LOSS (Rs)")
print(s[["supplier_name", "po_count", "A_billing_gap_inr", "A_excess_gap_inr", "B_rejection_loss_inr", "C_confirmed_returns_inr", "HEADLINE_LOSS_inr"]].head(8).round(0).to_string())
print("\nPanel total: gap %.0f | excess %.0f | rejection %.0f | confirmed returns %.0f" %
      (s.A_billing_gap_inr.sum(), s.A_excess_gap_inr.sum(), s.B_rejection_loss_inr.sum(), s.C_confirmed_returns_inr.sum()))
print("Returns: %d confirmed, %d capped, counted qty %.1f of %.1f" % (len(rv), (rv.note == "capped").sum(), rv.qty_counted.sum(), rv.quantity_returned.sum()))
print("\nPRICE SIGNAL (z = mean %% vs panel / std error)  top 5:")
print(s[["supplier_name", "price_vs_panel_pct_mean", "price_z", "price_premium_vs_panel_inr"]].sort_values("price_z", ascending=False).head(5).round(2).to_string())
print("\nArora pays late on avg: %.1f days (bad three: %s)" % (f.arora_days_late.mean(),
      s.loc[["VS03", "VS12", "VS19"], "arora_avg_days_late_paying"].round(1).to_dict()))
w = f[f.supplier_id == s.index[0]]
print("\nWORKED CHECK for", s.index[0], ": sum(billed) %.2f - sum(received x price) %.2f = %.2f" %
      (w.invoice_amount_billed.sum(), w.received_value.sum(), w.invoice_amount_billed.sum() - w.received_value.sum()))
