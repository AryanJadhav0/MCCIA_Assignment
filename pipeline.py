"""
Supplier Blindspot - Phase 1: data foundation (D1)
Run:  python pipeline.py [input_dir] [output_dir]
Outputs: po_fact_table.csv, data_quality_log.csv, supplier_category_map.csv, returns_clean.csv
"""
import sys, pandas as pd, numpy as np
IN  = sys.argv[1] if len(sys.argv) > 1 else "/mnt/user-data/uploads"
OUT = sys.argv[2] if len(sys.argv) > 2 else "/mnt/user-data/outputs"
log = []
def check(name, passed, detail=""):
    log.append({"check": name, "result": "PASS" if passed else "FAIL", "detail": detail})

# ---------- 1. Load ----------
po  = pd.read_csv(f"{IN}/purchase_orders.csv", parse_dates=["po_date", "delivery_promised_date"])
gr  = pd.read_csv(f"{IN}/goods_receipts.csv", parse_dates=["receipt_date"])
cr  = pd.read_csv(f"{IN}/customer_returns.csv", parse_dates=["return_date"])
mp  = pd.read_csv(f"{IN}/market_price_index.csv")
pay = pd.read_csv(f"{IN}/payment_records.csv", parse_dates=["invoice_date", "actual_payment_date"])
sm  = pd.read_csv(f"{IN}/supplier_master.csv")

# ---------- 2. Integrity checks ----------
for n, df, k in [("purchase_orders", po, "po_id"), ("goods_receipts", gr, "gr_id"),
                 ("goods_receipts", gr, "po_id"), ("payment_records", pay, "po_id"),
                 ("customer_returns", cr, "return_id"), ("supplier_master", sm, "supplier_id")]:
    check(f"{n}.{k} unique", df[k].is_unique, f"{df[k].duplicated().sum()} duplicates")
for n, df in [("purchase_orders", po), ("goods_receipts", gr), ("payment_records", pay), ("market_price_index", mp), ("supplier_master", sm)]:
    check(f"{n} no exact duplicate rows", df.duplicated().sum() == 0)
check("PO -> receipt 1:1", set(po.po_id) == set(gr.po_id))
check("PO -> payment 1:1", set(po.po_id) == set(pay.po_id))
check("all PO suppliers in master", set(po.supplier_id) <= set(sm.supplier_id))
check("all units are MT", (po.unit == "MT").all(), po.unit.value_counts().to_dict().__str__())

m = po.merge(gr, on="po_id", validate="1:1").merge(
    pay[["po_id", "supplier_id", "invoice_date", "agreed_payment_days", "actual_payment_date", "days_late"]]
    .rename(columns={"supplier_id": "supplier_id_pay", "days_late": "arora_days_late"}), on="po_id", validate="1:1")
check("supplier_id agrees PO vs payment", (m.supplier_id == m.supplier_id_pay).all())
check("payment terms agree PO vs payment", (m.payment_terms_days == m.agreed_payment_days).all())
check("receipt_date >= po_date", (m.receipt_date >= m.po_date).all())
check("promised date >= po_date", (m.delivery_promised_date >= m.po_date).all())
check("received <= ordered", (m.quantity_received <= m.quantity_ordered).all())
check("rejection_qty <= received", (m.rejection_qty <= m.quantity_received).all())
check("rejection_qty >=0 and days_late >=0", ((m.rejection_qty >= 0) & (m.arora_days_late >= 0)).all())
bill_ratio = m.invoice_amount_billed / (m.quantity_ordered * m.unit_price_quoted)
check("invoice = ordered qty x quoted price (within 0.01%)", bill_ratio.between(0.9999, 1.0001).all(),
      f"ratio min {bill_ratio.min():.6f} max {bill_ratio.max():.6f}")
check("rejection rows all have a reason", ((m.rejection_qty > 0) == m.rejection_reason.notna()).all(),
      f"{(m.rejection_qty>0).sum()} rows with rejection_qty>0; {m.rejection_reason.notna().sum()} rows with a reason (not identical sets)")
check("paid-after-invoice check: actual_payment_date >= invoice_date", (m.actual_payment_date >= m.invoice_date).all())
check("returns: dates inside PO window", cr.return_date.between(po.po_date.min(), gr.receipt_date.max() + pd.Timedelta(days=1)).all())
check("returns: material exists in POs", cr.material_id.isin(po.material_id).all())
check("returns: traced supplier in master", cr.supplier_id_traced.dropna().isin(sm.supplier_id).all())
miss = cr.supplier_id_traced.isna().mean()
check("returns: missing supplier share (info)", True, f"{cr.supplier_id_traced.isna().sum()} of {len(cr)} = {miss:.1%} -> inference required")

# ---------- 3. Master categories (source of truth) ----------
cat_map = (sm.assign(category=sm.material_categories.str.split(",")).explode("category")
             .assign(category=lambda d: d.category.str.strip())[["supplier_id", "supplier_name", "category"]])
master_pairs = set(zip(cat_map.supplier_id, cat_map.category))
m["in_master"] = [(s, c) in master_pairs for s, c in zip(m.supplier_id, m.material_id)]
check("PO materials inside supplier_master categories (info)", True,
      f"{m.in_master.sum()} in-master POs, {(~m.in_master).sum()} off-master ({(~m.in_master).mean():.1%}) -> tagged, not dropped")

# ---------- 4. Derived PO-level metrics ----------
m["short_qty"]   = m.quantity_ordered - m.quantity_received
m["short_pct"]   = m.short_qty / m.quantity_ordered * 100
m["days_late_delivery"] = (m.receipt_date - m.delivery_promised_date).dt.days
m["is_late"]     = m.days_late_delivery > 0
m["ordered_value"]       = m.quantity_ordered * m.unit_price_quoted
m["received_value"]      = m.quantity_received * m.unit_price_quoted
m["billing_gap_inr"]     = m.invoice_amount_billed - m.received_value          # judged formula
m["rejection_value_inr"] = m.rejection_qty * m.unit_price_quoted               # rejection is inside received qty, and billed
m["rejection_pct"]       = m.rejection_qty / m.quantity_received * 100

# price vs panel: same material, same calendar year, all suppliers (signed; aggregate per supplier later)
m["year"] = m.po_date.dt.year
m["panel_avg_price"] = m.groupby(["material_id", "year"]).unit_price_quoted.transform("mean")
m["price_vs_panel_pct"] = (m.unit_price_quoted / m.panel_avg_price - 1) * 100
m["premium_vs_panel_inr"] = (m.unit_price_quoted - m.panel_avg_price) * m.quantity_received

# market index (6 of 18 materials only)
m["month"] = m.po_date.dt.strftime("%Y-%m")
m = m.merge(mp[["month", "material_category", "market_price_per_mt"]].rename(columns={"material_category": "material_id"}),
            on=["month", "material_id"], how="left")
m["price_vs_index_pct"] = (m.unit_price_quoted / m.market_price_per_mt - 1) * 100
m["premium_vs_index_inr"] = (m.unit_price_quoted - m.market_price_per_mt) * m.quantity_received
check("market index coverage (info)", True,
      f"{m.market_price_per_mt.notna().sum()} of {len(m)} POs ({m.market_price_per_mt.notna().mean():.1%}) matched; materials: {sorted(m.loc[m.market_price_per_mt.notna(),'material_id'].unique())}")
check("market index level vs PO prices (info)", True,
      f"median PO price is {m.price_vs_index_pct.median():+.0f}% vs index on covered POs -> level gap, use panel comparison as primary")

# panel baseline for 'normal' shortage (no use of is_underperformer flag): median supplier short%
sup_short = m.groupby("supplier_id").apply(lambda d: (d.short_qty.sum() / d.quantity_ordered.sum()) * 100, include_groups=False)
baseline = float(sup_short.median())
m["baseline_short_pct"] = baseline
m["baseline_gap_inr"] = m.ordered_value * baseline / 100
m["excess_gap_inr"] = (m.billing_gap_inr - m.baseline_gap_inr)   # signed here; aggregate then floor at 0 per supplier
check("panel baseline short % (median supplier)", True, f"{baseline:.3f}%")

# ---------- 5. Returns: clean copy + price lookup helper ----------
cr["supplier_known"] = cr.supplier_id_traced.notna()

# ---------- 6. Write ----------
cols = ["po_id","po_date","supplier_id","material_id","in_master","quantity_ordered","unit_price_quoted","ordered_value",
        "delivery_promised_date","receipt_date","days_late_delivery","is_late","quantity_received","short_qty","short_pct",
        "invoice_amount_billed","received_value","billing_gap_inr","baseline_gap_inr","excess_gap_inr",
        "quality_grade","rejection_qty","rejection_pct","rejection_reason","rejection_value_inr",
        "panel_avg_price","price_vs_panel_pct","premium_vs_panel_inr","market_price_per_mt","price_vs_index_pct","premium_vs_index_inr",
        "payment_terms_days","arora_days_late","gr_id"]
m[cols].to_csv(f"{OUT}/po_fact_table.csv", index=False)
pd.DataFrame(log).to_csv(f"{OUT}/data_quality_log.csv", index=False)
cat_map.to_csv(f"{OUT}/supplier_category_map.csv", index=False)
cr.to_csv(f"{OUT}/returns_clean.csv", index=False)
print(pd.DataFrame(log).to_string(index=False))
print("\nfact table:", m[cols].shape)
