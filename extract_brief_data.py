import pandas as pd, numpy as np

f = pd.read_csv('output/po_fact_table.csv', parse_dates=['po_date','receipt_date'])
imp = pd.read_csv('output/supplier_rupee_impact_v2.csv', index_col=0)
sm = pd.read_csv('data/supplier_master.csv')
rc = pd.read_csv('output/returns_confirmed_valued.csv')
cat = pd.read_csv('output/supplier_category_impact.csv')

# Confirmed returns per supplier
print('=== CONFIRMED RETURNS ===')
for sid in ['VS12','VS19','VS03']:
    sub = rc[rc.supplier_id_traced==sid][['return_id','return_date','material_id','qty_counted','price_used','value_inr','note']].sort_values('value_inr', ascending=False).head(5)
    print(f'\n{sid} confirmed returns (top 5 by value):')
    print(sub.to_string(index=False))

# Category breakdown
print('\n=== CATEGORY IMPACT for VS12, VS19, VS03 ===')
sub = cat[cat.supplier_id.isin(['VS12','VS19','VS03'])][['supplier_id','in_master','material_id','po_count','short_pct','billing_gap_inr','excess_gap_inr','rejection_loss_inr','avg_days_late']].sort_values(['supplier_id','excess_gap_inr'], ascending=[True,False])
print(sub.to_string(index=False))

# Find good alternative suppliers for VS12, VS19, VS03 categories
print('\n=== ALTERNATIVE SUPPLIERS ===')
for cat_name in ['Spring Steel','HR Coils','MS Flats','MS Pipes','ERW Pipes','CR Sheets','Hex Bars']:
    candidates = sm[sm.material_categories.str.contains(cat_name, na=False)]['supplier_id'].tolist()
    acceptable = [s for s in candidates if s in imp.index and imp.loc[s,'rank_v2'] > 3]
    row = []
    for s in acceptable:
        name = imp.loc[s,'supplier_name']
        loss = imp.loc[s,'HEADLINE_LOSS_v2_inr']/1e5
        billed = imp.loc[s,'billed_inr']/1e7
        row.append(f"{s}({name}, loss={loss:.1f}L, billed={billed:.1f}Cr)")
    print(f'{cat_name}: {row}')

# VS16 and VS17 specifics
print('\n=== VS16 / VS17 watch-list data ===')
for sid in ['VS16','VS17']:
    row = imp.loc[sid]
    print(f"\n{sid} {row.supplier_name}:")
    print(f"  Billed: {row.billed_inr/1e7:.2f} Cr | Gap: {row.A_billing_gap_inr/1e5:.1f}L | Excess: {row.A_excess_gap_inr/1e5:.1f}L")
    print(f"  Confirmed returns: {row.C_confirmed_returns_inr/1e5:.1f}L ({int(row.C_confirmed_return_count)} returns)")
    print(f"  Inferred returns: {row.D_inferred_returns_expected_inr/1e5:.1f}L")
    print(f"  Headline loss v2: {row.HEADLINE_LOSS_v2_inr/1e5:.1f}L")
    print(f"  Price vs panel: {row.price_vs_panel_pct_mean:.2f}% (z={row.price_z:.2f})")
    print(f"  Returns per 1000 MT: {row.returns_per_1000MT_all:.2f}")
    print(f"  Avg days late: {row.avg_days_late:.1f}")
    mcat = sm[sm.supplier_id==sid]['material_categories'].values[0]
    print(f"  Master categories: {mcat}")
