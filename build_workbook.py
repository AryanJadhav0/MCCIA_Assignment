"""Build Supplier_Evidence_Workbook.xlsx: raw inputs in blue, every rupee figure a live formula. Run: python build_workbook.py <output_dir> <raw_dir>"""
import sys, pandas as pd, numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
D   = sys.argv[1] if len(sys.argv) > 1 else "/mnt/user-data/outputs"
RAW = sys.argv[2] if len(sys.argv) > 2 else "/mnt/user-data/uploads"
f   = pd.read_csv(f"{D}/po_fact_table.csv", parse_dates=["po_date", "receipt_date", "delivery_promised_date"])
imp = pd.read_csv(f"{D}/supplier_rupee_impact_v2.csv", index_col=0).sort_index()
sc  = pd.read_csv(f"{D}/supplier_scorecard.csv", index_col=0)
rc  = pd.read_csv(f"{D}/returns_confirmed_valued.csv", parse_dates=["return_date"])
ra  = pd.read_csv(f"{D}/returns_attributed.csv"); ri = pd.read_csv(f"{D}/returns_inferred_valued.csv")
cat = pd.read_csv(f"{D}/supplier_category_impact.csv"); dq = pd.read_csv(f"{D}/data_quality_log.csv")
val = pd.read_csv(f"{D}/attribution_validation.csv"); sens = pd.read_csv(f"{D}/scorecard_sensitivity.csv")
baseline = float(f.baseline_gap_inr.iloc[0] / f.ordered_value.iloc[0] * 100)

AR = "Arial"; BLUE = Font(name=AR, size=10, color="0000FF"); BLK = Font(name=AR, size=10); GRN = Font(name=AR, size=10, color="008000")
HDR = Font(name=AR, size=10, bold=True, color="FFFFFF"); HF = PatternFill("solid", fgColor="1F3864"); YEL = PatternFill("solid", fgColor="FFFF00")
BOLD = Font(name=AR, size=10, bold=True); TITLE = Font(name=AR, size=14, bold=True, color="1F3864")
RUP = '#,##0;(#,##0);-'; PCT = '0.00'; thin = Side(style="thin", color="BBBBBB")
wb = Workbook()

def header(ws, row, cols, widths=None):
    for j, c in enumerate(cols, 1):
        x = ws.cell(row, j, c); x.font = HDR; x.fill = HF; x.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[L(j)].width = (widths[j - 1] if widths else max(12, min(28, len(str(c)) + 3)))
    ws.row_dimensions[row].height = 32
def put(ws, r, c, v, font=BLK, fmt=None):
    x = ws.cell(r, c, v); x.font = font
    if fmt: x.number_format = fmt
    return x

# ============================================================ Assumptions
wa = wb.active; wa.title = "Assumptions"
put(wa, 1, 1, "Assumptions and inputs", TITLE)
put(wa, 3, 1, "Panel baseline short-delivery % (median supplier)", BOLD); put(wa, 3, 2, baseline, BLUE, '0.0000').fill = YEL
put(wa, 3, 3, "Median of the 34 suppliers' qty-weighted short %. Computed from data, never from the underperformer flag. Used to separate normal tolerance from excess shortage.")
put(wa, 4, 1, "Late-delivery carrying cost, % p.a. (memo only)", BOLD); put(wa, 4, 2, 0.12, BLUE, '0.0%').fill = YEL
put(wa, 4, 3, "Assumption (not from the data). Late-delivery cost proxy is NOT part of the headline loss.")
put(wa, 6, 1, "Scorecard weights (sum to 100%)", BOLD)
W = [("Short delivery %", .35), ("Late delivery (avg days)", .25), ("Rejection %", .15), ("Returns per 1,000 MT", .15), ("Price vs panel (scaled by z-score)", .10)]
for i, (n, w) in enumerate(W):
    put(wa, 7 + i, 1, n); put(wa, 7 + i, 2, w, BLUE, '0%').fill = YEL
put(wa, 12, 1, "Total", BOLD); put(wa, 12, 2, "=SUM(B7:B11)", BLK, '0%')
put(wa, 13, 1, "Price z-score at which price score reaches 0"); put(wa, 13, 2, 3, BLUE).fill = YEL
put(wa, 14, 1, "Tier cut-offs: POOR below / WATCH below"); put(wa, 14, 2, 60, BLUE).fill = YEL; put(wa, 14, 3, 85, BLUE).fill = YEL
notes = ["Colour code: blue = input taken from data or model output; black = formula; green = cross-sheet link. Yellow = assumption you may change.",
         "Not used anywhere: years_of_relationship, is_underperformer. Category source of truth: supplier_master.",
         "Inferred-return probabilities come from the Phase 2 model (see Validation). Per-return values are probability x quantity x batch price.",
         "Dataset scale: total billed is ~Rs 1,009 Cr over 3 years, ~37x the Rs 70-85L/month in the problem statement. Reported as found.",
         "Arora's own payment lateness (avg ~7.3 days) is in PO_Facts column M and Supplier_Impact: expect suppliers to raise it."]
for i, t in enumerate(notes): put(wa, 17 + i, 1, t)
wa.column_dimensions["A"].width = 52; wa.column_dimensions["B"].width = 12; wa.column_dimensions["C"].width = 90

# ============================================================ PO_Facts
wp = wb.create_sheet("PO_Facts")
cols = ["po_id", "po_date", "supplier_id", "material_id", "in_master", "qty_ordered_MT", "unit_price_quoted", "delivery_promised", "receipt_date", "qty_received_MT",
        "invoice_billed", "rejection_qty_MT", "arora_days_late_paying", "short_qty_MT", "days_late_delivery", "ordered_value", "received_value", "BILLING GAP (billed - received x price)",
        "baseline_gap", "excess_gap", "rejection_loss"]
header(wp, 1, cols, [14, 11, 8, 18, 8, 10, 12, 12, 12, 10, 14, 10, 10, 10, 9, 14, 14, 16, 14, 14, 13])
f = f.sort_values("po_id").reset_index(drop=True)
for i, r in enumerate(f.itertuples(), 2):
    vals = [r.po_id, r.po_date.to_pydatetime(), r.supplier_id, r.material_id, "Y" if r.in_master else "N", r.quantity_ordered, r.unit_price_quoted,
            r.delivery_promised_date.to_pydatetime(), r.receipt_date.to_pydatetime(), r.quantity_received, r.invoice_amount_billed, r.rejection_qty, r.arora_days_late]
    for j, v in enumerate(vals, 1): put(wp, i, j, v, BLUE, 'yyyy-mm-dd' if j in (2, 8, 9) else (RUP if j in (7, 11) else None))
    put(wp, i, 14, f"=F{i}-J{i}", BLK, '0.00'); put(wp, i, 15, f"=I{i}-H{i}", BLK, '0')
    put(wp, i, 16, f"=F{i}*G{i}", BLK, RUP); put(wp, i, 17, f"=J{i}*G{i}", BLK, RUP); put(wp, i, 18, f"=K{i}-Q{i}", BLK, RUP)
    put(wp, i, 19, f"=P{i}*Assumptions!$B$3/100", GRN, RUP); put(wp, i, 20, f"=R{i}-S{i}", BLK, RUP); put(wp, i, 21, f"=L{i}*G{i}", BLK, RUP)
wp.freeze_panes = "B2"; wp.auto_filter.ref = f"A1:U{len(f)+1}"; NP = len(f) + 1

# ============================================================ Returns_Confirmed
wc = wb.create_sheet("Returns_Confirmed")
header(wc, 1, ["return_id", "return_date", "client_id", "material_id", "qty_returned_MT", "reason", "supplier (as traced)", "matched_po (batch)", "qty_counted_MT (capped)",
               "batch_price", "VALUE (qty counted x price)", "batch received on/before return?"], [11, 11, 11, 18, 10, 18, 10, 14, 12, 12, 16, 22])
rc = rc.sort_values("return_id").reset_index(drop=True)
for i, r in enumerate(rc.itertuples(), 2):
    for j, v in enumerate([r.return_id, r.return_date.to_pydatetime(), r.client_id, r.material_id, r.quantity_returned, r.reason, r.supplier_id_traced, r.matched_po, r.qty_counted, r.price_used], 1):
        put(wc, i, j, v, BLUE, 'yyyy-mm-dd' if j == 2 else (RUP if j == 10 else None))
    put(wc, i, 11, f"=I{i}*J{i}", BLK, RUP)
    put(wc, i, 12, f'=IF(INDEX(PO_Facts!$I$2:$I${NP},MATCH(H{i},PO_Facts!$A$2:$A${NP},0))<=B{i},"yes","FALLBACK: no earlier batch")', GRN)
NC = len(rc) + 1; wc.freeze_panes = "A2"
put(wc, NC + 2, 1, "Cap rule: a batch cannot return more than (received - rejected) MT in total, so a return is never counted twice against rejection. FALLBACK rows price the return from the supplier's earliest batch of that material because no earlier batch exists.")

# ============================================================ Returns_Inferred (long format: one row per candidate supplier)
wi = wb.create_sheet("Returns_Inferred")
header(wi, 1, ["return_id", "candidate supplier", "probability", "matched_po (batch)", "qty_if_true_MT", "batch_price", "value if true", "EXPECTED VALUE (prob x value)"], [11, 12, 12, 16, 12, 12, 14, 18])
ri = ri.sort_values(["return_id", "supplier_id"]).reset_index(drop=True)
for i, r in enumerate(ri.itertuples(), 2):
    for j, v in enumerate([r.return_id, r.supplier_id, r.prob, r.matched_po, r.qty_if_true, r.price], 1): put(wi, i, j, v, BLUE, '0.0000' if j == 3 else (RUP if j == 6 else None))
    put(wi, i, 7, f"=E{i}*F{i}", BLK, RUP); put(wi, i, 8, f"=C{i}*G{i}", BLK, RUP)
NI = len(ri) + 1; wi.freeze_panes = "A2"; wi.auto_filter.ref = f"A1:H{NI}"

# ============================================================ Supplier_Impact
ws = wb.create_sheet("Supplier_Impact", 1)
put(ws, 1, 1, "Rupee impact per supplier, 3 years (all figures are live formulas over PO_Facts and the returns sheets)", TITLE)
heads = ["supplier_id", "supplier_name", "PO count", "Billed", "Received x price", "BILLING GAP (A)", "Panel-normal part of gap", "EXCESS GAP (A*)", "REJECTION LOSS (B)",
         "CONFIRMED RETURNS (C)", "INFERRED RETURNS, expected (D)", "HEADLINE LOSS = A*+B+C+D", "Rank", "Inferred range p10", "Inferred range p90", "Arora avg days late paying"]
header(ws, 3, heads, [10, 22, 8, 16, 16, 14, 14, 14, 14, 14, 16, 18, 6, 13, 13, 12])
ids = list(imp.index); r0 = 4; r1 = r0 + len(ids) - 1
for k, s in enumerate(ids):
    r = r0 + k; rng = lambda c: f"PO_Facts!${c}$2:${c}${NP}"
    put(ws, r, 1, s, BLK); put(ws, r, 2, imp.supplier_name[s], BLUE)
    put(ws, r, 3, f"=COUNTIFS({rng('C')},A{r})", GRN, '0')
    put(ws, r, 4, f"=SUMIFS({rng('K')},{rng('C')},A{r})", GRN, RUP); put(ws, r, 5, f"=SUMIFS({rng('Q')},{rng('C')},A{r})", GRN, RUP)
    put(ws, r, 6, f"=D{r}-E{r}", BLK, RUP); put(ws, r, 7, f"=SUMIFS({rng('S')},{rng('C')},A{r})", GRN, RUP)
    put(ws, r, 8, f"=MAX(0,SUMIFS({rng('T')},{rng('C')},A{r}))", GRN, RUP); put(ws, r, 9, f"=SUMIFS({rng('U')},{rng('C')},A{r})", GRN, RUP)
    put(ws, r, 10, f"=SUMIFS(Returns_Confirmed!$K$2:$K${NC},Returns_Confirmed!$G$2:$G${NC},A{r})", GRN, RUP)
    put(ws, r, 11, f"=SUMIFS(Returns_Inferred!$H$2:$H${NI},Returns_Inferred!$B$2:$B${NI},A{r})", GRN, RUP)
    put(ws, r, 12, f"=H{r}+I{r}+J{r}+K{r}", BOLD, RUP); put(ws, r, 13, f"=RANK(L{r},$L${r0}:$L${r1})", BLK, '0')
    put(ws, r, 14, float(imp.D_inferred_returns_p10_inr[s]), BLUE, RUP); put(ws, r, 15, float(imp.D_inferred_returns_p90_inr[s]), BLUE, RUP)
    put(ws, r, 16, f"=AVERAGEIFS({rng('M')},{rng('C')},A{r})", GRN, '0.0')
rt = r1 + 1; put(ws, rt, 2, "PANEL TOTAL", BOLD)
for c in range(3, 13): put(ws, rt, c, f"=SUM({L(c)}{r0}:{L(c)}{r1})", BOLD, RUP if c > 3 else '0')
put(ws, rt + 2, 1, "A = invoice_amount_billed - quantity_received x unit_price_quoted.  A* = A minus the panel-normal part (floored at 0 per supplier).  B = rejection_qty x unit price (rejected goods are inside received qty and were billed).", BLK)
put(ws, rt + 3, 1, "C = confirmed returns, capped per batch.  D = probability-weighted expected value of the 140 unattributed returns (80% range in columns N-O, from 4,000 simulated assignments).", BLK)
put(ws, rt + 4, 1, "Price premium and late-delivery carrying cost are NOT in the headline: no supplier shows a statistically significant price premium, and the late cost is an assumption.", BLK)
ws.freeze_panes = "C4"

# ============================================================ Scorecard
wsc = wb.create_sheet("Scorecard", 2)
put(wsc, 1, 1, "34-supplier scorecard - performance only (no relationship years). Score 100 = panel median or better, 0 = as bad as the worst supplier on that metric", TITLE)
heads = ["supplier_id", "supplier_name", "PO count", "Short delivery % of ordered qty", "Late POs %", "Avg days late", "Rejection % of received", "Returns per 1,000 MT received",
         "Price vs panel z-score", "Score: short", "Score: late", "Score: rejection", "Score: returns", "Score: price", "COMPOSITE", "Tier", "Rank (1 = worst)", "MT received"]
header(wsc, 3, heads, [10, 22, 8, 13, 9, 9, 12, 14, 11, 9, 9, 10, 9, 9, 11, 12, 9, 10])
for k, s in enumerate(ids):
    r = r0 + k; rng = lambda c: f"PO_Facts!${c}$2:${c}${NP}"
    put(wsc, r, 1, s); put(wsc, r, 2, imp.supplier_name[s], BLUE); put(wsc, r, 3, f"=COUNTIFS({rng('C')},A{r})", GRN, '0')
    put(wsc, r, 18, f"=SUMIFS({rng('J')},{rng('C')},A{r})", GRN, '#,##0.0')
    put(wsc, r, 4, f"=SUMIFS({rng('N')},{rng('C')},A{r})/SUMIFS({rng('F')},{rng('C')},A{r})*100", GRN, PCT)
    put(wsc, r, 5, f'=COUNTIFS({rng("C")},A{r},{rng("O")},">0")/C{r}*100', GRN, '0.0')
    put(wsc, r, 6, f"=AVERAGEIFS({rng('O')},{rng('C')},A{r})", GRN, PCT)
    put(wsc, r, 7, f"=SUMIFS({rng('L')},{rng('C')},A{r})/R{r}*100", GRN, PCT)
    put(wsc, r, 8, f"=(COUNTIFS(Returns_Confirmed!$G$2:$G${NC},A{r})+SUMIFS(Returns_Inferred!$C$2:$C${NI},Returns_Inferred!$B$2:$B${NI},A{r}))/R{r}*1000", GRN, PCT)
    put(wsc, r, 9, float(imp.price_z[s]), BLUE, '0.00')
    for col, src in ((10, "D"), (11, "F"), (12, "G"), (13, "H")):
        rg = f"${src}${r0}:${src}${r1}"
        put(wsc, r, col, f"=100*(1-MIN(1,MAX(0,({src}{r}-MEDIAN({rg}))/(MAX({rg})-MEDIAN({rg})))))", BLK, '0.0')
    put(wsc, r, 14, f"=100*(1-MIN(1,MAX(I{r},0)/Assumptions!$B$13))", GRN, '0.0')
    put(wsc, r, 15, f"=J{r}*Assumptions!$B$7+K{r}*Assumptions!$B$8+L{r}*Assumptions!$B$9+M{r}*Assumptions!$B$10+N{r}*Assumptions!$B$11", BOLD, '0.0')
    put(wsc, r, 16, f'=IF(O{r}<Assumptions!$B$14,"POOR",IF(O{r}<Assumptions!$C$14,"WATCH","ACCEPTABLE"))', GRN)
    put(wsc, r, 17, f"=RANK(O{r},$O${r0}:$O${r1},1)", BLK, '0')
put(wsc, r1 + 1, 2, "Panel median", BOLD)
for c in (4, 5, 6, 7, 8): put(wsc, r1 + 1, c, f"=MEDIAN({L(c)}{r0}:{L(c)}{r1})", BOLD, PCT)
put(wsc, r1 + 3, 1, "Price z-score = mean % above/below the same material's panel average that year, divided by its standard error (from the pipeline). Price is scored by statistical evidence, so noise is not penalised.")
wsc.freeze_panes = "C4"

# ============================================================ static sheets
def dump(name, df, widths=None, fmt=None):
    w = wb.create_sheet(name); header(w, 1, list(df.columns), widths)
    for i, row in enumerate(df.itertuples(index=False), 2):
        for j, v in enumerate(row, 1):
            if isinstance(v, (np.floating, float)) and np.isnan(v): v = None
            if isinstance(v, np.generic): v = v.item()
            put(w, i, j, v, BLUE)
    w.freeze_panes = "A2"; return w
dump("Returns_Attributed", ra, [11, 11, 11, 18, 10, 17, 9, 9, 9, 9, 9, 9, 11, 12, 22, 12])
dump("Category_Scorecard", cat.round(3), None)
v = dump("Validation", val, [58, 11, 11, 11])
r = len(val) + 3
for t in ["Attribution validated with 5 repeats of 5-fold cross-validation on the 240 known returns; rate/likelihoods re-learned inside each training fold.",
          "Date proximity test: true supplier's latest batch median 84 days before return vs 79 days for other suppliers (Mann-Whitney p = 0.83): no information, so excluded.",
          "Exact-supplier accuracy is modest (top-1 22%) because the three high-return suppliers are near-identical on available features; aggregate return counts per supplier are recovered closely (e.g. VS12 ~51 expected vs 54 actual).",
          "Rupee totals use summed probabilities, which is why they are far more reliable than any single assignment. See the p10-p90 range on Supplier_Impact."]:
    put(v, r, 1, t); r += 1
dump("Sensitivity", sens, [40, 10, 24, 10, 8, 14])
dump("Data_Quality_Log", dq, [55, 8, 110])

# ============================================================ Summary (first sheet)
sm_ = wb.create_sheet("Summary", 0)
put(sm_, 1, 1, "Arora Traders - Supplier Evidence Workbook (PS 4)", TITLE)
rows = [("Total billed to suppliers, 3 yrs (Rs)", f"=Supplier_Impact!D{rt}"), ("Total billing gap: billed minus received x price (Rs)", f"=Supplier_Impact!F{rt}"),
        ("  of which excess over panel-normal shortage (Rs)", f"=Supplier_Impact!H{rt}"),
        ("VS12 Nutan Tubes - headline loss (Rs)", f'=INDEX(Supplier_Impact!L{r0}:L{r1},MATCH("VS12",Supplier_Impact!A{r0}:A{r1},0))'),
        ("VS19 Thakur Steel - headline loss (Rs)", f'=INDEX(Supplier_Impact!L{r0}:L{r1},MATCH("VS19",Supplier_Impact!A{r0}:A{r1},0))'),
        ("VS03 Tata Steel Service - headline loss (Rs)", f'=INDEX(Supplier_Impact!L{r0}:L{r1},MATCH("VS03",Supplier_Impact!A{r0}:A{r1},0))'),
        ("Suppliers rated POOR on the scorecard", f'=COUNTIF(Scorecard!P{r0}:P{r1},"POOR")'),
        ("Lowest composite score among the other suppliers", f'=SMALL(Scorecard!O{r0}:O{r1},4)')]
for i, (a, b) in enumerate(rows):
    put(sm_, 3 + i, 1, a, BOLD if not a.startswith("  ") else BLK); put(sm_, 3 + i, 2, b, GRN, RUP if i < 6 else '0.0')
put(sm_, 12, 1, "Sheets: Assumptions | Supplier_Impact (rupees) | Scorecard | PO_Facts (every PO, formulas visible) | Returns_Confirmed | Returns_Inferred | Returns_Attributed | Category_Scorecard | Validation | Sensitivity | Data_Quality_Log")
put(sm_, 13, 1, "Worked example for VS12: filter PO_Facts on supplier_id = VS12 and sum column R (billed - received x price); it equals Supplier_Impact column F.")
sm_.column_dimensions["A"].width = 62; sm_.column_dimensions["B"].width = 20
for w in wb.worksheets: w.sheet_view.showGridLines = True
wb.save(f"{D}/Supplier_Evidence_Workbook.xlsx"); print("saved", NP, NC, NI)
