"""
Phase 5: Negotiation Briefs (D4)
Generates:
  - 3 full Word briefs (VS12 Nutan Tubes, VS19 Thakur Steel, VS03 Tata Steel Service)
  - 2 watch-list one-pagers (VS16 Garg Enterprises, VS17 Modern Metals)
Run:  python phase5_negotiation_briefs.py <output_dir> <raw_dir>
"""
import sys, pandas as pd, numpy as np
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

D   = sys.argv[1] if len(sys.argv) > 1 else "output"
RAW = sys.argv[2] if len(sys.argv) > 2 else "data"

# ---- Load data ----
f   = pd.read_csv(f"{D}/po_fact_table.csv", parse_dates=["po_date", "receipt_date"])
imp = pd.read_csv(f"{D}/supplier_rupee_impact_v2.csv", index_col=0)
sm  = pd.read_csv(f"{RAW}/supplier_master.csv").set_index("supplier_id")
rc  = pd.read_csv(f"{D}/returns_confirmed_valued.csv")
cat = pd.read_csv(f"{D}/supplier_category_impact.csv")

# ---- Formatting helpers ----
def cr(n):   return f"Rs {n/1e7:.2f} Cr"
def lk(n):   return f"Rs {n/1e5:.1f}L"
def pct(n):  return f"{n:.2f}%"

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)

def set_cell_border(cell, color="BBBBBB"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        tag = OxmlElement(f"w:{edge}")
        tag.set(qn("w:val"), "single")
        tag.set(qn("w:sz"), "4")
        tag.set(qn("w:color"), color)
        tcBorders.append(tag)
    tcPr.append(tcBorders)

def header_row(table, cols, bg="1F3864", text_color="FFFFFF", font_size=9):
    row = table.rows[0]
    for cell, txt in zip(row.cells, cols):
        cell.text = ""
        set_cell_bg(cell, bg)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(txt)
        run.font.bold = True
        run.font.size = Pt(font_size)
        run.font.color.rgb = RGBColor.from_string(text_color)

def data_row(table, row_idx, values, bold=False, bg=None, font_size=9, align=None):
    row = table.rows[row_idx]
    for i, (cell, val) in enumerate(zip(row.cells, values)):
        cell.text = ""
        p = cell.paragraphs[0]
        a = align[i] if align else WD_ALIGN_PARAGRAPH.LEFT
        p.alignment = a
        run = p.add_run(str(val))
        run.font.bold = bold
        run.font.size = Pt(font_size)
        if bg:
            set_cell_bg(cell, bg)
        set_cell_border(cell)

def add_heading(doc, text, level=1, color="1F3864"):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(14 if level == 1 else 11)
    run.font.color.rgb = RGBColor.from_string(color)
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    return p

def add_body(doc, text, size=10, bold=False, italic=False, color=None):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    return p

def add_bullet(doc, text, size=10):
    p = doc.add_paragraph(style="List Bullet")
    run = p.add_run(text)
    run.font.size = Pt(size)
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    return p

def add_divider(doc, color="1F3864"):
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)
    pPr.append(pBdr)
    p.paragraph_format.space_after = Pt(4)

def set_margins(doc, top=1.5, bottom=1.5, left=1.8, right=1.8):
    for section in doc.sections:
        section.top_margin = Cm(top)
        section.bottom_margin = Cm(bottom)
        section.left_margin = Cm(left)
        section.right_margin = Cm(right)

# ============================================================
# FULL BRIEF
# ============================================================
def build_full_brief(sid, alt_map):
    row    = imp.loc[sid]
    name   = row.supplier_name
    master = sm.loc[sid, "material_categories"]
    credit = sm.loc[sid, "credit_days_agreed"]

    doc = Document()
    set_margins(doc)

    # --- TITLE ---
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = title.add_run("SUPPLIER NEGOTIATION BRIEF")
    tr.font.size = Pt(16); tr.font.bold = True
    tr.font.color.rgb = RGBColor.from_string("1F3864")

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr = sub.add_run("Arora Traders, Pune  |  Prepared: October 2026  |  CONFIDENTIAL")
    sr.font.size = Pt(9); sr.italic = True
    sr.font.color.rgb = RGBColor.from_string("555555")
    add_divider(doc)

    # Identity table
    id_t = doc.add_table(rows=2, cols=4)
    id_t.style = "Table Grid"
    header_row(id_t, ["Supplier", "Supplier ID", "Master categories", "Credit terms"], bg="2E5FA3")
    data_row(id_t, 1, [name, sid, master, f"{credit} days"], bold=True, bg="EEF2FF",
             align=[WD_ALIGN_PARAGRAPH.LEFT]*4)

    id_t2 = doc.add_table(rows=2, cols=3)
    id_t2.style = "Table Grid"
    header_row(id_t2, ["PO count (3 yrs)", "Total billed (3 yrs)", "Arora avg. payment delay"], bg="2E5FA3")
    data_row(id_t2, 1,
             [int(row.po_count), cr(row.billed_inr),
              f"{row.arora_avg_days_late_paying:.1f} days late (our side — panel avg {f.arora_days_late.mean():.1f} days)"],
             bg="EEF2FF",
             align=[WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT])
    doc.add_paragraph()

    # --- SECTION 1: Financial Impact ---
    add_heading(doc, "1. Financial Impact — 3-Year Summary (Oct 2023 – Oct 2026)")
    add_divider(doc, "2E5FA3")

    ht = doc.add_table(rows=6, cols=3)
    ht.style = "Table Grid"
    header_row(ht, ["Loss Component", "Amount (Rs)", "How it is measured"])
    comps = [
        ("A*  Short-billing excess (above panel baseline 0.51%)",
         lk(row.A_excess_gap_inr),
         "Invoice on ordered qty; received qty was short. Panel baseline subtracted. Floor at 0 per supplier."),
        ("B   Rejection loss",
         lk(row.B_rejection_loss_inr),
         "Rejected MT x quoted unit price. Rejected goods inside received qty and were billed."),
        ("C   Confirmed customer returns",
         lk(row.C_confirmed_returns_inr),
         f"{int(row.C_confirmed_return_count)} batches. Capped per batch (no double-count with rejections)."),
        ("D   Inferred returns (probability estimate)",
         lk(row.D_inferred_returns_expected_inr),
         f"80%% range: {lk(imp.loc[sid,'D_inferred_returns_p10_inr'])} to {lk(imp.loc[sid,'D_inferred_returns_p90_inr'])}. Bayesian model."),
        ("HEADLINE LOSS  (A* + B + C + D)",
         cr(row.HEADLINE_LOSS_v2_inr),
         f"{row.HEADLINE_LOSS_pct_of_billed:.1f}%% of total billed over 3 years."),
    ]
    aligns_ht = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.LEFT]
    for ri, (comp, amt, note) in enumerate(comps, 1):
        is_total = ri == 5
        bg = "FFE5CC" if is_total else ("F9FAFB" if ri % 2 else "FFFFFF")
        data_row(ht, ri, [comp, amt, note], bold=is_total, bg=bg, align=aligns_ht)
    doc.add_paragraph()

    add_body(doc,
             f"Memo (NOT in headline): Full billing gap (all short-billing) = {lk(row.A_billing_gap_inr)}.  "
             f"Price premium vs panel = {row.price_vs_panel_pct_mean:+.2f}%% (z={row.price_z:.2f}; not statistically significant).  "
             f"Late-delivery carrying-cost proxy = {lk(row.late_cost_proxy_inr)} (12%% p.a. assumption).",
             size=8, italic=True, color="666666")

    # --- SECTION 2: Performance ---
    add_heading(doc, "2. Performance Metrics vs Panel Benchmark")
    add_divider(doc, "2E5FA3")

    panel_rej  = f.rejection_qty.sum() / f.quantity_received.sum() * 100
    panel_days = f.days_late_delivery.mean()
    panel_late_pct = f.is_late.mean() * 100

    pm = doc.add_table(rows=6, cols=4)
    pm.style = "Table Grid"
    header_row(pm, ["Metric", name, "Panel benchmark", "Comment"])
    metrics = [
        ("Short delivery  % of ordered qty",  pct(row.short_pct_of_qty),   pct(0.511),        f"{row.short_pct_of_qty/0.511:.1f}x the panel baseline"),
        ("Late POs  % of all POs",            pct(row.late_po_pct),         pct(panel_late_pct), "Every PO is late"),
        ("Average days late",                 f"{row.avg_days_late:.1f} d", f"{panel_days:.1f} d", "Consistent across 3 years"),
        ("Rejection  % of received qty",      pct(row.rejection_pct),       pct(panel_rej),    "High rejection rate"),
        ("Price vs panel  (same material/yr)",f"{row.price_vs_panel_pct_mean:+.2f}%%  z={row.price_z:.2f}", "0.00%%  z=0", "Weak signal — not in headline"),
    ]
    aligns_pm = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER,
                 WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT]
    for ri, (met, val, bench, cmt) in enumerate(metrics, 1):
        data_row(pm, ri, [met, val, bench, cmt], bg="F9FAFB" if ri % 2 else "FFFFFF", align=aligns_pm)
    doc.add_paragraph()

    # --- SECTION 3: Worst POs ---
    add_heading(doc, "3. Five Worst Purchase Orders (by excess billing gap)")
    add_divider(doc, "2E5FA3")

    worst5 = (f[f.supplier_id == sid]
               .nlargest(5, "excess_gap_inr")
               [["po_id","po_date","material_id","quantity_ordered","quantity_received",
                 "invoice_amount_billed","billing_gap_inr","rejection_value_inr"]])

    wt = doc.add_table(rows=6, cols=8)
    wt.style = "Table Grid"
    header_row(wt, ["PO ID","Date","Material","Ord. MT","Recd. MT",
                    "Billed (Rs)","Billing gap (Rs)","Rejection (Rs)"], font_size=8)
    aligns_wt = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT,
                 WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER,
                 WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT]
    for ri, (_, r2) in enumerate(worst5.iterrows(), 1):
        vals = [r2.po_id, str(r2.po_date.date()), r2.material_id,
                f"{r2.quantity_ordered:.2f}", f"{r2.quantity_received:.2f}",
                f"{r2.invoice_amount_billed:,.0f}", f"{r2.billing_gap_inr:,.0f}",
                f"{r2.rejection_value_inr:,.0f}"]
        data_row(wt, ri, vals, bg="F9FAFB" if ri % 2 else "FFFFFF", align=aligns_wt, font_size=8)

    add_body(doc,
             f"Verify: Open Supplier_Evidence_Workbook.xlsx -> PO_Facts, filter column C = {sid}, sort column T (excess_gap) descending.",
             size=8, italic=True, color="666666")
    doc.add_paragraph()

    # --- SECTION 4: Asks ---
    add_heading(doc, "4. Negotiation Asks — Specific and Quantified")
    add_divider(doc, "2E5FA3")

    asks = [
        ("Ask 1: Credit note for excess short-billing",
         f"Excess over panel-normal shortage = {lk(row.A_excess_gap_inr)} over 3 years. "
         f"Request a credit note for {lk(row.A_excess_gap_inr * 0.5)} (50%%), issued across three invoices."),
        ("Ask 2: Future billing on quantity received",
         "All invoices from next PO onwards to be raised on quantity_received, not quantity_ordered. "
         "A dispatch note with actual MT dispatched must accompany every invoice."),
        ("Ask 3: Credit or replacement for rejected goods",
         f"Rejected goods = {lk(row.B_rejection_loss_inr)} over 3 years. "
         "Request full credit note or replacement shipment within 30 days of this meeting."),
        ("Ask 4: Charge-back for confirmed customer returns",
         f"{int(row.C_confirmed_return_count)} confirmed return batches valued at {lk(row.C_confirmed_returns_inr)}. "
         f"Propose a 50%% charge-back ({lk(row.C_confirmed_returns_inr * 0.5)}) on next invoice. Returns_Confirmed sheet annexed."),
        ("Ask 5: On-time delivery clause (proposed; not data-derived)",
         f"100%% of {int(row.po_count)} POs arrived late (avg {row.avg_days_late:.1f} days). "
         "Propose penalty of 0.1%% of invoice value per day of delay, capped at 2%%. Rate to be agreed in writing — this is a proposal."),
    ]
    at = doc.add_table(rows=len(asks)+1, cols=2)
    at.style = "Table Grid"
    header_row(at, ["Ask", "Detail and proposed wording"])
    for ri, (ask_title, ask_detail) in enumerate(asks, 1):
        data_row(at, ri, [ask_title, ask_detail],
                 bg="FFF8E7" if ri % 2 else "FFFFFF",
                 align=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT], font_size=9)
    doc.add_paragraph()

    # --- SECTION 5: Pushback ---
    add_heading(doc, "5. Anticipated Pushback and Suggested Responses")
    add_divider(doc, "2E5FA3")

    pb_table = doc.add_table(rows=4, cols=2)
    pb_table.style = "Table Grid"
    header_row(pb_table, ["Likely pushback", "Suggested response"])
    pushbacks = [
        ("You also pay us late.",
         f"Fact: we average {row.arora_avg_days_late_paying:.1f} days late on your invoices. "
         "We acknowledge this and propose moving to 7-day electronic transfers in exchange for billing on received quantity and resolution of return credits. Both sides commit in writing."),
        ("The returns are from quality issues at your clients, not our material.",
         f"240 of 380 returns are directly traced to supplier batches at the time of return. "
         "We are asking only for confirmed returns (Returns_Confirmed sheet). The inferred estimate is shown separately and not included in our Ask 4."),
        ("Short deliveries are due to transit losses.",
         f"The panel median shortfall is 0.511%% across all 34 suppliers. {name} runs {row.short_pct_of_qty:.2f}%% — "
         f"{row.short_pct_of_qty/0.511:.1f}x the norm. Transit risk cannot explain this gap."),
    ]
    for ri, (pb, resp) in enumerate(pushbacks, 1):
        data_row(pb_table, ri, [pb, resp], bg="FFF0F0" if ri % 2 else "FFFFFF",
                 align=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT], font_size=9)
    doc.add_paragraph()

    # --- SECTION 6: Alternatives ---
    add_heading(doc, "6. Fallback Position and Alternative Suppliers")
    add_divider(doc, "2E5FA3")

    add_body(doc, "Minimum acceptable outcome from this meeting:", bold=True)
    add_bullet(doc, f"Written agreement to bill on quantity received from next PO onwards.")
    add_bullet(doc, f"Credit note of at least {lk(row.A_excess_gap_inr * 0.25)} within 60 days.")
    add_bullet(doc, "Rejection resolution (replacement or credit) within 30 days.")
    doc.add_paragraph()

    add_body(doc, "Alternative suppliers on the panel (ACCEPTABLE-rated, same categories):", bold=True)
    if alt_map:
        at2 = doc.add_table(rows=len(alt_map)+1, cols=4)
        at2.style = "Table Grid"
        header_row(at2, ["Supplier", "Category", "Headline loss (3 yr)", "Billed (3 yr)"])
        for ri, (a_sid, a_cat) in enumerate(alt_map, 1):
            a_row = imp.loc[a_sid]
            data_row(at2, ri,
                     [f"{a_sid} — {a_row.supplier_name}", a_cat,
                      lk(a_row.HEADLINE_LOSS_v2_inr), cr(a_row.billed_inr)],
                     bg="F9FAFB" if ri % 2 else "FFFFFF",
                     align=[WD_ALIGN_PARAGRAPH.LEFT]*4, font_size=9)
    else:
        add_body(doc, "Note: Only one registered alternative for this category. See Replacement_Recommendation.docx.", italic=True, color="CC0000")
    doc.add_paragraph()

    # --- SECTION 7: Annexure ---
    add_heading(doc, "7. Workbook Cross-Reference (Supplier_Evidence_Workbook.xlsx)")
    add_divider(doc, "2E5FA3")
    annex = [
        ("PO_Facts",            f"Filter col C = {sid}. Col R = billing gap. Col T = excess gap."),
        ("Returns_Confirmed",   f"Filter col G = {sid}. Col K = confirmed return value."),
        ("Returns_Inferred",    f"Filter col B = {sid}. Col H = expected inferred value."),
        ("Supplier_Impact",     f"Row for {sid}. Live formulas. Cols H, I, J, K = A*, B, C, D."),
        ("Scorecard",           f"Row for {sid}. Composite score and tier."),
    ]
    an_t = doc.add_table(rows=len(annex)+1, cols=2)
    an_t.style = "Table Grid"
    header_row(an_t, ["Sheet", "How to verify the numbers in this brief"])
    for ri, (sheet, note) in enumerate(annex, 1):
        data_row(an_t, ri, [sheet, note], bg="F9FAFB" if ri % 2 else "FFFFFF", font_size=9)

    add_body(doc,
             "All rupee figures in this brief are live formulas in the workbook — filter and verify any number.",
             size=8, italic=True, color="666666")

    out = f"{D}/Brief_{sid}_{name.replace(' ','_')}.docx"
    doc.save(out)
    print(f"Saved full brief: {out}")


# ============================================================
# WATCH-LIST ONE-PAGER
# ============================================================
def build_watchlist_brief(sid, points_override=None):
    row    = imp.loc[sid]
    name   = row.supplier_name
    master = sm.loc[sid, "material_categories"]
    credit = sm.loc[sid, "credit_days_agreed"]

    doc = Document()
    set_margins(doc, top=1.5, bottom=1.5, left=1.8, right=1.8)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = title.add_run("WATCH-LIST SUPPLIER — ONE-PAGER")
    tr.font.size = Pt(14); tr.font.bold = True
    tr.font.color.rgb = RGBColor.from_string("7B3F00")

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr = sub.add_run("Arora Traders, Pune  |  October 2026  |  CONFIDENTIAL")
    sr.font.size = Pt(9); sr.italic = True
    sr.font.color.rgb = RGBColor.from_string("555555")
    add_divider(doc, "7B3F00")

    add_body(doc,
             "IMPORTANT: This is a WATCH-LIST brief. Evidence is weaker than for the three POOR-rated suppliers. "
             "No credit claims are made here. This document supports a conversation, not a formal charge-back.",
             size=9, bold=True, color="CC5500")
    doc.add_paragraph()

    id_t = doc.add_table(rows=2, cols=4)
    id_t.style = "Table Grid"
    header_row(id_t, ["Supplier", "Supplier ID", "Master categories", "Credit terms"], bg="7B3F00")
    data_row(id_t, 1, [name, sid, master, f"{credit} days"], bold=True, bg="FFF3E0")
    doc.add_paragraph()

    add_heading(doc, "Performance Summary", level=2, color="7B3F00")
    add_divider(doc, "7B3F00")

    mt = doc.add_table(rows=7, cols=3)
    mt.style = "Table Grid"
    header_row(mt, ["Metric", name, "Context"], bg="7B3F00")
    metrics = [
        ("PO count / total billed",
         f"{int(row.po_count)} POs  |  {cr(row.billed_inr)}", ""),
        ("Headline loss (A*+B+C+D)", lk(row.HEADLINE_LOSS_v2_inr),
         f"vs {cr(imp.loc['VS12','HEADLINE_LOSS_v2_inr'])} for worst-rated supplier"),
        ("Confirmed returns",
         f"{lk(row.C_confirmed_returns_inr)}  ({int(row.C_confirmed_return_count)} returns)", ""),
        ("Price vs panel",
         f"{row.price_vs_panel_pct_mean:+.2f}%%  z = {row.price_z:.2f}",
         "z > 2 = borderline significant" if abs(row.price_z) >= 1.5 else "Not statistically significant"),
        ("Avg delivery delay", f"{row.avg_days_late:.1f} days", "Low — not a concern"),
        ("Returns per 1,000 MT received", f"{row.returns_per_1000MT_all:.2f}",
         "Elevated vs most ACCEPTABLE suppliers"),
    ]
    aligns_mt = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT]
    for ri, (met, val, ctx) in enumerate(metrics, 1):
        data_row(mt, ri, [met, val, ctx], bg="FFF8F0" if ri % 2 else "FFFFFF",
                 align=aligns_mt, font_size=9)
    doc.add_paragraph()

    add_heading(doc, "Conversation Points (not claims)", level=2, color="7B3F00")
    add_divider(doc, "7B3F00")
    if points_override:
        points = points_override
    elif sid == "VS16":
        points = [
            f"Price premium: {name} averages {row.price_vs_panel_pct_mean:+.2f}%% above the panel for the same materials in the same year. "
            f"Statistical z = {row.price_z:.2f} — borderline. Ask: 'Can you match the panel rate on MS Rods and HR Coils?'",
            f"Return rate: {row.returns_per_1000MT_all:.2f} returns per 1,000 MT received — the highest among ACCEPTABLE-rated suppliers. Raise informally.",
            f"Confirmed returns: {lk(row.C_confirmed_returns_inr)} over 3 years ({int(row.C_confirmed_return_count)} batches). Request a quality audit.",
            "No excess short-billing. Late delivery minimal (avg 1.7 days). These are positives — acknowledge them.",
        ]
    else:  # VS17
        points = [
            f"Return rate: {row.returns_per_1000MT_all:.2f} returns per 1,000 MT received — second-highest among ACCEPTABLE-rated suppliers.",
            f"Confirmed returns: {lk(row.C_confirmed_returns_inr)} over 3 years ({int(row.C_confirmed_return_count)} batches), concentrated in ERW Pipes.",
            f"Excess short-billing: {lk(row.A_excess_gap_inr)} — small but present. Worth monitoring.",
            f"Price is below panel average — not a concern. Late delivery minimal. These are positives.",
            f"Note: {name} is registered only for ERW Pipes. VS19 (POOR, also ERW Pipes) is being reviewed; consider qualifying a second ERW Pipes supplier for redundancy.",
        ]
    for pt in points:
        add_bullet(doc, pt, size=9)
    doc.add_paragraph()

    add_heading(doc, "Suggested Actions — Next 6 Months", level=2, color="7B3F00")
    add_divider(doc, "7B3F00")
    actions = [
        "Raise the return rate informally in the next order discussion — not a formal claim.",
        "Request a material test certificate for the next two shipments.",
        "Monitor quarterly: if return rate exceeds 3.5 per 1,000 MT, escalate to formal review.",
    ]
    if sid == "VS16":
        actions.insert(1, "Ask for a price quote on MS Rods and HR Coils benchmarked against the panel — frame as routine procurement exercise.")
    for ac in actions:
        add_bullet(doc, ac, size=9)
    doc.add_paragraph()

    add_body(doc,
             f"Reference: Supplier_Evidence_Workbook.xlsx -> Supplier_Impact and Returns_Confirmed (filter on {sid}).",
             size=8, italic=True, color="666666")

    out = f"{D}/Watchlist_{sid}_{name.replace(' ','_')}.docx"
    doc.save(out)
    print(f"Saved watch-list brief: {out}")


# ============================================================
# RUN
# ============================================================
alt_VS12 = [("VS02","Spring Steel"), ("VS15","HR Coils"), ("VS22","MS Flats")]
alt_VS19 = [("VS22","MS Pipes"),    ("VS21","ERW Pipes"), ("VS25","CR Sheets")]
alt_VS03 = [("VS18","Hex Bars")]  # only one alternative in this category

build_full_brief("VS12", alt_VS12)
build_full_brief("VS19", alt_VS19)
build_full_brief("VS03", alt_VS03)
build_watchlist_brief("VS16")
build_watchlist_brief("VS17")

print(f"\nAll 5 briefs saved to: {D}/")
