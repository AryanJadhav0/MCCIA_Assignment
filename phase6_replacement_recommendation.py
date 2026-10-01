"""
Phase 6: Replacement Recommendation (D5)
Generates Replacement_Recommendation.docx
Run:  python phase6_replacement_recommendation.py <output_dir> <raw_dir>
"""
import sys, pandas as pd, numpy as np
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

D   = sys.argv[1] if len(sys.argv) > 1 else "output"
RAW = sys.argv[2] if len(sys.argv) > 2 else "data"

# ---- Load ----
f   = pd.read_csv(f"{D}/po_fact_table.csv", parse_dates=["po_date"])
imp = pd.read_csv(f"{D}/supplier_rupee_impact_v2.csv", index_col=0)
sm  = pd.read_csv(f"{RAW}/supplier_master.csv").set_index("supplier_id")

spend = f.groupby("supplier_id").agg(
    billed=("invoice_amount_billed","sum"),
    qty_received=("quantity_received","sum"),
    po_count=("po_id","count"),
).round(0)
cat_spend = f.groupby(["supplier_id","material_id"]).agg(
    billed=("invoice_amount_billed","sum"),
    qty_received=("quantity_received","sum"),
    po_count=("po_id","count"),
).reset_index()

# ---- Doc helpers ----
def set_cell_bg(cell, hex_color):
    tc = cell._tc; tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)

def set_cell_border(cell, color="BBBBBB"):
    tc = cell._tc; tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge in ("top","left","bottom","right"):
        tag = OxmlElement(f"w:{edge}")
        tag.set(qn("w:val"),"single"); tag.set(qn("w:sz"),"4"); tag.set(qn("w:color"),color)
        tcBorders.append(tag)
    tcPr.append(tcBorders)

def header_row(table, cols, bg="1F3864", text_color="FFFFFF", font_size=9):
    row = table.rows[0]
    for cell, txt in zip(row.cells, cols):
        cell.text = ""
        set_cell_bg(cell, bg)
        p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(txt)
        run.font.bold = True; run.font.size = Pt(font_size)
        run.font.color.rgb = RGBColor.from_string(text_color)

def data_row(table, row_idx, values, bold=False, bg=None, font_size=9, align=None):
    row = table.rows[row_idx]
    for i, (cell, val) in enumerate(zip(row.cells, values)):
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = align[i] if align else WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(str(val))
        run.font.bold = bold; run.font.size = Pt(font_size)
        if bg: set_cell_bg(cell, bg)
        set_cell_border(cell)

def add_heading(doc, text, level=1, color="1F3864"):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True; run.font.size = Pt(14 if level==1 else 11)
    run.font.color.rgb = RGBColor.from_string(color)
    p.paragraph_format.space_before = Pt(8); p.paragraph_format.space_after = Pt(4)

def add_body(doc, text, size=10, bold=False, italic=False, color=None):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size = Pt(size); run.bold = bold; run.italic = italic
    if color: run.font.color.rgb = RGBColor.from_string(color)
    p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(2)

def add_bullet(doc, text, size=10):
    p = doc.add_paragraph(style="List Bullet")
    p.add_run(text).font.size = Pt(size)
    p.paragraph_format.space_before = Pt(1); p.paragraph_format.space_after = Pt(1)

def add_divider(doc, color="1F3864"):
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr(); pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"),"single"); bottom.set(qn("w:sz"),"6"); bottom.set(qn("w:color"),color)
    pBdr.append(bottom); pPr.append(pBdr)
    p.paragraph_format.space_after = Pt(4)

def set_margins(doc, top=1.5, bottom=1.5, left=1.8, right=1.8):
    for s in doc.sections:
        s.top_margin = Cm(top); s.bottom_margin = Cm(bottom)
        s.left_margin = Cm(left); s.right_margin = Cm(right)

def lk(n):  return f"Rs {n/1e5:.1f}L"
def cr(n):  return f"Rs {n/1e7:.2f} Cr"
def pct(n): return f"{n:.2f}%%"

# ---- Alternative finder ----
def get_alternatives(category, exclude_sids):
    cat_map = sm["material_categories"].dropna()
    candidates = [s for s, cats in cat_map.items()
                  if category in [c.strip() for c in cats.split(",")]
                  and s in imp.index and s not in exclude_sids]
    return sorted(candidates, key=lambda s: imp.loc[s,"rank_v2"], reverse=True)  # best rank first

# ============================================================
# BUILD DOC
# ============================================================
doc = Document()
set_margins(doc)

# TITLE
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
tr = title.add_run("SUPPLIER REPLACEMENT RECOMMENDATION")
tr.font.size = Pt(16); tr.font.bold = True
tr.font.color.rgb = RGBColor.from_string("1F3864")

sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
sr = sub.add_run("Arora Traders, Pune  |  Prepared: October 2026  |  CONFIDENTIAL  |  Deliverable D5")
sr.font.size = Pt(9); sr.italic = True
sr.font.color.rgb = RGBColor.from_string("555555")
add_divider(doc)

# EXECUTIVE SUMMARY
add_heading(doc, "Executive Summary")
add_divider(doc, "2E5FA3")

summary_t = doc.add_table(rows=5, cols=2)
summary_t.style = "Table Grid"
header_row(summary_t, ["Finding", "Detail"], bg="2E5FA3")
findings = [
    ("Suppliers recommended for replacement",
     "VS12 Nutan Tubes, VS19 Thakur Steel, VS03 Tata Steel Service — all rated POOR. "
     "Combined 3-year headline loss Rs 9.06 Cr (excess billing + rejections + returns)."),
    ("Feasibility",
     "Viable alternatives exist on the panel for all master-listed categories. "
     "Hex Bars (VS03's only master category) has only one registered alternative — thin cover, noted."),
    ("Off-master spend risk",
     "87.1% of all POs (Rs 874.7 Cr over 3 years) go to suppliers for materials outside their master list. "
     "Recommend a category registration update before re-allocating volume."),
    ("Transition approach",
     "Parallel-run the alternative suppliers for 2-3 months before exiting. "
     "Do not exit VS03 on Hex Bars until VS18 Kumar Hardware capacity is confirmed."),
]
aligns = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT]
for ri, (f_title, f_detail) in enumerate(findings, 1):
    data_row(summary_t, ri, [f_title, f_detail], bg="F0F4FF" if ri%2 else "FFFFFF",
             align=aligns, font_size=9)
doc.add_paragraph()

# ============================================================
# SECTION 1 — POOR SUPPLIER REPLACEMENT ANALYSIS
# ============================================================
add_heading(doc, "1. Replacement Analysis — Three POOR-Rated Suppliers")
add_divider(doc, "2E5FA3")

poor_suppliers = [
    ("VS12", "Nutan Tubes",       ["Spring Steel","HR Coils","MS Flats"]),
    ("VS19", "Thakur Steel",      ["MS Pipes","ERW Pipes","CR Sheets"]),
    ("VS03", "Tata Steel Service",["Hex Bars"]),
]

for poor_sid, poor_name, master_cats in poor_suppliers:
    poor_row = imp.loc[poor_sid]
    add_heading(doc, f"1.{poor_suppliers.index((poor_sid,poor_name,master_cats))+1}  {poor_sid} — {poor_name}", level=2, color="C00000")
    add_divider(doc, "C00000")

    # Summary stats for POOR supplier
    st = doc.add_table(rows=2, cols=5)
    st.style = "Table Grid"
    header_row(st, ["Scorecard rank","Headline loss (3 yr)","Short delivery %","Avg days late","Master categories"], bg="C00000")
    data_row(st, 1,
             [f"#{int(poor_row.rank_v2)} of 34 (POOR)",
              cr(poor_row.HEADLINE_LOSS_v2_inr),
              f"{poor_row.short_pct_of_qty:.2f}%%",
              f"{poor_row.avg_days_late:.1f} days",
              ", ".join(master_cats)],
             bg="FFE8E8", bold=True,
             align=[WD_ALIGN_PARAGRAPH.CENTER]*5, font_size=9)
    doc.add_paragraph()

    for category in master_cats:
        # Spend with the POOR supplier in this category
        poor_cat = cat_spend[(cat_spend.supplier_id==poor_sid)&(cat_spend.material_id==category)]
        poor_cat_billed  = poor_cat["billed"].sum()
        poor_cat_qty     = poor_cat["qty_received"].sum()
        poor_cat_po      = int(poor_cat["po_count"].sum())

        alts = get_alternatives(category, exclude_sids=["VS12","VS19","VS03"])

        n_rows = max(len(alts)+1, 2)
        alt_t = doc.add_table(rows=n_rows+1, cols=7)
        alt_t.style = "Table Grid"
        header_row(alt_t,
                   ["Supplier", "Category", "Panel rank", "Headline loss",
                    "Short %", "Avg late", "Capacity note"],
                   bg="1F5C2E")

        # Row 0 = current POOR supplier
        data_row(alt_t, 1,
                 [f"{poor_sid} — {poor_name} (CURRENT / TO EXIT)",
                  category,
                  f"#{int(poor_row.rank_v2)} POOR",
                  cr(poor_row.HEADLINE_LOSS_v2_inr),
                  f"{poor_row.short_pct_of_qty:.2f}%%",
                  f"{poor_row.avg_days_late:.1f} d",
                  f"Current: {lk(poor_cat_billed)} billed, {poor_cat_qty:.0f} MT, {poor_cat_po} POs"],
                 bg="FFE8E8", bold=True, font_size=8,
                 align=[WD_ALIGN_PARAGRAPH.LEFT]*7)

        for ri, alt_sid in enumerate(alts[:n_rows-1], 2):
            alt_row = imp.loc[alt_sid]
            alt_name = alt_row.supplier_name
            alt_cat = cat_spend[(cat_spend.supplier_id==alt_sid)&(cat_spend.material_id==category)]
            alt_cat_billed = alt_cat["billed"].sum()
            alt_cat_qty    = alt_cat["qty_received"].sum()

            # Capacity note: can they absorb the volume from POOR supplier?
            can_absorb = alt_cat_qty > 0
            if can_absorb:
                headroom_ratio = poor_cat_qty / alt_cat_qty if alt_cat_qty > 0 else 999
                if headroom_ratio < 0.5:
                    cap_note = f"Low load: currently {alt_cat_qty:.0f} MT; can absorb {poor_cat_qty:.0f} MT easily"
                elif headroom_ratio < 1.5:
                    cap_note = f"Moderate load: {alt_cat_qty:.0f} MT current; {poor_cat_qty:.0f} MT to absorb — confirm capacity"
                else:
                    cap_note = f"High load: {alt_cat_qty:.0f} MT current; {poor_cat_qty:.0f} MT to absorb — split across 2+ alts"
            else:
                cap_note = f"Not currently supplying this material — qualify first"

            bg_color = "E8F5E9" if ri % 2 == 0 else "F1F8F2"
            data_row(alt_t, ri,
                     [f"{alt_sid} — {alt_name}",
                      category,
                      f"#{int(alt_row.rank_v2)} ACCEPTABLE",
                      lk(alt_row.HEADLINE_LOSS_v2_inr),
                      f"{alt_row.short_pct_of_qty:.2f}%%",
                      f"{alt_row.avg_days_late:.1f} d",
                      cap_note],
                     bg=bg_color, font_size=8,
                     align=[WD_ALIGN_PARAGRAPH.LEFT]*7)

        if not alts:
            data_row(alt_t, 2,
                     ["NO REGISTERED ALTERNATIVES — panel has no other supplier for this category","","","","","",""],
                     bg="FFEECC", bold=True, font_size=8)

        doc.add_paragraph()

    # Per-supplier recommendation text
    if poor_sid == "VS12":
        add_body(doc,
                 "Recommendation for VS12 Nutan Tubes: Exit across all three master categories (Spring Steel, HR Coils, MS Flats). "
                 "Preferred transition: VS02 SAIL Distribution for Spring Steel (ranked #28, lowest loss on panel, Rs 29.2L billed in that category already). "
                 "VS15 Western Steel for HR Coils (ranked #32, lowest overall loss on panel at Rs 4.5L, currently supplying 212 MT). "
                 "VS29 Global Metals for MS Flats (ranked #22, moderate loss, 148 MT in category). "
                 "Parallel-run for 3 months before exit.",
                 size=9, color="1F3864")
    elif poor_sid == "VS19":
        add_body(doc,
                 "Recommendation for VS19 Thakur Steel: Exit across all three master categories (MS Pipes, ERW Pipes, CR Sheets). "
                 "Preferred transition: VS22 Ratan Steel for MS Pipes (ranked #34, lowest loss on entire panel at Rs 3.5L, already active in category). "
                 "VS21 Viraj Profiles for ERW Pipes (ranked #25, Rs 8.2L loss, 215 MT current — strong capacity). "
                 "VS25 Laxmi Hardware for CR Sheets (ranked #24, Rs 9.8L loss, 152 MT in category). "
                 "Parallel-run for 3 months before exit.",
                 size=9, color="1F3864")
    else:  # VS03
        add_body(doc,
                 "Recommendation for VS03 Tata Steel Service: Exit Hex Bars (only master category). "
                 "Single alternative: VS18 Kumar Hardware (ranked #19, Rs 13.7L loss, currently 128 MT in Hex Bars). "
                 "Note: VS03 supplies 234 MT of Hex Bars — VS18 must confirm capacity to absorb this before exit. "
                 "If VS18 cannot fully absorb, split with off-master supplier until a second Hex Bars supplier is formally registered. "
                 "Note also that VS03 handles significant off-master volume (Rs 28.1 Cr) — see Section 3 for how to redistribute that.",
                 size=9, color="1F3864")
    doc.add_paragraph()

# ============================================================
# SECTION 2 — SCORECARD COMPARISON
# ============================================================
add_heading(doc, "2. Scorecard Comparison — POOR vs Recommended Replacements")
add_divider(doc, "2E5FA3")

add_body(doc, "All 31 non-POOR suppliers score 91-100 composite (ACCEPTABLE). "
         "The three POOR suppliers score 9-11. The table below shows the proposed replacements alongside the suppliers being exited.",
         size=9)
doc.add_paragraph()

compare_pairs = [
    ("VS12","VS02","Spring Steel"),
    ("VS12","VS15","HR Coils"),
    ("VS12","VS29","MS Flats"),
    ("VS19","VS22","MS Pipes"),
    ("VS19","VS21","ERW Pipes"),
    ("VS19","VS25","CR Sheets"),
    ("VS03","VS18","Hex Bars"),
]

ct = doc.add_table(rows=len(compare_pairs)+1, cols=8)
ct.style = "Table Grid"
header_row(ct, ["Category","Exit (POOR)","Rank","Short %","Late (d)",
                "Replacement","Rank","Short %"])
aligns_ct = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT,
             WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER,
             WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER]
for ri, (ex_sid, rep_sid, cat_name) in enumerate(compare_pairs, 1):
    ex  = imp.loc[ex_sid]
    rep = imp.loc[rep_sid]
    data_row(ct, ri,
             [cat_name,
              f"{ex_sid} {ex.supplier_name}", f"#{int(ex.rank_v2)} POOR",
              f"{ex.short_pct_of_qty:.2f}%%", f"{ex.avg_days_late:.1f}",
              f"{rep_sid} {rep.supplier_name}", f"#{int(rep.rank_v2)} ACCEPTABLE",
              f"{rep.short_pct_of_qty:.2f}%%"],
             bg="FFE8E8" if ri % 2 else "FFF0F0",
             align=aligns_ct, font_size=8)
doc.add_paragraph()

# ============================================================
# SECTION 3 — OFF-MASTER SPEND
# ============================================================
add_heading(doc, "3. Off-Master Spend — The 87% Problem")
add_divider(doc, "2E5FA3")

total_spend  = f.invoice_amount_billed.sum()
off_spend    = f[~f.in_master].invoice_amount_billed.sum()
off_po_count = (~f.in_master).sum()

add_body(doc,
         f"{off_po_count:,} of 4,800 POs ({off_po_count/4800*100:.1f}%%) are placed with suppliers for materials "
         f"outside their master category list. This represents Rs {off_spend/1e7:.1f} Cr of the Rs {total_spend/1e7:.1f} Cr "
         f"total spend over 3 years — meaning supplier accountability is very hard to enforce for 87%% of purchases.",
         size=9)
doc.add_paragraph()

add_body(doc, "Root cause: supplier_master.csv lists very narrow master categories (e.g. VS03 is registered only for 'Hex Bars' "
         "but supplies 18 different material types). The master list has not been updated to reflect actual supply relationships.", size=9)
doc.add_paragraph()

add_heading(doc, "Recommended actions for off-master spend", level=2)
add_divider(doc, "4472C4")

actions = [
    "Conduct a category registration audit: for each material type actually supplied in the last 3 years, "
    "confirm which suppliers are qualified and update supplier_master accordingly. "
    "This brings 87%% of POs into scope for formal scorecard monitoring.",
    "Priority materials to register: MS Flats (Rs 60.0 Cr off-master spend), Square/Rect Tubes (Rs 55.5 Cr), "
    "MS Sheets (Rs 53.4 Cr), MS Rods (Rs 53.4 Cr), Spring Steel (Rs 52.5 Cr).",
    "For materials with only one registered supplier (e.g. Square/Rect Tubes — only VS30 Apex Hardware, "
    "GI Pipes — only VS32 Precision Steel): qualify at least one additional supplier before enforcing "
    "master-only purchasing.",
    "In the short term, continue using off-master suppliers for continuity but flag each off-master PO "
    "in the procurement system so the team knows no formal quality SLA is in force.",
    "After the registration audit, re-run the scorecard on the updated master list. "
    "The current 9.2 / 9.3 / 10.8 scores for the three POOR suppliers will remain — "
    "the methodology is performance-only, not category-dependent.",
]
for a in actions:
    add_bullet(doc, a, size=9)
doc.add_paragraph()

# Off-master top materials table
add_heading(doc, "Top off-master materials by spend (Rs, 3 years)", level=2)
add_divider(doc, "4472C4")

off_mat = f[~f.in_master].groupby("material_id").agg(
    billed=("invoice_amount_billed","sum"),
    po_count=("po_id","count")
).sort_values("billed", ascending=False).head(10).reset_index()

# Find who is registered for each
cat_map_dict = {}
for s, cats in sm["material_categories"].dropna().items():
    for c in cats.split(","):
        c = c.strip()
        cat_map_dict.setdefault(c, [])
        cat_map_dict[c].append(s)

omt = doc.add_table(rows=len(off_mat)+1, cols=4)
omt.style = "Table Grid"
header_row(omt, ["Material", "Off-master spend", "Off-master POs", "Registered suppliers on panel"])
aligns_omt = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT,
              WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT]
for ri, row_d in enumerate(off_mat.itertuples(), 1):
    reg = cat_map_dict.get(row_d.material_id, [])
    reg_str = ", ".join([f"{s} {imp.loc[s,'supplier_name']}" for s in reg if s in imp.index]) or "None registered"
    data_row(omt, ri,
             [row_d.material_id, lk(row_d.billed), int(row_d.po_count), reg_str],
             bg="F9FAFB" if ri%2 else "FFFFFF", align=aligns_omt, font_size=8)
doc.add_paragraph()

# ============================================================
# SECTION 4 — RISKS AND MITIGATIONS
# ============================================================
add_heading(doc, "4. Risks and Mitigations")
add_divider(doc, "2E5FA3")

risks = [
    ("Thin cover: Hex Bars",
     "VS03 is the primary Hex Bars supplier (234 MT, Rs 1.89 Cr). VS18 is the only registered alternative. "
     "If VS18 cannot absorb the volume, Arora must run off-master until a second supplier is qualified.",
     "Before exit: confirm VS18 capacity in writing. Identify and onboard one additional Hex Bars supplier within 6 months."),
    ("Disruption during transition",
     "The three POOR suppliers collectively handle Rs 93.1 Cr of spend. Simultaneous exit would be disruptive.",
     "Staggered exit: VS12 first (highest loss), then VS19, then VS03. Each with a 3-month parallel run. "
     "Target full exit within 12 months."),
    ("Replacement suppliers inherit volume spikes",
     "Doubling or tripling volume to a single replacement could degrade their performance.",
     "Spread volume across 2-3 alternatives per category. Monitor their scorecards quarterly after transition."),
    ("Supplier relationships and Kiran's pushback",
     "VS12, VS19, VS03 have 3, 12, and 2 years of relationship respectively. VS19 and VS03 may resist exit.",
     "Run the negotiation briefs (D4) first. If suppliers refuse to correct pricing and billing, "
     "the replacement conversation follows naturally from the documented data."),
    ("Off-master spend remains untracked",
     "Even after replacing the three POOR suppliers, 87%% of POs are off-master and unscored.",
     "Category registration audit (Section 3) is the highest-leverage systemic fix. "
     "Target: bring at least the top 10 materials into the master within 3 months."),
]

rt = doc.add_table(rows=len(risks)+1, cols=3)
rt.style = "Table Grid"
header_row(rt, ["Risk", "Description", "Mitigation"])
aligns_rt = [WD_ALIGN_PARAGRAPH.LEFT]*3
for ri, (risk, desc, mit) in enumerate(risks, 1):
    data_row(rt, ri, [risk, desc, mit],
             bg="FFF8F0" if ri%2 else "FFFFFF", align=aligns_rt, font_size=9)
doc.add_paragraph()

# ============================================================
# SECTION 5 — SUMMARY TABLE
# ============================================================
add_heading(doc, "5. Decision Summary")
add_divider(doc, "2E5FA3")

dec_t = doc.add_table(rows=8, cols=4)
dec_t.style = "Table Grid"
header_row(dec_t, ["Supplier", "Decision", "Preferred replacement(s)", "Timeline"])
decisions = [
    ("VS12 Nutan Tubes (POOR #1)",
     "EXIT all master categories",
     "Spring Steel: VS02 SAIL Distribution\nHR Coils: VS15 Western Steel\nMS Flats: VS29 Global Metals",
     "3-month parallel, then exit"),
    ("VS19 Thakur Steel (POOR #2)",
     "EXIT all master categories",
     "MS Pipes: VS22 Ratan Steel\nERW Pipes: VS21 Viraj Profiles\nCR Sheets: VS25 Laxmi Hardware",
     "After VS12 exit, 3-month parallel"),
    ("VS03 Tata Steel Service (POOR #3)",
     "EXIT Hex Bars (only master cat)\n— confirm capacity first",
     "Hex Bars: VS18 Kumar Hardware\n(only registered alternative)",
     "After VS19 exit; confirm capacity before exit"),
    ("VS16 Garg Enterprises (WATCH)",
     "RETAIN — monitor price and returns",
     "None (ACCEPTABLE rating)",
     "6-month review if return rate rises"),
    ("VS17 Modern Metals (WATCH)",
     "RETAIN — monitor return rate",
     "None (ACCEPTABLE rating)",
     "6-month review if return rate rises"),
    ("Off-master spend (87%% of POs)",
     "AUDIT — update supplier_master",
     "Register existing suppliers for categories they actually supply",
     "Start within 1 month; complete in 3 months"),
    ("Scorecard re-run",
     "Re-run after registration audit",
     "phase3_scorecard.py",
     "After master update"),
]
aligns_dec = [WD_ALIGN_PARAGRAPH.LEFT]*4
for ri, (sup, dec, rep, tl) in enumerate(decisions, 1):
    bg = "FFE8E8" if "EXIT" in dec else ("FFF8E7" if "WATCH" in dec else ("E8F5E9" if "AUDIT" in dec else "F9FAFB"))
    data_row(dec_t, ri, [sup, dec, rep, tl], bg=bg, align=aligns_dec, font_size=9)

doc.add_paragraph()
add_body(doc,
         "Methodology reference: Replacement candidates are selected from the panel based on (1) registration in the same master category as the POOR supplier, "
         "(2) ACCEPTABLE scorecard tier (composite score >= 85), and (3) existing volume in that category as a capacity proxy. "
         "years_of_relationship and is_underperformer were not used in any scoring or selection. "
         "All figures from Supplier_Evidence_Workbook.xlsx and the CSV outputs in the output/ folder.",
         size=8, italic=True, color="666666")

out = f"{D}/Replacement_Recommendation.docx"
doc.save(out)
print(f"Saved: {out}")
