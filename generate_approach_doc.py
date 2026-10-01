import sys
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

def add_heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = RGBColor(31, 56, 100) # Professional Navy Blue
        run.font.name = 'Calibri'
    return h

def main():
    doc = Document()
    
    # Title Section
    title = doc.add_heading('Arora Traders: Supplier Blindspot Resolution', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.color.rgb = RGBColor(31, 56, 100)
        
    subtitle = doc.add_paragraph('Analytical Framework, Technical Solution, and Strategic Assumptions')
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in subtitle.runs:
        run.font.italic = True
        run.font.size = Pt(14)
        run.font.color.rgb = RGBColor(89, 89, 89)
        
    doc.add_paragraph('\n')

    # 1. Approach
    add_heading(doc, '1. Executive Summary & Approach', level=1)
    p = doc.add_paragraph('The objective was to identify the root cause of financial leakages and supply chain inefficiencies ("The Supplier Blindspot") among 34 suppliers for Arora Traders over a 3-year period covering 4,800 POs.\n\nOur approach utilized a ')
    p.add_run('Multi-Stage Data Pipeline:').bold = True
    
    ul = doc.add_paragraph(style='List Bullet')
    ul.add_run('Attribution & Unification: ').bold = True
    ul.add_run('Mapped off-master POs to their correct supplier taxonomy using algorithmic matching on names and categorizations.')
    
    ul = doc.add_paragraph(style='List Bullet')
    ul.add_run('Financial Impact Modeling: ').bold = True
    ul.add_run('Transitioned from vague operational metrics to hard financial losses (A*+B+C+D Framework):')
    
    doc.add_paragraph('• A*: Excess Billing Gap (invoice quantities > received quantities)', style='List Bullet 2')
    doc.add_paragraph('• B: Rejection Losses (material billed in full but rejected)', style='List Bullet 2')
    doc.add_paragraph('• C: Confirmed Returns Traced', style='List Bullet 2')
    doc.add_paragraph('• D: Inferred Returns (probabilistic allocation for untraced returns)', style='List Bullet 2')
    
    ul = doc.add_paragraph(style='List Bullet')
    ul.add_run('Composite Scoring Engine: ').bold = True
    ul.add_run('Developed a weighted scoring algorithm to rank suppliers objectively based on delivery compliance, quality, and pricing deviations.')

    # 2. Technical Solution
    add_heading(doc, '2. The Technical Solution', level=1)
    
    ul = doc.add_paragraph(style='List Bullet')
    ul.add_run('Data Engineering Pipeline: ').bold = True
    ul.add_run('Built modular Python scripts to clean, merge, and evaluate the underlying POs, receipts, and market panel rates, segmenting suppliers into POOR, WATCH, and ACCEPTABLE tiers.')
    
    ul = doc.add_paragraph(style='List Bullet')
    ul.add_run('Interactive Web Dashboard: ').bold = True
    ul.add_run('Developed a fully interactive HTML/JS dashboard tracking KPIs, loss composition, and supplier scorecards. Hosted directly on Vercel, it allows stakeholders to visualize the data dynamically.')
    
    ul = doc.add_paragraph(style='List Bullet')
    ul.add_run('Automated Negotiation Briefs: ').bold = True
    ul.add_run('Auto-generated printable strategic briefs for the bottom 5 suppliers, outlining their exact financial leakage and the required action plans.')

    # 3. Key Assumptions
    add_heading(doc, '3. Key Assumptions & Constraints', level=1)
    
    ul = doc.add_paragraph(style='List Number')
    ul.add_run('Inferred Returns Distribution: ').bold = True
    ul.add_run('Untraced returns were probabilistically distributed among suppliers based on their historical failure rates (rejections/short deliveries) rather than assuming equal blame across the board.')
    
    ul = doc.add_paragraph(style='List Number')
    ul.add_run('Market Rate Baseline: ').bold = True
    ul.add_run('The "Panel Ideal" was established by aggregating the top quartile of supplier pricing, assuming this baseline represents the fair market value for comparative cost-analysis.')
    
    ul = doc.add_paragraph(style='List Number')
    ul.add_run('Cost of Capital Exclusions: ').bold = True
    ul.add_run('While late deliveries disrupt operations, direct financial penalties for delays were excluded from the Headline Loss calculation unless explicitly stipulated by SLAs. The financial model focuses strictly on direct cash-loss (shorting, rejecting).')

    # 4. Out of the box value
    add_heading(doc, '4. Out-of-the-Box Problem Solving Approach', level=1)
    
    ul = doc.add_paragraph(style='List Bullet')
    ul.add_run('Probabilistic Financial Attribution (Inferred Returns): ').bold = True
    ul.add_run('Instead of ignoring untraced returns or writing them off as a general business loss, we pioneered a probabilistic model. We assigned the financial penalty of untraced returns back to specific suppliers based on their historical quality failure rates. This ensures accountability even when direct traceability is broken.')
    
    ul = doc.add_paragraph(style='List Bullet')
    ul.add_run('Translating Operations to Unified Rupee Impact: ').bold = True
    ul.add_run('Traditional dashboards stop at operational metrics (e.g., "Supplier X is 5 days late"). We took the out-of-the-box approach of converting every operational failure—short deliveries, rejections, and pricing gaps—into a single, unified currency: "Headline Rupee Loss." This allows executives to instantly prioritize interventions based on absolute financial impact.')
    
    ul = doc.add_paragraph(style='List Bullet')
    ul.add_run('Bridging Data and Execution (Negotiation Briefs): ').bold = True
    ul.add_run('Rather than just presenting a dashboard of numbers, the solution extends directly into business execution. By automatically generating structured "Negotiation Briefs," we empower procurement managers with exact negotiation scripts and corrective action plans tailored to each underperforming supplier.')

    # Save
    out_path = 'Arora_Traders_Approach_Document.docx'
    doc.save(out_path)
    print(f"Successfully generated {out_path}")

if __name__ == '__main__':
    main()
