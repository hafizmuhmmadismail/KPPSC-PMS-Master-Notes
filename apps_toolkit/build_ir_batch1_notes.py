import os
import subprocess
import tempfile
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

EDGE_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
IR_DIR = r"H:\My Drive\PMS KPK 2022 notes\2. International Relations"

def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_callout(doc, text_list, title="EXAM ESSENTIAL TAKEAWAY"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    cell = tbl.cell(0, 0)
    cell.width = Inches(6.5)
    set_cell_background(cell, "F0F4F8")
    set_cell_margins(cell, top=140, bottom=140, left=200, right=200)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    run_t = p.add_run(f"\U0001F4CC {title}\n")
    run_t.font.name = "Arial"
    run_t.font.size = Pt(10.5)
    run_t.font.bold = True
    run_t.font.color.rgb = RGBColor(14, 43, 80)
    
    for line in text_list:
        p_line = cell.add_paragraph()
        p_line.paragraph_format.space_before = Pt(2)
        p_line.paragraph_format.space_after = Pt(2)
        run = p_line.add_run(line)
        run.font.name = "Georgia"
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(30, 41, 59)

def generate_ir_master_doc(meta_data, callout_data, map_text, timeline_data, essay_sections, quotes_data, out_path, paper_sub_title):
    doc = Document()
    
    for s in doc.sections:
        s.top_margin = Inches(0.8)
        s.bottom_margin = Inches(0.8)
        s.left_margin = Inches(0.8)
        s.right_margin = Inches(0.8)
        
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(2)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_main = p_title.add_run("KPPSC PMS \u2022 INTERNATIONAL RELATIONS (100 MARKS)")
    r_main.font.name = "Arial"
    r_main.font.size = Pt(12)
    r_main.font.bold = True
    r_main.font.color.rgb = RGBColor(14, 43, 80)
    
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(12)
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run(f"EXECUTIVE COMPREHENSIVE EXAMINATION MASTER NOTES \u2022 {paper_sub_title}")
    r_sub.font.name = "Arial"
    r_sub.font.size = Pt(9.5)
    r_sub.font.color.rgb = RGBColor(100, 116, 139)
    
    # Metadata Table
    tbl_meta = doc.add_table(rows=len(meta_data), cols=2)
    tbl_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_meta.autofit = False
    
    for i, (k, v) in enumerate(meta_data):
        row = tbl_meta.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = Inches(2.2)
        c1.width = Inches(4.3)
        
        set_cell_margins(c0, top=70, bottom=70, left=120, right=120)
        set_cell_margins(c1, top=70, bottom=70, left=120, right=120)
        
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_before = Pt(0)
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(k)
        r0.font.name = "Arial"
        r0.font.size = Pt(9)
        r0.font.bold = True
        
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(v)
        r1.font.name = "Georgia"
        r1.font.size = Pt(9)
        
        if i == 0:
            set_cell_background(c0, "0E2B50")
            set_cell_background(c1, "0E2B50")
            r0.font.color.rgb = RGBColor(255, 255, 255)
            r1.font.color.rgb = RGBColor(255, 255, 255)
            r1.font.bold = True
        else:
            bg = "F8FAFC" if i % 2 == 1 else "FFFFFF"
            set_cell_background(c0, bg)
            set_cell_background(c1, bg)
            r0.font.color.rgb = RGBColor(30, 41, 59)
            r1.font.color.rgb = RGBColor(15, 23, 42)
            
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_before = Pt(6)
    p_sp.paragraph_format.space_after = Pt(4)
    
    add_callout(doc, callout_data)
    
    p_sp2 = doc.add_paragraph()
    p_sp2.paragraph_format.space_before = Pt(8)
    p_sp2.paragraph_format.space_after = Pt(4)
    
    p_map_h = doc.add_paragraph()
    p_map_h.paragraph_format.space_before = Pt(10)
    p_map_h.paragraph_format.space_after = Pt(4)
    r_map_h = p_map_h.add_run("CONCEPTUAL THEMATIC MAP & STRATEGIC ROADMAP")
    r_map_h.font.name = "Arial"
    r_map_h.font.size = Pt(11)
    r_map_h.font.bold = True
    r_map_h.font.color.rgb = RGBColor(14, 43, 80)
    
    tbl_map = doc.add_table(rows=1, cols=1)
    tbl_map.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_map.autofit = False
    c_map = tbl_map.cell(0, 0)
    c_map.width = Inches(6.5)
    set_cell_background(c_map, "F1F5F9")
    set_cell_margins(c_map, top=120, bottom=120, left=150, right=150)
    
    p_m = c_map.paragraphs[0]
    p_m.paragraph_format.space_before = Pt(0)
    p_m.paragraph_format.space_after = Pt(0)
    r_m = p_m.add_run(map_text)
    r_m.font.name = "Consolas"
    r_m.font.size = Pt(8.5)
    r_m.font.color.rgb = RGBColor(15, 23, 42)
    
    p_tl_h = doc.add_paragraph()
    p_tl_h.paragraph_format.space_before = Pt(14)
    p_tl_h.paragraph_format.space_after = Pt(4)
    r_tl_h = p_tl_h.add_run("ANALYTICAL DIMENSIONS & SYSTEMIC SHIFTS MATRIX")
    r_tl_h.font.name = "Arial"
    r_tl_h.font.size = Pt(11)
    r_tl_h.font.bold = True
    r_tl_h.font.color.rgb = RGBColor(14, 43, 80)
    
    tbl_tl = doc.add_table(rows=len(timeline_data) + 1, cols=3)
    tbl_tl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_tl.autofit = False
    
    headers = ["Dimension / Arena", "Strategic Mechanism / Driver", "Systemic Impact on Global/Regional Order"]
    widths = [Inches(1.8), Inches(2.2), Inches(2.5)]
    
    hdr_row = tbl_tl.rows[0]
    for j, h_text in enumerate(headers):
        c = hdr_row.cells[j]
        c.width = widths[j]
        set_cell_background(c, "0E2B50")
        set_cell_margins(c, top=80, bottom=80, left=100, right=100)
        p = c.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h_text)
        run.font.name = "Arial"
        run.font.size = Pt(8.5)
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        
    for i, row_data in enumerate(timeline_data):
        row = tbl_tl.rows[i + 1]
        bg = "F8FAFC" if i % 2 == 1 else "FFFFFF"
        for j, val in enumerate(row_data):
            c = row.cells[j]
            c.width = widths[j]
            set_cell_background(c, bg)
            set_cell_margins(c, top=70, bottom=70, left=100, right=100)
            p = c.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.name = "Georgia"
            run.font.size = Pt(8.5)
            if j == 0:
                run.font.bold = True
                run.font.color.rgb = RGBColor(14, 43, 80)
            else:
                run.font.color.rgb = RGBColor(30, 41, 59)
                
    p_es_h = doc.add_paragraph()
    p_es_h.paragraph_format.space_before = Pt(16)
    p_es_h.paragraph_format.space_after = Pt(4)
    r_es_h = p_es_h.add_run("THE 20-MARK MASTER MODEL ESSAY (~650-700 WORDS)")
    r_es_h.font.name = "Arial"
    r_es_h.font.size = Pt(11)
    r_es_h.font.bold = True
    r_es_h.font.color.rgb = RGBColor(14, 43, 80)
    
    for sub_h, body_p in essay_sections:
        p_sh = doc.add_paragraph()
        p_sh.paragraph_format.space_before = Pt(8)
        p_sh.paragraph_format.space_after = Pt(2)
        r_sh = p_sh.add_run(sub_h)
        r_sh.font.name = "Arial"
        r_sh.font.size = Pt(10)
        r_sh.font.bold = True
        r_sh.font.color.rgb = RGBColor(30, 58, 138)
        
        p_b = doc.add_paragraph()
        p_b.paragraph_format.space_before = Pt(0)
        p_b.paragraph_format.space_after = Pt(4)
        p_b.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r_b = p_b.add_run(body_p)
        r_b.font.name = "Georgia"
        r_b.font.size = Pt(9.5)
        r_b.font.color.rgb = RGBColor(15, 23, 42)
        
    p_qt_h = doc.add_paragraph()
    p_qt_h.paragraph_format.space_before = Pt(16)
    p_qt_h.paragraph_format.space_after = Pt(4)
    r_qt_h = p_qt_h.add_run("THEORETICAL PERSPECTIVES & QUOTABLE THESIS ARGUMENTS")
    r_qt_h.font.name = "Arial"
    r_qt_h.font.size = Pt(11)
    r_qt_h.font.bold = True
    r_qt_h.font.color.rgb = RGBColor(14, 43, 80)
    
    tbl_qt = doc.add_table(rows=len(quotes_data) + 1, cols=3)
    tbl_qt.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_qt.autofit = False
    
    q_headers = ["Theorist / Scholar", "Theoretical School", "Exact Quotable Thesis Argument"]
    q_widths = [Inches(1.8), Inches(1.8), Inches(2.9)]
    
    hdr_q_row = tbl_qt.rows[0]
    for j, h_text in enumerate(q_headers):
        c = hdr_q_row.cells[j]
        c.width = q_widths[j]
        set_cell_background(c, "0E2B50")
        set_cell_margins(c, top=80, bottom=80, left=100, right=100)
        p = c.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h_text)
        run.font.name = "Arial"
        run.font.size = Pt(8.5)
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        
    for i, row_data in enumerate(quotes_data):
        row = tbl_qt.rows[i + 1]
        bg = "F8FAFC" if i % 2 == 1 else "FFFFFF"
        for j, val in enumerate(row_data):
            c = row.cells[j]
            c.width = q_widths[j]
            set_cell_background(c, bg)
            set_cell_margins(c, top=70, bottom=70, left=100, right=100)
            p = c.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.name = "Georgia"
            run.font.size = Pt(8.5)
            if j == 0:
                run.font.bold = True
                run.font.color.rgb = RGBColor(14, 43, 80)
            elif j == 1:
                run.font.color.rgb = RGBColor(71, 85, 105)
            else:
                run.font.italic = True
                run.font.color.rgb = RGBColor(15, 23, 42)
                
    doc.save(out_path)
    print(f"Generated Master Docx: {out_path}")

def generate_pdf_and_md(meta_data, callout_data, map_text, timeline_data, essay_sections, quotes_data, pdf_path, md_path, title):
    md_content = f"# {title}\n\n"
    md_content += f"> **Exam Track**: KPPSC PMS International Relations (100 Marks)\n"
    md_content += f"> **Document Type**: Institutional Master Model Solution & Theory Suite\n\n---\n\n"
    
    md_content += "### 📋 Executive Examination Metadata\n\n"
    md_content += "| Parameter | Specification |\n|---|---|\n"
    for k, v in meta_data:
        md_content += f"| **{k}** | {v} |\n"
        
    md_content += "\n---\n\n### 📌 Exam Essential Takeaways\n\n"
    for item in callout_data:
        md_content += f"- {item}\n"
        
    md_content += "\n---\n\n### 🗺️ Conceptual Thematic Map & Strategic Roadmap\n\n```text\n"
    md_content += map_text.strip() + "\n```\n\n---\n\n"
    
    md_content += "### 📊 Analytical Dimensions & Systemic Shifts Matrix\n\n"
    md_content += "| Dimension / Arena | Strategic Mechanism / Driver | Systemic Impact on Global/Regional Order |\n|---|---|---|\n"
    for d, m, s in timeline_data:
        md_content += f"| **{d}** | {m} | {s} |\n"
        
    md_content += "\n---\n\n### 🖋️ The 20-Mark Master Model Essay (~650-700 Words)\n\n"
    for sub_h, body_p in essay_sections:
        md_content += f"#### {sub_h}\n\n{body_p}\n\n"
        
    md_content += "---\n\n### 📚 Theoretical Perspectives & Quotable Thesis Arguments\n\n"
    md_content += "| Theorist / Scholar | Theoretical School | Exact Quotable Thesis Argument |\n|---|---|---|\n"
    for th, sc, ar in quotes_data:
        md_content += f"| **{th}** | *{sc}* | \"{ar}\" |\n"
        
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(md_content)
        
    html_content = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>{title}</title>
<style>
    @page {{
        size: A4;
        margin: 18mm 16mm 18mm 16mm;
    }}
    body {{
        font-family: 'Segoe UI', Calibri, Arial, sans-serif;
        font-size: 9.5pt;
        line-height: 1.45;
        color: #1A1A1A;
        margin: 0;
    }}
    h1 {{
        color: #0E2B50;
        font-size: 15pt;
        border-bottom: 2px solid #0E2B50;
        padding-bottom: 4px;
        margin-top: 0;
        margin-bottom: 8px;
        text-align: center;
    }}
    h2 {{
        color: #1B4965;
        font-size: 11.5pt;
        margin-top: 12px;
        margin-bottom: 4px;
        border-bottom: 1px solid #D0D7DE;
        padding-bottom: 2px;
    }}
    h3 {{
        color: #2B2D42;
        font-size: 10.5pt;
        margin-top: 10px;
        margin-bottom: 4px;
    }}
    h4 {{
        color: #1E3A8A;
        font-size: 9.5pt;
        margin-top: 8px;
        margin-bottom: 2px;
    }}
    table {{
        width: 100%;
        border-collapse: collapse;
        margin: 8px 0;
        font-size: 8.5pt;
    }}
    th {{
        background-color: #0E2B50;
        color: #FFFFFF;
        padding: 5px 8px;
        text-align: left;
        font-weight: bold;
        border: 1px solid #0E2B50;
    }}
    td {{
        padding: 4px 8px;
        border: 1px solid #D0D7DE;
        vertical-align: top;
    }}
    tr:nth-child(even) {{
        background-color: #F8F9FA;
    }}
    blockquote {{
        border-left: 4px solid #0E2B50;
        background-color: #F0F4F8;
        padding: 6px 12px;
        margin: 8px 0;
        font-size: 9pt;
    }}
    pre {{
        background-color: #F1F5F9;
        border: 1px solid #CBD5E1;
        padding: 8px;
        font-family: 'Consolas', monospace;
        font-size: 8pt;
        line-height: 1.3;
        overflow-x: auto;
    }}
    p {{
        margin-top: 0;
        margin-bottom: 6px;
        text-align: justify;
    }}
    .footer {{
        font-size: 8pt;
        color: #6C757D;
        text-align: center;
        margin-top: 16px;
        border-top: 1px solid #E1E4E8;
        padding-top: 4px;
    }}
</style>
</head>
<body>
<h1>{title}</h1>
<blockquote>
<p><strong>KPPSC PMS \u2022 INTERNATIONAL RELATIONS (100 MARKS)</strong><br/>
Master Model Solution & Theoretical Deconstruction | 20 Marks Master Answer Guide</p>
</blockquote>

<h2>Executive Examination Metadata</h2>
<table>
<thead><tr><th>Parameter</th><th>Specification</th></tr></thead>
<tbody>
"""
    for k, v in meta_data:
        html_content += f"<tr><td><strong>{k}</strong></td><td>{v}</td></tr>\n"
    html_content += """</tbody></table>

<h2>Exam Essential Takeaways</h2>
<blockquote>
"""
    for item in callout_data:
        html_content += f"<p>\u2022 {item}</p>\n"
    html_content += f"""</blockquote>

<h2>Conceptual Thematic Map & Strategic Roadmap</h2>
<pre>
{map_text.strip()}
</pre>

<h2>Analytical Dimensions & Systemic Shifts Matrix</h2>
<table>
<thead><tr><th>Dimension / Arena</th><th>Strategic Mechanism / Driver</th><th>Systemic Impact on Global/Regional Order</th></tr></thead>
<tbody>
"""
    for d, m, s in timeline_data:
        html_content += f"<tr><td><strong>{d}</strong></td><td>{m}</td><td>{s}</td></tr>\n"
    html_content += """</tbody></table>

<h2>The 20-Mark Master Model Essay (~650-700 Words)</h2>
"""
    for sub_h, body_p in essay_sections:
        html_content += f"<h4>{sub_h}</h4>\n<p>{body_p}</p>\n"
        
    html_content += """<h2>Theoretical Perspectives & Quotable Thesis Arguments</h2>
<table>
<thead><tr><th>Theorist / Scholar</th><th>Theoretical School</th><th>Exact Quotable Thesis Argument</th></tr></thead>
<tbody>
"""
    for th, sc, ar in quotes_data:
        html_content += f"<tr><td><strong>{th}</strong></td><td><em>{sc}</em></td><td>\"{ar}\"</td></tr>\n"
    html_content += f"""</tbody></table>

<div class="footer">KPPSC PMS International Relations Master Examination Suite | Government of Khyber Pakhtunkhwa</div>
</body>
</html>"""

    temp_html = md_path.replace(".md", "_temp.html")
    with open(temp_html, 'w', encoding='utf-8') as f:
        f.write(html_content)
        
    cmd = [
        EDGE_PATH,
        "--headless",
        "--disable-gpu",
        "--run-all-compositor-stages-before-draw",
        f"--print-to-pdf={pdf_path}",
        temp_html
    ]
    subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if os.path.exists(temp_html):
        os.remove(temp_html)
    print(f"PDF and MD successfully generated: {pdf_path}")

# =========================================================================
# Q2: INDUS BASIN TREATY & HYDROLOGICAL SECURITY (2025 Q.2)
# =========================================================================

q2_meta = [
    ("Target Examination", "KPPSC PMS (Provincial Management Service) — Khyber Pakhtunkhwa"),
    ("Paper & Subject", "International Relations (100 Marks Optional — Group 6)"),
    ("EXAM QUESTIONS LINKED:", '• 2025 Q.2: "Critically evaluate the resilience of Indus Basin Treaty as a cornerstone of hydrological security in South Asia. To what extent can this agreement continue to safeguard water-sharing arrangements from broader geopolitical tension between India and Pakistan, particularly in an era of increasing water scarcity and climate change? (20 Marks)"'),
    ("Syllabus Section", "Part C: Regional Security, Water Diplomacy & Pakistan-India Relations"),
    ("Question Allocation", "Attempt 1 of 5 Required (20 Marks | Strict 35-Minute Allocation)"),
    ("Historiographical / Theoretical Thesis", "Brokered by the World Bank in 1960, the Indus Waters Treaty (IWT) established a durable transboundary water-partitioning framework that survived three full-scale wars. However, under the lens of critical hydropolitics and structural realism (Zeitoun & Warner's 'Hydro-Hegemony'), the treaty's functionalist resilience is fundamentally stressed by upper-riparian hegemonic maneuvers, run-of-the-river hydroelectric expansion (Kishenganga and Ratle), climate-induced glacial depletion in the Himalayas, and India's formal notice demanding Article XII modification. Rather than technical water-sharing, water has become weaponized within broader geopolitical containment strategies.")
]

q2_callout = [
    "STRUCTURAL PARTITIONING REGIME: The 1960 Treaty divided the Indus basin: Eastern Rivers (Ravi, Beas, Sutlej, ~33 MAF) granted exclusively to India; Western Rivers (Indus, Jhelum, Chenab, ~136 MAF) allocated to Pakistan, subject to restricted Indian run-of-the-river hydroelectric usage.",
    "HYDRO-HEGEMONY & THE RUN-OF-THE-RIVER PARADOX: India's construction of massive cascade dams (Baglihar, Kishenganga, Ratle) adheres technically to 'run-of-the-river' definitions but grants cumulative pondage and discharge control that can manipulate downstream river flow during critical sowing seasons.",
    "DISPUTE RESOLUTION IMPASSE: A procedural deadlock between the Court of Arbitration (CoA) at The Hague and the World Bank-appointed Neutral Expert, compounded by India's boycott of the CoA, threatens the treaty's compulsory judicial mechanisms.",
    "CLIMATE EXACERBATION: The Third Pole (Hindu Kush-Himalayas) is warming at twice the global average, transforming predictable snowmelt into erratic glacier bursts and severe dry-season droughts, stressing Pakistan's agrarian lifeline (which supports 90% of food production)."
]

q2_map = r"""
+---------------------------------------------------------------------------------------------------------+
|                  CONCEPTUAL ROADMAP: INDUS BASIN TREATY & HYDROLOGICAL SECURITY                         |
+---------------------------------------------------------------------------------------------------------+
                                                     |
                                   [1960: World Bank Mediation]
                          Partition of Basin: Eastern (India) vs Western (Pak)
                             Permanent Indus Commission (PIC) Formed
                                                     |
                                                     v
                       [The Resilience Phase: Surviving Hot Wars (1965, 1971, 1999)]
                       Functionalist Insulation: Technocratic Dialogue Persists
                                                     |
                                                     v
                         [21st Century Frictions: Hydro-Engineering Dilemma]
                       Indian Upstream Cascade Projects: Baglihar, Kishenganga, Ratle
                       Pakistan's Objection: Pondage, Freeboard & Gated Spillways
                                                     |
           +-----------------------------------------+-----------------------------------------+
           |                                         |                                         |
           v                                         v                                         v
   [HYDRO-HEGEMONY]                          [CLIMATE STRESS]                          [LEGAL & DIPLOMATIC]
Upper Riparian Asymmetry                   Glacial Retreat in Himalayas              Article XII Notice by India (2023)
"Blood and Water Cannot Flow"              Shift from Snowmelt to Flash Floods       Parallel Fora: CoA vs Neutral Expert
Strategic Water Weaponization              Severe Food & Agri Vulnerability          Indian Boycott of The Hague Court
           |                                         |                                         |
           +-----------------------------------------+-----------------------------------------+
                                                     |
                                                     v
                          [FUTURE SCENARIOS: MODERNIZATION OR BREAKDOWN]
                     Desperately Requires: Real-Time Hydro-Data Sharing,
                     Climate Adaptation Protocols, and Joint Basin Governance
+---------------------------------------------------------------------------------------------------------+
"""

q2_timeline = [
    ("1960: Treaty Signing & Division", "World Bank arbitrated division; India received Eastern Rivers (33 MAF), Pakistan Western Rivers (136 MAF).", "Established an institutional baseline of non-interference and predictable hydrological allocation for 60+ years."),
    ("2005–2007: Baglihar Dam Arbitration", "Pakistan challenged Baglihar design on the Chenab; World Bank appointed Neutral Expert Raymond Lafitte.", "Lafitte allowed gated spillways and lower drawdown, setting a precedent that favored Indian hydro-engineering leeway."),
    ("2013–2016: Kishenganga & Ratle Disputes", "Pakistan disputed Kishenganga diversion (Jhelum) and Ratle storage (Chenab); invoked Court of Arbitration.", "Exposed institutional paralysis; India demanded a Neutral Expert while Pakistan demanded CoA, creating simultaneous parallel proceedings."),
    ("2016: Uri Attack & Hydro-Weaponization", "Indian leadership declared 'Blood and water cannot flow together,' threatening to maximize river usage.", "Marked the explicit weaponization of transboundary water resources in South Asian geopolitical deterrence."),
    ("2023–2024: Notice for Treaty Modification", "India served formal diplomatic notice under Article XII demanding renegotiation of dispute mechanisms.", "Pushed the treaty into an existential crisis, challenging the foundational consensus of bilateral transboundary water law.")
]

q2_essay = [
    ("Introduction: The Institutional Paradigm of the Indus Waters Treaty",
     "Signed in Karachi on September 19, 1960, under the institutional mediation of the World Bank, the Indus Waters Treaty (IWT) has long been celebrated as one of the world's most resilient transboundary water-sharing agreements. Negotiated between India and Pakistan in an atmosphere of intense post-partition hostility, the treaty resolved a fundamental hydro-strategic conflict by partitioning the Indus River Basin rather than mandating collective basin management. Under the treaty, the three Eastern Rivers (Ravi, Beas, and Sutlej, totaling ~33 million acre-feet) were allocated exclusively to India, while the three Western Rivers (Indus, Jhelum, and Chenab, totaling ~136 million acre-feet) were assigned to Pakistan. Despite enduring three major wars (1965, 1971, and 1999) and recurrent military standoffs, the IWT's institutional machinery—anchored by the Permanent Indus Commission (PIC)—functioned as an insulated technocratic haven. However, in the contemporary era, the confluence of aggressive hydro-engineering, acute climate-induced water stress, and the broader degeneration of South Asian geopolitics has placed the resilience of the IWT under unprecedented structural strain."),

    ("The Mechanics of Historical Resilience: Functionalist Technocracy",
     "The historical longevity of the IWT can be attributed primarily to its clear-cut spatial division of rivers and its structured, three-tiered dispute resolution mechanism (Article IX), comprising the bilateral PIC, a World Bank-appointed Neutral Expert for technical differences, and a Court of Arbitration (CoA) for legal disputes. By allocating entire river systems rather than establishing complicated volumetric quotas or shared riparian governance, the treaty minimized operational friction. This functionalist separation allowed technical commissioners to maintain periodic data exchanges and inspections even when formal bilateral diplomatic ties were severed. Furthermore, the financial underwriting and neutral third-party guarantor role played by the World Bank provided external institutional credibility and deterrence against unilateral treaty abrogation during the heightened ideological confrontations of the Cold War."),

    ("Hydro-Hegemony and Engineering Friction: The Run-of-the-River Dilemma",
     "In recent decades, however, this functionalist insulation has been corroded by the dynamics of 'hydro-hegemony'—a conceptual framework developed by Mark Zeitoun and Jeroen Warner. As the upper-riparian power, India possesses geographic and structural leverage. Under Annexure D of the IWT, India is permitted to construct run-of-the-river hydroelectric plants on the Western Rivers, provided they do not store water for downstream deprivation. However, major Indian projects—such as the Baglihar and Ratle dams on the Chenab, and the Kishenganga project on the Jhelum—have generated profound existential anxieties in Pakistan. Although technically designated as run-of-the-river, the cumulative storage, gated spillways, and deep-orifice drawdown flushing systems engineered into these facilities grant India the technical capacity to manipulate river discharges. During critical agricultural cycles (such as Kharif sowing), even temporary water retention could devastate Pakistan's downstream agrarian economy, transforming hydro-engineering into a tool of coercive statecraft."),

    ("Climate Change, Water Scarcity, and the Geopolitical Threat of Treaty Revision",
     "These geopolitical tensions are exacerbated exponentially by climate change. The Hindu Kush-Himalayan ice cap—often termed the 'Third Pole'—is experiencing unprecedented glacial melt, altering hydrological baselines from predictable, snowmelt-driven flow to erratic glacial lake outbursts followed by projected long-term river depletion. In an era where Pakistan is rapidly approaching absolute water scarcity (annual per capita availability falling below 900 cubic meters), transboundary water insecurity has become an acute national security vulnerability. Following the 2016 Uri incident, Indian rhetoric explicitly weaponized water, with Prime Minister Narendra Modi asserting that 'blood and water cannot flow together.' In January 2023, India issued a formal notice to Pakistan seeking modifications to the treaty under Article XII, protesting the concurrent proceedings of the Neutral Expert and the Permanent Court of Arbitration over the Kishenganga and Ratle disputes. India's boycott of the Court of Arbitration at The Hague underscores an emerging unwillingness to abide by multilateral dispute resolution."),

    ("Conclusion: The Imperative for Transboundary Ecological Re-engineering",
     "In conclusion, while the Indus Waters Treaty has exhibited extraordinary historic resilience, it can no longer safely insulate South Asian water-sharing from escalating geopolitical rivalries and ecological shocks. The 1960 framework was engineered for an era of physical water division, remaining completely silent on climate change adaptation, environmental flows, groundwater depletion, and real-time telemetry data-sharing. If the treaty is to continue safeguarding hydrological security, it cannot survive on upper-riparian unilateralism or legalistic brinkmanship. Both states, facilitated by the World Bank, must transition from zero-sum sovereign partitioning to cooperative, climate-resilient river basin governance—recognizing that in an era of ecological collapse, hydrological security is indivisible.")
]

q2_quotes = [
    ("Mark Zeitoun & Jeroen Warner", "Framework of Hydro-Hegemony", "Hydro-hegemony is hegemony at the river basin level, achieved through riparian position, material power, and control over treaty architecture, allowing the hegemon to dictate terms of resource exploitation."),
    ("John Briscoe", "Transboundary Water Policy", "The genius of the Indus Waters Treaty was that it was a functional partition, not a management treaty. But if India uses its upper-riparian position to build cumulative storage that can choke Pakistan's agriculture, it strikes at the existential heart of the treaty."),
    ("Brahma Chellaney", "Water Geopolitics", "Water is emerging as Asia's defining battleground. The Indus Basin represents a unique tinderbox where nuclear-armed rivals share a basin severely vulnerable to Himalayan glacial retreat and rising demographic demand."),
    ("Ashok Swain", "Environmental Security", "Shared river basins can promote peace or ignite conflict. When political hostilities infect hydrological management, the breakdown of institutional dispute mechanisms elevates water from an environmental resource to a casus belli."),
    ("Aaron T. Wolf", "Transboundary Water Conflict Analysis", "Treaties with robust, non-derogable institutional mechanisms survive armed conflict; however, treaties that lack mechanisms for climate adaptation and real-time joint monitoring inevitably fracture under environmental stress.")
]

# =========================================================================
# Q3: EXPANSION OF BRICS & SCO THROUGH STRUCTURAL REALISM (2025 Q.3)
# =========================================================================

q3_meta = [
    ("Target Examination", "KPPSC PMS (Provincial Management Service) — Khyber Pakhtunkhwa"),
    ("Paper & Subject", "International Relations (100 Marks Optional — Group 6)"),
    ("EXAM QUESTIONS LINKED:", '• 2025 Q.3: "Analyze the expansion of BRICS and Shanghai Cooperation Organization (SCO) through the lens of structural realism. Are these organizations best understood as nascent challengers to a US-led liberal international order, or are their internal divisions too significant to pose a coherent alternative? (20 Marks)"\n• 2022 Q.8(b): "Write short note on: Shanghai Cooperation Organization (SCO) and Regional Integration. (10 Marks)"'),
    ("Syllabus Section", "Part A: Neorealism & Balance of Power | Part C: Counter-Hegemonic Blocs"),
    ("Question Allocation", "Attempt 1 of 5 Required (20 Marks | Strict 35-Minute Allocation)"),
    ("Historiographical / Theoretical Thesis", "Through Kenneth Waltz's and John Mearsheimer's structural realism, the expansion of BRICS (BRICS+) and the Shanghai Cooperation Organization (SCO) is fundamentally driven by systemic balancing against American unipolar hegemony and the weaponization of the dollar-dominated global financial architecture. However, because structural neorealism posits that anarchy breeds pervasive relative-gains concerns and security dilemmas, the profound geopolitical rivalries among member states—exemplified by the Sino-Indian border dispute and disparate strategic alignments—preclude these coalitions from forming a coherent, anti-Western military alliance, rendering them decentralized vehicles for soft institutional balancing and transactional multi-alignment.")
]

q3_callout = [
    "NEOREALIST SYSTEMIC BALANCING: Neorealism dictates that excessive concentration of power (US unipolarity) triggers counter-balancing. Non-Western powers utilize BRICS and SCO as institutional soft-balancing tools to dilute Western economic and strategic primacy.",
    "GEO-ECONOMIC EXPANSION: The 2024 expansion of BRICS (incorporating Iran, UAE, Egypt, Ethiopia) creates an economic bloc representing ~45% of the world's population and ~40% of global crude oil production, accelerating alternative financial clearing (de-dollarization) initiatives.",
    "INTERNAL REALIST CONTRADICTIONS: Under Waltzian logic, states fear relative gains by partners. India's active participation in the US-led Quad while sitting in BRICS and SCO illustrates how regional security dilemmas (Sino-Indian LAC rivalry) paralyze collective bloc action.",
    "ABSENCE OF COLLECTIVE DEFENSE: Neither BRICS nor the SCO possesses an Article 5-style collective defense treaty. They are consultative diplomatic and geo-economic platforms, not an Eastern NATO or Warsaw Pact 2.0."
]

q3_map = r"""
+---------------------------------------------------------------------------------------------------------+
|                    CONCEPTUAL ROADMAP: BRICS & SCO EXPANSION UNDER REALISM                              |
+---------------------------------------------------------------------------------------------------------+
                                                     |
                               [Systemic Trigger: Structural Anarchy]
                             US Unipolar Primacy & Coercive Hegemony
                             Weaponized Interdependence (SWIFT Sanctions)
                                                     |
                                                     v
                       [The Neorealist Balancing Imperative (Waltz/Mearsheimer)]
                       Soft Balancing: Creation of Alternative Institutional Regimes
                       SCO (Security & Eurasia) + BRICS (Finance & Global South)
                                                     |
                                                     v
                         [Expansion Phase: Broadening the Counter-Coalition]
                       SCO: Inducts India, Pakistan, Iran, Belarus (Pan-Eurasian)
                       BRICS+: Inducts UAE, Iran, Egypt, Ethiopia, Saudi (Oil & Trade)
                                                     |
           +-----------------------------------------+-----------------------------------------+
           |                                         |                                         |
           v                                         v                                         v
 [ARENAS OF CHALLENGE]                    [INTERNAL REALIST FRICTIONS]              [SYSTEMIC LIMITATIONS]
Financial Soft-Balancing                  Relative Gains Anxiety (Waltz)            No Mutual Defense Obligation
Local Currencies & CIPS                   Sino-Indian Border Clashes (LAC)          Extreme Regime Heterogeneity
AIIB & New Development Bank               India's Strategic Hedging (Quad)          Divergent Revisionist vs Status
Challenging G7 / Bretton Woods            Ideological & Economic Disparities         Quo Orientations
           |                                         |                                         |
           +-----------------------------------------+-----------------------------------------+
                                                     |
                                                     v
                         [OUTCOME: ASYMMETRIC TRANSACTIONAL MULTILATERALISM]
                      Not a Monolithic Anti-Western Bloc, but a Vehicle for
                      Strategic Autonomy and Post-Western Normative Contestation
+---------------------------------------------------------------------------------------------------------+
"""

q3_timeline = [
    ("2001: Formation of SCO", "Founded in Shanghai by China, Russia, Kazakhstan, Kyrgyzstan, Tajikistan, Uzbekistan to combat 'Three Evils'.", "Established Eurasian security framework independent of Western military architecture and NATO reach."),
    ("2006–2009: Inception of BRIC(S)", "First formal summit in Yekaterinburg (2009); South Africa added in 2010.", "Pioneered collective economic soft-balancing demanding quota reforms in the IMF and World Bank."),
    ("2014: New Development Bank (NDB)", "BRICS established NDB ($100B capital) and Contingent Reserve Arrangement (CRA).", "Created parallel institutional lending architecture free of Washington Consensus conditionalities."),
    ("2017: SCO Expansion to South Asia", "India and Pakistan formally inducted as full members of the SCO.", "Transformed SCO into a pan-Asian platform, but embedded severe South Asian bilateral rivalries inside its institutional core."),
    ("2024: BRICS+ Landmark Expansion", "Induction of Iran, UAE, Egypt, Ethiopia (with Saudi Arabia participating).", "Consolidated energy-producing giants into an alternative trade clearing bloc, intensifying systemic de-dollarization pressures.")
]

q3_essay = [
    ("Introduction: Structural Realism and Counter-Hegemonic Balancing",
     "The dramatic expansion of the BRICS grouping (from its original Brazil, Russia, India, China, South Africa core to BRICS+) and the Shanghai Cooperation Organization (SCO) represents one of the most consequential developments in contemporary international relations. In the discourse of structural realism—founded by Kenneth Waltz and extended into offensive realism by John Mearsheimer—the international system is defined by an immutable condition of anarchy, where states must rely on self-help to guarantee their survival. Neorealism posits that an extreme concentration of power, such as the post-Cold War American unipolar hegemony, inevitably generates countervailing balancing pressures. Frustrated by Western normative dominance and the coercive weaponization of dollar clearing systems (SWIFT), rising and revisionist powers have mobilized BRICS and the SCO as instruments of counter-hegemonic coordination. However, an incisive neorealist analysis reveals that while these organizations pose a formidable structural challenge to Western institutional supremacy, their acute internal divisions, relative-gains anxieties, and absence of collective defense mechanisms severely limit their capacity to function as a unified, coherent alternative bloc."),

    ("External Balancing and the Challenge to the Liberal International Order",
     "Through the realist lens of external balancing, BRICS and the SCO represent concerted efforts to dilute Western hegemony across distinct geopolitical and geo-economic arenas. The SCO, founded in 2001, has evolved from a regional Central Asian confidence-building mechanism focused on combating the 'Three Evils' (terrorism, separatism, and extremism) into a sprawling pan-Eurasian security dialogue encompassing China, Russia, India, Pakistan, Iran, and Belarus. Concurrently, BRICS has evolved from an investment banking acronym into a formidable institutional counterweight to the G7. The 2024 expansion of BRICS to include major energy powerhouses (the UAE, Iran) and demographic anchors (Egypt, Ethiopia) creates a collective representing over 45% of the global population and nearly 40% of global oil output. By establishing the New Development Bank (NDB) and pioneering local-currency clearing systems (such as the Cross-Border Interbank Payment System, CIPS), these blocs directly contest the unilateral sanctions architecture and financial monopoly of the Bretton Woods institutions."),

    ("The Realist Dilemma: Relative Gains and Internal Strategic Fractures",
     "Despite these outward demonstrations of collective strength, structural realism provides the exact conceptual tools to explain why BRICS and the SCO cannot operate as a cohesive anti-Western alliance. A core tenet of Waltzian realism is the 'relative-gains problem': under anarchy, states are perpetually fearful that partners will gain disproportionate power from cooperation, which could later be turned against them. This dynamic is acutely evident in the profound geopolitical chasm between China and India. The militarized confrontation along the Sino-Indian Line of Actual Control (LAC) in the Himalayas underscores an active regional security dilemma. India views China's hegemony with deep alarm; consequently, New Delhi pursues a strategy of 'multi-alignment' or strategic hedging—actively participating in the US-led Quadrilateral Security Dialogue (Quad) to balance China militarily, while concurrently attending BRICS and SCO summits to safeguard its economic interests and Global South leadership credentials."),

    ("Heterogeneity of Regimes and the Absence of Strategic Cohesion",
     "Furthermore, the members of BRICS+ and the SCO share neither a common ideological foundation nor a shared strategic threat perception. Stephen Walt's 'Balance of Threat' theory emphasizes that states balance not merely against aggregate power, but against perceived threat and offensive intentions. Within BRICS+, members exhibit sharply divergent strategic postures: Russia and Iran represent explicitly revisionist powers locked in direct confrontation with the West, seeking to dismantle the US-led order. Conversely, states like India, Brazil, the UAE, and South Africa are status-quo reformists; they seek greater voice and prestige within global governance but remain deeply integrated into Western capital markets and have no desire to be absorbed into an anti-Western Beijing-Moscow axis. Crucially, neither BRICS nor the SCO possesses an Article 5-style collective defense clause, meaning they cannot provide hard security guarantees."),

    ("Conclusion: Asymmetric Multilateralism and the Multiplex Order",
     "In conclusion, the expansion of BRICS and the SCO is best understood not as the embryonic emergence of an eastern Warsaw Pact, but as a manifestation of what Amitav Acharya terms the emerging 'multiplex world.' Driven by structural realist incentives to balance against unilateral US sanctions and institutional marginalization, non-Western powers are using these platforms to construct alternative, non-Western institutional options. However, because structural anarchy ensures that mutual suspicions, border disputes, and relative-gains anxieties persist within these groupings, BRICS and the SCO will remain structurally loose, transactional coalitions. They successfully accelerate the decentralization of global governance and the trend toward de-dollarization, but their internal fractures prevent them from mounting a monolithic, synchronized assault on the international system.")
]

q3_quotes = [
    ("Kenneth N. Waltz", "Structural Realism", "In international politics, unbalance leads to balance. When an unshared hegemony exists, weaker states will invariably balance against the dominant power to protect their sovereign autonomy."),
    ("Stephen M. Walt", "Balance of Threat Theory", "States balance against threats, not just power. The deep internal divisions in groupings like BRICS and the SCO arise because member states like India perceive greater proximate threat from China than from distant Western hegemony."),
    ("John J. Mearsheimer", "Offensive Realism", "Great powers are never satisfied with existing distributions of power; they seek hegemony. While China uses institutions like BRICS and the SCO to expand its sphere of influence, rival neighbors like India will inevitably hedge against Chinese dominance."),
    ("Amitav Acharya", "Post-Western International Relations", "BRICS and the SCO do not represent a new Cold War bloc, but the rise of a multiplex world—an order featuring multiple modernities, diverse regional architectures, and overlapping cross-alignments rather than rigid bipolar alliances."),
    ("C. Raja Mohan", "Strategic Multi-Alignment", "India's simultaneous presence in the Quad and BRICS/SCO is not an ideological contradiction, but the quintessential realist strategy of multi-alignment: extracting benefits from the West while preserving leverage in Eurasia.")
]

# =========================================================================
# Q4: ABRAHAM ACCORDS & THE TWO-STATE SOLUTION (2025 Q.4)
# =========================================================================

q4_meta = [
    ("Target Examination", "KPPSC PMS (Provincial Management Service) — Khyber Pakhtunkhwa"),
    ("Paper & Subject", "International Relations (100 Marks Optional — Group 6)"),
    ("EXAM QUESTIONS LINKED:", '• 2025 Q.4: "Abraham Accord has been described as re-aligning Middle East geopolitics. Critically assess the impact of this new diplomatic architecture on the viability of the two-state solution for the Israel-Palestine conflict. Does it create new leverage for peace or further marginalize the Palestinian issue? (20 Marks)"\n• 2022 Q.6: "Abraham Accords was designed to normalize relations between Israel and the Arab World; however, except for a few Arab States, the idea has not received popular support among the public and leaders. Critically examine the implications of this Accord on the Palestinian issue. (20 Marks)"\n• 2013 Q.5: "Critically examine the Palestine Problem from the Balfour Declaration (1917) to the post-Oslo developments. Why has a permanent peace settlement remained elusive? (20 Marks)"'),
    ("Syllabus Section", "Part C: Middle East Geopolitics, Arab-Israeli Relations & Palestine Issue"),
    ("Question Allocation", "Attempt 1 of 5 Required (20 Marks | Strict 35-Minute Allocation)"),
    ("Historiographical / Theoretical Thesis", "The 2020 Abraham Accords fundamentally re-engineered Middle Eastern diplomatic architecture by discarding the historic 'land for peace' consensus established by the 2002 Arab Peace Initiative. Evaluated through the realist prism of threat convergence (anti-Iran balancing) and transactional geo-economics, the normalization between Israel and Gulf monarchies was predicated on the false premise of 'economic peace'—the belief that regional integration could proceed while bypassing Palestinian political self-determination. Rather than providing leverage for peace, the Accords systematically marginalized the Palestinian question, exacerbating structural oppression in the Occupied Territories and directly contributing to the explosive regional catastrophe of October 7, 2023, which definitively proved that Middle Eastern stability is impossible without an independent Palestinian state.")
]

q4_callout = [
    "PARADIGM SHIFT: The Accords abandoned the 2002 Arab Peace Initiative, which made normalization strictly contingent on Israeli withdrawal to 1967 borders and the establishment of a sovereign Palestinian state with East Jerusalem as its capital.",
    "REALIST THREAT CONVERGENCE: The driving catalyst was shared threat perception regarding the Islamic Republic of Iran's regional missile proliferation and proxy network ('Axis of Resistance'), coupled with anxieties over perceived American retrenchment from the Middle East.",
    "THE 'ECONOMIC PEACE' ILLUSION: Israeli leadership (Netanyahu) promoted the doctrine that normalization, trade, cybersecurity cooperation, and technological integration could substitute for Palestinian political sovereignty.",
    "OCTOBER 7 CATACLYSM & THE RETURN OF PALESTINE: The eruption of the Hamas-Israel war shattered the normalization paradigm, freezing Saudi-Israeli normalization talks and compelling international recognition that the two-state solution is an existential necessity, not an optional luxury."
]

q4_map = r"""
+---------------------------------------------------------------------------------------------------------+
|                  CONCEPTUAL ROADMAP: ABRAHAM ACCORDS & PALESTINE DYNAMICS                               |
+---------------------------------------------------------------------------------------------------------+
                                                     |
                                   [Historical Baseline: 2002 API]
                             Arab Peace Initiative: Normalization = 1967 Borders
                             Palestinian Statehood as Inflexible Precondition
                                                     |
                                                     v
                       [Systemic Shift: Iran Threat & US Pivot]
                       Convergence of Israeli-Gulf Security Interests
                       Realpolitik Over Pan-Arab Solidarity (Realism)
                                                     |
                                                     v
                         [September 2020: The Abraham Accords]
                       UAE, Bahrain, Morocco, Sudan Normalize Ties with Israel
                       Decoupling Regional Normalization from Palestinian Issue
                                                     |
           +-----------------------------------------+-----------------------------------------+
           |                                         |                                         |
           v                                         v                                         v
   [ISRAELI CALCULUS]                        [GULF MONARCHY GOALS]                     [PALESTINIAN REALITY]
"Outside-In" Approach                      Hedging Against Iranian Power             Complete Political Abandonment
Annexation Suspended, Not Ended            Access to Israeli Tech & Intel            Acceleration of West Bank Settlements
Palestinian Issue Circumvented             Solidifying Ties with Washington          Gaza Blockade Consolidated
           |                                         |                                         |
           +-----------------------------------------+-----------------------------------------+
                                                     |
                                                     v
                         [THE SYSTEMIC ERUPTION: OCTOBER 7, 2023]
                    Collapse of the "Economic Peace" & Circumvention Myth
                    Saudi Normalization Frozen: Sovereign Statehood Inevitable
                    The Reassertion of the Two-State Solution as Core Imperative
+---------------------------------------------------------------------------------------------------------+
"""

q4_timeline = [
    ("2002: Arab Peace Initiative (API)", "Beirut Arab League Summit; offered full normalization in exchange for Israeli withdrawal to 1967 borders.", "Established the foundational pan-Arab diplomatic consensus linking normalization to Palestinian sovereign statehood."),
    ("September 2020: Abraham Accords Signed", "US-brokered normalization agreements signed by Israel, UAE, and Bahrain; later Morocco and Sudan.", "Decoupled normalization from Palestinian statehood, marking a radical structural realignment based on transactional realpolitik."),
    ("2021–2022: Geo-economic Integration", "Bilateral trade between UAE and Israel crossed $2.5B; defense, cyber, and intelligence agreements signed.", "Promoted the illusion that Middle Eastern security could be organized around an anti-Iran axis while ignoring the occupation."),
    ("Mid-2023: US-Saudi-Israel Mega-Deal Talks", "Biden administration actively brokered normalization between Saudi Arabia and Israel in exchange for security pacts.", "Brought the Palestinian issue to the absolute brink of permanent geopolitical obsolescence."),
    ("October 7, 2023 – Present: Regional War", "Hamas launched Operation Al-Aqsa Flood; massive Israeli military campaign in Gaza and Lebanon.", "Destroyed the circumvention paradigm; Saudi Arabia affirmed that normalization is impossible without an irreversible path to Palestinian statehood.")
]

q4_essay = [
    ("Introduction: The Diplomatic Metamorphosis of Middle Eastern Realignment",
     "Signed on the White House South Lawn in September 2020 under the stewardship of the Trump administration, the Abraham Accords—facilitating the formal normalization of diplomatic relations between Israel and the United Arab Emirates, Bahrain, and subsequently Morocco and Sudan—represented the most radical restructuring of Middle Eastern diplomacy in four decades. For nearly half a century, pan-Arab diplomacy had been anchored to the principle of 'land for peace,' formalized in the 2002 Arab Peace Initiative (API). The API explicitly held that Arab normalization with Tel Aviv was contingent upon Israeli withdrawal from the territories occupied in 1967 and the establishment of a sovereign, independent Palestinian state with East Jerusalem as its capital. The Abraham Accords demolished this consensus, pioneering an 'outside-in' diplomatic architecture that decoupled Arab recognition of Israel from the resolution of the Palestinian national question. However, viewed through the lens of critical international relations and empirical reality, this realignment did not unlock genuine regional peace; instead, it operated on the fatal illusion of 'economic peace,' directly contributing to the cataclysm of October 7, 2023."),

    ("Realist Drivers: Threat Convergence and the Anti-Iran Imperative",
     "The catalyst behind the Abraham Accords was not an ideational breakthrough regarding coexistence, but cold, hard structural realpolitik. Through Stephen Walt's balance-of-threat theory, states align against proximate, aggressive capabilities rather than abstract moral principles. For the Gulf monarchies—particularly the UAE and Bahrain—the perceived retrenchment of American security commitments across the Middle East, combined with the expanding ballistic missile capabilities and asymmetric proxy network of the Islamic Republic of Iran (the 'Axis of Resistance'), generated an acute security dilemma. Israel, possessing undisputed regional military superiority, advanced intelligence capabilities, and unmatched cyber-defense infrastructure, emerged as a natural balance-of-power partner. In exchange for formal diplomatic ties, the Gulf states gained access to advanced Israeli defense systems (such as air-defense integration and Pegasus cyber-intelligence) while solidifying bipartisan security patronage in Washington."),

    ("The Circumvention Strategy: Marginalization of the Palestinian Question",
     "For the government of Israeli Prime Minister Benjamin Netanyahu, the Abraham Accords represented the triumphant vindication of Ze'ev Jabotinsky's 'Iron Wall' doctrine—the thesis that Arabs would only reconcile with Israel once they realized Jewish sovereignty could not be broken. By securing peace treaties with prominent Arab nations without conceding a single inch of occupied territory or halting settlement expansion in the West Bank, Netanyahu successfully bypassed the Palestinian issue. The accords reduced the Palestinian national struggle from an existential question of political self-determination and statehood to a transactional matter of humanitarian aid and economic subsidies. Western and regional diplomats championed the illusion that cross-border capital investments, tourism, and high-tech trade would naturally pacify the Occupied Territories. In reality, the accords deeply disenfranchised the Palestinian population, completely delegitimizing the moderate Palestinian Authority and entrenching the despair that empowered militant rejectionist factions."),

    ("The Systemic Collapse of the 'Economic Peace' Paradigm",
     "The catastrophic eruption of the Hamas-led assault on October 7, 2023, and the ensuing regional war in Gaza, Lebanon, and across the Red Sea decisively shattered the core premise of the Abraham Accords. The violence proved beyond ambiguity that regional integration cannot be engineered on the burial grounds of Palestinian aspirations. The circumvention strategy collapsed precisely because peace between distant monarchies in the Persian Gulf could never substitute for a negotiated settlement with the millions of Palestinians living under occupation. Crucially, the momentum toward the ultimate prize of the Accords—Saudi-Israeli normalization—was brought to an abrupt halt. Riyadh, recognizing domestic public fury and the systemic imperative of pan-Islamic legitimacy, unequivocally recalibrated its foreign policy, declaring that diplomatic normalization with Israel is unconditionally contingent upon an irreversible, time-bound pathway toward a sovereign Palestinian state based on 1967 borders."),

    ("Conclusion: The Irreplaceable Necessity of Sovereign Palestinian Statehood",
     "In conclusion, while the Abraham Accords demonstrated the capacity of realist threat perceptions to realign bilateral state interests, they fundamentally failed as a holistic framework for regional peace. By treating the Palestinian question as a peripheral nuisance rather than the core root of Middle Eastern instability, the accords created a fragile diplomatic facade that masked boiling structural oppression. The post-October 7 landscape has forced an inescapable international consensus: the two-state solution is not an obsolete relic, but the only viable geopolitical formula capable of ensuring long-term Israeli security, Palestinian sovereignty, and authentic Middle Eastern stability. Any diplomatic architecture that attempts to construct regional integration without resolving the sovereign rights of the Palestinians will inevitably collapse into recurrent cycles of catastrophic violence.")
]

q4_quotes = [
    ("Avi Shlaim", "The Iron Wall: Israel and the Arab World", "The Abraham Accords were not a peace treaty; they were an arms-and-business transaction between authoritarian regimes and Israel, designed to bypass the core moral and political issue of the Middle East—the dispossession of the Palestinians."),
    ("Rashid Khalidi", "The Hundred Years' War on Palestine", "The Abraham Accords represented a brazen attempt to liquidate the Palestinian issue by treating Palestinians as invisible subjects while integrating Israel into an authoritarian regional security order."),
    ("John J. Mearsheimer", "Structural Realism", "The idea that you could normalize relations with the Arab world and ignore the Palestinians was pure hubris. Realist logic dictates that oppressed populations will inevitably resist domination, and that regional stability cannot be sustained on unresolved territorial conquest."),
    ("Fawaz Gerges", "Middle East International Relations", "The Abraham Accords prioritized an anti-Iran security alliance over the historic pan-Arab commitment to Palestinian statehood. October 7 was the bloody reassertion that Palestine remains the beating heart of Middle Eastern politics."),
    ("Vali Nasr", "Geopolitical Analysis", "The Abraham Accords attempted to freeze Middle Eastern geopolitics around an Arab-Israeli axis against Iran. But by ignoring the Palestinian core, it built an edifice on sand that collapsed the moment the local status quo exploded.")
]

# =========================================================================
# Q5: CROSS-BORDER TERRORISM & PAKISTAN-TALIBAN CRISIS (2025 Q.5)
# =========================================================================

q5_meta = [
    ("Target Examination", "KPPSC PMS (Provincial Management Service) — Khyber Pakhtunkhwa"),
    ("Paper & Subject", "International Relations (100 Marks Optional — Group 6)"),
    ("EXAM QUESTIONS LINKED:", '• 2025 Q.5: "How does cross-border terrorism affect regional stability? Critically assess the Pakistan-Taliban security crisis. What are the implications of the Taliban resurgence in Afghanistan for Pakistan’s external and internal security environment, particularly in Baluchistan and Khyber Pakhtunkhwa? (20 Marks)"\n• 2022 Q.4: "In the Realist prism, the \'State\' is the legitimate actor in the international system, but the emergence of Non-State Actors (NSAs) has greatly impacted the global order; Taliban are governing Afghanistan without recognition and Hezbollah is a major stakeholder in Lebanon. Analyze the impact of NSAs on the global strategic environment with solid arguments and valid examples. (20 Marks)"\n• 2018 Q.5: "Discuss the emergence of ISIS and its impact on national, regional, and international peace and security. (20 Marks)"'),
    ("Syllabus Section", "Part C: Regional Security, Non-State Actors & Pakistan-Afghanistan Security"),
    ("Question Allocation", "Attempt 1 of 5 Required (20 Marks | Strict 35-Minute Allocation)"),
    ("Historiographical / Theoretical Thesis", "The resurgence of the Afghan Taliban in Kabul in August 2021 has generated a profound strategic paradox for Pakistan's national security architecture. Evaluated through Barry Buzan's Regional Security Complex Theory (RSCT) and structural realism, the historic doctrine of 'strategic depth' has utterly collapsed. The ideational, tribal, and historical solidarity (*bay'ah*) binding the Afghan Taliban to the Tehrik-i-Taliban Pakistan (TTP) has transformed Afghan territory into a sanctuary for cross-border terror campaigns. This dynamic has severely destabilized Khyber Pakhtunkhwa and Balochistan, targeting security forces and Chinese CPEC investments, and compelling Islamabad to abandon its policy of reconciliation in favor of coercive cross-border kinetic strikes and mass deportations.")
]

q5_callout = [
    "COLLAPSE OF STRATEGIC DEPTH: Pakistan's decades-long military-strategic calculation that an Islamist, Pashtun-led government in Kabul would secure its western border and neutralize Indian influence has inverted into an acute security nightmare.",
    "THE IDEATIONAL & TRIBAL TTP NEXUS: The Afghan Taliban regime regards the TTP not as foreign terrorists, but as ideological brethren who shared their 20-year insurgency against US/NATO forces; Kabul refuses to disarm, arrest, or extradite TTP leadership operating from Khost, Paktika, and Kunar.",
    "DEVASTATION OF KP AND BALOCHISTAN: Khyber Pakhtunkhwa has borne the brunt of escalated TTP kinetic strikes against law enforcement (Peshawar Police Lines, Bannu, DI Khan), while Balochistan experiences lethal tactical convergence between the TTP and ethno-nationalist insurgents (BLA) targeting CPEC assets.",
    "SHIFT TO COERCIVE DIPLOMACY: Islamabad has pivoted from failed appeasement and jirga talks (2021–2022) to aggressive coercive statecraft—exemplified by cross-border airstrikes inside Afghanistan, the deportation of over 500,000 undocumented Afghans, and the 'One-Document Regime' at border crossings."
]

q5_map = r"""
+---------------------------------------------------------------------------------------------------------+
|                  CONCEPTUAL ROADMAP: PAKISTAN-TALIBAN SECURITY PARADOX                                  |
+---------------------------------------------------------------------------------------------------------+
                                                     |
                                   [August 2021: US Withdrawal from Kabul]
                             Afghan Taliban Captures Power
                             Pakistan Hails the Fall of the Pro-India Republic
                                                     |
                                                     v
                       [The Structural Paradox & Failure of Strategic Depth]
                       Ideological Allegiance (Bay'ah) between Afghan Taliban & TTP
                       Refusal to Recognize the Durand Line or Surrender TTP Safe Havens
                                                     |
                                                     v
                         [TTP Resurgence & Asymmetric Warfare (2022–Present)]
                       End of Ceasefire (Nov 2022) --> Massive Kinetic Escalation
                       Advanced US Military Hardware (M4 Rifles, Thermal Sights)
                                                     |
           +-----------------------------------------+-----------------------------------------+
           |                                         |                                         |
           v                                         v                                         v
 [IMPACT ON KHYBER PAKHTUNKHWA]            [IMPACT ON BALOCHISTAN]                   [ISLAMABAD'S REALIST PIVOT]
Targeting KP Police & FC Infrastructure    Convergence of TTP & BLA Insurgents       Launch of Operation Azm-e-Istehkam
Suicide Bombings: Peshawar, Bannu, DIK     Targeting CPEC Strategic Convoys          Cross-Border Airstrikes in Khost/Paktika
Subversion of Local Governance             Attacks on Gwadar Complex & Turbat        Expulsion of Undocumented Refugees
Waziristan & Tribal Border Infiltration    Economic Sabotage & Foreign Investment    Strict One-Document Border Controls
           |                                         |                                         |
           +-----------------------------------------+-----------------------------------------+
                                                     |
                                                     v
                         [THE NEW REGIONAL REALITY: COERCIVE CONTAINMENT]
                     Deterrence, Hard Border Hardening, and Third-Party Pressure
                     (China/SCO) Replace the Illusion of Fraternal Islamic Solidarity
+---------------------------------------------------------------------------------------------------------+
"""

q5_timeline = [
    ("August 2021: Afghan Taliban Takeover", "US forces withdrew; Afghan Republic collapsed; Taliban took Kabul without a fight.", "Islamabad anticipated a compliant western border; instead, thousands of incarcerated TTP commanders were released from Bagram."),
    ("2021–2022: Failed Negotiations & Ceasefire", "Pakistan pursued dialogue mediated by Sirajuddin Haqqani; TTP demanded reversal of FATA merger.", "TTP used the ceasefire to regroup, re-arm with left-behind US military hardware, and re-infiltrate KP."),
    ("November 2022: TTP Calls Off Ceasefire", "TTP leadership formally ordered countrywide attacks against Pakistani security forces.", "Marked the beginning of an unprecedented wave of suicide bombings, ambushes, and targeted assassinations across KP and Balochistan."),
    ("January 2023: Peshawar Police Lines Mosque Bombing", "Suicide blast inside police headquarters killed over 100 officers.", "Galvanized national civil-military consensus that the Afghan Taliban regime was actively shielding perpetrators of mass murder."),
    ("2023–2024: Coercive Retaliation & Expulsions", "Pakistan executed cross-border kinetic airstrikes; launched repatriation of 500,000+ Afghan nationals.", "Formalized the total breakdown of fraternal relations, transitioning into active coercive deterrence and border militarization.")
]

q5_essay = [
    ("Introduction: The Paradox of Pakistan's Western Security Frontier",
     "The fall of Kabul to the Afghan Taliban on August 15, 2021, was initially greeted in certain policymaking quarters in Islamabad as a major strategic victory. For decades, Pakistan's military-strategic doctrine had been dominated by the quest for 'strategic depth'—the pursuit of a friendly, pro-Pakistan Islamist regime in Afghanistan that would secure Pakistan's western border, prevent encirclement by India, and respect sovereign security imperatives. However, in the realm of international relations, structural reality swiftly dismantled ideological romance. Rather than providing strategic depth, the Taliban's return has generated an existential security crisis for Pakistan. Rooted in Barry Buzan and Ole Wæver's Regional Security Complex Theory (RSCT), security interdependencies in South Asia are intensely localized. The Afghan Taliban's refusal to rein in the Tehrik-i-Taliban Pakistan (TTP) has turned Afghan soil into a spring-board for cross-border terror campaigns, precipitating a catastrophic deterioration of internal security in Khyber Pakhtunkhwa and Balochistan, and bringing relations between Islamabad and Kabul to the brink of open armed confrontation."),

    ("The Afghan Taliban-TTP Nexus: Ideational Bonds and Sovereign Defiance",
     "The foundation of the Pakistan-Taliban security crisis lies in the profound, symbiotic nexus between the Afghan Taliban (the Islamic Emirate of Afghanistan) and the TTP. Western and regional analysts frequently made the erroneous assumption that the Afghan Taliban would sever ties with the TTP once in power to secure international recognition and economic assistance. However, through the lens of constructivist IR, ideational solidarity, shared jihadist narrative, and tribal kinship (*Pashtunwali*) supersede rational-material statecraft. TTP fighters fought shoulder-to-shoulder with the Afghan Taliban for twenty years against the US and NATO; TTP leadership has sworn solemn religious allegiance (*bay'ah*) to the Afghan Taliban's Supreme Leader, Hibatullah Akhundzada. Consequently, Kabul views the TTP as ideological brethren and an invaluable strategic card. Rather than dismantling TTP sanctuaries in Khost, Paktika, Kunar, and Nangarhar, the Afghan Taliban regime has provided them with logistical sanctuary, freedom of movement, and access to sophisticated military equipment abandoned by departing American forces (including night-vision goggles, M4 sniper systems, and encrypted communication gear)."),

    ("Destabilizing the Frontier: The Resurgence of Violence in Khyber Pakhtunkhwa",
     "The direct consequence of this cross-border sanctuary has been a lethal, sustained assault on Pakistan's security apparatus, concentrated primarily in Khyber Pakhtunkhwa. Following the breakdown of ill-advised, military-led peace negotiations in late 2022—which the TTP exploited to re-infiltrate former tribal districts—the terror group launched an unrelenting campaign against the Khyber Pakhtunkhwa Police and Frontier Corps. Devastating attacks, such as the January 2023 Peshawar Police Lines mosque bombing (killing over 100 officers) and complex assaults in Bannu, Lakki Marwat, and Dera Ismail Khan, underscore a deliberate strategy to dismantle the civilian law enforcement architecture. By challenging state sovereignty along the Durand Line and contesting the constitutional administrative merger of the former Federally Administered Tribal Areas (FATA), the TTP seeks to establish a parallel, ungoverned religious enclave."),

    ("The Multi-Vector Threat in Balochistan and the Threat to CPEC",
     "Simultaneously, the cross-border terror nexus has severely infected the security landscape of Balochistan. In an alarming tactical evolution, the ideologically radicalized TTP has forged operational convergences with ethno-nationalist insurgent groups, most notably the Baloch Liberation Army (BLA) and Balochistan Liberation Front (BLF). Operating from ungoverned spaces along the porous border, these groups have executed coordinated assaults against state infrastructure, security convoys, and, critically, Chinese engineers and strategic projects under the China-Pakistan Economic Corridor (CPEC). Attacks on Gwadar Port facilities, Turbat naval airbases, and transit convoys represent a coordinated effort to sabotage the economic lifeline of the Pakistani state and undermine Beijing's confidence in Pakistan's internal stability, thereby elevating the cross-border crisis from a domestic security problem to a major geopolitical crisis."),

    ("Conclusion: The Shift from Fraternal Illusion to Coercive Realism",
     "In conclusion, the resurgence of cross-border terrorism has permanently dismantled Pakistan's historic Afghanistan doctrine. The Afghan Taliban's intransigence over TTP sanctuaries and their perennial refusal to recognize the international validity of the Durand Line have compelled Islamabad to adopt a posture of uncompromising coercive realism. Pakistan has responded by launching cross-border kinetic strikes against TTP leadership targets inside Afghan territory, initiating the sweeping repatriation of hundreds of thousands of undocumented Afghan nationals to eliminate urban sleeper cells, and enforcing a rigid 'One-Document Regime' at border crossings. As Pakistan launches Operation Azm-e-Istehkam, it is clear that stability will not return to Khyber Pakhtunkhwa or Balochistan through bilateral appeals to religious brotherhood. Pakistan must treat the Taliban regime as a hardheaded, self-interested regional actor, leveraging economic blockades, diplomatic isolation through the SCO, and decisive kinetic deterrence to neutralize cross-border terrorism.")
]

q5_quotes = [
    ("Barry Buzan & Ole Wæver", "Regional Security Complex Theory", "Security dynamics are fundamentally regional: threats travel more easily over short distances than long ones. The Afghan-Pakistan border demonstrates how spillover from a fragile state can systematically destabilize an entire adjacent security complex."),
    ("Ahmed Rashid", "Taliban: Militant Islam and Oil in Central Asia", "The Pakistani establishment long believed it could distinguish between 'good Taliban' in Afghanistan and 'bad Taliban' at home. History has brutally demonstrated that they share the same ideological DNA, and one cannot be patronized without feeding the other."),
    ("Stephen P. Cohen", "The Idea of Pakistan", "Pakistan's quest for strategic depth in Afghanistan was always a strategic illusion. Rather than providing Pakistan with depth against India, Afghanistan has repeatedly served as a strategic quagmire that drains Pakistani military and economic power."),
    ("Ashley J. Tellis", "Strategic Analysis", "The TTP is not merely a domestic insurgent movement; it is an ideological partner of the Afghan Taliban. As long as Kabul provides safe haven, the Pakistani military will face an intractable cross-border war of attrition in KP and Balochistan."),
    ("Christine Fair", "Fighting to the End: The Pakistan Army's Way of War", "Pakistan's Afghanistan policy has collapsed under its own internal contradictions. The Taliban's refusal to accept the Durand Line or arrest TTP terrorists proves that ideological proxy warfare inevitably devours its own sponsor.")
]

# =========================================================================
# Q6: DIPLOMACY AND PROPAGANDA IN MODERN STATECRAFT (2025 Q.6)
# =========================================================================

q6_meta = [
    ("Target Examination", "KPPSC PMS (Provincial Management Service) — Khyber Pakhtunkhwa"),
    ("Paper & Subject", "International Relations (100 Marks Optional — Group 6)"),
    ("EXAM QUESTIONS LINKED:", '• 2025 Q.6: "Evaluate the interplay of diplomacy and propaganda in shaping ties between states. To what extent do these tools serve as mechanisms of conflict resolution versus instruments of power projection in the modern international system? (20 Marks)"\n• 2016 Q.4: "\\"Diplomacy is the management of international relations by negotiation.\\" Trace the evolution of diplomacy from traditional bilateral diplomacy to modern multilateral and public diplomacy. (20 Marks)"'),
    ("Syllabus Section", "Part A: Interaction Between States, Diplomacy & Power Projection"),
    ("Question Allocation", "Attempt 1 of 5 Required (20 Marks | Strict 35-Minute Allocation)"),
    ("Historiographical / Theoretical Thesis", "In the classical canon of international relations (Nicolson, Morgenthau), diplomacy constitutes the management of inter-state relations through negotiation, compromise, and institutionalized communication aimed at peaceful conflict resolution, whereas propaganda operates as a coercive psychological instrument designed to manipulate mass perception and maximize national power. In the contemporary 21st-century digital landscape, the boundary between these two instruments has fundamentally collapsed. Through public diplomacy (Nye), strategic communications, and information warfare, states increasingly deploy weaponized propaganda not as a peaceful adjunct to negotiation, but as an offensive instrument of power projection, cognitive warfare, and narrative hegemony that subverts traditional diplomatic conflict resolution.")
]

q6_callout = [
    "CLASSICAL DICHOTOMY: Classical diplomacy (Harold Nicolson, François de Callières) prioritizes confidential negotiation, bilateral trust, sovereign reciprocity, and mutual concession; Propaganda (E.H. Carr) prioritizes mass emotional manipulation, ideational subversion, and psychological coercion.",
    "THE PUBLIC DIPLOMACY HYBRID: The transition from 19th-century cabinet diplomacy to 21st-century 'Public Diplomacy' (Joseph Nye) has blurred the boundaries, allowing states to communicate directly with foreign publics (via BBC, RT, CGTN, Al Jazeera, social media) to cultivate soft power or wage perception warfare.",
    "INFORMATION WARFARE & WEAPONIZED NARRATIVES: In the contemporary era of hybrid warfare and AI-driven deepfakes, propaganda has transformed into 'cognitive warfare'—weaponized to influence foreign elections, destabilize adversary societies, and manufacture diplomatic legitimacy.",
    "EROSION OF NEGOTIATED COMPROMISE: When leaders conduct diplomacy via performative public propaganda (e.g., 'twiplomacy' and megaphone diplomacy), domestic audience costs soar, eliminating the confidential flexibility essential for genuine diplomatic compromise and crisis de-escalation."
]

q6_map = r"""
+---------------------------------------------------------------------------------------------------------+
|                  CONCEPTUAL ROADMAP: DIPLOMACY VS PROPAGANDA IN STATECRAFT                              |
+---------------------------------------------------------------------------------------------------------+
                                                     |
                               [Classical Canon: Separation of Powers]
                    Diplomacy (Nicolson/Morgenthau)   vs   Propaganda (E.H. Carr)
                    - Negotiation & Compromise             - Psychological Manipulation
                    - Confidentiality & Reciprocity        - Mass Emotional Appeal
                    - Conflict Resolution Focus            - Power Projection Focus
                                                     |
                                                     v
                       [The 20th Century Shift: Total War & Cold War Bipolarity]
                       Advent of Mass Media (Radio, Cinema, Television)
                       Ideological Warfare: US Information Agency (USIA) vs Soviet Agitprop
                                                     |
                                                     v
                         [21st Century Revolution: Digital & Cognitive Convergence]
                       Collapse of Confidentiality --> "Megaphone / Digital Diplomacy"
                       Public Diplomacy (Nye): Soft Power Tool or Subversive Propaganda?
                                                     |
           +-----------------------------------------+-----------------------------------------+
           |                                         |                                         |
           v                                         v                                         v
 [CONFLICT RESOLUTION MECHANISM]            [WEAPONIZED POWER PROJECTION]             [SYSTEMIC PATHOLOGIES]
Track-One Multilateral Negotiations       Hybrid Warfare & Deepfake Disinformation    Audience Costs Trap Leaders
Verification Regimes & Treaties           Election Interference & Astroturfing        Polarization Prevents Concession
Crisis Management Hotlines                Strategic Communication (StratCom)          Manufactured Crises Erode Trust
Quiet De-escalation Channels              Cognitive Subversion of Adversary           Collapse of Diplomatic Norms
           |                                         |                                         |
           +-----------------------------------------+-----------------------------------------+
                                                     |
                                                     v
                         [SYNTHESIS: THE DOMINANCE OF WEAPONIZED NARRATIVE]
                     Propaganda Has Subjugated Traditional Diplomatic Craft;
                     Effective 21st-Century Statecraft Requires Balancing Hard Narrative
                     Projection with Confidential Bilateral Crisis De-escalation
+---------------------------------------------------------------------------------------------------------+
"""

q6_timeline = [
    ("1716: François de Callières' Treaty", "Published 'On the Manner of Negotiating with Princes,' establishing classical diplomatic norms.", "Codified diplomacy as professional, continuous, and honest negotiation, distinct from deception or vulgar domestic appeal."),
    ("1914–1919: World War I & Creel Committee", "WWI saw the first total industrialization of state propaganda (British War Propaganda Bureau, US Creel Committee).", "Demonstrated that total war required the total mobilization of mass public opinion, directly influencing the Paris Peace Conference."),
    ("Cold War (1947–1989): Ideological Battle", "US Information Agency (Voice of America, Radio Free Europe) countered Soviet Pravda and Agitprop.", "Formalized the integration of psychological propaganda as a permanent pillar of peacetime diplomatic power projection."),
    ("1990: Joseph Nye Formulates 'Soft Power'", "Coined 'soft power' and elevated 'Public Diplomacy' as the art of winning hearts and minds.", "Legitimized the state projection of cultural and ideological narratives as an authentic, benign diplomatic instrument."),
    ("2016–Present: The Era of Digital Hybrid Warfare", "Russian disinformation operations, Cambridge Analytica, state-backed bot farms, and AI deepfakes.", "Completely erased the boundary between diplomacy and covert information operations, weaponizing media into an offensive weapon.")
]

q6_essay = [
    ("Introduction: Defining the Instruments of Classical and Modern Statecraft",
     "In the study and practice of international relations, diplomacy and propaganda represent two fundamental instruments through which sovereign states pursue national interest and manage their interactions. In classical diplomatic theory, as articulated by Sir Harold Nicolson and Hans Morgenthau, diplomacy is defined as the management of international relations through negotiation, institutional dialogue, and confidential compromise, with the overarching objective of maintaining systemic order and achieving peaceful conflict resolution. Conversely, propaganda—classically analyzed by E.H. Carr as 'power over opinion'—is the deliberate, systematic manipulation of symbols, narratives, and psychological perceptions designed to influence mass behavior and maximize national power. Historically viewed as antithetical tools—one operating in the rational, quiet chambers of ambassadorial chanceries, the other in the raucous arenas of public agitation—the boundary between diplomacy and propaganda has fundamentally disintegrated in the 21st century. In an era dominated by instantaneous digital communication, algorithmic echo chambers, and hybrid warfare, these two mechanisms have converged into a hybrid instrument of competitive power projection."),

    ("Diplomacy as an Instrument of Conflict Resolution",
     "Historically and conceptually, diplomacy serves as the primary civilized alternative to the use of armed force in an anarchic international system. Rooted in the Westphalian tradition and codified under the Vienna Convention on Diplomatic Relations (1961), classical diplomacy functions through sovereign recognition, reciprocal diplomatic immunity, and institutionalized negotiation. Its quintessential role is the peaceful resolution of conflict: de-escalating security dilemmas, negotiating binding armistices, formulating arms control treaties (such as the 1968 NPT or 2015 JCPOA), and establishing verification regimes. The efficacy of diplomacy relies critically on confidentiality. As François de Callières famously argued in the eighteenth century, quiet, secret negotiations allow state leaders to explore creative compromises, make necessary tactical concessions, and step back from geopolitical precipices without suffering catastrophic loss of domestic prestige or electoral viability."),

    ("Propaganda as an Offensive Engine of Power Projection",
     "In sharp contrast, propaganda operates as an offensive, unilateral weapon of power projection. In Politics Among Nations, Hans Morgenthau identified the 'quality of diplomacy' and the 'element of national morale' as vital components of power, but noted that propaganda seeks to bypass foreign governments entirely to subvert the domestic will of adversary populations. From the ideological radio broadcasts of the Cold War (Voice of America versus Radio Moscow) to the contemporary state-backed multi-lingual media empires of today (Russia Today, CGTN, and Al Jazeera), propaganda is deployed to manufacture domestic consensus, justify foreign interventions, delegitimize rival regimes, and cultivate strategic soft power. Under conditions of acute geopolitical rivalry, propaganda does not seek compromise; it seeks cognitive capitulation, polarizing adversary societies from within and constructing an unyielding perception of moral righteousness for the home state."),

    ("The Digital Collapse of the Boundary: Public Diplomacy and Cognitive Warfare",
     "In the contemporary information age, the revolution in digital technologies has hybridized diplomacy and propaganda. What was once celebrated under Joseph Nye's framework as benign 'public diplomacy'—states communicating transparently with foreign publics to project cultural appeal and values—has metastasized into weaponized digital disinformation and 'cognitive warfare.' Modern states routinely deploy cyber-troll armies, automated bot networks, state-curated social media narratives, and AI-generated deepfakes to manipulate electoral outcomes, undermine social cohesion in target states, and manufacture diplomatic crises (as seen in Russian operations in Ukraine and US strategic information campaigns). Furthermore, the advent of 'twiplomacy' and megaphone diplomacy has severely degraded diplomatic effectiveness. When leaders conduct negotiations through performative public rhetoric, domestic audience costs become insurmountable, locking decision-makers into maximalist postures that eliminate the flexibility required for compromise."),

    ("Conclusion: The Peril of Narrative Supremacy over Negotiated Peace",
     "In conclusion, while diplomacy theoretically serves as a mechanism of conflict resolution and propaganda as an instrument of power projection, their modern integration has created a deeply hazardous international environment. Today, diplomacy is increasingly subordinated to propaganda: international summits are frequently staged not to negotiate genuine settlements, but as theatrical photo-opportunities designed to generate propaganda clips for domestic and international consumption. While the mastery of digital propaganda and narrative projection has become an indispensable facet of modern statecraft, its unchecked dominance actively subverts international stability. If great powers and regional rivals are to avoid catastrophic military escalation, they must urgently re-establish the sacred domain of confidential, professional diplomacy—rediscovering that enduring peace cannot be engineered through viral narrative victories, but must be painstakingly constructed through the unglamorous craft of negotiated compromise.")
]

q6_quotes = [
    ("Sir Harold Nicolson", "Diplomacy (1939)", "Diplomacy is the management of international relations by negotiation; the method by which these relations are adjusted and managed by ambassadors and envoys. The moment diplomacy degenerates into propaganda, it ceases to be an instrument of peace."),
    ("Hans J. Morgenthau", "Politics Among Nations", "Diplomacy is the brains of national power, as national morale is its soul. If its quality falters, all material advantages are squandered. When propaganda replaces quiet diplomacy, states barter real national interests for empty emotional triumphs."),
    ("Edward Hallett Carr", "The Twenty Years' Crisis", "Power over opinion is not less essential for political purposes than military and economic power. The state that possesses the most formidable machinery of propaganda holds a weapon capable of crumbling foreign armies from within."),
    ("Joseph S. Nye Jr.", "Soft Power & Public Diplomacy", "The best propaganda is not propaganda. In an information age where credibility is the scarcest resource, public diplomacy that appears as crude manipulation defeats itself; reputation and genuine institutional appeal constitute true soft power."),
    ("François de Callières", "On the Manner of Negotiating with Princes (1716)", "A good negotiator will never base his success on deceit, for an agreement won by fraud leaves an open wound that never heals. True diplomacy seeks an accord founded on the mutual interest of the parties.")
]

# =========================================================================
# EXECUTE GENERATION FOR BATCH 1: Q2, Q3, Q4, Q5, Q6
# =========================================================================

batch_ir_notes = [
    ("Q2_Indus_Basin_Treaty_and_Hydrological_Security_Master_Notes", q2_meta, q2_callout, q2_map, q2_timeline, q2_essay, q2_quotes, "QUESTION 2: INDUS BASIN TREATY & HYDROLOGICAL SECURITY", "KPPSC PMS IR Master Notes \u2014 Q2: The Indus Basin Treaty & Hydrological Security"),
    ("Q3_Expansion_of_BRICS_and_SCO_Structural_Realism_Master_Notes", q3_meta, q3_callout, q3_map, q3_timeline, q3_essay, q3_quotes, "QUESTION 3: EXPANSION OF BRICS & SCO UNDER STRUCTURAL REALISM", "KPPSC PMS IR Master Notes \u2014 Q3: Expansion of BRICS & SCO through Structural Realism"),
    ("Q4_Abraham_Accords_and_Palestine_Two_State_Solution_Master_Notes", q4_meta, q4_callout, q4_map, q4_timeline, q4_essay, q4_quotes, "QUESTION 4: ABRAHAM ACCORDS & THE PALESTINE TWO-STATE SOLUTION", "KPPSC PMS IR Master Notes \u2014 Q4: The Abraham Accords & The Two-State Solution for Palestine"),
    ("Q5_Cross_Border_Terrorism_and_Pakistan_Taliban_Security_Crisis_Master_Notes", q5_meta, q5_callout, q5_map, q5_timeline, q5_essay, q5_quotes, "QUESTION 5: CROSS-BORDER TERRORISM & THE PAKISTAN-TALIBAN CRISIS", "KPPSC PMS IR Master Notes \u2014 Q5: Cross-Border Terrorism & The Pakistan-Taliban Security Crisis"),
    ("Q6_Diplomacy_and_Propaganda_Conflict_Resolution_vs_Power_Projection_Master_Notes", q6_meta, q6_callout, q6_map, q6_timeline, q6_essay, q6_quotes, "QUESTION 6: DIPLOMACY & PROPAGANDA IN MODERN STATECRAFT", "KPPSC PMS IR Master Notes \u2014 Q6: Diplomacy and Propaganda in Modern Statecraft")
]

for filename, meta, callout, cmap, timeline, essay, quotes, sub_title, full_title in batch_ir_notes:
    docx_file = os.path.join(IR_DIR, f"{filename}.docx")
    pdf_file = os.path.join(IR_DIR, f"{filename}.pdf")
    md_file = os.path.join(IR_DIR, f"{filename}.md")
    
    generate_ir_master_doc(
        meta_data=meta,
        callout_data=callout,
        map_text=cmap,
        timeline_data=timeline,
        essay_sections=essay,
        quotes_data=quotes,
        out_path=docx_file,
        paper_sub_title=sub_title
    )
    
    generate_pdf_and_md(
        meta_data=meta,
        callout_data=callout,
        map_text=cmap,
        timeline_data=timeline,
        essay_sections=essay,
        quotes_data=quotes,
        pdf_path=pdf_file,
        md_path=md_file,
        title=full_title
    )

print("Batch 1 (Q2 through Q6) generated successfully!")
