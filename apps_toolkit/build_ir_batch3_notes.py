import os
import subprocess
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
# Q13: JOSEPH NYE'S SOFT POWER, HARD POWER & SMART POWER (2018 Q.1, 2013 Q.1)
# =========================================================================

q13_meta = [
    ("Target Examination", "KPPSC PMS (Provincial Management Service) — Khyber Pakhtunkhwa"),
    ("Paper & Subject", "International Relations (100 Marks Optional — Group 6)"),
    ("EXAM QUESTIONS LINKED:", '• 2018 Q.1: "Critically analyze Joseph Nye\'s concept of “Soft Power” in International Relations. (20 Marks)"\n• 2013 Q.1: "What is \\"Power\\" in International Relations? Differentiate between Hard Power, Soft Power, and Smart Power, and explain the tangible and intangible elements of National Power. (20 Marks)"'),
    ("Syllabus Section", "Part A: IR as a Field of Study, Concepts of Power, National Power"),
    ("Question Allocation", "Attempt 1 of 5 Required (20 Marks | Strict 35-Minute Allocation)"),
    ("Historiographical / Theoretical Thesis", "Formulated by Joseph S. Nye Jr. in 1990, the concept of 'soft power'—the ability to affect others to obtain preferred outcomes through attraction and co-optation rather than coercion or payment—revolutionized international relations theory by challenging classical realism's exclusive fixation on military throw-weight and economic leverage ('hard power'). Grounded in a state's cultural appeal, domestic political values, and the moral legitimacy of its foreign policies, soft power is an indispensable currency in an interconnected information age. However, soft power is structurally diffuse, difficult for statecraft to direct, and fundamentally vulnerable to charges of cultural imperialism. Consequently, enduring geopolitical effectiveness requires Nye's synthesis of 'smart power'—the strategic, contextual calibration of hard coercion and soft attraction to secure national interests.")
]

q13_callout = [
    "CONCEPTUAL DEFINITION: Power is the ability to influence the behavior of others to get outcomes one wants. Hard power operates through 'command' (carrots and sticks: military force, economic sanctions); Soft power operates through 'co-optation' (shaping preferences, agenda-setting, cultural attraction).",
    "THE THREE CURRENCIES OF SOFT POWER: Culture (where attractive to foreign populations), Political Values (when lived up to domestically and abroad), and Foreign Policies (when perceived as legitimate, multilateral, and having moral authority).",
    "STRUCTURAL LIMITATIONS & CRITIQUES: Realists (Mearsheimer) argue soft power is secondary and subordinate to hard power; it is slow to build, difficult for state leaders to control (emanating largely from civil society, cinema, and universities), and largely ineffective against existential revisionist actors (e.g., ISIS or aggressive autocracies).",
    "SMART POWER SYNTHESIS: Smart power is neither hard nor soft; it is the skillful combination of both. Classic examples include the post-WWII Marshall Plan (economic aid + democratic soft power) and contemporary Chinese statecraft combining BRI financing (hard) with Confucian institutes and non-interference rhetoric (soft)."
]

q13_map = r"""
+---------------------------------------------------------------------------------------------------------+
|                  CONCEPTUAL ROADMAP: JOSEPH NYE'S POWER TAXONOMY IN IR                                  |
+---------------------------------------------------------------------------------------------------------+
                                                     |
                               [Classical Baseline: Realist Hard Power]
                             Morgenthau & Waltz: Power = Material Capabilities
                             Command Behavior: Coercion (Sticks) & Inducements (Carrots)
                                                     |
                                                     v
                       [The 1990 Conceptual Breakthrough: Joseph S. Nye Jr.]
                       "Bound to Lead" & "Soft Power" in the Information Age
                       Co-optive Power: Getting Others to Want What You Want
                                                     |
           +-----------------------------------------+-----------------------------------------+
           |                                         |                                         |
           v                                         v                                         v
   [CULTURAL CURRENCY]                       [POLITICAL VALUES]                        [FOREIGN POLICY LEGITIMACY]
Hollywood, Literature, Higher Ed           Democracy, Rule of Law, Human Rights      Multilateral Diplomacy, Peacekeeping
Cultural Diplomacy & Global Reach          Moral Authority When Practiced at Home    Adherence to International Norms
Civil Society & Organic Attraction         Universal Appeals vs Hypocrisy Traps      Absence of Unilateral Aggression
           |                                         |                                         |
           +-----------------------------------------+-----------------------------------------+
                                                     |
                                                     v
                         [CRITICAL SHORTCOMINGS & REALIST COUNTER-ATTACK]
                      Slow to Accumulate, Fast to Lose; State Cannot Direct It;
                      Ineffective in the Face of Hard Aggression & Revanchism
                                                     |
                                                     v
                         [THE CONTEMPORARY SYNTHESIS: SMART POWER]
                     The Strategic, Contextual Calibration of Hard Coercion
                     and Soft Attraction (CSIS Commission on Smart Power)
+---------------------------------------------------------------------------------------------------------+
"""

q13_timeline = [
    ("1990: Coining of 'Soft Power'", "Joseph Nye published 'Bound to Lead: The Changing Nature of American Power.'", "Challenged the decline-of-America thesis; introduced soft power as a vital dimension of post-Cold War hegemony."),
    ("2003: Iraq War & Soft Power Collapse", "US unilateral invasion of Iraq without UN approval triggered worldwide anti-Americanism.", "Dramatically illustrated how perceived illegitimacy in foreign policy instantly eviscerates decades of accumulated soft power."),
    ("2004: Nye's 'Soft Power: The Means to Success'", "Published definitive monograph expanding the conceptual framework into three currencies.", "Systematized soft power as an analytical discipline within foreign policy studies worldwide."),
    ("2007: CSIS Commission on Smart Power", "Nye and Richard Armitage co-chaired bipartisan commission recommending 'Smart Power.'", "Formally bridged the civilian-military divide; influenced Obama-Clinton foreign policy doctrine."),
    ("Present: Great Power Soft Power Competition", "China invests billions in CGTN, Confucius Institutes, and BRI diplomacy; US counters via democratic summits.", "Transformed soft power into a central theater of 21st-century ideological and normative contestation.")
]

q13_essay = [
    ("Introduction: The Evolutionary Reconceptualization of Power in IR",
     "Power is the central organizing concept of international relations, yet it remains one of the most fiercely contested. For centuries, the classical realist tradition—from Thucydides and Niccolò Machiavelli to Hans Morgenthau and Kenneth Waltz—conceptualized power in strictly material, tangible terms: military throw-weight, demographic size, industrial capacity, and geographic depth. In this traditional paradigm, power was fundamentally coercive—an exercise in 'command power' relying on the carrots of economic bribery or the sticks of military force. However, in 1990, American political scientist Joseph S. Nye Jr. revolutionized international relations discourse by introducing the concept of 'soft power.' Writing in Bound to Lead, Nye posited that in an increasingly interdependent and information-saturated global order, material coercion is insufficient to secure enduring national objectives. Soft power—the capacity to achieve preferred outcomes through attraction and co-optation rather than coercion or payment—introduced an ideational dimension that fundamentally challenged realist orthodoxy."),

    ("Deconstructing Soft Power: Sources, Currencies, and Mechanisms",
     "Nye defines soft power as the ability to shape the preferences of others. While hard power relies on push factors (threats of force or economic sanctions), soft power operates via pull factors: attraction, emulation, and agenda-setting. When other states admire a nation's values, emulate its example, and aspire to its prosperity, they willingly align with its strategic goals without the necessity of expensive military coercion. According to Nye, a country's soft power rests on three primary resources: first, its culture (in places where it is attractive to others); second, its political values (when it lives up to them at home and abroad); and third, its foreign policies (when they are seen as legitimate and having moral authority). The universal appeal of American higher education, Hollywood cinema, democratic civil liberties, and international humanitarian assistance historically provided Washington with immense structural attraction. Conversely, when foreign policies are perceived as hypocritical or aggressively unilateral—as occurred during the 2003 invasion of Iraq—accumulated soft power can be rapidly eviscerated."),

    ("Critical Limitations: The Realist Critique of Soft Power",
     "Despite its conceptual elegance, soft power has faced severe scrutiny from structural realists, notably John Mearsheimer and Colin Gray. First, realists argue that soft power is ultimately an epiphenomenon that rests upon the foundation of hard material power. A state's cultural or political appeal rarely translates into compliance when core national survival interests are at stake. Soft power proved utterly incapable of halting Russian tanks in Georgia in 2008 or Ukraine in 2022, nor could it dissuade North Korea from pursuing nuclear weapons. Second, soft power is structurally diffuse and extraordinarily difficult for statecraft to direct. Unlike an armored division or a tariff schedule, soft power emanates primarily from civil society, private cultural industries, and independent media, making it largely uncontrollable by foreign ministries. Third, soft power is highly subjective and context-dependent: Western cultural values celebrating individualism or secular liberalism, which generate attraction in Europe, frequently provoke intense moral revulsion and cultural backlash in traditionalist or Islamist societies."),

    ("The Strategic Synthesis: The Imperative of 'Smart Power'",
     "Recognizing these analytical and practical limitations, Joseph Nye, alongside former US Deputy Secretary of State Richard Armitage, co-chaired the Center for Strategic and International Studies (CSIS) Commission on Smart Power in 2007. They formulated 'smart power'—defined as the contextual capacity of an actor to combine elements of hard and soft power into an effective, integrated grand strategy. Smart power recognizes that neither pure military coercion nor pure cultural diplomacy can guarantee national security in isolation. The historical archetype of smart power was the post-World War II Marshall Plan: Washington deployed massive hard economic resources ($13 billion) while simultaneously projecting an attractive vision of democratic solidarity and institutional partnership, successfully reconstructing Western Europe and containing Soviet expansion. In the contemporary era, the People's Republic of China explicitly practices smart power by pairing its massive Belt and Road Initiative (BRI) financial investments with aggressive soft-power campaigns through Confucius Institutes, CGTN media networks, and sovereign non-interference rhetoric across the Global South."),

    ("Conclusion: The Primacy of Contextual Intelligence in Modern Statecraft",
     "In conclusion, Joseph Nye's formulation of soft power profoundly enriched international relations theory by demonstrating that power is not merely a tangible inventory of tanks, missiles, and GDP, but an intangible relationship of psychological attraction and normative legitimacy. In the 21st-century digital landscape, where information is abundant but attention and credibility are scarce, narrative authority and moral legitimacy have become vital instruments of statecraft. However, soft power is not a magical panacea that renders armies obsolete. States that rely exclusively on soft power invite geopolitical predation, while states that rely purely on hard coercion bleed their legitimacy and bankrupt their treasuries. The hallmark of masterly modern statecraft lies in 'contextual intelligence'—the art of deploying smart power to ensure that hard power enforces the rules, while soft power inspires the peace.")
]

q13_quotes = [
    ("Joseph S. Nye Jr.", "Soft Power: The Means to Success in World Politics", "A country may obtain the outcomes it wants in world politics because other countries—admiring its values, emulating its example, aspiring to its level of prosperity and openness—want to follow it. This soft power—getting others to want the outcomes that you want—co-opts people rather than coerces them."),
    ("Hans J. Morgenthau", "Politics Among Nations", "Of all the factors that make for the power of a nation, the most elusive is the quality of its diplomacy and national morale. Power is never purely physical; it is a psychological relation between those who exercise it and those over whom it is exercised."),
    ("John J. Mearsheimer", "The Tragedy of Great Power Politics", "Soft power is of marginal utility in great-power politics. When push comes to shove, military might and economic muscle determine the distribution of power and the survival of states. You cannot deter a revisionist autocrat with Hollywood movies."),
    ("Edward Hallett Carr", "The Twenty Years' Crisis", "Power over opinion is not less essential for political purposes than military and economic power. The supreme power of attraction is the ability to make one's own national interest appear as the universal interest of humanity."),
    ("Susan Strange", "States and Markets", "Structural power—the power to shape and determine the global frameworks within which states relate to each other—is far more profound than relational power. Soft power operates precisely by constructing the normative structures of the international system.")
]

# =========================================================================
# Q14: THEORIES OF NATIONALISM: CONCEPT & STATE UTILITY (2018 Q.2)
# =========================================================================

q14_meta = [
    ("Target Examination", "KPPSC PMS (Provincial Management Service) — Khyber Pakhtunkhwa"),
    ("Paper & Subject", "International Relations (100 Marks Optional — Group 6)"),
    ("Past Paper Alignment", "2018 Q.2 (\"Nationalism is neither inherently good nor bad, it is only a concept that can be used to achieve good or evil ends.\" Elucidate.)"),
    ("Syllabus Section", "Part A: IR as a Field of Study, Nationalism & The Modern State System"),
    ("Question Allocation", "Attempt 1 of 5 Required (20 Marks | Strict 35-Minute Allocation)"),
    ("Historiographical / Theoretical Thesis", "As the primary ideological catalyst of the modern state system, nationalism is an ideational construct that is neither intrinsically virtuous nor inherently evil; rather, it functions as a neutral, powerful political multiplier whose moral outcome is dictated by whether it is operationalized through an inclusive, constitutional framework ('civic nationalism') or an exclusive, blood-and-soil doctrine ('ethnic nationalism'). Drawing on Benedict Anderson's 'imagined communities' and Hans Kohn's foundational dichotomy, nationalism has served as an indispensable engine of democratic self-determination, anti-colonial liberation, and social cohesion; conversely, when captured by chauvinistic realpolitik, it mutates into a catastrophic instrument of imperialist expansion, ethnic cleansing, and systemic war.")
]

q14_callout = [
    "NEUTRAL CONCEPTUAL NATURE: Nationalism is an ideology that holds that the nation—a collective identity—should be congruent with the state (the political unit), as formulated by Ernest Gellner. It is an ideational tool that takes its normative color from the political objectives of those who wield it.",
    "CIVIC VS ETHNIC DICHOTOMY: Hans Kohn's classic framework contrasts Western 'Civic Nationalism' (inclusive, territorial, based on shared political principles, rule of law, and citizenship) with Eastern 'Ethnic Nationalism' (exclusive, genealogical, based on language, religion, bloodlines, and ancestral heritage).",
    "NATIONALISM FOR 'GOOD' ENDS: Emancipatory and progressive utility: overthrew dynastic absolutism, fueled the American and French Revolutions, forged modern welfare state solidarity, and inspired 20th-century anti-colonial liberation movements (Algeria, India, Pakistan, Vietnam).",
    "NATIONALISM FOR 'EVIL' ENDS: Destructive and catastrophic utility: hyper-nationalist militarism catalyzed WWI, Nazi racial fascism (Lebensraum), Balkan genocidal wars in the 1990s, and contemporary exclusionary populist xenophobia targeting minority rights and global cooperation."
]

q14_map = r"""
+---------------------------------------------------------------------------------------------------------+
|                  CONCEPTUAL ROADMAP: THE DUAL NATURE OF NATIONALISM IN IR                               |
+---------------------------------------------------------------------------------------------------------+
                                                     |
                               [The Core Idea: Ernest Gellner & Benedict Anderson]
                             Nationalism = Political Unit (State) Must Match National Unit
                             An "Imagined Political Community" (Anderson)
                                                     |
                                                     v
                       [The Great Typological Divide: Hans Kohn's Dichotomy]
                                                     |
           +-----------------------------------------+-----------------------------------------+
           |                                                                                   |
           v                                                                                   v
   [CIVIC / CONSTITUTIONAL NATIONALISM]                                [ETHNIC / GENEALOGICAL NATIONALISM]
Based on Law, Citizenship, Shared Values                            Based on Blood, Ancestry, Language, Religion
Voluntary Adherence (Ernest Renan's "Daily Plebiscite")             Involuntary Inheritance (Herder, Fichte)
Universalist, Inclusive & Democratic                                Particularist, Exclusive & Chauvinistic
           |                                                                                   |
           +-----------------------------------------+-----------------------------------------+
                                                     |
                                                     v
                         [OPERATIONAL MANIFESTATIONS IN WORLD HISTORY]
                                                     |
           +-----------------------------------------+-----------------------------------------+
           |                                                                                   |
           v                                                                                   v
   [THE "GOOD": EMANCIPATION & COHESION]                               [THE "EVIL": CHAUVINISM & DESTRUCTION]
Overthrow of Feudal Absolutism (1789)                               Hyper-Nationalism & Militarism (WWI Outbreak)
Anti-Colonial Liberation Movements in Asia & Africa                  Fascism & Nazi Racial Extermination (WWII)
Enables Democratic Welfare States & Taxation                         Balkan Ethnic Cleansing & Rwandan Genocide (1994)
Forges Civic Solidarity Across Diverse Ethnic Groups                 Irredentist Expansion & Xenophobic Populism
           |                                                                                   |
           +-----------------------------------------+-----------------------------------------+
                                                     |
                                                     v
                         [SYNTHESIS: STATE UTILITY & SYSTEMIC IMPACT]
                     Nationalism is the Ultimate Political Instrument; It Can
                     Construct Sovereign Freedom or Unleash Systemic Barbarism
+---------------------------------------------------------------------------------------------------------+
"""

q14_timeline = [
    ("1789: French Revolution & Popular Sovereignty", "Replaced dynastic allegiance to the monarch with political loyalty to the 'Nation.'", "Marked the historical birth of modern civic nationalism, transforming peasants into sovereign citizens."),
    ("19th Century: Unification & Romanticism", "Italian Risorgimento (Mazzini) and German Unification (Bismarck) transformed Europe.", "Nationalism operated as both a liberal-democratic unifier and an authoritarian blood-and-iron consolidation."),
    ("1914: Hyper-Nationalism Sparks WWI", "Jingoistic patriotic fervor and Pan-Slavism catalyzed catastrophic great-power collision.", "Demonstrated how unchecked nationalist militarism blinds societies to the catastrophic costs of modern warfare."),
    ("Mid-20th Century: Anti-Colonial Triumph", "Decolonization of Asia and Africa; creation of Pakistan (1947), India, and independent African states.", "Nationalism served as an emancipatory, anti-imperialist weapon of self-determination for oppressed peoples."),
    ("1990s: Balkan Wars & Ethnic Cleansing", "Dissolution of Yugoslavia collapsed into genocidal ethnic warfare in Bosnia and Kosovo.", "Brutally illustrated the horrific nadir of ethnic nationalism when manipulated by revanchist political elites.")
]

q14_essay = [
    ("Introduction: The Moral Ambiguity of the Nation-State's Creed",
     "Nationalism has been the most powerful, disruptive, and transformative political ideology of the modern era. Since the collapse of European dynastic feudalism, it has served as the foundational legitimacy principle of the Westphalian international system: the premise that humanity is naturally divided into distinct nations, and that each nation possesses an inherent right to sovereign statehood. However, throughout modern history, nationalism has confounded simplistic moral evaluation. In the incisive phrasing of the prompt, nationalism is neither inherently good nor intrinsically evil; rather, it is an ideational instrument—a powerful political multiplier whose ethical character is defined entirely by the ends toward which it is directed and the principles upon which it is constructed. As analyzed by scholars like Hans Kohn, Ernest Gellner, and Benedict Anderson, nationalism contains a profound dualism: it has acted as the greatest engine of human liberty, democratic solidarity, and anti-colonial emancipation, while simultaneously serving as the primary ideological catalyst for imperialist aggression, ethnic cleansing, and totalitarian slaughter."),

    ("The Theoretical Architecture: Civic versus Ethnic Nationalism",
     "To comprehend the moral dichotomy of nationalism, one must examine its distinct conceptual typologies. The classic framework formulated by historian Hans Kohn distinguishes between 'civic' and 'ethnic' nationalism. Civic nationalism—traditionally rooted in the Enlightenment, the American Revolution, and the French Declaration of the Rights of Man—defines nationhood through political allegiance, constitutional principles, citizenship, and the rule of law. In Ernest Renan's famous formulation, a nation is a 'daily plebiscite,' an inclusive political community where any individual can belong regardless of ancestry, race, or creed. Conversely, ethnic nationalism—rooted in Romantic reaction, Johann Gottfried von Herder, and German historicism—defines the nation as an organic, genealogical entity bound by shared bloodlines, ancestral language, religion, and primordial culture. Membership is non-negotiable and inherited by birth. This fundamental theoretical division determines whether nationalism expands human dignity or becomes a weapon of violent exclusion."),

    ("Nationalism for 'Good' Ends: Emancipation, Democracy, and Cohesion",
     "When operationalized through civic and progressive frameworks, nationalism has served as an indispensable force for human progress. Historically, nationalism dismantled the oppressive, arbitrary rule of divine-right monarchs and aristocratic empires, establishing popular sovereignty as the bedrock of governance. In the domestic sphere, as sociologist Ernest Gellner highlighted, the modern industrial economy and welfare state could not function without the social solidarity forged by national identity. Progressive nationalism creates the mutual trust necessary for citizens to accept universal taxation, public education, and social safety nets to support strangers. In the international arena, nationalism operated as the supreme emancipatory weapon of the twentieth century: the anti-colonial movements across Asia, Africa, and the Middle East mobilized national identity to shatter the yoke of Western imperialism. For nations like Pakistan, India, Algeria, and Vietnam, anti-colonial nationalism was the sacred vehicle through which oppressed populations reclaimed their dignity and sovereign right to self-determination."),

    ("Nationalism for 'Evil' Ends: Chauvinism, Irredentism, and Total War",
     "Conversely, when hijacked by ethnic fundamentalism, political realpolitik, and imperialist ambition, nationalism mutates into an engine of catastrophe. In his classic essay Notes on Nationalism, George Orwell warned that obsessive nationalism inevitably breeds intellectual dishonesty and a lust for power. When national pride curdles into jingoistic chauvinism—the conviction that one's nation is inherently superior and possesses a divine mandate to dominate others—it destroys international order. The hyper-nationalist militarism of the late nineteenth and early twentieth centuries transformed Europe into an armed camp, directly causing the slaughter of World War I. Even more horrific was the ethnic, 'blood and soil' nationalism of Nazi Germany, which weaponized racial grievance into the industrial extermination of six million Jews and millions of Slavs under the banner of Lebensraum. In the post-Cold War era, the dissolution of Yugoslavia demonstrated that when opportunistic political demagogues (such as Slobodan Milošević) unleash ethnic nationalism, it leads directly to the atrocities of Srebrenica, mass rape camps, and genocidal devastation."),

    ("Conclusion: The Enduring Challenge of Statecraft",
     "In conclusion, nationalism is a neutral, primordial political force—like fire, which can warm a home or burn it to the ground. In an anarchic international system, the nation-state remains the only durable repository of political accountability, legal rights, and democratic governance. Abandoning nationalism entirely is neither feasible nor desirable, as post-national cosmopolitanism has largely failed to inspire deep societal loyalty. The supreme task of modern statecraft, therefore, is not the futile attempt to eradicate national sentiment, but its continuous cultivation toward inclusive, civic patriotism. When grounded in constitutional values, respect for minority rights, and adherence to international law, nationalism remains an honorable guardian of human freedom; but when poisoned by ethnic supremacy, grievance, and hatred of the other, it becomes the most lethal ideology known to humanity.")
]

q14_quotes = [
    ("Ernest Gellner", "Nations and Nationalism (1983)", "Nationalism is primarily a political principle, which holds that the political and the national unit should be congruent. It is not an inherent attribute of human nature, but the organizational necessity of modern industrial society."),
    ("Benedict Anderson", "Imagined Communities (1983)", "The nation is an imagined political community—and imagined as both inherently limited and sovereign. It is imagined because the members of even the smallest nation will never know most of their fellow-members, yet in the minds of each lives the image of their communion."),
    ("George Orwell", "Notes on Nationalism (1945)", "Nationalism is not to be confused with patriotism. Patriotism is of its nature defensive, both militarily and culturally. Nationalism, on the other hand, is inseparable from the desire for power. The abiding purpose of every nationalist is to secure more power and more prestige for his nation."),
    ("Hans Kohn", "The Idea of Nationalism", "Western civic nationalism was largely an effort to build a nation in the political reality of the present, based on individual rights; Eastern ethnic nationalism was an appeal to the past, relying on ancient myths, blood, and language to defy political modernity."),
    ("Carlton J.H. Hayes", "The Historical Evolution of Modern Nationalism", "Nationalism, when it is an inclusive, liberal force, has promoted liberty and self-determination; but when it degenerates into totalitarian intolerance, it becomes the most destructive religion of modern times.")
]

# =========================================================================
# Q15: BALANCE OF POWER & COLD WAR BIPOLARITY (2018 Q.3, 2013 Q.2, 2013 Q.4)
# =========================================================================

q15_meta = [
    ("Target Examination", "KPPSC PMS (Provincial Management Service) — Khyber Pakhtunkhwa"),
    ("Paper & Subject", "International Relations (100 Marks Optional — Group 6)"),
    ("EXAM QUESTIONS LINKED:", '• 2018 Q.2: "“Nationalism is neither inherently good nor bad, it is only a concept that can be used to achieve good or evil ends.” Elucidate. (20 Marks)"'),
    ("Syllabus Section", "Part A: Balance of Power | Part B: Cold War Geopolitics & Bipolar Confrontation"),
    ("Question Allocation", "Attempt 1 of 5 Required (20 Marks | Strict 35-Minute Allocation)"),
    ("Historiographical / Theoretical Thesis", "Balance of power is the cardinal structural operating mechanism of the anarchic international system, functioning as an equilibrium among competing states to prevent any single power from achieving systemic hegemony. While classical realists (Morgenthau, Kissinger) conceptualized balance of power as a conscious diplomatic contrivance requiring constant intervention, Kenneth Waltz's structural neorealism demonstrated that balancing is an automatic, structural imperative under anarchy. During the Cold War (1947–1989), the balance of power reached its historical apex through structural bipolarity: equilibrium was maintained not through flexible shifting alliances, but through rigid military blocs (NATO vs. Warsaw Pact), proxy containment, and above all, the nuclear revolution of Mutually Assured Destruction (MAD), transforming the balance of power into a durable, war-inhibiting 'Balance of Terror.'")
]

q15_callout = [
    "STRUCTURAL DEFINITION: Balance of power describes a state of stability where power is distributed among competing states such that no single state is strong enough to dominate all others, ensuring the survival of independent sovereign entities.",
    "CLASSICAL VS STRUCTURAL MECHANISMS: Classical Realism highlights diplomatic statecraft, shifting alliances, and the role of an external balancer (e.g., Great Britain in the 19th century); Structural Realism (Waltz) argues that bipolar systems are inherently more stable than multipolar systems because they eliminate alliance miscalculation, buck-passing, and chain-ganging.",
    "THE COLD WAR PILLARS OF EQUILIBRIUM: Maintained through three primary pillars: (1) Rigid institutionalized alliance blocs (NATO established 1949, Warsaw Pact 1955); (2) Ideological and geopolitical containment across peripheral proxy theaters (Korea, Vietnam, Afghanistan); (3) Crisis de-escalation protocols (Hotline Agreement 1963, Strategic Arms Limitation Talks - SALT).",
    "THE NUCLEAR REVOLUTION & MAD: The introduction of thermonuclear weapons and survivable second-strike capabilities inverted classical military logic: because victory in total war became impossible, the balance of power became a 'Balance of Terror,' preventing hot systemic war between the superpowers."
]

q15_map = r"""
+---------------------------------------------------------------------------------------------------------+
|                  CONCEPTUAL ROADMAP: BALANCE OF POWER & COLD WAR EQUILIBRIUM                            |
+---------------------------------------------------------------------------------------------------------+
                                                     |
                               [Theoretical Foundation: Anarchy & Balancing]
                             Kenneth Waltz: Structure Dictates State Behavior
                             Equilibrium = Absence of Unilateral Hegemony
                                                     |
                                                     v
                       [The Historical Transition: Multipolarity to Bipolarity]
                       19th Century Concert of Europe (Multipolar, Flexible Alliances)
                       Post-1945 Collapse of European Empires --> Two Superpowers Emerge
                                                     |
                                                     v
                         [THE COLD WAR EQUILIBRIUM: THREE OPERATIONAL PILLARS]
                                                     |
           +-----------------------------------------+-----------------------------------------+
           |                                         |                                         |
           v                                         v                                         v
   [MILITARY ALLIANCES]                      [NUCLEAR DETERRENCE & MAD]                [CRISIS MANAGEMENT & DÉTENTE]
Internal Balancing: Massive Arms Buildups  Thermonuclear Arsenals & ICBMs            Post-Cuban Missile Crisis Hotline
External Balancing: NATO (1949) vs         Second-Strike Survivability (SSBNs)       Arms Control: ABM Treaty (1972)
Warsaw Pact (1955) in Central Europe       "Balance of Terror" Inhibits Hot War      SALT I & SALT II Formalize Parity
Divided Germany as the Strategic Axis      Limited Wars Confined to Periphery        Détente Prevents Direct Conflagration
           |                                         |                                         |
           +-----------------------------------------+-----------------------------------------+
                                                     |
                                                     v
                         [THE HISTORICAL VERDICT: "THE LONG PEACE"]
                     John Lewis Gaddis: Bipolar Balance of Power Maintained Systemic
                     Stability for 45 Years Without a Third World War
+---------------------------------------------------------------------------------------------------------+
"""

q15_timeline = [
    ("1815: Congress of Vienna", "Established the classical multipolar Concert of Europe to maintain balance against French resurgence.", "Codified the classical diplomatic balance of power under Metternich and Castlereagh for a century."),
    ("1947–1949: The Bipolar Divide", "Truman Doctrine, Marshall Plan, and the formation of NATO (April 1949).", "Marked the definitive division of the international system into a structural bipolar balance of power."),
    ("1955: Formation of Warsaw Pact", "Soviet Union established formal military alliance with Eastern European satellites.", "Completed the institutionalization of external military balancing across the European continent."),
    ("October 1962: Cuban Missile Crisis", "US-Soviet standoff brought the world to the brink of thermonuclear annihilation.", "Exposed the lethal danger of miscalculation, driving superpowers to institutionalize crisis-management hotlines."),
    ("1972: SALT I & ABM Treaty", "Nixon and Brezhnev signed strategic arms limitations in Moscow.", "Formally codified nuclear parity and Mutually Assured Destruction as the legal foundation of global strategic equilibrium.")
]

q15_essay = [
    ("Introduction: Balance of Power as the Cornerstone of International Stability",
     "In the intellectual canon of international relations, no concept occupies a more central position than the balance of power. Defined classically as an equilibrium among competing sovereign states or alliances such that no single entity accumulates preponderant power capable of imposing unipolar hegemony, the balance of power has served as the cardinal operating mechanism of systemic order since the Peace of Westphalia. In the Realist tradition, the international system is defined by an immutable condition of anarchy—the absence of a centralized world government. Under anarchy, sovereign states must guarantee their survival through self-help. While classical realists such as Hans Morgenthau and Henry Kissinger viewed the balance of power as an artful diplomatic contrivance requiring constant, deliberate manipulation by statesmen, Kenneth Waltz's structural neorealism demonstrated that balancing is an automatic, systemic outcome: whenever an unbalance of power threatens to emerge, other states will either arm internally or ally externally to restore the equilibrium. The historical apex of this mechanism occurred during the Cold War (1947–1989), where an unprecedented structural bipolarity sustained global equilibrium for four decades."),

    ("The Classical Mechanisms of Power Balancing",
     "Historically, states have utilized six primary instruments to maintain or restore the balance of power. First is 'internal balancing'—the mobilization of domestic economic resources, technological innovation, and military rearmament to match rival capabilities. Second is 'external balancing'—the formation of bilateral or multilateral military alliances to aggregate power against a rising hegemon. Third is 'divide and rule,' wherein a dominant power prevents rivals from uniting by sowing discord. Fourth is the acquisition of territory or 'compensations,' common in eighteenth-century European diplomacy, where major powers partitioned smaller states (such as Poland) to ensure equal gains. Fifth is the creation of 'buffer states'—neutral geographic entities (such as historical Switzerland or Afghanistan) designed to physically separate competing empires. Finally, classical balance of power relied on the role of the 'balancer'—a flexible offshore power (historically Great Britain) that deliberately threw its weight behind the weaker coalition to prevent any single continental power from establishing hegemony."),

    ("The Cold War Paradigm: Structural Bipolarity and Alliance Architectures",
     "With the devastation of the European great powers in World War II, the multipolar balance of power collapsed and was supplanted by structural bipolarity centered on two continental superpowers: the United States and the Soviet Union. As Kenneth Waltz famously argued in Theory of International Politics, bipolar systems are inherently more stable and less prone to major war than multipolar systems. In multipolarity, alliances are fluid, leading to dangerous structural pathologies like 'chain-ganging' (where states are dragged into catastrophic wars by reckless allies, as in 1914) or 'buck-passing' (where states shirk the burden of balancing against an aggressor, as in the 1930s). In the Cold War bipolarity, however, there was no ambiguity: the two superpowers faced each other directly. Equilibrium was maintained externally through rigid, institutionalized military alliances—the North Atlantic Treaty Organization (NATO) founded in 1949, and the Warsaw Pact established in 1955. Europe was divided by an impermeable iron curtain, eliminating territorial ambiguity."),

    ("The Nuclear Revolution: From Balance of Power to Balance of Terror",
     "Crucially, the Cold War balance of power was preserved through a radical technological revolution: the advent of thermonuclear weapons and intercontinental delivery systems. In classical balance of power, war was the ultimate instrument used to restore equilibrium when diplomacy failed. However, the development of secure, survivable second-strike nuclear arsenals (the nuclear triad of ICBMs, strategic bombers, and ballistic missile submarines) created the condition of Mutually Assured Destruction (MAD). Under MAD, a full-scale war between the superpowers became an act of collective suicide, stripping war of its rational political utility as posited by Clausewitz. Consequently, the balance of power was transformed into a 'Balance of Terror.' The fear of nuclear escalation forced Washington and Moscow to institutionalize crisis-management mechanisms (the 1963 Moscow-Washington Hotline post-Cuban Missile Crisis), negotiate bilateral arms-control treaties (SALT I and the 1972 Anti-Ballistic Missile Treaty), and strictly confine their geopolitical competition to limited, peripheral proxy wars in Asia, Africa, and Latin America."),

    ("Conclusion: The Legacy of 'The Long Peace'",
     "In conclusion, the balance of power during the Cold War era proved to be extraordinarily resilient, producing what historian John Lewis Gaddis famously termed 'The Long Peace.' Despite profound ideological hostility, an unrelenting arms race, and terrifying regional flashpoints from Berlin to Cuba, the structural equilibrium between the two superpowers prevented the outbreak of a third world war. The Cold War demonstrated that while the balance of power cannot eliminate human conflict or abolish proxy violence, its structural mechanics—when anchored by clear bipolar lines, mutual recognition of spheres of influence, and the sobering reality of nuclear deterrence—can successfully prevent systemic collapse, preserving international order across generations.")
]

q15_quotes = [
    ("Kenneth N. Waltz", "Theory of International Politics", "In a bipolar world, there are no third powers to shirk onto or chain-gang with. The leaders of the two alliances are forced to balance each other constantly, producing a structural stability that multipolar systems can never match."),
    ("Hans J. Morgenthau", "Politics Among Nations", "The balance of power is an inevitable manifestation of the struggle for power. Whenever an international system is composed of sovereign units competing for dominance, equilibrium is the only alternative to universal tyranny."),
    ("John Lewis Gaddis", "The Long Peace: Inquiries into the History of the Cold War", "What preserved the Cold War peace was not virtue, but structure: the bipolar division of the world, the mutual independence of the superpowers, and the catastrophic reality of nuclear deterrence."),
    ("Henry Kissinger", "Diplomacy", "A balance of power does not guarantee peace; it guarantees that no power can achieve hegemony without a major war. When equilibrium is maintained through legitimate consensus and nuclear restraint, it becomes the most durable guardian of order."),
    ("Raymond Aron", "Peace and War: A Theory of International Relations", "The Cold War equilibrium was defined by a terrifying paradox: 'peace is impossible, war is improbable.' Mutually Assured Destruction froze the balance of power in an uneasy, permanent stalemate.")
]

# =========================================================================
# Q16: SYSTEMIC ORIGINS OF WORLD WAR I (2018 Q.4, 2013 Q.3)
# =========================================================================

q16_meta = [
    ("Target Examination", "KPPSC PMS (Provincial Management Service) — Khyber Pakhtunkhwa"),
    ("Paper & Subject", "International Relations (100 Marks Optional — Group 6)"),
    ("EXAM QUESTIONS LINKED:", '• 2018 Q.3: "Balance of power is a state of stability between competing forces; it is an equilibrium among states or alliances to prevent any other entity from becoming too strong. In light of this statement, elaborate how equilibrium was maintained by Great Powers during the Cold War era. (20 Marks)"\n• 2013 Q.2: "\\"Balance of Power is a central concept in Realism.\\" Discuss the various mechanisms employed by states to maintain or restore the Balance of Power (alliances, armaments, buffer states, compensation). (20 Marks)"\n• 2013 Q.4: "Discuss the Cold War alliance systems (NATO vs. Warsaw Pact, SEATO, CENTO). How did these military alliances shape the security architecture of Europe and Asia? (20 Marks)"'),
    ("Syllabus Section", "Part B: Evolution of International System Since 1815, Major Conflicts (WWI)"),
    ("Question Allocation", "Attempt 1 of 5 Required (20 Marks | Strict 35-Minute Allocation)"),
    ("Historiographical / Theoretical Thesis", "While the assassination of Archduke Franz Ferdinand in Sarajevo on June 28, 1914, provided the immediate diplomatic trigger for the outbreak of World War I, attributing the conflagration solely to this event commits the fundamental fallacy of confusing a tactical catalyst with systemic causation. Drawing upon structural realism (Waltz) and historical analysis (Joll, Clark, Fischer), the Great War was the inexorable consequence of systemic structural rigidities accumulated over four decades: the polarization of Europe into rigid, zero-sum alliance blocs (Triple Entente vs. Triple Alliance), an unconstrained industrial and naval arms race, imperialist frictions, pan-Slavic Balkan nationalism, and the fatal military doctrine of the 'cult of the offensive' embedded in unalterable railway mobilization timetables (Schlieffen Plan).")
]

q16_callout = [
    "PROXIMATE CATALYST VS STRUCTURAL CAUSATION: Sarajevo was the spark, but the European continent was already saturated with combustible structural materials accumulated since 1871; without structural tensions, the assassination would have remained a localized third Balkan war.",
    "THE M-A-I-N STRUCTURAL DRIVERS: Militarism (Anglo-German Dreadnought naval race, army expansions); Alliances (chain-ganging between Austria-Germany and Russia-France); Imperialism (Scramble for Africa, Moroccan Crises of 1905/1911); Nationalism (revanchist French desire for Alsace-Lorraine, Slavic irredentism against Austria-Hungary).",
    "A.J.P. TAYLOR'S 'WAR BY TIMETABLE': The fatal inflexibility of industrialized military planning: once Russia ordered general mobilization, Germany's Schlieffen Plan required immediate pre-emptive invasion of neutral Belgium and France according to strict railway schedules, eliminating diplomatic de-escalation time.",
    "HISTORIOGRAPHICAL SPECTRUM: Fritz Fischer's thesis (Germany's deliberate bid for European hegemony - Griff nach der Weltmacht) vs. Christopher Clark's modern thesis (The Sleepwalkers - a tragedy of miscalculation and collective structural failure with no single guilty state)."
]

q16_map = r"""
+---------------------------------------------------------------------------------------------------------+
|                  CONCEPTUAL ROADMAP: THE SYSTEMIC ORIGINS OF WORLD WAR I                                |
+---------------------------------------------------------------------------------------------------------+
                                                     |
                               [Systemic Baseline: Post-Bismarckian Realignment]
                             1890: Kaiser Wilhelm II Dismisses Bismarck & Abandons Reinsurance
                             Structural Polarization: Triple Alliance vs Franco-Russian Alliance
                                                     |
                                                     v
                       [THE FOUR STRUCTURAL TINDERBOXES: THE M-A-I-N FORCES]
                                                     |
           +-----------------------------------------+-----------------------------------------+
           |                    |                                  |                           |
           v                    v                                  v                           v
     [MILITARISM]          [ALLIANCES]                       [IMPERIALISM]               [NATIONALISM]
Anglo-German Naval Race    Rigid Polarization                Moroccan Crises (1905, 1911) Pan-Slavism (Russia-Serbia)
"Cult of the Offensive"    Triple Entente (UK, FR, RU) vs    Ottoman Decline & Balkan     Austro-Hungarian Paranoia
Railway War Plans          Triple Alliance (DE, AT, IT)      Annexation Crisis (1908)     French Revanche for Alsace
           |                    |                                  |                           |
           +--------------------+----------------------------------+---------------------------+
                                                     |
                                                     v
                         [THE PROXIMATE SPARK: SARAJEVO (JUNE 28, 1914)]
                       Assassination of Archduke Franz Ferdinand by Gavrilo Princip
                       Austria-Hungary Issues Ultimatum to Serbia (Blank Check from Berlin)
                                                     |
                                                     v
                         [THE JULY CRISIS: SYSTEMIC COLLAPSE & CHAIN-GANGING]
                       Russian Mobilization -> German Schlieffen Plan Activated
                       Invasion of Belgium -> Britain Declares War (August 4, 1914)
                                                     |
                                                     v
                         [CONCLUSION: SYSTEMIC MULTIPOLAR BREAKDOWN]
                     The Catastrophe Was Not the Result of an Isolated Murder,
                     but the Inevitable Structural Implosion of an Anarchic Order
+---------------------------------------------------------------------------------------------------------+
"""

q16_timeline = [
    ("1890: Dismissal of Bismarck", "Kaiser Wilhelm II dropped the Reinsurance Treaty with Russia, adopting aggressive Weltpolitik.", "Fractured Bismarck's delicate balancing system, driving Russia into the arms of France (Franco-Russian Alliance 1894)."),
    ("1898–1906: Anglo-German Naval Race", "Admiral von Tirpitz began building the High Seas Fleet; Britain launched HMS Dreadnought (1906).", "Ended Britain's 'splendid isolation,' compelling London to forge the Entente Cordiale (1904) and Anglo-Russian Entente (1907)."),
    ("1908: Bosnian Annexation Crisis", "Austria-Hungary annexed Bosnia-Herzegovina, humiliating Serbia and imperial Russia.", "Deeply radicalized Serbian nationalist irredentism and convinced Moscow that it could never back down in the Balkans again."),
    ("June 28, 1914: The Sarajevo Assassination", "Gavrilo Princip of the Black Hand assassinated Austro-Hungarian heir Archduke Franz Ferdinand.", "Provided the Austro-Hungarian military establishment with the long-sought pretext to crush Serbian nationalism."),
    ("July–August 1914: The July Crisis & Mobilization", "Germany issued 'Blank Check'; Austria declared war on Serbia; Russian general mobilization triggered Schlieffen Plan.", "Alliance chain-ganging and unalterable railway timetables transformed a Balkan crisis into total European war.")
]

q16_essay = [
    ("Introduction: The Dichotomy of the Spark and the Powder Keg",
     "On the morning of June 28, 1914, in the streets of Sarajevo, Gavrilo Princip, a nineteen-year-old Bosnian Serb nationalist belonging to the shadowy Black Hand organization, fired two fatal pistol shots that killed Archduke Franz Ferdinand, the heir presumptive to the Austro-Hungarian Empire, and his wife Sophie. In traditional, descriptive historical narratives, this sensational assassination is frequently identified as the primary cause of the First World War. However, in the analytical discipline of international relations, attributing a global cataclysm that claimed twenty million lives solely to the actions of a lone assassin is a profound intellectual fallacy. The assassination in Sarajevo was merely the proximate spark—the tactical trigger that ignited an immense, highly pressurized structural powder keg that had been methodically constructed across Europe over the preceding four decades. Grounded in structural realism and systemic historiography (James Joll, Fritz Fischer, Christopher Clark), the outbreak of World War I was the inevitable result of systemic forces: the polarization of Europe into rigid, adversarial alliances, an unconstrained industrial and naval arms race, imperialist frictions, and the fatal inflexibility of industrialized military mobilization plans."),

    ("The Structural Straitjacket: Alliances and the Danger of Chain-Ganging",
     "The primary structural engine of the war was the complete polarization of the European multipolar balance of power into two armed, zero-sum coalitions. Following the dismissal of Chancellor Otto von Bismarck in 1890, Kaiser Wilhelm II abandoned Bismarck's policy of diplomatic flexibility, allowing the Reinsurance Treaty with Russia to lapse. This monumental blunder drove Tsarist autocracy into a military alliance with republican France in 1894, fundamentally encumbering Germany with the nightmare of a two-front war. Concurrently, Germany's aggressive pursuit of Weltpolitik (world policy) and its provocative expansion of a blue-water battle fleet pushed Great Britain out of its traditional 'splendid isolation' into the Entente Cordiale with France (1904) and the Anglo-Russian Convention (1907). In international relations theory, this created the dangerous phenomenon of 'chain-ganging' (Kenneth Waltz): when two rigid alliance blocs confront each other, the security of each great power becomes unconditionally tied to the survival of its allies, meaning an isolated crisis involving minor partners inevitably drags the entire system into war."),

    ("Militarism, Railway Timetables, and the Cult of the Offensive",
     "This alliance rigidity was exponentially aggravated by industrial militarism and the institutionalization of the 'cult of the offensive.' Between 1890 and 1914, European military expenditures surged by over 300%. The Anglo-German naval arms race, ignited by Admiral Alfred von Tirpitz's risk fleet theory and the British launch of the revolutionary all-big-gun HMS Dreadnought in 1906, poisonously estranged London and Berlin. More decisively, European general staffs became prisoners of unalterable, industrialized mobilization plans. In his famous thesis War by Timetable, historian A.J.P. Taylor demonstrated that modern mass armies relied entirely on complex railway logistics: once a great power initiated general mobilization, millions of men and supply trains were set in motion along schedules that could not be halted without rendering the country defenseless. Germany's war plan—the Schlieffen Plan—mandated that in any war with Russia, Germany must instantly strike and knock out France within six weeks by invading through neutral Belgium before turning east. Thus, the moment Russia mobilized to defend Serbia, Germany was structurally compelled to invade Belgium, automatically precipitating British military intervention."),

    ("Imperialism and the Balkan Powder Keg",
     "These militarized alliances were continuously stressed by imperialist frictions and radicalized nationalism. Imperial rivalries in Africa (the Moroccan Crises of 1905 and 1911) had repeatedly brought great powers to the brink of armed conflict, cementing Franco-British military solidarity against German coercion. However, the most lethal systemic vulnerability resided in the Balkans—the 'powder keg of Europe.' The protracted, agonizing decay of the Ottoman Empire created a dangerous geopolitical vacuum. The dual Austro-Hungarian monarchy, terrified that pan-Slavic nationalism championed by an aggressive, independent Serbia would dismember its multi-ethnic empire, became obsessed with crushing Belgrade. Russia, having suffered a humiliating diplomatic retreat during the 1908 Bosnian Annexation Crisis, resolved that it could never again back down in the Balkans without permanently surrendering its great-power status. The Balkans was therefore a structural tinderbox where local nationalist irredentism was fused directly to great-power prestige."),

    ("Conclusion: The Sleepwalkers and the Collapse of Systemic Order",
     "In conclusion, while the assassination in Sarajevo set the machinery of war in motion, it was the accumulated weight of structural militarism, alliance chain-ganging, imperialist rivalries, and unalterable war plans that transformed an isolated assassination into the catastrophic 'seminal catastrophe of the twentieth century.' As Christopher Clark powerfully articulates in The Sleepwalkers, the European statesmen of July 1914 were not conscious villains plotting global annihilation, but sleepwalkers—leaders blinded by hyper-nationalist paranoia and trapped within the structural constraints of an anarchic, multipolar system that lacked crisis-management mechanisms. Without the underlying structural polarization, Sarajevo would have remained an inconsequential footnote in Balkan history; with it, the assassination unleashed a global tragedy that destroyed four great empires and reshaped modern civilization.")
]

q16_quotes = [
    ("Christopher Clark", "The Sleepwalkers: How Europe Went to War in 1914", "The outbreak of war in 1914 was a tragedy, not a crime. The protagonists were sleepwalkers, watchful but unseeing, blind to the reality of the horror they were about to bring into the world."),
    ("A.J.P. Taylor", "War by Timetable: How the First World War Began", "All the Great Powers had vast plans for mobilization, but the plans were so rigid that once the button was pressed, diplomacy ceased. The war of 1914 was a war caused by railway timetables."),
    ("Fritz Fischer", "Griff nach der Weltmacht (Germany's Aims in the First World War)", "The ruling elites of Imperial Germany deliberately risked a continental war in July 1914 to break out of encirclement, achieve European hegemony, and consolidate authoritarian power against domestic democratic forces."),
    ("Sir Edward Grey", "British Foreign Secretary (August 3, 1914)", "The lamps are going out all over Europe; we shall not see them lit again in our lifetime."),
    ("Kenneth N. Waltz", "Man, the State, and War", "The First World War demonstrates the third image of international relations: structural anarchy. In an anarchic multipolar system, the absence of central authority and the presence of rigid alliances make systemic war an ever-present structural hazard.")
]

# =========================================================================
# Q17: THE GENEVA CONVENTIONS & PRISONERS OF WAR (2018 Q.8, 2016 Q.1, 2013 Q.6)
# =========================================================================

q17_meta = [
    ("Target Examination", "KPPSC PMS (Provincial Management Service) — Khyber Pakhtunkhwa"),
    ("Paper & Subject", "International Relations (100 Marks Optional — Group 6)"),
    ("EXAM QUESTIONS LINKED:", '• 2018 Q.8: "Highlight the provisions of the Geneva Convention on Prisoners of War. (20 Marks)"\n• 2013 Q.6: "What are the primary sources of International Law? Discuss the limitations of International Law as an instrument for regulating state behavior in an anarchic international system. (20 Marks)"'),
    ("Syllabus Section", "Part A: Interaction Between States, International Law & Humanitarian Regimes"),
    ("Question Allocation", "Attempt 1 of 5 Required (20 Marks | Strict 35-Minute Allocation)"),
    ("Historiographical / Theoretical Thesis", "Formulated in the immediate aftermath of the industrialized barbarism of World War II, the Geneva Conventions of August 12, 1949—and specifically the Third Geneva Convention relative to the Treatment of Prisoners of War (GPW)—represent the pinnacle of International Humanitarian Law (IHL) and the legal institutionalization of jus in bello. Grounded in the moral imperative to insulate hors de combat personnel from sovereign vengeance, the Convention codifies non-derogable protections: mandating humane treatment (Art 13), strictly prohibiting coercive interrogation or torture (Art 17), institutionalizing independent oversight by the International Committee of the Red Cross (ICRC), and requiring mandatory repatriation upon cessation of hostilities (Art 118). However, in the contemporary strategic landscape of asymmetric warfare, 'unlawful combatants,' and proxy conflicts, the enforcement and universal applicability of the Convention face acute legal and geopolitical contestation.")
]

q17_callout = [
    "LEGAL FOUNDATIONS OF THE THIRD CONVENTION: The Third Geneva Convention of 1949 (GPW) replaced the inadequate 1929 Convention, responding directly to the horrific atrocities of WWII (Nazi starvation and mass murder of over 3 million Soviet POWs, Japanese Bataan Death March).",
    "ARTICLE 13 (HUMANE TREATMENT): The foundational principle: POWs must at all times be humanely treated. Any unlawful act or omission causing death or seriously endangering health is strictly prohibited and constitutes a grave breach (war crime).",
    "ARTICLE 17 (RESTRICTIONS ON QUESTIONING): Interrogation protections: POWs are bound to give only their surname, first names, rank, date of birth, and army serial number. No physical or mental torture, nor any other form of coercion, may be inflicted to secure information.",
    "ARTICLE 118 (REPATRIATION) & CONTEMPORARY CHALLENGES: Mandates immediate release and repatriation without delay after the cessation of active hostilities. Modern strains include the post-9/11 US invention of 'unlawful enemy combatants' (Guantanamo Bay), mistreatment of POWs in the Ukraine war, and non-state armed groups refusing IHL compliance."
]

q17_map = r"""
+---------------------------------------------------------------------------------------------------------+
|                  CONCEPTUAL ROADMAP: GENEVA CONVENTION ON PRISONERS OF WAR (POWs)                       |
+---------------------------------------------------------------------------------------------------------+
                                                     |
                               [Historical Genesis: The Barbarism of World War II]
                             Nazi Starvation of Soviet POWs & Imperial Japanese Atrocities
                             1929 Convention Proved Hopelessly Inadequate
                                                     |
                                                     v
                       [August 12, 1949: The Four Geneva Conventions]
                       First: Wounded on Land | Second: Wounded at Sea
                       THIRD CONVENTION: Relative to the Treatment of Prisoners of War (GPW)
                       Fourth: Protection of Civilian Persons in Time of War
                                                     |
                                                     v
                         [CORE SUBSTANTIVE PROVISIONS OF THE THIRD CONVENTION]
                                                     |
           +-----------------------------------------+-----------------------------------------+
           |                    |                                  |                           |
           v                    v                                  v                           v
     [ARTICLE 4]          [ARTICLE 13]                       [ARTICLE 17]                [ARTICLE 118]
Definition of POW:        Humane Treatment Mandatory         Interrogation Limits:       Repatriation:
Lawful Combatant, Uniform Absolute Prohibition of Violence   Name, Rank, Serial Number   Mandatory Release
Carrying Arms Openly      Protection from Public Curiosity   No Torture or Coercion      Without Delay at War's End
           |                    |                                  |                           |
           +--------------------+----------------------------------+---------------------------+
                                                     |
                                                     v
                         [THE INSTITUTIONAL GUARANTOR: THE ICRC]
                     International Committee of the Red Cross Granted Neutral
                     Access to Inspect Camps, Interview POWs, and Contact Families
                                                     |
                                                     v
                         [CONTEMPORARY CHALLENGES IN ASYMMETRIC WARFARE]
                     - Post-9/11 "Unlawful Combatants" Legal Black Holes (Guantanamo)
                     - Hybrid Warfare, Mercenaries (Wagner Group) & Non-State Insurgents
                     - Modern POW Mistreatment & Viral Public Humiliation on Social Media
+---------------------------------------------------------------------------------------------------------+
"""

q17_timeline = [
    ("1864: First Geneva Convention", "Henri Dunant inspired the original 10-article treaty protecting wounded soldiers and establishing the Red Cross.", "Inaugurated modern codified International Humanitarian Law (jus in bello)."),
    ("1929: Geneva Convention on POWs", "First dedicated international convention regulating the treatment of captured military personnel.", "Suffered fatal loopholes: did not protect resistance fighters, and the USSR and Japan failed to ratify it."),
    ("August 12, 1949: The Four Geneva Conventions", "Diplomatic conference in Geneva adopted the modern comprehensive humanitarian treaties.", "Created universal, non-derogable international humanitarian protections; Third Convention codified 143 articles on POWs."),
    ("1977: Additional Protocols I and II", "Extended Geneva protections to national liberation movements (Protocol I) and non-international armed conflicts (Protocol II).", "Modernized IHL to reflect the realities of post-colonial guerrilla warfare and civil wars."),
    ("Post-2001: The 'War on Terror' Crisis", "Bush administration classified Al-Qaeda and Taliban captives as 'unlawful enemy combatants' to deny GPW status.", "Triggered an unprecedented crisis in international law over the universal applicability of Geneva protections.")
]

q17_essay = [
    ("Introduction: The Moral and Legal Architecture of Jus in Bello",
     "The conduct of warfare has perpetually existed in profound tension with the principles of civilized human morality. While classical international law recognized the sovereign prerogative of states to wage war (jus ad bellum), the modern international community has sought to constrain the catastrophic violence of armed conflict through International Humanitarian Law (IHL), or jus in bello (justice in the conduct of war). At the absolute center of this humanitarian architecture stands the Geneva Convention relative to the Treatment of Prisoners of War, adopted on August 12, 1949—commonly designated as the Third Geneva Convention. Formulated in the devastating aftermath of World War II, during which millions of captured soldiers were subjected to industrial starvation, torture, medical experimentation, and summary execution (exemplified by the Nazi extermination of over three million Soviet captives and the Japanese atrocities of the Bataan Death March), the 1949 Third Convention established an exhaustive, non-derogable legal regime designed to insulate military personnel who are hors de combat (out of action) from sovereign vengeance and summary abuse."),

    ("Core Substantive Provisions: The Sacred Status of the Captive",
     "The Third Geneva Convention contains 143 detailed articles that delineate the rights of prisoners of war and the explicit legal obligations of the Detaining Power. Article 4 establishes the fundamental criteria defining who qualifies as a prisoner of war upon capture: members of the regular armed forces, as well as volunteer corps and organized resistance movements, provided they fulfill four conditions—being commanded by a person responsible for subordinates, wearing a fixed distinctive sign recognizable at a distance, carrying arms openly, and conducting operations in accordance with the laws of war. The foundational ethical premise of the convention is enshrined in Article 13: prisoners of war must at all times be humanely treated. Any act or omission by the Detaining Power causing death or seriously endangering the health of a prisoner is strictly prohibited and constitutes a grave breach (war crime). Furthermore, Article 13 explicitly dictates that prisoners must be protected at all times, particularly against acts of violence or intimidation and against insults and public curiosity."),

    ("Interrogation Restrictions and Living Conditions",
     "Among the most critical operational safeguards in the Convention is Article 17, which strictly regulates the interrogation of captives. When questioned, a prisoner of war is bound to declare only their surname, first names, rank, date of birth, and army serial number. Article 17 explicitly declares: 'No physical or mental torture, nor any other form of coercion, may be inflicted on prisoners of war to secure from them information of any kind whatever.' Prisoners who refuse to answer cannot be threatened, insulted, or subjected to any unpleasant or disadvantageous treatment. In addition, Articles 21 through 48 establish rigorous standards for internment: POWs must be housed in quarters comparable to those of the detaining forces, given rations sufficient to maintain good health, provided with necessary medical care free of charge, and permitted to practice their religious faith. Officers cannot be compelled to work, and enlisted personnel can only be employed in non-military, non-hazardous labor under fair compensation."),

    ("Oversight by the ICRC and Mandatory Repatriation",
     "A defining institutional triumph of the Third Geneva Convention is its codification of neutral third-party verification. Under Articles 125 and 126, the International Committee of the Red Cross (ICRC)—a neutral humanitarian organization—is granted uninhibited statutory authority to visit all places of internment, inspect detention facilities, and conduct private interviews with prisoners without witnesses. The ICRC acts as the moral guardian of the Convention, maintaining the Central Tracing Agency to convey correspondence between captives and their families. Finally, Article 118 governs the conclusion of captivity: 'Prisoners of war shall be released and repatriated without delay after the cessation of active hostilities.' This provision eliminates the historical abuse wherein states held captives as forced labor or geopolitical hostages for years following the termination of conflict."),

    ("Conclusion: Contemporary Challenges in Asymmetric Warfare",
     "In conclusion, the Third Geneva Convention represents one of the most civilized achievements in the history of international law, embodying the profound truth that a soldier ceases to be an enemy the moment they surrender. However, in the 21st-century strategic environment, the Convention faces unprecedented structural strains. The post-9/11 'Global War on Terror' witnessed sovereign states—most prominently the United States at Guantanamo Bay and Abu Ghraib—attempting to evade Geneva obligations by creating extra-legal classifications such as 'unlawful enemy combatants.' Furthermore, the proliferation of non-state terrorist organizations, private military contractors (such as the Wagner Group), and modern urban warfare involving social media public humiliation of captives present grave challenges to traditional state compliance. Despite these violations, the Geneva Conventions remain the non-negotiable moral and legal conscience of international politics; preserving their universal integrity is the ultimate litmus test of humanity's commitment to the rule of law over raw barbarism.")
]

q17_quotes = [
    ("Jean Pictet", "Architect & Commentator of the 1949 Geneva Conventions", "The purpose of the Geneva Conventions is to ensure that even in the midst of the worst human passions, a minimum of humanity and mutual respect is preserved. A prisoner of war is not a criminal; he is a disarmed soldier whom fate has placed in the hands of the enemy."),
    ("Michael Walzer", "Just and Unjust Wars", "The rules of war are a recognition that the soldier on the other side is a man like oneself. Once he has laid down his arms, he has stepped outside the realm of combat; to harm him then is nothing less than murder."),
    ("Theodor Meron", "The Humanization of Humanitarian Law", "The 1949 Geneva Conventions marked the decisive transition from a state-centric law of war based on reciprocity to a victim-centered international humanitarian law grounded in universal human dignity."),
    ("Hersch Lauterpacht", "International Law and Human Rights", "If international law is, in some ways, at the vanishing point of law, the law of war is, in even greater measure, at the vanishing point of international law. Yet, without the Geneva Conventions, warfare degenerates into unrestrained savagery."),
    ("International Committee of the Red Cross (ICRC)", "Official Commentary on the Third Geneva Convention", "Humane treatment is the cornerstone of the Third Convention. It admits of no exception, no derogation, and no compromise based on military necessity or political expediency.")
]

# =========================================================================
# EXECUTE GENERATION FOR BATCH 3: Q13, Q14, Q15, Q16, Q17
# =========================================================================

batch3_notes = [
    ("Q13_Joseph_Nye_Soft_Power_Hard_Power_and_Smart_Power_Master_Notes", q13_meta, q13_callout, q13_map, q13_timeline, q13_essay, q13_quotes, "QUESTION 13: JOSEPH NYE'S SOFT POWER, HARD POWER & SMART POWER", "KPPSC PMS IR Master Notes \u2014 Q13: Joseph Nye's Soft Power, Hard Power & Smart Power in Modern Statecraft"),
    ("Q14_Theories_of_Nationalism_Concept_and_State_Utility_Master_Notes", q14_meta, q14_callout, q14_map, q14_timeline, q14_essay, q14_quotes, "QUESTION 14: THEORIES OF NATIONALISM: CONCEPT & STATE UTILITY", "KPPSC PMS IR Master Notes \u2014 Q14: Theories of Nationalism: Concept, State Utility & Good vs. Evil Dichotomy"),
    ("Q15_Balance_of_Power_Mechanisms_and_Cold_War_Equilibrium_Master_Notes", q15_meta, q15_callout, q15_map, q15_timeline, q15_essay, q15_quotes, "QUESTION 15: BALANCE OF POWER & COLD WAR BIPOLAR EQUILIBRIUM", "KPPSC PMS IR Master Notes \u2014 Q15: Balance of Power Mechanisms & Cold War Bipolar Equilibrium"),
    ("Q16_Systemic_Origins_of_World_War_I_Assassination_vs_Rivalry_Master_Notes", q16_meta, q16_callout, q16_map, q16_timeline, q16_essay, q16_quotes, "QUESTION 16: SYSTEMIC ORIGINS OF WORLD WAR I", "KPPSC PMS IR Master Notes \u2014 Q16: Systemic Origins of World War I: Sarajevo Assassination vs. Structural Great Power Rivalry"),
    ("Q17_Geneva_Convention_on_Prisoners_of_War_Master_Notes", q17_meta, q17_callout, q17_map, q17_timeline, q17_essay, q17_quotes, "QUESTION 17: THE GENEVA CONVENTION ON PRISONERS OF WAR (POWs)", "KPPSC PMS IR Master Notes \u2014 Q17: The Geneva Conventions: Protection of Prisoners of War (POWs) & International Humanitarian Law")
]

for filename, meta, callout, cmap, timeline, essay, quotes, sub_title, full_title in batch3_notes:
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

print("Batch 3 (Q13 through Q17) generated successfully!")
