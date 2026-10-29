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
# Q7: THE UKRAINE WAR & EUROPEAN UNITY (2022 Q.1)
# =========================================================================

q7_meta = [
    ("Target Examination", "KPPSC PMS (Provincial Management Service) — Khyber Pakhtunkhwa"),
    ("Paper & Subject", "International Relations (100 Marks Optional — Group 6)"),
    ("EXAM QUESTIONS LINKED:", '• 2022 Q.1: "Critically examine if the Ukraine War is a local affair or a competition between the two superpowers, and how it is to impact European unity. Give arguments and examples to justify your answer. (20 Marks)"'),
    ("Syllabus Section", "Part B: Major Conflicts | Part C: US-Russia Rivalry & European Security"),
    ("Question Allocation", "Attempt 1 of 5 Required (20 Marks | Strict 35-Minute Allocation)"),
    ("Historiographical / Theoretical Thesis", "The 2022 Ukraine War is fundamentally an acute systemic confrontation between two superpowers rather than a localized territorial dispute, rooted in the structural clash between NATO's eastward expansion and Russia's near-abroad sphere-of-influence doctrine (Mearsheimer). While the conflict initially acted as a powerful centripetal catalyst for European unity—manifested in historic defense spending (Zeitenwende), massive joint weapons funding (European Peace Facility), and rapid energy decoupling from Russian hydrocarbons—it has simultaneously generated profound centrifugal fractures, exacerbating asymmetric economic burdens across the continent, deepening East-West political polarization, and entrenching Europe's strategic dependency on American military primacy.")
]

q7_callout = [
    "SUPERPOWER PROXY REALITY: The conflict features direct great-power stakes: US military aid exceeding $100B, real-time NATO battlefront intelligence, and explicit strategic goals to impose 'strategic defeat' on Russia, countered by Moscow's strategic alignments with China, Iran, and North Korea.",
    "REALIST VS LIBERAL THEORETICAL DICHOTOMY: Realism (Mearsheimer) explains the invasion as an inevitable, predictable reaction to NATO expansion violating Russia's core security perimeter; Liberal Institutionalism (Ikenberry) views it as an autocratic assault on the rules-based international order met by collective institutional solidarity.",
    "CENTRIPETAL EUROPEAN ALIGNMENT: Unprecedented short-term integration: 14+ sanction packages, defense budget surges (Germany's €100B fund), Nordic NATO expansion (Finland and Sweden), and fast-tracked EU candidate status for Ukraine.",
    "CENTRIFUGAL STRUCTURAL CRACKS: Persistent long-term vulnerabilities: Hungarian/Slovakian resistance to sanctions, widening East-West threat perception gaps, deindustrialization risks from high energy costs, and acute anxiety over American commitment reliability under shifting US administrations."
]

q7_map = r"""
+---------------------------------------------------------------------------------------------------------+
|                  CONCEPTUAL ROADMAP: THE UKRAINE WAR & EUROPEAN UNITY                                   |
+---------------------------------------------------------------------------------------------------------+
                                                     |
                               [Structural Genesis: Post-Cold War Order]
                             NATO Expansion Eastward (1999, 2004, 2008 Bucharest)
                             Russian Sphere-of-Influence Doctrine & Security Dilemma
                                                     |
                                                     v
                       [February 2022: Full-Scale Russian Invasion]
                       Local Flashpoint Inverts into Global Superpower Proxy Contest
                       US/NATO Proxy Support ($100B+ Arms) vs Sino-Russian Axis
                                                     |
           +-----------------------------------------+-----------------------------------------+
           |                                         |                                         |
           v                                         v                                         v
   [CENTRIPETAL FORCES]                      [CENTRIFUGAL FRACTURES]                   [SYSTEMIC DEPENDENCY]
Historic Defense Hikes (Zeitenwende)       Asymmetric Inflation & Energy Shocks      Erosion of EU "Strategic Autonomy"
European Peace Facility Arms Procurement   East-West Perception Gap (Poland vs Hung) Critical Reliance on US Nuclear &
Energy Decoupling from Russian Gas         Rise of Populist / Anti-Aid Factions      Conventional Deterrence Umbrella
Nordic NATO Entry (Finland & Sweden)       Deindustrialization Strains in Germany    Vulnerability to US Political Shifts
           |                                         |                                         |
           +-----------------------------------------+-----------------------------------------+
                                                     |
                                                     v
                         [THE PRECARIOUS OUTCOME: CRISIS-DRIVEN COHESION]
                     Europe Achieved Immediate Geopolitical Awakening, but Remains
                     Structurally Divided and Militarily Subordinated to Washington
+---------------------------------------------------------------------------------------------------------+
"""

q7_timeline = [
    ("2008: Bucharest Summit Declaration", "NATO declared that Ukraine and Georgia 'will become members of NATO.'", "Triggered Moscow's existential security dilemma, initiating active revanchism in Georgia and Ukraine."),
    ("February 2022: Full-Scale Invasion", "Russian military launched multi-axis invasion targeting Kyiv and Eastern Ukraine.", "Terminated the post-Cold War European security architecture; galvanized immediate transatlantic collective defense."),
    ("March–May 2022: European Geopolitical Awakening", "Germany announced €100B Zeitenwende; EU activated European Peace Facility to finance weapons.", "Marked the historic transformation of the European Union from a pure normative/economic bloc into a hard-power actor."),
    ("2023: Nordic NATO Expansion", "Finland officially joined NATO (doubling NATO-Russia border), followed by Sweden.", "Completely transformed the Baltic Sea into a 'NATO lake,' demonstrating decisive structural balancing against Russia."),
    ("2024–Present: Strains & Fracture Phase", "Hungary, Slovakia push back on aid packages; energy subsidies strain German/French industrial competitiveness.", "Highlighted the limits of crisis-driven unity, exposing long-term economic exhaustion and strategic divergence.")
]

q7_essay = [
    ("Introduction: Framing the Geopolitical Nature of the Ukraine Conflict",
     "The full-scale Russian invasion of Ukraine that commenced on February 24, 2022, represents the most catastrophic armed conflict on the European continent since the collapse of the Third Reich in 1945. While superficially presenting as an interstate border war between two post-Soviet neighbors, the conflict defies simplistic categorization as a localized dispute. Evaluated through the conceptual lenses of structural realism and international political economy, the Ukraine War is fundamentally an acute systemic confrontation between two nuclear-armed superpowers—the United States and the Russian Federation. Furthermore, the war has acted as a profound historical crucible for European unity. By violently shattering the foundational assumptions of the post-Cold War European security architecture, the conflict has generated a paradoxical dialectic: catalyzing unprecedented short-term institutional cohesion while concurrently deepening long-term structural fractures across the European project."),

    ("The Realist Dimension: A Superpower Competition in Local Dress",
     "The proposition that the Ukraine War is a great power rivalry is substantiated by structural realism. In his seminal analysis, John Mearsheimer argues that great powers are acutely sensitive to shifts in the balance of power along their geographical peripheries. From Moscow's vantage point, the eastward expansion of NATO—culminating in the fateful 2008 Bucharest Summit declaration that Ukraine would eventually become a member—posed an intolerable existential security threat. Conversely, for the United States, supporting Ukraine is not merely an act of normative solidarity; it is an exercise in grand-strategic balancing designed to impose 'strategic defeat' and permanent military attrition upon a major revanchist competitor. With Washington and its NATO allies providing over $100 billion in advanced armaments, real-time satellite intelligence, and battlefield operational planning, Ukraine serves as the primary battlefront of a wider proxy contest. This global dimension is further reinforced by Moscow's deepened strategic alignments with China, Iran, and North Korea, solidifying the conflict as a bipolar struggle over the rules governing the international system."),

    ("Centripetal Forces: Crisis-Driven European Integration",
     "For the European Union, the shock of the invasion initially produced an extraordinary surge of political solidarity and institutional integration. Confronted by an unprovoked violation of territorial sovereignty, European states enacted historic policy reversals. Germany abandoned its decades-long policy of Ostpolitik and military restraint, announcing a historic €100 billion defense fund (Zeitenwende). For the first time in its history, the EU financed the procurement and delivery of lethal weapons to a third party through the European Peace Facility. Simultaneously, Brussels enacted the REPowerEU roadmap, rapidly severing decades of economic reliance on cheap Russian piped gas through LNG diversification and green transition investments. The strategic perimeter of Europe was decisively consolidated as historically neutral Finland and Sweden abandoned non-alignment to join NATO, while the EU granted official candidate status to Ukraine and Moldova, demonstrating a unified geopolitical posture."),

    ("Centrifugal Fractures: Asymmetric Burdens and Strategic Vulnerabilities",
     "However, beneath this impressive surface consensus lie deep-seated, simmering fractures that threaten long-term European stability. First, the war has created severe economic asymmetries: Eastern and Southern European states have suffered disproportionately from runaway inflation, while Germany's heavy manufacturing sector faces deindustrialization due to the loss of inexpensive Russian energy. Second, sharp political divisions have emerged, exemplified by Hungary and Slovakia using their veto power to delay financial aid and sanctions packages, reflecting divergent threat perceptions between frontline Baltic/Polish hawks and skeptical Central/Southern members. Most critically, the war has exposed the illusion of European 'strategic autonomy' championed by French President Emmanuel Macron. Despite surging defense spending, European armies remain fundamentally dependent on American intelligence, logistics, and nuclear deterrence. The terrifying prospect of shifting political commitments in Washington has exposed Europe's acute strategic vulnerability."),

    ("Conclusion: The Precarious Future of European Unity",
     "In conclusion, the Ukraine War is definitively a competition between great powers played out on Ukrainian soil, carrying existential ramifications for global order. For European unity, the war has demonstrated that external security shocks can temporarily accelerate integration, unifying the continent around defense mobilization and energy sovereignty. Yet, as the conflict grinds into a war of attrition, the sustainability of this cohesion is increasingly fragile. If the European Union cannot translate its crisis-driven unity into genuine strategic autonomy, independent military capability, and shared economic burden-sharing, the centrifugal forces of inflation, political polarization, and geopolitical exhaustion will inevitably widen the fissures within Europe, leaving it deeply divided and permanently subordinated to great power realpolitik.")
]

q7_quotes = [
    ("John J. Mearsheimer", "Offensive Structural Realism", "The West, and especially America, is principally responsible for the crisis which began in February 2014. The taproot of the trouble is NATO enlargement, the central element of a larger strategy to move Ukraine out of Russia's orbit and integrate it into the West."),
    ("Stephen M. Walt", "Balance of Threat Theory", "The Ukraine war demonstrates that geography and proximity matter intensely. The European reaction was unified not because of abstract international law, but because Russian tanks on European borders represent a direct, immediate structural threat."),
    ("G. John Ikenberry", "Liberal Leviathan", "Putin's war in Ukraine was an attempt to drag Europe back to an imperial sphere-of-influence world. Its effect was the opposite: it revived the Western liberal democratic coalition, reinvigorating NATO and providing the EU with geopolitical purpose."),
    ("Barry Buzan", "Regional Security Complex Theory", "The Ukraine war illustrates the intense entanglement between a global superpower rivalry and a regional European security complex. The war has effectively welded transatlantic security into a single, highly militarized frontline."),
    ("Henry Kissinger", "Diplomacy & World Order", "To treat Ukraine as an outpost of either side will destroy for decades any prospect of bringing Russia and the West into an international system. Ukraine should function as a bridge between East and West, not a battleground for supremacy.")
]

# =========================================================================
# Q8: US INDO-PACIFIC STRATEGY & SOUTH ASIA (2022 Q.2)
# =========================================================================

q8_meta = [
    ("Target Examination", "KPPSC PMS (Provincial Management Service) — Khyber Pakhtunkhwa"),
    ("Paper & Subject", "International Relations (100 Marks Optional — Group 6)"),
    ("EXAM QUESTIONS LINKED:", '• 2022 Q.2: "Discuss the salient features of the US Indo-Pacific Strategy enunciated in February 2022. Analyze its implications for the South Asian Security Environment with particular reference to Pakistan. (20 Marks)"\n• 2016 Q.5: "Critically evaluate the foreign policy of the United States toward South Asia post-9/11. How has Washington\'s strategic partnership with India influenced the regional balance of power? (20 Marks)"\n• 2013 Q.8(d): "Write short note on: Geopolitical rivalry in the Indian Ocean. (10 Marks)"'),
    ("Syllabus Section", "Part C: US Foreign Policy, Indo-Pacific Strategy & Pakistan's Security"),
    ("Question Allocation", "Attempt 1 of 5 Required (20 Marks | Strict 35-Minute Allocation)"),
    ("Historiographical / Theoretical Thesis", "Enunciated by the Biden administration in February 2022, the US Indo-Pacific Strategy represents Washington's comprehensive grand-strategic framework designed to sustain American primacy by constructing minilateral counter-balancing networks against the People's Republic of China. By designating India as the indispensable regional 'net security provider' and the anchor of the Quad, the strategy has profoundly destabilized the South Asian strategic environment. For Pakistan, the institutionalization of US-India foundational defense agreements (BECA, COMCASA, LEMOA) and advanced dual-use military transfers severely erodes regional conventional and nuclear deterrence parity, compounds Islamabad's two-front security dilemma, and inexorably accelerates Pakistan's strategic reliance on Beijing.")
]

q8_callout = [
    "STRATEGIC FOUNDATIONS: Published in Feb 2022, the strategy focuses on five core pillars: advancing a free and open Indo-Pacific, building connections within and beyond the region, driving regional prosperity, bolstering deterrence, and building resilience to transnational threats.",
    "THE ELEVATION OF INDIA AS COUNTERWEIGHT: The strategy explicitly recognizes India as an indispensable partner, committing Washington to support India's rise and capacity to serve as a 'net security provider' in the Indian Ocean to balance China.",
    "EROSION OF SOUTH ASIAN STRATEGIC STABILITY: Advanced US technology transfers (Predator drones, jet engine co-production) and real-time geospatial intelligence sharing under BECA degrade Pakistan's conventional deterrence, compressing nuclear thresholds and fueling an escalatory arms race.",
    "PAKISTAN'S STRATEGIC SQUEEZE: Washington views South Asia almost exclusively through the prism of strategic competition with China; this marginalizes Pakistan's security concerns regarding India, intensifies CPEC vulnerability, and narrows Pakistan's non-aligned diplomatic maneuvering space."
]

q8_map = r"""
+---------------------------------------------------------------------------------------------------------+
|                  CONCEPTUAL ROADMAP: US INDO-PACIFIC STRATEGY & SOUTH ASIA                              |
+---------------------------------------------------------------------------------------------------------+
                                                     |
                               [Systemic Imperative: Balancing China's Rise]
                             US Pivot to Asia (Obama) --> Free & Open Indo-Pacific (Trump)
                             Biden Indo-Pacific Strategy (Feb 2022) Codifies Primacy
                                                     |
                                                     v
                       [The Architecture: Minilateral Balancing & Tech Alliances]
                       Quad (US, India, Japan, Australia) + AUKUS (US, UK, Australia)
                       Elevation of India: Major Defense Partner & Net Security Provider
                                                     |
                                                     v
                         [Institutionalizing US-India Strategic Interoperability]
                       Foundational Accords: LEMOA (Logistics), COMCASA (Comms), BECA (Geo-Spatial)
                       Critical Tech Transfers: GE F414 Jet Engines, MQ-9B Predator Drones
                                                     |
           +-----------------------------------------+-----------------------------------------+
           |                                         |                                         |
           v                                         v                                         v
 [CONVENTIONAL ASYMMETRY]                  [NUCLEAR DETERRENCE COLLAPSE]             [DIPLOMATIC MARGINALIZATION]
India Achieves Real-Time ISR Superiority   Precision Strike Capabilities Threaten    Pakistan Viewed Solely via China Lens
Anti-Access / Area-Denial in IOR           First-Strike Preemption Fears             Pressure on CPEC & Gwadar Port
Naval Modernization Overwhelms Pak Navy    Compression of Decision-Making Thresholds Polarization of South Asian Security
Compels Modernization of Submarines        Induction of Tactical Nuclear Weapons     Zero-Sum Binary Strategic Blocs
           |                                         |                                         |
           +-----------------------------------------+-----------------------------------------+
                                                     |
                                                     v
                         [THE REGIONAL OUTCOME: BIPOLAR STRATEGIC CLEAVAGE]
                     Deepening of the US-India Axis Counters the Sino-Pak Strategic
                     Partnership, Solidifying a Polarized, Crisis-Prone South Asia
+---------------------------------------------------------------------------------------------------------+
"""

q8_timeline = [
    ("2016: US Designates India Major Defense Partner", "US Congress formalized India's status, enabling defense technology transfers comparable to NATO allies.", "Marked the fundamental structural pivot in US South Asian policy toward explicit balancing against China."),
    ("2016–2020: Foundational Defense Pacts Signed", "US and India signed LEMOA (logistics), COMCASA (secure communications), and BECA (geospatial intelligence).", "Created seamless military interoperability, granting Indian armed forces real-time US satellite targeting data."),
    ("February 2022: US Indo-Pacific Strategy Released", "Biden administration formally published the strategy, defining India as the primary net security anchor.", "Codified the Indo-Pacific as the preeminent geopolitical theater of 21st-century great-power competition."),
    ("2023: US-India Initiative on Critical Tech (iCET)", "Bilateral agreements signed for GE jet engine co-production, semiconductor manufacturing, and defense tech.", "Accelerated the qualitative modernization of the Indian military-industrial complex."),
    ("Present: South Asian Strategic Repercussions", "Pakistan counters conventional asymmetry through HQ-9 air defense, Hangor submarines, and Shaheen/Babur missiles.", "Solidified a dangerous, polarized regional security complex characterized by razor-thin crisis escalation thresholds.")
]

q8_essay = [
    ("Introduction: The Genesis and Architecture of the US Indo-Pacific Strategy",
     "Enunciated by the Biden administration in February 2022, the United States Indo-Pacific Strategy represents the culmination of a decade-long grand-strategic reorientation aimed at arresting the decline of American hegemony in the face of China's economic and military ascent. Extending from the eastern coast of Africa to the western shores of the Americas, the concept of the 'Indo-Pacific' deliberately supplanted the traditional geographic nomenclature of the 'Asia-Pacific,' specifically to integrate the Indian Ocean into a unified geopolitical theater of containment. Grounded in the realist logic of external balancing, the strategy outlines five foundational objectives: advancing a free and open region, building connections within and beyond, driving economic prosperity through the Indo-Pacific Economic Framework (IPEF), bolstering collective deterrence, and reinforcing resilience against transnational threats. Central to this architecture is the construction of overlapping, minilateral security arrangements—most notably the Quadrilateral Security Dialogue (Quad)—and the strategic elevation of the Republic of India as Washington's preeminent regional anchor."),

    ("Salient Features: The Strategic and Military Elevation of India",
     "The centerpiece of the 2022 strategy is its unambiguous commitment to propel India's rise as a global power and a 'net security provider' in the Indian Ocean. In an explicit bid to establish an external balance of power against the People's Liberation Army Navy (PLAN), Washington has facilitated an unprecedented qualitative leap in defense and technological cooperation with New Delhi. Through the operationalization of four foundational defense agreements—the General Security of Military Information Agreement (GSOMIA), the Logistics Exchange Memorandum of Agreement (LEMOA), the Communications Compatibility and Security Agreement (COMCASA), and the Basic Exchange and Cooperation Agreement (BECA)—the US military has established profound interoperability with Indian armed forces. Furthermore, initiatives such as the US-India Initiative on Critical and Emerging Technologies (iCET) have provided India with sensitive dual-use military technologies, including the co-production of GE-414 jet engines, advanced semiconductor partnerships, and armed MQ-9B SeaGuardian drones."),

    ("Implications for South Asian Strategic Stability: The Erosion of Conventional Parity",
     "While calibrated to contain China, the influx of sophisticated American military hardware and real-time intelligence assets into India has profoundly destabilized the South Asian security environment, with direct, alarming implications for Pakistan. In strategic studies, deterrence stability rests upon the mutual conviction that neither side can execute a decisive, disarming first strike. However, the provision of US real-time geospatial intelligence under BECA significantly enhances India's precision-strike capabilities, providing Indian missile forces with pinpoint targeting coordinates against Pakistan's strategic assets. This technological asymmetry fuels dangerous doctrinal shifts within New Delhi, empowering notions of counter-force preemption and 'Cold Start' limited conventional strikes beneath Pakistan's nuclear threshold. The conventional military gap between the two nuclear-armed neighbors is widening into an unbridgeable chasm, severely degrading regional strategic equilibrium."),

    ("Pakistan's Security Dilemma: Compressing Nuclear Thresholds and Diplomatic Margins",
     "For Pakistan, the Indo-Pacific Strategy exacerbates an acute security dilemma. First, by eroding conventional deterrence parity, it forces Islamabad to lower its nuclear threshold, placing heavier strategic reliance on Full Spectrum Deterrence (FSD) and tactical nuclear weapons (such as the Nasr/Hatf-IX missile) to deter Indian conventional incursions. Second, Washington's single-minded prioritization of India has led to the structural marginalization of Pakistan in American foreign policy. Islamabad is increasingly viewed through the narrow, zero-sum prism of Sino-US rivalry, resulting in diplomatic pressure against Chinese investments under the China-Pakistan Economic Corridor (CPEC). Consequently, Pakistan's historical pursuit of strategic autonomy has been severely constrained, inexorably driving Islamabad into a tighter, deeper strategic, military, and economic alignment with Beijing."),

    ("Conclusion: The Peril of a Polarized Regional Security Matrix",
     "In conclusion, the US Indo-Pacific Strategy has permanently transformed the geopolitical landscape of South Asia, transforming a regional rivalry into a frontline theater of global superpower competition. By heavily arming India to counter China while disregarding the collateral impact on Pakistan's security imperatives, Washington has institutionalized an escalatory arms race and dismantled regional crisis management mechanisms. To navigate this hostile strategic environment, Pakistan must reject binary Cold War bloc alignment, modernize its conventional naval and electronic warfare capabilities (through Hangor-class submarines and air-defense integration), and pursue proactive, multidirectional diplomacy across the Global South, reminding the international community that a South Asian order engineered on deep strategic asymmetry is inherently fragile and acutely prone to catastrophic escalation.")
]

q8_quotes = [
    ("Ashley J. Tellis", "Carnegie Endowment for International Peace", "The US Indo-Pacific strategy is built on the bet that India's growing power will naturally balance China. But Washington must recognize that building up Indian military capacity inevitably alters the military balance with Pakistan, increasing instability in South Asia."),
    ("Stephen M. Walt", "Balance of Threat Theory", "The US-India partnership is a textbook example of balance-of-threat balancing. However, when great powers intervene in regional rivalries to balance distant competitors, they invariably heighten the regional security dilemma for intermediate powers like Pakistan."),
    ("C. Raja Mohan", "Indian Strategic Affairs", "The Indo-Pacific framework has broken the hyphenation of India and Pakistan in Washington's strategic calculus. India is now treated as a systemic actor in the global balance of power, while Pakistan has been relegated to a secondary regional consideration."),
    ("Ambassador Munir Akram", "Permanent Representative of Pakistan to the UN", "The massive transfer of advanced conventional and nuclear-capable technology to India under the guise of the Indo-Pacific strategy has demolished the strategic balance in South Asia, forcing Pakistan to rely on robust asymmetric deterrence."),
    ("John J. Mearsheimer", "The Tragedy of Great Power Politics", "In an anarchic system, the United States will do whatever it takes to contain a rising China, including empowering regional allies like India. The collateral damage to regional stability in South Asia is a price Washington is entirely willing to accept.")
]

# =========================================================================
# Q9: US-CHINA GEOECONOMIC RIVALRY & PAKISTAN'S PREDICAMENT (2022 Q.3)
# =========================================================================

q9_meta = [
    ("Target Examination", "KPPSC PMS (Provincial Management Service) — Khyber Pakhtunkhwa"),
    ("Paper & Subject", "International Relations (100 Marks Optional — Group 6)"),
    ("EXAM QUESTIONS LINKED:", '• 2022 Q.3: "The trade war between the US and China has jeopardized the global economic order impacting all regions and states; Pakistan caught between these two global powers is likely to be affected the most. Analyze Pakistan\'s predicament and suggest policy options to safeguard its national security objectives. (20 Marks)"\n• 2018 Q.6: "There may be a possible trade war between the USA and China that would rattle certain American businesses and force consumers to pay more for products. Discuss the term \\"trade war\\" and elaborate on why Trump was targeting China. (20 Marks)"'),
    ("Syllabus Section", "Part A: International Trade & Sanctions | Part C: US-China Rivalry & Pakistan"),
    ("Question Allocation", "Attempt 1 of 5 Required (20 Marks | Strict 35-Minute Allocation)"),
    ("Historiographical / Theoretical Thesis", "The trade and technology war between the United States and China represents a systemic struggle over the geo-economic architecture of the 21st century, defined by Henry Farrell and Abraham Newman as 'weaponized interdependence.' Through aggressive tariff barriers, export restrictions on advanced semiconductors (CHIPS Act), and financial sanctions, Washington seeks to decouple or 'de-risk' from Chinese supply chains. Caught precariously between these two global economic titans, Pakistan confronts a profound national security predicament: its macroeconomic solvency, export revenues, and multilateral financing remain critically tied to the Western financial sphere (IMF, US, EU), while its infrastructural modernization, energy security, and defense procurement are anchored to Beijing (CPEC). To survive, Pakistan must pursue rigorous non-camp hedging and geoeconomic revitalization.")
]

q9_callout = [
    "STRUCTURAL NATURE OF GEOECONOMIC RIVALRY: Transcends simple trade deficits; represents a comprehensive technological and financial conflict encompassing semiconductor supply chains, artificial intelligence, quantum computing, and alternative global reserve clearing currencies.",
    "WEAPONIZED INTERDEPENDENCE: Farrell & Newman's framework explains how centralized global hubs (SWIFT, US dollar clearing, advanced semiconductor fabrication) are mobilized as coercive geopolitical chokepoints against economic competitors.",
    "PAKISTAN'S ASYMMETRIC DUAL EXPOSURE: Pakistan's external economic survival requires Western cooperation (over 60% of exports flow to US/Europe, and IMF bailouts require US Treasury approval), yet its long-term infrastructural growth and military supply lines are unconditionally dependent on China ($30B+ CPEC liabilities).",
    "POLICY OPTIONS FOR NATIONAL SECURITY: Avoid binary alliance entrapment; operationalize CPEC Special Economic Zones (SEZs) for export diversification; accelerate trade integration with Central Asia (TAPI, CASA-1000); and leverage geoeconomics to position Pakistan as an indispensable connectivity corridor rather than a geopolitical frontline."
]

q9_map = r"""
+---------------------------------------------------------------------------------------------------------+
|                  CONCEPTUAL ROADMAP: US-CHINA TRADE WAR & PAKISTAN'S PREDICAMENT                        |
+---------------------------------------------------------------------------------------------------------+
                                                     |
                               [Systemic Genesis: The Geoeconomic Clash]
                             Trump Tariffs (2018) --> Biden's CHIPS Act & Decoupling (2022)
                             Weaponized Interdependence: Financial & Tech Chokepoints
                                                     |
                                                     v
                       [Bifurcation of Global Economic & Technological Regimes]
                       US Dollar / SWIFT Hegemony vs Chinese CIPS & Digital Yuan
                       Western 'De-Risking' vs China's Dual-Circulation & Belt and Road
                                                     |
                                                     v
                         [PAKISTAN'S ACUTE STRATEGIC & ECONOMIC PREDICAMENT]
                                  A Nation Caught in the Middle
                                                     |
           +-----------------------------------------+-----------------------------------------+
           |                                         |                                         |
           v                                         v                                         v
   [WESTERN LEVERAGE & RISKS]                [CHINESE ANCHOR & LIABILITIES]            [POLICY OPTIONS ROADMAP]
Over 60% of Exports to US & EU             $30B+ in CPEC Debt & Energy Payments      Strict Non-Alignment: No Camp Politics
Vulnerability to IMF Bailout Votes         Primary Supplier of High-Tech Weapons     Operationalize CPEC Industrial SEZs
FATF Compliance & FATF Monitoring          Vital Geopolitical Shield at UNSC         Expand Trade with Central Asia / SCO
Dependency on Western Capital Markets      Overland Energy & Connectivity Lifeline   Digital & Agricultural Export Modernization
           |                                         |                                         |
           +-----------------------------------------+-----------------------------------------+
                                                     |
                                                     v
                         [THE NATIONAL SECURITY IMPERATIVE: ACTIVE HEDGING]
                     Transforming Pakistan from a Geopolitical Pivot of Superpower
                     Friction into a Resilient Geoeconomic Hub of Interconnectivity
+---------------------------------------------------------------------------------------------------------+
"""

q9_timeline = [
    ("2018: Inception of US-China Tariff War", "Trump administration imposed tariffs on over $350B of Chinese imports under Section 301.", "Ended decades of economic integration, establishing trade warfare as an instrument of great-power competition."),
    ("2019–2020: Targeting Huawei & 5G Infrastructure", "US placed Huawei on Entity List, pressuring global allies to ban Chinese 5G telecommunications.", "Marked the formal transition from trade deficits to tech-nationalism and digital sovereignty contestation."),
    ("October 2022: The Biden CHIPS & Science Act", "Enacted sweeping export bans on advanced semiconductor lithography and AI chips to China.", "Jake Sullivan declared a 'small yard, high fence' doctrine, seeking to permanently suppress Chinese advanced computing."),
    ("2023: Western Pivot to 'De-Risking'", "G7 Hiroshima Summit adopted 'de-risking rather than decoupling' rhetoric.", "Formalized the strategic diversification of critical mineral and tech supply chains away from Chinese monopolies."),
    ("Present: Pakistan's Macroeconomic Squeeze", "IMF bailout reviews heavily scrutinize Chinese energy debt repayment and CPEC financial flows.", "Exposed the acute financial vulnerability of intermediate developing states caught between competing superpowers.")
]

q9_essay = [
    ("Introduction: The Geoeconomic Transformation of Great Power Rivalry",
     "The escalating trade and economic confrontation between the United States and the People's Republic of China represents far more than an ordinary dispute over bilateral trade imbalances or currency valuations. In the lexicon of contemporary international political economy, it constitutes a structural conflict over the control of the technological, financial, and supply-chain foundations of the 21st-century international order. Conceptualized by political scientists Henry Farrell and Abraham Newman as 'weaponized interdependence,' the globalized economic network has been transformed into a battleground where sovereign states leverage physical and digital network hubs for coercive geopolitical advantage. From the sweeping tariffs initiated in 2018 to the aggressive export controls of the 2022 CHIPS and Science Act, Washington has sought to 'de-risk' from Chinese manufacturing and constrain Beijing's technological ascendance. This systemic economic bifurcation has sent shockwaves through global markets, placing developing intermediate powers in profound jeopardy. Among these, Pakistan occupies the most precarious position—caught in an asymmetric vice between its vital Western financial lifelines and its existential strategic alignment with China."),

    ("The Mechanics of Geoeconomic Conflict: Weaponized Interdependence",
     "The instruments of the contemporary US-China economic clash are multi-dimensional. Washington has increasingly deployed its structural financial power—the unrivaled dominance of the US dollar in foreign exchange reserves and the Western oversight of the SWIFT interbank messaging system—to impose extraterritorial secondary sanctions. In the high-technology arena, the Biden administration's 'small yard, high fence' strategy explicitly prohibits the export of advanced semiconductor lithography, electronic design automation software, and artificial intelligence hardware to Chinese enterprises. In retaliation, Beijing has accelerated its 'dual-circulation' economic strategy, prioritizing domestic technological self-reliance while mobilizing its massive Belt and Road Initiative (BRI) network to secure alternative supply lines and promote de-dollarization through the Cross-Border Interbank Payment System (CIPS) and bilateral local-currency swaps."),

    ("Pakistan's Dual Predicament: Caught Between Two Titans",
     "For Pakistan, this geopolitical polarization strikes directly at the heart of national security. Pakistan suffers from an acute, structural dual-dependence. On one side stands Pakistan's macroeconomic solvency and commercial viability, which remain tethered to the Western financial sphere. Over 60% of Pakistan's total exports—predominantly textiles—are absorbed by consumer markets in the United States and the European Union. More critically, Pakistan's perennial balance-of-payments crises necessitate continuous bailout packages from the International Monetary Fund (IMF) and the World Bank, where Washington exercises decisive voting power. Conversely, on the other side stands Pakistan's long-term strategic, military, and infrastructural survival, which is unconditionally anchored to China. Under the $62 billion China-Pakistan Economic Corridor (CPEC), Beijing has modernized Pakistan's highway networks, expanded Gwadar Port, and alleviated catastrophic national energy deficits. Furthermore, China provides over 75% of Pakistan's advanced military hardware. When Washington pressures Islamabad to restrict Chinese investments, or when IMF negotiations scrutinize payments to Chinese Independent Power Producers (IPPs), Pakistan faces an existential squeeze."),

    ("Policy Options to Safeguard Pakistan's National Security",
     "To navigate this treacherous landscape, Pakistan must abandon obsolete Cold War paradigms and execute a sophisticated policy of geoeconomic hedging. First, Pakistan must uphold an absolute principle of non-camp alignment, rejecting any diplomatic or economic pressure to join exclusive anti-China or anti-Western coalitions. Second, Islamabad must operationalize CPEC Phase-II, swiftly developing Special Economic Zones (SEZs) by offering competitive tax incentives to attract Chinese manufacturing firms seeking to relocate production outside China to evade US tariffs. Third, Pakistan must diversify its external trade architecture by aggressively pivoting toward regional connectivity—integrating with the Shanghai Cooperation Organization (SCO) and unlocking trade corridors with Central Asian republics (via the Trans-Afghan railway and CASA-1000). Finally, Pakistan must implement radical domestic macroeconomic reforms to boost domestic tax mobilization and export competitiveness, thereby weaning itself off perpetual, politically compromised IMF structural adjustment loans."),

    ("Conclusion: From Geopolitical Frontline to Geoeconomic Bridge",
     "In conclusion, the US-China trade war has irreversibly transformed the global economic order from an era of frictionless neoliberal integration into a fractured landscape of techno-nationalism and weaponized trade. For Pakistan, continuing to treat national security as an exclusively military-territorial domain is a recipe for fiscal collapse. National security in the 21st century is fundamentally geoeconomic security. By rigorously preserving its strategic partnership with Beijing while proactively safeguarding its commercial and institutional ties with the West, Pakistan can transform its geographic location from a vulnerable frontline of superpower proxy competition into an indispensable, resilient transit bridge connecting Eurasian energy with global maritime commerce.")
]

q9_quotes = [
    ("Henry Farrell & Abraham Newman", "Weaponized Interdependence", "Globalized networks do not just generate mutual vulnerability; they create topological asymmetry. States that control central network hubs can weaponize informational and financial clearing nodes to coerce competitors and enforce global compliance."),
    ("Dani Rodrik", "The Globalization Paradox", "Hyper-globalization has collapsed because it proved incompatible with national sovereignty and domestic democratic priorities. The US-China conflict represents the violent reassertion of national economic security over global efficiency."),
    ("Robert O. Keohane", "Power and Interdependence", "Asymmetry in interdependence is a prime source of power. When a smaller state's financial solvency depends on one bloc while its military survival depends on another, it experiences the most acute vulnerability in international politics."),
    ("Shahid Javed Burki", "Pakistan's Economic Dilemma", "Pakistan's greatest foreign policy challenge in the 21st century is managing its external economic balance: maintaining access to Western markets and international lenders while securing Chinese investments under CPEC."),
    ("Graham Allison", "Destined for War", "The trade war is not about trade; it is about who will dominate the technologies that will define global power in the coming century. Smaller states caught in this vortex must hedge ruthlessly or risk economic ruin.")
]

# =========================================================================
# Q10: NON-STATE ACTORS & REALIST STATE PRIMACY (2022 Q.4)
# =========================================================================

q10_meta = [
    ("Target Examination", "KPPSC PMS (Provincial Management Service) — Khyber Pakhtunkhwa"),
    ("Paper & Subject", "International Relations (100 Marks Optional — Group 6)"),
    ("EXAM QUESTIONS LINKED:", '• 2022 Q.4: "In the Realist prism, the \'State\' is the legitimate actor in the international system, but the emergence of Non-State Actors (NSAs) has greatly impacted the global order; Taliban are governing Afghanistan without recognition and Hezbollah is a major stakeholder in Lebanon. Analyze the impact of NSAs on the global strategic environment with solid arguments and valid examples. (20 Marks)"\n• 2018 Q.5: "Discuss the emergence of ISIS and its impact on national, regional, and international peace and security. (20 Marks)"'),
    ("Syllabus Section", "Part A: Theories of IR (Realism) & The Modern State System"),
    ("Question Allocation", "Attempt 1 of 5 Required (20 Marks | Strict 35-Minute Allocation)"),
    ("Historiographical / Theoretical Thesis", "The axiomatic core of Classical and Structural Realism (Morgenthau, Waltz) posits the sovereign territorial state as the exclusive, unitary actor in international politics, possessing an uncontested Weberian monopoly over the legitimate use of violence. However, the contemporary strategic landscape fundamentally refutes this state-centric orthodoxy through the proliferation of highly capable Non-State Actors (NSAs) wielding state-level military, political, and administrative power. As demonstrated by the Afghan Taliban exercising de facto sovereign governance in Afghanistan without international de jure recognition, and Hezbollah operating as an armed sub-state actor with a precision-guided missile arsenal superior to the Lebanese Armed Forces, NSAs have transformed asymmetric warfare and regional deterrence, exposing the profound analytical blind spots of traditional realism.")
]

q10_callout = [
    "REALIST STATE-CENTRIC ORTHODOXY: Realism treats the state as a 'billiard ball' in an anarchic table; non-state actors are dismissed as secondary epiphenomena that either serve state masters or operate on the margins of world politics.",
    "THE TALIBAN AS DE FACTO STATEHOOD: The Afghan Taliban's defeat of the US/NATO intervention and subsequent administration of Afghanistan proves that non-state insurgent groups can capture, control, and govern sovereign territory completely outside the recognized Westphalian diplomatic recognition architecture.",
    "HEZBOLLAH AS A HYBRID SUB-STATE POWERHOUSE: Operating within Lebanon, Hezbollah maintains an arsenal of over 150,000 rockets, autonomous intelligence services, and cross-border expeditionary military units that project power across Syria and project credible strategic deterrence against Israel, surpassing the formal Lebanese state.",
    "PARADIGM SHIFT TO PLURALISM: The rise of militarized, ideational, and economic NSAs validates Transnational Pluralism (Keohane & Nye) and 'New Wars' frameworks (Mary Kaldor), demonstrating that systemic power is increasingly diffused among non-state networks that defy territorial sovereignty."
]

q10_map = r"""
+---------------------------------------------------------------------------------------------------------+
|                  CONCEPTUAL ROADMAP: NON-STATE ACTORS & THE REALIST CRISIS                              |
+---------------------------------------------------------------------------------------------------------+
                                                     |
                               [Westphalian Orthodoxy: The Realist State Prism]
                             Max Weber: State Holds Monopoly on Legitimate Force
                             Kenneth Waltz: States Are the Only Units that Matter
                                                     |
                                                     v
                       [The Historical Shift: Proliferation of Transnational NSAs]
                       Decolonization, State Fragility, Ideological Mobilization,
                       and Globalization Erode the Sovereign Monopoly of Violence
                                                     |
                                                     v
                         [CASE STUDY 1: AFGHAN TALIBAN]   [CASE STUDY 2: HEZBOLLAH (LEBANON)]
                         - Defeated Western Coalition     - Sub-State Armed Leviathan
                         - Governs Without UN Recognition - Possesses 150,000+ Missiles
                         - Enforces Taxes, Laws, Borders  - Deploys Expeditionary Forces
                         - De Facto Sovereign Entity      - Deters Regional States (Israel)
                                                     |
           +-----------------------------------------+-----------------------------------------+
           |                                         |                                         |
           v                                         v                                         v
   [ASYMMETRIC WARFARE]                      [SUBVERSION OF SOVEREIGNTY]               [THEORETICAL CRISIS]
Drones, Cyber & Guerrilla Tactics          States Can No Longer Control Borders      Realism Cannot Explain Sub-State
Cheap Offense Beats Expensive Defense      Hybrid Governance Overlaps Weak States    Actors Exercising Balance of Power
Nullification of Conventional Primacy      Parallel Service Provision & Ideology     Pluralism (Keohane/Nye) Validated
           |                                         |                                         |
           +-----------------------------------------+-----------------------------------------+
                                                     |
                                                     v
                         [THE NEW STRATEGIC REALITY: DIFFUSED GLOBAL ORDER]
                     The International System is No Longer Purely Inter-State, but a
                     Complex Matrix of State and Non-State Actors Competing for Power
+---------------------------------------------------------------------------------------------------------+
"""

q10_timeline = [
    ("1648: Peace of Westphalia", "Established the principle of state sovereignty and the state's monopoly over organized violence.", "Formed the foundational structural bedrock of modern international law and classical realist theory."),
    ("1982: Emergence of Hezbollah", "Founded during the Lebanese Civil War and Israeli invasion with Iranian Revolutionary Guard backing.", "Pioneered the modern model of an armed hybrid non-state actor combining military resistance, social welfare, and political dominance."),
    ("1996–2001: First Taliban Regime in Kabul", "Taliban captured Kabul; recognized only by Pakistan, Saudi Arabia, and UAE.", "First major instance in modern history of a radical non-state religious movement capturing national state machinery."),
    ("August 2021: Taliban Re-Capture Afghanistan", "Taliban seized Kabul following US withdrawal; established the Islamic Emirate without diplomatic recognition.", "Decisively demonstrated that non-state insurgencies can outlast Western state-building projects and govern through de facto control."),
    ("October 2023–Present: Regional Multi-Front NSA Warfare", "Hezbollah, Houthis in Yemen, and Iraqi militias engaged in synchronized cross-border military strikes.", "Demonstrated that non-state actors now possess the strategic capacity to close global maritime chokepoints (Red Sea) and reshape regional war.")
]

q10_essay = [
    ("Introduction: The Realist Axiom of State Primacy Under Siege",
     "For over three centuries, the intellectual discourse of international relations has been dominated by the Westphalian tradition, codified theoretically in the Classical Realism of Hans Morgenthau and the Structural Neorealism of Kenneth Waltz. The cornerstone of the realist paradigm is the axiom that the sovereign, territorial state is the primary, unitary, and legitimate actor in the international system. Grounded in Max Weber's classic formulation, the state is defined by its institutional monopoly over the legitimate use of physical force within a bounded territory. Under structural realism, non-state actors (NSAs) are relegated to trivial anomalies—epiphenomena devoid of systemic impact that must ultimately bend to the material power of sovereign governments. However, the realities of the 21st-century strategic environment have severely eroded this state-centric orthodoxy. The proliferation of powerful, highly armed, and politically sophisticated non-state actors—most visibly demonstrated by the Afghan Taliban governing Afghanistan without international recognition and Hezbollah dominating the strategic architecture of Lebanon—proves that non-state entities have emerged as decisive drivers of global geopolitics."),

    ("Case Study 1: The Afghan Taliban and De Facto Sovereignty Without Recognition",
     "The trajectory of the Afghan Taliban provides a devastating empirical rebuttal to the realist assumption that legitimate international authority requires de jure diplomatic recognition. In August 2021, a non-state insurgent movement successfully outlasted a twenty-year counter-insurgency campaign waged by the world's preeminent military superpower, toppled a Western-financed republic, and assumed total territorial control over Afghanistan. Today, the Taliban operate an effective, centralized administrative state: collecting customs revenues, enforcing nationwide legal codes, suppressing rival armed factions (such as ISKP), and projecting cross-border coercion against neighboring Pakistan. They execute every core function of Westphalian statehood, yet they remain entirely unrecognized by the United Nations and the international community. The Taliban's triumph exposes a fatal flaw in realism: that sovereign authority is derived from material compliance, ideational legitimacy, and local asymmetric power, rather than formal diplomatic admission into the interstate society of states."),

    ("Case Study 2: Hezbollah as an Armed Sub-State Leviathan",
     "If the Taliban represent an insurgent group capturing a state, the Lebanese movement Hezbollah represents an even more complex challenge to realism: the sub-state actor that dwarfs the sovereign government within which it operates. Established in the early 1980s, Hezbollah has evolved into what military analysts describe as the world's most heavily armed non-state military force. Boasting an arsenal of over 150,000 precision-guided ballistic missiles, advanced drones, and hardened combat brigades with extensive expeditionary experience in Syria, Hezbollah possesses a military apparatus far superior to the formal Lebanese Armed Forces. Operating within a weak, fragmented Lebanese confessional state, Hezbollah maintains an autonomous telecommunications network, independent banking institutions, and comprehensive social welfare systems. Crucially, Hezbollah enacts sovereign foreign policy—waging cross-border wars and establishing mutual deterrence with the state of Israel entirely outside the constitutional authorization of the Lebanese government, demonstrating that balance-of-power politics is no longer an exclusive monopoly of states."),

    ("Asymmetric Warfare and the Diffusion of Systemic Power",
     "The rise of these actors illustrates a profound democratization and diffusion of destructive power across the international system. In earlier epochs, the staggering capital costs of heavy artillery, armor, and air power preserved the state's military monopoly. In the contemporary era, the proliferation of low-cost, high-impact technologies—precision anti-tank guided missiles (ATGMs), commercial drone swarms, improvised explosive devices, and encrypted cyber networks—allows non-state actors to wage asymmetric warfare that neutralizes the conventional superiority of modern state militaries (as seen also with the Houthi movement in Yemen paralyzing global Red Sea shipping). Furthermore, non-state actors operate with profound structural advantages: lacking fixed sovereign territory that can be easily targeted for deterrence, they exploit porous borders and complex human terrain, effectively insulating themselves from classical coercive statecraft."),

    ("Conclusion: Towards a Post-Westphalian International Paradigm",
     "In conclusion, the emergence of non-state actors as pivotal geopolitical forces has permanently fractured the realist billiard-ball model of international relations. While sovereign states undeniably remain the most formidable concentrations of material resources and legal authority, they no longer hold a monopoly over war, peace, or governance. The global order has transitioned toward what Robert Keohane and Joseph Nye termed complex transnational pluralism, and what Hedley Bull envisioned as a 'New Medievalism'—an overlapping matrix of authority where sovereign states, armed sub-state militias, transnational insurgent networks, and religious movements simultaneously contest power. Scholars and policymakers who continue to view international security exclusively through the narrow lens of state-to-state relations will remain structurally blind to the profound asymmetric forces shaping modern global politics.")
]

q10_quotes = [
    ("Robert O. Keohane & Joseph S. Nye", "Transnational Relations and World Politics", "Realism's exclusive focus on the nation-state obscures the dramatic rise of transnational and sub-state actors who bypass sovereign governments and directly alter the international distribution of power."),
    ("Kenneth N. Waltz", "Theory of International Politics", "States set the terms of intercourse, whether of peace or of war, in an anarchic realm. Non-state actors may matter at the margins, but so long as the major states remain the ultimate arbiters of violence, they define the structure of international politics."),
    ("Mary Kaldor", "New and Old Wars", "In the era of globalization, the state's monopoly of legitimate organized violence is eroding from above and below. 'New Wars' are fought by networks of state and non-state actors using identity politics, terror, and asymmetric guerrilla warfare."),
    ("Max Weber", "Politics as a Vocation (1919)", "A state is a human community that successfully claims the monopoly of the legitimate use of physical force within a given territory. When non-state actors seize this monopoly, the foundational concept of the modern state dissolves."),
    ("Fawaz Gerges", "The New Middle East", "Hezbollah is not a mere proxy or militia; it is a hybrid political-military leviathan that has mastered asymmetric deterrence, redefining sovereignty in the Middle East by subordinating the formal state to non-state strategic priorities.")
]

# =========================================================================
# Q11: CHINA'S BELT AND ROAD INITIATIVE (BRI) (2022 Q.5)
# =========================================================================

q11_meta = [
    ("Target Examination", "KPPSC PMS (Provincial Management Service) — Khyber Pakhtunkhwa"),
    ("Paper & Subject", "International Relations (100 Marks Optional — Group 6)"),
    ("EXAM QUESTIONS LINKED:", '• 2022 Q.5: "The Chinese Belt and Road Initiative (BRI) has global outreach in Asia, Africa, and Europe at the cost of US hegemony. Discuss its main objectives and analyze challenges and prospects of its success. (20 Marks)"'),
    ("Syllabus Section", "Part C: Major Powers' Foreign Policy (China) & Geo-economic Competition"),
    ("Question Allocation", "Attempt 1 of 5 Required (20 Marks | Strict 35-Minute Allocation)"),
    ("Historiographical / Theoretical Thesis", "Launched by President Xi Jinping in 2013, the Belt and Road Initiative (BRI) represents the most audacious transcontinental geo-economic and infrastructural grand strategy in modern history. Grounded in Halford Mackinder's Heartland theory and the logic of structural statecraft, the BRI seeks to physically integrate the Eurasian and Afro-Asian landmasses, solve China's acute 'Malacca Dilemma,' and absorb industrial overcapacity. By offering non-Western developing nations over $1 trillion in unconditional infrastructure finance, the BRI poses a profound structural challenge to US maritime hegemony and the Washington Consensus. However, its long-term prospects are severely complicated by escalating sovereign debt defaults among host states, acute localized security threats (such as Baloch insurgent attacks in Pakistan), and determined Western counter-balancing initiatives (G7 PGII, IMEC).")
]

q11_callout = [
    "GRAND-STRATEGIC RATIONALE: The BRI comprises two primary prongs: the overland 'Silk Road Economic Belt' connecting China to Europe via Central Asia, and the '21st Century Maritime Silk Road' securing sea lanes through the South China Sea, Indian Ocean, and Mediterranean.",
    "MACKINDER'S HEARTLAND LOGIC: By constructing transcontinental overland freight, pipeline, and fiber-optic networks, the BRI re-centers global geopolitical power back to the Eurasian Heartland, reducing China's vulnerability to US naval blockades at maritime chokepoints (the Malacca Dilemma).",
    "STRUCTURAL CHALLENGE TO US HEGEMONY: Provides an alternative development paradigm that rejects IMF/World Bank governance conditionalities, anchors 150+ signatory states to Chinese industrial standards, and promotes the internationalization of the Renminbi via local-currency trade.",
    "MAJOR HEADWINDS & IMPEDIMENTS: The initiative faces immense structural hurdles: accusations of 'debt-trap diplomacy' (exemplified by the 99-year lease of Sri Lanka's Hambantota port), post-pandemic debt renegotiations, rising anti-Chinese local hostility, and the US-led India-Middle East-Europe Economic Corridor (IMEC)."
]

q11_map = r"""
+---------------------------------------------------------------------------------------------------------+
|                  CONCEPTUAL ROADMAP: CHINA'S BELT AND ROAD INITIATIVE (BRI)                             |
+---------------------------------------------------------------------------------------------------------+
                                                     |
                               [Strategic Genesis: Domestic & Geopolitical Drivers]
                             Absorption of Industrial Overcapacity (Steel, Cement)
                             Overcoming the 'Malacca Dilemma' (US Naval Chokepoints)
                                                     |
                                                     v
                       [The Twin Pillars of Xi Jinping's 2013 Vision]
                       Overland: Silk Road Economic Belt (Eurasian Railway/Pipelines)
                       Maritime: 21st Century Maritime Silk Road (Ports & Sea Lanes)
                                                     |
                                                     v
                         [FLAGSHIP PROJECTS: THE REACH ACROSS 150+ NATIONS]
                       CPEC (Pakistan): Direct Access to Indian Ocean via Gwadar Port
                       China-Laos Railway, Jakarta-Bandung High-Speed Rail, Piraeus (Greece)
                                                     |
           +-----------------------------------------+-----------------------------------------+
           |                                         |                                         |
           v                                         v                                         v
   [CHALLENGE TO US PRIMACY]                 [STRUCTURAL IMPEDIMENTS]                  [WESTERN COUNTER-INITIATIVES]
Bypasses American Sea Power                Sovereign Debt Distress in Host Nations   G7 PGII: Infrastructure & Investment
Erosion of Bretton Woods Monopoly          Hambantota Syndrome & 'Debt-Trap' Fears   IMEC: India-Middle East-Europe Corridor
Promotion of Petroyuan Settlements         Severe Local Security Risks (TTP, BLA)    US 'Build Back Better World' (B3W)
Sovereign Non-Interference Appeals         Environmental & Social Backlash           Tech Decoupling & Supply Chain Onshoring
           |                                         |                                         |
           +-----------------------------------------+-----------------------------------------+
                                                     |
                                                     v
                         [THE FUTURE PROSPECT: FROM MEGAPROJECTS TO 'SMALL & BEAUTIFUL']
                     Transitioning to High-Tech Digital Silk Road, Green Energy,
                     and Debt Restructuring to Sustain Eurasian Hegemonic Influence
+---------------------------------------------------------------------------------------------------------+
"""

q11_timeline = [
    ("September 2013: Xi Jinping Launches BRI", "Speeches in Kazakhstan and Indonesia introduced the Silk Road Economic Belt and Maritime Silk Road.", "Inaugurated China's definitive pivot from Deng Xiaoping's 'hide and bide' to assertive global leadership."),
    ("2015: Inception of CPEC Flagship Corridor", "President Xi visited Islamabad, signing $46B (later expanded to $62B) in infrastructure and energy deals.", "Designated the China-Pakistan Economic Corridor as the undisputed crown jewel of the global initiative."),
    ("2017: Hambantota Port Debt Equity Swap", "Sri Lanka leased Hambantota Port to a Chinese state firm for 99 years to repay debt liabilities.", "Sparked intense global debate and Western condemnation over alleged Chinese 'debt-trap diplomacy.'"),
    ("2021–2023: Western Counter-Initiatives", "G7 launched PGII ($600B fund); G20 announced India-Middle East-Europe Economic Corridor (IMEC).", "Formalized great-power competition within global infrastructure financing and trade corridor architecture."),
    ("October 2023: Third Belt and Road Forum", "Beijing recalibrated BRI toward 'small is beautiful' green, digital, and technologically resilient projects.", "Marked the strategic maturation of BRI away from massive capital-intensive debt outlays toward sustainable influence.")
]

q11_essay = [
    ("Introduction: The Scale and Vision of the Belt and Road Initiative",
     "Unveiled by President Xi Jinping during landmark addresses in Kazakhstan and Indonesia in late 2013, the Belt and Road Initiative (BRI)—originally titled 'One Belt, One Road' (OBOR)—represents the most expansive global infrastructure and economic integration endeavor ever conceived by a single nation. Comprising two interlocking prongs—the overland 'Silk Road Economic Belt' connecting China to Western Europe across Central Asia, and the oceanic '21st Century Maritime Silk Road' linking China's coastal littoral to Africa, the Mediterranean, and Latin America—the initiative has expanded to encompass over 150 countries and international organizations, mobilizing over $1 trillion in investments. Viewed through the theoretical lens of structural realism and classical geopolitics, the BRI is fundamentally a grand-strategic enterprise designed to reconfigure global economic geometry, overcome vulnerabilities in China's security perimeter, and challenge the post-World War II hegemony of the United States."),

    ("Grand Strategic Objectives: Escaping the Malacca Dilemma and Exporting Capital",
     "The driving impulses behind the BRI are rooted in both domestic economic compulsions and existential geopolitical imperatives. Domestically, following decades of state-subsidized capital expansion, China faced acute industrial overcapacity in steel, cement, and engineering infrastructure; the BRI provided a massive transcontinental outlet for Chinese state-owned enterprises (SOEs). Geopolitically, the initiative is an explicit operationalization of Halford Mackinder's Heartland theory. Mackinder famously posited that control of the vast Eurasian landmass would ultimately allow a terrestrial power to dominate global affairs. Crucially, over 80% of China's imported crude oil traverses the narrow, 1.7-mile chokepoint of the Strait of Malacca, which is vulnerable to blockade by the United States Navy during a Taiwan crisis. By constructing overland pipeline and transit corridors—most notably the China-Pakistan Economic Corridor (CPEC) terminating at the warm-water port of Gwadar on the Arabian Sea, and the China-Myanmar pipeline to the Bay of Bengal—Beijing bypasses vulnerable maritime bottlenecks, insulating its economic supply lines from American sea power."),

    ("Challenging American Hegemony and the Bretton Woods Consensus",
     "The BRI directly contests US hegemony across institutional, normative, and financial dimensions. For eight decades, global infrastructure financing was monopolized by the Washington Consensus—embodied by the World Bank and the IMF—which demanded rigorous structural adjustment, fiscal austerity, privatization, and political democratization as prerequisites for lending. China dismantled this monopoly by offering immediate, capital-intensive infrastructure financing under the banner of 'South-South cooperation' and the Westphalian principle of non-interference in internal affairs. Furthermore, the BRI fosters institutional soft-balancing against Western dominance through the establishment of parallel multilateral bodies like the Asian Infrastructure Investment Bank (AIIB) and the promotion of the Renminbi in bilateral trade settlements, laying the groundwork for a post-dollar global financial architecture."),

    ("Challenges, Vulnerabilities, and Western Counter-Offensives",
     "Despite its colossal scale, the long-term viability of the BRI confronts formidable headwinds. First, the initiative has generated severe sovereign debt distress among borrowing nations. The controversial 2017 case of Sri Lanka, which handed over the strategic port of Hambantota on a 99-year lease to a Chinese state conglomerate after defaulting on debt obligations, crystallized international accusations of 'debt-trap diplomacy.' While scholars debate the intentionality of this trap, acute sovereign defaults in Zambia, Pakistan, and Kenya have compelled Beijing to act as an unprecedented global debt-restructuring creditor. Second, BRI projects face severe localized security threats, exemplified by Baloch insurgent and TTP suicide bombings targeting Chinese engineers in Pakistan. Third, the United States and its allies have launched coordinated counter-balancing initiatives, including the G7's Partnership for Global Infrastructure and Investment (PGII) and the India-Middle East-Europe Economic Corridor (IMEC), designed to provide developing nations with transparent, democratic financing alternatives."),

    ("Conclusion: The Evolution Toward 'Small and Beautiful' Statecraft",
     "In conclusion, the Belt and Road Initiative has decisively re-engineered the physical and political geography of Eurasia and the Global South, successfully establishing China as the undisputed infrastructure titan of the modern era. While it has not yet supplanted American global hegemony, it has definitively broken the Western monopoly on development finance and created an expansive, Sinocentric economic orbit. Recognizing the economic unsustainability of massive, debt-fueled megaprojects, Beijing used the 2023 Third Belt and Road Forum to pivot toward a 'small and beautiful' paradigm—prioritizing the 'Digital Silk Road' (5G networks, artificial intelligence, satellite navigation), green energy transitions, and high-tech supply chains. The future of the BRI will not be measured by the raw tonnage of poured concrete, but by its capacity to institutionalize an alternative, multipolar order that permanently circumvents Western containment.")
]

q11_quotes = [
    ("Halford Mackinder", "The Geographical Pivot of History (1904)", "Who rules East Europe commands the Heartland; who rules the Heartland commands the World-Island; who rules the World-Island commands the world. The BRI is the modern engineering realization of the unified Heartland."),
    ("Alfred Thayer Mahan", "The Influence of Sea Power upon History", "Control of the sea by maritime commerce and naval supremacy is the decisive factor in world dominion. China's BRI is a deliberate, historic effort to break free from the containment of Anglo-American sea power."),
    ("David Harvey", "The New Imperialism", "The Belt and Road Initiative is the ultimate 'spatial-temporal fix'—a massive geographic expansion designed to absorb surplus capital and industrial overcapacity to stave off internal economic stagnation."),
    ("Minxin Pei", "China's Crony Capitalism", "The BRI is an extraordinarily ambitious gamble. If China cannot manage the sovereign debt crises and geopolitical blowback from host countries, the initiative risks becoming a multitrillion-dollar imperial overstretch."),
    ("John J. Mearsheimer", "The Tragedy of Great Power Politics", "China is doing exactly what the United States did in the Western Hemisphere: using its colossal wealth to build regional hegemony and push rival powers far away from its borders. The BRI is power politics dressed in the language of development.")
]

# =========================================================================
# Q12: THE COVID PANDEMIC & THE 'NEW NORMAL' IN IR (2022 Q.7)
# =========================================================================

q12_meta = [
    ("Target Examination", "KPPSC PMS (Provincial Management Service) — Khyber Pakhtunkhwa"),
    ("Paper & Subject", "International Relations (100 Marks Optional — Group 6)"),
    ("EXAM QUESTIONS LINKED:", '• 2022 Q.7: "The Covid pandemic has introduced the concept of the \'New Normal\' in the psycho-social attitudes of the public across the globe. Evaluate its structural impact on international relations, economic globalization, and global governance. (20 Marks)"\n• 2018 Q.7: "There are two schools of thought regarding the conjecture of globalization. First, view it as a blessing and a tool for prosperity; the other views it as a tool of frustration. Please elaborate. (20 Marks)"'),
    ("Syllabus Section", "Part A: International Trade & Globalization | Non-Traditional Security"),
    ("Question Allocation", "Attempt 1 of 5 Required (20 Marks | Strict 35-Minute Allocation)"),
    ("Historiographical / Theoretical Thesis", "The Covid-19 pandemic operated as a profound systemic accelerator that shattered the post-Cold War orthodoxy of frictionless hyper-globalization and exposed the vulnerabilities of 'weaponized medical interdependence.' Evaluated through international political economy and non-traditional security frameworks, the pandemic did not foster global collective solidarity; instead, it precipitated the resurgence of the sovereign Westphalian state, triggered aggressive 'vaccine nationalism,' and paralyzed multilateral institutions (WHO) through intense US-China geopolitical blame games. The resulting 'New Normal' in international relations is defined by the securitization of public health, the strategic shift from supply-chain efficiency to 'friend-shoring' resilience, and the entrenchment of fragmented globalization.")
]

q12_callout = [
    "THE COLLAPSE OF HYPER-GLOBALIZATION: The pandemic shattered the assumption that market efficiency and complex, single-source just-in-time supply chains (concentrated in China) could guarantee national security during biological disruptions.",
    "THE RESURGENCE OF THE WESTPHALIAN LEVIATHAN: In the face of existential crisis, multilateral governance retreated; the sovereign state reasserted its uncontested primacy through border closures, export bans on personal protective equipment (PPE), and unilateral lockdowns.",
    "VACCINE NATIONALISM & MULTILATERAL PARALYSIS: Rather than functioning as a neutral global health arbiter, the World Health Organization (WHO) became paralyzed by the US-China geopolitical rivalry, while wealthy nations hoarded vaccines, leaving the Global South to rely on bilateral vaccine diplomacy.",
    "THE 'NEW NORMAL' PARADIGM: International political economy has permanently pivoted from 'Just-in-Time' cost minimization to 'Just-in-Case' strategic resilience: nearshoring, onshoring, and friend-shoring critical pharmaceutical, energy, and semiconductor supply chains."
]

q12_map = r"""
+---------------------------------------------------------------------------------------------------------+
|                  CONCEPTUAL ROADMAP: COVID-19, THE 'NEW NORMAL' & GLOBALIZATION                         |
+---------------------------------------------------------------------------------------------------------+
                                                     |
                               [The Pre-Pandemic Baseline: Hyper-Globalization]
                             Neoliberal Orthodoxy: Borderless Capital & Frictionless Trade
                             Single-Source Dependencies (China as the World's Factory)
                                                     |
                                                     v
                       [Spring 2020: The Systemic Shock & Biological Rupture]
                       Worldwide Lockdowns, Border Closures, and Logistics Collapse
                       Exposure of Weaponized Interdependence in Critical Medical Supplies
                                                     |
                                                     v
                         [THE RETREAT OF MULTILATERALISM & VACCINE NATIONALISM]
                       WHO Politicized & Paralyzed by US-China Geopolitical Clashes
                       Western Powers Hoard Vaccines; COVAX Multilateral Mechanism Fails
                       Bilateral Vaccine Diplomacy: China & Russia Fill the Vacuum
                                                     |
           +-----------------------------------------+-----------------------------------------+
           |                                         |                                         |
           v                                         v                                         v
   [GEOECONOMIC RESHORING]                   [WESTPHALIAN RESTORATION]                 [PSYCHO-SOCIAL REALIGNMENT]
End of 'Just-in-Time' Supply Chains        State Reasserts Total Border Control      Normalization of Digital Surveillance
Rise of 'Friend-Shoring' & Nearshoring     Health Becomes Hard National Security     Technological Remote Work Paradigms
Strategic Autonomy in Pharma & Tech        Decline of Liberal Institutionalism       Deepening Inequality Between North
Western Push to Diversify Away from China  National Sovereignty Trumps Global Norms   and South (Debt Distress)
           |                                         |                                         |
           +-----------------------------------------+-----------------------------------------+
                                                     |
                                                     v
                         [THE CONTEMPORARY OUTCOME: THE 'NEW NORMAL']
                     A Fragmented, Multiplex World: Globalization Does Not Disappear,
                     but Becomes Securitized, Regionalized, and Subordinated to State Power
+---------------------------------------------------------------------------------------------------------+
"""

q12_timeline = [
    ("December 2019 – January 2020: Outbreak", "Novel coronavirus detected in Wuhan; declared a Public Health Emergency of International Concern (PHEIC).", "Inaugurated the most severe global public health catastrophe since the 1918 Spanish Flu pandemic."),
    ("March 2020: Border Closures & Supply Panic", "WHO declared global pandemic; over 100 countries enacted border closures and export restrictions on PPE.", "Marked the immediate, simultaneous collapse of global civil aviation and supply-chain logistics."),
    ("Mid-2020: Geopolitical Weaponization", "Trump administration announced withdrawal from WHO; accused Beijing of weaponizing health data.", "Transformed a humanitarian health crisis into an intense front of superpower geopolitical rivalry."),
    ("2021: Vaccine Nationalism vs Diplomacy", "US and Europe prioritized domestic populations; China and Russia launched bilateral vaccine diplomacy in Global South.", "Exposed the failure of multilateral global public goods mechanisms like the COVAX facility."),
    ("Post-2022: Institutionalization of 'New Normal'", "G7 enacted supply-chain resilience policies; widespread adoption of digital surveillance and hybrid governance.", "Formalized the transition from open hyper-globalization to securitized, fragmented economic regionalism.")
]

q12_essay = [
    ("Introduction: The Pandemic as a Systemic Stress-Test for Global Order",
     "The outbreak of the Covid-19 pandemic in early 2020 represented far more than an extraordinary public health emergency; it functioned as a profound systemic stress-test that stripped away the foundational illusions of the post-Cold War international order. For three decades, the ideology of hyper-globalization—championed by neoliberal economists and celebrated in Thomas Friedman's thesis of a 'flat world'—presupposed that cross-border economic integration, market efficiency, and frictionless supply chains had permanently diminished the relevance of the sovereign nation-state. However, the virus struck directly at the nervous system of globalized interdependence. Rather than inspiring a renaissance of multilateral cooperation, the pandemic acted as an intense geopolitical accelerator, exposing deep supply-chain vulnerabilities, entrenching state-centric nationalism, and establishing the 'New Normal' in international relations: a fragmented global architecture where economic efficiency is permanently subordinated to national security, sovereign self-reliance, and strategic resilience."),

    ("The Exposure of Hyper-Globalization and Weaponized Interdependence",
     "The primary structural casualty of the pandemic was the doctrine of unconstrained supply-chain optimization. Under neoliberal capitalism, global corporations had constructed hyper-efficient 'Just-in-Time' manufacturing models that concentrated critical production nodes in single geographic locations—predominantly the People's Republic of China. When factories across East Asia shuttered in early 2020, Western states suffered catastrophic shortages of essential medical supplies, pharmaceuticals, and active pharmaceutical ingredients (APIs). In international political economy, this revealed the dangerous vulnerability of 'weaponized interdependence.' Western policymakers suddenly realized that over-reliance on a potential geopolitical adversary for life-saving antibiotics and medical equipment constituted an unacceptable national security threat. The resulting disruption catalyzed an urgent, systemic shift toward 'Just-in-Case' supply-chain architectures, driving policies of onshoring, nearshoring, and 'friend-shoring' critical industries within trusted democratic alliances."),

    ("The Resurgence of the Westphalian State and Vaccine Nationalism",
     "In the realm of global governance, Covid-19 triggered the dramatic resurgence of the sovereign Westphalian state at the expense of international institutions. In times of existential peril, populations turned not to the United Nations or the European Union, but to their national governments for survival. Sovereign states unilaterally closed borders, requisitioned private manufacturing facilities, and enacted sweeping export bans on ventilators and therapeutics. This nationalist retreat was epitomized by the phenomenon of 'vaccine nationalism.' Despite grand rhetoric celebrating vaccines as 'global public goods,' wealthy industrialized nations pre-purchased billions of doses, severely underfunding the multilateral COVAX facility and leaving developing countries in Africa and South Asia desperately underserved. In the absence of genuine multilateral solidarity, great powers deployed 'vaccine diplomacy' as a transactional geopolitical weapon, with China and Russia leveraging Sinopharm and Sputnik V deliveries to extract political concessions across the Global South."),

    ("Institutional Paralysis and Superpower Blame Games",
     "Simultaneously, the pandemic exposed the acute institutional fragility of multilateral organizations in an era of great-power contestation. The World Health Organization (WHO), theoretically designed to serve as a neutral, technical coordinator of global epidemiological defense, became thoroughly paralyzed by the vitriolic geopolitical struggle between the United States and China. The Trump administration accused the WHO of serving as a 'puppet of Beijing' and formally withdrew from the body, while Chinese authorities resisted independent international investigations into the origins of the virus in Wuhan. The politicization of international health governance demonstrated that global institutions cannot function autonomously when great powers decide to treat a humanitarian crisis as a zero-sum propaganda battleground."),

    ("Conclusion: The Architecture of the 'New Normal'",
     "In conclusion, the 'New Normal' ushered in by the Covid-19 pandemic does not signify the absolute death of globalization, but its profound, permanent qualitative mutation. Globalization has not ended; it has been securitized, politicized, and fractured into competing regional blocs. The psycho-social attitudes of global populations have normalized digital surveillance, state border tracking, and skepticism toward external dependencies. In the post-pandemic international order, states recognize that health security, supply-chain resilience, and technological sovereignty are indivisible elements of national power. The romantic era of frictionless, borderless global governance has vanished, replaced by a hardheaded, competitive Westphalian reality where sovereign states must build robust domestic buffers to survive in a turbulent, fragmented world.")
]

q12_quotes = [
    ("Dani Rodrik", "The Globalization Paradox", "The pandemic exposed the fatal flaw of hyper-globalization: the idea that markets can exist without governments. When crisis strikes, the nation-state remains the only legitimate entity capable of providing public protection and economic rescue."),
    ("Thomas L. Friedman", "The World is Flat", "The post-Covid world will be divided into B.C. (Before Corona) and A.C. (After Corona). We learned that a world that is merely connected without resilience is profoundly fragile."),
    ("Henry Kissinger", "The Corona Crisis Will Forever Alter the World Order", "The founding legend of modern government is a walled city protected by powerful rulers. The pandemic has prompted a resurgence of the walled city in an age when prosperity depends on global trade."),
    ("Francis Fukuyama", "The Thing That Determines a Country's Resilience", "The pandemic revealed that state capacity, social trust, and leadership are the ultimate variables of national survival. Ideological systems mattered less than whether a state possessed competent bureaucratic institutions."),
    ("Joseph E. Stiglitz", "Globalization and Its Discontents", "We constructed an economic system that was optimized for short-term corporate profits at the expense of long-term societal resilience. Covid-19 proved that efficiency without security is an existential catastrophe.")
]

# =========================================================================
# EXECUTE GENERATION FOR BATCH 2: Q7, Q8, Q9, Q10, Q11, Q12
# =========================================================================

batch2_notes = [
    ("Q7_The_Ukraine_War_Superpower_Rivalry_and_European_Unity_Master_Notes", q7_meta, q7_callout, q7_map, q7_timeline, q7_essay, q7_quotes, "QUESTION 7: THE UKRAINE WAR & EUROPEAN UNITY", "KPPSC PMS IR Master Notes \u2014 Q7: The Ukraine War: Superpower Rivalry & European Unity"),
    ("Q8_US_Indo_Pacific_Strategy_and_South_Asian_Security_Master_Notes", q8_meta, q8_callout, q8_map, q8_timeline, q8_essay, q8_quotes, "QUESTION 8: US INDO-PACIFIC STRATEGY & SOUTH ASIA", "KPPSC PMS IR Master Notes \u2014 Q8: US Indo-Pacific Strategy & South Asian Strategic Stability"),
    ("Q9_US_China_Geoeconomic_Rivalry_Trade_Wars_and_Pakistan_Master_Notes", q9_meta, q9_callout, q9_map, q9_timeline, q9_essay, q9_quotes, "QUESTION 9: US-CHINA TRADE WAR & PAKISTAN'S PREDICAMENT", "KPPSC PMS IR Master Notes \u2014 Q9: US-China Geoeconomic Rivalry & Pakistan's Strategic Predicament"),
    ("Q10_Non_State_Actors_and_The_Crisis_of_Realist_State_Primacy_Master_Notes", q10_meta, q10_callout, q10_map, q10_timeline, q10_essay, q10_quotes, "QUESTION 10: NON-STATE ACTORS & REALIST STATE PRIMACY", "KPPSC PMS IR Master Notes \u2014 Q10: Non-State Actors & The Crisis of Realist State Primacy"),
    ("Q11_China_Belt_and_Road_Initiative_and_US_Hegemony_Master_Notes", q11_meta, q11_callout, q11_map, q11_timeline, q11_essay, q11_quotes, "QUESTION 11: CHINA'S BELT & ROAD INITIATIVE VS US HEGEMONY", "KPPSC PMS IR Master Notes \u2014 Q11: China's Belt and Road Initiative (BRI) vs US Hegemony"),
    ("Q12_Covid_Pandemic_New_Normal_and_Globalization_Master_Notes", q12_meta, q12_callout, q12_map, q12_timeline, q12_essay, q12_quotes, "QUESTION 12: COVID PANDEMIC & THE 'NEW NORMAL' IN IR", "KPPSC PMS IR Master Notes \u2014 Q12: The Covid-19 Pandemic & The 'New Normal' in International Relations")
]

for filename, meta, callout, cmap, timeline, essay, quotes, sub_title, full_title in batch2_notes:
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

print("Batch 2 (Q7 through Q12) generated successfully!")
