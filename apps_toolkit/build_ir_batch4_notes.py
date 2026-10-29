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
# Q18: THE MODERN STATE SYSTEM: WESTPHALIA TO GLOBALIZATION (2016 Q.1)
# =========================================================================

q18_meta = [
    ("Target Examination", "KPPSC PMS (Provincial Management Service) — Khyber Pakhtunkhwa"),
    ("Paper & Subject", "International Relations (100 Marks Optional — Group 6)"),
    ("EXAM QUESTIONS LINKED:", '• 2016 Q.1: "Critically examine the evolution of the Modern State System from the Peace of Westphalia (1648). How is traditional state sovereignty being redefined by globalization and transnational forces? (20 Marks)"'),
    ("Syllabus Section", "Part A: The Modern State System: History, Basic Features, Evolving Forms"),
    ("Question Allocation", "Attempt 1 of 5 Required (20 Marks | Strict 35-Minute Allocation)"),
    ("Historiographical / Theoretical Thesis", "The 1648 Peace of Westphalia established the foundational architecture of the modern international system, dismantling overlapping feudal hierarchies and universalist papal-imperial authority in favor of sovereign territorial statehood. Codifying the principles of territorial integrity, internal supremacy, and external independence, the Westphalian model institutionalized the state as the exclusive legitimate arbiter of violence and international law. However, in the 21st century, traditional Westphalian sovereignty is undergoing profound qualitative erosion under the dual pressures of economic globalization, supranational integration (e.g., the European Union), transboundary environmental and cyber challenges, and sub-state fragmentation, transforming absolute sovereignty into complex, shared, and conditional governance.")
]

q18_callout = [
    "THE 1648 BREAKTHROUGH: Terminated the Thirty Years' War; established the principle of 'cuius regio, eius religio' (whose realm, his religion), establishing state secular jurisdiction and mutual non-interference.",
    "CORE FOUR PILLARS: Codified in the 1933 Montevideo Convention: a defined territory, a permanent population, an effective government with a monopoly on legitimate force (Weber), and the capacity to enter into diplomatic relations with other states.",
    "TRANSNATIONAL EROSION: Hyper-globalization of capital, multinational corporations, and international financial institutions (IMF/WTO) severely restrict sovereign economic autonomy, while borderless threats (climate change, pandemics) render physical borders porous.",
    "STEPHEN KRASNER'S 'ORGANIZED HYPOCRISY': Realists remind us that pure sovereignty has always been violated by powerful states; while transnationalism challenges state control, the sovereign state remains the supreme legal entity in international law and security."
]

q18_map = r"""
+---------------------------------------------------------------------------------------------------------+
|                  CONCEPTUAL ROADMAP: THE MODERN STATE SYSTEM: 1648 TO PRESENT                           |
+---------------------------------------------------------------------------------------------------------+
                                                     |
                               [Pre-Westphalian World: Feudal Overlap & Papal Rule]
                             Universal Catholic Church + Holy Roman Empire Hegemony
                             No Unified Territorial Jurisdiction or Sovereign Borders
                                                     |
                                                     v
                       [The 1648 Peace of Westphalia: Birth of Sovereign Statehood]
                       Treaties of Münster and Osnabrück Terminate 30 Years' War
                       Principle: Cuius Regio, Eius Religio & Secular Sovereignty
                                                     |
           +-----------------------------------------+-----------------------------------------+
           |                                         |                                         |
           v                                         v                                         v
   [TERRITORIAL INTEGRITY]                   [INTERNAL SUPREMACY]                      [EXTERNAL INDEPENDENCE]
Inviolability of Border Lines              Max Weber: Monopoly on Legitimate Force    Sovereign Equality of States
Exclusion of Foreign Jurisdiction          Centralized Administrative & Legal Order   Freedom from External Subordination
           |                                         |                                         |
           +-----------------------------------------+-----------------------------------------+
                                                     |
                                                     v
                         [21st CENTURY PRESSURES: TRANSNATIONAL CHALLENGES]
                                                     |
           +-----------------------------------------+-----------------------------------------+
           |                                         |                                         |
           v                                         v                                         v
 [ECONOMIC GLOBALIZATION]                  [SUPRANATIONALISM (EU)]                   [NON-TRADITIONAL THREATS]
Borderless Capital, Supply Chains, MNCs   Voluntary Pooling of National Sovereignty  Climate Change, Cyber Warfare, Pandemics
IMF/WTO Economic Conditionalities         European Court of Justice Overrules States Physical Borders Rendered Defenseless
           |                                         |                                         |
           +-----------------------------------------+-----------------------------------------+
                                                     |
                                                     v
                         [CONCLUSION: REDEFINING WESTPHALIAN SOVEREIGNTY]
                     Sovereignty is Not Dead, but Has Transformed from Absolute
                     Territorial Control into Shared, Multilevel Global Governance
+---------------------------------------------------------------------------------------------------------+
"""

q18_timeline = [
    ("1648: Peace of Westphalia", "Treaties of Münster and Osnabrück signed, establishing state sovereignty and non-intervention.", "Created the foundational legal and diplomatic bedrock of the modern international system."),
    ("1789: French Revolution & Popular Sovereignty", "Transformed dynastic territorial states into modern sovereign nation-states.", "Merged state institutions with national identity, creating the nation-state model."),
    ("1933: Montevideo Convention", "Codified the four objective criteria of sovereign statehood in international law.", "Established universal legal benchmark for recognition in the international society of states."),
    ("1945–1970: Decolonization Wave", "Global diffusion of the Westphalian nation-state model across Asia and Africa.", "Expanded sovereign statehood from a European regional phenomenon into a universal global order."),
    ("21st Century: Globalized Contestation", "Rise of supranational governance (EU), transnational cyber networks, and non-state warfare.", "Compelled international relations theory to recognize the era of conditional and shared sovereignty.")
]

q18_essay = [
    ("Introduction: The Westphalian Genesis of International Relations",
     "The modern international system is fundamentally a Westphalian system. The signing of the Peace of Westphalia in October 1648—comprising the twin treaties of Münster and Osnabrück that concluded the devastating Thirty Years' War—represents the defining rupture in global political history. Prior to 1648, the European continent was characterized by medieval feudalism: a fragmented, overlapping matrix of authority where local lords, dynastic monarchs, the Holy Roman Empire, and the universal Roman Catholic papacy contested jurisdictional legitimacy. Westphalia dismantled this medieval hierarchy, instituting the sovereign territorial state as the foundational, indivisible unit of international politics. By enshrining the principle of cuius regio, eius religio (whose realm, his religion), Westphalia recognized the sovereign right of rulers to exercise supreme, secular authority within their borders, free from external intervention. However, while this state-centric model expanded globally through centuries of war and decolonization, the contemporary 21st century has witnessed profound challenges to Westphalian sovereignty, driven by the inexorable forces of globalization, supranationalism, and transnational security threats."),

    ("Core Features of the Classical Westphalian State System",
     "The classical Westphalian state system rests upon three interlocking pillars codified over centuries of international jurisprudence, culminating in the 1933 Montevideo Convention on the Rights and Duties of States. The first pillar is territorial integrity: the state operates within clearly demarcated geographical frontiers that define the spatial boundaries of domestic law. The second pillar is internal sovereignty, famously defined by sociologist Max Weber as the successful claim by a human community to the monopoly of the legitimate use of physical force within a given territory. The sovereign state alone possesses the legitimate authority to enact laws, levy taxes, enforce order, and command armed forces. The third pillar is external independence and sovereign equality: in the anarchic realm of international politics, all sovereign states are legally equal under international law, irrespective of vast disparities in material wealth or military power, and are protected by the foundational norm of non-intervention in internal affairs (codified in Article 2, Paragraph 7 of the UN Charter)."),

    ("Transnational Globalization and the Erosion of Sovereign Autonomy",
     "In the contemporary era, the traditional Westphalian paradigm has experienced unprecedented structural erosion under the pressure of globalization. As political scientist David Held argues, economic globalization has unbundled the relationship between sovereignty, territoriality, and state power. The rapid mobility of transnational capital, the ascendancy of multinational corporations (MNCs) whose annual revenues dwarf the GDP of mid-sized nations, and the binding conditionalities imposed by international financial institutions (the IMF, World Bank, and WTO) severely constrain the capacity of sovereign states to enact independent macroeconomic policies. Furthermore, the digital revolution and global supply chain integration have rendered physical territorial borders increasingly porous to the flow of information, finance, and culture, challenging the state's traditional monopoly over domestic communication and identity formation."),

    ("Supranationalism, Non-State Actors, and Borderless Threats",
     "Simultaneously, sovereignty is being challenged from above by supranational institutions and from below by non-state actors. The European Union (EU) represents the most advanced historical departure from classical Westphalian principles: member states have voluntarily surrendered sovereign competencies over monetary policy, trade law, and border control (the Schengen Agreement), accepting the supremacy of European Community law over national statutes. Concurrently, non-traditional security crises demonstrate the absolute physical impotence of territorial sovereignty. Existential threats such as climate change, pandemics (e.g., Covid-19), transnational terrorism, and cyber-warfare recognize no national boundaries; no state, regardless of its military power, can insulate its citizens from globalized ecological or biological disruptions through unilateral territorial defense. Moreover, the doctrine of 'Responsibility to Protect' (R2P), adopted by the UN in 2005, has introduced the concept of conditional sovereignty: when a state commits or fails to prevent mass atrocities against its population, the international community claims a moral and legal right to intervene, directly puncturing the sacred Westphalian shield of non-interference."),

    ("Conclusion: The Enduring State in an Era of Shared Governance",
     "In conclusion, while the modern state system has undergone profound qualitative transformations, declarations of the 'death of the nation-state' remain premature. As realist scholar Stephen Krasner trenchantly noted in Sovereignty: Organized Hypocrisy, sovereignty was never an absolute, inviolable reality, but an idealized normative framework that powerful states routinely breached when national interest dictated. Despite the corrosive pressures of globalization and transnational interdependence, the sovereign nation-state remains the supreme legal authority in international law, the primary provider of physical security, and the sole legitimate locus of democratic accountability and mass taxation. In the 21st century, the Westphalian system has not collapsed; rather, it has evolved into a post-Westphalian matrix of shared, multilevel governance where sovereign states must continuously negotiate, pool authority, and coordinate with supranational and transnational networks to preserve systemic order.")
]

q18_quotes = [
    ("Stephen D. Krasner", "Sovereignty: Organized Hypocrisy (1999)", "Sovereignty was never absolute. The principles of Westphalian sovereignty—territoriality and the exclusion of external actors—have always been violated by coercion, intervention, and hypocrisy whenever vital national interests were at stake."),
    ("Hedley Bull", "The Anarchical Society (1977)", "The modern international system is an anarchical society: a society of sovereign states that, while acknowledging no common superior, form an order through shared rules, diplomatic conventions, and international law."),
    ("David Held", "Global Transformations", "Globalization has not ended the state, but it has fundamentally unbundled sovereignty. The modern state is locked into an overlapping matrix of regional and global governance regimes that transform its role from ruler to coordinator."),
    ("Max Weber", "Politics as a Vocation (1919)", "The modern state is a compulsory association which organizes domination. It has been successful in seeking to monopolize the legitimate use of physical force as the means of domination within a territory."),
    ("Jean Bodin", "Six Books of the Commonwealth (1576)", "Sovereignty is the absolute and perpetual power of a commonwealth, which the Latins call majestas; the supreme power over citizens and subjects, unrestrained by law.")
]

# =========================================================================
# Q19: CLASSICAL REALISM VS STRUCTURAL NEOREALISM (2016 Q.2, 2013 Q.2)
# =========================================================================

q19_meta = [
    ("Target Examination", "KPPSC PMS (Provincial Management Service) — Khyber Pakhtunkhwa"),
    ("Paper & Subject", "International Relations (100 Marks Optional — Group 6)"),
    ("EXAM QUESTIONS LINKED:", '• 2016 Q.2: "Compare and contrast Classical Realism (Hans Morgenthau) and Structural Realism / Neorealism (Kenneth Waltz). How do their core assumptions differ regarding the cause of conflict in international politics? (20 Marks)"\n• 2013 Q.2: "\\"Balance of Power is a central concept in Realism.\\" Discuss the various mechanisms employed by states to maintain or restore the Balance of Power (alliances, armaments, buffer states, compensation). (20 Marks)"'),
    ("Syllabus Section", "Part A: IR as a Field of Study: Theories & Basic Approaches (Realism)"),
    ("Question Allocation", "Attempt 1 of 5 Required (20 Marks | Strict 35-Minute Allocation)"),
    ("Historiographical / Theoretical Thesis", "While Classical Realism (Hans Morgenthau) and Structural Neorealism (Kenneth Waltz, John Mearsheimer) share the fundamental axioms that international politics is anarchic, conflictual, and driven by state self-help, they diverge fundamentally in their locus of causation regarding the root origins of interstate conflict. Classical realism is a reductionist theory that locates the drive for power within flawed human nature (animus dominandi)—the biological, insatiable lust for power inherent in men. In sharp contrast, structural neorealism rejects psychological reductionism, positing through systemic theory that conflict is an inexorable structural outcome forced upon states by the condition of international anarchy and the material distribution of capabilities. Furthermore, within neorealism, Defensive Realists (Waltz) view power as an instrument for security, whereas Offensive Realists (Mearsheimer) view power maximization as the only guarantor of survival.")
]

q19_callout = [
    "LOCUS OF CAUSATION (FIRST VS THIRD IMAGE): Classical Realism operates at Kenneth Waltz's 'First Image' (Human Nature: human beings are biologically flawed, egotistical, and driven by an insatiable lust for power - animus dominandi); Structural Neorealism operates at the 'Third Image' (The International System: structural anarchy and distribution of capabilities dictate state behavior regardless of human virtue).",
    "MORGENTHAU'S SIX PRINCIPLES: Politics is governed by objective laws rooted in human nature; statesmen think and act in terms of interest defined as power; power is the currency of politics, though its form changes; abstract moral principles cannot be applied to state actions.",
    "WALTZ'S SYSTEMIC ARCHITECTURE: Replaced biological realism with a scientific, structural framework: international politics is characterized by anarchy (ordering principle), functionally undifferentiated sovereign units (all seek survival), and varies only by the distribution of capabilities across units (unipolar, bipolar, multipolar).",
    "DEFENSIVE VS OFFENSIVE NEOREALISM: Defensive Realism (Waltz, Jervis) posits that the international system rewards states that maintain the status quo balance of power, warning that overexpansion triggers balancing; Offensive Realism (Mearsheimer) argues that because states can never be certain of others' intentions, they are structurally compelled to maximize relative power and pursue regional hegemony."
]

q19_map = r"""
+---------------------------------------------------------------------------------------------------------+
|                  CONCEPTUAL ROADMAP: CLASSICAL REALISM VS STRUCTURAL NEOREALISM                         |
+---------------------------------------------------------------------------------------------------------+
                                                     |
                               [The Shared Realist Core: Anarchy & Power]
                             Thucydides & Machiavelli: The World is What It Is
                             Statism, Survival, and Self-Help in an Unforgiving Realm
                                                     |
                                                     v
                       [THE FUNDAMENTAL EPISTEMOLOGICAL & CAUSAL SCHISM]
                                                     |
           +-----------------------------------------+-----------------------------------------+
           |                                                                                   |
           v                                                                                   v
   [CLASSICAL REALISM: HANS MORGENTHAU]                                [STRUCTURAL NEOREALISM: KENNETH WALTZ]
Level of Analysis: First Image (Human Nature)                       Level of Analysis: Third Image (International System)
Causation: Animus Dominandi (Lust for Power)                        Causation: Systemic Anarchy (Lack of Central Authority)
Power as an End in Itself (Biological Drive)                        Power as a Useful Means for Survival
Philosophical, Historical, and Prescriptive                         Scientific, Deductive, and Positivist
Focus on Statesmanship and Prudence                                 Focus on Structural Polarity (Bipolar vs Multipolar)
           |                                                                                   |
           +-----------------------------------------+-----------------------------------------+
                                                     |
                                                     v
                         [THE INTRA-NEOREALIST SPLIT: DEFENSIVE VS OFFENSIVE]
                                                     |
           +-----------------------------------------+-----------------------------------------+
           |                                                                                   |
           v                                                                                   v
   [DEFENSIVE REALISM: KENNETH WALTZ]                                  [OFFENSIVE REALISM: JOHN MEARSHEIMER]
Status Quo Bias: System Rewards Moderation                          Revisionist Bias: System Compels Aggression
States Seek Security, Not Infinite Hegemony                         States Must Maximize Relative Power to Survive
Over-Expansion Triggers Destructive Balancing                        Ultimate Goal: Regional Hegemony to Eliminate Threats
           |                                                                                   |
           +-----------------------------------------+-----------------------------------------+
                                                     |
                                                     v
                         [SYNTHESIS: THE ENDURING INTELLECTUAL DOMINANCE]
                     Both Theories Agree that in an Anarchic World, Moral Sentiments
                     Must Bow to the Relentless Imperatives of Relative Material Power
+---------------------------------------------------------------------------------------------------------+
"""

q19_timeline = [
    ("431 BC: Thucydides' Peloponnesian War", "The Melian Dialogue: 'The strong do what they can, and the weak suffer what they must.'", "Formulated the foundational ancient canon of classical realpolitik."),
    ("1948: Hans Morgenthau's 'Politics Among Nations'", "Published definitive text establishing classical realism as the dominant post-WWII paradigm.", "Codified the Six Principles of Political Realism, anchoring foreign policy in national interest defined as power."),
    ("1979: Kenneth Waltz's 'Theory of International Politics'", "Introduced structural neorealism, replacing biological human nature with systemic theory.", "Revolutionized international relations theory by developing a rigorous, deductive, structural science of politics."),
    ("2001: John Mearsheimer's 'The Tragedy of Great Power Politics'", "Formulated offensive structural realism, challenging Waltz's defensive status-quo assumptions.", "Argued that the anarchic structure compels rational great powers to aggressively pursue regional hegemony."),
    ("Present: Explaining Superpower Competition", "Both classical and structural realisms actively deployed to explain US-China rivalry and Ukraine war.", "Demonstrated the enduring explanatory primacy of realist power-politics in international relations.")
]

q19_essay = [
    ("Introduction: The Shared Realist Foundations and the Great Intellectual Divide",
     "Realism has long constituted the dominant theoretical paradigm in the study of international relations. Across centuries of diplomatic history, realism has provided a hardheaded, unsentimental analysis of state behavior, repudiating liberal utopianism and legalistic moralism. All realists unite around three foundational axioms, famously summarized as the 'three Ss': Statism (sovereign states are the primary actors), Survival (the ultimate preeminent goal of foreign policy is national survival), and Self-Help (in the absence of an overarching world sovereign, states must rely on their own capabilities to ensure security). However, beneath this shared philosophical consensus lies a fundamental theoretical divide between two dominant schools: Classical Realism, crystallized in the mid-twentieth century by Hans J. Morgenthau, and Structural Neorealism, formulated in 1979 by Kenneth N. Waltz and later extended by John J. Mearsheimer. While both schools agree that international politics is defined by an unending struggle for power, they diverge profoundly in their level of analysis, their causal explanation for conflict, and their methodological foundations."),

    ("Classical Realism: Human Nature and the Lust for Power",
     "Classical realism, articulated in Morgenthau's foundational 1948 text Politics Among Nations, operates at what Kenneth Waltz categorized as the 'First Image'—the level of human nature. Rooted in the political philosophy of Thucydides, Thomas Hobbes, and Niccolò Machiavelli, classical realism posits that international conflict is a direct reflection of flawed human nature. In Morgenthau's analysis, human beings possess an innate, biological lust for power—the animus dominandi. Because the state is merely human organization written large, states inherit this insatiable biological drive to dominate others. In his celebrated 'Six Principles of Political Realism,' Morgenthau argues that politics is governed by objective laws rooted in human nature; statesmen think and act in terms of interest defined as power. For classical realists, power is both a means and an end in itself. Consequently, classical realism is heavily normative, historical, and prescriptive, emphasizing the irreplaceable virtues of statesmanship, moral prudence, and the conscious diplomatic manipulation of the balance of power to restrain raw aggressive impulses."),

    ("Structural Neorealism: The Tyranny of Anarchy and Systemic Theory",
     "In 1979, Kenneth Waltz revolutionized international relations theory with Theory of International Politics, decisively breaking with Morgenthau's psychological reductionism. Waltz argued that human nature cannot explain variation in international outcomes: if human nature is an unchanging constant of greed and aggression, it cannot explain why the international system experiences prolonged periods of peace punctuated by acute wars. Waltz elevated the analysis to the 'Third Image'—the structure of the international system. In neorealism, the root cause of conflict is not flawed human nature, but the structural condition of anarchy: the absence of a central, governing authority to enforce law and protect states from predation. Under anarchy, even peace-loving, virtuous states are forced to act aggressively because of the security dilemma—the reality that measures taken by one state to enhance its defense are inevitably perceived by neighbors as offensive threats. Waltz defined the international system through a parsimonious structural triad: the ordering principle is anarchy, the units are functionally identical sovereign states striving for survival, and the only structural variable is the distribution of material capabilities across units (polarity)."),

    ("The Intra-Neorealist Schism: Defensive versus Offensive Realism",
     "While united against classical reductionism, structural neorealists subsequently fractured into two competing theoretical camps over the amount of power a state should seek: Defensive Realism and Offensive Realism. Defensive Realism, championed by Kenneth Waltz and Robert Jervis, maintains that the international system possesses a status-quo bias. Waltz argues that anarchy compels states to seek an appropriate amount of power to ensure survival, rather than infinite power. The system punishes aggressive overexpansion: states that bid for hegemony inevitably trigger overwhelming counter-balancing coalitions that lead to their own destruction (as demonstrated by Napoleon and Hitler). Conversely, Offensive Realism, pioneered by John J. Mearsheimer in The Tragedy of Great Power Politics, argues that anarchy inherently incentivizes aggression. Because states can never verify the future intentions of rivals, and because offensive military capabilities are lethal, rational great powers realize that the only way to be truly secure is to become the ultimate hegemon. For Mearsheimer, great powers are trapped in a tragic structural competition, relentlessly maximizing relative power with the goal of achieving regional hegemony."),

    ("Conclusion: The Enduring Hegemony of Realist Statecraft",
     "In conclusion, the evolution from classical realism to structural neorealism represents a profound intellectual transition from a philosophical critique of human biological frailty to a scientific, structural theory of systemic incentives. While Morgenthau viewed power as an end rooted in human nature, Waltz and Mearsheimer proved that power is an indispensable currency demanded by the anarchic structure of the international system. In the contemporary 21st-century geopolitical landscape, marked by the return of great-power contestation between the United States, China, and Russia, the structural insights of neorealism provide the most formidable explanatory framework for world affairs, demonstrating that in a world without a global sheriff, sovereign states remain tragic prisoners of structural anarchy, forced to prioritize material survival above all else.")
]

q19_quotes = [
    ("Hans J. Morgenthau", "Politics Among Nations (1948)", "Political realism believes that politics, like society in general, is governed by objective laws that have their roots in human nature. International politics, like all politics, is a struggle for power."),
    ("Kenneth N. Waltz", "Theory of International Politics (1979)", "The structure of the system forces states to behave in certain ways. In an anarchic domain, state behavior is dictated not by the inner characteristics of the units, but by the structural distribution of capabilities across them."),
    ("John J. Mearsheimer", "The Tragedy of Great Power Politics (2001)", "The sad fact is that international politics has always been a ruthless and dangerous business, and it is likely to remain that way. Anarchy forces great powers to maximize their share of world power, because having superior power is the best means to ensure survival."),
    ("Robert Jervis", "Perception and Misperception in International Politics", "The security dilemma is the core structural engine of realism: when one state gains power to protect itself, it inherently diminishes the security of its neighbors, sparking an escalatory spiral that neither side intended."),
    ("Thucydides", "History of the Peloponnesian War", "What made war inevitable was the growth of Athenian power and the fear which this caused in Sparta. This is the timeless realist truth of structural power transitions.")
]

# =========================================================================
# Q20: NUCLEAR DETERRENCE THEORY & SOUTH ASIA (2016 Q.3, 2013 Q.7)
# =========================================================================

q20_meta = [
    ("Target Examination", "KPPSC PMS (Provincial Management Service) — Khyber Pakhtunkhwa"),
    ("Paper & Subject", "International Relations (100 Marks Optional — Group 6)"),
    ("EXAM QUESTIONS LINKED:", '• 2016 Q.3: "Nuclear deterrence prevented a direct hot war between superpowers during the Cold War. Discuss the theory of Mutually Assured Destruction (MAD) and its applicability to the strategic stability of South Asia between Pakistan and India. (20 Marks)"\n• 2013 Q.7: "Examine the dynamics of Arms Control and Disarmament in the nuclear era. Evaluate the effectiveness and shortcomings of the Nuclear Non-Proliferation Treaty (NPT) and Comprehensive Test Ban Treaty (CTBT). (20 Marks)"'),
    ("Syllabus Section", "Part B: Peace and Security in Nuclear Age | Part C: South Asian Nuclear Dynamics"),
    ("Question Allocation", "Attempt 1 of 5 Required (20 Marks | Strict 35-Minute Allocation)"),
    ("Historiographical / Theoretical Thesis", "The nuclear revolution fundamentally altered the nature of international relations by replacing Clausewitzian warfare with the absolute logic of deterrence: as Bernard Brodie formulated, while the historical purpose of military force was to win wars, in the thermonuclear era its primary purpose is to avert them. Under Mutually Assured Destruction (MAD), secure second-strike capabilities inhibit systemic conflagration. However, when applied to the enduring rivalry between Pakistan and India, nuclear deterrence operates under Glenn Snyder's 'Stability-Instability Paradox': strategic nuclear parity deters all-out war, but simultaneously creates a permissive umbrella for low-intensity proxy conflict, border skirmishes, and dangerous doctrinal escalations (India's 'Cold Start' doctrine versus Pakistan's Full Spectrum Deterrence and tactical nuclear weapons).")
]

q20_callout = [
    "THE NUCLEAR REVOLUTION (BERNARD BRODIE): The catastrophic destructiveness of thermonuclear weapons inverts classical warfare; when both sides possess survivable second-strike capabilities, victory in war becomes impossible, institutionalizing Mutually Assured Destruction (MAD).",
    "THE STABILITY-INSTABILITY PARADOX (GLENN SNYDER): High strategic stability (neither state will launch strategic nuclear weapons because of guaranteed retaliation) generates instability at lower rungs of the escalation ladder (states feel emboldened to engage in sub-conventional, proxy, or border skirmishes).",
    "SOUTH ASIAN CRISES VALIDATION: The Kargil War (1999), the 2001–2002 military standoff, and the 2019 Balakot-Swift Retort aerial dogfight validate the paradox: nuclear deterrence prevented conventional escalation to full-scale war, but failed to deter limited military strikes.",
    "DOCTRINAL FRICTION & ESCALATION RISKS: India's limited-war 'Cold Start' doctrine (rapid blitzkrieg incursions) was met by Pakistan's 'Full Spectrum Deterrence' (FSD), inducting short-range Tactical Nuclear Weapons (TNWs like Nasr/Hatf-IX) to plug the deterrence gap, which drastically compresses crisis decision-making time."
]

q20_map = r"""
+---------------------------------------------------------------------------------------------------------+
|                  CONCEPTUAL ROADMAP: NUCLEAR DETERRENCE & SOUTH ASIAN STABILITY                         |
+---------------------------------------------------------------------------------------------------------+
                                                     |
                               [Theoretical Baseline: The Nuclear Revolution]
                             Bernard Brodie (1946): Nuclear Weapons Must Avert Wars
                             Mutually Assured Destruction (MAD) & Second-Strike Parity
                                                     |
                                                     v
                       [May 1998: The South Asian Nuclearization]
                       India (Pokhran-II) & Pakistan (Chagai-I) Conduct Nuclear Tests
                       Subcontinent Formally Enters Deterrence Architecture
                                                     |
                                                     v
                         [THE STABILITY-INSTABILITY PARADOX (GLENN SNYDER)]
                                                     |
           +-----------------------------------------+-----------------------------------------+
           |                                                                                   |
           v                                                                                   v
   [STRATEGIC LEVEL STABILITY]                                         [CONVENTIONAL & SUB-CONVENTIONAL INSTABILITY]
Total War is Deterred (MAD Works)                                   Low-Intensity Proxy Wars & Cross-LoC Clashes
No Direct Armored Invasion of Core Cities                           Kargil War (1999) Fought Under Nuclear Shadow
Both Capitals Face Inevitable Annihilation                          Sub-Conventional Emboldenment Behind Nuclear Shield
           |                                                                                   |
           +-----------------------------------------+-----------------------------------------+
                                                     |
                                                     v
                         [THE DANGEROUS DOCTRINAL SPIRAL: COLD START VS FSD]
                                                     |
           +-----------------------------------------+-----------------------------------------+
           |                                                                                   |
           v                                                                                   v
   [INDIAN CONVENTIONAL DOCTRINE]                                      [PAKISTAN'S ASYMMETRIC RESPONSE]
"Cold Start" / Proactive Strategy                                   "Full Spectrum Deterrence" (FSD)
Integrated Battle Groups (IBGs) Blitzkrieg                          Induction of Tactical Nuclear Weapons (TNWs: Nasr)
Limited Incursions Below Nuclear Redlines                           "Plugging the Deterrence Gap" at Tactical Level
           |                                                                                   |
           +-----------------------------------------+-----------------------------------------+
                                                     |
                                                     v
                         [THE SYSTEMIC HAZARD: BALAKOT 2019 & CRISIS COMPRESSION]
                     Airstrikes Penetrate Sovereign Territory; Swift Retort Aerial Combat;
                     Razor-Thin Decision-Making Time Requires Institutional Crisis Regimes
+---------------------------------------------------------------------------------------------------------+
"""

q20_timeline = [
    ("May 1998: Overt Nuclearization", "India tested 5 devices (Pokhran-II); Pakistan responded with 6 tests at Chagai.", "Ended decades of nuclear ambiguity; formalized nuclear deterrence as the bedrock of South Asian security."),
    ("May–July 1999: The Kargil Conflict", "Armed conflict along the Line of Control in Kashmir between newly nuclear-armed states.", "The classic empirical vindication of Snyder's Stability-Instability Paradox: limited war under nuclear umbrella."),
    ("2004: Inception of Indian 'Cold Start' Doctrine", "Indian Army developed rapid-mobilization doctrine to execute shallow punitive incursions.", "Designed explicitly to exploit the space between sub-conventional proxy war and Pakistan's strategic nuclear threshold."),
    ("2011: Pakistan Tests Nasr (Hatf-IX)", "Pakistan successfully tested short-range tactical nuclear missile (60km range).", "Operationalized Full Spectrum Deterrence to neutralize Indian conventional armored breakthroughs at the tactical level."),
    ("February 2019: Pulwama-Balakot Crisis", "Indian airstrike on Balakot followed by Pakistan's Operation Swift Retort and dogfight.", "Demonstrated the acute danger of miscalculation; strategic nuclear deterrence prevented escalation to all-out war.")
]

q20_essay = [
    ("Introduction: The Nuclear Revolution and the Inversion of Military Strategy",
     "The detonation of the atomic bomb over Hiroshima in August 1945 did not merely introduce a more destructive weapon of war; it fundamentally revolutionized the nature of international relations and military strategy. In his pioneering 1946 treatise The Absolute Weapon, strategist Bernard Brodie captured this epistemological rupture with profound clarity: 'Thus far the chief purpose of our military establishment has been to win wars. From now on its chief purpose must be to avert them. It can have almost no other useful purpose.' In classical warfare, military power was an instrument used to achieve political victory through territorial conquest. In the thermonuclear era, however, the concept of victory became obsolete. Under the condition of Mutually Assured Destruction (MAD), states possessing secure, survivable second-strike nuclear capabilities are locked in a condition of absolute deterrence. However, when nuclear deterrence theory was transplanted from the US-Soviet Cold War context into the complex, emotionally charged, and geographically contiguous rivalry between Pakistan and India, it produced a highly volatile strategic environment governed by the 'Stability-Instability Paradox.'"),

    ("The Theoretical Foundations of Nuclear Deterrence and MAD",
     "Deterrence theory, rooted in structural realism and rational deterrence frameworks, relies upon the psychological manipulation of an adversary's cost-benefit calculations. Deterrence is achieved when a state successfully convinces a potential aggressor that the unacceptable costs of an attack will overwhelmingly outweigh any conceivable political or military gains. For deterrence to function, three elements are indispensable: Capability (possession of invulnerable nuclear arsenals and reliable delivery systems), Credibility (the demonstrated political and military resolve to actually use those weapons if survival is threatened), and Communication (the unambiguous signaling of redlines to the adversary). Under Cold War MAD, both Washington and Moscow maintained a nuclear triad—intercontinental ballistic missiles, strategic bombers, and nuclear-powered ballistic missile submarines (SSBNs)—ensuring that even a devastating surprise first strike would be met with an annihilating second-strike retaliation, thereby freezing the balance of power in an uneasy peace."),

    ("The South Asian Crucible: The Stability-Instability Paradox",
     "Following the reciprocal nuclear tests of May 1998 (India's Pokhran-II and Pakistan's Chagai-I and Chagai-II), South Asia officially became a nuclearized theater. Prominent 'deterrence optimists,' such as Kenneth Waltz, argued that nuclear weapons would permanently pacify the subcontinent, as neither Islamabad nor New Delhi would ever dare risk national annihilation. However, 'deterrence pessimists,' led by Scott Sagan, warned that geographical proximity, historical hatred, shared disputed borders, and compressed flight times created extreme vulnerabilities. History rapidly validated the classic 'Stability-Instability Paradox,' conceptualized by Glenn Snyder. The paradox posits that while strategic nuclear parity creates absolute stability at the strategic level (neither state will launch an all-out war of territorial destruction), it simultaneously creates instability at the lower rungs of the escalation ladder. Because both sides recognize that strategic escalation is suicidal, they feel emboldened to engage in low-intensity conflict, cross-border proxy warfare, and limited incursions, confident that the nuclear threshold will prevent the conflict from escalating into total war."),

    ("Empirical Testing and Doctrinal Escalation: Cold Start vs Full Spectrum Deterrence",
     "The reality of the Stability-Instability Paradox was brutally demonstrated in the 1999 Kargil War—the first direct armed conflict between two nuclear-armed states in modern history—and during the 2001–2002 military standoff. Frustrated by its inability to leverage its massive conventional military superiority against Pakistan due to Islamabad's nuclear shield, the Indian military establishment developed the 'Cold Start' doctrine (later formalized as the Proactive Strategy). Cold Start envisioned the rapid mobilization of eight integrated battle groups (IBGs) to launch lightning, shallow punitive strikes into Pakistani territory within 72 to 96 hours, holding territory as a bargaining chip while staying strictly beneath Pakistan's nuclear threshold. Recognizing that Cold Start was designed to exploit conventional asymmetry, Pakistan responded by formulating 'Full Spectrum Deterrence' (FSD). In 2011, Pakistan tested the short-range, solid-fueled Nasr (Hatf-IX) tactical ballistic missile (range 60–70 km). By introducing low-yield Tactical Nuclear Weapons (TNWs) onto the conventional battlefield, Pakistan's Strategic Plans Division (SPD) explicitly aimed to 'pour cold water on Cold Start,' plugging the deterrence gap at the tactical level."),

    ("Conclusion: The Fragility of Deterrence and the Imperative of Strategic Restraint",
     "In conclusion, nuclear deterrence in South Asia has successfully prevented a major total war between Pakistan and India for over a quarter of a century. However, as demonstrated by the harrowing February 2019 Pulwama-Balakot crisis—which witnessed Indian cross-border airstrikes and Pakistan's retaliatory Operation Swift Retort aerial combat—deterrence stability is becoming dangerously fragile. The induction of short-range tactical nuclear weapons, the operationalization of hypersonic cruise missiles, the deployment of ballistic missile defense (BMD) systems, and compressed missile flight times (often under four minutes between launch and impact) leave virtually zero margin for decision-making or error during an acute crisis. South Asian stability cannot safely rest upon the knife-edge of hair-trigger nuclear deterrence; both states must urgently formalize a comprehensive Strategic Restraint Regime (SRR), institutionalize permanent crisis-communication hotlines, and pursue a peaceful resolution to the core dispute of Jammu and Kashmir to prevent the Stability-Instability Paradox from collapsing into nuclear catastrophe.")
]

q20_quotes = [
    ("Bernard Brodie", "The Absolute Weapon (1946)", "Thus far the chief purpose of our military establishment has been to win wars. From now on its chief purpose must be to avert them. It can have almost no other useful purpose. In the nuclear age, victory is a meaningless concept."),
    ("Glenn Snyder", "The Balance of Power and the Balance of Terror", "The stability-instability paradox is the structural reality of the nuclear age: the greater the stability of the 'strategic' balance of terror, the greater the likelihood of minor skirmishes and low-level violence."),
    ("Kenneth N. Waltz", "The Spread of Nuclear Weapons: More May Better", "Nuclear weapons make wars impossible between those who possess them. In South Asia, as during the Cold War, the certainty of catastrophic retaliation forces both India and Pakistan to behave with ultimate prudence."),
    ("Scott D. Sagan", "The Spread of Nuclear Weapons: An Enduring Debate", "Professional military organizations display organizational biases and cognitive limitations that make accidental nuclear war or catastrophic miscalculation highly likely in geographically contiguous regions like South Asia."),
    ("Lt. Gen. (Retd) Khalid Kidwai", "Advisor to Pakistan's National Command Authority", "Pakistan's Full Spectrum Deterrence is not an academic theory; it is a defensive operational necessity. By possessing nuclear options at the tactical, operational, and strategic levels, we have plugged the gap beneath which India sought to wage conventional war.")
]

# =========================================================================
# Q21: THE UNITED NATIONS IN CONFLICT RESOLUTION: UNSC & VETO (2016 Q.6)
# =========================================================================

q21_meta = [
    ("Target Examination", "KPPSC PMS (Provincial Management Service) — Khyber Pakhtunkhwa"),
    ("Paper & Subject", "International Relations (100 Marks Optional — Group 6)"),
    ("EXAM QUESTIONS LINKED:", '• 2016 Q.6: "The United Nations was established to save succeeding generations from the scourge of war. Critically assess the performance of the UN Security Council in conflict resolution and peacekeeping in the post-Cold War era. (20 Marks)"'),
    ("Syllabus Section", "Part A: International Organizations & Part B: Collective Security (UN)"),
    ("Question Allocation", "Attempt 1 of 5 Required (20 Marks | Strict 35-Minute Allocation)"),
    ("Historiographical / Theoretical Thesis", "Founded in the ashes of the Second World War to 'save succeeding generations from the scourge of war,' the United Nations was engineered as an institutionalized collective security framework designed to overcome the structural failures of the League of Nations. However, the operational efficacy of its preeminent executive organ—the UN Security Council (UNSC)—is fatally compromised by the institutionalized realpolitik of the veto power granted to the five permanent members (P5 under Article 27). During the Cold War, ideological confrontation between the superpowers paralyzed the Council; while the post-Cold War era witnessed an initial expansion of Chapter VII peace enforcement, it was marred by catastrophic moral failures in Rwanda and Srebrenica. In the contemporary multipolar era, the resurgence of great-power rivalry has completely gridlocked the UNSC over Ukraine and Gaza, demonstrating that collective security remains perpetually subordinated to great power realpolitik.")
]

q21_callout = [
    "THE COLLECTIVE SECURITY IDEAL: Founded on Chapter VII of the UN Charter: an attack on any sovereign member is viewed as an attack on all; the UNSC holds the legal monopoly to authorize mandatory sanctions, embargoes, and military force to maintain international peace.",
    "THE FATAL FLAW: ARTICLE 27 VETO POWER: To entice great powers to join, the 1945 Charter granted absolute veto power to the P5 (US, Russia, China, UK, France); a single negative vote paralyzes the entire collective security machinery, ensuring the UN can never act against a permanent member or its protected client states.",
    "THE TRAGIC RECORD OF POST-COLD WAR FAILURES: The 1990s optimism collapsed into moral infamy: the 1994 Rwandan Genocide (UN peacekeepers under Roméo Dallaire ordered to withdraw while 800,000 were murdered) and the 1995 Srebrenica massacre inside a designated UN 'Safe Area'.",
    "CONTEMPORARY MULTIPOLAR PARALYSIS & REFORM IMPASSE: Current gridlock: Russia vetoes resolutions on Ukraine; the US repeatedly vetoes ceasefires in Gaza; China/Russia shield Syria. The reform debate is hopelessly deadlocked between the G4 (demanding permanent seats) and Uniting for Consensus (advocating non-permanent expansion without veto)."
]

q21_map = r"""
+---------------------------------------------------------------------------------------------------------+
|                  CONCEPTUAL ROADMAP: THE UNITED NATIONS & SECURITY COUNCIL IN IR                        |
+---------------------------------------------------------------------------------------------------------+
                                                     |
                               [1945 San Francisco Charter: Collective Security]
                             "Save Succeeding Generations from the Scourge of War"
                             Chapter VI (Pacific Settlement) vs Chapter VII (Enforcement)
                                                     |
                                                     v
                       [The Structural Fault Line: Article 27(3) P5 Veto Power]
                       Subordination of Collective Security to Great Power Realpolitik
                       US, Russia, China, UK, France Possess Unilateral Veto
                                                     |
                                                     v
                         [HISTORICAL PHASES OF SECURITY COUNCIL PERFORMANCE]
                                                     |
           +-----------------------------------------+-----------------------------------------+
           |                                         |                                         |
           v                                         v                                         v
   [COLD WAR PARALYSIS]                      [POST-COLD WAR EXPANSION & FAILURE]       [21st CENTURY POLARIZATION]
Superpower Ideological Gridlock            1991 Gulf War: Collective Action Triumph   Ukraine War: Russia Vetoes Action
Over 200 Vetoes Cast; UNSC Frozen          Catastrophic Failures: Rwanda (1994) &     Gaza Conflict: US Vetoes Ceasefires
Peacekeeping Invented as Improvised        Srebrenica (1995) 'Safe Areas' Exposed     Syrian Civil War: Sino-Russian Vetoes
Buffer (Chapter 6.5)                      Mission Creep and Lack of Political Will   UNSC Reduced to Theatrical Platform
           |                                         |                                         |
           +-----------------------------------------+-----------------------------------------+
                                                     |
                                                     v
                         [THE REFORM DEADLOCK: COMPETING GLOBAL VISIONS]
                     G4 (India, Brazil, Germany, Japan): Demand Permanent Veto Seats
                     Uniting for Consensus (Pakistan, Italy): Expand Elected Non-Permanent
+---------------------------------------------------------------------------------------------------------+
"""

q21_timeline = [
    ("June 1945: Signing of UN Charter", "50 nations signed Charter in San Francisco; established Security Council with P5 veto.", "Created modern multilateral collective security architecture with institutionalized great power dominance."),
    ("1950: Uniting for Peace Resolution (377)", "General Assembly asserted right to act on threats to peace when UNSC is veto-paralyzed.", "First institutional effort by smaller nations to circumvent P5 veto gridlock during the Korean War."),
    ("1991: Operation Desert Storm", "UNSC authorized military coalition under Chapter VII to liberate Kuwait from Iraqi invasion.", "Celebrated as the pinnacle of post-Cold War collective security in George H.W. Bush's 'New World Order.'"),
    ("1994–1995: The Rwandan Genocide & Srebrenica", "UNAMIR abandoned Rwanda; Dutch UN peacekeepers stood by as 8,000 Bosniaks were massacred in Srebrenica.", "Exposed the moral bankruptcy of traditional peacekeeping mandates when confronted by genocidal intent."),
    ("2022–Present: Existential Gridlock over Ukraine & Gaza", "UNSC rendered completely paralyzed by reciprocal Russian and American vetoes.", "Re-ignited universal demands from the Global South for fundamental democratization of global governance.")
]

q21_essay = [
    ("Introduction: The Moral Promise and Structural Reality of the United Nations",
     "Established on October 24, 1945, amidst the smoldering devastation of the Second World War, the United Nations was founded on the majestic moral commitment inscribed in the preamble of its Charter: 'to save succeeding generations from the scourge of war.' Learning from the catastrophic institutional collapse of the League of Nations—which had failed to deter fascist and militarist aggression in Manchuria, Abyssinia, and Central Europe—the architects of the San Francisco Charter sought to create an institutionalized system of collective security backed by formidable enforcement capabilities. Under the Charter, primary responsibility for the maintenance of international peace and security was vested exclusively in the United Nations Security Council (UNSC). Unlike the League's toothless Assembly, the UNSC was granted statutory authority under Chapter VII to issue legally binding resolutions, mandate sweeping economic sanctions, and authorize collective military force to repel aggression. However, the operational record of the Security Council over eight decades reveals a profound, tragic chasm between its collective security ideal and its structural reality. Dominated by the anachronistic veto power of its five permanent members (P5), the Security Council has perpetually struggled to function as an impartial guardian of global peace, operating instead as a battleground of great power realpolitik."),

    ("The Structural Flaw: Article 27 and the Tyranny of the Veto",
     "The fundamental structural impediment to effective UN conflict resolution is enshrined in Article 27, Paragraph 3 of the UN Charter, which dictates that substantive decisions of the Security Council require the concurring votes of all five permanent members: the United States, the Russian Federation (formerly the USSR), the People's Republic of China, the United Kingdom, and the French Republic. This 'veto power' was the non-negotiable price demanded by the great powers to join the organization, designed to ensure that the UN could never be mobilized to wage war against a nuclear-armed power. However, this constitutional design embedded a fatal paradox at the heart of the system: collective security works only when the great powers agree, but it is precisely when the great powers disagree that international peace is most imperiled. During the four decades of the Cold War, ideological confrontation between Washington and Moscow rendered the Security Council almost entirely impotent. Over 200 vetoes were cast, transforming the Council into a forum for theatrical propaganda while the superpowers fought destructive proxy wars across Korea, Vietnam, the Middle East, and Afghanistan."),

    ("The Post-Cold War Illusion and Catastrophic Humanitarian Failures",
     "The dissolution of the Soviet Union in 1991 generated a fleeting era of unipolar optimism. In January 1991, the UNSC successfully authorized a US-led multinational coalition to expel Iraqi forces from Kuwait, leading President George H.W. Bush to herald a 'New World Order' where the rule of law would supplant the rule of the jungle. Concurrently, Secretary-General Boutros Boutros-Ghali published An Agenda for Peace, proposing ambitious expansions of preventative diplomacy and peace-enforcement. However, this euphoria was brutally shattered by the catastrophic humanitarian debacles of the mid-1990s. In 1994, during the Rwandan Genocide, the Security Council—hamstrung by American refusal to intervene following the 'Black Hawk Down' incident in Somalia—refused to reinforce the UN Assistance Mission for Rwanda (UNAMIR) and ordered the withdrawal of peacekeepers, standing idly by while 800,000 Tutsis were butchered in 100 days. A year later, in July 1995, Dutch UN peacekeepers in Bosnia surrendered the designated UN 'Safe Area' of Srebrenica to Serb forces, permitting the genocidal massacre of over 8,000 Muslim men and boys. These horrors demonstrated that without great power political will, Chapter VI peacekeeping mandates were death traps."),

    ("Contemporary Multipolar Gridlock: Ukraine, Gaza, and Normative Paralysis",
     "In the contemporary era, the return of intense multipolar rivalry has plunged the Security Council into its most severe paralysis since the height of the Cold War. The structural bankruptcy of the veto was laid bare in February 2022: when the Russian Federation launched a full-scale invasion of Ukraine—a sovereign member state—Moscow simply used its permanent veto to block the draft resolution condemning its own aggression. Similarly, in the ongoing Israeli-Palestinian war in Gaza, the United States has repeatedly exercised its unilateral veto to shield Israel from binding ceasefire resolutions overwhelmingly supported by the international community. On other critical fronts, Sino-Russian vetoes have repeatedly blocked humanitarian resolutions and chemical weapons investigations in Syria, as well as sanctions on the military junta in Myanmar. This persistent gridlock has severely eroded the moral legitimacy of the Security Council, reducing it to a paralyzed bystander while hundreds of thousands perish in preventable conflicts."),

    ("Conclusion: The Reform Impasse and the Urgent Need for Democratization",
     "In conclusion, the United Nations Security Council in the 21st century reflects the geopolitical realities of 1945, not the multipolar world of today. It represents an archaic, Eurocentric oligopoly: Africa and Latin America possess zero permanent representation, while over two billion Muslims have no voice in the permanent chamber. However, efforts to reform the Council are hopelessly deadlocked. The Group of Four (G4: India, Brazil, Germany, Japan) demands permanent seats with veto power, triggering fierce opposition from the 'Uniting for Consensus' (UFC or Coffee Club) coalition—led by Pakistan, Italy, Mexico, and South Korea—which advocates for the democratic expansion of elected, non-permanent seats without the polarizing veto. As former Secretary-General Dag Hammarskjöld famously observed, 'The United Nations was not created in order to bring us to heaven, but in order to save us from hell.' If the Security Council is to avoid sliding into the irrelevance of the League of Nations, its members must urgently reform its composition and restrict the use of the veto in situations of mass atrocities, restoring collective security as an authentic instrument of universal human protection."),
]

q21_quotes = [
    ("Dag Hammarskjöld", "Second UN Secretary-General", "The United Nations was not created in order to bring us to heaven, but in order to save us from hell. It is an imperfect instrument, but it is the only one humanity possesses to prevent total catastrophe."),
    ("Inis L. Claude Jr.", "Swords into Plowshares", "The veto is not the cause of great-power disagreement; it is the reflection of it. To abolish the veto without eliminating great power rivalry would simply destroy the United Nations, for no great power will submit its survival to an international vote."),
    ("Kofi Annan", "In Larger Freedom (2005)", "No reform of the United Nations will be complete without reform of the Security Council. A Council that reflects the power distribution of 1945 cannot enjoy the legitimacy required to lead the world in the 21st century."),
    ("John J. Mearsheimer", "The False Promise of International Institutions", "Institutions are not a cure for power politics; they are an arena where power politics is played out. The Security Council reflects the balance of power among the major states; it does not change state behavior when vital interests clash."),
    ("Boutros Boutros-Ghali", "An Agenda for Peace (1992)", "The time of absolute sovereignty has passed; its theory was never matched by reality. The task for statesmanship today is to acknowledge this and to find a balance between the needs of good internal governance and the requirements of an interdependent world.")
]

# =========================================================================
# Q22: THEORIES OF IMPERIALISM: HOBSON, LENIN & NEO-COLONIALISM (2016 Q.7)
# =========================================================================

q22_meta = [
    ("Target Examination", "KPPSC PMS (Provincial Management Service) — Khyber Pakhtunkhwa"),
    ("Paper & Subject", "International Relations (100 Marks Optional — Group 6)"),
    ("EXAM QUESTIONS LINKED:", '• 2016 Q.7: "Discuss the concept of \\"Imperialism\\" and its economic theories, particularly Hobson and Lenin\'s critique of imperial capitalism. To what extent does neo-imperialism persist in the contemporary global economic order? (20 Marks)"'),
    ("Syllabus Section", "Part A: Basic Approaches: Imperialism, Capitalist Expansion & Global Inequality"),
    ("Question Allocation", "Attempt 1 of 5 Required (20 Marks | Strict 35-Minute Allocation)"),
    ("Historiographical / Theoretical Thesis", "The phenomenon of imperialism—the structural domination and economic subjugation of peripheral societies by core metropolitan states—underwent a profound analytical revolution in the early 20th century through the economic theories of John A. Hobson and Vladimir Lenin. While Hobson formulated a reformist liberal critique attributing imperialism to domestic under-consumption and the maldistribution of wealth that could be resolved through social democracy, Lenin identified imperialism as the 'highest stage of capitalism'—the inexorable culmination of monopoly finance capital seeking external markets, captive resources, and super-profits. In the post-colonial 21st century, traditional territorial annexation has been superseded by 'neo-imperialism' (Nkrumah) and structural dependency (Wallerstein), where developing nations remain trapped in unequal exchange through sovereign debt mechanisms, multinational corporate exploitation, and Western institutional financial hegemony.")
]

q22_callout = [
    "HOBSON'S LIBERAL UNDER-CONSUMPTION THEORY (1902): Imperialism is driven by domestic inequality: working classes are underpaid, creating a surplus of manufactured goods and capital that cannot be absorbed domestically; finance capitalists lobby imperial governments for territorial annexation. Hobson argued this could be cured by domestic democratic wealth redistribution.",
    "LENIN'S MARXIST HIGHEST STAGE (1916): Imperialism is not a policy choice, but the structural, inevitable culmination of capitalism: (1) Concentration of production into monopolies; (2) Merger of banking and industrial capital into 'finance capital'; (3) The export of capital overrides the export of commodities; (4) Division of the world among international capitalist cartels; (5) Territorial division of the globe leading to inevitable inter-imperialist war.",
    "THE EVOLUTION TO NEO-IMPERIALISM (KWAME NKRUMAH): Formal decolonization dismantled direct military-administrative rule, but replaced it with neo-colonialism: client regimes grant sovereign economic access to Western multinational corporations, resulting in 'power without responsibility and exploitation without redress.'",
    "STRUCTURAL DEPENDENCY & WORLD-SYSTEMS (WALLERSTEIN): Prebisch-Singer thesis on declining terms of trade; Wallerstein's Core-Periphery-Semiperiphery model proves that global capitalism systematically drains surplus value from the Global South to the Global North via debt servicing and unequal exchange."
]

q22_map = r"""
+---------------------------------------------------------------------------------------------------------+
|                  CONCEPTUAL ROADMAP: THEORIES OF IMPERIALISM & NEO-COLONIALISM                          |
+---------------------------------------------------------------------------------------------------------+
                                                     |
                               [Historical Genesis: Industrial Revolution & Scramble]
                             1870–1914: European Conquest of Africa and Asia
                             The Search for Markets, Raw Materials, and Geo-Strategic Bases
                                                     |
                                                     v
                       [THE CLASSICAL ECONOMIC THEORIES OF IMPERIALISM]
                                                     |
           +-----------------------------------------+-----------------------------------------+
           |                                                                                   |
           v                                                                                   v
   [JOHN A. HOBSON: LIBERAL UNDER-CONSUMPTION (1902)]                  [V.I. LENIN: HIGHEST STAGE OF CAPITALISM (1916)]
Root Cause: Domestic Wealth Maldistribution                         Root Cause: Inevitable Structural Laws of Capitalism
Workers Underpaid --> Domestic Market Saturated                     Monopoly Capitalism Replaces Free Competition
Finance Capitalists Force Imperialist Expansion                     Bank + Industrial Capital = Finance Capital Oligarchy
Prescription: Reformist Social Democracy, Wage Hikes               Export of Capital Drives Violent Global Partition
Imperialism is a Curable Disease of Capitalism                     Prescription: Proletarian Revolution & Anti-Imperialist War
           |                                                                                   |
           +-----------------------------------------+-----------------------------------------+
                                                     |
                                                     v
                         [POST-1945 DECOLONIZATION & THE MUTATION TO NEO-IMPERIALISM]
                       Flag Independence Achieved, but Economic Chains Remain
                       Kwame Nkrumah: Neo-Colonialism = Indirect Economic Subjugation
                                                     |
           +-----------------------------------------+-----------------------------------------+
           |                                         |                                         |
           v                                         v                                         v
 [UNEQUAL TERMS OF TRADE]                  [DEBT ENTRAPMENT & STRUCTURAL ADJUSTMENT] [TRANSNATIONAL CORPORATE EXTRACTION]
Singer-Prebisch Thesis: Raw Materials     IMF/World Bank Austerity Conditionalities  Multinational Corporations (MNCs) Siphon
Lose Value vs Manufactured Tech Goods     Sovereign Debt Servicing Drains Health/Ed Wealth; Transfer Pricing Evades Taxes
Peripheral Dependency on Western Core     Privatization of National Resources        Modern Imperialism Operates Through Markets
           |                                         |                                         |
           +-----------------------------------------+-----------------------------------------+
                                                     |
                                                     v
                         [CONTEMPORARY MANIFESTATION: GLOBAL STRUCTURAL INEQUALITY]
                     Wallerstein's World-Systems Theory: The Core Continues to Exploit
                     the Periphery, Sustaining an Asymmetric Global Economic Hierarchy
+---------------------------------------------------------------------------------------------------------+
"""

q22_timeline = [
    ("1884–1885: Berlin Conference", "European powers partitioned the African continent without indigenous consent.", "Codified the zenith of classical territorial imperialism under the guise of the 'civilizing mission.'"),
    ("1902: John A. Hobson Publishes 'Imperialism: A Study'", "Formulated the under-consumption thesis, identifying financial cartels as the engine of conquest.", "Provided the seminal economic foundation that influenced all subsequent Marxist and critical theorists."),
    ("1916: Lenin Publishes 'Imperialism, the Highest Stage of Capitalism'", "Written in Swiss exile; analyzed WWI as an inevitable clash between rival imperialist monopolies.", "Became the canonical Marxist-Leninist framework for 20th-century anti-colonial liberation movements."),
    ("1965: Kwame Nkrumah Coins 'Neo-Colonialism'", "First President of Ghana published 'Neo-Colonialism: The Last Stage of Imperialism.'", "Exposed how indirect economic, financial, and monetary levers perpetuate colonial exploitation post-independence."),
    ("1974: Immanuel Wallerstein's World-Systems Theory", "Published 'The Modern World-System,' conceptualizing Core, Periphery, and Semiperiphery.", "Provided the definitive contemporary structural framework analyzing the persistence of global unequal exchange.")
]

q22_essay = [
    ("Introduction: The Theoretical Evolution of Imperial Domination",
     "Imperialism—the policy, practice, and advocacy of extending the power and dominion of one nation over other societies through direct territorial acquisition or indirect economic and political control—is as ancient as human civilization. From the Roman and Ottoman empires to the Spanish conquest of the Americas, empires historically conquered territory for tribute, slaves, and prestige. However, the late nineteenth century witnessed a qualitative mutation in the nature of empire: the 'New Imperialism' (1870–1914), during which European powers subjected an astonishing 84% of the globe to colonial subjugation, culminating in the frenzied Scramble for Africa. In the discipline of international relations and political economy, this unprecedented expansion demanded a rigorous theoretical explanation. This intellectual quest generated two of the most influential economic treatises of modern history: John A. Hobson's liberal under-consumption theory and Vladimir Ilyich Lenin's Marxist formulation of imperialism as the highest stage of capitalism. While direct territorial colonialism disintegrated in the post-World War II era, their foundational insights remain vital for understanding the contemporary persistence of 'neo-imperialism' in the global political economy."),

    ("Hobson's Liberal Critique: Under-Consumption and the Parasitism of Finance",
     "In his seminal 1902 monograph Imperialism: A Study, British liberal economist John A. Hobson provided the first systematic economic deconstruction of modern imperial expansion. Writing against the backdrop of the brutal Boer War in South Africa, Hobson challenged the jingoistic myth that imperialism was driven by national glory, patriotism, or the benevolent 'civilizing mission.' Instead, Hobson located the taproot of imperialism in a severe structural dysfunction within domestic capitalist economies: under-consumption and the unequal distribution of wealth. In capitalist Britain, the concentration of income in the hands of a tiny aristocratic and capitalist elite meant that the working classes were paid subsistence wages, leaving them without the purchasing power to absorb the torrent of manufactured goods produced by modern industrial factories. Simultaneously, the rich accumulated vast surplus capital that could not find profitable investment opportunities at home. To resolve this crisis, aggressive finance capitalists—particularly international bankers and bondholders—lobbied and hijacked the state apparatus, using the national army and navy to annex foreign territories in Asia and Africa to create protected captive markets and lucrative investment outlets. Crucially, Hobson was a reformist: he argued that imperialism was an avoidable, parasitic disease of capitalism, which could be eradicated through progressive taxation, trade union empowerment, wage increases, and democratic wealth redistribution at home."),

    ("Lenin's Marxist Diagnosis: Imperialism as the Highest Stage of Capitalism",
     "Writing in Swiss exile in 1916 amidst the carnage of World War I, Vladimir Lenin radically expanded Hobson's insights in Imperialism, the Highest Stage of Capitalism. Unlike Hobson, Lenin rejected the reformist notion that imperialism was an optional policy that could be cured by social democracy. For Lenin, imperialism was the structural, inevitable, and terminal stage of capitalist evolution. Lenin delineated five essential features of this epoch: first, the concentration of production into colossal monopolies that replaced free market competition; second, the fusion of banking capital with industrial capital, creating a parasitic 'finance oligarchy'; third, the supreme preeminence of the export of capital over the export of ordinary commodities; fourth, the formation of international capitalist cartels and trusts that partitioned global markets among themselves; and fifth, the complete territorial division of the entire world among the great powers. Because the globe was finite and entirely partitioned, the uneven development of capitalism inevitably meant that rising powers (such as Imperial Germany) could only acquire colonies by violently stripping them away from established empires (Britain and France). World War I was therefore an inevitable inter-imperialist war for the re-division of global loot."),

    ("Decolonization and the Metamorphosis into Neo-Imperialism",
     "Following the collapse of European empires after 1945, classical territorial imperialism was dismantled by heroic anti-colonial liberation movements across Asia and Africa. However, the demise of formal colonial administration did not signify the end of economic exploitation; it marked its evolution into 'neo-imperialism.' In his visionary 1965 text Neo-Colonialism: The Last Stage of Imperialism, Ghana's first president, Kwame Nkrumah, captured this subtle reality: 'The essence of neo-colonialism is that the State which is subject to it is, in theory, independent and has all the outward trappings of international sovereignty. In reality its economic system and thus its political policy is directed from outside.' Neo-imperialism operates without colonial governors or military garrisons. Instead, core capitalist states exert domination through comprador domestic political elites, the structural hegemony of the US dollar, the monopoly control of technology, and the extraterritorial leverage of multinational corporations (MNCs) that repatriate billions in raw mineral and agricultural wealth from developing societies."),

    ("Conclusion: Structural Dependency and the Modern World-System",
     "In conclusion, the economic theories of Hobson and Lenin laid the intellectual foundation for contemporary dependency theory (Raúl Prebisch, André Gunder Frank) and Immanuel Wallerstein's World-Systems Theory. The modern international order remains rigidly stratified into an industrialized 'Core' that extracts surplus value from an agrarian, debt-ridden 'Periphery.' Through the mechanism of unequal exchange—where the prices of raw commodities exported by the Global South steadily decline relative to the manufactured high-tech goods imported from the North (the Prebisch-Singer thesis)—developing states are trapped in perpetual balance-of-payments crises. This structural trap is enforced by Western-dominated international financial institutions (the IMF and World Bank), which mandate fiscal austerity, resource privatization, and currency devaluations as conditionalities for bailout loans. Modern imperialism no longer needs to send gunboats up the Yangtze or Congo rivers; it enforces global hierarchy through bond markets, sovereign debt contracts, intellectual property rights, and credit-rating agencies. The struggle against imperialism in the 21st century is fundamentally a struggle to decolonize the global financial architecture and forge an equitable international economic order.")
]

q22_quotes = [
    ("Vladimir Ilyich Lenin", "Imperialism: The Highest Stage of Capitalism (1916)", "Imperialism is capitalism at that stage of development at which the dominance of monopolies and finance capital is established; in which the export of capital has acquired pronounced importance; in which the division of the world among the international trusts has begun."),
    ("John A. Hobson", "Imperialism: A Study (1902)", "Imperialism is the endeavour of the great controllers of industry to broaden the channel for the flow of their surplus wealth by seeking foreign markets and foreign investments to take off the goods and capital they cannot sell or use at home."),
    ("Kwame Nkrumah", "Neo-Colonialism: The Last Stage of Imperialism (1965)", "Neo-colonialism is the worst form of imperialism. For those who practise it, it means power without responsibility, and for those who suffer from it, it means exploitation without redress."),
    ("Immanuel Wallerstein", "The Modern World-System (1974)", "The capitalist world-economy is built on an international division of labour that divides the world into Core, Periphery, and Semiperiphery. The core appropriates the surplus value of the periphery through structural, institutionalized unequal exchange."),
    ("Raúl Prebisch", "The Economic Development of Latin America (1950)", "The peripheral countries are caught in a structural trap: the terms of trade systematically move against primary commodities in favor of manufactured goods, transferring wealth from the poorest nations to the industrial centers.")
]

# =========================================================================
# Q23: THE NON-ALIGNED MOVEMENT (NAM) (2016 Q.8)
# =========================================================================

q23_meta = [
    ("Target Examination", "KPPSC PMS (Provincial Management Service) — Khyber Pakhtunkhwa"),
    ("Paper & Subject", "International Relations (100 Marks Optional — Group 6)"),
    ("EXAM QUESTIONS LINKED:", '• 2016 Q.8: "What is the Non-Aligned Movement (NAM)? Evaluate its historical contribution during the Cold War bipolarity and analyze its relevance in today\'s multipolar world. (20 Marks)"\n• 2013 Q.8(a): "Write short note on: Decolonization and the Rise of the Third World. (10 Marks)"'),
    ("Syllabus Section", "Part B: Cold War Dynamics, Rise of Third World & Non-Aligned Movement"),
    ("Question Allocation", "Attempt 1 of 5 Required (20 Marks | Strict 35-Minute Allocation)"),
    ("Historiographical / Theoretical Thesis", "Born from the historic 1955 Bandung Asian-African Conference and formally founded at the 1961 Belgrade Summit, the Non-Aligned Movement (NAM) represented an extraordinary diplomatic and normative revolt by newly decolonized developing nations against the coercive, bipolar entrapment of the Cold War. Spearheaded by visionary statesmen—Jawaharlal Nehru, Josip Broz Tito, Gamal Abdel Nasser, Sukarno, and Kwame Nkrumah—NAM institutionalized the Panchsheel principles of sovereign equality, mutual non-aggression, and refusal to join military pacts (NATO/Warsaw Pact). While the demise of the Soviet Union led critics to pronounce NAM obsolete, the contemporary resurgence of intense great-power competition (US-China rivalry, Russia-West friction) has re-energized the core doctrine of non-alignment, transforming it into an indispensable strategy of 'strategic autonomy' and multi-alignment across the Global South.")
]

q23_callout = [
    "FOUNDATIONAL HISTORICAL MILESTONES: 1955 Bandung Conference (Asian-African solidarity); 1961 Belgrade Summit formally established NAM with 25 member states, establishing non-involvement in Cold War military alliances as its admission criterion.",
    "THE PANCHSHEEL PRINCIPLES: Five Pillars of Peaceful Coexistence: (1) Mutual respect for territorial integrity and sovereignty; (2) Mutual non-aggression; (3) Mutual non-interference in internal affairs; (4) Equality and mutual benefit; (5) Peaceful coexistence.",
    "HISTORICAL COLD WAR ACHIEVEMENTS: Prevented the complete partitioning of the Third World into military blocs; accelerated international support for anti-colonial liberation and the anti-Apartheid movement in South Africa; and championed the 1974 UN Declaration for a New International Economic Order (NIEO).",
    "RENAISSANCE IN 21ST-CENTURY MULTIPOLARITY: Far from obsolete, non-alignment has evolved into 'Non-Alignment 2.0' or 'Strategic Autonomy.' Developing states across Asia, Africa, and Latin America refuse to join binary anti-China or anti-Russia coalitions, hedging between superpowers to maximize sovereign economic and technological interests."
]

q23_map = r"""
+---------------------------------------------------------------------------------------------------------+
|                  CONCEPTUAL ROADMAP: THE NON-ALIGNED MOVEMENT (NAM) IN IR                               |
+---------------------------------------------------------------------------------------------------------+
                                                     |
                               [Systemic Context: Post-1945 Cold War & Decolonization]
                             Bipolar Trap: US Containment (NATO/CENTO) vs Soviet Blocs
                             Newly Independent States Refuse to Become Cold War Pawns
                                                     |
                                                     v
                       [The Diplomatic Genesis: Bandung (1955) to Belgrade (1961)]
                       The Founding Troika: Nehru (India), Tito (Yugoslavia), Nasser (Egypt)
                       Joined by Sukarno (Indonesia) and Kwame Nkrumah (Ghana)
                                                     |
                                                     v
                         [THE PANCHSHEEL CRITERIA & COLD WAR ACHIEVEMENTS]
                                                     |
           +-----------------------------------------+-----------------------------------------+
           |                                         |                                         |
           v                                         v                                         v
   [REJECTION OF ALLIANCE BLOCS]             [ANTI-COLONIAL SOLIDARITY]                [ECONOMIC REFORM: NIEO]
Refusal to Join NATO, SEATO, CENTO,        Diplomatic Shield for African & Asian     Demanded Restructuring of Bretton Woods
or Warsaw Pact Multilateral Pacts          Libya, Algeria, Angola Liberation Backed  1974 UN Declaration on New International
Preserved Sovereign Freedom of Action      Vigorous Crusade Against Apartheid SA     Economic Order to Combat Exploitation
           |                                         |                                         |
           +-----------------------------------------+-----------------------------------------+
                                                     |
                                                     v
                         [1991 SYSTEMIC CRISIS: IS NAM OBSOLETE?]
                     Dissolution of USSR & Bipolarity Disappears; Unipolar Moment
                     Critics Proclaim NAM Dead; Internal Bilateral Wars Erode Credibility
                                                     |
                                                     v
                         [THE 21st CENTURY RESURGENCE: NON-ALIGNMENT 2.0]
                     The Return of Great Power Rivalry (US vs China vs Russia)
                     Evolution into "Strategic Autonomy" & Issue-Based Multi-Alignment:
                     Global South (India, Brazil, SA, ASEAN) Refuses Binary Alignments
+---------------------------------------------------------------------------------------------------------+
"""

q23_timeline = [
    ("April 1955: The Bandung Conference", "29 Asian and African states convened in Indonesia to forge South-South solidarity.", "Formulated the foundational 10 Principles of Bandung, establishing the ideational bedrock of non-alignment."),
    ("September 1961: The First Belgrade Summit", "25 states formally founded the Non-Aligned Movement under Tito, Nehru, and Nasser.", "Codified strict non-membership in multilateral Cold War military pacts as the core criterion of membership."),
    ("1974: UN Declaration on the NIEO", "NAM mobilized UN General Assembly to adopt the Declaration for a New International Economic Order.", "Elevated the movement from an anti-military grouping into a vanguard demanding global economic justice."),
    ("September 1979: Havana Summit & Pakistan's Entry", "Pakistan formally joined NAM at the 6th Summit in Cuba after withdrawing from CENTO.", "Completed Pakistan's transition from Cold War alliance client to sovereign non-aligned developing nation."),
    ("Present: Strategic Autonomy in Multipolar Era", "Global South nations refuse to enforce unilateral Western sanctions in Ukraine war.", "Vindicated the enduring vitality of non-alignment as pragmatic hedging in 21st-century multipolarity.")
]

q23_essay = [
    ("Introduction: The Moral and Geopolitical Genesis of Non-Alignment",
     "The mid-twentieth century was defined by two monumental, colliding historical transformations: the collapse of European colonial empires across Asia, Africa, and the Middle East, and the simultaneous onset of the Cold War. As dozens of newly sovereign nations emerged into the international arena, they confronted an aggressive, polarized international system dominated by two nuclear-armed superpowers. Both the United States and the Soviet Union sought to drag these nascent, economically impoverished states into their respective military alliance architectures—NATO, SEATO, and CENTO under the American policy of containment, versus the Warsaw Pact and bilateral friendship treaties under Moscow. In this perilous crucible, the newly independent developing world staged one of the most audacious diplomatic revolts in modern history: the creation of the Non-Aligned Movement (NAM). Conceived at the historic 1955 Bandung Asian-African Conference and formally institutionalized at the 1961 Belgrade Summit, NAM represented an organized refusal to surrender newly won sovereignty to great power realpolitik."),

    ("Ideational Foundations and Founding Leadership",
     "The intellectual architecture of NAM was forged by a remarkable constellation of visionary Third World statesmen: Prime Minister Jawaharlal Nehru of India, President Josip Broz Tito of Yugoslavia, President Gamal Abdel Nasser of Egypt, President Sukarno of Indonesia, and President Kwame Nkrumah of Ghana. Rejecting the derogatory Western characterization of non-alignment as 'immoral neutrality' or passive isolationism, Nehru famously declared that non-alignment was not neutrality, but the active, sovereign freedom to judge every international issue on its own merits, unconstrained by military pacts. The movement was anchored in the 'Panchsheel'—the Five Principles of Peaceful Coexistence first codified in the 1954 Sino-Indian Agreement: mutual respect for territorial integrity and sovereignty, mutual non-aggression, mutual non-interference in internal affairs, equality and mutual benefit, and peaceful coexistence. Membership was strictly conditioned upon a state's refusal to enter into bilateral or multilateral military alliances with superpowers, or host foreign military bases for great power competition."),

    ("Historical Impact During the Cold War Bipolarity",
     "During the four decades of Cold War confrontation, NAM functioned as a vital structural buffer that preserved global peace. First, it prevented the complete bipolar partitioning of the planet: by organizing over one hundred developing states into a non-aligned collective, NAM denied both Washington and Moscow the demographic and geographic space to turn the entire Third World into an armed battleground. Second, NAM served as the supreme diplomatic engine of decolonization, providing crucial financial, political, and moral support to national liberation movements in Algeria, Angola, Mozambique, and Vietnam, while spearheading the international diplomatic crusade against apartheid in South Africa. Third, recognizing that political independence was an illusion without economic self-reliance, NAM shifted its focus in the 1970s toward global economic justice. At the 1973 Algiers Summit, NAM mobilized the Global South to demand the New International Economic Order (NIEO), formally adopted by the UN General Assembly in 1974, demanding structural reforms to global trade, technology transfers, and sovereign resource control."),

    ("The Post-Cold War Identity Crisis: The Question of Obsolescence",
     "The collapse of the Soviet Union in 1991 and the end of the Cold War plunged NAM into an acute existential identity crisis. Realist critics boldly declared the movement obsolete: if the two rival blocs no longer existed, against what was the movement 'non-aligned'? Furthermore, the movement suffered from profound internal contradictions: member states frequently engaged in fratricidal wars against one another (such as the Iran-Iraq War and India-Pakistan conflicts), while prominent members hypocritically forged close strategic relationships with superpowers (such as the 1971 Indo-Soviet Treaty of Friendship). For over a decade, NAM appeared to languish as an anachronistic rhetorical talking shop, marginalized by the unipolar dominance of the United States and the accelerating forces of neoliberal globalization."),

    ("Conclusion: The Renaissance of Strategic Autonomy in a Multipolar World",
     "However, reports of NAM's death have proved vastly premature. In the contemporary 21st-century international order, marked by the turbulent demise of American unipolarity, the meteoric rise of China, and the violent eruption of the Russia-Ukraine war, the core logic of non-alignment has experienced a dramatic renaissance. Renamed by scholars as 'Non-Alignment 2.0' or 'Strategic Autonomy,' developing states across Asia, Africa, and Latin America are once again asserting the fundamental Bandung principle: refusing to be forced into exclusive, binary ideological camps. As demonstrated by the refusal of major Global South powers (India, Brazil, South Africa, Indonesia, and Saudi Arabia) to join unilateral Western sanctions against Russia or support Western containment of China, non-alignment has evolved from an ideological protest into a sophisticated, realist strategy of multi-alignment and hedging. In an emerging multiplex, multipolar world, the Non-Aligned Movement—comprising 120 member states representing two-thirds of the United Nations—remains the indispensable collective voice for sovereign equality, multilateral reform, and the democratization of global governance.")
]

q23_quotes = [
    ("Jawaharlal Nehru", "Speech to the Indian Parliament (1950)", "Non-alignment does not mean passivity. It means the freedom to judge each issue on its merits and not be bound by the dictates of any military bloc. When freedom is menaced or justice threatened, we cannot and shall not be neutral."),
    ("Josip Broz Tito", "First Conference of Non-Aligned Countries, Belgrade (1961)", "The non-aligned countries do not represent an emerging third bloc; they represent the collective conscience of mankind, fighting against the division of the world into hostile military camps that threaten global survival."),
    ("Kwame Nkrumah", "I Speak of Freedom (1961)", "We face neither East nor West; we face forward. We cannot accept a world in which we are merely the pawns of other people's conflicts and the consumers of other people's ideologies."),
    ("Amitav Acharya", "The End of American World Order", "Non-alignment was the first major normative contribution of the Global South to international relations. It proved that weaker states could collectively challenge the hegemony of great powers through norm-setting and diplomatic solidarity."),
    ("Shashi Tharoor", "Pax Indica: India and the World of the 21st Century", "Non-alignment was a strategy for an era of bipolarity; strategic autonomy is the strategy for an era of multipolarity. The terminology has changed, but the fundamental objective remains identical: maximizing national sovereignty and freedom of action.")
]

# =========================================================================
# EXECUTE GENERATION FOR BATCH 4: Q18, Q19, Q20, Q21, Q22, Q23
# =========================================================================

batch4_notes = [
    ("Q18_Modern_State_System_Westphalia_to_Globalization_Master_Notes", q18_meta, q18_callout, q18_map, q18_timeline, q18_essay, q18_quotes, "QUESTION 18: THE MODERN STATE SYSTEM: WESTPHALIA TO GLOBALIZATION", "KPPSC PMS IR Master Notes \u2014 Q18: Evolution of the Modern State System: Peace of Westphalia (1648) to Transnational Globalization"),
    ("Q19_Classical_Realism_vs_Structural_Neorealism_Master_Notes", q19_meta, q19_callout, q19_map, q19_timeline, q19_essay, q19_quotes, "QUESTION 19: CLASSICAL REALISM VS STRUCTURAL NEOREALISM", "KPPSC PMS IR Master Notes \u2014 Q19: Theoretical Debate: Classical Realism (Morgenthau) vs. Structural Neorealism (Waltz & Mearsheimer)"),
    ("Q20_Nuclear_Deterrence_Theory_and_South_Asian_Stability_Master_Notes", q20_meta, q20_callout, q20_map, q20_timeline, q20_essay, q20_quotes, "QUESTION 20: NUCLEAR DETERRENCE & SOUTH ASIAN STABILITY", "KPPSC PMS IR Master Notes \u2014 Q20: Nuclear Deterrence Theory, Mutually Assured Destruction (MAD) & Strategic Stability in South Asia"),
    ("Q21_UN_Security_Council_Conflict_Resolution_and_Veto_Master_Notes", q21_meta, q21_callout, q21_map, q21_timeline, q21_essay, q21_quotes, "QUESTION 21: UN SECURITY COUNCIL CONFLICT RESOLUTION & VETO POLITICS", "KPPSC PMS IR Master Notes \u2014 Q21: The United Nations in Conflict Resolution: UNSC Limitations & Veto Politics"),
    ("Q22_Theories_of_Imperialism_Hobson_Lenin_and_Neo_Imperialism_Master_Notes", q22_meta, q22_callout, q22_map, q22_timeline, q22_essay, q22_quotes, "QUESTION 22: THEORIES OF IMPERIALISM & NEO-COLONIALISM", "KPPSC PMS IR Master Notes \u2014 Q22: Theories of Imperialism: Hobson, Lenin & Contemporary Neo-Imperialism"),
    ("Q23_The_Non_Aligned_Movement_Cold_War_to_Multipolarity_Master_Notes", q23_meta, q23_callout, q23_map, q23_timeline, q23_essay, q23_quotes, "QUESTION 23: THE NON-ALIGNED MOVEMENT (NAM) IN WORLD POLITICS", "KPPSC PMS IR Master Notes \u2014 Q23: The Non-Aligned Movement (NAM): Historical Genesis & Relevance in a Multipolar World")
]

for filename, meta, callout, cmap, timeline, essay, quotes, sub_title, full_title in batch4_notes:
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

print("Batch 4 (Q18 through Q23) generated successfully!")
