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
        
        lines = v.split("\n")
        for idx, line in enumerate(lines):
            if idx == 0:
                p1 = c1.paragraphs[0]
            else:
                p1 = c1.add_paragraph()
            p1.paragraph_format.space_before = Pt(1 if len(lines) > 2 else 0)
            p1.paragraph_format.space_after = Pt(1 if len(lines) > 2 else 0)
            r1 = p1.add_run(line)
            r1.font.name = "Georgia"
            r1.font.size = Pt(8.5 if len(lines) > 2 else 9)
            if i == 0:
                r1.font.color.rgb = RGBColor(255, 255, 255)
                r1.font.bold = True
            else:
                r1.font.color.rgb = RGBColor(15, 23, 42)
        
        if i == 0:
            set_cell_background(c0, "0E2B50")
            set_cell_background(c1, "0E2B50")
            r0.font.color.rgb = RGBColor(255, 255, 255)
        else:
            bg = "F8FAFC" if i % 2 == 1 else "FFFFFF"
            set_cell_background(c0, bg)
            set_cell_background(c1, bg)
            r0.font.color.rgb = RGBColor(14, 43, 80) if "EXAM QUESTIONS" in k else RGBColor(30, 41, 59)
            
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_before = Pt(6)
    p_sp.paragraph_format.space_after = Pt(4)
    
    # Callout Box
    add_callout(doc, callout_data)
    
    p_sp2 = doc.add_paragraph()
    p_sp2.paragraph_format.space_before = Pt(8)
    p_sp2.paragraph_format.space_after = Pt(4)
    
    # Flowchart / Map Heading
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
    
    # Timeline / Analytical Matrix Table
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
    
    headers = ["Dimension / Arena", "Strategic Mechanism / Driver", "Systemic Impact on Global Order"]
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
                
    # Model Essay Section
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
        
    # Theoretical / Quotable Perspectives Matrix Table
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
    md_content += "| Dimension / Arena | Strategic Mechanism / Driver | Systemic Impact on Global Order |\n|---|---|---|\n"
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
        
    # Generate HTML and compile PDF via Edge
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
<thead><tr><th>Dimension / Arena</th><th>Strategic Mechanism / Driver</th><th>Systemic Impact on Global Order</th></tr></thead>
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

# ==========================================
# DEFINING Q1: RISE OF CHINA & MULTIPOLARITY
# ==========================================

q1_meta = [
    ("Target Examination", "KPPSC PMS (Provincial Management Service) — Khyber Pakhtunkhwa"),
    ("Paper & Subject", "International Relations (100 Marks Optional — Group 6)"),
    ("EXAM QUESTIONS LINKED:", 
     "• 2025 Q.1: \"The rise of China is widely interpreted as a revival of balance of power politics, signaling a global shift toward multipolarity. Analyze the political, strategic and economic factors that are shaping the new international order. (20 Marks)\"\n"
     "• 2022 Q.3: \"The trade war between the US and China has jeopardized the global economic order impacting all regions and states; Pakistan caught between these two global powers is likely to be affected the most. Analyze Pakistan's predicament and suggest policy options to safeguard its national security objectives. (20 Marks)\"\n"
     "• 2022 Q.5: \"The Chinese Belt and Road Initiative (BRI) has global outreach in Asia, Africa, and Europe at the cost of US hegemony. Discuss its main objectives and analyze challenges and prospects of its success. (20 Marks)\"\n"
     "• 2018 Q.6: \"There may be a possible trade war between the USA and China that would rattle certain American businesses and force consumers to pay more for products. Discuss the term \\\"trade war\\\" and elaborate on why Trump was targeting China. (20 Marks)\""),
    ("Syllabus Section", "Part A: Balance of Power & State System | Part C: Major Powers & Global Order"),
    ("Question Allocation", "Attempt 1 of 5 Required (20 Marks | Strict 35-Minute Allocation)"),
    ("Historiographical / Theoretical Thesis", "The rapid rise of China represents a classic structural power transition (Organski) that ends post-Cold War American unipolarity (Krauthammer) and revives balance-of-power imperatives. As offensive realism (Mearsheimer) predicts, China seeks regional hegemony in the Indo-Pacific while balancing US containment, generating a 'Thucydides Trap' (Allison). Rather than a simple bipolar confrontation, the emerging order is an asymmetric, multiplex multipolarity (Acharya) defined by geo-economic counter-balancing (BRI, BRICS, SCO) and strategic hedging across the Global South.")
]

q1_callout = [
    "STRUCTURAL POWER TRANSITION: China's economic ascendance (overtaking US in PPP terms) has fundamentally shifted the material distribution of capabilities in the international system, triggering systemic balancing dynamics as posited by Kenneth Waltz's Neorealism.",
    "BIPOLARITY AT THE CORE, MULTIPOLARITY AT THE PERIPHERY: While the military and technological frontier forms a dual US-China strategic rivalry, the global political economy is thoroughly multipolar—with the EU, India, Russia, and regional middle powers resisting binary bloc alignment.",
    "GEO-ECONOMIC & INSTITUTIONAL COUNTER-BALANCING: Unlike the USSR's ideological-military balancing, China combines hard-power modernisation (PLAN blue-water navy, A2/AD systems) with infrastructural diplomacy (Belt and Road Initiative) and alternative financial architecture (AIIB, CIPS, BRICS New Development Bank).",
    "EXAMINER WARNING: High scoring in KPPSC PMS requires grounding your analysis in core IR theories—explicitly contrasting John Mearsheimer's Offensive Realism, Graham Allison's Thucydides Trap, G. John Ikenberry's Liberal Order persistence, and Amitav Acharya's Multiplex World."
]

q1_map = r"""
+---------------------------------------------------------------------------------------------------------+
|                  CONCEPTUAL ROADMAP: THE RISE OF CHINA & THE MULTIPOLAR ORDER                           |
+---------------------------------------------------------------------------------------------------------+
                                                     |
                                   [1991: End of Cold War]
                             Krauthammer's 'Unipolar Moment'
                               US Hegemony & Washington Consensus
                                                     |
                                                     v
                       [China's Peaceful Rise & Deng's Pragmatism]
                       Integration into WTO (2001) --> Double-Digit GDP Growth
                       Material Capability Accumulation (Waltzian Realism)
                                                     |
                                                     v
                         [Structural Friction: The Thucydides Trap]
                       Xi Jinping's 'China Dream' & Great Power Diplomacy
                       Power Transition: US Absolute Lead vs. Chinese Parity
                                                     |
           +-----------------------------------------+-----------------------------------------+
           |                                         |                                         |
           v                                         v                                         v
   [POLITICAL ARENA]                         [STRATEGIC ARENA]                         [GEO-ECONOMIC ARENA]
Rejection of Liberal Hegemony              Hard Balancing & Arms Race                Infrastructural Statecraft: BRI
Sovereign Non-Interference                 A2/AD Anti-Access in SCS/Taiwan           Alternative Finance: AIIB, NDB
Expanding BRICS+ & SCO Coalitions          Counter-Balancing: Quad, AUKUS            De-Dollarization & Local Currencies
Saudi-Iran Diplomatic Mediation            Sino-Russian Strategic Axis               Supply Chain Decoupling & CHIPS Act
           |                                         |                                         |
           +-----------------------------------------+-----------------------------------------+
                                                     |
                                                     v
                         [EMERGENT ORDER: ASYMMETRIC MULTIPOLARITY]
                    Neither Unipolar Primacy nor Pure Cold War Bipolarity
                    Multiplex Order: Plural Hegemonies, Strategic Hedging,
                    and Fragmented Global Governance (Amitav Acharya)
+---------------------------------------------------------------------------------------------------------+
"""

q1_timeline = [
    ("Material Distribution of Power", "China's GDP reaches ~$18T; defense modernization expands PLA Navy into the world's largest fleet by hull count.", "Erosion of US absolute unipolar primacy; return of structural anarchy where power parity drives intense systemic balancing (Kenneth Waltz)."),
    ("Indo-Pacific Strategic Architecture", "US establishes Quad (US-India-Japan-Australia) & AUKUS; China consolidates A2/AD bubbles and South China Sea bases.", "Regional security dilemma escalates; balance of threat (Walt) replaces global collective security with militarized containment rings."),
    ("Institutional Statecraft", "Launch of AIIB (100+ members), expansion of BRICS (BRICS+), and consolidation of Shanghai Cooperation Organization (SCO).", "Soft-balancing against Western-dominated Bretton Woods institutions (IMF, World Bank), fostering non-Western diplomatic convergence."),
    ("Geo-economics & Trade Wars", "Belt and Road Initiative (BRI) connects 150+ nations; US responds with tariffs, export controls (CHIPS Act), and 'de-risking'.", "Weaponized interdependence (Drezner/Farrell); fragmentation of global trade into competing techno-economic ecosystems."),
    ("Diplomatic Norms & Mediation", "Beijing-brokered Saudi-Iran Normalization (2023); Global Security Initiative (GSI) promoting sovereign non-interference.", "Challenging Western liberal interventionism; legitimizing a pluralistic, Westphalian-sovereigntist model of conflict management.")
]

q1_essay = [
    ("Introduction: The Demise of Unipolarity and the Return of Structural Balancing",
     "The contemporary international order is undergoing the most profound structural realignment since the dissolution of the Soviet Union in 1991. The post-Cold War era, famously characterized by Charles Krauthammer as the 'unipolar moment,' was defined by uncontested American military primacy, the global diffusion of neo-liberal economics under the Washington Consensus, and the perceived universality of Western institutionalism. However, the meteoric economic, technological, and military ascension of the People's Republic of China (PRC) has conclusively terminated this unipolar interregnum. Through the prism of structural realism, as articulated by Kenneth Waltz, changes in the international system are fundamentally dictated by shifts in the distribution of material capabilities across major powers. China's emergence as the world's second-largest economy—and the largest in Purchasing Power Parity (PPP)—has revived balance-of-power politics as the cardinal operating principle of global affairs. Rather than conforming to Francis Fukuyama's vision of liberal ideological convergence, global politics has entered a phase of intense systemic competition where China's bid for regional hegemony directly collides with American determination to preserve its global primacy, propelling the international system toward an asymmetric, multiplex multipolarity."),
    
    ("Political Drivers: Ideational Pluralism, Sovereignty, and Counter-Hegemonic Coalitions",
     "Politically, the rise of China has fractured the normative consensus of the US-led liberal international order. Beijing champions a rigorous defense of traditional Westphalian sovereignty, explicitly repudiating Western liberal doctrines of conditional sovereignty, democratic interventionism, and human rights conditionality. Through diplomatic frameworks such as the Global Security Initiative (GSI) and the Global Development Initiative (GDI), China promotes a model of 'sovereign pluralism' that resonates powerfully across the Global South. Furthermore, Beijing has systematically spearheaded the institutional expansion of counter-hegemonic platforms—most notably the Shanghai Cooperation Organization (SCO) and the expanded BRICS+ grouping. By providing non-Western powers with credible institutional venues that operate independently of G7 oversight, China has facilitated political balancing against Western hegemony. This diplomatic clout was demonstrated by the historic, Beijing-brokered Saudi-Iran rapprochement in March 2023, signaling China's emergence as a premier diplomatic arbiter capable of re-engineering strategic geometries in regions historically dominated by Western security umbrellas."),
     
    ("Strategic and Military Dynamics: Hard Balancing, Security Dilemmas, and the Indo-Pacific Pivot",
     "Strategically, the international system is witnessing a severe escalation of the security dilemma, particularly across the maritime commons of the Indo-Pacific. In his seminal text The Tragedy of Great Power Politics, John Mearsheimer posits under offensive structural realism that great powers inherently seek to maximize their share of world power, with the ultimate objective of achieving regional hegemony while denying it to rival peers. In accordance with this logic, China has developed formidable Anti-Access/Area-Denial (A2/AD) capabilities—spanning hypersonic glide vehicles, advanced submarine warfare, and militarized island outposts in the South China Sea—designed to push American carrier strike groups beyond the First and Second Island Chains. In response, the United States has enacted classic external balancing strategies, reinvigorating the Quadrilateral Security Dialogue (Quad) with India, Japan, and Australia, and establishing the AUKUS security pact with Britain and Australia. The strategic flashpoint of Taiwan epitomizes Graham Allison's 'Thucydides Trap': a rising power threatening to displace an established hegemon, generating structural friction where miscalculation could precipitate systemic conflagration."),
     
    ("Economic and Geo-economic Realities: Weaponized Interdependence and Parallel Financial Architectures",
     "Economically, balance-of-power politics has migrated from territorial conquest to geo-economic statecraft and the weaponization of interdependence. The Belt and Road Initiative (BRI), encompassing over a trillion dollars in transcontinental infrastructure, represents an unprecedented exercise in structural statecraft, physically reorienting Eurasian and African trade corridors toward Beijing. Simultaneously, recognizing vulnerability to US financial hegemony—manifested in the extraterritorial application of SWIFT sanctions against Russia—China has pioneered alternative institutional and financial clearing architecture. Initiatives such as the Asian Infrastructure Investment Bank (AIIB), the New Development Bank (NDB), and the Cross-Border Interbank Payment System (CIPS), coupled with bilateral local-currency trade settlements (petroyuan accords), represent deliberate soft-balancing against Bretton Woods dominance. While Washington pursues economic 'de-risking,' export restrictions on advanced semiconductors (the CHIPS Act), and tariff barriers, the reality of deep commercial entanglement ensures that global balancing remains economically fragmented rather than strictly bifurcated."),
     
    ("Conclusion: The Contours of an Asymmetric, Multiplex International Order",
     "In conclusion, the rise of China has undeniably catalyzed the revival of balance of power politics, definitively dismantling the unipolar architecture of the post-Cold War world. However, the emergent international order does not represent a simplistic replica of twentieth-century Cold War bipolarity. Because globalized supply chains and multilateral interdependencies persist, the international system is evolving into what Amitav Acharya terms a 'multiplex world'—a decentered, asymmetric multipolarity where the US and China form the principal strategic poles, yet regional middle powers (such as India, Turkey, Saudi Arabia, and Brazil) exercise unprecedented strategic autonomy through multi-alignment and hedging. Ultimately, the stability of this new order will depend on whether the reigning hegemon and the ascending challenger can establish mutual strategic reassurance, managing their balance of power without succumbing to the cataclysmic systemic conflict historically associated with power transitions.")
]

q1_quotes = [
    ("John J. Mearsheimer", "Offensive Structural Realism", "A wealthy China will not be a status quo power; it will try to dominate Asia the way the United States dominates the Western Hemisphere. The result will be an intense security competition with a high probability of war."),
    ("Graham Allison", "Power Transition Theory", "The defining question of global order for our time is whether China and the United States can escape Thucydides's Trap: when a rising power threatens to displace a ruling power, resulting structural stress makes violent clash the rule, not the exception."),
    ("Kenneth N. Waltz", "Neorealism / Structural Realism", "In an anarchic order, balance of power is not a policy choice but an inevitable systemic outcome; whenever power asymmetries emerge, other states will balance to ensure their survival."),
    ("G. John Ikenberry", "Liberal Institutionalism", "The rise of China does not have to trigger the demise of the Western order. The modern international order is unusually dense, institutionalized, and open—making it easy to join and hard to overturn."),
    ("Amitav Acharya", "Constructivist / Post-Western IR", "The emerging global architecture is not a unipolar hegemony, nor a new Cold War bipolarity, but a 'multiplex world'—like a cinema with multiple screens showing different movies, where diverse actors shape order across overlapping issue areas.")
]

# Generate Q1 Master Files
q1_docx = os.path.join(IR_DIR, "Q1_Rise_of_China_Balance_of_Power_and_Multipolarity_Master_Notes.docx")
q1_pdf = os.path.join(IR_DIR, "Q1_Rise_of_China_Balance_of_Power_and_Multipolarity_Master_Notes.pdf")
q1_md = os.path.join(IR_DIR, "Q1_Rise_of_China_Balance_of_Power_and_Multipolarity_Master_Notes.md")

generate_ir_master_doc(
    meta_data=q1_meta,
    callout_data=q1_callout,
    map_text=q1_map,
    timeline_data=q1_timeline,
    essay_sections=q1_essay,
    quotes_data=q1_quotes,
    out_path=q1_docx,
    paper_sub_title="QUESTION 1: RISE OF CHINA, BALANCE OF POWER & MULTIPOLARITY"
)

generate_pdf_and_md(
    meta_data=q1_meta,
    callout_data=q1_callout,
    map_text=q1_map,
    timeline_data=q1_timeline,
    essay_sections=q1_essay,
    quotes_data=q1_quotes,
    pdf_path=q1_pdf,
    md_path=q1_md,
    title="KPPSC PMS IR Master Notes \u2014 Q1: The Rise of China, Balance of Power & Multipolarity"
)

print("Q1 Master Notes generated successfully!")
