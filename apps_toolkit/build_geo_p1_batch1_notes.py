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
P1_DIR = r"H:\My Drive\PMS KPK 2022 notes\3. Geography Paper 1 and Paper 2\Paper 1 - Physical and General Geography"
os.makedirs(P1_DIR, exist_ok=True)

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

def generate_geo_master_doc(meta_data, callout_data, map_text, timeline_data, essay_sections, quotes_data, out_path, paper_sub_title):
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
    run_main = p_title.add_run("KPPSC PMS EXAMINATION — MASTER STUDY NOTES\n")
    run_main.font.name = "Arial"
    run_main.font.size = Pt(13)
    run_main.font.bold = True
    run_main.font.color.rgb = RGBColor(14, 43, 80)
    
    run_sub = p_title.add_run(f"SUBJECT: GEOGRAPHY (PAPER-I: PHYSICAL & GENERAL GEOGRAPHY) — 100 MARKS\n{paper_sub_title}")
    run_sub.font.name = "Arial"
    run_sub.font.size = Pt(11)
    run_sub.font.color.rgb = RGBColor(27, 73, 101)
    
    tbl_meta = doc.add_table(rows=len(meta_data), cols=2)
    tbl_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_meta.autofit = False
    
    for idx, (k, v) in enumerate(meta_data):
        row = tbl_meta.rows[idx]
        cell_k, cell_v = row.cells[0], row.cells[1]
        cell_k.width = Inches(2.2)
        cell_v.width = Inches(4.3)
        set_cell_background(cell_k, "0E2B50")
        set_cell_background(cell_v, "F4F6F9")
        set_cell_margins(cell_k, top=60, bottom=60, left=100, right=100)
        set_cell_margins(cell_v, top=60, bottom=60, left=100, right=100)
        
        p_k = cell_k.paragraphs[0]
        p_k.paragraph_format.space_before = Pt(0)
        p_k.paragraph_format.space_after = Pt(0)
        r_k = p_k.add_run(k)
        r_k.font.name = "Arial"
        r_k.font.size = Pt(9.5)
        r_k.font.bold = True
        r_k.font.color.rgb = RGBColor(255, 255, 255)
        
        p_v = cell_v.paragraphs[0]
        p_v.paragraph_format.space_before = Pt(0)
        p_v.paragraph_format.space_after = Pt(0)
        r_v = p_v.add_run(v)
        r_v.font.name = "Georgia"
        r_v.font.size = Pt(9.5)
        r_v.font.color.rgb = RGBColor(30, 41, 59)
        
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    add_callout(doc, callout_data, "EXAM ESSENTIAL TAKEAWAY")
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    
    p_hm = doc.add_paragraph()
    p_hm.paragraph_format.space_before = Pt(8)
    p_hm.paragraph_format.space_after = Pt(4)
    r_hm = p_hm.add_run("1. CONCEPTUAL MIND MAP & THEMATIC ROADMAP")
    r_hm.font.name = "Arial"
    r_hm.font.size = Pt(11)
    r_hm.font.bold = True
    r_hm.font.color.rgb = RGBColor(14, 43, 80)
    
    tbl_mm = doc.add_table(rows=1, cols=1)
    tbl_mm.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_mm.autofit = False
    c_mm = tbl_mm.cell(0, 0)
    c_mm.width = Inches(6.5)
    set_cell_background(c_mm, "F8FAFC")
    set_cell_margins(c_mm, top=100, bottom=100, left=140, right=140)
    p_mm_c = c_mm.paragraphs[0]
    p_mm_c.paragraph_format.space_before = Pt(0)
    p_mm_c.paragraph_format.space_after = Pt(0)
    r_mm_code = p_mm_c.add_run(map_text)
    r_mm_code.font.name = "Consolas"
    r_mm_code.font.size = Pt(8.5)
    r_mm_code.font.color.rgb = RGBColor(15, 23, 42)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    
    p_ht = doc.add_paragraph()
    p_ht.paragraph_format.space_before = Pt(8)
    p_ht.paragraph_format.space_after = Pt(4)
    r_ht = p_ht.add_run("2. ANALYTICAL DIMENSIONS & GEOMORPHIC/SYSTEMIC MATRIX")
    r_ht.font.name = "Arial"
    r_ht.font.size = Pt(11)
    r_ht.font.bold = True
    r_ht.font.color.rgb = RGBColor(14, 43, 80)
    
    tbl_time = doc.add_table(rows=len(timeline_data), cols=3)
    tbl_time.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_time.autofit = False
    
    for r_idx, row in enumerate(tbl_time.rows):
        d_phase, d_theme, d_impact = timeline_data[r_idx]
        c0, c1, c2 = row.cells[0], row.cells[1], row.cells[2]
        c0.width = Inches(1.8)
        c1.width = Inches(2.2)
        c2.width = Inches(2.5)
        
        is_hdr = (r_idx == 0)
        bg = "0E2B50" if is_hdr else ("EDF2F7" if r_idx % 2 == 1 else "FFFFFF")
        fg = RGBColor(255, 255, 255) if is_hdr else RGBColor(30, 41, 59)
        
        for c, text in [(c0, d_phase), (c1, d_theme), (c2, d_impact)]:
            set_cell_background(c, bg)
            set_cell_margins(c, top=60, bottom=60, left=80, right=80)
            p = c.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(text)
            r.font.name = "Arial" if is_hdr else "Georgia"
            r.font.size = Pt(9 if not is_hdr else 9.5)
            r.font.bold = is_hdr
            r.font.color.rgb = fg
            
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    p_he = doc.add_paragraph()
    p_he.paragraph_format.space_before = Pt(8)
    p_he.paragraph_format.space_after = Pt(4)
    r_he = p_he.add_run("3. 20-MARK MASTER MODEL ESSAY (KPPSC PMS BENCHMARK)")
    r_he.font.name = "Arial"
    r_he.font.size = Pt(11)
    r_he.font.bold = True
    r_he.font.color.rgb = RGBColor(14, 43, 80)
    
    for subhead, body in essay_sections:
        p_sh = doc.add_paragraph()
        p_sh.paragraph_format.space_before = Pt(6)
        p_sh.paragraph_format.space_after = Pt(2)
        r_sh = p_sh.add_run(subhead)
        r_sh.font.name = "Arial"
        r_sh.font.size = Pt(10)
        r_sh.font.bold = True
        r_sh.font.color.rgb = RGBColor(27, 73, 101)
        
        p_bd = doc.add_paragraph()
        p_bd.paragraph_format.space_before = Pt(0)
        p_bd.paragraph_format.space_after = Pt(5)
        p_bd.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r_bd = p_bd.add_run(body)
        r_bd.font.name = "Georgia"
        r_bd.font.size = Pt(9.5)
        r_bd.font.color.rgb = RGBColor(30, 41, 59)
        
    p_hq = doc.add_paragraph()
    p_hq.paragraph_format.space_before = Pt(8)
    p_hq.paragraph_format.space_after = Pt(4)
    r_hq = p_hq.add_run("4. SCHOLARLY PERSPECTIVES & GEOMORPHIC AXIOMS TABLE")
    r_hq.font.name = "Arial"
    r_hq.font.size = Pt(11)
    r_hq.font.bold = True
    r_hq.font.color.rgb = RGBColor(14, 43, 80)
    
    tbl_q = doc.add_table(rows=len(quotes_data), cols=3)
    tbl_q.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_q.autofit = False
    
    for q_idx, row in enumerate(tbl_q.rows):
        auth, schl, quote = quotes_data[q_idx]
        c0, c1, c2 = row.cells[0], row.cells[1], row.cells[2]
        c0.width = Inches(1.5)
        c1.width = Inches(1.8)
        c2.width = Inches(3.2)
        
        is_hdr = (q_idx == 0)
        bg = "0E2B50" if is_hdr else ("EDF2F7" if q_idx % 2 == 1 else "FFFFFF")
        fg = RGBColor(255, 255, 255) if is_hdr else RGBColor(30, 41, 59)
        
        for c, text in [(c0, auth), (c1, schl), (c2, quote)]:
            set_cell_background(c, bg)
            set_cell_margins(c, top=60, bottom=60, left=80, right=80)
            p = c.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(text)
            r.font.name = "Arial" if is_hdr else "Georgia"
            r.font.size = Pt(9 if not is_hdr else 9.5)
            r.font.bold = is_hdr
            r.font.color.rgb = fg
            
    doc.save(out_path)
    print(f"Generated Master Docx: {out_path}")

def generate_pdf_and_md(meta_data, callout_data, map_text, timeline_data, essay_sections, quotes_data, pdf_path, md_path, title):
    md_content = f"# {title}\n\n"
    md_content += "## KPPSC PMS Examination — Master Study Notes (Paper-I: Physical & General Geography)\n\n"
    
    md_content += "| Field | Value |\n|---|---|\n"
    for k, v in meta_data:
        md_content += f"| **{k}** | {v} |\n"
    md_content += "\n---\n\n"
    
    md_content += "### 📌 EXAM ESSENTIAL TAKEAWAY\n"
    for line in callout_data:
        md_content += f"> {line}\n"
    md_content += "\n---\n\n"
    
    md_content += "### 1. Conceptual Mind Map & Thematic Roadmap\n```\n" + map_text + "\n```\n\n---\n\n"
    
    md_content += "### 2. Analytical Dimensions & Geomorphic/Systemic Matrix\n\n"
    headers = timeline_data[0]
    md_content += f"| {headers[0]} | {headers[1]} | {headers[2]} |\n|---|---|---|\n"
    for row in timeline_data[1:]:
        md_content += f"| **{row[0]}** | {row[1]} | {row[2]} |\n"
    md_content += "\n---\n\n"
    
    md_content += "### 3. 20-Mark Master Model Essay\n\n"
    for sh, bd in essay_sections:
        md_content += f"#### {sh}\n{bd}\n\n"
    md_content += "---\n\n"
    
    md_content += "### 4. Scholarly Perspectives & Geomorphic Axioms\n\n"
    q_hdr = quotes_data[0]
    md_content += f"| {q_hdr[0]} | {q_hdr[1]} | {q_hdr[2]} |\n|---|---|---|\n"
    for row in quotes_data[1:]:
        md_content += f"| **{row[0]}** | {row[1]} | *\"{row[2]}\"* |\n"
        
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(md_content)
        
    # Build HTML for PDF
    html_content = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>{title}</title>
<style>
    @page {{ size: A4; margin: 18mm 16mm 18mm 16mm; }}
    body {{ font-family: 'Segoe UI', Calibri, Arial, sans-serif; font-size: 10pt; line-height: 1.45; color: #1A1A1A; margin: 0; }}
    h1 {{ color: #0E2B50; font-size: 15pt; border-bottom: 2px solid #0E2B50; padding-bottom: 3px; margin: 0 0 8px 0; text-align: center; }}
    h2 {{ color: #1B4965; font-size: 11pt; margin: 10px 0 6px 0; border-bottom: 1px solid #D0D7DE; padding-bottom: 2px; text-align: center; }}
    h3 {{ color: #0E2B50; font-size: 10.5pt; margin: 12px 0 4px 0; }}
    h4 {{ color: #1B4965; font-size: 10pt; margin: 8px 0 2px 0; }}
    table {{ width: 100%; border-collapse: collapse; margin: 8px 0; font-size: 9pt; }}
    th {{ background-color: #0E2B50; color: #FFFFFF; padding: 4px 8px; border: 1px solid #0E2B50; text-align: left; }}
    td {{ padding: 4px 8px; border: 1px solid #D0D7DE; vertical-align: top; }}
    tr:nth-child(even) td {{ background-color: #F8FAFC; }}
    blockquote {{ border-left: 4px solid #0E2B50; background-color: #F0F4F8; padding: 6px 12px; margin: 8px 0; font-size: 9.5pt; }}
    pre {{ background-color: #F8FAFC; border: 1px solid #E2E8F0; padding: 8px; font-family: 'Consolas', monospace; font-size: 8.5pt; line-height: 1.3; overflow-x: auto; }}
    p {{ margin: 4px 0 6px 0; text-align: justify; }}
    .footer {{ margin-top: 20px; font-size: 8pt; color: #666; text-align: center; border-top: 1px solid #E1E4E8; padding-top: 4px; }}
</style>
</head>
<body>
<h1>{title}</h1>
<h2>KPPSC PMS Master Examination Notes — Paper-I (Physical & General Geography)</h2>

<table>
<tbody>
{"".join(f"<tr><th style='width: 30%;'>{k}</th><td>{v}</td></tr>" for k, v in meta_data)}
</tbody>
</table>

<blockquote>
<strong style="color: #0E2B50;">📌 EXAM ESSENTIAL TAKEAWAY</strong><br/>
{"<br/>".join(callout_data)}
</blockquote>

<h3>1. Conceptual Mind Map & Thematic Roadmap</h3>
<pre>{map_text}</pre>

<h3>2. Analytical Dimensions & Geomorphic/Systemic Matrix</h3>
<table>
<thead>
<tr>{"".join(f"<th>{c}</th>" for c in timeline_data[0])}</tr>
</thead>
<tbody>
{"".join(f"<tr><td><strong>{r[0]}</strong></td><td>{r[1]}</td><td>{r[2]}</td></tr>" for r in timeline_data[1:])}
</tbody>
</table>

<h3>3. 20-Mark Master Model Essay (KPPSC PMS Benchmark)</h3>
{"".join(f"<h4>{sh}</h4><p>{bd}</p>" for sh, bd in essay_sections)}

<h3>4. Scholarly Perspectives & Geomorphic Axioms</h3>
<table>
<thead>
<tr>{"".join(f"<th>{c}</th>" for c in quotes_data[0])}</tr>
</thead>
<tbody>
{"".join(f"<tr><td><strong>{r[0]}</strong></td><td>{r[1]}</td><td><em>\"{r[2]}\"</em></td></tr>" for r in quotes_data[1:])}
</tbody>
</table>

<div class="footer">KPPSC PMS Physical Geography Master Suite — Verified Wordcraft & Print Standard</div>
</body>
</html>
"""
    temp_html = pdf_path.replace(".pdf", "_temp.html")
    with open(temp_html, "w", encoding="utf-8") as f:
        f.write(html_content)
        
    cmd = [
        EDGE_PATH,
        "--headless",
        "--disable-gpu",
        "--run-all-compositor-stages-before-draw",
        f"--print-to-pdf={pdf_path}",
        temp_html
    ]
    subprocess.run(cmd, check=True)
    if os.path.exists(temp_html):
        os.remove(temp_html)
    print(f"PDF and MD successfully generated: {pdf_path}")

print("Geography Batch 1 builder framework loaded successfully.")

# ==============================================================================
# QUESTION 1 (P1_Q1): INTERNAL STRUCTURE OF THE EARTH & SEISMIC DISCONTINUITIES
# Solved from: 2025 Q.1, 2016 Q.3, 2008 Q.2, 2008 Q.5
# ==============================================================================
p1_q1_meta = [
    ("Target Syllabus", "Section 2: Landforms — Internal Structure of the Earth & Seismology"),
    ("Exam Years Solved", "KPPSC PMS 2025 Q.1, 2016 Q.3, 2008 Q.2, 2008 Q.5"),
    ("Marks & Allocation", "20 Marks | Strict 35-Minute Descriptive Writing Limit"),
    ("Core Theoretical Thesis", "The Earth's concentric internal architecture is deciphered indirectly via the refraction, reflection, and shadow-zone mechanics of body seismic waves (P and S waves). The abrupt velocity shifts across first-order discontinuities (Conrad, Mohorovičić, Repetti, Gutenberg, Lehmann) demarcate sharp rheological and compositional boundaries—transitioning from brittle silicate crust to plastic peridotite mantle, liquid outer iron-nickel core, and crystalline solid inner core.")
]

p1_q1_callout = [
    "SEISMIC PROBING OF THE INTERIOR: Direct borehole drilling (e.g., Kola Superdeep Borehole at 12.26 km) penetrates less than 0.2% of Earth's 6,371 km radius. Geoscientists rely on body seismic waves generated by high-magnitude earthquakes as natural tomographic probes.",
    "P-WAVES VS. S-WAVES DICHOTOMY: Primary (P) waves are compressional longitudinal waves (propagating through solids, liquids, and gases), whose velocity varies with incompressibility and density (Vp = sqrt[(K + 4/3*mu)/rho]). Secondary (S) waves are transverse shear waves that cannot propagate through liquids (shear modulus mu = 0).",
    "THE TWO FIRST-ORDER DISCONTINUITIES: The Mohorovičić Discontinuity (Moho, ~30–70 km under continents, 5–10 km under oceans) marks the crust-mantle chemical boundary where P-wave velocity jumps from ~6.5 to ~8.1 km/s. The Gutenberg Discontinuity (2,891 km) marks the core-mantle boundary where S-waves abruptly terminate and P-waves refract sharply downward, creating the diagnostic seismic shadow zones.",
    "CORE COMPOSITION & GEODYNAMO: The core's high density (~10–13 g/cm³) and magnetic properties confirm an iron-nickel (Nife) composition with trace light elements (silicon, sulfur, oxygen). Convection currents in the liquid outer core drive the geodynamo that generates Earth's geomagnetic field."
]

p1_q1_map = """
                   [SEISMIC TOMOGRAPHY & EARTH'S INTERIOR]
                                      │
            ┌─────────────────────────┴─────────────────────────┐
            ▼                                                   ▼
  [COMPOSITIONAL LAYERS]                              [MECHANICAL/RHEOLOGICAL]
  • Crust: Silicic Sial (Granitic)                    • Lithosphere: Rigid / Brittle
    & Mafic Sima (Basaltic)                             (0-100 km; tectonic plates)
  • Mantle: Ultramafic Peridotite                     • Asthenosphere: Plastic / Semi-fluid
    (Fe-Mg Silicates; 84% volume)                       (100-410 km; low-velocity zone)
  • Core: Nife Metallic Alloy                         • Mesosphere: Solid lower mantle
    (85% Fe, 5% Ni, 10% light elements)               • Outer Core: Convecting Liquid
                                                      • Inner Core: Solid Crystalline
                                      │
                         [FIRST-ORDER DISCONTINUITIES]
  ┌───────────────────────────────────┼───────────────────────────────────┐
  ▼                                   ▼                                   ▼
[MOHOROVIČIĆ (MOHO)]          [GUTENBERG (CMB)]                   [LEHMANN BOUNDARY]
• Depth: ~35 km (avg)         • Depth: 2,891 km                   • Depth: 5,150 km
• Crust to Upper Mantle       • Mantle to Outer Core              • Outer Core to Inner Core
• P-wave: 6.5 -> 8.1 km/s     • S-wave drops to 0 km/s            • P-wave velocity jumps
• S-wave: 3.8 -> 4.6 km/s     • P-wave shadow: 104°-140°          • Liquid to solid NiFe
"""

p1_q1_timeline = [
    ("Internal Boundary / Layer", "Depth & Physical State", "Seismic Behavior & Petrographic Evidence"),
    ("Continental & Oceanic Crust", "0–35 km (avg); 5–10 km oceanic, 30–70 km continental (solid)", "P-waves travel at 6.0–6.8 km/s; granitic Sial (2.7 g/cm³) transitions across Conrad discontinuity to basaltic Sima (3.0 g/cm³)."),
    ("Mohorovičić Discontinuity (Moho)", "Depth: 30–70 km under continents; 7–10 km under oceanic basins", "First-order seismic jump discovered by Andrija Mohorovičić (1909); P-wave speed leaps from ~6.5 to 8.1 km/s as rock density shifts to ultramafic peridotite."),
    ("Asthenosphere (Low-Velocity Zone)", "100–410 km (partial melt: 1–5% liquid; ductile plastic)", "Pronounced drop in P and S wave velocities due to rheological softening; enables isostatic adjustment and lithospheric plate motion."),
    ("Mesosphere (Lower Mantle)", "660–2,891 km (solid high-density peridotite/bridgmanite)", "Continuous velocity increase under extreme confining pressure; P-waves reach ~13.7 km/s; terminates at D'' layer above core boundary."),
    ("Gutenberg Discontinuity (CMB)", "Depth: 2,891 km (mantle-to-core boundary)", "Identified by Beno Gutenberg (1914); S-waves vanish completely (shear modulus = 0); P-wave velocity plummets from 13.7 to 8.1 km/s due to liquid iron state."),
    ("Outer Core", "2,891–5,150 km (molten/liquid iron-nickel alloy)", "True liquid medium (density 9.9–12.2 g/cm³); vigorous thermal convection sustains geodynamo; creates global S-wave shadow zone beyond 104° epicentral distance."),
    ("Lehmann Discontinuity & Inner Core", "5,150–6,371 km (solid crystalline iron-nickel sphere)", "Discovered by Inge Lehmann (1936); weak P-wave reflections confirm solid state; P-wave velocity jumps from 10.3 to 11.2 km/s due to 3.6 million atm pressure.")
]

p1_q1_essay = [
    ("Introduction: The Subterranean Frontier and Seismological Remote Sensing",
     "Human direct empirical observation of the Earth's interior remains confined to the outermost skin of the lithosphere. The deepest artificial excavation in human history—the Kola Superdeep Borehole in Russia—penetrated a mere 12.26 kilometers, representing less than 0.2 percent of the planet's 6,371-kilometer mean radius. Consequently, understanding the deep interior relies upon geophysical remote sensing, predominantly the analysis of body seismic waves produced by high-energy tectonic earthquakes. By monitoring how Primary (P) compressional waves and Secondary (S) shear waves refract, reflect, accelerate, and extinguish as they traverse the Earth, seismologists reconstruct the concentric compositional and mechanical layers of our planet. This seismological tomography reveals a dynamic, highly differentiated terrestrial engine consisting of an outer brittle crust, a silicate plastic mantle, a convecting liquid metallic outer core, and a super-pressurized, crystalline solid inner core."),

    ("Mechanics of Body Seismic Waves: The Diagnostic Probes of Earth's Interior",
     "Seismic waves generated at earthquake foci divide fundamentally into surface waves (Love and Rayleigh waves) and body waves (Primary 'P' and Secondary 'S' waves). For deep interior profiling, body waves serve as the definitive diagnostic tools. Primary waves are longitudinal, compressional stress waves that vibrate rock particles parallel to the direction of wave propagation. Mathematically, P-wave velocity is governed by the formula Vp = sqrt[(K + 4/3*mu)/rho], where K represents the bulk modulus (incompressibility), mu is the shear modulus (rigidity), and rho is the density. Because both solids, liquids, and gases resist compression (K > 0), P-waves can propagate through all states of matter, though their velocity drops sharply when passing from solid to liquid due to the collapse of shear rigidity. In contrast, Secondary waves are transverse shear waves vibrating perpendicular to propagation, with velocity expressed as Vs = sqrt(mu/rho). Crucially, liquids possess zero resistance to shear stress (mu = 0), rendering S-waves physically incapable of traversing molten or liquid media. The differential travel times, refraction trajectories (governed by Snell’s Law in spherical coordinates), and shadow zones of these two wave types provide precise, mathematical evidence of deep structural boundaries."),

    ("The Concentric Interior: Crust, Mantle, and Rheological Stratification",
     "The Earth exhibits a twofold classification: a chemical/compositional division (Crust, Mantle, Core) and a mechanical/rheological division (Lithosphere, Asthenosphere, Mesosphere, Outer Core, Inner Core). The outermost crust comprises two distinct petrographic domains: the continental crust, averaging 35 km in thickness (reaching 70 km beneath orogenic belts like the Himalayas), composed of low-density (2.7 g/cm³) granitic-andesitic rocks rich in silica and aluminum (Sial); and the oceanic crust, 5 to 10 km thick, consisting of dense (3.0 g/cm³) basaltic-gabbroic rocks rich in silica and magnesium (Sima). Beneath the crust lies the mantle, extending to a depth of 2,891 km and constituting roughly 84 percent of Earth's volume. Composed primarily of ultramafic peridotite (iron- and magnesium-rich silicates), the mantle is subdivided mechanically. The uppermost solid mantle combines with the crust to form the rigid, brittle Lithosphere (0–100 km). Directly beneath lies the Asthenosphere (100–410 km), a zone of high geothermal heat where 1 to 5 percent partial melting occurs. This creates the seismic 'Low-Velocity Zone' (LVZ), imparting plastic, ductile rheology upon which lithospheric plates glide. Beneath the 660-km transition zone lies the solid Mesosphere, where intense lithostatic pressure prevents rock melting despite temperatures exceeding 2,500°C."),

    ("First-Order Seismic Discontinuities and Shadow Zone Geometry",
     "The internal layering is demarcated by first-order seismic discontinuities where wave velocities alter abruptly. The Mohorovičić Discontinuity ('Moho'), discovered in 1909 by Croatian seismologist Andrija Mohorovičić, separates the crust from the upper mantle. At the Moho, P-wave velocities leap abruptly from ~6.5 km/s to 8.1 km/s, reflecting the compositional transition to dense peridotite. In 1914, German geophysicist Beno Gutenberg identified the Core-Mantle Boundary (CMB) at 2,891 km depth. At the Gutenberg Discontinuity, S-waves terminate completely, demonstrating that the outer core is liquid, while P-waves suffer a catastrophic velocity collapse from ~13.7 km/s to ~8.1 km/s. This extreme refraction bends P-waves inward, producing a distinct 'P-wave Shadow Zone' between 104° and 140° angular distance from the earthquake epicenter. Furthermore, because S-waves cannot traverse the liquid core at all, an expansive 'S-wave Shadow Zone' extends across the entire hemisphere beyond 104° to 180°. In 1936, Danish seismologist Inge Lehmann discovered faint P-wave reflections within the shadow zone, proving the existence of a solid Inner Core at a depth of 5,150 km, where P-wave velocity abruptly accelerates back to 11.2 km/s."),

    ("The Core Composition, Iron Geochemistry, and the Geodynamo Mechanism",
     "Multiple independent lines of scientific reasoning validate that the Earth's core is predominantly composed of an iron-nickel alloy (Nife). First, global gravitational measurements and orbital dynamics determine Earth's mean density to be 5.515 g/cm³, whereas superficial crustal rocks average only 2.7 to 3.0 g/cm³, mathematically demanding a hyper-dense core (10.0 to 13.0 g/cm³). Second, geochemical analyses of primitive chondritic meteorites—representing the pristine primordial material from which the solar system accreted—demonstrate high metallic iron and nickel concentrations that must have segregated into the planetary center during Earth's early differentiation (the 'iron catastrophe'). Third, laboratory diamond-anvil cell experiments confirm that under pressures exceeding 3.3 to 3.6 million atmospheres and temperatures of 5,000°C to 6,000°C, iron assumes a hexagonal close-packed (hcp) solid crystal structure matching the inner core's seismic velocity. The liquid outer core's thermal and compositional convection, combined with Coriolis forces generated by planetary rotation, acts as a self-exciting geodynamo. This vigorous circulation of molten conductive iron sustains the geomagnetic dipolar field, shielding the biosphere from lethal solar cosmic radiation."),

    ("Conclusion: The Unified Seismotectonic Architecture of the Planet",
     "In evaluative summation, the internal structure of the Earth represents a masterwork of thermal differentiation and geochemical stratification. Seismic wave mechanics have transformed our understanding of deep Earth from mythological speculation into rigorous quantitative geophysics. The behavior of P and S waves across the Moho, Gutenberg, and Lehmann discontinuities provides unambiguous, replicable evidence of solid, plastic, liquid, and crystalline zones. Far from being a static interior, the deep Earth operates as a coupled thermodynamic heat engine: mantle plumes rise from the D'' thermal boundary layer above the Gutenberg discontinuity, subducting tectonic slabs sink to the core-mantle interface, and outer-core convection generates the protective magnetosphere. Mastery of these seismic principles is foundational to explaining surface volcanism, orogenesis, earthquake distribution in active belts like Pakistan's Chaman and MBT fault lines, and the overarching paradigm of plate tectonics.")
]

p1_q1_quotes = [
    ("Scholar / Geophysicist", "Theoretical Contribution / Concept", "Exact Quotable Academic Thesis / Axiom"),
    ("Andrija Mohorovičić (1909)", "Discovery of Crust-Mantle Boundary", "The sudden refraction and velocity jump of seismic waves at shallow depth proves an abrupt compositional transition from crustal silicates to dense ultramafic mantle rock."),
    ("Beno Gutenberg (1914)", "Core-Mantle Boundary & S-Wave Extinction", "The complete cessation of transverse shear waves at 2,891 km constitutes irrefutable mechanical proof that the Earth's outer core behaves as a true hydrodynamic liquid."),
    ("Inge Lehmann (1936)", "Discovery of the Solid Inner Core", "Weak P-wave arrivals refracted into the core shadow zone demonstrate the existence of a distinct, highly rigid inner core suspended within the molten outer envelope."),
    ("Arthur Holmes (1944)", "Principles of Physical Geology", "The Earth's internal thermal engine, sustained by radiogenic decay and primordial heat, drives convection currents in the plastic substratum that dictate all surface tectonic drama."),
    ("Francis Birch (1952)", "Elasticity and Constitution of the Earth's Interior", "High-pressure shockwave physics confirms that the core consists primarily of metallic iron alloyed with nickel and approximately ten percent light elements like silicon, sulfur, or oxygen.")
]

# ==============================================================================
# QUESTION 2 (P1_Q2): PLATE TECTONICS, PLATE BOUNDARIES & OROGENESIS
# Solved from: 2016 Q.3, 2010 Q.4, 2008 Q.2
# ==============================================================================
p1_q2_meta = [
    ("Target Syllabus", "Section 2: Landforms — Plate Tectonics, Earthquakes & Volcanicity"),
    ("Exam Years Solved", "KPPSC PMS 2016 Q.3, 2010 Q.4, 2008 Q.2"),
    ("Marks & Allocation", "20 Marks | Strict 35-Minute Descriptive Writing Limit"),
    ("Core Theoretical Thesis", "The Plate Tectonics paradigm synthesizes Continental Drift and Seafloor Spreading into a unifying kinematic model. The Earth's brittle lithosphere is fragmented into discrete mosaic plates that translate horizontally over the ductile asthenosphere. The tri-fold collision boundaries—Oceanic-Continental, Oceanic-Oceanic, and Continental-Continental—generate distinct tectonic regimes of Benioff-zone seismicity, calc-alkaline volcanism, and collisional orogenesis, exemplified regionally by the ongoing Indo-Eurasian collision shaping Pakistan's northern mountain arc.")
]

p1_q2_callout = [
    "PARADIGM SYNTHESIS: Plate Tectonics (formulated by Tuzo Wilson, McKenzie, Parker, and Morgan in the late 1960s) reconciled Wegener’s Continental Drift (1912) and Harry Hess’s Seafloor Spreading (1960) by identifying the lithosphere—not continents alone—as the mobile unit.",
    "THE THREE COLLISIONAL REGIMES: (1) Oceanic-Continental Subduction (e.g., Nazca-South American plate forming the Andes; generates deep Benioff earthquakes, trench, and volcanic arc); (2) Oceanic-Oceanic Subduction (e.g., Pacific-Marianas; generates deep ocean trenches and volcanic island arcs); (3) Continental-Continental Collision (e.g., Indo-Eurasian collision; produces double-thickness crust, extensive thrust faulting, and non-volcanic Himalayan orogeny).",
    "DRIVING MECHANISMS: Slab pull (negative buoyancy of cold, dense subducting oceanic lithosphere) is the dominant driving force, supported by ridge push (gravitational sliding away from elevated mid-ocean ridges) and basal mantle convection drag.",
    "PAKISTAN TECTONIC CONTEXT: Pakistan sits at the hyper-active triple junction of the Indian Plate, Eurasian Plate, and Arabian Plate. The collision drives the Main Boundary Thrust (MBT), Main Karakoram Thrust (MKT), and the left-lateral Chaman Strike-Slip Fault system responsible for major earthquakes (1935 Quetta, 2005 Kashmir, 2013 Awaran)."
]

p1_q2_map = """
                       [THE PLATE TECTONIC PARADIGM]
                                     │
           ┌─────────────────────────┴─────────────────────────┐
           ▼                                                   ▼
 [DIVERGENT BOUNDARIES]                              [CONVERGENT BOUNDARIES]
 (Constructive: Ridge Push)                          (Destructive: Slab Pull)
 • Mid-Atlantic Ridge                                          │
 • East African Rift                                           │
                                     ┌─────────────────────────┼─────────────────────────┐
                                     ▼                         ▼                         ▼
                           [OCEANIC-CONTINENTAL]     [OCEANIC-OCEANIC]         [CONTINENTAL-CONTINENTAL]
                           • Dense plate subducts    • Older/colder plate      • Buoyant crust resists
                           • Trench + Magmatic Arc     subducts                  subduction
                           • Wadati-Benioff Zone     • Deep trenches           • Intense folding & thrusting
                           • E.g., Andes Mountains   • Volcanic Island Arcs    • Double-thickness crust
                                                     • E.g., Japan, Marianas   • Non-volcanic orogeny
                                                                               • E.g., Himalayas & Karakoram
                                     │
                          [TRANSFORM BOUNDARIES]
                          (Conservative: Shear stress)
                          • E.g., San Andreas Fault, Chaman Fault (Pakistan)
"""

p1_q2_timeline = [
    ("Plate Boundary Type", "Geodynamic Mechanism & Stress Regime", "Diagnostic Landforms & Seismotectonic Manifestations"),
    ("Divergent (Constructive)", "Tensional stress; asthenospheric upwelling, decompression melting; seafloor spreading.", "Mid-ocean ridge systems, axial rift valleys (e.g., Mid-Atlantic Ridge, East African Rift); shallow-focus basaltic volcanism and normal fault earthquakes."),
    ("Oceanic-Continental Convergent", "Compressional stress; dense basaltic oceanic plate subducts beneath buoyant granitic continent at trench.", "Deep ocean trench (e.g., Peru-Chile Trench), continental magmatic/volcanic arc (Andes); Wadati-Benioff seismic zone with shallow to deep-focus earthquakes (up to 700 km)."),
    ("Oceanic-Oceanic Convergent", "Compressional stress; older, colder, denser oceanic lithosphere subducts beneath younger oceanic plate.", "Deep abyssal trenches (e.g., Mariana Trench, 11 km deep); curved volcanic island arcs (Aleutian Islands, Japan archipelago); back-arc basins and explosive andesitic volcanism."),
    ("Continental-Continental Convergent", "Intense compressional stress; neither plate subducts due to low density (buoyancy); crustal shortening and thickening.", "Massive collisional mountain ranges (Himalayas, Karakoram, Alps); crustal root doubling up to 70–80 km depth; intense shallow/intermediate thrust earthquakes; absence of active volcanism."),
    ("Transform Fault (Conservative)", "Horizontal shear stress; plates slide past each other along strike-slip transform faults.", "Linear strike-slip fault valleys, offset stream channels, shutter ridges; shallow-focus, high-magnitude shear earthquakes (e.g., San Andreas Fault, Chaman Fault in Balochistan)."),
    ("Indo-Eurasian Regional Collision", "Ongoing northward indentation of Indian plate into Eurasia at ~4–5 cm/year since early Eocene (~50 Ma).", "Formation of Himalayan-Karakoram orogen; Main Boundary Thrust (MBT), Main Mantle Thrust (MMT); deadly regional seismicity (2005 Muzaffarabad earthquake, Mw 7.6; Chaman Fault system).")
]

p1_q2_essay = [
    ("Introduction: The Unified Kinematic Architecture of the Lithosphere",
     "The theory of Plate Tectonics, formulated in the late 1960s through the synthesis of Alfred Wegener’s Continental Drift hypothesis (1912) and Harry Hess’s Seafloor Spreading concept (1960), stands as the unifying paradigm of modern geology and physical geography. The paradigm establishes that the Earth's cold, rigid outer shell—the Lithosphere (comprising the crust and outermost upper mantle, roughly 100 km thick)—is not a continuous, static eggshell. Instead, it is fragmented into a dynamic mosaic of approximately seven major and numerous minor lithospheric plates. These plates float upon and translate across the underlying ductile, semi-fluid Asthenosphere. Driven primarily by the thermal convection of the Earth's mantle, slab pull, and ridge push, lithospheric interactions along plate margins dictate the global spatial distribution of planetary volcanism, seismic activity, deep-sea trenches, and spectacular orogenic mountain belts."),

    ("Mechanisms of Plate Motion: Thermal Convection, Slab Pull, and Ridge Push",
     "The geodynamic propulsion of tectonic plates operates via coupled thermodynamic forces. Early models attributed plate motion exclusively to mantle convection currents dragging the lithosphere like a conveyor belt, as originally hypothesized by Arthur Holmes. Modern geophysical research, however, demonstrates that gravitational body forces within the plates themselves provide the primary kinetic drive. The foremost mechanism is 'Slab Pull': as an oceanic plate moves away from a spreading center, it cools, accumulates mantle lithosphere, and densifies until its density exceeds that of the underlying asthenosphere. Upon subduction at an oceanic trench, the cold, heavy sinking slab exerts a massive gravitational pull that drags the trailing plate seaward. This is complemented by 'Ridge Push', a gravitational sliding force generated by the elevated topography of mid-ocean ridges (standing 2 to 3 km above the abyssal plain), which drives oceanic lithosphere downslope toward subduction zones. Basal mantle convection drag and trench suction provide secondary kinematic inputs."),

    ("Convergent Boundaries: The Three Regimes of Plate Collision",
     "Convergent plate boundaries represent destructive margins where lithospheric slabs collide under intense compressional stress. The geomorphic and seismic outcome depends critically upon the crustal density and nature of the colliding plates, generating three distinct collisional regimes. First, Oceanic-Continental convergence occurs when a dense oceanic plate (e.g., the Nazca Plate) collides with a buoyant continental plate (e.g., South America). The denser oceanic slab subducts beneath the continent, carving a deep oceanic trench (such as the Peru-Chile Trench). As the slab plunges into the asthenosphere, dewatering of hydrous minerals lowers the melting point of the overlying mantle wedge (flux melting). This generates andesitic magma plumes that ascend to create a continental volcanic arc (the Andes Mountains). Subduction friction produces a Wadati-Benioff seismic zone, characterized by earthquake foci graduating from shallow near the trench to intermediate and deep-focus events down to 700 km."),

    ("Oceanic-Oceanic and Continental-Continental Collisional Orogenesis",
     "Second, Oceanic-Oceanic convergence occurs where two oceanic plates collide (e.g., the Pacific Plate subducting beneath the Philippine or Eurasian plates). The older, colder, and therefore denser of the two slabs subducts, plunging steeply into the mantle. This creates the planet's deepest abyssal depressions, such as the Mariana Trench (10,994 meters), and produces basaltic-to-andesitic volcanic island arcs (e.g., the Aleutian, Japanese, and Marianas archipelagos). Third, Continental-Continental convergence occurs when intervening oceanic basins close completely (the Wilson Cycle) and two buoyant continental landmasses collide. Because continental crust (Sial) is too low in density (2.7 g/cm³) to be subducted into the dense mantle (3.3 g/cm³), subduction halts. Instead, the collision zone undergoes colossal crustal shortening, intense folding, nappe formation, and thrust faulting. This process—collisional orogenesis—doubles the crustal thickness up to 70–80 km, thrusting up towering mountain massifs like the Himalayas, Tibetan Plateau, and Alps. Crucially, because no descending slab penetrates the deep melting zone, continental collision zones produce violent shallow-to-intermediate earthquakes but are entirely devoid of active volcanism."),

    ("Regional Seismotectonic Synthesis: Pakistan at the Triple Plate Junction",
     "The geodynamics of plate tectonics find profound empirical expression in South Asia. Pakistan occupies one of the most structurally complex and seismically hazardous tectonic zones in the world, positioned at the convergence zone where the Indian Plate indents northward into the Eurasian Plate at a rate of 4 to 5 centimeters per year, while the Arabian Plate subducts beneath the Makran coast in the southwest. This ongoing continent-continent collision, initiated approximately 50 million years ago, closed the ancient Tethys Sea and erected the Himalayan, Karakoram, and Hindu Kush mountain ranges. The collision front within Pakistan is marked by major thrust systems: the Main Karakoram Thrust (MKT), Main Mantle Thrust (MMT), and Main Boundary Thrust (MBT). Westward, the plate boundary transforms into an expansive transpressional strike-slip system—the 850-km-long Chaman Fault—which accommodates the differential motion between the Indian and Eurasian plates. This active tectonic setting generated the catastrophic 1935 Quetta earthquake (Mw 7.7), the 2005 Muzaffarabad/Kashmir earthquake (Mw 7.6, which ruptured the MBT), and the 2013 Awaran earthquake (Mw 7.7 along the Makran-Chaman transition)."),

    ("Conclusion: The Living Earth and Geomorphic Equilibrium",
     "In evaluative conclusion, Plate Tectonics provides the grand unified architectural blueprint for physical geography. It refutes early static theories of geosynclines and contracting Earth crusts, demonstrating that the planet's surface is in perpetual dynamic equilibrium. Endogenic forces driven by radioactive decay and slab-pull subduction continuously construct relief—raising mountain cordilleras, opening ocean basins, and venting magma—while exogenic weathering and fluvial processes strive to denude the landforms back to base level. Without plate tectonics, the Earth's geochemical cycles would cease, carbon sequestration through subduction would falter, and our planet would degrade into a flat, geologically inert sphere. The ongoing tremors along the northern arc of Pakistan remain a vivid reminder that the tectonic assembly of our planet remains an active, unfinished reality.")
]

p1_q2_quotes = [
    ("Scholar / Geologist", "Theoretical Contribution / Concept", "Exact Quotable Academic Thesis / Axiom"),
    ("Alfred Wegener (1912)", "The Origin of Continents and Oceans", "The continental landmasses, like icebergs of sial, have drifted through the simatic ocean floor over geological epochs, unified originally as Pangaea."),
    ("Harry H. Hess (1962)", "History of Ocean Basins (Seafloor Spreading)", "Mantle convection upwells at mid-ocean ridges to create new ocean crust, while old ocean floor is swallowed back into the mantle at deep-sea trenches."),
    ("J. Tuzo Wilson (1965)", "Nature (A New Class of Faults - Transform Faults)", "The earth's surface is composed of a mosaic of rigid mobile plates whose margins are bounded by ridges, subduction trenches, and transform faults."),
    ("Dan McKenzie & Robert Parker (1967)", "The North Pacific: An Example of Tectonics on a Sphere", "Plate movements can be mathematically modeled as rigid rotations on the surface of a sphere, governed by Euler's rotation poles and strict kinematic laws."),
    ("Xavier Le Pichon (1968)", "Sea-Floor Spreading and Continental Drift", "The quantitative division of the earth's lithosphere into six major rigid blocks provides a comprehensive explanation for all global seismic epicenters and mountain-building belts.")
]

# ==============================================================================
# QUESTION 3 (P1_Q3): ROCKS, SEDIMENTARY CLASSIFICATION & THE ROCK CYCLE
# Solved from: 2022 Q.2, 2018 Q.3, 2013 Q.3, 2008 Q.1
# ==============================================================================
p1_q3_meta = [
    ("Target Syllabus", "Section 2: Landforms — Rocks: Origin, Formation & Classification; Rock Cycle"),
    ("Exam Years Solved", "KPPSC PMS 2022 Q.2, 2018 Q.3, 2013 Q.3, 2008 Q.1"),
    ("Marks & Allocation", "20 Marks | Strict 35-Minute Descriptive Writing Limit"),
    ("Core Theoretical Thesis", "Rocks constitute solid mineral aggregates classified genetic-petrographically into Igneous, Sedimentary, and Metamorphic classes. Sedimentary rocks, representing 75% of continental surface exposures, evolve through weathering, transport, deposition, and diagenesis (compaction and cementation) into mechanically formed (clastic), chemically precipitated, and organically derived strata. The James Hutton Rock Cycle illustrates the continuous petrological transformation where endogenic tectonic heat and exogenic denudational processes cycle matter perpetually.")
]

p1_q3_callout = [
    "ROCK VS. MINERAL DISTINCTION: A mineral is a naturally occurring, inorganic crystalline solid with a definite chemical composition (e.g., Quartz SiO2, Calcite CaCO3). A rock is an aggregate of one or more minerals (e.g., Granite = Quartz + Feldspar + Mica; Limestone = Calcite).",
    "THE THREE GENETIC ROCK FAMILIES: (1) Igneous (Primary rocks formed by magma crystallization; intrusive plutonic vs. extrusive volcanic); (2) Sedimentary (Secondary rocks formed via lithification of eroded detritus or chemical/organic precipitation; characterized by stratification, fossils, and porosity); (3) Metamorphic (Tertiary rocks formed by solid-state recrystallization under heat, pressure, and chemically active fluids).",
    "DETAILED SEDIMENTARY CLASSIFICATION: Clastic/Mechanically formed (Conglomerate/Breccia > Sandstone > Siltstone > Shale based on Udden-Wentworth grain size scale); Chemically formed (Evaporites: Halite/Rock Salt, Gypsum; Carbonates: Travertine); Organically formed (Biogenic Carbonates: Chalk, Coral Limestone; Carbonaceous: Peat -> Lignite -> Bituminous Coal).",
    "HUTTONIAN ROCK CYCLE: Formulated by James Hutton (1785) under the principle of Uniformitarianism ('The present is the key to the past'), demonstrating that rocks are not static relics but dynamic stages in a closed terrestrial matter-energy cycle."
]

p1_q3_map = """
                         [PETROLOGICAL ROCK CYCLE]
                                     │
           ┌─────────────────────────┴─────────────────────────┐
           ▼                                                   ▼
   [IGNEOUS ROCKS]                                     [SEDIMENTARY ROCKS]
   (Primary Crustal Rocks)                             (Secondary Stratified Rocks)
   • Intrusive / Plutonic                              • Weathering -> Transport ->
     (Coarse crystalline: Granite,                       Deposition -> Diagenesis/Lithification
      Gabbro, Peridotite)                              • Covers 75% of land surface
   • Extrusive / Volcanic                                      │
     (Aphanitic / Glassy: Basalt,                              │
      Rhyolite, Obsidian)                                      ▼
                                                  [SEDIMENTARY CLASSIFICATION]
                                           ┌───────────────────┼───────────────────┐
                                           ▼                   ▼                   ▼
                                       [CLASTIC]          [CHEMICAL]          [ORGANIC]
                                     (Mechanical)        (Evaporites)        (Biogenic)
                                     • Conglomerate      • Rock Salt         • Limestone
                                     • Sandstone         • Gypsum            • Chalk
                                     • Shale / Mudstone  • Travertine        • Coal Series
                                           │
                                           ▼
                                 [METAMORPHIC ROCKS]
                                 (Solid-State Metamorphism)
                                 • Foliated: Slate -> Schist -> Gneiss
                                 • Non-foliated: Marble (from Limestone), Quartzite
"""

p1_q3_timeline = [
    ("Sedimentary Rock Class", "Sub-Types & Grain Size / Chemical Origin", "Diagnostic Lithological Examples & Environments"),
    ("Clastic (Mechanically Formed) — Rudaceous", "Coarse gravels (>2 mm); rounded pebbles (conglomerate) vs. angular fragments (breccia).", "Alluvial fans, high-energy mountain streams, scree slopes (e.g., Potohar Conglomerates, Salt Range Breccias)."),
    ("Clastic (Mechanically Formed) — Arenaceous", "Medium sand grains (0.0625 to 2 mm); dominated by resistant quartz minerals.", "Sandstones (quartz arenite, arkose, greywacke); desert sand seas, beach deposits, deltaic distributor channels (e.g., Khewra Sandstone)."),
    ("Clastic (Mechanically Formed) — Argillaceous", "Fine particles (<0.0625 mm); silt (0.004–0.0625 mm) and clay (<0.004 mm); highly fissile bedding.", "Shales, mudstones, claystones; low-energy environments: lake beds, floodplains, deep marine abyssal zones; primary source rock for petroleum."),
    ("Chemically Precipitated (Inorganic)", "Precipitation of dissolved mineral salts due to evaporation, supersaturation, or pH shifts.", "Evaporites: Rock Salt (Halite, NaCl) and Gypsum (CaSO4·2H2O) formed in barred basins; Travertine/Tufa in Karst limestone caves and hot springs."),
    ("Organically Formed (Biogenic Carbonate)", "Accumulation of skeletal remains, shells of marine organisms (calcite/aragonite CaCO3).", "Fossiliferous limestone, chalk (microscopic foraminifera), coquina, coral reef limestone (e.g., Margalla Limestone of Hazara/KPK)."),
    ("Organically Formed (Carbonaceous)", "Lithification of compressed plant/vegetal organic debris in anoxic swamp conditions.", "Coal series: Peat (50% carbon) -> Lignite (60-70%, Thar coal) -> Bituminous (80%, Dandot) -> Anthracite (>90% metamorphic).")
]

p1_q3_essay = [
    ("Introduction: Petrographic Foundations and the Earth's Crustal Architecture",
     "Rocks constitute the solid mineral aggregates that compose the Earth's crust and lithosphere, forming the elemental substrate of all physical landscapes. In geological science, a rock is distinguished from a mineral: whereas a mineral is a naturally occurring, inorganic crystalline solid possessing a definite chemical formula (e.g., quartz SiO2 or calcite CaCO3), a rock is an aggregate of one or more minerals bound together by chemical crystallization or physical cementation. Geoscientists classify rocks genetically into three great families based upon their mode of origin: Igneous rocks, formed through the cooling and solidification of molten magma or lava; Sedimentary rocks, produced through the denudation, deposition, and lithification of pre-existing materials; and Metamorphic rocks, formed through the solid-state recrystallization of existing rocks under extreme geothermal heat, directional pressure, and hydrothermal fluids. Although igneous and metamorphic rocks constitute 95 percent of the Earth's crust by volume, sedimentary rocks blanket over 75 percent of the continental surface, recording the environmental and biological history of the planet."),

    ("Genesis and Diagenesis of Sedimentary Rocks",
     "Sedimentary rocks originate through an uninterrupted sequence of exogenic geomorphic processes operating across four sequential phases: weathering, transportation, deposition, and diagenesis (lithification). Mechanical and chemical weathering disintegrates exposed bedrock into detrital fragments and dissolved ionic solutions. Surface agents—predominantly fluvial systems, but also glacial ice, wind, and oceanic currents—transport these sediments down hydraulic gradients into sedimentary basins (lakes, river plains, continental shelves, and deep marine trenches). Once transport energy wanes, sediments settle into horizontal strata under gravitational influence (the Law of Original Horizontality). The conversion of unconsolidated sediments into solid rock occurs during diagenesis. As overburden accumulates, 'compaction' expels interstitial pore fluids and packs mineral grains tightly. Simultaneously, 'cementation' precipitates mineral-rich fluids (calcite, silica, iron oxides) across pore spaces, binding mineral grains into cohesive sedimentary strata. Distinctive features of sedimentary rocks include stratification (bedding planes), cross-bedding, ripple marks, porosity, and the unique preservation of organic fossils."),

    ("Comprehensive Classification of Sedimentary Rocks",
     "Sedimentary rocks are categorized systematically into three principal genetic groups based on their petrological origin and composition: Mechanically Formed (Clastic), Chemically Precipitated, and Organically Formed (Biogenic) rocks. Clastic sedimentary rocks consist of weathered rock fragments (clasts) classified strictly according to grain diameter using the Udden-Wentworth scale. These subdivide into: Rudaceous rocks (coarse grains >2 mm), including rounded-clast Conglomerates deposited in high-energy braided rivers and angular-clast Breccias formed at the base of talus slopes; Arenaceous rocks (sand grains 0.0625–2 mm), dominated by durable quartz Sandstones found across desert ergs and beach shores; and Argillaceous rocks (fine silt and clay particles <0.0625 mm), forming fissile Shales and Mudstones in calm lacustrine and offshore marine environments. Shales constitute over 50 percent of all sedimentary rocks and serve as the paramount source and cap rocks in petroleum petroleum geology."),

    ("Chemical and Organic Sedimentary Strata: Carbonates and Evaporites",
     "Chemically formed sedimentary rocks arise from the direct inorganic precipitation of dissolved mineral salts from supersaturated aqueous solutions, driven primarily by arid evaporation. Foremost among these are the 'Evaporites', exemplified by Rock Salt (Halite, NaCl), Gypsum (CaSO4·2H2O), and Anhydrite, which crystallize in enclosed arid desert play lakes or restricted marine lagoons—such as the massive Neoproterozoic evaporite deposits of the Salt Range in Punjab/KPK. Travertine and calcareous tufa represent chemical carbonates precipitated around geothermal springs and within Karst caverns (stalactites/stalagmites). Conversely, Organically Formed (Biogenic) rocks derive from the accumulation of carbonaceous or calcareous organic debris. Biogenic carbonates include Fossiliferous Limestone, Coquina (cemented shell fragments), and Chalk (microscopic coccoliths and foraminifera shells). Carbonaceous rocks develop from the burial and biochemical coalification of dense plant matter in stagnant, anoxic swamps, progressing through the coal rank series: raw Peat (~50% carbon) to Lignite brown coal (~60–70% carbon, exemplified by the massive Thar coal deposits in Sindh), Bituminous coal (~80% carbon), and ultimately metamorphic Anthracite (>90% carbon)."),

    ("The Huttonian Rock Cycle: Closed-System Terrestrial Matter Cycling",
     "The relationship between igneous, sedimentary, and metamorphic rocks is synthesized through the petrological Rock Cycle, formulated in 1785 by Scottish geologist James Hutton under the doctrine of Uniformitarianism ('The present is the key to the past'). The rock cycle demonstrates that no rock type enjoys permanent existence; rather, Earth functions as a closed geochemical recycling system driven by two primary energy sources: exogenic solar radiation (powering climate, weathering, and erosion) and endogenic terrestrial radiogenic heat (powering plate tectonics and mantle melting). Primary igneous rocks (e.g., granitic plutons or basaltic plateaus) are elevated by tectonic orogeny and subjected to atmospheric weathering, yielding sediments that lithify into sedimentary rocks. As tectonic burial, subduction, or contact metamorphism exposes these sedimentary strata to elevated temperatures (>200°C) and lithostatic pressures, they recrystallize into metamorphic rocks (e.g., shale metamorphosing into slate, schist, and gneiss; limestone recrystallizing into marble). If temperatures surpass melting thresholds (~700°C–1,200°C), metamorphic rocks melt completely into magma, which upon cooling crystallizes into new igneous rock, completing the perpetual cycle."),

    ("Conclusion: Lithological Dominance in Geomorphic Evolution",
     "In evaluative summary, rocks and the geological rock cycle establish the fundamental lithological control over landform evolution. The differential resistance of rock types dictates topography: resistant sandstones and igneous granites form towering cliffs, cuestas, and mountain ridges, whereas soft, fissile shales and clays weather rapidly into broad alluvial strike valleys and badlands. In Pakistan's geomorphic landscape, the sedimentary sequence of the Salt Range, Potohar Plateau, and Sulaiman fold belts preserves a complete lithostratigraphic record spanning from the Precambrian salt marls to the Siwalik molasses. Understanding sedimentary classification and the dynamic rock cycle is indispensable not only for theoretical geomorphology, but also for economic resource extraction, including fossil fuel reservoirs, industrial mineral deposits, and civil engineering infrastructure.")
]

p1_q3_quotes = [
    ("Scholar / Geologist", "Theoretical Contribution / Concept", "Exact Quotable Academic Thesis / Axiom"),
    ("James Hutton (1785)", "Theory of the Earth & The Rock Cycle", "We find no vestige of a beginning, no prospect of an end; rocks decay to form strata, strata rise into mountains, in one unbroken cycle of terrestrial renovation."),
    ("Charles Lyell (1830)", "Principles of Geology (Uniformitarianism)", "The forces which alter rocks and shape the earth today—weathering, sediment deposition, volcanic intrusion—have operated with uniform energy across all geological time."),
    ("F.J. Pettijohn (1957)", "Sedimentary Rocks", "Sedimentary rocks are the definitive historical archives of the planet; their textures, mineral assemblages, and bedding planes record the hydraulic environments and life forms of vanished worlds."),
    ("Norman L. Bowen (1928)", "The Evolution of the Igneous Rocks", "Fractional crystallization of a cooling basaltic magma systematically dictates the sequence of silicate minerals formed, from olivine to quartz, determining crustal rock compositions."),
    ("Arthur N. Strahler (1969)", "Physical Geography", "The structural form of any landscape is a direct reflection of underlying lithology; rock resistance to exogenic denudation remains the primary determinant of landform relief.")
]

# Write builder executions for Batch 1 (Q1, Q2, Q3) and append Q4, Q5, Q6 definitions next
print("Batch 1 (Q1 to Q3) data loaded. Appending Q4 to Q6...")

# ==============================================================================
# QUESTION 4 (P1_Q4): WEATHERING PROCESSES & MASS WASTING MOVEMENTS
# Solved from: 2022 Q.1, 2018 Q.1, 2013 Q.4
# ==============================================================================
p1_q4_meta = [
    ("Target Syllabus", "Section 2: Landforms — Weathering (Physical & Chemical) & Mass Wasting"),
    ("Exam Years Solved", "KPPSC PMS 2022 Q.1, 2018 Q.1, 2013 Q.4"),
    ("Marks & Allocation", "20 Marks | Strict 35-Minute Descriptive Writing Limit"),
    ("Core Theoretical Thesis", "Weathering operates as the in-situ static disintegration (mechanical/physical) and decomposition (chemical) of bedrock, preparing material for geomorphic mobilization. Mass wasting constitutes the gravitationally driven, downslope movement of weathered regolith and rock without a primary transporting medium (fluid). Governed by Mohr-Coulomb shear strength criteria, mass movements classify across velocity and moisture gradients into creep, solifluction, slumping, debris flows, and catastrophic rockfalls, posing acute geomorphic hazards across Pakistan's Karakoram and Himalayan highway corridors.")
]

p1_q4_callout = [
    "WEATHERING VS. EROSION DISTINCTION: Weathering is an in-situ (static) breakdown of rocks in contact with the atmosphere, hydrosphere, and biosphere. Erosion is the dynamic mobilization and transportation of weathered debris by mobile agents (running water, glacial ice, wind, waves).",
    "PHYSICAL WEATHERING MECHANISMS: Operates without chemical change; dominant in arid and periglacial climates: (1) Frost Wedging (9% volumetric expansion of water freezing in joints); (2) Thermal Expansion/Insolation weathering (differential mineral expansion leading to granular disintegration); (3) Exfoliation/Unloading (sheeting in granites due to overburden removal); (4) Salt Wedging (haloclasty).",
    "CHEMICAL WEATHERING MECHANISMS: Alters mineral composition via aqueous reactions; dominant in warm, humid tropics: (1) Carbonation (atmospheric CO2 forming weak carbonic acid H2CO3, dissolving limestone CaCO3); (2) Hydrolysis (water reacting with feldspar to produce kaolinite clay); (3) Oxidation (ferrous Fe2+ rusting into ferric Fe3+ oxides/hematite); (4) Hydration.",
    "MASS WASTING CLASSIFICATION (SHARPE & VARNES): Governed by water content and velocity: Slow Flowage (soil creep, solifluction in permafrost); Rapid Flowage (earthflows, mudflows, debris avalanches); Slides (translational rockslides, rotational slumps with curved shear planes); Falls (free-fall rockfalls producing talus/scree cones)."
]

p1_q4_map = """
                     [EXOGENIC DENUDATION DYNAMICS]
                                   │
         ┌─────────────────────────┴─────────────────────────┐
         ▼                                                   ▼
 [WEATHERING: IN-SITU BREAKDOWN]                   [MASS WASTING: GRAVITY DRIVEN]
 (Prepares regolith; zero transport)               (Downslope translation; shear failure)
         │                                                   │
 ┌───────┴───────┐                                   ┌───────┴───────┐
 ▼               ▼                                   ▼               ▼
[PHYSICAL]      [CHEMICAL]                         [SLOW FLOW]     [RAPID / CATASTROPHIC]
• Frost Wedge   • Carbonation                      • Soil Creep    • Mudflow (lahars)
• Exfoliation   • Hydrolysis                         (terracetts)  • Debris Avalanche
• Thermal Shock • Oxidation                        • Solifluction  • Rotational Slump
• Salt Crystal  • Hydration                          (gelifluction)• Rockfall (talus cones)
                                                             │
                                               [GOVERNING SHEAR EQUATION]
                                               Safety Factor (FS) = Shear Strength / Shear Stress
                                               FS < 1.0 -> Mass failure occurs
"""

p1_q4_timeline = [
    ("Geomorphic Process / Class", "Physical / Chemical Mechanism", "Diagnostic Landforms, Field Expressions & Pakistan Examples"),
    ("Physical: Frost Wedging (Cryofracture)", "Water in bedrock joints expands by 9% upon freezing, exerting up to 2,100 kg/cm² pressure.", "Felsenmeer (block fields), angular jagged peaks, scree/talus slopes flanking high mountain roads (e.g., Karakoram Highway / Babusar Pass)."),
    ("Physical: Exfoliation / Unloading", "Erosion strips deep lithostatic overburden, causing plutonic rocks to expand upward along release joints.", "Exfoliation domes, sheeting joints, curved rock slabs peeling like onion layers (e.g., granitic domes of Nagar and Swat Valley)."),
    ("Chemical: Carbonation & Solution", "CO2 dissolves in rainwater: H2O + CO2 -> H2CO3; reacts with calcite: CaCO3 + H2CO3 -> Ca(HCO3)2 (soluble).", "Karst landscapes: sinkholes, solution grikes/clints, subterranean cavern systems (e.g., Margalla limestone caves, Salt Range)."),
    ("Chemical: Hydrolysis & Oxidation", "Hydrolysis: K-feldspar + H+ + H2O -> Kaolinite clay + K+ + Silica. Oxidation: 4FeSiO3 + O2 -> 2Fe2O3 + 4SiO2 (rusting).", "Saprolite deep weathering profiles, spheroidal weathering boulders, lateritic red soils in humid subtropical climates; structural weakening of bedrock."),
    ("Mass Wasting: Rotational Slump & Rockfall", "Shear stress exceeds shear strength; rotational shear failure along concave slip plane; free-fall of detached blocks.", "Crescentic head scarps, backward-tilted bench blocks, toe hummocks, extensive talus cones (e.g., Attabad Landslide 2010, Hunza, blocking river)."),
    ("Mass Wasting: Debris Flows & Mudflows", "Water-saturated regolith loses internal friction; behaves as high-density fluid moving down channels at 50 km/h.", "Alluvial debris fans, structural burial of settlements and bridges along active glaciated Himalayan/Karakoram valleys during monsoon deluges.")
]

p1_q4_essay = [
    ("Introduction: The Denudational Engine of the Earth's Surface",
     "The morphology of the continental lithosphere represents a continuous battlefield between endogenic forces, which construct relief through tectonic orogeny and volcanism, and exogenic processes, which seek to denude elevated topography back to base level. Denudation operates via a fourfold continuum: weathering, mass wasting, erosion, and deposition. Weathering stands as the foundational preparatory process—the in-situ (static) mechanical disintegration and chemical decomposition of exposed bedrock under atmospheric, hydrospheric, and biological conditions, without involving an active fluid transporting medium. Once rock cohesion is broken down into unconsolidated regolith, Mass Wasting mobilizes this material downslope under the direct pull of terrestrial gravity. Together, weathering and mass wasting govern slope evolution, soil genesis, and sediment supply for global drainage basins, while simultaneously constituting severe geomorphic hazards in high-relief orogenic belts like Pakistan's northern mountain arc."),

    ("Mechanisms of Physical (Mechanical) Weathering: Disintegration Without Mineral Alteration",
     "Physical weathering breaks cohesive bedrock into progressively smaller, angular fragments without altering its underlying chemical or mineralogical composition. It dominates in arid, hyper-arid, alpine, and polar environments where diurnal temperature fluctuations and freeze-thaw cycles are acute. The primary physical mechanism is 'Frost Wedging' (cryofracture). When liquid water enters micro-fractures and freezes at 0°C, it expands by approximately 9 percent in volume, exerting tensile pressures exceeding 2,100 kilograms per square centimeter—far surpassing the tensile strength of any rock. This produces jagged felsenmeer (block fields) and feeds extensive talus/scree cones along mountain flanks. In hyper-arid desert basins, 'Thermal Expansion' (insolation weathering) drives mechanical disintegration: because mineral grains (quartz, feldspar, biotite) possess unequal coefficients of thermal expansion, intense diurnal heating and nocturnal cooling generate internal stresses, producing granular disintegration and exfoliation. Furthermore, 'Pressure Release / Unloading' occurs when erosional stripping removes kilometers of overlying lithostatic weight, causing underlying plutonic granites to expand elastically and fracture along parallel sheeting joints, forming massive exfoliation domes."),

    ("Mechanisms of Chemical Weathering: Molecular Decomposition and Mineral Synthesis",
     "In contrast, chemical weathering involves the complex molecular alteration of primary rock minerals, transforming them into secondary minerals (such as clays and iron oxides) or soluble aqueous ions. It reaches maximum intensity in warm, humid equatorial and subtropical zones where water, atmospheric oxygen, and biological carbon dioxide are abundant. The dominant chemical process is 'Carbonation', wherein atmospheric or soil CO2 dissolves in meteoric rainwater to form weak carbonic acid (H2O + CO2 -> H2CO3). This acidic groundwater attacks calcium carbonate in limestones (CaCO3 + H2CO3 -> Ca(HCO3)2), converting insoluble rock into soluble calcium bicarbonate—the geochemical foundation of worldwide Karst topography. Second, 'Hydrolysis' represents a chemical reaction where water molecules dissociate into H+ and OH- ions, directly attacking silicate minerals. In granites, hydrolysis leaches potassium ions from orthoclase feldspar, transforming rigid crystals into soft, expansive kaolinite clay while releasing free silica. Third, 'Oxidation' occurs when dissolved oxygen reacts with iron-bearing mafic minerals (pyroxenes, amphiboles, olivine), producing reddish-yellow ferric oxides (hematite Fe2O3 and limonite) in a process analogous to rusting, fatally compromising rock structural integrity."),

    ("Mass Wasting Mechanics: Slope Stability and the Shear Strength Threshold",
     "Mass wasting (mass movement) encompasses all downslope movements of weathered bedrock and soil under the direct influence of gravity, distinct from fluvial or aeolian transport where a fluid medium acts as the carrying agent. Mechanically, slope stability is governed by the Coulomb-Terzaghi shear strength equation: tau = c + (sigma - u) * tan(phi), where tau is shear strength, c is cohesion, sigma is normal stress, u is pore water pressure, and phi is the internal angle of friction. A slope fails when the Factor of Safety (FS = Shear Strength / Shear Stress) falls below 1.0. Shear stress is increased by gravitational slope steepening (undercutting by rivers or road construction), tectonic seismic shaking, and heavy overburden. Conversely, shear strength is catastrophically degraded by water saturation: water lubricates slip planes, adds buoyant hydrostatic pore pressure that reduces effective normal stress (sigma - u), and adds immense gravitational weight to unconsolidated regolith."),

    ("Classification of Mass Movements: Creep, Slumping, Debris Flows, and Falls",
     "Following the geomorphic classifications of C.F.S. Sharpe and D.J. Varnes, mass wasting is categorized across two physical parameters: velocity (slow to catastrophic) and moisture content (dry to saturated). 'Slow Flowage' is exemplified by Soil Creep—the imperceptible, millimeter-scale annual downslope migration of regolith caused by freeze-thaw heaving, evidenced by tilted fence posts, bent tree trunks, and soil terracettes. In periglacial landscapes, 'Solifluction' (or gelifluction) occurs when the active soil layer thaws over impermeable permafrost, sliding slowly downslope as lobed sheets. 'Rapid Flowage' includes Mudflows and Debris Flows, where saturated, fine-grained material behaves as a viscous liquid, surging down mountain channels at speeds exceeding 50 km/h, often triggered by monsoon cloudbursts or glacial lake outbursts (GLOFs). 'Slides' involve shear displacement along distinct failure planes: Rotational Slumps exhibit a backward-rotated head scarp along a concave slip surface, whereas Translational Rockslides move rapidly down planar bedding or joint surfaces. Finally, 'Falls' entail the free fall of detached rock blocks down vertical cliffs, accumulating as massive scree slopes."),

    ("Conclusion: Geomorphic Hazards and Slope Dynamics in Pakistan",
     "In evaluative summation, weathering and mass wasting operate as coupled geomorphic drivers that continually sculpt the Earth's subaerial terrain. In mountainous territories like Khyber Pakhtunkhwa, Gilgit-Baltistan, and Azad Kashmir, these processes transcend academic geomorphology to present urgent socioeconomic and infrastructure hazards. The collision of the Indian and Eurasian plates maintains steep, unstable slopes, while monsoon rains and seismic tremors frequently trigger catastrophic mass failures. A tragic manifestation was the January 2010 Attabad Landslide in Hunza, where a colossal rock avalanche severed the Karakoram Highway, dammed the Hunza River, submerged multiple villages, and formed a 21-kilometer-long lake. Managing these relentless exogenic forces requires sophisticated terrain mapping, geotechnical slope stabilization (rock bolting, drainage adits, gabion retaining walls), and strict land-use zoning based on geomorphic principles.")
]

p1_q4_quotes = [
    ("Scholar / Geomorphologist", "Theoretical Contribution / Concept", "Exact Quotable Academic Thesis / Axiom"),
    ("C.F.S. Sharpe (1938)", "Landslides and Related Phenomena", "Mass wasting represents the universal gravitational adjustment of slopes, operating across a continuous spectrum from imperceptible soil creep to catastrophic rock avalanches."),
    ("Karl Terzaghi (1925)", "Theoretical Soil Mechanics & Effective Stress", "The catastrophic failure of slopes under saturated conditions is not caused by water acting as a simple lubricant, but by the development of neutral pore-water pressure which nullifies frictional resistance."),
    ("B.W. Sparks (1972)", "Geomorphology", "Weathering is the essential, indispensable precursor to all effective subaerial erosion; without chemical decay and mechanical shattering, rivers and glaciers would operate upon unyielding bedrock."),
    ("David J. Varnes (1978)", "Slope Movement Types and Processes", "Classification of slope failures must integrate the type of movement—fall, topple, slide, spread, or flow—with the nature of the material, whether solid bedrock or unconsolidated engineering soils."),
    ("Arthur N. Strahler (1952)", "Dynamic Basis of Geomorphology", "Slope forms and mass movement velocities reflect an exact equilibrium between the applied gravitational shear stresses and the resisting shearing strength of the rock-soil system.")
]

# ==============================================================================
# QUESTION 5 (P1_Q5): FLUVIAL GEOMORPHOLOGY & RIVER LANDFORMS
# Solved from: 2018 Q.2, 2008 Q.7
# ==============================================================================
p1_q5_meta = [
    ("Target Syllabus", "Section 2: Landforms — Fluvial Erosion & Deposition; River Valley Development"),
    ("Exam Years Solved", "KPPSC PMS 2018 Q.2, 2008 Q.7"),
    ("Marks & Allocation", "20 Marks | Strict 35-Minute Descriptive Writing Limit"),
    ("Core Theoretical Thesis", "Running water is the paramount subaerial geomorphic agent sculpting continental relief. Operating across a dynamic thermodynamic gradient from headwaters to ultimate base level, a river system evolves through three diagnostic stages—Youth, Maturity, and Old Age—governed by hydraulic geometry and sediment transport mechanics. Fluvial work transitions systematically from dominant vertical incisive erosion (gorges, waterfalls, potholes) in youthful mountain reaches to lateral planation and expansive depositional architecture (floodplains, meanders, oxbow lakes, natural levees, deltas) in senile lowland plains, as epitomized by the Indus River Basin.")
]

p1_q5_callout = [
    "PARAMOUNT EXOGENIC AGENT: Fluvial processes (running water) denude and transport more continental sediment to global oceans than glaciers, wind, and waves combined (delivering over 15–20 billion tons of sediment annually).",
    "THE THREE-FOLD FLUVIAL WORK: (1) Erosion (Hydraulic action, Corrasion/Abrasion, Attrition, Corrosion/Solution); (2) Transportation (Traction/rolling, Saltation/bouncing, Suspension of clays/silts, Solution of dissolved salts); (3) Deposition (occurs when stream velocity and kinetic energy fall below sediment-settling threshold).",
    "DAVISIAN THREE-STAGE PROFILE: Youth (steep gradient, high kinetic energy, dominant vertical downcutting, V-shaped valleys, gorges, rapids, waterfalls, absence of floodplains); Maturity (graded profile, lateral planation, open V-shaped valley, emergence of meander belt and narrow floodplain); Old Age (low gradient near ultimate base level, lateral shifting, expansive floodplains, oxbow lakes, braided networks, natural levees, deltas, and peneplain formation).",
    "INDUS BASIN CASE STUDY: The Indus River illustrates the complete fluvial spectrum: from torrential youthful vertical incision in the Indus Gorge (cutting 5,000 m deep through the Nanga Parbat syntaxis) to mature meanders in the Upper Indus Plain (Punjab) and senile deltaic bifurcation into 17 major distributaries entering the Arabian Sea."
]

p1_q5_map = """
                       [FLUVIAL GEOMORPHIC CONTINUUM]
                                     │
           ┌─────────────────────────┼─────────────────────────┐
           ▼                         ▼                         ▼
    [YOUTHFUL STAGE]          [MATURE STAGE]             [OLD AGE / SENILE]
    (Upper Mountain Course)   (Middle Valley Course)     (Lower Plain Course)
    • Dominant: Vertical      • Dominant: Lateral        • Dominant: Deposition
      Downcutting               Planation                  & Alluvial Building
    • Steep gradient          • Graded profile           • Low gradient / Base level
    • High kinetic energy     • Valley widening          • High sediment load
           │                         │                         │
    ┌──────┴──────┐           ┌──────┴──────┐           ┌──────┴──────┐
    ▼             ▼           ▼             ▼           ▼             ▼
[EROSIONAL]   [FEATURES]  [EROSIONAL]   [DEPOSITION] [EROSIONAL]   [DEPOSITION]
• V-Shaped    • Waterfalls • Open Valleys• Floodplain • Cut banks   • Oxbow Lakes
  Gorges      • Rapids    • River       • Initial    (limited)     • Natural Levees
• Potholes    • Interlock   Terraces      Meanders                 • Braided Channels
• River Canyons Spurs                                              • Alluvial Deltas
"""

p1_q5_timeline = [
    ("Fluvial Valley Stage", "Hydraulic Regime & Dominant Process", "Diagnostic Erosional & Depositional Landforms"),
    ("Youthful Stage (Upper Course)", "Steep hydraulic gradient; high velocity; dominant vertical downcutting (incision); hydraulic plucking & abrasion.", "Deep V-shaped valleys, vertical gorges and canyons (e.g., Indus Gorge at Gilgit/Chilas), waterfalls, cataracts, rapids, plunge pools, and potholes."),
    ("Potholes & Plunge Pools", "Cavitation and localized turbulent eddy currents drill abrasive boulders into bedrock depressions.", "Cylindrical bedrock hollows (potholes) in upper stream beds; massive plunge pools excavated at waterfall bases by falling hydraulic load."),
    ("Mature Stage (Middle Course)", "Moderate gradient; lateral erosion surpasses vertical incision; stream achieves graded equilibrium profile.", "Broad open valleys, bluffs, intermeshing meanders, slip-off slopes vs. undercut river cliffs, valley-floor floodplains, paired alluvial river terraces."),
    ("Old Age: Meanders & Oxbow Lakes", "Very low gradient; lateral corrasion shifts meander loops until high-water cut-off breaches narrow necks.", "Sinous meander loops, point bars (sand deposition on inner bank), abandoned oxbow lakes (mortlakes), and backswamp marshlands."),
    ("Old Age: Natural Levees & Braiding", "River overbank flooding deposits coarse sediment immediately along channel margins; heavy sediment load splits flow.", "Linear wedge-shaped natural levees flanking banks, braided river channels with mid-stream braid bars (e.g., Indus River near D.I. Khan and Mianwali)."),
    ("Old Age: Deltaic System (Terminal)", "Stream enters standing body of water (ocean/lake); velocity drops abruptly to zero; complete sediment settling.", "Alluvial deltas: Arcuate (Nile, Indus), Bird’s-foot (Mississippi), or Estuarine; extensive network of branching distributaries and mangrove swamps.")
]

p1_q5_essay = [
    ("Introduction: Running Water as the Master Architect of Continental Relief",
     "Among all the exogenic geomorphic agents operating across the subaerial Earth, running water is indisputably the most widespread, energetic, and morphologically significant. Fluvial systems—encompassing channelized streams and tributary networks—drain over 80 percent of the planetary landmass, annually transporting an estimated 15 to 20 billion metric tons of eroded bedrock debris and dissolved minerals to the world's oceans. A river is not merely a conduit of water; it operates as an integrated open physical system that transports mass and dissipates potential kinetic energy down a topographic gradient toward 'ultimate base level' (sea level, as defined by John Wesley Powell in 1875). From its high-altitude headwaters to its terminal marine mouth, a river system executes three inseparable geological tasks: erosion, transportation, and deposition. In doing so, it evolves systematically through three morphological stages—Youth, Maturity, and Old Age—creating an extraordinary repertoire of erosional and depositional landforms that structure continental geography."),

    ("Mechanisms of Fluvial Work: Erosion and Sediment Transportation Mechanics",
     "Fluvial erosion occurs through four interrelated hydraulic processes. 'Hydraulic Action' represents the direct mechanical tearing, lifting, and dragging of loose rock fragments by the sheer kinetic force of turbulent water, intensified by cavitation shockwaves in high-velocity rapids. 'Corrasion' (or Abrasion) is the scouring of the channel bed and banks using the rock debris carried by the river as natural geological chisels; it is the primary engine of bedrock channel downcutting. 'Attrition' is the continuous mechanical collision of boulders, pebbles, and sand grains against one another during transport, rapidly reducing angular blocks into rounded, polished gravels and fine silt. Finally, 'Corrosion' (Solution) represents the chemical dissolution of soluble rocks (such as chalk and limestone) by weak carbonic and humic acids present in river water. Sediment transportation proceeds through four physical modes: Traction (rolling and sliding of heavy bedload boulders), Saltation (bouncing of sand grains down-current), Suspension (clays and silts held aloft by hydraulic turbulence, constituting 70–80% of total load), and Solution (invisible dissolved mineral ions)."),

    ("The Youthful Stage: Incisive Downcutting and Torrential Mountain Landforms",
     "In its Youthful (Upper) Stage, a river cascades through high-elevation terrain characterized by steep gradients, turbulent flow, and high potential energy. Because the stream profile stands hundreds or thousands of meters above base level, vertical erosion (downcutting) violently outstrips lateral planation. The diagnostic erosional feature of this stage is the steep-sided, narrow 'V-shaped Valley', where rapid stream incision carves the valley floor while mass wasting and slope wash degrade the valley walls. Where vertical incision operates across massive, horizontally bedded or jointed resistant rock (such as granites or quartzites), the river excavates nearly vertical 'Gorges' and expansive 'Canyons'. The Indus River provides the ultimate global example: at the Nanga Parbat syntaxis in northern Pakistan, it has incised the world's deepest gorge, cutting over 5,000 meters deep between Nanga Parbat (8,126 m) and Rakaposhi. Structural anomalies along the youth profile create knickpoints, generating precipitous 'Waterfalls' and 'Rapids'. At the base of waterfalls, cavitation and impact excavate massive 'Plunge Pools', while circular eddy currents armed with quartz pebbles drill deep, cylindrical bedrock depressions known as 'Potholes' into channel floors."),

    ("The Mature Stage: Lateral Planation, Valley Widening, and River Terraces",
     "As the river descends into its Mature (Middle) Stage, its hydraulic gradient flattens and stream velocity moderates. At this juncture, vertical downcutting decelerates as the stream approaches graded equilibrium—a state described by J.H. Mackin (1948) as a 'graded stream' where slope and channel characteristics are precisely adjusted to transport the sediment load supplied by the basin. Lateral erosion now surpasses vertical incision: the river swings from side to side in wide arcs, undercutting the outer valley walls and planing away the spurs of the youth stage to produce a broad, flat valley floor. The diagnostic erosional and depositional feature of this stage is the emergence of the 'Meander Belt' and the initial development of an alluvial 'Floodplain'. Where tectonic uplift occurs or climate cycles cause sea level to drop, the river is rejuvenated, incising into its old floodplain to abandon elevated gravel benches known as 'Paired River Terraces', beautifully preserved along the Soan River and Indus Valley."),

    ("The Old Age (Senile) Stage: Meander Dynamics, Oxbow Lakes, and Levees",
     "In its Old Age (Lower) Stage, the river traverses broad alluvial plains near ultimate base level. The hydraulic gradient becomes negligible (often less than a few centimeters per kilometer), and the river's energy is overwhelmingly dedicated to transporting its massive sediment load. Lateral shifting reaches its zenith. Flow velocity is asymmetric across channel cross-sections: maximum velocity occurs along the outer curved bank (the 'cut bank'), driving intense lateral corrasion and bank slumping, while minimum velocity on the inner bank forces sediment deposition, constructing gentle 'Point Bars' (slip-off slopes). Over time, meander loops become increasingly exaggerated. During high-discharge seasonal floods, the river takes the shortest, steepest hydraulic path, breaching the narrow meander neck to carve a straight 'chute cut-off'. The abandoned, crescentic meander loop is sealed off by silt deposition, creating an 'Oxbow Lake' (locally termed 'dhand' in Sindh). Furthermore, during periodic overbank floods, coarse sand and silt settle immediately along channel margins, constructing elevated, linear ridges known as 'Natural Levees' that contain normal flow above the surrounding backswamps."),

    ("Deltaic Architecture and the Indus Geomorphic Synthesis",
     "The ultimate terminus of the fluvial cycle is the 'Delta', a depositional landform formed where a river enters a standing body of water (an ocean, sea, or lake). As flow enters the static basin, hydraulic velocity plummets abruptly to zero, forcing the entire remaining suspended sediment load to settle according to Stokes' Law, creating bottomset, foreset, and topset beds (G.K. Gilbert's model). The main channel becomes clogged with its own sediment, forcing the river to bifurcate into an intricate network of distributaries. The Indus Delta, covering over 41,000 square kilometers along the Arabian Sea coast of Sindh, represents an arcuate, wave-dominated delta historically fed by 400 million tons of annual silt. In evaluative summation, the fluvial system represents an exquisite, self-regulating physical cycle. From the incisive mountain torrents of Khyber Pakhtunkhwa to the sprawling meanders and deltas of Punjab and Sindh, fluvial geomorphology dictates the agricultural prosperity, settlement distribution, and physical geography of the entire Pakistan subcontinent.")
]

p1_q5_quotes = [
    ("Scholar / Geomorphologist", "Theoretical Contribution / Concept", "Exact Quotable Academic Thesis / Axiom"),
    ("William Morris Davis (1899)", "The Geographical Cycle of Landforms", "Landscape is a function of structure, process, and stage; rivers evolve systematically from torrential youth to graded maturity and peneplain old age."),
    ("G.K. Gilbert (1877)", "Report on the Geology of the Henry Mountains", "Channel equilibrium is governed by the law of uniform slope; river velocity, discharge, and sediment load maintain a dynamic adjustment across all reach geometries."),
    ("John Wesley Powell (1875)", "Exploration of the Colorado River (Base Level)", "Base level represents the grand limit of subaerial erosion below which the dry land cannot be denuded by the waters of the ocean."),
    ("J. Hoover Mackin (1948)", "Concept of the Graded River", "A graded stream is one in which, over a period of years, slope is delicately adjusted to provide, with available discharge, just the velocity required for the transportation of the load supplied from the drainage basin."),
    ("Luna B. Leopold & Thomas Maddock (1953)", "The Hydraulic Geometry of Stream Channels", "Width, depth, velocity, and suspended sediment load increase systematically as mathematical power functions of river discharge down-basin.")
]

# ==============================================================================
# QUESTION 6 (P1_Q6): GLACIAL GEOMORPHOLOGY & VALLEY GLACIERS VS. ICE SHEETS
# Solved from: 2022 Q.6, 2018 Q.8, 2010 Q.3
# ==============================================================================
p1_q6_meta = [
    ("Target Syllabus", "Section 2: Landforms — Glacial Landforms: Valley Glaciers vs. Ice Sheets"),
    ("Exam Years Solved", "KPPSC PMS 2022 Q.6, 2018 Q.8, 2010 Q.3"),
    ("Marks & Allocation", "20 Marks | Strict 35-Minute Descriptive Writing Limit"),
    ("Core Theoretical Thesis", "Glacial ice operates as a colossal plastic-solid geomorphic engine driven by cryospheric compaction, basal sliding, and internal deformation. Glaciers divide fundamentally into topographically confined Valley (Alpine) Glaciers and unconfined continental Ice Sheets. Operating via plucking and glacial abrasion, alpine glaciers transform pre-existing fluvial V-shaped valleys into steep-walled, flat-bottomed U-shaped glacial troughs flanked by hanging valleys, cirques, and arêtes, while depositional mechanics generate unsorted moraines, drumlins, and outwash plains—typified globally by the Karakoram giants (Siachen, Baltoro, Batura) of northern Pakistan.")
]

p1_q6_callout = [
    "CRYOSPHERIC ICE DYNAMICS: Snow transforms into glacial ice through compaction and recrystallization: Fresh Snow (density 0.1 g/cm³) -> Granular Firn/Névé (0.4–0.8 g/cm³) -> True Glacial Ice (>0.85 g/cm³). Glacial ice flows under gravity once ice thickness exceeds ~30–50 meters via internal plastic deformation (creep) and basal sliding.",
    "VALLEY GLACIERS VS. ICE SHEETS DICHOTOMY: Valley/Alpine Glaciers are linear ice tongues confined by pre-existing mountain valley walls (e.g., Baltoro, Siachen, Mer de Glace); Continental Ice Sheets are unconfined, dome-shaped continental ice masses submerging topography under kilometers of ice (e.g., Antarctica: 14M km², Greenland: 1.7M km²).",
    "EROSIONAL MASTER FEATURES: (1) Cirque/Corrie (horseshoe amphitheater hollow excavated by freeze-thaw plucking and rotational slip); (2) Arête (knife-edge ridge separating two cirques); (3) Horn Peak (pyramidal peak formed by 3+ intersecting cirques, e.g., Matterhorn, K2); (4) U-Shaped Trough (steep-sided, flat-floored glacial valley); (5) Hanging Valleys (tributary valleys left suspended above over-deepened main trough, forming waterfalls); (6) Fjords (submerged glacial troughs).",
    "DEPOSITIONAL ARCHITECTURE: Glacial Till / Moraines (unsorted, unstratified direct glacial debris): Lateral, Medial, Ground, and Terminal/Recessional Moraines; Fluvioglacial stratified deposits: Drumlins (streamlined teardrop till mounds), Eskers (sinuous subglacial meltwater river gravel ridges), Kames, and Outwash Plains (sandurs)."
]

p1_q6_map = """
                       [GLACIAL GEOMORPHIC SYSTEMS]
                                     │
           ┌─────────────────────────┴─────────────────────────┐
           ▼                                                   ▼
   [VALLEY / ALPINE GLACIERS]                          [CONTINENTAL ICE SHEETS]
   • Confined by mountain topography                   • Unconfined, massive ice domes
   • Linear, high relief                               • Submerges entire landforms
   • E.g., Siachen (76 km), Baltoro (63 km)            • E.g., Antarctica, Greenland
                                     │
           ┌─────────────────────────┴─────────────────────────┐
           ▼                                                   ▼
   [EROSIONAL LANDFORMS]                               [DEPOSITIONAL LANDFORMS]
   (Plucking & Basal Abrasion)                         (Glacial Till & Fluvioglacial)
   • Cirques / Corries (tarns)                         • Moraine Systems:
   • Arêtes & Horn Peaks (pyramidal)                     - Lateral & Medial Moraines
   • U-Shaped Glacial Trough                             - Terminal / End Moraines
   • Hanging Valleys (waterfalls)                        - Ground Moraine (till sheet)
   • Truncated Spurs & Roche Moutonnée                 • Drumlins (inverted spoon shape)
   • Fjords (drowned glacial valleys)                  • Eskers & Sandur Outwash Plains
"""

p1_q6_timeline = [
    ("Glacial Landform Class", "Geomorphic Genesis & Mechanism", "Diagnostic Diagnostic Examples & Pakistan Himalayan Context"),
    ("Valley Glaciers vs. Continental Ice Sheets", "Valley glaciers are constrained by mountain topography; continental sheets are unconfined, radial ice domes.", "Valley: Siachen (76 km), Biafo (67 km), Baltoro (63 km), Batura (57 km) in Karakoram. Ice Sheet: Antarctic Ice Sheet (>4,000 m thick)."),
    ("Cirque (Corrie / Cwm)", "Rotational slip of ice combined with frost plucking at bergschrund headwall; carves armchair hollow.", "Deep amphitheater-shaped rock basin; when ice melts, occupied by a circular alpine tarn lake (e.g., Saif-ul-Malook tarn in Kaghan Valley)."),
    ("Arêtes and Horn Peaks", "Backward recession of opposing cirque headwalls sharpens intervening mountain ridges and summits.", "Knife-edge saw-tooth arêtes; iconic pyramidal Horn peaks (e.g., Matterhorn in Alps; K2, Broad Peak in Karakoram Range)."),
    ("U-Shaped Glacial Trough", "Glacier widens, deepens, and straightens pre-existing fluvial V-shaped valley via plucking and abrasion.", "Steep, vertical valley walls, broad flat valley floor, truncated spurs, ribbon lakes (e.g., Swat, Hunza, and Gilgit alpine valleys)."),
    ("Hanging Valleys", "Main valley glacier incises far deeper than smaller tributary glaciers due to greater ice volume and erosive mass.", "Tributary valleys suspended high above main valley floor; tributaries cascade down as spectacular waterfalls (e.g., Hunza Valley tributaries)."),
    ("Moraine Architecture (Glacial Till)", "Direct, unstratified, unsorted dump of angular rock debris transported supraglacially and subglacially.", "Lateral moraines along valley flanks, medial moraines where two ice tongues fuse, and terminal end moraines damming glacial valleys.")
]

p1_q6_essay = [
    ("Introduction: The Cryosphere as a Geomorphic Force",
     "Glaciers constitute one of the most dynamic and powerful physical agents of subaerial denudation on Earth. Today, glacial ice blankets approximately 10 percent of the planet's subaerial land surface—predominantly locked within the colossal polar ice sheets of Antarctica and Greenland—while during the Pleistocene Epoch (the Quaternary Ice Age), glaciers expanded across more than 30 percent of continental landmasses. Outside the polar regions, the greatest concentration of glacial ice resides in the towering Karakoram, Himalayan, and Hindu Kush mountain ranges of northern Pakistan, rightly designated as Earth's 'Third Pole'. Glaciers originate wherever annual snowfall exceeds summer melting (ablation), causing fresh snow (density 0.1 g/cm³) to compact under successive layers into granular firn (or névé, 0.4–0.8 g/cm³) and ultimately recrystallize into dense, impermeable glacial ice (>0.85 g/cm³). Once ice thickness reaches a threshold of 30 to 50 meters, the shear stress generated by gravity surpasses the yield strength of ice, causing the glacier to deform internally like a hyper-viscous plastic solid and surge downslope."),

    ("Valley Glaciers vs. Continental Ice Sheets: A Morphological Dichotomy",
     "In physical geography, glaciers are categorized into two fundamental structural classes: Valley (Alpine) Glaciers and Continental Ice Sheets. Valley glaciers are linear tongues of ice that originate in high-altitude mountain snowfields (cirques) and flow down pre-existing stream valleys, strictly confined by surrounding bedrock topography. Their movement is highly channelized, their gradients are steep, and their morphological work is focused intensely upon valley deepening and relief accentuation. Pakistan hosts the world's longest alpine glaciers outside the polar circles, including the Siachen Glacier (76 km), the Biafo Glacier (67 km), the Baltoro Glacier (63 km), and the Batura Glacier (57 km). In stark contrast, Continental Ice Sheets (and Ice Caps) are unconfined, gargantuan domes of ice that submerge entire continental landmasses under ice sheets 2 to 4 kilometers thick. Their flow is radial, moving outward from elevated central domes toward marine margins regardless of underlying bedrock topography, as observed in the 14-million-square-kilometer Antarctic Ice Sheet. Whereas valley glaciers sharpen and steepen alpine peaks into jagged relief, continental ice sheets act as regional planes, grinding down continental shields into undulating, low-relief roche moutonnée landscapes."),

    ("Mechanics of Glacial Erosion: Plucking, Abrasion, and Subglacial Cavitation",
     "Glacial erosion operates primarily through two complementary mechanical processes: Plucking (quarrying) and Abrasion. 'Plucking' occurs along the fractured, jointed bedrock of the glacier's bed. Subglacial meltwater percolates into bedrock fissures, where pressure variations cause it to freeze, shattering rock blocks through frost wedging. As the overlying glacier advances, these shattered bedrock fragments are frozen into the basal ice and physically plucked from the bed, incorporated into the moving glacier as subglacial tools. 'Abrasion' is the subsequent rasping, scouring, and grooving of the bedrock bed by these embedded, angular rock fragments, functioning like coarse industrial sandpaper. Abrasion polishes bedrock surfaces to a mirror-like sheen, carves parallel linear 'glacial striations' that reveal past ice flow directions, and pulverizes rock into microscopic 'rock flour', which imparts a distinctive milky turquoise tint to downstream glacial meltwater streams."),

    ("Erosional Alpine Landforms: Cirques, Arêtes, Horns, and U-Shaped Troughs",
     "Valley glaciation completely restructures mountain topography. The foundational erosional feature is the 'Cirque' (or Corrie in Scotland, Cwm in Wales)—a steep-walled, armchair-shaped amphitheater hollow carved into the mountain headwall. Rotational slip of the ice mass combined with freeze-thaw plucking along the deep headwall crevasse (the bergschrund) steepens the backwall while scooping out a rock basin floor. When the glacier melts, this over-deepened basin frequently fills with water to form an enchanting circular alpine lake known as a 'Tarn' (such as Lake Saif-ul-Malook in Pakistan's Kaghan Valley). When two adjacent cirques erode backward toward each other, their intersecting headwalls sharpen the intervening ridge into a razor-thin, knife-edge serrated ridge termed an 'Arête'. Where three or more cirques erode radially into a single mountain summit, they quarry away all flanks to leave an isolated, towering pyramidal peak known as a 'Horn'—immortalized globally by the Swiss Matterhorn and regionally by the monumental pyramids of K2 and Broad Peak in the Karakoram. As the valley glacier bulldozes down-valley, it truncates pre-existing river spurs, widening, deepening, and straightening the valley into a monumental 'U-Shaped Glacial Trough' characterized by sheer vertical valley walls and a flat, broad floor. Furthermore, because tributary glaciers contain far less ice volume than the trunk glacier, they incise their beds far less deeply; when the ice vanishes, tributary valleys are left suspended high above the main trough floor as 'Hanging Valleys', from which meltwater streams plunge as spectacular waterfalls."),

    ("Depositional Architecture: Moraines, Drumlins, Eskers, and Outwash Plains",
     "Glacial deposition produces two distinct genetic suites of sediments: unstratified glacial till (moraine) dropped directly by melting ice, and stratified fluvioglacial drift reworked and sorted by meltwater streams. The primary depositional landform is the 'Moraine'—an unsorted, unstratified ridge of angular boulders, gravel, sand, and clay. 'Lateral Moraines' accumulate along the margins of valley glaciers from frost-shattered valley rockfall; where two tributary glaciers merge, their adjacent lateral moraines coalesce to form a prominent dark ribbon of debris running down the center of the trunk glacier, known as a 'Medial Moraine' (vividly visible along the Baltoro Glacier). At the maximum snout advance where ablation equals ice supply, debris is dumped to form a massive crescentic 'Terminal End Moraine', which frequently impounds meltwater to create proglacial lakes. Fluvioglacial landforms include 'Drumlins'—streamlined, teardrop-shaped hills of compacted till shaped like inverted spoons, with their steep blunt ends facing the direction of up-glacier advance. Subglacial meltwater streams running in pressurized sub-ice tunnels deposit sinuous, gravel ridges known as 'Eskers', while braided meltwater streams emerging from the glacier snout deposit expansive, sorted sand and gravel plains termed 'Outwash Plains' (or Sandurs)."),

    ("Conclusion: The Glacial Engine and Pakistan's Hydrological Security",
     "In evaluative summation, glacial geomorphology illustrates the supreme denudational capacity of the solid cryosphere. Glaciers act as natural bulldozers, transforming rugged fluvial terrains into dramatic, high-energy glacial landscapes that define the alpine world. Beyond their geomorphic splendor, glaciers serve as the irreplaceable freshwater towers of humanity. In Pakistan, the extensive glaciated massifs of the Karakoram, Himalayas, and Hindu Kush hold over 7,000 individual glaciers containing thousands of cubic kilometers of frozen freshwater. These glaciers feed the Indus River Basin, sustaining over 80 percent of Pakistan's agricultural irrigation, hydroelectric generation, and urban water supply. Today, anthropogenic climate change is accelerating glacial retreat, increasing the frequency of Glacial Lake Outburst Floods (GLOFs) across Chitral, Gilgit, and Hunza, and threatening the long-term hydrological stability of South Asia. Understanding glacial processes is thus not merely an academic exercise, but a strategic imperatve for Pakistan's national survival.")
]

p1_q6_quotes = [
    ("Scholar / Glaciologist", "Theoretical Contribution / Concept", "Exact Quotable Academic Thesis / Axiom"),
    ("Louis Agassiz (1840)", "Études sur les Glaciers (The Glacial Theory)", "The polished bedrock, striations, and erratic boulders across northern continents are not the work of a biblical deluge, but the undeniable signature of colossal vanished ice sheets."),
    ("J.T. Andrews (1975)", "Glacial Systems: An Approach to Glaciers and Their Landforms", "A glacier is a thermodynamic engine operating in open mass-balance equilibrium; its morphology is a direct response to inputs of snow accumulation and outputs of ablation."),
    ("John F. Nye (1952)", "The Mechanics of Glacier Flow", "Glacial ice behaves as a non-linear viscoplastic material; when shear stress at the bed exceeds a critical yield threshold, continuous plastic flow and basal sliding occur."),
    ("Clifford Embleton & C.A.M. King (1975)", "Glacial Geomorphology", "The conversion of a fluvial V-shaped valley into a glacial U-shaped trough represents one of the most complete geomorphic transformations in physical geography."),
    ("Kenneth Hewitt (2005)", "The Karakoram Anomaly & Glacial Dynamics", "The Karakoram cryosphere exhibits exceptional behavior, where high-altitude surging glaciers and debris-covered ablation zones maintain a delicate, resilient equilibrium amidst global warming trends.")
]

# ==============================================================================
# BATCH 1 EXECUTION RUNNER
# ==============================================================================
BATCH_1_TASKS = [
    ("Q1_Internal_Structure_of_Earth_and_Seismic_Evidence", "Q.1: Internal Structure of Earth & Seismic Discontinuities", p1_q1_meta, p1_q1_callout, p1_q1_map, p1_q1_timeline, p1_q1_essay, p1_q1_quotes),
    ("Q2_Plate_Tectonics_Theory_and_Orogenesis", "Q.2: Plate Tectonics Theory, Plate Collisions & Orogenesis", p1_q2_meta, p1_q2_callout, p1_q2_map, p1_q2_timeline, p1_q2_essay, p1_q2_quotes),
    ("Q3_Rocks_Classification_and_Rock_Cycle", "Q.3: Rocks, Sedimentary Classification & The Rock Cycle", p1_q3_meta, p1_q3_callout, p1_q3_map, p1_q3_timeline, p1_q3_essay, p1_q3_quotes),
    ("Q4_Weathering_Processes_and_Mass_Wasting", "Q.4: Weathering Processes (Physical & Chemical) & Mass Wasting", p1_q4_meta, p1_q4_callout, p1_q4_map, p1_q4_timeline, p1_q4_essay, p1_q4_quotes),
    ("Q5_Fluvial_Geomorphology_and_River_Landforms", "Q.5: Fluvial Geomorphology & River Valley Evolution", p1_q5_meta, p1_q5_callout, p1_q5_map, p1_q5_timeline, p1_q5_essay, p1_q5_quotes),
    ("Q6_Glacial_Geomorphology_Valley_Glaciers_vs_Ice_Sheets", "Q.6: Glacial Geomorphology: Valley Glaciers vs. Ice Sheets", p1_q6_meta, p1_q6_callout, p1_q6_map, p1_q6_timeline, p1_q6_essay, p1_q6_quotes),
]

for filename, title_text, meta, callout, map_txt, timeline, essay, quotes in BATCH_1_TASKS:
    docx_file = os.path.join(P1_DIR, f"{filename}_Master_Notes.docx")
    pdf_file = os.path.join(P1_DIR, f"{filename}_Master_Notes.pdf")
    md_file = os.path.join(P1_DIR, f"{filename}_Master_Notes.md")
    
    generate_geo_master_doc(meta, callout, map_txt, timeline, essay, quotes, docx_file, title_text)
    generate_pdf_and_md(meta, callout, map_txt, timeline, essay, quotes, pdf_file, md_file, title_text)

print("Batch 1 (Q1 through Q6) for Geography Paper 1 successfully generated!")
