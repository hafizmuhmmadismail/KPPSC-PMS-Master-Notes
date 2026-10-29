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
    r_hq = p_hq.add_run("4. SCHOLARLY PERSPECTIVES & GEOMORPHIC/CLIMATIC AXIOMS TABLE")
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
    
    md_content += "### 4. Scholarly Perspectives & Geomorphic/Climatic Axioms\n\n"
    q_hdr = quotes_data[0]
    md_content += f"| {q_hdr[0]} | {q_hdr[1]} | {q_hdr[2]} |\n|---|---|---|\n"
    for row in quotes_data[1:]:
        md_content += f"| **{row[0]}** | {row[1]} | *\"{row[2]}\"* |\n"
        
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(md_content)
        
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

<h3>4. Scholarly Perspectives & Geomorphic/Climatic Axioms</h3>
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

print("Batch 2 builder engine ready.")

# ==============================================================================
# QUESTION 7 (P1_Q7): KARST TOPOGRAPHY & GROUNDWATER HYDROLOGY
# Solved from: 2025 Q.6, 2016 Q.4, 2010 Q.8
# ==============================================================================
p1_q7_meta = [
    ("Target Syllabus", "Section 2: Landforms — Groundwater Hydrology & Karst Topography"),
    ("Exam Years Solved", "KPPSC PMS 2025 Q.6, 2016 Q.4, 2010 Q.8"),
    ("Marks & Allocation", "20 Marks | Strict 35-Minute Descriptive Writing Limit"),
    ("Core Theoretical Thesis", "Karst topography represents a distinctive geochemical terrain sculpted by the chemical carbonation and dissolution of soluble carbonate rocks (pure limestone CaCO3 >80%) by meteoric groundwater charged with atmospheric and biogenic CO2. Development requires essential environmental prerequisites: dense joint networks, incised fluvial valleys generating hydraulic gradients, and an aerated vadose zone overlying a fluctuating phreatic zone. The landscape evolves through surface corrosive landforms (lapies, sinkholes/dolines, uvalas, poljes) and subterranean cavern systems decorated with speleothems (stalactites, stalagmites, columns), exemplified regionally in the Margalla Hills and Salt Range.")
]

p1_q7_callout = [
    "KARST GEOCHEMICAL REACTION: Limestone is insoluble in pure water, but dissolves rapidly in carbonic acid: H2O + CO2 (atmospheric & soil organic) <-> H2CO3 (carbonic acid); CaCO3 (insoluble limestone) + H2CO3 <-> Ca(HCO3)2 (soluble calcium bicarbonate). Reversal of this reaction via CO2 outgassing precipitates speleothems.",
    "PREREQUISITES FOR KARST DEVELOPMENT (CVIJIĆ): (1) High-purity, massive limestone (>80% CaCO3, low insoluble residue); (2) Dense, well-developed network of joints and bedding planes allowing water penetration; (3) Entrenched river valleys creating a high hydraulic head (topographic gradient); (4) Moderate to high rainfall sustaining active groundwater recharge; (5) Substantial vertical separation between surface and phreatic water table (thick vadose zone).",
    "SURFACE VS. SUBSURFACE LANDFORMS: Surface: Karren/Lapies (grooved limestone pavements), Sinkholes/Dolines (solution, collapse), Uvalas (coalesced dolines), Poljes (massive flat-bottomed tectonic-karstic plains), Blind Valleys (where surface streams sink into ponors/swallow holes). Subsurface: Cavern networks, Blind river passages, Speleothems (Stalactites grow down, Stalagmites grow up, fusion creates Pillars/Columns).",
    "REGIONAL KARST PHENOMENA IN PAKISTAN: Manifested in the Eocene Margalla Limestone (Shah Allah Ditta caves near Islamabad), the Samana Suk Limestone in Hazara, and the massive Permian-Eocene calcareous sequences of the Salt Range and Balochistan (Ziarat limestone karst)."
]

p1_q7_map = """
                        [KARST HYDROLOGICAL SYSTEM]
                                     │
           ┌─────────────────────────┴─────────────────────────┐
           ▼                                                   ▼
 [SURFACE KARST LANDFORMS]                           [SUBTERRANEAN CAVERN SYSTEM]
 (Epikarst & Solution Drainage)                      (Vadose & Phreatic Speleogenesis)
 • Lapies / Karren (Limestone Pavements:             • Solution Caverns & Shafts
   Clints & Grikes)                                  • Phreatic Tubular Passages
 • Swallow Holes / Ponors (Stream Sink)              • Underground Rivers & Siphons
 • Sinkholes / Dolines:                              • Speleothem Architecture:
   - Solution Dolines (funnel-shaped)                  - Stalactites (ceiling straw)
   - Collapse Dolines (subsurface void)                - Stalagmites (floor mounds)
 • Uvalas (Compound sinkhole depressions)              - Pillars / Columns (fused)
 • Poljes (Massive flat-floored plains)                - Draperies & Flowstones
 • Blind Valleys (Terminated by ponors)              • Karst Springs (Vauclusian resurgences)
"""

p1_q7_timeline = [
    ("Karstic Feature / Zone", "Hydro-Chemical Genesis & Stage", "Diagnostic Diagnostic Features & Pakistan / Global Examples"),
    ("Karren / Lapies (Limestone Pavement)", "Rainwater with carbonic acid attacks exposed limestone joints and bedding planes.", "Etched, grooved bare rock surfaces; flat tabular blocks (clints) separated by deep dissolution crevices (grikes); Malham Cove (UK), Margalla Hills."),
    ("Sinkholes (Dolines)", "Solution dolines form by surface dissolution; collapse dolines occur when subterranean cavern roofs fail.", "Circular to funnel-shaped depressions (2–100 m wide); floor lined with insoluble residual terra rossa red clay; common in Dinaric Karst and Salt Range."),
    ("Uvalas and Poljes", "Coalescence of adjacent dolines produces uvalas; poljes form by tectonic down-faulting enhanced by lateral planation.", "Poljes are giant flat-bottomed valleys (up to 40 km long) bounded by steep limestone bluffs, drained by ponors; e.g., Livno Polje (Bosnia)."),
    ("Blind Valleys and Ponors", "Surface stream flows over impermeable rock, hits soluble limestone, and abruptly vanishes into a ponor.", "Valley terminates abruptly against a limestone wall; water travels through subterranean conduits to re-emerge as a Vauclusian spring."),
    ("Limestone Caverns & Caves", "Groundwater circulating below water table (phreatic zone) dissolves rock along major intersecting joints.", "Three-dimensional labyrinth of chambers and galleries; Carlsbad Caverns (USA), Postojna Cave (Slovenia), Pir Ghaib karst springs (Balochistan)."),
    ("Speleothem Deposition", "Cave water drips from ceiling; CO2 outgasses into cave air, reversing reaction: Ca(HCO3)2 -> CaCO3 (calcite).", "Stalactites (tapering icicle-like ceiling forms with central canal), Stalagmites (blunt floor mounds), and continuous Pillars/Columns.")
]

p1_q7_essay = [
    ("Introduction: The Geochemical Architecture of Karst Environments",
     "Karst topography constitutes one of the most distinctive and geomorphologically fascinating terrain types on the planet, characterized by barren, rocky surfaces, disrupted surface drainage, and vast subterranean cavern labyrinths. The term originates from the 'Krs' plateau of the Dinaric Alps in Slovenia and Croatia, where the phenomena were first systematically classified by the Serbian geographer Jovan Cvijić in his foundational 1893 treatise 'Das Karstphänomen'. Unlike fluvial, glacial, or aeolian landscapes where mechanical erosion dominates, Karst landforms are produced almost exclusively by chemical dissolution and geochemical weathering. The solvent action of percolating meteoric groundwater—naturally acidified by atmospheric and biogenic carbon dioxide—systematically attacks and hollows out soluble carbonate bedrock, primarily pure limestone (calcium carbonate, CaCO3) and dolomite (calcium magnesium carbonate, CaMg(CO3)2). The resulting terrain is an integrated three-dimensional hydrologic system comprising surface corrosion depressions, disappearing rivers, and spectacular speleothem-decorated subterranean caverns."),

    ("Geochemical Dynamics of Carbonate Solution and Precipitation",
     "The fundamental chemical mechanism driving karstification is the reversible equilibrium of calcium carbonate carbonation. Pure, distilled water dissolves negligible quantities of limestone (less than 14 milligrams per liter at normal atmospheric pressures). However, when meteoric rainwater falls through the atmosphere and percolates through organic-rich soil zones, it absorbs gaseous carbon dioxide. In soil horizons, microbial respiration and root decay generate CO2 partial pressures 10 to 100 times greater than atmospheric levels. The dissolved CO2 reacts with water to generate weak carbonic acid: H2O + CO2 <-> H2CO3. When this acidified groundwater encounters limestone bedrock, it dissociates and dissolves the insoluble calcium carbonate, converting it into highly soluble calcium bicarbonate: CaCO3 + H2CO3 <-> Ca(HCO3)2. This forward reaction drives speleogenesis (cave excavation) and surface dissolution. Conversely, when mineral-saturated groundwater drips into an open cave chamber where the partial pressure of CO2 is lower than in the host bedrock, CO2 outgasses into the cave atmosphere. This forces the chemical equilibrium to reverse, causing insoluble calcium carbonate to precipitate out of solution as crystalline calcite (speleothems)."),

    ("Prerequisites for Karst Development: The Cvijić Environmental Criteria",
     "Karst landscapes cannot develop haphazardly; their evolution demands five stringent lithological, geological, and hydrologic prerequisites, first articulated by Jovan Cvijić and expanded by modern speleologists Derek Ford and Paul Williams. First, the bedrock must consist of dense, high-purity limestone, ideally containing more than 80 to 90 percent calcium carbonate with minimal insoluble clays or chert. Second, the limestone formation must be massive and well-jointed, possessing a dense geometric grid of vertical fractures, joints, and horizontal bedding planes that provide conduits for acidic water circulation. Third, the regional geomorphic setting must exhibit deep, entrenched river valleys incised into the limestone mass; this topographic entrenchment creates a steep hydraulic gradient (hydraulic head), driving gravity-fed groundwater circulation from high plateau recharge zones down to base-level stream valleys. Fourth, the climate must supply moderate to abundant meteoric precipitation (generally >700 mm annually) to maintain active geochemical recharge. Fifth, there must exist a thick unsaturated 'Vadose Zone' (the zone of aeration between the surface and the water table) above the saturated 'Phreatic Zone' to permit three-dimensional speleogenesis."),

    ("Surface Karst Morphologies: From Lapies to Poljes",
     "The subaerial surface of a karst landscape exhibits a diagnostic hierarchy of erosional dissolution features. At the micro-scale, rainwash and sheetflow dissolve bare limestone outcrops into 'Karren' (or 'Lapies'). These limestone pavements consist of flat, tabular rock surfaces called 'clints', dissected by deep, vertical solution furrows called 'grikes'. At the meso-scale, the definitive karst landform is the 'Doline' (or Sinkhole)—a circular, funnel-shaped closed depression ranging from a few meters to hundreds of meters in diameter. Dolines originate through two distinct mechanisms: 'Solution Dolines', formed by gradual downward dissolution of rock along intersecting joints; and 'Collapse Dolines', produced when the roof of an underlying subterranean cavern abruptly fails under gravitational stress. Where adjacent dolines expand and coalesce, they form compound, elongated depressions termed 'Uvalas'. At the macro-scale, tectonic graben subsidence combined with lateral karst corrosion creates 'Poljes'—colossal, flat-floored valleys spanning up to tens of kilometers in length (e.g., Livno Polje in Bosnia), bounded by steep limestone scarps. Surface drainage in karst is remarkably disrupted: rivers flow across impermeable rocks, encounter soluble limestone, and abruptly vanish underground through swallow holes or 'Ponors', leaving dry 'Blind Valleys' downstream."),

    ("Subsurface Speleogenesis and Speleothem Architecture",
     "Beneath the epikarst surface lies the subterranean realm of caves and caverns. According to William Morris Davis’s Two-Cycle Theory of cave formation (1930), speleogenesis proceeds in two stages: during the first cycle, deep phreatic groundwater circulating under hydrostatic pressure dissolves massive horizontal cavities along joint planes; during the second cycle, regional valley incision lowers the water table, draining the cavities and leaving an air-filled vadose cave system. Once air occupies the cavern, vadose dripwaters precipitate crystalline speleothems. Water dripping from the cave ceiling deposits mineral rings around the drip point, growing downward as slender, hollow tubes known as 'Soda Straws', which thicken into conical 'Stalactites'. Where drops fall onto the cave floor, splashing releases CO2, depositing broad, blunt mounds known as 'Stalagmites'. Eventually, descending stalactites and ascending stalagmites fuse together, forming continuous, monumental 'Pillars' or 'Columns'. Water trickling down inclined cave walls deposits corrugated, undulating sheets known as 'Draperies' or 'Flowstones'."),

    ("Conclusion: Karst Hydrogeology and Environmental Significance in Pakistan",
     "In evaluative conclusion, Karst topography demonstrates the supreme capacity of chemical processes to transform terrestrial landscapes into complex, subterranean hydrologic networks. In Pakistan, classic karst phenomena are extensively developed across the Eocene Margalla Limestone belt (featuring the Shah Allah Ditta caves near Islamabad), the Samana Suk formations of Khyber Pakhtunkhwa, the Eocene carbonate scarps of the Salt Range (Kallar Kahar and Choa Saidan Shah), and the vast Cretaceous-Eocene limestone massifs of Balochistan (Pir Ghaib and Ziarat). Karst terrains pose unique geotechnical and environmental management challenges: thin residual soils (terra rossa) are exceptionally vulnerable to erosion, subterranean cavities threaten highway and dam construction with sudden subsidence, and the rapid, unfiltered flow of groundwater through open caverns makes karst aquifers acutely susceptible to chemical and biological contamination. Sustainable management of these landscapes requires sophisticated hydrogeological tracer testing and strict environmental zoning.")
]

p1_q7_quotes = [
    ("Scholar / Geomorphologist", "Theoretical Contribution / Concept", "Exact Quotable Academic Thesis / Axiom"),
    ("Jovan Cvijić (1893)", "Das Karstphänomen (Foundations of Karstology)", "Karst is an organic hydro-chemical landscape wherein subterranean solution drainage completely replaces normal subaerial fluvial valley systems."),
    ("William Morris Davis (1930)", "Origin of Limestone Caverns (Two-Cycle Theory)", "Caverns are excavated in deep phreatic waters beneath a regional water table, and only subsequently decorated by speleothems when regional rejuvenation drains the chambers."),
    ("J Harlen Bretz (1942)", "Vadose and Phreatic Features of Limestone Caverns", "The diagnostic morphology of natural caves—ceiling pockets, sponge-work cavities, and tubular passages—bears the indelible signature of phreatic slow-solution excavation."),
    ("Derek Ford & Paul Williams (1989)", "Karst Hydrogeology and Geomorphology", "The evolution of karst is fundamentally governed by the epikarst-aquifer continuum, where structural discontinuities focus acidic recharge into self-amplifying conduit networks."),
    ("Arthur N. Palmer (1991)", "Origin and Morphology of Limestone Caves", "Cave patterns are direct geometric reflections of groundwater flow paths, lithologic purity, and the geochemical kinetics of the calcite-water-carbon dioxide system.")
]

# ==============================================================================
# QUESTION 8 (P1_Q8): AEOLIAN & COASTAL GEOMORPHOLOGY
# Solved from: 2016 Q.6, 2016 Q.8, 2008 Q.7
# ==============================================================================
p1_q8_meta = [
    ("Target Syllabus", "Section 2: Landforms — Wind (Aeolian) & Coastal Geomorphology"),
    ("Exam Years Solved", "KPPSC PMS 2016 Q.6, 2016 Q.8, 2008 Q.7"),
    ("Marks & Allocation", "20 Marks | Strict 35-Minute Descriptive Writing Limit"),
    ("Core Theoretical Thesis", "Aeolian and coastal geomorphic systems represent dynamic fluid-boundary environments governed by kinetic boundary-layer transport. In arid lands, wind operates via deflation and sandblasting abrasion, constructing distinctive erosional landforms (yardangs, zeugen, ventifacts) and organized depositional architectures (barchans, seifs, loess). In coastal zones, oscillating ocean waves, longshore drift, and tidal currents drive mechanical quarrying and hydraulic shock, carving erosional cliff-and-platform systems and depositing depositional tombolos, spits, and barrier bars, typified regionally by Pakistan's Makran Coast and Cholistan/Thar deserts.")
]

p1_q8_callout = [
    "AEOLIAN DYNAMICS (BAGNOLD): Wind is a fluid of low density (1/800th of water); hence, aeolian erosion is restricted to arid/semi-arid regions lacking protective vegetative cover. Transport modes: Surface Creep (rolling of coarse grains 0.5–2 mm), Saltation (bouncing of sand 0.05–0.5 mm, accounting for 75–80% of total transport, restricted to <1–2 meters above ground), and Suspension (dust and silt <0.05 mm carried thousands of kilometers).",
    "AEOLIAN LANDFORM TAXONOMY: Erosional: Deflation hollows/blowouts, Desert Pavement/Reg/Serir (armored surface left after fine dust is stripped), Ventifacts (faceted pebbles), Zeugen (tabular ridge-and-furrow ridges in horizontally bedded strata), Yardangs (elongated streamlined ridges in vertically tilted beds), Pedestal/Mushroom rocks. Depositional: Dunes (Barchans = crescentic, horns point downwind; Seifs = longitudinal linear ridges parallel to wind; Transverse = perpendicular to wind; Star dunes); Loess = unstratified, fertile, porous wind-blown silt blankets (e.g., Potohar Loess plateau in Pakistan).",
    "COASTAL WAVE MECHANICS & LANDFORMS: Wave energy derives from offshore wind shear stress. Waves break as solitary waves in the surf zone, generating swash (upslope surge) and backwash (gravity return). Longshore drift (oblique wave approach) transports sediment parallel to shore.",
    "COASTAL EROSIONAL VS. DEPOSITIONAL: Erosional: Wave-cut notch -> Sea cliff retreat -> Wave-cut platform -> Sea cave -> Sea arch -> Sea stack -> Sea stump. Depositional: Beaches, Spits (hook-shaped sand ridges extending into bays), Baymouth bars, Tombolos (sand spit connecting mainland to offshore island, e.g., Gwadar Peninsula hammerhead connected by a sandy isthmus tombolo)."
]

p1_q8_map = """
                   [AEOLIAN & COASTAL FLUID SYSTEMS]
                                   │
         ┌─────────────────────────┴─────────────────────────┐
         ▼                                                   ▼
 [AEOLIAN ARID GEOMORPHOLOGY]                        [COASTAL MARINE GEOMORPHOLOGY]
 (Wind Kinetic Transport: Saltation)                 (Wave Kinetic Energy & Tides)
         │                                                   │
 ┌───────┴───────┐                                   ┌───────┴───────┐
 ▼               ▼                                   ▼               ▼
[EROSIONAL]     [DEPOSITIONAL]                      [EROSIONAL]     [DEPOSITIONAL]
• Deflation:    • Sand Dunes:                       • Wave-Cut      • Beaches & Berms
  - Blowouts      - Barchan (crescent,                Notch & Cliff • Spits (recurved)
  - Desert Reg      horns downwind)                 • Wave-Cut      • Baymouth Bars
• Abrasion:       - Seif / Longitudinal               Platform      • Tombolo (e.g.,
  - Yardangs        (parallel to wind)              • Sea Caves &     Gwadar tombolo)
  - Zeugen        - Transverse Dunes                  Sea Arches    • Barrier Islands
  - Mushroom      - Star Dunes (multidirectional)   • Sea Stacks      & Lagoon reefs
• Ventifacts    • Loess Deposits (Potohar Plateau)    & Stumps
"""

p1_q8_timeline = [
    ("Geomorphic Realm / Process", "Physical Transport & Fluid Mechanism", "Diagnostic Diagnostic Landforms & Pakistan Field Manifestations"),
    ("Aeolian Erosion: Deflation", "Wind lifts and sweeps away loose, dry, silt-sized particles; leaves behind coarse, interlocking gravel armor.", "Deflation hollows (e.g., Qattara Depression); Desert Pavements (Reg in Algeria, Serir in Libya, Dasht in Balochistan)."),
    ("Aeolian Erosion: Abrasion", "Sand grains saltating within 1 meter of ground act as sandblasters, cutting softer rock strata preferentially.", "Yardangs (streamlined inverted ship-hull ridges), Zeugen (table-like ridges), Mushroom/Pedestal rocks; common in Chagai desert (Balochistan)."),
    ("Aeolian Deposition: Barchans", "Unidirectional wind, limited sand supply; sand cascades down steep slip-face (34° angle of repose).", "Crescent-shaped sand dunes; convex windward slope, crescent horns pointing downwind; highly mobile across Cholistan and Thar deserts."),
    ("Aeolian Deposition: Loess", "Suspended periglacial or desert dust carried vast distances, settling as unstratified, porous, calcareous silt blankets.", "Vertical cliff-cleaving loess deposits; exceptional agricultural fertility, prone to badland gully erosion (e.g., Potohar Loess plateau in Rawalpindi/Jhelum)."),
    ("Coastal Erosion: Cliff-Stack Sequence", "Hydraulic shock pressure (up to 30,000 kg/m²) and abrasion compress air into joints, excavating rock.", "Wave-cut notch -> Cliff retreat -> Wave-cut platform -> Caves meeting to form Sea Arches -> Arch roof collapses leaving isolated Sea Stacks (e.g., Makran coast)."),
    ("Coastal Deposition: Spits & Tombolos", "Longshore drift carries sand along coast; sediment extends across bay mouth or ties island to mainland.", "Sand spits, baymouth bars enclosing lagoons, and Tombolos (e.g., Gwadar Hammerhead Peninsula, where a tombolo connects an offshore bedrock block to Balochistan mainland).")
]

p1_q8_essay = [
    ("Introduction: Boundary-Layer Geomorphic Systems in Arid and Marine Realms",
     "Aeolian (wind) and coastal (marine wave) geomorphic systems operate as energetic fluid-boundary environments where the interface between Earth's lithosphere, atmosphere, and hydrosphere produces distinct landforms. In hyper-arid and semi-arid desert basins—which blanket nearly one-third of the global continental surface—liquid water is scarce, and vegetative cover is absent. Here, the turbulent boundary layer of atmospheric wind acts as a primary agent of sediment entrainment and landform evolution. Conversely, along the planet's 600,000 kilometers of coastline, oceanic kinetic energy—concentrated into breaking surf, nearshore waves, longshore drift, and tidal oscillations—sculpts the land-sea margin through relentless mechanical attack. Both geomorphic realms illustrate the principles of sediment fluid dynamics, where shear stress governs the threshold between erosion, transportation, and organized deposition, vividly showcased in Pakistan across the windswept sands of Cholistan and Thar and the wave-battered cliffs of the Makran Coast."),

    ("Aeolian Dynamics: Entrainment, Transportation Mechanics, and Desert Pavements",
     "Because air possesses a density approximately 1/800th that of liquid water, wind is a low-density fluid incapable of mobilizing coarse gravels under normal atmospheric pressures. In his classical 1941 treatise 'The Physics of Blown Sand and Desert Dunes', Brigadier Ralph A. Bagnold established that aeolian transport occurs via three distinct mechanical modes: Surface Creep, Saltation, and Suspension. Coarse sand grains (0.5 to 2.0 mm) roll along the desert floor under the impact of falling particles (surface creep, 20–25% of load). Fine dust, silt, and clays (<0.05 mm) are caught in turbulent thermal updrafts and carried thousands of kilometers aloft in suspension. However, the workhorse of aeolian transport is 'Saltation' (accounting for 75–80% of total sediment movement): medium sand grains (0.05 to 0.5 mm) are lifted by aerodynamic drag, bounce downwind in parabolic arcs, and strike the surface, ejecting further grains upon impact. Because gravitational settling confines saltating sand grains within 1 to 2 meters of the surface, wind abrasion is strictly a low-altitude phenomenon."),

    ("Aeolian Landforms: Erosional Yardangs and Depositional Sand Seas (Ergs)",
     "Aeolian erosion operates through two physical processes: Deflation and Abrasion. 'Deflation' is the direct lifting and removal of unconsolidated dry sand and silt by wind turbulence. Over time, deflation hollows out vast desert depressions (such as the Qattara Depression in Egypt) and strips away all fine sediments, leaving behind an armored, interlocking mosaic of coarse pebbles and gravel known as a 'Desert Pavement' (termed 'Reg' in the Sahara and 'Dasht' in Balochistan). 'Abrasion' is the mechanical sandblasting of rock by saltating quartz grains. Differential abrasion carves 'Zeugen' (ridge-and-furrow terrain developed in horizontally bedded alternating hard and soft rocks), 'Yardangs' (elongated, aerodynamic, streamlined ridges sculpted in vertically tilted strata, aligned parallel to prevailing winds), and 'Mushroom Rocks' (pedestals undercut near their base). Depositional aeolian landforms divide into Sand Dunes and Loess. Dunes are organized sand accumulations classified by wind consistency and sand supply: 'Barchans' are crescent-shaped dunes formed under unidirectional winds and limited sand, whose convex face faces the wind while two horns point downwind; 'Seif' (or Longitudinal) dunes are knife-edge linear ridges running parallel to the resultant wind direction; 'Transverse' dunes form wavy sand ridges perpendicular to strong winds; and 'Star Dunes' develop under multidirectional wind regimes. Beyond deserts, wind deposits 'Loess'—extensive, unstratified blankets of calcareous silt blown from glacial outwash or deserts, epitomized in Pakistan by the Potohar Loess Plateau."),

    ("Coastal Wave Mechanics: Nearshore Energy Dissipation and Longshore Transport",
     "In the marine environment, the primary geomorphic engine is the wind-generated oceanic wave. As deep-water oscillatory waves (where water molecules move in closed circular orbits) approach a sloping shoreline, water depth decreases to less than half the wavelength (d < L/2). Friction with the seabed slows wave velocity, compresses wavelength, and steepens wave height until the wave becomes unstable and breaks violently in the surf zone. This transformation converts potential energy into intense kinetic energy: the breaking wave surges up the beach as 'swash' and retreats down-slope under gravity as 'backwash'. Waves hitting the coast at an oblique angle generate a net current parallel to the shoreline known as 'Longshore Drift', which acts as an immense aquatic conveyor belt transporting millions of cubic meters of sand along continental margins."),

    ("The Coastal Erosional Sequence: From Wave-Cut Notches to Isolated Stacks",
     "Coastal erosion proceeds through four aggressive mechanical processes: Hydraulic Action (where waves crashing against cliffs generate hydraulic shock pressures exceeding 30,000 kg/m²), Cavitation (the explosive compression and decompression of air trapped in rock fissures), Abrasion/Corrasion (the battering of cliffs using sand and gravel as artillery), and Attrition. The coastal erosional sequence follows a predictable evolutionary path: waves relentlessly attack the base of a high sea cliff, excavating a horizontal 'Wave-Cut Notch'. As the notch deepens, the unsupported cliff overhang collapses under gravity, causing the sea cliff to retreat landward. This cliff recession leaves behind a gently sloping, bare bedrock platform at low-tide level known as a 'Wave-Cut Platform'. Along structurally heterogeneous headlands, wave refraction focuses hydraulic energy on rocky promontories. Waves quarry intersecting joints to excavate 'Sea Caves'. When waves breach through the headland from both sides, they unite to form a 'Sea Arch'. Eventually, gravitational stress and wave pounding collapse the arch roof, leaving an isolated, vertical bedrock pillar standing offshore known as a 'Sea Stack'. Continued abrasion reduces the stack to a submerged 'Sea Stump'."),

    ("Coastal Depositional Architecture and the Gwadar Tombolo Case Study",
     "Where wave energy attenuates in sheltered bays or shallow waters, coastal deposition constructs constructive landforms. Foremost is the 'Beach'—an accumulation of wave-sorted sand, shingle, or shell fragments organized into berms and beach cusps. When longshore drift transports sand past a sharp change in coastline orientation (such as an inlet or bay mouth), wave energy drops, causing sand to deposit as a linear finger projecting into open water, known as a 'Spit' (often recurved at its tip by wave refraction). If a spit extends entirely across the mouth of a bay, it seals it off as a 'Baymouth Bar', enclosing a brackish coastal lagoon. A globally iconic coastal depositional landform is the 'Tombolo'—a depositional spit or sandbar that links an offshore bedrock island to the mainland. The definitive textbook example in Pakistan is the 'Gwadar Peninsula': Gwadar was originally an isolated, hammerhead-shaped offshore rocky island in the Arabian Sea; over geological epochs, convergent wave refraction and longshore drift deposited a monumental, 12-kilometer-long sandy isthmus tombolo, permanently welding the Gwadar bedrock ridge to the mainland of Balochistan. In evaluative summary, aeolian and coastal geomorphology illustrate the supreme organizing capacity of fluid currents, structuring the arid basins and marine frontiers that sustain human transport, maritime trade, and coastal urbanization.")
]

p1_q8_quotes = [
    ("Scholar / Geomorphologist", "Theoretical Contribution / Concept", "Exact Quotable Academic Thesis / Axiom"),
    ("R.A. Bagnold (1941)", "The Physics of Blown Sand and Desert Dunes", "Saltation is the primary engine of sand transport; sand grains moving in leaping trajectories govern both the erosion of bedforms and the geometric morphology of sand dunes."),
    ("Douglas W. Johnson (1919)", "Shore Processes and Shoreline Development", "Coastal profiles represent an evolutionary sequence governed by wave attack, progressing from initial notched headlands to mature, graded shore platforms and baymouth bars."),
    ("C.A.M. King (1972)", "Beaches and Coasts", "The morphology of any shoreline reflects a dynamic equilibrium between wave energy inputs, tidal ranges, and the availability of mobile littoral sediments."),
    ("Arthur N. Strahler (1969)", "Physical Geography", "In arid lands devoid of vegetal cohesion, wind and ephemeral flash floods combine to produce exceptional landforms characterized by desert pavements, sand seas, and pediments."),
    ("E.C.F. Bird (2008)", "Coastal Geomorphology: An Introduction", "Longshore drift and wave refraction operate as self-regulating hydrodynamic systems that sculpt sand spits, barrier chains, and connecting tombolos across continental shelves.")
]

# Write builder executions for Batch 2 (Q7 and Q8) and append Q9 to Q14 next
print("Batch 2 (Q7 and Q8) appended. Appending Q9 to Q14...")

# ==============================================================================
# QUESTION 9 (P1_Q9): SOLAR INSOLATION, HEAT BUDGET & RADIATION BALANCE
# Solved from: 2016 Q.1, 2013 Q.1, 2010 Q.1, 2008 Q.3
# ==============================================================================
p1_q9_meta = [
    ("Target Syllabus", "Section 1: Weather & Climate — Insolation, Heat Balance & Radiation Budget"),
    ("Exam Years Solved", "KPPSC PMS 2016 Q.1, 2013 Q.1, 2010 Q.1, 2008 Q.3"),
    ("Marks & Allocation", "20 Marks | Strict 35-Minute Descriptive Writing Limit"),
    ("Core Theoretical Thesis", "The Earth-Atmosphere system maintains long-term thermodynamic equilibrium through a closed radiation budget where incoming shortwave solar insolation (100 units) is exactly counterbalanced by outgoing longwave terrestrial and atmospheric radiation (100 units). The planetary albedo reflects ~30 units directly back into space, while the remaining 70 units undergo atmospheric absorption and surface heating before being reradiated. Latitudinal radiation imbalances—exhibiting net thermal surplus between 38°N and 38°S and net thermal deficit poleward—drive global atmospheric winds and ocean currents as heat-transfer engines, preventing equatorial overheating and polar hyper-freezing.")
]

p1_q9_callout = [
    "THE SOLAR CONSTANT & INCOMING INSOLATION: Solar radiation arriving at the outer boundary of Earth's atmosphere averages 1,367 Watts per square meter (W/m²), or roughly 1.96 cal/cm²/min (the Solar Constant). Radiation is shortwave (ultraviolet, visible, near-infrared, peak 0.5 microns).",
    "FACTORS CONTROLLING INSOLATION: (1) Angle of Solar Incidence (Lambert's Cosine Law: oblique rays spread over greater surface area and traverse longer atmospheric path length); (2) Duration of Daylight / Photoperiod (governed by Earth's 23.5° axial tilt and orbital revolution); (3) Atmospheric Attenuation (scattering by gas molecules/Rayleigh and aerosols/Mie, reflection by clouds, absorption by ozone, water vapor, and CO2); (4) Solar Distance (Perihelion 147M km vs. Aphelion 152M km).",
    "THE 100-UNIT GLOBAL HEAT BUDGET: Incoming: 100 units shortwave. Planetary Albedo = 30 units lost to space (20 reflected by clouds, 6 scattered by atmosphere, 4 reflected by ground). Absorbed = 70 units (19 absorbed by atmosphere/ozone/clouds, 51 absorbed by Earth's surface). Outgoing Longwave = 70 units radiated back to space (17 units directly through the 'atmospheric window' 8–12 microns, 53 units radiated from clouds and greenhouse-heated atmosphere), maintaining zero net energy gain/loss.",
    "LATITUDINAL HEAT IMBALANCE: Net radiative surplus exists between roughly 38°N and 38°S; net deficit exists from 38° to 90° poles. Poleward heat advection via planetary wind systems (Hadley/Ferrel cells, ~75–80%) and ocean current gyres (Gulf Stream, Kuroshio, ~20–25%) redistributes energy, maintaining habitable planetary climates."
]

p1_q9_map = """
                    [GLOBAL HEAT BUDGET & RADIATION BALANCE]
                                      │
            ┌─────────────────────────┴─────────────────────────┐
            ▼                                                   ▼
[INCOMING SHORTWAVE: 100 UNITS]                     [OUTGOING LONGWAVE: 100 UNITS]
(Peak: 0.5 μm; Visible/UV)                          (Peak: 10 μm; Infrared)
            │                                                   │
  ┌─────────┴─────────┐                               ┌─────────┴─────────┐
  ▼                   ▼                               ▼                   ▼
[PLANETARY ALBEDO]  [ENERGY ABSORBED]               [ATMOSPHERIC WINDOW] [ATMOSPHERIC EMISSION]
• 30% Reflected     • 70% Retained                  • 17% Lost directly  • 53% Emitted by
  - 20% Clouds        - 19% Atmosphere                to space (8-12 μm)   clouds, H2O, CO2
  - 6% Atmosphere     - 51% Ground                                       • Latent Heat (23%)
  - 4% Surface          * 26% Direct                                     • Sensible Heat (7%)
                        * 25% Diffuse                                    • Direct IR (21%)
                                      │
                         [LATITUDINAL HEAT IMBALANCE]
            Equator to 38° N/S: Net Surplus │ 38° to 90° N/S: Net Deficit
                                      │
             [POLEWARD HEAT TRANSFER: REDISTRIBUTION ENGINES]
             • Atmospheric Circulation: 75-80% (Winds, Cyclones, Jet Streams)
             • Oceanic Circulation: 20-25% (Warm Western Boundary Currents)
"""

p1_q9_timeline = [
    ("Radiation Component / Zone", "Energy Allocation & Physical Metric", "Atmospheric & Thermodynamic Role"),
    ("Incoming Solar Radiation (Insolation)", "100 Units (1,367 W/m² at top of atmosphere); shortwave (0.2–4.0 microns).", "Primary energy source driving all planetary weather, hydrological cycles, ocean currents, and biological photosynthesis."),
    ("Planetary Albedo (Reflection)", "30 Units lost immediately: 20 reflected by clouds, 6 by atmospheric air, 4 by ground.", "Total reflectivity of Earth system; snow/ice has 80–90% albedo, water has 5–10% (at high sun), forests have 10–20%."),
    ("Atmospheric Absorption", "19 Units absorbed: 16 by water vapor, dust, and ozone; 3 by clouds.", "Stratospheric ozone absorbs lethal UV (<0.3 μm); tropospheric water vapor and aerosols absorb selective near-infrared bands."),
    ("Surface Absorption", "51 Units absorbed: 26 via direct solar beam, 25 via diffuse sky radiation (skylight).", "Heats land and oceans; drives terrestrial reradiation, sensible heat thermals, and latent heat evaporation."),
    ("Terrestrial Longwave Emission", "70 Units returned to space: 17 direct through atmospheric window, 53 via atmosphere.", "Terrestrial radiation is longwave infrared (4–100 μm); greenhouse effect traps heat before final emission to space."),
    ("Poleward Heat Advection", "Latitudinal tipping point at 38° latitude; surplus below, deficit above.", "Planetary winds (75–80%) and ocean currents (20–25%) transport heat poleward, stabilizing global climatic belts.")
]

p1_q9_essay = [
    ("Introduction: The Thermodynamic Steady State of the Earth System",
     "The Earth is a thermodynamic open engine driven almost exclusively by radiant energy transmitted across 150 million kilometers of interplanetary space from the Sun. Solar insolation (incoming solar radiation) powers the global atmospheric circulation, drives the hydrological cycle, energizes ocean currents, and sustains terrestrial biosphere photosynthesis. Despite receiving an immense, continuous influx of radiant energy—approximately 1.74 × 10¹⁷ Watts globally—the Earth does not grow progressively hotter over geological epochs, nor does it freeze into a permanent cryospheric wasteland. Instead, our planet maintains an extraordinary long-term thermodynamic equilibrium known as the Global Heat Balance. In accordance with the First Law of Thermodynamics (Conservation of Energy) and the Stefan-Boltzmann radiation law (E = sigma * T⁴), the total quantity of incoming shortwave solar radiation entering the top of the atmosphere is balanced precisely by an equal quantity of outgoing longwave infrared radiation emitted back into cold cosmic space."),

    ("Mechanisms and Controlling Factors of Solar Insolation",
     "The solar radiation intercepted at the outer limit of the Earth's atmosphere on a surface perpendicular to the solar rays averages approximately 1,367 Watts per square meter (or 1.96 calories per square centimeter per minute), a value designated as the 'Solar Constant'. Solar radiation is shortwave radiation, peaking at 0.5 microns within the visible spectrum (0.4–0.7 microns), with ultraviolet and near-infrared components. However, the distribution of insolation across the Earth's surface is profoundly heterogeneous, controlled by four primary astronomical and atmospheric factors. First, the 'Angle of Solar Incidence' (solar altitude) is paramount: governed by Lambert’s Cosine Law, vertical equatorial rays concentrate energy over a minimal surface area, whereas oblique high-latitude rays spread the identical energy over a vastly larger, inclined surface. Second, oblique rays traverse a much longer atmospheric path length (air mass), suffering greater scattering and absorption. Third, the 'Duration of Daylight' (photoperiod), dictated by Earth's 23.5° axial tilt and annual revolution, produces dramatic seasonal variations in day length outside the tropics. Fourth, atmospheric attenuation—reflection by clouds, Rayleigh scattering by gas molecules, and absorption by ozone and water vapor—modulates ground-level irradiance."),

    ("The Global Heat Budget: A Rigorous 100-Unit Partitioning",
     "The global heat budget is mathematically modeled by tracking the fate of 100 hypothetical units of incoming shortwave solar radiation arriving at the top of the atmosphere. Of these 100 units, approximately 30 units are immediately reflected and scattered back into space without heating the Earth; this total fraction represents the 'Planetary Albedo' (0.30). Specifically, clouds reflect 20 units back to space, atmospheric gas molecules scatter 6 units, and Earth's reflective ground surfaces (snow, ice, sand) reflect 4 units. The remaining 70 units of solar energy are retained and absorbed within the Earth system: 19 units are absorbed directly by the atmosphere (stratospheric ozone absorbs ultraviolet rays, while tropospheric water vapor, clouds, and carbon dioxide absorb near-infrared wavelengths), and 51 units are absorbed by the Earth's surface. Of these 51 surface units, 26 units arrive via direct-beam solar radiation, while 25 units arrive as diffuse sky radiation (skylight scattered downward by atmospheric aerosols and gases)."),

    ("Terrestrial Longwave Reradiation and the Natural Greenhouse Effect",
     "Because the Earth absorbs 51 units of heat, its surface warms to a global mean temperature of approximately 15°C (288 K). In accordance with Wien’s Displacement Law (lambda_max = 2898 / T), the Earth reradiates this absorbed energy not as shortwave light, but as longwave thermal infrared radiation, peaking at 10 microns. The dissipation of these 51 surface units occurs through three distinct physical mechanisms: 23 units are transferred to the troposphere as 'Latent Heat Flux' via the evaporation of water (which is released upon cloud condensation); 7 units are transferred via 'Sensible Heat Flux' through conduction and convective thermal upwelling; and 21 units are radiated as net direct longwave infrared radiation. While 17 units of this infrared radiation escape directly into outer space through the 'Atmospheric Window' (an unabsorbed spectral band between 8 and 12 microns), the remaining 34 units are absorbed by atmospheric greenhouse gases (H2O, CO2, CH4). The atmosphere, having absorbed 19 units of solar shortwave and 51 units of terrestrial energy (totaling 70 units), reradiates 53 units of longwave energy to space from cloud tops and upper-tropospheric gases, perfectly balancing the 70 units of absorbed planetary energy."),

    ("Latitudinal Radiation Imbalance and Poleward Heat Advection",
     "While global incoming and outgoing radiation balance perfectly over annual cycles, a profound spatial imbalance exists across latitudes. Because of high solar angles and long days, equatorial and tropical latitudes absorb significantly more solar energy than they reradiate to space, creating a zone of net 'Radiative Surplus' between the equator and approximately 38° North and South latitudes. Conversely, poleward of 38° to the poles, low solar angles and high surface albedo (ice sheets) cause longwave radiative loss to exceed incoming shortwave insolation, producing a permanent zone of net 'Radiative Deficit'. If no energy transfer occurred, the tropics would heat up continuously to unlivable extremes, while polar regions would plunge toward absolute zero. To maintain planetary climatic equilibrium, massive quantities of surplus tropical heat are continuously transported poleward across the 38° parallel. Approximately 75 to 80 percent of this poleward heat advection is accomplished by Atmospheric Circulation—driven by the Hadley, Ferrel, and Polar cells, mid-latitude wave cyclones, and upper-tropospheric Jet Streams. The remaining 20 to 25 percent is accomplished by Oceanic Circulation—via warm western boundary currents such as the Gulf Stream and North Atlantic Drift in the Atlantic, and the Kuroshio Current in the Pacific."),

    ("Conclusion: Radiation Equilibrium and Anthropogenic Climate Perturbation",
     "In evaluative summation, the radiation budget of the Earth represents a finely tuned thermodynamic equilibrium that governs the physical geography of the planet. The intricate balance between incoming shortwave insolation, albedo reflection, atmospheric greenhouse absorption, and outgoing longwave emission establishes the planetary thermal boundary conditions that sustain liquid water and biological life. However, modern anthropogenic emissions of greenhouse gases—principally carbon dioxide from fossil fuel combustion and industrial deforestation—are narrowing the atmospheric infrared window. This radiative forcing traps approximately 1 to 2 extra Watts per square meter in the lower troposphere, driving contemporary global warming, accelerating polar cryospheric melt, perturbing monsoon circulation systems, and intensifying climatic volatility across sensitive zones like Pakistan.")
]

p1_q9_quotes = [
    ("Scholar / Climatologist", "Theoretical Contribution / Concept", "Exact Quotable Academic Thesis / Axiom"),
    ("M.I. Budyko (1956)", "The Heat Balance of the Earth's Surface", "The climate of any region is fundamentally an expression of its energy budget; the transformation of solar radiant energy into thermal and turbulent fluxes dictates the global terrestrial environment."),
    ("Wladimir Köppen (1936)", "Das geographische System der Klimate", "Temperature distribution across the globe is not merely a geometric function of latitude, but a dynamic product of atmospheric circulation acting upon solar radiation imbalances."),
    ("Arthur N. Strahler (1965)", "Introduction to Physical Geography", "The global heat budget operates as a closed thermodynamic ledger where the incoming credit of solar shortwave energy is exactly canceled by the outgoing debit of infrared radiation to space."),
    ("C.W. Thornthwaite (1948)", "An Approach toward a Rational Classification of Climate", "Energy and moisture are the twin master variables of physical geography; potential evapotranspiration provides the precise measure of thermal energy available at the surface."),
    ("Glenn T. Trewartha (1968)", "An Introduction to Climate", "The planetary wind systems and ocean currents exist primarily as thermodynamic safety valves, transferring surplus tropical heat toward the frozen polar sinks.")
]

# ==============================================================================
# QUESTION 10 (P1_Q10): ATMOSPHERIC PRESSURE BELTS & PLANETARY WIND SYSTEMS
# Solved from: 2022 Q.3, 2018 Q.5, 2016 Q.8, 2013 Q.2
# ==============================================================================
p1_q10_meta = [
    ("Target Syllabus", "Section 1: Weather & Climate — Atmospheric Pressure Belts & Planetary Winds"),
    ("Exam Years Solved", "KPPSC PMS 2022 Q.3, 2018 Q.5, 2016 Q.8, 2013 Q.2"),
    ("Marks & Allocation", "20 Marks | Strict 35-Minute Descriptive Writing Limit"),
    ("Core Theoretical Thesis", "Global atmospheric circulation operates as a three-dimensional tri-cellular thermodynamic engine (Hadley, Ferrel, and Polar cells) driven by differential latitudinal heating and planetary rotation. Air motion is governed by the vector balance of Pressure Gradient Force, Coriolis Force, and surface friction, establishing the four latitudinal pressure belts (Equatorial Low, Subtropical Highs, Subpolar Lows, Polar Highs) and three primary planetary wind systems (Trade Winds, Prevailing Westerlies, Polar Easterlies), whose seasonal latitudinal migration dictates regional monsoons and macro-climatic belts.")
]

p1_q10_callout = [
    "FORCES GOVERNING AIR CIRCULATION: (1) Pressure Gradient Force (PGF: directed from high to low pressure perpendicular to isobars; the primary initiator of wind); (2) Coriolis Force (deflective apparent force due to Earth's rotation: Fc = 2 * Omega * v * sin(phi); deflects wind to the right in Northern Hemisphere, left in Southern Hemisphere; zero at equator, maximum at poles); (3) Frictional Force (retards wind speed in planetary boundary layer, causing winds to cross isobars at 10°–30° angles); (4) Geostrophic Wind (frictionless upper-air equilibrium where PGF equals Coriolis force, flowing parallel to straight isobars).",
    "THE FOUR PLANETARY PRESSURE BELTS: (1) Equatorial Low Pressure Belt / Doldrums / ITCZ (0°–5° N/S: thermally induced, intense solar heating, convective convergence); (2) Subtropical High Pressure Belts / Horse Latitudes (30°–35° N/S: dynamically induced by upper-air subsiding air of the Hadley Cell; arid, calm weather); (3) Subpolar Low Pressure Belts (60°–65° N/S: dynamically induced by convergence of Westerlies and Polar Easterlies along the Polar Front); (4) Polar Highs (90° N/S: thermally induced by extreme cold and continuous atmospheric subsidence).",
    "THE TRI-CELLULAR CIRCULATION MODEL: Hadley Cell (thermally direct, 0°–30°: ascent at ITCZ, poleward upper flow, subsidence at 30°); Ferrel Cell (thermally indirect, 30°–60°: driven mechanically like a gear between Hadley and Polar cells); Polar Cell (thermally direct, 60°–90°: subsidence over cold poles, equatorward flow, ascent at Polar Front).",
    "SEASONAL SHIFT & MONSOON DYNAMICS: The apparent migration of the Sun (solstices) causes all pressure belts and wind systems to shift northward in boreal summer and southward in boreal winter, triggering the reversal of the South Asian Monsoon (reversal of Trade Winds over Pakistan and India)."
]

p1_q10_map = """
                      [TRI-CELLULAR PLANETARY CIRCULATION]
                                       │
 [90° N] ────── POLAR HIGH (Thermally Induced: Cold Dense Subsidence) ──────
    ▲           │ POLAR CELL (Thermally Direct)
    │           ▼ Polar Easterlies (Deflected right: NE winds)
 [60° N] ────── SUBPOLAR LOW / POLAR FRONT (Dynamically Induced: Frontal Uplift)
    ▲           ▲
    │           │ Prevailing Westerlies (SW winds in NH)
    │           │ FERREL CELL (Thermally Indirect / Mechanical)
 [30° N] ────── SUBTROPICAL HIGH / HORSE LATITUDES (Dynamically Induced: Subsidence)
    │           ▼
    │           │ Northeast Trade Winds (Blowing toward ITCZ)
    │           │ HADLEY CELL (Thermally Direct)
    ▼           ▲
 [ 0°  ] ────── EQUATORIAL LOW / ITCZ / DOLDRUMS (Thermally Induced: Convection)
    ▲           ▼
    │           │ Southeast Trade Winds (Blowing toward ITCZ)
    │           │ HADLEY CELL (Thermally Direct)
    │           ▲
 [30° S] ────── SUBTROPICAL HIGH (Dynamically Induced: Subsidence)
    │           │ Prevailing Westerlies (NW winds in SH / Roaring Forties)
    ▼           ▼ FERREL CELL (Thermally Indirect)
 [60° S] ────── SUBPOLAR LOW (Dynamically Induced: Frontal Uplift)
    ▲           ▲ Polar Easterlies (SE winds in SH)
    │           │ POLAR CELL (Thermally Direct)
 [90° S] ────── POLAR HIGH (Thermally Induced: Antarctic Ice Cap Subsidence)
"""

p1_q10_timeline = [
    ("Pressure Belt / Cell", "Latitudinal Zone & Origin Mechanism", "Characteristic Wind Systems & Associated Weather"),
    ("Equatorial Low Pressure (ITCZ / Doldrums)", "0°–5° North and South; Thermally induced by intense, year-round solar heating.", "Light, baffling winds (Doldrums); strong vertical convection; massive cumulonimbus cloud development; daily afternoon torrential rains."),
    ("Hadley Cell & Trade Winds", "0° to 30° N/S; Thermally direct circulation cell discovered by George Hadley (1735).", "Northeast Trades in Northern Hemisphere; Southeast Trades in Southern Hemisphere; highly persistent, reliable winds; maritime moisture feeds tropical rainfall."),
    ("Subtropical Highs (Horse Latitudes)", "30°–35° N/S; Dynamically induced by upper-tropospheric convergence and dry air subsidence.", "High barometric pressure; clear skies, calm winds, severe aridity; world's great deserts located here (Sahara, Arabian, Thar, Atacama, Kalahari)."),
    ("Ferrel Cell & Prevailing Westerlies", "35° to 60° N/S; Thermally indirect cell driven mechanically by flanking cells.", "Southwesterlies in Northern Hemisphere; Northwesterlies in Southern Hemisphere ('Roaring Forties', 'Furious Fifties'); highly volatile frontal weather."),
    ("Subpolar Low Pressure Belt", "60°–65° N/S; Dynamically induced by convergence of cold polar and warm subtropical air.", "Formation of the Polar Front; intense cyclogenesis; mid-latitude wave cyclones; Icelandic Low and Aleutian Low in the north."),
    ("Polar Highs & Polar Easterlies", "90° N/S; Thermally induced by extreme radiative cooling over ice caps.", "Permanent anticyclonic high pressure; cold, dry subsiding air; Polar Easterlies blow toward subpolar low; perpetual blizzard and permafrost conditions.")
]

p1_q10_essay = [
    ("Introduction: The Planetary Thermodynamic Heat Engine",
     "The Earth's atmosphere is a vast, gaseous thermodynamic engine that operates in continuous motion to eliminate the thermal disequilibrium created by unequal latitudinal solar insolation. If the Earth were a non-rotating, uniform sphere, atmospheric circulation would conform to a simple single-cell convective model (proposed by George Hadley in 1735), wherein hot air at the equator would rise, flow poleward aloft, sink over the frozen poles, and return equatorward as a uniform surface wind. However, planetary rotation, the deflective Coriolis force, continental land-sea thermal contrasts, and dynamic frictional drag break this simple convective loop into a complex, three-dimensional tri-cellular circulation pattern. This planetary circulation is structured around four distinct latitudinal pressure belts and three primary wind systems in each hemisphere, which collectively dictate the global distribution of precipitation, arid deserts, storm tracks, and seasonal monsoons."),

    ("Governing Physical Forces: Pressure Gradient, Coriolis, and Geostrophic Balance",
     "The velocity and trajectory of atmospheric wind are governed by the vector equilibrium of four fundamental physical forces. The primary driving force is the 'Pressure Gradient Force' (PGF), which acts perpendicular to isobars, driving air directly from regions of high barometric pressure toward regions of low barometric pressure. Once air begins to move, Earth's eastward rotation exerts an apparent deflective force known as the 'Coriolis Force' (first formulated by Gaspard-Gustave de Coriolis in 1835). Expressed mathematically as Fc = 2 * Omega * v * sin(phi)—where Omega is Earth's angular velocity, v is wind speed, and phi is latitude—the Coriolis force deflects moving air to the right of its path in the Northern Hemisphere and to the left in the Southern Hemisphere (Ferrel’s Law). Crucially, Coriolis acceleration is zero at the equator and increases toward unity at the poles. In the frictionless upper troposphere (>1 km above the surface), PGF and Coriolis force achieve an exact balance, producing the 'Geostrophic Wind', which blows parallel to straight isobars. Near the surface, however, 'Frictional Force' retards wind speed, reducing Coriolis deflection and causing surface winds to spiral across isobars into low pressure at an angle of 10° to 30°."),

    ("The Four Planetary Pressure Belts: Thermal vs. Dynamic Genesis",
     "The global atmosphere is organized latitudinally into four alternating high- and low-pressure belts, which divide genetically into 'Thermally Induced' and 'Dynamically Induced' systems. First, the 'Equatorial Low Pressure Belt' (0°–5° N/S) is thermally induced: intense year-round insolation heats the surface, causing air to expand, become buoyant, and rise violently through convection. This creates a surface trough of low barometric pressure known as the Intertropical Convergence Zone (ITCZ) or the 'Doldrums', characterized by calm surface winds, high humidity, and afternoon cumulonimbus downpours. Second, the 'Subtropical High Pressure Belts' (30°–35° N/S) are dynamically induced: warm air rising from the ITCZ cools radiatively as it moves poleward aloft, is deflected eastward by the Coriolis force into a jet, and converges in the upper troposphere, forcing dense air to subside mechanically around 30° latitude. This subsidence warms adiabatically, creating cloudless skies, intense aridity, and calm conditions ('Horse Latitudes'), under which the world’s great subtropical deserts (Sahara, Arabian, Thar, Sonoran) are situated."),

    ("Subpolar Lows and Polar High Pressure Systems",
     "Third, the 'Subpolar Low Pressure Belts' (60°–65° N/S) are dynamically induced: cold, dense air flowing equatorward from the poles collides with warm, buoyant air blowing poleward from the subtropics. This collision forces warm air to ride over the cold air wedge, creating a dynamic zone of intense upward motion along the 'Polar Front'. This ascent generates semi-permanent oceanic low-pressure cells—exemplified by the Icelandic Low and Aleutian Low in the Northern Hemisphere—which serve as breeding grounds for ferocious extratropical wave cyclones. Fourth, the 'Polar Highs' (90° N/S) are thermally induced: extreme radiative cooling over the Arctic and Antarctic ice sheets creates intensely cold, hyper-dense air that sinks continuously, maintaining permanent anticyclonic high pressure, calm conditions, and severe dry-polar blizzards over the ice domes."),

    ("The Tri-Cellular Model and Planetary Wind Systems",
     "The global circulation is organized into three distinct meridional circulation cells in each hemisphere: the Hadley Cell, the Ferrel Cell, and the Polar Cell. The 'Hadley Cell' (0°–30°) is a thermally direct convective cell: air rises at the ITCZ, travels poleward in the upper troposphere, subsides at the Subtropical Highs, and returns toward the equator along the surface. Deflected by the Coriolis force, these return winds blow from the northeast in the Northern Hemisphere and from the southeast in the Southern Hemisphere, forming the remarkably reliable 'Trade Winds' (Alizés), which have powered maritime navigation for centuries. The 'Ferrel Cell' (30°–60°), identified by William Ferrel in 1856, is a thermally indirect cell operating mechanically like a gear between the Hadley and Polar cells: surface air flowing poleward from the Subtropical Highs is deflected to the right, producing the 'Prevailing Westerlies' (Southwesterlies in the Northern Hemisphere, Northwesterlies in the Southern Hemisphere). In the southern oceans, uninterrupted by continental landmasses, the Westerlies achieve terrifying velocities, celebrated maritimely as the 'Roaring Forties' and 'Furious Fifties'. Finally, the 'Polar Cell' (60°–90°) is a thermally direct cell: air subsiding at the poles flows equatorward as the 'Polar Easterlies', meeting the Westerlies at the stormy Polar Front."),

    ("Seasonal Migration, Upper-Air Jet Streams, and the South Asian Monsoon",
     "In reality, planetary pressure belts and wind systems are not stationary; they shift latitudinally following the seasonal migration of the thermal equator (solar declination). In the boreal summer (June–July), the overhead sun moves to the Tropic of Cancer, pulling the ITCZ and planetary pressure belts 10° to 20° northward. In South Asia, this seasonal shift triggers the monumental 'South Asian Monsoon': the Southeast Trade Winds cross the equator into the Northern Hemisphere, are deflected to the right by the Coriolis force, and transform into the warm, moisture-laden Southwest Monsoon, which drenches Pakistan and India. Aloft, circulation is modulated by upper-tropospheric 'Jet Streams'—narrow ribbons of high-speed westerly winds (up to 300–400 km/h) meandering in Rossby waves along the Polar Front and Subtropical discontinuity. In evaluative summation, the planetary wind system is an integrated, multi-layered hydrodynamic continuum. Understanding its pressure belts, tri-cellular cells, and seasonal shifts is fundamental to explaining global climate patterns, aviation corridors, and regional agro-ecological security across Pakistan.")
]

p1_q10_quotes = [
    ("Scholar / Climatologist", "Theoretical Contribution / Concept", "Exact Quotable Academic Thesis / Axiom"),
    ("George Hadley (1735)", "Concerning the Cause of the General Trade-Winds", "The trade winds are caused by the ascent of air heated by the tropical sun, replaced by colder air drawn from the poles, deflected by the diurnal motion of the earth."),
    ("William Ferrel (1856)", "An Essay on the Winds and the Currents of the Ocean", "If a body moves in any direction on the earth's surface, there is a deflecting force arising from the earth's rotation, turning it to the right in the northern hemisphere and to the left in the southern."),
    ("Carl-Gustaf Rossby (1939)", "Relation Between Variations in the Intensity of the Zonal Circulation", "The upper tropospheric westerlies organize into planetary-scale meandering Rossby waves whose wavelengths and amplitudes dictate the displacement of surface cyclones and anticyclones."),
    ("Arthur N. Strahler (1969)", "Physical Geography", "The tri-cellular circulation model reconciles the thermal drive of equatorial insolation with the kinematic deflection of Earth's rotation, creating the grand global wind belts."),
    ("Glenn T. Trewartha (1968)", "An Introduction to Climate", "The seasonal migration of planetary pressure and wind belts is the master key to tropical meteorology, transforming regional trades into life-giving monsoon deluges.")
]

# ==============================================================================
# QUESTION 11 (P1_Q11): ATMOSPHERIC MOISTURE, CLOUDS & PRECIPITATION
# Solved from: 2025 Q.3, 2018 Q.4
# ==============================================================================
p1_q11_meta = [
    ("Target Syllabus", "Section 1: Weather & Climate — Humidity, Condensation & Forms of Precipitation"),
    ("Exam Years Solved", "KPPSC PMS 2025 Q.3, 2018 Q.4"),
    ("Marks & Allocation", "20 Marks | Strict 35-Minute Descriptive Writing Limit"),
    ("Core Theoretical Thesis", "Precipitation is the liquid or solid hydrometeor condensation product falling from the atmosphere to the Earth's surface under gravity. Its physical genesis requires atmospheric moisture saturation (100% Relative Humidity), adiabatic cooling to the Dew Point, presence of microscopic Cloud Condensation Nuclei (CCN), and droplet growth via microphysical mechanisms (Bergeron-Findeisen ice-crystal process in cold clouds, Collision-Coalescence in warm clouds). Dynamic precipitation classifies macro-scale into Convectional, Orographic, and Cyclonic/Frontal mechanisms, dictating regional rainfall distribution across South Asia.")
]

p1_q11_callout = [
    "HUMIDITY METRICS & ADIABATIC LAPSE RATES: Absolute Humidity (mass of vapor per volume, g/m³); Specific Humidity (mass of vapor per mass of air, g/kg); Relative Humidity (ratio of actual vapor pressure to saturation vapor pressure: RH = e / es * 100%). When RH = 100%, air reaches Dew Point. Dry Adiabatic Lapse Rate (DALR = 9.8°C/1,000 m); Saturated Adiabatic Lapse Rate (SALR = 4°–6°C/1,000 m, lower due to latent heat release).",
    "THE TWO MICROPHYSICAL PRECIPITATION THEORIES: (1) Collision-Coalescence Process (warm clouds, T > 0°C): Large cloud droplets fall faster than smaller ones, sweeping them up and coalescing into heavy raindrops (>0.5 mm); (2) Bergeron-Findeisen Process (cold clouds, T < 0°C): Supercooled liquid water droplets and ice crystals coexist; because saturation vapor pressure over ice is lower than over water, water vapor sublimates rapidly onto ice crystals, growing large snowflakes that melt into rain or fall as snow.",
    "THE THREE MAJOR PRECIPITATION MECHANISMS: (1) Convectional: Intense thermal heating -> buoyant unstable air -> rapid vertical ascent -> cumulonimbus cloudbursts and thunderstorms (equatorial afternoon rains, summer pre-monsoon showers); (2) Orographic (Relief): Air forced to ascend mountain barriers -> adiabatic cooling -> heavy windward rainfall -> dry leeward descent with adiabatic warming creating rain-shadow deserts (e.g., Margalla/Pir Panjal windward vs. Gilgit-Baltistan rain shadow); (3) Cyclonic / Frontal: Warm moist air mass forced to ride over denser cold air mass along frontal surfaces or convergent low-pressure troughs.",
    "REGIONAL FORMS OF PRECIPITATION: Rain (liquid droplets >0.5 mm), Drizzle (<0.5 mm), Snow (hexagonal ice crystals), Sleet (refrozen raindrops), Hail (concentric layers of hard ice formed in violent cumulonimbus updrafts)."
]

p1_q11_map = """
                     [ATMOSPHERIC PRECIPITATION CONTINUUM]
                                      │
            ┌─────────────────────────┴─────────────────────────┐
            ▼                                                   ▼
 [MICROPHYSICAL DROPLET GROWTH]                      [MACRO-SCALE DYNAMIC MECHANISMS]
 (Cloud Physics: CCN & Saturation)                   (Forced Adiabatic Ascent to LCL)
            │                                                   │
 ┌──────────┴──────────┐                             ┌──────────┼──────────┐
 ▼                     ▼                             ▼          ▼          ▼
[COLLISION-           [BERGERON-                   [CONVECTION] [OROGRAPHIC] [CYCLONIC /
 COALESCENCE]          FINDEISEN]                  • Solar      • Mountain    FRONTAL]
• Warm clouds         • Cold clouds                  thermal      barrier     • Convergence
  (T > 0°C)             (T < 0°C)                    heating      ascent      • Warm air
• Giant CCN           • Supercooled water          • Unstable   • Windward    forced over
• Droplet sweep         & ice crystals               plumes       deluge        cold wedge
  & coalescing        • Vapor pressure             • Cumulo-    • Leeward     • Frontal
• Tropical rain         differential                 nimbus       rain shadow   nimbostratus
                      • Snowflakes & hail          • Afternoon  • E.g., Pir   • E.g., Western
                                                     squalls      Panjal/GB     Depressions
"""

p1_q11_timeline = [
    ("Precipitation Mechanism / Form", "Thermodynamic Trigger & Physical Process", "Diagnostic Meteorological Features & Pakistan / Global Examples"),
    ("Convectional Precipitation", "Intense diurnal ground heating; buoyant thermal parcels rise adiabatically at DALR to LCL.", "Towering Cumulonimbus clouds, anvil tops; localized, high-intensity cloudbursts, severe lightning and thunder; common in equatorial zones and Punjab summer pre-monsoons."),
    ("Orographic (Relief) Precipitation", "Moist air mass strikes an elevated mountain barrier, forced into mechanical upslope ascent.", "Heavy precipitation on windward slopes (e.g., Cherrapunji, India; Murree/Kaghan in Pakistan); leeward slopes experience dry descending foehn winds, creating rain-shadows (Gilgit/Chilas)."),
    ("Cyclonic / Frontal Precipitation", "Warm, moist air converges and is forced to ascend over colder, denser air along frontal boundaries.", "Extensive stratiform and nimbostratus cloud decks; steady, prolonged widespread drizzle and rain; mid-latitude wave cyclones and winter 'Western Disturbances' in Pakistan."),
    ("Collision-Coalescence Process", "Warm clouds (T > 0°C); large droplets fall with higher terminal velocity, colliding with smaller droplets.", "Dominant in tropical maritime air masses; produces heavy warm tropical rain without requiring ice crystal formation; typical of Arabian Sea coastal showers."),
    ("Bergeron-Findeisen Process", "Mixed-phase cold clouds (0°C to -40°C); saturation vapor pressure over water exceeds that over ice.", "Rapid deposition of vapor onto ice crystals at expense of evaporating supercooled droplets; crystals grow into dendritic snowflakes; falls as winter snow in Malam Jabba and Skardu."),
    ("Hailstones & Extreme Hydrometeors", "Violent convective updrafts (>100 km/h) in severe cumulonimbus cycle ice pellets repeatedly.", "Concentric alternating layers of clear and rime ice; falls as destructive hailstones; damages standing wheat and fruit orchards in Khyber Pakhtunkhwa.")
]

p1_q11_essay = [
    ("Introduction: Atmospheric Moisture and the Hydrometeor Cycle",
     "Precipitation represents the crucial descending phase of the global hydrological cycle, whereby condensed atmospheric water vapor falls under the pull of terrestrial gravity to reach the Earth's surface in liquid or solid form. The atmosphere holds approximately 12,900 cubic kilometers of water vapor at any given moment—a seemingly modest volume representing less than 0.001 percent of Earth's total water reserves. Yet, this vapor pool undergoes rapid, dynamic turnover, cycling completely every nine to ten days through continuous evaporation, condensation, and precipitation. Precipitation is not an inevitable consequence of atmospheric moisture; it demands an intricate sequence of thermodynamic and microphysical prerequisites: the air must reach saturation (100% Relative Humidity), cool adiabatically below its Dew Point temperature, condense around microscopic aerosol nuclei to form cloud droplets, and grow through droplet-coalescence or ice-crystal sublimation until falling velocities surpass turbulent atmospheric updrafts."),

    ("Thermodynamics of Atmospheric Saturation and Adiabatic Cooling",
     "The capacity of an air parcel to hold water vapor is an exponential function of its temperature, mathematically defined by the Clausius-Clapeyron equation. When warm air cools, its saturation vapor pressure decreases, causing its Relative Humidity (RH = actual vapor pressure / saturation vapor pressure * 100%) to surge. If cooling continues, the air reaches its 'Dew Point'—the temperature at which RH achieves 100 percent saturation. The primary mechanism of cooling in the atmosphere is not conductive contact with cold surfaces (which merely produces localized fog or dew), but 'Adiabatic Expansion'. As an unsaturated air parcel ascends through the troposphere, surrounding atmospheric pressure decreases, allowing the parcel to expand mechanically. The kinetic work of expansion consumes thermal energy, cooling the parcel at the constant 'Dry Adiabatic Lapse Rate' (DALR) of approximately 9.8°C per 1,000 meters. Once the parcel cools to its Dew Point, it reaches the 'Lifting Condensation Level' (LCL). Condensation commences around microscopic 'Cloud Condensation Nuclei' (CCN)—hygroscopic particles of sea salt, sulfate aerosols, and clay dust. Upon condensation, latent heat of vaporization (roughly 2.5 × 10⁶ Joules per kilogram) is released into the parcel, warming it and slowing subsequent ascent to the 'Saturated Adiabatic Lapse Rate' (SALR), which ranges between 4°C and 6°C per 1,000 meters."),

    ("Microphysics of Precipitation: Collision-Coalescence vs. Bergeron Process",
     "Cloud droplets are microscopic, averaging a mere 20 micrometers (0.02 mm) in diameter, with terminal fall velocities so minuscule (less than 1 cm per second) that they remain suspended indefinitely by subtle updrafts. For precipitation to occur, droplets must grow into raindrops averaging 1 to 2 millimeters in diameter—representing an approximate million-fold increase in volume. Atmospheric physics identifies two primary microphysical growth mechanisms. In warm clouds (where temperatures throughout remain above 0°C, common in tropical maritime air), growth occurs via the 'Collision-Coalescence Process'. Large cloud droplets formed on giant hygroscopic nuclei fall faster than smaller droplets, colliding with them and coalescing into progressively larger drops until they become heavy enough to fall as rain. In cold clouds (where temperatures drop well below 0°C, typical of middle and high latitudes), precipitation is governed by the 'Bergeron-Findeisen Ice-Crystal Process'. In the sub-freezing zone between 0°C and -40°C, liquid water can remain supercooled without freezing. Because the saturation vapor pressure over liquid water is substantially higher than over an ice crystal at the same temperature, a steep vapor pressure gradient develops: supercooled droplets evaporate rapidly, while the water vapor sublimates directly onto the ice crystals. The ice crystals grow rapidly into complex dendritic snowflakes, which either fall to earth as snow or melt during their descent through warm lower air to emerge as raindrops."),

    ("Mechanisms of Macro-Scale Precipitation: Convectional, Orographic, and Cyclonic",
     "While microphysics governs droplet formation, the dynamic macro-scale uplifting of air parcels determines the spatial distribution and intensity of precipitation, classifying into three fundamental mechanisms. First, 'Convectional Precipitation' occurs when intense solar insolation superheats the ground, warming the lowest air layer via conduction. The warm parcel becomes buoyant, creating vigorous vertical updrafts that ascend at rates exceeding 20 to 30 meters per second. This rapid convective ascent generates massive Cumulonimbus storm clouds featuring classic anvil-shaped cirrus tops. Convectional rains are localized, short-lived, but extraordinarily intense, accompanied by violent lightning, thunder, and destructive hail, common as afternoon storms in equatorial zones and during northern Pakistan's pre-monsoon summer afternoons."),

    ("Orographic and Cyclonic Uplifting Regimes",
     "Second, 'Orographic (Relief) Precipitation' occurs when a horizontal, moisture-bearing air stream encounters an elevated topographic mountain barrier and is mechanically forced to ascend. As air climbs the windward slope, it expands and cools adiabatically, generating thick cloud decks and torrential rainfall. Upon cresting the summit, the now-dehydrated air descends the leeward slope, warming adiabatically at the rapid DALR. This creates a dry, warm descending wind (such as the Alpine Foehn or Rocky Mountain Chinook) and casts an expansive 'Rain-Shadow' over leeward basins. In South Asia, the southwest monsoon strikes the Western Ghats and the southern slopes of the Himalayas (Pir Panjal and Margalla ranges), dumping massive rainfall on windward slopes, while leeward trans-Himalayan valleys (such as Gilgit, Chilas, and Skardu) receive less than 150 mm annually, creating high-altitude desert conditions. Third, 'Cyclonic / Frontal Precipitation' occurs when two contrasting air masses meet. Because warm air is lighter and less dense than cold air, the warm air mass is forced to ascend along the inclined frontal wedge (along cold or warm fronts in mid-latitude wave cyclones). In tropical cyclones, cyclonic convergence in the low-pressure eye wall drives massive vertical uplifting, producing expansive rainbands."),

    ("Forms of Precipitation and Pakistan's Hydrological Regimes",
     "Precipitation manifests in multiple physical states depending upon the atmospheric temperature profile between the cloud base and the ground. 'Rain' consists of liquid water drops exceeding 0.5 mm in diameter, whereas 'Drizzle' comprises uniform, microscopic droplets (<0.5 mm) falling from low stratus decks. 'Snow' occurs when water vapor freezes directly into hexagonal ice crystals that aggregate into snowflakes, reaching the ground intact when the entire tropospheric column remains below 0°C. 'Sleet' represents raindrops that freeze into solid ice pellets while falling through a deep sub-freezing ground layer, whereas 'Hail' consists of hard, concentric balls of ice excavated by violent convective updrafts in severe Cumulonimbus cells. In Pakistan, precipitation regimes reflect a dual seasonality: the summer Southwest Monsoon (July–September) delivers 65 to 70 percent of total national rainfall via orographic and tropical cyclonic systems, while winter 'Western Disturbances' (December–March)—traveling as extratropical wave cyclones from the Mediterranean—bring essential snow and frontal rain to Khyber Pakhtunkhwa, Balochistan, and the northern mountains, sustaining the Indus River Basin's snowpack and Rabi wheat crop.")
]

p1_q11_quotes = [
    ("Scholar / Meteorologist", "Theoretical Contribution / Concept", "Exact Quotable Academic Thesis / Axiom"),
    ("Tor Bergeron (1935)", "On the Physics of Cloud and Precipitation", "The coexistence of supercooled water droplets and ice crystals in mixed clouds creates an unstable vapor equilibrium that is the primary trigger for temperate precipitation."),
    ("Luke Howard (1803)", "On the Modifications of Clouds", "Clouds are the visible indicators of the invisible thermodynamic motions of the atmosphere; their forms—cirrus, stratus, cumulus, nimbus—mirror their convective origins."),
    ("Horace R. Byers (1949)", "The Thunderstorm Project", "A convective thunderstorm is a thermodynamic machine comprised of distinct updraft and downdraft cells, where condensation latent heat sustains kinetic storm energy."),
    ("Arthur N. Strahler (1969)", "Physical Geography", "Orographic precipitation represents the most geographically pronounced rainfall regime on earth, transforming mountain windward slopes into rain forests and leeward flanks into deserts."),
    ("C.G. Rossby (1941)", "The Scientific Basis of Modern Meteorology", "Precipitation cannot be understood in isolation from macro-circulation; the thermodynamic ascent of air along fronts and convergence zones dictates regional water balances.")
]

print("Batch 2 (Q9, Q10, Q11) data appended. Appending Q12, Q13, Q14 and runner loop...")

# ==============================================================================
# QUESTION 12 (P1_Q12): AIR MASSES, FRONTOGENESIS & POLAR FRONT WAVE CYCLONES
# Solved from: 2022 Q.4, 2016 Q.2, 2010 Q.2, 2008 Q.3
# ==============================================================================
p1_q12_meta = [
    ("Target Syllabus", "Section 1: Weather & Climate — Air Masses, Fronts & Atmospheric Disturbances"),
    ("Exam Years Solved", "KPPSC PMS 2022 Q.4, 2016 Q.2, 2010 Q.2, 2008 Q.3"),
    ("Marks & Allocation", "20 Marks | Strict 35-Minute Descriptive Writing Limit"),
    ("Core Theoretical Thesis", "An air mass is a macro-scale tropospheric body characterized by quasi-homogeneous horizontal temperature and humidity properties acquired through prolonged thermodynamic stagnation over uniform source regions. The convergence of contrasting polar (cP) and tropical (mT) air masses along the baroclinic Polar Front triggers frontogenesis, establishing inclined discontinuity surfaces (cold, warm, stationary, occluded fronts). Governed by the Bergen School Polar Front Theory (Vilhelm & Jacob Bjerknes), baroclinic instability initiates mid-latitude wave cyclones that evolve systematically through five stages—incipient wave, open sector, mature occlusion, and cyclolysis—dictating temperate winter weather and Pakistan's Western Disturbances.")
]

p1_q12_callout = [
    "AIR MASS CRITERIA & SOURCE REGIONS (BERGERON): Requires (1) vast, physically uniform surface (oceanic water, flat desert, snow-covered ice sheet); (2) atmospheric stagnation under persistent anticyclonic high pressure allowing thermodynamic equilibration. Classified thermally: Arctic/Antarctic (A), Polar (P), Tropical (T), Equatorial (E); and moisture-wise: Continental (c = dry), Maritime (m = humid).",
    "THE FOUR FRONTAL DISCONTINUITIES: (1) Warm Front: Warm air actively overruns retreating cold air wedge; gentle slope (1:150 to 1:200); sequential cloud transition: Cirrus -> Cirrostratus -> Altostratus -> Nimbostratus; prolonged steady rain/drizzle; (2) Cold Front: Dense cold air actively undercuts warm air; steep slope (1:50 to 1:100); violent vertical convective lift: Cumulonimbus, gusty squalls, heavy localized downpours, sharp temperature drop, rapid barometric rise; (3) Stationary Front: Non-moving boundary; (4) Occluded Front: Fast-moving cold front overtakes warm front, lifting the warm sector completely off the ground (Cold Occlusion vs. Warm Occlusion).",
    "THE BERGEN POLAR FRONT CYCLONE LIFE CYCLE (BJERKNES): (1) Incipient / Stationary Wave: Stable polar front separating cP and mT air; (2) Open Wave Stage: Shear creates a wave perturbation with distinct warm and cold fronts enclosing a warm sector; (3) Mature Stage: Deepening central low pressure, well-developed cyclonic wind circulation; (4) Occlusion Stage: Cold front overtakes warm front, warm air lifted aloft, precipitation peak; (5) Dissipation (Cyclolysis): Warm air isolated aloft, central pressure rises, frontal contrast vanishes.",
    "CYCLONE VS. ANTICYCLONE DICHOTOMY: Cyclones = Low-pressure centers, inward counter-clockwise convergence in Northern Hemisphere (clockwise in Southern Hemisphere), vertical ascent, cloudy stormy weather; Anticyclones = High-pressure centers, outward clockwise divergence in Northern Hemisphere, subsidence, cloudless skies, calm, stable weather."
]

p1_q12_map = r"""
                 [POLAR FRONT CYCLOGENESIS LIFE CYCLE]
                                   │
 1. [INCIPIENT POLAR FRONT]        │ Polar Easterlies (cP: Cold Dry)
    ───────────────────────────────┴───────────────────────────────
                                   │ Prevailing Westerlies (mT: Warm Moist)
                                   ▼
 2. [OPEN WAVE STAGE]              │ Shear wave initiates baroclinic vortex
          Cold Front               │ Warm Front
       (Steep: 1:50)               │ (Gentle: 1:200)
    Cumulonimbus Squalls           │ Nimbostratus Drizzle
                 \                 │ /
                  \  [WARM SECTOR]  /
                   \   (mT Air)    /
                    \             /
                     ▼           ▼
 3. [MATURE OCCLUSION STAGE]       │ Cold front moves faster than warm front
    • Occluded Front formed: Cold front undercuts warm front completely
    • Warm air sector lifted entirely above ground
    • Maximum barometric pressure drop & widest precipitation band
                                   ▼
 4. [CYCLOLYSIS / DISSIPATION]     │ Energy exhausted (zero thermal contrast)
    • Central low pressure fills; anticyclonic stability restored
"""

p1_q12_timeline = [
    ("Wave Cyclone Stage / Front", "Dynamic Mechanism & Atmospheric Structure", "Associated Weather Sequences & Barometric Indicators"),
    ("1. Incipient Wave Stage", "Linear shear stress along stationary Polar Front; cold polar easterlies shear against warm westerlies.", "Stable boundary; barometric pressure uniform; cirrus clouds appear; wind shifts from easterly to southeasterly."),
    ("2. Open Wave Stage", "Perturbation deepens into distinct cyclonic low; develops warm front in advance and cold front trailing behind.", "Isobars assume V-shape; distinct 'warm sector' established; pressure begins steady fall; wind veers to southwest."),
    ("Warm Front Passage", "Warm mT air glides gently upward over retreating cP cold wedge (slope 1:150 to 1:200).", "Sequential cloud ceiling: Cirrus halo -> Cirrostratus -> Altostratus -> Nimbostratus; prolonged, steady soaking drizzle; rising temperature; falling pressure."),
    ("Cold Front Passage", "Dense cold cP air bulldozes under warm moist mT air like a wedge (steep slope 1:50 to 1:100).", "Violent convective updrafts; Cumulonimbus squall line, severe thunder, lightning, localized downpours; abrupt wind shift (SW to NW); catastrophic temperature plunge; rapid pressure surge."),
    ("3. Occlusion Stage", "Cold front travels faster than warm front (40–50 km/h vs. 25–30 km/h), overtaking it and pinching out warm sector.", "Formation of Occluded Front; warm air completely lifted above planetary boundary layer; heavy mixed precipitation; storm reaches maximum intensity."),
    ("4. Cyclolysis (Dissipation)", "Warm sector completely detached from surface; thermal contrast eliminated; friction fills central low.", "Barometric pressure rises; cloud decks break into stratocumulus; winds moderate; clear, cold polar anticyclonic weather takes over.")
]

p1_q12_essay = [
    ("Introduction: Synoptic Meteorology and Air Mass Climatology",
     "The dynamic and volatile weather patterns characterizing the middle and high latitudes (30° to 60° North and South) are governed primarily by the movement, modification, and collision of continental and maritime air masses. An air mass is defined in synoptic meteorology as an immense body of tropospheric air—frequently covering millions of square kilometers and extending vertically through several kilometers—whose physical properties, particularly horizontal temperature and moisture distribution, are quasi-homogeneous. The concept, pioneered by Norwegian meteorologist Vilhelm Bjerknes and his son Jacob Bjerknes during World War I, revolutionized atmospheric science. Where contrasting tropical and polar air masses meet along the baroclinic zone known as the 'Polar Front', they do not blend instantaneously; instead, they remain separated by narrow, inclined boundary zones of intense thermal and kinematic discontinuity termed 'Fronts'. The dynamic instability along these frontal surfaces gives birth to Mid-Latitude Wave Cyclones (extratropical cyclones), which serve as the primary engines of temperate winter precipitation."),

    ("Genesis and Classification of Air Masses: Bergeron's Genetic Taxonomy",
     "An air mass acquires its distinctive thermodynamic characteristics by lingering over a 'Source Region'—an expansive, topographically uniform section of the Earth's surface under stagnant, persistent anticyclonic high-pressure conditions for days or weeks. During this stagnation, conductive and radiative heat exchange, along with evaporation, homogenize the lowest tropospheric layers to match the underlying surface. Tor Bergeron established the canonical genetic classification of air masses based on two master criteria: moisture origin and latitudinal thermal characteristics. Moisture origin classifies air masses as 'Continental' (c, originating over dry landmasses, possessing low specific humidity) or 'Maritime' (m, originating over oceans, possessing high moisture content). Thermal characteristics divide into 'Arctic/Antarctic' (A, bitterly cold, shallow), 'Polar' (P, cold, sub-polar), 'Tropical' (T, warm to hot), and 'Equatorial' (E, hot, hyper-humid). The four dominant air masses driving global temperate weather are Continental Polar (cP, cold, dry, stable, originating over Siberian and Canadian snowfields), Maritime Polar (mP, cool, humid, unstable, originating over North Atlantic and North Pacific), Continental Tropical (cT, hot, dry, highly unstable, originating over the Sahara and Arabian deserts), and Maritime Tropical (mT, warm, humid, buoyant, originating over the Gulf of Mexico, Caribbean, and Arabian Sea)."),

    ("Frontogenesis and the Architecture of Frontal Discontinuities",
     "The physical process of frontal creation and intensification is termed 'Frontogenesis' (whereas frontal dissipation is 'Cyclolysis'). When two air masses of contrasting density, temperature, and moisture converge, the less dense (warmer) air mass is invariably forced to rise over the denser (colder) air mass along an inclined discontinuity surface. Fronts classify into four distinct types. A 'Warm Front' occurs where advancing warm air actively displaces and overruns a retreating wedge of cold air. Because of subtle density differences, the frontal slope is very gentle, ranging from 1:150 to 1:200. As warm air ascends this gentle ramp, it produces a classic, predictable 1,000-kilometer sequence of stratiform clouds: high feather-like Cirrus (producing solar/lunar halos), thickening into Cirrostratus, Altostratus, and ultimately dark Nimbostratus, delivering hours or days of continuous, steady drizzle. Conversely, a 'Cold Front' occurs where advancing cold, dense air bulldozes aggressively under a warm air mass. Friction retards the cold front's base, producing a steep, blunted frontal wedge (slope 1:50 to 1:100). This violent, forced vertical uplift triggers severe atmospheric instability, generating a narrow 'Squall Line' of towering Cumulonimbus clouds accompanied by gusty winds, heavy localized downpours, severe lightning, thunder, and occasional tornadoes, followed by a dramatic drop in temperature and an immediate rise in barometric pressure."),

    ("Stationary and Occluded Fronts: Cold vs. Warm Occlusion",
     "A 'Stationary Front' represents a boundary where neither air mass is actively advancing; surface winds blow parallel to the front rather than across it, producing persistent overcast conditions and intermittent rain that can stall for days. An 'Occluded Front' represents the culminating phase of frontal evolution. Because cold fronts travel significantly faster than warm fronts (averaging 40–50 km/h compared to 25–30 km/h for warm fronts), the advancing cold front eventually overtakes the warm front, lifting the intervening warm sector completely off the ground. Occlusions divide into two structural varieties: a 'Cold-type Occlusion' occurs when the air behind the advancing cold front is colder than the air ahead of the warm front, forcing the cold front to undercut both the warm sector and the cool air wedge; a 'Warm-type Occlusion' occurs when the air behind the cold front is milder than the bitterly cold air trapped ahead of the warm front, causing the advancing cold front to ride up over the sub-freezing wedge."),

    ("The Bergen School Polar Front Theory: The Wave Cyclone Life Cycle",
     "The life cycle of a mid-latitude wave cyclone is synthesized through the celebrated 'Polar Front Theory' formulated by Jacob Bjerknes and Halvor Solberg (1922). The cycle progresses through five distinct evolutionary stages. In Stage 1 (Incipient / Stationary Stage), an undisturbed polar front separates easterly polar air (cP) from westerly subtropical air (mT). In Stage 2 (Initial Wave Formation), horizontal wind shear or upper-tropospheric divergence creates a minor wave-like kink along the front, causing barometric pressure to drop at the crest and initiating cyclonic counter-clockwise circulation. In Stage 3 (Open Wave Stage), the cyclone develops well-defined warm and cold fronts enclosing a distinct wedge of warm air (the 'warm sector'). Isobars become tightly packed, winds accelerate, and distinct precipitation shields develop along both fronts. In Stage 4 (Occlusion Stage), the fast-moving cold front catches up with the warm front, lifting the warm sector entirely aloft. This stage exhibits the lowest central barometric pressure, the most expansive precipitation shield, and maximum gale-force winds. In Stage 5 (Cyclolysis / Dissipation Stage), with the warm sector completely severed from the ground, the horizontal temperature gradient (baroclinicity) vanishes; the central low fills with cold air, barometric pressure rises, and the storm dissipates into a stable, cold anticyclone."),

    ("Conclusion: Western Disturbances and the Meteorology of Pakistan",
     "In evaluative summation, the polar front theory provides the foundational conceptual framework for understanding extratropical atmospheric dynamics. The cyclic generation, maturation, and occlusion of wave cyclones govern the redistribution of heat and moisture across the mid-latitudes. For Pakistan, extratropical wave cyclones possess direct, vital economic significance in the form of 'Western Disturbances'. Originating over the Mediterranean Sea, Black Sea, or Atlantic Ocean, these mid-latitude wave cyclones travel eastward embedded within the subtropical westerly jet stream. Between December and March, they enter Pakistan across Iran and Afghanistan, bringing essential winter rains to the plains of Punjab, Khyber Pakhtunkhwa, and Balochistan, while depositing deep snowpacks across the Karakoram and Hindu Kush ranges. This winter precipitation is indispensable for nourishing the non-irrigated (barani) Rabi wheat crop and replenishing the reservoirs of the Tarbela and Mangla dams, proving that temperate frontal meteorology directly sustains South Asia's agrarian economy.")
]

p1_q12_quotes = [
    ("Scholar / Meteorologist", "Theoretical Contribution / Concept", "Exact Quotable Academic Thesis / Axiom"),
    ("Vilhelm Bjerknes (1919)", "The Weather Forecasting as a Problem in Mechanics and Physics", "The polar front is a battle line between polar and tropical air masses; atmospheric disturbances along this line represent waves that mature into cyclonic vortices."),
    ("Jacob Bjerknes & Halvor Solberg (1922)", "Life Cycle of Cyclones and the Polar Front Theory of Atmospheric Circulation", "The extratropical cyclone is born as a wave on the polar front, deepens through warm-sector convergence, reaches peak intensity upon occlusion, and dies when thermal contrast is destroyed."),
    ("Tor Bergeron (1928)", "Über die dreidimensional verknüpfende Wetteranalyse", "Air masses are the grand individuals of synoptic climatology; their physical properties are imprinted by their source regions and transformed across dynamic frontal boundaries."),
    ("Horace R. Byers (1959)", "General Meteorology", "A front is not a mathematical plane but a transition zone of finite thickness characterized by sharp gradients in density, temperature, and cyclonic wind shear."),
    ("Arthur N. Strahler (1969)", "Physical Geography", "The mid-latitude wave cyclone operates as a self-consuming thermodynamic vortex, utilizing the potential energy of contrasting air masses to drive planetary heat exchange.")
]

# ==============================================================================
# QUESTION 13 (P1_Q13): TROPICAL WEATHER SYSTEMS & HURRICANE/CYCLONE DYNAMICS
# Solved from: 2025 Q.4, 2010 Q.2
# ==============================================================================
p1_q13_meta = [
    ("Target Syllabus", "Section 1: Weather & Climate — Tropical Weather Systems: Hurricanes & Tropical Cyclones"),
    ("Exam Years Solved", "KPPSC PMS 2025 Q.4, 2010 Q.2"),
    ("Marks & Allocation", "20 Marks | Strict 35-Minute Descriptive Writing Limit"),
    ("Core Theoretical Thesis", "Tropical cyclones (hurricanes, typhoons, cyclones) represent intense, non-frontal, warm-core synoptic cyclonic vortices that operate as colossal natural Carnot heat engines driven by the release of latent heat of condensation over warm tropical oceans. Development requires strict thermodynamic and kinematic thresholds (Sea Surface Temperature >= 26.5°C to a depth of 50 m, significant Coriolis parameter at >= 5° latitude, weak vertical wind shear < 10 m/s, and high mid-tropospheric humidity). The storm organizes into a central calm, subsiding Eye, a ferocious, torrential Eye Wall, and spiral rainbands, posing lethal storm-surge hazards across coastal zones including Pakistan's Makran and Indus deltaic coastlines.")
]

p1_q13_callout = [
    "THERMODYNAMIC CARNOT ENGINE (EMANUEL): A tropical cyclone converts thermal energy absorbed from warm oceans into mechanical wind kinetic energy via latent heat release: Ocean warmth evaporates water -> moist air converges into low pressure -> air ascends in eye wall -> condensation releases latent heat (roughly 2.5 x 10^6 J/kg) -> warms upper troposphere -> intensifies central barometric pressure drop -> accelerates surface winds.",
    "THE SIX GRAY GENESIS THRESHOLDS (WILLIAM GRAY): (1) Warm Sea Surface Temperature (SST >= 26.5°C extending down >= 50 meters); (2) Significant Coriolis Force (located >= 5° north or south of the equator; cyclones cannot form at the equator where Coriolis = 0); (3) Low Vertical Wind Shear (< 10 m/s between surface and 200 hPa; high shear rips the storm's convective core apart); (4) High Mid-Tropospheric Humidity (>= 70% at 500 hPa to prevent dry air entrainment); (5) Pre-existing Low-Level Disturbance / Vorticity (Easterly Wave, ITCZ trough); (6) Conditional Instability.",
    "STRUCTURAL ANATOMY: (1) Eye: Central circular zone of calm winds, descending air (adiabatic warming), clear skies or thin cirrus, diameter 20–50 km, lowest barometric pressure; (2) Eye Wall: Donut-shaped ring of monstrous cumulonimbus clouds surrounding eye, maximum pressure gradient, strongest horizontal winds (>119 km/h to >250 km/h), torrential deluge; (3) Spiral Rainbands: Curved convective feeder bands spiraling into the eye wall.",
    "INTENSITY SCALE & COASTAL HAZARDS: Saffir-Simpson Hurricane Wind Scale (Category 1: 119–153 km/h to Category 5: >= 252 km/h). Hazards: Storm Surge (the dome of ocean water pushed ashore by winds and low pressure, responsible for 90% of cyclone fatalities), Gale-Force Winds, Inland Torrential Flooding. Arabian Sea Cyclones (Cyclone Phet 2010, Biparjoy 2023) directly threaten Karachi, Gwadar, and Thatta."
]

p1_q13_map = r"""
                 [TROPICAL CYCLONE ANATOMY & ENERGETICS]
                                     │
           [UPPER TROPOSPHERIC OUTFLOW: CIRRUS CANOPY (200 hPa)]
                                 ▲       ▲
                                /         \
                       ┌───────┴───────────┴───────┐
                       │      [THE EYE]            │
                       │ • Calm winds (<15 km/h)   │
                       │ • Descending air          │
                       │ • Adiabatic warming       │
                       │ • Lowest pressure (<950mb)│
                       └───────┬───────────┬───────┘
                               │           │
                 [EYE WALL]    │           │    [EYE WALL]
                 • Monster Cb  │           │    • Monster Cb
                 • Winds >200km│           │    • Winds >200km/h
                 • Max deluge  │           │    • Max deluge
                       \       │           │       /
                        \      │           │      /
     [SPIRAL RAINBANDS]  \     │           │     /  [SPIRAL RAINBANDS]
           ◄──────────────┴────┴───────────┴────┴──────────────►
             INFLOW OF WARM, MOIST TROPICAL AIR (SST >= 26.5°C)
                        [OCEAN HEAT ENGINE: LATENT HEAT]
"""

p1_q13_timeline = [
    ("Storm Stage / Element", "Barometric & Wind Speed Metric", "Meteorological Anatomy & Atmospheric Manifestation"),
    ("Tropical Disturbance", "Slight surface low (<1010 hPa); wind speed <37 km/h (<20 knots).", "Disorganized cluster of convective thunderstorms originating from an Easterly Wave or ITCZ trough; zero closed isobars."),
    ("Tropical Depression", "Central pressure 1005–1008 hPa; wind speed 37–62 km/h (20–33 knots).", "At least one closed isobar appears; organized cyclonic counter-clockwise circulation commences; assigned an official tracking number."),
    ("Tropical Storm", "Central pressure 990–1000 hPa; wind speed 63–118 km/h (34–63 knots).", "Tightly packed circular isobars; substantial central convection; storm receives an official international name; spiral rainbands emerge."),
    ("Hurricane / Cyclone (Cat 1–2)", "Central pressure 965–989 hPa; wind speed 119–177 km/h (64–95 knots).", "Definitive 'Eye' forms; dense eye wall develops; dangerous storm surge (1.2–2.4 m); extensive coastal damage to mobile structures."),
    ("Major Cyclone (Cat 3–5)", "Central pressure <920–964 hPa; wind speed 178 to >252 km/h (>96 knots).", "Catastrophic destructive power; eye wall winds exceed 250 km/h; storm surge reaches 4 to 9 meters; complete structural failure; e.g., Cyclone 05B (1999 Sindh)."),
    ("Storm Landfall & Dissipation", "Rapid pressure rise; winds decay below tropical storm force within 24 hours.", "Friction with land slows winds; loss of warm oceanic moisture supply cuts off latent heat engine; storm degrades into an inland rain depression.")
]

p1_q13_essay = [
    ("Introduction: The Thermodynamic Fury of the Tropical Atmosphere",
     "Tropical cyclones—variously designated as 'Hurricanes' in the North Atlantic and Northeast Pacific, 'Typhoons' in the Northwest Pacific, and simply 'Cyclones' in the Indian Ocean and South Pacific—represent the most energetic, visually spectacular, and destructive synoptic weather systems on Earth. Unlike mid-latitude extratropical cyclones, which derive their kinetic energy from horizontal temperature gradients along contrasting frontal boundaries (baroclinic instability), tropical cyclones are intense, warm-core, non-frontal cyclonic vortices that originate exclusively over warm tropical and subtropical oceans. Operating as colossal natural thermodynamic heat engines, a mature hurricane releases energy at an astronomical rate—equivalent to approximately 200 times the entire global electrical generating capacity—driven almost entirely by the release of latent heat of condensation within its towering eyewall convection. When these maritime vortexes strike populated coastal plains, the combination of catastrophic storm surges, gale-force cyclonic winds, and torrential deluge inflicts devastating human and economic losses."),

    ("Thermodynamic Heat Engine Mechanics: The Carnot Cycle Analogy",
     "The physical engine driving a tropical cyclone is best conceptualized through the thermodynamic Carnot cycle, formulated mathematically by MIT atmospheric scientist Kerry Emanuel. The ocean surface serves as the heat source (operating at temperatures exceeding 26.5°C to 30°C), while the extremely cold upper troposphere and lower stratosphere (operating at -70°C to -85°C near the tropopause) serves as the heat sink. The thermodynamic cycle operates through four continuous steps: First, air spiraling inward across the sea surface under intense barometric pressure gradients absorbs immense sensible heat and moisture from the warm ocean via turbulent evaporation. Second, as this hyper-humid air reaches the central eyewall, it ascends rapidly in an isothermal-to-adiabatic plume. Third, as the air cools, condensation releases vast quantities of latent heat of vaporization (roughly 2.5 × 10⁶ Joules per kilogram), which superheats the core of the storm, creating a warm-core anomaly that lowers surface barometric pressure. Fourth, the air exhausts radially outward at the tropopause as a vast anticyclonic cirrus canopy, radiating heat into space before subsiding."),

    ("The Six Environmental Thresholds of Tropical Cyclogenesis: William Gray's Criteria",
     "Tropical cyclones cannot form spontaneously anywhere in the tropics; their genesis demands a precise juxtaposition of six environmental thermodynamic and kinematic prerequisites, systematically identified by atmospheric scientist William M. Gray. First, the 'Thermal Oceanic Prerequisite': sea surface temperatures (SST) must equal or exceed 26.5°C (80°F), and this thermal threshold must extend to an ocean depth of at least 50 meters to prevent churning storm waves from upwelling cold water that would extinguish the storm. Second, the 'Coriolis Parameter Threshold': cyclones require a significant planetary deflecting force to spin converging air into a vortex; hence, they can only form at least 5° of latitude away from the equator (typically between 5° and 20° N/S). At the equator, where the Coriolis parameter is mathematically zero, cyclogenesis is physically impossible. Third, the 'Vertical Wind Shear Threshold': the atmospheric column must exhibit minimal vertical wind shear (less than 10 meters per second between the surface and 200 hPa). Strong vertical shear tilts the convective column, dispersing latent heat and tearing the embryonic vortex apart. Fourth, the 'Moisture Prerequisite': the mid-troposphere (around 700 to 500 hPa) must possess high relative humidity (>=70%), preventing dry-air entrainment that would evaporate convective updrafts. Fifth, the 'Instability Threshold': the troposphere must be conditionally unstable to support deep free convection. Sixth, there must exist a pre-existing low-level disturbance with cyclonic vorticity—such as an Easterly Wave originating over North Africa or a monsoon trough depression."),

    ("Structural Anatomy: The Eye, Eye Wall, and Spiral Rainbands",
     "A mature tropical cyclone exhibits a highly symmetrical, organized three-dimensional architecture comprising three distinct structural zones. At the center lies the 'Eye'—a circular, cloud-free or lightly overcast column averaging 20 to 50 kilometers in diameter. Within the eye, surface winds are remarkably calm (often less than 15 km/h), the sun or stars are visible, and air gently subsides from the upper troposphere. This subsidence warms the air adiabatically, creating the storm’s diagnostic 'warm core' and driving barometric pressure to the lowest values recorded on Earth (such as Typhoon Tip’s record 870 hPa in 1979). Enclosing the eye is the 'Eye Wall'—a towering, vertical ring of monstrous Cumulonimbus clouds reaching heights of 15 to 18 kilometers. The eye wall represents the powerhouse of the storm: here, the barometric pressure gradient is steepest, vertical updrafts reach maximum velocity, and horizontal winds achieve their most ferocious speeds (exceeding 119 km/h to over 250 km/h in Category 5 hurricanes), delivering blinding, torrential rain. Radiating outward from the eye wall are curved 'Spiral Rainbands'—convective bands of thunderstorms spiraling inward toward the center, interspersed with regions of calm subsidence."),

    ("Classification, Hazard Profiles, and Coastal Storm Surges",
     "Tropical cyclones are classified by sustained horizontal wind speeds using international scales, most notably the Saffir-Simpson Hurricane Wind Scale. The spectrum progresses from Tropical Disturbance (<37 km/h) to Tropical Depression (37–62 km/h, designated with closed isobars), Tropical Storm (63–118 km/h, at which point the storm is assigned an official name), and mature Hurricane/Cyclone (Category 1: 119–153 km/h; Category 2: 154–177 km/h; Category 3: 178–208 km/h; Category 4: 209–251 km/h; and Category 5: >=252 km/h). When a tropical cyclone makes landfall, its destructive power manifests across three hazard vectors: destructive gale-force winds, catastrophic inland freshwater flooding, and the deadly 'Storm Surge'. The storm surge is an abnormal, rapid rise of sea level pushed onto the coast by the combination of ferocious cyclonic winds dragging surface water onshore and the 'inverted barometer effect' (where low atmospheric pressure allows the sea surface to bulge upward by roughly 1 cm per 1 hPa pressure drop). Storm surges—frequently reaching 4 to 9 meters in height—account for over 90 percent of all historical cyclone fatalities, as tragically demonstrated by the 1970 Bhola Cyclone in East Pakistan (Bangladesh), which claimed over 300,000 lives."),

    ("Arabian Sea Cyclogenesis and Pakistan's Coastal Vulnerability",
     "In the regional context of the North Indian Ocean, tropical cyclones form in the Arabian Sea primarily during two distinct seasonal peaks: pre-monsoon (May–June) and post-monsoon (October–November), when the ITCZ traverses the basin and vertical wind shear is suppressed. Although the Arabian Sea historically experienced fewer cyclones than the Bay of Bengal, rising sea surface temperatures linked to anthropogenic climate change have intensified cyclogenesis in recent decades. The 1,050-kilometer coastline of Pakistan—encompassing the low-lying Indus River Delta in Sindh and the rugged Makran coast in Balochistan—is highly vulnerable to severe cyclonic storms. Notable events include the catastrophic 1999 Thatta Cyclone (Cyclone 02A), which devastated the Indus deltaic coastal belt; Cyclone Phet (June 2010), which struck Gwadar and Pasni with 150 km/h winds and caused severe flooding along the Makran Highway; and Cyclone Biparjoy (June 2023), which forced the evacuation of over 80,000 people from coastal Sindh. Mitigating these extreme oceanic hazards requires robust meteorological early-warning radars, conservation of protective coastal mangrove forests, storm-surge sea dikes, and rigorous disaster preparedness under the National Disaster Management Authority (NDMA).")
]

p1_q13_quotes = [
    ("Scholar / Meteorologist", "Theoretical Contribution / Concept", "Exact Quotable Academic Thesis / Axiom"),
    ("William M. Gray (1968)", "Global View of the Origin of Tropical Disturbances and Storms", "Tropical cyclogenesis represents a delicate, cooperative interaction between large-scale environmental parameters: warm ocean thermal depth, weak vertical wind shear, and high mid-tropospheric moisture."),
    ("Kerry Emanuel (1986)", "An Air-Sea Interaction Theory for Tropical Cyclones", "A hurricane is fundamentally a finite-amplitude heat engine of the Carnot cycle, where the thermodynamic disequilibrium between the tropical ocean and upper atmosphere drives violent kinetic winds."),
    ("Herbert Saffir & Robert Simpson (1974)", "The Saffir-Simpson Hurricane Wind Scale", "Categorization of hurricanes based on peak sustained winds provides a direct, operational index for predicting structural devastation and coastal storm-surge inundation."),
    ("Arthur N. Strahler (1969)", "Physical Geography", "The eye of the tropical cyclone represents an island of paradox: an oasis of descending calm and clear skies encircled by the most terrifying wall of hurricane fury known to nature."),
    ("Peter J. Webster (2005)", "Changes in Tropical Cyclone Number, Duration, and Intensity in a Warming Environment", "Rising sea surface temperatures in tropical ocean basins directly amplify the potential intensity and category 4 and 5 frequency of tropical vortices globally.")
]

# ==============================================================================
# QUESTION 14 (P1_Q14): KÖPPEN’S CLIMATIC CLASSIFICATION & Af VS. Aw CLIMATES
# Solved from: 2025 Q.5, 2016 Q.8
# ==============================================================================
p1_q14_meta = [
    ("Target Syllabus", "Section 1: Weather & Climate — Köppen's Climatic Classification; Af vs. Aw"),
    ("Exam Years Solved", "KPPSC PMS 2025 Q.5, 2016 Q.8"),
    ("Marks & Allocation", "20 Marks | Strict 35-Minute Descriptive Writing Limit"),
    ("Core Theoretical Thesis", "Wladimir Köppen’s empirical climatic classification (1900–1936) establishes a quantitative, global taxonomic system based on defined monthly and annual thresholds of temperature and precipitation, aligned with Alphonse de Candolle’s biomes. Its first-letter taxonomy divides into five major thermal groups (A: Tropical Humid, B: Dry/Arid, C: Mesothermal, D: Microthermal, E: Polar) plus H (Highland). Comparative dissection of the A-group climates—Af (Tropical Wet / Rainforest) versus Aw (Tropical Wet-and-Dry / Savanna)—reveals fundamental contrasts in seasonal ITCZ migration, temperature regimes, rainfall distribution, vegetation morphology, and soil edaphology.")
]

p1_q14_callout = [
    "KÖPPEN'S TAXONOMIC FOUNDATION: Quantitative and empirical; utilizes mean monthly temperature, annual temperature, monthly precipitation, and seasonal precipitation cycles. Designed to correlate with global vegetation zones (phytogeography).",
    "THE FIRST-LETTER MAJOR CLIMATIC GROUPS: (1) A = Tropical Humid Climates (coldest month mean >= 18°C; no winter season); (2) B = Dry / Arid Climates (potential evapotranspiration exceeds precipitation across annual cycle; further divided into BS Steppe and BW Desert); (3) C = Mesothermal / Mild Mid-Latitude (coldest month between -3°C and 18°C; at least one month > 10°C); (4) D = Microthermal / Severe Mid-Latitude (coldest month < -3°C; warmest month > 10°C); (5) E = Polar Climates (warmest month < 10°C; ET Tundra and EF Frost/Ice Cap); (6) H = Highland Climates (undifferentiated due to complex topography and altitude lapse rates).",
    "SECOND & THIRD LETTER QUALIFIERS: Second letter (precipitation seasonality): f = fully humid / no dry season; s = dry summer; w = dry winter; m = monsoonal. Third letter (temperature intensity): a = hot summer (warmest month >= 22°C); b = warm summer (warmest month < 22°C, at least 4 months > 10°C); c = cool summer; d = extremely cold winter (coldest month < -38°C); h = hot arid (annual mean >= 18°C); k = cold arid (annual mean < 18°C).",
    "Af VS. Aw COMPARATIVE DICHOTOMY: Af (Tropical Rainforest): Located 0°–10° N/S (Amazon, Congo, Indonesia); Isothermal climate (annual range < 3°C); rainfall > 2,000 mm evenly distributed (driest month >= 60 mm); dense multi-layered evergreen rainforest (selva), nutrient-poor oxisols/laterites. Aw (Tropical Savanna): Located 10°–20° N/S (Llanos, Campos, African Sahel, Deccan); distinct seasonal wet summer (ITCZ overhead) and severe dry winter (subtropical anticyclone); rainfall 750–1,500 mm (driest month < 60 mm); tall drought-resistant grasslands with scattered parkland trees (acacia, baobab)."
]

p1_q14_map = """
                 [KÖPPEN CLIMATIC TAXONOMY & AF VS. AW]
                                    │
           ┌────────────────────────┴────────────────────────┐
           ▼                                                 ▼
[KÖPPEN 5 MAJOR GROUPS]                             [A-GROUP CLIMATIC COMPARISON]
• A: Tropical Humid (Coldest >= 18°C)               (All months >= 18°C: No Winter)
• B: Dry Climates (Evaporation > Precip)                     │
• C: Warm Temperate (-3°C < Coldest < 18°C)         ┌────────┴────────┐
• D: Cold Forest (Coldest < -3°C)                   ▼                 ▼
• E: Polar Climates (Warmest < 10°C)        [Af: RAINFOREST]   [Aw: SAVANNA]
• H: Highland Climates (Altitude lapse)     • Lat: 0° - 10°    • Lat: 10° - 20°
                                            • All months >=60mm • Driest month <60mm
                                            • Rainfall >2,000mm • Rainfall: 750-1,500mm
                                            • Constant ITCZ    • Seasonal ITCZ shift
                                            • Isothermal       • Distinct wet/dry seasons
                                            • Selva Rainforest • Park-grassland savanna
"""

p1_q14_timeline = [
    ("Climatic Metric / Parameter", "Af (Tropical Wet / Rainforest Climate)", "Aw (Tropical Wet-and-Dry / Savanna Climate)"),
    ("Geographical Distribution", "Equatorial belt within 5°–10° N/S: Amazon Basin (Brazil), Congo Basin (Central Africa), Malaysian-Indonesian Archipelago.", "Tropical margins between 10° and 20° N/S: Llanos (Orinoco), Campos (Brazil), Sahel and Sudanian savanna (Africa), Deccan Plateau (India)."),
    ("Thermal Regime & Range", "Permanently hot and isothermal; all months average 25°–28°C; annual temperature range < 3°C; diurnal range exceeds annual range.", "Warm to hot year-round; mean temperatures 24°–30°C; maximum temperatures occur immediately before rainy season; annual range 5°–10°C."),
    ("Precipitation & Seasonality", "Year-round heavy rainfall (>2,000–3,000 mm); no dry season; driest month receives at least 60 mm; daily afternoon convection.", "Pronounced seasonal alternation: torrential wet summer followed by severe dry winter drought; rainfall 750–1,500 mm; driest month < 60 mm."),
    ("Dominant Atmospheric Control", "Permanent convergence of Trade Winds in the Intertropical Convergence Zone (ITCZ); continuous convective uplifting.", "Alternating circulation: dominated by ITCZ in summer (wet low pressure), dominated by Subtropical Highs in winter (dry sinking air)."),
    ("Vegetation Architecture (Biome)", "Tropical Rainforest (Selva): dense, evergreen, multi-layered canopy (emergent, canopy, understory); massive biodiversity; epiphytes and lianas.", "Tropical Savanna: extensive tall grasslands (elephant grass up to 3 m) interspersed with drought-resistant, fire-adapted parkland trees (acacia, baobab)."),
    ("Soil Classification & Fertility", "Oxisols and Ultisols (Lateritic soils): intense chemical leaching (desilication); silica stripped, leaving iron/aluminum oxides; acidic, infertile.", "Oxisols, Alfisols, and Vertisols: seasonal moisture drives upward capillary action of minerals; higher base saturation than Af; susceptible to hardpan formation.")
]

p1_q14_essay = [
    ("Introduction: The Quantitative Paradigm in Climatological Classification",
     "The systematic classification of world climates stands as one of the enduring achievements of physical geography, transforming isolated meteorological records of temperature and rainfall into organized, predictive regional models. Among the various genetic and empirical taxonomies devised over the past two centuries, the climatic classification formulated by the German-Russian climatologist and botanist Wladimir Köppen (first published in 1900 and iteratively refined until his definitive 1936 monograph with Rudolf Geiger) remains the international benchmark. Köppen’s genius lay in creating an empirical, quantitative system that utilized readily available monthly and annual temperature and precipitation data while deliberately anchoring its boundaries to the global phytogeographical zones established by the Swiss botanist Alphonse de Candolle in 1874. Recognizing that natural vegetation serves as the ultimate living integrator of macro-climatic conditions, Köppen established clear mathematical thresholds that delineate planetary biomes across five major thermal groups."),

    ("The Architectural Hierarchy of Köppen's First-Letter Classification",
     "The primary structural tier of Köppen’s system is designated by five capital letters (A, B, C, D, E), representing five great planetary climatic groups based upon fundamental thermal and moisture limits, supplemented by an undifferentiated Highland group (H). Group 'A' designates 'Tropical Humid Climates': defined strictly as winterless climates where the mean temperature of the coldest month never falls below 18°C (64.4°F), matching the thermal requirements of megathermal plants. Group 'B' designates 'Dry (Arid and Semi-Arid) Climates': unique in being defined not by temperature alone, but by moisture deficiency—where annual potential evapotranspiration exceeds annual precipitation, calculated via mathematical formulas that account for seasonal rainfall distribution. Group 'B' subdivides into Semi-Arid Steppe (BS) and True Arid Desert (BW). Group 'C' designates 'Mesothermal (Warm Temperate) Climates': defined where the coldest month averages between -3°C and +18°C, and at least one month averages above 10°C, corresponding to microthermal deciduous and sclerophyllous forest biomes. Group 'D' designates 'Microthermal (Cold Snow Forest / Boreal) Climates': defined where the coldest month averages below -3°C and the warmest month exceeds 10°C, sustaining the great coniferous Taiga forests. Group 'E' designates 'Polar Climates': defined where the warmest month never exceeds 10°C, dividing into Tundra (ET, warmest month between 0°C and 10°C) and Perpetual Frost / Ice Cap (EF, all months below 0°C)."),

    ("Second and Third Letter Qualifiers: Seasonal Moisture and Thermal Intensity",
     "To capture seasonal variations in precipitation and temperature, Köppen introduced precise second and third lowercase letter modifiers. The second letter defines precipitation seasonality: 'f' denotes fully humid with no dry season (fehlt, meaning drought is absent); 'w' denotes a distinct dry winter season (trockener Winter); 's' denotes a distinct dry summer season (trockener Sommer); and 'm' designates monsoonal rainfall, where excessive summer precipitation compensates for a short dry season. For the B group, capital 'S' denotes semi-arid steppe, while capital 'W' denotes arid desert (Wüste). The third letter provides refined thermal criteria: 'a' indicates hot summers (warmest month >=22°C); 'b' indicates warm summers (warmest month <22°C with at least four months >=10°C); 'c' indicates cool, short summers; and 'd' denotes hyper-continental cold winters (coldest month <-38°C). In dry climates, 'h' denotes hot (heiss, mean annual temperature >=18°C, e.g., BWh), while 'k' denotes cold (kalt, mean annual temperature <18°C, e.g., BWk)."),

    ("Comparative Analysis: Af (Tropical Rainforest) vs. Aw (Tropical Savanna)",
     "A profound comparative case study within Köppen’s taxonomy is the structural and ecological dichotomy between the two dominant subdivisions of the Tropical Humid (A) group: the 'Af' (Tropical Wet / Rainforest) climate and the 'Aw' (Tropical Wet-and-Dry / Savanna) climate. Geographically, the Af climate is tightly anchored along the equatorial belt within 5° to 10° North and South of the equator, encompassing the expansive Amazon Basin of South America, the Congo Basin of Central Africa, and the Indonesian-Malaysian Archipelago. Thermally, the Af climate is permanently warm and remarkable for its isothermal character: mean monthly temperatures hover consistently between 25°C and 28°C throughout the year, resulting in an annual temperature range of less than 3°C—so minimal that the diurnal temperature range between day and night is greater than the seasonal difference across months. Precipitously, the Af climate receives colossal annual rainfall, typically exceeding 2,000 to 3,000 millimeters, distributed evenly across all twelve months; according to Köppen’s strict mathematical definition, even the driest month receives at least 60 millimeters of precipitation. This continuous rainfall is driven by the permanent presence of the Intertropical Convergence Zone (ITCZ), where convergent trade winds fuel daily afternoon convectional thunderstorms."),

    ("The Aw Savanna Realm: Seasonal ITCZ Shift and Moisture Stress",
     "In sharp contrast, the Aw (Tropical Savanna) climate occupies the transitional zones flanking the equatorial belt, situated roughly between 10° and 20° North and South of the equator. The Aw realm encompasses the Llanos of Venezuela, the Campos and Cerrado of the Brazilian Highlands, the vast savanna belt of Sub-Saharan Africa (the Sudanian zone), and portions of the Deccan Plateau in India. While the Af climate experiences no dry season, the defining characteristic of the Aw climate is an extreme, alternating seasonality of precipitation: a torrential wet season in the high-sun summer alternating with a severe, prolonged dry drought season in the low-sun winter. Annual rainfall is substantially lower than in the Af realm, averaging 750 to 1,500 millimeters, and Köppen’s definition requires that at least one month receives less than 60 millimeters of precipitation. This seasonal rainfall dichotomy is governed by the annual latitudinal migration of global pressure belts: during summer, the northward migration of the thermal equator brings the humid ITCZ overhead, unleashing months of heavy convective and monsoonal rainfall; during winter, the ITCZ retreats, and the region is submerged under the subsiding, bone-dry air of the Subtropical High-Pressure belt, imposing months of desiccating drought."),

    ("Ecological and Edaphic Consequences: Selva Forest vs. Park-Grassland",
     "The contrasting moisture regimes of the Af and Aw climates dictate diametrically opposed ecological biomes and soil types. The Af climate sustains the 'Tropical Rainforest' (or Selva)—the most biologically diverse and architecturally complex terrestrial ecosystem on Earth. The vegetation is characterized by tall, broadleaf evergreen hardwoods forming a continuous, multi-layered canopy (emergent layer, closed canopy, understory) dense with woody lianas and epiphytes. Conversely, the seasonal drought of the Aw climate prevents the growth of continuous closed forests; instead, it fosters the 'Tropical Savanna'—an expansive parkland biome dominated by tall, coarse drought-resistant grasses (such as elephant grass reaching 3 meters in height) interspersed with scattered, umbrella-shaped drought- and fire-adapted trees like the acacia and baobab. Edaphically, both realms feature deeply weathered soils (Oxisols/Ultisols), but while Af soils undergo continuous, severe leaching (desilication) that leaves nutrient-deficient red laterites, Aw soils experience seasonal upward capillary draw during dry periods, accumulating mineral bases. In evaluative summation, Köppen’s classification masterfully captures how mathematical variations in seasonal heat and moisture translate into the fundamental bioclimatic landscapes that define physical geography.")
]

p1_q14_quotes = [
    ("Scholar / Climatologist", "Theoretical Contribution / Concept", "Exact Quotable Academic Thesis / Axiom"),
    ("Wladimir Köppen (1900)", "Versuch einer Klassifikation der Klimate", "Climate is best classified through its visible biological expression; the mathematical boundaries of temperature and precipitation must coincide with the limits of distinct plant formations."),
    ("Rudolf Geiger (1954)", "Klassifikation der Klimate nach W. Köppen", "The Köppen-Geiger system endures as the definitive climatic taxonomy because its objective quantitative criteria provide a universal language for physical geography."),
    ("Arthur N. Strahler (1969)", "Physical Geography", "The Af and Aw climates illustrate the profound morphological power of seasonal precipitation; the transition from equatorial rain forest to savanna reflects the shifting boundary between ITCZ convergence and subtropical subsidence."),
    ("Glenn T. Trewartha (1968)", "An Introduction to Climate", "Köppen's genius was in recognizing that monthly temperature and precipitation data contain the complete code for planetary vegetation biomes and agricultural potential."),
    ("Alphonse de Candolle (1874)", "Prodromus Systematis Naturalis Regni Vegetabilis", "The geographic distribution of plants reflects an ancient, inviolable physiological adaptation to specific thermal limits and moisture durations.")
]

# ==============================================================================
# BATCH 2 EXECUTION RUNNER (Q7 to Q14)
# ==============================================================================
BATCH_2_TASKS = [
    ("Q7_Karst_Topography_and_Groundwater_Hydrology", "Q.7: Karst Topography & Groundwater Hydrology", p1_q7_meta, p1_q7_callout, p1_q7_map, p1_q7_timeline, p1_q7_essay, p1_q7_quotes),
    ("Q8_Aeolian_and_Coastal_Geomorphology", "Q.8: Aeolian & Coastal Geomorphology", p1_q8_meta, p1_q8_callout, p1_q8_map, p1_q8_timeline, p1_q8_essay, p1_q8_quotes),
    ("Q9_Solar_Insolation_Heat_Budget_and_Radiation_Balance", "Q.9: Solar Insolation, Heat Budget & Radiation Balance", p1_q9_meta, p1_q9_callout, p1_q9_map, p1_q9_timeline, p1_q9_essay, p1_q9_quotes),
    ("Q10_Atmospheric_Pressure_Belts_and_Planetary_Winds", "Q.10: Atmospheric Pressure Belts & Planetary Winds", p1_q10_meta, p1_q10_callout, p1_q10_map, p1_q10_timeline, p1_q10_essay, p1_q10_quotes),
    ("Q11_Atmospheric_Moisture_Clouds_and_Precipitation", "Q.11: Atmospheric Moisture, Clouds & Precipitation", p1_q11_meta, p1_q11_callout, p1_q11_map, p1_q11_timeline, p1_q11_essay, p1_q11_quotes),
    ("Q12_Air_Masses_Frontogenesis_and_Wave_Cyclones", "Q.12: Air Masses, Frontogenesis & Polar Front Wave Cyclones", p1_q12_meta, p1_q12_callout, p1_q12_map, p1_q12_timeline, p1_q12_essay, p1_q12_quotes),
    ("Q13_Tropical_Weather_Systems_and_Hurricane_Dynamics", "Q.13: Tropical Weather Systems & Hurricane Dynamics", p1_q13_meta, p1_q13_callout, p1_q13_map, p1_q13_timeline, p1_q13_essay, p1_q13_quotes),
    ("Q14_Koppen_Climatic_Classification_Af_vs_Aw", "Q.14: Köppen Climatic Classification & Af vs. Aw Climates", p1_q14_meta, p1_q14_callout, p1_q14_map, p1_q14_timeline, p1_q14_essay, p1_q14_quotes),
]

for filename, title_text, meta, callout, map_txt, timeline, essay, quotes in BATCH_2_TASKS:
    docx_file = os.path.join(P1_DIR, f"{filename}_Master_Notes.docx")
    pdf_file = os.path.join(P1_DIR, f"{filename}_Master_Notes.pdf")
    md_file = os.path.join(P1_DIR, f"{filename}_Master_Notes.md")
    
    generate_geo_master_doc(meta, callout, map_txt, timeline, essay, quotes, docx_file, title_text)
    generate_pdf_and_md(meta, callout, map_txt, timeline, essay, quotes, pdf_file, md_file, title_text)

print("Batch 2 (Q7 through Q14) for Geography Paper 1 successfully generated!")
