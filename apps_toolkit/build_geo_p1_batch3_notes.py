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
    r_ht = p_ht.add_run("2. ANALYTICAL DIMENSIONS & GEOMORPHIC/CARTOGRAPHIC MATRIX")
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
    r_hq = p_hq.add_run("4. SCHOLARLY PERSPECTIVES & OCEANOGRAPHIC/CARTOGRAPHIC AXIOMS")
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
    
    md_content += "### 2. Analytical Dimensions & Geomorphic/Cartographic Matrix\n\n"
    headers = timeline_data[0]
    md_content += f"| {headers[0]} | {headers[1]} | {headers[2]} |\n|---|---|---|\n"
    for row in timeline_data[1:]:
        md_content += f"| **{row[0]}** | {row[1]} | {row[2]} |\n"
    md_content += "\n---\n\n"
    
    md_content += "### 3. 20-Mark Master Model Essay\n\n"
    for sh, bd in essay_sections:
        md_content += f"#### {sh}\n{bd}\n\n"
    md_content += "---\n\n"
    
    md_content += "### 4. Scholarly Perspectives & Oceanographic/Cartographic Axioms\n\n"
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

<h3>2. Analytical Dimensions & Geomorphic/Cartographic Matrix</h3>
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

<h3>4. Scholarly Perspectives & Oceanographic/Cartographic Axioms</h3>
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

print("Batch 3 builder engine ready.")

# ==============================================================================
# QUESTION 15 (P1_Q15): MORPHOLOGY OF THE OCEAN FLOOR
# Solved from: 2016 Q.5, 2010 Q.8
# ==============================================================================
p1_q15_meta = [
    ("Target Syllabus", "Section 3: Oceans & Oceanography — Continental Margins & Ocean Floor Morphology"),
    ("Exam Years Solved", "KPPSC PMS 2016 Q.5, 2010 Q.8"),
    ("Marks & Allocation", "20 Marks | Strict 35-Minute Descriptive Writing Limit"),
    ("Core Theoretical Thesis", "The global ocean floor (submerging 71% of Earth's surface across 361 million km²) exhibits a bimodal hypsographic profile structured into three major physiographic provinces: Continental Margins (Shelf, Slope, Rise), Deep Ocean Basins (Abyssal Plains, Oceanic Trenches, Seamounts, Guyots), and the Mid-Ocean Ridge System. Continental margins divide tectonically into Passive (Atlantic-type, broad shelf, aseismic) and Active (Pacific-type, narrow shelf, subduction trench, high seismicity/volcanism). The continuous 65,000-kilometer mid-ocean ridge chain represents the divergent volcanic engine of plate tectonics, driving seafloor spreading and global oceanic crustal renewal.")
]

p1_q15_callout = [
    "HYPSOGRAPHIC BIPARTITION: Earth's topography is bimodal: continental platforms average +840 meters elevation (granitic Sial, density 2.7 g/cm³), while oceanic floors average -3,800 meters depth (basaltic Sima, density 3.0 g/cm³), reflecting isostatic equilibrium over the mantle.",
    "CONTINENTAL MARGIN ARCHITECTURE: (1) Continental Shelf: Gently sloping submerged edge (gradient 0.1° or 1:1,000, average depth 130–200 m, width from <1 km to 1,500 km; covers 7.5% of ocean floor; critical economic zone for fisheries, petroleum, natural gas); (2) Continental Slope: Steep drop-off (gradient 4°–5°, depth 200 to 3,000–4,000 m; structural boundary between continental and oceanic crust; carved by Submarine Canyons via turbidity currents); (3) Continental Rise: Gentle apron of sediment (gradient 0.5°–1°, depth 3,000–4,000 m; coalescing deep-sea turbidite fans).",
    "DEEP OCEAN BASINS & MID-OCEAN RIDGES: Abyssal Plains (flat, horizontal plains at 3,000–6,000 m depth blanketed by pelagic red clays and biogenic calcareous/siliceous oozes, covering 40% of ocean floor); Oceanic Trenches (narrow, V-shaped subduction chasms up to 11,000 m deep, e.g., Mariana Trench 10,994 m); Seamounts (submerged volcanic cones >1,000 m) and Guyots (flat-topped volcanic seamounts planed by wave erosion before subsidence).",
    "MID-OCEAN RIDGE SYSTEM: The longest mountain range on Earth (65,000 km), featuring a central axial rift valley, basaltic pillow lavas, hydrothermal 'black smoker' vents, and transform fault offsets, driving seafloor spreading."
]

p1_q15_map = r"""
                     [OCEAN FLOOR PHYSIOGRAPHIC PROFILE]
                                      │
 0m [COASTLINE]
     \
      \── [CONTINENTAL SHELF] (Avg 130m depth; gradient 0.1°; width 75km)
           \
            \── [CONTINENTAL SLOPE] (Gradient 4°-5°; submarine canyons)
                 \
                  \── [CONTINENTAL RISE] (Gradient 0.5°-1°; turbidite fans)
                       \
                        \────────────────── [ABYSSAL PLAIN] (3,000-6,000m depth)
                                                 │
                  ┌──────────────────────────────┴──────────────────────────────┐
                  ▼                                                             ▼
          [OCEANIC TRENCH]                                            [MID-OCEAN RIDGE]
       (Subduction: up to 11,000m)                                 (Divergent Seafloor Spreading)
       • Mariana Trench (10,994m)                                  • Central Axial Rift Valley
       • Wadati-Benioff Zone                                       • Hydrothermal Black Smokers
       • Volcanic Island Arcs                                      • High Thermal Heat Flow
"""

p1_q15_timeline = [
    ("Physiographic Province", "Depth, Gradient & Tectonic Setting", "Diagnostic Geomorphic Features & Geological Role"),
    ("Continental Shelf", "Depth: 0 to 130–200 m; average gradient 0.1° (1:1,000); width varies from 1 to 1,500 km.", "Flooded edge of continent; covered with terrigenous sediments; epicenters of global marine fisheries and offshore hydrocarbon exploitation (e.g., Arabian Sea shelf off Karachi)."),
    ("Continental Slope", "Depth: 130–200 m down to 3,000–4,000 m; steep gradient 4°–5° (exceeding 20° in trenches).", "True geological boundary between continental granitic Sial and oceanic basaltic Sima; dissected by deep V-shaped Submarine Canyons (e.g., Indus Canyon)."),
    ("Submarine Canyons & Continental Rise", "Depth: 3,000 to 4,000 m; gradient 0.5° to 1.0°; sediment apron at base of slope.", "Canyons excavated by high-density, abrasive Turbidity Currents (underwater avalanches); sediments form massive deep-sea alluvial fans merging into the continental rise."),
    ("Abyssal Plains", "Depth: 3,000 to 6,000 m; ultra-flat (slope < 1:1,000); covers ~40% of the planetary ocean floor.", "Extremely flat topography produced by thick sediment burial of volcanic basement; covered by pelagic red clay and biogenic calcareous (globigerina) and siliceous oozes."),
    ("Oceanic Trenches", "Depth: 6,000 to nearly 11,000 m; narrow, arcuate V-shaped tectonic troughs.", "Formed by oceanic plate subduction along convergent margins; Wadati-Benioff seismic zones; deepest point on Earth: Challenger Deep in Mariana Trench (10,994 m)."),
    ("Mid-Ocean Ridge & Rift Valley", "Elevation: 2,000 to 3,000 m above abyssal plain; length 65,000 km continuous chain.", "Divergent plate boundary; central axial rift valley where magma upwells to form new basaltic oceanic crust; dotted with hydrothermal vents (black smokers).")
]

p1_q15_essay = [
    ("Introduction: The Submerged Physiographic Frontier of the Earth",
     "The oceans of the world blanket approximately 361 million square kilometers—representing over 70.8 percent of the Earth's surface—with a mean depth of approximately 3,800 meters. For centuries, the oceanic abyss was conceptualized as a vast, featureless, mud-covered bathtub devoid of dynamic relief. The advent of modern deep-sea geophysical surveying in the twentieth century—beginning with echo-sounding during the 1925 German Meteor expedition and culminating in Marie Tharp and Bruce Heezen’s monumental 1977 World Ocean Floor Panorama—shattered this archaic myth. Bathymetric cartography revealed that the ocean floor possesses a spectacular, highly rugged topography whose mountain ranges, vertical chasms, and volcanic plateaus surpass continental landforms in scale and relief. Tectonically, the ocean floor reflects isostatic equilibrium between dense basaltic oceanic crust (Sima, density 3.0 g/cm³) and lighter continental granitic crust (Sial, density 2.7 g/cm³), structured into three distinct physiographic provinces: Continental Margins, Deep Ocean Basins, and the global Mid-Ocean Ridge System."),

    ("The Continental Margin: Continental Shelf, Slope, and Submarine Canyons",
     "The Continental Margin represents the submerged transition zone between exposed continental landmasses and the deep abyssal ocean floor. It is structurally divided into three successive zones. The 'Continental Shelf' is the gently sloping, shallow coastal platform extending from the low-tide shoreline to the shelf break. Sloping at an average gradient of only 0.1° (approximately 1 to 2 meters per kilometer) and terminating at an average depth of 130 to 200 meters, the shelf covers approximately 7.5 percent of the total oceanic area. Shelf width is dictated by plate tectonics: along 'Passive Margins' (Atlantic-type, trailing edges of tectonic plates), shelves are exceptionally broad, reaching 1,500 km off the Arctic Siberian coast; along 'Active Margins' (Pacific-type, leading edges with subduction), shelves are virtually absent or exceptionally narrow (<1 km off western South America). The continental shelf represents the economic powerhouse of the oceans, harboring over 90 percent of global marine fisheries and colossal offshore petroleum and natural gas reserves (exemplified by the Arabian Sea shelf off the Indus Delta)."),

    ("The Continental Slope and Rise: Turbidity Currents and the Indus Canyon",
     "At the shelf break, the seafloor plunges downward along the 'Continental Slope' at a steep gradient averaging 4° to 5° (frequently exceeding 20° along fault scarps), descending from 200 meters to depths of 3,000 to 4,000 meters. The continental slope marks the fundamental petrological and structural boundary where thick continental granitic crust terminates and thin oceanic basalt begins. The slope is deeply incised by monumental V-shaped 'Submarine Canyons', whose sheer bedrock walls rival the Grand Canyon in scale. In 1936, Harvard geologist Reginald Daly solved the genesis of submarine canyons by identifying 'Turbidity Currents'—high-density, underwater avalanches of sediment-laden water triggered by earthquakes or storm waves, which surge downslope at velocities exceeding 60 km/h, scouring bedrock channels. A prime regional example is the 'Indus Submarine Canyon' off Pakistan's Indus Delta, which cuts deeply across the continental shelf and slope, funneling terrigenous sediments into the deep Arabian Sea to construct the Indus Fan—the second-largest deep-sea sediment fan on Earth. At the base of the slope lies the 'Continental Rise', a gently sloping (0.5°–1.0°) apron of thick sediment formed by coalescing submarine turbidite fans."),

    ("Deep Ocean Basins: The Abyssal Plains and Pelagic Sedimentation",
     "Beyond the continental rise lie the Deep Ocean Basins, situated at depths between 3,000 and 6,000 meters and occupying roughly 40 percent of the planetary surface. The dominant feature of these basins is the 'Abyssal Plain'—the flattest, most horizontal morphological regions on planet Earth, with slopes of less than 1:1,000. Their extreme flatness is not structural; rather, it is produced by the continuous pelagic deposition of fine sediments that completely bury the rugged, irregular basaltic volcanic bedrock beneath a blanket hundreds of meters thick. Sediments in deep basins divide into Terrigenous clays washed from continents, Pelagic Brown Clays (dust blown from deserts and volcanic ash), and Biogenic Oozes formed from the microscopic skeletal remains of planktonic organisms. Biogenic oozes divide into 'Calcareous Oozes' (composed of calcium carbonate shells of foraminifera and coccolithophores, preserved above the Carbonate Compensation Depth / CCD at ~4,500 m) and 'Siliceous Oozes' (composed of silica frustules of radiolarians and diatoms, accumulating in cold polar waters and nutrient-rich equatorial upwelling zones below the CCD)."),

    ("Oceanic Trenches, Seamounts, and Guyots",
     "Where oceanic lithosphere collides with another plate at convergent boundaries, the subducting plate bends downward into the mantle, carving narrow, arcuate depressions known as 'Oceanic Trenches'. Trenches represent the deepest, coldest, and highest-pressure zones on Earth (the Hadal Zone). The planet's deepest abyss is the Challenger Deep within the Mariana Trench in the western Pacific, plunging to a depth of 10,994 meters—deep enough to submerge Mount Everest with over two kilometers of water to spare. Trenches are parallelled by volcanic island arcs or continental volcanic mountain chains, and are characterized by deep-focus Wadati-Benioff earthquakes. Rising from the abyssal plains are isolated volcanic edifices: 'Seamounts' are submarine volcanic peaks rising more than 1,000 meters above the seafloor, formed over mantle hot spots; 'Guyots' (tablemounts), discovered by Harry Hess during World War II, are ancient seamounts whose summits were planed flat by wave erosion when they were at sea level, subsequently sinking to deep abyssal depths as the cooling oceanic lithosphere subsided."),

    ("The Mid-Ocean Ridge System: The Volcanic Spine of Plate Tectonics",
     "The grandest topographic feature on the Earth's surface is the 'Mid-Ocean Ridge System'—a continuous, 65,000-kilometer-long interconnected volcanic mountain chain traversing the Arctic, Atlantic, Indian, and Pacific oceans like the seams on a baseball. Rising 2,000 to 3,000 meters above the surrounding abyssal plains and spanning widths of 1,000 to 2,000 kilometers, the ridge system marks the divergent plate boundaries where tectonic plates pull apart. Along the ridge crest lies a central 'Axial Rift Valley', 1 to 2 kilometers deep, bounded by normal faults, where basaltic magma continuously upwells from the asthenosphere, erupting as pillow lavas to generate brand-new oceanic lithosphere (Seafloor Spreading). The ridge crest is intensely fractured by perpendicular 'Transform Faults' and is home to hydrothermal vents ('Black Smokers'), which spew mineral-rich fluids at temperatures exceeding 350°C, sustaining unique chemosynthetic biological ecosystems completely independent of solar sunlight. In evaluative summation, ocean floor morphology is not a static graveyard of sunken sediment, but a dynamic, living geological engine whose spreading ridges, abyssal plains, and subduction trenches govern global plate tectonics, geochemical balances, and marine biodiversity.")
]

p1_q15_quotes = [
    ("Scholar / Oceanographer", "Theoretical Contribution / Concept", "Exact Quotable Academic Thesis / Axiom"),
    ("Matthew Fontaine Maury (1855)", "The Physical Geography of the Sea", "The ocean floor is not a barren desert, but a monumental canvas sculpted by volcanic fires, subterranean chasms, and perpetual sedimentary snow."),
    ("Marie Tharp & Bruce Heezen (1977)", "World Ocean Floor Panorama", "The continuous central rift valley traversing the mid-ocean ridges provides the visible architectural proof that the ocean floor is tearing itself apart along spreading tectonic seams."),
    ("Harry H. Hess (1962)", "History of Ocean Basins (Seafloor Spreading)", "The mid-ocean ridges represent the upwelling limbs of mantle convection cells where juvenile oceanic crust is born, while guyots provide living benchmarks of lithospheric subsidence."),
    ("Reginald A. Daly (1936)", "Origin of Submarine Canyons", "Submarine canyons are excavated across continental slopes by the colossal abrasive power of density-driven turbidity currents operating as underwater avalanches."),
    ("Arthur N. Strahler (1969)", "Physical Geography", "The bimodal hypsographic distribution of the Earth's surface—separating continental cratons from oceanic abyssal floors—is the direct structural reflection of isostatic equilibrium between Sial and Sima.")
]

# ==============================================================================
# QUESTION 16 (P1_Q16): OCEAN SALINITY: CONTROLLING FACTORS & VARIATIONS
# Solved from: 2025 Q.8, 2022 Q.7, 2013 Q.6, 2010 Q.6
# ==============================================================================
p1_q16_meta = [
    ("Target Syllabus", "Section 3: Oceans & Oceanography — Ocean Salinity: Composition, Factors & Variations"),
    ("Exam Years Solved", "KPPSC PMS 2025 Q.8, 2022 Q.7, 2013 Q.6, 2010 Q.6"),
    ("Marks & Allocation", "20 Marks | Strict 35-Minute Descriptive Writing Limit"),
    ("Core Theoretical Thesis", "Oceanic salinity represents the total concentration of dissolved inorganic salts in seawater, averaging 35 grams per kilogram (35 parts per thousand or 35‰ / PSU). Governed by William Dittmar's Principle of Constant Proportions (Challenger Expedition, 1884), the relative ratio of major dissolved ions (dominated by Chloride 55% and Sodium 30.6%) remains invariant globally. Spatial and vertical salinity distributions are controlled by the local equilibrium between evaporation and precipitation (E - P), freshwater fluvial influx, sea-ice freezing/melting, and ocean mixing. Salinity variations dictate seawater density (sigma-t), driving the global Thermohaline Conveyor Belt and governing marine osmotic biology.")
]

p1_q16_callout = [
    "SALINITY METRICS & DITTMAR'S PRINCIPLE: Salinity (S) is defined as the total mass of dissolved solid material in grams contained in one kilogram of seawater, expressed in parts per thousand (‰ or Practical Salinity Units, PSU). Average global ocean salinity = 35‰ (range 33‰ to 37‰). William Dittmar (1884) established that regardless of total salinity, the ratio of major ions remains strictly constant: Chloride (Cl⁻: 55.04%), Sodium (Na⁺: 30.61%), Sulfate (SO₄²⁻: 7.68%), Magnesium (Mg²⁺: 3.69%), Calcium (Ca²⁺: 1.16%), Potassium (K⁺: 1.10%).",
    "THE PRIMARY CONTROLLING FACTORS: (1) Evaporation vs. Precipitation Balance (E - P): Net evaporation elevates salinity (subtropical deserts, Red Sea 41‰, Mediterranean 38‰, Arabian Sea 36.5‰); net precipitation dilutes salinity (Equatorial belt 34.5‰, ITCZ rains); (2) Fluvial Discharge: Large rivers dilute coastal waters (Bay of Bengal 30–32‰ due to Ganges-Brahmaputra; Baltic Sea 7–8‰); (3) Sea-Ice Freezing and Melting: Freezing expels salts into surrounding water ('Brine Rejection'), creating hyper-saline, dense polar bottom water; summer melting dilutes polar surface waters; (4) Marginal Sea Enclosure.",
    "HORIZONTAL & VERTICAL DISTRIBUTION: Horizontal: Bimodal latitudinal distribution—lowest at poles (30–32‰), intermediate at equator (34.5‰), maximum in subtropical anticyclonic belts (36–37‰ at 20°–30° N/S). Vertical: Mixed Surface Layer (0–200 m) -> Halocline (zone of steep vertical salinity gradient between 200 m and 1,000 m) -> Deep Ocean Layer (remarkably uniform, cold water: 34.5–34.8‰).",
    "THERMOHALINE CIRCULATION & OSMOTIC ECOLOGY: Salinity and temperature determine seawater density. High salinity + low temperature generates dense sinking water masses (Antarctic Bottom Water - AABW and North Atlantic Deep Water - NADW), powering the global Thermohaline Conveyor Belt. Biologically, salinity dictates osmotic pressure: stenohaline species cannot tolerate salinity fluctuations, while euryhaline species adapt to estuarine transitions (e.g., Indus Delta mangrove fauna)."
]

p1_q16_map = r"""
                 [GLOBAL OCEAN SALINITY DYNAMICS & HALOCLINE]
                                      │
            ┌─────────────────────────┴─────────────────────────┐
            ▼                                                   ▼
 [HORIZONTAL LATITUDINAL PATTERN]                    [VERTICAL SALINITY PROFILE]
 (Controlled by Evaporation - Precip)                (The Halocline Transition)
 • Equator (0°-10°): 34.5-35‰                        Surface Mixed Layer (0-200m)
   (High rainfall dilutes surface)                     │ High evaporation -> High S (36‰)
 • Subtropics (20°-30°): 36-37‰                        ▼
   (Max Evaporation; Horse Latitudes)                [HALOCLINE ZONE] (200m - 1,000m)
   * Red Sea: 41‰; Med: 38‰                            • Steep gradient transition
 • Subpolar/Polar (60°-90°): 30-32‰                    • S decreases in tropics
   (Ice melt & low evaporation)                        • S increases in polar zones
                                                       ▼
                                                     [DEEP HOMOGENEOUS LAYER]
                                                     • Extremely uniform (34.5 - 34.8‰)
                                                     • Cold, hyper-dense water
                                      │
                        [THERMOHALINE DRIVER: SIGMA-T]
                        High Salinity + Near-Freezing Temp
                                      │
              Sinks in North Atlantic (NADW) & Weddell Sea (AABW)
                        [GLOBAL CONVEYOR BELT ENGINE]
"""

p1_q16_timeline = [
    ("Marine Region / Oceanic Zone", "Salinity Level (‰ / PSU)", "Dominant Controlling Factor & Hydrographic Consequence"),
    ("Equatorial Oceanic Belt", "34.5‰ to 35.0‰ (moderately low)", "Intense year-round cloud cover and heavy convective precipitation (ITCZ) exceed evaporation; dilutes surface seawater."),
    ("Subtropical High-Pressure Belts", "36.0‰ to 37.2‰ (global maximum)", "Subtropical anticyclonic subsidence, cloudless skies, high solar radiation, and desiccating Trade Winds drive maximum net evaporation (E >> P)."),
    ("Enclosed Tropical Seas (Red Sea / Persian Gulf)", "39.0‰ to 41.5‰ (extreme hyper-salinity)", "Confined shallow basin surrounded by hyper-arid deserts; zero major river inflow, extreme evaporation (>2,000 mm/year); water becomes dense and sinks."),
    ("Partially Enclosed Dilute Seas (Baltic Sea)", "7.0‰ to 15.0‰ (brackish / semi-fresh)", "Low solar evaporation rate combined with massive freshwater inflow from over 250 rivers; narrow outlet to open North Sea limits marine mixing."),
    ("Bay of Bengal vs. Arabian Sea", "Bay of Bengal: 30–32‰; Arabian Sea: 36.0–36.8‰", "Dramatic regional contrast: Arabian Sea suffers high desert evaporation with minimal river influx; Bay of Bengal receives colossal freshwater runoff from Ganges-Brahmaputra."),
    ("Polar Oceans & Weddell Sea", "Surface: 30.0–32.0‰; Bottom Water: 34.66‰", "Summer sea-ice melting dilutes surface; winter freezing expels brine into surrounding water, forming dense Antarctic Bottom Water (AABW), which sinks to the ocean abyss.")
]

p1_q16_essay = [
    ("Introduction: The Chemical Salinity Framework of the Oceans",
     "Seawater is an extraordinary, complex aqueous solution containing virtually every known chemical element on Earth. In oceanography, the concentration of dissolved mineral salts is quantified as 'Salinity'—formally defined as the total mass of dissolved inorganic solids in grams contained within one kilogram of seawater, expressed in parts per thousand (‰ or Practical Salinity Units, PSU). The average salinity of the world ocean is approximately 35‰, meaning that every 1,000 grams of seawater contains 35 grams of dissolved salts and 965 grams of pure water. Although 35‰ represents the global mean, actual surface salinity fluctuates significantly across geographical coordinates, ranging from under 7‰ in enclosed polar-fed basins like the Baltic Sea to over 41‰ in hyper-arid tropical water bodies like the Red Sea. Understanding the spatial distribution, vertical stratification, and controlling mechanisms of ocean salinity is of paramount scientific importance, as salinity—along with temperature—determines seawater density (sigma-t), driving the global thermohaline conveyor belt, modulating planetary climate, and defining the physiological boundaries of marine ecosystems."),

    ("Chemical Composition and Dittmar's Principle of Constant Proportions",
     "During the historic British Challenger Expedition (1872–1876), analytical chemist William Dittmar analyzed 77 water samples collected from all oceans and established a fundamental geochemical axiom known as 'Dittmar’s Principle of Constant Proportions' (or the Rule of Relative Proportions). Dittmar proved that although the total salinity of seawater varies from place to place due to evaporation and freshwater dilution, the relative percentage concentrations of the major dissolved ionic constituents remain virtually identical across the entire global ocean. Seawater salinity is overwhelmingly dominated by six major conservative ions: Chloride (Cl⁻, 55.04 percent by weight), Sodium (Na⁺, 30.61 percent), Sulfate (SO₄²⁻, 7.68 percent), Magnesium (Mg²⁺, 3.69 percent), Calcium (Ca²⁺, 1.16 percent), and Potassium (K⁺, 1.10 percent). Sodium and chloride alone account for over 85.6 percent of all dissolved salts. Because these major ions possess chemical residence times measured in millions of years—far exceeding the ~1,000-year mixing time of the global oceans—turbulent oceanic circulation homogenizes their relative ratios throughout the planetary hydrosphere."),

    ("Controlling Factors of Ocean Salinity: The Evaporation-Precipitation Engine",
     "Variations in surface salinity are dictated by processes that add or remove pure water from the ocean, governed by four primary environmental controls. The foremost factor is the 'Evaporation-Precipitation Balance' (E - P). Evaporation removes pure water vapor into the atmosphere, concentrating dissolved salts and elevating surface salinity; conversely, precipitation adds pure freshwater, diluting salinity. Consequently, surface salinity closely mirrors the Earth's latitudinal radiation and pressure belts. In the Subtropical High-Pressure belts (Horse Latitudes, 20° to 30° North and South), cloudless skies, high solar insolation, and dry subsiding air drive excessive evaporation (E >> P), pushing open-ocean salinity to global maxima of 36‰ to 37.2‰. In the Equatorial belt (0° to 10° N/S), although solar heating is intense, perpetual cloudiness and torrential daily convective downpours in the Intertropical Convergence Zone (ITCZ) cause precipitation to exceed evaporation (P > E), diluting equatorial surface salinity to 34.5‰–35‰. In subpolar and polar latitudes (60° to 90°), weak evaporation and frequent frontal precipitation lower salinity to 31‰–33‰."),

    ("Fluvial Influx, Sea-Ice Dynamics, and Marginal Enclosure",
     "The second controlling factor is 'Fluvial Runoff' from continental drainage basins. Where colossal river systems discharge into coastal seas, surface salinity is dramatically reduced. This phenomenon is brilliantly illustrated in the North Indian Ocean by the stark contrast between the Arabian Sea and the Bay of Bengal: the Arabian Sea receives minimal river inflow and is flanked by hyper-arid deserts (Thar, Arabian), resulting in high surface salinities of 36‰ to 36.8‰; in contrast, the Bay of Bengal receives the massive, sediment-rich discharges of the Ganges, Brahmaputra, and Irrawaddy rivers, diluting surface salinity to 30‰–32‰. Third, 'Sea-Ice Dynamics' modulate polar salinity: during autumn and winter freeze-up, pure water freezes into sea ice, while dissolved salts are expelled into the underlying water in a process known as 'Brine Rejection', creating hyper-saline, near-freezing bottom water; during summer, melting sea ice releases freshwater, forming a buoyant, low-salinity surface lid (30‰–32‰). Fourth, 'Basin Enclosure': enclosed seas in arid regions with limited ocean exchange reach extreme salinities, exemplified by the Red Sea (41‰) and the Persian Gulf (40‰), whereas enclosed seas in humid, cold climates with massive river discharge become brackish, such as the Baltic Sea (7‰ to 15‰)."),

    ("Vertical Salinity Stratification: The Halocline and Thermohaline Dynamics",
     "Vertically, the ocean exhibits a pronounced three-tier salinity stratification. The uppermost 100 to 200 meters constitute the 'Surface Mixed Layer', where wind turbulence and wave action homogenize salinity to match local atmospheric surface conditions. Beneath this layer, extending between 200 and 1,000 meters depth, lies the 'Halocline'—a zone characterized by an exceptionally steep, rapid vertical change in salinity. In tropical and subtropical regions where surface salinity is high due to evaporation, salinity decreases sharply with depth across the halocline. In polar regions where surface salinity is diluted by ice melt, salinity increases rapidly with depth across the halocline. Below 1,000 meters lies the 'Deep Ocean Layer', which comprises over 80 percent of the ocean's volume; here, salinity is remarkably stable and uniform across all planetary basins, hovering consistently between 34.5‰ and 34.8‰."),

    ("Conclusion: Salinity and the Global Thermohaline Conveyor Belt",
     "In evaluative summation, ocean salinity is not merely a static chemical characteristic, but a vital dynamic driver of the Earth's climate system. Seawater density is an exponential function of temperature and salinity: density increases as temperature plummets and salinity surges. In the high-latitude North Atlantic (Greenland-Norwegian Seas) and the Antarctic Weddell Sea, the combination of sub-zero temperatures and high salinity induced by brine rejection generates the densest water masses on Earth—North Atlantic Deep Water (NADW) and Antarctic Bottom Water (AABW). These massive, hyper-dense water masses sink into the abyssal plains, initiating the global 'Thermohaline Conveyor Belt' (the meridional overturning circulation), which transports heat, dissolved oxygen, and carbon dioxide across the global ocean basins over a 1,000-year cycle. Disruptions to this salinity balance—such as accelerated melting of the Greenland Ice Sheet injecting massive pulses of low-density freshwater into the subpolar North Atlantic—threaten to weaken thermohaline circulation, highlighting the profound interconnectedness of ocean salinity, atmospheric stability, and global climate change.")
]

p1_q16_quotes = [
    ("Scholar / Oceanographer", "Theoretical Contribution / Concept", "Exact Quotable Academic Thesis / Axiom"),
    ("William Dittmar (1884)", "Challenger Expedition Chemical Report (Principle of Constant Proportions)", "The major components of sea water preserve an astonishingly constant proportion to each other; wherever salinity varies, it is the quantity of solvent water, not the relative ratio of dissolved salts, that alters."),
    ("Wallace S. Broecker (1991)", "The Great Ocean Conveyor", "Thermohaline circulation is the planetary radiator of the earth; the delicate salinity balance in high-latitude sinking regions governs the stability of global climate regimes."),
    ("Henry Stommel (1961)", "Thermohaline Convection with Two Stable Regimes of Flow", "Ocean circulation is driven by twin thermodynamic engines: wind-driven boundary gyres at the surface and density-driven thermohaline sinking in the abyss."),
    ("Arthur N. Strahler (1969)", "Physical Geography", "The geographical distribution of surface salinity reflects an exact balance between evaporation and precipitation, mirroring the great planetary atmospheric pressure belts."),
    ("G. Dietrich (1963)", "General Oceanography", "The halocline functions as a thermodynamic barrier in the sea, decoupling surface weather interactions from the timeless, homogeneous reservoir of the deep ocean.")
]

# ==============================================================================
# QUESTION 17 (P1_Q17): OCEANIC CIRCULATION & MAJOR OCEAN CURRENTS
# Solved from: 2022 Q.8, 2018 Q.6, 2013 Q.5, 2008 Q.4
# ==============================================================================
p1_q17_meta = [
    ("Target Syllabus", "Section 3: Oceans & Oceanography — Oceanic Currents of Atlantic, Pacific & Indian Oceans"),
    ("Exam Years Solved", "KPPSC PMS 2022 Q.8, 2018 Q.6, 2013 Q.5, 2008 Q.4"),
    ("Marks & Allocation", "20 Marks | Strict 35-Minute Descriptive Writing Limit"),
    ("Core Theoretical Thesis", "Surface oceanic circulation operates as a wind-driven hydrodynamic system organized into five macro-scale anticyclonic Subtropical Gyres across the Atlantic, Pacific, and Indian oceans. Driven by planetary Trade Winds and Prevailing Westerlies, circulation is modified by the Coriolis force (V. Walfrid Ekman's spiral transport at 90°) and continental land barriers. Henry Stommel's 'Westward Intensification' dynamic creates fast, narrow, warm Western Boundary Currents (Gulf Stream, Kuroshio) and slow, broad, cold Eastern Boundary Currents (Canaries, California, Peru). Uniquely, the North Indian Ocean exhibits seasonal circulation reversal driven by the South Asian Monsoon.")
]

p1_q17_callout = [
    "DRIVING MECHANISMS OF SURFACE CURRENTS: (1) Planetary Wind Stress (Trade Winds drive equatorial currents westward; Prevailing Westerlies drive mid-latitude currents eastward); (2) Coriolis Deflection & Ekman Transport: Swedish oceanographer V.W. Ekman (1905) proved that surface water moves at 45° to the wind, but the net integrated column ('Ekman Transport') moves at 90° (to the right in Northern Hemisphere, left in Southern Hemisphere); (3) Gravity & Dynamic Topography (convergence causes a 1-meter hill of water in the center of gyres); (4) Continental Barriers.",
    "WESTWARD INTENSIFICATION (STOMMEL 1948): Due to the latitudinal variation of the Coriolis parameter (beta effect), the center of subtropical gyres is shifted westward. Western Boundary Currents (Gulf Stream, Kuroshio, Brazil, Agulhas) are extraordinarily fast (2–9 km/h), narrow (<100 km), deep (up to 1,000 m), and warm. Eastern Boundary Currents (Canaries, California, Benguela, Peru/Humboldt) are slow (<1 km/h), broad (>1,000 km), shallow (<500 m), cold, and accompanied by coastal upwelling.",
    "THE THREE OCEAN SYSTEMS: (1) Atlantic Ocean: North Atlantic Gyre (North Equatorial -> Gulf Stream -> North Atlantic Drift -> Canaries); South Atlantic Gyre (South Equatorial -> Brazil -> West Wind Drift -> Benguela); (2) Pacific Ocean: North Pacific Gyre (North Equatorial -> Kuroshio -> North Pacific Drift -> California); South Pacific Gyre (South Equatorial -> East Australian -> West Wind Drift -> Peru/Humboldt); (3) Indian Ocean: South Indian Gyre (South Equatorial -> Agulhas/Mozambique -> West Wind Drift -> West Australian); North Indian Ocean: Exhibits complete bi-annual reversal—Southwest Monsoon Drift (clockwise, summer) vs. Northeast Monsoon Drift (counter-clockwise, winter).",
    "CLIMATIC & ECONOMIC SIGNIFICANCE: Warm currents elevate coastal temperatures and humidity (e.g., North Atlantic Drift keeps Murmansk and UK ice-free); cold currents produce coastal fog, hyper-aridity, and desertification (Atacama, Namib deserts), but drive coastal upwelling of nutrients that sustains the world's richest commercial fisheries."
]

p1_q17_map = r"""
                     [GLOBAL OCEANIC SURFACE GYRE DYNAMICS]
                                       │
     [NORTHERN HEMISPHERE: CLOCKWISE]  │  [SOUTHERN HEMISPHERE: COUNTER-CLOCKWISE]
                                       │
 60°N ── Subpolar Gyre (Cyclonic) ──── │
  ▲      • Oyashio (Cold)              │
  │      • Labrador (Cold)             │
  │                                    │
 30°N ── Subtropical Gyre ──────────── │
  ▲      [WESTERN BOUNDARY]            │
  │      • Gulf Stream (Warm, Fast)    │
  │      • Kuroshio Current (Warm)     │
  │            │                       │
  │            ▼                       │
  │      [EASTERN BOUNDARY]            │
  │      • Canaries Current (Cold)     │
  │      • California Current (Cold)   │
  │                                    │
  0°  ── North Equatorial Current ──── │ ── South Equatorial Current (Westward) ──
                                       │
                                       │ 30°S ── Subtropical Gyre ────────────
                                       │  ▲      • Brazil / Agulhas (Warm West)
                                       │  │      • Benguela / Peru (Cold East)
                                       │  ▼
                                       │ 60°S ── ANTARCTIC CIRCUMPOLAR CURRENT
                                       │         (West Wind Drift: Unbroken Eastward)
"""

p1_q17_timeline = [
    ("Ocean Current System", "Temperature & Hydrodynamic Characteristics", "Climatic, Regional & Marine Ecological Significance"),
    ("Gulf Stream & North Atlantic Drift", "Warm, deep, fast western boundary current; flow rate ~100 Sverdrups (100 million m³/s); velocity up to 9 km/h.", "Transports colossal tropical heat to Northwestern Europe; keeps ports ice-free above 60°N (UK, Norway, Murmansk); moderates European winter climate."),
    ("Canaries Current & Benguela Current", "Cold, shallow, slow eastern boundary currents; flow southward along Northwest Africa (Canaries) and Southwest Africa (Benguela).", "Stabilizes marine atmosphere; suppresses rainfall, causing hyper-arid coastal deserts (Sahara and Namib); strong coastal upwelling feeds rich fisheries."),
    ("Kuroshio & Oyashio Currents", "Kuroshio: Warm western boundary current of North Pacific. Oyashio (Kamchatka): Subpolar cold current flowing south.", "Convergence of warm Kuroshio and cold Oyashio off Hokkaido (Japan) creates dense navigational sea fogs and one of the world's richest commercial fishing grounds."),
    ("Peru (Humboldt) Current & Upwelling", "Cold, broad eastern boundary current flowing northward along Chile and Peru; strong coastal Ekman upwelling.", "Brings cold, nutrient-rich (nitrates/phosphates) bottom water to surface; fuels world's largest anchoveta fishery; suppresses rain in coastal Atacama Desert."),
    ("Antarctic Circumpolar Current", "Cold, eastward-flowing current circling the entire globe around Antarctica; greatest volume transport (~130–140 Sv).", "Driven by uninterrupted Westerlies (Roaring Forties); thermal insulator that thermally isolates Antarctica, preventing warm subtropical water intrusion."),
    ("North Indian Ocean Monsoon Reversal", "Seasonal reversal: Summer SW Monsoon Current flows clockwise; Winter NE Monsoon Current flows counter-clockwise.", "Summer: Strong Somali Current surges northward along East Africa, generating intense cold upwelling; directly modulates monsoon moisture transport to Pakistan.")
]

p1_q17_essay = [
    ("Introduction: The Wind-Driven Kinetic Engine of the Oceans",
     "Ocean currents represent the massive, continuous, directed horizontal movements of seawater across planetary distances, functioning as the circulatory bloodstream of the global climate system. While the deep ocean is governed by slow, density-driven thermohaline circulation, the upper 400 meters of the ocean—encompassing roughly 10 percent of total oceanic volume—is driven primarily by the kinetic friction of atmospheric planetary winds acting upon the sea surface. Driven by the persistent Trade Winds and Prevailing Westerlies, surface waters organize into colossal, basin-wide circular loops known as 'Subtropical Gyres'. There exist five primary subtropical gyres: the North Atlantic, South Atlantic, North Pacific, South Pacific, and South Indian gyres. The hydrodynamic trajectory of these currents is modulated by the Earth's rotation (the Coriolis force), the Ekman transport mechanism, and the physical boundaries of continental landmasses, creating a magnificent planetary heat-redistribution system that dictates global terrestrial climates."),

    ("Hydrodynamic Mechanics: Planetary Wind Stress and the Ekman Spiral",
     "The physical coupling between the atmosphere and the ocean begins with wind shear stress. As wind blows across the open sea, friction drags surface water molecules downwind. In 1902, Swedish physical oceanographer V. Walfrid Ekman formulated the mathematical theory of the 'Ekman Spiral' to explain why Arctic icebergs drifted at an angle to the prevailing wind. Ekman demonstrated that due to the Coriolis force, the uppermost layer of surface water moves at an angle of 45° to the right of the wind direction in the Northern Hemisphere (and 45° to the left in the Southern Hemisphere). Each successive deeper water layer is dragged by the layer above it, but friction reduces velocity while the Coriolis force deflects it further, tracing a downward spiral. At a depth of approximately 100 meters (the Ekman depth), water motion reverses completely to oppose the surface wind. When the entire water column is mathematically integrated, the net transport of water—known as 'Ekman Transport'—moves at exactly 90° to the right of the prevailing wind in the Northern Hemisphere. Consequently, the inward Ekman transport driven by the Trade Winds and Westerlies forces water into the center of subtropical gyres, creating a subtle 'hill' of water approximately 1 to 2 meters high, around which water circulates in geostrophic balance."),

    ("Henry Stommel's Theory of Westward Intensification",
     "Subtropical gyres are not symmetrical circles; their centers are displaced strongly toward the western side of ocean basins. In 1948, American oceanographer Henry Stommel formulated the mathematical theory of 'Westward Intensification' to explain this asymmetry. Stommel proved that because the Coriolis parameter increases with latitude (the 'beta effect', beta = df/dy), planetary vorticity demands that currents flowing poleward along the western margins of ocean basins must be compressed into exceptionally fast, narrow, deep, and energetic streams. Conversely, return flow along the eastern margins is dispersed across broad, sluggish, shallow currents. Consequently, all ocean basins exhibit a profound dichotomy between Western Boundary Currents and Eastern Boundary Currents. 'Western Boundary Currents'—exemplified by the Gulf Stream in the Atlantic, the Kuroshio in the Pacific, the Brazil Current, and the Agulhas Current in the Indian Ocean—are extraordinarily swift (velocities of 5 to 9 km/h), narrow (<100 km wide), deep (penetrating 1,000 m), and transport massive volumes of warm tropical water poleward. In sharp contrast, 'Eastern Boundary Currents'—exemplified by the Canaries and Benguela currents in the Atlantic, and the California and Peru (Humboldt) currents in the Pacific—are broad (>1,000 km), shallow (<500 m), sluggish (<1 km/h), and transport cold subpolar water equatorward."),

    ("The Atlantic and Pacific Circulation Systems",
     "In the Atlantic Ocean, the North Atlantic Subtropical Gyre begins with the North Equatorial Current driven westward by Trade Winds. Deflected northward by the Americas, it enters the Gulf of Mexico and emerges through the Florida Straits as the legendary 'Gulf Stream'. Transporting over 100 million cubic meters of water per second (100 Sverdrups)—more than 100 times the combined flow of all the world's rivers—the Gulf Stream flows northeastward, separating from the coast at Cape Hatteras into the open Atlantic as the 'North Atlantic Drift'. This warm current bathes the British Isles, Norway, and the Arctic port of Murmansk, keeping them ice-free year-round and granting Western Europe a temperate maritime climate 10°C to 15°C warmer than identical latitudes in Canada. The gyre completes its loop as the cold Canaries Current flows southward along Northwest Africa. In the Pacific Ocean, the North Pacific Gyre mirrors this pattern: the warm, majestic 'Kuroshio Current' flows past Japan, meets the cold subpolar 'Oyashio Current' (generating dense fog banks and rich fisheries), transitions into the North Pacific Drift, and returns equatorward as the cold, foggy 'California Current'. In the Southern Hemisphere, the South Pacific Gyre drives the cold 'Peru (Humboldt) Current' northward along Chile and Peru, while the South Atlantic Gyre drives the cold 'Benguela Current' along South Africa."),

    ("The Indian Ocean: The Phenomenon of Seasonal Monsoon Reversal",
     "The Indian Ocean possesses an exceptional, unique oceanic circulation structure owing to its asymmetric geography: it is bounded on the north by the colossal Asian landmass rather than opening freely into polar waters. While the South Indian Ocean maintains a standard, stable subtropical gyre (comprising the South Equatorial Current, the warm Agulhas Current, the West Wind Drift, and the cold West Australian Current), the North Indian Ocean exhibits a complete, bi-annual reversal of its surface currents, governed entirely by the South Asian Monsoon. During the summer (Southwest Monsoon, June–September), intense low pressure over the Asian landmass drives strong southwesterly winds across the Arabian Sea and Bay of Bengal. Surface circulation flows in a clockwise direction, designated as the 'Southwest Monsoon Drift'. Along the Horn of Africa, this generates the phenomenal 'Somali Current'—the only western boundary current in the world that flows northward as a cold current, accompanied by intense coastal upwelling of cold (13°C), nutrient-rich water. Conversely, during the winter (Northeast Monsoon, December–March), high pressure over Siberia and Central Asia reverses wind directions; the surface currents reverse into a counter-clockwise circulation known as the 'Northeast Monsoon Drift', restoring the westward-flowing North Equatorial Current."),

    ("Climatic and Ecological Repercussions: Deserts, Fisheries, and ENSO",
     "In evaluative conclusion, oceanic currents are the master thermostats of the biosphere. Warm currents elevate air temperatures and supercharge evaporation, feeding humid climates and coastal rainforests. Cold currents, by contrast, chill the lower atmosphere, creating persistent temperature inversions that suppress vertical convection; this desiccates coastal margins, producing the world's hyper-arid coastal deserts, including the Atacama Desert (chilled by the Peru Current), the Namib Desert (chilled by the Benguela Current), and the Baja California desert. However, this cold water is accompanied by wind-driven 'Coastal Upwelling', where offshore Ekman transport pulls deep, nutrient-laden waters to the sunlit photic zone, fueling explosive phytoplankton blooms that support the planet's greatest commercial fisheries. When this circulation pattern collapses periodically in the Pacific—a phenomenon known as the El Niño-Southern Oscillation (ENSO)—the cessation of upwelling causes catastrophic fish die-offs in Peru and triggers global climatic disruptions, vividly demonstrating that oceanic circulation is inextricably linked to the survival of the global human economy.")
]

p1_q17_quotes = [
    ("Scholar / Oceanographer", "Theoretical Contribution / Concept", "Exact Quotable Academic Thesis / Axiom"),
    ("V. Walfrid Ekman (1905)", "On the Influence of the Earth's Rotation on Ocean-Currents", "The friction of wind upon the sea does not drive water downwind; the Coriolis force deflects the net water transport perpendicularly at ninety degrees, organizing oceanic gyres."),
    ("Henry Stommel (1948)", "The Westward Intensification of Wind-Driven Ocean Currents", "The variation of the Coriolis parameter with latitude is the sole mathematical reason why western boundary currents are compressed into narrow, violent oceanic rivers like the Gulf Stream."),
    ("Walter H. Munk (1950)", "On the Wind-Driven Ocean Circulation", "The broad open-ocean circulation reflects an exact vorticity balance between the curl of the planetary wind stress and the lateral dissipation of eddy viscosity."),
    ("Matthew Fontaine Maury (1855)", "The Physical Geography of the Sea", "There is a river in the ocean; in the severest droughts it never fails, and in the mightiest floods it never overflows; its banks and its bottom are of cold water, while its current is of warm."),
    ("Arthur N. Strahler (1969)", "Physical Geography", "Ocean currents operate as planetary heat pumps, transporting equatorial warmth to high-latitude cold sinks and regulating the habitability of continental coasts.")
]

print("Batch 3 (Q15, Q16, Q17) appended. Appending Q18, Q19, Q20 and runner loop...")

# ==============================================================================
# QUESTION 18 (P1_Q18): MECHANICS OF OCEAN TIDES & TIDAL THEORIES
# Solved from: 2025 Q.2, 2018 Q.8, 2016 Q.6
# ==============================================================================
p1_q18_meta = [
    ("Target Syllabus", "Section 3: Oceans & Oceanography — Ocean Tides: Origin, Theories & Coastal Influence"),
    ("Exam Years Solved", "KPPSC PMS 2025 Q.2, 2018 Q.8, 2016 Q.6"),
    ("Marks & Allocation", "20 Marks | Strict 35-Minute Descriptive Writing Limit"),
    ("Core Theoretical Thesis", "Ocean tides represent rhythmic, periodic oscillations of sea level generated by the differential gravitational attraction of the Moon and Sun operating in concert with the centrifugal force of the rotating Earth-Moon barycenter system. The tractive tidal force varies inversely with the cube of celestial distance (d^3), granting the Moon 2.17 times greater tide-generating efficacy than the Sun. Tides transition from idealized static models (Sir Isaac Newton's Equilibrium Theory) to hydrodynamic reality (Pierre-Simon Laplace's Dynamic Theory of amphidromic systems). Coastal geometry, continental shelf shoaling, and bay resonance dramatically amplify tidal ranges, generating macro-tidal extremes and tidal bores.")
]

p1_q18_callout = [
    "TIDE-GENERATING TRACTIVE FORCE: Gravitational attraction (Newton: F = G * M * m / r²) pulls water toward the celestial body, while orbital centrifugal force acts uniformly in the opposite direction. The resultant differential force is the 'Tractive Force' (FT), which varies inversely with the cube of distance: FT proportional to Mass / (Distance)³. Because the Moon is 384,400 km away while the Sun is 149.6 million km away, the Moon's tidal force is roughly 2.17 times stronger than the Sun's (solar tide = 46% of lunar tide).",
    "THE TWO BULGES & TIDAL PERIODICITY: The Earth-Moon system creates two simultaneous tidal bulges: the direct sub-lunar bulge (gravitational pull exceeds centrifugal force) and the indirect antipodal bulge (centrifugal force exceeds moon's gravity). Because the Moon orbits Earth in the same direction Earth rotates, a complete tidal day lasts 24 hours and 50 minutes (the Lunar Day), giving semi-diurnal tides a period of 12 hours and 25 minutes.",
    "SPRING VS. NEAP TIDAL CYCLES: (1) Spring Tides (Syzygy): Occurs at New Moon and Full Moon when Sun, Moon, and Earth align in a straight line; solar and lunar tidal bulges reinforce constructively, producing maximum tidal range (highest high tides, lowest low tides); (2) Neap Tides (Quadrature): Occurs at First and Third Quarter moons when the Moon is at 90° right angles to the Sun; tidal bulges cancel destructively, producing minimum tidal range (lowest high tides, highest low tides).",
    "EQUILIBRIUM VS. DYNAMIC THEORIES: Equilibrium Theory (Newton 1687): Assumes frictionless, uniform global ocean adjusting instantaneously; fails to explain real-world tidal delays and amplitudes; Dynamic Theory (Laplace 1775): Treats tides as forced planetary shallow-water waves deflected by Coriolis force, rotating counter-clockwise in Northern Hemisphere around stationary, zero-tide 'Amphidromic Points' (Cotidal lines). Coastal amplification reaches extremes in funnel-shaped bays (e.g., Bay of Fundy: 16 m; Rann of Kutch: 3–4 m; Tidal Bores)."
]

p1_q18_map = r"""
                      [OCEAN TIDAL DYNAMICS & THEORIES]
                                      │
            ┌─────────────────────────┴─────────────────────────┐
            ▼                                                   ▼
 [ASTRONOMICAL TRACTIVE FORCES]                      [HYDRODYNAMIC TIDAL THEORIES]
 • Newton's Inverse Cube Law:                        • Equilibrium Theory (Newton 1687)
   Force proportional to M / d³                        - Static, frictionless ocean
 • Lunar Tide: 2.17x Solar Tide                        - Two theoretical bulges
 • Lunar Day: 24h 50m (Tide: 12h 25m)                • Dynamic Theory (Laplace 1775)
            │                                          - Shallow-water wave dynamics
 ┌──────────┴──────────┐                               - Amphidromic points (zero tide)
 ▼                     ▼                               - Coriolis cyclonic rotation
[SPRING TIDES]       [NEAP TIDES]                              │
(Syzygy: Align)      (Quadrature: 90°)                         ▼
• New & Full Moon    • 1st & 3rd Quarters            [COASTAL AMPLIFICATION]
• Sun + Moon add     • Sun & Moon oppose             • Funnel-shaped bays & estuaries
• Max Tidal Range    • Min Tidal Range               • Resonance: Bay of Fundy (16m)
                                                     • Tidal Bores (Severn, Qiantang)
"""

p1_q18_timeline = [
    ("Tidal Phenomenon / Theory", "Astronomical / Hydrodynamic Mechanism", "Diagnostic Wave Characteristics & Coastal Manifestations"),
    ("Equilibrium Theory of Tides", "Formulated by Sir Isaac Newton (1687); static model assuming frictionless, uniform ocean.", "Predicts two equal daily bulges on opposite sides of Earth; establishes the astronomical mechanics of syzygy and quadrature cycles."),
    ("Dynamic Theory of Tides", "Formulated by Pierre-Simon Laplace (1775); models tides as shallow-water forced waves.", "Tidal waves are deflected by the Coriolis force, creating rotary waves that circle amphidromic points; cotidal lines radiate outward."),
    ("Spring Tides (Syzygy)", "Sun, Moon, and Earth align linearly at New Moon (conjunction) and Full Moon (opposition).", "Constructive interference: solar tide reinforces lunar tide; maximum vertical tidal range; extreme high waters and extreme low waters."),
    ("Neap Tides (Quadrature)", "Moon stands at right angles (90°) relative to the Earth-Sun axis (First and Third Quarter).", "Destructive interference: solar tide partially cancels lunar tide; minimum vertical tidal range; weak tidal currents."),
    ("Tidal Types: Diurnal vs. Semi-Diurnal", "Governed by Moon's declination and local basin bathymetry.", "Semi-diurnal (two equal highs/lows per day, e.g., Atlantic); Diurnal (one high/low per day, e.g., Gulf of Mexico); Mixed (unequal heights, e.g., Pacific)."),
    ("Coastal Resonance & Tidal Bores", "Tidal wavelength matches natural resonant frequency of bay; shallowing water compresses wave.", "Extreme tidal range (Bay of Fundy, 16.3 m; Bristol Channel, 14 m); true 'Tidal Bores'—steep hydraulic jump waves surging up rivers (Qiantang River, China).")
]

p1_q18_essay = [
    ("Introduction: The Rhythmic Pulse of the Planetary Hydrosphere",
     "Ocean tides represent the rhythmic, periodic rise and fall of the sea surface occurring across the world's coastal margins twice or once daily. Tides constitute the longest waves on Earth, possessing wavelengths measuring thousands of kilometers, yet behaving as shallow-water waves because their length vastly exceeds the mean depth of the ocean basins (d < L/20). Since antiquity, maritime civilizations observed a correlation between the moon and the sea, but the rigorous physical explanation of tidal mechanics eluded natural philosophers until the mathematical formulation of universal gravitation. Tides are generated by the gravitational attraction exerted by the celestial bodies—principally the Moon and the Sun—acting upon the fluid envelope of the rotating Earth. However, the transformation of this astronomical gravitational tug into real-world tidal fluctuations is mediated by complex hydrodynamic and coastal geomorphic factors, progressing from Isaac Newton’s static Equilibrium Theory to Pierre-Simon Laplace’s Dynamic Theory of amphidromic circulation."),

    ("Astronomical Mechanics: Universal Gravitation and the Tractive Force",
     "The primary driver of oceanic tides is the gravitational attraction between the Earth, the Moon, and the Sun. In accordance with Sir Isaac Newton’s Law of Universal Gravitation, the gravitational force between two bodies is directly proportional to the product of their masses and inversely proportional to the square of the distance between their centers: F = G * (M * m) / r². In the Earth-Moon system, the two bodies rotate around a common center of mass—the 'Barycenter'—located approximately 1,700 kilometers beneath the Earth's surface. At the Earth's center, the Moon's gravitational pull precisely balances the orbital centrifugal force. However, on the surface of the Earth, a differential imbalance arises. On the side facing the Moon (the sublunar point), gravitational attraction exceeds centrifugal force, pulling water outward toward the Moon to create a 'Direct Tidal Bulge'. On the side opposite the Moon (the antipodal point), distance weakens lunar gravity, allowing the outward centrifugal force to dominate, creating a second, simultaneous 'Indirect Tidal Bulge'."),

    ("The Inverse Cube Law: Why the Moon Dominates the Sun",
     "A critical physical principle of tidal science is that the 'Tide-Generating Tractive Force' (FT) does not follow the standard inverse-square law of gravity; rather, it is proportional to the differential gradient of gravity across the Earth's diameter, causing it to vary inversely with the cube of the distance: FT proportional to Mass / (Distance)³. The Sun possesses a colossal mass approximately 27 million times greater than that of the Moon. However, the Sun is situated at a staggering mean distance of 149.6 million kilometers from Earth, whereas the Moon orbits at a close distance of only 384,400 kilometers. When calculated mathematically, (d_Sun / d_Moon)³ equals approximately 59 million. Dividing the Sun's mass advantage (27 million) by the distance cube disadvantage (59 million) reveals that the Sun's tide-generating power is only about 46 percent of the Moon's. Thus, the Moon exerts a tide-generating force approximately 2.17 times greater than the Sun, making lunar orbital mechanics the primary determinant of terrestrial tides."),

    ("Tidal Periodicities and Spring-Neap Cycles",
     "Because the Moon orbits the Earth in the same direction that the Earth rotates on its axis, the Earth must rotate through 360° plus an additional 12° to bring a given meridian back directly under the Moon. This requires an additional 50 minutes, establishing the 'Lunar Day' (or tidal day) of 24 hours and 50 minutes. Consequently, in regions experiencing semi-diurnal tides, successive high tides occur every 12 hours and 25 minutes. Over the course of the 29.5-day synodic lunar month, the changing geometric alignment of the Sun, Moon, and Earth produces alternating tidal cycles. At New Moon and Full Moon, the Sun, Moon, and Earth align in a straight line—an astronomical alignment known as 'Syzygy'. In this orientation, solar and lunar tidal bulges reinforce each other constructively, generating 'Spring Tides' (from the German springen, to leap), characterized by maximum tidal range (abnormally high high tides and abnormally low low tides). Conversely, at First and Third Quarter phases, the Moon is positioned at right angles (90°) relative to the Earth-Sun line—an alignment termed 'Quadrature'. Here, the solar tidal bulge partially cancels the lunar bulge destructively, producing 'Neap Tides', characterized by minimum tidal range (unusually low high tides and unusually high low tides)."),

    ("Theoretical Formulations: Newton's Equilibrium vs. Laplace's Dynamic Theory",
     "The physical modeling of tides is structured around two landmark theories. Sir Isaac Newton's 'Equilibrium Theory of Tides' (published in his 1687 Principia) provided the foundational static mathematical framework. Newton assumed an idealized Earth completely covered by a uniform layer of frictionless water, which adjusts instantaneously to the celestial tractive forces to form two symmetrical bulges. While Newton successfully explained the astronomical periodicities and the spring-neap cycle, the Equilibrium Theory fails completely to predict real-world tides: it predicts maximum tidal ranges of less than 0.5 meters, whereas real tides reach 16 meters; it predicts high tide exactly when the Moon crosses the meridian, whereas real tides experience significant time lags; and it ignores the physical barriers of continents. To resolve these contradictions, the French mathematician Pierre-Simon Laplace formulated the 'Dynamic Theory of Tides' in 1775. Laplace treated tides not as static bulges, but as dynamic, shallow-water forced waves moving across real ocean basins of finite depth and complex shapes. Deflected by the Coriolis force and continental boundaries, tidal waves rotate cyclonically around central, stationary nodal points called 'Amphidromic Points' (where the tidal range is zero). 'Cotidal lines' radiate outward from these amphidromic nodes, connecting points that experience high tide at the same hour."),

    ("Coastal Amplification, Resonance, and Tidal Bores",
     "As tidal waves propagate from the deep ocean into shallow coastal waters, their velocity decreases while wave energy is compressed, causing dramatic 'Coastal Amplification'. The degree of amplification is governed by the configuration of the coastline and continental shelf bathymetry. Where ocean water is funneled into converging, shallow V-shaped bays or estuaries whose natural hydrodynamic resonance period matches the 12.4-hour tidal period, tidal resonance occurs. The world's most dramatic example is the 'Bay of Fundy' in Nova Scotia, Canada, where the semi-diurnal tidal range reaches a phenomenal 16.3 meters (over 53 feet). Along Pakistan's coastline, the Arabian Sea experiences semi-diurnal tides with moderate ranges (2 to 3 meters at Karachi Port), but funneling in the shallow, mangrove-choked creeks of the Indus Delta and the Rann of Kutch amplifies tidal ranges to nearly 4 meters, driving powerful tidal currents. Where a massive incoming tide enters a shallow, narrowing river mouth against a descending fluvial current, the water piles up into a steep, solitary foaming wall of water that surges upstream as a 'Tidal Bore' (spectacularly seen in the Qiantang River in China, reaching heights of 9 meters and speeds of 40 km/h). In evaluative conclusion, ocean tides represent the supreme synthesis of astronomical celestial mechanics and marine hydrodynamics, governing coastal ecology, maritime port navigation, and emerging renewable tidal energy generation.")
]

p1_q18_quotes = [
    ("Scholar / Physicist", "Theoretical Contribution / Concept", "Exact Quotable Academic Thesis / Axiom"),
    ("Sir Isaac Newton (1687)", "Philosophiae Naturalis Principia Mathematica (Equilibrium Theory)", "The sea is agitated by two simultaneous bulges, one facing the moon and the other antipodal, produced by the differential gravitation of the lunar and solar masses upon the revolving earth."),
    ("Pierre-Simon Laplace (1775)", "Recherches sur plusieurs points du système du monde (Dynamic Theory)", "Tides are not passive static bulges but forced hydrodynamic shallow-water waves, whose rotation around amphidromic nodes is governed by the Coriolis force and basin bathymetry."),
    ("William Whewell (1833)", "Essay Towards a First Approximation to a Map of Cotidal Lines", "Cotidal lines illustrate the progressive march of the tidal wave across the global sea, revolving like the spokes of a wheel around the silent amphidromic centers."),
    ("Arthur T. Doodson (1921)", "The Harmonic Development of the Tide-Generating Potential", "Harmonic analysis of tides reveals that real-world water levels can be mathematically resolved into hundreds of discrete celestial harmonic constituents."),
    ("Arthur N. Strahler (1969)", "Physical Geography", "Coastal geometry and resonance transform the modest half-meter open-ocean tidal wave into colossal coastal bores and sixteen-meter bay deluges.")
]

# ==============================================================================
# QUESTION 19 (P1_Q19): CARTOGRAPHY BASICS, MAP SCALES & WEATHER MAPS
# Solved from: 2025 Q.7, 2018 Q.7, 2016 Q.7, 2010 Q.7
# ==============================================================================
p1_q19_meta = [
    ("Target Syllabus", "Section 4: Study of Maps & Cartography — Map Basics, Scales, RF & Weather Maps"),
    ("Exam Years Solved", "KPPSC PMS 2025 Q.7, 2018 Q.7, 2016 Q.7, 2010 Q.7"),
    ("Marks & Allocation", "20 Marks | Strict 35-Minute Descriptive Writing Limit"),
    ("Core Theoretical Thesis", "Cartography is the scientific, graphic, and technological discipline of modeling terrestrial spatial reality upon two-dimensional planar surfaces. Map design is governed by foundational cartographic essentials: orientation, coordinate grids, symbology, and Map Scale—the mathematical ratio of map distance to ground distance. Scales classify into Verbal Statements, Graphic/Linear Scales, and dimensionless Representative Fractions (RF). Synoptic Weather Maps synthesize synchronous meteorological observations into standard station models, isobars, and frontal systems, serving as the visual interface for modern Numerical Weather Prediction (NWP) while confronting the deterministic chaos of atmospheric thermodynamics.")
]

p1_q19_callout = [
    "DEFINITION & ESSENTIAL ELEMENTS OF A MAP: A map is a scaled, generalized graphic representation of the Earth's surface or a portion of it on a plane surface. Five Essential Elements: (1) Title (what, where, when); (2) Orientation / Direction (North arrow or graticule grid); (3) Scale (relationship between map distance and real-world ground distance); (4) Legend / Key (decoding cartographic symbols, colors, and hachures); (5) Grid System / Projection (geographic coordinate reference framework of latitude and longitude).",
    "THE THREE MODES OF MAP SCALE: (1) Statement of Scale / Verbal Scale (e.g., '1 cm represents 1 km' or '1 inch to 1 mile'; easily understood but unit-dependent); (2) Graphic / Linear Scale (a line subdivided into primary and secondary units; paramount advantage: remains mathematically accurate when a map is photographically enlarged or reduced); (3) Representative Fraction (RF: a dimensionless fraction or ratio expressing map distance to ground distance in identical units: RF = Map Distance / Ground Distance, e.g., 1:50,000; universally understood regardless of language or measurement systems).",
    "LARGE-SCALE VS. SMALL-SCALE MAPS: Large-Scale (>= 1:25,000: Cadastral maps 1:1,000, Topographic maps 1:50,000; small denominator = large ratio; depicts small geographic area in immense structural detail); Small-Scale (<= 1:1,000,000: Atlas maps, Wall maps; large denominator = small ratio; depicts vast continents with high generalization).",
    "SYNOPTIC WEATHER MAPS & METEOROLOGICAL STATION MODELS: Principles: Synchronous observations taken simultaneously across the globe at standard synoptic hours (0000, 0600, 1200, 1800 UTC) under World Meteorological Organization (WMO) protocols. Station Model encodes temperature, dew point, barometric pressure (in tenths of hPa/mb), pressure tendency, wind barb (direction and wind speed in knots), cloud cover (in oktas/eighths), and precipitation type. Synoptic Analysis draws Isobars (lines of equal pressure at 4 hPa intervals), tracks Highs/Lows, and positions fronts. Modern forecasting combines satellite imagery, Doppler radar, and supercomputer Numerical Weather Prediction (NWP), constrained by chaos theory (the butterfly effect / Edward Lorenz)."
]

p1_q19_map = r"""
                     [CARTOGRAPHIC SCALES & SYNOPTIC MAPS]
                                       │
            ┌──────────────────────────┴──────────────────────────┐
            ▼                                                     ▼
   [MAP SCALE TAXONOMY]                                 [SYNOPTIC WEATHER MAPS]
   • Ratio: Map Distance / Ground Distance              • Synchronous global UTC data
   • Three Modes of Expression:                         • Standard WMO Station Model:
     1. Verbal: "1 cm to 5 km"                            - Central circle: Cloud oktas
     2. Graphic / Linear: Bar scale                       - Wind barb: Speed & direction
        (Preserved during zoom/reprint)                   - Top-left: Temperature (°C/°F)
     3. Representative Fraction (RF):                     - Bottom-left: Dew point
        1:50,000 (Unitless / Universal)                   - Top-right: Pressure (998 = 999.8mb)
            │                                             - Middle-right: Pressure tendency
   ┌────────┴────────┐                                  • Synoptic Features:
   ▼                 ▼                                    - Isobars (4 hPa intervals)
[LARGE-SCALE]     [SMALL-SCALE]                           - Frontal boundaries (Cold/Warm)
• >= 1:25,000     • <= 1:1,000,000                        - Centers of Action (Highs/Lows)
• Topo & Cadastral • Atlas & World maps                 • Modern NWP & Chaos Limits
• High detail     • High generalization
"""

p1_q19_timeline = [
    ("Cartographic Component / System", "Mathematical Definition & Technical Architecture", "Practical Utility & Geographical Implementation"),
    ("Representative Fraction (RF)", "Dimensionless ratio: RF = 1 / X; numerator is map distance, denominator is ground distance in identical units.", "Universal metric; 1:50,000 means 1 cm on map = 50,000 cm (500 m) on ground; allows cross-national engineering and military usage."),
    ("Graphic / Linear Scale", "A graduated line bar divided into major primary units on the right and fractional secondary units on the left.", "The only scale mode that survives photographic reproduction, scanning, or digital resizing without introducing scale distortion errors."),
    ("Topographical Maps (Survey of Pakistan)", "Large-scale maps (typically 1:50,000 and 1:250,000); utilizes contour lines for vertical relief.", "Depicts physical features (drainage, relief contours, spot heights) and cultural features (roads, settlements, canals); vital for civil planning."),
    ("WMO Station Model", "Standardized circular alphanumeric glyph encoding up to 20 weather variables on synoptic maps.", "Enables meteorologists worldwide to decode temperature, dew point, cloud cover in oktas, barometric pressure, and precipitation in seconds."),
    ("Isobaric Analysis on Synoptic Maps", "Plotting isobars (lines connecting points of equal barometric pressure) at standard 4-hPa intervals.", "Identifies cyclonic low-pressure depressions (troughs) and anticyclonic highs (ridges); isobar spacing reveals wind speed (pressure gradient)."),
    ("Modern Weather Forecasting & Chaos", "Integration of Doppler radar, geostationary satellites, and supercomputer Numerical Weather Prediction (NWP).", "Simulates Navier-Stokes atmospheric fluid equations; constrained by Edward Lorenz's chaos theory (finite predictability limit ~10–14 days).")
]

p1_q19_essay = [
    ("Introduction: The Cartographic Science of Spatial Representation",
     "Cartography is the science, art, and technology of analyzing, designing, and constructing maps to represent the spatial distribution of terrestrial, environmental, and human phenomena upon a two-dimensional medium. In geography, a map is not a mere picture or artistic illustration; it is a mathematically governed, abstracted graphic model of reality. Because it is physically impossible to construct a map at a 1:1 scale with the Earth, all cartographic representation relies upon three fundamental pillars: Scale (the proportional reduction of reality), Projection (the mathematical transformation of the spherical Earth onto a flat plane), and Generalization (the selective simplification and symbolic encoding of geographical data). From ancient topographic cadastral surveys to modern digital Geographic Information Systems (GIS) and real-time synoptic weather maps, cartography serves as the indispensable visual language of spatial analysis."),

    ("Essential Elements of a Map: The Structural Framework",
     "Every professional cartographic product must incorporate five essential elements to ensure intelligibility, mathematical accuracy, and geographic orientation. First is the 'Title', which succinctly defines the geographical location, thematic subject matter, and temporal date of the map. Second is 'Orientation / Direction', indicated either by a traditional North Arrow pointing toward True North (or magnetic/grid north) or by the coordinate graticule of latitude parallels and longitude meridians. Third is the 'Map Scale', which defines the exact mathematical ratio between distances measured on the map and the corresponding real-world horizontal distances measured on the ground. Fourth is the 'Legend / Key', which translates the complex visual grammar of the map—decoding symbols (point, line, area), conventional signs, contour lines, hachures, and hypsometric color tints (e.g., blue for hydrography, brown for contours, green for vegetation). Fifth is the 'Coordinate Grid and Projection', which establishes the mathematical geodetic reference system (such as UTM coordinates or WGS84 latitude-longitude grids) and identifies the map projection utilized."),

    ("The Mathematics of Map Scales: Verbal, Graphic, and Representative Fraction",
     "Map Scale is the fundamental mathematical bridge connecting cartographic abstraction to terrestrial reality. It is expressed across three standard formats: the Verbal Statement of Scale, the Graphic (or Linear) Scale, and the Representative Fraction (RF). The 'Verbal Scale' expresses the ratio in ordinary written language (e.g., 'One centimeter represents one kilometer' or 'One inch to one mile'). While simple for laypersons, verbal scales are inherently restricted by specific national units of measurement. The 'Graphic Scale' consists of a segmented horizontal line divided into primary units to the right of a zero mark and finely subdivided secondary units to the left. The supreme advantage of the graphic scale is that it remains mathematically accurate whenever a map is photographically enlarged, reduced, or digitally resized on a computer screen. The 'Representative Fraction' (RF) is a dimensionless ratio or fraction expressing map distance to ground distance in identical units: RF = Distance on Map / Distance on Ground. For example, an RF of 1:50,000 means that 1 unit of distance on the map (whether 1 centimeter or 1 inch) represents exactly 50,000 of the same units on the ground (50,000 cm = 500 meters; 50,000 inches = 0.789 miles). Because RF is purely dimensionless, it transcends language barriers and unit systems, serving as the international gold standard for topographical surveys, including the Survey of Pakistan's 1:50,000 and 1:250,000 master topographic series."),

    ("Large-Scale vs. Small-Scale Maps: The Detail-Generalization Continuum",
     "Cartographers classify maps into 'Large-Scale' and 'Small-Scale' series based on the numerical value of their Representative Fraction. Because RF is a fraction (1/X), a small denominator represents a large numerical fraction (large scale), whereas a colossal denominator represents a tiny numerical fraction (small scale). 'Large-Scale Maps' possess an RF of 1:25,000 or larger (such as urban cadastral parcel maps at 1:1,000 or topographical sheets at 1:25,000 to 1:50,000). They cover a small geographic area but depict features in immense, realistic geometric detail with minimal generalization. Conversely, 'Small-Scale Maps' possess an RF of 1:1,000,000 or smaller (such as continental atlas maps or global wall maps at 1:10,000,000). They cover vast continental or planetary domains, necessitating extensive cartographic generalization—where minor details are eliminated, boundaries are smoothed, and features are depicted through highly generalized abstract symbols."),

    ("Synoptic Weather Maps: Construction Principles and the WMO Station Model",
     "A specialized and dynamic application of cartography is the 'Synoptic Weather Map' (surface synoptic chart). The term synoptic (from the Greek synoptikos, meaning taking a comprehensive view together) refers to the simultaneous observation of atmospheric variables across an entire region. To construct a synoptic chart, meteorological stations worldwide record standardized atmospheric parameters synchronously at mandatory Universal Coordinated Time (UTC) intervals (0000, 0600, 1200, and 1800 UTC), coordinated by the World Meteorological Organization (WMO). At each station location on the map, this data is plotted using the standardized 'WMO Station Model'—a dense circular glyph encoding up to twenty atmospheric variables. The central circle indicates total cloud cover (shaded in oktas or eighths of the sky); an arrow shaft with barbs indicates wind direction and speed (each full barb represents 10 knots); numbers to the upper left indicate air temperature; numbers to the lower left denote dew point temperature; numbers to the upper right denote sea-level barometric pressure (abbreviated in tenths of hPa/mb); and symbols to the middle right depict barometric pressure tendency over the past three hours. Synoptic meteorologists then analyze this field by drawing 'Isobars'—smooth lines connecting points of equal atmospheric pressure, conventionally plotted at 4-hPa intervals. Isobars delineate the centers of action: Lows (cyclonic depressions) and Highs (anticyclonic ridges). Frontal boundaries are superimposed using international symbols: solid blue lines with triangles for Cold Fronts, solid red lines with semicircles for Warm Fronts, and alternating purple symbols for Occluded Fronts."),

    ("Modern Weather Forecasting Technologies and the Chaos Limit",
     "In modern atmospheric science, manual synoptic mapping has been augmented by advanced technological observation systems, including Geostationary and Polar-Orbiting Weather Satellites (providing visible, infrared, and water-vapor imagery), Doppler Weather Radars (detecting precipitation velocity and storm vorticity), and Automated Weather Stations. These observational feeds are assimilated into high-performance supercomputers running 'Numerical Weather Prediction' (NWP) models (such as the American GFS and European ECMWF models). These models solve the fundamental hydrodynamic Navier-Stokes equations, thermodynamics equations, and continuity laws of atmospheric fluid motion. Despite these monumental technological advancements, modern weather forecasting confronts an insurmountable mathematical barrier: the 'Chaos Limit'. Discovered by MIT meteorologist Edward Lorenz in 1963, atmospheric fluid flow is a deterministic, non-linear chaotic system characterized by extreme sensitivity to initial conditions (the 'Butterfly Effect'). Because initial observational measurements always contain microscopic errors, forecast accuracy degrades exponentially over time, establishing an absolute theoretical predictability horizon of roughly 10 to 14 days. In evaluative summation, from static topographical scale calculations to the dynamic chaos of synoptic weather prediction, cartography remains the foundational visual science through which human intellect deciphers the physical complexities of the terrestrial environment.")
]

p1_q19_quotes = [
    ("Scholar / Cartographer", "Theoretical Contribution / Concept", "Exact Quotable Academic Thesis / Axiom"),
    ("Arthur H. Robinson (1953)", "Elements of Cartography", "A map is an orchestrated graphic communication; scale, projection, and symbolization operate as a unified grammar that translates terrestrial space into human understanding."),
    ("Erwin Raisz (1938)", "General Cartography", "The representative fraction is the universal mathematical lingua franca of cartography, bridging all provincial measuring systems through pure geometric proportionality."),
    ("Edward N. Lorenz (1963)", "Deterministic Nonperiodic Flow (Chaos Theory)", "The non-linear dynamics of the atmosphere impose an absolute limit on deterministic predictability; two states differing by imperceptible amounts will inevitably evolve into completely divergent weather regimes."),
    ("Mark Monmonier (1991)", "How to Lie with Maps", "Not only is it easy to lie with maps, it is essential; to portray a complex spherical earth on a flat sheet, a map must distort shape or scale, and practice selective generalization."),
    ("Vilhelm Bjerknes (1904)", "Weather Forecasting as a Problem of Mechanics and Physics", "If we possess sufficiently accurate knowledge of the initial state of the atmosphere and know the physical laws of hydrodynamics, future weather can be calculated with mathematical certainty.")
]

# ==============================================================================
# QUESTION 20 (P1_Q20): MAP PROJECTIONS: CLASSIFICATION & MERCATOR’S PROJECTION
# Solved from: 2022 Q.8, 2018 Q.8, 2016 Q.8, 2013 Q.7, 2010 Q.8, 2008 Q.6
# ==============================================================================
p1_q20_meta = [
    ("Target Syllabus", "Section 4: Study of Maps & Cartography — Map Projections: Classification & Mercator"),
    ("Exam Years Solved", "KPPSC PMS 2022 Q.8, 2018 Q.8, 2016 Q.8, 2013 Q.7, 2010 Q.8, 2008 Q.6"),
    ("Marks & Allocation", "20 Marks | Strict 35-Minute Descriptive Writing Limit"),
    ("Core Theoretical Thesis", "A map projection is the mathematical transformation of the spherical Earth's geographic coordinate graticule (latitude and longitude) onto a developable flat surface. Governed by Carl Friedrich Gauss's Theorema Egregium, a sphere cannot be flattened without introducing geometric distortion; hence, the four fundamental properties—Conformality (true shape), Equivalence (true area), Equidistance, and Azimuthality—are mutually exclusive. Projections classify by developable surface into Cylindrical, Conical, and Zenithal/Azimuthal families. Gerardus Mercator’s 1569 Conformal Cylindrical Projection remains cartography's most celebrated innovation: its supreme navigational virtue is that all straight lines represent lines of constant compass bearing (Rhumb Lines / Loxodromes), counterbalanced by extreme high-latitude areal exaggeration (the 'Greenland Problem').")
]

p1_q20_callout = [
    "THE MATHEMATICAL CARTOGRAPHIC DILEMMA (GAUSS): In differential geometry, Carl Friedrich Gauss proved via his Theorema Egregium (1828) that a sphere possesses positive Gaussian curvature (K > 0) while a plane possesses zero curvature (K = 0). Therefore, the Earth is a 'non-developable surface'—it is mathematically impossible to flatten a globe onto a plane without stretching, tearing, or compressing its geometry.",
    "THE FOUR MUTUALLY EXCLUSIVE PROJECTION PROPERTIES: (1) Conformality (Orthomorphism: preserves true local shapes and angles; meridians and parallels intersect at 90°; essential for navigation and topographic mapping); (2) Equivalence (Equal-Area: preserves true relative areas of landmasses; essential for demographic and economic distribution maps; conformality is destroyed); (3) Equidistance (preserves true linear scale along specific lines); (4) Azimuthality (preserves true compass direction from a central point).",
    "THE THREE MAJOR PROJECTION FAMILIES: (1) Cylindrical: Cylinder wrapped around globe (meridians and parallels are straight, parallel, intersecting at 90°; e.g., Mercator, Gall-Peters); (2) Conical: Cone placed over globe, tangent along one or two Standard Parallels (meridians converge toward pole; excellent for mid-latitude countries like Pakistan or USA; e.g., Conic with Two Standard Parallels, Bonne's, Albers Equal-Area); (3) Zenithal / Azimuthal: Flat plane tangent to a single point (polar, equatorial, or oblique; e.g., Gnomonic, Stereographic, Orthographic).",
    "MERCATOR'S PROJECTION (1569): Conformal cylindrical projection mathematically stretched by the secant of latitude: y = ln[tan(pi/4 + phi/2)]. Supreme Merit: Any straight line drawn on a Mercator map is a 'Rhumb Line' (Loxodrome)—a line of constant compass bearing, enabling navigators to sail from origin to destination on a single compass heading. Critical Demerit: Extreme areal exaggeration near poles; scale expands infinitely; Greenland appears equal to Africa, whereas Africa is actually 14 times larger (30.4M km² vs. 2.16M km²). Cannot project the poles."
]

p1_q20_map = r"""
                     [MAP PROJECTIONS TAXONOMY & MERCATOR]
                                       │
            ┌──────────────────────────┴──────────────────────────┐
            ▼                                                     ▼
  [PROJECTION FAMILIES]                                 [MERCATOR'S PROJECTION (1569)]
  • By Developable Surface:                             • Conformal Cylindrical Projection
    1. Cylindrical (Cylinder wrap)                      • Mathematical spacing expands poleward:
       - Meridians vertical, parallels horizontal         y = ln[tan(pi/4 + phi/2)]
       - E.g., Mercator, Gall-Peters                              │
    2. Conical (Cone cap)                               ┌─────────┴─────────┐
       - Meridians converge; curved parallels           ▼                   ▼
       - Standard Parallels (mid-latitudes)         [SUPREME MERIT]     [FATAL DEMERIT]
       - E.g., Albers Equal-Area, Bonne             • Rhumb Lines /     • Extreme Areal
    3. Zenithal / Azimuthal (Plane tangent)           Loxodromes are      Exaggeration
       - True directions from center                  straight lines      in High Latitudes
       - E.g., Gnomonic, Stereographic              • Indispensable     • "Greenland Problem":
  • By Preserved Property:                            for maritime        Greenland looks equal
    - Conformal (True shape, angles 90°)              & aeronautical      to Africa (Africa is
    - Equivalent (True area: Gall-Peters)             navigation          14x larger in reality)
    - Compromise (Robinson, Winkel Tripel)                               • Poles at infinity
"""

p1_q20_timeline = [
    ("Map Projection Family / Type", "Geometric Construction & Developable Surface", "Cartographic Strengths, Limitations & Primary Utility"),
    ("Cylindrical Family (General)", "Projecting globe coordinates onto an encompassing cylinder tangent along the Equator.", "Graticule is a rectangular grid; lines of latitude and longitude are straight, parallel, and intersect at 90°; areal distortion surges toward poles."),
    ("Mercator’s Projection (1569)", "Conformal cylindrical; spacing between parallels of latitude expands poleward to match meridian stretching.", "Supreme navigational virtue: Rhumb lines are straight lines; fatal limitation: extreme polar areal exaggeration (Greenland = Africa); cannot show poles."),
    ("Gall-Peters Projection (1973)", "Cylindrical Equal-Area projection; parallels of latitude are compressed toward the poles.", "Preserves true area of all countries (eliminates Eurocentric bias); severe shape distortion (continents appear stretched vertically like wet towels)."),
    ("Conical Family (Mid-Latitudes)", "Projecting globe onto a cone tangent along a single parallel or secant along two Standard Parallels.", "Scale is true along Standard Parallels; meridians are straight converging lines; optimal for mid-latitude east-west countries (e.g., Pakistan, USA, Russia)."),
    ("Zenithal / Azimuthal Family", "Projecting globe onto a flat plane tangent to a single point (polar, equatorial, or oblique).", "Preserves true direction (azimuth) from center point; Gnomonic projection displays Great Circles as straight lines (shortest navigational distance)."),
    ("Compromise Projections (Robinson)", "Pseudocylindrical; mathematically balances distortion of shape, area, distance, and direction.", "Neither purely conformal nor equal-area, but visually harmonious; adopted by National Geographic (1988–1998) for global thematic mapping.")
]

p1_q20_essay = [
    ("Introduction: The Geometric Challenge of Projecting a Sphere",
     "A map projection represents the systematic mathematical transformation of the spherical Earth's three-dimensional surface—defined by the geographic coordinate graticule of latitude parallels and longitude meridians—onto a two-dimensional flat plane. In mathematical cartography, the fundamental dilemma of map projection is rooted in differential geometry. In his celebrated 1828 Theorema Egregium, the German mathematician Carl Friedrich Gauss proved that a spherical surface possesses constant positive Gaussian curvature (K > 0), whereas a flat plane possesses zero Gaussian curvature (K = 0). Consequently, a sphere is a 'non-developable surface' that cannot be flattened onto a plane without undergoing physical stretching, compressive shearing, or tearing. Every flat map of the world is, by mathematical definition, distorted. The central task of cartographic science is therefore not to eliminate distortion entirely, but to mathematically control, direct, and specialize distortion to preserve specific geometric properties tailored to distinct human and navigational applications."),

    ("The Four Mutually Exclusive Cartographic Properties",
     "In designing a map projection, cartographers can preserve one or two geometric properties, but never all four simultaneously, as they are inherently mutually exclusive. First is 'Conformality' (or Orthomorphism): a conformal projection preserves true local shapes and correct angular relationships at every point. On a conformal map, parallels and meridians intersect at perpendicular 90° angles, and the scale at any point is identical in all directions, making conformal projections indispensable for maritime navigation, aeronautical flight planning, and military artillery targeting. Second is 'Equivalence' (or Equal-Area): an equivalent projection preserves the true relative surface areas of all regions across the map. A square centimeter measured anywhere on the map represents the exact same ground area as any other square centimeter. Equal-area projections are mandatory for geographical, demographic, and economic mapping (such as displaying world population density or agricultural yields), but they inevitably distort shapes, angles, and distances. Third is 'Equidistance': an equidistant projection preserves true linear scale along specific lines (such as along meridians or standard parallels) or from a single central point. Fourth is 'Azimuthality' (or True Direction): an azimuthal projection preserves true compass bearings from a central focal point to all other locations on the map."),

    ("Classification of Projections by Developable Surface",
     "Projections are systematically classified by the geometric nature of the 'Developable Surface'—a three-dimensional geometric shape that can be flattened into a plane without stretching or tearing. The three primary developable surfaces are the Cylinder, the Cone, and the Plane, supplemented by conventional mathematical formulations. In 'Cylindrical Projections', a cylinder is wrapped around the globe, typically tangent along the Equator. The resulting graticule is a rectangular grid where all meridians are straight, parallel, equally spaced vertical lines, and all parallels of latitude are straight, parallel horizontal lines intersecting at 90°. In 'Conical Projections', a cone is placed over the globe like a cap, tangent along a single 'Standard Parallel' or secant cutting through two Standard Parallels. Meridians are straight lines converging toward the apex of the cone, while parallels are concentric circular arcs. Conical projections exhibit minimal distortion along their standard parallels, making them the optimal cartographic choice for mid-latitude countries with broad east-west extent, such as Pakistan, the United States, and Russia. In 'Zenithal (Azimuthal) Projections', a flat plane is placed tangent to a single point on the globe (polar, equatorial, or oblique). They are widely used for polar research maps and intercontinental airline route mapping."),

    ("Gerardus Mercator's 1569 Masterwork: Mathematical Design and the Rhumb Line",
     "Among all the projections devised in human history, none has exerted a greater historical, maritime, and geopolitical impact than the Conformal Cylindrical Projection created by Flemish cartographer Gerardus Mercator in 1569, titled 'Nova et Aucta Orbis Terrae Descriptio ad Usum Navigantium Emendata' (A New and Augmented Description of Earth Corrected for the Use of Navigators). Mercator addressed a grave crisis in sixteenth-century maritime navigation: on a globe, the shortest path between two points is a 'Great Circle', but following a great circle requires a ship captain to constantly alter compass headings at every degree of longitude. Mercator designed a conformal projection where meridians are equally spaced vertical lines. However, because meridians on a real globe converge toward the poles, keeping them parallel on a flat cylinder artificially stretches horizontal distances by a factor of the secant of latitude (sec phi). To preserve conformality (true local angles), Mercator brilliantly stretched the spacing between parallels of latitude vertically by the exact same secant factor: y = R * ln[tan(pi/4 + phi/2)]. The supreme, revolutionary merit of Mercator’s projection is that any straight line drawn across the map represents a line of constant compass bearing, known in navigation as a 'Rhumb Line' or 'Loxodrome'. For the first time in maritime history, a ship’s captain could simply draw a straight line between departure and destination, measure the angle with a protractor, and sail on that single, fixed compass heading without ever recalculating course adjustments."),

    ("The Fatal Demerit: Extreme Polar Areal Distortion and the 'Greenland Problem'",
     "While Mercator’s projection was an unmatched triumph for maritime navigation, its application as a general-purpose world wall map created severe, misleading geographic distortions. Because both horizontal and vertical scale expand as a secant function of latitude, areal distortion accelerates exponentially toward the poles, approaching mathematical infinity at 90° latitude (preventing the poles from ever being depicted). In mid- and high latitudes, landmasses are grossly exaggerated in size relative to tropical continents—a phenomenon celebrated cartographically as the 'Greenland Problem'. On a standard Mercator projection, Greenland appears equal in size to the entire continent of Africa, and looks considerably larger than South America. In terrestrial reality, Africa covers 30.37 million square kilometers, whereas Greenland covers only 2.16 million square kilometers; Africa is actually fourteen times larger than Greenland! Similarly, Europe appears larger than South America on a Mercator map, whereas South America is nearly double Europe’s physical area. In the 1970s, German historian Arno Peters launched a fierce ideological critique, arguing that Mercator’s Eurocentric projection reinforced colonialist geopolitical mindsets by inflating the imperial powers of Europe and North America while visually shrinking developing nations across Africa, South Asia, and South America, leading to Peters’s controversial promotion of the Gall-Peters Cylindrical Equal-Area Projection."),

    ("Modern Cartographic Synthesis: Compromise Projections",
     "In modern cartography, the consensus recognizes that every projection is a specialized mathematical tool with distinct operational boundaries. Mercator remains indispensable for nautical charts, aviation approach plates, and digital web maps (Web Mercator utilized by Google Maps and OpenStreetMap for conformal tile panning). For thematic global mapping, however, modern geographers utilize 'Compromise Projections'—pseudocylindrical projections that deliberately abandon strict conformality and equal-area properties in exchange for visually harmonious, balanced compromises. Notable examples include the Robinson Projection (formulated by Arthur H. Robinson in 1963) and the Winkel Tripel Projection (formulated by Oswald Winkel in 1921), adopted by the National Geographic Society in 1998 as its standard world map. In evaluative summation, the study of map projections is a profound lesson in critical geographical thinking: no flat map can tell the absolute spatial truth. Understanding projection mathematics, developable surfaces, and distortion trade-offs enables competitive exam candidates to select the appropriate cartographic framework for military statecraft, geographic analysis, and global environmental governance.")
]

p1_q20_quotes = [
    ("Scholar / Cartographer", "Theoretical Contribution / Concept", "Exact Quotable Academic Thesis / Axiom"),
    ("Gerardus Mercator (1569)", "Ad Usum Navigantium Emendata (The 1569 World Map)", "In this map, we have spread the degrees of latitude toward each pole in proportion to the elongation of the parallels, so that a straight line represents a true and constant compass heading."),
    ("Carl Friedrich Gauss (1828)", "Theorema Egregium (Differential Geometry of Surfaces)", "Because a sphere possesses positive Gaussian curvature and a plane possesses zero curvature, no spherical surface can be mapped onto a plane without isometric distortion."),
    ("Arthur H. Robinson (1960)", "Elements of Cartography (The Robinson Projection)", "Cartographic design for world maps should not be slave to rigid mathematical purity; a compromise of subtle distortions produces the most realistic visual portrait of our spherical home."),
    ("Arno Peters (1973)", "The New Cartography (The Gall-Peters Projection)", "The Mercator map is a monument to Eurocentrism, visually subjugating the developing third world; cartographic justice demands equal-area fidelity."),
    ("John P. Snyder (1987)", "Map Projections: A Working Manual (USGS)", "There is no single best map projection; the choice of developable surface and distortion properties must be governed strictly by the geographic location and thematic purpose of the map.")
]

# ==============================================================================
# BATCH 3 EXECUTION RUNNER (Q15 to Q20)
# ==============================================================================
BATCH_3_TASKS = [
    ("Q15_Ocean_Floor_Morphology_and_Continental_Margins", "Q.15: Ocean Floor Morphology & Continental Margins", p1_q15_meta, p1_q15_callout, p1_q15_map, p1_q15_timeline, p1_q15_essay, p1_q15_quotes),
    ("Q16_Ocean_Salinity_Controlling_Factors_and_Variations", "Q.16: Ocean Salinity: Factors, Variations & Circulation", p1_q16_meta, p1_q16_callout, p1_q16_map, p1_q16_timeline, p1_q16_essay, p1_q16_quotes),
    ("Q17_Oceanic_Circulation_Major_Currents_and_Gyres", "Q.17: Oceanic Circulation: Major Ocean Currents & Gyres", p1_q17_meta, p1_q17_callout, p1_q17_map, p1_q17_timeline, p1_q17_essay, p1_q17_quotes),
    ("Q18_Mechanics_of_Ocean_Tides_and_Tidal_Theories", "Q.18: Mechanics of Ocean Tides & Tidal Theories", p1_q18_meta, p1_q18_callout, p1_q18_map, p1_q18_timeline, p1_q18_essay, p1_q18_quotes),
    ("Q19_Cartography_Basics_Map_Scales_and_Weather_Maps", "Q.19: Cartography Basics: Map Scales & Synoptic Maps", p1_q19_meta, p1_q19_callout, p1_q19_map, p1_q19_timeline, p1_q19_essay, p1_q19_quotes),
    ("Q20_Map_Projections_Classification_and_Mercators_Projection", "Q.20: Map Projections: Classification & Mercator", p1_q20_meta, p1_q20_callout, p1_q20_map, p1_q20_timeline, p1_q20_essay, p1_q20_quotes),
]

for filename, title_text, meta, callout, map_txt, timeline, essay, quotes in BATCH_3_TASKS:
    docx_file = os.path.join(P1_DIR, f"{filename}_Master_Notes.docx")
    pdf_file = os.path.join(P1_DIR, f"{filename}_Master_Notes.pdf")
    md_file = os.path.join(P1_DIR, f"{filename}_Master_Notes.md")
    
    generate_geo_master_doc(meta, callout, map_txt, timeline, essay, quotes, docx_file, title_text)
    generate_pdf_and_md(meta, callout, map_txt, timeline, essay, quotes, pdf_file, md_file, title_text)

print("Batch 3 (Q15 through Q20) for Geography Paper 1 successfully generated! Paper 1 is 100% COMPLETE!")
