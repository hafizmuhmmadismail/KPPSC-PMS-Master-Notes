import os
import re
import sys
import subprocess
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

EDGE_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
WORDCRAFT_CLI = r"H:\My Drive\PMS KPK 2022 notes\apps_toolkit\tools\wordcraft\wordcraft-0.8.0-windows-x64-portable\wordcraft-cli.exe"
GEO_P2_DIR = r"H:\My Drive\PMS KPK 2022 notes\3. Geography Paper 1 and Paper 2\Paper 2 - Human Economic and Regional Geography"

GEO_P2_EXAM_QUESTIONS_LINKED_MAP = {
    1: (
        "• 2022 Q.8(a): \"Write comprehensive short notes on: Environmental Determinism vs. Possibilism. (10 Marks)\"\n"
        "• 2018 Q.1: \"Discuss the concepts of Environmental Determinism and Possibilism with historical examples. How has modern technology altered this relationship? (20 Marks)\"\n"
        "• 2013 Q.8(c): \"Write comprehensive notes on: Man-Environment Relationship. (10 Marks)\"\n"
        "• 2010 Q.1: \"Differentiate between Environmental Determinism and Possibilism. Who were the key propagators? Discuss concept evolution. (20 Marks)\"\n"
        "• 2008 Q.1: \"Explain the concept of Environmental Determinism. How does the physical environment influence human activities? (20 Marks)\""
    ),
    2: (
        "• 2018 Q.2: \"Highlight the scope and status of Human Geography with special reference to its application for sustainable human development. (20 Marks)\""
    ),
    3: (
        "• 2022 Q.1: \"Discuss the major factors responsible for the distribution and density of world population with suitable examples. (20 Marks)\"\n"
        "• 2018 Q.4: \"What are the factors responsible for the uneven distribution of population with special reference to Pakistan? (20 Marks)\"\n"
        "• 2013 Q.6: \"To what extent can the major features of world population distribution be understood in terms of physical environment? Explain Malthusian Theory. (20 Marks)\"\n"
        "• 2010 Q.2: \"\\\"Population of the world is unevenly distributed.\\\" Elaborate with reference to physical and human controlling factors. (20 Marks)\"\n"
        "• 2008 Q.2: \"What is meant by Population Density? Write down areas of major world population concentrations. (20 Marks)\""
    ),
    4: (
        "• 2016 Q.1: \"Narrate the story of world population growth with the help of the J-shaped curve and present your views on future demographic projections. (20 Marks)\"\n"
        "• 2013 Q.6: \"To what extent can the major features of world population distribution be understood in terms of physical environment? Explain Malthusian Theory. (20 Marks)\""
    ),
    5: (
        "• 2025 Q.3: \"Analyze the major threats posed by overpopulation at global and national scales. Critically examine the demographic pressures facing Pakistan and their implications for economic development, resource management, and social stability. (20 Marks)\"\n"
        "• 2018 Q.4: \"What are the factors responsible for the uneven distribution of population with special reference to Pakistan? (20 Marks)\""
    ),
    6: (
        "• 2018 Q.8(d): \"Write short note on: Trends and patterns of population migration in Pakistan. (10 Marks)\"\n"
        "• 2016 Q.7: \"Explain the causes and socio-economic consequences of international migration (push-pull factors, remittances, brain drain). (20 Marks)\"\n"
        "• 2013 Q.5: \"Elaborate pioneer settlement patterns. \\\"The pattern of settlement characterizing human use of the earth is strongly related to migration.\\\" Discuss. (20 Marks)\"\n"
        "• 2010 Q.8(a): \"Write note on: Rural-urban migration. (10 Marks)\""
    ),
    7: (
        "• 2013 Q.5: \"Elaborate pioneer settlement patterns. \\\"The pattern of settlement characterizing human use of the earth is strongly related to migration.\\\" Discuss. (20 Marks)\"\n"
        "• 2010 Q.4: \"What yardsticks are used for defining Rural and Urban Settlements? Describe locational factors of rural settlements. (20 Marks)\""
    ),
    8: (
        "• 2022 Q.8(d): \"Write comprehensive short note on: Christaller’s Central Place Theory. (10 Marks)\"\n"
        "• 2013 Q.8(a): \"Write comprehensive note on: Christaller’s Central Place Theory. (10 Marks)\""
    ),
    9: (
        "• 2022 Q.2: \"Examine the process of rapid urbanization in Pakistan. Discuss the resultant socio-economic and environmental problems and suggest practical policy remedies. (20 Marks)\"\n"
        "• 2018 Q.3: \"What is urbanization? Discuss the urbanization process with special reference to Pakistan's major metropolitan centers. (20 Marks)\"\n"
        "• 2016 Q.2: \"Point out the major problems resulting from rapid urbanization in Pakistan and suggest workable remedies to overcome these crises. (20 Marks)\"\n"
        "• 2013 Q.8(b): \"Write comprehensive note on: Urbanization. (10 Marks)\"\n"
        "• 2010 Q.3: \"Define urbanization and discuss major problems associated with urbanization in both developing and developed countries. (20 Marks)\"\n"
        "• 2010 Q.8(b): \"Write note on: Mega-cities. (10 Marks)\""
    ),
    10: (
        "• 2025 Q.1: \"Explain the concept of bridging the “green” and “brown” agendas in sustainable urbanization. Evaluate how integrated environmental and infrastructural planning can address the challenges of rapid urban growth in developing countries. (20 Marks)\"\n"
        "• 2010 Q.8(b): \"Write note on: Mega-cities. (10 Marks)\""
    ),
    11: (
        "• 2016 Q.8(b): \"Write comprehensive note on: World's linguistic and spatial language patterns. (10 Marks)\"\n"
        "• 2008 Q.4: \"\\\"Despite modernization, religion still dominates human behavior.\\\" Geographical distribution of Islam and Christianity. (20 Marks)\""
    ),
    12: (
        "• 2022 Q.3: \"Define and classify economic activities into primary, secondary, and tertiary sectors. Discuss how economic development alters occupational structures globally. (20 Marks)\"\n"
        "• 2016 Q.3: \"Classify and explain economic activities by providing suitable examples from various regions of the world. (20 Marks)\"\n"
        "• 2013 Q.1: \"What is Economic Geography? Present a comparative survey of secondary and tertiary economic activities with occupational examples. (20 Marks)\"\n"
        "• 2010 Q.5: \"Define economic activities and discuss the evolution of secondary and tertiary activities with changing global consumption patterns. (20 Marks)\"\n"
        "• 2008 Q.3: \"Define Economic Geography. Detailed account of human economic activities related to production and services. (20 Marks)\""
    ),
    13: (
        "• 2025 Q.2: \"Discuss the physical, technological, and socio-economic factors that influence agricultural patterns around the world. Assess the significance of mixed farming as an adaptive agricultural system in modern economies. (20 Marks)\"\n"
        "• 2013 Q.4: \"Differentiate between intensive and extensive agriculture. Explain world agriculture under specializations, systems, and problems. (20 Marks)\"\n"
        "• 2008 Q.7(i): \"Write short note on: Mixed Farming. (10 Marks)\""
    ),
    14: (
        "• 2016 Q.5: \"Among agricultural commodities, which cereal and cash crops dominate world trade? Give an account of commercial wheat production. (20 Marks)\"\n"
        "• 2010 Q.6: \"\\\"Rice feeds more people and needs more labour force compared to other cereal crops.\\\" Describe geographical factors for rice cultivation in South Asia. (20 Marks)\"\n"
        "• 2008 Q.5: \"Distribution and production of cotton in the world with special reference to South Asia. (20 Marks)\""
    ),
    15: (
        "• 2025 Q.4: \"Describe the global spatial distribution of tropical savanna forests. Evaluate their economic significance and discuss the environmental challenges associated with resource extraction in these regions. (20 Marks)\"\n"
        "• 2016 Q.8(d): \"Write comprehensive note on: Environmental and economic benefits of forest resources. (10 Marks)\""
    ),
    16: (
        "• 2022 Q.8(c): \"Write comprehensive short note on: Major commercial fishing grounds of the world. (10 Marks)\"\n"
        "• 2016 Q.8(a): \"Write comprehensive note on: Commercial fishing regions of the world. (10 Marks)\"\n"
        "• 2010 Q.8(c): \"Write note on: Major fishing grounds. (10 Marks)\""
    ),
    17: (
        "• 2016 Q.6: \"\\\"The spatial distribution of power resources influences world geopolitics and economic supremacy.\\\" Discuss. (20 Marks)\"\n"
        "• 2013 Q.3: \"Discuss methods of conservation of raw material resources. Critically evaluate the potential for alternative energy sources. (20 Marks)\"\n"
        "• 2010 Q.7: \"\\\"Power resources are playing a key role in economic development.\\\" Describe the production and trade of Crude Oil in the world. (20 Marks)\"\n"
        "• 2008 Q.7(iv): \"Write short note on: Coal deposits of the USA. (10 Marks)\""
    ),
    18: (
        "• 2022 Q.4: \"Examine Alfred Weber’s Least Cost Theory of industrial location with reference to the Iron and Steel industry. (20 Marks)\"\n"
        "• 2013 Q.2: \"\\\"A large number of industrialized economies are located in the Western Hemisphere.\\\" Discuss. Examine the implications of Weber’s Model of industrial location. (20 Marks)\"\n"
        "• 2008 Q.6: \"Details about the Iron and Steel Industry of the USA. (20 Marks)\""
    ),
    19: (
        "• 2025 Q.5: \"Explain the geography of development and critically assess the factors responsible for regional disparities between developed and underdeveloped countries. How do these disparities influence global economic and political relations? (20 Marks)\"\n"
        "• 2025 Q.8: \"Explain the major economic and socio-demographic indicators used to assess national development. (20 Marks)\"\n"
        "• 2018 Q.8(c): \"Write short note on: Poverty alleviation strategies through spatial planning. (10 Marks)\""
    ),
    20: (
        "• 2016 Q.4: \"\\\"Pakistan is a resource-rich country.\\\" Critically examine this statement with reference to physical, mineral, and agricultural resources. (20 Marks)\"\n"
        "• 2008 Q.7(ii): \"Write short note on: The Indus Plain. (10 Marks)\""
    ),
    21: (
        "• 2025 Q.6: \"Evaluate the assertion that Pakistan is fundamentally an agro-based economy. Discuss the structural characteristics, strengths, weaknesses, and transformation potential of Pakistan's agricultural sector. (20 Marks)\"\n"
        "• 2022 Q.6: \"Analyze the agricultural sector of Pakistan: Major crops, canal irrigation network, and the emerging challenges of water scarcity and land degradation. (20 Marks)\"\n"
        "• 2018 Q.6: \"Discuss the main economic activities of Pakistan with special focus on the agricultural sector and rural livelihood. (20 Marks)\""
    ),
    22: (
        "• 2025 Q.7: \"Discuss the spatial distribution, raw material base, and locational factors of chemical industries in Pakistan. Evaluate their role in national economic development and identify the major constraints hindering their expansion. (20 Marks)\""
    ),
    23: (
        "• 2022 Q.8(b): \"Write comprehensive short note on: Structural shifts in the foreign trade of Pakistan. (10 Marks)\"\n"
        "• 2018 Q.8(a): \"Write short note on: Actionable policy measures for increasing Pakistan's export earnings. (10 Marks)\"\n"
        "• 2013 Q.7: \"Highlight the patterns of Foreign Trade of Pakistan. What major structural changes have taken place in exports and imports? (20 Marks)\"\n"
        "• 2010 Q.8(d): \"Write note on: Trade balance of Pakistan. (10 Marks)\"\n"
        "• 2008 Q.7(iii): \"Write short note on: Trade of Pakistan. (10 Marks)\""
    ),
    24: (
        "• 2022 Q.5: \"\\\"There is a paradigm shift from Geopolitics to Geoeconomics.\\\" In your opinion, how should Pakistan position itself to accelerate national economic development? (20 Marks)\"\n"
        "• 2022 Q.7: \"Critically evaluate the China-Pakistan Economic Corridor (CPEC) as a game-changer for regional transit, industrial modernization, and trade expansion. (20 Marks)\"\n"
        "• 2018 Q.5: \"There is a \\\"paradigm shift from Geo-political to Geo-Economic.\\\" In your opinion, which orientation should Pakistan prioritize for long-term prosperity? (20 Marks)\"\n"
        "• 2018 Q.7: \"In your opinion, how can Pakistan derive maximum socio-economic benefit from CPEC, the flagship project of the Belt and Road Initiative (BRI)? (20 Marks)\"\n"
        "• 2018 Q.8(b): \"Write short note on: Strategic and economic relations between Pakistan and China. (10 Marks)\""
    )
}

def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
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
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(30, 41, 59)

def generate_geo_p2_master_doc(meta_data, callout_data, map_text, timeline_data, essay_sections, quotes_data, out_path, paper_sub_title):
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
    run_main = p_title.add_run("KPPSC PMS • GEOGRAPHY (PAPER-II)")
    run_main.font.name = "Arial"
    run_main.font.size = Pt(12)
    run_main.font.bold = True
    run_main.font.color.rgb = RGBColor(14, 43, 80)
    
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(10)
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = p_sub.add_run(f"EXECUTIVE MASTER EXAMINATION NOTES • {paper_sub_title}")
    run_sub.font.name = "Arial"
    run_sub.font.size = Pt(10)
    run_sub.font.color.rgb = RGBColor(27, 73, 101)
    
    # Metadata Table
    tbl_meta = doc.add_table(rows=len(meta_data), cols=2)
    tbl_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_meta.autofit = False
    
    for idx, (k, v) in enumerate(meta_data):
        row = tbl_meta.rows[idx]
        cell_k, cell_v = row.cells[0], row.cells[1]
        cell_k.width = Inches(2.2)
        cell_v.width = Inches(4.3)
        
        set_cell_margins(cell_k, top=60, bottom=60, left=100, right=100)
        set_cell_margins(cell_v, top=60, bottom=60, left=100, right=100)
        
        p_k = cell_k.paragraphs[0]
        p_k.paragraph_format.space_before = Pt(0)
        p_k.paragraph_format.space_after = Pt(0)
        r_k = p_k.add_run(k)
        r_k.font.name = "Arial"
        r_k.font.size = Pt(9)
        r_k.font.bold = True
        
        lines = v.split("\n")
        for l_idx, line in enumerate(lines):
            if l_idx == 0:
                p_v = cell_v.paragraphs[0]
            else:
                p_v = cell_v.add_paragraph()
            p_v.paragraph_format.space_before = Pt(1 if len(lines) > 2 else 0)
            p_v.paragraph_format.space_after = Pt(1 if len(lines) > 2 else 0)
            r_v = p_v.add_run(line)
            r_v.font.name = "Georgia"
            r_v.font.size = Pt(8.5 if len(lines) > 2 else 9)
            if idx == 0:
                r_v.font.color.rgb = RGBColor(255, 255, 255)
                r_v.font.bold = True
            else:
                r_v.font.color.rgb = RGBColor(15, 23, 42)
                
        if idx == 0:
            set_cell_background(cell_k, "0E2B50")
            set_cell_background(cell_v, "0E2B50")
            r_k.font.color.rgb = RGBColor(255, 255, 255)
        else:
            bg = "F8FAFC" if idx % 2 == 1 else "FFFFFF"
            set_cell_background(cell_k, bg)
            set_cell_background(cell_v, bg)
            r_k.font.color.rgb = RGBColor(14, 43, 80) if "EXAM QUESTIONS" in k else RGBColor(30, 41, 59)
            
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    add_callout(doc, callout_data, "EXAM ESSENTIAL TAKEAWAY")
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    
    # 1. Conceptual Mind Map
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
    
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    # 2. Analytical Dimensions Matrix
    p_htl = doc.add_paragraph()
    p_htl.paragraph_format.space_before = Pt(8)
    p_htl.paragraph_format.space_after = Pt(4)
    r_htl = p_htl.add_run("2. ANALYTICAL DIMENSIONS & SOCIO-SPATIAL/THEORETICAL MATRIX")
    r_htl.font.name = "Arial"
    r_htl.font.size = Pt(11)
    r_htl.font.bold = True
    r_htl.font.color.rgb = RGBColor(14, 43, 80)
    
    num_cols = len(timeline_data[0])
    tbl_tl = doc.add_table(rows=len(timeline_data), cols=num_cols)
    tbl_tl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_tl.autofit = False
    
    if num_cols == 3:
        col_widths = [Inches(1.8), Inches(2.2), Inches(2.5)]
    elif num_cols == 4:
        col_widths = [Inches(1.4), Inches(1.6), Inches(1.8), Inches(1.7)]
    else:
        w = Inches(6.5 / num_cols)
        col_widths = [w] * num_cols
        
    for r_idx, row_data in enumerate(timeline_data):
        row = tbl_tl.rows[r_idx]
        is_hdr = (r_idx == 0)
        bg = "0E2B50" if is_hdr else ("EDF2F7" if r_idx % 2 == 1 else "FFFFFF")
        fg = RGBColor(255, 255, 255) if is_hdr else RGBColor(30, 41, 59)
        
        for c_idx, text in enumerate(row_data):
            c = row.cells[c_idx]
            c.width = col_widths[c_idx]
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
    
    # 3. Comprehensive Master Examination Essay
    p_he = doc.add_paragraph()
    p_he.paragraph_format.space_before = Pt(8)
    p_he.paragraph_format.space_after = Pt(4)
    r_he = p_he.add_run("3. 20-MARK MASTER MODEL ESSAY (KPPSC PMS BENCHMARK)")
    r_he.font.name = "Arial"
    r_he.font.size = Pt(11)
    r_he.font.bold = True
    r_he.font.color.rgb = RGBColor(14, 43, 80)
    
    for heading, text in essay_sections:
        p_sh = doc.add_paragraph()
        p_sh.paragraph_format.space_before = Pt(6)
        p_sh.paragraph_format.space_after = Pt(2)
        r_sh = p_sh.add_run(heading)
        r_sh.font.name = "Arial"
        r_sh.font.size = Pt(10)
        r_sh.font.bold = True
        r_sh.font.color.rgb = RGBColor(27, 73, 101)
        
        p_st = doc.add_paragraph()
        p_st.paragraph_format.space_before = Pt(0)
        p_st.paragraph_format.space_after = Pt(4)
        p_st.paragraph_format.line_spacing = 1.15
        p_st.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r_st = p_st.add_run(text)
        r_st.font.name = "Georgia"
        r_st.font.size = Pt(9.5)
        r_st.font.color.rgb = RGBColor(30, 41, 59)
        
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    # 4. Authoritative Lexicon & Citations
    p_hq = doc.add_paragraph()
    p_hq.paragraph_format.space_before = Pt(8)
    p_hq.paragraph_format.space_after = Pt(4)
    r_hq = p_hq.add_run("4. SCHOLARLY PERSPECTIVES & GEOGRAPHICAL AXIOMS TABLE")
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
    print(f"Native Wordcraft Docx generated: {out_path}")


def generate_pdf_and_md(meta_data, callout_data, map_text, timeline_data, essay_sections, quotes_data, pdf_path, md_path, title):
    # 1. Generate Markdown
    md_content = f"# {title}\n\n"
    md_content += "## KPPSC PMS Examination — Master Study Notes (Paper-II: Human, Economic & Regional Geography)\n\n"
    md_content += "| Field | Value |\n|---|---|\n"
    
    exam_questions_linked_raw = ""
    for k, v in meta_data:
        if "EXAM QUESTIONS" in k.upper():
            exam_questions_linked_raw = v
            val_clean = v.replace("\n", "<br>")
        else:
            val_clean = v.replace("\n", " ")
        md_content += f"| **{k}** | {val_clean} |\n"
        
    md_content += "\n---\n\n"
    
    if exam_questions_linked_raw:
        md_content += "### 🎯 EXAM QUESTIONS LINKED\n\n"
        for q_line in exam_questions_linked_raw.strip().split("\n"):
            md_content += f"> {q_line}\n>\n"
        md_content += "\n---\n\n"
        
    md_content += "### 📌 EXAM ESSENTIAL TAKEAWAY\n"
    for line in callout_data:
        md_content += f"> {line}\n"
    md_content += "\n---\n\n"
    
    md_content += "### 1. Conceptual Mind Map & Thematic Roadmap\n```\n"
    md_content += f"{map_text}\n"
    md_content += "```\n\n---\n\n"
    
    md_content += "### 2. Analytical Dimensions & Socio-Spatial/Theoretical Matrix\n\n"
    headers = timeline_data[0]
    md_content += f"| {' | '.join(headers)} |\n|{'|'.join(['---'] * len(headers))}|\n"
    for row in timeline_data[1:]:
        md_content += f"| {' | '.join([f'**{row[0]}**'] + list(row[1:]))} |\n"
    md_content += "\n---\n\n"
    
    md_content += "### 3. 20-Mark Master Model Essay\n\n"
    for heading, text in essay_sections:
        md_content += f"#### {heading}\n{text}\n\n"
    md_content += "---\n\n"
    
    md_content += "### 4. Scholarly Perspectives & Geographical Axioms Table\n\n"
    q_hdr = quotes_data[0]
    md_content += f"| {q_hdr[0]} | {q_hdr[1]} | {q_hdr[2]} |\n|---|---|---|\n"
    for row in quotes_data[1:]:
        md_content += f"| **{row[0]}** | {row[1]} | *\"{row[2]}\"* |\n"
        
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(md_content)
        
    # 2. Generate PDF via HTML
    html_meta_rows = ""
    for idx, (k, v) in enumerate(meta_data):
        if idx == 0:
            html_meta_rows += f"""<tr style="background-color: #0E2B50; color: white;">
                <td style="padding: 8px 12px; font-weight: bold; width: 28%; font-family: Arial, sans-serif; font-size: 9pt;">{k}</td>
                <td style="padding: 8px 12px; font-weight: bold; width: 72%; font-family: Arial, sans-serif; font-size: 9pt;">{v}</td>
            </tr>"""
        else:
            bg = "#F8FAFC" if idx % 2 == 1 else "#FFFFFF"
            val_formatted = "<br>".join(v.split("\n"))
            html_meta_rows += f"""<tr style="background-color: {bg};">
                <td style="padding: 6px 10px; font-weight: bold; color: #0E2B50; width: 28%; font-family: Arial, sans-serif; font-size: 8.5pt;">{k}</td>
                <td style="padding: 6px 10px; color: #1E293B; width: 72%; font-family: Georgia, serif; font-size: 8.5pt;">{val_formatted}</td>
            </tr>"""

    html_exam_questions = ""
    if exam_questions_linked_raw:
        html_exam_questions = f"""
        <div style="background-color: #f1f5f9; border-left: 4px solid #0E2B50; padding: 12px 16px; margin: 16px 0; border-radius: 4px;">
            <h4 style="margin: 0 0 8px 0; color: #0E2B50; font-size: 11pt; font-family: Arial, sans-serif;">🎯 EXAM QUESTIONS LINKED</h4>
            <div style="font-family: Georgia, serif; font-size: 9pt; color: #1e293b; line-height: 1.5;">
                {"<br><br>".join([f"<strong>{line.split(':', 1)[0]}:</strong>" + (f" {line.split(':', 1)[1]}" if ':' in line else '') for line in exam_questions_linked_raw.strip().split('\n')])}
            </div>
        </div>
        """

    html_callout_items = "".join([f"<p style='margin: 4px 0; font-family: Georgia, serif; font-size: 9.5pt; color: #1E293B;'>{line}</p>" for line in callout_data])
    
    html_tl_hdr = "".join(f"<th>{h}</th>" for h in timeline_data[0])
    html_tl_rows = "".join("<tr>" + "".join(f"<td>{'<strong>' + c + '</strong>' if idx == 0 else c}</td>" for idx, c in enumerate(r)) + "</tr>" for r in timeline_data[1:])

    html_essay = ""
    for heading, text in essay_sections:
        html_essay += f"""
        <h3 style="color: #1B4965; font-family: Arial, sans-serif; font-size: 10pt; margin: 10px 0 4px 0; border-bottom: 1px solid #E2E8F0; padding-bottom: 2px;">{heading}</h3>
        <p style="font-family: Georgia, serif; font-size: 9.5pt; line-height: 1.5; color: #1E293B; margin: 0 0 8px 0; text-align: justify;">{text}</p>
        """

    html_quotes_hdr = "".join(f"<th>{h}</th>" for h in quotes_data[0])
    html_quotes_rows = "".join(f"<tr><td><strong>{r[0]}</strong></td><td>{r[1]}</td><td><em>\"{r[2]}\"</em></td></tr>" for r in quotes_data[1:])

    html_content = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
    @page {{
        size: A4;
        margin: 18mm 16mm 18mm 16mm;
        @bottom-right {{
            content: counter(page);
            font-family: Arial, sans-serif;
            font-size: 8pt;
            color: #64748B;
        }}
    }}
    body {{
        font-family: Georgia, serif;
        color: #1E293B;
        margin: 0;
        padding: 0;
        line-height: 1.4;
    }}
    h1 {{
        color: #0E2B50;
        font-family: Arial, sans-serif;
        font-size: 14pt;
        text-align: center;
        margin: 0 0 2px 0;
        font-weight: bold;
    }}
    h2 {{
        color: #1B4965;
        font-family: Arial, sans-serif;
        font-size: 10pt;
        text-align: center;
        margin: 0 0 12px 0;
        font-weight: normal;
    }}
    h2.section-header {{
        color: #0E2B50;
        font-family: Arial, sans-serif;
        font-size: 11pt;
        margin: 14px 0 6px 0;
        border-bottom: 2px solid #0E2B50;
        padding-bottom: 3px;
        text-align: left;
        font-weight: bold;
    }}
    table {{
        width: 100%;
        border-collapse: collapse;
        margin: 8px 0;
    }}
    th {{
        background-color: #0E2B50;
        color: white;
        font-family: Arial, sans-serif;
        font-size: 8.5pt;
        padding: 6px 8px;
        text-align: left;
    }}
    td {{
        border: 1px solid #CBD5E1;
        padding: 6px 8px;
        font-size: 8.5pt;
        vertical-align: top;
    }}
    tr:nth-child(even) td {{
        background-color: #F8FAFC;
    }}
    .callout {{
        background-color: #F0F4F8;
        border-left: 4px solid #0E2B50;
        padding: 8px 14px;
        margin: 10px 0;
    }}
    .callout-title {{
        color: #0E2B50;
        font-family: Arial, sans-serif;
        font-size: 9.5pt;
        font-weight: bold;
        margin: 0 0 4px 0;
    }}
    pre {{
        background-color: #F8FAFC;
        border: 1px solid #CBD5E1;
        padding: 10px;
        font-family: Consolas, monospace;
        font-size: 7.5pt;
        line-height: 1.25;
        overflow-x: auto;
        white-space: pre-wrap;
    }}
</style>
</head>
<body>

<h1>KPPSC PMS • GEOGRAPHY (PAPER-II)</h1>
<h2>EXECUTIVE MASTER EXAMINATION NOTES • {title.split('—')[-1].strip() if '—' in title else title}</h2>

<table>
    {html_meta_rows}
</table>

{html_exam_questions}

<div class="callout">
    <div class="callout-title">📌 EXAM ESSENTIAL TAKEAWAY</div>
    {html_callout_items}
</div>

<h2 class="section-header">1. CONCEPTUAL MIND MAP & THEMATIC ROADMAP</h2>
<pre>{map_text}</pre>

<h2 class="section-header">2. ANALYTICAL DIMENSIONS & SOCIO-SPATIAL/THEORETICAL MATRIX</h2>
<table>
    <thead>
        <tr>{html_tl_hdr}</tr>
    </thead>
    <tbody>
        {html_tl_rows}
    </tbody>
</table>

<h2 class="section-header">3. 20-MARK MASTER MODEL ESSAY (KPPSC PMS BENCHMARK)</h2>
{html_essay}

<h2 class="section-header">4. SCHOLARLY PERSPECTIVES & GEOGRAPHICAL AXIOMS TABLE</h2>
<table>
    <thead>
        <tr>{html_quotes_hdr}</tr>
    </thead>
    <tbody>
        {html_quotes_rows}
    </tbody>
</table>

</body>
</html>"""

    temp_html = md_path.replace(".md", "_temp.html")
    with open(temp_html, "w", encoding="utf-8") as f:
        f.write(html_content)
        
    cmd = [EDGE_PATH, "--headless", "--disable-gpu", "--run-all-compositor-stages-before-draw", f"--print-to-pdf={pdf_path}", temp_html]
    subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if os.path.exists(temp_html):
        os.remove(temp_html)
    print(f"PDF and MD successfully generated: {pdf_path}")


def extract_batch_data(py_path, var_name):
    with open(py_path, 'r', encoding='utf-8') as f:
        code = f.read()
    
    clean_lines = []
    for line in code.splitlines():
        if re.match(r'^(if __name__ == "__main__":|for item in GEO_P2_BATCH)', line):
            break
        clean_lines.append(line)
    clean_code = "\n".join(clean_lines)
    
    ns = {}
    exec(clean_code, ns)
    return ns.get(var_name, [])

def update_and_build_all():
    batch1_file = r"H:\My Drive\PMS KPK 2022 notes\apps_toolkit\build_geo_p2_batch1_notes.py"
    batch2_file = r"H:\My Drive\PMS KPK 2022 notes\apps_toolkit\build_geo_p2_batch2_notes.py"
    batch3_file = r"H:\My Drive\PMS KPK 2022 notes\apps_toolkit\build_geo_p2_batch3_notes.py"
    
    print("Loading Batch 1 data...")
    batch1 = extract_batch_data(batch1_file, "GEO_P2_BATCH1_DATA")
    print(f"Batch 1 loaded: {len(batch1)} items")
    
    print("Loading Batch 2 data...")
    batch2 = extract_batch_data(batch2_file, "GEO_P2_BATCH2_DATA")
    print(f"Batch 2 loaded: {len(batch2)} items")
    
    print("Loading Batch 3 data...")
    batch3 = extract_batch_data(batch3_file, "GEO_P2_BATCH3_DATA")
    print(f"Batch 3 loaded: {len(batch3)} items")
    
    all_data = batch1 + batch2 + batch3
    print(f"Total topics loaded: {len(all_data)}")
    
    count = 0
    for idx, item in enumerate(all_data):
        q_num = idx + 1
        meta = item["meta"]
        callout = item["callout"]
        cmap = item["map"]
        timeline = item["timeline"]
        essay = item["essay"]
        quotes = item["quotes"]
        doc_name = item["doc_name"]
        sub_title = item["sub_title"]
        title = item["title"]
        
        # Build unified executive examination metadata
        new_meta = [
            ("Target Examination", "KPPSC PMS (Provincial Management Service) — Khyber Pakhtunkhwa"),
            ("Paper & Subject", "Geography Paper-II (Human, Economic & Regional Geography — 100 Marks)"),
            ("EXAM QUESTIONS LINKED:", GEO_P2_EXAM_QUESTIONS_LINKED_MAP[q_num])
        ]
        for k, v in meta:
            if k not in ["Exam Years Solved", "Target Examination", "Paper & Subject", "Subject & Paper", "Past Exam Appearances"] and "EXAM QUESTIONS" not in k.upper():
                new_meta.append((k, v))
                
        docx_file = os.path.join(GEO_P2_DIR, f"{doc_name}.docx")
        pdf_file = os.path.join(GEO_P2_DIR, f"{doc_name}.pdf")
        md_file = os.path.join(GEO_P2_DIR, f"{doc_name}.md")
        
        print(f"\n[{q_num}/24] Updating Q{q_num}: {doc_name}...")
        generate_geo_p2_master_doc(
            meta_data=new_meta,
            callout_data=callout,
            map_text=cmap,
            timeline_data=timeline,
            essay_sections=essay,
            quotes_data=quotes,
            out_path=docx_file,
            paper_sub_title=sub_title
        )
        
        full_title = f"KPPSC PMS Geography Paper-II — Q.{q_num}: {title}"
        generate_pdf_and_md(
            meta_data=new_meta,
            callout_data=callout,
            map_text=cmap,
            timeline_data=timeline,
            essay_sections=essay,
            quotes_data=quotes,
            pdf_path=pdf_file,
            md_path=md_file,
            title=full_title
        )
        count += 1

    print(f"\n=======================================================")
    print(f"Successfully generated all {count} Geography Paper-II Master Topics!")
    print(f"=======================================================")

if __name__ == "__main__":
    update_and_build_all()
