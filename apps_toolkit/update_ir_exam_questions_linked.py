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
IR_DIR = r"H:\My Drive\PMS KPK 2022 notes\2. International Relations"

EXAM_QUESTIONS_LINKED_MAP = {
    1: (
        "• 2025 Q.1: \"The rise of China is widely interpreted as a revival of balance of power politics, signaling a global shift toward multipolarity. Analyze the political, strategic and economic factors that are shaping the new international order. (20 Marks)\"\n"
        "• 2022 Q.3: \"The trade war between the US and China has jeopardized the global economic order impacting all regions and states; Pakistan caught between these two global powers is likely to be affected the most. Analyze Pakistan's predicament and suggest policy options to safeguard its national security objectives. (20 Marks)\"\n"
        "• 2022 Q.5: \"The Chinese Belt and Road Initiative (BRI) has global outreach in Asia, Africa, and Europe at the cost of US hegemony. Discuss its main objectives and analyze challenges and prospects of its success. (20 Marks)\"\n"
        "• 2018 Q.6: \"There may be a possible trade war between the USA and China that would rattle certain American businesses and force consumers to pay more for products. Discuss the term \\\"trade war\\\" and elaborate on why Trump was targeting China. (20 Marks)\""
    ),
    2: (
        "• 2025 Q.2: \"Critically evaluate the resilience of Indus Basin Treaty as a cornerstone of hydrological security in South Asia. To what extent can this agreement continue to safeguard water-sharing arrangements from broader geopolitical tension between India and Pakistan, particularly in an era of increasing water scarcity and climate change? (20 Marks)\""
    ),
    3: (
        "• 2025 Q.3: \"Analyze the expansion of BRICS and Shanghai Cooperation Organization (SCO) through the lens of structural realism. Are these organizations best understood as nascent challengers to a US-led liberal international order, or are their internal divisions too significant to pose a coherent alternative? (20 Marks)\"\n"
        "• 2022 Q.8(b): \"Write short note on: Shanghai Cooperation Organization (SCO) and Regional Integration. (10 Marks)\""
    ),
    4: (
        "• 2025 Q.4: \"Abraham Accord has been described as re-aligning Middle East geopolitics. Critically assess the impact of this new diplomatic architecture on the viability of the two-state solution for the Israel-Palestine conflict. Does it create new leverage for peace or further marginalize the Palestinian issue? (20 Marks)\"\n"
        "• 2022 Q.6: \"Abraham Accords was designed to normalize relations between Israel and the Arab World; however, except for a few Arab States, the idea has not received popular support among the public and leaders. Critically examine the implications of this Accord on the Palestinian issue. (20 Marks)\"\n"
        "• 2013 Q.5: \"Critically examine the Palestine Problem from the Balfour Declaration (1917) to the post-Oslo developments. Why has a permanent peace settlement remained elusive? (20 Marks)\""
    ),
    5: (
        "• 2025 Q.5: \"How does cross-border terrorism affect regional stability? Critically assess the Pakistan-Taliban security crisis. What are the implications of the Taliban resurgence in Afghanistan for Pakistan’s external and internal security environment, particularly in Baluchistan and Khyber Pakhtunkhwa? (20 Marks)\"\n"
        "• 2022 Q.4: \"In the Realist prism, the 'State' is the legitimate actor in the international system, but the emergence of Non-State Actors (NSAs) has greatly impacted the global order; Taliban are governing Afghanistan without recognition and Hezbollah is a major stakeholder in Lebanon. Analyze the impact of NSAs on the global strategic environment with solid arguments and valid examples. (20 Marks)\"\n"
        "• 2018 Q.5: \"Discuss the emergence of ISIS and its impact on national, regional, and international peace and security. (20 Marks)\""
    ),
    6: (
        "• 2025 Q.6: \"Evaluate the interplay of diplomacy and propaganda in shaping ties between states. To what extent do these tools serve as mechanisms of conflict resolution versus instruments of power projection in the modern international system? (20 Marks)\"\n"
        "• 2016 Q.4: \"\\\"Diplomacy is the management of international relations by negotiation.\\\" Trace the evolution of diplomacy from traditional bilateral diplomacy to modern multilateral and public diplomacy. (20 Marks)\""
    ),
    7: (
        "• 2022 Q.1: \"Critically examine if the Ukraine War is a local affair or a competition between the two superpowers, and how it is to impact European unity. Give arguments and examples to justify your answer. (20 Marks)\""
    ),
    8: (
        "• 2022 Q.2: \"Discuss the salient features of the US Indo-Pacific Strategy enunciated in February 2022. Analyze its implications for the South Asian Security Environment with particular reference to Pakistan. (20 Marks)\"\n"
        "• 2016 Q.5: \"Critically evaluate the foreign policy of the United States toward South Asia post-9/11. How has Washington's strategic partnership with India influenced the regional balance of power? (20 Marks)\"\n"
        "• 2013 Q.8(d): \"Write short note on: Geopolitical rivalry in the Indian Ocean. (10 Marks)\""
    ),
    9: (
        "• 2022 Q.3: \"The trade war between the US and China has jeopardized the global economic order impacting all regions and states; Pakistan caught between these two global powers is likely to be affected the most. Analyze Pakistan's predicament and suggest policy options to safeguard its national security objectives. (20 Marks)\"\n"
        "• 2018 Q.6: \"There may be a possible trade war between the USA and China that would rattle certain American businesses and force consumers to pay more for products. Discuss the term \\\"trade war\\\" and elaborate on why Trump was targeting China. (20 Marks)\""
    ),
    10: (
        "• 2022 Q.4: \"In the Realist prism, the 'State' is the legitimate actor in the international system, but the emergence of Non-State Actors (NSAs) has greatly impacted the global order; Taliban are governing Afghanistan without recognition and Hezbollah is a major stakeholder in Lebanon. Analyze the impact of NSAs on the global strategic environment with solid arguments and valid examples. (20 Marks)\"\n"
        "• 2018 Q.5: \"Discuss the emergence of ISIS and its impact on national, regional, and international peace and security. (20 Marks)\""
    ),
    11: (
        "• 2022 Q.5: \"The Chinese Belt and Road Initiative (BRI) has global outreach in Asia, Africa, and Europe at the cost of US hegemony. Discuss its main objectives and analyze challenges and prospects of its success. (20 Marks)\""
    ),
    12: (
        "• 2022 Q.7: \"The Covid pandemic has introduced the concept of the 'New Normal' in the psycho-social attitudes of the public across the globe. Evaluate its structural impact on international relations, economic globalization, and global governance. (20 Marks)\"\n"
        "• 2018 Q.7: \"There are two schools of thought regarding the conjecture of globalization. First, view it as a blessing and a tool for prosperity; the other views it as a tool of frustration. Please elaborate. (20 Marks)\""
    ),
    13: (
        "• 2018 Q.1: \"Critically analyze Joseph Nye's concept of “Soft Power” in International Relations. (20 Marks)\"\n"
        "• 2013 Q.1: \"What is \\\"Power\\\" in International Relations? Differentiate between Hard Power, Soft Power, and Smart Power, and explain the tangible and intangible elements of National Power. (20 Marks)\""
    ),
    14: (
        "• 2018 Q.2: \"“Nationalism is neither inherently good nor bad, it is only a concept that can be used to achieve good or evil ends.” Elucidate. (20 Marks)\""
    ),
    15: (
        "• 2018 Q.3: \"Balance of power is a state of stability between competing forces; it is an equilibrium among states or alliances to prevent any other entity from becoming too strong. In light of this statement, elaborate how equilibrium was maintained by Great Powers during the Cold War era. (20 Marks)\"\n"
        "• 2013 Q.2: \"\\\"Balance of Power is a central concept in Realism.\\\" Discuss the various mechanisms employed by states to maintain or restore the Balance of Power (alliances, armaments, buffer states, compensation). (20 Marks)\"\n"
        "• 2013 Q.4: \"Discuss the Cold War alliance systems (NATO vs. Warsaw Pact, SEATO, CENTO). How did these military alliances shape the security architecture of Europe and Asia? (20 Marks)\""
    ),
    16: (
        "• 2018 Q.4: \"The direct cause of World War I was the assassination of “Archduke Franz Ferdinand.” However, historians feel that a number of factors contributed to the rivalry between Great Powers that led to war on such a wide scale. Comment please. (20 Marks)\"\n"
        "• 2013 Q.3: \"Analyze the causes and global repercussions of the First World War (1914–1918). How did the failure of the Treaty of Versailles contribute to the outbreak of World War II? (20 Marks)\""
    ),
    17: (
        "• 2018 Q.8: \"Highlight the provisions of the Geneva Convention on Prisoners of War. (20 Marks)\"\n"
        "• 2013 Q.6: \"What are the primary sources of International Law? Discuss the limitations of International Law as an instrument for regulating state behavior in an anarchic international system. (20 Marks)\""
    ),
    18: (
        "• 2016 Q.1: \"Critically examine the evolution of the Modern State System from the Peace of Westphalia (1648). How is traditional state sovereignty being redefined by globalization and transnational forces? (20 Marks)\""
    ),
    19: (
        "• 2016 Q.2: \"Compare and contrast Classical Realism (Hans Morgenthau) and Structural Realism / Neorealism (Kenneth Waltz). How do their core assumptions differ regarding the cause of conflict in international politics? (20 Marks)\"\n"
        "• 2013 Q.2: \"\\\"Balance of Power is a central concept in Realism.\\\" Discuss the various mechanisms employed by states to maintain or restore the Balance of Power (alliances, armaments, buffer states, compensation). (20 Marks)\""
    ),
    20: (
        "• 2016 Q.3: \"Nuclear deterrence prevented a direct hot war between superpowers during the Cold War. Discuss the theory of Mutually Assured Destruction (MAD) and its applicability to the strategic stability of South Asia between Pakistan and India. (20 Marks)\"\n"
        "• 2013 Q.7: \"Examine the dynamics of Arms Control and Disarmament in the nuclear era. Evaluate the effectiveness and shortcomings of the Nuclear Non-Proliferation Treaty (NPT) and Comprehensive Test Ban Treaty (CTBT). (20 Marks)\""
    ),
    21: (
        "• 2016 Q.6: \"The United Nations was established to save succeeding generations from the scourge of war. Critically assess the performance of the UN Security Council in conflict resolution and peacekeeping in the post-Cold War era. (20 Marks)\""
    ),
    22: (
        "• 2016 Q.7: \"Discuss the concept of \\\"Imperialism\\\" and its economic theories, particularly Hobson and Lenin's critique of imperial capitalism. To what extent does neo-imperialism persist in the contemporary global economic order? (20 Marks)\""
    ),
    23: (
        "• 2016 Q.8: \"What is the Non-Aligned Movement (NAM)? Evaluate its historical contribution during the Cold War bipolarity and analyze its relevance in today's multipolar world. (20 Marks)\"\n"
        "• 2013 Q.8(a): \"Write short note on: Decolonization and the Rise of the Third World. (10 Marks)\""
    )
}

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
    
    col_count = len(timeline_data[0]) if timeline_data else 3
    tbl_tl = doc.add_table(rows=len(timeline_data) + 1, cols=col_count)
    tbl_tl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_tl.autofit = False
    
    if col_count == 4:
        headers = ["Dimension / Arena", "Strategic Mechanism / Driver", "Systemic Impact", "Empirical Case"]
        widths = [Inches(1.5), Inches(1.8), Inches(2.0), Inches(1.2)]
    else:
        headers = ["Dimension / Arena", "Strategic Mechanism / Driver", "Systemic Impact on Global Order"]
        widths = [Inches(1.8), Inches(2.2), Inches(2.5)]
    
    hdr_row = tbl_tl.rows[0]
    for j in range(col_count):
        c = hdr_row.cells[j]
        c.width = widths[j]
        set_cell_background(c, "0E2B50")
        set_cell_margins(c, top=80, bottom=80, left=100, right=100)
        p = c.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(headers[j] if j < len(headers) else f"Column {j+1}")
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
        v_table = "<br>".join(v.split("\n"))
        md_content += f"| **{k}** | {v_table} |\n"
        
    # Dedicated EXAM QUESTIONS LINKED Section in MD
    exam_questions = []
    for k, v in meta_data:
        if "EXAM QUESTIONS" in k.upper():
            exam_questions = v.split("\n")
            break
            
    if exam_questions:
        md_content += "\n---\n\n### 🎯 EXAM QUESTIONS LINKED\n\n"
        for q in exam_questions:
            md_content += f"> {q}\n>\n"
            
    md_content += "\n---\n\n### 📌 Exam Essential Takeaways\n\n"
    for item in callout_data:
        md_content += f"- {item}\n"
        
    md_content += "\n---\n\n### 🗺️ Conceptual Thematic Map & Strategic Roadmap\n\n```text\n"
    md_content += map_text.strip() + "\n```\n\n---\n\n"
    
    col_count = len(timeline_data[0]) if timeline_data else 3
    md_content += "### 📊 Analytical Dimensions & Systemic Shifts Matrix\n\n"
    if col_count == 4:
        md_content += "| Dimension / Arena | Strategic Mechanism / Driver | Systemic Impact | Empirical Case |\n|---|---|---|---|\n"
        for d, m, s, c in timeline_data:
            md_content += f"| **{d}** | {m} | {s} | {c} |\n"
    else:
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
    .callout-box {{
        background-color: #F0F4F8;
        border-left: 4px solid #0E2B50;
        padding: 8px 12px;
        margin: 8px 0;
        font-size: 8.5pt;
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
        v_html = "<br/>".join([f"<strong>{line[:line.find(':')+1]}</strong>{line[line.find(':')+1:]}" if ":" in line and any(yr in line for yr in ["2025", "2022", "2018", "2016", "2013"]) else line for line in v.split("\n")])
        html_content += f"<tr><td><strong>{k}</strong></td><td>{v_html}</td></tr>\n"
    html_content += """</tbody></table>\n"""

    if exam_questions:
        html_content += """<h2>🎯 Exam Questions Linked</h2>
<div class="callout-box">
"""
        for q in exam_questions:
            q_clean = q.strip()
            if ":" in q_clean:
                prefix, text = q_clean.split(":", 1)
                html_content += f"<p style='margin: 4px 0;'><strong>{prefix}:</strong>{text}</p>\n"
            else:
                html_content += f"<p style='margin: 4px 0;'>{q_clean}</p>\n"
        html_content += "</div>\n"

    html_content += """<h2>Exam Essential Takeaways</h2>
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
<thead><tr>
"""
    if col_count == 4:
        html_content += "<th>Dimension / Arena</th><th>Strategic Mechanism / Driver</th><th>Systemic Impact</th><th>Empirical Case</th>"
    else:
        html_content += "<th>Dimension / Arena</th><th>Strategic Mechanism / Driver</th><th>Systemic Impact on Global Order</th>"
    html_content += """</tr></thead>
<tbody>
"""
    for row_data in timeline_data:
        html_content += "<tr>"
        for idx, val in enumerate(row_data):
            if idx == 0:
                html_content += f"<td><strong>{val}</strong></td>"
            else:
                html_content += f"<td>{val}</td>"
        html_content += "</tr>\n"
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


def extract_data_from_file(py_path):
    """Safely extracts qX data from a python batch file by parsing AST or executing in namespace."""
    with open(py_path, 'r', encoding='utf-8') as f:
        code = f.read()
    
    # Strip any top-level execution calls at the end
    clean_lines = []
    for line in code.splitlines():
        if re.match(r'^(generate_ir_master_doc|generate_pdf_and_md|for filename, meta)', line):
            break
        clean_lines.append(line)
    clean_code = "\n".join(clean_lines)
    
    ns = {}
    exec(clean_code, ns)
    return ns

def update_and_build_all():
    script_map = [
        (r"H:\My Drive\PMS KPK 2022 notes\apps_toolkit\build_ir_master_notes.py", [1]),
        (r"H:\My Drive\PMS KPK 2022 notes\apps_toolkit\build_ir_batch1_notes.py", [2, 3, 4, 5, 6]),
        (r"H:\My Drive\PMS KPK 2022 notes\apps_toolkit\build_ir_batch2_notes.py", [7, 8, 9, 10, 11, 12]),
        (r"H:\My Drive\PMS KPK 2022 notes\apps_toolkit\build_ir_batch3_notes.py", [13, 14, 15, 16, 17]),
        (r"H:\My Drive\PMS KPK 2022 notes\apps_toolkit\build_ir_batch4_notes.py", [18, 19, 20, 21, 22, 23]),
    ]
    
    titles_map = {
        1: ("Q1_Rise_of_China_Balance_of_Power_and_Multipolarity_Master_Notes", "QUESTION 1: RISE OF CHINA, BALANCE OF POWER & MULTIPOLARITY", "KPPSC PMS IR Master Notes — Q1: The Rise of China, Balance of Power & Multipolarity"),
        2: ("Q2_Indus_Basin_Treaty_and_Hydrological_Security_Master_Notes", "QUESTION 2: INDUS BASIN TREATY & HYDROLOGICAL SECURITY", "KPPSC PMS IR Master Notes — Q2: The Indus Basin Treaty & Hydrological Security"),
        3: ("Q3_Expansion_of_BRICS_and_SCO_Structural_Realism_Master_Notes", "QUESTION 3: EXPANSION OF BRICS & SCO UNDER STRUCTURAL REALISM", "KPPSC PMS IR Master Notes — Q3: Expansion of BRICS & SCO through Structural Realism"),
        4: ("Q4_Abraham_Accords_and_Palestine_Two_State_Solution_Master_Notes", "QUESTION 4: ABRAHAM ACCORDS & THE PALESTINE TWO-STATE SOLUTION", "KPPSC PMS IR Master Notes — Q4: The Abraham Accords & The Two-State Solution for Palestine"),
        5: ("Q5_Cross_Border_Terrorism_and_Pakistan_Taliban_Security_Crisis_Master_Notes", "QUESTION 5: CROSS-BORDER TERRORISM & THE PAKISTAN-TALIBAN CRISIS", "KPPSC PMS IR Master Notes — Q5: Cross-Border Terrorism & The Pakistan-Taliban Security Crisis"),
        6: ("Q6_Diplomacy_and_Propaganda_Conflict_Resolution_vs_Power_Projection_Master_Notes", "QUESTION 6: DIPLOMACY & PROPAGANDA IN MODERN STATECRAFT", "KPPSC PMS IR Master Notes — Q6: Diplomacy and Propaganda in Modern Statecraft"),
        7: ("Q7_The_Ukraine_War_Superpower_Rivalry_and_European_Unity_Master_Notes", "QUESTION 7: THE UKRAINE WAR, SUPERPOWER RIVALRY & EUROPEAN UNITY", "KPPSC PMS IR Master Notes — Q7: The Ukraine War, Superpower Rivalry & European Unity"),
        8: ("Q8_US_Indo_Pacific_Strategy_and_South_Asian_Security_Master_Notes", "QUESTION 8: US INDO-PACIFIC STRATEGY & SOUTH ASIAN SECURITY", "KPPSC PMS IR Master Notes — Q8: US Indo-Pacific Strategy & Implications for South Asia"),
        9: ("Q9_US_China_Geoeconomic_Rivalry_Trade_Wars_and_Pakistan_Master_Notes", "QUESTION 9: US-CHINA GEOECONOMIC RIVALRY & PAKISTAN'S PREDICAMENT", "KPPSC PMS IR Master Notes — Q9: US-China Geoeconomic Rivalry & Pakistan's Strategic Predicament"),
        10: ("Q10_Non_State_Actors_and_The_Crisis_of_Realist_State_Primacy_Master_Notes", "QUESTION 10: NON-STATE ACTORS & THE CRISIS OF REALIST STATE PRIMACY", "KPPSC PMS IR Master Notes — Q10: Non-State Actors & The Crisis of Realist State Primacy"),
        11: ("Q11_China_Belt_and_Road_Initiative_and_US_Hegemony_Master_Notes", "QUESTION 11: CHINA'S BELT AND ROAD INITIATIVE & US HEGEMONY", "KPPSC PMS IR Master Notes — Q11: China's Belt and Road Initiative & The Challenge to US Primacy"),
        12: ("Q12_Covid_Pandemic_New_Normal_and_Globalization_Master_Notes", "QUESTION 12: COVID-19 'NEW NORMAL' & THE CRISIS OF GLOBALIZATION", "KPPSC PMS IR Master Notes — Q12: The Covid-19 'New Normal', Global Governance & Globalization"),
        13: ("Q13_Joseph_Nye_Soft_Power_Hard_Power_and_Smart_Power_Master_Notes", "QUESTION 13: JOSEPH NYE'S SOFT POWER, HARD POWER & SMART POWER", "KPPSC PMS IR Master Notes — Q13: Joseph Nye's Soft Power, Hard Power & Smart Power"),
        14: ("Q14_Theories_of_Nationalism_Concept_and_State_Utility_Master_Notes", "QUESTION 14: THEORIES OF NATIONALISM & STATE UTILITY", "KPPSC PMS IR Master Notes — Q14: Theories of Nationalism, Concept & State Utility"),
        15: ("Q15_Balance_of_Power_Mechanisms_and_Cold_War_Equilibrium_Master_Notes", "QUESTION 15: BALANCE OF POWER MECHANISMS & COLD WAR EQUILIBRIUM", "KPPSC PMS IR Master Notes — Q15: Balance of Power Mechanisms & Cold War Equilibrium"),
        16: ("Q16_Systemic_Origins_of_World_War_I_Assassination_vs_Rivalry_Master_Notes", "QUESTION 16: SYSTEMIC ORIGINS OF WWI: SARAJEVO VS STRUCTURAL RIVALRY", "KPPSC PMS IR Master Notes — Q16: Systemic Origins of WWI: Assassination vs. Structural Rivalry"),
        17: ("Q17_Geneva_Convention_on_Prisoners_of_War_Master_Notes", "QUESTION 17: GENEVA CONVENTION ON PRISONERS OF WAR (POWs)", "KPPSC PMS IR Master Notes — Q17: Geneva Convention on the Protection of Prisoners of War"),
        18: ("Q18_Modern_State_System_Westphalia_to_Globalization_Master_Notes", "QUESTION 18: EVOLUTION OF THE MODERN STATE SYSTEM: WESTPHALIA TO GLOBALIZATION", "KPPSC PMS IR Master Notes — Q18: Evolution of Modern State System from Westphalia to Globalization"),
        19: ("Q19_Classical_Realism_vs_Structural_Neorealism_Master_Notes", "QUESTION 19: CLASSICAL REALISM VS STRUCTURAL NEOREALISM", "KPPSC PMS IR Master Notes — Q19: Classical Realism vs. Structural Neorealism"),
        20: ("Q20_Nuclear_Deterrence_Theory_and_South_Asian_Stability_Master_Notes", "QUESTION 20: NUCLEAR DETERRENCE THEORY, MAD & SOUTH ASIAN STRATEGIC STABILITY", "KPPSC PMS IR Master Notes — Q20: Nuclear Deterrence Theory, MAD & Strategic Stability in South Asia"),
        21: ("Q21_UN_Security_Council_Conflict_Resolution_and_Veto_Master_Notes", "QUESTION 21: UN SECURITY COUNCIL IN CONFLICT RESOLUTION & VETO POLITICS", "KPPSC PMS IR Master Notes — Q21: The UN Security Council in Post-Cold War Conflict Resolution"),
        22: ("Q22_Theories_of_Imperialism_Hobson_Lenin_and_Neo_Imperialism_Master_Notes", "QUESTION 22: THEORIES OF IMPERIALISM: HOBSON, LENIN & NEO-IMPERIALISM", "KPPSC PMS IR Master Notes — Q22: Theories of Imperialism: Hobson, Lenin & Contemporary Neo-Imperialism"),
        23: ("Q23_The_Non_Aligned_Movement_Cold_War_to_Multipolarity_Master_Notes", "QUESTION 23: THE NON-ALIGNED MOVEMENT (NAM): BIPOLARITY TO MULTIPOLARITY", "KPPSC PMS IR Master Notes — Q23: The Non-Aligned Movement: Cold War Bipolarity to Contemporary Multipolarity")
    }

    generated_count = 0
    for script_path, q_nums in script_map:
        print(f"\n=======================================================")
        print(f"Loading data from: {os.path.basename(script_path)}")
        print(f"=======================================================")
        ns = extract_data_from_file(script_path)
        
        for q_num in q_nums:
            meta = ns.get(f"q{q_num}_meta")
            callout = ns.get(f"q{q_num}_callout")
            cmap = ns.get(f"q{q_num}_map")
            timeline = ns.get(f"q{q_num}_timeline")
            essay = ns.get(f"q{q_num}_essay")
            quotes = ns.get(f"q{q_num}_quotes")
            
            if not all([meta, callout, cmap, timeline, essay, quotes]):
                print(f"WARNING: Incomplete data for Q{q_num} in {script_path}")
                continue
                
            # Update meta with EXAM QUESTIONS LINKED:
            new_meta = []
            for k, v in meta:
                if k == "Past Paper Alignment" or "EXAM QUESTIONS" in k.upper():
                    new_meta.append(("EXAM QUESTIONS LINKED:", EXAM_QUESTIONS_LINKED_MAP[q_num]))
                else:
                    new_meta.append((k, v))
            
            # If EXAM QUESTIONS LINKED: wasn't in original meta, insert it at index 2
            if not any("EXAM QUESTIONS" in k.upper() for k, v in new_meta):
                new_meta.insert(2, ("EXAM QUESTIONS LINKED:", EXAM_QUESTIONS_LINKED_MAP[q_num]))
                
            fn_base, sub_t, full_t = titles_map[q_num]
            docx_file = os.path.join(IR_DIR, f"{fn_base}.docx")
            pdf_file = os.path.join(IR_DIR, f"{fn_base}.pdf")
            md_file = os.path.join(IR_DIR, f"{fn_base}.md")
            
            print(f"Generating Q{q_num}: {fn_base}...")
            generate_ir_master_doc(
                meta_data=new_meta,
                callout_data=callout,
                map_text=cmap,
                timeline_data=timeline,
                essay_sections=essay,
                quotes_data=quotes,
                out_path=docx_file,
                paper_sub_title=sub_t
            )
            
            generate_pdf_and_md(
                meta_data=new_meta,
                callout_data=callout,
                map_text=cmap,
                timeline_data=timeline,
                essay_sections=essay,
                quotes_data=quotes,
                pdf_path=pdf_file,
                md_path=md_file,
                title=full_t
            )
            generated_count += 1

    print(f"\nSuccessfully generated {generated_count} of 23 International Relations Master Topics!")

if __name__ == "__main__":
    update_and_build_all()
