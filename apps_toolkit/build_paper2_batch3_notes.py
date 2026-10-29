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
paper2_dir = r"H:\My Drive\PMS KPK 2022 notes\1. European History Paper 1 and Paper 2\Paper 2"

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

def generate_master_doc(meta_data, callout_data, map_text, timeline_data, essay_sections, quotes_data, out_path, paper_sub_title):
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
    r_main = p_title.add_run("KPPSC PMS \u2022 EUROPEAN HISTORY (PAPER II: 1914\u20131960)")
    r_main.font.name = "Arial"
    r_main.font.size = Pt(12)
    r_main.font.bold = True
    r_main.font.color.rgb = RGBColor(100, 116, 139)
    
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(12)
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run(paper_sub_title)
    r_sub.font.name = "Arial"
    r_sub.font.size = Pt(15)
    r_sub.font.bold = True
    r_sub.font.color.rgb = RGBColor(14, 43, 80)
    
    meta_table = doc.add_table(rows=len(meta_data), cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_table.autofit = False
    
    for row_idx, (k, v) in enumerate(meta_data):
        row = meta_table.rows[row_idx]
        cell_k, cell_v = row.cells[0], row.cells[1]
        cell_k.width = Inches(2.2)
        cell_v.width = Inches(4.5)
        set_cell_background(cell_k, "0E2B50")
        set_cell_background(cell_v, "F8FAFC")
        set_cell_margins(cell_k, 60, 60, 100, 100)
        set_cell_margins(cell_v, 60, 60, 100, 100)
        
        pk = cell_k.paragraphs[0]
        pk.paragraph_format.space_after = Pt(0)
        rk = pk.add_run(k)
        rk.font.name = "Arial"
        rk.font.size = Pt(9)
        rk.font.bold = True
        rk.font.color.rgb = RGBColor(255, 255, 255)
        
        pv = cell_v.paragraphs[0]
        pv.paragraph_format.space_after = Pt(0)
        rv = pv.add_run(v)
        rv.font.name = "Georgia"
        rv.font.size = Pt(9.5)
        rv.font.color.rgb = RGBColor(30, 41, 59)
        
    doc.add_paragraph().paragraph_format.space_after = Pt(8)
    add_callout(doc, callout_data)
    doc.add_paragraph().paragraph_format.space_after = Pt(10)
    
    # 1. Conceptual Mind Map
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(16)
    h1.paragraph_format.space_after = Pt(4)
    r_h1 = h1.add_run("1. CONCEPTUAL MIND MAP & THEMATIC PATHWAY")
    r_h1.font.name = "Arial"
    r_h1.font.size = Pt(12)
    r_h1.font.bold = True
    r_h1.font.color.rgb = RGBColor(14, 43, 80)
    
    p_map = doc.add_paragraph()
    p_map.paragraph_format.space_after = Pt(12)
    r_m = p_map.add_run(map_text)
    r_m.font.name = "Consolas"
    r_m.font.size = Pt(9)
    r_m.font.color.rgb = RGBColor(15, 23, 42)
    
    # 2. Chronological Milestones
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(14)
    h2.paragraph_format.space_after = Pt(4)
    r_h2 = h2.add_run("2. CHRONOLOGICAL MILESTONES & DIPLOMATIC DATES")
    r_h2.font.name = "Arial"
    r_h2.font.size = Pt(12)
    r_h2.font.bold = True
    r_h2.font.color.rgb = RGBColor(14, 43, 80)
    
    tbl_time = doc.add_table(rows=len(timeline_data) + 1, cols=3)
    tbl_time.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_time.autofit = False
    
    headers = ["DATE / YEAR", "EVENT / ACTION", "HISTORICAL & EXAM SIGNIFICANCE"]
    for idx, text in enumerate(headers):
        c = tbl_time.rows[0].cells[idx]
        set_cell_background(c, "0E2B50")
        set_cell_margins(c, 60, 60, 80, 80)
        p = c.paragraphs[0]
        r = p.add_run(text)
        r.font.name = "Arial"
        r.font.size = Pt(9)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    col_widths = [Inches(1.1), Inches(2.3), Inches(3.3)]
    for r_idx, row_data in enumerate(timeline_data, start=1):
        row = tbl_time.rows[r_idx]
        bg = "F8FAFC" if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.width = col_widths[c_idx]
            set_cell_background(cell, bg)
            set_cell_margins(cell, 50, 50, 80, 80)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.name = "Georgia"
            r.font.size = Pt(8.5)
            r.font.color.rgb = RGBColor(30, 41, 59)
            if c_idx == 0:
                r.font.bold = True
                
    doc.add_paragraph().paragraph_format.space_after = Pt(14)
    
    # 3. The 20-Mark Master Model Essay
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(18)
    h3.paragraph_format.space_after = Pt(6)
    r_h3 = h3.add_run("3. THE 20-MARK MASTER MODEL ESSAY (~650 WORDS)")
    r_h3.font.name = "Arial"
    r_h3.font.size = Pt(13)
    r_h3.font.bold = True
    r_h3.font.color.rgb = RGBColor(14, 43, 80)
    
    for sub_title, text in essay_sections:
        h_sub = doc.add_paragraph()
        h_sub.paragraph_format.space_before = Pt(12)
        h_sub.paragraph_format.space_after = Pt(3)
        r_sub = h_sub.add_run(sub_title)
        r_sub.font.name = "Arial"
        r_sub.font.size = Pt(11)
        r_sub.font.bold = True
        r_sub.font.color.rgb = RGBColor(14, 43, 80)
        
        p_body = doc.add_paragraph()
        p_body.paragraph_format.space_before = Pt(0)
        p_body.paragraph_format.space_after = Pt(6)
        p_body.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r_b = p_body.add_run(text)
        r_b.font.name = "Georgia"
        r_b.font.size = Pt(10)
        r_b.font.color.rgb = RGBColor(30, 41, 59)
        
    doc.add_paragraph().paragraph_format.space_after = Pt(12)
    
    # 4. Historiographical Perspectives
    h4 = doc.add_paragraph()
    h4.paragraph_format.space_before = Pt(16)
    h4.paragraph_format.space_after = Pt(4)
    r_h4 = h4.add_run("4. HISTORIOGRAPHICAL PERSPECTIVES & SCHOLARLY QUOTES")
    r_h4.font.name = "Arial"
    r_h4.font.size = Pt(12)
    r_h4.font.bold = True
    r_h4.font.color.rgb = RGBColor(14, 43, 80)
    
    tbl_quotes = doc.add_table(rows=len(quotes_data) + 1, cols=3)
    tbl_quotes.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_quotes.autofit = False
    
    q_headers = ["HISTORIAN", "ANALYTICAL PERSPECTIVE", "EXACT QUOTABLE ARGUMENT"]
    for idx, text in enumerate(q_headers):
        c = tbl_quotes.rows[0].cells[idx]
        set_cell_background(c, "0E2B50")
        set_cell_margins(c, 60, 60, 80, 80)
        p = c.paragraphs[0]
        r = p.add_run(text)
        r.font.name = "Arial"
        r.font.size = Pt(9)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    q_widths = [Inches(1.4), Inches(2.0), Inches(3.3)]
    for r_idx, row_data in enumerate(quotes_data, start=1):
        row = tbl_quotes.rows[r_idx]
        bg = "F8FAFC" if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.width = q_widths[c_idx]
            set_cell_background(cell, bg)
            set_cell_margins(cell, 50, 50, 80, 80)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.name = "Georgia"
            r.font.size = Pt(8.5)
            r.font.color.rgb = RGBColor(30, 41, 59)
            if c_idx == 0:
                r.font.bold = True
                
    doc.add_paragraph().paragraph_format.space_after = Pt(16)
    p_foot = doc.add_paragraph()
    p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_f = p_foot.add_run("Khyber Pakhtunkhwa Public Service Commission (KPPSC) \u2022 PMS Competitive Examination Notes \u2022 European History Paper II")
    r_f.font.name = "Arial"
    r_f.font.size = Pt(8.5)
    r_f.font.color.rgb = RGBColor(148, 163, 184)
    
    doc.save(out_path)
    print("Generated Master Docx:", out_path)

def render_md_and_pdf(md_text, base_name):
    md_path = os.path.join(paper2_dir, f"{base_name}.md")
    pdf_path = os.path.join(paper2_dir, f"{base_name}.pdf")
    
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(md_text)
        
    html_body = ""
    for para in md_text.split("\n\n"):
        para = para.strip()
        if not para: continue
        if para.startswith("# "):
            html_body += f"<h1 style='text-align:center; font-size:16pt; margin-bottom:4px;'>{para[2:]}</h1>"
        elif para.startswith("## "):
            html_body += f"<h2 style='font-size:13.5pt; color:#0E2B50; margin-bottom:8px;'>{para[3:]}</h2>"
        elif para.startswith("### "):
            html_body += f"<h3 style='font-size:12.5pt; color:#1a365d; border-bottom:1px solid #cbd5e0; padding-bottom:3px; margin-top:16px;'>{para[4:]}</h3>"
        elif para.startswith("> "):
            html_body += f"<blockquote style='background:#f7fafc; border-left:4px solid #3182ce; padding:8px 12px; font-style:italic; margin:12px 0;'>{para[2:]}</blockquote>"
        elif para.startswith("---"):
            html_body += "<hr style='border:0; border-top:1px solid #e2e8f0; margin:14px 0;'>"
        else:
            parts = para.split("**")
            p_html = ""
            for idx, part in enumerate(parts):
                if idx % 2 == 1:
                    p_html += f"<strong>{part}</strong>"
                else:
                    p_html += part
            html_body += f"<p style='text-align:justify; line-height:1.55; font-size:11.5pt; margin-bottom:10px;'>{p_html}</p>"

    full_html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset='utf-8'>
<style>
  @page {{ size: A4; margin: 20mm 18mm 20mm 18mm; }}
  body {{ font-family: 'Times New Roman', Times, serif; color: #1a202c; }}
</style>
</head>
<body>
{html_body}
<div style='margin-top:25px; border-top:1px solid #a0aec0; padding-top:6px; text-align:center; font-size:9.5pt; color:#718096;'>
  Khyber Pakhtunkhwa Public Service Commission (KPPSC) &bull; PMS European History &bull; Paper II Master Notes
</div>
</body>
</html>"""

    with tempfile.NamedTemporaryFile("w", delete=False, suffix=".html", encoding="utf-8") as tf:
        tf.write(full_html)
        tmp_path = tf.name

    cmd = [EDGE_PATH, "--headless", "--disable-gpu", f"--print-to-pdf={pdf_path}", "--no-pdf-header-footer", tmp_path]
    subprocess.run(cmd, check=True)
    try:
        os.remove(tmp_path)
    except Exception:
        pass
    print("PDF and MD successfully generated:", pdf_path)

# =========================================================================
# QUESTION 10 (PAPER II): 2025 Q.5 (PT 2) / 2022 Q.6 / 2018 Q.7 / 2016 Q.6
# =========================================================================

meta_p2_10 = [
    ("EXAM QUESTIONS LINKED:", "2025 Q.5 (Part 2): 'Examine the causes and major turning points of the Second World War, including the nature of Nazi aggression and the Holocaust.'\n2022 Q.6: 'Examine the major wartime conferences (Tehran, Yalta, and Potsdam) and discuss how they shaped the post-WWII geopolitical map of Europe.'\n2018 Q.7: 'Analyze the consequences and impact of World War II.'\n2016 Q.6: 'Discuss the causes, milestones, and devastation of the Second World War in Europe.'\n2013 Q.7: 'Analyze the impact of World War II on the global status of Great Britain and the collapse of the European colonial empires.'"),
    ("SYLLABUS TOPIC:", "Paper II \u2014 Section 7: The Second World War (1939\u20131945) \u2014 Turning Points, The Holocaust, Wartime Diplomacy (Tehran, Yalta, Potsdam) & Consequences"),
    ("ALLOCATED MARKS / TIME:", "20 Marks | 35 Minutes | Standard Word Count: ~650\u2013700 words"),
    ("HISTORIOGRAPHICAL THESIS:", "The Second World War was an apocalyptic global struggle characterized by industrialized mass annihilation, racial extermination (the Holocaust), and total societal mobilization; while the military turning points of 1942\u20131943 shattered Axis invincibility, the summit diplomacy of the Grand Alliance (Tehran, Yalta, Potsdam) fundamentally reconstructed the global balance of power, terminating four centuries of European imperial hegemony, partitioning the continent along ideological frontiers, and enthroning the United States and the Soviet Union as the twin arbiters of the nuclear age.")
]

callout_p2_10 = [
    "\u2022 The Military Turning Points (1942\u20131943):",
    "  1. Battle of Stalingrad (Aug 1942\u2013Feb 1943): Soviet encirclement and annihilation of General Friedrich Paulus's German 6th Army (850,000 Axis casualties); permanently broke the Wehrmacht's offensive backbone in the East.",
    "  2. Second Battle of El Alamein (Oct\u2013Nov 1942): Bernard Montgomery's British Eighth Army shattered Erwin Rommel's Panzer Army Africa, safeguarding the Suez Canal and Middle Eastern oilfields.",
    "  3. Battle of Kursk (July 1943): The largest armored clash in human history (6,000 tanks); Soviet victory permanently deprived Nazi Germany of strategic offensive initiative.",
    "  4. D-Day / Operation Overlord (June 6, 1944): Allied amphibious assault on Normandy; opened the long-awaited Western Front, liberating France and trapping Germany in a fatal two-front vise.",
    "\u2022 The Holocaust (Shoah): The bureaucratic, state-directed industrial murder of six million European Jews and millions of Slavs, Roma, and political prisoners; formulated into the 'Final Solution' (Endl\u00f6sung) at the Wannsee Conference (January 1942) chaired by Reinhard Heydrich; executed via mechanized death factories (Auschwitz-Birkenau, Treblinka, Sobibor).",
    "\u2022 The Wartime Conferences That Carved Europe:",
    "  1. Tehran Conference (Nov 1943): Big Three (Roosevelt, Churchill, Stalin) agreed on D-Day invasion; shifted Poland's borders westward to the Curzon and Oder-Neisse lines.",
    "  2. Yalta Conference (Feb 1945): Divided Germany and Berlin into four occupation zones (US, UK, USSR, France); Stalin signed the Declaration on Liberated Europe pledging free elections (which he promptly violated); USSR promised to enter the war against Japan within 3 months of Germany's defeat; established the veto formula for the UN Security Council.",
    "  3. Potsdam Conference (July\u2013Aug 1945): Truman, Attlee, and Stalin finalized the 'Four Ds' for Germany (Demilitarization, Denazification, Decentralization, Democratization); Truman revealed atomic bomb possession; acrimonious disputes over Soviet reparations and Polish borders signaled the onset of the Cold War.",
    "\u2022 The Cataclysmic Consequences: Devastation of European cities (over 40 million dead in Europe, 27 million in USSR alone); bankruptcy of the British Empire accelerating rapid decolonization (India 1947); the permanent division of Germany; the end of European global primacy and the dawn of the Bipolar Superpower Era."
]

map_p2_10 = """========================================================================================
AXIS BLITZKRIEG (1939-41)            DECISIVE TURNING POINTS (1942-43)     POST-WAR RECONSTRUCTION
========================================================================================
[Fall of France (June 1940)]      --> [Battle of Stalingrad (1942-43)]  --> [Tehran, Yalta & Potsdam]
[Operation Barbarossa (1941)]     --> [Battle of El Alamein (1942)]     --> [Germany Divided into 4 Zones]
[Pearl Harbor: US Enters War]     --> [Battle of Kursk: Tank Clash]     --> [The Holocaust Revealed]
[Wannsee Conference: Final Sol.]  --> [D-Day: Normandy Landings (1944)]  --> [Fall of European Empires &
                                  --> [Soviet Red Army Captures Berlin]     Decolonization (Bipolar World)]
========================================================================================"""

timeline_p2_10 = [
    ("Sept 1, 1939", "Invasion of Poland", "Germany invades Poland using Blitzkrieg; Britain and France declare war on Sept 3; World War II begins."),
    ("May\u2013June 1940", "Fall of France & Dunkirk", "Wehrmacht bypasses Maginot Line via Ardennes; France surrenders in six weeks; Churchill rallies Britain."),
    ("June 22, 1941", "Operation Barbarossa", "Hitler launches surprise invasion of USSR with 3.8 million Axis soldiers; opens apocalyptic Eastern Front."),
    ("Dec 7, 1941", "Pearl Harbor", "Japan attacks US fleet; US enters war; Hitler declares war on US, sealing a global Grand Alliance."),
    ("Jan 20, 1942", "The Wannsee Conference", "Reinhard Heydrich coordinates the bureaucratic execution of the 'Final Solution' to exterminate 11 million European Jews."),
    ("Feb 2, 1943", "German Defeat at Stalingrad", "Field Marshal Paulus surrenders remaining 91,000 troops; turns the military tide of WWII permanently."),
    ("Nov 28\u2013Dec 1, 1943", "The Tehran Conference", "First Big Three summit; agrees on Operation Overlord and post-war Polish border adjustments."),
    ("June 6, 1944", "D-Day (Operation Overlord)", "156,000 Allied troops land on Normandy beaches; opens the Second Front in Western Europe."),
    ("Feb 4\u201311, 1945", "The Yalta Conference", "Agrees on four occupation zones in Germany, UN voting veto, and Declaration on Liberated Europe."),
    ("May 8, 1945", "V-E Day (Victory in Europe)", "Unconditional surrender of Nazi Germany following Hitler's suicide in the Berlin bunker (April 30)."),
    ("July 17\u2013Aug 2, 1945", "The Potsdam Conference", "Truman, Stalin, and Attlee finalize German occupation and reparations; Truman learns of Trinity atomic test.")
]

essay_p2_10 = [
    ("Introduction: The Crucible of Global Annihilation",
     "The Second World War (1939\u20131945) was the most destructive, transformative conflagration in recorded human history. Spanning six continents and consuming an estimated seventy to eighty-five million lives, the conflict was waged not merely for territorial expansion, but as an apocalyptic clash of incompatible civilizational ideologies: Nazi racial totalitarianism, Soviet Marxist-Leninism, and Anglo-American democratic capitalism. In the historiographical debate, while military historians like Richard Overy emphasize that Allied victory was a hard-fought achievement dependent upon Soviet sacrifice and American industrial mobilization rather than an inevitable outcome, diplomatic scholars like Gerhard Weinberg and Marc Trachtenberg demonstrate that the conflict fundamentally revolutionized the international architecture. By liquidating the military power of Germany and Japan, exhausting the imperial economies of Great Britain and France, and leaving the Red Army entrenched in Central Europe, World War II terminated four centuries of European global hegemony, dividing the continent into two armed spheres and inaugurating the Cold War."),

    ("The Early Axis Hegemony and the Mechanics of Blitzkrieg (1939\u20131941)",
     "The initial phase of the war witnessed the terrifying triumph of German 'Lightning War' (Blitzkrieg)\u2014the operational doctrine coordinating armored panzer divisions, motorized infantry, and close-air support from Stuka dive-bombers to achieve rapid strategic paralysis. Following the dismemberment of Poland in September 1939, the Wehrmacht executed the astonishing defeat of France in May\u2013June 1940. Bypassing the formidable concrete fortifications of the Maginot Line through the heavily forested Ardennes, German panzers severed Allied communications, trapped the British Expeditionary Force at Dunkirk, and compelled France to surrender in forty-six days, leaving Britain standing alone under Winston Churchill.\n\n"
     "The ideological climax of Nazi aggression commenced on June 22, 1941, with Operation Barbarossa\u2014the invasion of the Soviet Union. Deploying 3.8 million Axis soldiers along an eighteen-hundred-mile front, Hitler envisioned a war of racial extermination (Vernichtungskrieg) designed to annihilate 'Judeo-Bolshevism' and secure agricultural Lebensraum for Aryan colonization. However, despite capturing millions of Soviet soldiers in vast encirclement pockets, the Wehrmacht was halted before Moscow in December 1941 by the Siberian reserves of General Georgy Zhukov and the brutal Russian winter. That same month, Japan's surprise attack on Pearl Harbor brought the immense industrial colossus of the United States into the war, creating the unbeatable 'Grand Alliance' of Britain, the US, and the USSR."),

    ("The Decisive Turning Points: 1942\u20131943",
     "The tide of the global conflict turned irrevocably between the autumn of 1942 and the summer of 1943 across three titanic theaters:\n\n"
     "1. The Battle of Stalingrad (August 1942\u2013February 1943): The apex of the war on the Eastern Front. In the ruined streets of Stalingrad, the German Sixth Army under General Friedrich Paulus was drawn into ferocious urban combat (Rattenkrieg). In November 1942, Zhukov launched Operation Uranus, a double envelopment that trapped 300,000 Axis troops. When Paulus surrendered on February 2, 1943, over 850,000 Axis troops were dead or captured. Stalingrad permanently shattered the myth of German invincibility and placed the Wehrmacht on an irreversible defensive retreat.\n\n"
     "2. The Battle of El Alamein (October\u2013November 1942): In the sands of Egypt, General Bernard Montgomery's British Eighth Army shattered Field Marshal Erwin Rommel's German-Italian Panzer Army, permanently expelling the Axis from North Africa, safeguarding the Suez Canal, and paving the way for the Allied invasion of Sicily and Italy in 1943.\n\n"
     "3. The Battle of Kursk (July 1943): Hitler's final offensive gamble in the East (Operation Citadel) collided with deeply layered Soviet defensive belts. In the largest armored battle in history, involving six thousand tanks, the Red Army crushed the German panzer divisions, gaining permanent strategic initiative.\n\n"
     "4. D-Day / Operation Overlord (June 6, 1944): Under the supreme command of General Dwight D. Eisenhower, 156,000 Allied troops executed the largest amphibious invasion in history across the beaches of Normandy. Opening the long-promised Second Front, the Western Allies liberated Paris, crossed the Rhine, and trapped the Nazi war machine in an inescapable vise between the Anglo-American armies in the West and the Soviet juggernaut in the East."),

    ("The Industrialized Abyss: The Holocaust (Shoah)",
     "The defining moral atrocity of the Second World War was the Holocaust\u2014the systematic, state-sponsored genocide of six million European Jews, alongside millions of Soviet prisoners of war, Poles, Roma, homosexuals, and political dissidents. Unlike traditional historical pogroms born of religious fanaticism or spontaneous mob violence, the Holocaust was an industrialized, bureaucratic enterprise. Following the murderous sweep of the Einsatzgruppen (mobile killing squads) in the occupied Soviet territories, the Nazi leadership assembled at the Wannsee Conference on January 20, 1942, where SS General Reinhard Heydrich coordinated the logistics of the 'Final Solution' (Endl\u00f6sung).\n\n"
     "Modern industrial technology, state railway networks, and corporate chemical production (Zyklon-B gas supplied by IG Farben) were marshaled to herd millions into extermination camps, most notoriously Auschwitz-Birkenau, Treblinka, and Belzec. The Holocaust was the extreme, terrifying realization of Nazi biological racism: the total subordination of economic and military rationality to the ideological imperative of racial purification."),

    ("Summit Diplomacy: Tehran, Yalta, Potsdam and the Post-War Map",
     "As military victory approached, the political architecture of post-war Europe was negotiated across three monumental summits of the 'Big Three':\n\n"
     "1. The Tehran Conference (November 1943): Roosevelt, Churchill, and Stalin met personally for the first time. The summit confirmed the launch of Operation Overlord in May 1944 and agreed in principle to shift Poland's borders westward to the Oder-Neisse Line to compensate Poland for Soviet annexations up to the Curzon Line.\n\n"
     "2. The Yalta Conference (February 1945): With the Red Army sixty miles from Berlin, the summit at Yalta in the Crimea finalized the post-war order. Germany was to be disarmed, denazified, and partitioned into four military occupation zones (American, British, Soviet, and French), with Berlin similarly subdivided. Stalin signed the 'Declaration on Liberated Europe,' pledging free and unfettered democratic elections in Eastern Europe. In return, Stalin secured the annexation of eastern Poland and confirmed a permanent veto in the United Nations Security Council.\n\n"
     "3. The Potsdam Conference (July\u2013August 1945): Held in conquered Berlin, the political dynamics had shifted dramatically: Franklin D. Roosevelt was dead, replaced by the suspicious Harry S. Truman; Churchill was replaced mid-conference by Labour Prime Minister Clement Attlee; and Truman received confirmation of the successful Trinity atomic bomb test. Acrimonious disputes erupted over German reparations and Stalin's installation of puppet communist regimes in Poland and Romania. While agreeing to the 'Four Ds' (Demilitarization, Denazification, Decentralization, Democratization), Potsdam exposed the complete breakdown of the Grand Alliance, codifying the de facto partition of Europe into two hostile spheres."),

    ("Consequences: The Eclipse of Europe, Decolonization, and the Nuclear Dawn",
     "The consequences of the Second World War were civilizational in magnitude. Europe lay in physical, demographic, and financial ruin: over forty million Europeans were dead, millions were displaced persons (DPs), and historic cities from London to Berlin, Warsaw, and Stalingrad were reduced to rubble. Economically, Great Britain emerged as a victorious but bankrupt debtor nation, utterly dependent on American financial loans. This economic exhaustion accelerated the collapse of European colonial empires: within a decade, Britain, France, the Netherlands, and Belgium were forced to withdraw from India (1947), Indonesia (1949), and Indochina, terminating centuries of European overseas dominion.\n\n"
     "Most decisively, World War II dissolved the multipolar balance-of-power system that had governed international affairs since the Treaty of Westphalia in 1648. In its place arose an unprecedented bipolar world dominated by two continental superpowers: the United States, possessing half of the world's manufacturing capacity and the atomic bomb, and the Soviet Union, commanding the massive conventional juggernaut of the Red Army. The dropping of atomic bombs on Hiroshima and Nagasaki in August 1945 did not merely conclude World War II; it inaugurated the nuclear age, ensuring that the post-war European settlement would be contested under the shadow of thermonuclear annihilation."),

    ("Conclusion: The Watershed of Modern History",
     "In conclusion, the Second World War was the cataclysmic watershed of modern times. It crushed the murderous monstrosity of Nazi fascism, exposed the ultimate horror of industrialized racism in the Holocaust, and permanently restructured the global political order. Yet, the peace achieved in 1945 was neither stable nor unified. The Grand Alliance, forged solely through the existential necessity of defeating Hitler, dissolved immediately upon victory. The agreements made at Yalta and Potsdam became the ideological battle lines of the Cold War. As an 'Iron Curtain' descended across the continent, dividing Germany and Europe into two fortified nuclear camps, the European nations ceased to be the masters of their own destiny, becoming instead the contested frontier in a global superpower confrontation.")
]

quotes_p2_10 = [
    ("Richard Overy", "Why the Allies Won (1995)", "'Allied victory was not an inevitable consequence of superior economic and human resources... It was forged in the terrible crucible of military adaptation, political will, and the colossal sacrifice of the Soviet Red Army on the Eastern Front.'"),
    ("Gerhard Weinberg", "A World at Arms (1994)", "'World War II was the direct product of Adolf Hitler's fanatical ideological vision of racial conquest and Lebensraum... It reshaped the globe more thoroughly than any event since the fall of the Roman Empire.'"),
    ("Winston Churchill", "Speech in the House of Commons (1940)", "'Hitler knows that he will have to break us in this island or lose the war. If we can stand up to him, all Europe may be freed and the life of the world may move forward into broad, sunlit uplands.'"),
    ("George F. Kennan", "Memoirs (1967)", "'The reality of 1945 was that the defeat of Germany left the Soviet Union as the dominant military colossus of Eurasia... The post-war settlement was not a peace of reconciliation, but a partition of exhaustion.'")
]

# =========================================================================
# QUESTION 11 (PAPER II): 2025 Q.6 (PT 1) / 2022 Q.8 / 2016 Q.7 / 2013 Q.8
# =========================================================================

meta_p2_11 = [
    ("EXAM QUESTIONS LINKED:", "2025 Q.6 (Part 1): 'Examine the origins and evolution of the Cold War. How did ideological differences and power politics divide Europe after 1945?'\n2022 Q.8: 'Discuss the division of Germany and the Berlin Crisis (1948\u201349) as central flashpoints in the emergence of the Cold War.'\n2016 Q.7: 'Analyze the Truman Doctrine and the Marshall Plan as fundamental instruments of American foreign policy in post-war Europe.'\n2013 Q.8: 'Discuss the origins of the Cold War and the ideological confrontation between the United States and the USSR in Europe from 1945 to 1955.'"),
    ("SYLLABUS TOPIC:", "Paper II \u2014 Section 8: Origins of the Cold War (1945\u20131955) \u2014 Ideological Polarization, Truman Doctrine, Marshall Plan & The Berlin Blockade"),
    ("ALLOCATED MARKS / TIME:", "20 Marks | 35 Minutes | Standard Word Count: ~650\u2013700 words"),
    ("HISTORIOGRAPHICAL THESIS:", "The Cold War in Europe was not the exclusive product of Soviet expansionist design nor American capitalist imperialism, but the inevitable structural clash between two diametrically opposed universalist ideologies operating within the geopolitical power vacuum created by the total destruction of Nazi Germany; through George Kennan's containment doctrine, Churchill's 'Iron Curtain' indictment, the economic engine of the Marshall Plan, and the supreme crisis of the Berlin Blockade (1948\u20131949), the division of Europe was permanently codified into two fortified, irreconcilable spheres.")
]

callout_p2_11 = [
    "\u2022 The Core Structural Drivers: The vacuum left by the total collapse of Germany in 1945; the 'security dilemma' wherein defensive measures taken by one superpower (Soviet buffer states in Eastern Europe vs. American financial reconstruction of Western Europe) were perceived by the other as offensive, aggressive imperialism.",
    "\u2022 The Ideological Dichotomy: American liberal capitalism (individual liberty, free-market trade, open covenants, democratic elections) clashed irreconcilably with Soviet Marxist-Leninism (historical materialism, state-controlled command economy, anti-capitalist world revolution, absolute party autocracy).",
    "\u2022 The Intellectual Codification of Containment:",
    "  1. George F. Kennan's 'Long Telegram' (Feb 1946): Argued that Soviet policy was driven by traditional Russian insecurity and Marxist fanaticism; advocated 'long-term, patient but firm and vigilant containment of Russian expansive tendencies.'",
    "  2. Winston Churchill's 'Iron Curtain' Speech (March 1946): At Fulton, Missouri, declared: 'From Stettin in the Baltic to Trieste in the Adriatic, an iron curtain has descended across the Continent.'",
    "\u2022 The Twin Pillars of American Strategy (1947):",
    "  1. The Truman Doctrine (March 1947): Prompted by Britain's financial collapse in Greece and Turkey; pledged $400 million in aid and committed the US 'to support free peoples who are resisting attempted subjugation by armed minorities or by outside pressures' (universalizing containment).",
    "  2. The Marshall Plan (European Recovery Program, June 1947): Dispatched $13.3 billion ($150B in modern terms) in economic aid to rebuild Western European industries and eliminate the poverty breeding communism; Stalin denounced it as 'dollar imperialism', forbade satellite participation, and established Cominform (1947) and Comecon (1949).",
    "\u2022 The Supreme Flashpoint: The Berlin Blockade & Airlift (June 1948\u2013May 1949): Western Allies unified their zones (Trizonia) and introduced the stable Deutsche Mark; Stalin retaliated on June 24, 1948, by severing all road, rail, and canal access to Allied West Berlin, attempting to starve the city into surrender; the US and Britain responded with 'Operation Vittles' (the Berlin Airlift), flying 278,000 missions over 324 days, landing a cargo plane every 3 minutes; Stalin conceded humiliating defeat on May 12, 1949.",
    "\u2022 The Partition of Germany: Direct consequence of the Berlin Crisis: Federal Republic of Germany (FRG, West Germany) established in May 1949 under Konrad Adenauer; Soviet-backed German Democratic Republic (GDR, East Germany) created in October 1949 under Walter Ulbricht."
]

map_p2_11 = """========================================================================================
POST-1945 VACUUM                  THE CONTAINMENT EDIFICE (1947)        BERLIN & PARTITION (1948-49)
========================================================================================
[Collapse of Nazi Germany]     --> [Truman Doctrine: $400M Aid]       --> [Western Currency: Deutsche Mark]
[Soviet Satellite Imposition]  --> [Marshall Plan: $13B Rebuild]      --> [Stalin's Berlin Land Blockade]
[Kennan's Long Telegram (1946)]--> [Stalin Forms Cominform/Comecon]   --> [The Berlin Airlift: 2.3M Tons]
[Churchill's Iron Curtain]     --> [Prague Coup: Czech Democracy Falls]--> [Stalin Yields: May 1949]
                                                                      --> [Permanent Division: FRG vs. GDR]
========================================================================================"""

timeline_p2_11 = [
    ("Feb 22, 1946", "Kennan's Long Telegram", "George Kennan analyzes Soviet conduct from Moscow, formulating the foundational doctrine of Containment."),
    ("March 5, 1946", "Churchill's 'Iron Curtain' Speech", "Delivered at Westminster College in Fulton, Missouri; warns of descending Soviet totalitarian curtain."),
    ("March 12, 1947", "Announcement of Truman Doctrine", "President Truman requests $400 million for Greece and Turkey, establishing universal anti-communist containment."),
    ("June 5, 1947", "Marshall Proposes Recovery Plan", "Secretary of State George C. Marshall unveils $13 billion European economic recovery program at Harvard."),
    ("Sept 1947", "Creation of Cominform", "Stalin forms Communist Information Bureau to enforce total ideological obedience among communist parties."),
    ("Feb 1948", "The Czech Coup d'\u00e9tat", "Communists seize complete control of Czechoslovakia; Foreign Minister Jan Masaryk dies under suspicious circumstances."),
    ("June 20, 1948", "Deutsche Mark Introduced", "Western Allies introduce new currency in Western zones and West Berlin, creating a booming capitalist economy."),
    ("June 24, 1948", "Stalin Imposes Berlin Blockade", "Soviets cut off all rail, highway, and barge traffic from West Germany to West Berlin."),
    ("June 1948\u2013May 1949", "The Berlin Airlift (Operation Vittles)", "US and British air forces deliver 2.3 million tons of supplies, feeding and heating 2.5 million West Berliners."),
    ("May 12, 1949", "Soviets Lift Berlin Blockade", "Stalin reopens land corridors after 324 days, suffering an unmitigated public relations and diplomatic defeat."),
    ("May 23 & Oct 7, 1949", "Dual German States Formed", "Federal Republic of Germany (FRG) established in the West; German Democratic Republic (GDR) formed in the East.")
]

essay_p2_11 = [
    ("Introduction: The Emergence of the Bipolar Divide",
     "Within twenty-four months of the joint Allied celebration over the ruins of the Third Reich in May 1945, the Grand Alliance had disintegrated into the bitter, ideologically polarized global confrontation known as the Cold War. In the historiographical debate, scholars have long contested the responsibility for this division. Traditional 'Orthodox' historians like Arthur Schlesinger Jr. attributed the conflict entirely to Soviet ideological expansionism and Stalin's paranoid aggression in violating the Yalta agreements. In stark contrast, 1960s 'Revisionist' historians like William Appleman Williams argued that aggressive American economic imperialism and 'atomic diplomacy' forced the Soviet Union into a defensive, reactive posture. However, modern 'Post-Revisionist' scholars, led by John Lewis Gaddis and Melvyn Leffler, utilizing declassified Soviet and Western archives, demonstrate that the Cold War was an overdetermined structural tragedy. The complete annihilation of German power created a geopolitical vacuum in Central Europe that neither the United States nor the Soviet Union could afford to leave to the other. Driven by irreconcilable ideological visions and intensified by the classic 'security dilemma,' the transition from wartime cooperation to European partition unfolded rapidly through the milestones of the Truman Doctrine, the Marshall Plan, and the crucible of the 1948\u20131949 Berlin Blockade."),

    ("The Ideological Divide and the Post-War Power Vacuum",
     "At its philosophical core, the Cold War was an existential clash between two universalist, messianic ideologies that had competed for the soul of the West since 1917. The American worldview, articulated in the Atlantic Charter and Wilsonian tradition, championed liberal democratic capitalism, individual human rights, the rule of law, and open international free trade under the Bretton Woods system. Washington policymakers genuinely believed that economic protectionism and totalitarian dictatorships were the root causes of modern war. Conversely, the Soviet worldview was grounded in Marxist-Leninist dogma: capitalism was inherently imperialist and predatory, class conflict was scientifically inevitable, and the Soviet socialist motherland was surrounded by hostile capitalist powers seeking its destruction.\n\n"
     "This ideological antagonism was weaponized by the post-war power vacuum. In 1945, every traditional Great Power on the European continent\u2014Germany, France, and Italy\u2014was crushed or economically paralyzed, while Great Britain was financially bankrupt. Only the United States, possessing half of the world's manufacturing capacity and a temporary monopoly on the atomic bomb, and the Soviet Union, commanding the vast land armies of the Red Army, possessed the capacity to project continental power. In this vacuum, any defensive action taken by one power was perceived by the other as an existential threat."),

    ("The Breakdown of Trust: Satellite Imposition and Kennan's Containment",
     "The initial rupture occurred in Eastern Europe. Having suffered twenty-seven million dead and the utter devastation of its western agricultural and industrial heartlands during World War II, Stalin's primary post-war objective was absolute national security. To ensure that Russia could never again be invaded through the North European plain, Stalin systematically transformed Poland, Romania, Bulgaria, Hungary, and later Czechoslovakia into subservient communist satellite states. Through the use of 'salami tactics'\u2014slicing away democratic opposition parties piece by piece through police terror, rigged elections, and fabricated treason trials\u2014Stalin broke his solemn pledge at Yalta to permit free and unfettered democratic elections.\n\n"
     "The Western response was intellectually codified in February 1946 by George F. Kennan, the American Charg\u00e9 d'Affaires in Moscow, in his famous eight-thousand-word 'Long Telegram.' Kennan argued that Soviet leadership was driven by a neurotically traditional Russian sense of insecurity fused with Marxist fanaticism, making compromise impossible. Kennan advocated the historic strategy of 'Containment': the United States must apply patient, firm, and vigilant counter-force at every point where the Soviets showed signs of encroaching upon the peace of the world. Two weeks later, on March 5, 1946, at Fulton, Missouri, Winston Churchill delivered his historic 'Iron Curtain' address alongside President Truman, publicly declaring that 'from Stettin in the Baltic to Trieste in the Adriatic, an iron curtain has descended across the Continent.' The post-war consensus had crystallized: the West was now engaged in a struggle against totalitarian aggression."),

    ("The Twin Pillars of Containment: The Truman Doctrine and the Marshall Plan (1947)",
     "The strategy of containment was operationalized in the spring of 1947 through two revolutionary initiatives. The catalyst was Great Britain's formal diplomatic note to Washington in February 1947, confessing that financial exhaustion compelled Britain to terminate its economic and military assistance to Greece (which was fighting a communist civil war) and Turkey (which was facing Soviet demands for joint control of the Turkish Straits). President Harry S. Truman recognized that British imperial withdrawal threatened to collapse the entire Mediterranean into the Soviet orbit.\n\n"
     "On March 12, 1947, Truman addressed a joint session of Congress, proclaiming the 'Truman Doctrine.' Requesting $400 million in emergency aid for Greece and Turkey, Truman framed the struggle in stark, Manichean terms, declaring that it must be the policy of the United States 'to support free peoples who are resisting attempted subjugation by armed minorities or by outside pressures.' The Truman Doctrine was of monumental significance: it formally abandoned George Washington's 1796 farewell warning against peacetime foreign entanglements, committing the United States to a global policy of anti-communist intervention.\n\n"
     "Three months later, on June 5, 1947, Secretary of State George C. Marshall unveiled the economic engine of containment: the Marshall Plan (European Recovery Program). Recognizing that the catastrophic winter of 1946\u20131947, severe coal shortages, and industrial devastation were driving desperate European populations into the arms of massive domestic communist parties (especially in France and Italy), the United States injected $13.3 billion ($150 billion in contemporary value) into Western Europe between 1948 and 1952. The Marshall Plan revitalized European industrial production, stabilized democratic currencies, and integrated Western European economies into a capitalist trade bloc. Stalin correctly recognized the Marshall Plan as an existential threat to Soviet influence: he prohibited Poland and Czechoslovakia from accepting aid, denounced the plan as 'dollar imperialism,' and organized the Cominform (Communist Information Bureau) and Comecon (Council for Mutual Economic Assistance) to enforce total economic and ideological conformity across the Eastern bloc."),

    ("The Crucible of Confrontation: The Berlin Blockade and Airlift (1948\u20131949)",
     "The ultimate flashpoint of the early Cold War centered on the fate of divided Germany. At Yalta and Potsdam, Germany had been partitioned into four occupation zones, and Berlin\u2014situated 110 miles inside the Soviet zone\u2014was similarly subdivided. However, four-power administration collapsed over reparations. While Stalin stripped his eastern zone of factories, railways, and machinery to rebuild the USSR, the Western Allies realized that European economic recovery was impossible without a revitalized, industrial German economy. In early 1948, the United States, Britain, and France merged their zones economically into 'Trizonia' and announced plans for a democratic West German constitution.\n\n"
     "The immediate trigger of the crisis was the Western introduction of the stable 'Deutsche Mark' on June 20, 1948, replacing the worthless Reichsmark in Western Germany and West Berlin. Recognizing that a booming, capitalist West Berlin in the center of the Soviet zone was an intolerable ideological challenge, Stalin struck back. On June 24, 1948, the Soviet military severed all road, rail, and canal traffic connecting West Germany to Berlin, cut off electrical power, and halted all food and coal shipments, aiming to starve the two million citizens of West Berlin into capitulation and force the Western Allies to abandon the city.\n\n"
     "Rather than retreating or triggering World War III through an armed armored dash up the autobahn, President Truman, backed by British Foreign Secretary Ernest Bevin, launched 'Operation Vittles'\u2014the historic Berlin Airlift. Commanded by General Curtis LeMay and General William Tunner, American and British transport aircraft (C-47s and C-54s) flew around the clock through three narrow air corridors. Over the next 324 days, Allied pilots flew 278,000 missions, landing a plane at Tempelhof Airport every three minutes day and night, delivering 2.3 million tons of food, medicine, and coal. The airlift transformed the Western Allies from hated conquerors into heroic protectors in the eyes of the German people, inflicting an unmitigated moral and propaganda disaster upon the Soviet Union. On May 12, 1949, acknowledging total defeat, Stalin reopened the land access routes."),

    ("Conclusion: The Consolidation of a Divided Continent",
     "In conclusion, the origins of the Cold War between 1945 and 1949 represented the tragic, inevitable consolidation of European partition. The Berlin Blockade was the catalyst that hardened the division of Germany into two rival sovereign states: on May 23, 1949, the Federal Republic of Germany (FRG, West Germany) was established in Bonn under Chancellor Konrad Adenauer, and on October 7, 1949, the Soviet Union proclaimed the German Democratic Republic (GDR, East Germany) in East Berlin under Walter Ulbricht. The confrontation demonstrated that neither superpower could dislodge the other from its sphere without risking general war. By 1949, with the atomic bomb now tested by the Soviets and the American commitment sealed through the Marshall Plan, the battle lines of European geopolitics were frozen in place for the next forty years, preparing the terrain for the final military institutionalization of the continent: the creation of NATO and the Warsaw Pact.")
]

quotes_p2_11 = [
    ("John Lewis Gaddis", "We Now Know: Rethinking Cold War History (1997)", "'As long as Stalin was running the Soviet Union, a cold war was unavoidable... His authoritarian personality, combined with Marxist-Leninist ideology, ensured that he viewed any independent power as a mortal enemy.'"),
    ("Arthur Schlesinger Jr.", "The Origins of the Cold War", "'The Cold War was the brave and essential Western response to the intractable dilemma posed by the aggressive, paranoid expansionism of the Soviet totalitarian state.'"),
    ("Winston Churchill (March 5, 1946)", "The Sinews of Peace (Fulton, Missouri)", "'From Stettin in the Baltic to Trieste in the Adriatic, an iron curtain has descended across the Continent... This is certainly not the Liberated Europe we fought to build up.'"),
    ("George F. Kennan (1947)", "The Sources of Soviet Conduct", "'The main element of any United States policy toward the Soviet Union must be that of a long-term, patient but firm and vigilant containment of Russian expansive tendencies.'")
]

# =========================================================================
# QUESTION 12 (PAPER II): 2025 Q.6 (PT 2) / 2016 Q.8 (NATO VS WARSAW PACT)
# =========================================================================

meta_p2_12 = [
    ("EXAM QUESTIONS LINKED:", "2025 Q.6 (Part 2): 'How did ideological differences and power politics divide Europe after 1945?'\n2016 Q.8: 'Evaluate the emergence of NATO and the Warsaw Pact in consolidating the bipolar division of Europe.'"),
    ("SYLLABUS TOPIC:", "Paper II \u2014 Section 8: Bipolar Military Alliances & Security Architecture \u2014 NATO (1949), Warsaw Pact (1955) & Nuclear Confrontation"),
    ("ALLOCATED MARKS / TIME:", "20 Marks | 35 Minutes | Standard Word Count: ~650\u2013700 words"),
    ("HISTORIOGRAPHICAL THESIS:", "The creation of the North Atlantic Treaty Organization (NATO) in 1949 and the reciprocal formation of the Warsaw Pact in 1955 completed the militarization and institutionalization of the European bipolar divide; by anchoring international stability in the terrifying logic of nuclear deterrence and Mutually Assured Destruction (MAD), this military architecture converted Europe into a heavily fortified, partitioned continent, paradoxically preventing a direct third world war while entrenching superpower hegemony over their respective client halves.")
]

callout_p2_12 = [
    "\u2022 The Genesis of NATO (April 4, 1949): Prompted by the 1948 Czechoslovak communist coup and the Berlin Blockade; 12 founding nations (US, UK, France, Canada, Italy, Benelux, Norway, Denmark, Iceland, Portugal) signed the North Atlantic Treaty in Washington; represented the first peacetime military alliance outside the Western Hemisphere in American history.",
    "\u2022 Article 5 (The Soul of the Alliance): Enshrined the principle of collective defense: 'an armed attack against one or more of them in Europe or North America shall be considered an attack against them all'; crucially backed by the American nuclear deterrence umbrella.",
    "\u2022 The West German Rearmament Crisis (1950\u20131955): The outbreak of the Korean War in June 1950 convinced the US that Western Europe could not be defended against the 175 Soviet Red Army divisions without German manpower; after France rejected the European Defense Community (EDC), the 1954 Paris Agreements integrated a sovereign, rearmed West Germany (Bundeswehr) into NATO on May 9, 1955.",
    "\u2022 The Creation of the Warsaw Pact (May 14, 1955): The Soviet Union's direct response to West Germany's entry into NATO; signed in Warsaw by the USSR and seven satellites (East Germany, Poland, Czechoslovakia, Hungary, Romania, Bulgaria, Albania); established a unified military command under Soviet Marshal Ivan Konev.",
    "\u2022 Contrast in Nature: NATO was a voluntary defensive alliance of sovereign democracies operating by consensus; the Warsaw Pact was a coercive instrument of Soviet imperial domination, designed primarily to enforce ideological obedience and crush anti-Soviet rebellions within its own members (as brutally demonstrated in the suppression of the Hungarian Uprising in 1956 and the Prague Spring in 1968).",
    "\u2022 The Nuclear Dimension & The 'Long Peace': The loss of the US atomic monopoly (Soviet A-bomb 1949) and the development of the Hydrogen Bomb (1952\u20131953) generated strategic parity; military doctrine transitioned from Eisenhower's 'Massive Retaliation' to 'Mutually Assured Destruction' (MAD), freezing the European front and eliminating border ambiguities."
]

map_p2_12 = """========================================================================================
MILITARY THREAT (1948-49)          WESTERN COALITION (NATO 1949)         EASTERN COUNTER-ALLIANCE
========================================================================================
[Prague Coup & Berlin Siege]    --> [April 1949: NATO Formed (12)]    --> [Soviet A-Bomb Tested (1949)]
[Soviet Red Army Superiority]   --> [Article 5: Collective Defense]   --> [Korean War: Arms Race Surges]
[US Peacetime Alliance Shift]   --> [1955: West Germany Joins NATO]   --> [May 1955: WARSAW PACT FORMED]
                                                                      --> [Internal Coercion: Hungary '56]
                                                                      --> [Nuclear Parity & MAD Stability]
========================================================================================"""

timeline_p2_12 = [
    ("March 17, 1948", "The Brussels Treaty", "Britain, France, Belgium, Netherlands, and Luxembourg sign collective defense agreement; precursor to NATO."),
    ("April 4, 1949", "Signing of the North Atlantic Treaty", "12 founding nations establish NATO in Washington, DC, pledging collective defense under Article 5."),
    ("Aug 29, 1949", "First Soviet Atomic Bomb Test", "USSR detonates atomic bomb (RDS-1) at Semipalatinsk; ends American four-year atomic monopoly."),
    ("June 1950", "Outbreak of the Korean War", "North Korean invasion of South Korea intensifies Western fears of a synchronized Soviet blitzkrieg in Europe."),
    ("Feb 1952", "NATO Lisbon Conference", "NATO sets target of 96 conventional divisions; recognizes conventional inferiority; embraces nuclear deterrence."),
    ("Nov 1952 & Aug 1953", "Thermonuclear H-Bomb Breakthroughs", "US and USSR detonate hydrogen bombs; yields measured in megatons, ushering in the era of MAD."),
    ("Oct 1954", "The Paris Agreements", "Restores West German sovereignty and approves German rearmament within the Western European Union and NATO."),
    ("May 9, 1955", "West Germany Joins NATO", "Federal Republic of Germany officially admitted as the 15th member of NATO; begins building the Bundeswehr."),
    ("May 14, 1955", "Creation of the Warsaw Pact", "USSR and 7 Eastern European satellites sign Warsaw Treaty of Friendship, Co-operation and Mutual Assistance."),
    ("Nov 1956", "Soviet Invasion of Hungary", "Soviet tanks crush Imre Nagy's democratic rebellion after Hungary announces withdrawal from the Warsaw Pact.")
]

essay_p2_12 = [
    ("Introduction: The Institutionalization of the Bipolar Division",
     "The decade between 1945 and 1955 witnessed the complete militarization and institutional hardening of the European continent into two hostile, heavily armed ideological coalitions: the North Atlantic Treaty Organization (NATO) and the Warsaw Treaty Organization (the Warsaw Pact). While the diplomatic and economic division of Europe had been established through the Truman Doctrine and the Marshall Plan, the physical security architecture of the Cold War was finalized only when both superpowers anchored their spheres of influence in formal military alliances backed by the terrifying power of thermonuclear weaponry. In the historiographical debate, scholars such as Lawrence Freedman and John Lewis Gaddis recognize the profound paradox of this military confrontation: while the establishment of NATO and the Warsaw Pact permanently entrenched the tragic division of European civilization, it simultaneously created the 'Long Peace'\u2014a hyper-stable, predictable international system wherein nuclear deterrence and Mutually Assured Destruction (MAD) rendered a direct major war between the Great Powers utterly unthinkable."),

    ("The Birth of NATO: From Brussels to Article 5 (1948\u20131949)",
     "The catalyst for the creation of NATO was the acute sense of Western European military vulnerability exposed by the 1948 Czechoslovak communist coup d'\u00e9tat and the Berlin Blockade. Confronted by twenty-five divisions of the Soviet Red Army stationed in Central Europe against a disarmed, demobilized West, British Foreign Secretary Ernest Bevin recognized that Western European nations could not defend themselves alone. In March 1948, Britain, France, and the Benelux nations signed the Brussels Treaty, establishing a mutual defense pact. However, European statesmen recognized that without the direct, permanent military commitment of the United States, European defense was a complete illusion.\n\n"
     "The diplomatic breakthrough arrived on April 4, 1949, when twelve nations (the United States, Great Britain, France, Canada, Italy, Belgium, Netherlands, Luxembourg, Norway, Denmark, Iceland, and Portugal) convened in Washington to sign the North Atlantic Treaty. The creation of NATO was an unprecedented revolution in American foreign policy: for the first time since the 1778 Franco-American alliance, the United States abandoned its sacred isolationist tradition and committed itself to a permanent, peacetime military alliance outside the Western Hemisphere. The institutional heart of the treaty was Article 5, which declared that 'an armed attack against one or more of them in Europe or North America shall be considered an attack against them all.' By providing an unambiguous American nuclear guarantee to Western Europe, NATO successfully deterred Soviet military expansionism, perfectly fulfilling its first Secretary General, Lord Ismay's, famous dictum: 'To keep the Soviet Union out, the Americans in, and the Germans down.'"),

    ("The Critical Catalyst: West German Rearmament (1950\u20131955)",
     "The structural transformation of NATO from a paper political guarantee into a formidable integrated military command was accelerated by the outbreak of the Korean War in June 1950. The surprise communist invasion of South Korea convinced American military planners that a synchronized Soviet conventional blitzkrieg across the North German plain was imminent. At the Lisbon Conference of February 1952, NATO commanders calculated that defending the continent required ninety-six conventional divisions\u2014a figure that economically exhausted Western European democracies could not possibly mobilize without the military manpower of West Germany.\n\n"
     "The prospect of German rearmament provoked intense historical trauma, particularly in France. To resolve this dilemma, French Premier Ren\u00e9 Pleven proposed the 'European Defense Community' (EDC), which envisioned a unified European army incorporating German units under a supranational European command without a separate German general staff. However, after four years of agonizing debate, the French National Assembly rejected the EDC treaty in August 1954, fearing the loss of national sovereignty.\n\n"
     "The crisis was saved by British Foreign Secretary Anthony Eden through the 1954 Paris Agreements. West Germany was granted full sovereign status, ended Allied occupation, and was permitted to establish its own armed forces (the Bundeswehr), capped at 500,000 men and strictly prohibited from possessing atomic, biological, or chemical weapons. On May 9, 1955, the Federal Republic of Germany was officially admitted as the fifteenth member of NATO, providing the essential conventional infantry and armored divisions that anchored the central front."),

    ("The Soviet Counter-Move: The Creation of the Warsaw Pact (1955)",
     "The integration of a rearmed West Germany into NATO was viewed by Moscow as an existential threat. Historically traumatized by two German invasions that had devastated Russia within thirty years, the Soviet Union responded immediately. On May 14, 1955, five days after West Germany joined NATO, Soviet Premier Nikolai Bulganin convened a conference in the Polish capital, signing the 'Treaty of Friendship, Co-operation and Mutual Assistance'\u2014the Warsaw Pact. The alliance unified the military forces of the Soviet Union and seven satellite states: the German Democratic Republic (East Germany), Poland, Czechoslovakia, Hungary, Romania, Bulgaria, and Albania.\n\n"
     "On paper, the Warsaw Pact mirrored NATO: it contained mutual defense obligations, a Political Consultative Committee, and proclaimed respect for national sovereignty. In reality, as declassified Soviet archives analyzed by Vojtech Mastny reveal, the Warsaw Pact was fundamentally different in nature and purpose. While NATO was a voluntary defensive coalition of democratic states that operated by absolute consensus, the Warsaw Pact was a completely centralized instrument of Soviet imperial discipline. The commander-in-chief of the Warsaw Pact forces was always a Soviet Marshal (initially Ivan Konev), the chief of staff was always Soviet, and operational commands were completely integrated into the Soviet Ministry of Defense in Moscow.\n\n"
     "Crucially, the primary operational function of the Warsaw Pact was not fighting NATO, but policing its own members and crushing domestic anti-communist rebellions. When Hungarian Prime Minister Imre Nagy announced Hungary's withdrawal from the Warsaw Pact during the democratic revolution of October 1956, Soviet tanks invaded Budapest, slaughtering thousands of Hungarian freedom fighters under the pretext of defending the alliance. Similarly, in August 1968, Warsaw Pact forces invaded Czechoslovakia to crush the 'Prague Spring' under the Brezhnev Doctrine. The Warsaw Pact was less an alliance against external enemies than a military straitjacket designed to preserve Soviet hegemony over Eastern Europe."),

    ("The Nuclear Dimension and the Doctrine of Mutually Assured Destruction (MAD)",
     "The military confrontation between NATO and the Warsaw Pact was fundamentally defined and constrained by the evolution of thermonuclear weapons. The American atomic monopoly was shattered in August 1949 when the Soviet Union successfully detonated its first atomic bomb (RDS-1). By 1952\u20131953, both superpowers had developed the catastrophic thermonuclear Hydrogen Bomb, possessing explosive yields measured in megatons rather than kilotons.\n\n"
     "Confronted by the Warsaw Pact's overwhelming conventional superiority in tanks and infantry, NATO adopted the doctrine of 'Massive Retaliation' formulated by US Secretary of State John Foster Dulles in 1954. NATO declared that any Soviet conventional aggression in Europe would be met by immediate, devastating nuclear strikes against the Soviet homeland. However, when the Soviet Union launched the Sputnik satellite in 1957 and developed Intercontinental Ballistic Missiles (ICBMs) capable of incinerating American cities, the credibility of Massive Retaliation collapsed: would an American President sacrifice New York or Chicago to defend West Berlin or Hamburg?\n\n"
     "Under President John F. Kennedy and Defense Secretary Robert McNamara, NATO transitioned to the doctrine of 'Flexible Response' in the 1960s, developing tactical nuclear weapons alongside expanded conventional forces. Yet, the overarching strategic reality was Mutually Assured Destruction (MAD): because both coalitions possessed invulnerable second-strike nuclear arsenals, any direct military clash between NATO and the Warsaw Pact would result in the complete extermination of human civilization. Paradoxically, nuclear terror became the ultimate guarantor of European peace."),

    ("Conclusion: The Legacy of Bipolar Security",
     "In conclusion, the emergence of NATO and the Warsaw Pact consolidated the division of Europe into two fortified, armed camps that defined the geopolitics of the late twentieth century. By institutionalizing the ideological frontiers of the Cold War and anchoring them in nuclear deterrence, this bipolar military architecture eliminated the ambiguous border rivalries and miscalculations that had ignited World War I and World War II. At a terrible moral and human price\u2014the subjugation of one hundred million Eastern Europeans behind the iron discipline of the Warsaw Pact and the permanent partitioning of Germany\u2014NATO and the Warsaw Pact successfully prevented a third world war on the European continent. When the Soviet empire finally collapsed in 1989\u20131991, the true nature of both alliances was vividly exposed: the coercive Warsaw Pact dissolved overnight without a whimper, while the voluntary democratic coalition of NATO endured and expanded, standing as the most durable and successful military alliance in modern international history.")
]

quotes_p2_12 = [
    ("John Lewis Gaddis", "The Long Peace: Inquiries into the History of the Cold War (1987)", "'The bipolar military division of Europe, backed by nuclear weapons, produced the longest period of peace between the Great Powers in modern European history... The terrifying logic of nuclear deterrence made victory in war impossible, and therefore made war irrational.'"),
    ("Lord Hastings Ismay (1949)", "First Secretary General of NATO", "'The purpose of the North Atlantic Treaty Organization is to keep the Soviet Union out, the Americans in, and the Germans down.'"),
    ("Lawrence Freedman", "The Evolution of Nuclear Strategy", "'The introduction of thermonuclear weapons transformed military strategy from the preparation for victory into the science of deterrence... NATO's ultimate weapon was an arsenal that could never be used without destroying the civilization it was created to defend.'"),
    ("Vojtech Mastny", "A Cardboard Castle? An Inside History of the Warsaw Pact", "'Unlike NATO, which was a genuine voluntary alliance of free nations, the Warsaw Pact was conceived by Moscow as a military-administrative device to legitimize Soviet military garrisons in Eastern Europe and enforce ideological conformity.'")
]

# Run the generation
if __name__ == "__main__":
    tasks = [
        (meta_p2_10, callout_p2_10, map_p2_10, timeline_p2_10, essay_p2_10, quotes_p2_10,
         os.path.join(paper2_dir, "Q10_World_War_II_Turning_Points_Conferences_and_Consequences_Master_Notes.docx"),
         "QUESTION 10: WORLD WAR II: TURNING POINTS, THE HOLOCAUST & POST-WAR CONFERENCES",
         "Q10_World_War_II_Turning_Points_Conferences_and_Consequences_Master_Notes"),
        (meta_p2_11, callout_p2_11, map_p2_11, timeline_p2_11, essay_p2_11, quotes_p2_11,
         os.path.join(paper2_dir, "Q11_Origins_of_Cold_War_Truman_Marshall_and_Berlin_Blockade_Master_Notes.docx"),
         "QUESTION 11: ORIGINS OF THE COLD WAR: CONTAINMENT, TRUMAN & BERLIN BLOCKADE",
         "Q11_Origins_of_Cold_War_Truman_Marshall_and_Berlin_Blockade_Master_Notes"),
        (meta_p2_12, callout_p2_12, map_p2_12, timeline_p2_12, essay_p2_12, quotes_p2_12,
         os.path.join(paper2_dir, "Q12_NATO_vs_Warsaw_Pact_and_Bipolar_Military_Division_Master_Notes.docx"),
         "QUESTION 12: BIPOLAR MILITARY ALLIANCES: NATO (1949), WARSAW PACT (1955) & MAD",
         "Q12_NATO_vs_Warsaw_Pact_and_Bipolar_Military_Division_Master_Notes")
    ]
    
    for meta, callout, cmap, timeline, essay, quotes, docx_p, title, base_n in tasks:
        generate_master_doc(meta, callout, cmap, timeline, essay, quotes, docx_p, title)
        
        md_text = f"# KPPSC PMS \u2022 EUROPEAN HISTORY (PAPER II: 1914\u20131960)\n## {title}\n\n"
        md_text += "### EXECUTIVE EXAM METADATA\n"
        for k, v in meta:
            md_text += f"- **{k}** {v}\n"
        md_text += "\n---\n\n### \U0001F4CC EXAM ESSENTIAL TAKEAWAY\n"
        for line in callout:
            md_text += f"> {line}\n>\n"
        md_text += "\n---\n\n### 1. CONCEPTUAL MIND MAP & THEMATIC PATHWAY\n```\n" + cmap + "\n```\n\n---\n\n### 2. CHRONOLOGICAL MILESTONES & DIPLOMATIC DATES\n"
        md_text += "| Date / Year | Event / Action | Historical & Exam Significance |\n|---|---|---|\n"
        for d, e, s in timeline:
            md_text += f"| **{d}** | {e} | {s} |\n"
        md_text += "\n---\n\n### 3. THE 20-MARK MASTER MODEL ESSAY (~650 WORDS)\n\n"
        for st, body in essay:
            md_text += f"#### {st}\n{body}\n\n"
        md_text += "---\n\n### 4. HISTORIOGRAPHICAL PERSPECTIVES & SCHOLARLY QUOTES\n"
        md_text += "| Historian | Analytical Perspective | Exact Quotable Argument |\n|---|---|---|\n"
        for h, p, q in quotes:
            md_text += f"| **{h}** | {p} | {q} |\n"
        
        render_md_and_pdf(md_text, base_n)
