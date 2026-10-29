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
paper1_dir = r"H:\My Drive\PMS KPK 2022 notes\1. European History Paper 1 and Paper 2\Paper 1"

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
    r_main = p_title.add_run("KPPSC PMS \u2022 EUROPEAN HISTORY (PAPER I)")
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
    r_f = p_foot.add_run("Khyber Pakhtunkhwa Public Service Commission (KPPSC) \u2022 PMS Competitive Examination Notes \u2022 European History")
    r_f.font.name = "Arial"
    r_f.font.size = Pt(8.5)
    r_f.font.color.rgb = RGBColor(148, 163, 184)
    
    doc.save(out_path)
    print("Generated Master Docx:", out_path)

def render_md_and_pdf(md_text, base_name):
    md_path = os.path.join(paper1_dir, f"{base_name}.md")
    pdf_path = os.path.join(paper1_dir, f"{base_name}.pdf")
    
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
  Khyber Pakhtunkhwa Public Service Commission (KPPSC) &bull; PMS European History &bull; Paper I Master Notes
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
# QUESTION 13: 2018 Q.4 / 2013 Q.5 (REVOLUTIONS OF 1848 - SPRING OF NATIONS)
# =========================================================================

meta_13 = [
    ("EXAM QUESTIONS LINKED:", "2018 Q.4: 'Discuss the Revolution of 1848 in France and explain why the revolutionary wave engulfed most of Europe.'\n2013 Q.5: 'Describe the effects of the Revolution of 1848, which transformed the whole European world.'"),
    ("SYLLABUS TOPIC:", "Paper I \u2014 Section 3: The Revolutions of 1848 ('The Spring of Nations') \u2014 Causes, Spread, Collapse & Transformation"),
    ("ALLOCATED MARKS / TIME:", "20 Marks | 35 Minutes | Standard Word Count: ~650\u2013700 words"),
    ("HISTORIOGRAPHICAL THESIS:", "The Revolutions of 1848 constituted the most widespread yet ephemeral revolutionary upheaval in modern European history, ignited by agrarian-commercial distress and an ideological coalition of middle-class liberalism, working-class socialism, and romantic nationalism; while swiftly crushed within eighteen months due to internal class fissures and the resurgence of dynastic armies, 1848 irrevocably destroyed the feudal social order, shattered the Metternichian Concert of Europe, and ushered in the era of authoritarian Realpolitik.")
]

callout_13 = [
    "\u2022 The Epicenter (France, Feb\u2013June 1848): Collapse of Louis Philippe's corrupt 'bourgeois monarchy' over the suppression of the banquet campaigns. Second Republic established; Louis Blanc's National Workshops experiment terminated; violent class war of the 'June Days' (over 3,000 workers killed, 12,000 deported), paving the path for Louis Napoleon Bonaparte's authoritarian presidency.",
    "\u2022 The Continental Conflagration: Over 50 separate revolutions exploded in weeks: Metternich forced to flee Vienna in disguise; Hungarian autonomy proclaimed under Lajos Kossuth; Prussian King Friedrich Wilhelm IV forced to uncover his head to fallen Berlin revolutionaries; Frankfurt Parliament convened; Milan (Five Days) and Venice expel Austrian garrisons.",
    "\u2022 Why Did It Engulf Europe? Shared structural vulnerability: 1845\u20131847 potato blight and agrarian famine ('Hungry Forties'), industrial overproduction slump, explosive urban population growth, expansion of railway and telegraph networks transmitting revolutionary news in hours.",
    "\u2022 The Anatomy of Defeat: The fatal divorce between moderate property-owning liberals (terrified of proletarian socialism) and radical workers; ethnic nationalism pitting revolutionaries against each other (Magyars vs. Croats/Serbs); unshakeable loyalty of professional dynastic armies (Radetzky, Windisch-Gr\u00e4tz); and Russian military intervention under Tsar Nicholas I.",
    "\u2022 The Permanent Transformation: Abolition of Austrian serfdom survived; Prussian and Piedmontese constitutions endured; established the political agenda (Italian and German unifications) accomplished two decades later via 'Blood and Iron' rather than parliamentary idealism."
]

map_13 = """========================================================================================
SOCIOECONOMIC CATALYSTS               EXPLOSION & WAVE (SPRING 1848)       REACTION & COLLAPSE (1849)
========================================================================================
[1845-47 Agrarian Famine]          --> [Paris (Feb): Monarchy Falls]   --> [June Days Class War in Paris]
[Urban Proletariat Slump]          --> [Vienna (Mar): Metternich Flees]--> [Radetzky Crushes Italy (Custoza)]
[Disenfranchised Bourgeoisie]      --> [Berlin (Mar): Barricades Win]  --> [Windisch-Gr\u00e4tz Re-takes Vienna]
[Nationalist Ferment (Italy/Ger)]  --> [Frankfurt Parliament Meets]    --> [Prussian King Rejects Crown]
                                   --> [Budapest: Kossuth Republic]    --> [Russian Invasion Crushes Hungary]
========================================================================================"""

timeline_13 = [
    ("Jan\u2013Feb 1848", "Palermo & Paris Uprisings", "Sicily revolts against Bourbons; Parisian banquet banned; Louis Philippe abdicates; Second Republic declared."),
    ("March 13, 1848", "Fall of Metternich in Vienna", "Students and workers storm the Hofburg; Prince Metternich resigns after 33 years in power and flees to England."),
    ("March 18\u201322, 1848", "Berlin & Milan Barricades", "Prussian King yields to street fighters; Milan's 'Five Days' forces Austrian Field Marshal Radetzky to retreat."),
    ("May 1848", "Frankfurt Parliament Convenes", "Over 800 German intellectuals convene to draft a unified, liberal German constitution (Kleindeutsch vs. Grossdeutsch)."),
    ("June 23\u201326, 1848", "The June Days in Paris", "General Cavaignac's troops butcher Parisian workers protesting closure of National Workshops; liberal-socialist rupture."),
    ("July 1848", "Battle of Custoza", "Austrian Field Marshal Radetzky decisively defeats King Charles Albert of Piedmont-Sardinia, reclaiming Lombardy."),
    ("Nov\u2013Dec 1848", "Rise of Reaction & Louis Napoleon", "Prince Schwarzenberg restores Austrian imperial control; Louis Napoleon elected French President by landslide."),
    ("April 1849", "The 'Crown from the Gutter'", "King Friedrich Wilhelm IV of Prussia rejects imperial crown offered by Frankfurt Parliament, killing liberal German unification."),
    ("June\u2013Aug 1849", "Russian Intervention in Hungary", "Tsar Nicholas I deploys 200,000 Russian troops to crush Kossuth's Hungarian Republic; surrender at Vil\u00e1gos."),
    ("Aug 1849", "Fall of Roman & Venetian Republics", "French troops overthrow Mazzini's Roman Republic; cholera and Austrian siege crush Venetian Republic.")
]

essay_13 = [
    ("Introduction: The Springtime and Winter of European Peoples",
     "The year 1848 witnessed the most synchronized, geographically pervasive revolutionary wave in European history, aptly designated the 'Spring of Nations' (V\u00f6lkerfr\u00fchling). From Paris to Budapest, and from Berlin to Palermo, dynastic regimes established at the Congress of Vienna collapsed in mere weeks before an insurgent combination of middle-class constitutionalism, proletarian radicalism, and romantic nationalism. Yet, within eighteen months, this astonishing continental triumph had deteriorated into catastrophic defeat, prompting A.J.P. Taylor's celebrated aphorism that 'German history reached its turning-point and failed to turn.' Nevertheless, dismissing 1848 as an outright failure is an error of historical perspective. As Eric Hobsbawm convincingly argues, 1848 was the watershed moment when traditional dynastic politics expired. While liberal intellectuals failed to capture power through parliamentary resolutions, the revolutions permanently liquidated European feudalism, shattered the conservative Metternichian equilibrium, and compelled monarchies to adopt the methods of mass politics, economic modernization, and state-directed nationalism."),

    ("The Epicenter: The February Revolution and the June Days in France",
     "The continental conflagration was ignited in Paris in February 1848. The constitutional July Monarchy of Louis Philippe and his Prime Minister Fran\u00e7ois Guizot had degenerated into a venal plutocracy serving an electoral oligarchy of merely 200,000 wealthy landowners. Guizot's cynical dismissal of democratic demands\u2014'Enrichissez-vous' (Enrich yourselves)\u2014provoked the political opposition to organize a nationwide 'Banquet Campaign.' When the government banned the Parisian banquet on February 22, violent street clashes erupted, barricades multiplied, the National Guard defected, and Louis Philippe abdicated on February 24. The newly declared Second Republic initially attempted an unprecedented socialist experiment: under pressure from Louis Blanc and the radical working class, the Provisional Government established 'National Workshops' (Ateliers Nationaux) guaranteeing the 'Right to Work' (droit au travail) to over 100,000 unemployed artisans.\n\n"
     "However, the April elections for the Constituent Assembly produced an overwhelmingly conservative, bourgeois-rural majority terrified of Parisian radicalism. On June 21, the government announced the dissolution of the National Workshops. The working class responded with the desperate, spontaneous insurrection known as the 'June Days' (June 23\u201326). Commanded by General Louis-Eug\u00e8ne Cavaignac, the army and bourgeois mobile guards engaged in brutal street warfare, slaughtering over 3,000 workers and arresting 15,000. The June Days were a turning point of immense ideological significance: Karl Marx identified it as the first open class war between the bourgeoisie and the industrial proletariat. The trauma of the June Days alienated the French masses from parliamentary republicanism, paving the way for the December 1848 presidential election of Louis-Napol\u00e9on Bonaparte, who capitalized on peasant conservatism and urban fatigue to dismantle the Republic by coup d'\u00e9tat in 1851."),

    ("The Continental Wildfire: Why the Wave Engulfed Europe",
     "The unprecedented speed with which the revolutionary impulse swept across Europe in March 1848 resulted from the convergence of acute socioeconomic misery and shared ideological aspirations. The preceding years (1845\u20131847) were marked by catastrophic agricultural crop failures, catastrophic potato blight across Central and Western Europe, and a deep commercial credit contraction that doubled basic food prices while shuttering nascent manufacturing shops. This 'Hungry Forties' crisis created an explosive subterranean powder keg of peasant and artisan desperation. Simultaneously, the rapid expansion of railways, steam printing, and telegraphic communication enabled news of the Parisian revolution to reach Vienna, Berlin, and Milan within hours, demoralizing dynastic rulers while electrifying urban crowds.\n\n"
     "In the Austrian Habsburg Empire, students and workers surged into the Hofburg on March 13, compelling Prince Klemens von Metternich\u2014the living embodiment of conservative restoration\u2014to resign after thirty-three years of rule and flee in disguise to London. In Hungary, the fiery orator Lajos Kossuth proclaimed the 'March Laws,' establishing near-total national autonomy. In Prussia, King Friedrich Wilhelm IV was forced by barricade combatants to remove his hat before the corpses of slain citizens and declare that 'Prussia is henceforth merged into Germany.' In Frankfurt, over eight hundred German jurists and professors convened the Frankfurt National Parliament to draft a liberal, constitutional unification of Germany. In Italy, revolts in Milan ('Five Days') and Venice compelled Austrian garrisons to evacuate, while King Charles Albert of Piedmont-Sardinia declared war on Austria to liberate the peninsula."),

    ("The Anatomy of Collapse: Why Did the Revolutions Fail?",
     "Despite their electrifying initial victories, the revolutions collapsed with equal velocity between autumn 1848 and summer 1849 due to three fatal structural vulnerabilities:\n\n"
     "1. The Class Schism Between Liberals and Proletarians: The fundamental paradox of 1848 was that the revolutions were made by working-class barricade fighters, but leadership was assumed by moderate bourgeois intellectuals and lawyers. Terrified by the radical specter of socialism and property redistribution vividly exposed during the Paris June Days, moderate liberals in Berlin, Vienna, and Frankfurt recoiled from their working-class allies and retreated into the protective embrace of traditional monarchical authority.\n\n"
     "2. Nationalist and Ethnic Antagonisms: Rather than creating a fraternal union of liberated peoples, romantic nationalism splintered the revolutionary front. In the Austrian Empire, Kossuth's Hungarian government insisted on fierce Magyar dominance, alienating Croats, Serbs, and Romanians, who promptly aligned with the Austrian Crown under Josip Jela\u010di\u0107 to crush the Hungarians. At the Frankfurt Parliament, German liberals debated whether to include Austria (Grossdeutschland) or exclude it (Kleindeutschland), debating endlessly while Prussian and Austrian monarchs gathered military strength.\n\n"
     "3. Cohesion of Standing Armies and Russian Intervention: While dynastic governments had been momentarily paralyzed, their professional aristocratic officer corps and peasant conscripts remained fundamentally loyal. Field Marshal Radetzky shattered Piedmontese ambitions at Custoza (July 1848) and Novara (March 1849); Field Marshal Windisch-Gr\u00e4tz bombarded and recaptured Prague and Vienna. Finally, when Hungary proclaimed full independence in April 1849, Austrian Emperor Franz Joseph appealed to Tsar Nicholas I. The Russian 'Gendarme of Europe' dispatched two hundred thousand troops into Hungary, crushing the republican armies at Vil\u00e1gos in August 1849."),

    ("Transformational Legacy: The Birth of Realpolitik",
     "The ultimate significance of 1848 lies not in its immediate defeats, but in its systemic destruction of the Ancien R\u00e9gime. Socially, the Austrian Empire permanently abolished serfdom and robot labor dues, emancipating millions of peasants and removing the social foundation of feudalism. Constitutionally, monarchies recognized that absolute rule was obsolete: Prussia retained its 1850 constitutional charter (with its three-class franchise), and Piedmont-Sardinia preserved the Statuto Albertino, transforming Piedmont into the liberal magnetic pole for Italian unification. Most critically, the failure of 1848 ended the romantic, idealistic phase of European politics. Statesmen like Count Cavour in Piedmont and Otto von Bismarck in Prussia observed the impotence of the Frankfurt and Roman assemblies. They realized that European boundaries would not be reshaped by parliamentary speeches and majority votes, but by state power, diplomatic maneuvering, and 'Blood and Iron'\u2014the triumph of Realpolitik."),

    ("Conclusion: The Critical Watershed of Modern European History",
     "In conclusion, the Revolutions of 1848 were neither an irrelevant episode nor a tragic waste of human life; they constituted the definitive turning point in the modernization of Europe. The 'Spring of Nations' demonstrated that dynastic legitimacy could no longer govern European states without popular or institutional accommodation. By stripping the conservative order of its Metternichian illusions, 1848 prepared the terrain for the mass political mobilizations of the late nineteenth century. The democratic and socialist demands formulated in the Parisian clubs and the national agendas articulated in Frankfurt, Budapest, and Rome became the unavoidable blueprint for modern European state-building, confirming that while the revolutionaries of 1848 lost the battles of the present, they irrevocably determined the architecture of the future.")
]

quotes_13 = [
    ("Eric Hobsbawm", "The Age of Capital (1975)", "'The revolution of 1848 was the only one in the history of Europe to affect both its developed and undeveloped parts... It was the most widespread and the least successful, yet it marked the definitive end of the pre-industrial dynastic order.'"),
    ("Lewis Namier", "1848: The Revolution of the Intellectuals", "'1848 was essentially a revolution of the intellectuals... and its tragedy was that middle-class liberal nationalism proved completely intolerant of other nationalities and terrified of social revolution.'"),
    ("A.J.P. Taylor", "The Course of German History", "'German history reached its turning point in 1848 and failed to turn... The Frankfurt professors debated world philosophy while the Prussian soldiers sharpened their swords.'"),
    ("Jonathan Sperber", "The European Revolutions, 1848-1851", "'Rather than looking forward to twentieth-century revolutions, 1848 was the last and greatest act of the revolutionary era that began in 1789, terminating the classic politics of the barricade.'")
]

# =========================================================================
# QUESTION 14: 2018 Q.5 (FOREIGN POLICY OF NAPOLEON III)
# =========================================================================

meta_14 = [
    ("EXAM QUESTIONS LINKED:", "2018 Q.5: 'Evaluate the foreign policy of Napoleon III with special reference to his role in Italian unification and the Franco-Prussian War.'"),
    ("SYLLABUS TOPIC:", "Paper I \u2014 Section 4: France Under the Second Empire (1852\u20131870) \u2014 Foreign Policy, Revisionism, Italian Risorgimento & Franco-Prussian War"),
    ("ALLOCATED MARKS / TIME:", "20 Marks | 35 Minutes | Standard Word Count: ~650\u2013700 words"),
    ("HISTORIOGRAPHICAL THESIS:", "Driven by an erratic synthesis of romantic nationalism, Napoleonic prestige, and opportunistic revisionism aimed at dismantling the 1815 Treaties of Vienna, Napoleon III's foreign policy initially established France as the arbiter of European diplomacy in the Crimea and Northern Italy, but ultimately destroyed the conservative equilibrium, alienated every potential European ally, and unwittingly catalyzed the unification of Italy and Hohenzollern Germany, culminating in the catastrophic collapse of the Second Empire at Sedan in 1870.")
]

callout_14 = [
    "\u2022 Overarching Foreign Policy Goal: The destruction of the 1815 Vienna Settlement that had contained and humiliated France; the elevation of France to continental hegemony through the championing of the 'Principle of Nationalities' (liberation of oppressed European peoples).",
    "\u2022 The Crimean War (1853\u20131856): Napoleon III joined Great Britain to check Russian expansionism; hosted the Congress of Paris (1856), replacing the Holy Alliance with French diplomatic leadership and splitting Austria from Russia.",
    "\u2022 Italian Intervention & Ambiguity (1858\u20131860): Signed the secret Plombi\u00e8res Agreement with Cavour; defeated Austria at Magenta and Solferino (1859); signed the sudden Armistice of Villafranca fearing Prussian mobilization and Catholic backlash; acquired Nice and Savoy (Treaty of Turin 1860), alienating Britain; maintained a French military garrison in Rome to protect the Papacy until 1870, permanently alienating the newly unified Kingdom of Italy.",
    "\u2022 Imperial Overreach & Mexico (1861\u20131867): Attempted to establish a Catholic Mexican Empire under puppet Archduke Maximilian; collapsed in financial bankruptcy and American diplomatic ultimatum under the Monroe Doctrine, ending in Maximilian's execution and devastating French prestige.",
    "\u2022 Bismarck's Nemesis & Franco-Prussian War (1870): Outmaneuvered during the 1866 Austro-Prussian War by demanding 'compensations' (Luxembourg, Rhineland, Belgium) which Bismarck leaked; isolated France completely; baited into declaring war by Bismarck's doctored Ems Telegram; crushed and captured at the Battle of Sedan (September 2, 1870)."
]

map_14 = """========================================================================================
FOREIGN POLICY GOALS                  PHASE 1: ASCENT (1853-1860)       PHASE 2: DOWNFALL (1861-1870)
========================================================================================
[Revision of 1815 Vienna Treaties]--> [Crimean War & Treaty of Paris] --> [Mexican Fiasco: Maximilian Shot]
[Principle of Nationalities]      --> [Plombi\u00e8res Pact with Cavour]   --> [Luxembourg/Rhineland Demands]
[French Imperial Gloire]          --> [Battles of Magenta & Solferino]--> [Diplomatic Isolation: No Allies]
[Mediterranean & Catholic Power]  --> [Annexation of Nice & Savoy]    --> [Bismarck's Ems Telegram Trap]
                                  --> [Garrisoning Rome (Papal Guard)]--> [Battle of Sedan: Empire Dies]
========================================================================================"""

timeline_14 = [
    ("Dec 2, 1852", "Proclamation of Second Empire", "Louis Napoleon proclaims himself Napoleon III; announces 'The Empire means peace' (L'Empire, c'est la paix)."),
    ("1854\u20131856", "The Crimean War", "Allies with Britain against Russia; Siege of Sevastopol; presides as European master at Congress of Paris (1856)."),
    ("July 21, 1858", "Secret Pact of Plombi\u00e8res", "Napoleon III and Count Cavour plot war against Austria to create an Italian Confederation under French protection."),
    ("June 1859", "Battles of Magenta & Solferino", "French forces defeat Austria in bloody battles; horrified by casualties and Prussian maneuvers, Napoleon signs Villafranca."),
    ("March 1860", "Treaty of Turin", "Cavour cedes Nice and Savoy to France in exchange for French recognition of central Italian annexations; alienates Great Britain."),
    ("1861\u20131867", "The Mexican Adventure", "Dispatches 30,000 troops to Mexico to install Emperor Maximilian; humiliating withdrawal; Maximilian executed (1867)."),
    ("July 1866", "Austro-Prussian War Neutrality", "Napoleon remains neutral expecting a prolonged stalemate; Bismarck wins in 7 weeks; denies France any 'compensations'."),
    ("1867", "Luxembourg Crisis", "Napoleon attempts to purchase Luxembourg from Dutch King; Bismarck blocks it publicly, humiliating Napoleon."),
    ("July 13, 1870", "The Ems Telegram", "Bismarck edits King Wilhelm I's dispatch to humiliate France; Parisian crowds demand war; Napoleon declares war on July 19."),
    ("Sept 2, 1870", "Catastrophe at Sedan", "Napoleon III surrounded and captured with 83,000 soldiers; Second Empire collapses; Third Republic declared in Paris.")
]

essay_14 = [
    ("Introduction: The Enigma and Ambition of Napoleon III",
     "Napoleon III remains one of the most enigmatic and paradoxical rulers of nineteenth-century Europe. Ascending to the imperial throne in 1852 with the reassuring proclamation that 'The Empire means peace' (L'Empire, c'est la paix), he systematically pursued a foreign policy characterized by relentless revisionism, military interventionism, and diplomatic audacity. Imbued with romantic nationalism, Saint-Simonian economic progressivism, and an obsessive desire to revive the imperial gloire of his uncle, Napoleon I, his paramount strategic ambition was the complete dismantling of the 1815 Vienna Settlement, which had contained France within cordons sanitaires and reduced her continental influence. In the historiographical assessment, scholars such as David Thomson and Henry Kissinger recognize that while Napoleon III achieved brilliant initial triumphs in the Crimean War and early Italian campaigns, his fatal intellectual incoherence, vacillation, and ideological adventurism ultimately shattered the European balance of power, isolated France diplomatically, and unleashed the unified Prussian powerhouse that brought his ruin at Sedan."),

    ("The Crimean Triumph and the Congress of Paris (1853\u20131856)",
     "The opening act of Napoleon III's foreign policy was an unmitigated diplomatic triumph. Capitalizing on a dispute between Catholic and Orthodox monks over the holy places in Ottoman Jerusalem, Napoleon III engineered an alliance with Great Britain against Tsarist Russia, precipitating the Crimean War (1853\u20131856). While the military campaign in the Crimea was marked by catastrophic logistical blunders and horrific cholera mortality, the capture of Sevastopol by French troops re-established the prestige of the French army. More decisively, Napoleon III hosted the Congress of Paris (1856), presiding as the supreme arbiter of European diplomacy. The Treaty of Paris neutralized the Black Sea, preserved the Ottoman Empire, and shattered the Holy Alliance that had bound Russia, Austria, and Prussia in counter-revolutionary solidarity since 1815. Austria's refusal to aid Russia alienated St. Petersburg, leaving Central Europe diplomatically atomized and providing Napoleon III with the strategic freedom to advance his next ambition: the liberation of Italy."),

    ("Role in Italian Unification: The Plombi\u00e8res Alliance and the Betrayal of Villafranca",
     "Napoleon III possessed a genuine, romantic sympathy for the Italian Risorgimento, having fought alongside the Carbonari in 1831. However, his vision for Italy was not a united, powerful unitary state on France's southeastern border, but a loose confederation of four states presided over by the Pope and dominated militarily by France. In July 1858, Napoleon met secretly with Count Camillo di Cavour at Plombi\u00e8res, agreeing to provoke Austria into war. In exchange for French military intervention to expel Austria from Lombardy and Venetia, Piedmont-Sardinia promised to cede Nice and Savoy to France.\n\n"
     "The Franco-Austrian War of 1859 unfolded rapidly: French armies won bloody victories at Magenta and Solferino (June 1859). However, horrified by the carnage of Solferino (which inspired the founding of the Red Cross), alarmed by Prussian troop mobilizations along the Rhine, and terrified that Italian revolutionary unrest would consume the Papal States, Napoleon III abruptly signed the separate Armistice of Villafranca with Austrian Emperor Franz Joseph without consulting Cavour. Lombardy was ceded to Piedmont, but Venetia remained under Austrian subjugation. Cavour was furious, yet Napoleon's betrayal could not halt the Italian momentum: central Italian duchies voted for annexation with Piedmont. Under the Treaty of Turin (1860), Napoleon accepted these annexations in exchange for Nice and Savoy. While France gained territory, the political cost was catastrophic: the annexation of Nice and Savoy deeply alarmed Great Britain, which began to view Napoleon as an aggressive imperialist predator, while Italian patriots bitterly resented his premature abandonment of their cause."),

    ("The Contradiction of Rome: Alienating the Italian Kingdom",
     "Napoleon III's Italian policy was paralyzed by an irreconcilable domestic contradiction. While he wished to be the champion of Italian national liberation, he depended politically upon French Catholic and conservative support at home. Consequently, when Garibaldi's Expedition of the Thousand united Southern Italy in 1860, Napoleon dispatched a French military garrison to Rome to defend Pope Pius IX's temporal sovereignty. For the next ten years, French bayonets prevented the Italian Kingdom from claiming Rome as its historic capital. When Garibaldi attempted to take the city in 1867, French troops slaughtered Italian volunteers at the Battle of Mentana. Napoleon thus achieved the absolute worst diplomatic outcome: he had helped create a united Kingdom of thirty million Italians on his frontier, but ensured that this new Italian state regarded France with profound hostility and distrust, preventing any future Franco-Italian defensive alliance."),

    ("Imperial Overreach: The Mexican Adventure (1861\u20131867)",
     "While his European diplomacy faltered, Napoleon III embarked on what he termed the 'grandest enterprise of the reign': the Mexican Expedition. Taking advantage of the American Civil War, which prevented Washington from enforcing the Monroe Doctrine, Napoleon deployed thirty thousand French troops to overthrow the republican government of Benito Ju\u00e1rez and install Austrian Archduke Maximilian as Emperor of Mexico. His fantasy was to create a Catholic Latin empire in North America to counter the Anglo-Saxon Protestant United States and provide markets for French industry. Instead, the campaign degenerated into a brutal, unwinnable guerrilla quagmire. Once the American Civil War concluded in 1865, the United States massed veteran armies along the Rio Grande and issued an ultimatum. Napoleon ignominiously evacuated his troops in 1866, abandoning the abandoned Maximilian to face a Mexican firing squad at Quer\u00e9taro in June 1867. The Mexican debacle depleted the French treasury, degraded army readiness, and severely diminished Napoleon's international prestige precisely as the Prussian threat reached acute proportions."),

    ("The Nemesis: Bismarck's Trap and the Franco-Prussian War (1870)",
     "The catastrophic finale of Napoleon III's foreign policy unfolded in his collision with Otto von Bismarck. During the Austro-Prussian War of 1866, Napoleon remained neutral, mistakenly expecting a protracted war of attrition that would allow him to intervene as arbiter and demand the Rhineland as 'compensation.' When Prussia obliterated Austria in seven weeks at K\u00f6niggr\u00e4tz, Napoleon was stunned. Bismarck subsequently outmaneuvered him at every turn: in 1867, Bismarck thwarted Napoleon's bid to purchase Luxembourg from the Dutch King, humiliating him publicly, and leaked Napoleon's written demands for Belgium and the Rhineland to the British, ensuring London's complete neutrality in any future war.\n\n"
     "In July 1870, the Hohenzollern candidature for the vacant Spanish throne offered Napoleon a diplomatic victory when Prince Leopold withdrew under French pressure. However, pressured by an aggressive press and court hawks (Empress Eug\u00e9nie), Napoleon demanded an insulting personal guarantee from King Wilhelm I that no Hohenzollern would ever claim Spain. Bismarck edited the King's courteous refusal into the stinging 'Ems Telegram,' releasing it on Bastille Day. Insulted and panicked, France declared war on July 19, 1870, with zero allies: Russia was hostile, Austria-Hungary hesitated, Italy demanded Rome as a precondition, and Britain was disgusted by French designs on Belgium. The French military machine proved hopelessly disorganized against Helmuth von Moltke's mobilized armies. On September 2, 1870, at the Battle of Sedan, Napoleon III was encircled, surrendered alongside eighty-three thousand French troops, and taken prisoner. Two days later, the Second Empire was deposed in Paris, and the Third Republic was declared."),

    ("Conclusion: The Tragic Legacy of Bonapartist Diplomacy",
     "In conclusion, the foreign policy of Napoleon III was a tragic study in the peril of ideological revisionism unmoored from realistic grand strategy. In his obsession with revising the 1815 settlement and championing the principle of nationalities, he destroyed the conservative concert that had preserved European peace and French security. By alienating Britain over Nice, estranging Italy over Rome, abandoning Austria after 1859, and outraging Russia in the Crimea, he dismantled every diplomatic barrier protecting France. Consequently, he paved the way for Bismarck to construct a united, militarized German Empire upon the ruins of French continental hegemony. The Treaty of Frankfurt (1871), which annexed Alsace and Lorraine to Germany, completed the utter collapse of Napoleon III's foreign design, demonstrating that while he possessed the imagination to disrupt the old world, he lacked the statesmanship to survive the forces he had unleashed.")
]

quotes_14 = [
    ("Henry Kissinger", "Diplomacy (1994)", "'Napoleon III was a tragic figure. He destroyed the Metternichian order that had preserved the peace of Europe and contained Prussia, only to find himself the author of a unified Germany that reduced France to a second-rate continental power.'"),
    ("A.J.P. Taylor", "The Struggle for Mastery in Europe", "'Napoleon III was a conspirator of genius, but an incompetent statesman... He lived on the prestige of a great name and the illusion of imperial grandeur, mistaking vague benevolent impulses for a foreign policy.'"),
    ("David Thomson", "Europe Since Napoleon", "'Napoleon III's tragedy was that he was both an idealist and an opportunist. His support for Italian nationality was genuine, yet he could never reconcile it with the defense of the Pope or the security of French borders.'"),
    ("Otto von Bismarck", "Contemporary Diplomatic Verdict", "'Napoleon III is a Sphinx without a riddle... from a distance, he looks like something significant; close up, he is a great unrecognized incapacity.'")
]

# Run the generation
if __name__ == "__main__":
    tasks = [
        (meta_13, callout_13, map_13, timeline_13, essay_13, quotes_13,
         os.path.join(paper1_dir, "Q13_Revolutions_of_1848_Spring_of_Nations_Master_Notes.docx"),
         "QUESTION 13: THE REVOLUTIONS OF 1848 ('THE SPRING OF NATIONS')",
         "Q13_Revolutions_of_1848_Spring_of_Nations_Master_Notes"),
        (meta_14, callout_14, map_14, timeline_14, essay_14, quotes_14,
         os.path.join(paper1_dir, "Q14_Napoleon_III_Foreign_Policy_Master_Notes.docx"),
         "QUESTION 14: FOREIGN POLICY OF NAPOLEON III (THE SECOND EMPIRE)",
         "Q14_Napoleon_III_Foreign_Policy_Master_Notes")
    ]
    
    for meta, callout, cmap, timeline, essay, quotes, docx_p, title, base_n in tasks:
        generate_master_doc(meta, callout, cmap, timeline, essay, quotes, docx_p, title)
        
        md_text = f"# KPPSC PMS \u2022 EUROPEAN HISTORY (PAPER I)\n## {title}\n\n"
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
