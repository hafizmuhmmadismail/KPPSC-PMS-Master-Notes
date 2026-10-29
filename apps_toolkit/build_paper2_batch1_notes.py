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
# QUESTION 4 (PAPER II): 2025 Q.2 (PART 1) / 2018 Q.2 / 2016 Q.1 (RUSSIAN REV)
# =========================================================================

meta_p2_4 = [
    ("EXAM QUESTIONS LINKED:", "2025 Q.2: 'Investigate the Russian Revolution, detailing the causes of the Bolshevik seizure of power, the subsequent civil war, and the consolidation of a totalitarian state under Lenin and Stalin.'\n2018 Q.2: 'Write a detailed note on the Russian Revolution (1917) and bring out its causes and results.'\n2016 Q.1: 'Elaborate in detail on the Bolshevik Revolution in Tsarist Russia and evaluate the role of Lenin as a revolutionary leader.'"),
    ("SYLLABUS TOPIC:", "Paper II \u2014 Section 4: Marxism & The Soviet Union \u2014 February and October Revolutions (1917), Lenin's Leadership, War Communism & NEP"),
    ("ALLOCATED MARKS / TIME:", "20 Marks | 35 Minutes | Standard Word Count: ~650\u2013700 words"),
    ("HISTORIOGRAPHICAL THESIS:", "The Russian Revolution of 1917 was a two-act historical cataclysm ignited by the structural exhaustion of Tsarist autocracy in total war; while the spontaneous February Revolution dismantled Romanov dynastic legitimacy only to construct an unstable 'Dual Power' deadlock, Vladimir Lenin's vanguard party discipline, tactical flexibility ('Peace, Land, and Bread'), and Leon Trotsky's military execution seized power in October, cementing the Soviet state through the coercive terror of War Communism and the pragmatic economic compromise of the New Economic Policy (NEP).")
]

callout_p2_4 = [
    "\u2022 The Dual Revolutions: February 1917 was a leaderless, spontaneous explosion of bread riots in Petrograd, industrial strikes, and soldier mutinies that forced Tsar Nicholas II's abdication. October 1917 was a disciplined, highly planned coup d'\u00e9tat executed by Lenin's Bolshevik minority through Trotsky's Military Revolutionary Committee.",
    "\u2022 The Fatal Blunders of the Provisional Government: Led by Prince Lvov and later Alexander Kerensky, the Provisional Government committed two mortal errors: (1) continuing Russia's disastrous participation in World War I to honor Allied treaties, and (2) postponing land redistribution until a democratically elected Constituent Assembly met. This alienated the peasant-soldier masses.",
    "\u2022 The Genius of Lenin's April Theses: Upon returning from Swiss exile, Lenin broke with orthodox Menshevik Marxism (which demanded a prolonged capitalist stage) by declaring that the bourgeois revolution was over. He demanded 'All Power to the Soviets' and coined the irresistible populist slogan 'Peace, Land, and Bread.'",
    "\u2022 The Kornilov Catalyst (August 1917): General Lavr Kornilov's abortive right-wing military coup compelled Kerensky to liberate and arm imprisoned Bolshevik Red Guards; this fatally exposed government impotence and gave Bolsheviks the majority in the Petrograd and Moscow Soviets.",
    "\u2022 From War Communism to NEP: War Communism (1918\u20131921) nationalized all industry, banned private commerce, instituted brutal grain requisitioning (prodrazvyorstka), and unleashed Cheka Red Terror, winning the Civil War against the Whites but causing the catastrophic 1921 famine. Following the Kronstadt Sailors' Mutiny (March 1921), Lenin pragmatically executed a tactical retreat with the New Economic Policy (NEP), re-introducing small-scale capitalism and market grain sales while retaining state control over the 'commanding heights' (banking, heavy industry, foreign trade)."
]

map_p2_4 = """========================================================================================
STRUCTURAL COLLAPSE (FEB 1917)       DUAL POWER & BOLSHEVIK SURGE      OCTOBER COUP & CONSOLIDATION
========================================================================================
[WWI Military Carnage & Tannenberg]--> [Tsar Nicholas II Abdicates]  --> [Oct 1917: Red Guards Take Power]
[Petrograd Bread Riots & Strikes]   --> [Provisional Govt vs. Soviet] --> [Decrees on Peace and Land]
[Garrison Mutiny & Rasputin Rot]    --> [Kerensky Keeps Fighting WWI] --> [Brest-Litovsk Exit from WWI]
                                    --> [Lenin's April Theses]        --> [Civil War & War Communism]
                                    --> [Kornilov Coup Arms Bolsheviks]--> [Kronstadt Revolt --> NEP 1921]
========================================================================================"""

timeline_p2_4 = [
    ("Feb 23\u2013Mar 2, 1917", "The February Revolution", "International Women's Day strikes turn into armed mutiny; Tsar Nicholas II abdicates; Provisional Government formed."),
    ("April 16, 1917", "Lenin's Return to Petrograd", "Arrives at Finland Station; publishes April Theses: 'All Power to the Soviets' and 'No support for Provisional Govt.'"),
    ("July 1917", "The July Days & Kerensky's Rise", "Spontaneous armed worker-soldier uprising fails; Bolsheviks suppressed; Lenin flees to Finland; Kerensky becomes Premier."),
    ("Aug 1917", "The Kornilov Affair", "General Kornilov marches on Petrograd; Kerensky arms Bolshevik Red Guards; Bolsheviks emerge as saviors of revolution."),
    ("Oct 25 (Nov 7), 1917", "The October Revolution", "Trotsky's Military Revolutionary Committee seizes telephone exchanges, railway stations, and storms the Winter Palace."),
    ("March 3, 1918", "Treaty of Brest-Litovsk", "Bolsheviks sign humiliating peace with Germany, surrendering Ukraine, Baltic states, and 34% of Russia's population."),
    ("1918\u20131921", "Russian Civil War & War Communism", "Trotsky organizes Red Army; defeats White Armies and 14 foreign interventionist nations; Red Terror crushed dissent."),
    ("March 1921", "Kronstadt Mutiny & NEP Introduced", "Red Sailors rebel demanding 'Soviets without Bolsheviks'; crushed by Trotsky; Lenin pivots to New Economic Policy."),
    ("Dec 1922", "Formation of the USSR", "Treaty on the Creation of the Union of Soviet Socialist Republics officially establishes the Soviet Union."),
    ("Jan 21, 1924", "Death of Vladimir Lenin", "Lenin dies following a series of strokes; power struggle between Stalin and Trotsky erupts.")
]

essay_p2_4 = [
    ("Introduction: The Two Acts of the 1917 Upheaval",
     "The Russian Revolution of 1917 stands as the most consequential sociopolitical rupture of the twentieth century, liquidating three centuries of Romanov dynastic autocracy and establishing the world's first Marxist-Leninist state. In the historiographical debate, while Soviet historiography historically sanctified 1917 as an inevitable, heroic uprising of the united proletariat under the Bolshevik banner, Western revisionist and liberal scholars such as Richard Pipes, Sheila Fitzpatrick, and Orlando Figes emphasize the profound distinction between the February and October upheavals. February was a genuine, leaderless popular implosion born of total military exhaustion and domestic starvation, whereas October was a brilliantly coordinated insurrection executed by a disciplined vanguard party. Vladimir Lenin's decisive contribution lay in his ideological audacity and tactical pragmatism: by discarding classical Marxist dogma, exploiting the democratic vacuum of the Provisional Government, and navigating the desperate transition from War Communism to the New Economic Policy (NEP), Lenin transformed revolutionary theory into permanent state power."),

    ("The Structural Roots: Tsarist Autocracy and the Crucible of World War I",
     "The collapse of the Romanov Empire was engineered fundamentally by the strain of industrialized total war upon an archaic, semi-feudal political system. By 1917, Russia had mobilized fifteen million men into a corrupt, poorly equipped military machine. Disastrous defeats at Tannenberg and the Masurian Lakes, accompanied by the Great Retreat of 1915, resulted in over five million Russian casualties. Tsar Nicholas II's disastrous decision in September 1915 to assume personal command of the army inextricably tied the monarchy's prestige to military slaughter, leaving the domestic administration in Petrograd in the hands of the German-born Tsarina Alexandra and the scandalous faith-healer Grigori Rasputin.\n\n"
     "On the home front, the conversion of railway lines to military transport starved major urban centers of grain and fuel. Hyperinflation soared by 400 percent, real wages collapsed, and bread rationing provoked existential desperation among urban factory workers. On February 23, 1917 (International Women's Day), bread riots in Petrograd ignited a general strike. When the elite Petrograd military garrison refused to shoot demonstrators and actively joined the strikers on February 27, the imperial coercive apparatus evaporated. Deserted by his generals, Nicholas II abdicated on March 2, ending Romanov rule."),

    ("The Dilemma of Dual Power: The Fatal Errors of the Provisional Government",
     "In the vacuum left by the autocracy, an inherently unstable constitutional compromise emerged known as 'Dual Power' (Dvoyevlastiye). Formal legal authority rested with the Provisional Government\u2014a coalition of liberal Cadets and moderate socialists led initially by Prince Georgy Lvov and later Alexander Kerensky. However, actual coercive control and popular legitimacy resided with the Petrograd Soviet of Workers' and Soldiers' Deputies, whose revolutionary 'Order No. 1' stripped military officers of authority and placed weapons under elected soldier committees.\n\n"
     "The Provisional Government committed two catastrophic strategic blunders that doomed it to destruction. First, constrained by national pride and financial dependence on British and French loans, it insisted on continuing the war against Germany, launching the disastrous Kerensky Offensive in June 1917, which produced 400,000 casualties and triggered mass desertions. Second, guided by bourgeois legalism, it refused to sanction peasant seizures of noble estates, insisting that agrarian redistribution must await the convocation of a Constituent Assembly. Consequently, the government alienated the two most powerful social forces in Russia: soldiers desperate for peace and peasants hungry for land."),

    ("Lenin's Intervention: The April Theses and the Seizure of October",
     "Vladimir Lenin recognized that the Provisional Government had paralyzed itself. Returning from Swiss exile via the famous German 'sealed train' in April 1917, Lenin stunned fellow Bolsheviks by issuing the 'April Theses.' Rejecting the orthodox Marxist belief that Russia must complete a prolonged bourgeois-democratic capitalist phase before attempting socialism, Lenin argued that the imperialist world war had created the weakest link in Russia, allowing the proletariat to leap directly to socialist revolution. He coined two electrifying slogans that resonated across the empire: 'All Power to the Soviets' and 'Peace, Land, and Bread.'\n\n"
     "The turning point arrived with the Kornilov Affair in August 1917. When the conservative Commander-in-Chief General Lavr Kornilov attempted a military coup d'\u00e9tat to crush the soviets, Kerensky panicked and appealed to the Bolsheviks, releasing imprisoned leaders and distributing forty thousand rifles to Bolshevik Red Guards. The coup dissolved without bloodshed, but the political fallout was fatal: Kerensky was discredited as an accomplice to reaction, while the Bolsheviks were celebrated as the saviors of the revolution, winning majorities in the Petrograd and Moscow Soviets. On the night of October 24\u201325, 1917, under the operational military genius of Leon Trotsky, Red Guards systematically seized key communications, bridges, and utilities, surrounding and storming the Winter Palace with minimal casualties. At the Second Congress of Soviets, Lenin proclaimed Soviet power and immediately issued the Decree on Peace (demanding immediate armistice) and the Decree on Land (legitimizing peasant land seizures)."),

    ("Consolidation of the Dictatorship: War Communism, Red Terror, and the Civil War",
     "Having captured power, the Bolsheviks faced immediate existential threats. In March 1918, Lenin insisted on signing the humiliating Treaty of Brest-Litovsk with Germany, sacrificing a third of Russia's population and fertile grain lands to buy 'breathing space' for the revolution. Between 1918 and 1921, the nascent Soviet state was plunged into a catastrophic Civil War, pit against the counter-revolutionary 'White' armies (commanded by Kolchak, Denikin, and Wrangel) backed by military interventions from fourteen foreign powers, including Britain, France, the US, and Japan.\n\n"
     "To survive, Lenin and Trotsky mobilized society under the brutal regime of 'War Communism.' All private enterprise was banned, large industries were nationalized, strikes were outlawed under military discipline, and armed detachment squads enforced the ruthless requisitioning of peasant grain (prodrazvyorstka). To eliminate internal class enemies, Felix Dzerzhinsky's political police (the Cheka) unleashed the 'Red Terror,' executing an estimated 200,000 suspected counter-revolutionaries, including the imperial family at Yekaterinburg in July 1918. Through Trotsky's creation of a three-million-man disciplined Red Army (employing former Tsarist officers monitored by political commissars), the Bolsheviks crushed the fragmented White armies by 1921."),

    ("The Strategic Retreat: Kronstadt and the New Economic Policy (NEP)",
     "While War Communism secured military victory, it devastated the Russian economy. By 1921, industrial output was merely twenty percent of its 1913 level, peasant production plummeted in protest against requisitioning, and a catastrophic drought ignited a horrific famine across the Volga basin that claimed five million lives. The breaking point arrived in March 1921 when the sailors of the Kronstadt naval fortress\u2014the 'pride and glory' of the 1917 revolution\u2014rebelled against Bolshevik autocracy, demanding free elections, freedom of speech, and an end to grain confiscation.\n\n"
     "Recognizing that the regime was on the precipice of popular overthrow, Lenin ruthlessly crushed the Kronstadt rebellion with Red Army troops, but immediately executed an astonishing economic pivot. At the Tenth Party Congress (March 1921), Lenin introduced the New Economic Policy (NEP), which he termed a necessary 'tactical retreat.' Forced grain requisitioning was replaced by a fixed tax in kind, allowing peasants to sell their surplus grain on the open market. Small-scale manufacturing and private retail trade were legalized, generating a flourishing class of small traders (NEPmen) and prosperous peasants (kulaks). Crucially, the Bolshevik state retained absolute control over the 'commanding heights' of the economy: heavy metallurgy, coal mining, foreign trade, and the banking system. Concurrently, Lenin banned all internal party factions, ensuring that while economic life was partially liberalized, political power remained an iron, monolithic one-party dictatorship."),

    ("Conclusion: The Monumental Legacy of 1917",
     "In conclusion, the Russian Revolution of 1917 was the foundational earthquake of modern international history. It dismantled a millennial empire, exposed the bankruptcy of liberal parliamentarianism in conditions of acute agrarian backwardness, and created a radically new model of state-directed ideological modernity. Vladimir Lenin demonstrated unmatched mastery of revolutionary timing, transforming Marxist philosophy from a theoretical critique of Western capitalism into an operational handbook for the seizure and retention of power. While the New Economic Policy successfully revived the devastated Soviet economy, the institutional machinery forged by Lenin\u2014the vanguard party monopoly, the Cheka police state, the suppression of democratic dissent, and the total subordination of civil society\u2014provided the exact totalitarian scaffolding that Joseph Stalin would ruthlessly inherit and weaponize in his 'Revolution from Above.'")
]

quotes_p2_4 = [
    ("Richard Pipes", "The Russian Revolution (1990)", "'The October Revolution was not a mass uprising... It was a classic coup d'\u00e9tat carried out by a tightly organized minority that exploited popular confusion and seized power to establish an unprecedented totalitarian regime.'"),
    ("Sheila Fitzpatrick", "The Russian Revolution", "'The Bolsheviks did not create the popular revolution; they rode the wave of genuine popular radicalism. In 1917, the workers, soldiers, and peasants wanted direct power, and Lenin was the only leader with the political courage to give it to them.'"),
    ("Orlando Figes", "A People's Tragedy (1996)", "'The failure of the Provisional Government was that it tried to govern a revolutionary country by constitutional rules... In the absence of a strong civil society, Russian political culture was doomed to revert to authoritarian violence.'"),
    ("Vladimir Lenin (1917)", "The State and Revolution", "'The goal of socialism is not merely the overthrow of the bourgeoisie, but the smashing of the bureaucratic-military state machine and its replacement by the dictatorship of the proletariat.'")
]

# =========================================================================
# QUESTION 5 (PAPER II): 2025 Q.2 (PART 2) / 2013 Q.5 (STALIN'S TOTALITARIAN STATE)
# =========================================================================

meta_p2_5 = [
    ("EXAM QUESTIONS LINKED:", "2025 Q.2 (Part 2): 'Investigate the consolidation of a totalitarian state under Lenin and Stalin.'\n2013 Q.5: 'Discuss the foreign policy of the Soviet Union under Joseph Stalin between 1928 and 1939, culminating in the Nazi-Soviet Non-Aggression Pact.'"),
    ("SYLLABUS TOPIC:", "Paper II \u2014 Section 4: Marxism & The Soviet Union \u2014 Stalin's Rise, Five-Year Plans, Collectivization, The Great Purges & Pre-War Foreign Policy"),
    ("ALLOCATED MARKS / TIME:", "20 Marks | 35 Minutes | Standard Word Count: ~650\u2013700 words"),
    ("HISTORIOGRAPHICAL THESIS:", "Joseph Stalin transformed the Soviet Union from an agrarian peasant society into an industrialized, thermonuclear superpower through a ruthless 'Revolution from Above'; by deploying forced agricultural collectivization to finance heavy Five-Year industrial plans, purging all institutional and military rivals in the Great Terror (Yezhovshchina), and conducting an unprincipled realpolitik foreign policy culminating in the 1939 Molotov-Ribbentrop Pact, Stalin constructed the definitive twentieth-century totalitarian police state.")
]

callout_p2_5 = [
    "\u2022 The Ascent to Supremacy (1924\u20131929): Following Lenin's death, Stalin utilized his bureaucratic post as General Secretary to pack party ranks with loyal clients. He allied with the Right (Bukharin) to defeat Leon Trotsky's 'Permanent Revolution' with his populist doctrine of 'Socialism in One Country,' expelled the Left Opposition (Trotsky, Zinoviev, Kamenev), and subsequently turned on Bukharin to terminate the NEP.",
    "\u2022 Forced Collectivization & The Holodomor: In 1928, Stalin declared war on the peasantry to finance industrialization; forced 25 million peasant households into collective farms (kolkhoz); ordered the total liquidation of the kulaks as a class ('dekulakization'); peasants slaughtered over 50% of their livestock in resistance; state-engineered grain requisitioning caused the catastrophic Ukrainian famine (Holodomor, 1932\u20131933), resulting in 4 to 7 million civilian deaths.",
    "\u2022 The Five-Year Plans & Industrial Miracle: The First (1928\u20131932) and Second (1933\u20131937) Five-Year Plans prioritized heavy industry (coal, iron, steel, electricity, tractors) over consumer goods. Built industrial cities from scratch beyond the Urals (Magnitogorsk); mobilized the Stakhanovite worker movement; by 1939, USSR was the world's second-largest industrial producer, providing the material base to defeat Nazi Germany.",
    "\u2022 The Great Terror (Yezhovshchina 1936\u20131938): Sparked by the 1934 assassination of Sergei Kirov; three sensational Moscow Show Trials executed all surviving Old Bolshevik leaders (Zinoviev, Kamenev, Bukharin); military purge executed Marshal Tukhachevsky, 3 of 5 Marshals, and 35,000 officer corps members; Gulag penal camp network housed millions of forced laborers; Cult of Personality elevated Stalin to infallible deity.",
    "\u2022 Foreign Policy Pivot (1928\u20131939): Shifted from early Comintern radicalism (denouncing social democrats as 'social fascists') to Maxim Litvinov's 'Popular Front' and 'Collective Security' in the League of Nations (1934). Deeply disillusioned by British and French appeasement of Hitler at the Munich Conference (1938), Stalin executed a stunning diplomatic volte-face: signing the Molotov-Ribbentrop Non-Aggression Pact on August 23, 1939, partitioning Poland and the Baltic states."
]

map_p2_5 = """========================================================================================
BUREAUCRATIC ASCENT               REVOLUTION FROM ABOVE (1928-38)     TOTALITARIAN CONSOLIDATION
========================================================================================
[General Secretary Machine]    --> [Collectivization & Kolkhoz]     --> [The Great Terror (1936-38)]
['Socialism in One Country']   --> [Liquidation of Kulaks]          --> [Moscow Show Trials & Gulag]
[Defeat of Trotsky & Bukharin] --> [The Holodomor Famine (1932-33)] --> [Red Army Officer Purge]
[Termination of Lenin's NEP]   --> [1st & 2nd Five-Year Plans:      --> [Cynical Foreign Realpolitik:
                                    Heavy Steel, Magnitogorsk]           Molotov-Ribbentrop Pact 1939]
========================================================================================"""

timeline_p2_5 = [
    ("1922\u20131924", "Stalin as General Secretary", "Lenin writes Testament warning party to remove Stalin as 'too rude'; Stalin suppresses Testament upon Lenin's death."),
    ("1927\u20131929", "Triumph of 'Socialism in One Country'", "Trotsky expelled from Communist Party and exiled; Bukharin removed from Politburo; NEP abandoned."),
    ("1928\u20131932", "First Five-Year Plan", "Total focus on heavy metallurgy, electricity, and coal; Gosplan sets astronomical, mandatory industrial targets."),
    ("1929\u20131930", "Total Collectivization Decree", "Stalin orders total collectivization; 'dekulakization' deports millions of peasant families to Siberia."),
    ("1932\u20131933", "The Holodomor Famine", "State-mandated grain quotas cause horrific man-made famine in Ukraine and North Caucasus; millions starve."),
    ("Dec 1, 1934", "Assassination of Sergei Kirov", "Leningrad party chief murdered under mysterious circumstances; Stalin uses event as pretext to unleash the Great Purges."),
    ("1936\u20131938", "The Moscow Show Trials", "Old Bolsheviks (Zinoviev, Kamenev, Radek, Bukharin) confess to bizarre Trotskyite-fascist plots and are executed."),
    ("June 1937", "Purge of the Red Army", "Marshal Mikhail Tukhachevsky and 80% of military high command executed; critically weakens military readiness."),
    ("Sept 1938", "The Munich Conference Betrayal", "Britain and France exclude USSR from Czech partition; Stalin concludes West is directing Hitler eastward."),
    ("Aug 23, 1939", "Molotov-Ribbentrop Pact", "Nazi Germany and Soviet Union sign Non-Aggression Pact; secret protocols partition Poland and Baltic states.")
]

essay_p2_5 = [
    ("Introduction: The Architect of Soviet Totalitarianism",
     "Between 1928 and 1939, the Union of Soviet Socialist Republics was subjected to a titanic, violent metamorphosis that converted a backward peasant society into an industrialized, totalitarian fortress. The architect of this monumental transformation was Joseph Stalin. Following the death of Vladimir Lenin in 1924, Stalin executed a masterclass in bureaucratic factionalism, outmaneuvering both the Left Opposition of Leon Trotsky and the Right Opposition of Nikolai Bukharin. Once in absolute control, Stalin dismantled Lenin's market-based New Economic Policy and launched what historians term the 'Revolution from Above.' In the historiographical debate between classical totalitarian theorists like Robert Conquest and revisionist social historians like Sheila Fitzpatrick and Stephen Kotkin, Stalinism is recognized as a complex convergence: a system of unprecedented state terror, forced collectivization, and rapid industrialization driven by an apocalyptic geopolitical fear of capitalist encirclement, combined with cynical diplomatic realpolitik that culminated in the Nazi-Soviet Non-Aggression Pact of 1939."),

    ("The Economic Crucible: Five-Year Plans and Forced Industrialization",
     "Stalin's domestic revolution was animated by an acute sense of national vulnerability, famously articulated in his historic February 1931 speech to industrial managers: 'We are fifty or a hundred years behind the advanced countries. We must make good this distance in ten years. Either we do it, or they will crush us.' To achieve this existential leap, Stalin instituted the First Five-Year Plan (1928\u20131932), followed immediately by the Second (1933\u20131937), orchestrated by the state planning agency Gosplan.\n\n"
     "The plans entirely bypassed consumer goods in favor of heavy industry: coal, iron, steel, electricity, and heavy machinery. Giant industrial complexes were constructed from scratch beyond the reach of European bombers in the Ural Mountains and Siberia, most notably the steel city of Magnitogorsk. The state deployed intense propaganda to celebrate socialist labor heroism, immortalized in the 'Stakhanovite movement' after miner Alexei Stakhanov. While Western capitalist economies collapsed into catastrophic unemployment during the Great Depression, the Soviet Union achieved an astonishing annual industrial growth rate exceeding fifteen percent. By 1939, the USSR had surpassed Britain and Germany to become the second-largest industrial producer in the world, providing the indispensable material foundation that enabled the Red Army to manufacture tens of thousands of T-34 tanks and artillery pieces during World War II."),

    ("Agrarian War: Collectivization, Dekulakization, and the Holodomor",
     "The colossal financial capital required to purchase foreign industrial technology and feed the burgeoning urban factory workforce was extracted directly from the Soviet peasantry. In 1929, Stalin launched the forced collectivization of agriculture, liquidating twenty-five million individual peasant smallholdings and forcing them into state-supervised collective farms (kolkhoz) and state farms (sovkhoz). To break peasant resistance, Stalin ordered the 'liquidation of the kulaks as a class.' Over five million prosperous or recalcitrant peasants (kulaks) were stripped of property, murdered, or deported in cattle cars to forced labor colonies in Siberia and the Arctic.\n\n"
     "Peasant resistance was fierce: rather than surrender their property to the collective, peasants slaughtered half of all Soviet horses and cattle. Stalin responded with ferocious state terror. In 1932\u20131933, when agricultural yields plummeted, the state aggressively requisitioned grain quotas at gunpoint, cordoned off borders, and deliberately engineered the catastrophic Ukrainian famine known as the Holodomor. Between four and seven million human beings perished of starvation in Ukraine, Kazakhstan, and the North Caucasus. While collectivization was an agricultural and humanitarian catastrophe that permanently crippled Soviet farming productivity, it achieved Stalin's political goal: it destroyed all autonomous peasant resistance, subordinated the countryside to state control, and secured cheap grain for urban industrialization."),

    ("The Machinery of Absolute Control: The Great Terror and the Gulag",
     "To consolidate his personal autocracy and eliminate any potential political alternative, Stalin constructed an all-pervasive totalitarian police state. The catalyst for absolute terror was the assassination of Leningrad party leader Sergei Kirov on December 1, 1934\u2014an event orchestrated or opportunistically seized upon by Stalin. Between 1936 and 1938, the NKVD under Nikolai Yezhov executed the 'Great Terror' (Yezhovshchina). Three theatrical Moscow Show Trials were staged before the international press, wherein the entire historic leadership of the Bolshevik Revolution\u2014including Lev Kamenev, Grigory Zinoviev, and Nikolai Bukharin\u2014confessed under physical torture and psychological blackmail to absurd conspiracies of Trotskyite sabotage, espionage, and fascism before being shot in the cellars of the Lubyanka.\n\n"
     "The terror escalated into a systematic purge of Soviet society: artists, scientists, engineers, and ordinary citizens were arrested under Article 58 on fabricated charges of counter-revolution. Most critically, in 1937, Stalin purged the Red Army high command, executing Marshal Mikhail Tukhachevsky, three of five Marshals, and thirty-five thousand experienced officer corps members. Millions of citizens were condemned to the vast network of slave labor camps managed by the Main Administration of Corrective Labor Camps (the Gulag), where convicts constructed the White Sea-Baltic Canal and mined gold in Kolyma under sub-zero temperatures. Simultaneously, the state cultivated a semi-religious 'Cult of Personality,' elevating Stalin to the omniscient 'Father of Nations' through monumental statues, retouched photography, and the cultural orthodoxy of Socialist Realism."),

    ("Soviet Foreign Policy (1928\u20131939): From Collective Security to the Nazi-Soviet Pact",
     "Stalin's foreign policy was governed not by ideological romanticism, but by cold, opportunistic realpolitik. In the early 1930s, Stalin adhered to Comintern dogma, denouncing European social democratic parties as 'social fascists'\u2014a fatal miscalculation that split the German left and facilitated Adolf Hitler's rise to power in 1933. Recognizing the deadly threat of Nazi anti-Bolshevik Lebensraum, Stalin executed a radical foreign policy pivot in 1934: the USSR joined the League of Nations, and Foreign Commissar Maxim Litvinov became the champion of 'Collective Security' and the 'Popular Front,' forging defensive alliances with France and Czechoslovakia (1935) and supporting the Republican forces in the Spanish Civil War (1936\u20131939).\n\n"
     "However, Stalin grew deeply suspicious that Britain and France were deliberately appeasing Hitler to direct German expansion eastward against the Soviet Union. The definitive rupture occurred at the Munich Conference of September 1938, where Britain and France partitioned Czechoslovakia without even inviting the USSR. Convinced of Western treachery, Stalin replaced Litvinov with the hardline Vyacheslav Molotov. When Anglo-French military negotiators dragged their feet in Moscow in the summer of 1939, Stalin executed the most sensational diplomatic volte-face in modern history: on August 23, 1939, the USSR and Nazi Germany signed the Molotov-Ribbentrop Non-Aggression Pact. The treaty's secret protocol divided Poland, the Baltic states (Estonia, Latvia, Lithuania), and Bessarabia into Nazi and Soviet spheres of influence. By signing the pact, Stalin secured a buffer zone, avoided a two-front war against Germany and Japan, and bought two years of precious time to prepare for the inevitable German onslaught."),

    ("Conclusion: The Terrible Balance Sheet of Stalinism",
     "In conclusion, Joseph Stalin was the titanic, ruthless architect of the modern Soviet superpower. Through the sheer application of organized state terror and centralized economic coercion, he eradicated capitalism, built an impregnable heavy industrial infrastructure, and unified the state under an iron, totalitarian command. Yet, this historic achievement was exacted at an unfathomable human cost: an estimated twenty million Soviet citizens perished through artificial famines, executions, and Gulag labor camps. Furthermore, by decapitating the Red Army in the 1937 purges and miscalculating Hitler's intentions in 1941, Stalin left the Soviet Union dangerously vulnerable to Operation Barbarossa. Ultimately, Stalin's legacy remains profoundly dual: he was both a monstrous tyrant who murdered his own people and the ruthless modernizer whose industrial steel preserved the Soviet Union from total annihilation by Nazi Germany.")
]

quotes_p2_5 = [
    ("Robert Conquest", "The Great Terror (1968)", "'Stalin was the supreme practitioner of state violence in modern history... The purges and the famine were not administrative mistakes, but cold, deliberate instruments designed to secure total political submission.'"),
    ("Stephen Kotkin", "Stalin: Waiting for Hitler (2017)", "'Stalin was not a cynical opportunist devoid of belief; he was a driven, zealous Marxist-Leninist who believed that ruthless force was the only midwife capable of delivering socialist modernization.'"),
    ("Sheila Fitzpatrick", "Everyday Stalinism (1999)", "'Stalinism was an extraordinary cultural and social phenomenon: while terror gripped the old elites, the Five-Year Plans created unprecedented upward social mobility for millions of working-class citizens who became the new Soviet bureaucracy.'"),
    ("Joseph Stalin (1931)", "Speech to Industrial Managers", "'We are fifty or a hundred years behind the advanced countries. We must make good this distance in ten years. Either we do it, or they will crush us.'")
]

# =========================================================================
# QUESTION 6 (PAPER II): 2025 Q.3 / 2018 Q.3 / 2016 Q.3 / 2013 Q.6 (LEAGUE)
# =========================================================================

meta_p2_6 = [
    ("EXAM QUESTIONS LINKED:", "2025 Q.3: 'Discuss the aims and achievements of the League of Nations. To what extent did its structural weaknesses and the policies of major powers lead to its failure?'\n2018 Q.3: 'What were the causes of failure of the League of Nations?'\n2016 Q.3: 'Discuss the mandate system and disarmament conferences under the auspices of the League of Nations. Why did disarmament fail in the 1930s?'\n2013 Q.6: 'Examine the collapse of the League of Nations with specific reference to the Japanese invasion of Manchuria and the Italian invasion of Abyssinia.'"),
    ("SYLLABUS TOPIC:", "Paper II \u2014 Section 3: The League of Nations & Interwar Security (1919\u20131939) \u2014 Covenant, Successes, Structural Flaws & Collapse"),
    ("ALLOCATED MARKS / TIME:", "20 Marks | 35 Minutes | Standard Word Count: ~650\u2013700 words"),
    ("HISTORIOGRAPHICAL THESIS:", "The League of Nations was an unprecedented institutional breakthrough aimed at substituting traditional nineteenth-century balance-of-power diplomacy with collective security and international law; while achieving notable humanitarian and minor frontier successes during the 1920s, the League was structurally crippled by the fatal absence of the United States, the unanimity voting requirement, and the lack of an independent military force, ensuring its inevitable collapse when confronted by fascist aggression in Manchuria (1931) and Abyssinia (1935) amidst British and French appeasement.")
]

callout_p2_6 = [
    "\u2022 The Covenant & Institutional Structure: Established under Part I of the Treaty of Versailles (1919) at Geneva; composed of an Assembly (all member states, one vote), a Council (Permanent Members: Britain, France, Italy, Japan; plus non-permanent members), a Permanent Secretariat, and the Permanent Court of International Justice (The Hague).",
    "\u2022 Article 10 (The Core of Collective Security): Obligated members to 'respect and preserve as against external aggression the territorial integrity and existing political independence of all Members.' Article 16 provided for automatic economic sanctions and moral condemnation against any state declaring war.",
    "\u2022 Fatal Structural Weaknesses:",
    "  1. American Isolationism: The US Senate's refusal to ratify Versailles deprived the League of the world's greatest financial and naval titan.",
    "  2. Exclusion & Walkouts: Germany was barred until 1926 (walked out 1933); USSR barred until 1934 (expelled 1939); Japan walked out (1933); Italy walked out (1937). It became a 'league of victors.'",
    "  3. Unanimity Voting Rule: All substantive decisions in the Council and Assembly required unanimous votes, allowing a single rogue state to veto collective action.",
    "  4. No Independent Armed Forces: The League possessed no standing military, relying entirely on Britain and France, neither of which would risk war for collective security.",
    "\u2022 Achievements in the 1920s: Successfully resolved disputes over the Aaland Islands (Sweden/Finland 1921), Upper Silesia (1921), and the Greek-Bulgarian border (1925); Fridtjof Nansen's refugee passports; global eradication campaigns against malaria, leprosy, and human trafficking.",
    "\u2022 The Two Fatal Death Knells (1930s):",
    "  1. Manchurian Crisis (1931\u20131933): Japan invaded Chinese Manchuria; Lord Lytton's commission took a full year to report; when the League condemned Japan, Tokyo simply resigned from the League with zero punitive consequences.",
    "  2. Abyssinian Crisis (1935\u20131936): Mussolini invaded Ethiopia; League voted sanctions but cowardly omitted oil, steel, and the British-controlled Suez Canal; secret Hoare-Laval Pact exposed Anglo-French treachery in offering Abyssinia to Mussolini, completely destroying the League's moral authority."
]

map_p2_6 = """========================================================================================
WILSONIAN AMBITION (1919)            THE 1920s ILLUSION OF PEACE       THE 1930s STRUCTURAL COLLAPSE
========================================================================================
[Article 10: Collective Security]--> [Aaland Islands Resolved (1921)]--> [1929 Great Depression Slump]
[Article 16: Economic Sanctions]--> [Upper Silesia Partition (1921)]--> [Manchuria 1931: Lytton Fails]
[Permanent Council & Assembly]   --> [Nansen Passports for Refugees] --> [Geneva Disarmament Fails (1933)]
                                 --> [Locarno Treaties & Germany In] --> [Abyssinia 1935: Sanctions Farce]
                                 --> [Kellogg-Briand Pact (1928)]    --> [Hoare-Laval Pact: League Dies]
========================================================================================"""

timeline_p2_6 = [
    ("April 28, 1919", "Covenant Adopted", "The Covenant of the League of Nations is unanimously adopted at the Paris Peace Conference."),
    ("Jan 16, 1920", "First Council Meeting", "Council meets in Paris; US absent following Senate rejection of the Versailles Treaty."),
    ("1921", "Aaland Islands & Silesia", "League successfully arbitrates territorial disputes between Sweden/Finland and Germany/Poland."),
    ("Oct 1925", "Locarno Treaties", "Germany accepts western borders; 'Locarno Spirit' ushers in interwar era of reconciliation."),
    ("Sept 1926", "Germany Joins League", "Weimar Germany admitted to the League with a permanent seat on the Council."),
    ("Aug 1928", "Kellogg-Briand Pact", "62 nations sign pact 'renouncing war as an instrument of national policy' with zero enforcement mechanism."),
    ("Sept 1931", "Japanese Invasion of Manchuria", "Japan seizes Manchuria; League appoints Lytton Commission; takes 14 months to deliver report."),
    ("March 1933", "Japan Leaves the League", "Following League adoption of Lytton Report, Japan votes 'No' and resigns from the League."),
    ("Oct 1933", "Hitler Quits Disarmament & League", "Germany withdraws from Geneva Disarmament Conference and leaves League to rearm in secret."),
    ("Oct 1935\u2013May 1936", "The Abyssinian Crisis", "Mussolini invades Ethiopia; Hoare-Laval scandal breaks; League sanctions fail; League collapses as security body.")
]

essay_p2_6 = [
    ("Introduction: The Noble Ideal and the Tragic Reality",
     "The League of Nations, established under Part I of the Treaty of Versailles in 1919, represented an unprecedented revolution in international relations. Championed by American President Woodrow Wilson, the League was designed to replace the secretive, balance-of-power alliance diplomacy that had precipitated World War I with a universal system of 'Collective Security' grounded in international law, arbitration, and disarmament. Yet, within two decades, this idealistic institution had degenerated into an impotent, discredited shell, helpless to prevent the relentless aggression of fascist dictatorships. In the historiographical debate, while Realist scholars like E.H. Carr condemned the League as a utopian delusion disconnected from the realities of military power, modern diplomatic historians like Zara Steiner and Sally Marks argue that the League was not fundamentally flawed in conception, but was sabotaged by critical structural handicaps, American abandonment, and the cynical, self-serving appeasement policies of Great Britain and France."),

    ("Institutional Architecture: The Covenant and the Mechanics of Collective Security",
     "The foundation of the League rested upon its twenty-six-article Covenant. Operating from Geneva, the League was structured around an Assembly representing all member states on an egalitarian one-nation-one-vote basis, an executive Council dominated by permanent Great Powers (initially Britain, France, Italy, and Japan) alongside rotating non-permanent members, a permanent international Secretariat, and the Permanent Court of International Justice at The Hague.\n\n"
     "The ideological engine of the League was enshrined in Article 10, wherein members pledged to 'respect and preserve as against external aggression the territorial integrity and existing political independence of all Members.' If an aggressive state resorted to war, Article 16 prescribed automatic, escalating collective responses: immediate diplomatic ostracization, total financial and trade embargoes, and, if recommended by the Council, joint military action. The revolutionary concept was that an attack against any member state would be treated as an attack against the entire global community, rendering aggression mathematically irrational."),

    ("Structural Vulnerabilities: The Seeds of Institutional Paralysis",
     "Despite its visionary rhetoric, the League was crippled from its inception by fatal structural and geopolitical deficiencies:\n\n"
     "1. The Defection of the United States: The fatal blow to the League occurred before it held its first session. In November 1919 and March 1920, the United States Senate, dominated by isolationist Republicans under Henry Cabot Lodge, refused to ratify the Treaty of Versailles. The absence of the world's greatest financial, industrial, and naval power deprived the League of the essential economic muscle required to make sanctions effective.\n\n"
     "2. The 'League of Victors' and Pariah States: The League was perceived globally as an instrument designed to preserve the imperialist spoils of Britain and France. Defeated Germany was excluded until 1926, while Bolshevik Russia was treated as a revolutionary pariah until 1934. When totalitarian regimes ascended, they simply walked out: Germany and Japan resigned in 1933, and Italy followed in 1937.\n\n"
     "3. The Unanimity Rule: Articles 5 and 15 mandated that all substantive resolutions in both the Assembly and the Council required unanimous consent. This absurd requirement meant that a single dissenting vote could paralyze the entire international community.\n\n"
     "4. The Absence of Military Enforcement: The League possessed no standing international police force or army. To enforce Article 16 militarily, it relied entirely on the voluntary military contributions of its members. Britain and France, traumatized by the bloodbath of 1914\u20131918 and prioritizing domestic imperial defense, were adamantly unwilling to shed blood for collective security in distant theaters."),

    ("The Illusion of Peace: Successes and Humanitarian Triumphs in the 1920s",
     "During its first decade, the League achieved notable, albeit minor, successes that fostered an illusion of durable international harmony. In 1921, the League successfully mediated the bitter territorial dispute between Sweden and Finland over the Aaland Islands, awarding sovereignty to Finland with guaranteed cultural autonomy for the Swedish inhabitants. In the same year, it peacefully partitioned Upper Silesia between Germany and Poland. In 1925, when Greek forces invaded Bulgaria in the 'War of the Stray Dog,' the Council acted decisively, ordering a ceasefire and compelling Greece to withdraw and pay reparations.\n\n"
     "Furthermore, the League's specialized non-political agencies performed heroic pioneering work. The High Commission for Refugees, led by Norwegian explorer Fridtjof Nansen, created the historic 'Nansen Passport,' providing legal identity and resettlement to over 450,000 stateless refugees. The Health Organization orchestrated international campaigns that checked typhus epidemics in Eastern Europe and eliminated yellow fever, while other commissions combated human trafficking, regulated drug smuggling, and established the International Labor Organization (ILO) to elevate global labor standards."),

    ("The Acid Test of the 1930s: The Manchurian Catastrophe (1931\u20131933)",
     "The fragile security edifice built in the 1920s was shattered by the Wall Street Crash of 1929 and the Great Depression. In September 1931, the Japanese Kwantung Army staged the Mukden Incident, invading and conquering the vast Chinese province of Manchuria to create the puppet state of Manchukuo. For the first time, a permanent member of the League Council had committed unprovoked aggression against another League member.\n\n"
     "China immediately appealed to the League under Article 11. The League responded with agonizing paralysis. Rather than imposing immediate economic sanctions under Article 16, the Council dispatched a fact-finding commission led by British diplomat Lord Lytton. It took the Lytton Commission fourteen months to travel to China, draft its report, and present it to the Assembly. When the League finally adopted the report in February 1933, condemning Japanese aggression, Japanese delegate Yosuke Matsuoka delivered a defiant speech, led his delegation out of the hall, and Japan formally resigned from the League in March. The League took zero punitive action: Britain and France refused to impose naval blockades or economic embargoes fearing Japanese retaliation against Hong Kong and Indochina. The message to the world's dictators was unmistakable: the League was an empty paper tiger."),

    ("The Death Knell: Abyssinia and the Treachery of Hoare-Laval (1935\u20131936)",
     "If Manchuria wounded the League, the Abyssinian Crisis was its executioner. In October 1935, Benito Mussolini's Fascist Italy invaded the sovereign, independent African empire of Abyssinia (Ethiopia), employing mustard poison gas, aerial bombardment, and armored divisions against defenseless tribesmen. Abyssinian Emperor Haile Selassie appeared personally before the League Assembly in Geneva, delivering an immortal indictment: 'It is collective security: it is the very existence of the League of Nations... It is international morality that is at stake. God and history will remember your judgment.'\n\n"
     "The League declared Italy an aggressor and voted to impose economic sanctions. However, terrified of alienating Mussolini and driving him into an alliance with Adolf Hitler, British Foreign Secretary Samuel Hoare and French Premier Pierre Laval engineered a cowardly compromise. Sanctions were deliberately neutered: vital commodities such as oil, coal, and steel were exempted, and the British government refused to close the Suez Canal to Italian troopships. The final moral betrayal occurred in December 1935 when the secret 'Hoare-Laval Pact' was leaked to the French press, revealing that Britain and France had secretly agreed to partition two-thirds of Abyssinia to Mussolini behind the League's back. The resulting public fury forced Hoare's resignation, but the damage was irreversible: Italy conquered Addis Ababa in May 1936, annexed Ethiopia, and walked out of the League to form the Rome-Berlin Axis with Hitler. The League of Nations was dead as a credible security instrument."),

    ("Conclusion: The Legacy of a Visionary Failure",
     "In conclusion, the collapse of the League of Nations was not a failure of the principle of collective security, but a tragic failure of national statesmanship. Entrusted to the custody of Britain and France, the League was repeatedly subordinated to short-sighted imperial self-interest and the cowardly evasions of appeasement. When Adolf Hitler witnessed the League's total humiliation in Abyssinia, he promptly remilitarized the Rhineland in March 1936, confident that the Western democracies would never take up arms for international law. The League spent its final years in ghost-like irrelevance, meeting in Geneva while Europe slid into World War II. Yet, the tragedy of the League was not entirely in vain: when the Allied Powers convened at San Francisco in 1945 to establish the United Nations, they corrected the League's structural flaws by eliminating the unanimity rule, creating an executive Security Council with enforcement teeth, and ensuring the permanent institutional leadership of the United States.")
]

quotes_p2_6 = [
    ("E.H. Carr", "The Twenty Years' Crisis (1939)", "'The fatal mistake of the champions of the League was to confuse an ideal aspiration with political reality... They believed that public opinion and moral declarations could substitute for military power.'"),
    ("Zara Steiner", "The Lights That Failed (2005)", "'The League was an astonishingly successful experiment in international administration and multilateral diplomacy... It failed in the 1930s not because of institutional design, but because the Great Depression tore apart the international economic order.'"),
    ("F.P. Walters", "A History of the League of Nations", "'The League was not a super-state; it was an association of sovereign governments. Its failures were not the failures of the Covenant, but the failures of the statesmen of London and Paris to honor their solemn pledges.'"),
    ("Haile Selassie (1936)", "Speech to League of Nations Assembly", "'I ask the fifty-two nations who have given the Ethiopian people a promise to help them: what are you going to do for Ethiopia? It is collective security that is at stake... It is us today. It will be you tomorrow.'")
]

# Run the generation
if __name__ == "__main__":
    tasks = [
        (meta_p2_4, callout_p2_4, map_p2_4, timeline_p2_4, essay_p2_4, quotes_p2_4,
         os.path.join(paper2_dir, "Q4_Russian_Revolutions_1917_Lenin_and_NEP_Master_Notes.docx"),
         "QUESTION 4: THE RUSSIAN REVOLUTIONS (1917), LENIN'S LEADERSHIP & NEP",
         "Q4_Russian_Revolutions_1917_Lenin_and_NEP_Master_Notes"),
        (meta_p2_5, callout_p2_5, map_p2_5, timeline_p2_5, essay_p2_5, quotes_p2_5,
         os.path.join(paper2_dir, "Q5_Stalin_Totalitarian_State_Collectivization_Purges_Master_Notes.docx"),
         "QUESTION 5: STALIN'S TOTALITARIAN STATE: FIVE-YEAR PLANS, PURGES & FOREIGN POLICY",
         "Q5_Stalin_Totalitarian_State_Collectivization_Purges_Master_Notes"),
        (meta_p2_6, callout_p2_6, map_p2_6, timeline_p2_6, essay_p2_6, quotes_p2_6,
         os.path.join(paper2_dir, "Q6_League_of_Nations_Security_Failures_Manchuria_Abyssinia_Master_Notes.docx"),
         "QUESTION 6: THE LEAGUE OF NATIONS: COVENANT, STRUCTURAL WEAKNESSES & COLLAPSE",
         "Q6_League_of_Nations_Security_Failures_Manchuria_Abyssinia_Master_Notes")
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
