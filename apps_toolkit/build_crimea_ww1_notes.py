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
    run_t = p.add_run(f"📌 {title}\n")
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
    r_main = p_title.add_run("KPPSC PMS &bull; EUROPEAN HISTORY (PAPER I)")
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
    r_f = p_foot.add_run("Khyber Pakhtunkhwa Public Service Commission (KPPSC) &bull; PMS Competitive Examination Notes &bull; European History")
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
# QUESTION 8: 2022 Q.5 / 2016 Q.6 / 2018 Q.7 (CRIMEAN WAR & EASTERN QUESTION)
# =========================================================================

meta_8 = [
    ("EXAM QUESTIONS LINKED:", "2022 Q.5: 'Highlight the causes and consequences of the Crimean War (1853–1856). Do you agree that it was a futile and unnecessary carnage?'\n2016 Q.6: 'Examine the causes, course, and historical significance of the Crimean War (1853–1856) in European diplomacy.'\n2018 Q.7: 'Trace the development of the Eastern Question from Greek Independence to the Treaty of Berlin.'"),
    ("TOTAL MARKS & TIME:", "20 Marks &bull; Descriptive Subjective Essay &bull; Time: 35 Minutes"),
    ("SYLLABUS CORRELATION:", "Paper I &bull; Section 5: The Eastern Question & Crimean War (1820–1878)"),
    ("HISTORIOGRAPHICAL THESIS:", "While militarily a blundering 'futile carnage' triggered by trivial religious disputes, diplomatically it was a tectonic earthquake that destroyed the 1815 Holy Alliance, isolated Austria, and unlocked Italian and German unification.")
]

callout_8 = [
    "1. The Core Paradox: Militarily, the Crimean War was an infamous display of administrative incompetence ('Charge of the Light Brigade', disease killing 4 out of 5 casualties); Diplomatically, it completely shattered the 1815 Holy Alliance between Austria and Russia.",
    "2. The Eastern Question Dilemma: The decline of the Ottoman Empire ('Sick Man of Europe') tempted Russian expansion towards Constantinople and the Mediterranean, threatening British overland and naval routes to India and French Catholic prestige.",
    "3. Geopolitical Earthquake of Treaty of Paris (1856): Neutralization of the Black Sea humiliated Russia; Austrian betrayal of Russia created eternal enmity between Vienna and St. Petersburg; Piedmont-Sardinia gained Great Power patronage (Cavour)."
]

map_8 = (
    "THE EASTERN QUESTION: Ottoman Decline ('Sick Man of Europe') ➔ Russian Ambition vs. British/French Fear\n"
    "  │\n"
    "  ├── PRETEXT: Holy Places in Palestine (Monks' Keys dispute: Roman Catholics vs. Orthodox Greeks)\n"
    "  │     └── Menshikov Ultimatum (1853): Tsar demands protectorate over 12 million Ottoman Christians\n"
    "  │\n"
    "  ├── ESCALATION: Battle of Sinope (Nov 1853) ➔ Russian fleet wipes out Turkish squadron\n"
    "  │     └── Anglo-French War Entry (March 1854): Protect Constantinople & preserve balance of power\n"
    "  │\n"
    "  ├── THE SEBASTOPOL CARNAGE (1854–1855): Balaclava, Inkerman, Malakoff ➔ Cholera, frostbite, tactical blunders\n"
    "  │     └── Austrian Ultimatum to Russia: Habsburg ingratitude permanently alienates the Tsar\n"
    "  │\n"
    "  └── TREATY OF PARIS (1856) & SHATTERED EUROPE:\n"
    "        ├── Black Sea neutralized; Ottoman integrity guaranteed\n"
    "        └── Holy Alliance Dead ➔ Isolated Austria + Embittered Russia = Door open for Cavour & Bismarck"
)

timeline_8 = [
    ("1844", "Tsar Nicholas I Visits England", "Proposes partition of Ottoman Empire ('We have on our hands a sick man, a very sick man')."),
    ("1850–1852", "Holy Places Dispute in Jerusalem", "Louis Napoleon (Napoleon III) champions Catholic monks; Tsar Nicholas I champions Orthodox monks."),
    ("May 1853", "Prince Menshikov's Ultimatum", "Russia demands sovereign protectorate over all Orthodox Christian subjects in Ottoman realms; Sultan refuses."),
    ("Nov 30, 1853", "Battle of Sinope", "Russian Black Sea fleet annihilates Turkish wooden flotilla; British press decries 'massacre at Sinope'."),
    ("March 1854", "Britain and France Declare War", "Anglo-French coalition enters the war to defend Ottoman integrity and check Russian naval expansion."),
    ("Oct 1854", "Battle of Balaclava", "Lord Cardigan leads the disastrous 'Charge of the Light Brigade'; immortalized by Tennyson."),
    ("Sept 1855", "Fall of Sevastopol", "French storm Malakoff redoubt; Russians evacuate naval citadel after 11-month brutal siege."),
    ("March 1856", "The Treaty of Paris", "Black Sea demilitarized; Danube opened to international commerce; Concert of Europe permanently fractured.")
]

essay_8 = [
    ("Introduction",
     "Between 1853 and 1856, the Great Powers of Europe fought their first major military conflict since the fall of Napoleon: the **Crimean War**. Pitting Imperial Russia against an alliance of Great Britain, France, the Ottoman Empire, and later Piedmont-Sardinia, the war arose out of the intractable **'Eastern Question'**—the geopolitical struggle to manage the agonizing decline of the Ottoman Empire (*'the Sick Man of Europe'*). Historians have frequently endorsed the biting verdict that the Crimean War was a **'futile and unnecessary carnage'**, ignited by a farcical dispute over the keys to Christian shrines in Bethlehem and waged with catastrophic administrative incompetence. Yet, while the military campaigns were marked by tactical blunders and horrific medical neglect, the war's geopolitical consequences were revolutionary: it destroyed the 1815 Holy Alliance, permanently isolated the Austrian Empire, embittered Russia, and created the diplomatic vacuum that enabled the unification of both Italy and Germany."),

    ("The Causes: The Eastern Question and the Holy Places Pretext",
     "The root cause of the war lay in Tsar Nicholas I’s ambition to exploit Ottoman decay and establish Russian hegemony over the Balkan peninsula and the Turkish Straits (Bosphorus and Dardanelles). To Great Britain, Russian access to the Mediterranean represented an existential threat to its maritime lines of communication to India. To Emperor Napoleon III of France, the conflict offered an opportunity to win domestic Catholic favor and tear up the hated 1815 Vienna treaties. The immediate catalyst was deceptively trivial: a squabble between Roman Catholic and Eastern Orthodox monks over the custody of the Holy Places in Ottoman-controlled Palestine. When Sultan Abdulmejid I favored Catholic claims, Tsar Nicholas dispatched **Prince Menshikov** in May 1853 with an unacceptable ultimatum demanding a Russian protectorate over all twelve million Orthodox Christian subjects in the Ottoman Empire. Encouraged by British ambassador Stratford de Redcliffe, the Sultan rejected the ultimatum, prompting Russia to invade the Danubian Principalities (Moldavia and Wallachia)."),

    ("The Outbreak of War and the Battle of Sinope (1853)",
     "The diplomatic crisis turned into open war following the **Battle of Sinope (November 30, 1853)**, where Russian naval shells annihilated an Ottoman flotilla in the Black Sea. Denounced by the British press as the 'Massacre of Sinope,' the engagement provoked overwhelming anti-Russian war hysteria in London and Paris. In March 1854, Britain and France declared war, dispatching an expeditionary force to the Black Sea with the objective of capturing and destroying **Sevastopol**, the fortified citadel of Russian naval power in the Crimea. In 1855, Count Cavour committed 15,000 Piedmontese troops to the allied siege, calculating that blood spilled on the Crimean battlefield would earn Piedmont a seat at the future peace table to plead the cause of Italian liberation."),

    ("Evaluation: Was it a 'Futile and Unnecessary Carnage'?",
     "From an operational and human standpoint, the Crimean War was unquestionably an appalling and futile slaughter. Over 650,000 soldiers perished, with four out of every five casualties succumbing not to enemy artillery, but to cholera, dysentery, typhus, and gangrene in frostbitten, rat-infested trenches. Military leadership on all sides displayed grotesque incompetence: the British cavalry was decimated in the suicidal **'Charge of the Light Brigade'** at Balaclava (October 1854) due to blundered orders, while Russian logistics collapsed under corrupt Tsarist bureaucracy. The war was only redeemed in the civilian imagination through the pioneering frontline journalism of William Howard Russell and the revolutionary nursing reforms of **Florence Nightingale** at Scutari. Furthermore, its immediate military objective—the neutralization of the Black Sea—proved utterly temporary, as Russia unilaterally repudiated the clause a mere fourteen years later in 1870."),

    ("Diplomatic Consequences: The Treaty of Paris and the Shattered Concert",
     "While militarily futile, the war produced a seismic diplomatic realignment at the **Congress of Paris (March 1856)**. The **Treaty of Paris** compelled Russia to cede southern Bessarabia, demilitarized the Black Sea, and placed the Ottoman Empire under collective European guarantee. Crucially, the war shattered the conservative consensus that had maintained European peace since 1815. By threatening Russia with war in 1854 despite Tsar Nicholas I having saved the Habsburgs during the 1848 Hungarian Revolution, **Austria committed an act of colossal diplomatic ingratitude**. Austria alienated Russia permanently without winning the friendship of Britain or France. Isolated between a vengeful Russia and an ambitious France, Austria was left entirely friendless when Cavour struck in 1859 and Bismarck attacked in 1866."),

    ("Conclusion",
     "In final historical judgment, the Crimean War fully justifies the charge of 'futile carnage' on the battlefield, but stands as the great diplomatic watershed of the nineteenth century. It settled neither the Eastern Question nor the fate of the Ottoman Empire, which continued to disintegrate until 1918. However, by destroying the conservative alliance between St. Petersburg and Vienna, the war liquidated the Metternichian order that had suppressed European nationalism for forty years. It demonstrated Tsarist backwardness, forcing Russia into the Great Reforms of Alexander II, while leaving Central Europe fluid. In the famous words of A.J.P. Taylor, the Crimean War was fought for trivial reasons, conducted with colossal blunders, and concluded with a treaty that did not last; yet without it, neither the Kingdom of Italy nor the German Empire could ever have been born.")
]

quotes_8 = [
    ("A.J.P. Taylor", "The Struggle for Mastery in Europe", "The Crimean War was a classic war of miscalculation, fought without any clear purpose and concluded with a peace treaty that vanished within a generation."),
    ("Lord Salisbury", "British Prime Minister", "British policy in the Crimea was an error; we put our money on the wrong horse by propping up the decaying Turkish Empire."),
    ("Paul W. Schroeder", "Austria, Great Britain and the Crimean War", "The Crimean War was the decisive turning point of nineteenth-century European history; it destroyed the system of 1815 and made the violent conflicts of the 1860s possible."),
    ("David Thomson", "Europe Since Napoleon", "The war was fought with the greatest gallantry by the troops, the greatest incompetence by the commanders, and the greatest irrelevance to any constructive purpose."),
    ("Prince Felix von Schwarzenberg", "Austrian Prime Minister", "Austria will astonish the world by the magnitude of her ingratitude.")
]

# =========================================================================
# QUESTION 9: 2025 Q.8 / 2022 Q.7 / 2018 Q.8 / 2016 Q.8 / 2013 Q.8 (ALLIANCE SYSTEM & ORIGINS OF WWI)
# =========================================================================

meta_9 = [
    ("EXAM QUESTIONS LINKED:", "2025 Q.8: 'Explain the long-term causes of WWI (Militarism, Alliances, Imperialism, Nationalism). How did the July Crisis act as the trigger?'\n2022 Q.7: 'Discuss the role of the Alliance System and the Arms Race as two major causes of WWI.'\n2018 Q.8: 'How far was the Triple Entente responsible for dividing Europe into two armed camps before 1914?'\n2016 Q.8: 'Examine the complex network of secret alliances and imperialism from 1890 to 1914 that made WWI inevitable.'\n2013 Q.8: 'Alliances and counter-alliances led the European world towards WWI. Discuss.'"),
    ("TOTAL MARKS & TIME:", "20 Marks &bull; Descriptive Subjective Essay &bull; Time: 35 Minutes"),
    ("SYLLABUS CORRELATION:", "Paper I &bull; Section 8: Armed Peace, The Alliance System & Road to WWI (1870–1914)"),
    ("HISTORIOGRAPHICAL THESIS:", "A tragedy of mechanistic entrapment: Bismarck's defensive web degenerated under Wilhelmine Weltpolitik into two rigid, armed coalitions bound by hair-trigger mobilization timetables, detonating when Balkan nationalism ignited the July Crisis.")
]

callout_9 = [
    "1. The MAIN Framework: Militarism (Anglo-German naval race, army expansions), Alliances (Triple Alliance vs. Triple Entente), Imperialism (Scramble for Africa, Moroccan crises), and Nationalism (Pan-Slavism, French Revanchism over Alsace-Lorraine).",
    "2. Bismarckian vs. Wilhelmine Diplomacy: Bismarck kept France isolated and maintained a Russian reinsurance wire; Kaiser Wilhelm II fired Bismarck (1890), allowed the Reinsurance Treaty to lapse, and drove republican France and autocratic Russia into military alliance (1894).",
    "3. The Mechanism of Escalation: Secret alliances turned a local Austro-Serbian clash into a global conflagration; military mobilization timetables (Schlieffen Plan) meant mobilization equaled war."
]

map_9 = (
    "GERMAN HEGEMONY (1871): Bismarck's Secret Web ➔ Isolate France + Maintain Russian Friendship\n"
    "  │\n"
    "  ├── 1890 SHIFT: Kaiser Wilhelm II fires Bismarck ➔ Drops Reinsurance Treaty with Russia\n"
    "  │     └── 1894 FRANCO-RUSSIAN ALLIANCE: Nightmare of two-front war born for Germany\n"
    "  │\n"
    "  ├── IMPERIAL CLASH & ARMS RACE (1898–1907):\n"
    "  │     ├── Anglo-German Naval Race (Tirpitz Naval Laws vs. HMS Dreadnought)\n"
    "  │     └── Britain abandons 'Splendid Isolation' ➔ Entente Cordiale (1904) & Anglo-Russian Entente (1907)\n"
    "  │\n"
    "  ├── THE TWO RIGID ARMED CAMPS: Triple Alliance (Germany, Austria, Italy) vs. Triple Entente (Britain, France, Russia)\n"
    "  │\n"
    "  └── BALKAN POWDER KEG & THE JULY CRISIS (1914):\n"
    "        ├── Serbian Pan-Slavism ➔ June 28: Sarajevo Assassination of Franz Ferdinand\n"
    "        ├── German 'Blank Check' to Austria ➔ Austro-Hungarian Ultimatum to Serbia\n"
    "        └── MOBILIZATION CASCADE: Russian mobilization ➔ Schlieffen Plan ➔ Belgian invasion ➔ Global War"
)

timeline_9 = [
    ("1879 & 1882", "Dual Alliance & Triple Alliance", "Bismarck binds Austria-Hungary and Italy to Germany in defensive anti-French alliances."),
    ("June 1887", "Reinsurance Treaty with Russia", "Bismarck’s masterpiece secret pact preventing a Franco-Russian encirclement of Germany."),
    ("March 1890", "Dropping the Pilot (Kaiser Wilhelm II)", "Bismarck dismissed; Wilhelm II abandons Reinsurance Treaty; begins aggressive Weltpolitik."),
    ("Jan 1894", "Franco-Russian Military Alliance", "Republican France and Tsarist Russia unite militarily; ends French diplomatic isolation."),
    ("1898–1906", "Anglo-German Naval Race & Dreadnought", "Admiral von Tirpitz expands high seas fleet; Britain launches revolutionary all-big-gun HMS Dreadnought."),
    ("April 1904", "Anglo-French Entente Cordiale", "Britain and France settle colonial disputes in Egypt and Morocco; diplomatic rapprochement."),
    ("Aug 1907", "Anglo-Russian Convention", "Britain and Russia settle Persian/Afghan rivalries; Triple Entente formally completed."),
    ("1908 & 1912", "Bosnian Crisis & Balkan Wars", "Austrian annexation of Bosnia humiliates Serbia and Russia; Balkans become the 'powder keg'."),
    ("June 28, 1914", "Assassination at Sarajevo", "Gavrilo Princip (Black Hand) shoots Archduke Franz Ferdinand; sparks the July Crisis."),
    ("July 28–Aug 4", "The Mobilization Domino Cascade", "Austria declares war on Serbia; Russia mobilizes; Germany executes Schlieffen Plan; Britain enters war.")
]

essay_9 = [
    ("Introduction",
     "The outbreak of the First World War in August 1914 shattered a century of relative European stability, unleashing a four-year industrialized slaughter that obliterated four empires and transformed global civilization. The catastrophe was not an accidental tragedy born in a vacuum, but the explosive culmination of four structural developments summarized by the **'MAIN' framework**: **Militarism**, the **Alliance System**, **Imperialism**, and **Nationalism**. While the assassination of Archduke Franz Ferdinand at Sarajevo on June 28, 1914, acted as the spark, the continental powder keg had been methodically packed over four decades. The transformation of Otto von Bismarck’s fluid, defensive alliance network into two armed, antagonistic camps—the **Triple Alliance** and the **Triple Entente**—subverted diplomatic flexibility into mechanistic entrapment, ensuring that a localized Balkan crisis inevitably detonated a global conflagration."),

    ("From Bismarckian Equilibrium to Two Armed Camps (1871–1907)",
     "Following the unification of Germany in 1871, Bismarck recognized that the new Reich was a satiated, yet vulnerable, continental power. To secure Germany, Bismarck constructed an intricate alliance network based upon two core principles: keeping revanchist France diplomatically isolated, and preventing Austria-Hungary and Russia from clashing over the Balkans. Through the **Dual Alliance with Austria (1879)**, the **Triple Alliance with Italy (1882)**, and the secret **Reinsurance Treaty with Russia (1887)**, Bismarck juggled conflicting imperial interests with peerless virtuosity. However, the catastrophic rupture occurred in 1890 when the impetuous **Kaiser Wilhelm II** dismissed Bismarck and allowed the Reinsurance Treaty to lapse. Seeking an exit from isolation, republican France seized the opportunity to woo autocratic St. Petersburg, culminating in the **Franco-Russian Military Convention of 1894**. Germany now faced its ultimate strategic nightmare: an armed, two-front encirclement."),

    ("Militarism, the Naval Race, and the British Diplomatic Revolution",
     "Simultaneously, the aggressive German pursuit of **Weltpolitik** (World Policy) and the enactment of the **Tirpitz Naval Laws (1898–1900)** challenged Great Britain's undisputed naval mastery. By constructing a battle fleet designed to challenge the Royal Navy in the North Sea, Germany forced Britain to abandon its cherished policy of 'Splendid Isolation.' Britain revolutionized maritime technology with **HMS Dreadnought (1906)** and resolved its historical colonial conflicts with former rivals: signing the **Entente Cordiale with France (1904)** over Egypt and Morocco, and the **Anglo-Russian Convention (1907)** over Persia and Afghanistan. By 1907, Europe was irrevocably partitioned into two rigid armed camps: the Triple Alliance (Germany, Austria-Hungary, Italy) and the Triple Entente (Britain, France, Russia)."),

    ("Imperial Crises and the Balkan Powder Keg (1908–1913)",
     "Between 1905 and 1913, a succession of geopolitical crises heightened military tensions and tested the solidarity of both alliances. German attempts to fracture the Entente in the **First and Second Moroccan Crises (1905, 1911)** boomeranged, driving Britain and France into joint military contingency planning. Concurrently, the collapse of Ottoman power in Europe transformed the **Balkans into the 'powder keg of Europe'**. The **Bosnian Crisis of 1908**, where Austria annexed Bosnia-Herzegovina with German backing, inflicted a humiliating diplomatic defeat upon Russia and enraged Serbian nationalism. Following the bloody **Balkan Wars (1912–1913)**, which doubled Serbian territory, an existential conflict loomed between a multi-ethnic Austro-Hungarian Empire terrified of dissolution and an aggressive Serbian kingdom aspiring to lead a South Slav union backed by St. Petersburg."),

    ("The July Crisis: Mobilization Timetables and the Slide to War",
     "The catalyst arrived on **June 28, 1914**, when the Serbian terrorist network 'Black Hand' assassinated the Austro-Hungarian heir, Archduke Franz Ferdinand, in Sarajevo. Determined to crush Serbia permanently, Austria secured a unconditional **'Blank Check'** from Kaiser Wilhelm II on July 5, guaranteeing full German military support. Austria issued an intentionally unacceptable ultimatum to Belgrade on July 23, declaring war on July 28. At this critical juncture, the rigid machinery of military mobilization overwhelmed diplomatic mediation. In nineteenth-century military doctrine, mobilization *meant* war: railway logistics and mobilization timetables were inflexible. Germany's war plan—the **Schlieffen Plan**—dictated that to survive a two-front war, Germany must rapidly knock out France by invading neutral Belgium before slowly mobilizing Russia could deploy. When Tsar Nicholas II ordered general mobilization to protect Serbia, Germany declared war on Russia and France, invading Belgium on August 3 and prompting Britain's immediate entry into the war."),

    ("Conclusion",
     "In final analysis, the outbreak of the First World War was the product of systemic failure rather than single-handed design. As historian Christopher Clark demonstrated in *The Sleepwalkers*, the leaders of Europe walked into tragedy blind to the magnitude of the catastrophe they were unleashing. The alliance system, initially constructed to deter aggression through balance, had transformed into an escalatory transmission belt: an assassination in a remote Balkan city dragged six Great Powers and sixty million men into the slaughter. Bound by secret pledges, seduced by militaristic hubris, and paralyzed by rigid railway timetables, European statesmen proved incapable of halting the descent into the abyss. The 'Armed Peace' of Europe died in August 1914, burying four grand empires and ushering in the violent dawn of the modern twentieth century.")
]

quotes_9 = [
    ("Christopher Clark", "The Sleepwalkers (2012)", "The protagonists of 1914 were sleepwalkers, watchful but unseeing, blind to the reality of the horror they were about to bring into the world."),
    ("Fritz Fischer", "Griff nach der Weltmacht (1961)", "Germany’s imperial leadership deliberately seized the crisis of 1914 to wage a preventive war for European hegemony and world power."),
    ("A.J.P. Taylor", "War by Timetable", "The First World War was produced by military timetables; the railway schedules of the mobilization plans left no time for diplomatic hesitation."),
    ("Sidney B. Fay", "The Origins of the World War", "No one power was solely responsible; the war was caused by the system of secret alliances, militarism, economic rivalries, and the poisoning of public opinion by the press."),
    ("Sir Edward Grey", "British Foreign Secretary (Aug 3, 1914)", "The lamps are going out all over Europe; we shall not see them lit again in our lifetime.")
]

out_8 = os.path.join(paper1_dir, "Q8_Crimean_War_and_Eastern_Question_Master_Notes.docx")
out_9 = os.path.join(paper1_dir, "Q9_Origins_of_WWI_and_Alliance_System_Master_Notes.docx")

generate_master_doc(meta_8, callout_8, map_8, timeline_8, essay_8, quotes_8, out_8, "MASTER STUDY NOTE & MODEL ANSWER &bull; CRIMEAN WAR & EASTERN QUESTION")
generate_master_doc(meta_9, callout_9, map_9, timeline_9, essay_9, quotes_9, out_9, "MASTER STUDY NOTE & MODEL ANSWER &bull; ALLIANCE SYSTEM & ORIGINS OF WWI")

q8_md = """# KPPSC PMS European History &bull; Paper I

## Question 8 (2022 Q.5 / 2016 Q.6 / 2018 Q.7):
> **"Highlight the causes and consequences of the Crimean War (1853–1856). Do you agree that it was a 'futile and unnecessary carnage'? Examine its impact on European diplomacy and the Treaty of Paris (1856)."**  
> *(Total Marks: 20 &bull; Word Count: ~650 words &bull; Standard Exam Length)*

---
"""
for heading, body in essay_8:
    q8_md += f"\n### {heading}\n{body}\n"

q9_md = """# KPPSC PMS European History &bull; Paper I

## Question 9 (2025 Q.8 / 2022 Q.7 / 2018 Q.8 / 2016 Q.8 / 2013 Q.8):
> **"Explain the long-term causes of World War I (Militarism, Alliances, Imperialism, Nationalism). How did Bismarck's Alliance System evolve into two rigid armed camps, and how did the July Crisis act as the trigger?"**  
> *(Total Marks: 20 &bull; Word Count: ~650 words &bull; Standard Exam Length)*

---
"""
for heading, body in essay_9:
    q9_md += f"\n### {heading}\n{body}\n"

render_md_and_pdf(q8_md, "Q8_Crimean_War_and_Eastern_Question_Master_Notes")
render_md_and_pdf(q9_md, "Q9_Origins_of_WWI_and_Alliance_System_Master_Notes")
