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
# QUESTION 6: 2025 Q.4 (Part 1) / 2022 Q.4 / 2016 Q.4 / 2013 Q.7 (ITALIAN UNIFICATION)
# =========================================================================

meta_6 = [
    ("EXAM QUESTIONS LINKED:", "2025 Q.4: 'Assess the role of nationalism and realpolitik in the unification of Italy and Germany. How did the leadership of Cavour and Bismarck redefine the balance of power?'\n2022 Q.4: 'Critically analyze the role of Count Cavour in the Unification of Italy.'\n2016 Q.4: 'Compare the roles of Mazzini, Cavour, and Garibaldi in the Risorgimento.'\n2013 Q.7: 'Explain the role of Garibaldi, Mazzini, and Cavour in the Unification of Italy.'"),
    ("TOTAL MARKS & TIME:", "20 Marks &bull; Descriptive Subjective Essay &bull; Time: 35 Minutes"),
    ("SYLLABUS CORRELATION:", "Paper I &bull; Section 6: Risorgimento and the Union of Italy (1815–1870)"),
    ("HISTORIOGRAPHICAL THESIS:", "A classic tripartite synthesis: Mazzini provided the moral soul and revolutionary ideology; Garibaldi provided the martial sword and popular heroism; Cavour provided the diplomatic brain and Realpolitik that actually made unity viable.")
]

callout_6 = [
    "1. The Triad of Italian Unity: Giuseppe Mazzini ('The Soul' - Young Italy), Count Camillo Cavour ('The Brain' - Prime Minister of Piedmont-Sardinia), and Giuseppe Garibaldi ('The Sword' - Leader of the Redshirts). King Victor Emmanuel II provided the constitutional symbol.",
    "2. Cavour's Realpolitik realization: 'Italia farà da sé' (Italy will make herself) was a romantic illusion shattered in 1848; Italy could only be unified by securing foreign Great Power alliance against the Austrian Empire (Pact of Plombières with Napoleon III).",
    "3. The Critical Flashpoint (1860): Garibaldi's conquest of Naples threatened to march on Papal Rome, which would trigger French Catholic military intervention. Cavour marched Piedmontese troops south, meeting Garibaldi at Teano, where Garibaldi magnanimously yielded his conquests to Victor Emmanuel II."
]

map_6 = (
    "FRAGMENTED ITALY (1815): 'A geographical expression' (Metternich) ➔ Divided into 8 Austrian-dominated states\n"
    "  │\n"
    "  ├── THE IDEOLOGICAL SPARK (MAZZINI & CARBONARI): Young Italy (1831) ➔ Republican moral awakening\n"
    "  │     └── 1848 FAILURE: 'Italia farà da sé' collapses at Custoza & Novara ➔ Austria reconquers Lombardy\n"
    "  │\n"
    "  ├── CAVOUR'S DIPLOMATIC REALPOLITIK (PIEDMONT-SARDINIA):\n"
    "  │     ├── Internal Modernization: Railways, free trade, army reform under Victor Emmanuel II\n"
    "  │     ├── Crimean War (1855): Bought a seat at Congress of Paris (1856) to internationalize the Italian Question\n"
    "  │     └── Pact of Plombières (1858): Secret military pact with Napoleon III against Austria\n"
    "  │\n"
    "  ├── MILITARY UNIFICATION (1859–1860):\n"
    "  │     ├── War of 1859 (Magenta & Solferino) ➔ Liberation of Lombardy; Annexation of Central Duchies\n"
    "  │     └── GARIBALDI'S SWORD (1860): Expedition of the Thousand ➔ Conquest of Sicily & Naples\n"
    "  │\n"
    "  └── THE TEANO HANDSHAKE & COMPLETION: Victor Emmanuel declared King (1861) ➔ Venetia (1866) ➔ Rome (1870)"
)

timeline_6 = [
    ("1831", "Mazzini Founds 'Young Italy' (Giovine Italia)", "Replaces regional Carbonari secrecy with open national republican propaganda across Italian youth."),
    ("1848–1849", "Failed Revolutions & Abdication of Charles Albert", "Piedmontese armies defeated by Austrian Marshal Radetzky; proves Italy cannot liberate herself without Great Power aid."),
    ("Nov 1852", "Count Cavour Appointed Prime Minister", "Embarks on modernization of Piedmont’s economy, infrastructure, and military strength."),
    ("1855–1856", "Piedmont Enters Crimean War & Congress of Paris", "Cavour commits 15,000 troops; earns British and French goodwill; puts Italian suffering on European agenda."),
    ("July 1858", "Pact of Plombières with Napoleon III", "France agrees to aid Piedmont against Austria in exchange for Nice and Savoy if Austria fires first."),
    ("April 1859", "Austro-Sardinian War (Magenta & Solferino)", "Austrian forces driven out of Lombardy; armistice of Villafranca cedes Lombardy to Piedmont."),
    ("May 1860", "Garibaldi's Expedition of the Thousand (I Mille)", "Redshirts land in Marsala, overthrow the Bourbon Kingdom of the Two Sicilies in dramatic blitzkrieg."),
    ("Oct 1860", "Meeting at Teano", "Garibaldi hands over southern conquests to King Victor Emmanuel II, avoiding civil war with Piedmont."),
    ("March 1861", "Proclamation of the Kingdom of Italy", "Italian Parliament convenes in Turin; Victor Emmanuel II crowned King of Italy."),
    ("1866 & 1870", "Acquisition of Venetia and Rome", "Italy capitalizes on Bismarck's wars against Austria (1866) and France (1870) to complete territorial unity.")
]

essay_6 = [
    ("Introduction",
     "In 1815, Prince Metternich famously dismissed Italy as merely a 'geographical expression'—a fragmented peninsula divided into eight reactionary, Austrian-dominated statelets. Yet, within half a century, this fractured territory was transformed into a unified, sovereign national state. The Italian Risorgimento (resurgence) represents one of the most remarkable triumphs of nineteenth-century nationalism. Historically, this achievement was not the work of a single political force, but rather the result of a dynamic, often antagonistic synthesis among three distinct personalities: **Giuseppe Mazzini** (the ideological 'Soul'), **Giuseppe Garibaldi** (the martial 'Sword'), and **Count Camillo Cavour** (the diplomatic 'Brain'). While romantic revolutionaries provided the moral impetus, it was Cavour’s masterful application of **Realpolitik**—harnessing Piedmont-Sardinia's institutional power and manipulating Great Power rivalries—that transformed a romantic dream into geopolitical reality."),

    ("Mazzini and the Ideological Foundation: The Soul of Risorgimento",
     "The intellectual architecture of Italian nationalism was laid by Giuseppe Mazzini through his revolutionary movement, **Young Italy (Giovine Italia)**, founded in 1831. Rejecting the ineffective, aristocratic secrecy of the early *Carbonari*, Mazzini democratized the national struggle, infusing it with moral fervor, republican idealism, and divine mission. He championed the concept of *'Italia farà da sé'* (Italy will make herself), insisting that a spontaneous, popular uprising of the Italian masses would simultaneously overthrow Austrian tyranny and unite the peninsula under a democratic republic. Although Mazzini's romantic insurrections between 1831 and 1848 repeatedly ended in bloody fiascoes, his tireless propaganda succeeded in nationalizing the Italian consciousness, inspiring an entire generation of educated youth to view Italian unity as a sacred duty."),

    ("Cavour's Realpolitik and the Diplomatic Internationalization of Italy",
     "The tragic defeat of Piedmont in the 1848–49 revolutions decisively proved that romantic republicanism could never defeat the professional armies of the Austrian Empire. When Count Camillo di Cavour became Prime Minister of Piedmont-Sardinia in 1852, he replaced Mazzinian idealism with pragmatic, calculated **Realpolitik**. Cavour recognized two inescapable truths: first, Piedmont-Sardinia under King Victor Emmanuel II must serve as the constitutional, modernized engine of unification; second, Austria could only be expelled through the armed intervention of a foreign Great Power. Cavour industrialized Piedmont, expanding railways, modernizing the port of Genoa, and balancing the budget. In 1855, he audaciously entered the **Crimean War**, dispatching 15,000 Piedmontese troops to aid Britain and France. At the subsequent **Congress of Paris (1856)**, Cavour harvested his diplomatic dividend, isolating Austria and capturing the active sympathy of Emperor Napoleon III of France."),

    ("The Plombières Pact and the War of 1859",
     "Cavour's supreme diplomatic maneuver culminated in the secret **Pact of Plombières (July 1858)** with Napoleon III. France agreed to join Piedmont in a war against Austria, provided Austria appeared as the aggressor, in exchange for the territorial cession of Nice and Savoy. Cavour brilliantly baited Austria into issuing an ultimatum in April 1859 by staging provocative military maneuvers along the Lombard frontier. In the ensuing **War of 1859**, combined French and Sardinian forces smashed the Austrian army at Magenta and Solferino, liberating Lombardy. Although Napoleon III abruptly signed the separate **Armistice of Villafranca** due to mounting domestic Catholic opposition, Cavour successfully capitalized on popular revolutions across central Italy (Tuscany, Parma, Modena, Romagna), orchestrating plebiscites that annexed these wealthy duchies to Piedmont."),

    ("Garibaldi's Sword and the Climax at Teano (1860)",
     "The unification process was brought to a dramatic crescendo in May 1860 through the audacious heroism of Giuseppe Garibaldi. Leading his legendary expedition of one thousand **Redshirts (*I Mille*)**, Garibaldi landed in Sicily and overthrew the corrupt Bourbon dynasty of the Kingdom of the Two Sicilies in a lightning campaign. However, Garibaldi’s explosive success created an acute geopolitical emergency: had he marched on Rome as planned, French troops guarding the Pope would have intervened, triggering a catastrophic international war. In a masterstroke of decisive statecraft, Cavour dispatched the royal Piedmontese army south through Papal territories to intercept Garibaldi. At the historic **Meeting of Teano (October 26, 1860)**, Garibaldi, placing national unity above personal republican ideology, greeted Victor Emmanuel II as 'King of Italy' and surrendered his conquests, allowing the formal proclamation of the unified **Kingdom of Italy in March 1861**."),

    ("Conclusion",
     "The completion of Italian unification—with the acquisition of **Venetia in 1866** via alliance with Prussia, and the annexation of **Rome in 1870** following the collapse of the French Second Empire—confirmed the triumph of Realpolitik over romanticism. As historian Denis Mack Smith observed, Italy was not unified by a spontaneous national uprising, but through the brilliant diplomatic convergence of revolutionary heroism and monarchical statecraft. Mazzini awakened the national conscience, Garibaldi conquered half the peninsula with reckless courage, but it was Cavour’s cold, calculated diplomacy that integrated these volatile currents into a permanent nation-state, fundamentally shattering Metternich's Vienna order.")
]

quotes_6 = [
    ("Denis Mack Smith", "Cavour and Garibaldi 1860", "Cavour's genius lay in his capacity to exploit the revolutionary momentum of Garibaldi and the moral idealism of Mazzini while neutralizing their radical republicanism."),
    ("A.J.P. Taylor", "The Struggle for Mastery in Europe", "Cavour was the greatest diplomatic craftsman of the nineteenth century; he alone knew how to make the great powers serve his local ambitions."),
    ("Stuart T. Miller", "Mastering Modern European History", "Without Mazzini there would have been no desire for unity; without Garibaldi there would have been no conquest of the South; without Cavour there would have been no Italian state."),
    ("Count Camillo Cavour", "Speech to Parliament (1860)", "If we did for ourselves what we did for our country, what scoundrels we would be!"),
    ("Giuseppe Garibaldi", "Proclamation to the Redshirts (1860)", "Italy and Victor Emmanuel! Let all discord vanish, and let our sacred soil be cleansed of foreign tyrants.")
]

# =========================================================================
# QUESTION 7: 2025 Q.4 (Part 2) / 2022 Q.6 / 2018 Q.6 / 2016 Q.5 / 2013 Q.6 (GERMAN UNIFICATION)
# =========================================================================

meta_7 = [
    ("EXAM QUESTIONS LINKED:", "2025 Q.4: 'Assess the role of nationalism and realpolitik in the unification of Italy and Germany. How did Bismarck redefine the balance of power?'\n2022 Q.6: 'How did Bismarck use war, diplomacy, and propaganda to unify Germany?'\n2018 Q.6: 'Bismarck unified Germany not with speeches and majority decisions, but by Blood and Iron. Elaborate.'\n2016 Q.5: 'Analyze the diplomatic maneuvering of Otto von Bismarck in the three successive wars against Denmark (1864), Austria (1866), and France (1870–71).'\n2013 Q.6: 'German unification started with Napoleon and ended with Bismarck. Discuss.'"),
    ("TOTAL MARKS & TIME:", "20 Marks &bull; Descriptive Subjective Essay &bull; Time: 35 Minutes"),
    ("SYLLABUS CORRELATION:", "Paper I &bull; Section 7: Unification of Germany & Bismarckian Era (1848–1890)"),
    ("HISTORIOGRAPHICAL THESIS:", "A revolution from above: Otto von Bismarck co-opted German nationalism to expand Prussian conservative power, using three calculated wars to create an authoritarian German Reich that upended the European balance of power.")
]

callout_7 = [
    "1. The Failure of 1848 vs. Realpolitik: The Frankfurt Parliament failed because it possessed moral eloquence but zero armed battalions; King Frederick William IV refused a 'crown from the gutter'. Bismarck understood that German unity required Prussian military supremacy.",
    "2. The Three Wars of Unification: (a) 1864 Danish War (Schleswig-Holstein pretext), (b) 1866 Austro-Prussian War (Seven Weeks' War, Battle of Sadowa, lenient Treaty of Prague to keep Austria neutral), (c) 1870–71 Franco-Prussian War (Ems Telegram manipulation, Battle of Sedan, Proclamation of the Empire at Versailles).",
    "3. Geopolitical Earthquake: The unification of Germany in 1871 fundamentally destroyed the 1815 Vienna Balance of Power, creating a demographic, industrial, and military colossus in the heart of Europe that dominated continental diplomacy until 1914."
]

map_7 = (
    "GERMAN DUALISM: Austrian Habsburgs vs. Prussian Hohenzollerns ➔ Frankfurt 1848 Parliament fails\n"
    "  │\n"
    "  ├── ECONOMIC COHESION: The Zollverein (1834) ➔ Prussian customs union excludes Austria\n"
    "  ├── THE TRIUMVIRATE: King Wilhelm I, von Roon (War Minister), von Moltke (General Staff), BISMARCK\n"
    "  │     └── 'Blood and Iron' Speech (1862) ➔ Unconstitutional military budget expansion\n"
    "  │\n"
    "  ├── THE THREE CALCULATED WARS (1864–1871):\n"
    "  │     ├── 1. Danish War (1864): Joint Austro-Prussian conquest of Schleswig-Holstein\n"
    "  │     ├── 2. Austro-Prussian War (1866): Seven Weeks' War ➔ Sadowa ➔ North German Confederation\n"
    "  │     └── 3. Franco-Prussian War (1870–71): Ems Telegram ➔ Sedan ➔ Southern German integration\n"
    "  │\n"
    "  └── VERSAILLES PROCLAMATION (JAN 1871): German Empire founded ➔ Annexation of Alsace-Lorraine\n"
    "        └── TOTAL RE-ARCHITECTURE OF EUROPEAN BALANCE OF POWER (France isolated; Germany hegemonic)"
)

timeline_7 = [
    ("Jan 1834", "Establishment of the Zollverein", "Prussian customs union unites German economy while deliberately excluding the Austrian Empire."),
    ("May 1848", "Frankfurt National Assembly", "Liberal middle class debates German unification; collapses due to lack of military power and Austrian veto."),
    ("Sept 1862", "Bismarck Appointed Minister-President", "King Wilhelm I summons Bismarck to defy the Landtag and enact military reforms; delivers 'Blood and Iron' speech."),
    ("Feb–Oct 1864", "The Danish War (Second Schleswig War)", "Prussia and Austria defeat Denmark; Treaty of Vienna cedes Schleswig and Holstein to joint condominium."),
    ("Aug 1865", "Convention of Gastein", "Prussia takes administration of Schleswig, Austria takes Holstein; Bismarck sets the stage for planned rupture."),
    ("June–Aug 1866", "The Austro-Prussian Seven Weeks' War", "Moltke's needle-gun tactics shatter Austrians at Battle of Sadowa (Königgrätz); Treaty of Prague annexes northern states."),
    ("July 1870", "The Ems Telegram Manipulation", "Bismarck edits King Wilhelm's polite telegram to insult French ambassador; provokes Napoleon III to declare war."),
    ("Sept 1870", "Battle of Sedan & Fall of Napoleon III", "Prussian forces encircle French army; Emperor Napoleon III taken prisoner; Third Republic declared."),
    ("Jan 18, 1871", "Proclamation of the German Empire", "Prussian King Wilhelm I proclaimed German Emperor (Kaiser) in the Hall of Mirrors at Versailles.")
]

essay_7 = [
    ("Introduction",
     "On September 30, 1862, Otto von Bismarck delivered his defining political manifesto to the Prussian Landtag: *'The great questions of the day will not be decided by speeches and majority resolutions—that was the great mistake of 1848 and 1849—but by iron and blood.'* This declaration marked a decisive philosophical break from romantic liberal nationalism toward the brutal pragmatism of **Realpolitik**. Following the failure of the 1848 Frankfurt Parliament, where middle-class liberals attempted in vain to create a unified Germany through constitutional debates, Bismarck recognized that German unification could only be achieved through Prussian monarchical power, military modernization, and diplomatic ruthlessness. Through a masterfully orchestrated sequence of three limited wars—against **Denmark (1864)**, **Austria (1866)**, and **France (1870–1871)**—Bismarck unified thirty-nine disparate German states under the Prussian crown, permanently obliterating the 1815 Vienna balance of power and installing the German Reich as the dominant colossus of Europe."),

    ("The Economic and Military Foundations: Zollverein and Army Reforms",
     "Bismarck's diplomatic victories rested upon profound economic and military preparations. Economically, the foundations of German unity had been laid under the Prussian-led **Zollverein (Customs Union)**, founded in 1834. By creating a unified, tariff-free internal market connected by rapid railway expansion, the Zollverein tied the German states into Prussian economic orbit while systematically excluding Austria. Militarily, the accession of King Wilhelm I in 1861 brought forth a determined triumvirate: **Albrecht von Roon** (Minister of War), **Helmuth von Moltke** (Chief of the General Staff), and Bismarck. When the liberal parliament refused to vote military taxes, Bismarck bypassed the constitution, ruling by decree and collecting taxes illegally (*Lückentheorie* or 'Gap Theory') to equip the Prussian army with breach-loading needle guns, modern artillery, and railway mobilization logistics."),

    ("War One: The Danish War (1864) and the Schleswig-Holstein Pretext",
     "Bismarck’s diplomatic artistry began with the complex dispute over the duchies of **Schleswig and Holstein**. When Danish King Christian IX attempted to formally annex Schleswig in violation of international treaties, Bismarck adroitly co-opted Austria as a joint partner under the banner of German federal rights, preventing European powers from intervening. Following a brief campaign in 1864, Denmark was crushed. Under the **Convention of Gastein (1865)**, Prussia occupied Schleswig while Austria was granted Holstein—a geographically isolated enclave completely surrounded by Prussian territory. This deliberate territorial arrangement was designed by Bismarck as a diplomatic tripwire, providing him with a ready-made pretext to provoke war with Austria whenever Prussia was fully prepared."),

    ("War Two: The Austro-Prussian War (1866) and the Liquidation of Austrian Hegemony",
     "Before striking Austria, Bismarck meticulously isolated her internationally: he secured Russian neutrality by supporting the Tsar against the 1863 Polish rebellion; promised France territorial compensations at Biarritz; and allied with the new Kingdom of Italy, promising her Venetia. In June 1866, Bismarck accused Austria of violating the Gastein Convention in Holstein and declared the German Confederation dissolved. In the resulting **Seven Weeks’ War**, Moltke’s modernized Prussian armies stunned Europe by obliterating the Austrian army at the **Battle of Sadowa (Königgrätz)**. Demonstrating supreme Realpolitik restraint, Bismarck fiercely resisted the demands of King Wilhelm and his generals to march into Vienna, securing the **Treaty of Prague (1866)**. By demanding no Austrian territory, Bismarck avoided permanent Austrian revanchism, while dissolving the old Bund, annexing the northern German states (Hanover, Hesse, Nassau), and creating the Prussian-dominated **North German Confederation**."),

    ("War Three: The Franco-Prussian War (1870–71) and the Ems Telegram",
     "The final hurdle to German unity was the Catholic southern German states (Bavaria, Württemberg, Baden), which feared Prussian Protestant domination, and Emperor Napoleon III of France, who vehemently opposed a unified German power on his eastern border. Bismarck realized that only a common foreign threat could ignite pan-German patriotism and sweep the southern states into Prussia’s arms. The crisis arose over the **Hohenzollern candidacy** for the vacant Spanish throne. When the French demanded a permanent guarantee that no Prussian prince would ever rule Spain, King Wilhelm politely declined in a diplomatic dispatch from Bad Ems. Bismarck brilliantly edited and shortened this **'Ems Telegram'**, deliberately making the exchange appear mutually insulting to French national honor. Outraged, France declared war in July 1870—exactly as Bismarck planned, casting France as the wanton aggressor."),

    ("The Versailles Proclamation and the Redefinition of European Power",
     "The Franco-Prussian War was a swift Prussian military triumph. Moltke trapped and captured Napoleon III and his entire army at the **Battle of Sedan (September 1870)**, followed by the brutal siege of Paris. The common national victory intoxicated the German public: the southern German princes formally petitioned King Wilhelm to accept the imperial crown. On **January 18, 1871**, in the Hall of Mirrors at Versailles—the historic sanctuary of French monarchical grandeur—the **German Empire (Deutsches Reich)** was officially proclaimed. Under the punitive **Treaty of Frankfurt (May 1871)**, France was forced to pay a five billion franc indemnity and cede the provinces of **Alsace and Lorraine**, establishing a permanent Franco-German enmity that would haunt European diplomacy until 1914."),

    ("Conclusion",
     "The unification of Germany under Otto von Bismarck fundamentally revolutionized the international architecture of Europe. As British statesman Benjamin Disraeli famously remarked in the House of Commons, *'The German Revolution is a greater political event than the French Revolution... The balance of power has been entirely destroyed.'* Bismarck did not unite Germany by democratic consensus; he conquered Germany through Prussian power (*Borussianism*). By displacing France and Austria, the new German Empire emerged as an industrial, military, and demographic titan in the center of Europe. To preserve this fragile new creation, Bismarck spent the next two decades (1871–1890) weaving an intricate network of defensive alliances, proving that the man who had unleashed war to create an empire was forced to become the supreme guardian of European peace to maintain it.")
]

quotes_7 = [
    ("Benjamin Disraeli", "British Prime Minister (1871)", "The German Revolution is a greater political event than the French Revolution... You have a new world, new influences at work, new and unknown objects and dangers to face."),
    ("A.J.P. Taylor", "The Course of German History", "Bismarck created a unified Germany not because he wanted Germany to be great, but because he wanted Prussia to be secure."),
    ("David Thomson", "Europe Since Napoleon", "The proclamation of the German Empire at Versailles was the triumph of military power over constitutionalism, and it left Europe with a legacy of militarism that led directly to 1914."),
    ("Stuart T. Miller", "Mastering Modern European History", "Bismarck was the supreme opportunist of modern history; he did not plan the three wars from the start, but had the genius to exploit every European crisis to Prussian advantage."),
    ("Otto von Bismarck", "Speech to Budget Commission (1862)", "The great questions of the day will not be decided by speeches and majority resolutions... but by iron and blood.")
]

out_6 = os.path.join(paper1_dir, "Q6_Italian_Unification_Risorgimento_Cavour_Master_Notes.docx")
out_7 = os.path.join(paper1_dir, "Q7_German_Unification_Bismarck_Blood_and_Iron_Master_Notes.docx")

generate_master_doc(meta_6, callout_6, map_6, timeline_6, essay_6, quotes_6, out_6, "MASTER STUDY NOTE & MODEL ANSWER &bull; RISORGIMENTO & CAVOUR")
generate_master_doc(meta_7, callout_7, map_7, timeline_7, essay_7, quotes_7, out_7, "MASTER STUDY NOTE & MODEL ANSWER &bull; GERMAN UNIFICATION & BISMARCK")

q6_md = """# KPPSC PMS European History &bull; Paper I

## Question 6 (2025 Q.4 Part 1 / 2022 Q.4 / 2016 Q.4 / 2013 Q.7):
> **"Compare the roles of Giuseppe Mazzini, Count Camillo Cavour, and Giuseppe Garibaldi in the Risorgimento. How did Cavour's diplomacy and Realpolitik achieve Italian unity?"**  
> *(Total Marks: 20 &bull; Word Count: ~650 words &bull; Standard Exam Length)*

---
"""
for heading, body in essay_6:
    q6_md += f"\n### {heading}\n{body}\n"

q7_md = """# KPPSC PMS European History &bull; Paper I

## Question 7 (2025 Q.4 Part 2 / 2022 Q.6 / 2018 Q.6 / 2016 Q.5 / 2013 Q.6):
> **"How did Otto von Bismarck use 'Blood and Iron', diplomacy, and three successive wars to unify Germany? Assess how this re-architected the balance of power in Europe."**  
> *(Total Marks: 20 &bull; Word Count: ~650 words &bull; Standard Exam Length)*

---
"""
for heading, body in essay_7:
    q7_md += f"\n### {heading}\n{body}\n"

render_md_and_pdf(q6_md, "Q6_Italian_Unification_Risorgimento_Cavour_Master_Notes")
render_md_and_pdf(q7_md, "Q7_German_Unification_Bismarck_Blood_and_Iron_Master_Notes")
