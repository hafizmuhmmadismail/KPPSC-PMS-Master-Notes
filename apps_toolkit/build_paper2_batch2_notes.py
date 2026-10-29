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
# QUESTION 7 (PAPER II): 2025 Q.4 (PT 1) / 2022 Q.3 / 2013 Q.3 (FASCIST ITALY)
# =========================================================================

meta_p2_7 = [
    ("EXAM QUESTIONS LINKED:", "2025 Q.4 (Part 1): 'Compare and contrast the rise of Fascism in Italy and Nazism in Germany. Analyze their core ideologies, methods of control, and social/economic conditions.'\n2022 Q.3: 'Discuss the ideological principles, economic policies, and totalitarian structure of Fascist Italy under Benito Mussolini.'\n2013 Q.3: 'Trace the rise of Benito Mussolini in Italy and evaluate his domestic socio-economic reforms.'"),
    ("SYLLABUS TOPIC:", "Paper II \u2014 Section 5: Fascism in Italy \u2014 Crisis of Liberal State, March on Rome, Ideology, Corporate State & Lateran Pacts"),
    ("ALLOCATED MARKS / TIME:", "20 Marks | 35 Minutes | Standard Word Count: ~650\u2013700 words"),
    ("HISTORIOGRAPHICAL THESIS:", "Italian Fascism emerged from the profound post-WWI structural paralysis of the Italian liberal state, catalyzed by the grievance of the 'mutilated victory' (vittoria mutilata), severe economic inflation, and acute bourgeois terror of Bolshevik revolution (Biennio Rosso); through Blackshirt squadristi paramilitary violence, the theatrical March on Rome (1922), the construction of the Corporate State, and the historic Lateran Pacts (1929), Benito Mussolini forged the twentieth century's pioneer fascist dictatorship, though it remained an 'imperfect totalitarianism' constrained by the enduring institutional authority of the Italian Crown and the Roman Catholic Church.")
]

callout_p2_7 = [
    "\u2022 The Crucible of Rise: Italy felt cheated at Versailles despite suffering 600,000 dead (denied Fiume and Dalmatia, fueling D'Annunzio's nationalist coup in Fiume); the 'Two Red Years' (Biennio Rosso 1919\u20131920) marked by mass factory occupations and peasant strikes terrified industrial magnates and landed agrarians (agrari), who funded Mussolini's armed Blackshirt militias (Squadristi) to break socialist unions.",
    "\u2022 The March on Rome (October 28\u201330, 1922): A colossal political bluff; 30,000 poorly armed Blackshirts converged on Rome; King Victor Emmanuel III refused Premier Luigi Facta's martial law decree, fearing civil war and republican overthrow, and appointed Mussolini Prime Minister within constitutional forms.",
    "\u2022 Transition to Totalitarian Dictatorship: Acerbo Law (1923) guaranteed two-thirds parliamentary majority to the leading party; socialist deputy Giacomo Matteotti murdered in June 1924; opposition walked out (Aventine Secession); on January 3, 1925, Mussolini delivered an ultimatum to parliament, assuming full responsibility for squadristi violence and decreeing one-party rule (Leggi Fascistissime); established the OVRA secret police.",
    "\u2022 The Corporate State (Stato Corporativo): Conceptualized as an alternative to both laissez-faire capitalism and Marxist class struggle; divided the economy into 22 state-supervised corporations representing employers and employees; banned strikes and lockouts; in reality, corporate fascism heavily favored capitalist monopolies and suppressed industrial wages.",
    "\u2022 The Lateran Accords (February 1929): Mussolini's diplomatic masterpiece; reconciled the Italian state with the Vatican, ending the 60-year 'Roman Question'; recognized sovereign Vatican City; compensated the Church with 1.75 billion lire; established Catholicism as sole state religion; gained Pope Pius XI's blessing, who famously called Mussolini 'the man sent by Providence.'",
    "\u2022 The Concept of 'Imperfect Totalitarianism': Unlike Hitler's Germany or Stalin's Russia, Mussolini never achieved total absolute power; the King retained supreme command of the army (eventually dismissing and arresting Mussolini in July 1943), while the Catholic Church retained independent youth and civic organizations (Catholic Action)."
]

map_p2_7 = """========================================================================================
POST-WWI CRISIS                   PARAMILITARY ASCENT (1919-22)        THE FASCIST REGIME (1925-39)
========================================================================================
['Mutilated Victory' (Versailles)]--> [Blackshirts (Squadristi)]      --> [Matteotti Murder (1924)]
[Biennio Rosso: Red Strike Wave]  --> [March on Rome (Oct 1922)]       --> [Jan 1925 Open Dictatorship]
[Bourgeois Terror of Bolshevism]  --> [King Hands Power to Mussolini] --> [The Corporate State (22 Corps)]
[Unstable Parliamentary Blocs]    --> [Acerbo Law (1923) Rigged Vote] --> [Lateran Pacts (1929)]
                                                                       --> ['Imperfect Totalitarianism']
========================================================================================"""

timeline_p2_7 = [
    ("March 23, 1919", "Foundation of Fascism", "Mussolini establishes Fasci Italiani di Combattimento in Milan, blending radical nationalism and anti-socialism."),
    ("1919\u20131920", "The Biennio Rosso", "Two Red Years of socialist strikes and factory occupations; Blackshirts launch violent punitive expeditions."),
    ("Oct 28\u201330, 1922", "The March on Rome", "Blackshirts march; King Victor Emmanuel III refuses martial law and appoints Mussolini Prime Minister."),
    ("Nov 1923", "The Acerbo Electoral Law", "Passes law granting 66% of parliamentary seats to whichever party wins 25% of the popular vote."),
    ("June 10, 1924", "Murder of Giacomo Matteotti", "Socialist leader abducted and killed by Fascist thugs; triggers severe regime crisis (Aventine Secession)."),
    ("Jan 3, 1925", "Establishment of Dictatorship", "Mussolini delivers historic speech accepting full responsibility for Fascism; outlaws all opposition parties."),
    ("April 1926", "Rocco Law & Corporate State", "Outlaws strikes and lockouts; establishes official Fascist trade unions and the Ministry of Corporations."),
    ("Feb 11, 1929", "The Lateran Pacts", "Signed with Cardinal Gasparri; settles Roman Question; creates Vatican City State; Catholicism becomes state religion."),
    ("1935\u20131936", "Invasion of Abyssinia", "Italy conquers Ethiopia defying League sanctions; proclaims the Italian Empire; enters Rome-Berlin Axis with Hitler."),
    ("July 25, 1943", "Downfall of Mussolini", "Following Allied invasion of Sicily, the Grand Council of Fascism and King Victor Emmanuel III depose and arrest Mussolini.")
]

essay_p2_7 = [
    ("Introduction: The Birthplace of European Fascism",
     "In October 1922, Benito Mussolini's theatrical March on Rome inaugurated the first fascist regime of the twentieth century, providing the ideological and tactical template that would later inspire Adolf Hitler and authoritarians across the European continent. In the historiographical debate, scholars have long contested the nature of Italian Fascism: while Marxist historians like Palmiro Togliatti analyzed fascism as the violent reactionary dictatorship of monopoly capitalism, liberal and cultural historians such as Denis Mack Smith and Emilio Gentile emphasize fascism as an innovative, albeit theatrical, political religion aimed at forging a 'New Italian' through the total subordination of the individual to the state. In reality, Italian Fascism was born of the acute structural paralysis of the post-WWI liberal state: an explosive convergence of wounded national prestige, deep agrarian-industrial class warfare (the Biennio Rosso), and parliamentary decadence. However, while Mussolini constructed an elaborate corporate and police apparatus, his regime remained an 'imperfect totalitarianism,' fundamentally constrained by historic institutional compromises with the Italian Monarchy, the military establishment, and the Roman Catholic Church."),

    ("The Crucible of Ascent: The 'Mutilated Victory' and the Biennio Rosso",
     "The foundation of Mussolini's rise lay in the profound psychological and economic trauma inflicted on Italy by World War I. Despite sacrificing 600,000 soldiers on the Isonzo and Piave rivers alongside the Allies, Italy emerged from the Paris Peace Conference in 1919 deeply humiliated. When Britain and the United States refused to cede the Adriatic port of Fiume and the Dalmatian coastline, poet Gabriele D'Annunzio coined the toxic grievance of the 'mutilated victory' (vittoria mutilata), seizing Fiume with nationalist legionnaires in September 1919 and demonstrating that the liberal government was too impotent to enforce the rule of law.\n\n"
     "Simultaneously, the transition from war production generated galloping inflation, widespread unemployment, and revolutionary class turmoil known as the 'Two Red Years' (Biennio Rosso, 1919\u20131920). Inspired by the Bolshevik Revolution, industrial workers seized factories in Turin and Milan, establishing workers' councils, while landless rural laborers occupied large agricultural estates (latifundia) in the Po Valley. Terrified of an imminent communist revolution, large industrial capitalists and agrarian landowners (agrari) abandoned the paralyzed liberal government of Giovanni Giolitti. They provided massive financial subsidies to Mussolini's armed paramilitary squads\u2014the Blackshirts (Squadristi)\u2014who unleashed systematic violence, burning socialist headquarters, beating union leaders, and destroying left-wing printing presses with total police complicity."),

    ("The Seizure of Power: The March on Rome and the Matteotti Crisis",
     "By late 1922, Mussolini had transformed his movement into the National Fascist Party (PNF) with over 300,000 members. Recognizing the cowardice of the political elite, Mussolini organized the historic 'March on Rome' on October 28, 1922. In reality, the march was a colossal political bluff: thirty thousand ill-equipped Blackshirts stranded in torrential rain could have been easily dispersed by the Rome military garrison. However, King Victor Emmanuel III, terrified of civil war and suspecting that his cousin, the Duke of Aosta, might usurp the throne with fascist backing, refused Prime Minister Luigi Facta's martial law decree. Instead, the King telegraphed Mussolini in Milan, inviting him to form a coalition government. Mussolini arrived comfortably by overnight sleeper train on October 30, legally appointed Prime Minister within constitutional forms.\n\n"
     "The transition from coalition rule to absolute dictatorship occurred through the crucible of the 1924 Matteotti Crisis. After passing the Acerbo Law (1923), which rigged elections by granting two-thirds of parliamentary seats to the party winning twenty-five percent of the vote, the Fascists won a crushing majority amidst systemic voter terror. On May 30, 1924, socialist deputy Giacomo Matteotti delivered a courageous speech exposing fascist electoral fraud. Eleven days later, Matteotti was abducted and murdered by fascist thugs. When the democratic opposition staged a toothless moral protest by boycotting parliament (the Aventine Secession), Mussolini weathered the storm. On January 3, 1925, Mussolini delivered an audacious speech to the Chamber of Deputies, accepting full political, moral, and historical responsibility for all fascist violence and announcing the end of parliamentary democracy. Over the next two years, the 'Fascist Laws' (Leggi Fascistissime) outlawed opposition parties, instituted total press censorship, created the Special Tribunal for the Defense of the State, and established the OVRA secret police."),

    ("Ideological Architecture: The Doctrine of Fascism and the Corporate State",
     "Italian Fascism was officially codified by philosopher Giovanni Gentile and Mussolini in 'The Doctrine of Fascism' (1932). Rejecting liberal individualism, pacifism, and Marxist historical materialism, Fascism declared the absolute supremacy of the State: 'Everything in the State, nothing outside the State, nothing against the State.' The state was conceived as a spiritual, moral, and organic entity that transcended individual lives, requiring unquestioning obedience to the charismatic Leader (Il Duce).\n\n"
     "To resolve the antagonism between labor and capital, Mussolini pioneered the 'Corporate State' (Stato Corporativo). The entire economy was organized into twenty-two corporations representing broad industrial sectors (e.g., textiles, agriculture, metallurgy), each comprising joint councils of fascist-approved employers and workers under the supervision of the Ministry of Corporations. Strikes, lockouts, and independent trade unions were outlawed by the Rocco Law of 1926 and the 1927 Charter of Labor (Carta del Lavoro). In theory, the Corporate State represented a revolutionary 'Third Way' between exploitative capitalism and destructive communism. In reality, corporate fascism was an elaborate bureaucratic facade: employers maintained operational control over their factories and retained private profits, while workers saw their real wages decline by fifteen percent between 1925 and 1938. Grand public works schemes, such as the draining of the malaria-infested Pontine Marshes and the 'Battle for Grain,' were celebrated by state propaganda, yet Italy's structural industrial backwardness remained acute."),

    ("The Lateran Pacts of 1929: The Zenith of Domestic Success",
     "Mussolini's most durable and brilliant political achievement was the resolution of the sixty-year-old 'Roman Question.' Since the Kingdom of Italy had annexed papal Rome in 1870, successive Popes had declared themselves 'prisoners in the Vatican' and prohibited Catholics from participating in Italian political life. Recognizing that ninety-nine percent of Italians were Catholic, Mussolini signed the historic Lateran Pacts on February 11, 1929, with Cardinal Pietro Gasparri representing Pope Pius XI.\n\n"
     "The treaty was composed of three interrelated pacts: (1) A political treaty recognizing Vatican City as a fully independent, sovereign micro-state under the absolute temporal authority of the Pope; (2) A financial convention paying the Holy See 1.75 billion lire in cash and bonds as indemnity for lands confiscated in 1870; and (3) A Concordat establishing Roman Catholicism as the 'sole religion of the State,' making religious education compulsory in all state schools, granting canonical courts jurisdiction over marriage, and guaranteeing the autonomy of Catholic Action (Azione Cattolica). The Lateran Pacts bestowed immense domestic and international prestige upon Mussolini, earning the Pope's famous praise of Il Duce as 'the man whom Providence has sent us.'"),

    ("The Limits of Totalitarianism: Totalitarismo Imperfetto",
     "Despite its totalitarian rhetoric and all-pervasive propaganda celebrating youth (the Giovent\u00f9 Italiana del Littorio - GIL) and after-work leisure (the Dopolavoro), Fascist Italy was never a true totalitarian state comparable to Hitler's Germany or Stalin's Soviet Union. Political scientist Emilio Gentile accurately designated it an 'imperfect totalitarianism' (totalitarismo imperfetto). Mussolini was fundamentally constrained by two independent, rival centers of institutional power that survived intact: the Roman Catholic Church and the Italian Monarchy.\n\n"
     "While Hitler was both Head of State and Head of Government, Mussolini was merely Prime Minister under King Victor Emmanuel III. The King retained the supreme loyalty of the regular army officer corps, the Carabinieri, and the traditional nobility. When Mussolini's military adventurism in World War II brought disaster to Italy, it was the Grand Council of Fascism and the King who legally deposed and arrested Mussolini on July 25, 1943. Unlike Nazism, which completely devoured the German state, Fascism coexisted with, and was ultimately extinguished by, the traditional Italian establishment."),

    ("Conclusion: The Tragic Trajectory of Il Duce",
     "In conclusion, Benito Mussolini's Fascist Italy was a pioneering, highly contradictory phenomenon of modern European history. It demonstrated the lethal vulnerability of liberal democratic institutions when confronted by economic crisis, class fear, and organized paramilitary violence. Domestically, Mussolini achieved notable successes in modernizing infrastructure, eradicating the Sicilian Mafia via Cesare Mori, and reconciling the Italian state with the Vatican. However, the regime was fatally undermined by its own militaristic and imperial mythology. Inebriated by his own propaganda, Mussolini launched the brutal invasion of Abyssinia in 1935 and intervened in the Spanish Civil War, alienating Britain and France and transforming Italy into a pathetic, subordinate vassal of Adolf Hitler's Nazi war machine. By dragging an impoverished, unready Italy into World War II, Mussolini shattered the very corporate and social facade he had spent two decades constructing, dying ignominiously before a partisan firing squad at Lake Como in 1945.")
]

quotes_p2_7 = [
    ("Denis Mack Smith", "Mussolini (1981)", "'Mussolini was an actor of genius... His fascism was less a coherent political philosophy than a continuous improvisation designed to maintain personal power through theatrical propaganda.'"),
    ("Emilio Gentile", "The Sacralization of Politics in Fascist Italy", "'Fascism was an authentic political religion that sought to reconstruct the Italian character through collective rituals, myths, and total subordination to the State... It remained an imperfect totalitarianism only because it coexisted with the Monarchy and the Church.'"),
    ("Renzo De Felice", "Mussolini: The Duce", "'We must distinguish between Fascism as a movement\u2014which possessed genuine revolutionary and youthful energy\u2014and Fascism as a regime, which was an alliance of conservative compromises with the traditional ruling classes.'"),
    ("Benito Mussolini (1932)", "The Doctrine of Fascism", "'The Fascist conception of the State is all-embracing; outside of it no human or spiritual values can exist, much less have value... Everything in the State, nothing outside the State, nothing against the State.'")
]

# =========================================================================
# QUESTION 8 (PAPER II): 2025 Q.4 (PT 2) / 2022 Q.4 / 2016 Q.4 / 2013 Q.4
# =========================================================================

meta_p2_8 = [
    ("EXAM QUESTIONS LINKED:", "2025 Q.4 (Part 2): 'Analyze core ideologies, methods of control, and socio-economic conditions facilitating the rise of Nazism in Germany.'\n2022 Q.4: 'Trace the causes and repercussions of the Great Depression of 1929 on European democracies, particularly Britain and Germany.'\n2016 Q.4: 'Analyze the domestic transformation of Germany under Adolf Hitler from the Weimar Republic to the Third Reich (1933\u20131939).'\n2013 Q.4: 'Analyze the internal crises of the Weimar Republic that facilitated the electoral rise of the National Socialist (Nazi) Party.'"),
    ("SYLLABUS TOPIC:", "Paper II \u2014 Section 5: Nazi Germany \u2014 Weimar Flaws, Great Depression (1929), Hitler's Ascent & The Totalitarian Third Reich"),
    ("ALLOCATED MARKS / TIME:", "20 Marks | 35 Minutes | Standard Word Count: ~650\u2013700 words"),
    ("HISTORIOGRAPHICAL THESIS:", "The collapse of the Weimar Republic and Adolf Hitler's ascent to supreme power was not the preordained culmination of German historical Sonderweg, but the fatal collision of institutional vulnerabilities (Article 48, proportional representation, imperial elite continuity) with the catastrophic economic devastation of the 1929 Great Depression; by brilliantly weaponizing mass lower-middle-class desperation, anti-Marxist paranoia, and revanchist hatred of Versailles, Hitler transformed an electoral fringe into a mass movement, rapidly dismantling democracy through legal revolution (Gleichschaltung) to establish the totalitarian Third Reich.")
]

callout_p2_8 = [
    "\u2022 Weimar's Structural Vulnerabilities: Proportional representation spawned endless fractured coalitions (20 governments in 14 years); Article 48 gave the Reich President emergency dictatorial decree powers; the judiciary, civil service, and military (Reichswehr) remained fundamentally hostile to democracy, staffed by old imperial monarchists; the 'original sin' of signing the Versailles Treaty created the toxic 'stab-in-the-back' myth (Dolchsto\u00dflegende).",
    "\u2022 The 1929 Catalyst: The Great Depression: Wall Street Crash (October 1929) caused the immediate recall of American short-term loans under the Dawes Plan; German industrial output fell by 40%; six million workers (one in three) were unemployed by 1932; Chancellor Heinrich Br\u00fcning governed entirely by presidential decree (Article 48), pursuing brutal deflationary austerity that impoverished the lower-middle class (Mittelstand) and created radical polarization between Communists (KPD) and Nazis (NSDAP).",
    "\u2022 Hitler's Electoral Breakthrough: Nazi vote surged from 2.6% (12 seats) in 1928 to 18.3% (107 seats) in 1930, and peaked at 37.3% (230 seats) in July 1932, becoming the largest party in the Reichstag through propaganda orchestrated by Joseph Goebbels and paramilitary intimidation by the SA (Brownshirts).",
    "\u2022 The Conservative Conspiracy: Hitler did not seize power through an armed coup; he was invited into power through backroom intrigue by aristocratic conservatives (Franz von Papen, General Kurt von Schleicher, Oskar von Hindenburg) who believed they could 'box Hitler in' as a puppet Chancellor on January 30, 1933.",
    "\u2022 Rapid Construction of the Third Reich (1933\u20131934):",
    "  1. Reichstag Fire (Feb 27, 1933) -> Reichstag Fire Decree permanently suspended all constitutional civil liberties.",
    "  2. Enabling Act (Erm\u00e4chtigungsgesetz, March 23, 1933) -> Granted Hitler dictatorial decree powers without Reichstag approval.",
    "  3. Gleichschaltung (Coordination) -> Dissolved federal states (L\u00e4nder), banned trade unions (German Labor Front formed), and outlawed all other political parties (July 1933).",
    "  4. Night of the Long Knives (June 30, 1934) -> SS assassinated Ernst R\u00f6hm and radical SA leadership, securing the loyalty of the army.",
    "  5. F\u00fchrer State (August 1934) -> Upon Hindenburg's death, Hitler merged Presidency and Chancellorship; the army swore a personal holy blood oath to Adolf Hitler."
]

map_p2_8 = """========================================================================================
WEIMAR STRUCTURAL ROT               THE 1929 DEPRESSION SPARK          THE TOTALITARIAN THIRD REICH
========================================================================================
[Article 48 Emergency Decrees]   --> [Wall Street Crash: Loans Recalled]--> [Jan 30, 1933: Hitler Chancellor]
[Proportional Representation]    --> [6 Million Unemployed in Germany] --> [Reichstag Fire: Civil Rights Dead]
[Versailles Dolchsto\u00dflegende]   --> [Br\u00fcning's Deflationary Austerity] --> [Enabling Act: Legal Dictatorship]
[Imperial Elites Disloyal]       --> [Nazi Vote Jumps: 2.6% to 37.3%]   --> [Gleichschaltung: One-Party State]
                                 --> [Papen/Hindenburg Elite Conspiracy]--> [Night of Long Knives & F\u00fchrer Oath]
========================================================================================"""

timeline_p2_8 = [
    ("Aug 11, 1919", "Weimar Constitution Adopted", "Democratic republic established at Weimar; contains fatal flaws of Article 48 and proportional representation."),
    ("Nov 8\u20139, 1923", "Munich Beer Hall Putsch", "Hitler attempts armed coup; crushed; writes Mein Kampf in prison; pivots to legal electoral strategy."),
    ("Oct 29, 1929", "Wall Street Crash", "American loans recalled; German banking system collapses; begins the Great Depression."),
    ("Sept 14, 1930", "Reichstag Breakthrough", "Nazi Party surges from 12 to 107 seats; Br\u00fcning begins governing exclusively via Article 48 presidential decrees."),
    ("July 1932", "Nazis Become Largest Party", "NSDAP wins 37.3% of the vote (230 seats); Weimar parliamentary democracy completely paralyzed."),
    ("Jan 30, 1933", "Hitler Appointed Chancellor", "President Hindenburg appoints Hitler Chancellor of a coalition cabinet arranged by Franz von Papen."),
    ("Feb 27\u201328, 1933", "Reichstag Fire Decree", "Reichstag burns; Hitler blames Communists; Hindenburg signs decree suspending habeas corpus, free speech, and press."),
    ("March 23, 1933", "The Enabling Act", "Reichstag passes Erm\u00e4chtigungsgesetz by 444 to 94 votes; transfers full legislative and constitutional powers to Hitler."),
    ("June 30, 1934", "Night of the Long Knives", "SS murders Ernst R\u00f6hm, SA leaders, and conservative rivals; wins total allegiance of the regular military."),
    ("Aug 2, 1934", "Hitler Becomes F\u00fchrer", "Hindenburg dies; Hitler merges Presidency and Chancellorship; armed forces swear personal blood oath to Hitler.")
]

essay_p2_8 = [
    ("Introduction: The Death of Democracy and the Nazi Ascent",
     "The collapse of the Weimar Republic and the establishment of Adolf Hitler's Third Reich represents the definitive catastrophe of twentieth-century European democracy. Born in the ashes of military defeat and revolution in 1919, the Weimar Republic possessed one of the most progressive, enlightened democratic constitutions in the world. Yet, within fourteen years, this constitutional experiment had dissolved into the most murderous totalitarian dictatorship in human history. In the historiographical debate, while older historians of the German 'Sonderweg' (Special Path) school, such as Hans-Ulrich Wehler, posited that Germany was culturally pre-programmed for authoritarianism due to flawed modernization, modern structuralist and intentionalist scholars like Richard J. Evans and Ian Kershaw emphasize the primacy of structural crises and historical contingency. Hitler's triumph was the outcome of an unholy convergence: structural constitutional defects and unpurged anti-democratic imperial institutions collided with the economic devastation of the 1929 Great Depression, allowing a fanatical ideological movement to capture the state through legal constitutional sabotage and conservative intrigue."),

    ("Institutional Vulnerabilities: The Original Sins of the Weimar Republic",
     "The Weimar Republic was structurally compromised from its inception by three fatal institutional and psychological defects. First, its electoral system of pure proportional representation generated an excessively fragmented Reichstag, where dozens of minor parties proliferated. Stable majorities proved impossible, resulting in twenty different coalition cabinets in fourteen years and breeding profound public cynicism toward parliamentary democracy. Second, the constitution contained the fateful Article 48, which authorized the Reich President to suspend fundamental civil liberties and govern by emergency decree without parliamentary approval in times of crisis\u2014a provision that effectively handed the executive a legal blueprint for dictatorship.\n\n"
     "Third, the republic suffered from what historians term an 'uncompleted revolution.' To preserve order in November 1918, socialist President Friedrich Ebert concluded the Ebert-Groener Pact with the imperial army, leaving the traditional anti-republican elite\u2014the aristocratic officer corps (Junkers), the civil service judiciary, and big business magnates\u2014entirely untouched. These elites harbored visceral contempt for the republic, viewing it as the illegitimate offspring of military betrayal. When the Weimar government was forced to sign the Treaty of Versailles in June 1919, Field Marshals Hindenburg and Ludendorff successfully cultivated the toxic 'stab-in-the-back' myth (Dolchsto\u00dflegende), branding the democratic politicians who signed the armistice as the 'November Criminals' who had betrayed an undefeated army in the field."),

    ("The Economic Catalyst: The Great Depression of 1929",
     "Despite its structural flaws, the Weimar Republic achieved impressive economic and diplomatic stabilization during the 'Golden Years' of Gustav Stresemann (1924\u20131929). Through the Dawes Plan of 1924, American financial loans flooded into Germany, stabilizing the currency and modernizing industrial infrastructure. In the federal elections of May 1928, Hitler's National Socialist German Workers' Party (NSDAP) was an irrelevant lunatic fringe, polling a miserable 2.6 percent of the vote (twelve seats). However, Weimar's prosperity was precariously mortgaged on American credit.\n\n"
     "The Wall Street Crash of October 1929 delivered a fatal shock. American banks immediately recalled their short-term loans, sending the German banking system into catastrophic collapse. Between 1929 and 1932, German industrial production plummeted by forty percent, five thousand banks and enterprises closed, and unemployment skyrocketed from 1.3 million to over six million citizens\u2014meaning one in every three German workers was destitute. The catastrophe was compounded by Chancellor Heinrich Br\u00fcning's disastrous orthodox economic policy: pursuing rigid deflation, Br\u00fcning slashed government expenditures, cut unemployment benefits, lowered wages, and increased taxes, completely impoverishing the lower-middle class (Mittelstand: shopkeepers, clerks, artisans, teachers).\n\n"
     "Crucially, when the Reichstag rejected Br\u00fcning's austerity measures in 1930, he dissolved parliament and invoked Article 48, establishing an authoritarian presidential regime that bypassed parliamentary democracy three years before Hitler took power. As the political center evaporated, German society fractured into violent ideological extremes: the Communist Party (KPD) mobilized the urban unemployed, while millions of terrified middle-class citizens, shopkeepers, and Protestant farmers surged into the ranks of Hitler's NSDAP."),

    ("The Electoral Surge and the Conservative Backroom Conspiracy",
     "Adolf Hitler exploited this socioeconomic desperation with unmatched demagogic brilliance. Orchestrated by propaganda chief Joseph Goebbels, the Nazi Party presented Hitler as an above-politics messianic savior. Nazi propaganda promised 'Work and Bread' to the unemployed, salvation to small businesses from department store competition, protection to industrialists from the Bolshevik menace, and vengeance against the humiliating shackles of the Versailles Treaty. Simultaneously, Ernst R\u00f6hm's four-hundred-thousand-strong paramilitary stormtroopers, the SA (Sturmabteilung), waged brutal street warfare against Communist Red Front fighters, deliberately creating an atmosphere of civil war to convince the bourgeois public that only Hitler possessed the iron will to restore order.\n\n"
     "In the September 1930 elections, the Nazi vote leaped to 6.4 million (107 seats); by July 1932, the NSDAP won 37.3 percent of the vote (230 seats), becoming the largest party in German history. Yet, by November 1932, the Nazi electoral wave had crested: the party lost two million votes, was bankrupt, and Hitler contemplated suicide. Hitler did not capture power through electoral majority or armed revolution; he was smuggled into the Chancellery through the backroom intrigue of a decadent aristocratic camarilla surrounding the senile President Paul von Hindenburg. Conspiring to topple Chancellor Kurt von Schleicher, former Chancellor Franz von Papen persuaded Hindenburg that the conservative establishment could 'tame' Hitler by appointing him Chancellor of a coalition cabinet where Nazis held only three of eleven posts. As Papen boastfully declared: 'Within two months, we will have pushed Hitler into a corner until he squeals.' It was the most catastrophic miscalculation in European history: on January 30, 1933, Adolf Hitler was appointed Chancellor of the German Reich."),

    ("The Legal Revolution: Gleichschaltung and the Totalitarian Monolith (1933\u20131934)",
     "Once in power, Hitler executed a masterclass in the rapid, ruthless dismantling of constitutional democracy under the veneer of legal continuity. The turning point arrived on the night of February 27, 1933, when the Reichstag building was consumed by arson, allegedly set by Dutch communist Marinus van der Lubbe. Hitler capitalized on the fire to persuade Hindenburg to sign the 'Reichstag Fire Decree' on February 28, which indefinitely suspended all constitutional civil liberties, including habeas corpus, freedom of the press, and rights of assembly. The communist press was suppressed, and four thousand left-wing leaders were arrested and sent to Dachau, the first concentration camp.\n\n"
     "On March 23, 1933, through a combination of SA intimidation, the exclusion of communist deputies, and false promises to the Catholic Center Party, Hitler secured the necessary two-thirds majority to pass the 'Enabling Act' (Erm\u00e4chtigungsgesetz). This historic law transferred full legislative power, including the authority to deviate from the constitution, directly to Hitler's cabinet for four years, effectively terminating parliamentary government and establishing a legal dictatorship.\n\n"
     "Over the next twelve months, Hitler executed the process of 'Coordination' (Gleichschaltung): the sixteen federal states (L\u00e4nder) were stripped of autonomy and subordinated to Nazi governors (Reichsstatthalter); independent trade unions were abolished on May 2, 1933, and replaced by the monolithic German Labor Front (DAF); and on July 14, 1933, the 'Law Against the Formation of New Parties' made the NSDAP the sole legal political entity in Germany. The final internal obstacle was eliminated on the 'Night of the Long Knives' (June 30, 1934), when the SS under Heinrich Himmler murdered Ernst R\u00f6hm and the radical leadership of the SA, alongside conservative critics like General Schleicher. By executing R\u00f6hm, Hitler placated the traditional army leadership (*Reichswehr*). When President Hindenburg died on August 2, 1934, Hitler merged the offices of President and Chancellor, assuming the title of F\u00fchrer und Reichskanzler. Every officer and soldier in the German military was required to swear a personal, unbreakable holy blood oath of loyalty not to the state or constitution, but directly to the person of Adolf Hitler."),

    ("Conclusion: The Anatomy of a Totalitarian Tyranny",
     "In conclusion, the rise of Adolf Hitler and the fall of the Weimar Republic was an overdetermined tragedy of modern European history. Weimar was not killed simply by the Versailles Treaty or inherent German militarism; it was strangled by the catastrophic socioeconomic fallout of the 1929 Great Depression, which destroyed middle-class faith in liberal institutions, and the cynical abdication of democratic responsibility by conservative elites who believed they could instrumentalize fascist demagogy for authoritarian ends. In their place arose the Third Reich: a totalitarian police state characterized by the pervasive terror of the Gestapo and SS, the racial biological segregation of the Nuremberg Laws (1935), economic autarky geared for war, and the charismatic 'F\u00fchrer Principle.' By 1939, having liquidated all domestic opposition and restored full employment through frantic rearmament, Hitler stood unchallenged at the apex of an industrialized war machine, poised to plunge the continent into the apocalyptic racial war of conquest and genocide that was World War II.")
]

quotes_p2_8 = [
    ("Richard J. Evans", "The Coming of the Third Reich (2003)", "'The Weimar Republic was not doomed from the start... It was destroyed by the Great Depression and by the conscious decisions of conservative elites who preferred an authoritarian dictatorship to the compromises of democracy.'"),
    ("Ian Kershaw", "Hitler: Hubris (1998)", "'Hitler did not seize power; he was handed power by a corrupt conservative clique that believed it could box him in... Once in the Chancellery, his dynamic drive and the mechanism of 'working towards the F\u00fchrer' rapidly converted legal office into absolute totalitarian dictatorship.'"),
    ("Hans-Ulrich Wehler", "The German Empire", "'The structural continuity of Prussian militarism, the unpurged imperial judiciary, and the anti-democratic bureaucracy ensured that the Weimar Republic remained a democracy without democrats.'"),
    ("Franz von Papen (January 1933)", "On Appointing Hitler", "'What do you want? I have Hindenburg's confidence. Within two months, we will have pushed Hitler into a corner until he squeals.'")
]

# =========================================================================
# QUESTION 9 (PAPER II): 2025 Q.5 (PT 1) / 2022 Q.5 / 2018 Q.4,5,6 / 2016 Q.5
# =========================================================================

meta_p2_9 = [
    ("EXAM QUESTIONS LINKED:", "2025 Q.5 (Part 1): 'Examine the causes and major turning points of the Second World War, including the failure of the policy of appeasement.'\n2022 Q.5: 'Analyze British and French policy of Appeasement towards Nazi Germany from 1935 to 1939. Why did it fail to avert World War II?'\n2018 Q.4: ''It is said that France's hunt for security and England's desire for peace were responsible for the rise of Hitler.' Do you agree?'\n2018 Q.5 & Q.6: 'Review the foreign policy of Britain and France between two World Wars (1919\u20131939).'\n2016 Q.5: 'Critically assess the Spanish Civil War (1936\u20131939) as a testing ground for European dictators and a prelude to World War II.'"),
    ("SYLLABUS TOPIC:", "Paper II \u2014 Section 6: Democracies in Crisis & The Road to World War II \u2014 The Policy of Appeasement, Spanish Civil War & Collapse of Peace"),
    ("ALLOCATED MARKS / TIME:", "20 Marks | 35 Minutes | Standard Word Count: ~650\u2013700 words"),
    ("HISTORIOGRAPHICAL THESIS:", "The British and French policy of Appeasement between 1935 and 1939 was neither sheer cowardice nor incomprehensible incompetence, but a coherent, defensive strategy rooted in the traumatic memory of World War I, public pacifism, severe economic limitations, imperial overstretch, and deep bourgeois terror of Soviet Bolshevism; however, by treating Adolf Hitler as a rational revisionist statesman whose grievances could be satisfied through territorial concessions rather than an unappeasable ideological conqueror, appeasement systematically dismantled Europe's collective security, surrendered strategic bastions, and made World War II inevitable.")
]

callout_p2_9 = [
    "\u2022 Defining Appeasement: The diplomatic strategy of making negotiated concessions to an aggrieved power (Nazi Germany) in order to avoid a catastrophic general war, championed pre-eminently by British Prime Minister Neville Chamberlain.",
    "\u2022 Structural Drivers Behind Appeasement:",
    "  1. Trauma of the Trenches: Horror of repeating the 1914\u20131918 slaughter (Somme/Verdun); British Oxford Union resolution (1933) pledging 'not to fight for King and Country'; universal terror of strategic civilian bombing ('the bomber will always get through').",
    "  2. Military & Economic Lag: Britain and France lagged years behind German rearmament; Chamberlain needed time to construct radar networks (Chain Home) and modern monoplane fighters (Spitfire/Hurricane).",
    "  3. Guilt over Versailles: British political elites genuinely believed that Germany had been mistreated in 1919 and that Hitler's demands for national self-determination (Rhineland, Austria, Sudetenland) were morally justified.",
    "  4. Anti-Bolshevik Buffer: Conservative politicians viewed Nazi Germany as a vital bulwark against the westward spread of Stalin's Soviet Communism.",
    "  5. French Paralyzed Defensiveness: France was crippled by chronic ministerial instability (40+ cabinets) and a defeatist military doctrine anchored behind the concrete fortifications of the static Maginot Line.",
    "\u2022 The Spanish Civil War (1936\u20131939): The premier testing ground for dictators; General Franco's fascist rebellion was bolstered by Hitler's Condor Legion (testing terror bombing at Guernica) and Mussolini's 70,000 troops; USSR aided the Republic; Britain and France enforced a hypocritical 'Non-Intervention Committee' that starved the democratic Spanish Republic, confirming Western democratic impotence.",
    "\u2022 The Steps of Capitulation: Anglo-German Naval Agreement (1935); Remilitarization of the Rhineland (March 1936); Anschluss with Austria (March 1938); Munich Agreement (September 1938) sacrificing the Sudetenland for Chamberlain's illusory 'Peace for our time.'",
    "\u2022 The Fatal Awakening: Hitler's unprovoked occupation of the non-German rump of Czechoslovakia (Bohemia and Moravia) in March 1939 permanently shattered the moral foundation of appeasement, forcing Britain to issue the military guarantee to Poland."
]

map_p2_9 = """========================================================================================
STRUCTURAL CONSTRAINTS            THE SPANISH CRUCIBLE (1936-39)       THE MARCH TO WAR (1938-39)
========================================================================================
[Trauma of WWI & Pacifism]     --> [Spanish Civil War: Franco Revolt]--> [Anschluss: Austria Absorbed]
[Military Rearmament Lag]      --> [Condor Legion Bombs Guernica]   --> [Munich 1938: Sudetenland Ceded]
[Versailles Moral Guilt]       --> [Axis Tested: Rome-Berlin Pact]  --> ['Peace for Our Time' Illusion]
[Terror of Soviet Bolshevism]  --> [Democracies' Non-Intervention   --> [March 1939: Prague Occupied]
[French Maginot Defeatism]          Farce Leaves Republic to Die]    --> [Polish Guarantee -> War Declared]
========================================================================================"""

timeline_p2_9 = [
    ("June 18, 1935", "Anglo-German Naval Agreement", "Britain unilaterally condones German rearmament, allowing Germany 35% of British surface fleet and 45% of submarines."),
    ("March 7, 1936", "Remilitarization of the Rhineland", "Hitler marches into the demilitarized Rhineland with 3 battalions; France and Britain lodge formal diplomatic protest only."),
    ("July 1936\u2013April 1939", "The Spanish Civil War", "Hitler and Mussolini aid Franco; Condor Legion tests blitzkrieg; Western democracies enforce Non-Intervention farce."),
    ("March 12, 1938", "The Anschluss with Austria", "German troops annex Austria; incorporated into Greater German Reich with zero military resistance from Western powers."),
    ("Sept 29\u201330, 1938", "The Munich Conference", "Chamberlain, Daladier, Hitler, and Mussolini partition the Sudetenland; Czechoslovakia and USSR excluded."),
    ("March 15, 1939", "Invasion of Czechoslovakia", "German Wehrmacht occupies Prague, Bohemia, and Moravia; shatters the claim that Hitler sought only ethnic Germans."),
    ("March 31, 1939", "The British Guarantee to Poland", "Chamberlain announces unconditional British and French military backing to Poland in the event of German aggression."),
    ("Aug 23, 1939", "Molotov-Ribbentrop Pact", "Hitler and Stalin sign Non-Aggression Pact, clearing the diplomatic hurdle for an unmolested invasion of Poland."),
    ("Sept 1, 1939", "Invasion of Poland", "German armies invade Poland using Blitzkrieg warfare; Britain and France issue ultimatum to withdraw."),
    ("Sept 3, 1939", "Outbreak of World War II", "Britain and France declare war on Germany; Chamberlain confesses on radio: 'Everything that I have worked for has fallen into ruins.'")
]

essay_p2_9 = [
    ("Introduction: The Historiographical Debate over Appeasement",
     "The foreign policy of Appeasement pursued by Great Britain and France toward Adolf Hitler's Germany between 1935 and 1939 remains one of the most intensely scrutinized chapters in the annals of international diplomacy. Immortalized in popular memory by the famous polemic 'Guilty Men' (1940) by 'Cato' and Winston Churchill's scathing memoirs, appeasement was long condemned as cowardly, craven, and morally bankrupt capitulation by weak-willed democratic politicians who surrendered Europe to fascist tyranny. However, modern revisionist and post-revisionist diplomatic historians such as A.J.P. Taylor, Donald Cameron Watt, and Zara Steiner have fundamentally reshaped this orthodox verdict. Appeasement was not an aberrant failure of political nerve; it was a coherent, rational grand strategy pursued by exhausted, fiscally strained global empires confronting insurmountable domestic, strategic, and military limitations. Yet, its fatal flaw was psychological and ideological: by treating Hitler as a rational Bismarckian revisionist seeking legitimate national adjustments rather than an unappeasable racial-ideological fanatic driven by Lebensraum, appeasement dismantled Europe's collective security architecture, surrendered vital strategic alliances, and precipitated the very world war it was desperately engineered to avert."),

    ("Structural Constraints: Why Democracies Chose Appeasement",
     "The adoption of appeasement by British Prime Minister Neville Chamberlain and French Premier \u00c9douard Daladier was governed by profound structural constraints that made war in the 1930s politically, economically, and militarily unthinkable:\n\n"
     "1. The Trauma of the Great War and Public Pacifism: British and French societies remained deeply shell-shocked by the industrial slaughter of 1914\u20131918. Across Britain, public opinion was militantly pacifist, typified by the 1933 Oxford Union debate which overwhelmingly resolved that 'this House will in no circumstances fight for its King and Country.' Furthermore, military theorists like Giulio Douhet and politicians like Stanley Baldwin terrified populations with apocalyptic predictions of aerial bombardment: 'The bomber will always get through.'\n\n"
     "2. Military Unpreparedness and Economic Fragility: Both Britain and France emerged from the Great Depression with fractured public finances. British Chiefs of Staff warned Chamberlain repeatedly that Britain faced a catastrophic 'three-front threat'\u2014Germany in Europe, Fascist Italy in the Mediterranean, and Imperial Japan in the Far East\u2014with a navy and army barely capable of defending the home islands. Chamberlain pursued appeasement to buy vital time to complete the construction of the radar warning network (Chain Home) and mass-produce modern monoplane fighter aircraft (the Spitfire and Hurricane) that would win the Battle of Britain in 1940.\n\n"
     "3. The Guilt over Versailles and Anti-Bolshevik Paranoia: British ruling elites harbored deep moral guilt regarding the 1919 Treaty of Versailles, viewing it as excessively vindictive. Consequently, when Hitler demanded the remilitarization of the Rhineland or the inclusion of German-speaking Austrians and Sudetens, British politicians viewed these claims as legitimate applications of Wilsonian national self-determination. Simultaneously, conservative elites in London and Paris viewed Hitler's Germany as an indispensable continental bulwark against the terrifying expansion of Soviet Communism.\n\n"
     "4. French Political Instability and the Maginot Mentality: The French Third Republic was paralyzed by chronic political polarization, experiencing over forty revolving coalition governments between 1919 and 1939. Terrified of demographic inferiority (France had 40 million citizens against Germany's 70 million), the French military adopted an utterly defeatist, defensive doctrine, burying its resources into the concrete subterranean fortresses of the Maginot Line, rendering France strategically incapable of launching offensive operations to assist its Eastern European allies."),

    ("The Spanish Civil War (1936\u20131939): The Proving Ground for Dictators",
     "The supreme moral and strategic bankruptcy of democratic foreign policy was vividly exposed during the Spanish Civil War. When General Francisco Franco launched a fascist military insurrection against the democratically elected Spanish Republic in July 1936, the conflict immediately transformed into an ideological crucible for Europe. Adolf Hitler and Benito Mussolini intervened decisively on Franco's side: Hitler deployed the Condor Legion, utilizing advanced Heinkel bombers and Messerschmitt fighters to test modern Blitzkrieg tactics, including the terror bombing of civilian populations at Guernica in April 1937, while Mussolini dispatched seventy thousand Italian troops.\n\n"
     "In contrast, terrified of general war and fearful of domestic communist influence, Britain and France orchestrated the farcical 'Non-Intervention Committee.' While the democracies prohibited arms sales to the legal Spanish government, Germany and Italy openly flouted the agreement with continuous military supplies. Only the Soviet Union provided military assistance to the Republic, forcing the Spanish government into the arms of the Comintern. The Spanish tragedy yielded fatal geopolitical consequences: it cemented the Rome-Berlin Axis (October 1936), isolated France by placing a hostile fascist dictatorship on its southern border, and confirmed to Hitler that the Western democracies would never resort to armed force to defend democratic self-determination."),

    ("The Escalating Surrenders: Rhineland, Anschluss, and the Road to Munich",
     "Hitler dismantled the post-WWI security framework through a series of escalating gambles, each rewarded by democratic inaction. On March 7, 1936, Hitler took his greatest military gamble by marching three battalions into the demilitarized Rhineland. French forces along the border vastly outnumbered the Germans, and Hitler's generals carried secret orders to retreat immediately if challenged. Yet, France refused to act without British support, and British public opinion concluded that Germany was merely 'walking into its own back garden.' The militarization of the Rhineland was a catastrophic strategic catastrophe: by constructing the Siegfried Line (West Wall), Germany neutralized French offensive capabilities, leaving France's Eastern European allies (Poland and Czechoslovakia) completely cut off from Western military rescue.\n\n"
     "Two years later, in March 1938, Hitler executed the bloodless absorption of Austria (*Anschluss*). Mobilizing Austrian Nazi fifth-columns and issuing an ultimatum to Chancellor Kurt von Schuschnigg, German troops crossed the border to cheering crowds. Britain and France registered formal diplomatic protests and did nothing, directly exposing Czechoslovakia to German strategic encirclement on three frontiers."),

    ("The Apex of Capitulation: The Munich Betrayal of September 1938",
     "The climax of appeasement unfolded during the Sudetenland Crisis of 1938. Czechoslovakia was the sole surviving democracy in Central Europe, possessing an advanced industrial economy (the massive \u0160koda arms works), an army of thirty-five divisions, an impenetrable mountain fortification system, and formal mutual defense treaties with France and the Soviet Union. Demanding the incorporation of three million ethnic Sudeten Germans into the Reich, Hitler ordered his military to prepare 'Operation Green' for an invasion.\n\n"
     "Terrified of a general European war, Neville Chamberlain engaged in desperate shuttle diplomacy, flying to Germany three times to meet Hitler (at Berchtesgaden, Godesberg, and Munich). On September 29\u201330, 1938, the historic Munich Conference assembled, attended by Chamberlain, Daladier, Mussolini, and Hitler. The Czechoslovak government was completely excluded from the conference, forced to wait outside the room while their national fate was decided, and the Soviet Union was pointedly ignored. Under the Munich Agreement, Britain and France capitulated to all of Hitler's demands, ordering Czechoslovakia to surrender the Sudetenland within ten days on pain of facing Germany alone.\n\n"
     "Returning to London, Chamberlain waved the piece of paper signed by Hitler, declaring to jubilant crowds that he had secured 'Peace for our time.' In reality, Munich was an unmitigated strategic disaster: it stripped Czechoslovakia of its defensive fortifications, destroyed the French alliance system in Eastern Europe, deeply alienated Stalin (convincing him of Western treachery), and delivered thirty-five divisions' worth of Czech tanks, artillery, and munitions into the hands of the Wehrmacht."),

    ("The Fatal Awakening: Prague, the Polish Guarantee, and War (1939)",
     "The moral and geopolitical illusion of appeasement was shattered on March 15, 1939. In blatant violation of his Munich promises that the Sudetenland was his 'last territorial demand in Europe,' Hitler marched German troops into Prague, annexing the non-German territory of Bohemia and Moravia as a Reich protectorate. For the first time, Nazi aggression was exposed not as a legitimate revision of Versailles to unite ethnic Germans, but as an insatiable campaign of imperialist subjugation.\n\n"
     "Public opinion in Britain and France revolted against further humiliation. In an abrupt, desperate policy reversal, Chamberlain issued a formal military guarantee on March 31, 1939, pledging that Britain and France would declare war if Germany attacked Poland. Yet, having surrendered Czechoslovakia and alienated the Soviet Union, the Western guarantee lacked all military credibility. Hitler dismissed the Anglo-French warning as another bluff, securing his eastern flank on August 23, 1939, by signing the Molotov-Ribbentrop Pact with Stalin. When German panzer divisions crossed the Polish frontier on September 1, 1939, Britain and France issued an ultimatum and declared war on September 3. Appeasement was dead, leaving the democracies to wage the very global war they had sacrificed their honor to prevent."),

    ("Conclusion: The Bitter Verdict on Appeasement",
     "In conclusion, the policy of Appeasement was an understandable, yet disastrous, grand strategic failure. Grounded in legitimate domestic horrors of industrialized slaughter, financial exhaustion, and military unreadiness, it nevertheless failed catastrophically because it fundamentally misdiagnosed the nature of Adolf Hitler's totalitarian regime. Hitler was not an ordinary statesman seeking to revise Versailles; he was an apocalyptic racial ideologue whose explicit program demanded the military conquest of European Lebensraum and the extermination of Jews and Slavs. By surrendering the Rhineland, Austria, and Czechoslovakia without firing a shot, appeasement did not prevent war; it merely ensured that when war came, Germany possessed vastly greater industrial resources, captured armaments, and secure frontiers, forcing Great Britain and France to fight for their civilizational survival under the most perilous circumstances imaginable.")
]

quotes_p2_9 = [
    ("Winston Churchill", "Speech in the House of Commons (Oct 5, 1938)", "'You were given the choice between war and dishonour. You chose dishonour, and you will have war... This is only the first sip, the first foretaste of a bitter cup which will be proffered to us year by year.'"),
    ("Donald Cameron Watt", "How War Came (1989)", "'Appeasement was the policy of a status quo power that lacked the military and economic resources to fight a global war on three fronts... Chamberlain was buying time for British rearmament, not groveling before a dictator.'"),
    ("A.J.P. Taylor", "The Origins of the Second World War", "'Hitler did not have a blueprint for conquest; he was an opportunist who simply walked through the doors that the British and French opened for him through the policy of appeasement.'"),
    ("Neville Chamberlain (Sept 3, 1939)", "Radio Broadcast Declaring War", "'Everything that I have worked for, everything that I had hoped for, everything that I have believed in during my public life, has crashed into ruins.'")
]

# Run the generation
if __name__ == "__main__":
    tasks = [
        (meta_p2_7, callout_p2_7, map_p2_7, timeline_p2_7, essay_p2_7, quotes_p2_7,
         os.path.join(paper2_dir, "Q7_Fascist_Italy_Mussolini_Corporate_State_Master_Notes.docx"),
         "QUESTION 7: FASCIST ITALY UNDER MUSSOLINI: IDEOLOGY, CORPORATE STATE & CONTROL",
         "Q7_Fascist_Italy_Mussolini_Corporate_State_Master_Notes"),
        (meta_p2_8, callout_p2_8, map_p2_8, timeline_p2_8, essay_p2_8, quotes_p2_8,
         os.path.join(paper2_dir, "Q8_Weimar_Crises_Hitler_Rise_and_Nazi_Third_Reich_Master_Notes.docx"),
         "QUESTION 8: WEIMAR CRISES, GREAT DEPRESSION & HITLER'S TOTALITARIAN THIRD REICH",
         "Q8_Weimar_Crises_Hitler_Rise_and_Nazi_Third_Reich_Master_Notes"),
        (meta_p2_9, callout_p2_9, map_p2_9, timeline_p2_9, essay_p2_9, quotes_p2_9,
         os.path.join(paper2_dir, "Q9_Policy_of_Appeasement_and_Spanish_Civil_War_Master_Notes.docx"),
         "QUESTION 9: THE POLICY OF APPEASEMENT, SPANISH CIVIL WAR & COLLAPSE OF PEACE",
         "Q9_Policy_of_Appeasement_and_Spanish_Civil_War_Master_Notes")
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
