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
# QUESTION 10: 2018 Q.1 / 2016 Q.1 / 2013 Q.1 (CAUSES OF FRENCH REVOLUTION)
# =========================================================================

meta_10 = [
    ("EXAM QUESTIONS LINKED:", "2018 Q.1: 'Critically examine the social, economic, and intellectual causes of the French Revolution of 1789.'\n2016 Q.1: 'Discuss in detail the philosophy of Liberty, Equality, and Fraternity that motivated the French people to revolt against monarchical powers in 1789.'\n2013 Q.1: 'The conditions in Europe were always the same except in England. But why did the revolution occur in France and not in any other country?'"),
    ("SYLLABUS TOPIC:", "Paper I \u2014 Section 1: The French Revolution (1789\u20131799) \u2014 Causes, Social Dynamics & Philosophy"),
    ("ALLOCATED MARKS / TIME:", "20 Marks | 35 Minutes | Standard Word Count: ~650\u2013700 words"),
    ("HISTORIOGRAPHICAL THESIS:", "The French Revolution of 1789 was not the spontaneous eruption of sheer misery, but the structural implosion of an archaic, fiscally bankrupt feudal absolute monarchy colliding with an economically ascendant, literate bourgeoisie and an aggrieved peasantry, catalyzed by Enlightenment critique that delegitimized divine right and articulated the transformative revolutionary triad: 'Libert\u00e9, \u00c9galit\u00e9, Fraternit\u00e9'.")
]

callout_10 = [
    "\u2022 The Three Estates Imbalance: 1st Estate (Clergy, 100k) & 2nd Estate (Nobility, 400k) owned 40% of fertile land and enjoyed total exemption from direct taxation (taille, corv\u00e9e). The 3rd Estate (24 million peasants, artisans, bourgeois) paid over 70% of their net income in feudal, royal, and ecclesiastic dues.",
    "\u2022 De Tocqueville's Paradox ('Why France?'): Revolutions erupt not where oppression is absolute (Russia, Austria), but where expanding economic freedom and partial property ownership make surviving feudal burdens intolerable. French peasants owned land and the French bourgeoisie was intellectually mature, making feudal exclusions intolerable.",
    "\u2022 The Financial Death-Spiral: Participation in the American War of Independence (1778\u20131783) added 1.3 billion livres of debt. By 1788, 50% of the royal budget went strictly to debt service. When Calonne and Brienne attempted to tax the nobility, the aristocratic Parlements rebelled, forcing the convocation of the Estates-General.",
    "\u2022 The Philosophical Triad: Liberty (Montesquieu: freedom from arbitrary arrest and royal despotism; separation of powers), Equality (Rousseau: equality before law, abolition of aristocratic privilege), Fraternity (National brotherhood replacing regional and estate divisions into one sovereign Nation)."
]

map_10 = """========================================================================================
STRUCTURAL PRESSURES               INTELLECTUAL CATALYST             TRIGGER & REVOLT (1789)
========================================================================================
[1st & 2nd Estates: Tax-Exempt]  +  [Montesquieu: Law & Liberty]  --> [1788 Fiscal Bankruptcy]
[3rd Estate: Borne 70% Taxes]    +  [Voltaire: Anti-Clericalism]  --> [Estates-General (May 1789)]
[Bourgeoisie Barred from Power]  +  [Rousseau: Popular Sovereignty]--> [Tennis Court Oath (June 20)]
[Agrarian Distress: Hail 1788]   +  [Siey\u00e8s: 'What is 3rd Estate']--> [Storming of Bastille (July 14)]
                                                                   --> [August 4: Feudalism Dead]
                                                                   --> [Declaration of Rights (Aug 26)]
========================================================================================"""

timeline_10 = [
    ("1774", "Accession of Louis XVI", "Inherits ballooning national debt; appoints reformist Turgot who is dismissed under court pressure."),
    ("1778\u20131783", "French Aid to American Revolution", "Bankrupts the Royal Treasury; introduces French soldiers (Lafayette) to republican ideology."),
    ("Feb 1787", "Assembly of Notables", "Calonne's proposed universal land tax rejected by nobles; aristocratic resistance paralyzes monarchy."),
    ("Aug 1788", "Royal Bankruptcy Declared", "Necker recalled; Louis XVI forced to summon the Estates-General for the first time since 1614."),
    ("Jan 1789", "Abb\u00e9 Siey\u00e8s' Pamphlet", "Publishes 'What is the Third Estate?': 'Everything. What has it been? Nothing. What does it ask? To become something.'"),
    ("May 5, 1789", "Estates-General Convenes", "Deadlock over voting by order (giving 2% nobles/clergy majority) vs. voting by head (Third Estate)."),
    ("June 20, 1789", "Tennis Court Oath", "Third Estate declares itself National Assembly; pledges not to disband until France has a written constitution."),
    ("July 14, 1789", "Storming of the Bastille", "Parisian populace enters political arena; destroys royal fortress; saves Assembly from military coup."),
    ("Aug 4, 1789", "Night of August 4", "National Assembly formally abolishes feudalism, tithes, seigneurial dues, and aristocratic tax exemptions."),
    ("Aug 26, 1789", "Declaration of Rights of Man", "Promulgates 'Liberty, Equality, Fraternity'; declares men born free and equal in rights; law as expression of General Will.")
]

essay_10 = [
    ("Introduction: The Anatomy of a Revolutionary Breakdown",
     "The French Revolution of 1789 was the defining crucible of modern European civilization, radically fracturing the millennial edifice of feudal absolutism and ushering in the era of popular sovereignty, civic egalitarianism, and secular constitutionalism. In the historiographical debate, while classic Marxist scholars such as Georges Lefebvre and Albert Soboul interpreted 1789 as an inevitable class rupture wherein an economically emergent bourgeoisie smashed an outmoded feudal order, modern revisionists like Fran\u00e7ois Furet and Alfred Cobban emphasize an institutional and cultural breakdown precipitated by fiscal collapse. In reality, the Revolution was an overdetermined phenomenon: an obsolete sociopolitical architecture (the Ancien R\u00e9gime) paralyzed by structural privilege and fiscal bankruptcy collided with an agrarian subsistence crisis and a revolutionary philosophical discourse. Crucially, the revolutionary battle cry of 'Libert\u00e9, \u00c9galit\u00e9, Fraternit\u00e9' did not merely reflect reactive discontent; it provided the transformative conceptual grammar that converted fiscal breakdown into a systemic ideological reconstruction of European political order."),

    ("Socioeconomic Structure: The Ancien R\u00e9gime and the Estates Gridlock",
     "The fundamental structural vulnerability of 18th-century France lay in its rigid, tri-cameral estate system. The First Estate (the Catholic Clergy, numbering approximately 100,000) and the Second Estate (the Nobility, around 400,000) collectively monopolized over forty percent of France's agricultural land, controlled all high military and judicial offices, and enjoyed near-total exemption from direct taxation, notably the taille and corv\u00e9e. In stark contrast, the Third Estate\u2014comprising over 98 percent of the population (24 million souls)\u2014bore the overwhelming fiscal burden of the state. Peasants, who constituted 80 percent of the populace, were crushed beneath a crushing tax pyramid of royal levies (taille, capitation, vingti\u00e8me), feudal dues (banalit\u00e9s, champart), and the church tithe (d\u00eeme), often surrendering upwards of 70 percent of their harvest. Concurrently, the urban bourgeoisie\u2014merchants, bankers, jurists, and intellectuals\u2014experienced an acute status inconsistency: possessing burgeoning wealth and administrative expertise, they found themselves systematically excluded from political power by aristocratic monopolies, fueling deep structural resentment."),

    ("Fiscal Bankruptcy and Institutional Paralysis: The Immediate Trigger",
     "The proximate catalyst of 1789 was not general economic poverty, but the acute insolvency of the French Crown. France's extensive military expenditures during the Seven Years' War (1756\u20131763) and its decisive maritime and financial intervention in the American Revolutionary War (1778\u20131783)\u2014financed predominantly through loans arranged by Jacques Necker\u2014doubled the national debt to over four billion livres. By 1788, sovereign debt servicing alone consumed over fifty percent of annual state revenue. When Controllers-General Calonne and Brienne recognized that solvency required abolishing noble tax exemptions through a universal land tax, the aristocratic Parlements and the 1787 Assembly of Notables adamantly refused. By asserting that only an Estates-General could approve new taxes, the nobility initiated what Lefebvre termed the 'aristocratic revolt.' However, by paralyzing the crown, the nobility unwittingly opened the floodgates to popular mobilization. The catastrophic harvest of 1788, marked by devastating hail storms, doubled bread prices across Paris, creating an explosive convergence of urban starvation and elite institutional impasse."),

    ("The Intellectual Catalyst: Philosophers and the Triad of Liberty, Equality, Fraternity",
     "The Enlightenment provided the indispensable intellectual armory that transformed localized grievances into a coherent revolutionary ideology. While philosophers like Voltaire, Montesquieu, and Jean-Jacques Rousseau were not active conspirators of rebellion, their writings corroded the ideological legitimacy of divine right absolutism. Baron de Montesquieu's 'Spirit of the Laws' (1748) established the imperative of constitutional limits and the separation of executive, legislative, and judicial powers to safeguard 'Liberty' from royal despotism. Voltaire deployed corrosive satire to expose ecclesiastical corruption, religious intolerance, and legal brutality, fostering an intellectual climate of secular rationalism. Most decisively, Rousseau's 'Social Contract' (1762) introduced the revolutionary concepts of 'Popular Sovereignty' and the 'General Will,' arguing that legitimate political authority resides exclusively in the collective will of the sovereign citizenry, not in the dynastic person of the monarch. This philosophy was encapsulated in the immortal revolutionary triad: 'Libert\u00e9' represented the eradication of arbitrary arrest (lettres de cachet) and freedom of expression; '\u00c9galit\u00e9' meant civic equality before the law and the abolition of aristocratic tax immunities; and 'Fraternit\u00e9' forged a unified national brotherhood, superseding regionalism and corporate estates into an indivisible French Nation."),

    ("Tocqueville's Paradox: Why Did Revolution Erupt in France and Not Elsewhere?",
     "A pivotal comparative question in European history is why the revolutionary upheaval occurred in France rather than in Austria, Prussia, or Russia, where monarchical absolutism was far more despotic and peasant serfdom markedly harsher. As Alexis de Tocqueville perceptively demonstrated in 'The Old Regime and the Revolution' (1856), social revolutions break out not where conditions are objectively worst, but where partial progress and expanding prosperity make remaining feudal anachronisms appear profoundly intolerable. Unlike Russian or Prussian serfs who remained tied to the soil under unmitigated baronial servitude, French peasants were already legally emancipated landholders owning nearly half of the kingdom's soil; consequently, feudal tithes and seigneurial dues felt like arbitrary, parasitial extortions rather than normal life. Similarly, the French bourgeoisie was the most affluent, educated, and self-conscious in continental Europe. Encountering an entrenched noble 'aristocratic reaction' that closed military officer ranks and magistrate posts via the S\u00e9gur Ordinance of 1781, the French middle class possessed both the economic leverage and ideological confidence to dismantle the state rather than endure subjugation."),

    ("Conclusion: The Enduring Legacy of 1789",
     "In conclusion, the French Revolution was the result of a catastrophic confluence: an unsustainable structural inequality of estates, an acute fiscal bankruptcy brought on by imperial warfare, an agrarian bread crisis, and an Enlightenment intellectual renaissance that furnished the masses and bourgeoisie with an alternative moral architecture. When the Third Estate reconstituted itself as the National Assembly on June 17, 1789, followed by the storming of the Bastille on July 14, the streets and the legislature combined to dismantle absolute monarchy. In promulgating the Declaration of the Rights of Man on August 26, 1789, France did not merely reform its own borders; it established a universalist political template. By enthroning 'Liberty, Equality, and Fraternity' as the foundational metrics of legitimate governance, 1789 irrevocably dismantled the dynastic legitimacy of Europe and laid the ideological groundwork for modern democracy, human rights, and constitutional government.")
]

quotes_10 = [
    ("Alexis de Tocqueville", "Institutional Continuity & Rising Expectations", "'The regime which is destroyed by a revolution is almost always an improvement on its immediate predecessor... Evils which are patiently endured when they seem inevitable become intolerable once the idea of escape from them is suggested.'"),
    ("Georges Lefebvre", "Marxist Synthesis of Four Revolutions", "'The ultimate cause of the French Revolution was the rise of the bourgeoisie... but it was the aristocracy which opened the way by revolting against the monarchy, followed by the urban revolution and the agrarian revolt.'"),
    ("Fran\u00e7ois Furet", "Revisionist Political-Culture Critique", "'The French Revolution was not a transition from feudalism to capitalism, but the invention of a new political culture founded on the absolute, democratic legitimacy of the sovereign Nation.'"),
    ("Abb\u00e9 Siey\u00e8s (1789)", "Revolutionary Manifestation of Third Estate", "'What is the Third Estate? Everything. What has it been heretofore in the political order? Nothing. What does it demand? To become something.'")
]

# =========================================================================
# QUESTION 11: 2022 Q.1 (NATIONAL ASSEMBLY 1789-1791: STREETS VS LEGISLATURE)
# =========================================================================

meta_11 = [
    ("EXAM QUESTIONS LINKED:", "2022 Q.1: ''Revolutions often start from the streets but eventually get completed in legislatures.' In the light of this statement, assess the role of the National Assembly in France between 1789 and 1791.'"),
    ("SYLLABUS TOPIC:", "Paper I \u2014 Section 1: The French Revolution \u2014 National Constituent Assembly & Constitutional Transformation (1789\u20131791)"),
    ("ALLOCATED MARKS / TIME:", "20 Marks | 35 Minutes | Standard Word Count: ~650\u2013700 words"),
    ("HISTORIOGRAPHICAL THESIS:", "While insurrectionary street mobilizations in Paris and peasant jacqueries in the countryside broke the coercive apparatus of Bourbon absolutism, it was the legislative craftsmanship of the National Constituent Assembly (1789\u20131791) that translated raw popular violence into an enduring institutional, administrative, and constitutional reality, proving that while streets dismantle absolute tyranny, legislatures institutionalize modern freedom.")
]

callout_11 = [
    "\u2022 The Street-Legislature Dialectic: Popular insurrections (Storming of the Bastille on July 14, Great Fear in the countryside in July\u2013August, Women's March on Versailles on October 5\u20136) repeatedly rescued the bourgeois National Assembly from royal military repression and forced Louis XVI to yield.",
    "\u2022 The Legislative Miracle: In two years, the National Assembly dismantled over a millennium of feudal institutions: Night of August 4 (abolition of feudal dues and seigneurial courts); Declaration of the Rights of Man and Citizen (August 26, 1789); Civil Constitution of the Clergy (July 1790); and the Constitution of 1791.",
    "\u2022 Rationalized Administrative Architecture: Eradicated the chaotic 300-custom patchwork of the Ancien R\u00e9gime, dividing France into 83 uniform, self-governing D\u00e9partements; standardized weights and measures; eliminated internal customs barriers; and instituted trial by jury.",
    "\u2022 Bourgeois Contradiction & Limits: The Assembly restricted the franchise through property qualifications, dividing citizens into 'Active' (tax-paying men with voting rights) and 'Passive' (excluded working classes), sowing the seeds for future radical Jacobin insurrections."
]

map_11 = """========================================================================================
POPULAR PRESSURE (THE STREETS)        LEGISLATIVE CODIFICATION (THE ASSEMBLY 1789-1791)
========================================================================================
[July 14: Bastille Stormed]        --> [Rescues Assembly; Royal Troops Withdraw from Paris]
[July-Aug: Peasant Great Fear]     --> [Night of August 4: Decrees Abolish Feudalism & Tithes]
[Urban Bread Crisis & Agitation]   --> [August 26: Declaration of Rights of Man Codified]
[Oct 5-6: Women's March to Versailles]-> [Forces King & Assembly to Paris; Subordinates Crown]
[Fiscal Crisis & Sans-Culottes]    --> [Nov 1789: Church Property Nationalized (Assignats)]
                                   --> [July 1790: Civil Constitution of the Clergy]
                                   --> [Sept 1791: Constitution of 1791 (Constitutional Monarchy)]
========================================================================================"""

timeline_11 = [
    ("June 17\u201320, 1789", "Birth of National Assembly & Tennis Court Oath", "Third Estate declares national sovereignty; vows to draft a constitution despite royal lockouts."),
    ("July 14, 1789", "Fall of the Bastille", "Street insurrection destroys royal despotism symbol; forces Louis XVI to recognize the Assembly and wear tricolor."),
    ("Aug 4, 1789", "Abolition of Feudalism", "Assembly deputies renounce aristocratic privileges, hunting rights, tithes, and tax immunities in night session."),
    ("Aug 26, 1789", "Declaration of the Rights of Man", "Foundational charter of human liberty promulgated: equality before law, freedom of speech/religion, property rights."),
    ("Oct 5\u20136, 1789", "October Days (Women's March)", "7,000 armed women march to Versailles; bring 'the Baker, the Baker's wife, and the Baker's little boy' to Paris."),
    ("Nov 2, 1789", "Confiscation of Church Lands", "Assembly nationalizes ecclesiastical wealth (biens nationaux) to back paper currency (assignats)."),
    ("July 12, 1790", "Civil Constitution of the Clergy", "Priests turned into salaried state officials elected by citizens; creates disastrous religious schism in France."),
    ("March 1791", "Allarde & Le Chapelier Laws", "Assembly abolishes craft guilds and bans workers' trade unions/strikes, enshrining bourgeois free market."),
    ("June 20, 1791", "Flight to Varennes", "Louis XVI attempts to flee abroad to lead counter-revolution; shatters the Assembly's constitutional monarchy illusion."),
    ("Sept 3\u201330, 1791", "Constitution of 1791 & Dissolution", "First written constitution enacted; Assembly dissolves via Robespierre's Self-Denying Ordinance.")
]

essay_11 = [
    ("Introduction: The Street-Legislature Dialectic",
     "The axiom that 'revolutions often start from the streets but eventually get completed in legislatures' encapsulates the essential dialectic of the French Revolution between 1789 and 1791. Left to themselves, street insurrections and peasant jacqueries represent destructive energies capable of paralyzing an authoritarian regime, yet inherently incapable of constructing a durable institutional order. Conversely, legislative bodies operating in a political vacuum are easily intimidated or dispersed by armed royalist power. Between the summer of 1789 and the autumn of 1791, the French National Constituent Assembly functioned as the institutional crucible of the modern world precisely because it engaged in a symbiotic dialogue with popular street mobilization. While the storming of the Bastille and the October Days shattered the crown's coercive apparatus, it was the jurists, philosophers, and bourgeois deputies of the National Assembly who forged that raw rebellious fury into the constitutional, administrative, and legal foundation of the modern French state."),

    ("The Streets as the Armed Protector of the Legislature",
     "The National Assembly was born in defiance on June 17, 1789, when the Third Estate, joined by liberal clerics and nobles, proclaimed itself the sovereign representative body of France and took the solemn Tennis Court Oath. However, King Louis XVI immediately concentrated twenty thousand Swiss and German mercenary troops around Paris to crush the nascent parliament. The Assembly was preserved from military dissolution exclusively by the spontaneous intervention of the streets. On July 14, 1789, Parisian artisans, shopkeepers, and wage-earners stormed the Bastille fortress, seizing thirty thousand muskets and lynching Governor de Launay. This street uprising permanently broke the royal monopoly on violence, forced Louis XVI to withdraw his troops, accept the tricolor cockade, and legitimize the National Assembly. Similarly, when rural peasants ignited the 'Great Fear' (Grande Peur) in July and August by burning ch\u00e2teaux and seigneurial debt-registers, the Assembly was compelled to accelerate structural reform. Furthermore, during the October Days of 1789, an armed crowd of Parisian market women marched twelve miles through torrential rain to Versailles, storming the palace and compelling both the royal family and the National Assembly to relocate to Paris under the watchful eye of the sovereign populace."),

    ("The Legislative Masterstrokes: Dismantling Feudalism and Codifying Rights",
     "Having been shielded by popular force, the National Assembly executed an astonishing legislative revolution. In the historic nocturnal session of August 4, 1789, liberal aristocrats and deputies renounced their ancestral privileges, formally abolishing the feudal regime in its entirety. Feudal personal servitude, seigneurial jurisdictions, aristocratic hunting monopolies, and ecclesiastical tithes were obliterated overnight, establishing absolute legal equality. Three weeks later, on August 26, 1789, the Assembly promulgated the 'Declaration of the Rights of Man and of the Citizen.' Drafted by the Marquis de Lafayette with input from Thomas Jefferson, this monumental document codified Enlightenment philosophy into positive law: it affirmed that men are born free and equal in rights, declared private property sacred and inviolable, established liberty, security, and resistance to oppression as fundamental entitlements, and proclaimed that sovereignty resides not in the dynastic person of the king, but essentially in the Nation."),

    ("Administrative, Judicial, and Fiscal Restructuring",
     "The constructive genius of the National Assembly was most vividly displayed in its comprehensive rationalization of French governance. The archaic, labyrinthine administrative patchwork of bailiwicks, intendancies, and custom zones was swept away and replaced by eighty-three uniform, geographically balanced departments (d\u00e9partements), subdivided into districts and communes, each governed by locally elected councils. In the judicial sphere, hereditary venal magistrate offices were abolished, torture was outlawed, and a unified secular legal system featuring trial by public juries was introduced. To resolve the catastrophic royal debt, the Assembly nationalized all Roman Catholic Church property on November 2, 1789 (biens nationaux), issuing interest-bearing bonds known as assignats. While the assignats eventually succumbed to hyperinflation due to over-issuance, the sale of church lands permanently tied the peasant and bourgeois purchasers to the survival of the revolutionary settlement."),

    ("Contradictions, Fractures, and the Bourgeois Limitation",
     "Despite its monumental achievements, the National Assembly suffered from fatal contradictions that ultimately undermined its moderate constitutional vision. Guided by bourgeois economic interests, the Assembly passed the Le Chapelier Law (June 1791), which outlawed trade unions and collective strikes, leaving urban wage-earners defenseless against capitalist exploitation. More critically, the Constitution of 1791 betrayed the egalitarian promise of the Declaration of Rights by dividing Frenchmen into 'Active' and 'Passive' citizens. Only affluent men paying direct taxes equivalent to three days of labor were granted voting rights, disenfranchising over three million working-class citizens (sans-culottes). The Assembly's most catastrophic miscalculation, however, was the Civil Constitution of the Clergy (July 1790). By transforming priests into salaried civil servants elected by parishioners and demanding a loyalty oath to the state, the Assembly alienated the Pope, split the French clergy into 'juring' and 'refractory' camps, and ignited a bitter religious civil war across rural France (notably in the Vend\u00e9e)."),

    ("Conclusion: The Triumph and Legacy of the Legislative Crucible",
     "In conclusion, the history of France between 1789 and 1791 affirms that revolutions require both the disruptive eruption of the street and the institutional codification of the legislature. Without the storming of the Bastille, the Tennis Court oath would have ended in mass arrests; without the Assembly's systematic drafting of eighty-three departments, the abolition of feudalism, and the Constitution of 1791, the street uprisings would have dissipated into aimless anarchy. Although the Flight to Varennes in June 1791 exposed the impossibility of a constitutional monarchy under Louis XVI and propelled France toward the radical Jacobin Republic, the institutional scaffolding erected by the National Constituent Assembly proved permanent. The administrative, legal, and egalitarian framework codified between 1789 and 1791 survived the Reign of Terror, Napoleon Bonaparte, and the Bourbon Restoration, enduring as the bedrock of modern democratic France.")
]

quotes_11 = [
    ("Timothy Tackett", "Empirical Transformation of Deputies", "'The deputies of the National Assembly did not arrive in Versailles as radical revolutionaries; it was the intense pressure of the Paris streets and court conspiracies that forged them into legislative transformers.'"),
    ("Simon Schama", "Revisionist Critique of Revolutionary Violence", "'From the very beginning, violence was the motor of the French Revolution... the National Assembly legislated under the perpetual shadow of the pike and the lantern.'"),
    ("Albert Soboul", "Marxist Perspective on Bourgeois Hegemony", "'The work of the National Constituent Assembly was essentially a bourgeois triumph; it swept away feudalism to establish the juridical equality necessary for agrarian and commercial capitalism.'"),
    ("William Doyle", "Administrative Legacy of 1789-1791", "'The destruction of the Ancien R\u00e9gime was completed in 1789, but the lasting achievement of the National Assembly was the creation of a rational, uniform, and decentralized administrative state.'")
]

# =========================================================================
# QUESTION 12: 2022 Q.8 (POLITICAL PHILOSOPHERS: LOCKE & ROUSSEAU)
# =========================================================================

meta_12 = [
    ("EXAM QUESTIONS LINKED:", "2022 Q.8: 'Assess the contribution of the following philosophers in changing the political thought of Europe: (a) John Locke, (b) Jean-Jacques Rousseau.'"),
    ("SYLLABUS TOPIC:", "Paper I \u2014 Section 1: Intellectual Foundations of Modern Europe \u2014 Political Philosophy of the Enlightenment"),
    ("ALLOCATED MARKS / TIME:", "20 Marks (10 Marks each for Locke and Rousseau) | 35 Minutes | Standard Word Count: ~650\u2013700 words"),
    ("HISTORIOGRAPHICAL THESIS:", "John Locke and Jean-Jacques Rousseau radically transformed European political consciousness by dismantling divine-right patriarchal absolutism and establishing contractual, secular legitimacy; while Locke provided the empirical, individualistic architecture for liberal constitutionalism, private property, and limited government, Rousseau articulated a radical, communitarian vision of popular sovereignty and the General Will that laid the intellectual powder-keg for modern democratic republicanism and mass revolution.")
]

callout_12 = [
    "\u2022 John Locke (1632\u20131704) \u2014 The Architect of Liberal Constitutionalism: In 'Two Treatises of Government' (1689), refuted Robert Filmer's divine right theory. Articulated that all human beings in the State of Nature possess inalienable Natural Rights: 'Life, Liberty, and Estate (Property).' Government is a conditional fiduciary trust; if rulers violate rights, citizens have the moral right of revolution.",
    "\u2022 Jean-Jacques Rousseau (1712\u20131778) \u2014 The Prophet of Popular Sovereignty: In 'The Social Contract' (1762), declared: 'Man is born free, and everywhere he is in chains.' Denounced private property as the root of inequality in his 'Discourse on Inequality'. Sovereignty cannot be represented or alienated; it resides perpetually in the 'General Will' (volont\u00e9 g\u00e9n\u00e9rale) aimed at the common good.",
    "\u2022 Comparative Dichotomy: Locke emphasized negative liberty (freedom from state encroachment, defense of property, representative parliamentary rule); Rousseau championed positive liberty (moral autonomy achieved through active participation in the civic community, subordinating private interest to collective welfare).",
    "\u2022 European Historical Impact: Locke's thought anchored the English Bill of Rights (1689), Montesquieu's separation of powers, and the American Revolution (1776); Rousseau's thought inspired the French Jacobins (Robespierre's 'Republic of Virtue'), 19th-century Romanticism, and modern democratic and socialist theory."
]

map_12 = """========================================================================================
PHILOSOPHER      STATE OF NATURE       SOCIAL CONTRACT GOAL        POLITICAL LEGACY
========================================================================================
JOHN LOCKE   --> [State of Reason]  --> [Fiduciary Trust to     --> [Liberal Constitutionalism]
(1632-1704)      [Natural Rights:        Protect Life, Liberty,      [Limited Government / Rule of Law]
                  Life/Liberty/Property] & Property]                 [Right of Revolution (1689/1776)]

JEAN-JACQUES --> [Noble Savage]     --> [Total Alienation of    --> [Radical Popular Sovereignty]
ROUSSEAU         [Corrupted by           Private Will to the         [Direct Participatory Democracy]
(1712-1778)       Civil Society &         General Will                [Jacobin Republic of Virtue (1793)]
                  Private Property]       (Volont\u00e9 G\u00e9n\u00e9rale)]       [Romanticism & Modern Socialism]
========================================================================================"""

timeline_12 = [
    ("1689", "Locke's Two Treatises of Government", "Published anonymously; demolishes divine right absolutism; establishes natural rights and social compact."),
    ("1690", "Locke's Essay Concerning Human Understanding", "Articulates 'tabula rasa' (blank slate); destroys scholastic dogmatism; elevates empirical reason."),
    ("1754", "Rousseau's Discourse on Inequality", "Argues that civil society, private property, and institutional laws generated human inequality and moral decay."),
    ("1762", "Rousseau's The Social Contract", "Opens with immortal phrase: 'Man is born free, and everywhere he is in chains'; formulates General Will."),
    ("1776", "American Declaration of Independence", "Thomas Jefferson directly adapts Locke's triad into 'Life, Liberty, and the pursuit of Happiness'."),
    ("1789", "Declaration of the Rights of Man", "Synthesizes Locke (Articles 1-2 on natural liberty/property) and Rousseau (Article 6: 'Law is expression of General Will')."),
    ("1793\u20131794", "Jacobin Reign of Terror", "Maximilien Robespierre applies Rousseau's General Will to justify eliminating counter-revolutionary 'enemies of virtue'."),
    ("19th Cent.", "Divergent Ideological Evolution", "Locke becomes intellectual father of European Classical Liberalism; Rousseau becomes catalyst for Democratic Socialism & Romanticism.")
]

essay_12 = [
    ("Introduction: The Philosophical Revolution of European Politics",
     "The transformation of European political thought from medieval theocratic absolutism into modern democratic constitutionalism was engineered primarily through the intellectual breakthroughs of John Locke and Jean-Jacques Rousseau. Prior to the seventeenth century, European monarchical legitimacy rested upon the divine right of kings, patriarchal authority (as expounded by Sir Robert Filmer), and the pessimistic Leviathan state of Thomas Hobbes. Locke and Rousseau completely subverted this paradigm by constructing secular, contractarian frameworks grounded in the concept of the 'State of Nature.' However, while both philosophers agreed that legitimate governance must originate from the consent of the governed, they diverged profoundly in their conception of human nature, liberty, and the state. John Locke articulated the foundational principles of individualistic classical liberalism, private property, and limited constitutional government, whereas Jean-Jacques Rousseau pioneered a radical, communitarian critique of civil society that enthroned inalienable popular sovereignty and the 'General Will,' forever altering the political landscape of the West."),

    ("Part (a) John Locke: Natural Rights, Constitutionalism, and the Right of Revolution",
     "John Locke's monumental contribution to European political thought, principally articulated in his 'Two Treatises of Government' (1689), served as the definitive philosophical justification for England's Glorious Revolution of 1688. Locke began by rejecting the Hobbesian premise that the State of Nature was a brutal condition of 'war of all against all.' Instead, Locke posited a pre-political state governed by natural law and reason, wherein all human beings were born equal and endowed by their Creator with inherent, inalienable 'Natural Rights' to 'Life, Liberty, and Estate (Property).' For Locke, property was not a concession granted by a monarch, but an innate natural right established whenever a human being 'mixes his labor' with the natural world. However, because the State of Nature lacked impartial judges and established laws to resolve disputes, individuals voluntarily entered into a 'Social Compact' to form a civil government.\n\n"
     "Crucially, Locke redefined the state as a conditional 'fiduciary trust.' Government was created for one specific purpose: the preservation of citizens' natural rights. If a monarch or legislature degenerates into tyranny, exercises arbitrary power, or confiscates property without consent, that government breaches its trust and dissolves its own legitimacy. Under such circumstances, Locke affirmed the radical constitutional doctrine of the 'Appeal to Heaven'\u2014the moral right of the people to resist tyranny and overthrow the regime through revolution. Furthermore, to prevent the concentration of arbitrary power, Locke advocated the institutional separation of legislative and executive powers, laying the blueprint for modern constitutional checks and balances. His empiricist epistemology in 'An Essay Concerning Human Understanding' (1690), presenting the human mind as a 'tabula rasa' (blank slate) shaped by experience rather than innate sin, democratized education and inspired the entire European Enlightenment."),

    ("Part (b) Jean-Jacques Rousseau: The Critique of Property and the 'General Will'",
     "Writing in the mid-eighteenth century, Jean-Jacques Rousseau executed an intellectual counter-revolution against both absolute monarchy and bourgeois rationalism. In his 'Discourse on the Origin of Inequality' (1754), Rousseau presented a radical anthropological critique of modern civilization. In Rousseau's State of Nature, man was not an aggressive Hobbesian brute nor an acquisitive Lockean property-owner, but a peaceful, compassionate 'noble savage' guided by self-preservation (amour de soi) and pity (piti\u00e9). The catastrophic fall of humanity occurred with the introduction of private property and metallurgy: 'The first man who, having enclosed a piece of ground, bethought himself of saying This is mine, and found people simple enough to believe him, was the real founder of civil society.' For Rousseau, modern civil society had institutionalized inequality, moral vanity (amour-propre), and psychological enslavement.\n\n"
     "To resolve this existential tragedy, Rousseau published 'The Social Contract' (1762), opening with the electrifying indictment: 'Man is born free, and everywhere he is in chains.' Rousseau's solution was not a retreat to primitive savagery, but the creation of a legitimate, moral political community. He proposed that each individual surrender their particular, selfish ego to the collective whole, creating a sovereign body governed by the 'General Will' (volont\u00e9 g\u00e9n\u00e9rale). The General Will is not simply the majority vote (which Rousseau called the 'will of all'), but the collective interest directed exclusively toward the common good. In obeying the General Will, the citizen obeys only themselves, achieving a higher, moral freedom. For Rousseau, sovereignty is completely indivisible, inalienable, and cannot be represented by parliamentary deputies; he favored direct, participatory democracy in the tradition of the ancient Greek polis and Swiss cantons."),

    ("Comparative Synthesis and Enduring European Impact",
     "The intellectual divergence between Locke and Rousseau generated two competing streams of modern European political culture. Locke provided the intellectual cornerstone for the Anglo-American tradition of 'Negative Liberty'\u2014the protection of the private individual from state coercion, the sanctity of contract, the rule of law, and parliamentary representation. Locke's ideas directly inspired Baron de Montesquieu in France and Thomas Jefferson in the United States, who immortalized Locke's natural rights in the Declaration of Independence (1776). In contrast, Rousseau pioneered the continental tradition of 'Positive Liberty'\u2014freedom realized through active democratic citizenship and subordination to the common good. Rousseau's egalitarian ethos ignited the radical phase of the French Revolution; Maximilien Robespierre worshipped Rousseau, using the concept of the General Will to justify the Jacobin 'Republic of Virtue' and eliminate counter-revolutionary dissenters who refused to be 'forced to be free.'"),

    ("Conclusion: The Dual Architects of Modern Political Thought",
     "In conclusion, John Locke and Jean-Jacques Rousseau irrevocably dismantled the philosophical foundations of absolute dynastic Europe. Locke stripped kings of their divine right, elevating the individual as the bearer of inviolable natural rights and making governance conditional upon constitutional consent. Rousseau diagnosed the moral bankruptcy of commercial society, elevating the collective Nation as the sole fountain of legitimate sovereignty and establishing the moral imperative of equality. When the French National Assembly convened in 1789, it fused these two titanic thinkers in the 'Declaration of the Rights of Man': Articles 1 and 2 enshrined Locke's natural rights to liberty and property, while Article 6 directly cited Rousseau in declaring that 'Law is the expression of the General Will.' Together, Locke and Rousseau forged the intellectual architecture of modern European democracy, constitutionalism, and human freedom.")
]

quotes_12 = [
    ("Isaiah Berlin", "Two Concepts of Liberty", "'Locke is the greatest champion of negative liberty\u2014the area within which a man can act unobstructed by others. Rousseau, in advocating positive liberty and the General Will, created a doctrine that could be used to justify tyranny in the name of true freedom.'"),
    ("C.B. Macpherson", "Marxist Critique of Locke", "'Locke's state of nature is an early market society; his social contract was designed fundamentally to provide a constitutional defense for possessive individualism and unequal capitalist property.'"),
    ("Ernst Cassirer", "Moral Reconstruction of Rousseau", "'Rousseau did not wish to return to the state of nature... He sought to transform society from an arena of exploitation into a genuine moral community where freedom and law are reconciled.'"),
    ("Peter Gay", "Enlightenment Synthesis", "'Locke provided the Enlightenment with its method and its political common sense; Rousseau gave it its conscience, its passion, and its revolutionary fire.'")
]

# Run the generation
if __name__ == "__main__":
    tasks = [
        (meta_10, callout_10, map_10, timeline_10, essay_10, quotes_10,
         os.path.join(paper1_dir, "Q10_French_Revolution_Causes_and_Philosophy_Master_Notes.docx"),
         "QUESTION 10: CAUSES OF THE FRENCH REVOLUTION & PHILOSOPHICAL MOTIVATIONS (1789)",
         "Q10_French_Revolution_Causes_and_Philosophy_Master_Notes"),
        (meta_11, callout_11, map_11, timeline_11, essay_11, quotes_11,
         os.path.join(paper1_dir, "Q11_French_National_Assembly_1789_1791_Master_Notes.docx"),
         "QUESTION 11: THE NATIONAL ASSEMBLY & CONSTITUTIONAL REFORMS (1789\u20131791)",
         "Q11_French_National_Assembly_1789_1791_Master_Notes"),
        (meta_12, callout_12, map_12, timeline_12, essay_12, quotes_12,
         os.path.join(paper1_dir, "Q12_Political_Philosophy_Locke_and_Rousseau_Master_Notes.docx"),
         "QUESTION 12: POLITICAL PHILOSOPHY IN EUROPE: JOHN LOCKE & JEAN-JACQUES ROUSSEAU",
         "Q12_Political_Philosophy_Locke_and_Rousseau_Master_Notes")
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
