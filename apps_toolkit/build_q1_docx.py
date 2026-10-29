import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

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

def build_docx():
    doc = Document()
    
    # Page setup
    sections = doc.sections
    for s in sections:
        s.top_margin = Inches(0.8)
        s.bottom_margin = Inches(0.8)
        s.left_margin = Inches(0.8)
        s.right_margin = Inches(0.8)
        
    # Title
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
    r_sub = p_sub.add_run("MASTER STUDY NOTE & MODEL ANSWER &bull; 2025 EXAMINATION")
    r_sub.font.name = "Arial"
    r_sub.font.size = Pt(16)
    r_sub.font.bold = True
    r_sub.font.color.rgb = RGBColor(14, 43, 80)
    
    # Meta Box Table
    meta_table = doc.add_table(rows=4, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_table.autofit = False
    
    meta_data = [
        ("EXAM QUESTION (2025):", "Q.1 Examine the internal and external factors that contributed to the rise of Napoleon Bonaparte. How far did Napoleon's rule continue or betray the ideals of the French Revolution?"),
        ("TOTAL MARKS & FORMAT:", "20 Marks &bull; Descriptive Analytical Essay &bull; Recommended Time: 35 Minutes"),
        ("CORE SYLLABUS TOPIC:", "Section 2: The Napoleonic Era (1799–1815) & Legacy of the 1789 Revolution"),
        ("KEY HISTORIOGRAPHICAL LENS:", "The 'Napoleonic Paradox': Authoritarian Consolidator (Social Equality) vs. Imperial Autocrat (Political Liberty)")
    ]
    
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
    
    # Callout
    add_callout(doc, [
        "1. The question has TWO distinct halves: (A) Factors of Rise (Internal vs. External), and (B) Evaluation of Ideals (Continuator vs. Betrayer). Both must receive equal analytical weight.",
        "2. The French Revolution's motto was threefold: Liberty (political rights), Equality (legal rights), and Fraternity (national brotherhood). Napoleon preserved Equality but destroyed Liberty.",
        "3. High scoring answers cite exact institutional reforms: Code Napoléon (1804), Concordat (1801), Bank of France (1800), and Lycées."
    ])
    
    doc.add_paragraph().paragraph_format.space_after = Pt(10)
    
    # Section 1: Conceptual Flowchart
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(16)
    h1.paragraph_format.space_after = Pt(4)
    r_h1 = h1.add_run("1. CONCEPTUAL MIND MAP: CAUSAL PATHWAY TO POWER")
    r_h1.font.name = "Arial"
    r_h1.font.size = Pt(12)
    r_h1.font.bold = True
    r_h1.font.color.rgb = RGBColor(14, 43, 80)
    
    map_text = (
        "1789: Outbreak of Revolution ➔ Overthrow of Feudal Privileges & Absolute Monarchy\n"
        "  └── 1793–1794: Reign of Terror (Robespierre) ➔ Guillotine, Paranoia, Mass Executions\n"
        "        └── 1795–1799: The Directory ➔ Extreme Corruption, Inflation, Bread Shortages, Chaos\n"
        "              ├── INTERNAL VACUUM: Bourgeoisie & Peasants demand Law, Order, and Property Security\n"
        "              ├── EXTERNAL GLORY: Coalition Wars ➔ Napoleon emerges as undefeated National Hero\n"
        "              └── 1799: Coup of 18 Brumaire ➔ Napoleon takes power ➔ Consulate ➔ Empire (1804)"
    )
    p_map = doc.add_paragraph()
    p_map.paragraph_format.space_after = Pt(12)
    r_m = p_map.add_run(map_text)
    r_m.font.name = "Consolas"
    r_m.font.size = Pt(9)
    r_m.font.color.rgb = RGBColor(15, 23, 42)
    
    # Section 2: Chronological Timeline
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(14)
    h2.paragraph_format.space_after = Pt(4)
    r_h2 = h2.add_run("2. CHRONOLOGICAL MILESTONES (1793 – 1815)")
    r_h2.font.name = "Arial"
    r_h2.font.size = Pt(12)
    r_h2.font.bold = True
    r_h2.font.color.rgb = RGBColor(14, 43, 80)
    
    tbl_time = doc.add_table(rows=8, cols=3)
    tbl_time.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_time.autofit = False
    
    headers = ["YEAR", "EVENT / ACTION", "STRATEGIC SIGNIFICANCE"]
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
        
    timeline_data = [
        ("1793", "Siege of Toulon", "Recaptured French naval base from British fleet; promoted to Brigadier General at age 24."),
        ("1795", "13 Vendémiaire ('Whiff of Grapeshot')", "Defended the National Convention from Royalist mob; earned total trust of Paris political leaders."),
        ("1796–1797", "Italian Campaign (Montenotte, Lodi, Rivoli)", "Defeated 4 Austrian armies; personally negotiated Treaty of Campo Formio; emerged as popular icon."),
        ("1798–1799", "Egyptian Expedition & Battle of Pyramids", "Military attempt to disrupt British trade routes to India; used scientific/archeological propaganda."),
        ("Nov 1799", "Coup of 18 Brumaire", "Overthrew the corrupt Directory with Abbé Sieyès; established the French Consulate as First Consul."),
        ("1801–1804", "Concordat & Code Napoléon", "Reconciled state with Catholic Church; published uniform Civil Code establishing legal equality."),
        ("1804–1815", "Coronation as Emperor & Waterloo (1815)", "Crowned himself Emperor; total militarization; ultimate defeat at Waterloo by Coalition forces.")
    ]
    
    col_widths = [Inches(1.0), Inches(2.3), Inches(3.4)]
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
    
    # Section 3: The 20-Mark Master Model Essay
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(18)
    h3.paragraph_format.space_after = Pt(6)
    r_h3 = h3.add_run("3. THE 20-MARK MASTER MODEL ESSAY (~650 WORDS)")
    r_h3.font.name = "Arial"
    r_h3.font.size = Pt(13)
    r_h3.font.bold = True
    r_h3.font.color.rgb = RGBColor(14, 43, 80)
    
    essay_sections = [
        ("Introduction", 
         "The French Revolution of 1789 began as a historic crusade for 'Liberty, Equality, and Fraternity,' dismantling absolute monarchy and centuries of feudal servitude. Yet, a turbulent decade of radical Jacobin terror, runaway inflation, and foreign invasions left France utterly exhausted. This vacuum of stability enabled Napoleon Bonaparte—a young Corsican artillery officer—to seize sovereign power via the Coup of 18 Brumaire in November 1799. Napoleon’s meteoric rise was forged at the intersection of acute domestic decay and unparalleled foreign military triumph. Once installed as ruler, his regime embodied a profound historical paradox: while he permanently institutionalized the Revolution’s social and legal conquests—principally civil equality, meritocracy, and secular law—he systematically extinguished its democratic spirit, popular sovereignty, and individual liberty."),
        
        ("Internal Factors: Domestic Collapse and the Demand for Order",
         "The domestic collapse of France under the Directory (1795–1799) created the precise social and political vacuum necessary for a military strongman. Following the bloody Reign of Terror (1793–1794), executive power was entrusted to a five-member Directory plagued by deep corruption, administrative ineptitude, and financial bankruptcy. The national paper currency (assignats) became worthless, bread scarcity triggered widespread urban misery, and royalist rebellions threatened civil war. Both the commercial middle class (bourgeoisie) and land-owning peasants lived in dread of either a resurgence of radical Jacobin bloodshed or a Bourbon restoration that would confiscate their newly gained lands. They desperately sought a non-partisan savior who could guarantee domestic tranquility and protect their property. Having demonstrated his loyalty by defending the government with a 'whiff of grapeshot' during the royalist uprising of 1795, Napoleon appeared as an incorruptible arbiter above petty factional strife."),
        
        ("External Factors: Revolutionary Warfare and Military Glory",
         "Externally, the prolonged French Revolutionary Wars against monarchical coalitions militarized the French state and elevated the army as the sole cohesive, respected institution in the nation. Napoleon capitalized on this reality through peerless tactical genius. His sensational Italian Campaign of 1796–1797—marked by victories at Lodi, Castiglione, and Rivoli—shattered Austrian forces and resulted in the Treaty of Campo Formio, which he negotiated independently of the Directory. Even the strategic setback of his Egyptian Expedition (1798) was transformed via state propaganda into a romantic cultural triumph. By 1799, as the War of the Second Coalition threatened French frontiers once more, a panic-stricken nation turned instinctively toward its most celebrated military prodigy. Napoleon’s unmatched martial charisma eclipsed all civilian politicians, paving the way for the bloodless overthrow of the Directory."),
        
        ("Continuation of Revolutionary Ideals: Equality, Meritocracy, and Institutional Reform",
         "In governance, Napoleon operated as the 'Son of the Revolution' by solidifying its social conquests into enduring state institutions. His crowning achievement, the Civil Code of 1804 (Code Napoléon), enshrined legal equality, abolished feudal privileges, recognized secular property rights, and guaranteed religious tolerance. Furthermore, Napoleon realized the revolutionary demand for meritocracy—'careers open to talents' (la carrière ouverte aux talents). Administrative posts, civil honors, and military ranks were awarded based on talent and loyalty rather than aristocratic lineage, symbolized by the prestigious Legion of Honour. Through the Concordat of 1801 with Pope Pius VII, Napoleon healed the religious schism while compelling the Papacy to accept the permanent confiscation of church estates and state supremacy, ensuring that the feudal privileges of the clergy could never be restored."),
        
        ("Betrayal of Revolutionary Ideals: Autocracy, Censorship, and the Destruction of Liberty",
         "Conversely, Napoleon ruthlessly liquidated the foremost tenet of 1789: political Liberty. He dismantled parliamentary democracy, reducing representative assemblies to impotent advisory bodies and substituting genuine elections with stage-managed plebiscites. Freedom of the press was systematically erased; independent publications were suppressed until only four government-monitored mouthpieces survived in Paris. Through his minister of police, Joseph Fouché, political opponents, royalists, and republicans alike were subjected to arbitrary surveillance, imprisonment, and exile without trial. In 1804, Napoleon buried the Republic entirely by crowning himself hereditary Emperor of the French. Domestically, the Civil Code stripped women of revolutionary legal gains, re-imposing complete patriarchal subjugation. Finally, he revived aristocratic court hierarchy and practiced blatant dynastic nepotism, placing his siblings on satellite thrones across Europe."),
        
        ("Conclusion",
         "In final analysis, Napoleon Bonaparte was neither the immaculate champion nor the mere assassin of the French Revolution; he was its authoritarian consolidator. He astutely recognized that the vast majority of French citizens treasured the tangible social equality and economic security won in 1789 far more than the volatile political freedom that had brought forth the guillotine. By entrenching civil equality, eradicating feudal bondage, and modernizing state administration, he rescued the revolutionary social order and exported it across the European continent. However, by sacrificing representative democracy and personal liberty upon the altar of personal ambition and imperial order, Napoleon proved that under his rule, the ideals of 1789 were selectively preserved only so far as they fortified his autocracy.")
    ]
    
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
    
    # Section 4: Historiographical Perspectives
    h4 = doc.add_paragraph()
    h4.paragraph_format.space_before = Pt(16)
    h4.paragraph_format.space_after = Pt(4)
    r_h4 = h4.add_run("4. HISTORIOGRAPHICAL PERSPECTIVES & EXAMINER QUOTES")
    r_h4.font.name = "Arial"
    r_h4.font.size = Pt(12)
    r_h4.font.bold = True
    r_h4.font.color.rgb = RGBColor(14, 43, 80)
    
    tbl_quotes = doc.add_table(rows=6, cols=3)
    tbl_quotes.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_quotes.autofit = False
    
    q_headers = ["HISTORIAN", "CRITICAL PERSPECTIVE", "EXACT QUOTABLE ARGUMENT"]
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
        
    quotes_data = [
        ("Georges Lefebvre", "Marxist / Social Consolidator", "Napoleon was an enlightened despot, the true son of the Revolution who reconciled social order with the legal conquests of 1789."),
        ("Albert Vandal", "Political Realist", "Napoleon did not destroy the Republic; the Republic had already exhausted and destroyed itself long before 18 Brumaire."),
        ("Pieter Geyl", "Democratic Skeptic", "Napoleon's rule was fundamentally a military dictatorship disguised in the trappings of Roman law and theatrical plebiscites."),
        ("Madame de Staël", "Contemporary Liberal Critic", "A man who divided all humanity into two categories: himself, and the instruments he used to achieve his absolute will."),
        ("Stuart T. Miller", "Palgrave Modern History", "The French wanted order and security above all else; Napoleon gave them both, but the price paid was total political submission.")
    ]
    
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
                
    # Footer Note
    doc.add_paragraph().paragraph_format.space_after = Pt(16)
    p_foot = doc.add_paragraph()
    p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_f = p_foot.add_run("Khyber Pakhtunkhwa Public Service Commission (KPPSC) &bull; PMS Competitive Examination Notes &bull; European History")
    r_f.font.name = "Arial"
    r_f.font.size = Pt(8.5)
    r_f.font.color.rgb = RGBColor(148, 163, 184)
    
    out_docx = r"H:\My Drive\PMS KPK 2022 notes\1. European History Paper 1 and Paper 2\Paper 1\Q1_Napoleon_Rise_and_Legacy_2025_Master_Notes.docx"
    doc.save(out_docx)
    print("Document successfully created:", out_docx)

if __name__ == "__main__":
    build_docx()
