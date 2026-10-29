import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
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
    
    # Meta Box Table
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
    
    # Section 1: Conceptual Flowchart
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
    
    # Section 2: Chronological Timeline
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(14)
    h2.paragraph_format.space_after = Pt(4)
    r_h2 = h2.add_run("2. CHRONOLOGICAL MILESTONES & INSTITUTIONAL DATES")
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
    
    # Section 3: The 20-Mark Master Model Essay
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
    
    # Section 4: Historiographical Perspectives
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
    
    q_headers = ["HISTORIAN", "ANALYTICAL SCHOOL", "EXACT QUOTABLE ARGUMENT"]
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

# =========================================================================
# QUESTION A: 2022 Q.2 / 2016 Q.2 / 2013 Q.2 (INTERNAL REFORMS & "I AM THE REVOLUTION")
# =========================================================================

meta_A = [
    ("EXAM QUESTIONS LINKED:", "2022 Q.2: 'In the light of his internal reforms, discuss whether Napoleon Bonaparte promoted or halted the French Revolution.'\n2016 Q.2: 'Napoleon Bonaparte declared: I am the Revolution. To what extent is this claim justified by his civil, judicial, and educational reforms?'\n2013 Q.2: 'Napoleonic reforms prove him to be a great administrator and a statesman. Discuss.'"),
    ("TOTAL MARKS & TIME:", "20 Marks &bull; Descriptive Subjective Essay &bull; Time: 35 Minutes"),
    ("SYLLABUS CORRELATION:", "Paper I &bull; Section 2: Napoleon as Statesman, Administrator & Lawgiver"),
    ("HISTORIOGRAPHICAL THESIS:", "Constructive Consolidator of Social Gains (Equality) vs. Liquidator of Political Ideals (Liberty).")
]

callout_A = [
    "1. Napoleon famously declared: 'I am the Revolution'—meaning he rescued its practical gains from anarchy, but he also stated: 'The romance of the Revolution is over; we must now begin its history.'",
    "2. The 5 pillars of Napoleonic statesmanship: (1) Code Napoléon (1804), (2) Concordat (1801), (3) Bank of France & Currency (1800), (4) Educational Lycées & University (1802/1806), (5) Prefectorial Administrative Centralization.",
    "3. Formula for 16+/20: Clearly distinguish between the 'Destructive Phase' of 1789–1794 (overthrow of monarchy) and the 'Constructive Phase' of 1799–1804 (institutionalization of law)."
]

map_A = (
    "REVOLUTIONARY CHAOS (1789–1799) ➔ 300+ Local Legal Systems, Currency Collapse, Church Schism\n"
    "  │\n"
    "  ├── LEGAL CODIFICATION: 1804 Code Napoléon ➔ Uniformity, Civil Equality, Abolition of Feudalism\n"
    "  ├── RELIGIOUS SETTLEMENT: 1801 Concordat ➔ Church Property Confiscated, State Supremacy Assured\n"
    "  ├── ECONOMIC STABILIZATION: 1800 Bank of France ➔ Franc de Germinal (stable for 125 years)\n"
    "  ├── EDUCATIONAL REGULATION: 1802 Lycées & 1806 Imperial University ➔ Meritocratic Bureaucrats\n"
    "  └── ADMINISTRATIVE CENTRALIZATION: Prefects in every Department ➔ Efficiency at cost of Local Democracy"
)

timeline_A = [
    ("Jan 1800", "Founding of the Bank of France", "Replaced worthless assignats with stable gold-backed 'Franc Germinal', ending hyperinflation."),
    ("Feb 1800", "Law on Administrative Organization", "Created system of Prefects and Sub-prefects directly appointed by Paris, restoring order."),
    ("July 1801", "The Concordat with Pope Pius VII", "Healed religious civil war while ensuring the state kept confiscated church lands."),
    ("May 1802", "Establishment of the Lycées & Legion of Honour", "Created state secular secondary schools and meritocratic civic order open to all classes."),
    ("March 1804", "Promulgation of the Civil Code (Code Napoléon)", "Unified French jurisprudence into 2,281 articles guaranteeing legal equality and property."),
    ("Dec 1804", "Imperial Coronation at Notre-Dame", "Ended republican experiment; established hereditary imperial monarchy."),
    ("May 1806", "Creation of the Imperial University", "Centralized national control over public education, teacher licensing, and state curriculum.")
]

essay_A = [
    ("Introduction",
     "Napoleon Bonaparte’s famous declaration, 'I am the Revolution,' encapsulates the central paradox of modern French history. When Napoleon seized power as First Consul in 1799, the ideals of 1789—Liberty, Equality, and Fraternity—were imperiled by a decade of administrative bankruptcy, religious civil war, and judicial confusion. Rather than viewing the Revolution as an ongoing political crusade, Napoleon regarded it as a completed earthquake whose valid structural gains needed urgent consolidation. Through a series of monumental internal reforms spanning law, religion, finance, and education, Napoleon proved to be an administrative genius. However, his reforms systematically separated equality from liberty: he permanently embedded the social, economic, and legal victories of the Revolution while firmly suppressing its democratic and representative spirit."),

    ("Legal and Judicial Reforms: The Code Napoléon (1804)",
     "The cornerstone of Napoleon’s claim to have fulfilled the Revolution rests upon the Civil Code of 1804 (Code Napoléon). Prior to 1789, France was fractured into over 300 conflicting local legal traditions, with Roman law dominating the south and feudal customary law in the north. Napoleon mobilized veteran jurists to distill French law into a single, uniform national code of 2,281 articles. The Code permanently preserved the core social conquests of 1789: absolute equality before the law, the irreversible abolition of feudalism, freedom of conscience, and the inviolability of private property. It cemented the middle-class and peasant land transfers resulting from the sale of church and noble estates. However, the Code also halted revolutionary egalitarianism in family law, reducing married women to legal minors under total patriarchal tutelage."),

    ("Religious Reconciliation: The Concordat of 1801",
     "The Civil Constitution of the Clergy (1790) had alienated millions of devout French Catholics and incited vicious civil uprisings in the Vendée. Recognizing that internal peace required religious reconciliation, Napoleon negotiated the Concordat of 1801 with Pope Pius VII. In doing so, he astutely promoted the Revolution's secular supremacy rather than surrendering to the Church. Catholicism was recognized not as the state religion, but merely as 'the religion of the great majority of French citizens.' In return, the Papacy renounced all claims to confiscated Church lands, agreed that bishops would be nominated by the French head of state, and accepted state payment of clergy salaries. By subordinating the Church to the civil state while guaranteeing freedom of worship to Protestants and Jews, Napoleon consolidated the secularization of French society."),

    ("Economic and Administrative Modernization",
     "The Revolution had collapsed under the weight of financial chaos and hyperinflated paper currency (*assignats*). Napoleon restored economic viability by founding the **Bank of France (1800)** and introducing the sound, silver-backed **Franc Germinal (1803)**, which remained stable for over a century. He instituted an efficient, equitable tax collection system overseen by professional state receivers, ending centuries of corrupt tax farming. Administratively, Napoleon reorganized France through the **Law of 28 Pluviôse (1800)**, placing each department under a centralized **Prefect** appointed directly by the central government. This eliminated chaotic local elections in favor of uniform, rapid execution of state policy, transforming France from a dysfunctional federation into the most streamlined administrative machine in Europe."),

    ("Educational Reforms and Meritocracy: 'Careers Open to Talents'",
     "To institutionalize meritocracy, Napoleon established a national secondary school system centered on state-run **Lycées (1802)** and later crowned by the **Imperial University (1806)**. Admission was determined by competitive examinations and scholarships rather than noble birth, training a new administrative elite of engineers, magistrates, and officers based on intellectual capability. This embodied the revolutionary principle of *la carrière ouverte aux talents* (careers open to talents). To reward civic and military excellence outside the feudal nobility, he created the **Legion of Honour (1802)**. While these reforms democratized upward social mobility, education was strictly regimented with military discipline and patriotic indoctrination, designed to produce obedient servants of the state rather than free-thinking democratic citizens."),

    ("Conclusion",
     "In ultimate balance, Napoleon Bonaparte promoted the social and legal Revolution while deliberately halting its political revolution. He institutionalized civil equality, destroyed the feudal social structure, secularized public life, and constructed an enduring bureaucratic apparatus that survives in France to this day. Yet, he achieved this stability by murdering political liberty: representative assemblies were neutered, the press was gagged, and the sovereign Republic was eclipsed by an imperial monarchy. Napoleon did not betray the Revolution in its entirety; rather, as historian Georges Lefebvre observed, he made a calculated compromise—rescuing the middle-class revolution of equality and property by anchoring it to an enlightened military autocracy.")
]

quotes_A = [
    ("Georges Lefebvre", "Social Historian", "Napoleon was the true executor of the Revolution's social conquests; he secured the triumph of the bourgeoisie and peasant proprietors over feudalism."),
    ("H.A.L. Fisher", "Classical Historian", "The Code Napoléon was the greatest legal monument of modern times, broadcasting the egalitarian principles of 1789 across the continent of Europe."),
    ("Alfred Cobban", "Revisionist Historian", "Napoleon did not save the Revolution; he ended it by transforming revolutionary nationalism into an instrument of personal imperial despotism."),
    ("Stuart T. Miller", "Palgrave Modern History", "His reforms were brilliant exercises in statecraft: he gave the French the equality they craved, while quietly depriving them of the liberty they could not manage."),
    ("Napoleon Bonaparte", "Primary Quote (1804)", "My true glory is not to have won forty battles; Waterloo will erase the memory of so many victories. What nothing will erase, what will live forever, is my Civil Code.")
]

# =========================================================================
# QUESTION B: 2018 Q.2 / 2013 Q.3 (THE CONTINENTAL SYSTEM & DOWNFALL)
# =========================================================================

meta_B = [
    ("EXAM QUESTIONS LINKED:", "2018 Q.2: 'Analyze the Continental System of Napoleon Bonaparte. Why did it fail to bring Great Britain to its knees, and how did it contribute to Napoleon's downfall?'\n2013 Q.3: 'Continental System became a catalyst for the downfall of Napoleon Bonaparte. Discuss.'"),
    ("TOTAL MARKS & TIME:", "20 Marks &bull; Descriptive Subjective Essay &bull; Time: 35 Minutes"),
    ("SYLLABUS CORRELATION:", "Paper I &bull; Section 2: Napoleon's Foreign Policy, Economic Warfare & Downfall"),
    ("HISTORIOGRAPHICAL THESIS:", "Economic warfare against British industrial supremacy boomeranged, compelling fatal territorial overreach in Spain (1808) and Russia (1812).")
]

callout_B = [
    "1. After the Battle of Trafalgar (1805), Napoleon could not invade Britain militarily. The Continental System was an economic weapon designed to starve Britain of European export markets and drain its bullion reserves.",
    "2. The Decisive Decrees: Berlin Decrees (Nov 1806) declared British Isles in a state of blockade; Milan Decrees (Dec 1807) authorized seizure of neutral ships obeying British Orders in Council.",
    "3. Why it doomed Napoleon: Britain possessed global trade routes and naval supremacy; European allies starved for colonial goods (coffee, sugar, cloth) grew rebellious, forcing Napoleon into the fatal 'Spanish Ulcer' and the disastrous Moscow campaign."
]

map_B = (
    "1805 TRAFALGAR: British Naval Mastery ➔ Direct French invasion of Britain impossible\n"
    "  │\n"
    "  ├── 1806–1807: Berlin & Milan Decrees ➔ Continental Blockade of all British commerce\n"
    "  │     └── British Counter: Orders in Council (1807) ➔ Royal Navy blockades Europe\n"
    "  │\n"
    "  ├── ECONOMIC BOOMERANG: Europe suffers inflation, shortages, and economic paralysis\n"
    "  │     ├── Smuggling flourishes; French industry cannot supply continental demand\n"
    "  │     └── Neutral and allied nations grow deeply alienated and rebellious\n"
    "  │\n"
    "  └── FATAL MILITARY OVEREXTENSION TO ENFORCE THE BLOCKADE:\n"
    "        ├── 1807–1808: Portugal & Spain Resist ➔ The Peninsular War ('Spanish Ulcer')\n"
    "        └── 1812: Tsar Alexander I rejects Blockade ➔ Invasion of Russia ➔ Total Ruin"
)

timeline_B = [
    ("Oct 1805", "Battle of Trafalgar", "Admiral Nelson destroys Franco-Spanish fleet; confirms undisputed British command of the seas."),
    ("Nov 1806", "The Berlin Decree", "Napoleon proclaims the Continental Blockade, banning all trade and communication with Britain from French-controlled Europe."),
    ("Jan 1807", "British Orders in Council", "Britain retaliates: neutral vessels must call at British ports and pay transit duties before trading with continental Europe."),
    ("Dec 1807", "The Milan Decree", "Napoleon orders that any ship submitting to British searches or paying British dues will be seized as prize of war."),
    ("1807–1808", "Invasion of Portugal & Spain", "Napoleon invades Iberian Peninsula to close Lisbon to British ships, triggering the brutal 6-year Peninsular War."),
    ("Dec 1810", "Russian Ukase (Defiance of Blockade)", "Tsar Alexander I withdraws Russia from the Continental System and taxes French luxuries, defying Napoleon."),
    ("June 1812", "Invasion of Russia (Grande Armée)", "Napoleon marches 600,000 troops into Russia primarily to compel Russian compliance with the Blockade; returns with under 40,000."),
    ("1813–1814", "Battle of Leipzig & Abdication", "Exhausted France defeated by the Sixth Coalition; Napoleon abdicates to Elba.")
]

essay_B = [
    ("Introduction",
     "Following the decisive destruction of the Franco-Spanish fleet at the Battle of Trafalgar in October 1805, Napoleon Bonaparte recognized that direct military invasion of Great Britain was impossible. Britain’s military power rested upon its industrial supremacy, commercial wealth, and global financial credit. To break this maritime hegemon, Napoleon initiated an ambitious campaign of economic warfare known as the **Continental System**. Inaugurated by the **Berlin Decree (1806)** and tightened by the **Milan Decree (1807)**, the system sought to starve Britain of European export markets, induce financial insolvency, and drain its bullion reserves. However, the Continental System was an economic fallacy that boomeranged catastrophically against France. Instead of bringing Britain to its knees, it alienated continental allies, strangled European commerce, and compelled Napoleon into fatal military overextensions in Spain and Russia that directly precipitated his downfall."),

    ("Nature and Mechanics of the Continental Blockade",
     "The Continental System was an attempt to weaponize land power against sea power. Under the Berlin Decree, the British Isles were declared in a state of blockade; all commerce and correspondence with Britain were prohibited, and British merchandise on the continent was confiscated. Britain retaliated through its **Orders in Council (1807)**, declaring that neutral vessels could only trade with continental ports if they first stopped at British ports to pay maritime duties. Napoleon countered with the Milan Decree, proclaiming that any ship submitting to British inspection would be treated as an enemy prize. By 1808, following the Treaties of Tilsit, Napoleon had coerced Russia, Prussia, Austria, and the Scandinavian states into the blockade, encompassing virtually the entire European coastline from Saint Petersburg to Cadiz."),

    ("Why the System Failed to Bring Britain to Its Knees",
     "The Continental System failed to cripple Britain due to four fundamental economic realities. First, Britain maintained **undisputed naval supremacy**, allowing the Royal Navy to enforce a counter-blockade that cut continental Europe off from raw materials and overseas colonies. Second, Britain held a near-monopoly on overseas trade routes: when European markets closed, British merchants aggressively expanded into **Latin America, North America, and Asia**, compensating for European losses. Third, British industrial manufacturing (cotton textiles, iron, machinery) was irreplaceable; French industry was technologically incapable of supplying the continental consumer demand. Fourth, **rampant smuggling** undermined the blockade at every turn. Islands like Heligoland in the North Sea and British-held Gibraltar became massive contraband depots; even Napoleon’s own military contractors were forced to illicitly smuggle 50,000 British-made overcoats and boots to clothe the *Grande Armée*."),

    ("Economic Backlash and the Alienation of Europe",
     "While Britain weathered the crisis, the Continental System inflicted devastating economic damage upon continental Europe. Coastal port cities such as Amsterdam, Bordeaux, Hamburg, and Genoa saw their maritime shipping ruined, docks deserted, and merchant classes bankrupted. Simultaneously, ordinary Europeans suffered severe shortages and astronomical price inflation on essential consumer goods such as coffee, sugar, tea, tobacco, and cotton cloth. Landowners in Prussia and Russia, who relied heavily on exporting timber, hemp, and grain to Britain, faced economic ruin. Instead of directing their fury toward London, European populations blamed Napoleonic tyranny for their impoverishment, transforming popular admiration for French revolutionary principles into bitter, anti-French nationalism across the German and Dutch lands."),

    ("The Catalyst for Downfall: The Spanish Ulcer and Russian Catastrophe",
     "The decisive contribution of the Continental System to Napoleon’s downfall lay in the catastrophic military campaigns it forced him to wage to plug gaps in the blockade. In 1807, Portugal refused to enforce the embargo. Napoleon invaded Portugal and recklessly deposed the Spanish Bourbon monarchy in 1808, installing his brother Joseph. This triggered the savage **Peninsular War (1808–1814)**, a six-year guerrilla bloodbath supported by British forces under Wellington that consumed 300,000 French troops—what Napoleon ruefully termed his **'Spanish Ulcer'**. Worse followed when **Tsar Alexander I of Russia**, facing domestic economic collapse, withdrew from the Continental System in December 1810. To force Russian compliance, Napoleon launched his catastrophic **Invasion of Russia in 1812**. The annihilation of the 600,000-strong *Grande Armée* in the snows of Moscow shattered the myth of Napoleonic invincibility, giving birth to the Sixth Coalition that ultimately captured Paris in 1814."),

    ("Conclusion",
     "In retrospect, the Continental System was Napoleon Bonaparte's greatest strategic miscalculation and the prime catalyst of his imperial ruin. It was rooted in the flawed assumption that agricultural Europe could survive without British industrial goods while Britain could not survive without European markets. In reality, the blockade weaponized European nationalism against France, drained French financial resources, and created an impossible geopolitical dilemma: to maintain the blockade, Napoleon had to control every mile of European coastline, compelling an endless cycle of conquests and annexations. As historian David Thomson observed, the Continental System ceased to be an economic policy and became a military trap. It led Napoleon straight into the guerrilla quagmire of Spain and the frozen graveyard of Russia, sealing the fate of the Napoleonic Empire.")
]

quotes_B = [
    ("David Thomson", "Europe Since Napoleon", "The Continental System was an economic blunder that dictated Napoleon's foreign policy and dragged him from one military adventure to another until he was ruined."),
    ("J. Holland Rose", "Napoleonic Scholar", "Napoleon set out to conquer the sea by the land, but in doing so, he made every consumer on the continent an enemy of his empire."),
    ("Stuart T. Miller", "Mastering Modern European History", "The blockade could only work if it was absolute; but to make it absolute, Napoleon had to conquer all Europe, and that proved his undoing."),
    ("Georges Lefebvre", "French Historian", "The Continental System divided the French bourgeoisie and alienated the peasantry; it made peace impossible and war perpetual."),
    ("Napoleon Bonaparte", "St. Helena Memoirs", "That unfortunate war in Spain was the cause of all my misfortunes. It divided my forces, multiplied my difficulties, and opened a school for the English soldiers.")
]

# Generate both files
paper1_dir = r"H:\My Drive\PMS KPK 2022 notes\1. European History Paper 1 and Paper 2\Paper 1"

out_A = os.path.join(paper1_dir, "Q2_Napoleon_Internal_Reforms_and_Revolution_Promoted_or_Halted_Master_Notes.docx")
out_B = os.path.join(paper1_dir, "Q3_Napoleon_Continental_System_and_Downfall_Master_Notes.docx")

generate_master_doc(meta_A, callout_A, map_A, timeline_A, essay_A, quotes_A, out_A, "MASTER STUDY NOTE & MODEL ANSWER &bull; REFORMS & 'I AM THE REVOLUTION'")
generate_master_doc(meta_B, callout_B, map_B, timeline_B, essay_B, quotes_B, out_B, "MASTER STUDY NOTE & MODEL ANSWER &bull; CONTINENTAL SYSTEM & DOWNFALL")
