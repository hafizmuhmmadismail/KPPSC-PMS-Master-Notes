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
    
    headers = ["DATE / YEAR", "EVENT / TREATY / ACTION", "HISTORICAL & EXAM SIGNIFICANCE"]
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
# QUESTION 4: 2025 Q.2 / 2016 Q.3 / 2013 Q.4 (THE METTERNICH SYSTEM)
# =========================================================================

meta_4 = [
    ("EXAM QUESTIONS LINKED:", "2025 Q.2: 'Examine how the Metternich system sought to suppress revolutionary movements in post-Napoleonic Europe. Why did it ultimately fail?'\n2016 Q.3: 'Critically assess Prince Metternich as the prime defender of the conservative order in Europe from 1815 to 1848.'\n2013 Q.4: 'The defeat of Napoleon brought the dark ages of old order under the Metternich system. Discuss.'"),
    ("TOTAL MARKS & TIME:", "20 Marks &bull; Descriptive Subjective Essay &bull; Time: 35 Minutes"),
    ("SYLLABUS CORRELATION:", "Paper I &bull; Section 3: The Metternich System, Concert of Europe & Revolutions of 1848"),
    ("HISTORIOGRAPHICAL THESIS:", "A static dam constructed to hold back dynamic subterranean currents: repressive diplomacy maintained international peace while suppressing unstoppable domestic nationalism and liberalism.")
]

callout_4 = [
    "1. Metternich's core philosophy was 'Restoration and Stability': he viewed the Austrian polyglot empire (comprising 11 nationalities) as vulnerable to dissolution if Nationalism or Liberalism took root.",
    "2. The 3 mechanisms of repression: (a) Domestic Police State (Carlsbad Decrees 1819 in Germany, strict press censorship, university spies), (b) The Congress System (Troppau Protocol 1820 authorizing armed intervention), (c) Military Subjugation (Austrian troops in Naples & Piedmont; French troops in Spain).",
    "3. Why it collapsed: Britain withdrew from interventionism (Canning's 'Every nation for itself'); the Greek Revolution (1821) divided the autocrats; the Industrial Revolution empowered a middle class that exploded in the 1848 Spring of Nations, forcing Metternich to flee in disguise."
]

map_4 = (
    "AUSTRIAN VULNERABILITY: Polyglot Empire (Germans, Magyars, Slavs, Italians) ➔ Nationalism = Extinction\n"
    "  │\n"
    "  ├── IDEOLOGY OF STABILITY: 'Presume nothing; change nothing' ➔ Repression of Liberalism & Constitutionalism\n"
    "  │\n"
    "  ├── INSTRUMENTS OF SUPPRESSION:\n"
    "  │     ├── Domestic: Carlsbad Decrees (1819) ➔ Censorship, banned Burschenschaften, blacklisted academics\n"
    "  │     └── International: Troppau Protocol (1820) ➔ Holy Alliance armed intervention in Naples, Spain, Piedmont\n"
    "  │\n"
    "  └── COLLAPSE OF THE DAM (CAUSES OF FAILURE):\n"
    "        ├── Great Power Schism: Britain (Castlereagh/Canning) rejects policing domestic politics\n"
    "        ├── Nationalist Breakthroughs: Greek Independence (1832) & Belgian Revolution (1830)\n"
    "        └── THE FLOOD OF 1848: Paris, Vienna, Berlin, Budapest erupt ➔ Metternich flees to London"
)

timeline_4 = [
    ("June 1815", "Congress of Vienna Final Act", "Prince Klemens von Metternich architects the post-Napoleonic European conservative order."),
    ("Sept 1815", "The Holy Alliance & Quadruple Alliance", "Tsar Alexander I's mystic alliance and the Great Power coalition form the diplomatic apparatus of stability."),
    ("Sept 1819", "The Carlsbad Decrees", "Metternich dissolves German student fraternities, imposes universal censorship, and installs university curators."),
    ("Nov 1820", "The Congress and Protocol of Troppau", "Austria, Prussia, and Russia formally claim the right to overthrow revolutionary regimes by military force."),
    ("1821–1823", "Military Interventions in Italy & Spain", "Austrian armies crush constitutionalists in Naples and Piedmont; French forces crush Spanish liberals."),
    ("1821–1832", "Greek War of Independence", "Tsarist support for fellow Orthodox Greeks shatters the Holy Alliance front against rebellion."),
    ("July 1830", "Revolutions of 1830 in France & Belgium", "Bourbons ousted in Paris; independent constitutional Belgium created, breaching Metternich's cordon sanitaire."),
    ("March 1848", "The March Revolution in Vienna", "Students and workers storm the Hofburg; Metternich resigns and flees in a laundry cart to England.")
]

essay_4 = [
    ("Introduction",
     "Following the final defeat of Napoleon at Waterloo in 1815, European diplomacy was dominated for over three decades by Prince Klemens von Metternich, the Austrian State Chancellor. Presiding over a fragile, multi-ethnic Habsburg Empire composed of Germans, Magyars, Czechs, Poles, Croats, and Italians, Metternich recognized that the twin forces unleashed by the French Revolution—Liberalism (the demand for constitutional freedom) and Nationalism (the demand for ethnic self-determination)—represented an existential death sentence for Austria. Consequently, he constructed the 'Metternich System', a vast continental framework designed to freeze the political status quo, extinguish revolutionary sparks, and enforce conservative legitimacy. While this system preserved general peace between European monarchs for over thirty years, it operated as a rigid, artificial dam against unstoppable socio-economic tides, culminating in its spectacular and total collapse during the pan-European Revolutions of 1848."),

    ("Mechanisms of Domestic Suppression: The Police State and Carlsbad Decrees",
     "Metternich’s internal strategy was rooted in the complete suppression of free expression, intellectual inquiry, and political association. Viewing universities as hotbeds of sedition, he capitalized on the 1819 assassination of the conservative playwright August von Kotzebue by a radical student to impose the draconian Carlsbad Decrees (1819) across the German Confederation. These decrees dissolved the patriotic student fraternities (Burschenschaften), placed state overseers (curators) in every university, purged liberal professors, and imposed strict pre-publication censorship on all books and newspapers under twenty pages. A central investigative commission was established at Mainz to track down political dissidents. Within the Austrian lands, Metternich deployed an extensive network of secret police, mail-opening interception bureaus (Kabinettsarchiv), and passport restrictions, effectively insulating Central Europe from modern political discourse."),

    ("The International Apparatus: Congress System and the Troppau Protocol",
     "To crush revolutions beyond Austrian borders, Metternich transformed the Concert of Europe into an international police force. Through periodic diplomatic summits—Aix-la-Chapelle (1818), Troppau (1820), Laibach (1821), and Verona (1822)—he orchestrated collective monarchical intervention against popular rebellions. At the Congress of Troppau (1820), Metternich drafted the radical Troppau Protocol, which audaciously asserted that any European state undergoing an internal revolution ceased to be a recognized member of the European alliance and authorized foreign armies to intervene by force to restore the legitimate sovereign. Under this mandate, Austrian troops marched south in 1821 to extinguish nascent constitutional governments in Naples and Piedmont, while French forces were commissioned at Verona in 1822 to invade Spain and restore the despotic King Ferdinand VII."),

    ("Why the System Failed: Diplomatic Fractures and the Rise of Britain",
     "The ultimate failure of the Metternich system was propelled by structural fractures within the Great Power coalition itself. Great Britain, under Foreign Secretaries Lord Castlereagh and George Canning, adamantly rejected Metternich’s doctrine of global counter-revolutionary policing. Britain maintained that the alliance was formed to contain French military aggression, not to meddle in the domestic constitutional evolutions of sovereign nations. When Metternich proposed intervention in Spanish America, Canning recognized the newly independent Latin American republics and famously declared: 'Every nation for itself, and God for us all.' Furthermore, the Greek War of Independence (1821–1832) fatally split the eastern autocracies: Russia, motivated by pan-Orthodox religious solidarity and imperial ambition to dismantle Ottoman power in the Balkans, supported the Greek rebels against Metternich’s frantic pleas to preserve Ottoman legitimacy."),

    ("The Irresistible Forces of Change: Industrialization and the 1848 Explosion",
     "Beyond diplomacy, Metternich failed because he waged war against the fundamental socio-economic transformation of Europe. The rapid advance of the Industrial Revolution created a wealthy, educated industrial bourgeoisie that demanded political representation, alongside an impoverished urban proletariat prone to insurrection. In 1830, the July Revolution in Paris toppled the reactionary Charles X, while Belgium successfully rebelled against Dutch rule, creating a liberal constitutional monarchy that Metternich was powerless to prevent. By 1848, a devastating continental potato blight and industrial recession merged with middle-class political fury into the 'Spring of Nations'. In March 1848, revolutionary street barricades erupted in Vienna itself. Abandoned by the imperial court, Metternich was forced to resign in humiliation, fleeing disguised in a laundry cart to London as revolutions swept across France, Italy, Germany, and Hungary."),

    ("Conclusion",
     "In retrospective evaluation, Prince Metternich was an extraordinary diplomatic craftsman who mistook the preservation of peace between states for the suppression of life within states. His system succeeded as a short-term diplomatic mechanism, sparing Europe from general continental warfare between 1815 and 1853. However, it was fundamentally doomed because it treated living ideas as police problems. By diagnosing legitimate aspirations for representative democracy and national unity as a 'moral gangrene' to be excised with bayonets, Metternich blocked all constitutional evolution. When the dam inevitably burst in 1848, it shattered the conservative consensus, proving that while armies can temporarily disperse legislatures and censor pamphlets, they cannot indefinitely suppress the structural march of modern European history.")
]

quotes_4 = [
    ("A.J.P. Taylor", "The Habsburg Monarchy", "Metternich was a master of diplomatic intrigue, but he ruled over a stagnant swamp; his system was an elaborate attempt to stop the clock of European development."),
    ("Henry Kissinger", "A World Restored", "Metternich was the quintessential conservative statesman; he understood that an empire without national cohesion could only survive by elevating international stability above national self-determination."),
    ("Stephen J. Lee", "Aspects of European History", "The Metternich system was essentially negative; it sought to prevent rather than create, treating the symptoms of revolution while ignoring the socio-economic disease."),
    ("Stuart T. Miller", "Mastering Modern European History", "Metternich failed because he identified the interests of Europe with the survival of the Austrian Empire; when the forces of 1848 arose, the empire was an anachronism."),
    ("Prince Klemens von Metternich", "Personal Memoirs (1820)", "I am holding the status quo together with my bare hands; I spend my days propping up rotten buildings that threaten to collapse over my head.")
]

# =========================================================================
# QUESTION 5: 2025 Q.3 / 2022 Q.3 / 2018 Q.3 (THE CONGRESS OF VIENNA)
# =========================================================================

meta_5 = [
    ("EXAM QUESTIONS LINKED:", "2025 Q.3: 'Evaluate the effectiveness of the Congress of Vienna in maintaining peace in Europe between 1815 and 1848. To what extent did it succeed in reconciling conservative principles with the force of change?'\n2022 Q.3: 'What were the principles adopted at the Congress of Vienna (1814–1815)? How were those principles implemented in the decisions of the Congress?'\n2018 Q.3: 'Describe the Concert of Europe and evaluate its achievements and failures from 1815 to 1848.'"),
    ("TOTAL MARKS & TIME:", "20 Marks &bull; Descriptive Subjective Essay &bull; Time: 35 Minutes"),
    ("SYLLABUS CORRELATION:", "Paper I &bull; Section 3: The Vienna Settlement (1814–1815) & The Concert of Europe"),
    ("HISTORIOGRAPHICAL THESIS:", "A triumph of geopolitical equilibrium that prevented general European warfare for a century, but a reactionary failure in accommodating domestic nationalism and democratic constitutionalism.")
]

callout_5 = [
    "1. The Core Guiding Principles: (a) Balance of Power (just equilibrium to contain France without destroying it), (b) Legitimacy (restoring pre-1789 Bourbon/Habsburg monarchs), (c) Compensation (territorial spoils for Coalition victors), (d) Encirclement / Buffer States.",
    "2. The Great Peacemakers: Metternich (Austria), Castlereagh (Britain), Alexander I (Russia), Hardenberg (Prussia), and the brilliant survivor Talleyrand (France).",
    "3. The Dual Verdict: Spectacular success in international diplomacy (no general war between European powers until 1914, barring localized conflicts); total failure in domestic justice (carved up Poland, partitioned Italy as a 'geographical expression', and ignored German unification)."
]

map_5 = (
    "FALL OF NAPOLEON (1814–1815) ➔ Redrawing the Political Map of Europe at Vienna\n"
    "  │\n"
    "  ├── GEOPOLITICAL RECONSTRUCTION (PRINCIPLES):\n"
    "  │     ├── Balance of Power: Moderate peace with France (1790 borders) ➔ Talleyrand admitted to Big Five\n"
    "  │     ├── Buffer States (Cordon Sanitaire): United Kingdom of the Netherlands, Prussian Rhineland, Kingdom of Sardinia\n"
    "  │     └── Dynastic Legitimacy: Bourbons restored in France/Spain; Habsburgs in Lombardy-Venetia\n"
    "  │\n"
    "  ├── STRATEGIC SUCCESS: THE 'LONG PEACE' (1815–1853/1914)\n"
    "  │     └── Replaced chaotic balance of terror with multilateral diplomatic conferences (Concert of Europe)\n"
    "  │\n"
    "  └── SYSTEMIC DEFECT: COMPLETE REJECTION OF FORCES OF CHANGE\n"
    "        ├── Nationalism Ignored: Poland partitioned; Italy divided into 8 autocracies; Belgium forced into Holland\n"
    "        └── Liberalism Suppressed: No constitutional guarantees ➔ Set the stage for 1830 & 1848 explosions"
)

timeline_5 = [
    ("March 1814", "Treaty of Chaumont", "Britain, Russia, Austria, and Prussia bind themselves into Quadruple Alliance to defeat Napoleon."),
    ("May 1814", "First Treaty of Paris", "Generous peace terms with France under restored Louis XVIII; borders rolled back to 1792."),
    ("Sept 1814", "Opening of the Congress of Vienna", "Diplomats assemble in Vienna; festivities and diplomatic maneuvering commence."),
    ("March 1815", "The Hundred Days & Waterloo", "Napoleon escapes Elba; Vienna diplomats outlaw him; final defeat at Waterloo (June 18, 1815)."),
    ("June 1815", "Signing of the Final Act of Vienna", "Comprehensive 121-article treaty creating the new territorial map and German Confederation."),
    ("Nov 1815", "Second Treaty of Paris & Quadruple Alliance", "Slightly harsher terms for France: 700m franc indemnity, 5-year Allied occupation, 1790 borders."),
    ("1818", "Congress of Aix-la-Chapelle", "Allied occupation ends; France readmitted as an equal Great Power into the Pentarchy."),
    ("1830", "Belgian Independence & London Conference", "Concert of Europe displays pragmatic flexibility by recognizing an independent, neutral Belgium.")
]

essay_5 = [
    ("Introduction",
     "The Congress of Vienna (1814–1815) assembled the foremost statesmen of Europe—Prince Metternich of Austria, Lord Castlereagh of Britain, Tsar Alexander I of Russia, Prince Karl von Hardenberg of Prussia, and later Prince Talleyrand of France—to reconstruct a continent devastated by twenty-five years of revolutionary turmoil and Napoleonic warfare. Charged with establishing a durable international order, the peacemakers operated upon conservative principles: Balance of Power, Dynastic Legitimacy, and Territorial Compensation. In assessing its historical effectiveness between 1815 and 1848, the Vienna Settlement presents a stark dichotomy. As an instrument of international diplomacy, it was brilliantly effective, establishing the 'Concert of Europe' and preventing a general continental conflagration for nearly a century. However, in its ideological task of reconciling conservative monarchism with the emergent forces of change—specifically Liberalism and Nationalism—the Congress was a reactionary failure, sowing the seeds of revolutionary defiance that erupted in 1830 and 1848."),

    ("The Architectural Pillars: Balance of Power and Cordon Sanitaire",
     "The preeminent triumph of the Congress was the construction of a resilient, self-equilibrating balance of power. Guided by Castlereagh and Metternich, the delegates wisely resisted the temptation to dismember or permanently humiliate defeated France, recognizing that a stable peace required French inclusion. Under the First and Second Treaties of Paris, France was returned to its pre-revolutionary 1790 frontiers, ordered to pay a moderate indemnity, and subsequently readmitted into the ranks of the Great Powers at Aix-la-Chapelle (1818). To deter future French revanchism, the peacemakers established a robust perimeter of buffer states (cordon sanitaire): the Austrian Netherlands (Belgium) was fused with the Dutch Republic into the United Kingdom of the Netherlands; Prussia was deliberately awarded the wealthy Rhineland provinces; the Kingdom of Sardinia-Piedmont was strengthened with Genoa; and Switzerland was declared perpetually neutral."),

    ("Effectiveness in Maintaining International Peace (1815–1848)",
     "Between 1815 and 1848, the Congress of Vienna fulfilled its primary mandate of interstate stability with unprecedented success. It replaced the uncoordinated, balance-of-terror politics of the eighteenth century with an institutionalized Concert of Europe—the earliest prototype of modern international security governance. Through regular congresses, the Great Powers negotiated disputes collectively rather than resorting to unilateral warfare. Even when revolutionary tremors shook Europe in 1830, the Concert demonstrated pragmatic flexibility: rather than triggering a general war over Belgium’s rebellion against the Dutch crown, the Powers met at the London Conference of 1830–31, successfully recognizing an independent and permanently neutral Belgium under British guarantees. Consequently, Europe experienced no major multi-power war until the Crimean War of 1853, and no continent-wide catastrophe until 1914."),

    ("The Failure to Reconcile Forces of Change: The Denial of Nationalism",
     "Where the Congress failed catastrophically was in its cynical disregard for the emergent principle of Nationalism. The peacemakers treated populations and territories like interchangeable chess pieces on an eighteenth-century diplomatic board. In the quest for 'territorial compensation,' millions of people were bartered away into foreign domination without consultation. Poland was dismembered once more, with its core transformed into 'Congress Poland' under the direct autocracy of the Russian Tsar. Italy was callously partitioned into eight separate reactionary monarchies, with the wealthy provinces of Lombardy and Venetia directly annexed by Austria, leading Metternich to dismiss Italy as merely a 'geographical expression.' In Germany, the romantic aspiration for a unified nation was smothered by creating a loose, impotent German Confederation (Bund) of 39 sovereign states dominated by Austrian conservative hegemony."),

    ("The Suppression of Liberalism and Constitutional Aspirations",
     "Equally defective was the Congress's absolute refusal to accommodate Liberalism and the demand for representative constitutional governance. Operating under the banner of Dynastic Legitimacy, the delegates restored antiquated, autocratic royal houses across Europe: the Bourbon monarchy under Louis XVIII in France and Ferdinand VII in Spain, the House of Orange in the Netherlands, and reactionary dukes across the Italian peninsula. The peacemakers viewed popular sovereignty, freedom of the press, and parliamentary democracy not as legitimate political expressions, but as dangerous infectious diseases that threatened civil order. By erecting an international treaty system dedicated to preserving unyielding monarchical absolutism, the Vienna settlement left middle-class reformers and progressive intellectuals with no peaceful avenue for constitutional reform except through conspiratorial secret societies (Carbonari) and violent revolution."),

    ("Conclusion",
     "In ultimate historical evaluation, the Congress of Vienna was a masterpiece of geopolitical statecraft but an ideological failure. In preventing general European war, managing interstate rivalries, and reintegrating a defeated aggressor into a collaborative diplomatic framework, the peacemakers constructed one of the most durable peace settlements in modern history—far surpassing the punitive Treaty of Versailles in 1919. Yet, this interstate peace was purchased at the tragic price of domestic oppression. By ignoring the defining forces of the modern age—national self-determination and popular democratic participation—the Congress of Vienna achieved international stability without achieving domestic justice. It successfully postponed the collapse of the Old Order for a generation, but by locking the lid on a boiling social cauldron, it ensured that the explosion of 1848 would be all the more violent.")
]

quotes_5 = [
    ("Henry Kissinger", "A World Restored (1957)", "The Vienna settlement was the most successful peace conference in modern history; it created a period of peace that lasted almost a hundred years because it established a balance that was accepted as legitimate by all major powers."),
    ("Paul W. Schroeder", "The Transformation of European Politics", "Vienna did not merely restore an old balance of power; it replaced an unstable 18th-century system with a consensus-driven system of mutual restraint and international law."),
    ("Harold Nicolson", "The Congress of Vienna", "The mistake of the Vienna negotiators was that they looked backwards rather than forwards; they treated the peoples of Europe as mere pawns in an arithmetic game of dynastic compensation."),
    ("David Thomson", "Europe Since Napoleon", "The settlement failed to recognize the great twin forces of the 19th century—nationalism and liberalism—and by attempting to ignore them, it doomed Europe to a cycle of revolution."),
    ("Lord Castlereagh", "British Foreign Secretary", "The purpose of the Alliance was to inspire the states of Europe with that sense of security which can alone preserve them from the hazards of war.")
]

paper1_dir = r"H:\My Drive\PMS KPK 2022 notes\1. European History Paper 1 and Paper 2\Paper 1"

out_4 = os.path.join(paper1_dir, "Q4_Metternich_System_Suppression_and_Failure_2025_Master_Notes.docx")
out_5 = os.path.join(paper1_dir, "Q5_Congress_of_Vienna_Peace_vs_Change_2025_Master_Notes.docx")

generate_master_doc(meta_4, callout_4, map_4, timeline_4, essay_4, quotes_4, out_4, "MASTER STUDY NOTE & MODEL ANSWER &bull; METTERNICH SYSTEM & SUPPRESSION")
generate_master_doc(meta_5, callout_5, map_5, timeline_5, essay_5, quotes_5, out_5, "MASTER STUDY NOTE & MODEL ANSWER &bull; CONGRESS OF VIENNA & BALANCE OF POWER")
