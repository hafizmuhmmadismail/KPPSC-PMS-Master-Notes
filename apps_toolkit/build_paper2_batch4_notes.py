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
# QUESTION 13 (PAPER II): 2025 Q.7 (SOVIET DOMINATION & HUNGARIAN UPRISING)
# =========================================================================

meta_p2_13 = [
    ("EXAM QUESTIONS LINKED:", "2025 Q.7: 'Assess the impact of the Soviet Union's domination of Eastern Europe (1945\u20131960). How did it consolidate its power and respond to challenges like the Hungarian Uprising (1956)?'"),
    ("SYLLABUS TOPIC:", "Paper II \u2014 Section 8: Eastern Europe under Soviet Hegemony \u2014 Satellite States, De-Stalinization, Comecon & The 1956 Hungarian Revolution"),
    ("ALLOCATED MARKS / TIME:", "20 Marks | 35 Minutes | Standard Word Count: ~650\u2013700 words"),
    ("HISTORIOGRAPHICAL THESIS:", "Between 1945 and 1960, the Soviet Union transformed Eastern Europe into an ideological, political, and strategic buffer zone through 'salami tactics', economic exploitation via Comecon, and the military garrison of the Warsaw Pact; however, Nikita Khrushchev's 1956 'Secret Speech' denouncing Stalin inadvertently fractured satellite legitimacy, culminating in the Hungarian Uprising of 1956, where the brutal deployment of Soviet tanks proved that Moscow's hegemony rested not on socialist consensus, but on naked military coercion.")
]

callout_p2_13 = [
    "\u2022 The Mechanics of Subjugation (1945\u20131948): Red Army military occupation; M\u00e1ty\u00e1s R\u00e1kosi's famous 'Salami Tactics' (slicing away democratic opposition parties piece by piece through police intimidation, judicial murder, and rigged elections); creation of the Cominform (1947) to enforce Moscow's line; Tito's Yugoslavia expelled in 1948 for defying Stalin's command.",
    "\u2022 Economic Subordination: The Council for Mutual Economic Assistance (Comecon, 1949) integrated satellite economies into Soviet Five-Year Plans; forced heavy industrialization, collectivization, and unequal trade terms that drained Eastern European agricultural and mineral wealth to Moscow.",
    "\u2022 The Thaw and De-Stalinization (1953\u20131956): Stalin's death in March 1953 sparked the East Berlin worker uprising (crushed by tanks); at the 20th Party Congress (Feb 1956), Nikita Khrushchev delivered his explosive 'Secret Speech' denouncing Stalin's crimes, purges, and personality cult; sparked the 'Polish October' where W\u0142adys\u0142aw Gomu\u0142ka negotiated domestic autonomy while preserving the Warsaw Pact.",
    "\u2022 The Hungarian Revolution (October\u2013November 1956):",
    "  1. Mass student and worker protests in Budapest toppled Stalin's monumental statue; state secret police (AVH) fired on crowds; Hungarian army units defected, handing weapons to freedom fighters.",
    "  2. Reform Communist Premier Imre Nagy formed a multi-party coalition government, abolished one-party rule, and declared Hungarian neutrality, announcing withdrawal from the Warsaw Pact on November 1.",
    "  3. Soviet Response (Operation Whirlwind, Nov 4, 1956): Khrushchev deployed 17 Soviet divisions and 1,130 T-54 tanks into Budapest; brutal street fighting killed 20,000 Hungarians; 200,000 fled as refugees to Austria; Nagy was abducted and secretly executed in 1958; J\u00e1nos K\u00e1d\u00e1r was installed as puppet dictator.",
    "\u2022 Western Paralysis & The Suez Distraction: President Eisenhower and NATO refused to intervene militarily, terrified of triggering a global nuclear war over a nation inside the recognized Soviet sphere; simultaneous Anglo-French-Israeli invasion of Egypt during the Suez Crisis (Oct 29, 1956) completely fractured Western moral authority and diverted world attention."
]

map_p2_13 = """========================================================================================
SATELLITE CREATION (1945-48)          DE-STALINIZATION CRISIS (1956)       THE CRUSHING OF BUDAPEST
========================================================================================
[Red Army Liberation/Occupation]--> [Stalin Dies: March 1953]           --> [Oct 23: Budapest Riots]
['Salami Tactics' & Police Terror]--> [Khrushchev's 'Secret Speech']    --> [Imre Nagy Multi-Party Govt]
[Tito-Stalin Split: Purges 1948] --> [Polish October: Gomu\u0142ka Rises]   --> [Nagy Quits Warsaw Pact]
[Comecon & Warsaw Pact Created]  --> [Hungarian Students Revolt]        --> [Operation Whirlwind Tanks]
                                                                        --> [West Paralyzed by Suez 1956]
========================================================================================"""

timeline_p2_13 = [
    ("1945\u20131948", "Sovietization of Eastern Europe", "Bolshevik one-party dictatorships established across Poland, East Germany, Hungary, Romania, and Bulgaria."),
    ("June 1948", "The Tito-Stalin Split", "Yugoslavia expelled from Cominform; Stalin launches bloody show-trial purges of 'national communists' (Rajk, Sl\u00e1nsk\u00fd)."),
    ("March 5, 1953", "Death of Joseph Stalin", "Stalin dies; collective leadership under Malenkov, Molotov, and Khrushchev initiates an uncertain 'thaw.'"),
    ("June 17, 1953", "East Berlin Worker Uprising", "Soviet tanks crush strikes by 100,000 East German construction workers protesting increased labor quotas."),
    ("Feb 25, 1956", "Khrushchev's 'Secret Speech'", "Delivers speech 'On the Cult of Personality and Its Consequences' at 20th Party Congress, shattering Stalinist myth."),
    ("Oct 19\u201324, 1956", "Polish October", "Workers strike in Pozna\u0144; Khrushchev permits W\u0142adys\u0142aw Gomu\u0142ka to assume power in exchange for Warsaw Pact loyalty."),
    ("Oct 23, 1956", "Hungarian Uprising Begins", "Budapest students march in solidarity with Poland; pull down Stalin statue; demand democratic freedoms."),
    ("Oct 28\u2013Nov 1, 1956", "Nagy's Neutrality Declaration", "Imre Nagy announces multi-party democracy, releases Cardinal Mindszenty, and withdraws Hungary from Warsaw Pact."),
    ("Nov 4, 1956", "Soviet Tanks Attack Budapest", "Operation Whirlwind launched; 60,000 Soviet troops assault Budapest; crushing freedom fighters after 6 days of bloody combat."),
    ("June 16, 1958", "Execution of Imre Nagy", "Nagy hanged for high treason after secret trial; J\u00e1nos K\u00e1d\u00e1r enforces 'Goulash Communism' under Soviet supervision.")
]

essay_p2_13 = [
    ("Introduction: The Nature of the Soviet Imperial Glacis",
     "Between the conclusion of World War II in 1945 and the dawn of the 1960s, the Soviet Union constructed the most rigid, centralized imperial sphere in modern European history. Stretching from the Baltic Sea to the Black Sea, one hundred million citizens across Poland, East Germany, Czechoslovakia, Hungary, Romania, and Bulgaria were subordinated to Moscow's ideological, economic, and military hegemony. In the historiographical debate, while older Marxist historians framed Soviet control as the voluntary fraternal solidarity of the socialist camp against Western imperialism, modern post-Soviet archive scholarship by Charles Gati, Mark Kramer, and Tony Judt demonstrates that Soviet domination was a coercive, neo-imperial enterprise. Stalin subordinated Eastern Europe to serve as an ideological buffer (glacis) protecting the Soviet motherland. However, when Nikita Khrushchev unleashed the forces of De-Stalinization in 1956, the brittle legitimacy of the satellite regimes shattered, culminating in the Hungarian Revolution\u2014a tragic conflagration that exposed the raw military violence upon which Soviet rule ultimately depended."),

    ("The Mechanics of Consolidation: Salami Tactics and Totalitarian Uniformity",
     "The imposition of Soviet control between 1945 and 1948 was executed through a masterful, deceptive strategy described by Hungarian communist leader M\u00e1ty\u00e1s R\u00e1kosi as 'salami tactics'\u2014slicing away democratic opposition slice by slice. Following the liberation of Eastern Europe by the Red Army, the Soviets initially tolerated broad anti-fascist coalition governments (National Fronts). However, local communist parties, directed by Moscow, systematically seized the crucial levers of state power: the Ministry of the Interior, the secret police, state radio, and the judiciary.\n\n"
     "Non-communist politicians were gradually intimidated, arrested, or exiled. In Poland, peasant leader Stanis\u0142aw Miko\u0142ajczyk was forced to flee for his life in 1947 after rigged elections; in Bulgaria, agrarian leader Nikola Petkov was framed for treason and hanged; in Czechoslovakia, the February 1948 coup d'\u00e9tat toppled the democratic government of Edvard Bene\u0161 and murdered Foreign Minister Jan Masaryk. Following the 1948 Tito-Stalin split\u2014wherein Marshal Josip Broz Tito successfully defended Yugoslavia's national independence against Moscow's diktat\u2014Stalin unleashed ferocious purges across the satellite states, staging show trials and executing 'Titoist national-deviationists' like L\u00e1szl\u00f3 Rajk in Hungary and Rudolf Sl\u00e1nsk\u00fd in Czechoslovakia to enforce absolute subservience."),

    ("Economic Exploitation and Military Subjection: Comecon and the Warsaw Pact",
     "Soviet political dominance was matched by structural economic exploitation. In January 1949, Moscow established the Council for Mutual Economic Assistance (Comecon) as the communist answer to the Marshall Plan. Satellite economies were forced to replicate the Stalinist model: rapid collectivization of peasant agriculture, abolition of private enterprise, and mandatory Five-Year Plans focused entirely on heavy metallurgy and coal extraction to feed the Soviet military machine. Comecon enforced unequal, bilateral barter trade agreements that drained satellite resources: Poland was forced to sell coal to the USSR at one-tenth of the world market price, while Romania surrendered its oil reserves to joint Soviet-Romanian enterprises (SovRoms).\n\n"
     "Militarily, the satellite states were locked into Moscow's strategic command through bilateral treaties, finalized in May 1955 with the creation of the Warsaw Pact. Soviet troops remained permanently garrisoned across Poland, East Germany, Hungary, and Romania, ensuring that any domestic threat to communist party supremacy would be met by immediate military intervention."),

    ("The Shock of De-Stalinization: Khrushchev's Secret Speech (1956)",
     "The death of Joseph Stalin on March 5, 1953, dealt an existential psychological shock to the Soviet empire. Three months later, on June 17, 1953, East Berlin construction workers staged mass strikes against increased production quotas, which rapidly escalated into a nationwide anti-communist uprising crushed only by Soviet T-34 tanks.\n\n"
     "The definitive crisis erupted at the 20th Congress of the Communist Party of the Soviet Union in February 1956. In an extraordinary four-hour nocturnal address known as the 'Secret Speech,' newly ascended First Secretary Nikita Khrushchev delivered a devastating indictment of Stalin's crimes, exposing the paranoia, mass purges of loyal communists, fabrication of show-trial confessions, and the megalomania of the 'Cult of Personality.' While Khrushchev intended De-Stalinization to cleanse the communist system from within, the speech leaked to the Western press and detonated across Eastern Europe. If Stalin was a criminal tyrant, then the puppet leaders he had installed in Warsaw, Budapest, and Prague were illegitimate accomplices. In June 1956, worker riots in Pozna\u0144, Poland, escalated into the 'Polish October': Khrushchev flew to Warsaw, prepared to invade, but ultimately compromised by permitting reform communist W\u0142adys\u0142aw Gomu\u0142ka to assume power in exchange for Gomu\u0142ka's solemn pledge that Poland would remain a loyal member of the Warsaw Pact."),

    ("The Hungarian Volcano: Revolution, Neutrality, and Soviet Tanks (October\u2013November 1956)",
     "In Hungary, the Polish compromise ignited an uncontrollable revolution. On October 23, 1956, twenty thousand students in Budapest staged a solidarity march, demanding the return of reformist Premier Imre Nagy, the withdrawal of Soviet troops, freedom of the press, and democratic elections. When the hated State Protection Authority (AVH, the secret police) opened fire on demonstrators outside the radio building, the protest transformed into an armed insurrection. Hungarian soldiers defected en masse, handing weapons to workers and students; freedom fighters attacked AVH headquarters, hunted down secret police agents, and toppled the monumental bronze statue of Stalin, decapitating it in the city center.\n\n"
     "To contain the conflagration, the Central Committee appointed Imre Nagy as Prime Minister. However, propelled by popular revolutionary momentum, Nagy moved far beyond Khrushchev's tolerated limits of reform: on October 30, Nagy abolished the one-party monopoly, established a multi-party coalition government, released political prisoners (including Catholic Primate Cardinal J\u00f3zsef Mindszenty), and demanded the immediate withdrawal of Soviet troops. When Soviet armored reinforcements crossed the Ukrainian border on November 1, Nagy took the fateful, revolutionary leap: he announced Hungary's unilateral withdrawal from the Warsaw Pact, proclaimed Hungarian neutrality, and appealed to the United Nations and the Western Great Powers for protection."),

    ("Operation Whirlwind and the Geopolitical Limits of Containment",
     "Khrushchev and the Soviet Presidium could not permit Hungary's defection. The withdrawal of Hungary would shatter the Warsaw Pact, leave the Soviet strategic buffer fractured, and trigger a domino collapse across the satellite empire. In the pre-dawn hours of November 4, 1956, the Soviet military launched 'Operation Whirlwind.' Seventeen Soviet mechanized divisions, 1,130 tanks, and heavy artillery under the supreme command of Marshal Ivan Konev assaulted Budapest.\n\n"
     "Hungarian freedom fighters\u2014many of them teenagers using Molotov cocktails and hunting rifles\u2014waged heroic, suicidal street warfare against Soviet tanks amidst the ruins of Budapest. Over six days of ferocious combat, an estimated twenty thousand Hungarians and twelve hundred Soviet soldiers were killed. Two hundred thousand Hungarians fled across the border into Austria as refugees. Imre Nagy sought sanctuary in the Yugoslav embassy; deceived by a safe-conduct guarantee issued by Soviet-installed puppet J\u00e1nos K\u00e1d\u00e1r, Nagy was abducted by Soviet agents, transported to Romania, and executed by hanging following a secret trial in June 1958.\n\n"
     "The Hungarian tragedy vividly exposed the limits of Western rhetoric. For years, the Eisenhower administration and Radio Free Europe had broadcast promises of 'Rollback' and the liberation of captive nations. Yet, when the Hungarians rose, the United States did not lift a finger. President Dwight D. Eisenhower recognized that military intervention in Hungary meant World War III and thermonuclear destruction over a nation recognized as falling within Moscow's sphere of influence since Yalta. Furthermore, the Western alliance was completely paralyzed by the simultaneous Suez Crisis: on October 29, 1956, Britain, France, and Israel invaded Egypt to seize the Suez Canal, infuriating Washington, shattering NATO unity, and granting Khrushchev a moral smokescreen to crush Budapest without unified Western diplomatic opposition."),

    ("Conclusion: The Moral Death of Soviet Ideology",
     "In conclusion, the Soviet domination of Eastern Europe between 1945 and 1960 achieved its immediate strategic objective: it preserved an impregnable military barrier safeguarding the Soviet heartland and locked one hundred million Europeans into a command economy. However, the Hungarian Uprising of 1956 was a mortal blow to the ideological soul of international communism. By deploying Soviet tanks to butcher Hungarian factory workers who flew the national flag with the hammer and sickle cut out of its center, the Soviet Union destroyed the myth that it represented the liberation of the working class. Eastern European communism ceased to be an inspiring utopian ideal, becoming instead an alien, bankrupt system of military occupation. While J\u00e1nos K\u00e1d\u00e1r's 'Goulash Communism' stabilized Hungary in the 1960s through consumer concessions, 1956 established the irreversible historical trajectory that culminated thirty-three years later in 1989, when the removal of Soviet military coercion caused the entire satellite empire to dissolve overnight.")
]

quotes_p2_13 = [
    ("Tony Judt", "Postwar: A History of Europe Since 1945 (2005)", "'The Hungarian Revolution of 1956 dealt a mortal wound to the Soviet empire... After Budapest, no one could pretend that communism in Eastern Europe was maintained by anything other than the naked threat of Soviet military violence.'"),
    ("Charles Gati", "Failed Illusions: Moscow, Washington, Budapest, and 1956", "'Imre Nagy was a tragic, contradictory figure: a loyal communist who, when swept along by the heroic passion of his people, chose to die as a Hungarian patriot rather than live as a Soviet collaborator.'"),
    ("Zbigniew Brzezinski", "The Soviet Bloc: Unity and Conflict", "'The Soviet empire in Eastern Europe was founded on an irreconcilable structural contradiction: Moscow demanded absolute ideological subservience, while the peoples of the region possessed distinct national cultures and historic affinities with Western civilization.'"),
    ("Imre Nagy (Nov 4, 1956)", "Final Radio Broadcast to the World", "'This is Imre Nagy, Chairman of the Council of Ministers of the Hungarian People's Republic. Today at daybreak Soviet forces launched an attack against our capital with the obvious intention of overthrowing the legal democratic Hungarian government... Our troops are in combat! Help Hungary! Help us!'")
]

# =========================================================================
# QUESTION 14 (PAPER II): 2022 Q.7 / 2018 Q.8 (UN VS LEAGUE OF NATIONS)
# =========================================================================

meta_p2_14 = [
    ("EXAM QUESTIONS LINKED:", "2022 Q.7: 'Evaluate the creation, structure, and early peacekeeping challenges of the United Nations Organization (UNO) compared to the League of Nations.'\n2018 Q.8: 'Discuss the establishment and importance of the United Nations Organization.'"),
    ("SYLLABUS TOPIC:", "Paper II \u2014 Section 8: The United Nations Organization (UNO) \u2014 Origins, Structure, Comparison with League of Nations & Early Peacekeeping"),
    ("ALLOCATED MARKS / TIME:", "20 Marks | 35 Minutes | Standard Word Count: ~650\u2013700 words"),
    ("HISTORIOGRAPHICAL THESIS:", "The creation of the United Nations Organization at the San Francisco Conference in 1945 represented the conscious, realistic reconstruction of international multilateralism designed to correct the fatal structural defects of the League of Nations; by abandoning the paralyzing requirement of unanimity, establishing an executive Security Council endowed with Great Power veto authority and Chapter VII military enforcement teeth, and anchoring the permanent leadership of the United States and the Soviet Union, the UN forged a durable institutional apparatus that avoided the catastrophic fate of the League, even as its collective security mechanisms were recurrently constrained by Cold War superpower deadlock.")
]

callout_p2_14 = [
    "\u2022 Wartime Genesis of the UN: Atlantic Charter (August 1941, Roosevelt & Churchill); Declaration by United Nations (Jan 1942, 26 nations); Moscow Conference (Oct 1943, Big Four pledge general international body); Dumbarton Oaks Conference (1944, drafting institutional charter); Yalta Conference (Feb 1945, resolving the Security Council voting veto formula); San Francisco Conference (April 25\u2013June 26, 1945, 50 nations sign the UN Charter).",
    "\u2022 Structural Comparison: UN vs. League of Nations:",
    "  1. Assembly: League Assembly required unanimity for all substantive action; UN General Assembly operates on a democratic two-thirds majority for vital questions (Article 18), functioning as a universal global parliament.",
    "  2. Executive Council: League Council required unanimous consent of all members; UN Security Council (15 members: Big Five permanent [US, USSR, UK, France, China] + 10 rotating non-permanent) operates on a qualified majority of 9 votes, incorporating the P5 Veto ('Yalta Formula') to prevent the UN from suicidal military action against a nuclear superpower.",
    "  3. Enforcement Teeth: League lacked military machinery (Article 16 economic sanctions only); UN Charter Chapter VII grants binding legal authority: Article 41 (economic sanctions) and Article 42 (military intervention by land, sea, and air via Military Staff Committee).",
    "  4. Socioeconomic Breadth: The League treated social matters as secondary; the UN created ECOSOC, specialized global agencies (WHO, UNESCO, ILO, FAO), and the Bretton Woods institutions (World Bank, IMF).",
    "\u2022 Why the UN Succeeded Where the League Failed: Universal Great Power participation (the US was the primary architect and financial patron; the USSR remained an active member); institutional resilience (no superpower walked out during crises); recognition of power realities via the veto.",
    "\u2022 Early Peacekeeping Innovations: Cold War paralysis (frequent Soviet vetoes) led to creative adaptations: (1) Uniting for Peace Resolution (377A, 1950) allowing General Assembly action during the Korean War, and (2) Dag Hammarskj\u00f6ld and Lester B. Pearson's invention of armed, neutral 'Chapter Six and a Half' Peacekeeping during the 1956 Suez Crisis (UNEF I)."
]

map_p2_14 = """========================================================================================
LEAGUE DEFICIENCIES (1919)           WARTIME PLANNING (1941-45)        THE MODERN UN (1945-1960)
========================================================================================
[Paralyzing Unanimity Rule]       --> [Atlantic Charter & Dumbarton Oaks]--> [General Assembly: 2/3 Vote]
[US Absent & Pariah Walkouts]     --> [Yalta Summit: Big Five Veto]     --> [Security Council: P5 Veto]
[Toothless Economic Sanctions]    --> [San Francisco Conference (1945)] --> [Chapter VII Military Teeth]
[No Standing Military Structure]  --> [Universal Great Power Buy-In]    --> [ECOSOC & Human Rights (1948)]
                                                                        --> [Invention of Peacekeeping:
                                                                             Suez Crisis UNEF I (1956)]
========================================================================================"""

timeline_p2_14 = [
    ("Aug 14, 1941", "The Atlantic Charter", "Roosevelt and Churchill formulate liberal war aims, promising a permanent system of general security."),
    ("Jan 1, 1942", "Declaration by United Nations", "26 nations endorse Atlantic Charter; first official use of term 'United Nations' coined by FDR."),
    ("Aug\u2013Oct 1944", "Dumbarton Oaks Conference", "US, UK, USSR, and China meet in Washington to draft structure of General Assembly and Security Council."),
    ("Feb 11, 1945", "Yalta Voting Compromise", "Big Three agree that substantive Security Council decisions require affirmative vote of 7 (now 9) members including all 5 permanent members."),
    ("April 25\u2013June 26, 1945", "San Francisco Conference", "50 nations debate and unanimously adopt the 111 articles of the Charter of the United Nations."),
    ("Oct 24, 1945", "UN Charter Enters into Force", "Ratified by the Big Five and majority of nations; officially establishes the United Nations (United Nations Day)."),
    ("Jan 10, 1946", "First General Assembly Session", "51 nations convene in London; Paul-Henri Spaak elected first Assembly President; Trygve Lie elected Secretary-General."),
    ("Dec 10, 1948", "Universal Declaration of Human Rights", "Chaired by Eleanor Roosevelt; General Assembly adopts landmark declaration establishing universal human dignity."),
    ("Nov 3, 1950", "Uniting for Peace Resolution", "Resolution 377A passed to bypass Soviet vetoes and authorize collective military action in the Korean War."),
    ("Nov 1956", "Invention of UN Peacekeeping (UNEF I)", "Dag Hammarskj\u00f6ld deploys 6,000 neutral blue-helmet troops to supervise British-French-Israeli withdrawal from Suez.")
]

essay_p2_14 = [
    ("Introduction: The Phoenix from the Ashes of Geneva",
     "On October 24, 1945, the Charter of the United Nations entered into legal force, inaugurating a revolutionary chapter in the history of international governance. Formulated in the catastrophic shadow of World War II, the United Nations Organization (UNO) was not an idealistic fantasy conceived in isolation, but the hard-headed, realistic reconstruction of international multilateralism designed to correct the catastrophic structural and political failures that had condemned the League of Nations to irrelevance. In the historiographical analysis of scholars such as Paul Kennedy, Inis Claude, and Mark Mazower, the creation of the United Nations represented a masterful synthesis of Wilsonian universalism and Realist balance-of-power diplomacy. By replacing the paralyzing rule of unanimity with an executive Security Council, incorporating the Great Power veto, providing military enforcement teeth under Chapter VII, and ensuring the permanent institutional leadership of both the United States and the Soviet Union, the architects of San Francisco created an enduring global institution that avoided the collapse of its predecessor and preserved humanity from a third world war."),

    ("The Lessons of Failure: Why the League of Nations Collapsed",
     "The institutional architecture of the United Nations was forged through a systematic autopsy of the League of Nations' structural vulnerabilities. The League had suffered from four fatal defects that made it helpless to resist the aggression of imperial Japan, fascist Italy, and Nazi Germany in the 1930s. First, the League was crippled by the fatal absence of the United States, whose Senate rejected ratification in 1919\u20131920, depriving the body of the world's greatest economic and naval titan. Second, the League was structured around the absurd procedural rule of unanimity (Articles 5 and 15), meaning that any single member state could paralyze collective action.\n\n"
     "Third, the League possessed no military enforcement mechanism: Article 16 provided only for voluntary economic sanctions, which Britain and France refused to enforce effectively during the 1935 Abyssinian crisis. Fourth, the League was perceived as a punitive 'league of victors' designed to enforce the reviled Versailles Treaty; whenever rogue powers were challenged, they simply resigned: Japan in 1933, Germany in 1933, and Italy in 1937."),

    ("Wartime Summitry and the Genesis of the UN (1941\u20131945)",
     "Unlike the League, which was hastily appended to the contentious Treaty of Versailles in 1919, the United Nations was meticulously planned across four years of inter-Allied wartime diplomacy. The ideological foundation was laid in August 1941 when Franklin D. Roosevelt and Winston Churchill drafted the Atlantic Charter, pledging a 'wider and permanent system of general security.' In January 1942, twenty-six nations signed the 'Declaration by United Nations,' cementing the wartime coalition against the Axis powers.\n\n"
     "The institutional blueprint was negotiated at the Dumbarton Oaks Conference in Washington (August\u2013October 1944) by the United States, Great Britain, the Soviet Union, and the Republic of China. The critical stumbling block\u2014the voting procedure in the executive council\u2014was resolved by Roosevelt, Churchill, and Stalin at the Yalta Conference in February 1945. Finally, from April 25 to June 26, 1945, delegates from fifty sovereign nations representing eighty percent of the world's population assembled at the San Francisco Opera House. In an atmosphere of intense debate, the 111 articles of the United Nations Charter were finalized and unanimously adopted on June 26, 1945."),

    ("Structural Innovations: The Security Council, the Veto, and Chapter VII",
     "The supreme structural breakthrough of the UN Charter was its pragmatic accommodation of global power realities across three key innovations:\n\n"
     "1. The General Assembly vs. The Executive Security Council: The UN established a clear functional division of labor. The General Assembly operates as a universal democratic forum where all member states possess one equal vote, passing non-binding recommendations by a two-thirds majority on vital matters (Article 18). However, actual executive authority was concentrated in the Security Council, which was entrusted with primary responsibility for the maintenance of international peace and security.\n\n"
     "2. The 'Big Five' and the Power of the Veto: To avoid the unanimity trap of the League, substantive decisions in the Security Council require a qualified majority of nine votes (originally seven). However, this majority must include the concurring votes of the five Permanent Members (P5: United States, Soviet Union/Russia, Great Britain, France, and China). The veto power (the 'Yalta Formula') was fiercely criticized by smaller nations at San Francisco. Yet, as Inis Claude correctly observes, the veto was the essential anchor of the UN: it acknowledged that collective security against a nuclear superpower was an impossibility that would result in World War III. The veto ensured that the UN would act only when the Great Powers were in agreement, preserving the institution from suicidal confrontation.\n\n"
     "3. Chapter VII Enforcement Teeth: Unlike the toothless Covenant of Geneva, Chapter VII of the UN Charter (Articles 39\u201351) provides the Security Council with binding legal authority. Under Article 39, the Council can identify threats to the peace; under Article 41, it can mandate comprehensive, legally binding economic and diplomatic sanctions; and under Article 42, it can authorize armed military intervention by air, sea, or land forces provided by member states through the Military Staff Committee. When North Korea invaded South Korea in June 1950, the Security Council\u2014capitalizing on a temporary Soviet boycott over the representation of Taiwan\u2014invoked Chapter VII to authorize a sixteen-nation military coalition commanded by General Douglas MacArthur to repel the aggression."),

    ("Socioeconomic Breadth: Beyond Pure Security",
     "A profound qualitative difference between the League and the United Nations was the UN's recognition that international peace is indivisible from economic stability, social justice, and human rights. Article 55 of the Charter established the Economic and Social Council (ECOSOC) to coordinate global development, public health, and human welfare.\n\n"
     "Under the visionary leadership of Eleanor Roosevelt, the UN Commission on Human Rights drafted the monumental 'Universal Declaration of Human Rights' (UDHR), adopted by the General Assembly on December 10, 1948, establishing the universal baseline for human dignity, civil liberties, and fundamental freedoms in international law. Furthermore, the UN integrated specialized agencies that revolutionized global governance: the World Health Organization (WHO), the United Nations Educational, Scientific and Cultural Organization (UNESCO), the International Labor Organization (ILO), the Food and Agriculture Organization (FAO), and the United Nations Children's Fund (UNICEF). In the economic sphere, the UN coordinated with the Bretton Woods institutions\u2014the International Monetary Fund (IMF) and the World Bank\u2014to ensure international monetary stability and post-war reconstruction."),

    ("Cold War Realities and the Invention of Peacekeeping (1956)",
     "As the Cold War escalated, the idyllic vision of great-power harmony dissolved. Between 1945 and 1955, the Soviet Union cast over eighty vetoes in the Security Council, paralyzing the Chapter VII collective security mechanism. Confronted by this structural deadlock, the United Nations demonstrated remarkable institutional flexibility and innovation. In November 1950, the General Assembly adopted the 'Uniting for Peace' Resolution (Resolution 377A), establishing that if the Security Council is paralyzed by a veto in the face of aggression, the General Assembly can immediately consider the matter and recommend collective measures, including armed force.\n\n"
     "The greatest institutional breakthrough occurred during the Suez Crisis of 1956. When Britain and France vetoed Security Council resolutions demanding their withdrawal from Egypt, UN Secretary-General Dag Hammarskj\u00f6ld and Canadian diplomat Lester B. Pearson invented an entirely new diplomatic concept: United Nations Peacekeeping (often designated 'Chapter Six and a Half'). Operating under the General Assembly, the UN deployed the First United Nations Emergency Force (UNEF I)\u2014six thousand neutral, blue-helmeted soldiers drawn from non-permanent powers (Canada, India, Sweden, Colombia)\u2014to supervise the ceasefire and facilitate the withdrawal of British, French, and Israeli forces. Peacekeeping was not designed to defeat an aggressor, but to act as a neutral buffer with the consent of the host state, defusing flashpoints before they triggered superpower nuclear confrontation."),

    ("Conclusion: The Enduring Parliament of Man",
     "In conclusion, the United Nations Organization stands as the most resilient, indispensable multilateral institution in world history. By systematically learning the lessons of the League of Nations' tragic demise, the founders of 1945 constructed an organization that reconciled the idealism of universal human rights with the unvarnished realities of great-power geopolitics. While the UN could not prevent the Cold War or eliminate regional conflicts, it prevented the ultimate catastrophe: a third world war between nuclear superpowers. As second Secretary-General Dag Hammarskj\u00f6ld famously summarized its realistic mandate: 'The United Nations was not created in order to bring us to heaven, but in order to save us from hell.' Through its universal membership, specialized agencies, and peacekeeping missions, the UN remains the irreplaceable parliament of humanity.")
]

quotes_p2_14 = [
    ("Paul Kennedy", "The Parliament of Man: The Past, Present, and Future of the United Nations (2006)", "'The United Nations survived the Cold War precisely because its founders were hard-headed realists who built the Great Power veto into the Charter... The UN is not a world government, but an indispensable scaffolding for global survival.'"),
    ("Inis L. Claude Jr.", "Swords into Plowshares (1971)", "'The veto power in the Security Council is not the cause of UN weakness; it is the reflection of the profound political reality that peace among the Great Powers cannot be enforced by majority votes.'"),
    ("Dag Hammarskj\u00f6ld", "Second UN Secretary-General", "'The United Nations was not created in order to bring us to heaven, but in order to save humanity from hell.'"),
    ("Franklin D. Roosevelt (1944)", "Address on the United Nations", "'We have learned that we cannot live alone, at peace; that our own well-being is dependent on the well-being of other nations, far away... We must build a structure of peace that is permanent.'")
]

# =========================================================================
# QUESTION 15 (PAPER II): 2025 Q.8 (EUROPEAN INTEGRATION: SCHUMAN TO ROME)
# =========================================================================

meta_p2_15 = [
    ("EXAM QUESTIONS LINKED:", "2025 Q.8: 'To what extent did the success of the Common Market inspire later movements toward deeper European integration, culminating in the European Union?'"),
    ("SYLLABUS TOPIC:", "Paper II \u2014 Section 8: Post-War European Integration \u2014 The Schuman Plan, ECSC (1951), Treaty of Rome (1957) & The Common Market"),
    ("ALLOCATED MARKS / TIME:", "20 Marks | 35 Minutes | Standard Word Count: ~650\u2013700 words"),
    ("HISTORIOGRAPHICAL THESIS:", "The post-WWII European integration movement was a revolutionary departure from centuries of destructive nationalism, engineered by Jean Monnet and Robert Schuman to eliminate the material and military possibility of Franco-German war through functional supranational pooling; by establishing the European Coal and Steel Community (ECSC) in 1951 and consolidating it through the 1957 Treaty of Rome into the European Economic Community (EEC / Common Market), integration delivered unprecedented economic prosperity, demonstrating that sectoral economic unification created an irresistible neo-functionalist 'spillover' that propelled Europe toward the European Union.")
]

callout_p2_15 = [
    "\u2022 The Core Dilemma: Following the devastation of World War II, Western Europe faced a strategic paradox: West Germany's heavy industrial core (the Ruhr and Saar basins) had to be revitalized to rebuild the European economy and resist Soviet expansion, yet a revived Germany terrified France with the specter of a fourth invasion in eighty years (1870, 1914, 1940).",
    "\u2022 The Functionalist Genius of Jean Monnet & Robert Schuman: Rather than pursuing an impossible, utopian federal political union, integrate specific key industrial sectors vital for waging war: Coal and Steel.",
    "\u2022 The Schuman Declaration (May 9, 1950): French Foreign Minister Robert Schuman proposed placing the entire Franco-German production of coal and steel under a common supranational 'High Authority', open to other European nations; famously declared that this would make war between France and Germany 'not merely unthinkable, but materially impossible' (celebrated today as Europe Day).",
    "\u2022 The European Coal and Steel Community (ECSC, Treaty of Paris 1951): Founded by 'The Six' (France, West Germany, Italy, Belgium, Netherlands, Luxembourg); abolished customs duties on coal and steel; created the first supranational institutions: the High Authority (precursor to European Commission), Common Assembly (European Parliament), and Court of Justice.",
    "\u2022 The Messina Conference (1955) & The Treaty of Rome (March 25, 1957): Following the collapse of the European Defense Community (EDC), the Six executed the 'Relance Europ\u00e9enne' at Messina, signing the Treaties of Rome establishing:",
    "  1. The European Economic Community (EEC / Common Market): Phased abolition of all internal tariffs, establishment of a Common External Tariff (CET), the 'Four Freedoms' (free movement of goods, services, capital, and labor), and the Common Agricultural Policy (CAP).",
    "  2. Euratom: Pooling atomic energy research for peaceful purposes.",
    "\u2022 Neo-Functionalist 'Spillover' (Ernst B. Haas): Sectoral economic integration created technical and economic pressures that compelled integration into adjacent sectors (coal/steel -> general trade -> agriculture -> single market -> monetary union / Euro -> European Union under the 1992 Maastricht Treaty).",
    "\u2022 Britain's Historic Blunder: Detached from the continent, Britain refused to join the ECSC in 1951 or the EEC in 1957; established the rival, weak European Free Trade Association (EFTA) in 1960; recognizing the EEC's phenomenal economic boom, Britain applied in 1961, only to be humiliated by Charles de Gaulle's vetoes in 1963 and 1967."
]

map_p2_15 = """========================================================================================
POST-WAR PARALYSIS                FUNCTIONAL SCHUMAN PLAN (1950)       THE COMMON MARKET & BEYOND
========================================================================================
[Devastated Post-1945 Europe]  --> [Monnet: Pool Coal & Steel]       --> [Treaty of Rome (March 1957)]
[Franco-German Historical Hate]--> [Schuman Declaration (May 1950)]   --> [EEC: The Common Market]
[Need to Rebuild Ruhr Industry]--> [ECSC Treaty of Paris (1951)]     --> [The Four Freedoms & CAP]
[Soviet Cold War Threat]       --> [The Six Founding Nations]        --> [Neo-Functionalist 'Spillover':
                               --> [EDC Military Failure (1954)]          Single Market -> Maastricht (EU)]
========================================================================================"""

timeline_p2_15 = [
    ("May 9, 1950", "The Schuman Declaration", "French Foreign Minister Robert Schuman proposes placing Franco-German coal and steel under a joint High Authority."),
    ("April 18, 1951", "Treaty of Paris (ECSC)", "The Six (France, West Germany, Italy, Belgium, Netherlands, Luxembourg) sign treaty creating the ECSC."),
    ("Aug 30, 1954", "Failure of the EDC", "French National Assembly rejects the European Defense Community treaty, terminating premature military federation."),
    ("June 1\u20133, 1955", "The Messina Conference", "Foreign ministers of the Six resolve to expand integration from coal and steel to a full customs union."),
    ("March 25, 1957", "The Treaties of Rome Signed", "The Six sign treaties establishing the European Economic Community (EEC) and Euratom on the Capitoline Hill."),
    ("Jan 1, 1958", "EEC Officially Established", "Treaties of Rome come into force; Walter Hallstein becomes first President of the European Commission."),
    ("May 1960", "Britain Establishes EFTA", "United Kingdom organizes seven non-EEC nations into the European Free Trade Association (EFTA) in Stockholm."),
    ("Aug 1961", "Britain Applies to Join EEC", "Prime Minister Harold Macmillan recognizes Common Market economic dynamism and formally applies for membership."),
    ("Jan 14, 1963", "De Gaulle's First Veto", "French President Charles de Gaulle vetoes British entry into the Common Market, accusing Britain of US vassalage."),
    ("Feb 7, 1992", "The Maastricht Treaty", "The Treaty on European Union signed, transforming the European Community into the EU and creating the Euro.")
]

essay_p2_15 = [
    ("Introduction: The Paradigm Shift of European Unity",
     "The movement toward European integration following the catastrophe of the Second World War stands as one of the most remarkable acts of political and economic statesmanship in modern world history. For centuries, the European continent had been the cockpit of fratricidal nationalist conflict, culminating in three devastating Franco-German wars within seventy years (1870, 1914, 1939). In 1945, Europe lay in physical, demographic, and moral ruin, squeezed between the two emergent superpowers of the Cold War. In the historiographical debate, while neo-functionalist scholars like Ernst B. Haas analyzed integration as an incremental, self-perpetuating technocratic process characterized by functional 'spillover', liberal intergovernmentalists like Andrew Moravcsik and Alan Milward argue that integration was a calculated strategy by sovereign nation-states to rescue their domestic economic legitimacy and manage interdependence. In reality, post-war integration was an extraordinary fusion of vision and pragmatism: engineered by French planners Jean Monnet and Robert Schuman, consolidated in the European Coal and Steel Community (ECSC) in 1951, and institutionalized in the 1957 Treaty of Rome (the Common Market), integration created an economic engine that permanently pacified the Franco-German frontier and laid the irreversible foundation for the modern European Union."),

    ("The Post-War Dilemma: The German Problem and the Vision of Jean Monnet",
     "The primary catalyst for European integration centered upon the intractable post-war 'German Question.' To rebuild Western European prosperity and provide a viable defense against the Soviet Union, the industrial potential of West Germany\u2014specifically the coal-rich Ruhr and the iron-producing Saar basins\u2014had to be revived. However, any unilateral German economic recovery terrified France, which was haunted by memories of Nazi occupation and demanded international control or detachment of the Ruhr. Traditional diplomatic solutions had failed: the Versailles Treaty had provoked German revanchism, while post-1945 punitive dismantling paralyzed the continental economy.\n\n"
     "The breakthrough was conceived by Jean Monnet, the head of the French General Planning Commissariat, and political leader Robert Schuman, the French Foreign Minister who had lived through two German annexations of his native Lorraine. Monnet and Schuman recognized that an immediate, comprehensive federal United States of Europe was a political impossibility. Instead, they championed 'Functionalism': pooling sovereign control over specific, vital economic sectors to build practical solidarity from the bottom up."),

    ("The Schuman Declaration and the ECSC (1950\u20131951)",
     "On May 9, 1950, in the historic Salon de l'Horloge of the Quai d'Orsay in Paris, Robert Schuman announced the 'Schuman Declaration' (celebrated today as Europe Day). Schuman proposed placing the entire Franco-German production of coal and steel under a common, supranational 'High Authority' within an organization open to other European countries. The economic and strategic genius of the plan was profound: coal and steel were the essential raw materials of heavy industry and armaments manufacture. By pooling these resources under an independent supranational body, Schuman famously declared that 'any war between France and Germany becomes not merely unthinkable, but materially impossible.'\n\n"
     "West German Chancellor Konrad Adenauer embraced the proposal immediately, recognizing it as the historic vehicle to restore German sovereignty, rehabilitate Germany as an equal partner in the Western community, and prevent punitive French reparations. On April 18, 1951, six nations\u2014France, West Germany, Italy, Belgium, the Netherlands, and Luxembourg (the historic 'Six')\u2014signed the Treaty of Paris, establishing the European Coal and Steel Community (ECSC). The ECSC eliminated all import duties, quotas, and discriminatory transport rates on coal, iron ore, and steel among the Six. More decisively, it erected the institutional prototype of the European Union: an executive High Authority possessing sovereign powers to bind member states, a Common Assembly of parliamentarians, a Special Council of Ministers, and a Court of Justice to adjudicate European law."),

    ("The Messina Relaunch and the Treaties of Rome (1955\u20131957)",
     "The momentum of integration suffered a severe crisis in August 1954 when the French National Assembly rejected the European Defense Community (EDC) treaty, which had attempted to create a supranational European army. The failure of the EDC demonstrated that integration could not proceed through premature military or political leaps; it had to remain grounded in economic realities.\n\n"
     "The historic breakthrough occurred in June 1955 at the Messina Conference in Sicily. Led by Belgian Foreign Minister Paul-Henri Spaak, the foreign ministers of the Six resolved to expand the functional approach from coal and steel to the entire economy by creating a general common market. The negotiations culminated on March 25, 1957, on the Capitoline Hill in Rome, where the Six signed the historic Treaties of Rome, establishing the European Economic Community (EEC, commonly known as the Common Market) and the European Atomic Energy Community (Euratom).\n\n"
     "The Treaty of Rome was a monumental diplomatic achievement. It established a twelve-year transition period to achieve four historic objectives: (1) The complete elimination of all internal customs tariffs and trade quotas among the Six; (2) The erection of a Common External Tariff (CET) on all goods entering from third-party nations; (3) The institutionalization of the 'Four Freedoms'\u2014the free, unrestricted movement of goods, services, capital, and labor across national borders; and (4) The establishment of common policies, most notably the Common Agricultural Policy (CAP) in 1962, which guaranteed price supports and subsidies to French and European farmers, cementing French rural support for the integration project."),

    ("The Phenomenon of Neo-Functionalist 'Spillover' toward the European Union",
     "The unprecedented economic success of the Common Market between 1958 and 1970 verified political scientist Ernst B. Haas's theory of 'Neo-Functionalist Spillover.' Haas posited that economic integration is inherently dynamic: integration in one economic sector creates functional frictions and technical bottlenecks that compel governments to integrate adjacent sectors.\n\n"
     "The Common Market generated extraordinary economic growth: between 1958 and 1970, trade among the Six quadrupled, industrial production surged, and living standards reached unprecedented heights. This economic success produced multiple layers of spillover:\n\n"
     "1. Technical and Economic Spillover: The creation of a customs union made independent national currency fluctuations and exchange-rate devaluations deeply disruptive to cross-border trade and the Common Agricultural Policy. This functional tension forced member states into monetary coordination, leading to the European Monetary System (EMS) in 1979 and culminating in the single European currency (the Euro) under the 1992 Maastricht Treaty.\n\n"
     "2. Legal and Institutional Spillover: The European Court of Justice (ECJ), through landmark rulings like Van Gend en Loos (1963) and Costa v. ENEL (1964), established the revolutionary doctrines of 'Direct Effect' and the 'Supremacy of European Law' over conflicting national statutes, effectively constitutionalizing the Treaty of Rome.\n\n"
     "3. Political Spillover: As national economic policies became intertwined, national interest groups, trade unions, and political parties shifted their lobbying from national capitals to the European Commission in Brussels, directly propelling the expansion of the European Parliament into a directly elected legislative body (1979). Ultimately, the economic architecture of the Common Market served as the indispensable stepping stone that culminated in the 1986 Single European Act and the 1992 Maastricht Treaty, which officially transformed the European Community into the European Union (EU), expanding its remit into foreign policy, justice, and common citizenship."),

    ("The British Detachment and De Gaulle's Vetoes",
     "The transformative success of the Common Market was thrown into sharp relief by the historic strategic blunder of Great Britain. Convinced of its global imperial status, proud of its 'special relationship' with the United States, and suspicious of continental supranational bureaucracy, Britain refused to join the ECSC in 1951 or the EEC in 1957. In 1960, Britain attempted to counter the Six by creating the European Free Trade Association (EFTA)\u2014the 'Outer Seven'\u2014which established a loose free trade area without common tariffs or supranational institutions.\n\n"
     "Within three years, Britain's economic growth lagged far behind the booming Common Market. Recognizing that exclusion meant economic stagnation and diplomatic irrelevance, Conservative Prime Minister Harold Macmillan executed an agonizing policy reversal, formally applying for British membership in the EEC in August 1961. However, on January 14, 1963, French President Charles de Gaulle held a sensational press conference, issuing a categorical 'Non' and vetoing British entry. De Gaulle argued that Britain was insular, maritime, and fundamentally bound to the United States, viewing British entry as a Trojan Horse for American domination of Europe. De Gaulle vetoed Britain a second time in 1967, and Britain was admitted to the Common Market only in 1973 following De Gaulle's death."),

    ("Conclusion: The Triumphant Architecture of European Peace",
     "In conclusion, the post-war European integration movement was the most successful, sophisticated experiment in voluntary supranational statecraft in modern human history. It fundamentally solved the centuries-old 'German Question' not by dismembering or subjugating Germany, but by anchoring a democratic German nation within a web of shared European institutions and mutual economic dependence. The journey from the Schuman Declaration in 1950 and the ECSC in 1951 to the Treaty of Rome in 1957 demonstrated that the path to European unity lay not in the violent conquest of empires or the utopian dreams of immediate political federalism, but in the patient, technical construction of economic solidarity. By establishing the Common Market, the founding fathers of Europe built the durable institutional and legal foundations that inspired later generations to create the Single Market, the Eurozone, and the European Union, transforming Western Europe from the world's bloodiest ideological battlefield into a zone of permanent democratic peace and collective prosperity.")
]

quotes_p2_15 = [
    ("Robert Schuman (May 9, 1950)", "The Schuman Declaration", "'Europe will not be made all at once, or according to a single plan. It will be built through concrete achievements which first create a de facto solidarity... The pooling of coal and steel will make war between France and Germany not merely unthinkable, but materially impossible.'"),
    ("Ernst B. Haas", "The Uniting of Europe (1958)", "'Integration is a process whereby political actors in distinct national settings are persuaded to shift their loyalties, expectations, and political activities toward a new centre, whose institutions possess or demand jurisdiction over the pre-existing national states.'"),
    ("Alan S. Milward", "The European Rescue of the Nation-State (1992)", "'European integration was not the contradiction of the nation-state, but its salvation... By pooling sovereignty in specific economic areas, European governments restored the social legitimacy and domestic prosperity of national democratic states after the catastrophe of 1945.'"),
    ("Jean Monnet (1976)", "Memoirs", "'We are not forming coalitions of states; we are uniting people... Human nature does not change, but institutions can change human behaviour.'")
]

# Run the generation
if __name__ == "__main__":
    tasks = [
        (meta_p2_13, callout_p2_13, map_p2_13, timeline_p2_13, essay_p2_13, quotes_p2_13,
         os.path.join(paper2_dir, "Q13_Soviet_Domination_of_Eastern_Europe_and_Hungarian_Uprising_Master_Notes.docx"),
         "QUESTION 13: SOVIET HEGEMONY IN EASTERN EUROPE & THE HUNGARIAN UPRISING (1956)",
         "Q13_Soviet_Domination_of_Eastern_Europe_and_Hungarian_Uprising_Master_Notes"),
        (meta_p2_14, callout_p2_14, map_p2_14, timeline_p2_14, essay_p2_14, quotes_p2_14,
         os.path.join(paper2_dir, "Q14_United_Nations_Organization_vs_League_of_Nations_Master_Notes.docx"),
         "QUESTION 14: THE UNITED NATIONS (UNO) VS. THE LEAGUE OF NATIONS: STRUCTURE & PEACE",
         "Q14_United_Nations_Organization_vs_League_of_Nations_Master_Notes"),
        (meta_p2_15, callout_p2_15, map_p2_15, timeline_p2_15, essay_p2_15, quotes_p2_15,
         os.path.join(paper2_dir, "Q15_European_Integration_Schuman_Plan_and_Common_Market_Master_Notes.docx"),
         "QUESTION 15: THE ROAD TO EUROPEAN UNITY: SCHUMAN PLAN, ECSC & THE COMMON MARKET",
         "Q15_European_Integration_Schuman_Plan_and_Common_Market_Master_Notes")
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
