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
# QUESTION 1 (PAPER II): 2025 Q.1 (PART 1) / 2013 Q.1 (WWI AS TOTAL WAR)
# =========================================================================

meta_p2_1 = [
    ("EXAM QUESTIONS LINKED:", "2025 Q.1: 'Analyze the First World War as a total war, examining its impact on European societies, economies, and political systems, and evaluate the consequences of the Treaty of Versailles as a long-term cause of future conflict.'\n2013 Q.1: 'Evaluate the underlying and fundamental causes that led to the outbreak of the First World War.'"),
    ("SYLLABUS TOPIC:", "Paper II \u2014 Section 1: The First World War (1914\u20131918) \u2014 Total War Mobilization, Home Front, Economic Disruption & Imperial Collapse"),
    ("ALLOCATED MARKS / TIME:", "20 Marks | 35 Minutes | Standard Word Count: ~650\u2013700 words"),
    ("HISTORIOGRAPHICAL THESIS:", "The First World War constituted the inaugural 'total war' of human history, obliterating the traditional nineteenth-century demarcation between military combatants and civilian non-combatants through state-directed economic regimentation, universal conscription, and pervasive mass propaganda; the war fundamentally reconstructed European society by propelling female labor emancipation, demolishing four historic dynastic empires, and unleashing deep socioeconomic traumas that catalyzed modern totalitarianism.")
]

callout_p2_1 = [
    "\u2022 The Mechanics of 'Total War': Coined conceptually by General Erich Ludendorff; characterized by total mobilization of societal human and industrial resources, total war aims (unconditional capitulation rather than negotiated frontier shifts), and the total elimination of the distinction between battlefield and home front.",
    "\u2022 The Command Economy & War Socialism: Walter Rathenau's War Raw Materials Department (Kriegsrohstoffabteilung - KRA) in Germany; Lloyd George's Ministry of Munitions in Britain; comprehensive price controls, food rationing, mandatory labor direction, and nationalization of railways, coal mines, and shipping.",
    "\u2022 Societal Mutation & Gender Emancipation: Mobilization of over 1.5 million British women and millions across the continent into heavy munitions ('munitionettes'), transportation, and agriculture dismantled the Victorian doctrine of separate domestic spheres, directly compelling post-war female enfranchisement (Britain's 1918 Representation of the People Act, German Weimar Constitution of 1919).",
    "\u2022 Political Cataclysm: The industrial meat-grinder caused the unprecedented collapse of four grand imperial dynasties: the Romanovs in Russia (Bolshevik Revolution 1917), the Hohenzollerns in Germany (Weimar Republic 1918), the Habsburgs in Austria-Hungary (fracturing into successor states), and the Ottoman Empire (partition and Turkish Republic)."
]

map_p2_1 = """========================================================================================
MILITARY ATTRITION                HOME FRONT REGIMENTATION          POLITICAL & SOCIAL MUTATION
========================================================================================
[Trench Deadlock (Somme/Verdun)]--> [State Command Economies]     --> [Female Labor Mobilization]
[Industrial Munitions Demand]   --> ['War Socialism' & Rationing] --> [Representation Acts (Suffrage)]
[Naval Blockades & U-Boats]     --> [Universal Labor Conscription]--> [Psychic Trauma: 'Lost Generation']
[Mass Propaganda & Censorship]  --> [Suppression of Civil Rights] --> [Fall of 4 Dynastic Empires:
                                                                       Russia, Germany, Austria, Ottomans]
========================================================================================"""

timeline_p2_1 = [
    ("Aug 1914", "Outbreak of WWI", "Schlieffen Plan halted at the Battle of the Marne; 400-mile trench line established from the Swiss border to the North Sea."),
    ("May 1915", "Ministry of Munitions Established", "David Lloyd George reorganizes British industry to resolve the catastrophic 'shell shortage' crisis."),
    ("Feb\u2013Dec 1916", "Battles of Verdun & The Somme", "Apex of industrial war of attrition; combined casualties exceed 1.7 million soldiers without strategic breakthrough."),
    ("Dec 1916", "German Auxiliary Services Law", "Hindenburg-Ludendorff dictatorship imposes universal labor conscription for all males aged 17 to 60 (Hilfsdienstgesetz)."),
    ("March & Nov 1917", "Dual Russian Revolutions", "Tsar Nicholas II abdicates under bread riot pressure; Bolsheviks seize power in October, exiting war at Brest-Litovsk (March 1918)."),
    ("April 1917", "US Enters the War", "German unrestricted submarine warfare and the Zimmermann Telegram compel President Woodrow Wilson to declare war."),
    ("Jan 1918", "Wilson's Fourteen Points", "Articulates liberal war aims: national self-determination, open diplomacy, and a League of Nations."),
    ("March\u2013July 1918", "Ludendorff's Spring Offensive", "Germany's final desperate gamble fails; Allied Hundred Days counter-offensive breaks the Hindenburg Line."),
    ("Nov 9\u201311, 1918", "German Collapse & Armistice", "Kaiser Wilhelm II abdicates; German Republic proclaimed; Armistice signed at Compi\u00e8gne, concluding combat.")
]

essay_p2_1 = [
    ("Introduction: The Genesis of Industrialized Total War",
     "The outbreak of the First World War in August 1914 marked the violent termination of the long nineteenth-century peace and the apocalyptic inaugurator of modern 'Total War.' Prior to 1914, European military doctrine, shaped by the brief cabinet wars of Bismarck, anticipated a rapid, mobile conflict resolved within months by professional armies. Instead, the lethal defensive supremacy of rapid-firing artillery, barbed wire, and the machine gun crystallized into an impenetrable, four-hundred-mile trench stalemate on the Western Front. To sustain this mechanized war of attrition, European combatants were compelled to pioneer a fundamentally new mode of warfare. As conceptualized by General Erich Ludendorff and later historicized by Arthur Marwick and Eric Hobsbawm, World War I was 'total' because it demanded the complete subordination of the civilian home front to the military machine, dissolved the legal boundaries between soldiers and non-combatants, centralized the economy under state control, and unleashed structural shocks that obliterated Europe's historic sociopolitical order."),

    ("Economic Regimentation: The Birth of the Command State and 'War Socialism'",
     "The primary structural casualty of total war was classical laissez-faire capitalism. The colossal expenditure of munitions\u2014typified by the British Army firing more artillery shells in the first week of the Battle of the Somme (1916) than during the entire preceding two years\u2014demanded an unprecedented reorganization of national industries. In Germany, Jewish industrialist Walter Rathenau established the War Raw Materials Department (Kriegsrohstoffabteilung - KRA), which rationed vital raw materials, developed synthetic substitutes (Ersatz goods like synthetic nitrates for gunpowder), and orchestrated the national economy under what became celebrated as 'War Socialism' (Kriegssozialismus). Under the Hindenburg Program and the Auxiliary Services Law of 1916, Germany instituted universal labor conscription for all males aged seventeen to sixty.\n\n"
     "In Great Britain, the catastrophic 'shell scandal' of May 1915 led to the creation of the Ministry of Munitions under David Lloyd George. The state seized direct operational control of coal mines, merchant shipping, and railways, constructed state-owned munitions factories, fixed national wages, outlawed strikes, and implemented strict food rationing to survive Germany's unrestricted submarine warfare. Government expenditure surged exponentially, permanently transforming the state into an economic Leviathan."),

    ("Societal Mutation: Gender Emancipation and the 'Lost Generation'",
     "Total war exerted a profound, permanent transformation upon European society, most visibly in the radical restructuring of gender roles. With millions of young men dispatched to the trenches, combatant nations faced catastrophic domestic labor shortages. Over one and a half million British women, and comparable numbers in France and Germany, entered the industrial workforce as factory machinists ('munitionettes'), tram conductors, railway stokers, bank clerks, and agricultural laborers in the Women's Land Army. By demonstrating that women could perform heavy, skilled labor previously monopolized by men, total war pulverized the Victorian bourgeois ideology of 'separate spheres.' Although demobilization in 1918 saw many women pushed back into the domestic sphere, the political debt could not be denied: Britain enfranchised women over thirty via the Representation of the People Act (1918), and the Weimar Republic granted full universal female suffrage in 1919.\n\n"
     "Concurrently, the horrific casualty rates shattered the psychological foundation of European society. Over ten million soldiers perished, accompanied by twenty million wounded and permanently maimed. The annihilation of an entire generation of young men\u2014particularly junior officers drawn from intellectual and aristocratic elites\u2014birthed the tragic cultural concept of the 'Lost Generation,' leaving European culture permeated by disillusionment, shell shock (PTSD), and existential nihilism."),

    ("The Information War: Psychological Mobilization and the Eradication of Civil Liberties",
     "Total war required not merely the mobilization of bodies and factories, but the regimentation of the human mind. To sustain domestic morale against horrific casualty lists and starvation rations, governments established centralized propaganda apparatuses. In Britain, the Ministry of Information at Crewe House orchestrated sophisticated campaigns demonizing the German 'Hun' as an uncivilized barbarian, fabricating atrocity stories (such as the bayoneting of Belgian babies) to inflame popular hatred. In the United States, George Creel's Committee on Public Information mobilized 75,000 'Four-Minute Men' to whip up patriotic fervor.\n\n"
     "Simultaneously, traditional civil liberties were ruthlessly suspended. Britain enacted the sweeping Defense of the Realm Act (DORA) in August 1914, authorizing warrantless arrests, total press censorship, and the prohibition of pacifist publications. In France, Georges Clemenceau assumed near-dictatorial powers in 1917, imprisoning defeatists and crushing labor strikes. Total war permanently legitimized state surveillance, the criminalization of political dissent, and the psychological manipulation of mass electorates."),

    ("Political Cataclysm: The Collapse of Four Empires and the Rise of Radicalism",
     "The ultimate political legacy of total war was the catastrophic dissolution of the dynastic architecture of Central and Eastern Europe. Under the unbearable strains of total mobilization, food shortages, and military slaughter, the autocratic regimes of Europe cracked. The Romanov dynasty in Russia fell first in March 1917, collapsing into the radical Bolshevik October Revolution that birthed the world's first Marxist-Leninist state. In November 1918, military mutinies in Kiel sparked the German Revolution, forcing Kaiser Wilhelm II into exile and establishing the fragile Weimar Republic. The multi-ethnic Austro-Hungarian Empire completely dissolved as Czechs, Poles, South Slavs, and Hungarians proclaimed independent nation-states. Finally, the Ottoman Empire collapsed, terminating six centuries of dynastic caliphal rule.\n\n"
     "Total war thus dismantled the traditional conservative ruling elites that had governed Europe since the Middle Ages. In their vacuum emerged an unstable continent carved into fragile successor states, polarized between radical Bolshevism and emergent fascist movements led by brutalized front-line veterans (the Frontgeneration) who brought the ruthless methods of the trenches into post-war civilian politics."),

    ("Conclusion: The Crucible of Twentieth-Century Catastrophe",
     "In conclusion, the First World War was the decisive historical watershed that terminated nineteenth-century liberal optimism and inaugurated the violent mass politics of the twentieth century. By transforming the entire nation into an industrial weapons-manufacturing collective, World War I demonstrated that the modern state possessed the terrifying bureaucratic capacity to mobilize millions, regiment economies, and annihilate human life on an industrial scale. The economic exhaustion, social dislocation, and political fracturing generated between 1914 and 1918 ensured that the post-war peace would be brittle. Rather than resolving the antagonisms of European civilization, total war unleashed the ideological monsters of Communism and Fascism, ensuring that the 'war to end all wars' functioned merely as the opening prologue to a greater twentieth-century catastrophe.")
]

quotes_p2_1 = [
    ("Arthur Marwick", "War and Social Change in the 20th Century", "'Total war is an enormous social catalyst; by testing institutions to destruction and forcing governments to mobilize women and the working class, it produced irreversibly egalitarian social consequences.'"),
    ("Eric Hobsbawm", "The Age of Extremes: 1914-1991", "'The First World War inaugurated the Age of Catastrophe... It was the first total war in history, in which the state mobilized all human and material resources to wage a war of unlimited aims and industrial destruction.'"),
    ("Jay Winter", "Sites of Memory, Sites of Mourning", "'The Great War destroyed an entire generation and with it the Victorian faith in progress, replacing civilizational optimism with a pervasive European culture of bereavement and trauma.'"),
    ("Erich Ludendorff", "Der Totale Krieg (Total War)", "'Total war is not only a matter for the armed forces; it touches the life and soul of every single citizen... the Home Front and the Battle Front have become one.'")
]

# =========================================================================
# QUESTION 2 (PAPER II): 2022 Q.2 / 2018 Q.1 / 2016 Q.2 / 2013 Q.2 (VERSAILLES)
# =========================================================================

meta_p2_2 = [
    ("EXAM QUESTIONS LINKED:", "2022 Q.2: 'Critically examine the Treaty of Versailles (1919). Was it a dictated peace (Diktat) that made the rise of Hitler inevitable?'\n2018 Q.1: ''The peace settlement of 1919 was the monument of peacemakers' inaptitude and short sightedness.' Discuss.'\n2016 Q.2: 'Examine the provisions of the Treaty of Versailles with special reference to war guilt clause, disarmament, and territorial adjustments imposed on Germany.'\n2013 Q.2: 'Woodrow Wilson's Fourteen Points promised a new international order based on justice. How far were these points reflected in the Paris Peace Settlement of 1919?'\n2025 Q.1 (Part 2): 'Evaluate the consequences of the Treaty of Versailles as a long-term cause of future conflict.'"),
    ("SYLLABUS TOPIC:", "Paper II \u2014 Section 2: The Paris Peace Settlement (1919\u20131923) \u2014 Treaty of Versailles, War Guilt, Reparations & Historical Debate"),
    ("ALLOCATED MARKS / TIME:", "20 Marks | 35 Minutes | Standard Word Count: ~650\u2013700 words"),
    ("HISTORIOGRAPHICAL THESIS:", "The Treaty of Versailles was a fatally compromised hybrid: too punitive to reconcile a defeated Germany, yet too lenient and feebly enforced to permanently incapacitate her; while its status as an unnegotiated 'Diktat' and the moral humiliation of Article 231 (War Guilt) provided Adolf Hitler with indispensable rhetorical dynamite, the rise of Nazism was not historically inevitable from the treaty alone, but was catalyzed by the 1929 Great Depression, the failure of Allied enforcement, and the collapse of the Weimar Republic.")
]

callout_p2_2 = [
    "\u2022 The Clash of Allied Visions: Woodrow Wilson arrived with the idealistic Fourteen Points (open covenants, national self-determination, collective security, League of Nations); Georges Clemenceau ('The Tiger') sought to crush Germany demographically, territorially, and economically to guarantee French security; David Lloyd George pragmatically balanced British naval dominance and imperial expansion against the danger of leaving Germany ripe for Bolshevism.",
    "\u2022 Core Penalties of Versailles (June 28, 1919):",
    "  1. Territorial: Alsace-Lorraine returned to France; Eupen-Malmedy to Belgium; Northern Schleswig to Denmark; Posen and West Prussia ceded to Poland (the 'Polish Corridor' dividing East Prussia); Danzig made a Free City; Saarland under League control for 15 years; all overseas colonies confiscated as League Mandates. Total loss: 13% of territory, 12% of population.",
    "  2. Military: German army restricted to 100,000 long-term volunteers; General Staff abolished; conscription banned; tanks, heavy artillery, military aircraft, and submarines completely prohibited; Rhineland permanently demilitarized and occupied by Allied troops for 15 years.",
    "  3. Financial & Moral: Article 231 ('War Guilt Clause') forced Germany to accept sole moral responsibility for all Allied loss and damage; 1921 London Ultimatum fixed reparations at the astronomical sum of 132 billion gold marks (\u00a36.6 billion).",
    "\u2022 The Reality of the 'Diktat': German delegates were completely excluded from the conference negotiations; presented with an ultimatum to sign within 100 hours or face the resumption of Allied military invasion and the starvation naval blockade; signed under protest in the Hall of Mirrors at Versailles.",
    "\u2022 Did Versailles Make Hitler Inevitable? While Versailles gave birth to the virulent 'stab-in-the-back' myth (Dolchsto\u00dflegende) and universal German revanchism, Hitler remained an electoral fringe (2.6% of vote in 1928) during the Locarno era (1924\u20131929). Versailles created the combustible timber, but the Wall Street Crash of 1929 was the spark that ignited the Weimar Republic."
]

map_p2_2 = """========================================================================================
ALLIED IMPASSE (1919)             VERSAILLES PROVISIONS             HISTORICAL CONSEQUENCES
========================================================================================
[Wilson: 14 Points Idealism]   --> [Territory: Polish Corridor,  --> ['Diktat' Backlash & Outrage]
[Clemenceau: Cripple Germany]       Alsace-Lorraine, Colonies]   --> [Dolchsto\u00dflegende ('Stab-in-Back')]
[Lloyd George: Balance Power]  --> [Military: 100k Army, No Air] --> [1923 Hyperinflation & Ruhr Crisis]
                               --> [Article 231: War Guilt]      --> [Locarno Golden Years (1924-29)]
                               --> [Reparations: 132B Marks]     --> [1929 Crash: Hitler Weaponizes Treaty]
========================================================================================"""

timeline_p2_2 = [
    ("Jan 8, 1918", "Wilson's Fourteen Points", "Promulgates blueprint for democratic peace, self-determination, disarmament, and a League of Nations."),
    ("Jan 18, 1919", "Paris Peace Conference Opens", "Leaders of 32 nations convene; defeated Central Powers and revolutionary Bolshevik Russia are excluded."),
    ("May 7, 1919", "Draft Treaty Presented to Germany", "Count Brockdorff-Rantzau rejects moral guilt; massive public demonstrations against treaty terms across Germany."),
    ("June 28, 1919", "Signing of the Treaty of Versailles", "German delegates sign in the Hall of Mirrors (Versailles) on the fifth anniversary of the Sarajevo assassination."),
    ("Nov 1919 & Mar 1920", "US Senate Rejects Treaty", "Isolationist Republicans led by Henry Cabot Lodge refuse ratification; US rejects League of Nations membership."),
    ("May 1921", "London Reparations Schedule", "Allied Reparations Commission assesses German liability at 132 billion gold marks (\u00a36.6 billion)."),
    ("Jan 1923", "French Invasion of the Ruhr", "Poincar\u00e9 orders troops into the Ruhr to seize coal; German passive resistance triggers catastrophic hyperinflation."),
    ("1924\u20131925", "Dawes Plan & Locarno Treaties", "US loans stabilize German currency; Germany voluntarily recognizes Western borders; joins League of Nations (1926)."),
    ("Oct 1929", "Wall Street Crash & Great Depression", "US recalls loans; German unemployment surges to 6 million; Nazi Party vote surges from 2.6% to 37.3% (1932)."),
    ("Jan 1933", "Hitler Appointed Chancellor", "Vows to dismantle the 'Versailles Diktat', remilitarize Germany, and reclaim lost territories.")
]

essay_p2_2 = [
    ("Introduction: The Historiographical Controversy of 1919",
     "The Treaty of Versailles, signed on June 28, 1919, in the Hall of Mirrors where the German Empire had been proclaimed forty-eight years earlier, remains the most fiercely contested diplomatic settlement in modern history. Designed to end the catastrophic conflagration of World War I and construct a durable international order, the treaty has been famously condemned by British economist John Maynard Keynes as a 'Carthaginian peace' and caricatured by subsequent generations as the direct midwife of Adolf Hitler's Third Reich. Conversely, modern revisionist historians such as Sally Marks and Margaret MacMillan argue that the settlement was neither uniquely draconian nor inherently unworkable. In reality, Versailles suffered from a fatal systemic compromise: it was too harsh to be accepted by any German political faction, yet too lenient and geographically preserving to prevent Germany's eventual re-emergence as the dominant industrial titan of Europe. Consequently, while the perceived indignity of the 'Diktat' provided the psychological fuel for right-wing extremism, it was the catastrophic economic shock of the Great Depression and the Allies' complete failure of political will to enforce the treaty that made the Second World War inevitable."),

    ("The Clash of the Big Three: Wilsonian Idealism vs. Continental Realpolitik",
     "The fundamental defect of the Paris Peace Conference lay in the irreconcilable philosophical and strategic divergence among the victorious 'Big Three.' United States President Woodrow Wilson approached the conference as a moral prophet, seeking to remake the international order through his Fourteen Points. Wilson advocated for national self-determination, universal disarmament, open diplomacy, and above all, the establishment of a League of Nations to guarantee collective security. In stark contrast, French Premier Georges Clemenceau ('The Tiger') was driven by an existential terror of German demographic and industrial superiority. Having witnessed two German invasions within his lifetime, Clemenceau pursued a purely defensive, punitive peace aimed at permanently crippling Germany through the dismemberment of the Rhineland, crushing financial indemnities, and extensive territorial amputations.\n\n"
     "Caught between these ideological extremes was British Prime Minister David Lloyd George. While publicly constrained by a jingoistic electorate demanding to 'Hang the Kaiser' and 'squeeze Germany until the pips squeak,' Lloyd George privately recognized that an excessively weakened Germany would destroy continental trade and leave Central Europe vulnerable to the spread of Soviet Bolshevism. The resulting treaty was an unprincipled compromise: it preserved Wilson's League of Nations while violating his principle of self-determination, and it punished Germany sufficiently to provoke incandescent outrage without giving France the permanent security of a fractured or detached Rhineland."),

    ("The Anatomy of Punishment: Territorial, Military, and Financial Clauses",
     "The specific provisions of Versailles dismantled Germany's imperial status across three critical dimensions:\n\n"
     "1. Territorial Amputations: Germany was stripped of over thirteen percent of its European territory and twelve percent of its population. In the west, Alsace-Lorraine was returned to France, and Eupen-Malmedy was ceded to Belgium. In the north, Northern Schleswig voted to join Denmark. Most disastrously for interwar stability, in the east, rich agricultural and industrial lands in Posen and West Prussia were surrendered to the resurrected state of Poland to create the 'Polish Corridor,' granting Poland access to the Baltic Sea but severing East Prussia from the rest of the Fatherland. Danzig was severed as an autonomous 'Free City,' and the coal-rich Saar basin was placed under League of Nations administration for fifteen years with its mines ceded to France. Furthermore, all German overseas colonies in Africa and the Pacific were confiscated as League of Nations mandates.\n\n"
     "2. Military Disarmament: To incapacitate Germany's offensive capability, the German army (Reichswehr) was capped at 100,000 long-term volunteers. Universal conscription was prohibited, the General Staff was abolished, and the manufacture of tanks, heavy artillery, military aircraft, and submarines was outlawed. The German navy was reduced to a token force of six obsolete battleships, and the Rhineland was placed under permanent demilitarization, with Allied forces occupying bridgeheads for fifteen years.\n\n"
     "3. Financial Indemnities and Article 231: To provide a legal justification for reparations, the Allies inserted the infamous Article 231\u2014the 'War Guilt Clause'\u2014which forced Germany to accept complete moral responsibility for causing all loss and damage to the Allied nations. In May 1921, the Allied Reparations Commission assessed Germany's total liability at 132 billion gold marks (\u00a36.6 billion), an astronomical figure that incensed German public opinion and destabilized the post-war European financial system."),

    ("The 'Diktat' and the Myth of the 'Stab in the Back'",
     "Beyond its material clauses, the procedural methodology of Versailles inflicted deep psychological trauma upon the German national consciousness. Germany was completely excluded from the conference deliberations. When the German delegation led by Count Ulrich von Brockdorff-Rantzau was finally summoned to Versailles in May 1919, they were presented with a completed draft and given an ultimatum: sign within days or face the resumption of Allied military invasion and the intensification of the naval blockade, which was already starving German civilians. To the German people, who believed they had concluded an honorable armistice based on Wilson's Fourteen Points, this was not a negotiated peace, but a shameful, coerced 'Diktat' (Dictated Peace).\n\n"
     "Crucially, because the German military had not suffered a catastrophic collapse on home territory in November 1918, Field Marshals Hindenburg and Ludendorff successfully cultivated the toxic 'stab-in-the-back' myth (Dolchsto\u00dflegende). The German public was propagandized into believing that their undefeated army in the field had been betrayed from within by socialists, liberals, and Jews who proclaimed the republic and signed the armistice. Consequently, the democratic Weimar Republic was indelibly stained from its birth with the humiliation of the 'Versailles Diktat' and the betrayal of the 'November Criminals.'"),

    ("Did Versailles Make the Rise of Hitler Inevitable?",
     "The assertion that Versailles made the rise of Adolf Hitler historically inevitable is an oversimplification that ignores the dynamics of the interwar decade. Between 1924 and 1929\u2014the era of Gustav Stresemann, the Dawes Plan, and the Locarno Treaties\u2014the Weimar Republic achieved remarkable economic stabilization and diplomatic normalization. Germany was admitted to the League of Nations with a permanent Council seat in 1926, and reparations were drastically scaled down via the Young Plan. In the May 1928 German federal elections, Hitler's National Socialist (Nazi) Party received a pitiful 2.6 percent of the vote, holding merely twelve seats in the Reichstag. If Versailles had made Hitler inevitable, the Nazi Party would have triumphed in the 1920s.\n\n"
     "Hitler's meteoric ascent was catalyzed not by the text of Versailles in 1919, but by the catastrophic external shock of the Wall Street Crash of October 1929. The Great Depression wiped out American capital loans, shuttered German banks, and generated six million unemployed citizens. It was in this environment of acute social despair that Hitler brilliantly weaponized Versailles. He merged economic agony with the dormant grievances of Article 231, the Polish Corridor, and the 'Diktat,' presenting his racist, militaristic program as the sole salvation to shatter the 'chains of Versailles.' Versailles was the indispensable gunpowder, but the Great Depression was the detonator."),

    ("Conclusion: The Anatomy of a Doomed Settlement",
     "In conclusion, the Treaty of Versailles was a tragic monument to the perils of diplomatic compromise without enforcement. In the prophetic words of French Marshal Ferdinand Foch upon reading the treaty in 1919: 'This is not peace. It is an armistice for twenty years.' The fundamental failure of Versailles was that it left Germany physically intact, demographically vibrant, and industrially potent, while surrounding her with small, militarily fragile successor states in Central Europe. When the United States retreated into isolationism and Great Britain adopted a guilt-ridden policy of appeasement, France was left alone, paralyzed by domestic instability and unable to enforce the treaty terms. By inflicting maximal moral insult while leaving the physical capacity for revenge intact, the peacemakers of 1919 created a diplomatic vacuum that totalitarian revanchism inevitably rushed to fill.")
]

quotes_p2_2 = [
    ("John Maynard Keynes", "The Economic Consequences of the Peace (1919)", "'The Treaty includes no provisions for the economic rehabilitation of Europe... It is a Carthaginian Peace that reduces an industrial nation to servitude and will sow the seeds of future ruin for the entire continent.'"),
    ("Sally Marks", "The Illusion of Peace (1976)", "'Versailles was not excessively harsh on Germany; its fatal flaw was that the victorious powers lacked the unity and the political will to enforce its terms once the immediate crisis of 1919 had passed.'"),
    ("Margaret MacMillan", "Peacemakers (2001)", "'It is too easy to blame Versailles for the horrors of the 1930s. The decisions that led to Hitler and World War II were made by politicians in the 1930s, not by the statesmen of 1919.'"),
    ("Marshal Ferdinand Foch (1919)", "Prophetic Verdict on Versailles", "'This is not peace. It is an armistice for twenty years.'")
]

# =========================================================================
# QUESTION 3 (PAPER II): 2022 Q.1 (OTTOMAN TURKEY, SÈVRES & LAUSANNE)
# =========================================================================

meta_p2_3 = [
    ("EXAM QUESTIONS LINKED:", "2022 Q.1: 'Highlighting the role of Ottoman Turkey in the First World War, critically analyze the treatment of Turkey by the Allied Powers in the post-war treaties (Treaty of S\u00e8vres and Treaty of Lausanne).'"),
    ("SYLLABUS TOPIC:", "Paper II \u2014 Section 2: Subordinate Treaties & Near Eastern Settlement \u2014 Ottoman Collapse, S\u00e8vres (1920) vs. Lausanne (1923)"),
    ("ALLOCATED MARKS / TIME:", "20 Marks | 35 Minutes | Standard Word Count: ~650\u2013700 words"),
    ("HISTORIOGRAPHICAL THESIS:", "Ottoman Turkey's entry into World War I fundamentally altered the geopolitical scope of the war by severing Imperial Russia from Western maritime supply and tying down massive Allied armies across the Middle East; while the victorious Allies sought at the Treaty of S\u00e8vres (1920) to partition the Anatolian heartland and reduce Turkey to an impotent protectorate, Mustafa Kemal Atat\u00fcrk's Turkish War of Independence destroyed the imperialist settlement, compelling the Allies at the Treaty of Lausanne (1923) to recognize full Turkish national sovereignty, making Turkey the only defeated Central Power to successfully overturn its post-war peace treaty.")
]

callout_p2_3 = [
    "\u2022 Strategic Impact of Ottoman Turkey in WWI: Entry in October 1914 closed the Turkish Straits (Bosphorus and Dardanelles), severing 90% of Russian maritime trade and directly accelerating Tsarist collapse; heroically repelled the Anglo-French Gallipoli Campaign (1915), forging Mustafa Kemal's legend; captured an entire British army at Kut-al-Amara in Mesopotamia (1916); diverted over one million British and imperial troops to the Middle Eastern theater.",
    "\u2022 The Allied Secret Treaties: Constantinople Agreement (1915 promising Straits to Russia), Sykes-Picot Agreement (1916 partitioning Arab provinces between Britain and France), Balfour Declaration (1917 pledging Jewish national home in Palestine), and Treaty of Saint-Jean-de-Maurienne (1917 promising southwestern Anatolia to Italy).",
    "\u2022 The Carthaginian Treaty of S\u00e8vres (August 10, 1920): Imposed on the collaborator Sultan Mehmed VI; partitioned Anatolia: Eastern Thrace and Smyrna/Izmir to Greece; independent Armenia; autonomous Kurdistan; Italian and French economic spheres in Adana/Antalya; Turkish Straits demilitarized under international commission; Capitulations restored; armed forces limited to 50,000 men. Reduced Turkey to a defenseless rump state in central Anatolia.",
    "\u2022 The Kemalist Resurgence & War of Independence: Mustafa Kemal landed at Samsun (May 19, 1919); convened the Grand National Assembly in Ankara; promulgated the National Pact (Misak-\u0131 Mill\u00ee); crushed the Armenian Republic in the east; defeated the Greek invasion armies at Sakarya (1921) and Dumlup\u0131nar (1922); liberated Izmir; precipitated the Chanak Crisis that broke British resolve and toppled Lloyd George's government.",
    "\u2022 The Triumph of the Treaty of Lausanne (July 24, 1923): Complete revision of S\u00e8vres; recognized complete Turkish sovereignty over Anatolia and Eastern Thrace; abolished all Capitulations, foreign financial control, and reparations; demilitarized the Straits under Turkish administration; implemented compulsory Greco-Turkish population exchange (1.2M Greeks, 400k Turks); laid the international foundation for the modern, secular Republic of Turkey."
]

map_p2_3 = """========================================================================================
WWI ROLE & DISMEMBERMENT             KEMALIST RESISTANCE               SOVEREIGN TRIUMPH (1923)
========================================================================================
[Straits Closed & Gallipoli 1915]--> [May 1919: Samsun Landing]    --> [Sakarya & Dumlup\u0131nar Victories]
[Arab Revolt & Mudros Armistice] --> [National Pact: Misak-\u0131 Mill\u00ee]--> [Chanak Crisis: UK Backs Down]
[Allied Occupation of Istanbul]  --> [Ankara National Assembly]    --> [Nov 1922: Sultanate Abolished]
[TREATY OF S\u00c8VRES (1920):       --> [War of Independence Against  --> [TREATY OF LAUSANNE (1923):
 Partitioned Anatolia, Greece in       Greeks, Armenians, French]       Complete Sovereignty, No Reparations,
 Smyrna, Capitulations Restored]                                        Capitulations Eradicated]
========================================================================================"""

timeline_p2_3 = [
    ("Oct 29, 1914", "Ottoman Entry into WWI", "German-Ottoman warships Goeben and Breslau bombard Russian Black Sea ports, drawing Empire into war."),
    ("Feb 1915\u2013Jan 1916", "Gallipoli Campaign", "Mustafa Kemal's brilliant defense of Gallipoli repels Allied invasion, costing Allies 250,000 casualties."),
    ("May 1916", "Sykes-Picot Agreement", "Britain and France secretly partition Ottoman Arab territories (Syria, Lebanon, Iraq, Transjordan)."),
    ("Oct 30, 1918", "Armistice of Mudros", "Ottoman Empire capitulates; Allied warships anchor in Istanbul; military forces demobilized."),
    ("May 19, 1919", "Kemal Lands at Samsun", "Mustafa Kemal ignites the Turkish National Movement; rejects Istanbul government's subservience to Allies."),
    ("Aug 10, 1920", "Treaty of S\u00e8vres Signed", "Sultan's representatives sign Carthaginian partition treaty; Grand National Assembly in Ankara rejects it as treason."),
    ("Aug\u2013Sept 1921", "Battle of Sakarya", "Kemalist armies halt the advancing Greek army 50 miles from Ankara; 22-day battle of national survival."),
    ("Aug\u2013Sept 1922", "Great Offensive & Liberation of Izmir", "Turkish forces shatter Greek lines at Dumlup\u0131nar; reclaim Smyrna; Chanak Crisis topples British PM Lloyd George."),
    ("Nov 1, 1922", "Abolition of the Sultanate", "Grand National Assembly abolishes the Ottoman monarchy, ending 623 years of dynastic rule."),
    ("July 24, 1923", "Signing of the Treaty of Lausanne", "Allied Powers recognize sovereign borders of modern Turkey; abolishes Capitulations; proclaimed Republic (Oct 1923).")
]

essay_p2_3 = [
    ("Introduction: From Imperial Collapse to National Rebirth",
     "The fate of Ottoman Turkey across the First World War and its aftermath represents one of the most remarkable geopolitical reversals in modern diplomatic history. Entering the conflict in October 1914 on the side of the Central Powers, the Ottoman Empire\u2014long disparaged by European statesmen as the 'Sick Man of Europe'\u2014exercised a profound, transformative impact on the course of global hostilities. However, upon its military defeat in 1918, the Allied Powers sought to execute a centuries-old imperialist ambition: the complete dismemberment of the Ottoman state through the predatory Treaty of S\u00e8vres (1920). While Germany, Austria, Hungary, and Bulgaria succumbed to the dictated peace of Paris, Turkey alone mounted a successful armed revolution. Under the military and political genius of Mustafa Kemal Atat\u00fcrk, the Turkish national resistance demolished the S\u00e8vres partition on the battlefield, compelling the victorious Great Powers to assemble at Lausanne in 1923 and negotiate a radically revised treaty that recognized full Turkish sovereignty, establishing modern secular Turkey upon the ruins of dynastic theocracy."),

    ("Ottoman Turkey's Strategic Role in the First World War",
     "The intervention of the Ottoman Empire in October 1914 transformed a localized European war into a global conflagration. By closing the Turkish Straits (the Dardanelles and Bosphorus), the Ottomans severed Imperial Russia from its British and French allies, cutting off over ninety percent of Russia's grain export revenue and choking off Western ammunition shipments. This maritime isolation critically exacerbated Russia's industrial collapse, accelerating the domestic breakdown that precipitated the 1917 Bolshevik Revolution.\n\n"
     "Furthermore, the Ottoman army mounted an astonishing military resistance that confounded Allied expectations of an easy collapse. In 1915, during the epic Gallipoli Campaign, Ottoman forces commanded on the ground by Lieutenant-Colonel Mustafa Kemal repelled nine months of relentless British, French, and ANZAC amphibious assaults, inflicting over 250,000 casualties and forcing an ignominious Allied evacuation. In Mesopotamia, Ottoman forces encircled and captured an entire British army under General Townshend at Kut-al-Amara (April 1916). Across Palestine, Syria, and the Caucasus, the Ottoman Empire forced the British Empire to commit over one million troops, diverting vital manpower and material resources away from the Western Front."),

    ("Allied Carve-Up: Secret Pacts and the Carthaginian Treaty of S\u00e8vres (1920)",
     "Following the Armistice of Mudros (October 1918), Allied warships anchored in the Golden Horn, occupying Constantinople. Guided by a web of secret wartime accords\u2014the Sykes-Picot Agreement (1916) dividing Arab provinces into British and French spheres, and the Balfour Declaration (1917) promising a Jewish national home\u2014the Allied Powers imposed the Treaty of S\u00e8vres on August 10, 1920. Signed by the craven representatives of Sultan Mehmed VI, S\u00e8vres was the most draconian of all post-war treaties.\n\n"
     "S\u00e8vres did not merely detach non-Turkish Arab provinces; it partitioned the Anatolian heartland itself. Greece was awarded Eastern Thrace, Gallipoli, and the Aegean province of Smyrna (Izmir); an independent Armenian state was established in Eastern Anatolia; an autonomous Kurdistan was designated in the southeast; and extensive Italian and French 'spheres of influence' were carved out in Antalya and Cilicia. Turkey's capital, Constantinople, and the Straits were demilitarized and placed under an international Allied Commission. Economically, the humiliating Capitulations were restored, placing the Ottoman budget under foreign financial supervision, while the Turkish military was capped at a token force of 50,000 men. S\u00e8vres effectively extinguished Turkish independence, reducing the nation to an impoverished, landlocked client state."),

    ("Mustafa Kemal and the War of National Liberation (1919\u20131922)",
     "The implementation of S\u00e8vres was shattered by the spontaneous emergence of the Turkish National Movement. Landing at Samsun on May 19, 1919, Mustafa Kemal galvanized national resistance across the Anatolian plateau. In 1920, the Grand National Assembly convened in Ankara, issuing the 'National Pact' (Misak-\u0131 Mill\u00ee), which renounced Arab territories but demanded absolute sovereignty over all lands inhabited by an Ottoman Muslim majority within the armistice lines.\n\n"
     "When the Greek army, backed by British Prime Minister David Lloyd George, launched an aggressive invasion into inner Anatolia to enforce S\u00e8vres, Kemal organized a disciplined national army. In a series of brilliant military campaigns, the Kemalists secured their eastern flank by defeating Armenia (Treaty of G\u00fcmr\u00fc), negotiated the withdrawal of French troops from Cilicia (Treaty of Ankara 1921), and obtained vital Soviet financial and arms aid (Treaty of Moscow 1921). In August\u2013September 1921, Kemal halted the Greek advance at the epic twenty-two-day Battle of Sakarya. In August 1922, Kemal launched the Great Offensive, shattering Greek defenses at Dumlup\u0131nar and liberating Izmir on September 9, 1922. When Turkish forces marched on the British-garrisoned neutral zone at the Dardanelles (the Chanak Crisis), Britain's Dominions refused to fight, British public opinion rebelled, and Lloyd George's coalition government collapsed, forcing the Allies to sue for a negotiated peace."),

    ("The Diplomatic Masterstroke: The Treaty of Lausanne (1923)",
     "Signed on July 24, 1923, following eight months of acrimonious negotiations led by Ismet In\u00f6n\u00fc, the Treaty of Lausanne completely superseded and obliterated the Treaty of S\u00e8vres. Unlike the dictated treaties of Versailles, Saint-Germain, and Trianon, Lausanne was an authentically negotiated peace between equals. The treaty recognized the sovereign boundaries of the modern Turkish state, restoring Eastern Thrace, Izmir, and the entire Anatolian mainland to Turkish rule. All foreign military occupation zones, independent Armenian claims, and Kurdish autonomy provisions were liquidated.\n\n"
     "Crucially, Lausanne abolished the century-old Capitulations, freeing Turkey from foreign judicial and financial tutelage, and waived all Allied claims for war reparations. To permanently resolve ethnic conflict, Lausanne mandated a compulsory population exchange, transferring 1.2 million Orthodox Christians to Greece and 400,000 Muslims to Turkey. Although the Turkish Straits were temporarily demilitarized under a Turkish-chaired commission (fully remilitarized later under the 1936 Montreux Convention), Lausanne was an unmitigated triumph. On October 29, 1923, Mustafa Kemal proclaimed the Republic of Turkey, serving as its first President and launching sweeping secular modernization reforms."),

    ("Conclusion: The Unique Legacy of the Turkish Settlement",
     "In conclusion, the post-war treatment of Ottoman Turkey offers a profound historical study in the limits of imperialist coercion when confronted by organized national resistance. The Treaty of S\u00e8vres failed because it attempted to treat a proud, militarized nation as a passive colonial conquest. By contrast, the Treaty of Lausanne succeeded because it reflected the military and political realities established on the ground by the Turkish people. While Versailles left Germany humiliated, resentful, and primed for revanchism, Lausanne provided Turkey with authentic, recognized borders and total legal sovereignty. Consequently, modern Turkey evolved into a stable, status-quo power that maintained peace and neutrality during the interwar years, proving that lasting peace is forged not through punitive imperial diktats, but through the mutual recognition of national sovereignty.")
]

quotes_p2_3 = [
    ("David Fromkin", "A Peace to End All Peace (1989)", "'The Treaty of S\u00e8vres was the high-water mark of European imperialist arrogance... In attempting to carve up the Turkish homeland, the Allies unwittingly provoked the birth of modern Turkish nationalism.'"),
    ("Bernard Lewis", "The Emergence of Modern Turkey", "'The Treaty of Lausanne was a triumph for Turkish diplomacy... Turkey was the only one of the defeated Central Powers of 1918 that was able to negotiate with the Allies as an equal and tear up the dictated peace imposed upon it.'"),
    ("Margaret MacMillan", "Peacemakers (2001)", "'Of all the treaties signed at the Paris Peace Conference, S\u00e8vres was the most completely dead letter... Atat\u00fcrk demonstrated that military force combined with diplomatic skill could reverse the verdict of 1918.'"),
    ("Mustafa Kemal Atat\u00fcrk (1923)", "Speech on the Treaty of Lausanne", "'This treaty is a monument to the defeat of a colossal conspiracy against the Turkish nation, conceived through the centuries and carried out through the Treaty of S\u00e8vres... It is an unprecedented political victory for the Turkish nation.'")
]

# Run the generation
if __name__ == "__main__":
    tasks = [
        (meta_p2_1, callout_p2_1, map_p2_1, timeline_p2_1, essay_p2_1, quotes_p2_1,
         os.path.join(paper2_dir, "Q1_WWI_Total_War_and_Sociopolitical_Impact_Master_Notes.docx"),
         "QUESTION 1: WORLD WAR I AS TOTAL WAR & SOCIETAL TRANSFORMATION (1914\u20131918)",
         "Q1_WWI_Total_War_and_Sociopolitical_Impact_Master_Notes"),
        (meta_p2_2, callout_p2_2, map_p2_2, timeline_p2_2, essay_p2_2, quotes_p2_2,
         os.path.join(paper2_dir, "Q2_Treaty_of_Versailles_1919_Diktat_and_Hitler_Rise_Master_Notes.docx"),
         "QUESTION 2: THE TREATY OF VERSAILLES (1919) \u2014 DIKTAT, WAR GUILT & THE RISE OF HITLER",
         "Q2_Treaty_of_Versailles_1919_Diktat_and_Hitler_Rise_Master_Notes"),
        (meta_p2_3, callout_p2_3, map_p2_3, timeline_p2_3, essay_p2_3, quotes_p2_3,
         os.path.join(paper2_dir, "Q3_Ottoman_Turkey_WWI_Sevres_and_Lausanne_Master_Notes.docx"),
         "QUESTION 3: OTTOMAN TURKEY IN WWI & POST-WAR TREATIES (S\u00c8VRES TO LAUSANNE)",
         "Q3_Ottoman_Turkey_WWI_Sevres_and_Lausanne_Master_Notes")
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
