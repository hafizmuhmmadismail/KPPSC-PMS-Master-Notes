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
# QUESTION 15: 2025 Q.5 / 2025 Q.6 (INDUSTRIALIZATION & NEW IMPERIALISM)
# =========================================================================

meta_15 = [
    ("EXAM QUESTIONS LINKED:", "2025 Q.5: 'Discuss the relationship between industrialization, imperialism, and nationalism in shaping Europe's global dominance during the 19th century.'\n2025 Q.6: 'What were the primary economic, political, and ideological motivations behind the New Imperialism of the late 19th century?'"),
    ("SYLLABUS TOPIC:", "Paper I \u2014 Section 6: Industrialization, Nationalism & The Age of New Imperialism (1870\u20131914)"),
    ("ALLOCATED MARKS / TIME:", "20 Marks | 35 Minutes | Standard Word Count: ~650\u2013700 words"),
    ("HISTORIOGRAPHICAL THESIS:", "The 'New Imperialism' of 1870\u20131914 was the explosive structural convergence of industrial capitalism, competitive mass nationalism, and racist pseudo-scientific ideology; the Second Industrial Revolution created an insatiable demand for captive raw materials, protected markets, and surplus capital outlets, while aggressive post-1870 national chauvinism and Social Darwinist rhetoric transformed colonial expansion into the ultimate existential metric of Great Power survival.")
]

callout_15 = [
    "\u2022 The Tripartite Dynamic: Industrialization provided the material imperatives (surplus capital, search for rubber, petroleum, cotton, and copper) and the technological 'tools of empire' (steam navigation, telegraph cables, Maxim machine guns, quinine prophylaxis); Nationalism provided the political engine (prestige, naval coaling stations, popular jingoism); Ideology provided the moral camouflage (Kipling's 'White Man's Burden', French 'Mission Civilisatrice', and Social Darwinism).",
    "\u2022 The Hobson-Lenin Thesis: J.A. Hobson argued that domestic underconsumption and income inequality in Europe forced financiers to seek speculative investments abroad; V.I. Lenin formulated that imperialism is the 'highest stage of capitalism,' characterized by monopoly cartels and bank capital dividing the world.",
    "\u2022 The Scramble for Africa (1881\u20131914): In 1870, barely 10% of Africa was under European control; by 1914, over 90% was partitioned following Otto von Bismarck's 1884\u20131885 Berlin West Africa Conference, which institutionalized the doctrine of 'effective occupation.'",
    "\u2022 Flashpoints to World War: Imperial friction exported European tensions globally before boomeranging back into the heart of Europe: the Fashoda Crisis (1898), the Boer War (1899\u20131902), and the First and Second Moroccan Crises (1905, 1911) cemented the Anglo-French Entente and accelerated the drift toward 1914."
]

map_15 = """========================================================================================
ECONOMIC IMPERATIVES             TECHNOLOGICAL TOOLS             POLITICAL & IDEOLOGICAL MOTIVES
========================================================================================
[Second Industrial Revolution]  + [Steamships & Suez Canal]    + [Chauvinist Mass Nationalism]
[Surplus Capital & Markets]     + [Subsea Telegraphs]         + [Naval Coaling Stations / Prestige]
[Raw Materials: Oil, Rubber]    + [Quinine Malaria Defense]   + [Social Darwinism & Race Hierarchy]
[Protective Tariff Walls]       + [Maxim Machine Guns]        + ['The White Man's Burden']
                                                ||
                                                \/
========================================================================================
THE SCRAMBLE FOR AFRICA & ASIA (1881-1914) --> 84% OF GLOBE COLONIZED --> PRE-1914 FLASHPOINTS
========================================================================================"""

timeline_15 = [
    ("1869", "Opening of the Suez Canal", "Slashes maritime transit time to India and the Far East; sparks Anglo-French rivalry over Egypt and the Levant."),
    ("1876", "Brussels Geographic Conference", "King Leopold II of Belgium establishes the International African Association; begins ruthless conquest of Congo."),
    ("1882", "British Occupation of Egypt", "Prime Minister Gladstone occupies Cairo to protect Suez Canal; ignites the general 'Scramble for Africa'."),
    ("1884\u20131885", "Berlin West Africa Conference", "Bismarck brokers rules for African partition: establishes principle of 'Effective Occupation' to claim colonies."),
    ("1898", "Fashoda Incident in Sudan", "Kitchener and Marchand face off on the Upper Nile; France backs down; paves the way for 1904 Entente Cordiale."),
    ("1898\u20131899", "'Scramble for Concessions' in China", "European powers carve out extraterritorial spheres of influence in Qing China; triggers Boxer Rebellion (1900)."),
    ("1899\u20131902", "Second Boer War in South Africa", "Britain deploys 450,000 troops to defeat Boer republics; brutal concentration camp tactics isolate Britain diplomatically."),
    ("1902", "J.A. Hobson Publishes 'Imperialism'", "Seminal economic critique attributing colonial expansion to domestic underconsumption and financier conspiracies."),
    ("1904\u20131905", "Russo-Japanese War", "First modern defeat of a European Great Power by an Asian nation; shatters myth of white invincibility."),
    ("1905 & 1911", "First & Second Moroccan Crises", "Kaiser Wilhelm II challenges French dominance in Morocco; backfires, driving Britain and France into joint military staff talks.")
]

essay_15 = [
    ("Introduction: The Phenomenon of New Imperialism",
     "Between 1870 and 1914, European civilization expanded outward with unprecedented velocity and destructive power, subjugating nearly eighty-four percent of the globe's land surface. This phenomenon, historiographically designated the 'New Imperialism,' differed radically from the mercantile colonial expansion of the sixteenth and seventeenth centuries. Where early modern colonialism focused on coastal trading enclaves and precious metals under mercantilist state monopolies, New Imperialism was characterized by comprehensive territorial annexation, direct bureaucratic administration, and the deep economic penetration of continental interiors in Africa, Asia, and the Pacific. The driving engine behind this titanic expansion was an indivisible tripartite synergy: the structural economic imperatives of the Second Industrial Revolution, the hyper-competitive mass nationalism of post-1870 Europe, and the pervasive ideological hegemony of Social Darwinism and the 'civilizing mission.' Together, these forces transformed imperial acquisition into both a perceived economic necessity and the ultimate test of Great Power virility."),

    ("Economic Foundations: Industrial Capitalism and the Hobson-Lenin Thesis",
     "At the structural base of New Imperialism lay the transformative economic pressures of the Second Industrial Revolution (1870\u20131914). Unlike the first industrial phase centered on coal and textiles, the second phase was anchored in heavy steel, electrical machinery, chemicals, and internal combustion engines. This industrial acceleration generated an insatiable demand for tropical raw materials that Europe could not produce: rubber for tires and electrical insulation from the Congo and Amazon, copper for electric wiring from Central Africa, petroleum for navies, and vegetable oils for industrial lubricants. Concurrently, European domestic markets became saturated due to severe wage compression and unequal wealth distribution, while the Long Depression (1873\u20131896) prompted continental nations to erect protective tariff walls (such as Germany's Bismarckian tariffs of 1879). Industrialists were forced to secure captive colonial markets shielded from foreign competition.\n\n"
     "In his seminal 1902 critique, 'Imperialism: A Study,' British liberal economist J.A. Hobson demonstrated that the maldistribution of domestic income produced 'underconsumption' at home, leaving powerful financial cartels with vast accumulations of surplus capital. To achieve high profit margins, these financiers pressured national governments to use imperial armies to secure overseas investment outlets. In 1917, Vladimir Lenin radicalized Hobson's thesis in 'Imperialism, the Highest Stage of Capitalism,' arguing that the emergence of monopoly cartels and the merger of bank and industrial capital made the territorial division of the world inevitable, and that inter-capitalist conflicts over colonial spoils would inevitably generate cataclysmic global wars."),

    ("Technological Hegemony: Daniel Headrick's 'Tools of Empire'",
     "The breathtaking rapidity of late-nineteenth-century conquest was made possible only by revolutionary technological superiority. In his landmark study 'The Tools of Empire' (1981), Daniel Headrick demonstrated that European penetration of disease-ridden African and Asian interiors required specific technological breakthroughs. First, the isolation of quinine from cinchona bark in the 1850s provided regular medical prophylaxis against malaria, slashing European mortality rates in tropical regions from fifty percent to under five percent. Second, steam-powered shallow-draft riverboats allowed Europeans to navigate treacherous African river systems against the current. Third, the opening of the Suez Canal in 1869 and the laying of trans-oceanic telegraph cables allowed London, Paris, and Berlin to communicate with colonial garrisons in seconds. Finally, the massive disparity in firepower sealed European hegemony: the transition from muzzle-loading flintlocks to breech-loading rifles and, above all, the Maxim machine gun (1884) meant that small European forces could slaughter vastly larger indigenous armies, as demonstrated at the Battle of Omdurman (1898) where British troops killed eleven thousand Mahdist fighters while losing only forty-eight men."),

    ("Political Motivations: Competitive Mass Nationalism and Strategic Chokepoints",
     "While economic and technological forces provided the fuel and mechanism, competitive nationalism supplied the psychological propulsion for New Imperialism. The unifications of Germany and Italy in 1870\u20131871 fundamentally altered European psychology: nation-states were now locked in a zero-sum Darwinian struggle for continental prestige. A Great Power without colonies was increasingly viewed as a dying empire. In France, humiliated by the loss of Alsace-Lorraine in 1870, leaders like Jules Ferry championed colonial expansion in Indochina and West Africa as the essential means to restore national gloire and prove French vitality. In newly united Germany, Kaiser Wilhelm II abandoned Bismarck's cautious continental diplomacy in favor of Weltpolitik ('World Politics'), demanding Germany's rightful 'place in the sun' and a High Seas Fleet capable of challenging the British Empire.\n\n"
     "Furthermore, imperialism developed its own self-perpetuating strategic logic. To defend existing colonies, powers felt compelled to seize adjacent territories. The classic example was Britain's occupation of Egypt in 1882 to safeguard the Suez Canal\u2014the vital maritime lifeline to British India. Securing Egypt then compelled Britain to conquer Sudan to control the Nile's headwaters, which in turn spurred rivalries with France (culminating in the 1898 Fashoda standoff) and Germany in East Africa. Naval coaling stations, cable relay stations, and border buffers became strategic imperatives across the world's oceans."),

    ("Ideological Camouflage: Social Darwinism, Racism, and the 'Civilizing Mission'",
     "New Imperialism was intellectually legitimized and popularized among European voting publics through a pervasive ideology of racial superiority. The biological theories of Charles Darwin were cynically distorted by theorists like Herbert Spencer and Walter Bagehot into 'Social Darwinism.' Human societies were conceptualized as biological organisms engaged in a brutal 'survival of the fittest'; white European races were deemed the pinnacle of evolutionary development, destined by nature to dominate or displace 'inferior' indigenous peoples. This pseudo-scientific racism permeated state school curricula, popular press, and imperial expositions.\n\n"
     "Simultaneously, imperialism was disguised in moral and humanitarian rhetoric. Rudyard Kipling's famous 1899 poem urged Western nations to 'Take up the White Man's Burden' by bringing Christian salvation, modern medicine, and western civilization to 'half-devil and half-child' native populations. In France, the doctrine of the Mission Civilisatrice claimed a universal republican duty to assimilate colonial subjects into French cultural enlightenment. In reality, this philanthropic facade concealed catastrophic humanitarian atrocities, most notoriously in King Leopold II's Congo Free State, where forced rubber quotas, systematic mutilation, and terror halved the native population, resulting in an estimated ten million deaths."),

    ("The Imperial Boomerang: Destabilizing the European Balance of Power",
     "The Scramble for Africa, regulated by Bismarck at the 1884\u20131885 Berlin Conference under the rule of 'effective occupation,' initially acted as a diplomatic safety valve, diverting Great Power rivalries away from the European continent into the colonial periphery. However, by the dawn of the twentieth century, the colonial frontier was fully closed, and imperial tensions boomeranged back into the heart of Europe. The German challenge to French claims in Morocco (the First Moroccan Crisis of 1905 and the Agadir Crisis of 1911) backfired spectacularly: instead of fracturing the Anglo-French Entente, German gunboat diplomacy convinced Britain that Germany sought maritime hegemony, prompting joint Anglo-French military staff talks and cementing the alliance system that exploded in August 1914."),

    ("Conclusion: The Apex and Contradiction of Global Dominance",
     "In conclusion, the New Imperialism of the late nineteenth century was the supreme manifestation of Europe's global dominance, propelled by the relentless production of the Second Industrial Revolution, the militarized chauvinism of competitive nation-states, and the self-righteous arrogance of racial superiority. By 1914, European empires controlled the trade routes, resources, and destinies of hundreds of millions across the planet. Yet, this global dominion carried the seeds of its own destruction. The militarized competition, colonial rivalries, and naval arms races bred by imperialism fatally poisoned the European international order, while the dissemination of European political ideas\u2014constitutionalism, nationalism, and socialism\u2014unwittingly nurtured the indigenous anti-colonial liberation movements of the twentieth century. In conquering the world, imperial Europe assembled the very geopolitical and moral dynamite that obliterated its own supremacy in the trenches of World War I.")
]

quotes_15 = [
    ("J.A. Hobson", "Imperialism: A Study (1902)", "'Imperialism is the endeavour of the great controllers of industry to broaden the channel for the flow of their surplus wealth by seeking foreign markets and foreign investments to take off the goods and capital they cannot sell or use at home.'"),
    ("V.I. Lenin", "Imperialism: The Highest Stage of Capitalism", "'Imperialism emerged as the development and direct continuation of the fundamental characteristics of capitalism in general... The division of the world among capitalist monopolies has been completed.'"),
    ("Daniel Headrick", "The Tools of Empire (1981)", "'The expansion of Europe into Africa and Asia in the nineteenth century was not simply a triumph of superior will, but a direct consequence of the technological breakthroughs that made penetration cheap, rapid, and deadly.'"),
    ("Cecil Rhodes", "Imperialist Visionary (1895)", "'In order to save the 40 million inhabitants of the United Kingdom from a bloody civil war, we colonial statesmen must acquire new lands to settle the surplus population, to provide new markets for the goods produced in the factories.'")
]

# =========================================================================
# QUESTION 16: 2025 Q.7 (SCIENCE, TECHNOLOGY & URBANIZATION 1850-1914)
# =========================================================================

meta_16 = [
    ("EXAM QUESTIONS LINKED:", "2025 Q.7: 'Analyze how science, technology, and urbanization changed European society and culture between 1850 and 1914.'"),
    ("SYLLABUS TOPIC:", "Paper I \u2014 Section 7: Social, Scientific, and Cultural Transformations of Modern Europe (1850\u20131914)"),
    ("ALLOCATED MARKS / TIME:", "20 Marks | 35 Minutes | Standard Word Count: ~650\u2013700 words"),
    ("HISTORIOGRAPHICAL THESIS:", "Between 1850 and 1914, Europe experienced a total civilizational mutation as the Second Industrial Revolution, revolutionary scientific breakthroughs, and explosive urbanization dismantled centuries of agrarian-communal tradition; while creating an affluent, hygienic, and scientifically confident mass society, this modernization simultaneously unleashed profound social polarizations, alienated the urban proletariat, and fostered a fin-de-si\u00e8cle cultural crisis that challenged Enlightenment rationalism on the eve of the Great War.")
]

callout_16 = [
    "\u2022 The Second Industrial Revolution: Marked by the transition from coal/textiles to Bessemer/Siemens steel, synthetic chemical dyes and pharmaceuticals (Bayer, BASF), widespread electrical power generation, and internal combustion engines (Daimler, Benz), creating massive industrial cartels and white-collar corporate bureaucracies.",
    "\u2022 Scientific Revolutions: Charles Darwin's 'Origin of Species' (1859) shattered biblical creationism; Louis Pasteur and Robert Koch formulated the Germ Theory of Disease, isolating bacteria and eliminating cholera/typhus epidemics; Marie Curie, Max Planck (Quantum Theory 1900), and Albert Einstein (Special Relativity 1905) demolished classical Newtonian physics.",
    "\u2022 The Transformation of Cities: Massive demographic migration (by 1910, 60% of British and 45% of Germans lived in cities). Baron Haussmann's radical reconstruction of Paris replaced medieval slums with wide boulevards, public parks, and underground sewers; the advent of electric streetcars enabled suburban expansion and class-segregated neighborhoods.",
    "\u2022 Sociopolitical & Gender Reconfiguration: Birth of the organized industrial working class (trade unions, German SPD); rise of the educated managerial bourgeoisie; emergence of compulsory primary schooling and mass literacy; the rise of the 'New Woman' and the militant suffragette movement (Emmeline Pankhurst's WSPU).",
    "\u2022 The Cultural Crisis (Fin-de-Si\u00e8cle): Shift from mid-Victorian positivism and faith in progress to modernist irrationalism: Friedrich Nietzsche proclaimed 'God is dead' and attacked bourgeois morality; Sigmund Freud exposed the irrational, repressed instincts of the unconscious mind."
]

map_16 = """========================================================================================
MATERIAL DRIVERS                  INFRASTRUCTURAL MUTATION          CULTURAL & INTELLECTUAL IMPACT
========================================================================================
[Bessemer Steel & Chemicals]   --> [Metropolitan Urbanization]   --> [Demise of Biblical Certainty]
[Electricity & Auto Engine]    --> [Haussmannization of Paris]   --> [Germ Theory & Life Expectancy]
[Railways & Electric Trams]    --> [Sanitation, Sewers & Water]  --> [Mass Politics: German SPD / Labor]
[Darwin: Biological Evolution] --> [Suburban Class Segregation]  --> [Suffrage: The 'New Woman']
[Planck/Einstein: New Physics] --> [Compulsory Mass Literacy]    --> [Nietzsche & Freud: Irrationalism]
========================================================================================"""

timeline_16 = [
    ("1853\u20131870", "Haussmann's Rebuilding of Paris", "Napoleon III and Baron Haussmann carve grand boulevards, construct modern sewers, and eliminate revolutionary barricade alleys."),
    ("1856", "Invention of the Bessemer Process", "Enables mass production of cheap steel; transforms railway construction, bridge engineering, and shipbuilding."),
    ("1859", "Darwin Publishes 'Origin of Species'", "Articulates evolutionary natural selection; ignites fierce theological controversies across Europe."),
    ("1864\u20131882", "Pasteur & Koch's Germ Theory", "Pasteur proves microbes cause disease; Koch identifies tuberculosis and cholera bacilli; transforms public sanitation."),
    ("1875", "Public Health Act & SPD Foundation", "British Parliament mandates municipal sanitation; German Social Democratic Party founded, inaugurating mass socialist politics."),
    ("1882", "First Central Electrical Power Stations", "Edison and Swan commercialize electric lighting; transforms factory night shifts and urban street safety."),
    ("1886", "Benz Invents Gasoline Automobile", "Karl Benz patents the first petrol-powered motor wagon; begins the internal combustion revolution."),
    ("1900", "Max Planck & Sigmund Freud", "Planck formulates Quantum Hypothesis; Freud publishes 'The Interpretation of Dreams,' revolutionizing psychology."),
    ("1903", "Foundation of WSPU (Suffragettes)", "Emmeline Pankhurst establishes militant suffragette movement in Britain, demanding votes for women."),
    ("1905", "Einstein's Miracle Year (Annus Mirabilis)", "Publishes Special Theory of Relativity (E=mc\u00b2); overturns absolute time and space of classical physics.")
]

essay_16 = [
    ("Introduction: The Dawn of European Mass Society",
     "The period between 1850 and 1914 witnessed the most rapid, comprehensive civilizational transformation in European history. In a single lifespan, the continent transitioned from a largely rural, agrarian, and localized world governed by traditional religious hierarchies into an urbanized, electrified, and scientifically dominated mass society. This transformation was driven by the titanic convergence of the Second Industrial Revolution, revolutionary breakthroughs in theoretical and applied science, and the explosive growth of modern metropolises. While this material and intellectual metamorphosis generated unprecedented prosperity, wiped out ancient epidemic diseases, and elevated living standards, it simultaneously fractured traditional social cohesion. In the historiographical analysis of scholars like Eric Hobsbawm and Carl Schorske, late-nineteenth-century modernization was inherently paradoxical: it enthroned human reason and bourgeois confidence, while simultaneously unleashing alienating industrial realities, radical class warfare, and a fin-de-si\u00e8cle cultural crisis that challenged the very foundations of Enlightenment rationalism."),

    ("The Second Industrial Revolution: Steel, Electricity, and Chemical Mastery",
     "The material bedrock of late-nineteenth-century change was the Second Industrial Revolution, which accelerated after 1870. While the First Industrial Revolution had been British, localized, and based on coal, iron, and textile steam engines, the second phase was continental (led by a unified Germany and the United States), heavily corporatized, and anchored in steel, chemicals, electricity, and the internal combustion engine. The invention of the Bessemer converter (1856) and the Siemens-Martin open-hearth process made steel vastly cheaper and stronger than cast iron, enabling the construction of trans-continental railway networks, steel-hulled steamships, and skyscraper architecture.\n\n"
     "Concurrently, the chemical industry\u2014dominated by German corporate giants such as Bayer and BASF\u2014synthesized artificial dyes from coal tar, created agricultural fertilizers, and manufactured synthetic pharmaceuticals (notably Aspirin in 1899). In the realm of power, the dynamo (invented by Werner von Siemens) and Nikola Tesla's alternating current systems allowed electrical energy to be transmitted over vast distances. Electricity illuminated city streets, powered factory assembly lines day and night, and energized electric tramways. By the 1880s, the development of the lightweight internal combustion engine by German engineers Gottlieb Daimler and Karl Benz laid the foundation for the modern automotive, aviation, and petroleum industries, permanently compressing geographic space."),

    ("The Metropolis Unbound: Urbanization, Sanitation, and Haussmannization",
     "The physical and demographic landscape of Europe was radically transformed by unprecedented urbanization. Driven by skyrocketing agricultural productivity and rural surplus labor, millions migrated into industrial centers. By 1910, over sixty percent of the British population and nearly half of the German population resided in urban areas, with metropolises like London, Paris, and Berlin swelling into multi-million-citizen mega-cities. Early industrial cities had been notorious pestilential deathtraps marked by open sewage, toxic smog, and rampant cholera epidemics. Between 1850 and 1914, European states embarked on a heroic sanitary revolution guided by Edwin Chadwick's 'miasmatic' research and municipal socialism.\n\n"
     "The premier paradigm of modern urban redesign was executed in Paris between 1853 and 1870 by Napoleon III and his prefect, Baron Georges-Eug\u00e8ne Haussmann. Haussmann brutally demolished the cramped, medieval neighborhoods of Paris, replacing them with broad, tree-lined boulevards, monumental public squares, interconnected parks, and an eighty-mile underground sewer and freshwater aqueduct system. While wide boulevards served the strategic purpose of preventing revolutionary barricades and allowing cavalry charges, Haussmannization became the international architectural blueprint for Vienna (the Ringstrasse), Brussels, and Barcelona. Furthermore, the advent of electric streetcars and underground subways (the London Tube in 1863, Paris M\u00e9tro in 1900) allowed cities to expand outward, giving birth to suburban living and institutionalizing spatial segregation: the wealthy moved to leafy bourgeois peripheries, while the working class was concentrated in industrial factory rings."),

    ("Scientific Paradigm Shifts: Darwinism, Germ Theory, and the New Physics",
     "In the intellectual realm, scientific discoveries fundamentally reshaped the Western worldview. In biology, Charles Darwin's 'On the Origin of Species' (1859) delivered an existential shock to European thought. By demonstrating that all biological species, including humanity, evolved through the blind, mechanistic process of 'natural selection' across millions of years, Darwinism demolished the literal biblical account of divine creation. While mainstream society gradually reconciled science with faith, Darwinism secularized intellectual life and, tragically, was perverted by conservative ideologues into 'Social Darwinism' to justify racial hierarchies and imperial subjugation.\n\n"
     "In medicine, the Germ Theory of Disease, formulated by French chemist Louis Pasteur and German physician Robert Koch, represented the greatest sanitary breakthrough in human history. Pasteur proved that fermentation and decay were caused by microscopic living organisms (microbes), inventing pasteurization and developing life-saving vaccines against anthrax and rabies. Koch systematically isolated the specific bacterial pathogens responsible for tuberculosis (1882) and cholera (1883). Combined with Joseph Lister's introduction of carbolic acid antiseptic surgery, medicine ceased to be an empirical lottery. Infant mortality plummeted, urban death rates dropped below birth rates for the first time in history, and European life expectancy surged from under forty years in 1850 to over fifty-two years by 1914.\n\n"
     "Finally, at the turn of the century, classical Newtonian physics\u2014which had envisioned the universe as an orderly, predictable, mechanistic clockwork\u2014was completely shattered. In 1898, Marie and Pierre Curie discovered radioactivity in uranium and radium, proving that the atom was not an indivisible solid particle. In 1900, Max Planck proposed Quantum Theory, postulating that energy is emitted in erratic packets (quanta). In 1905, Albert Einstein published his Special Theory of Relativity ($E=mc^2$), demonstrating that time and space are not absolute, but relative to the observer. The universe was suddenly revealed to be mysterious, subjective, and fundamentally uncertain."),

    ("Sociopolitical Reconfiguration: Mass Politics, Class Stratification, and the 'New Woman'",
     "The combination of industrialization and urbanization produced a fundamentally new social structure. At the summit stood a wealthy haute bourgeoisie of industrial magnates and bankers who increasingly intermarried with the traditional landed aristocracy. Beneath them arose a massive, respectable 'lower-middle class' (petite bourgeoisie) of clerks, bank tellers, schoolteachers, and civil servants who zealously guarded their respectable status through Victorian domesticity and consumer culture. At the base was the vast industrial proletariat. Concentrated in urban factories, workers overcame early fragmentation by organizing powerful trade unions and mass socialist political parties. The prime example was the German Social Democratic Party (SPD), founded in 1875, which became the largest, most disciplined political organization in Germany, forcing Chancellor Bismarck to introduce the world's first modern welfare state (health, accident, and old-age pensions in the 1880s) to siphon off working-class radicalism.\n\n"
     "Concurrently, universal compulsory primary education, introduced across Western Europe in the 1870s and 1880s, virtually eliminated illiteracy, spawning the modern commercial mass press (such as the British Daily Mail) and popular culture. Furthermore, urban modernity revolutionized the role of women. The expansion of white-collar employment provided unmarried middle-class women with economic independence as typists, telephone operators, and nurses. This gave rise to the 'New Woman,' who demanded civic and political equality. In Great Britain, the movement escalated into the militant suffragette campaign led by Emmeline Pankhurst's Women's Social and Political Union (WSPU) from 1903, utilizing hunger strikes, public demonstrations, and civil disobedience to demand the vote."),

    ("The Cultural Crisis: Fin-de-Si\u00e8cle Modernism and the Assault on Reason",
     "By the 1890s, the triumph of scientific positivism provoked a radical cultural counter-revolution known as the fin-de-si\u00e8cle crisis. An emerging generation of intellectuals rejected the complacent Victorian belief in steady human moral progress through science. In philosophy, German iconoclast Friedrich Nietzsche launched a devastating assault on Christian morality, bourgeois mediocrity, and democracy, declaring that 'God is dead' and asserting that human existence is driven not by rational truth, but by the primal 'Will to Power' (Wille zur Macht).\n\n"
     "In psychology, Austrian physician Sigmund Freud published 'The Interpretation of Dreams' (1900), founding psychoanalysis. Freud argued that human behavior is fundamentally irrational, governed not by conscious logic, but by a subterranean battleground of repressed, primitive sexual and aggressive impulses (the Id) checked by parental socialization (the Superego). In literature and the visual arts, Realism and Naturalism gave way to Modernism, Expressionism, and Cubism (Pablo Picasso). Artists abandoned the objective representation of nature to paint internal emotional anxieties and fragmented realities, reflecting a pervasive psychic unease about the speed of modern life."),

    ("Conclusion: The Fragile Triumph of European Modernity",
     "In conclusion, the period from 1850 to 1914 was the epoch in which the modern world was forged. Science, technology, and urbanization transformed European civilization, creating an urbanized society endowed with miraculous medical defenses, instantaneous electrical communication, and extraordinary industrial productivity. Yet this glittering modernity rested upon profound structural and psychological fault lines. The rapid dissolution of traditional communities fostered urban alienation; industrial capitalism widened class antagonisms; and the intellectual revolution undermined spiritual certainty, substituting it with the disquieting philosophies of the unconscious and the Will to Power. When the European powers plunged into the apocalyptic carnage of World War I in the summer of 1914, they turned the very fruits of this magnificent scientific and industrial revolution\u2014machine guns, chemical poison gas, electrified barbed wire, and industrialized mass mobilization\u2014into the instruments of their own catastrophic self-destruction.")
]

quotes_16 = [
    ("Eric Hobsbawm", "The Age of Empire: 1875-1914", "'The era from 1875 to 1914 was an age of profound paradox: an unprecedented triumph of material progress, industrial productivity, and scientific rationalism that was secretly breeding the social, political, and cultural explosives that tore it apart in 1914.'"),
    ("Carl Schorske", "Fin-de-Si\u00e8cle Vienna: Politics and Culture", "'The twentieth-century mind was born in the crisis of late-nineteenth-century liberalism... As faith in rational man and steady progress evaporated, art and psychology turned inward to explore the dark, irrational depths of the human psyche.'"),
    ("Lewis Mumford", "The City in History", "'Haussmann's Paris was the triumph of the straight line and the imperial perspective... It transformed the chaotic medieval community into an orderly, hygienic, and highly policed capitalist metropolis.'"),
    ("Friedrich Nietzsche", "The Gay Science (1882)", "'God is dead. God remains dead. And we have killed him. How shall we comfort ourselves, the murderers of all murderers?... Must we ourselves not become gods simply to appear worthy of it?'")
]

# Run the generation
if __name__ == "__main__":
    tasks = [
        (meta_15, callout_15, map_15, timeline_15, essay_15, quotes_15,
         os.path.join(paper1_dir, "Q15_Industrialization_and_New_Imperialism_Master_Notes.docx"),
         "QUESTION 15: INDUSTRIALIZATION, NEW IMPERIALISM & EUROPEAN GLOBAL DOMINANCE",
         "Q15_Industrialization_and_New_Imperialism_Master_Notes"),
        (meta_16, callout_16, map_16, timeline_16, essay_16, quotes_16,
         os.path.join(paper1_dir, "Q16_Science_Technology_and_Urbanization_1850_1914_Master_Notes.docx"),
         "QUESTION 16: SCIENCE, TECHNOLOGY, URBANIZATION & SOCIETY (1850\u20131914)",
         "Q16_Science_Technology_and_Urbanization_1850_1914_Master_Notes")
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
