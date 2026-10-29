import os
import subprocess
import tempfile

EDGE_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

def render_html_to_pdf(html_content, output_pdf_path):
    with tempfile.NamedTemporaryFile("w", delete=False, suffix=".html", encoding="utf-8") as f:
        f.write(html_content)
        temp_html = f.name
    
    cmd = [
        EDGE_PATH,
        "--headless",
        "--disable-gpu",
        f"--print-to-pdf={output_pdf_path}",
        "--no-pdf-header-footer",
        temp_html
    ]
    subprocess.run(cmd, check=True)
    try:
        os.remove(temp_html)
    except Exception:
        pass
    print(f"Generated: {output_pdf_path}")

HTML_TEMPLATE = """<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  @page {
    size: A4;
    margin: 20mm 18mm 20mm 18mm;
  }
  body {
    font-family: "Times New Roman", Times, serif;
    color: #111;
    line-height: 1.45;
    font-size: 13.5pt;
  }
  .header {
    text-align: center;
    border-bottom: 2px solid #000;
    padding-bottom: 12px;
    margin-bottom: 20px;
  }
  .header h2 {
    margin: 0 0 4px 0;
    font-size: 16pt;
    letter-spacing: 0.5px;
    text-transform: uppercase;
  }
  .header h3 {
    margin: 0 0 6px 0;
    font-size: 13pt;
    font-weight: normal;
  }
  .header .subject {
    font-size: 14pt;
    font-weight: bold;
    text-transform: uppercase;
    margin: 6px 0;
  }
  .meta-table {
    width: 100%;
    margin-top: 8px;
    font-weight: bold;
    font-size: 12pt;
    border-collapse: collapse;
  }
  .meta-table td {
    padding: 2px 0;
  }
  .note-box {
    margin: 14px 0 20px 0;
    padding: 8px 12px;
    border: 1px dashed #444;
    font-style: italic;
    font-size: 12pt;
    background: #fbfbfb;
  }
  .questions-list {
    margin-top: 15px;
  }
  .q-item {
    margin-bottom: 18px;
    display: flex;
    align-items: baseline;
  }
  .q-num {
    font-weight: bold;
    min-width: 45px;
    flex-shrink: 0;
  }
  .q-text {
    flex-grow: 1;
    text-align: justify;
  }
  .q-marks {
    font-weight: bold;
    margin-left: 10px;
    flex-shrink: 0;
  }
  .sub-item {
    margin-top: 6px;
    margin-left: 20px;
  }
  .footer {
    margin-top: 35px;
    text-align: center;
    font-size: 11pt;
    border-top: 1px solid #aaa;
    padding-top: 8px;
    color: #444;
  }
</style>
</head>
<body>
  <div class="header">
    <h2>Khyber Pakhtunkhwa Public Service Commission</h2>
    <h3>Competitive Examination for the Posts of Provincial Management Service (BPS-17)</h3>
    <div class="subject">{subject_title} ({year})</div>
    <table class="meta-table">
      <tr>
        <td style="text-align: left;">TIME ALLOWED: 03 HOURS</td>
        <td style="text-align: right;">MAXIMUM MARKS: 100</td>
      </tr>
    </table>
  </div>

  <div class="note-box">
    <strong>Note:</strong> Attempt any <strong>FIVE</strong> questions. All questions carry equal marks (20 Marks each).
  </div>

  <div class="questions-list">
    {questions_html}
  </div>

  <div class="footer">
    Khyber Pakhtunkhwa Public Service Commission &bull; Provincial Management Service (PMS)
  </div>
</body>
</html>
"""

def make_paper_html(subject_title, year, questions):
    q_html_parts = []
    for i, q in enumerate(questions, 1):
        if isinstance(q, tuple):
            main_q, sub_items = q
            sub_html = "".join([f"<div class='sub-item'><strong>({chr(97+k)})</strong> {s}</div>" for k, s in enumerate(sub_items)])
            item = f"""<div class="q-item">
              <div class="q-num">Q.{i}</div>
              <div class="q-text">{main_q}{sub_html}</div>
              <div class="q-marks">(20)</div>
            </div>"""
        else:
            item = f"""<div class="q-item">
              <div class="q-num">Q.{i}</div>
              <div class="q-text">{q}</div>
              <div class="q-marks">(20)</div>
            </div>"""
        q_html_parts.append(item)
    
    content = HTML_TEMPLATE.replace("{subject_title}", subject_title)
    content = content.replace("{year}", year)
    content = content.replace("{questions_html}", "\n".join(q_html_parts))
    return content

paper1_dir = r"H:\My Drive\PMS KPK 2022 notes\1. European History Paper 1 and Paper 2\Paper 1"
paper2_dir = r"H:\My Drive\PMS KPK 2022 notes\1. European History Paper 1 and Paper 2\Paper 2"

# ----------------- PAPER 1 PAPERS -----------------

p1_2025 = [
    "Examine the internal and external factors that contributed to the rise of Napoleon Bonaparte. How far did Napoleon's rule continue or betray the ideals of the French Revolution?",
    "Examine how the Metternich system sought to suppress revolutionary movements in post-Napoleonic Europe. Why did it ultimately fail?",
    "Evaluate the effectiveness of the Congress of Vienna in maintaining peace in Europe between 1815 and 1848. To what extent did it succeed in reconciling conservative principles with the force of change?",
    "Assess the role of nationalism and realpolitik in the unification of Italy and Germany. How did the leadership of Cavour and Bismarck redefine the balance of power in Europe?",
    "Discuss the relationship between industrialization, imperialism, and nationalism in shaping Europe's global dominance during the 19th century.",
    "What were the primary economic, political, and ideological motivations behind the 'New Imperialism' of the late 19th century?",
    "Analyze how science, technology, and urbanization changed European society and culture between 1850 and 1914.",
    "Explain the long-term causes of World War I (Militarism, Alliances, Imperialism, Nationalism). How did the July Crisis of 1914 act as the immediate trigger for the conflict?"
]

p1_2022 = [
    "'Revolutions often start from the streets but eventually get completed in legislatures.' In the light of this statement, assess the role of the National Assembly in France between 1789 and 1791.",
    "In the light of his internal reforms, discuss whether Napoleon Bonaparte promoted or halted the French Revolution.",
    "What were the principles adopted at the Congress of Vienna (1814–1815)? How were those principles implemented in the decisions of the Congress?",
    "Critically analyze the role of Count Cavour in the Unification of Italy.",
    "Highlight the causes and consequences of the Crimean War (1853–1856). Do you agree that it was a 'futile and unnecessary carnage'?",
    "How did Bismarck use war, diplomacy, and propaganda to unify Germany?",
    "Discuss the role of the Alliance System and the Arms Race as two major causes of World War I.",
    ("Assess the contribution of the following philosophers in changing the political thought of Europe:", [
        "John Locke",
        "Jean-Jacques Rousseau"
    ])
]

p1_2018 = [
    "Critically examine the social, economic, and intellectual causes of the French Revolution of 1789.",
    "Analyze the Continental System of Napoleon Bonaparte. Why did it fail to bring Great Britain to its knees, and how did it contribute to Napoleon's downfall?",
    "Describe the Concert of Europe and evaluate its achievements and failures from 1815 to 1848.",
    "Discuss the Revolution of 1848 in France and explain why the revolutionary wave engulfed most of Europe.",
    "Evaluate the foreign policy of Napoleon III with special reference to his role in Italian unification and the Franco-Prussian War.",
    "Bismarck unified Germany not with speeches and majority decisions, but by 'Blood and Iron'. Elaborate.",
    "Trace the development of the Eastern Question from the Greek War of Independence (1821) to the Treaty of Berlin (1878).",
    "How far was the Triple Entente responsible for dividing Europe into two armed camps before 1914?"
]

p1_2016 = [
    "Discuss in detail the philosophy of 'Liberty, Equality, and Fraternity' that motivated the French people to revolt against monarchical powers in 1789.",
    "Napoleon Bonaparte declared: 'I am the Revolution.' To what extent is this claim justified by his civil, judicial, and educational reforms?",
    "Critically assess Prince Metternich as the prime defender of the conservative order in Europe from 1815 to 1848.",
    "Compare the roles of Giuseppe Mazzini, Count Camillo Cavour, and Giuseppe Garibaldi in the Risorgimento (Unification of Italy).",
    "Analyze the diplomatic maneuvering of Otto von Bismarck in the three successive wars against Denmark (1864), Austria (1866), and France (1870–71).",
    "Examine the causes, course, and historical significance of the Crimean War (1853–1856) in European diplomacy.",
    "What was the significance of the Congress of Berlin (1878) in resolving the Balkan crisis? Did it prevent or sow the seeds of future conflict?",
    "Examine the complex network of secret alliances and imperialism from 1890 to 1914 that made the First World War inevitable."
]

p1_2013 = [
    "The conditions in Europe were always the same except in England. But why did the revolution (1789) occur in France and not in any other country?",
    "Napoleonic reforms prove him to be a great administrator and a statesman. Some of them have existed nowadays in France. Discuss.",
    "Continental System became a catalyst for the downfall of Napoleon Bonaparte. Discuss.",
    "The defeat of Napoleon Bonaparte in Waterloo brought the dark ages of old order under the Metternich system. Discuss.",
    "Describe the effects of the Revolution of 1848, which transformed the whole European world.",
    "German unification started with Napoleon and ended with Bismarck. Discuss.",
    "Explain the role of Garibaldi, Mazzini, and Cavour in the Unification of Italy.",
    "Alliances and counter-alliances led the European world towards the First World War. Discuss."
]

# ----------------- PAPER 2 PAPERS -----------------

p2_2025 = [
    "Analyze the First World War as a 'total war,' examining its impact on European societies, economies, and political systems, and evaluate the consequences of the Treaty of Versailles as a long-term cause of future conflict.",
    "Investigate the Russian Revolution, detailing the causes of the Bolshevik seizure of power, the subsequent civil war, and the consolidation of a totalitarian state under Lenin and Stalin.",
    "Discuss the aims and achievements of the League of Nations. To what extent did its structural weaknesses and the policies of major powers lead to its failure?",
    "Compare and contrast the rise of Fascism in Italy and Nazism in Germany. Analyze their core ideologies, methods of control, and the social and economic conditions that facilitated their ascent to power during the interwar period.",
    "Examine the causes and major turning points of the Second World War, including the failure of the policy of appeasement, the nature of Nazi aggression, and the Holocaust.",
    "Examine the origins and evolution of the Cold War. How did ideological differences and power politics divide Europe after 1945?",
    "Assess the impact of the Soviet Union's domination of Eastern Europe (1945–1960). How did it consolidate its power and respond to challenges like the Hungarian Uprising (1956)?",
    "To what extent did the success of the Common Market inspire later movements toward deeper European integration, culminating in the European Union?"
]

p2_2022 = [
    "Highlighting the role of Ottoman Turkey in the First World War, critically analyze the treatment of Turkey by the Allied Powers in the post-war treaties (Treaty of Sèvres and Treaty of Lausanne).",
    "Critically examine the Treaty of Versailles (1919). Was it a dictated peace (Diktat) that made the rise of Hitler inevitable?",
    "Discuss the ideological principles, economic policies, and totalitarian structure of Fascist Italy under Benito Mussolini.",
    "Trace the causes and repercussions of the Great Depression of 1929 on European democracies, particularly Britain and Germany.",
    "Analyze British and French policy of Appeasement towards Nazi Germany from 1935 to 1939. Why did it fail to avert World War II?",
    "Examine the major wartime conferences (Tehran, Yalta, and Potsdam) and discuss how they shaped the post-WWII geopolitical map of Europe.",
    "Evaluate the creation, structure, and early peacekeeping challenges of the United Nations Organization (UNO) compared to the League of Nations.",
    "Discuss the division of Germany and the Berlin Crisis (1948–49) as central flashpoints in the emergence of the Cold War."
]

p2_2018 = [
    "'The peace settlement of 1919 was the monument of peacemakers' inaptitude and short sightedness.' Discuss.",
    "Write a detailed note on the Russian Revolution (1917) and bring out its causes and results.",
    "What were the causes of failure of the League of Nations?",
    "'It is said that France's hunt for security and England's desire for peace were responsible for the rise of Hitler.' Do you agree with this statement?",
    "Review the foreign policy of Britain between two World Wars.",
    "Describe the Foreign Policy of France between 1919–1939.",
    "Analyze the consequences and impact of World War II.",
    "Discuss the establishment and importance of the United Nations Organization."
]

p2_2016 = [
    "Elaborate in detail on the Bolshevik Revolution in Tsarist Russia and evaluate the role of Lenin as a revolutionary leader.",
    "Examine the provisions of the Treaty of Versailles with special reference to war guilt clause, disarmament, and territorial adjustments imposed on Germany.",
    "Discuss the mandate system and disarmament conferences under the auspices of the League of Nations. Why did disarmament fail in the 1930s?",
    "Analyze the domestic transformation of Germany under Adolf Hitler from the Weimar Republic to the Third Reich (1933–1939).",
    "Critically assess the Spanish Civil War (1936–1939) as a testing ground for European dictators and a prelude to World War II.",
    "Discuss the causes, milestones, and devastation of the Second World War in Europe.",
    "Analyze the Truman Doctrine and the Marshall Plan as fundamental instruments of American foreign policy in post-war Europe.",
    "Evaluate the emergence of NATO and the Warsaw Pact in consolidating the bipolar division of Europe."
]

p2_2013 = [
    "Evaluate the underlying and fundamental causes that led to the outbreak of the First World War.",
    "Woodrow Wilson's Fourteen Points promised a new international order based on justice. How far were these points reflected in the Paris Peace Settlement of 1919?",
    "Trace the rise of Benito Mussolini in Italy and evaluate his domestic socio-economic reforms.",
    "Analyze the internal crises of the Weimar Republic that facilitated the electoral rise of the National Socialist (Nazi) Party.",
    "Discuss the foreign policy of the Soviet Union under Joseph Stalin between 1928 and 1939, culminating in the Nazi-Soviet Non-Aggression Pact.",
    "Examine the collapse of the League of Nations with specific reference to the Japanese invasion of Manchuria and the Italian invasion of Abyssinia.",
    "Analyze the impact of World War II on the global status of Great Britain and the collapse of the European colonial empires.",
    "Discuss the origins of the Cold War and the ideological confrontation between the United States and the USSR in Europe from 1945 to 1955."
]

def generate_all():
    # Paper 1
    p1_items = [
        ("2025", p1_2025),
        ("2022", p1_2022),
        ("2018", p1_2018),
        ("2016", p1_2016),
        ("2013", p1_2013),
    ]
    for yr, qs in p1_items:
        html = make_paper_html("European History &bull; Paper I (1789–1914)", yr, qs)
        out_pdf = os.path.join(paper1_dir, f"KPPSC_PMS_European_History_Paper_1_{yr}.pdf")
        render_html_to_pdf(html, out_pdf)
    
    # Paper 2
    p2_items = [
        ("2025", p2_2025),
        ("2022", p2_2022),
        ("2018", p2_2018),
        ("2016", p2_2016),
        ("2013", p2_2013),
    ]
    for yr, qs in p2_items:
        html = make_paper_html("European History &bull; Paper II (1914–1960)", yr, qs)
        out_pdf = os.path.join(paper2_dir, f"KPPSC_PMS_European_History_Paper_2_{yr}.pdf")
        render_html_to_pdf(html, out_pdf)

if __name__ == "__main__":
    generate_all()
