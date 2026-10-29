import os
import subprocess

BASE_DIR = r"H:\My Drive\PMS KPK 2022 notes\2. International Relations"
PAPERS_DIR = os.path.join(BASE_DIR, "Past Papers")
os.makedirs(PAPERS_DIR, exist_ok=True)

EDGE_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

def compile_pdf(md_path, pdf_path, title):
    html_content = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>{title}</title>
<style>
    @page {{
        size: A4;
        margin: 18mm 16mm 18mm 16mm;
    }}
    body {{
        font-family: 'Segoe UI', Calibri, Arial, sans-serif;
        font-size: 10pt;
        line-height: 1.45;
        color: #1A1A1A;
        margin: 0;
    }}
    h1 {{
        color: #0E2B50;
        font-size: 16pt;
        border-bottom: 2px solid #0E2B50;
        padding-bottom: 4px;
        margin-top: 0;
        margin-bottom: 10px;
        text-align: center;
    }}
    h2 {{
        color: #1B4965;
        font-size: 12pt;
        margin-top: 14px;
        margin-bottom: 6px;
        border-bottom: 1px solid #D0D7DE;
        padding-bottom: 2px;
    }}
    h3 {{
        color: #2B2D42;
        font-size: 10.5pt;
        margin-top: 10px;
        margin-bottom: 4px;
    }}
    table {{
        width: 100%;
        border-collapse: collapse;
        margin: 10px 0;
        font-size: 9pt;
    }}
    th {{
        background-color: #0E2B50;
        color: #FFFFFF;
        padding: 5px 8px;
        text-align: left;
        font-weight: bold;
        border: 1px solid #0E2B50;
    }}
    td {{
        padding: 4px 8px;
        border: 1px solid #D0D7DE;
        vertical-align: top;
    }}
    tr:nth-child(even) {{
        background-color: #F8F9FA;
    }}
    blockquote {{
        border-left: 4px solid #0E2B50;
        background-color: #F0F4F8;
        padding: 6px 12px;
        margin: 8px 0;
        font-size: 9.5pt;
    }}
    ol, ul {{
        margin-top: 4px;
        margin-bottom: 8px;
        padding-left: 22px;
    }}
    li {{
        margin-bottom: 4px;
    }}
    .box {{
        background-color: #F0F4F8;
        border: 1px solid #C8D6E5;
        border-radius: 4px;
        padding: 10px 14px;
        margin: 10px 0;
    }}
    .footer {{
        font-size: 8pt;
        color: #6C757D;
        text-align: center;
        margin-top: 20px;
        border-top: 1px solid #E1E4E8;
        padding-top: 5px;
    }}
</style>
</head>
<body>
"""
    # Simple markdown to html
    with open(md_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    in_table = False
    in_block = False
    
    for line in lines:
        stripped = line.strip()
        if not stripped:
            if in_block:
                html_content += "</blockquote>\n"
                in_block = False
            continue
        
        if stripped.startswith("# "):
            html_content += f"<h1>{stripped[2:]}</h1>\n"
        elif stripped.startswith("## "):
            html_content += f"<h2>{stripped[3:]}</h2>\n"
        elif stripped.startswith("### "):
            html_content += f"<h3>{stripped[4:]}</h3>\n"
        elif stripped.startswith("> "):
            if not in_block:
                html_content += "<blockquote>\n"
                in_block = True
            html_content += f"<p>{stripped[2:]}</p>\n"
        elif stripped.startswith("|") and stripped.endswith("|"):
            if "---" in stripped:
                continue
            cells = [c.strip() for c in stripped.split("|")[1:-1]]
            if not in_table:
                html_content += "<table>\n<thead><tr>\n"
                for c in cells:
                    html_content += f"<th>{c}</th>\n"
                html_content += "</tr></thead>\n<tbody>\n"
                in_table = True
            else:
                html_content += "<tr>\n"
                for c in cells:
                    html_content += f"<td>{c}</td>\n"
                html_content += "</tr>\n"
        else:
            if in_table:
                html_content += "</tbody></table>\n"
                in_table = False
            if in_block:
                html_content += "</blockquote>\n"
                in_block = False
            
            if stripped.startswith("- ") or stripped.startswith("* "):
                html_content += f"<ul><li>{stripped[2:]}</li></ul>\n"
            elif stripped.startswith("1. ") or stripped.startswith("2. ") or stripped.startswith("3. ") or stripped.startswith("4. ") or stripped.startswith("5. ") or stripped.startswith("6. ") or stripped.startswith("7. ") or stripped.startswith("8. "):
                html_content += f"<ol><li>{stripped[3:]}</li></ol>\n"
            elif stripped.startswith("---"):
                html_content += "<hr style='border: 0; border-top: 1px solid #D0D7DE; margin: 12px 0;'/>\n"
            else:
                html_content += f"<p>{stripped}</p>\n"
                
    if in_table:
        html_content += "</tbody></table>\n"
    if in_block:
        html_content += "</blockquote>\n"
        
    html_content += f"""
<div class="footer">KPPSC PMS International Relations Official Examination Suite | Government of Khyber Pakhtunkhwa</div>
</body>
</html>"""

    temp_html = md_path.replace(".md", "_temp.html")
    with open(temp_html, 'w', encoding='utf-8') as f:
        f.write(html_content)
        
    cmd = [
        EDGE_PATH,
        "--headless",
        "--disable-gpu",
        "--run-all-compositor-stages-before-draw",
        f"--print-to-pdf={pdf_path}",
        temp_html
    ]
    subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if os.path.exists(temp_html):
        os.remove(temp_html)
    print(f"Generated PDF: {pdf_path}")

# 1. Official Syllabus
syllabus_md = os.path.join(BASE_DIR, "KPPSC_PMS_International_Relations_Official_Syllabus.md")
syllabus_pdf = os.path.join(BASE_DIR, "KPPSC_PMS_International_Relations_Official_Syllabus.pdf")

syllabus_text = """# KHYBER PAKHTUNKHAWA PUBLIC SERVICE COMMISSION (KPPSC)
## PROVINCIAL MANAGEMENT SERVICE (PMS) COMPETITIVE EXAMINATION
### SYLLABUS FOR OPTIONAL SUBJECT: INTERNATIONAL RELATIONS (100 MARKS)

> **Subject Group**: Group 6 (Optional Subjects)  
> **Total Marks**: 100 Marks | **Time Allowed**: 3 Hours  
> **Examination Pattern**: Descriptive / Subjective Examination (Attempt ANY 5 Questions out of 8, each carrying 20 Marks)  
> **Statutory Combination Rule**: Candidates opting for *International Relations* **CANNOT** opt for *International Law*.

---

### PART A: Foundational Theories, Core Concepts & The State System

1. **The Modern State System**:
   - Historical evolution from the Treaty of Westphalia (1648) to the contemporary era.
   - Core Westphalian features: Territorial integrity, internal sovereignty, external independence, and diplomatic recognition.
   - Evolving forms and contemporary challenges: Supranational integration (EU), globalization, failed/fragile states, and the rise of non-state actors.

2. **International Relations as a Field of Study & Basic Approaches**:
   - Scope and disciplinary evolution of International Relations.
   - Foundational theoretical traditions:
     - **Realism**: Classical Realism (Thucydides, Machiavelli, Morgenthau) vs. Structural Realism / Neorealism (Waltz, Mearsheimer - Defensive vs. Offensive).
     - **Liberalism**: Classical Liberalism, Neoliberal Institutionalism (Keohane, Nye), Democratic Peace Theory, and Complex Interdependence.
     - **Constructivism**: Social construction of anarchy, identity, norms, and culture (Alexander Wendt).
     - **Critical & Marxist Perspectives**: Dependency Theory, World-Systems Analysis (Wallerstein).
   - Core Concepts:
     - **Power**: Elements of national power (tangible vs. intangible); Hard Power, Soft Power (Joseph Nye), and Smart Power.
     - **Balance of Power**: Meaning, theoretical foundations, historical models, and instruments (alliances, armaments, buffer states, partition).
     - **Imperialism**: Historical colonialism vs. modern imperialism; economic theories of Hobson and Lenin; neo-colonialism and dependency.
     - **Nationalism**: Historical rise, civic vs. ethnic nationalism, anti-colonial nationalism, and contemporary populist resurgence.

3. **Interaction Between States**:
   - **Diplomacy**: Evolution of diplomatic institutions, bilateral vs. multilateral diplomacy, summit diplomacy, public diplomacy, coercive diplomacy, and track-two channels.
   - **International Law**: Nature, historical development, sources (treaties, customary law, general principles, judicial decisions), subjects, enforcement mechanisms, and systemic limitations in an anarchic order.
   - **International Economic and Trade Linkages**: Mercantilism, economic liberalism, and structuralism; global trade regimes (WTO, GATT), international financial institutions (IMF, World Bank), trade wars, economic sanctions, and geo-economic competition.

---

### PART B: Historical Evolution, Cold War Dynamics & Nuclear Security

1. **Evolution of the International System Since 1815**:
   - The Concert of Europe (1815–1854): Balance of power and Congress diplomacy.
   - The Bismarckian alliance system and the collapse of multipolar equilibrium.

2. **Major Global Conflicts & Systemic Transformations**:
   - **World War I (1914–1918)**: Systemic causes (militarism, alliances, imperialism, nationalism), breakdown of crisis management, and the Paris Peace Settlement (Treaty of Versailles 1919).
   - **World War II (1939–1945)**: Interwar instability, failure of collective security (League of Nations), appeasement, and total war dynamics.
   - **Decolonization & The Third World**: Disintegration of European empires in Asia and Africa, post-colonial state vulnerabilities, and the emergence of the Global South.

3. **Cold War Geopolitics & Bipolar Confrontation**:
   - Origins and structural nature of bipolarity (USA vs. USSR).
   - Military Alliance Systems: NATO (1949), Warsaw Pact (1955), SEATO (1954), and CENTO (1955).
   - Ideological warfare, proxy conflicts (Korea, Vietnam, Afghanistan), and crises (Berlin, Cuban Missile Crisis).
   - The doctrine of **Peaceful Coexistence** and the era of **Détente** (1969–1979).
   - **The Non-Aligned Movement (NAM)**: Origins (Bandung Conference 1955, Belgrade 1961), principles of non-alignment, historical achievements during the Cold War, and relevance in the post-Cold War multipolar era.

4. **Peace and Security in the Nuclear Age**:
   - Theories of Peace and Security: Collective security vs. balance of power vs. deterrence.
   - Nuclear Deterrence Theory: Mutually Assured Destruction (MAD), first-strike vs. second-strike stability, deterrence stability vs. instability paradox.
   - Limited war, escalation control, and crisis management in nuclearized environments.
   - Arms Control and Disarmament: Strategic Arms Limitation Talks (SALT I & II), START, Nuclear Non-Proliferation Treaty (NPT), and Comprehensive Nuclear-Test-Ban Treaty (CTBT).

---

### PART C: Major Powers' Foreign Policies & Contemporary Global / Regional Flashpoints

1. **Foreign Policies of Superpowers and Major Powers**:
   - **United States**: Grand strategy from Containment to Unipolarity; Pivot to Asia, Indo-Pacific Strategy, and strategic partnership with India.
   - **Russian Federation**: Post-Soviet transition, revanchism, Near Abroad doctrine, NATO expansion pushback, and the Ukraine War (2022–present).
   - **People's Republic of China**: Peaceful Rise, Belt and Road Initiative (BRI), global economic statecraft, and balance of power in Asia.
   - **Special Reference to South Asia**: Great power competition and its direct strategic, political, and economic impacts on Pakistan and India.

2. **Contemporary Issues in Global & Regional Politics**:
   - **The Shifting Global Order**: Decline of Western unipolarity, the Rise of China, multipolarity, and the expansion of counter-hegemonic blocs (BRICS, SCO).
   - **South Asian Security Matrix**:
     - India-Pakistan strategic rivalry, Kashmir dispute, and nuclearized crisis dynamics.
     - Hydrological Security & Climate Change: The Indus Waters Treaty (1960), cross-border river disputes, and environmental challenges.
     - Afghanistan Security Crisis: Taliban resurgence (August 2021), cross-border terrorism (TTP/ISKP), and security implications for Khyber Pakhtunkhwa and Balochistan.
   - **Middle Eastern Realignments**:
     - The Abraham Accords (2020) and normalization of Arab-Israeli diplomatic ties.
     - The Israel-Palestine Conflict, viability of the Two-State Solution, and post-October 7 regional fallout.
     - Iran-Saudi reconciliation (Beijing-brokered 2023) and regional security rebalancing.
   - **Geo-economics & Strategic Competition**:
     - US-China Trade Wars, supply chain decoupling, and tariff diplomacy.
     - The Indo-Pacific Strategy vs. China's Maritime Silk Road in the Indian Ocean.
   - **Non-State Actors (NSAs) & Non-Traditional Security**:
     - Transnational terrorism (ISIS, Al-Qaeda), insurgencies, and hybrid warfare.
     - Global pandemics (Covid-19 "New Normal"), global supply chain fragilities, and multilateral institutional paralysis.
"""

with open(syllabus_md, 'w', encoding='utf-8') as f:
    f.write(syllabus_text)
compile_pdf(syllabus_md, syllabus_pdf, "KPPSC PMS International Relations Official Syllabus")

# 2. Past Papers Generation
past_papers_data = [
    {
        "year": "2025",
        "title": "KPPSC PMS International Relations Past Paper 2025",
        "md_file": "KPPSC_PMS_International_Relations_Paper_2025.md",
        "pdf_file": "KPPSC_PMS_International_Relations_Paper_2025.pdf",
        "questions": [
            "Q.1 The rise of China is widely interpreted as a revival of balance of power politics, signaling a global shift toward multipolarity. Analyze the political, strategic and economic factors that are shaping the new international order. (20 Marks)",
            "Q.2 Critically evaluate the resilience of Indus Basin Treaty as a cornerstone of hydrological security in South Asia. To what extent can this agreement continue to safeguard water-sharing arrangements from broader geopolitical tension between India and Pakistan, particularly in an era of increasing water scarcity and climate change? (20 Marks)",
            "Q.3 Analyze the expansion of BRICS and Shanghai Cooperation Organization (SCO) through the lens of structural realism. Are these organizations best understood as nascent challengers to a US-led liberal international order, or are their internal divisions too significant to pose a coherent alternative? (20 Marks)",
            "Q.4 Abraham Accord has been described as re-aligning Middle East geopolitics. Critically assess the impact of this new diplomatic architecture on the viability of the two-state solution for the Israel-Palestine conflict. Does it create new leverage for peace or further marginalize the Palestinian issue? (20 Marks)",
            "Q.5 How does cross-border terrorism affect regional stability? Critically assess the Pakistan-Taliban security crisis. What are the implications of the Taliban resurgence in Afghanistan for Pakistan’s external and internal security environment, particularly in Baluchistan and Khyber Pakhtunkhwa? (20 Marks)",
            "Q.6 Evaluate the interplay of diplomacy and propaganda in shaping ties between states. To what extent do these tools serve as mechanisms of conflict resolution versus instruments of power projection in the modern international system? (20 Marks)"
        ]
    },
    {
        "year": "2022",
        "title": "KPPSC PMS International Relations Past Paper 2022",
        "md_file": "KPPSC_PMS_International_Relations_Paper_2022.md",
        "pdf_file": "KPPSC_PMS_International_Relations_Paper_2022.pdf",
        "questions": [
            "Q.1 Critically examine if the Ukraine War is a local affair or a competition between the two superpowers, and how it is to impact European unity. Give arguments and examples to justify your answer. (20 Marks)",
            "Q.2 Discuss the salient features of the US Indo-Pacific Strategy enunciated in February 2022. Analyze its implications for the South Asian Security Environment with particular reference to Pakistan. (20 Marks)",
            "Q.3 The trade war between the US and China has jeopardized the global economic order impacting all regions and states; Pakistan caught between these two global powers is likely to be affected the most. Analyze Pakistan's predicament and suggest policy options to safeguard its national security objectives. (20 Marks)",
            "Q.4 In the Realist prism, the 'State' is the legitimate actor in the international system, but the emergence of Non-State Actors (NSAs) has greatly impacted the global order; Taliban are governing Afghanistan without recognition and Hezbollah is a major stakeholder in Lebanon. Analyze the impact of NSAs on the global strategic environment with solid arguments and valid examples. (20 Marks)",
            "Q.5 The Chinese Belt and Road Initiative (BRI) has global outreach in Asia, Africa, and Europe at the cost of US hegemony. Discuss its main objectives and analyze challenges and prospects of its success. (20 Marks)",
            "Q.6 Abraham Accords was designed to normalize relations between Israel and the Arab World; however, except for a few Arab States, the idea has not received popular support among the public and leaders. Critically examine the implications of this Accord on the Palestinian issue. (20 Marks)",
            "Q.7 The Covid pandemic has introduced the concept of the 'New Normal' in the psycho-social attitudes of the public across the globe. Evaluate its structural impact on international relations, economic globalization, and global governance. (20 Marks)"
        ]
    },
    {
        "year": "2018",
        "title": "KPPSC PMS International Relations Past Paper 2018",
        "md_file": "KPPSC_PMS_International_Relations_Paper_2018.md",
        "pdf_file": "KPPSC_PMS_International_Relations_Paper_2018.pdf",
        "questions": [
            "Q.1 Critically analyze Joseph Nye's concept of “Soft Power” in International Relations. (20 Marks)",
            "Q.2 “Nationalism is neither inherently good nor bad, it is only a concept that can be used to achieve good or evil ends.” Elucidate. (20 Marks)",
            "Q.3 Balance of power is a state of stability between competing forces; it is an equilibrium among states or alliances to prevent any other entity from becoming too strong. In light of this statement, elaborate how equilibrium was maintained by Great Powers during the Cold War era. (20 Marks)",
            "Q.4 The direct cause of World War I was the assassination of “Archduke Franz Ferdinand.” However, historians feel that a number of factors contributed to the rivalry between Great Powers that led to war on such a wide scale. Comment please. (20 Marks)",
            "Q.5 Discuss the emergence of ISIS and its impact on national, regional, and international peace and security. (20 Marks)",
            "Q.6 There may be a possible trade war between the USA and China that would rattle certain American businesses and force consumers to pay more for products. Discuss the term \"trade war\" and elaborate on why Trump was targeting China. (20 Marks)",
            "Q.7 There are two schools of thought regarding the conjecture of globalization. First, view it as a blessing and a tool for prosperity; the other views it as a tool of frustration. Please elaborate. (20 Marks)",
            "Q.8 Highlight the provisions of the Geneva Convention on Prisoners of War. (20 Marks)"
        ]
    },
    {
        "year": "2016",
        "title": "KPPSC PMS International Relations Past Paper 2016",
        "md_file": "KPPSC_PMS_International_Relations_Paper_2016.md",
        "pdf_file": "KPPSC_PMS_International_Relations_Paper_2016.pdf",
        "questions": [
            "Q.1 Critically examine the evolution of the Modern State System from the Peace of Westphalia (1648). How is traditional state sovereignty being redefined by globalization and transnational forces? (20 Marks)",
            "Q.2 Compare and contrast Classical Realism (Hans Morgenthau) and Structural Realism / Neorealism (Kenneth Waltz). How do their core assumptions differ regarding the cause of conflict in international politics? (20 Marks)",
            "Q.3 Nuclear deterrence prevented a direct hot war between superpowers during the Cold War. Discuss the theory of Mutually Assured Destruction (MAD) and its applicability to the strategic stability of South Asia between Pakistan and India. (20 Marks)",
            "Q.4 \"Diplomacy is the management of international relations by negotiation.\" Trace the evolution of diplomacy from traditional bilateral diplomacy to modern multilateral and public diplomacy. (20 Marks)",
            "Q.5 Critically evaluate the foreign policy of the United States toward South Asia post-9/11. How has Washington's strategic partnership with India influenced the regional balance of power? (20 Marks)",
            "Q.6 The United Nations was established to save succeeding generations from the scourge of war. Critically assess the performance of the UN Security Council in conflict resolution and peacekeeping in the post-Cold War era. (20 Marks)",
            "Q.7 Discuss the concept of \"Imperialism\" and its economic theories, particularly Hobson and Lenin's critique of imperial capitalism. To what extent does neo-imperialism persist in the contemporary global economic order? (20 Marks)",
            "Q.8 What is the Non-Aligned Movement (NAM)? Evaluate its historical contribution during the Cold War bipolarity and analyze its relevance in today's multipolar world. (20 Marks)"
        ]
    },
    {
        "year": "2013",
        "title": "KPPSC PMS International Relations Past Paper 2013",
        "md_file": "KPPSC_PMS_International_Relations_Paper_2013.md",
        "pdf_file": "KPPSC_PMS_International_Relations_Paper_2013.pdf",
        "questions": [
            "Q.1 What is \"Power\" in International Relations? Differentiate between Hard Power, Soft Power, and Smart Power, and explain the tangible and intangible elements of National Power. (20 Marks)",
            "Q.2 \"Balance of Power is a central concept in Realism.\" Discuss the various mechanisms employed by states to maintain or restore the Balance of Power (alliances, armaments, buffer states, compensation). (20 Marks)",
            "Q.3 Analyze the causes and global repercussions of the First World War (1914–1918). How did the failure of the Treaty of Versailles contribute to the outbreak of World War II? (20 Marks)",
            "Q.4 Discuss the Cold War alliance systems (NATO vs. Warsaw Pact, SEATO, CENTO). How did these military alliances shape the security architecture of Europe and Asia? (20 Marks)",
            "Q.5 Critically examine the Palestine Problem from the Balfour Declaration (1917) to the post-Oslo developments. Why has a permanent peace settlement remained elusive? (20 Marks)",
            "Q.6 What are the primary sources of International Law? Discuss the limitations of International Law as an instrument for regulating state behavior in an anarchic international system. (20 Marks)",
            "Q.7 Examine the dynamics of Arms Control and Disarmament in the nuclear era. Evaluate the effectiveness and shortcomings of the Nuclear Non-Proliferation Treaty (NPT) and Comprehensive Test Ban Treaty (CTBT). (20 Marks)",
            "Q.8 Write short notes on any TWO of the following: (a) Decolonization and the Rise of the Third World; (b) The New International Economic Order (NIEO); (c) Peaceful Coexistence vs. Cold War Containment; (d) Geopolitical rivalry in the Indian Ocean. (20 Marks)"
        ]
    }
]

for p in past_papers_data:
    md_path = os.path.join(PAPERS_DIR, p["md_file"])
    pdf_path = os.path.join(PAPERS_DIR, p["pdf_file"])
    
    paper_text = f"""# KHYBER PAKHTUNKHWA PUBLIC SERVICE COMMISSION (KPPSC)
## COMPETITIVE EXAMINATION FOR THE POSTS OF PMS OFFICERS (BPS-17)
### SUBJECT: INTERNATIONAL RELATIONS (100 MARKS) — {p['year']}

> **Time Allowed**: Three (03) Hours | **Maximum Marks**: 100 Marks  
> **Instructions**: Attempt any **FIVE (05)** questions in all. All questions carry equal marks (20 Marks each). Answers must be analytical, objective, theoretically grounded, and supported by contemporary empirical evidence.

---

### QUESTIONS

"""
    for q in p["questions"]:
        paper_text += f"**{q}**\n\n---\n\n"
        
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(paper_text)
    compile_pdf(md_path, pdf_path, p["title"])

print("All IR past papers and syllabus generated successfully!")
