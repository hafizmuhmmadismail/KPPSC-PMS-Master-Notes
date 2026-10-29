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

syllabus_html = """<!DOCTYPE html>
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
    font-size: 13pt;
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
    text-transform: uppercase;
  }
  .header h3 {
    margin: 0 0 6px 0;
    font-size: 13pt;
    font-weight: normal;
  }
  .subject {
    font-size: 15pt;
    font-weight: bold;
    text-transform: uppercase;
    margin: 8px 0;
  }
  .marks-badge {
    display: inline-block;
    padding: 3px 12px;
    background: #eee;
    font-weight: bold;
    border: 1px solid #999;
    margin-top: 4px;
  }
  .section-title {
    font-size: 14pt;
    font-weight: bold;
    text-transform: uppercase;
    background: #0E2B50;
    color: #fff;
    padding: 6px 12px;
    margin-top: 25px;
    margin-bottom: 12px;
  }
  .topic-card {
    margin-bottom: 14px;
    padding-left: 6px;
  }
  .topic-head {
    font-weight: bold;
    font-size: 13pt;
    color: #0E2B50;
    margin-bottom: 4px;
  }
  .topic-body {
    margin-left: 20px;
    font-size: 12pt;
    color: #333;
  }
  ul {
    margin: 4px 0 8px 0;
    padding-left: 20px;
  }
  li {
    margin-bottom: 4px;
  }
  .page-break {
    page-break-before: always;
  }
  .footer {
    margin-top: 30px;
    text-align: center;
    font-size: 10pt;
    color: #555;
    border-top: 1px solid #ccc;
    padding-top: 6px;
  }
</style>
</head>
<body>
  <div class="header">
    <h2>Khyber Pakhtunkhwa Public Service Commission</h2>
    <h3>Provincial Management Service (PMS) Competitive Examination</h3>
    <div class="subject">Official Syllabus &bull; European History</div>
    <div class="marks-badge">Total Marks: 200 (Paper I: 100 Marks | Paper II: 100 Marks)</div>
  </div>

  <div class="section-title">PAPER I: (1789 – 1914) &bull; 100 MARKS</div>
  
  <div class="topic-card">
    <div class="topic-head">1. The French Revolution (1789–1799)</div>
    <div class="topic-body">
      <ul>
        <li>Background of the Revolution: Ancien Régime (Old Order), Social Estates, Economic Crisis.</li>
        <li>Intellectual origins: Contributions of Philosophers (Voltaire, Montesquieu, Rousseau, John Locke).</li>
        <li>From Estates-General to National Assembly (1789–1791) & Tennis Court Oath.</li>
        <li>Fall of Monarchy, the Legislative Assembly, and European Coalitions against Revolutionary France.</li>
        <li>The Reign of Terror, Committee of Public Safety (Robespierre), Thermidorian Reaction, and the Directory.</li>
        <li>Partitions of Poland (Second & Third) and European reactions.</li>
      </ul>
    </div>
  </div>

  <div class="topic-card">
    <div class="topic-head">2. Napoleon Bonaparte & The Napoleonic Empire (1799–1815)</div>
    <div class="topic-body">
      <ul>
        <li>Rise of Napoleon: Coup of 18 Brumaire, Napoleon as First Consul and Emperor.</li>
        <li>Internal Reforms: Code Napoléon (Civil Code), Concordat of 1801, Educational & Financial administration.</li>
        <li>Napoleon as a military strategist: Expansion across Europe and the Continental System (Berlin & Milan Decrees).</li>
        <li>Peninsular War ('Spanish Ulcer'), Invasion of Russia (1812), and British Naval Supremacy.</li>
        <li>Battle of Leipzig, Abdication, The Hundred Days, and the Battle of Waterloo (1815).</li>
      </ul>
    </div>
  </div>

  <div class="topic-card">
    <div class="topic-head">3. Vienna Settlement & The Concert of Europe (1814–1848)</div>
    <div class="topic-body">
      <ul>
        <li>Treaty of Chaumont, Treaties of Paris, and the Congress of Vienna (1814–1815).</li>
        <li>Guiding principles: Legitimacy, Balance of Power, Restoration, and Compensations.</li>
        <li>Holy Alliance and Quadruple/Quintuple Alliance; Working of the Congress System (Aix-la-Chapelle, Troppau, Laibach, Verona).</li>
        <li>The Metternich System: Defense of Conservatism vs. Rising forces of Nationalism, Liberalism, and Democracy.</li>
        <li>Revolutions of 1830 in France, Belgium, and Poland; Revolutions of 1848 across Continental Europe.</li>
      </ul>
    </div>
  </div>

  <div class="topic-card">
    <div class="topic-head">4. Industrial Revolution, Social Movements & England (1814–1848)</div>
    <div class="topic-body">
      <ul>
        <li>Impact of the Industrial Revolution on European economy, urbanization, and class structures.</li>
        <li>Rise of Socialism, Marxism, and working-class movements (Chartism in Britain).</li>
        <li>Constitutional developments in Britain: Reform Act of 1832, Repeal of Corn Laws, and Factory Legislation.</li>
      </ul>
    </div>
  </div>

  <div class="topic-card">
    <div class="topic-head">5. The Eastern Question (1820–1878)</div>
    <div class="topic-body">
      <ul>
        <li>Decline of the Ottoman Empire and Great Power rivalries in the Balkans and the Straits.</li>
        <li>Greek War of Independence (1821–1832) and the Treaty of Adrianople.</li>
        <li>The Crimean War (1853–1856): Causes, Course, Peace Treaty of Paris, and Diplomatic Consequences.</li>
        <li>Pan-Slavism, Russo-Turkish War (1877), Treaty of San Stefano, and the Congress and Treaty of Berlin (1878).</li>
      </ul>
    </div>
  </div>

  <div class="topic-card">
    <div class="topic-head">6. Unification of Italy (Risorgimento) (1815–1870)</div>
    <div class="topic-body">
      <ul>
        <li>Obstacles to Italian unity; Carbonari and early uprisings.</li>
        <li>The Triad of Unification: Mazzini (Soul), Cavour (Brain), Garibaldi (Sword), and Victor Emmanuel II.</li>
        <li>Piedmont-Sardinia's leadership, Alliance with Napoleon III (Plombières), War of 1859, Annexation of Central Italy, the Expedition of the Thousand, and Annexation of Rome (1870).</li>
      </ul>
    </div>
  </div>

  <div class="topic-card">
    <div class="topic-head">7. Unification of Germany & Bismarckian Era (1848–1890)</div>
    <div class="topic-body">
      <ul>
        <li>Prussian leadership, Frankfurt Parliament of 1848, and the economic foundation (Zollverein).</li>
        <li>Otto von Bismarck and the policy of "Blood and Iron".</li>
        <li>The Three Wars of Unification: Danish War (1864), Austro-Prussian Seven Weeks' War (1866), and Franco-Prussian War (1870–71). Proclamation of the German Empire at Versailles.</li>
        <li>Bismarck’s Foreign Policy (1871–1890): System of Alliances (Three Emperors' League, Dual Alliance, Triple Alliance, Reinsurance Treaty) aimed at the diplomatic isolation of France.</li>
      </ul>
    </div>
  </div>

  <div class="topic-card">
    <div class="topic-head">8. Armed Peace, Imperialism & Road to WWI (1890–1914)</div>
    <div class="topic-body">
      <ul>
        <li>Fall of Bismarck (1890) and Kaiser Wilhelm II’s *Weltpolitik* (World Policy).</li>
        <li>Franco-Russian Alliance (1894), Anglo-French Entente Cordiale (1904), Anglo-Russian Convention (1907) -> Division into Triple Alliance and Triple Entente.</li>
        <li>New Imperialism, Scramble for Africa, and the Anglo-German Naval Arms Race.</li>
        <li>Pre-war Crises: Moroccan Crises (1905, 1911), Bosnian Crisis (1908), Balkan Wars (1912–1913).</li>
        <li>The July Crisis of 1914: Sarajevo Assassination and the outbreak of World War I.</li>
      </ul>
    </div>
  </div>

  <div class="page-break"></div>

  <div class="section-title">PAPER II: (1914 – 1960) &bull; 100 MARKS</div>

  <div class="topic-card">
    <div class="topic-head">1. The First World War (1914–1918)</div>
    <div class="topic-body">
      <ul>
        <li>Nature of WWI as a Total War: Trench warfare, technological innovations, economic mobilization.</li>
        <li>Entry of the United States (1917) and Woodrow Wilson's Fourteen Points.</li>
        <li>Ottoman Empire in WWI: Middle Eastern front, McMahon-Hussein Correspondence, Sykes-Picot Agreement, and the Balfour Declaration (1917).</li>
        <li>Collapse of the Central Powers (Germany, Austria-Hungary, Ottoman Empire) and the Armistice of November 1918.</li>
      </ul>
    </div>
  </div>

  <div class="topic-card">
    <div class="topic-head">2. The Paris Peace Settlement (1919–1923)</div>
    <div class="topic-body">
      <ul>
        <li>The Paris Peace Conference, the 'Big Four' (Wilson, Lloyd George, Clemenceau, Orlando) and conflicting national interests.</li>
        <li>Treaty of Versailles (1919): War Guilt Clause (Article 231), Reparations, Disarmament, and Territorial clauses.</li>
        <li>Treaties of St. Germain (Austria), Trianon (Hungary), Neuilly (Bulgaria), Sèvres (Ottoman), and the Treaty of Lausanne (1923).</li>
      </ul>
    </div>
  </div>

  <div class="topic-card">
    <div class="topic-head">3. The League of Nations & Interwar Security (1919–1939)</div>
    <div class="topic-body">
      <ul>
        <li>Structure and Covenant of the League of Nations; Principle of Collective Security.</li>
        <li>Successes in the 1920s: Humanitarian work, minor boundary disputes.</li>
        <li>Search for Security: Treaty of Mutual Assistance, Geneva Protocol, Locarno Treaties (1925), Kellogg-Briand Pact (1928).</li>
        <li>Failures in the 1930s: Japanese invasion of Manchuria (1931), Italian invasion of Abyssinia (1935), German rearmament, and structural shortcomings.</li>
      </ul>
    </div>
  </div>

  <div class="topic-card">
    <div class="topic-head">4. Marxism & The Soviet Union (1917–1939)</div>
    <div class="topic-body">
      <ul>
        <li>Russian Revolutions of 1917: February Revolution (Fall of Tsardom) and October Bolshevik Revolution (Lenin & Trotsky).</li>
        <li>Russian Civil War (1918–1921), War Communism, and the New Economic Policy (NEP).</li>
        <li>Rise of Joseph Stalin vs. Leon Trotsky; "Socialism in One Country".</li>
        <li>Collectivization of Agriculture, Five-Year Plans, Industrialization, and the Great Purges (1936–1938).</li>
        <li>Soviet Foreign Policy between the wars: Comintern, Rapallo, Popular Fronts, and the Nazi-Soviet Non-Aggression Pact (1939).</li>
      </ul>
    </div>
  </div>

  <div class="topic-card">
    <div class="topic-head">5. Fascism in Italy & Nazism in Germany</div>
    <div class="topic-body">
      <ul>
        <li>Post-WWI crises, "Mutilated Victory", and the rise of Benito Mussolini and the Fascist Party (March on Rome 1922).</li>
        <li>The Corporate State, Lateran Accords (1929), and Italian aggressive foreign policy.</li>
        <li>Weimar Republic: Structural defects, Hyperinflation (1923), and the devastating impact of the Great Depression (1929).</li>
        <li>Rise of Adolf Hitler: Ideology of *Mein Kampf*, totalitarian consolidation (*Gleichschaltung*), Anti-Semitism, and rearmament.</li>
        <li>Comparison between Fascism and Nazism: Ideology, methods of social control, and economic mobilization.</li>
      </ul>
    </div>
  </div>

  <div class="topic-card">
    <div class="topic-head">6. Democracies in Crisis & Road to World War II (1919–1939)</div>
    <div class="topic-body">
      <ul>
        <li>France: Political instability under the Third Republic, Maginot Line psychology, and quest for security.</li>
        <li>Great Britain: Economic stagnation, Great Depression, Commonwealth evolution (Statute of Westminster 1931).</li>
        <li>Spanish Civil War (1936–1939) as an ideological proxy war.</li>
        <li>Policy of Appeasement: Neville Chamberlain, Rhineland remilitarization, Anschluss of Austria, and the Munich Agreement (1938).</li>
      </ul>
    </div>
  </div>

  <div class="topic-card">
    <div class="topic-head">7. The Second World War (1939–1945)</div>
    <div class="topic-body">
      <ul>
        <li>Invasion of Poland, Blitzkrieg in Western Europe, Battle of Britain, and Operation Barbarossa.</li>
        <li>Turning points: Battle of Stalingrad, El Alamein, Midway, and the Normandy Landings (D-Day).</li>
        <li>The Holocaust and Nazi occupation policies.</li>
        <li>Major Allied Conferences: Atlantic Charter (1941), Tehran (1943), Yalta (1945), and Potsdam (1945).</li>
      </ul>
    </div>
  </div>

  <div class="topic-card">
    <div class="topic-head">8. Post-War Settlement, Cold War & European Integration (1945–1960)</div>
    <div class="topic-body">
      <ul>
        <li>Creation and Charter of the United Nations Organization (UNO); comparison with the League.</li>
        <li>Emergence of the Bipolar World: Division of Germany and the Berlin Blockade (1948–49).</li>
        <li>Containment Doctrine: Truman Doctrine (1947), Marshall Plan (1947), and creation of NATO (1949).</li>
        <li>Sovietization of Eastern Europe, Warsaw Pact (1955), De-Stalinization, and the Hungarian Uprising (1956).</li>
        <li>Origins of European Integration: European Coal and Steel Community (ECSC 1951), Treaty of Rome (1957), and the Common Market (EEC).</li>
      </ul>
    </div>
  </div>

  <div class="footer">
    Khyber Pakhtunkhwa Public Service Commission (KPPSC) &bull; PMS Competitive Examination Syllabus &bull; European History
  </div>
</body>
</html>
"""

syllabus_pdf_path = r"H:\My Drive\PMS KPK 2022 notes\1. European History Paper 1 and Paper 2\KPPSC_PMS_European_History_Official_Syllabus.pdf"
render_html_to_pdf(syllabus_html, syllabus_pdf_path)

# Also write as clean Markdown
syllabus_md = """# KPPSC PMS European History Syllabus (Official Reference)

**Total Marks:** 200 Marks  
**Paper I (1789 – 1914):** 100 Marks  
**Paper II (1914 – 1960):** 100 Marks  
**Examination Format:** Descriptive / Subjective (Attempt 5 out of 8 questions, 20 marks each).

---

## PAPER I: (1789 – 1914) — 100 Marks

### 1. The French Revolution (1789–1799)
- **Background of the Revolution:** The *Ancien Régime* (Old Order), the Three Estates, financial crisis and bankruptcy under Louis XVI.
- **Intellectual Origins:** Enlightenment philosophers (Voltaire, Montesquieu, Jean-Jacques Rousseau, John Locke).
- **Phases of the Revolution:**
  - Estates-General meeting and formation of the National Assembly (1789).
  - Storming of the Bastille, Declaration of the Rights of Man and of the Citizen.
  - The Civil Constitution of the Clergy and Constitution of 1791.
  - Legislative Assembly, foreign coalitions, and the fall of the Monarchy.
  - National Convention, the Reign of Terror (Robespierre & Committee of Public Safety), and the Thermidorian Reaction.
  - The Directory (1795–1799).
- **Partitions of Poland:** Second (1793) and Third (1795) partitions and European diplomacy.

### 2. Napoleon Bonaparte & The Napoleonic Empire (1799–1815)
- **Rise to Power:** Coup of 18 Brumaire (1799), Consulate, and proclamation of the Empire (1804).
- **Internal Reforms:** *Code Napoléon* (Civil Code), financial reorganization (Bank of France), education reform, Concordat of 1801 with the Catholic Church, administrative centralization.
- **Foreign Policy & Military Campaigns:** Austerlitz, Jena-Auerstedt, Friedland, Treaties of Tilsit.
- **The Continental System:** Berlin and Milan Decrees; British Orders in Council; causes of economic and administrative failure.
- **Decline and Fall:** The Peninsular War ("Spanish Ulcer"), Moscow campaign (1812), War of the Sixth Coalition (Battle of Leipzig 1813), Abdication, The Hundred Days, and the Battle of Waterloo (1815).

### 3. Vienna Settlement & The Concert of Europe (1814–1848)
- **The Congress of Vienna (1814–1815):** The Big Four (Metternich, Castlereagh, Alexander I, Hardenberg) and Talleyrand.
- **Core Principles:** Legitimacy, Restoration, Balance of Power, and Territorial Compensation.
- **The Congress System:** Holy Alliance vs. Quadruple/Quintuple Alliance. Congresses of Aix-la-Chapelle (1818), Troppau (1820), Laibach (1821), and Verona (1822); causes of its eventual breakdown.
- **The Metternich Era:** Repression of Liberalism and Nationalism; Carlsbad Decrees (1819).
- **Revolutions in France and Europe:**
  - The July Revolution of 1830 in France (Louis Philippe and the July Monarchy); independence of Belgium.
  - The Revolutions of 1848 ("Spring of Nations"): February Revolution in Paris, uprisings in the Austrian Empire, German states, and Italian states; causes of their ultimate defeat.

### 4. Great Britain (1814–1848) & Social Movements
- **Post-Napoleonic Britain:** Economic distress, Peterloo Massacre (1819), Corn Laws.
- **Constitutional Progress:** Catholic Emancipation (1829), the Great Reform Act of 1832, Repeal of the Corn Laws (1846).
- **The Industrial Revolution:** Factory system, social transformation, rise of Chartism, and early utopian/Marxian socialism.

### 5. The Eastern Question (1820–1878)
- **Core Dilemma:** The decline of the Ottoman Empire ("Sick Man of Europe") and the strategic rivalry of Britain, Russia, France, and Austria.
- **Greek War of Independence (1821–1832):** Battle of Navarino (1827) and Treaty of London/Adrianople.
- **The Crimean War (1853–1856):** Holy Places controversy, Sinope, Siege of Sevastopol, and the Peace Treaty of Paris (1856) (neutralization of the Black Sea).
- **The Balkan Crisis (1875–1878):** Pan-Slavism, Russo-Turkish War (1877), Treaty of San Stefano, and the Congress and Treaty of Berlin (1878) under Bismarck's presidency.

### 6. Unification of Italy (Risorgimento) (1815–1870)
- **Pre-conditions & Obstacles:** Austrian hegemony, Papal opposition, regional fragmentation.
- **Key Architects:**
  - Giuseppe Mazzini & Young Italy (*The Prophet/Soul*).
  - Count Camillo Cavour & Kingdom of Sardinia-Piedmont (*The Brain/Diplomat*).
  - Giuseppe Garibaldi & Red Shirts (*The Sword*).
  - King Victor Emmanuel II (*The Constitutional Figurehead*).
- **Milestones:** Crimean War participation, Pact of Plombières with Napoleon III, War of 1859 (Villafranca), annexation of Parma/Modena/Tuscany, Garibaldi's conquest of Naples and Sicily (1860), acquisition of Venetia (1866) and Rome (1870).

### 7. Unification of Germany (1848–1871) & Bismarckian Era (1871–1890)
- **Prussian Dominance:** Zollverein (customs union), King Wilhelm I, and military reforms of Roon and Moltke.
- **Otto von Bismarck:** Realpolitik, "Blood and Iron" philosophy.
- **Three Wars of Unification:**
  - Danish War (1864) over Schleswig-Holstein.
  - Austro-Prussian War / Seven Weeks' War (1866): Battle of Sadowa/Königgrätz; Treaty of Prague; creation of North German Confederation.
  - Franco-Prussian War (1870–1871): Ems Telegram, Battle of Sedan; Siege of Paris; Treaty of Frankfurt (Loss of Alsace-Lorraine); Proclamation of the German Empire at Versailles.
- **Bismarck's Foreign Policy (1871–1890):** Maintaining the European status quo and isolating France:
  - Three Emperors' League (*Dreikaiserbund* 1873, 1881).
  - Dual Alliance (1879) and Triple Alliance (1882) with Austria and Italy.
  - Reinsurance Treaty with Russia (1887).

### 8. Armed Peace, Imperialism & Origins of World War I (1890–1914)
- **Post-Bismarckian Shift:** Accession of Kaiser Wilhelm II, "Dropping the Pilot" (1890), and aggressive *Weltpolitik*.
- **Counter-Alliances:** Franco-Russian Alliance (1894), Anglo-French *Entente Cordiale* (1904), Anglo-Russian Convention (1907) -> Formation of the Triple Entente.
- **New Imperialism & Naval Arms Race:** Scramble for Africa, Dreadnought competition between Britain and Germany.
- **Crises Leading to War:**
  - First (1905) and Second (1911) Moroccan Crises.
  - Bosnian Annexation Crisis (1908).
  - Balkan Wars (1912–1913).
  - The July Crisis (1914): Assassination of Archduke Franz Ferdinand at Sarajevo and the escalation into general war.

---

## PAPER II: (1914 – 1960) — 100 Marks

### 1. The First World War (1914–1918)
- Nature of the War as "Total War": War of attrition, trench warfare, blockade, propaganda, and home front mobilization.
- US entry into the war (1917) and Woodrow Wilson's Fourteen Points.
- The Ottoman Empire: Gallipoli campaign, Middle Eastern partition agreements (Sykes-Picot, Balfour Declaration 1917).
- Collapse of the German, Austro-Hungarian, Russian, and Ottoman Empires.

### 2. The Paris Peace Settlement (1919–1923)
- The Big Four: Woodrow Wilson (USA), David Lloyd George (UK), Georges Clemenceau (France), Vittorio Orlando (Italy).
- Treaty of Versailles (1919): War Guilt (Art. 231), reparations, disarmament, territorial adjustments, Polish Corridor, Rhineland demilitarization.
- Subordinate Treaties: St. Germain, Trianon, Neuilly, Sèvres, and the revised Treaty of Lausanne (1923) under Mustafa Kemal Atatürk.

### 3. The League of Nations & Interwar Security (1919–1939)
- Principles, structure (Council, Assembly, Secretariat), and the Covenant.
- Successes in the 1920s (Aaland Islands, Upper Silesia, humanitarian work).
- Search for Collective Security: Locarno Treaties (1925), Kellogg-Briand Pact (1928), Geneva Disarmament Conference (1932–34).
- Failures in the 1930s: Japanese aggression in Manchuria (1931), Italian invasion of Abyssinia (1935), Spanish Civil War, and withdrawal of major powers.

### 4. Marxism & The Soviet Union (1917–1939)
- Russian Revolutions of 1917: February Revolution (fall of the Romanovs) and October Revolution (Bolshevik seizure of power).
- Lenin's leadership, War Communism, and the New Economic Policy (NEP 1921).
- The Stalin-Trotsky power struggle; "Socialism in One Country".
- Stalin's Five-Year Plans, collectivization of agriculture, rapid industrialization, and the Great Purges (1936–1938).
- Soviet foreign policy: The Comintern, diplomatic normalization, and the Molotov-Ribbentrop Pact (1939).

### 5. Fascism in Italy & Nazism in Germany
- **Fascist Italy:** Crisis of liberal Italy, Mussolini’s March on Rome (1922), the Corporate State, Lateran Treaty (1929), and invasion of Ethiopia (1935).
- **Nazi Germany:** Flaws of the Weimar Republic, hyperinflation, impact of the Great Depression, Hitler's rise to Chancellorship (1933), *Enabling Act*, totalitarian police state, Nuremberg Laws, anti-Semitism, and economic autarky.

### 6. Democracies in Crisis & The Road to World War II
- **France:** Ministerial instability under the Third Republic, economic stagnation, Popular Front government (1936), defensive mindset.
- **Britain:** Economic hardship, General Strike (1926), Great Depression, Imperial Preference.
- **The Policy of Appeasement:** Reasons for British/French appeasement under Neville Chamberlain:
  - German remilitarization of the Rhineland (1936).
  - *Anschluss* with Austria (March 1938).
  - Sudetenland Crisis and Munich Conference (September 1938).
  - Occupation of Czechoslovakia (March 1939) and the Polish Guarantee.

### 7. The Second World War (1939–1945)
- Course of the War: Blitzkrieg, Fall of France, Battle of Britain, Operation Barbarossa (Invasion of Russia), Pearl Harbor.
- Turning Points (1942–43): Stalingrad, Second Battle of El Alamein, Midway, Italian surrender.
- D-Day (Normandy landings June 1944) and the downfall of Nazi Germany.
- Wartime Diplomacy & Conferences: Atlantic Charter, Casablanca, Tehran, Yalta, and Potsdam conferences.

### 8. Post-War Settlement, The Cold War & European Integration (1945–1960)
- **The United Nations Organization (UNO):** Origins, San Francisco Conference (1945), structural comparison with the League.
- **Origins of the Cold War:** Division of Germany, Iron Curtain speech, Truman Doctrine (1947), Marshall Plan (1947), Berlin Blockade and Airlift (1948–49).
- **Bipolar Military Alliances:** Creation of NATO (1949) and the Warsaw Pact (1955).
- **Eastern Europe under Soviet Hegemony:** Satellite states, Comecon, Cominform, De-Stalinization under Khrushchev, and the Hungarian Uprising (1956).
- **European Integration:** Schuman Plan, European Coal and Steel Community (ECSC 1951), Treaty of Rome (1957), and founding of the European Economic Community (Common Market).
"""

with open(r"H:\My Drive\PMS KPK 2022 notes\1. European History Paper 1 and Paper 2\KPPSC_PMS_European_History_Official_Syllabus.md", "w", encoding="utf-8") as f:
    f.write(syllabus_md)

print("Syllabus files generated successfully.")
