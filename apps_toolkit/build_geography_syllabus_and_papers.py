import os
import subprocess

BASE_DIR = r"H:\My Drive\PMS KPK 2022 notes\3. Geography Paper 1 and Paper 2"
PAPERS_DIR = os.path.join(BASE_DIR, "Past Papers")
os.makedirs(PAPERS_DIR, exist_ok=True)

EDGE_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

def compile_pdf(md_path, pdf_path, title):
    with open(md_path, "r", encoding="utf-8") as f:
        md_text = f.read()

    # Simple Markdown to HTML converter
    lines = md_text.split("\n")
    html_body = []
    in_table = False
    in_ul = False
    in_ol = False
    in_blockquote = False

    for line in lines:
        sline = line.strip()
        
        # Blockquote
        if sline.startswith(">"):
            if not in_blockquote:
                html_body.append("<blockquote>")
                in_blockquote = True
            bq_content = sline.lstrip("> ").replace("**", "<strong>").replace("*", "<em>")
            html_body.append(f"<p>{bq_content}</p>")
            continue
        elif in_blockquote:
            html_body.append("</blockquote>")
            in_blockquote = False

        # Table
        if sline.startswith("|") and sline.endswith("|"):
            cells = [c.strip() for c in sline.split("|")[1:-1]]
            if all(set(c) <= set("-: ") for c in cells):
                continue  # divider
            if not in_table:
                html_body.append("<table><tbody>")
                in_table = True
                html_body.append("<tr>" + "".join(f"<th>{c}</th>" for c in cells) + "</tr>")
            else:
                html_body.append("<tr>" + "".join(f"<td>{c}</td>" for c in cells) + "</tr>")
            continue
        elif in_table:
            html_body.append("</tbody></table>")
            in_table = False

        # Unordered list
        if sline.startswith("- ") or sline.startswith("* "):
            if not in_ul:
                html_body.append("<ul>")
                in_ul = True
            item_text = sline[2:].replace("**", "<strong>").replace("*", "<em>")
            html_body.append(f"<li>{item_text}</li>")
            continue
        elif in_ul:
            html_body.append("</ul>")
            in_ul = False

        # Headings
        if sline.startswith("# "):
            html_body.append(f"<h1>{sline[2:]}</h1>")
        elif sline.startswith("## "):
            html_body.append(f"<h2>{sline[3:]}</h2>")
        elif sline.startswith("### "):
            html_body.append(f"<h3>{sline[4:]}</h3>")
        elif sline.startswith("#### "):
            html_body.append(f"<h4>{sline[5:]}</h4>")
        elif sline == "---":
            html_body.append("<hr/>")
        elif sline == "":
            html_body.append("<br/>")
        else:
            p_text = sline.replace("**", "<strong>").replace("*", "<em>")
            html_body.append(f"<p>{p_text}</p>")

    if in_table:
        html_body.append("</tbody></table>")
    if in_ul:
        html_body.append("</ul>")
    if in_blockquote:
        html_body.append("</blockquote>")

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
    h4 {{
        color: #333333;
        font-size: 10pt;
        margin-top: 8px;
        margin-bottom: 2px;
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
        border: 1px solid #0E2B50;
    }}
    td {{
        padding: 4px 8px;
        border: 1px solid #D0D7DE;
        vertical-align: top;
    }}
    tr:nth-child(even) td {{
        background-color: #F8FAFC;
    }}
    blockquote {{
        border-left: 4px solid #0E2B50;
        background-color: #F0F4F8;
        padding: 6px 12px;
        margin: 8px 0;
        font-size: 9.5pt;
    }}
    blockquote p {{
        margin: 3px 0;
    }}
    ul, ol {{
        margin: 4px 0 8px 18px;
        padding: 0;
    }}
    li {{
        margin-bottom: 3px;
    }}
    hr {{
        border: none;
        border-top: 1px solid #D0D7DE;
        margin: 12px 0;
    }}
    p {{
        margin: 4px 0;
    }}
    .footer {{
        margin-top: 20px;
        font-size: 8pt;
        color: #666;
        text-align: center;
        border-top: 1px solid #E1E4E8;
        padding-top: 6px;
    }}
</style>
</head>
<body>
{"".join(html_body)}
<div class="footer">KPPSC PMS Examination Suite — Master Academic Repository</div>
</body>
</html>
"""

    temp_html = pdf_path.replace(".pdf", "_temp.html")
    with open(temp_html, "w", encoding="utf-8") as f:
        f.write(html_content)

    cmd = [
        EDGE_PATH,
        "--headless",
        "--disable-gpu",
        "--run-all-compositor-stages-before-draw",
        f"--print-to-pdf={pdf_path}",
        temp_html
    ]
    subprocess.run(cmd, check=True)
    if os.path.exists(temp_html):
        os.remove(temp_html)
    print(f"Generated PDF: {pdf_path}")

# ==================== 1. SYLLABUS ====================
SYLLABUS_MD = os.path.join(BASE_DIR, "KPPSC_PMS_Geography_Official_Syllabus.md")
SYLLABUS_PDF = os.path.join(BASE_DIR, "KPPSC_PMS_Geography_Official_Syllabus.pdf")

syllabus_text = """# Khyber Pakhtunkhwa Public Service Commission (KPPSC)
## Syllabus for Provincial Management Service (PMS) Competitive Examination
### GEOGRAPHY — OPTIONAL SUBJECT (TOTAL MARKS: 200)

> **Examination Structure**: Geography in KPPSC PMS is a full **200-mark optional subject** divided into **TWO distinct papers**:
> - **Paper-I: Physical and General Geography (100 Marks | 3 Hours)**
> - **Paper-II: Human, Economic and Regional Geography (100 Marks | 3 Hours)**
> **Examination Pattern for Both Papers**: Attempt **ANY FIVE (5)** questions out of Eight (8). Each question carries **20 Marks**. Time allowed: **3 Hours** per paper.

---

## 🌍 PAPER-I: PHYSICAL AND GENERAL GEOGRAPHY (100 MARKS)

### Section 1: Elements of Weather and Climate (Climatology)
1. **Insolation and Atmospheric Temperature**:
   - Mechanism of solar insolation; Factors affecting solar radiation distribution (latitude, angle of incidence, duration of daylight, atmospheric absorption).
   - Global radiation balance and terrestrial heat budget; Heat deficit and surplus zones; Latitudinal heat transfer.
   - Horizontal and vertical distribution of temperature; Temperature inversion (types, causes, and environmental significance).
2. **Atmospheric Pressure and Wind Systems**:
   - Planetary atmospheric pressure belts (Equatorial Low, Subtropical Highs, Subpolar Lows, Polar Highs).
   - Mechanics of atmospheric circulation; Pressure gradient force, Coriolis force, frictional force, and geostrophic balance.
   - Primary planetary wind systems (Trade Winds, Westerlies, Polar Easterlies); Secondary circulations (Monsoons, Land/Sea breezes, Anabatic/Katabatic winds).
3. **Humidity, Condensation, and Precipitation**:
   - Atmospheric moisture: Absolute humidity, relative humidity, specific humidity, and dew point.
   - Forms of condensation (clouds, fog, dew, frost); Cloud classification (high, middle, low, clouds of vertical development).
   - Mechanisms of precipitation: Convectional, Orographic, and Cyclonic/Frontal precipitation; Global spatial distribution of precipitation.
4. **Air Masses, Fronts, and Atmospheric Disturbances**:
   - Air masses: Definition, source regions, thermal and moisture classifications (cP, cT, mP, mT), and modification mechanisms.
   - Frontogenesis and frontal types: Cold fronts, warm fronts, occluded fronts, and stationary fronts.
   - Atmospheric disturbances: Life cycle and structure of mid-latitude wave cyclones (Polar Front Theory); Tropical cyclones (Hurricanes/Typhoons: origins, environmental conditions, eye wall mechanics, and track patterns); Anticyclones and associated weather.
5. **Climatic Classifications and World Climatic Regions**:
   - Wladimir Köppen’s climatic classification system: First-letter designations (A, B, C, D, E), second and third-letter criteria (precipitation seasonality and temperature characteristics).
   - Detailed comparative analysis of major climatic types: Tropical Rainforest (Af), Tropical Savanna (Aw), Mediterranean (Cs), Marine West Coast (Cfb), Steppe (BS), and Desert (BW) climates.

### Section 2: Landforms and Geomorphology
1. **Internal Structure of the Earth and Seismology**:
   - Earth's interior layering: Crust (Sial and Sima), Mohorovičić discontinuity, Mantle (Asthenosphere and Mesosphere), Gutenberg discontinuity, Outer Core (liquid iron-nickel), and Inner Core (solid iron-nickel).
   - Seismic wave behavior: Primary (P) waves, Secondary (S) waves, shadow zones, and seismic evidence for core composition and mantle rheology.
2. **Plate Tectonics and Tectonic Landforms**:
   - Theory of continental drift (Wegener) and seafloor spreading (Hess); Plate Tectonics Paradigm.
   - Three major types of plate boundaries and collisions: Divergent (rift valleys, mid-ocean ridges), Convergent (oceanic-continental, oceanic-oceanic, continental-continental orogenesis), and Transform faults.
   - Earthquakes and Volcanicity: Mechanisms, epicenters, Wadati-Benioff zones; Volcanic landforms (intrusive: batholiths, laccoliths, sills, dykes; extrusive: shield volcanoes, composite cones, calderas); Tectonic and seismic zones of Pakistan (Chaman Fault, Main Boundary Thrust).
3. **Rocks and the Rock Cycle**:
   - Definition and classification of rocks: Igneous (intrusive vs. extrusive, felsic vs. mafic mineral composition), Sedimentary (mechanically formed, chemically precipitated, organically formed), and Metamorphic (foliated vs. non-foliated, contact vs. regional metamorphism).
   - The geological rock cycle and petrographic transformations.
4. **Geomorphic Processes: Endogenic and Exogenic Forces**:
   - Weathering: Physical/mechanical weathering (frost wedging, thermal expansion/exfoliation, salt wedging) and Chemical weathering (carbonation, hydration, oxidation, hydrolysis); Biological weathering.
   - Mass Wasting: Slope stability factors; Slow flowage (soil creep, solifluction), Rapid flowage (earthflow, mudflow, debris avalanche), Landslides (rockfall, rockslide, slumping).
5. **Fluvial, Glacial, Aeolian, Karst, and Coastal Systems**:
   - Fluvial Geomorphology: Stages of a river (youth, maturity, old age); Erosional features (gorges, waterfalls, potholes, river capture); Depositional features (floodplains, meanders, oxbow lakes, braided streams, natural levees, deltas).
   - Glacial Landforms: Valley/Alpine glaciers vs. Continental ice sheets; Erosional features (cirques, arêtes, horn peaks, U-shaped valleys, hanging valleys); Depositional features (moraines: terminal, lateral, medial, ground; drumlins, eskers, kames).
   - Groundwater and Karst Topography: Prerequisites for Karst development; Surface features (sinkholes/dolines, uvalas, poljes, lapies); Underground features (caves, stalactites, stalagmites, columns).
   - Aeolian Landforms: Wind erosion (deflation, abrasion, yardangs, zeugen); Wind deposition (dunes: barchans, seifs, transverse, longitudinal; loess deposits).
   - Coastal Geomorphology: Wave action, longshore drift; Erosional coastal features (sea cliffs, wave-cut platforms, sea arches, stacks); Depositional coastal features (beaches, spits, bars, tombolos).
6. **Cycle of Erosion**:
   - W.M. Davis's Geographical Cycle (Youth, Maturity, Old Age / Peneplain); Walther Penck's Morphological Analysis (Aufsteigende, Gleichförmige, Absteigende Entwicklung).

### Section 3: Oceans and Oceanography
1. **Morphology of the Ocean Floor**:
   - Major physiographic divisions: Continental shelf, continental slope, continental rise, deep ocean abyssal plains, oceanic trenches, mid-ocean ridges, and seamounts/guyots.
2. **Physical and Chemical Properties of Seawater**:
   - Chemical composition: Salinity (average 35‰), major constituent salts (NaCl, MgCl2, etc.); Factors determining spatial and vertical salinity variations (evaporation, precipitation, river inflow, ice melting).
   - Temperature distribution of ocean waters: Horizontal latitudinal variations, vertical thermal structure (thermocline layer).
3. **Ocean Water Movements: Waves, Currents, and Tides**:
   - Oceanic Circulation: Major surface current gyres of the Atlantic Ocean (Gulf Stream, North Atlantic Drift, Canaries Current, Benguela), Pacific Ocean (Kuroshio, California, Humboldt/Peru Current), and Indian Ocean (Monsoon drift reversal).
   - Tides: Equilibrium Theory (Newton) and Dynamic Theory (Laplace); Spring tides, neap tides, semi-diurnal, diurnal, and mixed tides; Tidal bores and coastal amplification.

### Section 4: Study of Maps and Cartography
1. **Map Basics and Scales**:
   - Definition and essentials of a map (title, orientation, scale, legend, grid/projection).
   - Map scales: Statement of scale, Graphic/linear scale, Representative Fraction (RF: definition, mathematical calculation, merits, and demerits).
2. **Types of Maps**:
   - Topographical maps (contour lines, relief representation, cultural features), Cadastral maps, Thematic maps.
   - Synoptic Weather Maps: Construction principles, international meteorological symbols (isobars, fronts, precipitation, cloud cover), interpretation, and modern weather forecasting methods and constraints.
   - Remote Sensing and Aerial Photography: Principles, stereoscopic viewing, photogrammetry, and GIS integration.
3. **Map Projections**:
   - Definition, necessity of projections, and spherical distortion (shape, area, distance, direction).
   - Classification: Cylindrical, Conical, Zenithal/Azimuthal, and Conventional projections.
   - Properties, merits, and limitations: Mercator’s Projection (rhumb line conformity vs. high-latitude areal exaggeration), Conical Projection with one/two standard parallels, Bonne’s Projection, and Gnomonic Projection.

---

## 🏙️ PAPER-II: HUMAN, ECONOMIC AND REGIONAL GEOGRAPHY (100 MARKS)

### Section 1: Human Geography
1. **Man and Habitat: Philosophical Paradigms**:
   - Environmental Determinism: Friedrich Ratzel, Ellen Churchill Semple, Ellsworth Huntington.
   - Possibilism: Paul Vidal de la Blache, Lucien Febvre.
   - Neo-Determinism / Stop-and-Go Determinism: Griffith Taylor; Probabilism and Cultural Determinism.
2. **World Population Dynamics**:
   - Spatial distribution and density of world population; Physical, cultural, and socio-economic controlling factors; Major ecumene and non-ecumene zones.
   - Population Growth: Historical trajectory, J-shaped exponential curve, S-shaped logistic curve; Demographic Transition Model (5 stages); Malthusian Theory of Population vs. Neo-Malthusianism; Optimum Population Theory.
   - Demographic Pressures in Pakistan: Population explosion, high dependency ratio, youth bulge, pressure on arable land, water resources, and urban infrastructure.
3. **Migration and Spatial Mobility**:
   - Types of migration: Internal (rural-to-urban, seasonal) and International migration.
   - Ravenstein's Laws of Migration; Lee's Push-Pull Model; Socio-economic and demographic consequences for sending and receiving regions (remittances, brain drain, urban slums).
4. **Human Settlements and Urbanization**:
   - Rural Settlements: Site and situation factors, nucleated vs. dispersed settlements, pioneer settlement patterns.
   - Urban Settlements: Definition, hierarchy, urban functions; Christaller’s Central Place Theory (k=3 marketing, k=4 transport, k=7 administrative principles; threshold and range of a good).
   - Urbanization: Global urbanization trends; Problems of rapid urbanization in developing countries (slums, katchi abadis, environmental degradation, water and sanitation deficits); Bridging the "green agenda" (ecological protection, open spaces) and "brown agenda" (safe water, sanitation, waste management) in sustainable urban planning.
5. **Society, Culture, and Geography of Religions**:
   - Spatial distribution of major human races and language families.
   - Geography of World Religions: Spatial distribution, historical diffusion, and geopolitical imprint of Islam and Christianity.

### Section 2: Economic Geography
1. **Classification of Economic Activities**:
   - Primary activities (agriculture, forestry, fishing, mining); Secondary activities (manufacturing, processing, construction); Tertiary activities (services, transport, trade); Quaternary (knowledge/information) and Quinary (high-level executive decision-making) sectors.
2. **Agricultural Systems and Global Commodities**:
   - Classification of agricultural systems: Subsistence vs. commercial, intensive vs. extensive agriculture, shifting cultivation, plantation agriculture, mechanized cereal farming, and mixed farming (livestock and crop integration).
   - Major cereal and cash crops: Wheat, Rice, Cotton, and Sugarcane (geographical requirements, soil and climate, global production, spatial distribution, and international trade).
   - Forest Resources: Tropical rainforests, tropical savanna woodlands, temperate coniferous forests; Economic benefits, deforestation, and environmental conservation.
   - Marine and Freshwater Fisheries: Major commercial fishing grounds of the world (North-West Pacific/Japan, North-East Atlantic/Dogger Bank, North-West Atlantic/Grand Banks, South-East Pacific/Peru); Factors favoring commercial fishing.
3. **Power Resources and Energy Geopolitics**:
   - Fossil fuels: Coal (types: anthracite, bituminous, lignite; major global coalfields), Petroleum / Crude Oil (formation, major oil-producing regions: Middle East, North America, Russia, offshore fields; OPEC and global trade), Natural Gas.
   - Alternative and Renewable Energy: Hydroelectric power, nuclear, solar, wind, and geothermal energy; Energy security and geopolitics of energy pipelines.
4. **Industrial Location and Major Manufacturing Regions**:
   - Theories of Industrial Location: Alfred Weber’s Least-Cost Theory (Material Index, locational triangle, labor cost, agglomeration economies); August Lösch’s Maximum Profit Location.
   - Major World Industries: Iron and Steel Industry (locational shift from raw material to market/coastal sites; US Rust Belt, European Ruhr, Chinese industrial nodes); Textile Industry (natural vs. synthetic fibers; historical shifts); Chemical and Petrochemical Industries.
5. **Geography of Development and Global Economic Relations**:
   - Indicators of Development: Economic (GDP, GNP, per capita income), socio-demographic (Human Development Index - HDI, Physical Quality of Life Index - PQLI, Gini Coefficient).
   - Core-Periphery Model and Dependency Theory; Politico-economic relations between More Developed Countries (MDCs) and Less Developed Countries (LDCs); Regional economic blocs (EU, NAFTA/USMCA, ASEAN).

### Section 3: Regional Geography (South Asia with Special Focus on Pakistan)
1. **Physiography and Drainage Base of Pakistan**:
   - Physical divisions: Northern and Western Mountain Ranges (Karakoram, Himalayas, Hindu Kush, Safed Koh, Sulaiman, Kirthar); Potohar Plateau and Salt Range; Balochistan Plateau; Indus River Plain (Upper Indus Plain, Lower Indus Plain, Deltaic region).
   - Drainage System: Indus River System and its eastern (Jhelum, Chenab, Ravi, Sutlej) and western tributaries (Kabul, Swat, Kurram, Gomal).
2. **Climatic Zones and Water Resources**:
   - Climatic divisions of Pakistan (Highland, Lowland, Coastal, Arid); Monsoon circulation and Western Depressions.
   - Water Resources and Irrigation Network: Canal irrigation system (inundation, perennial canals), dams and barrages; Indus Basin Treaty and hydrological challenges; Waterlogging, salinity, and climate change impacts.
3. **Agricultural Sector of Pakistan**:
   - Structural characteristics of Pakistan's agro-economy; Kharif and Rabi cropping seasons; Major food crops (wheat, rice) and cash crops (cotton, sugarcane).
   - Challenges facing agriculture: Land fragmentation, water scarcity, low per-acre yield, archaic practices, post-harvest losses, and climate volatility.
4. **Industrial Base and Mineral Wealth**:
   - Major industries of Pakistan: Textiles, sugar, cement, fertilizer, steel, and chemical industries (locational factors, raw materials, domestic and export role).
   - Mineral wealth: Gas fields (Sui), coal deposits (Thar coalfield), copper-gold reserves (Reko Diq, Saindak), rock salt, and limestone.
5. **Foreign Trade, Transport, and Geoeconomics**:
   - Foreign trade patterns: Structural composition of exports (textiles, rice, leather, surgical instruments) and imports (petroleum, machinery, edible oil, chemicals); Persistent trade deficit causes and export expansion strategies.
   - Transport infrastructure: National Highway Authority (NHA) network, Pakistan Railways, major seaports (Karachi Port, Port Qasim, Gwadar deep-sea port).
   - Geoeconomic Paradigm & CPEC: Shift from geopolitics to geoeconomics; China-Pakistan Economic Corridor (CPEC) as the flagship corridor of the Belt and Road Initiative (BRI); Gwadar port connectivity, Special Economic Zones (SEZs), and regional transit potential for Central Asia.
"""

with open(SYLLABUS_MD, "w", encoding="utf-8") as f:
    f.write(syllabus_text)
compile_pdf(SYLLABUS_MD, SYLLABUS_PDF, "KPPSC PMS Geography Official Syllabus")

# ==================== 2. PAST PAPERS COMPILATION ====================
P1_ALL_MD = os.path.join(PAPERS_DIR, "KPPSC_PMS_Geography_Paper_1_Past_Papers_2008_2025.md")
P1_ALL_PDF = os.path.join(PAPERS_DIR, "KPPSC_PMS_Geography_Paper_1_Past_Papers_2008_2025.pdf")

p1_all_text = """# Khyber Pakhtunkhwa Public Service Commission (KPPSC)
## Provincial Management Service (PMS) Competitive Examination
### GEOGRAPHY — PAPER-I (PHYSICAL AND GENERAL GEOGRAPHY)
#### Compendium of Solved Past Papers (2008 – 2025)

> **Time Allowed**: 3 Hours | **Maximum Marks**: 100 Marks  
> **Instructions**: Attempt **ANY FIVE (5)** questions. All questions carry equal marks (**20 Marks** each). Support your answers with relevant geographical sketches, diagrams, and maps where appropriate.

---

### 📝 KPPSC PMS GEOGRAPHY PAPER-I — 2025
* **Q.No.1**: Describe the internal structure of the Earth and explain how seismic wave behavior provides evidence for the layering of the Earth's interior. Evaluate the scientific reliability of seismic methods in understanding deep Earth processes. *(20 Marks)*
* **Q.No.2**: Explain the origin of ocean tides and critically examine the major theories proposed to explain tidal behavior. How do local coastal configurations influence tidal patterns? *(20 Marks)*
* **Q.No.3**: Discuss the major forms of precipitation and the atmospheric mechanisms that produce them. Support your answer with suitable regional and local examples illustrating each type. *(20 Marks)*
* **Q.No.4**: Analyze the formation and development of tropical weather systems with reference to hurricanes. Describe the weather patterns associated with hurricanes and the environmental factors influencing their intensity. *(20 Marks)*
* **Q.No.5**: Explain the basis of the First Letter Classification in the Köppen Climate Classification System. Compare and contrast the Af (tropical rainforest) and Aw (tropical savanna) climates in terms of distribution, temperature-precipitation regimes, and ecological characteristics. *(20 Marks)*
* **Q.No.6**: Discuss the environmental conditions necessary for the development of Karst Landscapes. Describe the major Karst Landforms and evaluate the processes responsible for their evolution. *(20 Marks)*
* **Q.No.7**: Explain the principles behind the construction and interpretation of weather maps. How are modern weather forecasts generated, and what are the limitations of current forecasting techniques? *(20 Marks)*
* **Q.No.8**: Describe the factors controlling salinity of sea water. Critically analyze the spatial and vertical variations in salinity across the world's oceans and explain their influence on ocean circulation and marine life. *(20 Marks)*

---

### 📝 KPPSC PMS GEOGRAPHY PAPER-I — 2022
* **Q.No.1**: Discuss various processes of physical and chemical weathering with support from sketches/diagrams. Explain their geomorphic significance. *(20 Marks)*
* **Q.No.2**: Provide a detailed account of sedimentary rocks and their most workable classification based on origin and composition. *(20 Marks)*
* **Q.No.3**: Discuss the distribution of pressure belts over the globe and the resulting planetary wind circulation systems. *(20 Marks)*
* **Q.No.4**: Air masses and fronts: Give a comprehensive classification and elaborate their weather characteristics with sketches. *(20 Marks)*
* **Q.No.5**: Describe the composition and vertical structure of the atmosphere. Explain the temperature profile of each layer. *(20 Marks)*
* **Q.No.6**: Discuss valley glaciers and the resultant erosional and depositional landforms in detail. *(20 Marks)*
* **Q.No.7**: Analyze the major causes of salinity variation in the oceans of the world. *(20 Marks)*
* **Q.No.8**: Write comprehensive short notes on any TWO of the following: *(10 + 10 = 20 Marks)*  
  (a) Maps and types of maps  
  (b) Merits and demerits of cylindrical projections  
  (c) Map projections and their choice in cartography  
  (d) Ocean currents of the Atlantic Ocean

---

### 📝 KPPSC PMS GEOGRAPHY PAPER-I — 2018
* **Q.No.1**: Discuss the endogenic and exogenic geomorphic processes in detail. How do they shape the Earth's relief? *(20 Marks)*
* **Q.No.2**: What are the major erosional and depositional landforms produced by rivers? Illustrate their evolution from youth to old age. *(20 Marks)*
* **Q.No.3**: Discuss the main types of rocks along with examples and describe the geological rock cycle. *(20 Marks)*
* **Q.No.4**: Discuss the horizontal distribution of temperature and factors affecting the temperature distribution across the globe. *(20 Marks)*
* **Q.No.5**: Describe the different layers and composition of the atmosphere with reference to lapse rate and pause zones. *(20 Marks)*
* **Q.No.6**: Explain with the help of maps the major ocean currents of the Indian, Pacific, and Atlantic Oceans. *(20 Marks)*
* **Q.No.7**: Define map and discuss its types along with their practical uses in modern geographical studies. *(20 Marks)*
* **Q.No.8**: Write short notes on any TWO of the following: *(10 + 10 = 20 Marks)*  
  (a) Glacial landforms  
  (b) Mercator’s Projection  
  (c) Ocean tides  
  (d) Climate change and coping mechanisms

---

### 📝 KPPSC PMS GEOGRAPHY PAPER-I — 2016
* **Q.No.1**: "The heat energy entering the atmosphere and its part reaching the earth surface is balanced by an equal amount of heat going out into outer space." Discuss the global heat balance. *(20 Marks)*
* **Q.No.2**: What are the distinguishing characteristics of an air mass? How does a source region influence the type of an air mass that forms over it? Explain the life cycle of a wave cyclone. *(20 Marks)*
* **Q.No.3**: What are the different layers of the Earth? Relate the three types of plate collisions associated with seismicity, volcanism, and mountain building. *(20 Marks)*
* **Q.No.4**: Describe the nature of groundwater and define the elements of groundwater environment. Review the processes and features associated with Karst topography. *(20 Marks)*
* **Q.No.5**: Explain the chemical composition of seawater and the basic physical structure of the ocean. *(20 Marks)*
* **Q.No.6**: "The coastal system is the scene of complex tidal fluctuation, winds, waves, and ocean currents." Discuss coastal processes and coastal landforms. *(20 Marks)*
* **Q.No.7**: Define cartography and mapping basics: Explain map scale (Representative Fraction) and the principles of map projections. *(20 Marks)*
* **Q.No.8**: Write comprehensive notes on any TWO of the following: *(10 + 10 = 20 Marks)*  
  (a) Planetary wind system  
  (b) The Köppen climatic classification  
  (c) Wind deposition / Aeolian landforms  
  (d) Mercator’s projection

---

### 📝 KPPSC PMS GEOGRAPHY PAPER-I — 2013
* **Q.No.1**: Write an essay on the Radiation Balance of the Earth and Atmosphere. *(20 Marks)*
* **Q.No.2**: What are the causes of Atmospheric Circulation? Explain the large-scale wind systems of the globe. *(20 Marks)*
* **Q.No.3**: Define rocks and minerals. Discuss the circumstances under which sedimentary rocks are formed and classified. *(20 Marks)*
* **Q.No.4**: Discuss the factors promoting mass movements in weathered materials. Explain the various types of mass movements. *(20 Marks)*
* **Q.No.5**: Describe the currents of the Pacific Ocean and their climatic significance. *(20 Marks)*
* **Q.No.6**: Write a comprehensive note on the salinity of ocean water and factors governing its variations. *(20 Marks)*
* **Q.No.7**: What are Map Projections? Give the merits and demerits of Cylindrical Projections. *(20 Marks)*

---

### 📝 KPPSC PMS GEOGRAPHY PAPER-I — 2010
* **Q.No.1**: Write an essay on the influence of latitude on the annual cycle of insolation. *(20 Marks)*
* **Q.No.2**: Define Cyclone and Anticyclone with particular reference to air movement and circulation in Northern and Southern hemispheres. What weather is associated with each? *(20 Marks)*
* **Q.No.3**: Differentiate between Valley Glacier and Ice Sheet. Discuss the depositional landforms of Valley Glacier in detail. *(20 Marks)*
* **Q.No.4**: What is an earthquake, and how does it arise? Describe the relationship between earthquakes and plate tectonics. *(20 Marks)*
* **Q.No.5**: Describe in detail the various movements of ocean water and the factors responsible. *(20 Marks)*
* **Q.No.6**: Why does salinity of ocean water vary from place to place? What are the determining factors? *(20 Marks)*
* **Q.No.7**: Define Map. What are the essentials of a map? Why do we need projections for drawing a map? *(20 Marks)*
* **Q.No.8**: Write notes on any TWO: (a) Karst Topography (b) Ocean floor morphology (c) Temperature inversion (d) Mercator's Projection. *(20 Marks)*

---

### 📝 KPPSC PMS GEOGRAPHY PAPER-I — 2008
* **Q.No.1**: Describe the rock cycle. How can we identify igneous rocks on the basis of mineral contents? Explain types of sedimentary rocks. *(20 Marks)*
* **Q.No.2**: Describe the mechanism and types of seismic waves. Define the causes of earthquakes and write an account of earthquakes in Pakistan. *(20 Marks)*
* **Q.No.3**: "Solar input energy is equal to Earth's output energy." Explain. Describe the classification of air masses with examples. *(20 Marks)*
* **Q.No.4**: Explain the mechanism of oceanic currents that produce circulation in ocean water. *(20 Marks)*
* **Q.No.5**: Define the internal structure of the Earth from Asthenosphere to inner core. Give reasons why the core is composed of iron. *(20 Marks)*
* **Q.No.6**: How can we identify different types of maps? Characteristics of Cylindrical and Conical Projections. *(20 Marks)*
* **Q.No.7**: Short notes on any TWO: (1) Weathering (2) Volcanic activity (3) River landforms (4) Distribution of temperature. *(20 Marks)*
"""

with open(P1_ALL_MD, "w", encoding="utf-8") as f:
    f.write(p1_all_text)
compile_pdf(P1_ALL_MD, P1_ALL_PDF, "KPPSC PMS Geography Paper-I Past Papers Compendium (2008-2025)")

# ==================== 3. PAPER-II COMPILATION ====================
P2_ALL_MD = os.path.join(PAPERS_DIR, "KPPSC_PMS_Geography_Paper_2_Past_Papers_2008_2025.md")
P2_ALL_PDF = os.path.join(PAPERS_DIR, "KPPSC_PMS_Geography_Paper_2_Past_Papers_2008_2025.pdf")

p2_all_text = """# Khyber Pakhtunkhwa Public Service Commission (KPPSC)
## Provincial Management Service (PMS) Competitive Examination
### GEOGRAPHY — PAPER-II (HUMAN, ECONOMIC AND REGIONAL GEOGRAPHY)
#### Compendium of Solved Past Papers (2008 – 2025)

> **Time Allowed**: 3 Hours | **Maximum Marks**: 100 Marks  
> **Instructions**: Attempt **ANY FIVE (5)** questions. All questions carry equal marks (**20 Marks** each). Support your answers with relevant geographical sketches, diagrams, and maps where appropriate.

---

### 📝 KPPSC PMS GEOGRAPHY PAPER-II — 2025
* **Q.No.1**: Explain the concept of bridging the “green” and “brown” agendas in sustainable urbanization. Evaluate how integrated environmental and infrastructural planning can address the challenges of rapid urban growth in developing countries. *(20 Marks)*
* **Q.No.2**: Discuss the physical, technological, and socio-economic factors that influence agricultural patterns around the world. Assess the significance of mixed farming as an adaptive agricultural system in modern economies. *(20 Marks)*
* **Q.No.3**: Analyze the major threats posed by overpopulation at global and national scales. Critically examine the demographic pressures facing Pakistan and their implications for economic development, resource management, and social stability. *(20 Marks)*
* **Q.No.4**: Describe the global spatial distribution of tropical savanna forests. Evaluate their economic significance and discuss the environmental challenges associated with resource extraction in these regions. *(20 Marks)*
* **Q.No.5**: Explain the geography of development and critically assess the factors responsible for regional disparities between developed and underdeveloped countries. How do these disparities influence global economic and political relations? *(20 Marks)*
* **Q.No.6**: Evaluate the assertion that Pakistan is fundamentally an agro-based economy. Discuss the structural characteristics, strengths, weaknesses, and transformation potential of Pakistan's agricultural sector. *(20 Marks)*
* **Q.No.7**: Discuss the spatial distribution, raw material base, and locational factors of chemical industries in Pakistan. Evaluate their role in national economic development and identify the major constraints hindering their expansion. *(20 Marks)*
* **Q.No.8**: Explain the major economic and socio-demographic indicators used to assess national development. *(20 Marks)*

---

### 📝 KPPSC PMS GEOGRAPHY PAPER-II — 2022
* **Q.No.1**: Discuss the major factors responsible for the distribution and density of world population with suitable examples. *(20 Marks)*
* **Q.No.2**: Examine the process of rapid urbanization in Pakistan. Discuss the resultant socio-economic and environmental problems and suggest practical policy remedies. *(20 Marks)*
* **Q.No.3**: Define and classify economic activities into primary, secondary, and tertiary sectors. Discuss how economic development alters occupational structures globally. *(20 Marks)*
* **Q.No.4**: Examine Alfred Weber’s Least Cost Theory of industrial location with reference to the Iron and Steel industry. *(20 Marks)*
* **Q.No.5**: "There is a paradigm shift from Geopolitics to Geoeconomics." In your opinion, how should Pakistan position itself to accelerate national economic development? *(20 Marks)*
* **Q.No.6**: Analyze the agricultural sector of Pakistan: Major crops, canal irrigation network, and the emerging challenges of water scarcity and land degradation. *(20 Marks)*
* **Q.No.7**: Critically evaluate the China-Pakistan Economic Corridor (CPEC) as a game-changer for regional transit, industrial modernization, and trade expansion. *(20 Marks)*
* **Q.No.8**: Write comprehensive short notes on any TWO of the following: *(10 + 10 = 20 Marks)*  
  (a) Environmental Determinism vs. Possibilism  
  (b) Structural shifts in the foreign trade of Pakistan  
  (c) Major commercial fishing grounds of the world  
  (d) Christaller’s Central Place Theory

---

### 📝 KPPSC PMS GEOGRAPHY PAPER-II — 2018
* **Q.No.1**: Discuss the concepts of Environmental Determinism and Possibilism with historical examples. How has modern technology altered this relationship? *(20 Marks)*
* **Q.No.2**: Highlight the scope and status of Human Geography with special reference to its application for sustainable human development. *(20 Marks)*
* **Q.No.3**: What is urbanization? Discuss the urbanization process with special reference to Pakistan's major metropolitan centers. *(20 Marks)*
* **Q.No.4**: What are the factors responsible for the uneven distribution of population with special reference to Pakistan? *(20 Marks)*
* **Q.No.5**: There is a "paradigm shift from Geo-political to Geo-Economic." In your opinion, which orientation should Pakistan prioritize for long-term prosperity? *(20 Marks)*
* **Q.No.6**: Discuss the main economic activities of Pakistan with special focus on the agricultural sector and rural livelihood. *(20 Marks)*
* **Q.No.7**: In your opinion, how can Pakistan derive maximum socio-economic benefit from CPEC, the flagship project of the Belt and Road Initiative (BRI)? *(20 Marks)*
* **Q.No.8**: Write short notes on any TWO of the following: *(10 + 10 = 20 Marks)*  
  (a) Actionable policy measures for increasing Pakistan's export earnings  
  (b) Strategic and economic relations between Pakistan and China  
  (c) Poverty alleviation strategies through spatial planning  
  (d) Trends and patterns of population migration in Pakistan

---

### 📝 KPPSC PMS GEOGRAPHY PAPER-II — 2016
* **Q.No.1**: Narrate the story of world population growth with the help of the J-shaped curve and present your views on future demographic projections. *(20 Marks)*
* **Q.No.2**: Point out the major problems resulting from rapid urbanization in Pakistan and suggest workable remedies to overcome these crises. *(20 Marks)*
* **Q.No.3**: Classify and explain economic activities by providing suitable examples from various regions of the world. *(20 Marks)*
* **Q.No.4**: "Pakistan is a resource-rich country." Critically examine this statement with reference to physical, mineral, and agricultural resources. *(20 Marks)*
* **Q.No.5**: Among agricultural commodities, which cereal and cash crops dominate world trade? Give an account of commercial wheat production. *(20 Marks)*
* **Q.No.6**: "The spatial distribution of power resources influences world geopolitics and economic supremacy." Discuss. *(20 Marks)*
* **Q.No.7**: Explain the causes and socio-economic consequences of international migration (push-pull factors, remittances, brain drain). *(20 Marks)*
* **Q.No.8**: Write comprehensive notes on any TWO of the following: *(10 + 10 = 20 Marks)*  
  (a) Commercial fishing regions of the world  
  (b) World's linguistic and spatial language patterns  
  (c) Intra-regional political relations and economic cooperation in South Asia (SAARC)  
  (d) Environmental and economic benefits of forest resources

---

### 📝 KPPSC PMS GEOGRAPHY PAPER-II — 2013
* **Q.No.1**: What is Economic Geography? Present a comparative survey of secondary and tertiary economic activities with occupational examples. *(20 Marks)*
* **Q.No.2**: "A large number of industrialized economies are located in the Western Hemisphere." Discuss. Examine the implications of Weber’s Model of industrial location. *(20 Marks)*
* **Q.No.3**: Discuss methods of conservation of raw material resources. Critically evaluate the potential for alternative energy sources. *(20 Marks)*
* **Q.No.4**: Differentiate between intensive and extensive agriculture. Explain world agriculture under specializations, systems, and problems. *(20 Marks)*
* **Q.No.5**: Elaborate pioneer settlement patterns. "The pattern of settlement characterizing human use of the earth is strongly related to migration." Discuss. *(20 Marks)*
* **Q.No.6**: To what extent can the major features of world population distribution be understood in terms of physical environment? Explain Malthusian Theory. *(20 Marks)*
* **Q.No.7**: Highlight the patterns of Foreign Trade of Pakistan. What major structural changes have taken place in exports and imports? *(20 Marks)*
* **Q.No.8**: Write comprehensive notes on any TWO: (a) Christaller’s Central Place Theory (b) Urbanization (c) Man-Environment Relationship (d) Major geographic qualities of South Asia. *(20 Marks)*

---

### 📝 KPPSC PMS GEOGRAPHY PAPER-II — 2010
* **Q.No.1**: Differentiate between Environmental Determinism and Possibilism. Who were the key propagators? Discuss concept evolution. *(20 Marks)*
* **Q.No.2**: "Population of the world is unevenly distributed." Elaborate with reference to physical and human controlling factors. *(20 Marks)*
* **Q.No.3**: Define urbanization and discuss major problems associated with urbanization in both developing and developed countries. *(20 Marks)*
* **Q.No.4**: What yardsticks are used for defining Rural and Urban Settlements? Describe locational factors of rural settlements. *(20 Marks)*
* **Q.No.5**: Define economic activities and discuss the evolution of secondary and tertiary activities with changing global consumption patterns. *(20 Marks)*
* **Q.No.6**: "Rice feeds more people and needs more labour force compared to other cereal crops." Describe geographical factors for rice cultivation in South Asia. *(20 Marks)*
* **Q.No.7**: "Power resources are playing a key role in economic development." Describe the production and trade of Crude Oil in the world. *(20 Marks)*
* **Q.No.8**: Write notes on any TWO: (a) Rural-urban migration (b) Mega-cities (c) Major fishing grounds (d) Trade balance of Pakistan. *(20 Marks)*

---

### 📝 KPPSC PMS GEOGRAPHY PAPER-II — 2008
* **Q.No.1**: Explain the concept of Environmental Determinism. How does the physical environment influence human activities? *(20 Marks)*
* **Q.No.2**: What is meant by Population Density? Write down areas of major world population concentrations. *(20 Marks)*
* **Q.No.3**: Define Economic Geography. Detailed account of human economic activities related to production and services. *(20 Marks)*
* **Q.No.4**: "Despite modernization, religion still dominates human behavior." Geographical distribution of Islam and Christianity. *(20 Marks)*
* **Q.No.5**: Distribution and production of cotton in the world with special reference to South Asia. *(20 Marks)*
* **Q.No.6**: Details about the Iron and Steel Industry of the USA. *(20 Marks)*
* **Q.No.7**: Short notes on any TWO: (i) Mixed Farming (ii) The Indus Plain (iii) Trade of Pakistan (iv) Coal deposits of the USA. *(20 Marks)*
"""

with open(P2_ALL_MD, "w", encoding="utf-8") as f:
    f.write(p2_all_text)
compile_pdf(P2_ALL_MD, P2_ALL_PDF, "KPPSC PMS Geography Paper-II Past Papers Compendium (2008-2025)")

print("Syllabus and Past Paper Compendiums successfully compiled!")
