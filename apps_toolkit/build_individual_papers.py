import os
import subprocess

BASE_DIR = r"H:\My Drive\PMS KPK 2022 notes\3. Geography Paper 1 and Paper 2\Past Papers"
EDGE_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

for paper_num, fname in [(1, "KPPSC_PMS_Geography_Paper_1_Past_Papers_2008_2025.md"), (2, "KPPSC_PMS_Geography_Paper_2_Past_Papers_2008_2025.md")]:
    with open(os.path.join(BASE_DIR, fname), "r", encoding="utf-8") as f:
        content = f.read()
    
    sections = content.split("### 📝 KPPSC PMS GEOGRAPHY PAPER-")
    for sec in sections[1:]:
        year = sec.split("\n")[0].split("—")[1].strip()
        sec_content = f"""# KPPSC PMS Geography Paper-{paper_num} ({year})

> **Examination**: KPPSC PMS Optional Subject (100 Marks | 3 Hours)  
> **Instructions**: Attempt ANY FIVE questions (20 Marks each). All questions carry equal marks. Support your answers with sketches and diagrams.

---

### Examination Questions
""" + "\n".join(sec.split("\n")[1:]).strip()
        
        md_file = os.path.join(BASE_DIR, f"KPPSC_PMS_Geography_Paper_{paper_num}_{year}.md")
        pdf_file = os.path.join(BASE_DIR, f"KPPSC_PMS_Geography_Paper_{paper_num}_{year}.pdf")
        with open(md_file, "w", encoding="utf-8") as mf:
            mf.write(sec_content)
            
        items_html = ""
        for line in sec.split("\n")[1:]:
            s = line.strip()
            if s.startswith("* **"):
                formatted = s.replace("* **", "<strong>").replace("**: ", "</strong>: ").replace("*(", "<em>(").replace(")*", ")</em>")
                items_html += f"<li style='margin-bottom: 12px;'>{formatted}</li>"

        html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>KPPSC PMS Geography Paper {paper_num} ({year})</title>
<style>
    @page {{ size: A4; margin: 20mm; }}
    body {{ font-family: 'Segoe UI', Calibri, Arial, sans-serif; font-size: 11pt; line-height: 1.6; color: #1A1A1A; margin: 0; }}
    h1 {{ color: #0E2B50; font-size: 16pt; border-bottom: 2px solid #0E2B50; padding-bottom: 4px; }}
    blockquote {{ border-left: 4px solid #0E2B50; background: #F0F4F8; padding: 10px 14px; margin: 12px 0; }}
    ul {{ margin-left: 20px; }}
</style>
</head>
<body>
<h1>KPPSC PMS Geography Paper-{paper_num} ({year})</h1>
<blockquote>
<strong>KPPSC Provincial Management Service (PMS) Examination</strong><br>
Subject: Geography Paper-{paper_num} | Time Allowed: 3 Hours | Total Marks: 100 Marks<br>
Attempt ANY FIVE questions (20 Marks each). All questions carry equal marks.
</blockquote>
<hr/>
<ul>
{items_html}
</ul>
</body>
</html>
"""
        temp_html = pdf_file.replace(".pdf", "_temp.html")
        with open(temp_html, "w", encoding="utf-8") as hf:
            hf.write(html)
        subprocess.run([EDGE_PATH, "--headless", "--disable-gpu", f"--print-to-pdf={pdf_file}", temp_html], check=True)
        if os.path.exists(temp_html):
            os.remove(temp_html)
        print(f"Compiled: Paper {paper_num} - {year}")

print("All individual past papers successfully generated and compiled!")
