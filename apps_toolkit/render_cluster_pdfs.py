import subprocess
import tempfile
import os

EDGE_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
paper1_dir = r"H:\My Drive\PMS KPK 2022 notes\1. European History Paper 1 and Paper 2\Paper 1"

def render_md_to_pdf_and_file(md_text, base_name, title):
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

# Build text for Q2
from build_napoleon_cluster_notes import essay_A, essay_B

q2_md = """# KPPSC PMS European History &bull; Paper I

## Question 2:
> **"In the light of his internal reforms, discuss whether Napoleon Bonaparte promoted or halted the French Revolution. To what extent is his claim 'I am the Revolution' justified by his civil, judicial, and educational reforms?"**  
> *(Total Marks: 20 &bull; Word Count: ~650 words &bull; Covers 2022 Q.2, 2016 Q.2 & 2013 Q.2)*

---
"""
for heading, body in essay_A:
    q2_md += f"\n### {heading}\n{body}\n"

# Build text for Q3
q3_md = """# KPPSC PMS European History &bull; Paper I

## Question 3:
> **"Analyze the Continental System of Napoleon Bonaparte. Why did it fail to bring Great Britain to its knees, and how did it contribute to Napoleon's downfall?"**  
> *(Total Marks: 20 &bull; Word Count: ~650 words &bull; Covers 2018 Q.2 & 2013 Q.3)*

---
"""
for heading, body in essay_B:
    q3_md += f"\n### {heading}\n{body}\n"

render_md_to_pdf_and_file(q2_md, "Q2_Napoleon_Internal_Reforms_and_Revolution_Promoted_or_Halted_Master_Notes", "Reforms & Legacy")
render_md_to_pdf_and_file(q3_md, "Q3_Napoleon_Continental_System_and_Downfall_Master_Notes", "Continental System & Downfall")
