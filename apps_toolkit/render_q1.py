import subprocess
import tempfile
import os

EDGE_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
md_path = r"H:\My Drive\PMS KPK 2022 notes\1. European History Paper 1 and Paper 2\Paper 1\Q1_Napoleon_Rise_and_Legacy_2025_Solved.md"
out_pdf = r"H:\My Drive\PMS KPK 2022 notes\1. European History Paper 1 and Paper 2\Paper 1\Q1_Napoleon_Rise_and_Legacy_2025_Solved.pdf"

with open(md_path, "r", encoding="utf-8") as f:
    text = f.read()

html_body = ""
for para in text.split("\n\n"):
    para = para.strip()
    if not para:
        continue
    if para.startswith("# "):
        html_body += f"<h1 style='text-align:center; font-size:17pt; margin-bottom:4px;'>{para[2:]}</h1>"
    elif para.startswith("## "):
        html_body += f"<h2 style='font-size:14pt; color:#0E2B50; margin-bottom:8px;'>{para[3:]}</h2>"
    elif para.startswith("### "):
        html_body += f"<h3 style='font-size:13pt; color:#1a365d; border-bottom:1px solid #cbd5e0; padding-bottom:3px; margin-top:16px;'>{para[4:]}</h3>"
    elif para.startswith("> "):
        html_body += f"<blockquote style='background:#f7fafc; border-left:4px solid #3182ce; padding:8px 12px; font-style:italic; margin:12px 0;'>{para[2:]}</blockquote>"
    elif para.startswith("---"):
        html_body += "<hr style='border:0; border-top:1px solid #e2e8f0; margin:14px 0;'>"
    else:
        # replace bold formatting
        parts = para.split("**")
        p_html = ""
        for idx, part in enumerate(parts):
            if idx % 2 == 1:
                p_html += f"<strong>{part}</strong>"
            else:
                p_html += part
        html_body += f"<p style='text-align:justify; line-height:1.55; font-size:12pt; margin-bottom:10px;'>{p_html}</p>"

full_html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  @page {{ size: A4; margin: 20mm 18mm 20mm 18mm; }}
  body {{ font-family: 'Times New Roman', Times, serif; color: #1a202c; }}
</style>
</head>
<body>
{html_body}
<div style='margin-top:25px; border-top:1px solid #a0aec0; padding-top:6px; text-align:center; font-size:10pt; color:#718096;'>
  Khyber Pakhtunkhwa Public Service Commission (KPPSC) &bull; PMS European History &bull; Paper I Model Answers
</div>
</body>
</html>"""

with tempfile.NamedTemporaryFile("w", delete=False, suffix=".html", encoding="utf-8") as tf:
    tf.write(full_html)
    tmp_path = tf.name

cmd = [EDGE_PATH, "--headless", "--disable-gpu", f"--print-to-pdf={out_pdf}", "--no-pdf-header-footer", tmp_path]
subprocess.run(cmd, check=True)
try:
    os.remove(tmp_path)
except Exception:
    pass
print("Q1 PDF successfully generated:", out_pdf)
