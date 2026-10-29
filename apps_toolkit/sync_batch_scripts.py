import re
import os
import sys
from update_ir_exam_questions_linked import EXAM_QUESTIONS_LINKED_MAP

files = [
    (r"H:\My Drive\PMS KPK 2022 notes\apps_toolkit\build_ir_batch1_notes.py", [2, 3, 4, 5, 6]),
    (r"H:\My Drive\PMS KPK 2022 notes\apps_toolkit\build_ir_batch2_notes.py", [7, 8, 9, 10, 11, 12]),
    (r"H:\My Drive\PMS KPK 2022 notes\apps_toolkit\build_ir_batch3_notes.py", [13, 14, 15, 16, 17]),
    (r"H:\My Drive\PMS KPK 2022 notes\apps_toolkit\build_ir_batch4_notes.py", [18, 19, 20, 21, 22, 23]),
]

for fpath, q_nums in files:
    with open(fpath, "r", encoding="utf-8") as f:
        content = f.read()
    
    for q in q_nums:
        q_val = EXAM_QUESTIONS_LINKED_MAP[q]
        meta_pat = rf"(q{q}_meta\s*=\s*\[.*?)(?:\(\"Past Paper Alignment\",\s*\"[^\"]*\"\))(.*?\])"
        m = re.search(meta_pat, content, re.DOTALL)
        if m:
            content = content[:m.start()] + m.group(1) + f'("EXAM QUESTIONS LINKED:", {repr(q_val)})' + m.group(2) + content[m.end():]
            print(f"Updated Q{q} in {os.path.basename(fpath)}")
        else:
            print(f"Could not find meta block for Q{q} in {os.path.basename(fpath)}")
            
    with open(fpath, "w", encoding="utf-8") as f:
        f.write(content)

print("Batch scripts synchronization complete!")
