---
name: pms-exam-notes-processor
description: Systematically decomposes KPPSC PMS and CSS syllabi and past exam papers (2008–2025) to generate publishable, 5-page institutional master study notes with verbatim question linking, ASCII mindmaps, comparative matrices, and scholar quotation banks.
---

# PMS Exam Notes Processor Skill

## 🎯 Purpose
This skill coordinates the end-to-end extraction, thematic clustering, past paper cross-referencing, and generation of competitive examination notes for civil service aspirants preparing for the Khyber Pakhtunkhwa Public Service Commission (KPPSC PMS) and Federal Public Service Commission (CSS).

---

## 📋 Standard Workflow

### Phase 1: Syllabus & Past Paper Ingestion
1. **Official Syllabus Mapping**:
   - Extract official syllabus sections, recommended reading lists, and structural weightage.
2. **Multi-Decade Past Paper Deconstruction (2008–2025)**:
   - Ingest all past examination papers from official KPPSC PDF archives.
   - Transcribe every question verbatim, recording Year, Question Number, and Marks.
   - Cluster repeated, overlapping, and rephrased questions under unified Master Topics.

### Phase 2: Master Topic Synthesis
For every clustered topic, execute the **7-Part Institutional Blueprint**:
1. **Header Metadata**: Subject, Paper, Exact Verbatim Past Paper Questions Linked, Syllabus Theme, Historiographical / Theoretical Lens.
2. **Executive Exam Takeaways**: 3–5 bulleted core analytical assertions.
3. **Conceptual Flowchart / Mind Map**: Monospace ASCII tree diagram showing causal mechanisms.
4. **Structured Chronology / Comparative Analytical Matrix**: 4–6 row comparison or chronological milestone table.
5. **Comprehensive Essay Body**: 2,200 – 2,800 words containing academic introduction, 6–8 bolded thematic subheadings, institutional data, and forward-looking conclusion.
6. **Historiographical & Theorist Quotation Bank**: 5 verbatim scholar quotations with work, date, and exam application.
7. **Strategic Exam Recommendations & Traps to Avoid**: Concrete examiner rubric advice and common candidate pitfalls.

### Phase 3: Triple-Format Publication
1. **GitHub Flavored Markdown (`.md`)**:
   - Clean headings, Markdown tables, callout blocks (`> [!NOTE]`), and code blocks for ASCII diagrams.
   - Formatted for instant viewing on GitHub, mobile devices, and Obsidian/Notion.
2. **Native Microsoft Word (`.docx`)**:
   - Styled using python-docx with institutional executive theme: `#0E2B50` (Primary Navy), `#1B4965` (Secondary Teal), `#F0F4F8` (Callout background), `Consolas` font for mindmaps, and 0.8-inch margins.
3. **Publishable Print PDF (`.pdf`)**:
   - High-fidelity typography with zero layout shifts, suitable for double-sided spiral binding.

---

## 🛠️ Verification Checklist
Before releasing any master topic:
- [ ] Verbatim Question Check: Are all past paper questions quoted word-for-word with year and question number?
- [ ] Word Count Check: Is the essay body substantive (minimum 2,000 words)?
- [ ] Mind Map Check: Does the ASCII diagram render cleanly without wrapping?
- [ ] Quotation Check: Are all 5 quotations verified, dated, and attributed to genuine scholars?
- [ ] Strategy Check: Are examiner traps specific and actionable rather than generic advice?
