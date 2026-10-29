# Apps Toolkit & Reference Repositories

This toolkit contains reference implementations scanned, audited, and cataloged for reverse engineering and feature reconstruction:

## Reference Repositories

### 1. [NoteGenius](./reference_repos/NoteGenius)
- **Source**: `AAC-Open-Source-Pool/NoteGenius`
- **Core Features**:
  - `flashcards.py`: Keyword extraction via RAKE and fuzzy sentence matching.
  - `mcq_generator.py` & `ques_gen.py`: SpaCy + WordNet distractor MCQ generation.
  - `mindmap.py`: Syntactic clause splitting and Graphviz diagram generation.
  - `mnemonics.py`: KeyBERT keyword extraction with acronym mnemonic generator.
  - `pptgen.py`: Auto-generation of PowerPoint slides from text chunks.
  - `text to speech.py`: Audio synthesis via gTTS.
  - `answer to the questions.py`: Local Hugging Face QA pipeline (`deepset/tinyroberta-squad2`).

### 2. [AI-Notes-Summarizer](./reference_repos/AI-Notes-Summarizer)
- **Source**: `KaisoX24/AI-Notes-Summarizer`
- **Core Features**:
  - Streamlit UI for PDF upload and processing.
  - Structured prompt engineering for educational summarization + MCQ generation.
  - Downloadable summary exports.

### 3. [question_extractor](./reference_repos/question_extractor)
- **Source**: `nestordemeure/question_extractor`
- **Core Features**:
  - Recursive markdown chunking by heading hierarchy to match model context tokens.
  - Dual-phase extraction: generates question candidate list first, then generates grounded answers based strictly on source text.
  - Async throttled batch execution.

## Agent Team
- **`feature-reverse-engineer`**: Dissects reference code, maps algorithms, and produces reconstruction plans.
- **`study-app-builder`**: Builds unified, high-performance study tools tailored for PMS KPK notes.
