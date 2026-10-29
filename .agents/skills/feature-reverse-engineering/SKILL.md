---
name: feature-reverse-engineering
description: Systematically reverse engineers, dissects, and reconstructs features from third-party repositories and apps. Use when analyzing an external codebase to replicate or adapt its functionality.
---

# Feature Reverse Engineering & Reconstruction Guide

This skill guides the deconstruction and clean reconstruction of features from reference applications.

## 4-Step Reverse Engineering Workflow

### Step 1: Feature Inventory & Mapping
- Scan reference repos (e.g., in `apps_toolkit/reference_repos/`).
- Identify the core user-facing features (e.g., Summarization, MCQ generation, Mindmap rendering, QA pair extraction).
- Trace entry points (UI route, CLI command, or script trigger) to the backend implementation.

### Step 2: Dependency & Pipeline Dissection
- **Inputs**: What format does the feature ingest? (Raw text, PDF pages, Markdown headers).
- **Processing Logic**: Is it rule-based NLP (spaCy, NLTK, KeyBERT) or LLM prompt orchestration (system prompt, temperature, few-shot examples)?
- **Dependencies**: Filter out heavy or obsolete libraries (e.g. outdated LangChain, legacy OpenAI 0.28) and identify lightweight modern equivalents.
- **Outputs**: Formats (JSON, Markdown, Graphviz DOT/PNG, PPTX, MP3).

### Step 3: Security & Code Sanitization
- Check for hardcoded API keys or vulnerable dependencies.
- Remove synchronous blocking network calls or unvetted external scripts.
- Ensure all execution happens inside controlled virtual environments or local scripts.

### Step 4: Clean Modular Reconstruction
- Extract pure logic into modular Python classes/functions.
- Separate UI layer (Streamlit/Web) from Core Engine.
- Provide clean CLI and programmatic interfaces so features can be chained together or run as batch jobs against study notes.
