# Prompt Log

This document tracks key prompts and decisions used to develop the Footnote Checker.

---

## Product Vision

- Build a Word add-in that checks footnotes directly inside Word
- Show footnotes one at a time with navigation
- Split each footnote into sources
- Display rules, warnings, and suggested labels

---

## Architecture Decisions

- Use a Word add-in instead of a file-upload web app
- Use FastAPI for the backend
- Preserve formatting using OOXML instead of plain text
- Maintain run-level structure throughout the pipeline

---

## Development Approach

- Incremental “Step X” workflow
- Modify one file at a time
- Validate each step before proceeding
- Preserve existing logic from `footnote_processor.py` wherever possible

---

## Key Technical Prompts

- Extract footnotes directly from Word using Word API
- Use OOXML to reconstruct runs
- Preserve formatting in the UI
- Replace naive splitting with `split_sources`
- Replace custom alignment with `assign_source_runs`
- Integrate `_classify_source` incrementally
- Integrate `_suggest_filename` after classification is validated

---

## Lessons Learned

- Plain text pipelines break citation fidelity
- Formatting is essential for legal citation parsing
- Reusing proven parsing logic is better than rewriting
- Incremental debugging prevents major breakage
- Run-to-source alignment is the hardest part of the pipeline

---

## Current Development State

- OOXML → runs pipeline working
- Source splitting and alignment working
- Classification working
- Filename suggestions working

---

This file will be updated as new functionality is added.