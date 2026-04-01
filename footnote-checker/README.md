# Footnote Checker for Word

A Microsoft Word task pane add-in for legal footnote review.

This tool extracts footnotes directly from Word, preserves formatting via OOXML, splits authorities into sources, classifies them, and displays structured citation data in a side pane.

---

## Features

- Extracts footnotes directly from Word
- Preserves formatting (italics, bold, small caps, underline) using OOXML
- Splits footnotes into individual sources
- Aligns sources with original Word run structure
- Displays sources in a navigable side pane
- Renders formatting inside the add-in UI
- Classifies sources (case, statute, etc.)
- Generates suggested filenames for each source

---

## Current Status

- Working prototype
- OOXML pipeline integrated
- Source splitting and run alignment implemented (using original parser logic)
- Source classification integrated
- Filename suggestion integrated
- Warnings and validation not yet implemented

---

## Architecture Overview

Word Add-in (React)  
→ Extracts footnotes + OOXML  
→ Sends structured payload to FastAPI backend  
→ Backend reconstructs runs and splits sources  
→ Applies classification + filename logic  
→ Returns structured data  
→ UI renders sources with formatting  

---

## Project Structure