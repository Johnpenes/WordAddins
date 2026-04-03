# Footnote Checker (Word Add-in)

A Microsoft Word add-in that analyzes footnotes, extracts their source structure, and maps each footnote to the exact sentence it supports in the document.

## 🚀 Overview

This tool is designed for legal writing, academic papers, and citation-heavy workflows. It:

- Parses Word documents and extracts footnotes
- Identifies the exact **body sentence** associated with each footnote
- Splits footnotes into **individual sources**
- Applies rule-based analysis to each source
- Suggests structured file names for research organization
- Navigates directly to the relevant sentence inside Word

## 🧠 Key Features

### 1. Precise Sentence Mapping
Each footnote is linked to the **exact sentence it supports**, using XML-aware parsing (not naive text matching).

### 2. Smart Word Navigation
- Click **Next / Prev / Go**
- The add-in selects the corresponding sentence directly in Word
- Uses **anchor-based range expansion** to bypass Word’s search limits

### 3. Source Extraction
Footnotes are split into individual sources:
- Preserves formatting (italic, bold, small caps, etc.)
- Handles multiple citations within one footnote

### 4. Rule Detection
Each source is analyzed for:
- Citation structure
- Formatting issues
- Potential inconsistencies

### 5. Suggested File Naming
Automatically generates clean, research-friendly filenames for each source.

---

## 🏗 Architecture

### Frontend (Word Add-in)
- React + Office.js
- Handles:
  - UI
  - Navigation
  - Word selection logic

### Backend (Python)
- FastAPI
- Handles:
  - Footnote parsing
  - XML processing
  - Sentence extraction
  - Source classification

---

## 🔍 How Sentence Selection Works

Word’s search API cannot handle long strings reliably. This project uses a custom approach:

1. Extract full sentence from backend
2. Generate:
   - **start anchor** (first ~80 chars)
   - **end anchor** (last ~80 chars)
3. In Word:
   - Search for both anchors
   - Use `range.expandTo()` to select the full sentence

This avoids:
- Word’s search length limits
- Partial or incorrect highlighting

---

## 🛠 Setup

### 1. Clone the repo
```bash
git clone https://github.com/YOUR_USERNAME/footnote-checker.git
cd footnote-checker

### 2. Backend setup
cd backend
python -m venv venv
source venv/bin/activate   # macOS
pip install -r requirements.txt
uvicorn footnote_api:app --reload --port 3000

### 3. Frontend setup
cd frontend
npm install
npm start

### 4. Load into Word
# Use Office Add-in sideloading
# Requires WordApi 1.5+, macOS or Windows, Node.js + Python 3.9+


