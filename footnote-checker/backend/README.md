# SCU Law Review — Footnote Formatter

A web application that automates footnote cite-checking for law review editors. Upload a `.docx` file, specify a footnote range, and receive a formatted PDF or Excel report with each footnote broken down into individual citation sources — with rich text formatting, body sentence context, and Bluebook rule suggestions preserved.

Built with Claude Code and VS Code. Deployed on Vercel: [lawreview.vercel.app](https://lawreview.vercel.app)

---

## What it does

Law review cite-checking requires editors to identify every source cited in a footnote, verify its format against the Bluebook, and cross-reference it with the body text. This tool automates the extraction and organization step.

Given a `.docx` manuscript, the formatter:

1. **Parses footnote XML directly** from the Word document — preserving italics, bold, and small caps that plain-text export loses
2. **Splits each footnote** into individual citation sources using a rule-based parser that detects citation signals (`See`, `Id.`, `Cf.`), sentence boundaries, semicolons, and abbreviation patterns
3. **Filters out commentary** — prose sentences that don't contain bibliographic markers are excluded from the output
4. **Extracts body sentence context** — the sentence in the main text that corresponds to each footnote, so editors can verify citation relevance
5. **Classifies each source by Bluebook rule** using formatting cues (e.g. italic "v." → Rule 10 for cases; volume + L. Rev. → Rule 16 for journal articles)
6. **Exports a structured report** as either a PDF table or Excel spreadsheet with merged cells, wrapped text, and per-source rule suggestions

---

## Output columns

| FN # | Body Sentence Context | Source Text | Bluebook Rules |
|------|-----------------------|-------------|----------------|
| 42 | The court held that… | *Murphy v. State*, 123 F.3d 456 (9th Cir. 2020). | Rule 10 (Cases); Rule 10.2 (Case Names); Rule 10.4 (Reporters) |

The FN # and Body Sentence cells span all source rows belonging to the same footnote.

---

## Tech stack

| Layer | Technology |
|-------|-----------|
| Backend | Python 3, Flask |
| Docx parsing | lxml, zipfile (raw Word XML) |
| PDF generation | ReportLab |
| Excel generation | XlsxWriter |
| Frontend | Vanilla JS, HTML/CSS |
| Deployment | Vercel (serverless) |

---

## Local setup

```bash
git clone https://github.com/yourhandle/footnote-web
cd footnote-web
pip install -r requirements.txt
python app.py
```

Then open `http://localhost:5000` in your browser.

---

## Deployment

The app is configured for Vercel via `vercel.json`. To deploy your own instance:

```bash
npm i -g vercel
vercel
```

Temporary files are written to `/tmp` to comply with Vercel's serverless filesystem constraints.

---

## Limitations

- Input must be a `.docx` file (not `.doc` or PDF)
- Maximum footnote range: 500 footnotes per request
- Bluebook rule classification is heuristic (~85–90% accuracy on standard citation types); edge cases are flagged with a NOTE in the output
- Rich text in the Excel export requires Microsoft Excel — Numbers (macOS) does not render per-run formatting from XlsxWriter correctly

---

## License

MIT
