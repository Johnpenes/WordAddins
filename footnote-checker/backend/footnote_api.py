from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
import tempfile
import shutil

from footnote_processor import _build_rows, split_sources, assign_source_runs, _classify_source, _suggest_filename
from lxml import etree


app = FastAPI()

# Allow Word add-in (localhost) to call this
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.post("/process")
async def process_docx(file: UploadFile = File(...)):
    # Save uploaded file temporarily
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        shutil.copyfileobj(file.file, tmp)
        temp_path = tmp.name

    # Run your existing logic
    rows, _, _, _ = _build_rows(temp_path, 1, 200)

    # Convert rows -> structured JSON
    footnotes = {}
    for fn, body, src_runs, bb, _, label, _, _ in rows:
        fn_num = int(fn)
        if fn_num not in footnotes:
            footnotes[fn_num] = {
                "number": fn_num,
                "context": body,
                "sources": [],
            }

        text = "".join(t for t, _ in src_runs)

        footnotes[fn_num]["sources"].append(
            {
                "text": text,
                "rules": bb,
                "warnings": "",
                "fileLabel": label,
            }
        )

    return {"footnotes": list(footnotes.values())}


@app.post("/process-text")
async def process_text(data: dict):
    text = data.get("text", "")

    # TEMP: basic response for Word integration
    return {
        "footnotes": [
            {
                "number": 1,
                "context": text[:200],
                "sources": [],
            }
        ]
    }


# Helper to parse OOXML and extract runs (text + formatting)
def _runs_from_ooxml(ooxml: str):
    if not ooxml:
        return []
    try:
        root = etree.fromstring(ooxml.encode("utf-8"))
    except Exception:
        return []

    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    runs = []

    for r in root.xpath(".//w:r", namespaces=ns):
        text_parts = []
        for t in r.xpath("./w:t", namespaces=ns):
            text_parts.append(t.text or "")
        for tab in r.xpath("./w:tab", namespaces=ns):
            text_parts.append("\t")
        for br in r.xpath("./w:br", namespaces=ns):
            text_parts.append("\n")

        text = "".join(text_parts)
        if not text:
            continue

        rpr = r.find("w:rPr", namespaces=ns)
        fmt = {
            "bold": False,
            "italic": False,
            "small_caps": False,
            "underline": False,
        }
        if rpr is not None:
            fmt["bold"] = rpr.find("w:b", namespaces=ns) is not None
            fmt["italic"] = rpr.find("w:i", namespaces=ns) is not None
            fmt["small_caps"] = rpr.find("w:smallCaps", namespaces=ns) is not None
            fmt["underline"] = rpr.find("w:u", namespaces=ns) is not None

        runs.append((text, fmt))

    return runs


@app.post("/process-footnotes")
async def process_footnotes(data: dict):
    footnotes = data.get("footnotes", [])

    normalized = []
    for fn in footnotes:
        number = fn.get("number")
        text = fn.get("text", "")
        ooxml = fn.get("ooxml", "")

        src_runs = _runs_from_ooxml(ooxml)
        raw_text = "".join(t for t, _ in src_runs) if src_runs else text

        parts = [part.strip() for part in split_sources(raw_text) if part.strip()]

        # Use the original footnote_processor run-assignment logic when we have runs.
        runs_per_part = assign_source_runs(src_runs, parts) if src_runs else []

        sources = []
        for idx, part in enumerate(parts):
            part_runs = runs_per_part[idx] if idx < len(runs_per_part) else []

            # Serialize runs to JSON-friendly structure
            runs_json = [
                {
                    "text": t,
                    "bold": fmt.get("bold", False),
                    "italic": fmt.get("italic", False),
                    "small_caps": fmt.get("small_caps", False),
                    "underline": fmt.get("underline", False),
                }
                for (t, fmt) in part_runs
            ]

            classification = _classify_source(part_runs) if part_runs else []
            if isinstance(classification, str):
                rules = [classification]
            elif isinstance(classification, list):
                rules = classification
            else:
                rules = []

            file_label = _suggest_filename(number, part_runs, classification) if part_runs else ""

            sources.append(
                {
                    "text": part,
                    "runs": runs_json,
                    "rules": rules,
                    "warnings": [],
                    "fileLabel": file_label or "",
                }
            )

        if not sources and raw_text.strip():
            # Fallback: single source with either runs or plain text
            if src_runs:
                classification = _classify_source(src_runs)
                runs_json = [
                    {
                        "text": t,
                        "bold": fmt.get("bold", False),
                        "italic": fmt.get("italic", False),
                        "small_caps": fmt.get("small_caps", False),
                        "underline": fmt.get("underline", False),
                    }
                    for (t, fmt) in src_runs
                ]
                sources = [
                    {
                        "text": raw_text.strip(),
                        "runs": runs_json,
                        "rules": ([classification] if isinstance(classification, str) else classification) if classification else [],
                        "warnings": [],
                        "fileLabel": _suggest_filename(number, src_runs, classification) or "",
                    }
                ]
            else:
                sources = [
                    {
                        "text": raw_text.strip(),
                        "runs": [],
                        "rules": [],
                        "warnings": [],
                        "fileLabel": "",
                    }
                ]

        normalized.append(
            {
                "number": number,
                "context": raw_text[:200],
                "sources": sources,
            }
        )

    return {"footnotes": normalized}