from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
import asyncio
from concurrent.futures import ThreadPoolExecutor
import tempfile
import shutil
import re

from footnote_processor import _build_rows, split_sources, assign_source_runs, _classify_source, _suggest_filename, _get_body_sentence
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


def _extract_xml_roots_from_package_ooxml(package_ooxml: str):
    if not package_ooxml:
        return None, None
    try:
        pkg_root = etree.fromstring(package_ooxml.encode("utf-8"))
    except Exception:
        return None, None

    pkg_ns = {"pkg": "http://schemas.microsoft.com/office/2006/xmlPackage"}

    def _part_root(part_name: str):
        node = pkg_root.find(f".//pkg:part[@pkg:name='{part_name}']/pkg:xmlData/*", namespaces=pkg_ns)
        return node

    doc_root = _part_root("/word/document.xml")
    fn_root = _part_root("/word/footnotes.xml")
    return doc_root, fn_root


def _extract_body_contexts_from_package_ooxml(package_ooxml: str) -> dict:
    """Mirror extract_body_contexts(...) using packaged OOXML from Word.getOoxml()."""
    contexts = {}
    doc_root, fn_root = _extract_xml_roots_from_package_ooxml(package_ooxml)
    if doc_root is None or fn_root is None:
        return contexts

    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    w_ns = ns["w"]

    auto_ids = set()
    for fn_el in fn_root.findall("w:footnote", ns):
        raw_id = fn_el.get(f"{{{w_ns}}}id")
        if raw_id and fn_el.findall(".//w:footnoteRef", ns):
            try:
                auto_ids.add(int(raw_id))
            except Exception:
                continue

    display_num = 0
    xml_to_display = {}
    for ref in doc_root.findall(".//w:footnoteReference", ns):
        raw_id = ref.get(f"{{{w_ns}}}id")
        if not raw_id:
            continue
        try:
            xml_id = int(raw_id)
        except Exception:
            continue
        if xml_id in auto_ids:
            display_num += 1
            xml_to_display[xml_id] = display_num

    for para in doc_root.findall(".//w:p", ns):
        for ref in para.findall(".//w:footnoteReference", ns):
            raw_id = ref.get(f"{{{w_ns}}}id")
            if raw_id is None:
                continue
            try:
                xml_id = int(raw_id)
            except Exception:
                continue
            d = xml_to_display.get(xml_id)
            if d is not None and d not in contexts:
                try:
                    contexts[d] = _get_body_sentence(para, xml_id, auto_ids)
                except Exception:
                    contexts[d] = ""

    return contexts


def _process_single_footnote(fn: dict, body_contexts: dict) -> dict:
    number = fn.get("number")
    text = fn.get("text", "")
    ooxml = fn.get("ooxml", "")

    try:
        display_number = int(number)
    except Exception:
        display_number = number
    body_text = body_contexts.get(display_number, "")

    src_runs = _runs_from_ooxml(ooxml)
    raw_text = "".join(t for t, _ in src_runs) if src_runs else text

    parts = [part.strip() for part in split_sources(raw_text) if part.strip()]
    runs_per_part = assign_source_runs(src_runs, parts) if src_runs else []

    sources = []
    for idx, part in enumerate(parts):
        part_runs = runs_per_part[idx] if idx < len(runs_per_part) else []

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

    return {
        "number": number,
        "context": raw_text[:200],
        "bodyText": body_text,
        "sources": sources,
    }


_executor = ThreadPoolExecutor()


@app.post("/process-footnotes")
async def process_footnotes(data: dict):
    footnotes = data.get("footnotes", [])
    package_ooxml = data.get("documentOoxml", "")
    body_contexts = _extract_body_contexts_from_package_ooxml(package_ooxml)

    loop = asyncio.get_event_loop()
    tasks = [
        loop.run_in_executor(_executor, _process_single_footnote, fn, body_contexts)
        for fn in footnotes
    ]
    normalized = list(await asyncio.gather(*tasks))

    return {"footnotes": normalized}