#!/usr/bin/env python3
"""
Core footnote processing logic with rich text support.

Extracts footnotes from .docx XML, splits into individual citation sources
(same logic as footnote_breakdown.py), and generates a PDF that preserves
italic, bold, small caps, and underline formatting from the original document.

PDF table columns:
  FN #  |  Body Sentence Context  |  Source Text

The Body Sentence cell spans all source rows that belong to the same footnote.
"""

import io
import re
import zipfile

from lxml import etree

import xlsxwriter

from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    HRFlowable,
    Paragraph,
    SimpleDocTemplate,
    Table,
    TableStyle,
)

# ── Word XML namespace ─────────────────────────────────────────────────────────
W  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}

_QUOTES = "\u201c\u201d\u2018\u2019\"'"

_CITATION_SIGNALS      = frozenset({"See", "But", "Cf", "Id", "Accord", "Compare", "Contra"})
_CITATION_SIGNALS_LOWER = frozenset(s.lower() for s in _CITATION_SIGNALS)

# Regex to strip citation signal prefixes from the start of a source string
# before author/title extraction (e.g. "E.g., Oliver Kunzler…" → "Oliver Kunzler…")
_SIGNAL_PREFIX_RE = re.compile(
    r"^\s*(?:"
    r"See\s+generally"
    r"|See\s+also"
    r"|See\s+e\.g\."
    r"|E\.g\."
    r"|See"
    r"|But\s+see"
    r"|Cf\."
    r"|Accord"
    r"|Compare"
    r"|Contra"
    r")"
    r"[,.]?\s*",
    re.IGNORECASE,
)

_COMMENTARY_STARTERS = frozenset({
    # Articles / determiners
    "The", "This", "These", "A", "An", "There", "That", "Those",
    # Prepositions / conjunctions commonly starting commentary
    "In", "At", "As", "Although", "While", "However", "Moreover",
    "Furthermore", "Additionally", "Note", "Notably", "Indeed",
    "Section", "Subsection", "Harmonizing", "Under", "Here",
    "Because", "Since", "When", "Where", "If", "It",
    # Discourse connectors / transitional phrases
    "For", "Also", "Despite", "Yet", "Thus", "By", "With", "Without",
    "Among", "Both", "Each", "Such", "Other", "Many", "Most",
    "Several", "Some", "Subsequent", "Only", "Two", "Three",
})


# ── Run-level formatting extraction ───────────────────────────────────────────

def _get_rpr(run_el):
    """Return formatting dict for a w:r element."""
    rpr   = run_el.find("w:rPr", NS)
    props = {"bold": False, "italic": False, "small_caps": False, "underline": False}
    if rpr is None:
        return props

    def _on(tag):
        el = rpr.find(tag, NS)
        if el is None:
            return False
        val = el.get(f"{{{W}}}val", "true")
        return val.lower() not in ("false", "0", "off")

    props["bold"]       = _on("w:b")
    props["italic"]     = _on("w:i")
    props["small_caps"] = _on("w:smallCaps")

    u = rpr.find("w:u", NS)
    if u is not None:
        uval = u.get(f"{{{W}}}val", "")
        props["underline"] = uval.lower() not in ("none", "false", "0", "")

    return props


# ── Body sentence extraction (ported from sanity_check.py) ────────────────────

def _is_abbreviation(word: str) -> bool:
    """True if word looks like a citation abbreviation rather than a sentence end."""
    w = word.strip(".,;:()" + _QUOTES)
    return len(w) <= 2 or "." in w


def _last_sentence(text: str) -> str:
    """Return the last complete sentence in text."""
    start = 0
    for m in re.finditer(r"[.!?]", text):
        pos  = m.start()
        rest = text[pos + 1:].lstrip(" " + _QUOTES)
        if rest and not rest[0].isupper():
            continue
        word_m = re.search(r"(\S+)$", text[:pos])
        if not word_m or _is_abbreviation(word_m.group(1)):
            continue
        next_start = pos + 1 + (len(text[pos + 1:]) - len(text[pos + 1:].lstrip(" " + _QUOTES)))
        if next_start < len(text):
            start = next_start
    return text[start:]


def _get_body_sentence(para, target_xml_id: int, auto_ids: set) -> str:
    """
    Walk the paragraph in document order.  Collect text up to the target
    footnote reference, resetting at each other auto-numbered reference so
    each footnote gets only the text since the previous footnote.
    """
    text_before: list[str] = []
    for el in para.iter():
        tag = el.tag.split("}")[-1] if "}" in el.tag else el.tag
        if tag == "footnoteReference":
            ref_id = el.get(f"{{{W}}}id")
            if ref_id and int(ref_id) == target_xml_id:
                break
            if ref_id and int(ref_id) in auto_ids:
                text_before = []
        elif tag == "t":
            text_before.append(el.text or "")
    return _last_sentence("".join(text_before)).strip()


def extract_body_contexts(docx_path: str, start_fn: int, end_fn: int) -> dict:
    """Return {display_number: body_sentence_str} for footnotes in range."""
    contexts = {}

    with zipfile.ZipFile(docx_path) as z:
        fn_root  = etree.fromstring(z.read("word/footnotes.xml"))
        doc_root = etree.fromstring(z.read("word/document.xml"))

        auto_ids = set()
        for fn_el in fn_root.findall("w:footnote", NS):
            raw_id = fn_el.get(f"{{{W}}}id")
            if raw_id and fn_el.findall(".//w:footnoteRef", NS):
                auto_ids.add(int(raw_id))

        # Build xml_id → display_number map
        display_num   = 0
        xml_to_display = {}
        for ref in doc_root.findall(".//w:footnoteReference", NS):
            raw_id = ref.get(f"{{{W}}}id")
            if raw_id and int(raw_id) in auto_ids:
                display_num += 1
                xml_to_display[int(raw_id)] = display_num

        # Walk paragraphs to extract body sentence per footnote
        for para in doc_root.findall(".//w:p", NS):
            for ref in para.findall(".//w:footnoteReference", NS):
                raw_id = ref.get(f"{{{W}}}id")
                if raw_id is None:
                    continue
                xml_id = int(raw_id)
                d = xml_to_display.get(xml_id)
                if d is not None and start_fn <= d <= end_fn and d not in contexts:
                    contexts[d] = _get_body_sentence(para, xml_id, auto_ids)

    return contexts


# ── Docx footnote run extraction ───────────────────────────────────────────────

def extract_footnote_runs(docx_path: str, start_fn: int, end_fn: int) -> dict:
    """
    Returns {display_number: [(text_str, props_dict), ...]}
    Preserves run-level formatting (bold, italic, smallCaps, underline).
    """
    result = {}

    with zipfile.ZipFile(docx_path) as z:
        fn_root  = etree.fromstring(z.read("word/footnotes.xml"))
        doc_root = etree.fromstring(z.read("word/document.xml"))

        auto_ids   = set()
        runs_by_id = {}

        for fn_el in fn_root.findall("w:footnote", NS):
            raw_id = fn_el.get(f"{{{W}}}id")
            if raw_id is None:
                continue
            xml_id = int(raw_id)
            if not fn_el.findall(".//w:footnoteRef", NS):
                continue
            auto_ids.add(xml_id)

            fn_runs    = []
            paragraphs = fn_el.findall(".//w:p", NS)
            for p_idx, para in enumerate(paragraphs):
                for run in para.findall(".//w:r", NS):
                    if run.find("w:footnoteRef", NS) is not None:
                        continue
                    text = "".join((t.text or "") for t in run.findall("w:t", NS))
                    if text:
                        fn_runs.append((text, _get_rpr(run)))
                if p_idx < len(paragraphs) - 1 and fn_runs:
                    fn_runs.append((" ", {"bold": False, "italic": False,
                                          "small_caps": False, "underline": False}))

            runs_by_id[xml_id] = fn_runs

        display_num = 0
        for ref in doc_root.findall(".//w:footnoteReference", NS):
            raw_id = ref.get(f"{{{W}}}id")
            if raw_id and int(raw_id) in auto_ids:
                display_num += 1
                if start_fn <= display_num <= end_fn:
                    result[display_num] = runs_by_id.get(int(raw_id), [])

    return result


# ── Citation splitting ─────────────────────────────────────────────────────────

# Compiled once: matches any bibliographic/citation marker inside a segment.
# If a segment's first word is not a known citation signal or commentary
# starter, we require at least one of these markers before treating it as
# a real source — this rejects pure commentary sentences (e.g. "Courts have
# recognized...") whose opening word just isn't enumerated in
# _COMMENTARY_STARTERS.
_SOURCE_MARKER_RE = re.compile(
    r"(\(\d{4}\)"                   # year in parens: (2020)
    r"|\bsupra\b|\binfra\b"         # supra / infra note references
    r"|https?://"                   # URL
    r"|[§¶]"                        # statutory section / paragraph
    r"|\bat\s+\d+\b"                # page citation: at 5
    r"|\b\d+\s+[A-Z][a-z]{0,4}\."  # volume + reporter abbreviation: 50 L.
    r"|\bF\.[234](?:d|th)\b"        # federal reporters: F.2d F.3d F.4th
    r"|\bU\.S\.\b"                  # U.S. reporter
    r"|\bS\.\s*Ct\."                # S. Ct.
    r"|\bL\.\s*(?:Rev|J)\."         # L. Rev. / L.J.
    r"|\b(?:ed|eds|vol|rev|no)\."   # common bibliographic abbreviations
    r")",
    re.IGNORECASE,
)


def _has_source_markers(text):
    """Return True if text contains at least one bibliographic/citation marker."""
    return bool(_SOURCE_MARKER_RE.search(text))


def _is_source(text):
    m = re.match(r"\S+", text.lstrip(_QUOTES + " "))
    if not m:
        return False
    fw = m.group(0).rstrip(".,;:-()" + _QUOTES)
    if fw in _CITATION_SIGNALS:
        return True
    if fw and (fw[0].isdigit() or fw.startswith("§")):
        return True
    if fw in _COMMENTARY_STARTERS:
        return False
    # First word is neither a recognized citation signal nor a known commentary
    # starter.  Only accept as a source if it contains actual bibliographic
    # markers (year, reporter, supra/infra, URL, etc.).  This rejects prose
    # commentary sentences that happen to begin with a word not in the
    # starter list (e.g. "Courts have recognized…", "Respondents argue…").
    return _has_source_markers(text)


# ── Bluebook rule classification ───────────────────────────────────────────────

def _classify_source(runs):
    """
    Classify a citation source using plain text + italic/formatting cues.
    Returns a string listing 2+ Bluebook rules to consult, or a NOTE for
    ambiguous/edge-case citations.

    Classification priority:
      1. Id.           → Rule 4.1
      2. supra/infra   → Rule 4.2
      3. Statute (§, U.S.C., Pub. L.)  → Rule 12
      4. Constitution  → Rule 11
      5. Restatement / Model Code      → Rule 12.9
      6. Regulation (C.F.R., Fed. Reg.) → Rule 14
      7. Case (italic "v.")            → Rule 10
      8. URL           → Rule 18
      9. Journal article (vol + L. Rev./J.) → Rule 16
     10. Book (italic title + year/ed.) → Rule 15
     11. Year present but type unclear → NOTE
     12. Fully ambiguous              → NOTE
    """
    plain       = "".join(t for t, _ in runs).strip()
    italic_text = " ".join(t for t, p in runs if p.get("italic")).strip()

    # 1. Id. short form
    if re.match(r"^Id\b", plain, re.IGNORECASE):
        return "Rule 4.1 (Id.); Rule for underlying source type"

    # 2. supra / infra
    if re.search(r"\bsupra\b|\binfra\b", plain, re.IGNORECASE):
        return "Rule 4.2 (supra/infra); Rule for underlying source type"

    # 3. Statute
    if re.search(r"[§¶]|\bU\.S\.C\.|\bPub\.\s*L\.", plain):
        return "Rule 12 (Statutes); Rule 12.3–12.4 (Codified/Session Laws)"

    # 4. Constitution
    if re.search(r"\bConst\b|\bConstitution\b", plain, re.IGNORECASE):
        return "Rule 11 (Constitutions); Rule 11.1 (U.S. Constitution)"

    # 5. Restatement / Model Code / Uniform Act
    if re.search(r"\bRestatement\b|\bModel\s+(?:Code|Rules|Penal)\b|\bUniform\s+\w+\s+Act\b",
                 plain, re.IGNORECASE):
        return "Rule 12.9 (Restatements & Model Codes); Rule 4.2 (supra short form)"

    # 6. Administrative regulation / executive material
    if re.search(r"\bC\.F\.R\.|\bFed\.\s*Reg\.|\bExec\.\s*Order\b", plain):
        return "Rule 14 (Administrative Materials); Rule 14.2 (Rules & Regulations)"

    # 7. Case: italic run containing " v. "
    if re.search(r"\bv\.\s", italic_text):
        return "Rule 10 (Cases); Rule 10.2 (Case Names); Rule 10.4 (Reporters)"

    # 8. Internet / URL
    if re.search(r"https?://", plain, re.IGNORECASE):
        return "Rule 18 (Internet Sources); Rule 18.2 (Direct Internet Citations)"

    # 9. Journal / law review article: volume + abbreviated journal name + page
    if (re.search(r"\b\d+\s+[A-Z][a-z]{0,4}\.\s*(?:Rev|J|L)\b", plain)
            or re.search(r"\bL\.\s*(?:Rev|J)\.", plain)):
        # Check for student note/comment signals
        student = re.search(r"\bNote\b|\bComment\b|\bRecent\s+Case\b", plain)
        r2 = "Rule 16.4 (Student-Written Materials)" if student else "Rule 16.3 (Author & Title)"
        return f"Rule 16 (Periodical Materials); {r2}"

    # 10. Book / non-periodical: italic title without "v.", plus year or ed.
    if italic_text and not re.search(r"\bv\.\s", italic_text):
        if re.search(r"\(\d{4}\)", plain) or re.search(r"\beds?\b", plain, re.IGNORECASE):
            return "Rule 15 (Books & Non-Periodicals); Rule 15.1–15.4 (Author, Title, Edition, Year)"

    # 11. Has a year but type is unclear
    if re.search(r"\(\d{4}\)", plain):
        if italic_text:
            return ("NOTE: type unclear (Book or Article?) — "
                    "Rule 15 (Books); Rule 16 (Articles); verify italic title structure")
        return ("NOTE: type unclear (Case or Book?) — "
                "Rule 10 (Cases); Rule 15 (Books); verify citation structure")

    # 12. Fully ambiguous / edge case
    return ("NOTE: citation type unclear — "
            "Rule 10 (Cases), Rule 15 (Books), or Rule 16 (Articles); verify manually")


def _is_abbrev(word):
    w = word.strip(".,;:()" + _QUOTES)
    # len <= 4 catches common 4-char legal abbreviations: Conf, Corp, Auth, Comm, Supp, Univ …
    return len(w) <= 2 or "." in w or "'" in w or len(w) <= 4


def _citation_segment_complete(text, start, end):
    """
    Return True if the accumulated citation text from start..end already
    looks like a complete citation entry — meaning a real sentence boundary
    could follow.  Incomplete segments are still mid-citation (e.g. the
    publisher date and URL haven't appeared yet), so a period within them
    is likely an abbreviation period, not a sentence-ending period.

    Completion markers:
      • a year in parentheses  (2021)
      • any closing parenthetical ending the segment
      • a URL (http)
    """
    seg = text[start:end]
    return bool(
        re.search(r'\(\d{4}\)', seg)   # year in parens
        or seg.rstrip().endswith(')')  # closing parenthetical
        or 'http' in seg.lower()       # URL present
    )


def split_sources(raw_text):
    """Split footnote plain text into individual citation source strings."""
    text = raw_text.replace("\xa0", " ")
    text = re.sub(r"\s*\bDOC\d+\b\s*", " ", text)
    text = re.sub(r"\s+", " ", text).strip()

    sources    = []
    prev_start = 0

    for m in re.finditer(r"[.!?;]", text):
        pos  = m.start()
        ch   = m.group(0)
        rest = text[pos + 1:].lstrip(" " + _QUOTES)

        if not rest:
            if ch != ";":
                seg = text[prev_start:pos + 1].strip()
                if seg:
                    sources.append(seg)
                prev_start = len(text)
            break

        if ch == ";":
            nw      = re.match(r"(\S+)", rest)
            nw_bare = nw.group(1).rstrip(".,;:-()" + _QUOTES).lower() if nw else ""
            # Never split on lowercase-start unless it's a citation signal
            if not rest[0].isupper() and nw_bare not in _CITATION_SIGNALS_LOWER:
                continue
            # Uppercase after ";": a citation signal always splits.
            # Otherwise, check the token immediately before the ";" — if it is a
            # plain alphabetic word (no digits, not closing a parenthetical), the
            # semicolon is most likely a title/subtitle separator, not a citation
            # boundary (e.g. "Stress in America; Generation Z" or "the Job;
            # Practicing Attorneys").  A genuine citation boundary has a closing
            # parenthetical "(…)" or a number/page range before the ";".
            if nw_bare not in _CITATION_SIGNALS_LOWER:
                wb_m = re.search(r"(\S+)$", text[:pos])
                if wb_m:
                    wb_raw  = wb_m.group(1)
                    wb_core = wb_raw.strip(".,;:-()" + _QUOTES)
                    # Plain word before ";" with no closing paren → mid-title
                    if (re.match(r"^[A-Za-z'\u2019\-]+$", wb_core)
                            and not wb_raw.rstrip(".,;").endswith((")", "]"))):
                        continue
            seg = text[prev_start:pos].strip()
            if seg:
                sources.append(seg)
            skip = len(text[pos + 1:]) - len(text[pos + 1:].lstrip(" " + _QUOTES))
            prev_start = pos + 1 + skip
            continue

        if not rest[0].isupper():
            continue

        word_m      = re.search(r"(\S+)$", text[:pos])
        next_word_m = re.match(r"(\S+)", rest)
        if not word_m:
            continue

        word      = word_m.group(1)
        next_word = next_word_m.group(1) if next_word_m else ""
        next_bare = next_word.rstrip(".,;:()" + _QUOTES)

        if next_bare in _CITATION_SIGNALS:
            pass  # P1: citation signal → always split
        elif next_word.endswith("."):
            continue  # P2: next word ends with "." → abbreviation chain, never
                      # split (e.g. "U.S.C.A.", "Jr.", "Corp.").  Must come
                      # before P1b so single-letter starters like "A." in
                      # "U.S.C.A." are not mistaken for the article "A".
        elif next_bare in _COMMENTARY_STARTERS:
            pass  # P1b: commentary starter → split so _is_source can filter it
                  # (overrides the P4 abbreviation guard, e.g. "at 4. The …")
        elif word.endswith((")", "]")):
            pass  # P3: parenthetical just closed → split
        elif _is_abbrev(word):
            continue  # P4: abbreviation before period → no split
        elif word[0].isupper() and not _citation_segment_complete(text, prev_start, pos):
            continue  # P4b: title-case word but citation not yet complete —
                      # period is likely inside an abbreviation ("Assoc.", "Psychol.")
                      # rather than a genuine sentence boundary.
                      # Safe because P1/P1b already fire for any citation signal
                      # or commentary starter that follows, so those splits still happen.
        else:
            pass  # P5: ordinary word → split

        seg = text[prev_start:pos + 1].strip()
        if seg:
            sources.append(seg)
        skip = len(text[pos + 1:]) - len(text[pos + 1:].lstrip(" " + _QUOTES))
        prev_start = pos + 1 + skip

    remaining = text[prev_start:].strip()
    if remaining:
        sources.append(remaining)

    return [s for s in sources if s and _is_source(s)]


# ── Map source segments back to their formatted runs ──────────────────────────

def _normalize(text):
    return re.sub(r"\s+", " ", text.replace("\xa0", " ")).strip()


def assign_source_runs(all_runs, sources):
    """
    Map each source string to the corresponding (text, props) runs.
    Returns a list of run-groups, one per source, preserving formatting.
    """
    _plain = {"bold": False, "italic": False, "small_caps": False, "underline": False}

    if not all_runs or not sources:
        return [[(s, _plain)] for s in sources]

    # Build collapsed char list with per-char run-index tracking
    collapsed = []
    in_space  = False
    for run_idx, (text, _) in enumerate(all_runs):
        for ch in text.replace("\xa0", " "):
            if ch in " \t\r\n":
                if not in_space:
                    collapsed.append((" ", run_idx))
                    in_space = True
            else:
                collapsed.append((ch, run_idx))
                in_space = False

    norm_text    = "".join(c for c, _ in collapsed)
    search_start = 0
    result       = []

    for source in sources:
        norm_src = _normalize(source)
        idx = norm_text.find(norm_src, search_start)

        if idx == -1:
            result.append([(source, _plain)])
            continue

        end_idx = idx + len(norm_src)

        current_run_idx = None
        current_text    = ""
        src_runs        = []

        for c, ri in collapsed[idx:end_idx]:
            if ri != current_run_idx:
                if current_run_idx is not None and current_text:
                    _, p = all_runs[current_run_idx]
                    src_runs.append((current_text, p))
                current_run_idx = ri
                current_text    = c
            else:
                current_text += c

        if current_run_idx is not None and current_text:
            _, p = all_runs[current_run_idx]
            src_runs.append((current_text, p))

        result.append(src_runs if src_runs else [(source, _plain)])
        search_start = end_idx

    return result


# ── ReportLab markup ───────────────────────────────────────────────────────────

def runs_to_markup(runs):
    """Convert (text, props) runs to ReportLab paragraph XML markup."""
    parts = []
    for text, props in runs:
        if not text:
            continue
        s = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

        if props.get("small_caps"):
            s = f'<font size="7">{s.upper()}</font>'
        if props.get("underline"):
            s = f"<u>{s}</u>"
        if props.get("bold") and props.get("italic"):
            s = f"<b><i>{s}</i></b>"
        elif props.get("bold"):
            s = f"<b>{s}</b>"
        elif props.get("italic"):
            s = f"<i>{s}</i>"

        parts.append(s)
    return "".join(parts)


# ── PDF generation ─────────────────────────────────────────────────────────────

def _style(name, **kw):
    base = {"fontName": "Times-Roman", "fontSize": 9, "leading": 12, "wordWrap": "LTR"}
    base.update(kw)
    return ParagraphStyle(name, **base)


def build_pdf(rows, docx_name, start_fn, end_fn, source_count, commentary_count):
    """
    Build and return a PDF as bytes.

    rows: [(fn_num_str, body_sentence_str, [(text, props), ...]), ...]

    When multiple consecutive rows share the same fn_num, the Body Sentence
    cell is rendered only in the first row and spans all rows for that footnote.
    """
    buf    = io.BytesIO()
    page_w, _ = landscape(letter)

    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(letter),
        leftMargin=0.5 * inch,
        rightMargin=0.5 * inch,
        topMargin=0.65 * inch,
        bottomMargin=0.65 * inch,
    )

    normal     = _style("normal")
    col_header = _style("colh", fontName="Times-Bold")
    h1         = _style("h1",  fontName="Times-Bold", fontSize=13, leading=16)
    sub        = _style("sub", fontSize=8.5, leading=11,
                         textColor=colors.HexColor("#555555"))

    avail_w    = page_w - doc.leftMargin - doc.rightMargin
    fn_col_w   = 0.45 * inch
    body_col_w = 2.10 * inch
    bb_col_w   = 1.50 * inch
    label_col_w = 1.80 * inch
    src_col_w  = avail_w - fn_col_w - body_col_w - bb_col_w - label_col_w

    root_style = _style("root", fontSize=7.5, leading=10,
                        textColor=colors.HexColor("#666666"),
                        fontName="Times-Italic")

    # ── Build table data and collect SPAN commands ────────────────────────────
    table_data    = [[
        Paragraph("FN #",                  col_header),
        Paragraph("Body Sentence Context", col_header),
        Paragraph("Source Text",           col_header),
        Paragraph("Bluebook Rules",        col_header),
        Paragraph("File Label",            col_header),
    ]]
    span_commands = []
    row_idx       = 1  # 0 is the header

    i = 0
    while i < len(rows):
        fn_num = rows[i][0]

        # Count consecutive rows with the same footnote number
        j = i + 1
        while j < len(rows) and rows[j][0] == fn_num:
            j += 1
        count = j - i  # number of source rows for this footnote
        body  = rows[i][1]

        for k in range(count):
            _, _, src_runs, bb_rules, id_root_str, file_label, _, label_span = rows[i + k]
            # Body sentence and FN # only appear in the first row of the group
            body_cell = Paragraph(body, normal) if k == 0 else Paragraph("", normal)
            fn_cell   = Paragraph(fn_num, normal) if k == 0 else Paragraph("", normal)

            # Source cell: append Id. root note when present
            src_markup = runs_to_markup(src_runs)
            if id_root_str:
                escaped = (id_root_str
                           .replace("&", "&amp;")
                           .replace("<", "&lt;")
                           .replace(">", "&gt;"))
                src_markup += (f'<br/><font size="7.5" color="#666666">'
                               f'<i>{escaped}</i></font>')

            # File Label cell: only first row of a chain span gets content
            label_cell = (Paragraph(file_label, normal)
                          if label_span != 0
                          else Paragraph("", normal))

            table_data.append([
                fn_cell,
                body_cell,
                Paragraph(src_markup, normal),
                Paragraph(bb_rules, normal),
                label_cell,
            ])

            # Emit SPAN command for File Label chain coalescing
            if label_span > 1:
                span_commands.append(
                    ("SPAN", (4, row_idx + k), (4, row_idx + k + label_span - 1))
                )

        # SPAN the FN # and Body Sentence cells across all rows in the group
        if count > 1:
            span_commands += [
                ("SPAN", (0, row_idx), (0, row_idx + count - 1)),
                ("SPAN", (1, row_idx), (1, row_idx + count - 1)),
            ]

        row_idx += count
        i = j

    # ── Style ─────────────────────────────────────────────────────────────────
    base_style = [
        ("TEXTCOLOR",     (0, 0), (-1, 0), colors.black),
        ("GRID",          (0, 0), (-1, -1), 0.5, colors.HexColor("#aaaaaa")),
        ("BOX",           (0, 0), (-1, -1), 0.7, colors.black),
        ("ALIGN",         (0, 0), (0, -1), "CENTER"),
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 5),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 5),
        ("TOPPADDING",    (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1),
         [colors.white, colors.HexColor("#f5f5f5")]),
    ]

    tbl = Table(
        table_data,
        colWidths=[fn_col_w, body_col_w, src_col_w, bb_col_w, label_col_w],
        repeatRows=1,
    )
    tbl.setStyle(TableStyle(base_style + span_commands))

    elements = [
        Paragraph("SCU Law Review — Footnote Breakdown", h1),
        Paragraph(f"FN {start_fn}–{end_fn} &nbsp;|&nbsp; Source: {docx_name}", sub),
        Paragraph(
            f"Total source entries: {source_count}"
            f" &nbsp;|&nbsp; Commentary-only footnotes: {commentary_count}", sub
        ),
        HRFlowable(width="100%", thickness=1, color=colors.black,
                   spaceBefore=4, spaceAfter=8),
        tbl,
    ]

    doc.build(elements)
    buf.seek(0)
    return buf.read()


# ── Id. root resolution & file-label helpers ──────────────────────────────────

def _build_id_root_map(runs_map):
    """
    Walk footnotes in ascending order, tracking the most recently cited
    non-short-form authority.

    Both Id. and supra are treated as back-references: they receive a pointer
    to last_root but do NOT update it (so a subsequent Id. after a supra still
    resolves to the original source, not the supra).

    Returns {fn_num: [root_or_None, ...]} where the list index matches the
    source index within that footnote (as produced by split_sources).
    """
    last_root = None   # (fn_num, plain_source_text)
    result    = {}

    for fn_num in sorted(runs_map):
        runs = runs_map[fn_num]
        if not runs:
            result[fn_num] = []
            continue

        plain    = "".join(t for t, _ in runs)
        sources  = split_sources(plain)
        fn_roots = []

        for src in sources:
            is_short = (re.match(r"^Id\b", src.strip(), re.IGNORECASE)
                        or re.search(r"\bsupra\b", src, re.IGNORECASE))
            if is_short:
                fn_roots.append(last_root)   # None if chain predates the buffer
            else:
                last_root = (fn_num, src)
                fn_roots.append(None)

        result[fn_num] = fn_roots

    return result


def _fs_safe(s, maxlen=55):
    """Return a filesystem-safe version of s, truncated to maxlen."""
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", s)
    s = re.sub(r"\s+", " ", s).strip().rstrip(".")
    return s[:maxlen]


def _label_from_plain(plain_text):
    """
    Extract an author_last + title (or case + reporter) label from plain-text
    citation.  Used when we only have plain text for the root source (Id./supra
    resolution) and can't rely on italic run metadata.
    """
    plain = _SIGNAL_PREFIX_RE.sub("", plain_text.strip()).strip()

    # Case: has "v. CapitalLetter" pattern
    if re.search(r"\bv\.\s+[A-Z]", plain):
        case_m = re.match(r"([^\d]+?)(?=,?\s*\d)", plain)
        case_part = _fs_safe(case_m.group(1).strip() if case_m else plain[:50], 50)
        rep_m = re.search(r"\b(\d+\s+\S+\.?\s*\d+)", plain)
        reporter  = _fs_safe(rep_m.group(1)) if rep_m else ""
        return "_".join(p for p in [case_part, reporter] if p)

    # Statute
    if re.search(r"[§¶]|\bU\.S\.C", plain):
        stat = re.sub(r"\(\d{4}\).*$", "", plain).strip()
        return _fs_safe(stat, 55)

    # General: try "Author, Title, Source" split
    parts = plain.split(",", 2)
    if len(parts) >= 2 and 1 <= len(parts[0].split()) <= 3:
        name_words = parts[0].strip().split()
        author_last = _fs_safe(name_words[-1], 20) if name_words else ""
        rest = ",".join(parts[1:]).strip()
        rest = re.sub(r"\s*\(\d{4}\).*$", "", rest)
        title = _fs_safe(rest[:50])
        return f"{author_last}_{title}" if title else author_last

    return _fs_safe(plain[:60])

def _author_section_from_plain(author_raw):
    """Return `Last` or `Last1 & Last2` from a plain-text author block."""
    author_raw = _SIGNAL_PREFIX_RE.sub("", author_raw.strip()).strip().rstrip(",")
    if not author_raw:
        return ""

    norm = re.sub(r"\s+(?:and|&)\s+", " & ", author_raw, flags=re.IGNORECASE)
    parts = [p.strip() for p in norm.split(" & ") if p.strip()]

    last_names = []
    for part in parts:
        cleaned = re.sub(r"\s+", " ", part.strip().strip(","))
        if not cleaned:
            continue
        words = cleaned.split()
        last_names.append(_fs_safe(words[-1], 20))

    return " & ".join(last_names)

def _internet_author_and_title(plain):
    """
    Best-effort parser for internet/article-like citations of the form:
      Title, AUTHOR NAME (n.d./year), URL ...
    Returns (author_section, title).
    """
    before_url = re.split(r"https?://", plain, maxsplit=1, flags=re.IGNORECASE)[0].strip().rstrip(",")
    m = re.match(
        r"^(?P<title>.+?),\s*(?P<author>[^,]+?)\s*\((?:n\.d\.|\d{4})\)\s*$",
        before_url,
        re.IGNORECASE,
    )
    if m:
        title = _fs_safe(m.group("title").strip(), 80)
        author_section = _author_section_from_plain(m.group("author"))
        return author_section, title
    return "", ""

def _suggest_filename(fn_num, src_runs, classification, root_info=None, supra_ref_label=None):
    """
    Suggest a file-label string for a source entry.

    Format varies by citation type:
      Articles / Books  →  FN#_AuthorLastName_Title
      Cases             →  FN#_CaseName_Reporter   (punctuation preserved)
      Statutes          →  FN#_StatuteRef           (§, dots, etc. preserved)
      Internet          →  FN#_AuthorLastName_Title (if article-like)
                           FN#_PageTitle            (if bare URL/page)
      Id. / supra       →  FN{root_fn}-{fn_num}_RootAuthor_RootTitle
                           (chains resolve back to the ultimate root footnote)
    """
    plain       = "".join(t for t, _ in src_runs).strip()
    italic_text = "".join(t for t, p in src_runs if p.get("italic")).strip()
    prefix      = f"FN{fn_num}"

    # ── Id. ───────────────────────────────────────────────────────────────
    if re.match(r"^Id\b", plain, re.IGNORECASE):
        if root_info:
            root_fn, root_text = root_info
            return f"FN{root_fn}-{fn_num}_{_label_from_plain(root_text)}"
        return f"{prefix}_Id"

    # ── supra ─────────────────────────────────────────────────────────────
    # supra_ref_label: pre-computed canonical label of the referenced footnote
    # (passed in from _build_rows after resolving "supra note N")
    if re.search(r"\bsupra\b", plain, re.IGNORECASE):
        note_m = re.search(r"\bsupra\s+note\s+(\d+)\b", plain, re.IGNORECASE)
        ref_fn = note_m.group(1) if note_m else None
        if supra_ref_label and ref_fn:
            rest = re.sub(rf"^FN{ref_fn}_?", "", supra_ref_label)
            return f"FN{ref_fn}-{fn_num}_{rest}" if rest else f"FN{ref_fn}-{fn_num}_supra"
        if root_info:
            root_fn, root_text = root_info
            return f"FN{root_fn}-{fn_num}_{_label_from_plain(root_text)}"
        if ref_fn:
            return f"FN{ref_fn}-{fn_num}_supra"
        return f"{prefix}_supra"

    # ── Case (Rule 10) ────────────────────────────────────────────────────
    if "Rule 10" in classification:
        # Preserve case name punctuation (v., commas, etc.) — only strip
        # trailing page number references
        if re.search(r"\bv\.\s", italic_text):
            cn = re.sub(r",\s*\d+.*$", "", italic_text).strip()
            case_name = _fs_safe(cn, 50)
        else:
            case_name = _fs_safe(italic_text[:45] or plain[:45], 50)
        # Reporter: preserve abbreviation dots (F.3d, U.S., S. Ct., etc.)
        rep_m = re.search(r"\b(\d+\s+\S+\.?\s*\d+)", plain)
        reporter = _fs_safe(rep_m.group(1)) if rep_m else ""
        return "_".join(p for p in [prefix, case_name, reporter] if p)

    # ── Statute (Rule 12) ─────────────────────────────────────────────────
    # Preserve § symbol, U.S.C.A. dots, etc. — only strip the trailing year
    if "Rule 12" in classification:
        # Keep only the main statute reference and strip subsection / commentary parentheticals.
        stat_m = re.search(r"(\d+\s+U\.S\.C\.?\s+§+\s*\d+)", plain)
        if stat_m:
            stat = stat_m.group(1)
        else:
            stat = re.sub(r"\s*\([^)]*\)", "", plain).strip()
        stat = _fs_safe(stat, 55)
        return f"{prefix}_{stat}" if stat else prefix

    # ── Shared helper: extract author last name + italic title from runs ──
    def _author_and_title():
        author_buf = []
        for text, props in src_runs:
            if props.get("italic"):
                break
            author_buf.append(text)
        author_raw = "".join(author_buf).strip().rstrip(",").strip()
        # Strip any leading citation signal (e.g. "E.g.,", "See generally,")
        author_raw = _SIGNAL_PREFIX_RE.sub("", author_raw).strip().rstrip(",").strip()
        # Use last word as last name (natural order: "Oliver Kunzler" → "Kunzler")
        author_last = _author_section_from_plain(author_raw)
        title_raw = italic_text.split(",")[0] if italic_text else ""
        title     = _fs_safe(title_raw, 50)
        return author_last, title

    # ── Journal article (Rule 16) or Book (Rule 15) ───────────────────────
    if "Rule 16" in classification or "Rule 15" in classification:
        author_last, title = _author_and_title()
        return "_".join(p for p in [prefix, author_last, title] if p)

    # ── Internet / URL (Rule 18) ──────────────────────────────────────────
    # Many online articles have a clear author + italic title — detect that
    # structure and treat them identically to journal articles.
    if "Rule 18" in classification:
        #First, try article-like internet citations where the title precedes the author.
        author_section, title = _internet_author_and_title(plain)
        if author_section or title:
            return "_".join(p for p in [prefix, author_section, title] if p)

        # Next, handle sources that still expose a conventional author + italic title structure.
        has_non_italic = any(t.strip() and not p.get("italic") for t, p in src_runs)
        if has_non_italic and italic_text:
            author_last, title = _author_and_title()
            return "_".join(p for p in [prefix, author_last, title] if p)

        # Bare URL or no clear structure.
        title_text = re.sub(r"https?://\S+", "", plain).strip()
        title_text = re.sub(r"\(last visited[^)]*\)", "", title_text, flags=re.IGNORECASE).strip().rstrip(",")
        title_text = re.sub(r"^(?:See|Cf|But\s+see|Accord)[,.]?\s+",
                            "", title_text, flags=re.IGNORECASE)
        title = _fs_safe(title_text, 80) or "Online Source"
        return f"{prefix}_{title}"

    # ── Administrative / regulatory (Rule 14) ────────────────────────────
    if "Rule 14" in classification:
        title = _fs_safe(italic_text or plain[:55], 55)
        return f"{prefix}_{title}" if title else prefix

    # ── Fallback ──────────────────────────────────────────────────────────
    title = _fs_safe(italic_text or plain[:55], 55)
    return f"{prefix}_{title}" if title else prefix


# ── Excel generation ───────────────────────────────────────────────────────────

def build_xlsx(rows, docx_name, start_fn, end_fn, source_count, commentary_count):
    """
    Build and return an Excel workbook as bytes using xlsxwriter.

    xlsxwriter's write_rich_string() stores rich text in the shared-strings
    table (not as inline strings), avoiding the Excel repair warnings that
    openpyxl's CellRichText produces when cell-level styles are also applied.
    """
    buf = io.BytesIO()
    wb  = xlsxwriter.Workbook(buf, {"in_memory": True})
    ws  = wb.add_worksheet(f"FN {start_fn}-{end_fn}")

    # ── Format definitions ────────────────────────────────────────────────────
    _cell = dict(
        font_name="Times New Roman", font_size=10,
        border=1, border_color="#AAAAAA",
        text_wrap=True, valign="top",
    )
    fmt_header = wb.add_format({**_cell, "bold": True,  "align": "center"})
    fmt_normal = wb.add_format(_cell)
    fmt_center = wb.add_format({**_cell, "align": "center"})

    # Character-only formats (no borders/wrap) — used as per-run format args
    # inside write_rich_string().
    _char_cache  = {}

    def _char_fmt(bold, italic, underline, small_caps):
        key = (bold, italic, underline, small_caps)
        if key not in _char_cache:
            _char_cache[key] = wb.add_format({
                "font_name": "Times New Roman",
                "font_size": 7 if small_caps else 10,
                "bold":      bold,
                "italic":    italic,
                "underline": 1 if underline else 0,
            })
        return _char_cache[key]

    # Combined char + cell formats — used when a source has only one run
    # (write_rich_string requires 2+ fragments, so single runs use write()).
    _combo_cache = {}

    def _combo_fmt(bold, italic, underline, small_caps):
        key = (bold, italic, underline, small_caps)
        if key not in _combo_cache:
            _combo_cache[key] = wb.add_format({
                **_cell,
                "font_size": 7 if small_caps else 10,
                "bold":      bold,
                "italic":    italic,
                "underline": 1 if underline else 0,
            })
        return _combo_cache[key]

    fmt_root = wb.add_format({   # small italic grey — for Id. root note
        **_cell,
        "font_size": 8, "italic": True, "font_color": "#666666",
    })

    def _write_rich(row, col, src_runs, id_root_str=""):
        """Write a source-text cell with per-run character formatting.
        If id_root_str is non-empty, appends the Id. root note as a small
        italic line below the source text.
        """
        parts = []
        for text, props in src_runs:
            if not text:
                continue
            sc   = props.get("small_caps", False)
            disp = text.upper() if sc else text
            parts.append((
                disp,
                props.get("bold",      False),
                props.get("italic",    False),
                props.get("underline", False),
                sc,
            ))

        if id_root_str:
            parts.append(("\n" + id_root_str, False, True, False, False))

        if not parts:
            ws.write_blank(row, col, None, fmt_normal)
        elif len(parts) == 1:
            disp, bold, italic, underline, sc = parts[0]
            ws.write(row, col, disp, _combo_fmt(bold, italic, underline, sc))
        else:
            rich_args = []
            for disp, bold, italic, underline, sc in parts:
                rich_args.extend([_char_fmt(bold, italic, underline, sc), disp])
            ws.write_rich_string(row, col, *rich_args, fmt_normal)

    # ── Column widths & freeze ────────────────────────────────────────────────
    ws.set_column(0, 0,  6)
    ws.set_column(1, 1, 34)
    ws.set_column(2, 2, 58)
    ws.set_column(3, 3, 28)
    ws.set_column(4, 4, 35)
    ws.freeze_panes(1, 0)

    # ── Header row ────────────────────────────────────────────────────────────
    for col, h in enumerate(["FN #", "Body Sentence Context",
                              "Source Text", "Bluebook Rules", "File Label"]):
        ws.write(0, col, h, fmt_header)

    # ── Data rows ─────────────────────────────────────────────────────────────
    row_idx = 1
    i = 0
    while i < len(rows):
        fn_num = rows[i][0]
        body   = rows[i][1]

        j = i + 1
        while j < len(rows) and rows[j][0] == fn_num:
            j += 1
        count = j - i

        # FN # and Body Sentence cells — merge when multiple sources per footnote
        if count == 1:
            ws.write(row_idx, 0, int(fn_num), fmt_center)
            ws.write(row_idx, 1, body,         fmt_normal)
        else:
            ws.merge_range(row_idx, 0, row_idx + count - 1, 0,
                           int(fn_num), fmt_center)
            ws.merge_range(row_idx, 1, row_idx + count - 1, 1,
                           body, fmt_normal)

        for k in range(count):
            _, _, src_runs, bb, id_root_str, file_label, _, label_span = rows[i + k]
            _write_rich(row_idx + k, 2, src_runs, id_root_str)
            ws.write(row_idx + k, 3, bb, fmt_normal)

            # File Label column — merge for Id. chain spans
            if label_span > 1:
                ws.merge_range(
                    row_idx + k, 4, row_idx + k + label_span - 1, 4,
                    file_label, fmt_normal,
                )
            elif label_span == 1:
                ws.write(row_idx + k, 4, file_label, fmt_normal)
            # label_span == 0: cell covered by earlier merge_range — skip

        row_idx += count
        i = j

    wb.close()
    buf.seek(0)
    return buf.read()


# ── Public entry point ─────────────────────────────────────────────────────────

def _build_rows(docx_path, start_fn, end_fn):
    """
    Shared parsing logic.
    Returns (rows, fn_count, source_count, commentary_count).

    Row tuple:
      (fn_num_str, body, src_runs, bb_rules, id_root_str,
       file_label, chain_root_fn, label_span)

      chain_root_fn — fn number of the Id./supra root (None for regular sources).
      label_span    — number of table rows the File Label cell spans from this
                      row (>1 for the first row of an Id. chain; 0 for rows
                      covered by an earlier span; 1 for standalone rows).
    """
    buffer_start  = max(1, start_fn - 50)
    all_fn_runs   = extract_footnote_runs(docx_path, buffer_start, end_fn)
    body_contexts = extract_body_contexts(docx_path, start_fn, end_fn)
    id_root_map   = _build_id_root_map(all_fn_runs)

    # ── Pass 1: pre-compute canonical labels for supra note-N resolution ─────
    # For each footnote in buffer+range, compute the label of its first source
    # (without supra_ref_label, to avoid circular lookups).
    canonical_labels = {}
    for fn_num in range(buffer_start, end_fn + 1):
        runs = all_fn_runs.get(fn_num)
        if not runs:
            continue
        plain    = "".join(t for t, _ in runs)
        sources  = split_sources(plain)
        fn_roots = id_root_map.get(fn_num, [])
        if sources:
            src_runs_list = assign_source_runs(runs, sources)
            src_r     = src_runs_list[0]
            root_info = fn_roots[0] if fn_roots else None
            bb        = _classify_source(src_r)
            canonical_labels[fn_num] = _suggest_filename(
                fn_num, src_r, bb, root_info=root_info
            )
        else:
            canonical_labels[fn_num] = ""

    # ── Pass 2: build raw rows for the requested range ───────────────────────
    # 7-element tuples (label_span not yet set).
    raw_rows         = []
    source_count     = 0
    commentary_count = 0

    for fn_num in range(start_fn, end_fn + 1):
        runs = all_fn_runs.get(fn_num)
        if not runs:
            continue

        body     = body_contexts.get(fn_num, "")
        plain    = "".join(t for t, _ in runs)
        sources  = split_sources(plain)
        fn_roots = id_root_map.get(fn_num, [])

        if sources:
            src_runs_list = assign_source_runs(runs, sources)
            for src_idx, src_runs in enumerate(src_runs_list):
                root_info = fn_roots[src_idx] if src_idx < len(fn_roots) else None
                bb        = _classify_source(src_runs)

                # Supra resolution: look up the canonical label of the
                # referenced footnote so we can format "FN{ref}-{fn}_..." .
                src_plain       = "".join(t for t, _ in src_runs)
                supra_ref_label = None
                if re.search(r"\bsupra\b", src_plain, re.IGNORECASE):
                    note_m = re.search(
                        r"\bsupra\s+note\s+(\d+)\b", src_plain, re.IGNORECASE
                    )
                    if note_m:
                        ref_fn          = int(note_m.group(1))
                        supra_ref_label = canonical_labels.get(ref_fn, "")

                file_label = _suggest_filename(
                    fn_num, src_runs, bb,
                    root_info=root_info,
                    supra_ref_label=supra_ref_label,
                )

                id_root_str = ""
                if root_info:
                    root_fn_num, root_text = root_info
                    snippet = root_text[:100].rstrip()
                    if len(root_text) > 100:
                        snippet += "…"
                    id_root_str = f"↳ FN {root_fn_num}: {snippet}"

                chain_root_fn = root_info[0] if root_info else None
                raw_rows.append(
                    (str(fn_num), body, src_runs, bb,
                     id_root_str, file_label, chain_root_fn)
                )
                source_count += 1
        else:
            commentary_count += 1
            raw_rows.append((str(fn_num), body, runs, "", "", "", None))

    fn_count = sum(1 for fn in range(start_fn, end_fn + 1) if fn in all_fn_runs)

    # ── Pass 3: compute label_span for Id./supra chain coalescing ────────────
    # Consecutive rows that share the same non-None chain_root_fn are merged
    # into one File Label cell.  The merged label uses the range
    # FN{root}-{last_fn_in_chain}_... so editors see the full span at once.
    rows = []
    n    = len(raw_rows)
    i    = 0
    while i < n:
        chain_root_fn = raw_rows[i][6]
        if chain_root_fn is None:
            rows.append(raw_rows[i] + (1,))
            i += 1
        else:
            # Find the full run of consecutive rows with the same root
            j = i + 1
            while j < n and raw_rows[j][6] == chain_root_fn:
                j += 1
            chain_len = j - i
            last_fn   = int(raw_rows[j - 1][0])

            # Reformat the label to cover the whole chain range
            first_label = raw_rows[i][5]
            chain_label = re.sub(
                rf"^(FN{chain_root_fn})(?:-\d+)?_",
                rf"\1-{last_fn}_",
                first_label,
            )
            # Fallback if the label didn't match the expected prefix
            if chain_label == first_label and not re.search(r"-\d+_", first_label):
                chain_label = re.sub(
                    rf"^FN{chain_root_fn}(?=_|$)",
                    f"FN{chain_root_fn}-{last_fn}",
                    first_label,
                )

            for k in range(chain_len):
                span = chain_len if k == 0 else 0
                r    = raw_rows[i + k]
                rows.append((r[0], r[1], r[2], r[3], r[4], chain_label, r[6], span))

            i = j

    return rows, fn_count, source_count, commentary_count


def process_footnotes_to_pdf(docx_path, start_fn, end_fn,
                              original_filename="document.docx"):
    """Parse the docx and return (pdf_bytes, fn_count, source_count)."""
    rows, fn_count, source_count, commentary_count = _build_rows(
        docx_path, start_fn, end_fn
    )
    pdf_bytes = build_pdf(
        rows, original_filename, start_fn, end_fn,
        source_count, commentary_count,
    )
    return pdf_bytes, fn_count, source_count


def process_footnotes_to_xlsx(docx_path, start_fn, end_fn,
                               original_filename="document.docx"):
    """Parse the docx and return (xlsx_bytes, fn_count, source_count)."""
    rows, fn_count, source_count, commentary_count = _build_rows(
        docx_path, start_fn, end_fn
    )
    xlsx_bytes = build_xlsx(
        rows, original_filename, start_fn, end_fn,
        source_count, commentary_count,
    )
    return xlsx_bytes, fn_count, source_count
