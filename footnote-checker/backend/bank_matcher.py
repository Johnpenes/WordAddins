"""
bank_matcher.py — suggest filenames for citations using the file bank.

Two strategies, in order:
  1. Bank lookup — if the citation closely matches an existing bank entry
     (high token overlap), return that filename with the current FN# prefix.
  2. Structured generation — classify the citation type, learn the naming
     format from existing bank entries of that type (separator style, year
     inclusion, spacing), extract author/title/reporter/year, and assemble
     following the learned pattern.

Bank format (bank.json): list of { "filename": str, "tokens": [str, ...], "type": str }
Type field is populated lazily on first load via _migrate_bank_if_needed.
"""

import json
import re
import threading
from dataclasses import dataclass
from pathlib import Path

_BANK_PATH = Path(__file__).parent / "bank.json"

# Minimum token overlap to trust a bank lookup over generation
_LOOKUP_THRESHOLD = 4
_MIN_TOKEN_LEN = 3

_SIGNALS = re.compile(
    r"^(see(\s+also|\s+generally)?|but\s+see|accord|cf\.?|e\.g\.,?|id\.?|supra|infra"
    r"|compare|contra|quoting|citing)\s*",
    re.IGNORECASE,
)

_lock = threading.Lock()
_bank: list[dict] | None = None


# ── FormatSpec ────────────────────────────────────────────────────────────────

@dataclass
class FormatSpec:
    separator: str = "-"              # "-" (bare) or ", " (cases)
    space_around_sep: bool = False    # True → "Author - Title", False → "Author-Title"
    include_year: bool = False        # append (YYYY) at end
    year_for_cases: bool = False      # append (YYYY) for case filenames specifically
    include_parenthetical: bool = False  # keep trailing descriptive parentheticals


_DEFAULT_FORMATS: dict[str, FormatSpec] = {
    "case":     FormatSpec(separator=", ", space_around_sep=False, include_year=False, year_for_cases=False),
    "statute":  FormatSpec(),
    "book":     FormatSpec(separator="-",  space_around_sep=False, include_year=True),
    "article":  FormatSpec(separator="-",  space_around_sep=False, include_year=False),
    "internet": FormatSpec(separator="-",  space_around_sep=False, include_year=False),
    "case_doc": FormatSpec(separator="-",  space_around_sep=True,  include_year=False),
    "unknown":  FormatSpec(separator="-",  space_around_sep=False, include_year=False),
}

# Invalidated whenever add_to_bank writes a new entry
_format_cache: dict[str, FormatSpec] = {}


# ── Type helpers ──────────────────────────────────────────────────────────────

def _classification_to_type(classification: str) -> str:
    cls = classification.lower()
    if any(r in cls for r in ("rule 10", "rule 11", "rule 12")):
        return "case"
    if any(r in cls for r in ("rule 12.9", "rule 14")):
        return "statute"
    if "rule 15" in cls:
        return "book"
    if "rule 16" in cls:
        return "article"
    if "rule 18" in cls:
        return "internet"
    return "unknown"


def _infer_type_from_filename(filename: str) -> str:
    """Classify an existing bank entry from its filename alone (used for migration)."""
    name = _strip_fn_prefix(re.sub(r"\.(pdf|png|docx)$", "", filename, flags=re.IGNORECASE))

    # Case (with optional case_doc sub-type)
    if re.search(r"\bv\.?\s+[A-Z]", name):
        tail = re.search(
            r"[-–]\s*(Complaint|Brief|Order|Motion|Memorandum|Petition|Reply"
            r"|Injunction|Judgment|Amended|Declaration)",
            name, re.IGNORECASE,
        )
        return "case_doc" if tail else "case"

    # Statute
    if re.search(
        r"[§¶]|U\.S\.C\.?A?|Revised\s+Statute|Code\s+Ann\.|Municipal\s+Code"
        r"|\bOrdinance\b|\bConst\b",
        name, re.IGNORECASE,
    ):
        return "statute"

    # Book: year at end
    if re.search(r"\(\d{4}\)\s*$", name):
        return "book"

    # Article: single author last name before separator
    m = re.match(r"^([A-Z][A-Za-z'\-]+(?:\s+[A-Z][A-Za-z'\-]+){0,2})\s*[-–_]", name)
    if m and len(m.group(1).split()) == 1:
        return "article"

    # Internet: org name followed by year
    if re.match(r"^[A-Z][A-Za-z]+\s*[-–]\s*\d{4}", name):
        return "internet"

    # Case doc keywords
    if re.search(
        r"\b(Brief|Complaint|Amicus|Petition|Order|Memorandum|Report|Assessment)\b",
        name, re.IGNORECASE,
    ):
        return "case_doc"

    return "unknown"


# ── Bank I/O ──────────────────────────────────────────────────────────────────

def _migrate_bank_if_needed(bank: list[dict]) -> list[dict]:
    """One-time migration: add 'type' field to entries that lack it, then persist."""
    if not any("type" not in e for e in bank):
        return bank
    for e in bank:
        if "type" not in e:
            e["type"] = _infer_type_from_filename(e["filename"])
    _BANK_PATH.write_text(json.dumps(bank, indent=2))
    return bank


def _load_bank() -> list[dict]:
    global _bank
    with _lock:
        if _bank is None:
            raw = json.loads(_BANK_PATH.read_text()) if _BANK_PATH.exists() else []
            _bank = _migrate_bank_if_needed(raw)
    return _bank


def _strip_fn_prefix(filename: str) -> str:
    return re.sub(r"^(?:FN)?[\d,\s\-–]*_", "", filename)


def add_to_bank(filename: str, classification: str = "") -> None:
    """Add a confirmed filename to the bank."""
    bank = _load_bank()
    with _lock:
        if any(e["filename"] == filename for e in bank):
            return
        tokens = _tokenize_filename(filename, classification)
        type_name = _classification_to_type(classification)
        bank.append({"filename": filename, "tokens": tokens, "type": type_name})
        _BANK_PATH.write_text(json.dumps(bank, indent=2))
        _format_cache.clear()  # invalidate so _learn_format picks up new entry


def _tokenize_filename(filename: str, classification: str = "") -> list[str]:
    name = re.sub(r"\.(pdf|png|docx)$", "", filename, flags=re.IGNORECASE)
    name = _strip_fn_prefix(name)
    words = [w.strip(".,;:()-_") for w in re.split(r"[\s_\-]+", name.lower())]
    tokens = []
    for w in words:
        # Skip purely numeric tokens (page numbers, volumes, years, WL numbers)
        if re.match(r"^\d+$", w):
            continue
        # Skip mixed alphanumeric IDs (docket numbers like 25cv05989, 4209227)
        if re.match(r"^[\da-z]*\d[\da-z]*$", w) and len(w) > 4:
            continue
        if len(w) > _MIN_TOKEN_LEN:
            tokens.append(w)
    # Add classification as a pseudo-token so type is part of the match
    if classification:
        for rule_m in re.finditer(r"rule\s*([\d.]+)", classification, re.IGNORECASE):
            tokens.append(f"_rule{rule_m.group(1)}")
    return tokens


# ── Format learning ───────────────────────────────────────────────────────────

def _learn_format(type_name: str) -> FormatSpec:
    """
    Analyze bank entries of the given type and return a FormatSpec that
    reflects the user's observed naming conventions (separator style, year,
    spacing).  Falls back to _DEFAULT_FORMATS when no entries exist yet.
    Results are cached until the bank changes.
    """
    if type_name in _format_cache:
        return _format_cache[type_name]

    bank = _load_bank()
    names = [
        _strip_fn_prefix(re.sub(r"\.(pdf|png|docx)$", "", e["filename"], flags=re.IGNORECASE))
        for e in bank if e.get("type") == type_name
    ]

    if not names:
        return _DEFAULT_FORMATS.get(type_name, FormatSpec())

    # Separator style: spaced dash " - " vs bare dash "word-word"
    spaced = sum(1 for n in names if re.search(r"\s[-–]\s", n))
    bare   = sum(1 for n in names if re.search(r"[A-Za-z]-[A-Za-z]", n))
    space_around = spaced > bare

    # Year inclusion: majority vote
    total = len(names)
    with_year = sum(1 for n in names if re.search(r"\(\d{4}\)\s*$", n))
    include_year = with_year > total / 2

    # Parenthetical detection: does the bank include trailing descriptive parens?
    # Uses lowercase-start heuristic to distinguish "(defining X)" from "(2024)"
    has_descriptive_paren = sum(1 for n in names if re.search(r"\([a-z]", n))
    include_parenthetical = has_descriptive_paren > len(names) / 2

    if type_name == "case":
        # Cases always use ", " reporter separator; year is separate flag
        fmt = FormatSpec(
            separator=", ",
            space_around_sep=False,
            include_year=False,
            year_for_cases=include_year,
            include_parenthetical=include_parenthetical,
        )
    else:
        fmt = FormatSpec(
            separator="-",
            space_around_sep=space_around,
            include_year=include_year,
            include_parenthetical=include_parenthetical,
        )

    _format_cache[type_name] = fmt
    return fmt


# ── Token helpers ─────────────────────────────────────────────────────────────

def _normalise(text: str) -> str:
    s = _SIGNALS.sub("", text).strip().lstrip(".,; ")
    return re.sub(r"[^a-z0-9\s]", " ", s.lower())


def _citation_tokens(text: str) -> list[str]:
    norm = _normalise(text)
    stop = {
        "the", "and", "for", "with", "from", "this", "that", "have", "has",
        "been", "were", "they", "their", "into", "also", "over", "such",
        "after", "under", "upon", "about", "would", "which", "when", "than",
        "more", "some",
    }
    tokens = []
    for w in norm.split():
        if w in stop:
            continue
        if w.isdigit() or len(w) >= _MIN_TOKEN_LEN:
            tokens.append(w)
    return tokens


def _score(citation_tokens: list[str], bank_tokens: list[str]) -> int:
    citation_set = set(citation_tokens)
    return sum(1 for t in bank_tokens if t in citation_set)


# ── Structured extraction ─────────────────────────────────────────────────────

def _safe(s: str) -> str:
    """Strip characters illegal in filenames (keep colons — Mac allows them)."""
    s = re.sub(r'[<>"/\\|?*\x00-\x1f]', "", s)
    return re.sub(r"\s+", " ", s).strip().rstrip(".,;")


def _strip_signals(text: str) -> str:
    return _SIGNALS.sub("", text).strip().lstrip(".,; ")


def _strip_commentary(text: str) -> str:
    """Remove pincite and trailing prose parentheticals; keep year paren."""
    text = re.sub(r",\s*\d+(?=\s*\(\d{4}\))", "", text)
    m = re.search(r"\(\d{4}\)", text)
    return text[: m.end()].rstrip(".,; ") if m else text.strip()


def _extract_case(text: str) -> str | None:
    """
    For cases: return 'CaseName, Reporter Volume' (no year paren).
    Pattern: ... v. ... , Volume Reporter Page (Year)
    """
    s = _strip_signals(text)
    # Find 'v.' or 'v ' as case indicator
    v_m = re.search(r"\bv\.?\s+[A-Z]", s)
    if not v_m:
        return None
    # Strip trailing year paren and beyond
    s = re.sub(r"\s*\(\d{4}\).*$", "", s).strip().rstrip(".,;")
    return _safe(s)


def _extract_author_last(text: str) -> str:
    """
    Extract the last name from an author citation.
    Handles: 'First Last,', 'Last, First,', 'Last et al.,'
    Returns empty string if no clear author found.
    """
    s = _strip_signals(text)
    # Match author segment before first comma
    m = re.match(r"^([A-Z][A-Za-z\-']+(?:\s+[A-Z]?[A-Za-z\-'.]+){0,4})\s*,", s)
    if not m:
        return ""
    author = m.group(1).strip()
    words = author.split()
    # et al
    if len(words) >= 2 and words[-2].lower() == "et" and words[-1].rstrip(".").lower() == "al":
        return words[-3] + " et al" if len(words) >= 3 else "et al"
    # Last word is the last name (handles 'William P. Quigley' → 'Quigley')
    last = words[-1].rstrip(".")
    # Reject if it looks like an institution or acronym (all caps, >4 chars)
    if last.isupper() and len(last) > 3:
        return ""
    return last


def _extract_title(text: str) -> str:
    """
    Extract title portion — everything after 'Author,' up to the source/volume info.
    Strips year parens and trailing commentary.
    """
    s = _strip_signals(text)
    # Drop author segment
    first_comma = s.find(",")
    if first_comma == -1:
        return _safe(_strip_commentary(s))
    after_author = s[first_comma + 1:].strip().lstrip(".,; ")
    # Title ends at the first ', <digit>' (volume/page boundary)
    boundary = re.search(r",\s*\d", after_author)
    title = after_author[: boundary.start()].strip() if boundary else after_author
    # Strip [hereinafter ...] and year parens
    title = re.sub(r"\s*\[[^\]]*\]\s*$", "", title).strip()
    title = re.sub(r"\s*\(\d{4}\)\s*$", "", title).strip()
    return _safe(title.rstrip(".,;"))


def _extract_internet(text: str) -> tuple[str, str]:
    """Return (author_last, title) for internet/news citations."""
    s = _strip_signals(text)
    # Strip URL
    s = re.split(r"https?://", s, maxsplit=1)[0].strip().rstrip(",")
    # Strip date paren
    s = re.sub(r"\s*\([^)]*(?:\d{4}|n\.d\.)[^)]*\)\s*$", "", s, flags=re.IGNORECASE).strip().rstrip(",")

    parts = [p.strip() for p in s.split(",")]
    if not parts:
        return "", s

    # First segment is author if 1–4 title-cased words, no digits
    author_last = ""
    title_start = 0
    candidate = parts[0]
    words = candidate.split()
    if (1 <= len(words) <= 4
            and all(w[0].isupper() for w in words if w)
            and not any(c.isdigit() for c in candidate)):
        author_last = _extract_author_last(candidate + ",")
        title_start = 1

    title_parts = parts[title_start: title_start + 2]
    title = ", ".join(title_parts).strip().rstrip(".,;")
    return author_last, _safe(title)


# ── Short cite detection ──────────────────────────────────────────────────────

_REPORTER_RE = re.compile(
    r"\b\d+\s+[A-Z][A-Za-z.]*(?:\s+[A-Z][A-Za-z.]*)?\s+\d+\b"  # e.g. 603 U.S. 520
)

def _is_short_cite(stripped: str) -> bool:
    """
    Return True if the text is a back-reference / short cite that has no
    independent file name — these should resolve to a root footnote instead.
    """
    s = stripped

    # Id. / Id. at NNN
    if re.match(r"^id\.?(\s+at\s+[\d\-–,n.]+)?\.?$", s, re.IGNORECASE):
        return True

    # Bare supra / infra at start
    if re.match(r"^(supra|infra)\b", s, re.IGNORECASE):
        return True

    # "Author, supra" or "Title, supra note N"
    if re.search(r"\bsupra\b", s, re.IGNORECASE):
        return True

    # Reporter-only short cite: "603 U.S. at 543" or "50 F.4th at 812"
    if re.match(r"^\d+\s+[A-Z][A-Za-z.]+(?:\s+[A-Z][A-Za-z.]*)?\s+at\s+[\d\-–,]+", s):
        return True

    # Partial case + reporter + "at": "Grants Pass, 603 U.S. at 589"
    if re.match(r"^[A-Z][A-Za-z\s]{1,30},\s*\d+\s+[A-Z][A-Za-z.]+.*\bat\b", s):
        return True

    # Very short after stripping — fewer than 3 meaningful tokens, likely a stub
    tokens = [w for w in re.split(r"\W+", s) if len(w) > 2]
    if len(tokens) < 3:
        return True

    return False


# ── Generation by citation type ───────────────────────────────────────────────

def _generate_filename(text: str, fn_number: int | str, classification: str) -> str | None:
    """
    Generate a filename using the format learned from existing bank entries
    of the same citation type.  Falls back to sensible defaults when the bank
    has no entries for a type yet.
    """
    prefix = f"FN{fn_number}"
    cls = classification.lower()
    s = _strip_signals(text)

    # ── Back-references: Id., supra, infra — no filename possible ─────────────
    if any(r in cls for r in ("rule 4.1", "rule 4.2")):
        return None
    stripped = s.strip().rstrip(".,; ")
    if _is_short_cite(stripped):
        return None

    # ── Determine type and learn format from bank ─────────────────────────────
    type_name = _classification_to_type(classification)

    # Override for statute signals in text even if classification is ambiguous
    if type_name == "unknown" and re.search(r"[§¶]|\bU\.S\.C\b", s):
        type_name = "statute"

    fmt = _learn_format(type_name)

    # ── Assembly helpers ──────────────────────────────────────────────────────
    def join(a: str, b: str) -> str:
        if a and b:
            sep = f" {fmt.separator} " if fmt.space_around_sep else fmt.separator
            return f"{a}{sep}{b}"
        return b or a

    def yr(raw: str) -> str:
        m = re.search(r"\((\d{4})\)", raw)
        return f" ({m.group(1)})" if m and fmt.include_year else ""

    # ── Cases ─────────────────────────────────────────────────────────────────
    if type_name == "case":
        case = _extract_case(text)
        if case:
            suffix = ""
            if fmt.year_for_cases:
                ym = re.search(r"\((\d{4})\)", s)
                suffix = f" ({ym.group(1)})" if ym else ""
            return f"{prefix}_{case}{suffix}"

    # ── Statutes ──────────────────────────────────────────────────────────────
    if type_name == "statute":
        stat = re.sub(r"\(\d{4}\).*$", "", s).strip().rstrip(".,;")
        if not fmt.include_parenthetical:
            stat = re.sub(r"\s*\([^)]*\)\s*$", "", stat).strip().rstrip(".,;")
        return f"{prefix}_{_safe(stat)}" if stat else None

    # ── Books ─────────────────────────────────────────────────────────────────
    if type_name == "book":
        a = _extract_author_last(text)
        t = _extract_title(text)
        body = join(a, t)
        return f"{prefix}_{body}{yr(s)}" if body else None

    # ── Articles ──────────────────────────────────────────────────────────────
    if type_name == "article":
        a = _extract_author_last(text)
        t = _extract_title(text)
        body = join(a, t)
        # Articles never include year, even if bank majority happens to have one
        return f"{prefix}_{body}" if body else None

    # ── Internet / news ───────────────────────────────────────────────────────
    if type_name == "internet":
        al, t = _extract_internet(text)
        body = join(al, t)
        return f"{prefix}_{body}{yr(s)}" if body else None

    # ── Unknown / ambiguous — generic author+title extraction ─────────────────
    a = _extract_author_last(text)
    t = _extract_title(text)
    if a and t:
        return f"{prefix}_{a}-{t}"

    # Last resort: first 80 chars of stripped text
    clean = _safe(_strip_commentary(s))
    return f"{prefix}_{clean[:80]}" if clean else None


# ── Public API ────────────────────────────────────────────────────────────────

def is_back_reference(citation_text: str, classification: str = "") -> bool:
    """Public helper — True if this source is a back-reference with no own file."""
    cls = classification.lower()
    if any(r in cls for r in ("rule 4.1", "rule 4.2")):
        return True
    s = _strip_signals(citation_text).strip().rstrip(".,; ")
    return _is_short_cite(s)


def match_filename(citation_text: str, fn_number: int | str, classification: str = "") -> str | None:
    """
    Suggest a filename for the given citation.

    1. Try bank lookup — if token overlap is high enough, the citation is
       already catalogued; return the existing name with updated FN prefix.
    2. Otherwise generate a new name using the format learned from the bank
       for citations of this type.
    """
    bank = _load_bank()
    c_tokens = _citation_tokens(citation_text)
    # Inject classification pseudo-tokens so type factors into scoring
    for rule_m in re.finditer(r"rule\s*([\d.]+)", classification, re.IGNORECASE):
        c_tokens.append(f"_rule{rule_m.group(1)}")

    # Bank lookup
    if c_tokens and bank:
        best_score, best_entry = 0, None
        for entry in bank:
            s = _score(c_tokens, entry.get("tokens", []))
            if s > best_score:
                best_score = s
                best_entry = entry
        if best_score >= _LOOKUP_THRESHOLD and best_entry:
            fname_no_prefix = _strip_fn_prefix(best_entry["filename"])
            fname_no_prefix = re.sub(r"\.(pdf|png|docx)$", "", fname_no_prefix, flags=re.IGNORECASE)
            return f"FN{fn_number}_{fname_no_prefix}"

    # Structured generation using learned format
    return _generate_filename(citation_text, fn_number, classification)
