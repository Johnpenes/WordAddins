"""
bank_matcher.py — suggest filenames for citations using the file bank.

Two strategies, in order:
  1. Bank lookup — if the citation closely matches an existing bank entry
     (high token overlap), return that filename with the current FN# prefix.
  2. Structured generation — classify the citation type, extract author/title/
     reporter/year, and format following the naming conventions the bank
     demonstrates for that type.

Bank format (bank.json): list of { "filename": str, "tokens": [str, ...] }
"""

import json
import re
import threading
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


# ── Bank I/O ──────────────────────────────────────────────────────────────────

def _load_bank() -> list[dict]:
    global _bank
    with _lock:
        if _bank is None:
            _bank = json.loads(_BANK_PATH.read_text()) if _BANK_PATH.exists() else []
    return _bank


def _strip_fn_prefix(filename: str) -> str:
    return re.sub(r"^(?:FN)?[\d,\s\-–]*_", "", filename)


def add_to_bank(filename: str) -> None:
    """Add a confirmed filename to the bank."""
    bank = _load_bank()
    with _lock:
        if any(e["filename"] == filename for e in bank):
            return
        tokens = _tokenize_filename(filename)
        bank.append({"filename": filename, "tokens": tokens})
        _BANK_PATH.write_text(json.dumps(bank, indent=2))


def _tokenize_filename(filename: str) -> list[str]:
    name = re.sub(r"\.(pdf|png|docx)$", "", filename, flags=re.IGNORECASE)
    name = _strip_fn_prefix(name)
    words = [w.strip(".,;:()-_") for w in re.split(r"[\s_\-]+", name.lower())]
    return [w for w in words if w.isdigit() or len(w) > _MIN_TOKEN_LEN]


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
    # Pattern: digits + reporter + "at" + digits — no case name
    if re.match(r"^\d+\s+[A-Z][A-Za-z.]+(?:\s+[A-Z][A-Za-z.]*)?\s+at\s+[\d\-–,]+", s):
        return True

    # Partial case + reporter + "at": "Grants Pass, 603 U.S. at 589"
    # One or two words, comma, reporter, "at"
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
    Generate a filename from structured extraction, following bank conventions:
      Cases     →  CaseName, Reporter Volume.pdf
      Articles  →  LastName-Title.pdf
      Books     →  LastName-Title (Year).pdf   [year preserved for books]
      Internet  →  LastName-Title.pdf  or  Title.pdf
      Statutes  →  StatuteRef.pdf
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

    # ── Cases (Rule 10, 11, 12) ───────────────────────────────────────────────
    if any(r in cls for r in ("rule 10", "rule 11", "rule 12")):
        case = _extract_case(text)
        if case:
            return (f"{prefix}_{case}")

    # ── Statutes (Rule 12.9, 14) ──────────────────────────────────────────────
    if any(r in cls for r in ("rule 12.9", "rule 14")) or re.search(r"[§¶]|\bU\.S\.C\b", s):
        stat = re.sub(r"\(\d{4}\).*$", "", s).strip().rstrip(".,;")
        return (f"{prefix}_{_safe(stat)}") if stat else None

    # ── Books (Rule 15) ───────────────────────────────────────────────────────
    if "rule 15" in cls:
        author = _extract_author_last(text)
        title = _extract_title(text)
        year_m = re.search(r"\((\d{4})\)", s)
        year = f" ({year_m.group(1)})" if year_m else ""
        if author and title:
            return (f"{prefix}_{author}-{title}{year}")
        if title:
            return (f"{prefix}_{title}{year}")

    # ── Journal articles (Rule 16) ────────────────────────────────────────────
    if "rule 16" in cls:
        author = _extract_author_last(text)
        title = _extract_title(text)
        if author and title:
            return (f"{prefix}_{author}-{title}")
        if title:
            return (f"{prefix}_{title}")

    # ── Internet / news (Rule 18) ─────────────────────────────────────────────
    if "rule 18" in cls:
        author_last, title = _extract_internet(text)
        if author_last and title:
            return (f"{prefix}_{author_last}-{title}")
        if title:
            return (f"{prefix}_{title}")

    # ── NOTE / ambiguous — try generic author+title extraction ────────────────
    author = _extract_author_last(text)
    title = _extract_title(text)
    if author and title:
        return (f"{prefix}_{author}-{title}")

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
    2. Otherwise generate a new name from structured extraction, following
       the naming conventions demonstrated by the bank.
    """
    bank = _load_bank()
    c_tokens = _citation_tokens(citation_text)

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

    # Structured generation
    return _generate_filename(citation_text, fn_number, classification)
