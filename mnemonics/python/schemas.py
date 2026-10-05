"""CSV schema constants and validators for Memory Palace."""

from __future__ import annotations

import html
import re
from pathlib import Path
from typing import Any
from urllib.parse import quote

STUDIES_HEADERS = ["id", "title", "notes_rel_path", "sort_order", "active"]
PALACES_HEADERS = [
    "id",
    "study_id",
    "palace_number",
    "title",
    "character_name",
    "image_rel_path",
    "depth_slots_used",
    "image_prompt",
    "palace_notes",
]
PALACE_IMAGES_HEADERS = [
    "id",
    "palace_id",
    "image_rel_path",
    "caption",
    "sort_order",
]
STUDY_IMAGES_HEADERS = [
    "id",
    "study_id",
    "image_rel_path",
    "caption",
    "sort_order",
]
BEASTS_HEADERS = [
    "id",
    "palace_id",
    "peg_code",
    "beast_name",
    "beast_source",
    "sensory_channel",
    "is_smashed",
    "sort_order",
]
ATOMS_HEADERS = [
    "id",
    "beast_id",
    "kind",
    "zone",
    "zone_label",
    "concept",
    "keywords",
    "quote",
    "story",
    "ipa",
    "sort_order",
]

# Legacy read cap. New or rewritten atoms use exactly two bracket groups.
ATOM_KEYWORDS_MAX_PAIRS = 10
ATOM_KEYWORDS_NEW_PAIRS = 2


_CONCEPT_NOTE_SEPS = (" — Note:", " – Note:", " - Note:")


def split_concept_note(raw: str) -> tuple[str, str]:
    """Split compressed core from optional ` — Note:` suffix."""
    text = (raw or "").strip()
    for sep in _CONCEPT_NOTE_SEPS:
        idx = text.find(sep)
        if idx >= 0:
            return text[:idx].rstrip(), text[idx:]
    return text, ""


def format_concept_thought_groups(raw: str | None) -> str:
    """Show concept thought groups as `[chunk] [chunk]`; keep Note unbracketed.

    Legacy cores delimited with ` | ` are converted. Already-bracketed cores
    and unchunked cores are left unchanged. Quote/story cites like `[cite: 1]`
    stay inside their chunk.
    """
    text = (raw or "").strip()
    if not text:
        return ""
    core, note = split_concept_note(text)
    if " | " in core:
        parts = [p.strip() for p in core.split(" | ") if p.strip()]
        if len(parts) >= 2:
            core = " ".join(f"[{p}]" for p in parts)
    return f"{core}{note}"


def concept_bracket_groups(raw: str | None) -> list[str] | None:
    """Return top-level concept groups, or None when the core is not only groups."""
    core, _note = split_concept_note(raw or "")
    groups: list[str] = []
    outside: list[str] = []
    current: list[str] = []
    depth = 0
    for char in core:
        if char == "[":
            if depth == 0:
                if "".join(outside).strip():
                    return None
                outside = []
                current = []
            else:
                current.append(char)
            depth += 1
        elif char == "]":
            if depth == 0:
                return None
            depth -= 1
            if depth == 0:
                group = "".join(current).strip()
                if not group:
                    return None
                groups.append(group)
            else:
                current.append(char)
        elif depth:
            current.append(char)
        else:
            outside.append(char)
    if depth or "".join(outside).strip():
        return None
    return groups


def normalize_atom_keywords(raw: str | None) -> str:
    """Normalize `keywords` to `Keyword | Word || …` with at most ATOM_KEYWORDS_MAX_PAIRS pairs."""
    text = (raw or "").strip()
    if not text:
        return ""
    chunks = [c.strip() for c in re.split(r"\s*\|\|\s*", text) if c.strip()]
    pairs: list[str] = []
    for chunk in chunks:
        if " | " in chunk:
            left, right = chunk.split(" | ", 1)
        elif "|" in chunk:
            left, right = chunk.split("|", 1)
        else:
            continue
        left, right = left.strip(), right.strip()
        if not left or not right:
            continue
        pairs.append(f"{left} | {right}")
        if len(pairs) >= ATOM_KEYWORDS_MAX_PAIRS:
            break
    return " || ".join(pairs)


def _strict_keyword_pairs(raw: str | None) -> tuple[list[tuple[str, str]], str | None]:
    text = (raw or "").strip()
    if not text:
        return [], "keywords are required"
    pairs: list[tuple[str, str]] = []
    for index, chunk in enumerate(re.split(r"\s*\|\|\s*", text), start=1):
        if not chunk.strip():
            return [], f"keyword pair {index} is empty"
        fields = [part.strip() for part in chunk.split("|")]
        if len(fields) != 2 or not all(fields):
            return [], (f"keyword pair {index} must be `Keyword | RecognizableWord`")
        pairs.append((fields[0], fields[1]))
    return pairs, None


def _term_in_group(group: str, term: str) -> bool:
    return term.casefold() in group.casefold()


def _pairs_match_groups_one_to_one(
    groups: list[str], pairs: list[tuple[str, str]]
) -> bool:
    """Exactly one pair per group; RecognizableWord must appear in that group."""
    if len(groups) != len(pairs):
        return False
    return all(
        _term_in_group(group, right) for group, (_left, right) in zip(groups, pairs)
    )


def _opening_clause_ok(group: str) -> bool:
    """Group 1 is `I use {name} to {verb}` and does not end the sentence."""
    text = group.strip()
    if not text or text.endswith("."):
        return False
    return re.fullmatch(r"I use .+ to \S.*", text) is not None


def validate_atom_mnemonics(concept: str | None, keywords: str | None) -> str | None:
    """Validate the two-part sentence contract for a new or rewritten atom."""
    groups = concept_bracket_groups(concept)
    if groups is None:
        return (
            "concept core must contain only square-bracket groups followed by an "
            "optional unbracketed ` — Note:`"
        )
    if len(groups) != ATOM_KEYWORDS_NEW_PAIRS:
        return (
            f"concept must contain exactly {ATOM_KEYWORDS_NEW_PAIRS} bracket groups "
            "([I use {Name} to {verb}] [continuation])"
        )
    if not _opening_clause_ok(groups[0]):
        return "concept opening must be `I use … to …` without a closing period"

    pairs, error = _strict_keyword_pairs(keywords)
    if error:
        return error
    if len(pairs) != len(groups):
        return "exactly one keyword pair is required per concept bracket group"
    if len(pairs) != ATOM_KEYWORDS_NEW_PAIRS:
        return f"keywords must contain exactly {ATOM_KEYWORDS_NEW_PAIRS} pairs"
    if not _pairs_match_groups_one_to_one(groups, pairs):
        return (
            "each keyword RecognizableWord must appear in its matching bracket group "
            "in left-to-right source order"
        )
    return None


def iter_keyword_pairs(raw: str | None) -> list[tuple[str, str]]:
    """Return (tangible Keyword, concept RecognizableWord) pairs."""
    normalized = normalize_atom_keywords(raw)
    if not normalized:
        return []
    out: list[tuple[str, str]] = []
    for chunk in normalized.split(" || "):
        chunk = chunk.strip()
        if " | " in chunk:
            left, right = chunk.split(" | ", 1)
        elif "|" in chunk:
            left, right = chunk.split("|", 1)
        else:
            continue
        left, right = left.strip(), right.strip()
        if left and right:
            out.append((left, right))
    return out


def _term_pattern(term: str) -> str | None:
    if len(term) < 2:
        return None
    return rf"(?i)(?<![A-Za-z0-9]){re.escape(term)}(?![A-Za-z0-9])"


def _find_term(text: str, term: str) -> re.Match[str] | None:
    pattern = _term_pattern(term)
    if not pattern:
        return None
    return re.search(pattern, text)


# Light fields, dark text. GitHub's <mark> wash is too dim on a black page.
# The first concept phrase is blue; every later phrase stays orange.
_KEYWORD_ORANGE = "ffd966"
_KEYWORD_BLUE = "9fd4ff"


def _keyword_chip_md(phrase: str, mnemonic: str, *, lead: bool = False) -> str:
    """Highlight only the concept phrase. The mnemonic stays plain text.

    GitHub strips custom colors, and its `<mark>` wash has almost no contrast
    in dark mode. A flat badge is a light chip with dark text, which is the
    same pairing the web app paints. Only the first phrase in the concept is blue.
    """
    safe_phrase = html.escape(phrase, quote=True)
    safe_mnemonic = html.escape(mnemonic, quote=True)
    message = quote(f"[{phrase}]", safe="")
    color = _KEYWORD_BLUE if lead else _KEYWORD_ORANGE
    src = (
        "https://img.shields.io/static/v1?style=flat-square"
        f"&label=&message={message}&color={color}"
    )
    return f'<img alt="[{safe_phrase}]" src="{src}" />({safe_mnemonic})'


def _top_level_group_spans(core: str) -> list[tuple[int, int]] | None:
    """Inclusive `[start, end]` of each top-level group, or None if not only groups."""
    spans: list[tuple[int, int]] = []
    outside: list[str] = []
    depth = 0
    open_at = -1
    for index, char in enumerate(core):
        if char == "[":
            if depth == 0:
                if "".join(outside).strip():
                    return None
                outside = []
                open_at = index
            depth += 1
        elif char == "]":
            if depth == 0:
                return None
            depth -= 1
            if depth == 0:
                if open_at < 0:
                    return None
                inner = core[open_at + 1 : index].strip()
                if not inner:
                    return None
                spans.append((open_at, index))
                open_at = -1
        elif depth == 0:
            outside.append(char)
    if depth or "".join(outside).strip():
        return None
    return spans


def _apply_span_replacements(text: str, reps: list[tuple[int, int, str]]) -> str:
    for start, end, chip in sorted(reps, key=lambda item: item[0], reverse=True):
        text = text[:start] + chip + text[end:]
    return text


def _embed_aligned(
    core: str,
    spans: list[tuple[int, int]],
    pairs: list[tuple[str, str]],
) -> str:
    reps: list[tuple[int, int, str]] = []
    for index, ((start, end), (mnemonic, term)) in enumerate(zip(spans, pairs)):
        inner = core[start + 1 : end]
        match = _find_term(inner, term)
        if match is None:
            continue
        lead = len(inner) - len(inner.lstrip())
        trail_end = len(inner.rstrip())
        chip = _keyword_chip_md(match.group(0), mnemonic, lead=index == 0)
        if match.start() == lead and match.end() == trail_end:
            reps.append((start, end + 1, chip))
        else:
            reps.append((start + 1 + match.start(), start + 1 + match.end(), chip))
    return _apply_span_replacements(core, reps)


def _embed_legacy(core: str, pairs: list[tuple[str, str]]) -> str:
    """First non-overlapping core match for each pair. Unmatched pairs are omitted."""
    claimed: list[tuple[int, int]] = []
    reps: list[tuple[int, int, str]] = []
    lead_chip = True
    for mnemonic, term in pairs:
        pattern = _term_pattern(term)
        if not pattern:
            continue
        for match in re.finditer(pattern, core):
            overlaps = any(
                match.start() < end and match.end() > start for start, end in claimed
            )
            if overlaps:
                continue
            claimed.append((match.start(), match.end()))
            reps.append(
                (
                    match.start(),
                    match.end(),
                    _keyword_chip_md(match.group(0), mnemonic, lead=lead_chip),
                )
            )
            lead_chip = False
            break
    return _apply_span_replacements(core, reps)


def embed_keyword_mnemonics(text: str, keywords: str | None) -> str:
    """Inline each keyword phrase. The first phrase is blue; the rest are orange.

    When pairs line up with bracket groups, each phrase is rewritten inside its
    own group and a phrase that fills the group reuses that group's brackets.
    Otherwise each pair wraps its first match in the core. Note text is unchanged.
    """
    if not text:
        return text
    pairs = iter_keyword_pairs(keywords)
    if not pairs:
        return text
    core, note = split_concept_note(text)
    spans = _top_level_group_spans(core)
    aligned = (
        spans is not None
        and len(spans) == len(pairs)
        and all(
            _find_term(core[start + 1 : end], term) is not None
            for (start, end), (_mnemonic, term) in zip(spans, pairs)
        )
    )
    if aligned and spans is not None:
        rewritten = _embed_aligned(core, spans, pairs)
    else:
        rewritten = _embed_legacy(core, pairs)
    return rewritten + note


PLANS_HEADERS = ["id", "study_id", "title", "sort_order", "active"]
PLAN_ITEMS_HEADERS = [
    "id",
    "plan_id",
    "section_path",
    "text",
    "checked",
    "sort_order",
]
PLAN_RESOURCES_HEADERS = ["id", "plan_id", "section_path", "line", "sort_order"]
ENTERTAINMENT_HEADERS = [
    "id",
    "parent_id",
    "title",
    "url",
    "notes",
    "done",
    "sort_order",
    "image",
    "score",
]

HEADERS = {
    "studies": STUDIES_HEADERS,
    "palaces": PALACES_HEADERS,
    "palace_images": PALACE_IMAGES_HEADERS,
    "study_images": STUDY_IMAGES_HEADERS,
    "beasts": BEASTS_HEADERS,
    "atoms": ATOMS_HEADERS,
    "plans": PLANS_HEADERS,
    "plan_items": PLAN_ITEMS_HEADERS,
    "plan_resources": PLAN_RESOURCES_HEADERS,
    "entertainment": ENTERTAINMENT_HEADERS,
}


def validate_beast_atoms(atoms_for_beast: list[dict[str, Any]]) -> str | None:
    singles = 0
    zoned = 0
    for a in atoms_for_beast:
        kind = (a.get("kind") or "single").strip().lower()
        if kind in ("zoned", "subtopic"):
            zoned += 1
        else:
            singles += 1
    if singles and zoned:
        return "Beast cannot mix a single Knowledge Atom with zoned Knowledge Atoms."
    if singles > 1:
        return "Beast may carry only one comprehensive Knowledge Atom."
    if zoned > 4:
        return "Beast may carry at most four zoned Knowledge Atoms (Z1–Z4)."
    return None


def validate_dataset(
    studies: list[dict[str, str]],
    palaces: list[dict[str, str]],
    beasts: list[dict[str, str]],
    atoms: list[dict[str, str]],
    palace_images: list[dict[str, str]] | None = None,
) -> list[str]:
    issues: list[str] = []
    study_ids = {s["id"] for s in studies if s.get("id")}
    palace_ids = {s["id"] for s in palaces if s.get("id")}
    beast_ids = {b["id"] for b in beasts if b.get("id")}

    for st in palaces:
        if st.get("study_id") not in study_ids:
            issues.append(
                f"Palace {st.get('id')} has unknown study_id {st.get('study_id')}"
            )
        if not (st.get("character_name") or "").strip():
            issues.append(f"Palace {st.get('id')} missing character_name")
        elif (st.get("character_name") or "").startswith("(unassigned"):
            issues.append(
                f"Palace {st.get('id')} has placeholder character (fill via Palaces module)"
            )

    chars_by_study: dict[str, set[str]] = {}
    for st in palaces:
        sid = st.get("study_id") or ""
        ch = (st.get("character_name") or "").strip()
        if not ch:
            continue
        chars_by_study.setdefault(sid, set())
        if ch in chars_by_study[sid]:
            issues.append(f"Duplicate character '{ch}' in study {sid}")
        chars_by_study[sid].add(ch)

    beasts_on_palace: dict[str, int] = {}
    for b in beasts:
        if b.get("palace_id") not in palace_ids:
            issues.append(
                f"Beast {b.get('id')} has unknown palace_id {b.get('palace_id')}"
            )
        palace = b.get("palace_id") or ""
        beasts_on_palace[palace] = beasts_on_palace.get(palace, 0) + 1
        peg = (b.get("peg_code") or "").strip()
        if peg.isdigit():
            issues.append(f"Beast {b.get('id')} has numeric peg_code")

    for palace, count in beasts_on_palace.items():
        if count > 5:
            issues.append(f"Palace {palace} has {count} beasts (max 5)")

    by_beast: dict[str, list[dict[str, str]]] = {}
    for a in atoms:
        bid = a.get("beast_id") or ""
        if bid not in beast_ids:
            issues.append(f"Atom {a.get('id')} has unknown beast_id {bid}")
        by_beast.setdefault(bid, []).append(a)

    for bid, group in by_beast.items():
        err = validate_beast_atoms(group)
        if err:
            issues.append(f"Beast {bid}: {err}")

    for img in palace_images or []:
        if img.get("palace_id") not in palace_ids:
            issues.append(
                f"Palace image {img.get('id')} has unknown palace_id {img.get('palace_id')}"
            )

    return issues


def ensure_data_dir(data_dir: Path) -> None:
    data_dir.mkdir(parents=True, exist_ok=True)
    (data_dir / "imported").mkdir(parents=True, exist_ok=True)
