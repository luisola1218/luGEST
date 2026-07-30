from __future__ import annotations

import re
import unicodedata
from collections.abc import Iterable
from typing import Any


_KNOWN_UNITS = {
    "g",
    "kg",
    "m",
    "m2",
    "m3",
    "mm",
    "mm2",
    "mm3",
}


def search_normalize(value: Any) -> str:
    """Return a stable, accent-insensitive representation for user searches."""
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.casefold()
    text = re.sub(r"(?<=\d),(?=\d)", ".", text)
    text = re.sub(r"[^a-z0-9./]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def numeric_search_key(value: Any) -> str:
    text = search_normalize(value).strip().strip("./")
    if not re.fullmatch(r"\d+(?:\.\d+)?", text):
        return ""
    try:
        number = float(text)
    except (TypeError, ValueError):
        return ""
    return f"{number:.4f}".rstrip("0").rstrip(".")


def search_terms(value: Any) -> list[str]:
    """Split a query and understand compact measurements such as 6mm or 1,5kg."""
    raw_terms = [term for term in search_normalize(value).split() if term]
    terms: list[str] = []
    attached_unit = False
    for term in raw_terms:
        match = re.fullmatch(r"(\d+(?:\.\d+)?)(mm3|mm2|mm|kg|m3|m2|m|g)", term)
        if match:
            terms.append(numeric_search_key(match.group(1)) or match.group(1))
            attached_unit = True
            continue
        terms.append(term)

    # "6 mm" and "6mm" must behave identically. A unit entered by itself is
    # retained, so users can still deliberately search for unit labels.
    has_numeric = attached_unit or any(numeric_search_key(term) for term in terms)
    if has_numeric and len(terms) > 1:
        terms = [term for term in terms if term not in _KNOWN_UNITS]
    return list(dict.fromkeys(terms))


def _search_bucket(values: Iterable[Any]) -> tuple[str, set[str]]:
    normalized = search_normalize(" ".join(str(value or "") for value in values))
    tokens = set(normalized.split())
    expanded = set(tokens)
    for token in tokens:
        numeric = numeric_search_key(token)
        if numeric:
            expanded.add(numeric)
        if "/" in token:
            for part in token.split("/"):
                part = part.strip()
                if not part:
                    continue
                expanded.add(part)
                part_numeric = numeric_search_key(part)
                if part_numeric:
                    expanded.add(part_numeric)
    return normalized, expanded


def search_matches(values: Iterable[Any], query: Any) -> bool:
    """Match every query term, allowing each one to come from a different field."""
    terms = search_terms(query)
    if not terms:
        return True
    normalized, tokens = _search_bucket(values)
    for term in terms:
        numeric = numeric_search_key(term)
        if numeric:
            if numeric not in tokens and term not in tokens:
                return False
            continue
        if "/" in term:
            parts = [part for part in term.split("/") if part]
            if parts and all((numeric_search_key(part) or part) in tokens for part in parts):
                continue
        if term not in normalized:
            return False
    return True
