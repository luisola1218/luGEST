from __future__ import annotations

import re
import unicodedata
from difflib import SequenceMatcher
from typing import Any


def _fold(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(char for char in text if not unicodedata.combining(char))
    return re.sub(r"\s+", " ", text.casefold()).strip()


def _first(text: str, rules: tuple[tuple[str, tuple[str, ...]], ...]) -> str:
    for label, patterns in rules:
        if any(re.search(pattern, text) for pattern in patterns):
            return label
    return ""


def extract_product_attributes(description: str) -> dict[str, str]:
    """Extract commercial attributes without changing business data.

    Values are deliberately compact so they can be shown in the product
    inspector and indexed by future natural-language search.
    """

    text = _fold(description)
    metric = re.search(r"\bm\s*(\d+(?:[.,]\d+)?)(?:\s*[x×]\s*(\d+(?:[.,]\d+)?))?\b", text)
    commercial_size = re.search(
        r"\b(\d+(?:[.,]\d+)?)\s*[x×]\s*(\d+(?:[.,]\d+)?)\s*mm\b",
        text,
    )
    ral = re.search(r"\bral\s*[-:]?\s*(\d{3,4})\b", text)
    standard = re.search(r"\b(?:din|iso|en)\s*[-:]?\s*([a-z0-9.-]+)\b", text)
    tente_reference = re.search(r"\b(347[078])\s*(ufr)\s*(\d{3})\s*(p\d{2})\b", text)
    material = _first(text, (
        ("Inox A4 / AISI 316", (r"\ba4\b", r"\baisi\s*316", r"\binox\s*316")),
        ("Inox A2 / AISI 304", (r"\ba2\b", r"\baisi\s*304", r"\binox\s*304")),
        ("Inox", (r"\binox\b", r"\binoxidavel")),
        ("Alumínio", (r"\baluminio\b",)),
        ("Latão", (r"\blatao\b",)),
        ("Nylon", (r"\bnylon\b",)),
        ("Aço galvanizado", (r"\baco\s+galvan", r"\bgalvaniz")),
        ("Aço zincado", (r"\baco\s+zinc", r"\bzincad")),
        ("Aço", (r"\baco\b", r"\bs235", r"\bs275", r"\bs355")),
        ("PVC", (r"\bpvc\b",)),
        ("PTFE", (r"\bptfe\b", r"\bteflon\b")),
    ))
    drive = _first(text, (
        ("Torx", (r"\btorx\b", r"\btx\s*\d")),
        ("Pozidriv", (r"\bpozidriv\b", r"\bpz\s*\d")),
        ("Phillips/Cruz", (r"\bphillips\b", r"\bphilips\b", r"\bcruz\b", r"\bph\s*\d")),
        ("Allen/Umbrako", (r"\ballen\b", r"\bumbrak", r"\bunbrak", r"\bsextavad[oa]\s+interior")),
        ("Fenda", (r"\bfenda\b", r"\bslotted\b")),
        ("Sextavado exterior", (r"\bsextavad[oa]\b", r"\bhexagonal\b")),
    ))
    head = _first(text, (
        ("Cilíndrica", (r"\bcilindric",)),
        ("Escareada", (r"\bescaread",)),
        ("Sextavada", (r"\bcabeca\s+sextavad",)),
        ("Aba larga", (r"\baba\s+larga",)),
        ("Panela", (r"\bpanela\b",)),
    ))
    finish = _first(text, (
        ("Niquelado", (r"\bniquelad",)),
        ("Zincado", (r"\bzincad",)),
        ("Galvanizado", (r"\bgalvaniz",)),
        ("Preto", (r"\bpreto\b", r"\bblack\b")),
        ("Polido", (r"\bpolid",)),
        ("Escovado", (r"\bescovad",)),
    ))

    attributes: dict[str, str] = {}
    if metric:
        diameter = metric.group(1).replace(",", ".")
        length = str(metric.group(2) or "").replace(",", ".")
        attributes["medida"] = f"M{diameter}" + (f"x{length}" if length else "")
    elif commercial_size:
        first = commercial_size.group(1).replace(",", ".")
        second = commercial_size.group(2).replace(",", ".")
        attributes["medida"] = f"{first}x{second} mm"
    if re.search(r"\btente\b", text) or tente_reference:
        attributes["fabricante"] = "TENTE"
    if tente_reference:
        attributes["modelo"] = (
            f"{tente_reference.group(1)}"
            f"{tente_reference.group(2).upper()}"
            f"{tente_reference.group(3)}"
            f"{tente_reference.group(4).upper()}"
        )
        attributes["serie"] = "Alpha / UFR"
        attributes["travagem"] = (
            "Bloqueio total"
            if tente_reference.group(1) == "3477"
            else "Sem travão"
        )
        attributes["banda"] = "Borracha elástica não marcante"
    if material:
        attributes["material"] = material
    if drive:
        attributes["acionamento"] = drive
    if head:
        attributes["cabeca"] = head
    if finish:
        attributes["acabamento"] = finish
    color = _first(text, (
        ("Azul", (r"\bazul\b", r"\bblue\b")),
        ("Cinza", (r"\bcinz", r"\bgrey\b", r"\bgray\b")),
        ("Preto", (r"\bpreto\b", r"\bblack\b")),
        ("Branco", (r"\bbranco\b", r"\bwhite\b")),
        ("Vermelho", (r"\bvermelh", r"\bred\b")),
    ))
    if color:
        attributes["cor"] = color
    if standard:
        attributes["norma"] = standard.group(0).upper().replace("  ", " ")
    if ral:
        attributes["cor"] = f"RAL {ral.group(1)}"
    return attributes


def normalize_product_description(description: str) -> str:
    """Return a conservative, presentation-ready product designation."""

    value = re.sub(r"\s+", " ", str(description or "")).strip(" ,;")
    if not value:
        return ""
    if re.search(r"\bparafus", value, flags=re.IGNORECASE):
        value = re.sub(
            r"\bc/\s*(?=(?:cil[ií]ndric|sextavad|escaread|panela|aba))",
            "cabeça ",
            value,
            flags=re.IGNORECASE,
        )
    replacements = (
        (r"\bporcas?\b", "Porca"),
        (r"\bparafusos?\b", "Parafuso"),
        (r"\banilhas?\b", "Anilha"),
        (r"\brebites?\b", "Rebite"),
        (r"\bauto[\s-]*bloqueio\b", "autoblocante"),
        (r"\bauto[\s-]*travante\b", "autoblocante"),
        (r"\bnyloc\b", "Nyloc"),
        (r"\bumbrak(?:o)?\b", "Umbrako"),
        (r"\btorx\b", "Torx"),
        (r"\bphillips\b", "Phillips"),
        (r"\bpozidriv\b", "Pozidriv"),
        (r"\binox(?:idavel)?\b", "Inox"),
        (r"\baluminio\b", "Alumínio"),
        (r"\bc/\s*", "com "),
    )
    for pattern, replacement in replacements:
        value = re.sub(pattern, replacement, value, flags=re.IGNORECASE)
    value = re.sub(
        r"\bm\s*(\d+(?:[.,]\d+)?)(?:\s*[x×]\s*(\d+(?:[.,]\d+)?))?\b",
        lambda match: "M"
        + match.group(1).replace(",", ".")
        + (f"x{match.group(2).replace(',', '.')}" if match.group(2) else ""),
        value,
        flags=re.IGNORECASE,
    )
    value = re.sub(r"\s+", " ", value).strip()
    return value[:1].upper() + value[1:] if value else ""


def product_similarity(left: str, right: str) -> float:
    """Similarity tuned for duplicate product designations."""

    left_norm = _fold(normalize_product_description(left))
    right_norm = _fold(normalize_product_description(right))
    if not left_norm or not right_norm:
        return 0.0
    if left_norm == right_norm:
        return 1.0
    left_tokens = set(re.findall(r"[a-z0-9.]+", left_norm))
    right_tokens = set(re.findall(r"[a-z0-9.]+", right_norm))
    union = left_tokens | right_tokens
    token_score = len(left_tokens & right_tokens) / len(union) if union else 0.0
    sequence_score = SequenceMatcher(None, left_norm, right_norm).ratio()
    return round((token_score * 0.64) + (sequence_score * 0.36), 4)
