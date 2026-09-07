from __future__ import annotations

import re
import unicodedata
from lugest_core.materials import (
    PROFILE_SERIES_META,
    detect_profile_designation as _detect_profile_designation,
    profile_mass_tables as _profile_mass_tables,
    profile_series as _profile_series,
)
from typing import Any


_STEEL_DENSITY_G_CM3 = 7.85


def _search_normalize(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.casefold()
    text = re.sub(r"(?<=\d),(?=\d)", ".", text)
    text = re.sub(r"[^a-z0-9.]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _search_terms(value: Any) -> list[str]:
    return [term for term in _search_normalize(value).split() if term]


def _commercial_thickness_base_mm(value: Any) -> float:
    text = str(value or "").strip().replace(",", ".")
    match = re.search(r"(\d+(?:\.\d+)?)\s*/\s*(\d+(?:\.\d+)?)", text)
    if not match:
        return 0.0
    try:
        return float(match.group(1))
    except Exception:
        return 0.0


def _commercial_thickness_label(value: Any) -> str:
    text = str(value or "").strip().replace(",", ".")
    match = re.search(r"(\d+(?:\.\d+)?)\s*/\s*(\d+(?:\.\d+)?)", text)
    if not match:
        return ""

    def _clean(part: str) -> str:
        return part.rstrip("0").rstrip(".") if "." in part else part

    left = _clean(match.group(1))
    right = _clean(match.group(2))
    return f"{left}/{right}"


_CHECKER_PLATE_WEIGHT_KG: dict[tuple[int, int], dict[str, float]] = {
    (2000, 1000): {
        "3/5": 51.1,
        "4/6": 66.8,
        "5/7": 82.5,
        "6/8": 98.2,
        "8/10": 129.6,
    },
    (2500, 1250): {
        "3/5": 79.84,
        "4/6": 104.38,
        "5/7": 128.91,
        "6/8": 153.44,
        "8/10": 202.5,
    },
    (3000, 1500): {
        "3/5": 112.23,
        "4/6": 147.55,
        "5/7": 182.88,
        "6/8": 218.2,
        "8/10": 288.85,
        "10/12": 360.0,
    },
}


def _checker_plate_weight_kg(length_mm: float, width_mm: float, thickness_label: str) -> float:
    dims = (int(round(float(length_mm or 0))), int(round(float(width_mm or 0))))
    label = _commercial_thickness_label(thickness_label) or str(thickness_label or "").strip()
    if not label:
        return 0.0
    exact = float(_CHECKER_PLATE_WEIGHT_KG.get(dims, {}).get(label, 0.0) or 0.0)
    if exact > 0:
        return exact
    area = max(0.0, float(length_mm or 0.0) * float(width_mm or 0.0))
    if area <= 0:
        return 0.0
    candidates: list[tuple[float, float]] = []
    target_ratio = max(float(length_mm or 0.0), float(width_mm or 0.0)) / max(1.0, min(float(length_mm or 0.0), float(width_mm or 0.0)))
    for (base_length, base_width), rows in _CHECKER_PLATE_WEIGHT_KG.items():
        base_weight = float(rows.get(label, 0.0) or 0.0)
        if base_weight <= 0:
            continue
        base_area = float(base_length * base_width)
        base_ratio = max(base_length, base_width) / max(1.0, min(base_length, base_width))
        score = abs(base_area - area) / max(area, base_area) + abs(base_ratio - target_ratio)
        candidates.append((score, base_weight / base_area))
    if not candidates:
        return 0.0
    _score, kg_per_mm2 = min(candidates, key=lambda item: item[0])
    return round(area * kg_per_mm2, 4)


_PROFILE_STANDARD_KG_M: dict[str, dict[str, float]] = {
    "IPN": {
        "80": 5.94,
        "100": 8.34,
        "120": 11.1,
        "140": 14.3,
        "160": 17.9,
        "180": 21.9,
        "200": 26.2,
        "220": 31.1,
        "240": 36.2,
        "260": 41.9,
        "280": 47.9,
        "300": 54.2,
        "320": 61.0,
    },
    "IPE": {
        "80": 6.0,
        "100": 8.1,
        "120": 10.4,
        "140": 12.9,
        "160": 15.8,
        "180": 18.8,
        "200": 22.4,
        "220": 26.2,
        "240": 30.7,
        "270": 36.1,
        "300": 42.2,
        "330": 49.1,
        "360": 57.1,
        "400": 66.3,
        "450": 77.6,
        "500": 90.7,
        "550": 106.0,
        "600": 122.0,
    },
    "HEA": {
        "100": 16.7,
        "120": 19.9,
        "140": 24.7,
        "160": 30.4,
        "180": 35.5,
        "200": 42.3,
        "220": 50.5,
        "240": 60.3,
        "260": 68.2,
        "280": 76.4,
        "300": 88.3,
        "320": 97.6,
        "340": 105.0,
        "360": 112.0,
        "400": 125.0,
        "450": 140.0,
        "500": 155.0,
        "550": 166.0,
        "600": 178.0,
        "650": 190.0,
        "700": 204.0,
        "800": 224.0,
        "900": 252.0,
        "1000": 272.0,
    },
    "HEB": {
        "100": 20.4,
        "120": 26.7,
        "140": 33.7,
        "160": 42.6,
        "180": 51.2,
        "200": 61.3,
        "220": 71.5,
        "240": 83.2,
        "260": 93.0,
        "280": 103.0,
        "300": 117.0,
        "320": 127.0,
        "340": 134.0,
        "360": 142.0,
        "400": 155.0,
        "450": 171.0,
        "500": 187.0,
        "550": 199.0,
        "600": 212.0,
        "650": 225.0,
        "700": 241.0,
        "800": 262.0,
        "900": 291.0,
        "1000": 314.0,
    },
    "HEM": {
        "100": 41.8,
        "120": 52.1,
        "140": 63.2,
        "160": 76.2,
        "180": 88.9,
        "200": 103.0,
        "220": 117.0,
        "240": 157.0,
        "260": 172.0,
        "280": 189.0,
        "300": 238.0,
        "320": 245.0,
        "340": 248.0,
        "360": 250.0,
        "400": 256.0,
        "450": 263.0,
        "500": 270.0,
        "550": 278.0,
        "600": 285.0,
        "650": 293.0,
        "700": 301.0,
        "800": 317.0,
        "900": 333.0,
        "1000": 349.0,
    },
    "UPN": {
        "80": 8.65,
        "100": 10.6,
        "120": 13.4,
        "140": 16.0,
        "160": 18.8,
        "180": 22.0,
        "200": 25.3,
        "220": 29.4,
        "240": 33.2,
        "260": 37.9,
        "280": 41.8,
        "300": 46.2,
        "320": 59.5,
        "400": 71.8,
    },
}


_PROFILE_STANDARD_KG_M = _profile_mass_tables()


_TUBE_SECTION_OPTIONS: list[dict[str, Any]] = [
    {"key": "quadrado", "label": "Quadrado"},
    {"key": "retangular", "label": "Retangular"},
    {"key": "redondo", "label": "Redondo"},
]


_PROFILE_SECTION_OPTIONS: list[dict[str, Any]] = [
    {
        "key": series,
        "label": str(PROFILE_SERIES_META.get(series, {}).get("label", series) or series),
        "catalog": True,
    }
    for series in _profile_series()
] + [
    {"key": "L", "label": "L (kg/m manual)", "catalog": False},
    {"key": "T", "label": "T (kg/m manual)", "catalog": False},
    {"key": "U", "label": "U genérico (kg/m manual)", "catalog": False},
    {"key": "I", "label": "I genérico (kg/m manual)", "catalog": False},
    {"key": "H", "label": "H genérico (kg/m manual)", "catalog": False},
    {"key": "OUTRO", "label": "Outro / manual", "catalog": False},
]


_ANGLE_SECTION_OPTIONS: list[dict[str, Any]] = [
    {"key": "abas_iguais", "label": "Abas iguais"},
    {"key": "abas_desiguais", "label": "Abas desiguais"},
]


_BAR_SECTION_OPTIONS: list[dict[str, Any]] = [
    {"key": "chata", "label": "Barra chata"},
    {"key": "quadrada", "label": "Barra quadrada"},
    {"key": "retangular", "label": "Barra retangular"},
]


def _profile_catalog_lookup_key(value: Any) -> str:
    token = str(value or "").strip().upper()
    return token if token in _PROFILE_STANDARD_KG_M else ""


def _profile_size_lookup_key(value: Any) -> str:
    text = str(value or "").strip().replace(",", ".")
    if not text:
        return ""
    match = re.search(r"[-+]?\d+(?:\.\d+)?", text)
    if not match:
        return ""
    try:
        number = float(match.group(0))
    except Exception:
        return ""
    if abs(number - round(number)) < 1e-6:
        return str(int(round(number)))
    return f"{number:.3f}".rstrip("0").rstrip(".")


def _detect_profile_catalog_from_text(value: Any) -> tuple[str, str]:
    return _detect_profile_designation(value)


class _ValueHolder:
    def __init__(self, value: str = "") -> None:
        self._value = str(value or "")

    def get(self) -> str:
        return self._value
