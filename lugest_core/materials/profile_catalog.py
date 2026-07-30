from __future__ import annotations

import re
from typing import Any


# European hot-rolled section catalogue. Dimensions are millimetres and mass is
# kg/m. The five complete ranges below mirror the technical tables supplied for
# the LuGEST material catalogue and are identified by their applicable European
# product/tolerance standards.
PROFILE_SERIES_META: dict[str, dict[str, str]] = {
    "IPE": {
        "label": "IPE — Viga I europeia (abas paralelas)",
        "standard": "EN 10365 / EN 10034",
    },
    "IPN": {
        "label": "IPN — Viga I europeia (abas inclinadas)",
        "standard": "EN 10365 / EN 10024",
    },
    "UPN": {
        "label": "UPN — Canal U europeu",
        "standard": "EN 10365 / EN 10279",
    },
    "HEA": {
        "label": "HEA — Viga H leve",
        "standard": "EN 10365 / EN 10034",
    },
    "HEB": {
        "label": "HEB — Viga H normal",
        "standard": "EN 10365 / EN 10034",
    },
    "HEM": {
        "label": "HEM — Viga H reforçada",
        "standard": "EN 10365 / EN 10034",
    },
    "UPE": {
        "label": "UPE — Canal U de abas paralelas",
        "standard": "EN 10365 / EN 10279",
    },
}


def _row(h: float, b: float, tw: float, tf: float, kg_m: float) -> dict[str, float]:
    return {"h": h, "b": b, "tw": tw, "tf": tf, "kg_m": kg_m}


PROFILE_CATALOG: dict[str, dict[str, dict[str, float]]] = {
    "IPE": {
        "80": _row(80, 46, 3.8, 5.2, 6.0),
        "100": _row(100, 55, 4.1, 5.7, 8.1),
        "120": _row(120, 64, 4.4, 6.3, 10.4),
        "140": _row(140, 73, 4.7, 6.9, 12.9),
        "160": _row(160, 82, 5.0, 7.4, 15.8),
        "180": _row(180, 91, 5.3, 8.0, 18.8),
        "200": _row(200, 100, 5.6, 8.5, 22.4),
        "220": _row(220, 110, 5.9, 9.2, 26.2),
        "240": _row(240, 120, 6.2, 9.8, 30.7),
        "270": _row(270, 135, 6.6, 10.2, 36.1),
        "300": _row(300, 150, 7.1, 10.7, 42.2),
        "330": _row(330, 160, 7.5, 11.5, 49.1),
        "360": _row(360, 170, 8.0, 12.7, 57.1),
        "400": _row(400, 180, 8.6, 13.5, 66.3),
        "450": _row(450, 190, 9.4, 14.6, 77.6),
        "500": _row(500, 200, 10.2, 16.0, 90.7),
        "550": _row(550, 210, 11.1, 17.2, 106.0),
        "600": _row(600, 220, 12.0, 19.0, 122.0),
    },
    "IPN": {
        "80": _row(80, 42, 3.9, 5.9, 5.95),
        "100": _row(100, 50, 4.5, 6.8, 8.32),
        "120": _row(120, 58, 5.1, 7.7, 11.2),
        "140": _row(140, 66, 5.7, 8.6, 14.4),
        "160": _row(160, 74, 6.3, 9.5, 17.9),
        "180": _row(180, 82, 6.9, 10.4, 21.9),
        "200": _row(200, 90, 7.5, 11.3, 26.3),
        "220": _row(220, 98, 8.1, 12.2, 31.1),
        "240": _row(240, 106, 8.7, 13.1, 36.2),
        "260": _row(260, 113, 9.4, 14.1, 41.9),
        "280": _row(280, 119, 10.1, 15.2, 48.0),
        "300": _row(300, 125, 10.8, 16.2, 54.2),
        "320": _row(320, 131, 11.5, 17.3, 61.1),
        "340": _row(340, 137, 12.2, 18.3, 68.1),
        "360": _row(360, 143, 13.0, 19.5, 76.2),
        "380": _row(380, 149, 13.7, 20.5, 84.0),
        "400": _row(400, 155, 14.4, 21.6, 92.6),
        "450": _row(450, 170, 16.2, 24.3, 115.0),
        "500": _row(500, 185, 18.0, 27.0, 141.0),
        "550": _row(550, 200, 19.0, 30.0, 167.0),
        "600": _row(600, 215, 21.6, 32.4, 199.0),
    },
    "UPN": {
        "80": _row(80, 45, 6.0, 4.0, 8.65),
        "100": _row(100, 50, 6.0, 4.5, 10.6),
        "120": _row(120, 55, 7.0, 4.5, 13.4),
        "140": _row(140, 60, 7.0, 5.0, 16.0),
        "160": _row(160, 65, 7.5, 5.5, 18.8),
        "180": _row(180, 70, 8.0, 5.5, 22.0),
        "200": _row(200, 75, 8.5, 6.0, 25.3),
        "220": _row(220, 80, 9.0, 6.5, 29.4),
        "240": _row(240, 85, 9.5, 6.5, 33.2),
        "260": _row(260, 90, 10.0, 7.0, 37.9),
        "280": _row(280, 95, 10.0, 7.5, 41.8),
        "300": _row(300, 100, 10.0, 8.0, 46.2),
        "320": _row(320, 100, 14.0, 8.8, 59.5),
        "350": _row(350, 100, 14.0, 8.0, 60.7),
        "380": _row(380, 102, 13.5, 8.0, 63.1),
        "400": _row(400, 110, 14.0, 9.0, 71.8),
    },
    "HEA": {
        "100": _row(96, 100, 5.0, 8.0, 16.7),
        "120": _row(114, 120, 5.0, 8.0, 19.9),
        "140": _row(133, 140, 5.5, 8.5, 24.7),
        "160": _row(152, 160, 6.0, 9.0, 30.4),
        "180": _row(171, 180, 6.0, 9.5, 35.5),
        "200": _row(190, 200, 6.5, 10.0, 42.3),
        "220": _row(210, 220, 7.0, 11.0, 50.5),
        "240": _row(230, 240, 7.5, 12.0, 60.3),
        "260": _row(250, 260, 7.5, 12.5, 68.2),
        "280": _row(270, 280, 8.0, 13.0, 76.4),
        "300": _row(290, 300, 8.5, 14.0, 88.3),
        "320": _row(310, 300, 9.0, 15.5, 97.6),
        "340": _row(330, 300, 9.5, 16.5, 105.0),
        "360": _row(350, 300, 10.0, 17.5, 112.0),
        "400": _row(390, 300, 11.0, 19.0, 125.0),
        "450": _row(440, 300, 11.5, 21.0, 140.0),
        "500": _row(490, 300, 12.0, 23.0, 155.0),
        "550": _row(540, 300, 12.5, 24.0, 166.0),
        "600": _row(590, 300, 13.0, 25.0, 178.0),
    },
    "HEB": {
        "100": _row(100, 100, 6.0, 10.0, 20.4),
        "120": _row(120, 120, 6.5, 11.0, 26.7),
        "140": _row(140, 140, 7.0, 12.0, 33.7),
        "160": _row(160, 160, 8.0, 13.0, 42.6),
        "180": _row(180, 180, 8.5, 14.0, 51.2),
        "200": _row(200, 200, 9.0, 15.0, 61.3),
        "220": _row(220, 220, 9.5, 16.0, 71.5),
        "240": _row(240, 240, 10.0, 17.0, 83.2),
        "260": _row(260, 260, 10.0, 17.5, 93.0),
        "280": _row(280, 280, 10.5, 18.0, 103.0),
        "300": _row(300, 300, 11.0, 19.0, 117.0),
        "320": _row(320, 300, 11.5, 20.5, 127.0),
        "340": _row(340, 300, 12.0, 21.5, 134.0),
        "360": _row(360, 300, 12.5, 22.5, 142.0),
        "400": _row(400, 300, 13.5, 24.0, 155.0),
        "450": _row(450, 300, 14.0, 26.0, 171.0),
        "500": _row(500, 300, 14.5, 28.0, 187.0),
        "550": _row(550, 300, 15.0, 29.0, 199.0),
        "600": _row(600, 300, 15.5, 30.0, 212.0),
    },
    # Existing mass-only ranges retained for backwards compatibility. The form
    # identifies them as mass-table entries until full dimensional data is added.
    "HEM": {
        size: {"kg_m": mass}
        for size, mass in {
            "100": 41.8, "120": 52.1, "140": 63.2, "160": 76.2,
            "180": 88.9, "200": 103.0, "220": 117.0, "240": 157.0,
            "260": 172.0, "280": 189.0, "300": 238.0, "320": 245.0,
            "340": 248.0, "360": 250.0, "400": 256.0, "450": 263.0,
            "500": 270.0, "550": 278.0, "600": 285.0, "650": 293.0,
            "700": 301.0, "800": 317.0, "900": 333.0, "1000": 349.0,
        }.items()
    },
    "UPE": {
        size: {"kg_m": mass}
        for size, mass in {
            "80": 7.93, "100": 9.82, "120": 12.1, "140": 14.5,
            "160": 17.0, "180": 19.7, "200": 22.8, "220": 26.6,
            "240": 30.2, "270": 35.2, "300": 40.5,
        }.items()
    },
}


def profile_series() -> list[str]:
    return list(PROFILE_CATALOG)


def profile_sizes(series: Any) -> list[str]:
    key = str(series or "").strip().upper()
    return sorted(
        PROFILE_CATALOG.get(key, {}),
        key=lambda value: float(value) if str(value).replace(".", "", 1).isdigit() else float("inf"),
    )


def profile_entry(series: Any, size: Any) -> dict[str, Any]:
    series_key = str(series or "").strip().upper()
    size_text = str(size or "").strip().replace(",", ".")
    match = re.search(r"\d+(?:\.\d+)?", size_text)
    if match:
        number = float(match.group(0))
        size_key = str(int(number)) if number.is_integer() else f"{number:g}"
    else:
        size_key = ""
    row = dict(PROFILE_CATALOG.get(series_key, {}).get(size_key, {}) or {})
    if not row:
        return {}
    meta = dict(PROFILE_SERIES_META.get(series_key, {}) or {})
    return {
        "series": series_key,
        "size": size_key,
        "designation": f"{series_key} {size_key}",
        "label": meta.get("label", series_key),
        "standard": meta.get("standard", "EN 10365"),
        **row,
    }


def profile_mass_tables() -> dict[str, dict[str, float]]:
    return {
        series: {
            size: float(row.get("kg_m", 0.0) or 0.0)
            for size, row in sizes.items()
        }
        for series, sizes in PROFILE_CATALOG.items()
    }


def detect_profile_designation(value: Any) -> tuple[str, str]:
    raw = str(value or "").upper()
    series_pattern = "|".join(sorted((re.escape(key) for key in PROFILE_CATALOG), key=len, reverse=True))
    match = re.search(rf"\b({series_pattern})\s*[- ]?\s*(\d{{2,4}})\b", raw)
    if not match:
        # Natural commands often place the steel grade between the series and
        # the nominal size: "IPE S275JR tamanho 220".
        match = re.search(
            rf"\b({series_pattern})\b.{{0,60}}?\b(?:TAMANHO|ALTURA|DIMENSAO|DIMENSÃO)\s*[:#=-]?\s*(\d{{2,4}})\b",
            raw,
        )
    if not match:
        return "", ""
    series, size = match.group(1), match.group(2)
    return (series, size) if profile_entry(series, size) else ("", "")
