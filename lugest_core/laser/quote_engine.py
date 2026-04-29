from __future__ import annotations

import copy
import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

from lugest_core.cad.profile_analysis import analyze_profile_cut_features

try:
    import ezdxf  # type: ignore[import-not-found]
except Exception:  # pragma: no cover - optional dependency
    ezdxf = None


DXF_BINARY_SIGNATURE = b"AutoCAD Binary DXF"
DXF_SNAP_TOLERANCE_MM = 0.05
MATERIAL_FAMILY_ALIASES = {
    "FERRO": "Aco carbono",
    "ACOCARBONO": "Aco carbono",
    "CARBONSTEEL": "Aco carbono",
    "INOX": "Aco inox",
    "ACOINOX": "Aco inox",
    "STAINLESS": "Aco inox",
}
MATERIAL_FAMILY_LABELS = {
    "Aco carbono": "Ferro",
    "Aco inox": "INOX",
}


def _as_float(value: Any, default: float = 0.0) -> float:
    try:
        return float(value)
    except Exception:
        return float(default)


def _as_int(value: Any, default: int = 0) -> int:
    try:
        return int(float(value))
    except Exception:
        return int(default)


def _round_mm(value: Any, digits: int = 3) -> float:
    return round(_as_float(value, 0.0), digits)


def _round_money(value: Any, digits: int = 2) -> float:
    return round(_as_float(value, 0.0), digits)


def _safe_div(numerator: float, denominator: float, default: float = 0.0) -> float:
    if abs(float(denominator or 0.0)) <= 1e-12:
        return float(default)
    return float(numerator or 0.0) / float(denominator or 0.0)


def _distance(a: tuple[float, float], b: tuple[float, float]) -> float:
    return math.hypot(float(b[0]) - float(a[0]), float(b[1]) - float(a[1]))


def _bbox_from_points(points: list[tuple[float, float]]) -> dict[str, float]:
    if not points:
        return {"min_x": 0.0, "min_y": 0.0, "max_x": 0.0, "max_y": 0.0, "width": 0.0, "height": 0.0}
    xs = [float(point[0]) for point in points]
    ys = [float(point[1]) for point in points]
    min_x = min(xs)
    min_y = min(ys)
    max_x = max(xs)
    max_y = max(ys)
    return {
        "min_x": round(min_x, 3),
        "min_y": round(min_y, 3),
        "max_x": round(max_x, 3),
        "max_y": round(max_y, 3),
        "width": round(max_x - min_x, 3),
        "height": round(max_y - min_y, 3),
    }


def _shoelace_area(points: list[tuple[float, float]]) -> float:
    if len(points) < 3:
        return 0.0
    raw = list(points)
    if raw[0] != raw[-1]:
        raw.append(raw[0])
    area = 0.0
    for index in range(len(raw) - 1):
        x1, y1 = raw[index]
        x2, y2 = raw[index + 1]
        area += (x1 * y2) - (x2 * y1)
    return area / 2.0


def _point_in_polygon(point: tuple[float, float], polygon: list[tuple[float, float]]) -> bool:
    if len(polygon) < 3:
        return False
    x, y = point
    bbox = _bbox_from_points(polygon)
    if x < bbox["min_x"] or x > bbox["max_x"] or y < bbox["min_y"] or y > bbox["max_y"]:
        return False
    inside = False
    raw = list(polygon)
    if raw[0] != raw[-1]:
        raw.append(raw[0])
    for index in range(len(raw) - 1):
        x1, y1 = raw[index]
        x2, y2 = raw[index + 1]
        if ((y1 > y) != (y2 > y)) and (x < ((x2 - x1) * (y - y1) / max(1e-12, (y2 - y1))) + x1):
            inside = not inside
    return inside


def _polygon_centroid(points: list[tuple[float, float]]) -> tuple[float, float]:
    if len(points) < 3:
        if points:
            bbox = _bbox_from_points(points)
            return ((bbox["min_x"] + bbox["max_x"]) / 2.0, (bbox["min_y"] + bbox["max_y"]) / 2.0)
        return (0.0, 0.0)
    raw = list(points)
    if raw[0] != raw[-1]:
        raw.append(raw[0])
    signed_area = _shoelace_area(raw)
    if abs(signed_area) <= 1e-12:
        bbox = _bbox_from_points(raw)
        return ((bbox["min_x"] + bbox["max_x"]) / 2.0, (bbox["min_y"] + bbox["max_y"]) / 2.0)
    cx = 0.0
    cy = 0.0
    for index in range(len(raw) - 1):
        x1, y1 = raw[index]
        x2, y2 = raw[index + 1]
        cross = (x1 * y2) - (x2 * y1)
        cx += (x1 + x2) * cross
        cy += (y1 + y2) * cross
    factor = 1.0 / (6.0 * signed_area)
    centroid = (cx * factor, cy * factor)
    if _point_in_polygon(centroid, raw):
        return centroid
    bbox = _bbox_from_points(raw)
    bbox_center = ((bbox["min_x"] + bbox["max_x"]) / 2.0, (bbox["min_y"] + bbox["max_y"]) / 2.0)
    if _point_in_polygon(bbox_center, raw):
        return bbox_center
    return centroid


def _sample_arc_points(
    center: tuple[float, float],
    radius: float,
    start_deg: float,
    end_deg: float,
    *,
    ccw: bool = True,
    max_step_deg: float = 7.5,
) -> list[tuple[float, float]]:
    cx, cy = center
    start = float(start_deg)
    end = float(end_deg)
    if ccw:
        while end <= start:
            end += 360.0
    else:
        while end >= start:
            end -= 360.0
    span = abs(end - start)
    steps = max(8, int(math.ceil(span / max(0.5, float(max_step_deg or 7.5)))))
    points: list[tuple[float, float]] = []
    for index in range(steps + 1):
        ratio = index / float(steps)
        angle_deg = start + ((end - start) * ratio)
        angle_rad = math.radians(angle_deg)
        points.append((cx + (radius * math.cos(angle_rad)), cy + (radius * math.sin(angle_rad))))
    return points


def _sample_circle_points(center: tuple[float, float], radius: float) -> list[tuple[float, float]]:
    points = _sample_arc_points(center, radius, 0.0, 360.0, ccw=True, max_step_deg=6.0)
    if points and points[0] != points[-1]:
        points.append(points[0])
    return points


def _sample_bulge_segment(
    start: tuple[float, float],
    end: tuple[float, float],
    bulge: float,
) -> list[tuple[float, float]]:
    if abs(float(bulge or 0.0)) <= 1e-9:
        return [start, end]
    x1, y1 = start
    x2, y2 = end
    chord = _distance(start, end)
    if chord <= 1e-9:
        return [start, end]
    radius = abs(chord * (1.0 + (bulge * bulge)) / (4.0 * float(bulge)))
    mid_x = (x1 + x2) / 2.0
    mid_y = (y1 + y2) / 2.0
    ux = (x2 - x1) / chord
    uy = (y2 - y1) / chord
    perp_x = -uy
    perp_y = ux
    half_chord = chord / 2.0
    offset = math.sqrt(max((radius * radius) - (half_chord * half_chord), 0.0))
    sign = 1.0 if bulge > 0 else -1.0
    cx = mid_x + (perp_x * offset * sign)
    cy = mid_y + (perp_y * offset * sign)
    start_angle = math.degrees(math.atan2(y1 - cy, x1 - cx))
    end_angle = math.degrees(math.atan2(y2 - cy, x2 - cx))
    return _sample_arc_points((cx, cy), radius, start_angle, end_angle, ccw=(bulge > 0), max_step_deg=6.0)


def _snap_key(point: tuple[float, float], tolerance_mm: float = DXF_SNAP_TOLERANCE_MM) -> tuple[int, int]:
    return (
        int(round(float(point[0]) / max(1e-9, tolerance_mm))),
        int(round(float(point[1]) / max(1e-9, tolerance_mm))),
    )


def _coerce_label(text: str) -> str:
    return str(text or "").strip().replace("_", " ")


def _norm_material_token(value: Any) -> str:
    return "".join(char for char in str(value or "").upper() if char.isalnum())


def _canonical_material_family(value: Any) -> str:
    raw = str(value or "").strip()
    if not raw:
        return ""
    token = _norm_material_token(raw)
    return MATERIAL_FAMILY_ALIASES.get(token, raw)


def _display_material_family(value: Any) -> str:
    canonical = _canonical_material_family(value)
    if not canonical:
        return ""
    return MATERIAL_FAMILY_LABELS.get(canonical, canonical)


def _profile_section_perimeter_mm(section: Any, family: Any = "") -> float:
    text = str(section or "").upper().replace(",", ".")
    values = [float(match) for match in re.findall(r"\d+(?:\.\d+)?", text)]
    family_norm = _norm_material_token(family)
    if not values:
        return 0.0
    if any(token in family_norm for token in ("CANTONEIRA", "ANGLE")) and len(values) >= 2:
        leg_a = values[0]
        leg_b = values[1]
        return max(0.0, 2.0 * (leg_a + leg_b))
    if any(token in text for token in ("RHS", "SHS")) or "X" in text:
        if len(values) >= 2:
            return max(0.0, 2.0 * (values[0] + values[1]))
    if any(token in text for token in ("CHS", "DN", "D", "Ø", "ROUND", "REDONDO")):
        return max(0.0, math.pi * values[0])
    if len(values) >= 2:
        return max(0.0, 2.0 * (values[0] + values[1]))
    return 0.0


def _profile_section_kg_m(section: Any, family: Any, thickness_mm: float, density_kg_m3: float) -> float:
    text = str(section or "").upper().replace(",", ".")
    values = [float(match) for match in re.findall(r"\d+(?:\.\d+)?", text)]
    family_norm = _norm_material_token(family)
    density_g_cm3 = max(0.001, float(density_kg_m3 or 7800.0) / 1000.0)
    thickness = max(0.0, float(thickness_mm or 0.0))
    if not values:
        return 0.0
    area_mm2 = 0.0
    if any(token in family_norm for token in ("TUBO", "TUBE")):
        if any(token in text for token in ("CHS", "DN", "D", "Ã˜", "Ø", "ROUND", "REDONDO")) and values:
            diameter = values[0]
            inner = max(0.0, diameter - (2.0 * thickness))
            if diameter > 0.0 and thickness > 0.0:
                area_mm2 = math.pi * ((diameter ** 2) - (inner ** 2)) / 4.0
        elif len(values) >= 2 and thickness > 0.0:
            width = values[0]
            height = values[1]
            inner_w = max(0.0, width - (2.0 * thickness))
            inner_h = max(0.0, height - (2.0 * thickness))
            area_mm2 = max(0.0, (width * height) - (inner_w * inner_h))
    elif any(token in family_norm for token in ("CANTONEIRA", "ANGLE")) and len(values) >= 2 and thickness > 0.0:
        leg_a = values[0]
        leg_b = values[1]
        area_mm2 = max(0.0, thickness * ((leg_a + leg_b) - thickness))
    elif any(token in family_norm for token in ("BARRA", "BAR")):
        if len(values) >= 2:
            area_mm2 = max(0.0, values[0] * values[1])
        elif len(values) == 1 and thickness > 0.0:
            area_mm2 = max(0.0, values[0] * thickness)
    return round((area_mm2 * density_g_cm3) / 1000.0, 4) if area_mm2 > 0.0 else 0.0


def _estimate_profile_cut_length_m(
    *,
    total_cut_count: int,
    hole_count: int,
    slot_count: int,
    outer_cut_count: int,
    thickness_mm: float,
    section: Any = "",
    family: Any = "",
) -> float:
    section_perimeter_mm = _profile_section_perimeter_mm(section, family)
    end_length_mm = max(0, int(outer_cut_count or 0)) * (section_perimeter_mm if section_perimeter_mm > 0 else max(30.0, thickness_mm * 12.0))
    hole_diameter_mm = max(6.0, thickness_mm * 2.0)
    hole_length_mm = max(0, int(hole_count or 0)) * math.pi * hole_diameter_mm
    slot_width_mm = max(6.0, thickness_mm * 2.0)
    slot_length_mm = max(18.0, thickness_mm * 8.0)
    slots_length_mm = max(0, int(slot_count or 0)) * (2.0 * (slot_length_mm + slot_width_mm))
    accounted = max(0, int(hole_count or 0)) + max(0, int(slot_count or 0)) + max(0, int(outer_cut_count or 0))
    generic_count = max(0, int(total_cut_count or 0) - accounted)
    generic_length_mm = generic_count * max(25.0, thickness_mm * 10.0)
    return round(max(0.0, end_length_mm + hole_length_mm + slots_length_mm + generic_length_mm) / 1000.0, 4)


def _infer_material_family_and_subtype(settings: dict[str, Any], material_value: Any, subtype_value: Any = "") -> tuple[str, str]:
    raw_material = str(material_value or "").strip()
    raw_subtype = str(subtype_value or "").strip()
    canonical_family = _canonical_material_family(raw_material)
    if canonical_family in MATERIAL_FAMILY_LABELS:
        return canonical_family, raw_subtype
    subtype_token = _norm_material_token(raw_subtype or raw_material)
    if not subtype_token:
        return canonical_family or raw_material, raw_subtype
    subtype_map = dict(settings.get("material_subtypes", {}) or {})
    for family, values in subtype_map.items():
        for value in list(values or []):
            clean = str(value or "").strip()
            if clean and _norm_material_token(clean) == subtype_token:
                return _canonical_material_family(family) or str(family or "").strip(), clean
    commercial_profiles = dict(settings.get("commercial_profiles", {}) or {})
    for profile in commercial_profiles.values():
        if not isinstance(profile, dict):
            continue
        catalog = dict(profile.get("material_catalog", {}) or {})
        for family, family_catalog in catalog.items():
            for key in dict(family_catalog or {}).keys():
                clean = str(key or "").strip()
                if clean and _norm_material_token(clean) == subtype_token:
                    return _canonical_material_family(family) or str(family or "").strip(), clean
    return canonical_family or raw_material, raw_subtype


def _sanitize_file_stem(path: str) -> str:
    stem = Path(str(path or "").strip()).stem.strip()
    if not stem:
        return "PECA-LASER"
    clean = []
    for char in stem.upper():
        clean.append(char if char.isalnum() or char in ("-", "_") else "-")
    text = "".join(clean).strip("-_")
    return text or "PECA-LASER"


def _resolve_cad_analysis_input(path: str | Path) -> tuple[Path, Path, list[str]]:
    source_path = Path(path).expanduser()
    if not source_path.exists():
        raise ValueError(f"Ficheiro CAD nao encontrado: {source_path}")
    suffix = source_path.suffix.lower()
    if suffix == ".dxf":
        return source_path, source_path, []
    if suffix != ".dwg":
        raise ValueError(f"Formato nao suportado: {source_path.suffix or '(sem extensao)'}. Seleciona DXF ou DWG.")

    sibling_candidates = [
        source_path.with_suffix(".dxf"),
        source_path.with_suffix(".DXF"),
    ]
    for candidate in sibling_candidates:
        if candidate.exists():
            return source_path, candidate, [f"DWG analisado via DXF associado: {candidate.name}."]

    oda_exec = _find_oda_file_converter()
    if oda_exec is not None:
        converted = _convert_dwg_with_oda(source_path, oda_exec)
        return source_path, converted, [f"DWG convertido automaticamente por ODA: {converted.name}."]

    command_template = str(os.environ.get("LUGEST_DWG_CONVERTER_CMD", "") or "").strip()
    if not command_template:
        raise ValueError(
            "Ficheiro DWG detetado, mas este posto nao tem conversor DWG->DXF configurado. "
            "Instala/configura o ODA File Converter ou define LUGEST_DWG_CONVERTER_CMD no ambiente."
        )

    temp_dir = Path(tempfile.mkdtemp(prefix="lugest_dwg_"))
    expected_output = temp_dir / f"{source_path.stem}.dxf"
    context = {
        "input": str(source_path),
        "output": str(expected_output),
        "input_dir": str(source_path.parent),
        "output_dir": str(temp_dir),
        "input_name": source_path.name,
        "output_name": expected_output.name,
        "input_stem": source_path.stem,
        "output_stem": expected_output.stem,
    }
    try:
        command = command_template.format(**context)
    except Exception as exc:
        raise ValueError(
            "A configuracao LUGEST_DWG_CONVERTER_CMD esta invalida. "
            "Usa placeholders como {input} e {output}."
        ) from exc
    completed = subprocess.run(command, shell=True, capture_output=True, text=True)
    generated_output = expected_output if expected_output.exists() else None
    if generated_output is None:
        generated = sorted(temp_dir.rglob("*.dxf"))
        if generated:
            generated_output = generated[0]
    if completed.returncode != 0 or generated_output is None or not generated_output.exists():
        detail = str((completed.stderr or completed.stdout or "")).strip()
        if detail:
            detail = f"\nDetalhe: {detail[:400]}"
        raise ValueError(
            "Falha ao converter DWG para DXF neste posto. "
            "Confirma o conversor configurado em LUGEST_DWG_CONVERTER_CMD." + detail
        )
    return source_path, generated_output, [f"DWG convertido para DXF automaticamente: {generated_output.name}."]


def _find_oda_file_converter() -> Path | None:
    env_candidate = str(os.environ.get("LUGEST_DWG_CONVERTER_EXE", "") or "").strip().strip('"')
    candidates: list[Path] = []
    if env_candidate:
        candidates.append(Path(env_candidate))
    base_dir = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parents[2]
    bundled_candidates = [
        base_dir / "tools" / "ODAFileConverter" / "ODAFileConverter.exe",
        base_dir / "tools" / "ODA" / "ODAFileConverter.exe",
        base_dir / "ODAFileConverter" / "ODAFileConverter.exe",
        base_dir / "third_party" / "ODAFileConverter" / "ODAFileConverter.exe",
    ]
    candidates.extend(bundled_candidates)
    which_path = shutil.which("ODAFileConverter.exe") or shutil.which("ODAFileConverter")
    if which_path:
        candidates.append(Path(which_path))
    for root_txt in (r"C:\Program Files\ODA", r"C:\Program Files (x86)\ODA"):
        root = Path(root_txt)
        if not root.exists():
            continue
        for entry in sorted(root.glob("ODAFileConverter*"), reverse=True):
            exe = entry / "ODAFileConverter.exe"
            if exe.is_file():
                candidates.append(exe)
    seen: set[str] = set()
    for candidate in candidates:
        normalized = str(candidate).strip().lower()
        if not normalized or normalized in seen:
            continue
        seen.add(normalized)
        if candidate.is_file():
            return candidate
    return None


def _convert_dwg_with_oda(source_path: Path, oda_exec: Path) -> Path:
    temp_dir = Path(tempfile.mkdtemp(prefix="lugest_dwg_oda_"))
    input_dir = temp_dir / "in"
    output_dir = temp_dir / "out"
    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)
    local_source = input_dir / source_path.name
    shutil.copy2(source_path, local_source)
    args = [
        str(oda_exec),
        str(input_dir),
        str(output_dir),
        "ACAD2018",
        "DXF",
        "0",
        "1",
        source_path.name,
    ]
    completed = subprocess.run(args, capture_output=True, text=True)
    generated = output_dir / f"{source_path.stem}.dxf"
    if not generated.exists():
        alt = sorted(output_dir.rglob("*.dxf"))
        if alt:
            generated = alt[0]
    if completed.returncode != 0 or not generated.exists():
        detail = str((completed.stderr or completed.stdout or "")).strip()
        if detail:
            detail = f"\nDetalhe: {detail[:400]}"
        raise ValueError("Falha ao converter DWG para DXF pelo ODA File Converter." + detail)
    return generated


def default_laser_quote_settings() -> dict[str, Any]:
    thickness_rows = [
        {"thickness_mm": 1, "speed_min_m_min": 9.0, "speed_max_m_min": 12.0, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 3.0, "gas_pressure_bar_max": 3.0, "focus_mm": 3.0, "nozzle": "1.2", "power_w": 1000},
        {"thickness_mm": 2, "speed_min_m_min": 5.0, "speed_max_m_min": 6.0, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 3.0, "gas_pressure_bar_max": 3.0, "focus_mm": 3.0, "nozzle": "1.2", "power_w": 2000},
        {"thickness_mm": 3, "speed_min_m_min": 3.6, "speed_max_m_min": 4.5, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 4.0, "gas_pressure_bar_max": 4.0, "focus_mm": 6.5, "nozzle": "1.2", "power_w": 3000},
        {"thickness_mm": 4, "speed_min_m_min": 3.3, "speed_max_m_min": 3.6, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 0.6, "gas_pressure_bar_max": 0.6, "focus_mm": 7.0, "nozzle": "1.2", "power_w": 3000},
        {"thickness_mm": 5, "speed_min_m_min": 3.0, "speed_max_m_min": 3.3, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 0.6, "gas_pressure_bar_max": 0.6, "focus_mm": 7.0, "nozzle": "1.2", "power_w": 4000},
        {"thickness_mm": 6, "speed_min_m_min": 2.8, "speed_max_m_min": 3.0, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 0.65, "gas_pressure_bar_max": 0.65, "focus_mm": 7.5, "nozzle": "1.2", "power_w": 4000},
        {"thickness_mm": 8, "speed_min_m_min": 2.4, "speed_max_m_min": 2.6, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 0.7, "gas_pressure_bar_max": 0.7, "focus_mm": 8.0, "nozzle": "1.2", "power_w": 5000},
        {"thickness_mm": 10, "speed_min_m_min": 1.9, "speed_max_m_min": 2.2, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 0.75, "gas_pressure_bar_max": 0.75, "focus_mm": 8.5, "nozzle": "1.2", "power_w": 6500},
        {"thickness_mm": 12, "speed_min_m_min": 1.8, "speed_max_m_min": 2.1, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 0.75, "gas_pressure_bar_max": 0.75, "focus_mm": 9.0, "nozzle": "1.2", "power_w": 7000},
        {"thickness_mm": 14, "speed_min_m_min": 1.7, "speed_max_m_min": 2.0, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 0.75, "gas_pressure_bar_max": 0.75, "focus_mm": 11.0, "nozzle": "1.4", "power_w": 8000},
        {"thickness_mm": 16, "speed_min_m_min": 1.6, "speed_max_m_min": 1.8, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 0.8, "gas_pressure_bar_max": 0.8, "focus_mm": 12.5, "nozzle": "1.4", "power_w": 9600},
        {"thickness_mm": 20, "speed_min_m_min": 1.3, "speed_max_m_min": 1.5, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 0.85, "gas_pressure_bar_max": 0.85, "focus_mm": 13.0, "nozzle": "1.4", "power_w": 12000},
        {"thickness_mm": 22, "speed_min_m_min": 1.2, "speed_max_m_min": 1.4, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 0.85, "gas_pressure_bar_max": 0.85, "focus_mm": 13.0, "nozzle": "1.6", "power_w": 12000},
        {"thickness_mm": 25, "speed_min_m_min": 1.0, "speed_max_m_min": 1.3, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 0.8, "gas_pressure_bar_max": 1.0, "focus_mm": 14.0, "nozzle": "1.6", "power_w": 12000},
        {"thickness_mm": 30, "speed_min_m_min": 0.5, "speed_max_m_min": 0.8, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 0.8, "gas_pressure_bar_max": 1.1, "focus_mm": 16.0, "nozzle": "1.6", "power_w": 12000},
        {"thickness_mm": 35, "speed_min_m_min": 0.4, "speed_max_m_min": 0.6, "nozzle_distance_mm": 1.2, "gas_pressure_bar_min": 0.8, "gas_pressure_bar_max": 1.2, "focus_mm": 16.5, "nozzle": "1.6/1.8", "power_w": 12000},
        {"thickness_mm": 40, "speed_min_m_min": 0.2, "speed_max_m_min": 0.4, "nozzle_distance_mm": 1.2, "gas_pressure_bar_min": 0.8, "gas_pressure_bar_max": 1.3, "focus_mm": 17.0, "nozzle": "1.6/1.8", "power_w": 12000},
        {"thickness_mm": 45, "speed_min_m_min": 0.1, "speed_max_m_min": 0.3, "nozzle_distance_mm": 1.2, "gas_pressure_bar_min": 0.8, "gas_pressure_bar_max": 1.3, "focus_mm": 17.0, "nozzle": "1.6/1.8", "power_w": 12000},
        {"thickness_mm": 50, "speed_min_m_min": 0.1, "speed_max_m_min": 0.2, "nozzle_distance_mm": 1.2, "gas_pressure_bar_min": 0.8, "gas_pressure_bar_max": 1.4, "focus_mm": 17.5, "nozzle": "1.6/1.8", "power_w": 12000},
    ]
    inox_nitrogen_rows = [
        {"thickness_mm": 1, "speed_min_m_min": 30.0, "speed_max_m_min": 50.0, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 6.0, "gas_pressure_bar_max": 6.0, "focus_mm": 0.0, "nozzle": "3.0", "power_w": 12000, "duty_pct": 100.0, "frequency_hz": 5000.0},
        {"thickness_mm": 2, "speed_min_m_min": 30.0, "speed_max_m_min": 38.0, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 6.0, "gas_pressure_bar_max": 6.0, "focus_mm": -0.5, "nozzle": "3.0", "power_w": 12000, "duty_pct": 100.0, "frequency_hz": 5000.0},
        {"thickness_mm": 3, "speed_min_m_min": 23.0, "speed_max_m_min": 30.0, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 6.0, "gas_pressure_bar_max": 6.0, "focus_mm": -0.5, "nozzle": "3.0", "power_w": 12000, "duty_pct": 100.0, "frequency_hz": 5000.0},
        {"thickness_mm": 4, "speed_min_m_min": 20.0, "speed_max_m_min": 25.0, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 6.0, "gas_pressure_bar_max": 6.0, "focus_mm": -1.0, "nozzle": "3.0", "power_w": 12000, "duty_pct": 100.0, "frequency_hz": 5000.0},
        {"thickness_mm": 5, "speed_min_m_min": 14.0, "speed_max_m_min": 18.0, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 6.0, "gas_pressure_bar_max": 6.0, "focus_mm": -1.5, "nozzle": "3.0", "power_w": 12000, "duty_pct": 100.0, "frequency_hz": 5000.0},
        {"thickness_mm": 6, "speed_min_m_min": 12.0, "speed_max_m_min": 15.0, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 6.0, "gas_pressure_bar_max": 6.0, "focus_mm": -2.0, "nozzle": "3.0", "power_w": 12000, "duty_pct": 100.0, "frequency_hz": 5000.0},
        {"thickness_mm": 8, "speed_min_m_min": 9.0, "speed_max_m_min": 11.0, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 6.0, "gas_pressure_bar_max": 6.0, "focus_mm": -3.0, "nozzle": "3.0", "power_w": 12000, "duty_pct": 100.0, "frequency_hz": 5000.0},
        {"thickness_mm": 10, "speed_min_m_min": 6.0, "speed_max_m_min": 8.0, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 6.0, "gas_pressure_bar_max": 6.0, "focus_mm": -4.0, "nozzle": "5.0", "power_w": 12000, "duty_pct": 100.0, "frequency_hz": 5000.0},
        {"thickness_mm": 12, "speed_min_m_min": 4.5, "speed_max_m_min": 5.5, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 6.0, "gas_pressure_bar_max": 6.0, "focus_mm": -5.0, "nozzle": "5.0", "power_w": 12000, "duty_pct": 100.0, "frequency_hz": 5000.0},
        {"thickness_mm": 14, "speed_min_m_min": 3.0, "speed_max_m_min": 3.8, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 6.0, "gas_pressure_bar_max": 6.0, "focus_mm": -7.5, "nozzle": "5.0", "power_w": 12000, "duty_pct": 100.0, "frequency_hz": 5000.0},
        {"thickness_mm": 16, "speed_min_m_min": 2.2, "speed_max_m_min": 2.6, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 7.0, "gas_pressure_bar_max": 7.0, "focus_mm": -9.0, "nozzle": "7.0", "power_w": 12000, "duty_pct": 100.0, "frequency_hz": 5000.0},
        {"thickness_mm": 20, "speed_min_m_min": 1.6, "speed_max_m_min": 1.8, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 7.0, "gas_pressure_bar_max": 7.0, "focus_mm": -12.0, "nozzle": "7.0", "power_w": 12000, "duty_pct": 100.0, "frequency_hz": 5000.0},
        {"thickness_mm": 25, "speed_min_m_min": 0.4, "speed_max_m_min": 1.0, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 7.0, "gas_pressure_bar_max": 7.0, "focus_mm": -15.0, "nozzle": "7.0", "power_w": 12000, "duty_pct": 100.0, "frequency_hz": 5000.0},
        {"thickness_mm": 30, "speed_min_m_min": 0.35, "speed_max_m_min": 0.6, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 7.0, "gas_pressure_bar_max": 7.0, "focus_mm": -17.0, "nozzle": "7.0", "power_w": 12000, "duty_pct": 100.0, "frequency_hz": 5000.0},
        {"thickness_mm": 35, "speed_min_m_min": 0.2, "speed_max_m_min": 0.4, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 7.0, "gas_pressure_bar_max": 7.0, "focus_mm": -20.0, "nozzle": "7.0", "power_w": 12000, "duty_pct": 65.0, "frequency_hz": 200.0},
        {"thickness_mm": 40, "speed_min_m_min": 0.15, "speed_max_m_min": 0.3, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 7.0, "gas_pressure_bar_max": 7.0, "focus_mm": -25.0, "nozzle": "7.0", "power_w": 12000, "duty_pct": 60.0, "frequency_hz": 275.0},
        {"thickness_mm": 45, "speed_min_m_min": 0.1, "speed_max_m_min": 0.2, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 7.0, "gas_pressure_bar_max": 7.0, "focus_mm": -30.0, "nozzle": "7.0", "power_w": 12000, "duty_pct": 75.0, "frequency_hz": 275.0},
        {"thickness_mm": 50, "speed_min_m_min": 0.1, "speed_max_m_min": 0.2, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 7.0, "gas_pressure_bar_max": 7.0, "focus_mm": -34.0, "nozzle": "7.0", "power_w": 12000, "duty_pct": 55.0, "frequency_hz": 125.0},
    ]
    inox_air_rows = [
        {"thickness_mm": 1, "speed_min_m_min": 30.0, "speed_max_m_min": 50.0, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 6.0, "gas_pressure_bar_max": 6.0, "focus_mm": 0.0, "nozzle": "3.0", "power_w": 12000, "duty_pct": 100.0, "frequency_hz": 5000.0},
        {"thickness_mm": 2, "speed_min_m_min": 30.0, "speed_max_m_min": 38.0, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 6.0, "gas_pressure_bar_max": 6.0, "focus_mm": -0.5, "nozzle": "3.0", "power_w": 12000, "duty_pct": 100.0, "frequency_hz": 5000.0},
        {"thickness_mm": 3, "speed_min_m_min": 24.0, "speed_max_m_min": 32.0, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 6.0, "gas_pressure_bar_max": 6.0, "focus_mm": -0.5, "nozzle": "3.0", "power_w": 12000, "duty_pct": 100.0, "frequency_hz": 5000.0},
        {"thickness_mm": 4, "speed_min_m_min": 22.0, "speed_max_m_min": 26.0, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 6.0, "gas_pressure_bar_max": 6.0, "focus_mm": -1.0, "nozzle": "3.0", "power_w": 12000, "duty_pct": 100.0, "frequency_hz": 5000.0},
        {"thickness_mm": 5, "speed_min_m_min": 16.0, "speed_max_m_min": 20.0, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 6.0, "gas_pressure_bar_max": 6.0, "focus_mm": -1.5, "nozzle": "3.0", "power_w": 12000, "duty_pct": 100.0, "frequency_hz": 5000.0},
        {"thickness_mm": 6, "speed_min_m_min": 12.5, "speed_max_m_min": 15.0, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 6.0, "gas_pressure_bar_max": 6.0, "focus_mm": -2.0, "nozzle": "3.0", "power_w": 12000, "duty_pct": 100.0, "frequency_hz": 5000.0},
        {"thickness_mm": 8, "speed_min_m_min": 10.0, "speed_max_m_min": 12.0, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 6.0, "gas_pressure_bar_max": 6.0, "focus_mm": -3.0, "nozzle": "3.0", "power_w": 12000, "duty_pct": 100.0, "frequency_hz": 5000.0},
        {"thickness_mm": 10, "speed_min_m_min": 6.5, "speed_max_m_min": 8.0, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 6.0, "gas_pressure_bar_max": 6.0, "focus_mm": -4.0, "nozzle": "5.0", "power_w": 12000, "duty_pct": 100.0, "frequency_hz": 5000.0},
        {"thickness_mm": 12, "speed_min_m_min": 4.6, "speed_max_m_min": 5.6, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 6.0, "gas_pressure_bar_max": 6.0, "focus_mm": -5.0, "nozzle": "5.0", "power_w": 12000, "duty_pct": 100.0, "frequency_hz": 5000.0},
        {"thickness_mm": 14, "speed_min_m_min": 3.2, "speed_max_m_min": 4.0, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 6.0, "gas_pressure_bar_max": 6.0, "focus_mm": -7.5, "nozzle": "5.0", "power_w": 12000, "duty_pct": 100.0, "frequency_hz": 5000.0},
        {"thickness_mm": 16, "speed_min_m_min": 2.3, "speed_max_m_min": 2.8, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 7.0, "gas_pressure_bar_max": 7.0, "focus_mm": -9.0, "nozzle": "7.0", "power_w": 12000, "duty_pct": 100.0, "frequency_hz": 5000.0},
        {"thickness_mm": 20, "speed_min_m_min": 1.6, "speed_max_m_min": 2.0, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 7.0, "gas_pressure_bar_max": 7.0, "focus_mm": -12.0, "nozzle": "7.0", "power_w": 12000, "duty_pct": 100.0, "frequency_hz": 5000.0},
        {"thickness_mm": 25, "speed_min_m_min": 0.7, "speed_max_m_min": 1.1, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 7.0, "gas_pressure_bar_max": 7.0, "focus_mm": -15.0, "nozzle": "7.0", "power_w": 12000, "duty_pct": 100.0, "frequency_hz": 5000.0},
        {"thickness_mm": 30, "speed_min_m_min": 0.5, "speed_max_m_min": 0.6, "nozzle_distance_mm": 0.8, "gas_pressure_bar_min": 7.0, "gas_pressure_bar_max": 7.0, "focus_mm": -17.0, "nozzle": "7.0", "power_w": 12000, "duty_pct": 100.0, "frequency_hz": 5000.0},
    ]
    return {
        "active_machine": "BLT641 12kW",
        "active_commercial": "Laser Oficina",
        "layer_rules": {
            "mark_patterns": ["mark", "scribe", "engrave", "grav", "texto", "text", "txt"],
            "ignore_patterns": ["fold", "bend", "guide", "construction", "dim", "assist", "layer0-ignore"],
        },
        "nesting": {
            "default_part_spacing_mm": 8.0,
            "default_edge_margin_mm": 8.0,
            "allow_rotate": True,
            "auto_select_sheet": False,
            "use_stock_first": False,
            "allow_purchase_fallback": True,
            "shape_aware": True,
            "shape_grid_mm": 10.0,
            "common_line_estimate": True,
            "common_line_tolerance_mm": 1.0,
            "lead_optimization": True,
            "lead_optimization_pct": 8.0,
            "sheet_profiles": [
                {"name": "1000 x 2000", "width_mm": 1000.0, "height_mm": 2000.0},
                {"name": "1250 x 2500", "width_mm": 1250.0, "height_mm": 2500.0},
                {"name": "1500 x 3000", "width_mm": 1500.0, "height_mm": 3000.0},
                {"name": "2000 x 4000", "width_mm": 2000.0, "height_mm": 4000.0},
            ],
        },
        "machine_profiles": {
            "BLT641 12kW": {
                "name": "BLT641 12kW",
                "motion": {
                    "rapid_speed_mm_s": 200.0,
                    "travel_acc_mm_s2": 2000.0,
                    "cut_acc_mm_s2": 2000.0,
                    "mark_speed_m_min": 18.0,
                    "effective_speed_factor_pct": 92.0,
                    "lead_in_mm": 2.0,
                    "lead_out_mm": 2.0,
                    "lead_move_speed_mm_s": 3.0,
                    "pierce_base_ms": 400.0,
                    "pierce_per_mm_ms": 35.0,
                    "first_gas_delay_ms": 200.0,
                    "gas_delay_ms": 0.0,
                    "motion_overhead_pct": 4.0,
                },
                "materials": {
                    "Aco carbono": {
                        "density_kg_m3": 7800.0,
                        "default_gas": "Oxigenio",
                        "gases": {
                            "Oxigenio": {
                                "rows": thickness_rows,
                            }
                        },
                    },
                    "Aco inox": {
                        "density_kg_m3": 7930.0,
                        "default_gas": "Azoto",
                        "gases": {
                            "Azoto": {
                                "rows": inox_nitrogen_rows,
                            },
                            "Ar comprimido": {
                                "rows": inox_air_rows,
                            },
                        },
                    }
                },
            }
        },
        "commercial_profiles": {
            "Laser Oficina": {
                "name": "Laser Oficina",
                "currency": "EUR",
                "cost_mode": "hybrid_max",
                "minimum_line_eur": 0.0,
                "margin_pct": 18.0,
                "setup_time_min": 3.0,
                "handling_eur": 0.0,
                "include_profile_event_rates": False,
                "profile_reference_thickness_mm": 3.0,
                "profile_min_thickness_rate_factor": 0.45,
                "profile_max_thickness_rate_factor": 8.0,
                "material_utilization_pct": 82.0,
                "fallback_fill_pct": 72.0,
                "use_scrap_credit": True,
                "series_pricing": {
                    "tiers": [
                        {"key": "single", "label": "Peca unica", "qty_min": 1, "qty_max": 1, "margin_delta_pct": 6.0, "setup_multiplier": 1.6},
                        {"key": "small", "label": "Pequena serie", "qty_min": 2, "qty_max": 9, "margin_delta_pct": 2.0, "setup_multiplier": 1.15},
                        {"key": "medium", "label": "Media serie", "qty_min": 10, "qty_max": 49, "margin_delta_pct": 0.0, "setup_multiplier": 1.0},
                        {"key": "large", "label": "Grande serie", "qty_min": 50, "qty_max": 999999, "margin_delta_pct": -3.0, "setup_multiplier": 0.8},
                    ]
                },
                "materials": {
                    "Aco carbono": {
                        "density_kg_m3": 7800.0,
                        "price_per_kg": 1.30,
                        "scrap_credit_per_kg": 1.20,
                    },
                    "Aco inox": {
                        "density_kg_m3": 7930.0,
                        "price_per_kg": 0.0,
                        "scrap_credit_per_kg": 0.0,
                    }
                },
                "material_catalog": {
                    "Aco carbono": {
                        "S235JR": {"price_per_kg": 1.30, "scrap_credit_per_kg": 1.20},
                        "S275JR": {"price_per_kg": 1.30, "scrap_credit_per_kg": 1.20},
                        "S355JR": {"price_per_kg": 1.30, "scrap_credit_per_kg": 1.20},
                        "S355J2+N": {"price_per_kg": 1.30, "scrap_credit_per_kg": 1.20},
                        "S355MC": {"price_per_kg": 1.30, "scrap_credit_per_kg": 1.20},
                        "S420MC": {"price_per_kg": 1.30, "scrap_credit_per_kg": 1.20},
                        "DX51D+Z": {"price_per_kg": 1.30, "scrap_credit_per_kg": 1.20},
                        "DX53D+Z": {"price_per_kg": 1.30, "scrap_credit_per_kg": 1.20},
                        "DD11": {"price_per_kg": 1.30, "scrap_credit_per_kg": 1.20},
                        "DC01": {"price_per_kg": 1.30, "scrap_credit_per_kg": 1.20},
                        "CORTEN": {"price_per_kg": 1.30, "scrap_credit_per_kg": 1.20},
                        "S355JOW": {"price_per_kg": 1.30, "scrap_credit_per_kg": 1.20},
                        "HARDOX 400": {"price_per_kg": 1.30, "scrap_credit_per_kg": 1.20},
                        "HARDOX 450": {"price_per_kg": 1.30, "scrap_credit_per_kg": 1.20},
                    },
                    "Aco inox": {
                        "INOX304 2B": {"price_per_kg": 0.0, "scrap_credit_per_kg": 0.0},
                        "INOX304 BA": {"price_per_kg": 0.0, "scrap_credit_per_kg": 0.0},
                        "INOX304 Escovado": {"price_per_kg": 0.0, "scrap_credit_per_kg": 0.0},
                        "INOX304L 2B": {"price_per_kg": 0.0, "scrap_credit_per_kg": 0.0},
                        "INOX316 2B": {"price_per_kg": 0.0, "scrap_credit_per_kg": 0.0},
                        "INOX316L 2B": {"price_per_kg": 0.0, "scrap_credit_per_kg": 0.0},
                        "INOX316L Escovado": {"price_per_kg": 0.0, "scrap_credit_per_kg": 0.0},
                        "INOX430 2B": {"price_per_kg": 0.0, "scrap_credit_per_kg": 0.0},
                        "INOX430 BA": {"price_per_kg": 0.0, "scrap_credit_per_kg": 0.0},
                        "AISI304": {"price_per_kg": 0.0, "scrap_credit_per_kg": 0.0},
                        "AISI304L": {"price_per_kg": 0.0, "scrap_credit_per_kg": 0.0},
                        "AISI316": {"price_per_kg": 0.0, "scrap_credit_per_kg": 0.0},
                        "AISI316L": {"price_per_kg": 0.0, "scrap_credit_per_kg": 0.0},
                        "1.4301": {"price_per_kg": 0.0, "scrap_credit_per_kg": 0.0},
                        "1.4307": {"price_per_kg": 0.0, "scrap_credit_per_kg": 0.0},
                        "1.4401": {"price_per_kg": 0.0, "scrap_credit_per_kg": 0.0},
                        "1.4404": {"price_per_kg": 0.0, "scrap_credit_per_kg": 0.0},
                        "1.4016": {"price_per_kg": 0.0, "scrap_credit_per_kg": 0.0},
                    },
                },
                "rates": {
                    "cut_per_m_eur": 1.50,
                    "defilm_per_m_eur": 0.50,
                    "marking_per_m_eur": 0.66,
                    "pierce_eur": 0.15,
                    "machine_hour_eur": 90.0,
                },
                "profile_operations": {
                    "outer_cut_eur": 1.25,
                    "hole_cut_eur": 0.85,
                    "slot_cut_eur": 1.15,
                },
            }
        },
        "material_subtypes": {
            "Aco carbono": [
                "S235JR",
                "S275JR",
                "S355JR",
                "S355J2+N",
                "S355MC",
                "S420MC",
                "DX51D+Z",
                "DX53D+Z",
                "DD11",
                "DC01",
                "CORTEN",
                "S355JOW",
                "HARDOX 400",
                "HARDOX 450",
            ],
            "Aco inox": [
                "INOX304 2B",
                "INOX304 BA",
                "INOX304 Escovado",
                "INOX304L 2B",
                "INOX316 2B",
                "INOX316L 2B",
                "INOX316L Escovado",
                "INOX430 2B",
                "INOX430 BA",
                "AISI304",
                "AISI304L",
                "AISI316",
                "AISI316L",
                "1.4301",
                "1.4307",
                "1.4401",
                "1.4404",
                "1.4016",
            ],
        },
    }


def merge_laser_quote_settings(stored: dict[str, Any] | None = None) -> dict[str, Any]:
    defaults = default_laser_quote_settings()
    base = copy.deepcopy(defaults)
    incoming = copy.deepcopy(dict(stored or {}))

    def merge_dict(target: dict[str, Any], source: dict[str, Any]) -> dict[str, Any]:
        for key, value in list(source.items()):
            if isinstance(value, dict) and isinstance(target.get(key), dict):
                merge_dict(target[key], value)
            else:
                target[key] = copy.deepcopy(value)
        return target

    merge_dict(base, incoming)
    default_profiles = dict(defaults.get("commercial_profiles", {}) or {})
    for profile_name, profile in dict(base.get("commercial_profiles", {}) or {}).items():
        if not isinstance(profile, dict):
            continue
        default_profile = dict(default_profiles.get(profile_name, {}) or {})
        default_catalog = dict(default_profile.get("material_catalog", {}) or {})
        current_catalog = dict(profile.get("material_catalog", {}) or {})
        hidden_catalog = {
            str(family or "").strip(): [str(item or "").strip() for item in list(values or []) if str(item or "").strip()]
            for family, values in dict(profile.get("material_catalog_hidden", {}) or {}).items()
            if str(family or "").strip()
        }
        merged_catalog: dict[str, Any] = {}
        for family in list(default_catalog.keys()) + [key for key in current_catalog.keys() if key not in default_catalog]:
            default_family = dict(default_catalog.get(family, {}) or {})
            current_family = dict(current_catalog.get(family, {}) or {})
            merged_family = dict(default_family)
            merged_family.update(current_family)
            hidden_items = set(hidden_catalog.get(family, []) or [])
            if hidden_items:
                merged_family = {
                    str(key or "").strip(): dict(value or {})
                    for key, value in merged_family.items()
                    if str(key or "").strip() not in hidden_items
                }
            merged_catalog[family] = merged_family
        profile["material_catalog"] = merged_catalog
    default_subtypes = dict(defaults.get("material_subtypes", {}) or {})
    current_subtypes = dict(base.get("material_subtypes", {}) or {})
    hidden_subtypes_cfg = dict(base.get("material_subtypes_hidden", {}) or {})
    merged_subtypes: dict[str, list[str]] = {}
    for family in list(default_subtypes.keys()) + [key for key in current_subtypes.keys() if key not in default_subtypes]:
        hidden_values = {
            str(value or "").strip()
            for value in list(hidden_subtypes_cfg.get(family, []) or [])
            if str(value or "").strip()
        }
        values: list[str] = []
        for value in list(default_subtypes.get(family, []) or []) + list(current_subtypes.get(family, []) or []):
            clean = str(value or "").strip()
            if clean and clean not in hidden_values and clean not in values:
                values.append(clean)
        merged_subtypes[family] = values
    base["material_subtypes"] = merged_subtypes
    for profile in dict(base.get("commercial_profiles", {}) or {}).values():
        if isinstance(profile, dict):
            profile["minimum_line_eur"] = 0.0
    return base


def _iter_pairs(path: str | Path) -> list[tuple[int, str]]:
    payload = Path(path).read_bytes()
    if payload.startswith(DXF_BINARY_SIGNATURE):
        raise ValueError("DXF binario nao suportado nesta versao. Exporta em DXF ASCII.")
    try:
        text = payload.decode("utf-8")
    except Exception:
        text = payload.decode("latin-1", errors="ignore")
    lines = text.splitlines()
    if len(lines) < 4:
        raise ValueError("Ficheiro DXF invalido ou vazio.")
    pairs: list[tuple[int, str]] = []
    limit = len(lines) - (len(lines) % 2)
    for index in range(0, limit, 2):
        code_raw = str(lines[index] or "").strip()
        value = str(lines[index + 1] or "").rstrip("\r\n")
        try:
            code = int(code_raw)
        except Exception:
            continue
        pairs.append((code, value))
    return pairs


def _entity_common(entity_pairs: list[tuple[int, str]]) -> dict[str, Any]:
    layer = ""
    for code, value in entity_pairs:
        if code == 8:
            layer = str(value or "").strip()
            break
    return {"layer": layer}


def _parse_lwpolyline(entity_pairs: list[tuple[int, str]]) -> dict[str, Any]:
    common = _entity_common(entity_pairs)
    closed = False
    vertices: list[dict[str, float]] = []
    current: dict[str, float] | None = None
    for code, value in entity_pairs:
        if code == 70:
            closed = bool(_as_int(value, 0) & 1)
        elif code == 10:
            if current is not None and "x" in current and "y" in current:
                vertices.append(current)
            current = {"x": _as_float(value, 0.0), "bulge": 0.0}
        elif code == 20:
            if current is None:
                current = {"x": 0.0, "bulge": 0.0}
            current["y"] = _as_float(value, 0.0)
        elif code == 42:
            if current is None:
                current = {"x": 0.0, "y": 0.0}
            current["bulge"] = _as_float(value, 0.0)
    if current is not None and "x" in current and "y" in current:
        vertices.append(current)
    if len(vertices) < 2:
        raise ValueError("LWPOLYLINE sem vertices suficientes.")
    points: list[tuple[float, float]] = []
    length_mm = 0.0
    vertex_count = len(vertices)
    segment_limit = vertex_count if closed else vertex_count - 1
    for index in range(segment_limit):
        current_vertex = vertices[index]
        next_vertex = vertices[(index + 1) % vertex_count]
        start = (float(current_vertex["x"]), float(current_vertex["y"]))
        end = (float(next_vertex["x"]), float(next_vertex["y"]))
        segment_points = _sample_bulge_segment(start, end, float(current_vertex.get("bulge", 0.0) or 0.0))
        if not points:
            points.extend(segment_points)
        else:
            points.extend(segment_points[1:])
        length_mm += sum(_distance(segment_points[i], segment_points[i + 1]) for i in range(len(segment_points) - 1))
    if closed and points and points[0] != points[-1]:
        points.append(points[0])
    return {
        **common,
        "entity_type": "LWPOLYLINE",
        "points": points,
        "closed": closed,
        "length_mm": round(length_mm, 4),
    }


def _parse_line(entity_pairs: list[tuple[int, str]]) -> dict[str, Any]:
    common = _entity_common(entity_pairs)
    x1 = y1 = x2 = y2 = 0.0
    for code, value in entity_pairs:
        if code == 10:
            x1 = _as_float(value, 0.0)
        elif code == 20:
            y1 = _as_float(value, 0.0)
        elif code == 11:
            x2 = _as_float(value, 0.0)
        elif code == 21:
            y2 = _as_float(value, 0.0)
    points = [(x1, y1), (x2, y2)]
    return {
        **common,
        "entity_type": "LINE",
        "points": points,
        "closed": False,
        "length_mm": round(_distance(points[0], points[1]), 4),
    }


def _parse_circle(entity_pairs: list[tuple[int, str]]) -> dict[str, Any]:
    common = _entity_common(entity_pairs)
    cx = cy = radius = 0.0
    for code, value in entity_pairs:
        if code == 10:
            cx = _as_float(value, 0.0)
        elif code == 20:
            cy = _as_float(value, 0.0)
        elif code == 40:
            radius = abs(_as_float(value, 0.0))
    points = _sample_circle_points((cx, cy), radius)
    return {
        **common,
        "entity_type": "CIRCLE",
        "points": points,
        "closed": True,
        "length_mm": round(2.0 * math.pi * radius, 4),
    }


def _parse_arc(entity_pairs: list[tuple[int, str]]) -> dict[str, Any]:
    common = _entity_common(entity_pairs)
    cx = cy = radius = start_deg = end_deg = 0.0
    for code, value in entity_pairs:
        if code == 10:
            cx = _as_float(value, 0.0)
        elif code == 20:
            cy = _as_float(value, 0.0)
        elif code == 40:
            radius = abs(_as_float(value, 0.0))
        elif code == 50:
            start_deg = _as_float(value, 0.0)
        elif code == 51:
            end_deg = _as_float(value, 0.0)
    points = _sample_arc_points((cx, cy), radius, start_deg, end_deg, ccw=True, max_step_deg=6.0)
    span_deg = end_deg - start_deg
    if span_deg <= 0:
        span_deg += 360.0
    return {
        **common,
        "entity_type": "ARC",
        "points": points,
        "closed": False,
        "length_mm": round((abs(span_deg) / 360.0) * (2.0 * math.pi * radius), 4),
    }


def _parse_polyline(polyline_pairs: list[tuple[int, str]], vertex_groups: list[list[tuple[int, str]]]) -> dict[str, Any]:
    common = _entity_common(polyline_pairs)
    closed = False
    for code, value in polyline_pairs:
        if code == 70:
            closed = bool(_as_int(value, 0) & 1)
            break
    vertices: list[dict[str, float]] = []
    for vertex_pairs in vertex_groups:
        current = {"x": 0.0, "y": 0.0, "bulge": 0.0}
        for code, value in vertex_pairs:
            if code == 10:
                current["x"] = _as_float(value, 0.0)
            elif code == 20:
                current["y"] = _as_float(value, 0.0)
            elif code == 42:
                current["bulge"] = _as_float(value, 0.0)
        vertices.append(current)
    if len(vertices) < 2:
        raise ValueError("POLYLINE sem vertices suficientes.")
    points: list[tuple[float, float]] = []
    length_mm = 0.0
    vertex_count = len(vertices)
    segment_limit = vertex_count if closed else vertex_count - 1
    for index in range(segment_limit):
        current_vertex = vertices[index]
        next_vertex = vertices[(index + 1) % vertex_count]
        start = (float(current_vertex["x"]), float(current_vertex["y"]))
        end = (float(next_vertex["x"]), float(next_vertex["y"]))
        segment_points = _sample_bulge_segment(start, end, float(current_vertex.get("bulge", 0.0) or 0.0))
        if not points:
            points.extend(segment_points)
        else:
            points.extend(segment_points[1:])
        length_mm += sum(_distance(segment_points[i], segment_points[i + 1]) for i in range(len(segment_points) - 1))
    if closed and points and points[0] != points[-1]:
        points.append(points[0])
    return {
        **common,
        "entity_type": "POLYLINE",
        "points": points,
        "closed": closed,
        "length_mm": round(length_mm, 4),
    }


def _path_length_mm(points: list[tuple[float, float]]) -> float:
    if len(points) < 2:
        return 0.0
    return sum(_distance(points[index], points[index + 1]) for index in range(len(points) - 1))


def _parse_splines_with_ezdxf(path: str | Path) -> tuple[list[dict[str, Any]], list[str], int]:
    if ezdxf is None:
        return [], [], 0
    contours: list[dict[str, Any]] = []
    warnings: list[str] = []
    try:
        document = ezdxf.readfile(str(path))
        modelspace = document.modelspace()
    except Exception as exc:
        return [], [f"SPLINE ignorada: nao foi possivel ler com ezdxf ({exc})."], 0

    for entity in modelspace.query("SPLINE"):
        try:
            raw_points = list(entity.flattening(0.20, segments=8))
            points: list[tuple[float, float]] = []
            for point in raw_points:
                x = round(_as_float(point[0], 0.0), 4)
                y = round(_as_float(point[1], 0.0), 4)
                candidate = (x, y)
                if points and points[-1] == candidate:
                    continue
                points.append(candidate)
            closed = bool(getattr(entity, "closed", False))
            if len(points) < 2:
                continue
            if closed and points[0] != points[-1]:
                points.append(points[0])
            contours.append(
                {
                    "layer": str(getattr(entity.dxf, "layer", "") or "").strip(),
                    "entity_type": "SPLINE",
                    "points": points,
                    "closed": closed,
                    "length_mm": round(_path_length_mm(points), 4),
                }
            )
        except Exception as exc:
            warnings.append(f"SPLINE ignorada: {exc}")
    return contours, warnings, len(contours)


def _parse_entities(path: str | Path) -> tuple[list[dict[str, Any]], list[str], dict[str, int]]:
    pairs = _iter_pairs(path)
    contours: list[dict[str, Any]] = []
    warnings: list[str] = []
    counts: Counter[str] = Counter()
    in_entities = False
    index = 0
    while index < len(pairs):
        code, value = pairs[index]
        value_upper = str(value or "").strip().upper()
        if code == 0 and value_upper == "SECTION":
            if index + 1 < len(pairs) and pairs[index + 1][0] == 2 and str(pairs[index + 1][1] or "").strip().upper() == "ENTITIES":
                in_entities = True
                index += 2
                continue
        if code == 0 and value_upper == "ENDSEC":
            in_entities = False
            index += 1
            continue
        if not in_entities or code != 0:
            index += 1
            continue
        if value_upper == "POLYLINE":
            polyline_pairs: list[tuple[int, str]] = []
            vertex_groups: list[list[tuple[int, str]]] = []
            index += 1
            while index < len(pairs) and pairs[index][0] != 0:
                polyline_pairs.append(pairs[index])
                index += 1
            while index < len(pairs):
                sub_code, sub_value = pairs[index]
                sub_type = str(sub_value or "").strip().upper()
                if sub_code != 0:
                    index += 1
                    continue
                if sub_type == "VERTEX":
                    index += 1
                    vertex_pairs: list[tuple[int, str]] = []
                    while index < len(pairs) and pairs[index][0] != 0:
                        vertex_pairs.append(pairs[index])
                        index += 1
                    vertex_groups.append(vertex_pairs)
                    continue
                if sub_type == "SEQEND":
                    index += 1
                    break
                break
            try:
                contour = _parse_polyline(polyline_pairs, vertex_groups)
                contours.append(contour)
                counts[contour["entity_type"]] += 1
            except Exception as exc:
                warnings.append(f"POLYLINE ignorada: {exc}")
            continue
        entity_pairs: list[tuple[int, str]] = []
        entity_type = value_upper
        index += 1
        while index < len(pairs) and pairs[index][0] != 0:
            entity_pairs.append(pairs[index])
            index += 1
        try:
            if entity_type == "LWPOLYLINE":
                contour = _parse_lwpolyline(entity_pairs)
            elif entity_type == "LINE":
                contour = _parse_line(entity_pairs)
            elif entity_type == "CIRCLE":
                contour = _parse_circle(entity_pairs)
            elif entity_type == "ARC":
                contour = _parse_arc(entity_pairs)
            else:
                counts[f"UNSUPPORTED:{entity_type}"] += 1
                continue
            contours.append(contour)
            counts[entity_type] += 1
        except Exception as exc:
            warnings.append(f"{entity_type} ignorada: {exc}")
    unsupported_splines = int(counts.get("UNSUPPORTED:SPLINE", 0) or 0)
    if unsupported_splines > 0:
        spline_contours, spline_warnings, spline_count = _parse_splines_with_ezdxf(path)
        warnings.extend(spline_warnings)
        if spline_contours:
            contours.extend(spline_contours)
            counts["SPLINE"] += spline_count
            remaining_unsupported = max(0, unsupported_splines - spline_count)
            if remaining_unsupported > 0:
                counts["UNSUPPORTED:SPLINE"] = remaining_unsupported
            else:
                counts.pop("UNSUPPORTED:SPLINE", None)
    if not contours:
        raise ValueError("Nao foram encontradas entidades DXF suportadas (LINE, ARC, CIRCLE, LWPOLYLINE ou POLYLINE).")
    return contours, warnings, dict(counts)


def _classify_layer(layer_name: str, layer_rules: dict[str, Any]) -> str:
    layer_txt = str(layer_name or "").strip().lower()
    ignore_patterns = [str(value or "").strip().lower() for value in list(layer_rules.get("ignore_patterns", []) or []) if str(value or "").strip()]
    mark_patterns = [str(value or "").strip().lower() for value in list(layer_rules.get("mark_patterns", []) or []) if str(value or "").strip()]
    for pattern in ignore_patterns:
        if pattern and pattern in layer_txt:
            return "ignore"
    for pattern in mark_patterns:
        if pattern and pattern in layer_txt:
            return "mark"
    return "cut"


def _component_from_single_contour(contour: dict[str, Any]) -> dict[str, Any]:
    points = list(contour.get("points", []) or [])
    bbox = _bbox_from_points(points)
    closed = bool(contour.get("closed"))
    area = abs(_shoelace_area(points)) if closed else 0.0
    center = _polygon_centroid(points) if closed else ((bbox["min_x"] + bbox["max_x"]) / 2.0, (bbox["min_y"] + bbox["max_y"]) / 2.0)
    return {
        "closed": closed,
        "points": points,
        "length_mm": _as_float(contour.get("length_mm", 0), 0.0),
        "bbox": bbox,
        "center": center,
        "area_mm2": abs(area),
        "entity_type": str(contour.get("entity_type", "") or "").strip(),
        "branching": False,
    }


def _assemble_open_components(open_contours: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[str]]:
    warnings: list[str] = []
    if not open_contours:
        return [], warnings
    edges: list[dict[str, Any]] = []
    node_to_edges: defaultdict[tuple[int, int], list[int]] = defaultdict(list)
    for contour in open_contours:
        points = list(contour.get("points", []) or [])
        if len(points) < 2:
            continue
        start = tuple(points[0])
        end = tuple(points[-1])
        edge = {
            "points": points,
            "length_mm": _as_float(contour.get("length_mm", 0), 0.0),
            "start": start,
            "end": end,
            "start_key": _snap_key(start),
            "end_key": _snap_key(end),
            "entity_type": str(contour.get("entity_type", "") or "").strip(),
        }
        edge_index = len(edges)
        edges.append(edge)
        node_to_edges[edge["start_key"]].append(edge_index)
        node_to_edges[edge["end_key"]].append(edge_index)
    visited_edges: set[int] = set()
    components: list[dict[str, Any]] = []
    for edge_index, edge in enumerate(edges):
        if edge_index in visited_edges:
            continue
        stack = [edge_index]
        component_edges: set[int] = set()
        component_nodes: set[tuple[int, int]] = set()
        while stack:
            current_edge = stack.pop()
            if current_edge in component_edges:
                continue
            component_edges.add(current_edge)
            current = edges[current_edge]
            for node_key in (current["start_key"], current["end_key"]):
                component_nodes.add(node_key)
                for linked_edge in list(node_to_edges.get(node_key, []) or []):
                    if linked_edge not in component_edges:
                        stack.append(linked_edge)
        visited_edges.update(component_edges)
        degree_map = {node_key: len(list(node_to_edges.get(node_key, []) or [])) for node_key in component_nodes}
        branching = any(int(degree or 0) > 2 for degree in degree_map.values())
        ordered_points: list[tuple[float, float]] = []
        total_length_mm = 0.0
        bbox_points: list[tuple[float, float]] = []
        if branching:
            warnings.append("Foram detetados segmentos DXF com ramificacoes. O numero de perfuracoes e a area podem ficar aproximados.")
            for current_edge in list(component_edges):
                edge_obj = edges[current_edge]
                total_length_mm += _as_float(edge_obj.get("length_mm", 0), 0.0)
                bbox_points.extend(list(edge_obj.get("points", []) or []))
            bbox = _bbox_from_points(bbox_points)
            center = ((bbox["min_x"] + bbox["max_x"]) / 2.0, (bbox["min_y"] + bbox["max_y"]) / 2.0)
            components.append(
                {
                    "closed": False,
                    "points": bbox_points,
                    "length_mm": total_length_mm,
                    "bbox": bbox,
                    "center": center,
                    "area_mm2": 0.0,
                    "entity_type": "ASSEMBLED",
                    "branching": True,
                }
            )
            continue
        open_nodes = [node_key for node_key, degree in list(degree_map.items()) if int(degree or 0) == 1]
        start_node = open_nodes[0] if open_nodes else next(iter(component_nodes))
        used_edges: set[int] = set()
        current_node = start_node
        while len(used_edges) < len(component_edges):
            next_edge_index = -1
            next_node = current_node
            for candidate in list(node_to_edges.get(current_node, []) or []):
                if candidate in component_edges and candidate not in used_edges:
                    next_edge_index = candidate
                    candidate_edge = edges[candidate]
                    next_node = candidate_edge["end_key"] if candidate_edge["start_key"] == current_node else candidate_edge["start_key"]
                    break
            if next_edge_index < 0:
                remaining = [candidate for candidate in list(component_edges) if candidate not in used_edges]
                if not remaining:
                    break
                next_edge_index = remaining[0]
                candidate_edge = edges[next_edge_index]
                current_node = candidate_edge["start_key"]
                next_node = candidate_edge["end_key"]
            edge_obj = edges[next_edge_index]
            used_edges.add(next_edge_index)
            edge_points = list(edge_obj.get("points", []) or [])
            if edge_obj["start_key"] != current_node:
                edge_points = list(reversed(edge_points))
            if not ordered_points:
                ordered_points.extend(edge_points)
            else:
                ordered_points.extend(edge_points[1:])
            total_length_mm += _as_float(edge_obj.get("length_mm", 0), 0.0)
            current_node = next_node
        closed = bool(not open_nodes and ordered_points)
        if closed and ordered_points and ordered_points[0] != ordered_points[-1]:
            ordered_points.append(ordered_points[0])
        bbox = _bbox_from_points(ordered_points)
        center = _polygon_centroid(ordered_points) if closed else ((bbox["min_x"] + bbox["max_x"]) / 2.0, (bbox["min_y"] + bbox["max_y"]) / 2.0)
        components.append(
            {
                "closed": closed,
                "points": ordered_points,
                "length_mm": total_length_mm,
                "bbox": bbox,
                "center": center,
                "area_mm2": abs(_shoelace_area(ordered_points)) if closed else 0.0,
                "entity_type": "ASSEMBLED",
                "branching": False,
            }
        )
    return components, warnings


def _estimate_rapid_length_mm(components: list[dict[str, Any]]) -> float:
    centers = [tuple(component.get("center", (0.0, 0.0))) for component in list(components or []) if component.get("center")]
    if len(centers) <= 1:
        return 0.0
    remaining = list(centers[1:])
    current = centers[0]
    total = 0.0
    while remaining:
        next_index = min(range(len(remaining)), key=lambda idx: _distance(current, remaining[idx]))
        target = remaining.pop(next_index)
        total += _distance(current, target)
        current = target
    return round(total, 3)


def analyze_dxf_geometry(path: str | Path, layer_rules: dict[str, Any] | None = None) -> dict[str, Any]:
    source_path, analysis_path, resolution_warnings = _resolve_cad_analysis_input(path)
    rules = dict(layer_rules or {})
    contours, warnings, counts = _parse_entities(analysis_path)
    warnings = list(resolution_warnings) + list(warnings)
    cut_contours: list[dict[str, Any]] = []
    mark_contours: list[dict[str, Any]] = []
    ignored_count = 0
    for contour in contours:
        mode = _classify_layer(str(contour.get("layer", "") or ""), rules)
        contour["mode"] = mode
        if mode == "ignore":
            ignored_count += 1
            continue
        if mode == "mark":
            mark_contours.append(contour)
        else:
            cut_contours.append(contour)
    closed_cut = [_component_from_single_contour(contour) for contour in cut_contours if bool(contour.get("closed"))]
    open_cut = [contour for contour in cut_contours if not bool(contour.get("closed"))]
    assembled_open, open_warnings = _assemble_open_components(open_cut)
    warnings.extend(open_warnings)
    cut_components = closed_cut + assembled_open
    if not cut_components and mark_contours:
        warnings.append("Foram detetadas apenas camadas de marcacao. Confirma se o DXF contem geometrias de corte.")
    cut_length_mm = round(sum(_as_float(component.get("length_mm", 0), 0.0) for component in cut_components), 3)
    mark_length_mm = round(sum(_as_float(contour.get("length_mm", 0), 0.0) for contour in mark_contours), 3)
    cut_points: list[tuple[float, float]] = []
    for component in cut_components:
        cut_points.extend(list(component.get("points", []) or []))
    bbox = _bbox_from_points(cut_points)
    closed_polygons = [component for component in cut_components if bool(component.get("closed")) and len(list(component.get("points", []) or [])) >= 4]
    open_components = [component for component in cut_components if not bool(component.get("closed"))]
    net_area_mm2 = 0.0
    outer_closed = 0
    hole_closed = 0
    outer_polygons_data: list[list[tuple[float, float]]] = []
    hole_polygons_data: list[list[tuple[float, float]]] = []
    estimated_outer_from_open = False
    polygon_areas = [abs(_as_float(component.get("area_mm2", 0), 0.0)) for component in list(closed_polygons or [])]
    polygon_bboxes = [dict(component.get("bbox", {}) or {}) for component in list(closed_polygons or [])]
    parent_index_map: dict[int, int] = {}
    for index, component in enumerate(closed_polygons):
        polygon = list(component.get("points", []) or [])
        rep_point = tuple(component.get("center", (0.0, 0.0)))
        current_area = polygon_areas[index] if index < len(polygon_areas) else abs(_as_float(component.get("area_mm2", 0), 0.0))
        candidates: list[tuple[float, int]] = []
        for other_index, other in enumerate(closed_polygons):
            if other_index == index:
                continue
            other_area = polygon_areas[other_index] if other_index < len(polygon_areas) else abs(_as_float(other.get("area_mm2", 0), 0.0))
            if other_area <= (current_area + 1e-6):
                continue
            other_polygon = list(other.get("points", []) or [])
            if other_polygon and _point_in_polygon(rep_point, other_polygon):
                candidates.append((other_area, other_index))
        parent_index_map[index] = min(candidates, key=lambda row: row[0])[1] if candidates else -1

    if closed_polygons and bbox["width"] > 0.0 and bbox["height"] > 0.0:
        global_bbox_area = bbox["width"] * bbox["height"]
        dominant_index = max(range(len(closed_polygons)), key=lambda idx: polygon_areas[idx])
        dominant_area = polygon_areas[dominant_index] if dominant_index < len(polygon_areas) else 0.0
        dominant_bbox = polygon_bboxes[dominant_index] if dominant_index < len(polygon_bboxes) else {}
        dominant_like_part = (
            dominant_area >= (global_bbox_area * 0.60)
            and _as_float(dominant_bbox.get("width", 0.0), 0.0) >= (bbox["width"] * 0.90)
            and _as_float(dominant_bbox.get("height", 0.0), 0.0) >= (bbox["height"] * 0.90)
        )
        if dominant_like_part:
            tol = 0.5
            dom_min_x = _as_float(dominant_bbox.get("min_x", 0.0), 0.0) - tol
            dom_min_y = _as_float(dominant_bbox.get("min_y", 0.0), 0.0) - tol
            dom_max_x = _as_float(dominant_bbox.get("max_x", 0.0), 0.0) + tol
            dom_max_y = _as_float(dominant_bbox.get("max_y", 0.0), 0.0) + tol
            for index, component_bbox in enumerate(list(polygon_bboxes or [])):
                if index == dominant_index:
                    continue
                min_x = _as_float(component_bbox.get("min_x", 0.0), 0.0)
                min_y = _as_float(component_bbox.get("min_y", 0.0), 0.0)
                max_x = _as_float(component_bbox.get("max_x", 0.0), 0.0)
                max_y = _as_float(component_bbox.get("max_y", 0.0), 0.0)
                if min_x >= dom_min_x and min_y >= dom_min_y and max_x <= dom_max_x and max_y <= dom_max_y:
                    parent_index_map[index] = dominant_index

    global_bbox_area = bbox["width"] * bbox["height"] if bbox["width"] > 0.0 and bbox["height"] > 0.0 else 0.0
    largest_closed_area = max(polygon_areas) if polygon_areas else 0.0
    dominant_open = None
    if open_components:
        dominant_open = max(
            open_components,
            key=lambda component: (
                (_as_float(dict(component.get("bbox", {}) or {}).get("width", 0.0), 0.0) * _as_float(dict(component.get("bbox", {}) or {}).get("height", 0.0), 0.0)),
                _as_float(component.get("length_mm", 0.0), 0.0),
            ),
        )
    if dominant_open and global_bbox_area > 0.0:
        dominant_open_bbox = dict(dominant_open.get("bbox", {}) or {})
        dominant_open_area = _as_float(dominant_open_bbox.get("width", 0.0), 0.0) * _as_float(dominant_open_bbox.get("height", 0.0), 0.0)
        dominant_open_like_part = (
            dominant_open_area >= (global_bbox_area * 0.85)
            and _as_float(dominant_open_bbox.get("width", 0.0), 0.0) >= (bbox["width"] * 0.90)
            and _as_float(dominant_open_bbox.get("height", 0.0), 0.0) >= (bbox["height"] * 0.90)
        )
        dominant_closed_is_weak = (not closed_polygons) or (largest_closed_area < (global_bbox_area * 0.25))
        if dominant_open_like_part and dominant_closed_is_weak:
            tol = 0.5
            min_x = _as_float(dominant_open_bbox.get("min_x", 0.0), 0.0)
            min_y = _as_float(dominant_open_bbox.get("min_y", 0.0), 0.0)
            max_x = _as_float(dominant_open_bbox.get("max_x", 0.0), 0.0)
            max_y = _as_float(dominant_open_bbox.get("max_y", 0.0), 0.0)
            holes_area_mm2 = 0.0
            detected_holes: list[list[tuple[float, float]]] = []
            for index, component in enumerate(closed_polygons):
                component_bbox = polygon_bboxes[index] if index < len(polygon_bboxes) else {}
                hole_min_x = _as_float(component_bbox.get("min_x", 0.0), 0.0)
                hole_min_y = _as_float(component_bbox.get("min_y", 0.0), 0.0)
                hole_max_x = _as_float(component_bbox.get("max_x", 0.0), 0.0)
                hole_max_y = _as_float(component_bbox.get("max_y", 0.0), 0.0)
                if (
                    hole_min_x >= (min_x - tol)
                    and hole_min_y >= (min_y - tol)
                    and hole_max_x <= (max_x + tol)
                    and hole_max_y <= (max_y + tol)
                ):
                    holes_area_mm2 += polygon_areas[index] if index < len(polygon_areas) else abs(_as_float(component.get("area_mm2", 0.0), 0.0))
                    polygon = list(component.get("points", []) or [])
                    if polygon:
                        detected_holes.append([(round(float(x), 3), round(float(y), 3)) for x, y in polygon])
            net_area_mm2 = max(0.0, dominant_open_area - holes_area_mm2)
            outer_closed = 1
            hole_closed = len(detected_holes)
            outer_polygons_data = [[
                (round(float(min_x), 3), round(float(min_y), 3)),
                (round(float(max_x), 3), round(float(min_y), 3)),
                (round(float(max_x), 3), round(float(max_y), 3)),
                (round(float(min_x), 3), round(float(max_y), 3)),
                (round(float(min_x), 3), round(float(min_y), 3)),
            ]]
            hole_polygons_data = detected_holes
            estimated_outer_from_open = True
            warnings.append("Contorno exterior aberto/ramificado: a area liquida foi estimada pela caixa do contorno principal.")

    if not estimated_outer_from_open:
        for index, component in enumerate(closed_polygons):
            polygon = list(component.get("points", []) or [])
            area_abs = polygon_areas[index] if index < len(polygon_areas) else abs(_as_float(component.get("area_mm2", 0), 0.0))
            if parent_index_map.get(index, -1) < 0:
                net_area_mm2 += area_abs
                outer_closed += 1
                outer_polygons_data.append([(round(float(x), 3), round(float(y), 3)) for x, y in polygon])
            else:
                net_area_mm2 -= area_abs
                hole_closed += 1
                hole_polygons_data.append([(round(float(x), 3), round(float(y), 3)) for x, y in polygon])
    net_area_mm2 = max(0.0, net_area_mm2)
    pierce_count = len(cut_components)
    rapid_length_mm = _estimate_rapid_length_mm(cut_components)
    unsupported = {key.replace("UNSUPPORTED:", ""): int(value) for key, value in list(counts.items()) if key.startswith("UNSUPPORTED:")}
    if unsupported:
        warnings.append("Existem entidades DXF nao suportadas nesta versao: " + ", ".join(sorted(unsupported.keys())) + ".")
    if outer_closed > 1:
        warnings.append(f"Foram detetados {outer_closed} contornos exteriores. Confirma se o ficheiro representa uma unica peca ou um conjunto.")
    if cut_length_mm <= 0.0 and mark_length_mm <= 0.0:
        raise ValueError("O DXF nao contem comprimentos validos para orcamentar.")
    area_quality = "estimada_contorno_aberto" if estimated_outer_from_open else ("exacta" if net_area_mm2 > 0 and closed_polygons else "estimada")
    cut_preview_paths = [
        [(round(float(x), 3), round(float(y), 3)) for x, y in list(contour.get("points", []) or [])]
        for contour in list(cut_contours or [])
        if len(list(contour.get("points", []) or [])) >= 2
    ]
    mark_preview_paths = [
        [(round(float(x), 3), round(float(y), 3)) for x, y in list(contour.get("points", []) or [])]
        for contour in list(mark_contours or [])
        if len(list(contour.get("points", []) or [])) >= 2
    ]
    return {
        "file_path": str(source_path),
        "file_name": source_path.name,
        "file_stem": source_path.stem,
        "analysis_file_path": str(analysis_path),
        "analysis_file_name": analysis_path.name,
        "entity_counts": counts,
        "unsupported_entities": unsupported,
        "warnings": list(dict.fromkeys([str(item or "").strip() for item in warnings if str(item or "").strip()])),
        "layer_summary": {
            "cut_entities": len(cut_contours),
            "mark_entities": len(mark_contours),
            "ignored_entities": ignored_count,
        },
        "bbox_mm": bbox,
        "nesting_shape": {
            "outer_polygons": outer_polygons_data,
            "hole_polygons": hole_polygons_data,
            "shape_quality": "exacta" if outer_polygons_data else "bbox",
        },
        "preview_paths": {
            "cut_paths": cut_preview_paths,
            "mark_paths": mark_preview_paths,
        },
        "metrics": {
            "cut_length_mm": round(cut_length_mm, 3),
            "cut_length_m": round(cut_length_mm / 1000.0, 4),
            "mark_length_mm": round(mark_length_mm, 3),
            "mark_length_m": round(mark_length_mm / 1000.0, 4),
            "rapid_length_mm": round(rapid_length_mm, 3),
            "rapid_length_m": round(rapid_length_mm / 1000.0, 4),
            "pierce_count": int(max(0, pierce_count)),
            "closed_contours": int(max(0, outer_closed + hole_closed)),
            "outer_contours": int(max(0, outer_closed)),
            "inner_contours": int(max(0, hole_closed)),
            "part_count_hint": 1 if estimated_outer_from_open else int(max(1 if cut_components else 0, outer_closed + max(0, len([component for component in cut_components if not component.get("closed")])))),
            "net_area_mm2": round(net_area_mm2, 2),
            "net_area_m2": round(net_area_mm2 / 1_000_000.0, 6),
            "bbox_area_mm2": round((bbox["width"] * bbox["height"]), 2),
            "bbox_area_m2": round((bbox["width"] * bbox["height"]) / 1_000_000.0, 6),
            "area_quality": area_quality,
        },
    }


def _active_machine(settings: dict[str, Any], machine_name: str = "") -> dict[str, Any]:
    machine_profiles = dict(settings.get("machine_profiles", {}) or {})
    selected = str(machine_name or settings.get("active_machine", "") or "").strip()
    if selected and selected in machine_profiles:
        return copy.deepcopy(dict(machine_profiles.get(selected, {}) or {}))
    if machine_profiles:
        key = next(iter(machine_profiles.keys()))
        return copy.deepcopy(dict(machine_profiles.get(key, {}) or {}))
    return {}


def _active_commercial(settings: dict[str, Any], commercial_name: str = "") -> dict[str, Any]:
    profiles = dict(settings.get("commercial_profiles", {}) or {})
    selected = str(commercial_name or settings.get("active_commercial", "") or "").strip()
    if selected and selected in profiles:
        return copy.deepcopy(dict(profiles.get(selected, {}) or {}))
    if profiles:
        key = next(iter(profiles.keys()))
        return copy.deepcopy(dict(profiles.get(key, {}) or {}))
    return {}


def _resolve_material_profile(materials: dict[str, Any], material_name: str, material_subtype: str = "") -> tuple[str, dict[str, Any]]:
    subtype = str(material_subtype or "").strip()
    family = _canonical_material_family(material_name) or str(material_name or "").strip()
    for candidate in (subtype, family):
        if candidate and candidate in materials:
            return candidate, dict(materials.get(candidate, {}) or {})
    family_token = _norm_material_token(family)
    for key, value in dict(materials or {}).items():
        if _norm_material_token(key) == family_token:
            return str(key), dict(value or {})
    if materials:
        key = next(iter(materials.keys()))
        return key, dict(materials.get(key, {}) or {})
    return "", {}


def _machine_material(machine_profile: dict[str, Any], material_name: str, material_subtype: str = "") -> dict[str, Any]:
    materials = dict(machine_profile.get("materials", {}) or {})
    return _resolve_material_profile(materials, material_name, material_subtype)[1]


def _commercial_material(profile: dict[str, Any], material_name: str, material_subtype: str = "") -> dict[str, Any]:
    materials = dict(profile.get("materials", {}) or {})
    return _resolve_material_profile(materials, material_name, material_subtype)[1]


def _commercial_material_subtype_override(profile: dict[str, Any], material_name: str, material_subtype: str = "") -> dict[str, Any]:
    subtype = str(material_subtype or "").strip()
    if not subtype:
        return {}
    catalog = dict(profile.get("material_catalog", {}) or {})
    family = _canonical_material_family(material_name) or str(material_name or "").strip()
    family_catalog = dict(catalog.get(family, {}) or {})
    if subtype in family_catalog:
        return dict(family_catalog.get(subtype, {}) or {})
    subtype_token = _norm_material_token(subtype)
    for key, value in family_catalog.items():
        if _norm_material_token(key) == subtype_token:
            return dict(value or {})
    return {}


def _series_pricing_tiers(profile: dict[str, Any]) -> list[dict[str, Any]]:
    series_cfg = dict(profile.get("series_pricing", {}) or {})
    tiers = list(series_cfg.get("tiers", []) or [])
    normalized: list[dict[str, Any]] = []
    for index, tier in enumerate(tiers):
        row = dict(tier or {})
        qty_min = max(1, _as_int(row.get("qty_min", 1), 1))
        qty_max = max(qty_min, _as_int(row.get("qty_max", qty_min), qty_min))
        normalized.append(
            {
                "key": str(row.get("key", f"tier_{index + 1}") or f"tier_{index + 1}").strip(),
                "label": str(row.get("label", f"Serie {index + 1}") or f"Serie {index + 1}").strip(),
                "qty_min": qty_min,
                "qty_max": qty_max,
                "margin_delta_pct": _as_float(row.get("margin_delta_pct", 0.0), 0.0),
                "setup_multiplier": max(0.05, _as_float(row.get("setup_multiplier", 1.0), 1.0)),
            }
        )
    normalized.sort(key=lambda item: (int(item.get("qty_min", 1)), int(item.get("qty_max", 1))))
    return normalized


def _pick_series_tier(profile: dict[str, Any], quantity: int) -> dict[str, Any]:
    tiers = _series_pricing_tiers(profile)
    qty = max(1, int(quantity or 1))
    for tier in tiers:
        if int(tier.get("qty_min", 1)) <= qty <= int(tier.get("qty_max", qty)):
            return dict(tier)
    if tiers:
        return dict(tiers[-1])
    return {
        "key": "default",
        "label": "Sem serie",
        "qty_min": 1,
        "qty_max": 999999,
        "margin_delta_pct": 0.0,
        "setup_multiplier": 1.0,
    }


def _gas_rows(machine_profile: dict[str, Any], material_name: str, gas_name: str, material_subtype: str = "") -> tuple[str, list[dict[str, Any]], dict[str, Any]]:
    material_profile = _machine_material(machine_profile, material_name, material_subtype)
    gases = dict((material_profile.get("gases") or {}))
    selected_gas = str(gas_name or material_profile.get("default_gas", "") or "").strip()
    if selected_gas and selected_gas in gases:
        return selected_gas, list(dict(gases.get(selected_gas, {}) or {}).get("rows", []) or []), material_profile
    if gases:
        selected_gas = next(iter(gases.keys()))
        return selected_gas, list(dict(gases.get(selected_gas, {}) or {}).get("rows", []) or []), material_profile
    return selected_gas, [], material_profile


def _lookup_cut_row(rows: list[dict[str, Any]], thickness_mm: float) -> dict[str, Any]:
    normalized = sorted(
        [dict(row or {}) for row in list(rows or []) if _as_float((row or {}).get("thickness_mm", 0), 0.0) > 0.0],
        key=lambda row: _as_float(row.get("thickness_mm", 0), 0.0),
    )
    if not normalized:
        return {}
    target = float(thickness_mm or 0.0)
    for row in normalized:
        if abs(_as_float(row.get("thickness_mm", 0), 0.0) - target) <= 1e-9:
            return dict(row)
    lower = None
    upper = None
    for row in normalized:
        value = _as_float(row.get("thickness_mm", 0), 0.0)
        if value < target:
            lower = row
        if value > target and upper is None:
            upper = row
            break
    if lower is None:
        return dict(normalized[0])
    if upper is None:
        return dict(normalized[-1])
    lower_t = _as_float(lower.get("thickness_mm", 0), 0.0)
    upper_t = _as_float(upper.get("thickness_mm", 0), 0.0)
    ratio = _safe_div(target - lower_t, upper_t - lower_t, 0.0)
    out = {"thickness_mm": target}
    numeric_keys = {
        "speed_min_m_min",
        "speed_max_m_min",
        "nozzle_distance_mm",
        "gas_pressure_bar_min",
        "gas_pressure_bar_max",
        "focus_mm",
        "duty_pct",
        "frequency_hz",
        "power_w",
    }
    for key in numeric_keys:
        out[key] = round(
            _as_float(lower.get(key, 0), 0.0) + ((_as_float(upper.get(key, 0), 0.0) - _as_float(lower.get(key, 0), 0.0)) * ratio),
            4,
        )
    out["nozzle"] = str(lower.get("nozzle", "") or upper.get("nozzle", "") or "").strip()
    out["interpolated"] = True
    return out


def _effective_speed_m_min(cut_row: dict[str, Any], motion_cfg: dict[str, Any]) -> float:
    speed_min = _as_float(cut_row.get("speed_min_m_min", 0), 0.0)
    speed_max = _as_float(cut_row.get("speed_max_m_min", 0), 0.0)
    if speed_max <= 0.0 and speed_min > 0.0:
        speed_max = speed_min
    if speed_min <= 0.0 and speed_max > 0.0:
        speed_min = speed_max
    nominal = (speed_min + speed_max) / 2.0 if speed_max > 0.0 else speed_min
    factor = max(0.0, _as_float(motion_cfg.get("effective_speed_factor_pct", 100.0), 100.0)) / 100.0
    return round(max(0.01, nominal * factor), 4)


def _profile_thickness_rate_factor(
    gas_rows: list[dict[str, Any]],
    cut_row: dict[str, Any],
    motion_cfg: dict[str, Any],
    thickness_mm: float,
    commercial_profile: dict[str, Any],
    payload: dict[str, Any],
) -> tuple[float, float, float]:
    reference_thickness = max(
        0.1,
        _as_float(
            payload.get("profile_reference_thickness_mm", commercial_profile.get("profile_reference_thickness_mm", 3.0)),
            3.0,
        ),
    )
    reference_row = _lookup_cut_row(gas_rows, reference_thickness)
    current_speed = _effective_speed_m_min(cut_row, motion_cfg)
    reference_speed = _effective_speed_m_min(reference_row or cut_row, motion_cfg)
    raw_factor = _safe_div(reference_speed, current_speed, 1.0)
    min_factor = max(
        0.1,
        _as_float(payload.get("profile_min_thickness_rate_factor", commercial_profile.get("profile_min_thickness_rate_factor", 0.45)), 0.45),
    )
    max_factor = max(
        min_factor,
        _as_float(payload.get("profile_max_thickness_rate_factor", commercial_profile.get("profile_max_thickness_rate_factor", 8.0)), 8.0),
    )
    factor = min(max(raw_factor, min_factor), max_factor)
    return round(factor, 4), round(reference_speed, 4), round(current_speed, 4)


def _build_description(file_path: str, material_name: str, thickness_mm: float, bbox_mm: dict[str, Any], part_count_hint: int) -> str:
    stem = _coerce_label(Path(str(file_path or "").strip()).stem)
    size_txt = ""
    width = _as_float(bbox_mm.get("width", 0), 0.0)
    height = _as_float(bbox_mm.get("height", 0), 0.0)
    if width > 0.0 and height > 0.0:
        size_txt = f" {int(round(width))}x{int(round(height))} mm"
    part_txt = f" x{part_count_hint}" if int(part_count_hint or 0) > 1 else ""
    thickness_txt = f"{_round_mm(thickness_mm, 1):g} mm"
    return f"{stem}{part_txt} | Corte laser {material_name} {thickness_txt}{size_txt}".strip()


def _build_profile_description(
    file_path: str,
    material_name: str,
    thickness_mm: float,
    family_name: str,
    section_name: str,
    total_cut_count: int,
    hole_count: int,
    slot_count: int,
) -> str:
    stem = _coerce_label(Path(str(file_path or "").strip()).stem)
    profile_txt = str(family_name or "perfil").strip().lower()
    section_txt = str(section_name or "").strip()
    features: list[str] = []
    if total_cut_count > 0:
        features.append(f"{int(total_cut_count)} cortes")
    if hole_count > 0:
        features.append(f"{int(hole_count)} furos")
    if slot_count > 0:
        features.append(f"{int(slot_count)} rasgos")
    suffix = ""
    if features:
        suffix = " | " + " / ".join(features)
    return (
        f"{stem} | Corte laser STEP/IGS {profile_txt} {material_name} "
        f"{_round_mm(thickness_mm, 1):g} mm"
        f"{f' | {section_txt}' if section_txt else ''}"
        f"{suffix}"
    ).strip()


def estimate_laser_quote(payload: dict[str, Any], settings: dict[str, Any] | None = None) -> dict[str, Any]:
    payload = dict(payload or {})
    merged_settings = merge_laser_quote_settings(settings)
    machine_profile = _active_machine(merged_settings, str(payload.get("machine_name", "") or "").strip())
    commercial_profile = _active_commercial(merged_settings, str(payload.get("commercial_name", "") or "").strip())
    layer_rules = dict(merged_settings.get("layer_rules", {}) or {})
    geometry = analyze_dxf_geometry(str(payload.get("path", "") or "").strip(), layer_rules)
    thickness_mm = max(0.1, _as_float(payload.get("thickness_mm", 0), 0.0))
    quantity = max(1, _as_int(payload.get("qtd", payload.get("quantity", 1)), 1))
    requested_material_family = str(payload.get("material", "") or "").strip() or "Aco carbono"
    requested_material_subtype = str(payload.get("material_subtype", "") or "").strip()
    material_family, material_subtype = _infer_material_family_and_subtype(merged_settings, requested_material_family, requested_material_subtype)
    material_family = _canonical_material_family(material_family) or "Aco carbono"
    display_material_family = _display_material_family(material_family)
    display_material = material_subtype or display_material_family
    material_supplied_by_client = bool(
        payload.get("material_supplied_by_client", False)
        or payload.get("material_fornecido_cliente", False)
        or payload.get("exclude_material_cost", False)
    )
    material_cost_included = not material_supplied_by_client
    gas_name, gas_rows, machine_material = _gas_rows(machine_profile, material_family, str(payload.get("gas", "") or "").strip(), material_subtype)
    if not gas_rows:
        raise ValueError(f"Nao existe tabela de corte configurada para {material_family} / {gas_name or 'gas'}")
    cut_row = _lookup_cut_row(gas_rows, thickness_mm)
    if not cut_row:
        raise ValueError("Nao foi possivel encontrar parametros de corte para esta espessura.")
    motion_cfg = dict(machine_profile.get("motion", {}) or {})
    commercial_material = _commercial_material(commercial_profile, material_family, material_subtype)
    commercial_subtype_override = _commercial_material_subtype_override(commercial_profile, material_family, material_subtype)
    pricing_material = dict(commercial_material or {})
    pricing_material.update({key: value for key, value in dict(commercial_subtype_override or {}).items() if value is not None})
    metrics = dict(geometry.get("metrics", {}) or {})
    cut_length_m = _as_float(payload.get("cut_length_m_override", metrics.get("cut_length_m", 0)), metrics.get("cut_length_m", 0))
    mark_length_m = _as_float(payload.get("mark_length_m_override", metrics.get("mark_length_m", 0)), metrics.get("mark_length_m", 0))
    pierce_count = max(0, _as_int(payload.get("pierce_count_override", metrics.get("pierce_count", 0)), metrics.get("pierce_count", 0)))
    net_area_m2 = _as_float(payload.get("net_area_m2_override", metrics.get("net_area_m2", 0)), metrics.get("net_area_m2", 0))
    bbox_area_m2 = _as_float(metrics.get("bbox_area_m2", 0), 0.0)
    rapid_length_m = _as_float(metrics.get("rapid_length_m", 0), 0.0)
    if net_area_m2 <= 0.0:
        fallback_fill_pct = max(1.0, _as_float(payload.get("fallback_fill_pct", commercial_profile.get("fallback_fill_pct", 72.0)), 72.0))
        net_area_m2 = bbox_area_m2 * (fallback_fill_pct / 100.0)
    density_kg_m3 = max(
        1.0,
        _as_float(
            payload.get("density_kg_m3", pricing_material.get("density_kg_m3", machine_material.get("density_kg_m3", 7800.0))),
            7800.0,
        ),
    )
    material_price_source = "subtype" if commercial_subtype_override else "family"
    configured_material_price_per_kg = max(0.0, _as_float(payload.get("material_price_per_kg", pricing_material.get("price_per_kg", 0.0)), 0.0))
    configured_scrap_credit_per_kg = max(0.0, _as_float(payload.get("scrap_credit_per_kg", pricing_material.get("scrap_credit_per_kg", 0.0)), 0.0))
    material_price_per_kg = 0.0 if material_supplied_by_client else configured_material_price_per_kg
    scrap_credit_per_kg = 0.0 if material_supplied_by_client else configured_scrap_credit_per_kg
    utilization_pct = max(1.0, _as_float(payload.get("material_utilization_pct", commercial_profile.get("material_utilization_pct", 82.0)), 82.0))
    rates = dict(commercial_profile.get("rates", {}) or {})
    cut_rate_per_m = max(0.0, _as_float(payload.get("cut_rate_per_m", rates.get("cut_per_m_eur", 0.0)), 0.0))
    mark_rate_per_m = max(0.0, _as_float(payload.get("mark_rate_per_m", rates.get("marking_per_m_eur", 0.0)), 0.0))
    defilm_rate_per_m = max(0.0, _as_float(payload.get("defilm_rate_per_m", rates.get("defilm_per_m_eur", 0.0)), 0.0))
    pierce_rate = max(0.0, _as_float(payload.get("pierce_rate", rates.get("pierce_eur", 0.0)), 0.0))
    machine_hour_eur = max(0.0, _as_float(payload.get("machine_hour_eur", rates.get("machine_hour_eur", 0.0)), 0.0))
    base_margin_pct = max(0.0, _as_float(payload.get("margin_pct", commercial_profile.get("margin_pct", 0.0)), 0.0))
    base_setup_time_min = max(0.0, _as_float(payload.get("setup_time_min", commercial_profile.get("setup_time_min", 0.0)), 0.0))
    handling_eur = max(0.0, _as_float(payload.get("handling_eur", commercial_profile.get("handling_eur", 0.0)), 0.0))
    include_marking = bool(payload.get("include_marking", mark_length_m > 0.0))
    include_defilm = bool(payload.get("include_defilm", False))
    use_scrap_credit = bool(payload.get("use_scrap_credit", commercial_profile.get("use_scrap_credit", True))) and not material_supplied_by_client
    cost_mode = str(payload.get("cost_mode", commercial_profile.get("cost_mode", "hybrid_max")) or "hybrid_max").strip().lower()
    series_tier = _pick_series_tier(commercial_profile, quantity)
    series_margin_delta_pct = _as_float(series_tier.get("margin_delta_pct", 0.0), 0.0)
    series_setup_multiplier = max(0.05, _as_float(series_tier.get("setup_multiplier", 1.0), 1.0))
    margin_pct = max(0.0, base_margin_pct + series_margin_delta_pct)
    setup_time_min = max(0.0, base_setup_time_min * series_setup_multiplier)
    effective_cut_speed_m_min = _effective_speed_m_min(cut_row, motion_cfg)
    effective_mark_speed_m_min = max(0.1, _as_float(payload.get("mark_speed_m_min", motion_cfg.get("mark_speed_m_min", 18.0)), 18.0))
    rapid_speed_mm_s = max(1.0, _as_float(payload.get("rapid_speed_mm_s", motion_cfg.get("rapid_speed_mm_s", 200.0)), 200.0))
    lead_in_mm = max(0.0, _as_float(payload.get("lead_in_mm", motion_cfg.get("lead_in_mm", 2.0)), 2.0))
    lead_out_mm = max(0.0, _as_float(payload.get("lead_out_mm", motion_cfg.get("lead_out_mm", 2.0)), 2.0))
    lead_move_speed_mm_s = max(0.1, _as_float(payload.get("lead_move_speed_mm_s", motion_cfg.get("lead_move_speed_mm_s", 3.0)), 3.0))
    pierce_sec_each = max(
        0.0,
        (
            _as_float(payload.get("pierce_base_ms", motion_cfg.get("pierce_base_ms", 400.0)), 400.0)
            + (thickness_mm * _as_float(payload.get("pierce_per_mm_ms", motion_cfg.get("pierce_per_mm_ms", 35.0)), 35.0))
        )
        / 1000.0,
    )
    first_gas_delay_sec = max(0.0, _as_float(payload.get("first_gas_delay_ms", motion_cfg.get("first_gas_delay_ms", 200.0)), 200.0) / 1000.0)
    gas_delay_sec = max(0.0, _as_float(payload.get("gas_delay_ms", motion_cfg.get("gas_delay_ms", 0.0)), 0.0) / 1000.0)
    motion_overhead_factor = 1.0 + (max(0.0, _as_float(payload.get("motion_overhead_pct", motion_cfg.get("motion_overhead_pct", 4.0)), 4.0)) / 100.0)
    cut_time_sec = _safe_div(cut_length_m, effective_cut_speed_m_min / 60.0, 0.0)
    mark_time_sec = _safe_div(mark_length_m if include_marking else 0.0, effective_mark_speed_m_min / 60.0, 0.0)
    rapid_time_sec = _safe_div(rapid_length_m * 1000.0, rapid_speed_mm_s, 0.0)
    lead_time_sec = pierce_count * _safe_div(lead_in_mm + lead_out_mm, lead_move_speed_mm_s, 0.0)
    pierce_time_sec = pierce_count * pierce_sec_each
    gas_time_sec = (first_gas_delay_sec if pierce_count > 0 else 0.0) + (max(0, pierce_count - 1) * gas_delay_sec)
    setup_time_sec_share = _safe_div(setup_time_min * 60.0, quantity, 0.0)
    machine_time_sec_unit = (cut_time_sec + mark_time_sec + rapid_time_sec + lead_time_sec + pierce_time_sec + gas_time_sec) * motion_overhead_factor
    machine_time_total_unit = machine_time_sec_unit + setup_time_sec_share
    thickness_m = thickness_mm / 1000.0
    net_mass_kg = max(0.0, net_area_m2 * thickness_m * density_kg_m3)
    gross_mass_kg = max(net_mass_kg, net_mass_kg / max(0.01, utilization_pct / 100.0))
    scrap_mass_kg = max(0.0, gross_mass_kg - net_mass_kg)
    material_cost_unit = gross_mass_kg * material_price_per_kg
    scrap_credit_unit = scrap_mass_kg * scrap_credit_per_kg if use_scrap_credit else 0.0
    material_net_cost_unit = max(0.0, material_cost_unit - scrap_credit_unit)
    cut_cost_meter_unit = cut_length_m * cut_rate_per_m
    mark_cost_unit = (mark_length_m if include_marking else 0.0) * mark_rate_per_m
    defilm_cost_unit = (cut_length_m if include_defilm else 0.0) * defilm_rate_per_m
    pierce_cost_unit = pierce_count * pierce_rate
    machine_runtime_cost_unit = (machine_time_total_unit / 3600.0) * machine_hour_eur
    if cost_mode == "per_meter":
        effective_cutting_cost_unit = cut_cost_meter_unit
        effective_cutting_label = "Corte por metro"
    elif cost_mode == "machine_time":
        effective_cutting_cost_unit = machine_runtime_cost_unit
        effective_cutting_label = "Tempo maquina"
    else:
        effective_cutting_cost_unit = max(cut_cost_meter_unit, machine_runtime_cost_unit)
        effective_cutting_label = "Hibrido (maior entre metro e tempo)"
    subtotal_cost_unit = (
        material_net_cost_unit
        + effective_cutting_cost_unit
        + mark_cost_unit
        + defilm_cost_unit
        + pierce_cost_unit
        + handling_eur
    )
    unit_price_before_min = subtotal_cost_unit * (1.0 + (margin_pct / 100.0))
    total_price_before_min = unit_price_before_min * quantity
    minimum_applies = False
    minimum_line_eur = 0.0
    final_total = total_price_before_min
    final_unit = _safe_div(final_total, quantity, unit_price_before_min)
    file_path = str(payload.get("path", "") or geometry.get("file_path", "")).strip()
    description = str(payload.get("description", "") or "").strip() or _build_description(
        file_path,
        display_material,
        thickness_mm,
        dict(geometry.get("bbox_mm", {}) or {}),
        _as_int(metrics.get("part_count_hint", 1), 1),
    )
    ref_externa = str(payload.get("ref_externa", "") or "").strip() or _sanitize_file_stem(file_path)
    operations = ["Corte Laser"]
    if include_marking and mark_length_m > 0.0:
        operations.append("Marcacao")
    if include_defilm:
        operations.append("Defilm")
    warnings = list(geometry.get("warnings", []) or [])
    if _as_float(metrics.get("net_area_m2", 0), 0.0) <= 0.0:
        warnings.append("A area liquida foi estimada a partir da caixa envolvente. Confirma o custo do material.")
    if material_supplied_by_client:
        warnings.append("Materia-prima fornecida pelo cliente: o orcamento considera apenas trabalho e processo.")
    elif configured_material_price_per_kg <= 0.0:
        warnings.append(f"Sem preco comercial configurado para {display_material}. Define EUR/kg no perfil comercial antes de usar este valor como orçamento final.")
    return {
        "file_path": file_path,
        "geometry": geometry,
        "machine": {
            "name": str(machine_profile.get("name", "") or merged_settings.get("active_machine", "")),
            "material": display_material,
            "material_family": material_family,
            "material_subtype": material_subtype,
            "material_supplied_by_client": material_supplied_by_client,
            "material_cost_included": material_cost_included,
            "gas": gas_name,
            "thickness_mm": round(thickness_mm, 3),
            "cut_row": cut_row,
            "motion": motion_cfg,
            "effective_cut_speed_m_min": round(effective_cut_speed_m_min, 4),
            "mark_speed_m_min": round(effective_mark_speed_m_min, 4),
        },
        "commercial": {
            "name": str(commercial_profile.get("name", "") or merged_settings.get("active_commercial", "")),
            "cost_mode": cost_mode,
            "effective_cutting_label": effective_cutting_label,
            "series_key": str(series_tier.get("key", "") or ""),
            "series_label": str(series_tier.get("label", "") or ""),
            "series_margin_delta_pct": round(series_margin_delta_pct, 3),
            "series_setup_multiplier": round(series_setup_multiplier, 4),
            "base_margin_pct": round(base_margin_pct, 3),
            "margin_pct": round(margin_pct, 3),
            "minimum_line_eur": round(minimum_line_eur, 2),
            "minimum_applied": bool(minimum_applies),
            "base_setup_time_min": round(base_setup_time_min, 3),
            "setup_time_min": round(setup_time_min, 3),
            "handling_eur": round(handling_eur, 2),
            "material_utilization_pct": round(utilization_pct, 3),
            "material_supplied_by_client": material_supplied_by_client,
            "material_cost_included": material_cost_included,
            "material_price_source": material_price_source,
        },
        "metrics": {
            **metrics,
            "cut_length_m": round(cut_length_m, 4),
            "mark_length_m": round(mark_length_m, 4),
            "rapid_length_m": round(rapid_length_m, 4),
            "pierce_count": int(pierce_count),
            "net_area_m2": round(net_area_m2, 6),
            "density_kg_m3": round(density_kg_m3, 3),
            "net_mass_kg": round(net_mass_kg, 4),
            "gross_mass_kg": round(gross_mass_kg, 4),
            "scrap_mass_kg": round(scrap_mass_kg, 4),
        },
        "times": {
            "cut_sec": round(cut_time_sec, 3),
            "mark_sec": round(mark_time_sec, 3),
            "rapid_sec": round(rapid_time_sec, 3),
            "lead_sec": round(lead_time_sec, 3),
            "pierce_sec": round(pierce_time_sec, 3),
            "gas_sec": round(gas_time_sec, 3),
            "setup_share_sec": round(setup_time_sec_share, 3),
            "machine_total_sec": round(machine_time_total_unit, 3),
            "machine_total_min": round(machine_time_total_unit / 60.0, 3),
        },
        "pricing": {
            "material_cost_unit": _round_money(material_net_cost_unit, 4),
            "material_gross_cost_unit": _round_money(material_cost_unit, 4),
            "scrap_credit_unit": _round_money(scrap_credit_unit, 4),
            "process_only_cost_unit": _round_money(max(0.0, subtotal_cost_unit - material_net_cost_unit), 4),
            "material_cost_included": material_cost_included,
            "cut_cost_meter_unit": _round_money(cut_cost_meter_unit, 4),
            "machine_runtime_cost_unit": _round_money(machine_runtime_cost_unit, 4),
            "effective_cutting_cost_unit": _round_money(effective_cutting_cost_unit, 4),
            "mark_cost_unit": _round_money(mark_cost_unit, 4),
            "defilm_cost_unit": _round_money(defilm_cost_unit, 4),
            "pierce_cost_unit": _round_money(pierce_cost_unit, 4),
            "handling_unit": _round_money(handling_eur, 4),
            "subtotal_cost_unit": _round_money(subtotal_cost_unit, 4),
            "unit_price_before_min": _round_money(unit_price_before_min, 4),
            "unit_price": _round_money(final_unit, 4),
            "total_price": _round_money(final_total, 2),
            "quantity": quantity,
        },
        "line_suggestion": {
            "tipo_item": "peca",
            "ref_externa": ref_externa,
            "descricao": description,
            "material": display_material,
            "material_family": material_family,
            "material_subtype": material_subtype,
            "material_supplied_by_client": material_supplied_by_client,
            "material_fornecido_cliente": material_supplied_by_client,
            "material_cost_included": material_cost_included,
            "espessura": f"{_round_mm(thickness_mm, 3):g}",
            "operacao": " + ".join(operations),
            "tempo_peca_min": round(machine_time_total_unit / 60.0, 3),
            "qtd": quantity,
            "preco_unit": round(final_unit, 4),
            "desenho": file_path,
        },
        "warnings": list(dict.fromkeys([str(item or "").strip() for item in warnings if str(item or "").strip()])),
        "debug_json": json.dumps(
            {
                "machine": str(machine_profile.get("name", "") or ""),
                "commercial": str(commercial_profile.get("name", "") or ""),
                "material_family": material_family,
                "material_subtype": material_subtype,
                "material": display_material,
                "material_supplied_by_client": material_supplied_by_client,
                "material_price_source": material_price_source,
                "series_key": str(series_tier.get("key", "") or ""),
                "series_label": str(series_tier.get("label", "") or ""),
                "gas": gas_name,
                "thickness_mm": thickness_mm,
                "quantity": quantity,
            },
            ensure_ascii=False,
        ),
    }


def estimate_profile_laser_quote(payload: dict[str, Any], settings: dict[str, Any] | None = None) -> dict[str, Any]:
    payload = dict(payload or {})
    merged_settings = merge_laser_quote_settings(settings)
    machine_profile = _active_machine(merged_settings, str(payload.get("machine_name", "") or "").strip())
    commercial_profile = _active_commercial(merged_settings, str(payload.get("commercial_name", "") or "").strip())
    file_path = str(payload.get("path", "") or "").strip()
    thickness_mm = max(0.1, _as_float(payload.get("thickness_mm", 0), 0.0))
    quantity = max(1, _as_int(payload.get("qtd", payload.get("quantity", 1)), 1))
    requested_material_family = str(payload.get("material", "") or "").strip() or "Aco carbono"
    requested_material_subtype = str(payload.get("material_subtype", "") or "").strip()
    material_family, material_subtype = _infer_material_family_and_subtype(merged_settings, requested_material_family, requested_material_subtype)
    material_family = _canonical_material_family(material_family) or "Aco carbono"
    display_material_family = _display_material_family(material_family)
    display_material = material_subtype or display_material_family
    if any(key in payload for key in ("material_supplied_by_client", "material_fornecido_cliente", "exclude_material_cost")):
        material_supplied_by_client = bool(
            payload.get("material_supplied_by_client", False)
            or payload.get("material_fornecido_cliente", False)
            or payload.get("exclude_material_cost", False)
        )
    else:
        material_supplied_by_client = True
    material_cost_included = not material_supplied_by_client
    gas_name, gas_rows, machine_material = _gas_rows(machine_profile, material_family, str(payload.get("gas", "") or "").strip(), material_subtype)
    if not gas_rows:
        raise ValueError(f"Nao existe tabela de corte configurada para {material_family} / {gas_name or 'gas'}")
    cut_row = _lookup_cut_row(gas_rows, thickness_mm)
    if not cut_row:
        raise ValueError("Nao foi possivel encontrar parametros de corte para esta espessura.")

    analyzed_metrics = dict(payload.get("profile_metrics", {}) or {})
    include_external_cuts = str(payload.get("include_external_profile_cuts", "0") or "0").strip().lower() in {
        "1",
        "true",
        "yes",
        "sim",
        "on",
    }
    explicit_counts = any(
        payload.get(key) not in (None, "")
        for key in ("cuts", "holes", "slots", "outer_cuts", "end_cut_count", "generic_cuts")
    )
    if not analyzed_metrics and file_path and not explicit_counts:
        try:
            analyzed_metrics = dict(analyze_profile_cut_features(file_path) or {})
        except Exception:
            analyzed_metrics = {}

    hole_count = max(0, _as_int(payload.get("holes", analyzed_metrics.get("holes", 0)), analyzed_metrics.get("holes", 0)))
    slot_count = max(0, _as_int(payload.get("slots", analyzed_metrics.get("slots", 0)), analyzed_metrics.get("slots", 0)))
    raw_end_cut_count = max(0, _as_int(payload.get("end_cut_count", analyzed_metrics.get("end_cut_count", 0)), analyzed_metrics.get("end_cut_count", 0)))
    end_cut_count = raw_end_cut_count if include_external_cuts else 0
    generic_cut_count = max(0, _as_int(payload.get("generic_cuts", analyzed_metrics.get("generic_cuts", 0)), analyzed_metrics.get("generic_cuts", 0)))
    fallback_total_cut_count = hole_count + slot_count + generic_cut_count + end_cut_count
    total_cut_count = max(0, _as_int(payload.get("cuts", fallback_total_cut_count), fallback_total_cut_count))
    derived_outer_cuts = max(0, total_cut_count - hole_count - slot_count)
    outer_cut_count = max(
        0,
        _as_int(
            payload.get("outer_cuts", generic_cut_count + end_cut_count),
            generic_cut_count + end_cut_count,
        ),
    )
    if not include_external_cuts:
        outer_cut_count = min(outer_cut_count, generic_cut_count)
        total_cut_count = hole_count + slot_count + outer_cut_count
    else:
        total_cut_count = max(total_cut_count, hole_count + slot_count + outer_cut_count)

    motion_cfg = dict(machine_profile.get("motion", {}) or {})
    commercial_material = _commercial_material(commercial_profile, material_family, material_subtype)
    commercial_subtype_override = _commercial_material_subtype_override(commercial_profile, material_family, material_subtype)
    pricing_material = dict(commercial_material or {})
    pricing_material.update({key: value for key, value in dict(commercial_subtype_override or {}).items() if value is not None})
    profile_operations = dict(commercial_profile.get("profile_operations", {}) or {})
    include_profile_event_rates = str(
        payload.get("include_profile_event_rates", commercial_profile.get("include_profile_event_rates", "0")) or "0"
    ).strip().lower() in {"1", "true", "yes", "sim", "on"}
    include_profile_setup = str(
        payload.get("include_profile_setup", commercial_profile.get("include_profile_setup", "0")) or "0"
    ).strip().lower() in {"1", "true", "yes", "sim", "on"}
    density_kg_m3 = max(
        1.0,
        _as_float(
            payload.get("density_kg_m3", pricing_material.get("density_kg_m3", machine_material.get("density_kg_m3", 7800.0))),
            7800.0,
        ),
    )
    material_price_source = "subtype" if commercial_subtype_override else "family"
    configured_material_price_per_kg = max(0.0, _as_float(payload.get("material_price_per_kg", pricing_material.get("price_per_kg", 0.0)), 0.0))
    configured_scrap_credit_per_kg = max(0.0, _as_float(payload.get("scrap_credit_per_kg", pricing_material.get("scrap_credit_per_kg", 0.0)), 0.0))
    material_price_per_kg = 0.0 if material_supplied_by_client else configured_material_price_per_kg
    scrap_credit_per_kg = 0.0 if material_supplied_by_client else configured_scrap_credit_per_kg
    utilization_pct = max(1.0, _as_float(payload.get("material_utilization_pct", commercial_profile.get("material_utilization_pct", 82.0)), 82.0))
    rates = dict(commercial_profile.get("rates", {}) or {})
    cut_rate_per_m = max(0.0, _as_float(payload.get("cut_rate_per_m", rates.get("cut_per_m_eur", 0.0)), 0.0))
    pierce_rate = max(0.0, _as_float(payload.get("pierce_rate", rates.get("pierce_eur", 0.0)), 0.0))
    outer_cut_rate = max(0.0, _as_float(payload.get("outer_cut_rate", profile_operations.get("outer_cut_eur", 1.25)), 1.25))
    hole_cut_rate = max(0.0, _as_float(payload.get("hole_cut_rate", profile_operations.get("hole_cut_eur", 0.85)), 0.85))
    slot_cut_rate = max(0.0, _as_float(payload.get("slot_cut_rate", profile_operations.get("slot_cut_eur", 1.15)), 1.15))
    machine_hour_eur = max(0.0, _as_float(payload.get("machine_hour_eur", rates.get("machine_hour_eur", 0.0)), 0.0))
    base_margin_pct = max(0.0, _as_float(payload.get("margin_pct", commercial_profile.get("margin_pct", 0.0)), 0.0))
    base_setup_time_min = max(0.0, _as_float(payload.get("setup_time_min", commercial_profile.get("setup_time_min", 0.0)), 0.0))
    handling_eur = max(0.0, _as_float(payload.get("handling_eur", commercial_profile.get("handling_eur", 0.0)), 0.0))
    cost_mode = str(payload.get("cost_mode", commercial_profile.get("cost_mode", "hybrid_max")) or "hybrid_max").strip().lower()
    series_tier = _pick_series_tier(commercial_profile, quantity)
    series_margin_delta_pct = _as_float(series_tier.get("margin_delta_pct", 0.0), 0.0)
    series_setup_multiplier = max(0.05, _as_float(series_tier.get("setup_multiplier", 1.0), 1.0))
    margin_pct = max(0.0, base_margin_pct + series_margin_delta_pct)
    setup_time_min = max(0.0, base_setup_time_min * series_setup_multiplier) if include_profile_setup else 0.0
    effective_cut_speed_m_min = _effective_speed_m_min(cut_row, motion_cfg)
    thickness_rate_factor, reference_cut_speed_m_min, current_cut_speed_m_min = _profile_thickness_rate_factor(
        gas_rows,
        cut_row,
        motion_cfg,
        thickness_mm,
        commercial_profile,
        payload,
    )
    scaled_cut_rate_per_m = cut_rate_per_m * thickness_rate_factor
    scaled_pierce_rate = pierce_rate * max(1.0, math.sqrt(max(0.0, thickness_rate_factor)))
    scaled_outer_cut_rate = outer_cut_rate * thickness_rate_factor
    scaled_hole_cut_rate = hole_cut_rate * thickness_rate_factor
    scaled_slot_cut_rate = slot_cut_rate * thickness_rate_factor
    rapid_speed_mm_s = max(1.0, _as_float(payload.get("rapid_speed_mm_s", motion_cfg.get("rapid_speed_mm_s", 200.0)), 200.0))
    lead_in_mm = max(0.0, _as_float(payload.get("lead_in_mm", motion_cfg.get("lead_in_mm", 2.0)), 2.0))
    lead_out_mm = max(0.0, _as_float(payload.get("lead_out_mm", motion_cfg.get("lead_out_mm", 2.0)), 2.0))
    lead_move_speed_mm_s = max(0.1, _as_float(payload.get("lead_move_speed_mm_s", motion_cfg.get("lead_move_speed_mm_s", 3.0)), 3.0))
    pierce_sec_each = max(
        0.0,
        (
            _as_float(payload.get("pierce_base_ms", motion_cfg.get("pierce_base_ms", 400.0)), 400.0)
            + (thickness_mm * _as_float(payload.get("pierce_per_mm_ms", motion_cfg.get("pierce_per_mm_ms", 35.0)), 35.0))
        )
        / 1000.0,
    )
    first_gas_delay_sec = max(0.0, _as_float(payload.get("first_gas_delay_ms", motion_cfg.get("first_gas_delay_ms", 200.0)), 200.0) / 1000.0)
    gas_delay_sec = max(0.0, _as_float(payload.get("gas_delay_ms", motion_cfg.get("gas_delay_ms", 0.0)), 0.0) / 1000.0)
    motion_overhead_factor = 1.0 + (max(0.0, _as_float(payload.get("motion_overhead_pct", motion_cfg.get("motion_overhead_pct", 4.0)), 4.0)) / 100.0)
    setup_time_sec_share = _safe_div(setup_time_min * 60.0, quantity, 0.0) if include_profile_setup else 0.0
    pierce_count = total_cut_count
    lead_time_sec = pierce_count * _safe_div(lead_in_mm + lead_out_mm, lead_move_speed_mm_s, 0.0)
    pierce_time_sec = pierce_count * pierce_sec_each
    gas_time_sec = (first_gas_delay_sec if pierce_count > 0 else 0.0) + (max(0, pierce_count - 1) * gas_delay_sec)
    rapid_jump_mm = max(0.0, _as_float(payload.get("profile_rapid_between_cuts_mm", 12.0), 12.0))
    rapid_length_m = max(0.0, (max(0, pierce_count - 1) * rapid_jump_mm) / 1000.0)
    rapid_time_sec = _safe_div(rapid_length_m * 1000.0, rapid_speed_mm_s, 0.0)
    family_txt = str(payload.get("profile_family", payload.get("family", "Perfil")) or "Perfil").strip()
    section_txt = str(payload.get("section", payload.get("section_label", "")) or "").strip()
    cut_length_m = max(
        0.0,
        _as_float(
            payload.get("cut_length_m_override", analyzed_metrics.get("cut_length_m", 0.0)),
            _as_float(analyzed_metrics.get("cut_length_m", 0.0), 0.0),
        ),
    )
    if not include_external_cuts:
        internal_length_m = (
            _as_float(analyzed_metrics.get("feature_cut_length_m", 0.0), 0.0)
            or (
                (_as_float(analyzed_metrics.get("hole_cut_length_mm", 0.0), 0.0)
                + _as_float(analyzed_metrics.get("slot_cut_length_mm", 0.0), 0.0)
                + _as_float(analyzed_metrics.get("generic_cut_length_mm", 0.0), 0.0))
                / 1000.0
            )
        )
        if internal_length_m > 0.0:
            cut_length_m = internal_length_m
    if cut_length_m <= 0.0:
        cut_length_m = _estimate_profile_cut_length_m(
            total_cut_count=total_cut_count,
            hole_count=hole_count,
            slot_count=slot_count,
            outer_cut_count=outer_cut_count,
            thickness_mm=thickness_mm,
            section=section_txt,
            family=family_txt,
        )
    cut_time_sec = _safe_div(cut_length_m, effective_cut_speed_m_min / 60.0, 0.0)
    machine_time_sec_unit = (cut_time_sec + lead_time_sec + pierce_time_sec + gas_time_sec + rapid_time_sec) * motion_overhead_factor
    machine_time_total_unit = machine_time_sec_unit + setup_time_sec_share
    net_area_m2 = max(0.0, _as_float(payload.get("net_area_m2_override", 0.0), 0.0))
    thickness_m = thickness_mm / 1000.0
    net_mass_kg = max(0.0, net_area_m2 * thickness_m * density_kg_m3)
    gross_mass_kg = max(net_mass_kg, net_mass_kg / max(0.01, utilization_pct / 100.0)) if net_area_m2 > 0.0 else 0.0
    scrap_mass_kg = max(0.0, gross_mass_kg - net_mass_kg)
    profile_length_m = max(0.0, _as_float(payload.get("profile_length_m", payload.get("material_length_m", 0.0)), 0.0))
    profile_kg_m = max(0.0, _as_float(payload.get("profile_kg_m", payload.get("kg_m", 0.0)), 0.0))
    if profile_kg_m <= 0.0 and profile_length_m > 0.0:
        profile_kg_m = _profile_section_kg_m(section_txt, family_txt, thickness_mm, density_kg_m3)
    profile_mass_kg = profile_kg_m * profile_length_m if profile_kg_m > 0.0 and profile_length_m > 0.0 else 0.0
    material_price_unit = str(payload.get("profile_material_price_unit", payload.get("material_price_unit", "EUR/kg")) or "EUR/kg").strip().lower()
    family_norm = _norm_material_token(family_txt)
    if any(token in family_norm for token in ("TUBO", "TUBE")):
        material_price_unit = "eur/m"
    profile_material_price = max(0.0, _as_float(payload.get("profile_material_price", payload.get("material_price", 0.0)), 0.0))
    explicit_profile_material_cost = 0.0
    if profile_material_price > 0.0 and profile_length_m > 0.0:
        if material_price_unit in {"eur/m", "m", "metro", "metros"}:
            explicit_profile_material_cost = profile_length_m * profile_material_price
        elif material_price_unit in {"eur/t", "eur/ton", "ton", "tonelada", "toneladas"}:
            explicit_profile_material_cost = profile_mass_kg * (profile_material_price / 1000.0)
        else:
            explicit_profile_material_cost = profile_mass_kg * profile_material_price
    if explicit_profile_material_cost > 0.0 and not material_supplied_by_client:
        material_cost_unit = explicit_profile_material_cost
        scrap_credit_unit = 0.0
        material_net_cost_unit = material_cost_unit
        gross_mass_kg = max(gross_mass_kg, profile_mass_kg)
        net_mass_kg = max(net_mass_kg, profile_mass_kg)
        scrap_mass_kg = 0.0
        material_price_source = "profile"
    else:
        material_cost_unit = gross_mass_kg * material_price_per_kg
        scrap_credit_unit = scrap_mass_kg * scrap_credit_per_kg if gross_mass_kg > 0.0 else 0.0
        material_net_cost_unit = max(0.0, material_cost_unit - scrap_credit_unit)
    profile_operation_cost_unit = 0.0
    if include_profile_event_rates:
        profile_operation_cost_unit = (
            (outer_cut_count * scaled_outer_cut_rate)
            + (hole_count * scaled_hole_cut_rate)
            + (slot_count * scaled_slot_cut_rate)
        )
    cut_cost_meter_unit = cut_length_m * scaled_cut_rate_per_m
    pierce_cost_unit = pierce_count * scaled_pierce_rate
    profile_process_cost_unit = cut_cost_meter_unit + pierce_cost_unit + profile_operation_cost_unit
    machine_runtime_cost_unit = (machine_time_total_unit / 3600.0) * machine_hour_eur
    if cost_mode == "machine_time":
        effective_cutting_cost_unit = machine_runtime_cost_unit + pierce_cost_unit + profile_operation_cost_unit
        effective_cutting_label = "Tempo maquina + penetracoes"
    elif cost_mode == "per_meter":
        effective_cutting_cost_unit = profile_process_cost_unit
        effective_cutting_label = "Metros + penetracoes"
    else:
        effective_cutting_cost_unit = max(cut_cost_meter_unit, machine_runtime_cost_unit) + pierce_cost_unit + profile_operation_cost_unit
        effective_cutting_label = "Hibrido DXF (maior entre metro e tempo + penetracoes)"
    subtotal_cost_unit = material_net_cost_unit + effective_cutting_cost_unit + handling_eur
    unit_price_before_min = subtotal_cost_unit * (1.0 + (margin_pct / 100.0))
    final_total = unit_price_before_min * quantity
    final_unit = _safe_div(final_total, quantity, unit_price_before_min)
    description = str(payload.get("description", "") or "").strip() or _build_profile_description(
        file_path,
        display_material,
        thickness_mm,
        family_txt,
        section_txt,
        total_cut_count,
        hole_count,
        slot_count,
    )
    ref_externa = str(payload.get("ref_externa", "") or "").strip() or _sanitize_file_stem(file_path)
    warnings: list[str] = []
    if material_supplied_by_client:
        warnings.append("Perfil/tubo assumido como materia fornecida pelo cliente: o calculo considera apenas processo laser.")
    if not include_external_cuts and raw_end_cut_count > 0:
        warnings.append(f"Cortes de extremidade ignorados por criterio comercial: {raw_end_cut_count}.")
    if total_cut_count <= 0:
        warnings.append("Sem cortes detetados automaticamente. Confirma os contadores antes de aplicar.")
    elif hole_count > total_cut_count:
        warnings.append("Os furos excedem o numero total de cortes. O total foi ajustado para evitar dupla contagem.")
    if _as_float(analyzed_metrics.get("cut_length_m", 0.0), 0.0) <= 0.0 and cut_length_m > 0.0:
        warnings.append("Comprimento de corte estimado por fallback. Se possivel confirma com preview/FreeCAD antes de fechar o preco.")
    if thickness_rate_factor > 1.15:
        warnings.append(
            f"Precos de perfil ajustados pela tabela da maquina para {thickness_mm:g} mm (fator x{thickness_rate_factor:.2f})."
        )
    if configured_material_price_per_kg <= 0.0 and not material_supplied_by_client:
        if explicit_profile_material_cost <= 0.0:
            warnings.append(f"Sem preco comercial configurado para {display_material}.")
    if (
        any(token in family_norm for token in ("TUBO", "TUBE"))
        and not material_supplied_by_client
        and profile_material_price <= 0.0
    ):
        warnings.append("Tubo com materia-prima incluida: indica o preco manual em EUR/m para calcular o custo do material.")
    if not material_supplied_by_client and explicit_profile_material_cost <= 0.0 and profile_length_m <= 0.0:
        warnings.append("Material do perfil incluido, mas falta comprimento/preco do perfil para calcular materia-prima.")
    analysis_note = str(analyzed_metrics.get("note", "") or "").strip()
    if analysis_note:
        warnings.append(analysis_note)
    return {
        "file_path": file_path,
        "geometry": {
            "file_path": file_path,
            "bbox_mm": {},
            "source": "step_profile",
        },
        "machine": {
            "name": str(machine_profile.get("name", "") or merged_settings.get("active_machine", "")),
            "material": display_material,
            "material_family": material_family,
            "material_subtype": material_subtype,
            "material_supplied_by_client": material_supplied_by_client,
            "material_cost_included": material_cost_included,
            "gas": gas_name,
            "thickness_mm": round(thickness_mm, 3),
            "cut_row": cut_row,
            "motion": motion_cfg,
            "effective_cut_speed_m_min": round(effective_cut_speed_m_min, 4),
            "reference_cut_speed_m_min": round(reference_cut_speed_m_min, 4),
            "current_cut_speed_m_min": round(current_cut_speed_m_min, 4),
        },
        "commercial": {
            "name": str(commercial_profile.get("name", "") or merged_settings.get("active_commercial", "")),
            "cost_mode": cost_mode,
            "effective_cutting_label": effective_cutting_label,
            "series_key": str(series_tier.get("key", "") or ""),
            "series_label": str(series_tier.get("label", "") or ""),
            "series_margin_delta_pct": round(series_margin_delta_pct, 3),
            "series_setup_multiplier": round(series_setup_multiplier, 4),
            "base_margin_pct": round(base_margin_pct, 3),
            "margin_pct": round(margin_pct, 3),
            "minimum_line_eur": 0.0,
            "minimum_applied": False,
            "base_setup_time_min": round(base_setup_time_min, 3),
            "setup_time_min": round(setup_time_min, 3),
            "handling_eur": round(handling_eur, 2),
            "material_utilization_pct": round(utilization_pct, 3),
            "material_supplied_by_client": material_supplied_by_client,
            "material_cost_included": material_cost_included,
            "material_price_source": material_price_source,
            "include_external_profile_cuts": include_external_cuts,
            "include_profile_event_rates": include_profile_event_rates,
            "include_profile_setup": include_profile_setup,
        },
        "metrics": {
            "cut_length_m": round(cut_length_m, 4),
            "cut_length_mm": round(cut_length_m * 1000.0, 3),
            "mark_length_m": 0.0,
            "rapid_length_m": round(rapid_length_m, 4),
            "pierce_count": int(pierce_count),
            "cut_event_count": int(total_cut_count),
            "outer_cut_count": int(outer_cut_count),
            "hole_count": int(hole_count),
            "slot_count": int(slot_count),
            "end_cut_count": int(end_cut_count),
            "raw_end_cut_count": int(raw_end_cut_count),
            "generic_cut_count": int(generic_cut_count),
            "feature_cut_count": int(max(0, hole_count + slot_count + generic_cut_count)),
            "feature_cut_length_m": round(_as_float(analyzed_metrics.get("feature_cut_length_m", 0.0), 0.0), 4),
            "end_cut_length_m": round(_as_float(analyzed_metrics.get("end_cut_length_m", 0.0), 0.0), 4),
            "thickness_rate_factor": round(thickness_rate_factor, 4),
            "net_area_m2": round(net_area_m2, 6),
            "density_kg_m3": round(density_kg_m3, 3),
            "profile_length_m": round(profile_length_m, 4),
            "profile_kg_m": round(profile_kg_m, 4),
            "profile_mass_kg": round(profile_mass_kg, 4),
            "net_mass_kg": round(net_mass_kg, 4),
            "gross_mass_kg": round(gross_mass_kg, 4),
            "scrap_mass_kg": round(scrap_mass_kg, 4),
        },
        "times": {
            "cut_sec": round(cut_time_sec, 3),
            "mark_sec": 0.0,
            "rapid_sec": round(rapid_time_sec, 3),
            "lead_sec": round(lead_time_sec, 3),
            "pierce_sec": round(pierce_time_sec, 3),
            "gas_sec": round(gas_time_sec, 3),
            "setup_share_sec": round(setup_time_sec_share, 3),
            "machine_total_sec": round(machine_time_total_unit, 3),
            "machine_total_min": round(machine_time_total_unit / 60.0, 3),
        },
        "pricing": {
            "material_cost_unit": _round_money(material_net_cost_unit, 4),
            "material_gross_cost_unit": _round_money(material_cost_unit, 4),
            "scrap_credit_unit": _round_money(scrap_credit_unit, 4),
            "profile_material_price": _round_money(profile_material_price, 4),
            "profile_material_price_unit": material_price_unit,
            "process_only_cost_unit": _round_money(max(0.0, subtotal_cost_unit - material_net_cost_unit), 4),
            "material_cost_included": material_cost_included,
            "profile_operation_cost_unit": _round_money(profile_operation_cost_unit, 4),
            "cut_cost_meter_unit": _round_money(cut_cost_meter_unit, 4),
            "pierce_cost_unit": _round_money(pierce_cost_unit, 4),
            "cut_rate_per_m_unit": _round_money(scaled_cut_rate_per_m, 4),
            "base_cut_rate_per_m_unit": _round_money(cut_rate_per_m, 4),
            "pierce_rate_unit": _round_money(scaled_pierce_rate, 4),
            "base_pierce_rate_unit": _round_money(pierce_rate, 4),
            "hole_cut_rate_unit": _round_money(scaled_hole_cut_rate, 4),
            "slot_cut_rate_unit": _round_money(scaled_slot_cut_rate, 4),
            "outer_cut_rate_unit": _round_money(scaled_outer_cut_rate, 4),
            "profile_process_cost_unit": _round_money(profile_process_cost_unit, 4),
            "machine_runtime_cost_unit": _round_money(machine_runtime_cost_unit, 4),
            "effective_cutting_cost_unit": _round_money(effective_cutting_cost_unit, 4),
            "mark_cost_unit": 0.0,
            "defilm_cost_unit": 0.0,
            "handling_unit": _round_money(handling_eur, 4),
            "subtotal_cost_unit": _round_money(subtotal_cost_unit, 4),
            "unit_price_before_min": _round_money(unit_price_before_min, 4),
            "unit_price": _round_money(final_unit, 4),
            "total_price": _round_money(final_total, 2),
            "quantity": quantity,
        },
        "line_suggestion": {
            "tipo_item": "servico",
            "ref_externa": ref_externa,
            "descricao": description,
            "material": display_material,
            "material_family": material_family,
            "material_subtype": material_subtype,
            "material_supplied_by_client": material_supplied_by_client,
            "material_fornecido_cliente": material_supplied_by_client,
            "material_cost_included": material_cost_included,
            "espessura": f"{_round_mm(thickness_mm, 3):g}",
            "operacao": "Corte Laser STEP/IGS",
            "tempo_peca_min": round(machine_time_total_unit / 60.0, 3),
            "qtd": quantity,
            "preco_unit": round(final_unit, 4),
            "desenho": file_path,
        },
        "warnings": list(dict.fromkeys([str(item or "").strip() for item in warnings if str(item or "").strip()])),
        "debug_json": json.dumps(
            {
                "machine": str(machine_profile.get("name", "") or ""),
                "commercial": str(commercial_profile.get("name", "") or ""),
                "material_family": material_family,
                "material_subtype": material_subtype,
                "material": display_material,
                "material_supplied_by_client": material_supplied_by_client,
                "material_price_source": material_price_source,
                "series_key": str(series_tier.get("key", "") or ""),
                "series_label": str(series_tier.get("label", "") or ""),
                "gas": gas_name,
                "thickness_mm": thickness_mm,
                "quantity": quantity,
                "cuts": total_cut_count,
                "holes": hole_count,
                "slots": slot_count,
                "outer_cuts": outer_cut_count,
                "generic_cuts": generic_cut_count,
            },
            ensure_ascii=False,
        ),
    }
