from __future__ import annotations

from collections import defaultdict
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
from typing import Any


_STEP_ENTITY_RE = re.compile(r"^(#\d+)\s*=\s*(.+);$")
_STEP_REF_RE = re.compile(r"#\d+")


def _step_refs(expr: str) -> list[str]:
    return _STEP_REF_RE.findall(str(expr or ""))


def _step_kind(expr: str) -> str:
    return str(expr or "").split("(", 1)[0].strip()


def _step_load_entities(path: str | Path) -> dict[str, str]:
    raw = Path(path).read_text(encoding="latin-1", errors="ignore")
    if "ISO-10303-21" not in raw[:128]:
        return {}
    entities: dict[str, str] = {}
    current: list[str] = []
    for line in raw.splitlines():
        chunk = line.rstrip()
        if not current and not chunk.startswith("#"):
            continue
        if chunk.startswith("#") and current:
            block = " ".join(current)
            match = _STEP_ENTITY_RE.match(block)
            if match:
                entities[match.group(1)] = match.group(2)
            current = [chunk]
            continue
        current.append(chunk)
        if chunk.endswith(";"):
            block = " ".join(current)
            match = _STEP_ENTITY_RE.match(block)
            if match:
                entities[match.group(1)] = match.group(2)
            current = []
    if current:
        block = " ".join(current)
        match = _STEP_ENTITY_RE.match(block)
        if match:
            entities[match.group(1)] = match.group(2)
    return entities


def _step_parse_numeric_tuple(expr: str) -> tuple[float, ...]:
    match = re.search(r"\(\s*'[^']*'\s*,\s*\(([^()]*)\)\s*\)\s*$", str(expr or ""))
    if not match:
        return ()
    values: list[float] = []
    for token in match.group(1).split(","):
        try:
            values.append(float(token.strip()))
        except Exception:
            return ()
    return tuple(values)


def _step_cartesian_point(entities: dict[str, str], ref: str) -> tuple[float, float, float] | None:
    values = _step_parse_numeric_tuple(entities.get(ref, ""))
    if len(values) != 3:
        return None
    return (float(values[0]), float(values[1]), float(values[2]))


def _step_direction(entities: dict[str, str], ref: str) -> tuple[float, float, float] | None:
    values = _step_parse_numeric_tuple(entities.get(ref, ""))
    if len(values) != 3:
        return None
    return (float(values[0]), float(values[1]), float(values[2]))


def _step_axis_placement(entities: dict[str, str], ref: str) -> dict[str, tuple[float, float, float]] | None:
    refs = _step_refs(entities.get(ref, ""))
    if len(refs) < 3:
        return None
    origin = _step_cartesian_point(entities, refs[0])
    axis = _step_direction(entities, refs[1])
    ref_dir = _step_direction(entities, refs[2])
    if origin is None or axis is None or ref_dir is None:
        return None
    return {"origin": origin, "axis": axis, "ref_dir": ref_dir}


def _step_vertex_point(entities: dict[str, str], ref: str) -> tuple[float, float, float] | None:
    refs = _step_refs(entities.get(ref, ""))
    if not refs:
        return None
    return _step_cartesian_point(entities, refs[0])


def _dominant_axis(vector: tuple[float, float, float]) -> int:
    values = [abs(float(part or 0.0)) for part in vector]
    return max(range(3), key=lambda index: values[index])


def _signed_offset(origin: tuple[float, float, float], normal: tuple[float, float, float]) -> float:
    return sum(float(origin[index]) * float(normal[index]) for index in range(3))


def _loop_geometry_kinds(entities: dict[str, str], loop_ref: str) -> list[str]:
    loop_expr = entities.get(loop_ref, "")
    kinds: list[str] = []
    for oriented_edge_ref in _step_refs(loop_expr):
        oriented_expr = entities.get(oriented_edge_ref, "")
        edge_refs = _step_refs(oriented_expr)
        if not edge_refs:
            continue
        edge_curve_expr = entities.get(edge_refs[0], "")
        geometry_refs = _step_refs(edge_curve_expr)
        if not geometry_refs:
            continue
        geometry_expr = entities.get(geometry_refs[-1], "")
        geometry_kind = _step_kind(geometry_expr)
        if geometry_kind:
            kinds.append(geometry_kind)
    return kinds


def _step_trailing_numbers(expr: str) -> tuple[float, ...]:
    match = re.search(r"\(([^()]*)\)\s*$", str(expr or ""))
    if not match:
        return ()
    values: list[float] = []
    for token in match.group(1).split(","):
        token = str(token or "").strip()
        if not token or token.startswith("#") or token.startswith(".") or token.startswith("'"):
            continue
        try:
            values.append(float(token))
        except Exception:
            continue
    return tuple(values)


def _loop_projection_signature(
    entities: dict[str, str],
    face_bound_ref: str,
    plane_axis_index: int,
) -> tuple[Any, ...] | None:
    bound_refs = _step_refs(entities.get(face_bound_ref, ""))
    if not bound_refs:
        return None
    loop_ref = bound_refs[0]
    projected_points: list[tuple[float, float]] = []
    projection_axes = [index for index in range(3) if index != int(plane_axis_index)]
    if len(projection_axes) != 2:
        return None
    for oriented_edge_ref in _step_refs(entities.get(loop_ref, "")):
        oriented_expr = entities.get(oriented_edge_ref, "")
        edge_refs = _step_refs(oriented_expr)
        if not edge_refs:
            continue
        edge_curve_expr = entities.get(edge_refs[0], "")
        edge_curve_refs = _step_refs(edge_curve_expr)
        for vertex_ref in edge_curve_refs[:2]:
            point = _step_vertex_point(entities, vertex_ref)
            if point is not None:
                projected_points.append((float(point[projection_axes[0]]), float(point[projection_axes[1]])))
        if not edge_curve_refs:
            continue
        geometry_ref = edge_curve_refs[-1]
        geometry_expr = entities.get(geometry_ref, "")
        geometry_kind = _step_kind(geometry_expr)
        if geometry_kind == "CIRCLE":
            geometry_refs = _step_refs(geometry_expr)
            placement = _step_axis_placement(entities, geometry_refs[0]) if geometry_refs else None
            radius_values = _step_trailing_numbers(geometry_expr)
            radius = float(radius_values[-1]) if radius_values else 0.0
            if placement is not None and radius > 0.0:
                origin = placement["origin"]
                projected_points.extend(
                    [
                        (float(origin[projection_axes[0]]) - radius, float(origin[projection_axes[1]])),
                        (float(origin[projection_axes[0]]) + radius, float(origin[projection_axes[1]])),
                        (float(origin[projection_axes[0]]), float(origin[projection_axes[1]]) - radius),
                        (float(origin[projection_axes[0]]), float(origin[projection_axes[1]]) + radius),
                    ]
                )
    if not projected_points:
        return None
    xs = [point[0] for point in projected_points]
    ys = [point[1] for point in projected_points]
    min_x = min(xs)
    max_x = max(xs)
    min_y = min(ys)
    max_y = max(ys)
    return (
        _classify_step_loop(entities, face_bound_ref),
        round((min_x + max_x) / 2.0, 3),
        round((min_y + max_y) / 2.0, 3),
        round(max_x - min_x, 3),
        round(max_y - min_y, 3),
    )


def _classify_step_loop(entities: dict[str, str], face_bound_ref: str) -> str:
    bound_refs = _step_refs(entities.get(face_bound_ref, ""))
    if not bound_refs:
        return "cut"
    geometry_kinds = _loop_geometry_kinds(entities, bound_refs[0])
    if not geometry_kinds:
        return "cut"
    unique_kinds = set(geometry_kinds)
    if unique_kinds == {"CIRCLE"}:
        return "hole"
    if "ELLIPSE" in unique_kinds:
        return "slot"
    if "B_SPLINE_CURVE_WITH_KNOTS" in unique_kinds or "B_SPLINE_CURVE" in unique_kinds:
        return "slot"
    if unique_kinds == {"LINE"}:
        return "slot"
    if "CIRCLE" in unique_kinds and "LINE" in unique_kinds:
        return "slot"
    if len(unique_kinds) > 1:
        return "slot"
    return "cut"


def _pick_principal_axis(plane_faces: list[dict[str, Any]], spans: list[float]) -> int:
    axis_candidates: list[tuple[float, int, int, int]] = []
    for axis_index in range(3):
        axis_faces = [face for face in plane_faces if int(face.get("axis_index", -1)) == axis_index]
        if not axis_faces:
            continue
        side_signs = {int(face.get("side_sign", 0)) for face in axis_faces}
        if side_signs != {-1, 1}:
            continue
        outer_faces: list[dict[str, Any]] = []
        for side_sign in (-1, 1):
            sign_faces = [
                face
                for face in axis_faces
                if int(face.get("side_sign", 0)) == side_sign and list(face.get("outer_bounds", []) or [])
            ]
            if not sign_faces:
                continue
            outer_faces.append(max(sign_faces, key=lambda face: abs(float(face.get("position_delta", 0.0)))))
        if len(outer_faces) < 2:
            continue
        total_internal_bounds = sum(len(list(face.get("face_bounds", []) or [])) for face in outer_faces)
        axis_span = float(spans[axis_index] if axis_index < len(spans) else 0.0)
        axis_candidates.append((axis_span, total_internal_bounds, -len(axis_faces), axis_index))
    if axis_candidates:
        axis_candidates.sort(reverse=True)
        return int(axis_candidates[0][3])
    return max(range(3), key=lambda index: spans[index])


def find_freecad_executable(*, prefer_gui: bool = False) -> str:
    env_names = ["LUGEST_FREECAD_EXE", "FREECAD_EXE"] if prefer_gui else ["LUGEST_FREECADCMD_EXE", "FREECADCMD_EXE"]
    candidates: list[Path] = []
    for env_name in env_names:
        value = str(os.environ.get(env_name, "") or "").strip().strip('"')
        if value:
            candidates.append(Path(value))
    common_roots = [
        Path(r"C:\Program Files"),
        Path(r"C:\Program Files (x86)"),
        Path.home() / "AppData" / "Local" / "Programs",
    ]
    gui_names = ("FreeCAD.exe",)
    cmd_names = ("FreeCADCmd.exe", "freecadcmd.exe")
    names = gui_names if prefer_gui else cmd_names
    for root in common_roots:
        if not root.exists():
            continue
        for version_dir in sorted(root.glob("FreeCAD*"), reverse=True):
            for exe_name in names:
                candidates.append(version_dir / "bin" / exe_name)
                candidates.append(version_dir / exe_name)
    which_candidates = ["FreeCAD"] if prefer_gui else ["FreeCADCmd"]
    for command_name in which_candidates:
        resolved = shutil.which(command_name)
        if resolved:
            candidates.append(Path(resolved))
    seen: set[str] = set()
    for candidate in candidates:
        key = str(candidate).strip().lower()
        if not key or key in seen:
            continue
        seen.add(key)
        if candidate.is_file():
            return str(candidate)
    return ""


def _freecad_bridge_script_path() -> Path:
    return Path(__file__).resolve().parents[2] / "scripts" / "freecad_step_bridge.py"


def _run_freecad_step_bridge(mode: str, source_path: str | Path) -> dict[str, Any]:
    bridge_path = _freecad_bridge_script_path()
    if not bridge_path.exists():
        return {}
    freecad_cmd = find_freecad_executable(prefer_gui=False)
    if not freecad_cmd:
        return {}
    source = Path(source_path).expanduser()
    if not source.exists():
        return {}
    cache_dir = Path(tempfile.gettempdir()) / "lugest_step_bridge"
    cache_dir.mkdir(parents=True, exist_ok=True)
    output_path = cache_dir / (
        hashlib.sha1(
            f"{mode}|{source.resolve()}|{source.stat().st_mtime_ns}|{bridge_path.stat().st_mtime_ns}".encode(
                "utf-8",
                errors="ignore",
            )
        ).hexdigest()
        + ".json"
    )
    completed = subprocess.run(
        [freecad_cmd, str(bridge_path), "--pass", str(mode), str(source), str(output_path)],
        capture_output=True,
        text=True,
        timeout=180,
    )
    if not output_path.exists():
        note_txt = str((completed.stderr or completed.stdout or "")).strip()
        return {"note": note_txt[:400]} if note_txt else {}
    try:
        payload = json.loads(output_path.read_text(encoding="utf-8"))
    except Exception:
        return {}
    if isinstance(payload, dict):
        return payload
    return {}


def analyze_step_profile_freecad(path: str | Path) -> dict[str, Any]:
    payload = _run_freecad_step_bridge("analyze", path)
    if not isinstance(payload, dict):
        return {}
    if str(payload.get("engine", "") or "").strip().lower() != "freecad":
        return {}
    if max(0, int(payload.get("cuts", 0) or 0)) <= 0 and not str(payload.get("note", "") or "").strip():
        return {}
    return payload


def render_step_preview_image(path: str | Path, *, size_px: int = 720) -> dict[str, Any]:
    source_path = Path(path).expanduser()
    if not source_path.exists():
        return {"available": False, "note": f"Ficheiro STEP/IGS nao encontrado: {source_path}"}
    suffix = str(source_path.suffix or "").strip().lower()
    if suffix not in {".step", ".stp", ".igs", ".iges"}:
        return {"available": False, "note": "Preview FreeCAD disponivel apenas para STEP/IGS."}
    if not find_freecad_executable(prefer_gui=False):
        return {"available": False, "note": "FreeCAD nao encontrado neste posto."}

    cache_dir = Path(tempfile.gettempdir()) / "lugest_step_preview"
    cache_dir.mkdir(parents=True, exist_ok=True)
    cache_key = hashlib.sha1(
        f"{source_path.resolve()}|{source_path.stat().st_mtime_ns}|{int(size_px)}".encode("utf-8", errors="ignore")
    ).hexdigest()
    image_path = cache_dir / f"{cache_key}.png"
    if image_path.exists():
        return {"available": True, "engine": "freecad", "image_path": str(image_path)}
    payload = _run_freecad_step_bridge("preview", source_path)
    polylines = list(payload.get("polylines", []) or []) if isinstance(payload, dict) else []
    if not polylines:
        note_txt = str((payload or {}).get("note", "") or "").strip()
        return {"available": False, "engine": "freecad", "image_path": "", "note": note_txt}
    try:
        from PIL import Image, ImageDraw
    except Exception:
        return {"available": False, "engine": "freecad", "image_path": "", "note": "PIL nao disponivel para gerar preview."}

    target_size = int(max(240, min(1600, size_px)))
    padding = max(12, int(round(target_size * 0.06)))
    all_points = [
        (float(point[0]), float(point[1]))
        for polyline in polylines
        for point in list(dict(polyline or {}).get("points", []) or [])
        if isinstance(point, (list, tuple)) and len(point) >= 2
    ]
    if not all_points:
        return {"available": False, "engine": "freecad", "image_path": "", "note": "Preview FreeCAD sem pontos projetados."}
    min_x = min(point[0] for point in all_points)
    max_x = max(point[0] for point in all_points)
    min_y = min(point[1] for point in all_points)
    max_y = max(point[1] for point in all_points)
    span_x = max(1e-6, max_x - min_x)
    span_y = max(1e-6, max_y - min_y)
    scale = min((target_size - (padding * 2)) / span_x, (target_size - (padding * 2)) / span_y)
    image = Image.new("RGB", (target_size, target_size), "#ffffff")
    draw = ImageDraw.Draw(image)
    ordered_polylines = sorted(polylines, key=lambda item: float(dict(item or {}).get("depth", 0.0) or 0.0))
    for polyline in ordered_polylines:
        points = []
        for raw_point in list(dict(polyline or {}).get("points", []) or []):
            if not isinstance(raw_point, (list, tuple)) or len(raw_point) < 2:
                continue
            px = padding + ((float(raw_point[0]) - min_x) * scale)
            py = target_size - (padding + ((float(raw_point[1]) - min_y) * scale))
            points.append((round(px, 2), round(py, 2)))
        if len(points) >= 2:
            draw.line(points, fill="#111827", width=max(1, int(round(target_size / 240.0))), joint="curve")
    image.save(image_path)
    return {"available": image_path.exists(), "engine": "freecad", "image_path": str(image_path), "note": ""}


def analyze_step_profile(path: str | Path) -> dict[str, Any]:
    freecad_result = analyze_step_profile_freecad(path)
    if freecad_result:
        return freecad_result
    entities = _step_load_entities(path)
    if not entities:
        return {}

    points = [
        point
        for entity_ref, expr in entities.items()
        if _step_kind(expr) == "CARTESIAN_POINT"
        for point in [_step_cartesian_point(entities, entity_ref)]
        if point is not None
    ]
    if not points:
        return {}

    spans = [max(point[index] for point in points) - min(point[index] for point in points) for index in range(3)]
    plane_faces: list[dict[str, Any]] = []
    for face_ref, expr in entities.items():
        if _step_kind(expr) != "ADVANCED_FACE":
            continue
        refs = _step_refs(expr)
        if len(refs) < 2:
            continue
        surface_ref = refs[-1]
        surface_expr = entities.get(surface_ref, "")
        if _step_kind(surface_expr) != "PLANE":
            continue
        placement_refs = _step_refs(surface_expr)
        if not placement_refs:
            continue
        placement = _step_axis_placement(entities, placement_refs[0])
        if placement is None:
            continue
        bound_refs = refs[:-1]
        face_bound_refs = [ref for ref in bound_refs if _step_kind(entities.get(ref, "")) == "FACE_BOUND"]
        face_outer_refs = [ref for ref in bound_refs if _step_kind(entities.get(ref, "")) == "FACE_OUTER_BOUND"]
        normal = placement["axis"]
        axis_index = _dominant_axis(normal)
        position = float(placement["origin"][axis_index])
        sign = 1 if normal[axis_index] >= 0 else -1
        plane_faces.append(
            {
                "face_ref": face_ref,
                "origin": placement["origin"],
                "normal": normal,
                "axis_index": axis_index,
                "sign": sign,
                "position": position,
                "offset": _signed_offset(placement["origin"], normal),
                "face_bounds": face_bound_refs,
                "outer_bounds": face_outer_refs,
            }
        )

    if not plane_faces:
        return {}

    axis_centers: dict[int, float] = {}
    for axis_index in range(3):
        axis_positions = [float(face.get("position", 0.0)) for face in plane_faces if int(face.get("axis_index", -1)) == axis_index]
        if not axis_positions:
            continue
        axis_centers[axis_index] = (min(axis_positions) + max(axis_positions)) / 2.0
    for face in plane_faces:
        axis_index = int(face.get("axis_index", -1))
        center = float(axis_centers.get(axis_index, 0.0))
        position_delta = float(face.get("position", 0.0)) - center
        face["position_delta"] = position_delta
        face["side_sign"] = 1 if position_delta >= 0 else -1

    principal_axis = _pick_principal_axis(plane_faces, spans)
    principal_axis_name = ("x", "y", "z")[principal_axis]

    lateral_groups: dict[tuple[int, int], list[dict[str, Any]]] = defaultdict(list)
    end_groups: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for face in plane_faces:
        axis_index = int(face["axis_index"])
        if axis_index == principal_axis:
            end_groups[int(face["side_sign"])].append(face)
        elif list(face["face_bounds"]):
            lateral_groups[(axis_index, int(face["side_sign"]))].append(face)

    selected_lateral_faces: list[dict[str, Any]] = []
    for group_faces in lateral_groups.values():
        selected_lateral_faces.append(max(group_faces, key=lambda face: abs(float(face.get("position_delta", 0.0)))))

    selected_end_faces: list[dict[str, Any]] = []
    for group_faces in end_groups.values():
        outer_candidates = [face for face in group_faces if list(face["outer_bounds"])]
        if not outer_candidates:
            continue
        selected_end_faces.append(max(outer_candidates, key=lambda face: abs(float(face.get("position_delta", 0.0)))))

    hole_count = 0
    slot_count = 0
    generic_cut_count = 0
    seen_feature_signatures: set[tuple[Any, ...]] = set()
    for face in selected_lateral_faces:
        plane_axis_index = int(face.get("axis_index", -1))
        for face_bound_ref in list(face.get("face_bounds", []) or []):
            loop_kind = _classify_step_loop(entities, face_bound_ref)
            loop_signature = _loop_projection_signature(entities, face_bound_ref, plane_axis_index)
            dedupe_key = loop_signature or (loop_kind, str(face_bound_ref))
            if dedupe_key in seen_feature_signatures:
                continue
            seen_feature_signatures.add(dedupe_key)
            if loop_kind == "hole":
                hole_count += 1
            elif loop_kind == "slot":
                slot_count += 1
            else:
                generic_cut_count += 1

    side_feature_count = hole_count + slot_count + generic_cut_count
    end_cut_count = sum(1 for face in selected_end_faces if list(face.get("outer_bounds", []) or []))
    outer_cuts = int(generic_cut_count + end_cut_count)
    cuts = int(side_feature_count + end_cut_count)

    notes = [
        f"topologia STEP analisada no eixo principal {principal_axis_name}",
        f"eventos adicionais de corte no perfil: {side_feature_count}",
        f"cortes terminais base: {end_cut_count}",
    ]
    if cuts:
        notes.append(f"eventos laser totais: {cuts}")
    if hole_count:
        notes.append(f"furos circulares detetados: {hole_count}")
    if slot_count:
        notes.append(f"rasgos/contornos nao circulares detetados: {slot_count}")
    if generic_cut_count:
        notes.append(f"aberturas/cortes genericos detetados: {generic_cut_count}")

    return {
        "mode": "step_topology",
        "cuts": cuts,
        "holes": hole_count,
        "slots": slot_count,
        "generic_cuts": generic_cut_count,
        "feature_cuts": side_feature_count,
        "outer_cuts": outer_cuts,
        "base_cuts": end_cut_count,
        "principal_axis": principal_axis_name,
        "side_feature_count": side_feature_count,
        "end_cut_count": end_cut_count,
        "selected_lateral_faces": [str(face.get("face_ref", "") or "") for face in selected_lateral_faces],
        "selected_end_faces": [str(face.get("face_ref", "") or "") for face in selected_end_faces],
        "note": ". ".join(note for note in notes if note).strip(),
    }


def analyze_profile_cut_features(path: str | Path) -> dict[str, Any]:
    suffix = str(Path(path).suffix or "").strip().lower()
    if suffix in {".step", ".stp"}:
        return analyze_step_profile(path)
    return {}
